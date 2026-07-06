Attribute VB_Name = "modStornoContext"
Option Explicit

' ============================================================
' modStornoContext - persistentni correction/storno context (tblStornoVeze)
'
' Svaka storno/ispravka dobija red u tblStornoVeze: staro -> novo trag, mod,
' status, recovery flag. Kontekst je PERSISTENTAN (preziveljava zatvaranje
' forme/Excela) -- nije oslonjen na module-level promenljive. Module-level
' pending je dozvoljen SAMO kao kratkorocni UI state dok je forma otvorena
' (modStornoFlow drzi aktivni CorrectionID), ali poslovni trag zivi ovde.
'
' Statusi: PENDING / COMPLETED / FAILED / MANUAL_REQUIRED / CANCELLED.
' NeedsRecovery = "Da" za sve sto nije COMPLETED/CANCELLED -> to je izvor
' recovery panela (Osiroceni dokumenti) i badge brojaca.
'
' Reuse: modDataAccess (AppendRow/UpdateCell/GetTableData), clsTransaction,
' RequireColumnIndex/RequireUpdateCell, Monitor_Event, LogErr. Bez MsgBox.
' ============================================================

Private Const MOD_NAME As String = "modStornoContext"
Private Const FLAG_DA As String = "Da"
Private Const FLAG_NE As String = "Ne"

' ============================================================
' CREATE - novi PENDING context. Vraca CorrectionID (prazan string na gresku).
' Sopstvena transakcija -> red se odmah trajno upisuje (nezavisno od poslovnog
' storna koji se moze rollback-ovati; pending trag mora ostati vidljiv).
' ============================================================
Public Function CreateCorrectionContext(ByVal mode As String, _
        ByVal oldDocType As String, ByVal oldDocID As String, ByVal oldBroj As String, _
        Optional ByVal newDocType As String = "", Optional ByVal newDocID As String = "", _
        Optional ByVal newBroj As String = "", _
        Optional ByVal parentDocType As String = "", Optional ByVal parentDocID As String = "", _
        Optional ByVal parentBroj As String = "", _
        Optional ByVal message As String = "") As String
    Const SRC As String = MOD_NAME & ".CreateCorrectionContext"
    Dim tx As clsTransaction
    On Error GoTo EH

    EnsureTable

    Dim newID As String
    newID = GetNextID(TBL_STORNO_VEZE, COL_SV_ID, "COR-")

    ' Redosled MORA pratiti EnsureStornoVezeSchemaCore (AppendRow po poziciji):
    ' CorrectionID, Mode, Status, OldDocType, OldDocID, OldBroj, NewDocType,
    ' NewDocID, NewBroj, ParentDocType, ParentDocID, ParentBroj, CreatedAt,
    ' CreatedBy, CompletedAt, Message, NeedsRecovery, RecoveryAction
    Dim rowData(0 To 17) As Variant
    rowData(0) = newID
    rowData(1) = mode
    rowData(2) = SV_STATUS_PENDING
    rowData(3) = oldDocType
    rowData(4) = oldDocID
    rowData(5) = oldBroj
    rowData(6) = newDocType
    rowData(7) = newDocID
    rowData(8) = newBroj
    rowData(9) = parentDocType
    rowData(10) = parentDocID
    rowData(11) = parentBroj
    rowData(12) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    rowData(13) = CurrentUser()
    rowData(14) = ""                       ' CompletedAt
    rowData(15) = message
    rowData(16) = FLAG_DA                   ' NeedsRecovery (PENDING dok se ne zavrsi)
    rowData(17) = ""                        ' RecoveryAction

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_STORNO_VEZE
    If AppendRow(TBL_STORNO_VEZE, rowData) = 0 Then
        Err.Raise ERR_STORNO_FW_BASE + 10, SRC, "AppendRow tblStornoVeze nije uspeo."
    End If
    tx.CommitTx
    Set tx = Nothing

    CreateCorrectionContext = newID
    LogContext "STORNO_CTX_CREATE", "INFO", _
        "Correction " & newID & " [" & mode & "] " & oldDocType & " " & oldBroj, _
        newID, oldDocType, oldBroj
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    CreateCorrectionContext = ""
End Function

' ============================================================
' COMPLETE - context uspesno zavrsen (staro -> novo povezano/rekalkulisano).
' ============================================================
Public Function CompleteCorrectionContext(ByVal correctionID As String, _
        Optional ByVal newDocID As String = "", Optional ByVal newBroj As String = "", _
        Optional ByVal message As String = "") As Boolean
    Const SRC As String = MOD_NAME & ".CompleteCorrectionContext"
    Dim tx As clsTransaction
    On Error GoTo EH

    Dim rowIdx As Long
    rowIdx = RequireCorrectionRow(correctionID, SRC)

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_STORNO_VEZE

    RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_STATUS, SV_STATUS_COMPLETED, SRC
    RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_COMPLETED_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss"), SRC
    RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_NEEDS_RECOVERY, FLAG_NE, SRC
    If Len(newDocID) > 0 Then RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_NEW_DOCID, newDocID, SRC
    If Len(newBroj) > 0 Then RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_NEW_BROJ, newBroj, SRC
    If Len(message) > 0 Then RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_MESSAGE, message, SRC

    tx.CommitTx
    Set tx = Nothing

    CompleteCorrectionContext = True
    LogContext "STORNO_CTX_COMPLETE", "INFO", _
        "Correction " & correctionID & " COMPLETED. Novo: " & newBroj, _
        correctionID, "", newBroj
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    CompleteCorrectionContext = False
End Function

' ============================================================
' FAIL - automatski storno/reassign/recalc nije uspeo. Ostaje vidljiv u recovery
' panelu (NeedsRecovery = Da). Best-effort: nikad ne baca (poziva se iz EH bloka).
' ============================================================
Public Function FailCorrectionContext(ByVal correctionID As String, _
        ByVal message As String) As Boolean
    FailCorrectionContext = SetTerminalState(correctionID, SV_STATUS_FAILED, FLAG_DA, _
                                             message, "", "STORNO_CTX_FAIL", "ERROR")
End Function

' ============================================================
' MANUAL_REQUIRED - potreban rucni korak (npr. relink nije uspeo, osirocene
' palete cekaju re-point). RecoveryAction opisuje sta operater treba da uradi.
' ============================================================
Public Function MarkCorrectionManual(ByVal correctionID As String, _
        ByVal recoveryAction As String, ByVal message As String) As Boolean
    MarkCorrectionManual = SetTerminalState(correctionID, SV_STATUS_MANUAL, FLAG_DA, _
                                            message, recoveryAction, "STORNO_CTX_MANUAL", "WARN")
End Function

' ============================================================
' CANCEL - operater odustao (npr. odbio potvrdu). Nije za recovery.
' ============================================================
Public Function CancelCorrectionContext(ByVal correctionID As String, _
        Optional ByVal message As String = "") As Boolean
    CancelCorrectionContext = SetTerminalState(correctionID, SV_STATUS_CANCELLED, FLAG_NE, _
                                               message, "", "STORNO_CTX_CANCEL", "INFO")
End Function

' Zajednicki upis terminalnog statusa (+CompletedAt + NeedsRecovery + Recovery).
Private Function SetTerminalState(ByVal correctionID As String, ByVal status As String, _
        ByVal needsRecovery As String, ByVal message As String, ByVal recoveryAction As String, _
        ByVal evtType As String, ByVal severity As String) As Boolean
    Const SRC As String = MOD_NAME & ".SetTerminalState"
    Dim tx As clsTransaction
    On Error GoTo EH

    Dim rowIdx As Long
    rowIdx = RequireCorrectionRow(correctionID, SRC)

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_STORNO_VEZE

    RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_STATUS, status, SRC
    RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_NEEDS_RECOVERY, needsRecovery, SRC
    RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_COMPLETED_AT, Format$(Now, "yyyy-mm-dd hh:nn:ss"), SRC
    If Len(message) > 0 Then RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_MESSAGE, message, SRC
    If Len(recoveryAction) > 0 Then RequireUpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_RECOVERY_ACTION, recoveryAction, SRC

    tx.CommitTx
    Set tx = Nothing

    SetTerminalState = True
    LogContext evtType, severity, status & ": " & correctionID & " " & message, _
               correctionID, "", ""
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    SetTerminalState = False
End Function

' ============================================================
' READ helpers
' ============================================================

' Red index (1-based, u DataBodyRange) za dati CorrectionID; 0 ako nema.
Public Function GetCorrectionRowByID(ByVal correctionID As String) As Long
    On Error GoTo EH
    correctionID = Trim$(correctionID)
    If Len(correctionID) = 0 Then Exit Function
    If GetTable(TBL_STORNO_VEZE) Is Nothing Then Exit Function

    Dim c As Collection
    Set c = FindRows(TBL_STORNO_VEZE, COL_SV_ID, correctionID)
    If c Is Nothing Then Exit Function
    If c.count = 0 Then Exit Function
    GetCorrectionRowByID = CLng(c(c.count))     ' poslednji (najsvezi) ako ima duplikat
    Exit Function
EH:
    LogErr MOD_NAME & ".GetCorrectionRowByID"
End Function

' Vrednost jednog polja context reda (prazno ako nema).
Public Function GetCorrectionField(ByVal correctionID As String, ByVal colName As String) As String
    On Error GoTo EH
    Dim v As Variant
    v = LookupValue(TBL_STORNO_VEZE, COL_SV_ID, correctionID, colName)
    If Not IsEmpty(v) Then GetCorrectionField = Trim$(CStr(v))
    Exit Function
EH:
    LogErr MOD_NAME & ".GetCorrectionField"
End Function

' Kolekcija Dictionary-ja za sve redove kojima treba recovery (NeedsRecovery=Da).
' Svaki dict nosi kljuceve: id, mode, status, oldDocType, oldBroj, newBroj,
' parentBroj, message, recoveryAction, createdAt.
Public Function GetPendingCorrections() As Collection
    Const SRC As String = MOD_NAME & ".GetPendingCorrections"
    Dim result As New Collection
    Set GetPendingCorrections = result

    On Error GoTo EH
    If GetTable(TBL_STORNO_VEZE) Is Nothing Then Exit Function
    Dim data As Variant
    data = GetTableData(TBL_STORNO_VEZE)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cMode As Long, cStatus As Long, cNeeds As Long
    Dim cOldT As Long, cOldB As Long, cNewB As Long, cParB As Long, cMsg As Long
    Dim cRec As Long, cCreated As Long
    cId = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_ID)
    cMode = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_MODE)
    cStatus = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_STATUS)
    cNeeds = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_NEEDS_RECOVERY)
    cOldT = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_OLD_DOCTYPE)
    cOldB = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_OLD_BROJ)
    cNewB = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_NEW_BROJ)
    cParB = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_PARENT_BROJ)
    cMsg = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_MESSAGE)
    cRec = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_RECOVERY_ACTION)
    cCreated = GetColumnIndex(TBL_STORNO_VEZE, COL_SV_CREATED_AT)
    If cNeeds = 0 Or cId = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(CStr(data(i, cNeeds)))) = "DA" Then
            Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
            d("id") = CStr(data(i, cId))
            d("mode") = TxCell(data, i, cMode)
            d("status") = TxCell(data, i, cStatus)
            d("oldDocType") = TxCell(data, i, cOldT)
            d("oldBroj") = TxCell(data, i, cOldB)
            d("newBroj") = TxCell(data, i, cNewB)
            d("parentBroj") = TxCell(data, i, cParB)
            d("message") = TxCell(data, i, cMsg)
            d("recoveryAction") = TxCell(data, i, cRec)
            d("createdAt") = TxCell(data, i, cCreated)
            result.Add d
        End If
    Next i
    Exit Function
EH:
    LogErr SRC
End Function

' Broj context redova koji cekaju recovery (za badge / brzu proveru).
Public Function CountPendingRecovery() As Long
    On Error Resume Next
    Dim c As Collection
    Set c = GetPendingCorrections()
    If Not c Is Nothing Then CountPendingRecovery = c.count
End Function

' Najsveziji (poslednji upisan) PENDING context za dati OldDocType + Mode. Koristi
' se za "Zavrsi ispravku" u novoj sesiji kad module-level UI state ne postoji.
' Prazan string ako nema takvog. (Prazan oldDocType = bilo koji tip.)
Public Function FindLatestPending(ByVal oldDocType As String, ByVal mode As String) As String
    On Error GoTo EH
    Dim c As Collection
    Set c = GetPendingCorrections()
    If c Is Nothing Then Exit Function
    Dim i As Long, res As String, d As Object
    For i = 1 To c.count
        Set d = c(i)
        If UCase$(CStr(d("status"))) = UCase$(SV_STATUS_PENDING) _
           And StrComp(CStr(d("mode")), mode, vbTextCompare) = 0 Then
            If Len(oldDocType) = 0 Or StrComp(CStr(d("oldDocType")), oldDocType, vbTextCompare) = 0 Then
                res = CStr(d("id"))     ' zadnji match u redosledu tabele = najsveziji
            End If
        End If
    Next i
    FindLatestPending = res
    Exit Function
EH:
    LogErr MOD_NAME & ".FindLatestPending"
End Function

' ============================================================
' PRIVATE
' ============================================================

Private Function RequireCorrectionRow(ByVal correctionID As String, ByVal SRC As String) As Long
    Dim rowIdx As Long
    rowIdx = GetCorrectionRowByID(correctionID)
    If rowIdx = 0 Then
        Err.Raise ERR_STORNO_FW_BASE + 11, SRC, _
                  "Correction context nije pronadjen: " & correctionID
    End If
    RequireCorrectionRow = rowIdx
End Function

Private Sub EnsureTable()
    If GetTable(TBL_STORNO_VEZE) Is Nothing Then
        On Error Resume Next
        modSetup.EnsureStornoVezeSchemaCore
        On Error GoTo 0
    End If
End Sub

Private Function CurrentUser() As String
    On Error Resume Next
    Dim u As String
    u = modAuth.GetCurrentUser()
    If Len(Trim$(u)) = 0 Then u = Environ$("USERNAME")
    If Len(Trim$(u)) = 0 Then u = "Operator"
    CurrentUser = u
End Function

Private Function TxCell(ByVal data As Variant, ByVal r As Long, ByVal c As Long) As String
    If c > 0 Then TxCell = Trim$(CStr(data(r, c)))
End Function

Private Sub LogContext(ByVal evtType As String, ByVal severity As String, ByVal message As String, _
                       ByVal correctionID As String, ByVal entityType As String, ByVal entityID As String)
    On Error Resume Next
    Monitor_Event eventType:=evtType, severity:=severity, message:=message, _
        moduleName:=MOD_NAME, procedureName:="", _
        entityType:="Correction", entityID:=correctionID, correlationId:=entityID, _
        payloadJson:="{""docType"":""" & entityType & """}"
End Sub
