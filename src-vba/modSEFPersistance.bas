Attribute VB_Name = "modSEFPersistance"
Option Explicit

' v. SefLogPadTestSet -- seam za pad citanja dnevnika dogadjaja.
Private mSefLogPadTest As String

' =========================================================
' modSEFPersistence
' Alle SEF-Reads/Writes laufen ueber modDataAccess
' =========================================================
Private Const TBL_SEF_SUBMISSION As String = "tblSEFSubmission"
Private Const TBL_SEF_EVENT_LOG As String = "tblSEFEventLog"

' =========================
' READ HELPERS
' =========================

Public Function GetFakturaSEFWorkflowState(ByVal fakturaID As String) As String
    GetFakturaSEFWorkflowState = GetFakturaSEFFieldText( _
        fakturaID, "SEFWorkflowState", "modSEFPersistance.GetFakturaSEFWorkflowState")
End Function

Public Function GetFakturaSEFDocumentId(ByVal fakturaID As String) As String
    GetFakturaSEFDocumentId = GetFakturaSEFFieldText( _
        fakturaID, "SEFDocumentId", "modSEFPersistance.GetFakturaSEFDocumentId")
End Function

Public Function GetLastSEFSubmissionID(ByVal fakturaID As String) As String
    GetLastSEFSubmissionID = GetFakturaSEFFieldText( _
        fakturaID, "SEFSubmissionIDLast", "modSEFPersistance.GetLastSEFSubmissionID")
End Function

Public Function GetNextSEFVersionNo(ByVal fakturaID As String) As Long
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetNextSEFVersionNo"

    Dim rawValue As String
    Dim currentVersion As Long

    rawValue = GetFakturaSEFFieldText(fakturaID, "SEFVersionNo", SRC)

    If rawValue = "" Then
        GetNextSEFVersionNo = 1
    ElseIf Not TryParseLong(rawValue, currentVersion) Then
        GetNextSEFVersionNo = 1
    Else
        GetNextSEFVersionNo = currentVersion + 1
    End If

    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Function

Public Function GetCurrentSEFVersionNo(ByVal fakturaID As String) As Long
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetCurrentSEFVersionNo"

    Dim rawValue As String
    Dim currentVersion As Long

    rawValue = GetFakturaSEFFieldText(fakturaID, "SEFVersionNo", SRC)

    If rawValue = "" Then
        GetCurrentSEFVersionNo = 0
    ElseIf Not TryParseLong(rawValue, currentVersion) Then
        GetCurrentSEFVersionNo = 0
    Else
        GetCurrentSEFVersionNo = currentVersion
    End If

    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Function

' =========================
' WRITE HELPERS
' =========================
' NOTE:
' newState controls internal workflow transitions only.
' sefStatus stores the latest known external SEF API status.
' Do not assume newState and sefStatus must be identical.

Public Sub UpdateFakturaSEFState_Row( _
    ByVal fakturaID As String, _
    ByVal newState As String, _
    Optional ByVal sefStatus As String = "", _
    Optional ByVal sefDocumentId As String = "", _
    Optional ByVal errorCode As String = "", _
    Optional ByVal errorMessage As String = "", _
    Optional ByVal payloadHash As String = "", _
    Optional ByVal submissionID As String = "", _
    Optional ByVal versionNo As Long = 0)

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.UpdateFakturaSEFState_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    If Len(Trim$(newState)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "newState is required."
    End If

    RequireFaktureSEFSchema SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    Dim oldState As String
    oldState = GetFakturaSEFWorkflowState(fakturaID)

    If Len(oldState) > 0 Then
        ValidateAllowedTransition oldState, newState
    End If

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFWorkflowState", newState, SRC

    If Len(sefStatus) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFStatus", sefStatus, SRC
    End If

    If Len(sefDocumentId) > 0 Then
        EnsureSEFDocumentIdTextFormat TBL_FAKTURE, rowIndex
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFDocumentId", sefDocumentId, SRC
    End If

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorCode", errorCode, SRC
    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorMessage", errorMessage, SRC

    If Len(payloadHash) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFPayloadHash", payloadHash, SRC
    End If

    If Len(submissionID) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFSubmissionIDLast", submissionID, SRC
    End If

    If versionNo > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFVersionNo", versionNo, SRC
    End If

    Select Case newState
        Case WF_SEF_SENT, WF_SEF_ACCEPTED
            RequireUpdateCell TBL_FAKTURE, rowIndex, "PoslatNaSEF", "Da", SRC

            If Len(Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFSentAt")))) = 0 Then
                RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFSentAt", Now, SRC
            End If

        Case WF_SEF_SENDING
            RequireUpdateCell TBL_FAKTURE, rowIndex, "PoslatNaSEF", "Ne", SRC
    End Select

    Select Case newState
        Case WF_SEF_SENT, WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_SYNC_ERROR
            RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastSyncAt", Now, SRC
    End Select

    Exit Sub

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Sub

' NOTE:
' This helper updates refresh-related fields without performing
' a workflow transition. Use it when SEFStatus changes but the
' internal workflow state should remain unchanged.

Public Sub UpdateFakturaSEFRefreshFields_Row( _
    ByVal fakturaID As String, _
    Optional ByVal sefStatus As String = "", _
    Optional ByVal sefDocumentId As String = "", _
    Optional ByVal errorCode As String = "", _
    Optional ByVal errorMessage As String = "")

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.UpdateFakturaSEFRefreshFields_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    RequireFaktureSEFSchema SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    If Len(sefStatus) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFStatus", sefStatus, SRC
    End If

    If Len(sefDocumentId) > 0 Then
        EnsureSEFDocumentIdTextFormat TBL_FAKTURE, rowIndex
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFDocumentId", sefDocumentId, SRC
    End If

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorCode", errorCode, SRC
    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorMessage", errorMessage, SRC

    Exit Sub

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Sub

Public Function CreateSEFSubmission_Row( _
    ByVal fakturaID As String, _
    ByVal versionNo As Long, _
    ByVal workflowState As String, _
    ByVal payloadHash As String, _
    ByVal requestBody As String, _
    ByVal requestFormat As String) As String

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.CreateSEFSubmission_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    If versionNo <= 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "VersionNo must be > 0."
    End If

    If Len(Trim$(workflowState)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "WorkflowStateAtSubmit is required."
    End If

    If Len(Trim$(requestBody)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "RequestBody is required."
    End If

    If Len(Trim$(requestFormat)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "RequestFormat is required."
    End If

    RequireSEFSubmissionSchema SRC

    Dim submissionID As String
    submissionID = GetNextID(TBL_SEF_SUBMISSION, "SEFSubmissionID", "SFS-")

    If Len(Trim$(submissionID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "GetNextID did not return SEFSubmissionID."
    End If

    Dim rowData(1 To 20) As Variant

    rowData(1) = submissionID
    rowData(2) = fakturaID
    rowData(3) = versionNo
    rowData(4) = workflowState
    rowData(5) = Now
    rowData(6) = Empty
    rowData(7) = SEF_SUB_CREATED
    rowData(8) = payloadHash
    rowData(9) = requestFormat
    rowData(10) = requestBody
    rowData(11) = Empty
    rowData(12) = Empty
    rowData(13) = Empty
    rowData(14) = Empty
    rowData(15) = Empty
    rowData(16) = Empty
    rowData(17) = Empty
    rowData(18) = GetCurrentOperatorName()
    rowData(19) = "Ne"
    rowData(20) = Empty

    Dim newRowIndex As Long
    newRowIndex = AppendRow(TBL_SEF_SUBMISSION, rowData)

    If newRowIndex <= 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Could not append row to tblSEFSubmission."
    End If

    CreateSEFSubmission_Row = submissionID
    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Function

Public Sub SaveSEFSubmissionResult_Row( _
    ByVal submissionID As String, _
    ByVal response As clsSEFResponse)

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.SaveSEFSubmissionResult_Row"

    If Len(Trim$(submissionID)) = 0 Then Exit Sub

    If response Is Nothing Then
        Err.Raise ERR_SEF_RESPONSE_PARSE, SRC, _
                  "Response object is Nothing."
    End If

    RequireSEFSubmissionSchema SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, True)

    Dim subStatus As String

    If response.Accepted Then
        subStatus = SEF_SUB_ACCEPTED
    ElseIf response.Rejected Then
        subStatus = SEF_SUB_REJECTED
    ElseIf response.Success Then
        subStatus = SEF_SUB_SENT
    Else
        subStatus = SEF_SUB_FAILED
    End If

    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "SubmittedAt", Now, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "FinishedAt", Now, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "SubmissionStatus", subStatus, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "HttpStatus", response.httpStatus, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ApiStatus", response.apiStatus, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "CorrelationId", response.correlationId, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ResponseBody", response.rawBody, SRC
    ' Ensure text format on tblSEFSubmission.SEFDocumentId column too
    Dim subDocCol As ListColumn
    On Error Resume Next
    Set subDocCol = GetTable(TBL_SEF_SUBMISSION).ListColumns("SEFDocumentId")
    If Not subDocCol Is Nothing Then
        If Not subDocCol.DataBodyRange Is Nothing Then
            subDocCol.DataBodyRange.cells(rowIndex, 1).NumberFormat = "@"
        End If
    End If
    On Error GoTo EH
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "SEFDocumentId", response.sefDocumentId, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ErrorCode", response.errorCode, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ErrorMessage", response.errorMessage, SRC

    Exit Sub

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Sub

Public Sub AppendSEFEvent_Row( _
    ByVal fakturaID As String, _
    ByVal submissionID As String, _
    ByVal eventType As String, _
    ByVal message As String, _
    Optional ByVal details As String = "")

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.AppendSEFEvent_Row"

    RequireSEFEventLogSchema SRC

    Dim eventID As String
    eventID = GetNextID(TBL_SEF_EVENT_LOG, "SEFEventID", "SFE-")

    If Len(Trim$(eventID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "GetNextID did not return SEFEventID."
    End If

    Dim rowData(1 To 9) As Variant

    rowData(1) = eventID
    rowData(2) = fakturaID
    rowData(3) = submissionID
    rowData(4) = Now
    rowData(5) = eventType
    rowData(6) = message
    rowData(7) = details
    rowData(8) = GetCurrentOperatorName()
    rowData(9) = "Ne"

    Dim newRowIndex As Long
    newRowIndex = AppendRow(TBL_SEF_EVENT_LOG, rowData)

    If newRowIndex <= 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Could not append row to tblSEFEventLog."
    End If

    Exit Sub

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Sub

' =========================
' INTERNAL HELPERS
' =========================

Private Function GetSingleRowIndexByKey( _
    ByVal tblName As String, _
    ByVal keyColName As String, _
    ByVal keyValue As Variant, _
    Optional ByVal raiseIfNotFound As Boolean = False) As Long
    
    Dim rowsFound As Collection
    
    Set rowsFound = FindRows(tblName, keyColName, keyValue)
    
    If rowsFound.count = 0 Then
        If raiseIfNotFound Then
            Err.Raise ERR_SEF_STATE, "GetSingleRowIndexByKey", _
                "Row not found in " & tblName & " for " & keyColName & "=" & CStr(keyValue)
        End If
        GetSingleRowIndexByKey = 0
        Exit Function
    End If
    
    If rowsFound.count > 1 Then
        Err.Raise ERR_SEF_DUPLICATE, "GetSingleRowIndexByKey", _
            "Multiple rows found in " & tblName & " for " & keyColName & "=" & CStr(keyValue)
    End If
    
    GetSingleRowIndexByKey = CLng(rowsFound(1))
End Function

Private Sub EnsureSEFDocumentIdTextFormat(ByVal tblName As String, _
                                          ByVal rowIndex As Long)
    ' Forces the SEFDocumentId cell to Text format BEFORE RequireUpdateCell
    ' writes the value. This prevents Excel from:
    '   - converting "5317568" to Double (precision loss for 12+ digits)
    '   - rendering large IDs as scientific notation ("5.31757E+06")
    '   - rejecting GUID values as #NAME? errors
    '
    ' RequireUpdateCell remains the canonical write helper - this just
    ' prepares the cell so the write is preserved exactly as text.
    
    Dim lo As ListObject
    Dim col As ListColumn
    Dim cell As Range
    
    On Error Resume Next
    
    Set lo = GetTable(tblName)
    If lo Is Nothing Then Exit Sub
    
    Set col = lo.ListColumns("SEFDocumentId")
    If col Is Nothing Then Exit Sub
    If col.DataBodyRange Is Nothing Then Exit Sub
    
    Set cell = col.DataBodyRange.cells(rowIndex, 1)
    cell.NumberFormat = "@"
    
    On Error GoTo 0
End Sub

Public Sub UpdateSEFLastSyncAt_Row(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.UpdateSEFLastSyncAt_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", SRC
    RequireColumnIndex TBL_FAKTURE, "SEFLastSyncAt", SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastSyncAt", Now, SRC

    Exit Sub

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Sub

Private Function GetCurrentOperatorName() As String
    On Error Resume Next
    
    GetCurrentOperatorName = Environ$("Username")
    
    If Len(Trim$(GetCurrentOperatorName)) = 0 Then
        GetCurrentOperatorName = Application.userName
    End If
    
    If Len(Trim$(GetCurrentOperatorName)) = 0 Then
        GetCurrentOperatorName = "UNKNOWN"
    End If
End Function


Public Function GetSEFSubmissionsForFaktura(ByVal fakturaID As String) As Variant
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSEFSubmissionsForFaktura"

    Dim data As Variant
    data = GetTableData(TBL_SEF_SUBMISSION)

    If IsEmpty(data) Then
        GetSEFSubmissionsForFaktura = Empty
        Exit Function
    End If

    Dim filters As Collection
    Dim fp As clsFilterParam

    Set filters = New Collection
    Set fp = New clsFilterParam

    filters.Add fp.Init(RequireColumnIndex(TBL_SEF_SUBMISSION, "FakturaID", SRC), "=", fakturaID)

    GetSEFSubmissionsForFaktura = FilterArray(data, filters)
    Exit Function

EH:
    LogErr SRC
    GetSEFSubmissionsForFaktura = Empty
End Function

' TEST SEAM: dnevnik se cita iz tabele KOJE NEMA.
'
' Prva verzija je samo dizala gresku i time merila da EH prosledjuje -- ali
' NIJE prolazila kroz putanju koja je i bila kvar: GetTableData za nepostojecu
' tabelu vraca Empty (modDataAccess), pa se rano izlazi PRE nego sto se EH i
' RequireColumnIndex uopste aktiviraju. Nepostojeca tabela je tako i dalje
' izgledala kao prazan dnevnik.
'
' Zato seam podmece IME nepostojece tabele: put je onda stvaran -- RequireTable
' je jedino sto stoji izmedju te tabele i tihog Empty. Sema se u harnessu ne sme
' lomiti, pa je ovo jedini nacin da se ta razlika izmeri.
'
' Vezan je za TEST REZIM pri svakom citanju, ne samo pri postavljanju: seam koji
' ostane upaljen (test koji je pukao pre ciscenja) postaje inertan cim
' RunAllTests vrati prethodni rezim.
' DVA REZIMA, jer su to dva razlicita kvara:
'   "TABELA"  -- dnevnika NEMA. GetTableData za nepostojecu tabelu vraca isti
'                Empty kao za praznu, pa se izlazi PRE EH-a; jedino RequireTable
'                tu razliku pravi.
'   "CITANJE" -- tabela postoji, ali citanje puca (izgubljena kolona, pad
'                filtera). Tu se meri da EH gresku PROSLEDI, a ne proguta.
' Prazno gasi seam.
Public Sub SefLogPadTestSet(ByVal rezim As String)
    If Not IsTestMode() Then Exit Sub
    mSefLogPadTest = UCase$(Trim$(rezim))
End Sub

Private Function SefLogTabela() As String
    SefLogTabela = TBL_SEF_EVENT_LOG
    If mSefLogPadTest = "TABELA" And IsTestMode() Then _
        SefLogTabela = "tblSEFEventLogNePostoji"
End Function

Public Function GetSEFEventsForFaktura(ByVal fakturaID As String) As Variant
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSEFEventsForFaktura"

    Dim tabela As String
    tabela = SefLogTabela()

    ' NEPOSTOJECA TABELA NIJE PRAZAN DNEVNIK.
    '
    ' GetTableData za tabelu koje nema vraca Empty, isto kao za tabelu koja je
    ' prazna. Bez ove kapije bi se rano izaslo PRE nego sto se EH i
    ' RequireColumnIndex uopste aktiviraju, pa bi nedostajuca tabela operateru
    ' izgledala kao 'faktura nema dogadjaja'. RequireTable je jedino sto tu
    ' razliku pravi -- i mora da stoji PRE citanja, ne posle.
    RequireTable tabela, SRC

    ' Drugi rezim seam-a: tabela POSTOJI, ali citanje puca. Namerno UNUTAR
    ' On Error GoTo EH -- meri se bas to da EH gresku prosledjuje.
    If mSefLogPadTest = "CITANJE" And IsTestMode() Then _
        Err.Raise ERR_SEF_STATE, SRC, "SEAM: citanje dnevnika pada"

    Dim data As Variant
    data = GetTableData(tabela)

    ' Ovde Empty vise moze da znaci SAMO jedno: tabela postoji i prazna je.
    If IsEmpty(data) Then
        GetSEFEventsForFaktura = Empty
        Exit Function
    End If

    Dim filters As Collection
    Dim fp As clsFilterParam

    Set filters = New Collection
    Set fp = New clsFilterParam

    filters.Add fp.Init(RequireColumnIndex(tabela, "FakturaID", SRC), "=", fakturaID)

    GetSEFEventsForFaktura = FilterArray(data, filters)
    Exit Function

' PAD CITANJA NIJE PRAZAN DNEVNIK.
'
' Do v2.39 je svaka greska ovde postajala Empty, a pozivalac Empty cita kao
' 'faktura nema dogadjaja'. Izgubljena kolona FakturaID ili pad filtera time su
' operateru izgledali kao uredan prazan log -- a ovo je AUDIT trag, jedino cime
' se dokazuje sta je kome poslato. Isti razlog zbog kog
' CountStrongKeyReadyBankaImport ne guta gresku (AUD-014).
'
' Nedostajucu TABELU ovaj EH ne bi ni video: GetTableData je za nju vracao
' Empty, pa se izlazilo rano. Nju hvata RequireTable iznad.
'
' Empty od sada znaci TACNO jedno: citanje je uspelo i redova nema.
'
' Err se cita PRE LogErr-a -- LogErr ume da obrise Err, pa bi Raise ispod njega
' dizao gresku bez broja i opisa (v. provera MRTAV_LOG u vba_check).
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC
    GetSEFEventsForFaktura = Empty
    Err.Raise errNum, errSrc, errDesc
End Function

' AUD-032b: razduzi TACNO ONU submisiju koju korektivni resubmit zamenjuje.
' Vraca True ako je red stvarno razduzen.
'
' Zasto uopste treba: status refresh NAMERNO ne dira submission red (poziv
' SaveSEFSubmissionResult_Row u modSEFStatusSync je zakomentarisan, da se podaci
' o originalnom submit HTTP pozivu ne prepisu podacima iz status upita). Zato
' faktura koju je SEF odbio TEK NA REFRESH-u zadrzava submission red u statusu
' SENT, pa `HasSuccessfulSEFSubmission` (fail-closed duplicate guard, AUD-031d)
' obara i onaj resubmit koji je `PrepareRejectedInvoiceForResubmit` upravo
' pripremio -- dokumentovan tok je time bio mrtav.
'
' Upisuje se SEF_SUB_REJECTED, dakle tacno ono sto bi `SaveSEFSubmissionResult_Row`
' upisao za `response.Rejected` -- ne izmislja se novo stanje, primenjuje se
' postojece mapiranje sistema.
'
' NAMERNO uzak zahvat -- samo prosledjeni (poslednji) red, uz proveru vlasnistva:
'   * SEF_SUB_ACCEPTED se ne dira (prihvacena submisija uz odbijen workflow je
'     neuskladjen podatak -> duplicate guard s pravom nastavlja da blokira),
'   * stariji SENT redovi iste fakture se ne diraju (ne prepisujemo istoriju o
'     kojoj nista ne znamo). Ako takav red postoji, resubmit ostaje blokiran i
'     pozivalac to prijavi kao rucnu proveru -- fail-closed.
'
' `_Row` = pozivalac obezbedjuje transakciju (snapshot nad TBL_SEF_SUBMISSION).
Public Function DischargeSEFSubmission_Row(ByVal submissionID As String, _
                                           ByVal fakturaID As String, _
                                           ByVal expectedDocumentId As String) As Boolean

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.DischargeSEFSubmission_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    ' Nema poslednje submisije -> nema sta da se razduzi. Pozivalac posle ovoga
    ' ionako proverava duplicate guard, pa se tisina ovde ne pretvara u uspeh.
    If Len(Trim$(submissionID)) = 0 Then Exit Function

    RequireSEFSubmissionSchema SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, True)

    Dim ownerFakturaID As String
    ownerFakturaID = Trim$(CStr(LookupValue(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, "FakturaID")))

    If ownerFakturaID <> Trim$(fakturaID) Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Submission " & submissionID & " belongs to faktura " & ownerFakturaID & _
                  ", not " & fakturaID & "."
    End If

    ' Fiskalni lineage: kad OBA identiteta postoje, moraju da se poklope.
    ' `SEFSubmissionIDLast` je pokazivac koji moze da bude zastareo (npr. red
    ' zaostao iz ranijeg pokusaja), pa bi se bez ove provere kao "odbijen" mogao
    ' obeleziti pogresan pokusaj -- a to je zapis o predaji poreskom organu.
    ' Ako se ne poklapaju, ne pretpostavljamo koji je tacan: pucamo i trazimo
    ' rucnu proveru (pozivalac je u TX-u, pa se sve vraca).
    Dim submissionDocumentId As String
    submissionDocumentId = Trim$(CStr(LookupValue(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, "SEFDocumentId")))

    If Len(submissionDocumentId) > 0 And Len(Trim$(expectedDocumentId)) > 0 Then
        If submissionDocumentId <> Trim$(expectedDocumentId) Then
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Submission " & submissionID & " points to SEF document " & _
                      submissionDocumentId & ", but faktura " & fakturaID & _
                      " carries " & Trim$(expectedDocumentId) & ". Manual review required."
        End If
    End If

    Dim currentStatus As String
    currentStatus = Trim$(CStr(LookupValue(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, "SubmissionStatus")))

    If currentStatus <> SEF_SUB_SENT Then Exit Function

    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "SubmissionStatus", SEF_SUB_REJECTED, SRC

    DischargeSEFSubmission_Row = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Function

Public Function HasSuccessfulSEFSubmission(ByVal fakturaID As String) As Boolean
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.HasSuccessfulSEFSubmission"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    ' AUD-031d: this duplicate guard must be fail-CLOSED. Require the submission
    ' schema FIRST so a missing/absent tblSEFSubmission raises here instead of
    ' GetTableData returning Empty and the guard silently reporting "no prior
    ' submission" (-> double send). Read the table directly (not via the
    ' fail-soft GetSEFSubmissionsForFaktura, whose EH returns Empty) so a
    ' genuine read error also propagates. After the schema check, IsEmpty means
    ' the table legitimately has zero rows.
    RequireSEFSubmissionSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_SEF_SUBMISSION)

    If IsEmpty(data) Then
        HasSuccessfulSEFSubmission = False
        Exit Function
    End If

    Dim colFakturaID As Long
    Dim colStatus As Long
    colFakturaID = RequireColumnIndex(TBL_SEF_SUBMISSION, "FakturaID", SRC)
    colStatus = RequireColumnIndex(TBL_SEF_SUBMISSION, "SubmissionStatus", SRC)

    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakturaID))) = Trim$(fakturaID) Then
            Select Case Trim$(CStr(data(i, colStatus)))
                Case SEF_SUB_SENT, SEF_SUB_ACCEPTED
                    HasSuccessfulSEFSubmission = True
                    Exit Function
            End Select
        End If
    Next i

    HasSuccessfulSEFSubmission = False
    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Function

Public Function GetLastSEFSubmissionStatus(ByVal fakturaID As String) As String
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetLastSEFSubmissionStatus"

    Dim data As Variant
    data = GetSEFSubmissionsForFaktura(fakturaID)

    If IsEmpty(data) Then
        GetLastSEFSubmissionStatus = ""
        Exit Function
    End If

    Dim colCreatedAt As Long
    Dim colStatus As Long

    colCreatedAt = RequireColumnIndex(TBL_SEF_SUBMISSION, "CreatedAt", SRC)
    colStatus = RequireColumnIndex(TBL_SEF_SUBMISSION, "SubmissionStatus", SRC)

    data = SortArray(data, colCreatedAt, False)

    GetLastSEFSubmissionStatus = Trim$(CStr(data(1, colStatus)))
    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If errNum = 0 Then errNum = ERR_SEF_STATE

    Err.Raise errNum, SRC, errDesc
End Function

Public Function GetSubmissionRequestBody(ByVal submissionID As String) As String
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSubmissionRequestBody"

    If Len(Trim$(submissionID)) = 0 Then Exit Function

    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFSubmissionID", SRC
    RequireColumnIndex TBL_SEF_SUBMISSION, "RequestBody", SRC

    Dim v As Variant
    v = LookupValue(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, "RequestBody")

    If IsEmpty(v) Or IsNull(v) Then
        GetSubmissionRequestBody = ""
    Else
        GetSubmissionRequestBody = CStr(v)
    End If

    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim ehErrNum As Long
    Dim ehErrDesc As String

    ehErrNum = Err.Number
    ehErrDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If ehErrNum = 0 Then ehErrNum = ERR_SEF_STATE

    Err.Raise ehErrNum, SRC, ehErrDesc
End Function

Public Function GetSubmissionPayloadHash(ByVal submissionID As String) As String
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSubmissionPayloadHash"

    If Len(Trim$(submissionID)) = 0 Then Exit Function

    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFSubmissionID", SRC
    RequireColumnIndex TBL_SEF_SUBMISSION, "PayloadHash", SRC

    Dim v As Variant
    v = LookupValue(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, "PayloadHash")

    If IsEmpty(v) Or IsNull(v) Then
        GetSubmissionPayloadHash = ""
    Else
        GetSubmissionPayloadHash = Trim$(CStr(v))
    End If

    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim ehErrNum As Long
    Dim ehErrDesc As String

    ehErrNum = Err.Number
    ehErrDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If ehErrNum = 0 Then ehErrNum = ERR_SEF_STATE

    Err.Raise ehErrNum, SRC, ehErrDesc
End Function

Public Sub ClearFakturaLastSubmission_Row(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.ClearFakturaLastSubmission_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", SRC
    RequireColumnIndex TBL_FAKTURE, "SEFSubmissionIDLast", SRC
    RequireColumnIndex TBL_FAKTURE, "SEFDocumentId", SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFSubmissionIDLast", "", SRC

    ' AUD-031: a corrected resubmit must not carry the previous attempt's
    ' SEFDocumentId. Clear it so a stale docId can never drive a later status
    ' refresh / cancel / storno against the wrong SEF document; the next
    ' successful submission writes the fresh docId.
    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFDocumentId", "", SRC

    Exit Sub

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim ehErrNum As Long
    Dim ehErrDesc As String

    ehErrNum = Err.Number
    ehErrDesc = Err.description

    LogErr SRC
    On Error Resume Next
    On Error GoTo 0

    If ehErrNum = 0 Then ehErrNum = ERR_SEF_STATE

    Err.Raise ehErrNum, SRC, ehErrDesc
End Sub

Private Sub RequireFaktureSEFSchema(ByVal sourceName As String)
    RequireColumnIndex TBL_FAKTURE, "FakturaID", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFWorkflowState", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFStatus", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFDocumentId", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFLastErrorCode", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFLastErrorMessage", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFPayloadHash", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFSubmissionIDLast", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFVersionNo", sourceName
    RequireColumnIndex TBL_FAKTURE, "PoslatNaSEF", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFSentAt", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFLastSyncAt", sourceName
End Sub

Private Sub RequireSEFSubmissionSchema(ByVal sourceName As String)
    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFSubmissionID", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "FakturaID", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "VersionNo", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "WorkflowStateAtSubmit", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "CreatedAt", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "SubmittedAt", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "SubmissionStatus", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "PayloadHash", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "RequestFormat", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "RequestBody", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ResponseBody", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "HttpStatus", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ApiStatus", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "CorrelationId", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFDocumentId", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ErrorCode", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ErrorMessage", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "OperatorName", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "Stornirano", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "FinishedAt", sourceName
End Sub

Private Sub RequireSEFEventLogSchema(ByVal sourceName As String)
    RequireColumnIndex TBL_SEF_EVENT_LOG, "SEFEventID", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "FakturaID", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "SEFSubmissionID", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "EventTime", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "EventType", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "Message", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "Details", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "OperatorName", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "Stornirano", sourceName
End Sub


Private Function GetFakturaSEFFieldText(ByVal fakturaID As String, _
                                        ByVal fieldName As String, _
                                        ByVal sourceName As String) As String
    On Error GoTo EH

    If Len(Trim$(fakturaID)) = 0 Then Exit Function

    RequireColumnIndex TBL_FAKTURE, "FakturaID", sourceName
    RequireColumnIndex TBL_FAKTURE, fieldName, sourceName

    Dim v As Variant
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, fieldName)

    If IsEmpty(v) Or IsNull(v) Then
        GetFakturaSEFFieldText = ""
        Exit Function
    End If
    
    ' Defensive guard for legacy rows where Excel may have stored a numeric
    ' SEFDocumentId as Double before the Text format migration. CStr on
    ' a Double can produce scientific notation for large values, which
    ' would corrupt the ID downstream. Format$(raw, "0") preserves all
    ' integer digits without scientific notation or decimal artifacts.
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger
            GetFakturaSEFFieldText = Format$(v, "0")
        Case Else
            GetFakturaSEFFieldText = Trim$(CStr(v))
    End Select
    
    Exit Function

EH:
    ' AUD-054: greska se hvata PRE LogErr-a. LogError interno radi
    ' "On Error Resume Next" / "On Error GoTo 0", a svaka On Error naredba
    ' resetuje Err objekat -- zatecno "Err.Raise Err.Number" je time postajalo
    ' "Err.Raise 0", pa se greska GUTALA umesto da se propagira pozivaocu.
    ' RF-22 se oslanja bas na ovu propagaciju (rollback TX-a, fail-closed kapije).
    Dim ehErrNum As Long
    Dim ehErrDesc As String

    ehErrNum = Err.Number
    ehErrDesc = Err.description

    LogErr sourceName
    On Error Resume Next
    On Error GoTo 0

    If ehErrNum = 0 Then ehErrNum = ERR_SEF_STATE

    Err.Raise ehErrNum, sourceName, ehErrDesc
End Function

