Attribute VB_Name = "modStornoZurnal"
Option Explicit

' ============================================================
' modStornoZurnal - append-only cell-level zurnal storno operacija za LOSSLESS
' "Vrati storno" (pravi inverz storna).
'
' Ambient op-kontekst: storno primitiva (StornoOtkup / StornoOMKoopByBrDok) pozove
' BeginStornoOp na ulazu; mutacione tacke (MarkRowStornirano preko StornoOtkup,
' StornoAmbalazaByDokument, ResetNovacOtkupLink) usput JournalCell-uju STARU
' vrednost PRE mutacije; EndStornoOp zatvara. UndoOperation_TX vrati svaku celiju
' na staru vrednost i cilja SAMO tu operaciju -> resava reused-broj rizik (stari
' undo je vracao SVE generacije istog broja) i vraca tblNovac.OtkupID koji je
' storno nepovratno obrisao.
'
' Opseg (faza 1): Otkup + Revers (OM-koop). Chain dokumenti (prijemnica/faktura/
' paleta/prerada) se instrumentiraju u kasnijem PR-u; njihov undo se i dalje odbija.
'
' Zurnal upisi teku UNUTAR storno transakcije (entry _TX snapshot-uju TBL_STORNO_ZURNAL)
' pa rollback storna povlaci i zurnal (nema orphan redova).
' ============================================================

Private Const MOD_NAME As String = "modStornoZurnal"
Private Const ERR_SZ_BASE As Long = vbObjectError + 3100

Private mOpID As String
Private mActive As Boolean
Private mDocType As String
Private mBroj As String

' ============================================================
' AMBIENT OP KONTEKST
' ============================================================
' Vraca True ako je OVAJ poziv otvorio operaciju (pa je on i zatvara). Ako je op
' vec aktivan (ugnjezden poziv), vraca False i cell-ovi se pridruzuju postojecoj.
Public Function BeginStornoOp(ByVal docType As String, ByVal broj As String) As Boolean
    On Error GoTo EH
    If mActive Then Exit Function
    mOpID = GetNextID(TBL_STORNO_ZURNAL, COL_SZ_OP_ID, "SOP-")
    If Len(mOpID) = 0 Then mOpID = "SOP-1"
    mDocType = docType: mBroj = broj
    mActive = True
    BeginStornoOp = True
    Exit Function
EH:
    mActive = False: mOpID = ""
    LogErr MOD_NAME & ".BeginStornoOp"
End Function

Public Sub EndStornoOp(ByVal owns As Boolean)
    If Not owns Then Exit Sub
    mActive = False: mOpID = "": mDocType = "": mBroj = ""
End Sub

' Force-reset op-konteksta (za EH putanje entry _TX-ova) -> nikad ne ostavi op
' otvoren posle greske (inace bi sledeci storno pisao u pogresnu/mrtvu op).
Public Sub AbortStornoOp()
    mActive = False: mOpID = "": mDocType = "": mBroj = ""
End Sub

Public Function StornoOpActive() As Boolean
    StornoOpActive = mActive
End Function

' JournalCell: zabelezi (tabela, RowID(PK), kolona, STARA vrednost) za tekucu op.
' No-op ako operacija nije aktivna (backward-compatible za ne-instrumentirane putanje).
' FAIL-CLOSED: ako upis padne dok je op aktivan, DIZE gresku -> storno primitiva to
' reraise-uje -> _TX rollback-uje i podatke i (parcijalni) zurnal. Bez tihog gubitka
' upisa: lossless undo zavisi od KOMPLETNOG zurnala.
Public Sub JournalCell(ByVal tbl As String, ByVal rowID As String, _
                       ByVal col As String, ByVal oldVal As Variant)
    Const SRC As String = MOD_NAME & ".JournalCell"
    If Not mActive Then Exit Sub
    Dim zid As String: zid = GetNextID(TBL_STORNO_ZURNAL, COL_SZ_ID, "ZUR-")
    ' Redosled MORA pratiti EnsureStornoZurnalSchemaCore:
    ' ZurnalID, OperationID, Timestamp, DocType, Broj, Tabela, RowID, Kolona, StaraVrednost
    If AppendRow(TBL_STORNO_ZURNAL, Array(zid, mOpID, Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
        mDocType, mBroj, tbl, CStr(rowID), col, CStr(oldVal))) = 0 Then
        Err.Raise ERR_SZ_BASE + 10, SRC, "Zurnal upis nije uspeo (" & tbl & "." & col & _
            ") -> storno se prekida (lossless garancija)."
    End If
End Sub

' ============================================================
' UNDO - pravi inverz jedne operacije (vrati svaku celiju na staru vrednost).
' ============================================================
Public Function UndoOperation_TX(ByVal opID As String) As Boolean
    Const SRC As String = MOD_NAME & ".UndoOperation_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    opID = Trim$(opID)
    If Len(opID) = 0 Then Err.Raise ERR_SZ_BASE + 1, SRC, "OperationID je obavezan."

    Dim data As Variant: data = GetTableData(TBL_STORNO_ZURNAL)
    If IsEmpty(data) Then Err.Raise ERR_SZ_BASE + 2, SRC, "Storno zurnal je prazan."

    Dim cOp As Long, cTbl As Long, cRow As Long, cCol As Long, cOld As Long, cDoc As Long, cBroj As Long
    cOp = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_OP_ID)
    cTbl = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_TABELA)
    cRow = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_ROWID)
    cCol = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_KOLONA)
    cOld = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_STARA)
    cDoc = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_DOCTYPE)
    cBroj = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_BROJ)
    If cOp = 0 Then Err.Raise ERR_SZ_BASE + 6, SRC, "Zurnal sema nije kompletna."

    Dim rows As Collection: Set rows = New Collection
    Dim docType As String, broj As String
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOp))) = opID Then
            rows.Add Array(CStr(data(i, cTbl)), CStr(data(i, cRow)), CStr(data(i, cCol)), CStr(data(i, cOld)))
            docType = CStr(data(i, cDoc)): broj = CStr(data(i, cBroj))
        End If
    Next i
    If rows.count = 0 Then Err.Raise ERR_SZ_BASE + 3, SRC, "Operacija nije nadjena: " & opID

    ' ZAJEDNICKA garda (isti guard i za legacy i za zurnal put): Otkup active-dup +
    ' mrtav-roditelj; revers (OM) active-dup ambalaze (#134 garda - zurnal put ju je
    ' ranije zaobilazio). Bez ovoga journaled revers undo bi duplirao ledger.
    Dim gr As String: gr = UndoGuardReason(docType, broj)
    If Len(gr) > 0 Then Err.Raise ERR_SZ_BASE + 4, SRC, gr

    ' Pre-validacija (SVE-ILI-NISTA): svaka celija mora biti restore-abilna PRE nego
    ' sto diramo ijedan red -> undo je atomican i ne laze uspeh nad delimicnim vracanjem.
    For i = 1 To rows.count
        Dim vt As String: vt = CStr(rows(i)(0))
        Dim vpk As String: vpk = PkColForTable(vt)
        If Len(vpk) = 0 Then Err.Raise ERR_SZ_BASE + 7, SRC, "Nepodrzana tabela u zurnalu: " & vt
        If GetColumnIndex(vt, CStr(rows(i)(2))) = 0 Then _
            Err.Raise ERR_SZ_BASE + 8, SRC, "Kolona ne postoji: " & vt & "." & CStr(rows(i)(2))
        If FindRowIndexByKey(vt, vpk, CStr(rows(i)(1))) = 0 Then _
            Err.Raise ERR_SZ_BASE + 9, SRC, "Ciljni red ne postoji: " & vt & " " & CStr(rows(i)(1))
    Next i

    Set tx = New clsTransaction
    tx.BeginTx
    Dim snapped As Object: Set snapped = CreateObject("Scripting.Dictionary")
    snapped.CompareMode = vbTextCompare
    For i = 1 To rows.count
        Dim t As String: t = CStr(rows(i)(0))
        If Len(t) > 0 And Not snapped.Exists(t) Then tx.AddTableSnapshot t: snapped(t) = True
    Next i
    For i = 1 To rows.count
        RestoreCell CStr(rows(i)(0)), CStr(rows(i)(1)), CStr(rows(i)(2)), CStr(rows(i)(3)), SRC
    Next i
    tx.CommitTx: Set tx = Nothing

    UndoOperation_TX = True
    Monitor_Event eventType:="STORNO_UNDO_OP", severity:="INFO", _
        message:="Vrati storno (op " & opID & "): " & docType & " " & broj & " -> " & _
                 rows.count & " celija vraceno.", _
        moduleName:=MOD_NAME, procedureName:="UndoOperation_TX", _
        entityType:=docType, entityID:=broj, correlationId:=opID
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    UndoOperation_TX = False
End Function

Private Sub RestoreCell(ByVal tbl As String, ByVal rowID As String, _
                        ByVal col As String, ByVal oldVal As String, ByVal SRC As String)
    Dim pkCol As String: pkCol = PkColForTable(tbl)
    If Len(pkCol) = 0 Then Exit Sub
    Dim ri As Long: ri = FindRowIndexByKey(tbl, pkCol, rowID)
    If ri > 0 Then RequireUpdateCell tbl, ri, col, oldVal, SRC
End Sub

' PK kolona po tabeli (opseg faze 1: Otkup + Revers -> tblOtkup/Ambalaza/Novac).
Private Function PkColForTable(ByVal tbl As String) As String
    Select Case tbl
        Case TBL_OTKUP: PkColForTable = COL_OTK_ID
        Case TBL_AMBALAZA: PkColForTable = COL_AMB_ID
        Case TBL_NOVAC: PkColForTable = COL_NOV_ID
    End Select
End Function

Private Function FindRowIndexByKey(ByVal tbl As String, ByVal keyCol As String, _
                                   ByVal keyVal As String) As Long
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim c As Long: c = GetColumnIndex(tbl, keyCol)
    If c = 0 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, c))) = Trim$(keyVal) Then FindRowIndexByKey = i: Exit Function
    Next i
End Function

' Najskoriji OperationID za (docType, broj) - za undo-by-broj iz UI-a.
Public Function LatestOpFor(ByVal docType As String, ByVal broj As String) As String
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_STORNO_ZURNAL)
    If IsEmpty(data) Then Exit Function
    Dim cOp As Long, cDoc As Long, cBroj As Long
    cOp = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_OP_ID)
    cDoc = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_DOCTYPE)
    cBroj = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_BROJ)
    If cOp = 0 Then Exit Function
    Dim i As Long, best As String, bestN As Long
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, cDoc))), docType, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(data(i, cBroj))), broj, vbTextCompare) = 0 Then
            Dim opv As String: opv = Trim$(CStr(data(i, cOp)))
            Dim numv As Long: numv = OpNum(opv)
            If numv >= bestN Then bestN = numv: best = opv
        End If
    Next i
    LatestOpFor = best
    Exit Function
EH:
    LogErr MOD_NAME & ".LatestOpFor"
End Function

Private Function OpNum(ByVal opID As String) As Long
    On Error Resume Next
    If Left$(opID, 4) = "SOP-" Then OpNum = CLng(Mid$(opID, 5))
End Function
