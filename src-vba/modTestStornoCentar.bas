Attribute VB_Name = "modTestStornoCentar"
Option Explicit

' ============================================================
' modTestStornoCentar - automatski regres-testovi za Storno centar / Faza 7.
' Pokriva korektnosne dodatke: Guard C (BlockStornoDriftReason) i sledljivost
' (StampIspravkaTrace). FindSingleActiveRow ima svoj test u modDokumentInvariant.
'
' SVAKI test je ROLLBACK-SAFE: clsTransaction snapshot -> seed fixture (SVT- prefiks)
' -> assert -> RollbackTx (fixture NE ostaje u podacima).
' Pokretanje: Alt+F8 -> Test_StornoCentar_All. Rezultat u Immediate (Ctrl+G).
' Napomena: pre pokretanja EnsureRuntimeSchema (da trace kolone postoje), inace
' StampIspravkaTrace test pada (guarded no-op).
' ============================================================

Private mPass As Long
Private mFail As Long

Public Sub Test_StornoCentar_All()
    mPass = 0: mFail = 0
    Test_StampIspravkaTrace_Auto
    Test_BlockStornoDriftReason_Auto
    Debug.Print "=== StornoCentar: " & mPass & " OK, " & mFail & " FAIL ==="
End Sub

' Sledljivost: novi red nosi IspravkaOd + CorrectionID; stari (storniran) nosi ZamenjenSa.
Public Sub Test_StampIspravkaTrace_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_STORNIRANO), _
              Array("SVT-ST-OLD", "SVT-ST-B1", "I", "Da")          ' stari, storniran
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA), _
              Array("SVT-ST-NEW", "SVT-ST-B2", "I")                ' novi, aktivan

    StampIspravkaTrace TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-ST-B2", "SVT-ST-B1", "SVT-CID-1"

    TcChk NzS(LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-ST-B2", COL_TRACE_ISPRAVKA_OD)) = "SVT-ST-B1", _
          "novi red IspravkaOd = stari broj"
    TcChk NzS(LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-ST-B2", COL_TRACE_CORRECTION_ID)) = "SVT-CID-1", _
          "novi red CorrectionID upisan"
    TcChk NzS(LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-ST-B1", COL_TRACE_ZAMENJEN_SA)) = "SVT-ST-B2", _
          "stari (storniran) red ZamenjenSa = novi broj"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StampIspravkaTrace_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Guard C: blok-storno nad ZIVOM otpremnicom -> drift (odbij); mrtva/PONISTENJE/unbound -> dozvoljeno.
Public Sub Test_BlockStornoDriftReason_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTPREMNICA
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_KLASA), _
              Array("SVT-DR-OTP", "SVT-DR-O1", "I")                ' aktivna otpremnica
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_OTPREMNICA_ID, COL_OTK_BR_DOK), _
              Array("SVT-DR-BLK", "SVT-DR-OTP", "SVT-DR-BD")       ' blok vezan za nju

    Dim ids As Collection: Set ids = New Collection: ids.Add "SVT-DR-BLK"
    TcChk Len(BlockStornoDriftReason(FLOW_DOC_PRIJEMNICA, SV_MODE_DUPLI, ids)) > 0, _
          "DUPLI prijemnice + blok na ZIVOJ otpremnici -> drift (odbij)"
    TcChk Len(BlockStornoDriftReason(FLOW_DOC_PRIJEMNICA, SV_MODE_PONISTENJE, ids)) = 0, _
          "PONISTENJE -> dozvoljeno (roditelj umire u kaskadi)"

    ' storniraj otpremnicu -> mrtav roditelj -> DUPLI blok dozvoljen (nema zive da precenjuje)
    Dim c As Collection: Set c = FindRows(TBL_OTPREMNICA, COL_OTP_ID, "SVT-DR-OTP")
    If Not c Is Nothing Then If c.count > 0 Then UpdateCell TBL_OTPREMNICA, CLng(c(1)), COL_STORNIRANO, "Da"
    TcChk Len(BlockStornoDriftReason(FLOW_DOC_PRIJEMNICA, SV_MODE_DUPLI, ids)) = 0, _
          "mrtva otpremnica -> DUPLI blok dozvoljen"

    Dim ids2 As Collection: Set ids2 = New Collection: ids2.Add "SVT-DR-NONE"
    TcChk Len(BlockStornoDriftReason(FLOW_DOC_PRIJEMNICA, SV_MODE_DUPLI, ids2)) = 0, _
          "nepoznat/unbound blok -> dozvoljen"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_BlockStornoDriftReason_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' ============================================================
' HELPERS
' ============================================================
Private Sub TcSeedRow(ByVal tbl As String, ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject: Set lo = GetTable(tbl)
    If lo Is Nothing Then Exit Sub
    Dim nr As ListRow: Set nr = lo.ListRows.Add
    Dim i As Long, ci As Long
    For i = LBound(cols) To UBound(cols)
        ci = GetColumnIndex(tbl, CStr(cols(i)))
        If ci > 0 Then nr.Range.cells(1, ci).value = vals(i)
    Next i
End Sub

Private Sub TcChk(ByVal cond As Boolean, ByVal nm As String)
    If cond Then
        mPass = mPass + 1
        Debug.Print "OK   " & nm
    Else
        mFail = mFail + 1
        Debug.Print "FAIL " & nm
    End If
End Sub

Private Function NzS(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzS = "" Else NzS = Trim$(CStr(v))
End Function
