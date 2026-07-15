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
    Test_DocIsIssued_Auto
    Test_OtkupBlockDeadParent_Auto
    Test_BuildStornoImpact_Auto
    Test_GetActiveDocumentsForStorno_Auto
    Test_StornoSelectedBlocks_Auto
    Test_GetNedovrseno_Auto
    Debug.Print "=== StornoCentar: " & mPass & " OK, " & mFail & " FAIL ==="
End Sub

' #4 objedinjeni recovery: GetNedovrseno nosi CorrectionID i DEDUPLIKUJE osirocene
' protiv PENDING context-a (isti poslovni broj se ne prikazuje dvaput).
Public Sub Test_GetNedovrseno_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_STORNO_VEZE

    ' aktivna prijemnica cija zbirna ne postoji -> osirocena
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_BROJ_ZBIRNE), _
              Array("SVT-ND-PID", "SVT-ND-P1", "I", "SVT-ND-ZDEAD")

    ' pre context-a: SVT-ND-P1 vidljiv kao osiroce, bez CorrectionID
    TcChk NedRefCount("SVT-ND-P1") = 1, "osirocena prijemnica -> 1 red u Nedovrseno"
    TcChk Len(NedRefCorrectionID("SVT-ND-P1")) = 0, "osirocen red nema CorrectionID"

    ' PENDING context za ISTI broj (RESI_KASNIJE, NeedsRecovery=Da)
    TcSeedRow TBL_STORNO_VEZE, Array(COL_SV_ID, COL_SV_MODE, COL_SV_STATUS, COL_SV_OLD_DOCTYPE, _
              COL_SV_OLD_BROJ, COL_SV_NEEDS_RECOVERY), _
              Array("SVT-ND-COR", SV_MODE_RESI_KASNIJE, SV_STATUS_PENDING, FLOW_DOC_PRIJEMNICA, _
              "SVT-ND-P1", "Da")

    ' posle: i dalje 1 red (dedup), ali sada nosi CorrectionID (context "pobedi")
    TcChk NedRefCount("SVT-ND-P1") = 1, "context + osiroce isti broj -> deduplikovano na 1 red"
    TcChk NedRefCorrectionID("SVT-ND-P1") = "SVT-ND-COR", "dedup red nosi CorrectionID iz context-a"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_GetNedovrseno_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

Private Function NedRefCount(ByVal ref As String) As Long
    Dim c As Collection: Set c = GetNedovrseno()
    If c Is Nothing Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To c.count
        If StrComp(Trim$(CStr(c(i)("ref"))), ref, vbTextCompare) = 0 Then n = n + 1
    Next i
    NedRefCount = n
End Function

Private Function NedRefCorrectionID(ByVal ref As String) As String
    Dim c As Collection: Set c = GetNedovrseno()
    If c Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To c.count
        If StrComp(Trim$(CStr(c(i)("ref"))), ref, vbTextCompare) = 0 Then
            NedRefCorrectionID = CStr(c(i)("correctionID"))
            Exit Function
        End If
    Next i
End Function

' Undo garda: blok sa storniranim roditeljem -> siroce (odbij); ziv roditelj/unbound -> ok.
Public Sub Test_OtkupBlockDeadParent_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTPREMNICA
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_KLASA), _
              Array("SVT-DP-OTP-A", "SVT-DP-OA", "I")                          ' aktivna otpremnica
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_KLASA, COL_STORNIRANO), _
              Array("SVT-DP-OTP-D", "SVT-DP-OD", "I", "Da")                    ' stornirana otpremnica
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_OTPREMNICA_ID, COL_STORNIRANO), _
              Array("SVT-DP-K1", "SVT-DP-B1", "SVT-DP-OTP-A", "Da")            ' ziv roditelj
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_OTPREMNICA_ID, COL_STORNIRANO), _
              Array("SVT-DP-K2", "SVT-DP-B2", "SVT-DP-OTP-D", "Da")            ' mrtav roditelj
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_OTPREMNICA_ID, COL_STORNIRANO), _
              Array("SVT-DP-K3", "SVT-DP-B3", "", "Da")                        ' unbound

    TcChk Len(OtkupBlockDeadParent("SVT-DP-B1")) = 0, "blok sa ZIVOM otpremnicom -> undo dozvoljen"
    TcChk Len(OtkupBlockDeadParent("SVT-DP-B2")) > 0, "blok sa STORNIRANOM otpremnicom -> mrtav roditelj (odbij)"
    TcChk Len(OtkupBlockDeadParent("SVT-DP-B3")) = 0, "unbound blok -> undo dozvoljen"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_OtkupBlockDeadParent_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Impact agregator: header + summary iz stvarnih (seed) redova.
Public Sub Test_BuildStornoImpact_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_KUPAC, COL_PRJ_KOLICINA, COL_PRJ_BROJ_ZBIRNE), _
              Array("SVT-BI-ID", "SVT-BI-P1", "I", "SVT-BI-KUP", 123, "SVT-BI-Z1")

    Dim m As Object: Set m = BuildStornoImpact(FLOW_DOC_PRIJEMNICA, "SVT-BI-P1")
    Dim h As Object: Set h = m("header")
    Dim sm As Object: Set sm = m("summary")
    TcChk NzS(h("partnerID")) = "SVT-BI-KUP", "impact header partnerID iz reda"
    TcChk Val(NzS(h("kolicina"))) = 123, "impact header kolicina = 123"
    TcChk CLng(sm("blockCount")) = 0, "impact summary blockCount = 0"
    TcChk CLng(sm("paleteCount")) = 0, "impact summary paleteCount = 0"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_BuildStornoImpact_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Browse: distinct po broju (2 klase -> 1x), filter tip, iskljuci stornirano.
Public Sub Test_GetActiveDocumentsForStorno_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA), Array("SVT-GA-1", "SVT-GA-P1", "I")
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA), Array("SVT-GA-2", "SVT-GA-P1", "II")   ' isti broj
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA), Array("SVT-GA-3", "SVT-GA-P2", "I")
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_STORNIRANO), Array("SVT-GA-4", "SVT-GA-P3", "I", "Da")

    Dim c As Collection: Set c = GetActiveDocumentsForStorno("Prijemnica", "SVT-GA-")
    Dim nP1 As Long, nP2 As Long, nP3 As Long, i As Long
    If Not c Is Nothing Then
        For i = 1 To c.count
            Dim br As String: br = NzS(c(i)(1))
            If br = "SVT-GA-P1" Then nP1 = nP1 + 1
            If br = "SVT-GA-P2" Then nP2 = nP2 + 1
            If br = "SVT-GA-P3" Then nP3 = nP3 + 1
        Next i
    End If
    TcChk nP1 = 1, "distinct po broju: P1 (2 klase) -> 1x"
    TcChk nP2 = 1, "P2 aktivan -> 1x"
    TcChk nP3 = 0, "stornirana P3 -> iskljucena"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_GetActiveDocumentsForStorno_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' StornoSelectedBlocks_TX: atomican storno N blokova; -1 + rollback na los ID.
Public Sub Test_StornoSelectedBlocks_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KOLICINA, COL_OTK_KLASA), Array("SVT-SB-1", "SVT-SB-D1", 10, "I")
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KOLICINA, COL_OTK_KLASA), Array("SVT-SB-2", "SVT-SB-D2", 20, "I")

    Dim good As Collection: Set good = New Collection: good.Add "SVT-SB-1": good.Add "SVT-SB-2"
    TcChk StornoSelectedBlocks_TX(good) = 2, "storno 2 bloka -> vraca 2"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-SB-1", COL_STORNIRANO))) = "DA", "blok 1 storniran"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-SB-2", COL_STORNIRANO))) = "DA", "blok 2 storniran"

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KOLICINA, COL_OTK_KLASA), Array("SVT-SB-3", "SVT-SB-D3", 30, "I")
    Dim mix As Collection: Set mix = New Collection: mix.Add "SVT-SB-3": mix.Add "SVT-SB-BAD"
    TcChk StornoSelectedBlocks_TX(mix) = -1, "los ID -> -1 (rollback)"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-SB-3", COL_STORNIRANO))) <> "DA", "atomicnost: blok 3 ostao AKTIVAN"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoSelectedBlocks_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' IzdatoStatus gate: prazno/IZDATO -> izdato; DRAFT -> nije izdato.
Public Sub Test_DocIsIssued_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA), _
              Array("SVT-IZ-1", "SVT-IZ-EMPTY", "I")                 ' prazan IzdatoStatus
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_TRACE_IZDATO_STATUS), _
              Array("SVT-IZ-2", "SVT-IZ-DRAFT", "I", IZDATO_DRAFT)
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_TRACE_IZDATO_STATUS), _
              Array("SVT-IZ-3", "SVT-IZ-IZD", "I", IZDATO_IZDATO)

    TcChk DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-IZ-EMPTY") = True, "prazan IzdatoStatus -> izdato"
    TcChk DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-IZ-DRAFT") = False, "DRAFT -> nije izdato"
    TcChk DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-IZ-IZD") = True, "IZDATO -> izdato"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_DocIsIssued_Auto GRESKA: " & Err.description: mFail = mFail + 1
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
