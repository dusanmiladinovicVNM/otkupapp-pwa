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
    Test_UndoReverseGuard_Auto
    Test_ZbirnaRecalcInPlace_Auto
    Test_PonistenjePrijemniceKaskada_Auto
    Test_StornoJournalUndo_Auto
    Test_StornoJournalDualClass_Auto
    Test_StornoJournalReversGuard_Auto
    Test_StornoJournalUndoValidation_Auto
    Test_StornoJournalDrift_Auto
    Test_StornoJournalPartialClass_Auto
    Test_StornoJournalMixedOp_Auto
    Test_StornoJournalEmptyBrDok_Auto
    Test_StornoJournalReusedBroj_Auto
    Test_StornoJournalDeadParentOtherGen_Auto
    Test_StornoJournalEmptyBrDokUndo_Auto
    Test_ImpactHeaderSum_Auto
    Debug.Print "=== StornoCentar: " & mPass & " OK, " & mFail & " FAIL ==="
End Sub

' #6: impact header kolicina = SUMA aktivnih Klasa I+II (ranije citao samo prvu klasu
' -> potceni dvoklasni dokument u uvidu pre potvrde). Storniran red se ne broji.
Public Sub Test_ImpactHeaderSum_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_KOLICINA), Array("SVT-IH-1", "SVT-IH-P", "I", 100)
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_KOLICINA), Array("SVT-IH-2", "SVT-IH-P", "II", 50)
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_KOLICINA, COL_STORNIRANO), Array("SVT-IH-3", "SVT-IH-P", "I", 999, "Da")

    Dim m As Object: Set m = BuildStornoImpact(FLOW_DOC_PRIJEMNICA, "SVT-IH-P")
    Dim h As Object: Set h = m("header")
    TcChk Val(NzS(h("kolicina"))) = 150, "impact header kolicina = suma aktivnih Klasa I+II (100+50), storniran izuzet"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_ImpactHeaderSum_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Operation-centric UI koren: reused poslovni broj -> undo STAROG op vraca STARU
' generaciju (ne najnoviju). Dokazuje da ciljanje po OperationID resava reused-broj.
Public Sub Test_StornoJournalReusedBroj_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP: tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC: tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-RB-A", "SVT-RB-OTK", "I")
    TcChk StornoOtkup_TX("SVT-RB-A") = True, "storno gen A -> True"
    Dim opA As String: opA = TcDistinctOpsForRow(TBL_OTKUP, "SVT-RB-A")
    ' druga generacija istog broja (nov aktivan red) -> storno
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-RB-B", "SVT-RB-OTK", "I")
    TcChk StornoOtkup_TX("SVT-RB-B") = True, "storno gen B (isti broj) -> True"

    ' undo STAROG op (A) mora vratiti A, a B ostaje storniran (ne najnoviji!)
    TcChk UndoOperation_TX(opA) = True, "undo STAROG op (A) -> True"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-RB-A", COL_STORNIRANO))) <> "DA", "gen A vracena (bas ta operacija)"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-RB-B", COL_STORNIRANO))) = "DA", "gen B ostaje stornirana (nije dirnut najnoviji)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalReusedBroj_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Dead-parent DRUGE generacije istog broja NE sme preblokirati undo bezbedne operacije
' (per-red OtkupBlockDeadParentByID, ne broj-level).
Public Sub Test_StornoJournalDeadParentOtherGen_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP: tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA: tx.AddTableSnapshot TBL_NOVAC: tx.AddTableSnapshot TBL_STORNO_ZURNAL

    ' gen A: VEC stornirana (bez zurnala), mrtav roditelj (stornirana otpremnica)
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_KLASA, COL_STORNIRANO), Array("SVT-DG-OTP", "SVT-DG-OB", "I", "Da")
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA, COL_OTK_OTPREMNICA_ID, COL_STORNIRANO), Array("SVT-DG-A", "SVT-DG-OTK", "I", "SVT-DG-OTP", "Da")
    ' gen B: unbound aktivna -> storno (journaled)
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-DG-B", "SVT-DG-OTK", "I")
    TcChk StornoOtkup_TX("SVT-DG-B") = True, "storno gen B (unbound) -> True"

    ' undo B mora PROCI iako gen A (isti broj) ima mrtvog roditelja
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-DG-OTK") = True, "undo B prolazi (mrtav roditelj je na DRUGOJ generaciji)"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-DG-B", COL_STORNIRANO))) <> "DA", "gen B vracena"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalDeadParentOtherGen_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Prazan BrDok end-to-end: unbound blok se moze VRATITI preko OperationID (broj nije potreban).
Public Sub Test_StornoJournalEmptyBrDokUndo_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP: tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC: tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-EU-1", "", "I")
    TcChk StornoOtkup_TX("SVT-EU-1") = True, "storno unbound -> True"
    Dim op As String: op = TcDistinctOpsForRow(TBL_OTKUP, "SVT-EU-1")
    TcChk Len(op) > 0, "unbound -> op zabelezen"
    TcChk UndoOperation_TX(op) = True, "undo unbound preko OperationID -> True (broj nije potreban)"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-EU-1", COL_STORNIRANO))) <> "DA", "unbound blok vracen"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalEmptyBrDokUndo_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' P1 drift: ako se posle storna novac re-linkuje na drugi otkup, undo NE gazi noviju
' vezu (optimistic-concurrency: trenutna vrednost != NovaVrednost -> odbij).
Public Sub Test_StornoJournalDrift_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP: tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC: tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-DR-OID", "SVT-DR-B", "I")
    TcSeedRow TBL_NOVAC, Array(COL_NOV_ID, COL_NOV_OTKUP_ID), Array("SVT-DR-NID", "SVT-DR-OID")
    TcChk StornoOtkup_TX("SVT-DR-OID") = True, "storno (drift setup) -> True"
    ' DRIFT: drugi tok re-linkuje isti novac red na drugi otkup
    Dim ri As Long: ri = TcRowIndex(TBL_NOVAC, COL_NOV_ID, "SVT-DR-NID")
    If ri > 0 Then UpdateCell TBL_NOVAC, ri, COL_NOV_OTKUP_ID, "SVT-DR-DRUGI"
    ' undo MORA biti odbijen (ne gazi noviju vezu)
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-DR-B") = False, "undo uz drift novca -> ODBIJEN"
    TcChk NzS(LookupValue(TBL_NOVAC, COL_NOV_ID, "SVT-DR-NID", COL_NOV_OTKUP_ID)) = "SVT-DR-DRUGI", "novija veza netaknuta"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalDrift_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' P2 partial-class: storno SAMO Klase I; Klasa II ostaje aktivna; undo Klase I mora
' PROCI (per (broj,klasa) guard, ne broj-level koji bi preblokirao).
Public Sub Test_StornoJournalPartialClass_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP: tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC: tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-PC-1", "SVT-PC-B", "I")
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-PC-2", "SVT-PC-B", "II")
    ' storniraj SAMO Klasu I (selektivno)
    Dim sel As Collection: Set sel = New Collection: sel.Add "SVT-PC-1"
    TcChk StornoSelectedBlocks_TX(sel) = 1, "selektivni storno Klase I -> 1"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-PC-2", COL_STORNIRANO))) <> "DA", "Klasa II ostaje aktivna"
    ' undo Klase I MORA proci iako je Klasa II aktivna (nije dup po (broj,klasa))
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-PC-B") = True, "undo Klase I uz aktivnu Klasu II -> PROLAZI"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-PC-1", COL_STORNIRANO))) <> "DA", "Klasa I vracena"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalPartialClass_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' P2 pomesana operacija: jedan OperationID sa dva razlicita Broja -> undo odbijen (corrupt).
Public Sub Test_StornoJournalMixedOp_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_STORNO_ZURNAL, Array(COL_SZ_ID, COL_SZ_OP_ID, COL_SZ_DOCTYPE, COL_SZ_BROJ, COL_SZ_TABELA, COL_SZ_ROWID, COL_SZ_KOLONA, COL_SZ_STARA, COL_SZ_NOVA), _
              Array("ZUR-M1", "SOP-MIX", DOK_TIP_OTKUP, "SVT-MX-A", TBL_OTKUP, "SVT-MX-1", COL_STORNIRANO, "", "Da")
    TcSeedRow TBL_STORNO_ZURNAL, Array(COL_SZ_ID, COL_SZ_OP_ID, COL_SZ_DOCTYPE, COL_SZ_BROJ, COL_SZ_TABELA, COL_SZ_ROWID, COL_SZ_KOLONA, COL_SZ_STARA, COL_SZ_NOVA), _
              Array("ZUR-M2", "SOP-MIX", DOK_TIP_OTKUP, "SVT-MX-B", TBL_OTKUP, "SVT-MX-2", COL_STORNIRANO, "", "Da")
    TcChk UndoOperation_TX("SOP-MIX") = False, "pomesan op (dva broja) -> undo odbijen (corrupt)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalMixedOp_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' P2 prazan BrDok: dva unbound bloka (bez broja) -> DVE odvojene operacije (ne spajaju se).
Public Sub Test_StornoJournalEmptyBrDok_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP: tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC: tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-EB-1", "", "I")
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA), Array("SVT-EB-2", "", "I")
    Dim sel As Collection: Set sel = New Collection: sel.Add "SVT-EB-1": sel.Add "SVT-EB-2"
    TcChk StornoSelectedBlocks_TX(sel) = 2, "storno 2 unbound bloka -> 2"
    ' oba zurnalisana pod ZASEBNIM OperationID (broj je "" ali RowID/PK ih razdvaja)
    TcChk TcDistinctOpsForRow(TBL_OTKUP, "SVT-EB-1") <> TcDistinctOpsForRow(TBL_OTKUP, "SVT-EB-2"), "unbound blokovi -> razliciti OperationID (ne spojeni)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalEmptyBrDok_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Storno-zurnal: LOSSLESS "Vrati storno" za otkup -> storno obrise tblNovac.OtkupID,
' undo ga preko zurnala VRACA (glavni bug review #5). Egzekucija: StornoOtkup_TX ->
' UndoStorno_TX(Otkup, broj). Rollback-safe.
Public Sub Test_StornoJournalUndo_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore                 ' tabela mora postojati za snapshot
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA, COL_OTK_KOLICINA), _
              Array("SVT-SJ-OID", "SVT-SJ-B", "I", 10)
    TcSeedRow TBL_AMBALAZA, Array(COL_AMB_ID, COL_AMB_DOK_ID, COL_AMB_DOK_TIP), _
              Array("SVT-SJ-AID", "SVT-SJ-OID", DOK_TIP_OTKUP)
    TcSeedRow TBL_NOVAC, Array(COL_NOV_ID, COL_NOV_OTKUP_ID), _
              Array("SVT-SJ-NID", "SVT-SJ-OID")

    ' --- STORNO (journaled) ---
    TcChk StornoOtkup_TX("SVT-SJ-OID") = True, "StornoOtkup_TX -> True"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-SJ-OID", COL_STORNIRANO))) = "DA", "otkup stornirano"
    TcChk UCase$(NzS(LookupValue(TBL_AMBALAZA, COL_AMB_ID, "SVT-SJ-AID", COL_STORNIRANO))) = "DA", "ambalaza stornirana"
    TcChk NzS(LookupValue(TBL_NOVAC, COL_NOV_ID, "SVT-SJ-NID", COL_NOV_OTKUP_ID)) = "", "novac OtkupID obrisan (storno)"
    TcChk Len(LatestOpFor(DOK_TIP_OTKUP, "SVT-SJ-B")) > 0, "zurnal operacija zabelezena"

    ' --- UNDO (lossless preko zurnala) ---
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-SJ-B") = True, "UndoStorno_TX (zurnal) -> True"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-SJ-OID", COL_STORNIRANO))) <> "DA", "otkup vracen (aktivan)"
    TcChk UCase$(NzS(LookupValue(TBL_AMBALAZA, COL_AMB_ID, "SVT-SJ-AID", COL_STORNIRANO))) <> "DA", "ambalaza vracena"
    TcChk NzS(LookupValue(TBL_NOVAC, COL_NOV_ID, "SVT-SJ-NID", COL_NOV_OTKUP_ID)) = "SVT-SJ-OID", "novac OtkupID VRACEN (lossless)"

    ' P2 7: ponovni undo iste op -> odbijen (drift guard: Stornirano je sada "" != NovaVrednost "Da")
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-SJ-B") = False, "ponovni undo iste op -> odbijen (drift)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalUndo_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Blocker 3 fix: dvoklasni otkup (StornoOtkupByBrDok_TX) -> JEDAN OperationID ->
' undo vraca OBE klase (ranije je LatestOpFor davao samo poslednju klasu).
Public Sub Test_StornoJournalDualClass_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA, COL_OTK_KOLICINA), Array("SVT-DC-1", "SVT-DC-B", "I", 10)
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KLASA, COL_OTK_KOLICINA), Array("SVT-DC-2", "SVT-DC-B", "II", 5)

    TcChk StornoOtkupByBrDok_TX("SVT-DC-B") = True, "StornoOtkupByBrDok_TX (dvoklasni) -> True"
    TcChk Len(LatestOpFor(DOK_TIP_OTKUP, "SVT-DC-B")) > 0, "dvoklasni -> zabelezen op"
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-DC-B") = True, "undo dvoklasnog -> True"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-DC-1", COL_STORNIRANO))) <> "DA", "Klasa I vracena"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-DC-2", COL_STORNIRANO))) <> "DA", "Klasa II vracena (isti op)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalDualClass_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' Blocker 2 fix: journaled revers undo NE zaobilazi #134 dup-gardu.
Public Sub Test_StornoJournalReversGuard_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_ZURNAL

    ' aktivan revers -> storno (kreira zurnal op)
    TcSeedRow TBL_AMBALAZA, Array(COL_AMB_ID, COL_AMB_DOK_ID, COL_AMB_DOK_TIP), Array("SVT-RG-1", "SVT-RG-R", DOK_TIP_OM_IZLAZ_KOOP)
    TcChk StornoOMKoopByBrDok_TX("SVT-RG-R", DOK_TIP_OM_IZLAZ_KOOP) = True, "revers storno (journaled) -> True"
    ' unesi NOVI aktivan revers istog broj+tip
    TcSeedRow TBL_AMBALAZA, Array(COL_AMB_ID, COL_AMB_DOK_ID, COL_AMB_DOK_TIP), Array("SVT-RG-2", "SVT-RG-R", DOK_TIP_OM_IZLAZ_KOOP)
    ' undo preko ZURNALA mora biti ODBIJEN (dup guard #134, ranije zaobidjen)
    TcChk UndoStorno_TX(DOK_TIP_OM_IZLAZ_KOOP, "SVT-RG-R") = False, "journaled revers undo uz aktivan dup -> ODBIJEN"
    TcChk UCase$(NzS(LookupValue(TBL_AMBALAZA, COL_AMB_ID, "SVT-RG-1", COL_STORNIRANO))) = "DA", "stari revers ostao storniran (nije dupliran)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalReversGuard_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' P2 5: undo je SVE-ILI-NISTA -> zurnal red sa nepostojecim ciljem -> undo False, bez mutacije.
Public Sub Test_StornoJournalUndoValidation_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_STORNO_ZURNAL

    TcSeedRow TBL_STORNO_ZURNAL, Array(COL_SZ_ID, COL_SZ_OP_ID, COL_SZ_DOCTYPE, COL_SZ_BROJ, _
              COL_SZ_TABELA, COL_SZ_ROWID, COL_SZ_KOLONA, COL_SZ_STARA, COL_SZ_NOVA), _
              Array("ZUR-X", "SOP-VALX", DOK_TIP_OTKUP, "SVT-VL-B", TBL_OTKUP, "SVT-VL-NEPOSTOJI", COL_STORNIRANO, "", "Da")
    TcChk UndoOperation_TX("SOP-VALX") = False, "undo sa nepostojecim ciljem -> False (pre-validacija, sve-ili-nista)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalUndoValidation_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' PONISTENJE prijemnice = uzvodna kaskada (649e904): zbirna 1:1 -> storno zbirne +
' njenih otpremnica + prijemnice. DUPLI NAMERNO ostaje list. Egzekucija (ne samo odluka).
Public Sub Test_PonistenjePrijemniceKaskada_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_VEZE

    ' --- Lanac A: PONISTENJE -> uzvodna kaskada ---
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB), _
              Array("SVT-KA-ZID", "SVT-KA-Z", "I", 100, 10)
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE, COL_OTP_KLASA, COL_OTP_KOLICINA, COL_OTP_KOL_AMB), _
              Array("SVT-KA-OID", "SVT-KA-O", "SVT-KA-Z", "I", 100, 10)
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_BROJ_ZBIRNE), _
              Array("SVT-KA-PID", "SVT-KA-P", "I", "SVT-KA-Z")

    Dim rA As Object: Set rA = RunPrijemnicaCorrection("SVT-KA-P", SV_MODE_PONISTENJE, True)
    TcChk CBool(rA("success")), "PONISTENJE prijemnice -> success"
    TcChk TcCountActive(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-KA-Z") = 0, "zbirna stornirana (uzvodna kaskada)"
    TcChk TcCountActive(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, "SVT-KA-Z") = 0, "otpremnica te zbirne stornirana"
    TcChk TcCountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-KA-P") = 0, "prijemnica stornirana"

    ' --- Lanac B: DUPLI -> NAMERNO list (zbirna/otpremnica prezivljavaju) ---
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB), _
              Array("SVT-KB-ZID", "SVT-KB-Z", "I", 100, 10)
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE, COL_OTP_KLASA, COL_OTP_KOLICINA, COL_OTP_KOL_AMB), _
              Array("SVT-KB-OID", "SVT-KB-O", "SVT-KB-Z", "I", 100, 10)
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_BROJ_ZBIRNE), _
              Array("SVT-KB-PID", "SVT-KB-P", "I", "SVT-KB-Z")

    Dim rB As Object: Set rB = RunPrijemnicaCorrection("SVT-KB-P", SV_MODE_DUPLI, True)
    TcChk CBool(rB("success")), "DUPLI prijemnica -> success"
    TcChk TcCountActive(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-KB-Z") = 1, "DUPLI: zbirna ostaje AKTIVNA (list, ne kaskada)"
    TcChk TcCountActive(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, "SVT-KB-Z") = 1, "DUPLI: otpremnica ostaje AKTIVNA"
    TcChk TcCountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-KB-P") = 0, "DUPLI: prijemnica stornirana (list)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_PonistenjePrijemniceKaskada_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' 3.2 odluka: auto-recalc IZDATE zbirne ostaje IN-PLACE (izveden agregat), NE
' re-verzionise se (nov BrojZbirne bi razbio lookup-e; sync je bezbedan ali interni
' join nije). Audit-trag ide u Monitoring. Test: recalc bez otpremnica -> stara
' vrednost -> 0 na ISTOM redu, bez novog zbirna reda.
Public Sub Test_ZbirnaRecalcInPlace_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    ' izdata (default) aktivna zbirna sa zastarelim totalom, bez otpremnica
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB), _
              Array("SVT-ZR-ID", "SVT-ZR-Z1", "I", 999, 9)

    ResetIssuedZbirnaAudit
    TcChk RecalculateZbirnaFromOtpremnice_TX("SVT-ZR-Z1", "SVT-ZR-COR", "test") = True, "recalk izdate zbirne -> True"
    TcChk Val(NzS(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, "SVT-ZR-ID", COL_ZBR_KOLICINA))) = 0, "total spusten na 0 (nema otpremnica)"
    TcChk UCase$(NzS(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, "SVT-ZR-ID", COL_STORNIRANO))) <> "DA", "isti red ostaje AKTIVAN (in-place, ne re-verzija)"
    TcChk TcCountActive(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-ZR-Z1") = 1, "nema novog zbirna reda (bez re-verzionisanja)"
    ' promena izdate zbirne (999->0) -> audit MORA da okine (gate: izdato + promena)
    TcChk InStr(LastIssuedZbirnaAudit(), "SVT-ZR-Z1") > 0, "audit emitovan za izmenu izdate zbirne"

    ' NEGATIVNO: recalk bez promene (total vec 0) -> audit NE sme da okine
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB), _
              Array("SVT-ZR-ID0", "SVT-ZR-Z0", "I", 0, 0)
    ResetIssuedZbirnaAudit
    RecalculateZbirnaFromOtpremnice_TX "SVT-ZR-Z0"
    TcChk Len(LastIssuedZbirnaAudit()) = 0, "nema audita kad se nista ne menja (0->0)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_ZbirnaRecalcInPlace_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' #5 undo: reverse dup-guard -> ne vraca revers ako vec postoji AKTIVAN isti broj+tip.
Public Sub Test_UndoReverseGuard_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA

    ' A: samo storniran revers (nema aktivnog) -> undo prolazi (reaktivira)
    TcSeedRow TBL_AMBALAZA, Array(COL_AMB_ID, COL_AMB_DOK_ID, COL_AMB_DOK_TIP, COL_STORNIRANO), _
              Array("SVT-UR-A1", "SVT-UR-RA", DOK_TIP_OM_IZLAZ_KOOP, "Da")
    TcChk UndoStorno_TX(DOK_TIP_OM_IZLAZ_KOOP, "SVT-UR-RA") = True, "revers undo bez aktivnog -> prolazi"

    ' B: AKTIVAN revers + storniran isti broj+tip -> guard odbija (bez ove garde bi duplirao)
    TcSeedRow TBL_AMBALAZA, Array(COL_AMB_ID, COL_AMB_DOK_ID, COL_AMB_DOK_TIP, COL_STORNIRANO), _
              Array("SVT-UR-B1", "SVT-UR-RB", DOK_TIP_OM_IZLAZ_KOOP, "")
    TcSeedRow TBL_AMBALAZA, Array(COL_AMB_ID, COL_AMB_DOK_ID, COL_AMB_DOK_TIP, COL_STORNIRANO), _
              Array("SVT-UR-B2", "SVT-UR-RB", DOK_TIP_OM_IZLAZ_KOOP, "Da")
    TcChk UndoStorno_TX(DOK_TIP_OM_IZLAZ_KOOP, "SVT-UR-RB") = False, "revers undo uz AKTIVAN duplikat -> odbijeno"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_UndoReverseGuard_Auto GRESKA: " & Err.description: mFail = mFail + 1
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

    ' #7: broj sa DVE generacije -> status se cita sa AKTIVNOG reda, ne sa storniranog
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_TRACE_IZDATO_STATUS, COL_STORNIRANO), _
              Array("SVT-IZ-4S", "SVT-IZ-MIX", "I", IZDATO_DRAFT, "Da")     ' STORNIRAN red = DRAFT
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_TRACE_IZDATO_STATUS), _
              Array("SVT-IZ-4A", "SVT-IZ-MIX", "I", IZDATO_IZDATO)          ' AKTIVAN red = IZDATO

    TcChk DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-IZ-EMPTY") = True, "prazan IzdatoStatus -> izdato"
    TcChk DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-IZ-DRAFT") = False, "DRAFT -> nije izdato"
    TcChk DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-IZ-IZD") = True, "IZDATO -> izdato"
    TcChk DocIsIssued(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-IZ-MIX") = True, "#7: status sa aktivnog (IZDATO), ne storniranog (DRAFT)"

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

' Broj AKTIVNIH (ne-storniranih) redova gde col=val (CountActive u modStornoFlow je Private).
Private Function TcCountActive(ByVal tbl As String, ByVal col As String, ByVal val As String) As Long
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim cKey As Long, cSt As Long
    cKey = GetColumnIndex(tbl, col)
    cSt = GetColumnIndex(tbl, COL_STORNIRANO)
    If cKey = 0 Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cKey))) = Trim$(val) Then
            Dim isStor As Boolean: isStor = False
            If cSt > 0 Then isStor = (UCase$(Trim$(CStr(data(i, cSt)))) = "DA")
            If Not isStor Then n = n + 1
        End If
    Next i
    TcCountActive = n
End Function

Private Function TcRowIndex(ByVal tbl As String, ByVal col As String, ByVal val As String) As Long
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim c As Long: c = GetColumnIndex(tbl, col)
    If c = 0 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, c))) = Trim$(val) Then TcRowIndex = i: Exit Function
    Next i
End Function

' OperationID koji je u zurnalu zabelezio dati (Tabela, RowID) - za proveru grupisanja.
Private Function TcDistinctOpsForRow(ByVal tbl As String, ByVal rowID As String) As String
    Dim data As Variant: data = GetTableData(TBL_STORNO_ZURNAL)
    If IsEmpty(data) Then Exit Function
    Dim cTab As Long, cRow As Long, cOp As Long
    cTab = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_TABELA)
    cRow = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_ROWID)
    cOp = GetColumnIndex(TBL_STORNO_ZURNAL, COL_SZ_OP_ID)
    If cTab = 0 Or cRow = 0 Or cOp = 0 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, cTab))), tbl, vbTextCompare) = 0 _
           And Trim$(CStr(data(i, cRow))) = Trim$(rowID) Then
            TcDistinctOpsForRow = Trim$(CStr(data(i, cOp))): Exit Function
        End If
    Next i
End Function
