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
' Imena palih provera. Idu u poruku gate-a: runner vidi samo tu poruku, ne
' Immediate -- bez imena sabotaza ne moze da se potvrdi PO IMENU.
Private mFailImena As String

' Gate: bez ovoga runner vidi suite kao "blind" -- proslo bez greske, sto NIJE
' isto sto i sve provere prosle. Pojedinacni Test_*_Auto vec hvataju gresku i
' broje mFail; ovde se taj zbir pretvara u verdikt. Konvencija: modTestBanka.
Private Const ERR_STORNOCENTAR_SUITE_FAILED As Long = vbObjectError + 2963

Public Sub Test_StornoCentar_All()
    mPass = 0: mFail = 0: mFailImena = ""
    Test_StampIspravkaTrace_Auto
    Test_BlockStornoDriftReason_Auto
    Test_DocIsIssued_Auto
    Test_OtkupBlockDeadParent_Auto
    Test_BuildStornoImpact_Auto
    Test_GetActiveDocumentsForStorno_Auto
    Test_GetNedovrseno_Auto
    Test_UndoReverseGuard_Auto
    Test_ZbirnaRecalcInPlace_Auto
    Test_PonistenjePrijemniceKaskada_Auto
    Test_StornoJournalUndo_Auto
    Test_StornoJournalReversGuard_Auto
    Test_StornoReversPoStanici_Auto
    Test_StornoReversGranicaRID_Auto
    Test_StornoReversOpisStanice_Auto
    Test_StornoJournalUndoValidation_Auto
    Test_StornoJournalDrift_Auto
    Test_StornoJournalMixedOp_Auto
    Test_StornoJournalReusedBroj_Auto
    Test_StornoJournalDeadParentOtherGen_Auto
    Test_StornoJournalEmptyBrDokUndo_Auto
    Test_ImpactHeaderSum_Auto
    Debug.Print "=== StornoCentar: " & mPass & " OK, " & mFail & " FAIL ==="

    If mFail > 0 Then
        Err.Raise ERR_STORNOCENTAR_SUITE_FAILED, "modTestStornoCentar.Test_StornoCentar_All", _
            "Test_StornoCentar_All: " & CStr(mFail) & " provera palo (PASS=" & _
            CStr(mPass) & "). Pale:" & mFailImena
    End If
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

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), Array("SVT-RB-A", "SVT-RB-OTK")
    TcChk StornoOtkup_TX("SVT-RB-A") = True, "storno gen A -> True"
    Dim opA As String: opA = TcDistinctOpsForRow(TBL_OTKUP, "SVT-RB-A")
    ' druga generacija istog broja (nov aktivan red) -> storno
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), Array("SVT-RB-B", "SVT-RB-OTK")
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
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_OTPREMNICA_ID, COL_STORNIRANO), Array("SVT-DG-A", "SVT-DG-OTK", "SVT-DG-OTP", "Da")
    ' gen B: unbound aktivna -> storno (journaled)
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), Array("SVT-DG-B", "SVT-DG-OTK")
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

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), Array("SVT-EU-1", "")
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

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), Array("SVT-DR-OID", "SVT-DR-B")
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

    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), _
              Array("SVT-SJ-OID", "SVT-SJ-B")
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

' Blocker 2 fix: journaled revers undo NE zaobilazi #134 dup-gardu.
Public Sub Test_StornoJournalReversGuard_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_ZURNAL

    ' Redovi nose ReversID (identitet reversa) i nogu Stanica sa datumom: red bez
    ' ReversID-a storno i undo odbijaju (fail-closed), a duplikat se meri u nizu
    ' (stanica, dan).
    ' aktivan revers -> storno (kreira zurnal op)
    TcSeedRevNoga "SVT-RG-1", Date, "SVT-ST-RG", "Stanica", "SVT-RG-R", DOK_TIP_OM_IZLAZ_KOOP, NoviReversID()
    TcChk StornoOMKoopByBrDok_TX("SVT-RG-R", DOK_TIP_OM_IZLAZ_KOOP) = True, "revers storno (journaled) -> True"
    ' unesi NOVI aktivan revers istog broja, stanice i dana (drugi dokument)
    TcSeedRevNoga "SVT-RG-2", Date, "SVT-ST-RG", "Stanica", "SVT-RG-R", DOK_TIP_OM_IZLAZ_KOOP, NoviReversID()
    ' undo preko ZURNALA mora biti ODBIJEN (dup guard #134, ranije zaobidjen)
    TcChk UndoStorno_TX(DOK_TIP_OM_IZLAZ_KOOP, "SVT-RG-R") = False, "journaled revers undo uz aktivan dup -> ODBIJEN"
    TcChk UCase$(NzS(LookupValue(TBL_AMBALAZA, COL_AMB_ID, "SVT-RG-1", COL_STORNIRANO))) = "DA", "stari revers ostao storniran (nije dupliran)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalReversGuard_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' REVERS PO ReversID-u (REV-IDENT-01) -- ARCHITECTURE_CONTRACT A2 red REV.
' Isti broj legalno nose reversi dve stanice istog dana, pa storno i undo biraju
' po ReversID-u, nikad po (broj, tip), i bez uparivanja noge Kooperant sa nogom
' Stanica preko (broj, dan).
' Fixture: oblik nogu je produkcioni (SaveOMUlaz_TX: FIRMA = noga Stanica, KOOP =
' noga Kooperant + noga Stanica, sve noge jednog dokumenta nose jedan ReversID iz
' NoviReversID); seed u rollback-u. Fault injection su red bez ReversID-a i
' ambalaza uz otkup SA ReversID-om -- oba storno mora da odbije. KOOP isti broj na
' dve stanice istog dana od Faze 2b pravi i pisac (BFP
' Test_BKTX_ReversKoopIstiBrojDveStanice); ovde se meri citalac nad seed-om.
' Nivo merenja: fizicki red (koja noga je stornirana).
'
' SABOTAZE: u ReversRedoviRID ne poredi ReversID -> pukne "revers ISTOG broja na
' drugoj stanici ostaje aktivan"; izbaci proveru uz-otkup u ReversIDRazresi ->
' pukne "ambalaza uz otkup se ne stornira kao revers"; pusti prazan ReversID u
' ReversIDRazresi I u ReversRedoviRID (dva sloja) -> pukne "red bez ReversID-a se
' ne stornira"; UndoGuardReasonZaOp bez ReversID-a -> pukne "undo po operaciji";
' UndoStorno_TX na LatestOpFor umesto LatestOpForRevers -> pukne "undo po broju:
' bira operaciju reversa A".
Public Sub Test_StornoReversPoStanici_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_ZURNAL
    tx.AddTableSnapshot TBL_OTKUP

    Dim d As Date: d = DateSerial(2031, 5, 7)
    Dim ridA As String, ridB As String, ridK1 As String, ridK1d2 As String
    Dim ridK2a As String, ridK2b As String, ridUA As String, ridUB As String, ridO As String
    ridA = NoviReversID(): ridB = NoviReversID()
    ridK1 = NoviReversID(): ridK1d2 = NoviReversID()
    ridK2a = NoviReversID(): ridK2b = NoviReversID()
    ridUA = NoviReversID(): ridUB = NoviReversID()
    ridO = NoviReversID()

    ' FIRMA: isti broj istog dana na dve stanice (samo noge Stanica).
    TcSeedRevNoga "SVT-RS-A", d, "SVT-ST-A", "Stanica", "SVT-RS-F", DOK_TIP_OM_ULAZ_FIRMA, ridA
    TcSeedRevNoga "SVT-RS-B", d, "SVT-ST-B", "Stanica", "SVT-RS-F", DOK_TIP_OM_ULAZ_FIRMA, ridB

    TcChk StornoOMKoopByBrDok_TX("SVT-RS-F", DOK_TIP_OM_ULAZ_FIRMA) = False, _
          "revers FIRMA: bez identiteta dvosmislen broj -> odbijen"
    TcChk TcAmbStorno("SVT-RS-A") = "" And TcAmbStorno("SVT-RS-B") = "", _
          "revers FIRMA: odbijen storno nije dirao nijednu nogu"
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-F", DOK_TIP_OM_ULAZ_FIRMA, "SVT-RS-A") = True, _
          "revers FIRMA: storno po identitetu reda prolazi"
    TcChk TcAmbStorno("SVT-RS-A") = "DA", "revers FIRMA: stornirana je izabrana stanica"
    TcChk TcAmbStorno("SVT-RS-B") = "", "revers FIRMA: revers ISTOG broja na drugoj stanici ostaje aktivan"

    ' Undo garda: ReversID bira dokument, duplikat je pitanje niza (stanica, dan).
    TcChk Len(UndoGuardReason(DOK_TIP_OM_ULAZ_FIRMA, "SVT-RS-F", ridA)) = 0, _
          "undo garda: aktivan revers druge stanice NIJE duplikat"
    TcChk Len(UndoGuardReason(DOK_TIP_OM_ULAZ_FIRMA, "SVT-RS-F", ridB)) > 0, _
          "undo garda: aktivan revers na stanici i danu tog reversa blokira"
    TcChk Len(UndoGuardReason(DOK_TIP_OM_ULAZ_FIRMA, "SVT-RS-F")) > 0, _
          "undo garda: revers bez ReversID-a se odbija (fail-closed)"
    Dim opA As String: opA = LatestOpFor(DOK_TIP_OM_ULAZ_FIRMA, "SVT-RS-F")
    TcChk UndoOperation_TX(opA) = True, _
          "undo po operaciji: ReversID iz AmbID-eva op-a -- revers druge stanice ne blokira"
    TcChk TcAmbStorno("SVT-RS-A") = "", "undo po operaciji: A je ponovo aktivan"

    ' KOOP, jedna stanica: storno sa noge Kooperant stornira obe noge svog
    ' ReversID-a, a revers istog broja iste stanice drugog dana ostaje.
    TcSeedRevNoga "SVT-RK1-K", d, "SVT-KOOP-1", "Kooperant", "SVT-RS-K1", DOK_TIP_OM_ULAZ_KOOP, ridK1
    TcSeedRevNoga "SVT-RK1-S", d, "SVT-ST-A", "Stanica", "SVT-RS-K1", DOK_TIP_OM_ULAZ_KOOP, ridK1
    TcSeedRevNoga "SVT-RK1-K2", DateAdd("d", 1, d), "SVT-KOOP-1", "Kooperant", "SVT-RS-K1", _
                  DOK_TIP_OM_ULAZ_KOOP, ridK1d2
    TcSeedRevNoga "SVT-RK1-S2", DateAdd("d", 1, d), "SVT-ST-A", "Stanica", "SVT-RS-K1", _
                  DOK_TIP_OM_ULAZ_KOOP, ridK1d2
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-K1", DOK_TIP_OM_ULAZ_KOOP, "SVT-RK1-K") = True, _
          "revers KOOP: storno sa noge Kooperant prolazi"
    TcChk TcAmbStorno("SVT-RK1-K") = "DA" And TcAmbStorno("SVT-RK1-S") = "DA", _
          "revers KOOP: stornirane su obe noge izabranog reversa"
    TcChk TcAmbStorno("SVT-RK1-K2") = "" And TcAmbStorno("SVT-RK1-S2") = "", _
          "revers KOOP: isti broj iste stanice drugog dana ostaje aktivan"

    ' KOOP, dve stanice istog dana: noga Kooperant nosi ReversID svog dokumenta, pa
    ' se stornira TACNO on. Do Faze 2a ovo je bilo odbijeno kao nerazlucivo.
    TcSeedRevNoga "SVT-RK2-K1", d, "SVT-KOOP-1", "Kooperant", "SVT-RS-K2", DOK_TIP_OM_IZLAZ_KOOP, ridK2a
    TcSeedRevNoga "SVT-RK2-S1", d, "SVT-ST-A", "Stanica", "SVT-RS-K2", DOK_TIP_OM_IZLAZ_KOOP, ridK2a
    TcSeedRevNoga "SVT-RK2-K2", d, "SVT-KOOP-2", "Kooperant", "SVT-RS-K2", DOK_TIP_OM_IZLAZ_KOOP, ridK2b
    TcSeedRevNoga "SVT-RK2-S2", d, "SVT-ST-B", "Stanica", "SVT-RS-K2", DOK_TIP_OM_IZLAZ_KOOP, ridK2b
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-K2", DOK_TIP_OM_IZLAZ_KOOP) = False, _
          "revers KOOP dve stanice: bez identiteta dvosmislen broj -> odbijen"
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-K2", DOK_TIP_OM_IZLAZ_KOOP, "SVT-RK2-K1") = True, _
          "revers KOOP dve stanice: storno sa noge Kooperant prolazi po ReversID-u"
    TcChk TcAmbStorno("SVT-RK2-K1") = "DA" And TcAmbStorno("SVT-RK2-S1") = "DA", _
          "revers KOOP dve stanice: stornirane su obe noge izabranog reversa"
    TcChk TcAmbStorno("SVT-RK2-K2") = "" And TcAmbStorno("SVT-RK2-S2") = "", _
          "revers KOOP dve stanice: revers istog broja i dana na drugoj stanici ostaje aktivan"

    ' Red bez ReversID-a (fault injection): identitet se ne pogadja ni po broju, ni
    ' po stanici i danu.
    TcSeedRevNoga "SVT-RN-S", d, "SVT-ST-C", "Stanica", "SVT-RS-N", DOK_TIP_OM_ULAZ_FIRMA, ""
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-N", DOK_TIP_OM_ULAZ_FIRMA, "SVT-RN-S") = False, _
          "red bez ReversID-a se ne stornira (sa identitetom reda)"
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-N", DOK_TIP_OM_ULAZ_FIRMA) = False, _
          "red bez ReversID-a se ne stornira (po broju)"
    TcChk TcAmbStorno("SVT-RN-S") = "", "red bez ReversID-a se ne stornira -- nije dirnut"

    ' Ambalaza uz otkup (DokumentID = OtkupID) nije revers -- ni kad bi nosila
    ' ReversID (fault injection: pisac otkupa ga ne pise).
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), Array("SVT-RS-OTK", "SVT-RS-OTKBR")
    TcSeedRevNoga "SVT-RO-K", d, "SVT-KOOP-1", "Kooperant", "SVT-RS-OTK", DOK_TIP_OM_IZLAZ_KOOP, ridO
    TcSeedRevNoga "SVT-RO-S", d, "SVT-ST-A", "Stanica", "SVT-RS-OTK", DOK_TIP_OM_IZLAZ_KOOP, ridO
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-OTK", DOK_TIP_OM_IZLAZ_KOOP, "SVT-RO-K") = False, _
          "ambalaza uz otkup se ne stornira kao revers"
    TcChk TcAmbStorno("SVT-RO-K") = "" And TcAmbStorno("SVT-RO-S") = "", _
          "ambalaza uz otkup: nijedna noga nije dirnuta"

    ' Undo po broju (UndoStorno_TX) bira operaciju ISTOG ReversID-a: noviji storno
    ' reversa iste oznake na drugoj stanici, vec vracen po operaciji, nije ta.
    TcSeedRevNoga "SVT-RU-A", d, "SVT-ST-A", "Stanica", "SVT-RS-U", DOK_TIP_OM_ULAZ_FIRMA, ridUA
    TcSeedRevNoga "SVT-RU-B", DateAdd("d", 1, d), "SVT-ST-B", "Stanica", "SVT-RS-U", _
                  DOK_TIP_OM_ULAZ_FIRMA, ridUB
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-U", DOK_TIP_OM_ULAZ_FIRMA, "SVT-RU-A") = True, _
          "undo po broju: preduslov -- storno A"
    TcChk StornoOMKoopByBrDok_TX("SVT-RS-U", DOK_TIP_OM_ULAZ_FIRMA, "SVT-RU-B") = True, _
          "undo po broju: preduslov -- storno B (noviji)"
    TcChk UndoOperation_TX(LatestOpFor(DOK_TIP_OM_ULAZ_FIRMA, "SVT-RS-U")) = True, _
          "undo po broju: preduslov -- B vracen po operaciji"
    TcChk UndoStorno_TX(DOK_TIP_OM_ULAZ_FIRMA, "SVT-RS-U") = True, _
          "undo po broju: bira operaciju reversa A, ne noviju vracenu operaciju B"
    TcChk TcAmbStorno("SVT-RU-A") = "" And TcAmbStorno("SVT-RU-B") = "", _
          "undo po broju: A je vracen, B ostaje aktivan"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoReversPoStanici_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' REV-IDENT-01 GRANICA DOKUMENTA: ReversID bira redove za storno, undo i stampu, pa
' svaki red koji ga nosi mora pripadati ISTOM reversu. Red tudjeg dokumenta sa
' istim ReversID-om obara celu operaciju -- nijedan red se ne menja (fail-closed).
' Fixture: fault injection (pisac ReversID pise samo na noge jednog reversa), seed u
' rollback-u; svaki slucaj je ispravan revers + JEDAN los red. Negativne kontrole:
' ispravan KOOP revers i FIRMA revers sa dve noge istog vozaca su cisti.
' Nivo merenja: fizicki red (da li je ijedan red storniran) + razlog granice.
'
' SABOTAZE: u ReversIDGranica preskoci proveru tipa -> pukne "granica: ReversID samo
' na redu koji nije revers" (uz nogu reversa tip hvata i provera smera -- zato taj
' slucaj nema nogu reversa); proveru broja -> "granica: red drugog broja"; dana ->
' "granica: red drugog dana"; vozaca -> "granica: noga drugog vozaca"; stanice ->
' "granica: noga Stanica druge stanice"; kooperanta -> "granica: noga Kooperant
' drugog kooperanta"; tipa entiteta -> "granica: red koji nije noga (Kupac)";
' uz-otkup -> "granica: ReversID na ambalazi uz otkup"; izbaci poziv granice iz
' ReversRedoviRID -> "storno: red koji nije revers sa istim ReversID-om -> storno
' odbijen"; izbaci granicu iz UndoGuardReason -> "undo po operaciji: tudj red sa
' istim ReversID-om -> undo odbijen" (zurnal-put nema drugi sloj; legacy undo po
' broju hvata i ReversRedoviRID pri vracanju, pa "undo: tudj red ..." drze dva
' sloja); izbaci B10 nalaz -> "B10: ReversID na redu koji nije revers prijavljen".
Public Sub Test_StornoReversGranicaRID_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_ZURNAL
    tx.AddTableSnapshot TBL_OTKUP

    Dim d As Date: d = DateSerial(2031, 6, 11)
    Dim rid As String, colsV As Variant
    colsV = Array(COL_AMB_ID, COL_AMB_DATUM, COL_AMB_ENTITET, COL_AMB_ENTITET_TIP, _
                  COL_AMB_DOK_ID, COL_AMB_DOK_TIP, COL_AMB_VOZAC, COL_AMB_REVERS_ID)

    ' Negativne kontrole: granica ne sme da odbije legitiman dokument.
    rid = TcSeedKoopRevers("SVT-GR-OK", d)
    TcChk Len(ReversIDGranica(rid)) = 0, "granica: ispravan KOOP revers je cist"
    rid = NoviReversID()
    TcSeedRow TBL_AMBALAZA, colsV, Array("SVT-GR-V-S1", d, "SVT-ST-GR1", "Stanica", "SVT-GR-V", _
                                         DOK_TIP_OM_ULAZ_FIRMA, "SVT-VOZ-GR1", rid)
    TcSeedRow TBL_AMBALAZA, colsV, Array("SVT-GR-V-S2", d, "SVT-ST-GR1", "Stanica", "SVT-GR-V", _
                                         DOK_TIP_OM_ULAZ_FIRMA, "SVT-VOZ-GR1", rid)
    TcChk Len(ReversIDGranica(rid)) = 0, "granica: FIRMA revers sa dve noge istog vozaca je cist"
    TcSeedRow TBL_AMBALAZA, colsV, Array("SVT-GR-V-S3", d, "SVT-ST-GR1", "Stanica", "SVT-GR-V", _
                                         DOK_TIP_OM_ULAZ_FIRMA, "SVT-VOZ-GR2", rid)
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: noga drugog vozaca"

    ' ReversID samo na redu koji nije revers (bez ijedne noge reversa): meri pravilo
    ' tipa samo za sebe -- uz nogu reversa isti red odbija i provera smera.
    rid = NoviReversID()
    TcSeedRevNoga "SVT-GR-T-X", d, "SVT-ST-GR1", "Stanica", "SVT-GR-T", DOK_TIP_OTPREMNICA, rid
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: ReversID samo na redu koji nije revers"

    ' Red koji nije revers (otpremnica) sa istim ReversID-om -- slucaj iz review-a.
    rid = TcSeedKoopRevers("SVT-GR-A", d)
    TcSeedRevNoga "SVT-GR-A-X", d, "SVT-ST-GR1", "Stanica", "SVT-GR-OTP", DOK_TIP_OTPREMNICA, rid
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: red koji nije revers (otpremnica) sa istim ReversID-om"
    TcChk StornoOMKoopByBrDok_TX("SVT-GR-A", DOK_TIP_OM_IZLAZ_KOOP, "SVT-GR-A-K") = False, _
          "storno: red koji nije revers sa istim ReversID-om -> storno odbijen"
    TcChk StornoOMKoopByBrDok_TX("SVT-GR-A", DOK_TIP_OM_IZLAZ_KOOP) = False, _
          "storno po broju: red koji nije revers sa istim ReversID-om -> storno odbijen"
    TcChk TcAmbStorno("SVT-GR-A-K") = "" And TcAmbStorno("SVT-GR-A-S") = "" And _
          TcAmbStorno("SVT-GR-A-X") = "", _
          "storno odbijen na granici: nijedan red nije promenjen (ni red otpremnice)"
    TcChk InStr(1, TcIntegritetSa("SVT-GR-A-X"), "ReversID na redu koji nije revers", vbBinaryCompare) > 0, _
          "B10: ReversID na redu koji nije revers prijavljen"

    ' Red drugog broja (isti smer) pod istim ReversID-om.
    rid = TcSeedKoopRevers("SVT-GR-B", d)
    TcSeedRevNoga "SVT-GR-B-X", d, "SVT-ST-GR1", "Stanica", "SVT-GR-B2", DOK_TIP_OM_IZLAZ_KOOP, rid
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: red drugog broja"
    TcChk StornoOMKoopByBrDok_TX("SVT-GR-B", DOK_TIP_OM_IZLAZ_KOOP, "SVT-GR-B-K") = False, _
          "storno: red drugog broja pod istim ReversID-om -> storno odbijen"
    TcChk TcAmbStorno("SVT-GR-B-K") = "" And TcAmbStorno("SVT-GR-B-X") = "", _
          "storno odbijen na granici: red drugog broja nije promenjen"

    ' Red drugog dana pod istim ReversID-om.
    rid = TcSeedKoopRevers("SVT-GR-C", d)
    TcSeedRevNoga "SVT-GR-C-X", DateAdd("d", 1, d), "SVT-ST-GR1", "Stanica", "SVT-GR-C", DOK_TIP_OM_IZLAZ_KOOP, rid
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: red drugog dana"

    ' Noga Stanica druge stanice.
    rid = TcSeedKoopRevers("SVT-GR-D", d)
    TcSeedRevNoga "SVT-GR-D-X", d, "SVT-ST-GR2", "Stanica", "SVT-GR-D", DOK_TIP_OM_IZLAZ_KOOP, rid
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: noga Stanica druge stanice"

    ' Noga Kooperant drugog kooperanta.
    rid = TcSeedKoopRevers("SVT-GR-E", d)
    TcSeedRevNoga "SVT-GR-E-X", d, "SVT-KOOP-GR2", "Kooperant", "SVT-GR-E", DOK_TIP_OM_IZLAZ_KOOP, rid
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: noga Kooperant drugog kooperanta"

    ' Red koji nije noga (Kupac) pod istim ReversID-om.
    rid = TcSeedKoopRevers("SVT-GR-F", d)
    TcSeedRevNoga "SVT-GR-F-X", d, "SVT-KUP-GR", "Kupac", "SVT-GR-F", DOK_TIP_OM_IZLAZ_KOOP, rid
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: red koji nije noga (Kupac)"

    ' Ambalaza uz otkup sa ReversID-om.
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), Array("SVT-GR-OTK", "SVT-GR-OTKBR")
    rid = TcSeedKoopRevers("SVT-GR-OTK", d)
    TcChk Len(ReversIDGranica(rid)) > 0, "granica: ReversID na ambalazi uz otkup"

    ' Undo: storniran revers + aktivan tudj red (otpremnica) sa istim ReversID-om.
    rid = NoviReversID()
    TcSeedRevNoga "SVT-GR-U-S", d, "SVT-ST-GR3", "Stanica", "SVT-GR-U", DOK_TIP_OM_ULAZ_FIRMA, rid, "Da"
    TcSeedRevNoga "SVT-GR-U-X", d, "SVT-ST-GR3", "Stanica", "SVT-GR-OTPU", DOK_TIP_OTPREMNICA, rid
    TcChk UndoStorno_TX(DOK_TIP_OM_ULAZ_FIRMA, "SVT-GR-U") = False, _
          "undo: tudj red sa istim ReversID-om -> undo odbijen"
    TcChk TcAmbStorno("SVT-GR-U-S") = "DA" And TcAmbStorno("SVT-GR-U-X") = "", _
          "undo odbijen na granici: nijedan red nije promenjen"

    ' Undo po OPERACIJI (zurnal-put, i ekran Oporavak): UndoOperation_TX vraca celije
    ' po AmbID-u i ne zove ReversRedoviRID, pa je granica u UndoGuardReason tu JEDINA
    ' kapija. Revers se stornira ispravan, a tudj red sa njegovim ReversID-om nastane
    ' posle storna.
    rid = NoviReversID()
    TcSeedRevNoga "SVT-GR-J-S", d, "SVT-ST-GR4", "Stanica", "SVT-GR-J", DOK_TIP_OM_ULAZ_FIRMA, rid
    TcChk StornoOMKoopByBrDok_TX("SVT-GR-J", DOK_TIP_OM_ULAZ_FIRMA, "SVT-GR-J-S") = True, _
          "undo po operaciji: preduslov -- ispravan revers storniran kroz zurnal"
    TcSeedRevNoga "SVT-GR-J-X", d, "SVT-ST-GR4", "Stanica", "SVT-GR-OTPJ", DOK_TIP_OTPREMNICA, rid
    Dim opJ As String: opJ = LatestOpFor(DOK_TIP_OM_ULAZ_FIRMA, "SVT-GR-J")
    TcChk Len(opJ) > 0, "undo po operaciji: preduslov -- operacija storna postoji"
    TcChk Len(UndoGuardReason(DOK_TIP_OM_ULAZ_FIRMA, "SVT-GR-J", rid)) > 0, _
          "undo garda: tudj red sa istim ReversID-om blokira"
    TcChk Len(UndoGuardReasonZaOp(opJ, DOK_TIP_OM_ULAZ_FIRMA, "SVT-GR-J")) > 0, _
          "undo garda po operaciji: tudj red sa istim ReversID-om blokira"
    TcChk UndoOperation_TX(opJ) = False, _
          "undo po operaciji: tudj red sa istim ReversID-om -> undo odbijen"
    TcChk TcAmbStorno("SVT-GR-J-S") = "DA" And TcAmbStorno("SVT-GR-J-X") = "", _
          "undo po operaciji odbijen na granici: nijedan red nije promenjen"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoReversGranicaRID_Auto GRESKA: " & Err.description: mFail = mFail + 1
End Sub

' REV-IDENT-01 Faza 2b: isti KOOP broj, smer i dan legalno nose reversi DVE stanice,
' i to istog kooperanta. Red liste Storno i red liste "Vrati storno" ih po (tip,
' broj) ne razlikuju, pa opis koji operater potvrdjuje mora da imenuje stanicu i
' dan. MsgBox se u headless run-u ne meri -- meri se tekst koji mu se predaje
' (modStornoDok.DokumentOpis, modStornoZurnal.UndoOpisOperacije).
' Fixture: produkcioni oblik nogu (Kooperant + Stanica, jedan ReversID), seed u
' rollback-u. Nivo merenja: logicki dokument -- tekst potvrde po dokumentu.
'
' SABOTAZE: DokumentOpis bez ReversOpis-a -> pukne "storno opis: dva reversa istog
' broja, smera i dana razlikuju se po stanici"; UndoOpisOperacije bez ReversOpis-a
' -> pukne "undo opis: dve operacije istog broja razlikuju se po stanici";
' UndoPotvrdaTekst po (tip, broj) -> pukne "undo potvrda: tekst MsgBox-a imenuje
' stanicu"; izbaci ReversStanicaDan iz StornoRazlog -> pukne "storno kapija: revers
' bez noge Stanica se odbija pre potvrde"; IspravkaReversOpis bez opisa -> pukne
' "ispravka opis: dve ispravke istog broja razlikuju se po stanici"; izbaci
' IspravkaReversOpis iz GetNedovrseno -> pukne "Nedovrseno: redovi dve ispravke
' istog broja razlikuju se po stanici".
Public Sub Test_StornoReversOpisStanice_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    EnsureStornoZurnalSchemaCore
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_ZURNAL
    tx.AddTableSnapshot TBL_STORNO_VEZE

    Dim d As Date: d = DateSerial(2031, 7, 9)
    Dim ridA As String: ridA = NoviReversID()
    Dim ridB As String: ridB = NoviReversID()
    TcSeedRevNoga "SVT-OP-KA", d, "SVT-KOOP-OP", "Kooperant", "SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, ridA
    TcSeedRevNoga "SVT-OP-SA", d, "SVT-ST-OPA", "Stanica", "SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, ridA
    TcSeedRevNoga "SVT-OP-KB", d, "SVT-KOOP-OP", "Kooperant", "SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, ridB
    TcSeedRevNoga "SVT-OP-SB", d, "SVT-ST-OPB", "Stanica", "SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, ridB

    Dim opisA As String, opisB As String
    opisA = modStornoDok.DokumentOpis(STIP_REVERSI, "SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-KA")
    opisB = modStornoDok.DokumentOpis(STIP_REVERSI, "SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-KB")
    TcChk InStr(1, opisA, "SVT-ST-OPA", vbBinaryCompare) > 0 And _
          InStr(1, opisB, "SVT-ST-OPB", vbBinaryCompare) > 0 And opisA <> opisB, _
          "storno opis: dva reversa istog broja, smera i dana razlikuju se po stanici"
    TcChk InStr(1, opisA, Format$(d, "dd.mm.yyyy"), vbBinaryCompare) > 0, "storno opis: nosi dan reversa"
    TcChk InStr(1, opisA, "SVT-OP-7", vbBinaryCompare) > 0, "storno opis: i dalje nosi broj"

    Dim opA As String, opB As String, uA As String, uB As String
    TcChk StornoOMKoopByBrDok_TX("SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-KA") = True, _
          "undo opis: preduslov -- storno reversa A"
    opA = LatestOpFor(DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-7")
    TcChk StornoOMKoopByBrDok_TX("SVT-OP-7", DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-KB") = True, _
          "undo opis: preduslov -- storno reversa B"
    opB = LatestOpFor(DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-7")
    TcChk Len(opA) > 0 And Len(opB) > 0 And opA <> opB, "undo opis: preduslov -- dve operacije istog broja"
    uA = UndoOpisOperacije(opA, DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-7")
    uB = UndoOpisOperacije(opB, DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-7")
    TcChk InStr(1, uA, "SVT-ST-OPA", vbBinaryCompare) > 0 And _
          InStr(1, uB, "SVT-ST-OPB", vbBinaryCompare) > 0 And uA <> uB, _
          "undo opis: dve operacije istog broja razlikuju se po stanici"
    TcChk InStr(1, modScrOporavak.UndoPotvrdaTekst(opA, DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-7"), "SVT-ST-OPA", _
                vbBinaryCompare) > 0, "undo potvrda: tekst MsgBox-a imenuje stanicu"

    ' Dve otvorene ispravke reversa istog broja, smera i dana na dve stanice: u listi
    ' Nedovrseno i u potvrdi "Odbaci" razlikuju ih stanica i dan (OldDocID = ReversID).
    Dim ridD As String: ridD = NoviReversID()
    Dim ridE As String: ridE = NoviReversID()
    TcSeedRevNoga "SVT-OP-KD", d, "SVT-KOOP-OP", "Kooperant", "SVT-OP-9", DOK_TIP_OM_IZLAZ_KOOP, ridD
    TcSeedRevNoga "SVT-OP-SD", d, "SVT-ST-OPD", "Stanica", "SVT-OP-9", DOK_TIP_OM_IZLAZ_KOOP, ridD
    TcSeedRevNoga "SVT-OP-KE", d, "SVT-KOOP-OP", "Kooperant", "SVT-OP-9", DOK_TIP_OM_IZLAZ_KOOP, ridE
    TcSeedRevNoga "SVT-OP-SE", d, "SVT-ST-OPE", "Stanica", "SVT-OP-9", DOK_TIP_OM_IZLAZ_KOOP, ridE
    Dim resD As Object, resE As Object, cidD As String, cidE As String
    Set resD = modStornoFlow.RunReversCorrection("SVT-OP-9", DOK_TIP_OM_IZLAZ_KOOP, SV_MODE_RESI_KASNIJE, "SVT-OP-KD")
    Set resE = modStornoFlow.RunReversCorrection("SVT-OP-9", DOK_TIP_OM_IZLAZ_KOOP, SV_MODE_RESI_KASNIJE, "SVT-OP-KE")
    cidD = CStr(resD("correctionID"))
    cidE = CStr(resE("correctionID"))
    TcChk Len(cidD) > 0 And Len(cidE) > 0 And cidD <> cidE, _
          "ispravka opis: preduslov -- dve otvorene ispravke istog broja"
    TcChk InStr(1, modStornoDok.IspravkaReversOpis(cidD), "SVT-ST-OPD", vbBinaryCompare) > 0 And _
          InStr(1, modStornoDok.IspravkaReversOpis(cidE), "SVT-ST-OPE", vbBinaryCompare) > 0, _
          "ispravka opis: dve ispravke istog broja razlikuju se po stanici"
    ' Isto u redovima liste Nedovrseno (GetNedovrseno) -- iz njih operater bira.
    Dim ned As Collection, nr As Long, opisD As String, opisE As String
    Set ned = modStornoRecovery.GetNedovrseno()
    For nr = 1 To ned.count
        If CStr(ned(nr)("correctionID")) = cidD Then opisD = CStr(ned(nr)("opis"))
        If CStr(ned(nr)("correctionID")) = cidE Then opisE = CStr(ned(nr)("opis"))
    Next nr
    TcChk InStr(1, opisD, "SVT-ST-OPD", vbBinaryCompare) > 0 And _
          InStr(1, opisE, "SVT-ST-OPE", vbBinaryCompare) > 0, _
          "Nedovrseno: redovi dve ispravke istog broja razlikuju se po stanici"

    ' Revers bez noge Stanica (B10 nalaz): stanica i dan nisu poznati, pa ga kapija
    ' storna odbija PRE potvrde -- potvrda bez stanice bila bi upravo dvosmislena.
    Dim ridC As String: ridC = NoviReversID()
    TcSeedRevNoga "SVT-OP-KC", d, "SVT-KOOP-OP", "Kooperant", "SVT-OP-8", DOK_TIP_OM_IZLAZ_KOOP, ridC
    TcChk Len(modStornoDok.StornoRazlog(STIP_REVERSI, "SVT-OP-8", DOK_TIP_OM_IZLAZ_KOOP, "SVT-OP-KC")) > 0, _
          "storno kapija: revers bez noge Stanica se odbija pre potvrde"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoReversOpisStanice_Auto GRESKA: " & Err.description: mFail = mFail + 1
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

    ' Redovi nose ReversID (identitet reversa) i nogu Stanica sa datumom (duplikat
    ' se meri u nizu (stanica, dan)); red bez ReversID-a undo odbija (fail-closed).
    ' A: samo storniran revers (nema aktivnog) -> undo prolazi (reaktivira)
    TcSeedRevNoga "SVT-UR-A1", Date, "SVT-ST-UR", "Stanica", "SVT-UR-RA", DOK_TIP_OM_IZLAZ_KOOP, NoviReversID(), "Da"
    TcChk UndoStorno_TX(DOK_TIP_OM_IZLAZ_KOOP, "SVT-UR-RA") = True, "revers undo bez aktivnog -> prolazi"

    ' B: AKTIVAN revers + storniran istog broja, stanice i dana -> guard odbija (bez ove garde bi duplirao)
    TcSeedRevNoga "SVT-UR-B1", Date, "SVT-ST-UR", "Stanica", "SVT-UR-RB", DOK_TIP_OM_IZLAZ_KOOP, NoviReversID(), ""
    TcSeedRevNoga "SVT-UR-B2", Date, "SVT-ST-UR", "Stanica", "SVT-UR-RB", DOK_TIP_OM_IZLAZ_KOOP, NoviReversID(), "Da"
    TcChk UndoStorno_TX(DOK_TIP_OM_IZLAZ_KOOP, "SVT-UR-RB") = False, "revers undo uz AKTIVAN duplikat -> odbijeno"

    ' C: cetiri smera dele jedan niz (stanica, dan) -- aktivan revers istog broja u
    ' DRUGOM smeru je takodje duplikat (anomalija: pisac je ne pravi, a recovery je
    ' fail-closed). SABOTAZA: u UndoGuardReason pitaj samo docType -> pukne "undo
    ' garda: aktivan revers DRUGOG smera istog broja, stanice i dana blokira".
    Dim ridUC As String: ridUC = NoviReversID()
    TcSeedRevNoga "SVT-UR-C1", Date, "SVT-ST-UR", "Stanica", "SVT-UR-RC", DOK_TIP_OM_ULAZ_KOOP, ridUC, "Da"
    TcSeedRevNoga "SVT-UR-C2", Date, "SVT-ST-UR", "Stanica", "SVT-UR-RC", DOK_TIP_OM_IZLAZ_FIRMA, NoviReversID(), ""
    TcChk Len(UndoGuardReason(DOK_TIP_OM_ULAZ_KOOP, "SVT-UR-RC", ridUC)) > 0, _
          "undo garda: aktivan revers DRUGOG smera istog broja, stanice i dana blokira"
    TcChk UndoStorno_TX(DOK_TIP_OM_ULAZ_KOOP, "SVT-UR-RC") = False, _
          "revers undo uz aktivan revers drugog smera -> odbijeno"
    TcChk TcAmbStorno("SVT-UR-C1") = "DA", "revers undo drugog smera: storniran revers ostao storniran"

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
    tx.AddTableSnapshot TBL_OTPREMNICA_IZVORI
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_KLASA), _
              Array("SVT-DR-OTP", "SVT-DR-O1", "I")                ' aktivna otpremnica
    TcSeedRow TBL_OTKUP, Array(COL_OTK_ID, COL_OTK_BR_DOK), _
              Array("SVT-DR-BLK", "SVT-DR-BD")                     ' blok
    ' Pripadnost ide KANONOM (tblOtpremnicaIzvori), ne kolonom Otkup.OtpremnicaID:
    ' nju od S3a ne pise nijedan zivi put, pa je kapija citajuci nju uvek
    ' vracala "bezbedno je" (S3c). Seed mora da govori isti jezik kao kapija.
    TcSeedRow TBL_OTPREMNICA_IZVORI, Array(COL_OPI_ID, COL_OPI_OTPREMNICA_ID, COL_OPI_OTKUP_ID), _
              Array("SVT-DR-IZV", "SVT-DR-OTP", "SVT-DR-BLK")      ' blok vezan za nju

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
    TcPecatiGeneraciju tbl, nr
End Sub

' ZBR-IDENT-01: aktivan red tblZbirna MORA da nosi GeneracijaID -- v. isti
' komentar uz modTestStorno.PecatiGeneracijuAkoZbirna. Pecat ide kroz produkcionu
' rutinu, pa dve klase istog broja i vlasnika dele generaciju.
Private Sub TcPecatiGeneraciju(ByVal tbl As String, ByVal nr As ListRow)
    If StrComp(tbl, TBL_ZBIRNA, vbTextCompare) <> 0 Then Exit Sub
    ' RequireColumnIndex, ne GetColumnIndex: bez kolone vlasnika pao bi tek
    ' Cells(1, 0), a to je greska koja ne kaze sta nedostaje.
    Dim cBr As Long, cVo As Long, cKu As Long
    cBr = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modTestStornoCentar.TcPecatiGeneraciju")
    cVo = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, "modTestStornoCentar.TcPecatiGeneraciju")
    cKu = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, "modTestStornoCentar.TcPecatiGeneraciju")
    ApplyGeneracijaID TBL_ZBIRNA, nr.Index, _
                      COL_ZBR_BROJ, NzToText(nr.Range.cells(1, cBr).value), _
                      COL_ZBR_VOZAC, NzToText(nr.Range.cells(1, cVo).value), _
                      COL_ZBR_KUPAC, NzToText(nr.Range.cells(1, cKu).value)
    ' Seed koji tiho prekrsi invarijantu obara sve nizvodno, a suite ostane
    ' zelena iz pogresnog razloga. Zato glasno, ovde, a ne po tvrdnjama.
    Dim cGen As Long
    cGen = RequireColumnIndex(TBL_ZBIRNA, COL_GENERACIJA_ID, "modTestStornoCentar.TcPecatiGeneraciju")
    If Len(NzToText(nr.Range.cells(1, cGen).value)) = 0 Then
        Err.Raise vbObjectError + 2802, "modTestStornoCentar.TcPecatiGeneraciju", _
                  "Seed tblZbirna bez GeneracijaID (ZBR-IDENT-01)."
    End If
End Sub

Private Sub TcChk(ByVal cond As Boolean, ByVal nm As String)
    If cond Then
        mPass = mPass + 1
        Debug.Print "OK   " & nm
    Else
        mFail = mFail + 1
        mFailImena = mFailImena & " | " & nm
        Debug.Print "FAIL " & nm
    End If
End Sub

Private Function NzS(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzS = "" Else NzS = Trim$(CStr(v))
End Function

' Noga reversa u produkcionom obliku (SaveOMUlaz_TX): AmbID, datum, entitet, broj,
' tip, ReversID (REV-IDENT-01; prazan = red bez identiteta, fault injection).
Private Sub TcSeedRevNoga(ByVal ambID As String, ByVal d As Date, ByVal entID As String, _
                          ByVal entTip As String, ByVal broj As String, ByVal dokTip As String, _
                          ByVal rid As String, Optional ByVal stornirano As String = "")
    TcSeedRow TBL_AMBALAZA, _
        Array(COL_AMB_ID, COL_AMB_DATUM, COL_AMB_ENTITET, COL_AMB_ENTITET_TIP, _
              COL_AMB_DOK_ID, COL_AMB_DOK_TIP, COL_STORNIRANO, COL_AMB_REVERS_ID), _
        Array(ambID, d, entID, entTip, broj, dokTip, stornirano, rid)
End Sub

' Ispravan KOOP revers (izdavanje): noga Kooperant + noga Stanica pod jednim
' ReversID-om iz produkcione fabrike. Vraca ReversID.
Private Function TcSeedKoopRevers(ByVal broj As String, ByVal d As Date) As String
    Dim rid As String: rid = NoviReversID()
    TcSeedRevNoga broj & "-K", d, "SVT-KOOP-GR1", "Kooperant", broj, DOK_TIP_OM_IZLAZ_KOOP, rid
    TcSeedRevNoga broj & "-S", d, "SVT-ST-GR1", "Stanica", broj, DOK_TIP_OM_IZLAZ_KOOP, rid
    TcSeedKoopRevers = rid
End Function

' Redovi nalaza integriteta (modIntegritet) koji sadrze dati tekst, spojeni vbLf.
Private Function TcIntegritetSa(ByVal sadrzi As String) As String
    Dim nal As Variant, i As Long, spoj As String
    nal = modIntegritet.GetIntegritetRows()
    If Not IsArray(nal) Then Exit Function
    For i = LBound(nal, 1) To UBound(nal, 1)
        If InStr(1, CStr(nal(i, 2)), sadrzi, vbBinaryCompare) > 0 Then spoj = spoj & CStr(nal(i, 2)) & vbLf
    Next i
    TcIntegritetSa = spoj
End Function

' Oznaka storna reda ambalaze po AmbID-u, velikim slovima ("" = aktivan).
Private Function TcAmbStorno(ByVal ambID As String) As String
    TcAmbStorno = UCase$(NzS(LookupValue(TBL_AMBALAZA, COL_AMB_ID, ambID, COL_STORNIRANO)))
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
