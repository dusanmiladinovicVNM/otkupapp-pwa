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
    Test_BuildStornoImpact_Auto
    Test_GetActiveDocumentsForStorno_Auto
    Test_GetNedovrseno_Auto
    Test_PonistenjePrijemniceKaskada_Auto
    Test_StornoJournalUndo_Auto
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

' #6: impact header kolicina = SUMA STAVKI dokumenta (I + II). Ranije je citao prvi
' red klase, pa je potcenio dvoklasni dokument u uvidu pre potvrde. Od S6 koraka 4
' dokument je jedno zaglavlje sa stavkama, a uvid ga bira po PrijemnicaID-u:
' storniran dokument istog broja ne ulazi, jer ga identitet ne bira.
Public Sub Test_ImpactHeaderSum_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA_STAVKE
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ), Array("SVT-IH-1", "SVT-IH-P")
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_STORNIRANO), Array("SVT-IH-3", "SVT-IH-P", "Da")
    TcSeedStavkaPrijemnice "SVT-IH-1", 1, "I", 100
    TcSeedStavkaPrijemnice "SVT-IH-1", 2, "II", 50
    TcSeedStavkaPrijemnice "SVT-IH-3", 1, "I", 999

    Dim m As Object: Set m = BuildStornoImpact(FLOW_DOC_PRIJEMNICA, "SVT-IH-P", "", "SVT-IH-1")
    Dim h As Object: Set h = m("header")
    TcChk Val(NzS(h("kolicina"))) = 150, "impact header kolicina = suma stavki dokumenta (100+50), storniran istog broja izuzet"

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
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_STORNIRANO), Array("SVT-DG-OTP", "SVT-DG-OB", "Da")
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
'
' Otkup je BEZ ambalaze, i to je pravilo, ne pojednostavljenje: undo otkupa bez
' ambalaznog efekta prolazi, a otkup SA ambalazom je konacan (AMB-10-ODL-25) --
' to meri BFP Test_Amb_UndoStornaOdbijenNadKnjigom. Do 10e-2 je test sejao red
' STAROG oblika i tvrdio zastavicu Stornirano na njemu; zastavica je obrisana, a
' red starog oblika posle 10e-3 nije ni valjan red knjige.
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
    TcSeedRow TBL_NOVAC, Array(COL_NOV_ID, COL_NOV_OTKUP_ID), _
              Array("SVT-SJ-NID", "SVT-SJ-OID")

    ' --- STORNO (journaled) ---
    TcChk StornoOtkup_TX("SVT-SJ-OID") = True, "StornoOtkup_TX -> True"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-SJ-OID", COL_STORNIRANO))) = "DA", "otkup stornirano"
    TcChk NzS(LookupValue(TBL_NOVAC, COL_NOV_ID, "SVT-SJ-NID", COL_NOV_OTKUP_ID)) = "", "novac OtkupID obrisan (storno)"
    TcChk Len(LatestOpFor(DOK_TIP_OTKUP, "SVT-SJ-B")) > 0, "zurnal operacija zabelezena"

    ' --- UNDO (lossless preko zurnala) ---
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-SJ-B") = True, "UndoStorno_TX (zurnal) -> True"
    TcChk UCase$(NzS(LookupValue(TBL_OTKUP, COL_OTK_ID, "SVT-SJ-OID", COL_STORNIRANO))) <> "DA", "otkup vracen (aktivan)"
    TcChk NzS(LookupValue(TBL_NOVAC, COL_NOV_ID, "SVT-SJ-NID", COL_NOV_OTKUP_ID)) = "SVT-SJ-OID", "novac OtkupID VRACEN (lossless)"

    ' P2 7: ponovni undo iste op -> odbijen (drift guard: Stornirano je sada "" != NovaVrednost "Da")
    TcChk UndoStorno_TX(DOK_TIP_OTKUP, "SVT-SJ-B") = False, "ponovni undo iste op -> odbijen (drift)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_StornoJournalUndo_Auto GRESKA: " & Err.description: mFail = mFail + 1
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
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA_STAVKE
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_VEZE
    tx.AddTableSnapshot TBL_ZBIRNA_IZVORI

    ' Prijemnica se seje u NOVOM obliku, zaglavlje + stavka (S6 korak 5): kapija
    ' storna pita izvedenu fakturisanost, a ona cita stavke strogim citacem, koji
    ' zaglavlje bez stavke odbija (invarijanta dokumenta iz koraka 3).

    ' --- Lanac A: PONISTENJE -> uzvodna kaskada ---
    TcSeedZbirna "SVT-KA-ZID", "SVT-KA-Z", "I", 100, 10
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE), _
              Array("SVT-KA-OID", "SVT-KA-O", "SVT-KA-Z")
    TcSeedClanstvo "SVT-KA-ZID", "SVT-KA-OID"
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_BROJ_ZBIRNE), _
              Array("SVT-KA-PID", "SVT-KA-P", "I", "SVT-KA-Z")
    TcSeedStavkaPrijemnice "SVT-KA-PID", 1, "I", 100

    ' Identitet je PrijemnicaID izabranog reda (S6 korak 4) -- bez njega okvir
    ' dokument ne nalazi.
    Dim rA As Object: Set rA = RunPrijemnicaCorrection("SVT-KA-P", SV_MODE_PONISTENJE, True, False, "SVT-KA-PID")
    TcChk CBool(rA("success")), "PONISTENJE prijemnice -> success"
    TcChk TcCountActive(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-KA-Z") = 0, "zbirna stornirana (uzvodna kaskada)"
    TcChk TcOtpAktivnihUZbirni("SVT-KA-ZID") = 0, "otpremnica te zbirne stornirana"
    TcChk TcCountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-KA-P") = 0, "prijemnica stornirana"

    ' --- Lanac B: DUPLI -> NAMERNO list (zbirna/otpremnica prezivljavaju) ---
    TcSeedZbirna "SVT-KB-ZID", "SVT-KB-Z", "I", 100, 10
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE), _
              Array("SVT-KB-OID", "SVT-KB-O", "SVT-KB-Z")
    TcSeedClanstvo "SVT-KB-ZID", "SVT-KB-OID"
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KLASA, COL_PRJ_BROJ_ZBIRNE), _
              Array("SVT-KB-PID", "SVT-KB-P", "I", "SVT-KB-Z")
    TcSeedStavkaPrijemnice "SVT-KB-PID", 1, "I", 100

    Dim rB As Object: Set rB = RunPrijemnicaCorrection("SVT-KB-P", SV_MODE_DUPLI, True, False, "SVT-KB-PID")
    TcChk CBool(rB("success")), "DUPLI prijemnica -> success"
    TcChk TcCountActive(TBL_ZBIRNA, COL_ZBR_BROJ, "SVT-KB-Z") = 1, "DUPLI: zbirna ostaje AKTIVNA (list, ne kaskada)"
    TcChk TcOtpAktivnihUZbirni("SVT-KB-ZID") = 1, "DUPLI: otpremnica ostaje AKTIVNA"
    TcChk TcCountActive(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-KB-P") = 0, "DUPLI: prijemnica stornirana (list)"

    tx.RollbackTx: Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    Debug.Print "FAIL Test_PonistenjePrijemniceKaskada_Auto GRESKA: " & Err.description: mFail = mFail + 1
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

' Impact agregator: header + summary iz stvarnih (seed) redova.
Public Sub Test_BuildStornoImpact_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA_STAVKE
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KUPAC, COL_PRJ_BROJ_ZBIRNE), _
              Array("SVT-BI-ID", "SVT-BI-P1", "SVT-BI-KUP", "SVT-BI-Z1")
    TcSeedStavkaPrijemnice "SVT-BI-ID", 1, "I", 123

    ' Identitet je PrijemnicaID izabranog reda (S6 korak 4), kolicina sa stavki.
    Dim m As Object: Set m = BuildStornoImpact(FLOW_DOC_PRIJEMNICA, "SVT-BI-P1", "", "SVT-BI-ID")
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

' Browse: JEDAN RED PO DOKUMENTU, filter tip, iskljuci stornirano.
'
' Do S6 koraka 4 je lista spajala redove PO BROJU, jer su dva reda istog broja
' bila dve KLASE jednog dokumenta. Od koraka 2 je dokument jedno zaglavlje, pa
' su dva reda istog broja dva DOKUMENTA (dva kupca) -- i spajanje po broju je
' jedan od njih precutalo. Promena premise testa, ne samo tvrdnje.
Public Sub Test_GetActiveDocumentsForStorno_Auto()
    Dim tx As clsTransaction
    On Error GoTo EH
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KUPAC), Array("SVT-GA-1", "SVT-GA-P1", "SVT-GA-KUP1")
    TcSeedRow TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_KUPAC), Array("SVT-GA-2", "SVT-GA-P1", "SVT-GA-KUP2")   ' isti broj, drugi kupac
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
    TcChk nP1 = 2, "jedan red po dokumentu: P1 (dva kupca, isti broj) -> 2x"
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
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE
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
    tx.AddTableSnapshot TBL_ZBIRNA_STAVKE
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
    TcSeedRow TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_BROJ), _
              Array("SVT-DR-OTP", "SVT-DR-O1")                     ' aktivna otpremnica
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
' Zbirna u NOVOM obliku (S4-1): zaglavlje + stavka. Isti razlog kao
' modTestStorno.SeedZbirna -- strog citalac sadrzaja odbija zaglavlje bez
' stavki, pa bi golo zaglavlje oborilo svaki test koji dodirne liste, uvid ili
' prefill. Kolone zaglavlja ostaju popunjene dok ih okvir jos cita (S4-3).
Private Sub TcSeedZbirna(ByVal zbrID As String, ByVal broj As String, _
                         ByVal klasa As String, ByVal kg As Double, _
                         ByVal amb As Long)
    TcSeedRow TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_KLASA, COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB), _
              Array(zbrID, broj, klasa, kg, amb)
    TcSeedRow TBL_ZBIRNA_STAVKE, _
              Array(COL_ZBS_ID, COL_ZBS_ZBIRNA_ID, COL_ZBS_RB, COL_ZBS_KLASA, _
                    COL_ZBS_KOLICINA, COL_ZBS_KOL_AMB), _
              Array(zbrID & "-S1", zbrID, 1, klasa, kg, amb)
End Sub

' Stavka prijemnice uz seed zaglavlja (S6 korak 4): uvid cita stavke, a strogi
' citalac stavki trazi neprazan ID, kolicinu > 0 i cenu kao broj.
Private Sub TcSeedStavkaPrijemnice(ByVal prijemnicaID As String, ByVal rb As Long, _
                                   ByVal klasa As String, ByVal kolicina As Double)
    TcSeedRow TBL_PRIJEMNICA_STAVKE, _
              Array(COL_PRS_ID, COL_PRS_PRIJEMNICA_ID, COL_PRS_RB, COL_PRS_KLASA, _
                    COL_PRS_KOLICINA, COL_PRS_CENA, COL_PRS_KOL_AMB), _
              Array(prijemnicaID & "-S" & CStr(rb), prijemnicaID, rb, klasa, kolicina, 0, 0)
End Sub

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
        mFailImena = mFailImena & " | " & nm
        Debug.Print "FAIL " & nm
    End If
End Sub

Private Function NzS(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzS = "" Else NzS = Trim$(CStr(v))
End Function

' Broj AKTIVNIH (ne-storniranih) redova gde col=val (CountActive u modStornoFlow je Private).
' Red clanstva otpremnica -> zbirna. Od S5-3b kaskada bira decu odavde, pa
' fixture koji upise samo labelu pravi zbirnu BEZ ijednog izvora.
Private Sub TcSeedClanstvo(ByVal zbirnaID As String, ByVal otpremnicaID As String)
    TcSeedRow TBL_ZBIRNA_IZVORI, _
              Array(COL_ZBI_ID, COL_ZBI_ZBIRNA_ID, COL_ZBI_OTPREMNICA_ID), _
              Array(otpremnicaID & "-ZI", zbirnaID, otpremnicaID)
End Sub

' Koliko je otpremnica te zbirne JOS AKTIVNO.
'
' Namerno ZbrClanovi, a ne AktivnaZbirnaZaOtpremnicu: ova druga izbacuje decu
' stornirane zbirne, pa bi posle PONISTENJA vratila 0 i onda kad otpremnica NIJE
' stornirana -- tvrdnja bi prolazila iz pogresnog razloga. Meri se storniranost
' DETETA, kao i pre; menja se samo po cemu se dete prepoznaje kao njeno.
Private Function TcOtpAktivnihUZbirni(ByVal zbirnaID As String) As Long
    Dim clan As Variant, n As Long
    For Each clan In KolekcijaUNizTc(modDokumenta.ZbrClanovi(zbirnaID))
        If StrComp(Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                     CStr(clan), COL_STORNIRANO))), "Da", vbTextCompare) <> 0 Then
            n = n + 1
        End If
    Next clan
    TcOtpAktivnihUZbirni = n
End Function

Private Function KolekcijaUNizTc(ByVal c As Collection) As Variant
    If c Is Nothing Then KolekcijaUNizTc = Array(): Exit Function
    If c.count = 0 Then KolekcijaUNizTc = Array(): Exit Function

    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    KolekcijaUNizTc = a
End Function

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
