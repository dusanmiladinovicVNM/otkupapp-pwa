Attribute VB_Name = "modTestStorno"
Option Explicit

' ============================================================
' modTestStorno - test suite za centralni storno/ispravka framework
'
' Pokretanje: Alt+F8 -> RunStornoTestSuite
'
' Sopstveni harness (mPass/mFail/mReport + Chk/ChkEq/ChkEqD), NEZAVISAN od
' modTestPalete (koji ima Private accumulatore). Ceo run je u JEDNOJ transakciji
' koja se UVEK rollback-uje -> nula ostavljenih podataka. Inner *_TX commituju
' svoje TX; spoljni snapshot-restore ih ponisti (isti obrazac kao modTestPalete).
'
' Prefiks svih test podataka: "SVT-" (Storno-Veze-Test) -> izolacija.
'
' Pokriva 8 scenarija (vidi zadatak FAZA 5):
'  T01 storno otpremnice rekalkulise zbirnu
'  T02 prevezivanje otpremnice validira OBE zbirne
'  T03 storno zbirne sa aktivnim otpremnicama ne ostavlja mismatch
'  T04 ispravka zbirne prevezuje otpremnice i prijemnicu
'  T05 paleta-stavke dobijaju novu zbirnu pri ispravci
'  T06 revers ispravka ne duplira saldo
'  T07 revers ponistenje uklanja uticaj na saldo
'  T08 pending correction ostaje vidljiv ako flow ne uspe
' ============================================================

Private mPass As Long
Private mFail As Long
Private mFails As String
Private mReport As String

' Gate: suite mora da PODIGNE gresku kad provera padne, inace je runner vidi kao
' "blind" -- proslo bez greske, sto NIJE isto sto i sve provere prosle. Konvencija
' preuzeta iz modTestBanka.ERR_BIT_SUITE_FAILED.
Private Const ERR_STORNO_SUITE_FAILED As Long = vbObjectError + 2962

' Deterministicki hladnjaca-kupac za testove (config se snapshotuje pa vraca
' rollback-om). Zbirna sa ovim KupacID = interni hladnjaca-tok (kaskada); bilo koji
' drugi KupacID = eksterni kupac (prijemnica netaknuta).
Private Const HLAD_KUP As String = "SVT-HLAD-KUPAC"

Public Sub RunStornoTestSuite()
    Dim tx As clsTransaction
    On Error GoTo EH

    ' Tabela context-a mora postojati PRE snapshot-a (AddTableSnapshot cita tabelu).
    modSetup.EnsureStornoVezeSchemaCore

    If GetTable(TBL_OTPREMNICA) Is Nothing Or GetTable(TBL_ZBIRNA) Is Nothing Then
        MsgBox "Tabele otpremnica/zbirna ne postoje. Prekid.", vbExclamation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Pokrenuti storno/ispravka test suite?" & vbCrLf & vbCrLf & _
              "Svi podaci (SVT-*) se prave u transakciji i UVEK se ponistavaju " & _
              "(rollback). Nista ne ostaje u tabelama.", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    mPass = 0: mFail = 0: mFails = "": mReport = ""

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_SEF_CONFIG
    tx.AddTableSnapshot TBL_STORNO_VEZE
    tx.AddTableSnapshot TBL_NOVAC              ' T23-T27 (storno izvoda)
    tx.AddTableSnapshot TBL_BANKA_IMPORT       ' T23-T27 (storno izvoda)

    ' Deterministicki config (rollback vraca original): hladnjaca-kupac za gate,
    ' i MALINA_MODE=NO da simple storno otpremnice NE kaskadira zbirnu preko malina
    ' putanje -> testira se framework-ovo storno/rekalk ponasanje (T22).
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, HLAD_KUP
    SetConfigValue CFG_KEY_MALINA_MODE, "NO"

    T01_StornoOtpremniceRekalkuliseZbirnu
    T02_PrevezivanjeValidiraObeZbirne
    T03_StornoZbirneBezMismatch
    T04_IspravkaZbirnePrevezuje
    T05_PaletaStavkeNovaZbirna
    T06_ReversIspravkaNeDupliraSaldo
    T07_ReversPonistenjeUklanjaSaldo
    T08_PendingCorrectionVidljivNaFail
    T09_SimpleStornoZbirna
    T10_SmartTriggerGate
    T11_OtpremnicaIspravkaPrevezujePrijemnicuIPalete
    T12_ReversCompleteTraziAktivanNoviRevers
    T13_ReversCompleteSaAktivnimNovimReversom
    T14_AutoCompleteNeBiraLatestKadImaVisePending
    T15_PonistenjeUvekTraziPotvrdu
    T16_DupliOtpremniceOslobadjaBlokove
    T17_PonistenjeZbirnaHladnjacaKaskada
    T18_PonistenjeZbirnaEksterniNeDiraPrijemnicu
    T19_DupliVsPonistenjeZbirnaOtpremnice
    T20_PonistenjeOtpremniceDeljenaNeObaraZbirnu
    T21_PonistenjeOtpremniceJedinaKaskadaCeoTok
    T22_SimpleStornoOtpremniceNeOstavljaZbirnu00
    T23_StornoIzvodaRemapVracaStavkeUObradu
    T24_StornoIzvodaReimportOslobadjaPonovniUvoz
    T25_AvansSplitNasledjujeMarkerIPada
    T26_IzvodNaViseRacunaTraziRacun
    T27_StornoIzvodaOsvezavaOtkup
    T28_OMAvansBrojiObaKanala
    T29_OMAvansUIzvestajimaObaKanala
    T30_StagingBezPKOdbijaStorno
    T31_IzvodBezRacunaOdbijen
    T32_SamoMarkerOdredjujePripadnost
    T33_ObradjenaStavkaBezNovcaBlokira
    T34_BrojIzvodaSaKosomCrtom
    T35_NeslaganjeIznosaBlokira
    T36_RucniRedSaIstimBrojemOstajeStornirljiv

    tx.RollbackTx
    Set tx = Nothing

    On Error GoTo 0
    ReportResults

    If mFail > 0 Then
        Err.Raise ERR_STORNO_SUITE_FAILED, "modTestStorno.RunStornoTestSuite", _
            "RunStornoTestSuite: " & CStr(mFail) & " provera palo (PASS=" & _
            CStr(mPass) & "). Detalji u Immediate prozoru."
    End If

    Exit Sub
EH:
    Dim errDesc As String
    errDesc = Err.description

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "modTestStorno.RunStornoTestSuite"
    On Error GoTo 0

    ' Prekid je pad, ne "nije se desilo": ubroji ga i podigni, da runner ne
    ' prijavi zeleno za suite koja nije dosla do kraja.
    Fail "SUITE prekinut greskom: " & errDesc
    ReportResults

    Err.Raise ERR_STORNO_SUITE_FAILED, "modTestStorno.RunStornoTestSuite", _
        "RunStornoTestSuite prekinut: " & errDesc
End Sub

' ============================================================
' T01 - storno (DUPLI) jedne otpremnice rekalkulise zbirnu na preostale.
' ============================================================
Private Sub T01_StornoOtpremniceRekalkuliseZbirnu()
    Const S As String = "T01 storno otpremnice -> recalc zbirne: "

    SeedZbirna "SVT-Z1", "I", 100, 10
    SeedOtpremnica "SVT-OA1", "SVT-Z1", "I", 60, 6
    SeedOtpremnica "SVT-OB1", "SVT-Z1", "I", 40, 4

    ' Pocetno stanje: 100 = 60 + 40 -> OK
    Chk modDokumentInvariant.IsZbirnaConsistent("SVT-Z1"), S & "pocetni invariant OK"

    ' Storno OB1 kao DUPLI/FANTOM -> zbirna se rekalkulise na preostalu OA1 (60/6).
    Dim res As Object
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-OB1", SV_MODE_DUPLI)
    Chk CBool(res("success")), S & "RunOtpremnicaCorrection(DUPLI) uspeo"

    Dim inv As Object: Set inv = modDokumentInvariant.ValidateZbirnaInvariant("SVT-Z1")
    ChkEqD CDbl(inv("kgZbrI")), 60, S & "zbirna KG I rekalkulisana na 60"
    ChkEq CLng(inv("ambZbrTotal")), 6, S & "zbirna AMB rekalkulisana na 6"
    Chk CBool(inv("isValid")), S & "invariant OK posle storna (nema tihi mismatch)"
End Sub

' ============================================================
' T02 - ValidateOtpremnicaZbirnaImpact hvata mismatch na OBE zbirne.
' ============================================================
Private Sub T02_PrevezivanjeValidiraObeZbirne()
    Const S As String = "T02 prevezivanje validira obe zbirne: "

    SeedZbirna "SVT-Z2A", "I", 100, 10
    SeedOtpremnica "SVT-OA2", "SVT-Z2A", "I", 60, 6
    SeedOtpremnica "SVT-OB2", "SVT-Z2A", "I", 40, 4
    SeedZbirna "SVT-Z2B", "I", 50, 5
    SeedOtpremnica "SVT-OC2", "SVT-Z2B", "I", 50, 5

    Dim impact As Object
    Set impact = modDokumentInvariant.ValidateOtpremnicaZbirnaImpact("SVT-Z2A", "SVT-Z2B")
    Chk CBool(impact("bothValid")), S & "obe zbirne validne pre premestaja"

    ' Premesti OB2 (40) sa Z2A na Z2B BEZ rekalkulacije -> obe postaju mismatch.
    ForceSetOtpremnicaZbirna "SVT-OB2", "SVT-Z2B"

    Set impact = modDokumentInvariant.ValidateOtpremnicaZbirnaImpact("SVT-Z2A", "SVT-Z2B")
    Chk Not CBool(impact("oldValid")), S & "stara zbirna (Z2A) sada mismatch"
    Chk Not CBool(impact("newValid")), S & "nova zbirna (Z2B) sada mismatch"
    Chk Not CBool(impact("bothValid")), S & "bothValid = False (uhvacen mismatch)"
End Sub

' ============================================================
' T03 - storno zbirne (DUPLI) sa aktivnim otpremnicama: otpremnice se odvezu
' ("ceka zbirnu"), ne ostaje aktivna zbirna sa mismatch-om.
' ============================================================
Private Sub T03_StornoZbirneBezMismatch()
    Const S As String = "T03 storno zbirne bez mismatch: "

    SeedZbirna "SVT-Z3", "I", 100, 10
    SeedOtpremnica "SVT-OA3", "SVT-Z3", "I", 60, 6
    SeedOtpremnica "SVT-OB3", "SVT-Z3", "I", 40, 4

    Dim res As Object
    Set res = modStornoFlow.RunZbirnaCorrection("SVT-Z3", SV_MODE_DUPLI)
    Chk CBool(res("success")), S & "RunZbirnaCorrection(DUPLI) uspeo"

    Chk Not ZbirnaPostoji("SVT-Z3"), S & "zbirna vise nije aktivna (stornirana)"
    ChkEq OtpBrojZbirne("SVT-OA3"), "", S & "OA3 vracena u 'ceka zbirnu' (BrojZbirne prazno)"
    ChkEq OtpBrojZbirne("SVT-OB3"), "", S & "OB3 vracena u 'ceka zbirnu' (BrojZbirne prazno)"
End Sub

' ============================================================
' T04 - ispravka zbirne: nova zbirna, otpremnice + prijemnica prevezane na nju,
' invariant nove zbirne OK.
' ============================================================
Private Sub T04_IspravkaZbirnePrevezuje()
    Const S As String = "T04 ispravka zbirne prevezuje: "

    SeedZbirna "SVT-Z4", "I", 100, 10
    SeedOtpremnica "SVT-OA4", "SVT-Z4", "I", 60, 6
    SeedOtpremnica "SVT-OB4", "SVT-Z4", "I", 40, 4
    SeedPrijemnica "SVT-P4", "SVT-Z4", "I", 100, 10

    ' Faza 1: storno stare + context
    Dim res As Object
    Set res = modStornoFlow.RunZbirnaCorrection("SVT-Z4", SV_MODE_ISPRAVKA)
    Chk CBool(res("needsForm")), S & "ISPRAVKA trazi novu zbirnu (needsForm)"
    Dim cid As String: cid = CStr(res("correctionID"))
    Chk Len(cid) > 0, S & "correction context kreiran"

    ' Operater snima NOVU zbirnu (drugaciji broj) -> recalc ce popuniti tacne KG.
    SeedZbirna "SVT-Z4B", "I", 0, 0

    ' Faza 2: complete -> prevezi otpremnice+prijemnicu, recalc, validiraj.
    Set res = modStornoFlow.CompleteZbirnaIspravka(cid, "SVT-Z4B")
    Chk CBool(res("success")), S & "CompleteZbirnaIspravka uspeo"

    ChkEq OtpBrojZbirne("SVT-OA4"), "SVT-Z4B", S & "OA4 prevezana na novu zbirnu"
    ChkEq OtpBrojZbirne("SVT-OB4"), "SVT-Z4B", S & "OB4 prevezana na novu zbirnu"
    ChkEq PrjBrojZbirne("SVT-P4"), "SVT-Z4B", S & "prijemnica prevezana na novu zbirnu"
    Chk modDokumentInvariant.IsZbirnaConsistent("SVT-Z4B"), S & "nova zbirna = zbir otpremnica (OK)"
End Sub

' ============================================================
' T05 - pri ispravci zbirne, paleta-stavke prijemnice dobijaju novu zbirnu.
' ============================================================
Private Sub T05_PaletaStavkeNovaZbirna()
    Const S As String = "T05 paleta-stavke nova zbirna: "

    SeedZbirna "SVT-Z5", "I", 100, 10
    SeedOtpremnica "SVT-OA5", "SVT-Z5", "I", 100, 10
    SeedPrijemnica "SVT-P5", "SVT-Z5", "I", 100, 10
    SeedPaletaStavka "SVT-PS5", "SVT-P5", "SVT-Z5", "I", 100, 10

    Dim res As Object
    Set res = modStornoFlow.RunZbirnaCorrection("SVT-Z5", SV_MODE_ISPRAVKA)
    Dim cid As String: cid = CStr(res("correctionID"))

    SeedZbirna "SVT-Z5B", "I", 0, 0
    Set res = modStornoFlow.CompleteZbirnaIspravka(cid, "SVT-Z5B")
    Chk CBool(res("success")), S & "CompleteZbirnaIspravka uspeo"

    ChkEq PalsBrojZbirne("SVT-PS5"), "SVT-Z5B", S & "paleta-stavka dobila novu zbirnu"
    ChkEq PrjBrojZbirne("SVT-P5"), "SVT-Z5B", S & "prijemnica prevezana na novu zbirnu"
End Sub

' ============================================================
' T06 - revers ispravka: stari storniran, novi aktivan -> saldo racuna SAMO novi.
' ============================================================
Private Sub T06_ReversIspravkaNeDupliraSaldo()
    Const S As String = "T06 revers ispravka ne duplira saldo: "

    ' Stari revers: kooperant SVT-K6 Ulaz 10 (tip SVT-A), stanica Izlaz 10.
    SeedRevers "SVT-R6", DOK_TIP_OM_IZLAZ_KOOP, "SVT-K6", "SVT-S6", "SVT-A", 10
    ChkEq AmbSaldo("SVT-K6", "Kooperant", "SVT-A"), 10, S & "baseline saldo K6 = 10"

    Dim res As Object
    Set res = modStornoFlow.RunReversCorrection("SVT-R6", DOK_TIP_OM_IZLAZ_KOOP, SV_MODE_ISPRAVKA)
    Chk CBool(res("needsForm")), S & "ISPRAVKA reversa trazi novi (needsForm)"
    ChkEq AmbSaldo("SVT-K6", "Kooperant", "SVT-A"), 0, S & "posle storna starog saldo = 0"

    ' Operater unosi NOVI revers 12.
    SeedRevers "SVT-R6B", DOK_TIP_OM_IZLAZ_KOOP, "SVT-K6", "SVT-S6", "SVT-A", 12
    Dim cid As String: cid = CStr(res("correctionID"))
    Set res = modStornoFlow.CompleteReversIspravka(cid, "SVT-R6B")
    Chk CBool(res("success")), S & "CompleteReversIspravka uspeo"

    ChkEq AmbSaldo("SVT-K6", "Kooperant", "SVT-A"), 12, S & "saldo = 12 (samo novi, NE 22)"
End Sub

' ============================================================
' T07 - revers ponistenje uklanja uticaj na saldo (bez kontra-stavke).
' ============================================================
Private Sub T07_ReversPonistenjeUklanjaSaldo()
    Const S As String = "T07 revers ponistenje uklanja saldo: "

    SeedRevers "SVT-R7", DOK_TIP_OM_IZLAZ_KOOP, "SVT-K7", "SVT-S7", "SVT-A", 8
    ChkEq AmbSaldo("SVT-K7", "Kooperant", "SVT-A"), 8, S & "baseline saldo K7 = 8"

    Dim res As Object
    Set res = modStornoFlow.RunReversCorrection("SVT-R7", DOK_TIP_OM_IZLAZ_KOOP, SV_MODE_PONISTENJE)
    Chk CBool(res("success")), S & "RunReversCorrection(PONISTENJE) uspeo"

    ChkEq AmbSaldo("SVT-K7", "Kooperant", "SVT-A"), 0, S & "saldo = 0 posle ponistenja"
End Sub

' ============================================================
' T08 - neuspesan flow: pending/failed context ostaje vidljiv u recovery listi.
' ============================================================
Private Sub T08_PendingCorrectionVidljivNaFail()
    Const S As String = "T08 pending correction vidljiv na fail: "

    Dim cid As String
    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_OTPREMNICA, _
            "SVT-OTPID8", "SVT-OTP8", , , , FLOW_DOC_ZBIRNA, , "SVT-Z8", "Test fail flow.")
    Chk Len(cid) > 0, S & "context kreiran (PENDING)"

    Dim n0 As Long: n0 = modStornoContext.CountPendingRecovery()
    Chk n0 >= 1, S & "PENDING context se broji u recovery (>=1)"

    ' Simuliraj neuspeh automatskog koraka.
    Chk modStornoContext.FailCorrectionContext(cid, "Simulirani neuspeh reassign-a."), S & "FailCorrectionContext uspeo"

    ChkEq modStornoContext.GetCorrectionField(cid, COL_SV_STATUS), SV_STATUS_FAILED, S & "status = FAILED"
    ChkEq modStornoContext.GetCorrectionField(cid, COL_SV_NEEDS_RECOVERY), "Da", S & "NeedsRecovery = Da"

    ' I dalje vidljiv u recovery listi (nije tihi nestanak).
    Dim found As Boolean, c As Collection, i As Long
    Set c = modStornoContext.GetPendingCorrections()
    For i = 1 To c.count
        If CStr(c(i)("id")) = cid Then found = True
    Next i
    Chk found, S & "FAILED context ostaje u recovery listi (vidljiv)"
End Sub

' ============================================================
' T09 - SIMPLE storno zbirne (bez dijaloga): storno + odvezivanje otpremnica.
' ============================================================
Private Sub T09_SimpleStornoZbirna()
    Const S As String = "T09 simple storno zbirne: "

    SeedZbirna "SVT-Z9", "I", 100, 10
    SeedOtpremnica "SVT-OA9", "SVT-Z9", "I", 60, 6
    SeedOtpremnica "SVT-OB9", "SVT-Z9", "I", 40, 4

    ' Nema prijemnice/paleta -> smart trigger NE trazi dijalog.
    Chk Not modStornoFlow.CorrectionNeedsDialog(FLOW_DOC_ZBIRNA, "SVT-Z9"), _
        S & "CorrectionNeedsDialog = False (nema nizvodnog toka)"

    Dim res As Object
    Set res = modStornoFlow.RunSimpleStornoZbirna("SVT-Z9")
    Chk CBool(res("success")), S & "RunSimpleStornoZbirna uspeo"

    Chk Not ZbirnaPostoji("SVT-Z9"), S & "zbirna stornirana"
    ChkEq OtpBrojZbirne("SVT-OA9"), "", S & "OA9 odvezana ('ceka zbirnu')"
    ChkEq OtpBrojZbirne("SVT-OB9"), "", S & "OB9 odvezana ('ceka zbirnu')"
End Sub

' ============================================================
' T10 - smart trigger gate: dijalog se trazi TEK kad postoji nizvodni tok.
' ============================================================
Private Sub T10_SmartTriggerGate()
    Const S As String = "T10 smart trigger gate: "

    SeedZbirna "SVT-Z10", "I", 100, 10
    SeedOtpremnica "SVT-OA10", "SVT-Z10", "I", 100, 10

    ' Bez prijemnice/paleta -> otpremnica NE trazi dijalog (obican storno).
    Chk Not modStornoFlow.CorrectionNeedsDialog(FLOW_DOC_OTPREMNICA, "SVT-OA10"), _
        S & "otpremnica bez nizvodnog toka -> False"

    ' Dodaj prijemnicu preko zbirne -> sada TRAZI dijalog (odluka o prijemnici).
    SeedPrijemnica "SVT-P10", "SVT-Z10", "I", 100, 10
    Chk modStornoFlow.CorrectionNeedsDialog(FLOW_DOC_OTPREMNICA, "SVT-OA10"), _
        S & "otpremnica sa prijemnicom -> True (eskalira na dijalog)"
    Chk modStornoFlow.CorrectionNeedsDialog(FLOW_DOC_ZBIRNA, "SVT-Z10"), _
        S & "zbirna sa prijemnicom -> True"

    ' Revers nikad ne trazi dijalog (list, bez nizvodnog toka).
    Chk Not modStornoFlow.CorrectionNeedsDialog(FLOW_DOC_REVERS, "SVT-R10", DOK_TIP_OM_IZLAZ_KOOP), _
        S & "revers -> uvek False (nema lanca)"
End Sub

' ============================================================
' T11 - ISPRAVKA otpremnice: prevezuje otkupne listove, prijemnicu I paleta-stavke
' na novu zbirnu; stara zbirna se STORNIRA (ne ostaje aktivna 0/0). Regres test.
' ============================================================
Private Sub T11_OtpremnicaIspravkaPrevezujePrijemnicuIPalete()
    Const S As String = "T11 ispravka otpremnice prevezuje prijemnicu/palete: "

    ' Stara: otpremnica OA11 (kg 100) u zbirni Z11 + otkupni list + prijemnica + paleta-stavka.
    SeedZbirna "SVT-Z11", "I", 100, 10
    SeedOtpremnica "SVT-OA11", "SVT-Z11", "I", 100, 10
    SeedOtkupBlok "SVT-BLK11", "SVT-OA11-ID-I", "SVT-Z11"
    SeedPrijemnica "SVT-P11", "SVT-Z11", "I", 100, 10
    SeedPaletaStavka "SVT-PS11", "SVT-P11", "SVT-Z11", "I", 100, 10

    ' Faza 1: ISPRAVKA_ODMAH -> storno stare otpremnice + context.
    Dim res As Object
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-OA11", SV_MODE_ISPRAVKA)
    Chk CBool(res("needsForm")), S & "ISPRAVKA trazi novu otpremnicu"
    Dim cid As String: cid = CStr(res("correctionID"))

    ' Operater snima NOVU otpremnicu (manji kg 90) sa NOVOM zbirnom (malina 1:1).
    SeedOtpremnica "SVT-OB11", "SVT-Z11B", "I", 90, 9
    SeedZbirna "SVT-Z11B", "I", 0, 0

    ' Faza 2: complete -> prevezi blok + prijemnicu/palete, recalc nove, storno stare.
    Set res = modStornoFlow.CompleteOtpremnicaIspravka(cid, "SVT-OB11")
    Chk CBool(res("success")), S & "CompleteOtpremnicaIspravka uspeo"

    ' otkupni listovi na novoj otpremnici
    ChkEq OtkOtpremnicaID("SVT-BLK11"), "SVT-OB11-ID-I", S & "otkupni list prevezan na novu otpremnicu"
    ' nova zbirna konzistentna (= zbir svojih otpremnica = 90)
    Chk modDokumentInvariant.IsZbirnaConsistent("SVT-Z11B"), S & "nova zbirna = zbir otpremnica (90)"
    ' aktivna prijemnica ima BrojZbirne = nova zbirna
    ChkEq PrjBrojZbirne("SVT-P11"), "SVT-Z11B", S & "prijemnica.BrojZbirne = nova zbirna"
    ' paletaStavka ima BrojZbirne = nova zbirna
    ChkEq PalsBrojZbirne("SVT-PS11"), "SVT-Z11B", S & "paletaStavka.BrojZbirne = nova zbirna"
    ' stara zbirna NE ostaje aktivna 0/0 -> stornirana
    Chk Not ZbirnaPostoji("SVT-Z11"), S & "stara zbirna STORNIRANA (ne aktivna 0/0)"
    ' context COMPLETED (jer je downstream relink uspeo)
    ChkEq modStornoContext.GetCorrectionField(cid, COL_SV_STATUS), SV_STATUS_COMPLETED, _
        S & "context COMPLETED (downstream relink uspeo)"
End Sub

' ============================================================
' T12 - CompleteReversIspravka NE sme kompletirati ako novi revers ne postoji.
' ============================================================
Private Sub T12_ReversCompleteTraziAktivanNoviRevers()
    Const S As String = "T12 revers complete trazi aktivan novi revers: "

    SeedRevers "SVT-R12", DOK_TIP_OM_IZLAZ_KOOP, "SVT-K12", "SVT-S12", "SVT-A", 10
    Dim res As Object
    Set res = modStornoFlow.RunReversCorrection("SVT-R12", DOK_TIP_OM_IZLAZ_KOOP, SV_MODE_ISPRAVKA)
    Dim cid As String: cid = CStr(res("correctionID"))
    Chk Len(cid) > 0, S & "context kreiran"

    ' NE seed-ujemo novi revers -> complete mora da odbije.
    Set res = modStornoFlow.CompleteReversIspravka(cid, "SVT-NEMA-REV")
    Chk Not CBool(res("success")), S & "success=False (novi revers ne postoji)"
    Chk modStornoContext.GetCorrectionField(cid, COL_SV_STATUS) <> SV_STATUS_COMPLETED, _
        S & "status NIJE COMPLETED"
    ChkEq modStornoContext.GetCorrectionField(cid, COL_SV_STATUS), SV_STATUS_MANUAL, S & "status MANUAL_REQUIRED"
    ChkEq modStornoContext.GetCorrectionField(cid, COL_SV_NEEDS_RECOVERY), "Da", S & "NeedsRecovery = Da"
End Sub

' ============================================================
' T13 - CompleteReversIspravka sa AKTIVNIM novim reversom -> COMPLETED; saldo
' racuna samo novi (ne stari + novi).
' ============================================================
Private Sub T13_ReversCompleteSaAktivnimNovimReversom()
    Const S As String = "T13 revers complete sa aktivnim novim: "

    SeedRevers "SVT-R13", DOK_TIP_OM_IZLAZ_KOOP, "SVT-K13", "SVT-S13", "SVT-A", 10
    Dim res As Object
    Set res = modStornoFlow.RunReversCorrection("SVT-R13", DOK_TIP_OM_IZLAZ_KOOP, SV_MODE_ISPRAVKA)
    Dim cid As String: cid = CStr(res("correctionID"))
    ChkEq AmbSaldo("SVT-K13", "Kooperant", "SVT-A"), 0, S & "posle storna starog saldo = 0"

    ' Operater unosi NOVI revers (12).
    SeedRevers "SVT-R13B", DOK_TIP_OM_IZLAZ_KOOP, "SVT-K13", "SVT-S13", "SVT-A", 12
    Set res = modStornoFlow.CompleteReversIspravka(cid, "SVT-R13B")
    Chk CBool(res("success")), S & "success=True (novi revers aktivan)"
    ChkEq modStornoContext.GetCorrectionField(cid, COL_SV_STATUS), SV_STATUS_COMPLETED, S & "status COMPLETED"
    ChkEq AmbSaldo("SVT-K13", "Kooperant", "SVT-A"), 12, S & "saldo = 12 (samo novi, NE 22)"
End Sub

' ============================================================
' T14 - safe-stop: kad ima VISE pending ispravki istog tipa, helper vraca count>1
' (UI tada ne bira naslepo najnoviji). CountPendingCorrectionsByDocType.
' ============================================================
Private Sub T14_AutoCompleteNeBiraLatestKadImaVisePending()
    Const S As String = "T14 multi-pending safe-stop: "

    Dim before As Long
    before = modStornoContext.CountPendingCorrectionsByDocType(FLOW_DOC_OTPREMNICA, SV_MODE_ISPRAVKA)

    Dim c1 As String, c2 As String
    c1 = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_OTPREMNICA, _
            "SVT-ID14A", "SVT-OTPA14", , , , FLOW_DOC_ZBIRNA, , "SVT-ZA14", "Test multi 1.")
    c2 = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_OTPREMNICA, _
            "SVT-ID14B", "SVT-OTPB14", , , , FLOW_DOC_ZBIRNA, , "SVT-ZB14", "Test multi 2.")
    Chk Len(c1) > 0 And Len(c2) > 0, S & "dva pending context-a kreirana"

    ChkEq modStornoContext.CountPendingCorrectionsByDocType(FLOW_DOC_OTPREMNICA, SV_MODE_ISPRAVKA), _
        before + 2, S & "count raste tacno za 2"
    Chk modStornoContext.CountPendingCorrectionsByDocType(FLOW_DOC_OTPREMNICA, SV_MODE_ISPRAVKA) > 1, _
        S & "count > 1 -> UI safe-stop (ne bira naslepo latest)"
End Sub

' ============================================================
' T15 - PONISTENJE UVEK trazi svesnu potvrdu (blocked) pre nego sto bilo sta uradi,
' cak i bez zavisnih dokumenata; sa forceConfirm se izvrsava. (Razlika od DUPLI.)
' NAPOMENA: ovo je BUSINESS-API garancija (RunOtpremnicaCorrection/RunZbirnaCorrection).
' UI (frmDokumenta) za dokument BEZ nizvodnog toka ide smart-trigger shortcut na
' obican storno (RunSimpleStorno*) i ne prikazuje 4-mode dijalog -- svesna odluka.
' ============================================================
Private Sub T15_PonistenjeUvekTraziPotvrdu()
    Const S As String = "T15 ponistenje uvek trazi potvrdu: "

    ' Lone otpremnica (zbirna ne postoji) -> nema zavisnih.
    SeedOtpremnica "SVT-OP15", "SVT-Z15-NEMA", "I", 50, 5

    ' Bez potvrde -> mora BLOKIRATI (prikaz posledica), nista ne dira.
    Dim res As Object
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-OP15", SV_MODE_PONISTENJE)
    Chk CBool(res("blocked")), S & "PONISTENJE blokira dok se ne potvrdi (i bez zavisnih)"
    Chk Not CBool(res("success")), S & "nista nije izvrseno pre potvrde"
    Chk LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-OP15", COL_OTP_ID) <> "", _
        S & "otpremnica jos aktivna (nije stornirana)"

    ' Sa svesnom potvrdom (forceConfirm) -> izvrsava se.
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-OP15", SV_MODE_PONISTENJE, True)
    Chk CBool(res("success")), S & "sa potvrdom (forceConfirm) ponistava"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-OP15", COL_OTP_ID), "", _
        S & "otpremnica sada stornirana"
End Sub

' ============================================================
' T16 - DUPLI otpremnica: storno + OSLOBODI otkup blokove (za reveze) + NE dira
' prijemnicu (nije kaskada); prazna zbirna -> STORNO (ne aktivna 0/0).
' ============================================================
Private Sub T16_DupliOtpremniceOslobadjaBlokove()
    Const S As String = "T16 DUPLI otpremnica oslobadja blokove: "

    SeedZbirna "SVT-Z16", "I", 100, 10, HLAD_KUP
    SeedOtpremnica "SVT-O16", "SVT-Z16", "I", 100, 10
    SeedOtkupBlok "SVT-BLK16", "SVT-O16-ID-I", "SVT-Z16"
    SeedPrijemnica "SVT-P16", "SVT-Z16", "I", 100, 10

    Dim res As Object
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-O16", SV_MODE_DUPLI)
    Chk CBool(res("success")), S & "DUPLI uspeo"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-O16", COL_OTP_ID), "", S & "otpremnica stornirana"
    ' blok OSLOBODJEN (OtpremnicaID prazan) ali i dalje AKTIVAN (realna kupovina).
    ChkEq OtkOtpremnicaID("SVT-BLK16"), "", S & "blok oslobodjen (OtpremnicaID prazan)"
    Chk LookupActiveID(TBL_OTKUP, COL_OTK_BR_DOK, "SVT-BLK16", COL_OTK_ID) <> "", S & "blok i dalje aktivan (nije storniran)"
    ' DUPLI NE kaskadira -> prijemnica ostaje aktivna (osirocena za reveze).
    Chk LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-P16", COL_PRJ_ID) <> "", S & "prijemnica NETAKNUTA (DUPLI ne kaskadira)"
    ' prazna zbirna -> STORNO (nulling fix), NE aktivna 0/0.
    Chk Not ZbirnaPostoji("SVT-Z16"), S & "prazna zbirna STORNIRANA (ne aktivna 0/0)"
End Sub

' ============================================================
' T17 - PONISTENJE zbirne (hladnjaca kupac): kaskadno stornira ceo interni tok
' (otpremnice + prijemnica + paletne stavke).
' ============================================================
Private Sub T17_PonistenjeZbirnaHladnjacaKaskada()
    Const S As String = "T17 PONISTENJE zbirna (hladnjaca) kaskada + palete: "

    SeedZbirna "SVT-Z17", "I", 100, 10, HLAD_KUP
    SeedOtpremnica "SVT-OA17", "SVT-Z17", "I", 60, 6
    SeedOtpremnica "SVT-OB17", "SVT-Z17", "I", 40, 4
    SeedPrijemnica "SVT-P17", "SVT-Z17", "I", 100, 10
    ' Kao u realnim podacima: DELJENA paleta (co-tenant druge zbirne, tip PAL-00150)
    ' i paleta koja sadrzi SAMO nasu robu (tip PAL-00151) -> posle skidanja prazna.
    SeedPaleta "SVT-PAL17A", 100, 200, 20, 100, PAL_STATUS_ZATVORENA
    SeedPaleta "SVT-PAL17B", 50, 100, 10, 50, PAL_STATUS_ZATVORENA
    SeedPaletaStavka "SVT-PS17A", "SVT-P17", "SVT-Z17", "I", 100, 10, "SVT-PAL17A", 50
    SeedPaletaStavka "SVT-PS17X", "SVT-P17X", "SVT-Z17X", "I", 100, 10, "SVT-PAL17A", 50
    SeedPaletaStavka "SVT-PS17B", "SVT-P17", "SVT-Z17", "I", 100, 10, "SVT-PAL17B", 50

    Dim res As Object
    Set res = modStornoFlow.RunZbirnaCorrection("SVT-Z17", SV_MODE_PONISTENJE, True)
    Chk CBool(res("success")), S & "PONISTENJE (forceConfirm) uspeo"
    Chk Not ZbirnaPostoji("SVT-Z17"), S & "zbirna stornirana"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-OA17", COL_OTP_ID), "", S & "OA17 STORNIRANA (kaskada)"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-OB17", COL_OTP_ID), "", S & "OB17 STORNIRANA (kaskada)"
    ChkEq LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-P17", COL_PRJ_ID), "", S & "prijemnica STORNIRANA (hladnjaca)"
    ' nase paletne stavke skinute (kroz paletni motor)
    ChkEq PalsStornirano("SVT-PS17A"), "Da", S & "nasa stavka na deljenoj paleti skinuta"
    ChkEq PalsStornirano("SVT-PS17B"), "Da", S & "nasa stavka na punoj paleti skinuta"
    ' co-tenant (druga zbirna) NETAKNUT
    Chk PalsStornirano("SVT-PS17X") <> "Da", S & "co-tenant stavka (druga zbirna) NETAKNUTA"
    ' PRAZNA paleta (samo nasa) -> STORNIRANA (kljucni fix: nije ostala 'puna')
    ChkEq PalStornirano("SVT-PAL17B"), "Da", S & "prazna paleta STORNIRANA (ne ostaje puna)"
    ' DELJENA paleta -> NE stornirana (co-tenant ostaje), reopen + gajbe 100->50
    Chk PalStornirano("SVT-PAL17A") <> "Da", S & "deljena paleta NIJE stornirana (co-tenant)"
    ChkEq PalStatus("SVT-PAL17A"), PAL_STATUS_OTVORENA, S & "deljena paleta reopen (ispod kapaciteta)"
    ChkEq PalBrGajbica("SVT-PAL17A"), 50, S & "deljena paleta gajbe 100->50 (skinuto nasih 50)"
End Sub

' ============================================================
' T18 - PONISTENJE zbirne (EKSTERNI kupac): interne otpremnice se storniraju, ali
' prijemnica (eksterna) ostaje NETAKNUTA (zbirna je poslednji interni dok).
' ============================================================
Private Sub T18_PonistenjeZbirnaEksterniNeDiraPrijemnicu()
    Const S As String = "T18 PONISTENJE zbirna (eksterni) ne dira prijemnicu: "

    SeedZbirna "SVT-Z18", "I", 100, 10, "SVT-EXT-KUPAC"
    SeedOtpremnica "SVT-O18", "SVT-Z18", "I", 100, 10
    SeedPrijemnica "SVT-P18", "SVT-Z18", "I", 100, 10

    Dim res As Object
    Set res = modStornoFlow.RunZbirnaCorrection("SVT-Z18", SV_MODE_PONISTENJE, True)
    Chk CBool(res("success")), S & "PONISTENJE uspeo"
    Chk Not ZbirnaPostoji("SVT-Z18"), S & "zbirna stornirana"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-O18", COL_OTP_ID), "", S & "otpremnica STORNIRANA (interno)"
    Chk LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-P18", COL_PRJ_ID) <> "", S & "prijemnica NETAKNUTA (eksterni kupac)"
End Sub

' ============================================================
' T19 - funkcionalna razlika: DUPLI zbirne ODVEZUJE otpremnice (prezivljavaju),
' PONISTENJE zbirne ih STORNIRA.
' ============================================================
Private Sub T19_DupliVsPonistenjeZbirnaOtpremnice()
    Const S As String = "T19 DUPLI odvezuje vs PONISTENJE stornira: "

    ' DUPLI: otpremnica prezivljava (aktivna, BrojZbirne prazno).
    SeedZbirna "SVT-Z19A", "I", 50, 5, HLAD_KUP
    SeedOtpremnica "SVT-O19A", "SVT-Z19A", "I", 50, 5
    modStornoFlow.RunZbirnaCorrection "SVT-Z19A", SV_MODE_DUPLI
    Chk LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-O19A", COL_OTP_ID) <> "", S & "DUPLI: otpremnica PREZIVLJAVA (aktivna)"
    ChkEq OtpBrojZbirne("SVT-O19A"), "", S & "DUPLI: otpremnica odvezana (BrojZbirne prazno)"

    ' PONISTENJE: otpremnica stornirana.
    SeedZbirna "SVT-Z19B", "I", 50, 5, HLAD_KUP
    SeedOtpremnica "SVT-O19B", "SVT-Z19B", "I", 50, 5
    modStornoFlow.RunZbirnaCorrection "SVT-Z19B", SV_MODE_PONISTENJE, True
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-O19B", COL_OTP_ID), "", S & "PONISTENJE: otpremnica STORNIRANA"
End Sub

' ============================================================
' T20 - PONISTENJE JEDNE otpremnice sa DELJENOM zbirnom NE obara zbirnu (sestre):
' storno te otpremnice + rekalk; sestra + prijemnica prezivljavaju.
' ============================================================
Private Sub T20_PonistenjeOtpremniceDeljenaNeObaraZbirnu()
    Const S As String = "T20 PONISTENJE otpremnice (deljena) ne obara zbirnu: "

    SeedZbirna "SVT-Z20", "I", 100, 10, HLAD_KUP
    SeedOtpremnica "SVT-OA20", "SVT-Z20", "I", 60, 6
    SeedOtpremnica "SVT-OB20", "SVT-Z20", "I", 40, 4
    SeedPrijemnica "SVT-P20", "SVT-Z20", "I", 100, 10

    Dim res As Object
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-OA20", SV_MODE_PONISTENJE, True)
    Chk CBool(res("success")), S & "PONISTENJE (forceConfirm) uspeo"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-OA20", COL_OTP_ID), "", S & "OA20 stornirana"
    Chk LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-OB20", COL_OTP_ID) <> "", S & "OB20 PREZIVLJAVA (deljena zbirna)"
    Chk ZbirnaPostoji("SVT-Z20"), S & "zbirna PREZIVLJAVA (deljena -> nije oborena)"
    Dim inv As Object: Set inv = modDokumentInvariant.ValidateZbirnaInvariant("SVT-Z20")
    ChkEqD CDbl(inv("kgZbrI")), 40, S & "zbirna rekalk na 40 (preostala OB20)"
    Chk LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-P20", COL_PRJ_ID) <> "", S & "prijemnica NETAKNUTA (zbirna ziva)"
End Sub

' ============================================================
' T21 - malina 1:1 (blok = okidac celog lanca; stanica=hladnjaca -> NIKAD dve
' otpremnice na zbirnoj / dva bloka na otpremnici): PONISTENJE JEDINE otpremnice
' ekskluzivno drzi zbirnu -> kaskada celog toka (zbirna+prijemnica+palete),
' blok oslobodjen (aktivan). Pokriva sole-owner PONISTENJE granu za otpremnicu.
' ============================================================
Private Sub T21_PonistenjeOtpremniceJedinaKaskadaCeoTok()
    Const S As String = "T21 PONISTENJE jedine otpremnice (malina 1:1) kaskada: "

    SeedZbirna "SVT-Z21", "I", 100, 10, HLAD_KUP
    SeedOtpremnica "SVT-O21", "SVT-Z21", "I", 100, 10
    SeedOtkupBlok "SVT-BLK21", "SVT-O21-ID-I", "SVT-Z21"
    SeedPrijemnica "SVT-P21", "SVT-Z21", "I", 100, 10
    SeedPaleta "SVT-PAL21", 50, 100, 10, 50, PAL_STATUS_ZATVORENA
    SeedPaletaStavka "SVT-PS21", "SVT-P21", "SVT-Z21", "I", 100, 10, "SVT-PAL21", 50

    Dim res As Object
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-O21", SV_MODE_PONISTENJE, True)
    Chk CBool(res("success")), S & "PONISTENJE (forceConfirm) uspeo"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-O21", COL_OTP_ID), "", S & "otpremnica stornirana"
    Chk Not ZbirnaPostoji("SVT-Z21"), S & "zbirna stornirana (jedina otpremnica -> ceo tok)"
    ChkEq LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, "SVT-P21", COL_PRJ_ID), "", S & "prijemnica stornirana (kaskada)"
    ChkEq PalsStornirano("SVT-PS21"), "Da", S & "paletna stavka skinuta"
    ChkEq PalStornirano("SVT-PAL21"), "Da", S & "prazna paleta STORNIRANA (ne ostaje puna)"
    ' blok OSLOBODJEN ali AKTIVAN (realna kupovina, za reveze).
    ChkEq OtkOtpremnicaID("SVT-BLK21"), "", S & "blok oslobodjen (OtpremnicaID prazan)"
    Chk LookupActiveID(TBL_OTKUP, COL_OTK_BR_DOK, "SVT-BLK21", COL_OTK_ID) <> "", S & "blok i dalje aktivan (nije storniran)"
End Sub

' ============================================================
' T22 - SIMPLE storno otpremnice (bez nizvodnog toka) NE ostavlja aktivnu zbirnu
' 0/0: jedina otpremnica, bez prijemnice/paleta -> smart trigger ide simple path;
' prazna zbirna se STORNIRA (nulling fix i u simple putanji, dosledno DUPLI/
' PONISTENJE grani). Regres za review tacku 1.
' ============================================================
Private Sub T22_SimpleStornoOtpremniceNeOstavljaZbirnu00()
    Const S As String = "T22 simple storno otpremnice ne ostavlja zbirnu 0/0: "

    ' Eksterni kupac + MALINA_MODE=NO (suite) -> nema auto/malina kaskade zbirne;
    ' zbirnu mora oboriti sam framework (RecalcOrStornoEmptyZbirna_TX).
    SeedZbirna "SVT-Z22", "I", 100, 10, "SVT-EXT-KUPAC"
    SeedOtpremnica "SVT-O22", "SVT-Z22", "I", 100, 10

    Chk Not modStornoFlow.CorrectionNeedsDialog(FLOW_DOC_OTPREMNICA, "SVT-O22"), _
        S & "CorrectionNeedsDialog = False (nema nizvodnog toka -> simple path)"

    Dim res As Object
    Set res = modStornoFlow.RunSimpleStornoOtpremnica("SVT-O22")
    Chk CBool(res("success")), S & "RunSimpleStornoOtpremnica uspeo"
    ChkEq LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, "SVT-O22", COL_OTP_ID), "", S & "otpremnica stornirana"
    ' KLJUC: zbirna NIJE ostavljena aktivna 0/0 -> stornirana (nema vise otpremnica).
    Chk Not ZbirnaPostoji("SVT-Z22"), S & "prazna zbirna STORNIRANA (ne aktivna 0/0)"
End Sub

' ============================================================
' SEED HELPERS (upis po IMENU kolone -> otporno na redosled)
' ============================================================

' Zbirna: jedan red po klasi. BrojZbirne + Klasa + KG + AMB (+ ZbirnaID, Datum).
' kupac (opciono) = KupacID; = HLAD_KUP -> interni hladnjaca-tok, inace eksterni.
' ============================================================
' T30 - staging red BEZ BankaImportID mora da odbije ceo storno.
'
' Bez PK-a bi prazan kljuc usao u skup izvoda, a BimIdFromNapomena za svaki RUCNO
' unet red vraca "" -> storno bi zahvatio tudj novac. Ovo je regresioni cuvar za
' taj scenario: proverava se i da rucni novac DRUGOG dokumenta ostane netaknut.
' ============================================================
Private Sub T30_StagingBezPKOdbijaStorno()
    Const S As String = "T30 staging bez PK: "

    SeedBimStavka "SVT-BIM-6A", "SVT-IZV-6", "SVT-RAC-6", "SVT-PARTNER-6", 900, 0
    SeedBimStavka "", "SVT-IZV-6", "SVT-RAC-6", "SVT-PARTNER-6", 100, 0      ' pokvaren red
    SeedNovacBim "SVT-NOV-6A", "SVT-IZV-6", "SVT-BIM-6A", "SVT-P6", 900, 0
    SeedNovacSplit "SVT-NOV-RUC", "SVT-RUCNI-BROJ", "SVT-P9", 1234, 0        ' tudj rucni novac

    Dim info As String
    Chk Not StornoIzvod_TX("SVT-IZV-6", "SVT-RAC-6", IZVOD_STORNO_REMAP, info), S & "storno odbijen"
    ChkEq NovStornirano("SVT-NOV-6A"), "", S & "novac izvoda netaknut (nema delimicnog storna)"
    ChkEq NovStornirano("SVT-NOV-RUC"), "", S & "TUDJ rucni novac netaknut"
    ChkEq BimObradjeno("SVT-BIM-6A"), "Da", S & "staging netaknut"
End Sub

' ============================================================
' T31 - prazan racun je dvosmislen identitet: javna TX funkcija ga mora odbiti
' sama, ne oslanjajuci se na to da je forma prvo pozvala resolver.
' ============================================================
Private Sub T31_IzvodBezRacunaOdbijen()
    Const S As String = "T31 izvod bez racuna: "

    SeedBimStavka "SVT-BIM-7A", "SVT-IZV-7", "SVT-RAC-7", "SVT-PARTNER-7", 500, 0
    SeedNovacBim "SVT-NOV-7A", "SVT-IZV-7", "SVT-BIM-7A", "SVT-P7", 500, 0

    Dim info As String
    Chk Not StornoIzvod_TX("SVT-IZV-7", "", IZVOD_STORNO_REMAP, info), S & "prazan racun -> odbijeno"
    ChkEq NovStornirano("SVT-NOV-7A"), "", S & "novac netaknut"
    ChkEq BimObradjeno("SVT-BIM-7A"), "Da", S & "staging netaknut"
End Sub

' ============================================================
' T32 - avans split NASLEDJUJE BIM marker (BuildAvansSplitNapomena) pa se hvata
' direktno, i kad ima drugi BrojDokumenta. Markerless red BEZ partnera se NE
' pogadja - to je bila preteska heuristika.
' ============================================================
Private Sub T32_SamoMarkerOdredjujePripadnost()
    Const S As String = "T32 pripadnost samo po markeru: "

    SeedBimStavka "SVT-BIM-8A", "SVT-IZV-8", "SVT-RAC-8", "SVT-PARTNER-8", 0, 3000
    SeedNovacBim "SVT-NOV-8A", "SVT-IZV-8", "SVT-BIM-8A", "SVT-P8", 0, 1800
    ' Red sa markerom pripada izvodu i kad ima DRUGI BrojDokumenta.
    SeedNovacBim "SVT-NOV-8B", "SVT-DRUGI-BROJ-8", "SVT-BIM-8A", "SVT-P8", 0, 1200
    ' Rucni red sa ISTIM brojem I ISTIM partnerom, ali BEZ markera -> ne sme se
    ' pokupiti (nema heuristike po broju/partneru).
    SeedNovacSplit "SVT-NOV-8C", "SVT-IZV-8", "SVT-P8", 0, 500

    Dim info As String
    Chk StornoIzvod_TX("SVT-IZV-8", "SVT-RAC-8", IZVOD_STORNO_REMAP, info), S & "storno uspeo"
    ChkEq NovStornirano("SVT-NOV-8A"), "Da", S & "red sa markerom storniran"
    ChkEq NovStornirano("SVT-NOV-8B"), "Da", S & "marker vazi i uz drugi BrojDokumenta"
    ChkEq NovStornirano("SVT-NOV-8C"), "", S & "rucni red (isti broj I partner, bez markera) NETAKNUT"
End Sub

' ============================================================
' T36 - rucni red koji DELI BROJ sa izvodom mora ostati stornirljiv pojedinacno.
'
' Zastita "ne stornira se izvod parcijalno" mora da gleda MARKER konkretnog reda,
' ne BrojDokumenta. Provera po broju bi ovaj rucni red ostavila netaknutim pri
' stornu izvoda (dobro), ali i trajno nestornirljivim normalnom putanjom (lose).
' ============================================================
Private Sub T36_RucniRedSaIstimBrojemOstajeStornirljiv()
    Const S As String = "T36 rucni red deli broj sa izvodom: "

    SeedBimStavka "SVT-BIM-12A", "SVT-IZV-12", "SVT-RAC-12", "SVT-PARTNER-12", 2000, 0
    SeedNovacBim "SVT-NOV-12A", "SVT-IZV-12", "SVT-BIM-12A", "SVT-P12", 2000, 0
    SeedNovacSplit "SVT-NOV-12R", "SVT-IZV-12", "SVT-P12", 700, 0        ' rucni, isti broj i partner

    ' Izvodni red se NE sme stornirati pojedinacno.
    Dim razlogBim As String
    Chk Len(ResolveNovacForStorno("SVT-NOV-12A", razlogBim)) = 0, S & "izvodni red odbijen pojedinacno"
    Chk Len(razlogBim) > 0, S & "razlog odbijanja objasnjen"

    ' Rucni red MORA proci - i kroz razresavanje i kroz stvarni storno.
    Dim razlogRuc As String
    ChkEq ResolveNovacForStorno("SVT-NOV-12R", razlogRuc), "SVT-NOV-12R", S & "rucni red prihvacen"
    ChkEq razlogRuc, "", S & "nema razloga za odbijanje rucnog reda"
    Chk StornoNovac_TX("SVT-NOV-12R"), S & "StornoNovac_TX stornirao rucni red"
    ChkEq NovStornirano("SVT-NOV-12R"), "Da", S & "rucni red stvarno storniran"
    ChkEq NovStornirano("SVT-NOV-12A"), "", S & "izvodni red i dalje netaknut"
End Sub

' ============================================================
' T35 - rekonsilijacija iznosa: zbir aktivnog novca po stavci mora biti jednak
' iznosu stavke izvoda. Nesklad znaci da je nesto vec dirnuto -> odbij ceo storno.
' ============================================================
Private Sub T35_NeslaganjeIznosaBlokira()
    Const S As String = "T35 neslaganje iznosa: "

    SeedBimStavka "SVT-BIM-11A", "SVT-IZV-11", "SVT-RAC-11", "SVT-PARTNER-11", 1000, 0
    SeedNovacBim "SVT-NOV-11A", "SVT-IZV-11", "SVT-BIM-11A", "SVT-P11", 600, 0   ' fali 400

    Chk Len(GetIzvodStornoBlokade("SVT-IZV-11", "SVT-RAC-11")) > 0, S & "preflight prijavljuje razliku"

    Dim info As String
    Chk Not StornoIzvod_TX("SVT-IZV-11", "SVT-RAC-11", IZVOD_STORNO_REMAP, info), S & "storno odbijen"
    ChkEq NovStornirano("SVT-NOV-11A"), "", S & "novac netaknut"
    ChkEq BimObradjeno("SVT-BIM-11A"), "Da", S & "staging netaknut"
End Sub

' ============================================================
' T33 - stavka oznacena kao obradjena, a nema nijedan novac red: vracanje u
' "za obradu" bi dozvolilo ponovno mapiranje i dvostruko knjizenje -> blokada.
' ============================================================
Private Sub T33_ObradjenaStavkaBezNovcaBlokira()
    Const S As String = "T33 obradjena stavka bez novca: "

    SeedBimStavka "SVT-BIM-9A", "SVT-IZV-9", "SVT-RAC-9", "SVT-PARTNER-9", 700, 0

    Chk Len(GetIzvodStornoBlokade("SVT-IZV-9", "SVT-RAC-9")) > 0, S & "preflight objasnjava blokadu"

    Dim info As String
    Chk Not StornoIzvod_TX("SVT-IZV-9", "SVT-RAC-9", IZVOD_STORNO_REMAP, info), S & "storno odbijen"
    ChkEq BimObradjeno("SVT-BIM-9A"), "Da", S & "staging NIJE tiho vracen u 'za obradu'"
End Sub

' ============================================================
' T34 - broj izvoda sme da sadrzi kosu crtu (npr. "42/2026"); ne sme se raseci
' kao "broj/racun".
' ============================================================
Private Sub T34_BrojIzvodaSaKosomCrtom()
    Const S As String = "T34 broj izvoda sa kosom crtom: "

    SeedBimStavka "SVT-BIM-10A", "SVT-IZV/2026", "SVT-RAC-10", "SVT-PARTNER-10", 100, 0

    Dim br As String, rac As String, razlog As String
    Chk ResolveIzvodZaStorno("SVT-IZV/2026", br, rac, razlog), S & "ceo unos prihvacen kao broj"
    ChkEq br, "SVT-IZV/2026", S & "broj nije rasecen na kosoj crti"
    ChkEq rac, "SVT-RAC-10", S & "racun izveden iz staginga"
End Sub

' ============================================================
' T28 - OM avans broji OBA kanala placanja.
'
' Posle razdvajanja kanala (Tip nosi kes vs virman) redovi uvezeni iz izvoda nose
' VIRMAN tip, a redovi uvezeni PRE toga i dalje nose KES tip. Svaki citalac OM
' avansa mora racunati oba - inace iznosi tiho padnu (bankovni novac ispadne iz
' salda, ili legacy zapisi ispadnu posle migracije). Ovo je regresioni cuvar koji
' zamenjuje rucno "uporedi saldo pre i posle importa grane".
'
' Citaoci: modNovac.GetOMAvansSaldo (ovde), modIzvestaj x2 (T29),
' modStammdatenSync.ExportSaldoOM - Private, pa se pokriva ugovorom klasifikatora.
' ============================================================
Private Sub T28_OMAvansBrojiObaKanala()
    Const S As String = "T28 OM avans oba kanala: "

    ' Ugovor klasifikatora - od njega zavise SVI citaoci.
    Chk IsFirmaOtkupacAvansTip(NOV_KES_FIRMA_OTKUPAC), S & "KES Firma-Otkupac je OM avans (legacy zapis)"
    Chk IsFirmaOtkupacAvansTip(NOV_VIRMAN_FIRMA_OTKUPAC), S & "VIRMAN Firma-Otkupac je OM avans (iz izvoda)"
    Chk Not IsFirmaOtkupacAvansTip(NOV_KES_OTKUPAC_KOOP), S & "isplata kooperantu NIJE OM avans"
    Chk Not IsFirmaOtkupacAvansTip(NOV_VIRMAN_FIRMA_KOOP), S & "virman kooperantu NIJE OM avans"
    Chk IsKesNovacTip(NOV_KES_OTKUPAC_KOOP), S & "KesOtkupacKoop je kes"
    Chk Not IsKesNovacTip(NOV_VIRMAN_FIRMA_OTKUPAC), S & "VirmanFirmaOtkupac NIJE kes"

    ' Saldo = 1000 (kes, legacy) + 500 (virman, iz izvoda) - 300 (podeljeno kooperantu)
    SeedNovacOM "SVT-NOV-OM1", "SVT-OM-1", "", NOV_KES_FIRMA_OTKUPAC, 1000
    SeedNovacOM "SVT-NOV-OM2", "SVT-OM-1", "", NOV_VIRMAN_FIRMA_OTKUPAC, 500
    SeedNovacOM "SVT-NOV-OM3", "SVT-OM-1", "SVT-KOOP-OM", NOV_KES_OTKUPAC_KOOP, 300

    ChkEqD GetOMAvansSaldo("SVT-OM-1"), 1200, S & "saldo = 1000 kes + 500 virman - 300 podeljeno"
    ' Da pad bude citljiv: 700 = virman kanal ispao, 900 = kes (legacy) kanal ispao.
    Chk GetOMAvansSaldo("SVT-OM-1") <> 700, S & "virman kanal NIJE ispao iz salda"
    Chk GetOMAvansSaldo("SVT-OM-1") <> 200, S & "kes (legacy) kanal NIJE ispao iz salda"
End Sub

' ============================================================
' T29 - isti kanal-invariant kroz IZVESTAJE (ReportSaldoOM + ReportIsplata).
' Lokalni EH: neocekivana greska u report sloju daje FAIL, ne obara ceo suite.
' ============================================================
Private Sub T29_OMAvansUIzvestajimaObaKanala()
    Const S As String = "T29 OM avans u izvestajima: "
    On Error GoTo EH

    SeedNovacOM "SVT-NOV-OM4", "SVT-OM-2", "", NOV_KES_FIRMA_OTKUPAC, 800
    SeedNovacOM "SVT-NOV-OM5", "SVT-OM-2", "", NOV_VIRMAN_FIRMA_OTKUPAC, 200
    SeedNovacOM "SVT-NOV-OM6", "SVT-OM-2", "SVT-KOOP-OM", NOV_KES_OTKUPAC_KOOP, 250

    Dim od As Date: od = Date - 1
    Dim doD As Date: doD = Date + 1

    ' ReportSaldoOM: red "OM AVANS (nerasporedjen)", kolona 4 = primljeno - podeljeno.
    Dim rs As Variant: rs = ReportSaldoOM("SVT-OM-2", od, doD)
    ChkEqD ReportCellByLabel(rs, "OM AVANS (nerasporedjen)", 4), 750, _
        S & "ReportSaldoOM = 800 kes + 200 virman - 250 podeljeno"

    ' ReportIsplata: kontrolni redovi (primljeno = oba kanala; podeljeno = kes koop).
    Dim ri As Variant: ri = ReportIsplata("OM", "SVT-OM-2", od, doD)
    ChkEqD ReportCellByLabel(ri, "OM Avans (primljeno)", 5), 1000, _
        S & "ReportIsplata primljeno = 800 kes + 200 virman"
    ChkEqD ReportCellByLabel(ri, "OM Avans (podeljeno)", 5), 250, _
        S & "ReportIsplata podeljeno = 250"
    Exit Sub
EH:
    Chk False, S & "neocekivana greska u report sloju: " & Err.description
End Sub

' ============================================================
' T23 - storno izvoda, ishod REMAP (PDF ispravan, mapiranje pogresno):
' novac pada, staging stavke se vracaju u "za obradu" i NISU ugasene.
' ============================================================
Private Sub T23_StornoIzvodaRemapVracaStavkeUObradu()
    Const S As String = "T23 storno izvoda REMAP: "

    SeedBimStavka "SVT-BIM-1A", "SVT-IZV-1", "SVT-RAC-1", "SVT-PARTNER", 5000, 0
    SeedBimStavka "SVT-BIM-1B", "SVT-IZV-1", "SVT-RAC-1", "SVT-PARTNER", 3000, 0
    SeedNovacBim "SVT-NOV-1A", "SVT-IZV-1", "SVT-BIM-1A", "SVT-P1", 5000, 0
    SeedNovacBim "SVT-NOV-1B", "SVT-IZV-1", "SVT-BIM-1B", "SVT-P1", 3000, 0

    Dim info As String
    Chk StornoIzvod_TX("SVT-IZV-1", "SVT-RAC-1", IZVOD_STORNO_REMAP, info), S & "StornoIzvod_TX uspeo"

    ChkEq NovStornirano("SVT-NOV-1A"), "Da", S & "novac 1A storniran"
    ChkEq NovStornirano("SVT-NOV-1B"), "Da", S & "novac 1B storniran"
    ChkEq BimObradjeno("SVT-BIM-1A"), "", S & "staging 1A vracen u 'za obradu'"
    ChkEq BimObradjeno("SVT-BIM-1B"), "", S & "staging 1B vracen u 'za obradu'"
    ChkEq BimStornirano("SVT-BIM-1A"), "", S & "staging 1A NIJE ugasen (PDF ispravan)"
End Sub

' ============================================================
' T24 - storno izvoda, ishod REIMPORT (PDF los): staging se gasi, a dedupe ga vise
' ne vidi -> isti PDF se moze uvesti PONOVO (to je cela poenta ishoda).
' ============================================================
Private Sub T24_StornoIzvodaReimportOslobadjaPonovniUvoz()
    Const S As String = "T24 storno izvoda REIMPORT: "

    SeedBimStavka "SVT-BIM-2A", "SVT-IZV-2", "SVT-RAC-2", "SVT-PARTNER-2", 7000, 0
    SeedNovacBim "SVT-NOV-2A", "SVT-IZV-2", "SVT-BIM-2A", "SVT-P2", 7000, 0

    ' Pre storna: ponovni uvoz iste stavke bi bio odbijen kao duplikat.
    Chk IsDuplicateBankaImport("SVT-IZV-2", Date, 7000, 0, "SVT-PARTNER-2", "", "SVT-RAC-2"), _
        S & "pre storna dedupe prepoznaje duplikat"

    Dim info As String
    Chk StornoIzvod_TX("SVT-IZV-2", "SVT-RAC-2", IZVOD_STORNO_REIMPORT, info), S & "StornoIzvod_TX uspeo"

    ChkEq NovStornirano("SVT-NOV-2A"), "Da", S & "novac storniran"
    ChkEq BimStornirano("SVT-BIM-2A"), "Da", S & "staging ugasen"
    Chk Not IsDuplicateBankaImport("SVT-IZV-2", Date, 7000, 0, "SVT-PARTNER-2", "", "SVT-RAC-2"), _
        S & "posle storna ponovni uvoz istog PDF-a NIJE duplikat"
End Sub

' ============================================================
' T25 - KANONSKI lineage test kroz STVARNI tok, ne kroz seed ocekivanog rezultata.
'
' Pusta pravi ApplyAvansToOtkup_TX da napravi split (umanji original, napravi novi
' red), pa proverava da je split NASLEDIO BIM marker i da ga storno izvoda obara
' zajedno sa originalom. Usput dokazuje i rekonsilijaciju: original + split moraju
' i dalje davati tacan iznos stavke izvoda (2500 + 1500 = 4000), inace bi preflight
' odbio storno.
' ============================================================
Private Sub T25_AvansSplitNasledjujeMarkerIPada()
    Const S As String = "T25 avans split (stvarni tok): "

    SeedBimStavka "SVT-BIM-3A", "SVT-IZV-3", "SVT-RAC-3", "SVT-PARTNER-3", 0, 4000
    SeedNovacBim "SVT-NOV-3A", "SVT-IZV-3", "SVT-BIM-3A", "SVT-P3", 0, 4000, "", "SVT-K3"
    SeedOtkupZaAvans "SVT-OTK-3", "SVT-K3", 100, 15      ' vrednost 1500 -> split

    ChkEq AktivnihSaMarkerom("SVT-BIM-3A"), 1, S & "pre raspodele jedan red izvoda"

    Dim applied As Double
    Chk ApplyAvansToOtkup_TX("SVT-K3", "SVT-OTK-3", applied), S & "ApplyAvansToOtkup_TX uspeo"
    ChkEqD applied, 1500, S & "primenjeno tacno 1500 (vrednost bloka)"
    ChkEq AktivnihSaMarkerom("SVT-BIM-3A"), 2, S & "split NASLEDIO marker (2 reda pod istim BIM-om)"

    Dim info As String
    Chk StornoIzvod_TX("SVT-IZV-3", "SVT-RAC-3", IZVOD_STORNO_REMAP, info), S & "StornoIzvod_TX uspeo"
    ChkEq NovStornirano("SVT-NOV-3A"), "Da", S & "originalni red storniran"
    ChkEq AktivnihSaMarkerom("SVT-BIM-3A"), 0, S & "nijedan red izvoda nije ostao aktivan (ni split)"
End Sub

' ============================================================
' T26 - isti broj izvoda na dva racuna (dve banke): razresavanje mora TRAZITI
' racun, a ne tiho uzeti jedan.
' ============================================================
Private Sub T26_IzvodNaViseRacunaTraziRacun()
    Const S As String = "T26 izvod na vise racuna: "

    SeedBimStavka "SVT-BIM-4A", "SVT-IZV-4", "SVT-RAC-A", "SVT-PARTNER-4", 1000, 0
    SeedBimStavka "SVT-BIM-4B", "SVT-IZV-4", "SVT-RAC-B", "SVT-PARTNER-4", 2000, 0

    Dim br As String, rac As String, razlog As String
    Chk Not ResolveIzvodZaStorno("SVT-IZV-4", br, rac, razlog), S & "sam broj je odbijen (dvosmislen)"
    Chk Len(razlog) > 0, S & "razlog objasnjen operateru"

    Chk ResolveIzvodZaStorno("SVT-IZV-4/SVT-RAC-B", br, rac, razlog), S & "broj/racun je prihvacen"
    ChkEq rac, "SVT-RAC-B", S & "izabran tacan racun"
End Sub

' ============================================================
' T27 - storno izvoda osvezava otkup: isplata iz izvoda je pokrivala blok; posle
' storna blok vise NIJE isplacen (inace ostaje nevidljiv u listi za isplatu).
' ============================================================
Private Sub T27_StornoIzvodaOsvezavaOtkup()
    Const S As String = "T27 storno izvoda osvezava otkup: "

    SeedOtkupPlacen "SVT-OTK-5", 100, 10          ' vrednost 1000, oznacen kao isplacen
    SeedBimStavka "SVT-BIM-5A", "SVT-IZV-5", "SVT-RAC-5", "SVT-PARTNER-5", 0, 1000
    SeedNovacBim "SVT-NOV-5A", "SVT-IZV-5", "SVT-BIM-5A", "SVT-P5", 0, 1000, "SVT-OTK-5"

    ChkEq OtkIsplaceno("SVT-OTK-5"), STATUS_ISPLACENO, S & "pre storna blok je isplacen"

    Dim info As String
    Chk StornoIzvod_TX("SVT-IZV-5", "SVT-RAC-5", IZVOD_STORNO_REMAP, info), S & "StornoIzvod_TX uspeo"

    ChkEq NovStornirano("SVT-NOV-5A"), "Da", S & "isplata stornirana"
    ChkEq OtkIsplaceno("SVT-OTK-5"), "", S & "blok vise nije isplacen"
    ChkEq OtkDatumIsplate("SVT-OTK-5"), "", S & "datum isplate ocisceni"
End Sub

Private Sub SeedZbirna(ByVal broj As String, ByVal klasa As String, _
                       ByVal kg As Double, ByVal amb As Long, _
                       Optional ByVal kupac As String = "")
    SvAppend TBL_ZBIRNA, _
        Array(COL_ZBR_ID, COL_ZBR_DATUM, COL_ZBR_BROJ, COL_ZBR_KUPAC, COL_ZBR_KOLICINA, _
              COL_ZBR_TIP_AMB, COL_ZBR_KOL_AMB, COL_ZBR_VRSTA, COL_ZBR_SORTA, COL_ZBR_KLASA), _
        Array(broj & "-ID-" & klasa, Date, broj, kupac, kg, "SVT-A", amb, "SVT-VOCE", "SVT-SORTA", klasa)
End Sub

Private Sub SeedOtpremnica(ByVal broj As String, ByVal brojZbirne As String, _
                           ByVal klasa As String, ByVal kg As Double, ByVal amb As Long)
    SvAppend TBL_OTPREMNICA, _
        Array(COL_OTP_ID, COL_OTP_DATUM, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE, _
              COL_OTP_KOLICINA, COL_OTP_TIP_AMB, COL_OTP_KOL_AMB, COL_OTP_VRSTA, _
              COL_OTP_SORTA, COL_OTP_KLASA), _
        Array(broj & "-ID-" & klasa, Date, broj, brojZbirne, kg, "SVT-A", amb, _
              "SVT-VOCE", "SVT-SORTA", klasa)
End Sub

Private Sub SeedPrijemnica(ByVal broj As String, ByVal brojZbirne As String, _
                           ByVal klasa As String, ByVal kg As Double, ByVal amb As Long)
    SvAppend TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_DATUM, COL_PRJ_BROJ, COL_PRJ_BROJ_ZBIRNE, _
              COL_PRJ_KOLICINA, COL_PRJ_TIP_AMB, COL_PRJ_KOL_AMB, COL_PRJ_KLASA), _
        Array(broj & "-ID-" & klasa, Date, broj, brojZbirne, kg, "SVT-A", amb, klasa)
End Sub

' paletaID/gajbe opciono: postave se kad test ide kroz paletni motor (DetachOsiroce
' u PONISTENJE kaskadi). Kad su prazni (ISPRAVKA testovi) stavka je "bez palete".
Private Sub SeedPaletaStavka(ByVal stavkaID As String, ByVal brojPrij As String, _
                             ByVal brojZbirne As String, ByVal klasa As String, _
                             ByVal kg As Double, ByVal amb As Long, _
                             Optional ByVal paletaID As String = "", _
                             Optional ByVal gajbe As Long = 0)
    SvAppend TBL_PALETA_STAVKA, _
        Array(COL_PALS_ID, COL_PALS_PALETA_ID, COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, _
              COL_PALS_KLASA, COL_PALS_BR_GAJBICA, COL_PALS_NETO, COL_PALS_AMBALAZA), _
        Array(stavkaID, paletaID, brojPrij, brojZbirne, klasa, gajbe, kg, amb)
End Sub

' Minimalna realna paleta (header) da DetachOsirocenePaletaStavke_TX radi end-to-end:
' skida gajbe/neto/amb, reopen ispod kapaciteta, storno kad ostane prazna.
Private Sub SeedPaleta(ByVal palID As String, ByVal gajbe As Long, ByVal neto As Double, _
                       ByVal amb As Double, ByVal kapacitet As Long, ByVal status As String)
    SvAppend TBL_PALETA, _
        Array(COL_PAL_ID, COL_PAL_BR_GAJBICA, COL_PAL_NETO, COL_PAL_AMBALAZA, COL_PAL_BRUTO, _
              COL_PAL_KAPACITET, COL_PAL_STATUS, COL_PAL_PRERADJENO, COL_PAL_KLASA), _
        Array(palID, gajbe, neto, amb, neto + amb, kapacitet, status, "Ne", "I")
End Sub

' Otkupni blok ("list") vezan za otpremnicu (OtpremnicaID) + denorm. BrojZbirne.
Private Sub SeedOtkupBlok(ByVal blkID As String, ByVal otpID As String, ByVal brojZbirne As String)
    SvAppend TBL_OTKUP, _
        Array(COL_OTK_ID, COL_OTK_OTPREMNICA_ID, COL_OTK_BROJ_ZBIRNE, COL_OTK_BR_DOK), _
        Array(blkID, otpID, brojZbirne, blkID)
End Sub

' Revers = dvojni upis (kooperant Ulaz + stanica Izlaz), oba dele DokumentID+Tip.
Private Sub SeedRevers(ByVal brDok As String, ByVal dokTip As String, _
                       ByVal koopID As String, ByVal stanicaID As String, _
                       ByVal tipAmb As String, ByVal kol As Long)
    SeedAmb brDok & "-K", tipAmb, kol, "Ulaz", koopID, "Kooperant", brDok, dokTip
    SeedAmb brDok & "-S", tipAmb, kol, "Izlaz", stanicaID, "Stanica", brDok, dokTip
End Sub

Private Sub SeedAmb(ByVal id As String, ByVal tip As String, ByVal kol As Long, _
                    ByVal smer As String, ByVal entID As String, ByVal entTip As String, _
                    ByVal dokID As String, ByVal dokTip As String)
    SvAppend TBL_AMBALAZA, _
        Array(COL_AMB_ID, COL_AMB_DATUM, COL_AMB_TIP, COL_AMB_KOLICINA, COL_AMB_SMER, _
              COL_AMB_ENTITET, COL_AMB_ENTITET_TIP, COL_AMB_DOK_ID, COL_AMB_DOK_TIP), _
        Array(id, Date, tip, kol, smer, entID, entTip, dokID, dokTip)
End Sub

' --- Banka izvod (staging + novac) ---

' Stavka uvezenog izvoda; Obradjeno="Da" = vec mapirana (to je stanje iz kog se
' storno izvoda i poziva).
Private Sub SeedBimStavka(ByVal bimID As String, ByVal brojIzvoda As String, _
                          ByVal racun As String, ByVal partner As String, _
                          ByVal uplata As Double, ByVal isplata As Double)
    SvAppend TBL_BANKA_IMPORT, _
        Array(COL_BIM_ID, COL_BIM_BROJ_DOKUMENTA, COL_BIM_BROJ_RACUNA, COL_BIM_DATUM_TRANSAKCIJE, _
              COL_BIM_PARTNER, COL_BIM_UPLATA, COL_BIM_ISPLATA, COL_BIM_OBRADJENO), _
        Array(bimID, brojIzvoda, racun, Date, partner, uplata, isplata, "Da")
End Sub

' Novac red nastao mapiranjem izvoda: Napomena nosi BIM marker (jedina veza
' novac -> izvod), BrojDokumenta = broj izvoda.
Private Sub SeedNovacBim(ByVal novID As String, ByVal brojIzvoda As String, ByVal bimID As String, _
                         ByVal partnerID As String, ByVal uplata As Double, ByVal isplata As Double, _
                         Optional ByVal otkupID As String = "", Optional ByVal koopID As String = "")
    SvAppend TBL_NOVAC, _
        Array(COL_NOV_ID, COL_NOV_BROJ_DOK, COL_NOV_DATUM, COL_NOV_PARTNER_ID, COL_NOV_KOOP_ID, _
              COL_NOV_TIP, COL_NOV_UPLATA, COL_NOV_ISPLATA, COL_NOV_NAPOMENA, COL_NOV_OTKUP_ID), _
        Array(novID, brojIzvoda, Date, partnerID, koopID, NOV_VIRMAN_AVANS_KOOP, uplata, isplata, _
              NOV_NAPOMENA_BIM_PREFIX & bimID & "; Ref:SVT", otkupID)
End Sub

' Otkup blok spreman za raspodelu avansa (kooperant + vrednost, NIJE isplacen).
Private Sub SeedOtkupZaAvans(ByVal otkID As String, ByVal koopID As String, _
                             ByVal kolicina As Double, ByVal cena As Double)
    SvAppend TBL_OTKUP, _
        Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KOOPERANT, COL_OTK_KOLICINA, COL_OTK_CENA), _
        Array(otkID, otkID, koopID, kolicina, cena)
End Sub

' Avans split naslednik: isti broj i partner kao original, ali BEZ BIM markera
' (tako ga ApplyAvansToFaktura/ApplyAvansToOtkup stvarno prave).
Private Sub SeedNovacSplit(ByVal novID As String, ByVal brojIzvoda As String, _
                           ByVal partnerID As String, ByVal uplata As Double, ByVal isplata As Double)
    SvAppend TBL_NOVAC, _
        Array(COL_NOV_ID, COL_NOV_BROJ_DOK, COL_NOV_DATUM, COL_NOV_PARTNER_ID, COL_NOV_TIP, _
              COL_NOV_UPLATA, COL_NOV_ISPLATA, COL_NOV_NAPOMENA), _
        Array(novID, brojIzvoda, Date, partnerID, NOV_VIRMAN_AVANS_KOOP, uplata, isplata, _
              "Avans raspodela")
End Sub

' Novac vezan za otkupno mesto: OM avans (koopID prazan) ili isplata kooperantu iz
' tog avansa (koopID popunjen). Datum = danas -> ulazi u report period T29.
Private Sub SeedNovacOM(ByVal novID As String, ByVal omID As String, ByVal koopID As String, _
                        ByVal tip As String, ByVal isplata As Double)
    SvAppend TBL_NOVAC, _
        Array(COL_NOV_ID, COL_NOV_BROJ_DOK, COL_NOV_DATUM, COL_NOV_OM_ID, COL_NOV_KOOP_ID, _
              COL_NOV_ENTITET_TIP, COL_NOV_TIP, COL_NOV_UPLATA, COL_NOV_ISPLATA), _
        Array(novID, novID, Date, omID, koopID, "OM", tip, 0, isplata)
End Sub

' Otkup blok sa kolicinom/cenom, unapred oznacen kao isplacen (T27 proverava da
' storno izvoda to ponisti).
Private Sub SeedOtkupPlacen(ByVal otkID As String, ByVal kolicina As Double, ByVal cena As Double)
    SvAppend TBL_OTKUP, _
        Array(COL_OTK_ID, COL_OTK_BR_DOK, COL_OTK_KOLICINA, COL_OTK_CENA, _
              COL_OTK_ISPLACENO, COL_OTK_DATUM_ISPLATE), _
        Array(otkID, otkID, kolicina, cena, STATUS_ISPLACENO, Date)
End Sub

' Append red po IMENU kolone (preskace kolone kojih nema). Vraca nista; raise
' ako tabela ne postoji.
Private Sub SvAppend(ByVal tblName As String, ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then Err.Raise vbObjectError + 2800, "modTestStorno.SvAppend", "Nema tabele: " & tblName
    Dim nr As ListRow
    Set nr = lo.ListRows.Add
    Dim i As Long, ci As Long
    For i = LBound(cols) To UBound(cols)
        ci = GetColumnIndex(tblName, CStr(cols(i)))
        If ci > 0 Then nr.Range.cells(1, ci).value = vals(i)
    Next i
End Sub

' ============================================================
' READ HELPERS
' ============================================================

Private Function OtpBrojZbirne(ByVal otpBroj As String) As String
    OtpBrojZbirne = NzTx(LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, otpBroj, COL_OTP_BROJ_ZBIRNE))
End Function

Private Function PrjBrojZbirne(ByVal prjBroj As String) As String
    PrjBrojZbirne = NzTx(LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ, prjBroj, COL_PRJ_BROJ_ZBIRNE))
End Function

Private Function PalsBrojZbirne(ByVal stavkaID As String) As String
    PalsBrojZbirne = NzTx(LookupValue(TBL_PALETA_STAVKA, COL_PALS_ID, stavkaID, COL_PALS_BROJ_ZBIRNE))
End Function

Private Function PalsStornirano(ByVal stavkaID As String) As String
    PalsStornirano = NzTx(LookupValue(TBL_PALETA_STAVKA, COL_PALS_ID, stavkaID, COL_STORNIRANO))
End Function

Private Function PalStornirano(ByVal palID As String) As String
    PalStornirano = NzTx(LookupValue(TBL_PALETA, COL_PAL_ID, palID, COL_STORNIRANO))
End Function

Private Function PalStatus(ByVal palID As String) As String
    PalStatus = NzTx(LookupValue(TBL_PALETA, COL_PAL_ID, palID, COL_PAL_STATUS))
End Function

Private Function PalBrGajbica(ByVal palID As String) As Long
    Dim v As Variant: v = LookupValue(TBL_PALETA, COL_PAL_ID, palID, COL_PAL_BR_GAJBICA)
    If IsNumeric(v) Then PalBrGajbica = CLng(v)
End Function

' Vrednost iz kolone `col` u redu ciji je prvi stubac = label. Sentinel -99999 kad
' red ne postoji -> test vidljivo pada umesto da tiho poredi nulu.
Private Function ReportCellByLabel(ByVal arr As Variant, ByVal label As String, _
                                   ByVal col As Long) As Double
    If Not IsArray(arr) Then
        ReportCellByLabel = -99999
        Exit Function
    End If
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If StrComp(Trim$(CStr(arr(i, 1))), label, vbTextCompare) = 0 Then
            If IsNumeric(arr(i, col)) Then ReportCellByLabel = CDbl(arr(i, col))
            Exit Function
        End If
    Next i
    ReportCellByLabel = -99999
End Function

' Broj AKTIVNIH tblNovac redova koji nose dati BIM marker. Meri pripadnost izvodu
' onako kako je meri i produkcijski kod (samo marker, bez pogadjanja).
Private Function AktivnihSaMarkerom(ByVal bimID As String) As Long
    Dim data As Variant: data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    Dim cNap As Long, cSt As Long
    cNap = GetColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA)
    cSt = GetColumnIndex(TBL_NOVAC, COL_STORNIRANO)
    If cNap = 0 Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(data, 1)
        If cSt = 0 Or UCase$(Trim$(CStr(data(i, cSt)))) <> "DA" Then
            If BimIdFromNapomena(CStr(data(i, cNap))) = Trim$(bimID) Then n = n + 1
        End If
    Next i
    AktivnihSaMarkerom = n
End Function

Private Function NovStornirano(ByVal novID As String) As String
    NovStornirano = NzTx(LookupValue(TBL_NOVAC, COL_NOV_ID, novID, COL_STORNIRANO))
End Function

Private Function BimObradjeno(ByVal bimID As String) As String
    BimObradjeno = NzTx(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_OBRADJENO))
End Function

Private Function BimStornirano(ByVal bimID As String) As String
    BimStornirano = NzTx(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_STORNIRANO))
End Function

Private Function OtkIsplaceno(ByVal otkID As String) As String
    OtkIsplaceno = NzTx(LookupValue(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_ISPLACENO))
End Function

Private Function OtkDatumIsplate(ByVal otkID As String) As String
    OtkDatumIsplate = NzTx(LookupValue(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_DATUM_ISPLATE))
End Function

Private Function OtkOtpremnicaID(ByVal blkID As String) As String
    OtkOtpremnicaID = NzTx(LookupValue(TBL_OTKUP, COL_OTK_ID, blkID, COL_OTK_OTPREMNICA_ID))
End Function

' Saldo (Ulaz +, Izlaz -) za entitet+tip -> iz produkcijskog GetAmbalazeStanje.
Private Function AmbSaldo(ByVal entID As String, ByVal entTip As String, ByVal tip As String) As Long
    On Error GoTo EH
    Dim arr As Variant
    arr = GetAmbalazeStanje(entID, entTip)
    If Not IsArray(arr) Then Exit Function
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Trim$(CStr(arr(i, 1))) = tip Then
            AmbSaldo = CLng(arr(i, 2))
            Exit Function
        End If
    Next i
    Exit Function
EH:
    AmbSaldo = -99999      ' sentinel -> test vidljivo pada
End Function

' Direktan upis BrojZbirne na otpremnicu (za T02 simulaciju losega premestaja).
Private Sub ForceSetOtpremnicaZbirna(ByVal otpBroj As String, ByVal novaZbirna As String)
    Dim c As Collection: Set c = FindRows(TBL_OTPREMNICA, COL_OTP_BROJ, otpBroj)
    Dim k As Long
    For k = 1 To c.count
        UpdateCell TBL_OTPREMNICA, CLng(c(k)), COL_OTP_BROJ_ZBIRNE, novaZbirna
    Next k
End Sub

Private Function NzTx(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzTx = "" Else NzTx = Trim$(CStr(v))
End Function

' ============================================================
' ASSERT + REPORT (mirror modTestPalete stila)
' ============================================================

Private Sub Chk(ByVal cond As Boolean, ByVal nm As String)
    If cond Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbCrLf
    Else
        Fail nm
    End If
End Sub

Private Sub ChkEq(ByVal act As Variant, ByVal exp As Variant, ByVal nm As String)
    If CStr(act) = CStr(exp) Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbCrLf
    Else
        Fail nm & " [dobijeno=" & CStr(act) & " ocekivano=" & CStr(exp) & "]"
    End If
End Sub

Private Sub ChkEqD(ByVal act As Double, ByVal exp As Double, ByVal nm As String)
    If Abs(act - exp) <= 0.001 Then
        mPass = mPass + 1
        mReport = mReport & "OK    " & nm & vbCrLf
    Else
        Fail nm & " [dobijeno=" & CStr(act) & " ocekivano=" & CStr(exp) & "]"
    End If
End Sub

Private Sub Fail(ByVal nm As String)
    mFail = mFail + 1
    mFails = mFails & " - " & nm & vbCrLf
    mReport = mReport & "PAO   " & nm & vbCrLf
End Sub

Private Sub ReportResults()
    Dim hdr As String
    hdr = "STORNO TEST SUITE  ->  PASS=" & mPass & "  FAIL=" & mFail
    Debug.Print String(60, "=")
    Debug.Print hdr
    Debug.Print String(60, "=")
    Debug.Print mReport

    Dim msg As String
    msg = hdr & vbCrLf & vbCrLf
    If mFail = 0 Then
        msg = msg & "Svi testovi PROSLI. (Detalji: Immediate / Ctrl+G)"
    Else
        msg = msg & "PALI testovi:" & vbCrLf & mFails & vbCrLf & "(Detalji: Immediate / Ctrl+G)"
    End If
    MsgBox msg, IIf(mFail = 0, vbInformation, vbExclamation), APP_NAME
End Sub
