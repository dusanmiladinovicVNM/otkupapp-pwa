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
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_STORNO_VEZE

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
    T11_IspravkaOtpremniceNovaZbirna

    tx.RollbackTx
    Set tx = Nothing

    ReportResults
    Exit Sub
EH:
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "modTestStorno.RunStornoTestSuite"
    MsgBox "Greska u test suite-u: " & Err.description & vbCrLf & vbCrLf & _
           "Do sada: PASS=" & mPass & " FAIL=" & mFail, vbCritical, APP_NAME
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
' T11 - ISPRAVKA otpremnice sa NOVOM zbirnom (malina 1:1) + prijemnica/palete.
' Regres test za bug: stara zbirna nulirana (umesto storno), prijemnica/palete
' zaglavljene na staroj. Ocekivano: stara zbirna STORNIRANA, prijemnica + paleta-
' stavke PRESELJENE na novu, blok prevezan, nova zbirna konzistentna.
' ============================================================
Private Sub T11_IspravkaOtpremniceNovaZbirna()
    Const S As String = "T11 ispravka otpremnice -> nova zbirna: "

    ' Stara: otpremnica OA11 (kg 100) u zbirni Z11 + blok + prijemnica + paleta-stavka.
    SeedZbirna "SVT-Z11", "I", 100, 10
    SeedOtpremnica "SVT-OA11", "SVT-Z11", "I", 100, 10
    SeedOtkupBlok "SVT-BLK11", "SVT-OA11-ID-I", "SVT-Z11"
    SeedPrijemnica "SVT-P11", "SVT-Z11", "I", 100, 10
    SeedPaletaStavka "SVT-PS11", "SVT-P11", "SVT-Z11", "I", 100, 10

    ' Faza 1: ISPRAVKA -> storno stare otpremnice + context.
    Dim res As Object
    Set res = modStornoFlow.RunOtpremnicaCorrection("SVT-OA11", SV_MODE_ISPRAVKA)
    Chk CBool(res("needsForm")), S & "ISPRAVKA trazi novu otpremnicu"
    Dim cid As String: cid = CStr(res("correctionID"))

    ' Operater snima NOVU otpremnicu (kg 90) sa NOVOM zbirnom (malina 1:1).
    SeedOtpremnica "SVT-OB11", "SVT-Z11B", "I", 90, 9
    SeedZbirna "SVT-Z11B", "I", 0, 0

    ' Faza 2: complete -> prevezi blok + prijemnicu/palete, recalc nove, storno stare.
    Set res = modStornoFlow.CompleteOtpremnicaIspravka(cid, "SVT-OB11")
    Chk CBool(res("success")), S & "CompleteOtpremnicaIspravka uspeo"

    Chk Not ZbirnaPostoji("SVT-Z11"), S & "STARA zbirna STORNIRANA (ne nulirana)"
    ChkEq OtkOtpremnicaID("SVT-BLK11"), "SVT-OB11-ID-I", S & "blok prevezan na novu otpremnicu"
    ChkEq PrjBrojZbirne("SVT-P11"), "SVT-Z11B", S & "prijemnica preseljena na novu zbirnu"
    ChkEq PalsBrojZbirne("SVT-PS11"), "SVT-Z11B", S & "paleta-stavka preseljena na novu zbirnu"
    Chk modDokumentInvariant.IsZbirnaConsistent("SVT-Z11B"), S & "nova zbirna = zbir otpremnica (90)"
End Sub

' ============================================================
' SEED HELPERS (upis po IMENU kolone -> otporno na redosled)
' ============================================================

' Zbirna: jedan red po klasi. BrojZbirne + Klasa + KG + AMB (+ ZbirnaID, Datum).
Private Sub SeedZbirna(ByVal broj As String, ByVal klasa As String, _
                       ByVal kg As Double, ByVal amb As Long)
    SvAppend TBL_ZBIRNA, _
        Array(COL_ZBR_ID, COL_ZBR_DATUM, COL_ZBR_BROJ, COL_ZBR_KOLICINA, _
              COL_ZBR_TIP_AMB, COL_ZBR_KOL_AMB, COL_ZBR_VRSTA, COL_ZBR_SORTA, COL_ZBR_KLASA), _
        Array(broj & "-ID-" & klasa, Date, broj, kg, "SVT-A", amb, "SVT-VOCE", "SVT-SORTA", klasa)
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

Private Sub SeedPaletaStavka(ByVal stavkaID As String, ByVal brojPrij As String, _
                             ByVal brojZbirne As String, ByVal klasa As String, _
                             ByVal kg As Double, ByVal amb As Long)
    SvAppend TBL_PALETA_STAVKA, _
        Array(COL_PALS_ID, COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, COL_PALS_KLASA, _
              COL_PALS_NETO, COL_PALS_AMBALAZA), _
        Array(stavkaID, brojPrij, brojZbirne, klasa, kg, amb)
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
