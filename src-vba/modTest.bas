Attribute VB_Name = "modTest"
Option Explicit

' ============================================================
' modTest
' Test suite koja pada na PONASANJU, ne na sintaksi. Cilj je da razlikuje
' ispravan od pokvarenog koda -- suite koja je zelena nad cistim kodom, a nije
' dokazano crvena nad pokvarenim, ne dokazuje nista.
'
' Pokretanje: tools/run_vba.py zove Run("RunAllTests") nad temp kopijom
' fixture-a (tests/fixtures/otkup_test.xlsm, pravi ga tools/make_fixture.py).
' Rezultat ide u last_run.txt PORED SVESKE (dakle u temp folder), prvi red
' "TESTS=n FAIL=m". Driver taj fajl cita; nema fajla = pad.
'
' Compile signal stize sam: da bi se RunAllTests uopste pokrenuo, VBA mora da
' kompajlira modTest i sve sto on referencira -- a to je bas kod pod testom
' (frmOtkup, modOtkup, modOtkupBlok). Zato ovde nema posebnog compile gate-a.
'
' Greska se hvata PO TESTU: jedan pad ne obara ostale, i u ispisu stoji ime
' bas tog testa.
'
' NOVI UI (frmOtkupUI + modOtkupUI) ima svoja tri testa, 4-6. Legacy forma se
' NE gasi (docs/UI_MIGRACIJA_KATALOG.md), pa oba skupa stoje jedan pored drugog
' -- ugovor je isti, kod je namerno dvostruk.
'
' UPIS ZBIRNE I PRIJEMNICE (F3/F4, modDokUnos) ima testove 7-11. Oni ne diraju
' formu: pravilo unosa zivi u modulu bez ijedne kontrole, pa se i proverava
' tamo -- brzo, bez gradnje ekrana i bez stanja koje ostaje za sobom.
' ============================================================

' --- Fixture konstante (moraju da prate tools/make_fixture.py) --------------
Private Const FX_DATUM As String = "15.3.2026"      ' FIXTURE_DATE, d.m.yyyy
Private Const FX_ZBIRNA As String = "ZB-TEST-1"     ' zbirna na OTP-TEST-1
Private Const FX_BROJ_OTP As String = "1/TEST"      ' BrojOtpremnice OTP-TEST-1
Private Const FX_KOOPERANT As String = "KOOP-TEST-1"
Private Const FX_KOOPERANT2 As String = "KOOP-TEST-2"
Private Const FX_OTP_ID As String = "OTP-TEST-1"    ' otpremnica koja nosi FX_BROJ_OTP
Private Const FX_PARCELA As String = "PAR-TEST-1"   ' parcela kooperanta KOOP-TEST-1
Private Const FX_VOZAC As String = "VOZ-TEST-1"
' Otkupno mesto IZ FIXTURE-a. Mora biti stvarno: kapija IsplataBlokProblem
' poredi otkupno mesto bloka sa ovim, pa izmisljen ID vise ne prolazi.
Private Const FX_STANICA As String = "STA-TEST-1"
Private Const FX_STANICA2 As String = "STA-DRUGO"    ' ne postoji -- "tudje OM"
' Otkupni blokovi iz fixture-a: prvi je KOOP-TEST-1, drugi KOOP-TEST-2, oba na
' FX_STANICA. Vrednost prvog = Kolicina 400 * Cena 50; fixture nema tblNovac,
' pa je neisplaceni ostatak jednak vrednosti -- ali test ga i dalje racuna
' kroz GetUplataForOtkup, da tvrdnja ne padne ako fixture jednom dobije uplate.
Private Const FX_BLOK As String = "OTK-TEST-1"
Private Const FX_BLOK_TUDJ As String = "OTK-TEST-2"
Private Const FX_BLOK_VREDNOST As Double = 20000
' Faktura iz fixture-a (tblFakture): pripada FX_KUPAC, iznos FX_FAKTURA_IZNOS.
Private Const FX_FAKTURA As String = "FAK-TEST-1"
Private Const FX_FAKTURA_IZNOS As Double = 10000
' Faktura BEZ evidentiranog iznosa: kapija nad uplatom se na nju ne primenjuje,
' i to pravilo mora da prezivi popravku te kapije.
Private Const FX_FAKTURA_BEZ_IZNOSA As String = "FAK-TEST-0"
' Broj dokumenta za novac/ambalazu koji NE postoji ni u tblAmbalaza ni u
' tblNovac -- provera duplikata mora da ga propusti.
Private Const FX_BROJ_NOVAC As String = "NOVUNOS-TEST-1"
Private Const FX_VRSTA As String = "TESTVOCE"
Private Const FX_SORTA As String = "TESTSORTA"
Private Const FX_TIP_AMB As String = "12/1"         ' AMB_12_1, TezinaGajbiceKg = 1
' Kupca u fixture-u NEMA i ne treba ga: provere gledaju samo da li je izabran
' (Len > 0). Upis, koji bi trazio postojeceg, ovi testovi ne voze.
Private Const FX_KUPAC As String = "KUP-TEST-1"
' Zbir OTP-TEST-1 -- jedine otpremnice koja nosi FX_ZBIRNA. Zbirna mora tacno
' toliko da prijavi, inace je kapija obara.
Private Const FX_ZBIRNA_KG As Double = 1000
Private Const FX_ZBIRNA_AMB As Long = 100

Private Const ERR_ASSERT As Long = vbObjectError + 9500
Private Const ERR_GOLDEN As Long = vbObjectError + 9501

Private m_Total As Long
Private m_Failed As Long
Private m_Report As String

' ============================================================
' Ulazna tacka
' ============================================================
Public Sub RunAllTests()
    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    m_Total = 0
    m_Failed = 0
    m_Report = ""

    RunOne 1
    RunOne 2
    RunOne 3
    RunOne 4
    RunOne 5
    RunOne 6
    RunOne 7
    RunOne 8
    RunOne 9
    RunOne 10
    RunOne 11
    RunOne 12
    RunOne 13
    RunOne 14
    RunOne 15
    RunOne 16
    RunOne 17
    RunOne 18
    RunOne 19
    RunOne 20
    RunOne 21

    SetTestMode prevMode
    WriteResultFile
End Sub

' Svaki test se zove kroz ovu omotnicu: broji se, greska mu se hvata i upisuje
' pod NJEGOVIM imenom. Ime se razresava pre poziva da bi bilo poznato i kad
' test pukne.
Private Sub RunOne(ByVal idx As Long)
    Dim nm As String
    Dim errNum As Long, errDesc As String
    nm = TestName(idx)

    On Error GoTo EH
    m_Total = m_Total + 1
    InvokeTest idx
    AppendReport nm, "OK", ""
    Exit Sub

EH:
    ' Err se cita PRE ciscenja: CleanupPosleTesta ide kroz On Error Resume Next
    ' (OtkupUI_Release ga ima), a to brise Err -- bez ovoga bi izvestaj o padu
    ' ostao prazan.
    errNum = Err.Number
    errDesc = Err.description
    m_Failed = m_Failed + 1
    CleanupPosleTesta
    ' Pad bez opisa je vec jednom kostao dva rana dijagnostike: "FAIL T_X" bez
    ' razloga ne kaze operateru nista. Broj greske je tada jedini trag.
    If Len(errDesc) = 0 Then errDesc = "greska bez opisa (Err.Number=" & errNum & ")"
    AppendReport nm, "FAIL", errDesc
End Sub

' Test koji je pao NIJE stigao do svog ReleaseOtkupUIForm, pa modul novog UI-ja
' (mFrm, Btns, kes tabela) i aktivna otpremnica u modScrDokumenti ostaju
' zaprljani. Sledeci test bi tada gradio ekran nad ostacima prethodnog i pao BEZ
' SVOJE KRIVICE -- jedna sabotaza obarala bi dva testa, pa bi drugi pad bio lazan
' trag. (Dokazano: sabotaza parcela-tekst obarala je i T_ClearForm_Ugovor, sa
' Err.Number=0 i praznim opisom.)
'
' Ciscenje je idempotentno (OtkupUI_Release je ceo pod On Error Resume Next,
' Scr_OtpOtkazi samo prazni tri promenljive), pa je bezbedno i posle testa koji
' formu nikad nije napravio. Samu formu otpusta odmotavanje steka -- ovde ostaje
' ono sto zivi na MODULIMA i sto odmotavanje ne dira.
Private Sub CleanupPosleTesta()
    On Error Resume Next
    modOtkupUI.OtkupUI_Release
    modScrDokumenti.Scr_OtpOtkazi
End Sub

Private Function TestName(ByVal idx As Long) As String
    Select Case idx
        Case 1: TestName = "T_PosleSnimanja_ZadrzavaKontekstOtpremnice"
        Case 2: TestName = "T_PosleSnimanja_ZadrzavaZbirnu"
        Case 3: TestName = "T_ClearForm_BrisePartnera"
        Case 4: TestName = "T_ParseDatum_Ugovor"
        Case 5: TestName = "T_ParcelaID_IzSkriveneKolone"
        Case 6: TestName = "T_ClearForm_Ugovor"
        Case 7: TestName = "T_ZbirnaValidiraj_TraziVozaca"
        Case 8: TestName = "T_ZbirnaValidiraj_MoraDaSeSlazeSaOtpremnicama"
        Case 9: TestName = "T_PrijemnicaValidiraj_TraziKupca"
        Case 10: TestName = "T_BrutoNeto_PoRezimu"
        Case 11: TestName = "T_ScrSave_RutaPoRezimu"
        Case 12: TestName = "T_IsplataValidiraj_TipNovcaPoIzboru"
        Case 13: TestName = "T_UplataValidiraj_FakturaOdlucujeTip"
        Case 14: TestName = "T_ReversValidiraj_SmerJeObavezan"
        Case 15: TestName = "T_IsplataBlokGuard_VlasnistvoITrenutniOstatak"
        Case 16: TestName = "T_NerazresenIzbor_NeProlaziKaoPrazno"
        Case 17: TestName = "T_WriterGuard_OdbijaTudjBlok"
        Case 18: TestName = "T_UplataGuard_VecPlacenaFaktura"
        Case 19: TestName = "T_WriterGuard_AvansSaldoOM"
        Case 20: TestName = "T_F8_TipBiraTabeluIKolone"
        Case 21: TestName = "T_StornoDok_KapijePreUpisa"
        Case Else: TestName = "T_Nepoznat_" & idx
    End Select
End Function

' Direktan poziv (ne Application.Run) -- tako VBA mora da kompajlira i test i
' sve sto test referencira.
Private Sub InvokeTest(ByVal idx As Long)
    Select Case idx
        Case 1: T_PosleSnimanja_ZadrzavaKontekstOtpremnice
        Case 2: T_PosleSnimanja_ZadrzavaZbirnu
        Case 3: T_ClearForm_BrisePartnera
        Case 4: T_ParseDatum_Ugovor
        Case 5: T_ParcelaID_IzSkriveneKolone
        Case 6: T_ClearForm_Ugovor
        Case 7: T_ZbirnaValidiraj_TraziVozaca
        Case 8: T_ZbirnaValidiraj_MoraDaSeSlazeSaOtpremnicama
        Case 9: T_PrijemnicaValidiraj_TraziKupca
        Case 10: T_BrutoNeto_PoRezimu
        Case 11: T_ScrSave_RutaPoRezimu
        Case 12: T_IsplataValidiraj_TipNovcaPoIzboru
        Case 13: T_UplataValidiraj_FakturaOdlucujeTip
        Case 14: T_ReversValidiraj_SmerJeObavezan
        Case 15: T_IsplataBlokGuard_VlasnistvoITrenutniOstatak
        Case 16: T_NerazresenIzbor_NeProlaziKaoPrazno
        Case 17: T_WriterGuard_OdbijaTudjBlok
        Case 18: T_UplataGuard_VecPlacenaFaktura
        Case 19: T_WriterGuard_AvansSaldoOM
        Case 20: T_F8_TipBiraTabeluIKolone
        Case 21: T_StornoDok_KapijePreUpisa
    End Select
End Sub

' ============================================================
' Testovi
' ============================================================

' Posle snimanja otkupnog lista kontekst otpremnice mora da ostane: datum se NE
' brise, jer sledeci blok ide u niz istog datuma. Pada ako se u ClearOtkupFields
' vrati brisanje datuma (txtDatum.value = "").
Private Sub T_PosleSnimanja_ZadrzavaKontekstOtpremnice()
    Dim f As frmOtkup
    Set f = NewOtkupForm()

    f.ClearOtkupFields

    AssertEq f.txtDatum.value, FX_DATUM, _
             "datum posle snimanja mora da ostane datum otpremnice"

    AssertSnapshot DumpKontrole(f), "PosleSnimanja_KontekstOtpremnice"

    Unload f
End Sub

' Broj zbirne ostaje popunjen posle snimanja: sledeci blok iste otpremnice mora
' da dobije istu zbirnu, inace operater kuca broj iznova na svaki unos. Pada ako
' se u ClearOtkupFields vrati txtBrojZbirne.value = "".
Private Sub T_PosleSnimanja_ZadrzavaZbirnu()
    Dim f As frmOtkup
    Set f = NewOtkupForm()

    f.ClearOtkupFields
    AssertEq f.txtBrojZbirne.value, FX_ZBIRNA, _
             "broj zbirne mora da ostane popunjen posle snimanja"

    ' Drugi blok nad istom otpremnicom -- posle jos jednog snimanja zbirna je ista.
    f.cmbKooperant.value = FX_KOOPERANT2
    f.ClearOtkupFields
    AssertEq f.txtBrojZbirne.value, FX_ZBIRNA, _
             "drugi blok mora da dobije istu zbirnu"

    Unload f
End Sub

' Kooperant se BRISE posle snimanja -- sledeci unos je nov partner. Suprotno od
' prethodna dva testa: ovde je brisanje trazeno ponasanje. Pada ako se iz
' ClearOtkupFields ukloni cmbKooperant.value = "".
Private Sub T_ClearForm_BrisePartnera()
    Dim f As frmOtkup
    Set f = NewOtkupForm()

    ' Preduslov: bez ovoga bi test bio zelen i kad kontrola uopste ne prima
    ' vrednost, pa ne bi merio nista.
    AssertEq f.cmbKooperant.value, FX_KOOPERANT, _
             "preduslov: kooperant je postavljen pre ciscenja"

    f.ClearOtkupFields

    AssertEq f.cmbKooperant.value, "", _
             "kooperant mora da bude obrisan posle snimanja"

    Unload f
End Sub

' ============================================================
' Novi UI (frmOtkupUI + modOtkupUI)
' ============================================================

' DATUM DOKUMENTA ide u tblOtkup i u kontekst (predlog broja, zakljucavanje
' stanice), pa "necitljivo" mora da bude 0 -- nikad priblizan datum. Parser je
' NAMERNO deterministican (modParse.TryParseDateValue): CDate isti tekst cita po
' Windows locale-u, pa bi "01.02.2026" na MDY masini bio 2. januar a na DMY
' masini 1. februar. Pada ako se ParseDatum vrati na IsDate/CDate ili ako se
' izgubi skidanje trailing tacke.
Private Sub T_ParseDatum_Ugovor()
    AssertEq modOtkupUI.ParseDatum(""), 0, "prazno polje nije datum"
    AssertEq modOtkupUI.ParseDatum("   "), 0, "sami razmaci nisu datum"
    AssertEq modOtkupUI.ParseDatum("besmislica"), 0, "necitljiv tekst nije datum"

    AssertEq modOtkupUI.ParseDatum("11.08.2026"), CDbl(DateSerial(2026, 8, 11)), _
             "d.m.yyyy se cita kao dan.mesec.godina"

    ' Trailing tacka je nacin na koji se datum kod nas pise ("11.08.2026."), pa
    ' se skida umesto da obori unos. Petlja, ne jedno skidanje.
    AssertEq modOtkupUI.ParseDatum("11.08.2026."), CDbl(DateSerial(2026, 8, 11)), _
             "trailing tacka se skida, ne obara unos"
    AssertEq modOtkupUI.ParseDatum("11.08.2026.."), CDbl(DateSerial(2026, 8, 11)), _
             "skidaju se SVE trailing tacke, ne samo poslednja"

    ' AUD-007: DateSerial se na nemogucem datumu PRELIVA (30.02 -> 2.3, mesec 13
    ' -> januar sledece godine) umesto da pukne. Round-trip u parseru to odbija --
    ' inace bi dokument tiho dobio pomeren datum.
    AssertEq modOtkupUI.ParseDatum("30.02.2026"), 0, _
             "nepostojeci dan se odbija, ne preliva u sledeci mesec"
    AssertEq modOtkupUI.ParseDatum("01.13.2026"), 0, _
             "mesec 13 se odbija, ne preliva u sledecu godinu"

    ' Kapija poslovnih godina dolazi iz zajednickog parsera (modParse), ali se
    ' vidi kroz ovo polje -- zato stoji ovde, uz ostatak ugovora.
    AssertEq modOtkupUI.ParseDatum("11.08.1899"), 0, "godina van poslovnog opsega"
End Sub

' ID PARCELE JE SKRIVENA DRUGA KOLONA combo-a, kao kod svih ostalih dropdown-a
' (PartnerID / modComboBinding.GetComboID). Regres koji ovaj test cuva: ID se
' nekad vadio iz prikaznog teksta trazenjem " - ", a FillParcele gradi prikaz sa
' " " & ChrW(183) & " " -- separator se nikad nije nasao, pa je ceo prikazni
' string odlazio u ParcelaID i u tblOtkup. Pada ako se ID opet cita iz teksta,
' ili ako se izgubi provera vidljivosti polja.
Private Sub T_ParcelaID_IzSkriveneKolone()
    Dim f As frmOtkupUI, fr As Object, CB As MSForms.ComboBox
    Set f = NewOtkupUIForm()

    Set fr = f.Controls("zForm").Controls("fgParcela")
    Set CB = fr.Controls("fgParcelaT")
    fr.Visible = True

    ' Isti oblik koji gradi FillParcele: prikaz u koloni 1, ID u koloni 2.
    ' Prikaz NAMERNO nosi separator koji nije " - ".
    CB.Clear
    CB.ColumnCount = 2
    CB.BoundColumn = 1
    CB.TextColumn = 1
    CB.AddItem "1001   " & ChrW(183) & "   Malina   " & ChrW(183) & "   1,20 ha"
    CB.List(0, 1) = FX_PARCELA
    CB.ListIndex = 0

    ' Preduslov: bez ovoga bi test bio zelen i kad combo uopste ne prima stavke.
    AssertEq CB.ListCount, 1, "preduslov: parcela je u listi"

    AssertEq modOtkupUI.ParcelaID(), FX_PARCELA, _
             "ID parcele dolazi iz skrivene kolone, ne iz prikaznog teksta"

    CB.ListIndex = -1
    AssertEq modOtkupUI.ParcelaID(), "", "bez izabrane parcele dokument ne dobija ID"

    ' PRACENJE_PARCELA iskljuceno -> polje je sakriveno. Zatecen izbor tada NE
    ' sme da procuri u dokument.
    CB.ListIndex = 0
    fr.Visible = False
    AssertEq modOtkupUI.ParcelaID(), "", "sakriveno polje ne salje parcelu u dokument"

    ReleaseOtkupUIForm f
End Sub

' UGOVOR ClearForm-a, isti kao frmOtkup.ClearOtkupFields (.claude/rules/
' otkup-i-dokumenta.md odeljak 1 i 5): datum i broj zbirne su KONTEKST
' otpremnice i ostaju, partner se brise. Uz to i nova razlika koju legacy nema:
' bez aktivne otpremnice datum se vraca na danas.
'
' Zasto datum: otpremnica 8/220726 od 22.07 dobijala je blok 8/110826 od 11.08 --
' vracanje na danas je i broj i datum bloka odvlacilo iz niza otpremnice.
Private Sub T_ClearForm_Ugovor()
    Dim f As frmOtkupUI, zf As Object, ctx As Object
    Dim datumBloka As String, danas As String
    Set f = NewOtkupUIForm()
    Set zf = f.Controls("zForm")
    Set ctx = f.Controls("zCtx")

    ' Datum se izvodi iz danasnjeg, da NIKAD ne bude jednak "danas" -- zakucan
    ' datum bi jednog dana u godini prosao test i kad pravilo ne radi.
    datumBloka = Format$(Date - 30, "dd.mm.yyyy")

    ' Blok koji se upravo snimio nad aktivnom otpremnicom.
    '
    ' Datum i zbirna se postavljaju kroz ApplyPrefill, ne pisanjem u kontrolu:
    ' to je put kojim ih i produkcija dobija (izbor otpremnice), i jedini koji
    ' ide pod mLoading. Direktan upis u fgDatum okine OnDatumChanged, a on trazi
    ' stanica-lock i predlog broja SA PITANJEM GOOGLE-U -- mreza u testu.
    ' Kilogrami i ambalaza su TextBox-evi: njihova promena samo preracunava
    ' vrednost, pa idu direktno.
    modScrDokumenti.Scr_OtpTestSet FX_OTP_ID, FX_BROJ_OTP
    modOtkupUI.ApplyPrefill "datum=" & datumBloka & "|brzbirne=" & FX_ZBIRNA
    SetPolje zf, "fgKgI", "123,4"
    SetPolje zf, "fgKolAmb", "10"
    ctx.Controls("cbKupac").value = FX_KOOPERANT

    ' Preduslovi: bez njih bi test bio zelen i kad kontrole uopste ne primaju
    ' vrednost, pa ne bi merio nista.
    AssertEq Polje(zf, "fgDatum"), datumBloka, "preduslov: datum otpremnice je upisan"
    AssertEq Polje(zf, "fgBrZbir"), FX_ZBIRNA, "preduslov: broj zbirne je upisan"
    AssertEq Polje(zf, "fgKgI"), "123,4", "preduslov: kilogrami su upisani"
    AssertEq ctx.Controls("cbKupac").value, FX_KOOPERANT, "preduslov: partner je upisan"

    modOtkupUI.ClearForm

    ' 1) DATUM OSTAJE -- sledeci blok ide u niz istog datuma otpremnice.
    AssertEq Polje(zf, "fgDatum"), datumBloka, _
             "dok je otpremnica aktivna datum se NE vraca na danas"
    ' 2) BROJ ZBIRNE OSTAJE -- svi blokovi jedne otpremnice idu na istu zbirnu.
    AssertEq Polje(zf, "fgBrZbir"), FX_ZBIRNA, _
             "broj zbirne je kontekst -- ne brise se posle snimanja"
    ' 3) PARTNER SE BRISE -- sledeci unos je nov kooperant. Obrnut smer od prva
    '    dva: ovde je brisanje trazeno ponasanje.
    AssertEq ctx.Controls("cbKupac").value, "", _
             "partner mora da bude obrisan posle snimanja"
    ' ... a podaci bloka odlaze sa njim.
    AssertEq Polje(zf, "fgKgI"), "", "kilogrami se brisu posle snimanja"
    AssertEq Polje(zf, "fgKolAmb"), "", "kolicina ambalaze se brise posle snimanja"

    ' BEZ AKTIVNE OTPREMNICE datum se vraca na danas: prazno ili staro polje bi
    ' bila greska koju operater mora da ispravlja pri svakom novom dokumentu.
    modScrDokumenti.Scr_OtpOtkazi
    modOtkupUI.ApplyPrefill "datum=" & datumBloka & "|brzbirne=" & FX_ZBIRNA
    danas = Format$(Date, "dd.mm.yyyy")
    modOtkupUI.ClearForm
    AssertEq Polje(zf, "fgDatum"), danas, _
             "bez aktivne otpremnice datum se vraca na danas"

    ReleaseOtkupUIForm f
End Sub

' ============================================================
' Upis zbirne (F3) i prijemnice (F4) -- modDokUnos + ruta u modScrDokumenti
'
' Ovi testovi ne grade formu: pravilo unosa zivi u modulu bez ijedne kontrole,
' pa se tamo i proverava. Datum se ne postavlja jer ga nijedna provera ne cita
' (njega proverava ljuska, pre poziva ekrana -- modOtkupUI.CommitDokument).
' ============================================================

' VOZAC JE ENTITET NIZA ZBIRNE (Z3a): po njemu se broji, njegova je tura i on
' nosi robu kupcu. Zato je prva provera, pre kupca i pre broja -- isti redosled
' kao frmDokumenta.btnUnosZbr_Click. Pada ako se provera ukloni ili spusti
' ispod ostalih.
Private Sub T_ZbirnaValidiraj_TraziVozaca()
    Dim p As Object, fokus As String, res As String

    Set p = ZbirnaUnosKojiSeSlaze()
    p("vozacID") = ""
    res = modDokUnos.ZbirnaValidiraj(p, fokus)
    AssertEq res, Poruka("DOKUNOS_ERR_VOZAC"), "zbirna bez vozaca se odbija"
    AssertEq fokus, "vozacID", "fokus se vraca na vozaca"

    ' Obrnut smer: sa vozacem ta poruka vise ne dolazi. Bez ovoga bi test bio
    ' zelen i kad rutina odbija SVE.
    Set p = ZbirnaUnosKojiSeSlaze()
    p("kupacID") = ""
    res = modDokUnos.ZbirnaValidiraj(p, fokus)
    AssertEq res, Poruka("DOKUNOS_ERR_KUPAC"), "sa vozacem zbirna staje tek na kupcu"
    AssertEq fokus, "kupacID", "fokus se vraca na kupca"
End Sub

' ZBIRNA JE POKLOPAC NAD OTPREMNICAMA, ne slobodan unos: kilogrami i ambalaza
' moraju da se poklope sa nestorniranim otpremnicama tog broja. U legacy je to
' hard-gate (btnUnosZbr_Click: "If Not UpdateValidacija()") koji NE zavisi od
' podesavanja VALIDACIJA_UNOSA. Pada ako se kapija ukloni, gejtuje tim
' podesavanjem, ili ako se ambalaza prestane porediti.
Private Sub T_ZbirnaValidiraj_MoraDaSeSlazeSaOtpremnicama()
    Dim p As Object, fokus As String, res As String
    Dim prevVal As String, resBezVal As String

    Set p = ZbirnaUnosKojiSeSlaze()
    p("kolicinaI") = FX_ZBIRNA_KG - 100          ' 100 kg manje nego sto izvor nosi
    res = modDokUnos.ZbirnaValidiraj(p, fokus)
    AssertEq res, Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA"), _
             "zbirna koja ne prijavljuje sve kilograme otpremnica se odbija"

    ' Ambalaza se poredi zasebno od kilograma.
    Set p = ZbirnaUnosKojiSeSlaze()
    p("kolAmb") = FX_ZBIRNA_AMB - 10
    res = modDokUnos.ZbirnaValidiraj(p, fokus)
    AssertEq res, Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA"), _
             "zbirna sa pogresnim brojem gajbi se odbija"

    ' Zbirna bez unete ambalaze se ne pusta ni kad se kilogrami slazu: razlika bi
    ' ispala 0 i kad izvor ambalazu ima (legacy uslov "zbrAmb > 0").
    Set p = ZbirnaUnosKojiSeSlaze()
    p("kolAmb") = 0
    res = modDokUnos.ZbirnaValidiraj(p, fokus)
    AssertEq res, Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA"), _
             "zbirna bez unete ambalaze ne prolazi kapiju"

    ' Obrnut smer: zbirna koja se u svemu slaze PROLAZI kapiju. Dalje je
    ' zaustavlja samo duplikat (FX_ZBIRNA vec postoji u fixture-u) -- druga
    ' provera i druga poruka, pa je razlika merljiva.
    Set p = ZbirnaUnosKojiSeSlaze()
    res = modDokUnos.ZbirnaValidiraj(p, fokus)
    AssertEq (res = Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA")), False, _
             "zbirna koja se poklapa sa otpremnicama prolazi kapiju"

    ' Kapija NE zavisi od VALIDACIJA_UNOSA: i sa iskljucenim podesavanjem zbirna
    ' koja ne pokriva svoje otpremnice mora da padne (legacy zove UpdateValidacija
    ' bezuslovno, za razliku od provera vrste/sorte/gajbi iznad).
    prevVal = GetConfigValue(CFG_VALIDACIJA_UNOSA)
    SetConfigValue CFG_VALIDACIJA_UNOSA, "NE"
    Set p = ZbirnaUnosKojiSeSlaze()
    p("kolicinaI") = FX_ZBIRNA_KG - 100
    resBezVal = modDokUnos.ZbirnaValidiraj(p, fokus)
    ' Podesavanje se vraca PRE tvrdnje: pala tvrdnja ne sme da ostavi iskljucenu
    ' validaciju ostatku suite-a nad istom sveskom.
    SetConfigValue CFG_VALIDACIJA_UNOSA, prevVal
    AssertEq resBezVal, Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA"), _
             "kapija vazi i kad je VALIDACIJA_UNOSA iskljucena"
End Sub

' PRIJEMNICA JE PRIJEM KOD KUPCA: bez kupca ne postoji, pa je on prva provera
' (kod zbirne je prvi vozac -- redosled nije stil nego navika operatera, isto
' kao u frmDokumenta.btnUnosPrij_Click). Uz to: broj zbirne je obavezan, jer je
' prijemnica prijem po JEDNOJ zbirnoj.
Private Sub T_PrijemnicaValidiraj_TraziKupca()
    Dim p As Object, fokus As String, res As String

    Set p = PrijemnicaUnosKojiProlazi()
    p("kupacID") = ""
    res = modDokUnos.PrijemnicaValidiraj(p, fokus)
    AssertEq res, Poruka("DOKUNOS_ERR_KUPAC"), "prijemnica bez kupca se odbija"
    AssertEq fokus, "kupacID", "fokus se vraca na kupca"

    ' Obrnut smer + red provera: sa kupcem prijemnica staje tek na vozacu.
    Set p = PrijemnicaUnosKojiProlazi()
    p("vozacID") = ""
    res = modDokUnos.PrijemnicaValidiraj(p, fokus)
    AssertEq res, Poruka("DOKUNOS_ERR_VOZAC"), "sa kupcem prijemnica staje tek na vozacu"
    AssertEq fokus, "vozacID", "fokus se vraca na vozaca"

    Set p = PrijemnicaUnosKojiProlazi()
    p("brojZbirne") = ""
    res = modDokUnos.PrijemnicaValidiraj(p, fokus)
    AssertEq res, Poruka("DOKUNOS_ERR_BROJ_ZBIRNE"), "prijemnica bez broja zbirne se odbija"
    AssertEq fokus, "brojZbirne", "fokus se vraca na broj zbirne"
End Sub

' BRUTO -> NETO PO REZIMU, i to razlicito -- zato je jedan test za oba:
'
'   PRIJEMNICA ima BrutoKg (tblPrijemnica): uneti bruto se zamrzava, u Kolicinu
'   ide neto, i to po klasama zasebno jer su im gajbice zasebne.
'
'   ZBIRNA ga NEMA i ne sme da ga dobije: ona je zbir SVOJIH otpremnica, a one
'   su vec u netu. Oduzimanje tare i drugi put spustilo bi kilograme ispod
'   izvora i oborilo bas kapiju iz testa 8. tblZbirna zato nema kolonu BrutoKg.
Private Sub T_BrutoNeto_PoRezimu()
    Dim pp As Object, pz As Object, fokus As String
    Dim resP As String, resZ As String, prevBruto As String

    prevBruto = GetConfigValue(CFG_OTKUP_BRUTO_UNOS)
    SetConfigValue CFG_OTKUP_BRUTO_UNOS, "DA"

    Set pp = PrijemnicaUnosKojiProlazi()
    pp("kolicinaI") = 110               ' bruto = 100 kg voca + 10 gajbica po 1 kg
    pp("kolAmb") = 10
    pp("dveKlase") = True
    pp("kolicinaII") = 55               ' bruto = 50 kg + 5 gajbica
    pp("cenaII") = 40
    pp("kolAmbII") = 5
    resP = modDokUnos.PrijemnicaValidiraj(pp, fokus)

    Set pz = ZbirnaUnosKojiSeSlaze()
    resZ = modDokUnos.ZbirnaValidiraj(pz, fokus)

    ' Podesavanje se vraca PRE tvrdnji: pala tvrdnja ne sme da ostavi ukljucen
    ' bruto rezim ostatku suite-a nad istom sveskom.
    SetConfigValue CFG_OTKUP_BRUTO_UNOS, prevBruto

    AssertEq pp("brutoKgI"), 110, "uneti bruto Kl.I se zamrzava u BrutoKg"
    AssertEq pp("kolicinaI"), 100, "u Kolicinu Kl.I ide neto (bruto - tara)"
    AssertEq pp("brutoKgII"), 55, "uneti bruto Kl.II se zamrzava u BrutoKg"
    AssertEq pp("kolicinaII"), 50, "u Kolicinu Kl.II ide neto (bruto - tara)"
    AssertEq resP, "", "prijemnica sa ispravnim bruto unosom prolazi provere"

    AssertEq pz("kolicinaI"), FX_ZBIRNA_KG, "zbirna se NE preracunava iz bruta"
    AssertEq pz.Exists("brutoKgI"), False, "zbirna nema BrutoKg (nema ga ni tabela)"
    AssertEq (resZ = Poruka("DOK_MSG_VALIDACIJA_NIJE_PROSLA")), False, _
             "zbirna se i u bruto rezimu poklapa sa (neto) otpremnicama"
End Sub

' RUTA PO REZIMU. Ekran samo prevodi polja i zove pravi modul; rezim koji jos
' nema svoj upis mora da ostane iskren -- "nije vezano", ne lazna potvrda.
' Pada ako se ruta izgubi (F3-F7 padnu na Case Else) ili ako se nepokriven
' rezim tiho propusti u neki od upisa.
'
' Dokaz rute je poruka koju SAMO taj modul ume da vrati. Zato se svakom rezimu
' popune polja koja dele sa ostalima (otkupno mesto, vrsta), pa ga zaustavi
' pravilo koje je iskljucivo njegovo.
Private Sub T_ScrSave_RutaPoRezimu()
    Dim p As Object

    ' F8 (storno) jos nema upis -- jedini preostali nepokriven rezim.
    Set p = PoljaEkrana(modScrDokumenti.modeKey("F8"))
    AssertEq modScrDokumenti.Scr_Save(p), Poruka("OTKUI_TODO_NEVEZANO"), _
             "nepokriven rezim vraca OTKUI_TODO_NEVEZANO"

    ' F3 i F4 su vezani: prazna polja ih zaustavljaju na PRVOM pravilu svog
    ' dokumenta -- a koje je to pravilo, dokazuje do kog modula je poziv stigao.
    Set p = PoljaEkrana(modScrDokumenti.modeKey("F3"))
    AssertEq modScrDokumenti.Scr_Save(p), Poruka("DOKUNOS_ERR_VOZAC"), _
             "zbirna ide u modDokUnos.ZbirnaValidiraj"
    AssertEq CStr(p("fokus")), "vozacID", "ekran vraca i polje na koje ide fokus"

    Set p = PoljaEkrana(modScrDokumenti.modeKey("F4"))
    AssertEq modScrDokumenti.Scr_Save(p), Poruka("DOKUNOS_ERR_KUPAC"), _
             "prijemnica ide u modDokUnos.PrijemnicaValidiraj"
    AssertEq CStr(p("fokus")), "kupacID", "ekran vraca i polje na koje ide fokus"

    ' F5: iznos veci od neisplacenog ostatka bloka -- pravilo koje postoji SAMO
    ' na putu isplate. Prazan broj dokumenta preskace proveru duplikata, pa ovaj
    ' test ne dira nijednu tabelu.
    Set p = PoljaEkrana(modScrDokumenti.modeKey("F5"))
    p("stanicaID") = FX_STANICA
    p("vrsta") = FX_VRSTA
    p("kooperantID") = FX_KOOPERANT
    p("partnerTip") = "KOOP"
    p("otkupID") = FX_BLOK
    p("blokTekst") = FX_BLOK
    p("novac") = OstatakFixtureBloka() + 1
    AssertEq modScrDokumenti.Scr_Save(p), _
             Poruka("NOVUNOS_ERR_VECI_OD_OSTATKA") & " " & _
             Format$(OstatakFixtureBloka(), "#,##0.00"), _
             "isplata ide u modNovacUnos.IsplataValidiraj"
    AssertEq CStr(p("fokus")), "novac", "ekran vraca i polje na koje ide fokus"

    ' F6: kupac je prvo pravilo uplate. Poruku deli sa prijemnicom, ali fokus ne
    ' -- prijemnica vraca "kupacID", uplata "partnerID".
    Set p = PoljaEkrana(modScrDokumenti.modeKey("F6"))
    AssertEq modScrDokumenti.Scr_Save(p), Poruka("DOKUNOS_ERR_KUPAC"), _
             "uplata ide u modNovacUnos.UplataValidiraj"
    AssertEq CStr(p("fokus")), "partnerID", "uplata trazi partnera, ne kupca prijemnice"

    ' F7: kolicina ambalaze je pravilo koje postoji samo na putu reversa.
    Set p = PoljaEkrana(modScrDokumenti.modeKey("F7"))
    p("stanicaID") = FX_STANICA
    p("vrsta") = FX_VRSTA
    AssertEq modScrDokumenti.Scr_Save(p), Poruka("NOVUNOS_ERR_KOL_AMB"), _
             "revers ide u modNovacUnos.ReversValidiraj"
    AssertEq CStr(p("fokus")), "kolAmb", "ekran vraca i polje na koje ide fokus"
End Sub

' F5 ISPLATA -- TIP NOVCA NIJE KOZMETIKA. Isti iznos knjizen pod pogresnim tipom
' ne razduzuje otkupni blok i ne vidi se u avansu otkupnog mesta, pa je izbor
' primaoca / bloka / prekidaca "isplata iz" jedino sto ovaj rezim odlucuje.
' Pada ako neka od cetiri grane nestane ili se zameni.
Private Sub T_IsplataValidiraj_TipNovcaPoIzboru()
    Dim p As Object, fokus As String, res As String
    Dim saldo As Double, prevVal As String
    Dim resStrogo As String, resLabavo As String

    ' (1) redosled: bez otkupnog mesta se ne ide dalje ni sa svim ostalim
    Set p = IsplataUnosKojiProlazi()
    p("stanicaID") = ""
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), Poruka("OTKUNOS_ERR_OM"), _
             "otkupno mesto je prvo pravilo isplate"
    AssertEq fokus, "stanicaID", "fokus ide na otkupno mesto"

    ' (2) iznos mora postojati -- rezim bez robe nema sta drugo da knjizi
    Set p = IsplataUnosKojiProlazi()
    p("novac") = 0
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), Poruka("NOVUNOS_ERR_NOVAC"), _
             "isplata bez iznosa se ne knjizi"

    ' (3) kooperant BEZ izabranog bloka -- to nije razduzenje nego avans.
    ' Cisti se I tekst polja: prazan ID uz nepraznu tekst znaci "ukucano a
    ' nerazreseno" i tada kapija namerno staje (T_NerazresenIzbor...).
    Set p = IsplataUnosKojiProlazi()
    p("otkupID") = ""
    p("blokTekst") = ""
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), "", "avans kooperantu prolazi"
    AssertEq CStr(p("tipNovca")), NOV_VIRMAN_AVANS_KOOP, _
             "bez bloka isplata kooperantu je avans, ne razduzenje"

    ' (4) kooperant SA blokom, virmanom -- razduzenje bloka
    Set p = IsplataUnosKojiProlazi()
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), "", "isplata po bloku prolazi"
    AssertEq CStr(p("tipNovca")), NOV_VIRMAN_FIRMA_KOOP, _
             "uz blok bez prekidaca isplata je virman firme"

    ' (5) iznos preko neisplacenog ostatka bloka -- blokada. Ostatak se cita iz
    ' PODATAKA; prosledjeni snimak (999999) se namerno ne poklapa sa njim, pa
    ' bi povratak na "veruj ekranu" ovu tvrdnju odmah oborio.
    Set p = IsplataUnosKojiProlazi()
    p("novac") = OstatakFixtureBloka() + 1
    res = modNovacUnos.IsplataValidiraj(p, fokus)
    AssertEq res, Poruka("NOVUNOS_ERR_VECI_OD_OSTATKA") & " " & _
             Format$(OstatakFixtureBloka(), "#,##0.00"), _
             "preko TRENUTNOG ostatka bloka se ne isplacuje"

    ' (6) "kes iz OM avansa" preko raspolozivog salda -- blokada. Saldo se cita
    ' iz istog read-modela koji koristi i pravilo, pa tvrdnja ne zavisi od toga
    ' koliko ga u fixture-u ima.
    saldo = GetOMAvansSaldo(FX_STANICA)
    Set p = IsplataUnosKojiProlazi()
    p("izAvansa") = True
    p("otkupOstatak") = saldo + 1000
    p("novac") = saldo + 1
    res = modNovacUnos.IsplataValidiraj(p, fokus)
    AssertEq res, Poruka("DOK_MSG_NEDOVOLJNO_AVANSA_RASPOLOZIVO") & " " & _
             Format$(saldo, "#,##0.00") & " RSD", _
             "iz OM avansa se ne isplacuje vise nego sto ga ima"

    ' (7) primalac je otkupno mesto -- kooperanta nema, pa nema ni bloka:
    ' red bez kooperanta sa OtkupID-em bio bi razduzenje nicijeg bloka.
    Set p = IsplataUnosKojiProlazi()
    p("partnerTip") = "OM"
    p("partnerID") = "OM-DRUGO"
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), "", "isplata otkupnom mestu prolazi"
    AssertEq CStr(p("tipNovca")), NOV_KES_FIRMA_OTKUPAC, _
             "primalac otkupno mesto -> kes firma-otkupac"
    AssertEq CStr(p("otkupID")), "", "uz primaoca otkupno mesto blok se odbacuje"
    AssertEq CStr(p("stanicaID")), "OM-DRUGO", _
             "izabrano otkupno mesto JESTE entitet novca"

    ' (8) BROJ DOKUMENTA JE POSLEDNJA PROVERA I VISI O VALIDACIJA_UNOSA -- isto
    ' kao u legacy. Isti unos mora da padne uz ukljucenu i prodje uz iskljucenu
    ' validaciju; bez oba smera se ne vidi da li kapija uopste meri taj flag.
    prevVal = GetConfigValue(CFG_VALIDACIJA_UNOSA)

    SetConfigValue CFG_VALIDACIJA_UNOSA, "DA"
    Set p = IsplataUnosKojiProlazi()
    p("brDok") = ""
    resStrogo = modNovacUnos.IsplataValidiraj(p, fokus)

    SetConfigValue CFG_VALIDACIJA_UNOSA, "NE"
    Set p = IsplataUnosKojiProlazi()
    p("brDok") = ""
    resLabavo = modNovacUnos.IsplataValidiraj(p, fokus)

    ' Podesavanje se vraca PRE tvrdnji: pala tvrdnja ne sme da ostavi iskljucenu
    ' validaciju ostatku suite-a nad istom sveskom.
    SetConfigValue CFG_VALIDACIJA_UNOSA, prevVal

    AssertEq resStrogo, Poruka("OTKUI_ERR_BROJ"), _
             "uz VALIDACIJA_UNOSA broj dokumenta je obavezan"
    AssertEq resLabavo, "", "bez VALIDACIJA_UNOSA isplata sme i bez broja"
End Sub

' F6 UPLATA -- IZABRANA FAKTURA ODLUCUJE STA JE RED. Bez nje uplata ne zatvara
' nijednu fakturu (UpdateFakturaStatus se ne pokrece), pa je razlika izmedju
' "uplata po fakturi" i "avans kupca" jedina odluka ovog rezima.
Private Sub T_UplataValidiraj_FakturaOdlucujeTip()
    Dim p As Object, fokus As String

    ' Kupac je prvo pravilo -- SaveKupciIzlaz_TX ga i sam trazi. Prazni se i
    ' tekst: ukucano ime bez izbora ima svoju, precizniju poruku.
    Set p = UplataUnosKojiProlazi()
    p("partnerID") = ""
    p("partnerTekst") = ""
    AssertEq modNovacUnos.UplataValidiraj(p, fokus), Poruka("DOKUNOS_ERR_KUPAC"), _
             "kupac je prvo pravilo uplate"
    AssertEq fokus, "partnerID", "fokus ide na partnera"

    ' bez fakture -> avans kupca (prazno je i polje, ne samo ID)
    Set p = UplataUnosKojiProlazi()
    p("fakturaID") = ""
    p("fakturaTekst") = ""
    AssertEq modNovacUnos.UplataValidiraj(p, fokus), "", "uplata bez fakture prolazi"
    AssertEq CStr(p("tipNovca")), NOV_KUPCI_AVANS, "bez fakture uplata je avans kupca"
    AssertEq CStr(p("napomena")), Poruka("NOVUNOS_NAP_AVANS_KUP"), _
             "napomena avansa ne pominje fakturu"

    ' sa fakturom -> uplata po fakturi, i napomena nosi njen broj
    Set p = UplataUnosKojiProlazi()
    AssertEq modNovacUnos.UplataValidiraj(p, fokus), "", "uplata po fakturi prolazi"
    AssertEq CStr(p("tipNovca")), NOV_KUPCI_UPLATA, "uz fakturu uplata zatvara fakturu"
    AssertEq CStr(p("napomena")), Poruka("NOVUNOS_NAP_FAKTURA") & " FAK-TEST-1", _
             "napomena nosi broj fakture"

    ' preko preostalog iznosa fakture -- blokada, po TRENUTNOM stanju
    Set p = UplataUnosKojiProlazi()
    p("novac") = OstatakFixtureFakture() + 1
    AssertEq modNovacUnos.UplataValidiraj(p, fokus), _
             Poruka("NOVUNOS_ERR_VECI_OD_FAKTURE") & " " & _
             Format$(OstatakFixtureFakture(), "#,##0.00"), _
             "preko TRENUTNOG preostalog iznosa fakture se ne uplacuje"

    ' faktura drugog kupca -- vlasnistvo se proverava u kapiji, ne u ekranu
    Set p = UplataUnosKojiProlazi()
    p("partnerID") = "KUP-DRUGI"
    AssertEq modNovacUnos.UplataValidiraj(p, fokus), _
             Poruka("NOVAC_ERR_FAK_TUDJ_KUPAC") & " " & FX_FAKTURA, _
             "uplata se ne vezuje za fakturu drugog kupca"

    ' nepostojeca faktura -- ID koji nije iz liste ne sme da prodje kao avans
    Set p = UplataUnosKojiProlazi()
    p("fakturaID") = "FAK-NEPOSTOJI"
    p("fakturaTekst") = "FAK-NEPOSTOJI"
    AssertEq modNovacUnos.UplataValidiraj(p, fokus), _
             Poruka("NOVAC_ERR_FAK_NEMA") & " FAK-NEPOSTOJI", _
             "nepostojeca faktura se odbija, ne knjizi kao avans"
End Sub

' KAPIJA VLASNISTVA I TRENUTNOG OSTATKA (modNovac.IsplataBlokProblem).
' Ovo je pravilo koje UI ne moze da odbrani: iznos je proveren nad snimkom iz
' trenutka kad je lista punjena, a izmedju punjenja i potvrde stanje se moze
' promeniti. Zato kapija cita podatke SADA, i zato je istu podize i writer.
' Pada ako se bilo koja od cetiri provere ukloni ili ako se vrati oslanjanje na
' vrednost koju je poslao ekran.
Private Sub T_IsplataBlokGuard_VlasnistvoITrenutniOstatak()
    Dim p As Object, fokus As String

    ' prazan blok NIJE greska -- to je avans kooperantu
    AssertEq modNovac.IsplataBlokProblem("", FX_KOOPERANT, FX_STANICA, 100), "", _
             "bez izabranog bloka kapija propusta (avans)"

    ' blok koji ne postoji
    AssertEq modNovac.IsplataBlokProblem("OTK-NEPOSTOJI", FX_KOOPERANT, FX_STANICA, 100), _
             Poruka("NOVAC_ERR_BLOK_NEMA") & " OTK-NEPOSTOJI", _
             "nepostojeci blok se odbija"

    ' blok drugog kooperanta
    AssertEq modNovac.IsplataBlokProblem(FX_BLOK_TUDJ, FX_KOOPERANT, FX_STANICA, 100), _
             Poruka("NOVAC_ERR_BLOK_TUDJ_KOOP") & " " & FX_BLOK_TUDJ, _
             "blok drugog kooperanta se odbija"

    ' blok sa drugog otkupnog mesta -- red novca se knjizi na aktivno OM, pa bi
    ' ovo razduzilo jedno mesto a teretilo drugo
    AssertEq modNovac.IsplataBlokProblem(FX_BLOK, FX_KOOPERANT, FX_STANICA2, 100), _
             Poruka("NOVAC_ERR_BLOK_TUDJ_OM") & " " & FX_BLOK, _
             "blok sa drugog otkupnog mesta se odbija"

    ' ispravna kombinacija prolazi
    AssertEq modNovac.IsplataBlokProblem(FX_BLOK, FX_KOOPERANT, FX_STANICA, 100), "", _
             "sopstveni blok na sopstvenom OM prolazi"

    ' iznos preko TRENUTNOG ostatka
    AssertEq modNovac.IsplataBlokProblem(FX_BLOK, FX_KOOPERANT, FX_STANICA, _
                                         OstatakFixtureBloka() + 1), _
             Poruka("NOVUNOS_ERR_VECI_OD_OSTATKA") & " " & _
             Format$(OstatakFixtureBloka(), "#,##0.00"), _
             "iznos preko trenutnog ostatka se odbija"

    ' Ista kapija kroz put unosa: cross-OM kombinacija mora da padne i kad ekran
    ' posalje savrsen snimak. Ovo je tvrdnja koju stari testovi nisu imali.
    Set p = IsplataUnosKojiProlazi()
    p("stanicaID") = FX_STANICA2
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), _
             Poruka("NOVAC_ERR_BLOK_TUDJ_OM") & " " & FX_BLOK, _
             "IsplataValidiraj zove kapiju, ne veruje snimku ekrana"
End Sub

' WRITER SE BRANI SAM. Svi ostali testovi voze put unosa (ekran -> modul), pa
' bi prosli i kad bi kapija postojala SAMO u modulu. Ovaj zove writer direktno,
' zaobilazeci ceo UI sloj -- kao sto ga zove legacy frmDokumenta ili bilo koji
' drugi pozivalac. Kombinacija je nemoguca (blok sa FX_STANICA, kontekst
' FX_STANICA2), pa upis mora da padne i NISTA ne sme da ostane u tabelama:
' guard puca posle BeginTx, a EH grana radi RollbackTx.
Private Sub T_WriterGuard_OdbijaTudjBlok()
    Dim ok As Boolean, uplataPre As Double, uplataPosle As Double

    uplataPre = GetUplataForOtkup(FX_BLOK)

    ok = SaveOMUlaz_TX(datum:=Date, _
                       brojDok:=FX_BROJ_NOVAC & "-W", _
                       stanicaNaziv:=FX_STANICA2, _
                       stanicaID:=FX_STANICA2, _
                       vozacID:="", _
                       tipAmb:="", _
                       kolAmb:=0, _
                       vrstaVoca:=FX_VRSTA, _
                       novac:=100, _
                       kooperantID:=FX_KOOPERANT, _
                       primalacDisplay:=FX_KOOPERANT, _
                       otkupID:=FX_BLOK, _
                       tipNovca:=NOV_VIRMAN_FIRMA_KOOP, _
                       koopSmer:="")

    uplataPosle = GetUplataForOtkup(FX_BLOK)

    AssertEq ok, False, "writer odbija blok sa drugog otkupnog mesta i bez UI provere"
    AssertEq uplataPosle, uplataPre, "odbijen upis ne ostavlja red u tblNovac"
End Sub

' POTPUNO PLACENA FAKTURA NE SME DA PRIMI JOS JEDNU UPLATU.
'
' Kapija je ranije glasila "If preostalo > 0 And iznos > preostalo", i to je
' bila rupa u samom mehanizmu koji je uveden da spreci zastarelo stanje:
'
'   faktura 10.000, operater otvorio ekran dok je preostalo 500,
'   u medjuvremenu je faktura zatvorena -> preostalo = 0
'   -> "0 > 0" je False -> kapija cuti -> jos jedna uplata prolazi.
'
' Test to vrti kroz PRAVI writer: prvo plati fakturu u celosti, pa pokusa jos
' jednu uplatu. Uslov "faktura bez iznosa ne blokira" je i dalje tu i proverava
' se zasebno - to je razlog zbog koga je provera uopste bila uslovna.
Private Sub T_UplataGuard_VecPlacenaFaktura()
    Dim ok As Boolean, pre As Double, posle As Double

    ' Preduslov: faktura ima iznos i nije placena.
    AssertEq FakturaIznos(FX_FAKTURA), FX_FAKTURA_IZNOS, "preduslov: faktura ima iznos"
    AssertEq modNovac.UplataFakturaProblem(FX_FAKTURA, FX_KUPAC, 1), "", _
             "dok ima ostatka, uplata prolazi kapiju"

    ' Plati je u CELOSTI, kroz pravi writer.
    ok = SaveKupciIzlaz_TX(datum:=Date, brojDok:=FX_BROJ_NOVAC & "-FULL", _
                           kupacNaziv:=FX_KUPAC, kupacID:=FX_KUPAC, vozacID:="", _
                           tipAmb:="", kolAmb:=0, vrstaVoca:=FX_VRSTA, _
                           novac:=FX_FAKTURA_IZNOS, fakturaID:=FX_FAKTURA, _
                           napomena:="test: puna uplata", tipNovca:=NOV_KUPCI_UPLATA)
    AssertEq ok, True, "puna uplata je proknjizena"
    AssertEq GetUplataForFaktura(FX_FAKTURA), FX_FAKTURA_IZNOS, "faktura je zatvorena"

    ' Sada je preostalo TACNO nula - stara kapija je bas tu cutala.
    AssertEq (Len(modNovac.UplataFakturaProblem(FX_FAKTURA, FX_KUPAC, 1)) > 0), True, _
             "vec placena faktura se odbija (preostalo = 0)"

    ' I writer mora da odbije, bez ijedne UI provere.
    pre = GetUplataForFaktura(FX_FAKTURA)
    ok = SaveKupciIzlaz_TX(datum:=Date, brojDok:=FX_BROJ_NOVAC & "-VISAK", _
                           kupacNaziv:=FX_KUPAC, kupacID:=FX_KUPAC, vozacID:="", _
                           tipAmb:="", kolAmb:=0, vrstaVoca:=FX_VRSTA, _
                           novac:=1, fakturaID:=FX_FAKTURA, _
                           napomena:="test: uplata preko pune", tipNovca:=NOV_KUPCI_UPLATA)
    posle = GetUplataForFaktura(FX_FAKTURA)

    AssertEq ok, False, "writer odbija uplatu na vec placenu fakturu"
    AssertEq posle, pre, "odbijen upis ne ostavlja red u tblNovac"

    ' PREPLACENA faktura (preostalo < 0) mora da se ponasa isto - ranije je i
    ' negativan ostatak prolazio kroz "preostalo > 0".
    AssertEq (Len(modNovac.UplataFakturaProblem(FX_FAKTURA, FX_KUPAC, 0.01)) > 0), True, _
             "ni najmanji iznos ne prolazi na zatvorenu fakturu"

    ' Faktura BEZ iznosa i dalje ne blokira - to pravilo se ne gubi uz popravku.
    AssertEq modNovac.UplataFakturaProblem(FX_FAKTURA_BEZ_IZNOSA, FX_KUPAC, 5000), "", _
             "faktura bez evidentiranog iznosa ne blokira uplatu"
End Sub

' ISPLATA IZ OM AVANSA NE SME DA PREDJE SALDO - I TO PROVERAVA WRITER.
'
' Blok i faktura su vec bili zasticeni u writer-u; avans je ostajao samo na UI
' sloju (modNovacUnos.IsplataValidiraj). Writer je time bio poslednja linija za
' dve od tri stvari koje isti dokument moze da prekoraci.
'
' Test je DIFERENCIJALAN: ista suma, isti prazan saldo, dva tipa novca. Samo
' kes isplata kooperantu trosi OM avans (GetOMAvansSaldo je i racuna kao
' odbitak), pa samo ona sme da bude odbijena - inace bi kapija bila obicna
' blokada svake isplate, a test to ne bi razlikovao.
Private Sub T_WriterGuard_AvansSaldoOM()
    Dim ok As Boolean, pre As Long

    AssertEq GetOMAvansSaldo(FX_STANICA), 0, _
             "preduslov: otkupno mesto nema avans salda"

    pre = NovacRedova()
    ok = SaveOMUlaz_TX(datum:=Date, brojDok:=FX_BROJ_NOVAC & "-AV", _
                       stanicaNaziv:=FX_STANICA, stanicaID:=FX_STANICA, _
                       vozacID:="", tipAmb:="", kolAmb:=0, vrstaVoca:=FX_VRSTA, _
                       novac:=100, kooperantID:=FX_KOOPERANT, _
                       primalacDisplay:=FX_KOOPERANT, otkupID:="", _
                       tipNovca:=NOV_KES_OTKUPAC_KOOP, koopSmer:="")
    AssertEq ok, False, "writer odbija kes isplatu preko avans salda OM"
    AssertEq NovacRedova(), pre, "odbijen upis ne ostavlja red u tblNovac"

    ' KONTROLA: virman firme NE trosi OM avans, pa isti iznos mora da prodje.
    ' Bez ove grane test ne bi razlikovao ciljanu kapiju od opste blokade.
    ok = SaveOMUlaz_TX(datum:=Date, brojDok:=FX_BROJ_NOVAC & "-VIR", _
                       stanicaNaziv:=FX_STANICA, stanicaID:=FX_STANICA, _
                       vozacID:="", tipAmb:="", kolAmb:=0, vrstaVoca:=FX_VRSTA, _
                       novac:=100, kooperantID:=FX_KOOPERANT, _
                       primalacDisplay:=FX_KOOPERANT, otkupID:="", _
                       tipNovca:=NOV_VIRMAN_FIRMA_KOOP, koopSmer:="")
    AssertEq ok, True, "virman firme ne trosi OM avans i prolazi"
    AssertEq NovacRedova(), pre + 1, "prosao upis JESTE ostavio red"
End Sub

' NumVal se NE koristi: postoji u modOtkupBlok i modScrDokumenti, ali je u oba
' Private. vba_check to ne prijavljuje (ime jeste definisano), pa se videlo tek
' kao "Cannot run the macro" - ceo projekat se nije kompajlirao.
Private Function FakturaIznos(ByVal fakturaID As String) As Double
    Dim v As Variant
    On Error Resume Next
    v = LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS)
    If IsNumeric(v) Then FakturaIznos = CDbl(v)
End Function

Private Function NovacRedova() As Long
    Dim d As Variant
    d = GetTableData(TBL_NOVAC)
    If IsEmpty(d) Then Exit Function
    NovacRedova = UBound(d, 1)
End Function

' F8 CITA TABELU IZABRANOG TIPA, ne uvek tblOtpremnica.
'
' Do v6-ui-118 je "STORNO" bio tih sinonim za "OTPREMNICA" u ModeTable i u
' desetak Col* funkcija, pa je storno centar mogao da pokaze samo otpremnice.
' Pada ako se EffKey ukloni ili ako se neki tip izgubi iz TabelaTipa: tada
' F8 opet svira po jednoj tabeli, a mreza tiho pokazuje pogresne dokumente
' pod pravim naslovom -- greska koju operater ne moze da vidi.
'
' Kolone se proveravaju ZAJEDNO sa tabelom: tabela bez odgovarajucih kolona
' daje praznu mrezu, sto izgleda kao "nema dokumenata".
Private Sub T_F8_TipBiraTabeluIKolone()
    Dim tipovi As Variant, tabele As Variant, i As Long, cols As Variant

    tipovi = Array(STIP_OTKUP, STIP_OTPREMNICA, STIP_ZBIRNA, STIP_PRIJEMNICA, _
                   STIP_ISPLATE, STIP_UPLATE, STIP_REVERSI, STIP_FAKTURA, STIP_IZVOD)
    tabele = Array(TBL_OTKUP, TBL_OTPREMNICA, TBL_ZBIRNA, TBL_PRIJEMNICA, _
                   TBL_NOVAC, TBL_NOVAC, TBL_AMBALAZA, TBL_FAKTURE, TBL_BANKA_IMPORT)

    For i = 0 To UBound(tipovi)
        modScrDokumenti.Scr_StornoTipTestSet CStr(tipovi(i))
        AssertEq modScrDokumenti.ModeTable("F8"), CStr(tabele(i)), _
                 "F8 / " & CStr(tipovi(i)) & " cita svoju tabelu"
        AssertEq modScrDokumenti.EffKey("STORNO"), CStr(tipovi(i)), _
                 "F8 / " & CStr(tipovi(i)) & " razresava kljuc rezima u kljuc tipa"
        cols = modScrDokumenti.GridCols("STORNO")
        AssertEq (IsArray(cols)), True, _
                 "F8 / " & CStr(tipovi(i)) & " ima opis kolona"
        AssertEq (UBound(cols) >= 3), True, _
                 "F8 / " & CStr(tipovi(i)) & " ima bar cetiri kolone"
    Next i

    ' Broj zbirne postoji samo tamo gde ga dokument NOSI. Dok je F8 bio
    ' otpremnica, cip "Bez zbirne" je bio ukljucen i nad novcem, gde tblNovac
    ' tu kolonu nema.
    modScrDokumenti.Scr_StornoTipTestSet STIP_ISPLATE
    AssertEq modScrDokumenti.ModeHasZbirna("F8"), False, _
             "novac u F8 nema pojam zbirne"
    modScrDokumenti.Scr_StornoTipTestSet STIP_OTKUP
    AssertEq modScrDokumenti.ModeHasZbirna("F8"), True, _
             "otkupni list u F8 ima pojam zbirne"

    modScrDokumenti.Scr_StornoTipTestSet STIP_OTKUP
End Sub

' KAPIJA STOJI PRE UPISA, I VRACA RAZLOG.
'
' Ekran pita StornoRazlog pre nego sto uopste ponudi potvrdu -- da operater
' vidi zasto se nesto ne moze stornirati, umesto tihog neuspeha posle "Da".
' Pada ako se neka grana Select Case-a izgubi (tada nepostojeci dokument
' prolazi kapiju i ide pravo u Storno*_TX) ili ako se izgubi zahtev za
' smerom reversa (cetiri smera dele isti brojevni niz, pa bi bez smera
' StornoOMKoopByBrDok_TX gadjao pogresne redove).
'
' Test NE upisuje nista: svi brojevi su izmisljeni, pa nijedna grana ne
' stigne do transakcije.
Private Sub T_StornoDok_KapijePreUpisa()
    Dim tipovi As Variant, i As Long, r As String
    Const NEMA As String = "NE-POSTOJI-9999"

    ' 1) Nepostojeci dokument -- svaki tip mora da vrati razlog.
    tipovi = Array(STIP_OTKUP, STIP_OTPREMNICA, STIP_ZBIRNA, STIP_PRIJEMNICA, _
                   STIP_FAKTURA, STIP_ISPLATE, STIP_UPLATE)
    For i = 0 To UBound(tipovi)
        r = modStornoDok.StornoRazlog(CStr(tipovi(i)), NEMA, "")
        AssertEq (Len(r) > 0), True, _
                 "kapija zaustavlja nepostojeci dokument, tip " & CStr(tipovi(i))
    Next i

    ' 2) Prazan broj (red bez broja) -- ne sme da prodje ni za jedan tip.
    AssertEq (Len(modStornoDok.StornoRazlog(STIP_OTKUP, "", "")) > 0), True, _
             "kapija zaustavlja prazan broj"

    ' 3) Revers bez smera. Broj postoji ili ne -- svejedno: bez smera se ne
    '    zna koji je od cetiri dokumenta, pa se ne sme ni pokusati.
    AssertEq modStornoDok.StornoRazlog(STIP_REVERSI, NEMA, ""), _
             Poruka("STORNO_ERR_NEMA_SMERA"), _
             "revers bez smera se odbija PRE trazenja dokumenta"

    ' 4) Nepoznat tip ne sme tiho da ne uradi nista.
    AssertEq (Len(modStornoDok.StornoRazlog("NEPOSTOJECI_TIP", NEMA, "")) > 0), True, _
             "nepoznat tip vraca razlog"

    ' 5) Ime tipa je ono sto operater vidi u potvrdi -- prazno ime bi dalo
    '    "Stornirati  12/010826?" i operater ne bi znao STA stornira.
    For i = 0 To UBound(tipovi)
        AssertEq (Len(modStornoDok.TipNaziv(CStr(tipovi(i)), "")) > 0), True, _
                 "tip ima ime za potvrdu: " & CStr(tipovi(i))
    Next i
    AssertEq (Len(modStornoDok.TipNaziv(STIP_REVERSI, DOK_TIP_OM_IZLAZ_KOOP)) > 0), True, _
             "revers ima ime po SMERU"
End Sub

' UKUCAN A NERAZRESEN IZBOR NIJE "NIJE IZABRANO".
' Combo dopusta kucanje, a ID stize iz skrivene kolone koja postoji samo uz
' stvarno izabranu stavku. Bez ove kapije operater vidi ime u polju, pritisne
' Sacuvaj, a dokument se knjizi na nekog drugog (ili kao avans).
Private Sub T_NerazresenIzbor_NeProlaziKaoPrazno()
    Dim p As Object, fokus As String

    ' F5 partner: ime u polju, ID prazan
    Set p = IsplataUnosKojiProlazi()
    p("partnerID") = ""
    p("otkupID") = ""
    p("blokTekst") = ""
    p("partnerTekst") = "Petar Petrovic"
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), _
             Poruka("NOVUNOS_ERR_PARTNER_NEIZABRAN"), _
             "ukucan partner bez izbora ne prolazi kao isplata otkupnom mestu"
    AssertEq fokus, "partnerID", "fokus ide na partnera"

    ' F5 blok: tekst u polju, ID prazan -> ne sme da postane avans
    Set p = IsplataUnosKojiProlazi()
    p("otkupID") = ""
    p("blokTekst") = "1/TEST"
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), _
             Poruka("NOVUNOS_ERR_BLOK_NEIZABRAN"), _
             "ukucan blok bez izbora ne prolazi kao avans"

    ' F6 faktura: tekst u polju, ID prazan -> ne sme da postane avans kupca
    Set p = UplataUnosKojiProlazi()
    p("fakturaID") = ""
    p("fakturaTekst") = "12/2026"
    AssertEq modNovacUnos.UplataValidiraj(p, fokus), _
             Poruka("NOVUNOS_ERR_FAKTURA_NEIZABRANA"), _
             "ukucana faktura bez izbora ne prolazi kao avans kupca"

    ' F7 partner uz kooperantski smer
    Set p = ReversUnosKojiProlazi()
    p("partnerID") = ""
    p("partnerTip") = ""
    p("partnerTekst") = "Petar Petrovic"
    AssertEq modNovacUnos.ReversValidiraj(p, fokus), _
             Poruka("NOVUNOS_ERR_PARTNER_NEIZABRAN"), _
             "ukucan partner bez izbora ne prolazi ni u reversu"

    ' KONTROLA (obrnut smer): PRAZNO polje i dalje znaci "nije izabrano" i
    ' prolazi -- kapija sme da hvata samo tekst bez ID-a.
    Set p = IsplataUnosKojiProlazi()
    p("partnerID") = ""
    p("partnerTip") = ""
    p("partnerTekst") = ""
    p("otkupID") = ""
    p("blokTekst") = ""
    AssertEq modNovacUnos.IsplataValidiraj(p, fokus), "", _
             "prazan partner i dalje znaci isplata otkupnom mestu"
    AssertEq CStr(p("tipNovca")), NOV_KES_FIRMA_OTKUPAC, _
             "bez partnera tip ostaje kes firma-otkupac"
End Sub

' F7 REVERS -- SMER JE OBAVEZAN I NIJE PRIKAZ. Bez smera je SaveOMUlaz_TX ranije
' tiho knjizio "OM prima od vozaca", pa je prazan smer davao pogresan red bez
' ijedne poruke. Uz to: kooperantski smerovi traze bas kooperanta (kupac u
' reversu ne postoji ni u legacy), a firma<->OM smerovi traze vozaca UVEK.
Private Sub T_ReversValidiraj_SmerJeObavezan()
    Dim p As Object, fokus As String

    Set p = ReversUnosKojiProlazi()
    p("smerRev") = 0
    AssertEq modNovacUnos.ReversValidiraj(p, fokus), Poruka("NOVUNOS_ERR_SMER"), _
             "revers bez izabranog smera se ne knjizi"
    AssertEq fokus, "smerRev", "fokus ide na smer"

    ' kolicina i tip ambalaze idu PRE smera -- revers bez ambalaze nema smisla
    Set p = ReversUnosKojiProlazi()
    p("kolAmb") = 0
    AssertEq modNovacUnos.ReversValidiraj(p, fokus), Poruka("NOVUNOS_ERR_KOL_AMB"), _
             "revers bez kolicine se ne knjizi"

    Set p = ReversUnosKojiProlazi()
    p("tipAmb") = ""
    AssertEq modNovacUnos.ReversValidiraj(p, fokus), Poruka("DOK_MSG_IZABERITE_TIP_AMBALAZE"), _
             "revers bez tipa ambalaze se ne knjizi"

    ' "Izdato koop." sa kupcem kao partnerom -- blokada, ne tiha knjizba
    Set p = ReversUnosKojiProlazi()
    p("partnerTip") = "KUP"
    AssertEq modNovacUnos.ReversValidiraj(p, fokus), Poruka("NOVUNOS_ERR_SMER_KOOP"), _
             "kooperantski smer ne prima kupca"
    AssertEq fokus, "partnerID", "fokus ide na partnera"

    ' Firma <-> OM ide preko vozaca -- vozac je obavezan i BEZ stroge validacije
    Set p = ReversUnosKojiProlazi()
    p("smerRev") = modNovacUnos.SMER_REV_IZD_OM
    p("vozacID") = ""
    AssertEq modNovacUnos.ReversValidiraj(p, fokus), Poruka("NOVUNOS_ERR_VOZAC_OM"), _
             "revers firma-OM bez vozaca se ne knjizi"

    ' Prevod segmenta u ono sto core poznaje. Pogresan prevod ne bi pao na
    ' proveri nego bi proknjizio suprotan smer.
    AssertEq modNovacUnos.SmerRevKljuc(modNovacUnos.SMER_REV_IZD_KOOP), "IZDAVANJE", _
             "segment 1 = izdavanje kooperantu"
    AssertEq modNovacUnos.SmerRevKljuc(modNovacUnos.SMER_REV_PRI_KOOP), "PRIJEM", _
             "segment 2 = prijem od kooperanta"
    AssertEq modNovacUnos.SmerRevKljuc(modNovacUnos.SMER_REV_IZD_OM), "IZDATO_OM", _
             "segment 3 = izdato OM"
    AssertEq modNovacUnos.SmerRevKljuc(modNovacUnos.SMER_REV_PRI_OM), "PRIJEM_OD_OM", _
             "segment 4 = prijem od OM"
    AssertEq modNovacUnos.SmerRevKljuc(0), "", _
             "neizabran smer nema prevod -- core guard puca umesto da knjizi"
End Sub

' Isplata koja prolazi sve provere: kooperant sa izabranim otkupnim blokom.
' Testovi je onda kvare po jednom polju.
Private Function IsplataUnosKojiProlazi() As Object
    Dim p As Object
    Set p = modNovacUnos.NoviIsplataUnos()
    p("stanicaID") = FX_STANICA
    p("stanicaTekst") = FX_STANICA
    p("partnerID") = FX_KOOPERANT
    p("partnerTip") = "KOOP"
    p("partnerTekst") = FX_KOOPERANT
    p("vrsta") = FX_VRSTA
    ' Broj MORA biti popunjen: uz VALIDACIJA_UNOSA (podrazumevano ukljucena)
    ' prazan broj obara i inace ispravan unos. Vrednost je izmisljena bas da je
    ' provera duplikata ne nadje ni u tblAmbalaza ni u tblNovac.
    p("brDok") = FX_BROJ_NOVAC
    p("novac") = 500
    p("otkupID") = FX_BLOK
    p("blokTekst") = FX_BLOK
    ' Namerno LAZAN snimak ostatka. Modul ga vise ne koristi -- cita trenutno
    ' stanje kroz IsplataBlokProblem. Da se na njega vrati, testovi koji ovde
    ' salju nemoguce vrednosti bi to odmah pokazali.
    p("otkupOstatak") = 999999
    Set IsplataUnosKojiProlazi = p
End Function

' Trenutni neisplaceni ostatak fixture bloka -- racuna se isto kao u kapiji,
' pa tvrdnja ne zavisi od toga da li fixture ima uplate.
Private Function OstatakFixtureBloka() As Double
    OstatakFixtureBloka = FX_BLOK_VREDNOST - GetUplataForOtkup(FX_BLOK)
End Function

Private Function UplataUnosKojiProlazi() As Object
    Dim p As Object
    Set p = modNovacUnos.NoviUplataUnos()
    p("partnerID") = FX_KUPAC
    p("partnerTekst") = FX_KUPAC
    p("vrsta") = FX_VRSTA
    p("brDok") = FX_BROJ_NOVAC
    p("novac") = 500
    p("fakturaID") = FX_FAKTURA
    p("fakturaTekst") = FX_FAKTURA
    ' Lazan snimak, iz istog razloga kao kod bloka.
    p("fakturaOstatak") = 999999
    Set UplataUnosKojiProlazi = p
End Function

Private Function OstatakFixtureFakture() As Double
    OstatakFixtureFakture = FX_FAKTURA_IZNOS - GetUplataForFaktura(FX_FAKTURA)
End Function

' Revers kome fali samo ono sto test pokvari. Broj je POPUNJEN namerno: prazan
' bi na kraju provera okinuo SuggestNextBroj, koji ume da pita Google. Tu granu
' ne vozi nijedna tvrdnja u ispravnom kodu, ali je vozi sabotaza "revers-smer"
' -- a test koji zavisi od mreze nije test.
Private Function ReversUnosKojiProlazi() As Object
    Dim p As Object
    Set p = modNovacUnos.NoviReversUnos()
    p("stanicaID") = FX_STANICA
    p("stanicaTekst") = FX_STANICA
    p("partnerID") = FX_KOOPERANT
    p("partnerTip") = "KOOP"
    p("partnerTekst") = FX_KOOPERANT
    p("vozacID") = FX_VOZAC
    p("vrsta") = FX_VRSTA
    p("brDok") = FX_BROJ_NOVAC
    p("tipAmb") = FX_TIP_AMB
    p("kolAmb") = 20
    p("smerRev") = modNovacUnos.SMER_REV_IZD_KOOP
    Set ReversUnosKojiProlazi = p
End Function

' Zbirna koja se u SVEMU poklapa sa fixture otpremnicom OTP-TEST-1 (jedina koja
' nosi FX_ZBIRNA). Testovi je onda kvare po jednom polju.
Private Function ZbirnaUnosKojiSeSlaze() As Object
    Dim p As Object
    Set p = modDokUnos.NoviZbirnaUnos()
    p("vozacID") = FX_VOZAC
    p("kupacID") = FX_KUPAC
    p("brDok") = FX_ZBIRNA              ' u F3 broj dokumenta JESTE broj zbirne
    p("vrsta") = FX_VRSTA
    p("sorta") = FX_SORTA
    p("tipAmb") = FX_TIP_AMB
    p("kolicinaI") = FX_ZBIRNA_KG
    p("kolAmb") = FX_ZBIRNA_AMB
    Set ZbirnaUnosKojiSeSlaze = p
End Function

' Prijemnica koja prolazi sve provere: broj koji nije zauzet, postojeca zbirna,
' popunjena roba i cena. Testovi je kvare po jednom polju.
Private Function PrijemnicaUnosKojiProlazi() As Object
    Dim p As Object
    Set p = modDokUnos.NoviPrijemnicaUnos()
    p("kupacID") = FX_KUPAC
    p("vozacID") = FX_VOZAC
    p("brDok") = "PR-TEST-1"
    p("brojZbirne") = FX_ZBIRNA
    p("vrsta") = FX_VRSTA
    p("sorta") = FX_SORTA
    p("tipAmb") = FX_TIP_AMB
    p("kolicinaI") = 100
    p("cenaI") = 50
    p("kolAmb") = 10
    Set PrijemnicaUnosKojiProlazi = p
End Function

' Recnik kakav ljuska predaje ekranu (modOtkupUI.SkupiPolja), sa praznim
' vrednostima. Imena kljuceva su deo ugovora izmedju ljuske i ekrana, pa ih
' test navodi eksplicitno umesto da ih pozajmi iz ljuske.
Private Function PoljaEkrana(ByVal rezim As String) As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("rezim") = rezim
    p("datum") = Date
    p("stanicaID") = ""
    p("kooperantID") = ""
    p("partnerTekst") = ""
    p("vrsta") = ""
    p("sorta") = ""
    p("vozacID") = ""
    p("brDok") = ""
    p("brojZbirne") = ""
    p("parcelaID") = ""
    p("tipAmb") = ""
    p("kolicinaI") = 0#
    p("cenaI") = 0#
    p("kolAmb") = 0&
    p("kolAmbIzdata") = 0&
    p("dveKlase") = False
    p("kolicinaII") = 0#
    p("cenaII") = 0#
    p("kolAmbII") = 0&
    p("novac") = 0#
    ' gotovinski rezimi i reversi -- isti kljucevi koje salje modOtkupUI.SkupiPolja
    p("stanicaTekst") = ""
    p("partnerTip") = ""
    p("otkupID") = ""
    p("blokTekst") = ""
    p("otkupOstatak") = 0#
    p("izAvansa") = False
    p("fakturaID") = ""
    p("fakturaTekst") = ""
    p("fakturaOstatak") = 0#
    p("smerRev") = 0&
    p("stampajUvek") = False
    Set PoljaEkrana = p
End Function

' Novi UI bez prikaza. Gradnja se okida dodirom Controls.count, isto kao kod
' frmOtkup; .Show se NE zove -- GoFullScreen, raspored i punjenje mreze idu tek
' u UserForm_Activate, a nista od toga ovi testovi ne mere.
Private Function NewOtkupUIForm() As frmOtkupUI
    Dim f As frmOtkupUI
    Set f = New frmOtkupUI

    Dim ctlCount As Long
    ctlCount = f.Controls.count          ' bez ovoga se UserForm_Initialize ne okine

    ' UserForm_Initialize hvata pad gradnje i salje ga u OtkupUI_BuildFailed, pa
    ' greska NE stize ovamo. Bez ove provere bi svaka sledeca tvrdnja padala na
    ' "Could not find the specified object" -- pad na trazenju kontrole, a ne na
    ' ponasanju koje test meri.
    If ctlCount < 2 Then
        Err.Raise ERR_ASSERT, "modTest.NewOtkupUIForm", _
                  "frmOtkupUI nije izgradjen (kontrola: " & ctlCount & ")"
    End If

    Set NewOtkupUIForm = f
End Function

' Unload gasi formu (Terminate -> OtkupUI_FormClosed), a OtkupUI_Release pusta i
' ono sto ostaje na modulu (Btns, kes tabela, num-polja) -- inace sledeci test
' gradi ekran nad ostacima prethodnog. Aktivna otpremnica zivi u TRECEM modulu
' (modScrDokumenti) i nju OtkupUI_Release ne dira, pa se otpusta ovde.
Private Sub ReleaseOtkupUIForm(f As frmOtkupUI)
    Unload f
    modOtkupUI.OtkupUI_Release
    modScrDokumenti.Scr_OtpOtkazi
End Sub

' Polja novog UI-ja su ugnjezdena: zona -> okvir polja -> kontrola (ime + "T").
' Test se kroz to stablo krece SAM, ne kroz modOtkupUI.FldText/SetFld: rutina
' koja se testira ne sme da bude i merni instrument.
Private Function Polje(z As Object, ByVal grp As String) As String
    Polje = z.Controls(grp).Controls(grp & "T").text
End Function

Private Sub SetPolje(z As Object, ByVal grp As String, ByVal v As String)
    z.Controls(grp).Controls(grp & "T").text = v
End Sub

' Forma sa kontekstom otpremnice OTP-TEST-1 iz fixture-a, bez .Show.
Private Function NewOtkupForm() As frmOtkup
    Dim f As frmOtkup
    Set f = New frmOtkup

    Dim ctlCount As Long
    ctlCount = f.Controls.count          ' bez ovoga se UserForm_Initialize ne okine

    f.txtDatum.value = FX_DATUM
    f.txtBrojZbirne.value = FX_ZBIRNA
    f.txtBrojDokumenta.value = FX_BROJ_OTP
    f.cmbKooperant.value = FX_KOOPERANT

    Set NewOtkupForm = f
End Function

' ============================================================
' Assert-i
' ============================================================
Public Sub AssertEq(ByVal actual As Variant, ByVal expected As Variant, _
                    ByVal label As String)
    Dim a As String, e As String
    a = SafeStr(actual)
    e = SafeStr(expected)
    If a <> e Then
        Err.Raise ERR_ASSERT, "modTest.AssertEq", _
                  label & " -- ocekivano [" & e & "], dobijeno [" & a & "]"
    End If
End Sub

' Prazna kontrola ume da vrati Null umesto "", a CStr(Null) puca ("Invalid use
' of Null") -- test bi pao na toj gresci umesto na ponasanju koje meri.
Private Function SafeStr(ByVal v As Variant) As String
    If IsNull(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

' Snapshot hvata i polja koja niko nije trazio da se provere. Kad golden ne
' postoji, upisuje ga i PADA -- nov golden mora da prodje ljudski pregled pre
' nego sto postane merilo.
Public Sub AssertSnapshot(ByVal tekuci As String, ByVal imeGolden As String)
    Dim path As String
    path = GoldenDir() & imeGolden & ".txt"

    If Len(Dir$(path)) = 0 Then
        WriteTextFile path, tekuci
        Err.Raise ERR_GOLDEN, "modTest.AssertSnapshot", _
                  "Golden nije postojao -- upisan je (" & imeGolden & _
                  ".txt). Pregledaj ga i commit-uj, pa pokreni ponovo."
    End If

    Dim golden As String
    golden = ReadTextFile(path)
    If golden <> tekuci Then
        Err.Raise ERR_ASSERT, "modTest.AssertSnapshot", _
                  "snapshot " & imeGolden & " se razlikuje od golden-a -- " & _
                  FirstDiff(golden, tekuci)
    End If
End Sub

' ============================================================
' Pomocno
' ============================================================

' Sve kontrole forme kao sortirano "ime=vrednost", jedan par po liniji.
' Sortira se postojecim modArrayUtils.SortArray (nema novog sorta).
Public Function DumpKontrole(ByVal f As Object) As String
    Dim n As Long
    n = f.Controls.count
    If n = 0 Then
        DumpKontrole = ""
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 1)

    Dim ctl As Object
    Dim i As Long
    i = 0
    For Each ctl In f.Controls
        i = i + 1
        arr(i, 1) = AsciiEscape(ctl.name & "=" & ControlValue(ctl))
    Next ctl

    Dim sorted As Variant
    sorted = SortArray(arr, 1, True)

    Dim sb As String
    For i = 1 To n
        sb = sb & CStr(sorted(i, 1)) & vbLf
    Next i

    DumpKontrole = sb
End Function

' Sve van stampanog ASCII-ja ide kao \uXXXX. Bez ovoga je golden neupotrebljiv:
' VBA Print # pise u ANSI kodnu stranu, koja "Vrsta voca" sa ch ne moze da
' predstavi (cp1252) -- snimi se osakaceno, pa svako sledece poredjenje pada, a
' poruka o razlici izgleda kao da su stringovi isti jer se i ona gubi na istom
' mestu. Uz escape je golden cist ASCII, round-trip je tacan, a razlika citljiva.
Private Function AsciiEscape(ByVal s As String) As String
    Dim i As Long
    Dim ch As Long
    Dim out As String

    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Then ch = ch + 65536      ' AscW je Integer: > 32767 dolazi negativno
        If ch = 92 Then
            out = out & "\\"            ' inace bi putanja "C:\users" izgledala kao escape
        ElseIf ch >= 32 And ch <= 126 Then
            out = out & Chr$(ch)
        Else
            out = out & "\u" & Right$("000" & Hex$(ch), 4)
        End If
    Next i

    AsciiEscape = out
End Function

' Kontrole nemaju sve .Value (Label/Frame imaju Caption, neke nemaju nista).
Private Function ControlValue(ByVal ctl As Object) As String
    Dim s As String

    On Error Resume Next
    Err.Clear
    s = CStr(ctl.value)
    If Err.Number <> 0 Then
        Err.Clear
        s = CStr(ctl.caption)
        If Err.Number <> 0 Then
            Err.Clear
            s = "<n/a>"
        End If
    End If
    On Error GoTo 0

    ControlValue = s
End Function

Private Function FirstDiff(ByVal a As String, ByVal b As String) As String
    Dim la As Variant, lb As Variant
    la = Split(a, vbLf)
    lb = Split(b, vbLf)

    Dim n As Long
    n = UBound(la)
    If UBound(lb) < n Then n = UBound(lb)

    Dim i As Long
    For i = 0 To n
        If la(i) <> lb(i) Then
            FirstDiff = "prva razlika: golden [" & la(i) & "] vs tekuci [" & lb(i) & "]"
            Exit Function
        End If
    Next i

    FirstDiff = "razlicit broj linija: golden " & (UBound(la) + 1) & _
                ", tekuci " & (UBound(lb) + 1)
End Function

' Golden fajlovi zive pored sveske; run_vba.py ih kopira iz tests/golden pre
' rana i vraca posle, da nov golden zavrsi u repou na pregled.
Private Function GoldenDir() As String
    Dim d As String
    d = ThisWorkbook.path & Application.PathSeparator & "golden"
    If Len(Dir$(d, vbDirectory)) = 0 Then MkDir d
    GoldenDir = d & Application.PathSeparator
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, content;
    Close #fnum
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim raw As String
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Input As #fnum
    raw = Input$(LOF(fnum), fnum)
    Close #fnum

    ' CR se izbacuje: .gitattributes drzi golden na LF, ali klon sa drugim
    ' podesavanjem (ili rucno editovanje u Notepad-u) vrati CRLF, a tada golden
    ' vise nije jednak dump-u koji se spaja sa vbLf. Pravi CR u sadrzaju ne
    ' postoji -- AsciiEscape ga pretvara u \u000D.
    ReadTextFile = Replace$(raw, vbCr, "")
End Function

Private Sub AppendReport(ByVal testNm As String, ByVal status As String, _
                         ByVal detail As String)
    m_Report = m_Report & status & " " & testNm
    If Len(detail) > 0 Then m_Report = m_Report & " -- " & detail
    m_Report = m_Report & vbLf
End Sub

Private Sub WriteResultFile()
    Dim path As String
    path = ThisWorkbook.path & Application.PathSeparator & "last_run.txt"
    WriteTextFile path, "TESTS=" & m_Total & " FAIL=" & m_Failed & vbLf & m_Report
End Sub
