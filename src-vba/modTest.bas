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
' Stornirana zbirna -- postoji SAMO da bi lista ciljeva na ekranu Oporavak
' imala sta da izostavi (bez nje ta tvrdnja nema nad cim da padne).
Private Const FX_ZBIRNA_STORNO As String = "ZB-TEST-STORNO"
Private Const FX_ZBIRNA2 As String = "ZB-TEST-2"
' Zbirna u koju nijedan test ne upisuje -- kolizioni par aktivnih
' prijemnica pocinje na njoj, da upis iz drugog testa ne otvori dijalog.
Private Const FX_ZBIRNA_MIRNA As String = "ZB-TEST-4"
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
' Kolizija brojeva: PRJ-TEST-A (KUP-TEST-1) i PRJ-TEST-B (KUP-TEST-2) nose ISTI
' BrojPrijemnice. Tako je i u produkciji -- broj se racuna po kupcu.
Private Const FX_KUPAC2 As String = "KUP-TEST-2"
Private Const FX_PRIJ_BROJ As String = "1/150326"
Private Const FX_PRIJ_STORNO As String = "9/150326"
' Kolizioni par storniranih: isti broj, dva kupca, dve palete.
Private Const FX_PRIJ_KOLIZIJA As String = "8/150326"
' Kolizioni par AKTIVNIH prijemnica, svaka sa svojom paletom. Prevezivanje na
' zbirnu sme da dira samo svoj dokument -- i u tblPrijemnica I u tblPaletaStavka.
Private Const FX_PRIJ_ZBR_KOLIZIJA As String = "6/150326"
' Dva dokumenta istog broja i iste robe koja DELE fizicku paletu, i cilj
' druge vrste -- da prevezivanje uopste bude relabel.
Private Const FX_PRIJ_DELJENA As String = "5/150326"
Private Const FX_PRIJ_CILJ_V2 As String = "4/150326"
' Dva AKTIVNA reda tblNovac pod istim brojem -- avans raspodela.
Private Const FX_NOVAC_DUPLI As String = "NOV-DUPLI-1"
' Dve AKTIVNE prijemnice istog broja, za ispravku pod kolizijom.
Private Const FX_PRIJ_ISPRAVKA As String = "3/150326"
' Isti broj na dva otkupna mesta / dve stanice -- oba niza su scoped po stanici.
Private Const FX_OTKUP_KOLIZIJA As String = "7/150326"
Private Const FX_OTPREMNICA_KOLIZIJA As String = "8/TEST"
' Svez par zbirnih za kaskadu (test 38 potrosi ZB-TEST-DUPL).
Private Const FX_ZBIRNA_KASK As String = "ZB-TEST-KASK"
' Zatecen par BEZ generacije + zamena, za zavrsetak ispravke.
Private Const FX_OTPREMNICA_LEGACY As String = "6/TEST"
Private Const FX_OTPREMNICA_ZAMENA As String = "7/TEST"
Private Const FX_OTPREMNICA_STALE As String = "9/TEST"
Private Const FX_OTPREMNICA_STALE_NOVA As String = "10/TEST"
Private Const FX_ZBIRNA_STALE As String = "ZB-TEST-STL"
Private Const FX_PRIJEMNICA_STALE As String = "12/TEST"
Private Const FX_OTPREMNICA_BLOK As String = "18/TEST"
Private Const FX_PRIJEMNICA_OLD_U As String = "16/TEST"
Private Const FX_ZBIRNA_TGT As String = "ZB-TEST-TGT"
Private Const FX_ZBIRNA_OLDU As String = "ZB-TEST-OLDU"
Private Const FX_OTPREMNICA_OLD_U As String = "13/TEST"
Private Const FX_OTPREMNICA_NEW_T As String = "15/TEST"
' Dve zbirne ISTOG broja i ISTOG kupca, dva vozaca. Broj zbirne se generise po
' vozacu, pa su to dva dokumenta -- ciljna lista mora da ponudi oba.
Private Const FX_ZBIRNA_DUPL As String = "ZB-TEST-DUPL"
Private Const FX_VOZAC2 As String = "VOZ-TEST-2"
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
    RunOne 22
    RunOne 23
    RunOne 24
    RunOne 25
    RunOne 26
    RunOne 27
    RunOne 28
    RunOne 29
    RunOne 30
    RunOne 31
    RunOne 32
    RunOne 33
    RunOne 34
    RunOne 35
    RunOne 36
    RunOne 37
    RunOne 38
    RunOne 39
    RunOne 40
    RunOne 41
    RunOne 42
    RunOne 43
    RunOne 44
    RunOne 45
    RunOne 46
    RunOne 47
    RunOne 48
    RunOne 49
    RunOne 50
    RunOne 51
    RunOne 52
    RunOne 53
    RunOne 54
    RunOne 55
    RunOne 56
    RunOne 57
    RunOne 58
    RunOne 59

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
        Case 20: TestName = "T_Storno_TipBiraTabeluIKolone"
        Case 21: TestName = "T_StornoDok_KapijePreUpisa"
        Case 22: TestName = "T_PrefillIzStorniranog_CitaSvojuTabelu"
        Case 23: TestName = "T_FrameworkIspravke_SamoCetiriTipa"
        Case 24: TestName = "T_Prefill_PoIdentitetuNePoBroju"
        Case 25: TestName = "T_IspravkaDetekcija_FailClosed"
        Case 26: TestName = "T_Oporavak_UgovorIRadnje"
        Case 27: TestName = "T_Oporavak_CiljneListe"
        Case 28: TestName = "T_IspravkaPrijemnice_SkipIRelink"
        Case 29: TestName = "T_RelinkPoGeneraciji_NeDiraTudjDokument"
        Case 30: TestName = "T_PrevezivanjeNaZbirnu_PaletaIdePoIdentitetu"
        Case 31: TestName = "T_ZadataGeneracijaKojeNema_Staje"
        Case 32: TestName = "T_VerdiktPoIdentitetu_RelabelSeNePreskace"
        Case 33: TestName = "T_DeljenaPaleta_SuStanarPoIdentitetu"
        Case 34: TestName = "T_IstiBrojRazliciteGeneracije_NijeIstiDokument"
        Case 55: TestName = "T_StornoJeEkranNeRezim"
        Case 56: TestName = "T_Storno_UgovorIRadnje"
        Case 57: TestName = "T_StornoEkran_KolonaIdentiteta"
        Case 58: TestName = "T_StornoEkran_SvakaListaVracaRedove"
        Case 59: TestName = "T_PrefillBezBroja_PredlaziBroj"
        Case 54: TestName = "T_MapaImena_KljucNosiKolone"
        Case 53: TestName = "T_KesTabela_NeMemoiseNeuspeh"
        Case 52: TestName = "T_StornoIzvrsi_ZbirnaImenujeVezanuPrijemnicu"
        Case 51: TestName = "T_StorniranSibling_ZadrzavaSvojBlok"
        Case 50: TestName = "T_BlokoviF8_PoIdentitetu"
        Case 49: TestName = "T_IspravkaZbirne_KapijaNaObeStrane"
        Case 48: TestName = "T_CiljnaZbirnaDvosmislena_Staje"
        Case 47: TestName = "T_KapijaZbirne_FailClosedNaSvojuGresku"
        Case 46: TestName = "T_ZatecenContext_NePrevezujeTudjePrijemnice"
        Case 45: TestName = "T_OtpremnicaNadDvosmislenomZbirnom_Staje"
        Case 44: TestName = "T_StorniranVlasnik_JosImaAktivnuDecu"
        Case 43: TestName = "T_ZavrsetakIspravke_NeDegradiraOldDocID"
        Case 42: TestName = "T_ZamenaZbirne_NeDiraDecuTudje"
        Case 41: TestName = "T_ZbirnaKaskada_StajeNaDvosmislenom"
        Case 40: TestName = "T_SoleOwner_MeriDokumenteNeBrojeve"
        Case 39: TestName = "T_OtkupBezGeneracije_NeStorniraTudjeOM"
        Case 38: TestName = "T_Zbirna_ZaglavljePoGeneracijiKaskadaStaje"
        Case 37: TestName = "T_IspravkaPrijemnice_PodKolizijomBroja"
        Case 36: TestName = "T_Preflight_KoristiIdentitet"
        Case 35: TestName = "T_F8_IzabranRedOstajeIzabran"
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
        Case 20: T_Storno_TipBiraTabeluIKolone
        Case 21: T_StornoDok_KapijePreUpisa
        Case 22: T_PrefillIzStorniranog_CitaSvojuTabelu
        Case 23: T_FrameworkIspravke_SamoCetiriTipa
        Case 24: T_Prefill_PoIdentitetuNePoBroju
        Case 25: T_IspravkaDetekcija_FailClosed
        Case 26: T_Oporavak_UgovorIRadnje
        Case 27: T_Oporavak_CiljneListe
        Case 28: T_IspravkaPrijemnice_SkipIRelink
        Case 29: T_RelinkPoGeneraciji_NeDiraTudjDokument
        Case 30: T_PrevezivanjeNaZbirnu_PaletaIdePoIdentitetu
        Case 31: T_ZadataGeneracijaKojeNema_Staje
        Case 32: T_VerdiktPoIdentitetu_RelabelSeNePreskace
        Case 33: T_DeljenaPaleta_SuStanarPoIdentitetu
        Case 34: T_IstiBrojRazliciteGeneracije_NijeIstiDokument
        Case 55: T_StornoJeEkranNeRezim
        Case 56: T_Storno_UgovorIRadnje
        Case 57: T_StornoEkran_KolonaIdentiteta
        Case 58: T_StornoEkran_SvakaListaVracaRedove
        Case 59: T_PrefillBezBroja_PredlaziBroj
        Case 54: T_MapaImena_KljucNosiKolone
        Case 53: T_KesTabela_NeMemoiseNeuspeh
        Case 52: T_StornoIzvrsi_ZbirnaImenujeVezanuPrijemnicu
        Case 51: T_StorniranSibling_ZadrzavaSvojBlok
        Case 50: T_BlokoviF8_PoIdentitetu
        Case 49: T_IspravkaZbirne_KapijaNaObeStrane
        Case 48: T_CiljnaZbirnaDvosmislena_Staje
        Case 47: T_KapijaZbirne_FailClosedNaSvojuGresku
        Case 46: T_ZatecenContext_NePrevezujeTudjePrijemnice
        Case 45: T_OtpremnicaNadDvosmislenomZbirnom_Staje
        Case 44: T_StorniranVlasnik_JosImaAktivnuDecu
        Case 43: T_ZavrsetakIspravke_NeDegradiraOldDocID
        Case 42: T_ZamenaZbirne_NeDiraDecuTudje
        Case 41: T_ZbirnaKaskada_StajeNaDvosmislenom
        Case 40: T_SoleOwner_MeriDokumenteNeBrojeve
        Case 39: T_OtkupBezGeneracije_NeStorniraTudjeOM
        Case 38: T_Zbirna_ZaglavljePoGeneracijiKaskadaStaje
        Case 37: T_IspravkaPrijemnice_PodKolizijomBroja
        Case 36: T_Preflight_KoristiIdentitet
        Case 35: T_F8_IzabranRedOstajeIzabran
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

    ' Do v6-ui-142 je ovde prva tvrdnja bila da F8 (storno) vraca
    ' OTKUI_TODO_NEVEZANO -- "nepokriven rezim". Ta tvrdnja je nestala sa
    ' rezimom: storno je svoj ekran, ciji Scr_Meta kaze "upis=ne", pa se
    ' Scr_Save nad njim uopste ne zove. SVIH SEDAM preostalih rezima je
    ' vezano, i test to sada i tvrdi -- nijedan ne sme da padne u Case Else.

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

    ' F1 i F2 zatvaraju spisak: nijedan od sedam rezima ne sme da vrati
    ' "nije vezano". Ako neka ruta nestane, ovde pada po imenu -- ranije je
    ' isti simptom bio nevidljiv, jer je F8 legitimno vracao bas tu poruku.
    Set p = PoljaEkrana(modScrDokumenti.modeKey("F1"))
    AssertEq (modScrDokumenti.Scr_Save(p) <> Poruka("OTKUI_TODO_NEVEZANO")), True, _
             "otkupni list je vezan na svoju rutinu"
    Set p = PoljaEkrana(modScrDokumenti.modeKey("F2"))
    AssertEq (modScrDokumenti.Scr_Save(p) <> Poruka("OTKUI_TODO_NEVEZANO")), True, _
             "otpremnica je vezana na svoju rutinu"
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

' STORNO CITA TABELU IZABRANOG TIPA, ne uvek tblOtpremnica.
'
' Do v6-ui-118 je "STORNO" bio tih sinonim za "OTPREMNICA" u ModeTable i u
' desetak Col* funkcija, pa je storno centar mogao da pokaze samo otpremnice.
' Pada ako se neki tip izgubi iz TabelaTipa ili iz Col* funkcija: tada storno
' opet svira po jednoj tabeli, a mreza tiho pokazuje pogresne dokumente pod
' pravim naslovom -- greska koju operater ne moze da vidi.
'
' PRETARGETIRAN u v6-ui-142: mera je ista, ali seam vise nije rezim F8 nego
' KLJUC TIPA. Storno je svoj ekran, pa "koja tabela" vise ne zavisi od
' ActiveMode -- i test to sada trazi tako kako produkcija stvarno pita.
'
' Kolone se proveravaju ZAJEDNO sa tabelom: tabela bez odgovarajucih kolona
' daje praznu mrezu, sto izgleda kao "nema dokumenata".
Private Sub T_Storno_TipBiraTabeluIKolone()
    Dim tipovi As Variant, tabele As Variant, i As Long, cols As Variant

    tipovi = Array(STIP_OTKUP, STIP_OTPREMNICA, STIP_ZBIRNA, STIP_PRIJEMNICA, _
                   STIP_ISPLATE, STIP_UPLATE, STIP_REVERSI, STIP_FAKTURA, STIP_IZVOD)
    tabele = Array(TBL_OTKUP, TBL_OTPREMNICA, TBL_ZBIRNA, TBL_PRIJEMNICA, _
                   TBL_NOVAC, TBL_NOVAC, TBL_AMBALAZA, TBL_FAKTURE, TBL_BANKA_IMPORT)

    For i = 0 To UBound(tipovi)
        AssertEq modScrDokumenti.TabelaTipa(CStr(tipovi(i))), CStr(tabele(i)), _
                 "Storno / " & CStr(tipovi(i)) & " cita svoju tabelu"
        cols = modScrDokumenti.GridCols(CStr(tipovi(i)), True)
        AssertEq (IsArray(cols)), True, _
                 "Storno / " & CStr(tipovi(i)) & " ima opis kolona"
        AssertEq (UBound(cols) >= 3), True, _
                 "Storno / " & CStr(tipovi(i)) & " ima bar cetiri kolone"
    Next i

    ' Rezim i dalje mora da stigne do iste tabele -- ModeTable je od v6-ui-142
    ' samo TabelaTipa(modeKey()), pa bi razilazenje ta dva puta znacilo da
    ' unosni ekran i storno gledaju u razlicite tabele za isti dokument.
    AssertEq modScrDokumenti.ModeTable("F4"), TBL_PRIJEMNICA, _
             "rezim i tip vode u istu tabelu"

    ' Broj zbirne postoji samo tamo gde ga dokument NOSI. Dok je storno bio
    ' otpremnica, cip "Bez zbirne" je bio ukljucen i nad novcem, gde tblNovac
    ' tu kolonu nema.
    AssertEq modScrDokumenti.ModeHasZbirna("F5"), False, _
             "novac nema pojam zbirne"
    AssertEq modScrDokumenti.ModeHasZbirna("F1"), True, _
             "otkupni list ima pojam zbirne"
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

' PREFILL POSLE STORNA CITA TABELU SVOG TIPA, I TO PO PRAVOM IMENU KOLONE.
'
' Fixture je za ovo idealan: otkupni list OTK-TEST-1 i otpremnica OTP-TEST-1
' dele BROJ "1/TEST", a nose razlicite kolicine (400 vs 1000). Ako prefill
' pogresi tabelu, ispravka otkupa ponudi kilograme otpremnice - i to niko
' ne bi primetio, jer je dokument i dalje "ispravan", samo pogresan.
'
' Zbirna je drugi razlog za ovaj test: njena kolicina se u semi zove
' "UkupnoKolicina", a ambalaza "UkupnoAmbalaze". Literal "Kolicina" bi tiho
' vratio nulu, pa bi ispravka zbirne dosla prazna.
'
' Tri pravila koja se lako izgube i koja test drzi:
'   datum se preuzima iz storniranog (ispravka sutradan ne menja dan)
'   broj dokumenta se NE preuzima (ispravka je NOV dokument, nov broj)
'   nula se ne salje (prazno polje i "0" su dva razlicita stanja)
Private Sub T_PrefillIzStorniranog_CitaSvojuTabelu()
    Dim s As String

    ' --- OTKUPNI LIST: 400 kg, 40 gajbi, kooperant, parcela ---
    s = modStornoDok.PrefillIzStorniranog(STIP_OTKUP, "1/TEST", "")
    AssertEq (Len(s) > 0), True, "prefill otkupa nije prazan"
    AssertEq SpecVal(s, "kol1"), "400", "otkup daje SVOJU kolicinu (ne otpremnicinu)"
    AssertEq SpecVal(s, "amb1"), "40", "otkup daje svoje gajbe"
    AssertEq SpecVal(s, "cena"), "50", "otkup daje svoju cenu"
    AssertEq SpecVal(s, "partnerid"), FX_KOOPERANT, "partner otkupa je kooperant"
    AssertEq SpecVal(s, "omid"), FX_STANICA, "otkup nosi otkupno mesto"
    AssertEq SpecVal(s, "parcela"), FX_PARCELA, "parcelu ima SAMO otkup"
    AssertEq SpecVal(s, "brzbirne"), FX_ZBIRNA, "otkup nosi broj zbirne"
    AssertEq SpecVal(s, "dveklase"), "1", "jednoklasni dokument"
    ' Datum iz storniranog, ne danasnji.
    AssertEq SpecVal(s, "datum"), "15.03.2026", "datum se preuzima iz storniranog"
    ' Broj se NE preuzima: ispravka je nov dokument sa novim brojem.
    AssertEq SpecVal(s, "brdok"), "", "broj dokumenta se NE preuzima"
    ' Nula se ne salje: fixture nema izdatu ambalazu na otkupu.
    AssertEq SpecVal(s, "ambpr"), "", "nula se ne salje kao vrednost"

    ' --- OTPREMNICA: isti BROJ, druga tabela, druge kolicine ---
    s = modStornoDok.PrefillIzStorniranog(STIP_OTPREMNICA, "1/TEST", "")
    AssertEq SpecVal(s, "kol1"), "1000", "otpremnica daje SVOJU kolicinu"
    AssertEq SpecVal(s, "amb1"), "100", "otpremnica daje svoje gajbe"
    AssertEq SpecVal(s, "omid"), FX_STANICA, "otpremnica nosi otkupno mesto"
    AssertEq SpecVal(s, "parcela"), "", "otpremnica nema parcelu"
    AssertEq SpecVal(s, "partnerid"), "", "otpremnica nema partnera (njen je stanica)"

    ' --- ZBIRNA: UkupnoKolicina / UkupnoAmbalaze, bez cene i bez zbirne ---
    s = modStornoDok.PrefillIzStorniranog(STIP_ZBIRNA, FX_ZBIRNA, "")
    AssertEq SpecVal(s, "kol1"), "1000", "zbirna cita UkupnoKolicina, ne Kolicina"
    AssertEq SpecVal(s, "amb1"), "100", "zbirna cita UkupnoAmbalaze, ne KolAmbalaze"
    AssertEq SpecVal(s, "cena"), "", "zbirna NEMA cenu (nema je ni tabela)"
    AssertEq SpecVal(s, "brzbirne"), "", "zbirna ne preuzima samu sebe kao broj zbirne"

    ' --- nepostojeci dokument i tip koji se ne prefiluje ---
    AssertEq modStornoDok.PrefillIzStorniranog(STIP_OTKUP, "NE-POSTOJI", ""), "", _
             "nepostojeci dokument ne daje prefill"
    AssertEq modStornoDok.PrefillIzStorniranog(STIP_IZVOD, "1/TEST", ""), "", _
             "izvod se ne prefiluje (nije dokument unosa)"
End Sub

' FRAMEWORK ISPRAVKE VAZI SAMO ZA CETIRI TIPA.
'
' Otkup, novac, faktura i izvod nemaju nizvodni tok o kome se odlucuje, pa
' im je storno obican - isto kao u legacy formi, gde TryRunCorrectionFramework
' za njih vraca False i posao preuzima obican Select Case.
'
' Pada ako se neki tip ubaci u framework: tada bi npr. storno isplate poceo
' da nudi "ISPRAVKA / DUPLIKAT / PONISTENJE", a modStornoFlow za novac nema
' nijednu od tih grana - dokument bi ostao neopisan i nestorniran.
Private Sub T_FrameworkIspravke_SamoCetiriTipa()
    Dim jesu As Variant, nisu As Variant, i As Long

    jesu = Array(STIP_OTPREMNICA, STIP_ZBIRNA, STIP_PRIJEMNICA, STIP_REVERSI)
    nisu = Array(STIP_OTKUP, STIP_ISPLATE, STIP_UPLATE, STIP_FAKTURA, STIP_IZVOD)

    For i = 0 To UBound(jesu)
        AssertEq (Len(modStornoDok.TipUFlowDoc(CStr(jesu(i)))) > 0), True, _
                 "framework tip: " & CStr(jesu(i))
    Next i
    For i = 0 To UBound(nisu)
        AssertEq modStornoDok.TipUFlowDoc(CStr(nisu(i))), "", _
                 "obican storno, bez framework-a: " & CStr(nisu(i))
        ' Ako tip nije framework tip, ne sme ni da trazi izbor moda niti da
        ' ima sta da izvrsi kroz njega.
        AssertEq modStornoDok.StornoTraziIzborModa(CStr(nisu(i)), "1/TEST", ""), False, _
                 "ne trazi izbor moda: " & CStr(nisu(i))
        AssertEq (modStornoDok.StornoIzvrsiMod(CStr(nisu(i)), "1/TEST", "", _
                  SV_MODE_ISPRAVKA, False, False) Is Nothing), True, _
                 "framework ne izvrsava nista nad: " & CStr(nisu(i))
    Next i
End Sub

' PREFILL BIRA DOKUMENT PO PK-u, NE PO BROJU.
'
' Fixture ima dve AKTIVNE prijemnice sa istim brojem i razlicitim kupcem
' (PRJ-TEST-A / KUP-TEST-1, PRJ-TEST-B / KUP-TEST-2). Tako i mora da bude u
' produkciji: GenerateBrojPrijemnice racuna sekvencu PO KUPCU, pa dva kupca
' istog dana dobiju isti "1/ddmmyy".
'
' Ako prefill krene od broja, ispravka jednog kupca ponudi kolicine i cenu
' DRUGOG. Dokument bi i dalje bio "ispravan" - samo tudji.
'
' Prethodna verzija ovog testa zvala je PrefillIzStorniranog sa oldDocID:="",
' pa bas ovaj slucaj nije ni doticala.
Private Sub T_Prefill_PoIdentitetuNePoBroju()
    Dim sA As String, sB As String

    sA = modStornoDok.PrefillIzStorniranog(STIP_PRIJEMNICA, FX_PRIJ_BROJ, "PRJ-TEST-A")
    sB = modStornoDok.PrefillIzStorniranog(STIP_PRIJEMNICA, FX_PRIJ_BROJ, "PRJ-TEST-B")

    AssertEq (Len(sA) > 0), True, "prefill po PK-u A nije prazan"
    AssertEq (Len(sB) > 0), True, "prefill po PK-u B nije prazan"

    ' Isti broj, dva PK-a -> DVE razlicite vrednosti. Da prefill ide po broju,
    ' obe strane bi vratile isti dokument.
    AssertEq SpecVal(sA, "kol1"), "300", "PK A daje SVOJU kolicinu"
    AssertEq SpecVal(sB, "kol1"), "700", "PK B daje SVOJU kolicinu"
    AssertEq SpecVal(sA, "partnerid"), FX_KUPAC, "PK A daje svog kupca"
    AssertEq SpecVal(sB, "partnerid"), FX_KUPAC2, "PK B daje svog kupca"
    AssertEq SpecVal(sA, "cena"), "60", "PK A daje svoju cenu"
    AssertEq SpecVal(sB, "cena"), "80", "PK B daje svoju cenu"

    ' Nepoznat PK ne sme da "padne nazad" na prvi red istog broja - to bi bila
    ' ista greska, samo tise.
    AssertEq modStornoDok.PrefillIzStorniranog(STIP_PRIJEMNICA, FX_PRIJ_BROJ, "PRJ-NE-POSTOJI"), "", _
             "nepoznat PK ne pogadja tudji dokument istog broja"

    ' Ista kolizija je razlog zasto prevezivanje po broju mora da stane.
    ' Racun je jedan i deli ga sa storno jezgrom (RequireJedanVlasnikPoBroju),
    ' sa kompozitnim vlasnistvom po tipu dokumenta.
    AssertEq (VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, FX_PRIJ_BROJ, _
              "test", False, Array(COL_PRJ_KUPAC)).count > 1), True, _
             "broj sa dva kupca se prijavljuje kao dvosmislen"
    AssertEq VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, FX_PRIJ_STORNO, _
             "test", False, Array(COL_PRJ_KUPAC)).count, 0, _
             "storniran dokument se ne broji medju AKTIVNIM vlasnicima"
    ' ...ali se broji kad se izricito trazi i stornirane - to je slucaj IZVORA
    ' prevezivanja, gde je izvor bas storniran dokument.
    ' Isti racun, sa ukljucenim storniranima, vidi i KOLIZIONI par -- a bas to
    ' je slucaj IZVORA prevezivanja, gde je izvor storniran dokument.
    AssertEq VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, FX_PRIJ_KOLIZIJA, _
             "test", True, Array(COL_PRJ_KUPAC)).count, 2, _
             "sa storniranima se vidi da broj nose DVA dokumenta"
    AssertEq VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, FX_PRIJ_KOLIZIJA, _
             "test", False, Array(COL_PRJ_KUPAC)).count, 0, _
             "medju AKTIVNIMA ih nema - oba su stornirana"
End Sub

' NEIZVESNOST ZAUSTAVLJA UPIS, NE PROPUSTA GA.
'
' Kad se ne zna da li ispravka na cekanju postoji, "nastavi kao obican unos"
' znaci: nova prijemnica dobija SVEZE palete, stare ostaju osirocene, a
' correction ostaje PENDING i ceka jos jednu prijemnicu. Zato je detekcija
' fail-closed.
'
' Fixture ima DVE ispravke na cekanju nad otpremnicom - namerno ne nad
' prijemnicom, jer detekcija prijemnice pita operatera kroz MsgBox, a MsgBox
' u headless runu visi. Pravilo je isto i deli ga ista rutina.
Private Sub T_IspravkaDetekcija_FailClosed()
    Dim cid As String, stari As String, parent As String, razlog As String
    Dim ishod As Long

    ' Dve na cekanju -> STOP, sa razlogom (safe-stop).
    ishod = modDokUnos.NadjiIspravku(FLOW_DOC_OTPREMNICA, cid, stari, parent, razlog)
    AssertEq ishod, -1, "dve ispravke na cekanju zaustavljaju upis"
    AssertEq (Len(razlog) > 0), True, "safe-stop nosi razlog za operatera"

    ' Nijedna za drugi tip -> obican unos. Dokazuje i da se tipovi ne mesaju:
    ' otpremnicke ispravke ne smeju da zaustave unos prijemnice.
    ishod = modDokUnos.NadjiIspravku(FLOW_DOC_PRIJEMNICA, cid, stari, parent, razlog)
    AssertEq ishod, 0, "ispravka drugog tipa ne dira ovaj unos"
    AssertEq razlog, "", "bez ispravke nema ni razloga"
    AssertEq cid, "", "bez ispravke nema ni CorrectionID"
End Sub

' EKRAN OPORAVAK JE U REGISTRU I ODGOVARA NA UGOVOR.
'
' Ljuska ne poznaje nijedan ekran po imenu - sve ide kroz Application.Run po
' redu iz registra. Zato je "ekran postoji" isto sto i "modul odgovara na
' Scr_Meta": ako se ime modula u registru omakne, sidebar ga samo prikaze
' prigusenog i niko ne zna zasto. Ovaj test je jedino mesto koje to hvata.
'
' Radnje po listi su drugi deo: ciljne liste (Zbirne, Prijemnice) NE SMEJU
' da imaju dugme "Prevezi". Cilj se bira klikom na red; dugme nad ciljem bi
' prevezivalo cilj na samog sebe.
Private Sub T_Oporavak_UgovorIRadnje()
    Dim liste As Variant, i As Long, kljucevi As String

    ' 1) Registar zna za ekran, i modul odgovara.
    AssertEq (Len(modUiScreens.ScrRowByKey("OPORAVAK")) > 0), True, _
             "OPORAVAK postoji u registru ekrana"
    AssertEq modUiScreens.ScrPostoji("OPORAVAK"), True, _
             "modul ekrana odgovara na Scr_Meta (kasno vezivanje radi)"
    AssertEq (InStr(modUiScreens.ScrMeta("OPORAVAK"), "kljuc=OPORAVAK") > 0), True, _
             "Scr_Meta prijavljuje svoj kljuc"

    ' 2) Sest lista, i to bas ovih sest.
    liste = modScrOporavak.Scr_Liste()
    AssertEq (UBound(liste) + 1), 6, "ekran ima sest lista"
    For i = 0 To UBound(liste)
        kljucevi = kljucevi & "|" & Split(CStr(liste(i)), "|")(0)
    Next i
    AssertEq kljucevi, "|NEDOVRSENO|PRIJEMNICE|ZBIRNE|PALETE|CILJPRIJ|UNDO", _
             "redosled i kljucevi lista"

    ' 3) Radnje po listi. Prazno = lista je samo pregled ili izbor cilja.
    modScrOporavak.Scr_OpoTestSet "NEDOVRSENO", "", ""
    AssertEq modScrOporavak.Scr_Radnje(), "", "Nedovrseno je samo pregled"
    modScrOporavak.Scr_OpoTestSet "ZBIRNE", "", ""
    AssertEq modScrOporavak.Scr_Radnje(), "", "ciljna lista zbirnih nema radnju"
    modScrOporavak.Scr_OpoTestSet "CILJPRIJ", "", ""
    AssertEq modScrOporavak.Scr_Radnje(), "", "ciljna lista prijemnica nema radnju"
    modScrOporavak.Scr_OpoTestSet "PRIJEMNICE", "", ""
    AssertEq (InStr(modScrOporavak.Scr_Radnje(), "prevezipri:") = 1), True, _
             "osirotele prijemnice imaju Prevezi"
    modScrOporavak.Scr_OpoTestSet "PALETE", "", ""
    AssertEq (InStr(modScrOporavak.Scr_Radnje(), "prevezipal:") = 1), True, _
             "osirotele palete imaju Prevezi"
    modScrOporavak.Scr_OpoTestSet "UNDO", "", ""
    AssertEq (InStr(modScrOporavak.Scr_Radnje(), "vrati:") = 1), True, _
             "Undo lista ima Vrati storno"
    ' Radnja koja menja podatke i tesko se poziva nazad mora da bude crvena.
    AssertEq (InStr(modScrOporavak.Scr_Radnje(), ":danger:") > 0), True, _
             "Vrati storno nosi danger stil"

    modScrOporavak.Scr_OpoTestSet "NEDOVRSENO", "", ""
End Sub

' CILJNE LISTE: SAMO AKTIVNI, I JEDAN RED PO DOKUMENTU - NE PO BROJU.
'
' Prva verzija ovog testa tvrdila je "cilj prevezivanja JESTE broj". To je
' bila greska koju je test KODIFIKOVAO umesto da je uhvati - i prosla je
' jer fixture tada nije imao nijedan par dokumenata koji dele broj.
'
' Broj se racuna PO KUPCU (GenerateBrojPrijemnice), pa dva kupca istog dana
' dobiju isti "1/ddmmyy". To su DVA dokumenta. Dedup po samom broju sveo bi
' ih na jedan red: operater bi izabrao "cilj" ne znajuci ciji je, a
' prevezivanje bi otislo na onaj koji zatekne poslednji.
'
' Klase I i II i dalje dele jedan red - one dele i broj i vlasnika, pa jesu
' jedan dokument.
Private Sub T_Oporavak_CiljneListe()
    Dim d As Variant, redovi As Variant, n As Long, i As Long, nadjen As Long
    Dim storniranih As Long, istiBroj As Long, vlasnici As String

    modScrOporavak.Scr_OpoTestSet "ZBIRNE", "", ""
    d = modScrOporavak.Scr_Rows("sve", "")
    AssertEq IsArray(d), True, "ciljna lista zbirnih vraca ugovor mreze"
    n = CLng(d(2))
    AssertEq (n > 0), True, "fixture ima bar jednu aktivnu zbirnu"

    redovi = d(1)
    For i = 1 To n
        If CStr(redovi(i, 1)) = FX_ZBIRNA Then nadjen = nadjen + 1
        If CStr(redovi(i, 1)) = FX_ZBIRNA_STORNO Then storniranih = storniranih + 1
    Next i
    AssertEq nadjen, 1, "zbirna iz fixture-a stoji TACNO jednom"
    ' Zasto ovo mora da postoji: prevezivanje na STORNIRAN cilj bi napravilo
    ' drugu siroticu umesto da resi prvu.
    AssertEq storniranih, 0, "stornirana zbirna se NE nudi kao cilj"

    ' --- KOLIZIJA ZBIRNIH: isti broj, ISTI kupac, dva vozaca ---
    '
    ' Broj zbirne generator DRZI JEDINSTVENIM: SuggestNextBroj za ZBR bumpuje
    ' sekvencu dok BrojZbirneExists ne kaze da je slobodan, a ApplyMirrorPrefix
    ' dodaje 'S' da se mirror-vozac ne sudari sa realnim. Dva reda istog broja
    ' zato mogu nastati SAMO mimo generatora -- rucnim unosom (auto-broj se
    ' iskljucuje u Podesavanjima), uvozom ili ispravkom u tabeli.
    '
    ' Test brani bas taj slucaj: lista je vlasnikom smatrala samo kupca i spajala
    ' ih u JEDAN red, pa operater ne bi mogao da izabere onaj koji mu treba.
    Dim duplih As Long, vozaci As String
    For i = 1 To n
        If CStr(redovi(i, 1)) = FX_ZBIRNA_DUPL Then
            duplih = duplih + 1
            vozaci = vozaci & "|" & CStr(redovi(i, 3))
        End If
    Next i
    AssertEq duplih, 2, "isti broj zbirne kod dva vozaca daje DVA ciljna dokumenta"
    AssertEq (InStr(vozaci, FX_VOZAC2) > 0), True, "drugi vozac je vidljiv u listi"
    AssertEq (InStr(vozaci, FX_KUPAC) > 0), True, _
             "vlasnik je kompozit -- uz vozaca se vidi i kupac"

    ' --- KOLIZIJA BROJEVA: dva kupca, isti broj -> DVA reda ---
    modScrOporavak.Scr_OpoTestSet "CILJPRIJ", "", ""
    d = modScrOporavak.Scr_Rows("sve", "")
    AssertEq IsArray(d), True, "ciljna lista prijemnica vraca ugovor mreze"
    n = CLng(d(2))
    redovi = d(1)
    For i = 1 To n
        If CStr(redovi(i, 1)) = FX_PRIJ_BROJ Then
            istiBroj = istiBroj + 1
            vlasnici = vlasnici & "|" & CStr(redovi(i, 3))
        End If
        AssertEq (CStr(redovi(i, 1)) = FX_PRIJ_STORNO), False, _
                 "stornirana prijemnica se NE nudi kao cilj"
    Next i
    AssertEq istiBroj, 2, "dva kupca sa istim brojem daju DVA ciljna dokumenta"
    AssertEq (InStr(vlasnici, FX_KUPAC) > 0), True, "prvi vlasnik je vidljiv u listi"
    AssertEq (InStr(vlasnici, FX_KUPAC2) > 0), True, "drugi vlasnik je vidljiv u listi"

    ' Pretraga suzava istu listu.
    modScrOporavak.Scr_OpoTestSet "ZBIRNE", "", ""
    d = modScrOporavak.Scr_Rows("sve", "NE-POSTOJI-NIGDE")
    AssertEq CLng(d(2)), 0, "pretraga koja nista ne pogadja daje praznu listu"

    modScrOporavak.Scr_OpoTestSet "NEDOVRSENO", "", ""
End Sub


' ISPRAVKA PRIJEMNICE OD KRAJA DO KRAJA: preskoci paletizaciju, prevezi palete.
'
' Ovo je najrizicniji put celog paketa i do sada je bio samo na operaterskoj
' checklisti. Pokriva ga BEZ ijednog novog seam-a u produkcionom kodu: recnik
' koji PrijemnicaUpisi prima je vec javni ulazni ugovor (NoviPrijemnicaUnos ga
' i objavljuje), pa test postavlja "ispravkaID" direktno. MsgBox iz
' PrepoznajIspravkuPrijemnice tako uopste nije na putanji -- a odluka koju taj
' dijalog donosi vec je pokrivena kroz NadjiIspravku.
'
' Test je DIFERENCIJALAN, i to namerno. "Nema svezih paleta" samo po sebi ne
' dokazuje nista: isto bi se videlo da je paletiranje ugaseno u Podesavanjima,
' ili da paletizacija uopste ne radi nad ovim fixture-om. Zato se ISTI upis
' izvrsi dvaput:
'
'   A) bez ispravke  -> sveza paletizacija MORA da napravi stavku
'   B) kao ispravka  -> nema sveze, a stara stavka je PREVEZANA
'
' Tek razlika izmedju A i B dokazuje da SetPaletizeSkip radi.
'
' Gajbice su namerno jednake starim (40): tada ReassignPaleteToPrijemnica_TX
' ne prijavljuje razliku, pa PaletaAdjustPrompt (koji ume da pita operatera)
' ne ulazi u igru.
Private Sub T_IspravkaPrijemnice_SkipIRelink()
    Dim p As Object, res As String, poruke As String
    Dim cid As String, prevPal As String

    ' Preduslov: bez ukljucenog paletiranja ceo test meri prazno.
    prevPal = GetConfigValue(CFG_PALETIRANJE)
    SetConfigValue CFG_PALETIRANJE, "DA"
    AssertEq IsPaletiranjeEnabled(), True, "preduslov: paletiranje je ukljuceno"
    AssertEq StavkiZaPrijemnicu(FX_PRIJ_STORNO), 1, _
             "preduslov: stornirana prijemnica nosi jednu paletnu stavku"

    ' --- A) KONTROLA: obican upis -> sveza paletizacija RADI ---
    Set p = NoviPrijemnicaUnos()
    PopuniPrijemnicu p, "T-KTRL-1"
    res = modDokUnos.PrijemnicaUpisi(p, poruke)
    AssertEq (Len(res) > 0), True, "kontrolni upis je prosao"
    AssertEq (StavkiZaPrijemnicu("T-KTRL-1") > 0), True, _
             "bez ispravke sveza paletizacija pravi stavku"

    ' --- B) ISPRAVKA: nema sveze paletizacije, stara stavka se prevezuje ---
    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, _
              FLOW_DOC_PRIJEMNICA, "PRJ-TEST-S", FX_PRIJ_STORNO, _
              FLOW_DOC_PRIJEMNICA, , , FLOW_DOC_ZBIRNA, , FX_ZBIRNA, _
              "test: ispravka prijemnice")
    AssertEq (Len(cid) > 0), True, "correction context je kreiran"

    Set p = NoviPrijemnicaUnos()
    PopuniPrijemnicu p, "T-ISPR-1"
    p("ispravkaID") = cid
    p("ispravkaStariBroj") = FX_PRIJ_STORNO

    res = modDokUnos.PrijemnicaUpisi(p, poruke)
    AssertEq (Len(res) > 0), True, "upis ispravke je prosao"

    AssertEq StavkiZaPrijemnicu("T-ISPR-1"), 1, _
             "ispravka nosi prevezenu paletnu stavku"
    AssertEq StavkiZaPrijemnicu(FX_PRIJ_STORNO), 0, _
             "stara prijemnica vise ne nosi paletnu stavku"

    AssertEq GajbicaZaPrijemnicu("T-ISPR-1"), 40, _
             "prevezena roba je 40 gajbica"

    ' KLJUCNA TVRDNJA, i broji SVE redove - i stornirane.
    '
    ' Dve ranije verzije ovog testa bile su placebo: brojale su samo AKTIVNE
    ' stavke, pa je sabotaza koja ukloni SetPaletizeSkip ostajala ZELENA.
    ' Izmereno stanje pokazuje zasto: bez preskakanja se sveza paleta ipak
    ' napravi, a onda je ReassignPaleteToPrijemnica_TX odmah STORNIRA. U
    ' aktivnom preseku se zato ne vidi nista - ostaje samo trag:
    '
    '   sa preskakanjem : [ST T-ISPR-1 gaj=40]
    '   bez preskakanja : [ST T-ISPR-1 gaj=40] + [ST T-ISPR-1 gaj=40 st=Da]
    '
    ' To je tacno ono sto komentar uz SetPaletizeSkip i opisuje kao stetu:
    ' "kreirale bi se palete koje se odmah storniraju (prazna otvorena paleta
    ' + potrosen broj)". Broj palete se ne vraca.
    AssertEq SvihStavkiZaPrijemnicu("T-ISPR-1"), 1, _
             "nema paletizacije-pa-storna: nijedna stavka nije nastala uzalud"

    ' Kontekst mora da bude ZATVOREN - inace bi sledeci unos opet bio ponudjen
    ' kao zamena za isti stari dokument.
    AssertEq modStornoContext.GetCorrectionField(cid, COL_SV_STATUS), SV_STATUS_COMPLETED, _
             "correction context je zatvoren posle uspesnog prevezivanja"
    ' Recnik se TROSI: ponovljen poziv ne sme da prevezuje drugi put.
    AssertEq CStr(p("ispravkaID")), "", "ispravka je potrosena iz recnika"

    SetConfigValue CFG_PALETIRANJE, prevPal
End Sub

' Zajednicka polja za oba upisa iz gornjeg testa. Kolicine i gajbice su iste
' kao na storniranoj prijemnici (400 kg / 40 gajbica) - v. napomenu o
' PaletaAdjustPrompt.
Private Sub PopuniPrijemnicu(ByVal p As Object, ByVal broj As String)
    p("datum") = CDate(FX_DATUM)
    p("kupacID") = FX_KUPAC
    p("vozacID") = FX_VOZAC
    p("brDok") = broj
    p("brojZbirne") = FX_ZBIRNA
    p("vrsta") = FX_VRSTA
    p("sorta") = FX_SORTA
    p("tipAmb") = FX_TIP_AMB
    p("kolicinaI") = 400
    p("cenaI") = 50
    p("kolAmb") = 40
End Sub

' Zbir gajbica na AKTIVNIM paletnim stavkama date prijemnice. Ovo je mera
' ROBE - jedina koja razlikuje "prevezeno" od "paletizovano pa prevezeno",
' jer se stavke na istoj paleti spajaju u jedan red.
Private Function GajbicaZaPrijemnicu(ByVal brojPrij As String) As Long
    Dim d As Variant, i As Long, cBr As Long, cSt As Long, cGa As Long
    d = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(d) Then Exit Function
    cBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    cGa = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If cBr = 0 Or cGa = 0 Then Exit Function
    For i = 1 To UBound(d, 1)
        If Trim$(NzToText(d(i, cBr))) = brojPrij Then
            If cSt = 0 Then
                GajbicaZaPrijemnicu = GajbicaZaPrijemnicu + NzL(d(i, cGa))
            ElseIf UCase$(Trim$(NzToText(d(i, cSt)))) <> "DA" Then
                GajbicaZaPrijemnicu = GajbicaZaPrijemnicu + NzL(d(i, cGa))
            End If
        End If
    Next i
End Function

' SVE paletne stavke datog broja, ukljucujuci stornirane. Aktivan presek ne
' razlikuje "nije paletizovano" od "paletizovano pa stornirano" - a bas ta
' razlika je ono sto SetPaletizeSkip sprecava.
Private Function SvihStavkiZaPrijemnicu(ByVal brojPrij As String) As Long
    Dim d As Variant, i As Long, cBr As Long
    d = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(d) Then Exit Function
    cBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    If cBr = 0 Then Exit Function
    For i = 1 To UBound(d, 1)
        If Trim$(NzToText(d(i, cBr))) = brojPrij Then _
            SvihStavkiZaPrijemnicu = SvihStavkiZaPrijemnicu + 1
    Next i
End Function

' Broj AKTIVNIH paletnih stavki koje pokazuju na dati broj prijemnice.
Private Function StavkiZaPrijemnicu(ByVal brojPrij As String) As Long
    Dim d As Variant, i As Long, cBr As Long, cSt As Long
    d = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(d) Then Exit Function
    cBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If cBr = 0 Then Exit Function
    For i = 1 To UBound(d, 1)
        If Trim$(NzToText(d(i, cBr))) = brojPrij Then
            If cSt = 0 Then
                StavkiZaPrijemnicu = StavkiZaPrijemnicu + 1
            ElseIf UCase$(Trim$(NzToText(d(i, cSt)))) <> "DA" Then
                StavkiZaPrijemnicu = StavkiZaPrijemnicu + 1
            End If
        End If
    Next i
End Function

' PREVEZIVANJE DIRA SAMO SVOJ DOKUMENT, I KAD DVA DELE BROJ.
'
' Ovo je test koji je nedostajao. Prethodni E2E test dokazuje da mehanizam
' radi, ali koristi jedinstven broj - pa ne dokazuje IZOLACIJU. Fixture zato
'  ima dve STORNIRANE prijemnice istog broja (8/150326), razlicitih kupaca,
' svaka sa svojom paletom:
'
'   PRJ-TEST-C1  KUP-TEST-1   40 gajbica
'   PRJ-TEST-C2  KUP-TEST-2   25 gajbica
'
' Tako je i u produkciji: BrojPrijemnice se racuna PO KUPCU.
'
' Identitet nosi GeneracijaID - kolonu pravi EnsureSledljivostSchema na svakom
' startu, a pecate je writeri. Fixture redovi su sejani mimo writera, pa im
' test sam upisuje generacije: to je jedini nacin da dva dokumenta budu
' razlucena bas onako kako bi ih razlucio pravi upis.
Private Sub T_RelinkPoGeneraciji_NeDiraTudjDokument()
    Dim upoz As String, gajbDiff As Boolean, ok As Boolean

    ' Preduslov: kolona postoji (Ensure je odradio svoje) i oba dokumenta nose
    ' svoje palete pod ISTIM brojem.
    AssertEq (GetColumnIndex(TBL_PRIJEMNICA, COL_GENERACIJA_ID) > 0), True, _
             "preduslov: EnsureSledljivostSchema je napravio GeneracijaID"
    AssertEq StavkiZaPrijemnicu(FX_PRIJ_KOLIZIJA), 2, _
             "preduslov: dva dokumenta istog broja nose dve paletne stavke"

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-C1", "GEN-TEST-A"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-C2", "GEN-TEST-B"
    AssertEq modDokumenta.GeneracijaPoID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-C1"), _
             "GEN-TEST-A", "preduslov: generacija A je upisana"

    ' CILJ JE ISTO TAKO KOLIZIONO. Broj 1/150326 nose DVE AKTIVNE prijemnice,
    ' PRJ-TEST-A (kupac 1) i PRJ-TEST-B (kupac 2). Raniji oblik ovog testa je tu
    ' pisalo "svejedno koja, bitno je da postoji" -- a nije bilo svejedno: cilj se
    ' birao po golom broju, pa je roba isla onom dokumentu koji je slucajno
    ' POSLEDNJI u tabeli. Izvor po identitetu a cilj po labeli i dalje moze da
    ' odnese palete pogresnom kupcu, samo na drugom kraju.
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-A", "GEN-CILJ-A"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-B", "GEN-CILJ-B"

    ' Prvo FAIL-CLOSED: bez generacije cilja broj je dvosmislen i writer odbija.
    ok = ReassignPaleteToPrijemnica_TX(FX_PRIJ_KOLIZIJA, FX_PRIJ_BROJ, upoz, True, _
                                       gajbDiff, "GEN-TEST-A")
    AssertEq ok, False, "bez generacije CILJA dvosmislen broj se odbija"
    AssertEq (Len(upoz) > 0), True, "odbijanje nosi razlog za operatera"
    AssertEq GajbicaZaDokument("PRJ-TEST-A"), 0, _
             "odbijeno prevezivanje nije nista pomerilo"

    ' Pa po identitetu, na dokument kupca 1.
    ok = ReassignPaleteToPrijemnica_TX(FX_PRIJ_KOLIZIJA, FX_PRIJ_BROJ, upoz, True, _
                                       gajbDiff, "GEN-TEST-A", "GEN-CILJ-A")
    AssertEq ok, True, "prevezivanje po generaciji je proslo"

    ' Tvrdnja na IZVORNOJ strani: dokument B se nije pomerio.
    AssertEq StavkiZaPrijemnicu(FX_PRIJ_KOLIZIJA), 1, _
             "tudji dokument istog broja OSTAJE na svom mestu"
    AssertEq GajbicaZaPrijemnicu(FX_PRIJ_KOLIZIJA), 25, _
             "na starom broju ostaje bas roba dokumenta B (25 gajbica)"

    ' Tvrdnja na CILJNOJ strani: roba je otisla bas izabranom kupcu.
    AssertEq GajbicaZaDokument("PRJ-TEST-A"), 40, _
             "roba je stigla na dokument kupca 1 (40 gajbica)"
    AssertEq GajbicaZaDokument("PRJ-TEST-B"), 0, _
             "dokument drugog kupca istog broja NIJE nista dobio"
End Sub

' ============================================================
' 35. F8: izabran red ostaje izabran do correction context-a
' ============================================================
' Ovo je jedina putanja koju ni jedan owner guard ne pokriva.
'
' Kod modova ISPRAVKA i DUPLI se posle kreiranja context-a zove guarded writer,
' pa dvosmislen broj tamo pukne i context bude obelezen neuspelim. Kod moda
' RESI KASNIJE writer se NE ZOVE UOPSTE -- napravi se samo trajan recovery
' zapis. Ako je dokument razresen po broju, taj zapis moze zauvek da pokazuje
' na TUDJI dokument, i nista to ne prijavljuje.
'
'   PRJ-TEST-A  KUP-TEST-1  \  isti broj 1/150326
'   PRJ-TEST-B  KUP-TEST-2  /  dva aktivna dokumenta
'
' Bira se A. Tvrdnja je da OldDocID u context-u bude BAS A.
' Tvrdnja je namerno napisana tako da NE zavisi od toga koji red je "prvi" u
' tabeli. Prva verzija ovog testa je birala dokument i poredila OldDocID sa njim
' -- i prolazila je i kad se identitet potpuno ignorise, jer je razresavanje po
' broju SLUCAJNO davalo bas taj dokument. Sabotaza je to pokazala; bez nje bi
' test bio placebo.
'
' Zato se meri RAZLIKA U PONASANJU, ne konkretan PK:
'   bez identiteta  -> dvosmislen broj se odbija, recovery zapisa NEMA
'   sa identitetom  -> zapis postoji i pokazuje na izabran dokument
' Prva tvrdnja pada cim se identitet zaobidje, bez obzira na redosled redova.
Private Sub T_F8_IzabranRedOstajeIzabran()
    Dim res As Object, cid As String, ocekivan As String

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z1", "GEN-F8-1"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2", "GEN-F8-2"

    ' Preduslov: broj je stvarno dvosmislen -- inace test ne meri nista.
    AssertEq VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, FX_PRIJ_ZBR_KOLIZIJA, _
                             "T_F8", False, Array(COL_PRJ_KUPAC)).count, 2, _
             "preduslov: broj nose DVA aktivna kupca"

    ' BEZ identiteta: dokument se ne moze utvrditi, pa se ne sme napraviti
    ' trajan recovery zapis. Kod RESI KASNIJE guarded writer se ne zove, pa je
    ' ovo jedina kapija na toj putanji.
    Set res = modStornoDok.StornoIzvrsiMod(STIP_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, "", _
                                           SV_MODE_RESI_KASNIJE, False, False, "")
    AssertEq (Not res Is Nothing), True, "framework je vratio rezultat i bez identiteta"
    AssertEq Len(Trim$(NzToText(res("correctionID")))), 0, _
             "bez identiteta se NE pravi recovery zapis nad dvosmislenim brojem"

    ' SA identitetom: zapis postoji i pokazuje na bas taj dokument.
    Set res = modStornoDok.StornoIzvrsiMod(STIP_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, "", _
                                           SV_MODE_RESI_KASNIJE, False, False, "GEN-F8-2")
    cid = Trim$(NzToText(res("correctionID")))
    AssertEq (Len(cid) > 0), True, "sa identitetom se recovery zapis pravi"

    ocekivan = modDokumenta.GeneracijaPoID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2")
    AssertEq ocekivan, "GEN-F8-2", "preduslov: generacija je upisana na Z2"
    AssertEq OldDocIDKonteksta(cid), "PRJ-TEST-Z2", _
             "recovery zapis pokazuje na IZABRAN dokument"
End Sub

' ============================================================
' 36. Preflight koristi identitet umesto da ga ignorise
' ============================================================
' StornoRazlog je dobio docID pa ga nije koristio. Za novac je zato i dalje
' zvao ResolveNovacForStorno(broj), koji kod dva aktivna reda istog broja kaze
' "treba NovacID" -- iako mu je F8 NovacID upravo poslao. StornoIzvrsi nize je
' vec bio ispravan, ali se do njega nije stizalo: kapija iznad je zaustavljala
' operaciju. Popravka jednog sloja bez drugog izgleda kao da radi.
Private Sub T_Preflight_KoristiIdentitet()
    Dim razlog As String

    ' Bez identiteta: broj je dvosmislen i preflight to kaze.
    razlog = modStornoDok.StornoRazlog(STIP_ISPLATE, FX_NOVAC_DUPLI, "")
    AssertEq (Len(razlog) > 0), True, _
             "bez NovacID-a dvosmislen broj se odbija u preflight-u"

    ' Sa identitetom: pita se BAS taj red, pa nema sta da se razresava.
    razlog = modStornoDok.StornoRazlog(STIP_ISPLATE, FX_NOVAC_DUPLI, "", "NOV-TEST-D2")
    AssertEq razlog, "", "sa NovacID-em preflight propusta izabran red"
End Sub

' ============================================================
' 37. Ispravka prijemnice pod kolizijom broja
' ============================================================
' REZIM RESI KASNIJE je bio identity-aware, a ISPRAVKA i DUPLI su se vracali na
' broj -- pa je owner guard u writeru obarao potpuno legitimnu operaciju.
' Storno je bio bezbedan, ali funkcija nije radila.
Private Sub T_IspravkaPrijemnice_PodKolizijomBroja()
    Dim res As Object

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-I1", "GEN-ISP-1"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-I2", "GEN-ISP-2"
    AssertEq VlasniciPoBroju(TBL_PRIJEMNICA, COL_PRJ_BROJ, FX_PRIJ_ISPRAVKA, _
                             "T_Isp", False, Array(COL_PRJ_KUPAC)).count, 2, _
             "preduslov: broj nose dva dokumenta"

    Set res = modStornoDok.StornoIzvrsiMod(STIP_PRIJEMNICA, FX_PRIJ_ISPRAVKA, "", _
                                           SV_MODE_ISPRAVKA, False, True, "GEN-ISP-1")
    AssertEq (Not res Is Nothing), True, "framework je vratio rezultat"
    AssertEq CBool(res("needsForm")), True, _
             "ISPRAVKA pod kolizijom broja PROLAZI kad je identitet poznat"

    ' Tudji dokument nije ni takao.
    AssertEq (UCase$(Trim$(NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, _
             "PRJ-TEST-I2", COL_STORNIRANO)))) = "DA"), False, _
             "dokument drugog kupca istog broja OSTAJE aktivan"
End Sub

' ============================================================
' 38. Zbirna: zaglavlje po generaciji, kaskada fail-closed
' ============================================================
' Broj zbirne generator drzi jedinstvenim (v. T_Oporavak_CiljneListe), pa ovaj
' test brani RUCNI UNOS. Uz to, identitet se ni tada ne moze provuci do kraja
' lanca -- i to nije previd nego OGRANICENJE SEME: otpremnice, prijemnice i paletne stavke vezuju zbirnu
' KOLONOM BrojZbirne -- ZbirnaID im nije strani kljuc nigde. Deca dva dokumenta
' istog broja su nerazluciva podatkom koji postoji.
'
' Zato: zaglavlje se stornira po generaciji (tacno), a putanje koje bi menjale
' DECU staju kad je broj dvosmislen (postene).
Private Sub T_Zbirna_ZaglavljePoGeneracijiKaskadaStaje()
    Dim ok As Boolean

    StampGeneraciju TBL_ZBIRNA, COL_ZBR_ID, "ZBI-DUPL-1", "GEN-ZB-1"
    StampGeneraciju TBL_ZBIRNA, COL_ZBR_ID, "ZBI-DUPL-2", "GEN-ZB-2"
    AssertEq VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, FX_ZBIRNA_DUPL, "T_Zbr", _
                             False, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count, 2, _
             "preduslov: broj zbirne nose dva aktivna dokumenta"

    ok = StornoZbirna_TX(FX_ZBIRNA_DUPL, "GEN-ZB-2")
    AssertEq ok, True, "zaglavlje izabrane generacije se stornira"
    AssertEq (UCase$(Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, _
             "ZBI-DUPL-2", COL_STORNIRANO)))) = "DA"), True, _
             "izabrana zbirna je stornirana"
    AssertEq (UCase$(Trim$(NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_ID, _
             "ZBI-DUPL-1", COL_STORNIRANO)))) = "DA"), False, _
             "zbirna drugog vozaca istog broja OSTAJE aktivna"
End Sub

' ============================================================
' 39. Otkup bez generacije NE SME da stornira oba otkupna mesta
' ============================================================
' BrojDokumenta otkupa je scoped PO OTKUPNOM MESTU (KIND_OTK, entitet je
' stanica), pa isti broj na dva OM-a postoji legitimno. Writer je do sada bez
' generacije skupljao SVE aktivne redove tog broja -- zatecen zapis bez
' generacije je tako mogao da obori i tudji dokument.
'
' Test je na WRITERU, ne na preflight-u: preflight se moze zaobici (legacy
' forma, kaskada), writer ne moze.
Private Sub T_OtkupBezGeneracije_NeStorniraTudjeOM()
    Dim ok As Boolean, greska As String

    AssertEq VlasniciPoBroju(TBL_OTKUP, COL_OTK_BR_DOK, FX_OTKUP_KOLIZIJA, _
                             "T_Otk", False, Array(COL_OTK_STANICA)).count, 2, _
             "preduslov: isti broj na DVA otkupna mesta"

    ' BEZ generacije -- mora stati, i nista ne sme da se promeni.
    On Error Resume Next
    ok = StornoOtkupByBrDok_TX(FX_OTKUP_KOLIZIJA)
    greska = Err.description
    On Error GoTo 0
    AssertEq ok, False, "bez generacije dvosmislen broj otkupa se odbija"
    AssertEq StorniranoNaID(TBL_OTKUP, COL_OTK_ID, "OTK-KOL-A"), False, _
             "posle odbijanja dokument A nije diran"
    AssertEq StorniranoNaID(TBL_OTKUP, COL_OTK_ID, "OTK-KOL-B"), False, _
             "posle odbijanja dokument B nije diran"

    ' SA generacijom -- prolazi, i dira samo svoj dokument.
    StampGeneraciju TBL_OTKUP, COL_OTK_ID, "OTK-KOL-A", "GEN-OTK-A"
    StampGeneraciju TBL_OTKUP, COL_OTK_ID, "OTK-KOL-B", "GEN-OTK-B"
    AssertEq StornoOtkupByBrDok_TX(FX_OTKUP_KOLIZIJA, "GEN-OTK-A"), True, _
             "sa generacijom storno prolazi"
    AssertEq StorniranoNaID(TBL_OTKUP, COL_OTK_ID, "OTK-KOL-A"), True, _
             "storniran je izabran dokument"
    AssertEq StorniranoNaID(TBL_OTKUP, COL_OTK_ID, "OTK-KOL-B"), False, _
             "dokument drugog otkupnog mesta OSTAJE aktivan"
End Sub

' ============================================================
' 40. "Jedini vlasnik" zbirne se meri DOKUMENTIMA, ne brojevima
' ============================================================
' Zbirna je po invarijanti zbir SVIH svojih aktivnih otpremnica, pa je vise
' otpremnica u jednoj zbirni normalno stanje. Broj otpremnice je scoped po
' stanici, pa dve otpremnice istog broja sa razlicitih stanica u istoj zbirni
' daju JEDAN distinct broj -- i stara provera je tada rekla "jedini vlasnik".
'
' Posledica: PONISTENJE izabrane otpremnice ulazilo bi u punu kaskadu nad
' zbirnom i oborilo i tudju otpremnicu.
Private Sub T_SoleOwner_MeriDokumenteNeBrojeve()
    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-KOL-A", "GEN-OTP-A"
    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-KOL-B", "GEN-OTP-B"

    AssertEq modStornoFlow.OtpremnicaJeJediniVlasnik_Test(FX_ZBIRNA_KASK, _
                             FX_OTPREMNICA_KOLIZIJA, "GEN-OTP-A"), False, _
             "dve otpremnice istog broja u istoj zbirni NISU jedini vlasnik"

    ' Kontrola: kad je stvarno sama, tvrdnja mora biti True -- inace bi test
    ' prolazio i da provera uvek vraca False.
    AssertEq modStornoFlow.OtpremnicaJeJediniVlasnik_Test(FX_ZBIRNA, "1/TEST", ""), _
                                                          True, _
             "jedina otpremnica svoje zbirne JESTE jedini vlasnik"
End Sub

' ============================================================
' 41. Kaskada zbirne staje dok broj nose dva aktivna dokumenta
' ============================================================
' Ovaj test cilja KASKADNU kapiju (PonistiZbirnaChain_TX), a ne onu na nivou
' moda zbirne -- do nje se ovim putem i ne stize. PONISTENJE PRIJEMNICE zove
' kaskadu nad SVOJIM RODITELJEM, pa je to jedini put koji je dohvata.
'
' Svez par (ZBI-KASK-1/2), jer test 38 stornira jedno zaglavlje -- posle njega
' bi ostao jedan aktivan vlasnik i kapija ne bi imala sta da detektuje.
Private Sub T_ZbirnaKaskada_StajeNaDvosmislenom()
    Dim res As Object

    StampGeneraciju TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-1", "GEN-ZB-K1"
    StampGeneraciju TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-2", "GEN-ZB-K2"
    AssertEq VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, FX_ZBIRNA_KASK, "T_Kask", _
                             False, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count, 2, _
             "preduslov: broj nose dva aktivna dokumenta"

    Set res = modStornoDok.StornoIzvrsiMod(STIP_PRIJEMNICA, "2/150326", "", _
                                           SV_MODE_PONISTENJE, True, False, "")
    AssertEq (Not res Is Nothing), True, "framework je vratio rezultat"
    AssertEq CBool(res("success")), False, _
             "ponistenje lanca staje dok broj nose dva aktivna dokumenta"
    ' Sama BEZBEDNOST dolazi od zatecene kapije u StornoZbirna -- kaskada bi
    ' pala i bez moje provere. Ono sto moja provera dodaje je RAZLOG: staje
    ' pre transakcije i kaze operateru sta je problem, umesto generickog
    ' "nije uspelo". Bas to se ovde tvrdi.
    AssertEq (InStr(1, CStr(res("message")), "pripadao VISE vlasnika", _
                    vbTextCompare) > 0), True, _
             "odbijanje imenuje dvosmislen broj, ne samo neuspeh"

    ' Nista nije poniisteno -- ni zaglavlja ni deca.
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-2"), False, _
             "tudja zbirna istog broja nije dirana"
    AssertEq StorniranoNaID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-KOL-B"), False, _
             "otpremnica tudjeg dokumenta nije dirana"
End Sub

' ============================================================
' 42. Zamena zbirne ne sme da odnese decu TUDJE zbirne
' ============================================================
' Ovo je najtisi kvar u celom lancu. Pocetak ISPRAVKE je tacan: zaglavlje se
' stornira po generaciji, tudje ostaje aktivno. Ali CompleteZbirnaIspravka --
' koja se izvrsava TEK POSLE snimanja zamene -- prevezuje otpremnice i
' prijemnice po BrojZbirne, jer drugog kljuca u semi nema.
'
' Ishod bi bio: storniram tacno SVOJE zaglavlje, pa TUDJOJ zbirni odnesem decu.
' Nista ne izgleda pokvareno u trenutku storna.
'
' Dok child mutacije ne budu scoped, jedina postena opcija je stati PRE nego
' sto se ista promeni -- i to je ono sto se ovde tvrdi.
Private Sub T_ZamenaZbirne_NeDiraDecuTudje()
    Dim res As Object

    AssertEq VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, FX_ZBIRNA_KASK, "T_Zam", _
                             False, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count, 2, _
             "preduslov: broj nose dva aktivna dokumenta"
    AssertEq ZbirnaNaOtpremnici("OTP-KOL-B"), FX_ZBIRNA_KASK, _
             "preduslov: tudja otpremnica visi na tom broju"

    Set res = modStornoFlow.RunZbirnaCorrection(FX_ZBIRNA_KASK, SV_MODE_ISPRAVKA, _
                                                False, "GEN-ZB-K1")
    AssertEq CBool(res("success")), False, _
             "ISPRAVKA staje dok broj nose dva aktivna dokumenta"
    AssertEq CBool(res("needsForm")), False, _
             "forma za zamenu se NE otvara -- inace bi zamena stigla do relinka"

    ' Nista nije dirano: ni izabrano zaglavlje, ni tudje, ni deca.
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-1"), False, _
             "izabrano zaglavlje nije stornirano pre nego sto se zna da relink moze"
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-2"), False, _
             "tudje zaglavlje nije dirano"
    AssertEq ZbirnaNaOtpremnici("OTP-KOL-B"), FX_ZBIRNA_KASK, _
             "tudja otpremnica je OSTALA na svojoj zbirni"
End Sub

' ============================================================
' 43. Zavrsetak ispravke: tacan OldDocID se NE degradira na broj
' ============================================================
' Zatecen dokument nema GeneracijaID. Completion je iz correction context-a
' citao OldDocID, pa GeneracijaPoID vracao "" -- i onda je prazan opseg znacio
' "izaberi po poslovnom broju". Broj otpremnice je scoped PO STANICI, pa su
' blokovi dokumenta sa DRUGE stanice mogli da udju u relink.
'
' OldDocID je bio tacan sve vreme; gubio se jedan korak kasnije.
Private Sub T_ZavrsetakIspravke_NeDegradiraOldDocID()
    Dim cid As String, res As Object

    AssertEq modDokumenta.GeneracijaPoID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-LEG-A"), "", _
             "preduslov: zatecen dokument NEMA generaciju"
    AssertEq VlasniciPoBroju(TBL_OTPREMNICA, COL_OTP_BROJ, FX_OTPREMNICA_LEGACY, _
                             "T_Zav", False, Array(COL_OTP_STANICA)).count, 2, _
             "preduslov: isti broj na dve stanice"

    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_OTPREMNICA, _
                                                   "OTP-LEG-A", FX_OTPREMNICA_LEGACY)
    AssertEq (Len(cid) > 0), True, "correction context je napravljen"

    Set res = modStornoFlow.CompleteOtpremnicaIspravka(cid, FX_OTPREMNICA_ZAMENA)
    AssertEq (Not res Is Nothing), True, "completion je vratio rezultat"

    ' DVA SMERA. Sama tvrdnja "tudji blok nije pomeren" prolazi i kod verzije
    ' koja ne preveze NIJEDAN blok, pa uz nju ide i pozitivna kontrola.
    AssertEq CBool(res("success")), True, "zavrsetak ispravke je uspeo"
    AssertEq OtpremnicaNaBloku("OTK-LEG-A"), "OTP-LEG-N", _
             "MOJ blok JESTE prevezan na zamensku otpremnicu"
    AssertEq OtpremnicaNaBloku("OTK-LEG-B"), "OTP-LEG-B", _
             "blok dokumenta sa druge stanice OSTAJE na svojoj otpremnici"
End Sub

' ============================================================
' 44. Storniran vlasnik nestaje iz racuna, njegova deca ne
' ============================================================
' StornoZbirna_TX stornira SAMO redove tblZbirna -- otpremnice, prijemnice i
' palete ne dira. Zato je ovo dostizno stanje, ne teorija:
'
'   Zbirna A  STORNIRANA   ali OTP-A jos AKTIVNA
'   Zbirna B  AKTIVNA      isti broj
'
' Sa brojanjem samo AKTIVNIH vlasnika, izbor B daje "broj je jednoznacan", pa
' detach i kaskada -- koje idu PO BROJU -- odvezu i decu stornirane A.
Private Sub T_StorniranVlasnik_JosImaAktivnuDecu()
    Dim res As Object

    ' Korak 1: storniraj SAMO zaglavlje A.
    AssertEq StornoZbirna_TX(FX_ZBIRNA_KASK, "GEN-ZB-K1"), True, _
             "zaglavlje A je stornirano"
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-1"), True, _
             "A je stornirana"
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-2"), False, _
             "B je ostala aktivna"

    ' Korak 2: dete stornirane A je i dalje AKTIVNO -- to je cela poenta.
    AssertEq StorniranoNaID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-KOL-A"), False, _
             "dete stornirane zbirne je i dalje aktivno"

    ' Korak 3: operacija nad B koja dira DECU mora da stane, iako je sada
    ' samo jedan AKTIVAN vlasnik tog broja.
    Set res = modStornoDok.StornoIzvrsiMod(STIP_ZBIRNA, FX_ZBIRNA_KASK, "", _
                                           SV_MODE_DUPLI, True, False, "GEN-ZB-K2")
    AssertEq CBool(res("success")), False, _
             "DUPLI staje jer broj je IKAD pripadao dvama vlasnicima"
    ' Ishod cuvaju DVE nezavisne kapije (na nivou moda i u detach-u), pa ga
    ' jedna sabotaza ne moze oboriti. Zato se tvrdi i KOJA je stala: kapija
    ' na nivou moda staje PRE transakcije i objasnjava razlog, dok bi detach
    ' pukao iznutra i dao samo "Storno zbirne nije uspeo".
    AssertEq (InStr(1, CStr(res("message")), "Zamena bi prevezala decu", _
                    vbTextCompare) > 0), True, _
             "staje kapija na nivou moda, pre transakcije, sa razlogom"

    ' Nista nije odvezano ni stornirano.
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-KASK-2"), False, _
             "B zaglavlje nije dirano"
    AssertEq ZbirnaNaOtpremnici("OTP-KOL-A"), FX_ZBIRNA_KASK, _
             "dete stornirane A nije odvezano"
    AssertEq ZbirnaNaOtpremnici("OTP-KOL-B"), FX_ZBIRNA_KASK, _
             "dete aktivne B nije odvezano"
End Sub

' ============================================================
' 45. Otpremnica ne sme da mutira dvosmislenu RODITELJSKU zbirnu
' ============================================================
' Identitet same otpremnice je bio resen, ali su ISPRAVKA/DUPLI/PONISTENJE svi
' dirali RODITELJSKU zbirnu PO BROJU: rekalkulacija, storno prazne zbirne, i
' relink njenih prijemnica u completion-u. Kad broj roditelja nije jednoznacan,
' nijedno od toga ne zna cije je.
'
' RecalculateZbirnaFromOtpremnice_TX je posebno podmukao: sabira otpremnice po
' broju, pa tim zbirom azurira JEDAN nadjen red -- moglo je da rekalkulise
' zaglavlje B vrednostima koje ukljucuju otpremnice oba dokumenta.
Private Sub T_OtpremnicaNadDvosmislenomZbirnom_Staje()
    Dim res As Object

    AssertEq VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, FX_ZBIRNA_KASK, "T_Otp", _
                             True, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count, 2, _
             "preduslov: broj roditeljske zbirne je pripadao dvama vlasnicima"
    AssertEq ZbirnaNaOtpremnici("OTP-KOL-A"), FX_ZBIRNA_KASK, _
             "preduslov: otpremnica visi na tom broju"

    Set res = modStornoDok.StornoIzvrsiMod(STIP_OTPREMNICA, FX_OTPREMNICA_KOLIZIJA, "", _
                                           SV_MODE_DUPLI, True, False, "GEN-OTP-A")
    AssertEq CBool(res("success")), False, _
             "DUPLI staje kad je broj roditeljske zbirne dvosmislen"
    AssertEq (InStr(1, CStr(res("message")), "roditeljske zbirne", vbTextCompare) > 0), _
             True, "razlog imenuje RODITELJSKU zbirnu, ne samo neuspeh"

    ' Nijedna otpremnica nije stornirana -- staje se PRE mutacije.
    AssertEq StorniranoNaID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-KOL-A"), False, _
             "izabrana otpremnica nije stornirana"
    AssertEq StorniranoNaID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-KOL-B"), False, _
             "tudja otpremnica nije dirana"
End Sub

' ============================================================
' 46. Zatecen context ne sme da preveze prijemnice tudje zbirne
' ============================================================
' Kapija na startu ne pomaze za context koji je napravljen PRE nje: correction
' context je persistentan i prezivljava upgrade. Zato completion pita ponovo.
'
' Bez toga bi CompleteOtpremnicaIspravka skupila prijemnice po oldZbirna BROJU
' i prevezala i tudje na novu zbirnu.
Private Sub T_ZatecenContext_NePrevezujeTudjePrijemnice()
    Dim cid As String, res As Object

    ' Scenario mora imati DVA razlicita roditelja pod istim brojem otpremnice,
    ' inace test ne meri nista: ako oba siblinga vise na istoj zbirni, lookup po
    ' broju slucajno daje tacan odgovor i kapija prolazi i kad je kod pogresan.
    AssertEq ZbirnaNaOtpremnici("OTP-STL-A"), FX_ZBIRNA_KASK, _
             "preduslov: izabrana otpremnica visi na DVOSMISLENOJ zbirni"
    AssertEq ZbirnaNaOtpremnici("OTP-STL-B"), FX_ZBIRNA_STALE, _
             "preduslov: sibling istog broja visi na JEDNOZNACNOJ zbirni"
    AssertEq PrviRoditeljPoBroju(FX_OTPREMNICA_STALE), FX_ZBIRNA_STALE, _
             "preduslov: prvi red po broju daje POGRESNOG roditelja"
    ' Tvrdnja se meri nad NAMENSKOM prijemnicom. Naslanjanje na PRJ-KASK-1 je
    ' tvrdnju cinilo vakuumskom -- do ovog testa je vec bila u drugom stanju, pa
    ' relink nije imao sta da preveze i sabotaza je obarala samo tudju tvrdnju.
    AssertEq ZbirnaNaPrijemnici("PRJ-STL-T"), FX_ZBIRNA_KASK, _
             "preduslov: tudja prijemnica visi na dvosmislenoj zbirni"
    AssertEq StorniranoNaID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-STL-T"), False, _
             "preduslov: tudja prijemnica je AKTIVNA (inace relink nema sta da uzme)"
    AssertEq ZbirnaNaOtpremnici("OTP-STL-N"), FX_ZBIRNA_STALE, _
             "preduslov: cilj zamene ima NEPRAZNU zbirnu (inace relinka nema)"
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-STL-1"), False, _
             "preduslov: ciljna zbirna je aktivna"

    ' Context kakav pravi produkcija: OldDocID + OldBroj + ParentBroj. Napravljen
    ' je "pre kapije" -- persistentan je i prezivljava upgrade, pa kapija na
    ' startu za njega nije ni postojala.
    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_OTPREMNICA, _
                                                  "OTP-STL-A", FX_OTPREMNICA_STALE, _
                                                  "", "", "", "", "", FX_ZBIRNA_KASK)
    AssertEq (Len(cid) > 0), True, "zatecen context je napravljen"

    Set res = modStornoFlow.CompleteOtpremnicaIspravka(cid, FX_OTPREMNICA_STALE_NOVA)

    ' POSLOVNA TVRDNJA IDE PRVA. AssertEq puca na prvom padu, pa tvrdnja koju
    ' sabotaza navodi mora biti i prva koja pada -- inace sabotaza obori tvrdnju
    ' o success-u, test se prekine, a ova ostane nemerena. Tako je i izgledalo
    ' zeleno: "prijemnica nije prevezana" se nije ni izvrsavalo.
    AssertEq ZbirnaNaPrijemnici("PRJ-STL-T"), FX_ZBIRNA_KASK, _
             "tudja prijemnica NIJE prevezana na novu zbirnu"
    AssertEq CBool(res("success")), False, _
             "completion staje nad zatecenim context-om dvosmislene zbirne"
    AssertEq (InStr(1, CStr(res("message")), "stare zbirne", vbTextCompare) > 0), True, _
             "razlog imenuje staru zbirnu"
End Sub

' ============================================================
' 53. Kes tabela ne sme da memoise NEUSPEH
' ============================================================
' Zatecen incident sa prave instalacije: operater je gledao PRAZNE liste za svaki
' tip dokumenta, bez ijedne greske, dok je tblOtkup bio pun. Dijagnostika:
'
'     IsArray(modUiData.CachedTable("tblOtkup"))  -> False
'     IsArray(GetTableData("tblOtkup"))           -> True
'
' Tabela je pri PRVOM citanju bila prazna (podaci stizu posle -- sync, uvoz,
' legacy forma), pa je Empty ostao u kesu do kraja sesije. ResetCache se zove
' samo pri gradnji ekrana i posle upisa kroz novi UI. Potvrda: rucni
' "modUiData.ResetCache: modOtkupUI.RefreshFromData" je vratio sve liste.
'
' Vraceno stanje se NE moze izmeriti kroz vracenu vrednost -- ona je Empty i kad
' je neuspeh kesiran i kad nije. Zato test gleda SAM KES, kroz seam.
Private Sub T_KesTabela_NeMemoiseNeuspeh()
    modUiData.ResetCache

    AssertEq IsArray(modUiData.CachedTable("tblNePostojiUopste")), False, _
             "necitljiva tabela vraca Empty"
    AssertEq modUiData.KesImaKljuc("tblNePostojiUopste"), False, _
             "neuspeh se NE kesira -- inace tabela ostaje prazna do kraja sesije"

    ' Pozitivna kontrola: uspeh se i dalje kesira, inace bi ispravka znacila
    ' skeniranje tabele pri svakom crtanju mreze.
    AssertEq IsArray(modUiData.CachedTable(TBL_OTKUP)), True, _
             "citljiva tabela se procita"
    AssertEq modUiData.KesImaKljuc(TBL_OTKUP), True, "uspeh se kesira"

    ' Necitljiva tabela mora da se RAZLIKUJE od prazne.
    AssertEq modUiData.TabelaCitljiva("tblNePostojiUopste"), False, _
             "nepostojeca tabela nije citljiva"
    AssertEq modUiData.TabelaCitljiva(TBL_OTKUP), True, "postojeca tabela je citljiva"
End Sub

' ============================================================
' 54. Mapa imena: kljuc nosi kolone, prazna se ne kesira
' ============================================================
' Drugi simptom istog korena: u listi je stajalo "KOOP-00022" umesto imena.
' PartnerMap je zvao isti kes, dobio Empty, napravio PRAZAN recnik -- i kesirao
' ga. Posle toga je svako ime do kraja sesije padalo na goli ID.
'
' Uz to je kljuc kesa bio SAMO ime tabele, iako mapa zavisi i od kolona: prvi
' pozivalac je time odlucivao sta svi ostali dobijaju.
Private Sub T_MapaImena_KljucNosiKolone()
    Dim samoIme As Object, imePrezime As Object

    modUiData.ResetCache
    ' Isti ID, dve razlicite kolone imena. Sa kljucem samo po tabeli drugi poziv
    ' vraca prvi recnik, pa su vrednosti identicne -- a nisu iste stvari.
    Set samoIme = modOtkupUI.PartnerMap(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "")
    Set imePrezime = modOtkupUI.PartnerMap(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    AssertEq (samoIme.count > 0), True, "preduslov: mapa kooperanata nije prazna"
    AssertEq (CStr(imePrezime("KOOP-TEST-1")) <> CStr(samoIme("KOOP-TEST-1"))), True, _
             "kljuc kesa nosi KOLONE -- ime+prezime nije isto sto i samo ime"

    ' Prazna mapa (nepostojeca tabela) ne sme da se kesira.
    AssertEq modOtkupUI.PartnerMap("tblNePostojiUopste", "X", "Y", "").count, 0, _
             "nepostojeca tabela daje praznu mapu"
    AssertEq modOtkupUI.MapaImenaKesirana("tblNePostojiUopste|X|Y|"), False, _
             "prazna mapa se NE kesira -- inace svako ime pada na goli ID"
End Sub

' ============================================================
' 55. Storno je EKRAN, ne unosni rezim
' ============================================================
' F8 je crtao celu unosnu formu i primarno dugme "Storniraj dokument", a
' Scr_Save za STORNO je padao u Case Else i vracao "Nije vezano na postojecu
' rutinu". Dakle dugme je bilo mrtvo, a forma je pozivala operatera da ukuca
' podatke dokumenta koji hoce da stornira. #201 je to sakrio grid-maxom --
' privremeno, jer je forma i dalje postojala i samo se nije videla.
'
' Od v6-ui-142 forme nema: storno je ekran u registru, sa "upis=ne". Ovaj test
' zamenjuje raniji T_StornoNijeUnosniRezim, koji je merio grid-max -- meru koja
' je sa rezimom prestala da postoji.
Private Sub T_StornoJeEkranNeRezim()
    ' NAJVAZNIJE PRVO (AssertEq dize gresku, pa se test prekida na prvom padu):
    ' ekran mora da postoji u registru i da odgovara na ugovor. Ako se ime modula
    ' u registru omakne, sidebar ga samo prikaze prigusenog -- i storno nestane
    ' iz aplikacije bez ijedne greske.
    AssertEq modUiScreens.ScrPostoji("STORNO"), True, _
             "modul ekrana Storno odgovara na Scr_Meta (kasno vezivanje radi)"
    AssertEq (InStr(modUiScreens.ScrMeta("STORNO"), "upis=ne") > 0), True, _
             "Storno nema upis -- forma i primarno dugme mu ne pripadaju"

    ' F8 vise NIJE rezim. modeKey za nepoznat rezim pada u Case Else ("OTKUP"),
    ' pa se odsustvo meri time da vise ne daje "STORNO".
    AssertEq (modScrDokumenti.modeKey("F8") <> "STORNO"), True, _
             "F8 vise ne razresava u rezim STORNO"
    AssertEq (InStr(modUiScreens.ScrMeta("DOKUMENTI"), "rezima=7") > 0), True, _
             "unosni ekran ima SEDAM rezima, ne osam"

    ' Odluka se donosi u zoni, uz posledice -- ne dugmetom u redu mreze, koje bi
    ' vodilo pravo u izvrsenje.
    AssertEq modScrStorno.Scr_Radnje(), "", _
             "Storno nema radnju nad redom -- odluka je u zoni, uz posledice"
End Sub

' ============================================================
' 56. Ugovor ekrana Storno: deset lista, prva je navigaciona
' ============================================================
' Isti oblik kao T_Oporavak_UgovorIRadnje. Devet tipova su preneti iz F8; deseta
' ("Svi") je pogled preko tipova koji legacy ima kao "Nadji dokument", a novi UI
' do v6-ui-142 nije imao.
Private Sub T_Storno_UgovorIRadnje()
    Dim liste As Variant, i As Long, kljucevi As String

    AssertEq (Len(modUiScreens.ScrRowByKey("STORNO")) > 0), True, _
             "STORNO postoji u registru ekrana"
    AssertEq (InStr(modUiScreens.ScrMeta("STORNO"), "kljuc=STORNO") > 0), True, _
             "Scr_Meta prijavljuje svoj kljuc"

    liste = modScrStorno.Scr_Liste()
    ' NAJVAZNIJE PRVO: ljuska mora da nacrta SVE liste koje ekran prijavi.
    ' MAX_SEG je bio 9 dok ih ekran ima 10, pa je "Izvodi" tiho nestajao --
    ' LayoutGrid nacrta prvih MAX_SEG i stane, bez greske i bez traga. Tvrdnja
    ' vazi za svaki buduci ekran, ne samo za ovaj.
    AssertEq (UBound(liste) + 1 <= modOtkupUI.MaxPrekidaca()), True, _
             "ljuska crta sve liste ekrana -- nijedna se ne odseca tiho"
    AssertEq (UBound(liste) + 1), 10, "ekran ima deset lista"
    For i = 0 To UBound(liste)
        kljucevi = kljucevi & "|" & Split(CStr(liste(i)), "|")(0)
    Next i
    AssertEq kljucevi, "|SVI|OTKUP|OTPREMNICA|ZBIRNA|PRIJEMNICA|AMB_ISPLATE|" & _
             "AMB_UPLATE|REVERSI|FAKTURA|IZVOD", _
             "redosled i kljucevi lista -- 'Svi' je prva"

    ' Kljucevi lista JESU kljucevi tipova (STIP_*), pa prevodne tabele nema.
    ' Ako se razidju, Scr_Rows bi trazio tabelu za nepostojeci tip i tiho vratio
    ' otkupne listove pod tudjim naslovom.
    For i = 1 To UBound(liste)
        AssertEq (Len(modScrDokumenti.TabelaTipa(Split(CStr(liste(i)), "|")(0))) > 0), True, _
                 "lista " & Split(CStr(liste(i)), "|")(0) & " ima svoju tabelu"
    Next i
End Sub

' ============================================================
' 57. Ekran Storno isporucuje IDENTITET, ne samo broj
' ============================================================
' Ovo je najskuplja tvrdnja u celoj migraciji storna u svoj ekran.
'
' Nevidljiva kolona identiteta se do v6-ui-141 dodavala pod uslovom
' "If modOtkupUI.ActiveMode = "F8"". Ekran nema rezim -- da je taj uslov ostao,
' bio bi cutke False, kolona bi nestala, IdentIzReda bi vracao prazno, i ceo
' lanac iz #198 (correctionID / OldDocID / GeneracijaID) bi pao na fail-closed
' po broju. NIJEDNA postojeca suite to ne bi videla: testovi identiteta (35, 45,
' 46, 48-52) mere sloj ISPOD mreze, kome se docID prosledjuje direktno.
'
' Zato se ovde meri bas spoj: da li opis kolona koji dobija MREZA nosi kolonu
' identiteta, i da li je unosni ekran i dalje NE nosi.
Private Sub T_StornoEkran_KolonaIdentiteta()
    Dim cols As Variant, poslednja As String, i As Long, ima As Boolean

    ' NAJVAZNIJE PRVO: kolona identiteta MORA biti tu kad se trazi.
    cols = modScrDokumenti.GridCols(STIP_PRIJEMNICA, True)
    poslednja = modScrDokumenti.ColF(CStr(cols(UBound(cols))), 1)
    AssertEq poslednja, COL_GENERACIJA_ID, _
             "opis kolona za Storno nosi kolonu identiteta, i to POSLEDNJU"

    ' I mora biti NEVIDLJIVA: prioritet 4, dok petlja vidljivosti ide 3 -> 1.
    AssertEq modScrDokumenti.ColF(CStr(cols(UBound(cols))), 4), "4", _
             "kolona identiteta je prioriteta 4 -- nikad vidljiva"

    ' Unosni ekran je NE sme dobiti: GridCols je zajednicki za rezim unosa i za
    ' Storno nad istim tipom (F4 i Storno/Prijemnica daju isti kljuc), pa bi
    ' bezuslovno dodavanje menjalo i mrezu unosa.
    cols = modScrDokumenti.GridCols(STIP_PRIJEMNICA, False)
    For i = 0 To UBound(cols)
        If modScrDokumenti.ColF(CStr(cols(i)), 1) = COL_GENERACIJA_ID Then ima = True
    Next i
    AssertEq ima, False, "unosni rezim NE dobija kolonu identiteta"

    ' Tipovi koji identitet nemaju (revers, izvod) ga i ne dobijaju -- kolona bez
    ' izvora bi mrezi dala prazan string koji izgleda kao "zatecen zapis".
    cols = modScrDokumenti.GridCols(STIP_REVERSI, True)
    ima = False
    For i = 0 To UBound(cols)
        If modScrDokumenti.ColF(CStr(cols(i)), 1) = COL_GENERACIJA_ID Then ima = True
    Next i
    AssertEq ima, False, "revers nema kanonski identitet, pa ni kolonu"

    ' I na kraju: ono sto ekran zapamti pri izboru reda je ono sto salje nizvodno.
    modScrStorno.Scr_IzborTestSet STIP_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, "GEN-F8-2", ""
    AssertEq modScrStorno.Scr_IzabranDocID(), "GEN-F8-2", _
             "ekran nosi identitet izabranog reda, ne samo broj"
End Sub

' ============================================================
' 58. Svaka lista ekrana Storno stvarno vraca redove
' ============================================================
' Operater je prijavio da cip "Svi" "nema funkciju": klik ne menja ni mrezu ni
' naslov. Uzrok nije bio u prekidacu nego DVA sloja nize, i bio je nevidljiv:
'
'   modUiScreens.ScrGridData ima "On Error Resume Next", pa greska iz Scr_Rows
'   ne stigne do ReloadGrid nego se vrati kao Empty. LoadGridFromScreen na
'   ne-niz radi "Exit Sub" -- i mreza OSTANE na prethodnoj listi, sa prethodnim
'   naslovom. Nema greske, nema toasta, izgleda kao da dugme ne radi.
'
' Zato ovaj test zove Scr_Rows za SVAKU listu direktno, mimo tog gutaca: ako
' neka pukne, ovde pukne po imenu liste umesto da tiho ne uradi nista.
Private Sub T_StornoEkran_SvakaListaVracaRedove()
    Dim liste As Variant, i As Long, kljuc As String, d As Variant

    liste = modScrStorno.Scr_Liste()
    For i = 0 To UBound(liste)
        kljuc = Split(CStr(liste(i)), "|")(0)
        modScrStorno.Scr_TipTestSet kljuc
        AssertEq modScrStorno.Scr_Lista(), kljuc, "lista " & kljuc & " je izabrana"
        d = modScrStorno.Scr_Rows("sve", "")
        AssertEq IsArray(d), True, "lista " & kljuc & " vraca niz, ne Empty"
        AssertEq (UBound(d) >= 4), True, _
                 "lista " & kljuc & " vraca pun oblik (kolone, redovi, n, kg, val)"
        AssertEq IsArray(d(0)), True, "lista " & kljuc & " vraca opis kolona"
    Next i

    modScrStorno.Scr_TipTestSet STIP_OTKUP
End Sub

' ============================================================
' 59. Prefill bez broja MORA da predlozi broj
' ============================================================
' Posle ispravke je forma bila popunjena, a BROJ OTPREMNICE prazan. Prefill ga
' namerno ne donosi -- stari broj pripada storniranom dokumentu, novi mora da
' dobije svoj -- ali predlog se ni ne racuna:
'
'   RefreshBrojPredlog visi o promeni STANICE ili DATUMA, a prefill oba
'   postavlja pod "mLoading = True", pa se nijedan event ne okine.
'   SelectModeCore ga zove ranije, ali tada stanice jos nema (forma je tek
'   ocisceno), pa EntitetZaBroj vrati prazno i predlog se preskoci.
'
' Rezultat: dokument koji operater treba samo da potvrdi ostaje bez broja, i to
' bez ijedne poruke. Ovaj test meri POSLEDICU (polje je popunjeno), ne put.
Private Sub T_PrefillBezBroja_PredlaziBroj()
    Dim f As frmOtkupUI, zf As Object, broj As String

    Set f = NewOtkupUIForm()
    ' Combo-i moraju biti PUNJENI: generator broja cita stanicu iz cbOM
    ' (EntitetZaBroj -> GetComboID), a SetComboByID ne moze da izabere stavku u
    ' praznoj listi. U produkciji ih puni StartApp; u testu se forma gradi bez
    ' .Show, pa se punjenje trazi izricito.
    modOtkupUI.FillCombos f
    modOtkupUI.SelectMode f, "F2"
    Set zf = f.Controls("zForm")

    ' Preduslov: polje je prazno pre prefilla -- inace test meri zatecenu
    ' vrednost umesto onoga sto prefill uradi.
    zf.Controls("fgBrOtpr").Controls("fgBrOtprT").text = ""

    ' Spec BEZ "brdok", sa stanicom i datumom -- tacno ono sto ispravka salje.
    modOtkupUI.ApplyPrefill "datum=" & Format$(Date, "dd.mm.yyyy") & _
                            "|omid=" & FX_STANICA & "|vrsta=" & FX_VRSTA

    ' PREDUSLOV: stanica je stvarno izabrana. Generator broja je cita iz combo-a
    ' (EntitetZaBroj -> GetComboID), pa bez nje ne bi bilo predloga ni kad je
    ' pravilo ispravno -- test bi merio prazan combo umesto pravila.
    AssertEq (Len(Trim$(CStr(f.Controls("zCtx").Controls("cbOM").value))) > 0), True, _
             "preduslov: prefill je izabrao otkupno mesto"

    broj = Trim$(CStr(zf.Controls("fgBrOtpr").Controls("fgBrOtprT").text))
    AssertEq (Len(broj) > 0), True, _
             "prefill bez broja predlaze broj dokumenta za svoj kontekst"

    ' Suprotan smer: kad prefill DONESE broj, predlog ga ne sme pregaziti --
    ' inace bi izbor otpremnice u F1 gubio broj koji je stigao uz nju.
    modOtkupUI.ApplyPrefill "datum=" & Format$(Date, "dd.mm.yyyy") & _
                            "|omid=" & FX_STANICA & "|brdok=TEST-BR-1"
    AssertEq Trim$(CStr(zf.Controls("fgBrOtpr").Controls("fgBrOtprT").text)), "TEST-BR-1", _
             "broj koji prefill donese se ne pregazuje predlogom"
End Sub

' ============================================================
' 52. Prost F8 storno zbirne mora da se IZVRSI, ne samo da postoji
' ============================================================
' Ovaj test postoji zbog compile greske koja je zivela od v6-ui-119 i koju je
' nasao operater rucnim Debug > Compile, a ne suite:
'
'   poruka = Poruka("STORNO_MSG_ZBIRNA_PRIJ")   ' Expected array
'
' Izlazni parametar procedure se zove "poruka", VBA je case-insensitive, pa je
' nekvalifikovan poziv postao indeksiranje tog String parametra. Nijedna suite
' nije zvala StornoIzvrsi, a VBA proceduru kompajlira TEK KAD SE POZOVE -- zato
' je 51 zelen test mirno stajao nad kodom koji se ne kompajlira.
'
' Zato ovaj test ne meri samo poruku: on tu proceduru IZVRSAVA. To je jedini
' nacin da compile greska u njoj postane crvena suite, a ne tek nalaz operatera.
Private Sub T_StornoIzvrsi_ZbirnaImenujeVezanuPrijemnicu()
    Dim ok As Boolean, msg As String

    AssertEq ZbirnaNaPrijemnici("PRJ-OLD-U"), FX_ZBIRNA_OLDU, _
             "preduslov: aktivna prijemnica visi na toj zbirni"
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-OLDU-1"), False, _
             "preduslov: zbirna je aktivna"

    ok = modStornoDok.StornoIzvrsi(STIP_ZBIRNA, FX_ZBIRNA_OLDU, "", msg, "")

    AssertEq ok, True, "prost storno zbirne je prosao"
    ' StornoZbirna namerno NE kaskadira, pa prijemnica ostaje vezana za storniranu
    ' zbirnu. Operater to mora da vidi, inace mu sledljivost visi bez upozorenja.
    AssertEq (InStr(1, msg, FX_PRIJEMNICA_OLD_U, vbTextCompare) > 0), True, _
             "poruka imenuje prijemnicu koja je ostala vezana"
End Sub

' ============================================================
' 50. Spisak blokova za F8 je po IDENTITETU, ne po broju
' ============================================================
' ActiveBlocksForFlow je za otpremnicu radio GetOtpremnicaIDsByBroj(broj) bez
' generacije -- pa je spisak sadrzao blokove SVIH dokumenata tog broja. Isti
' BrojOtpremnice na dve stanice je legitiman, sto ostatak ovog PR-a i modeluje.
Private Sub T_BlokoviF8_PoIdentitetu()
    Dim po As Collection, sviRedovi As Collection

    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-A", "GEN-BLK-A"
    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-B", "GEN-BLK-B"

    ' Scenario je stvaran samo ako broj sam po sebi daje OBA bloka.
    Set sviRedovi = modStornoFlow.GetStornoBlockRows(FLOW_DOC_OTPREMNICA, _
                                                    FX_OTPREMNICA_BLOK, "", "")
    AssertEq sviRedovi.count, 2, _
             "preduslov: po golom broju spisak nosi blokove OBA dokumenta"

    Set po = modStornoFlow.GetStornoBlockRows(FLOW_DOC_OTPREMNICA, FX_OTPREMNICA_BLOK, _
                                             "", "GEN-BLK-A")
    AssertEq po.count, 1, "sa identitetom spisak nosi SAMO blok izabranog dokumenta"
    AssertEq CStr(po(1)(0)), "OTK-BLK-A", "i to bas njegov blok"

    ' Isti kvar je bio i u PREGLEDU: ScanOtpremnica razresi dokument po identitetu
    ' pa blockCount racuna po broju. Operater bi video tudje blokove, a correction
    ' dijalog bi se otvorio i nad dokumentom koji blokove nema.
    Dim pregled As String
    pregled = modStornoDok.StornoPregledLanca(STIP_OTPREMNICA, FX_OTPREMNICA_BLOK, _
                                              "", "GEN-BLK-A")
    AssertEq (InStr(1, pregled, "Otkupni blokovi: 1", vbTextCompare) > 0), True, _
             "pregled broji blokove IZABRANOG dokumenta, ne svih tog broja"
End Sub

' ============================================================
' 51. Storniran sibling ne sme da izgubi svoj blok
' ============================================================
' Ovo je mutacija, ne pregled. Kapija BlockStornoDriftReason tu ne pomaze: prva
' linija joj je "If ModeStornoBlokParent(docType, mode) Then Exit Function", a to
' je True za svaki PONISTENJE i za OTPREMNICA+DUPLI/ISPRAVKA -- dakle za tacno
' one modove koji jedini stizu do dodatnog storna blokova. Njena pretpostavka
' ("roditelj umire, pa je blok-storno bezbedan") vazi samo za blokove IZABRANOG
' dokumenta.
'
' Test radi ono sto radi UI posle uspesnog moda: uzme spisak blokova i stornira
' ga. Sam StornirajBlokoveAko se ne moze zvati iz testa (MsgBox), pa se meri
' sloj ispod -- ista dva poziva, bez dijaloga.
Private Sub T_StorniranSibling_ZadrzavaSvojBlok()
    Dim res As Object, redovi As Collection, ids As Collection, i As Long

    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-A", "GEN-BLK-A"
    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-B", "GEN-BLK-B"
    AssertEq StorniranoNaID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-B"), True, _
             "preduslov: sibling je STORNIRAN (pa nema zivog roditelja za kapiju)"
    AssertEq StorniranoNaID(TBL_OTKUP, COL_OTK_ID, "OTK-BLK-B"), False, _
             "preduslov: blok siblinga je i dalje AKTIVAN"

    Set res = modStornoDok.StornoIzvrsiMod(STIP_OTPREMNICA, FX_OTPREMNICA_BLOK, "", _
                                           SV_MODE_DUPLI, True, False, "GEN-BLK-A")
    AssertEq CBool(res("success")), True, "DUPLI nad izabranom otpremnicom je prosao"

    ' Dodatni storno blokova -- isto sto UI radi posle uspesnog moda.
    Set redovi = modStornoFlow.GetStornoBlockRows(FLOW_DOC_OTPREMNICA, FX_OTPREMNICA_BLOK, _
                                                 "", "GEN-BLK-A")
    Set ids = New Collection
    For i = 1 To redovi.count
        ids.Add CStr(redovi(i)(0))
    Next i
    If ids.count > 0 Then modStornoFlow.StornoSelectedBlocks_TX ids

    ' Poslovna tvrdnja PRVA (v. zamka 6 u sabotaza.py).
    AssertEq StorniranoNaID(TBL_OTKUP, COL_OTK_ID, "OTK-BLK-B"), False, _
             "blok storniranog siblinga je ostao AKTIVAN"
    AssertEq StorniranoNaID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-A"), True, _
             "izabrana otpremnica je stornirana (mod je odradio svoje)"
End Sub

' ============================================================
' 48. I CILJNA zbirna mora da prodje kapiju, ne samo izvorna
' ============================================================
' Zastita je bila nesimetricna: stara zbirna je od v6-ui-137 imala kapiju, ciljna
' nijednu -- a nizvodne operacije nad ciljem idu PO GOLOM BROJU.
'
' Zatecena kapija u writeru (RequireJedanVlasnikPoBroju) ovo ne pokriva jer broji
' samo AKTIVNE vlasnike. Ovde je owner A STORNIRAN a njegovo dete OTP-HIST je
' AKTIVNO, pa writer vidi jednog vlasnika i pusti relink.
'
' Najgori deo nije kontaminacija nego to sto je SAMA VALIDACIJA potvrdi: i
' SumOtpremniceByKlasa i ValidateZbirnaInvariant sabiraju po broju, pa zaglavlje B
' sa zbirom dece OBA vlasnika prolazi kao konzistentno.
Private Sub T_CiljnaZbirnaDvosmislena_Staje()
    Dim cid As String, res As Object
    Dim pre As Double

    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-TGT-A"), True, _
             "preduslov: jedan vlasnik ciljnog broja je STORNIRAN"
    AssertEq StorniranoNaID(TBL_ZBIRNA, COL_ZBR_ID, "ZBI-TGT-B"), False, _
             "preduslov: drugi vlasnik ciljnog broja je AKTIVAN"
    AssertEq StorniranoNaID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-HIST"), False, _
             "preduslov: dete storniranog vlasnika je AKTIVNO"
    ' Bas ovo zatecena kapija ne vidi -- pa je i propustala.
    AssertEq VlasniciPoBroju(TBL_ZBIRNA, COL_ZBR_BROJ, FX_ZBIRNA_TGT, "T_Cilj", _
                             False, Array(COL_ZBR_VOZAC, COL_ZBR_KUPAC)).count, 1, _
             "preduslov: po AKTIVNIM vlasnicima ciljni broj izgleda jednoznacan"
    AssertEq ZbirnaNaPrijemnici("PRJ-OLD-U"), FX_ZBIRNA_OLDU, _
             "preduslov: izvorna prijemnica visi na izvornoj zbirni"
    ' Zasto je kontaminacija tako podmukla: invarijanta je i sama po BROJU. U
    ' zdravom stanju kaze NEISPRAVNO (sabira decu oba vlasnika, 400, protiv
    ' jednog aktivnog zaglavlja, 100) -- a posle kontaminirane rekalkulacije bi
    ' oba iznosa bila 400 i rekla bi ISPRAVNO. Validacija bi, dakle, potvrdila
    ' pokvareno vlasnistvo kao konzistentno.
    AssertEq CBool(modDokumentInvariant.ValidateZbirnaInvariant(FX_ZBIRNA_TGT)("isValid")), _
             False, "preduslov: invarijanta je number-based, ne ownership-aware"

    pre = KolicinaZbirne("ZBI-TGT-B")
    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_OTPREMNICA, _
                                                  "OTP-OLD-U", FX_OTPREMNICA_OLD_U, _
                                                  "", "", "", "", "", FX_ZBIRNA_OLDU)
    AssertEq (Len(cid) > 0), True, "context je napravljen"

    Set res = modStornoFlow.CompleteOtpremnicaIspravka(cid, FX_OTPREMNICA_NEW_T)

    ' Poslovni ishod PRVI (v. zamka 6 u sabotaza.py): bez kapije zaglavlje B
    ' dobije zbir dece OBA vlasnika, pa mu se kolicina promeni.
    AssertEq KolicinaZbirne("ZBI-TGT-B"), pre, _
             "aktivno ciljno zaglavlje NIJE rekalkulisano preko tudje dece"
    AssertEq ZbirnaNaPrijemnici("PRJ-OLD-U"), FX_ZBIRNA_OLDU, _
             "izvorna prijemnica NIJE prevezana na dvosmislen cilj"
    AssertEq CBool(res("success")), False, "completion staje pred dvosmislenim ciljem"
    AssertEq (InStr(1, CStr(res("message")), "ciljne zbirne", vbTextCompare) > 0), True, _
             "razlog imenuje CILJNU zbirnu, ne staru"
End Sub

' ============================================================
' 49. Ispravka ZBIRNE: ista kapija, obe strane
' ============================================================
' CompleteZbirnaIspravka je imala istu rupu kao ispravka otpremnice, samo sirju:
' po broju idu i izvor i cilj -- RelinkOtpremniceToZbirna_TX(oldBroj, newBroj),
' DistinctActiveValues po oldBroj, ReassignPrijemnicaToZbirna_TX na newBroj,
' RecalculateZbirnaFromOtpremnice_TX(newBroj). Nijedna strana nije bila proverena.
'
' Dvosmislen CILJ znaci "cije zaglavlje dobija zbir", dvosmislen IZVOR znaci
' "cija deca se sele". Zato test meri obe grane, jedna po jedna.
Private Sub T_IspravkaZbirne_KapijaNaObeStrane()
    Dim cid As String, res As Object
    Dim preB As Double

    preB = KolicinaZbirne("ZBI-TGT-B")

    ' (a) DVOSMISLEN CILJ: izvor je jednoznacan, cilj je nekad imao dva vlasnika.
    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_ZBIRNA, _
                                                  "ZBI-OLDU-1", FX_ZBIRNA_OLDU)
    Set res = modStornoFlow.CompleteZbirnaIspravka(cid, FX_ZBIRNA_TGT)
    AssertEq KolicinaZbirne("ZBI-TGT-B"), preB, _
             "dvosmislen CILJ: aktivno zaglavlje nije dobilo zbir tudje dece"
    AssertEq ZbirnaNaOtpremnici("OTP-OLD-U"), FX_ZBIRNA_OLDU, _
             "dvosmislen CILJ: otpremnica izvora nije prevezana"
    AssertEq CBool(res("success")), False, "dvosmislen CILJ zaustavlja ispravku zbirne"
    AssertEq (InStr(1, CStr(res("message")), "ciljne zbirne", vbTextCompare) > 0), True, _
             "razlog imenuje CILJNU stranu"

    ' (b) DVOSMISLEN IZVOR: cilj je jednoznacan, izvor je nekad imao dva vlasnika.
    ' Bez ove grane bi se selila deca oba vlasnika izvornog broja.
    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_ZBIRNA, _
                                                  "ZBI-KASK-1", FX_ZBIRNA_KASK)
    Set res = modStornoFlow.CompleteZbirnaIspravka(cid, FX_ZBIRNA_STALE)
    AssertEq ZbirnaNaOtpremnici("OTP-KOL-A"), FX_ZBIRNA_KASK, _
             "dvosmislen IZVOR: otpremnica nije odseljena sa dvosmislenog broja"
    AssertEq CBool(res("success")), False, "dvosmislen IZVOR zaustavlja ispravku zbirne"
    AssertEq (InStr(1, CStr(res("message")), "stare zbirne", vbTextCompare) > 0), True, _
             "razlog imenuje STARU stranu"
End Sub

Private Function KolicinaZbirne(ByVal zbrID As String) As Double
    Dim v As Variant
    v = LookupValue(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_KOLICINA)
    If IsNumeric(v) Then KolicinaZbirne = CDbl(v)
End Function

' ============================================================
' 47. Kapija ne sme da bude fail-open na sopstvenu gresku
' ============================================================
' "On Error Resume Next" je davao False -- to jest "broj je jednoznacan, mutiraj"
' -- bas kad se nista ne zna: nedostajuca owner kolona, schema drift, greska
' resolvera. Za kapiju je "ne mogu da dokazem jednoznacnost" isto sto i
' "ne mutiraj".
'
' Drift se pravi stvarno (preimenovanje kolone), ne simulira: poenta je da
' RequireColumnIndex digne gresku unutar kapije, a da kapija to pretvori u
' blokadu. Sema se vraca u istom testu.
Private Sub T_KapijaZbirne_FailClosedNaSvojuGresku()
    Dim lo As ListObject
    Dim podDriftom As Boolean, semaVracena As Boolean

    ' Pozitivna kontrola: nad zdravom semom kapija NE blokira jednoznacan broj.
    ' Bez nje bi test prosao i kad kapija uvek vraca True (blokira sve).
    AssertEq modStornoFlow.ZbirnaDvosmislenaIkad_Test(FX_ZBIRNA_MIRNA), False, _
             "pozitivna kontrola: jednoznacan broj prolazi kroz kapiju"

    Set lo = GetTable(TBL_ZBIRNA)
    On Error GoTo VRATI
    lo.ListColumns(COL_ZBR_VOZAC).Name = COL_ZBR_VOZAC & "_DRIFT"
    podDriftom = modStornoFlow.ZbirnaDvosmislenaIkad_Test(FX_ZBIRNA_MIRNA)
VRATI:
    On Error Resume Next
    lo.ListColumns(COL_ZBR_VOZAC & "_DRIFT").Name = COL_ZBR_VOZAC
    On Error GoTo 0
    semaVracena = (GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC) > 0)

    AssertEq podDriftom, True, _
             "nerazresena jednoznacnost se tretira kao dvosmislena"
    ' Ako sema nije vracena, svi testovi posle ovog mere pokvarenu tabelu.
    AssertEq semaVracena, True, "sema je vracena posle testa"
End Sub

' Roditelj koji vraca lookup po poslovnom broju -- to jest PRVI red tog broja.
' Postoji samo da preduslov testa 46 bude proveren, a ne pretpostavljen.
Private Function PrviRoditeljPoBroju(ByVal brojOtp As String) As String
    PrviRoditeljPoBroju = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, _
                                                     brojOtp, COL_OTP_BROJ_ZBIRNE)))
End Function

Private Function ZbirnaNaOtpremnici(ByVal otpID As String) As String
    ZbirnaNaOtpremnici = Trim$(NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, _
                                                    otpID, COL_OTP_BROJ_ZBIRNE)))
End Function

Private Function OtpremnicaNaBloku(ByVal otkupID As String) As String
    OtpremnicaNaBloku = Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, _
                                                   otkupID, COL_OTK_OTPREMNICA_ID)))
End Function

Private Function StorniranoNaID(ByVal tbl As String, ByVal idCol As String, _
                                ByVal id As String) As Boolean
    StorniranoNaID = (UCase$(Trim$(NzToText(LookupValue(tbl, idCol, id, _
                                                        COL_STORNIRANO)))) = "DA")
End Function

' OldDocID iz correction context-a po njegovom PK-u.
Private Function OldDocIDKonteksta(ByVal correctionID As String) As String
    OldDocIDKonteksta = Trim$(NzToText(LookupValue(TBL_STORNO_VEZE, COL_SV_ID, _
                                                   correctionID, COL_SV_OLD_DOCID)))
End Function

' Gajbice vezane za JEDAN dokument (po PrijemnicaID), ne za broj. Broj je
' labela i dele ga dva kupca, pa zbir po broju ne moze da razlikuje ciljeve --
' bas ono sto ovaj test treba da dokaze.
Private Function GajbicaZaDokument(ByVal prijemnicaID As String) As Long
    Dim d As Variant, i As Long, cPid As Long, cSt As Long, cGa As Long
    d = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(d) Then Exit Function
    cPid = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    cGa = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If cPid = 0 Or cGa = 0 Then Exit Function
    For i = 1 To UBound(d, 1)
        If Trim$(NzToText(d(i, cPid))) = prijemnicaID Then
            If cSt = 0 Then
                GajbicaZaDokument = GajbicaZaDokument + NzL(d(i, cGa))
            ElseIf UCase$(Trim$(NzToText(d(i, cSt)))) <> "DA" Then
                GajbicaZaDokument = GajbicaZaDokument + NzL(d(i, cGa))
            End If
        End If
    Next i
End Function

' ============================================================
' 28. Prevezivanje prijemnice na zbirnu ne sme da povuce TUDJU paletu
' ============================================================
' Prvi deo ReassignPrijemnicaToZbirna_TX je birao redove tblPrijemnica po
' generaciji -- tacno. Ali je zatim NOVU BrojZbirne propagirao u tblPaletaStavka
' po BrojPrijemnice, cime je ponistavao ceo taj izbor.
'
' Posledica nije bila "prevezano malo vise" nego dokument koji SAM SEBI
' PROTIVRECI: prijemnica drugog kupca ostaje na staroj zbirni, a njena paleta
' zavrsi na novoj. Sledljivost paleta -> zbirna -> kooperanti tada laze.
'
'   PRJ-TEST-Z1  KUP-TEST-1   paleta PST-TEST-Z1   <- prevezuje se
'   PRJ-TEST-Z2  KUP-TEST-2   paleta PST-TEST-Z2   <- ne sme da se pomeri
'
' Oba nose broj 6/150326, kao u produkciji: broj se racuna po kupcu.
Private Sub T_PrevezivanjeNaZbirnu_PaletaIdePoIdentitetu()
    Dim ok As Boolean

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z1", "GEN-ZBR-1"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2", "GEN-ZBR-2"

    AssertEq ZbirnaNaPrijemnici("PRJ-TEST-Z2"), FX_ZBIRNA_MIRNA, _
             "preduslov: tudji dokument pocinje na staroj zbirni"
    AssertEq ZbirnaNaStavci("PST-TEST-Z2"), FX_ZBIRNA_MIRNA, _
             "preduslov: tudja paleta pocinje na staroj zbirni"

    ok = ReassignPrijemnicaToZbirna_TX(FX_PRIJ_ZBR_KOLIZIJA, FX_ZBIRNA2, "GEN-ZBR-1")
    AssertEq ok, True, "prevezivanje po generaciji je proslo"

    ' Svoj dokument -- oba reda, i prijemnica i njena paleta.
    AssertEq ZbirnaNaPrijemnici("PRJ-TEST-Z1"), FX_ZBIRNA2, _
             "svoja prijemnica je presla na novu zbirnu"
    AssertEq ZbirnaNaStavci("PST-TEST-Z1"), FX_ZBIRNA2, _
             "svoja paleta je presla na novu zbirnu"

    ' TUDJI dokument -- nijedan od ta dva reda. Druga tvrdnja je ono sto je
    ' propustala: prijemnica je ostajala, paleta nije.
    AssertEq ZbirnaNaPrijemnici("PRJ-TEST-Z2"), FX_ZBIRNA_MIRNA, _
             "tudja prijemnica OSTAJE na staroj zbirni"
    AssertEq ZbirnaNaStavci("PST-TEST-Z2"), FX_ZBIRNA_MIRNA, _
             "tudja paleta OSTAJE na staroj zbirni"
End Sub

' ============================================================
' 29. Zadata generacija koje nema NIJE poziv na fallback po broju
' ============================================================
' Prazan argument i zadat-ali-nepostojeci su dva razlicita stanja:
'
'   ""        pozivalac ne zna identitet (zatecen zapis) -> fallback po broju,
'             ali tek kroz kapiju nad jednoznacnoscu
'   "GEN-X"   pozivalac je rekao BAS TAJ dokument. Ako ga nema, pad na broj bi
'             znacio da se dira NESTO DRUGO -- tise i gore od greske.
Private Sub T_ZadataGeneracijaKojeNema_Staje()
    Dim ok As Boolean, upoz As String, gajbDiff As Boolean
    Dim preZbirna As String, preStavki As Long

    preZbirna = ZbirnaNaPrijemnici("PRJ-TEST-Z2")
    preStavki = StavkiZaPrijemnicu(FX_PRIJ_KOLIZIJA)

    ok = ReassignPrijemnicaToZbirna_TX(FX_PRIJ_ZBR_KOLIZIJA, FX_ZBIRNA2, "GEN-NE-POSTOJI")
    AssertEq ok, False, "zadata generacija prijemnice koje nema zaustavlja upis"
    AssertEq ZbirnaNaPrijemnici("PRJ-TEST-Z2"), preZbirna, _
             "posle odbijanja nijedan dokument nije pomeren"

    ok = ReassignPrijemnicaToZbirna_TX(FX_PRIJ_ZBR_KOLIZIJA, FX_ZBIRNA2, _
                                       "GEN-ZBR-1", "GEN-ZBIRNE-NEMA")
    AssertEq ok, False, "zadata generacija CILJNE zbirne koje nema zaustavlja upis"

    ok = ReassignPaleteToPrijemnica_TX(FX_PRIJ_KOLIZIJA, FX_PRIJ_BROJ, upoz, True, _
                                       gajbDiff, "GEN-NE-POSTOJI", "GEN-CILJ-A")
    AssertEq ok, False, "zadata generacija izvora paleta koje nema zaustavlja upis"
    AssertEq (Len(upoz) > 0), True, "odbijanje nosi razlog za operatera"
    AssertEq StavkiZaPrijemnicu(FX_PRIJ_KOLIZIJA), preStavki, _
             "posle odbijanja nijedna paletna stavka nije pomerena"
End Sub

' ============================================================
' 30. Presuda o RELABEL-u mora da opisuje BAS izabran dokument
' ============================================================
' Writer je dokument birao po GeneracijaID, a onda zvao
' EvaluatePaletaReassign(oldBroj, newBroj), koja ga je PONOVO trazila po
' poslovnom broju i uzimala PRVI red. Kod kolizije je presuda opisivala drugi
' dokument.
'
' Kvar je tisi od pogresnog prevezivanja: upis ide na tacnu prijemnicu, samo se
' RELABEL preskoci, pa paleta ostane oznacena starom robom.
'
'   PRJ-TEST-C1  TESTVOCE    <- prvi red broja 8/150326; ISTA vrsta kao cilj
'   PRJ-TEST-C2  TESTVOCE2   <- stvarni izvor (GEN-TEST-B)
'   cilj 1/150326            TESTVOCE
'
' Presuda po broju vidi C1 i kaze CLEAN. Presuda po identitetu vidi C2 i mora
' reci RELABEL.
Private Sub T_VerdiktPoIdentitetu_RelabelSeNePreskace()
    Dim v As Variant, ok As Boolean, upoz As String, gajbDiff As Boolean

    v = EvaluatePaletaReassign(FX_PRIJ_KOLIZIJA, FX_PRIJ_BROJ, "GEN-TEST-B", "GEN-CILJ-A")
    AssertEq CStr(v(0)), "RELABEL", _
             "presuda po identitetu vidi razliku vrste izabranog dokumenta"

    ok = ReassignPaleteToPrijemnica_TX(FX_PRIJ_KOLIZIJA, FX_PRIJ_BROJ, upoz, True, _
                                       gajbDiff, "GEN-TEST-B", "GEN-CILJ-A")
    AssertEq ok, True, "prevezivanje uz relabel je proslo"

    ' Ono zbog cega presuda uopste postoji: etiketa na stavci mora da prati robu
    ' na koju je stavka prevezana. Bez relabela ostaje stara vrsta.
    AssertEq VrstaNaStavci("PST-TEST-C2"), FX_VRSTA, _
             "stavka je prelabelirana na vrstu ciljnog dokumenta"
End Sub

' ============================================================
' 31. Su-stanar na deljenoj paleti je DRUGI DOKUMENT, ne drugi broj
' ============================================================
' Pred relabel se proverava da li fizicka paleta nosi i tudju robu: ako nosi,
' promena headera bi iskvarila i nju, pa se operacija blokira. Ideja je tacna,
' ali se "tudja" merilo poredjenjem BROJEVA:
'
'   If bpg <> oldBroj And bpg <> newBroj Then ...
'
' Dva kupca istog broja i ISTE robe smeju legitimno da dele paletu -- roba im
' je identicna, nema sta da se razlikuje. Za tu kapiju su izgledali kao ista
' prijemnica (bpg = oldBroj), pa nije okidala: STEP 2b bi prepravio header CELE
' palete na novu robu, a su-stanar ostaje stara. Paleta i njena stavka bi od
' tog trenutka tvrdile razlicito.
'
'   PRJ-TEST-D1  KUP-TEST-1  TESTVOCE  \  ista paleta PAL-TEST-D
'   PRJ-TEST-D2  KUP-TEST-2  TESTVOCE  /  isti broj 5/150326
'   cilj PRJ-TEST-T2         TESTVOCE2 -> relabel
Private Sub T_DeljenaPaleta_SuStanarPoIdentitetu()
    Dim ok As Boolean, upoz As String, gajbDiff As Boolean

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-D1", "GEN-DEL-1"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-D2", "GEN-DEL-2"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-T2", "GEN-CILJ-V2"

    AssertEq VrstaNaPaleti("PAL-TEST-D"), FX_VRSTA, _
             "preduslov: deljena paleta pocinje sa starom robom"

    ' allowRelabel = True: potvrda relabela POSTOJI, a operacija se svejedno
    ' odbija -- deljena paleta se ne moze prelabelirati ni uz potvrdu.
    ok = ReassignPaleteToPrijemnica_TX(FX_PRIJ_DELJENA, FX_PRIJ_CILJ_V2, upoz, True, _
                                       gajbDiff, "GEN-DEL-1", "GEN-CILJ-V2")
    AssertEq ok, False, "relabel deljene palete se odbija i uz potvrdu"
    AssertEq (InStr(upoz, "BLOKIRANO") > 0), True, "odbijanje kaze zasto"

    ' Nista se nije pomerilo ni prelabeliralo.
    AssertEq VrstaNaPaleti("PAL-TEST-D"), FX_VRSTA, _
             "header deljene palete OSTAJE stara roba"
    AssertEq VrstaNaStavci("PST-TEST-D2"), FX_VRSTA, _
             "su-stanar OSTAJE stara roba"
    AssertEq PrijemnicaNaStavci("PST-TEST-D1"), "PRJ-TEST-D1", _
             "izvorna stavka nije prevezana"
End Sub

' ============================================================
' 32. Isti broj sa RAZLICITIM generacijama nije isti dokument
' ============================================================
' Ova kapija je popravljena u ekranu, a u writeru je ostala stara -- pravilo je
' time bilo samo preseljeno iz UI-ja u core. Writer se testira direktno, bez
' forme: ekranska putanja otvara MsgBox potvrde i headless se ne vozi.
Private Sub T_IstiBrojRazliciteGeneracije_NijeIstiDokument()
    Dim ok As Boolean, upoz As String, gajbDiff As Boolean

    ok = ReassignPaleteToPrijemnica_TX(FX_PRIJ_BROJ, FX_PRIJ_BROJ, upoz, True, _
                                       gajbDiff, "GEN-CILJ-A", "GEN-CILJ-A")
    AssertEq ok, False, "ista generacija sa obe strane je ISTI dokument -- odbija se"

    ' Isti broj, druge generacije: dva dokumenta, operacija je legitimna.
    '
    ' Meri se DELTA, ne apsolutan zbir: na PRJ-TEST-A stoji ono sto su tamo
    ' ostavili raniji testovi, a testovi dele svesku. Tvrdnja je "sve sto je bilo
    ' na A preslo je na B", i ona ne zavisi od toga koliko je toga bilo.
    Dim preA As Long, preB As Long
    preA = GajbicaZaDokument("PRJ-TEST-A")
    preB = GajbicaZaDokument("PRJ-TEST-B")
    AssertEq (preA > 0), True, "preduslov: izvorni dokument uopste nosi robu"

    ok = ReassignPaleteToPrijemnica_TX(FX_PRIJ_BROJ, FX_PRIJ_BROJ, upoz, True, _
                                       gajbDiff, "GEN-CILJ-A", "GEN-CILJ-B")
    AssertEq ok, True, "isti broj a razlicite generacije PROLAZI"
    AssertEq GajbicaZaDokument("PRJ-TEST-B"), preB + preA, _
             "sva roba je presla na drugi dokument ISTOG broja"
    AssertEq GajbicaZaDokument("PRJ-TEST-A"), 0, _
             "na izvornom dokumentu vise nema robe"
End Sub

Private Function VrstaNaPaleti(ByVal paletaID As String) As String
    VrstaNaPaleti = Trim$(NzToText(LookupValue(TBL_PALETA, COL_PAL_ID, _
                                               paletaID, COL_PAL_VRSTA)))
End Function

Private Function PrijemnicaNaStavci(ByVal stavkaID As String) As String
    PrijemnicaNaStavci = Trim$(NzToText(LookupValue(TBL_PALETA_STAVKA, COL_PALS_ID, _
                                                    stavkaID, COL_PALS_PRIJEMNICA_ID)))
End Function

Private Function VrstaNaStavci(ByVal stavkaID As String) As String
    VrstaNaStavci = Trim$(NzToText(LookupValue(TBL_PALETA_STAVKA, COL_PALS_ID, _
                                               stavkaID, COL_PALS_VRSTA)))
End Function

' BrojZbirne sa jednog reda -- po PK-u, ne po broju dokumenta.
Private Function ZbirnaNaPrijemnici(ByVal prijemnicaID As String) As String
    ZbirnaNaPrijemnici = Trim$(NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, _
                                                    prijemnicaID, COL_PRJ_BROJ_ZBIRNE)))
End Function

Private Function ZbirnaNaStavci(ByVal stavkaID As String) As String
    ZbirnaNaStavci = Trim$(NzToText(LookupValue(TBL_PALETA_STAVKA, COL_PALS_ID, _
                                                stavkaID, COL_PALS_BROJ_ZBIRNE)))
End Function

' Upisi generaciju na red po PK-u. Fixture redovi se seju mimo writera, pa
' generacije nemaju; test ih postavlja da bi dva dokumenta bila razluciva.
Private Sub StampGeneraciju(ByVal tbl As String, ByVal idCol As String, _
                            ByVal docID As String, ByVal gen As String)
    Dim rows As Collection
    Set rows = FindRows(tbl, idCol, docID)
    If rows Is Nothing Then Exit Sub
    If rows.count = 0 Then Exit Sub
    UpdateCell tbl, rows(1), COL_GENERACIJA_ID, gen
End Sub

' Vrednost jednog polja iz prefill opisa ("kljuc=vrednost|kljuc=vrednost").
' Prazno = kljuca nema, sto je za ovaj opis isto sto i "nema vrednosti"
' (Spoji prazne vrednosti uopste ne upisuje).
Private Function SpecVal(ByVal spec As String, ByVal kljuc As String) As String
    Dim par As Variant, kv As Variant
    If Len(spec) = 0 Then Exit Function
    For Each par In Split(spec, "|")
        kv = Split(CStr(par), "=")
        If UBound(kv) >= 1 Then
            If CStr(kv(0)) = kljuc Then
                SpecVal = CStr(kv(1))
                Exit Function
            End If
        End If
    Next par
End Function

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
