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
Private Const FX_KOOPERANT3 As String = "KOOP-TEST-3"   ' bez agro odbitka
Private Const FX_ABZUG_KOOP1 As Double = 500   ' 300 + 200; 999 je storniran
Private Const FX_ABZUG_KOOP2 As Double = 100
' Kolizija broja po godini: isti broj, dve godine, dva identiteta.
Private Const FX_PAL_KOL_BROJ As String = "12"
Private Const FX_PAL_KOL_STARA As String = "PAL-TEST-Y25"   ' 12/2025
Private Const FX_PAL_KOL_NOVA As String = "PAL-TEST-Z2"     ' 12/2026
Private Const FX_PRE_KOL_BROJ As String = "7"
Private Const FX_PRE_KOL_STARA As String = "PRE-TEST-Y25"   ' neto 200
Private Const FX_PRE_KOL_NOVA As String = "PRE-TEST-Y26"    ' neto 300
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
' Fakture ekrana Fakturisanje. FAK-TEST-1 i FAK-TEST-0 nemaju broj, datum ni
' status (i jedan raniji test FAK-TEST-1 u toku suite-a PLATI), pa se tvrdnje
' o listi i cipovima oslanjaju iskljucivo na ove tri.
Private Const FX_FAK_NEPL As String = "FAK-TEST-N"     ' KUPAC2, bez uplate
Private Const FX_FAK_NEPL_IZNOS As Double = 5000
Private Const FX_FAK_PLAC As String = "FAK-TEST-P"     ' KUPAC, uplacena u celosti
Private Const FX_FAK_PLAC_IZNOS As Double = 4000
Private Const FX_FAK_STORNO As String = "FAK-TEST-X"   ' ne sme da se pojavi u listi
' Prijemnice ekrana Fakturisanje, sve tri kupca FX_KUPAC:
'   FAK1 uredno fakturisana | FAK2 obelezena ali BEZ FakturaID | FAK3 slobodna
Private Const FX_PRJ_FAK1 As String = "PRJ-FAK-1"
Private Const FX_PRJ_FAK2 As String = "PRJ-FAK-2"
Private Const FX_PRJ_FAK3 As String = "PRJ-FAK-3"

' BANKA UVOZ (Faza E/17). tblBankaImport je do ovog fixture-a bio PRAZAN, pa je
' svaka tvrdnja o listi stavki, cipovima i integritetu izvoda radila nad praznim
' skupom.
Private Const FX_BIM_JAKI_FAK As String = "BIM-FIX-1"    ' poziv = broj fakture 2/2026
Private Const FX_BIM_JAKI_BLOK As String = "BIM-FIX-2"   ' poziv = BrojDokumenta 1/TEST
Private Const FX_BIM_BEZ_KLJUCA As String = "BIM-FIX-3"  ' bez poziva i bez konta
Private Const FX_BIM_KOLIZIJA As String = "BIM-FIX-K"    ' DRUGI racun, ISTI broj izvoda
Private Const FX_BIM_BLOK3 As String = "BIM-FIX-3K"      ' blok sa 3 otvorene stavke
Private Const FX_BIM_ERROR As String = "BIM-FIX-ER"      ' auto pokusao pa vratio
Private Const FX_BIM_DA As String = "BIM-FIX-DA"
Private Const FX_BIM_SKIP As String = "BIM-FIX-SK"
Private Const FX_BIM_STORNO As String = "BIM-FIX-ST"
Private Const FX_BIM_IZVOD1 As String = "IZV-FIX-1"
Private Const FX_BIM_IZVOD2 As String = "IZV-FIX-2"
Private Const FX_BIM_IZVOD3 As String = "IZV-FIX-3"
' DVA reda pod ISTIM BankaImportID-em. Bez njih bi tvrdnja "dvosmislen ID
' nosi PRAZAN identitet" merila odsustvo reda.
Private Const FX_BIM_DUP As String = "BIM-FIX-DUP"
Private Const FX_BIM_RACUN1 As String = "160-0000000111111-11"
Private Const FX_BIM_RACUN2 As String = "265-0000000222222-22"
' Blok kooperanta KOOP-TEST-3 sa TRI otvorene stavke -- preko granice koju
' automatska raspodela sme da podeli.
Private Const FX_BIM_BLOK3_BR As String = "BLK-BIM-3"
' Blok kooperanta KOOP-TEST-1 sa jednom otvorenom stavkom (OTK-TEST-1).
Private Const FX_BIM_BLOK1_BR As String = "1/TEST"
' ISTI kooperant, ISTI broj bloka, DVE stanice. Broj otkupa je jedinstven PO
' STANICI, pa je ovo legitiman podatak -- i jedini nacin da se izmeri da rucno
' mapiranje nosi scope otkupnog mesta, a ne samo broj.
Private Const FX_BIM_BLOK_OM As String = "BLK-BIM-OM"
Private Const FX_OTK_OM_A As String = "OTK-BIM-OMA"   ' STANICA
Private Const FX_OTK_OM_B As String = "OTK-BIM-OMB"   ' druga stanica
' ISTI blok, BEZ upisanog otkupnog mesta -- legacy oblik koji danasnji pisci
' odbijaju, a zatecene sveske ga imaju.
Private Const FX_OTK_OM_X As String = "OTK-BIM-OMX"
' FX_STANICA2 je namerno NEPOSTOJECA ("tudje OM"); scope trazi pravu drugu.
Private Const FX_STANICA_B As String = "STA-TEST-2"
' Blok koji je u CELOSTI placen. Lista blokova ga i dalje nudi (ne proverava dug),
' a kandidata za placanje nema -- writer bi takav izbor tiho preveo u avans.
Private Const FX_BIM_BLOK_PLACEN As String = "BLK-BIM-PLAC"
' Izvod cija DVA REDA nose RAZLICITE zbirove.
Private Const FX_BIM_IZVOD_NES As String = "IZV-FIX-NES"
' Isti broj izvoda i isti racun, DRUGI ciklus.
Private Const FX_BIM_PY As String = "BIM-FIX-PY"
' Red se u mrezi nalazi PO PARTNERU, jer BankaImportID vise nije prikazan --
' interna sifra ne ide operateru pred oci (nalaz iz smoke-a).
Private Const FX_BIM_P_JAKI_FAK As String = "Kupac Prvi doo"
Private Const FX_BIM_P_JAKI_BLOK As String = "Prvi Testni"
Private Const FX_BIM_P_KOLIZIJA As String = "Drugi Platilac"
Private Const FX_BIM_P_DA As String = "Obradjeni Platilac"
Private Const FX_BIM_P_DUP As String = "Dvojnik Prvi"
Private Const FX_BIM_P_STORNO As String = "Stornirani Platilac"
Private Const FX_BIM_SVE As Long = 13       ' 14 redova minus jedan storniran
Private Const FX_BIM_OTVORENIH As Long = 7  ' status "" ili "Error"
Private Const FX_BIM_OBRADJENIH As Long = 5 ' DA + dva dvojnika + prosli ciklus + jedan nesaglasan
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
' AGROHEMIJA (tools/make_fixture.py, SEED tblArtikli / tblMagacin).
' ART-TEST-1: Pakovanje 5, DozaPoHa 2, cena 500. ULAZ 20 - IZLAZ 5 = STANJE 15.
' Pakovanje 5 uz dozu 2 je izabrano tako da se ZAOKRUZENJE NAGORE vidi:
' 1.5 ha * 2 = 3 l, a izdaje se jedno pakovanje od 5 l.
Private Const FX_ARTIKAL As String = "ART-TEST-1"
Private Const FX_ARTIKAL_BEZ_PAK As String = "ART-TEST-2"
Private Const FX_ARTIKAL_BEZ_STANJA As String = "ART-TEST-3"
' Artikal sa velikom zalihom i pakovanjem od 1. Traka korpe se drugacije ne moze
' izmeriti: ART-TEST-1 kroz kapiju stanja pusta najvise TRI pakovanja, a za
' preliv trake mora da udje vise stavki nego sto ona ima redova.
Private Const FX_ARTIKAL_ZALIHA As String = "ART-TEST-Z"
Private Const FX_ART_PAKOVANJE As Double = 5
Private Const FX_ART_STANJE As Double = 15
' Kooperant ISTOG IMENA kao FX_KOOPERANT ("Prvi Testni"), drugi identitet.
' Postoji da bi "dvosmislen prikaz se odbija" imalo nad cim da padne.
Private Const FX_KOOP_ISTOIME As String = "KOOP-TEST-IME"
Private Const FX_KOOP_PRIKAZ As String = "Prvi Testni"
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
    RunOne 60
    RunOne 61
    RunOne 62
    RunOne 63
    RunOne 64
    RunOne 65
    RunOne 66
    RunOne 67
    RunOne 68
    RunOne 69
    RunOne 70
    RunOne 71
    RunOne 72
    RunOne 73
    RunOne 74
    RunOne 75
    RunOne 76
    RunOne 77
    RunOne 78
    RunOne 79
    RunOne 80
    RunOne 81
    RunOne 82
    RunOne 83
    RunOne 84
    RunOne 85
    RunOne 86
    RunOne 87
    RunOne 88
    RunOne 89
    RunOne 90
    RunOne 91
    RunOne 92
    RunOne 93
    RunOne 94
    RunOne 95
    RunOne 96
    RunOne 97
    RunOne 98
    RunOne 99
    RunOne 100
    RunOne 101
    RunOne 102
    RunOne 103
    RunOne 104
    RunOne 105
    RunOne 106
    RunOne 107
    RunOne 108
    RunOne 109
    RunOne 110
    RunOne 111
    RunOne 112
    RunOne 113
    RunOne 114
    RunOne 115
    RunOne 116
    RunOne 117
    RunOne 118
    RunOne 119
    RunOne 120
    RunOne 121
    RunOne 122
    RunOne 123

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
        Case 60: TestName = "T_StornoImpact_PoIdentitetu"
        Case 61: TestName = "T_StornoAkcije_RefreshInvalidiraOdluku"
        Case 62: TestName = "T_StornoBezUvida_NemaAkcije"
        Case 63: TestName = "T_StornoImpact_SchemaDriftJeInvalidan"
        Case 64: TestName = "T_StornoImpact_IdentitetNeDegradira"
        Case 65: TestName = "T_StornoImpact_BlokSekcijaDriftJeInvalidna"
        Case 66: TestName = "T_StornoEkran_NeCuriGreska"
        Case 67: TestName = "T_StornoImpact_PrijemnicaBlokDriftJeInvalidan"
        Case 68: TestName = "T_LogErr_NeVidiErrPosleResumeNext"
        Case 69: TestName = "T_PorukeUnosa_UpozorenjeNosiOznaku"
        Case 70: TestName = "T_StornoImpact_NestaoIdentitetJeInvalidan"
        Case 71: TestName = "T_Oporavak_OdbaciIspravku_PoIdentitetu"
        Case 72: TestName = "T_Oporavak_OdbaciIspravku_GasiSamoSvoj"
        Case 73: TestName = "T_ImpactPalete_ZaglavljeIzPraveVrste"
        Case 74: TestName = "T_StornoEfekat_TekstIzKataloga"
        Case 75: TestName = "T_StornoBlokovi_PodrazumevanoNijedan"
        Case 76: TestName = "T_NavBrojac_SamoEkranKojiBroji"
        Case 77: TestName = "T_NovaPrerada_IzborINeto"
        Case 78: TestName = "T_PaletaDvoklik_OtvaraStavke"
        Case 79: TestName = "T_CipoviEkrana_UgovorIFilter"
        Case 80: TestName = "T_ZonaPrerade_SvaPoljaVidljiva"
        Case 81: TestName = "T_BazenLjuske_ViseNegoStoStaje"
        Case 82: TestName = "T_Agro_UgovorEkrana"
        Case 83: TestName = "T_Agro_KapijaStanjaBrojiKorpu"
        Case 84: TestName = "T_Agro_SmartDozaZaokruzujeNagore"
        Case 85: TestName = "T_ZonaAgro_PoljaPostojeIPrateRezim"
        Case 86: TestName = "T_Agro_CipoviSuzavajuListu"
        Case 87: TestName = "T_Agro_BrojacIDvoklikPoIdentitetu"
        Case 88: TestName = "T_Agro_AbzugMapaPratiPojedinacni"
        Case 89: TestName = "T_ZonaAgro_PrekidacRezimaZadrzavaBoju"
        Case 90: TestName = "T_Agro_TrakaKorpe_NajnovijePrvoIPreliv"
        Case 91: TestName = "T_Agro_KorpaUklanjaPoIdentitetu"
        Case 92: TestName = "T_Agro_ZnackaPratiKorpuVanKorpeListe"
        Case 93: TestName = "T_PaleteIdentitet_PoIDNePoBroju"
        Case 94: TestName = "T_PreradeIdentitet_PoIDNePoBroju"
        Case 95: TestName = "T_GridTelo_NePokrivaToast"
        Case 96: TestName = "T_PaleteScrEvent_NeCuriGreska"
        Case 97: TestName = "T_Fak_UgovorEkrana"
        Case 98: TestName = "T_Fak_IdentitetURedu_NeCrtaSe"
        Case 99: TestName = "T_Fak_DostupnostSePrenosiURedu"
        Case 100: TestName = "T_Fak_KorpaZnackaITraka"
        Case 101: TestName = "T_Fak_CipoviPrateStatusFakture"
        Case 102: TestName = "T_Fak_NerazresenKupacNeDiraKorpu"
        Case 103: TestName = "T_Fak_GreskaNePreziviLogErr"
        Case 104: TestName = "T_BankaUvoz_UgovorEkrana"
        Case 105: TestName = "T_BankaUvoz_IdentitetURedu_NeCrtaSe"
        Case 106: TestName = "T_BankaUvoz_RedNosiSmerIOtvorenost"
        Case 107: TestName = "T_BankaUvoz_CipJakihPratiBrojac"
        Case 108: TestName = "T_BankaUvoz_IzvodiSuAgregatPoRacunu"
        Case 109: TestName = "T_BankaUvoz_RucnoMapiranjePravila"
        Case 110: TestName = "T_ZonaBankaUvoz_PoljaIRaspored"
        Case 111: TestName = "T_MrezaDatum_BrojKojiNijeDatum"
        Case 112: TestName = "T_MrezaGeometrija_PratiOpisKolona"
        Case 113: TestName = "T_MrezaCelija_NeostavljaTudjiTekst"
        Case 114: TestName = "T_LegacyBanka_PadUcitavanjaNijePraznaLista"
        Case 115: TestName = "T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana"
        Case 116: TestName = "T_Mreza_PodnozjeDvaNovcanaSlota"
        Case 117: TestName = "T_Kolona_TrazenjeNeGutaGresku"
        Case 118: TestName = "T_MrezaPilula_PozadinaSeCisti"
        Case 119: TestName = "T_LegacyDok_PadListeBlokovaNijeAvans"
        Case 120: TestName = "T_LegacyDok_PadListeFakturaNijeAvans"
        Case 121: TestName = "T_Ljuska_PadListeNovcaNijeAvans"
        Case 122: TestName = "T_StornoFilter_NedostajucaKolonaNijeTisina"
        Case 123: TestName = "T_KesKolone_NeMemoiseNulu"
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
        Case 60: T_StornoImpact_PoIdentitetu
        Case 61: T_StornoAkcije_RefreshInvalidiraOdluku
        Case 62: T_StornoBezUvida_NemaAkcije
        Case 63: T_StornoImpact_SchemaDriftJeInvalidan
        Case 64: T_StornoImpact_IdentitetNeDegradira
        Case 65: T_StornoImpact_BlokSekcijaDriftJeInvalidna
        Case 66: T_StornoEkran_NeCuriGreska
        Case 67: T_StornoImpact_PrijemnicaBlokDriftJeInvalidan
        Case 68: T_LogErr_NeVidiErrPosleResumeNext
        Case 69: T_PorukeUnosa_UpozorenjeNosiOznaku
        Case 70: T_StornoImpact_NestaoIdentitetJeInvalidan
        Case 71: T_Oporavak_OdbaciIspravku_PoIdentitetu
        Case 72: T_Oporavak_OdbaciIspravku_GasiSamoSvoj
        Case 73: T_ImpactPalete_ZaglavljeIzPraveVrste
        Case 74: T_StornoEfekat_TekstIzKataloga
        Case 75: T_StornoBlokovi_PodrazumevanoNijedan
        Case 76: T_NavBrojac_SamoEkranKojiBroji
        Case 77: T_NovaPrerada_IzborINeto
        Case 78: T_PaletaDvoklik_OtvaraStavke
        Case 79: T_CipoviEkrana_UgovorIFilter
        Case 80: T_ZonaPrerade_SvaPoljaVidljiva
        Case 81: T_BazenLjuske_ViseNegoStoStaje
        Case 82: T_Agro_UgovorEkrana
        Case 83: T_Agro_KapijaStanjaBrojiKorpu
        Case 84: T_Agro_SmartDozaZaokruzujeNagore
        Case 85: T_ZonaAgro_PoljaPostojeIPrateRezim
        Case 86: T_Agro_CipoviSuzavajuListu
        Case 87: T_Agro_BrojacIDvoklikPoIdentitetu
        Case 88: T_Agro_AbzugMapaPratiPojedinacni
        Case 89: T_ZonaAgro_PrekidacRezimaZadrzavaBoju
        Case 90: T_Agro_TrakaKorpe_NajnovijePrvoIPreliv
        Case 91: T_Agro_KorpaUklanjaPoIdentitetu
        Case 92: T_Agro_ZnackaPratiKorpuVanKorpeListe
        Case 93: T_PaleteIdentitet_PoIDNePoBroju
        Case 94: T_PreradeIdentitet_PoIDNePoBroju
        Case 95: T_GridTelo_NePokrivaToast
        Case 96: T_PaleteScrEvent_NeCuriGreska
        Case 97: T_Fak_UgovorEkrana
        Case 98: T_Fak_IdentitetURedu_NeCrtaSe
        Case 99: T_Fak_DostupnostSePrenosiURedu
        Case 100: T_Fak_KorpaZnackaITraka
        Case 101: T_Fak_CipoviPrateStatusFakture
        Case 102: T_Fak_NerazresenKupacNeDiraKorpu
        Case 103: T_Fak_GreskaNePreziviLogErr
        Case 104: T_BankaUvoz_UgovorEkrana
        Case 105: T_BankaUvoz_IdentitetURedu_NeCrtaSe
        Case 106: T_BankaUvoz_RedNosiSmerIOtvorenost
        Case 107: T_BankaUvoz_CipJakihPratiBrojac
        Case 108: T_BankaUvoz_IzvodiSuAgregatPoRacunu
        Case 109: T_BankaUvoz_RucnoMapiranjePravila
        Case 110: T_ZonaBankaUvoz_PoljaIRaspored
        Case 111: T_MrezaDatum_BrojKojiNijeDatum
        Case 112: T_MrezaGeometrija_PratiOpisKolona
        Case 113: T_MrezaCelija_NeostavljaTudjiTekst
        Case 114: T_LegacyBanka_PadUcitavanjaNijePraznaLista
        Case 115: T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana
        Case 116: T_Mreza_PodnozjeDvaNovcanaSlota
        Case 117: T_Kolona_TrazenjeNeGutaGresku
        Case 118: T_MrezaPilula_PozadinaSeCisti
        Case 119: T_LegacyDok_PadListeBlokovaNijeAvans
        Case 120: T_LegacyDok_PadListeFakturaNijeAvans
        Case 121: T_Ljuska_PadListeNovcaNijeAvans
        Case 122: T_StornoFilter_NedostajucaKolonaNijeTisina
        Case 123: T_KesKolone_NeMemoiseNulu
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

    Unload f
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

    Unload f
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

    ' Do v6-ui-143 je ovde prva tvrdnja bila da F8 (storno) vraca
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
' PRETARGETIRAN u v6-ui-143: mera je ista, ali seam vise nije rezim F8 nego
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

    ' Rezim i dalje mora da stigne do iste tabele -- ModeTable je od v6-ui-143
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
    AssertEq (InStr(modScrOporavak.Scr_Radnje(), "odbaci:") = 1), True, _
             "Nedovrseno ima Odbaci ispravku"
    AssertEq (InStr(modScrOporavak.Scr_Radnje(), ":danger:") > 0), True, _
             "Odbaci ispravku nosi danger stil"
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
' Od v6-ui-143 forme nema: storno je ekran u registru, sa "upis=ne". Ovaj test
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
' do v6-ui-143 nije imao.
Private Sub T_Storno_UgovorIRadnje()
    Dim liste As Variant, i As Long, kljucevi As String, d As Variant

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

    ' A CRTANJE NIJE ISTO STO I DISPECOVANJE. Klik na cip je isao kroz uslov
    ' Len(tag) = 6, sto pokriva samo lsSeg0..lsSeg9 -- pa je jedanaesti cip
    ' (lsSeg10) imao sedam znakova, propadao kroz granu i nije radio NISTA:
    ' crta se, boji se na hover, a klik nema kome da stigne. Operater je to
    ' prijavio kao 'cip postoji ali je mrtav'.
    '
    ' Zato se meri i druga kapija, i to za POSLEDNJI cip -- prvi je radio i pre.
    AssertEq modOtkupUI.SegIndeksIzTaga("lsSeg" & UBound(liste)), UBound(liste), _
             "ljuska razresava klik na POSLEDNJI cip, ne samo na jednocifrene"
    AssertEq modOtkupUI.SegIndeksIzTaga("lsSeg0"), 0, "i na prvi, i dalje"
    AssertEq modOtkupUI.SegIndeksIzTaga("btnAct0"), -1, "tudji tag nije cip"
    AssertEq modOtkupUI.SegIndeksIzTaga("lsSeg"), -1, "cip bez rednog broja nije cip"
    AssertEq (UBound(liste) + 1), 11, "ekran ima jedanaest lista"
    For i = 0 To UBound(liste)
        kljucevi = kljucevi & "|" & Split(CStr(liste(i)), "|")(0)
    Next i
    AssertEq kljucevi, "|LANAC|OTKUP|OTPREMNICA|ZBIRNA|PRIJEMNICA|AMB_ISPLATE|" & _
             "AMB_UPLATE|REVERSI|FAKTURA|IZVOD|BLOKOVI", _
             "redosled i kljucevi lista -- navigaciona je prva, blokovi poslednji"

    ' Kljucevi TIPIZIRANIH lista JESU kljucevi tipova (STIP_*), pa prevodne tabele
    ' nema. Ako se razidju, Scr_Rows bi trazio tabelu za nepostojeci tip i tiho
    ' vratio otkupne listove pod tudjim naslovom.
    '
    ' Prva (LANAC) i poslednja (BLOKOVI) su POGLEDI, ne tipovi: prva nad svim
    ' framework dokumentima, poslednja nad blokovima vec izabranog dokumenta.
    For i = 1 To UBound(liste) - 1
        AssertEq (Len(modScrDokumenti.TabelaTipa(Split(CStr(liste(i)), "|")(0))) > 0), True, _
                 "lista " & Split(CStr(liste(i)), "|")(0) & " ima svoju tabelu"
    Next i

    ' Pogled se NE SME provuci kroz RedoviZaTip: TabelaTipa na nepoznat kljuc
    ' vraca tblOtkup (Case Else), pa bi lista blokova tiho prikazala otkupne
    ' listove pod svojim naslovom. Meri se posledica -- koje kolone lista vrati.
    modScrStorno.Scr_TipTestSet "BLOKOVI"
    d = modScrStorno.Scr_Rows("sve", "")
    AssertEq Split(CStr(d(0)(0)), "|")(0), "OTKUI_HD_OZN", _
             "lista blokova pocinje kolonom izbora, ne kolonama otkupnog lista"
    AssertEq Split(CStr(d(0)(UBound(d(0)))), "|")(0), "OTKUI_HD_IDENT", _
             "i zavrsava nevidljivom kolonom identiteta bloka"
    modScrStorno.Scr_TipTestSet STIP_OTKUP
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
' 60. Uvid je identity-scoped u CELOSTI, ne samo u lancu
' ============================================================
' Do v6-ui-143 je BuildStornoImpact provlacio docID kroz zaglavlje, lanac,
' blokove i zastavice -- a PALETE i FAKTURU je i dalje trazio po BROJU:
'
'     Set d("palete")  = ImpactPalete(docType, broj)
'     Set d("faktura") = ImpactFaktura(docType, broj)
'
' Pod kolizijom broja to znaci: zaglavlje, lanac i blokovi pokazuju izabran
' dokument, a palete pokazuju palete OBA. Writer nizvodno mutira samo izabrani.
' Ekran koji obecava "ovo su posledice" tvrdio bi posledice koje se nece desiti
' -- tacno klasa greske koju je #198 devet rundi vadio iz poslovnog sloja.
'
' Fixture ima tacno taj par: PRJ-TEST-Z1 i PRJ-TEST-Z2 dele broj, imaju razlicite
' kupce i SVOJE palete (PAL-TEST-Z1 = 10 gajbi / 100 kg, PAL-TEST-Z2 = 20 / 200).
Private Sub T_StornoImpact_PoIdentitetu()
    Dim m As Object, pal As Collection, sm As Object

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z1", "GEN-IMP-1"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2", "GEN-IMP-2"

    ' PREDUSLOV: po golom broju uvid stvarno vidi palete OBA dokumenta -- inace
    ' test ne meri suzavanje nego prazan skup.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA)
    Set pal = m("palete")
    AssertEq pal.count, 2, "preduslov: po golom broju uvid nosi palete OBA dokumenta"

    ' NAJVAZNIJE: sa identitetom uvid nosi SAMO palete izabranog dokumenta.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "GEN-IMP-2")
    Set pal = m("palete")
    AssertEq pal.count, 1, "sa identitetom uvid nosi SAMO palete izabranog dokumenta"

    ' I to bas njegove: Z2 nosi 20 gajbi / 200 kg, Z1 nosi 10 / 100.
    Set sm = m("summary")
    AssertEq CLng(sm("detachGajb")), 20, "zbir uticaja je zbir NJEGOVIH paleta"
    AssertEq CDbl(sm("detachNeto")), 200#, "i njegovih kilograma"
    AssertEq CLng(sm("paleteCount")), 1, "sazetak broji iste palete koje su prikazane"

    ' Zaglavlje mora da opisuje isti dokument -- Z2 je kupca KUP-TEST-2.
    AssertEq CStr(m("header")("partnerID")), "KUP-TEST-2", _
             "zaglavlje opisuje IZABRAN dokument, ne prvi po broju"

    ' Model koji je prosao mora i da se prijavi kao valjan.
    AssertEq CBool(m("valid")), True, "kompletan uvid se prijavljuje kao valjan"
End Sub

' ============================================================
' 61. Promena podataka ponistava vec izracunatu odluku
' ============================================================
' Red odluke se kesira po tip|broj|docID -- dakle po DOKUMENTU, ne po stanju
' podataka. Dok je Scr_ResetCache brisao samo uvid, posle sync-a je ostajala
' STARA odluka: dokument koji je u 10:00 bio bez nizvodnog toka (pa je dobio samo
' "obican storno") zadrzao bi taj red dugmadi i posle sync-a koji mu je doneo
' zbirnu, prijemnicu i palete. Operater bi tako preskocio ceo izbor moda.
'
' StornoRazlog to ne hvata: on pita sme li se dokument stornirati, ne da li sada
' treba framework ispravke.
Private Sub T_StornoAkcije_RefreshInvalidiraOdluku()
    modScrStorno.Scr_IzborTestSet STIP_OTKUP, FX_BLOK, "", ""

    ' Odluka se izracuna i kesira.
    AssertEq (modScrStorno.Scr_BrojAkcija() > 0), True, _
             "preduslov: izabran dokument ima red odluke"
    AssertEq (Len(modScrStorno.Scr_OdlukaKes()) > 0), True, _
             "preduslov: odluka je kesirana"

    ' Promena podataka mora da je ponisti. NAJVAZNIJE u ovom testu.
    modScrStorno.Scr_ResetCache
    AssertEq modScrStorno.Scr_OdlukaKes(), "", _
             "promena podataka ponistava kes odluke -- inace vazi odluka od pre sync-a"
    AssertEq modScrStorno.Scr_IzabranDocID(), "", _
             "i sam izbor, jer se nad zastarelim izborom ne sme odlucivati"
    AssertEq modScrStorno.Scr_BrojAkcija(), 0, _
             "bez izbora nema nijednog dugmeta za mutaciju"
End Sub

' ============================================================
' 62. Bez uvida nema odluke (fail-closed)
' ============================================================
' Ceo smisao ovog ekrana je "prvo vidi posledice, pa odluci". Ako uvid ne uspe,
' dugmad za mutaciju NE SMEJU da se ponude -- inace ekran pita isto sto je i
' MsgBox pitao, samo bez posledica pred sobom.
'
' Scr_IzborTestSet postavlja izbor BEZ gradnje uvida, sto je tacno stanje posle
' neuspelog BuildStornoImpact-a (mImpact ostaje Nothing).
Private Sub T_StornoBezUvida_NemaAkcije()
    ' NAJVAZNIJE PRVO: framework tip bez uvida ne nudi NIJEDNU radnju.
    '
    ' Neuspeh uvida se pravi STVARNO -- zadatom generacijom koju nijedan red ne
    ' nosi. Pod strict rezimom to DIZE gresku, pa uvid ostane prazan, sto je bas
    ' stanje koje kapija treba da pokrije.
    '
    ' Ranije je isto stanje dolazilo otud sto test seam nije gradio uvid. To je
    ' bila greska u testu: merilo se stanje koje u aplikaciji ne postoji, jer
    ' produkcija uvid gradi pri svakom izboru reda.
    modScrStorno.Scr_IzborTestSet STIP_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, "GEN-NE-POSTOJI", ""
    AssertEq modScrStorno.Scr_BrojAkcija(), 0, _
             "framework dokument bez uvida ne nudi nijednu radnju"

    ' Kontrola u suprotnom smeru: ISTI dokument sa razresivim identitetom uvid
    ' dobija, pa radnje postoje. Bez ovoga bi tvrdnja iznad prolazila i kad bi
    ' kapija bila zaglavljena na nuli.
    modScrStorno.Scr_IzborTestSet STIP_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, "GEN-IMP-2", ""
    AssertEq (modScrStorno.Scr_BrojAkcija() > 0), True, _
             "sa valjanim uvidom radnje POSTOJE -- kapija nije zaglavljena"

    ' Tip koji uvid i NEMA (otkup nije framework tip) i dalje nudi obican storno --
    ' inace bi kapija zakljucala i ono sto uvid nikad nije ni imalo.
    modScrStorno.Scr_IzborTestSet STIP_OTKUP, FX_BLOK, "", ""
    AssertEq modScrStorno.Scr_BrojAkcija(), 1, _
             "tip bez uvida i dalje nudi obican storno"

    ' Revers isto: nema uvid po prirodi (list u lancu), ali ima svoja dva izbora.
    modScrStorno.Scr_IzborTestSet STIP_REVERSI, "REV-NEMA", "", DOK_TIP_OM_IZLAZ_KOOP
    AssertEq modScrStorno.Scr_BrojAkcija(), 2, _
             "revers nema uvid po prirodi, pa kapija ne sme da ga zakljuca"

    modScrStorno.Scr_ResetCache
End Sub
' ============================================================
' 63. Necitljiva sekcija cini CEO uvid nevalidnim
' ============================================================
' "valid = True" je ugovor: znaci da je SVIH SEDAM sekcija pouzdano procitano.
' Dok su citaci gutali greske, taj ugovor je bio prazan -- BuildStornoImpact bi
' uredno stigao do kraja i postavio valid = True i kad je npr. paletna sekcija
' vratila praznu kolekciju zato sto kolone nema:
'
'     ne mogu da procitam palete -> prazna Collection -> valid = True
'                                -> ekran kaze "nema paleta" -> nudi mutaciju
'
' A tacan odgovor nije "nema paleta" nego "ne znam da li ih ima".
'
' Drift se pravi STVARNO (preimenovanje kolone), ne simulira -- isti obrazac kao
' test 47. Sema se vraca u istom testu, jer bi inace svi testovi posle ovog merili
' pokvarenu tabelu.
Private Sub T_StornoImpact_SchemaDriftJeInvalidan()
    Dim lo As ListObject, m As Object
    Dim validPodDriftom As Boolean, semaVracena As Boolean, imaoPalete As Boolean
    Dim validPodPalDriftom As Boolean, palSemaVracena As Boolean

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2", "GEN-IMP-2"

    ' POZITIVNA KONTROLA: nad zdravom semom uvid je valjan i ima palete. Bez nje
    ' bi test prosao i kad BuildStornoImpact uvek vraca valid = False.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "GEN-IMP-2", True)
    AssertEq CBool(m("valid")), True, "pozitivna kontrola: zdrava sema daje valjan uvid"
    imaoPalete = (m("palete").count > 0)
    AssertEq imaoPalete, True, "pozitivna kontrola: dokument stvarno ima paletu"

    Set lo = GetTable(TBL_PALETA_STAVKA)
    On Error GoTo VRATI
    lo.ListColumns(COL_PALS_PRIJEMNICA_ID).name = COL_PALS_PRIJEMNICA_ID & "_DRIFT"
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "GEN-IMP-2", True)
    validPodDriftom = CBool(m("valid"))
VRATI:
    On Error Resume Next
    lo.ListColumns(COL_PALS_PRIJEMNICA_ID & "_DRIFT").name = COL_PALS_PRIJEMNICA_ID
    On Error GoTo 0
    semaVracena = (GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID) > 0)

    ' DRIFT KOLONE KOJU CITA SAMO PALETNA SEKCIJA.
    '
    ' Prethodni drift gasi PrijemnicaID -- ali po BAS toj koloni filtrira i
    ' GetChainFlags (kroz CountActive), a on ide RANIJE u BuildStornoImpact. Uvid
    ' zato padne pre nego sto se do paleta uopste stigne, pa tvrdnja iznad meri
    ' tudju kapiju. Mereno: poruka je bila "Kolona PrijemnicaID ne postoji u
    ' tblPaletaStavka" iz modStornoFlow.CountActive.
    '
    ' PaletaID u BuildStornoImpact cita JEDINO GetPaleteImpactByField, pa je ovo
    ' jedina tvrdnja koja meri strogost same paletne sekcije.
    Set lo = GetTable(TBL_PALETA_STAVKA)
    On Error GoTo VRATI_PAL
    lo.ListColumns(COL_PALS_PALETA_ID).name = COL_PALS_PALETA_ID & "_DRIFT"
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "GEN-IMP-2", True)
    validPodPalDriftom = CBool(m("valid"))
VRATI_PAL:
    On Error Resume Next
    lo.ListColumns(COL_PALS_PALETA_ID & "_DRIFT").name = COL_PALS_PALETA_ID
    On Error GoTo 0
    palSemaVracena = (GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID) > 0)

    ' NAJVAZNIJE: necitljiva paletna sekcija cini CEO uvid nevalidnim.
    AssertEq validPodDriftom, False, _
             "necitljiva paletna sekcija cini CEO uvid nevalidnim"
    AssertEq validPodPalDriftom, False, _
             "...i kad nedostaje kolona koju cita SAMO paletna sekcija"
    AssertEq palSemaVracena, True, "i ta sema je vracena posle testa"
    ' Ako sema nije vracena, svi testovi posle ovog mere pokvarenu tabelu.
    AssertEq semaVracena, True, "sema je vracena posle testa"
End Sub

' ============================================================
' 64. Zadat identitet NIKAD ne degradira na broj
' ============================================================
' Druga polovina istog ugovora. Kad je docID zadat a ne moze da se razresi,
' povratak na poslovni broj vraca tacno ono sto je #198 vadio -- i to unutar
' modela koji se posle oznacava kao valid.
'
' Prazan docID je DRUGA prica i mora da nastavi da radi: zatecen zapis nema
' generaciju, pa je broj sve sto postoji. Test meri obe strane.
Private Sub T_StornoImpact_IdentitetNeDegradira()
    Dim m As Object

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z1", "GEN-IMP-1"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2", "GEN-IMP-2"

    ' Identitet KOJI NE POSTOJI. Sema je zdrava, kolona generacije je tu -- samo
    ' nijedan red tog broja ne nosi ovu generaciju. To je stanje u kome se sme
    ' uraditi tacno jedna stvar: stati. Povratak na broj bi dao palete OBA
    ' dokumenta, i to unutar modela koji se posle oznacava kao valid.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "GEN-NE-POSTOJI", True)
    AssertEq CBool(m("valid")), False, _
             "zadat identitet koji se ne moze razresiti obara uvid, ne pada na broj"

    ' GENERACIJA KOJA PRIPADA DRUGOM BROJU.
    '
    ' Prethodna tvrdnja meri kapiju u modStornoFlow.PkPoIdentitetu: generacija
    ' koje NEMA nigde. Ova meri kapiju u modStornoImpact.ImpactPalete, i to je
    ' druga kapija -- jer se dva trazenja razlikuju:
    '
    '   IdoviGeneracije          trazi generaciju kroz CELU tabelu
    '   PrijemniceIDPoIdentitetu trazi broj I generaciju
    '
    ' GEN-IMP-1 postoji (na PRJ-TEST-Z1, broj 6/150326), pa ga prvo trazenje
    ' razresi i uvid ide dalje. Drugo ga ne nadje pod brojem 1/150326 -- i bez
    ' kapije bi palete bile procitane PO BROJU, dakle tudje, unutar modela koji
    ' se posle oznacava kao valid. Tacno degradacija zbog koje je nastao 198.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_BROJ, _
                                             "", "GEN-IMP-1", True)
    AssertEq CBool(m("valid")), False, _
             "generacija koja pripada DRUGOM broju ne tumaci ovaj dokument"

    ' Bez identiteta uvid i dalje radi po broju -- zatecen zapis nema generaciju,
    ' pa je broj sve sto postoji. Kapija ne sme da zakljuca ni taj slucaj.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "", True)
    AssertEq CBool(m("valid")), True, _
             "bez identiteta uvid i dalje radi po broju (zatecen zapis)"
    AssertEq m("palete").count, 2, "i tada legitimno vidi oba dokumenta tog broja"
End Sub
' ============================================================
' 65. Necitljiva BLOCK sekcija obara ceo uvid
' ============================================================
' Test 63 je pokrio paletnu sekciju, koju cita modStornoImpact. Block sekcija
' dolazi iz modStornoFlow (GetStornoBlockRows -> ActiveBlocksForFlow), i tamo je
' fail-open obrazac ziveo jos jednu rundu duze:
'
'     If cId = 0 Then Exit Function        ' nedostaje OtkupID
'     EH: LogErr ... : End Function        ' greska -> prazan spisak
'
' Za operatera to znaci poruku "nema pogodjenih blokova" nad odlukom koja
' blokove STORNIRA. Prazan spisak sme da znaci samo "uspesno sam proverio i
' nema ih", nikad "ne umem da proverim".
'
' Drift se pravi STVARNO (preimenovanje kolone), i sema se vraca u istom testu.
Private Sub T_StornoImpact_BlokSekcijaDriftJeInvalidna()
    Dim lo As ListObject, m As Object
    Dim validPodDriftom As Boolean, semaVracena As Boolean, imaoBlokove As Boolean

    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-A", "GEN-BLK-B"
    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-B", "GEN-BLK-B"

    ' POZITIVNA KONTROLA: nad zdravom semom uvid je valjan i blokovi POSTOJE.
    ' Bez nje bi test prosao i kad BuildStornoImpact uvek vraca valid = False.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_OTPREMNICA, FX_OTPREMNICA_BLOK, _
                                             "", "GEN-BLK-B", True)
    AssertEq CBool(m("valid")), True, "pozitivna kontrola: zdrava sema daje valjan uvid"
    imaoBlokove = (m("blocks").count > 0)
    AssertEq imaoBlokove, True, "pozitivna kontrola: dokument stvarno ima otkupni blok"

    Set lo = GetTable(TBL_OTKUP)
    On Error GoTo VRATI
    lo.ListColumns(COL_OTK_ID).name = COL_OTK_ID & "_DRIFT"
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_OTPREMNICA, FX_OTPREMNICA_BLOK, _
                                             "", "GEN-BLK-B", True)
    validPodDriftom = CBool(m("valid"))
VRATI:
    On Error Resume Next
    lo.ListColumns(COL_OTK_ID & "_DRIFT").name = COL_OTK_ID
    On Error GoTo 0
    semaVracena = (GetColumnIndex(TBL_OTKUP, COL_OTK_ID) > 0)

    ' NAJVAZNIJE: necitljiva block sekcija obara CEO uvid.
    AssertEq validPodDriftom, False, "necitljiva block sekcija obara CEO uvid"
    ' Ako sema nije vracena, svi testovi posle ovog mere pokvarenu tabelu.
    AssertEq semaVracena, True, "sema je vracena posle testa"
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
' ============================================================
' 66. Ekran ne pusta obradjenu gresku u ljusku
' ============================================================
' Operater je posle USPESNE ispravke dobijao crven toast:
'
'     X Radnja nije uspela: modScrStorno.Scr_Event scrStA0
'
' preko uredno popunjene forme. Uzrok nije bio pad radnje nego zivot Err-a:
' "On Error Resume Next" PRIGUSUJE gresku, ali je NE BRISE. OtvoriIspravku ima
' bas takav gard, pa je greska prigusena u njemu prezivela povratak kroz
' StornoPoModu i PokreniAkciju sve do modUiScreens.ScrEvent -- a on posle
' Application.Run cita Err.Number i, ako nije nula, javlja neuspeh.
'
' Dodavanje Err.Clear u EH handlere to NIJE resilo: EH se na uspesnom putu
' uopste ne izvrsava. Zato Scr_Event sada ima JEDAN izlaz, i na njemu cisti Err.
'
' Test ide kroz tag koji sigurno prolazi kroz "On Error Resume Next" region
' (OsveziZonu -> Zona -> ScreenZone, a forma u testu nije izgradjena).
Private Sub T_StornoEkran_NeCuriGreska()
    Dim brojPosle As Long

    modScrStorno.Scr_IzborTestSet STIP_OTKUP, FX_BLOK, "", ""

    ' Handler koji je gresku PROGUTAO -- bez toga se Err.Clear u Scr_Event ne
    ' moze izmeriti: nijedna danasnja grana ne ostavlja Err ziv, pa je tvrdnja
    ' bila zelena i kad tog Err.Clear nema.
    modScrStorno.Scr_ErrTestPrljav
    Err.Clear
    modScrStorno.Scr_Event "scrStPal", "Click"
    ' Err se cita ODMAH: svaki poziv ispod (pa i AssertEq) ume da ga promeni.
    brojPosle = Err.Number
    Err.Clear

    AssertEq brojPosle, 0, _
             "Scr_Event vraca cist Err -- inace ljuska javi neuspeh za radnju koja je prosla"

    ' Kontrola u drugom smeru: prekidac je stvarno obradjen, nije se samo
    ' progutao. Bez ovoga bi test prosao i kad Scr_Event ne radi nista.
    modScrStorno.Scr_Event "lsOTPREMNICA", "Click"
    AssertEq modScrStorno.Scr_Lista(), STIP_OTPREMNICA, _
             "kontrola: Scr_Event i dalje obradjuje dogadjaj"

    modScrStorno.Scr_ResetCache
    modScrStorno.Scr_TipTestSet STIP_OTKUP
End Sub
' ============================================================
' 67. Blok sekcija PRIJEMNICE (preko zbirne) -- druga grana istog dispecera
' ============================================================
' Test 65 je pokrio OTPREMNICU, koja u ActiveBlocksForFlow ide kroz
' GetBlokOtkupIDs. Zbirna i prijemnica idu kroz ActiveOtkupIDsByZbirna -- i tamo
' se strict gubio jos jednu rundu:
'
'     tblOtkup.BrojZbirne drift -> ActiveOtkupIDsByZbirna vrati prazno
'                                -> GetStornoBlockRows izadje na ids.count = 0
'                                -> dakle PRE svoje kapije
'                                -> blocks = 0, valid = True
'
' Isti kvar kao u 65, samo druga grana istog Select Case-a. Zato test 65 nije
' bio dovoljan: on tu granu uopste ne dodiruje.
'
' PRJ-TEST-Z2 ide preko zbirne ZB-TEST-4, koja nosi otkupne blokove.
Private Sub T_StornoImpact_PrijemnicaBlokDriftJeInvalidan()
    Dim lo As ListObject, m As Object
    Dim validZbirna As Boolean, validPrij As Boolean
    Dim semaVracena As Boolean, imaoBlokove As Boolean

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2", "GEN-IMP-2"

    ' POZITIVNA KONTROLA nad ZBIRNOM: ona u fixture-u pouzdano nosi aktivan
    ' otkupni blok (OTK-TEST-1 na ZB-TEST-1). Prijemnica se za pozitivnu kontrolu
    ' ne koristi: njene blokove raniji testovi u ovom istom prolazu storniraju,
    ' pa bi kontrola merila redosled testova umesto pravila.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_ZBIRNA, FX_ZBIRNA, "", "", True)
    AssertEq CBool(m("valid")), True, "pozitivna kontrola: zdrava sema daje valjan uvid"
    imaoBlokove = (m("blocks").count > 0)
    AssertEq imaoBlokove, True, "pozitivna kontrola: zbirna stvarno nosi otkupni blok"

    Set lo = GetTable(TBL_OTKUP)
    On Error GoTo VRATI
    lo.ListColumns(COL_OTK_BROJ_ZBIRNE).name = COL_OTK_BROJ_ZBIRNE & "_DRIFT"
    ' Obe grane koje idu kroz ActiveOtkupIDsByZbirna, ne samo jedna: zbirna
    ' direktno, prijemnica preko svoje zbirne. Test 65 pokriva TRECU granu
    ' (otpremnica -> GetBlokOtkupIDs) i ove dve ne dodiruje.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_ZBIRNA, FX_ZBIRNA, "", "", True)
    validZbirna = CBool(m("valid"))
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "GEN-IMP-2", True)
    validPrij = CBool(m("valid"))
VRATI:
    On Error Resume Next
    lo.ListColumns(COL_OTK_BROJ_ZBIRNE & "_DRIFT").name = COL_OTK_BROJ_ZBIRNE
    On Error GoTo 0
    semaVracena = (GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE) > 0)

    ' NAJVAZNIJE: obe grane obaraju uvid, ne samo grana otpremnice iz testa 65.
    AssertEq validZbirna, False, "necitljiva blok sekcija ZBIRNE obara CEO uvid"
    AssertEq validPrij, False, "necitljiva blok sekcija PRIJEMNICE obara CEO uvid"
    ' Ako sema nije vracena, svi testovi posle ovog mere pokvarenu tabelu.
    AssertEq semaVracena, True, "sema je vracena posle testa"
End Sub
' ============================================================
' 68. "On Error Resume Next" resetuje Err -- pa LogErr posle njega ne pise
' ============================================================
' Operater je prijavio pad upisa otpremnice, a Log fajl je bio PRAZAN -- nijedna
' ERROR linija. Uzrok nije u upisu nego u dijagnostici:
'
'     LogErr pise samo "If Err.Number <> 0"
'     a EH blokovi rade:  errDesc = Err.description
'                         On Error Resume Next      <- resetuje Err
'                         LogErr "SaveOtpremnicaMulti_TX"   <- vidi 0, ne pise
'
' Ovaj test ne meri nas kod nego SEMANTIKU VBA na kojoj taj zakljucak stoji.
' Ako VBA to jednog dana promeni, tvrdnja pada ovde, a ne kroz prazan log posle
' incidenta.
Private Sub T_LogErr_NeVidiErrPosleResumeNext()
    Dim preN As Long, posleN As Long

    On Error Resume Next
    Err.Raise 5, "T_LogErr", "namerna greska"
    preN = Err.Number
    ' isti potez koji EH blokovi rade pre poziva LogErr-a
    On Error Resume Next
    posleN = Err.Number
    Err.Clear
    On Error GoTo 0

    AssertEq (preN <> 0), True, "preduslov: greska je stvarno podignuta"
    AssertEq posleN, 0, _
             "'On Error Resume Next' resetuje Err -- LogErr posle njega nema sta da vidi"
End Sub
' ============================================================
' 69. Upozorenje uz uspesan upis mora da nosi svoju oznaku
' ============================================================
' Dokument moze da bude snimljen, a da uz njega NESTO ne prodje: prevezivanje
' paleta, auto-zbirna, ili zavrsetak ispravke koji stane na safe-stopu ("vise
' ispravki na cekanju"). Te poruke stizu u isti izlazni string kao i obicna
' informacija o uspehu.
'
' CommitDokument ih razdvaja po OZNACI: ChrW(10007) ide i u MsgBox (operater
' mora da vidi da mu je ostao posao), ChrW(10003) ostaje u toastu. Ako neko doda
' novo upozorenje bez te oznake, ono ce se tiho izgubiti -- toast ga sece, a
' uspesan toast se jos i sam sakrije posle cetiri sekunde.
'
' Test cuva bas tu podelu, jer se ona iz koda ne vidi -- oba su samo stringovi.
Private Sub T_PorukeUnosa_UpozorenjeNosiOznaku()
    Dim upoz As Variant, info As Variant, i As Long, t As String

    ' Katalog se PRVO osvezava iz koda: Poruka() cita tblPoruke, a fixture nosi
    ' onaj katalog kakav je bio u donoru -- bez ovoga bi test merio zatecene
    ' podatke umesto ugovora iz UpsertPoruke. EnsurePoruke je idempotentan i
    ' bez MsgBox-a (isti obrazac koji vec koristi Test_PorukeKatalogPokrivaDokumenta).
    modSetup.EnsurePoruke
    modPoruke.InvalidateCache

    upoz = Array("DOKUNOS_MSG_VISE_ISPRAVKI", "DOKUNOS_MSG_PALETE_NISU", _
                 "DOKUNOS_MSG_ZBIRNA_NIJE", "DOKUNOS_MSG_ISPRAVKA_NIJE")
    For i = 0 To UBound(upoz)
        t = Poruka(CStr(upoz(i)))
        AssertEq (Len(t) > 0), True, CStr(upoz(i)) & " postoji u katalogu"
        AssertEq (InStr(1, t, ChrW(10007)) > 0), True, _
                 CStr(upoz(i)) & " nosi oznaku upozorenja -- inace se ne vidi"
    Next i

    ' Druga strana: cista informacija NE SME da nosi oznaku upozorenja, inace bi
    ' svaki uspesan upis otvarao dijalog bez razloga.
    info = Array("DOKUNOS_MSG_PALETE_PREVEZANE", "DOKUNOS_MSG_ISPRAVKA_OK")
    For i = 0 To UBound(info)
        t = Poruka(CStr(info(i)))
        AssertEq (Len(t) > 0), True, CStr(info(i)) & " postoji u katalogu"
        AssertEq (InStr(1, t, ChrW(10007)) > 0), False, _
                 CStr(info(i)) & " je informacija, ne upozorenje"
    Next i

    ' TRECA tvrdnja: oznaka je SIGNAL ZA RUTIRANJE, ne deo recenice. MsgBox crta
    ' kroz ANSI kodnu stranu u kojoj ChrW(10007) ne postoji, pa ju je operater
    ' video kao vodece '?' ispred teksta. Pred dijalog se zato skida -- a u traci
    ' poruka, koja je Unicode, OSTAJE, jer tamo nosi znacenje.
    For i = 0 To UBound(upoz)
        t = Poruka(CStr(upoz(i)))
        AssertEq (InStr(1, modOtkupUI.PorukaZaDijalog(t), ChrW(10007)) > 0), False, _
                 CStr(upoz(i)) & " u dijalogu ide BEZ oznake"
        AssertEq (Len(modOtkupUI.PorukaZaDijalog(t)) > 0), True, _
                 CStr(upoz(i)) & " posle skidanja oznake nije prazna"
        AssertEq Left$(modOtkupUI.PorukaZaDijalog(t), 1), UCase$(Left$(modOtkupUI.PorukaZaDijalog(t), 1)), _
                 CStr(upoz(i)) & " u dijalogu pocinje slovom, ne razmakom"
    Next i
End Sub

' ============================================================
' 70. Nestao identitet obara uvid -- i za otpremnicu i za zbirnu
' ============================================================
' Test 64 je pokrio PRIJEMNICU, gde identitet cuva ImpactPalete. Otpremnica i
' zbirna idu kroz PkPoIdentitetu, koji je dobio parametar strict ali ga NIJE
' koristio:
'
'     If ids.count = 0 Then Exit Function     ' komentar iznad tvrdi da je greska
'
' Nizvodno je to izgledalo kao "dokument ne postoji" umesto "ne mogu da ga
' razresim" -- a model se posle svega oznacavao kao valid. Zbirna je uz to
' prekidala propagaciju i u ScanZbirna, koji strict nije ni prosledjivao.
'
' Scenario je stvaran: mreza nosi broj i docID iz reda, a dokument te generacije
' je u medjuvremenu nestao (storniran, prevezan, obrisan).
Private Sub T_StornoImpact_NestaoIdentitetJeInvalidan()
    Dim m As Object

    StampGeneraciju TBL_OTPREMNICA, COL_OTP_ID, "OTP-BLK-B", "GEN-BLK-B"

    ' POZITIVNA KONTROLA: postojeca generacija daje valjan uvid. Bez nje bi test
    ' prosao i kad BuildStornoImpact uvek vraca False.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_OTPREMNICA, FX_OTPREMNICA_BLOK, _
                                             "", "GEN-BLK-B", True)
    AssertEq CBool(m("valid")), True, "pozitivna kontrola: postojeca generacija daje valjan uvid"

    ' NAJVAZNIJE: generacija koje NEMA obara uvid, umesto da prodje kao
    ' "dokument ne postoji".
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_OTPREMNICA, FX_OTPREMNICA_BLOK, _
                                             "", "GEN-NE-POSTOJI", True)
    AssertEq CBool(m("valid")), False, _
             "nestao identitet OTPREMNICE obara uvid"

    ' Zbirna je isla svojim putem (ScanZbirna nije prosledjivao strict), pa se
    ' meri zasebno -- ista tvrdnja, druga grana.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_ZBIRNA, FX_ZBIRNA, _
                                             "", "GEN-NE-POSTOJI", True)
    AssertEq CBool(m("valid")), False, _
             "nestao identitet ZBIRNE obara uvid"

    ' Bez identiteta oba i dalje rade po broju -- zatecen zapis nema generaciju.
    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_ZBIRNA, FX_ZBIRNA, "", "", True)
    AssertEq CBool(m("valid")), True, "bez identiteta zbirna i dalje radi po broju"
End Sub

' ============================================================
' 71. Odbaci zaostalu ispravku -- red nosi CorrectionID, ne samo broj
' ============================================================
' Ekran Oporavak je listu "Nedovrseno" prikazivao kao cist pregled: operater
' vidi da ga safe-stop blokira, a nema cime da to razresi. Jedini izlaz je bio
' legacy frmDokumenta. CancelCorrectionContext je postojao sve vreme -- falio
' mu je ulaz iz novog UI-ja.
'
' Radnja MORA da cilja CorrectionID, ne poslovni broj. Nad istim brojem moze da
' stoji vise contexta (storno, pa opet storno istog dokumenta), pa bi izbor po
' broju zatvorio onaj koji zatekne prvi -- a operater je gledao drugi red.
' Zato red nosi identitet u nevidljivoj koloni, isto kao GeneracijaID na ekranu
' Storno.
Private Sub T_Oporavak_OdbaciIspravku_PoIdentitetu()
    Dim d As Variant, cols As Variant, r As Variant
    Dim i As Long, n As Long, ctxRedova As Long, saCID As Long
    Dim vidjen1 As Boolean, vidjen2 As Boolean

    modScrOporavak.Scr_OpoTestSet "NEDOVRSENO", "", ""
    d = modScrOporavak.Scr_Rows("", "")
    AssertEq IsArray(d), True, "Nedovrseno vraca niz"
    cols = d(0)

    ' Kolona identiteta je POSLEDNJA i NEVIDLJIVA. Prioritet 4 nikad ne prolazi
    ' petlju vidljivosti (ide 3 -> 1), pa je operater ne vidi, a GridCell je cita
    ' iz mView -- sirina kolone ne dira podatak.
    AssertEq (UBound(cols) + 1), modScrOporavak.NED_COL_CID, _
             "opis kolona se zavrsava BAS na koloni koju radnja cita"
    AssertEq Split(CStr(cols(modScrOporavak.NED_COL_CID - 1)), "|")(0), "OTKUI_HDO_CID", _
             "na toj koloni stoji CorrectionID"
    AssertEq modScrDokumenti.ColF(CStr(cols(UBound(cols))), 4), "4", _
             "kolona CID je prioriteta 4 -- nikad vidljiva"

    ' PREDUSLOV: fixture stvarno ima ispravke na cekanju. Bez ovoga bi petlja
    ' ispod prosla nad nula redova i test bi bio zelen ne merivsi nista.
    n = CLng(d(2))
    AssertEq (n >= 2), True, "fixture ima bar dve stavke u Nedovrsenom"

    ' Svaki CONTEXT red nosi svoj CorrectionID; osirotele stavke ga nemaju --
    ' one se ne odbacuju nego prevezuju, pa radnja nad njima mora da stane.
    r = d(1)
    For i = 1 To n
        If Left$(CStr(r(i, 2)), 8) = "CONTEXT/" Then
            ctxRedova = ctxRedova + 1
            If Len(Trim$(CStr(r(i, modScrOporavak.NED_COL_CID)))) > 0 Then saCID = saCID + 1
            If CStr(r(i, modScrOporavak.NED_COL_CID)) = "SV-TEST-1" Then vidjen1 = True
            If CStr(r(i, modScrOporavak.NED_COL_CID)) = "SV-TEST-2" Then vidjen2 = True
        Else
            AssertEq CStr(r(i, modScrOporavak.NED_COL_CID)), "", _
                     "osirotela stavka nema CorrectionID -- resava se prevezivanjem"
        End If
    Next i

    AssertEq (ctxRedova >= 2), True, "fixture ima bar dva context reda"
    AssertEq saCID, ctxRedova, _
             "SVAKI context red nosi svoj CorrectionID u koloni 6"

    ' I to bas ONE iz fixture-a: dva razlicita ID-ja, ne dva puta isti. Test koji
    ' bi merio samo "nije prazno" prosao bi i kad bi svi redovi nosili isti CID.
    AssertEq vidjen1, True, "red za SV-TEST-1 nosi svoj identitet"
    AssertEq vidjen2, True, "red za SV-TEST-2 nosi svoj identitet"
End Sub

' ============================================================
' 72. Odbacivanje gasi IZABRANU ispravku -- i nijednu drugu
' ============================================================
' Test 71 dokazuje da identitet STIGNE do reda mreze. To nije isto sto i
' "radnja gadja bas njega": hard-kodovan `CancelCorrectionContext("SV-TEST-1")`,
' ili `GridCell(red - 1, ...)`, prosli bi 71 netaknuti. Ovde se meri POSLEDICA.
'
' MsgBox u headless runu visi, pa se ne zove `OdbaciIspravku` nego njegovo
' jezgro -- sve osim potvrde i toast-a.
'
' Test MUTIRA podatke i zato ih VRACA: fixture nosi tacno dve ispravke na
' cekanju, a test 25 bas na tome meri safe-stop ("dve ili vise = ne biraj
' naslepo"). Bez vracanja bi ovaj test menjao ishod tudjeg, zavisno od redosleda.
Private Sub T_Oporavak_OdbaciIspravku_GasiSamoSvoj()
    Dim d As Variant, r As Variant, i As Long, n As Long
    Dim ostao1 As Boolean, ostao2 As Boolean

    ' PREDUSLOV: oba su na cekanju. Bez ovoga bi test merio zatecen ostatak
    ' ranijeg testa umesto posledice ove radnje.
    AssertEq SvPolje("SV-TEST-1", COL_SV_STATUS), SV_STATUS_PENDING, _
             "SV-TEST-1 je na cekanju PRE radnje"
    AssertEq SvPolje("SV-TEST-2", COL_SV_STATUS), SV_STATUS_PENDING, _
             "SV-TEST-2 je na cekanju PRE radnje"

    AssertEq modScrOporavak.OdbaciIspravkuCore("SV-TEST-2"), True, _
             "odbacivanje je proslo"

    ' NAJVAZNIJE PRVO: sused je NETAKNUT. Radnja koja gadja prvi red, susedni
    ' red ili poslovni broj pada bas ovde -- a sve tri bi prosle tvrdnju koja
    ' meri samo da je izabrani ugasen.
    AssertEq SvPolje("SV-TEST-1", COL_SV_STATUS), SV_STATUS_PENDING, _
             "SV-TEST-1 ostaje netaknut"
    AssertEq SvPolje("SV-TEST-1", COL_SV_NEEDS_RECOVERY), "Da", _
             "SV-TEST-1 i dalje ceka zamenski dokument"

    ' I tek onda: izabrani JESTE ugasen, u oba polja.
    AssertEq SvPolje("SV-TEST-2", COL_SV_STATUS), SV_STATUS_CANCELLED, _
             "SV-TEST-2 je otkazan"
    AssertEq SvPolje("SV-TEST-2", COL_SV_NEEDS_RECOVERY), "Ne", _
             "SV-TEST-2 vise ne ceka nista"

    ' Posledica koju operater vidi: red nestaje iz liste, drugi ostaje.
    modScrOporavak.Scr_OpoTestSet "NEDOVRSENO", "", ""
    d = modScrOporavak.Scr_Rows("", "")
    n = CLng(d(2))
    AssertEq (n > 0), True, "lista Nedovrseno nije prazna posle radnje"
    r = d(1)
    For i = 1 To n
        If CStr(r(i, modScrOporavak.NED_COL_CID)) = "SV-TEST-1" Then ostao1 = True
        If CStr(r(i, modScrOporavak.NED_COL_CID)) = "SV-TEST-2" Then ostao2 = True
    Next i
    AssertEq ostao1, True, "SV-TEST-1 je i dalje u listi Nedovrseno"
    AssertEq ostao2, False, "SV-TEST-2 je nestao iz liste Nedovrseno"

    ' Ciscenje se i PROVERAVA. Nevereno vracanje je isto sto i nikakvo: test
    ' dodat ispod ovog nasledio bi tiho izmenjen fixture, a pao bi po tudjem
    ' imenu.
    VratiContextNaCekanje "SV-TEST-2"
    AssertEq SvPolje("SV-TEST-2", COL_SV_STATUS), SV_STATUS_PENDING, _
             "fixture je vracen: SV-TEST-2 je opet na cekanju"
    AssertEq SvPolje("SV-TEST-2", COL_SV_NEEDS_RECOVERY), "Da", _
             "fixture je vracen: SV-TEST-2 opet ceka zamenski dokument"
End Sub

' ============================================================
' 73. Uvid o paleti cita zaglavlje BAS te palete
' ============================================================
' Operater je prijavio 2-3 sekunde po kliku na red. Merenje je pokazalo da 95%
' vremena odlazi na sekciju paleta: svaka paleta u rezultatu je izazivala TRI
' linearna prolaza kroz tblPaleta i TRI kopije cele tabele. Tabela se sada cita
' jednom, a red se nalazi kroz recnik.
'
' Zatecena suita to NE bi uhvatila: postojeci testovi tvrde samo KOLIKO paleta
' uvid nosi i koliki im je zbir -- a zbir dolazi iz druge petlje, koju izmena ne
' dira. Polja iz zaglavlja palete (oznaka, popunjenost, kapacitet, neto,
' preradjenost) nije merilo nista, a bas njih izmena preracunava.
'
' Fixture: PAL-TEST-Z2 je paleta 12/2026, 20 gajbi od 100, 200 kg.
Private Sub T_ImpactPalete_ZaglavljeIzPraveVrste()
    Dim m As Object, pal As Collection, d As Object

    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z1", "GEN-IMP-1"
    StampGeneraciju TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-TEST-Z2", "GEN-IMP-2"

    Set m = modStornoImpact.BuildStornoImpact(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA, _
                                             "", "GEN-IMP-2")
    Set pal = m("palete")

    ' PREDUSLOV: tacno jedna paleta, i to njegova. Bez ovoga bi tvrdnje ispod
    ' merile pogresan red i prolazile iz pogresnog razloga.
    AssertEq pal.count, 1, "preduslov: uvid nosi tacno paletu izabranog dokumenta"
    Set d = pal(1)
    AssertEq CStr(d("paletaID")), "PAL-TEST-Z2", "preduslov: i to bas PAL-TEST-Z2"

    ' NAJVAZNIJE: zaglavlje dolazi iz REDA TE palete. Radnja koja bi uzela prvi
    ' red tabele, ili susedni, pada ovde -- a prosla bi svaki postojeci test, jer
    ' oni broje palete i sabiraju stavke, a zaglavlje ne diraju.
    AssertEq CLng(d("used")), 20, "popunjenost je iz reda BAS te palete"
    AssertEq CLng(d("cap")), 100, "kapacitet je iz istog reda"
    AssertEq CDbl(d("neto")), 200#, "neto je iz istog reda"
    AssertEq CStr(d("label")), "12/2026", "oznaka je broj/godina iz istog reda"
    AssertEq CBool(d("preradjena")), False, "preradjenost je iz istog reda"

    ' I da se zaglavlje ne pobrka sa zbirom stavki: to su dva razlicita racuna nad
    ' istom paletom -- koliko je na njoj UKUPNO, i koliko od toga nosi OVAJ
    ' dokument. Ovde se poklapaju, ali dolaze iz razlicitih tabela.
    AssertEq CLng(d("thisGajb")), 20, "zbir stavki OVOG dokumenta ostaje svoj racun"
    AssertEq CDbl(d("thisNeto")), 200#, "isto i za kilograme"
End Sub

' ============================================================
' 74. Efekat storna se sklapa IZ KATALOGA, ne iz literala
' ============================================================
' Tekstovi posledica su do v6-ui-148 bili ASCII literali u modStornoFlow. Zato su
' i bili pisani telegrafski ("preracun, NE pada") -- bez dijakritike se poslovna
' recenica ne moze napisati, a VBA izvor mora ostati ASCII.
'
' Selidba u katalog uvodi nov nacin da se ekran pokvari TIHO: kljuc koji katalog
' ne zna vraca prazan string, pa najvaznija kolona ekrana ostane prazna, bez
' greske i bez traga. vba_check hvata kljuc bez para u UpsertPoruke, ali ne i
' katalog koji nije osvezen -- a to je bas ono sto se desava posle importa.
Private Sub T_StornoEfekat_TekstIzKataloga()
    Dim ch As Collection, red As Variant
    Dim i As Long, spojenih As Long, razdvojenih As Long

    ' Katalog se osvezava izricito: test meri TABELU poruka, a ne kes koji je
    ' zatekao (isti razlog kao u testu 69).
    modSetup.EnsurePoruke
    modPoruke.InvalidateCache

    ' PREDUSLOV: katalog stvarno nosi tekst. Da kljuc nedostaje, sve tvrdnje ispod
    ' bi poredile prazno sa praznim i prosle.
    AssertEq (Len(Poruka("STEF_STORNO_AMB")) > 10), True, _
             "katalog nosi tekst efekta, nije prazan kljuc"
    AssertEq (Len(Poruka("STEF_PRE_OBA")) > 5), True, _
             "katalog nosi spojen prefiks odluke"

    Set ch = modStornoFlow.GetStornoChainRows(FLOW_DOC_PRIJEMNICA, FX_PRIJ_ZBR_KOLIZIJA)
    AssertEq (ch.count > 0), True, "lanac ima bar jedan red"

    ' NAJVAZNIJE: napomena prvog reda je sklopljena iz kataloga. Prvi red je sam
    ' dokument, a njegov efekat je isti za oba osnova -- pa mora doci kroz SPOJEN
    ' prefiks, ne kao dva razdvojena.
    red = ch(1)
    AssertEq CStr(red(2)), Poruka("STEF_PRE_OBA") & Poruka("STEF_STORNO_AMB"), _
             "napomena se sklapa iz kataloga, sa spojenim prefiksom"


    ' I obrnut slucaj: gde se osnovi RAZLIKUJU moraju se videti OBA prefiksa.
    ' Bez ove tvrdnje bi ChainEff koji uvek spaja prosao neprimecen. Ne trazi se
    ' odredjen red -- lanac otpremnice ih nosi oba, a redosled nije predmet ovog
    ' testa.
    Set ch = modStornoFlow.GetStornoChainRows(FLOW_DOC_OTPREMNICA, FX_BROJ_OTP)
    For i = 1 To ch.count
        red = ch(i)
        If InStr(CStr(red(2)), Poruka("STEF_PRE_OBA")) = 1 Then spojenih = spojenih + 1
        If InStr(CStr(red(2)), Poruka("STEF_PRE_DUPLI")) = 1 And _
           InStr(CStr(red(2)), Poruka("STEF_PRE_PONIST")) > 1 Then razdvojenih = razdvojenih + 1
    Next i
    AssertEq (spojenih > 0), True, _
             "isti efekat za oba osnova ide kroz JEDAN spojen prefiks"
    AssertEq (razdvojenih > 0), True, _
             "razlicit efekat nosi OBA prefiksa u istom redu"
End Sub

' ============================================================
' 75. Otkupni blokovi: podrazumevano NIJEDAN, i storno gadja bas oznacene
' ============================================================
' Legacy panel je imao multiselect: podrazumevano nijedan blok nije cekiran, a
' cekiran je znacio DODATNO storniran. Nov ekran je taj izbor izgubio i na
' potvrdu stornirao SVE blokove -- dakle bio je destruktivniji od legacy-ja, i to
' ne namerno nego zato sto multiselect nije bio prenet.
'
' Ovaj test meri dve stvari koje suita inace ne bi videla:
'   1. da je podrazumevano stanje PRAZNO (nijedan oznacen),
'   2. da oznacavanje gadja BAS taj blok, po OtkupID-u a ne po broju otkupa.
'
' Druga tvrdnja nije formalnost: broj otkupa se racuna po kooperantu, pa dva
' bloka lako dele isti -- a spisak zavrsava u StornoSelectedBlocks_TX, u mutaciji.
Private Sub T_StornoBlokovi_PodrazumevanoNijedan()
    Dim d As Variant, r As Variant, n As Long, i As Long, ident As String

    modScrStorno.Scr_IzborTestSet STIP_OTPREMNICA, FX_BROJ_OTP, "", ""
    modScrStorno.Scr_TipTestSet "BLOKOVI"
    d = modScrStorno.Scr_Rows("sve", "")
    n = CLng(d(2))

    ' PREDUSLOV: dokument stvarno ima blokove. Bez ovoga bi tvrdnje ispod merile
    ' prazan skup i prolazile iz pogresnog razloga.
    AssertEq (n > 0), True, "preduslov: izabrana otpremnica nosi otkupne blokove"
    r = d(1)

    ' NAJVAZNIJE: podrazumevano NIJEDAN nije oznacen. Kolona izbora je prazna u
    ' svakom redu, i brojac to potvrdjuje.
    For i = 1 To n
        AssertEq CStr(r(i, 1)), "", "red " & i & " nije oznacen bez izricitog izbora"
    Next i
    AssertEq modScrStorno.BlokOznacenihBroj(), 0, "brojac oznacenih je nula"

    ' Oznaci TACNO jedan, po identitetu iz nevidljive kolone.
    ident = Trim$(CStr(r(1, modScrStorno.ST_BLOK_COL_ID)))
    AssertEq (Len(ident) > 0), True, "red nosi OtkupID u nevidljivoj koloni"
    modScrStorno.Scr_BlokTestSet ident

    d = modScrStorno.Scr_Rows("sve", "")
    r = d(1)
    AssertEq modScrStorno.BlokOznacenihBroj(), 1, "oznacen je tacno jedan blok"
    AssertEq CStr(r(1, 1)), ChrW(10003), "oznaceni red nosi kvacicu"
    For i = 2 To CLng(d(2))
        AssertEq CStr(r(i, 1)), "", "ostali redovi ostaju neoznaceni"
    Next i

    ' Promena izabranog dokumenta PONISTAVA izbor: oznake pripadaju dokumentu nad
    ' kojim su napravljene. Ostavljene bi na sledecem stornirale blokove koje
    ' operater nikad nije video.
    modScrStorno.Scr_IzborTestSet STIP_OTPREMNICA, FX_BROJ_OTP, "", ""
    AssertEq modScrStorno.BlokOznacenihBroj(), 0, _
             "promena izbora dokumenta ponistava oznacene blokove"

    ' Red o blokovima u zoni mora da nosi STANJE, ne pravilo. Operater je prijavio
    ' da cip Blokovi postoji, ali da nista ne kaze da tamo ima sta da se odluci --
    ' pa taj red menja tekst prema izboru.
    AssertEq (InStr(modScrStorno.BlokStatusTekst(), Poruka("STEF_BLOK_BIRAJ_1")) = 1), True, _
             "bez izbora red poziva na izbor i imenuje listu"
    modScrStorno.Scr_BlokTestSet ident
    AssertEq (InStr(modScrStorno.BlokStatusTekst(), Poruka("STEF_BLOK_IZABRANO_1")) = 1), True, _
             "sa izborom red prijavljuje KOLIKO ih je izabrano"
    AssertEq (InStr(modScrStorno.BlokStatusTekst(), Poruka("STEF_BLOK_BIRAJ_1")) > 0), False, _
             "i vise ne poziva na izbor koji je vec napravljen"

    modScrStorno.Scr_TipTestSet STIP_OTKUP
End Sub

' ============================================================
' 76. Brojac uz stavku menija: opcion deo ugovora, bez imena ekrana
' ============================================================
' Operater je pitao ima li smisla da Storno i Oporavak stoje ravnopravno jedan
' ispod drugog. Nemaju: Storno je RADNJA, Oporavak je POSLEDICA -- spisak onoga
' sto je ostalo nedovrseno, uglavnom zato sto je neki storno stao na safe-stopu.
' Kod operatera se nakupilo 44 stavke a da nista nije reklo; sidebar je izgledao
' isto i kad je iza stavke nula i kad je 44.
'
' Brojac to razlikuje. Ali NE SME da uvede ljusku u poznavanje ekrana po imenu --
' ceo ugovor postoji da bi ljuska ostala neuka. Zato je Scr_Brojac OPCION clan
' ugovora: ekran koji nema sta da broji ga ne implementira i dobija nulu, a ljuska
' pita sve redom i ne zna ko je odgovorio.
Private Sub T_NavBrojac_SamoEkranKojiBroji()
    Dim n As Long

    ' NAJVAZNIJE PRVO: ekran koji broji vraca broj, i to kroz KASNO VEZIVANJE --
    ' isto kao svaki drugi clan ugovora. Da ljuska zove GetNedovrseno direktno,
    ' ova tvrdnja bi prolazila a ugovor bi bio probijen.
    n = modUiScreens.ScrBrojac("OPORAVAK")
    AssertEq (n >= 0), True, "ekran Oporavak odgovara na Scr_Brojac"
    AssertEq n, modScrOporavak.Scr_Brojac(), _
             "ljuska dobija BAS ono sto ekran broji, bez posrednika"

    ' Isti broj koji ekran prikazuje kao Nedovrseno -- da se meni i ekran ne mogu
    ' razici. Fixture ima bar dve ispravke na cekanju, pa nije nula.
    AssertEq (n >= 2), True, "brojac vidi ispravke na cekanju iz fixture-a"

    ' Ekran koji brojac NEMA ne sme da obori poziv niti da izmisli broj. Scr_Brojac
    ' je opcion: Application.Run na nepostojecu proceduru DIZE gresku, pa se bez
    ' gutanja te greske sidebar ne bi ni iscrtao.
    AssertEq modUiScreens.ScrBrojac("DOKUMENTI"), 0, _
             "ekran bez brojaca daje nulu, ne gresku"
    AssertEq modUiScreens.ScrBrojac("STORNO"), 0, _
             "ni Storno nema sta da broji -- on je radnja, ne zaostatak"

    ' Nepoznat kljuc takodje mora da prodje mirno: registar se menja, a sidebar
    ' ne sme da padne na stavku koja je u medjuvremenu izbacena.
    AssertEq modUiScreens.ScrBrojac("NE-POSTOJI"), 0, "nepoznat ekran daje nulu"

    ' I greska ne sme da PROCURI. Application.Run na nepostojecu proceduru je
    ' podigne; On Error Resume Next je proguta, ali je ostavi POSTAVLJENU -- pa bi
    ' prvi sledeci LogErr u ljusci zapisao tudju gresku, a prvi Err.Number <> 0
    ' skrenuo tok. Ista klasa nalaza kao test 66.
    Err.Clear
    n = modUiScreens.ScrBrojac("DOKUMENTI")
    AssertEq Err.Number, 0, "poziv ekrana bez brojaca ne ostavlja Err postavljen"
End Sub

' 81. Bazen ljuske je konacan, i prekoracenje se PRIJAVLJUJE.
'
' Segmenata, radnji, cipova i kolona ima tacno onoliko koliko je kontrola
' napravljeno. Ekran koji zatrazi vise nije gresio -- ali visak se gubio bez
' ijedne poruke, i to se desilo dvaput: jedanaesti cip se crtao a nije reagovao,
' a sesta radnja nad redom je tiho izbacila 'Nepotpune palete'.
'
' Test tvrdi oboje: da cuvar odseca i imenuje, i da nijedan DANASNJI ekran ne
' prekoracuje -- druga tvrdnja je ta koja ce pasti kad se doda osmi cip.
Private Sub T_BazenLjuske_ViseNegoStoStaje()
    Dim pre As Long, r As Variant, kljuc As String, i As Long
    Dim liste As Variant, spec As String

    ' 1) Sto staje -- prolazi netaknuto. Bez ove tvrdnje bi cuvar mogao da odseca
    ' i ono sto je u redu, pa bi ekrani tiho gubili poslednji cip.
    AssertEq modOtkupUI.BazenStaje(3, 5, "proba"), 3, _
             "sto staje u bazen prolazi neodseceno"
    AssertEq modOtkupUI.BazenStaje(5, 5, "proba"), 5, _
             "tacno pun bazen nije prekoracenje"

    ' 2) Visak se odseca NA velicinu bazena, ne na nulu i ne na trazeno.
    pre = modOtkupUI.BazenPrijavaBroj()
    AssertEq modOtkupUI.BazenStaje(9, 5, "probaX"), 5, _
             "visak se odseca na velicinu bazena"
    AssertEq modOtkupUI.BazenPrijavaBroj(), pre + 1, _
             "prekoracenje se prijavljuje"

    ' 3) I to JEDNOM. Ovo se zove pri svakom crtanju mreze; prijava po pozivu bi
    ' napunila log i sakrila ono sto se stvarno desava.
    AssertEq modOtkupUI.BazenStaje(9, 5, "probaX"), 5, _
             "ponovljeno prekoracenje i dalje odseca"
    AssertEq modOtkupUI.BazenPrijavaBroj(), pre + 1, _
             "isto prekoracenje se ne prijavljuje dvaput"

    ' 4) Nijedan danasnji ekran ne trazi vise nego sto staje. Ide se kroz REGISTAR,
    ' pa provera pokriva i ekrane koji tek dolaze -- ne treba je dopunjavati.
    r = modUiScreens.ScrRows()
    If Not IsArray(r) Then Exit Sub
    For i = 0 To UBound(r)
        kljuc = modUiScreens.ScrField(CStr(r(i)), modUiScreens.SCR_KLJUC)
        If Len(kljuc) > 0 Then
            liste = modUiScreens.ScrListe(kljuc)
            If IsArray(liste) Then
                AssertEq ((UBound(liste) + 1) <= modOtkupUI.MAX_SEG), True, _
                         "ekran " & kljuc & ": prekidac lista staje u bazen"
            End If
            spec = modUiScreens.ScrCipovi(kljuc)
            If Len(spec) > 0 Then
                AssertEq ((UBound(Split(spec, "|")) + 1) <= modOtkupUI.MAX_CHIP), _
                         True, "ekran " & kljuc & ": cipovi staju u bazen"
            End If
        End If
    Next i
End Sub

' ============================================================
' 80. Zona liste 'Nova prerada' ima SVA polja, i sva su vidljiva.
'
' Operater je prijavio zonu u kojoj se vide samo prvo polje (Bruto) i poslednje
' (Tip kese), bez bele podloge, naslova, NETO i dugmeta. Raspored je bio tacan --
' oba vidljiva polja su stajala BAS gde ih Scr_Layout salje -- pa kvar nije u
' merama nego u tome sto ostale kontrole ne postoje ili se ne pale.
'
' Zona se gradi nad obicnim Frame-om, pa se ceo taj put moze izmeriti bez .Show.
Private Sub T_ZonaPrerade_SvaPoljaVidljiva()
    Dim f As frmOtkupUI, z As Object, nm As Variant
    Dim nema As String, neupaljene As String, neugasene As String

    Set f = NewOtkupUIForm()
    Set z = f.Controls.Add("Forms.Frame.1", "zProba", True)
    z.width = 1200: z.Height = 300
    modScrPalete.Scr_Build z

    ' Nalazi se SKUPLJAJU, a tvrde tek posle Unload-a. Dok forma zivi, njena
    ' masinerija obrise Err izmedju Err.Raise i omotnice testa, pa bi pad stigao
    ' kao 'greska bez opisa' -- test bi padao tacno, a ne bi umeo da kaze zasto.
    ' Uz to spisak imena kaze KOJE kontrole fale, a ne samo da nesto fali.
    For Each nm In Array("preBg", "preCap", "preUlazL", "preUlazV", _
                         "preNetoL", "preNetoV", "preIzbor", _
                         "scrPreBruto", "scrPreTezPal", "scrPreGP", "scrPreNap", _
                         "scrPreKut", "scrPreTipKut", "scrPreKes", "scrPreTipKes")
        If Not KontrolaPostoji(z, CStr(nm)) Then nema = nema & " " & CStr(nm)
    Next nm

    modScrPalete.Scr_PalTestSet "NOVAPRERADA"
    modScrPalete.Scr_Layout z, 1200, 300
    For Each nm In Array("preBg", "preCap", "preUlazL", "preUlazV", _
                         "preNetoL", "preNetoV", "preIzbor", _
                         "scrPreBruto", "scrPreTezPal", "scrPreGP", "scrPreNap", _
                         "scrPreKut", "scrPreTipKut", "scrPreKes", "scrPreTipKes")
        If Not VidljivaKontrola(z, CStr(nm)) Then _
            neupaljene = neupaljene & " " & CStr(nm)
    Next nm

    modScrPalete.Scr_PalTestSet "PALETE"
    modScrPalete.Scr_Layout z, 1200, 300
    For Each nm In Array("preBg", "preCap", "scrPreBruto", "scrPreTipKes")
        If VidljivaKontrola(z, CStr(nm)) Then _
            neugasene = neugasene & " " & CStr(nm)
    Next nm

    modScrPalete.Scr_PalTestSet "PALETE"
    Unload f

    ' 1) Sve kontrole panela POSTOJE. Kontrola koje nema Scr_Layout tiho
    ' preskoci, pa operater vidi rupu na mestu polja.
    AssertEq nema, "", "panel za unos prerade nema nijednu kontrolu manje"

    ' 2) Na listi za unos su SVE upaljene -- ovo je bas ono sto je operater
    ' prijavio kao zonu u kojoj se vide samo prvo i poslednje polje.
    AssertEq neupaljene, "", "na listi za unos je upaljen ceo panel"

    ' 3) Na listama pregleda su ugasene -- inace bi polja unosa visila nad
    ' listom koja se samo cita.
    AssertEq neugasene, "", "u pregledu panel ostaje ugasen"
End Sub

' Postoji li kontrola pod tim imenom. Bez ovoga bi test morao da hvata gresku
' na svakom mestu gde pita.
Private Function KontrolaPostoji(z As Object, ByVal nm As String) As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = z.Controls(nm)
    KontrolaPostoji = Not (c Is Nothing)
    Err.Clear
End Function

' Vidljivost kontrole; kontrola koje NEMA nije vidljiva.
Private Function VidljivaKontrola(z As Object, ByVal nm As String) As Boolean
    On Error Resume Next
    VidljivaKontrola = z.Controls(nm).Visible
    Err.Clear
End Function

' 79. Cipovi pripadaju EKRANU, ne ljusci: ugovor, bazen i stvarno suzavanje.
'
' Ljuska je do sada znala jedan ekran po imenu -- 'ako je lista OTPREMNICE,
' pokazi chipSve i chipOtvorene'. Sada svaki ekran prijavi svoje cipove, a
' ljuska pozajmljuje slotove svog bazena. Test meri obe strane: opis koji ekran
' daje i pravilo po kom se lista suzava.
Private Sub T_CipoviEkrana_UgovorIFilter()
    Dim spec As String, e As Variant, p As Variant, n As Long
    Dim d As Variant, uk As Long, otv As Long, zat As Long, i As Long

    ' 1) UGOVOR: pet cipova nad listom paleta, svaki kljuc:KATALOG:sirina, i
    ' svaki natpis postoji u katalogu -- cip bez natpisa je prazno dugme.
    modScrPalete.Scr_PalTestSet "PALETE"
    spec = modScrPalete.Scr_Cipovi()
    AssertEq (Len(spec) > 0), True, "lista paleta prijavljuje svoje cipove"
    For Each e In Split(spec, "|")
        p = Split(CStr(e), ":")
        AssertEq (UBound(p) = 2), True, "cip je oblika kljuc:KATALOG:sirina"
        ' Katalog na nepostojeci kljuc vraca "[KLJUC]", pa bi provera duzine
        ' uvek prolazila i ne bi merila nista. Meri se da natpis NIJE ta oznaka.
        AssertEq (Left$(Poruka(CStr(p(1))), 1) <> "["), True, _
                 "natpis cipa " & CStr(p(0)) & " postoji u katalogu"
        AssertEq (val(p(2)) > 0), True, "cip ima sirinu"
        n = n + 1
    Next e
    AssertEq n, 5, "lista paleta ima pet cipova"
    ' Bazen ljuske je konacan: visak bi se izgubio bez ijedne poruke.
    AssertEq (n <= modOtkupUI.MAX_CHIP), True, _
             "ekran ne trazi vise cipova nego sto bazen ljuske ima"
    AssertEq Split(CStr(Split(spec, "|")(0)), ":")(0), "sve", _
             "prvi cip je SVE -- na njega se pada kad filter ne pripada listi"

    ' 2) Ostale liste istog ekrana nemaju sta da suze.
    modScrPalete.Scr_PalTestSet "STAVKE"
    AssertEq modScrPalete.Scr_Cipovi(), "", "lista stavki nema cipove"
    modScrPalete.Scr_PalTestSet "PRERADE"
    AssertEq modScrPalete.Scr_Cipovi(), "", "lista prerada nema cipove"

    ' 3) PRAVILO cipa, bez mreze. Prazan i nepoznat kljuc puste sve -- ekran koji
    ' dobije filter koji ne poznaje pokazuje punu listu, ne praznu.
    AssertEq modScrPalete.PalCipProlaz("otvorene", "Otvorena", "2026", ""), True, _
             "otvorena paleta prolazi kroz cip Otvorene"
    AssertEq modScrPalete.PalCipProlaz("otvorene", "ZATVORENA", "2026", ""), False, _
             "zatvorena paleta ne prolazi kroz cip Otvorene"
    AssertEq modScrPalete.PalCipProlaz("zatvorene", "Zatvorena", "2026", ""), True, _
             "cip Zatvorene ne gleda velika i mala slova"
    AssertEq modScrPalete.PalCipProlaz("preradjene", "Otvorena", "2026", "DA"), True, _
             "preradjena paleta prolazi kroz cip Preradjene"
    AssertEq modScrPalete.PalCipProlaz("preradjene", "Otvorena", "2026", ""), False, _
             "nepreradjena paleta ne prolazi kroz cip Preradjene"
    AssertEq modScrPalete.PalCipProlaz("godina", "Otvorena", CStr(Year(Date)), ""), _
             True, "paleta ove godine prolazi kroz cip Ova godina"
    AssertEq modScrPalete.PalCipProlaz("godina", "Otvorena", "1999", ""), False, _
             "paleta iz ranije godine ne prolazi kroz cip Ova godina"
    AssertEq modScrPalete.PalCipProlaz("nepoznat", "Otvorena", "1999", ""), True, _
             "nepoznat filter pusta sve, da lista ne ostane prazna"

    ' 4) Cip STVARNO suzava mrezu, i to bez gubitka: svaka paleta je ili otvorena
    ' ili zatvorena, pa dva cipa moraju da daju tacno ono sto daje 'sve'.
    modScrPalete.Scr_PalTestSet "PALETE"
    d = modScrPalete.Scr_Rows("sve", "")
    uk = CLng(d(2))
    AssertEq (uk > 0), True, "preduslov: fixture ima paleta"
    d = modScrPalete.Scr_Rows("otvorene", "")
    otv = CLng(d(2))
    d = modScrPalete.Scr_Rows("zatvorene", "")
    zat = CLng(d(2))
    AssertEq otv + zat, uk, "Otvorene i Zatvorene zajedno daju sve palete"
    AssertEq (otv < uk Or zat < uk), True, "cip stvarno suzava, ne vraca sve"

    ' 5) Unosni ekran: lista otpremnica nosi svoje cipove, lista dokumenata NE --
    ' njeni cipovi zavise od rezima (zbirna, faktura) pa ostaju ljuskini.
    ' Lista otpremnica postoji samo u rezimu OTKUP, pa se sam ugovor ne moze
    ' dovesti u to stanje bez forme -- meri se pravilo, koje je zato izdvojeno.
    spec = modScrDokumenti.CipoviZaListu("OTPREMNICE")
    AssertEq (InStr(spec, "otvorene:") > 0), True, _
             "lista otpremnica prijavljuje svoj cip Neraspodeljene"
    AssertEq modScrDokumenti.CipoviZaListu("SVI"), "", _
             "lista dokumenata prepusta cipove ljusci -- oni zavise od rezima"
    ' i ugovor stvarno ide kroz to pravilo, a ne pored njega
    AssertEq modScrDokumenti.Scr_Cipovi(), _
             modScrDokumenti.CipoviZaListu(modScrDokumenti.Scr_Lista()), _
             "Scr_Cipovi vraca bas ono sto pravilo kaze za aktivnu listu"

    modScrPalete.Scr_PalTestSet "PALETE"
End Sub

' 78. Dvoklik na paletu otvara njene stavke; obican klik i dalje samo BIRA.
'
' Zasto oba smera u istom testu: da klik prebacuje listu, radnje nad redom
' (zatvori paletu, storniraj, stampaj) postale bi nedostupne -- operater ne bi
' stigao da ih pritisne. Tvrdnja nije samo 'dvoklik radi' nego 'dvoklik radi, a
' klik je ostao netaknut'.
Private Sub T_PaletaDvoklik_OtvaraStavke()
    Dim broj As String, opis As String, d As Variant

    ' Mreza se puni pravim podacima ekrana, bez forme -- klik cita bas nju.
    modScrPalete.Scr_PalTestSet "PALETE"
    ' Zona koje NEMA ne sme da izgleda kao pad ekrana. Ekran u Scr_Rows dopunjava
    ' svoju zonu; kad zone nema, On Error Resume Next gresku preskoci ali je ostavi
    ' postavljenu, pa je ScrGridData procita kao pad i isprazni mrezu. Ovo je merenje
    ' bas tog puta -- harness nema zonu, kao ni ekran koji je pukao u gradnji.
    d = modUiScreens.ScrGridData("PALETE", "sve", "")
    AssertEq modUiScreens.ScrLastErr, "", _
             "procitana lista se ne prijavljuje kao pad ekrana"
    AssertEq IsArray(d), True, "ljuska je dobila redove"
    modOtkupUI.GridTestLoad "PALETE"
    broj = Trim$(CStr(modOtkupUI.GridCell(1, 1)))
    AssertEq (Len(broj) > 0), True, "preduslov: fixture ima paletu u prvom redu"

    ' 1) OBICAN KLIK bira i nista vise: lista ostaje, mreza se ne cita ponovo.
    AssertEq modUiScreens.ScrEvent("PALETE", "row:1", "Click"), False, _
             "izbor reda ne trazi ponovno citanje mreze"
    AssertEq modScrPalete.Scr_Lista(), "PALETE", _
             "obican klik BIRA, ne otvara -- inace radnje nad redom postaju nedostupne"

    ' 2) DVOKLIK otvara stavke te palete i trazi ponovno citanje.
    AssertEq modUiScreens.ScrEvent("PALETE", "dbl:1", "Click"), True, _
             "dvoklik menja listu, pa mreza mora da se procita ponovo"
    AssertEq modScrPalete.Scr_Lista(), "STAVKE", _
             "dvoklik na paletu otvara njene stavke"

    ' 3) Zona i naslov i dalje pokazuju KOJA je paleta otvorena -- bez toga se sa
    ' liste stavki ne bi znalo cije su.
    opis = modScrPalete.Scr_NaslovDopuna()
    AssertEq opis, broj, "naslov liste stavki nosi broj otvorene palete"

    ' 4) Stavke se otvaraju SAMO dvoklikom -- ne i sestom radnjom nad redom.
    ' Bazen dugmadi je MAX_ACT i lista paleta ga vec puni do vrha; sesta radnja bi
    ' tiho ostala bez dugmeta, jer RefreshRowActions radi Exit For. Tako je jednom
    ' vec izbacila 'Nepotpune palete'.
    modScrPalete.Scr_PalTestSet "PALETE"
    Dim r As Variant
    r = Split(modScrPalete.Scr_Radnje(), "|")
    AssertEq ((UBound(r) + 1) <= modOtkupUI.MAX_ACT), True, _
             "lista ne trazi vise radnji nego sto ljuska ima dugmadi"

    modScrPalete.Scr_PalTestSet "PALETE"
    modOtkupUI.GridTestLoad ""
End Sub

' 77. Nova prerada: cetvrta lista, izbor po identitetu, neto kao racun
' ============================================================
' Faza C, stavka 10. Legacy je unos prerade radio panelom sa sedam polja i
' multiselektom liste; ovde su polja u zoni, a izbor ide stikliranjem u mrezi.
'
' Sto se moze izmeriti bez forme: ugovor liste, kolone, izbor po PaletaID-u i
' sam racun neta. Sto NE moze: da se polje vidi i da kucanje osvezi brojku --
' zona se crta nad formom koju harness gradi bez .Show. To ostaje na smoke-u.
Private Sub T_NovaPrerada_IzborINeto()
    Dim liste As Variant, d As Variant, r As Variant, i As Long, n As Long
    Dim ident As String, kljucevi As String

    ' 1) Ekran prijavljuje cetvrtu listu, i to POSLEDNJU -- prve tri su pregledi,
    ' ova je radna, pa stoji na kraju prekidaca.
    liste = modScrPalete.Scr_Liste()
    AssertEq (UBound(liste) + 1), 4, "ekran Palete ima cetiri liste"
    For i = 0 To UBound(liste)
        kljucevi = kljucevi & "|" & Split(CStr(liste(i)), "|")(0)
    Next i
    AssertEq kljucevi, "|PALETE|STAVKE|PRERADE|NOVAPRERADA", _
             "redosled i kljucevi lista"

    ' 2) Kolone: kvacica napred, identitet pozadi i NEVIDLJIV. Radnja gadja
    ' PaletaID, ne broj palete -- spisak zavrsava u SavePrerada_TX.
    modScrPalete.Scr_PalTestSet "NOVAPRERADA"
    d = modScrPalete.Scr_Rows("sve", "")
    AssertEq Split(CStr(d(0)(0)), "|")(0), "OTKUI_HDP_OZN", _
             "prva kolona je izbor"
    AssertEq (UBound(d(0)) + 1), modScrPalete.PAL_NOVA_COL_ID, _
             "opis kolona se zavrsava BAS na koloni koju radnja cita"
    AssertEq Split(CStr(d(0)(modScrPalete.PAL_NOVA_COL_ID - 1)), "|")(0), "OTKUI_HD_IDENT", _
             "poslednja kolona je identitet palete"
    AssertEq modScrDokumenti.ColF(CStr(d(0)(modScrPalete.PAL_NOVA_COL_ID - 1)), 4), "4", _
             "kolona identiteta je prioriteta 4 -- nikad vidljiva"

    ' 3) PODRAZUMEVANO NIJEDNA nije oznacena.
    n = CLng(d(2))
    AssertEq (n > 0), True, "preduslov: fixture ima paleta"
    r = d(1)
    For i = 1 To n
        AssertEq CStr(r(i, 1)), "", "red " & i & " nije oznacen bez izricitog izbora"
    Next i
    AssertEq modScrPalete.PalOznacenihBroj(), 0, "brojac oznacenih je nula"

    ' 4) Oznaka gadja BAS tu paletu, po identitetu iz nevidljive kolone.
    ident = Trim$(CStr(r(1, modScrPalete.PAL_NOVA_COL_ID)))
    AssertEq (Len(ident) > 0), True, "red nosi PaletaID u nevidljivoj koloni"
    modScrPalete.Scr_PreTestSet ident
    d = modScrPalete.Scr_Rows("sve", "")
    r = d(1)
    AssertEq modScrPalete.PalOznacenihBroj(), 1, "oznacena je tacno jedna paleta"
    AssertEq CStr(r(1, 1)), ChrW(10003), "oznaceni red nosi kvacicu"
    ' Neto ULAZ je zbir neto kilaze izabranih paleta -- operater po njemu vidi sa
    ' koliko sveze robe ulazi u preradu, pre nego sto unese izlaz.
    AssertEq modScrPalete.NetoUlazIzabranih(), CDbl(r(1, 10)), _
             "neto ulaz je zbir neto izabranih paleta"
    For i = 2 To CLng(d(2))
        AssertEq CStr(r(i, 1)), "", "ostale palete ostaju neoznacene"
    Next i

    ' 5) NETO je racun, ne unos: bruto minus tezina palete minus ambalaza.
    ' Bez ambalaze (nepoznat tip -> tezina 0) ostaje cista razlika.
    AssertEq modScrPalete.NetoIzracun(100, 20, 0, "", 0, ""), 80#, _
             "neto je bruto minus tezina palete"

    ' I donja granica: negativan neto nije podatak nego znak da unos jos nije
    ' potpun. Nula je iskrenija od minusa.
    AssertEq modScrPalete.NetoIzracun(10, 40, 0, "", 0, ""), 0#, _
             "neto se ne spusta ispod nule"

    modScrPalete.Scr_PalTestSet "PALETE"
End Sub


' ============================================================
' 82. Ugovor ekrana Agrohemija: cetiri liste, radnja samo nad korpom
' ============================================================
' Isti oblik kao T_Storno_UgovorIRadnje. Postoji zato sto ekran koji nije u
' registru ili ne odgovara na ugovor NE PADA -- sidebar ga samo prikaze
' prigusenog, pa agrohemija nestane iz aplikacije bez ijedne greske.
Private Sub T_Agro_UgovorEkrana()
    Dim liste As Variant, i As Long, kljucevi As String, d As Variant
    Dim kljuc As String, spec As String

    AssertEq (Len(modUiScreens.ScrRowByKey("AGRO")) > 0), True, _
             "AGRO postoji u registru ekrana"
    AssertEq modUiScreens.ScrPostoji("AGRO"), True, _
             "modul ekrana Agrohemija odgovara na Scr_Meta (kasno vezivanje radi)"
    AssertEq (InStr(modUiScreens.ScrMeta("AGRO"), "kljuc=AGRO") > 0), True, _
             "Scr_Meta prijavljuje svoj kljuc"
    AssertEq modUiScreens.ScrField(modUiScreens.ScrRowByKey("AGRO"), SCR_OBLAST), _
             OBL_AGROHEMIJA, "ekran trazi pravo na oblast Agrohemija"

    liste = modScrAgro.Scr_Liste()
    ' NAJVAZNIJE PRVO: ljuska mora da nacrta SVE liste koje ekran prijavi.
    ' LayoutGrid crta prvih MAX_SEG i stane, bez greske i bez traga.
    AssertEq (UBound(liste) + 1 <= modOtkupUI.MaxPrekidaca()), True, _
             "ljuska crta sve liste ekrana -- nijedna se ne odseca tiho"
    For i = 0 To UBound(liste)
        kljucevi = kljucevi & "|" & Split(CStr(liste(i)), "|")(0)
    Next i
    AssertEq kljucevi, "|KORPA|STANJE|PROMET|DUGOVI", _
             "redosled i kljucevi lista -- korpa je prva"

    ' Radnja nad redom postoji SAMO u korpi: stanje, promet i dugovi su
    ' pregledi, a storno magacin stavke je posao ekrana Storno.
    modScrAgro.Scr_ListaTestSet "KORPA"
    AssertEq (Len(modScrAgro.Scr_Radnje()) > 0), True, _
             "korpa ima radnju nad redom"
    modScrAgro.Scr_ListaTestSet "STANJE"
    AssertEq modScrAgro.Scr_Radnje(), "", _
             "stanje je pregled -- nema radnje nad redom"
    modScrAgro.Scr_ListaTestSet "PROMET"
    AssertEq modScrAgro.Scr_Radnje(), "", _
             "promet je pregled -- nema radnje nad redom"
    modScrAgro.Scr_ListaTestSet "DUGOVI"
    AssertEq modScrAgro.Scr_Radnje(), "", _
             "dugovi su pregled -- nema radnje nad redom"

    ' Svaka lista mora da vrati ISPRAVAN niz. Lista koja pukne se u ljusci
    ' pretvara u Empty, LoadGridFromScreen na ne-niz radi Exit Sub -- pa mreza
    ' ostane na prethodnoj listi i prekidac izgleda kao da ne radi.
    For i = 0 To UBound(liste)
        kljuc = Split(CStr(liste(i)), "|")(0)
        modScrAgro.Scr_ListaTestSet kljuc
        d = modScrAgro.Scr_Rows("sve", "")
        AssertEq IsArray(d), True, "lista " & kljuc & " vraca niz"
        AssertEq (UBound(d) >= 4), True, _
                 "lista " & kljuc & " vraca pun oblik (kolone, redovi, n, kg, vrednost)"
        AssertEq IsArray(d(0)), True, "lista " & kljuc & " prijavljuje svoje kolone"
    Next i

    ' Dve korpe zive istovremeno -- prekidac rezima ne prazni ni jednu.
    modScrAgro.Scr_RezimTestSet "ULAZ"
    AssertEq modScrAgro.Scr_Rezim(), "ULAZ", "prekidac rezima menja rezim"
    modScrAgro.Scr_RezimTestSet "IZLAZ"
    AssertEq modScrAgro.Scr_Rezim(), "IZLAZ", "povratak na izdavanje"

    ' GRANICE BAZENA LJUSKE. Visak se ne prijavljuje kao greska nego se TIHO
    ' odseca: LayoutGrid nacrta prvih MAX_SEG i stane, RefreshRowActions prvih
    ' MAX_ACT, bazen cipova prvih MAX_CHIP, SetGridColsArr odseca kolone na
    ' MAX_COLS. Operater tada vidi ekran kome fali dugme, bez ijedne poruke.
    modScrAgro.Scr_ListaTestSet "KORPA"
    AssertEq (UBound(Split(modScrAgro.Scr_Radnje(), "|")) + 1 <= modOtkupUI.MAX_ACT), _
             True, "korpa ne trazi vise radnji nego sto ljuska ima dugmadi"
    For i = 0 To UBound(liste)
        kljuc = Split(CStr(liste(i)), "|")(0)
        modScrAgro.Scr_ListaTestSet kljuc
        spec = modScrAgro.Scr_Cipovi()
        If Len(spec) > 0 Then
            AssertEq (UBound(Split(spec, "|")) + 1 <= modOtkupUI.MAX_CHIP), True, _
                     "lista " & kljuc & " ne trazi vise cipova nego sto bazen ima"
        End If
        d = modScrAgro.Scr_Rows("sve", "")
        AssertEq (UBound(d(0)) + 1 <= modOtkupUI.MAX_COLS), True, _
                 "lista " & kljuc & " ne trazi vise kolona nego sto mreza pravi"
    Next i

    modScrAgro.Scr_ListaTestSet "KORPA"
End Sub

' ============================================================
' 83. Kapija stanja broji i ono sto je VEC u korpi
' ============================================================
' Dve tvrdnje, obe iz legacy forme:
'   1. dodavanje u korpu sabira sa onim sto je u korpi (btnDodajIzlaz), pa se
'      ista roba ne moze dodati dva puta preko stanja;
'   2. pred upis se stanje proverava JOS JEDNOM, agregirano po artiklu
'      (ValidateKorpaIzlazStanje) -- jer se stanje izmedju dodavanja i upisa
'      moglo promeniti (drugi operater, sync).
' Bez druge kapije bi upis krenuo pa pao na pola petlje i vratio se rollback-om,
' a operater bi dobio 4301 umesto recenice.
Private Sub T_Agro_KapijaStanjaBrojiKorpu()
    Dim korpa As Collection, fokus As String, mapa As Object
    Dim magID As String

    ' PREDUSLOV: stanje je ono iz fixture-a. Bez ovoga bi test merio ostatak
    ' ranijeg testa umesto posledice ove radnje.
    Set mapa = modAgroUnos.AgroStanjeMapa()
    AssertEq (mapa.Exists(FX_ARTIKAL)), True, "artikal iz fixture-a ima stanje"
    AssertEq CDbl(mapa(FX_ARTIKAL)), FX_ART_STANJE, _
             "PREDUSLOV: stanje artikla je " & FX_ART_STANJE

    Set korpa = modAgroUnos.NovaAgroKorpa()

    ' Pola pakovanja se ne izdaje -- kolicina se kuca u PAKOVANJIMA.
    AssertEq (Len(modAgroUnos.AgroDodajIzlaz(korpa, FX_ARTIKAL, 1.5, FX_PARCELA, fokus)) > 0), _
             True, "pola pakovanja se ne izdaje"
    AssertEq korpa.count, 0, "odbijena stavka ne ulazi u korpu"

    ' Dva pa jos jedno pakovanje = 15 l, tacno stanje.
    AssertEq modAgroUnos.AgroDodajIzlaz(korpa, FX_ARTIKAL, 2, FX_PARCELA, fokus), "", _
             "dva pakovanja staju u stanje"
    AssertEq modAgroUnos.AgroDodajIzlaz(korpa, FX_ARTIKAL, 1, FX_PARCELA, fokus), "", _
             "i trece pakovanje staje -- granica se ne odbija"
    AssertEq korpa.count, 2, "korpa ima dve stavke"

    ' NAJVAZNIJE: cetvrto pada BAS zbog onoga sto je vec u korpi. Kapija koja
    ' gleda samo stanje bi ga propustila (5 < 15) i upis bi pao tek u petlji.
    AssertEq (Len(modAgroUnos.AgroDodajIzlaz(korpa, FX_ARTIKAL, 1, FX_PARCELA, fokus)) > 0), _
             True, "kapija stanja broji i ono sto je vec u korpi"
    AssertEq korpa.count, 2, "odbijena stavka ne ulazi u korpu"

    ' Artikal koji nema nijedan magacin red: stanje 0, pa izdavanja nema.
    AssertEq (Len(modAgroUnos.AgroDodajIzlaz(korpa, FX_ARTIKAL_BEZ_STANJA, 1, FX_PARCELA, fokus)) > 0), _
             True, "artikal bez stanja se ne izdaje"
    ' Artikal bez popunjenog Pakovanja: invarijanta, ne stanje.
    AssertEq (Len(modAgroUnos.AgroDodajIzlaz(korpa, FX_ARTIKAL_BEZ_PAK, 1, FX_PARCELA, fokus)) > 0), _
             True, "artikal bez Pakovanja se ne izdaje"
    AssertEq korpa.count, 2, "nijedna odbijena stavka nije usla u korpu"

    ' Agregirana kapija: korpa tacno na stanju PROLAZI.
    AssertEq modAgroUnos.AgroProveriKorpuIzlaz(korpa), "", _
             "korpa tacno na stanju prolazi kapiju pre upisa"

    ' A sada se stanje promeni IZA ledja korpe -- tacno zbog toga druga kapija
    ' i postoji. Ista korpa vise ne sme da prodje.
    ' Parcela se PROSLEDJUJE: PRACENJE_PARCELA je u fixture-u ukljuceno, pa
    ' bi prazna parcela ovde podigla 4215 i test bi pao na svom cistacu
    ' umesto na tvrdnji koju meri.
    magID = SaveMagacinCore(Date, FX_ARTIKAL, MAG_IZLAZ, 10, FX_KOOPERANT, FX_PARCELA, _
                            "AGRO-TEST-TX")
    AssertEq (Len(Trim$(magID)) > 0), True, "kontrolni izlaz je proknjizen"
    AssertEq (Len(modAgroUnos.AgroProveriKorpuIzlaz(korpa)) > 0), True, _
             "korpa vise ne staje u stanje -- kapija pre upisa to hvata"

    ' Ciscenje se i PROVERAVA. Nevereno vracanje je isto sto i nikakvo: test
    ' dodat ispod ovog nasledio bi tiho izmenjen fixture i pao bi po tudjem imenu.
    StornirajMagacinRed magID
    Set mapa = modAgroUnos.AgroStanjeMapa()
    AssertEq CDbl(mapa(FX_ARTIKAL)), FX_ART_STANJE, _
             "fixture je vracen: stanje je opet " & FX_ART_STANJE
End Sub

' ============================================================
' 84. Smart doza se zaokruzuje NAGORE, na cela pakovanja
' ============================================================
' Doza je racun po hektaru, ali se roba izdaje u pakovanjima -- pola pakovanja
' ne postoji. Fixture je namesten tako da se razlika vidi: doza 2 l/ha na
' 1.5 ha = 3 l, a pakovanje je 5 l. Zaokruzenje nanize dalo bi 0 pakovanja,
' matematicko zaokruzenje takodje 1 -- ali na 3.75 ha (7.5 l) matematicko daje
' 2 i nanize 1, dok nagore mora dati 2.
Private Sub T_Agro_SmartDozaZaokruzujeNagore()
    Dim pre As Object, info As Object

    Set pre = modAgroUnos.AgroPreporukaInfo(FX_ARTIKAL, 1.5)
    AssertEq CStr(pre("greska")), "", "artikal iz fixture-a nema smetnju"
    AssertEq CDbl(pre("dozaKg")), 3#, "doza za 1.5 ha je 3 l"
    AssertEq CDbl(pre("pakovanje")), FX_ART_PAKOVANJE, "pakovanje iz sifarnika"
    AssertEq CLng(pre("brojPak")), 1&, "3 l trazi JEDNO pakovanje od 5 l"
    AssertEq CDbl(pre("izdajKol")), 5#, "izdaje se celo pakovanje, ne 3 l"

    ' 3.75 ha -> 7.5 l -> dva pakovanja (nagore), ne jedno.
    Set pre = modAgroUnos.AgroPreporukaInfo(FX_ARTIKAL, 3.75)
    AssertEq CDbl(pre("dozaKg")), 7.5, "doza za 3.75 ha je 7.5 l"
    AssertEq CLng(pre("brojPak")), 2&, "7.5 l trazi DVA pakovanja -- nagore"
    AssertEq CDbl(pre("izdajKol")), 10#, "izdaju se dva cela pakovanja"

    ' Bez izabrane parcele nema ni preporuke -- ne sme da izmisli jedno pakovanje.
    Set pre = modAgroUnos.AgroPreporukaInfo(FX_ARTIKAL, 0)
    AssertEq CLng(pre("brojPak")), 0&, "bez hektara nema preporuke"

    ' Invarijanta nad Pakovanjem je kapija, i prijavljuje se kao smetnja a ne
    ' kao nula: nula bi izgledala kao "nema sta da se izda".
    Set info = modAgroUnos.AgroArtikalInfo(FX_ARTIKAL_BEZ_PAK)
    AssertEq (Len(CStr(info("greska"))) > 0), True, _
             "artikal bez Pakovanja prijavljuje smetnju"
    Set pre = modAgroUnos.AgroPreporukaInfo(FX_ARTIKAL_BEZ_PAK, 1.5)
    AssertEq (Len(CStr(pre("greska"))) > 0), True, _
             "preporuka nad artiklom bez Pakovanja prijavljuje smetnju"
    AssertEq CLng(pre("brojPak")), 0&, "i ne predlaze nijedno pakovanje"
End Sub

' ============================================================
' 85. Zona agrohemije: polja postoje i prate rezim
' ============================================================
' Isti oblik kao T_ZonaPrerade_SvaPoljaVidljiva, i iz istog razloga: kontrolu
' koje NEMA Scr_Layout tiho preskoci (On Error Resume Next), pa operater vidi
' rupu na mestu polja, a log ne kaze nista.
'
' Ovde je jos jedna stvar pod merenjem: prekidac rezima. Izdavanje i prijem
' dele polja (artikal, kolicina, broj dokumenta) a razlikuju se u ostalima --
' ako se vidljivost i raspored raziju, polja se preklope jedno preko drugog.
' Zato oba rezima imaju i svoj spisak koji MORA da bude ugasen.
Private Sub T_ZonaAgro_PoljaPostojeIPrateRezim()
    Dim f As frmOtkupUI, z As Object, nm As Variant
    Dim nema As String, izlNema As String, izlVisak As String
    Dim ulzNema As String, ulzVisak As String

    Set f = NewOtkupUIForm()
    Set z = f.Controls.Add("Forms.Frame.1", "zProbaAg", True)
    z.width = 1200: z.Height = 300
    modScrAgro.Scr_Build z

    ' Nalazi se SKUPLJAJU, a tvrde tek posle Unload-a: dok forma zivi, njena
    ' masinerija obrise Err izmedju Err.Raise i omotnice testa, pa bi pad stigao
    ' kao "greska bez opisa".
    For Each nm In Array("agBg", "agCap", "agParTxt", "agHint", "agVred", "agLnB", _
                         "agKL0", "agKV0", "agKL3", "agKV3", _
                         "scrAgSegI", "scrAgSegU", "scrAgParAdd", "scrAgParClr", _
                         "scrAgDodaj", "scrAgZavrsi", "scrAgPocDug", "scrAgOcisti", _
                         "scrAgKoop", "scrAgArt", "scrAgPar", "scrAgDob", _
                         "scrAgKol", "scrAgCena", "scrAgDok")
        If Not KontrolaPostoji(z, CStr(nm)) Then nema = nema & " " & CStr(nm)
    Next nm

    ' Kombo u zoni MORA biti polje (okvir nm + kontrola nmT): panel za izbor
    ' (modOtkupUI.FindCombo) trazi bas taj oblik. Gola kontrola bi imala
    ' strelicu koja "ne radi" i listu koja se ne otvara.
    For Each nm In Array("scrAgKoop", "scrAgArt", "scrAgPar", _
                         "scrAgKol", "scrAgCena", "scrAgDok", "scrAgDob")
        If KontrolaPostoji(z, CStr(nm)) Then
            If Not KontrolaPostoji(z.Controls(CStr(nm)), CStr(nm) & "T") Then _
                nema = nema & " " & CStr(nm) & "T"
        End If
    Next nm

    ' IZDAVANJE: kooperant i parcele postoje, dobavljac i cena ne.
    modScrAgro.Scr_RezimTestSet "IZLAZ"
    modScrAgro.Scr_Layout z, 1200, 300
    For Each nm In Array("scrAgKoop", "scrAgArt", "scrAgKol", "scrAgDok", "scrAgPocDug")
        If Not VidljivaKontrola(z, CStr(nm)) Then izlNema = izlNema & " " & CStr(nm)
    Next nm
    For Each nm In Array("scrAgDob", "scrAgCena")
        If VidljivaKontrola(z, CStr(nm)) Then izlVisak = izlVisak & " " & CStr(nm)
    Next nm

    ' PRIJEM: obrnuto. Kooperant i parcele nemaju sta da traze u prijemu --
    ' roba ulazi od dobavljaca, ne izlazi kooperantu.
    modScrAgro.Scr_RezimTestSet "ULAZ"
    modScrAgro.Scr_Layout z, 1200, 300
    For Each nm In Array("scrAgArt", "scrAgDob", "scrAgCena", "scrAgKol", "scrAgDok")
        If Not VidljivaKontrola(z, CStr(nm)) Then ulzNema = ulzNema & " " & CStr(nm)
    Next nm
    For Each nm In Array("scrAgKoop", "scrAgPar", "scrAgPocDug")
        If VidljivaKontrola(z, CStr(nm)) Then ulzVisak = ulzVisak & " " & CStr(nm)
    Next nm

    modScrAgro.Scr_RezimTestSet "IZLAZ"
    Unload f

    AssertEq nema, "", "zona agrohemije nema nijednu kontrolu manje"
    AssertEq izlNema, "", "u izdavanju su upaljena sva polja izdavanja"
    AssertEq izlVisak, "", "u izdavanju su ugasena polja prijema"
    AssertEq ulzNema, "", "u prijemu su upaljena sva polja prijema"
    AssertEq ulzVisak, "", "u prijemu su ugasena polja izdavanja"
End Sub

' ============================================================
' 86. Cipovi agrohemije: ugovor i stvarno suzavanje
' ============================================================
' Dve stvari, i obe padaju tiho ako se pokvare. Ugovor: cip bez natpisa u
' katalogu je prazno dugme, a cip bez sirine se ne vidi. Pravilo: cip koji ne
' suzava nista izgleda kao da radi -- lista je ista i pre i posle klika.
'
' Pravila se mere BEZ mreze, kao PalCipProlaz na Paletama: mreza bi uvela
' sortiranje, stranice i pretragu u tvrdnju koja je o jednom uslovu.
Private Sub T_Agro_CipoviSuzavajuListu()
    Dim spec As String, e As Variant, p As Variant, n As Long

    ' 1) UGOVOR nad svakom listom koja cipove ima.
    modScrAgro.Scr_ListaTestSet "STANJE"
    spec = modScrAgro.Scr_Cipovi()
    AssertEq (Len(spec) > 0), True, "lista stanja prijavljuje svoje cipove"
    For Each e In Split(spec, "|")
        p = Split(CStr(e), ":")
        AssertEq (UBound(p) = 2), True, "cip je oblika kljuc:KATALOG:sirina"
        ' Katalog na nepostojeci kljuc vraca "[KLJUC]", pa bi provera duzine
        ' uvek prolazila i ne bi merila nista.
        AssertEq (Left$(Poruka(CStr(p(1))), 1) <> "["), True, _
                 "natpis cipa " & CStr(p(0)) & " postoji u katalogu"
        AssertEq (val(p(2)) > 0), True, "cip " & CStr(p(0)) & " ima sirinu"
        n = n + 1
    Next e
    AssertEq n, 3, "stanje ima tri cipa"
    AssertEq Split(CStr(Split(spec, "|")(0)), ":")(0), "sve", _
             "prvi cip je SVE -- na njega ljuska pada kad filter ne pripada listi"

    ' Korpa je nekoliko upravo unetih redova -- tu se ne trazi nego se gleda.
    modScrAgro.Scr_ListaTestSet "KORPA"
    AssertEq modScrAgro.Scr_Cipovi(), "", "korpa nema cipove"

    ' 2) PRAVILA. Prazan i nepoznat kljuc PUSTAJU sve: ekran koji dobije filter
    ' koji ne poznaje pokazuje punu listu, ne praznu.
    AssertEq modScrAgro.AgCipStanje("", 0), True, "prazan filter pusta sve"
    AssertEq modScrAgro.AgCipStanje("nepoznato", 0), True, "nepoznat filter pusta sve"
    AssertEq modScrAgro.AgCipStanje("ima", 0.5), True, "pola jedinice JESTE na stanju"
    AssertEq modScrAgro.AgCipStanje("ima", 0), False, "nula nije na stanju"
    AssertEq modScrAgro.AgCipStanje("nema", 0), True, "nula je bez zaliha"
    ' Negativno stanje je greska u knjizenju, ali se MORA videti -- sakriveno bi
    ' ostalo neispravljeno.
    AssertEq modScrAgro.AgCipStanje("nema", -3), True, "negativno stanje je bez zaliha"
    AssertEq modScrAgro.AgCipStanje("ima", -3), False, "negativno stanje nije na stanju"

    AssertEq modScrAgro.AgCipPromet("ulaz", MAG_ULAZ, 0), True, "ulaz prolazi kroz cip Ulazi"
    AssertEq modScrAgro.AgCipPromet("ulaz", MAG_IZLAZ, 0), False, "izlaz ne prolazi kroz Ulazi"
    AssertEq modScrAgro.AgCipPromet("izlaz", MAG_IZLAZ, 0), True, "izlaz prolazi kroz cip Izlazi"
    AssertEq modScrAgro.AgCipPromet("ulaz", UCase$(MAG_ULAZ), 0), True, _
             "tip se ne poredi po velicini slova"
    AssertEq modScrAgro.AgCipPromet("godina", MAG_ULAZ, CDbl(CDate(Date))), True, _
             "danasnji red je iz ove godine"
    AssertEq modScrAgro.AgCipPromet("godina", MAG_ULAZ, CDbl(CDate(DateSerial(Year(Date) - 1, 6, 1)))), _
             False, "prosla godina ne prolazi kroz cip Ova godina"
    ' Red bez citljivog datuma NE sme da prodje: propustiti ga znacilo bi
    ' tvrditi da je iz tekuce godine, a ne zna se.
    AssertEq modScrAgro.AgCipPromet("godina", MAG_ULAZ, 0), False, _
             "red bez datuma ne prolazi kroz cip godine"

    AssertEq modScrAgro.AgCipDugovi("duguju", 1), True, "dug veci od nule duguje"
    AssertEq modScrAgro.AgCipDugovi("duguju", 0), False, "nula ne duguje"
    AssertEq modScrAgro.AgCipDugovi("duguju", -500), False, _
             "pretplata nije dug -- kooperant kome je vise oduzeto ne duguje"

    modScrAgro.Scr_ListaTestSet "KORPA"
End Sub

' ============================================================
' 87. Brojac ceka na korpi; dvoklik bira po IDENTITETU
' ============================================================
' Brojac: korpa je jedino sto na ovom ekranu ceka operatera -- sve ostalo je
' vec u tabelama. Bez brojke operater koji predje na drugi ekran nema nijedan
' znak da mu je ostala puna.
'
' Dvoklik: lista dugova pokazuje IME, a dvoklik bira KOOPERANTA. Kad dva
' kooperanta nose isto ime, prikaz je dvosmislen i izbor se ODBIJA -- isto
' pravilo kao "dvosmislen broj -> MANUAL" u storno okviru. Fixture zato ima
' KOOP-TEST-1 i KOOP-TEST-IME, oba "Prvi Testni".
Private Sub T_Agro_BrojacIDvoklikPoIdentitetu()
    Dim d As Variant, greska As String

    ' --- BROJAC ---
    modScrAgro.Scr_KorpaTestReset
    AssertEq modScrAgro.Scr_Brojac(), 0, "prazna korpa ne ceka nista"

    greska = modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL, 1, FX_PARCELA)
    AssertEq greska, "", "stavka je usla u korpu"
    AssertEq modScrAgro.Scr_Brojac(), 1, "brojac vidi stavku koja ceka upis"

    greska = modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL, 1, FX_PARCELA)
    AssertEq greska, "", "i druga stavka je usla"
    AssertEq modScrAgro.Scr_Brojac(), 2, "brojac broji SVE sto ceka, ne samo prvu"

    modScrAgro.Scr_KorpaTestReset
    AssertEq modScrAgro.Scr_Brojac(), 0, "praznjenje korpe gasi brojac"

    ' --- DVOKLIK: mapa identiteta ---
    ' Mapu puni citac liste, pa se lista mora procitati pre nego sto se tvrdi.
    modScrAgro.Scr_ListaTestSet "DUGOVI"
    d = modScrAgro.Scr_Rows("sve", "")
    AssertEq IsArray(d), True, "preduslov: lista dugova je procitana"

    ' NAJVAZNIJE PRVO: dvosmislen prikaz nosi PRAZAN identitet. Mapa koja bi
    ' zapamtila prvog pogodjenog izgledala bi ispravno u svakoj drugoj tvrdnji,
    ' a dvoklik bi izdao robu pogresnom coveku.
    AssertEq modScrAgro.Scr_DugIdTest(FX_KOOP_PRIKAZ), "", _
             "dva kooperanta istog imena daju DVOSMISLEN prikaz, ne prvog"

    ' Jednoznacan prikaz i dalje daje svoj identitet -- kapija ne sme da obori
    ' sve redom.
    AssertEq modScrAgro.Scr_DugIdTest("Drugi Testni"), FX_KOOPERANT2, _
             "jednoznacan prikaz daje svoj KooperantID"
    AssertEq modScrAgro.Scr_DugIdTest("Ne Postoji"), "", _
             "nepoznat prikaz nema identitet"

    ' Ista kapija nad listom stanja: naziv artikla -> ArtikalID.
    modScrAgro.Scr_ListaTestSet "STANJE"
    d = modScrAgro.Scr_Rows("sve", "")
    AssertEq modScrAgro.Scr_ArtIdTest("Test Preparat"), FX_ARTIKAL, _
             "jednoznacan naziv artikla daje svoj ArtikalID"

    modScrAgro.Scr_ListaTestSet "KORPA"
    modScrAgro.Scr_KorpaTestReset
End Sub

' ============================================================
' 88. Mapa odbitaka i pojedinacni racun daju ISTO
' ============================================================
' GetAgroAbzugMapa je DRUGA implementacija pravila koje vec zivi u
' GetAgroAbzug. Postoji samo zbog brzine (lista dugova bi inace citala celu
' tblNovac po redu liste), ali obe su ZIVE u istoj funkciji: mapu zove lista
' dugova (GetAgroDugoviForGrid), a pojedinacnu kes ekrana (modScrAgro).
'
' Dve kopije istog pravila se tiho razilaze. Dodas tip uplate u jednu ili
' promenis izuzimanje storniranih, i ista aplikacija na dva mesta pokazuje
' RAZLICIT dug istom coveku -- bez ijednog crvenog testa.
Private Sub T_Agro_AbzugMapaPratiPojedinacni()
    Dim mapa As Object, k As Variant, n As Long

    Set mapa = modNovac.GetAgroAbzugMapa()
    AssertEq (mapa Is Nothing), False, "mapa odbitaka postoji"

    ' PREDUSLOV: mapa NIJE prazna. Bez ovoga bi petlja ispod prosla nula puta
    ' i test bi bio zelen ne merivsi nista -- tacno oblik placeba.
    AssertEq (mapa.count >= 2), True, _
             "PREDUSLOV: fixture ima odbitke za bar dva kooperanta"

    ' Tacne brojke, ne samo slaganje: 300 + 200 = 500. Storniranih 999 i
    ' uplata drugog tipa (777) se NE broje. Da se dve implementacije slome
    ' na ISTI nacin, puko poredjenje bi i dalje bilo zeleno.
    AssertEq CDbl(mapa(FX_KOOPERANT)), FX_ABZUG_KOOP1, _
             "mapa SABIRA odbitke i izuzima stornirane"
    AssertEq modNovac.GetAgroAbzug(FX_KOOPERANT), FX_ABZUG_KOOP1, _
             "pojedinacni racun daje isti zbir"
    AssertEq CDbl(mapa(FX_KOOPERANT2)), FX_ABZUG_KOOP2, _
             "mapa razdvaja kooperante -- ne slije sve u jedan zbir"

    ' NAJVAZNIJE: slaganje nad SVAKIM kooperantom koga mapa zna, ne samo nad
    ' dva imenovana. Kad neko sutra doda tip uplate u jednu implementaciju,
    ' ovo je tvrdnja koja pukne.
    For Each k In mapa.keys
        AssertEq CDbl(mapa(CStr(k))), modNovac.GetAgroAbzug(CStr(k)), _
                 "odbitak za " & CStr(k) & " isti u mapi i pojedinacno"
        n = n + 1
    Next k
    AssertEq (n >= 2), True, "petlja je stvarno prosla kroz kooperante"

    ' Kooperant BEZ ijednog odbitka: mapa ga ne zna, pojedinacni daje nulu.
    ' Odsustvo kljuca i nula moraju da znace isto, inace lista dugova za
    ' njega prikaze prazno umesto 0.
    AssertEq mapa.Exists(FX_KOOPERANT3), False, _
             "kooperant bez odbitka nije u mapi"
    AssertEq modNovac.GetAgroAbzug(FX_KOOPERANT3), 0, _
             "pojedinacni racun mu daje nulu -- isto znacenje"
End Sub

' ============================================================
' 89. Prekidac rezima ZADRZAVA boju posle izlaska pokazivaca
' ============================================================
' Kvar koji je prijavio operater na prvom smoke-u: izabran rezim je bio zelen
' samo dok je pokazivac nad njim, a cim predje dalje ispuna se vrati na BELU --
' natpis ostane krem (njegova labela je vezana kao "chev", nju reset ne dira),
' pa aktivno dugme postane skoro necitljivo.
'
' Uzrok nije bojenje nego PAMCENJE: clsFlatBtn zapamti osnovnu boju pri Bind-u i
' vraca je u ResetVisual kad pokazivac ode. BoxState menja kontrolu, ali ne i tu
' zapamcenu osnovu -- zato render koji promeni boju mora da javi novu kroz
' RebaseSink. Isti kvar je vec jednom placen u modScrStorno (StilDugmeta).
'
' Test ne trazi mis: ResetVisual se zove direktno nad sink-om, sto je tacno ono
' sto se desi kad pokazivac napusti dugme. Boja se cita PRE i POSLE.
Private Sub T_ZonaAgro_PrekidacRezimaZadrzavaBoju()
    Dim f As frmOtkupUI, z As Object, snimak As String, rez As String
    Dim izlI As Long, izlU As Long, izlRstI As Long, izlRstU As Long
    Dim ulzI As Long, ulzU As Long, ulzRstI As Long, ulzRstU As Long
    Dim gI As Long, gU As Long, gIC As Long, gUC As Long
    Dim rI As Long, rU As Long, rIC As Long, rUC As Long

    ' MERI SE DOK FORMA ZIVI, TVRDI SE POSLE Unload-a -- isti razlog kao u
    ' T_ZonaAgro_PoljaPostojeIPrateRezim: dok forma zivi, njena masinerija
    ' obrise Err izmedju Err.Raise i omotnice testa.
    '
    ' REZ SE MERI PREKO Font.Weight (400 normalan, 700 bold), ne preko
    ' Font.Bold. Font.Bold je iz tezine izveden i sam po sebi ume da prevari --
    ' sonda je kroz sest krugova pokazala i upis koji se izgubi i upis koji
    ' vrati staru vrednost. Tezina je broj i ne moze da bude "skoro tacna".
    Set f = NewOtkupUIForm()
    Set z = f.Controls.Add("Forms.Frame.1", "zProbaAgSeg", True)
    z.width = 1200: z.Height = 300
    modScrAgro.Scr_Build z

    ' stanje ODMAH posle gradnje -- ispuna i natpis oba segmenta
    gI = z.Controls("scrAgSegI").Font.Weight
    gU = z.Controls("scrAgSegU").Font.Weight
    gIC = z.Controls("scrAgSegIC").Font.Weight
    gUC = z.Controls("scrAgSegUC").Font.Weight

    ' --- IZDAVANJE je izabrano ---
    modScrAgro.Scr_RezimTestSet "IZLAZ"
    modScrAgro.Scr_Layout z, 1200, 300
    rez = modScrAgro.Scr_Rezim()
    rI = z.Controls("scrAgSegI").Font.Weight
    rU = z.Controls("scrAgSegU").Font.Weight
    rIC = z.Controls("scrAgSegIC").Font.Weight
    rUC = z.Controls("scrAgSegUC").Font.Weight
    izlI = z.Controls("scrAgSegI").BackColor
    izlU = z.Controls("scrAgSegU").BackColor
    ' Posle izlaska pokazivaca boja mora da OSTANE.
    ResetSinkVizual "scrAgSegI"
    ResetSinkVizual "scrAgSegU"
    izlRstI = z.Controls("scrAgSegI").BackColor
    izlRstU = z.Controls("scrAgSegU").BackColor

    ' --- PRIJEM: boje se zamene, i opet prezive ---
    modScrAgro.Scr_RezimTestSet "ULAZ"
    modScrAgro.Scr_Layout z, 1200, 300
    ulzI = z.Controls("scrAgSegI").BackColor
    ulzU = z.Controls("scrAgSegU").BackColor
    ResetSinkVizual "scrAgSegI"
    ResetSinkVizual "scrAgSegU"
    ulzRstI = z.Controls("scrAgSegI").BackColor
    ulzRstU = z.Controls("scrAgSegU").BackColor

    modScrAgro.Scr_RezimTestSet "IZLAZ"
    ReleaseOtkupUIForm f

    ' REZ, u jednoj tvrdnji. AssertEq staje na prvoj razlici, pa bi osam
    ' zasebnih tvrdnji pokazalo samo prvu vrednost -- a bas se meri da li se
    ' gradnja i raspored SLAZU. Natpis (IC/UC) ide uz ispunu (I/U) jer im
    ' BoxState pise isti rez: razlika izmedju njih dvoje je sama po sebi kvar.
    snimak = "gradnja I=" & gI & " U=" & gU & " IC=" & gIC & " UC=" & gUC & _
             " | raspored rez=" & rez & _
             " I=" & rI & " U=" & rU & " IC=" & rIC & " UC=" & rUC
    AssertEq snimak, _
             "gradnja I=700 U=400 IC=700 UC=400" & _
             " | raspored rez=IZLAZ I=700 U=400 IC=700 UC=400", _
             "rez prekidaca rezima (Font.Weight: 400 normalan, 700 bold)"

    ' PREDUSLOV: bojenje uopste radi. Bez ovoga bi tvrdnja ispod prolazila i nad
    ' dugmetom koje nikad nije ni pozelenelo.
    AssertEq izlI, CLng(modOtkupUI.C_FOREST), "preduslov: izabran rezim je obojen"
    AssertEq izlU, CLng(modOtkupUI.C_WHITE), "preduslov: neizabran rezim je beo"

    ' NAJVAZNIJE: posle izlaska pokazivaca boja OSTAJE. Ovo je jedina tvrdnja
    ' koja pada kad se izgubi RebaseSink -- sve ostalo izgleda ispravno.
    AssertEq izlRstI, CLng(modOtkupUI.C_FOREST), _
             "izabran rezim ostaje zelen i kad pokazivac ode"
    AssertEq izlRstU, CLng(modOtkupUI.C_WHITE), _
             "neizabran rezim ostaje beo i kad pokazivac ode"

    ' Bez ove polovine bi sabotaza koja zamrzne boje na prvoj vrednosti prosla:
    ' prvi rezim bi i dalje bio tacan.
    AssertEq ulzU, CLng(modOtkupUI.C_FOREST), "prekidac je presao na prijem"
    AssertEq ulzI, CLng(modOtkupUI.C_WHITE), "izdavanje je prestalo da bude izabrano"
    AssertEq ulzRstU, CLng(modOtkupUI.C_FOREST), _
             "prijem ostaje zelen i kad pokazivac ode"
    AssertEq ulzRstI, CLng(modOtkupUI.C_WHITE), _
             "izdavanje ostaje belo i kad pokazivac ode"
End Sub

' ============================================================
' 90. Traka korpe: najnovije prvo, i preliv se PRIJAVLJUJE
' ============================================================
' Nalaz sa smoke-a: korpa se videla samo dok je izabrana lista "Korpa", pa
' operater koji gleda stanje ili dugove nije imao nijedan znak sta je upravo
' dodao. Zona je zato dobila traku sa korpom.
'
' Dva pravila, oba se tiho kvare:
'   1. NAJNOVIJE PRVO -- operater upravo nesto doda, pa mu je potvrda ono sto
'      trazi. Traka koja pokazuje najstarije izgleda ispravno dok se korpa ne
'      napuni preko cetiri reda.
'   2. PRELIV SE PRIJAVLJUJE -- lista koja se tiho odseca izgleda kao cela.
'      Bas to je pravilo koje ljuska nad sobom vec ima (BazenStaje).
'
' Racun je odvojen od crtanja, pa se meri bez forme.
Private Sub T_Agro_TrakaKorpe_NajnovijePrvoIPreliv()
    Dim i As Long, greska As String, r0 As String

    modScrAgro.Scr_KorpaTestReset
    AssertEq modScrAgro.Scr_TrakaRedTest(0), "", "prazna korpa ne pise nista u traku"

    ' Jedna stavka: stoji u prvom redu, ostali su prazni.
    greska = modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL_ZALIHA, 1, FX_PARCELA)
    AssertEq greska, "", "preduslov: prva stavka je usla"
    r0 = modScrAgro.Scr_TrakaRedTest(0)
    AssertEq (Len(r0) > 0), True, "prva stavka se vidi u traci"
    ' Izdavanje se meri PAKOVANJIMA, pa red mora da nosi broj pakovanja.
    AssertEq (InStr(1, r0, "1 " & ChrW(215), vbTextCompare) > 0), True, _
             "red trake nosi broj pakovanja (1 x)"
    AssertEq (InStr(1, r0, "Test Zaliha", vbTextCompare) > 0), True, _
             "red trake nosi naziv artikla"
    AssertEq modScrAgro.Scr_TrakaRedTest(1), "", "drugi red je prazan dok ima jedna stavka"

    ' Cetiri stavke: sve staju, ali OBRNUTO -- poslednja dodata je prva.
    ' Svaka je razlicitog broja pakovanja da bi se redosled uopste video.
    modScrAgro.Scr_KorpaTestReset
    For i = 1 To 3
        AssertEq modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL_ZALIHA, i, FX_PARCELA), "", _
                 "preduslov: stavka " & i & " je usla"
    Next i
    ' NAJVAZNIJE: prvi red trake je POSLEDNJA dodata (3 pakovanja), ne prva.
    AssertEq (InStr(1, modScrAgro.Scr_TrakaRedTest(0), "3 " & ChrW(215), vbTextCompare) > 0), _
             True, "prvi red trake je POSLEDNJA dodata stavka"
    AssertEq (InStr(1, modScrAgro.Scr_TrakaRedTest(2), "1 " & ChrW(215), vbTextCompare) > 0), _
             True, "poslednji red trake je PRVA dodata stavka"
    AssertEq modScrAgro.Scr_TrakaRedTest(3), "", "cetvrti red je prazan dok ih ima tri"

    ' Vise nego sto staje: poslednji red kaze KOLIKO ih je sakriveno.
    ' Korpa ima sest stavki, traka nosi cetiri reda -> tri stavke + preliv "2".
    modScrAgro.Scr_KorpaTestReset
    For i = 1 To 6
        AssertEq modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL_ZALIHA, 1, FX_PARCELA), "", _
                 "preduslov: stavka " & i & " od sest je usla"
    Next i
    AssertEq (Len(modScrAgro.Scr_TrakaRedTest(2)) > 0), True, _
             "treci red jos nosi stavku"
    AssertEq (InStr(1, modScrAgro.Scr_TrakaRedTest(3), ChrW(8230), vbTextCompare) > 0), _
             True, "poslednji red je prelivni, ne cetvrta stavka"
    ' Sest stavki, tri prikazane -> sakriveno ih je TRI, i to mora da pise.
    AssertEq (InStr(1, modScrAgro.Scr_TrakaRedTest(3), "3", vbTextCompare) > 0), True, _
             "prelivni red kaze KOLIKO stavki je sakriveno"

    ' Traka ne izmislja redove preko svoje visine.
    AssertEq modScrAgro.Scr_TrakaRedTest(4), "", "traka nema peti red"

    modScrAgro.Scr_KorpaTestReset
End Sub

' ============================================================
' 91. Korpa se uklanja po IDENTITETU, ne po prikazu
' ============================================================
' Nalaz iz review-a PR #213: "Ukloni stavku" je stavku trazio po nazivu artikla
' i kolicini iz PRIKAZANOG reda. Dve iste stavke su tada nerazlucive, a to nije
' izmisljen slucaj -- "dva pakovanja sada, dva kasnije" daje dva reda iste robe
' i iste kolicine. Klik na drugi red je tada izbacivao PRVI, tiho: red koji
' nestane izgleda isto kao onaj koji je trebalo da nestane.
'
' Isto pravilo kao "dvosmislen broj -> MANUAL" u storno okviru, samo sto se
' ovde dvosmislenost moze SPRECITI umesto prijaviti: stavka nosi svoj identitet.
'
' Mereno bez mreze: mreza bi uvela sortiranje i stranice u tvrdnju koja je o
' identitetu. Ono sto mreza mora da uradi -- da identitet PRENESE i da ga ne
' nacrta -- tvrdi se nad opisom kolona i nad redovima koje Scr_Rows vraca.
Private Sub T_Agro_KorpaUklanjaPoIdentitetu()
    Dim id1 As String, id2 As String, d As Variant
    Dim kolone As Variant, redovi As Variant, n As Long

    modScrAgro.Scr_KorpaTestReset
    modScrAgro.Scr_ListaTestSet "KORPA"

    AssertEq modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL_ZALIHA, 2, FX_PARCELA), "", _
             "preduslov: prva stavka je usla"
    AssertEq modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL_ZALIHA, 2, FX_PARCELA), "", _
             "preduslov: druga -- u prikazu identicna -- stavka je usla"
    AssertEq modScrAgro.Scr_KorpaBroj(), 2, "preduslov: u korpi su dve stavke"

    id1 = modScrAgro.Scr_StavkaIdTest(1)
    id2 = modScrAgro.Scr_StavkaIdTest(2)
    AssertEq (Len(id1) > 0), True, "prva stavka nosi identitet"
    AssertEq (Len(id2) > 0), True, "druga stavka nosi identitet"
    ' NAJVAZNIJE: prikaz je isti, identitet NIJE.
    AssertEq (id1 <> id2), True, "dve stavke istog prikaza imaju RAZLICIT identitet"

    ' Identitet mora da stigne do mreze -- inace ga "Ukloni" nema odakle da cita.
    d = modScrAgro.Scr_Rows("sve", "")
    kolone = d(0): redovi = d(1): n = CLng(d(2))
    AssertEq n, 2, "mreza korpe ima dva reda"
    AssertEq UBound(kolone), 7, "korpa ima osam kolona -- osma nosi identitet"
    ' Prioritet 4, a mreza crta do 3: vrednost postoji u modelu, celija se nikad
    ' ne pravi. Bez toga bi operater u korpi gledao internu sifru.
    AssertEq Split(CStr(kolone(7)), "|")(4), "4", _
             "kolona identiteta je prioriteta 4 -- mreza je ne crta"
    AssertEq (CStr(redovi(1, 8)) <> CStr(redovi(2, 8))), True, _
             "redovi mreze nose razlicite identitete"

    ' Uklanjanje DRUGE stavke ostavlja PRVU. To je tvrdnja koju pretraga po
    ' nazivu i kolicini ne moze da zadovolji -- ona bi izbacila prvi red koji lici.
    AssertEq modScrAgro.Scr_UkloniStavkuTest(id2), True, _
             "uklanjanje po identitetu je proslo"
    AssertEq modScrAgro.Scr_KorpaBroj(), 1, "u korpi je ostala jedna stavka"
    AssertEq modScrAgro.Scr_StavkaIdTest(1), id1, _
             "ostala je bas ona stavka koja NIJE pokazana"

    ' Identitet kog nema ne sme nista da izbaci -- ni prazan, ni nepostojeci.
    ' Prazan stize sa reda mreze koji identitet nije poneo; tada se ne pogadja.
    AssertEq modScrAgro.Scr_UkloniStavkuTest("K-NEMA-OVAKVE"), False, _
             "nepoznat identitet ne uklanja nista"
    AssertEq modScrAgro.Scr_UkloniStavkuTest(""), False, _
             "prazan identitet ne uklanja nista"
    AssertEq modScrAgro.Scr_KorpaBroj(), 1, "korpa je posle promasaja nedirnuta"

    modScrAgro.Scr_KorpaTestReset
    modScrAgro.Scr_ListaTestSet "KORPA"
End Sub

' ============================================================
' 92. Znacka prati korpu i kad korpa NIJE prikazana lista
' ============================================================
' Nalaz iz review-a PR #213: ljuska brojace uz stavke menija pita samo kroz
' RefreshFromData, a nju zove tek kad ekran na klik javi True = "podaci su
' promenjeni". Ekran to javlja samo kad je korpa PRIKAZANA lista, jer bi inace
' terao ponovno citanje stanja ili prometa koje se nije menjalo.
'
' Posledica u pogonu: operater gleda STANJE, doda tri stavke, a znacka i dalje
' pise nulu -- pa predje na drugi ekran misleci da nema sta da proknjizi.
' Korpa nije podatak u tabeli, pa "podaci su promenjeni" i "korpa je promenjena"
' nisu ista stvar i ne smeju da dele isti kanal.
'
' Sta ovaj test NE pokriva: da bas DodajUKorpu / IsprazniKorpu / ZavrsiUnos zovu
' KorpaPromenjena. Te tri rutine citaju zonu, a zone u testu nema. Pokriveno je
' u kodu (nigde na tim mestima ne stoji goli OsveziZonu) i sabotazom
' agro-znacka-ne-prati-korpu.
Private Sub T_Agro_ZnackaPratiKorpuVanKorpeListe()
    Dim iD As String

    modScrAgro.Scr_KorpaTestReset
    ' LISTA NIJE KORPA -- to je ceo smisao testa.
    modScrAgro.Scr_ListaTestSet "STANJE"
    AssertEq modScrAgro.Scr_Lista(), "STANJE", "preduslov: prikazana lista nije korpa"
    AssertEq modScrAgro.Scr_ZnackaTest(), 0, "prazna korpa -> znacka je nula"

    AssertEq modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL_ZALIHA, 1, FX_PARCELA), "", _
             "preduslov: stavka je usla"
    AssertEq modScrAgro.Scr_KorpaBroj(), 1, "preduslov: korpa ima jednu stavku"
    AssertEq modScrAgro.Scr_ZnackaTest(), 1, _
             "znacka prati korpu i kad korpa NIJE prikazana lista"
    ' Znacka mora da dobije BAS ono sto ljuska cita iz ugovora ekrana -- inace
    ' bi ekran osvezavao neku svoju brojku koja sa sidebarom nema veze.
    AssertEq modScrAgro.Scr_ZnackaTest(), modScrAgro.Scr_Brojac(), _
             "znacka je dobila ono sto ljuska cita iz Scr_Brojac"

    AssertEq modScrAgro.Scr_KorpaTestDodaj(FX_ARTIKAL_ZALIHA, 1, FX_PARCELA), "", _
             "preduslov: druga stavka je usla"
    AssertEq modScrAgro.Scr_ZnackaTest(), 2, "znacka prati i drugu stavku"

    ' I uklanjanje je promena korpe.
    iD = modScrAgro.Scr_StavkaIdTest(1)
    AssertEq modScrAgro.Scr_UkloniStavkuTest(iD), True, "preduslov: stavka je uklonjena"
    AssertEq modScrAgro.Scr_ZnackaTest(), 1, "znacka prati uklanjanje"

    ' I praznjenje korpe.
    modScrAgro.Scr_KorpaTestReset
    AssertEq modScrAgro.Scr_ZnackaTest(), 0, "znacka prati praznjenje korpe"

    modScrAgro.Scr_ListaTestSet "KORPA"
End Sub

' Vrati kontrolu u "nije pod pokazivacem" stanje -- tacno ono sto clsFlatBtn
' uradi kad pokazivac napusti dugme. Isti tag nose i ispuna i natpis, pa se
' resetuju oba; za natpis ("chev") ResetVisual sam izlazi.
Private Sub ResetSinkVizual(ByVal tag As String)
    Dim b As clsFlatBtn
    On Error Resume Next
    If modOtkupUI.Btns Is Nothing Then Exit Sub
    For Each b In modOtkupUI.Btns
        If b.SinkTag = tag Then b.ResetVisual
    Next b
End Sub

' Ugasi magacin red koji je test napravio. Vracanje fixture-a, ne poslovna
' radnja -- zato ide kroz UpdateCell, a ne kroz storno rutinu.
Private Sub StornirajMagacinRed(ByVal magID As String)
    Dim redovi As Collection, i As Long
    Set redovi = FindRows(TBL_MAGACIN, COL_MAG_ID, magID)
    For i = 1 To redovi.count
        UpdateCell TBL_MAGACIN, CLng(redovi(i)), COL_STORNIRANO, "Da"
    Next i
End Sub

' Jedno polje contexta, po CorrectionID.

' ============================================================
' 93. Radnja nad paletom gadja RED, ne broj
' ============================================================
' Broj palete se RESETUJE po godini: GenerateBrojPalete racuna maxN+1 unutar
' Year(Date), pa 12/2025 i 12/2026 postoje istovremeno. Dok je ekran identitet
' resavao preko recnika broj->ID, taj recnik je za oba imao TACNO JEDAN unos --
' pa je stampa, zatvaranje i storno nad starijom paletom pogadjalo noviju.
'
' Kvar je tih: operater vidi red koji je izabrao, a radnja ode na drugi zapis.
'
' Test ide kroz PRAVI put: GridTestLoad puni mrezu (uz sortiranje), a
' Scr_IdZaRedTest vraca bas ono sto bi mutation path poslao nizvodno.
Private Sub T_PaleteIdentitet_PoIDNePoBroju()
    Dim d As Variant, i As Long, n As Long
    Dim redStara As Long, redNova As Long
    Dim idStara As String, idNova As String

    modScrPalete.Scr_PalTestSet "PALETE"

    ' 1) UGOVOR: lista se zavrsava kolonom identiteta, i ta kolona je
    ' NEVIDLJIVA. Vidljiva bi operateru prikazala internu sifru.
    d = modScrPalete.Scr_Rows("sve", "")
    AssertEq (UBound(d(0)) + 1), modScrPalete.PAL_KOL_ID, _
             "opis kolona se zavrsava BAS na koloni koju radnja cita"
    AssertEq Split(CStr(d(0)(modScrPalete.PAL_KOL_ID - 1)), "|")(0), "OTKUI_HD_IDENT", _
             "poslednja kolona je identitet palete"
    AssertEq modScrDokumenti.ColF(CStr(d(0)(modScrPalete.PAL_KOL_ID - 1)), 4), "4", _
             "kolona identiteta je prioriteta 4 -- nikad vidljiva"

    ' 2) Mreza se puni kao u ljusci -- preko nje radnja i cita red.
    modOtkupUI.GridTestLoad "PALETE"
    n = modOtkupUI.GridBrojRedova()
    AssertEq (n > 0), True, "preduslov: mreza je napunjena"

    ' PREDUSLOV koji nosi ceo test: fixture STVARNO ima dve palete istog
    ' broja u dve godine. Bez njega bi sve ispod prolazilo nad jednim redom
    ' i ne bi merilo nista.
    For i = 1 To n
        If Trim$(CStr(modOtkupUI.GridCell(i, 1))) = FX_PAL_KOL_BROJ Then
            If Trim$(CStr(modOtkupUI.GridCell(i, 2))) = "2025" Then redStara = i
            If Trim$(CStr(modOtkupUI.GridCell(i, 2))) = "2026" Then redNova = i
        End If
    Next i
    AssertEq (redStara > 0 And redNova > 0), True, _
             "PREDUSLOV: fixture ima paletu " & FX_PAL_KOL_BROJ & " i u 2025 i u 2026"

    ' 3) NAJVAZNIJE: svaki red daje SVOJ identitet.
    idStara = modScrPalete.Scr_IdZaRedTest("palprint", redStara)
    idNova = modScrPalete.Scr_IdZaRedTest("palprint", redNova)
    AssertEq idStara, FX_PAL_KOL_STARA, "stariji red daje SVOJ PaletaID"
    AssertEq idNova, FX_PAL_KOL_NOVA, "noviji red ostaje netaknut -- svoj PaletaID"
    AssertEq (idStara <> idNova), True, _
             "dva reda istog broja NE smeju da daju isti identitet"

    ' 4) Isti identitet vidi i izbor reda (aktivna paleta), ne samo radnje.
    AssertEq modScrPalete.Scr_AktivnaPaletaTest(redStara), FX_PAL_KOL_STARA, _
             "izbor reda postavlja aktivnu paletu po identitetu reda"
    AssertEq modScrPalete.Scr_AktivnaPaletaTest(redNova), FX_PAL_KOL_NOVA, _
             "izbor drugog reda ne vuce identitet prethodnog"

    ' 5) Red van skupa nema identitet -- pozivalac tada ne sme nista da menja.
    AssertEq modScrPalete.Scr_IdZaRedTest("palprint", n + 5), "", _
             "red van mreze nema identitet"

    modOtkupUI.GridTestLoad ""
    modScrPalete.Scr_PalTestSet "PALETE"
End Sub

' ============================================================
' 94. Radnja nad preradom gadja RED, ne broj
' ============================================================
' Isti kvar, druga lista: GenerateBrojPrerade takodje racuna maxN+1 unutar
' Year(Date). Zaseban test jer je i resolver imao zasebnu granu ("pre").
'
' Lista prerada NEMA kolonu godine, pa se dva reda istog broja razlikuju po
' netu (300 / 200) -- i bas zato je identitet iz reda jedini nacin da se
' pogodi pravi zapis.
Private Sub T_PreradeIdentitet_PoIDNePoBroju()
    Dim d As Variant, i As Long, n As Long
    Dim redStara As Long, redNova As Long
    Dim idStara As String, idNova As String

    modScrPalete.Scr_PalTestSet "PRERADE"

    d = modScrPalete.Scr_Rows("sve", "")
    AssertEq (UBound(d(0)) + 1), modScrPalete.PRE_KOL_ID, _
             "opis kolona se zavrsava BAS na koloni koju radnja cita"
    AssertEq Split(CStr(d(0)(modScrPalete.PRE_KOL_ID - 1)), "|")(0), "OTKUI_HD_IDENT", _
             "poslednja kolona je identitet prerade"
    AssertEq modScrDokumenti.ColF(CStr(d(0)(modScrPalete.PRE_KOL_ID - 1)), 4), "4", _
             "kolona identiteta je prioriteta 4 -- nikad vidljiva"

    modOtkupUI.GridTestLoad "PALETE"
    n = modOtkupUI.GridBrojRedova()
    AssertEq (n > 0), True, "preduslov: mreza prerada je napunjena"

    For i = 1 To n
        If Trim$(CStr(modOtkupUI.GridCell(i, 1))) = FX_PRE_KOL_BROJ Then
            If CDbl(val(CStr(modOtkupUI.GridCell(i, 6)))) = 200 Then redStara = i
            If CDbl(val(CStr(modOtkupUI.GridCell(i, 6)))) = 300 Then redNova = i
        End If
    Next i
    AssertEq (redStara > 0 And redNova > 0), True, _
             "PREDUSLOV: fixture ima preradu " & FX_PRE_KOL_BROJ & " u dve godine"

    idStara = modScrPalete.Scr_IdZaRedTest("prestorno", redStara)
    idNova = modScrPalete.Scr_IdZaRedTest("prestorno", redNova)
    AssertEq idStara, FX_PRE_KOL_STARA, "stariji red daje SVOJ PreradaID"
    AssertEq idNova, FX_PRE_KOL_NOVA, "noviji red ostaje netaknut -- svoj PreradaID"
    AssertEq (idStara <> idNova), True, _
             "dva reda istog broja NE smeju da daju isti identitet"

    ' Grana resolvera se bira po PREFIKSU radnje: 'pre' cita kolonu prerade,
    ' sve ostalo kolonu palete. Da grane nema, storno prerade bi citao praznu
    ' kolonu 13 i tiho odbio radnju.
    AssertEq modScrPalete.Scr_IdZaRedTest("preprint", redStara), FX_PRE_KOL_STARA, _
             "i preprint ide kroz istu granu resolvera"

    modOtkupUI.GridTestLoad ""
    modScrPalete.Scr_PalTestSet "PALETE"
End Sub

' ============================================================
' 95. Telo mreze ne ulazi u traku poruka
' ============================================================
' Traka poruka stoji tacno iznad podnozja (PostaviToast: footTop - TOAST_H - 4),
' a telo mreze je racunato sa rezervom od svega 6pt -- pa je poslednji red
' ulazio 24pt U traku. Poruka se crtala PREKO reda i drzala se samo ZOrder-om,
' sto resava redosled crtanja, ne prostor: red ispod poruke je bio necitljiv.
'
' Meri se na VISE visina, jer je racun linearan po zh i greska bi se na jednoj
' visini mogla slucajno poklopiti.
'
' Ispod ~195pt pobedjuje pod od tri reda (bodyH < GRID_ROW_H * 3) -- svesno:
' mreza koja pokaze manje od tri reda nije upotrebljiva. Zato se meri od 200.
Private Sub T_GridTelo_NePokrivaToast()
    Dim f As frmOtkupUI, z As Object, body As Object, toast As Object
    Dim h As Variant, nalaz As String, dodir As Long

    Set f = NewOtkupUIForm()
    Set z = f.Controls("zGrid")

    ' Nalazi se SKUPLJAJU, a tvrde tek posle Unload-a: dok forma zivi, njena
    ' masinerija obrise Err izmedju Err.Raise i omotnice testa.
    For Each h In Array(200, 260, 300, 420, 560, 700)
        modOtkupUI.GridLayoutTest z, 1200, CSng(h)
        Set body = z.Controls("grdBody")
        Set toast = z.Controls("tstScr")
        If body.top + body.Height > toast.top Then
            nalaz = nalaz & " zh=" & CStr(h) & "[telo do " & _
                    CStr(body.top + body.Height) & ", traka od " & _
                    CStr(toast.top) & "]"
        End If
        If body.top + body.Height = toast.top Then dodir = dodir + 1
    Next h
    Unload f

    AssertEq nalaz, "", _
             "telo mreze staje pre trake poruka -- body.Bottom <= toast.Top"
    ' Kontrola u drugom smeru: rezervacija ne sme da bude i prevelika. Bar
    ' jedna visina mora telo da dovede TACNO do trake -- inace bi test prolazio
    ' i kad bi mreza bila proizvoljno niska i gubila redove bez razloga.
    AssertEq (dodir > 0), True, _
             "rezervacija je tacno TOAST_H -- telo bar jednom stigne do trake"
End Sub

' ============================================================
' 96. Scr_Event ekrana Palete vraca cist Err
' ============================================================
' Isti ugovor koji modScrStorno.Scr_Event vec drzi (test 66): na USPESNOM
' izlazu Err.Number mora biti 0.
'
' Ovde je cela funkcija stajala pod On Error Resume Next i Err nikad nije
' cistila, pa je progutana greska iznutra ostajala u Err i posle povratka --
' ljuska je prijavljivala neuspeh za radnju koja je PROSLA.
Private Sub T_PaleteScrEvent_NeCuriGreska()
    Dim brojPosle As Long, brojLos As Long

    modScrPalete.Scr_PalTestSet "PALETE"

    ' 1) Obicna, uspesna radnja.
    Err.Clear
    modScrPalete.Scr_Event "lsPRERADE", "Click"
    ' Err se cita ODMAH: svaki poziv ispod (pa i AssertEq) ume da ga promeni.
    brojPosle = Err.Number
    Err.Clear

    ' 2) Dogadjaj koji IZNUTRA puca: CLng nad ne-brojem. Ovo je oblik koji je
    ' i curio -- Resume Next ga proguta, a Err ostane postavljen.
    modScrPalete.Scr_PalTestSet "PALETE"
    Err.Clear
    modScrPalete.Scr_Event "row:xyz", "Click"
    brojLos = Err.Number
    Err.Clear

    AssertEq brojPosle, 0, _
             "Scr_Event vraca cist Err -- inace ljuska javi neuspeh za radnju koja je prosla"
    AssertEq brojLos, 0, _
             "i kad dogadjaj iznutra pukne, Err ne curi kroz Scr_Event"

    ' Kontrola u drugom smeru: prekidac je stvarno obradjen, nije se samo
    ' progutao. Bez ovoga bi test prosao i kad Scr_Event ne radi nista.
    modScrPalete.Scr_Event "lsPRERADE", "Click"
    AssertEq modScrPalete.Scr_Lista(), "PRERADE", _
             "kontrola: Scr_Event i dalje obradjuje dogadjaj"

    modScrPalete.Scr_PalTestSet "PALETE"
End Sub
Private Function SvPolje(ByVal cid As String, ByVal kol As String) As String
    SvPolje = Trim$(NzToText(LookupValue(TBL_STORNO_VEZE, COL_SV_ID, cid, kol)))
End Function

' Vrati context u stanje iz fixture-a. Ne kroz modStornoContext -- tamo nema
' rutine koja terminalni status ponistava, i ne treba je ni biti: u produkciji
' je CANCELLED konacan. Ovo je iskljucivo ciscenje posle testa.
Private Sub VratiContextNaCekanje(ByVal cid As String)
    Dim rowIdx As Long
    rowIdx = modStornoContext.GetCorrectionRowByID(cid)
    If rowIdx = 0 Then Exit Sub
    UpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_STATUS, SV_STATUS_PENDING
    UpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_NEEDS_RECOVERY, "Da"
    UpdateCell TBL_STORNO_VEZE, rowIdx, COL_SV_COMPLETED_AT, ""
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

'=====================================================================
' EKRAN FAKTURISANJE (Faza E, stavka 16 -- v6-ui-176)
'=====================================================================

' Koliko stavki ima opis oblika "a|b|c". Prazan opis nosi NULA stavki, ne
' jednu -- Split("", "|") vraca niz od jednog praznog clana.
Private Function BrojStavkiOpisa(ByVal spec As String) As Long
    If Len(spec) = 0 Then Exit Function
    BrojStavkiOpisa = UBound(Split(spec, "|")) + 1
End Function

' Red citaca ciji je i-ti clan jednak trazenom. Vraca 0 kad ga nema.
Private Function RedSaVrednoscu(ByVal src As Variant, ByVal kol As Long, _
                                ByVal trazeno As String) As Long
    Dim i As Long
    If Not IsArray(src) Then Exit Function
    For i = 1 To UBound(src, 1)
        If Trim$(CStr(src(i, kol))) = trazeno Then
            RedSaVrednoscu = i
            Exit Function
        End If
    Next i
End Function

' UGOVOR EKRANA i GRANICE BAZENA LJUSKE. Visak se ne prijavljuje kao greska
' nego se TIHO odseca -- operater vidi ekran kome fali dugme, bez ijedne
' poruke (v6-ui-162, v6-ui-170). Lista SEF-a stoji TACNO na granici, pa je
' ovo jedino mesto koje primeti sestu radnju pre nego sto nestane.
Private Sub T_Fak_UgovorEkrana()
    Dim liste As Variant, i As Long, kljucevi As String
    Dim kljuc As String, spec As String, d As Variant, kolone As Variant
    ' For Each trazi Variant ili Object -- String iterator je "Type mismatch".
    Dim kv As Variant

    AssertEq (Len(modUiScreens.ScrRowByKey("FAKTURE")) > 0), True, _
             "FAKTURE postoji u registru ekrana"
    AssertEq modUiScreens.ScrField(modUiScreens.ScrRowByKey("FAKTURE"), SCR_MODUL), _
             "modScrFakture", "registar vodi ekran na modScrFakture"
    AssertEq modUiScreens.ScrPostoji("FAKTURE"), True, _
             "modul ekrana odgovara na Scr_Meta -- stavka menija vise nije prigusena"
    AssertEq (InStr(modUiScreens.ScrMeta("FAKTURE"), "kljuc=FAKTURE") > 0), True, _
             "Scr_Meta prijavljuje svoj kljuc"
    AssertEq modUiScreens.ScrField(modUiScreens.ScrRowByKey("FAKTURE"), SCR_OBLAST), _
             OBL_FAKTURISANJE, "ekran trazi pravo na oblast Fakturisanje"

    ' SEF LISTA POSTOJI UVEK. Citanje stanja su kolone tblFakture i ne trazi
    ' nikakvu vezu; kapiju trazi samo RADNJA, i ona je ima. Uslovna lista je
    ' novi UI cinila uzim od legacy-ja, koji frmSEF otvara bezuslovno.
    liste = modScrFakture.Scr_Liste()
    AssertEq (UBound(liste) + 1 <= modOtkupUI.MaxPrekidaca()), True, _
             "ljuska crta sve liste ekrana -- nijedna se ne odseca tiho"
    AssertEq UBound(liste) + 1, 3, _
             "tri liste, i kad SEF nije podesen"
    ' Ekran koji stoji na SEF listi tu i ostaje -- nema uslovnog vracanja.
    modScrFakture.Scr_FkListaTestSet "SEF"
    AssertEq modScrFakture.Scr_Lista(), "SEF", _
             "SEF lista se ne napusta zbog konfiguracije"
    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
    For i = 0 To UBound(liste)
        kljucevi = kljucevi & "|" & Split(CStr(liste(i)), "|")(0)
    Next i
    AssertEq kljucevi, "|ZAFAKT|FAKTURE|SEF", _
             "redosled i kljucevi lista -- prijemnice su prve, one su posao"

    ' RADNJE PO LISTI. Citaju se po KLJUCU, ne kroz Scr_Lista: Scr_Lista je
    ' gate-ovana SEF konfiguracijom, a fixture je donor-zavisan (tblConfig se
    ' u make_fixture ne brise), pa bi test vezan za nju bio lutrija.
    AssertEq BrojStavkiOpisa(modScrFakture.FkRadnjeZaListu("ZAFAKT")), 2, _
             "prijemnice: dodaj i ukloni"
    AssertEq BrojStavkiOpisa(modScrFakture.FkRadnjeZaListu("FAKTURE")), 2, _
             "fakture: stampaj i osvezi status"
    AssertEq BrojStavkiOpisa(modScrFakture.FkRadnjeZaListu("SEF")), modOtkupUI.MAX_ACT, _
             "SEF ima TACNO MAX_ACT radnji -- sesta bi se tiho odsekla"

    ' CIPOVI. Prvi je svuda 'sve' -- ljuska na njega pada kad zatecen filter ne
    ' pripada listi (RefreshChipsForScreen), pa prvi mora da bude NAJSIRI;
    ' povratak na uzi cip bi tiho sakrio redove.
    For Each kv In Array("ZAFAKT", "FAKTURE", "SEF")
        spec = modScrFakture.FkCipoviZaListu(CStr(kv))
        AssertEq (BrojStavkiOpisa(spec) <= modOtkupUI.MAX_CHIP), True, _
                 "lista " & kv & " ne trazi vise cipova nego sto bazen ima"
        AssertEq Split(Split(spec, "|")(0), ":")(0), "sve", _
                 "prvi cip liste " & kv & " je najsiri ('sve')"
        kolone = modScrFakture.FkKoloneZaListu(CStr(kv))
        AssertEq (UBound(kolone) + 1 <= modOtkupUI.MAX_COLS), True, _
                 "lista " & kv & " ne trazi vise kolona nego sto mreza pravi"
    Next kv

    ' Svaka DOSTUPNA lista mora da vrati ispravan niz. Lista koja pukne se u
    ' ljusci pretvara u Empty, LoadGridFromScreen na ne-niz radi Exit Sub -- pa
    ' mreza ostane na prethodnoj listi i prekidac izgleda kao da ne radi.
    modScrFakture.Scr_FkKupacTestSet FX_KUPAC
    For i = 0 To UBound(liste)
        kljuc = Split(CStr(liste(i)), "|")(0)
        modScrFakture.Scr_FkListaTestSet kljuc
        AssertEq modScrFakture.Scr_Lista(), kljuc, "lista " & kljuc & " je prihvacena"
        d = modScrFakture.Scr_Rows("sve", "")
        AssertEq IsArray(d), True, "lista " & kljuc & " vraca niz"
        AssertEq (UBound(d) >= 4), True, _
                 "lista " & kljuc & " vraca pun oblik (kolone, redovi, n, kg, vrednost)"
        AssertEq IsArray(d(0)), True, "lista " & kljuc & " prijavljuje svoje kolone"
    Next i

    modScrFakture.Scr_FkKorpaTestReset
    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
End Sub

' IDENTITET IDE U RED I NE CRTA SE. Mreza redove sortira i deli na strane, pa
' bi svaka mapa "prikaz -> ID" koju ekran drzi sa strane zastarela na prvi
' klik po zaglavlju. Kolona je prioriteta 4, a LayoutGrid crta do 3.
Private Sub T_Fak_IdentitetURedu_NeCrtaSe()
    Dim kljuc As Variant, kolone As Variant, poslednja As String
    Dim src As Variant, r As Long, d As Variant, redovi As Variant
    Dim dupli As Object

    ' Svaka lista nosi identitet u POSLEDNJOJ koloni, prioriteta 4.
    For Each kljuc In Array("ZAFAKT", "FAKTURE", "SEF")
        kolone = modScrFakture.FkKoloneZaListu(CStr(kljuc))
        poslednja = CStr(kolone(UBound(kolone)))
        AssertEq Split(poslednja, "|")(4), "4", _
                 "lista " & kljuc & ": kolona identiteta je prioriteta 4"
        AssertEq Split(poslednja, "|")(3), "1", _
                 "lista " & kljuc & ": kolona identiteta ne trazi sirinu"
    Next kljuc

    ' PRAVILO DVOSMISLENOSTI, mereno direktno. Duplikat ID-a se u fixture ne
    ' moze posejati a da ne obori kapije koje o njemu nista ne znaju
    ' (RequireSingleFakturaRow, CreateFaktura), pa se meri sam racun.
    Set dupli = CreateObject("Scripting.Dictionary")
    dupli("JEDAN") = 1
    dupli("DVA") = 2
    AssertEq modFaktura.IdIliPrazno(dupli, "JEDAN"), "JEDAN", _
             "jedinstven ID prolazi kao identitet"
    AssertEq modFaktura.IdIliPrazno(dupli, "DVA"), "", _
             "ID koji postoji dvaput NIJE identitet -- radnja ne sme da pogadja"
    AssertEq modFaktura.IdIliPrazno(dupli, "NEMA"), "", "nepoznat ID nije identitet"
    AssertEq modFaktura.IdIliPrazno(dupli, ""), "", "prazan ID nije identitet"

    ' Brojac gleda SIROVU tabelu -- i stornirani red istog ID-a cini ID
    ' dvosmislenim, jer ga FindRows (koji na kraju odlucuje) i dalje vidi.
    Set dupli = modFaktura.BrojacIdova(TBL_FAKTURE, COL_FAK_ID)
    AssertEq CLng(dupli(FX_FAK_STORNO)), 1, _
             "brojac vidi i storniranu fakturu -- njen ID nije slobodan"

    ' Identitet stize iz citaca u red mreze i tamo ostaje citljiv radnji.
    src = modFaktura.GetFaktureForGrid()
    r = RedSaVrednoscu(src, 1, FX_FAK_PLAC)
    AssertEq (r > 0), True, "citac faktura vraca placenu fakturu iz fixture-a"

    modScrFakture.Scr_FkListaTestSet "FAKTURE"
    d = modScrFakture.Scr_Rows("sve", "")
    redovi = d(1)
    AssertEq (CLng(d(2)) > 0), True, "lista faktura ima redova"
    r = RedSaVrednoscu(redovi, 8, FX_FAK_PLAC)
    AssertEq (r > 0), True, "red mreze NOSI FakturaID u skrivenoj koloni 8"
    AssertEq CStr(redovi(r, 1)), "3/2026", "isti red pokazuje broj fakture"

    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
End Sub

' DOSTUPNOST SE PRENOSI U REDU, ne izvodi se iz prikaza. Prijemnica
' obelezena kao fakturisana a BEZ FakturaID iz prikaza izgleda slobodna
' (kolona broja fakture je prazna), a CreateFaktura je odbija. Ko dostupnost
' cita iz onoga sto se vidi, ponudi je operateru pa padne u transakciji.
Private Sub T_Fak_DostupnostSePrenosiURedu()
    Dim src As Variant, r As Long, d As Variant, redovi As Variant, i As Long
    Dim greska As String

    ' PRAVILO, samo po sebi. Jedno mesto -- deli ga i kapija
    ' IsPrijemnicaAvailableForFaktura i citac mreze.
    AssertEq modFaktura.PrijemnicaDostupna("", "", ""), True, _
             "neobelezena prijemnica sme u fakturu"
    AssertEq modFaktura.PrijemnicaDostupna("Da", "", ""), False, _
             "stornirana ne sme"
    AssertEq modFaktura.PrijemnicaDostupna("", "Da", ""), False, _
             "obelezena kao fakturisana ne sme -- i kad FakturaID nedostaje"
    AssertEq modFaktura.PrijemnicaDostupna("", "", "FAK-1"), False, _
             "vezana za fakturu ne sme -- i kad oznaka nedostaje"

    ' NAD PRAVIM REDOVIMA. FAK2 je red zbog koga ovaj test postoji.
    src = modFaktura.GetPrijemniceZaFakturisanjeForGrid(FX_KUPAC)
    AssertEq IsArray(src), True, "citac vraca prijemnice kupca"

    r = RedSaVrednoscu(src, 1, FX_PRJ_FAK1)
    AssertEq (r > 0), True, "preduslov: uredno fakturisana prijemnica je u listi"
    AssertEq CBool(src(r, 9)), False, "uredno fakturisana nije dostupna"
    AssertEq CStr(src(r, 10)), "3/2026", "uredno fakturisana pokazuje svoj broj fakture"

    r = RedSaVrednoscu(src, 1, FX_PRJ_FAK2)
    AssertEq (r > 0), True, "preduslov: nepotpuno obelezena prijemnica je u listi"
    AssertEq CStr(src(r, 10)), "", "PRIKAZ je prazan -- iz njega izgleda slobodna"
    AssertEq CBool(src(r, 9)), False, "PRAVILO kaze da nije slobodna"

    r = RedSaVrednoscu(src, 1, FX_PRJ_FAK3)
    AssertEq (r > 0), True, "preduslov: slobodna prijemnica je u listi"
    AssertEq CBool(src(r, 9)), True, "slobodna prijemnica je dostupna"

    ' Red mreze mora da PRENESE dostupnost -- inace je ekran u trenutku klika
    ' nema odakle da procita.
    modScrFakture.Scr_FkKupacTestSet FX_KUPAC
    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
    d = modScrFakture.Scr_Rows("sve", "")
    redovi = d(1)
    r = RedSaVrednoscu(redovi, 10, FX_PRJ_FAK2)
    AssertEq (r > 0), True, "red nepotpuno obelezene prijemnice je u mrezi"
    AssertEq CStr(redovi(r, 11)), "", "red NOSI 'nije dostupna' u koloni 11"
    r = RedSaVrednoscu(redovi, 10, FX_PRJ_FAK3)
    AssertEq CStr(redovi(r, 11)), "1", "red slobodne prijemnice nosi '1'"

    ' Cip 'ceka' propusta tacno one koje pravilo pusta.
    d = modScrFakture.Scr_Rows("ceka", "")
    redovi = d(1)
    AssertEq RedSaVrednoscu(redovi, 10, FX_PRJ_FAK2), 0, _
             "cip 'za fakturisanje' ne pokazuje nepotpuno obelezenu"
    AssertEq (RedSaVrednoscu(redovi, 10, FX_PRJ_FAK3) > 0), True, _
             "cip 'za fakturisanje' pokazuje slobodnu"

    ' I kapija korpe. Nedostupna prijemnica se ODBIJA porukom, ne tiho.
    modScrFakture.Scr_FkKorpaTestReset
    modScrFakture.Scr_FkKupacTestSet FX_KUPAC
    greska = modScrFakture.Scr_FkKorpaTestDodaj(FX_PRJ_FAK2, "21/150326", 100, 40, False)
    AssertEq (Len(greska) > 0), True, "nedostupna prijemnica ne ulazi u korpu"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 0, "korpa je posle odbijanja prazna"

    modScrFakture.Scr_FkKorpaTestReset
End Sub

' KORPA, ZNACKA I TRAKA. Korpa NIJE podatak u tabeli, pa "podaci su promenjeni"
' i "korpa je promenjena" nisu ista stvar i ne smeju da dele isti kanal:
' ljuska brojace pita samo kroz RefreshFromData, a nju zove tek kad ekran na
' klik javi True. Bez sopstvenog kanala operater koji gleda listu faktura doda
' stavke, a znacka i dalje pise nulu.
Private Sub T_Fak_KorpaZnackaITraka()
    Dim i As Long

    modScrFakture.Scr_FkKorpaTestReset
    ' LISTA NIJE KORPA -- to je ceo smisao prve polovine testa.
    modScrFakture.Scr_FkListaTestSet "FAKTURE"
    AssertEq modScrFakture.Scr_Lista(), "FAKTURE", "preduslov: prikazana lista nije korpa"
    AssertEq modScrFakture.Scr_FkZnackaTest(), 0, "prazna korpa -> znacka je nula"

    AssertEq modScrFakture.Scr_FkKorpaTestDodaj(FX_PRJ_FAK3, "22/150326", 100, 40, True), "", _
             "preduslov: slobodna prijemnica je usla u korpu"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 1, "korpa ima jednu stavku"
    AssertEq modScrFakture.Scr_FkZnackaTest(), 1, _
             "znacka prati korpu i kad korpa NIJE prikazana lista"
    AssertEq modScrFakture.Scr_FkZnackaTest(), modScrFakture.Scr_Brojac(), _
             "znacka je dobila ono sto ljuska cita iz Scr_Brojac"

    ' Ista prijemnica dvaput je ista prijemnica -- fakturisala bi se dvaput.
    AssertEq (Len(modScrFakture.Scr_FkKorpaTestDodaj(FX_PRJ_FAK3, "22/150326", 100, 40, True)) > 0), _
             True, "ista prijemnica ne moze dvaput u korpu"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 1, "odbijeno dodavanje ne menja korpu"

    ' Druga stavka, pa uklanjanje PO IDENTITETU. Prikaz je namerno isti
    ' (isti broj, ista kolicina, ista cena): po prikazu se ova dva reda ne bi
    ' razlikovala, a to je bas greska koju je Agrohemija vec platila (P1, 7.9).
    AssertEq modScrFakture.Scr_FkKorpaTestDodaj(FX_PRJ_FAK1, "22/150326", 100, 40, True), "", _
             "preduslov: druga -- u prikazu identicna -- stavka je usla"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 2, "u korpi su dve stavke"
    AssertEq modScrFakture.Scr_FkZnackaTest(), 2, "znacka prati i drugu stavku"

    AssertEq modScrFakture.Scr_FkUkloniStavkuTest(FX_PRJ_FAK1), True, _
             "uklanjanje po identitetu je proslo"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 1, "ostala je jedna stavka"
    AssertEq modScrFakture.Scr_FkStavkaIdTest(1), FX_PRJ_FAK3, _
             "ostala je bas ona koja NIJE pokazana"
    AssertEq modScrFakture.Scr_FkZnackaTest(), 1, "znacka prati uklanjanje"

    AssertEq modScrFakture.Scr_FkUkloniStavkuTest("PRJ-NEMA-OVAKVE"), False, _
             "nepoznat identitet ne uklanja nista"
    AssertEq modScrFakture.Scr_FkUkloniStavkuTest(""), False, _
             "prazan identitet ne uklanja nista"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 1, "korpa je posle promasaja nedirnuta"

    modScrFakture.Scr_FkKorpaTestReset
    AssertEq modScrFakture.Scr_FkZnackaTest(), 0, "znacka prati praznjenje korpe"

    ' TRAKA KORPE. Dva pravila, oba se tiho kvare: NAJNOVIJE PRVO (operater
    ' upravo nesto doda, pa mu je potvrda ono sto trazi) i PRELIV SE PRIJAVLJUJE
    ' (lista koja se tiho odseca izgleda kao cela). Traka prima cetiri reda, pa
    ' se pet stavki koristi da se preliv uopste pojavi.
    For i = 1 To 5
        AssertEq modScrFakture.Scr_FkKorpaTestDodaj("PRJ-T" & i, i & "/T", 1, 100 * i, True), "", _
                 "preduslov: stavka " & i & " je usla u korpu"
    Next i
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 5, "preduslov: pet stavki u korpi"
    AssertEq (InStr(modScrFakture.Scr_FkTrakaRedTest(0), "5/T") > 0), True, _
             "prvi red trake je POSLEDNJA dodata stavka"
    AssertEq (InStr(modScrFakture.Scr_FkTrakaRedTest(1), "4/T") > 0), True, _
             "drugi red trake je pretposlednja"
    AssertEq (InStr(modScrFakture.Scr_FkTrakaRedTest(3), ChrW(8230)) > 0), True, _
             "poslednji red trake PRIJAVLJUJE preliv, ne cuti o njemu"
    AssertEq (InStr(modScrFakture.Scr_FkTrakaRedTest(3), "2") > 0), True, _
             "preliv kaze KOLIKO ih se ne vidi (5 stavki, 3 reda + preliv)"

    modScrFakture.Scr_FkKorpaTestReset
    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
End Sub

' CIPOVI LISTE FAKTURA I ZNAK U REDU MORAJU DA SE SLAZU -- medjusobno i sa
' modNovac.GetOpenFakture, jedinim read-modelom otvorenih faktura kupca.
' Pravilo 'otvorena faktura' zivi na dva mesta, pa moze da se razidje; ovo je
' isti oblik tvrdnje kao T_Agro_AbzugMapaPratiPojedinacni.
Private Sub T_Fak_CipoviPrateStatusFakture()
    Dim src As Variant, r As Long, i As Long, kup As Variant
    Dim otvorene As Variant, j As Long, iD As String
    Dim izCitaca As Long, izOpen As Long

    ' SIFRA ZNAKA U REDU (paypill). Ljuska crta iz sifre, ne iz teksta.
    AssertEq modScrFakture.FkPayKod(4000, 4000), PAY_PLACENO, "pun iznos -> placeno"
    AssertEq modScrFakture.FkPayKod(4000, 5000), PAY_PLACENO, "preplaceno je i dalje placeno"
    AssertEq modScrFakture.FkPayKod(5000, 2000), PAY_DELIM, "delimicna uplata -> delimicno"
    AssertEq modScrFakture.FkPayKod(5000, 0), PAY_NEPLAC, "bez uplate -> neplaceno"
    ' Faktura iznosa 0 nije placena nego PRAZNA. Da je 'placena', cip i znak u
    ' istom redu bi tvrdili suprotno.
    AssertEq modScrFakture.FkPayKod(0, 0), PAY_NEPLAC, "faktura bez iznosa nije placena"
    AssertEq modScrFakture.FkCipFaktura("plac", "Placeno", 0, 0, FX_DATUM), False, _
             "cip 'placene' ne uzima fakturu bez iznosa -- slaze se sa znakom u redu"
    AssertEq modScrFakture.FkCipFaktura("plac", "", 4000, 4000, FX_DATUM), True, _
             "cip 'placene' gleda uplatu, ne zapisan status"

    ' Cip 'neplacene' ima ISTA dva uslova kao GetOpenFakture: zapisan status
    ' Neplaceno I nesto stvarno preostalo.
    AssertEq modScrFakture.FkCipFaktura("nepl", STATUS_NEPLACENO, 5000, 0, FX_DATUM), True, _
             "neplacena sa ostatkom prolazi"
    AssertEq modScrFakture.FkCipFaktura("nepl", STATUS_NEPLACENO, 5000, 5000, FX_DATUM), False, _
             "bez ostatka ne prolazi ni sa statusom Neplaceno"
    AssertEq modScrFakture.FkCipFaktura("nepl", STATUS_PLACENO, 5000, 0, FX_DATUM), False, _
             "sa statusom Placeno ne prolazi ni sa ostatkom"
    AssertEq modScrFakture.FkCipFaktura("nepoznat", STATUS_PLACENO, 5000, 0, FX_DATUM), True, _
             "nepoznat filter pusta sve -- ekran ne pokazuje praznu listu"

    ' NAD PRAVIM REDOVIMA. Stornirana faktura ne sme da se pojavi -- inace bi
    ' joj operater nudio stampu i slanje na SEF.
    src = modFaktura.GetFaktureForGrid()
    AssertEq IsArray(src), True, "citac faktura vraca redove"
    AssertEq RedSaVrednoscu(src, 1, FX_FAK_STORNO), 0, _
             "stornirana faktura NIJE u listi"

    r = RedSaVrednoscu(src, 1, FX_FAK_PLAC)
    AssertEq CDbl(src(r, 6)), FX_FAK_PLAC_IZNOS, "uplata je sabrana po fakturi"
    AssertEq CDbl(src(r, 7)), 0, "placena faktura nema ostatak"
    r = RedSaVrednoscu(src, 1, FX_FAK_NEPL)
    AssertEq CDbl(src(r, 6)), 0, "neplacena faktura nema uplate"
    AssertEq CDbl(src(r, 7)), FX_FAK_NEPL_IZNOS, "ceo iznos je preostao"
    AssertEq CStr(src(r, 9)), FX_KUPAC2, "citac vraca i IDENTITET kupca, ne samo naziv"

    ' SLAGANJE SA GetOpenFakture, po SVAKOM kupcu. Da se dve implementacije
    ' istog pravila raziju, ovde bi se brojevi razisli.
    For Each kup In Array(FX_KUPAC, FX_KUPAC2)
        izCitaca = 0
        For i = 1 To UBound(src, 1)
            If CStr(src(i, 9)) = CStr(kup) Then
                If modScrFakture.FkCipFaktura("nepl", CStr(src(i, 8)), _
                                              CDbl(src(i, 5)), CDbl(src(i, 6)), src(i, 3)) Then
                    izCitaca = izCitaca + 1
                End If
            End If
        Next i
        otvorene = GetOpenFakture(CStr(kup))
        izOpen = 0
        If IsArray(otvorene) Then izOpen = UBound(otvorene, 1)
        AssertEq izCitaca, izOpen, _
                 "cip 'neplacene' i GetOpenFakture vide isto za " & kup
        ' I to bas iste fakture, ne samo isti broj.
        For j = 1 To izOpen
            iD = Trim$(CStr(otvorene(j, 2)))
            r = RedSaVrednoscu(src, 1, iD)
            AssertEq (r > 0), True, "otvorena faktura " & iD & " je u listi ekrana"
            AssertEq modScrFakture.FkCipFaktura("nepl", CStr(src(r, 8)), _
                                                CDbl(src(r, 5)), CDbl(src(r, 6)), src(r, 3)), _
                     True, "cip 'neplacene' propusta otvorenu fakturu " & iD
        Next j
    Next kup
End Sub

' NERAZRESEN UNOS NIJE PROMENA KUPCA. Ljuska Change salje ekranu na SVAKI
' znak, a GetComboID daje stabilan ID samo dok je stavka stvarno izabrana
' (ListIndex >= 0); cim operater krene da kuca, fallback iz parcijalnog teksta
' vrati "". Bez ove razlike bi prvo otkucano slovo bacilo celu neproknjizenu
' korpu -- a da drugi kupac nije ni izabran.
'
' Ovo je gubitak operaterskog rada, ne pokvaren podatak, ali je jedina stvar
' na ovom ekranu koja bez traga unistava ono sto je covek vec uradio.
' Znacka se ovde NE tvrdi iako je operater i nju gubio: to je posao testa
' 100. Prepisana tvrdnja bi znacila da sabotaza znacke obara DVA testa, pa
' dvosmerni dokaz vise ne bi pokazivao 'tacno jedan test, po imenu'.
Private Sub T_Fak_NerazresenKupacNeDiraKorpu()
    ' PRAVILO, samo po sebi.
    AssertEq modScrFakture.FkKupacPromenjen("", FX_KUPAC), False, _
             "prazan ID nije promena kupca -- to je nerazresen unos"
    AssertEq modScrFakture.FkKupacPromenjen("   ", FX_KUPAC), False, _
             "ni sam razmak nije promena kupca"
    AssertEq modScrFakture.FkKupacPromenjen(FX_KUPAC2, FX_KUPAC), True, _
             "razresen DRUGI kupac jeste promena"
    AssertEq modScrFakture.FkKupacPromenjen(FX_KUPAC, FX_KUPAC), False, _
             "isti kupac nije promena"
    AssertEq modScrFakture.FkKupacPromenjen(FX_KUPAC, ""), True, _
             "prvi izbor kupca jeste promena"

    ' NAD PRAVOM KORPOM, istim putem kojim ide Change dogadjaj.
    modScrFakture.Scr_FkKorpaTestReset
    modScrFakture.Scr_FkKupacTestSet FX_KUPAC
    AssertEq modScrFakture.Scr_FkKupacUnosTest(FX_KUPAC), True, _
             "preduslov: kupac je izabran"
    AssertEq modScrFakture.Scr_FkKorpaTestDodaj(FX_PRJ_FAK3, "22/150326", 100, 40, True), "", _
             "preduslov: prva stavka je usla u korpu"
    AssertEq modScrFakture.Scr_FkKorpaTestDodaj(FX_PRJ_FAK1, "20/150326", 100, 40, True), "", _
             "preduslov: druga stavka je usla u korpu"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 2, "preduslov: u korpi su dve stavke"

    ' KUCANJE: nerazresen unos NE SME nista da dirne.
    AssertEq modScrFakture.Scr_FkKupacUnosTest(""), False, _
             "nerazresen unos se ne tretira kao promena kupca"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 2, _
             "korpa prezivljava kucanje po polju kupca"

    ' STVARAN IZBOR DRUGOG KUPCA: tek tada se korpa prazni.
    AssertEq modScrFakture.Scr_FkKupacUnosTest(FX_KUPAC2), True, _
             "razresen drugi kupac JESTE promena"
    AssertEq modScrFakture.Scr_FkKorpaBroj(), 0, _
             "korpa se prazni tek kad je drugi kupac stvarno izabran"

    modScrFakture.Scr_FkKorpaTestReset
    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
End Sub

' GRESKA SE CITA PRE LogErr-a, INACE JE VISE NEMA.
'
' modLogError.LogError pocinje sa `On Error Resume Next`, a svaka On Error
' naredba u VBA BRISE Err objekat. Zato `Err.Raise Err.Number, SRC,
' Err.Description` POSLE LogErr-a postane `Err.Raise 0, SRC, ""` -- pozivalac
' dobije prazan opis, a citac mreze koji je trebalo da propagira pad seme
' stigne do ekrana kao 'nema redova'.
'
' Test ima dva dela: prvo dokazuje da je opasnost STVARNA (LogErr brise Err),
' pa onda meri PRAVI put -- modFaktura.PrintFaktura nad fakturom koje nema.
' Bas tu funkciju zove radnja ekrana 'Stampaj', pa je ovo i tvrdnja o tome
' sta operater vidi kad stampa padne.
Private Sub T_Fak_GreskaNePreziviLogErr()
    Dim n1 As Long, d1 As String, n2 As Long
    Dim n3 As Long, d3 As String

    ' --- 1) opasnost je stvarna: LogErr BRISE Err ---
    On Error Resume Next
    Err.Raise vbObjectError + 9911, "T_Fak", "kontrolna greska"
    n1 = Err.Number
    d1 = Err.description
    AssertEq (n1 <> 0), True, "preduslov: kontrolna greska je podignuta"
    AssertEq (Len(d1) > 0), True, "preduslov: kontrolna greska ima opis"
    LogErr "T_Fak_GreskaNePreziviLogErr"
    n2 = Err.Number
    Err.Clear
    On Error GoTo 0
    AssertEq n2, 0, _
             "LogErr BRISE Err -- zato se greska mora citati PRE njega"

    ' --- 2) pravi put: stampa fakture koje nema ---
    ' RequireSingleFakturaRow digne gresku PRE ijednog upisa, pa je ovo
    ' bezbedno; PrintFaktura je jedan od cetiri EH bloka koji su popravljeni.
    On Error Resume Next
    modFaktura.PrintFaktura "FAK-NE-POSTOJI"
    n3 = Err.Number
    d3 = Err.description
    Err.Clear
    On Error GoTo 0

    AssertEq (n3 <> 0), True, "greska je uopste stigla do pozivaoca"
    ' NAJVAZNIJE: opis prezivljava. Prazan opis znaci da je Err citan POSLE
    ' LogErr-a -- tada operater na ekranu vidi radnju koja 'ne radi', bez razloga.
    AssertEq (Len(d3) > 0), True, _
             "opis greske NIJE prazan -- Err je citan PRE LogErr-a"
    AssertEq (InStr(d3, "FAK-NE-POSTOJI") > 0), True, _
             "opis imenuje fakturu koje nema, ne neku drugu gresku"
End Sub


'=====================================================================
' EKRAN UVOZ IZVODA (Faza E/17, v6-ui-177)
'=====================================================================

' Koliko redova cip propusta na TEKUCOJ listi ekrana. Odvojeno, jer je
' `Scr_Rows(...)(2)` nad Variant-om koji sadrzi niz nepouzdan zapis.
Private Function BuBrojRedova(ByVal filter As String) As Long
    Dim d As Variant
    d = modScrBankaUvoz.Scr_Rows(filter, "")
    If Not IsArray(d) Then Exit Function
    BuBrojRedova = CLng(d(2))
End Function

' UGOVOR EKRANA i GRANICE BAZENA LJUSKE. Visak se ne prijavljuje kao greska
' nego se TIHO odseca -- operater vidi ekran kome fali dugme, bez ijedne
' poruke. Lista stavki stoji TACNO na granici MAX_ACT, pa je ovo jedino mesto
' koje bi sestu radnju primetilo pre nego sto nestane.
Private Sub T_BankaUvoz_UgovorEkrana()
    Dim liste As Variant, i As Long, kljucevi As String
    Dim kljuc As String, spec As String, d As Variant, kolone As Variant
    ' For Each trazi Variant ili Object -- String iterator je "Type mismatch".
    Dim kv As Variant
    Dim redovi As Variant, j As Long, r2 As Long

    AssertEq (Len(modUiScreens.ScrRowByKey("BANKA_UVOZ")) > 0), True, _
             "BANKA_UVOZ postoji u registru ekrana"
    AssertEq modUiScreens.ScrField(modUiScreens.ScrRowByKey("BANKA_UVOZ"), SCR_MODUL), _
             "modScrBankaUvoz", "registar vodi ekran na modScrBankaUvoz"
    AssertEq modUiScreens.ScrPostoji("BANKA_UVOZ"), True, _
             "modul ekrana odgovara na Scr_Meta -- stavka menija vise nije prigusena"
    AssertEq (InStr(modUiScreens.ScrMeta("BANKA_UVOZ"), "kljuc=BANKA_UVOZ") > 0), True, _
             "Scr_Meta prijavljuje svoj kljuc"
    AssertEq modUiScreens.ScrField(modUiScreens.ScrRowByKey("BANKA_UVOZ"), SCR_OBLAST), _
             OBL_BANKA, "ekran trazi pravo na oblast Banka"

    liste = modScrBankaUvoz.Scr_Liste()
    AssertEq (UBound(liste) + 1 <= modOtkupUI.MaxPrekidaca()), True, _
             "ljuska crta sve liste ekrana -- nijedna se ne odseca tiho"
    AssertEq UBound(liste) + 1, 2, "dve liste: stavke i izvodi"
    For i = 0 To UBound(liste)
        kljucevi = kljucevi & "|" & Split(CStr(liste(i)), "|")(0)
    Next i
    AssertEq kljucevi, "|STAVKE|IZVODI", _
             "redosled lista -- stavke su prve, one su posao"

    ' RADNJE PO KLJUCU LISTE, ne kroz Scr_Lista: ugovor svake liste mora da se
    ' meri bez prebacivanja stanja ekrana.
    AssertEq BrojStavkiOpisa(modScrBankaUvoz.BuRadnjeZaListu("STAVKE")), _
             modOtkupUI.MAX_ACT, _
             "stavke nose TACNO MAX_ACT radnji -- sesta bi se tiho odsekla"
    AssertEq BrojStavkiOpisa(modScrBankaUvoz.BuRadnjeZaListu("IZVODI")), 0, _
             "izvodi su pregled -- nijedna radnja nad redom"

    ' Prvi cip je svuda 'sve' -- ljuska na njega pada kad zatecen filter ne
    ' pripada listi (RefreshChipsForScreen), pa prvi mora da bude NAJSIRI;
    ' povratak na uzi cip bi tiho sakrio redove.
    For Each kv In Array("STAVKE", "IZVODI")
        spec = modScrBankaUvoz.BuCipoviZaListu(CStr(kv))
        AssertEq (BrojStavkiOpisa(spec) <= modOtkupUI.MAX_CHIP), True, _
                 "lista " & kv & " ne trazi vise cipova nego sto bazen ima"
        AssertEq Split(Split(spec, "|")(0), ":")(0), "sve", _
                 "prvi cip liste " & kv & " je najsiri ('sve')"
        kolone = modScrBankaUvoz.BuKoloneZaListu(CStr(kv))
        AssertEq (UBound(kolone) + 1 <= modOtkupUI.MAX_COLS), True, _
                 "lista " & kv & " ne trazi vise kolona nego sto mreza pravi"
    Next kv

    ' Svaka lista mora da vrati ispravan niz. Lista koja pukne se u ljusci
    ' pretvara u Empty, LoadGridFromScreen na ne-niz radi Exit Sub -- pa mreza
    ' ostane na prethodnoj listi i prekidac izgleda kao da ne radi.
    For i = 0 To UBound(liste)
        kljuc = Split(CStr(liste(i)), "|")(0)
        modScrBankaUvoz.Scr_BuListaTestSet kljuc
        AssertEq modScrBankaUvoz.Scr_Lista(), kljuc, "lista " & kljuc & " je prihvacena"
        d = modScrBankaUvoz.Scr_Rows("sve", "")
        AssertEq IsArray(d), True, "lista " & kljuc & " vraca niz"
        AssertEq (UBound(d) >= 4), True, _
                 "lista " & kljuc & " vraca pun oblik (kolone, redovi, n, kg, vrednost)"
        AssertEq IsArray(d(0)), True, "lista " & kljuc & " prijavljuje svoje kolone"
        AssertEq (CLng(d(2)) > 0), True, _
                 "lista " & kljuc & " ima redove u fixture-u -- tvrdnje nisu prazne"

        ' DATUM MORA DA STIGNE MREZI KAO SERIJSKI BROJ. Ljuskin FmtDatumKratko
        ' odbija sve sto nije IsNumeric, a IsNumeric je nad Date-om FALSE -- pa
        ' celija ostane PRAZNA, bez ijedne greske. To je naslo tek pustanje nad
        ' pravim podacima; suite nije video jer nijedna tvrdnja nije citala datum.
        kolone = d(0)
        redovi = d(1)
        For j = 0 To UBound(kolone)
            If Split(CStr(kolone(j)), "|")(2) = "date" Then
                ' Svaki red, ne samo prvi: dovoljna je JEDNA losa vrednost da
                ' celija ostane sa natpisom od ranijeg crtanja.
                For r2 = 1 To CLng(d(2))
                    AssertEq IsNumeric(redovi(r2, j + 1)), True, _
                             "lista " & kljuc & ", red " & CStr(r2) & ", kolona " & _
                             CStr(j + 1) & ": datum je BROJ -- inace ga mreza ne crta"
                    ' Opseg (da CDate sme da primi broj) NIJE ovde: to je
                    ' pravilo ljuske i ima svoj test -- T_MrezaDatum_BrojKojiNijeDatum.
                    ' Da se tvrdi i ovde, jedna sabotaza bi obarala dva testa.
                Next r2
            End If
        Next j
    Next i

    modScrBankaUvoz.Scr_BuTestReset
End Sub

' IDENTITET IDE U RED I NE CRTA SE. Mreza redove sortira i deli na strane, pa
' bi svaka mapa "prikaz -> ID" koju ekran drzi sa strane zastarela na prvi klik
' po zaglavlju.
'
' BROJ IZVODA NIJE IDENTITET: dedupe kljuc (IsDuplicateBankaImport) pocinje od
' BROJA RACUNA -- "Drugi racun = druga transakcija, bez obzira na broj izvoda i
' iznos" -- pa dva racuna firme legitimno nose izvod istog broja.
Private Sub T_BankaUvoz_IdentitetURedu_NeCrtaSe()
    Dim kolone As Variant, spec() As String, d As Variant, redovi As Variant
    Dim r As Long, i As Long

    ' Identitet stavke je kolona koju radnja cita kroz GridCell (BU_STV_KOL_ID),
    ' a to NIJE poslednja kolona: iza nje stoje jos dve koje red takodje samo
    ' PRENOSI (otvorenost i smer). Tvrdnja mora da gadja BAS TU kolonu --
    ' tvrdnja o "poslednjoj" meri susedovu i propusta pomeren prioritet
    ' identiteta. (Nadjeno sabotazom banka-uvoz-identitet-vidljiv, koja je nad
    ' prvom verzijom ovog testa prolazila neprimeceno.)
    kolone = modScrBankaUvoz.BuKoloneZaListu("STAVKE")
    spec = Split(CStr(kolone(9)), "|")
    AssertEq spec(0), "OTKUI_HDB_BIMKEY", "deseta kolona stavke je identitet"
    AssertEq spec(4), "4", "identitet stavke je prioriteta 4 -- ne crta se"
    ' Sve tri prenosne kolone moraju ostati van prikaza: mreza crta do 3.
    For i = 9 To UBound(kolone)
        AssertEq Split(CStr(kolone(i)), "|")(4), "4", _
                 "prenosna kolona " & CStr(i + 1) & " se ne crta"
    Next i
    ' INTERNA SIFRA NE IDE OPERATERU PRED OCI. BankaImportID sme da postoji samo
    ' u prenosnoj koloni; medju vidljivima ga nema.
    For i = 0 To 8
        AssertEq (Split(CStr(kolone(i)), "|")(0) <> "OTKUI_HDB_BIMKEY"), True, _
                 "vidljiva kolona " & CStr(i + 1) & " nije interna sifra"
    Next i

    kolone = modScrBankaUvoz.BuKoloneZaListu("IZVODI")
    spec = Split(CStr(kolone(UBound(kolone))), "|")
    AssertEq spec(0), "OTKUI_HDB_IZVKEY", "poslednja kolona izvoda je identitet"
    AssertEq spec(4), "4", "identitet izvoda je prioriteta 4 -- ne crta se"

    modScrBankaUvoz.Scr_BuListaTestSet "STAVKE"
    d = modScrBankaUvoz.Scr_Rows("sve", "")
    redovi = d(1)

    r = RedSaVrednoscu(redovi, 3, FX_BIM_P_JAKI_FAK)
    AssertEq (r > 0), True, "stavka partnera " & FX_BIM_P_JAKI_FAK & " je u listi"
    AssertEq Trim$(CStr(redovi(r, 10))), FX_BIM_JAKI_FAK, _
             "skrivena kolona nosi identitet stavke"
    ' Prva kolona je BROJ IZVODA -- jedini POSLOVNI broj koji stavka nosi.
    AssertEq Trim$(CStr(redovi(r, 1))), FX_BIM_IZVOD1, "prva kolona je broj izvoda"
    AssertEq Trim$(CStr(redovi(r, 9))), FX_BIM_RACUN1, "red nosi broj racuna"

    r = RedSaVrednoscu(redovi, 3, FX_BIM_P_KOLIZIJA)
    AssertEq (r > 0), True, "kolizioni red je u listi"
    AssertEq Trim$(CStr(redovi(r, 1))), FX_BIM_IZVOD1, _
             "kolizioni red nosi ISTI broj izvoda"
    AssertEq Trim$(CStr(redovi(r, 9))), FX_BIM_RACUN2, _
             "ali DRUGI broj racuna -- zato broj izvoda ne moze biti identitet"
    AssertEq Trim$(CStr(redovi(r, 10))), FX_BIM_KOLIZIJA, _
             "identitet je BankaImportID i razlikuje ta dva reda"

    ' DVOSMISLEN ID NOSI PRAZAN IDENTITET, a red se i dalje VIDI. Radnja tada
    ' odbija da bira umesto da pogodi -- bez toga bi svakako pukla
    ' (RequireSingleRow fail-close-uje na duplikat), ali kao greska transakcije
    ' umesto kao poruka operateru.
    r = RedSaVrednoscu(redovi, 3, FX_BIM_P_DUP)
    AssertEq (r > 0), True, "dvosmislena stavka se i dalje VIDI u listi"
    AssertEq Trim$(CStr(redovi(r, 10))), "", _
             "dvosmislen ID nosi PRAZAN identitet -- radnja odbija da bira"

    AssertEq RedSaVrednoscu(redovi, 3, FX_BIM_P_STORNO), 0, _
             "storniran red nije u listi stavki"

    modScrBankaUvoz.Scr_BuTestReset
End Sub

' STO RADNJA MORA DA ZNA A IZ PRIKAZA SE NE VIDI JEDNOZNACNO -- to red NOSI.
'
' Smer: red sa I uplatom I isplatom u mrezi izgleda kao uplata (kolona uplate je
' popunjena), a writer ga odbija. Otvorenost: nov red ima PRAZAN status, pa se u
' prikazu ne razlikuje od reda kome status nije upisan.
Private Sub T_BankaUvoz_RedNosiSmerIOtvorenost()
    Dim d As Variant, redovi As Variant, r As Long

    AssertEq modBankaMapiranje.ClassifyBimSmer(100, 100), BIM_SMER_NEJASAN, _
             "i uplata i isplata = NEJASAN smer, iako izgleda kao uplata"
    AssertEq modBankaMapiranje.ClassifyBimSmer(0, 0), BIM_SMER_NEJASAN, _
             "ni uplata ni isplata = nejasan smer"
    AssertEq modBankaMapiranje.ClassifyBimSmer(100, 0), BIM_SMER_UPLATA, "cista uplata"
    AssertEq modBankaMapiranje.ClassifyBimSmer(0, 100), BIM_SMER_ISPLATA, "cista isplata"

    AssertEq modBankaMapiranje.BimOtvoren(""), True, "nov red je otvoren"
    AssertEq modBankaMapiranje.BimOtvoren(BIM_OBR_ERROR), True, _
             "red oznacen za rucno je JOS UVEK otvoren -- auto ga nije zatvorio"
    AssertEq modBankaMapiranje.BimOtvoren(BIM_OBR_DA), False, "obradjen nije otvoren"
    AssertEq modBankaMapiranje.BimOtvoren(BIM_OBR_SKIP), False, "preskocen nije otvoren"

    modScrBankaUvoz.Scr_BuListaTestSet "STAVKE"
    d = modScrBankaUvoz.Scr_Rows("sve", "")
    redovi = d(1)

    r = RedSaVrednoscu(redovi, 3, FX_BIM_P_JAKI_BLOK)
    AssertEq (r > 0), True, "red isplate je u listi"
    AssertEq Trim$(CStr(redovi(r, 12))), BIM_SMER_ISPLATA, "red NOSI svoj smer"
    AssertEq Trim$(CStr(redovi(r, 11))), "1", "otvoren red NOSI otvorenost"

    r = RedSaVrednoscu(redovi, 3, FX_BIM_P_DA)
    AssertEq (r > 0), True, "obradjen red je u listi pod cipom 'sve'"
    AssertEq Trim$(CStr(redovi(r, 11))), "", _
             "obradjen red NE nosi otvorenost -- radnja ga odbija"

    ' Zatvorena stavka nema sta da predlozi: predlog racuna resolvere, a nad
    ' zatvorenim redom nema sta da se mapira.
    AssertEq modScrBankaUvoz.BuPredlogTekst(BIM_SMER_UPLATA, BIM_CILJ_FAKTURA, _
                                            "2/2026", False), "", _
             "zatvorena stavka nema predlog"
    AssertEq (Len(modScrBankaUvoz.BuPredlogTekst(BIM_SMER_NEJASAN, "", "", True)) > 0), True, _
             "nejasan smer se PRIJAVLJUJE u predlogu, ne cuti"

    modScrBankaUvoz.Scr_BuTestReset
End Sub

' CIP 'JAKI KLJUCEVI' I BROJAC MORAJU DA VIDE ISTI SKUP.
'
' Pravilo "da li bi jak kljuc zavrsio ovaj red" zivi na dva mesta -- u citacu
' mreze i u CountStrongKeyReadyBankaImport (koji stoji u natpisu dugmeta) -- i
' moze da se razidje. Ovo je jedino sto bi to primetilo. Isti oblik kao
' T_Fak_CipoviPrateStatusFakture naspram GetOpenFakture.
Private Sub T_BankaUvoz_CipJakihPratiBrojac()
    Dim nSve As Long, nZa As Long, nJaki As Long, nRucno As Long
    Dim nObr As Long, nPre As Long
    Dim otvorene As Variant

    modScrBankaUvoz.Scr_BuListaTestSet "STAVKE"
    nSve = BuBrojRedova("sve")
    nZa = BuBrojRedova("zaobradu")
    nJaki = BuBrojRedova("jaki")
    nRucno = BuBrojRedova("rucno")
    nObr = BuBrojRedova("obradjeno")
    nPre = BuBrojRedova("preskoceno")

    AssertEq nSve, FX_BIM_SVE, "cip 'sve' vidi sve nestornirane stavke"
    AssertEq nSve, nZa + nObr + nPre, _
             "'sve' je tacno unija tri stanja -- nijedan red ne ispada iz svih cipova"
    AssertEq nZa, FX_BIM_OTVORENIH, "sedam stavki ceka operatera"

    ' PREDUSLOV ZA DOKAZ, ne za ponasanje. Znacka broji OTVORENE, a sabotaza
    ' banka-uvoz-znacka-broji-mapirane joj podmece MAPIRANE. Ako fixture ikad
    ' izjednaci te dve brojke, sabotaza prestaje da obara bilo sta -- suite
    ' ostaje zelena, a dokaz tiho nestane. To se vec desilo: dva nova reda sa
    ' "Da" dala su 6 i 6.
    AssertEq (nZa <> nObr), True, _
             "otvorenih i mapiranih MORA biti razlicito -- inace sabotaza znacke ne meri nista"
    AssertEq nRucno, 1, "tacno jedan red je auto pokusao pa vratio operateru"
    ' 'Za rucno' je PODSKUP otvorenih, ne suprotnost: red sa statusom "Error"
    ' je i dalje otvoren.
    AssertEq (nRucno <= nZa), True, "'za rucno' je podskup otvorenih"
    AssertEq nObr, FX_BIM_OBRADJENIH, _
             "cetiri obradjena reda (jedan + dva dvojnika + prosli ciklus)"

    ' NEUSPEH CITANJA NIJE NULA. Znacka odgovara na "ima li posla"; ako citanje
    ' pukne a vratimo nule, operater dobija "nema posla" umesto "ne znam".
    AssertEq CLng(modScrBankaUvoz.BuKpiPosleGreske(Array(7, 1, 8, 0#, 0#))(0)), 7, _
             "posle greske se zadrzava POSLEDNJA POZNATA brojka"
    AssertEq modScrBankaUvoz.BuKpiNepoznat(Array(7, 1, 8, 0#, 0#)), False, _
             "poslednja poznata brojka JESTE podatak"

    ' A prvi pad u sesiji -- kad poslednje poznate vrednosti nema -- daje
    ' NEPOZNATO, ne nulu. Nula bi kroz BrojacTekst dala praznu znacku, a prazna
    ' znacka u ovom UI-ju znaci "nema sta da ceka".
    AssertEq modScrBankaUvoz.BuKpiNepoznat(modScrBankaUvoz.BuKpiPosleGreske(Empty)), True, _
             "bez ijedne poznate brojke stanje je NEPOZNATO"
    AssertEq (CLng(modScrBankaUvoz.BuKpiPosleGreske(Empty)(0)) < 0), True, _
             "nepoznato se nosi kao negativan broj -- ugovor Scr_Brojac je Long"
    AssertEq modScrBankaUvoz.BuKpiNepoznat(Empty), True, _
             "ni skup koji uopste nije niz nije podatak"
    AssertEq nPre, 1, "tacno jedan preskocen"

    AssertEq (nJaki > 0), True, "fixture ima bar jedan jak kljuc -- tvrdnja nije prazna"
    AssertEq (nJaki <= nZa), True, "jaki kljucevi su podskup otvorenih"
    AssertEq nJaki, modBankaMapiranje.CountStrongKeyReadyBankaImport(), _
             "cip 'jaki kljucevi' i BROJAC vide ISTI skup"

    ' Cip 'za obradu' je tacno ono sto GetBankaImportOpen vraca...
    otvorene = modBankaMapiranje.GetBankaImportOpen()
    AssertEq IsArray(otvorene), True, "GetBankaImportOpen vraca redove"
    AssertEq nZa, UBound(otvorene, 1), _
             "cip 'za obradu' je tacno skup GetBankaImportOpen"

    ' ...i to je isti broj koji nosi znacka uz stavku menija. Znacka NEMA svoj
    ' kanal: red za mapiranje je podatak u tabeli, pa ga ljuska osvezi kroz
    ' RefreshFromData posle svakog upisa.
    modScrBankaUvoz.Scr_ResetCache
    AssertEq modScrBankaUvoz.Scr_Brojac(), nZa, _
             "znacka broji ISTI skup kao cip 'za obradu'"

    modScrBankaUvoz.Scr_BuTestReset
End Sub

' IZVODI SU AGREGAT PO (BROJ IZVODA + BROJ RACUNA), i to je jedino mesto na kom
' se vidi da li se izvod slaze. Legacy je isti racun imao u JEDNOJ labeli i samo
' za najnoviji izvod (UpdateIzvodSummaryLabel).
Private Sub T_BankaUvoz_IzvodiSuAgregatPoRacunu()
    Dim d As Variant, redovi As Variant, i As Long, n As Long
    Dim r1 As Long, r2 As Long, rPY As Long, istihBrojeva As Long
    Dim nes As Long, nesEkran As Long
    Dim okC As Boolean, prometSve As Double
    Dim sirovi As Variant

    ' PRAVILO SLAGANJA, izmereno bez mreze.
    AssertEq modBankaImport.BimSaldoStatus(0, 0, 0, 0), BIM_SALDO_NEMA, _
             "legacy red bez saldo metapodataka NIJE neslaganje nego odsustvo podatka"
    AssertEq modBankaImport.BimSaldoStatus(1000, 1200, 300, 500), BIM_SALDO_OK, _
             "1000 + 500 - 300 = 1200 se slaze"
    AssertEq modBankaImport.BimSaldoStatus(1000, 1300, 300, 500), BIM_SALDO_RAZLIKA, _
             "sto dinara razlike je neslaganje"
    AssertEq modBankaImport.BimSaldoRazlika(1000, 1300, 300, 500), -100, _
             "razlika nosi znak -- zavrsno je VECE od izracunatog"

    ' KLJUC GRUPE, izmeren direktno. Agregat ispod je posledica; da se pravilo
    ' meri samo preko broja redova, obe polovine kljuca bi obarale ISTU tvrdnju
    ' i sabotaza ne bi umela da ih razlikuje.
    AssertEq (modBankaImport.BimIzvodKljuc("15", "111", DateSerial(2026, 3, 16)) <> _
              modBankaImport.BimIzvodKljuc("15", "222", DateSerial(2026, 3, 16))), True, _
             "isti broj izvoda na DVA RACUNA daje dva kljuca"
    AssertEq (modBankaImport.BimIzvodKljuc("15", "111", DateSerial(2026, 3, 16)) <> _
              modBankaImport.BimIzvodKljuc("15", "111", DateSerial(2025, 3, 16))), True, _
             "isti broj i isti racun iz DVA CIKLUSA daju dva kljuca"
    AssertEq (modBankaImport.BimIzvodKljuc("15", "111", CDbl(DateSerial(2026, 3, 16))) = _
              modBankaImport.BimIzvodKljuc("15", "111", DateSerial(2026, 3, 16))), True, _
             "isti dan zapisan kao broj i kao datum je ISTI izvod"

    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    d = modScrBankaUvoz.Scr_Rows("sve", "")
    redovi = d(1)
    n = CLng(d(2))

    For i = 1 To n
        If Trim$(CStr(redovi(i, 1))) = FX_BIM_IZVOD1 Then
            istihBrojeva = istihBrojeva + 1
            ' Isti broj I isti racun postoje DVAPUT -- u dva ciklusa. Red se zato
            ' bira i po datumu; da se bira samo po broju i racunu, tvrdnja bi
            ' merila onaj koji je slucajno poslednji.
            If Trim$(CStr(redovi(i, 2))) = FX_BIM_RACUN1 _
               And CDbl(redovi(i, 3)) = CDbl(DateSerial(2026, 3, 16)) Then r1 = i
            If Trim$(CStr(redovi(i, 2))) = FX_BIM_RACUN2 Then r2 = i
            If Trim$(CStr(redovi(i, 2))) = FX_BIM_RACUN1 _
               And CDbl(redovi(i, 3)) = CDbl(DateSerial(2025, 3, 16)) Then rPY = i
        End If
    Next i

    AssertEq istihBrojeva, 3, _
             "isti broj izvoda daje TRI reda: dva racuna i dva ciklusa"
    ' DRUGA POLOVINA IDENTITETA. Isti broj i isti racun, ali drugi datum, NISU
    ' isti izvod. Da jesu, saldo i datum bi se uzeli sa prvog reda a broj stavki
    ' sabrao preko oba -- sinteticki izvod koji nikad nije postojao.
    AssertEq (rPY > 0), True, "izvod iz proslog ciklusa ima SVOJ red"
    AssertEq (rPY <> r1), True, _
             "isti broj + isti racun + drugi datum su DVA izvoda"
    AssertEq CStr(redovi(rPY, 9)), "0 / 1", _
             "stavke se NE sabiraju preko dva ciklusa"
    AssertEq CDbl(redovi(rPY, 4)), 1000, "saldo se ne uzima sa tudjeg reda"
    AssertEq (r1 > 0 And r2 > 0), True, "oba racuna su u listi"
    AssertEq (Trim$(CStr(redovi(r1, 10))) <> Trim$(CStr(redovi(r2, 10)))), True, _
             "identiteti dva izvoda istog broja su RAZLICITI"

    ' Zbirovi izvoda se NE SABIRAJU po redovima -- parser ih pise na SVAKI red
    ' grupe, pa bi sabiranje dalo iznos pomnozen brojem stavki.
    AssertEq CDbl(redovi(r1, 4)), 10000, "pocetno stanje se uzima sa reda, ne sabira"
    AssertEq CDbl(redovi(r1, 7)), 9500, "zavrsno stanje se uzima sa reda"
    ' Broj otvorenih i broj stavki stoje u JEDNOJ koloni -- v. IzvodiKolone.
    AssertEq CStr(redovi(r1, 9)), "3 / 3", _
             "izvod 1 / racun 1 ima tri stavke i sve tri su otvorene"
    AssertEq modScrBankaUvoz.BuStavkiTekst(10, 16), "10 / 16", _
             "zapis je isti kao u traci iznad mreze"

    ' Storniran red ne ulazi ni u grupu ni u brojace: izvod 2 ima pet redova, a
    ' jedan od njih je storniran.
    For i = 1 To n
        If Trim$(CStr(redovi(i, 1))) = FX_BIM_IZVOD2 Then
            AssertEq CStr(redovi(i, 9)), "2 / 4", _
                     "storniran red se ne broji u stavke izvoda"
        End If
    Next i

    ' NESAGLASAN IZVOD: redovi istog izvoda nose razlicite zbirove.
    '
    ' Agregat brojke UZIMA sa prvog reda umesto da ih sabira (sabiranje bi ih
    ' pomnozilo brojem stavki) -- a to vazi samo dok su svi redovi saglasni.
    ' Kad nisu, brojka prvog reda nije istina o izvodu, i to se mora reci.
    ' STATUS se cita iz CITACA, ne iz reda ekrana: red ekrana nosi TEKST kolone
    ' "Slaganje" (v. IzvodiKolone), a deseta kolona mu je identitet. Ovde se meri
    ' odluka, a nize i to da se ta odluka vidi u prikazu.
    sirovi = modBankaImport.GetBankaIzvodiForGrid()
    nes = 0
    For i = 1 To UBound(sirovi, 1)
        If Trim$(CStr(sirovi(i, 2))) = FX_BIM_IZVOD_NES Then nes = i
    Next i
    AssertEq (nes > 0), True, "preduslov: nesaglasan izvod je u citacu"
    AssertEq CLng(sirovi(nes, 10)), BIM_SALDO_NEKONZISTENTAN, _
             "izvod cija se dva reda razlikuju je NESAGLASAN"
    AssertEq CLng(sirovi(nes, 11)), 2, "obe stavke su u istoj grupi"
    ' NESAGLASNO NADJACAVA i "slaze se" i "ne slaze se". Prvi red ovog izvoda
    ' sam za sebe DAJE slaganje (4500 + 500 - 0 = 5000), pa bi bez pravila stajalo
    ' "slaze se" -- najgori moguci ishod, jer tvrdi tacnost o brojkama kojih nema.
    AssertEq modBankaImport.BimSaldoStatus(4500, 5000, 0, 500), BIM_SALDO_OK, _
             "kontrola: prvi red sam za sebe se SLAZE"

    ' Pravilo poredjenja, izmereno direktno.
    AssertEq modBankaImport.BimSaldoIsti(1, 2, 3, 4, 1, 2, 3, 4), True, _
             "isti zbirovi su saglasni"
    AssertEq modBankaImport.BimSaldoIsti(1, 2, 3, 4, 1, 2, 3, 4.02), False, _
             "razlika veca od centa je nesaglasnost"
    AssertEq modBankaImport.BimSaldoIsti(1, 2, 3, 4, 1, 2, 3, 4.005), True, _
             "polovina centa nije -- prag je isti kao kod slaganja"

    ' Cip "ne slaze se" NE broji nesaglasne: on nosi jedno tvrdjenje, a o
    ' nesaglasnom izvodu se ne zna nista. Da ih broji, brojka bi bila
    ' neupotrebljiva za oba stanja.
    AssertEq modScrBankaUvoz.BuCipIzvod("razlika", BIM_SALDO_NEKONZISTENTAN, 0), False, _
             "nesaglasan izvod nije 'ne slaze se'"
    AssertEq modScrBankaUvoz.BuCipIzvod("sve", BIM_SALDO_NEKONZISTENTAN, 0), True, _
             "...ali se vidi u 'sve'"
    ' Poredi se sa KATALOGOM, ne sa samom funkcijom: tvrdnja koja obe strane
    ' racuna istom funkcijom prolazi i kad funkcija vrati pogresnu poruku.
    AssertEq modScrBankaUvoz.BuSlaganjeTekst(BIM_SALDO_NEKONZISTENTAN, 0), _
             Poruka("OTKUI_LBL_BU_SALDO_NESAGLASAN"), _
             "nesaglasan izvod dobija SVOJU poruku"
    AssertEq (InStr(1, modScrBankaUvoz.BuSlaganjeTekst(BIM_SALDO_NEKONZISTENTAN, 0), _
                    Poruka("OTKUI_LBL_BU_SALDO_RAZLIKA")) = 0), True, _
             "...i u njoj NE stoji 'ne slaze se'"

    ' I to stvarno stigne u red mreze -- osma kolona je SLAGANJE (IzvodiKolone).
    nesEkran = 0
    For i = 1 To n
        If Trim$(CStr(redovi(i, 1))) = FX_BIM_IZVOD_NES Then nesEkran = i
    Next i
    AssertEq (nesEkran > 0), True, "nesaglasan izvod je i u listi ekrana"
    AssertEq CStr(redovi(nesEkran, 8)), Poruka("OTKUI_LBL_BU_SALDO_NESAGLASAN"), _
             "kolona Slaganje nosi bas taj tekst"

    ' STATUS I BROJKE MORAJU DA SE SLAZU. Reci "ne zna se koji zbirovi vaze" a
    ' pored toga prikazati brojku PRVOG reda znaci ponuditi tudji podatak kao
    ' saldo -- ista klasa kao natpis prethodnog ekrana u koloni datuma.
    ' Kolone su "rest", pa nula znaci PRAZNA celija.
    AssertEq CDbl(redovi(nesEkran, 4)), 0, "nesaglasan izvod nema pocetno stanje"
    AssertEq CDbl(redovi(nesEkran, 5)), 0, "...ni uplate"
    AssertEq CDbl(redovi(nesEkran, 6)), 0, "...ni isplate"
    AssertEq CDbl(redovi(nesEkran, 7)), 0, "...ni zavrsno stanje"
    AssertEq modOtkupUI.CelijaTekst("rest", 0, okC), "", _
             "...a nula u 'rest' koloni je PRAZNO, ne 0,00"

    ' PODNOZJE. Nesaglasan izvod ne sme da ulazi u promet: njegov UkupanPotrazuje
    ' je 500 na prvom redu i 700 na drugom, pa bi zbir tvrdio promet koji nikad
    ' nije izracunat.
    d = modScrBankaUvoz.Scr_Rows("sve", "")
    prometSve = CDbl(d(4))
    AssertEq (prometSve > 0), True, "preduslov: podnozje uopste racuna promet"

    ' PODNOZJE MORA I DA SE VIDI, ne samo da se izracuna.
    '
    ' Scr_Rows vraca zbir, ali ljuska odlucuje hoce li ga NACRTATI --
    ' ModeHasValCol gleda vrste kolona. Kad su novcane kolone izvoda presle na
    ' "rest", ta odluka je postala False i podnozje se sakrilo, dok je zbir bio
    ' savrseno tacan i test zelen. Zato se tvrdi i ODLUKA LJUSKE.
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    modOtkupUI.GridTestLoad "BANKA_UVOZ"
    AssertEq modOtkupUI.GridImaValKolonuTest(), True, _
             "ljuska za listu IZVODA crta zbir vrednosti u podnozju"

    modScrBankaUvoz.Scr_BuListaTestSet "STAVKE"
    modOtkupUI.GridTestLoad "BANKA_UVOZ"
    AssertEq modOtkupUI.GridImaValKolonuTest(), True, _
             "...i za listu STAVKI"

    modOtkupUI.GridTestLoad ""
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"

    ' PRETRAGA I PODNOZJE ide PRVO, i to namerno (zamka 5): zbir koji ne postuje
    ' pretragu obara i tvrdnju o nesaglasnom izvodu ispod, pa bi dve sabotaze
    ' padale na istoj tvrdnji i ne bi se razlikovale. Meri se na izvodu koji
    ' JESTE saglasan, pa mu promet nije nula.
    d = modScrBankaUvoz.Scr_Rows("sve", FX_BIM_IZVOD2)
    AssertEq (CDbl(d(4)) > 0), True, "preduslov: izolovan izvod ima svoj promet"
    AssertEq (CDbl(d(4)) < prometSve), True, _
             "pretraga smanjuje i PROMET, ne samo broj redova"

    ' SAM NESAGLASAN IZVOD, izolovan pretragom. Tvrdnja je ostra: njegov promet
    ' mora biti TACNO nula. Poredjenje "manje od ukupnog" ne bi merilo nista --
    ' i da ulazi u zbir, ukupno bi samo bilo vece.
    d = modScrBankaUvoz.Scr_Rows("sve", FX_BIM_IZVOD_NES)
    AssertEq CLng(d(2)), 1, "preduslov: pretraga izoluje bas taj izvod"
    AssertEq CDbl(d(4)), 0, _
             "nesaglasan izvod ne donosi NISTA u promet -- ne zna se koji zbirovi vaze"

    AssertEq BuBrojRedova("razlika"), 1, "tacno jedan izvod se ne slaze"
    AssertEq BuBrojRedova("sve"), 6, _
             "sest grupa: dva racuna i dva ciklusa pod istim brojem, pa jos tri"
    ' Izvod 3 je ceo obradjen, pa cip "sa otvorenim" ima sta da iskljuci --
    ' inace bi propustao sve i bio prazna tvrdnja.
    AssertEq BuBrojRedova("otvoreni"), 4, "cetiri izvoda jos imaju otvorenih stavki"

    modScrBankaUvoz.Scr_BuTestReset
End Sub

' RUCNO MAPIRANJE -- pravila koja su do sada zivela Private u frmBankaImport.
'
' Najskuplje od njih je FAIL-CLOSED nad listom faktura: prazna lista i PAD
' ucitavanja izgledaju isto, a znace suprotno -- prazan izbor fakture knjizi
' AVANS umesto zatvaranja duga.
Private Sub T_BankaUvoz_RucnoMapiranjePravila()
    Dim ok As Boolean, greska As String, razlog As String
    Dim src As Variant, i As Long
    Dim nasao As Boolean
    Dim omBlokova As Long, omBezStanice As Long
    Dim placenUListi As Boolean
    Dim errBezKolone As Long, errBezScope As Long
    Dim errNemaTabele As Long, errImaTabele As Long
    Dim punjenoPre As Long

    ' SMER-KAPIJA PRE KLIKA -- ista koju RequireBimSmer sprovodi u writeru
    ' (Kupac -> UPLATA, Kooperant -> ISPLATA, OM -> bilo koji CIST smer).
    AssertEq modBankaMapiranje.BimSmerOdgovaraTipu(BIM_SMER_UPLATA, BIM_TIP_KUPAC), True, _
             "kupac prima UPLATU"
    AssertEq modBankaMapiranje.BimSmerOdgovaraTipu(BIM_SMER_ISPLATA, BIM_TIP_KUPAC), False, _
             "kupac NE prima isplatu"
    AssertEq modBankaMapiranje.BimSmerOdgovaraTipu(BIM_SMER_ISPLATA, BIM_TIP_KOOPERANT), True, _
             "kooperant prima ISPLATU"
    AssertEq modBankaMapiranje.BimSmerOdgovaraTipu(BIM_SMER_UPLATA, BIM_TIP_KOOPERANT), False, _
             "kooperant NE prima uplatu"
    AssertEq modBankaMapiranje.BimSmerOdgovaraTipu(BIM_SMER_UPLATA, BIM_TIP_OM), True, _
             "OM prima uplatu"
    AssertEq modBankaMapiranje.BimSmerOdgovaraTipu(BIM_SMER_ISPLATA, BIM_TIP_OM), True, _
             "OM prima i isplatu"
    AssertEq modBankaMapiranje.BimSmerOdgovaraTipu(BIM_SMER_NEJASAN, BIM_TIP_OM), False, _
             "nejasan smer ne prolazi ni za OM"

    ' NERAZRESEN UNOS NIJE TIP. Ljuska Change salje na svaki otkucaj.
    AssertEq modScrBankaUvoz.BuTipIliPrazno("Koo"), "", _
             "poluukucan tip nije tip"
    AssertEq modScrBankaUvoz.BuTipIliPrazno(BIM_TIP_KOOPERANT), BIM_TIP_KOOPERANT, _
             "razresen tip prolazi"

    ' EFEKTIVNI BLOK: prazan izbor NIJE "nema bloka" nego "uzmi poziv na broj iz
    ' izvoda". U formi je prazan combo bio DEFAULT slucaj, pa je blok sa 3+
    ' stavki bez ovog pravila zavrsavao generickom greskom.
    AssertEq modBankaMapiranje.BimEfektivniBlok(FX_BIM_JAKI_BLOK, ""), FX_BIM_BLOK1_BR, _
             "prazan izbor uzima poziv na broj iz izvoda"
    AssertEq modBankaMapiranje.BimEfektivniBlok(FX_BIM_JAKI_BLOK, "BLOK-RUCNO"), "BLOK-RUCNO", _
             "izbor operatera pobedjuje poziv na broj"

    ' BLOK PREKO GRANICE trazi izricitu potvrdu podele; blok u granicama ne pita.
    AssertEq modBankaMapiranje.BimBlokTraziPotvrdu(FX_KOOPERANT3, FX_BIM_BLOK3_BR, razlog), _
             True, "blok sa tri otvorene stavke trazi potvrdu podele"
    AssertEq (Len(razlog) > 0), True, "razlog imenuje blok -- operater vidi zasto"
    AssertEq modBankaMapiranje.BimBlokTraziPotvrdu(FX_KOOPERANT, FX_BIM_BLOK1_BR, razlog), _
             False, "blok sa jednom otvorenom stavkom ne pita nista"
    AssertEq razlog, "", "kad se ne pita, razloga nema"

    ' PODELA se racuna ISTIM planerom po kome se knjizi, pa operater pre klika
    ' vidi TACNO onu podelu koja ce biti proknjizena.
    src = modBankaMapiranje.GetOtkupCandidatesForKooperantBlock( _
              FX_KOOPERANT3, FX_BIM_BLOK3_BR, True)
    AssertEq IsArray(src), True, "blok preko granice ipak vraca kandidate uz allowOverMax"
    AssertEq UBound(src, 1), 3, "blok ima tri otvorena kandidata"
    AssertEq (Len(modScrBankaUvoz.TekstPodele(src, 3000)) > 0), True, _
             "predlog podele nije prazan"

    ' FAKTURE ZA RUCNO MAPIRANJE: samo one sa otvorenim saldom, a "otvoreno" u
    ' listi mora da bude ISTO ono koje writer racuna -- prikaz i knjizenje jedan
    ' izvor.
    '
    ' Kupac se bira PAZLJIVO: FX_KUPAC do ovog testa vise nema nijednu otvorenu
    ' fakturu, jer je raniji test (uplata na fakturu) zatvorio FX_FAKTURA u
    ' celosti. Tvrdnja vezana za njega bi merila posledicu REDOSLEDA testova,
    ' ne pravilo. FX_KUPAC2 i njegova FAK-TEST-N se ne diraju ni u jednom testu.
    src = modBankaMapiranje.GetFaktureZaBimMapiranje(FX_KUPAC2, ok, greska)
    AssertEq ok, True, "citanje faktura je uspelo"
    AssertEq greska, "", "uspesno citanje ne prijavljuje gresku"
    AssertEq IsArray(src), True, "kupac sa neplacenom fakturom je dobija u listi"
    AssertEq (RedSaVrednoscu(src, 1, FX_FAK_NEPL) > 0), True, _
             "neplacena faktura je ponudjena za rucno mapiranje"
    AssertEq RedSaVrednoscu(src, 1, FX_FAKTURA), 0, _
             "faktura DRUGOG kupca nije u listi"
    AssertEq RedSaVrednoscu(src, 1, FX_FAK_STORNO), 0, _
             "stornirana faktura nije u listi za mapiranje"
    For i = 1 To UBound(src, 1)
        AssertEq CDbl(src(i, 3)), _
                 modBankaMapiranje.GetOtvorenoNaFakturi(CStr(src(i, 1))), _
                 "otvoreno u listi = otvoreno koje racuna writer (" & CStr(src(i, 1)) & ")"
    Next i

    ' ZATVORENA FAKTURA NE ULAZI. FAK-TEST-P je placena u samom fixture-u
    ' (jedina uplata koja nosi FakturaID), pa ovo ne zavisi od redosleda.
    AssertEq modBankaMapiranje.GetOtvorenoNaFakturi(FX_FAK_PLAC), 0, _
             "preduslov: placena faktura nema otvoreno"
    src = modBankaMapiranje.GetFaktureZaBimMapiranje(FX_KUPAC, ok, greska)
    AssertEq ok, True, "citanje faktura drugog kupca je proslo"
    AssertEq RedSaVrednoscu(src, 1, FX_FAK_PLAC), 0, _
             "placena faktura nije u listi za mapiranje"
    If IsArray(src) Then
        For i = 1 To UBound(src, 1)
            AssertEq (CDbl(src(i, 3)) > 0), True, _
                     "svaka ponudjena faktura ima otvoreno (" & CStr(src(i, 1)) & ")"
        Next i
    End If

    ' PRAZNA LISTA UZ USPESNO CITANJE je druga polovina fail-closed pravila:
    ' prazno sme da znaci "nema faktura" SAMO kad je citanje proslo.
    src = modBankaMapiranje.GetFaktureZaBimMapiranje("KUP-NE-POSTOJI", ok, greska)
    AssertEq ok, True, "citanje je proslo i za kupca kog nema"
    AssertEq IsArray(src), False, "kupac bez faktura dobija praznu listu, bez greske"

    ' NEDOSTAJUCA TABELA NIJE PRAZNA TABELA. GetTableData vraca Empty za oba, pa
    ' bi citac koji gleda samo IsEmpty tumacio kvar kao "kupac nema faktura" --
    ' a prazan izbor fakture znaci AVANS. RequireColumnIndex ovo ne pokriva: do
    ' provere kolona se ne bi ni stiglo.
    On Error Resume Next
    Err.Clear
    modSchemaGuard.RequireTable "tblNePostojiNikako", "T_BankaUvoz"
    errNemaTabele = Err.Number
    Err.Clear
    modSchemaGuard.RequireTable TBL_FAKTURE, "T_BankaUvoz"
    errImaTabele = Err.Number
    Err.Clear
    On Error GoTo 0

    AssertEq (errNemaTabele <> 0), True, _
             "nedostajuca tabela PUCA -- ne prolazi kao prazna lista"
    AssertEq errImaTabele, 0, "postojeca tabela prolazi"

    ' FAIL-CLOSED. Pad ucitavanja se u testu ne moze izazvati bez lomljenja
    ' seme, pa se meri kroz seam -- ali seam ide kroz ISTU kapiju (CiljUcitan)
    ' kroz koju idu i obe rucne rute.
    '
    ' Kapija je ZAJEDNICKA namerno. Prvo je stajala samo kod kupca, pa je pad
    ' punjenja liste blokova ostajao neprimecen: prazan combo je izgledao kao
    ' "operater nije birao blok", odatle fallback na poziv na broj sa PRAZNIM
    ' scope-om, a ako kandidata nema -- ceo iznos se knjizi kao avans kooperanta
    ' i stavka se oznacava obradjenom. Kvar bi postao uspesno knjizenje drugog
    ' poslovnog ishoda.
    AssertEq modScrBankaUvoz.Scr_BuCiljStanjeTest(False), False, _
             "pad ucitavanja ZAUSTAVLJA rucno mapiranje -- prazan izbor bi bio avans ili poziv na broj"
    AssertEq modScrBankaUvoz.Scr_BuCiljStanjeTest(True), True, _
             "uredno procitana lista pusta mapiranje"

    ' NEUSPEH SE NE PAMTI. Dve grane greske javljaju razlicito: blokovi je dizu
    ' (pa EH obrise kes), a fakture je vracaju kroz zastavicu i punjenje mirno
    ' stigne do kraja. Da se i tada kesira, radnja bi ostala tacno blokirana --
    ' ali sledeci klik ne bi ni pokusao ponovo, pa bi izbor ostao zakljucan.
    AssertEq modScrBankaUvoz.Scr_BuCiljKesTest("Kupac|K1", True), "Kupac|K1", _
             "uspesno punjenje se pamti"
    AssertEq modScrBankaUvoz.Scr_BuCiljKesTest("Kupac|K1", False), "", _
             "neuspelo punjenje se NE pamti -- sledeci klik pokusava ponovo"

    ' KAPIJA MORA DA PUNI LISTU pre nego sto presudi -- inace zastavica opisuje
    ' PRETHODNI izbor, a odluka se donosi nad ovim. Bez forme se to ne vidi ni
    ' po cemu drugom (PuniCiljCombo bez kontrole izlazi odmah), pa se meri
    ' brojacem poziva.
    punjenoPre = modScrBankaUvoz.Scr_BuCiljPunjenoTest()
    modScrBankaUvoz.Scr_BuCiljStanjeTest True
    AssertEq (modScrBankaUvoz.Scr_BuCiljPunjenoTest() > punjenoPre), True, _
             "kapija PUNI listu cilja pre nego sto presudi"

    ' SCOPE SE NE SME TIHO IZGUBITI KAD KOLONE NEMA.
    ' Ovo je najtisi moguci kvar: zadat scope, kolona nedokaziva, filtriranje
    ' otpada, i pozivalac dobija kandidate sa SVIH otkupnih mesta u listi koja
    ' izgleda savrseno ispravno. Zato "ne mogu da dokazem scope" mora da bude
    ' greska, a ne tihi nastavak.
    On Error Resume Next
    Err.Clear
    modBankaMapiranje.BimScopeKolona FX_STANICA, "NemaOvakveKoloneUOtkupu"
    errBezKolone = Err.Number
    Err.Clear
    ' A kad scope NIJE zadat, ista nedokaziva kolona je legitimna: automatsko
    ' mapiranje otkupno mesto nema odakle da zna i radi bez njega, kao i pre.
    modBankaMapiranje.BimScopeKolona "", "NemaOvakveKoloneUOtkupu"
    errBezScope = Err.Number
    Err.Clear
    On Error GoTo 0

    AssertEq (errBezKolone <> 0), True, _
             "zadat scope nad nedokazivom kolonom PUCA -- ne vraca nescope-ovane kandidate"
    AssertEq errBezScope, 0, _
             "bez zadatog scope-a ista kolona ostaje opciona"
    AssertEq modBankaMapiranje.BimScopeKolona(FX_STANICA, COL_OTK_STANICA) > 0, True, _
             "nad zdravom semom scope kolona ima indeks"

    ' BLOKOVI kooperanta. Kljuc je (broj + OTKUPNO MESTO), ne samo broj: broj
    ' otkupa je jedinstven po stanici, pa isti broj pripada dvama razlicitim
    ' blokovima. Ko ponudi samo broj, posle izbora ne zna KOJI je -- a od toga
    ' zavisi na koji otkupni lanac ide novac.
    src = modBankaMapiranje.GetBlokoviZaBimMapiranje(FX_KOOPERANT3)
    AssertEq IsArray(src), True, "kooperant ima blokove"
    For i = 1 To UBound(src, 1)
        If CStr(src(i, 1)) = FX_BIM_BLOK3_BR Then nasao = True
        If CStr(src(i, 1)) = FX_BIM_BLOK_PLACEN Then placenUListi = True
        If CStr(src(i, 1)) = FX_BIM_BLOK_OM Then
            omBlokova = omBlokova + 1
            If Len(Trim$(CStr(src(i, 2)))) = 0 Then omBezStanice = omBezStanice + 1
        End If
    Next i
    AssertEq nasao, True, "blok sa tri stavke je u listi"
    AssertEq UBound(src, 1), 5, _
             "tri stavke istog bloka daju jedan red; blok na tri mesta TRI; placen blok JOS jedan"
    AssertEq omBlokova, 3, "isti broj bloka na tri otkupna mesta daje TRI reda"
    ' Red BEZ otkupnog mesta se NE precutkuje -- postoji u podacima, pa se nudi;
    ' ono sto se menja je da radnja nad njim STAJE (v. nize).
    AssertEq omBezStanice, 1, "blok bez otkupnog mesta ostaje u listi"

    ' SCOPE STVARNO SUZAVA. Bez njega su kandidati sva tri otkupna mesta -- a to
    ' je novac na tri razlicita poslovna lanca u JEDNOJ raspodeli.
    src = modBankaMapiranje.GetOtkupCandidatesForKooperantBlock( _
              FX_KOOPERANT3, FX_BIM_BLOK_OM, True)
    AssertEq UBound(src, 1), 3, "bez scope-a ulaze kandidati sa SVIH otkupnih mesta"

    src = modBankaMapiranje.GetOtkupCandidatesForKooperantBlock( _
              FX_KOOPERANT3, FX_BIM_BLOK_OM, True, FX_STANICA)
    AssertEq UBound(src, 1), 1, "sa scope-om ulazi samo jedno otkupno mesto"
    AssertEq CStr(src(1, 1)), FX_OTK_OM_A, "i to bas ono koje je izabrano"

    ' Kontrola u drugom smeru: druga stanica daje DRUGI otkup, ne prazno.
    src = modBankaMapiranje.GetOtkupCandidatesForKooperantBlock( _
              FX_KOOPERANT3, FX_BIM_BLOK_OM, True, FX_STANICA_B)
    AssertEq IsArray(src), True, "druga stanica ima svoje kandidate"
    AssertEq CStr(src(1, 1)), FX_OTK_OM_B, "scope B nikad ne vraca otkup iz scope-a A"

    ' IZABRAN BLOK BEZ OTVORENIH STAVKI -- STOP, ne tihi AVANS.
    '
    ' Lista blokova nudi SVAKI nestorniran broj otkupa i ne proverava da li blok
    ' jos duguje; kandidati se biraju samo ako je "otvoreno > 0.009". Placen blok
    ' zato legitimno stoji u listi a daje NULA kandidata -- a writer na
    ' IsEmpty(kandidati) ceo iznos knjizi kao avans kooperanta i stavku oznacava
    ' obradjenom. Operater je rekao KOJI dug placa; tiha promena u avans je druga
    ' finansijska semantika od one koju je izabrao.
    AssertEq modBankaMapiranje.BimBlokBezOtvorenih(FX_KOOPERANT3, FX_BIM_BLOK_PLACEN), _
             True, "potpuno placen blok NEMA otvorenih stavki"
    AssertEq modBankaMapiranje.BimBlokBezOtvorenih(FX_KOOPERANT3, FX_BIM_BLOK3_BR), _
             False, "blok sa tri otvorene stavke ima sta da plati"

    ' Blok JE u listi -- ne precutkuje se, jer postoji u podacima. Ono sto se
    ' menja je da radnja nad njim staje.
    AssertEq placenUListi, True, "placen blok je i dalje u listi blokova"

    ' KAPIJA VAZI SAMO ZA RUCNI IZBOR. Kad blok dolazi iz poziva na broj,
    ' izabranBlok je prazan i avans i dalje JESTE namerno ponasanje -- to je
    ' bezbedan izlaz dok je poreklo dvosmisleno.
    AssertEq modScrBankaUvoz.BuBlokZatvoren(FX_KOOPERANT3, FX_BIM_BLOK_PLACEN, _
                                            FX_BIM_BLOK_PLACEN, ""), True, _
             "rucno izabran placen blok ZAUSTAVLJA knjizenje"
    AssertEq modScrBankaUvoz.BuBlokZatvoren(FX_KOOPERANT3, "", _
                                            FX_BIM_BLOK_PLACEN, ""), False, _
             "isti blok iz POZIVA NA BROJ ne prolazi kroz kapiju -- avans ostaje namerno ponasanje"
    AssertEq modScrBankaUvoz.BuBlokZatvoren(FX_KOOPERANT3, FX_BIM_BLOK3_BR, _
                                            FX_BIM_BLOK3_BR, ""), False, _
             "blok sa otvorenim stavkama prolazi"

    ' IZABRAN BLOK BEZ OTKUPNOG MESTA -- STOP, ne nescope-ovan upis.
    ' Ovo je najvaznija tvrdnja ovog dela: prazan scope izgleda isto kao "scope
    ' nije ni trazen", a znaci nesto sasvim drugo. Da se prazan propusti, writer
    ' bi raspodelio novac preko sva tri otkupna mesta sa istim brojem bloka.
    modScrBankaUvoz.Scr_BuIzborTestSet BIM_TIP_KOOPERANT, FX_KOOPERANT3, _
                                       FX_BIM_BLOK_OM, ""
    AssertEq modScrBankaUvoz.Scr_BuStopBezOmTest(), True, _
             "izabran blok bez otkupnog mesta ZAUSTAVLJA rucno mapiranje"

    ' PRAZAN IZBOR BLOKA NEMA SCOPE, i to NIJE isti slucaj. Blok tada dolazi iz
    ' poziva na broj, koji otkupno mesto ne nosi -- pa se ekran ponasa kao
    ' automatsko mapiranje i radnja se NE zaustavlja.
    modScrBankaUvoz.Scr_BuIzborTestSet BIM_TIP_KOOPERANT, FX_KOOPERANT3, "", ""
    AssertEq modScrBankaUvoz.Scr_BuScopeBlokaTest(), "", _
             "bez izabranog bloka nema ni scope-a"
    AssertEq modScrBankaUvoz.Scr_BuStopBezOmTest(), False, _
             "poziv na broj nije 'blok bez otkupnog mesta' -- radnja ide dalje"

    modScrBankaUvoz.Scr_BuIzborTestSet BIM_TIP_KOOPERANT, FX_KOOPERANT3, _
                                       FX_BIM_BLOK_OM, FX_STANICA
    AssertEq modScrBankaUvoz.Scr_BuScopeBlokaTest(), FX_STANICA, _
             "izabran blok nosi svoje otkupno mesto do writera"
    AssertEq modScrBankaUvoz.Scr_BuStopBezOmTest(), False, _
             "blok sa otkupnim mestom prolazi"

    modScrBankaUvoz.Scr_BuTestReset
End Sub


' ZONA SE STVARNO GRADI I RASPOREDJUJE.
'
' Ovaj test postoji zbog jednog compile kvara koji nijedan drugi nije mogao da
' vidi: RasporediPolja je koristila GAP, koja je u modOtkupUI PRIVATE. VBA takvu
' gresku prijavljuje tek kad se procedura PRVI PUT IZVRSI -- a nijedan test do
' tada nije crtao zonu ovog ekrana, pa je suite bila zelena, a Excel je na uvozu
' javio "Variable not defined". Sve ostale tvrdnje o ekranu rade nad citacima i
' pravilima, gde zone nema.
'
' Nalazi se SKUPLJAJU, a tvrde tek posle Unload-a: dok forma zivi, njena
' masinerija obrise Err izmedju Err.Raise i omotnice testa, pa bi pad stigao kao
' "greska bez opisa".
Private Sub T_ZonaBankaUvoz_PoljaIRaspored()
    Dim f As frmOtkupUI, z As Object, nm As Variant
    Dim nema As String, koopNema As String, omVisak As String
    Dim visina As Single

    Set f = NewOtkupUIForm()
    Set z = f.Controls.Add("Forms.Frame.1", "zProbaBu", True)
    z.width = 1200: z.Height = 300
    modScrBankaUvoz.Scr_Build z

    For Each nm In Array("buBg", "buCap", "buHint", "buLnB", _
                         "buKL0", "buKV0", "buKL3", "buKV3", _
                         "scrBuTip", "scrBuPartner", "scrBuCilj")
        If Not KontrolaPostoji(z, CStr(nm)) Then nema = nema & " " & CStr(nm)
    Next nm

    ' Kombo u zoni MORA biti polje (okvir nm + kontrola nmT): panel za izbor
    ' (modOtkupUI.FindCombo) trazi bas taj oblik. Gola kontrola bi imala
    ' strelicu koja "ne radi" i listu koja se ne otvara.
    For Each nm In Array("scrBuTip", "scrBuPartner", "scrBuCilj")
        If KontrolaPostoji(z, CStr(nm)) Then
            If Not KontrolaPostoji(z.Controls(CStr(nm)), CStr(nm) & "T") Then _
                nema = nema & " " & CStr(nm) & "T"
        End If
    Next nm

    ' KOOPERANT: cilj je blok otkupa, pa sva tri polja rade.
    modScrBankaUvoz.Scr_BuIzborTestSet BIM_TIP_KOOPERANT, "", ""
    visina = modScrBankaUvoz.Scr_Layout(z, 1200, 300)
    For Each nm In Array("scrBuTip", "scrBuPartner", "scrBuCilj")
        If Not VidljivaKontrola(z, CStr(nm)) Then koopNema = koopNema & " " & CStr(nm)
    Next nm

    ' OM: ni faktura ni blok se ne biraju, pa se polje cilja GASI. Polje koje ne
    ' radi nista poziva da se u njega nesto upise.
    modScrBankaUvoz.Scr_BuIzborTestSet BIM_TIP_OM, "", ""
    modScrBankaUvoz.Scr_Layout z, 1200, 300
    If VidljivaKontrola(z, "scrBuCilj") Then omVisak = "scrBuCilj"

    modScrBankaUvoz.Scr_BuTestReset
    Unload f

    AssertEq nema, "", "zona uvoza izvoda nema nijednu kontrolu manje"
    AssertEq (visina > 0), True, "Scr_Layout prijavljuje visinu zone"
    AssertEq koopNema, "", "za kooperanta su upaljena sva tri polja"
    AssertEq omVisak, "", "za OM je polje cilja UGASENO"
End Sub


' ============================================================
' 111. Broj koji NIJE datum ne sme u kolonu datuma
' ============================================================
' Mreza nad kolonom tipa "date" radi CDate, a CDate van opsega baca Overflow.
' RenderGrid radi pod "On Error Resume Next", pa upis celije bude PRESKOCEN i u
' njoj ostane natpis od RANIJEG crtanja -- operater vidi tudji tekst, bez ijedne
' greske i bez traga u logu.
'
' Nadjeno merenjem nad pravom sveskom (Diag_BuRedovi): tblBankaImport ume da
' nosi DatumTransakcije kao BROJ oblika ddmmyyyy. Isti podatak je, posejan u
' fixture, obarao SEDAM testova sa "Overflow" -- dakle ne pogadja jedan ekran
' nego citavu mrezu.
Private Sub T_MrezaDatum_BrojKojiNijeDatum()
    Dim a(1 To 1, 1 To 1) As Variant

    ' PRAVILO, izmereno bez mreze.
    AssertEq modUiData.DatumSerijskiValidan(26062026#), False, _
             "ddmmyyyy kao broj NIJE datum"
    AssertEq modUiData.DatumSerijskiValidan(0), False, "nula nije datum"
    AssertEq modUiData.DatumSerijskiValidan(modUiData.DATUM_SERIJSKI_MAX + 1), False, _
             "preko 31.12.9999 CDate baca Overflow"
    AssertEq modUiData.DatumSerijskiValidan(CDbl(DateSerial(2026, 8, 21))), True, _
             "stvaran datum prolazi"

    ' CITAC ODBIJA, ne prosledjuje. Da prosledi, mreza bi pukla pri crtanju --
    ' tiho, jer RenderGrid gresku guta.
    a(1, 1) = 26062026#
    AssertEq modUiData.CellDate(a, 1, 1), 0, _
             "broj van opsega se NE prosledjuje mrezi"

    a(1, 1) = DateSerial(2026, 8, 21)
    AssertEq modUiData.CellDate(a, 1, 1), CDbl(DateSerial(2026, 8, 21)), _
             "pravi datum prolazi kao serijski broj"

    ' Kontrola u drugom smeru: kapija ne sme da bude presiroka i pojede validne.
    a(1, 1) = 1#
    AssertEq modUiData.CellDate(a, 1, 1), 1, "prvi dan Excel kalendara prolazi"
End Sub

' ============================================================
' 112. Geometrija kolona prati OPIS kolona
' ============================================================
' LayoutGrid (koji puni mColX / mColW) zove se iz RASPOREDA ekrana. ReloadGrid --
' promena liste, cipa, pretrage -- zove samo LoadGridFromScreen i RenderGrid.
' Bez zastavice je RenderGrid crtao sa sirinama PRETHODNE liste: kolona koja je
' tamo bila skrivena (prioritet 4 -> sirina 0) ostajala je nevidljiva i u novoj
' listi, ma koliko joj vrednost bila ispravna. Zaglavlje je pri tom umelo da
' bude vidljivo, pa je izgled bio najgori moguci: naslov stoji, celije prazne.
'
' Mereno na Fakturisanju, jer su mu liste razlicite sirine (FAKTURE ima sedam
' vidljivih kolona, ZAFAKT devet) -- dakle prelazak sa uze na siru je bas onaj
' smer u kom se kolone gube.
'
' Nalazi se SKUPLJAJU, a tvrde tek posle Unload-a: dok forma zivi, njena
' masinerija obrise Err izmedju Err.Raise i omotnice testa.

' LEGACY FORMA: PAD UCITAVANJA LISTE NIJE PRAZNA LISTA.
'
' Ova pravila su do sada bila proverljiva samo rukom -- stajala su u click
' handleru i nad Private stanjem forme. Ista klasa greske ("prazna lista je
' protumacena kao izbor") nadjena je TRI PUTA u poslednja tri PR-a, svaki put u
' review-u a ne u suite.
'
' Forma se NE prikazuje: frmBankaImport nema UserForm_Initialize, a
' UserForm_Activate (koji cita tabele) ide tek na .Show. Zato je "New" jeftin i
' bez ijednog upisa.
Private Sub T_LegacyBanka_PadUcitavanjaNijePraznaLista()
    Dim f As frmBankaImport

    Set f = New frmBankaImport

    ' PAD UCITAVANJA -> STOP. Prazan combo bi inace znacio "operater nije birao
    ' blok", odatle poziv na broj, a ako iz njega ne ispadne nijedna otkupna
    ' stavka -- ceo iznos se knjizi kao AVANS kooperanta.
    f.BiTestSetUcitanost False, "test greska"
    AssertEq f.BiTestKooperantSme(), False, _
             "pad ucitavanja liste blokova ZAUSTAVLJA rucno mapiranje"
    AssertEq (InStr(1, f.BiTestKooperantPoruka(), "NIJE") > 0), True, _
             "...i operater dobija objasnjenje, ne cutanje"

    ' Uredno ucitana lista pusta dalje -- kapija ne sme da bude siroka.
    f.BiTestSetUcitanost True, ""
    AssertEq f.BiTestKooperantSme(), True, "uredno ucitana lista pusta mapiranje"
    AssertEq f.BiTestKooperantPoruka(), "", "...bez poruke"

    ' "BLOK JE IZABRAN" se cita iz combo-a, i to je podatak koji ide writeru:
    ' od njega zavisi sme li prazan skup kandidata da postane avans.
    f.BiTestSetIzbor BIM_TIP_KOOPERANT, "1/TEST"
    AssertEq f.BiTestBlokIzabran(), True, "izabran blok se prijavljuje writeru kao izabran"

    f.BiTestSetIzbor BIM_TIP_KOOPERANT, ""
    AssertEq f.BiTestBlokIzabran(), False, _
             "prazan combo NIJE izbor -- tada blok dolazi iz poziva na broj"

    ' Drugi tip mapiranja nema blok, pa ne sme da ga ni prijavi.
    f.BiTestSetIzbor BIM_TIP_KUPAC, "1/TEST"
    AssertEq f.BiTestBlokIzabran(), False, "kod kupca se blok ne prijavljuje uopste"

    Unload f
End Sub

' CELIJA MREZE NIKAD NE OSTAVLJA TUDJI TEKST.
'
' RenderGrid radi pod "On Error Resume Next" -- namerno, da pad jedne celije ne
' obori crtanje cele mreze. Ali dok se tekst racunao U SAMOM UPISU
' (".caption = FmtBroj(CDbl(v), 0)"), pad konverzije nije preskakao samo racun
' nego i UPIS: u celiji je OSTAJAO natpis od ranijeg crtanja. Tako je 26062026
' u koloni datuma dalo "OSIROCENE_PAL" -- vrstu reda sa ekrana Oporavak.
'
' Datum je zatvoren zasebno (DatumSerijskiValidan), ali to je bio JEDAN ULAZ, ne
' klasa: isti ishod daju CDbl nad tekstom u kg/num/rsd i CLng u pill kolonama.
Private Sub T_MrezaCelija_NeostavljaTudjiTekst()
    Dim f As frmOtkupUI, body As Object
    Dim ok As Boolean
    Dim pre As String, posle As String
    Dim i As Long, kol As Long, kolPil As Long
    Dim poravnanjePre As Long, boldPre As Boolean
    Dim pilPre As String, pilSirinaPre As Single, pilNazad As String

    ' --- PRAVILO, bez forme -----------------------------------------------
    AssertEq modOtkupUI.CelijaTekst("num", "nije broj", ok), "", _
             "vrednost koja nije broj daje PRAZNU celiju"
    AssertEq ok, False, "...i to se prijavljuje kao kvar prikaza"

    AssertEq modOtkupUI.CelijaTekst("num", 5, ok), "5", "ispravan broj se crta"
    AssertEq ok, True, "ispravan broj nije kvar"

    modOtkupUI.CelijaTekst "kg", "nije broj", ok
    AssertEq ok, False, "isto vazi za kilograme"
    modOtkupUI.CelijaTekst "rsd", "nije broj", ok
    AssertEq ok, False, "isto vazi za dinare"

    ' PRAZNO ZBOG NULE NIJE KVAR. Kolona "rest" namerno ne crta 0,00 -- prazno
    ' je tu istina o podatku, ne neuspeh prikaza. Da se to broji kao kvar, log
    ' bi bio pun poruka o urednim redovima i prestao bi da se cita.
    AssertEq modOtkupUI.CelijaTekst("rest", 0, ok), "", "nula u koloni duga se ne crta"
    AssertEq ok, True, "...ali to NIJE kvar prikaza"

    ' DATUM VAN OPSEGA JESTE KVAR. FmtDatumKratko ga sam odbija, pa bi bez ove
    ' grane bas prvi nalaz ove vrste ostao neprebrojan i nevidljiv u logu.
    AssertEq modOtkupUI.CelijaTekst("date", 26062026#, ok), "", _
             "ddmmyyyy kao broj nije datum -- celija ostaje prazna"
    AssertEq ok, False, "...i to se prijavljuje"
    AssertEq modOtkupUI.CelijaTekst("date", 0, ok), "", "prazan datum je prazna celija"
    AssertEq ok, True, "...i nije kvar"

    ' Date SE PRIMA DIREKTNO. "IsNumeric" nad Date-om je False, pa je ekran koji
    ' vrednost preda onakvu kakva u tabeli jeste dobijao PRAZNU celiju -- bez
    ' greske i bez traga. Lista FAKTURA je tako imala prazan datum u SVAKOM redu,
    ' i to se nije videlo jer nijedan test nije citao NACRTAN datum.
    AssertEq modOtkupUI.CelijaTekst("date", DateSerial(2026, 3, 15), ok), "15.03.", _
             "prava Date vrednost se crta -- IsNumeric je nad njom False"
    AssertEq ok, True, "...i nije kvar"
    AssertEq modOtkupUI.CelijaTekst("date", CDbl(DateSerial(2026, 3, 15)), ok), "15.03.", _
             "serijski broj istog dana daje isti tekst"

    ' Pilula: neuspeh NE sme da postane nula, jer je nula ODREDJEN status
    ' ("Sacuvana" / "Neplaceno") -- to je svoja vrsta lazi.
    modOtkupUI.CelijaBroj "nije broj", ok
    AssertEq ok, False, "pilula nad nebrojem je kvar, ne status"
    AssertEq modOtkupUI.CelijaBroj(2, ok), 2, "ispravan kod pilule prolazi"
    AssertEq ok, True, "...i nije kvar"

    ' --- CRTANJE, nad pravom formom ---------------------------------------
    ' Pravilo iznad je tacno i kad se upis preskoci; ceo kvar je bio bas u tome.
    ' Zato se meri i sam RenderGrid.
    Set f = NewOtkupUIForm()
    modScrFakture.Scr_FkListaTestSet "FAKTURE"
    modOtkupUI.GridTestLoad "FAKTURE"
    modOtkupUI.GridRenderTest f, 1200, 600

    Set body = f.Controls("zGrid").Controls("grdBody")
    AssertEq (modOtkupUI.GridRedovaTest() > 0), True, _
             "preduslov: mreza ima sta da nacrta"
    AssertEq modOtkupUI.GridKvarCelijaTest(), 0, _
             "uredni podaci ne daju nijedan kvar (prva: " & _
             modOtkupUI.GridKvarKindTest() & ")"

    ' Kolona nad kojom konverzija UOPSTE moze da pukne. Nad txt kolonom svaka
    ' vrednost prolazi (CStr), pa bi tvrdnja tamo merila nista -- raspored tudjeg
    ' ekrana se ne pretpostavlja nego se pita.
    kol = -1
    For i = 0 To MAX_COLS - 1
        If modOtkupUI.GridSirinaKoloneTest(i) > 0 Then
            Select Case modOtkupUI.GridKindKoloneTest(i)
                Case "num", "sum0", "rsd", "mult", "kg"
                    If Len(CStr(body.Controls("c0_" & i).caption)) > 0 Then
                        kol = i
                        Exit For
                    End If
            End Select
        End If
    Next i

    AssertEq (kol >= 0), True, _
             "preduslov: bar jedna brojcana celija prvog reda je nacrtana"
    pre = CStr(body.Controls("c0_" & kol).caption)

    ' STIL KOLONE JE LAYOUT-OV POSAO, NE CRTANJEV.
    ' LayoutGrid brojcane kolone poravnava DESNO, a StyleGridCell prvoj koloni i
    ' kolonama novca daje bold. Crtanje koje bi celiju "vracalo u neutralno"
    ' pre upisa pokvarilo bi oboje -- na SVAKOM ekranu, a nijedna tvrdnja o
    ' natpisu to ne bi primetila. Zato se stil meri PRE i POSLE crtanja.
    poravnanjePre = body.Controls("c0_" & kol).TextAlign
    AssertEq poravnanjePre, fmTextAlignRight, _
             "preduslov: brojcana kolona je poravnata DESNO"
    boldPre = (body.Controls("c0_0").Font.bold = True)
    AssertEq boldPre, True, "preduslov: prva kolona je podebljana"

    ' Kolona pilule -- njen neuspeh se ne cisti isto kao obicna celija.
    kolPil = -1
    For i = 0 To MAX_COLS - 1
        If modOtkupUI.GridSirinaKoloneTest(i) > 0 Then
            Select Case modOtkupUI.GridKindKoloneTest(i)
                Case "pill", "paypill": kolPil = i: Exit For
            End Select
        End If
    Next i
    AssertEq (kolPil >= 0), True, "preduslov: lista ima kolonu sa statusnom oznakom"
    pilPre = CStr(body.Controls("c0_" & kolPil).caption)
    AssertEq (Len(pilPre) > 0), True, "preduslov: statusna oznaka je naslikana"
    pilSirinaPre = body.Controls("c0_" & kolPil).width

    ' Ista mreza, iste kolone -- samo vrednosti koje se ne mogu prikazati.
    modOtkupUI.GridTestVrednost 1, kol + 1, "NIJE-BROJ"
    modOtkupUI.GridTestVrednost 1, kolPil + 1, "NIJE-BROJ"
    modOtkupUI.GridRenderTest f, 1200, 600
    posle = CStr(body.Controls("c0_" & kol).caption)

    AssertEq posle, "", _
             "celija koja se ne moze prikazati OSTAJE PRAZNA -- ne zadrzava tudji tekst"
    AssertEq (posle <> pre), True, "stara vrednost je stvarno prepisana"
    AssertEq (modOtkupUI.GridKvarCelijaTest() > 0), True, _
             "kvar prikaza se broji, pa ostaje trag u logu"

    ' CRTANJE NE SME DA POKVARI STIL. Ovo je tvrdnja o REGRESIJI, ne o kvaru:
    ' celija koja nije mogla da se prikaze i dalje pripada svojoj koloni.
    AssertEq body.Controls("c0_" & kol).TextAlign, poravnanjePre, _
             "crtanje NE menja poravnanje kolone"
    AssertEq (body.Controls("c0_0").Font.bold = True), True, _
             "crtanje NE skida bold koji je layout postavio"

    ' PILULA SE BRISE CELA. PaintPill menja i pozadinu, boju, sirinu i BackStyle,
    ' a PaintRow pill kolone pri vracanju pozadine NAMERNO preskace -- pa bi
    ' celija kojoj je obrisan samo natpis ostala kao PRAZNA OBOJENA KUTIJA.
    AssertEq CStr(body.Controls("c0_" & kolPil).caption), "", _
             "statusna oznaka koja se ne moze naslikati NESTAJE"
    ' SIRINA SE NE SME POMERITI. Dve vrste pilule imaju dva ugovora: pravoj
    ' ("pill") sirinu racuna PaintPill, a ovoj ("paypill") je drzi LayoutGrid
    ' (mColW - 16). Ciscenje koje bi je tretiralo kao pravu pilulu postavilo bi
    ' PUNU sirinu kolone -- i ona bi takva ostala, jer PaintPayPill sirinu ne
    ' vraca, a LayoutGrid se ponovo pusta tek kad se promeni opis kolona.
    AssertEq body.Controls("c0_" & kolPil).width, pilSirinaPre, _
             "ciscenje statusne oznake NE menja sirinu celije"

    ' ROUND-TRIP: vrednost se popravlja i oznaka mora da se VRATI kakva je bila.
    ' Samo "valid -> invalid" ne bi video zaostalo stanje -- ono se vidi tek kad
    ' se posle kvara opet crta uredan podatak.
    modOtkupUI.GridTestVrednost 1, kolPil + 1, 0
    modOtkupUI.GridRenderTest f, 1200, 600
    pilNazad = CStr(body.Controls("c0_" & kolPil).caption)

    AssertEq (Len(pilNazad) > 0), True, "ispravna vrednost opet daje statusnu oznaku"
    AssertEq body.Controls("c0_" & kolPil).width, pilSirinaPre, _
             "...i celija je iste sirine kao pre kvara"

    ' Pozadina se OVDE ne tvrdi: FAKTURE nose "paypill", a PaintPayPill pozadinu
    ' ne dira -- vraca je PaintRow, pa bi tvrdnja o njoj prolazila i bez
    ' ispravke. Ciscenje POZADINE vazi za prave "pill" kolone (Dokumenta), cija
    ' se lista bez izabranog rezima ne puni. Zapisano kao NEIZMERENO, ne kao
    ' pokriveno -- v. katalog paragraf 10.5.

    modOtkupUI.GridTestLoad ""
    modOtkupUI.GridOtkaciFormuTest
    Unload f
End Sub

Private Sub T_MrezaGeometrija_PratiOpisKolona()
    Dim f As frmOtkupUI, z As Object
    Dim stara As Boolean, posle As Boolean
    Dim s7 As Single, s8 As Single

    Set f = NewOtkupUIForm()
    Set z = f.Controls("zGrid")

    ' Uza lista prvo.
    modScrFakture.Scr_FkListaTestSet "FAKTURE"
    modOtkupUI.GridTestLoad "FAKTURE"
    modOtkupUI.GridLayoutTest z, 1200, 600

    ' Pa sira: opis kolona se promenio.
    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
    modOtkupUI.GridTestLoad "FAKTURE"
    stara = modOtkupUI.GridGeomStaraTest()

    ' Ovo je ono sto RenderGrid uradi pre crtanja.
    modOtkupUI.GridOsveziGeomTest z
    s7 = modOtkupUI.GridSirinaKoloneTest(7)
    s8 = modOtkupUI.GridSirinaKoloneTest(8)
    posle = modOtkupUI.GridGeomStaraTest()

    modScrFakture.Scr_FkKorpaTestReset
    modScrFakture.Scr_FkListaTestSet "ZAFAKT"
    modOtkupUI.GridTestLoad ""
    Unload f

    AssertEq stara, True, _
             "promena opisa kolona proglasava geometriju ZASTARELOM"
    AssertEq (s7 > 0), True, "osma kolona sire liste dobija sirinu"
    AssertEq (s8 > 0), True, "deveta kolona sire liste dobija sirinu"
    AssertEq posle, False, "posle osvezavanja geometrija vise nije zastarela"
End Sub


' TEST 115: podnozje ne sme da broji novac u komadima.
'
' Ljuska je jedinicu i broj decimala citala iz `ActiveMode` -- a to je rezim
' UNOSA DOKUMENATA. Na ugovornom ekranu taj rezim ostaje onakav kakav ga je
' Dokumenta ostavila, pa je operater koji je bio na F7 (Reversi) pa presao na
' Uvoz izvoda video promet kao "Ukupno 8.950 kom" umesto "Vrednost 8.950,00
' RSD": novac izbrojan u komadima, i jos bez para. Ista klasa kao traka `zOtp`
' koja je na tudjem ekranu ostajala upaljena.
'
' Meri se na DVA nivoa. Sama tvrdnja "na ugovornom ekranu su dinari" prolazi i
' kad se komadi ugase svima -- Dokumenta bi tiho izgubila svoju jedinicu, a
' suite bi ostala zelena. Zato se tvrdi i da Dokumenta i dalje broje komade.
Private Sub T_Mreza_PodnozjeJedinicaIdeIzUgovoraEkrana()
    Dim staraMode As String, staraLista As String
    Dim natpis As String, saParama As String
    Dim dokF7 As Boolean, dokF5 As Boolean, bankaF7 As Boolean
    Dim stornoRev As Boolean, stornoFak As Boolean

    staraMode = modOtkupUI.ActiveMode

    ' UGOVOR. Zatrovan globalni rezim: operater je ostao na reversima.
    modOtkupUI.ActiveMode = "F7"
    dokF7 = modUiScreens.ScrBrojiKomade("DOKUMENTI")
    bankaF7 = modUiScreens.ScrBrojiKomade("BANKA_UVOZ")

    modOtkupUI.ActiveMode = "F5"
    dokF5 = modUiScreens.ScrBrojiKomade("DOKUMENTI")

    ' STORNO nema JEDNU semantiku: prikazuje osam tipova dokumenata, medju njima
    ' i REVERSE -- i to preko ISTOG citaca (RedoviZaTip), pa mu u podnozje stize
    ' zbir komada. Ekran koji ne odgovori dobija dinare, sto je dobar podrazumevan
    ' odgovor za Banku i Fakture, ali bi ovde 125 reversa prikazalo kao
    ' "Vrednost 125,00 RSD". Zato se pita SVAKI korisnik reversa, ne samo prvi.
    staraLista = modScrStorno.Scr_Lista()
    modScrStorno.Scr_TipTestSet STIP_REVERSI
    stornoRev = modUiScreens.ScrBrojiKomade("STORNO")
    modScrStorno.Scr_TipTestSet STIP_FAKTURA
    stornoFak = modUiScreens.ScrBrojiKomade("STORNO")
    modScrStorno.Scr_TipTestSet staraLista

    ' LJUSKA. Natpis se cita sa ekrana koji je STVARNO ucitan, sa istim
    ' zatrovanim rezimom -- bas onako kako je operater dosao.
    modOtkupUI.ActiveMode = "F7"
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    modOtkupUI.GridTestLoad "BANKA_UVOZ"
    natpis = modOtkupUI.GridPodnozjeValTest(8950)
    saParama = modOtkupUI.GridPodnozjeValTest(1234.56)

    modOtkupUI.GridTestLoad ""
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    modOtkupUI.ActiveMode = staraMode

    AssertEq dokF7, True, "Dokumenta na reversima i dalje broje komade"
    AssertEq dokF5, False, "...a na ambalazi ne -- ekran prati SVOJ rezim"
    AssertEq bankaF7, False, _
             "ugovorni ekran ne nasledjuje rezim unosa dokumenata"

    ' Dve tvrdnje, ne jedna: "Storno broji komade" bi prosla i da ekran uvek
    ' odgovara True, cime bi fakture i izvodi na njemu postali komadi.
    AssertEq stornoRev, True, "Storno lista Reversi broji komade"
    AssertEq stornoFak, False, "...a ostali tipovi na Stornu ne broje komade"

    AssertEq (InStr(1, natpis, Poruka("OTKUI_UNIT_KOM")) = 0), True, _
             "podnozje ugovornog ekrana ne pominje komade"
    AssertEq (Right$(natpis, Len(Poruka("OTKUI_UNIT_RSD"))) = Poruka("OTKUI_UNIT_RSD")), _
             True, "...nego dinare"
    AssertEq (Left$(natpis, Len(Poruka("OTKUI_FT_VREDNOST"))) = Poruka("OTKUI_FT_VREDNOST")), _
             True, "...i zove se Vrednost, ne Ukupno"

    ' Jedinica i decimale su ISTA odluka, pa se i tvrde odvojeno: komadi seku
    ' pare, a 1234.56 bez para nema "56" nigde u natpisu.
    AssertEq (InStr(1, saParama, "56") > 0), True, _
             "novac u podnozju ide sa parama"
End Sub


' TEST 116: podnozje nosi DVA novcana broja, ne njihov zbir.
'
' Promet (uplate + isplate) je jedan broj koji operater ne moze da uporedi ni sa
' cim -- izvod u ruci ima uplate i isplate odvojeno, i ne moze se rastaviti
' unazad. Ekran ih zato salje kroz sedmi clan ugovora Scr_Rows.
'
' NAJVAZNIJE JE DA SU TO NJEGOVI BROJEVI: dva broja koja ne prate filtere gora su
' od jednog koji ih prati, jer izgledaju preciznije.
Private Sub T_Mreza_PodnozjeDvaNovcanaSlota()
    Dim d As Variant, d2 As Variant, ds As Variant, sl As Variant
    Dim n As Long, nPal As Long
    Dim t0 As String, t1 As String
    Dim f As frmOtkupUI, ft As Object
    Dim vid1 As Boolean, vid2 As Boolean, vid2Pal As Boolean
    Dim capA As String, capB As String

    ' Lista se postavlja IZRICITO. Bez ovoga test meri ono sto je prethodni
    ' ostavio, a zove se po IZVODIMA -- state-dependent test je zelen dok se ne
    ' promeni redosled.
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    d = modScrBankaUvoz.Scr_Rows("sve", "")
    AssertEq (UBound(d) >= 6), True, "ugovor nosi sedmi clan -- novcane slotove"
    sl = d(6)
    AssertEq (UBound(sl) + 1), 2, "lista izvoda salje DVA novcana slota"
    AssertEq CStr(sl(0)(0)), "OTKUI_FT_UPLATE", "prvi slot su uplate"
    AssertEq CStr(sl(1)(0)), "OTKUI_FT_ISPLATE", "drugi slot su isplate"

    ' Preduslovi: bez njih bi tvrdnje ispod prolazile i nad praznim brojkama.
    AssertEq (CDbl(sl(0)(1)) > 0), True, "preduslov: test-sveska ima uplata"
    AssertEq (CDbl(sl(1)(1)) > 0), True, "preduslov: test-sveska ima isplata"
    AssertEq (CDbl(sl(0)(1)) <> CDbl(sl(1)(1))), True, _
             "preduslov: uplate i isplate se RAZLIKUJU -- inace tvrdnja ispod ne meri"

    ' ISTI FILTERI KAO REDOVI. Zbir oba slota mora biti bas promet iz petog
    ' clana -- inace su brojani po drugom pravilu od onoga sto se vidi u listi.
    AssertEq (CDbl(sl(0)(1)) + CDbl(sl(1)(1))), CDbl(d(4)), _
             "zbir slotova je promet -- brojani su pod istim filterima"

    ' I pretraga mora da ih smanji, ne samo broj redova.
    d2 = modScrBankaUvoz.Scr_Rows("sve", FX_BIM_IZVOD2)
    AssertEq (CLng(d2(2)) < CLng(d(2))), True, "preduslov: pretraga suzava listu"
    AssertEq (CDbl(d2(6)(0)(1)) < CDbl(sl(0)(1))), True, _
             "pretraga smanjuje i slotove, ne samo redove"
    ' Oba, ne samo prvi: implementacija koja filtrira uplate a isplate ostavi
    ' globalne prosla bi tvrdnju iznad.
    AssertEq (CDbl(d2(6)(1)(1)) < CDbl(sl(1)(1))), True, _
             "...i to OBA slota, ne samo prvi"
    ' I nad suzenom listom zbir mora da se poklopi sa njenim prometom.
    AssertEq (CDbl(d2(6)(0)(1)) + CDbl(d2(6)(1)(1))), CDbl(d2(4)), _
             "i u suzenoj listi je zbir slotova njen promet"

    ' STAVKE nose ISTI ugovor -- ekran ima dva citaca, i oba su izmenjena.
    modScrBankaUvoz.Scr_BuListaTestSet "STAVKE"
    ds = modScrBankaUvoz.Scr_Rows("sve", "")
    AssertEq (UBound(ds) >= 6), True, "i lista stavki nosi sedmi clan"
    AssertEq (UBound(ds(6)) + 1), 2, "lista stavki salje DVA novcana slota"
    AssertEq CStr(ds(6)(0)(0)), "OTKUI_FT_UPLATE", "stavke: prvi slot su uplate"
    AssertEq CStr(ds(6)(1)(0)), "OTKUI_FT_ISPLATE", "stavke: drugi slot su isplate"
    AssertEq (CDbl(ds(6)(0)(1)) + CDbl(ds(6)(1)(1))), CDbl(ds(4)), _
             "stavke: zbir slotova je promet"

    ' LJUSKA. Da li je uopste preuzela oba, i sa svojim natpisima.
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    modOtkupUI.GridTestLoad "BANKA_UVOZ"
    n = modOtkupUI.GridPodnozjeSlotBrojTest()
    t0 = modOtkupUI.GridPodnozjeSlotTest(0)
    t1 = modOtkupUI.GridPodnozjeSlotTest(1)

    ' Ekran koji slotove ne salje mora da ostane na starom putu.
    modOtkupUI.GridTestLoad "PALETE"
    nPal = modOtkupUI.GridPodnozjeSlotBrojTest()

    ' CRTANJE, nad pravom formom. Model i natpis nisu isto sto i NACRTAN slot:
    ' kod ove iste liste je vec jednom bio tacan zbir uz nevidljivu kontrolu
    ' (prelazak novcanih kolona na "rest", v6-ui-181).
    Set f = NewOtkupUIForm()
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    modOtkupUI.GridTestLoad "BANKA_UVOZ"
    modOtkupUI.GridRenderTest f, 1200, 600
    Set ft = f.Controls("zGrid").Controls("grdFoot")
    vid1 = ft.Controls("ftVal").Visible
    vid2 = ft.Controls("ftVal2").Visible
    capA = CStr(ft.Controls("ftVal").caption)
    capB = CStr(ft.Controls("ftVal2").caption)

    ' Ekran bez slotova mora da drugi slot UGASI -- inace tudja brojka ostaje
    ' na ekranu, kao nekad traka otpremnice.
    modOtkupUI.GridTestLoad "PALETE"
    modOtkupUI.GridRenderTest f, 1200, 600
    vid2Pal = ft.Controls("ftVal2").Visible

    modOtkupUI.GridTestLoad ""
    modScrBankaUvoz.Scr_BuListaTestSet "IZVODI"
    Unload f
    modOtkupUI.GridOtkaciFormuTest

    AssertEq n, 2, "ljuska je preuzela oba slota"
    AssertEq (InStr(1, t0, Poruka("OTKUI_FT_UPLATE")) = 1), True, _
             "prvi slot nosi natpis Uplate"
    AssertEq (InStr(1, t1, Poruka("OTKUI_FT_ISPLATE")) = 1), True, _
             "drugi slot nosi natpis Isplate"
    AssertEq (Right$(t0, Len(Poruka("OTKUI_UNIT_RSD"))) = Poruka("OTKUI_UNIT_RSD")), _
             True, "slot je u dinarima"
    ' Ceo natpis se razlikuje vec po natpisu, pa se poredi ono STO OSTANE kad
    ' se natpis skine -- inace tvrdnja prolazi i kad oba slota crtaju isti iznos.
    AssertEq (Mid$(t0, Len(Poruka("OTKUI_FT_UPLATE")) + 1) <> _
              Mid$(t1, Len(Poruka("OTKUI_FT_ISPLATE")) + 1)), True, _
             "slotovi ne nose isti IZNOS"
    AssertEq nPal, 0, "ekran bez slotova ostaje na zbiru vrednosti"

    AssertEq vid2, True, "drugi slot je STVARNO nacrtan, ne samo izracunat"
    AssertEq vid1, True, "...uz prvi, oba u isto vreme"
    AssertEq (capA <> capB), True, "nacrtani slotovi nose razlicite natpise"
    AssertEq vid2Pal, False, _
             "na ekranu bez slotova drugi slot se GASI -- ne ostaje tudja brojka"
End Sub


' TEST 117: poruka o nedostajucoj koloni razlikuje TRI stanja, i kaze da li je
' bas trazena kolona vidjena u svezem prolazu.
'
' Povod: u logu nad radnom sveskom stoji "Nedostaje kolona 'VozacID' u tabeli
' 'tblZbirna'" -- a ta kolona u toj svesci POSTOJI (provereno dump_schema-om).
' Poruka ne razlikuje tri razlicita stanja: kolone stvarno nema, tabele nema, i
' zaglavlje je drugacije od ocekivanog. Zato sada nosi i ono STO JE VIDELA.
'
' Sam uzrok nije reprodukovan i ovde se ne popravlja naslepo. Zabelezeno je sta
' se zna: nula iz GetColumnIndex se KESIRA za ceo BeginTableCache prozor, pa bi
' jedan trenutan neuspeh postao trajan, a kapije koje su fail-closed (npr.
' ZbirnaBrojJeDvosmislenIkad) na to staju. Dijagnostika ispod je da sledeci put
' ne gadjamo.
Private Sub T_Kolona_TrazenjeNeGutaGresku()
    Dim iPostoji As Long, iMalim As Long, iNema As Long, iKes As Long
    Dim errPosle As Long
    Dim poruka As String, porukaKes As String, porukaTbl As String

    iPostoji = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    AssertEq (iPostoji > 0), True, "preduslov: postojeca kolona se nalazi"

    ' Zateceno ponasanje, ovde samo zapisano da se ne izgubi: nepostojeca kolona
    ' daje nulu, a Err ostaje cist (On Error GoTo 0 u VBA cisti Err).
    '
    ' Ni jedna od te dve tvrdnje NEMA svoju sabotazu i to je namerno: proverio
    ' sam merenjem da ih zatecen kod vec zadovoljava, pa bi sabotaza koja ih
    ' obara morala da izmisli kvar koji se u ovom kodu ne moze desiti.
    Err.Clear
    iNema = GetColumnIndex(TBL_OTKUP, "NemaOvakveKoloneNigde")
    errPosle = Err.Number
    Err.Clear

    AssertEq iNema, 0, "nepostojeca kolona daje nulu"
    AssertEq errPosle, 0, "...i ne ostavlja Err ziv"

    iMalim = GetColumnIndex(TBL_OTKUP, LCase$(COL_OTK_ID))
    AssertEq iMalim, iPostoji, "ime kolone se poredi bez obzira na velicinu slova"

    ' SIMPTOM IZ POGONA, izmeren. Trazenje kaze NULA za kolonu koja u zaglavlju
    ' POSTOJI -- tacno ono sto je stajalo u logu nad radnom sveskom.
    '
    ' Izaziva se kroz kes: nula se pamti za ceo BeginTableCache prozor. Time se
    ' NE tvrdi da je kes uzrok prvog neuspeha (nije reprodukovan) -- tvrdi se da
    ' jednom zapamcena nula prezivi prozor, i da poruka to sada ume da razlikuje
    ' od stvarnog nedostatka kolone.
    BeginTableCache
    modDataAccess.KesKoloneTestSet TBL_OTKUP, COL_OTK_ID, 0
    iKes = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    On Error Resume Next
    Err.Clear
    RequireColumnIndex TBL_OTKUP, COL_OTK_ID, "T_Kolona"
    porukaKes = Err.description
    Err.Clear
    On Error GoTo 0
    EndTableCache

    AssertEq iKes, 0, "kesirana nula prezivi prozor -- simptom iz loga se ponavlja"
    AssertEq (InStr(1, porukaKes, "VIDJENA") > 0), True, _
             "poruka kaze da je trazena kolona VIDJENA u svezem prolazu"
    AssertEq (InStr(1, porukaKes, "pozicija") > 0), True, _
             "...i na kojoj je poziciji"

    ' DIJAGNOSTIKA nad stvarno nedostajucom kolonom. Bez zaglavlja bi isti tekst
    ' opisivao tri razlicita stanja: kolone nema, tabele nema, citanje je puklo.
    On Error Resume Next
    Err.Clear
    RequireColumnIndex TBL_OTKUP, "NemaOvakveKoloneNigde", "T_Kolona"
    poruka = Err.description
    Err.Clear
    On Error GoTo 0

    AssertEq (InStr(1, poruka, "NemaOvakveKoloneNigde") > 0), True, _
             "poruka imenuje kolonu koja nedostaje"
    AssertEq (InStr(1, poruka, COL_OTK_ID) > 0), True, _
             "...i zaglavlje koje je stvarno videla"
    AssertEq (InStr(1, poruka, "NIJE vidjena") > 0), True, _
             "...i da trazene kolone u zaglavlju stvarno nema"

    ' Trece stanje: tabele uopste nema. Mora da se razlikuje od prva dva.
    On Error Resume Next
    Err.Clear
    RequireColumnIndex "tblNePostojiNikako", "BiloSta", "T_Kolona"
    porukaTbl = Err.description
    Err.Clear
    On Error GoTo 0

    AssertEq (InStr(1, porukaTbl, "tabela nije nadjena") > 0), True, _
             "za nepostojecu tabelu poruka kaze da TABELE nema"
End Sub


' TEST 123: nula iz trazenja kolone se NE pamti.
'
' Nastavak testa 117, koji je isti simptom samo ZAPISAO. Kes indeksa kolone je
' pamtio i nulu, pa je jedan trenutan neuspeh vazio za ceo BeginTableCache
' prozor: svaki sledeci poziv nad istom kolonom dobijao je istu nulu bez ijednog
' novog pokusaja, a fail-closed kapije (RequireColumnIndex) na to staju.
'
' Isto pravilo vec vazi za kes TABELA -- test 53, modUiData.CachedTable kesira
' samo uspeh. Rupa je bila u kesu KOLONA, na istom mestu i iz istog razloga.
'
' Sta ovo NE tvrdi: da je uzrok PRVOG neuspeha popravljen. On nije reprodukovan
' (postmortem par 11) i ovde se ne dira naslepo. Skinuta mu je TRAJNOST.
Private Sub T_KesKolone_NeMemoiseNulu()
    Dim iPostoji As Long, iNema As Long
    Dim imaUspeh As Boolean, imaNulu As Boolean
    Const KOL_KOJE_NEMA As String = "NemaOvakveKoloneNigde"

    ' MERI se unutar prozora, TVRDI posle EndTableCache. AssertEq puca na prvom
    ' padu, pa bi tvrdnja unutar prozora ostavila kes otvoren svim narednim
    ' testovima -- a ovaj test je pisan da bude i crven. Isto kao u testu 117.
    BeginTableCache
    iPostoji = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    imaUspeh = modDataAccess.KesKoloneImaKljuc(TBL_OTKUP, COL_OTK_ID)
    iNema = GetColumnIndex(TBL_OTKUP, KOL_KOJE_NEMA)
    imaNulu = modDataAccess.KesKoloneImaKljuc(TBL_OTKUP, KOL_KOJE_NEMA)
    EndTableCache

    ' POZITIVNA KONTROLA PRVA: bez nje bi test prosao i nad verzijom koja kes
    ' prosto ugasi -- a kes postoji da bi se izbeglo ~80 skenova po prozoru.
    AssertEq (iPostoji > 0), True, "preduslov: postojeca kolona se nalazi"
    AssertEq imaUspeh, True, "uspesan indeks se i dalje pamti"

    AssertEq iNema, 0, "nepostojeca kolona daje nulu"
    AssertEq imaNulu, False, _
             "nula se NE pamti -- trenutan neuspeh ne postaje trajan"
End Sub

' TEST 118: pozadina PRAVE pilule se cisti kad se vrednost ne moze prikazati.
'
' Ovo je rupa zapisana u katalogu 10.6 kao NEIZMERENA. Razlog je tada bio: jedina
' lista koja se puni bez forme a ima statusnu oznaku je FAKTURE, a njena je
' "paypill" -- PaintPayPill pozadinu ne dira, pa bi tvrdnja o njoj prolazila i bez
' ispravke. Prava "pill" kolona zivi na Dokumentima, "cija se lista bez izabranog
' rezima ne puni".
'
' TAJ ZAKLJUCAK JE BIO NETACAN. Lista Dokumenata se puni i bez forme -- treba joj
' samo rezim (ActiveMode) i podlista, a podlista se bira PRODUKCIONIM putem
' (Scr_Event "lsSVI"), bez ijednog novog seam-a. Izmereno: 14 redova, 13 kolona,
' trinaesta je bas "pill".
'
' Dve vrste pilule su dva ugovora (10.4): "pill" se brise CELA -- natpis, pozadina
' i sirina -- jer je pilula bez natpisa i dalje obojen pravougaonik koji tvrdi
' stanje. "paypill" menja samo natpis. Ovaj test meri onu prvu.
Private Sub T_MrezaPilula_PozadinaSeCisti()
    Dim f As frmOtkupUI, body As Object
    Dim staraMode As String, staraLista As String
    Dim k As Long, kPil As Long
    Dim stilPre As Long, stilPosle As Long, stilNazad As Long
    Dim capPre As String, capPosle As String

    staraMode = modOtkupUI.ActiveMode
    staraLista = modScrDokumenti.Scr_Lista()

    Set f = NewOtkupUIForm()
    modOtkupUI.ActiveMode = "F1"
    ' Lista dokumenata, istim putem kojim je bira operater (klik na cip).
    modScrDokumenti.Scr_Event "lsSVI", "Click"
    modOtkupUI.GridTestLoad "DOKUMENTI"
    modOtkupUI.GridRenderTest f, 1200, 600
    Set body = f.Controls("zGrid").Controls("grdBody")

    ' Kolona se TRAZI, ne pretpostavlja: nad "txt" kolonom bi svaka vrednost
    ' prolazila, pa bi tvrdnja tamo merila nista.
    kPil = -1
    For k = 0 To 13
        If modOtkupUI.GridKindKoloneTest(k) = "pill" Then kPil = k: Exit For
    Next k

    If kPil >= 0 Then
        stilPre = CLng(body.Controls("c0_" & kPil).BackStyle)
        capPre = CStr(body.Controls("c0_" & kPil).caption)

        ' Vrednost koja se ne moze prikazati ide POSLE ucitavanja -- takav red u
        ' tabeli je jednom vec oborio sedam tudjih testova sa Overflow.
        modOtkupUI.GridTestVrednost 1, kPil + 1, "NIJE-BROJ"
        modOtkupUI.GridRenderTest f, 1200, 600
        stilPosle = CLng(body.Controls("c0_" & kPil).BackStyle)
        capPosle = CStr(body.Controls("c0_" & kPil).caption)

        ' ROUND-TRIP: uredna vrednost mora da vrati i pozadinu, ne samo natpis.
    '
    ' Ova tvrdnja NEMA svoju sabotazu, i to je namerno: pozadinu i prvi put i
    ' posle popravke slika ISTA rutina (PaintPill), pa svaka sabotaza nad njom
    ' obara preduslov iznad umesto ove tvrdnje (zamka 6). Izmereno, ne
    ' pretpostavljeno.
        modOtkupUI.GridTestVrednost 1, kPil + 1, 0
        modOtkupUI.GridRenderTest f, 1200, 600
        stilNazad = CLng(body.Controls("c0_" & kPil).BackStyle)
    End If

    modOtkupUI.GridTestLoad ""
    modScrDokumenti.Scr_Event "ls" & staraLista, "Click"
    modOtkupUI.ActiveMode = staraMode
    Unload f
    modOtkupUI.GridOtkaciFormuTest

    AssertEq (kPil >= 0), True, _
             "preduslov: lista Dokumenata se puni bez forme i ima 'pill' kolonu"
    AssertEq stilPre, CLng(fmBackStyleOpaque), _
             "preduslov: uredna pilula je NASLIKANA (neprozirna pozadina)"

    ' OVO JE TVRDNJA ZBOG KOJE TEST POSTOJI. Natpis se brisao i pre; pozadina je
    ' ostajala, pa je celija tvrdila stanje koje nema pokrice.
    AssertEq stilPosle, CLng(fmBackStyleTransparent), _
             "pilula koja se ne moze prikazati gubi i POZADINU, ne samo natpis"
    AssertEq capPosle, "", "...i natpis"
    AssertEq (Len(capPre) > 0), True, "kontrola: pre kvara je natpis postojao"
    AssertEq stilNazad, CLng(fmBackStyleOpaque), _
             "uredna vrednost VRACA pozadinu -- ciscenje ne ostaje zauvek"
End Sub


' TEST 119: u frmDokumenta prazna lista blokova NIJE izbor.
'
' Ista klasa greske koju je frmBankaImport imao i koja je tamo zatvorena u PR
' #220. Ovde je stajala nedirnuta:
'
'   btnUnosOMUlaz_Click:  If cmbOtkupBlok.ListIndex >= 0 Then ... Else AVANS
'
' Prazan kombo i NEUSPELO ucitavanje izgledaju isto -- ListIndex je -1 u oba
' slucaja. Punjenje (FillOpenOtkupi) je do sada padalo bez traga u stanju forme,
' a pozivalac (cmbPrimalacOMUlaz_Change) nema rukovaoca -- pa je novac tiho
' postajao AVANS kooperanta umesto da se knjizi na blok.
'
' Forma se u testu NE prikazuje: frmDokumenta nema UserForm_Initialize, pa je
' New frmDokumenta jeftin i ne cita nijednu tabelu (isti razlog kao 11.1).
Private Sub T_LegacyDok_PadListeBlokovaNijeAvans()
    Dim f As frmDokumenta
    Dim smePad As Boolean, smeOk As Boolean
    Dim porukaPad As String, porukaOk As String
    Dim padSaIzborom As Boolean, padBezIzbora As Boolean, okSaIzborom As Boolean

    Set f = New frmDokumenta

    ' PAD UCITAVANJA -> STOP. Nalazi se skupljaju pa tvrde POSLE Unload-a: pad
    ' tvrdnje nad zivom formom ostavlja formu u memoriji i poruka se ne vidi.
    f.DokTestSetBlokUcitanost False, "test greska"
    smePad = f.DokTestBlokSme()
    porukaPad = f.DokTestBlokPoruka()

    ' Uredno ucitana lista pusta dalje -- kapija ne sme da bude sira od kvara.
    ' Prazna lista posle USPESNOG citanja stvarno znaci "nema otvorenih blokova",
    ' i avans je tada ispravan.
    ' IZBOR NE SME DA ZAOBIDJE KAPIJU.
    '
    ' Pad usred punjenja ostavlja kombo DELIMICNO napunjen: operater tada bira
    ' red iz nepotpune liste, ListIndex je >= 0, i kapija koja stoji samo u AVANS
    ' grani se nikad ne pita. Odluka zato ne sme da zavisi od toga da li je red
    ' izabran -- i to se ovde tvrdi za obe vrednosti.
    padSaIzborom = f.DokTestKnjizenjeSme(True)
    padBezIzbora = f.DokTestKnjizenjeSme(False)

    f.DokTestSetBlokUcitanost True, ""
    smeOk = f.DokTestBlokSme()
    porukaOk = f.DokTestBlokPoruka()
    okSaIzborom = f.DokTestKnjizenjeSme(True)

    Unload f

    AssertEq smePad, False, _
             "pad ucitavanja liste blokova ZAUSTAVLJA knjizenje avansa"
    AssertEq (InStr(1, porukaPad, "NIJE") > 0), True, _
             "...i operater dobija objasnjenje, ne cutanje"
    AssertEq (InStr(1, porukaPad, "test greska") > 0), True, _
             "...u kojem stoji i sta je puklo"
    AssertEq padSaIzborom, False, _
             "ni IZABRAN blok ne prolazi kad je ucitavanje palo"
    AssertEq padBezIzbora, False, "...ni prazan izbor"

    AssertEq smeOk, True, "uredno ucitana lista pusta avans"
    AssertEq porukaOk, "", "...bez poruke"
    AssertEq okSaIzborom, True, "...i pusta izabran blok"
End Sub

' ============================================================
' 120: ISTA GRESKA NA STRANI KUPCA (F6 / "Izlaz").
'
' FillOpenFakture je pad citanja gubio isto kao FillOpenOtkupi: kombo ostane
' prazan, a btnUnosIzlaz_Click iz praznog polja zakljucuje "nema fakture" i
' knjizi NOV_KUPCI_AVANS. Razlika se vidi tek u saldu kupca.
'
' Kapija je ovde NAMERNO uza nego kod blokova: rezim pusta i unos same ambalaze
' bez novca, a tada odluke faktura/avans nema -- pa lista ne sme da ga zaustavi.
' Ta uzina je tvrdnja, ne izuzetak, i ima svoju sabotazu.
' ============================================================
Private Sub T_LegacyDok_PadListeFakturaNijeAvans()
    Dim f As frmDokumenta
    Dim smePad As Boolean, smeOk As Boolean
    Dim porukaPad As String, porukaOk As String
    Dim padSaIzborom As Boolean, bezNovca As Boolean, okSaIzborom As Boolean

    Set f = New frmDokumenta

    ' Nalazi se skupljaju pa tvrde POSLE Unload-a -- pad tvrdnje nad zivom formom
    ' ostavlja formu u memoriji i poruka se ne vidi.
    f.DokTestSetFaktUcitanost False, "test greska"
    smePad = f.DokTestUplataSme(1000#, False)
    porukaPad = f.DokTestUplataPoruka()
    ' Izbor ne sme da zaobidje kapiju: delimicno napunjena lista IMA izbor.
    padSaIzborom = f.DokTestUplataSme(1000#, True)
    ' ...ali unos bez novca kapija ne dira.
    bezNovca = f.DokTestUplataSme(0#, False)

    f.DokTestSetFaktUcitanost True, ""
    smeOk = f.DokTestUplataSme(1000#, False)
    porukaOk = f.DokTestUplataPoruka()
    okSaIzborom = f.DokTestUplataSme(1000#, True)

    Unload f

    AssertEq smePad, False, _
             "pad ucitavanja liste faktura ZAUSTAVLJA knjizenje avansa kupca"
    AssertEq (InStr(1, porukaPad, "NIJE") > 0), True, _
             "...i operater dobija objasnjenje, ne cutanje"
    AssertEq (InStr(1, porukaPad, "test greska") > 0), True, _
             "...u kojem stoji i sta je puklo"
    AssertEq padSaIzborom, False, _
             "ni IZABRANA faktura ne prolazi kad je ucitavanje palo"
    AssertEq bezNovca, True, _
             "unos bez novca NE staje zbog liste faktura"

    AssertEq smeOk, True, "uredno ucitana lista pusta uplatu"
    AssertEq porukaOk, "", "...bez poruke"
    AssertEq okSaIzborom, True, "...i pusta izabranu fakturu"
End Sub

' Recnik kakav modOtkupUI.SkupiPolja salje ekranu -- samo kljucevi od kojih
' zavisi kapija ucitanosti.
Private Function LjuskaNovacPolja(ByVal rezim As String, ByVal novac As Double, _
                                  ByVal partnerTip As String) As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("rezim") = rezim
    p("novac") = novac
    p("partnerTip") = partnerTip
    p("kooperantID") = "P-1"
    Set LjuskaNovacPolja = p
End Function

' ============================================================
' 121: ISTA GRESKA U NOVOJ LJUSCI (F5 i F6).
'
' modOtkupUI.FillOpenBlokovi / FillOpenFakture su pad prijavljivali u Debug.Print
' -- prozor koji u pogonu niko ne gleda -- pa je prazan combo isao dalje kao
' "nema otvorenih". modNovacUnos iz praznog otkupID/fakturaID bira AVANS i ne
' moze da zna razliku: nju zna samo onaj ko je listu punio.
'
' Kapija se meri nad RECNIKOM, bez forme, jer se odluka i donosi nad njim.
' ============================================================
Private Sub T_Ljuska_PadListeNovcaNijeAvans()
    Dim padF5 As Boolean, padF6 As Boolean, poruka As String
    Dim omIsplata As Boolean, bezNovca As Boolean, drugiRezim As Boolean
    Dim okF5 As Boolean, okF6 As Boolean, porukaOk As String

    modOtkupUI.UiTestSetListaUcitanost False, "test greska", False, "test greska"

    padF5 = modOtkupUI.UiTestNovacListaSme(LjuskaNovacPolja("AMB_ISPLATE", 1000#, "KOOP"))
    poruka = modOtkupUI.UiTestNovacListaPoruka(LjuskaNovacPolja("AMB_ISPLATE", 1000#, "KOOP"))
    padF6 = modOtkupUI.UiTestNovacListaSme(LjuskaNovacPolja("AMB_UPLATE", 1000#, "KUP"))

    ' Kapija ne sme da bude sira od kvara -- tri slucaja koje NE dodiruje:
    omIsplata = modOtkupUI.UiTestNovacListaSme(LjuskaNovacPolja("AMB_ISPLATE", 1000#, "OM"))
    bezNovca = modOtkupUI.UiTestNovacListaSme(LjuskaNovacPolja("AMB_ISPLATE", 0#, "KOOP"))
    drugiRezim = modOtkupUI.UiTestNovacListaSme(LjuskaNovacPolja("OTKUP", 1000#, "KOOP"))

    modOtkupUI.UiTestSetListaUcitanost True, "", True, ""
    okF5 = modOtkupUI.UiTestNovacListaSme(LjuskaNovacPolja("AMB_ISPLATE", 1000#, "KOOP"))
    okF6 = modOtkupUI.UiTestNovacListaSme(LjuskaNovacPolja("AMB_UPLATE", 1000#, "KUP"))
    porukaOk = modOtkupUI.UiTestNovacListaPoruka(LjuskaNovacPolja("AMB_ISPLATE", 1000#, "KOOP"))

    ' REDOSLED JE DEO KONSTRUKCIJE. AssertEq staje na prvoj palo tvrdnji, pa
    ' svaka sabotaza mora prva da sretne BAS svoju: sabotaza koja kapiju siri na
    ' sve rezime usput obara i padF6, pa 'drugiRezim' mora doci pre njega.
    AssertEq padF5, False, "pad liste blokova ZAUSTAVLJA isplatu kooperantu u ljusci"
    AssertEq (InStr(1, poruka, "test greska") > 0), True, _
             "...uz poruku u kojoj stoji sta je puklo"

    AssertEq omIsplata, True, "isplata otkupnom mestu ne zavisi od liste blokova"
    AssertEq bezNovca, True, "unos bez novca ne staje zbog liste"
    AssertEq drugiRezim, True, "rezim bez tih listi kapiju ne oseca"

    AssertEq padF6, False, "pad liste faktura ZAUSTAVLJA uplatu kupca u ljusci"

    AssertEq okF5, True, "uredno ucitana lista blokova pusta isplatu"
    AssertEq okF6, True, "uredno ucitana lista faktura pusta uplatu"
    AssertEq porukaOk, "", "...bez poruke"
End Sub

' ============================================================
' 122: FILTER STORNIRANIH NA NEDOSTAJUCU KOLONU VISE NE CUTI.
'
' ExcludeStornirano je pitao GetColumnIndex za kolonu Stornirano i na NULU tiho
' vracao NEFILTRIRANE podatke -- iz 183 poziva, ukljucujuci read-modele
' otvorenih faktura i otkupnih blokova. Posledica je gora od pogresne
' klasifikacije novca iz testova 119-121: tamo je novac dobijao pogresan TIP,
' ovde storniran dokument dobija pogresno POSTOJANJE, pa uplata moze da ode na
' otkazanu fakturu.
'
' Nula je imala dva znacenja i to je bilo celo pitanje: "ova tabela storno pojam
' nema" (maticni podaci -- prolaz je tacan) i "kolona nije nadjena" (drift).
' Registar u modSchemaGuard ih razdvaja, i obe strane se ovde mere.
' ============================================================
Private Sub T_StornoFilter_NedostajucaKolonaNijeTisina()
    Dim nosiDok As Boolean, nosiMat As Boolean
    Dim poruka As String, greskaMat As String
    Dim porukaPrazno As String, porukaNepoznata As String
    Dim dataDok As Variant, dataMat As Variant, prosloMat As Variant
    Dim prazno As Variant, nepoznato As Variant
    Dim redovaPre As Long, redovaPosle As Long

    nosiDok = modSchemaGuard.TabelaNosiStorno(TBL_OTKUP)
    nosiMat = modSchemaGuard.TabelaNosiStorno(TBL_KOOPERANTI)

    ' (1) TABELA IZ REGISTRA BEZ KOLONE -> GLASAN PAD.
    '
    ' Nula se izaziva kroz kes, istim putem kao u testu 117. Time se NE tvrdi da
    ' je kes uzrok ijednog pada iz pogona -- tvrdi se da ExcludeStornirano na
    ' nulu vise ne propusta nefiltrirane podatke.
    BeginTableCache
    dataDok = GetTableData(TBL_OTKUP)
    modDataAccess.KesKoloneTestSet TBL_OTKUP, COL_STORNIRANO, 0
    On Error Resume Next
    Err.Clear
    dataDok = ExcludeStornirano(dataDok, TBL_OTKUP)
    poruka = Err.description
    Err.Clear
    On Error GoTo 0
    EndTableCache

    ' (2) PRAZNA TABELA IZ REGISTRA BEZ KOLONE -- TAKODJE PAD.
    '
    ' Ugovor je "tabela iz STORNO_TABELE bez kolone znaci drift", a ne "drift se
    ' prijavljuje samo dok tabela ima redova". Dok je IsEmpty izlazio prvi, prazna
    ' tabela je kapiju preskakala i tvrdnja iznad to nije videla -- fixture ima
    ' redove, pa se ta grana nikad nije ni takla.
    BeginTableCache
    modDataAccess.KesKoloneTestSet TBL_OTKUP, COL_STORNIRANO, 0
    On Error Resume Next
    Err.Clear
    prazno = ExcludeStornirano(Empty, TBL_OTKUP)
    porukaPrazno = Err.description
    Err.Clear
    On Error GoTo 0
    EndTableCache

    ' (3) TABELA KOJU REGISTAR NE POZNAJE -- PAD, ne tihi prolaz.
    '
    ' Bez ove kapije "TabelaNosiStorno = False" opet znaci dve stvari: eksplicitno
    ' BEZ_STORNA i "niko je nije klasifikovao". Staticka provera to ne zatvara,
    ' jer namerno preskace pozive sa promenljivim imenom tabele -- a takvih ima
    ' (modIntegritet.CollectBrojZbirne, modDokumenta.SumByBroj).
    BeginTableCache
    dataDok = GetTableData(TBL_OTKUP)
    On Error Resume Next
    Err.Clear
    nepoznato = ExcludeStornirano(dataDok, "tblNijeURegistru")
    porukaNepoznata = Err.description
    Err.Clear
    On Error GoTo 0
    EndTableCache

    ' (4) MATICNI PODACI STORNO POJAM NEMAJU -- prolaz je TACAN ishod.
    ' Kapija sme da bude siroka tacno koliko i kvar; bez ove tvrdnje bi
    ' fail-closed zaustavio i citanje kooperanata, koje nikad nije bilo u pitanju.
    BeginTableCache
    dataMat = GetTableData(TBL_KOOPERANTI)
    If IsArray(dataMat) Then redovaPre = UBound(dataMat, 1)
    On Error Resume Next
    Err.Clear
    prosloMat = ExcludeStornirano(dataMat, TBL_KOOPERANTI)
    greskaMat = Err.description
    Err.Clear
    On Error GoTo 0
    EndTableCache
    If IsArray(prosloMat) Then redovaPosle = UBound(prosloMat, 1)

    AssertEq nosiDok, True, "dokument tabela je u registru storna"
    AssertEq nosiMat, False, "maticni podaci nisu u registru storna"
    AssertEq (InStr(1, poruka, COL_STORNIRANO) > 0), True, _
             "nedostajuca kolona storna PADA i imenuje kolonu, ne propusta tiho"
    AssertEq (InStr(1, porukaPrazno, COL_STORNIRANO) > 0), True, _
             "...i kad je tabela PRAZNA -- drift ne ceka da bude redova"
    AssertEq (InStr(1, porukaNepoznata, "registru") > 0), True, _
             "tabela koju registar ne poznaje PADA, ne prolazi kao da nema storno"

    AssertEq greskaMat, "", "tabela bez storno pojma prolazi bez greske"
    AssertEq (redovaPre > 0), True, "...nad tabelom koja stvarno ima redove"
    AssertEq redovaPosle, redovaPre, "...i vraca sve svoje redove"
End Sub
