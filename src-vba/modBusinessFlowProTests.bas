Attribute VB_Name = "modBusinessFlowProTests"
'Attribute VB_Name = "modBusinessFlowProTests"
Option Explicit

' ============================================================
' modBusinessFlowProTests
'
' Professional smoke/regression suite for an empty OtkupApp workbook.
'
' What this suite does:
'   1) Seeds minimum master data if missing.
'   2) Runs isolated happy-path document chain:
'        Otkup -> Otpremnica -> Zbirna -> Prijemnica -> Faktura
'   3) Runs rollback/invalid-save checks.
'   4) Runs traceability auto-link regression tests.
'   5) Runs global data quality audit for cross-zbirna wrong links.
'
' Important:
'   - Tests create TST-* rows and do not physically delete data.
'   - Cleanup is optional and soft-storno based where possible.
'   - The cross-zbirna regression is expected to FAIL until
'     AutoLinkOtkupOtpremnica uses BrojZbirne as part of the preferred key.
'
' Recommended run order:
'   RunBusinessFlowProSuite
'
' Optional:
'   RunBusinessFlowProSeedOnly
'   RunBusinessFlowProTraceabilityOnly
'   RunBusinessFlowProAuditOnly
'   SoftStornoBusinessFlowTestRows
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long

' Gate: verdikt se podize u EndRun, koji zovu SVA cetiri Run* runnera -- i na
' uspesnoj putanji i iz EH. LogFatal takodje inkrementira m_Failed, pa prekinuta
' suite izlazi kao pad, ne kao "nije se desilo". Bez ovoga bi runner video
' "proslo bez greske", sto NIJE isto sto i sve provere prosle.
' Konvencija: modTestBanka.ERR_BIT_SUITE_FAILED.
Private Const ERR_BFP_SUITE_FAILED As Long = vbObjectError + 2966
Private m_Skipped As Long
Private m_RunID As String
Private m_DateSeq As Long

Private Const TEST_LOG_SHEET As String = "BUSINESS_FLOW_PRO_TEST_LOG"

' Izvestaj na DISK, u istom formatu kao modTest.WriteResultFile.
'
' Debug.Print ne izlazi iz Excela, a opis podignute greske ne prezivi COM
' granicu (pywin32 vidi golo "Exception occurred"). Bez ovoga run_vba vidi
' samo da je suite pala, ne i KOJA tvrdnja -- pa je dvosmerni dokaz nad njom
' slep: sabotaza koja obori bas ciljanu tvrdnju i ona koja obori neku drugu
' izgledaju isto. Isti razlog i isti format kao modTestBanka.
Private m_Report As String

Private Const TEST_ST_ID As String = "ST-90001"
Private Const TEST_KOOP_ID As String = "KOOP-90001"
Private Const TEST_VOZ_ID As String = "VOZ-90001"
' Drugi vozac: zbirna sme da nosi otpremnice SAMO svog vozaca (A15/domen).
Private Const TEST_VOZ_ID_B As String = "VOZ-90002"
Private Const TEST_KUP_ID As String = "KUP-90001"
Private Const TEST_KULTURA_ID As String = "KUL-90001"
' Drugi kooperant: parcela sme da pripada nekom KO POSTOJI, a nije vlasnik.
' Bez njega bi test vlasnistva parcele zapravo merio postojanje kooperanta.
Private Const TEST_KOOP2_ID As String = "KOOP-90002"
' Kultura BEZ sorte: prazna sorta na dokumentu je legitimna tacno uz nju.
Private Const TEST_KUL_BEZ_SORTE_ID As String = "KUL-90002"
Private Const TEST_PAR_ID As String = "PAR-90001"

Private Const TEST_VRSTA As String = "Test Jabuka"
Private Const TEST_SORTA As String = "Test Sorta"
Private Const TEST_TIP_AMB As String = "Test Gajba"
' Drugi tip ambalaze: clanovi otpremnice moraju biti homogeni -- header nosi
' jedan TipAmbalaze, pa 20 plasticnih + 30 drvenih gajbi nije 50 gajbi.
Private Const TEST_TIP_AMB_B As String = "Test Letvarica"
Private Const TEST_VRSTA_BEZ_SORTE As String = "Test Dunja"

Private Const TEST_PREFIX As String = "TST-PRO"

' Hladnjaca lanac (modAutoHladnjaca): zasebna stanica sa JeHladnjaca="Da" (TEST_ST_ID
' to NIJE, da ostali testovi ne okinu auto-lanac) i drugi kupac za proveru izolacije
' backfill mapa po kupcu.
Private Const TEST_HLAD_ST_ID As String = "ST-HLADTEST-90001"

' Druga stanica sa DRUGACIJIM numerickim delom -- preduslov svakog dvosmernog
' dokaza kapije konteksta broja. TEST_ST_ID i TEST_HLAD_ST_ID oba daju 90001,
' pa nad njima kapija ne moze da razlikuje vlasnike niza.
Private Const BKTX_ST2 As String = "ST-BKTX-90007"
Private Const TEST_KUP2_ID As String = "KUP-90002"

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub RunBusinessFlowProSuite()
    On Error GoTo EH

    BeginRun "BUSINESS FLOW PROFESSIONAL SUITE"

    Test_CoreTablesAndColumnsExist
    SeedBusinessFlowProMasterData
    Test_SeedMasterDataAvailable

    Test_DuplicateFakturaIsBlocked
    Test_InvalidSavesDoNotAppend
    ' OBRISANI u Otkup cutover-u: Test_OtkupAtomicMultiClassSave i
    ' Test_OtkupClassIIAmbalaza su tvrdili PO-KLASNE redove na zaglavlju
    ' ("appends exactly two rows", KolAmbalaze po klasi, BrojZbirne po
    ' klasi) -- bas model koji se uklanja. Njihove zive tvrdnje nose:
    '   jedan header + dve stavke      Test_OTK_HeaderIStavke
    '   atomicnost dvoklasnog upisa    Test_OTK_LosaDrugaStavkaRollback
    '   ambalaza dvoklasnog dokumenta  Test_OTK_AmbalazaIdeNaDokument
    Test_OtkupInputValidationHardening
    Test_OtkupReadHelpersExcludeStornirano
    Test_DokumentaInputValidationHardening
    Test_MalinaVozacMirror

    ' RF-28 (MasterSync integritet -- AUD-041/042/043)
    Test_RF28_BrojZbirneRupaNeDajeDuplikat
    Test_ZBR_ImportDvaUredjajaNeStapaDokumente
    Test_RF28_LinkKonfliktNePrepisuje
    Test_RF28_MembershipKoristiSvojuZbirnu
    Test_RF28_MembershipDanskiProzor
    Test_RF28_NevalidanDatumJeSyncError
    Test_RF28_VozacIDUpdateIshodi

    ' RF-05 (frmDokumenta unos + storno set)
    Test_ManjakPreviewJeZbirnaMinusPrijem
    Test_OpenFaktureExcludeStornirano
    Test_PrefillBiraPoslednjuGeneraciju
    Test_GeneracijaNePrelaziVlasnika
    Test_StornoPoBrojuOdbijaDvaVlasnika
    Test_ZBR_PaletaNasledjujeGeneracijuPrijemnice
    Test_ZBR_MasterSyncNePrepisujeGeneracijuDeteta
    Test_ZBR_KapijaPustaKadJeIzborScoped
    Test_MalinaAutoZbirnaFailSignal
    Test_ZbirnaRowDataColumnMapped
    Test_OMUlazSmerObavezan
    Test_PorukeKatalogPokrivaDokumenta


    Test_NoCrossZbirnaLinksAudit

    ' PR3 -- Zbirna: header + stavke.  Nov pisac jos nema nijednog pozivaoca;
    ' cutover citalaca, invarijante i storna je Zbirna cutover.
    Test_PR3_CreateZbirnaHeaderIStavke
    Test_PR3_DveOtpremniceIsteKlaseSeSabiraju
    Test_PR3_HeaderNeNosiKolicinu
    Test_PR3_ZbirnaIDJeOpaque
    Test_PR3_PrazanHeaderIDNeProlazi
    Test_PR3_PrazanStavkaIDNeProlazi
    Test_PR3_IstaOtpremnicaDvaputNeProlazi
    Test_PR3_VecVezanaOtpremnicaSeNePreuzima
    Test_PR3_DvaAktivnaClanstvaSuGreska
    Test_PR3_DupliIstiZapisClanstvaJeGreska
    Test_PR3_StorniranIzvorNeOstavljaPolaDokumenta
    Test_PR3_RazlicitaVrstaNeProlazi
    Test_PR3_RazlicitVozacNeProlazi
    Test_PR3_HeaderJeEksplicitnoIzdat
    Test_PR3_ZbirnaBezIzvoraNeProlazi
    Test_PR3_NepoznatKljucUHeaderuPada
    Test_PR3_NedostajuciObavezniKljucPada
    Test_PR3_OcekivanoKojeSeNeSlazePada
    Test_PR3_RucniUnosTraziOcekivano
    Test_PR3_AmbalazaMoraBitiCeoBroj
    Test_PR3_ClanstvoJeZapisanoPoVerziji
    Test_PR3_CitacDajeIstuZbirnuKaoClanstvo
    Test_PR3_PrazanIzvorIDNeProlazi
    Test_PR3_OtpremnicaNemaZbirnaID

    ' Otkup skela -- header + stavke. Nov pisac jos nema pozivaoca;
    ' citaoci, ambalaza i novac su Otkup cutover.
    Test_OTK_HeaderIStavke
    Test_OTK_HeaderNeNosiLinePolja
    Test_OTK_KulturaSeNeFabrikuje
    Test_OTK_KulturaSeMoraSlagatiSaVrstom
    Test_OTK_ParcelaPripadaKooperantu
    Test_OTK_BrutoINetoSuZamrznuti
    Test_OTK_CenaJeStvarnoPrimenjena
    Test_OTK_DuplaKlasaPada
    Test_OTK_LosaDrugaStavkaRollback
    Test_OTK_PrazanIDFailClosed
    Test_OTK_KolAmbIzdataJeHeader
    Test_OTK_NepoznatKljucUHeaderuPada
    Test_OTK_BezStavkiNeProlazi
    Test_OTK_AmbalazaMoraBitiCeoBroj
    Test_OTK_NepoznatKljucUStavciPada
    Test_OTK_KooperantMoraPostojati
    Test_OTK_StanicaMoraPostojati
    Test_OTK_RedosledKlasaJeKanonski
    Test_OTK_SamoKlasaII
    Test_OTK_SortaPraznaSamoUzKulturuBezSorte
    Test_OTK_TipAmbalazeVezujeSvakaAmbalaza
    Test_OTK_AmbalazaIdeNaDokument
    Test_OTK_OdbijenDokumentNeKnjiziAmbalazu
    Test_OTK_EkranPiseNovimModelom
    Test_OTK_EkranNerazresivaKulturaPada
    Test_OTK_EkranPauziraAutoLanac
    Test_OTK_PrintNetoUnosNeRekonstruiseBruto
    Test_OTK_IspravkaPauziranaNeTrosiPending
    Test_OTK_BrojJedinstvenPoStaniciIDanu
    Test_OTK_VrednostBezStavkiPada
    Test_OTK_VrednostPunUgovor
    Test_OTK_StatusIsplateJeIzveden
    Test_OTK_IspravkaNovDokumentINovBroj
    Test_OTK_IspravkaKapije
    Test_OTK_IspravkaNeGubiNovac
    Test_OTK_IspravkaPrenosiKes
    Test_OTK_IspravkaPrijavljujePreplatu
    Test_OTK_IspravkaPrijavljujePreplatuVirmanom
    Test_OTK_IsplataNaNovDokumentProlazi
    Test_OTK_CitaociCitajuStavke
    Test_OTK_CitaociStavkiFailClosed
    Test_OTK_ZaglavljeBezStavkiObaraCitaoce
    Test_OTK_StavkaBezZaglavljaObaraCitaoce
    Test_OTK_ZaglavljeBezIDObaraCitaoce
    Test_OTK_DupliOtkupIDObaraCitaoce
    Test_BIM_NovOtkupJeOtvorenBlok
    Test_PWA_StanicaJeUredjajNeKooperant
    Test_BKTX_NekanonskiBrojNeOdbija
    Test_BKTX_VlasnikOsaOdbijaTudjuStanicu
    Test_BKTX_DanOsaOdbijaTudjiDan
    Test_BKTX_GeneratorUvekProlaziKapiju
    Test_BKTX_MirrorSPrefiksProlazi
    Test_BKTX_ZbirnaTudjegVlasnikaOdbijena
    Test_BKTX_PrijemnicaNikadNeOdbija
    Test_BKTX_ReversSudiSamoAmbalazu
    Test_BKTX_RucniRezimNeGasiPravilo
    Test_BKTX_UvozOtkupaOdbijaTudjBroj
    Test_BKTX_UvozZbirneOdbijaTudjBroj
    Test_OTP_BrojZauzetPoStaniciIDanu
    Test_OTP_DraftBrojIzuzimaSebe
    Test_ZBR_StorniranBrojIstogVozacaOdbijen
    Test_BKTX_ReversPisacOdbijaZauzet
    Test_BKTX_ReversKoopIstiBrojDveStanice
    Test_BKTX_ReversIDNaSvimNogama
    Test_BKTX_ReversIDJedanDokument
    Test_PWA_BezUredjajaUvozPada
    Test_OTK_OdvezanVirmanJeRaspolozivAvans
    Test_OTK_IspravkaRoditeljFailClosed
    Test_OTK_IspravkaRollbackVracaSve
    Test_OTK_SelfHealMigracijeKolona
    Test_OTK_BrojStorniranogSeNePonovoKoristi
    Test_OTK_EkranIPisacImajuIstoPravilo
    Test_OTK_StornoJednimID
    Test_OTK_PrefillStornaDveKlaseSaStavki
    Test_OTK_PrefillStornaBezStavkiPada
    Test_OTK_IzvozDveKlaseIzStavki
    Test_OTK_IzvozBezStavkiPada
    Test_OTK_PushStavkiIdempotentan

    ' PWA ingest -- produkcioni put od Otkup cutover-a. RunMasterSyncSmokeSuite
    ' je zatecena crvena (9/26) i nije u FULL prolazu, pa pokrice mora ovde.
    Test_PWA_IngestPraviHeaderIStavku
    Test_PWA_NerazresivaKulturaObaraUvoz
    Test_PWA_IstiCridIstiSadrzajJeNoOp
    Test_PWA_IstiCridDrugiSadrzajPada
    Test_PWA_RazresivacImenujeRazlog
    Test_PWA_KonfliktPoParceliITipu
    Test_PWA_PrenosiVremeNastanka
    Test_PWA_IzvedeniLanacJePauziran

    ' Otpremnica skela -- header + stavke + clanstvo. Izvori su otkupi po
    ' NOVOM modelu, pa ovi testovi mere i da se dva nova pisca slazu.
    Test_OTP_JedanBrojJedanHeader
    Test_OTP_PredlogCeneJePoKlasi
    Test_OTP_AmbalazaSeKnjiziPriIzdavanju
    Test_OTP_F2OtvaraNacrt
    Test_OTP_MalinaZbirnaPauzirana
    Test_OTP_NacrtNijeZavrsetakIspravke
    Test_OTP_MrezaCitaStavke
    Test_OTP_ZaglavljeBezStavkiObaraCitaoce
    Test_OTP_IzvestajOMRedPoKlasi
    Test_OTP_OtpremljenoJeSamoIzdato
    Test_OTP_VrednostIzIzvoraNePredlogCene
    Test_OTP_IzdatoStatusPravilo
    Test_OTP_DveStavkeIsteKlaseObaraCitaoce
    Test_OTP_InvarijantaSabiraStavke
    Test_OTP_PrefillIspravkeCitaStavke
    Test_OTP_StavkeSuIzvedene
    Test_OTP_HeaderNeNosiLinePolja
    Test_OTP_NepoznatKljucUHeaderuPada
    Test_OTP_HeaderFKovi
    Test_OTP_DraftNosiOcekivanje
    Test_OTP_NapredakPoKlasi
    Test_OTP_IzdavanjeTraziJednakost
    Test_OTP_IzdavanjeRevalidiraIzvore
    Test_OTP_BrutoSeNeSabiraParcijalno
    Test_OTP_KulturaSeSlaziSaIzvorima
    Test_OTP_UpdateDraftaMenjaOcekivanje
    Test_OTP_DraftNothingOcekivanjePada
    Test_OTP_UpdateStaniceSaPostojecimIzvoromPada
    Test_OTP_UpdateKultureSaPostojecimIzvoromPada
    Test_OTP_DvaTipaAmbalazeNeUlazeUDraft
    Test_OTP_NeizdatOtkupNeUlazi
    Test_OTP_TipAmbalazeJeHeaderCinjenica
    Test_OTP_IzvorBezGajbiNeOdredjujeTip
    Test_OTP_ClanstvoNaNepostojeciOtkupPada
    Test_OTP_DupliParUClanstvuPada
    Test_OTP_ClanstvoMutabilnoUDraftu
    Test_OTP_PosleIzdavanjaClanstvoZamrznuto
    Test_OTP_IzvorNeSmeDvaPutaAktivno
    Test_OTP_StorniranIzvorNeUlazi
    Test_OTP_DveStaniceNeProlaze
    Test_OTP_BezIzvoraNeProlazi
    Test_OTP_StariOtkupNeUlazi
    Test_OTP_PrazanIDFailClosed
    Test_OTP_OtkupOtpremnicaIDNetaknut

    On Error GoTo 0        ' verdikt podize EndRun -- bez ovoga bi skocio u EH i dvaput brojao
    EndRun
    Exit Sub

EH:
    LogFatal "RunBusinessFlowProSuite", Err.Number, Err.description
    EndRun
End Sub

Public Sub RunBusinessFlowProSeedOnly()
    On Error GoTo EH

    BeginRun "BUSINESS FLOW PRO SEED ONLY"

    Test_CoreTablesAndColumnsExist
    SeedBusinessFlowProMasterData
    Test_SeedMasterDataAvailable

    On Error GoTo 0        ' verdikt podize EndRun -- bez ovoga bi skocio u EH i dvaput brojao
    EndRun
    Exit Sub

EH:
    LogFatal "RunBusinessFlowProSeedOnly", Err.Number, Err.description
    EndRun
End Sub

Public Sub RunBusinessFlowProTraceabilityOnly()
    On Error GoTo EH

    BeginRun "BUSINESS FLOW PRO TRACEABILITY ONLY"

    Test_CoreTablesAndColumnsExist
    SeedBusinessFlowProMasterData
    Test_NoCrossZbirnaLinksAudit

    On Error GoTo 0        ' verdikt podize EndRun -- bez ovoga bi skocio u EH i dvaput brojao
    EndRun
    Exit Sub

EH:
    LogFatal "RunBusinessFlowProTraceabilityOnly", Err.Number, Err.description
    EndRun
End Sub

Public Sub RunBusinessFlowProAuditOnly()
    On Error GoTo EH

    BeginRun "BUSINESS FLOW PRO AUDIT ONLY"

    Test_CoreTablesAndColumnsExist
    Test_NoCrossZbirnaLinksAudit

    On Error GoTo 0        ' verdikt podize EndRun -- bez ovoga bi skocio u EH i dvaput brojao
    EndRun
    Exit Sub

EH:
    LogFatal "RunBusinessFlowProAuditOnly", Err.Number, Err.description
    EndRun
End Sub

' ============================================================
' CORE TESTS
' ============================================================

Private Sub Test_CoreTablesAndColumnsExist()
    On Error GoTo EH

    RequireTableExists TBL_STANICE
    RequireTableExists TBL_KOOPERANTI
    RequireTableExists TBL_VOZACI
    RequireTableExists TBL_KUPCI
    RequireTableExists TBL_KULTURE

    RequireTableExists TBL_OTKUP
    RequireTableExists TBL_OTPREMNICA
    RequireTableExists TBL_ZBIRNA
    RequireTableExists TBL_PRIJEMNICA
    RequireTableExists TBL_FAKTURE
    RequireTableExists TBL_FAKTURA_STAVKE
    RequireTableExists TBL_AMBALAZA
    RequireTableExists TBL_NOVAC

    RequireTableExists TBL_OTKUP_STAVKE

    ' Otkup je zaglavlje + stavke (S1d): kolicina, cena, klasa i gajbe su na stavci.
    RequireColumnsExist TBL_OTKUP, Array( _
        "OtkupID", "Datum", "KooperantID", "StanicaID", "VrstaVoca", _
        "SortaVoca", "TipAmbalaze", "VozacID", "BrojDokumenta", "BrojZbirne", _
        "OtpremnicaID")

    RequireColumnsExist TBL_OTKUP_STAVKE, Array( _
        COL_OKS_ID, COL_OKS_OTKUP_ID, COL_OKS_RB, COL_OKS_KLASA, COL_OKS_KOLICINA, _
        COL_OKS_CENA, COL_OKS_KOL_AMB, COL_OKS_BRUTO)

    RequireColumnsExist TBL_OTPREMNICA, Array( _
        "OtpremnicaID", "Datum", "StanicaID", "VozacID", "BrojOtpremnice", _
        "BrojZbirne", "VrstaVoca", "SortaVoca", "Kolicina", "Cena", _
        "TipAmbalaze", "KolAmbalaze", "Klasa")

    RequireColumnsExist TBL_ZBIRNA, Array( _
        "ZbirnaID", "Datum", "VozacID", "BrojZbirne", "KupacID", _
        "VrstaVoca", "SortaVoca", "UkupnoKolicina", "TipAmbalaze", _
        "UkupnoAmbalaze", "Klasa")

    RequireColumnsExist TBL_PRIJEMNICA, Array( _
        "PrijemnicaID", "Datum", "KupacID", "VozacID", "BrojPrijemnice", _
        "BrojZbirne", "VrstaVoca", "SortaVoca", "Kolicina", "Cena", _
        "TipAmbalaze", "KolAmbalaze", "KolAmbVracena", "Klasa", _
        "Fakturisano", "FakturaID")

    RequireColumnsExist TBL_FAKTURE, Array( _
        "FakturaID", "BrojFakture", "Datum", "KupacID", "Iznos")

    RequireColumnsExist TBL_FAKTURA_STAVKE, Array( _
        "StavkaID", "FakturaID", "PrijemnicaID", "Kolicina", "Cena", _
        "Klasa", "BrojPrijemnice")

    LogPass "Core tables and required columns exist"
    Exit Sub

EH:
    LogFail "Core tables and required columns exist", Err.description
End Sub

Private Sub Test_SeedMasterDataAvailable()
    On Error GoTo EH

    AssertTrue RowExists(TBL_STANICE, "StanicaID", TEST_ST_ID), "Seed station exists"
    AssertTrue RowExists(TBL_KOOPERANTI, "KooperantID", TEST_KOOP_ID), "Seed kooperant exists"
    AssertTrue RowExists(TBL_VOZACI, "VozacID", TEST_VOZ_ID), "Seed vozac exists"
    AssertTrue RowExists(TBL_KUPCI, "KupacID", TEST_KUP_ID), "Seed kupac exists"
    AssertTrue RowExists(TBL_KULTURE, "KulturaID", TEST_KULTURA_ID), "Seed kultura exists"

    If Not GetTable(TBL_PARCELE) Is Nothing Then
        AssertTrue RowExists(TBL_PARCELE, "ParcelaID", TEST_PAR_ID), "Seed parcela exists"
    Else
        LogSkip "Seed parcela exists", "tblParcele not found"
    End If

    Exit Sub

EH:
    LogFail "Seed master data available", Err.description
End Sub

Private Sub Test_DuplicateFakturaIsBlocked()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("DUPFAK")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    Dim brojPrij As String

    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-" & scenario

    ' Minimal prijemnica fixture for faktura duplicate test.
    ' Zbirna mora da postoji pre prijemnice (PRIJEMNICA_ZBIRNA_PROVERA guard).
    Dim zbrFix As String
    zbrFix = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                           "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                           100#, TEST_TIP_AMB, 0, "I")
    AssertTrue Len(zbrFix) > 0, "Duplicate faktura fixture zbirna created"

    Dim prjI As String
    prjI = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 0, 0, "I")

    AssertTrue Len(prjI) > 0, "Duplicate faktura fixture prijemnica created"

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add Array(prjI, 100#, 100#, "I", brojPrij)

    Dim beforeFak As Long
    beforeFak = CountRows(TBL_FAKTURE)

    Dim f1 As String
    f1 = CreateFaktura_TX(TEST_KUP_ID, stavke)

    AssertTrue Len(f1) > 0, "First faktura for duplicate test created"
    AssertEquals CStr(beforeFak + 1), CStr(CountRows(TBL_FAKTURE)), "First faktura increments count"

    Dim f2 As String
    On Error Resume Next
    f2 = CreateFaktura_TX(TEST_KUP_ID, stavke)

    If Err.Number <> 0 Then
        LogPass "Duplicate faktura attempt raises/blocks"
        Err.Clear
        On Error GoTo EH
    Else
        On Error GoTo EH
        AssertTrue Len(Trim$(f2)) = 0, "Duplicate faktura attempt returns empty"
    End If

    AssertEquals CStr(beforeFak + 1), CStr(CountRows(TBL_FAKTURE)), _
                 "Duplicate faktura did not append second faktura"

    Exit Sub

EH:
    LogFail "Duplicate faktura is blocked", Err.description
End Sub

Private Sub Test_InvalidSavesDoNotAppend()
    On Error GoTo EH

    Test_InvalidOtkupDoesNotAppend
    Test_InvalidPrijemnicaDoesNotAppend

    Exit Sub

EH:
    LogFail "Invalid saves do not append", Err.description
End Sub

Private Sub Test_InvalidOtkupDoesNotAppend()
    On Error GoTo ExpectedError

    Dim beforeCount As Long
    beforeCount = CountRows(TBL_OTKUP)

    ' Prazan kooperantID treba da blokira
    Dim result As String
    ' Prazan KooperantID -> kanonski pisac ga odbija kao FK (S4.1f).
    Dim badH As Object
    Set badH = OtkHeader(TEST_PREFIX & "-BAD-OTK-" & NewScenarioCode("BAD"))
    badH("KooperantID") = ""
    result = CreateOtkup_TX(badH, OtkStavke(100#, 100#, 0, 0#, 0#, 0))

    If Len(Trim$(result)) = 0 Then
        AssertEquals CStr(beforeCount), CStr(CountRows(TBL_OTKUP)), _
                     "Invalid otkup did not append row"
        Exit Sub
    End If

    LogFail "Invalid otkup rejected", "CreateOtkup_TX returned ID: " & result
    Exit Sub

ExpectedError:
    AssertEquals CStr(beforeCount), CStr(CountRows(TBL_OTKUP)), _
                 "Invalid otkup raised and did not append row"
End Sub

Private Sub Test_InvalidPrijemnicaDoesNotAppend()
    On Error GoTo ExpectedError

    Dim beforeCount As Long
    beforeCount = CountRows(TBL_PRIJEMNICA)

    Dim result As String
    result = SavePrijemnica_TX(NextTestDate(), "", TEST_VOZ_ID, TEST_PREFIX & "-BAD-PRJ-" & NewScenarioCode("BAD"), _
                               TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("BAD"), _
                               TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 1, 1, "I")

    If Len(Trim$(result)) = 0 Then
        AssertEquals CStr(beforeCount), CStr(CountRows(TBL_PRIJEMNICA)), _
                     "Invalid prijemnica did not append row"
        Exit Sub
    End If

    LogFail "Invalid prijemnica rejected", "SavePrijemnica_TX returned ID: " & result
    Exit Sub

ExpectedError:
    AssertEquals CStr(beforeCount), CStr(CountRows(TBL_PRIJEMNICA)), _
                 "Invalid prijemnica raised and did not append row"
End Sub

Private Sub Test_OtkupInputValidationHardening()
    On Error GoTo EH

    Test_InvalidOtkupNegativeCenaDoesNotAppend
    Test_InvalidOtkupInvalidClassDoesNotAppend

    Exit Sub

EH:
    LogFail "Otkup input validation hardening", Err.description
End Sub

Private Sub Test_InvalidOtkupNegativeCenaDoesNotAppend()
    On Error GoTo EH

    Dim beforeOtkup As Long
    beforeOtkup = CountRows(TBL_OTKUP)

    Dim result As String
    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add OtkStavka(KLASA_I, 100#, -1#, 1, 0#)

    Dim beforeStavke As Long
    beforeStavke = CountRows(TBL_OTKUP_STAVKE)

    Dim razlog As String
    result = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-BAD-OTK-" & NewScenarioCode("NEGPRICE")), _
                            stavke, razlog)
    AssertTrue InStr(1, razlog, "Cena mora biti veca od nule", vbTextCompare) > 0, _
               "Invalid otkup negative cena: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(beforeStavke), CStr(CountRows(TBL_OTKUP_STAVKE)), _
                 "Invalid otkup negative cena did not append stavka"

    AssertEquals "", result, "Invalid otkup negative cena returns empty"
    AssertEquals CStr(beforeOtkup), CStr(CountRows(TBL_OTKUP)), _
                 "Invalid otkup negative cena did not append row"

    Exit Sub

EH:
    LogFail "Invalid otkup negative cena", Err.description
End Sub

Private Sub Test_InvalidOtkupInvalidClassDoesNotAppend()
    On Error GoTo EH

    Dim beforeOtkup As Long
    beforeOtkup = CountRows(TBL_OTKUP)

    Dim result As String
    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add OtkStavka("BAD", 100#, 10#, 1, 0#)

    Dim beforeStavke As Long
    beforeStavke = CountRows(TBL_OTKUP_STAVKE)

    Dim razlog As String
    result = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-BAD-OTK-" & NewScenarioCode("BADCLASS")), _
                            stavke, razlog)
    AssertTrue Len(razlog) > 0, "Invalid otkup class: kapija vraca razlog"
    AssertEquals CStr(beforeStavke), CStr(CountRows(TBL_OTKUP_STAVKE)), _
                 "Invalid otkup class did not append stavka"

    AssertEquals "", result, "Invalid otkup class returns empty"
    AssertEquals CStr(beforeOtkup), CStr(CountRows(TBL_OTKUP)), _
                 "Invalid otkup class did not append row"

    Exit Sub

EH:
    LogFail "Invalid otkup invalid class", Err.description
End Sub

Private Sub Test_DokumentaInputValidationHardening()
    On Error GoTo EH

    Test_InvalidZbirnaInvalidClassDoesNotAppend
    Test_InvalidPrijemnicaNegativeAmbalazaDoesNotAppend
    Test_PrijemnicaMissingZbirnaDoesNotAppend

    Exit Sub

EH:
    LogFail "Dokumenta input validation hardening", Err.description
End Sub

Private Sub Test_InvalidZbirnaInvalidClassDoesNotAppend()
    On Error GoTo EH

    Dim beforeCount As Long
    beforeCount = CountRows(TBL_ZBIRNA)

    Dim result As String
    result = SaveZbirna_TX( _
        NextTestDate(), TEST_VOZ_ID, _
        TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("BADCLASS"), _
        TEST_KUP_ID, "Test Hladnjaca", "Test Pogon", _
        TEST_VRSTA, TEST_SORTA, _
        100#, TEST_TIP_AMB, 1, "BAD")

    AssertEquals "", result, "Invalid zbirna class returns empty"
    AssertEquals CStr(beforeCount), CStr(CountRows(TBL_ZBIRNA)), _
                 "Invalid zbirna class did not append row"

    Exit Sub

EH:
    LogFail "Invalid zbirna invalid class", Err.description
End Sub

Private Sub Test_InvalidPrijemnicaNegativeAmbalazaDoesNotAppend()
    On Error GoTo EH

    Dim beforePrj As Long
    Dim beforeAmb As Long

    beforePrj = CountRows(TBL_PRIJEMNICA)
    beforeAmb = CountRows(TBL_AMBALAZA)

    ' Validna zbirna mora da postoji -> jedini razlog odbijanja je negativna
    ' ambalaza (a ne PRIJEMNICA_ZBIRNA_PROVERA guard, koji bi inace prekinuo pre).
    Dim scenario As String: scenario = NewScenarioCode("NEGAMB")
    Dim testDate As Date: testDate = NextTestDate()
    Dim brojZbirne As String: brojZbirne = TEST_PREFIX & "-BAD-ZBR-" & scenario

    Dim zbrFix As String
    zbrFix = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                           "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                           100#, TEST_TIP_AMB, 0, KLASA_I)
    AssertTrue Len(zbrFix) > 0, "Negative ambalaza fixture zbirna created"

    Dim result As String
    result = SavePrijemnica_TX( _
        testDate, TEST_KUP_ID, TEST_VOZ_ID, _
        TEST_PREFIX & "-BAD-PRJ-" & scenario, _
        brojZbirne, _
        TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, -1, 0, KLASA_I)

    AssertEquals "", result, "Invalid prijemnica negative ambalaza returns empty"
    AssertEquals CStr(beforePrj), CStr(CountRows(TBL_PRIJEMNICA)), _
                 "Invalid prijemnica negative ambalaza did not append prijemnica"
    AssertEquals CStr(beforeAmb), CStr(CountRows(TBL_AMBALAZA)), _
                 "Invalid prijemnica negative ambalaza did not append ambalaza"

    Exit Sub

EH:
    LogFail "Invalid prijemnica negative ambalaza", Err.description
End Sub

Private Sub Test_PrijemnicaMissingZbirnaDoesNotAppend()
    On Error GoTo EH

    ' PRIJEMNICA_ZBIRNA_PROVERA guard: u BLOK modu prijemnica sa nepostojecom
    ' zbirnom mora biti odbijena (referencijalni integritet, bez orphan reda).
    Dim prevMode As String
    prevMode = GetConfigValue(CFG_PRIJEMNICA_ZBIRNA_PROVERA)
    SetConfigValue CFG_PRIJEMNICA_ZBIRNA_PROVERA, "BLOK"

    Dim beforePrj As Long
    beforePrj = CountRows(TBL_PRIJEMNICA)

    Dim scenario As String
    scenario = NewScenarioCode("NOZBR")

    Dim result As String
    result = SavePrijemnica_TX( _
        NextTestDate(), TEST_KUP_ID, TEST_VOZ_ID, _
        TEST_PREFIX & "-PRJ-" & scenario, _
        TEST_PREFIX & "-ZBR-MISSING-" & scenario, _
        TEST_VRSTA, TEST_SORTA, _
        100#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)

    AssertEquals "", result, "Prijemnica with missing zbirna is rejected (BLOK)"
    AssertEquals CStr(beforePrj), CStr(CountRows(TBL_PRIJEMNICA)), _
                 "Rejected prijemnica did not append a row"

    SetConfigValue CFG_PRIJEMNICA_ZBIRNA_PROVERA, prevMode
    Exit Sub

EH:
    SetConfigValue CFG_PRIJEMNICA_ZBIRNA_PROVERA, prevMode
    LogFail "Prijemnica without zbirna is blocked", Err.description
End Sub

' Fixture je kanonski otkup (CreateOtkup_TX kroz NoviOtkupFixture); test se
' vraca rollback-om jer markira storno i zigose BrojZbirne.
Private Sub Test_OtkupReadHelpersExcludeStornirano()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKSTO")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojActive As String
    Dim brojStorno As String
    Dim brojZbirne As String

    brojActive = TEST_PREFIX & "-OTK-A-" & scenario
    brojStorno = TEST_PREFIX & "-OTK-S-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim activeID As String
    Dim stornoID As String

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    ' CreateOtkup_TX zove ApplyAvansToOtkup: fixture sme da potrosi ili
    ' podeli slobodan avans kooperanta, pa i tblNovac mora nazad.
    tx.AddTableSnapshot TBL_NOVAC

    activeID = NoviOtkupFixture(testDate, TEST_ST_ID, brojActive, brojZbirne, _
                                100#, 10#, 1, 0#, 0#, 0)

    stornoID = NoviOtkupFixture(testDate, TEST_ST_ID, brojStorno, brojZbirne, _
                                200#, 10#, 1, 0#, 0#, 0)

    AssertTrue Len(activeID) > 0 And Len(stornoID) > 0, _
               "Otkup storno filter fixture rows created"

    MarkTestRowStornirano TBL_OTKUP, "OtkupID", stornoID

    AssertFalse ArrayContainsKeyValue(GetOtkupByStation(TEST_ST_ID, testDate, testDate), _
                                      TBL_OTKUP, "OtkupID", stornoID), _
                "GetOtkupByStation excludes stornirano"

    AssertFalse ArrayContainsKeyValue(GetOtkupByKooperant(TEST_KOOP_ID, testDate, testDate), _
                                      TBL_OTKUP, "OtkupID", stornoID), _
                "GetOtkupByKooperant excludes stornirano"

    tx.RollbackTx
    Set tx = Nothing
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "Otkup read helpers exclude stornirano", errDesc
End Sub

' ============================================================
' TRACEABILITY / AUTOLINK REGRESSION TESTS
' ============================================================

Private Sub Test_NoCrossZbirnaLinksAudit()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        LogSkip "Cross-zbirna link audit", "tblOtkup empty"
        Exit Sub
    End If

    Dim colOtkID As Long
    Dim colOtkZbr As Long
    Dim colOtkOtp As Long

    colOtkID = RequireCol(TBL_OTKUP, "OtkupID")
    colOtkZbr = RequireCol(TBL_OTKUP, "BrojZbirne")
    colOtkOtp = RequireCol(TBL_OTKUP, "OtpremnicaID")

    Dim badCount As Long
    Dim details As String

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim otkID As String
        Dim otkZbr As String
        Dim otpID As String
        Dim otpZbr As String

        otkID = Trim$(CStr(data(i, colOtkID)))
        otkZbr = Trim$(CStr(data(i, colOtkZbr)))
        otpID = Trim$(CStr(data(i, colOtkOtp)))

        If Len(otpID) > 0 Then
            otpZbr = Trim$(CStr(GetValueByKey(TBL_OTPREMNICA, "OtpremnicaID", otpID, "BrojZbirne")))

            If Len(otkZbr) > 0 And Len(otpZbr) > 0 Then
                If otkZbr <> otpZbr Then
                    badCount = badCount + 1
                    details = details & otkID & " -> " & otpID & _
                              " | Otkup.BrojZbirne=" & otkZbr & _
                              " | Otp.BrojZbirne=" & otpZbr & vbCrLf
                End If
            End If
        End If
    Next i

    If badCount = 0 Then
        LogPass "Cross-zbirna link audit found no mismatches"
    Else
        LogFail "Cross-zbirna link audit found " & badCount & " mismatch(es)", details
    End If

    Exit Sub

EH:
    LogFail "Cross-zbirna link audit", Err.description
End Sub

' ============================================================
' OPTIONAL CLEANUP
' ============================================================

Public Sub SoftStornoBusinessFlowTestRows()
    On Error GoTo EH

    BeginRun "SOFT STORNO BUSINESS FLOW TEST ROWS"

    SoftStornoByTestMarkers TBL_OTKUP, Array("BrojDokumenta", "BrojZbirne")
    SoftStornoByTestMarkers TBL_OTPREMNICA, Array("BrojOtpremnice", "BrojZbirne")
    SoftStornoByTestMarkers TBL_ZBIRNA, Array("BrojZbirne")
    SoftStornoByTestMarkers TBL_PRIJEMNICA, Array("BrojPrijemnice", "BrojZbirne")
    SoftStornoByTestMarkers TBL_FAKTURE, Array("BrojFakture")
    SoftStornoByTestMarkers TBL_FAKTURA_STAVKE, Array("BrojPrijemnice")

    EndRun
    Exit Sub

EH:
    LogFatal "SoftStornoBusinessFlowTestRows", Err.Number, Err.description
    EndRun
End Sub

Private Sub SoftStornoByTestMarkers(ByVal tableName As String, ByVal markerColumns As Variant)
    On Error GoTo EH

    If GetTable(tableName) Is Nothing Then
        LogSkip "Soft-storno " & tableName, "Table not found"
        Exit Sub
    End If

    If GetColumnIndex(tableName, "Stornirano") = 0 Then
        LogSkip "Soft-storno " & tableName, "No Stornirano column"
        Exit Sub
    End If

    Dim data As Variant
    data = GetTableData(tableName)

    If IsEmpty(data) Then
        LogSkip "Soft-storno " & tableName, "No rows"
        Exit Sub
    End If

    Dim changed As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If RowHasTestMarker(data, i, tableName, markerColumns) Then
            RequireUpdateCell tableName, i, "Stornirano", "Da", "modBusinessFlowProTests.SoftStornoByTestMarkers"
            changed = changed + 1
        End If
    Next i

    LogPass "Soft-storno " & tableName & " changed " & changed & " row(s)"
    Exit Sub

EH:
    LogFail "Soft-storno " & tableName, Err.description
End Sub

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RowHasTestMarker(ByRef data As Variant, ByVal rowIndex As Long, _
                                  ByVal tableName As String, ByVal markerColumns As Variant) As Boolean
    Dim c As Variant

    For Each c In markerColumns
        Dim colIdx As Long
        colIdx = GetColumnIndex(tableName, CStr(c))

        If colIdx > 0 Then
            If InStr(1, CStr(data(rowIndex, colIdx)), TEST_PREFIX, vbTextCompare) > 0 Then
                RowHasTestMarker = True
                Exit Function
            End If
        End If
    Next c
End Function

' ============================================================
' SEED DATA
' ============================================================

Private Sub SeedBusinessFlowProMasterData()
    On Error GoTo EH

    SeedStanica
    SeedHladnjacaStanica
    SeedVozac
    SeedVozac2
    SeedKupac
    SeedKupac2
    SeedKultura
    SeedKulturaBezSorte
    SeedKooperant
    SeedKooperant2
    SeedParcelaIfAvailable

    LogPass "Seed master data ready"
    Exit Sub

EH:
    LogFail "Seed master data", Err.description
End Sub

Private Sub SeedStanica()
    If RowExists(TBL_STANICE, "StanicaID", TEST_ST_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_STANICE)

    SetRequiredField rowData, TBL_STANICE, "StanicaID", TEST_ST_ID
    SetRequiredField rowData, TBL_STANICE, "Naziv", "TEST STANICA"
    SetOptionalField rowData, TBL_STANICE, "Mesto", "Test Mesto"
    SetOptionalField rowData, TBL_STANICE, "Kontakt", "Test Kontakt"
    SetOptionalField rowData, TBL_STANICE, "Telefon", "060000000"
    SetOptionalField rowData, TBL_STANICE, "Aktivan", "Aktivan"
    SetOptionalField rowData, TBL_STANICE, "Ime", "Test"
    SetOptionalField rowData, TBL_STANICE, "Prezime", "Stanica"
    SetOptionalField rowData, TBL_STANICE, "PIN", "9001"

    RequireAppend TBL_STANICE, rowData, "SeedStanica"
End Sub

' Idempotentan seed stanice po zadatom ID-u. AUD-046: mirror/stamp testovi vise ne
' smeju da rade sa StanicaID-em koji ne postoji u tblStanice.
Private Sub SeedStanicaByID(ByVal stanicaID As String, ByVal naziv As String)
    If RowExists(TBL_STANICE, "StanicaID", stanicaID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_STANICE)

    SetRequiredField rowData, TBL_STANICE, "StanicaID", stanicaID
    SetRequiredField rowData, TBL_STANICE, "Naziv", naziv
    SetOptionalField rowData, TBL_STANICE, "Mesto", "Test Mesto"
    SetOptionalField rowData, TBL_STANICE, "Kontakt", "Test Kontakt"
    SetOptionalField rowData, TBL_STANICE, "Aktivan", "Aktivan"

    RequireAppend TBL_STANICE, rowData, "SeedStanicaByID"
End Sub

' Stanica oznacena kao hladnjaca -> IsHladnjacaStanica = True (auto-lanac).
Private Sub SeedHladnjacaStanica()
    If RowExists(TBL_STANICE, "StanicaID", TEST_HLAD_ST_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_STANICE)

    SetRequiredField rowData, TBL_STANICE, "StanicaID", TEST_HLAD_ST_ID
    SetRequiredField rowData, TBL_STANICE, "Naziv", "TEST HLADNJACA STANICA"
    SetOptionalField rowData, TBL_STANICE, "Mesto", "Test Mesto"
    SetOptionalField rowData, TBL_STANICE, "Kontakt", "Test Kontakt"
    SetOptionalField rowData, TBL_STANICE, "Aktivan", "Aktivan"
    SetOptionalField rowData, TBL_STANICE, "Ime", "Test"
    SetOptionalField rowData, TBL_STANICE, "Prezime", "Hladnjaca"
    SetOptionalField rowData, TBL_STANICE, "PIN", "9011"
    SetOptionalField rowData, TBL_STANICE, COL_STA_JE_HLADNJACA, "Da"

    RequireAppend TBL_STANICE, rowData, "SeedHladnjacaStanica"
End Sub

' Drugi kupac -- koristi se samo da dokaze da backfill mape ignorisu prijemnice
' koje ne pripadaju hladnjaca-kupcu.
Private Sub SeedKupac2()
    If RowExists(TBL_KUPCI, "KupacID", TEST_KUP2_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_KUPCI)

    SetRequiredField rowData, TBL_KUPCI, "KupacID", TEST_KUP2_ID
    SetRequiredField rowData, TBL_KUPCI, "Naziv", "TEST KUPAC DVA DOO"
    SetOptionalField rowData, TBL_KUPCI, "Mesto", "Test Grad"
    SetRequiredField rowData, TBL_KUPCI, "PIB", "109000002"
    SetOptionalField rowData, TBL_KUPCI, "MaticniBroj", "20900002"
    SetOptionalField rowData, TBL_KUPCI, "Ulica", "Test ulica 2"
    SetOptionalField rowData, TBL_KUPCI, "PostanskiBroj", "11000"
    SetOptionalField rowData, TBL_KUPCI, "Drzava", "RS"
    SetOptionalField rowData, TBL_KUPCI, "Hladnjaca", "Test Hladnjaca 2"
    SetOptionalField rowData, TBL_KUPCI, "Aktivan", "Aktivan"
    SetOptionalField rowData, TBL_KUPCI, "TekuciRacun", "160-0000000000002-00"

    RequireAppend TBL_KUPCI, rowData, "SeedKupac2"
End Sub

' Drugi vozac. Koristi ga PR3 (zbirna nosi otpremnice SAMO svog vozaca) i
' izmena drafta otpremnice. Do sada NIJE bio zasejan: stari pisac otpremnice
' nema FK proveru, pa je nepostojeci VozacID prolazio neprimetno.
Private Sub SeedVozac2()
    If RowExists(TBL_VOZACI, "VozacID", TEST_VOZ_ID_B) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_VOZACI)

    SetRequiredField rowData, TBL_VOZACI, "VozacID", TEST_VOZ_ID_B
    SetRequiredField rowData, TBL_VOZACI, "Ime", "Test"
    SetRequiredField rowData, TBL_VOZACI, "Prezime", "Vozac Drugi"
    SetOptionalField rowData, TBL_VOZACI, "Aktivan", "Aktivan"

    RequireAppend TBL_VOZACI, rowData, "SeedVozac2"
End Sub

Private Sub SeedVozac()
    If RowExists(TBL_VOZACI, "VozacID", TEST_VOZ_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_VOZACI)

    SetRequiredField rowData, TBL_VOZACI, "VozacID", TEST_VOZ_ID
    SetRequiredField rowData, TBL_VOZACI, "Ime", "Test"
    SetRequiredField rowData, TBL_VOZACI, "Prezime", "Vozac"
    SetOptionalField rowData, TBL_VOZACI, "Telefon", "060000001"
    SetOptionalField rowData, TBL_VOZACI, "Aktivan", "Aktivan"
    SetOptionalField rowData, TBL_VOZACI, "PIN", "9002"
    SetOptionalField rowData, TBL_VOZACI, "KapacitetKG", 10000

    RequireAppend TBL_VOZACI, rowData, "SeedVozac"
End Sub

' ============================================================
' MALINA MOD -- vozac mirror: nova stanica -> par-vozac sa istim ID-em
' ============================================================
Private Sub Test_MalinaVozacMirror()
    Dim prevMode As String

    On Error GoTo EH

    ' Fiksni test ID -> idempotentno; ne gomila redove kroz vise run-ova suite-a.
    Const MIR_ST As String = "ST-MIRTEST-90001"

    ' AUD-046: stanica MORA da postoji da bi mirror smeo da se napravi, pa je
    ' test-stanica sada deo pripreme (ranije se Ensure zvao za nepostojeci ID).
    SeedStanicaByID MIR_ST, "TEST MIRROR STANICA"

    prevMode = GetConfigValue(CFG_KEY_MALINA_MODE)
    SetConfigValue CFG_KEY_MALINA_MODE, "YES"

    ' Posle Ensure vozac mora postojati (kreiran sad ili od ranijeg run-a).
    Call EnsureVozacMirrorForStanica(MIR_ST, "Test Naziv", "Test Mesto", "")
    AssertTrue RowExists(TBL_VOZACI, "VozacID", MIR_ST), _
        "Malina mirror: vozac VozacID==StanicaID postoji posle Ensure"

    ' Idempotencija: ponovni poziv NE sme da kreira nov red.
    AssertFalse EnsureVozacMirrorForStanica(MIR_ST, "Test Naziv", "Test Mesto", ""), _
        "Malina mirror: ponovni poziv ne kreira duplikat (idempotentno)"

    ' AUD-046: canonical par-provera vidi kompletan mirror.
    AssertTrue IsManagedStationMirror(MIR_ST), _
        "Malina mirror: IsManagedStationMirror True za kompletan par (tblStanice+tblVozaci)"

    ' AUD-046: stanica koja NE postoji nije mirror i Ensure za nju MORA da padne
    ' (ne sme da napravi vozaca bez stanice, ni da tiho vrati False).
    Const MIR_NEPOSTOJI As String = "ST-MIRTEST-NEMA-90002"

    AssertFalse IsManagedStationMirror(MIR_NEPOSTOJI), _
        "Malina mirror: IsManagedStationMirror False za nepostojecu stanicu"

    Dim raised As Boolean
    On Error Resume Next
    Call EnsureVozacMirrorForStanica(MIR_NEPOSTOJI, "X", "Y", "")
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    AssertTrue raised, _
        "Malina mirror: Ensure re-raise-uje za nepostojecu stanicu (ne guta gresku)"
    AssertFalse RowExists(TBL_VOZACI, "VozacID", MIR_NEPOSTOJI), _
        "Malina mirror: nema vozaca bez stanice (nije kreiran shadow)"

    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    Exit Sub

EH:
    On Error Resume Next
    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    On Error GoTo 0
    LogFatal "Test_MalinaVozacMirror", Err.Number, Err.description
End Sub

' ============================================================
' RF-28 -- MasterSync integritet (AUD-041/042/043) regresija
'
' Svaki test radi u sopstvenoj clsTransaction i ROLLBACK-uje se, pa fixture redovi
' ne ostaju u svesci i suite je ponovljiv. Privatne rutine se zovu kroz
' modMasterSync.TestHook_* (bez Google/HTTP zavisnosti).
'
' MSVOZ_* ishodi su Private consts u modMasterSync, pa se ovde porede kao
' literali ("UPDATED"/"NOCHANGE"/"CONFLICT"/"NOTFOUND") -- ako se preimenuju,
' ovi testovi moraju da se azuriraju zajedno sa njima.
' ============================================================

' AUD-041(b): rupa u nizu ne sme da proizvede vec zauzet broj. Row-count generator
' je za {"N/ddmmyy", "N/ddmmyy-3"} vracao "-3" ponovo; MAX-seq vraca "-4".
' ZBR-IDENT-01 / A21 (KR-001): DVA UREDJAJA, isti vozac, isti kupac, isti broj.
'
' To NIJE dvoklasna zbirna. PWA obe klase sabira u JEDAN red ("I/II"), pa dva reda
' iz uvoza znace dve odvojene terenske cinjenice -- dva dokumenta koja slucajno
' dele broj, jer su uredjaji bili offline (KR-001).
'
' Do v6-ui-224 je ImportRowToTblZbirna zvao ApplyGeneracijaID, koji generaciju
' NASLEDJUJE od aktivnog reda istog broja u istom scope-u (vozac + kupac). Drugi
' uvoz je zato dobijao generaciju prvog, pa su dve cinjenice postajale JEDAN
' identitet. Posledice koje ovaj test meri, sve odjednom:
'   activeLogicalCount broji GENERACIJE -> ostajao 1 -> resolutionStatus UNIQUE
'   -> F4 (ZbirnaRoditeljRazlog) pusta prijemnicu na spojen dokument
'   -> PrijaviKolizijuBrojaZbirne i B8 (koji presudu uzimaju od resolvera) cute.
'
' Meri se OBA smera: da uvoz nije odbijen (oba reda postoje) i da identitet nije
' stopljen. Bez prve grane bi tvrdnja bila zelena i da import blokira, sto je bas
' ono sto korak 5 nije smeo da uradi.
Private Sub Test_ZBR_ImportDvaUredjajaNeStapaDokumente()
    Dim tx As clsTransaction
    Dim testDate As Date
    Dim broj As String
    Dim idA As String, idB As String
    Dim genA As String, genB As String
    Dim ident As ZbirnaIdent

    On Error GoTo EH

    testDate = NextTestDate()
    broj = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    idA = TestHook_ImportZbirnaRowPWA("CRID-ZBRIDENT-A-" & m_RunID, TEST_VOZ_ID, _
                                      TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, _
                                      100, broj)
    idB = TestHook_ImportZbirnaRowPWA("CRID-ZBRIDENT-B-" & m_RunID, TEST_VOZ_ID, _
                                      TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, _
                                      120, broj)

    genA = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, idA)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, idB)

    ident = ZbirnaIdentResolve(broj, TEST_VOZ_ID, TEST_KUP_ID)

    Dim nalazi As Variant
    nalazi = modIntegritet.GetIntegritetRows()

    ' INGEST: nijedan red nije odbijen, uvoz nije rollback-ovan.
    AssertTrue (Len(idA) > 0 And Len(idB) > 0 And idA <> idB), _
        "A21 ingest: oba PWA reda su upisana kao zasebne zbirne"

    ' ZBR-IDENT-01: dve cinjenice, dva identiteta.
    AssertTrue (Len(genA) > 0 And Len(genB) > 0), _
        "A21: oba uvezena reda nose GeneracijaID"
    AssertTrue (genA <> genB), _
        "A21/KR-001: drugi uredjaj NE nasledjuje generaciju prvog"

    ' Detekcija je time dobila sta da vidi.
    AssertEquals "2", CStr(ident.activeLogicalCount), _
        "A21: broj nosi DVA aktivna logicka dokumenta"
    AssertEquals "1", CStr(ident.activeOwnerCount), _
        "A21/A17: dvosmislenost postoji i kod JEDNOG vlasnika"
    AssertEquals ZBR_RES_AMBIGUOUS, ident.resolutionStatus, _
        "A21: resolver kaze CURRENT_AMBIGUOUS"
    AssertEquals ZBR_PARENT_DVOSMISLEN, ZbirnaRoditeljRazlog(ident), _
        "A21: F4 fail-closed odbija taj broj"
    AssertTrue modTest.NalazSadrzi(nalazi, "B8", broj), _
        "A21: revizija integriteta (B8) prijavljuje broj"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-IDENT-01 A21 uvoz dva uredjaja", bfpErrDesc
End Sub

Private Sub Test_RF28_BrojZbirneRupaNeDajeDuplikat()
    Dim tx As clsTransaction
    Dim prevAuto As String

    On Error GoTo EH

    Dim testDate As Date
    testDate = NextTestDate()

    Dim baza As String
    baza = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")

    prevAuto = GetConfigValue(CFG_AUTO_BROJ_DOK)
    SetConfigValue CFG_AUTO_BROJ_DOK, "DA"

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    ' Niz sa rupom: postoje seq 1 i seq 3 (seq 2 obrisan/storniran).
    AppendRF28ZbirnaFixture "ZBR-RF28G1-" & m_RunID, testDate, TEST_VOZ_ID, baza
    AppendRF28ZbirnaFixture "ZBR-RF28G2-" & m_RunID, testDate, TEST_VOZ_ID, baza & "-3"

    Dim predlog As String
    predlog = TestHook_GenerateBrojZbirne(TEST_VOZ_ID, testDate)

    AssertEquals baza & "-4", predlog, _
        "RF-28 AUD-041b: rupa u nizu daje MAX+1 (-4), ne duplikat"

    tx.RollbackTx
    SetConfigValue CFG_AUTO_BROJ_DOK, prevAuto
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    SetConfigValue CFG_AUTO_BROJ_DOK, prevAuto
    On Error GoTo 0
    LogFail "RF-28 AUD-041b broj zbirne rupa", bfpErrDesc
End Sub

' AUD-043(b): otkup koji je vec u DRUGOJ zbirnoj ne sme da bude tiho prepisan.
Private Sub Test_RF28_LinkKonfliktNePrepisuje()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("RF28LNK")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim otkID As String, crid As String, zbrID As String
    Dim brojA As String, brojB As String

    otkID = "OTK-RF28LNK-" & scenario
    crid = "CRID-RF28LNK-" & scenario
    zbrID = "ZBR-RF28LNK-" & scenario
    brojA = "RF28-A-" & scenario
    brojB = "RF28-B-" & scenario

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA

    ' Otkup je VEC vezan na zbirnu A.
    AppendRF28OtkupFixture otkID, testDate, TEST_VOZ_ID, crid, brojA
    AppendRF28ZbirnaFixture zbrID, testDate, TEST_VOZ_ID, brojB

    Dim raised As Boolean
    On Error Resume Next
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrID, brojB, crid
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    AssertTrue raised, _
        "RF-28 AUD-043b: link na otkup sa drugim BrojZbirne podize konflikt"
    AssertEquals brojA, Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_BROJ_ZBIRNE))), _
        "RF-28 AUD-043b: postojeci BrojZbirne NIJE prepisan"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "RF-28 AUD-043b link konflikt", bfpErrDesc
End Sub

' AUD-043(b): membership se razresava preko ZbirnaID (PK), NE preko BrojZbirne.
' Dve zbirne sa ISTIM poslovnim brojem (multi-device kolizija): LookupValue bi
' vratio PRVU (drugi vozac) i lazno prijavio konflikt vozaca -- PK putanja mora
' da procita vozaca SVOG reda i da link prode.
Private Sub Test_RF28_MembershipKoristiSvojuZbirnu()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("RF28PK")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojIsti As String
    brojIsti = "RF28-DUP-" & scenario

    Dim vozacDrugi As String
    vozacDrugi = "VOZ-RF28-OTHER"

    Dim otkID As String, crid As String
    Dim zbrStara As String, zbrNova As String

    otkID = "OTK-RF28PK-" & scenario
    crid = "CRID-RF28PK-" & scenario
    zbrStara = "ZBR-RF28PK-OLD-" & scenario
    zbrNova = "ZBR-RF28PK-NEW-" & scenario

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA

    ' Redosled je bitan: STARA (tudji vozac) je PRVI match za BrojZbirne.
    AppendRF28ZbirnaFixture zbrStara, testDate, vozacDrugi, brojIsti
    AppendRF28ZbirnaFixture zbrNova, testDate, TEST_VOZ_ID, brojIsti

    AppendRF28OtkupFixture otkID, testDate, TEST_VOZ_ID, crid

    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrNova, brojIsti, crid

    AssertEquals brojIsti, Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_BROJ_ZBIRNE))), _
        "RF-28 AUD-043b: membership preko PK povezuje otkup sa SVOJOM zbirnom"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "RF-28 AUD-043b membership preko PK", bfpErrDesc
End Sub

' AUD-043(b): dan je stvarni guard -- susedni dan prolazi (utovar posle ponoci),
' veca razlika pada.
Private Sub Test_RF28_MembershipDanskiProzor()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("RF28DAY")

    Dim zbrDate As Date
    zbrDate = NextTestDate()

    Dim brojZ As String
    brojZ = "RF28-DAY-" & scenario

    Dim zbrID As String
    zbrID = "ZBR-RF28DAY-" & scenario

    Dim otkBlizu As String, cridBlizu As String
    Dim otkDaleko As String, cridDaleko As String

    otkBlizu = "OTK-RF28DAY-N-" & scenario
    cridBlizu = "CRID-RF28DAY-N-" & scenario
    otkDaleko = "OTK-RF28DAY-F-" & scenario
    cridDaleko = "CRID-RF28DAY-F-" & scenario

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA

    AppendRF28ZbirnaFixture zbrID, zbrDate, TEST_VOZ_ID, brojZ

    ' Susedni dan -> dozvoljeno (samo LogWarn).
    AppendRF28OtkupFixture otkBlizu, zbrDate - 1, TEST_VOZ_ID, cridBlizu
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrID, brojZ, cridBlizu

    AssertEquals brojZ, Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkBlizu, COL_OTK_BROJ_ZBIRNE))), _
        "RF-28 AUD-043b: otkup od prethodnog dana prolazi (post-midnight)"

    ' 10 dana razlike -> nije membership.
    AppendRF28OtkupFixture otkDaleko, zbrDate - 10, TEST_VOZ_ID, cridDaleko

    Dim raised As Boolean
    On Error Resume Next
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrID, brojZ, cridDaleko
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    AssertTrue raised, _
        "RF-28 AUD-043b: otkup 10 dana od zbirne je odbijen (nije samo upozorenje)"
    AssertEquals "", Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkDaleko, COL_OTK_BROJ_ZBIRNE))), _
        "RF-28 AUD-043b: odbijen otkup nije dobio BrojZbirne"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "RF-28 AUD-043b danski prozor", bfpErrDesc
End Sub

' AUD-042(b): nevalidan datum je SyncError, ne tihi danasnji datum.
Private Sub Test_RF28_NevalidanDatumJeSyncError()
    On Error GoTo EH

    AssertEquals "", TestHook_ValidatePWAOtkupDatum(TEST_KOOP_ID, DateSerial(2090, 5, 5)), _
        "RF-28 AUD-042b: validan OTK datum prolazi"

    ' STVARNI format iz pipeline-a: PWA salje ISO string (getTodayIsoDate ->
    ' "yyyy-mm-dd"), a ne native Date serijal. Testira se bas taj oblik, jer se
    ' validacija i import oslanjaju na CDate nad tim stringom.
    AssertEquals "", TestHook_ValidatePWAOtkupDatum(TEST_KOOP_ID, "2090-05-05"), _
        "RF-28 AUD-042b: ISO string datum (PWA format) prolazi"
    AssertEquals "", TestHook_ValidatePWAOtkupDatum(TEST_KOOP_ID, "2026-01-31"), _
        "RF-28 AUD-042b: backdate ISO string prolazi (donja granica ne odbija realne datume)"
    AssertEquals "", TestHook_ValidatePWAZbirnaDatum(TEST_VOZ_ID, TEST_KUP_ID, "2090-05-05"), _
        "RF-28 AUD-042b: ISO string datum prolazi i na VOZ putanji"

    AssertTrue Len(TestHook_ValidatePWAOtkupDatum(TEST_KOOP_ID, "")) > 0, _
        "RF-28 AUD-042b: prazan OTK datum je greska"
    AssertTrue Len(TestHook_ValidatePWAOtkupDatum(TEST_KOOP_ID, "nije datum")) > 0, _
        "RF-28 AUD-042b: neparsiran OTK datum je greska"
    AssertTrue Len(TestHook_ValidatePWAOtkupDatum(TEST_KOOP_ID, "12:30")) > 0, _
        "RF-28 AUD-042b: samo-vreme nije OTK datum"
    AssertTrue Len(TestHook_ValidatePWAOtkupDatum(TEST_KOOP_ID, "1899-12-30")) > 0, _
        "RF-28 AUD-042b: 1899 baseline nije poslovni datum"

    AssertEquals "", TestHook_ValidatePWAZbirnaDatum(TEST_VOZ_ID, TEST_KUP_ID, DateSerial(2090, 5, 5)), _
        "RF-28 AUD-042b: validan VOZ datum prolazi"
    AssertTrue Len(TestHook_ValidatePWAZbirnaDatum(TEST_VOZ_ID, TEST_KUP_ID, "")) > 0, _
        "RF-28 AUD-042b: prazan VOZ datum je greska"
    AssertTrue Len(TestHook_ValidatePWAZbirnaDatum(TEST_VOZ_ID, TEST_KUP_ID, "nije datum")) > 0, _
        "RF-28 AUD-042b: neparsiran VOZ datum je greska"

    Exit Sub

EH:
    LogFail "RF-28 AUD-042b nevalidan datum", Err.description
End Sub

' AUD-042(a): ishodi VozacID update-a se razlikuju. CONFLICT/NOTFOUND ne smeju da
' izgledaju kao obican Duplicate (pozivalac ih zato salje u SyncError).
Private Sub Test_RF28_VozacIDUpdateIshodi()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("RF28VOZ")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim otkPrazan As String, cridPrazan As String
    Dim otkZauzet As String, cridZauzet As String
    Dim otkPad As String, cridPad As String

    otkPrazan = "OTK-RF28VOZ-E-" & scenario
    cridPrazan = "CRID-RF28VOZ-E-" & scenario
    otkZauzet = "OTK-RF28VOZ-F-" & scenario
    cridZauzet = "CRID-RF28VOZ-F-" & scenario
    otkPad = "OTK-RF28VOZ-X-" & scenario
    cridPad = "CRID-RF28VOZ-X-" & scenario

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    AppendRF28OtkupFixture otkPrazan, testDate, "", cridPrazan
    AppendRF28OtkupFixture otkZauzet, testDate, TEST_VOZ_ID, cridZauzet
    AppendRF28OtkupFixture otkPad, testDate, "", cridPad

    Dim detail As String

    AssertEquals "UPDATED", TestHook_TryUpdateVozacID(cridPrazan, TEST_VOZ_ID, detail), _
        "RF-28 AUD-042a: prazan VozacID se popunjava (UPDATED)"
    AssertEquals TEST_VOZ_ID, Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkPrazan, COL_OTK_VOZAC))), _
        "RF-28 AUD-042a: VozacID je stvarno upisan"

    AssertEquals "NOCHANGE", TestHook_TryUpdateVozacID(cridPrazan, TEST_VOZ_ID, detail), _
        "RF-28 AUD-042a: isti VozacID je NOCHANGE (bezopasno -> Duplicate)"

    AssertEquals "CONFLICT", TestHook_TryUpdateVozacID(cridZauzet, "VOZ-RF28-OTHER", detail), _
        "RF-28 AUD-042a: drugi VozacID je CONFLICT (ne tihi Duplicate)"
    AssertEquals TEST_VOZ_ID, Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkZauzet, COL_OTK_VOZAC))), _
        "RF-28 AUD-042a: konflikt NE prepisuje postojeci VozacID"

    AssertEquals "NOTFOUND", TestHook_TryUpdateVozacID("CRID-RF28-NEMA-" & scenario, TEST_VOZ_ID, detail), _
        "RF-28 AUD-042a: nepostojeci ClientRecordID je NOTFOUND (greska, ne preskok)"

    ' Armiran pad upisa (UpdateCell se ne moze naterati da padne "prirodno").
    ' Ovo je putanja zbog koje je AUD-042a i postojao: stari kod je vracao True.
    TestHook_ArmFailSeam "VOZAC_WRITE"

    AssertEquals "FAILED", TestHook_TryUpdateVozacID(cridPad, TEST_VOZ_ID, detail), _
        "RF-28 AUD-042a: neuspeo UpdateCell je FAILED (ne tihi uspeh)"
    AssertTrue Len(detail) > 0, _
        "RF-28 AUD-042a: FAILED nosi detalj za SyncError/log"
    AssertEquals "", Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkPad, COL_OTK_VOZAC))), _
        "RF-28 AUD-042a: posle neuspelog upisa VozacID je i dalje prazan"

    ' Seam je jednokratan -- sledeci poziv mora ponovo da radi normalno.
    AssertEquals "UPDATED", TestHook_TryUpdateVozacID(cridPad, TEST_VOZ_ID, detail), _
        "RF-28 AUD-042a: fail seam je jednokratan (sledeci upis prolazi)"

    TestHook_ArmFailSeam ""

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    TestHook_ArmFailSeam ""
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "RF-28 AUD-042a VozacID ishodi", bfpErrDesc
End Sub

' ------------------------------------------------------------
' RF-28 fixture helperi (direktan append -- kontrolisemo tacno polja koja
' grupisanje/membership citaju, bez zavisnosti od validacija save putanje)
' ------------------------------------------------------------
Private Sub AppendRF28OtkupFixture(ByVal otkupID As String, _
                                   ByVal datum As Date, _
                                   ByVal vozacID As String, _
                                   ByVal clientRecordID As String, _
                                   Optional ByVal brojZbirne As String = "", _
                                   Optional ByVal vrsta As String = TEST_VRSTA, _
                                   Optional ByVal sorta As String = TEST_SORTA, _
                                   Optional ByVal tipAmb As String = TEST_TIP_AMB)
    Dim rowData As Variant
    rowData = BlankRow(TBL_OTKUP)

    SetRequiredField rowData, TBL_OTKUP, COL_OTK_ID, otkupID
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_DATUM, datum
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_KOOPERANT, TEST_KOOP_ID
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_STANICA, TEST_ST_ID
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_VRSTA, vrsta
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_SORTA, sorta
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_KULTURA, TEST_KULTURA_ID
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_TIP_AMB, tipAmb
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_VOZAC, vozacID
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_BR_DOK, "RF28-" & otkupID
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_BROJ_ZBIRNE, brojZbirne
    SetOptionalField rowData, TBL_OTKUP, "ClientRecordID", clientRecordID
    SetOptionalField rowData, TBL_OTKUP, "SyncSource", "RF28TEST"

    RequireAppend TBL_OTKUP, rowData, "AppendRF28OtkupFixture"
End Sub

Private Sub AppendRF28OtpremnicaFixture(ByVal otpremnicaID As String, _
                                        ByVal datum As Date, _
                                        ByVal vozacID As String, _
                                        ByVal brojOtpremnice As String)
    Dim rowData As Variant
    rowData = BlankRow(TBL_OTPREMNICA)

    SetRequiredField rowData, TBL_OTPREMNICA, COL_OTP_ID, otpremnicaID
    SetRequiredField rowData, TBL_OTPREMNICA, COL_OTP_DATUM, datum
    SetRequiredField rowData, TBL_OTPREMNICA, COL_OTP_STANICA, TEST_ST_ID
    SetRequiredField rowData, TBL_OTPREMNICA, COL_OTP_VOZAC, vozacID
    SetRequiredField rowData, TBL_OTPREMNICA, COL_OTP_BROJ, brojOtpremnice
    SetOptionalField rowData, TBL_OTPREMNICA, COL_OTP_VRSTA, TEST_VRSTA
    SetOptionalField rowData, TBL_OTPREMNICA, COL_OTP_SORTA, TEST_SORTA
    SetOptionalField rowData, TBL_OTPREMNICA, COL_OTP_KOLICINA, 100#
    SetOptionalField rowData, TBL_OTPREMNICA, COL_OTP_CENA, 10#
    SetOptionalField rowData, TBL_OTPREMNICA, COL_OTP_TIP_AMB, TEST_TIP_AMB
    SetOptionalField rowData, TBL_OTPREMNICA, COL_OTP_KOL_AMB, 0
    SetOptionalField rowData, TBL_OTPREMNICA, COL_OTP_KLASA, "I"

    ' BrojZbirne i ZbirnaGeneracijaID ostaju PRAZNI -- dete pre roditelja, sto je
    ' za otpremnicu legitimno (auto-lanac je snima pre zbirne).
    RequireAppend TBL_OTPREMNICA, rowData, "AppendRF28OtpremnicaFixture"
End Sub

Private Sub AppendRF28ZbirnaFixture(ByVal zbirnaID As String, _
                                    ByVal datum As Date, _
                                    ByVal vozacID As String, _
                                    ByVal brojZbirne As String)
    Dim rowData As Variant
    rowData = BlankRow(TBL_ZBIRNA)

    SetRequiredField rowData, TBL_ZBIRNA, COL_ZBR_ID, zbirnaID
    SetRequiredField rowData, TBL_ZBIRNA, COL_ZBR_DATUM, datum
    SetRequiredField rowData, TBL_ZBIRNA, COL_ZBR_VOZAC, vozacID
    SetRequiredField rowData, TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne
    SetRequiredField rowData, TBL_ZBIRNA, COL_ZBR_KUPAC, TEST_KUP_ID
    SetOptionalField rowData, TBL_ZBIRNA, COL_ZBR_VRSTA, TEST_VRSTA
    SetOptionalField rowData, TBL_ZBIRNA, COL_ZBR_SORTA, TEST_SORTA
    SetOptionalField rowData, TBL_ZBIRNA, COL_ZBR_KOLICINA, 100#
    SetOptionalField rowData, TBL_ZBIRNA, COL_ZBR_TIP_AMB, TEST_TIP_AMB
    SetOptionalField rowData, TBL_ZBIRNA, COL_ZBR_KOL_AMB, 0
    SetOptionalField rowData, TBL_ZBIRNA, COL_ZBR_KLASA, "I"

    RequireAppend TBL_ZBIRNA, rowData, "AppendRF28ZbirnaFixture"
End Sub

' ============================================================
' RF-05 -- frmDokumenta unos + storno set (regresija)
'   R01 prosek gajbe ne racuna stornirane redove (SumByBroj)
'   R02 stornirana faktura ne ulazi u listu za placanje/avans (FillOpenFakture)
'   R03 izvor sa Klasom II blokira zbirnu bez "Dve klase" (ZbirnaIzvorImaKlasuII)
'   R04 prefill bira POSLEDNJU GENERACIJU (GeneracijaID, ne datum/ID kontinuitet)
'   R05 malina auto-zbirna signalizira pad (Err / created=0), scoped na svoj broj
'   R06 katalog poruka sadrzi kljuceve koje frmDokumenta koristi (EnsurePoruke)
'   R07 SaveZbirna upisuje po IMENU kolone (BuildZbirnaRowData)
'   R08 OM ulaz: smer ambalaze je obavezan (core guard u SaveOMUlaz_TX)
'   R09 storno po broju sa dva vlasnika je odbijen (ne stornira tudji dokument)
'   R10 isti guard vazi i na ISPRAVKA/DUPLI/SIMPLE correction putanjama
'   R11 guard vazi i u malina/autohladnjaca kaskadama (ulaz je BrojZbirne)
'   R12 kaskade mutiraju samo redove razresenog lanca (scope), fail-closed bez parenta
' ============================================================

' MIG-004. Manjak koji F4 crta nije novi racun nego CalculateManjakPreview --
' funkcija koja je od brisanja frmDokumenta ostala bez ijednog pozivaoca. Ovo
' meri BAS NJU, nad pravim redovima: da sabira po broju zbirne, da dodaje
' NEUPISANE kilograme iz forme, i da stornirane redove ne broji ni sa jedne
' strane. Bez toga bi ekran mogao da pokaze zelenu nulu nad podacima koji se
' ne slazu.
Private Sub Test_ManjakPreviewJeZbirnaMinusPrijem()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("MANJAK")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String, brojPrij As String
    brojZbirne = TEST_PREFIX & "-ZBR-MJ-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-MJ-" & scenario

    ' Zbirna 300 kg u dva reda (100 + 200) -- manjak se meri po BROJU, ne po
    ' jednom redu; sa jednim redom bi i pogresno "uzmi prvi" prolazilo.
    Dim zbrI As String, zbrII As String
    zbrI = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    zbrII = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          200#, TEST_TIP_AMB, 20, KLASA_II)
    AssertTrue Len(zbrI) > 0 And Len(zbrII) > 0, "Manjak: fixture zbirna I+II kreirana"

    Dim prj As String
    prj = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                            TEST_VRSTA, TEST_SORTA, 200#, 100#, TEST_TIP_AMB, 20, 0, "I")
    AssertTrue Len(prj) > 0, "Manjak: fixture prijemnica kreirana"

    ' 300 upisano na zbirnoj, 200 stiglo prijemnicom -> manjak 100 kg = 33,33%.
    Dim m As Variant
    m = CalculateManjakPreview(brojZbirne, 0#, 0#)
    AssertTrue IsArray(m), "Manjak: preview vraca niz"
    AssertDoubleNear 300#, CDbl(m(0)), 0.01, "Manjak: zbirna je ZBIR svih svojih redova"
    AssertDoubleNear 200#, CDbl(m(1)), 0.01, "Manjak: prijemnica je zbir upisanih redova"
    AssertDoubleNear 100#, CDbl(m(2)), 0.01, "Manjak: manjak je zbirna minus prijemnica"
    AssertDoubleNear 33.3333, CDbl(m(3)), 0.01, "Manjak: procenat je udeo u ZBIRNOJ"

    ' NEUPISANI kilogrami iz forme ulaze u prijemnicu -- to je cela svrha
    ' "Preview" varijante: operater vidi manjak PRE snimanja, ne posle.
    m = CalculateManjakPreview(brojZbirne, 50#, 25#)
    AssertDoubleNear 275#, CDbl(m(1)), 0.01, "Manjak: neupisane kg obe klase ulaze u prijemnicu"
    AssertDoubleNear 25#, CDbl(m(2)), 0.01, "Manjak: manjak pada za neupisane kilograme"

    ' Stornirana prijemnica NE ulazi -- inace bi ispravka dokumenta izgledala
    ' kao da je roba stigla dvaput.
    MarkTestRowStornirano TBL_PRIJEMNICA, "PrijemnicaID", prj
    m = CalculateManjakPreview(brojZbirne, 0#, 0#)
    AssertDoubleNear 0#, CDbl(m(1)), 0.01, "Manjak: stornirana prijemnica se NE broji"
    AssertDoubleNear 300#, CDbl(m(2)), 0.01, "Manjak: posle storna fali cela zbirna"

    ' Ista uzica na drugoj strani: storniran red zbirne smanjuje ocekivanje.
    MarkTestRowStornirano TBL_ZBIRNA, "ZbirnaID", zbrII
    m = CalculateManjakPreview(brojZbirne, 0#, 0#)
    AssertDoubleNear 100#, CDbl(m(0)), 0.01, "Manjak: storniran red zbirne se NE broji"

    ' Broj koji ne postoji nije "sve se slaze" nego NEMA ZBIRNE: zbirna ostaje
    ' nula, i bas po toj nuli ekran zna da liniju ne sme da nacrta.
    m = CalculateManjakPreview(TEST_PREFIX & "-ZBR-NEMA-" & scenario, 40#, 0#)
    AssertDoubleNear 0#, CDbl(m(0)), 0.01, "Manjak: nepostojeca zbirna daje nula kg"
    AssertDoubleNear 40#, CDbl(m(1)), 0.01, "Manjak: neupisane kg se i tada vide"

    Exit Sub

EH:
    LogFatal "Test_ManjakPreviewJeZbirnaMinusPrijem", Err.Number, Err.description
End Sub

Private Sub Test_OpenFaktureExcludeStornirano()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("FAKSTO")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String, brojPrij As String
    brojZbirne = TEST_PREFIX & "-ZBR-FS-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-FS-" & scenario

    Dim zbrFix As String
    zbrFix = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                           "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                           100#, TEST_TIP_AMB, 0, KLASA_I)
    AssertTrue Len(zbrFix) > 0, "Storno faktura: fixture zbirna kreirana"

    Dim prjI As String
    prjI = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    AssertTrue Len(prjI) > 0, "Storno faktura: fixture prijemnica kreirana"

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add Array(prjI, 100#, 100#, KLASA_I, brojPrij)

    Dim fakID As String
    fakID = CreateFaktura_TX(TEST_KUP_ID, stavke)
    AssertTrue Len(fakID) > 0, "Storno faktura: fixture faktura kreirana"

    ' Pre storna: faktura JESTE u produkcionom read-modelu koji forma zove
    ' (modNovac.GetOpenFakture -- FillOpenFakture vise nema sopstveni filter).
    AssertTrue OpenFaktureSadrzi(TEST_KUP_ID, fakID), _
               "Otvorena faktura je u GetOpenFakture (read-model koji forma zove)"
    AssertTrue OpenFaktureImaDatum(TEST_KUP_ID, fakID), _
               "GetOpenFakture vraca i Datum (6. kolona za prikaz u formi)"

    MarkTestRowStornirano TBL_FAKTURE, COL_FAK_ID, fakID

    ' Stornirana faktura NIJE "Placeno" -- stari filter forme (Status <> Placeno)
    ' bi je pustio nazad u listu.
    AssertTrue CStr(nz(GetValueByKey(TBL_FAKTURE, COL_FAK_ID, fakID, COL_FAK_STATUS), "")) <> STATUS_PLACENO, _
               "Storno faktura: status i dalje nije 'Placeno' (stari filter bi je pustio)"

    AssertFalse OpenFaktureSadrzi(TEST_KUP_ID, fakID), _
                "Stornirana faktura ne ulazi u listu za placanje/avans"

    Exit Sub

EH:
    LogFatal "Test_OpenFaktureExcludeStornirano", Err.Number, Err.description
End Sub

Private Sub Test_PrefillBiraPoslednjuGeneraciju()
    On Error GoTo EH

    ' Sinteticka 2D tabela (1-based): 1=Broj 2=Klasa 3=ID 4=GeneracijaID.
    ' REGRESIJA: uzastopni ID-evi preko granice generacije (30=I i 31=II stare,
    ' 32=I nove) -- heuristika ID kontinuiteta bi spojila novu Kl.I sa starom Kl.II.
    Dim d As Variant
    ReDim d(1 To 3, 1 To 4)
    d(1, 1) = "DOK-1": d(1, 2) = "I":  d(1, 3) = "OTP-00030": d(1, 4) = "GEN-00001"
    d(2, 1) = "DOK-1": d(2, 2) = "II": d(2, 3) = "OTP-00031": d(2, 4) = "GEN-00001"
    d(3, 1) = "DOK-1": d(3, 2) = "I":  d(3, 3) = "OTP-00032": d(3, 4) = "GEN-00002"

    Dim rI As Long, rII As Long

    ' Anchor = PK stornirane (novi Kl.I red).
    PickPrefillRows d, 1, 2, 3, 4, "DOK-1", "OTP-00032", rI, rII
    AssertEquals "3", CStr(rI), "Prefill: Kl.I iz generacije anchor reda"
    AssertEquals "0", CStr(rII), _
                 "Prefill: stara Kl.II (ID 31) se NE spaja sa novom Kl.I (ID 32)"

    ' Anchor na STAROJ generaciji -> prefiluje se ona, ne najnovija.
    PickPrefillRows d, 1, 2, 3, 4, "DOK-1", "OTP-00030", rI, rII
    AssertEquals "1", CStr(rI), "Prefill: anchor odredjuje generaciju (Kl.I stare)"
    AssertEquals "2", CStr(rII), "Prefill: Kl.II iste (stare) generacije"

    ' Bez anchor PK-a -> poslednje upisan red datog broja + njegova generacija.
    PickPrefillRows d, 1, 2, 3, 4, "DOK-1", "", rI, rII
    AssertEquals "3", CStr(rI), "Prefill fallback: poslednje upisan red broja"
    AssertEquals "0", CStr(rII), "Prefill fallback: ostaje u generaciji tog reda"

    ' KLJUCNO: dva vlasnika dele isti BROJ (razlicite generacije) -- prefill po PK
    ' ostaje kod svog dokumenta i ne prelazi na tudji.
    Dim x As Variant
    ReDim x(1 To 4, 1 To 4)
    x(1, 1) = "1/050826": x(1, 2) = "I":  x(1, 3) = "PRJ-00010": x(1, 4) = "GEN-00100"
    x(2, 1) = "1/050826": x(2, 2) = "II": x(2, 3) = "PRJ-00011": x(2, 4) = "GEN-00100"
    x(3, 1) = "1/050826": x(3, 2) = "I":  x(3, 3) = "PRJ-00012": x(3, 4) = "GEN-00101"
    x(4, 1) = "1/050826": x(4, 2) = "II": x(4, 3) = "PRJ-00013": x(4, 4) = "GEN-00101"

    PickPrefillRows x, 1, 2, 3, 4, "1/050826", "PRJ-00010", rI, rII
    AssertEquals "1", CStr(rI), "Prefill: Kl.I ostaje kod svog vlasnika (isti broj, drugi kupac)"
    AssertEquals "2", CStr(rII), "Prefill: Kl.II ostaje kod svog vlasnika"

    ' Bez generacije (red stariji od kolone) -> samo anchor.
    Dim f As Variant
    ReDim f(1 To 2, 1 To 4)
    f(1, 1) = "DOK-4": f(1, 2) = "I":  f(1, 3) = "OTP-00060": f(1, 4) = ""
    f(2, 1) = "DOK-4": f(2, 2) = "II": f(2, 3) = "OTP-00061": f(2, 4) = ""

    PickPrefillRows f, 1, 2, 3, 4, "DOK-4", "OTP-00060", rI, rII
    AssertEquals "1", CStr(rI), "Prefill bez generacije: samo anchor red"
    AssertEquals "0", CStr(rII), "Prefill bez generacije: druga klasa ostaje prazna"

    ' Nepoznat broj / nepoznat PK -> nista.
    PickPrefillRows d, 1, 2, 3, 4, "DOK-NEMA", "", rI, rII
    AssertEquals "00", CStr(rI) & CStr(rII), "Prefill: nepoznat broj ne vraca red"

    Exit Sub

EH:
    LogFatal "Test_PrefillBiraPoslednjuGeneraciju", Err.Number, Err.description
End Sub

' Storno po BROJU zahvata sve aktivne redove tog broja. Kad broj nije jedinstven
' (dva kupca), to bi tiho storniralo i tudji dokument -> mora biti ODBIJENO.
Private Sub Test_StornoPoBrojuOdbijaDvaVlasnika()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("STOVLAS")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String, brojPrij As String
    brojZbirne = TEST_PREFIX & "-ZBR-SV-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-SV-" & scenario     ' ISTI broj za oba kupca

    AssertTrue Len(SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                                 "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                                 100#, TEST_TIP_AMB, 0, KLASA_I)) > 0, _
               "Storno guard: fixture zbirna kreirana"

    Dim prjA As String, prjB As String
    prjA = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    prjB = SavePrijemnica_TX(testDate, TEST_KUP2_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 80#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    AssertTrue Len(prjA) > 0 And Len(prjB) > 0, _
               "Storno guard: obe prijemnice (isti broj, dva kupca) kreirane"

    ' Dvosmislen number-only storno mora pasti...
    AssertFalse StornoPrijemnicaByBroj_TX(brojPrij), _
                "Storno guard: storno po broju sa dva vlasnika je ODBIJEN"

    ' ...i ne sme ostaviti nijedan storniran red (rollback / nista nije dirano).
    AssertTrue Not RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, prjA), _
               "Storno guard: dokument kupca A ostaje aktivan"
    AssertTrue Not RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, prjB), _
               "Storno guard: dokument kupca B ostaje aktivan"

    ' Kontrola: jedinstven broj (jedan vlasnik, obe klase) i dalje prolazi.
    Dim brojPrijOK As String
    brojPrijOK = TEST_PREFIX & "-PRJ-SV1-" & scenario

    Dim okI As String, okII As String
    okI = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrijOK, brojZbirne, _
                            TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    okII = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrijOK, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 40#, 90#, TEST_TIP_AMB, 0, 0, KLASA_II)
    AssertTrue Len(okI) > 0 And Len(okII) > 0, "Storno guard: fixture jednog vlasnika kreiran"

    AssertTrue StornoPrijemnicaByBroj_TX(brojPrijOK), _
               "Storno guard: jedinstven broj (jedan vlasnik) i dalje prolazi"
    AssertTrue RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, okI), _
               "Storno guard: Kl.I stornirana"
    AssertTrue RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, okII), _
               "Storno guard: Kl.II stornirana (obe klase istog broja)"

    Exit Sub

EH:
    LogFatal "Test_StornoPoBrojuOdbijaDvaVlasnika", Err.Number, Err.description
End Sub

Private Function RowIsStornirano(ByVal tableName As String, ByVal idColumn As String, _
                                 ByVal idValue As String) As Boolean
    If Len(Trim$(idValue)) = 0 Then Exit Function

    RowIsStornirano = (UCase$(Trim$(CStr(nz(GetValueByKey(tableName, idColumn, idValue, _
                                                          COL_STORNIRANO), "")))) = "DA")
End Function

' ZBR-CHILD-01: paleta nasledjuje generaciju OD PRIJEMNICE, ne razresava po broju.
'
' Kanonski lanac je PaletaStavka -> Prijemnica -> Zbirna, i prijemnica svoj
' ZbirnaGeneracijaID vec nosi. Pitanje "koja je zbirna SADA pod ovim brojem" je
' zato i suvisno i pogresno -- pravilo je "nikad ne pogadjaj kad vec znas".
'
' Grana A sama NE razlikuje tacno od pogresnog: kad je prijemnica vezana za
' jedinu zbirnu pod tim brojem, nasledjivanje i pogadjanje vracaju ISTU vrednost.
' Prva verzija ovog testa je imala samo granu A i sabotaza je prosla neprimeceno
' (dokaz.py: NE OBARA NISTA).
'
' Grana B je ZATECEN red: broj stoji, generacija prazna. Druga verzija je tu
' pokusala redosled "dete pre roditelja" i pukla na ValidatePrijemnicaInput:
' PrijemnicaZbirnaBlokira() je po defaultu True (Case Else hvata i prazno), pa
' prijemnica bez postojece zbirne uopste ne prolazi. Taj redosled je stvaran za
' OTPREMNICU iz auto-lanca (modAutoHladnjaca), ne za prijemnicu.
'
' Zatecen red je pak stvaran za oba: tako izgleda svaki red pre migracije, i
' takav ostaje dok backfill ne prodje -- a pod dvosmislenim brojem ostaje prazan
' zauvek. Tu nasledjivanje daje prazno, a pogadjanje generaciju. Stavka koja bi
' pogodila tvrdila bi sledljivost koju njen sopstveni roditelj nema.
Private Sub Test_ZBR_PaletaNasledjujeGeneracijuPrijemnice()
    Dim tx As clsTransaction
    Dim testDate As Date, scenario As String
    Dim brojA As String
    Dim brPrijA As String, brPrijB As String
    Dim zbrA As String, prjA As String, prjB As String
    Dim genPrj As String

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRPAL")
    testDate = NextTestDate()
    brojA = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    brPrijA = TEST_PREFIX & "-PRJ-PALA-" & scenario
    brPrijB = TEST_PREFIX & "-PRJ-PALB-" & scenario

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    ' --- A) uobicajen redosled: roditelj pa dete ---
    zbrA = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojA, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    AssertTrue Len(zbrA) > 0, "ZBR-PAL preduslov: zbirna je snimljena"

    prjA = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brPrijA, brojA, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, 0, _
                             KLASA_I, 0)
    AssertTrue Len(prjA) > 0, "ZBR-PAL preduslov: prijemnica je snimljena"

    genPrj = NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prjA, COL_DETE_ZBIRNA_GEN))
    AssertTrue Len(genPrj) > 0, "ZBR-PAL preduslov: prijemnica nosi generaciju roditelja"
    AssertTrue BrojPaletnihStavki(prjA) > 0, _
        "ZBR-PAL preduslov: paletizacija je napravila stavku (grana A)"
    AssertEquals genPrj, PrvaGeneracijaPaletneStavke(prjA), _
        "ZBR-PAL: paletna stavka nosi ISTU generaciju kao njena prijemnica"

    ' --- B) ZATECEN red: broj stoji, generacija prazna ---
    prjB = SavePrijemnica(testDate, TEST_KUP_ID, TEST_VOZ_ID, brPrijB, brojA, _
                          TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 0, 0, _
                          KLASA_I, 0)
    AssertTrue Len(prjB) > 0, "ZBR-PAL preduslov: druga prijemnica je snimljena"

    IsprazniGeneracijuDeteta TBL_PRIJEMNICA, COL_PRJ_ID, prjB
    AssertEquals "", _
        NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prjB, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-PAL preduslov: prijemnica je u zatecenom obliku (generacija prazna)"
    AssertEquals brojA, _
        NzToText(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prjB, COL_PRJ_BROJ_ZBIRNE)), _
        "ZBR-PAL preduslov: prijemnica je zadrzala broj"
    AssertEquals genPrj, ZbirnaGeneracijaZaBroj(brojA), _
        "ZBR-PAL preduslov: broj razresava na generaciju (ima sta da se pogodi)"

    PaletizePrijemnica prijemnicaID:=prjB, brojPrij:=brPrijB, brojZbirne:=brojA, _
                       vrstaVoca:=TEST_VRSTA, sortaVoca:=TEST_SORTA, klasa:=KLASA_I, _
                       netoKg:=100#, brGajbica:=10, tipAmb:=TEST_TIP_AMB

    AssertTrue BrojPaletnihStavki(prjB) > 0, _
        "ZBR-PAL preduslov: paletizacija je napravila stavku (grana B)"
    AssertEquals "", PrvaGeneracijaPaletneStavke(prjB), _
        "ZBR-PAL: prazna generacija roditelja ostaje prazna, ne pogadja se po broju"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 paleta nasledjuje od prijemnice", bfpErrDesc
End Sub

' Generacija prve paletne stavke date prijemnice, ili prazno. Prazno je i
' legitiman rezultat i "nema stavke", pa se postojanje meri BrojPaletnihStavki.
Private Function PrvaGeneracijaPaletneStavke(ByVal prijemnicaID As String) As String
    Dim dat As Variant: dat = GetTableData(TBL_PALETA_STAVKA)
    If Not IsArray(dat) Then Exit Function
    Dim cP As Long, cG As Long, r As Long
    cP = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    cG = GetColumnIndex(TBL_PALETA_STAVKA, COL_DETE_ZBIRNA_GEN)
    If cP = 0 Or cG = 0 Then Exit Function
    For r = 1 To UBound(dat, 1)
        If Trim$(NzToText(dat(r, cP))) = Trim$(prijemnicaID) Then
            PrvaGeneracijaPaletneStavke = Trim$(NzToText(dat(r, cG)))
            Exit Function
        End If
    Next r
End Function

Private Function BrojPaletnihStavki(ByVal prijemnicaID As String) As Long
    Dim rows As Collection
    Set rows = FindRows(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID, prijemnicaID)
    If rows Is Nothing Then Exit Function
    BrojPaletnihStavki = rows.count
End Function

' Vraca red u oblik kakav ima pre migracije: broj zbirne stoji, generacija ne.
Private Sub IsprazniGeneracijuDeteta(ByVal tableName As String, _
                                     ByVal idColumn As String, _
                                     ByVal idValue As String)
    Const SRC As String = "IsprazniGeneracijuDeteta"

    Dim rows As Collection
    Set rows = FindRows(tableName, idColumn, idValue)
    If rows Is Nothing Or rows.count = 0 Then
        Err.Raise vbObjectError + 9311, SRC, _
                  "Red nije nadjen. Tabela=" & tableName & " ID=" & idValue
    End If

    RequireUpdateCell tableName, CLng(rows(1)), COL_DETE_ZBIRNA_GEN, "", SRC
End Sub

' ZBR-CHILD-01 / P1: ingest NE SME da premesti dete na drugi dokument.
'
' Dok je PoveziDeteNaZbirnu pisao samo broj, drugi link pod istim brojem je bio
' idempotentan -- ista vrednost preko sebe. Otkad pise i generaciju, isti put
' menja ROdITELJA deteta, a stara kapija (samo broj) to ne vidi. Regresiju je
' uveo upis, ne kapija.
'
' Scenario je KR-001, koji ugovor izricito dozvoljava: dva uredjaja bez veze
' posalju zbirnu pod istim brojem, istim vozacem i istim kupcem. Membership
' kapije (vozac, poslovni dan) tu prolaze, pa dete legitimno stigne u oba skupa.
'
' Mere se OBA pozivna mesta iste kapije:
'   korak 2 -- otkup je vec dete GEN-A
'   korak 3 -- otkup je cist, ali otpremnica na koju pokazuje je dete GEN-A
Private Sub Test_ZBR_MasterSyncNePrepisujeGeneracijuDeteta()
    Dim tx As clsTransaction
    Dim scenario As String, testDate As Date
    Dim broj As String, brojOtp As String
    Dim zbrA As String, zbrB As String, genA As String, genB As String
    Dim otkID As String, otkID2 As String
    Dim crid As String, crid2 As String, otpID As String
    Dim raised As Boolean

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRFK")
    testDate = NextTestDate()
    broj = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    otkID = "OTK-ZBRFK-A-" & scenario
    otkID2 = "OTK-ZBRFK-B-" & scenario
    crid = "CRID-ZBRFK-A-" & scenario
    crid2 = "CRID-ZBRFK-B-" & scenario
    otpID = "OTP-ZBRFK-" & scenario
    brojOtp = TEST_PREFIX & "-OTP-ZBRFK-" & scenario

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTPREMNICA

    zbrA = TestHook_ImportZbirnaRowPWA("CRID-ZBRFK-ZA-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, _
                                       100, broj)
    zbrB = TestHook_ImportZbirnaRowPWA("CRID-ZBRFK-ZB-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, _
                                       120, broj)
    genA = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrA)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrB)
    AssertTrue (Len(genA) > 0 And Len(genB) > 0 And genA <> genB), _
        "ZBR-FK preduslov: dva dokumenta pod istim brojem nose RAZLICITE generacije"

    AppendRF28OtpremnicaFixture otpID, testDate, TEST_VOZ_ID, brojOtp
    AppendRF28OtkupFixture otkID, testDate, TEST_VOZ_ID, crid, ""
    VeziOtkupZaOtpremnicuFixture otkID, otpID

    ' --- 1) prvi link DOVRSAVA praznu vezu ---
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrA, broj, crid
    AssertEquals genA, DeteGeneracija(TBL_OTKUP, COL_OTK_ID, otkID), _
        "ZBR-FK preduslov: prvi link je upisao generaciju A na otkup"
    AssertEquals genA, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpID), _
        "ZBR-FK preduslov: prvi link je upisao generaciju A na otpremnicu"

    ' --- 2) drugi dokument, ISTI broj -> kapija na otkupu ---
    raised = False
    On Error Resume Next
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrB, broj, crid
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    AssertTrue raised, _
        "ZBR-FK: drugi dokument pod istim brojem ne prolazi tiho"
    AssertEquals genA, DeteGeneracija(TBL_OTKUP, COL_OTK_ID, otkID), _
        "ZBR-FK: otkup ostaje na svojoj originalnoj generaciji"
    AssertEquals broj, _
        NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_BROJ_ZBIRNE)), _
        "ZBR-FK: otkup zadrzava broj -- blokira se generacija, ne broj"

    ' --- 3) ista kapija na otpremnickom pozivnom mestu ---
    ' Otkup2 je cist, pa njegova kapija pusta; otpremnica na koju pokazuje je vec
    ' dete GEN-A. Bez ovog koraka drugo pozivno mesto ostaje nemereno.
    AppendRF28OtkupFixture otkID2, testDate, TEST_VOZ_ID, crid2, ""
    VeziOtkupZaOtpremnicuFixture otkID2, otpID

    raised = False
    On Error Resume Next
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrB, broj, crid2
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    AssertTrue raised, _
        "ZBR-FK: kapija radi i na otpremnickom pozivnom mestu"
    AssertEquals genA, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpID), _
        "ZBR-FK: otpremnica ostaje na svojoj originalnoj generaciji"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 MasterSync ne prepisuje generaciju deteta", bfpErrDesc
End Sub

Private Function DeteGeneracija(ByVal tableName As String, _
                                ByVal idColumn As String, _
                                ByVal idValue As String) As String
    DeteGeneracija = NzToText(LookupValue(tableName, idColumn, idValue, COL_DETE_ZBIRNA_GEN))
End Function

Private Sub VeziOtkupZaOtpremnicuFixture(ByVal otkupID As String, _
                                         ByVal otpremnicaID As String)
    Const SRC As String = "VeziOtkupZaOtpremnicuFixture"

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If rows Is Nothing Or rows.count = 0 Then
        Err.Raise vbObjectError + 9321, SRC, "Otkup nije nadjen. ID=" & otkupID
    End If

    RequireUpdateCell TBL_OTKUP, CLng(rows(1)), COL_OTK_OTPREMNICA_ID, otpremnicaID, SRC
End Sub

Private Function StornoOznaka(ByVal tableName As String, ByVal idColumn As String, _
                                      ByVal idValue As String) As String
    StornoOznaka = UCase$(NzToText(LookupValue(tableName, idColumn, idValue, COL_STORNIRANO)))
End Function

Private Function RedJeStorniran(ByVal tableName As String, ByVal idColumn As String, _
                                ByVal idValue As String) As Boolean
    RedJeStorniran = (StornoOznaka(tableName, idColumn, idValue) = "DA")
End Function

' ZAUZETOST BROJA OTPREMNICE -- ekran i pisac, po nizu (stanica, dan), sa
' storniranima (A2 tacke 2 i 4, odluke 14.09.2026).
'
' Nivo merenja: poslovni broj u nizu. Dvoklasna otpremnica je JEDAN dokument na
' dva reda -- provera je jednom po dokumentu, pa oba reda moraju da se upisu.
'
' SABOTAZE: preskoci stornirane u BrojZauzetUNizu -> pukne "storno ne oslobadja
' broj"; premesti proveru iz Multi_TX u SaveOtpremnica -> pukne "dvoklasna
' otpremnica upisuje obe klase"; vrati CheckDuplicate u OtpremnicaValidiraj ->
' pukne "druga stanica istog dana prolazi ekran".
Private Sub Test_OTP_BrojZauzetPoStaniciIDanu()
    On Error GoTo EH

    Dim scenario As String: scenario = NewScenarioCode("OTPBZ")
    Dim d As Date: d = NextTestDate()
    Dim broj As String: broj = TEST_PREFIX & "-OTP-BZ-" & scenario
    Dim zauzeto As String: zauzeto = Poruka("DOKUNOS_ERR_BROJ_ZAUZET")

    ' Produkcioni pisac (S3b-1): dvoklasna otpremnica je JEDNO zaglavlje sa dve
    ' stavke. Do S3b-1 je ovde stajao stari pisac, koji je pravio dva reda.
    Dim preRedova As Long: preRedova = CountRows(TBL_OTPREMNICA)
    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpBrojHeader(broj, d, TEST_ST_ID), _
                                     OtpOcek(100#, 10#, 50#, 0#), razlog)
    AssertTrue Len(otpID) > 0, _
               "OTP broj: dvoklasna otpremnica upisana (bilo: " & razlog & ")"
    AssertEquals CStr(preRedova + 1), CStr(CountRows(TBL_OTPREMNICA)), _
                 "OTP broj: dvoklasna otpremnica je JEDNO zaglavlje"

    ' Unos koji prolazi SVE provere pre provere broja -- inace bi tvrdnje o
    ' broju pale ili prosle iz pogresnog razloga.
    Dim p As Object, fokus As String, r As String
    Set p = modDokUnos.NoviOtpremnicaUnos()
    p("stanicaID") = TEST_ST_ID
    p("vozacID") = TEST_VOZ_ID
    p("vrsta") = TEST_VRSTA
    p("sorta") = TEST_SORTA
    p("tipAmb") = TEST_TIP_AMB
    p("kolicinaI") = 100#
    p("cenaI") = 10#
    p("kolAmb") = 10
    p("datum") = d
    p("brDok") = broj

    r = modDokUnos.OtpremnicaValidiraj(p, fokus)
    AssertTrue InStr(1, r, zauzeto, vbBinaryCompare) = 1, _
               "OTP broj: ekran odbija isti broj, stanicu i dan (bilo: " & r & ")"
    AssertEquals "brDok", fokus, "OTP broj: fokus ide na broj"

    p("brDok") = "  " & LCase$(broj) & " "
    r = modDokUnos.OtpremnicaValidiraj(p, fokus)
    AssertTrue InStr(1, r, zauzeto, vbBinaryCompare) = 1, _
               "OTP broj: razmaci i mala slova ne otvaraju rupu (bilo: " & r & ")"

    p("brDok") = broj
    p("stanicaID") = TEST_HLAD_ST_ID
    r = modDokUnos.OtpremnicaValidiraj(p, fokus)
    AssertEquals "", r, "OTP broj: druga stanica istog dana prolazi ekran (A2)"

    p("stanicaID") = TEST_ST_ID
    p("datum") = DateAdd("d", 1, d)
    r = modDokUnos.OtpremnicaValidiraj(p, fokus)
    AssertEquals "", r, "OTP broj: drugi dan iste stanice prolazi ekran"

    ' Pisac: isti niz odbijen.
    preRedova = CountRows(TBL_OTPREMNICA)
    AssertEquals "", CreateOtpremnicaDraft_TX(OtpBrojHeader(broj, d, TEST_ST_ID), _
                                              OtpOcek(100#, 10#, 0#, 0#), razlog), _
                 "OTP broj: pisac odbija isti broj, stanicu i dan"
    AssertEquals CStr(preRedova), CStr(CountRows(TBL_OTPREMNICA)), _
                 "OTP broj: odbijen upis nije ostavio red"

    ' Storno ne oslobadja broj.
    MarkTestRowStornirano TBL_OTPREMNICA, "OtpremnicaID", otpID

    p("datum") = d
    r = modDokUnos.OtpremnicaValidiraj(p, fokus)
    AssertTrue InStr(1, r, zauzeto, vbBinaryCompare) = 1, _
               "OTP broj: storno ne oslobadja broj -- ekran (A9) (bilo: " & r & ")"
    AssertEquals "", CreateOtpremnicaDraft_TX(OtpBrojHeader(broj, d, TEST_ST_ID), _
                                              OtpOcek(100#, 10#, 0#, 0#), razlog), _
                 "OTP broj: storno ne oslobadja broj -- pisac (A9)"

    Exit Sub
EH:
    LogFatal "Test_OTP_BrojZauzetPoStaniciIDanu", Err.Number, Err.description
End Sub

' DRAFT OTPREMNICE: drugi draft istog broja i dana je odbijen, a izmena
' SOPSTVENOG drafta sme da zadrzi broj (izuzimanje po ID-u).
'
' SABOTAZE: ukloni izuzmiID u OtpIzmeniDraft -> pukne "izmena drafta sa svojim
' brojem prolazi"; ukloni poziv u OtpNapraviDraft -> pukne "drugi draft istog
' broja i dana odbijen".
Private Sub Test_OTP_DraftBrojIzuzimaSebe()
    On Error GoTo EH

    Dim scenario As String: scenario = NewScenarioCode("OTPDRZ")
    Dim broj As String: broj = TEST_PREFIX & "-OTP-DRZ-" & scenario

    Dim h1 As Object: Set h1 = OtpHeader(broj)
    Dim razlog As String, id1 As String
    id1 = CreateOtpremnicaDraft_TX(h1, OtpOcek(1000#, 50#, 500#, 25#), razlog)
    AssertTrue Len(id1) > 0, "OTP draft broj: prvi draft nastao (bilo: " & razlog & ")"

    ' OtpHeader pomera datum na svaki poziv, pa se dan izricito izjednacava --
    ' inace drugi draft ide na drugi dan i test ne meri zauzetost.
    Dim h2 As Object: Set h2 = OtpHeader(broj)
    h2("Datum") = h1("Datum")
    Dim razlog2 As String
    AssertEquals "", CreateOtpremnicaDraft_TX(h2, OtpOcek(1000#, 50#, 500#, 25#), razlog2), _
                 "OTP draft broj: drugi draft istog broja i dana odbijen"
    AssertTrue InStr(1, razlog2, id1, vbTextCompare) > 0, _
               "OTP draft broj: razlog imenuje zauzimaca (bilo: " & razlog2 & ")"

    Dim razlog3 As String
    AssertTrue UpdateOtpremnicaDraft_TX(id1, h1, OtpOcek(900#, 45#, 500#, 25#), razlog3), _
               "OTP draft broj: izmena drafta sa svojim brojem prolazi (bilo: " & razlog3 & ")"

    Exit Sub
EH:
    LogFatal "Test_OTP_DraftBrojIzuzimaSebe", Err.Number, Err.description
End Sub

' ZBIRNA: storniran broj ISTOG vozaca istog dana ne sme ponovo (A9, odluka
' 14.09.2026). Dvoklasna zbirna je jedan dokument na dva reda.
'
' SABOTAZE: ukloni poziv u SaveZbirnaMulti_TX -> pukne "storniran broj istog
' vozaca"; premesti ga u SaveZbirna -> pukne "dvoklasna zbirna upisuje obe klase".
Private Sub Test_ZBR_StorniranBrojIstogVozacaOdbijen()
    On Error GoTo EH

    Dim scenario As String: scenario = NewScenarioCode("ZBRBZ")
    Dim d As Date: d = NextTestDate()
    Dim broj As String: broj = TEST_PREFIX & "-ZBR-BZ-" & scenario

    Dim pre As Long: pre = CountRows(TBL_ZBIRNA)
    Dim res As String
    res = SaveZbirnaMulti_TX(d, TEST_VOZ_ID, broj, TEST_KUP_ID, "Test Hladnjaca", "Test Pogon", _
                             TEST_VRSTA, TEST_SORTA, 100#, TEST_TIP_AMB, 10, True, 50#, 5)
    AssertTrue InStr(1, res, " + ", vbBinaryCompare) > 0, _
               "ZBR broj: dvoklasna zbirna upisuje obe klase (bilo: " & res & ")"
    AssertEquals CStr(pre + 2), CStr(CountRows(TBL_ZBIRNA)), _
                 "ZBR broj: dvoklasna zbirna je dva reda"

    Dim ids() As String
    ids = Split(res, " + ")
    MarkTestRowStornirano TBL_ZBIRNA, COL_ZBR_ID, Trim$(ids(0))
    MarkTestRowStornirano TBL_ZBIRNA, COL_ZBR_ID, Trim$(ids(1))

    pre = CountRows(TBL_ZBIRNA)
    AssertEquals "", SaveZbirnaMulti_TX(d, TEST_VOZ_ID, broj, TEST_KUP_ID, "Test Hladnjaca", "Test Pogon", _
                                        TEST_VRSTA, TEST_SORTA, 100#, TEST_TIP_AMB, 10), _
                 "ZBR broj: storniran broj istog vozaca istog dana ne upisuje nov red (A9)"
    AssertEquals CStr(pre), CStr(CountRows(TBL_ZBIRNA)), _
                 "ZBR broj: odbijen upis nije ostavio red"

    AssertTrue Len(SaveZbirnaMulti_TX(d, TEST_VOZ_ID, broj & "-2", TEST_KUP_ID, "Test Hladnjaca", _
                                      "Test Pogon", TEST_VRSTA, TEST_SORTA, 100#, TEST_TIP_AMB, 10)) > 0, _
               "ZBR broj: nov broj istog vozaca istog dana prolazi"

    Exit Sub
EH:
    LogFatal "Test_ZBR_StorniranBrojIstogVozacaOdbijen", Err.Number, Err.description
End Sub

' ZBR-CHILD-01 faza 4: kapija pusta dva aktivna dokumenta kad izbor JESTE scoped.
'
' Ovo je korist zbog koje su faze 1-3 placene. KR-001 scenario -- dva uredjaja bez
' veze posalju zbirnu pod istim brojem, isti vozac i kupac -- danas zaustavlja i
' storno i ponistenje, iako svaki dokument ima svoju decu.
'
' Kontrast u istom testu je bitan:
'   bez generacije -> kapija STOJI. Pozivalac koji ne kaze KOJI dokument stornira
'                     ne moze biti pusten -- pod tim brojem ih je dva.
'   sa generacijom -> kapija PUSTA. Selekcija posle faze 3 dira samo svoju decu.
'
' Deca moraju da dobiju generaciju kroz MasterSync exact-link, ne kroz obican
' upis: cim su oba dokumenta aktivna, ZbirnaGeneracijaZaBroj je fail-closed i
' otpremnica snimljena po broju ostaje bez generacije. Link preko ZbirnaID zna
' tacno cija je.
Private Sub Test_ZBR_KapijaPustaKadJeIzborScoped()
    Dim tx As clsTransaction
    Dim scenario As String, testDate As Date
    Dim broj As String
    Dim zbrA As String, zbrB As String, genA As String, genB As String
    Dim otpA As String, otpB As String
    Dim otkA As String, otkB As String, cridA As String, cridB As String
    Dim r As Object

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRF4")
    testDate = NextTestDate()
    broj = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    otpA = "OTP-ZBRF4-A-" & scenario
    otpB = "OTP-ZBRF4-B-" & scenario
    otkA = "OTK-ZBRF4-A-" & scenario
    otkB = "OTK-ZBRF4-B-" & scenario
    cridA = "CRID-ZBRF4-OA-" & scenario
    cridB = "CRID-ZBRF4-OB-" & scenario

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP

    zbrA = TestHook_ImportZbirnaRowPWA("CRID-ZBRF4-ZA-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 100, broj)
    zbrB = TestHook_ImportZbirnaRowPWA("CRID-ZBRF4-ZB-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 120, broj)
    genA = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrA)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrB)
    AssertTrue (Len(genA) > 0 And Len(genB) > 0 And genA <> genB), _
        "ZBR-F4 preduslov: dva aktivna dokumenta pod istim brojem, razlicite generacije"

    AppendRF28OtpremnicaFixture otpA, testDate, TEST_VOZ_ID, TEST_PREFIX & "-OA-" & scenario
    AppendRF28OtpremnicaFixture otpB, testDate, TEST_VOZ_ID, TEST_PREFIX & "-OB-" & scenario
    AppendRF28OtkupFixture otkA, testDate, TEST_VOZ_ID, cridA, ""
    AppendRF28OtkupFixture otkB, testDate, TEST_VOZ_ID, cridB, ""
    VeziOtkupZaOtpremnicuFixture otkA, otpA
    VeziOtkupZaOtpremnicuFixture otkB, otpB

    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrA, broj, cridA
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrB, broj, cridB

    AssertEquals genA, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
        "ZBR-F4 preduslov: otpremnica A nosi generaciju A"
    AssertEquals genB, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpB), _
        "ZBR-F4 preduslov: otpremnica B nosi generaciju B"

    ' --- BEZ generacije: pozivalac ne kaze KOJI dokument -> kapija STOJI ---
    Set r = RunSimpleStornoZbirna(broj)
    AssertFalse CBool(r("success")), _
        "ZBR-F4: storno BEZ generacije i dalje staje na dva aktivna dokumenta"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrA), _
        "ZBR-F4: posle odbijenog storna dokument A je netaknut"

    ' --- SA generacijom: izbor je scoped -> kapija PUSTA ---
    Set r = RunSimpleStornoZbirna(broj, genB)
    AssertTrue CBool(r("success")), _
        "ZBR-F4: storno SA generacijom prolazi iako broj nosi dva dokumenta"
    AssertTrue RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrB), _
        "ZBR-F4: stornira se bas izabrani dokument B"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrA), _
        "ZBR-F4: dokument A ostaje aktivan"
    AssertEquals "", _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpB, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F4: sopstvena otpremnica B je odvezana"
    AssertEquals broj, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpA, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F4: otpremnica dokumenta A NIJE dirnuta"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 faza 4 kapija pusta scoped izbor", bfpErrDesc
End Sub

' Dva kupca mogu istog dana dobiti ISTI BrojPrijemnice (GenerateBrojPrijemnice
' racuna sekvencu po kupcu). Generacije im moraju biti razlicite.
Private Sub Test_GeneracijaNePrelaziVlasnika()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("GENVLAS")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String, brojPrij As String
    brojZbirne = TEST_PREFIX & "-ZBR-VL-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-VL-" & scenario     ' ISTI broj za oba kupca

    Dim zbrFix As String
    zbrFix = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                           "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                           100#, TEST_TIP_AMB, 0, KLASA_I)
    AssertTrue Len(zbrFix) > 0, "Vlasnik scope: fixture zbirna kreirana"

    Dim prjA As String, prjB As String
    prjA = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    prjB = SavePrijemnica_TX(testDate, TEST_KUP2_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 80#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)

    AssertTrue Len(prjA) > 0 And Len(prjB) > 0, _
               "Vlasnik scope: obe prijemnice (isti broj, razliciti kupci) kreirane"

    Dim genA As String, genB As String
    genA = DokGeneracija(TBL_PRIJEMNICA, COL_PRJ_ID, prjA)
    genB = DokGeneracija(TBL_PRIJEMNICA, COL_PRJ_ID, prjB)

    AssertTrue Len(genA) > 0 And Len(genB) > 0, "Vlasnik scope: obe prijemnice imaju generaciju"
    AssertTrue genA <> genB, _
               "Vlasnik scope: isti broj kod DVA kupca ne deli generaciju"

    ' Klasa II kupca A mora naslediti generaciju kupca A (ne kupca B).
    Dim prjAII As String
    prjAII = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                               TEST_VRSTA, TEST_SORTA, 40#, 90#, TEST_TIP_AMB, 0, 0, KLASA_II)
    AssertTrue Len(prjAII) > 0, "Vlasnik scope: Kl.II kupca A kreirana"
    AssertEquals genA, DokGeneracija(TBL_PRIJEMNICA, COL_PRJ_ID, prjAII), _
                 "Vlasnik scope: Kl.II nasledjuje generaciju SVOG kupca"

    ' Prefill po PK-u kupca A vraca redove kupca A.
    Dim d As Variant
    d = GetTableData(TBL_PRIJEMNICA)

    Dim rI As Long, rII As Long
    PickPrefillRows d, GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ), _
                    GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA), _
                    GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID), _
                    GetColumnIndex(TBL_PRIJEMNICA, COL_GENERACIJA_ID), _
                    brojPrij, prjA, rI, rII

    AssertTrue rI > 0 And rII > 0, "Vlasnik scope: prefill nasao obe klase kupca A"

    Dim cKup As Long
    cKup = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC)
    If rI > 0 Then
        AssertEquals TEST_KUP_ID, Trim$(CStr(nz(d(rI, cKup), ""))), _
                     "Vlasnik scope: prefill Kl.I je kupca A (ne prelazi na kupca B)"
    End If
    If rII > 0 Then
        AssertEquals TEST_KUP_ID, Trim$(CStr(nz(d(rII, cKup), ""))), _
                     "Vlasnik scope: prefill Kl.II je kupca A"
    End If

    Exit Sub

EH:
    LogFatal "Test_GeneracijaNePrelaziVlasnika", Err.Number, Err.description
End Sub

' Generacija reda po ID-u. Prazan ID daje "" -> AssertEquals nad dva prazna bi
' lazno prosao, pa pozivaoci uz poredjenje tvrde i da generacija NIJE prazna.
Private Function DokGeneracija(ByVal tableName As String, ByVal idColumn As String, _
                               ByVal idValue As String) As String
    If Len(Trim$(idValue)) = 0 Then Exit Function

    DokGeneracija = Trim$(CStr(nz(GetValueByKey(tableName, idColumn, idValue, _
                                                COL_GENERACIJA_ID), "")))
End Function

Private Sub Test_MalinaAutoZbirnaFailSignal()
    Dim prevMode As String, prevKupac As String

    On Error GoTo EH

    prevMode = GetConfigValue(CFG_KEY_MALINA_MODE)
    prevKupac = GetConfigValue(CFG_MALINA_DEFAULT_KUPAC)

    SetConfigValue CFG_KEY_MALINA_MODE, "YES"
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, ""

    ' Uslov 1 koji frmDokumenta prijavljuje: poziv baca gresku (nedostaje config).
    ' Scope na sopstveni (nepostojeci) broj -> test ne dira tudje otpremnice.
    Dim scenario As String
    scenario = NewScenarioCode("MALFAIL")

    Dim brojOtpNema As String
    brojOtpNema = TEST_PREFIX & "-OTP-NEMA-" & scenario

    ' Zove se JEZGRO, ne _TX ulaz: produkcioni ulaz je PAUZIRAN dok izvedeni
    ' lanac ne predje na nov model (modMasterSync.IzvedeniLanacIzPwaDostupan).
    ' Obe provere koje ovaj test meri -- raise bez MALINA_DEFAULT_KUPAC i povrat
    ' 0 bez otvorene otpremnice -- zive u jezgru (modMasterSync:1108-1114), pa se
    ' nista ne gubi. Da je ostao na _TX ulazu, prva tvrdnja bi prolazila iz
    ' POGRESNOG razloga: raise bi dolazio od pauze, ne od nedostajuceg configa.
    Dim raised As Boolean
    On Error Resume Next
    Call modMasterSync.AutoCreateZbirnaFromOtpremnice(brojOtpNema)
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    AssertTrue raised, _
               "Malina: bez MALINA_DEFAULT_KUPAC auto-zbirna baca gresku (forma prikazuje poruku)"

    ' Uslov 2: nema otvorene otpremnice u scope-u -> povrat 0 (forma to tretira kao pad).
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, TEST_KUP_ID

    Dim zbrBefore As Long
    zbrBefore = CountRows(TBL_ZBIRNA)

    Dim created As Long
    created = modMasterSync.AutoCreateZbirnaFromOtpremnice(brojOtpNema)
    AssertEquals "0", CStr(created), _
                 "Malina: bez otvorene otpremnice povrat je 0 (forma javlja da zbirna NIJE kreirana)"
    AssertEquals CStr(zbrBefore), CStr(CountRows(TBL_ZBIRNA)), _
                 "Malina: neuspeo run ne dira nepovezane otpremnice (scoped)"

    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    Exit Sub

EH:
    On Error Resume Next
    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    On Error GoTo 0
    LogFatal "Test_MalinaAutoZbirnaFailSignal", Err.Number, Err.description
End Sub

Private Sub Test_PorukeKatalogPokrivaDokumenta()
    On Error GoTo EH

    ' EnsurePoruke je MsgBox-free i idempotentan -> bezbedan u suite-u.
    modSetup.EnsurePoruke
    modPoruke.InvalidateCache

    AssertTrue PorukaPostoji("DOK_MSG_VALIDACIJA_NIJE_PROSLA"), _
               "Poruke: DOK_MSG_VALIDACIJA_NIJE_PROSLA postoji u katalogu"
    AssertTrue PorukaPostoji("DOK_MSG_GRESKA_PRI_CUVANJU"), _
               "Poruke: DOK_MSG_GRESKA_PRI_CUVANJU postoji u katalogu"
    AssertTrue PorukaPostoji("DOK_MSG_GRESKA_PRI_CUVANJU_3"), _
               "Poruke: DOK_MSG_GRESKA_PRI_CUVANJU_3 postoji u katalogu"
    AssertTrue PorukaPostoji("DOK_LBL_NEISPRAVNA_KOLICINA_AMBALAZE"), _
               "Poruke: DOK_LBL_NEISPRAVNA_KOLICINA_AMBALAZE postoji u katalogu"
    AssertTrue PorukaPostoji("DOK_ERR_GRESKA"), _
               "Poruke: DOK_ERR_GRESKA postoji u katalogu"

    Exit Sub

EH:
    LogFatal "Test_PorukeKatalogPokrivaDokumenta", Err.Number, Err.description
End Sub

Private Sub Test_ZbirnaRowDataColumnMapped()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("ZBRMAP")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    brojZbirne = TEST_PREFIX & "-ZBR-MAP-" & scenario

    ' SaveZbirna gradi red PO IMENU kolone (BuildZbirnaRowData), pa svaka vrednost
    ' mora zavrsiti u SVOJOJ koloni -- pozicijski Array(...) je to garantovao samo
    ' dok je redosled kolona tacno onakav kakav je kod pretpostavljao.
    Dim zbrID As String
    zbrID = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          123.45, TEST_TIP_AMB, 7, KLASA_II)

    AssertTrue Len(zbrID) > 0, "Zbirna mapiranje: red snimljen"

    AssertEquals TEST_VOZ_ID, ZbrPolje(zbrID, COL_ZBR_VOZAC), "Zbirna mapiranje: VozacID"
    AssertEquals brojZbirne, ZbrPolje(zbrID, COL_ZBR_BROJ), "Zbirna mapiranje: BrojZbirne"
    AssertEquals TEST_KUP_ID, ZbrPolje(zbrID, COL_ZBR_KUPAC), "Zbirna mapiranje: KupacID"
    AssertEquals "Test Hladnjaca", ZbrPolje(zbrID, COL_ZBR_HLADNJACA), "Zbirna mapiranje: Hladnjaca"
    AssertEquals "Test Pogon", ZbrPolje(zbrID, COL_ZBR_POGON), "Zbirna mapiranje: Pogon"
    AssertEquals TEST_VRSTA, ZbrPolje(zbrID, COL_ZBR_VRSTA), "Zbirna mapiranje: VrstaVoca"
    AssertEquals TEST_SORTA, ZbrPolje(zbrID, COL_ZBR_SORTA), "Zbirna mapiranje: SortaVoca"
    AssertEquals TEST_TIP_AMB, ZbrPolje(zbrID, COL_ZBR_TIP_AMB), "Zbirna mapiranje: TipAmbalaze"
    AssertEquals KLASA_II, ZbrPolje(zbrID, COL_ZBR_KLASA), "Zbirna mapiranje: Klasa"
    AssertEquals "7", ZbrPolje(zbrID, COL_ZBR_KOL_AMB), "Zbirna mapiranje: UkupnoAmbalaze"

    Dim kol As Double
    AssertTrue TryParseDouble(ZbrPolje(zbrID, COL_ZBR_KOLICINA), kol), _
               "Zbirna mapiranje: UkupnoKolicina je broj"
    AssertTrue Abs(kol - 123.45) < 0.001, "Zbirna mapiranje: UkupnoKolicina vrednost"

    ' Datum se cita kao sirova vrednost: Excel ga vraca kao Date ILI kao serijski
    ' broj (zavisi od formata kolone), pa poredjenje ne sme da ide preko CStr.
    Dim vDat As Variant
    vDat = GetValueByKey(TBL_ZBIRNA, COL_ZBR_ID, zbrID, COL_ZBR_DATUM)

    Dim datOk As Boolean
    If IsDate(vDat) Then
        datOk = (Int(CDbl(CDate(vDat))) = Int(CDbl(testDate)))
    ElseIf IsNumeric(vDat) Then
        datOk = (Int(CDbl(vDat)) = Int(CDbl(testDate)))
    End If

    AssertTrue datOk, "Zbirna mapiranje: Datum vrednost u Datum koloni"

    If GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO) > 0 Then
        AssertEquals "", ZbrPolje(zbrID, COL_STORNIRANO), _
                     "Zbirna mapiranje: Stornirano ostaje prazno"
    End If

    Exit Sub

EH:
    LogFatal "Test_ZbirnaRowDataColumnMapped", Err.Number, Err.description
End Sub

Private Sub Test_OMUlazSmerObavezan()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OMSMER")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojDok As String
    brojDok = TEST_PREFIX & "-OMU-" & scenario

    Dim ambBefore As Long
    ambBefore = CountRows(TBL_AMBALAZA)

    ' Prazan smer uz kolicinu ambalaze: ranije je tiho knjizen legacy Stanica ULAZ.
    ' Sada core guard odbija upis (UI dodatno blokira pre poziva).
    Dim ok As Boolean
    ok = SaveOMUlaz_TX(datum:=testDate, brojDok:=brojDok, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=10, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="")

    AssertFalse ok, "OM ulaz: prazan smer uz kolicinu ambalaze je odbijen"
    AssertEquals CStr(ambBefore), CStr(CountRows(TBL_AMBALAZA)), _
                 "OM ulaz: odbijen upis nije ostavio ambalaza red"

    ' Nepoznat smer takodje pada (nije jedan od cetiri dozvoljena).
    ok = SaveOMUlaz_TX(datum:=testDate, brojDok:=brojDok & "-X", _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=10, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="NEPOSTOJECI")

    AssertFalse ok, "OM ulaz: nepoznat smer je odbijen"
    AssertEquals CStr(ambBefore), CStr(CountRows(TBL_AMBALAZA)), _
                 "OM ulaz: nepoznat smer nije ostavio ambalaza red"

    ' Kontrola: eksplicitan smer prolazi (IZDATO_OM = vozac predaje na OM).
    ok = SaveOMUlaz_TX(datum:=testDate, brojDok:=brojDok & "-OK", _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=10, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")

    AssertTrue ok, "OM ulaz: eksplicitan smer IZDATO_OM prolazi"
    AssertEquals CStr(ambBefore + 1), CStr(CountRows(TBL_AMBALAZA)), _
                 "OM ulaz: eksplicitan smer upisao tacno jedan ambalaza red"

    Exit Sub

EH:
    LogFatal "Test_OMUlazSmerObavezan", Err.Number, Err.description
End Sub

Private Function ZbrPolje(ByVal zbirnaID As String, ByVal columnName As String) As String
    ZbrPolje = Trim$(CStr(nz(GetValueByKey(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID, columnName), "")))
End Function

' GetOpenFakture: 1=BrojFakture 2=FakturaID 3=Iznos 4=Uplaceno 5=Preostalo 6=Datum
Private Function OpenFaktureRed(ByVal kupacID As String, ByVal fakturaID As String) As Long
    Dim d As Variant
    d = GetOpenFakture(kupacID)
    If Not IsArray(d) Then Exit Function

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If Trim$(NzToText(d(i, 2))) = Trim$(fakturaID) Then
            OpenFaktureRed = i
            Exit Function
        End If
    Next i
End Function

Private Function OpenFaktureSadrzi(ByVal kupacID As String, ByVal fakturaID As String) As Boolean
    OpenFaktureSadrzi = (OpenFaktureRed(kupacID, fakturaID) > 0)
End Function

Private Function OpenFaktureImaDatum(ByVal kupacID As String, ByVal fakturaID As String) As Boolean
    Dim r As Long
    r = OpenFaktureRed(kupacID, fakturaID)
    If r = 0 Then Exit Function

    Dim d As Variant
    d = GetOpenFakture(kupacID)
    If Not IsArray(d) Then Exit Function

    OpenFaktureImaDatum = IsDate(d(r, 6))
End Function

' Poruka() za nepoznat kljuc vraca "[KLJUC]" -> to je "nedostaje u katalogu".
Private Function PorukaPostoji(ByVal kljuc As String) As Boolean
    Dim t As String
    t = Poruka(kljuc)
    PorukaPostoji = (Len(t) > 0) And (t <> "[" & kljuc & "]")
End Function

Private Sub SeedKupac()
    If RowExists(TBL_KUPCI, "KupacID", TEST_KUP_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_KUPCI)

    SetRequiredField rowData, TBL_KUPCI, "KupacID", TEST_KUP_ID
    SetRequiredField rowData, TBL_KUPCI, "Naziv", "TEST KUPAC DOO"
    SetOptionalField rowData, TBL_KUPCI, "Mesto", "Test Grad"
    SetRequiredField rowData, TBL_KUPCI, "PIB", "109000001"
    SetOptionalField rowData, TBL_KUPCI, "MaticniBroj", "20900001"
    SetOptionalField rowData, TBL_KUPCI, "Ulica", "Test ulica 1"
    SetOptionalField rowData, TBL_KUPCI, "PostanskiBroj", "11000"
    SetOptionalField rowData, TBL_KUPCI, "Drzava", "RS"
    SetOptionalField rowData, TBL_KUPCI, "Email", "test@example.com"
    SetOptionalField rowData, TBL_KUPCI, "Hladnjaca", "Test Hladnjaca"
    SetOptionalField rowData, TBL_KUPCI, "Aktivan", "Aktivan"
    SetOptionalField rowData, TBL_KUPCI, "TekuciRacun", "160-0000000000000-00"

    RequireAppend TBL_KUPCI, rowData, "SeedKupac"
End Sub

Private Sub SeedKultura()
    If RowExists(TBL_KULTURE, "KulturaID", TEST_KULTURA_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_KULTURE)

    SetRequiredField rowData, TBL_KULTURE, "KulturaID", TEST_KULTURA_ID
    SetRequiredField rowData, TBL_KULTURE, "VrstaVoca", TEST_VRSTA
    SetRequiredField rowData, TBL_KULTURE, "SortaVoca", TEST_SORTA
    SetOptionalField rowData, TBL_KULTURE, "Aktivan", "Aktivan"

    RequireAppend TBL_KULTURE, rowData, "SeedKultura"
End Sub

' Kultura bez sorte. Postoji u stvarnosti (dunja se ne vodi po sorti), a
' writer je koristi da dokaze da prazna sorta prolazi TACNO tamo gde je i
' master prazan -- ne uvek i ne nikad.
Private Sub SeedKulturaBezSorte()
    If RowExists(TBL_KULTURE, "KulturaID", TEST_KUL_BEZ_SORTE_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_KULTURE)

    SetRequiredField rowData, TBL_KULTURE, "KulturaID", TEST_KUL_BEZ_SORTE_ID
    SetRequiredField rowData, TBL_KULTURE, "VrstaVoca", TEST_VRSTA_BEZ_SORTE
    SetOptionalField rowData, TBL_KULTURE, "SortaVoca", ""
    SetOptionalField rowData, TBL_KULTURE, "Aktivan", "Aktivan"

    RequireAppend TBL_KULTURE, rowData, "SeedKulturaBezSorte"
End Sub

Private Sub SeedKooperant()
    If RowExists(TBL_KOOPERANTI, "KooperantID", TEST_KOOP_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_KOOPERANTI)

    SetRequiredField rowData, TBL_KOOPERANTI, "KooperantID", TEST_KOOP_ID
    SetRequiredField rowData, TBL_KOOPERANTI, "Ime", "Test"
    SetRequiredField rowData, TBL_KOOPERANTI, "Prezime", "Kooperant"
    SetOptionalField rowData, TBL_KOOPERANTI, "Mesto", "Test Selo"
    SetOptionalField rowData, TBL_KOOPERANTI, "Telefon", "060000002"
    SetRequiredField rowData, TBL_KOOPERANTI, "StanicaID", TEST_ST_ID
    SetOptionalField rowData, TBL_KOOPERANTI, "Aktivan", "Da"
    SetOptionalField rowData, TBL_KOOPERANTI, "BPGBroj", "BPG-TEST-90001"
    SetOptionalField rowData, TBL_KOOPERANTI, "TekuciRacun", "160-0000000000001-00"
    SetOptionalField rowData, TBL_KOOPERANTI, "PIN", "9003"
    SetOptionalField rowData, TBL_KOOPERANTI, "Adresa", "Test adresa 1"
    SetOptionalField rowData, TBL_KOOPERANTI, "JMBG", "0101000710000"

    RequireAppend TBL_KOOPERANTI, rowData, "SeedKooperant"
End Sub

Private Sub SeedParcelaIfAvailable()
    If GetTable(TBL_PARCELE) Is Nothing Then Exit Sub
    If RowExists(TBL_PARCELE, "ParcelaID", TEST_PAR_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_PARCELE)

    SetRequiredField rowData, TBL_PARCELE, "ParcelaID", TEST_PAR_ID
    SetRequiredField rowData, TBL_PARCELE, "KooperantID", TEST_KOOP_ID
    SetRequiredField rowData, TBL_PARCELE, "KatBroj", "TEST-1"
    SetOptionalField rowData, TBL_PARCELE, "KatOpstina", "Test KO"
    SetOptionalField rowData, TBL_PARCELE, "Kultura", TEST_SORTA
    SetOptionalField rowData, TBL_PARCELE, "PovrsinaHa", 1.25
    SetOptionalField rowData, TBL_PARCELE, "GGAPStatus", "DA"
    SetOptionalField rowData, TBL_PARCELE, "Napomena", "Auto test parcela"
    SetOptionalField rowData, TBL_PARCELE, "Aktivna", "Da"
    SetOptionalField rowData, TBL_PARCELE, "Aktivan", "Aktivan"

    RequireAppend TBL_PARCELE, rowData, "SeedParcelaIfAvailable"
End Sub

' ============================================================
' GENERIC TABLE HELPERS
' ============================================================

Private Sub RequireTableExists(ByVal tableName As String)
    If GetTable(tableName) Is Nothing Then
        Err.Raise vbObjectError + 9200, "modBusinessFlowProTests.RequireTableExists", _
                  "Table missing: " & tableName
    End If
End Sub

Private Sub RequireColumnsExist(ByVal tableName As String, ByVal columnNames As Variant)
    Dim c As Variant

    For Each c In columnNames
        RequireCol tableName, CStr(c)
    Next c
End Sub

Private Function RequireCol(ByVal tableName As String, ByVal columnName As String) As Long
    RequireCol = GetColumnIndex(tableName, columnName)

    If RequireCol = 0 Then
        Err.Raise vbObjectError + 9201, "modBusinessFlowProTests.RequireCol", _
                  "Missing column: " & tableName & "." & columnName
    End If
End Function

Private Function BlankRow(ByVal tableName As String) As Variant
    Dim lo As ListObject
    Set lo = GetTable(tableName)

    If lo Is Nothing Then
        Err.Raise vbObjectError + 9202, "modBusinessFlowProTests.BlankRow", _
                  "Table not found: " & tableName
    End If

    Dim arr() As Variant
    ReDim arr(1 To lo.ListColumns.count)
    BlankRow = arr
End Function

Private Sub SetRequiredField(ByRef rowData As Variant, ByVal tableName As String, _
                             ByVal columnName As String, ByVal value As Variant)
    Dim colIdx As Long
    colIdx = RequireCol(tableName, columnName)
    rowData(colIdx) = value
End Sub

' Drugi kooperant, na ISTOJ stanici -- razlika prema prvom je samo vlasnistvo
' parcele. Sve ostalo namerno isto, da test vlasnistva ne bi prosao iz nekog
' drugog razloga.
Private Sub SeedKooperant2()
    If RowExists(TBL_KOOPERANTI, "KooperantID", TEST_KOOP2_ID) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_KOOPERANTI)

    SetRequiredField rowData, TBL_KOOPERANTI, "KooperantID", TEST_KOOP2_ID
    SetRequiredField rowData, TBL_KOOPERANTI, "Ime", "Test"
    SetRequiredField rowData, TBL_KOOPERANTI, "Prezime", "Kooperant Drugi"
    SetOptionalField rowData, TBL_KOOPERANTI, "Mesto", "Test Selo"
    SetRequiredField rowData, TBL_KOOPERANTI, "StanicaID", TEST_ST_ID
    SetOptionalField rowData, TBL_KOOPERANTI, "Aktivan", "Da"

    RequireAppend TBL_KOOPERANTI, rowData, "SeedKooperant2"
End Sub

Private Sub SetOptionalField(ByRef rowData As Variant, ByVal tableName As String, _
                             ByVal columnName As String, ByVal value As Variant)
    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, columnName)

    If colIdx > 0 Then
        rowData(colIdx) = value
    End If
End Sub

Private Sub RequireAppend(ByVal tableName As String, ByVal rowData As Variant, ByVal sourceName As String)
    If AppendRow(tableName, rowData) <= 0 Then
        Err.Raise vbObjectError + 9203, sourceName, "AppendRow failed for " & tableName
    End If
End Sub

Private Function RowExists(ByVal tableName As String, ByVal keyColumn As String, ByVal keyValue As String) As Boolean
    On Error GoTo EH

    If GetTable(tableName) Is Nothing Then Exit Function

    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, keyColumn)
    If colIdx = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colIdx)) = CStr(keyValue) Then
            RowExists = True
            Exit Function
        End If
    Next i

    Exit Function

EH:
    RowExists = False
End Function

Private Function CountRows(ByVal tableName As String) As Long
    On Error GoTo EH

    Dim lo As ListObject
    Set lo = GetTable(tableName)

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    CountRows = lo.DataBodyRange.rows.count
    Exit Function

EH:
    CountRows = 0
End Function

Private Function GetValueByKey(ByVal tableName As String, ByVal keyColumn As String, _
                               ByVal keyValue As String, ByVal returnColumn As String) As Variant
    On Error GoTo EH

    GetValueByKey = LookupValue(tableName, keyColumn, keyValue, returnColumn)
    Exit Function

EH:
    GetValueByKey = Empty
End Function

Private Function GetTestParcelaID() As String
    If GetTable(TBL_PARCELE) Is Nothing Then
        GetTestParcelaID = ""
    ElseIf RowExists(TBL_PARCELE, "ParcelaID", TEST_PAR_ID) Then
        GetTestParcelaID = TEST_PAR_ID
    Else
        GetTestParcelaID = ""
    End If
End Function

Private Sub AssertFalse(ByVal condition As Boolean, ByVal testName As String)
    AssertTrue Not condition, testName
End Sub

Private Sub MarkTestRowStornirano(ByVal tableName As String, _
                                  ByVal idColumn As String, _
                                  ByVal idValue As String)
    Const SRC As String = "MarkTestRowStornirano"

    Dim rows As Collection
    Set rows = FindRows(tableName, idColumn, idValue)

    If rows Is Nothing Or rows.count = 0 Then
        Err.Raise vbObjectError + 9301, SRC, _
                  "Row not found. Table=" & tableName & " ID=" & idValue
    End If

    RequireUpdateCell tableName, CLng(rows(1)), COL_STORNIRANO, "Da", SRC
End Sub

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function ArrayContainsKeyValue(ByRef data As Variant, _
                                       ByVal tableName As String, _
                                       ByVal keyColumn As String, _
                                       ByVal keyValue As String) As Boolean
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function

    Dim colKey As Long
    colKey = RequireColumnIndex(tableName, keyColumn, "ArrayContainsKeyValue")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colKey))) = Trim$(keyValue) Then
            ArrayContainsKeyValue = True
            Exit Function
        End If
    Next i
End Function

Private Function FindIDByTwoColumns(ByVal tableName As String, _
                                    ByVal idColumn As String, _
                                    ByVal keyColumn1 As String, _
                                    ByVal keyValue1 As String, _
                                    ByVal keyColumn2 As String, _
                                    ByVal keyValue2 As String) As String
    Dim data As Variant
    data = GetTableData(tableName)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, tableName)

    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    Dim colKey1 As Long
    Dim colKey2 As Long

    colID = RequireColumnIndex(tableName, idColumn, "FindIDByTwoColumns")
    colKey1 = RequireColumnIndex(tableName, keyColumn1, "FindIDByTwoColumns")
    colKey2 = RequireColumnIndex(tableName, keyColumn2, "FindIDByTwoColumns")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colKey1))) = Trim$(keyValue1) And _
           Trim$(CStr(data(i, colKey2))) = Trim$(keyValue2) Then

            FindIDByTwoColumns = Trim$(CStr(data(i, colID)))
            Exit Function
        End If
    Next i
End Function

' BrojPrijemnice za (BrojZbirne | Klasa | KupacID). FindPrijemnicaIDByBrojAndKlasa
' ne moze ovde: trazi po BROJU prijemnice, a kod izolacije po kupcu dve prijemnice
' dele isti BrojZbirne pa je kupac deo kljuca.
Private Function FindPrijBrojByZbirnaKlasaKupac(ByVal brZbr As String, ByVal klasa As String, _
                                                ByVal kupacID As String) As String
    Const SRC As String = "FindPrijBrojByZbirnaKlasaKupac"
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function

    Dim cZbr As Long, cKla As Long, cKup As Long, cBroj As Long
    cZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
    cKla = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, SRC)
    cKup = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, SRC)
    cBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = Trim$(brZbr) And _
           Trim$(CStr(data(i, cKla))) = Trim$(klasa) And _
           StrComp(Trim$(CStr(data(i, cKup))), Trim$(kupacID), vbTextCompare) = 0 Then
            FindPrijBrojByZbirnaKlasaKupac = Trim$(CStr(data(i, cBroj)))
            Exit Function
        End If
    Next i
End Function

' ============================================================
' RUN / SCENARIO HELPERS
' ============================================================

Private Sub BeginRun(ByVal suiteName As String)
    ResetCounters
    InitTestLog

    Randomize
    m_RunID = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int((9999 - 1000 + 1) * Rnd + 1000))
    m_DateSeq = 0

    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "RunID=" & m_RunID
    Debug.Print String$(70, "=")

    ' Disclaimer -- NAMERNO ne-modalni (Debug.Print, ne MsgBox): suite se pokrece i
    ' automatizovano iz modE2EReleaseGate (Application.Run), gde bi modal blokirao
    ' ceo release gate. Suite PISE u radnu svesku (otkup/dokumenti/fakture/palete,
    ' i privremeno menja config), pa se pokrece nad TEST kopijom, ne nad klijentskim
    ' fajlom. Test podaci nose prefiks TST-PRO / ID-eve *-9000x i ne ciste se.
    Debug.Print "UPOZORENJE: suite UPISUJE podatke u ovu svesku (i privremeno menja"
    Debug.Print "            config). Pokretati SAMO nad test kopijom, ne nad"
    Debug.Print "            klijentskim fajlom. Test redovi (TST-PRO / *-9000x) ostaju."
    Debug.Print String$(70, "-")

    AppendTestLog "SUITE", suiteName, "START", "RunID=" & m_RunID
    AppendTestLog "SUITE", suiteName, "WARN", _
                  "Suite upisuje u svesku -- pokretati samo nad test kopijom."
End Sub

Private Sub EndRun()
    Dim summary As String

    summary = "RunID=" & m_RunID & _
              " | Total=" & m_Total & _
              " | Passed=" & m_Passed & _
              " | Failed=" & m_Failed & _
              " | Skipped=" & m_Skipped

    Debug.Print String$(70, "-")
    Debug.Print "BUSINESS FLOW PRO TEST SUMMARY: " & summary
    Debug.Print String$(70, "-")

    AppendTestLog "SUITE", "SUMMARY", "INFO", summary

    ' PRE MsgBox-a i PRE Err.Raise: gate iskace iz procedure, pa bi upis
    ' posle njega izostao bas kad suite padne -- dakle kad je jedino i treba.
    On Error Resume Next
    WriteResultFileBFP
    On Error GoTo 0

    If m_Failed > 0 Then
        MsgBox "Business Flow Pro tests finished with failures." & vbCrLf & summary, _
               vbExclamation, APP_NAME
    Else
        MsgBox "Business Flow Pro tests finished." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If

    ' Gate. Zove se iz sva cetiri runnera, pa je ovo jedina tacka verdikta.
    If m_Failed > 0 Then
        Err.Raise ERR_BFP_SUITE_FAILED, "modBusinessFlowProTests.EndRun", _
            "Business Flow Pro: " & CStr(m_Failed) & " provera palo (PASS=" & _
            CStr(m_Passed) & "). Detalji u Immediate prozoru."
    End If
End Sub

Private Sub ResetCounters()
    m_Report = ""
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

' ============================================================
' PR3 -- Zbirna: header + stavke, izvedene iz izvornih otpremnica
' ============================================================
'
' Sta se ovde MERI, a sta ne:
'
'   meri se     da CreateZbirna_TX pravi JEDAN header, IZVODI stavke iz izvornih
'               otpremnica, upisuje clanstvo u tblZbirnaIzvori u ISTOJ
'               transakciji, i da header vise ne nosi kolicinu
'   ne meri se  ponasanje citalaca, invarijante i storna -- to je Zbirna
'               cutover. Do tada je stari writer (SaveZbirnaMulti_TX) jedini
'               put, a golden scenariji to i dalje dokazuju, nepromenjeni.

Private Sub Test_PR3_CreateZbirnaHeaderIStavke()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3HS")

    Dim otpI As String, otpII As String
    otpI = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3-" & scenario, KLASA_I, 400#, 20)
    otpII = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3-" & scenario, KLASA_II, 600#, 30)
    AssertTrue Len(otpI) > 0 And Len(otpII) > 0, "PR3: izvorne otpremnice napravljene"

    Dim zbrID As String
    zbrID = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3-" & scenario), Pr3Izvor(otpI, otpII))

    AssertTrue Len(zbrID) > 0, "PR3: CreateZbirna_TX vraca ID"

    ' JEDAN header, ne dva reda po klasi.
    Dim headeri As Collection
    Set headeri = FindRows(TBL_ZBIRNA, COL_ZBR_ID, zbrID)
    AssertTrue Not headeri Is Nothing, "PR3: header pronadjen"
    AssertEquals "1", CStr(headeri.count), "PR3: tacno jedan header red"

    AssertEquals TEST_PREFIX & "-ZBR-PR3-" & scenario, _
                 ZbrPolje(zbrID, COL_ZBR_BROJ), "PR3: BrojZbirne"
    AssertEquals TEST_KUP_ID, ZbrPolje(zbrID, COL_ZBR_KUPAC), "PR3: KupacID"
    AssertEquals TEST_VOZ_ID, ZbrPolje(zbrID, COL_ZBR_VOZAC), "PR3: VozacID"

    ' Vrsta/sorta/tip NISU dosli iz headera -- izvedeni su iz otpremnica.
    AssertEquals TEST_VRSTA, ZbrPolje(zbrID, COL_ZBR_VRSTA), _
                 "PR3: VrstaVoca izvedena iz otpremnice"
    AssertEquals TEST_SORTA, ZbrPolje(zbrID, COL_ZBR_SORTA), _
                 "PR3: SortaVoca izvedena iz otpremnice"
    AssertEquals TEST_TIP_AMB, ZbrPolje(zbrID, COL_ZBR_TIP_AMB), _
                 "PR3: TipAmbalaze izveden iz otpremnice"

    ' DVE stavke, po jedna po klasi, RedniBroj u kanonskom redu (I pa II).
    AssertEquals "2", CStr(Pr3BrojStavki(zbrID)), "PR3: dve stavke"
    AssertEquals "1", Pr3StavkaPolje(zbrID, KLASA_I, COL_ZBS_RB), "PR3: I ima RB 1"
    AssertEquals "2", Pr3StavkaPolje(zbrID, KLASA_II, COL_ZBS_RB), "PR3: II ima RB 2"

    AssertTrue Abs(Pr3StavkaBroj(zbrID, KLASA_I, COL_ZBS_KOLICINA) - 400#) < 0.001, _
               "PR3: Klasa I kolicina 400"
    AssertTrue Abs(Pr3StavkaBroj(zbrID, KLASA_II, COL_ZBS_KOLICINA) - 600#) < 0.001, _
               "PR3: Klasa II kolicina 600"
    AssertTrue Abs(Pr3StavkaBroj(zbrID, KLASA_I, COL_ZBS_KOL_AMB) - 20#) < 0.001, _
               "PR3: Klasa I ambalaza 20"
    AssertTrue Abs(Pr3StavkaBroj(zbrID, KLASA_II, COL_ZBS_KOL_AMB) - 30#) < 0.001, _
               "PR3: Klasa II ambalaza 30"

    AssertEquals zbrID, Pr3StavkaPolje(zbrID, KLASA_I, COL_ZBS_ZBIRNA_ID), _
                 "PR3: stavka pokazuje na ZbirnaID"

    ' MEMBERSHIP: obe izvorne otpremnice nose novi ZbirnaID. Bez ovoga bi zbirna
    ' postojala bez ijednog izvora, a invarijanta nema sta da sabira.
    AssertEquals zbrID, Pr3OtpZbirnaID(otpI), "PR3: otpremnica I vezana za zbirnu"
    AssertEquals zbrID, Pr3OtpZbirnaID(otpII), "PR3: otpremnica II vezana za zbirnu"

    Exit Sub

EH:
    LogFatal "Test_PR3_CreateZbirnaHeaderIStavke", Err.Number, Err.description
End Sub

' Dve otpremnice ISTE klase daju JEDNU stavku sa zbirom -- stavka je "jedna klasa
' jedne zbirne", ne "jedna otpremnica".
Private Sub Test_PR3_DveOtpremniceIsteKlaseSeSabiraju()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3SB")

    Dim a As String, b As String
    a = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3A-" & scenario, KLASA_I, 400#, 20)
    b = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3B-" & scenario, KLASA_I, 250#, 12)

    Dim zbrID As String
    zbrID = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3S-" & scenario), Pr3Izvor(a, b))

    AssertTrue Len(zbrID) > 0, "PR3 zbir: dokument napravljen"
    AssertEquals "1", CStr(Pr3BrojStavki(zbrID)), "PR3 zbir: jedna stavka za jednu klasu"
    AssertTrue Abs(Pr3StavkaBroj(zbrID, KLASA_I, COL_ZBS_KOLICINA) - 650#) < 0.001, _
               "PR3 zbir: 400 + 250 = 650"
    AssertTrue Abs(Pr3StavkaBroj(zbrID, KLASA_I, COL_ZBS_KOL_AMB) - 32#) < 0.001, _
               "PR3 zbir: 20 + 12 = 32"

    AssertEquals zbrID, Pr3OtpZbirnaID(a), "PR3 zbir: prva otpremnica vezana"
    AssertEquals zbrID, Pr3OtpZbirnaID(b), "PR3 zbir: druga otpremnica vezana"

    Exit Sub

EH:
    LogFatal "Test_PR3_DveOtpremniceIsteKlaseSeSabiraju", Err.Number, Err.description
End Sub

' Header NE nosi kolicinu, ambalazu ni klasu -- to su kolone koje u ciljnoj semi
' ne postoje. Prazno je tacan odgovor: "ne pitaj header za kolicinu".
'
' Bez ovog testa bi neko u cutover-u mogao "za svaki slucaj" da upise i zbir
' i time napravio dva izvora istine za istu vrednost -- tacno bolest koju
' header+stavke uklanja.
Private Sub Test_PR3_HeaderNeNosiKolicinu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3HP")

    Dim zbrID As String
    zbrID = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3P-" & scenario), _
        Pr3Izvor(Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3P-" & scenario, _
                               KLASA_I, 400#, 20), ""))

    AssertTrue Len(zbrID) > 0, "PR3 header: dokument napravljen"

    AssertEquals "", ZbrPolje(zbrID, COL_ZBR_KOLICINA), _
                 "PR3 header: UkupnoKolicina ostaje prazna"
    AssertEquals "", ZbrPolje(zbrID, COL_ZBR_KOL_AMB), _
                 "PR3 header: UkupnoAmbalaze ostaje prazna"
    AssertEquals "", ZbrPolje(zbrID, COL_ZBR_KLASA), _
                 "PR3 header: Klasa ostaje prazna"

    ' GeneracijaID je kompenzacija za nepostojeci header i brise se u cutover-u.
    ' Nov pisac je ne sme ozivljavati.
    If GetColumnIndex(TBL_ZBIRNA, COL_GENERACIJA_ID) > 0 Then
        AssertEquals "", ZbrPolje(zbrID, COL_GENERACIJA_ID), _
                     "PR3 header: GeneracijaID se ne pise"
    End If

    Exit Sub

EH:
    LogFatal "Test_PR3_HeaderNeNosiKolicinu", Err.Number, Err.description
End Sub

' ZbirnaID i ZbirnaStavkaID su OPAQUE. Stari GetNextID je davao ZBR-1, ZBR-2 ...
' -- brojac po kome se moze pogoditi "sledeci" i iz koga se cita koliko je
' dokumenata u sistemu. Novi ID nosi samo identitet.
Private Sub Test_PR3_ZbirnaIDJeOpaque()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3ID")

    Dim a As String, b As String
    a = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3A-" & scenario), _
        Pr3Izvor(Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3IA-" & scenario, _
                               KLASA_I, 100#, 5), ""))
    b = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3B-" & scenario), _
        Pr3Izvor(Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3IB-" & scenario, _
                               KLASA_I, 100#, 5), ""))

    AssertTrue Len(a) > 0 And Len(b) > 0, "PR3 ID: oba dokumenta napravljena"
    AssertTrue a <> b, "PR3 ID: dva poziva daju razlicit ID"

    AssertTrue Pr3JeOpaqueID(a, "ZBR-"), "PR3 ID: header ZBR- + 32 hex"
    AssertTrue Pr3JeOpaqueID(Pr3StavkaPolje(a, KLASA_I, COL_ZBS_ID), "ZBS-"), _
               "PR3 ID: stavka ZBS- + 32 hex"

    Exit Sub

EH:
    LogFatal "Test_PR3_ZbirnaIDJeOpaque", Err.Number, Err.description
End Sub

' Ista otpremnica navedena dvaput bi joj duplirala kolicinu u kesu.
Private Sub Test_PR3_IstaOtpremnicaDvaputNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3DV")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3D-" & scenario, KLASA_I, 400#, 20)

    Dim preH As Long, preS As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)
    preS = Pr3BrojRedova(TBL_ZBIRNA_STAVKE)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3D-" & scenario), _
                          Pr3Izvor(otp, otp), razlog)

    AssertEquals "", rez, "PR3 duplikat: upis odbijen"
    AssertTrue InStr(1, razlog, "navedena dvaput", vbTextCompare) > 0, _
               "PR3 duplikat: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 duplikat: nijedan header nije ostao"
    AssertEquals CStr(preS), CStr(Pr3BrojRedova(TBL_ZBIRNA_STAVKE)), _
                 "PR3 duplikat: nijedna stavka nije ostala"
    AssertEquals "", Pr3OtpZbirnaID(otp), "PR3 duplikat: otpremnica nije vezana"

    Exit Sub

EH:
    LogFatal "Test_PR3_IstaOtpremnicaDvaputNeProlazi", Err.Number, Err.description
End Sub

' Otpremnica koja vec pripada nekoj zbirnoj ne sme se tiho preuzeti.
Private Sub Test_PR3_VecVezanaOtpremnicaSeNePreuzima()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3VV")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3V-" & scenario, KLASA_I, 400#, 20)

    Dim prva As String
    prva = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3V1-" & scenario), _
                           Pr3Izvor(otp, ""))
    AssertTrue Len(prva) > 0, "PR3 preuzimanje: prva zbirna napravljena"

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3V2-" & scenario), _
                          Pr3Izvor(otp, ""), razlog)

    AssertEquals "", rez, "PR3 preuzimanje: druga zbirna odbijena"

    ' Poruka mora doci iz KANONSKE grane -- iz tblZbirnaIzvori. Tvrdnja ide na
    ' tekst razloga, ne samo na ishod: kad bi kapija odbila upis iz nekog drugog
    ' razloga, ovaj test bi i dalje bio zelen a clanstvo neprovereno.
    AssertTrue InStr(1, razlog, "u sastavu aktivne zbirne", vbTextCompare) > 0, _
               "PR3 preuzimanje: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 preuzimanje: drugi header nije ostao"
    ' Citac i dalje daje PRVU zbirnu -- odbijen pokusaj nije nista pomerio.
    AssertEquals prva, Pr3OtpZbirnaID(otp), _
                 "PR3 preuzimanje: otpremnica ostaje na prvoj zbirnoj"

    Exit Sub

EH:
    LogFatal "Test_PR3_VecVezanaOtpremnicaSeNePreuzima", Err.Number, Err.description
End Sub

' Stornirana otpremnica nije izvor. Pad na DRUGOJ otpremnici ne sme da ostavi
' prvu vezanu -- prevalidacija ide pre ijednog upisa, transakcija je mreza ispod.
Private Sub Test_PR3_StorniranIzvorNeOstavljaPolaDokumenta()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3ST")

    Dim dobra As String, losa As String
    dobra = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3G-" & scenario, KLASA_I, 400#, 20)
    losa = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3L-" & scenario, KLASA_II, 600#, 30)
    AssertTrue StornoOtpremnica_TX(losa), "PR3 storno: druga otpremnica stornirana"

    Dim preH As Long, preS As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)
    preS = Pr3BrojRedova(TBL_ZBIRNA_STAVKE)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3T-" & scenario), _
                          Pr3Izvor(dobra, losa), razlog)

    AssertEquals "", rez, "PR3 storno: upis odbijen"
    AssertTrue InStr(1, razlog, "stornirana", vbTextCompare) > 0, _
               "PR3 storno: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 storno: header nije ostao"
    AssertEquals CStr(preS), CStr(Pr3BrojRedova(TBL_ZBIRNA_STAVKE)), _
                 "PR3 storno: stavka nije ostala"
    AssertEquals "", Pr3OtpZbirnaID(dobra), _
                 "PR3 storno: ispravna otpremnica NIJE vezana (nema pola dokumenta)"

    Exit Sub

EH:
    LogFatal "Test_PR3_StorniranIzvorNeOstavljaPolaDokumenta", Err.Number, Err.description
End Sub

' Dokument ima JEDNU vrstu voca. Otpremnica koja se ne slaze ne pripada ovoj
' zbirnoj -- inace bi header nosio vrstu prve, a sadrzao robu druge.
Private Sub Test_PR3_RazlicitaVrstaNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3VR")

    Dim a As String, b As String
    a = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3R1-" & scenario, KLASA_I, 400#, 20)
    b = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3R2-" & scenario, KLASA_II, 600#, 30, _
                      TEST_KUL_BEZ_SORTE_ID)

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3R-" & scenario), _
                          Pr3Izvor(a, b), razlog)

    AssertEquals "", rez, "PR3 vrsta: upis odbijen"
    AssertTrue InStr(1, razlog, "VrstaVoca", vbTextCompare) > 0, _
               "PR3 vrsta: kapija imenuje polje (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 vrsta: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_PR3_RazlicitaVrstaNeProlazi", Err.Number, Err.description
End Sub

' Zbirna bez ijedne izvorne otpremnice nije zbirna.
'
' Ovo je nalaz zbog kog je writer i prepravljen: prva verzija je primala gotove
' stavke, pa je bilo legalno napraviti zbirnu od izmisljenih 400+600 kg bez
' ijedne otpremnice -- kes koji se ne slaze sa izvorom, kroz kanonski writer.
Private Sub Test_PR3_ZbirnaBezIzvoraNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3BI")

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3E-" & scenario), _
                          New Collection, razlog)

    AssertEquals "", rez, "PR3 bez izvora: upis odbijen"
    AssertTrue InStr(1, razlog, "bar jednu izvornu otpremnicu", vbTextCompare) > 0, _
               "PR3 bez izvora: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 bez izvora: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_PR3_ZbirnaBezIzvoraNeProlazi", Err.Number, Err.description
End Sub

' Nepoznat kljuc u headeru je GRESKA, ne tiho ignorisanje.
'
' Tipfeler u OPCIONOM polju je inace nevidljiv: "Hladnjca" ne obara nista, samo
' tiho ostavi prazno. Isto vazi za polja koja vise NE PRIPADAJU headeru --
' VrstaVoca dolazi iz otpremnica, pa pozivalac koji je salje pravi drugi izvor
' istine i mora to da cuje.
Private Sub Test_PR3_NepoznatKljucUHeaderuPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3NK")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3N-" & scenario, KLASA_I, 400#, 20)

    Dim h As Object
    Set h = Pr3Header(TEST_PREFIX & "-ZBR-PR3N-" & scenario)
    h.Add "VrstaVoca", TEST_VRSTA           ' vise ne pripada headeru

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(h, Pr3Izvor(otp, ""), razlog)

    AssertEquals "", rez, "PR3 nepoznat kljuc: upis odbijen"
    AssertTrue InStr(1, razlog, "nepoznat kljuc", vbTextCompare) > 0 And _
               InStr(1, razlog, "VrstaVoca", vbTextCompare) > 0, _
               "PR3 nepoznat kljuc: kapija imenuje kljuc (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 nepoznat kljuc: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_PR3_NepoznatKljucUHeaderuPada", Err.Number, Err.description
End Sub

Private Sub Test_PR3_NedostajuciObavezniKljucPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3OK")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3O-" & scenario, KLASA_I, 400#, 20)

    Dim h As Object
    Set h = Pr3Header(TEST_PREFIX & "-ZBR-PR3O-" & scenario)
    h.Remove "KupacID"

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(h, Pr3Izvor(otp, ""), razlog)

    AssertEquals "", rez, "PR3 obavezan kljuc: upis odbijen"
    AssertTrue InStr(1, razlog, "KupacID", vbTextCompare) > 0, _
               "PR3 obavezan kljuc: kapija imenuje polje (bilo: " & razlog & ")"
    AssertEquals "", Pr3OtpZbirnaID(otp), _
                 "PR3 obavezan kljuc: otpremnica nije vezana"

    Exit Sub

EH:
    LogFatal "Test_PR3_NedostajuciObavezniKljucPada", Err.Number, Err.description
End Sub

' Ono sto je operater otkucao mora da se slaze sa onim sto daju otpremnice.
' Neslaganje je danas tiha greska: ekran prikaze svoje brojeve, tabela nosi druge.
Private Sub Test_PR3_OcekivanoKojeSeNeSlazePada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3OC")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3C-" & scenario, KLASA_I, 400#, 20)

    Dim ocek As Collection
    Set ocek = New Collection
    ocek.Add Pr3Ocekivano(KLASA_I, 450#, 20)      ' 450 != 400

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirna_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3C-" & scenario), _
                          Pr3Izvor(otp, ""), ocek, razlog)

    AssertEquals "", rez, "PR3 ocekivano: upis odbijen"
    AssertTrue InStr(1, razlog, "ne slaze sa otpremnicama", vbTextCompare) > 0, _
               "PR3 ocekivano: kapija imenuje neslaganje (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 ocekivano: header nije ostao"

    ' Isto ocekivano, ali TACNO -> prolazi. Bez ovoga bi provera koja uvek pada
    ' izgledala isto kao provera koja radi.
    Dim ocekOk As Collection
    Set ocekOk = New Collection
    ocekOk.Add Pr3Ocekivano(KLASA_I, 400#, 20)

    AssertTrue Len(CreateZbirna_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3C2-" & scenario), _
        Pr3Izvor(otp, ""), ocekOk, razlog)) > 0, _
        "PR3 ocekivano: tacno ocekivanje prolazi"

    Exit Sub

EH:
    LogFatal "Test_PR3_OcekivanoKojeSeNeSlazePada", Err.Number, Err.description
End Sub

' Ambalaza je BROJ KOMADA. Legacy ValidateZbirnaInput je to drzao tipom
' (ukupnoAmb As Long); dictionary writer nema tip koji to cuva, pa se trazi
' izricito. Odbija se, ne zaokruzuje: "20.5 gajbica" je kvar u izvoru, a tiha
' ispravka ga sakriva.
Private Sub Test_PR3_AmbalazaMoraBitiCeoBroj()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3AM")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3M-" & scenario, KLASA_I, 400#, 20)
    Pr3PostaviAmbalazu otp, 1.5

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3M-" & scenario), _
                          Pr3Izvor(otp, ""), razlog)

    AssertEquals "", rez, "PR3 ambalaza: upis odbijen"
    AssertTrue InStr(1, razlog, "mora biti ceo broj", vbTextCompare) > 0, _
               "PR3 ambalaza: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 ambalaza: header nije ostao"

    ' Pokvarena stavka se VRACA: citalac stavki je strog za celu tabelu, pa bi
    ' 1,5 gajbe oborio svaki sledeci test koji cita otpremnice.
    Pr3PostaviAmbalazu otp, 20
    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    On Error Resume Next
    Pr3PostaviAmbalazu otp, 20
    On Error GoTo 0
    LogFatal "Test_PR3_AmbalazaMoraBitiCeoBroj", errNum, errDesc
End Sub

' Otpremnica NEMA kolonu koja pokazuje na zbirnu -- pripadnost je u
' tblZbirnaIzvori. Test to i tvrdi: kanonska pozicija "ZbirnaID" je nula.
Private Sub Test_PR3_OtpremnicaNemaZbirnaID()
    On Error GoTo EH

    ' Clanstvo je jedina veza -- otpremnica NEMA kolonu koja pokazuje na zbirnu.
    ' Kanon je izvor: kolona je izbacena iz schema.json, pa nova sveska je nema.
    ' Zatecena razvojna sveska moze imati mrtvu kolonu iza kanonskog prefiksa;
    ' to je bezopasno i nestaje pri sledecoj izgradnji fixture-a.
    AssertEquals "0", CStr(Pr3KanonskaPozicija(TBL_OTPREMNICA, "ZbirnaID")), _
                 "PR3: ZbirnaID nije u kanonu tblOtpremnica"

    AssertTrue GetColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID) > 0, _
               "PR3: clanstvo nosi OtpremnicaID"

    ' Stavke nemaju svoj Stornirano: line-level storno u domenu ne postoji,
    ' status drzi header. Kolona koje nema ne moze da se filtrira, pa je tabela
    ' u modSchemaGuard.BEZ_STORNA.
    AssertEquals "0", CStr(GetColumnIndex(TBL_ZBIRNA_STAVKE, COL_STORNIRANO)), _
                 "PR3: tblZbirnaStavke nema kolonu Stornirano"

    Exit Sub

EH:
    LogFatal "Test_PR3_OtpremnicaNemaZbirnaID", Err.Number, Err.description
End Sub

' Red bez identiteta je gori od pada: niko ga posle ne moze ni naci ni vezati.
' NewEntityID zato vraca "" kad CoCreateGuid ne uspe, a pozivalac MORA da stane.
'
' Ta grana se u praksi nikad ne desi, pa je provera kod pozivaoca bila
' NEDOKAZIVA -- sabotaza koja je ukloni ostavljala je suite zelen. Seam
' NewEntityIDPadniTest je jedini nacin da se pokaze da kapija stvarno grize.
Private Sub Test_PR3_PrazanHeaderIDNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3IH")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3IH-" & scenario, KLASA_I, 400#, 20)

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    ' Seam je tvrdo gejtovan test rezimom (modDataAccess), a
    ' RunBusinessFlowProSuite ga -- za razliku od RunAllTests i RunGoldenSuite --
    ' ne pali. Bez ovoga bi seam bio inertan i test bi merio da kapija NE grize.
    ' Rezim se vraca i na putu greske: test rezim samo SUZAVA ponasanje, ali
    ' zaostao bi progutao zavrsni MsgBox suite-a i run_vba bi ostao bez verdikta.
    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    modDataAccess.NewEntityIDPadniTest True          ' pada odmah -> header ID
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3IH-" & scenario), _
                          Pr3Izvor(otp, ""), razlog)
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode

    AssertEquals "", rez, "PR3 prazan ID: upis odbijen"
    AssertTrue InStr(1, razlog, "nije vratio ZbirnaID", vbTextCompare) > 0, _
               "PR3 prazan ID: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 prazan ID: header nije ostao"
    AssertEquals "", Pr3OtpZbirnaID(otp), "PR3 prazan ID: otpremnica nije vezana"

    Exit Sub

EH:
    ' Seam se gasi i na putu greske -- zaostao bi svakom sledecem upisu dao
    ' prazan ID, pa bi sledeci test pao BEZ SVOJE KRIVICE.
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode
    LogFatal "Test_PR3_PrazanHeaderIDNeProlazi", Err.Number, Err.description
End Sub

' Isto za STAVKU, i to je druga kapija: header je vec upisan sa ispravnim ID-em,
' pa bi bez provere stavka legla sa praznim PK i transakcija bi commitovala.
'
' Seam propusta PRVI poziv (header) pa obara sledeci (stavka). Bez brojanja bi
' pao header i stavka se ne bi ni pokusala -- merila bi se ista kapija dvaput.
Private Sub Test_PR3_PrazanStavkaIDNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3IS")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3IS-" & scenario, KLASA_I, 400#, 20)

    Dim preH As Long, preS As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)
    preS = Pr3BrojRedova(TBL_ZBIRNA_STAVKE)

    Dim rez As String, razlog As String
    ' Seam je tvrdo gejtovan test rezimom (modDataAccess), a
    ' RunBusinessFlowProSuite ga -- za razliku od RunAllTests i RunGoldenSuite --
    ' ne pali. Bez ovoga bi seam bio inertan i test bi merio da kapija NE grize.
    ' Rezim se vraca i na putu greske: test rezim samo SUZAVA ponasanje, ali
    ' zaostao bi progutao zavrsni MsgBox suite-a i run_vba bi ostao bez verdikta.
    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    modDataAccess.NewEntityIDPadniTest True, 1       ' header prodje, stavka pada
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3IS-" & scenario), _
                          Pr3Izvor(otp, ""), razlog)
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode

    AssertEquals "", rez, "PR3 prazan ZBS: upis odbijen"
    AssertTrue InStr(1, razlog, "nije vratio ZbirnaStavkaID", vbTextCompare) > 0, _
               "PR3 prazan ZBS: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 prazan ZBS: vec upisan header je ROLLBACK-ovan"
    AssertEquals CStr(preS), CStr(Pr3BrojRedova(TBL_ZBIRNA_STAVKE)), _
                 "PR3 prazan ZBS: stavka nije ostala"
    AssertEquals "", Pr3OtpZbirnaID(otp), "PR3 prazan ZBS: otpremnica nije vezana"

    Exit Sub

EH:
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode
    LogFatal "Test_PR3_PrazanStavkaIDNeProlazi", Err.Number, Err.description
End Sub

' Sastav verzije se cita iz tblZbirnaIzvori, ne iz trenutnog stanja otpremnica.
'
' Ovo je tabela zbog koje jedan mutable FK nije dovoljan: posle ispravke jedne
' otpremnice nastaje NOVA verzija zbirne, a sestre koje se nisu menjale pripadaju
' i staroj i novoj. Otpremnica.ZbirnaID moze da pokaze samo jednu.
Private Sub Test_PR3_ClanstvoJeZapisanoPoVerziji()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3CL")

    Dim a As String, b As String, c As String
    a = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3C1-" & scenario, KLASA_I, 400#, 20)
    b = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3C2-" & scenario, KLASA_I, 250#, 12)
    c = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3C3-" & scenario, KLASA_II, 600#, 30)

    Dim izvor As Collection
    Set izvor = New Collection
    izvor.Add a
    izvor.Add b
    izvor.Add c

    Dim zbrID As String
    zbrID = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3C-" & scenario), izvor)
    AssertTrue Len(zbrID) > 0, "PR3 clanstvo: dokument napravljen"

    AssertEquals "3", CStr(Pr3BrojIzvora(zbrID)), _
                 "PR3 clanstvo: tri zapisa clanstva"
    AssertTrue Pr3JeIzvor(zbrID, a), "PR3 clanstvo: prva otpremnica u sastavu"
    AssertTrue Pr3JeIzvor(zbrID, b), "PR3 clanstvo: druga otpremnica u sastavu"
    AssertTrue Pr3JeIzvor(zbrID, c), "PR3 clanstvo: treca otpremnica u sastavu"

    ' Isti odgovor i kroz javni citac, koji ga RACUNA iz clanstva.
    AssertEquals zbrID, Pr3OtpZbirnaID(a), "PR3 clanstvo: citac daje istu zbirnu"
    AssertEquals zbrID, Pr3OtpZbirnaID(b), "PR3 clanstvo: citac daje istu zbirnu (b)"
    AssertEquals zbrID, Pr3OtpZbirnaID(c), "PR3 clanstvo: citac daje istu zbirnu (c)"

    Exit Sub

EH:
    LogFatal "Test_PR3_ClanstvoJeZapisanoPoVerziji", Err.Number, Err.description
End Sub

' Javni citac daje ISTU zbirnu koju kaze zapis clanstva.
'
' AktivnaZbirnaZaOtpremnicu je jedina zamena za obrisanu kolonu
' Otpremnica.ZbirnaID -- racuna odgovor iz tblZbirnaIzvori umesto da ga cuva na
' drugom mestu. Test poredi citac sa sirovim zapisom, da racunanje ne bi tiho
' odgovaralo drugacije od kanona.
Private Sub Test_PR3_CitacDajeIstuZbirnuKaoClanstvo()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3PK")

    Dim a As String, b As String
    a = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3P1-" & scenario, KLASA_I, 400#, 20)
    b = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3P2-" & scenario, KLASA_II, 600#, 30)

    Dim zbrID As String
    zbrID = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3P-" & scenario), _
                            Pr3Izvor(a, b))
    AssertTrue Len(zbrID) > 0, "PR3 citac: dokument napravljen"

    AssertEquals Pr3ZbirnaIzClanstva(a), Pr3OtpZbirnaID(a), _
                 "PR3 citac: citac i clanstvo daju istu zbirnu (a)"
    AssertEquals Pr3ZbirnaIzClanstva(b), Pr3OtpZbirnaID(b), _
                 "PR3 citac: citac i clanstvo daju istu zbirnu (b)"

    Exit Sub

EH:
    LogFatal "Test_PR3_CitacDajeIstuZbirnuKaoClanstvo", Err.Number, Err.description
End Sub

' Zapis clanstva nosi opaque ID i fail-closed je, kao header i stavka.
Private Sub Test_PR3_PrazanIzvorIDNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3IZ")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3IZ-" & scenario, KLASA_I, 400#, 20)

    Dim preH As Long, preI As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)
    preI = Pr3BrojRedova(TBL_ZBIRNA_IZVORI)

    Dim rez As String, razlog As String
    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    ' header (1) + stavka (1) prolaze, clanstvo pada
    modDataAccess.NewEntityIDPadniTest True, 2
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3IZ-" & scenario), _
                          Pr3Izvor(otp, ""), razlog)
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode

    AssertEquals "", rez, "PR3 prazan ZBI: upis odbijen"
    AssertTrue InStr(1, razlog, "nije vratio ZbirnaIzvorID", vbTextCompare) > 0, _
               "PR3 prazan ZBI: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 prazan ZBI: header je rollback-ovan"
    AssertEquals CStr(preI), CStr(Pr3BrojRedova(TBL_ZBIRNA_IZVORI)), _
                 "PR3 prazan ZBI: clanstvo nije ostalo"
    AssertEquals "", Pr3OtpZbirnaID(otp), "PR3 prazan ZBI: otpremnica nije vezana"

    Exit Sub

EH:
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode
    LogFatal "Test_PR3_PrazanIzvorIDNeProlazi", Err.Number, Err.description
End Sub

' Zbirna je JEDAN transport JEDNOG vozaca.
'
' BrojZbirne je scoped po vozacu, pa bi zbirna sa otpremnicama dva vozaca bila i
' nepretraziva. Bez ove kapije je bilo moguce napraviti header sa VOZ-A, a sve
' izvore sa VOZ-B -- korupcija domena koju nista ne prijavljuje.
Private Sub Test_PR3_RazlicitVozacNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3VZ")

    Dim otp As String
    otp = Pr3OtpremnicaVozac(TEST_PREFIX & "-OTP-PR3VZ-" & scenario, KLASA_I, _
                             400#, 20, TEST_VOZ_ID_B)

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3VZ-" & scenario), _
                          Pr3Izvor(otp, ""), razlog)

    AssertEquals "", rez, "PR3 vozac: upis odbijen"
    AssertTrue InStr(1, razlog, "VozacID", vbTextCompare) > 0, _
               "PR3 vozac: kapija imenuje polje (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 vozac: header nije ostao"
    AssertEquals "", Pr3OtpZbirnaID(otp), "PR3 vozac: otpremnica nije vezana"

    Exit Sub

EH:
    LogFatal "Test_PR3_RazlicitVozacNeProlazi", Err.Number, Err.description
End Sub

' Nov model ne koristi legacy konvenciju "prazno = IZDATO".
'
' Ovaj writer pravi i ODMAH finalizuje dokument, pa to i upisuje. Kad se pojavi
' draft-first tok, razlika izmedju "niko nije upisao" i "izdato" vise ne sme da
' bude pretpostavka citaoca.
Private Sub Test_PR3_HeaderJeEksplicitnoIzdat()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3IZD")

    Dim zbrID As String
    zbrID = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3IZD-" & scenario), _
        Pr3Izvor(Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3IZD-" & scenario, _
                               KLASA_I, 400#, 20), ""))

    AssertTrue Len(zbrID) > 0, "PR3 izdato: dokument napravljen"
    AssertEquals IZDATO_IZDATO, ZbrPolje(zbrID, COL_TRACE_IZDATO_STATUS), _
                 "PR3 izdato: IzdatoStatus je upisan eksplicitno"

    Exit Sub

EH:
    LogFatal "Test_PR3_HeaderJeEksplicitnoIzdat", Err.Number, Err.description
End Sub

' Dva zapisa clanstva za istu otpremnicu -- cak i kad pokazuju na ISTU zbirnu.
'
' Pravilo je "tacno 0 ili 1 aktivan zapis", bez izuzetka. Dupli red iste veze ne
' menja kojoj zbirnoj otpremnica pripada, ali bi ga obican join sabrao DVAPUT --
' pa bi zbirna dobila dvostruku kolicinu iz jedne otpremnice.
Private Sub Test_PR3_DupliIstiZapisClanstvaJeGreska()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3DI")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3DI-" & scenario, KLASA_I, 400#, 20)

    Dim zbr As String
    zbr = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3DI-" & scenario), _
                          Pr3Izvor(otp, ""))
    AssertTrue Len(zbr) > 0, "PR3 dupli par: zbirna napravljena"

    ' Isti par (ZbirnaID, OtpremnicaID) jos jednom.
    Pr3DodajClanstvo zbr, otp
    AssertEquals "2", CStr(Pr3BrojClanstavaZa(otp)), _
                 "PR3 dupli par: dva zapisa iste veze"

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3DIX-" & scenario), _
        Pr3Izvor(Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3DIX-" & scenario, _
                               KLASA_I, 100#, 5), ""), razlog)

    ' Ciscenje PRE tvrdnji -- korumpiran kanon obara svaki sledeci test.
    Pr3UkloniClanstvo zbr, otp
    AssertEquals "1", CStr(Pr3BrojClanstavaZa(otp)), _
                 "PR3 dupli par: kanon vracen u konzistentno stanje"

    AssertEquals "", rez, "PR3 dupli par: upis odbijen"
    AssertTrue InStr(1, razlog, "dva aktivna zapisa", vbTextCompare) > 0, _
               "PR3 dupli par: kapija imenuje dupli zapis (bilo: " & razlog & ")"

    Exit Sub

EH:
    On Error Resume Next
    Pr3UkloniClanstvo zbr, otp
    On Error GoTo 0
    LogFatal "Test_PR3_DupliIstiZapisClanstvaJeGreska", Err.Number, Err.description
End Sub

' Ista otpremnica u DVE aktivne zbirne je korupcija kanona, ne rubni slucaj.
'
' Loader je ranije radio prosto mapa(otpID) = zbrID, pa bi drugi red tiho
' pregazio prvi. Upis bi i tada bio odbijen -- ali iz pogresnog razloga i sa
' pogresnom porukom, a stvarni problem (vec postoje dva clanstva) ostao bi
' neprijavljen. Kapija koja nelegalno stanje normalizuje u legalno radi protiv
' sebe.
Private Sub Test_PR3_DvaAktivnaClanstvaSuGreska()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3D2")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3D2-" & scenario, KLASA_I, 400#, 20)

    Dim prva As String
    prva = CreateZbirnaIzIzvora_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3D2A-" & scenario), _
                           Pr3Izvor(otp, ""))
    AssertTrue Len(prva) > 0, "PR3 dva clanstva: prva zbirna napravljena"

    ' Druga AKTIVNA zbirna, pa joj se rucno doda clanstvo iste otpremnice.
    Dim druga As String
    druga = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3D2B-" & scenario), _
        Pr3Izvor(Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3D2X-" & scenario, _
                               KLASA_I, 100#, 5), ""))
    AssertTrue Len(druga) > 0, "PR3 dva clanstva: druga zbirna napravljena"

    Pr3DodajClanstvo druga, otp
    AssertEquals "2", CStr(Pr3BrojClanstavaZa(otp)), _
                 "PR3 dva clanstva: kanon je sada nekonzistentan"

    ' Bilo koji sledeci upis mora da stane i da IMENUJE obe zbirne.
    '
    ' "Bilo koji" je namerno: korumpiran kanon blokira SVAKI upis zbirne, ne samo
    ' onaj koji dira sporni ID. To je tacan fail-closed ishod -- integritet
    ' clanstva nije pitanje jednog dokumenta.
    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    Dim rez As String, razlog As String
    rez = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3D2C-" & scenario), _
        Pr3Izvor(Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3D2Y-" & scenario, _
                               KLASA_I, 100#, 5), ""), razlog)

    ' CISCENJE ODMAH: korumpiran kanon ostaje u fixture-u i obara SVAKI sledeci
    ' test koji pravi zbirnu -- pad bez svoje krivice, i to devet puta zaredom.
    ' Ide PRE tvrdnji, da ga ne preskoci ni pad tvrdnje.
    Pr3UkloniClanstvo druga, otp
    AssertEquals "1", CStr(Pr3BrojClanstavaZa(otp)), _
                 "PR3 dva clanstva: kanon je vracen u konzistentno stanje"

    AssertEquals "", rez, "PR3 dva clanstva: upis odbijen"
    AssertTrue InStr(1, razlog, "nekonzistentno", vbTextCompare) > 0, _
               "PR3 dva clanstva: kapija imenuje nekonzistentnost (bilo: " & razlog & ")"
    AssertTrue InStr(1, razlog, prva, vbTextCompare) > 0 And _
               InStr(1, razlog, druga, vbTextCompare) > 0, _
               "PR3 dva clanstva: poruka imenuje OBE zbirne (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 dva clanstva: header nije ostao"

    Exit Sub

EH:
    ' I na putu greske -- inace ostatak suite-a pada bez svoje krivice.
    On Error Resume Next
    Pr3UkloniClanstvo druga, otp
    On Error GoTo 0
    LogFatal "Test_PR3_DvaAktivnaClanstvaSuGreska", Err.Number, Err.description
End Sub

' Kontrola "ocekivano vs izvedeno" se ne moze iskljuciti.
'
' Ranije je bio jedan ulaz sa Optional ocekivano, pa je Nothing (ili prazna
' kolekcija) tiho preskakao proveru -- bez ikakvog traga na pozivu. Sada rucni
' ulaz to odbija, a automatski tok ima svoj ulaz koji NAMERU kaze naglas.
Private Sub Test_PR3_RucniUnosTraziOcekivano()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PR3OB")

    Dim otp As String
    otp = Pr3Otpremnica(TEST_PREFIX & "-OTP-PR3OB-" & scenario, KLASA_I, 400#, 20)

    Dim preH As Long
    preH = Pr3BrojRedova(TBL_ZBIRNA)

    ' 1) bez ocekivanog
    Dim rez As String, razlog As String
    rez = CreateZbirna_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3OB1-" & scenario), _
                          Pr3Izvor(otp, ""), Nothing, razlog)

    AssertEquals "", rez, "PR3 obavezno ocekivano: Nothing je odbijen"
    AssertTrue InStr(1, razlog, "obavezne za rucni unos", vbTextCompare) > 0, _
               "PR3 obavezno ocekivano: kapija imenuje razlog (bilo: " & razlog & ")"

    ' 2) prazna kolekcija je isto sto i nijedna
    rez = CreateZbirna_TX(Pr3Header(TEST_PREFIX & "-ZBR-PR3OB2-" & scenario), _
                          Pr3Izvor(otp, ""), New Collection, razlog)

    AssertEquals "", rez, "PR3 obavezno ocekivano: prazna kolekcija je odbijena"
    AssertTrue InStr(1, razlog, "prazne", vbTextCompare) > 0, _
               "PR3 obavezno ocekivano: prazno se imenuje (bilo: " & razlog & ")"

    AssertEquals CStr(preH), CStr(Pr3BrojRedova(TBL_ZBIRNA)), _
                 "PR3 obavezno ocekivano: nijedan header nije ostao"

    ' 3) automatski tok ISTU zbirnu pravi bez ocekivanog -- namera je izricita
    Dim auto As String
    auto = CreateZbirnaIzIzvora_TX( _
        Pr3Header(TEST_PREFIX & "-ZBR-PR3OB3-" & scenario), Pr3Izvor(otp, ""))
    AssertTrue Len(auto) > 0, _
               "PR3 obavezno ocekivano: automatski ulaz prolazi bez njega"

    Exit Sub

EH:
    LogFatal "Test_PR3_RucniUnosTraziOcekivano", Err.Number, Err.description
End Sub

' ============================================================
' OTKUP skela -- header + stavke
' ============================================================
'
' Sta se meri:  da CreateOtkup_TX pravi JEDAN header i N stavki, da KulturaID
'               PRIMA a ne fabrikuje, da parcela mora biti kooperantova, i da
'               bruto/neto i cena ostaju zamrznute cinjenice.
' Sta se NE meri: ponasanje citalaca -- ona su predmet Otkup cutover-a, koji je
'               ambalazu i novac doveo pod isti pisac i obrisao starog.

Private Sub Test_OTK_HeaderIStavke()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKHS")

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-" & scenario), _
                           OtkStavke(400#, 50#, 20, 600#, 40#, 30))

    AssertTrue Len(otkID) > 0, "OTK: CreateOtkup_TX vraca ID"

    Dim headeri As Collection
    Set headeri = FindRows(TBL_OTKUP, COL_OTK_ID, otkID)
    AssertTrue Not headeri Is Nothing, "OTK: header pronadjen"
    AssertEquals "1", CStr(headeri.count), "OTK: tacno jedan header red"

    AssertEquals TEST_KOOP_ID, OtkPolje(otkID, COL_OTK_KOOPERANT), "OTK: KooperantID"
    AssertEquals TEST_ST_ID, OtkPolje(otkID, COL_OTK_STANICA), "OTK: StanicaID"
    AssertEquals TEST_KULTURA_ID, OtkPolje(otkID, COL_OTK_KULTURA), "OTK: KulturaID"
    AssertEquals TEST_VRSTA, OtkPolje(otkID, COL_OTK_VRSTA), "OTK: VrstaVoca"

    AssertEquals "2", CStr(OtkBrojStavki(otkID)), "OTK: dve stavke"
    AssertEquals "1", OtkStavkaPolje(otkID, KLASA_I, COL_OKS_RB), "OTK: I ima RB 1"
    AssertEquals "2", OtkStavkaPolje(otkID, KLASA_II, COL_OKS_RB), "OTK: II ima RB 2"

    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_I, COL_OKS_KOLICINA) - 400#) < 0.001, _
               "OTK: Klasa I kolicina 400"
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_I, COL_OKS_CENA) - 50#) < 0.001, _
               "OTK: Klasa I cena 50"
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_II, COL_OKS_KOLICINA) - 600#) < 0.001, _
               "OTK: Klasa II kolicina 600"
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_II, COL_OKS_CENA) - 40#) < 0.001, _
               "OTK: Klasa II cena 40 (razlicita od Klase I)"

    AssertEquals otkID, OtkStavkaPolje(otkID, KLASA_I, COL_OKS_OTKUP_ID), _
                 "OTK: stavka pokazuje na OtkupID"

    Exit Sub

EH:
    LogFatal "Test_OTK_HeaderIStavke", Err.Number, Err.description
End Sub

' Header ne nosi nista sto je stavka, ni polja koja u ciljnom modelu ne postoje.
'
' Od S1d kolone stavke na zaglavlju ne postoje (kanon), pa se tvrdi njihovo odsustvo.
Private Sub Test_OTK_HeaderNeNosiLinePolja()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKHP")

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-HP-" & scenario), _
                           OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK header: dokument napravljen"

    ' S1d: kolone stavke i polja kojih u ciljnom modelu nema su OBRISANE iz
    ' kanona (schema.json) i iz sveske -- tvrdi se da ih nema, ne da su prazne.
    Dim odlazi As Variant, k As Long
    odlazi = Array("Kolicina", "Cena", "Klasa", "KolAmbalaze", "BrutoKg", _
                   "Novac", "PrimalacNovca", "VremeUnosa")
    For k = LBound(odlazi) To UBound(odlazi)
        AssertEquals "0", CStr(GetColumnIndex(TBL_OTKUP, CStr(odlazi(k)))), _
                     "OTK header: kolona " & CStr(odlazi(k)) & " ne postoji na zaglavlju"
    Next k

    AssertEquals "", OtkPolje(otkID, COL_OTK_VOZAC), _
                 "OTK header: VozacID prazan (vozac pripada otpremnici)"

    If GetColumnIndex(TBL_OTKUP, COL_GENERACIJA_ID) > 0 Then
        AssertEquals "", OtkPolje(otkID, COL_GENERACIJA_ID), _
                     "OTK header: GeneracijaID se ne pise"
    End If

    AssertEquals IZDATO_IZDATO, OtkPolje(otkID, COL_TRACE_IZDATO_STATUS), _
                 "OTK header: IzdatoStatus je upisan eksplicitno"

    Exit Sub

EH:
    LogFatal "Test_OTK_HeaderNeNosiLinePolja", Err.Number, Err.description
End Sub

' KulturaID se PRIMA, ne fabrikuje.
'
' Zatecen kod na dva mesta sklopi "vrsta-sorta" string kad lookup ne uspe
' (modOtkup.bas:556, modMasterSync.bas:1959) -- to izgleda kao FK a ne pokazuje
' ni na sta. Nov writer takav ID odbija jer takvog reda u tblKulture nema.
Private Sub Test_OTK_KulturaSeNeFabrikuje()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKKF")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-KF-" & scenario)
    h("KulturaID") = TEST_VRSTA & "-" & TEST_SORTA     ' tacno oblik koji stari kod pravi

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK kultura: fabrikovan ID odbijen"
    AssertTrue InStr(1, razlog, "KulturaID ne postoji", vbTextCompare) > 0, _
               "OTK kultura: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK kultura: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_OTK_KulturaSeNeFabrikuje", Err.Number, Err.description
End Sub

' Postojeci KulturaID nije dovoljan -- snapshot vrsta/sorta mora da mu odgovara.
'
' Bez ove provere bi dokument nosio jednu vrstu u tekstu a drugu preko FK-a, pa
' bi izvestaj po kulturi i izvestaj po vrsti davali razlicite brojeve.
Private Sub Test_OTK_KulturaSeMoraSlagatiSaVrstom()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKKS")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-KS-" & scenario)
    h("VrstaVoca") = TEST_VRSTA & " DRUGA"

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK kultura/vrsta: upis odbijen"
    AssertTrue InStr(1, razlog, "ne slazu sa kulturom", vbTextCompare) > 0, _
               "OTK kultura/vrsta: kapija imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_KulturaSeMoraSlagatiSaVrstom", Err.Number, Err.description
End Sub

' Tudja parcela ne prolazi kanonski writer (HARD, S4.1f).
Private Sub Test_OTK_ParcelaPripadaKooperantu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPP")

    ' Ista parcela, ali otkup za DRUGOG kooperanta -- koji POSTOJI.
    ' Sa izmisljenim ID-em bi otkup pao na FK proveri, pa bi test tvrdio
    ' vlasnistvo a merio postojanje.
    AssertTrue RowExists(TBL_KOOPERANTI, "KooperantID", TEST_KOOP2_ID), _
               "OTK parcela: drugi kooperant postoji"

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-PP-" & scenario)
    h("ParcelaID") = TEST_PAR_ID
    h("KooperantID") = TEST_KOOP2_ID

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK parcela: tudja parcela odbijena"
    AssertTrue InStr(1, razlog, "pripada kooperantu", vbTextCompare) > 0, _
               "OTK parcela: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK parcela: header nije ostao"

    ' Kontrola: ISTA parcela sa SVOJIM kooperantom prolazi -- inace bi kapija
    ' koja uvek odbija izgledala isto kao kapija koja radi.
    Dim h2 As Object
    Set h2 = OtkHeader(TEST_PREFIX & "-OTK-PP2-" & scenario)
    h2("ParcelaID") = TEST_PAR_ID

    AssertTrue Len(CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0))) > 0, _
               "OTK parcela: svoja parcela prolazi"

    Exit Sub

EH:
    LogFatal "Test_OTK_ParcelaPripadaKooperantu", Err.Number, Err.description
End Sub

' Bruto unos cuva OBA broja; neto unos ne izmislja bruto.
'
' Prazan BrutoKg je PODATAK ("unet je neto"), ne nula. Tezina gajbice se koristi
' samo u trenutku nastanka -- izdat dokument se nikad ne rekalkulise (S4.1d).
Private Sub Test_OTK_BrutoINetoSuZamrznuti()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKBN")

    ' bruto 1100, ambalaza 100 -> neto 1000 (tara 1 kg/gajbi u trenutku unosa)
    Dim sBruto As Collection
    Set sBruto = New Collection
    sBruto.Add OtkStavka(KLASA_I, 1000#, 50#, 100, 1100#)

    Dim brutoID As String
    brutoID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-BN1-" & scenario), sBruto)
    AssertTrue Len(brutoID) > 0, "OTK bruto: dokument napravljen"

    AssertTrue Abs(OtkStavkaBrojP(brutoID, KLASA_I, COL_OKS_BRUTO) - 1100#) < 0.001, _
               "OTK bruto: BrutoKg je tacno ono sto je uneto"
    AssertTrue Abs(OtkStavkaBrojP(brutoID, KLASA_I, COL_OKS_KOLICINA) - 1000#) < 0.001, _
               "OTK bruto: Kolicina je neto"

    ' neto unos -> BrutoKg ostaje PRAZAN
    Dim netoID As String
    netoID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-BN2-" & scenario), _
                            OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertEquals "", OtkStavkaPolje(netoID, KLASA_I, COL_OKS_BRUTO), _
                 "OTK neto: BrutoKg ostaje prazan, ne nula"

    ' bruto manji od neta = zamenjene vrednosti, ne rubni slucaj
    Dim sLos As Collection
    Set sLos = New Collection
    sLos.Add OtkStavka(KLASA_I, 1000#, 50#, 100, 900#)

    Dim razlog As String
    AssertEquals "", CreateOtkup_TX( _
        OtkHeader(TEST_PREFIX & "-OTK-BN3-" & scenario), sLos, razlog), _
        "OTK bruto: bruto manji od neta je odbijen"
    AssertTrue InStr(1, razlog, "manji od neto", vbTextCompare) > 0, _
               "OTK bruto: kapija imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_BrutoINetoSuZamrznuti", Err.Number, Err.description
End Sub

' Cenovnik je PREDLOG. Writer trazi samo Cena > 0 -- override je legitiman i
' sacuvana cena je istorijska cinjenica dokumenta.
Private Sub Test_OTK_CenaJeStvarnoPrimenjena()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKCE")

    ' cena koja sigurno nije iz cenovnika
    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-CE-" & scenario), _
                           OtkStavke(400#, 137.5, 20, 0#, 0#, 0))

    AssertTrue Len(otkID) > 0, "OTK cena: override prolazi"
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_I, COL_OKS_CENA) - 137.5) < 0.001, _
               "OTK cena: sacuvana je bas uneta cena"

    Dim razlog As String
    AssertEquals "", CreateOtkup_TX( _
        OtkHeader(TEST_PREFIX & "-OTK-CE0-" & scenario), _
        OtkStavke(400#, 0#, 20, 0#, 0#, 0), razlog), _
        "OTK cena: nula je odbijena"
    AssertTrue InStr(1, razlog, "Cena mora biti veca od nule", vbTextCompare) > 0, _
               "OTK cena: kapija imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_CenaJeStvarnoPrimenjena", Err.Number, Err.description
End Sub

' Dve stavke iste klase su bas bug koji header+stavke uklanja.
Private Sub Test_OTK_DuplaKlasaPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKDK")

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add OtkStavka(KLASA_I, 400#, 50#, 20, 0#)
    stavke.Add OtkStavka(KLASA_I, 600#, 40#, 30, 0#)

    Dim preH As Long, preS As Long
    preH = OtkBrojRedova(TBL_OTKUP)
    preS = OtkBrojRedova(TBL_OTKUP_STAVKE)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-DK-" & scenario), stavke, razlog)

    AssertEquals "", rez, "OTK dupla klasa: upis odbijen"
    AssertTrue InStr(1, razlog, "Dve stavke iste klase", vbTextCompare) > 0, _
               "OTK dupla klasa: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK dupla klasa: header nije ostao"
    AssertEquals CStr(preS), CStr(OtkBrojRedova(TBL_OTKUP_STAVKE)), _
                 "OTK dupla klasa: stavka nije ostala"

    Exit Sub

EH:
    LogFatal "Test_OTK_DuplaKlasaPada", Err.Number, Err.description
End Sub

' Neispravna DRUGA stavka ne sme da ostavi header i prvu stavku.
Private Sub Test_OTK_LosaDrugaStavkaRollback()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKLS")

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add OtkStavka(KLASA_I, 400#, 50#, 20, 0#)
    stavke.Add OtkStavka(KLASA_II, 0#, 40#, 30, 0#)      ' kolicina 0

    Dim preH As Long, preS As Long
    preH = OtkBrojRedova(TBL_OTKUP)
    preS = OtkBrojRedova(TBL_OTKUP_STAVKE)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-LS-" & scenario), stavke, razlog)

    AssertEquals "", rez, "OTK losa stavka: upis odbijen"
    AssertTrue InStr(1, razlog, "Kolicina mora biti veca od nule", vbTextCompare) > 0, _
               "OTK losa stavka: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK losa stavka: header nije ostao"
    AssertEquals CStr(preS), CStr(OtkBrojRedova(TBL_OTKUP_STAVKE)), _
                 "OTK losa stavka: prva stavka nije ostala"

    Exit Sub

EH:
    LogFatal "Test_OTK_LosaDrugaStavkaRollback", Err.Number, Err.description
End Sub

' Red bez identiteta je gori od pada. Seam broji pozive, pa se header i stavka
' mere odvojeno.
Private Sub Test_OTK_PrazanIDFailClosed()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKID")

    Dim preH As Long, preS As Long
    preH = OtkBrojRedova(TBL_OTKUP)
    preS = OtkBrojRedova(TBL_OTKUP_STAVKE)

    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    ' 1) header ID pada odmah
    Dim rez As String, razlog As String
    modDataAccess.NewEntityIDPadniTest True
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-ID1-" & scenario), _
                         OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)
    modDataAccess.NewEntityIDPadniTest False

    AssertEquals "", rez, "OTK prazan ID: header ID odbijen"
    AssertTrue InStr(1, razlog, "nije vratio OtkupID", vbTextCompare) > 0, _
               "OTK prazan ID: kapija imenuje razlog (bilo: " & razlog & ")"

    ' 2) header prodje, stavka padne -> header mora biti rollback-ovan
    modDataAccess.NewEntityIDPadniTest True, 1
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-ID2-" & scenario), _
                         OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode

    AssertEquals "", rez, "OTK prazan OKS: upis odbijen"
    AssertTrue InStr(1, razlog, "nije vratio OtkupStavkaID", vbTextCompare) > 0, _
               "OTK prazan OKS: kapija imenuje razlog (bilo: " & razlog & ")"

    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK prazan ID: nijedan header nije ostao"
    AssertEquals CStr(preS), CStr(OtkBrojRedova(TBL_OTKUP_STAVKE)), _
                 "OTK prazan ID: nijedna stavka nije ostala"

    Exit Sub

EH:
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode
    LogFatal "Test_OTK_PrazanIDFailClosed", Err.Number, Err.description
End Sub

' Izdata ambalaza je dokument-level cinjenica sa otkupnog lista, ne stavka.
Private Sub Test_OTK_KolAmbIzdataJeHeader()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKAI")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-AI-" & scenario)
    h("KolAmbIzdata") = 45

    Dim otkID As String
    otkID = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 600#, 40#, 30))

    AssertTrue Len(otkID) > 0, "OTK izdata: dokument napravljen"
    AssertEquals "45", OtkPolje(otkID, COL_OTK_KOL_AMB_IZDATA), _
                 "OTK izdata: vrednost je na HEADERU"
    AssertEquals "2", CStr(OtkBrojStavki(otkID)), _
                 "OTK izdata: dve stavke, a izdata ambalaza se ne deli po klasi"

    Exit Sub

EH:
    LogFatal "Test_OTK_KolAmbIzdataJeHeader", Err.Number, Err.description
End Sub

' Nepoznat kljuc u headeru je GRESKA -- ukljucujuci polja koja su u STAROM
' modelu bila na headeru. Pozivalac koji salje VozacID ili Kolicina radi po
' starom modelu i mora to da cuje.
Private Sub Test_OTK_NepoznatKljucUHeaderuPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKNK")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-NK-" & scenario)
    h.Add "VozacID", TEST_VOZ_ID

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK nepoznat kljuc: upis odbijen"
    AssertTrue InStr(1, razlog, "nepoznat kljuc", vbTextCompare) > 0 And _
               InStr(1, razlog, "VozacID", vbTextCompare) > 0, _
               "OTK nepoznat kljuc: kapija imenuje kljuc (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK nepoznat kljuc: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_OTK_NepoznatKljucUHeaderuPada", Err.Number, Err.description
End Sub

' Otkup bez ijedne stavke nije dokument.
Private Sub Test_OTK_BezStavkiNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKBS")

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-BS-" & scenario), _
                         New Collection, razlog)

    AssertEquals "", rez, "OTK bez stavki: upis odbijen"
    AssertTrue InStr(1, razlog, "bar jednu stavku", vbTextCompare) > 0, _
               "OTK bez stavki: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK bez stavki: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_OTK_BezStavkiNeProlazi", Err.Number, Err.description
End Sub

' Ambalaza je BROJ KOMADA. Isto pravilo kao na zbirnoj -- odbija se, ne
' zaokruzuje: "20.5 gajbica" je kvar u izvoru, a tiha ispravka ga sakriva.
Private Sub Test_OTK_AmbalazaMoraBitiCeoBroj()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKAM")

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add OtkStavka(KLASA_I, 400#, 50#, 1.5, 0#)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-AM-" & scenario), stavke, razlog)

    AssertEquals "", rez, "OTK ambalaza: upis odbijen"
    AssertTrue InStr(1, razlog, "mora biti ceo broj", vbTextCompare) > 0, _
               "OTK ambalaza: kapija imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_AmbalazaMoraBitiCeoBroj", Err.Number, Err.description
End Sub

' --- OTK pomocne -------------------------------------------------------------

' Stavka ima zatvoren spisak kljuceva -- kao header.
'
' Tipfeler bas u OPCIONOM polju je jedini koji nema svoj glas: "BruttoKg" se ne
' procita, BrutoKg ostane prazan, i bruto unos tiho postane neto. Zato test
' cilja bas njega, a ne neko obavezno polje koje bi palo i bez whitelist-a.
Private Sub Test_OTK_NepoznatKljucUStavciPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKNS")

    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", KLASA_I
    s.Add "Kolicina", 1000#
    s.Add "Cena", 50#
    s.Add "KolAmbalaze", 100#
    s.Add "BruttoKg", 1100#                  ' tipfeler u opcionom polju

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add s

    Dim preH As Long, preS As Long
    preH = OtkBrojRedova(TBL_OTKUP)
    preS = OtkBrojRedova(TBL_OTKUP_STAVKE)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-NS-" & scenario), stavke, razlog)

    AssertEquals "", rez, "OTK stavka kljuc: upis odbijen"
    AssertTrue InStr(1, razlog, "nepoznat kljuc", vbTextCompare) > 0 And _
               InStr(1, razlog, "BruttoKg", vbTextCompare) > 0, _
               "OTK stavka kljuc: kapija imenuje kljuc (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK stavka kljuc: header nije ostao"
    AssertEquals CStr(preS), CStr(OtkBrojRedova(TBL_OTKUP_STAVKE)), _
                 "OTK stavka kljuc: stavka nije ostala"

    Exit Sub

EH:
    LogFatal "Test_OTK_NepoznatKljucUStavciPada", Err.Number, Err.description
End Sub

' KooperantID je FK, ne string. Neprazan tekst nije dokaz da kooperant postoji.
Private Sub Test_OTK_KooperantMoraPostojati()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKKP")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-KP-" & scenario)
    h("KooperantID") = "KOOP-NE-POSTOJI-" & scenario

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK kooperant: nepostojeci FK odbijen"
    AssertTrue InStr(1, razlog, "KooperantID ne postoji", vbTextCompare) > 0, _
               "OTK kooperant: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK kooperant: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_OTK_KooperantMoraPostojati", Err.Number, Err.description
End Sub

' StanicaID je FK ka tblStanice, i NE izvodi se iz kooperanta.
'
' Kontrola na kraju je poslovna, ne kozmeticka: otkup na stanici koja NIJE
' maticna stanica kooperanta mora da prodje. Kod desktopa stanica dolazi iz
' zakljucane sesije (modStanicaLock), pa kooperant sme da preda robu bilo gde.
Private Sub Test_OTK_StanicaMoraPostojati()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKSP")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-SP-" & scenario)
    h("StanicaID") = "ST-NE-POSTOJI-" & scenario

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK stanica: nepostojeci FK odbijen"
    AssertTrue InStr(1, razlog, "StanicaID ne postoji", vbTextCompare) > 0, _
               "OTK stanica: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK stanica: header nije ostao"

    ' Stanica koja NIJE maticna stanica kooperanta i dalje prolazi.
    AssertTrue StrComp(TEST_HLAD_ST_ID, _
                       Trim$(CStr(nz(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, _
                                                 TEST_KOOP_ID, COL_KOOP_STANICA), ""))), _
                       vbTextCompare) <> 0, _
               "OTK stanica: druga stanica zaista nije maticna (inace kontrola ne meri nista)"

    Dim h2 As Object
    Set h2 = OtkHeader(TEST_PREFIX & "-OTK-SP2-" & scenario)
    h2("StanicaID") = TEST_HLAD_ST_ID

    AssertTrue Len(CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0))) > 0, _
               "OTK stanica: otkup na nematicnoj stanici prolazi"

    Exit Sub

EH:
    LogFatal "Test_OTK_StanicaMoraPostojati", Err.Number, Err.description
End Sub

' RedniBroj nosi dokument, ne redosled poziva.
'
' Adapter sme da sklopi stavke bilo kojim redom; ista poslovna cinjenica mora
' dati isti dokument. Test salje II pa I -- suprotno od kanonskog reda, i
' suprotno od onoga sto OtkStavke() pravi, pa zelena boja ovde nije slucajna.
Private Sub Test_OTK_RedosledKlasaJeKanonski()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKRK")

    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add OtkStavka(KLASA_II, 600#, 40#, 30#, 0#)     ' II je PRVA u ulazu
    stavke.Add OtkStavka(KLASA_I, 400#, 50#, 20#, 0#)

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-RK-" & scenario), stavke)

    AssertTrue Len(otkID) > 0, "OTK red klasa: upis prosao"
    AssertEquals "2", CStr(OtkBrojStavki(otkID)), "OTK red klasa: dve stavke"

    AssertEquals "1", OtkStavkaPolje(otkID, KLASA_I, COL_OKS_RB), _
                 "OTK red klasa: I ima RB 1 iako je poslata druga"
    AssertEquals "2", OtkStavkaPolje(otkID, KLASA_II, COL_OKS_RB), _
                 "OTK red klasa: II ima RB 2 iako je poslata prva"

    ' Preslozen RedniBroj ne sme da preslozi i sadrzaj -- klasa i njeni brojevi
    ' moraju ostati zajedno.
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_I, COL_OKS_KOLICINA) - 400#) < 0.001, _
               "OTK red klasa: I zadrzala svoju kolicinu"
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_II, COL_OKS_CENA) - 40#) < 0.001, _
               "OTK red klasa: II zadrzala svoju cenu"

    Exit Sub

EH:
    LogFatal "Test_OTK_RedosledKlasaJeKanonski", Err.Number, Err.description
End Sub

' Otkup samo druge klase je stvaran tok koji zatecen ekran podrzava
' (modOtkupUnos: Klasa I sme da ostane prazna kad je ukljucena Klasa II).
Private Sub Test_OTK_SamoKlasaII()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKS2")

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-S2-" & scenario), _
                           OtkStavke(0#, 0#, 0, 600#, 40#, 30))

    AssertTrue Len(otkID) > 0, "OTK samo II: upis prosao"
    AssertEquals "1", CStr(OtkBrojStavki(otkID)), "OTK samo II: tacno jedna stavka"
    AssertEquals "1", OtkStavkaPolje(otkID, KLASA_II, COL_OKS_RB), _
                 "OTK samo II: II ima RB 1 kad je sama"
    AssertEquals "", OtkStavkaPolje(otkID, KLASA_I, COL_OKS_RB), _
                 "OTK samo II: stavke klase I nema"
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_II, COL_OKS_KOLICINA) - 600#) < 0.001, _
               "OTK samo II: kolicina na stavci"

    Exit Sub

EH:
    LogFatal "Test_OTK_SamoKlasaII", Err.Number, Err.description
End Sub

' Prazna sorta je legitimna TACNO kad je i sama kultura bez sorte.
'
' Writer ne donosi tu odluku: pravilo je "snapshot mora da odgovara kulturi", pa
' prazno prolazi samo tamo gde je i master prazan. Kljuc ipak mora da postoji --
' inace bi tipfeler u imenu polja prosao kao "kultura nema sortu".
Private Sub Test_OTK_SortaPraznaSamoUzKulturuBezSorte()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPS")

    ' (a) kultura IMA sortu -- prazna sorta na dokumentu je neslaganje
    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-PS-" & scenario)
    h("SortaVoca") = ""

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK prazna sorta: uz kulturu SA sortom odbijena"
    AssertTrue InStr(1, razlog, "ne slazu sa kulturom", vbTextCompare) > 0, _
               "OTK prazna sorta: kapija imenuje razlog (bilo: " & razlog & ")"

    ' (b) kultura NEMA sortu -- prazna sorta je tacan podatak
    Dim h2 As Object
    Set h2 = OtkHeader(TEST_PREFIX & "-OTK-PS2-" & scenario)
    h2("KulturaID") = TEST_KUL_BEZ_SORTE_ID
    h2("VrstaVoca") = TEST_VRSTA_BEZ_SORTE
    h2("SortaVoca") = ""

    Dim otkID As String
    otkID = CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertTrue Len(otkID) > 0, _
               "OTK prazna sorta: uz kulturu BEZ sorte prolazi (bilo: " & razlog & ")"
    AssertEquals "", OtkPolje(otkID, COL_OTK_SORTA), _
                 "OTK prazna sorta: sorta ostaje prazna, ne izmisljena"

    ' (c) kljuc mora da postoji i onda kad sme da bude prazan
    Dim h3 As Object
    Set h3 = OtkHeader(TEST_PREFIX & "-OTK-PS3-" & scenario)
    h3.Remove "SortaVoca"

    rez = CreateOtkup_TX(h3, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK prazna sorta: kljuc koji fali je greska"
    AssertTrue InStr(1, razlog, "nema obavezan kljuc", vbTextCompare) > 0 And _
               InStr(1, razlog, "SortaVoca", vbTextCompare) > 0, _
               "OTK prazna sorta: kapija imenuje kljuc (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_SortaPraznaSamoUzKulturuBezSorte", Err.Number, Err.description
End Sub

' Tip ambalaze vezuje SVAKA ambalaza -- i primljena na stavkama i izdata na
' headeru. Bez ambalaze je prazan tip tacan podatak, ne propust.
'
' Slucaj (c) je onaj koji je zatecen ekran vec pokrivao (modOtkupUnos:158
' gleda i kolAmbIzdata), a nov writer umalo nije: izdata ambalaza bez tipa je
' gajba koja je otisla kooperantu a ne zna se koja.
Private Sub Test_OTK_TipAmbalazeVezujeSvakaAmbalaza()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKTA")

    ' (a) nema nikakve ambalaze -- prazan tip prolazi
    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-TA-" & scenario)
    h("TipAmbalaze") = ""

    Dim razlog As String
    Dim otkID As String
    otkID = CreateOtkup_TX(h, OtkStavke(400#, 50#, 0, 0#, 0#, 0), razlog)

    AssertTrue Len(otkID) > 0, _
               "OTK tip ambalaze: bez ambalaze prazan tip prolazi (bilo: " & razlog & ")"

    ' (b) primljena ambalaza na stavci
    Dim h2 As Object
    Set h2 = OtkHeader(TEST_PREFIX & "-OTK-TA2-" & scenario)
    h2("TipAmbalaze") = ""

    Dim rez As String
    rez = CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK tip ambalaze: primljena ambalaza bez tipa odbijena"
    AssertTrue InStr(1, razlog, "Tip ambalaze", vbTextCompare) > 0, _
               "OTK tip ambalaze: kapija imenuje razlog (bilo: " & razlog & ")"

    ' (c) IZDATA ambalaza, na stavkama je nema
    Dim h3 As Object
    Set h3 = OtkHeader(TEST_PREFIX & "-OTK-TA3-" & scenario)
    h3("TipAmbalaze") = ""
    h3.Add "KolAmbIzdata", 5#

    rez = CreateOtkup_TX(h3, OtkStavke(400#, 50#, 0, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK tip ambalaze: izdata ambalaza bez tipa odbijena"
    AssertTrue InStr(1, razlog, "Tip ambalaze", vbTextCompare) > 0, _
               "OTK tip ambalaze: izdata ambalaza imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_TipAmbalazeVezujeSvakaAmbalaza", Err.Number, Err.description
End Sub

Private Function OtkHeader(ByVal brDok As String) As Object
    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", NextTestDate()
    h.Add "KooperantID", TEST_KOOP_ID
    h.Add "StanicaID", TEST_ST_ID
    h.Add "KulturaID", TEST_KULTURA_ID
    h.Add "VrstaVoca", TEST_VRSTA
    h.Add "SortaVoca", TEST_SORTA
    h.Add "TipAmbalaze", TEST_TIP_AMB
    h.Add "BrojDokumenta", brDok
    Set OtkHeader = h
End Function

Private Function OtkStavka(ByVal klasa As String, ByVal kol As Double, _
                           ByVal cena As Double, ByVal amb As Double, _
                           ByVal bruto As Double) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", klasa
    s.Add "Kolicina", kol
    s.Add "Cena", cena
    s.Add "KolAmbalaze", amb
    If bruto > 0 Then s.Add "BrutoKg", bruto
    Set OtkStavka = s
End Function

' Klasa II se izostavlja kad je kolII = 0 -- dokument sme da ima samo jednu klasu.
Private Function OtkStavke(ByVal kolI As Double, ByVal cenaI As Double, _
                           ByVal ambI As Double, ByVal kolII As Double, _
                           ByVal cenaII As Double, ByVal ambII As Double) As Collection
    Dim c As Collection
    Set c = New Collection
    If kolI > 0 Then c.Add OtkStavka(KLASA_I, kolI, cenaI, ambI, 0#)
    If kolII > 0 Then c.Add OtkStavka(KLASA_II, kolII, cenaII, ambII, 0#)
    Set OtkStavke = c
End Function

Private Function OtkPolje(ByVal otkupID As String, ByVal columnName As String) As String
    OtkPolje = Trim$(CStr(nz(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkupID, columnName), "")))
End Function

Private Function OtkBrojRedova(ByVal tblName As String) As Long
    Dim d As Variant
    d = GetTableData(tblName)
    If Not IsArray(d) Then Exit Function
    OtkBrojRedova = UBound(d, 1)
End Function

Private Function OtkBrojStavki(ByVal otkupID As String) As Long
    Dim redovi As Collection
    Set redovi = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkupID)
    If redovi Is Nothing Then Exit Function
    OtkBrojStavki = redovi.count
End Function

Private Function OtkStavkaPolje(ByVal otkupID As String, ByVal klasa As String, _
                                ByVal columnName As String) As String
    Dim d As Variant
    d = GetTableData(TBL_OTKUP_STAVKE)
    If Not IsArray(d) Then Exit Function

    Dim cOtk As Long, cKlasa As Long, cTraz As Long
    cOtk = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, "OtkStavkaPolje")
    cKlasa = RequireColumnIndex(TBL_OTKUP_STAVKE, COL_OKS_KLASA, "OtkStavkaPolje")
    cTraz = RequireColumnIndex(TBL_OTKUP_STAVKE, columnName, "OtkStavkaPolje")

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(i, cOtk), "")), otkupID, vbTextCompare) = 0 Then
            If StrComp(Trim$(nz(d(i, cKlasa), "")), klasa, vbTextCompare) = 0 Then
                OtkStavkaPolje = Trim$(CStr(nz(d(i, cTraz), "")))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function OtkStavkaBrojP(ByVal otkupID As String, ByVal klasa As String, _
                                ByVal columnName As String) As Double
    Dim t As String
    t = OtkStavkaPolje(otkupID, klasa, columnName)
    If IsNumeric(t) Then OtkStavkaBrojP = CDbl(t)
End Function

' --- OTPREMNICA skela (PR5) --------------------------------------------------
'
' Stavke drafta su OCEKIVANJE (sta je operater prijavio), clanstvo daje POVEZANO
' (sta su otkupni listovi dokumentovali), izdavanje trazi jednakost. Izvori su
' otkupi po NOVOM modelu, pa ovi testovi usput mere i da se dva nova pisca slazu.

' Dvoklasni otkup -> otpremnica: JEDAN header, dve stavke, kanonski red klasa.
Private Sub Test_OTP_JedanBrojJedanHeader()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPJH")

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add OtpNoviOtkup(scenario, 400#, 600#)

    Dim razlog As String
    Dim otpID As String
    otpID = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-" & scenario), _
                                        izvori, razlog)

    AssertTrue Len(otpID) > 0, "OTP: upis prosao (bilo: " & razlog & ")"

    AssertEquals "1", CStr(FindRows(TBL_OTPREMNICA, COL_OTP_ID, otpID).count), _
                 "OTP: tacno jedan header red"
    AssertEquals "2", CStr(OtpBrojStavki(otpID)), "OTP: dve stavke"
    AssertEquals "1", OtpStavkaPolje(otpID, KLASA_I, COL_OPS_RB), "OTP: I ima RB 1"
    AssertEquals "2", OtpStavkaPolje(otpID, KLASA_II, COL_OPS_RB), "OTP: II ima RB 2"
    AssertEquals IZDATO_IZDATO, OtpPolje(otpID, COL_TRACE_IZDATO_STATUS), _
                 "OTP: dokument je IZDATO"
    AssertEquals "1", CStr(OtpBrojClanova(otpID)), "OTP: jedan clan"

    Exit Sub

EH:
    LogFatal "Test_OTP_JedanBrojJedanHeader", Err.Number, Err.description
End Sub

' PREDLOG CENE JE PO KLASI (S3a, odluka S14.8 t. 2).
'
' Do S3a je predlog stajao na zaglavlju, kao JEDAN broj. Otpremnica koja nosi i
' prvu i drugu klasu je time obe prefilovala istom cenom -- a druga klasa je
' jeftinija, pa je operater tu cenu ispravljao na svakom otkupnom bloku ili je,
' gore, ostavljao. Zaglavlje vise cenu ne prima: kljuc se ODBIJA, da pozivalac
' koji je i dalje salje to i sazna.
Private Sub Test_OTP_PredlogCeneJePoKlasi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPPC")

    Dim c As Collection
    Set c = New Collection
    c.Add OtpOcekStavkaSaCenom(KLASA_I, 400#, 20#, 250#)
    c.Add OtpOcekStavkaSaCenom(KLASA_II, 600#, 30#, 120#)

    Dim razlog As String
    Dim otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-PC-" & scenario), _
                                     c, razlog)

    AssertTrue Len(otpID) > 0, "OTP predlog cene: draft prosao (bilo: " & razlog & ")"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_I, COL_OPS_PREDLOG_CENA) - 250#) < 0.001, _
               "OTP predlog cene: Klasa I nosi svoju cenu"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_II, COL_OPS_PREDLOG_CENA) - 120#) < 0.001, _
               "OTP predlog cene: Klasa II nosi SVOJU cenu, ne cenu prve"
    AssertEquals "", OtpPolje(otpID, COL_OTP_CENA), _
                 "OTP predlog cene: zaglavlje vise ne nosi cenu"

    ' Stavka bez predloga je legitimna ("cena jos nije dogovorena") i ostaje
    ' PRAZNA -- nula bi u prefillu otkupa bila tvrdnja da je cena nula.
    Dim otpID2 As String
    otpID2 = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-PC2-" & scenario), _
                                      OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertTrue Len(otpID2) > 0, "OTP predlog cene: draft bez cene prosao"
    AssertEquals "", OtpStavkaPolje(otpID2, KLASA_I, COL_OPS_PREDLOG_CENA), _
                 "OTP predlog cene: bez predloga celija ostaje PRAZNA, ne 0"

    ' Cena na zaglavlju: odbijena, i to imenovano.
    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-PC3-" & scenario)
    h.Add "Cena", 250#

    Dim razlog2 As String
    AssertEquals "", CreateOtpremnicaDraft_TX(h, OtpOcek(400#, 20#, 0#, 0#), razlog2), _
                 "OTP predlog cene: cena na zaglavlju NE prolazi"
    AssertTrue InStr(1, razlog2, "nepoznat kljuc", vbTextCompare) > 0, _
               "OTP predlog cene: kapija imenuje razlog (bilo: " & razlog2 & ")"

    ' Negativan predlog je greska, kao i negativna cena na zaglavlju pre S3a.
    Dim c2 As Collection
    Set c2 = New Collection
    c2.Add OtpOcekStavkaSaCenom(KLASA_I, 400#, 20#, -1#)

    Dim razlog3 As String
    AssertEquals "", CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-PC4-" & scenario), _
                                              c2, razlog3), _
                 "OTP predlog cene: negativan predlog NE prolazi"

    Exit Sub

EH:
    LogFatal "Test_OTP_PredlogCeneJePoKlasi", Err.Number, Err.description
End Sub

' AMBALAZA SE KNJIZI PRI IZDAVANJU, NE PRI OTVARANJU NACRTA (odluka S14.8 t. 1).
'
' Gajbe odlaze sa stanice kad ih vozac preuzme. Nacrt je najava: sme da se menja
' i sme da ostane neizdat. Da se ambalaza knjizila sa nacrtom, napusten nacrt bi
' trajno umanjio stanje gajbi na otkupnom mestu, a svaka izmena ocekivanja bi
' trazila storniranje knjizenja.
Private Sub Test_OTP_AmbalazaSeKnjiziPriIzdavanju()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPAMB")

    Dim izvor As String
    izvor = OtpNoviOtkup(scenario, 400#, 0#)      ' Klasa I: 400 kg, 20 gajbi

    Dim razlog As String
    Dim otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-AM-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertTrue Len(otpID) > 0, "OTP ambalaza: draft prosao (bilo: " & razlog & ")"

    AssertEquals "0", CStr(AmbRedovaZaDokument(otpID)), _
                 "OTP ambalaza: NACRT ne knjizi nijednu gajbu"

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, izvor, razlog), _
               "OTP ambalaza: izvor vezan (bilo: " & razlog & ")"
    AssertEquals "0", CStr(AmbRedovaZaDokument(otpID)), _
                 "OTP ambalaza: ni vezivanje izvora ne knjizi gajbe"

    AssertTrue IzdajOtpremnicu_TX(otpID, razlog), _
               "OTP ambalaza: izdavanje proslo (bilo: " & razlog & ")"

    AssertEquals "1", CStr(AmbRedovaZaDokument(otpID)), _
                 "OTP ambalaza: izdavanje knjizi TACNO jedan red"
    AssertTrue Abs(AmbKolicinaZaDokument(otpID) - 20#) < 0.001, _
               "OTP ambalaza: kolicina je ZBIR STAVKI izdate otpremnice"
    AssertEquals "Izlaz", AmbPoljeZaDokument(otpID, COL_AMB_SMER), _
                 "OTP ambalaza: smer je izlaz sa stanice"
    AssertEquals TEST_ST_ID, AmbPoljeZaDokument(otpID, COL_AMB_ENTITET), _
                 "OTP ambalaza: gajbe odlaze sa OTKUPNOG MESTA otpremnice"
    AssertEquals TEST_TIP_AMB, AmbPoljeZaDokument(otpID, COL_AMB_TIP), _
                 "OTP ambalaza: knjizi se PO TIPU gajbe"

    Exit Sub

EH:
    LogFatal "Test_OTP_AmbalazaSeKnjiziPriIzdavanju", Err.Number, Err.description
End Sub

' F2 OTVARA NACRT (S3a cutover).
'
' Ekran je do sada zvao SaveOtpremnicaMulti_TX, koji je pravio GOTOV dokument sa
' linijskim poljima na zaglavlju -- i to po jedan RED PO KLASI, pod istim brojem.
' Sada isti unos otvara JEDAN nacrt sa ocekivanjem po klasi.
'
' Tvrdnja o povratnoj vrednosti nije kozmeticka: ekran mora da dobije ID, jer se
' nacrt kasnije izdaje i menja PO ID-u, a broj je jedinstven tek po (otkupno
' mesto, dan).
Private Sub Test_OTP_F2OtvaraNacrt()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPF2")

    Dim p As Object, fokus As String, greska As String, poruke As String
    Dim res As String

    Set p = modDokUnos.NoviOtpremnicaUnos()
    p("datum") = NextTestDate()
    p("stanicaID") = TEST_ST_ID
    p("vozacID") = TEST_VOZ_ID
    p("brDok") = TEST_PREFIX & "-OTP-F2-" & scenario
    p("vrsta") = TEST_VRSTA
    p("sorta") = TEST_SORTA
    p("tipAmb") = TEST_TIP_AMB
    p("kolicinaI") = 400#
    p("cenaI") = 250#
    p("kolAmb") = 20
    p("dveKlase") = True
    p("kolicinaII") = 600#
    p("cenaII") = 120#
    p("kolAmbII") = 30

    greska = modDokUnos.OtpremnicaValidiraj(p, fokus)
    AssertEquals "", greska, "OTP F2: unos prosao provere (fokus: " & fokus & ")"

    res = modDokUnos.OtpremnicaUpisi(p, poruke)
    AssertTrue Len(res) > 0, "OTP F2: upis prosao (bilo: " & poruke & ")"
    AssertTrue Left$(res, 4) = "OTP-", _
               "OTP F2: vraca se OtpremnicaID, ne broj dokumenta (bilo: " & res & ")"

    AssertEquals "1", CStr(FindRows(TBL_OTPREMNICA, COL_OTP_ID, res).count), _
                 "OTP F2: dvoklasan unos je JEDAN header, ne dva reda"
    AssertEquals IZDATO_DRAFT, OtpPolje(res, COL_TRACE_IZDATO_STATUS), _
                 "OTP F2: dokument je NACRT dok mu se ne vezu blokovi"
    AssertEquals "2", CStr(OtpBrojStavki(res)), "OTP F2: dve ocekivane stavke"

    ' Kolicina se poredi sa recnikom POSLE validacije: u bruto rezimu je
    ' validacija vec pretvorila uneto u neto, pa bi fiksan broj merio rezim.
    AssertTrue Abs(OtpStavkaBrojP(res, KLASA_I, COL_OPS_KOLICINA) - CDbl(p("kolicinaI"))) < 0.001, _
               "OTP F2: ocekivana kolicina I je ono sto je operater uneo"
    AssertTrue Abs(OtpStavkaBrojP(res, KLASA_II, COL_OPS_KOLICINA) - CDbl(p("kolicinaII"))) < 0.001, _
               "OTP F2: ocekivana kolicina II je ono sto je operater uneo"
    AssertTrue Abs(OtpStavkaBrojP(res, KLASA_I, COL_OPS_PREDLOG_CENA) - 250#) < 0.001, _
               "OTP F2: cena I sa ekrana je predlog cene KLASE I"
    AssertTrue Abs(OtpStavkaBrojP(res, KLASA_II, COL_OPS_PREDLOG_CENA) - 120#) < 0.001, _
               "OTP F2: cena II sa ekrana je predlog cene KLASE II"

    ' Kulturu razresava ADAPTER iz vrste i sorte -- pisac je samo proverava.
    AssertEquals TEST_KULTURA_ID, OtpPolje(res, COL_OTP_KULTURA), _
                 "OTP F2: kultura razresena iz vrste i sorte"
    AssertEquals CStr(p("brDok")), OtpPolje(res, COL_OTP_BROJ), _
                 "OTP F2: broj ostaje labela na zaglavlju"

    ' Nacrt ne knjizi gajbe (v. Test_OTP_AmbalazaSeKnjiziPriIzdavanju).
    AssertEquals "0", CStr(AmbRedovaZaDokument(res)), _
                 "OTP F2: otvaranje nacrta ne knjizi ambalazu"

    ' Nepoznata vrsta/sorta ne sme da napravi otpremnicu bez kulture.
    Dim p2 As Object, res2 As String, poruke2 As String
    Set p2 = modDokUnos.NoviOtpremnicaUnos()
    p2("datum") = NextTestDate()
    p2("stanicaID") = TEST_ST_ID
    p2("vozacID") = TEST_VOZ_ID
    p2("brDok") = TEST_PREFIX & "-OTP-F2X-" & scenario
    p2("vrsta") = "NEPOSTOJECA VRSTA " & scenario
    p2("sorta") = "NEPOSTOJECA SORTA"
    p2("tipAmb") = TEST_TIP_AMB
    p2("kolicinaI") = 400#
    p2("kolAmb") = 20

    Dim preRedova As Long
    preRedova = CountRows(TBL_OTPREMNICA)
    res2 = modDokUnos.OtpremnicaUpisi(p2, poruke2)
    AssertEquals "", res2, "OTP F2: nepoznata kultura ne pravi otpremnicu"
    AssertEquals CStr(preRedova), CStr(CountRows(TBL_OTPREMNICA)), _
                 "OTP F2: pad razresavanja ne ostavlja red u tabeli"

    Exit Sub

EH:
    LogFatal "Test_OTP_F2OtvaraNacrt", Err.Number, Err.description
End Sub

' MALINA AUTO-ZBIRNA JE PAUZIRANA, I TO GLASNO (S3a).
'
' AutoCreateZbirnaFromOtpremnice cita Kolicina / Klasa / KolAmbalaze sa
' ZAGLAVLJA otpremnice i ne gleda IzdatoStatus -- nad nacrtom bi napravila
' zbirnu sa 0 kg, od dokumenta koji jos nije isporuka. Zbirna prelazi na nov
' model u S4.
'
' Tvrdnja ima dva dela i oba su potrebna: da zbirna NIJE nastala, i da je
' operater o tome OBAVESTEN. Tiha pauza bi znacila da malina operater ceka
' zbirnu koja nikad nece doci.
Private Sub Test_OTP_MalinaZbirnaPauzirana()
    Dim prevMode As String, prevKupac As String

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPMAL")

    prevMode = GetConfigValue(CFG_KEY_MALINA_MODE)
    prevKupac = GetConfigValue(CFG_MALINA_DEFAULT_KUPAC)
    SetConfigValue CFG_KEY_MALINA_MODE, "YES"
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, TEST_KUP_ID

    Dim p As Object, fokus As String, poruke As String, res As String
    Set p = modDokUnos.NoviOtpremnicaUnos()
    p("datum") = NextTestDate()
    p("stanicaID") = TEST_ST_ID
    p("vozacID") = TEST_VOZ_ID
    p("brDok") = TEST_PREFIX & "-OTP-ML-" & scenario
    p("vrsta") = TEST_VRSTA
    p("sorta") = TEST_SORTA
    p("tipAmb") = TEST_TIP_AMB
    p("kolicinaI") = 400#
    p("cenaI") = 250#
    p("kolAmb") = 20

    Call modDokUnos.OtpremnicaValidiraj(p, fokus)

    Dim zbrPre As Long
    zbrPre = CountRows(TBL_ZBIRNA)

    res = modDokUnos.OtpremnicaUpisi(p, poruke)
    AssertTrue Len(res) > 0, "OTP malina: nacrt otvoren (bilo: " & poruke & ")"

    AssertEquals CStr(zbrPre), CStr(CountRows(TBL_ZBIRNA)), _
                 "OTP malina: nad nacrtom NE nastaje zbirna"
    AssertTrue InStr(1, poruke, Poruka("DOKUNOS_MSG_ZBIRNA_PAUZIRANA"), vbTextCompare) > 0, _
               "OTP malina: operater je OBAVESTEN da zbirne nema (bilo: " & poruke & ")"

    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    Exit Sub

EH:
    On Error Resume Next
    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    On Error GoTo 0
    LogFatal "Test_OTP_MalinaZbirnaPauzirana", Err.Number, Err.description
End Sub

' NOV NACRT NIJE ZAMENA ZA STORNIRANU OTPREMNICU (S3a, review #361 P1).
'
' Do S3a je F2 posle upisa zvao ZavrsiIspravkuAko FLOW_DOC_OTPREMNICA. Taj tok
' nije zatvaranje konteksta nego pisac STAROG modela: CompleteOtpremnicaIspravka
' preko ReassignOtkupToOtpremnica_TX upisuje Otkup.OtpremnicaID i BrojZbirne, pa
' rekalkulise zbirnu.
'
' Nad upravo otvorenim NACRTOM to je dvostruko pogresno: nov dokument bi se
' vezivao STAROM vezom, i dokument koji jos nema nijedan izvor ni status IZDATO
' bio bi proglasen zamenom izdate otpremnice -- a zamena sme da bude gotova tek
' posle clanstva, jednakosti i izdavanja.
'
' Tvrdnja ima cetiri dela, i svaki meri po jednu posledicu tog poziva.
Private Sub Test_OTP_NacrtNijeZavrsetakIspravke()
    Dim cid As String

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPISP")

    ' Stara, uredno IZDATA otpremnica sa jednim izvorom.
    Dim izvor As String
    izvor = OtpNoviOtkup(scenario, 400#, 0#)

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add izvor

    Dim razlog As String
    Dim staraOtp As String
    staraOtp = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-IS-" & scenario), _
                                           izvori, razlog)
    AssertTrue Len(staraOtp) > 0, "OTP ispravka: stara otpremnica izdata (bilo: " & razlog & ")"

    ' SYNTHETIC ANOMALY: stara veza Otkup.OtpremnicaID. Nijedan ziv pisac je vise
    ' ne postavlja (kolona odlazi u S3e) -- postavlja se rucno bas zato da bi se
    ' videlo da je legacy tok NE dira.
    VeziOtkupZaOtpremnicuFixture izvor, staraOtp

    cid = modStornoContext.CreateCorrectionContext(SV_MODE_ISPRAVKA, FLOW_DOC_OTPREMNICA, _
                                                   staraOtp, OtpPolje(staraOtp, COL_OTP_BROJ))
    AssertTrue Len(cid) > 0, "OTP ispravka: correction kontekst napravljen"

    Dim zbrPre As Long
    zbrPre = CountRows(TBL_ZBIRNA)

    ' F2: operater unosi novu otpremnicu dok ispravka stoji na cekanju.
    Dim p As Object, fokus As String, poruke As String, res As String
    Set p = modDokUnos.NoviOtpremnicaUnos()
    p("datum") = NextTestDate()
    p("stanicaID") = TEST_ST_ID
    p("vozacID") = TEST_VOZ_ID
    p("brDok") = TEST_PREFIX & "-OTP-ISN-" & scenario
    p("vrsta") = TEST_VRSTA
    p("sorta") = TEST_SORTA
    p("tipAmb") = TEST_TIP_AMB
    p("kolicinaI") = 400#
    p("cenaI") = 250#
    p("kolAmb") = 20

    Call modDokUnos.OtpremnicaValidiraj(p, fokus)
    res = modDokUnos.OtpremnicaUpisi(p, poruke)
    AssertTrue Len(res) > 0, "OTP ispravka: nacrt otvoren (bilo: " & poruke & ")"

    ' (1) Kontekst ostaje otvoren -- nacrt nije zamena.
    AssertEquals SV_STATUS_PENDING, _
                 modStornoContext.GetCorrectionField(cid, COL_SV_STATUS), _
                 "OTP ispravka: correction NIJE zavrsen nad nacrtom"

    ' (2) Stara veza je netaknuta -- legacy relink nije radio.
    AssertEquals staraOtp, OtkPolje(izvor, COL_OTK_OTPREMNICA_ID), _
                 "OTP ispravka: Otkup.OtpremnicaID se NE prevezuje na nacrt"

    ' (3) Zbirna nije dirana.
    AssertEquals CStr(zbrPre), CStr(CountRows(TBL_ZBIRNA)), _
                 "OTP ispravka: zbirna se NE rekalkulise"

    ' (4) Operater to ZNA. Tiho preskakanje bi znacilo da misli da je ispravka
    ' zavrsena, a ona i dalje ceka na ekranu Oporavak.
    AssertTrue InStr(1, poruke, Poruka("DOKUNOS_MSG_OTP_ISPRAVKA_PAUZIRANA"), _
                     vbTextCompare) > 0, _
               "OTP ispravka: operater je OBAVESTEN da ispravka ceka (bilo: " & poruke & ")"

    Call modStornoContext.CancelCorrectionContext(cid, "S3a test cleanup")
    Exit Sub

EH:
    On Error Resume Next
    If Len(cid) > 0 Then Call modStornoContext.CancelCorrectionContext(cid, "S3a test cleanup")
    On Error GoTo 0
    LogFatal "Test_OTP_NacrtNijeZavrsetakIspravke", Err.Number, Err.description
End Sub

' === S3b: citaoci otpremnice citaju STAVKE ===================================

' MREZA F2 CITA STAVKE, NE ZAGLAVLJE (S3b).
'
' Od S3a zaglavlje otpremnice ostaje bez Klase, Kolicine, KolAmbalaze i Cene --
' nacrt ih pise na stavke. Mreza koja bi i dalje citala zaglavlje pokazala bi
' svaki nov dokument kao prazan red: 0 kg, bez klase, bez vrednosti. Isti kvar
' koji je otkup imao pre S14.7, samo na drugom dokumentu.
'
' DVE KLASE = JEDAN RED. Pre S3a su dve klase bile dva zaglavlja pod istim
' brojem, pa ih je mreza crtala kao dva dokumenta.
Private Sub Test_OTP_MrezaCitaStavke()
    On Error GoTo EH

    Dim scenario As String, broj As String
    scenario = NewScenarioCode("OTPMR")
    broj = TEST_PREFIX & "-OTP-MR-" & scenario

    Dim c As Collection
    Set c = New Collection
    c.Add OtpOcekStavkaSaCenom(KLASA_I, 400#, 20#, 250#)
    c.Add OtpOcekStavkaSaCenom(KLASA_II, 600#, 30#, 120#)

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(broj), c, razlog)
    AssertTrue Len(otpID) > 0, "OTP mreza: nacrt napravljen (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    ' Zaglavlje je JEDNO, pa je i red JEDAN -- dve klase nisu dva dokumenta.
    AssertEquals "1", CStr(OtpMrezaBrojRedova(broj)), _
                 "OTP mreza: jedan dokument = jedan red"

    Dim red As Variant
    red = OtpMrezaRed(broj)
    AssertTrue IsArray(red), "OTP mreza: red dokumenta nadjen"
    If Not IsArray(red) Then Exit Sub

    AssertTrue Abs(CDbl(red(0)) - 1000#) < 0.001, _
               "OTP mreza: kg su zbir stavki (400 + 600), ne prazno zaglavlje"
    AssertTrue Abs(CDbl(red(1)) - 50#) < 0.001, _
               "OTP mreza: gajbe su zbir stavki (20 + 30)"
    AssertEquals KLASA_I & ", " & KLASA_II, CStr(red(2)), _
                 "OTP mreza: kolona klase nabraja obe klase dokumenta"

    ' Kolone vrednosti NEMA (review #362, P1): PredlogCena je predlog za
    ' prefill, a Kolicina x PredlogCena nije vrednost dokumenta.
    AssertTrue Not OtpMrezaImaKolonu("OTKUI_HD_VREDNOST"), _
               "OTP mreza: nema kolone vrednosti -- predlog cene nije finansijska cifra"

    Exit Sub

EH:
    LogFatal "Test_OTP_MrezaCitaStavke", Err.Number, Err.description
End Sub

' DRUGA BRANA CITAOCA: otpremnica bez ijedne stavke PADA PO IMENU.
'
' Pisac takav dokument ne moze da napravi (OtpUpisiOcekivano odbija prazno
' ocekivanje), pa je ovo SINTETICKA ANOMALIJA -- pravi se brisanjem stavki u
' transakciji koja se vraca. Meri se tacno ono sto je kod otkupa bio kvar
' (review #334, P1): citalac koji nedostajuci kljuc procita kao nulu nacrta
' dokument sa 0 kg umesto da kaze koji dokument je pokvaren.
Private Sub Test_OTP_ZaglavljeBezStavkiObaraCitaoce()
    Const SRC As String = "Test_OTP_ZaglavljeBezStavkiObaraCitaoce"
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String, broj As String
    scenario = NewScenarioCode("OTPBS")
    broj = TEST_PREFIX & "-OTP-BS-" & scenario

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(broj), OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertTrue Len(otpID) > 0, "OTP bez stavki: nacrt napravljen (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    AssertEquals "", OtpMrezaGreska(broj), "OTP bez stavki: mreza prolazi pre kvarenja"

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA_STAVKE

    Dim rows As Collection
    Set rows = FindRows(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, otpID)
    AssertTrue Not rows Is Nothing, "OTP bez stavki: stavke nadjene"
    If rows Is Nothing Then Exit Sub

    Dim k As Long
    For k = rows.count To 1 Step -1
        RequireDeleteRow TBL_OTPREMNICA_STAVKE, CLng(rows(k)), SRC
    Next k

    AssertTrue InStr(1, OtpMrezaGreska(broj), "nema nijednu stavku", vbTextCompare) > 0, _
               "OTP bez stavki: mreza pada po imenu, ne crta 0 kg"

    tx.RollbackTx
    Set tx = Nothing

    AssertEquals "", OtpMrezaGreska(broj), _
                 "OTP bez stavki: citalac prolazi posle vracanja"

    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogFatal SRC, errNum, errDesc
End Sub

' IZVESTAJ PO OTKUPNOM MESTU: JEDAN RED PO KLASI (S3b).
'
' Manjak se razresava kroz stavku zbirne, a njen kljuc nosi KLASU -- kad bi
' dvoklasna otpremnica dala jedan red, prijem obe klase bi se sabrao i pripisao
' jednoj (u malina modu bukvalno duplo). Zato zaglavlje ostaje jedno, a red
' izvestaja i dalje pripada jednoj klasi; menja se samo odakle klasa dolazi.
Private Sub Test_OTP_IzvestajOMRedPoKlasi()
    On Error GoTo EH

    Dim scenario As String, broj As String
    scenario = NewScenarioCode("OTPIOM")
    broj = TEST_PREFIX & "-OTP-IOM-" & scenario

    ' Izvestaj broji samo IZDATE otpremnice (review #362), pa se ova pravi iz
    ' izvora -- jednim potezom, koji je i izdaje.
    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add OtpNoviOtkup(scenario, 400#, 600#)

    Dim h As Object
    Set h = OtpHeader(broj)
    Dim dan As Date
    dan = CDate(h("Datum"))

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaIzIzvora_TX(h, izvori, razlog)
    AssertTrue Len(otpID) > 0, "OTP izvestaj OM: otpremnica izdata (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    modUiData.ResetCache
    Dim r As Variant
    r = ReportOtkupRoba("OM", TEST_ST_ID, dan, dan)
    AssertTrue IsArray(r), "OTP izvestaj OM: izvestaj je vratio redove"
    If Not IsArray(r) Then Exit Sub

    ' Kolone: (2) BrOtp (4) Klasa (6) Otp kg -- v. IzKoloneZaListu, tip "OM".
    Dim i As Long, n As Long, kgI As Double, kgII As Double
    For i = 1 To UBound(r, 1)
        If Trim$(CStr(r(i, 2))) = broj Then
            n = n + 1
            If Trim$(CStr(r(i, 4))) = KLASA_I Then kgI = CDbl(r(i, 6))
            If Trim$(CStr(r(i, 4))) = KLASA_II Then kgII = CDbl(r(i, 6))
        End If
    Next i

    AssertEquals "2", CStr(n), "OTP izvestaj OM: dvoklasna otpremnica daje DVA reda"
    AssertTrue Abs(kgI - 400#) < 0.001, "OTP izvestaj OM: red klase I nosi svoju kilazu"
    AssertTrue Abs(kgII - 600#) < 0.001, "OTP izvestaj OM: red klase II nosi SVOJU kilazu"

    Exit Sub

EH:
    LogFatal "Test_OTP_IzvestajOMRedPoKlasi", Err.Number, Err.description
End Sub

' OTPREMLJENO JE SAMO IZDATO (review #362, P1).
'
' Nacrt je najava: nema izvora, gajbe nisu knjizene i sme da ostane neizdat. Do
' ovog review-a su "roba po vozacu" i "roba po otkupnom mestu" brojale svaku
' nestorniranu otpremnicu -- pa je nacrt od 1000 kg, bez ijednog povezanog
' otkupa, vec bio otpremljena roba. Test prati jedan dokument kroz ceo tok:
' nacrt ne ulazi, izdat ulazi TACNO jednom, storniran izlazi.
Private Sub Test_OTP_OtpremljenoJeSamoIzdato()
    On Error GoTo EH

    Dim scenario As String, broj As String
    scenario = NewScenarioCode("OTPIZD")
    broj = TEST_PREFIX & "-OTP-IZD-" & scenario

    Dim izvor As String
    izvor = OtpNoviOtkup(scenario, 600#, 400#)      ' I: 600 kg / 20 gajbi, II: 400 / 30

    Dim h As Object
    Set h = OtpHeader(broj)
    Dim dan As Date
    dan = CDate(h("Datum"))

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(h, OtpOcek(600#, 20#, 400#, 30#), razlog)
    AssertTrue Len(otpID) > 0, "OTP otpremljeno: nacrt napravljen (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    AssertEquals Format$(0#, "0.00"), OtpVozacRoba(dan, 3), _
                 "OTP otpremljeno: nacrt nije otpremljena roba (roba po vozacu)"
    AssertEquals "0", CStr(OtpOmRedova(dan, broj)), _
                 "OTP otpremljeno: nacrt nije otpremljena roba (roba po OM)"

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, izvor, razlog), _
               "OTP otpremljeno: izvor vezan (bilo: " & razlog & ")"
    AssertTrue IzdajOtpremnicu_TX(otpID, razlog), _
               "OTP otpremljeno: izdavanje proslo (bilo: " & razlog & ")"

    AssertEquals Format$(1000#, "0.00"), OtpVozacRoba(dan, 3), _
                 "OTP otpremljeno: izdata ulazi TACNO jednom -- 1000 kg"
    AssertEquals "2", CStr(OtpOmRedova(dan, broj)), _
                 "OTP otpremljeno: izdata je u robi po OM, red po klasi"

    MarkTestRowStornirano TBL_OTPREMNICA, "OtpremnicaID", otpID
    AssertEquals Format$(0#, "0.00"), OtpVozacRoba(dan, 3), _
                 "OTP otpremljeno: stornirana izdata otpremnica ne ulazi"

    Exit Sub

EH:
    LogFatal "Test_OTP_OtpremljenoJeSamoIzdato", Err.Number, Err.description
End Sub

' VREDNOST OTPREMNICE JE VREDNOST NJENIH IZVORA (review #362, P1).
'
' Dva otkupna bloka iste klase, placena RAZLICITO (50 i 40 din), idu u jednu
' otpremnicu ciji je predlog cene 999. Vrednost mora biti ono sto je placeno --
' 300 x 50 + 200 x 40 = 23000 -- a ne 500 x 999. Predlog je polje za prefill
' otkupa i ne sme da zameni stvarne cene.
Private Sub Test_OTP_VrednostIzIzvoraNePredlogCene()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPVRI")

    Dim otkA As String, otkB As String
    otkA = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-VA-" & scenario), _
                          OtkStavke(300#, 50#, 10, 0#, 0#, 0))
    otkB = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-VB-" & scenario), _
                          OtkStavke(200#, 40#, 10, 0#, 0#, 0))
    AssertTrue Len(otkA) > 0 And Len(otkB) > 0, "OTP vrednost: izvori napravljeni"

    Dim c As Collection
    Set c = New Collection
    c.Add OtpOcekStavkaSaCenom(KLASA_I, 500#, 20#, 999#)

    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-VRI-" & scenario)
    Dim dan As Date
    dan = CDate(h("Datum"))

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(h, c, razlog)
    AssertTrue Len(otpID) > 0, "OTP vrednost: nacrt napravljen (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, otkA, razlog), _
               "OTP vrednost: prvi blok vezan (bilo: " & razlog & ")"
    AssertTrue DodajOtpremnicaIzvor_TX(otpID, otkB, razlog), _
               "OTP vrednost: drugi blok vezan (bilo: " & razlog & ")"
    AssertTrue IzdajOtpremnicu_TX(otpID, razlog), _
               "OTP vrednost: izdavanje proslo (bilo: " & razlog & ")"

    AssertEquals Format$(500#, "0.00"), OtpVozacRoba(dan, 3), _
                 "OTP vrednost: kilaza je 300 + 200"
    AssertEquals Format$(23000#, "0.00"), OtpVozacRoba(dan, 4), _
                 "OTP vrednost: vrednost je ono sto je placeno (300x50 + 200x40), ne 500x999"

    Exit Sub

EH:
    LogFatal "Test_OTP_VrednostIzIzvoraNePredlogCene", Err.Number, Err.description
End Sub

' PRAVILO "IZDATA" ZA SVIH PET STANJA (review #362, drugi krug).
'
' Operativni citaoci (roba po vozacu, roba po OM, stampa) odlucuju kroz
' IzdatoStatusJeIzdato. PROSLEDJENO je izdat dokument koji je i otisao dalje --
' da ga pravilo ne broji, buduci sync bi retroaktivno izbrisao otpremljenu robu
' iz izvestaja. Prazan i nepoznat status NISU izdati: u novom modelu otpremnica
' nastaje kao nacrt, pa samo imenovan status dokazuje izdavanje.
Private Sub Test_OTP_IzdatoStatusPravilo()
    On Error GoTo EH

    AssertTrue Not modDokumenta.IzdatoStatusJeIzdato(IZDATO_DRAFT), _
               "OTP status: DRAFT nije izdat"
    AssertTrue modDokumenta.IzdatoStatusJeIzdato(IZDATO_IZDATO), _
               "OTP status: IZDATO je izdat"
    AssertTrue modDokumenta.IzdatoStatusJeIzdato(IZDATO_PROSLEDJENO), _
               "OTP status: PROSLEDJENO je izdat -- sync ne brise otpremljenu robu"
    AssertTrue Not modDokumenta.IzdatoStatusJeIzdato(""), _
               "OTP status: prazan status nije izdat"
    AssertTrue Not modDokumenta.IzdatoStatusJeIzdato("NESTO"), _
               "OTP status: nepoznat status nije izdat"

    ' Celija sme da nosi razmake i mala slova -- pravilo poredi normalizovano.
    AssertTrue modDokumenta.IzdatoStatusJeIzdato("  " & LCase$(IZDATO_PROSLEDJENO) & " "), _
               "OTP status: razmaci i mala slova ne menjaju ishod"

    Exit Sub

EH:
    LogFatal "Test_OTP_IzdatoStatusPravilo", Err.Number, Err.description
End Sub

' JEDNA STAVKA PO KLASI -- citalac drzi isto sto i pisac (review #362, P2).
'
' Pisac odbija dve stavke iste klase (OtpUpisiOcekivano), pa je dokument sa
' dve stavke klase I SINTETICKA ANOMALIJA -- pravi se dodavanjem reda u
' transakciji koja se vraca. Bez kapije u citaocu takav dokument bi u
' izvestaju bio sabran kao 2 x I, a u stampi dao dva reda iste klase.
Private Sub Test_OTP_DveStavkeIsteKlaseObaraCitaoce()
    Const SRC As String = "Test_OTP_DveStavkeIsteKlaseObaraCitaoce"
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String, broj As String
    scenario = NewScenarioCode("OTPDK")
    broj = TEST_PREFIX & "-OTP-DK-" & scenario

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(broj), OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertTrue Len(otpID) > 0, "OTP dve iste klase: nacrt napravljen (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA_STAVKE

    Dim rowData As Variant
    rowData = BlankRow(TBL_OTPREMNICA_STAVKE)
    SetRequiredField rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_ID, otpID & "-DUPLA"
    SetRequiredField rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, otpID
    SetRequiredField rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_RB, 2
    SetRequiredField rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_KLASA, KLASA_I
    SetRequiredField rowData, TBL_OTPREMNICA_STAVKE, COL_OPS_KOLICINA, 100#
    RequireAppend TBL_OTPREMNICA_STAVKE, rowData, SRC

    AssertTrue InStr(1, OtpMrezaGreska(broj), "Dve stavke iste klase", vbTextCompare) > 0, _
               "OTP dve iste klase: mreza pada po imenu, ne sabira 2 x I"

    tx.RollbackTx
    Set tx = Nothing

    AssertEquals "", OtpMrezaGreska(broj), _
                 "OTP dve iste klase: citalac prolazi posle vracanja"

    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogFatal SRC, errNum, errDesc
End Sub

' INVARIJANTA ZBIRNE SABIRA STAVKE (S3b).
'
' Zbirna je po invarijanti tacno zbir svojih aktivnih otpremnica, PO KLASI. Kad
' bi taj zbir i dalje dolazio sa zaglavlja, svaka bi zbirna od S3a bila
' poredjena sa nulom i proglasena neispravnom -- i to ne jednom, nego na svakoj
' izmeni koja invariant proverava.
'
' BrojZbirne se ovde upisuje rucno, u transakciji koja se vraca: vezivanje
' otpremnice za zbirnu je stari model i prelazi tek u S4.
Private Sub Test_OTP_InvarijantaSabiraStavke()
    Const SRC As String = "Test_OTP_InvarijantaSabiraStavke"
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String, broj As String, brZbr As String
    scenario = NewScenarioCode("OTPINV")
    broj = TEST_PREFIX & "-OTP-INV-" & scenario
    brZbr = TEST_PREFIX & "-ZBR-INV-" & scenario

    Dim c As Collection
    Set c = New Collection
    c.Add OtpOcekStavka(KLASA_I, 400#, 20#)
    c.Add OtpOcekStavka(KLASA_II, 600#, 30#)

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(broj), c, razlog)
    AssertTrue Len(otpID) > 0, "OTP invarijanta: nacrt napravljen (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA

    Dim hdr As Collection
    Set hdr = FindRows(TBL_OTPREMNICA, COL_OTP_ID, otpID)
    AssertTrue Not hdr Is Nothing, "OTP invarijanta: zaglavlje nadjeno"
    If hdr Is Nothing Then Exit Sub
    RequireUpdateCell TBL_OTPREMNICA, CLng(hdr(1)), COL_OTP_BROJ_ZBIRNE, brZbr, SRC

    modUiData.ResetCache
    Dim d As Object
    Set d = modDokumentInvariant.SumOtpremniceByKlasa(brZbr)

    AssertTrue Abs(CDbl(d("kgI")) - 400#) < 0.001, _
               "OTP invarijanta: kg klase I dolaze sa stavke"
    AssertTrue Abs(CDbl(d("kgII")) - 600#) < 0.001, _
               "OTP invarijanta: kg klase II dolaze sa stavke"
    AssertTrue Abs(CDbl(d("kgTotal")) - 1000#) < 0.001, _
               "OTP invarijanta: ukupno je zbir stavki, ne prazno zaglavlje"
    AssertEquals "50", CStr(CLng(d("ambTotal"))), _
                 "OTP invarijanta: gajbe su zbir stavki (20 + 30)"
    ' Druga strana poredjenja (tblZbirna) je red po klasi, pa i ovde broje STAVKE.
    AssertEquals "2", CStr(CLng(d("nRows"))), _
                 "OTP invarijanta: jedno zaglavlje sa dve klase broji DVE stavke"

    tx.RollbackTx
    Set tx = Nothing

    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogFatal SRC, errNum, errDesc
End Sub

' PREFILL ISPRAVKE CITA STAVKE (S3b).
'
' Ispravka otpremnice je od S3a PAUZIRANA na zavrsetku (B-038 vraca S3c), ali se
' prefill i dalje nudi kad operater otvori ispravku sa ekrana Oporavak. Prefill
' sa praznog zaglavlja bi mu ponudio dokument bez kilaze i bez klase, pa bi
' ispravka "sacuvala" nesto sto stornirani dokument nikad nije bio.
Private Sub Test_OTP_PrefillIspravkeCitaStavke()
    On Error GoTo EH

    Dim scenario As String, broj As String
    scenario = NewScenarioCode("OTPPF")
    broj = TEST_PREFIX & "-OTP-PF-" & scenario

    Dim c As Collection
    Set c = New Collection
    c.Add OtpOcekStavkaSaCenom(KLASA_I, 400#, 20#, 250#)
    c.Add OtpOcekStavkaSaCenom(KLASA_II, 600#, 30#, 120#)

    Dim razlog As String, otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(broj), c, razlog)
    AssertTrue Len(otpID) > 0, "OTP prefill: nacrt napravljen (bilo: " & razlog & ")"
    If Len(otpID) = 0 Then Exit Sub

    modUiData.ResetCache
    Dim spec As String
    spec = modStornoDok.PrefillIzStorniranog(STIP_OTPREMNICA, broj, otpID)

    AssertTrue InStr(1, spec, "kol1=400", vbTextCompare) > 0, _
               "OTP prefill: klasa I nosi svoju kilazu (bilo: " & spec & ")"
    AssertTrue InStr(1, spec, "kol2=600", vbTextCompare) > 0, _
               "OTP prefill: klasa II nosi svoju kilazu"
    AssertTrue InStr(1, spec, "amb1=20", vbTextCompare) > 0, _
               "OTP prefill: gajbe klase I dolaze sa stavke"
    AssertTrue InStr(1, spec, "dveklase=2", vbTextCompare) > 0, _
               "OTP prefill: broj klasa se broji sa stavki"
    AssertTrue InStr(1, spec, "cena=250", vbTextCompare) > 0, _
               "OTP prefill: predlog cene klase I"
    AssertTrue InStr(1, spec, "cena2=120", vbTextCompare) > 0, _
               "OTP prefill: predlog cene klase II je SVOJ, ne cena prve"

    Exit Sub

EH:
    LogFatal "Test_OTP_PrefillIspravkeCitaStavke", Err.Number, Err.description
End Sub

' Jednopotezni ulaz izvodi ocekivanje iz izvora -- tu nezavisnog operaterskog
' ocekivanja nema.
Private Sub Test_OTP_StavkeSuIzvedene()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPIZ")

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add OtpNoviOtkup(scenario & "A", 400#, 600#)
    izvori.Add OtpNoviOtkup(scenario & "B", 100#, 50#)

    Dim razlog As String
    Dim otpID As String
    otpID = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-IZ-" & scenario), _
                                        izvori, razlog)

    AssertTrue Len(otpID) > 0, "OTP izvedeno: upis prosao (bilo: " & razlog & ")"
    AssertEquals "2", CStr(OtpBrojClanova(otpID)), "OTP izvedeno: dva clana"

    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_I, COL_OPS_KOLICINA) - 500#) < 0.001, _
               "OTP izvedeno: Klasa I = 400 + 100"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_II, COL_OPS_KOLICINA) - 650#) < 0.001, _
               "OTP izvedeno: Klasa II = 600 + 50"

    ' TipAmbalaze dolazi IZ IZVORA -- header ga ni ne prima.
    AssertEquals TEST_TIP_AMB, OtpPolje(otpID, COL_OTP_TIP_AMB), "OTP izvedeno: TipAmbalaze"

    Exit Sub

EH:
    LogFatal "Test_OTP_StavkeSuIzvedene", Err.Number, Err.description
End Sub

' DRAFT NOSI OCEKIVANJE.
'
' Ovo je ono zbog cega panel postoji: operater prijavi sta otpremnica nosi, pa
' unosi otkupne listove gledajuci koliko je preostalo. Danas to ocekivanje zivi
' na Otpremnica.Kolicina, koja u ciljnom modelu odlazi na stavku (S4.2a).
Private Sub Test_OTP_DraftNosiOcekivanje()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPDR")

    Dim razlog As String
    Dim otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-DR-" & scenario), _
                                     OtpOcek(1000#, 50#, 500#, 25#), razlog)

    AssertTrue Len(otpID) > 0, "OTP draft: nastao (bilo: " & razlog & ")"
    AssertEquals IZDATO_DRAFT, OtpPolje(otpID, COL_TRACE_IZDATO_STATUS), _
                 "OTP draft: status je DRAFT"

    ' Stavke POSTOJE odmah -- one su ocekivanje, ne izveden kes.
    AssertEquals "2", CStr(OtpBrojStavki(otpID)), "OTP draft: dve stavke ocekivanja"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_I, COL_OPS_KOLICINA) - 1000#) < 0.001, _
               "OTP draft: ocekivano I = 1000"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_II, COL_OPS_KOL_AMB) - 25#) < 0.001, _
               "OTP draft: ocekivana ambalaza II = 25"
    AssertEquals "0", CStr(OtpBrojClanova(otpID)), "OTP draft: jos nista nije povezano"

    ' Kultura se zna OD OTVARANJA -- panel njome prefiluje formu otkupa.
    AssertEquals TEST_KULTURA_ID, OtpPolje(otpID, COL_OTP_KULTURA), "OTP draft: KulturaID"
    AssertEquals TEST_VRSTA, OtpPolje(otpID, COL_OTP_VRSTA), "OTP draft: VrstaVoca snapshot"
    AssertEquals TEST_SORTA, OtpPolje(otpID, COL_OTP_SORTA), "OTP draft: SortaVoca snapshot"

    ' Bruto operater ne prijavljuje -- dolazi iz izvora pri izdavanju.
    AssertEquals "", OtpStavkaPolje(otpID, KLASA_I, COL_OPS_BRUTO), _
                 "OTP draft: BrutoKg prazan na draftu"

    ' Draft bez ijedne stavke nema sta da meri.
    Dim rez As String
    rez = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-DR2-" & scenario), _
                                   New Collection, razlog)
    AssertEquals "", rez, "OTP draft: prazno ocekivanje odbijeno"
    AssertTrue InStr(1, razlog, "nema sta da meri", vbTextCompare) > 0, _
               "OTP draft: kapija imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_DraftNosiOcekivanje", Err.Number, Err.description
End Sub

' ocekivano / povezano / preostalo po klasi -- read-model panela.
Private Sub Test_OTP_NapredakPoKlasi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPNP")

    Dim otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-NP-" & scenario), _
                                     OtpOcek(1000#, 50#, 500#, 25#))

    Dim p As Object
    Set p = GetOtpremnicaProgress(otpID)

    AssertTrue Abs(p(UCase$(KLASA_I))("ocekivano") - 1000#) < 0.001, _
               "OTP napredak: ocekivano I = 1000"
    AssertTrue Abs(p(UCase$(KLASA_I))("povezano")) < 0.001, _
               "OTP napredak: povezano I = 0 pre izvora"
    AssertTrue Abs(p(UCase$(KLASA_I))("preostalo") - 1000#) < 0.001, _
               "OTP napredak: preostalo I = 1000 pre izvora"

    Dim razlog As String
    AssertTrue DodajOtpremnicaIzvor_TX(otpID, OtpNoviOtkup(scenario, 400#, 200#), razlog), _
               "OTP napredak: izvor dodat (bilo: " & razlog & ")"

    Set p = GetOtpremnicaProgress(otpID)
    AssertTrue Abs(p(UCase$(KLASA_I))("povezano") - 400#) < 0.001, _
               "OTP napredak: povezano I = 400"
    AssertTrue Abs(p(UCase$(KLASA_I))("preostalo") - 600#) < 0.001, _
               "OTP napredak: preostalo I = 600"
    AssertTrue Abs(p(UCase$(KLASA_II))("preostalo") - 300#) < 0.001, _
               "OTP napredak: preostalo II = 300"
    AssertTrue Abs(p(UCase$(KLASA_I))("povezanoAmb") - 20#) < 0.001, _
               "OTP napredak: povezana ambalaza I = 20"

    Exit Sub

EH:
    LogFatal "Test_OTP_NapredakPoKlasi", Err.Number, Err.description
End Sub

' Izdavanje trazi ocekivano = povezano. I manjak i VISAK obaraju.
Private Sub Test_OTP_IzdavanjeTraziJednakost()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPJD")

    ' manjak: prijavljeno 1000, povezano 400
    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-JD-" & scenario), _
                                     OtpOcek(1000#, 20#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID, OtpNoviOtkup(scenario & "A", 400#, 0#)

    AssertTrue Not IzdajOtpremnicu_TX(otpID, razlog), "OTP jednakost: manjak obara izdavanje"
    AssertTrue InStr(1, razlog, "preostalo", vbTextCompare) > 0, _
               "OTP jednakost: kapija imenuje preostalo (bilo: " & razlog & ")"
    AssertEquals IZDATO_DRAFT, OtpPolje(otpID, COL_TRACE_IZDATO_STATUS), _
                 "OTP jednakost: posle manjka ostaje DRAFT"

    ' visak: prijavljeno 400, povezano 500
    Dim otpID2 As String
    otpID2 = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-JD2-" & scenario), _
                                      OtpOcek(400#, 20#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID2, OtpNoviOtkup(scenario & "B", 500#, 0#)

    AssertTrue Not IzdajOtpremnicu_TX(otpID2, razlog), "OTP jednakost: visak obara izdavanje"
    AssertTrue InStr(1, razlog, "povezano", vbTextCompare) > 0, _
               "OTP jednakost: visak imenuje razlog (bilo: " & razlog & ")"

    ' tacno: prijavljeno 400 / 20, povezano 400 / 20
    Dim otpID3 As String
    otpID3 = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-JD3-" & scenario), _
                                      OtpOcek(400#, 20#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID3, OtpNoviOtkup(scenario & "C", 400#, 0#)

    AssertTrue IzdajOtpremnicu_TX(otpID3, razlog), _
               "OTP jednakost: jednako prolazi (bilo: " & razlog & ")"
    AssertEquals IZDATO_IZDATO, OtpPolje(otpID3, COL_TRACE_IZDATO_STATUS), _
                 "OTP jednakost: dokument izdat"

    Exit Sub

EH:
    LogFatal "Test_OTP_IzdavanjeTraziJednakost", Err.Number, Err.description
End Sub

' TOCTOU: izmedju Dodaj i Izdaj prolazi vreme.
'
' Bez revalidacije je "DRAFT -> dodaj OTK1 -> storno OTK1 -> Izdaj" izdavalo
' dokument iz storniranog izvora, jer je Izdaj verovao onome sto je Dodaj vec
' proverio.
Private Sub Test_OTP_IzdavanjeRevalidiraIzvore()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPRV")

    Dim otkID As String
    otkID = OtpNoviOtkup(scenario, 400#, 0#)

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-RV-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, otkID, razlog), _
               "OTP revalidacija: izvor dodat dok je bio ispravan"

    ' Izvor se u medjuvremenu stornira -- bas ono sto se u panelu desava.
    RequireUpdateCell TBL_OTKUP, FindRows(TBL_OTKUP, COL_OTK_ID, otkID)(1), _
                      COL_STORNIRANO, "Da", "Test_OTP_IzdavanjeRevalidiraIzvore"

    AssertTrue Not IzdajOtpremnicu_TX(otpID, razlog), _
               "OTP revalidacija: izdavanje iz storniranog izvora odbijeno"
    AssertTrue InStr(1, razlog, "storniran", vbTextCompare) > 0, _
               "OTP revalidacija: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals IZDATO_DRAFT, OtpPolje(otpID, COL_TRACE_IZDATO_STATUS), _
                 "OTP revalidacija: dokument ostaje DRAFT"

    Exit Sub

EH:
    LogFatal "Test_OTP_IzdavanjeRevalidiraIzvore", Err.Number, Err.description
End Sub

' Bruto se ne sabira parcijalno -- prazno nije nula.
'
' Jedan izvor 500 bruto / 480 neto, drugi 300 neto bez bruta: zbir bi dao
' Kolicina 780, BrutoKg 500, dakle bruto MANJI od neta. Fizicki nemoguc red.
Private Sub Test_OTP_BrutoSeNeSabiraParcijalno()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPBR")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-BR-" & scenario), _
                                     OtpOcek(780#, 40#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID, OtpNoviOtkupSaBrutom(scenario & "A", 480#, 500#)
    DodajOtpremnicaIzvor_TX otpID, OtpNoviOtkup(scenario & "B", 300#, 0#)

    AssertTrue IzdajOtpremnicu_TX(otpID, razlog), _
               "OTP bruto: izdavanje proslo (bilo: " & razlog & ")"
    AssertEquals "", OtpStavkaPolje(otpID, KLASA_I, COL_OPS_BRUTO), _
                 "OTP bruto: parcijalno poznat bruto ostaje PRAZAN"

    ' Kontrola: kad ga nose SVI izvori, bruto se upisuje. Bez ovoga bi kapija
    ' koja uvek ostavlja prazno izgledala isto kao kapija koja radi.
    Dim otpID2 As String
    otpID2 = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-BR2-" & scenario), _
                                      OtpOcek(780#, 40#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID2, OtpNoviOtkupSaBrutom(scenario & "C", 480#, 500#)
    DodajOtpremnicaIzvor_TX otpID2, OtpNoviOtkupSaBrutom(scenario & "D", 300#, 320#)

    AssertTrue IzdajOtpremnicu_TX(otpID2, razlog), _
               "OTP bruto: kontrola izdata (bilo: " & razlog & ")"
    AssertTrue Abs(OtpStavkaBrojP(otpID2, KLASA_I, COL_OPS_BRUTO) - 820#) < 0.001, _
               "OTP bruto: pun bruto se sabira (500 + 320)"

    Exit Sub

EH:
    LogFatal "Test_OTP_BrutoSeNeSabiraParcijalno", Err.Number, Err.description
End Sub

' Izvor mora da bude iste kulture kao otpremnica.
Private Sub Test_OTP_KulturaSeSlaziSaIzvorima()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPKU")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-KU-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))

    AssertTrue Not DodajOtpremnicaIzvor_TX(otpID, _
                       OtpNoviOtkupDrugeKulture(scenario), razlog), _
               "OTP kultura: izvor druge kulture odbijen"
    AssertTrue InStr(1, razlog, "KulturaID", vbTextCompare) > 0, _
               "OTP kultura: kapija imenuje polje (bilo: " & razlog & ")"

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, OtpNoviOtkup(scenario, 400#, 0#), razlog), _
               "OTP kultura: izvor iste kulture prolazi (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_KulturaSeSlaziSaIzvorima", Err.Number, Err.description
End Sub

' Draft se sme ispraviti dok nije izdat -- i zaglavlje i ocekivanje.
Private Sub Test_OTP_UpdateDraftaMenjaOcekivanje()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPUP")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-UP-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))

    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-UP-" & scenario)
    h("VozacID") = TEST_VOZ_ID_B

    AssertTrue UpdateOtpremnicaDraft_TX(otpID, h, OtpOcek(1000#, 50#, 0#, 0#), razlog), _
               "OTP update: prosao (bilo: " & razlog & ")"

    AssertEquals "1", CStr(OtpBrojStavki(otpID)), "OTP update: i dalje jedna stavka"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_I, COL_OPS_KOLICINA) - 1000#) < 0.001, _
               "OTP update: ocekivano promenjeno na 1000"
    AssertEquals TEST_VOZ_ID_B, OtpPolje(otpID, COL_OTP_VOZAC), "OTP update: vozac promenjen"
    AssertEquals IZDATO_DRAFT, OtpPolje(otpID, COL_TRACE_IZDATO_STATUS), _
                 "OTP update: i dalje DRAFT"

    Exit Sub

EH:
    LogFatal "Test_OTP_UpdateDraftaMenjaOcekivanje", Err.Number, Err.description
End Sub

' Header ne nosi nista sto je stavka.
Private Sub Test_OTP_HeaderNeNosiLinePolja()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPHP")

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add OtpNoviOtkup(scenario, 400#, 600#)

    Dim otpID As String
    otpID = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-HP-" & scenario), _
                                        izvori)

    AssertTrue Len(otpID) > 0, "OTP header: upis prosao"
    AssertEquals "", OtpPolje(otpID, COL_OTP_KOLICINA), "OTP header: Kolicina prazna"
    AssertEquals "", OtpPolje(otpID, COL_OTP_KOL_AMB), "OTP header: KolAmbalaze prazna"
    AssertEquals "", OtpPolje(otpID, COL_OTP_KLASA), "OTP header: Klasa prazna"
    AssertEquals "", OtpPolje(otpID, COL_OTP_BRUTO), "OTP header: BrutoKg prazan"

    ' A vozac JESTE na headeru -- otpremnica ga poseduje (S4.1c).
    AssertEquals TEST_VOZ_ID, OtpPolje(otpID, COL_OTP_VOZAC), "OTP header: VozacID"

    Exit Sub

EH:
    LogFatal "Test_OTP_HeaderNeNosiLinePolja", Err.Number, Err.description
End Sub

' Zatvoren spisak kljuceva: snapshot i izvedena polja se ne primaju.
Private Sub Test_OTP_NepoznatKljucUHeaderuPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPNK")

    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-NK-" & scenario)
    h.Add "VrstaVoca", TEST_VRSTA

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaDraft_TX(h, OtpOcek(400#, 20#, 0#, 0#), razlog)

    AssertEquals "", rez, "OTP nepoznat kljuc: upis odbijen"
    AssertTrue InStr(1, razlog, "nepoznat kljuc", vbTextCompare) > 0 And _
               InStr(1, razlog, "VrstaVoca", vbTextCompare) > 0, _
               "OTP nepoznat kljuc: kapija imenuje kljuc (bilo: " & razlog & ")"

    ' Isto vazi i za ocekivanu stavku -- tipfeler u opcionom polju je nevidljiv.
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", KLASA_I
    s.Add "Kolicina", 400#
    s.Add "KolAmbalaze", 20#
    s.Add "BrutoKg", 420#

    Dim c As Collection
    Set c = New Collection
    c.Add s

    rez = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-NK2-" & scenario), c, razlog)
    AssertEquals "", rez, "OTP nepoznat kljuc: stavka odbijena"
    AssertTrue InStr(1, razlog, "BrutoKg", vbTextCompare) > 0, _
               "OTP nepoznat kljuc: stavka imenuje kljuc (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_NepoznatKljucUHeaderuPada", Err.Number, Err.description
End Sub

' FK-ovi headera: stanica, vozac i kultura moraju postojati.
Private Sub Test_OTP_HeaderFKovi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPFK")

    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-FK-" & scenario)
    h("StanicaID") = "ST-NE-POSTOJI-" & scenario

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaDraft_TX(h, OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertEquals "", rez, "OTP FK: nepostojeca stanica odbijena"
    AssertTrue InStr(1, razlog, "StanicaID ne postoji", vbTextCompare) > 0, _
               "OTP FK: stanica imenovana (bilo: " & razlog & ")"

    Dim h2 As Object
    Set h2 = OtpHeader(TEST_PREFIX & "-OTP-FK2-" & scenario)
    h2("VozacID") = "VOZ-NE-POSTOJI-" & scenario

    rez = CreateOtpremnicaDraft_TX(h2, OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertEquals "", rez, "OTP FK: nepostojeci vozac odbijen"
    AssertTrue InStr(1, razlog, "VozacID ne postoji", vbTextCompare) > 0, _
               "OTP FK: vozac imenovan (bilo: " & razlog & ")"

    Dim h3 As Object
    Set h3 = OtpHeader(TEST_PREFIX & "-OTP-FK3-" & scenario)
    h3("KulturaID") = TEST_VRSTA & "-" & TEST_SORTA     ' oblik koji stari kod fabrikuje

    rez = CreateOtpremnicaDraft_TX(h3, OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertEquals "", rez, "OTP FK: fabrikovana kultura odbijena"
    AssertTrue InStr(1, razlog, "KulturaID ne postoji", vbTextCompare) > 0, _
               "OTP FK: kultura imenovana (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_HeaderFKovi", Err.Number, Err.description
End Sub

' Clanstvo je promenljivo DOK je otpremnica DRAFT (A15).
Private Sub Test_OTP_ClanstvoMutabilnoUDraftu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPMU")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-MU-" & scenario), _
                                     OtpOcek(500#, 40#, 0#, 0#))

    Dim otkA As String, otkB As String
    otkA = OtpNoviOtkup(scenario & "A", 400#, 0#)
    otkB = OtpNoviOtkup(scenario & "B", 100#, 0#)

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, otkA, razlog), "OTP mutacija: A dodat"
    AssertTrue DodajOtpremnicaIzvor_TX(otpID, otkB, razlog), "OTP mutacija: B dodat"
    AssertEquals "2", CStr(OtpBrojClanova(otpID)), "OTP mutacija: dva clana"

    AssertTrue UkloniOtpremnicaIzvor_TX(otpID, otkA, razlog), _
               "OTP mutacija: A uklonjen (bilo: " & razlog & ")"
    AssertEquals "1", CStr(OtpBrojClanova(otpID)), "OTP mutacija: ostao jedan clan"

    ' Uklonjen izvor je SLOBODAN -- moze u drugu otpremnicu. Da je ostao
    ' tombstone, kapija "vec u aktivnoj otpremnici" bi ga i dalje drzala.
    Dim otpID2 As String
    otpID2 = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-MU2-" & scenario), _
                                      OtpOcek(400#, 20#, 0#, 0#))
    AssertTrue DodajOtpremnicaIzvor_TX(otpID2, otkA, razlog), _
               "OTP mutacija: uklonjen izvor je slobodan (bilo: " & razlog & ")"

    AssertTrue UkloniOtpremnicaIzvor_TX(otpID2, otkA, razlog), "OTP mutacija: A opet uklonjen"
    AssertTrue DodajOtpremnicaIzvor_TX(otpID, otkA, razlog), "OTP mutacija: A vracen"
    AssertEquals "2", CStr(OtpBrojClanova(otpID)), "OTP mutacija: opet dva clana"

    Exit Sub

EH:
    LogFatal "Test_OTP_ClanstvoMutabilnoUDraftu", Err.Number, Err.description
End Sub

' Posle izdavanja je sastav istorijska cinjenica (A13).
Private Sub Test_OTP_PosleIzdavanjaClanstvoZamrznuto()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPZM")

    Dim otpID As String, razlog As String
    Dim otkA As String, otkB As String
    otkA = OtpNoviOtkup(scenario & "A", 400#, 0#)
    otkB = OtpNoviOtkup(scenario & "B", 100#, 0#)

    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-ZM-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID, otkA
    AssertTrue IzdajOtpremnicu_TX(otpID, razlog), "OTP zamrznuto: izdato (bilo: " & razlog & ")"

    AssertTrue Not DodajOtpremnicaIzvor_TX(otpID, otkB, razlog), _
               "OTP zamrznuto: dodavanje odbijeno"
    AssertTrue InStr(1, razlog, "nije DRAFT", vbTextCompare) > 0, _
               "OTP zamrznuto: kapija imenuje razlog (bilo: " & razlog & ")"

    AssertTrue Not UkloniOtpremnicaIzvor_TX(otpID, otkA, razlog), _
               "OTP zamrznuto: uklanjanje odbijeno"

    AssertTrue Not UpdateOtpremnicaDraft_TX(otpID, _
                       OtpHeader(TEST_PREFIX & "-OTP-ZM-" & scenario), _
                       OtpOcek(999#, 20#, 0#, 0#), razlog), _
               "OTP zamrznuto: izmena drafta odbijena"

    AssertEquals "1", CStr(OtpBrojClanova(otpID)), "OTP zamrznuto: sastav netaknut"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_I, COL_OPS_KOLICINA) - 400#) < 0.001, _
               "OTP zamrznuto: stavka netaknuta"

    ' Ni ponovno izdavanje: to bi napravilo drugi komplet stavki.
    AssertTrue Not IzdajOtpremnicu_TX(otpID, razlog), "OTP zamrznuto: reizdavanje odbijeno"
    AssertEquals "1", CStr(OtpBrojStavki(otpID)), "OTP zamrznuto: jedan komplet stavki"

    Exit Sub

EH:
    LogFatal "Test_OTP_PosleIzdavanjaClanstvoZamrznuto", Err.Number, Err.description
End Sub

' Jedan otkup ne sme da bude u dve aktivne otpremnice.
Private Sub Test_OTP_IzvorNeSmeDvaPutaAktivno()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPDV")

    Dim otkID As String
    otkID = OtpNoviOtkup(scenario, 400#, 0#)

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add otkID

    AssertTrue Len(CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-DV-" & scenario), _
                                               izvori)) > 0, _
               "OTP dvaput: prva otpremnica prosla"

    Dim izvori2 As Collection
    Set izvori2 = New Collection
    izvori2.Add otkID

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-DV2-" & scenario), _
                                      izvori2, razlog)

    AssertEquals "", rez, "OTP dvaput: druga otpremnica odbijena"
    AssertTrue InStr(1, razlog, "vec u sastavu aktivne otpremnice", vbTextCompare) > 0, _
               "OTP dvaput: kapija imenuje razlog (bilo: " & razlog & ")"

    ' Isti otkup dvaput u ISTOM pozivu je isto greska.
    Dim izvori3 As Collection
    Set izvori3 = New Collection
    Dim otkC As String: otkC = OtpNoviOtkup(scenario & "C", 400#, 0#)
    izvori3.Add otkC
    izvori3.Add otkC

    rez = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-DV3-" & scenario), _
                                      izvori3, razlog)
    AssertEquals "", rez, "OTP dvaput: isti izvor dvaput u istom pozivu odbijen"

    Exit Sub

EH:
    LogFatal "Test_OTP_IzvorNeSmeDvaPutaAktivno", Err.Number, Err.description
End Sub

' Storniran otkup nije roba.
Private Sub Test_OTP_StorniranIzvorNeUlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPST")

    Dim otkID As String
    otkID = OtpNoviOtkup(scenario, 400#, 0#)

    RequireUpdateCell TBL_OTKUP, FindRows(TBL_OTKUP, COL_OTK_ID, otkID)(1), _
                      COL_STORNIRANO, "Da", "Test_OTP_StorniranIzvorNeUlazi"

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add otkID

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-ST-" & scenario), _
                                      izvori, razlog)

    AssertEquals "", rez, "OTP storno: storniran izvor odbijen"
    AssertTrue InStr(1, razlog, "storniran", vbTextCompare) > 0, _
               "OTP storno: kapija imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_StorniranIzvorNeUlazi", Err.Number, Err.description
End Sub

' Otpremnica je isporuka sa JEDNOG otkupnog mesta.
Private Sub Test_OTP_DveStaniceNeProlaze()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPS2")

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add OtpNoviOtkup(scenario & "A", 400#, 0#)
    izvori.Add OtpNoviOtkupNaStanici(scenario & "B", TEST_HLAD_ST_ID)

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-S2-" & scenario), _
                                      izvori, razlog)

    AssertEquals "", rez, "OTP dve stanice: upis odbijen"
    AssertTrue InStr(1, razlog, "StanicaID", vbTextCompare) > 0, _
               "OTP dve stanice: kapija imenuje polje (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_DveStaniceNeProlaze", Err.Number, Err.description
End Sub

' Otpremnica bez otkupa nije isporuka -- ali prazan DRAFT sa ocekivanjem jeste.
Private Sub Test_OTP_BezIzvoraNeProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPBI")

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTPREMNICA)

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-BI-" & scenario), _
                                      New Collection, razlog)

    AssertEquals "", rez, "OTP bez izvora: jednopotezni upis odbijen"
    AssertTrue InStr(1, razlog, "nema nijedan izvor", vbTextCompare) > 0, _
               "OTP bez izvora: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTPREMNICA)), _
                 "OTP bez izvora: header nije ostao"

    ' DRAFT bez ijednog izvora je legitiman -- to je bas ono sto panel pravi.
    Dim otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-BI2-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))
    AssertTrue Len(otpID) > 0, "OTP bez izvora: DRAFT sa ocekivanjem je legitiman"

    ' Ali izdavanje bez ijednog izvora nije.
    AssertTrue Not IzdajOtpremnicu_TX(otpID, razlog), "OTP bez izvora: izdavanje odbijeno"
    AssertTrue InStr(1, razlog, "nema nijedan izvor", vbTextCompare) > 0, _
               "OTP bez izvora: izdavanje imenuje razlog (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_BezIzvoraNeProlazi", Err.Number, Err.description
End Sub

' Otkup po STAROM modelu nema stavke, pa ne moze u kanonsku otpremnicu.
Private Sub Test_OTP_StariOtkupNeUlazi()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPSM")

    ' Zaglavlje bez stavki se od review-a #334 ne sme ostaviti u svesci:
    ' oborilo bi svakog sledeceg citaoca vrednosti u suite-u.
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    ' CreateOtkup_TX zove ApplyAvansToOtkup: fixture sme da potrosi ili
    ' podeli slobodan avans kooperanta, pa i tblNovac mora nazad.
    tx.AddTableSnapshot TBL_NOVAC

    Dim stariID As String
    stariID = OtkupBezStavkiFixture(TEST_PREFIX & "-OTK-SM-" & scenario)

    AssertTrue Len(stariID) > 0, "OTP stari otkup: zaglavlje bez stavki napravljeno"
    AssertEquals "0", CStr(OtkBrojStavkiZaOtkup(stariID)), _
                 "OTP stari otkup: nema stavki (inace test ne meri nista)"

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add stariID

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-SM-" & scenario), _
                                      izvori, razlog)

    AssertEquals "", rez, "OTP stari otkup: upis odbijen"
    AssertTrue InStr(1, razlog, "nema stavke", vbTextCompare) > 0, _
               "OTP stari otkup: kapija imenuje razlog (bilo: " & razlog & ")"

    tx.RollbackTx
    Set tx = Nothing
    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFatal "Test_OTP_StariOtkupNeUlazi", errNum, errDesc
End Sub

' Prazan ID je fail-closed, i header ne sme da ostane bez stavki.
Private Sub Test_OTP_PrazanIDFailClosed()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPFC")

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add OtpNoviOtkup(scenario, 400#, 600#)

    Dim preH As Long, preS As Long
    preH = OtkBrojRedova(TBL_OTPREMNICA)
    preS = OtkBrojRedova(TBL_OTPREMNICA_STAVKE)

    Dim rez As String, razlog As String
    Dim prevMode As Boolean
    prevMode = IsTestMode()
    SetTestMode True

    ' header (1) + clanstvo (2) prodju, STAVKA pada -- ID-evi idu tim redom
    modDataAccess.NewEntityIDPadniTest True, 2
    rez = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-FC-" & scenario), _
                                      izvori, razlog)
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode

    AssertEquals "", rez, "OTP prazan ID: upis odbijen"
    AssertTrue InStr(1, razlog, "nije vratio OtpremnicaStavkaID", vbTextCompare) > 0, _
               "OTP prazan ID: kapija imenuje KOJI id (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTPREMNICA)), _
                 "OTP prazan ID: header nije ostao"
    AssertEquals CStr(preS), CStr(OtkBrojRedova(TBL_OTPREMNICA_STAVKE)), _
                 "OTP prazan ID: stavka nije ostala"

    Exit Sub

EH:
    modDataAccess.NewEntityIDPadniTest False
    SetTestMode prevMode
    LogFatal "Test_OTP_PrazanIDFailClosed", Err.Number, Err.description
End Sub

' Skela je ADITIVNA: stara kolona ostaje netaknuta.
'
' 39 ne-test citalaca u 15 modula jos zivi na Otkup.OtpremnicaID. Bez ove
' tvrdnje bi "aditivno" bila namera, ne mereno svojstvo.
Private Sub Test_OTP_OtkupOtpremnicaIDNetaknut()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPNT")

    Dim otkID As String
    otkID = OtpNoviOtkup(scenario, 400#, 0#)

    AssertEquals "", Trim$(CStr(nz(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, _
                                                 COL_OTK_OTPREMNICA_ID), ""))), _
                 "OTP netaknuto: pre otpremnice je prazno"

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add otkID

    Dim otpID As String
    otpID = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-NT-" & scenario), _
                                        izvori)
    AssertTrue Len(otpID) > 0, "OTP netaknuto: otpremnica nastala"

    AssertEquals "", Trim$(CStr(nz(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, _
                                                 COL_OTK_OTPREMNICA_ID), ""))), _
                 "OTP netaknuto: nov pisac NE pise staru kolonu"
    AssertEquals otkID, CStr(OtpClanoviTest(otpID)(1)), _
                 "OTP netaknuto: pripadnost zna clanstvo, ne kolona"

    Exit Sub

EH:
    LogFatal "Test_OTP_OtkupOtpremnicaIDNetaknut", Err.Number, Err.description
End Sub

' Nothing je PRIVATAN signal jednopoteznog puta -- javni rucni ulaz ga ne prima.
'
' Bez ove kapije je CreateOtpremnicaDraft_TX(h, Nothing) pravio validan DRAFT BEZ
' ocekivanja: dokument koji nema sta da meri, a izgleda ispravno. Prazna
' Collection je NESTO DRUGO (i vec je odbijena) -- ovo je bas Nothing.
Private Sub Test_OTP_DraftNothingOcekivanjePada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPNO")

    Dim preH As Long, preS As Long
    preH = OtkBrojRedova(TBL_OTPREMNICA)
    preS = OtkBrojRedova(TBL_OTPREMNICA_STAVKE)

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-NO-" & scenario), _
                                   Nothing, razlog)

    AssertEquals "", rez, "OTP Nothing: upis odbijen"
    AssertTrue InStr(1, razlog, "Ocekivanje nije prosledjeno", vbTextCompare) > 0, _
               "OTP Nothing: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTPREMNICA)), _
                 "OTP Nothing: header nije ostao"
    AssertEquals CStr(preS), CStr(OtkBrojRedova(TBL_OTPREMNICA_STAVKE)), _
                 "OTP Nothing: stavka nije ostala"

    Exit Sub

EH:
    LogFatal "Test_OTP_DraftNothingOcekivanjePada", Err.Number, Err.description
End Sub

' Izmena zaglavlja ne sme da pokvari vec validno clanstvo.
'
' Draft sa stanicom ST1 i clanom sa ST1, prebacen na ST2, nosio bi clana koga
' Dodaj nikad ne bi primio -- a GetOtpremnicaProgress bi ga do izdavanja uredno
' racunao. Invarijanta ne sme da bude prekrsena izmedju dva klika.
Private Sub Test_OTP_UpdateStaniceSaPostojecimIzvoromPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPUS")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-US-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID, OtpNoviOtkup(scenario, 400#, 0#)

    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-US-" & scenario)
    h("StanicaID") = TEST_HLAD_ST_ID

    AssertTrue Not UpdateOtpremnicaDraft_TX(otpID, h, OtpOcek(999#, 30#, 0#, 0#), razlog), _
               "OTP update stanice: odbijen zbog postojeceg izvora"
    AssertTrue InStr(1, razlog, "StanicaID", vbTextCompare) > 0, _
               "OTP update stanice: kapija imenuje polje (bilo: " & razlog & ")"

    ' Rollback mora da vrati I zaglavlje I ocekivanje.
    AssertEquals TEST_ST_ID, OtpPolje(otpID, COL_OTP_STANICA), _
                 "OTP update stanice: stara stanica netaknuta"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_I, COL_OPS_KOLICINA) - 400#) < 0.001, _
               "OTP update stanice: staro ocekivanje netaknuto"
    AssertEquals "1", CStr(OtpBrojStavki(otpID)), "OTP update stanice: jedna stavka"
    AssertEquals "1", CStr(OtpBrojClanova(otpID)), "OTP update stanice: clan netaknut"

    Exit Sub

EH:
    LogFatal "Test_OTP_UpdateStaniceSaPostojecimIzvoromPada", Err.Number, Err.description
End Sub

Private Sub Test_OTP_UpdateKultureSaPostojecimIzvoromPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPUK")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-UK-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID, OtpNoviOtkup(scenario, 400#, 0#)

    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-UK-" & scenario)
    h("KulturaID") = TEST_KUL_BEZ_SORTE_ID

    AssertTrue Not UpdateOtpremnicaDraft_TX(otpID, h, OtpOcek(400#, 20#, 0#, 0#), razlog), _
               "OTP update kulture: odbijen zbog postojeceg izvora"
    AssertTrue InStr(1, razlog, "KulturaID", vbTextCompare) > 0, _
               "OTP update kulture: kapija imenuje polje (bilo: " & razlog & ")"

    AssertEquals TEST_KULTURA_ID, OtpPolje(otpID, COL_OTP_KULTURA), _
                 "OTP update kulture: stara kultura netaknuta"
    AssertEquals TEST_VRSTA, OtpPolje(otpID, COL_OTP_VRSTA), _
                 "OTP update kulture: stari snapshot vrste netaknut"

    ' Kontrola: izmena koja NE dira stanicu ni kulturu i dalje prolazi -- inace bi
    ' kapija koja sve odbija izgledala isto kao kapija koja radi.
    Dim h2 As Object
    Set h2 = OtpHeader(TEST_PREFIX & "-OTP-UK2-" & scenario)
    AssertTrue UpdateOtpremnicaDraft_TX(otpID, h2, OtpOcek(1000#, 50#, 0#, 0#), razlog), _
               "OTP update kulture: bezopasna izmena prolazi (bilo: " & razlog & ")"
    AssertTrue Abs(OtpStavkaBrojP(otpID, KLASA_I, COL_OPS_KOLICINA) - 1000#) < 0.001, _
               "OTP update kulture: ocekivanje promenjeno"

    Exit Sub

EH:
    LogFatal "Test_OTP_UpdateKultureSaPostojecimIzvoromPada", Err.Number, Err.description
End Sub

' Header nosi JEDAN TipAmbalaze, pa 20 plasticnih + 30 drvenih gajbi nije 50.
'
' Provera je pri DODAVANJU, ne tek pri izdavanju: do tada bi
' GetOtpremnicaProgress sabirao dve razlicite stvari i prikazivao broj koji
' semanticki ne znaci nista.
Private Sub Test_OTP_DvaTipaAmbalazeNeUlazeUDraft()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPTA")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-TA-" & scenario), _
                                     OtpOcek(800#, 40#, 0#, 0#))

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, OtpNoviOtkup(scenario & "A", 400#, 0#), razlog), _
               "OTP dva tipa: prvi izvor dodat"

    AssertTrue Not DodajOtpremnicaIzvor_TX(otpID, _
                       OtpNoviOtkupDrugogTipa(scenario & "B"), razlog), _
               "OTP dva tipa: drugi tip ambalaze odbijen VEC pri dodavanju"
    AssertTrue InStr(1, razlog, "TipAmbalaze", vbTextCompare) > 0, _
               "OTP dva tipa: kapija imenuje polje (bilo: " & razlog & ")"
    AssertEquals "1", CStr(OtpBrojClanova(otpID)), "OTP dva tipa: clanstvo netaknuto"

    ' Read-model nije stigao da sabere dve razlicite gajbe.
    Dim p As Object
    Set p = GetOtpremnicaProgress(otpID)
    AssertTrue Abs(p(UCase$(KLASA_I))("povezanoAmb") - 20#) < 0.001, _
               "OTP dva tipa: povezana ambalaza je samo prvog tipa"

    ' Kontrola: isti tip prolazi.
    AssertTrue DodajOtpremnicaIzvor_TX(otpID, OtpNoviOtkup(scenario & "C", 400#, 0#), razlog), _
               "OTP dva tipa: isti tip prolazi (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTP_DvaTipaAmbalazeNeUlazeUDraft", Err.Number, Err.description
End Sub

' Otpremnica se sastavlja od IZDATIH otkupnih listova.
'
' Danas svaki otkup iz CreateOtkup_TX jeste IZDATO, pa ovo nije ziv bug -- ali
' kanonska veza treba da kaze sta trazi, a ne da se oslanja na to sto drugi pisac
' trenutno ne pravi drugacije redove.
Private Sub Test_OTP_NeizdatOtkupNeUlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPNI")

    Dim otkID As String
    otkID = OtpNoviOtkup(scenario, 400#, 0#)

    AssertEquals IZDATO_IZDATO, _
                 Trim$(CStr(nz(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, _
                                             COL_TRACE_IZDATO_STATUS), ""))), _
                 "OTP neizdat: CreateOtkup_TX pise IZDATO (inace test ne meri nista)"

    RequireUpdateCell TBL_OTKUP, FindRows(TBL_OTKUP, COL_OTK_ID, otkID)(1), _
                      COL_TRACE_IZDATO_STATUS, IZDATO_DRAFT, "Test_OTP_NeizdatOtkupNeUlazi"

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-NI-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))

    AssertTrue Not DodajOtpremnicaIzvor_TX(otpID, otkID, razlog), _
               "OTP neizdat: neizdat otkup odbijen"
    AssertTrue InStr(1, razlog, "nije izdat", vbTextCompare) > 0, _
               "OTP neizdat: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals "0", CStr(OtpBrojClanova(otpID)), "OTP neizdat: clanstvo prazno"

    Exit Sub

EH:
    LogFatal "Test_OTP_NeizdatOtkupNeUlazi", Err.Number, Err.description
End Sub

' TipAmbalaze je HEADER cinjenica, primljena pri otvaranju.
'
' Ocekivanje "50 gajbi" mora da zna KOJIH 50 vec pri otvaranju -- inace se tip
' saznaje tek iz prvog izvora. Zatecen posao to vec resava ovako: legacy
' SaveOtpremnicaMulti_TX prima tipAmb JEDNOM, kao header podatak, i bas njime
' knjizi ambalazu pri nastanku otpremnice (modDokumenta:382 TrackAmbalaza).
Private Sub Test_OTP_TipAmbalazeJeHeaderCinjenica()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPTH")

    ' Ocekuje se ambalaza, a tip nije rekao KOJA.
    Dim h As Object
    Set h = OtpHeader(TEST_PREFIX & "-OTP-TH-" & scenario)
    h("TipAmbalaze") = ""

    Dim rez As String, razlog As String
    rez = CreateOtpremnicaDraft_TX(h, OtpOcek(400#, 20#, 0#, 0#), razlog)

    AssertEquals "", rez, "OTP tip header: ocekivana ambalaza bez tipa odbijena"
    AssertTrue InStr(1, razlog, "Tip ambalaze je obavezan", vbTextCompare) > 0, _
               "OTP tip header: kapija imenuje razlog (bilo: " & razlog & ")"

    ' Bez ocekivane ambalaze prazan tip je tacan podatak.
    Dim h2 As Object
    Set h2 = OtpHeader(TEST_PREFIX & "-OTP-TH2-" & scenario)
    h2("TipAmbalaze") = ""

    AssertTrue Len(CreateOtpremnicaDraft_TX(h2, OtpOcek(400#, 0#, 0#, 0#), razlog)) > 0, _
               "OTP tip header: bez ambalaze prazan tip prolazi (bilo: " & razlog & ")"

    ' Sa tipom: DRAFT ga nosi ODMAH, ne tek posle prvog izvora.
    Dim otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-TH3-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#), razlog)
    AssertTrue Len(otpID) > 0, "OTP tip header: sa tipom prolazi (bilo: " & razlog & ")"
    AssertEquals TEST_TIP_AMB, OtpPolje(otpID, COL_OTP_TIP_AMB), _
                 "OTP tip header: DRAFT nosi tip pre ijednog izvora"

    Exit Sub

EH:
    LogFatal "Test_OTP_TipAmbalazeJeHeaderCinjenica", Err.Number, Err.description
End Sub

' Izvor koji NE nosi gajbe ne odredjuje transportnu ambalazu.
'
' Otkup sme da ima TipAmbalaze zbog KolAmbIzdata -- gajbi koje su OTISLE
' kooperantu -- a da njegove stavke ne nose nijednu gajbu u otpremnicu.
' Poredjenje golih header stringova svih otkupa bi takav izvor pogresno odbilo.
Private Sub Test_OTP_IzvorBezGajbiNeOdredjujeTip()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPBG")

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-BG-" & scenario), _
                                     OtpOcek(800#, 20#, 0#, 0#), razlog)

    ' Nosi gajbe i slaze se sa headerom.
    AssertTrue DodajOtpremnicaIzvor_TX(otpID, OtpNoviOtkup(scenario & "A", 400#, 0#), razlog), _
               "OTP bez gajbi: izvor sa gajbama dodat"

    ' DRUGI TIP na headeru otkupa, ali NULA gajbi na stavkama -- prolazi.
    Dim otkB As String
    otkB = OtpNoviOtkupDrugogTipaBezGajbi(scenario & "B")

    AssertTrue Abs(OtkStavkaBrojP(otkB, KLASA_I, COL_OKS_KOL_AMB)) < 0.001, _
               "OTP bez gajbi: taj otkup zaista ne nosi gajbe (inace test ne meri nista)"
    AssertEquals TEST_TIP_AMB_B, _
                 Trim$(CStr(nz(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkB, _
                                             COL_OTK_TIP_AMB), ""))), _
                 "OTP bez gajbi: a header otkupa nosi DRUGI tip"

    AssertTrue DodajOtpremnicaIzvor_TX(otpID, otkB, razlog), _
               "OTP bez gajbi: izvor bez gajbi prolazi uprkos drugom tipu (bilo: " & razlog & ")"
    AssertEquals "2", CStr(OtpBrojClanova(otpID)), "OTP bez gajbi: dva clana"

    Exit Sub

EH:
    LogFatal "Test_OTP_IzvorBezGajbiNeOdredjujeTip", Err.Number, Err.description
End Sub

' Clanstvo na nepostojeci otkup je INTEGRITET, ne manji zbir.
'
' Citalac koji tiho izracuna manje pokazuje operateru broj koji izgleda ispravno,
' a finalizacija istu korupciju prijavi tek sat kasnije.
Private Sub Test_OTP_ClanstvoNaNepostojeciOtkupPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPCN")

    Dim otpID As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-CN-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))

    Dim red As Long
    red = OtpUpisiSirovoClanstvo(otpID, "OTK-NE-POSTOJI-" & scenario)

    Dim greska As String
    greska = OtpProgressGreska(otpID)

    ' CISCENJE PRE TVRDNJI: korumpiran red truje AktivnoOtpClanstvoPoKanonu za
    ' svaki sledeci test, pa se sklanja pre nego sto bilo sta moze da padne.
    DeleteRow TBL_OTPREMNICA_IZVORI, red

    AssertTrue InStr(1, greska, "ne postoji", vbTextCompare) > 0, _
               "OTP clanstvo: read-model pada na nepostojeci otkup (bilo: " & greska & ")"
    AssertTrue InStr(1, greska, "clanstvo", vbTextCompare) > 0, _
               "OTP clanstvo: poruka kaze da je rec o clanstvu (bilo: " & greska & ")"

    ' Posle ciscenja read-model opet radi -- inace bi test dokazao samo da nesto puca.
    AssertEquals "", OtpProgressGreska(otpID), "OTP clanstvo: posle ciscenja read-model radi"

    Exit Sub

EH:
    LogFatal "Test_OTP_ClanstvoNaNepostojeciOtkupPada", Err.Number, Err.description
End Sub

' Isti par (otpremnica, otkup) dvaput je korupcija, ne "jedan clan".
Private Sub Test_OTP_DupliParUClanstvuPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPDP")

    Dim otkID As String
    otkID = OtpNoviOtkup(scenario, 400#, 0#)

    Dim otpID As String, razlog As String
    otpID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-DP-" & scenario), _
                                     OtpOcek(400#, 20#, 0#, 0#))
    DodajOtpremnicaIzvor_TX otpID, otkID

    Dim red As Long
    red = OtpUpisiSirovoClanstvo(otpID, otkID)

    Dim greska As String
    greska = OtpProgressGreska(otpID)

    DeleteRow TBL_OTPREMNICA_IZVORI, red

    AssertTrue InStr(1, greska, "postoji vise puta", vbTextCompare) > 0, _
               "OTP dupli par: read-model pada (bilo: " & greska & ")"
    AssertEquals "", OtpProgressGreska(otpID), "OTP dupli par: posle ciscenja read-model radi"
    AssertEquals "1", CStr(OtpBrojClanova(otpID)), "OTP dupli par: ostao jedan clan"

    Exit Sub

EH:
    LogFatal "Test_OTP_DupliParUClanstvuPada", Err.Number, Err.description
End Sub

' Ambalaza se knjizi JEDNOM po dokumentu, nad zbirom stavki.
'
' Zatecen pisac pravi par redova PO KLASI, samo zato sto ima dva OtkupID-a.
' tblAmbalaza nema kolonu Klasa, pa bi ta dva reda bila dva reda koja se
' razlikuju samo u kolicini. Sa jednim headerom taj razlog nestaje.
Private Sub Test_OTK_AmbalazaIdeNaDokument()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKAM")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-AM-" & scenario)
    h.Add "KolAmbIzdata", 7#

    Dim otkID As String
    otkID = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 600#, 40#, 30))

    AssertTrue Len(otkID) > 0, "OTK ambalaza: upis prosao"

    ' Primljene gajbe: 20 + 30 = 50, JEDAN par redova.
    AssertEquals "2", CStr(AmbBrojRedova(otkID, DOK_TIP_OTKUP)), _
                 "OTK ambalaza: primljeno = jedan dvojni upis, ne po klasi"
    AssertTrue Abs(AmbKolicina(otkID, DOK_TIP_OTKUP, "Izlaz") - 50#) < 0.001, _
               "OTK ambalaza: kooperant IZLAZ nosi zbir stavki"
    AssertTrue Abs(AmbKolicina(otkID, DOK_TIP_OTKUP, "Ulaz") - 50#) < 0.001, _
               "OTK ambalaza: OM ULAZ nosi isti zbir"
    AssertEquals TEST_KOOP_ID, AmbEntitet(otkID, DOK_TIP_OTKUP, "Izlaz"), _
                 "OTK ambalaza: izlazna noga je kooperantova"
    AssertEquals TEST_ST_ID, AmbEntitet(otkID, DOK_TIP_OTKUP, "Ulaz"), _
                 "OTK ambalaza: ulazna noga je stanicina"

    ' Izdate gajbe: obrnut smer, svoj tip dokumenta.
    AssertEquals "2", CStr(AmbBrojRedova(otkID, DOK_TIP_OM_IZLAZ_KOOP)), _
                 "OTK ambalaza: izdato = jedan dvojni upis"
    AssertTrue Abs(AmbKolicina(otkID, DOK_TIP_OM_IZLAZ_KOOP, "Ulaz") - 7#) < 0.001, _
               "OTK ambalaza: kooperant ULAZ prima prazne"

    ' Vozac se NE zigose -- otkup ga u ciljnom modelu nema.
    AssertEquals "", AmbVozac(otkID, DOK_TIP_OTKUP, "Izlaz"), _
                 "OTK ambalaza: nema vozaca na otkupnoj nozi"

    ' Bez gajbi nema ni reda -- prazan upis nije nula, nego odsustvo.
    Dim otkID2 As String
    otkID2 = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-AM2-" & scenario), _
                            OtkStavke(400#, 50#, 0, 0#, 0#, 0))
    AssertEquals "0", CStr(AmbBrojRedova(otkID2, DOK_TIP_OTKUP)), _
                 "OTK ambalaza: bez gajbi nema reda"

    Exit Sub

EH:
    LogFatal "Test_OTK_AmbalazaIdeNaDokument", Err.Number, Err.description
End Sub

' Odbijen dokument ne knjizi ambalazu.
'
' IME JE ISPRAVLJENO POSLE SABOTAZE. Prvo se zvao "AmbalazaUIstojTransakciji" i
' tvrdio da rollback vraca redove ambalaze -- ali sabotaza koja SKLONI
' AddTableSnapshot TBL_AMBALAZA nije ugrizla. Razlog: pad je u prevalidaciji, PRE
' ijednog knjizenja, pa nema sta ni da se vrati. Test je merio odsustvo upisa, a
' tvrdio rollback.
'
' Ono sto sada tvrdi je i dalje vredno: knjizenje se ne sme pomeriti ISPRED
' validacije. Pad IZMEDJU dva TrackAmbalaza poziva snapshot stvarno pokriva, ali
' se iz javnog API-ja ne moze izazvati bez novog seam-a -- to je namerno
' neizmereno, a ne previdjeno.
Private Sub Test_OTK_OdbijenDokumentNeKnjiziAmbalazu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKAT")

    Dim preA As Long
    preA = OtkBrojRedova(TBL_AMBALAZA)

    ' Druga stavka ne valja -> ceo dokument pada, pa i ambalaza prve.
    Dim stavke As Collection
    Set stavke = New Collection
    stavke.Add OtkStavka(KLASA_I, 400#, 50#, 20#, 0#)
    stavke.Add OtkStavka(KLASA_II, 600#, 0#, 30#, 0#)      ' cena 0 -> odbijeno

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-AT-" & scenario), stavke, razlog)

    AssertEquals "", rez, "OTK ambalaza tx: upis odbijen"
    AssertTrue InStr(1, razlog, "Cena", vbTextCompare) > 0, _
               "OTK ambalaza tx: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertEquals CStr(preA), CStr(OtkBrojRedova(TBL_AMBALAZA)), _
                 "OTK ambalaza: odbijen dokument nije knjizio nijedan red"

    Exit Sub

EH:
    LogFatal "Test_OTK_OdbijenDokumentNeKnjiziAmbalazu", Err.Number, Err.description
End Sub

' EKRAN PISE NOVIM MODELOM.
'
' OtkupUpisi je jedini produkcioni put do pisca, a zove ga samo ekran
' (modScrDokumenti:778) -- do sada ga nijedan test nije izvrsavao. Prelazak sa
' SaveOtkupMulti_TX na CreateOtkup_TX je najveca izmena ponasanja u cutover-u,
' pa bez ovog testa zelena suite ne bi dokazivala nista o njoj.
Private Sub Test_OTK_EkranPiseNovimModelom()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKEK")

    Dim p As Object
    Set p = OtkEkranParam(TEST_PREFIX & "-OTK-EK-" & scenario)
    p("kolicinaI") = 400#
    p("cenaI") = 50#
    p("kolAmb") = 20&
    p("dveKlase") = True
    p("kolicinaII") = 600#
    p("cenaII") = 40#
    p("kolAmbII") = 30&

    Dim poruke As String
    Dim res As String
    res = modOtkupUnos.OtkupUpisi(p, poruke)

    AssertTrue Len(res) > 0, "OTK ekran: upis prosao (poruke: " & poruke & ")"

    ' JEDAN ID -- ne "ID1 + ID2".
    AssertEquals "0", CStr(InStr(1, res, " + ")), "OTK ekran: vraca JEDAN ID, bez spajanja"
    AssertEquals "1", CStr(FindRows(TBL_OTKUP, COL_OTK_ID, res).count), _
                 "OTK ekran: tacno jedan header red za dvoklasni blok"

    ' Stavke nose brojeve, header ih ne nosi.
    AssertEquals "2", CStr(OtkBrojStavkiZaOtkup(res)), "OTK ekran: dve stavke"
    AssertTrue Abs(OtkStavkaBrojP(res, KLASA_I, COL_OKS_KOLICINA) - 400#) < 0.001, _
               "OTK ekran: Klasa I kolicina sa ekrana"
    AssertTrue Abs(OtkStavkaBrojP(res, KLASA_II, COL_OKS_CENA) - 40#) < 0.001, _
               "OTK ekran: Klasa II cena sa ekrana"

    ' Kultura je razresena iz (vrsta, sorta) -- ekran je adapter, ne pisac.
    AssertEquals TEST_KULTURA_ID, OtkPolje(res, COL_OTK_KULTURA), _
                 "OTK ekran: KulturaID razresen iz izbora"

    ' Vozac se vise ne pise na otkup.
    AssertEquals "", OtkPolje(res, COL_OTK_VOZAC), "OTK ekran: header ne nosi vozaca"

    ' Ambalaza: jedan dvojni upis nad zbirom (20 + 30).
    AssertEquals "2", CStr(AmbBrojRedova(res, DOK_TIP_OTKUP)), "OTK ekran: jedan dvojni upis"
    AssertTrue Abs(AmbKolicina(res, DOK_TIP_OTKUP, "Izlaz") - 50#) < 0.001, _
               "OTK ekran: knjizen zbir gajbi obe klase"

    Exit Sub

EH:
    LogFatal "Test_OTK_EkranPiseNovimModelom", Err.Number, Err.description
End Sub

' Nerazresiva kultura obara upis PRE pisca, sa porukom operateru.
Private Sub Test_OTK_EkranNerazresivaKulturaPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKEN")

    Dim p As Object
    Set p = OtkEkranParam(TEST_PREFIX & "-OTK-EN-" & scenario)
    p("sorta") = TEST_SORTA & " NEPOSTOJECA"
    p("kolicinaI") = 400#
    p("cenaI") = 50#
    p("kolAmb") = 20&

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim poruke As String
    Dim res As String
    res = modOtkupUnos.OtkupUpisi(p, poruke)

    AssertEquals "", res, "OTK ekran kultura: upis odbijen"
    AssertTrue Len(poruke) > 0, "OTK ekran kultura: operater je dobio poruku"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK ekran kultura: header nije ostao"

    Exit Sub

EH:
    LogFatal "Test_OTK_EkranNerazresivaKulturaPada", Err.Number, Err.description
End Sub

' Auto-lanac hladnjace je PAUZIRAN do PR7, i to se kaze operateru.
'
' Lanac deli dokument po klasi, a nov pisac daje jedan OtkupID -- veza bi bila
' polovicna. Kod lanca ostaje netaknut; pauzira se poziv.
Private Sub Test_OTK_EkranPauziraAutoLanac()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPL")

    Dim p As Object
    Set p = OtkEkranParam(TEST_PREFIX & "-OTK-PL-" & scenario)
    p("stanicaID") = TEST_HLAD_ST_ID
    p("kolicinaI") = 400#
    p("cenaI") = 50#
    p("kolAmb") = 20&

    Dim preOtp As Long
    preOtp = OtkBrojRedova(TBL_OTPREMNICA)

    Dim poruke As String
    Dim res As String
    res = modOtkupUnos.OtkupUpisi(p, poruke)

    AssertTrue Len(res) > 0, "OTK lanac: otkup je upisan (poruke: " & poruke & ")"
    AssertTrue InStr(1, poruke, "PAUZIRAN", vbTextCompare) > 0, _
               "OTK lanac: operater je obavesten (bilo: " & poruke & ")"
    AssertEquals CStr(preOtp), CStr(OtkBrojRedova(TBL_OTPREMNICA)), _
                 "OTK lanac: nijedna otpremnica nije nastala"

    ' Kontrola: van hladnjace nema ni poruke -- inace bi se javljala uvek.
    Dim p2 As Object
    Set p2 = OtkEkranParam(TEST_PREFIX & "-OTK-PL2-" & scenario)
    p2("kolicinaI") = 400#
    p2("cenaI") = 50#
    p2("kolAmb") = 20&

    Dim poruke2 As String
    AssertTrue Len(modOtkupUnos.OtkupUpisi(p2, poruke2)) > 0, "OTK lanac: obican unos prosao"
    AssertEquals "0", CStr(InStr(1, poruke2, "PAUZIRAN", vbTextCompare)), _
                 "OTK lanac: van hladnjace nema poruke"

    Exit Sub

EH:
    LogFatal "Test_OTK_EkranPauziraAutoLanac", Err.Number, Err.description
End Sub

' PWA INGEST IDE KROZ KANONSKI PISAC.
'
' Zatecen uvoz je radio AppendRow(TBL_OTKUP) sa golim Array-em od 24 elementa nad
' tabelom od 39 kolona, fabrikovao KulturaID i NIJE pravio stavke. Bio je drugi
' put do istog dokumenta -- i drugi model.
Private Sub Test_PWA_IngestPraviHeaderIStavku()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PWAIN")

    Dim crid As String
    crid = TEST_PREFIX & "-CRID-" & scenario

    Dim red As Variant
    red = PwaRed(crid, TEST_PREFIX & "-OTK-PWA-" & scenario, 400#, 50#, 20)

    Dim otkID As String
    otkID = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)

    AssertTrue Len(otkID) > 0, "PWA: uvoz vratio OtkupID"
    AssertEquals "1", CStr(FindRows(TBL_OTKUP, COL_OTK_ID, otkID).count), _
                 "PWA: tacno jedan header"
    AssertEquals "1", CStr(OtkBrojStavkiZaOtkup(otkID)), "PWA: jedna stavka"

    ' Kultura je RAZRESENA, ne fabrikovana.
    AssertEquals TEST_KULTURA_ID, OtkPolje(otkID, COL_OTK_KULTURA), _
                 "PWA: KulturaID je pravi FK, ne 'vrsta-sorta' string"

    ' Brojevi su na stavci, header ih ne nosi.
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_I, COL_OKS_KOLICINA) - 400#) < 0.001, _
               "PWA: kolicina na stavci"
    AssertEquals "", OtkPolje(otkID, COL_OTK_VOZAC), "PWA: header ne nosi vozaca"

    ' Trag porekla.
    AssertEquals crid, OtkPolje(otkID, COL_OTK_CLIENT_RECORD_ID), "PWA: ClientRecordID"
    AssertEquals "PWA", OtkPolje(otkID, COL_OTK_SYNC_SOURCE), "PWA: SyncSource"

    ' Ambalazu knjizi pisac, jednom po dokumentu.
    AssertEquals "2", CStr(AmbBrojRedova(otkID, DOK_TIP_OTKUP)), "PWA: dvojni upis ambalaze"

    Exit Sub

EH:
    LogFatal "Test_PWA_IngestPraviHeaderIStavku", Err.Number, Err.description
End Sub

' Nerazresiva kultura obara uvoz umesto da fabrikuje FK.
Private Sub Test_PWA_NerazresivaKulturaObaraUvoz()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PWAKU")

    Dim crid As String
    crid = TEST_PREFIX & "-CRID-KU-" & scenario

    Dim red As Variant
    red = PwaRed(crid, TEST_PREFIX & "-OTK-PWAKU-" & scenario, 400#, 50#, 20)
    red(1, 13) = TEST_SORTA & " NEPOSTOJECA"        ' GS_SORTA

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim otkID As String
    otkID = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)

    AssertEquals "", otkID, "PWA kultura: uvoz odbijen"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "PWA kultura: nijedan red nije upisan"
    AssertEquals "", modOtkup.OtkupPoClientRecordID(crid), _
                 "PWA kultura: CRID nije zauzet neuspelim uvozom"

    Exit Sub

EH:
    LogFatal "Test_PWA_NerazresivaKulturaObaraUvoz", Err.Number, Err.description
End Sub

' ISTI CRID + ISTI SADRZAJ -> NO-OP.
'
' Retry i ponovljen sync su normalni; smeju da naprave SAMO JEDAN dokument.
Private Sub Test_PWA_IstiCridIstiSadrzajJeNoOp()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PWANO")

    Dim crid As String
    crid = TEST_PREFIX & "-CRID-NO-" & scenario

    Dim red As Variant
    red = PwaRed(crid, TEST_PREFIX & "-OTK-PWANO-" & scenario, 400#, 50#, 20)

    Dim prvi As String
    prvi = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)
    AssertTrue Len(prvi) > 0, "PWA no-op: prvi uvoz prosao"

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim drugi As String
    drugi = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)

    AssertEquals prvi, drugi, "PWA no-op: drugi uvoz vraca ISTI OtkupID"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "PWA no-op: nijedan nov red nije nastao"

    Exit Sub

EH:
    LogFatal "Test_PWA_IstiCridIstiSadrzajJeNoOp", Err.Number, Err.description
End Sub

' ISTI CRID + DRUGI SADRZAJ -> TVRDA GRESKA.
'
' Zatecen kod je svaki poznat CRID preskakao, pa je izmenjen sadrzaj tiho
' nestajao: PWA misli da je poslala ispravku, master je nema i niko ne sazna.
' Ispravka ide kroz storno i nov dokument (A13).
Private Sub Test_PWA_IstiCridDrugiSadrzajPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PWAKF")

    Dim crid As String
    crid = TEST_PREFIX & "-CRID-KF-" & scenario

    Dim red As Variant
    red = PwaRed(crid, TEST_PREFIX & "-OTK-PWAKF-" & scenario, 400#, 50#, 20)

    Dim prvi As String
    prvi = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)
    AssertTrue Len(prvi) > 0, "PWA konflikt: prvi uvoz prosao"

    ' Isti CRID, promenjena SAMO kolicina -- kopija, ne nov PwaRed poziv.
    ' Nov poziv bi pomerio i datum (NextTestDate), pa bi test merio razliku
    ' datuma umesto razlike kolicine i ostao zelen i sa ugasenom kapijom.
    Dim izmenjen As Variant
    izmenjen = red
    izmenjen(1, 15) = 999#                             ' GS_KOLICINA

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim drugi As String
    drugi = modMasterSync.ImportRowToTblOtkup_RowTX(izmenjen, 1, crid)

    AssertEquals "", drugi, "PWA konflikt: izmenjen sadrzaj pod istim CRID-om odbijen"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "PWA konflikt: nijedan nov red nije nastao"

    ' Postojeci dokument je NETAKNUT -- konflikt ne sme da ga prepise.
    AssertTrue Abs(OtkStavkaBrojP(prvi, KLASA_I, COL_OKS_KOLICINA) - 400#) < 0.001, _
               "PWA konflikt: prvi dokument nepromenjen"

    Exit Sub

EH:
    LogFatal "Test_PWA_IstiCridDrugiSadrzajPada", Err.Number, Err.description
End Sub

' STANICA DOKUMENTA JE STANICA UREDJAJA, NE MATICNA STANICA KOOPERANTA.
'
' Kooperant nije zakljucan za stanicu -- svaki moze da preda na svakoj, a otkupni
' list pripada stanici na kojoj je roba predata. tblKooperanti.StanicaID je
' MATICNA stanica i sluzi samo kao filter padajuce liste pri unosu
' (KOOP_FILTER_BY_OM, modOtkupUI.bas:7034) -- nikad za knjizenje.
'
' Do 13.09.2026. je PWA ingest citao stanicu IZ KOOPERANTA, a uredjaj mu je bio
' samo rezerva. Dokument je zavrsavao na pogresnom otkupnom mestu, dok mu je broj
' (koji PWA pravi po uredjaju) tvrdio drugu stanicu.
Private Sub Test_PWA_StanicaJeUredjajNeKooperant()
    On Error GoTo EH

    Dim maticna As String
    maticna = Trim$(CStr(nz(LookupValue(TBL_KOOPERANTI, "KooperantID", _
                                        TEST_KOOP_ID, COL_KOOP_STANICA), "")))

    ' Preduslov: bez razlicite stanice test ne meri nista.
    AssertTrue StrComp(maticna, TEST_HLAD_ST_ID, vbTextCompare) <> 0, _
               "PWA stanica: preduslov -- maticna (" & maticna & ") NIJE stanica uredjaja"

    Dim crid As String: crid = "CRID-STA-" & NewScenarioCode("PWAST")
    Dim red As Variant
    red = PwaRed(crid, "PWA stanica", 100#, 100#, 0)
    red(1, 8) = TEST_HLAD_ST_ID          ' GS_OTKUPAC_ID -- uredjaj DRUGE stanice

    Dim otkID As String
    otkID = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)
    AssertTrue Len(otkID) > 0, "PWA stanica: dokument uvezen"

    Dim upisana As String
    upisana = Trim$(CStr(nz(LookupValue(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_STANICA), "")))

    ' Tvrdnja koja razlikuje uzrok: pre popravke je ovde stajala MATICNA.
    AssertEquals TEST_HLAD_ST_ID, upisana, _
                 "PWA stanica: dokument je knjizen na stanicu UREDJAJA"

    Exit Sub
EH:
    LogFatal "Test_PWA_StanicaJeUredjajNeKooperant", Err.Number, Err.description
End Sub

' FAIL-CLOSED: bez uredjaja se ne zna gde je roba predata, pa se dokument NE PRAVI.
'
' Pogadjanje po kooperantu je bas greska koja je zatvorena, pa prazan OtkupacID
' ne sme da se tiho popuni maticnom stanicom.
'
' Meri se POVRATNA VREDNOST i odsustvo reda, ne dignuta greska: _RowTX po ugovoru
' gresku GUTA -- EH loguje, radi rollback i vraca "" (modMasterSync.bas:2105).
' Prva verzija ovog testa je tvrdila da greska stigne do pozivaoca i pala je iz
' tog razloga, ne zato sto kapija ne radi. Unutrasnji ImportRowToTblOtkup je
' Private, pa se imenovan razlog odavde ne moze procitati -- i to se ne
' pretvara da moze.
Private Sub Test_PWA_BezUredjajaUvozPada()
    On Error GoTo EH

    Dim crid As String: crid = "CRID-NOST-" & NewScenarioCode("PWANO")
    Dim red As Variant
    red = PwaRed(crid, "PWA bez uredjaja", 100#, 100#, 0)
    red(1, 8) = ""                        ' GS_OTKUPAC_ID prazan

    Dim pre As Long
    pre = CountRows(TBL_OTKUP)

    Dim rezultat As String
    rezultat = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)

    AssertEquals "", rezultat, _
                 "PWA bez uredjaja: uvoz NE vraca OtkupID"
    AssertTrue CountRows(TBL_OTKUP) = pre, _
               "PWA bez uredjaja: nijedan red nije upisan (rollback)"

    Exit Sub
EH:
    LogFatal "Test_PWA_BezUredjajaUvozPada", Err.Number, Err.description
End Sub


' ============================================================
' KAPIJA KONTEKSTA BROJA -- modBrojevi.BrojOdgovaraKontekstu
'
' Pravilo: broj dokumenta pripada nizu (vrsta, vlasnik, dan). Kapija ne dokazuje
' da je broj tacan nego da je TUDJ; sto ne govori kanonski jezik ne sudi se.
'
' Zasto je inverzija nosiva: pre nje modBrojevi nije imao NIJEDAN test --
' FormatBroj, IsValidBrojFormat, ApplyMirrorPrefix i SuggestNextBroj su bili
' nemereni. Naivna kapija ("broj mora biti kanonski") oborila bi oko 337
' "TST-PRO-*" brojeva iz ove suite i celu "N/TEST" fixture porodicu.
'
' Druga stanica sa DRUGACIJIM numerickim delom je preduslov svakog dvosmernog
' dokaza: TEST_ST_ID ("ST-90001") i TEST_HLAD_ST_ID ("ST-HLADTEST-90001") oba
' daju 90001, pa nad njima kapija ne moze da razlikuje vlasnike.
' ============================================================

' BKTX_ST2 je u deklaracionoj sekciji, uz ostale TEST_* konstante -- VBA ne
' kompajlira Const ubacen izmedju dve procedure.
Private Sub SeedBktxDrugaStanica()
    SeedStanicaByID BKTX_ST2, "TEST BKTX DRUGA STANICA"
End Sub

' NOSIVI TEST CELOG DIZAJNA. Ako ovaj padne, kapija je postala format-kapija i
' obara sve sto nije kanonski broj -- ukljucujuci ceo fixture.
'
' SABOTAZA: u BrojOdgovaraKontekstu zameni "If Not IsValidBrojFormat(s) Then
' Exit Function" sa dodelom BROJ_KTX_TUDJ_VLASNIK -> pukne po imenu na prvoj
' tvrdnji, a za njim padne i veci deo suite.
Private Sub Test_BKTX_NekanonskiBrojNeOdbija()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("BKTXNK")

    Dim d As Date
    d = NextTestDate()

    ' Predikat cuti na svemu sto nije nas jezik.
    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTK, _
                      TEST_ST_ID, d, TEST_PREFIX & "-OTK-" & scenario)), _
                 "BKTX nekanonski: test-prefiks se ne sudi"

    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTP, _
                      TEST_ST_ID, d, "HL-130926-143205")), _
                 "BKTX nekanonski: fallback broj auto-lanca hladnjace se ne sudi"

    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTP, _
                      TEST_ST_ID, d, "8/TEST")), _
                 "BKTX nekanonski: fixture oblik N/TEST se ne sudi"

    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_ZBR, _
                      TEST_VOZ_ID, d, "ZB-TEST-1")), _
                 "BKTX nekanonski: fixture broj zbirne se ne sudi"

    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTK, _
                      TEST_ST_ID, d, "KUP-RN-2026/117")), _
                 "BKTX nekanonski: eksterni kupcev broj se ne sudi"

    ' I kroz pisca, ne samo kao predikat.
    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-BKTXNK-" & scenario)

    Dim razlog As String
    AssertTrue Len(CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)) > 0, _
               "BKTX nekanonski: pisac prima nekanonski broj (bilo: " & razlog & ")"

    Exit Sub
EH:
    LogFatal "Test_BKTX_NekanonskiBrojNeOdbija", Err.Number, Err.description
End Sub

' OSA VLASNIKA. Broj koji imenuje drugu stanicu ne sme na ovaj dokument.
'
' SABOTAZA: ukloni poredjenje Left$(s, slashPos - 1) <> ocekNum -> pukne po imenu.
Private Sub Test_BKTX_VlasnikOsaOdbijaTudjuStanicu()
    Dim tx As clsTransaction

    On Error GoTo EH

    SeedBktxDrugaStanica

    ' Kontrolni upis prolazi i ostaje u svesci -- test se vraca rollback-om.
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    ' CreateOtkup_TX zove ApplyAvansToOtkup: fixture sme da potrosi ili
    ' podeli slobodan avans kooperanta, pa i tblNovac mora nazad.
    tx.AddTableSnapshot TBL_NOVAC

    ' Preduslov: bez razlicitog numerickog dela test ne meri nista.
    AssertTrue modBrojevi.ExtractNumericFromEntityID(BKTX_ST2) <> _
               modBrojevi.ExtractNumericFromEntityID(TEST_ST_ID), _
               "BKTX vlasnik: preduslov -- dve stanice imaju razlicit numericki deo"

    Dim h As Object
    Set h = OtkHeader("")
    Dim d As Date
    d = h("Datum")

    h("BrojDokumenta") = modBrojevi.FormatBroj(BKTX_ST2, d, 1)

    Dim razlog As String
    AssertEquals "", CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog), _
                 "BKTX vlasnik: broj druge stanice je ODBIJEN"
    AssertTrue InStr(1, razlog, "vlasnik", vbTextCompare) > 0, _
               "BKTX vlasnik: kapija imenuje osu (bilo: " & razlog & ")"

    ' Kontrola: isti dan, isti pisac, broj OVE stanice -> prolazi.
    Dim h2 As Object
    Set h2 = OtkHeader("")
    h2("Datum") = d
    h2("BrojDokumenta") = modBrojevi.FormatBroj(TEST_ST_ID, d, 1)

    AssertTrue Len(CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)) > 0, _
               "BKTX vlasnik: broj ove stanice prolazi (bilo: " & razlog & ")"

    ' Vodeca nula ne menja vlasnika u oba smera: "090001/..." je i dalje stanica
    ' 90001, a "090007/..." je i dalje TUDJA stanica 90007.
    AssertEquals CStr(BROJ_KTX_OK), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTK, _
                      TEST_ST_ID, d, "0" & modBrojevi.FormatBroj(TEST_ST_ID, d, 1))), _
                 "BKTX vlasnik: vodeca nula ne cini broj tudjim"
    AssertEquals CStr(BROJ_KTX_TUDJ_VLASNIK), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTK, _
                      TEST_ST_ID, d, "0" & modBrojevi.FormatBroj(BKTX_ST2, d, 1))), _
                 "BKTX vlasnik: vodeca nula ne sakriva tudjeg vlasnika"

    tx.RollbackTx
    Set tx = Nothing
    Exit Sub
EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFatal "Test_BKTX_VlasnikOsaOdbijaTudjuStanicu", errNum, errDesc
End Sub

' OSA DANA. Broj sa juceradnjim ddmmyy ne sme na danasnji dokument.
'
' SABOTAZA: ukloni poredjenje Mid$(s, slashPos + 1, 6) <> Format$(datum,"ddmmyy")
' -> pukne po imenu.
Private Sub Test_BKTX_DanOsaOdbijaTudjiDan()
    On Error GoTo EH

    Dim h As Object
    Set h = OtkHeader("")

    ' NextTestDate se POMERA na svaki poziv -- dan se pamti u promenljivu.
    Dim d As Date
    d = h("Datum")

    h("BrojDokumenta") = modBrojevi.FormatBroj(TEST_ST_ID, DateAdd("d", -1, d), 1)

    Dim razlog As String
    AssertEquals "", CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog), _
                 "BKTX dan: broj od juce je ODBIJEN"
    AssertTrue InStr(1, razlog, "danu", vbTextCompare) > 0, _
               "BKTX dan: kapija imenuje osu (bilo: " & razlog & ")"

    Exit Sub
EH:
    LogFatal "Test_BKTX_DanOsaOdbijaTudjiDan", Err.Number, Err.description
End Sub

' POZITIVNA KONTROLA PROTIV DRIFTA. Sto generator napravi, kapija mora primiti.
' Bez ovoga bi promena formata u jednom od dva mesta ostala nevidljiva dok
' neko ne izgubi dan posla.
'
' SABOTAZA: u kapiji promeni Format$(datum, "ddmmyy") u "ddmmyyyy", ili
' ExtractNumericFromEntityID u sopstveni parser -> pukne po imenu za svaku vrstu.
Private Sub Test_BKTX_GeneratorUvekProlaziKapiju()
    On Error GoTo EH

    Dim d As Date
    d = NextTestDate()

    Dim broj As String

    broj = modBrojevi.GenerateBrojDokumenta(TEST_ST_ID, d)
    AssertTrue Len(broj) > 0, "BKTX generator: OTK broj je generisan"
    AssertFalse modBrojevi.BrojKontekstOdbija( _
                    modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTK, _
                                                     TEST_ST_ID, d, broj)), _
                "BKTX generator: OTK broj prolazi kapiju (" & broj & ")"

    broj = modBrojevi.GenerateBrojOtpremnice(TEST_ST_ID, d)
    AssertTrue Len(broj) > 0, "BKTX generator: OTP broj je generisan"
    AssertFalse modBrojevi.BrojKontekstOdbija( _
                    modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTP, _
                                                     TEST_ST_ID, d, broj)), _
                "BKTX generator: OTP broj prolazi kapiju (" & broj & ")"

    ' checkRemote:=False -- suite ne sme da zavisi od Drive lookup-a.
    broj = modBrojevi.SuggestNextBroj(modBrojevi.KIND_ZBR, TEST_VOZ_ID, d, False)
    AssertTrue Len(broj) > 0, "BKTX generator: ZBR predlog je generisan"
    AssertFalse modBrojevi.BrojKontekstOdbija( _
                    modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_ZBR, _
                                                     TEST_VOZ_ID, d, broj)), _
                "BKTX generator: ZBR predlog prolazi kapiju (" & broj & ")"

    broj = modBrojevi.SuggestNextBroj(modBrojevi.KIND_REV, TEST_ST_ID, d, False)
    AssertTrue Len(broj) > 0, "BKTX generator: REV predlog je generisan"
    AssertFalse modBrojevi.BrojKontekstOdbija( _
                    modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_REV, _
                                                     TEST_ST_ID, d, broj)), _
                "BKTX generator: REV predlog prolazi kapiju (" & broj & ")"

    Exit Sub
EH:
    LogFatal "Test_BKTX_GeneratorUvekProlaziKapiju", Err.Number, Err.description
End Sub

' MALINA MOD. Zbirna mirror-stanice nosi "S" prefiks i to je NAMERA, ne propust:
' u malina modu je otpremnica = zbirna, jer svako otkupno mesto je i vozac.
' Kapija prefiks skida i ne tumaci -- IsStanicaMirrorVozac je fail-open, pa bi
' pravilo "S mora biti opravdan" pretvorilo svaki neuspeo lookup u odbijanje.
'
' SABOTAZA: ukloni skidanje "S" u ZBR grani -> prve dve tvrdnje puknu po imenu.
Private Sub Test_BKTX_MirrorSPrefiksProlazi()
    On Error GoTo EH

    Const MIR_ST As String = "ST-BKTXMIR-90008"
    SeedStanicaByID MIR_ST, "TEST BKTX MIRROR"

    Dim d As Date
    d = NextTestDate()

    Dim bez As String
    bez = modBrojevi.FormatBroj(MIR_ST, d, 1)

    AssertEquals CStr(BROJ_KTX_OK), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_ZBR, _
                      MIR_ST, d, "S" & bez)), _
                 "BKTX mirror: S-broj mirror-stanice prolazi"

    AssertEquals CStr(BROJ_KTX_OK), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_ZBR, _
                      MIR_ST, d, "s" & bez)), _
                 "BKTX mirror: malo 's' se ponasa isto (poredjenja broja su textcompare)"

    ' PWA nikad ne salje "S" (zbirna.js ne zna za mirror) -- i bez njega prolazi.
    AssertEquals CStr(BROJ_KTX_OK), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_ZBR, _
                      MIR_ST, d, bez)), _
                 "BKTX mirror: isti broj bez prefiksa takodje prolazi"

    ' "S" nije deo jezika drugih vrsta -- tamo string ispada kao "nije nas broj".
    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_OTK, _
                      MIR_ST, d, "S" & bez)), _
                 "BKTX mirror: S na otkupu se ne sudi (nije nas oblik)"

    Exit Sub
EH:
    LogFatal "Test_BKTX_MirrorSPrefiksProlazi", Err.Number, Err.description
End Sub

' SIDRO ZA DUG PR7/PR8, imenovano u komentarima na modMasterSync i
' modAutoHladnjaca. Oba mesta prave BrojZbirne iz BrojOtpremnice, pa numericki
' deo pripada STANICI a vlasnik niza je VOZAC. Danas se poklapa samo zato sto je
' par-vozac mirror (VozacID == StanicaID kao string). Kad se izvedeni lanac
' odmrzne sa realnim vozacem, kapija ce ga odbiti -- i popravka je na IZVORU
' (zbirna dobija svoj broj), ne relaksacija kapije.
'
' SABOTAZA: relaksiraj ZBR granu da prima i vlasnika-stanicu -> pukne po imenu.
Private Sub Test_BKTX_ZbirnaTudjegVlasnikaOdbijena()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim d As Date
    d = NextTestDate()

    ' Broj koji pripada stanici, a vozac je realan (VOZ-90001 -> 90001).
    Dim brojStanice As String
    brojStanice = modBrojevi.FormatBroj(BKTX_ST2, d, 1)

    AssertEquals CStr(BROJ_KTX_TUDJ_VLASNIK), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu(modBrojevi.KIND_ZBR, _
                      TEST_VOZ_ID, d, brojStanice)), _
                 "BKTX zbirna: broj stanice u nizu realnog vozaca je TUDJ"

    ' I kroz pisca.
    AssertEquals "", SaveZbirna_TX(d, TEST_VOZ_ID, brojStanice, TEST_KUP_ID, _
                                   "Test Hladnjaca", "Test Pogon", TEST_VRSTA, _
                                   TEST_SORTA, 1000#, TEST_TIP_AMB, 100, "I"), _
                 "BKTX zbirna: pisac odbija broj tudjeg vlasnika"

    ' Kontrola: vozacev sopstveni broj prolazi kroz istog pisca.
    AssertTrue Len(SaveZbirna_TX(d, TEST_VOZ_ID, _
                                 modBrojevi.FormatBroj(TEST_VOZ_ID, d, 1), _
                                 TEST_KUP_ID, "Test Hladnjaca", "Test Pogon", _
                                 TEST_VRSTA, TEST_SORTA, 1000#, TEST_TIP_AMB, _
                                 100, "I")) > 0, _
               "BKTX zbirna: vozacev sopstveni broj prolazi"

    Exit Sub
EH:
    LogFatal "Test_BKTX_ZbirnaTudjegVlasnikaOdbijena", Err.Number, Err.description
End Sub

' PRIJEMNICA SE NE SUDI, NIKAD. Numericki deo prijemnice hladnjace je konstanta
' "1" (GenerateBrojPrijemnice) i ne kodira kupca; eksterni kupac nosi svoj niz;
' backfill nasledjuje broj preko dana; fixture nosi 22 broja stare konvencije.
' Nijedna tvrdnja o PRJ broju nije istinita u SVIM legitimnim slucajevima.
'
' SABOTAZA: dodaj KIND_PRJ u Select Case dozvoljenih vrsta -> pukne po imenu.
Private Sub Test_BKTX_PrijemnicaNikadNeOdbija()
    On Error GoTo EH

    Dim d As Date
    d = NextTestDate()

    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu("PRJ", TEST_KUP_ID, d, _
                      "77/010199")), _
                 "BKTX prijemnica: ni tudj vlasnik ni tudj dan se ne sude"

    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu("PRJ", TEST_KUP_ID, d, _
                      modBrojevi.FormatBroj("1", d, 1))), _
                 "BKTX prijemnica: ni kanonski oblik hladnjace se ne sudi"

    ' Nepoznata vrsta takodje cuti -- kapija ne obara pozivaoca zbog svog neznanja.
    AssertEquals CStr(BROJ_KTX_NEPRIMENLJIVO), _
                 CStr(modBrojevi.BrojOdgovaraKontekstu("NEPOSTOJECA", TEST_ST_ID, d, _
                      modBrojevi.FormatBroj(TEST_ST_ID, d, 1))), _
                 "BKTX prijemnica: nepoznata vrsta se ne sudi"

    Exit Sub
EH:
    LogFatal "Test_BKTX_PrijemnicaNikadNeOdbija", Err.Number, Err.description
End Sub

' REVERS: kapija sudi SAMO ambalaznu granu. Cist gotovinski promet (F5 isplata /
' F6 uplata) ide kroz istu proceduru sa kolAmb = 0, nema svoj brojevni niz, a
' stanicaID tamo postaje partner-OM -- kapija nad celom procedurom odbijala bi
' legitimnu isplatu na koju je u polju zaostao broj iz prethodnog rezima.
'
' SABOTAZA: pomeri poziv kapije iznad "If kolAmb > 0 Then" -> pukne po imenu na
' trecoj tvrdnji (gotovina).
Private Sub Test_BKTX_ReversSudiSamoAmbalazu()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim d As Date
    d = NextTestDate()

    Dim tudjBroj As String
    tudjBroj = modBrojevi.FormatBroj(BKTX_ST2, d, 1)

    Dim ambPre As Long
    ambPre = CountRows(TBL_AMBALAZA)

    Dim ok As Boolean
    ok = SaveOMUlaz_TX(datum:=d, brojDok:=tudjBroj, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=10, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")

    AssertFalse ok, "BKTX revers: broj tudje stanice je ODBIJEN"
    AssertEquals CStr(ambPre), CStr(CountRows(TBL_AMBALAZA)), _
                 "BKTX revers: odbijen upis nije ostavio ambalaza red"

    ' Kontrola: broj ove stanice prolazi.
    ok = SaveOMUlaz_TX(datum:=d, brojDok:=modBrojevi.FormatBroj(TEST_ST_ID, d, 1), _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=10, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")

    AssertTrue ok, "BKTX revers: broj ove stanice prolazi"

    ' GOTOVINA: isti "tudj" broj, ali kolAmb = 0 -- nema niza, kapija cuti.
    ok = SaveOMUlaz_TX(datum:=d, brojDok:=tudjBroj, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:="", tipAmb:="", kolAmb:=0, _
                       vrstaVoca:=TEST_VRSTA, novac:=5000#, kooperantID:="", _
                       primalacDisplay:="Test primalac", otkupID:="", _
                       tipNovca:=NOV_KES_FIRMA_OTKUPAC, koopSmer:="")

    AssertTrue ok, "BKTX revers: cist gotovinski promet NIJE sudjen po broju"

    Exit Sub
EH:
    LogFatal "Test_BKTX_ReversSudiSamoAmbalazu", Err.Number, Err.description
End Sub

' ZAUZETOST BROJA REVERSA U PISCU -- SaveOMUlaz_TX, po nizu (stanica, dan), sa
' storniranima (A2 red REV, A9). Plus storno i undo po kljucu: isti broj legalno
' nose reversi druge stanice i drugog dana, pa ni storno ni undo ne biraju po broju.
'
' Nivo merenja: poslovni broj u nizu. Smer IZDATO_OM pise JEDNU nogu (Stanica),
' pa je broj redova tblAmbalaza isto sto i broj dokumenata. Broj je nekanonski,
' da kapija konteksta ne sudi -- meri se samo zauzetost.
'
' SABOTAZE: ukloni RequireBrojSlobodanUNizu iz SaveOMUlaz_TX -> pukne "pisac
' odbija isti broj, stanicu i dan"; preskoci stornirane u BrojZauzetRevers ->
' pukne "storno ne oslobadja broj reversa"; u ReversIDRazresi uzmi prvi
' ReversID -> pukne "storno bez identiteta odbija dvosmislen broj".
Private Sub Test_BKTX_ReversPisacOdbijaZauzet()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim scenario As String: scenario = NewScenarioCode("REVBZ")
    Dim d As Date: d = NextTestDate()
    Dim d2 As Date: d2 = DateAdd("d", 1, d)
    Dim broj As String: broj = TEST_PREFIX & "-REV-BZ-" & scenario

    Dim ok As Boolean, pre As Long
    ok = SaveOMUlaz_TX(datum:=d, brojDok:=broj, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=10, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")
    AssertTrue ok, "REV broj: prvi revers se upisuje"

    pre = CountRows(TBL_AMBALAZA)
    ok = SaveOMUlaz_TX(datum:=d, brojDok:=broj, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=5, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")
    AssertFalse ok, "REV broj: pisac odbija isti broj, stanicu i dan"
    ok = SaveOMUlaz_TX(datum:=d, brojDok:=broj, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=5, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="PRIJEM_OD_OM")
    AssertFalse ok, "REV broj: drugi smer istog broja je isti niz"
    AssertEquals CStr(pre), CStr(CountRows(TBL_AMBALAZA)), _
                 "REV broj: odbijeni upisi nisu ostavili red"

    ok = SaveOMUlaz_TX(datum:=d, brojDok:=broj, _
                       stanicaNaziv:="Test OM 2", stanicaID:=BKTX_ST2, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=7, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")
    AssertTrue ok, "REV broj: druga stanica istog dana prima isti broj (A2)"
    ok = SaveOMUlaz_TX(datum:=d2, brojDok:=broj, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=9, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")
    AssertTrue ok, "REV broj: drugi dan iste stanice prima isti broj"

    Dim ambA As String, ambB As String, ambC As String
    ambA = AmbIDNogeStanice(broj, TEST_ST_ID, d)
    ambB = AmbIDNogeStanice(broj, BKTX_ST2, d)
    ambC = AmbIDNogeStanice(broj, TEST_ST_ID, d2)
    AssertTrue Len(ambA) > 0 And Len(ambB) > 0 And Len(ambC) > 0, _
               "REV storno: preduslov -- tri reversa istog broja postoje"

    ' Storno: bez identiteta broj je dvosmislen; sa identitetom pada samo taj revers.
    AssertFalse StornoOMKoopByBrDok_TX(broj, DOK_TIP_OM_ULAZ_FIRMA), _
                "REV storno: bez identiteta odbija dvosmislen broj"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, ambA) Or _
                RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, ambB) Or _
                RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, ambC), _
                "REV storno: odbijen storno nije dirao nijedan revers"
    AssertTrue StornoOMKoopByBrDok_TX(broj, DOK_TIP_OM_ULAZ_FIRMA, ambA), _
               "REV storno: storno po identitetu reda prolazi"
    AssertTrue RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, ambA), _
               "REV storno: storniran je izabrani revers"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, ambB), _
                "REV storno: revers istog broja na drugoj stanici ostaje aktivan"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, ambC), _
                "REV storno: revers istog broja drugog dana ostaje aktivan"

    ' Storno ne oslobadja broj (A9): ni pisac ni ekran.
    pre = CountRows(TBL_AMBALAZA)
    ok = SaveOMUlaz_TX(datum:=d, brojDok:=broj, _
                       stanicaNaziv:="Test OM", stanicaID:=TEST_ST_ID, _
                       vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=10, _
                       vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:="", _
                       primalacDisplay:="", otkupID:="", tipNovca:="", _
                       koopSmer:="IZDATO_OM")
    AssertFalse ok, "REV broj: storno ne oslobadja broj reversa -- pisac (A9)"
    AssertEquals CStr(pre), CStr(CountRows(TBL_AMBALAZA)), _
                 "REV broj: odbijen upis posle storna nije ostavio red"
    AssertEquals ambA, modBrojevi.BrojZauzetUNizu(modBrojevi.KIND_REV, TEST_ST_ID, d, broj), _
                 "REV broj: storno ne oslobadja broj reversa -- provera vraca storniranu nogu"

    ' Undo po operaciji: garda pita kljuc operacije, pa aktivni reversi istog broja
    ' na drugoj stanici i drugog dana ne blokiraju vracanje.
    AssertTrue UndoOperation_TX(LatestOpFor(DOK_TIP_OM_ULAZ_FIRMA, broj)), _
               "REV undo: reversi istog broja drugde ne blokiraju vracanje (garda po kljucu)"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, ambA), _
                "REV undo: izabrani revers je ponovo aktivan"

    Exit Sub
EH:
    LogFatal "Test_BKTX_ReversPisacOdbijaZauzet", Err.Number, Err.description
End Sub

' KOOP REVERS, ISTI BROJ NA DVE STANICE (REV-IDENT-01 Faza 2b). Zabrana istog
' (broj, KOOP smer, dan) na drugoj stanici je uklonjena: noge jednog reversa povezuje
' ReversID, pa su dva takva reversa -- i za ISTOG kooperanta -- dva nezavisna
' dokumenta. Broj i dalje zauzima niz (stanica, dan) za SVA CETIRI smera, sa
' storniranima (A2, A9).
'
' Nivo merenja: poslovni broj u nizu + logicki dokument (storno, undo). Kroz pravi
' pisac, ne seed; noge se nalaze po ReversID-u, jer (broj, kooperant, dan) vise
' nije jednoznacan.
'
' SABOTAZE: vrati u SaveOMUlaz_TX (IZDAVANJE) odbijanje broja i smera koji vec nosi
' aktivan revers -> pukne "REV KOOP 2b: isti broj, smer i dan na drugoj stanici se
' upisuje (i za istog kooperanta)"; isto u grani PRIJEM -> pukne "REV KOOP 2b: povrat
' istog broja i dana na drugoj stanici se upisuje"; u ReversRedoviRID ne poredi ReversID -> pukne
' "REV KOOP 2b: storno S2 ne dira revers S1"; izbaci RequireBrojSlobodanUNizu iz
' SaveOMUlaz_TX -> pukne "REV KOOP 2b: drugi smer istog broja na istoj stanici i
' danu je isti niz".
Private Sub Test_BKTX_ReversKoopIstiBrojDveStanice()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim scenario As String: scenario = NewScenarioCode("REVKS")
    Dim d As Date: d = NextTestDate()
    Dim broj As String: broj = TEST_PREFIX & "-REV-KS-" & scenario
    Dim pre As Long

    AssertTrue UpisiReversTest(d, broj, TEST_ST_ID, TEST_KOOP_ID, "IZDAVANJE"), _
               "REV KOOP 2b: izdavanje kooperantu na S1 upisano"
    pre = CountRows(TBL_AMBALAZA)
    AssertTrue UpisiReversTest(d, broj, BKTX_ST2, TEST_KOOP_ID, "IZDAVANJE"), _
               "REV KOOP 2b: isti broj, smer i dan na drugoj stanici se upisuje (i za istog kooperanta)"
    AssertEquals CStr(pre + 2), CStr(CountRows(TBL_AMBALAZA)), _
                 "REV KOOP 2b: revers na S2 ima obe noge"

    Dim s1 As String, s2 As String, rid1 As String, rid2 As String, k1 As String, k2 As String
    s1 = AmbIDNogeStanice(broj, TEST_ST_ID, d)
    s2 = AmbIDNogeStanice(broj, BKTX_ST2, d)
    AssertTrue Len(s1) > 0 And Len(s2) > 0, "REV KOOP 2b: preduslov -- noge Stanica oba reversa postoje"
    rid1 = ReversIDReda(s1)
    rid2 = ReversIDReda(s2)
    AssertTrue Len(rid1) > 0 And Len(rid2) > 0 And rid1 <> rid2, _
               "REV KOOP 2b: dva reversa istog broja nose razlicit ReversID"
    k1 = NogaPoReversID(rid1, "Kooperant")
    k2 = NogaPoReversID(rid2, "Kooperant")
    AssertTrue Len(k1) > 0 And Len(k2) > 0 And k1 <> k2, _
               "REV KOOP 2b: svaki revers ima svoju nogu Kooperant"

    ' Niz (stanica, dan) i dalje vazi za sva cetiri smera.
    pre = CountRows(TBL_AMBALAZA)
    AssertFalse UpisiReversTest(d, broj, BKTX_ST2, TEST_KOOP2_ID, "PRIJEM"), _
                "REV KOOP 2b: drugi smer istog broja na istoj stanici i danu je isti niz"
    AssertFalse UpisiReversTest(d, broj, TEST_ST_ID, "", "IZDATO_OM"), _
                "REV KOOP 2b: FIRMA smer istog broja na istoj stanici i danu je isti niz"
    AssertEquals CStr(pre), CStr(CountRows(TBL_AMBALAZA)), _
                 "REV KOOP 2b: odbijeni upisi nisu ostavili red"

    ' Isto za povrat (PRIJEM), zasebna grana pisca: isti broj, smer i dan na dve
    ' stanice, drugi broj niza.
    Dim brojP As String: brojP = TEST_PREFIX & "-REV-KSP-" & scenario
    pre = CountRows(TBL_AMBALAZA)
    AssertTrue UpisiReversTest(d, brojP, TEST_ST_ID, TEST_KOOP_ID, "PRIJEM"), _
               "REV KOOP 2b: povrat od kooperanta na S1 upisan"
    AssertTrue UpisiReversTest(d, brojP, BKTX_ST2, TEST_KOOP_ID, "PRIJEM"), _
               "REV KOOP 2b: povrat istog broja i dana na drugoj stanici se upisuje"
    AssertEquals CStr(pre + 4), CStr(CountRows(TBL_AMBALAZA)), _
                 "REV KOOP 2b: oba povrata imaju obe noge"

    ' Bez identiteta reda broj je dvosmislen -- odbija se, nista se ne dira.
    AssertFalse StornoOMKoopByBrDok_TX(broj, DOK_TIP_OM_IZLAZ_KOOP), _
                "REV KOOP 2b: storno bez identiteta reda odbija broj koji nose dva reversa"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, s1) Or RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, s2), _
                "REV KOOP 2b: odbijen storno nije dirao nijedan revers"

    ' Storno sa noge Kooperant S2 dira samo S2 -- isti kooperant, broj, smer i dan.
    AssertTrue StornoOMKoopByBrDok_TX(broj, DOK_TIP_OM_IZLAZ_KOOP, k2), _
               "REV KOOP 2b: storno S2 sa noge Kooperant prolazi"
    AssertTrue RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, k2) And RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, s2), _
               "REV KOOP 2b: stornirane su obe noge S2"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, k1) Or RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, s1), _
                "REV KOOP 2b: storno S2 ne dira revers S1"

    ' Storno ne oslobadja broj u nizu S2 (A9), a treca stanica ga istog dana prima.
    pre = CountRows(TBL_AMBALAZA)
    AssertFalse UpisiReversTest(d, broj, BKTX_ST2, TEST_KOOP2_ID, "IZDAVANJE"), _
                "REV KOOP 2b: storno ne oslobadja broj u nizu S2 (A9)"
    AssertTrue UpisiReversTest(d, broj, TEST_HLAD_ST_ID, TEST_KOOP2_ID, "IZDAVANJE"), _
               "REV KOOP 2b: treca stanica istog dana prima isti broj -- niz je (stanica, dan)"
    AssertEquals CStr(pre + 2), CStr(CountRows(TBL_AMBALAZA)), _
                 "REV KOOP 2b: noge je ostavio samo upis trece stanice"

    ' Undo po operaciji vraca S2, a S1 ostaje netaknut.
    AssertTrue UndoOperation_TX(LatestOpFor(DOK_TIP_OM_IZLAZ_KOOP, broj)), _
               "REV KOOP 2b: undo storna S2 prolazi"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, k2) Or RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, s2), _
                "REV KOOP 2b: undo je vratio obe noge S2"
    AssertFalse RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, k1) Or RedJeStorniran(TBL_AMBALAZA, COL_AMB_ID, s1), _
                "REV KOOP 2b: undo S2 ne dira revers S1"

    Exit Sub
EH:
    LogFatal "Test_BKTX_ReversKoopIstiBrojDveStanice", Err.Number, Err.description
End Sub

' AmbID noge date vrste entiteta pod ReversID-om; prazno kad nema. Od Faze 2b
' (broj, kooperant, dan) nije jednoznacan -- noge jednog reversa nalazi ReversID.
Private Function NogaPoReversID(ByVal rid As String, ByVal entitetTip As String) As String
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If Not IsArray(data) Then Exit Function
    Dim cID As Long, cR As Long, cET As Long, i As Long
    cID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ID)
    cR = GetColumnIndex(TBL_AMBALAZA, COL_AMB_REVERS_ID)
    cET = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    If cID = 0 Or cR = 0 Or cET = 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cR))) = rid And Trim$(NzToText(data(i, cET))) = entitetTip Then
            NogaPoReversID = Trim$(NzToText(data(i, cID)))
            Exit Function
        End If
    Next i
End Function

' REV-IDENT-01 (Faza 1): identitet logickog reversa je ReversID -- ISTI na svim
' nogama jednog dokumenta, RAZLICIT izmedju dokumenata. Kroz pravi pisac.
'
' Nivo merenja: logicki dokument. Preduslov je fizicki (KOOP revers ima dve noge,
' FIRMA jednu), pa se tek onda tvrdi da te noge nose isti, neprazan ReversID. Uz to
' B10 (modIntegritet) prijavljuje nogu kojoj je ReversID obrisan i grupu kojoj time
' fali noga Kooperant -- a ispravan revers ne prijavljuje.
'
' SABOTAZE: izostavi reversID u nozi Kooperant IZDAVANJE grane SaveOMUlaz_TX ->
' pukne "REV-ID KOOP izdavanje: obe noge nose isti ReversID"; izostavi ga u
' PRIJEM_OD_OM grani -> pukne "REV-ID FIRMA prijem od OM: noga Stanica nosi
' ReversID"; u Chk_B10 preskoci prazan ReversID -> pukne "REV-ID B10: noga bez
' ReversID prijavljena"; vrati GetNextID u NoviReversID -> pukne "REV-ID: format
' je opaque RID-<32 hex>".
Private Sub Test_BKTX_ReversIDNaSvimNogama()
    On Error GoTo EH

    Dim scenario As String: scenario = NewScenarioCode("REVID")
    Dim d As Date: d = NextTestDate()
    Dim brojI As String: brojI = TEST_PREFIX & "-REV-IDI-" & scenario
    Dim brojP As String: brojP = TEST_PREFIX & "-REV-IDP-" & scenario
    Dim brojF As String: brojF = TEST_PREFIX & "-REV-IDF-" & scenario
    Dim brojO As String: brojO = TEST_PREFIX & "-REV-IDO-" & scenario

    AssertTrue UpisiReversTest(d, brojI, TEST_ST_ID, TEST_KOOP_ID, "IZDAVANJE"), _
               "REV-ID: izdavanje kooperantu upisano"
    AssertTrue UpisiReversTest(d, brojP, TEST_ST_ID, TEST_KOOP_ID, "PRIJEM"), _
               "REV-ID: povrat od kooperanta upisan"
    AssertTrue UpisiReversTest(d, brojF, TEST_ST_ID, "", "IZDATO_OM"), _
               "REV-ID: FIRMA izdato OM upisan"
    AssertTrue UpisiReversTest(d, brojO, TEST_ST_ID, "", "PRIJEM_OD_OM"), _
               "REV-ID: FIRMA prijem od OM upisan"

    Dim kI As String, sI As String, kP As String, sP As String, sF As String, sO As String
    kI = AmbIDNoge(brojI, TEST_KOOP_ID, "Kooperant", d)
    sI = AmbIDNogeStanice(brojI, TEST_ST_ID, d)
    kP = AmbIDNoge(brojP, TEST_KOOP_ID, "Kooperant", d)
    sP = AmbIDNogeStanice(brojP, TEST_ST_ID, d)
    sF = AmbIDNogeStanice(brojF, TEST_ST_ID, d)
    sO = AmbIDNogeStanice(brojO, TEST_ST_ID, d)
    AssertTrue Len(kI) > 0 And Len(sI) > 0 And Len(kP) > 0 And Len(sP) > 0 And Len(sF) > 0 And Len(sO) > 0, _
               "REV-ID: preduslov -- KOOP reversi imaju obe noge, FIRMA nogu Stanica"
    AssertEquals "", AmbIDNoge(brojF, TEST_KOOP_ID, "Kooperant", d), _
                 "REV-ID: preduslov -- FIRMA revers nema nogu Kooperant"

    Dim ridI As String, ridP As String, ridF As String, ridO As String
    ridI = ReversIDReda(sI)
    ridP = ReversIDReda(sP)
    ridF = ReversIDReda(sF)
    ridO = ReversIDReda(sO)
    AssertTrue Len(ridI) > 0 And Len(ridP) > 0 And Len(ridF) > 0, _
               "REV-ID: svaka noga Stanica nosi ReversID"
    AssertTrue Len(ridO) > 0, "REV-ID FIRMA prijem od OM: noga Stanica nosi ReversID"
    AssertTrue ridI Like "RID-" & String$(32, "?") And Len(ridI) = 36 And Not (Mid$(ridI, 5) Like "*[!0-9A-F]*"), _
               "REV-ID: format je opaque RID-<32 hex>"
    AssertEquals ridI, ReversIDReda(kI), "REV-ID KOOP izdavanje: obe noge nose isti ReversID"
    AssertEquals ridP, ReversIDReda(kP), "REV-ID KOOP povrat: obe noge nose isti ReversID"
    AssertTrue ridI <> ridP And ridI <> ridF And ridP <> ridF And ridO <> ridI And ridO <> ridP And ridO <> ridF, _
               "REV-ID: cetiri dokumenta imaju cetiri razlicita ReversID-a"

    ' B10: identitet se ne pogadja po broju -- noga bez ReversID-a je nalaz.
    Dim r As Long: r = RedAmbalaze(kP)
    AssertTrue r > 0, "REV-ID B10: preduslov -- red noge Kooperant povrata nadjen"
    If r > 0 Then RequireUpdateCell TBL_AMBALAZA, r, COL_AMB_REVERS_ID, "", "Test_BKTX_ReversIDNaSvimNogama"
    AssertTrue InStr(1, IntegritetRedoviSa(kP), "nema ReversID", vbBinaryCompare) > 0, _
               "REV-ID B10: noga bez ReversID prijavljena"
    AssertTrue InStr(1, IntegritetRedoviSa(ridP), "KOOP ocekuje 1", vbBinaryCompare) > 0, _
               "REV-ID B10: revers kome fali noga Kooperant prijavljen"
    AssertEquals "", IntegritetRedoviSa(ridI), "REV-ID B10: ispravan revers nije prijavljen"

    ' Jedan revers sme da nosi VISE tipova ambalaze (odluka 15.09.2026). Fixture
    ' REV-IZV-1 (12/1 + LETVA) je jedan dokument pod jednim ReversID-om: cetiri
    ' aktivne noge, dva tipa. B10 broji noge PO TIPU, pa ga ne prijavljuje.
    Const RID_DVA_TIPA As String = "RID-00000000000000000000000000000001"
    Dim nNogu As Long, nTipova As Long
    NogeReversID RID_DVA_TIPA, nNogu, nTipova
    AssertTrue nNogu = 4 And nTipova = 2, _
               "REV-ID: preduslov -- fixture REV-IZV-1 ima 4 noge dva tipa pod jednim ReversID-om"
    AssertEquals "", IntegritetRedoviSa(RID_DVA_TIPA), _
                 "REV-ID B10: revers sa dva tipa ambalaze nije prijavljen"

    Exit Sub
EH:
    LogFatal "Test_BKTX_ReversIDNaSvimNogama", Err.Number, Err.description
End Sub

' REV-IDENT-01, B10 kao ugovor kome ce Faza 2 verovati: ReversID je JEDAN dokument.
' Sme da nosi vise tipova ambalaze, ali ne sme da spoji delove dva dokumenta --
' drugu stanicu, drugog kooperanta ili drugog vozaca.
'
' To stanje pisac ne moze da napravi (jedan tip i jedan ReversID po pozivu), pa se
' prave dva ISPRAVNA reversa i drugi se PREPISE u prvi: isti ReversID, broj i dan,
' drugi tip ambalaze. Po tipu i po kljucu oblik ostaje ispravan (preduslov), pa
' nalaz mora biti bas stanica / kooperant / vozac.
'
' SABOTAZE: u Chk_B10 ne uskladjuj stanicu -> pukne "REV-ID B10: jedan ReversID na
' dve stanice prijavljen"; ne uskladjuj kooperanta -> "REV-ID B10: jedan KOOP
' ReversID sa dva kooperanta prijavljen"; ne uskladjuj vozaca -> "REV-ID B10: jedan
' FIRMA ReversID sa dva vozaca prijavljen".
Private Sub Test_BKTX_ReversIDJedanDokument()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim scenario As String: scenario = NewScenarioCode("REVJD")
    Dim d As Date: d = NextTestDate()
    Dim brojA As String: brojA = TEST_PREFIX & "-REV-JDA-" & scenario
    Dim brojB As String: brojB = TEST_PREFIX & "-REV-JDB-" & scenario
    Dim brojF1 As String: brojF1 = TEST_PREFIX & "-REV-JDF1-" & scenario
    Dim brojF2 As String: brojF2 = TEST_PREFIX & "-REV-JDF2-" & scenario
    Dim tipB As String: tipB = TEST_TIP_AMB & "-B10"

    AssertTrue UpisiReversTest(d, brojA, TEST_ST_ID, TEST_KOOP_ID, "IZDAVANJE"), _
               "REV-ID JD: KOOP revers A upisan"
    AssertTrue UpisiReversTest(d, brojB, BKTX_ST2, TEST_KOOP2_ID, "IZDAVANJE"), _
               "REV-ID JD: KOOP revers B (druga stanica i kooperant) upisan"
    AssertTrue UpisiReversTest(d, brojF1, TEST_ST_ID, "", "IZDATO_OM"), _
               "REV-ID JD: FIRMA revers F1 upisan"
    AssertTrue UpisiReversTest(d, brojF2, TEST_ST_ID, "", "IZDATO_OM"), _
               "REV-ID JD: FIRMA revers F2 upisan"

    Dim kA As String, sA As String, kB As String, sB As String, sF1 As String, sF2 As String
    kA = AmbIDNoge(brojA, TEST_KOOP_ID, "Kooperant", d)
    sA = AmbIDNogeStanice(brojA, TEST_ST_ID, d)
    kB = AmbIDNoge(brojB, TEST_KOOP2_ID, "Kooperant", d)
    sB = AmbIDNogeStanice(brojB, BKTX_ST2, d)
    sF1 = AmbIDNogeStanice(brojF1, TEST_ST_ID, d)
    sF2 = AmbIDNogeStanice(brojF2, TEST_ST_ID, d)
    AssertTrue Len(kA) > 0 And Len(sA) > 0 And Len(kB) > 0 And Len(sB) > 0 And Len(sF1) > 0 And Len(sF2) > 0, _
               "REV-ID JD: preduslov -- sve noge postoje"

    Dim ridA As String: ridA = ReversIDReda(sA)
    Dim ridF As String: ridF = ReversIDReda(sF1)
    AssertEquals "", IntegritetRedoviSa(ridA), "REV-ID JD: preduslov -- ispravan KOOP revers A nije nalaz"
    AssertEquals "", IntegritetRedoviSa(ridF), "REV-ID JD: preduslov -- ispravan FIRMA revers F1 nije nalaz"

    ' B se prepise u A (druga stanica, drugi kooperant); F2 u F1 (ista stanica, drugi vozac).
    PrepisiNoguURevers kB, ridA, brojA, tipB
    PrepisiNoguURevers sB, ridA, brojA, tipB
    PrepisiNoguURevers sF2, ridF, brojF1, tipB
    If RedAmbalaze(sF2) > 0 Then RequireUpdateCell TBL_AMBALAZA, RedAmbalaze(sF2), COL_AMB_VOZAC, _
                                                     "VOZ-B10-DRUGI", "Test_BKTX_ReversIDJedanDokument"

    Dim nalA As String: nalA = IntegritetRedoviSa(ridA)
    Dim nalF As String: nalF = IntegritetRedoviSa(ridF)
    AssertTrue InStr(1, nalA & nalF, "nogu ", vbBinaryCompare) = 0 And _
               InStr(1, nalA & nalF, "nisu istog", vbBinaryCompare) = 0, _
               "REV-ID JD: preduslov -- po tipu i po kljucu oblik je ispravan"
    AssertTrue InStr(1, nalA, "razlicite stanice", vbBinaryCompare) > 0, _
               "REV-ID B10: jedan ReversID na dve stanice prijavljen"
    AssertTrue InStr(1, nalA, "razlicite kooperante", vbBinaryCompare) > 0, _
               "REV-ID B10: jedan KOOP ReversID sa dva kooperanta prijavljen"
    AssertTrue InStr(1, nalF, "razlicite vozace", vbBinaryCompare) > 0, _
               "REV-ID B10: jedan FIRMA ReversID sa dva vozaca prijavljen"
    AssertTrue InStr(1, nalF, "razlicite stanice", vbBinaryCompare) = 0, _
               "REV-ID B10: FIRMA revers iste stanice nije nalaz za stanicu"

    Exit Sub
EH:
    LogFatal "Test_BKTX_ReversIDJedanDokument", Err.Number, Err.description
End Sub

' Test-only adversarijalno stanje: noga se prepise u drugi revers (ReversID, broj,
' tip ambalaze). Kroz pisac se ne moze napraviti.
Private Sub PrepisiNoguURevers(ByVal ambID As String, ByVal rid As String, _
                               ByVal broj As String, ByVal tipAmb As String)
    Dim r As Long: r = RedAmbalaze(ambID)
    If r <= 0 Then
        Err.Raise vbObjectError + 9901, "PrepisiNoguURevers", "Noga " & ambID & " nije nadjena."
    End If
    RequireUpdateCell TBL_AMBALAZA, r, COL_AMB_REVERS_ID, rid, "PrepisiNoguURevers"
    RequireUpdateCell TBL_AMBALAZA, r, COL_AMB_DOK_ID, broj, "PrepisiNoguURevers"
    RequireUpdateCell TBL_AMBALAZA, r, COL_AMB_TIP, tipAmb, "PrepisiNoguURevers"
End Sub

' Aktivne noge tblAmbalaza pod datim ReversID-om: broj nogu i broj razlicitih
' tipova ambalaze.
Private Sub NogeReversID(ByVal rid As String, ByRef nNogu As Long, ByRef nTipova As Long)
    nNogu = 0
    nTipova = 0
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If Not IsArray(data) Then Exit Sub
    Dim cR As Long, cT As Long, i As Long, tipovi As Object
    cR = GetColumnIndex(TBL_AMBALAZA, COL_AMB_REVERS_ID)
    cT = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    If cR = 0 Or cT = 0 Then Exit Sub
    Set tipovi = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cR))) = rid Then
            nNogu = nNogu + 1
            tipovi(Trim$(NzToText(data(i, cT)))) = True
        End If
    Next i
    nTipova = tipovi.count
End Sub

' ReversID reda tblAmbalaza po AmbID; prazno kad nema.
Private Function ReversIDReda(ByVal ambID As String) As String
    ReversIDReda = Trim$(NzToText(LookupValue(TBL_AMBALAZA, COL_AMB_ID, ambID, COL_AMB_REVERS_ID)))
End Function

' Indeks reda tblAmbalaza sa datim AmbID; 0 kad nema.
Private Function RedAmbalaze(ByVal ambID As String) As Long
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If Not IsArray(data) Then Exit Function
    Dim cID As Long, i As Long
    cID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ID)
    If cID = 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cID))) = ambID Then
            RedAmbalaze = i
            Exit Function
        End If
    Next i
End Function

' Redovi nalaza integriteta (modIntegritet) koji sadrze dati tekst, spojeni vbLf.
Private Function IntegritetRedoviSa(ByVal sadrzi As String) As String
    Dim nal As Variant, i As Long, out As String
    nal = modIntegritet.GetIntegritetRows()
    If Not IsArray(nal) Then Exit Function
    For i = LBound(nal, 1) To UBound(nal, 1)
        If InStr(1, CStr(nal(i, 2)), sadrzi, vbBinaryCompare) > 0 Then out = out & CStr(nal(i, 2)) & vbLf
    Next i
    IntegritetRedoviSa = out
End Function

' Revers kroz pravi pisac (SaveOMUlaz_TX): 5 gajbi test tipa, bez novca.
Private Function UpisiReversTest(ByVal d As Date, ByVal broj As String, _
                                 ByVal stanicaID As String, ByVal kooperantID As String, _
                                 ByVal koopSmer As String) As Boolean
    UpisiReversTest = SaveOMUlaz_TX(datum:=d, brojDok:=broj, _
                                    stanicaNaziv:="Test OM", stanicaID:=stanicaID, _
                                    vozacID:=TEST_VOZ_ID, tipAmb:=TEST_TIP_AMB, kolAmb:=5, _
                                    vrstaVoca:=TEST_VRSTA, novac:=0, kooperantID:=kooperantID, _
                                    primalacDisplay:="", otkupID:="", tipNovca:="", _
                                    koopSmer:=koopSmer)
End Function

' AmbID noge Stanica reversa (broj, stanica, dan) -- identitet reda, onako kako ga
' salje ekran Storno. Prazno kad nema.
Private Function AmbIDNogeStanice(ByVal broj As String, ByVal stanicaID As String, _
                                  ByVal d As Date) As String
    AmbIDNogeStanice = AmbIDNoge(broj, stanicaID, "Stanica", d)
End Function

' AmbID noge reversa (broj, entitet, tip entiteta, dan). Prazno kad nema.
Private Function AmbIDNoge(ByVal broj As String, ByVal entitetID As String, _
                           ByVal entitetTip As String, ByVal d As Date) As String
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If Not IsArray(data) Then Exit Function
    Dim cID As Long, cDok As Long, cEnt As Long, cEntTip As Long, cDat As Long, i As Long
    cID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ID)
    cDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cEnt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    cEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    cDat = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    If cID = 0 Or cDok = 0 Or cEnt = 0 Or cEntTip = 0 Or cDat = 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cDok))) = broj And Trim$(NzToText(data(i, cEnt))) = entitetID _
           And Trim$(NzToText(data(i, cEntTip))) = entitetTip Then
            If IsDate(data(i, cDat)) Then
                If Int(CDbl(CDate(data(i, cDat)))) = Int(CDbl(d)) Then
                    AmbIDNoge = Trim$(NzToText(data(i, cID)))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' RUCNI REZIM. Kad je auto-broj iskljucen u Podesavanjima, operater kuca svoj
' broj -- i ostaje slobodan: broj koji ne govori kanonski jezik se ne sudi. Ali
' rucno otkucan kanonski broj TUDJE stanice i dalje tvrdi tu stanicu, pa ga
' pisac odbija isto kao u auto-rezimu. Istina zapisana u broju ne zavisi od
' globalnog prekidaca.
'
' SABOTAZA: vrati "If Not IsAutoBrojDokumenta() Then Exit Sub" na pocetak
' RequireBrojUKontekstu -> pukne po imenu na drugoj tvrdnji.
Private Sub Test_BKTX_RucniRezimNeGasiPravilo()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim prevMode As String
    prevMode = GetConfigValue(CFG_AUTO_BROJ_DOK)
    SetConfigValue CFG_AUTO_BROJ_DOK, "NO"

    Dim scenario As String
    scenario = NewScenarioCode("BKTXRUC")

    ' Slobodan rucni broj.
    Dim h1 As Object
    Set h1 = OtkHeader("MOJ-OTKUP-" & scenario)
    Dim razlog1 As String, rez1 As String
    rez1 = CreateOtkup_TX(h1, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog1)

    ' Rucno otkucan kanonski broj DRUGE stanice.
    Dim h2 As Object
    Set h2 = OtkHeader("")
    Dim d As Date
    d = h2("Datum")
    h2("BrojDokumenta") = modBrojevi.FormatBroj(BKTX_ST2, d, 1)
    Dim razlog2 As String, rez2 As String
    rez2 = CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog2)

    SetConfigValue CFG_AUTO_BROJ_DOK, prevMode      ' vrati stanje PRE tvrdnji

    AssertTrue Len(rez1) > 0, _
               "BKTX rucni: slobodan rucni broj prolazi (bilo: " & razlog1 & ")"
    AssertEquals "", rez2, _
                 "BKTX rucni: rucno otkucan broj druge stanice je ODBIJEN"
    AssertTrue InStr(1, razlog2, "vlasnik", vbTextCompare) > 0, _
               "BKTX rucni: kapija imenuje osu (bilo: " & razlog2 & ")"

    Exit Sub
EH:
    SetConfigValue CFG_AUTO_BROJ_DOK, prevMode
    LogFatal "Test_BKTX_RucniRezimNeGasiPravilo", Err.Number, Err.description
End Sub

' PWA UVOZ, OBE OSE TVRDO -- i uvoz i pisac govore isto.
'
' Uvoz zove CreateOtkup_TX, gde kapija vec stoji fail-closed. Meka grana na
' uvozu zato nista ne bi propustila, samo bi pomerila poruku sa mesta koje zna
' ClientRecordID na mesto koje ga ne zna. Ovaj test cuva bas to: da se dve
' kapije nad istim brojem ne raziidju.
'
' Ponocna trka u PWA (dva citanja sata sa mreznim await-om izmedju) popravljena
' je na IZVORU u istom PR-u -- otkup-form.js sada cita dan jednom i deli ga i
' broju i zapisu -- pa nov klijent tu neuskladjenost ne moze da proizvede.
'
' DOMET OVOG TESTA, izricito: on meri da broj tudjeg vlasnika ili tudjeg dana NE
' udje u tabelu kroz PWA put. NE moze da razlikuje da li ga je odbio uvoz
' (Err 8109, poruka imenuje ClientRecordID) ili nizvodni CreateOtkup -- interni
' ImportRowToTblOtkup je Private, a _RowTX gresku po ugovoru guta i vraca "".
' Ta razlika se ovde ne pretvara da je merena.
'
' SABOTAZA: ukloni poziv kapije iz CreateOtkup -> prve dve tvrdnje puknu po
' imenu. Ukloni ga sa uvoza -> ostaju zelene, jer pisac i dalje odbija; to je
' poznata granica ovog testa, a ne tiha rupa.
Private Sub Test_BKTX_UvozOtkupaOdbijaTudjBroj()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim scenario As String
    scenario = NewScenarioCode("BKTXPWA")

    ' --- osa vlasnika ---
    Dim cridV As String: cridV = "CRID-BKTXV-" & scenario
    Dim redV As Variant
    redV = PwaRed(cridV, "BKTX vlasnik", 100#, 100#, 0)
    Dim dV As Date: dV = CDate(redV(1, 9))
    redV(1, 23) = modBrojevi.FormatBroj(BKTX_ST2, dV, 1)

    Dim preV As Long: preV = OtkBrojRedova(TBL_OTKUP)

    AssertEquals "", modMasterSync.ImportRowToTblOtkup_RowTX(redV, 1, cridV), _
                 "BKTX uvoz: broj druge stanice NE vraca OtkupID"
    AssertEquals CStr(preV), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "BKTX uvoz: odbijen red nije upisan"

    ' --- osa dana ---
    Dim cridD As String: cridD = "CRID-BKTXD-" & scenario
    Dim redD As Variant
    redD = PwaRed(cridD, "BKTX dan", 100#, 100#, 0)
    Dim dD As Date: dD = CDate(redD(1, 9))
    redD(1, 23) = modBrojevi.FormatBroj(TEST_ST_ID, DateAdd("d", -1, dD), 1)

    Dim preD As Long: preD = OtkBrojRedova(TBL_OTKUP)

    AssertEquals "", modMasterSync.ImportRowToTblOtkup_RowTX(redD, 1, cridD), _
                 "BKTX uvoz: broj od juce NE vraca OtkupID"
    AssertEquals CStr(preD), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "BKTX uvoz: odbijen dan nije ostavio red"

    ' Kontrola: broj ove stanice i ovog dana prolazi kroz isti put.
    Dim cridOK As String: cridOK = "CRID-BKTXOK-" & scenario
    Dim redOK As Variant
    redOK = PwaRed(cridOK, "BKTX kontrola", 100#, 100#, 0)
    Dim dOK As Date: dOK = CDate(redOK(1, 9))
    redOK(1, 23) = modBrojevi.FormatBroj(TEST_ST_ID, dOK, 1)

    AssertTrue Len(modMasterSync.ImportRowToTblOtkup_RowTX(redOK, 1, cridOK)) > 0, _
               "BKTX uvoz: kanonski broj ove stanice i ovog dana prolazi"

    Exit Sub
EH:
    LogFatal "Test_BKTX_UvozOtkupaOdbijaTudjBroj", Err.Number, Err.description
End Sub

' PWA UVOZ ZBIRNE -- fail-closed na kontekst, isto kao oblik i isto kao otkup.
'
' Broj koji protivreci SOPSTVENOM redu ne ulazi u kanonsku tblZbirna. To nije
' kolizija dva dokumenta (legalna po A2, zato PrijaviKolizijuBrojaZbirne samo
' prijavljuje) nego dokument koji laze o sebi -- a BrojZbirne je i danas join
' kljuc u modDokumenta.
'
' SABOTAZA: zameni uslov "If modBrojevi.BrojKontekstOdbija(zbrVerdikt) Then" sa
' "If False Then" -> prve tri tvrdnje puknu po imenu.
Private Sub Test_BKTX_UvozZbirneOdbijaTudjBroj()
    On Error GoTo EH

    SeedBktxDrugaStanica

    Dim scenario As String
    scenario = NewScenarioCode("BKTXZBR")

    Dim d As Date
    d = NextTestDate()

    Dim pre As Long
    pre = OtkBrojRedova(TBL_ZBIRNA)

    ' Broj stanice 90007 na realnom vozacu VOZ-90001.
    AssertEquals "", TestHook_ImportZbirnaRowPWA("CRID-BKTXZV-" & scenario, TEST_VOZ_ID, _
                         TEST_KUP_ID, d, TEST_VRSTA, TEST_SORTA, 100#, _
                         modBrojevi.FormatBroj(BKTX_ST2, d, 1)), _
                 "BKTX uvoz zbirne: broj tudjeg vlasnika NE vraca ZbirnaID"
    AssertEquals CStr(pre), CStr(OtkBrojRedova(TBL_ZBIRNA)), _
                 "BKTX uvoz zbirne: odbijen red nije upisan"

    ' Vozacev broj od juce na danasnjem redu.
    AssertEquals "", TestHook_ImportZbirnaRowPWA("CRID-BKTXZD-" & scenario, TEST_VOZ_ID, _
                         TEST_KUP_ID, d, TEST_VRSTA, TEST_SORTA, 100#, _
                         modBrojevi.FormatBroj(TEST_VOZ_ID, DateAdd("d", -1, d), 1)), _
                 "BKTX uvoz zbirne: broj od juce NE vraca ZbirnaID"

    ' Kontrola: vozacev broj ovog dana prolazi istim putem.
    AssertTrue Len(TestHook_ImportZbirnaRowPWA("CRID-BKTXZOK-" & scenario, TEST_VOZ_ID, _
                         TEST_KUP_ID, d, TEST_VRSTA, TEST_SORTA, 100#, _
                         modBrojevi.FormatBroj(TEST_VOZ_ID, d, 1))) > 0, _
               "BKTX uvoz zbirne: vozacev broj ovog dana prolazi"

    Exit Sub
EH:
    LogFatal "Test_BKTX_UvozZbirneOdbijaTudjBroj", Err.Number, Err.description
End Sub

' Red kakav PWA salje u OTK sheet-u. Indeksi su GS_* kolone modMasterSync-a;
' one su Private tamo, pa se ovde imenuju komentarom, ne konstantom.
'
' ZAMKA: GS_DATUM se puni sa NextTestDate(), koji se POMERA na svaki poziv.
' Dva poziva PwaRed sa istim argumentima zato daju redove koji se razlikuju u
' DATUMU. Test koji tako gradi "izmenjen" red meri razliku datuma, ne razliku
' koju je hteo -- i ostaje zelen i kad se ciljana kapija ugasi (mereno
' sabotazom nad poredjenjem parcele: nije oborila nijednu tvrdnju).
'
' Za konflikt-testove: pozovi JEDNOM, pa kopiraj (VBA niz se kopira dodelom) i
' promeni tacno jedno polje.
Private Function PwaRed(ByVal crid As String, ByVal opisPoziva As String, _
                        ByVal kolicina As Double, ByVal cena As Double, _
                        ByVal kolAmb As Long) As Variant
    Dim r As Variant
    ReDim r(1 To 1, 1 To 23)

    r(1, 1) = crid                  ' GS_CLIENT_RECORD_ID
    r(1, 6) = "PENDING"             ' GS_SYNC_STATUS
    r(1, 8) = TEST_ST_ID            ' GS_OTKUPAC_ID
    r(1, 9) = NextTestDate()        ' GS_DATUM
    r(1, 10) = TEST_KOOP_ID         ' GS_KOOPERANT_ID
    r(1, 12) = TEST_VRSTA           ' GS_VRSTA
    r(1, 13) = TEST_SORTA           ' GS_SORTA
    r(1, 14) = KLASA_I              ' GS_KLASA
    r(1, 15) = kolicina             ' GS_KOLICINA
    r(1, 16) = cena                 ' GS_CENA
    r(1, 17) = TEST_TIP_AMB         ' GS_TIP_AMB
    r(1, 18) = kolAmb               ' GS_KOL_AMB
    r(1, 19) = ""                   ' GS_PARCELA_ID
    r(1, 20) = ""                   ' GS_VOZAC_ID
    ' GS_BROJ_DOKUMENTA se NE salje: kanonski format je ^\d+/\d{6}(-\d+)?$ i
    ' ne trpi test-prefiks, pa ingest generise broj lokalno -- to je i realan
    ' PWA pre-rollout put (modMasterSync: "BrojDokumenta fallback-generated").
    r(1, 23) = ""                   ' GS_BROJ_DOKUMENTA

    PwaRed = r
End Function

' DELJENI RAZRESIVAC (Vrsta, Sorta) -> KulturaID.
'
' Meri se ODVOJENO od ingesta, i to je nalaz iz sabotaze: kad adapter fabrikuje
' kulturu, pisac je odbije kao FK, pa ingest test i dalje prolazi -- odbrana u
' dubinu radi, ali sam razresivac time ostaje nemeren.
'
' Nula i vise od jedan su ISTA greska: izbor se ne prevodi u jedan maticni podatak.
Private Sub Test_PWA_RazresivacImenujeRazlog()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PWARK")

    Dim razlog As String

    ' Postojeci par -> tacno jedan FK.
    AssertEquals TEST_KULTURA_ID, _
                 modOtkup.RazresiKulturuIzVrsteSorte(TEST_VRSTA, TEST_SORTA, razlog), _
                 "PWA razresivac: postojeci par daje FK"
    AssertEquals "", razlog, "PWA razresivac: bez greske kad je jednoznacno"

    ' Nepostojeca sorta -> prazno I imenovan razlog, ne fabrikovan string.
    Dim rez As String
    rez = modOtkup.RazresiKulturuIzVrsteSorte(TEST_VRSTA, TEST_SORTA & " NEMA", razlog)
    AssertEquals "", rez, "PWA razresivac: nepostojeci par ne daje FK"
    AssertTrue InStr(1, razlog, "ne prevodi", vbTextCompare) > 0, _
               "PWA razresivac: imenuje razlog (bilo: " & razlog & ")"
    AssertEquals "0", CStr(InStr(1, rez, "-")), _
                 "PWA razresivac: NE sklapa vrsta-sorta string"

    ' Sorta koja pripada DRUGOJ vrsti -> takodje nema pogotka.
    rez = modOtkup.RazresiKulturuIzVrsteSorte(TEST_VRSTA_BEZ_SORTE, TEST_SORTA, razlog)
    AssertEquals "", rez, "PWA razresivac: sorta druge vrste ne prolazi"

    Exit Sub

EH:
    LogFatal "Test_PWA_RazresivacImenujeRazlog", Err.Number, Err.description
End Sub

' STORNO PREFILL OTKUPA CITA STAVKE (S1b-2, B-036).
'
' Otkup je jedan red zaglavlja; do S1b-2 je prefill trazio red po klasi i citao
' Kolicina/Cena/KolAmbalaze sa zaglavlja, gde CreateOtkup_TX ne pise -- ispravka
' je dolazila prazna. Dvoklasni dokument mora da vrati OBE stavke.
Private Sub Test_OTK_PrefillStornaDveKlaseSaStavki()
    On Error GoTo EH

    Dim otkID As String, brDok As String, spec As String
    brDok = TEST_PREFIX & "-OTK-PF-" & NewScenarioCode("OTKPF")
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(600#, 50#, 12, 400#, 30#, 8))
    AssertTrue Len(otkID) > 0, "OTK prefill: dvoklasni otkup upisan"

    spec = "|" & modStornoDok.PrefillIzStorniranog(STIP_OTKUP, brDok, otkID) & "|"
    AssertTrue InStr(1, spec, "|dveklase=2|", vbBinaryCompare) > 0, _
               "OTK prefill: dve klase (spec: " & spec & ")"
    AssertTrue InStr(1, spec, "|kol1=600|", vbBinaryCompare) > 0, "OTK prefill: kol1 = stavka I"
    AssertTrue InStr(1, spec, "|amb1=12|", vbBinaryCompare) > 0, "OTK prefill: amb1 = stavka I"
    AssertTrue InStr(1, spec, "|cena=50|", vbBinaryCompare) > 0, "OTK prefill: cena = stavka I"
    AssertTrue InStr(1, spec, "|kol2=400|", vbBinaryCompare) > 0, "OTK prefill: kol2 = stavka II"
    AssertTrue InStr(1, spec, "|amb2=8|", vbBinaryCompare) > 0, "OTK prefill: amb2 = stavka II"
    AssertTrue InStr(1, spec, "|cena2=30|", vbBinaryCompare) > 0, "OTK prefill: cena2 = stavka II"

    Exit Sub

EH:
    LogFatal "Test_OTK_PrefillStornaDveKlaseSaStavki", Err.Number, Err.description
End Sub

' STORNO PREFILL OTKUPA BEZ STAVKI NE VRACA DELIMICAN SPEC (review #355, P1).
'
' Synthetic anomaly: kanonski otkup, pa brisanje njegovih stavki u transakciji
' testa. Delimican spec (datum, partner, vrsta -- bez kolicine i cene) izgledao
' bi operateru kao legitimna prazna ispravka. Kontrola pre brisanja dokazuje da
' prazan rezultat posle nije "dokument nije nadjen".
Private Sub Test_OTK_PrefillStornaBezStavkiPada()
    Dim tx As clsTransaction
    On Error GoTo EH

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC

    Dim otkID As String, brDok As String, spec As String
    brDok = TEST_PREFIX & "-OTK-PB-" & NewScenarioCode("OTKPB")
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK prefill bez stavki: otkup upisan"

    spec = modStornoDok.PrefillIzStorniranog(STIP_OTKUP, brDok, otkID)
    AssertTrue InStr(1, "|" & spec & "|", "|kol1=400|", vbBinaryCompare) > 0, _
               "OTK prefill bez stavki: kontrola -- sa stavkama spec postoji (spec: " & spec & ")"

    Dim rows As Collection, k As Long
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID)
    For k = rows.count To 1 Step -1
        RequireDeleteRow TBL_OTKUP_STAVKE, CLng(rows(k)), "Test_OTK_PrefillStornaBezStavkiPada"
    Next k
    AssertEquals "0", CStr(FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID).count), _
                 "OTK prefill bez stavki: stavke zaista obrisane"

    spec = modStornoDok.PrefillIzStorniranog(STIP_OTKUP, brDok, otkID)
    AssertEquals "", spec, "OTK prefill bez stavki: NE vraca delimican spec"

    tx.RollbackTx
    Set tx = Nothing
    Exit Sub

EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogFatal "Test_OTK_PrefillStornaBezStavkiPada", Err.Number, Err.description
End Sub

' IZVOZ OTKUPA CITA STAVKE (S1c, REFAKTOR S14.8 t. 13).
'
' Dvoklasni otkup (I 100 kg x 50, II 40 kg x 30 = 6200) mora u svakom izvozu da
' se vidi kao dve klase sa tacnim kg i vrednoscu: OtkupPoOM (zbir po klasi),
' OtkupiAllStavke (red po stavci), SaldoOMDetail (zbir po kooperantu) i push
' ka stanici (zaglavlje bez linijskih polja + dve stavke u OTK_STAVKE).
' Opseg je sopstveni OtkupID -- izvoz ne sme da zavisi od ostatka sveske.
Private Sub Test_OTK_IzvozDveKlaseIzStavki()
    On Error GoTo EH

    Dim otkID As String, brDok As String
    brDok = TEST_PREFIX & "-OTK-IZ-" & NewScenarioCode("OTKIZ")
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(100#, 50#, 5, 40#, 30#, 2))
    AssertTrue Len(otkID) > 0, "OTK izvoz: dvoklasni otkup upisan"

    Dim samo As Object
    Set samo = CreateObject("Scripting.Dictionary")
    samo.Add otkID, True

    ' OtkupPoOM: naslov + klasa I + klasa II
    Dim om As Variant, r As Long, nI As Long, nII As Long
    om = modStammdatenSync.OtkupPoOMRedovi(samo)
    AssertEquals "3", CStr(UBound(om, 1)), "OTK izvoz OtkupPoOM: naslov + dve klase"
    For r = 2 To UBound(om, 1)
        Select Case CStr(om(r, 3))
            Case KLASA_I
                nI = nI + 1
                AssertEquals "100", CStr(om(r, 4)), "OTK izvoz OtkupPoOM: kg klase I"
                AssertEquals "5", CStr(om(r, 5)), "OTK izvoz OtkupPoOM: gajbe klase I"
                AssertEquals "5000", CStr(om(r, 6)), "OTK izvoz OtkupPoOM: vrednost klase I"
                AssertEquals "1", CStr(om(r, 7)), "OTK izvoz OtkupPoOM: jedan dokument u klasi I"
            Case KLASA_II
                nII = nII + 1
                AssertEquals "40", CStr(om(r, 4)), "OTK izvoz OtkupPoOM: kg klase II"
                AssertEquals "1200", CStr(om(r, 6)), "OTK izvoz OtkupPoOM: vrednost klase II"
        End Select
    Next r
    AssertTrue nI = 1 And nII = 1, "OTK izvoz OtkupPoOM: tacno po jedan red klase I i II"

    ' OtkupiAllStavke: red po stavci, po imenu kolone
    Dim st As Variant, kol As Variant, k As Long
    Dim cOtk As Long, cKl As Long, cId As Long, cKol As Long
    st = modStammdatenSync.OtkupiAllStavkeRedovi(samo)
    kol = modMasterSync.OtkStavkeKolone()
    For k = LBound(kol) To UBound(kol)
        Select Case CStr(kol(k))
            Case COL_OKS_OTKUP_ID: cOtk = k - LBound(kol) + 1
            Case COL_OKS_KLASA: cKl = k - LBound(kol) + 1
            Case COL_OKS_ID: cId = k - LBound(kol) + 1
            Case COL_OKS_KOLICINA: cKol = k - LBound(kol) + 1
        End Select
    Next k
    AssertEquals "3", CStr(UBound(st, 1)), "OTK izvoz OtkupiAllStavke: naslov + dve stavke"
    Dim zbirKg As Double, klase As String
    For r = 2 To UBound(st, 1)
        AssertEquals otkID, CStr(st(r, cOtk)), "OTK izvoz OtkupiAllStavke: roditelj je OtkupID"
        AssertTrue Len(CStr(st(r, cId))) > 0, "OTK izvoz OtkupiAllStavke: stavka nosi OtkupStavkaID"
        zbirKg = zbirKg + CDbl(st(r, cKol))
        klase = klase & "|" & CStr(st(r, cKl))
    Next r
    AssertEquals "140", CStr(zbirKg), "OTK izvoz OtkupiAllStavke: kg obe stavke"
    AssertTrue InStr(1, klase & "|", "|" & KLASA_I & "|", vbBinaryCompare) > 0 And _
               InStr(1, klase & "|", "|" & KLASA_II & "|", vbBinaryCompare) > 0, _
               "OTK izvoz OtkupiAllStavke: obe klase (" & klase & ")"

    ' SaldoOMDetail: kooperant iz stavki
    Dim saldo As Object, v As Variant
    Set saldo = modStammdatenSync.OtkupSaldoPoKooperantu(samo)
    AssertTrue saldo.Exists(TEST_KOOP_ID), "OTK izvoz SaldoOMDetail: kooperant postoji"
    v = saldo(TEST_KOOP_ID)
    AssertEquals "140", CStr(v(1)), "OTK izvoz SaldoOMDetail: kg"
    AssertEquals "6200", CStr(v(2)), "OTK izvoz SaldoOMDetail: vrednost = zbir stavki"
    AssertEquals "7", CStr(v(3)), "OTK izvoz SaldoOMDetail: gajbe"

    ' Push ka stanici: zaglavlje bez linijskih polja, dve stavke
    Dim lo As ListObject, iID As Long, red As Variant, zk As Variant
    Set lo = GetTable(TBL_OTKUP)
    iID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    For r = 1 To lo.DataBodyRange.rows.count
        If CStr(lo.DataBodyRange.cells(r, iID).value) = otkID Then Exit For
    Next r
    red = modStanicaLock.BuildOTKSheetRowForOtkup(otkID, TEST_ST_ID, lo, r, iID)
    AssertTrue IsArray(red), "OTK push: red zaglavlja sastavljen"
    zk = modMasterSync.OtkZaglavljeKolone()
    AssertEquals CStr(UBound(zk) - LBound(zk)), CStr(UBound(red)), "OTK push: red prati spisak kolona"
    For k = LBound(zk) To UBound(zk)
        Select Case CStr(zk(k))
            Case "ServerRecordID"
                AssertEquals otkID, CStr(red(k - LBound(zk))), "OTK push: ServerRecordID = OtkupID"
            Case "Klasa", "Kolicina", "Cena", "KolAmbalaze"
                AssertEquals "", CStr(red(k - LBound(zk))), "OTK push: " & CStr(zk(k)) & " je na stavci, ne u zaglavlju"
        End Select
    Next k
    Dim poOtk As Object
    Set poOtk = modMasterSync.OtkStavkeRedoviPoOtkupu()
    AssertEquals "2", CStr(poOtk(otkID).count), "OTK push: dve stavke za OTK_STAVKE"

    Exit Sub

EH:
    LogFatal "Test_OTK_IzvozDveKlaseIzStavki", Err.Number, Err.description
End Sub

' IZVOZ OTKUPA BEZ STAVKI PADA PO IMENU, NIKAD 0 KG (S1c).
'
' Synthetic anomaly: kanonski otkup, pa brisanje stavki u transakciji testa.
' Kontrola pre brisanja dokazuje da je pad posle zbog stavki, ne zbog opsega.
Private Sub Test_OTK_IzvozBezStavkiPada()
    Dim tx As clsTransaction
    On Error GoTo EH

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC

    Dim otkID As String, brDok As String
    brDok = TEST_PREFIX & "-OTK-IB-" & NewScenarioCode("OTKIB")
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK izvoz bez stavki: otkup upisan"

    Dim samo As Object
    Set samo = CreateObject("Scripting.Dictionary")
    samo.Add otkID, True
    AssertEquals "2", CStr(UBound(modStammdatenSync.OtkupPoOMRedovi(samo), 1)), _
                 "OTK izvoz bez stavki: kontrola -- sa stavkama red postoji"

    Dim rows As Collection, k As Long
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID)
    For k = rows.count To 1 Step -1
        RequireDeleteRow TBL_OTKUP_STAVKE, CLng(rows(k)), "Test_OTK_IzvozBezStavkiPada"
    Next k

    AssertTrue InStr(1, IzvozGreska("POOM", samo), otkID, vbTextCompare) > 0, _
               "OTK izvoz bez stavki: OtkupPoOM pada po imenu dokumenta"
    AssertTrue InStr(1, IzvozGreska("STAVKE", samo), otkID, vbTextCompare) > 0, _
               "OTK izvoz bez stavki: OtkupiAllStavke pada po imenu dokumenta"
    AssertTrue InStr(1, IzvozGreska("SALDO", samo), otkID, vbTextCompare) > 0, _
               "OTK izvoz bez stavki: SaldoOMDetail pada po imenu dokumenta"
    AssertTrue InStr(1, IzvozGreska("PUSH", samo), otkID, vbTextCompare) > 0, _
               "OTK izvoz bez stavki: push stavki pada po imenu dokumenta"

    tx.RollbackTx
    Set tx = Nothing
    Exit Sub

EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogFatal "Test_OTK_IzvozBezStavkiPada", Err.Number, Err.description
End Sub

' Poruka greske izvoza ili "" ako je prosao.
Private Function IzvozGreska(ByVal koji As String, ByVal samo As Object) As String
    Dim x As Variant
    On Error Resume Next
    Err.Clear
    Select Case koji
        Case "POOM": x = modStammdatenSync.OtkupPoOMRedovi(samo)
        Case "STAVKE": x = modStammdatenSync.OtkupiAllStavkeRedovi(samo)
        Case "SALDO": Set x = modStammdatenSync.OtkupSaldoPoKooperantu(samo)
        Case "PUSH": Set x = modMasterSync.OtkStavkeRedoviPoOtkupu()
    End Select
    If Err.Number <> 0 Then IzvozGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' PUSH STAVKI JE IDEMPOTENTAN PO OtkupStavkaID (review #357, P1).
'
' Scenario mreznog pada: prva stavka ode, druga padne, zaglavlje se ne salje.
' Ponovljen pokusaj mora da da TACNO jedan red po stavci i tacnu kolicinu --
' ne 200 kg iz dva reda iste stavke. Isti ID sa drugim sadrzajem je konflikt
' bez upisa; naslov taba u pogresnom redosledu pada (P2).
' Google se simulira kroz TestHook_OtkStavkeSimulacija; indeks se svaki put
' gradi iz "taba" istim putem kao u produkciji (OtkStavkeIndeksIzTaba).
Private Sub Test_OTK_PushStavkiIdempotentan()
    On Error GoTo EH

    Dim otkID As String, brDok As String
    brDok = TEST_PREFIX & "-OTK-PI-" & NewScenarioCode("OTKPI")
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(100#, 50#, 5, 40#, 30#, 2))
    AssertTrue Len(otkID) > 0, "OTK push retry: dvoklasni otkup upisan"

    Dim poOtk As Object, stavke As Collection
    Set poOtk = modMasterSync.OtkStavkeRedoviPoOtkupu()
    Set stavke = poOtk(otkID)
    AssertEquals "2", CStr(stavke.count), "OTK push retry: preduslov -- dve stavke"

    Dim simTab As Collection, greska As String, ok As Boolean
    Set simTab = New Collection

    ' 1) prvi pokusaj: druga stavka padne
    modStanicaLock.TestHook_OtkStavkeSimulacija simTab, 2
    ok = modStanicaLock.PosaljiStavkeOtkupa("SIM", _
            modStanicaLock.OtkStavkeIndeksIzTaba(SimTabStavki(simTab)), stavke, greska)
    AssertTrue Not ok, "OTK push retry: pad druge stavke vraca neuspeh"
    AssertEquals "1", CStr(simTab.count), "OTK push retry: posle pada u tabu je prva stavka"

    ' 2) retry bez pada: tab se cita ponovo, prva stavka se NE salje opet
    modStanicaLock.TestHook_OtkStavkeSimulacija simTab, 0
    ok = modStanicaLock.PosaljiStavkeOtkupa("SIM", _
            modStanicaLock.OtkStavkeIndeksIzTaba(SimTabStavki(simTab)), stavke, greska)
    AssertTrue ok, "OTK push retry: ponovljen pokusaj zavrsava (" & greska & ")"
    AssertEquals "2", CStr(simTab.count), "OTK push retry: tacno jedan red po stavci"

    ' 3) treci pokusaj (npr. palo zaglavlje): nista novo
    ok = modStanicaLock.PosaljiStavkeOtkupa("SIM", _
            modStanicaLock.OtkStavkeIndeksIzTaba(SimTabStavki(simTab)), stavke, greska)
    AssertTrue ok And simTab.count = 2, "OTK push retry: treci pokusaj ne dodaje redove"

    Dim kol As Variant, k As Long, cId As Long, cKol As Long
    kol = modMasterSync.OtkStavkeKolone()
    For k = LBound(kol) To UBound(kol)
        If CStr(kol(k)) = COL_OKS_ID Then cId = k - LBound(kol)
        If CStr(kol(k)) = COL_OKS_KOLICINA Then cKol = k - LBound(kol)
    Next k
    Dim red As Variant, ids As Object, kg As Double
    Set ids = CreateObject("Scripting.Dictionary")
    For Each red In simTab
        AssertTrue Not ids.Exists(CStr(red(cId))), "OTK push retry: OtkupStavkaID jedinstven u tabu"
        ids(CStr(red(cId))) = True
        kg = kg + CDbl(red(cKol))
    Next red
    AssertEquals "140", CStr(kg), "OTK push retry: kolicina = 140 kg, ne dupla"

    ' 4) isti ID, drugi sadrzaj -> konflikt, bez upisa
    Dim lazni As Collection, izmenjen As Variant
    Set lazni = New Collection
    izmenjen = simTab(1)
    izmenjen(cKol) = 999
    lazni.Add izmenjen
    modStanicaLock.TestHook_OtkStavkeSimulacija lazni, 0
    ok = modStanicaLock.PosaljiStavkeOtkupa("SIM", _
            modStanicaLock.OtkStavkeIndeksIzTaba(SimTabStavki(lazni)), stavke, greska)
    AssertTrue Not ok And InStr(1, greska, "konflikt", vbTextCompare) > 0, _
               "OTK push retry: isti ID sa drugim sadrzajem je konflikt (" & greska & ")"
    AssertEquals "1", CStr(lazni.count), "OTK push retry: konflikt ne upisuje nista"

    ' 5) naslov taba u pogresnom redosledu pada po imenu (P2)
    Dim pogresan As Variant
    pogresan = SimTabStavki(simTab)
    Dim tmp As Variant
    tmp = pogresan(1, 1)
    pogresan(1, 1) = pogresan(1, 2)
    pogresan(1, 2) = tmp
    Dim errOpis As String
    On Error Resume Next
    Set ids = modStanicaLock.OtkStavkeIndeksIzTaba(pogresan)
    errOpis = Err.description
    Err.Clear
    On Error GoTo EH
    AssertTrue InStr(1, errOpis, "Naslov taba", vbTextCompare) > 0, _
               "OTK push retry: pogresan redosled naslova pada (" & errOpis & ")"

    modStanicaLock.TestHook_OtkStavkeSimulacija Nothing, 0
    Exit Sub

EH:
    modStanicaLock.TestHook_OtkStavkeSimulacija Nothing, 0
    LogFatal "Test_OTK_PushStavkiIdempotentan", Err.Number, Err.description
End Sub

' Simulirani OTK_STAVKE kao sto ga TryReadSheetData vraca: 2D, red 1 = naslov.
Private Function SimTabStavki(ByVal redovi As Collection) As Variant
    Dim kol As Variant, nk As Long, k As Long, r As Long
    kol = modMasterSync.OtkStavkeKolone()
    nk = UBound(kol) - LBound(kol) + 1
    Dim out() As Variant
    ReDim out(1 To redovi.count + 1, 1 To nk)
    For k = 1 To nk
        out(1, k) = kol(LBound(kol) + k - 1)
    Next k
    Dim red As Variant
    r = 1
    For Each red In redovi
        r = r + 1
        For k = 1 To nk
            out(r, k) = red(LBound(red) + k - 1)
        Next k
    Next red
    SimTabStavki = out
End Function

' STORNO DVOKLASNOG DOKUMENTA JE JEDAN POZIV NAD HEADER-ID.
'
' U starom modelu je dvoklasni blok bio dva reda sa istim BrojDokumenta, pa je
' storno morao da ih grupise -- StornoOtkupByBrDok_TX i GeneracijaID postoje bas
' zbog toga. Sa jednim headerom ta mehanika nema sta da radi.
Private Sub Test_OTK_StornoJednimID()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKSJ")

    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-SJ-" & scenario)
    h.Add "KolAmbIzdata", 7#

    Dim otkID As String
    otkID = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 600#, 40#, 30))
    AssertTrue Len(otkID) > 0, "OTK storno: dvoklasni dokument upisan"
    AssertEquals "2", CStr(OtkBrojStavkiZaOtkup(otkID)), "OTK storno: dve stavke"

    ' JEDAN poziv, nad header-ID.
    AssertTrue StornoOtkup_TX(otkID), "OTK storno: jedan poziv je dovoljan"

    AssertEquals "Da", OtkPolje(otkID, COL_STORNIRANO), "OTK storno: header storniran"

    ' Obe noge ambalaze idu sa dokumentom -- primljena i izdata.
    AssertEquals "Da", AmbPolje(otkID, DOK_TIP_OTKUP, "Izlaz", COL_STORNIRANO), _
                 "OTK storno: primljena ambalaza stornirana"
    AssertEquals "Da", AmbPolje(otkID, DOK_TIP_OM_IZLAZ_KOOP, "Ulaz", COL_STORNIRANO), _
                 "OTK storno: izdata ambalaza stornirana"

    ' Stavke NEMAJU svoj storno: aktivnost stavke je pitanje za header (S7).
    AssertEquals "2", CStr(OtkBrojStavkiZaOtkup(otkID)), _
                 "OTK storno: stavke ostaju, njihovu aktivnost drzi header"

    ' Drugi storno istog dokumenta ne prolazi -- nema sta da se stornira dvaput.
    AssertTrue Not StornoOtkup_TX(otkID), "OTK storno: ponovljeni storno odbijen"

    Exit Sub

EH:
    LogFatal "Test_OTK_StornoJednimID", Err.Number, Err.description
End Sub

' EKRAN I PISAC IMAJU ISTO PRAVILO ZA BROJ.
'
' Zatecena UI provera je isla kroz CheckDuplicate(broj, datum) -- BEZ stanice --
' pa je bila UZA od pisca: dokument koji CreateOtkup_TX smatra legalnim (isti
' broj, druga stanica, isti dan) ekran bi odbio pre nego sto pisac dobije priliku.
'
' Test meri bas taj razmak: isti broj na DRUGOJ stanici mora da prodje i kroz
' ekran, ne samo kroz pisca.
Private Sub Test_OTK_EkranIPisacImajuIstoPravilo()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKEP")

    Dim broj As String
    broj = TEST_PREFIX & "-OTK-EP-" & scenario

    Dim p1 As Object
    Set p1 = OtkEkranParam(broj)
    p1("kolicinaI") = 400#
    p1("cenaI") = 50#
    p1("kolAmb") = 20&

    Dim poruke As String
    AssertTrue Len(modOtkupUnos.OtkupUpisi(p1, poruke)) > 0, _
               "OTK isto pravilo: prvi dokument upisan"

    ' ISTA stanica, isti dan, isti broj -> ekran mora da odbije.
    Dim p2 As Object
    Set p2 = OtkEkranParam(broj)
    p2("datum") = p1("datum")
    p2("kolicinaI") = 400#
    p2("cenaI") = 50#
    p2("kolAmb") = 20&

    Dim fokus As String
    Dim greska As String
    greska = modOtkupUnos.OtkupValidiraj(p2, fokus)

    AssertTrue Len(greska) > 0, "OTK isto pravilo: ista stanica odbijena na ekranu"
    AssertEquals "brDok", fokus, "OTK isto pravilo: fokus je na broju"

    ' DRUGA stanica, isti dan, isti broj -> ekran NE sme da odbije, jer pisac ne bi.
    Dim p3 As Object
    Set p3 = OtkEkranParam(broj)
    p3("datum") = p1("datum")
    p3("stanicaID") = TEST_HLAD_ST_ID
    p3("kolicinaI") = 400#
    p3("cenaI") = 50#
    p3("kolAmb") = 20&

    Dim fokus3 As String
    Dim greska3 As String
    greska3 = modOtkupUnos.OtkupValidiraj(p3, fokus3)

    AssertEquals "", fokus3, "OTK isto pravilo: druga stanica NE pada na broju"
    AssertEquals "", greska3, "OTK isto pravilo: ekran je propustio ono sto pisac dozvoljava"

    ' I pisac ga stvarno prima -- inace bi test dokazao samo da ekran cuti.
    Dim poruke3 As String
    AssertTrue Len(modOtkupUnos.OtkupUpisi(p3, poruke3)) > 0, _
               "OTK isto pravilo: pisac prima drugu stanicu (poruke: " & poruke3 & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_EkranIPisacImajuIstoPravilo", Err.Number, Err.description
End Sub

' STORNO NE OSLOBADJA POSLOVNI BROJ (A9).
'
' Prva verzija kapije je radila ExcludeStornirano, uz obrazlozenje "inace
' ispravka ne bi mogla da zadrzi isti broj" -- a to je bas ono sto A9 zabranjuje:
' ispravka lanca dobija NOV BrojDokumenta, da dva papira razlicitog sadrzaja ne
' bi delila broj.
'
'   OTK120  storniran/zamenjen
'   OTK121  ispravka OTK120     <- nov broj, ne recikliran 120
Private Sub Test_OTK_BrojStorniranogSeNePonovoKoristi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKBS")

    Dim broj As String
    broj = TEST_PREFIX & "-OTK-BS-" & scenario

    Dim h1 As Object
    Set h1 = OtkHeader(broj)
    Dim datum As Date
    datum = h1("Datum")

    Dim prvi As String
    prvi = CreateOtkup_TX(h1, OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(prvi) > 0, "OTK broj storno: prvi dokument prosao"

    ' Storno se ovde radi direktno -- StornoOtkup_TX nad headerom je sledeci korak
    ' cutover-a. Tvrdnja se tice broja, ne mehanike storna.
    RequireUpdateCell TBL_OTKUP, FindRows(TBL_OTKUP, COL_OTK_ID, prvi)(1), _
                      COL_STORNIRANO, "Da", "Test_OTK_BrojStorniranogSeNePonovoKoristi"

    Dim h2 As Object
    Set h2 = OtkHeader(broj)
    h2("Datum") = datum

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK broj storno: broj storniranog se NE oslobadja"
    AssertTrue InStr(1, razlog, "vec izdat", vbTextCompare) > 0, _
               "OTK broj storno: kapija imenuje razlog (bilo: " & razlog & ")"
    AssertTrue InStr(1, razlog, "A9", vbTextCompare) > 0, _
               "OTK broj storno: poruka upucuje na pravilo (bilo: " & razlog & ")"

    ' Ispravka sa NOVIM brojem prolazi -- to je put koji A9 predvidja.
    Dim h3 As Object
    Set h3 = OtkHeader(broj & "-B")
    h3("Datum") = datum
    AssertTrue Len(CreateOtkup_TX(h3, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)) > 0, _
               "OTK broj storno: ispravka sa novim brojem prolazi (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_BrojStorniranogSeNePonovoKoristi", Err.Number, Err.description
End Sub

' Dokument BEZ stavki nije dokument vrednosti nula.
'
' Nula je legitiman odgovor samo kad stavke postoje a zbir im je nula. Bez te
' razlike ApplyAvansToOtkup cita 0 kao "nema sta da se plati" i TIHO preskoci
' primenu avansa -- kvar koji je golden vec jednom prijavio (B2/B3).
' PUN UGOVOR VREDNOSTI: zaglavlje tacno jednom, stavke brojcane i pozitivne.
'
' Ugovor je do koraka 4 bio nepotpun jer su postojala dva pisca zaglavlja bez
' stavki. Oba su zatvorena, pa kapije sad smeju da stoje -- a test postoji da se
' ne vrate tiho. Meri se PORUKOM, ne samo padom: kapija koja padne iz drugog
' razloga ne dokazuje nista.
Private Sub Test_OTK_VrednostPunUgovor()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKVPU")

    ' --- zaglavlje koje ne postoji ---
    AssertTrue InStr(1, VrednostGreska("OTK-NE-POSTOJI-" & scenario), _
                     "ne nalazi tacno jednom", vbTextCompare) > 0, _
               "OTK ugovor: nepostojece zaglavlje pada po imenu"

    ' --- prazan OtkupID ---
    AssertTrue InStr(1, VrednostGreska(""), "Prazan OtkupID", vbTextCompare) > 0, _
               "OTK ugovor: prazan OtkupID pada po imenu"

    ' --- kontrola: ispravan dokument daje broj ---
    Dim okID As String
    okID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-VPU-" & scenario), _
                          OtkStavke(400#, 50#, 20, 100#, 80#, 5))
    AssertTrue Abs(modOtkup.VrednostOtkupa(okID) - (400# * 50# + 100# * 80#)) < 0.001, _
               "OTK ugovor: dve stavke se sabiraju (20000 + 8000)"

    ' --- stavka sa nulom: pisac je ne pravi, pa se pravi RUCNO ---
    ' Nula na stavci nije "dokument vrednosti manje" nego neispravan red: kolicina
    ' i cena su na upisu vec obavezno > 0 (modOtkup:252/261). Citalac mora da drzi
    ' ISTO pravilo, inace se razilaze sa piscem i zbir postaje tisi od istine.
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, okID)
    AssertTrue Not rows Is Nothing, "OTK ugovor: stavke nadjene"
    RequireUpdateCell TBL_OTKUP_STAVKE, rows(1), COL_OKS_CENA, 0#, "Test_OTK_VrednostPunUgovor"

    AssertTrue InStr(1, VrednostGreska(okID), "vece od nule", vbTextCompare) > 0, _
               "OTK ugovor: stavka sa cenom 0 pada po imenu"

    RequireUpdateCell TBL_OTKUP_STAVKE, rows(1), COL_OKS_CENA, "n/d", _
                      "Test_OTK_VrednostPunUgovor"

    AssertTrue InStr(1, VrednostGreska(okID), "nije brojcana", vbTextCompare) > 0, _
               "OTK ugovor: nebrojcana stavka pada po imenu"

    ' CISCENJE JE DEO TESTA, ne kozmetika: pokvarena stavka ostaje u tabeli i
    ' obara SVAKI sledeci citalac koji sabira stavke (GetOpenOtkupi je pao bas
    ' tako). Kvar se pravi namerno, pa se namerno i vraca.
    RequireUpdateCell TBL_OTKUP_STAVKE, rows(1), COL_OKS_CENA, 50#, _
                      "Test_OTK_VrednostPunUgovor"
    AssertTrue Abs(modOtkup.VrednostOtkupa(okID) - (400# * 50# + 100# * 80#)) < 0.001, _
               "OTK ugovor: vrednost vracena posle ciscenja"

    Exit Sub

EH:
    LogFatal "Test_OTK_VrednostPunUgovor", Err.Number, Err.description
End Sub

' Poruka greske koju VrednostOtkupa podigne, ili "" kad prodje.
Private Function VrednostGreska(ByVal otkupID As String) As String
    Dim v As Double
    On Error Resume Next
    Err.Clear
    v = modOtkup.VrednostOtkupa(otkupID)
    If Err.Number <> 0 Then VrednostGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' CITAOCI KOLICINE I VREDNOSTI OTKUPA CITAJU STAVKE (REFAKTOR S14.7, kvarovi 2/3/9).
'
' CreateOtkup_TX linijska polja zaglavlja ostavlja PRAZNA. Svaki citalac koji ih je
' sabirao davao je za nov dokument 0 kg i 0 dinara bez ijedne greske: saldo OM,
' kartica, rekapitulacija robe, otkupne liste, prosecna cena, zbirni OM, rang,
' detalj liste, KPI "danas" i mreza otkupa sa pilulom placanja. Fixture to NE vidi
' -- nosi iste brojeve na zaglavlju i na stavci -- pa se meri nad dokumentom koji
' vrednost nosi SAMO na stavkama. Dan dokumenta je jedinstven (NextTestDate), pa
' izvestaji nad tim danom vide samo njega.
Private Sub Test_OTK_CitaociCitajuStavke()
    On Error GoTo EH

    Const OCEK_KG As Double = 1000#
    Const OCEK_VR As Double = 44000#     ' 400 x 50 + 600 x 40
    Const OCEK_AMB As Double = 50#

    Dim scenario As String, brDok As String, otkID As String, dan As Date
    scenario = NewScenarioCode("OTKCIT")
    brDok = TEST_PREFIX & "-OTK-CIT-" & scenario
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(400#, 50#, 20, 600#, 40#, 30))
    AssertTrue Len(otkID) > 0, "OTK citaoci: dokument napravljen"
    If Len(otkID) = 0 Then Exit Sub
    dan = CDate(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_DATUM))

    ' Vrednost je SAMO na stavkama: zaglavlje od S1d nema kolone Kolicina/Cena
    ' (Test_OTK_HeaderNeNosiLinePolja).

    Dim r As Variant, i As Long, u As Long, nasao As Boolean

    ' --- saldo OM: UKUPNO dana te stanice ---
    r = ReportSaldoOM(TEST_ST_ID, dan, dan)
    AssertTrue IsArray(r), "OTK citaoci: saldo OM postoji"
    If IsArray(r) Then
        u = UBound(r, 1)
        AssertTrue Abs(CDbl(r(u, 2)) - OCEK_KG) < 0.001, "OTK citaoci: saldo OM kg = zbir stavki"
        AssertTrue Abs(CDbl(r(u, 3)) - OCEK_VR) < 0.001, "OTK citaoci: saldo OM vrednost = zbir stavki"
    End If

    ' --- kartica kooperanta: red DOKUMENTA ---
    r = ReportKarticaKooperanta(TEST_KOOP_ID, dan, dan)
    nasao = False
    If IsArray(r) Then
        For i = 1 To UBound(r, 1)
            If CStr(r(i, 9)) = "OTK|" & otkID Then
                nasao = True
                AssertTrue Abs(CDbl(r(i, 5)) - OCEK_VR) < 0.001, _
                           "OTK citaoci: kartica zaduzuje vrednost stavki"
                AssertTrue InStr(1, CStr(r(i, 4)), "I, II", vbBinaryCompare) > 0, _
                           "OTK citaoci: opis kartice nabraja klase stavki"
                ' Saldo gajbi reda = prethodni red + (izdate - primljene). Izdatih
                ' nema, primljene su zbir stavki.
                Dim prethAmb As Double
                prethAmb = 0
                If i > 1 Then
                    If IsNumeric(r(i - 1, 8)) Then prethAmb = CDbl(r(i - 1, 8))
                End If
                AssertTrue Abs((CDbl(r(i, 8)) - prethAmb) + OCEK_AMB) < 0.001, _
                           "OTK citaoci: kartica razduzuje primljene gajbe iz stavki"
            End If
        Next i
    End If
    AssertTrue nasao, "OTK citaoci: kartica ima red dokumenta"

    ' --- rekapitulacija robe: red po klasi STAVKE ---
    Dim kgI As Double, kgII As Double
    kgI = -1: kgII = -1
    r = ReportKarticaRobaRekap(TEST_KOOP_ID, dan, dan)
    AssertTrue IsArray(r), "OTK citaoci: rekapitulacija postoji"
    If IsArray(r) Then
        u = UBound(r, 1)
        AssertEquals "3", CStr(u), "OTK citaoci: rekapitulacija = klasa I + klasa II + UKUPNO"
        AssertTrue Abs(CDbl(r(u, 4)) - OCEK_KG) < 0.001, _
                   "OTK citaoci: rekapitulacija UKUPNO kg = zbir stavki"
        For i = 1 To u - 1
            If CStr(r(i, 3)) = KLASA_I Then kgI = CDbl(r(i, 4))
            If CStr(r(i, 3)) = KLASA_II Then kgII = CDbl(r(i, 4))
        Next i
    End If
    AssertTrue Abs(kgI - 400#) < 0.001, "OTK citaoci: rekapitulacija klasa I = kg stavke"
    AssertTrue Abs(kgII - 600#) < 0.001, "OTK citaoci: rekapitulacija klasa II = kg stavke"

    ' --- otkupne liste: red = DOKUMENT ---
    r = ReportOtkupListe(TEST_ST_ID, dan, dan)
    AssertTrue IsArray(r), "OTK citaoci: otkupne liste postoje"
    If IsArray(r) Then
        AssertEquals "1", CStr(UBound(r, 1)), _
                     "OTK citaoci: otkupne liste -- dvoklasni dokument je jedan red"
        AssertEquals "I, II", CStr(r(1, 5)), "OTK citaoci: otkupne liste nabrajaju klase stavki"
        AssertTrue Abs(CDbl(r(1, 6)) - OCEK_KG) < 0.001, "OTK citaoci: otkupne liste kg = zbir stavki"
        AssertTrue Abs(CDbl(r(1, 7)) - OCEK_VR) < 0.001, _
                   "OTK citaoci: otkupne liste vrednost = zbir stavki"
    End If

    ' --- prosecna cena OM ---
    r = ReportProsecnaCena("OM", TEST_ST_ID, dan, dan)
    AssertTrue IsArray(r), "OTK citaoci: prosecna cena postoji"
    If IsArray(r) Then
        AssertTrue Abs(CDbl(r(1, 2)) - OCEK_KG) < 0.001, "OTK citaoci: prosecna cena kg = zbir stavki"
        AssertTrue Abs(CDbl(r(1, 3)) - OCEK_VR) < 0.001, _
                   "OTK citaoci: prosecna cena vrednost = zbir stavki"
        AssertTrue Abs(CDbl(r(1, 4)) - OCEK_VR / OCEK_KG) < 0.001, _
                   "OTK citaoci: prosecna cena = vrednost / kg"
    End If

    ' --- zbirni OM: UKUPNO dana (sve stanice) ---
    r = ReportZbirni("OM", dan, dan)
    AssertTrue IsArray(r), "OTK citaoci: zbirni OM postoji"
    If IsArray(r) Then
        u = UBound(r, 1)
        AssertTrue Abs(CDbl(r(u, 3)) - OCEK_KG) < 0.001, "OTK citaoci: zbirni OM kg = zbir stavki"
        AssertTrue Abs(CDbl(r(u, 4)) - OCEK_VR) < 0.001, "OTK citaoci: zbirni OM vrednost = zbir stavki"
    End If

    ' --- rang kooperanata u danu ---
    Dim rKg As Double, rVal As Double, eKg As Double, eVal As Double
    r = modOtkupBlok.KoopRangRows(rKg, rVal, eKg, eVal, CDbl(Int(CDbl(dan))), CDbl(Int(CDbl(dan))))
    AssertTrue Abs(rKg - OCEK_KG) < 0.001, "OTK citaoci: rang kg = zbir stavki"
    AssertTrue Abs(rVal - OCEK_VR) < 0.001, "OTK citaoci: rang iznos = zbir stavki"

    ' --- detalj reda liste: linije = stavke dokumenta ---
    Dim det As Variant
    det = modScrIzvestaji.IzDetaljOtkupLista(otkID)
    AssertTrue IsArray(det), "OTK citaoci: detalj dokumenta postoji"
    If IsArray(det) Then
        ' Nov dokument nema vozaca ni zbirnu na zaglavlju, pa nema ni linije konteksta.
        AssertEquals "3", CStr(UBound(det) - LBound(det) + 1), _
                     "OTK citaoci: detalj = dve stavke + UKUPNO"
        AssertTrue InStr(1, CStr(det(LBound(det))), " x ", vbBinaryCompare) > 0, _
                   "OTK citaoci: prva linija detalja je stavka sa cenom"
        AssertTrue InStr(1, CStr(det(UBound(det))), "UKUPNO", vbBinaryCompare) = 1, _
                   "OTK citaoci: detalj dvoklasnog dokumenta nosi UKUPNO"
    End If

    ' --- KPI "danas" ---
    AssertTrue Abs(modOtkup.KgOtkupaZaDan(dan) - OCEK_KG) < 0.001, _
               "OTK citaoci: KPI kg dana = zbir stavki"

    ' --- mreza otkupa: kolone i pilula placanja ---
    Dim g As Variant
    g = OtkMrezaRed(brDok)
    AssertTrue IsArray(g), "OTK citaoci: mreza ima red dokumenta"
    If IsArray(g) Then
        AssertTrue Abs(CDbl(g(0)) - OCEK_KG) < 0.001, "OTK citaoci: mreza kg = zbir stavki"
        AssertTrue Abs(CDbl(g(1)) - OCEK_VR) < 0.001, "OTK citaoci: mreza vrednost = zbir stavki"
        AssertTrue Abs(CDbl(g(2)) - OCEK_AMB) < 0.001, "OTK citaoci: mreza gajbe = zbir stavki"
        AssertEquals "I, II", CStr(g(3)), "OTK citaoci: mreza klasa nabraja klase stavki"
        AssertEquals CStr(PAY_NEPLAC), CStr(g(4)), "OTK citaoci: pilula -- neplacen dokument"
        AssertTrue Abs(CDbl(g(5)) - OCEK_VR) < 0.001, "OTK citaoci: ostatak neplacenog = vrednost"
    End If

    SaveNovac TEST_PREFIX & "-NOV-CIT-D-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, 4000#, "delimicno", otkID
    g = OtkMrezaRed(brDok)
    AssertTrue IsArray(g), "OTK citaoci: mreza posle delimicne isplate"
    If IsArray(g) Then
        AssertEquals CStr(PAY_DELIM), CStr(g(4)), _
                     "OTK citaoci: pilula -- delimicna isplata je DELIMICNO, ne placeno"
        AssertTrue Abs(CDbl(g(5)) - (OCEK_VR - 4000#)) < 0.001, _
                   "OTK citaoci: ostatak = vrednost - placeno"
    End If

    SaveNovac TEST_PREFIX & "-NOV-CIT-P-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, OCEK_VR - 4000#, "ostatak", otkID
    g = OtkMrezaRed(brDok)
    AssertTrue IsArray(g), "OTK citaoci: mreza posle pune isplate"
    If IsArray(g) Then
        AssertEquals CStr(PAY_PLACENO), CStr(g(4)), "OTK citaoci: pilula -- puna isplata je placeno"
    End If

    Exit Sub

EH:
    LogFatal "Test_OTK_CitaociCitajuStavke", Err.Number, Err.description
End Sub

' Red dokumenta u mrezi otkupa -- ISTI poziv koji crta ekran
' (modScrDokumenti.RedoviZaTip), nadjen pretragom po jedinstvenom broju. Kes se
' prazni jer mreza cita kesirane tabele. Vraca Array(kg, vrednost, gajbe, klasa,
' pilula, ostatak), ili Empty kad red nije tacno jedan.
Private Function OtkMrezaRed(ByVal brDok As String) As Variant
    Dim d As Variant, cols As Variant, redovi As Variant, c As Long
    Dim iKg As Long, iVr As Long, iAmb As Long, iKl As Long, iPill As Long, iRest As Long

    modUiData.ResetCache
    d = modScrDokumenti.RedoviZaTip("OTKUP", "", brDok)
    If Not IsArray(d) Then Exit Function
    If CLng(d(2)) <> 1 Then Exit Function

    cols = d(0)
    redovi = d(1)
    For c = 0 To UBound(cols)
        Select Case modScrDokumenti.ColF(CStr(cols(c)), 2)
            Case "kg":      iKg = c + 1
            Case "mult":    iVr = c + 1
            Case "paypill": iPill = c + 1
            Case "rest":    iRest = c + 1
        End Select
        Select Case modScrDokumenti.ColF(CStr(cols(c)), 1)
            Case COL_OKS_KOL_AMB: iAmb = c + 1
            Case COL_OKS_KLASA:   iKl = c + 1
        End Select
    Next c
    If iKg = 0 Or iVr = 0 Or iAmb = 0 Or iKl = 0 Or iPill = 0 Or iRest = 0 Then Exit Function

    OtkMrezaRed = Array(redovi(1, iKg), redovi(1, iVr), redovi(1, iAmb), _
                        redovi(1, iKl), redovi(1, iPill), redovi(1, iRest))
End Function

' Red dokumenta u mrezi OTPREMNICA -- isti obrazac kao OtkMrezaRed, isti poziv
' koji crta ekran. Vraca Array(kg, gajbe, klasa) ili Empty kad red nije tacno
' jedan. Vrednosti nema: otpremnica u mrezi ne nosi finansijsku cifru.
Private Function OtpMrezaRed(ByVal broj As String) As Variant
    Dim d As Variant, cols As Variant, redovi As Variant, c As Long
    Dim iKg As Long, iAmb As Long, iKl As Long

    modUiData.ResetCache
    d = modScrDokumenti.RedoviZaTip("OTPREMNICA", "", broj)
    If Not IsArray(d) Then Exit Function
    If CLng(d(2)) <> 1 Then Exit Function

    cols = d(0)
    redovi = d(1)
    For c = 0 To UBound(cols)
        Select Case modScrDokumenti.ColF(CStr(cols(c)), 0)
            Case "OTKUI_HD_KG":        iKg = c + 1
            Case "OTKUI_HD_KOL_AMB":   iAmb = c + 1
            Case "OTKUI_HD_KLASA":     iKl = c + 1
        End Select
    Next c
    If iKg = 0 Or iAmb = 0 Or iKl = 0 Then Exit Function

    OtpMrezaRed = Array(redovi(1, iKg), redovi(1, iAmb), redovi(1, iKl))
End Function

' Da li mreza OTPREMNICA ima kolonu sa datim kljucem naslova.
Private Function OtpMrezaImaKolonu(ByVal kljuc As String) As Boolean
    Dim cols As Variant, c As Long
    cols = modScrDokumenti.GridCols("OTPREMNICA")
    For c = 0 To UBound(cols)
        If modScrDokumenti.ColF(CStr(cols(c)), 0) = kljuc Then
            OtpMrezaImaKolonu = True
            Exit Function
        End If
    Next c
End Function

' Roba po vozacu (zbirno) za TEST vozaca na JEDAN dan: kg (kolona 3) ili
' vrednost (kolona 4) njegovog reda, formatirano "0.00". Nema reda = "0.00".
' Greska izvestaja se VRACA kao tekst, da tvrdnja padne po imenu umesto da
' ceo test padne na prvom pozivu.
Private Function OtpVozacRoba(ByVal dan As Date, ByVal kolona As Long) As String
    Dim r As Variant, i As Long
    On Error GoTo EH
    modUiData.ResetCache
    OtpVozacRoba = Format$(0#, "0.00")
    r = ReportRobaVozaciZbirni(dan, dan)
    If Not IsArray(r) Then Exit Function
    For i = 1 To UBound(r, 1)
        If Trim$(CStr(r(i, 1))) = TEST_VOZ_ID Then
            OtpVozacRoba = Format$(CDbl(r(i, kolona)), "0.00")
            Exit Function
        End If
    Next i
    Exit Function
EH:
    OtpVozacRoba = "GRESKA: " & Err.description
End Function

' Broj redova otpremnice (po broju) u izvestaju "Otkupljena roba (OM)" za TEST
' stanicu na jedan dan; greska izvestaja = -1.
Private Function OtpOmRedova(ByVal dan As Date, ByVal broj As String) As Long
    Dim r As Variant, i As Long
    On Error GoTo EH
    modUiData.ResetCache
    r = ReportOtkupRoba("OM", TEST_ST_ID, dan, dan)
    If Not IsArray(r) Then Exit Function
    For i = 1 To UBound(r, 1)
        If Trim$(CStr(r(i, 2))) = broj Then OtpOmRedova = OtpOmRedova + 1
    Next i
    Exit Function
EH:
    OtpOmRedova = -1
End Function

Private Function OtpMrezaBrojRedova(ByVal broj As String) As Long
    Dim d As Variant
    modUiData.ResetCache
    d = modScrDokumenti.RedoviZaTip("OTPREMNICA", "", broj)
    If Not IsArray(d) Then Exit Function
    OtpMrezaBrojRedova = CLng(d(2))
End Function

' Poruka greske koju mreza otpremnica podigne, ili "" kad prodje.
Private Function OtpMrezaGreska(ByVal broj As String) As String
    Dim d As Variant
    On Error Resume Next
    Err.Clear
    modUiData.ResetCache
    d = modScrDokumenti.RedoviZaTip("OTPREMNICA", "", broj)
    If Err.Number <> 0 Then OtpMrezaGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' Citalac stavki drzi ISTA pravila kao kanon (VrednostOtkupa): pokvarena stavka
' obara zbir PO IMENU umesto da se preskoci -- preskakanje bi izvestaju tiho
' umanjilo kolicinu i vrednost. Kvar se pravi rucno (pisac ga ne pravi) i vraca.
Private Sub Test_OTK_CitaociStavkiFailClosed()
    Dim rows As Collection
    On Error GoTo EH

    Dim scenario As String, otkID As String
    scenario = NewScenarioCode("OTKCFC")
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-CFC-" & scenario), _
                           OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK citaoci kapija: dokument napravljen"
    If Len(otkID) = 0 Then Exit Sub

    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID)
    AssertTrue Not rows Is Nothing, "OTK citaoci kapija: stavka nadjena"
    If rows Is Nothing Then Exit Sub

    RequireUpdateCell TBL_OTKUP_STAVKE, rows(1), COL_OKS_CENA, 0#, _
                      "Test_OTK_CitaociStavkiFailClosed"
    AssertTrue InStr(1, ZbirStavkiGreska(), "vece od nule", vbTextCompare) > 0, _
               "OTK citaoci kapija: stavka sa cenom 0 obara zbir po imenu"

    RequireUpdateCell TBL_OTKUP_STAVKE, rows(1), COL_OKS_CENA, "n/d", _
                      "Test_OTK_CitaociStavkiFailClosed"
    AssertTrue InStr(1, ZbirStavkiGreska(), "nije brojcana", vbTextCompare) > 0, _
               "OTK citaoci kapija: nebrojcana stavka obara zbir po imenu"

    ' CISCENJE JE DEO TESTA: pokvarena stavka obara svaki sledeci citalac stavki.
    RequireUpdateCell TBL_OTKUP_STAVKE, rows(1), COL_OKS_CENA, 50#, _
                      "Test_OTK_CitaociStavkiFailClosed"
    AssertEquals "", ZbirStavkiGreska(), "OTK citaoci kapija: zbir prolazi posle ciscenja"
    Exit Sub

EH:
    LogFatal "Test_OTK_CitaociStavkiFailClosed", Err.Number, Err.description
    On Error Resume Next
    If Not rows Is Nothing Then
        RequireUpdateCell TBL_OTKUP_STAVKE, rows(1), COL_OKS_CENA, 50#, _
                          "Test_OTK_CitaociStavkiFailClosed"
    End If
End Sub

' Poruka greske koju modOtkup.ZbirStavkiPoOtkupu podigne, ili "" kad prodje.
Private Function ZbirStavkiGreska() As String
    Dim z As Object
    On Error Resume Next
    Err.Clear
    Set z = modOtkup.ZbirStavkiPoOtkupu()
    If Err.Number <> 0 Then ZbirStavkiGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' Poruka greske koju saldo OM podigne, ili "" kad prodje.
Private Function SaldoOMGreska(ByVal dan As Date) As String
    Dim r As Variant
    On Error Resume Next
    Err.Clear
    r = ReportSaldoOM(TEST_ST_ID, dan, dan)
    If Err.Number <> 0 Then SaldoOMGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' Poruka greske koju mreza otkupa podigne, ili "" kad prodje. Isti poziv
' koji crta ekran -- greska ide kroz RedoviZaTip i nosi ime koraka.
Private Function MrezaGreska(ByVal brDok As String) As String
    Dim d As Variant
    On Error Resume Next
    Err.Clear
    modUiData.ResetCache
    d = modScrDokumenti.RedoviZaTip("OTKUP", "", brDok)
    If Err.Number <> 0 Then MrezaGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' Poruka greske koju lista otvorenih obaveza podigne, ili "" kad prodje.
Private Function OtvoreniGreska() As String
    Dim r As Variant
    On Error Resume Next
    Err.Clear
    r = GetOpenOtkupi("")
    If Err.Number <> 0 Then OtvoreniGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' ZAGLAVLJE BEZ STAVKI NE SME DA IZGLEDA PLACENO (review #334, P1).
'
' Dokument se napravi kanonski i isplati DO KRAJA, pa mu se stavke OBRISU --
' tacno oblik koji je ranije prolazio tiho: duguje se racunao iz zbira stavki
' kojih nema, pa je ispadao 0, a PayCode(0, placeno > 0) je pilulu bojio u
' PLACENO. Izvestaji su isti dokument prikazivali sa 0 kg i 0 dinara, a lista
' za isplatu ga je preskakala uz jedan red u logu.
'
' Meri se PORUKOM, ne samo padom: citalac koji padne iz drugog razloga ne
' dokazuje nista. Brisanje je u transakciji i vraca se ROLLBACK-om.
Private Sub Test_OTK_ZaglavljeBezStavkiObaraCitaoce()
    Const OCEK_VR As Double = 20000#
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String, brDok As String, otkID As String, dan As Date
    scenario = NewScenarioCode("OTKHBS")
    brDok = TEST_PREFIX & "-OTK-HBS-" & scenario
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK bez stavki: dokument napravljen"
    If Len(otkID) = 0 Then Exit Sub
    dan = CDate(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_DATUM))

    ' Placen DO KRAJA: bez toga se "placeno" ne razlikuje od "nema duga".
    SaveNovac TEST_PREFIX & "-NOV-HBS-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, OCEK_VR, "puna isplata", otkID

    Dim g As Variant
    g = OtkMrezaRed(brDok)
    AssertTrue IsArray(g), "OTK bez stavki: kontrola -- mreza ima red dokumenta"
    If IsArray(g) Then
        AssertEquals CStr(PAY_PLACENO), CStr(g(4)), _
                     "OTK bez stavki: kontrola -- placen dokument je PLACENO dok stavke postoje"
    End If

    ' --- stavke se BRISU: dokument postaje zaglavlje bez stavki ---
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP_STAVKE

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID)
    AssertTrue Not rows Is Nothing, "OTK bez stavki: stavke nadjene"
    If rows Is Nothing Then Exit Sub

    Dim k As Long
    For k = rows.count To 1 Step -1
        RequireDeleteRow TBL_OTKUP_STAVKE, CLng(rows(k)), _
                         "Test_OTK_ZaglavljeBezStavkiObaraCitaoce"
    Next k
    AssertEquals "0", CStr(OtkBrojStavki(otkID)), _
                 "OTK bez stavki: preduslov -- dokument je ostao bez stavki"

    ' Svaki citalac PADA i imenuje dokument -- nijedan ne vraca nulu.
    AssertTrue InStr(1, ZbirStavkiGreska(), "nema nijednu stavku", vbTextCompare) > 0, _
               "OTK bez stavki: zbir stavki pada po imenu"
    AssertTrue InStr(1, ZbirStavkiGreska(), otkID, vbTextCompare) > 0, _
               "OTK bez stavki: poruka imenuje sporan dokument"
    AssertTrue InStr(1, SaldoOMGreska(dan), "nema nijednu stavku", vbTextCompare) > 0, _
               "OTK bez stavki: saldo OM pada po imenu, ne vraca 0 kg"
    AssertTrue InStr(1, MrezaGreska(brDok), "nema nijednu stavku", vbTextCompare) > 0, _
               "OTK bez stavki: mreza pada po imenu -- pilula ne moze da kaze placeno"
    AssertTrue InStr(1, OtvoreniGreska(), "nema nijednu stavku", vbTextCompare) > 0, _
               "OTK bez stavki: lista za isplatu pada po imenu, ne preskace red"

    ' --- vracanje: isti citaoci su opet zeleni ---
    tx.RollbackTx
    Set tx = Nothing

    AssertEquals "", ZbirStavkiGreska(), _
                 "OTK bez stavki: zbir prolazi posle vracanja stavki"
    g = OtkMrezaRed(brDok)
    AssertTrue IsArray(g), "OTK bez stavki: mreza opet ima red dokumenta"
    If IsArray(g) Then
        AssertEquals CStr(PAY_PLACENO), CStr(g(4)), _
                     "OTK bez stavki: pilula je opet PLACENO posle vracanja stavki"
    End If

    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFatal "Test_OTK_ZaglavljeBezStavkiObaraCitaoce", errNum, errDesc
End Sub

' STAVKA MORA IMATI SVOJ DOKUMENT, I POLJA KAO KOD PISCA (review #334, P1).
'
' Citalac je stavku bez OtkupID-a TIHO PRESKAKAO, a nebrojcanu KolAmbalaze
' pretvarao u 0 -- iako pisac oba odbija na upisu (RequireValidOtkupClass,
' kolAmb >= 0, RequireCeoBrojOtk). Razlika izmedju pisca i citaoca je bila
' tiha: kolicina dokumenta u izvestaju manja nego sto jeste, gajbe nestale.
'
' DUPLIKAT zaglavlja se ovde NE meri: taj rod ima svoje imenovane kapije
' (ERR_ISPLATA_DUPLI_OTKUPID, VrednostOtkupa) i svoj dokaz (modTestBanka T17).
Private Sub Test_OTK_StavkaBezZaglavljaObaraCitaoce()
    Const SRC As String = "Test_OTK_StavkaBezZaglavljaObaraCitaoce"
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String, otkID As String
    scenario = NewScenarioCode("OTKSBZ")
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-SBZ-" & scenario), _
                           OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK siroce: dokument napravljen"
    If Len(otkID) = 0 Then Exit Sub

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID)
    AssertTrue Not rows Is Nothing, "OTK siroce: stavka nadjena"
    If rows Is Nothing Then Exit Sub

    Dim red As Long
    red = CLng(rows(1))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP_STAVKE

    ' --- stavka bez OtkupID-a: ranije tiho preskocena ---
    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_OTKUP_ID, "", SRC
    AssertTrue InStr(1, ZbirStavkiGreska(), "bez OtkupID-a", vbTextCompare) > 0, _
               "OTK siroce: stavka bez OtkupID-a obara citaoca po imenu"

    ' --- stavka ciji dokument ne postoji ---
    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_OTKUP_ID, _
                      "OTK-NE-POSTOJI-" & scenario, SRC
    AssertTrue InStr(1, ZbirStavkiGreska(), "Zaglavlje otkupa ne postoji", vbTextCompare) > 0, _
               "OTK siroce: stavka bez zaglavlja obara citaoca po imenu"
    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_OTKUP_ID, otkID, SRC

    ' --- klasa: ISTA kapija koju pisac trazi na upisu ---
    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_KLASA, "III", SRC
    AssertTrue InStr(1, ZbirStavkiGreska(), "Neispravna klasa", vbTextCompare) > 0, _
               "OTK siroce: nevalidna klasa stavke obara citaoca po imenu"
    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_KLASA, KLASA_I, SRC

    ' --- KolAmbalaze: broj, nenegativan, ceo -- kao kod pisca ---
    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_KOL_AMB, "n/d", SRC
    AssertTrue InStr(1, ZbirStavkiGreska(), "KolAmbalaze stavke nije brojcana", vbTextCompare) > 0, _
               "OTK siroce: nebrojcana KolAmbalaze obara citaoca (pisac je vec odbija)"

    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_KOL_AMB, -1#, SRC
    AssertTrue InStr(1, ZbirStavkiGreska(), "ne sme biti negativna", vbTextCompare) > 0, _
               "OTK siroce: negativna KolAmbalaze obara citaoca"

    RequireUpdateCell TBL_OTKUP_STAVKE, red, COL_OKS_KOL_AMB, 2.5, SRC
    AssertTrue InStr(1, ZbirStavkiGreska(), "mora biti ceo broj", vbTextCompare) > 0, _
               "OTK siroce: decimalna KolAmbalaze obara citaoca"

    ' CISCENJE JE DEO TESTA: pokvarena stavka obara svaki sledeci citalac.
    tx.RollbackTx
    Set tx = Nothing

    AssertEquals "", ZbirStavkiGreska(), _
                 "OTK siroce: citalac prolazi posle vracanja stavke"

    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFatal "Test_OTK_StavkaBezZaglavljaObaraCitaoce", errNum, errDesc
End Sub

' Poruka greske koju razresenje bloka u banci podigne, ili "" kad prodje.
Private Function BankaKandidatGreska(ByVal koopID As String, _
                                     ByVal brDok As String) As String
    Dim s As String
    On Error Resume Next
    Err.Clear
    s = modBankaMapiranje.BimOtkupIzBroja(koopID, brDok)
    If Err.Number <> 0 Then BankaKandidatGreska = Err.description
    Err.Clear
    On Error GoTo 0
End Function

' PRAZAN OtkupID NA ZAGLAVLJU: dva citaoca, dva razlicita odgovora (FM-0021 #5).
'
' Takav red se NE moze vrednovati, ali se ne sme ni izgubiti iz liste za
' isplatu -- to je bas taj kvar (otvorena obaveza tiho izostane iz pregleda).
' Zato ga StavkeOtkupaRedovi NAMERNO pusta, a obara ga ZbirStavkiZaOtkup na
' MESTU UPOTREBE: mreza, izvestaj i kandidat bloka padaju po imenu, dok lista
' otvorenih obaveza red zadrzava i prepusta ga imenovanom vlasniku
' (BuildBlokIsplataList -> ERR_ISPLATA_PRAZAN_OTKUPID).
'
' ZASTO BAS OVAJ SLUCAJ: ovo je JEDINI test koji meri drugu branu samu za sebe.
' Za dokument bez stavki izvor pada ranije, pa sabotaza nad citaocem tamo ne
' pokazuje crveno -- dokaz.py je to i prijavio (NE OBARA NISTA).
Private Sub Test_OTK_ZaglavljeBezIDObaraCitaoce()
    Const SRC As String = "Test_OTK_ZaglavljeBezIDObaraCitaoce"
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String, brDok As String, otkID As String, dan As Date
    scenario = NewScenarioCode("OTKBID")
    brDok = TEST_PREFIX & "-OTK-BID-" & scenario
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK bez ID: dokument napravljen"
    If Len(otkID) = 0 Then Exit Sub
    dan = CDate(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_DATUM))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE

    ' REDOSLED JE DEO TESTA: prvo se brisu stavke, pa se prazni ID. Obrnuto bi
    ' stavka ostala siroce i pao bi IZVOR -- a ovde se meri druga brana.
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID)
    AssertTrue Not rows Is Nothing, "OTK bez ID: stavke nadjene"
    If rows Is Nothing Then Exit Sub

    Dim k As Long
    For k = rows.count To 1 Step -1
        RequireDeleteRow TBL_OTKUP_STAVKE, CLng(rows(k)), SRC
    Next k

    Dim hdr As Collection
    Set hdr = FindRows(TBL_OTKUP, COL_OTK_ID, otkID)
    AssertTrue Not hdr Is Nothing, "OTK bez ID: zaglavlje nadjeno"
    If hdr Is Nothing Then Exit Sub
    RequireUpdateCell TBL_OTKUP, CLng(hdr(1)), COL_OTK_ID, "", SRC

    ' Izvor NAMERNO cuti: red bez ID-a i bez stavki nije ni u recniku.
    AssertEquals "", ZbirStavkiGreska(), _
                 "OTK bez ID: izvor namerno ne pada na redu bez OtkupID-a"

    ' Druga brana pada, i to po imenu.
    AssertTrue InStr(1, MrezaGreska(brDok), "bez OtkupID-a", vbTextCompare) > 0, _
               "OTK bez ID: mreza pada po imenu, ne crta 0 kg"
    AssertTrue InStr(1, SaldoOMGreska(dan), "bez OtkupID-a", vbTextCompare) > 0, _
               "OTK bez ID: saldo OM pada po imenu, ne sabira nulu"
    AssertTrue InStr(1, BankaKandidatGreska(TEST_KOOP_ID, brDok), _
                     "bez OtkupID-a", vbTextCompare) > 0, _
               "OTK bez ID: kandidat bloka pada po imenu, ne knjizi uplatu kao avans"

    ' A lista otvorenih obaveza ga NE gubi i NE obara se (FM-0021 #5).
    AssertEquals "", OtvoreniGreska(), _
                 "OTK bez ID: lista otvorenih ne pada -- red bez ID-a se ne sme izgubiti"

    tx.RollbackTx
    Set tx = Nothing

    AssertEquals "", ZbirStavkiGreska(), "OTK bez ID: citalac prolazi posle vracanja"

    Exit Sub

EH:
    Dim errNum2 As Long, errDesc2 As String
    errNum2 = Err.Number
    errDesc2 = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFatal "Test_OTK_ZaglavljeBezIDObaraCitaoce", errNum2, errDesc2
End Sub

' DUPLI OtkupID: isti teret se ne sme izbrojati dvaput (review #334, drugi krug).
'
' Dokument-level citaoci iteriraju ZAGLAVLJA i za svako uzimaju zbir po ID-u
' (ReportSaldoOM, KPI, mreza, rang). Dva zaglavlja sa istim OtkupID zato daju
' 2x kg nad JEDNOM fizickom stavkom, a kad nose razlicit KooperantID, iste
' kilograme pripisu dvojici kooperanata. Centralni citalac zato pada PRE zbira.
'
' Kontrola meri i tacan broj: pre duplikata dan nosi 1000 kg i posle vracanja
' opet 1000 -- nikad 2000.
Private Sub Test_OTK_DupliOtkupIDObaraCitaoce()
    Const SRC As String = "Test_OTK_DupliOtkupIDObaraCitaoce"
    Const OCEK_KG As Double = 1000#
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String, brDok As String, otkID As String, dan As Date
    scenario = NewScenarioCode("OTKDUP")
    brDok = TEST_PREFIX & "-OTK-DUP-" & scenario
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(OCEK_KG, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK dupli ID: dokument napravljen"
    If Len(otkID) = 0 Then Exit Sub
    dan = CDate(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkID, COL_OTK_DATUM))

    ' Drugi dokument -- njegovo ZAGLAVLJE postaje duplikat prvog.
    Dim drugiID As String
    drugiID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-DUP2-" & scenario), _
                             OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(drugiID) > 0, "OTK dupli ID: drugi dokument napravljen"
    If Len(drugiID) = 0 Then Exit Sub

    ' Kontrola PRE: dan nosi tacno jedan dokument i njegovih 1000 kg.
    Dim r As Variant, u As Long
    r = ReportSaldoOM(TEST_ST_ID, dan, dan)
    AssertTrue IsArray(r), "OTK dupli ID: kontrola -- saldo OM postoji"
    If IsArray(r) Then
        u = UBound(r, 1)
        AssertTrue Abs(CDbl(r(u, 2)) - OCEK_KG) < 0.001, _
                   "OTK dupli ID: kontrola -- dan nosi 1000 kg pre duplikata"
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE

    ' REDOSLED JE DEO TESTA: prvo se brisu stavke drugog dokumenta, pa mu se
    ' zaglavlju upisuje TUDJ OtkupID. Obrnuto bi njegove stavke ostale siroce
    ' i pao bi drugi kvar (1908), pa test ne bi merio duplikat.
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, drugiID)
    AssertTrue Not rows Is Nothing, "OTK dupli ID: stavke drugog dokumenta nadjene"
    If rows Is Nothing Then Exit Sub

    Dim k As Long
    For k = rows.count To 1 Step -1
        RequireDeleteRow TBL_OTKUP_STAVKE, CLng(rows(k)), SRC
    Next k

    Dim hdr As Collection
    Set hdr = FindRows(TBL_OTKUP, COL_OTK_ID, drugiID)
    AssertTrue Not hdr Is Nothing, "OTK dupli ID: zaglavlje drugog dokumenta nadjeno"
    If hdr Is Nothing Then Exit Sub
    RequireUpdateCell TBL_OTKUP, CLng(hdr(1)), COL_OTK_ID, otkID, SRC
    ' Isti dan kao original -- bez kapije bi izvestaj tog dana vratio 2000 kg.
    RequireUpdateCell TBL_OTKUP, CLng(hdr(1)), COL_OTK_DATUM, dan, SRC

    AssertEquals "2", CStr(OtkBrojZaglavlja(otkID)), _
                 "OTK dupli ID: preduslov -- dva zaglavlja nose isti OtkupID"
    AssertEquals "1", CStr(OtkBrojStavki(otkID)), _
                 "OTK dupli ID: preduslov -- stavka je i dalje jedna"

    ' Nijedan citalac ne sme da vrati zbir -- svi padaju i imenuju ID.
    AssertTrue InStr(1, ZbirStavkiGreska(), "ne nalazi tacno jednom", vbTextCompare) > 0, _
               "OTK dupli ID: zbir stavki pada po imenu"
    AssertTrue InStr(1, ZbirStavkiGreska(), otkID, vbTextCompare) > 0, _
               "OTK dupli ID: poruka imenuje sporan OtkupID"
    AssertTrue InStr(1, SaldoOMGreska(dan), "ne nalazi tacno jednom", vbTextCompare) > 0, _
               "OTK dupli ID: saldo OM pada umesto da vrati 2x kg"
    AssertTrue InStr(1, MrezaGreska(brDok), "ne nalazi tacno jednom", vbTextCompare) > 0, _
               "OTK dupli ID: mreza pada po imenu"
    AssertTrue InStr(1, OtvoreniGreska(), "ne nalazi tacno jednom", vbTextCompare) > 0, _
               "OTK dupli ID: lista za isplatu pada po imenu"

    tx.RollbackTx
    Set tx = Nothing

    ' Posle vracanja: opet 1000 kg, nikad 2000.
    AssertEquals "", ZbirStavkiGreska(), "OTK dupli ID: zbir prolazi posle vracanja"
    r = ReportSaldoOM(TEST_ST_ID, dan, dan)
    AssertTrue IsArray(r), "OTK dupli ID: saldo OM postoji posle vracanja"
    If IsArray(r) Then
        u = UBound(r, 1)
        AssertTrue Abs(CDbl(r(u, 2)) - OCEK_KG) < 0.001, _
                   "OTK dupli ID: posle vracanja dan opet nosi 1000 kg, ne 2000"
    End If

    Exit Sub

EH:
    Dim errNum3 As Long, errDesc3 As String
    errNum3 = Err.Number
    errDesc3 = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFatal "Test_OTK_DupliOtkupIDObaraCitaoce", errNum3, errDesc3
End Sub

' Koliko ZAGLAVLJA nosi dati OtkupID (2 = kvar identiteta).
Private Function OtkBrojZaglavlja(ByVal otkupID As String) As Long
    Dim redovi As Collection
    Set redovi = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If redovi Is Nothing Then Exit Function
    OtkBrojZaglavlja = redovi.count
End Function

' STATUS ISPLATE JE IZVEDEN, NE KESIRAN.
'
' UpdateOtkupStatus je odrzavao tblOtkup.Isplaceno i racunao vrednost kao
' Kolicina x Cena SA ZAGLAVLJA -- posle prelaska na stavke to je uvek nula, pa
' nijedan nov otkup ne bi nikad bio oznacen kao placen, tiho i bez greske.
'
' Sada listu otvorenih obaveza odlucuje sam novac. Test to i meri: dokument je
' otvoren, delimicna isplata ga ostavlja otvorenim, puna ga zatvara.
Private Sub Test_OTK_StatusIsplateJeIzveden()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKISP")

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-ISP-" & scenario), _
                           OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK isplata: dokument napravljen"

    AssertTrue OtkUOtvorenim(otkID), "OTK isplata: nov dokument je otvorena obaveza"

    ' Kolone Isplaceno/DatumIsplate vise NE POSTOJE (korak 7), pa se ni ne
    ' tvrde. Da se vrate, kapija je staticka: kanon i modSchema moraju u korak
    ' (gen_schema_module --check), a zatecena sveska pada na poziciji kolone.

    SaveNovac TEST_PREFIX & "-NOV-D-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, 4000#, "delimicno", otkID

    AssertTrue OtkUOtvorenim(otkID), _
               "OTK isplata: delimicna isplata NE zatvara obavezu"

    SaveNovac TEST_PREFIX & "-NOV-P-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, 6000#, "ostatak", otkID

    AssertTrue Not OtkUOtvorenim(otkID), _
               "OTK isplata: puna isplata zatvara obavezu"

    Exit Sub

EH:
    LogFatal "Test_OTK_StatusIsplateJeIzveden", Err.Number, Err.description
End Sub

Private Function OtkUOtvorenim(ByVal otkupID As String) As Boolean
    Dim r As Variant
    r = GetOpenOtkupi("")
    If Not IsArray(r) Then Exit Function

    Dim i As Long
    For i = 1 To UBound(r, 1)
        If StrComp(Trim$(CStr(r(i, 2))), otkupID, vbTextCompare) = 0 Then
            OtkUOtvorenim = True
            Exit Function
        End If
    Next i
End Function

' ISPRAVKA JE NOV DOKUMENT, NOV BROJ I VEZA PO ID-u (A9).
'
' Zatecen aparat je vezu drzao POSLOVNIM BROJEM (modStornoFlow.StampIspravkaTrace)
' -- pa dve verzije istog dokumenta nisu bile razlucive. Ovaj test tvrdi ono sto
' A9 trazi: nov ID, nov broj, IspravkaOdID na novom, ZamenjenSaID na starom, isti
' CorrectionID na oba, i CorrectionID koji STVARNO postoji u tblStornoVeze.
Private Sub Test_OTK_IspravkaNovDokumentINovBroj()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKISPR")

    Dim brStari As String, brNovi As String
    brStari = TEST_PREFIX & "-OTK-I1-" & scenario
    brNovi = TEST_PREFIX & "-OTK-I2-" & scenario

    Dim stariID As String
    stariID = CreateOtkup_TX(OtkHeader(brStari), OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(stariID) > 0, "Ispravka: stari dokument napravljen"

    Dim greska As String
    Dim noviID As String
    noviID = modOtkup.IspravkaOtkupa_TX(stariID, OtkHeader(brNovi), _
                                        OtkStavke(380#, 50#, 19, 0#, 0#, 0), greska)

    AssertTrue Len(noviID) > 0, "Ispravka: nov dokument nastao (" & greska & ")"
    AssertTrue StrComp(noviID, stariID, vbTextCompare) <> 0, "Ispravka: NOV OtkupID"

    ' Stari je stornirani, nov je aktivan -- jedna ziva verzija.
    AssertTrue RowIsStornirano(TBL_OTKUP, COL_OTK_ID, stariID), _
               "Ispravka: stari dokument je storniran"
    AssertTrue Not RowIsStornirano(TBL_OTKUP, COL_OTK_ID, noviID), _
               "Ispravka: nov dokument je aktivan"

    ' Veza na OBA kraja, po ID-u.
    AssertEquals stariID, OtkPolje(noviID, COL_TRACE_ISPRAVKA_OD_ID), _
                 "Ispravka: nov nosi IspravkaOdID"
    AssertEquals noviID, OtkPolje(stariID, COL_TRACE_ZAMENJEN_SA_ID), _
                 "Ispravka: stari nosi ZamenjenSaID"

    ' Isti CorrectionID, i to PRAV -- postoji u tblStornoVeze.
    Dim cid As String
    cid = OtkPolje(noviID, COL_TRACE_CORRECTION_ID)
    AssertTrue Len(cid) > 0, "Ispravka: nov nosi CorrectionID"
    AssertEquals cid, OtkPolje(stariID, COL_TRACE_CORRECTION_ID), _
                 "Ispravka: oba dokumenta nose ISTI CorrectionID"
    AssertTrue RowExists(TBL_STORNO_VEZE, COL_SV_ID, cid), _
               "Ispravka: CorrectionID postoji u tblStornoVeze (nije izmisljen)"

    ' Nov broj je stvarno nov i stoji na novom dokumentu.
    AssertEquals brNovi, OtkPolje(noviID, COL_OTK_BR_DOK), "Ispravka: nov broj na novom"
    AssertEquals brStari, OtkPolje(stariID, COL_OTK_BR_DOK), "Ispravka: stari broj netaknut"

    ' Stavke pripadaju NOVOM dokumentu; stari ih zadrzava (istorija je citljiva).
    AssertEquals "1", CStr(OtkBrojStavkiZaOtkup(noviID)), "Ispravka: nov ima svoju stavku"
    AssertEquals "1", CStr(OtkBrojStavkiZaOtkup(stariID)), _
                 "Ispravka: stari zadrzava svoju stavku (istorija se ne brise)"
    AssertTrue Abs(modOtkup.VrednostOtkupa(noviID) - 19000#) < 0.001, _
               "Ispravka: vrednost novog je 380 x 50"

    ' Poslednja verzija se cita bez rucnog pracenja lanca.
    AssertEquals noviID, modOtkup.PoslednjaVerzijaOtkupa(stariID), _
                 "Ispravka: PoslednjaVerzijaOtkupa vodi na naslednika"

    Exit Sub

EH:
    LogFatal "Test_OTK_IspravkaNovDokumentINovBroj", Err.Number, Err.description
End Sub

' TRI KAPIJE ISPRAVKE -- svaka po imenu.
'
' Kapije se ne mere padom nego PORUKOM: tri razlicita razloga koja daju istu
' poruku su jedna kapija sa tri ulaza, a ne tri kapije.
Private Sub Test_OTK_IspravkaKapije()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKISPK")

    Dim br1 As String: br1 = TEST_PREFIX & "-OTK-K1-" & scenario
    Dim br2 As String: br2 = TEST_PREFIX & "-OTK-K2-" & scenario
    Dim br3 As String: br3 = TEST_PREFIX & "-OTK-K3-" & scenario

    Dim id1 As String
    id1 = CreateOtkup_TX(OtkHeader(br1), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(id1) > 0, "Ispravka kapije: polazni dokument"

    ' 1) ISTI BROJ -- A9 trazi nov.
    Dim g As String
    Dim r As String
    r = modOtkup.IspravkaOtkupa_TX(id1, OtkHeader(br1), _
                                   OtkStavke(100#, 100#, 0, 0#, 0#, 0), g)
    AssertEquals "", r, "Ispravka kapije: isti broj je odbijen"
    AssertTrue InStr(1, g, "NOV broj", vbTextCompare) > 0, _
               "Ispravka kapije: poruka imenuje pravilo o broju (bilo: " & g & ")"
    AssertTrue Not RowIsStornirano(TBL_OTKUP, COL_OTK_ID, id1), _
               "Ispravka kapije: odbijena ispravka NIJE stornirala izvor"

    ' 2) DRUGI PUT nad istim dokumentom -- prvi je vec dao naslednika.
    Dim id2 As String
    id2 = modOtkup.IspravkaOtkupa_TX(id1, OtkHeader(br2), _
                                     OtkStavke(90#, 100#, 0, 0#, 0#, 0), g)
    AssertTrue Len(id2) > 0, "Ispravka kapije: prva ispravka prosla"

    r = modOtkup.IspravkaOtkupa_TX(id1, OtkHeader(br3), _
                                   OtkStavke(80#, 100#, 0, 0#, 0#, 0), g)
    AssertEquals "", r, "Ispravka kapije: druga ispravka istog dokumenta odbijena"
    AssertTrue InStr(1, g, "vec zamenjen", vbTextCompare) > 0, _
               "Ispravka kapije: poruka imenuje postojeceg naslednika (bilo: " & g & ")"

    ' 3) STORNIRAN IZVOR -- to nije ispravka nego nov unos.
    Dim id3 As String
    id3 = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-K4-" & scenario), _
                         OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue modStorno.StornoOtkup_TX(id3), "Ispravka kapije: izvor storniran"

    r = modOtkup.IspravkaOtkupa_TX(id3, OtkHeader(TEST_PREFIX & "-OTK-K5-" & scenario), _
                                   OtkStavke(100#, 100#, 0, 0#, 0#, 0), g)
    AssertEquals "", r, "Ispravka kapije: storniran izvor odbijen"
    AssertTrue InStr(1, g, "vec storniran", vbTextCompare) > 0, _
               "Ispravka kapije: poruka imenuje storno (bilo: " & g & ")"

    ' Lanac ispravki: id1 -> id2 je jedina veza, treca nije nastala.
    AssertEquals id2, modOtkup.PoslednjaVerzijaOtkupa(id1), _
                 "Ispravka kapije: lanac ima tacno jednog naslednika"

    Exit Sub

EH:
    LogFatal "Test_OTK_IspravkaKapije", Err.Number, Err.description
End Sub

' BANKA VIDI NOV OTKUP KAO OTVOREN BLOK.
'
' Otvoreni iznos bloka se racunao kao Kolicina * Cena SA ZAGLAVLJA. Nov pisac te
' kolone ne puni (od S1d ih u semi i nema), pa je vrednost ostajala 0 i blok je
' izgledao placen.
'
' Posledica nije kozmeticka: BimOtkupBezOtvorenog tada vrati True, pa se uplata
' knjizi kao AVANS umesto na blok, a stavka izvoda se oznaci obradjenom. Novac
' ode na pogresno mesto i niko ne dobije poruku.
'
' Postojeci banka testovi to ne vide jer koriste FX_BIM_BLOK* iz fixture-a --
' stari model, zaglavlje popunjeno. Ovaj test blok PRAVI kroz CreateOtkup_TX.
Private Sub Test_BIM_NovOtkupJeOtvorenBlok()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("BIMNOV")

    Dim br As String: br = TEST_PREFIX & "-BIM-" & scenario
    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(br), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "Banka nov blok: dokument iz NOVOG pisca napravljen"

    ' Broj iz izvoda se razresava u BAS taj dokument -- jednom, na granici (S2).
    AssertEquals otkID, modBankaMapiranje.BimOtkupIzBroja(TEST_KOOP_ID, br), _
                 "Banka nov blok: poziv na broj razresen u OtkupID dokumenta"

    ' I otvoreni iznos mora biti pun -- 10000, iz stavki.
    AssertTrue Abs(modBankaMapiranje.BimOtvorenoNaOtkupu(otkID) - 10000#) < 0.001, _
               "Banka nov blok: otvoreno je 10000 (iz stavki, ne sa zaglavlja)"

    ' Drugi smer: mapiranje ga NE sme videti kao blok bez otvorenog --
    ' to je tacka na kojoj bi uplata otisla u avans.
    AssertTrue Not modBankaMapiranje.BimOtkupBezOtvorenog(otkID), _
               "Banka nov blok: NIJE 'blok bez otvorenog' (inace uplata ide u avans)"

    Exit Sub
EH:
    LogFatal "Test_BIM_NovOtkupJeOtvorenBlok", Err.Number, Err.description
End Sub

Private Sub Test_OTK_IsplataNaNovDokumentProlazi()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKISPL")

    Dim br As String: br = TEST_PREFIX & "-OTK-IB-" & scenario
    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(br), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "IsplataBlok: dokument iz NOVOG pisca napravljen"

    ' Preduslov koji objasnjava ceo kvar: zaglavlje NEMA kolicinu.
    AssertTrue Abs(modOtkup.VrednostOtkupa(otkID) - 10000#) < 0.001, _
               "IsplataBlok: kanonska vrednost dolazi iz STAVKI (10000)"

    ' Isplata unutar vrednosti mora PROCI. Pre popravke je vracala
    ' "veci od ostatka", jer je vrednost sa zaglavlja bila 0.
    AssertEquals "", modNovac.IsplataBlokProblem(otkID, TEST_KOOP_ID, "", 5000#), _
                 "IsplataBlok: isplata unutar vrednosti NIJE odbijena"

    ' Drugi smer: preko vrednosti se i dalje odbija -- kapija nije ukinuta.
    AssertTrue Len(modNovac.IsplataBlokProblem(otkID, TEST_KOOP_ID, "", 15000#)) > 0, _
               "IsplataBlok: isplata PREKO vrednosti se i dalje odbija"

    Exit Sub
EH:
    LogFatal "Test_OTK_IsplataNaNovDokumentProlazi", Err.Number, Err.description
End Sub

' PUT D SUZENO: odvezan VirmanFirmaKoop se VIDI kao raspoloziv avans.
'
' Obican storno (bez ispravke) skida OtkupID. Do odluke 13.09.2026. su ga posle
' toga videli samo citaci VirmanAvansKoop-a, pa je VirmanFirmaKoop ispadao iz
' svake masinerije koja bira sta se placa. Sada ga vide -- ali KesOtkupacKoop i
' dalje NE, jer kes na otkupnom mestu nije avans nego zatvoren posao.
Private Sub Test_OTK_OdvezanVirmanJeRaspolozivAvans()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKAVD")

    Dim pre As Double
    pre = GetKooperantUnallocatedAvans(TEST_KOOP_ID)

    Dim br As String: br = TEST_PREFIX & "-OTK-AV-" & scenario
    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(br), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "Odvezan virman: polazni dokument"

    SaveNovac TEST_PREFIX & "-NOV-AV-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, 7000#, "virman firme na blok", otkID

    ' Dok je VEZAN, nije raspoloziv -- inace bi se dvaput trosio.
    AssertTrue Abs(GetKooperantUnallocatedAvans(TEST_KOOP_ID) - pre) < 0.001, _
               "Odvezan virman: dok je vezan za blok NIJE raspoloziv"

    AssertTrue modStorno.StornoOtkup(otkID), "Odvezan virman: storno prosao"

    ' Posle storna JESTE raspoloziv. Pre odluke je ovde bilo 0.
    AssertTrue Abs(GetKooperantUnallocatedAvans(TEST_KOOP_ID) - pre - 7000#) < 0.001, _
               "Odvezan virman: posle storna JESTE raspoloziv avans"

    Exit Sub
EH:
    LogFatal "Test_OTK_OdvezanVirmanJeRaspolozivAvans", Err.Number, Err.description
End Sub

' PUT A meri se KESOM, ne virmanom -- i to je nalaz, ne detalj.
'
' Prva verzija ovog dokaza je koristila VirmanFirmaKoop i NIJE merila prenos:
' sabotaza koja ukloni PrevezaNovacNaOtkup nije oborila nijednu tvrdnju
' (mereno 13.09.2026). Razlog nije previd u testu nego arhitektura -- put D
' suzeno cini VirmanFirmaKoop vidljivim kao avans, pa ga ApplyAvansToOtkup
' unutar CreateOtkup sam pokupi i veze za nov dokument. Za taj tip je prenos
' SUVISAN.
'
' Put A ima posmatriv efekat samo na KesOtkupacKoop, koji D namerno iskljucuje:
' kes placen za TAJ posao treba da prati ispravljen dokument, ali ne sme
' slobodno da pluta na tudje dokumente.
' PREPLATA VIRMANOM -- interakcija A + D, koju nijedan raniji test nije merio.
'
' Test preplate kesom dokazuje put A izolovano (kes ne ulazi u avans-petlju), i
' bas zato NE meri ovaj slucaj. Kod virmana rade OBA mehanizma i sudaraju se:
'
'   1. NovacIDsZaOtkup zapamti ORIGINALNI NovacID od 10.000
'   2. storno skine OtkupID
'   3. CreateOtkup napravi nov dokument od 8.000
'   4. ApplyAvansToOtkup vidi virman kao avans (put D suzeno) i posto je
'      10.000 > 8.000, DELI ga: original smanji na 2.000, a za primenjenih
'      8.000 napravi NOV red vezan za nov dokument
'   5. PrevezaNovacNaOtkup radi nad ZAPAMCENIM originalnim ID-em i prenese
'      preostalih 2.000
'
' Konacno: nov dokument ima 10.000 placeno na dug od 8.000. Prva verzija koda je
' upozorenje racunala kao `preneto > dug` -- a `preneto` je u tom trenutku bilo
' samo 2.000, pa uslov nije opalio i preplata je prosla TIHO. Nadjeno u recenziji
' 13.09.2026; upozorenje se sada racuna iz konacnog stanja dokumenta.
Private Sub Test_OTK_IspravkaPrijavljujePreplatuVirmanom()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPRV")

    Dim brStari As String: brStari = TEST_PREFIX & "-OTK-V1-" & scenario
    Dim brNovi As String: brNovi = TEST_PREFIX & "-OTK-V2-" & scenario

    Dim preAvans As Double
    preAvans = GetKooperantUnallocatedAvans(TEST_KOOP_ID)

    Dim stariID As String
    stariID = CreateOtkup_TX(OtkHeader(brStari), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(stariID) > 0, "Preplata virman: polazni dokument (10000)"

    SaveNovac TEST_PREFIX & "-NOV-V-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, 10000#, "virman pre ispravke", stariID

    Dim g As String, g2 As String
    Dim noviID As String
    noviID = modOtkup.IspravkaOtkupa_TX(stariID, OtkHeader(brNovi), _
                                        OtkStavke(80#, 100#, 0, 0#, 0#, 0), g, g2)

    AssertTrue Len(noviID) > 0, "Preplata virman: ispravka PROLAZI (" & g & ")"

    ' Sav novac je zavrsio na novom dokumentu -- i split deo i ostatak.
    AssertTrue Abs(modNovac.GetIsplataForOtkup(noviID) - 10000#) < 0.001, _
               "Preplata virman: nov dokument drzi SVIH 10000 (split + ostatak)"

    ' Nista nije ostalo da pluta kao slobodan avans kooperanta.
    AssertTrue Abs(GetKooperantUnallocatedAvans(TEST_KOOP_ID) - preAvans) < 0.001, _
               "Preplata virman: nema slobodnog ostatka starog placanja"

    ' I operater to MORA da sazna, sa iznosom.
    AssertTrue Len(Trim$(g2)) > 0, "Preplata virman: operater dobija upozorenje"
    AssertTrue InStr(1, g2, "2.000,00", vbTextCompare) > 0 _
               Or InStr(1, g2, "2,000.00", vbTextCompare) > 0, _
               "Preplata virman: upozorenje imenuje 2000 (bilo: " & g2 & ")"

    Exit Sub
EH:
    LogFatal "Test_OTK_IspravkaPrijavljujePreplatuVirmanom", Err.Number, Err.description
End Sub

Private Sub Test_OTK_IspravkaPrenosiKes()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKKES")

    Dim brStari As String: brStari = TEST_PREFIX & "-OTK-K1-" & scenario
    Dim brNovi As String: brNovi = TEST_PREFIX & "-OTK-K2-" & scenario

    Dim stariID As String
    stariID = CreateOtkup_TX(OtkHeader(brStari), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(stariID) > 0, "Ispravka kes: polazni dokument (vrednost 10000)"

    SaveNovac TEST_PREFIX & "-NOV-K-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_KES_OTKUPAC_KOOP, 0#, 10000#, "kes na otkupnom mestu", stariID

    AssertTrue Abs(GetIsplataForOtkup(stariID) - 10000#) < 0.001, _
               "Ispravka kes: stari dokument je placen kesom"

    Dim g As String, g2 As String
    Dim noviID As String
    noviID = modOtkup.IspravkaOtkupa_TX(stariID, OtkHeader(brNovi), _
                                        OtkStavke(100#, 100#, 0, 0#, 0#, 0), g, g2)
    AssertTrue Len(noviID) > 0, "Ispravka kes: ispravka prosla (" & g & ")"

    ' OVO meri put A: kes ne moze da stigne kroz avans-masineriju, jer ga
    ' JeAvansKooperanta namerno iskljucuje. Ako je ovde 10000, preneo ga je
    ' PrevezaNovacNaOtkup i nista drugo.
    AssertTrue Abs(GetIsplataForOtkup(noviID) - 10000#) < 0.001, _
               "Ispravka kes: nov dokument PREUZIMA kes (put A, ne avans-petlja)"

    ' I dalje nije slobodan avans -- kes to nikad ne postaje.
    AssertTrue Abs(GetKooperantUnallocatedAvans(TEST_KOOP_ID)) < 0.001, _
               "Ispravka kes: kes NIJE postao slobodan avans"

    Exit Sub
EH:
    LogFatal "Test_OTK_IspravkaPrenosiKes", Err.Number, Err.description
End Sub

' PREPLATA SE PRIJAVLJUJE, ISPRAVKA PROLAZI (odluka 13.09.2026).
' Ispravka smanjuje 100 kg na 80 kg, a placeno je punih 10000 -- razlika od 2000
' mora stici operateru kroz outUpozorenje, a nov dokument mora nastati.
Private Sub Test_OTK_IspravkaPrijavljujePreplatu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPRE")

    Dim brStari As String: brStari = TEST_PREFIX & "-OTK-P1-" & scenario
    Dim brNovi As String: brNovi = TEST_PREFIX & "-OTK-P2-" & scenario

    Dim stariID As String
    stariID = CreateOtkup_TX(OtkHeader(brStari), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(stariID) > 0, "Preplata: polazni dokument (10000)"

    SaveNovac TEST_PREFIX & "-NOV-P-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_KES_OTKUPAC_KOOP, 0#, 10000#, "kes pre ispravke", stariID

    Dim g As String, g2 As String
    Dim noviID As String
    noviID = modOtkup.IspravkaOtkupa_TX(stariID, OtkHeader(brNovi), _
                                        OtkStavke(80#, 100#, 0, 0#, 0#, 0), g, g2)

    AssertTrue Len(noviID) > 0, "Preplata: ispravka PROLAZI, ne blokira se (" & g & ")"
    AssertTrue Len(Trim$(g2)) > 0, "Preplata: operater dobija upozorenje"
    AssertTrue InStr(1, g2, "2.000,00", vbTextCompare) > 0 _
               Or InStr(1, g2, "2,000.00", vbTextCompare) > 0, _
               "Preplata: upozorenje imenuje IZNOS razlike (bilo: " & g2 & ")"

    Exit Sub
EH:
    LogFatal "Test_OTK_IspravkaPrijavljujePreplatu", Err.Number, Err.description
End Sub

Private Sub Test_OTK_IspravkaNeGubiNovac()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKISPN")

    Dim brStari As String: brStari = TEST_PREFIX & "-OTK-N1-" & scenario
    Dim brNovi As String: brNovi = TEST_PREFIX & "-OTK-N2-" & scenario

    Dim stariID As String
    stariID = CreateOtkup_TX(OtkHeader(brStari), OtkStavke(100#, 100#, 0, 0#, 0#, 0))
    AssertTrue Len(stariID) > 0, "Ispravka novac: polazni dokument (vrednost 10000)"

    SaveNovac TEST_PREFIX & "-NOV-I-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, 10000#, "isplata pre ispravke", stariID

    AssertTrue Abs(GetIsplataForOtkup(stariID) - 10000#) < 0.001, _
               "Ispravka novac: stari dokument je placen"

    Dim g As String
    Dim noviID As String
    Dim g2 As String
    noviID = modOtkup.IspravkaOtkupa_TX(stariID, OtkHeader(brNovi), _
                                        OtkStavke(100#, 100#, 0, 0#, 0#, 0), g, g2)
    AssertTrue Len(noviID) > 0, "Ispravka novac: ispravka prosla (" & g & ")"

    ' Stari vise nema vezan novac -- veza je skinuta, ne stornirana.
    AssertTrue Abs(GetIsplataForOtkup(stariID)) < 0.001, _
               "Ispravka novac: stari dokument vise ne drzi isplatu"

    ' PUT A (odluka 13.09.2026): nov dokument PREUZIMA placeni iznos.
    ' Do te odluke je ovde stajala obrnuta tvrdnja -- test je namerno merio
    ' ZATECENO stanje, da nalaz ne zivi u komentaru. Sada je okrenut, ne obrisan.
    AssertTrue Abs(GetIsplataForOtkup(noviID) - 10000#) < 0.001, _
               "Ispravka novac: nov dokument PREUZIMA odvezanu isplatu"

    ' Nije ni slobodan avans -- ali sada zato sto je VEZAN za naslednika, a ne
    ' zato sto je ispao iz svake masinerije. Ista tvrdnja, drugi razlog.
    AssertTrue Abs(GetKooperantUnallocatedAvans(TEST_KOOP_ID)) < 0.001, _
               "Ispravka novac: preneta isplata nije slobodan avans (vezana je)"

    ' Zato nov dokument vise NIJE otvorena obaveza -- placen je.
    AssertTrue Not OtkUOtvorenim(noviID), _
               "Ispravka novac: nov dokument NIJE otvorena obaveza (placen je)"

    ' Ista vrednost pre i posle -- nema preplate, pa nema ni upozorenja.
    AssertTrue Len(Trim$(g2)) = 0, _
               "Ispravka novac: bez smanjenja iznosa nema upozorenja o preplati (bilo: " & g2 & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_IspravkaNeGubiNovac", Err.Number, Err.description
End Sub

' A13: OTKUP KOJI IMA RODITELJA SE NE ISPRAVLJA -- NI IZDATOG, NI DRAFT.
'
' Ranija verzija ovog testa je DRAFT roditelja pustala kroz, uz obrazlozenje da je
' clanstvo drafta mutabilno. Merenje iz review-a je pokazalo da to nije dovoljno:
' IspravkaOtkupa_TX ne dira tblOtpremnicaIzvori, pa bi draft ostao sa izvorom koji
' pokazuje na STORNIRAN otkup, dok naslednik stoji van njega.
'
' To je isto medjustanje koje je PR5 vec odbio kod UpdateOtpremnicaDraft_TX:
' invarijanta mora da vazi IZMEDJU dva klika, ne tek pri izdavanju. Test zato sada
' meri OBA stanja kao ODBIJENA, i u oba slucaja tvrdi da je odbijanje POTPUNO.
Private Sub Test_OTK_IspravkaRoditeljFailClosed()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKA13")

    ' --- DRAFT roditelj ---
    Dim otkD As String
    otkD = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-A13D-" & scenario), _
                          OtkStavke(100#, 100#, 10, 0#, 0#, 0))
    AssertTrue Len(otkD) > 0, "A13: otkup za draft scenario"

    Dim gD As String
    Dim draftID As String
    draftID = CreateOtpremnicaDraft_TX(OtpHeader(TEST_PREFIX & "-OTP-A13D-" & scenario), _
                                       OtpOcek(100#, 10#, 0#, 0#), gD)
    AssertTrue Len(draftID) > 0, "A13: draft otpremnica napravljena (" & gD & ")"

    AssertTrue DodajOtpremnicaIzvor_TX(draftID, otkD, gD), _
               "A13: otkup je clan drafta (" & gD & ")"
    AssertTrue Not modDokumenta.OtpremnicaJeIzdata(draftID), "A13: draft nije izdat"

    IspravkaOdbijena otkD, draftID, "DRAFT", scenario & "-D"

    ' Clanstvo je NETAKNUTO -- draft i dalje pokazuje na ISTI, aktivan otkup.
    AssertEquals draftID, modDokumenta.OtpremnicaZaOtkup(otkD), _
                 "A13: draft i dalje ima svoj izvor"

    ' --- IZDATA otpremnica ---
    Dim otkI As String
    otkI = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-A13I-" & scenario), _
                          OtkStavke(200#, 100#, 20, 0#, 0#, 0))
    AssertTrue Len(otkI) > 0, "A13: otkup za izdat scenario"

    Dim izvori As Collection
    Set izvori = New Collection
    izvori.Add otkI

    Dim gI As String
    Dim izdataID As String
    izdataID = CreateOtpremnicaIzIzvora_TX(OtpHeader(TEST_PREFIX & "-OTP-A13I-" & scenario), _
                                           izvori, gI)
    AssertTrue Len(izdataID) > 0, "A13: izdata otpremnica napravljena (" & gI & ")"
    AssertTrue modDokumenta.OtpremnicaJeIzdata(izdataID), "A13: ta otpremnica JESTE izdata"

    IspravkaOdbijena otkI, izdataID, "IZDATA", scenario & "-I"

    Exit Sub

EH:
    LogFatal "Test_OTK_IspravkaRoditeljFailClosed", Err.Number, Err.description
End Sub

' Ispravka otkupa sa roditeljem mora biti odbijena POTPUNO -- i poruka mora da
' imenuje otpremnicu, a NE da nudi obilazak.
Private Sub IspravkaOdbijena(ByVal otkupID As String, ByVal otpID As String, _
                             ByVal stanje As String, ByVal scenario As String)
    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim g As String
    Dim r As String
    r = modOtkup.IspravkaOtkupa_TX(otkupID, OtkHeader(TEST_PREFIX & "-OTK-A13X-" & scenario), _
                                   OtkStavke(90#, 100#, 9, 0#, 0#, 0), g)

    AssertEquals "", r, "A13 " & stanje & ": ispravka je ODBIJENA"
    AssertTrue InStr(1, g, otpID, vbTextCompare) > 0, _
               "A13 " & stanje & ": poruka imenuje BAS tu otpremnicu (bilo: " & g & ")"
    AssertTrue InStr(1, g, "nije dostupna do PR7", vbTextCompare) > 0, _
               "A13 " & stanje & ": poruka upucuje na PR7"

    ' KAPIJA NAD PORUKOM: ranija verzija je govorila "storniraj otpremnicu pa
    ' ponovi" -- uputstvo za obilazak same kapije, jer OtpremnicaZaOtkup gleda samo
    ' AKTIVNE otpremnice. Test to sada zabranjuje po tekstu.
    AssertTrue InStr(1, g, "Storniraj otpremnicu", vbTextCompare) = 0, _
               "A13 " & stanje & ": poruka NE nudi obilazak preko storna roditelja"

    ' Odbijanje je potpuno: ni nov red, ni storniran izvor, ni naslednik.
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "A13 " & stanje & ": nijedan red nije upisan"
    AssertTrue Not RowIsStornirano(TBL_OTKUP, COL_OTK_ID, otkupID), _
               "A13 " & stanje & ": izvor je ostao aktivan"
    AssertEquals "", OtkPolje(otkupID, COL_TRACE_ZAMENJEN_SA_ID), _
                 "A13 " & stanje & ": izvor nije dobio naslednika"
End Sub

' JEDNA TRANSAKCIJA -- dokazano padom IZMEDJU storna i naslednika.
'
' IspravkaOtkupa_TX prvo stornira stari dokument, pa tek onda pravi nov. Ako pisac
' novog padne, sve mora nazad. Do review-a #308 to NIJE bilo tacno: transakcija
' nije snapshotovala tblStornoZurnal, u koji StornoOtkup pise kroz JournalCell --
' pa je posle rollback-a ostajao zapis storna koji se nije desio, i "Ponisti
' storno" bi nudio operaciju nad dokumentom koji je i dalje aktivan.
'
' Pad se izaziva BEZ test seam-a: stavka sa cenom nula prolazi sve kapije ispravke
' (roditelj, naslednik, storno, nov broj) i pada tek u CreateOtkup (modOtkup:261).
' Seam bi merio granu koja u pogonu ne postoji; ovako pada pravi pisac.
Private Sub Test_OTK_IspravkaRollbackVracaSve()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKROLL")

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-RB-" & scenario), _
                           OtkStavke(100#, 100#, 10, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "Rollback: polazni dokument"

    ' Vezan novac -- da rollback ima sta da vrati i na toj strani.
    SaveNovac TEST_PREFIX & "-NOV-RB-" & scenario, NextTestDate(), _
              "TEST KOOPERANT", TEST_KOOP_ID, "Kooperant", _
              "", TEST_KOOP_ID, "", "", _
              NOV_VIRMAN_FIRMA_KOOP, 0#, 10000#, "pre rollbacka", otkID

    Dim preH As Long: preH = OtkBrojRedova(TBL_OTKUP)
    Dim preS As Long: preS = OtkBrojRedova(TBL_OTKUP_STAVKE)
    Dim preA As Long: preA = OtkBrojRedova(TBL_AMBALAZA)
    Dim preV As Long: preV = OtkBrojRedova(TBL_STORNO_VEZE)
    Dim preZ As Long: preZ = OtkBrojRedova(TBL_STORNO_ZURNAL)
    Dim preN As Double: preN = GetIsplataForOtkup(otkID)

    AssertTrue Abs(preN - 10000#) < 0.001, "Rollback: novac je vezan pre pada"

    ' Cena nula -> pisac odbija stavku, ali TEK POSLE storna starog dokumenta.
    Dim g As String
    Dim r As String
    r = modOtkup.IspravkaOtkupa_TX(otkID, OtkHeader(TEST_PREFIX & "-OTK-RB2-" & scenario), _
                                   OtkStavke(90#, 0#, 9, 0#, 0#, 0), g)

    AssertEquals "", r, "Rollback: ispravka je pala"
    AssertTrue InStr(1, g, "Cena mora biti veca od nule", vbTextCompare) > 0, _
               "Rollback: pala je BAS na stavci (bilo: " & g & ")"

    ' --- sve mora biti kao pre ---
    AssertTrue Not RowIsStornirano(TBL_OTKUP, COL_OTK_ID, otkID), _
               "Rollback: stari dokument je opet AKTIVAN"
    AssertEquals "", OtkPolje(otkID, COL_TRACE_ZAMENJEN_SA_ID), _
                 "Rollback: stari nema naslednika"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "Rollback: nema novog otkup reda"
    AssertEquals CStr(preS), CStr(OtkBrojRedova(TBL_OTKUP_STAVKE)), _
                 "Rollback: nema novih stavki"
    AssertEquals CStr(preA), CStr(OtkBrojRedova(TBL_AMBALAZA)), _
                 "Rollback: ambalaza vracena"
    AssertEquals CStr(preV), CStr(OtkBrojRedova(TBL_STORNO_VEZE)), _
                 "Rollback: tblStornoVeze bez ostatka"
    AssertEquals CStr(preZ), CStr(OtkBrojRedova(TBL_STORNO_ZURNAL)), _
                 "Rollback: tblStornoZurnal bez ostatka (fantom zapis storna)"
    AssertTrue Abs(GetIsplataForOtkup(otkID) - 10000#) < 0.001, _
               "Rollback: novac je i dalje vezan za stari dokument"

    ' Dokument je i dalje upotrebljiv -- rollback ga nije ostavio polu-mrtvog.
    AssertTrue Abs(modOtkup.VrednostOtkupa(otkID) - 10000#) < 0.001, _
               "Rollback: vrednost starog dokumenta netaknuta"

    Exit Sub

EH:
    LogFatal "Test_OTK_IspravkaRollbackVracaSve", Err.Number, Err.description
End Sub

' SELF-HEAL MIGRACIJE KOLONA -- destruktivan put mora da ima meru.
'
' Kanon je u koraku 5 PREIMENOVAO dve kolone na tblOtkup, a u koraku 7 OBRISAO
' druge dve. Zatecena sveska mora da prati: kolona koja ostane u SREDINI pomera
' sve iza sebe, a AppendRow pise POZICIONO -- vrednosti bi tiho otisle u pogresna
' polja. modSchema to hvata i staje, ali sveska ostaje neupotrebljiva.
'
' Fixture je migriran generatorom (make_fixture: DROP_COLS / RENAME_COLS), pa
' VBA put do ovog testa nije izvrsavao NIKO. Kod koji brise kolonu, a nikad nije
' pokrenut, je najgora vrsta nemerene odbrane.
'
' Test radi nad kolonom koju sam doda NA KRAJ tabele: visak na kraju ne pomera
' nijednu kanonsku poziciju (otisak se racuna nad kanonskim prefiksom), pa ni
' pad testa ne ostavlja svesku u losem stanju.
Private Sub Test_OTK_SelfHealMigracijeKolona()
    On Error GoTo EH

    Const PROBA As String = "ZZTestKolonaProba"
    Const PROBA2 As String = "ZZTestKolonaProbaID"

    Dim lo As ListObject
    Set lo = GetTable(TBL_OTKUP)
    AssertTrue Not lo Is Nothing, "SelfHeal: tblOtkup postoji"

    Dim preKolona As Long
    preKolona = lo.ListColumns.count

    ' --- PREIMENOVANJE ---
    lo.ListColumns.Add().name = PROBA
    AssertTrue GetColumnIndex(TBL_OTKUP, PROBA) > 0, "SelfHeal: proba kolona dodata"

    modSetup.PreimenujKolonuAko TBL_OTKUP, PROBA, PROBA2

    AssertEquals "0", CStr(GetColumnIndex(TBL_OTKUP, PROBA)), _
                 "SelfHeal: staro ime vise ne postoji"
    AssertTrue GetColumnIndex(TBL_OTKUP, PROBA2) > 0, _
               "SelfHeal: novo ime postoji"
    AssertEquals CStr(preKolona + 1), CStr(GetTable(TBL_OTKUP).ListColumns.count), _
                 "SelfHeal: preimenovanje NE dodaje kolonu"

    ' Idempotentno: drugi prolaz nema sta da radi i ne sme da pogazi.
    modSetup.PreimenujKolonuAko TBL_OTKUP, PROBA, PROBA2
    AssertTrue GetColumnIndex(TBL_OTKUP, PROBA2) > 0, _
               "SelfHeal: drugi prolaz preimenovanja ne kvari nista"

    ' OBA IMENA ODJEDNOM -- stanje koje pravi medjuverzija.
    '
    ' Sveska koju je stara grana vec dopunila novim imenom (EnsureColumnOnTable
    ' dodaje NA KRAJ), a staro jos nosi. Bez kapije 'vec migrirano' preimenovanje
    ' bi napravilo DVE kolone istog imena -- Excel ih tada sam preimenuje u
    ' 'ime2' i pozicioni upis dobija polje koje niko ne trazi.
    '
    ' MERENO: gasenje kapije 'vec migrirano' NE obara ovaj test, i to je tacan
    ' rezultat -- Excel sam odbija drugu ListColumn istog imena, pa se ishod ne
    ' menja. Kapija stedi LogError na svakom startu takve sveske, ne podatak.
    ' Test zato tvrdi ISHOD (nema duplikata, nista se nije pomerilo), a razlog
    ' zbog kog kapija ipak stoji pise uz nju u modSetup.
    Dim loM As ListObject
    Set loM = GetTable(TBL_OTKUP)
    loM.ListColumns.Add().name = PROBA
    AssertTrue GetColumnIndex(TBL_OTKUP, PROBA) > 0, "SelfHeal: staro ime vraceno"
    AssertEquals CStr(preKolona + 2), CStr(GetTable(TBL_OTKUP).ListColumns.count), _
                 "SelfHeal: sada postoje OBA imena"

    modSetup.PreimenujKolonuAko TBL_OTKUP, PROBA, PROBA2

    AssertEquals CStr(preKolona + 2), CStr(GetTable(TBL_OTKUP).ListColumns.count), _
                 "SelfHeal: sa oba imena preimenovanje NE radi nista"
    AssertTrue GetColumnIndex(TBL_OTKUP, PROBA) > 0, _
               "SelfHeal: staro ime je netaknuto (nema duplikata)"

    modSetup.ObrisiKolonuAko TBL_OTKUP, PROBA
    AssertEquals CStr(preKolona + 1), CStr(GetTable(TBL_OTKUP).ListColumns.count), _
                 "SelfHeal: pomocna kolona sklonjena"

    ' --- BRISANJE ---
    modSetup.ObrisiKolonuAko TBL_OTKUP, PROBA2

    AssertEquals "0", CStr(GetColumnIndex(TBL_OTKUP, PROBA2)), _
                 "SelfHeal: kolona je obrisana"
    AssertEquals CStr(preKolona), CStr(GetTable(TBL_OTKUP).ListColumns.count), _
                 "SelfHeal: tabela je vracena na polazni broj kolona"

    ' Idempotentno i u drugom smeru.
    modSetup.ObrisiKolonuAko TBL_OTKUP, PROBA2
    AssertEquals CStr(preKolona), CStr(GetTable(TBL_OTKUP).ListColumns.count), _
                 "SelfHeal: brisanje nepostojece kolone ne dira tabelu"

    ' --- KANONSKE KOLONE SE NE DIRAJU ---
    ' Kapija je uska po imenu, ne po pravilu "sve sto nije u kanonu": modSetup
    ' legitimno dodaje kolone NA KRAJ pre nego sto ih kanon preuzme.
    modSetup.ObrisiKolonuAko TBL_OTKUP, "NemaOvakveKolone"
    AssertEquals CStr(preKolona), CStr(GetTable(TBL_OTKUP).ListColumns.count), _
                 "SelfHeal: nepoznato ime ne obara nijednu kanonsku kolonu"
    AssertTrue GetColumnIndex(TBL_OTKUP, COL_OTK_ID) > 0, _
               "SelfHeal: OtkupID je netaknut"

    Exit Sub

EH:
    ' Ciscenje i posle pada -- visak kolone ne sme da ostane iza testa.
    On Error Resume Next
    modSetup.ObrisiKolonuAko TBL_OTKUP, PROBA
    modSetup.ObrisiKolonuAko TBL_OTKUP, PROBA2
    On Error GoTo 0
    LogFatal "Test_OTK_SelfHealMigracijeKolona", Err.Number, Err.description
End Sub

' CRID KONFLIKT SE MERI I PO PARCELI I PO TIPU AMBALAZE.
'
' Poredjenje sadrzaja je prvo gledalo samo kooperanta, kulturu, datum i stavku.
' Nalaz iz review-a: isti ClientRecordID sa parcele P1 i sa parcele P2 prolazio je
' kao "isti sadrzaj" -- pa bi ispravljena parcela TIHO nestala. Isto za tip
' ambalaze, koji odlucuje ceo dvojni upis gajbi.
'
' Oba menjaju STA dokument tvrdi, pa oba moraju biti konflikt, ne no-op.
Private Sub Test_PWA_KonfliktPoParceliITipu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PWAPT")

    ' --- parcela ---
    Dim cridP As String
    cridP = TEST_PREFIX & "-CRID-PAR-" & scenario

    Dim redP As Variant
    redP = PwaRed(cridP, TEST_PREFIX & "-OTK-PWAPAR-" & scenario, 400#, 50#, 20)

    Dim prviP As String
    prviP = modMasterSync.ImportRowToTblOtkup_RowTX(redP, 1, cridP)
    AssertTrue Len(prviP) > 0, "PWA parcela: prvi uvoz prosao"

    ' PwaRed salje PRAZNU parcelu, pa prvi dokument nema parcelu. Drugi je salje.
    ' Oba su LEGITIMNE vrednosti za istog kooperanta -- test tako meri bas kapiju
    ' jednakosti, a ne FK proveru parcele (PAR-TEST-2 pripada drugom kooperantu).
    AssertEquals "", OtkPolje(prviP, COL_OTK_PARCELA), "PWA parcela: prvi je bez parcele"

    ' KOPIJA polaznog reda -- menja se TACNO jedno polje (v. zamku uz PwaRed).
    Dim izmenjenP As Variant
    izmenjenP = redP
    izmenjenP(1, 19) = GetTestParcelaID()              ' GS_PARCELA_ID

    Dim preP As Long: preP = OtkBrojRedova(TBL_OTKUP)
    AssertEquals "", modMasterSync.ImportRowToTblOtkup_RowTX(izmenjenP, 1, cridP), _
                 "PWA parcela: druga parcela pod istim CRID-om je ODBIJENA"
    AssertEquals CStr(preP), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "PWA parcela: nijedan nov red nije nastao"

    ' --- tip ambalaze ---
    Dim cridT As String
    cridT = TEST_PREFIX & "-CRID-TIP-" & scenario

    Dim redT As Variant
    redT = PwaRed(cridT, TEST_PREFIX & "-OTK-PWATIP-" & scenario, 400#, 50#, 20)

    Dim prviT As String
    prviT = modMasterSync.ImportRowToTblOtkup_RowTX(redT, 1, cridT)
    AssertTrue Len(prviT) > 0, "PWA tip: prvi uvoz prosao"

    Dim izmenjenT As Variant
    izmenjenT = redT
    ' TipAmbalaze NIJE FK (v. modOtkup:474 -- FK su kooperant, stanica, kultura,
    ' parcela), pa je drugi tip legitiman ulaz i kapija jednakosti je jedino sto
    ' ga moze odbiti.
    izmenjenT(1, 17) = TEST_TIP_AMB & "-DRUGI"         ' GS_TIP_AMB

    Dim preT As Long: preT = OtkBrojRedova(TBL_OTKUP)
    AssertEquals "", modMasterSync.ImportRowToTblOtkup_RowTX(izmenjenT, 1, cridT), _
                 "PWA tip: drugi tip ambalaze pod istim CRID-om je ODBIJEN"
    AssertEquals CStr(preT), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "PWA tip: nijedan nov red nije nastao"

    ' --- broj dokumenta: USLOVNO poredjenje ---
    '
    ' Prazan incoming broj se IGNORISE (master ga generise lokalno), a izricito
    ' poslat broj JESTE deo payload-a. Test meri obe strane tog uslova.
    Dim cridB As String
    cridB = TEST_PREFIX & "-CRID-BR-" & scenario

    Dim redB As Variant
    redB = PwaRed(cridB, TEST_PREFIX & "-OTK-PWABR-" & scenario, 400#, 50#, 20)

    ' Broj se gradi iz stanice i dana SAMOG REDA. Ranije su ovde stajali literali
    ' "77/090926" i "78/090926" -- oni tvrde stanicu 77, a red nosi ST-90001 i
    ' datum iz 2090. Test je time merio idempotenciju po CRID-u nad brojem koji
    ' je i pre kapije konteksta bio pogresan; sada ga modBrojevi kapija odbija na
    ' uvozu, pa bi tvrdnja o CRID-u prolazila iz pogresnog razloga.
    Dim datumB As Date
    datumB = CDate(redB(1, 9))                         ' GS_DATUM
    Dim brojB As String, brojB2 As String
    brojB = modBrojevi.FormatBroj(TEST_ST_ID, datumB, 1)
    brojB2 = modBrojevi.FormatBroj(TEST_ST_ID, datumB, 2)
    redB(1, 23) = brojB                                ' GS_BROJ_DOKUMENTA

    Dim prviB As String
    prviB = modMasterSync.ImportRowToTblOtkup_RowTX(redB, 1, cridB)
    AssertTrue Len(prviB) > 0, "PWA broj: prvi uvoz sa izricitim brojem prosao"
    AssertEquals brojB, OtkPolje(prviB, COL_OTK_BR_DOK), _
                 "PWA broj: izricit broj je zapisan"

    Dim izmenjenB As Variant
    izmenjenB = redB
    izmenjenB(1, 23) = brojB2

    Dim preB As Long: preB = OtkBrojRedova(TBL_OTKUP)
    AssertEquals "", modMasterSync.ImportRowToTblOtkup_RowTX(izmenjenB, 1, cridB), _
                 "PWA broj: drugi broj pod istim CRID-om je ODBIJEN"
    AssertEquals CStr(preB), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "PWA broj: nijedan nov red nije nastao"

    ' Kontrola: NEPROMENJEN red je i dalje no-op, ne konflikt -- i za red BEZ
    ' broja (master ga je generisao), sto dokazuje da uslov ne lomi taj put.
    AssertEquals prviT, modMasterSync.ImportRowToTblOtkup_RowTX(redT, 1, cridT), _
                 "PWA kontrola: nepromenjen sadrzaj je i dalje NO-OP"
    AssertEquals prviB, modMasterSync.ImportRowToTblOtkup_RowTX(redB, 1, cridB), _
                 "PWA kontrola: isti izricit broj je i dalje NO-OP"

    Exit Sub

EH:
    LogFatal "Test_PWA_KonfliktPoParceliITipu", Err.Number, Err.description
End Sub

' VREME NASTANKA NA TERENU SE PRENOSI, ne baca.
'
' PWA sema polje TRAZI (RequireOTKHeaderValue nad GS_CREATED_AT), a CreateOtkup_TX
' ga prima kao opcion header kljuc -- ali adapter ga nije prosledjivao. Bez njega
' je jedini vremenski trag CreatedAt, koji nosi trenutak SINHRONIZACIJE; posle
' prekida veze to ume da bude i nekoliko dana kasnije od stvarnog otkupa.
Private Sub Test_PWA_PrenosiVremeNastanka()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PWASCA")

    Dim crid As String
    crid = TEST_PREFIX & "-CRID-SCA-" & scenario

    Dim red As Variant
    red = PwaRed(crid, TEST_PREFIX & "-OTK-PWASCA-" & scenario, 400#, 50#, 20)

    Dim nastalo As String
    nastalo = "2026-08-14T06:30:00Z"
    red(1, 3) = nastalo                                ' GS_CREATED_AT

    Dim otkID As String
    otkID = modMasterSync.ImportRowToTblOtkup_RowTX(red, 1, crid)
    AssertTrue Len(otkID) > 0, "PWA vreme: uvoz prosao"

    AssertEquals nastalo, OtkPolje(otkID, COL_OTK_SOURCE_CREATED_AT), _
                 "PWA vreme: SourceCreatedAt nosi vreme sa terena"

    ' Bez prenosa bi polje ostalo prazno -- to je stanje koje je nalaz i opisao.
    AssertTrue Len(OtkPolje(otkID, COL_OTK_SOURCE_CREATED_AT)) > 0, _
               "PWA vreme: polje nije ostalo prazno"

    Exit Sub

EH:
    LogFatal "Test_PWA_PrenosiVremeNastanka", Err.Number, Err.description
End Sub

' IZVEDENI LANAC JE PAUZIRAN NA OBA PREOSTALA ULAZA (auto-otpremnica iz PWA je
' obrisana u S1c; vraca je S5).
'
' Nalaz iz review-a: pauzirana je bila samo auto-otpremnica, a nizvodni koraci su
' nastavljali -- i oba PISU NAZAD NA ZAGLAVLJE OTKUPA:
'
'   AutoCreateZbirnaFromOtpremnice  -> BackfillOtkupBrojZbirneByOtpremnica
'   ImportVOZRow_RowTX              -> LinkZbirnaToOtkupAndOtpremnica
'
' Pun PWA sync je time mogao da napravi canonical otkup, pa da ga odmah
' KONTAMINIRA starim backlink modelom. Kapija je zato JEDNA i pokriva ceo lanac;
' test tvrdi da nijedan ulaz ne prolazi, i to PO PORUCI.
Private Sub Test_PWA_IzvedeniLanacJePauziran()
    On Error GoTo EH

    AssertTrue Not modMasterSync.IzvedeniLanacIzPwaDostupan(), _
               "Lanac: kapija je zatvorena"

    ' 2) malina auto-zbirna iz otpremnica
    AssertTrue InStr(1, UlazPada("ZBR"), "PAUZIRANA", vbTextCompare) > 0, _
               "Lanac: auto-zbirna iz otpremnica je pauzirana"

    ' 3) VOZ/zbirna uvoz -- BACA sa svojom porukom.
    '
    ' Ranije je vracao False, pa se pauza nije razlikovala od "nema VOZ fajlova":
    ' sabotaza kapije nije obarala nista. Tvrdnja je zato po PORUCI.
    AssertTrue InStr(1, UlazPada("VOZ"), "PAUZIRAN", vbTextCompare) > 0, _
               "Lanac: VOZ/zbirna uvoz je pauziran"

    Exit Sub

EH:
    LogFatal "Test_PWA_IzvedeniLanacJePauziran", Err.Number, Err.description
End Sub

' Poruka greske sa ulaza koji mora biti pauziran, ili "" ako je prosao.
Private Function UlazPada(ByVal koji As String) As String
    On Error Resume Next
    Err.Clear

    Select Case koji
        Case "ZBR": Call modMasterSync.AutoCreateZbirnaFromOtpremnice_TX("NEMA-" & NewScenarioCode("LNC"))
        Case "VOZ": Call modMasterSync.ImportZbirneFromPWA_Core(False)
    End Select

    If Err.Number <> 0 Then UlazPada = Err.description
    Err.Clear
    On Error GoTo 0
End Function

Private Sub Test_OTK_VrednostBezStavkiPada()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKVS")

    ' Red bez stavki se pravi NAMERNO, pa se namerno i vraca: od review-a
    ' #334 on obara svakog citaoca vrednosti, ne samo kanon ispod.
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTKUP_STAVKE
    tx.AddTableSnapshot TBL_AMBALAZA
    ' CreateOtkup_TX zove ApplyAvansToOtkup: fixture sme da potrosi ili
    ' podeli slobodan avans kooperanta, pa i tblNovac mora nazad.
    tx.AddTableSnapshot TBL_NOVAC

    ' Zaglavlje BEZ stavki (sinteticka anomalija) -- oblik koji kapija mora da uhvati.
    Dim stariID As String
    stariID = OtkupBezStavkiFixture(TEST_PREFIX & "-OTK-VS-" & scenario)

    AssertTrue Len(stariID) > 0, "OTK vrednost: stari red napravljen"
    AssertEquals "0", CStr(OtkBrojStavkiZaOtkup(stariID)), _
                 "OTK vrednost: taj red zaista nema stavke"

    Dim greska As String
    Dim v As Double
    On Error Resume Next
    Err.Clear
    v = modOtkup.VrednostOtkupa(stariID)
    greska = Err.description
    If Err.Number = 0 Then greska = ""
    Err.Clear
    On Error GoTo EH

    AssertTrue InStr(1, greska, "nema nijednu stavku", vbTextCompare) > 0, _
               "OTK vrednost: kapija pada po imenu (bilo: " & greska & ")"

    ' Kontrola: dokument SA stavkama daje broj, ne gresku.
    Dim noviID As String
    noviID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-VS2-" & scenario), _
                            OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Abs(modOtkup.VrednostOtkupa(noviID) - 20000#) < 0.001, _
               "OTK vrednost: 400 x 50 = 20000"

    tx.RollbackTx
    Set tx = Nothing
    Exit Sub

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFatal "Test_OTK_VrednostBezStavkiPada", errNum, errDesc
End Sub

' STAMPA NE REKONSTRUISE ISTORIJSKI BRUTO.
'
' Zatecena stampa je za neto unos racunala bruto iz TRENUTNE tare gajbice. Test
' menja taru POSLE nastanka dokumenta -- to je dokaz koji obican assert ne daje:
' da isti istorijski dokument ne menja smisao kad se sifarnik promeni (S4.1d).
Private Sub Test_OTK_PrintNetoUnosNeRekonstruiseBruto()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPB")

    SeedTaraGajbice TEST_TIP_AMB, 2#

    ' NETO unos: 400 kg, 20 gajbi. Rekonstrukcija bi dala 400 + 20*2 = 440.
    Dim netoID As String
    netoID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-PB-" & scenario), _
                            OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(netoID) > 0, "OTK print bruto: neto otkup upisan"
    AssertEquals "", OtkStavkaPolje(netoID, KLASA_I, COL_OKS_BRUTO), _
                 "OTK print bruto: neto unos nema zamrznut bruto"

    ' BRUTO unos: 480 neto / 500 bruto.
    Dim brutoStavke As Collection
    Set brutoStavke = New Collection
    brutoStavke.Add OtkStavka(KLASA_I, 480#, 50#, 20#, 500#)

    Dim brutoID As String
    brutoID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-PB2-" & scenario), brutoStavke)
    AssertTrue Len(brutoID) > 0, "OTK print bruto: bruto otkup upisan"

    AssertTrue Not PrintListSadrzi(netoID, 440#), _
               "OTK print bruto: neto dokument NE stampa rekonstruisanih 440"
    AssertTrue PrintListSadrzi(brutoID, 500#), _
               "OTK print bruto: bruto dokument stampa zamrznutih 500"

    ' Tara se menja POSLE nastanka. Istorijski dokument ne sme da se pomeri.
    SeedTaraGajbice TEST_TIP_AMB, 3#

    AssertTrue Not PrintListSadrzi(netoID, 460#), _
               "OTK print bruto: promena tare ne pravi nov 'istorijski' bruto"
    AssertTrue Not PrintListSadrzi(netoID, 440#), _
               "OTK print bruto: ni stari rekonstruisani se ne vraca"
    AssertTrue PrintListSadrzi(brutoID, 500#), _
               "OTK print bruto: zamrznut bruto je nepromenjen posle izmene tare"

    Exit Sub

EH:
    LogFatal "Test_OTK_PrintNetoUnosNeRekonstruiseBruto", Err.Number, Err.description
End Sub

' Ispravka hladnjackog dokumenta je FAIL-CLOSED dok je lanac pauziran.
'
' Bez ove kapije bi nastao nov otkup, pending bi bio POTROSEN, a operater bi
' dobio samo "nema prijemnice" -- pola ispravke, i to nepovratno.
Private Sub Test_OTK_IspravkaPauziranaNeTrosiPending()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKIP")

    Dim stariPending As String
    stariPending = GetHladnjacaRelinkPending()
    SetHladnjacaRelinkPending TEST_PREFIX & "-PRJ-STARA-" & scenario

    Dim p As Object
    Set p = OtkEkranParam(TEST_PREFIX & "-OTK-IP-" & scenario)
    p("stanicaID") = TEST_HLAD_ST_ID
    p("kolicinaI") = 400#
    p("cenaI") = 50#
    p("kolAmb") = 20&

    Dim preH As Long
    preH = OtkBrojRedova(TBL_OTKUP)

    Dim poruke As String
    Dim res As String
    res = modOtkupUnos.OtkupUpisi(p, poruke)

    Dim pendingPosle As String
    pendingPosle = GetHladnjacaRelinkPending()
    SetHladnjacaRelinkPending stariPending          ' vrati stanje pre tvrdnji

    AssertEquals "", res, "OTK ispravka: otkup NIJE nastao"
    AssertEquals CStr(preH), CStr(OtkBrojRedova(TBL_OTKUP)), _
                 "OTK ispravka: nijedan red nije upisan"
    AssertEquals TEST_PREFIX & "-PRJ-STARA-" & scenario, pendingPosle, _
                 "OTK ispravka: pending NIJE potrosen"
    AssertTrue InStr(1, poruke, "nedostupna", vbTextCompare) > 0, _
               "OTK ispravka: operater je obavesten (bilo: " & poruke & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_IspravkaPauziranaNeTrosiPending", Err.Number, Err.description
End Sub

' Broj otkupnog lista je jedinstven po STANICI I DANU -- i to cuva PISAC.
'
' Zatecena provera je bila samo u UI-ju (modOtkupUnos:229) i nije gledala
' stanicu. Invarijanta koja zivi u UI-ju nije invarijanta nego navika: PWA ne
' prolazi kroz OtkupValidiraj.
Private Sub Test_OTK_BrojJedinstvenPoStaniciIDanu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKBJ")

    Dim broj As String
    broj = TEST_PREFIX & "-OTK-BJ-" & scenario

    Dim h1 As Object
    Set h1 = OtkHeader(broj)
    Dim datum As Date
    datum = h1("Datum")

    AssertTrue Len(CreateOtkup_TX(h1, OtkStavke(400#, 50#, 20, 0#, 0#, 0))) > 0, _
               "OTK broj: prvi dokument prosao"

    ' Isti broj, ista stanica, isti dan -> odbijeno.
    Dim h2 As Object
    Set h2 = OtkHeader(broj)
    h2("Datum") = datum

    Dim rez As String, razlog As String
    rez = CreateOtkup_TX(h2, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)

    AssertEquals "", rez, "OTK broj: duplikat na istoj stanici istog dana odbijen"
    AssertTrue InStr(1, razlog, "vec izdat", vbTextCompare) > 0, _
               "OTK broj: kapija imenuje razlog (bilo: " & razlog & ")"

    ' Isti broj, DRUGA stanica, isti dan -> prolazi. Generator skopira po stanici,
    ' pa dve stanice legitimno mogu imati isti redni broj istog dana.
    Dim h3 As Object
    Set h3 = OtkHeader(broj)
    h3("Datum") = datum
    h3("StanicaID") = TEST_HLAD_ST_ID

    AssertTrue Len(CreateOtkup_TX(h3, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)) > 0, _
               "OTK broj: druga stanica istog dana prolazi (bilo: " & razlog & ")"

    ' Isti broj, ista stanica, DRUGI dan -> prolazi.
    Dim h4 As Object
    Set h4 = OtkHeader(broj)
    h4("Datum") = DateAdd("d", 1, datum)

    AssertTrue Len(CreateOtkup_TX(h4, OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)) > 0, _
               "OTK broj: drugi dan na istoj stanici prolazi (bilo: " & razlog & ")"

    Exit Sub

EH:
    LogFatal "Test_OTK_BrojJedinstvenPoStaniciIDanu", Err.Number, Err.description
End Sub

' Tara gajbice u sifarniku -- upisuje se ili azurira.
Private Sub SeedTaraGajbice(ByVal tip As String, ByVal tezina As Double)
    Dim redovi As Collection
    Set redovi = FindRows(TBL_TIP_AMBALAZE, COL_TAMB_TIP, tip)
    If redovi.count > 0 Then
        RequireUpdateCell TBL_TIP_AMBALAZE, redovi(1), COL_TAMB_TEZINA, tezina, _
                          "SeedTaraGajbice"
        Exit Sub
    End If

    Dim rowData As Variant
    rowData = BlankRow(TBL_TIP_AMBALAZE)
    SetRequiredField rowData, TBL_TIP_AMBALAZE, COL_TAMB_TIP, tip
    SetRequiredField rowData, TBL_TIP_AMBALAZE, COL_TAMB_TEZINA, tezina
    SetOptionalField rowData, TBL_TIP_AMBALAZE, "Aktivan", "Aktivan"
    RequireAppend TBL_TIP_AMBALAZE, rowData, "SeedTaraGajbice"
End Sub

' Da li popunjen otkupni list igde sadrzi bas taj broj.
'
' Ne trazi se odredjena celija nego PRISUSTVO vrednosti -- tvrdnja je o tome sta
' dokument kaze, ne o geometriji sablona, pa test ne puca kad se sablon preuredi.
Private Function PrintListSadrzi(ByVal otkupID As String, ByVal broj As Double) As Boolean
    Dim ws As Worksheet
    Set ws = modPrint.FillOtkupSablon(otkupID)
    If ws Is Nothing Then Exit Function

    Dim c As Range
    For Each c In ws.UsedRange
        If IsNumeric(c.value) And Not IsEmpty(c.value) Then
            If Abs(CDbl(c.value) - broj) < 0.001 Then
                PrintListSadrzi = True
                Exit Function
            End If
        End If
    Next c
End Function

' Otkup po NOVOM modelu, za fixture nizvodnih testova.
'
' Zamenjuje SaveOtkupMulti_TX u testovima ciji SUBJEKT nije pisac otkupa nego
' nesto nizvodno: AutoLink, storno kaskada, hladnjacki lanac, SEF faktura. Otkup
' se zato pravi kanonski (jedan header + stavke), a BrojZbirne se ZIGOSE posle --
' nov pisac ga ne prima (broj nije veza, A2), ali kolona jos zivi i ti testovi je
' citaju. Odlazi u PR8, zajedno sa njihovim tvrdnjama.
Private Function NoviOtkupFixture(ByVal datum As Date, ByVal stanicaID As String, _
                                  ByVal brDok As String, ByVal brojZbirne As String, _
                                  ByVal kolI As Double, ByVal cenaI As Double, _
                                  ByVal kolAmb As Double, _
                                  ByVal kolII As Double, ByVal cenaII As Double, _
                                  ByVal kolAmbII As Double) As String
    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", datum
    h.Add "KooperantID", TEST_KOOP_ID
    h.Add "StanicaID", stanicaID
    h.Add "KulturaID", TEST_KULTURA_ID
    h.Add "VrstaVoca", TEST_VRSTA
    h.Add "SortaVoca", TEST_SORTA
    h.Add "TipAmbalaze", TEST_TIP_AMB
    h.Add "BrojDokumenta", brDok
    h.Add "ParcelaID", GetTestParcelaID()

    Dim stavke As Collection
    Set stavke = New Collection
    If kolI > 0 Then stavke.Add OtkStavka(KLASA_I, kolI, cenaI, kolAmb, 0#)
    If kolII > 0 Then stavke.Add OtkStavka(KLASA_II, kolII, cenaII, kolAmbII, 0#)

    Dim greska As String
    NoviOtkupFixture = CreateOtkup_TX(h, stavke, greska)
    If Len(NoviOtkupFixture) = 0 Then
        Err.Raise vbObjectError + 9400, "NoviOtkupFixture", _
                  "CreateOtkup_TX nije vratio ID: " & greska
    End If

    If Len(Trim$(brojZbirne)) > 0 Then
        RequireUpdateCell TBL_OTKUP, _
                          FindRows(TBL_OTKUP, COL_OTK_ID, NoviOtkupFixture)(1), _
                          COL_OTK_BROJ_ZBIRNE, brojZbirne, "NoviOtkupFixture"
    End If
End Function

' Parametri kakve ekran salje OtkupUpisi -- isti kljucevi kao NoviOtkupUnos.
Private Function OtkEkranParam(ByVal brDok As String) As Object
    Dim p As Object
    Set p = CreateObject("Scripting.Dictionary")
    p.CompareMode = vbTextCompare
    p("datum") = NextTestDate()
    p("stanicaID") = TEST_ST_ID
    p("kooperantID") = TEST_KOOP_ID
    p("vrsta") = TEST_VRSTA
    p("sorta") = TEST_SORTA
    p("tipAmb") = TEST_TIP_AMB
    p("vozacID") = TEST_VOZ_ID
    p("brDok") = brDok
    p("brojZbirne") = ""
    p("parcelaID") = ""
    p("primalac") = ""
    p("kolicinaI") = 0#
    p("cenaI") = 0#
    p("kolAmb") = 0&
    p("kolAmbIzdata") = 0&
    p("dveKlase") = False
    p("kolicinaII") = 0#
    p("cenaII") = 0#
    p("kolAmbII") = 0&
    p("novac") = 0#
    p("brutoKgI") = 0#
    p("brutoKgII") = 0#
    Set OtkEkranParam = p
End Function

' --- tblAmbalaza, po dokumentu -----------------------------------------------
Private Function AmbRedovi(ByVal dokID As String, ByVal dokTip As String) As Collection
    Dim c As Collection
    Set c = New Collection
    Set AmbRedovi = c

    Dim d As Variant
    d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Function

    Dim cDok As Long, cTip As Long
    cDok = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, "AmbRedovi")
    cTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, "AmbRedovi")

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(i, cDok), "")), dokID, vbTextCompare) = 0 Then
            If StrComp(Trim$(nz(d(i, cTip), "")), dokTip, vbTextCompare) = 0 Then
                c.Add i
            End If
        End If
    Next i
End Function

Private Function AmbBrojRedova(ByVal dokID As String, ByVal dokTip As String) As Long
    AmbBrojRedova = AmbRedovi(dokID, dokTip).count
End Function

Private Function AmbPolje(ByVal dokID As String, ByVal dokTip As String, _
                          ByVal smer As String, ByVal columnName As String) As String
    Dim d As Variant
    d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Function

    Dim cSmer As Long, cTraz As Long
    cSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, "AmbPolje")
    cTraz = RequireColumnIndex(TBL_AMBALAZA, columnName, "AmbPolje")

    Dim redovi As Collection
    Set redovi = AmbRedovi(dokID, dokTip)

    Dim k As Long, i As Long
    For k = 1 To redovi.count
        i = CLng(redovi(k))
        If StrComp(Trim$(nz(d(i, cSmer), "")), smer, vbTextCompare) = 0 Then
            AmbPolje = Trim$(CStr(nz(d(i, cTraz), "")))
            Exit Function
        End If
    Next k
End Function

Private Function AmbKolicina(ByVal dokID As String, ByVal dokTip As String, _
                             ByVal smer As String) As Double
    Dim t As String
    t = AmbPolje(dokID, dokTip, smer, COL_AMB_KOLICINA)
    If IsNumeric(t) Then AmbKolicina = CDbl(t)
End Function

Private Function AmbEntitet(ByVal dokID As String, ByVal dokTip As String, _
                            ByVal smer As String) As String
    AmbEntitet = AmbPolje(dokID, dokTip, smer, COL_AMB_ENTITET)
End Function

Private Function AmbVozac(ByVal dokID As String, ByVal dokTip As String, _
                          ByVal smer As String) As String
    AmbVozac = AmbPolje(dokID, dokTip, smer, COL_AMB_VOZAC)
End Function

' --- OTP pomocne -------------------------------------------------------------
Private Function OtpHeader(ByVal brojOtp As String) As Object
    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", NextTestDate()
    h.Add "StanicaID", TEST_ST_ID
    h.Add "VozacID", TEST_VOZ_ID
    h.Add "KulturaID", TEST_KULTURA_ID
    h.Add "TipAmbalaze", TEST_TIP_AMB
    h.Add "BrojOtpremnice", brojOtp
    Set OtpHeader = h
End Function

' Zaglavlje sa ZADATIM danom i stanicom -- za testove broja, gde je
' (stanica, dan) upravo ono sto se meri.
Private Function OtpBrojHeader(ByVal brojOtp As String, ByVal dan As Date, _
                               ByVal stanicaID As String) As Object
    Dim h As Object
    Set h = OtpHeader(brojOtp)
    h("Datum") = dan
    h("StanicaID") = stanicaID
    Set OtpBrojHeader = h
End Function

' Ocekivanje: sta je operater prijavio da otpremnica nosi.
Private Function OtpOcek(ByVal kolI As Double, ByVal ambI As Double, _
                         ByVal kolII As Double, ByVal ambII As Double) As Collection
    Dim c As Collection
    Set c = New Collection
    If kolI > 0 Then c.Add OtpOcekStavka(KLASA_I, kolI, ambI)
    If kolII > 0 Then c.Add OtpOcekStavka(KLASA_II, kolII, ambII)
    Set OtpOcek = c
End Function

Private Function OtpOcekStavka(ByVal klasa As String, ByVal kol As Double, _
                               ByVal amb As Double) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", klasa
    s.Add "Kolicina", kol
    s.Add "KolAmbalaze", amb
    Set OtpOcekStavka = s
End Function

' Ocekivana stavka SA predlogom cene (S3a). Predlog je izricito ne-finansijsko
' polje: prefiluje formu otkupa, ne knjizi se i ne sabira.
Private Function OtpOcekStavkaSaCenom(ByVal klasa As String, ByVal kol As Double, _
                                      ByVal amb As Double, _
                                      ByVal predlogCena As Double) As Object
    Dim s As Object
    Set s = OtpOcekStavka(klasa, kol, amb)
    s.Add "PredlogCena", predlogCena
    Set OtpOcekStavkaSaCenom = s
End Function

' --- knjizenje ambalaze po dokumentu -----------------------------------------
' Redovi tblAmbalaza vezani za dati dokument. Trazi se po DokumentID-u, ne po
' broju: broj otpremnice je jedinstven tek po (otkupno mesto, dan).
Private Function AmbRedoviZaDokument(ByVal dokID As String) As Collection
    Dim res As Collection
    Set res = New Collection
    Set AmbRedoviZaDokument = res

    Dim d As Variant
    d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Function

    Dim cDok As Long
    cDok = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, "AmbRedoviZaDokument")

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cDok))), Trim$(dokID), vbTextCompare) = 0 Then
            res.Add i
        End If
    Next i
End Function

Private Function AmbRedovaZaDokument(ByVal dokID As String) As Long
    AmbRedovaZaDokument = AmbRedoviZaDokument(dokID).count
End Function

Private Function AmbKolicinaZaDokument(ByVal dokID As String) As Double
    Dim redovi As Collection
    Set redovi = AmbRedoviZaDokument(dokID)
    If redovi.count = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Function

    Dim cKol As Long
    cKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, "AmbKolicinaZaDokument")

    Dim i As Long
    For i = 1 To redovi.count
        AmbKolicinaZaDokument = AmbKolicinaZaDokument + _
                                CDbl(nz(d(CLng(redovi(i)), cKol), 0))
    Next i
End Function

' Vrednost polja PRVOG reda ambalaze tog dokumenta. Tvrdnje koje ga koriste prvo
' proveravaju da red postoji tacno jedan.
Private Function AmbPoljeZaDokument(ByVal dokID As String, _
                                    ByVal columnName As String) As String
    Dim redovi As Collection
    Set redovi = AmbRedoviZaDokument(dokID)
    If redovi.count = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Function

    AmbPoljeZaDokument = Trim$(NzToText(d(CLng(redovi(1)), _
                         RequireColumnIndex(TBL_AMBALAZA, columnName, "AmbPoljeZaDokument"))))
End Function

' Otkup po NOVOM modelu -- izvor kakav kanonska otpremnica prima.
Private Function OtpNoviOtkup(ByVal scenario As String, ByVal kolI As Double, _
                              ByVal kolII As Double) As String
    OtpNoviOtkup = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-" & scenario), _
                                  OtkStavke(kolI, 50#, 20, kolII, 40#, 30))
End Function

Private Function OtpNoviOtkupSaBrutom(ByVal scenario As String, ByVal kol As Double, _
                                      ByVal bruto As Double) As String
    Dim c As Collection
    Set c = New Collection
    c.Add OtkStavka(KLASA_I, kol, 50#, 20#, bruto)
    OtpNoviOtkupSaBrutom = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-" & scenario), c)
End Function

Private Function OtpNoviOtkupDrugogTipa(ByVal scenario As String) As String
    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-" & scenario)
    h("TipAmbalaze") = TEST_TIP_AMB_B
    OtpNoviOtkupDrugogTipa = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0))
End Function

' Drugi tip ambalaze na headeru otkupa, ali NULA gajbi na stavkama: tip postoji
' zbog izdate ambalaze (KolAmbIzdata), a u otpremnicu ne ide nijedna gajba.
Private Function OtpNoviOtkupDrugogTipaBezGajbi(ByVal scenario As String) As String
    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-" & scenario)
    h("TipAmbalaze") = TEST_TIP_AMB_B
    h.Add "KolAmbIzdata", 10#

    Dim c As Collection
    Set c = New Collection
    c.Add OtkStavka(KLASA_I, 400#, 50#, 0#, 0#)
    OtpNoviOtkupDrugogTipaBezGajbi = CreateOtkup_TX(h, c)
End Function

' Sirov upis clanstva, mimo writera -- samo da se napravi korupcija koju
' strikt loader mora da vidi. Vraca indeks reda, da se moze skloniti.
Private Function OtpUpisiSirovoClanstvo(ByVal otpID As String, _
                                        ByVal otkupID As String) As Long
    Dim rowData As Variant
    rowData = BlankRow(TBL_OTPREMNICA_IZVORI)

    SetRequiredField rowData, TBL_OTPREMNICA_IZVORI, COL_OPI_ID, _
                     "OPI-SAB-" & Format$(Timer * 1000, "0")
    SetRequiredField rowData, TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, otpID
    SetRequiredField rowData, TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, otkupID

    OtpUpisiSirovoClanstvo = AppendRow(TBL_OTPREMNICA_IZVORI, rowData)
    If OtpUpisiSirovoClanstvo <= 0 Then
        Err.Raise vbObjectError + 9300, "OtpUpisiSirovoClanstvo", "AppendRow nije uspeo."
    End If
End Function

' Greska koju read-model digne, kao tekst -- "" znaci da je prosao.
Private Function OtpProgressGreska(ByVal otpID As String) As String
    Dim p As Object
    On Error Resume Next
    Err.Clear
    Set p = GetOtpremnicaProgress(otpID)
    OtpProgressGreska = Err.description
    If Err.Number = 0 Then OtpProgressGreska = ""
    Err.Clear
    On Error GoTo 0
End Function

Private Function OtpNoviOtkupNaStanici(ByVal scenario As String, _
                                       ByVal stanicaID As String) As String
    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-" & scenario)
    h("StanicaID") = stanicaID
    OtpNoviOtkupNaStanici = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0))
End Function

Private Function OtpNoviOtkupDrugeKulture(ByVal scenario As String) As String
    Dim h As Object
    Set h = OtkHeader(TEST_PREFIX & "-OTK-" & scenario)
    h("KulturaID") = TEST_KUL_BEZ_SORTE_ID
    h("VrstaVoca") = TEST_VRSTA_BEZ_SORTE
    h("SortaVoca") = ""
    OtpNoviOtkupDrugeKulture = CreateOtkup_TX(h, OtkStavke(400#, 50#, 20, 0#, 0#, 0))
End Function

Private Function OtpPolje(ByVal otpID As String, ByVal columnName As String) As String
    OtpPolje = Trim$(CStr(nz(GetValueByKey(TBL_OTPREMNICA, COL_OTP_ID, otpID, _
                                           columnName), "")))
End Function

Private Function OtpBrojStavki(ByVal otpID As String) As Long
    OtpBrojStavki = FindRows(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, otpID).count
End Function

Private Function OtpBrojClanova(ByVal otpID As String) As Long
    OtpBrojClanova = FindRows(TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, otpID).count
End Function

' Zaglavlje otkupa BEZ STAVKI -- synthetic anomaly / fault injection.
'
' Jedini pisac (CreateOtkup_TX) takav red ne pravi; kapije koje ga odbijaju se
' mere tako sto se kanonskom otkupu obrisu stavke. Pozivalac drzi transakciju
' sa snapshot-om TBL_OTKUP i TBL_OTKUP_STAVKE i vraca je rollback-om.
Private Function OtkupBezStavkiFixture(ByVal brDok As String) As String
    Const SRC As String = "OtkupBezStavkiFixture"

    Dim razlog As String
    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(brDok), OtkStavke(400#, 50#, 20, 0#, 0#, 0), razlog)
    If Len(otkID) = 0 Then
        Err.Raise vbObjectError + 9401, SRC, "CreateOtkup_TX nije vratio ID: " & razlog
    End If

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkID)
    Dim k As Long
    If Not rows Is Nothing Then
        For k = rows.count To 1 Step -1
            RequireDeleteRow TBL_OTKUP_STAVKE, CLng(rows(k)), SRC
        Next k
    End If

    OtkupBezStavkiFixture = otkID
End Function

Private Function OtkBrojStavkiZaOtkup(ByVal otkupID As String) As Long
    OtkBrojStavkiZaOtkup = FindRows(TBL_OTKUP_STAVKE, COL_OKS_OTKUP_ID, otkupID).count
End Function

Private Function OtpClanoviTest(ByVal otpID As String) As Collection
    Dim c As Collection
    Set c = New Collection
    Set OtpClanoviTest = c

    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA_IZVORI)
    If Not IsArray(d) Then Exit Function

    Dim cOtp As Long, cOtk As Long
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, "OtpClanoviTest")
    cOtk = RequireColumnIndex(TBL_OTPREMNICA_IZVORI, COL_OPI_OTKUP_ID, "OtpClanoviTest")

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(i, cOtp), "")), otpID, vbTextCompare) = 0 Then
            c.Add Trim$(CStr(nz(d(i, cOtk), "")))
        End If
    Next i
End Function

Private Function OtpStavkaPolje(ByVal otpID As String, ByVal klasa As String, _
                                ByVal columnName As String) As String
    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA_STAVKE)
    If Not IsArray(d) Then Exit Function

    Dim cOtp As Long, cKlasa As Long, cTraz As Long
    cOtp = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, "OtpStavkaPolje")
    cKlasa = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, COL_OPS_KLASA, "OtpStavkaPolje")
    cTraz = RequireColumnIndex(TBL_OTPREMNICA_STAVKE, columnName, "OtpStavkaPolje")

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(i, cOtp), "")), otpID, vbTextCompare) = 0 Then
            If StrComp(Trim$(nz(d(i, cKlasa), "")), klasa, vbTextCompare) = 0 Then
                OtpStavkaPolje = Trim$(CStr(nz(d(i, cTraz), "")))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function OtpStavkaBrojP(ByVal otpID As String, ByVal klasa As String, _
                                ByVal columnName As String) As Double
    Dim t As String
    t = OtpStavkaPolje(otpID, klasa, columnName)
    If IsNumeric(t) Then OtpStavkaBrojP = CDbl(t)
End Function

' --- PR3 pomocne -------------------------------------------------------------

' Header BEZ vrste/sorte/tipa ambalaze -- oni se izvode iz izvornih otpremnica.
Private Function Pr3Header(ByVal brojZbirne As String) As Object
    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    h.Add "Datum", NextTestDate()
    h.Add "VozacID", TEST_VOZ_ID
    h.Add "BrojZbirne", brojZbirne
    h.Add "KupacID", TEST_KUP_ID
    h.Add "Hladnjaca", ""
    h.Add "Pogon", ""
    Set Pr3Header = h
End Function

' Otpremnica BEZ broja zbirne -- slobodna da je novi writer preuzme.
'
' Pravi je PRODUKCIONI pisac otpremnice (CreateOtpremnicaDraft_TX), sa jednom
' ocekivanom stavkom. Do S3b-1 je ovde stajao stari pisac (SaveOtpremnica_TX),
' koji je pravio zaglavlje bez stavki -- dokument koji novi citaoci odbijaju.
' Vrsta dolazi iz KULTURE (tako je i u F2): druga vrsta = druga kultura.
Private Function Pr3Otpremnica(ByVal broj As String, ByVal klasa As String, _
                               ByVal kol As Double, ByVal amb As Long, _
                               Optional ByVal kulturaID As String = "") As String
    Pr3Otpremnica = Pr3OtpremnicaVozac(broj, klasa, kol, amb, TEST_VOZ_ID, kulturaID)
End Function

Private Function Pr3Izvor(ByVal a As String, ByVal b As String) As Collection
    Dim c As Collection
    Set c = New Collection
    If Len(a) > 0 Then c.Add a
    If Len(b) > 0 Then c.Add b
    Set Pr3Izvor = c
End Function

Private Function Pr3Ocekivano(ByVal klasa As String, ByVal kol As Double, _
                              ByVal amb As Double) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s.Add "Klasa", klasa
    s.Add "Kolicina", kol
    s.Add "KolAmbalaze", amb
    Set Pr3Ocekivano = s
End Function

' Gajbe otpremnice su na STAVCI (S3b-1) -- zaglavlje ih vise ne nosi, pa bi upis
' u zaglavlje bio kvar koji niko ne cita. Otpremnica iz Pr3Otpremnica ima tacno
' jednu stavku.
Private Sub Pr3PostaviAmbalazu(ByVal otpID As String, ByVal amb As Double)
    Dim redovi As Collection
    Set redovi = FindRows(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, otpID)
    If redovi Is Nothing Then Exit Sub
    If redovi.count <> 1 Then Exit Sub
    RequireUpdateCell TBL_OTPREMNICA_STAVKE, CLng(redovi(1)), COL_OPS_KOL_AMB, amb, _
                      "Pr3PostaviAmbalazu"
End Sub

' Na kojoj je AKTIVNOJ zbirnoj otpremnica -- kroz PRODUKCIONI citac.
'
' Ranije je citalo kolonu Otpremnica.ZbirnaID. Kolone vise nema: clanstvo je
' jedini zapis, a odgovor se racuna. Test namerno ide kroz javni API, da citac
' koji zamenjuje kolonu ima pokrice.
Private Function Pr3OtpZbirnaID(ByVal otpID As String) As String
    Pr3OtpZbirnaID = modDokumenta.AktivnaZbirnaZaOtpremnicu(otpID)
End Function

' Prefiks + tacno 32 hex znaka. NE proverava "nije broj": 32 hex znaka smeju
' slucajno biti sve cifre, pa bi takva tvrdnja bila flaky bez razloga.
Private Function Pr3JeOpaqueID(ByVal id As String, ByVal prefiks As String) As Boolean
    Dim i As Long

    If StrComp(Left$(id, Len(prefiks)), prefiks, vbTextCompare) <> 0 Then Exit Function
    If Len(id) - Len(prefiks) <> 32 Then Exit Function

    For i = Len(prefiks) + 1 To Len(id)
        If InStr(1, "0123456789ABCDEF", UCase$(Mid$(id, i, 1))) = 0 Then Exit Function
    Next i

    Pr3JeOpaqueID = True
End Function

Private Function Pr3BrojRedova(ByVal tblName As String) As Long
    Dim d As Variant
    d = GetTableData(tblName)
    If Not IsArray(d) Then Exit Function
    Pr3BrojRedova = UBound(d, 1)
End Function

Private Function Pr3BrojStavki(ByVal zbirnaID As String) As Long
    Dim redovi As Collection
    Set redovi = FindRows(TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, zbirnaID)
    If redovi Is Nothing Then Exit Function
    Pr3BrojStavki = redovi.count
End Function

Private Function Pr3StavkaPolje(ByVal zbirnaID As String, ByVal klasa As String, _
                                ByVal columnName As String) As String
    Dim d As Variant
    d = GetTableData(TBL_ZBIRNA_STAVKE)
    If Not IsArray(d) Then Exit Function

    Dim cZbr As Long, cKlasa As Long, cTraz As Long
    cZbr = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, "Pr3StavkaPolje")
    cKlasa = RequireColumnIndex(TBL_ZBIRNA_STAVKE, COL_ZBS_KLASA, "Pr3StavkaPolje")
    cTraz = RequireColumnIndex(TBL_ZBIRNA_STAVKE, columnName, "Pr3StavkaPolje")

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(i, cZbr), "")), zbirnaID, vbTextCompare) = 0 Then
            If StrComp(Trim$(nz(d(i, cKlasa), "")), klasa, vbTextCompare) = 0 Then
                Pr3StavkaPolje = Trim$(CStr(nz(d(i, cTraz), "")))
                Exit Function
            End If
        End If
    Next i
End Function

Private Function Pr3StavkaBroj(ByVal zbirnaID As String, ByVal klasa As String, _
                               ByVal columnName As String) As Double
    Dim t As String
    t = Pr3StavkaPolje(zbirnaID, klasa, columnName)
    If IsNumeric(t) Then Pr3StavkaBroj = CDbl(t)
End Function

Private Function Pr3BrojIzvora(ByVal zbirnaID As String) As Long
    Dim redovi As Collection
    Set redovi = FindRows(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, zbirnaID)
    If redovi Is Nothing Then Exit Function
    Pr3BrojIzvora = redovi.count
End Function

Private Function Pr3JeIzvor(ByVal zbirnaID As String, ByVal otpID As String) As Boolean
    Dim d As Variant
    d = GetTableData(TBL_ZBIRNA_IZVORI)
    If Not IsArray(d) Then Exit Function

    Dim cZbr As Long, cOtp As Long, i As Long
    cZbr = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, "Pr3JeIzvor")
    cOtp = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, "Pr3JeIzvor")

    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(i, cZbr), "")), zbirnaID, vbTextCompare) = 0 Then
            If StrComp(Trim$(nz(d(i, cOtp), "")), otpID, vbTextCompare) = 0 Then
                Pr3JeIzvor = True
                Exit Function
            End If
        End If
    Next i
End Function

' Kojoj zbirnoj otpremnica pripada PO ZAPISU CLANSTVA (ne po kesu).
Private Function Pr3ZbirnaIzClanstva(ByVal otpID As String) As String
    Dim d As Variant
    d = GetTableData(TBL_ZBIRNA_IZVORI)
    If Not IsArray(d) Then Exit Function

    Dim cZbr As Long, cOtp As Long, i As Long
    cZbr = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, "Pr3ZbirnaIzClanstva")
    cOtp = RequireColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, "Pr3ZbirnaIzClanstva")

    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(nz(d(i, cOtp), "")), otpID, vbTextCompare) = 0 Then
            Pr3ZbirnaIzClanstva = Trim$(nz(d(i, cZbr), ""))
            Exit Function
        End If
    Next i
End Function

Private Function Pr3OtpremnicaVozac(ByVal broj As String, ByVal klasa As String, _
                                    ByVal kol As Double, ByVal amb As Long, _
                                    ByVal vozac As String, _
                                    Optional ByVal kulturaID As String = "") As String
    Dim h As Object
    Set h = OtpHeader(broj)
    h("VozacID") = vozac
    If Len(kulturaID) > 0 Then h("KulturaID") = kulturaID

    Dim c As Collection
    Set c = New Collection
    c.Add OtpOcekStavka(klasa, kol, CDbl(amb))

    Dim razlog As String
    Pr3OtpremnicaVozac = CreateOtpremnicaDraft_TX(h, c, razlog)
End Function


' Rucno ubaci zapis clanstva -- SAMO za test korupcije kanona. Produkcioni put
' je iskljucivo CreateZbirna_TX.
Private Sub Pr3DodajClanstvo(ByVal zbirnaID As String, ByVal otpremnicaID As String)
    Dim lo As ListObject
    Set lo = GetTable(TBL_ZBIRNA_IZVORI)
    If lo Is Nothing Then Exit Sub

    Dim rowData() As Variant
    ReDim rowData(0 To lo.ListColumns.count - 1)

    rowData(GetColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ID) - 1) = _
        modDataAccess.NewEntityID("ZBI-")
    rowData(GetColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID) - 1) = zbirnaID
    rowData(GetColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID) - 1) = otpremnicaID

    AppendRow TBL_ZBIRNA_IZVORI, rowData
End Sub

' Ukloni tacno jedan zapis clanstva. SAMO za ciscenje posle testa korupcije;
' produkcija zapise clanstva ne brise (A15 -- istorija sastava).
Private Sub Pr3UkloniClanstvo(ByVal zbirnaID As String, ByVal otpremnicaID As String)
    Dim lo As ListObject
    Set lo = GetTable(TBL_ZBIRNA_IZVORI)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim d As Variant
    d = lo.DataBodyRange.Value2
    If IsEmpty(d) Then Exit Sub

    Dim cZbr As Long, cOtp As Long, i As Long
    cZbr = GetColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID)
    cOtp = GetColumnIndex(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID)
    If cZbr = 0 Or cOtp = 0 Then Exit Sub

    For i = UBound(d, 1) To 1 Step -1
        If StrComp(Trim$(nz(d(i, cZbr), "")), zbirnaID, vbTextCompare) = 0 Then
            If StrComp(Trim$(nz(d(i, cOtp), "")), otpremnicaID, vbTextCompare) = 0 Then
                lo.ListRows(i).Delete
                Exit Sub
            End If
        End If
    Next i
End Sub

' Pozicija kolone u KANONU (modSchema), ne u zatecenoj svesci. Nula = nema je.
Private Function Pr3KanonskaPozicija(ByVal tblName As String, _
                                     ByVal colName As String) As Long
    Dim kolone As Collection
    Set kolone = modSchema.SchemaTableColumns(tblName)
    If kolone Is Nothing Then Exit Function

    Dim i As Long
    For i = 1 To kolone.count
        If StrComp(CStr(kolone(i)), colName, vbTextCompare) = 0 Then
            Pr3KanonskaPozicija = i
            Exit Function
        End If
    Next i
End Function

Private Function Pr3BrojClanstavaZa(ByVal otpID As String) As Long
    Dim redovi As Collection
    Set redovi = FindRows(TBL_ZBIRNA_IZVORI, COL_ZBI_OTPREMNICA_ID, otpID)
    If redovi Is Nothing Then Exit Function
    Pr3BrojClanstavaZa = redovi.count
End Function

Private Function NewScenarioCode(ByVal scenarioName As String) As String
    NewScenarioCode = scenarioName & "-" & m_RunID & "-" & CStr(m_Total + 1)
End Function

Private Function NextTestDate() As Date
    m_DateSeq = m_DateSeq + 1
    NextTestDate = DateSerial(2090, 1, 1) + m_DateSeq
End Function

' ============================================================
' ASSERTIONS
' ============================================================

Private Sub AssertTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        LogPass testName
    Else
        LogFail testName, "Assertion failed."
    End If
End Sub

Private Sub AssertEquals(ByVal expected As String, ByVal actual As String, ByVal testName As String)
    If CStr(expected) = CStr(actual) Then
        LogPass testName
    Else
        LogFail testName, "Expected [" & CStr(expected) & "], got [" & CStr(actual) & "]."
    End If
End Sub

Private Sub AssertDoubleNear(ByVal expected As Double, ByVal actual As Double, _
                             ByVal TOLERANCE As Double, ByVal testName As String)
    If Abs(expected - actual) <= TOLERANCE Then
        LogPass testName
    Else
        LogFail testName, "Expected [" & CStr(expected) & "], got [" & CStr(actual) & "]."
    End If
End Sub

' ============================================================
' LOGGING
' ============================================================

Private Sub LogPass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1

    Debug.Print "[PASS] " & testName
    AppendTestLog "TEST", testName, "PASS", ""
    m_Report = m_Report & "PASS " & testName & vbLf
End Sub

Private Sub LogFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendTestLog "TEST", testName, "FAIL", details
    ' Separator je " :: ", ne " -- ": tekst tvrdnje SME da sadrzi " -- ", pa bi
    ' se ime na njemu odseklo pri citanju (dokaz.py bi javio NE OBARA SVOJ TEST
    ' nad sabotazom koja radi savrseno). Isti separator modul vec koristi u
    ' Debug.Print, pa je izlaz i konzistentan.
    m_Report = m_Report & "FAIL " & testName & " :: " & details & vbLf
End Sub

Private Sub WriteResultFileBFP()
    Dim path As String
    Dim fnum As Integer
    path = ThisWorkbook.path & Application.PathSeparator & "last_run_bfp.txt"
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, "TESTS=" & m_Total & " FAIL=" & m_Failed & vbLf & m_Report;
    Close #fnum
End Sub

Private Sub LogSkip(ByVal testName As String, ByVal reason As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1

    Debug.Print "[SKIP] " & testName & " :: " & reason
    AppendTestLog "TEST", testName, "SKIP", reason
    m_Report = m_Report & "SKIP " & testName & " :: " & reason & vbLf
End Sub

Private Sub LogFatal(ByVal sourceName As String, ByVal errNum As Long, ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
    AppendTestLog "FATAL", sourceName, "FAIL", CStr(errNum) & " - " & errDesc
    m_Report = m_Report & "FAIL " & sourceName & " :: FATAL " & _
               CStr(errNum) & " " & errDesc & vbLf
End Sub

Private Sub InitTestLog()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TEST_LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = TEST_LOG_SHEET
        ws.Range("A1:G1").value = Array("Timestamp", "RunID", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendTestLog(ByVal kindText As String, ByVal nameText As String, _
                          ByVal statusText As String, ByVal detailsText As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TEST_LOG_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = ws.cells(ws.rows.count, 1).End(xlUp).row + 1

    ws.cells(r, 1).value = Now
    ws.cells(r, 2).value = m_RunID
    ws.cells(r, 3).value = kindText
    ws.cells(r, 4).value = nameText
    ws.cells(r, 5).value = statusText
    ws.cells(r, 6).value = Left$(detailsText, 2000)
    ws.cells(r, 7).value = Environ$("Username")
End Sub







Public Sub HardDeleteBusinessFlowTestRows()
    On Error GoTo EH

    Dim answer As String
    answer = InputBox( _
        "Ovo CE FIZICKI OBRISATI sve TST-PRO-* redove iz svih tabela." & vbCrLf & _
        "Ova operacija je NEPOVRATNA." & vbCrLf & vbCrLf & _
        "Ukucaj BRISI da nastavis:", _
        "Potvrda brisanja test podataka")

    If answer <> "BRISI" Then
        MsgBox "Brisanje otkazano.", vbInformation
        Exit Sub
    End If

    Dim deleted As Long
    Dim total As Long

    deleted = DeleteTestRowsFromTable(TBL_FAKTURA_STAVKE, Array("FakturaID", "BrojPrijemnice"))
    total = total + deleted
    Debug.Print "tblFakturaStavke: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_FAKTURE, Array("BrojFakture"))
    total = total + deleted
    Debug.Print "tblFakture: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_PRIJEMNICA, Array("BrojPrijemnice", "BrojZbirne"))
    total = total + deleted
    Debug.Print "tblPrijemnica: " & deleted & " obrisano"

    ' Deca zbirne PRE roditelja, i po FK -- ne po markeru: tblZbirnaStavke i
    ' tblZbirnaIzvori nemaju nijednu kolonu sa TST-PRO tekstom, pa bi ih
    ' DeleteTestRowsFromTable preskocio i ostavio orphan redove.
    deleted = DeleteChildRowsByParent(TBL_ZBIRNA_STAVKE, COL_ZBS_ZBIRNA_ID, _
                                      TBL_ZBIRNA, COL_ZBR_ID, "BrojZbirne")
    total = total + deleted
    Debug.Print "tblZbirnaStavke: " & deleted & " obrisano"

    deleted = DeleteChildRowsByParent(TBL_ZBIRNA_IZVORI, COL_ZBI_ZBIRNA_ID, _
                                      TBL_ZBIRNA, COL_ZBR_ID, "BrojZbirne")
    total = total + deleted
    Debug.Print "tblZbirnaIzvori: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_ZBIRNA, Array("BrojZbirne"))
    total = total + deleted
    Debug.Print "tblZbirna: " & deleted & " obrisano"

    ' Deca otpremnice PRE roditelja i po FK -- isti razlog kao kod zbirne.
    deleted = DeleteChildRowsByParent(TBL_OTPREMNICA_STAVKE, COL_OPS_OTPREMNICA_ID, _
                                      TBL_OTPREMNICA, COL_OTP_ID, "BrojOtpremnice")
    total = total + deleted
    Debug.Print "tblOtpremnicaStavke: " & deleted & " obrisano"

    deleted = DeleteChildRowsByParent(TBL_OTPREMNICA_IZVORI, COL_OPI_OTPREMNICA_ID, _
                                      TBL_OTPREMNICA, COL_OTP_ID, "BrojOtpremnice")
    total = total + deleted
    Debug.Print "tblOtpremnicaIzvori: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_OTPREMNICA, Array("BrojOtpremnice", "BrojZbirne"))
    total = total + deleted
    Debug.Print "tblOtpremnica: " & deleted & " obrisano"

    ' ClientRecordID je u spisku zbog PWA uvoza: njihov BrojDokumenta je
    ' GENERISAN (kanonski format ne trpi TST-PRO), pa marker nosi samo CRID.
    deleted = DeleteTestRowsFromTable(TBL_OTKUP, _
                  Array("BrojDokumenta", "BrojZbirne", "ClientRecordID"))
    total = total + deleted
    Debug.Print "tblOtkup: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_AMBALAZA, Array("DokumentID"))
    total = total + deleted
    Debug.Print "tblAmbalaza: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_NOVAC, Array("BrojDokumenta"))
    total = total + deleted
    Debug.Print "tblNovac: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable("tblSEFSubmission", Array("FakturaID"))
    total = total + deleted
    Debug.Print "tblSEFSubmission: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable("tblSEFEventLog", Array("FakturaID"))
    total = total + deleted
    Debug.Print "tblSEFEventLog: " & deleted & " obrisano"

    MsgBox "Obrisano ukupno " & total & " test redova.", vbInformation
    Exit Sub

EH:
    MsgBox "Greska pri brisanju: " & Err.description, vbCritical
End Sub

' Obrisi decu cija je RODITELJSKA vrednost test red.
'
' DeleteTestRowsFromTable trazi TST-PRO u koloni same tabele; tabele stavki i
' clanstva nose samo ID-eve, pa im marker mora doci od roditelja.
Private Function DeleteChildRowsByParent(ByVal childTable As String, _
                                         ByVal childFkCol As String, _
                                         ByVal parentTable As String, _
                                         ByVal parentIdCol As String, _
                                         ByVal parentMarkerCol As String) As Long
    On Error GoTo EH

    Dim pd As Variant
    pd = GetTableData(parentTable)
    If Not IsArray(pd) Then Exit Function

    Dim cPid As Long, cMark As Long
    cPid = GetColumnIndex(parentTable, parentIdCol)
    cMark = GetColumnIndex(parentTable, parentMarkerCol)
    If cPid = 0 Or cMark = 0 Then Exit Function

    Dim meta As Object
    Set meta = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(pd, 1)
        If InStr(1, CStr(pd(i, cMark)), "TST-PRO", vbTextCompare) > 0 Then
            meta(UCase$(Trim$(nz(pd(i, cPid), "")))) = True
        End If
    Next i
    If meta.count = 0 Then Exit Function

    Dim lo As ListObject
    Set lo = GetTable(childTable)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cFk As Long
    cFk = GetColumnIndex(childTable, childFkCol)
    If cFk = 0 Then Exit Function

    Dim cd As Variant
    cd = lo.DataBodyRange.Value2
    If IsEmpty(cd) Then Exit Function

    Dim obrisano As Long
    For i = UBound(cd, 1) To 1 Step -1
        If meta.Exists(UCase$(Trim$(nz(cd(i, cFk), "")))) Then
            lo.ListRows(i).Delete
            obrisano = obrisano + 1
        End If
    Next i

    DeleteChildRowsByParent = obrisano
    Exit Function

EH:
    Debug.Print "DeleteChildRowsByParent greska (" & childTable & "): " & Err.description
End Function

Private Function DeleteTestRowsFromTable(ByVal tableName As String, _
                                         ByVal markerColumns As Variant) As Long
    On Error GoTo EH

    Dim lo As ListObject
    Set lo = GetTable(tableName)

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim data As Variant
    data = lo.DataBodyRange.Value2

    If IsEmpty(data) Then Exit Function

    ' Sakupi indekse redova koji treba brisati -- od dna ka vrhu
    Dim toDelete() As Long
    Dim deleteCount As Long
    ReDim toDelete(1 To lo.DataBodyRange.rows.count)

    Dim i As Long
    For i = UBound(data, 1) To 1 Step -1
        If RowHasTestPrefix(data, i, tableName, markerColumns) Then
            deleteCount = deleteCount + 1
            toDelete(deleteCount) = i
        End If
    Next i

    If deleteCount = 0 Then Exit Function

    ' Brisi od dna ka vrhu da ne pomeramo indekse
    Dim j As Long
    For j = 1 To deleteCount
        lo.ListRows(toDelete(j)).Delete
    Next j

    DeleteTestRowsFromTable = deleteCount
    Exit Function

EH:
    Debug.Print "DeleteTestRowsFromTable greska (" & tableName & "): " & Err.description
    DeleteTestRowsFromTable = 0
End Function

' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function RowHasTestPrefix(ByRef data As Variant, ByVal rowIndex As Long, _
                                   ByVal tableName As String, _
                                   ByVal markerColumns As Variant) As Boolean
    Const prefix As String = "TST-PRO"

    Dim c As Variant
    For Each c In markerColumns
        Dim colIdx As Long
        colIdx = GetColumnIndex(tableName, CStr(c))

        If colIdx > 0 Then
            If InStr(1, CStr(data(rowIndex, colIdx)), prefix, vbTextCompare) > 0 Then
                RowHasTestPrefix = True
                Exit Function
            End If
        End If
    Next c
End Function


