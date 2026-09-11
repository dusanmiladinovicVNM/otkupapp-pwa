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

    Test_OtkupAtomicMultiClassSave
    Test_OtkupClassIIAmbalaza
    Test_FullDocumentChainHappyPath
    Test_DuplicateFakturaIsBlocked
    Test_InvalidSavesDoNotAppend
    Test_OtkupInputValidationHardening
    Test_OtkupReadHelpersExcludeStornirano
    Test_DokumentaInputValidationHardening
    Test_DokumentaReadHelpersExcludeStornirano
    Test_DualClassDocumentWrappers
    Test_MalinaAutoZbirnaFromOtpremnice
    Test_MalinaVozacMirror

    ' RF-28 (MasterSync integritet -- AUD-041/042/043)
    Test_RF28_AutoOtpremnicaNeMesaArtikle
    Test_RF28_BrojZbirneRupaNeDajeDuplikat
    Test_ZBR_ImportDvaUredjajaNeStapaDokumente
    Test_RF28_LinkKonfliktNePrepisuje
    Test_RF28_MembershipKoristiSvojuZbirnu
    Test_RF28_MembershipDanskiProzor
    Test_RF28_NevalidanDatumJeSyncError
    Test_RF28_VozacIDUpdateIshodi

    ' RF-05 (frmDokumenta unos + storno set)
    Test_ProsekGajbeExcludesStornirano
    Test_ManjakPreviewJeZbirnaMinusPrijem
    Test_OpenFaktureExcludeStornirano
    Test_ZbirnaKlasaIIGuard
    Test_PrefillBiraPoslednjuGeneraciju
    Test_GeneracijaIDNaSavePutanji
    Test_GeneracijaNePrelaziVlasnika
    Test_StornoPoBrojuOdbijaDvaVlasnika
    Test_StornoGuardNaSvimPutanjama
    Test_ZBR_MutacijaPoBrojuStajeNaDvaDokumenta
    Test_ZBR_DeteNosiGeneracijuRoditelja
    Test_ZBR_PaletaNasledjujeGeneracijuPrijemnice
    Test_ZBR_BackfillNeVezeStaroDeteNaNovuGeneraciju
    Test_ZBR_MasterSyncNePrepisujeGeneracijuDeteta
    Test_ZBR_KaskadaNeDiraDecuDrugogDokumenta
    Test_ZBR_RezimJeZaCeluOperacijuNePoTabeli
    Test_ZBR_IspravkaVezeSvojuDecuNeTudju
    Test_ZBR_KapijaPustaKadJeIzborScoped
    Test_ZBR_TudjaGeneracijaNeOtvaraKapiju
    Test_StornoGuardUKaskadi
    Test_StornoKaskadaScopePoLancu
    Test_MalinaAutoZbirnaFailSignal
    Test_ZbirnaRowDataColumnMapped
    Test_ZbirnaEkranNosiOdrediste
    Test_OMUlazSmerObavezan
    Test_PorukeKatalogPokrivaDokumenta

    Test_HladnjacaChainHappyPath
    Test_HladnjacaChainFailFastOtpremnica
    Test_HladnjacaChainFailFastZbirna
    Test_HladnjacaChainPrijemnicaFailNoBroj
    Test_HladnjacaChainLinkFailureIsReported
    Test_BackfillHladnjacaDeliBrojPoZbirnoj
    Test_BackfillHladnjacaIgnorisePrijemniceDrugogKupca

    Test_AutoLinkPositiveUniqueMatch
    Test_AutoLinkMustNotCrossBrojZbirne
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
    Test_OTK_BrojStorniranogSeNePonovoKoristi
    Test_OTK_EkranIPisacImajuIstoPravilo
    Test_OTK_StornoJednimID
    Test_OTK_PanelNapredakJePauziran

    ' PWA ingest -- produkcioni put od Otkup cutover-a. RunMasterSyncSmokeSuite
    ' je zatecena crvena (9/26) i nije u FULL prolazu, pa pokrice mora ovde.
    Test_PWA_IngestPraviHeaderIStavku
    Test_PWA_NerazresivaKulturaObaraUvoz
    Test_PWA_IstiCridIstiSadrzajJeNoOp
    Test_PWA_IstiCridDrugiSadrzajPada
    Test_PWA_RazresivacImenujeRazlog

    ' Otpremnica skela -- header + stavke + clanstvo. Izvori su otkupi po
    ' NOVOM modelu, pa ovi testovi mere i da se dva nova pisca slazu.
    Test_OTP_JedanBrojJedanHeader
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
    Test_AutoLinkPositiveUniqueMatch
    Test_AutoLinkMustNotCrossBrojZbirne
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

    RequireColumnsExist TBL_OTKUP, Array( _
        "OtkupID", "Datum", "KooperantID", "StanicaID", "VrstaVoca", _
        "SortaVoca", "Kolicina", "Cena", "TipAmbalaze", "KolAmbalaze", _
        "VozacID", "BrojDokumenta", "Klasa", "BrojZbirne", "OtpremnicaID")

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

Private Sub Test_OtkupAtomicMultiClassSave()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTK")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojDok As String
    Dim brojZbirne As String

    brojDok = TEST_PREFIX & "-OTK-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-OTK-" & scenario

    Dim beforeOtkup As Long
    Dim beforeAmb As Long
    beforeOtkup = CountRows(TBL_OTKUP)
    beforeAmb = CountRows(TBL_AMBALAZA)

    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        1000#, 120#, TEST_TIP_AMB, 100, TEST_VOZ_ID, brojDok, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        True, 200#, 80#)

    AssertTrue Len(Trim$(result)) > 0, "Otkup multi wrapper returns ID(s)"
    AssertEquals CStr(beforeOtkup + 2), CStr(CountRows(TBL_OTKUP)), "Otkup multi wrapper appends exactly two rows"
    AssertTrue CountRows(TBL_AMBALAZA) >= beforeAmb + 1, "Otkup class I ambalaza movement created"

    Dim otkI As String
    Dim otkII As String
    otkI = FindOtkupIDByBrojAndKlasa(brojDok, "I")
    otkII = FindOtkupIDByBrojAndKlasa(brojDok, "II")

    AssertTrue Len(otkI) > 0, "Otkup class I can be found by document number"
    AssertTrue Len(otkII) > 0, "Otkup class II can be found by document number"

    AssertEquals "100", CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkI, "KolAmbalaze")), _
                 "Otkup class I carries ambalaza"

    AssertEquals "0", CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkII, "KolAmbalaze")), _
                 "Otkup class II carries zero ambalaza"

    AssertEquals brojZbirne, CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkI, "BrojZbirne")), _
                 "Otkup class I has scenario BrojZbirne"

    AssertEquals brojZbirne, CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkII, "BrojZbirne")), _
                 "Otkup class II has scenario BrojZbirne"

    Exit Sub

EH:
    LogFail "Otkup atomic multi-class save", Err.description
End Sub

' #3 Klasa II ima SVOJU kolicinu ambalaze (kolAmbII) -> red Klase II nosi te gajbe
' i kreira sopstvene pokrete u ambalaznom ledgeru (ranije: uvek 0 na Klasi II).
Private Sub Test_OtkupClassIIAmbalaza()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTK2A")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojDok As String
    Dim brojZbirne As String
    brojDok = TEST_PREFIX & "-OTK2A-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-OTK2A-" & scenario

    Dim beforeAmb As Long
    beforeAmb = CountRows(TBL_AMBALAZA)

    ' Dve klase, OBE sa svojim gajbama (Klasa I = 100, Klasa II = 30).
    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        1000#, 120#, TEST_TIP_AMB, 100, TEST_VOZ_ID, brojDok, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        hasKlasaII:=True, kolicinaII:=200#, cenaII:=80#, kolAmbII:=30)

    AssertTrue Len(Trim$(result)) > 0, "Otkup multi (II amb) returns ID(s)"

    Dim otkI As String
    Dim otkII As String
    otkI = FindOtkupIDByBrojAndKlasa(brojDok, "I")
    otkII = FindOtkupIDByBrojAndKlasa(brojDok, "II")

    AssertEquals "100", CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkI, "KolAmbalaze")), _
                 "Otkup class I carries its ambalaza (100)"
    AssertEquals "30", CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkII, "KolAmbalaze")), _
                 "Otkup class II carries its own ambalaza (30)"

    ' Obe klase sa gajbama -> kreirani su ambalazni pokreti (Klasa II vise nije 0).
    AssertTrue CountRows(TBL_AMBALAZA) > beforeAmb, _
               "Two-class otkup with crates creates ambalaza movements"

    Exit Sub

EH:
    LogFail "Otkup class II ambalaza", Err.description
End Sub

Private Sub Test_FullDocumentChainHappyPath()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("FLOW")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojOtk As String
    Dim brojOtp As String
    Dim brojZbirne As String
    Dim brojPrij As String

    brojOtk = TEST_PREFIX & "-OTK-" & scenario
    brojOtp = TEST_PREFIX & "-OTP-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-" & scenario

    Dim beforeOtp As Long
    Dim beforeZbr As Long
    Dim beforePrj As Long
    Dim beforeFak As Long
    Dim beforeStavke As Long

    beforeOtp = CountRows(TBL_OTPREMNICA)
    beforeZbr = CountRows(TBL_ZBIRNA)
    beforePrj = CountRows(TBL_PRIJEMNICA)
    beforeFak = CountRows(TBL_FAKTURE)
    beforeStavke = CountRows(TBL_FAKTURA_STAVKE)

    Dim otkupResult As String
    otkupResult = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        1000#, 120#, TEST_TIP_AMB, 100, TEST_VOZ_ID, brojOtk, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        True, 200#, 80#)

    AssertTrue Len(otkupResult) > 0, "Flow setup creates otkup rows"

    Dim otpI As String
    Dim otpII As String

    otpI = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 1000#, 120#, TEST_TIP_AMB, 100, "I")

    otpII = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                              TEST_VRSTA, TEST_SORTA, 200#, 80#, TEST_TIP_AMB, 0, "II")

    AssertTrue Len(otpI) > 0, "Otpremnica class I created"
    AssertTrue Len(otpII) > 0, "Otpremnica class II created"
    AssertEquals CStr(beforeOtp + 2), CStr(CountRows(TBL_OTPREMNICA)), "Exactly two otpremnica rows appended"

    Dim preVal As Variant
    preVal = ValidateZbirnaPreUnosa(brojZbirne, 1000#, 200#, 100)

    AssertTrue CBool(preVal(3)), "Pre-zbirna class I kg validation green"
    AssertTrue CBool(preVal(7)), "Pre-zbirna class II kg validation green"
    AssertEquals "0", CStr(preVal(10)), "Pre-zbirna ambalaza difference is zero"

    Dim zbrI As String
    Dim zbrII As String

    zbrI = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         1000#, TEST_TIP_AMB, 100, "I")

    zbrII = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          200#, TEST_TIP_AMB, 0, "II")

    AssertTrue Len(zbrI) > 0, "Zbirna class I created"
    AssertTrue Len(zbrII) > 0, "Zbirna class II created"
    AssertEquals CStr(beforeZbr + 2), CStr(CountRows(TBL_ZBIRNA)), "Exactly two zbirna rows appended"

    Dim zVal As Variant
    zVal = ValidateZbirna(brojZbirne)

    AssertTrue CBool(zVal(3)), "Post-zbirna kg validation green"
    AssertDoubleNear 0#, CDbl(zVal(2)), 0.01, "Post-zbirna kg difference zero"

    Dim prjI As String
    Dim prjII As String

    prjI = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 990#, 120#, TEST_TIP_AMB, 100, 95, "I")

    prjII = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                              TEST_VRSTA, TEST_SORTA, 190#, 80#, TEST_TIP_AMB, 0, 0, "II")

    AssertTrue Len(prjI) > 0, "Prijemnica class I created"
    AssertTrue Len(prjII) > 0, "Prijemnica class II created"
    AssertEquals CStr(beforePrj + 2), CStr(CountRows(TBL_PRIJEMNICA)), "Exactly two prijemnica rows appended"
    
    ' Kooperant je trebao dobiti Izlaz na otkupu
    Dim koopAmbSaldo As Variant
    koopAmbSaldo = GetAmbalazeStanje(TEST_KOOP_ID, "Kooperant")
    AssertTrue Not IsEmpty(koopAmbSaldo), "Kooperant has ambalaza movements after otkup"

    ' Vozac je trebao dobiti Izlaz na otpremnici
    Dim vozAmbSaldo As Variant
    vozAmbSaldo = GetVozacAmbSaldo(TEST_VOZ_ID)
    AssertTrue Not IsEmpty(vozAmbSaldo), "Vozac has ambalaza movements after otpremnica"

    Dim manjak As Variant
    manjak = CalculateManjak(brojZbirne)

    AssertDoubleNear 1200#, CDbl(manjak(0)), 0.01, "Manjak zbirna kg"
    AssertDoubleNear 1180#, CDbl(manjak(1)), 0.01, "Manjak prijemnica kg"
    AssertDoubleNear 20#, CDbl(manjak(2)), 0.01, "Manjak kg"

    Dim linked As Long
    linked = AutoLinkOtkupOtpremnica_TX()
    AssertTrue linked >= 2, "Auto-link links the scenario otkup rows"

    Dim otkI As String
    Dim otkII As String
    otkI = FindOtkupIDByBrojAndKlasa(brojOtk, "I")
    otkII = FindOtkupIDByBrojAndKlasa(brojOtk, "II")

    AssertEquals otpI, CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkI, "OtpremnicaID")), _
                 "Otkup class I linked to matching otpremnica"

    AssertEquals otpII, CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkII, "OtpremnicaID")), _
                 "Otkup class II linked to matching otpremnica"

    Dim trace As Variant
    trace = TraceByZbirna(brojZbirne)
    AssertTrue Not IsEmpty(trace), "TraceByZbirna returns rows"

    If Not IsEmpty(trace) Then
        AssertTrue UBound(trace, 1) >= 2, "TraceByZbirna returns at least two rows"
    End If

    Dim stavke As Collection
    Set stavke = New Collection

    stavke.Add Array(prjI, 990#, 120#, "I", brojPrij)
    stavke.Add Array(prjII, 190#, 80#, "II", brojPrij)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(TEST_KUP_ID, stavke)

    AssertTrue Len(fakturaID) > 0, "CreateFaktura_TX returns FakturaID"
    Dim expectedIznos As Double
    expectedIznos = (990# * 120#) + (190# * 80#)   ' 118800 + 15200 = 134000

    Dim actualIznos As Double
    Dim iznosVal As Variant
    iznosVal = GetValueByKey(TBL_FAKTURE, "FakturaID", fakturaID, "Iznos")
    If IsNumeric(iznosVal) Then actualIznos = CDbl(iznosVal)

    AssertDoubleNear expectedIznos, actualIznos, 0.01, _
                 "Faktura iznos matches sum of prijemnica stavke"
    AssertEquals CStr(beforeFak + 1), CStr(CountRows(TBL_FAKTURE)), "Exactly one faktura row appended"
    AssertTrue CountRows(TBL_FAKTURA_STAVKE) >= beforeStavke + 2, "At least two faktura stavke appended"

    AssertEquals "Da", CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjI, "Fakturisano")), _
                 "Prijemnica class I marked Fakturisano"

    AssertEquals "Da", CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjII, "Fakturisano")), _
                 "Prijemnica class II marked Fakturisano"

    AssertEquals fakturaID, CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjI, "FakturaID")), _
                 "Prijemnica class I linked to faktura"

    AssertEquals fakturaID, CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjII, "FakturaID")), _
                 "Prijemnica class II linked to faktura"

    LogInfo "Happy path: OTK=" & otkupResult & _
            " | OTP=" & otpI & "/" & otpII & _
            " | ZBR=" & zbrI & "/" & zbrII & _
            " | PRJ=" & prjI & "/" & prjII & _
            " | FAK=" & fakturaID

    Exit Sub

EH:
    LogFail "Full document chain happy path", Err.description
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
    Test_InvalidOtpremnicaDoesNotAppend
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
    result = SaveOtkupMulti_TX( _
        NextTestDate(), "", TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 100#, TEST_TIP_AMB, 0, TEST_VOZ_ID, _
        TEST_PREFIX & "-BAD-OTK-" & NewScenarioCode("BAD"), _
        0#, "", "", "", False, 0#, 0#)

    If Len(Trim$(result)) = 0 Then
        AssertEquals CStr(beforeCount), CStr(CountRows(TBL_OTKUP)), _
                     "Invalid otkup did not append row"
        Exit Sub
    End If

    LogFail "Invalid otkup rejected", "SaveOtkupMulti_TX returned ID: " & result
    Exit Sub

ExpectedError:
    AssertEquals CStr(beforeCount), CStr(CountRows(TBL_OTKUP)), _
                 "Invalid otkup raised and did not append row"
End Sub

Private Sub Test_InvalidOtpremnicaDoesNotAppend()
    On Error GoTo ExpectedError

    Dim beforeCount As Long
    beforeCount = CountRows(TBL_OTPREMNICA)

    Dim result As String
    result = SaveOtpremnica_TX(NextTestDate(), "", TEST_VOZ_ID, TEST_PREFIX & "-BAD-OTP-" & NewScenarioCode("BAD"), _
                               TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("BAD"), _
                               TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 1, "I")

    If Len(Trim$(result)) = 0 Then
        AssertEquals CStr(beforeCount), CStr(CountRows(TBL_OTPREMNICA)), _
                     "Invalid otpremnica did not append row"
        Exit Sub
    End If

    LogFail "Invalid otpremnica rejected", "SaveOtpremnica_TX returned ID: " & result
    Exit Sub

ExpectedError:
    AssertEquals CStr(beforeCount), CStr(CountRows(TBL_OTPREMNICA)), _
                 "Invalid otpremnica raised and did not append row"
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
    result = SaveOtkup_TX( _
        NextTestDate(), TEST_KOOP_ID, TEST_ST_ID, _
        TEST_VRSTA, TEST_SORTA, _
        100#, -1#, TEST_TIP_AMB, 1, _
        TEST_VOZ_ID, TEST_PREFIX & "-BAD-OTK-" & NewScenarioCode("NEGPRICE"), _
        0#, "TEST OPERATOR", KLASA_I, GetTestParcelaID(), _
        TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("NEGPRICE"))

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
    result = SaveOtkup_TX( _
        NextTestDate(), TEST_KOOP_ID, TEST_ST_ID, _
        TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 1, _
        TEST_VOZ_ID, TEST_PREFIX & "-BAD-OTK-" & NewScenarioCode("BADCLASS"), _
        0#, "TEST OPERATOR", "BAD", GetTestParcelaID(), _
        TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("BADCLASS"))

    AssertEquals "", result, "Invalid otkup class returns empty"
    AssertEquals CStr(beforeOtkup), CStr(CountRows(TBL_OTKUP)), _
                 "Invalid otkup class did not append row"

    Exit Sub

EH:
    LogFail "Invalid otkup invalid class", Err.description
End Sub

Private Sub Test_DokumentaInputValidationHardening()
    On Error GoTo EH

    Test_InvalidOtpremnicaNegativeCenaDoesNotAppend
    Test_InvalidOtpremnicaMissingAmbTypeDoesNotAppend
    Test_InvalidZbirnaInvalidClassDoesNotAppend
    Test_InvalidPrijemnicaNegativeAmbalazaDoesNotAppend
    Test_PrijemnicaMissingZbirnaDoesNotAppend

    Exit Sub

EH:
    LogFail "Dokumenta input validation hardening", Err.description
End Sub

Private Sub Test_InvalidOtpremnicaNegativeCenaDoesNotAppend()
    On Error GoTo EH

    Dim beforeCount As Long
    beforeCount = CountRows(TBL_OTPREMNICA)

    Dim result As String
    result = SaveOtpremnica_TX( _
        NextTestDate(), TEST_ST_ID, TEST_VOZ_ID, _
        TEST_PREFIX & "-BAD-OTP-" & NewScenarioCode("NEGPRICE"), _
        TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("NEGPRICE"), _
        TEST_VRSTA, TEST_SORTA, _
        100#, -1#, TEST_TIP_AMB, 1, KLASA_I)

    AssertEquals "", result, "Invalid otpremnica negative cena returns empty"
    AssertEquals CStr(beforeCount), CStr(CountRows(TBL_OTPREMNICA)), _
                 "Invalid otpremnica negative cena did not append row"

    Exit Sub

EH:
    LogFail "Invalid otpremnica negative cena", Err.description
End Sub

Private Sub Test_InvalidOtpremnicaMissingAmbTypeDoesNotAppend()
    On Error GoTo EH

    Dim beforeOtp As Long
    Dim beforeAmb As Long

    beforeOtp = CountRows(TBL_OTPREMNICA)
    beforeAmb = CountRows(TBL_AMBALAZA)

    Dim result As String
    result = SaveOtpremnica_TX( _
        NextTestDate(), TEST_ST_ID, TEST_VOZ_ID, _
        TEST_PREFIX & "-BAD-OTP-" & NewScenarioCode("NOAMBTYPE"), _
        TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("NOAMBTYPE"), _
        TEST_VRSTA, TEST_SORTA, _
        100#, 10#, "", 1, KLASA_I)

    AssertEquals "", result, "Invalid otpremnica missing amb type returns empty"
    AssertEquals CStr(beforeOtp), CStr(CountRows(TBL_OTPREMNICA)), _
                 "Invalid otpremnica missing amb type did not append otpremnica"
    AssertEquals CStr(beforeAmb), CStr(CountRows(TBL_AMBALAZA)), _
                 "Invalid otpremnica missing amb type did not append ambalaza"

    Exit Sub

EH:
    LogFail "Invalid otpremnica missing amb type", Err.description
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

Private Sub Test_DokumentaReadHelpersExcludeStornirano()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("STOFILTER")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    Dim brojOtpActive As String
    Dim brojOtpStorno As String
    Dim brojPrijActive As String
    Dim brojPrijStorno As String

    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario
    brojOtpActive = TEST_PREFIX & "-OTP-A-" & scenario
    brojOtpStorno = TEST_PREFIX & "-OTP-S-" & scenario
    brojPrijActive = TEST_PREFIX & "-PRJ-A-" & scenario
    brojPrijStorno = TEST_PREFIX & "-PRJ-S-" & scenario

    Dim otpActive As String
    Dim otpStorno As String
    Dim zbrActive As String
    Dim zbrStorno As String
    Dim prjActive As String
    Dim prjStorno As String

    otpActive = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtpActive, brojZbirne, _
                                  TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 1, KLASA_I)

    otpStorno = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtpStorno, brojZbirne, _
                                  TEST_VRSTA, TEST_SORTA, 200#, 10#, TEST_TIP_AMB, 1, KLASA_I)

    zbrActive = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                              "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                              100#, TEST_TIP_AMB, 1, KLASA_I)

    zbrStorno = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                              "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                              200#, TEST_TIP_AMB, 1, KLASA_I)

    prjActive = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrijActive, brojZbirne, _
                                  TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 1, 0, KLASA_I)

    prjStorno = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrijStorno, brojZbirne, _
                                  TEST_VRSTA, TEST_SORTA, 200#, 10#, TEST_TIP_AMB, 1, 0, KLASA_I)

    AssertTrue Len(otpActive) > 0 And Len(otpStorno) > 0, "Storno filter fixture otpremnice created"
    AssertTrue Len(zbrActive) > 0 And Len(zbrStorno) > 0, "Storno filter fixture zbirne created"
    AssertTrue Len(prjActive) > 0 And Len(prjStorno) > 0, "Storno filter fixture prijemnice created"

    MarkTestRowStornirano TBL_OTPREMNICA, "OtpremnicaID", otpStorno
    MarkTestRowStornirano TBL_ZBIRNA, "ZbirnaID", zbrStorno
    MarkTestRowStornirano TBL_PRIJEMNICA, "PrijemnicaID", prjStorno

    AssertFalse ArrayContainsKeyValue(GetOtpremniceByZbirna(brojZbirne), TBL_OTPREMNICA, _
                                      "OtpremnicaID", otpStorno), _
                "GetOtpremniceByZbirna excludes stornirano"

    AssertFalse ArrayContainsKeyValue(GetOtpremniceByStation(TEST_ST_ID, testDate, testDate), TBL_OTPREMNICA, _
                                      "OtpremnicaID", otpStorno), _
                "GetOtpremniceByStation excludes stornirano"

    AssertFalse ArrayContainsKeyValue(GetZbirnaByKupac(TEST_KUP_ID, testDate, testDate), TBL_ZBIRNA, _
                                      "ZbirnaID", zbrStorno), _
                "GetZbirnaByKupac excludes stornirano"

    AssertFalse ArrayContainsKeyValue(GetPrijemniceByKupac(TEST_KUP_ID, testDate, testDate), TBL_PRIJEMNICA, _
                                      "PrijemnicaID", prjStorno), _
                "GetPrijemniceByKupac excludes stornirano"

    Exit Sub

EH:
    LogFail "Dokumenta read helpers exclude stornirano", Err.description
End Sub

Private Sub Test_OtkupReadHelpersExcludeStornirano()
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

    activeID = SaveOtkup_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, _
        TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 1, _
        TEST_VOZ_ID, brojActive, _
        0#, "TEST OPERATOR", KLASA_I, GetTestParcelaID(), brojZbirne)

    stornoID = SaveOtkup_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, _
        TEST_VRSTA, TEST_SORTA, _
        200#, 10#, TEST_TIP_AMB, 1, _
        TEST_VOZ_ID, brojStorno, _
        0#, "TEST OPERATOR", KLASA_I, GetTestParcelaID(), brojZbirne)

    AssertTrue Len(activeID) > 0 And Len(stornoID) > 0, _
               "Otkup storno filter fixture rows created"

    MarkTestRowStornirano TBL_OTKUP, "OtkupID", stornoID

    AssertFalse ArrayContainsKeyValue(GetOtkupByStation(TEST_ST_ID, testDate, testDate), _
                                      TBL_OTKUP, "OtkupID", stornoID), _
                "GetOtkupByStation excludes stornirano"

    AssertFalse ArrayContainsKeyValue(GetOtkupByKooperant(TEST_KOOP_ID, testDate, testDate), _
                                      TBL_OTKUP, "OtkupID", stornoID), _
                "GetOtkupByKooperant excludes stornirano"

    Exit Sub

EH:
    LogFail "Otkup read helpers exclude stornirano", Err.description
End Sub

Private Sub Test_DualClassDocumentWrappers()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("DOCMULTI")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojOtp As String
    Dim brojZbirne As String
    Dim brojPrij As String

    brojOtp = TEST_PREFIX & "-OTP-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-" & scenario

    Dim beforeOtp As Long
    Dim beforeZbr As Long
    Dim beforePrj As Long

    beforeOtp = CountRows(TBL_OTPREMNICA)
    beforeZbr = CountRows(TBL_ZBIRNA)
    beforePrj = CountRows(TBL_PRIJEMNICA)

    Dim otpResult As String
    otpResult = SaveOtpremnicaMulti_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                                       TEST_VRSTA, TEST_SORTA, 111#, 10#, TEST_TIP_AMB, 5, _
                                       True, 22#, 8#)

    AssertTrue Len(otpResult) > 0, "SaveOtpremnicaMulti_TX returns IDs"
    AssertEquals CStr(beforeOtp + 2), CStr(CountRows(TBL_OTPREMNICA)), _
                 "SaveOtpremnicaMulti_TX appends two rows"

    Dim otpI As String
    Dim otpII As String
    otpI = FindOtpremnicaIDByBrojAndKlasa(brojOtp, KLASA_I)
    otpII = FindOtpremnicaIDByBrojAndKlasa(brojOtp, KLASA_II)

    AssertTrue Len(otpI) > 0, "Dual otpremnica class I found"
    AssertTrue Len(otpII) > 0, "Dual otpremnica class II found"

    AssertEquals "5", CStr(GetValueByKey(TBL_OTPREMNICA, "OtpremnicaID", otpI, "KolAmbalaze")), _
                 "Otpremnica class I carries ambalaza"

    AssertEquals "0", CStr(GetValueByKey(TBL_OTPREMNICA, "OtpremnicaID", otpII, "KolAmbalaze")), _
                 "Otpremnica class II carries zero ambalaza"

    Dim zbrResult As String
    zbrResult = SaveZbirnaMulti_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                                   "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                                   111#, TEST_TIP_AMB, 5, True, 22#)

    AssertTrue Len(zbrResult) > 0, "SaveZbirnaMulti_TX returns IDs"
    AssertEquals CStr(beforeZbr + 2), CStr(CountRows(TBL_ZBIRNA)), _
                 "SaveZbirnaMulti_TX appends two rows"

    Dim zbrI As String
    Dim zbrII As String
    zbrI = FindZbirnaIDByBrojAndKlasa(brojZbirne, KLASA_I)
    zbrII = FindZbirnaIDByBrojAndKlasa(brojZbirne, KLASA_II)

    AssertTrue Len(zbrI) > 0, "Dual zbirna class I found"
    AssertTrue Len(zbrII) > 0, "Dual zbirna class II found"

    AssertEquals "5", CStr(GetValueByKey(TBL_ZBIRNA, "ZbirnaID", zbrI, "UkupnoAmbalaze")), _
                 "Zbirna class I carries ambalaza"

    AssertEquals "0", CStr(GetValueByKey(TBL_ZBIRNA, "ZbirnaID", zbrII, "UkupnoAmbalaze")), _
                 "Zbirna class II carries zero ambalaza"

    Dim prjResult As String
    prjResult = SavePrijemnicaMulti_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                                       TEST_VRSTA, TEST_SORTA, 111#, 10#, TEST_TIP_AMB, 5, 4, _
                                       True, 22#, 8#)

    AssertTrue Len(prjResult) > 0, "SavePrijemnicaMulti_TX returns IDs"
    AssertEquals CStr(beforePrj + 2), CStr(CountRows(TBL_PRIJEMNICA)), _
                 "SavePrijemnicaMulti_TX appends two rows"

    Dim prjI As String
    Dim prjII As String
    prjI = FindPrijemnicaIDByBrojAndKlasa(brojPrij, KLASA_I)
    prjII = FindPrijemnicaIDByBrojAndKlasa(brojPrij, KLASA_II)

    AssertTrue Len(prjI) > 0, "Dual prijemnica class I found"
    AssertTrue Len(prjII) > 0, "Dual prijemnica class II found"

    AssertEquals "5", CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjI, "KolAmbalaze")), _
                 "Prijemnica class I carries ambalaza"

    AssertEquals "0", CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjII, "KolAmbalaze")), _
                 "Prijemnica class II carries zero ambalaza"

    AssertEquals "4", CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjI, "KolAmbVracena")), _
                 "Prijemnica class I carries returned ambalaza"

    AssertEquals "0", CStr(GetValueByKey(TBL_PRIJEMNICA, "PrijemnicaID", prjII, "KolAmbVracena")), _
                 "Prijemnica class II carries zero returned ambalaza"

    Exit Sub

EH:
    LogFail "Dual-class document wrappers", Err.description
End Sub


' ============================================================
' TRACEABILITY / AUTOLINK REGRESSION TESTS
' ============================================================

Private Sub Test_AutoLinkPositiveUniqueMatch()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("LINKOK")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojOtk As String
    Dim brojOtp As String
    Dim brojZbirne As String

    brojOtk = TEST_PREFIX & "-OTK-" & scenario
    brojOtp = TEST_PREFIX & "-OTP-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim otkupID As String
    Dim otpID As String

    Dim otkupResult As String
    otkupResult = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 100#, TEST_TIP_AMB, 10, TEST_VOZ_ID, brojOtk, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        False, 0#, 0#)

    otkupID = FindOtkupIDByBrojAndKlasa(brojOtk, "I")

    otpID = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                              TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 10, "I")

    AssertTrue Len(otkupID) > 0, "Positive autolink fixture otkup exists"
    AssertTrue Len(otpID) > 0, "Positive autolink fixture otpremnica exists"

    AutoLinkOtkupOtpremnica_TX

    AssertEquals otpID, CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkupID, "OtpremnicaID")), _
                 "Positive autolink links exact unique scenario"

    Exit Sub

EH:
    LogFail "Auto-link positive unique match", Err.description
End Sub

Private Sub Test_AutoLinkMustNotCrossBrojZbirne()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("LINKBUG")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojOtkA As String
    Dim brojOtkB As String
    Dim brojOtpB As String

    Dim brojZbrA As String
    Dim brojZbrB As String

    brojOtkA = TEST_PREFIX & "-OTK-A-" & scenario
    brojOtkB = TEST_PREFIX & "-OTK-B-" & scenario
    brojOtpB = TEST_PREFIX & "-OTP-B-" & scenario

    brojZbrA = TEST_PREFIX & "-ZBR-A-" & scenario
    brojZbrB = TEST_PREFIX & "-ZBR-B-" & scenario

    ' Two otkup rows share Station/Date/Vozac/Class but have different BrojZbirne.
    ' Only B has matching otpremnica. A must remain unlinked.
    Dim resA As String
    Dim resB As String

    resA = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 100#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojOtkA, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbrA, _
        False, 0#, 0#)

    resB = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 100#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojOtkB, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbrB, _
        False, 0#, 0#)

    Dim otkA As String
    Dim otkB As String

    otkA = FindOtkupIDByBrojAndKlasa(brojOtkA, "I")
    otkB = FindOtkupIDByBrojAndKlasa(brojOtkB, "I")

    Dim otpB As String
    otpB = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtpB, brojZbrB, _
                             TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 0, "I")

    AssertTrue Len(otkA) > 0, "Cross-zbirna fixture A otkup exists"
    AssertTrue Len(otkB) > 0, "Cross-zbirna fixture B otkup exists"
    AssertTrue Len(otpB) > 0, "Cross-zbirna fixture B otpremnica exists"

    AutoLinkOtkupOtpremnica_TX

    Dim linkA As String
    Dim linkB As String

    linkA = CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkA, "OtpremnicaID"))
    linkB = CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkB, "OtpremnicaID"))

    AssertEquals "", linkA, _
                 "Auto-link must NOT link otkup with different BrojZbirne"

    AssertEquals otpB, linkB, _
                 "Auto-link should link matching BrojZbirne row"

    Exit Sub

EH:
    LogFail "Auto-link must not cross BrojZbirne", Err.description
End Sub

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
' MALINA MOD -- D: auto-zbirna iz otpremnice (1:1; BrojZbirne==BrojOtpremnice)
' ============================================================
Private Sub Test_MalinaAutoZbirnaFromOtpremnice()
    Dim prevMode As String, prevKupac As String

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("MALINA")

    Dim testDate As Date
    testDate = NextTestDate()

    ' Testovi inace ne diraju config -> sacuvaj pa vrati.
    prevMode = GetConfigValue(CFG_KEY_MALINA_MODE)
    prevKupac = GetConfigValue(CFG_MALINA_DEFAULT_KUPAC)
    SetConfigValue CFG_KEY_MALINA_MODE, "YES"
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, TEST_KUP_ID

    ' Otpremnica (Klasa I + II) sa PRAZNIM BrojZbirne (malina konvencija).
    Dim brojOtp As String
    brojOtp = TEST_PREFIX & "-MAL-" & scenario

    Dim otpResult As String
    otpResult = SaveOtpremnicaMulti_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, "", _
                                       TEST_VRSTA, TEST_SORTA, 1000#, 100#, TEST_TIP_AMB, 50, _
                                       True, 200#, 90#)
    AssertTrue Len(otpResult) > 0, "Malina: otpremnica I+II sa praznim BrojZbirne snimljena"

    ' Act
    Dim created As Long
    created = AutoCreateZbirnaFromOtpremnice(brojOtp)   ' scoped na sopstvenu otpremnicu
    AssertTrue created >= 1, "Malina: AutoCreateZbirnaFromOtpremnice kreirao zbirnu"

    ' BrojZbirne == BrojOtpremnice; zbirna I i II postoje
    Dim zbrI As String, zbrII As String
    zbrI = FindZbirnaIDByBrojAndKlasa(brojOtp, KLASA_I)
    zbrII = FindZbirnaIDByBrojAndKlasa(brojOtp, KLASA_II)
    AssertTrue Len(zbrI) > 0, "Malina: zbirna Klasa I (BrojZbirne==BrojOtpremnice) postoji"
    AssertTrue Len(zbrII) > 0, "Malina: zbirna Klasa II postoji (hasKlasaII)"

    ' Auto-zbirna pise red po red (dva zasebna SaveZbirna_TX poziva), ali obe klase
    ' dele BrojZbirne -> moraju deliti i generaciju.
    AssertTrue Len(DokGeneracija(TBL_ZBIRNA, COL_ZBR_ID, zbrI)) > 0, _
        "Malina: auto-zbirna ima generaciju"
    AssertEquals DokGeneracija(TBL_ZBIRNA, COL_ZBR_ID, zbrI), _
                 DokGeneracija(TBL_ZBIRNA, COL_ZBR_ID, zbrII), _
        "Malina: obe klase auto-zbirne dele generaciju"

    ' kg zbirne == kg otpremnice (1:1)
    AssertEquals "1000", _
        CStr(GetValueByKey(TBL_ZBIRNA, "ZbirnaID", zbrI, "UkupnoKolicina")), _
        "Malina: kg Klasa I zbirne == otpremnica"

    ' Backfill BrojZbirne na otpremnicu (GetOtpremniceByZbirna mora vratiti redove)
    AssertTrue Not IsEmpty(GetOtpremniceByZbirna(brojOtp)), _
        "Malina: BrojZbirne backfilovan na otpremnicu"

    ' Idempotencija: ponovni poziv ne pravi novu zbirnu
    Dim zbrBefore As Long
    zbrBefore = CountRows(TBL_ZBIRNA)
    Call AutoCreateZbirnaFromOtpremnice(brojOtp)
    AssertEquals CStr(zbrBefore), CStr(CountRows(TBL_ZBIRNA)), _
        "Malina: ponovni poziv ne duplira zbirnu (idempotentno)"

    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    Exit Sub

EH:
    On Error Resume Next
    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    On Error GoTo 0
    LogFatal "Test_MalinaAutoZbirnaFromOtpremnice", Err.Number, Err.description
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

' AUD-043(a): otkupi istog Stanica|Datum|Vozac|Klasa koji se razlikuju po BILO KOM
' artikal-atributu moraju dati ZASEBNE otpremnice. Stari kljuc (bez
' Vrsta|Sorta|Cena|TipAmb) ih je spajao u jednu i citao metadata sa PRVOG reda ->
' pogresna vrsta, sorta, novac i ambalaza na otpremnici.
'
' Testiraju se sva cetiri polja zasebno (jedna promenljiva po redu, baseline je
' red A) -- da regresija u samo jednom segmentu kljuca ne prode neopazeno.
Private Sub Test_RF28_AutoOtpremnicaNeMesaArtikle()
    Dim tx As clsTransaction

    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("RF28OTP")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim otkBase As String, otkCena As String, otkVrsta As String
    Dim otkSorta As String, otkTipAmb As String

    otkBase = "OTK-RF28-BASE-" & scenario
    otkCena = "OTK-RF28-CENA-" & scenario
    otkVrsta = "OTK-RF28-VRSTA-" & scenario
    otkSorta = "OTK-RF28-SORTA-" & scenario
    otkTipAmb = "OTK-RF28-AMB-" & scenario

    Dim vrstaB As String, sortaB As String, tipAmbB As String
    vrstaB = TEST_VRSTA & " RF28-2"
    sortaB = TEST_SORTA & " RF28-2"
    tipAmbB = TEST_TIP_AMB & " RF28-2"

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    ' Svi dele Stanica|Datum|Vozac|Klasa; svaki se od baseline-a razlikuje po
    ' TACNO JEDNOM artikal-atributu.
    AppendRF28OtkupFixture otkBase, testDate, TEST_VOZ_ID, "I", 120#, ""
    AppendRF28OtkupFixture otkCena, testDate, TEST_VOZ_ID, "I", 175#, ""
    AppendRF28OtkupFixture otkVrsta, testDate, TEST_VOZ_ID, "I", 120#, "", "", vrstaB
    AppendRF28OtkupFixture otkSorta, testDate, TEST_VOZ_ID, "I", 120#, "", "", TEST_VRSTA, sortaB
    AppendRF28OtkupFixture otkTipAmb, testDate, TEST_VOZ_ID, "I", 120#, "", "", TEST_VRSTA, TEST_SORTA, tipAmbB

    ' Scope na test-dan -- run ne sme da zahvati nepovezane otkupe u svesci.
    Call AutoCreateOtpremniceFromPWA_TX(testDate)

    Dim otpBase As String, otpCena As String, otpVrsta As String
    Dim otpSorta As String, otpTipAmb As String

    otpBase = RF28OtpremnicaZaOtkup(otkBase)
    otpCena = RF28OtpremnicaZaOtkup(otkCena)
    otpVrsta = RF28OtpremnicaZaOtkup(otkVrsta)
    otpSorta = RF28OtpremnicaZaOtkup(otkSorta)
    otpTipAmb = RF28OtpremnicaZaOtkup(otkTipAmb)

    AssertTrue Len(otpBase) > 0 And Len(otpCena) > 0 And Len(otpVrsta) > 0 _
               And Len(otpSorta) > 0 And Len(otpTipAmb) > 0, _
        "RF-28 AUD-043a: svih pet otkupa je povezano na otpremnicu"

    ' Pet razlicitih kombinacija -> pet RAZLICITIH otpremnica.
    Dim jedinstvene As Object
    Set jedinstvene = CreateObject("Scripting.Dictionary")
    jedinstvene(otpBase) = True
    jedinstvene(otpCena) = True
    jedinstvene(otpVrsta) = True
    jedinstvene(otpSorta) = True
    jedinstvene(otpTipAmb) = True

    AssertEquals "5", CStr(jedinstvene.count), _
        "RF-28 AUD-043a: pet artikal-kombinacija daje PET otpremnica (ne jednu mesanu)"

    ' Svaka otpremnica nosi SVOJ atribut, ne onaj sa prvog reda grupe.
    AssertDoubleNear 120#, CDbl(GetValueByKey(TBL_OTPREMNICA, COL_OTP_ID, otpBase, COL_OTP_CENA)), _
        0.001, "RF-28 AUD-043a: baseline otpremnica nosi svoju cenu"
    AssertDoubleNear 175#, CDbl(GetValueByKey(TBL_OTPREMNICA, COL_OTP_ID, otpCena, COL_OTP_CENA)), _
        0.001, "RF-28 AUD-043a: razlicita Cena je zasebna otpremnica sa svojom cenom"
    AssertEquals vrstaB, Trim$(CStr(GetValueByKey(TBL_OTPREMNICA, COL_OTP_ID, otpVrsta, COL_OTP_VRSTA))), _
        "RF-28 AUD-043a: razlicita VrstaVoca je zasebna otpremnica sa svojom vrstom"
    AssertEquals sortaB, Trim$(CStr(GetValueByKey(TBL_OTPREMNICA, COL_OTP_ID, otpSorta, COL_OTP_SORTA))), _
        "RF-28 AUD-043a: razlicita SortaVoca je zasebna otpremnica sa svojom sortom"
    AssertEquals tipAmbB, Trim$(CStr(GetValueByKey(TBL_OTPREMNICA, COL_OTP_ID, otpTipAmb, COL_OTP_TIP_AMB))), _
        "RF-28 AUD-043a: razlicit TipAmbalaze je zasebna otpremnica sa svojim tipom"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "RF-28 AUD-043a auto-otpremnica ne mesa artikle", bfpErrDesc
End Sub

Private Function RF28OtpremnicaZaOtkup(ByVal otkupID As String) As String
    RF28OtpremnicaZaOtkup = _
        Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_OTPREMNICA_ID)))
End Function

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
    AppendRF28OtkupFixture otkID, testDate, TEST_VOZ_ID, "I", 100#, crid, brojA
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

    AppendRF28OtkupFixture otkID, testDate, TEST_VOZ_ID, "I", 100#, crid

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
    AppendRF28OtkupFixture otkBlizu, zbrDate - 1, TEST_VOZ_ID, "I", 100#, cridBlizu
    TestHook_LinkZbirnaToOtkupAndOtpremnica zbrID, brojZ, cridBlizu

    AssertEquals brojZ, Trim$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkBlizu, COL_OTK_BROJ_ZBIRNE))), _
        "RF-28 AUD-043b: otkup od prethodnog dana prolazi (post-midnight)"

    ' 10 dana razlike -> nije membership.
    AppendRF28OtkupFixture otkDaleko, zbrDate - 10, TEST_VOZ_ID, "I", 100#, cridDaleko

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

    AppendRF28OtkupFixture otkPrazan, testDate, "", "I", 100#, cridPrazan
    AppendRF28OtkupFixture otkZauzet, testDate, TEST_VOZ_ID, "I", 100#, cridZauzet
    AppendRF28OtkupFixture otkPad, testDate, "", "I", 100#, cridPad

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
                                   ByVal klasa As String, _
                                   ByVal cena As Double, _
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
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_KOLICINA, 100#
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_CENA, cena
    SetRequiredField rowData, TBL_OTKUP, COL_OTK_KLASA, klasa
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_KULTURA, TEST_KULTURA_ID
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_TIP_AMB, tipAmb
    SetOptionalField rowData, TBL_OTKUP, COL_OTK_KOL_AMB, 0
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

Private Sub Test_ProsekGajbeExcludesStornirano()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("PROSGAJ")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojOtp As String, brojZbirne As String
    brojOtp = TEST_PREFIX & "-OTP-PG-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-PG-" & scenario

    ' Dvoklasna otpremnica: (100+200) kg / (10+10) gajbi = 15 kg po gajbi.
    Dim otpI As String, otpII As String
    otpI = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    otpII = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                              TEST_VRSTA, TEST_SORTA, 200#, 10#, TEST_TIP_AMB, 10, KLASA_II)

    AssertTrue Len(otpI) > 0 And Len(otpII) > 0, "Prosek gajbe: fixture otpremnica I+II kreirana"
    AssertTrue Abs(CalculateProsekGajbe(brojOtp) - 15#) < 0.001, _
               "Prosek gajbe (otpremnica) pre storna = 15"

    MarkTestRowStornirano TBL_OTPREMNICA, "OtpremnicaID", otpII

    ' Posle storna Kl.II ostaje samo 100 kg / 10 gajbi = 10.
    AssertTrue Abs(CalculateProsekGajbe(brojOtp) - 10#) < 0.001, _
               "Prosek gajbe (otpremnica) ne racuna stornirani red"

    ' Isto na zbirnoj (CalculateProsekGajbeByZbirna -> isti SumByBroj).
    Dim zbrI As String, zbrII As String
    zbrI = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    zbrII = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          200#, TEST_TIP_AMB, 10, KLASA_II)

    AssertTrue Len(zbrI) > 0 And Len(zbrII) > 0, "Prosek gajbe: fixture zbirna I+II kreirana"
    AssertTrue Abs(CalculateProsekGajbeByZbirna(brojZbirne) - 15#) < 0.001, _
               "Prosek gajbe (zbirna) pre storna = 15"

    MarkTestRowStornirano TBL_ZBIRNA, "ZbirnaID", zbrII

    AssertTrue Abs(CalculateProsekGajbeByZbirna(brojZbirne) - 10#) < 0.001, _
               "Prosek gajbe (zbirna) ne racuna stornirani red"

    Exit Sub

EH:
    LogFatal "Test_ProsekGajbeExcludesStornirano", Err.Number, Err.description
End Sub

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

Private Sub Test_ZbirnaKlasaIIGuard()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("KLIIGUARD")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    brojZbirne = TEST_PREFIX & "-ZBR-K2-" & scenario

    AssertFalse ZbirnaIzvorImaKlasuII(""), "Kl.II guard: prazan broj zbirne ne blokira"

    Dim otpI As String
    otpI = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-K2A-" & scenario, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    AssertTrue Len(otpI) > 0, "Kl.II guard: fixture otpremnica Kl.I kreirana"
    AssertFalse ZbirnaIzvorImaKlasuII(brojZbirne), "Kl.II guard: izvor samo sa Kl.I ne blokira"

    Dim otpII As String
    otpII = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                              TEST_PREFIX & "-OTP-K2B-" & scenario, brojZbirne, _
                              TEST_VRSTA, TEST_SORTA, 50#, 8#, TEST_TIP_AMB, 5, KLASA_II)
    AssertTrue Len(otpII) > 0, "Kl.II guard: fixture otpremnica Kl.II kreirana"
    AssertTrue ZbirnaIzvorImaKlasuII(brojZbirne), _
               "Kl.II guard: izvor sa Kl.II blokira unos bez 'Dve klase'"

    MarkTestRowStornirano TBL_OTPREMNICA, "OtpremnicaID", otpII
    AssertFalse ZbirnaIzvorImaKlasuII(brojZbirne), _
                "Kl.II guard: stornirana Kl.II otpremnica ne blokira"

    ' Posledica koju blokada sprecava: hasKlasaII:=False tiho odbacuje Kl.II izvor.
    Dim brojZbirne2 As String
    brojZbirne2 = TEST_PREFIX & "-ZBR-K2X-" & scenario

    Dim zbrRes As String
    zbrRes = SaveZbirnaMulti_TX(datum:=testDate, vozacID:=TEST_VOZ_ID, _
                                brojZbirne:=brojZbirne2, kupacID:=TEST_KUP_ID, _
                                hladnjaca:="Test Hladnjaca", pogon:="Test Pogon", _
                                vrstaVoca:=TEST_VRSTA, sortaVoca:=TEST_SORTA, _
                                ukupnoKolI:=100#, tipAmb:=TEST_TIP_AMB, ukupnoAmb:=10, _
                                hasKlasaII:=False, ukupnoKolII:=50#, ukupnoAmbII:=5)

    AssertTrue Len(zbrRes) > 0, "Kl.II guard: kontrolna zbirna (hasKlasaII=False) snimljena"
    AssertEquals "", FindZbirnaIDByBrojAndKlasa(brojZbirne2, KLASA_II), _
                 "Kl.II guard: bez 'Dve klase' Kl.II se NE upisuje (zato blokada)"

    Exit Sub

EH:
    LogFatal "Test_ZbirnaKlasaIIGuard", Err.Number, Err.description
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

' Guard mora da vazi na SVIM number-only putanjama, ne samo na direktnom
' StornoPrijemnicaByBroj_TX: ISPRAVKA/DUPLI otpremnice idu kroz atomic helper,
' a SIMPLE/DUPLI zbirna kroz core StornoZbirna.
' ZBR-MUT-01: mutacija po BROJU staje kad broj nosi DVA AKTIVNA dokumenta,
' makar bili istog vlasnika.
'
' Zatecena kapija je brojala VLASNIKE, pa je ovo stanje prolazilo. Posledice su
' bile razlicite po putanji, a obe destruktivne PREKO granice dokumenta:
'   SIMPLE  -- zaglavlje se stornira tacno (po generaciji), ali
'              DetachOtpremniceInline nize ide po BROJU i prazni BrojZbirne
'              deci OBA dokumenta;
'   ISPRAVKA -- relink i rekalkulacija po broju zahvataju oba.
'
' Stanje pravi PRAVI uvoz (dva ClientRecordID-a), jer je bas on jedini put kojim
' redovno nastaje: F3 kapija ga ne pusta, a Excel writer dva reda istog broja i
' vlasnika stapa u JEDAN dokument. Zato je i negativna kontrola dole bas taj
' slucaj -- da kapija ne pocne da odbija dvoklasnu zbirnu.
' ZBR-CHILD-01 (Faza 1): generacija roditelja se na detetu menja U KORAKU sa
' BrojZbirne -- i kad se postavlja, i kad se brise.
'
' Meri se oba smera i oba ishoda razresenja:
'   roditelj postoji i jednoznacan  -> dete nosi NJEGOVU generaciju
'   roditelja nema (dete pre zbirne) -> dete nosi PRAZNO, ne pogodjenu vrednost
'   odvezivanje                      -> i broj i generacija prazni
'
' Treca grana je razlog zasto ova kolona uopste moze da se uvede postepeno:
' prazno je legitimno stanje i znaci "citaj po broju", pa Faza 1 ne menja nista
' za citaoce. Bez te tvrdnje bi neko kasnije "popravio" prazno na pogadjanje.
Private Sub Test_ZBR_DeteNosiGeneracijuRoditelja()
    Dim tx As clsTransaction
    Dim testDate As Date
    Dim scenario As String
    Dim broj As String, brojBezZbirne As String
    Dim zbrID As String, genZbr As String
    Dim otpSaRod As String, otpBezRod As String
    Dim r As Object

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRCHILD")
    testDate = NextTestDate()
    broj = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    brojBezZbirne = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & _
                    Format$(NextTestDate(), "ddmmyy")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP

    ' --- A) roditelj postoji: dete nosi njegovu generaciju ---
    zbrID = SaveZbirna_TX(testDate, TEST_VOZ_ID, broj, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          100#, TEST_TIP_AMB, 10, KLASA_I)
    genZbr = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrID)
    AssertTrue Len(genZbr) > 0, "ZBR-CHILD preduslov: zbirna nosi svoju generaciju"

    otpSaRod = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                                 TEST_PREFIX & "-OTP-CHLD-A-" & scenario, broj, _
                                 TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    AssertTrue Len(otpSaRod) > 0, "ZBR-CHILD preduslov: otpremnica sa roditeljem je snimljena"
    AssertEquals genZbr, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpSaRod, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-CHILD: dete nosi generaciju roditelja"

    ' --- B) roditelja NEMA: prazno, ne pogodjeno ---
    otpBezRod = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                                  TEST_PREFIX & "-OTP-CHLD-B-" & scenario, brojBezZbirne, _
                                  TEST_VRSTA, TEST_SORTA, 50#, 10#, TEST_TIP_AMB, 5, KLASA_I)
    AssertTrue Len(otpBezRod) > 0, "ZBR-CHILD preduslov: otpremnica bez roditelja je snimljena"
    AssertEquals "", _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpBezRod, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-CHILD: bez roditelja generacija ostaje PRAZNA"
    AssertEquals brojBezZbirne, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpBezRod, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-CHILD: broj se svejedno upisuje (dete pre roditelja je normalno)"

    ' --- C) odvezivanje brise OBOJE ---
    Set r = RunSimpleStornoZbirna(broj)
    AssertTrue CBool(r("success")), "ZBR-CHILD preduslov: storno zbirne je prosao"
    AssertEquals "", _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpSaRod, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-CHILD: odvezivanje brise broj"
    AssertEquals "", _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpSaRod, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-CHILD: odvezivanje brise i generaciju roditelja"

    ' --- D) roditelj STORNIRAN: red pod tim brojem POSTOJI, ali nije aktivan ---
    '
    ' Ovo je slucaj koji razdvaja RAZRESAVANJE od POGADJANJA. Grane B i C ne bi
    ' ga uhvatile: kad zbirne uopste nema, i naivni LookupValue po broju vrati
    ' prazno, pa bi sabotaza koja uvodi pogadjanje prosla neprimeceno. Ovde
    ' pogadjanje vraca generaciju STORNIRANE zbirne, a tacan odgovor je prazno.
    Dim otpPosleStorna As String
    otpPosleStorna = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                                       TEST_PREFIX & "-OTP-CHLD-D-" & scenario, broj, _
                                       TEST_VRSTA, TEST_SORTA, 30#, 10#, TEST_TIP_AMB, 3, KLASA_I)
    AssertTrue Len(otpPosleStorna) > 0, _
        "ZBR-CHILD preduslov: otpremnica pod storniranim brojem je snimljena"
    AssertTrue Len(GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrID)) > 0, _
        "ZBR-CHILD preduslov: stornirana zbirna I DALJE nosi generaciju (ima sta da se pogodi)"
    AssertEquals "", _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpPosleStorna, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-CHILD: stornirana zbirna NIJE roditelj -- generacija ostaje prazna"

    ' --- E) RE-ENTRY: backfill NE SME da veze staro dete na novu generaciju ---
    '
    ' Ugovor par.5 izricito dozvoljava da isti vlasnik posle storna ponovo unese
    ' zbirnu pod ISTIM brojem. Tada pod tim brojem stoje stornirana GEN-A i aktivna
    ' GEN-B, a staro dete (jos bez generacije) istorijski pripada GEN-A.
    '
    ' "Ko je roditelj SADA" tu vraca GEN-B -- tacno za nov upis, POGRESNO za
    ' rekonstrukciju starog reda. Backfill zato pita "ko je IKAD bio pod ovim
    ' brojem" i cuti kad ih je bilo vise. Lazna sledljivost je gora od prazne
    ' kolone: prazna bar ne tvrdi nista.
    Dim zbrB As String, genB As String
    zbrB = SaveZbirna_TX(testDate, TEST_VOZ_ID, broj, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         70#, TEST_TIP_AMB, 7, KLASA_I)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrB)
    AssertTrue (Len(genB) > 0 And genB <> genZbr), _
        "ZBR-CHILD preduslov: re-entry pod istim brojem dao je NOVU generaciju"

    AssertEquals genB, ZbirnaGeneracijaZaBroj(broj), _
        "ZBR-CHILD: 'ko je roditelj SADA' vraca novu generaciju (tacno za nov upis)"
    AssertEquals "", ZbirnaJedinaGeneracijaIkadZaBroj(broj), _
        "ZBR-CHILD: 'ko je IKAD' cuti kad su pod brojem bile DVE generacije"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 dete nosi generaciju roditelja", bfpErrDesc
End Sub

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

' ZBR-CHILD-01 faza 2: backfill rekonstruise identitet STARIH redova.
'
' Zasto poseban test, a ne oslanjanje na granu E testa DeteNosiGeneracijuRoditelja:
' ta grana meri PRIMITIVU (ZbirnaJedinaGeneracijaIkadZaBroj) pozivajuci je
' direktno, a backfill je nikad nije zvao ni u jednom testu. Sabotaza koja je
' menjala njegov izbor kriterijuma zato nije obarala NISTA -- menjala je red koda
' koji suite ne izvrsava. Pokrivena primitiva nije pokriven pozivalac.
'
' Test vozi BackfillDeteZbirnaGeneracija_Core nad dva broja odjednom:
'   X -- pod njim je IKAD bila jedna generacija  -> mora da POPUNI
'   Y -- pod njim su IKAD bile dve (storno + re-entry) -> mora da CUTI
' Grana X je anti-placebo: bez nje bi "ostalo prazno" prolazilo i kad backfill
' uopste nije radio.
Private Sub Test_ZBR_BackfillNeVezeStaroDeteNaNovuGeneraciju()
    Dim tx As clsTransaction
    Dim testDate As Date, scenario As String
    Dim brojX As String, brojY As String
    Dim zbrX As String, zbrYA As String, zbrYB As String
    Dim genX As String, genYB As String
    Dim otpX As String, otpY As String
    Dim popunjeno As Long, preskoceno As Long
    Dim r As Object

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRBF")
    testDate = NextTestDate()
    brojX = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    brojY = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & _
            Format$(NextTestDate(), "ddmmyy")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA_STAVKA
    tx.AddTableSnapshot TBL_OTKUP

    ' --- X: jedna generacija ikad ---
    zbrX = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojX, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    genX = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrX)
    AssertTrue Len(genX) > 0, "ZBR-BACKFILL preduslov: zbirna X nosi generaciju"

    otpX = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-BFX-" & scenario, brojX, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    AssertTrue Len(otpX) > 0, "ZBR-BACKFILL preduslov: otpremnica X je snimljena"

    ' --- Y: dve generacije ikad (storno pa re-entry istog vlasnika, ugovor par.5) ---
    zbrYA = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojY, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          80#, TEST_TIP_AMB, 8, KLASA_I)
    AssertTrue Len(zbrYA) > 0, "ZBR-BACKFILL preduslov: zbirna Y-A je snimljena"

    Set r = RunSimpleStornoZbirna(brojY)
    AssertTrue CBool(r("success")), "ZBR-BACKFILL preduslov: storno zbirne Y-A je prosao"

    zbrYB = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojY, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          70#, TEST_TIP_AMB, 7, KLASA_I)
    genYB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrYB)
    AssertTrue (Len(genYB) > 0 And genYB <> GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrYA)), _
        "ZBR-BACKFILL preduslov: re-entry pod brojem Y dao je NOVU generaciju"

    otpY = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-BFY-" & scenario, brojY, _
                             TEST_VRSTA, TEST_SORTA, 70#, 10#, TEST_TIP_AMB, 7, KLASA_I)
    AssertTrue Len(otpY) > 0, "ZBR-BACKFILL preduslov: otpremnica Y je snimljena"

    ' --- oblik ZATECENOG reda: broj postoji, generacija ne ---
    ' Tacno stanje svakog reda pre migracije. Bez ovog koraka backfill nema sta
    ' da radi (preskace popunjene), pa bi test bio zelen ne merivsi nista.
    IsprazniGeneracijuDeteta TBL_OTPREMNICA, COL_OTP_ID, otpX
    IsprazniGeneracijuDeteta TBL_OTPREMNICA, COL_OTP_ID, otpY

    AssertEquals "", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpX, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-BACKFILL preduslov: red X je u zatecenom obliku (generacija prazna)"
    AssertEquals brojX, NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpX, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-BACKFILL preduslov: red X je zadrzao broj"
    AssertEquals "", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpY, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-BACKFILL preduslov: red Y je u zatecenom obliku (generacija prazna)"
    AssertEquals brojY, NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpY, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-BACKFILL preduslov: red Y je zadrzao broj"
    AssertEquals genYB, ZbirnaGeneracijaZaBroj(brojY), _
        "ZBR-BACKFILL preduslov: 'ko je roditelj SADA' pod Y vraca novu generaciju"

    BackfillDeteZbirnaGeneracija_Core False, popunjeno, preskoceno

    AssertEquals genX, NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpX, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-BACKFILL: jednoznacan broj se popunjava"
    AssertEquals "", NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpY, COL_DETE_ZBIRNA_GEN)), _
        "ZBR-BACKFILL: broj koji je IKAD nosio dve generacije ostaje PRAZAN"
    AssertTrue popunjeno >= 1, "ZBR-BACKFILL preduslov: backfill je nesto upisao"
    AssertTrue preskoceno >= 1, "ZBR-BACKFILL preduslov: backfill je nesto preskocio"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 backfill ne veze staro dete na novu generaciju", bfpErrDesc
End Sub

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
    AppendRF28OtkupFixture otkID, testDate, TEST_VOZ_ID, "I", 100#, crid, ""
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
    AppendRF28OtkupFixture otkID2, testDate, TEST_VOZ_ID, "I", 100#, crid2, ""
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

' ZBR-CHILD-01 faza 3: kaskada dira SVOJU decu, ne svu decu pod tim brojem.
'
' Scenario nije hipotetican nego postoji danas. Creation path:
' `modStornoDok` STIP_ZBIRNA zove `modStorno.StornoZbirna_TX`, koji snapshot-uje
' SAMO tblZbirna i stornira ZAGLAVLJE -- decu ne dira. Zato posle njega postoji
' stornirana zbirna sa jos AKTIVNOM decom.
'
' Kapija ZBR-MUT-01 to NE zaustavlja kad je vlasnik isti: istorijska grana broji
' VLASNIKE (`ikadVl` po ZbirnaVlasnikKljuc), a re-entry istog vozaca i kupca daje
' 1; aktivnih dokumenata je takodje 1, jer je A storniran. Kapija pusta, a Detach
' po broju odvezuje i decu A.
'
' Deo 2 meri fallback: cim jedno dete nema generaciju, suzavanje se ne desava i
' ishod je BIT-IDENTICAN zatecenom -- ukljucujuci i njegovu manu. To je cena koja
' je svesno placena da faza 3 ne pomeri nijednu zatecenu brojku.
Private Sub Test_ZBR_KaskadaNeDiraDecuDrugogDokumenta()
    Dim tx As clsTransaction
    Dim scenario As String, testDate As Date
    Dim brojX As String, brojY As String
    Dim zbrA As String, zbrB As String, genA As String, genB As String
    Dim otpA As String, otpB As String
    Dim zbrC As String, zbrD As String, genC As String
    Dim otpC As String, otpD As String
    Dim r As Object

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRF3")
    testDate = NextTestDate()
    brojX = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    brojY = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & _
            Format$(NextTestDate(), "ddmmyy")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP

    ' ================= DEO 1: sva deca nose generaciju -> suzavanje radi =========
    zbrA = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3-A-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 100, brojX)
    genA = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrA)
    AssertTrue Len(genA) > 0, "ZBR-F3 preduslov: zbirna A nosi generaciju"

    otpA = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3A-" & scenario, brojX, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    AssertEquals genA, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
        "ZBR-F3 preduslov: otpremnica A nosi generaciju A"

    ' Operaterski storno zaglavlja -- deca ostaju AKTIVNA i zadrzavaju broj.
    AssertTrue StornoZbirna_TX(brojX, genA), _
        "ZBR-F3 preduslov: zaglavlje A je stornirano"
    AssertEquals brojX, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpA, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F3 preduslov: otpremnica A je i posle storna zaglavlja jos vezana"

    zbrB = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3-B-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 120, brojX)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrB)
    AssertTrue (Len(genB) > 0 And genB <> genA), _
        "ZBR-F3 preduslov: re-entry istog vlasnika dao je NOVU generaciju"

    otpB = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3B-" & scenario, brojX, _
                             TEST_VRSTA, TEST_SORTA, 80#, 10#, TEST_TIP_AMB, 8, KLASA_I)
    AssertEquals genB, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpB), _
        "ZBR-F3 preduslov: otpremnica B nosi generaciju B"

    ' Kapija PUSTA -- i to je deo nalaza, ne slucajnost.
    ' Jezgro, ne modStornoFlow.ZbirnaMutRazlog: taj je Private i tanak omotac nad
    ' bas ovom funkcijom, pa je iz drugog modula i nedostupan i suvisan.
    AssertEquals "", modDokumenta.ZbirnaMutacijaPoBrojuRazlogZaBroj(brojX), _
        "ZBR-F3 preduslov: kapija ZBR-MUT-01 pusta (isti vlasnik, jedan aktivan)"

    Set r = RunSimpleStornoZbirna(brojX, genB)
    AssertTrue CBool(r("success")), "ZBR-F3 preduslov: storno zbirne B je prosao"

    AssertEquals "", _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpB, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F3 preduslov: sopstvena otpremnica B JESTE odvezana"
    AssertEquals brojX, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpA, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F3: kaskada NE odvezuje dete drugog dokumenta pod istim brojem"
    AssertEquals genA, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
        "ZBR-F3: dete drugog dokumenta zadrzava svoju generaciju"

    ' ================= DEO 2: jedno dete bez generacije -> fallback na broj =======
    zbrC = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3-C-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 100, brojY)
    genC = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrC)
    otpC = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3C-" & scenario, brojY, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    AssertTrue StornoZbirna_TX(brojY, genC), _
        "ZBR-F3 preduslov: zaglavlje C je stornirano"

    zbrD = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3-D-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 120, brojY)
    otpD = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3D-" & scenario, brojY, _
                             TEST_VRSTA, TEST_SORTA, 80#, 10#, TEST_TIP_AMB, 8, KLASA_I)

    ' Zatecen red: broj stoji, generacija ne. Dovoljan je JEDAN takav.
    IsprazniGeneracijuDeteta TBL_OTPREMNICA, COL_OTP_ID, otpD
    AssertEquals "", DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpD), _
        "ZBR-F3 preduslov: otpremnica D je u zatecenom obliku"

    Set r = RunSimpleStornoZbirna(brojY, GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrD))
    AssertTrue CBool(r("success")), "ZBR-F3 preduslov: storno zbirne D je prosao"

    AssertEquals "", _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpC, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F3: jedno dete bez generacije vraca CEO izbor na broj (zatecen ishod)"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 faza 3 kaskada po generaciji", bfpErrDesc
End Sub

' ZBR-CHILD-01 faza 3 / P1: rezim je po OPERACIJI, ne po tabeli.
'
' Kaskada bira decu iz tri skupa. Kad svaki odlucuje sam, jedna poslovna radnja
' zna da bude pola scoped a pola po broju:
'
'   otpremnice: obe nose generaciju -> suzi na GEN-B -> OTP-A prezivi
'   prijemnice: jedna je legacy     -> fallback      -> PRJ-A stornirana
'
' Ovaj test NE tvrdi da je fallback ishod pozeljan -- tvrdi da je JEDINSTVEN.
' Kad bilo koji skup padne na broj, pada CELA operacija; mesavina je gora od oba
' cista rezima, jer ostavlja pola dokumenta.
Private Sub Test_ZBR_RezimJeZaCeluOperacijuNePoTabeli()
    Dim tx As clsTransaction
    Dim scenario As String, testDate As Date
    Dim broj As String
    Dim zbrA As String, zbrB As String, genA As String, genB As String
    Dim otpA As String, otpB As String, prjA As String, prjB As String
    Dim prevKupac As String
    Dim r As Object

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRF3X")
    testDate = NextTestDate()
    broj = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")

    ' Bez ovoga je ownsChain = False, pa kaskada prijemnice UOPSTE ne dira -- i
    ' glavna tvrdnja prolazi ne merivsi nista. Prva verzija testa je bas tako
    ' pala: preduslov "sopstvena prijemnica B je stornirana" je javio da lanac
    ' nije vlasnicki. Testovi inace ne diraju config -> sacuvaj pa vrati.
    prevKupac = GetConfigValue(CFG_MALINA_DEFAULT_KUPAC)
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, TEST_KUP_ID

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

    ' --- dokument A: otpremnica sa generacijom, prijemnica ZATECENA (bez nje) ---
    zbrA = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3X-A-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 100, broj)
    genA = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrA)
    otpA = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3XA-" & scenario, broj, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    prjA = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-PRJ-F3XA-" & scenario, broj, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 0, 0, KLASA_I, 0)
    AssertTrue (Len(otpA) > 0 And Len(prjA) > 0), _
        "ZBR-F3X preduslov: dokument A ima i otpremnicu i prijemnicu"

    IsprazniGeneracijuDeteta TBL_PRIJEMNICA, COL_PRJ_ID, prjA
    AssertEquals genA, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
        "ZBR-F3X preduslov: otpremnica A NOSI generaciju"
    AssertEquals "", DeteGeneracija(TBL_PRIJEMNICA, COL_PRJ_ID, prjA), _
        "ZBR-F3X preduslov: prijemnica A je ZATECENA (bez generacije)"

    AssertTrue StornoZbirna_TX(broj, genA), _
        "ZBR-F3X preduslov: zaglavlje A je stornirano (deca ostaju aktivna)"

    ' --- dokument B: oba deteta nose generaciju ---
    zbrB = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3X-B-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 120, broj)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrB)
    otpB = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3XB-" & scenario, broj, _
                             TEST_VRSTA, TEST_SORTA, 80#, 10#, TEST_TIP_AMB, 8, KLASA_I)
    prjB = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-PRJ-F3XB-" & scenario, broj, _
                             TEST_VRSTA, TEST_SORTA, 80#, 10#, TEST_TIP_AMB, 0, 0, KLASA_I, 0)
    AssertTrue (Len(genB) > 0 And genB <> genA), _
        "ZBR-F3X preduslov: B je NOVA generacija pod istim brojem"
    AssertEquals genB, DeteGeneracija(TBL_PRIJEMNICA, COL_PRJ_ID, prjB), _
        "ZBR-F3X preduslov: prijemnica B nosi generaciju"

    Set r = RunZbirnaCorrection(broj, SV_MODE_PONISTENJE, True, genB)
    AssertTrue CBool(r("success")), "ZBR-F3X preduslov: ponistenje B je proslo"

    ' Sopstvena deca su svakako dirnuta -- bez toga kaskada nije ni radila.
    AssertTrue RedJeStorniran(TBL_OTPREMNICA, COL_OTP_ID, otpB), _
        "ZBR-F3X preduslov: sopstvena otpremnica B je stornirana"
    AssertTrue RedJeStorniran(TBL_PRIJEMNICA, COL_PRJ_ID, prjB), _
        "ZBR-F3X preduslov: sopstvena prijemnica B je stornirana (lanac je vlasnicki)"

    ' JEZGRO: prijemnica A je legacy, pa CELA operacija pada na broj -- ukljucujuci
    ' i otpremnice. Mesavina bi ostavila OTP-A a stornirala PRJ-A.
    AssertEquals StornoOznaka(TBL_PRIJEMNICA, COL_PRJ_ID, prjA), _
                 StornoOznaka(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
        "ZBR-F3X: otpremnica i prijemnica drugog dokumenta zavrse u ISTOM stanju"

    tx.RollbackTx
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 faza 3 rezim po operaciji", bfpErrDesc
End Sub

Private Function StornoOznaka(ByVal tableName As String, ByVal idColumn As String, _
                                      ByVal idValue As String) As String
    StornoOznaka = UCase$(NzToText(LookupValue(tableName, idColumn, idValue, COL_STORNIRANO)))
End Function

Private Function RedJeStorniran(ByVal tableName As String, ByVal idColumn As String, _
                                ByVal idValue As String) As Boolean
    RedJeStorniran = (StornoOznaka(tableName, idColumn, idValue) = "DA")
End Function

' ZBR-CHILD-01 faza 3 / P1: ISPRAVKA uzima identitet STAROG dokumenta.
'
' Lifecycle je: context sa OldDocID -> StornoZbirna_TX -> operater snimi novu ->
' CompleteZbirnaIspravka -> relink. U trenutku relinka stara zbirna VISE NIJE
' AKTIVNA, pa razresavanje po broju ne moze da je nadje.
'
' Prva verzija je bas tu zvala ZbirnaGeneracijaZaBroj(oldBroj). Dva ishoda:
'   nema drugog dokumenta pod tim brojem -> prazno -> suzavanje mrtvo
'   ima aktivnog GEN-B pod istim brojem  -> GEN-B  -> relink precizno izabere
'                                                     POGRESAN dokument
' Drugi je gori od stanja pre faze 3: nekad je prevozio i svoju i tudju decu, a
' tako bi prevezao SAMO tudju, a svoju ostavio. Ovaj test meri bas taj slucaj.
Private Sub Test_ZBR_IspravkaVezeSvojuDecuNeTudju()
    Dim tx As clsTransaction
    Dim scenario As String, testDate As Date
    Dim brojStari As String, brojNovi As String
    Dim zbrA As String, zbrB As String, zbrC As String
    Dim genA As String, genB As String, genC As String
    Dim otpA As String, otpB As String
    Dim cid As String
    Dim r As Object

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRF3I")
    testDate = NextTestDate()
    brojStari = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    brojNovi = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & _
               Format$(NextTestDate(), "ddmmyy")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_STORNO_VEZE

    zbrA = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3I-A-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 100, brojStari)
    genA = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrA)
    otpA = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3IA-" & scenario, brojStari, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    AssertEquals genA, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
        "ZBR-F3I preduslov: otpremnica A nosi generaciju A"

    ' Ispravka: zaglavlje A se stornira, context pamti njegov OldDocID.
    Set r = RunZbirnaCorrection(brojStari, SV_MODE_ISPRAVKA, True, genA)
    cid = CStr(r("correctionID"))
    AssertTrue (CBool(r("success")) And Len(cid) > 0), _
        "ZBR-F3I preduslov: ispravka je otvorena i zaglavlje A stornirano"

    ' IZMEDJU storna i zavrsetka pod ISTIM brojem nastane drugi dokument.
    zbrB = TestHook_ImportZbirnaRowPWA("CRID-ZBRF3I-B-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 120, brojStari)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrB)
    otpB = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-F3IB-" & scenario, brojStari, _
                             TEST_VRSTA, TEST_SORTA, 80#, 10#, TEST_TIP_AMB, 8, KLASA_I)
    AssertEquals genB, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpB), _
        "ZBR-F3I preduslov: otpremnica B nosi generaciju B"
    AssertEquals genB, ZbirnaGeneracijaZaBroj(brojStari), _
        "ZBR-F3I preduslov: razresavanje po STAROM broju sada vraca TUDJU generaciju"

    ' Zamena: nova zbirna pod NOVIM brojem.
    zbrC = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojNovi, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    genC = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrC)
    AssertTrue Len(genC) > 0, "ZBR-F3I preduslov: zamenska zbirna je snimljena"

    Set r = CompleteZbirnaIspravka(cid, brojNovi)
    AssertTrue CBool(r("success")), "ZBR-F3I preduslov: zavrsetak ispravke je prosao"

    AssertEquals brojNovi, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpA, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F3I: ispravka prevezuje SVOJU otpremnicu na novi broj"
    AssertEquals brojStari, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpB, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-F3I: otpremnica drugog dokumenta ostaje NETAKNUTA"
    AssertEquals genB, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpB), _
        "ZBR-F3I: otpremnica drugog dokumenta zadrzava svoju generaciju"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 faza 3 ispravka veze svoju decu", bfpErrDesc
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
    AppendRF28OtkupFixture otkA, testDate, TEST_VOZ_ID, "I", 100#, cridA, ""
    AppendRF28OtkupFixture otkB, testDate, TEST_VOZ_ID, "I", 100#, cridB, ""
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

' ZBR-CHILD-01 / P1: neprazna generacija NIJE dokaz da akter zna dokument.
'
' Faza 4 popusta kapiju uz ugovor "akter zna identitet". Ali `gen <> ""` znaci
' samo da akter drzi NEKU generaciju -- ne nuzno onu koja pripada prosledjenom
' broju. A `RedJeIzabranogDokumenta` kad dobije generaciju bira red ISKLJUCIVO po
' njoj: broj se tada vise i ne gleda.
'
'   broj X:  GEN-A, GEN-B   (oba aktivna, sva deca nose generaciju)
'   broj Y:  GEN-C
'
'   RunSimpleStornoZbirna("X", "GEN-C")
'     bez provere para -> kapija popusti (gen neprazna, deca scoped)
'                      -> StornoZbirna bira GEN-C, jer broj vise ne ucestvuje
'                      -> stornira se dokument DRUGOG poslovnog broja
'
' Rupa je STARIJA od faze 4 -- i kad X nosi jedan dokument, nespojiv par prolazi.
' Faza 4 je samo uklonila kapiju koja ju je maskirala kad je X dvosmislen.
Private Sub Test_ZBR_TudjaGeneracijaNeOtvaraKapiju()
    Dim tx As clsTransaction
    Dim scenario As String, testDate As Date
    Dim brojX As String, brojY As String
    Dim zbrA As String, zbrB As String, zbrC As String
    Dim genA As String, genB As String, genC As String
    Dim otpA As String, otpC As String
    Dim r As Object

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRPAR")
    testDate = NextTestDate()
    brojX = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")
    brojY = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & _
            Format$(NextTestDate(), "ddmmyy")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP

    ' broj X: dva aktivna dokumenta istog vlasnika
    zbrA = TestHook_ImportZbirnaRowPWA("CRID-ZBRPAR-A-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 100, brojX)
    genA = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrA)
    otpA = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-PARA-" & scenario, brojX, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    zbrB = TestHook_ImportZbirnaRowPWA("CRID-ZBRPAR-B-" & scenario, TEST_VOZ_ID, _
                                       TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 120, brojX)
    genB = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrB)
    AssertTrue (Len(genA) > 0 And Len(genB) > 0 And genA <> genB), _
        "ZBR-PAR preduslov: broj X nosi dva aktivna dokumenta"

    ' broj Y: sasvim drugi dokument
    zbrC = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojY, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         90#, TEST_TIP_AMB, 9, KLASA_I)
    genC = GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbrC)
    otpC = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTP-PARC-" & scenario, brojY, _
                             TEST_VRSTA, TEST_SORTA, 90#, 10#, TEST_TIP_AMB, 9, KLASA_I)
    AssertTrue Len(genC) > 0, "ZBR-PAR preduslov: broj Y nosi svoj dokument"
    AssertEquals genC, DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, otpC), _
        "ZBR-PAR preduslov: otpremnica Y nosi generaciju C"

    AssertFalse ZbirnaGeneracijaPripadaBroju(brojX, genC), _
        "ZBR-PAR preduslov: GEN-C ne pripada broju X"

    ' Nespojiv par: broj X, generacija sa broja Y.
    Set r = RunSimpleStornoZbirna(brojX, genC)
    AssertFalse CBool(r("success")), _
        "ZBR-PAR: nespojiv par (broj, generacija) ne prolazi"

    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrC), _
        "ZBR-PAR: dokument DRUGOG broja ostaje netaknut"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrA), _
        "ZBR-PAR: dokument A ostaje aktivan"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrB), _
        "ZBR-PAR: dokument B ostaje aktivan"
    AssertEquals brojY, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpC, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-PAR: dete drugog broja nije odvezano"
    AssertEquals brojX, _
        NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpA, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-PAR: dete broja X nije odvezano"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-CHILD-01 tudja generacija ne otvara kapiju", bfpErrDesc
End Sub

Private Sub Test_ZBR_MutacijaPoBrojuStajeNaDvaDokumenta()
    Dim tx As clsTransaction
    Dim testDate As Date
    Dim scenario As String
    Dim broj As String, brojDvoklasna As String
    Dim idA As String, idB As String, otpID As String
    Dim ident As ZbirnaIdent
    Dim r As Object
    Dim zbr1 As String, zbr2 As String

    On Error GoTo EH

    scenario = NewScenarioCode("ZBRMUT")
    testDate = NextTestDate()
    broj = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & Format$(testDate, "ddmmyy")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_OTKUP

    idA = TestHook_ImportZbirnaRowPWA("CRID-ZBRMUT-A-" & scenario, TEST_VOZ_ID, _
                                      TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 100, broj)
    idB = TestHook_ImportZbirnaRowPWA("CRID-ZBRMUT-B-" & scenario, TEST_VOZ_ID, _
                                      TEST_KUP_ID, testDate, TEST_VRSTA, TEST_SORTA, 120, broj)

    ' Dete koje visi o BROJU -- ono sto je detach ranije odvezivao preko granice.
    otpID = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                              TEST_PREFIX & "-OTP-ZBRMUT-" & scenario, broj, _
                              TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)

    ident = ZbirnaIdentResolve(broj, TEST_VOZ_ID, TEST_KUP_ID)

    ' Preduslov: bas A17 oblik. Bez ovoga bi test mogao da meri dva VLASNIKA,
    ' sto je zatecena kapija i ranije hvatala.
    AssertEquals "2", CStr(ident.activeLogicalCount), _
        "ZBR-MUT preduslov: broj nosi DVA aktivna dokumenta"
    AssertEquals "1", CStr(ident.activeOwnerCount), _
        "ZBR-MUT preduslov: oba su ISTOG vlasnika"
    AssertTrue Len(otpID) > 0, "ZBR-MUT preduslov: otpremnica visi o tom broju"

    ' --- SIMPLE ---
    Set r = RunSimpleStornoZbirna(broj)
    AssertFalse CBool(r("success")), _
        "ZBR-MUT: SIMPLE storno staje na dva aktivna dokumenta istog vlasnika"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, idA), _
        "ZBR-MUT: dokument A ostaje aktivan"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, idB), _
        "ZBR-MUT: dokument B ostaje aktivan"
    AssertEquals broj, NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-MUT: otpremnica NIJE odvezana preko granice dokumenta"

    ' --- ISPRAVKA i DUPLI: RAZLICITE PUTANJE, obe se mere ---
    '
    ' Do v6-ui-225 je ovaj blok pisao "ISPRAVKA" a vrteo SV_MODE_DUPLI. Tvrdnja
    ' je bila zelena, ali ne iz razloga koji je imenovala: DUPLI staje tek u
    ' StornoZbirnaIDetach_TX, dok ISPRAVKA do te rutine uopste ne dolazi --
    ' ona ide na CreateCorrectionContext pa StornoZbirna_TX, i njena jedina
    ' odbrana je PRED-MUTACIONA kapija u RunZbirnaCorrection. Ta kapija je do
    ' istog koraka brojala VLASNIKE, pa je A17 kroz nju prolazio: zaglavlje bi
    ' bilo stornirano, a blokada stigla tek na CompleteZbirnaIspravka -- dakle
    ' posle izmene, u MANUAL stanju.
    Set r = RunZbirnaCorrection(broj, SV_MODE_ISPRAVKA, True)
    AssertFalse CBool(r("success")), _
        "ZBR-MUT: ISPRAVKA staje PRE mutacije na dva aktivna dokumenta"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, idA), _
        "ZBR-MUT: ISPRAVKA nije stornirala zaglavlje"

    Set r = RunZbirnaCorrection(broj, SV_MODE_DUPLI, True)
    AssertFalse CBool(r("success")), _
        "ZBR-MUT: DUPLI staje na dva aktivna dokumenta istog vlasnika"
    AssertEquals broj, NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE)), _
        "ZBR-MUT: DUPLI nije odvezao otpremnicu"

    ' --- NEGATIVNA KONTROLA: dvoklasna zbirna (dva reda, JEDNA generacija) ---
    ' Kapija sme da odbija samo dva DOKUMENTA. Ako pocne da odbija i ovo, obara
    ' redovan storno svake dvoklasne zbirne -- pa bi tvrdnje gore bile zelene iz
    ' pogresnog razloga.
    brojDvoklasna = CStr(ExtractNumericFromEntityID(TEST_VOZ_ID)) & "/" & _
                    Format$(NextTestDate(), "ddmmyy")
    zbr1 = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojDvoklasna, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    zbr2 = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojDvoklasna, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         50#, TEST_TIP_AMB, 5, KLASA_II)
    AssertEquals GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbr1), _
                 GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, zbr2), _
        "ZBR-MUT preduslov: dva reda dvoklasne dele generaciju"

    Set r = RunSimpleStornoZbirna(brojDvoklasna)
    AssertTrue CBool(r("success")), _
        "ZBR-MUT negativna kontrola: dvoklasna zbirna se i dalje stornira"

    tx.RollbackTx
    Exit Sub

EH:
    ' Err se brise SVAKIM 'On Error' -- opis se hvata PRE rollback-a.
    Dim bfpErrDesc As String: bfpErrDesc = Err.Number & ": " & Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogFail "ZBR-MUT-01 dva dokumenta istog vlasnika", bfpErrDesc
End Sub

Private Sub Test_StornoGuardNaSvimPutanjama()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("STOPUT")

    Dim testDate As Date
    testDate = NextTestDate()

    ' --- OTPREMNICA: isti broj na DVE stanice -> ISPRAVKA i DUPLI moraju pasti ---
    Dim brojOtp As String
    brojOtp = TEST_PREFIX & "-OTP-2ST-" & scenario

    Dim otpA As String, otpB As String
    otpA = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, "", _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    otpB = SaveOtpremnica_TX(testDate, TEST_HLAD_ST_ID, TEST_VOZ_ID, brojOtp, "", _
                             TEST_VRSTA, TEST_SORTA, 80#, 10#, TEST_TIP_AMB, 8, KLASA_I)
    AssertTrue Len(otpA) > 0 And Len(otpB) > 0, _
               "Guard putanje: otpremnice istog broja na dve stanice kreirane"

    Dim rOtp As Object
    Set rOtp = RunOtpremnicaCorrection(brojOtp, SV_MODE_DUPLI, True)
    AssertFalse CBool(rOtp("success")), _
                "Guard putanje: DUPLI otpremnice sa dva vlasnika je odbijen"
    AssertTrue Not RowIsStornirano(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
               "Guard putanje: otpremnica stanice A ostaje aktivna"
    AssertTrue Not RowIsStornirano(TBL_OTPREMNICA, COL_OTP_ID, otpB), _
               "Guard putanje: otpremnica stanice B ostaje aktivna"

    Set rOtp = RunOtpremnicaCorrection(brojOtp, SV_MODE_ISPRAVKA, True)
    AssertTrue Not RowIsStornirano(TBL_OTPREMNICA, COL_OTP_ID, otpA), _
               "Guard putanje: ISPRAVKA ne stornira otpremnicu stanice A"
    AssertTrue Not RowIsStornirano(TBL_OTPREMNICA, COL_OTP_ID, otpB), _
               "Guard putanje: ISPRAVKA ne stornira otpremnicu stanice B"

    ' --- ZBIRNA: isti broj kod dva kupca -> SIMPLE i DUPLI moraju pasti ---
    Dim brojZbr As String
    brojZbr = TEST_PREFIX & "-ZBR-2KUP-" & scenario

    Dim zbrA As String, zbrB As String
    zbrA = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbr, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    zbrB = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbr, TEST_KUP2_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         80#, TEST_TIP_AMB, 8, KLASA_I)
    AssertTrue Len(zbrA) > 0 And Len(zbrB) > 0, _
               "Guard putanje: zbirne istog broja kod dva kupca kreirane"

    Dim rZbr As Object
    Set rZbr = RunSimpleStornoZbirna(brojZbr)
    AssertFalse CBool(rZbr("success")), _
                "Guard putanje: SIMPLE storno zbirne sa dva vlasnika je odbijen"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrA), _
               "Guard putanje: zbirna kupca A ostaje aktivna"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrB), _
               "Guard putanje: zbirna kupca B ostaje aktivna"

    Set rZbr = RunZbirnaCorrection(brojZbr, SV_MODE_DUPLI, True)
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrA), _
               "Guard putanje: DUPLI zbirne ne stornira kupca A"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrB), _
               "Guard putanje: DUPLI zbirne ne stornira kupca B"

    ' --- PRIJEMNICA kroz correction dispatch (ne samo direktan helper) ---
    Dim brojZbrOK As String, brojPrij As String
    brojZbrOK = TEST_PREFIX & "-ZBR-PC-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-2KUP-" & scenario

    AssertTrue Len(SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbrOK, TEST_KUP_ID, _
                                 "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                                 100#, TEST_TIP_AMB, 0, KLASA_I)) > 0, _
               "Guard putanje: fixture zbirna za prijemnice kreirana"

    Dim prjA As String, prjB As String
    prjA = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbrOK, _
                             TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    prjB = SavePrijemnica_TX(testDate, TEST_KUP2_ID, TEST_VOZ_ID, brojPrij, brojZbrOK, _
                             TEST_VRSTA, TEST_SORTA, 80#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    AssertTrue Len(prjA) > 0 And Len(prjB) > 0, _
               "Guard putanje: prijemnice istog broja kod dva kupca kreirane"

    Dim rPrj As Object
    Set rPrj = RunPrijemnicaCorrection(brojPrij, SV_MODE_DUPLI, True)
    AssertTrue Not RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, prjA), _
               "Guard putanje: correction prijemnice ne stornira kupca A"
    AssertTrue Not RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, prjB), _
               "Guard putanje: correction prijemnice ne stornira kupca B"

    Exit Sub

EH:
    LogFatal "Test_StornoGuardNaSvimPutanjama", Err.Number, Err.description
End Sub

' Kaskade (malina / autohladnjaca) mutiraju lanac po BrojZbirne. Ako taj broj nije
' jedinstven, kaskada bi oborila TUDJI lanac -- guard mora vaziti i tu, ne samo na
' direktnim storno putanjama.
Private Sub Test_StornoGuardUKaskadi()
    Dim prevMode As String
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("STOKASK")

    Dim testDate As Date
    testDate = NextTestDate()

    ' Dve zbirne ISTOG broja kod dva kupca (isti vozac) -> broj je dvosmislen.
    Dim brojZbr As String
    brojZbr = TEST_PREFIX & "-ZBR-KASK-" & scenario

    Dim zbrA As String, zbrB As String
    zbrA = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbr, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         100#, TEST_TIP_AMB, 10, KLASA_I)
    zbrB = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbr, TEST_KUP2_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         80#, TEST_TIP_AMB, 8, KLASA_I)
    AssertTrue Len(zbrA) > 0 And Len(zbrB) > 0, _
               "Kaskada guard: dve zbirne istog broja kod dva kupca kreirane"

    ' Otpremnica vezana na taj (dvosmislen) BrojZbirne.
    Dim brojOtp As String
    brojOtp = TEST_PREFIX & "-OTP-KASK-" & scenario

    Dim otpID As String
    otpID = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbr, _
                              TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, KLASA_I)
    AssertTrue Len(otpID) > 0, "Kaskada guard: otpremnica na dvosmislenu zbirnu kreirana"

    ' Malina mod: storno otpremnice kaskadira na njenu zbirnu (StornoZbirnaCascade).
    prevMode = GetConfigValue(CFG_KEY_MALINA_MODE)
    SetConfigValue CFG_KEY_MALINA_MODE, "YES"

    AssertFalse StornoOtpremnicaByBroj_TX(brojOtp), _
                "Kaskada guard: storno otpremnice sa dvosmislenom zbirnom je odbijen"

    SetConfigValue CFG_KEY_MALINA_MODE, prevMode

    ' Rollback: ni otpremnica ni ijedna zbirna nisu dirane.
    AssertTrue Not RowIsStornirano(TBL_OTPREMNICA, COL_OTP_ID, otpID), _
               "Kaskada guard: otpremnica ostaje aktivna (TX rollback)"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrA), _
               "Kaskada guard: zbirna kupca A ostaje aktivna"
    AssertTrue Not RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrB), _
               "Kaskada guard: zbirna kupca B (tudji lanac) ostaje aktivna"

    Exit Sub

EH:
    On Error Resume Next
    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    On Error GoTo 0
    LogFatal "Test_StornoGuardUKaskadi", Err.Number, Err.description
End Sub

' Kaskade mutiraju tblOtpremnica/tblPrijemnica po BrojZbirne, a vlasnik se cita iz
' zbirne -- zato se scope lanca razresava JEDNOM pre prve mutacije i child redovi se
' filtriraju po njemu. Pokriva javni ulaz (StornoOtkupByBrDok_TX), sve tri kaskade,
' single-owner happy path i fail-closed granu bez aktivnog parenta.
Private Sub Test_StornoKaskadaScopePoLancu()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim scenario As String
    scenario = NewScenarioCode("KASKSCOPE")

    ' --- Deo 1: happy path + TUDJI aktivan child pod istim BrojZbirne ---
    Dim brDok As String
    brDok = TEST_PREFIX & "-KSC-" & scenario

    Dim brPrij As String, w As String
    w = RunHladnjacaChain(brDok, NextTestDate(), "", brPrij)
    AssertEquals "", w, "Kaskada scope: hladnjaca lanac kreiran bez upozorenja"

    Dim otpI As String, zbrI As String, prjI As String
    otpI = FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I)
    zbrI = FindZbirnaIDByBrojAndKlasa(brDok, KLASA_I)
    prjI = FindPrijemnicaIDByBrojAndKlasa(brPrij, KLASA_I)
    AssertTrue Len(otpI) > 0 And Len(zbrI) > 0 And Len(prjI) > 0, _
               "Kaskada scope: otpremnica/zbirna/prijemnica lanca postoje"

    ' Tudja prijemnica DRUGOG kupca vezana na ISTI BrojZbirne (co-tenant / osirocena).
    Dim tudjaPrij As String
    tudjaPrij = SavePrijemnica_TX(NextTestDate(), TEST_KUP2_ID, TEST_VOZ_ID, _
                                  TEST_PREFIX & "-KSC-TUDJA-" & scenario, brDok, _
                                  TEST_VRSTA, TEST_SORTA, 60#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    AssertTrue Len(tudjaPrij) > 0, "Kaskada scope: tudja prijemnica na isti BrojZbirne kreirana"

    AssertTrue StornoOtkupByBrDok_TX(brDok), _
               "Kaskada scope: storno otkup bloka (single owner) prolazi"

    AssertTrue RowIsStornirano(TBL_OTPREMNICA, COL_OTP_ID, otpI), _
               "Kaskada scope: otpremnica lanca stornirana"
    AssertTrue RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrI), _
               "Kaskada scope: zbirna lanca stornirana"
    AssertTrue RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, prjI), _
               "Kaskada scope: prijemnica lanca stornirana (kaskada radi)"
    AssertTrue Not RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, tudjaPrij), _
               "Kaskada scope: prijemnica DRUGOG kupca pod istim BrojZbirne NETAKNUTA"

    ' --- Deo 2: zbirna stornirana, njena prijemnica JOS AKTIVNA -> fail-closed ---
    ' (Prijemnica se ne moze kreirati bez zbirne -- PRIJEMNICA_ZBIRNA_PROVERA -- pa
    '  se osiroceno stanje pravi legitimno: zbirna, pa prijemnica, pa storno zbirne.)
    Dim brDok2 As String
    brDok2 = TEST_PREFIX & "-KSC2-" & scenario

    Dim testDate2 As Date
    testDate2 = NextTestDate()

    Dim zbrB As String
    zbrB = SaveZbirna_TX(testDate2, TEST_VOZ_ID, brDok2, TEST_KUP2_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         50#, TEST_TIP_AMB, 0, KLASA_I)
    AssertTrue Len(zbrB) > 0, "Kaskada scope: zbirna drugog kupca kreirana"

    Dim orphanPrij As String
    orphanPrij = SavePrijemnica_TX(testDate2, TEST_KUP2_ID, TEST_VOZ_ID, _
                                   TEST_PREFIX & "-KSC2-PRJ-" & scenario, brDok2, _
                                   TEST_VRSTA, TEST_SORTA, 50#, 100#, TEST_TIP_AMB, 0, 0, KLASA_I)
    AssertTrue Len(orphanPrij) > 0, "Kaskada scope: prijemnica drugog kupca kreirana"

    ' Zbirna se stornira, prijemnica ostaje aktivna -> osiroceni nizvodni dokument.
    MarkTestRowStornirano TBL_ZBIRNA, "ZbirnaID", zbrB
    AssertTrue RowIsStornirano(TBL_ZBIRNA, COL_ZBR_ID, zbrB), _
               "Kaskada scope: zbirna stornirana, prijemnica ostala aktivna"

    ' Otkup blok na hladnjaca stanici sa istim BrojZbirne (nema aktivne zbirne).
    Dim otkIDs As String
    otkIDs = SaveOtkupMulti_TX(testDate2, TEST_KOOP_ID, TEST_HLAD_ST_ID, TEST_VRSTA, TEST_SORTA, _
                               100#, 100#, TEST_TIP_AMB, 10, TEST_VOZ_ID, brDok2, _
                               0#, "TEST OPERATOR", GetTestParcelaID(), brDok2)
    AssertTrue Len(otkIDs) > 0, "Kaskada scope: otkup blok bez aktivne zbirne kreiran"

    Dim otkID As String
    otkID = FindOtkupIDByBrojAndKlasa(brDok2, KLASA_I)

    AssertFalse StornoOtkupByBrDok_TX(brDok2), _
                "Kaskada scope: bez aktivne zbirne uz aktivan child -> storno je ODBIJEN"
    AssertTrue Len(orphanPrij) > 0 And Not RowIsStornirano(TBL_PRIJEMNICA, COL_PRJ_ID, orphanPrij), _
               "Kaskada scope: osirocena prijemnica ostaje netaknuta"
    If Len(otkID) > 0 Then
        AssertTrue Not RowIsStornirano(TBL_OTKUP, "OtkupID", otkID), _
                   "Kaskada scope: otkup red ostaje aktivan (TX rollback)"
    End If

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub

EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFatal "Test_StornoKaskadaScopePoLancu", Err.Number, Err.description
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

' End-to-end: save putanja stvarno pise GeneracijaID, i to ISTU za obe klase
' jednog Multi_TX upisa, a NOVU za ispravku istog broja.
Private Sub Test_GeneracijaIDNaSavePutanji()
    On Error GoTo EH

    ' Kolona je obavezan invariant (EnsureSledljivostSchema je pravi na svakom
    ' startu) -> nedostatak je FAIL, ne SKIP; inace suite ostaje zelen bez pokrica.
    AssertTrue GetColumnIndex(TBL_OTPREMNICA, COL_GENERACIJA_ID) > 0, _
               "GeneracijaID: kolona postoji na tblOtpremnica"
    AssertTrue GetColumnIndex(TBL_ZBIRNA, COL_GENERACIJA_ID) > 0, _
               "GeneracijaID: kolona postoji na tblZbirna"
    AssertTrue GetColumnIndex(TBL_PRIJEMNICA, COL_GENERACIJA_ID) > 0, _
               "GeneracijaID: kolona postoji na tblPrijemnica"

    Dim scenario As String
    scenario = NewScenarioCode("GENID")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojOtp As String, brojZbirne As String
    brojOtp = TEST_PREFIX & "-OTP-GEN-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-GEN-" & scenario

    ' Generacija 1: dvoklasna otpremnica (jedan Multi_TX poziv).
    Dim res1 As String
    res1 = SaveOtpremnicaMulti_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                                  TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 10, _
                                  True, 50#, 8#)
    AssertTrue Len(res1) > 0, "GeneracijaID: dvoklasna otpremnica snimljena"

    Dim genI As String, genII As String
    genI = OtpGeneracija(brojOtp, KLASA_I)
    genII = OtpGeneracija(brojOtp, KLASA_II)

    AssertTrue Len(genI) > 0, "GeneracijaID: Klasa I ima generaciju"
    AssertEquals genI, genII, "GeneracijaID: obe klase jednog upisa dele generaciju"

    ' Generacija 2: ispravka istog broja, samo Klasa I.
    MarkTestRowStornirano TBL_OTPREMNICA, "OtpremnicaID", FindOtpremnicaIDByBrojAndKlasa(brojOtp, KLASA_I)
    MarkTestRowStornirano TBL_OTPREMNICA, "OtpremnicaID", FindOtpremnicaIDByBrojAndKlasa(brojOtp, KLASA_II)

    Dim res2 As String
    res2 = SaveOtpremnicaMulti_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                                  TEST_VRSTA, TEST_SORTA, 120#, 10#, TEST_TIP_AMB, 12)
    AssertTrue Len(res2) > 0, "GeneracijaID: ispravka (samo Kl.I) snimljena"

    ' Nasledjivanje ide samo od AKTIVNIH redova -> posle storna nema sta da se
    ' nasledi i ispravka dobija NOVU generaciju.
    AssertTrue OtpGeneracija(brojOtp, KLASA_I) <> genI, _
               "GeneracijaID: ispravka posle storna dobija NOVU generaciju"

    ' Prefill nad REALNOM tabelom mora dati novi Kl.I red i praznu Kl.II.
    Dim d As Variant
    d = GetTableData(TBL_OTPREMNICA)

    Dim cBr As Long, cKl As Long, cId As Long, cGen As Long, cKol As Long
    cBr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    cKl = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA)
    cId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    cGen = GetColumnIndex(TBL_OTPREMNICA, COL_GENERACIJA_ID)
    cKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)

    Dim rI As Long, rII As Long
    PickPrefillRows d, cBr, cKl, cId, cGen, brojOtp, _
                    FindOtpremnicaIDByBrojAndKlasa(brojOtp, KLASA_I), rI, rII

    AssertTrue rI > 0, "GeneracijaID: prefill nasao Kl.I poslednje generacije"
    AssertEquals "0", CStr(rII), _
                 "GeneracijaID: stara Kl.II se NE prefiluje uz novu Kl.I"

    If rI > 0 And cKol > 0 Then
        Dim kolI As Double
        kolI = CDbl(nz(d(rI, cKol), 0))
        AssertTrue Abs(kolI - 120#) < 0.001, _
                   "GeneracijaID: prefill uzima kolicinu IZ ISPRAVKE (120), ne original"
    End If

    Exit Sub

EH:
    LogFatal "Test_GeneracijaIDNaSavePutanji", Err.Number, Err.description
End Sub

' Generacija reda po ID-u. Prazan ID daje "" -> AssertEquals nad dva prazna bi
' lazno prosao, pa pozivaoci uz poredjenje tvrde i da generacija NIJE prazna.
Private Function DokGeneracija(ByVal tableName As String, ByVal idColumn As String, _
                               ByVal idValue As String) As String
    If Len(Trim$(idValue)) = 0 Then Exit Function

    DokGeneracija = Trim$(CStr(nz(GetValueByKey(tableName, idColumn, idValue, _
                                                COL_GENERACIJA_ID), "")))
End Function

Private Function OtpGeneracija(ByVal brojOtp As String, ByVal klasa As String) As String
    OtpGeneracija = DokGeneracija(TBL_OTPREMNICA, COL_OTP_ID, _
                                  FindOtpremnicaIDByBrojAndKlasa(brojOtp, klasa))
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

    Dim raised As Boolean
    On Error Resume Next
    Call AutoCreateZbirnaFromOtpremnice_TX(brojOtpNema)
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
    created = AutoCreateZbirnaFromOtpremnice_TX(brojOtpNema)
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

' MIG-001: EKRAN -> WRITER -> TABELA za hladnjacu i pogon.
'
' Test_ZbirnaRowDataColumnMapped iznad meri drugu kariku: writer prima dve
' vrednosti kao argumente i pise ih u SVOJE kolone. Ta tvrdnja je bila zelena i
' dok je ekran slao prazno -- writer je dobijao "" i uredno ga upisivao.
'
' Ovde se meri put kojim vrednost STVARNO ide u produkciji: recnik ljuske ->
' modScrDokumenti.Scr_Save -> modDokUnos.ZbirnaUpisi -> SaveZbirnaMulti_TX ->
' tblZbirna. Pukne cim bilo koja karika ispusti kljuc (npr. mapiranje u
' SaveZbirna), sto tvrdnja nad writerom ne vidi.
Private Sub Test_ZbirnaEkranNosiOdrediste()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("ZBREKR")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojOtp As String, brojZbirne As String
    brojOtp = TEST_PREFIX & "-OTP-EKR-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-EKR-" & scenario

    ' Izvor zbirne: jedna otpremnica jedne klase. Zbirna mora da prijavi TACNO
    ' njene kilograme i gajbe, inace je zaustavi ZbirnaValidiraj i test bi merio
    ' kapiju umesto prenosa vrednosti.
    Dim otpRes As String
    otpRes = SaveOtpremnicaMulti_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                                    TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 4)
    AssertTrue Len(otpRes) > 0, "Zbirna ekran: izvorna otpremnica snimljena"

    ' Recnik je isti oblik koji ljuska salje (modOtkupUI.SkupiPolja): partner je
    ' pod kljucem "kooperantID" jer je to ista kontrola u svim rezimima.
    Dim polja As Object
    Set polja = CreateObject("Scripting.Dictionary")
    polja.CompareMode = vbTextCompare
    polja("rezim") = "ZBIRNA"
    polja("datum") = testDate
    polja("vozacID") = TEST_VOZ_ID
    polja("kooperantID") = TEST_KUP_ID
    polja("brDok") = brojZbirne
    polja("hladnjaca") = "Hladnjaca " & scenario
    polja("pogon") = "Pogon " & scenario
    polja("vrsta") = TEST_VRSTA
    polja("sorta") = TEST_SORTA
    polja("tipAmb") = TEST_TIP_AMB
    polja("kolicinaI") = 100#
    polja("kolAmb") = 4
    polja("dveKlase") = False
    polja("kolicinaII") = 0#
    polja("kolAmbII") = 0

    Dim greska As String
    greska = modScrDokumenti.Scr_Save(polja)
    AssertEquals "", greska, "Zbirna ekran: Scr_Save prolazi"

    Dim zbrID As String
    zbrID = Trim$(CStr(polja("rezultat")))
    AssertTrue Len(zbrID) > 0, "Zbirna ekran: upis vraca ZbirnaID"

    ' Tvrdnja koja nosi ceo test: ono sto je operater izabrao stiglo je u red.
    AssertEquals "Hladnjaca " & scenario, ZbrPolje(zbrID, COL_ZBR_HLADNJACA), _
                 "Zbirna ekran: hladnjaca iz recnika je u tblZbirna"
    AssertEquals "Pogon " & scenario, ZbrPolje(zbrID, COL_ZBR_POGON), _
                 "Zbirna ekran: pogon iz recnika je u tblZbirna"

    ' Kontrola u drugom smeru: bez ta dva kljuca upis i dalje prolazi i kolone
    ' ostaju prazne. Prazno je legitimno stanje (kupac bez hladnjace), pa novo
    ' polje ne sme da postane kapija.
    Dim brojZbirne2 As String, otpRes2 As String, zbrID2 As String
    brojZbirne2 = TEST_PREFIX & "-ZBR-EKR2-" & scenario
    otpRes2 = SaveOtpremnicaMulti_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                                     TEST_PREFIX & "-OTP-EKR2-" & scenario, brojZbirne2, _
                                     TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 4)
    AssertTrue Len(otpRes2) > 0, "Zbirna ekran: druga izvorna otpremnica snimljena"

    polja("brDok") = brojZbirne2
    polja("hladnjaca") = ""
    polja("pogon") = ""
    greska = modScrDokumenti.Scr_Save(polja)
    AssertEquals "", greska, "Zbirna ekran: prazno odrediste ne blokira upis"
    zbrID2 = Trim$(CStr(polja("rezultat")))
    AssertTrue Len(zbrID2) > 0, "Zbirna ekran: drugi upis vraca ZbirnaID"
    AssertEquals "", ZbrPolje(zbrID2, COL_ZBR_HLADNJACA), _
                 "Zbirna ekran: prazna hladnjaca ostaje prazna"

    Exit Sub

EH:
    LogFatal "Test_ZbirnaEkranNosiOdrediste", Err.Number, Err.description
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

Private Function FindOtkupIDByBrojAndKlasa(ByVal brojDok As String, ByVal klasa As String) As String
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    Dim colBroj As Long
    Dim colKlasa As Long

    colID = RequireCol(TBL_OTKUP, "OtkupID")
    colBroj = RequireCol(TBL_OTKUP, "BrojDokumenta")
    colKlasa = RequireCol(TBL_OTKUP, "Klasa")

    Dim i As Long
    For i = UBound(data, 1) To 1 Step -1
        If CStr(data(i, colBroj)) = brojDok _
           And CStr(data(i, colKlasa)) = klasa Then
            FindOtkupIDByBrojAndKlasa = CStr(data(i, colID))
            Exit Function
        End If
    Next i

    Exit Function

EH:
    FindOtkupIDByBrojAndKlasa = ""
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

Private Function FindOtpremnicaIDByBrojAndKlasa(ByVal brojOtp As String, _
                                                ByVal klasa As String) As String
    FindOtpremnicaIDByBrojAndKlasa = FindIDByTwoColumns( _
        TBL_OTPREMNICA, "OtpremnicaID", "BrojOtpremnice", brojOtp, "Klasa", klasa)
End Function

Private Function FindZbirnaIDByBrojAndKlasa(ByVal brojZbirne As String, _
                                            ByVal klasa As String) As String
    FindZbirnaIDByBrojAndKlasa = FindIDByTwoColumns( _
        TBL_ZBIRNA, "ZbirnaID", "BrojZbirne", brojZbirne, "Klasa", klasa)
End Function

Private Function FindPrijemnicaIDByBrojAndKlasa(ByVal brojPrij As String, _
                                                ByVal klasa As String) As String
    FindPrijemnicaIDByBrojAndKlasa = FindIDByTwoColumns( _
        TBL_PRIJEMNICA, "PrijemnicaID", "BrojPrijemnice", brojPrij, "Klasa", klasa)
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

' ============================================================
' HLADNJACA AUTO-LANAC (modAutoHladnjaca) -- RF-04
'
' Pokriva: fail-fast nizvodno (pad koraka NE sme da ostavi nizvodne dokumente),
' outBrPrij tek posle STVARNO kreirane prijemnice, propagaciju pada back-linka,
' i backfill (deljen broj po BrojZbirne + izolacija mapa po kupcu).
'
' Pad pojedinacnog koraka se izaziva test seam-om ArmHladnjacaTestFail
' (modAutoHladnjaca). Seam je jednokratan -- AutoChainHladnjaca ga trosi na ulazu.
' ============================================================
Private Sub ArrangeHladnjacaConfig(ByRef prevAuto As String, ByRef prevKupac As String)
    prevAuto = GetConfigValue(CFG_AUTO_PRIJEMNICA_HLADNJACA)
    prevKupac = GetConfigValue(CFG_MALINA_DEFAULT_KUPAC)
    SetConfigValue CFG_AUTO_PRIJEMNICA_HLADNJACA, "YES"
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, TEST_KUP_ID
End Sub

Private Sub RestoreHladnjacaConfig(ByVal prevAuto As String, ByVal prevKupac As String)
    On Error Resume Next
    SetConfigValue CFG_AUTO_PRIJEMNICA_HLADNJACA, prevAuto
    SetConfigValue CFG_MALINA_DEFAULT_KUPAC, prevKupac
    ArmHladnjacaTestFail ""      ' seam ne sme da ostane armiran ni posle pada testa
    On Error GoTo 0
End Sub

' Otkup (Klasa I + II) na hladnjaca stanici -> pa auto-lanac. Vraca upozorenje
' lanca; outBrPrij nosi broj prijemnice (prazan ako nijedna nije kreirana).
Private Function RunHladnjacaChain(ByVal brDok As String, ByVal testDate As Date, _
                                   ByVal failStep As String, _
                                   ByRef outBrPrij As String) As String
    Dim otkupIDs As String
    otkupIDs = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_HLAD_ST_ID, TEST_VRSTA, TEST_SORTA, _
                                 100#, 100#, TEST_TIP_AMB, 10, TEST_VOZ_ID, brDok, _
                                 0#, "TEST OPERATOR", GetTestParcelaID(), brDok, _
                                 True, 50#, 80#, 0, 0#, 5, 0#)

    If Len(failStep) > 0 Then ArmHladnjacaTestFail failStep

    RunHladnjacaChain = AutoChainHladnjaca(testDate, TEST_HLAD_ST_ID, TEST_VRSTA, TEST_SORTA, _
                                           TEST_VOZ_ID, TEST_TIP_AMB, 10, 100#, 100#, _
                                           True, 50#, 80#, brDok, otkupIDs, _
                                           0#, 5, 0#, outBrPrij)
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

' Kontrolna grupa: bez simulacije pada ceo lanac mora da prodje.
Private Sub Test_HladnjacaChainHappyPath()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim brDok As String
    brDok = TEST_PREFIX & "-HLOK-" & NewScenarioCode("HLADOK")

    Dim brPrij As String, w As String
    w = RunHladnjacaChain(brDok, NextTestDate(), "", brPrij)

    AssertEquals "", w, "Hladnjaca lanac: kompletan lanac ne vraca upozorenje"
    AssertTrue Len(brPrij) > 0, "Hladnjaca lanac: outBrPrij izlozen posle kreirane prijemnice"

    AssertTrue Len(FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I)) > 0, _
        "Hladnjaca lanac: otpremnica Klasa I kreirana"
    AssertTrue Len(FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_II)) > 0, _
        "Hladnjaca lanac: otpremnica Klasa II kreirana"
    AssertTrue Len(FindZbirnaIDByBrojAndKlasa(brDok, KLASA_I)) > 0, _
        "Hladnjaca lanac: zbirna Klasa I kreirana"
    AssertTrue Len(FindZbirnaIDByBrojAndKlasa(brDok, KLASA_II)) > 0, _
        "Hladnjaca lanac: zbirna Klasa II kreirana"

    ' Jedna prijemnica = jedan broj: obe klase nose isti BrojPrijemnice.
    AssertEquals brPrij, FindPrijBrojByZbirnaKlasaKupac(brDok, KLASA_I, TEST_KUP_ID), _
        "Hladnjaca lanac: prijemnica Klasa I nosi izlozeni broj"
    AssertEquals brPrij, FindPrijBrojByZbirnaKlasaKupac(brDok, KLASA_II, TEST_KUP_ID), _
        "Hladnjaca lanac: prijemnica Klasa II nosi ISTI broj"

    ' Generacija: lanac pise Klasu I i II ZASEBNIM _TX pozivima, ali obe klase
    ' istog dokumenta moraju deliti generaciju (inace prefill vidi samo jednu).
    AssertEquals DokGeneracija(TBL_OTPREMNICA, COL_OTP_ID, FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I)), _
                 DokGeneracija(TBL_OTPREMNICA, COL_OTP_ID, FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_II)), _
        "Hladnjaca lanac: otpremnica Kl.I i Kl.II dele generaciju"
    AssertTrue Len(DokGeneracija(TBL_OTPREMNICA, COL_OTP_ID, FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I))) > 0, _
        "Hladnjaca lanac: otpremnica ima generaciju"
    AssertEquals DokGeneracija(TBL_ZBIRNA, COL_ZBR_ID, FindZbirnaIDByBrojAndKlasa(brDok, KLASA_I)), _
                 DokGeneracija(TBL_ZBIRNA, COL_ZBR_ID, FindZbirnaIDByBrojAndKlasa(brDok, KLASA_II)), _
        "Hladnjaca lanac: zbirna Kl.I i Kl.II dele generaciju"
    AssertEquals DokGeneracija(TBL_PRIJEMNICA, COL_PRJ_ID, FindPrijemnicaIDByBrojAndKlasa(brPrij, KLASA_I)), _
                 DokGeneracija(TBL_PRIJEMNICA, COL_PRJ_ID, FindPrijemnicaIDByBrojAndKlasa(brPrij, KLASA_II)), _
        "Hladnjaca lanac: prijemnica Kl.I i Kl.II dele generaciju"

    ' ZBR-CHILD-01: lanac snima otpremnicu PRE zbirne, pa joj je veza u tom
    ' trenutku prazna; ZavrsiVezuOtpremniceNaZbirnu je dovrsava posle. Taj helper
    ' je fail-soft (tri Exit Sub-a i LogErr) i njegov neuspeh NE ulazi u failLink,
    ' pa lanac moze da prijavi uspeh a veza da ostane nerazresena. Merenje je
    ' jedini nacin da se to vidi -- odsustvo upozorenja ovde ne dokazuje nista.
    AssertEquals DokGeneracija(TBL_ZBIRNA, COL_ZBR_ID, FindZbirnaIDByBrojAndKlasa(brDok, KLASA_I)), _
                 DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I)), _
        "Hladnjaca lanac: otpremnica Kl.I nosi generaciju SVOJE zbirne"
    AssertEquals DokGeneracija(TBL_ZBIRNA, COL_ZBR_ID, FindZbirnaIDByBrojAndKlasa(brDok, KLASA_II)), _
                 DeteGeneracija(TBL_OTPREMNICA, COL_OTP_ID, FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_II)), _
        "Hladnjaca lanac: otpremnica Kl.II nosi generaciju SVOJE zbirne"

    ' Back-link u otkup red.
    Dim otkID As String: otkID = FindOtkupIDByBrojAndKlasa(brDok, KLASA_I)
    AssertTrue Len(CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkID, "OtpremnicaID"))) > 0, _
        "Hladnjaca lanac: otkup red povezan sa otpremnicom"
    AssertEquals brDok, CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkID, "BrojZbirne")), _
        "Hladnjaca lanac: otkup red nosi BrojZbirne"

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub
EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFail "Hladnjaca chain happy path", Err.description
End Sub

' P1: pad OTPREMNICE mora da zaustavi lanac -- bez zbirne i bez prijemnice.
Private Sub Test_HladnjacaChainFailFastOtpremnica()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim brDok As String
    brDok = TEST_PREFIX & "-HLFO-" & NewScenarioCode("HLADFO")

    Dim brPrij As String, w As String
    w = RunHladnjacaChain(brDok, NextTestDate(), "OTP", brPrij)

    AssertTrue InStr(w, "OTPREMNICA nije kreirana") > 0, _
        "Fail-fast OTP: upozorenje prijavljuje pad otpremnice"
    AssertEquals "", FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I), _
        "Fail-fast OTP: otpremnica nije kreirana"
    AssertEquals "", FindZbirnaIDByBrojAndKlasa(brDok, KLASA_I), _
        "Fail-fast OTP: ZBIRNA nije kreirana (lanac zaustavljen)"
    AssertEquals "", FindPrijBrojByZbirnaKlasaKupac(brDok, KLASA_I, TEST_KUP_ID), _
        "Fail-fast OTP: PRIJEMNICA nije kreirana (lanac zaustavljen)"
    AssertEquals "", brPrij, _
        "Fail-fast OTP: outBrPrij ostaje prazan"

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub
EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFail "Hladnjaca fail-fast otpremnica", Err.description
End Sub

' P1: pad ZBIRNE mora da zaustavi lanac -- otpremnica ostaje, prijemnice nema.
Private Sub Test_HladnjacaChainFailFastZbirna()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim brDok As String
    brDok = TEST_PREFIX & "-HLFZ-" & NewScenarioCode("HLADFZ")

    Dim brPrij As String, w As String
    w = RunHladnjacaChain(brDok, NextTestDate(), "ZBR", brPrij)

    AssertTrue Len(FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I)) > 0, _
        "Fail-fast ZBR: otpremnica (uzvodni korak) jeste kreirana"
    AssertTrue InStr(w, "ZBIRNA nije kreirana") > 0, _
        "Fail-fast ZBR: upozorenje prijavljuje pad zbirne"
    AssertEquals "", FindPrijBrojByZbirnaKlasaKupac(brDok, KLASA_I, TEST_KUP_ID), _
        "Fail-fast ZBR: PRIJEMNICA nije kreirana (lanac zaustavljen)"
    ' Prijemnica nije ni pokusana -> ne sme se pojaviti u upozorenju.
    AssertTrue InStr(w, "PRIJEMNICA nije kreirana") = 0, _
        "Fail-fast ZBR: upozorenje ne prijavljuje korak koji nije ni pokusan"
    AssertEquals "", brPrij, "Fail-fast ZBR: outBrPrij ostaje prazan"

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub
EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFail "Hladnjaca fail-fast zbirna", Err.description
End Sub

' Fix #2: outBrPrij se NE sme izloziti ako prijemnica nije kreirana (caller bi
' relinkovao osirocene palete na nepostojecu prijemnicu).
Private Sub Test_HladnjacaChainPrijemnicaFailNoBroj()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim brDok As String
    brDok = TEST_PREFIX & "-HLFP-" & NewScenarioCode("HLADFP")

    Dim brPrij As String, w As String
    w = RunHladnjacaChain(brDok, NextTestDate(), "PRJ", brPrij)

    AssertTrue Len(FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I)) > 0, _
        "Pad prijemnice: otpremnica jeste kreirana"
    AssertTrue Len(FindZbirnaIDByBrojAndKlasa(brDok, KLASA_I)) > 0, _
        "Pad prijemnice: zbirna jeste kreirana"
    AssertTrue InStr(w, "PRIJEMNICA nije kreirana") > 0, _
        "Pad prijemnice: upozorenje prijavljuje pad prijemnice"
    AssertEquals "", brPrij, _
        "Pad prijemnice: outBrPrij ostaje prazan (nema relinka na nepostojecu)"

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub
EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFail "Hladnjaca prijemnica fail", Err.description
End Sub

' Fix #5: pad back-linka se prijavljuje (ranije je lanac javljao uspeh).
Private Sub Test_HladnjacaChainLinkFailureIsReported()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim brDok As String
    brDok = TEST_PREFIX & "-HLFL-" & NewScenarioCode("HLADFL")

    Dim brPrij As String, w As String
    w = RunHladnjacaChain(brDok, NextTestDate(), "LINK", brPrij)

    AssertTrue Len(FindOtpremnicaIDByBrojAndKlasa(brDok, KLASA_I)) > 0, _
        "Pad linka: dokumenti su kreirani (link je poslednji korak)"
    AssertTrue Len(brPrij) > 0, _
        "Pad linka: prijemnica jeste kreirana pa je outBrPrij izlozen"
    AssertTrue InStr(w, "nije povezan sa dokumentom") > 0, _
        "Pad linka: upozorenje prijavljuje nepovezan otkup red"

    Dim otkID As String: otkID = FindOtkupIDByBrojAndKlasa(brDok, KLASA_I)
    AssertEquals "", CStr(GetValueByKey(TBL_OTKUP, "OtkupID", otkID, "OtpremnicaID")), _
        "Pad linka: otkup red stvarno NIJE povezan"

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub
EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFail "Hladnjaca link failure reported", Err.description
End Sub

' Fix #4: obe klase istog dokumenta dele broj i kad je sestrinska klasa vec
' backfill-ovana u ranijem prolazu.
Private Sub Test_BackfillHladnjacaDeliBrojPoZbirnoj()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim testDate As Date: testDate = NextTestDate()
    Dim brZbr As String
    brZbr = TEST_PREFIX & "-HLBF-" & NewScenarioCode("HLADBF")

    ' Otpremnice obe klase na hladnjaca stanici + zbirne (prijemnica ih zahteva).
    AssertTrue Len(SaveOtpremnica_TX(testDate, TEST_HLAD_ST_ID, TEST_VOZ_ID, brZbr, brZbr, _
        TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 10, KLASA_I)) > 0, _
        "Backfill fixture: otpremnica Klasa I"
    AssertTrue Len(SaveOtpremnica_TX(testDate, TEST_HLAD_ST_ID, TEST_VOZ_ID, brZbr, brZbr, _
        TEST_VRSTA, TEST_SORTA, 50#, 80#, TEST_TIP_AMB, 5, KLASA_II)) > 0, _
        "Backfill fixture: otpremnica Klasa II"
    SaveZbirna_TX testDate, TEST_VOZ_ID, brZbr, TEST_KUP_ID, "Test Hladnjaca", "", _
        TEST_VRSTA, TEST_SORTA, 100#, TEST_TIP_AMB, 10, KLASA_I
    SaveZbirna_TX testDate, TEST_VOZ_ID, brZbr, TEST_KUP_ID, "Test Hladnjaca", "", _
        TEST_VRSTA, TEST_SORTA, 50#, TEST_TIP_AMB, 5, KLASA_II

    ' Klasa I VEC ima prijemnicu; Klasa II je nema.
    Dim brPostojeci As String
    brPostojeci = GenerateBrojPrijemnice(TEST_KUP_ID, testDate)
    AssertTrue Len(SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brPostojeci, brZbr, _
        TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 10, 0, KLASA_I)) > 0, _
        "Backfill fixture: prijemnica Klasa I postoji"

    Dim ok As Long, fail As Long
    ' Opseg = SAMO ovaj dokument: bez toga backfill skenira sve hladnjaca-otpremnice
    ' u svesci, pa bi suite nad realnim fajlom dirao prave dokumente.
    BackfillPrijemniceHladnjacaCore True, brZbr, ok, fail

    AssertEquals brPostojeci, FindPrijBrojByZbirnaKlasaKupac(brZbr, KLASA_II, TEST_KUP_ID), _
        "Backfill: Klasa II nasledjuje broj prijemnice Klase I (isti dokument)"

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub
EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFail "Backfill hladnjaca deli broj po zbirnoj", Err.description
End Sub

' P2b: prijemnica DRUGOG kupca sa istim BrojZbirne ne sme ni da preskoci kandidata
' (idempotentnost) ni da mu pozajmi broj (numeracija je per-kupac).
Private Sub Test_BackfillHladnjacaIgnorisePrijemniceDrugogKupca()
    Dim prevAuto As String, prevKupac As String
    On Error GoTo EH
    ArrangeHladnjacaConfig prevAuto, prevKupac

    Dim testDate As Date: testDate = NextTestDate()
    Dim brZbr As String
    brZbr = TEST_PREFIX & "-HLBF2-" & NewScenarioCode("HLADBF2")

    AssertTrue Len(SaveOtpremnica_TX(testDate, TEST_HLAD_ST_ID, TEST_VOZ_ID, brZbr, brZbr, _
        TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 10, KLASA_I)) > 0, _
        "Backfill izolacija: otpremnica Klasa I"
    SaveZbirna_TX testDate, TEST_VOZ_ID, brZbr, TEST_KUP_ID, "Test Hladnjaca", "", _
        TEST_VRSTA, TEST_SORTA, 100#, TEST_TIP_AMB, 10, KLASA_I

    ' Prijemnica DRUGOG kupca na ISTOM BrojZbirne i istoj klasi.
    Dim brTudji As String
    brTudji = TEST_PREFIX & "-TUDJI-" & NewScenarioCode("HLADTUD")
    AssertTrue Len(SavePrijemnica_TX(testDate, TEST_KUP2_ID, TEST_VOZ_ID, brTudji, brZbr, _
        TEST_VRSTA, TEST_SORTA, 100#, 100#, TEST_TIP_AMB, 10, 0, KLASA_I)) > 0, _
        "Backfill izolacija: prijemnica drugog kupca kreirana"

    Dim ok As Long, fail As Long
    ' Opseg = SAMO ovaj dokument: bez toga backfill skenira sve hladnjaca-otpremnice
    ' u svesci, pa bi suite nad realnim fajlom dirao prave dokumente.
    BackfillPrijemniceHladnjacaCore True, brZbr, ok, fail

    Dim brNas As String
    brNas = FindPrijBrojByZbirnaKlasaKupac(brZbr, KLASA_I, TEST_KUP_ID)
    AssertTrue Len(brNas) > 0, _
        "Backfill izolacija: kandidat NIJE preskocen zbog prijemnice drugog kupca"
    AssertTrue brNas <> brTudji, _
        "Backfill izolacija: broj NIJE pozajmljen iz prijemnice drugog kupca"

    RestoreHladnjacaConfig prevAuto, prevKupac
    Exit Sub
EH:
    RestoreHladnjacaConfig prevAuto, prevKupac
    LogFail "Backfill hladnjaca izolacija po kupcu", Err.description
End Sub

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
                      TEST_VRSTA & " DRUGA")

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

    Exit Sub

EH:
    LogFatal "Test_PR3_AmbalazaMoraBitiCeoBroj", Err.Number, Err.description
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
' Sta se NE meri: ponasanje citalaca, ambalaza i novac -- to je Otkup cutover.
'               Do tada je stari writer (SaveOtkupMulti_TX) jedini put, a golden
'               scenariji to dokazuju nepromenjeni.

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
' Kolone JOS postoje u tabeli -- stari writer ih puni i brisu se tek u cutover-u.
' Zato se ovde meri da ih NOV writer ostavlja prazne. Tvrdnja "kolone nema"
' postaje moguca tek posle cutover-a.
Private Sub Test_OTK_HeaderNeNosiLinePolja()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKHP")

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-HP-" & scenario), _
                           OtkStavke(400#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK header: dokument napravljen"

    AssertEquals "", OtkPolje(otkID, COL_OTK_KOLICINA), "OTK header: Kolicina prazna"
    AssertEquals "", OtkPolje(otkID, COL_OTK_CENA), "OTK header: Cena prazna"
    AssertEquals "", OtkPolje(otkID, COL_OTK_KLASA), "OTK header: Klasa prazna"
    AssertEquals "", OtkPolje(otkID, COL_OTK_KOL_AMB), "OTK header: KolAmbalaze prazna"
    AssertEquals "", OtkPolje(otkID, COL_OTK_BRUTO), "OTK header: BrutoKg prazan"

    AssertEquals "", OtkPolje(otkID, COL_OTK_VOZAC), _
                 "OTK header: VozacID prazan (vozac pripada otpremnici)"
    AssertEquals "", OtkPolje(otkID, COL_OTK_ISPLACENO), _
                 "OTK header: Isplaceno prazno (read-model)"
    AssertEquals "", OtkPolje(otkID, COL_OTK_DATUM_ISPLATE), _
                 "OTK header: DatumIsplate prazan"
    AssertEquals "", OtkPolje(otkID, COL_OTK_VREME_UNOSA), _
                 "OTK header: VremeUnosa prazno (CreatedAt/SourceCreatedAt)"

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
    AssertEquals "", OtkPolje(otkID, COL_OTK_KOLICINA), _
                 "OTK samo II: header i dalje ne nosi kolicinu"

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
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTPSM")

    Dim stariID As String
    stariID = SaveOtkup_TX(NextTestDate(), TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, _
                           TEST_SORTA, 400#, 50#, TEST_TIP_AMB, 20, TEST_VOZ_ID, _
                           TEST_PREFIX & "-OTK-SM-" & scenario, 0#, "", KLASA_I)

    AssertTrue Len(stariID) > 0, "OTP stari otkup: stari pisac je napravio red"
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

    Exit Sub

EH:
    LogFatal "Test_OTP_StariOtkupNeUlazi", Err.Number, Err.description
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
    AssertEquals "", OtkPolje(res, COL_OTK_KOLICINA), "OTK ekran: header ne nosi kolicinu"
    AssertEquals "", OtkPolje(res, COL_OTK_CENA), "OTK ekran: header ne nosi cenu"

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
    AssertEquals "", OtkPolje(otkID, COL_OTK_KOLICINA), "PWA: header ne nosi kolicinu"
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

    ' Isti CRID, promenjena kolicina.
    Dim izmenjen As Variant
    izmenjen = PwaRed(crid, TEST_PREFIX & "-OTK-PWAKF-" & scenario, 999#, 50#, 20)

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

' Red kakav PWA salje u OTK sheet-u. Indeksi su GS_* kolone modMasterSync-a;
' one su Private tamo, pa se ovde imenuju komentarom, ne konstantom.
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

' PANEL BLOKOVA JE PAUZIRAN DO PR7 -- i ovaj test meri ZASTO.
'
' SumKolByOtp sabira tblOtkup.Kolicina po Otkup.OtpremnicaID. Nov pisac tu kolonu
' ostavlja praznu, pa otkup od 1000 kg panelu izgleda kao NULA. Test to tvrdi
' izricito, umesto da se oslanja na to sto UI niko ne izvrsava u suite-u.
'
' U PR7 ce ovaj test POCRVENETI -- citalac tada prelazi na tblOtpremnicaIzvori i
' GetOtpremnicaProgress. To je namerno: pauza se tada mora SVESNO skinuti, a ne
' zaboraviti ukljucenom.
Private Sub Test_OTK_PanelNapredakJePauziran()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPN")

    Dim otkID As String
    otkID = CreateOtkup_TX(OtkHeader(TEST_PREFIX & "-OTK-PN-" & scenario), _
                           OtkStavke(1000#, 50#, 20, 0#, 0#, 0))
    AssertTrue Len(otkID) > 0, "OTK panel: otkup od 1000 kg upisan"
    AssertTrue Abs(OtkStavkaBrojP(otkID, KLASA_I, COL_OKS_KOLICINA) - 1000#) < 0.001, _
               "OTK panel: kolicina je NA STAVCI"

    ' Panel vezuje blok za otpremnicu upisom OtpremnicaID -- to i dalje radi.
    Dim otpID As String
    otpID = SaveOtpremnica_TX(NextTestDate(), TEST_ST_ID, TEST_VOZ_ID, _
                              TEST_PREFIX & "-OTP-PN-" & scenario, "", _
                              TEST_VRSTA, TEST_SORTA, 1000#, 50#, TEST_TIP_AMB, 20)
    AssertTrue Len(otpID) > 0, "OTK panel: otpremnica napravljena"

    RequireUpdateCell TBL_OTKUP, FindRows(TBL_OTKUP, COL_OTK_ID, otkID)(1), _
                      COL_OTK_OTPREMNICA_ID, otpID, "Test_OTK_PanelNapredakJePauziran"

    ' MERENA REGRESIJA: stari citalac vidi nulu tamo gde je 1000 kg.
    AssertTrue Abs(modOtkupBlok.SumKolByOtp(otpID)) < 0.001, _
               "OTK panel: stari citalac vidi 0 -- brojevi su na stavkama"

    ' Zato panel te brojeve NE prikazuje.
    AssertTrue Not modOtkupBlok.NapredakBlokaDostupan(), _
               "OTK panel: napredak je pauziran dok citalac ne predje (PR7)"

    Exit Sub

EH:
    LogFatal "Test_OTK_PanelNapredakJePauziran", Err.Number, Err.description
End Sub

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
Private Sub Test_OTK_VrednostBezStavkiPada()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKVS")

    ' Stari pisac pravi red BEZ stavki -- tacno oblik koji kapija mora da uhvati.
    Dim stariID As String
    stariID = SaveOtkup_TX(NextTestDate(), TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, _
                           TEST_SORTA, 400#, 50#, TEST_TIP_AMB, 20, TEST_VOZ_ID, _
                           TEST_PREFIX & "-OTK-VS-" & scenario, 0#, "", KLASA_I)

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

    Exit Sub

EH:
    LogFatal "Test_OTK_VrednostBezStavkiPada", Err.Number, Err.description
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
Private Function Pr3Otpremnica(ByVal broj As String, ByVal klasa As String, _
                               ByVal kol As Double, ByVal amb As Long, _
                               Optional ByVal vrsta As String = "") As String
    If Len(vrsta) = 0 Then vrsta = TEST_VRSTA
    Pr3Otpremnica = SaveOtpremnica_TX(NextTestDate(), TEST_ST_ID, TEST_VOZ_ID, _
                                      broj, "", vrsta, TEST_SORTA, kol, 50#, _
                                      TEST_TIP_AMB, amb, klasa)
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

Private Sub Pr3PostaviAmbalazu(ByVal otpID As String, ByVal amb As Double)
    Dim redovi As Collection
    Set redovi = FindRows(TBL_OTPREMNICA, COL_OTP_ID, otpID)
    If redovi Is Nothing Then Exit Sub
    If redovi.count <> 1 Then Exit Sub
    RequireUpdateCell TBL_OTPREMNICA, CLng(redovi(1)), COL_OTP_KOL_AMB, amb, _
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
                                    ByVal vozac As String) As String
    Pr3OtpremnicaVozac = SaveOtpremnica_TX(NextTestDate(), TEST_ST_ID, vozac, _
                                           broj, "", TEST_VRSTA, TEST_SORTA, _
                                           kol, 50#, TEST_TIP_AMB, amb, klasa)
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

Private Sub LogInfo(ByVal message As String)
    Debug.Print "[INFO] " & message
    AppendTestLog "INFO", "", "INFO", message
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







Public Function CreateSEFLiveTestFaktura() As String
    On Error GoTo EH

    BeginRun "CREATE SEF LIVE TEST FAKTURA"

    SeedBusinessFlowProMasterData

    Dim scenario As String
    scenario = NewScenarioCode("SEFLIVE")

    Dim testDate As Date
    testDate = Date

    Dim brojOtk As String
    Dim brojOtp As String
    Dim brojZbirne As String
    Dim brojPrij As String

    brojOtk = TEST_PREFIX & "-OTK-" & scenario
    brojOtp = TEST_PREFIX & "-OTP-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-" & scenario

    Dim otkupResult As String
    otkupResult = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        1000#, 120#, TEST_TIP_AMB, 100, TEST_VOZ_ID, brojOtk, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        True, 200#, 80#)

    Dim otpI As String
    Dim otpII As String

    otpI = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 1000#, 120#, TEST_TIP_AMB, 100, "I")

    otpII = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
                              TEST_VRSTA, TEST_SORTA, 200#, 80#, TEST_TIP_AMB, 0, "II")

    Dim zbrI As String
    Dim zbrII As String

    zbrI = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                         "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                         1000#, TEST_TIP_AMB, 100, "I")

    zbrII = SaveZbirna_TX(testDate, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
                          "Test Hladnjaca", "Test Pogon", TEST_VRSTA, TEST_SORTA, _
                          200#, TEST_TIP_AMB, 0, "II")

    Dim prjI As String
    Dim prjII As String

    prjI = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 990#, 120#, TEST_TIP_AMB, 100, 95, "I")

    prjII = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
                              TEST_VRSTA, TEST_SORTA, 190#, 80#, TEST_TIP_AMB, 0, 0, "II")

    AutoLinkOtkupOtpremnica_TX

    Dim stavke As Collection
    Set stavke = New Collection

    stavke.Add Array(prjI, 990#, 120#, "I", brojPrij)
    stavke.Add Array(prjII, 190#, 80#, "II", brojPrij)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(TEST_KUP_ID, stavke)

    LogInfo "Created SEF live test faktura=" & fakturaID

    CreateSEFLiveTestFaktura = fakturaID

    EndRun
    Exit Function

EH:
    LogFatal "CreateSEFLiveTestFaktura", Err.Number, Err.description
    CreateSEFLiveTestFaktura = ""
    EndRun
End Function


Public Function CreateSEFLiveDummyFaktura() As String
    On Error GoTo EH

    BeginRun "CREATE SEF LIVE DUMMY FAKTURA"

    SeedBusinessFlowProMasterData

    Dim scenario As String
    scenario = NewScenarioCode("SEFLIVE")

    Dim d As Date
    d = Date

    Dim brojOtk As String
    Dim brojOtp As String
    Dim brojZbirne As String
    Dim brojPrij As String

    brojOtk = TEST_PREFIX & "-OTK-" & scenario
    brojOtp = TEST_PREFIX & "-OTP-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario
    brojPrij = TEST_PREFIX & "-PRJ-" & scenario
    
    Dim otkupResult As String

    otkupResult = SaveOtkupMulti_TX( _
        d, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        1000#, 120#, TEST_TIP_AMB, 100, TEST_VOZ_ID, brojOtk, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        True, 200#, 80#)

    If Len(Trim$(otkupResult)) = 0 Then          ' ? ovde
        Err.Raise vbObjectError + 9301, "CreateSEFLiveDummyFaktura", _
              "SaveOtkupMulti_TX failed."
    End If
    
    Dim otpI As String
    Dim otpII As String

    otpI = SaveOtpremnica_TX( _
        d, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
        TEST_VRSTA, TEST_SORTA, 1000#, 120#, TEST_TIP_AMB, 100, "I")

    otpII = SaveOtpremnica_TX( _
        d, TEST_ST_ID, TEST_VOZ_ID, brojOtp, brojZbirne, _
        TEST_VRSTA, TEST_SORTA, 200#, 80#, TEST_TIP_AMB, 0, "II")

    If Len(Trim$(otpI)) = 0 Or Len(Trim$(otpII)) = 0 Then     ' ? ovde
        Err.Raise vbObjectError + 9302, "CreateSEFLiveDummyFaktura", _
              "SaveOtpremnica_TX failed."
    End If

    Dim zbrI As String
    Dim zbrII As String
    
    zbrI = SaveZbirna_TX( _
        d, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
        "Test Hladnjaca", "Test Pogon", _
        TEST_VRSTA, TEST_SORTA, 1000#, TEST_TIP_AMB, 100, "I")

    zbrII = SaveZbirna_TX( _
        d, TEST_VOZ_ID, brojZbirne, TEST_KUP_ID, _
        "Test Hladnjaca", "Test Pogon", _
        TEST_VRSTA, TEST_SORTA, 200#, TEST_TIP_AMB, 0, "II")

    If Len(Trim$(zbrI)) = 0 Or Len(Trim$(zbrII)) = 0 Then     ' ? ovde
        Err.Raise vbObjectError + 9303, "CreateSEFLiveDummyFaktura", _
              "SaveZbirna_TX failed."
    End If
    
    Dim prjI As String
    Dim prjII As String

    prjI = SavePrijemnica_TX( _
        d, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
        TEST_VRSTA, TEST_SORTA, 990#, 120#, TEST_TIP_AMB, 100, 95, "I")

    prjII = SavePrijemnica_TX( _
        d, TEST_KUP_ID, TEST_VOZ_ID, brojPrij, brojZbirne, _
        TEST_VRSTA, TEST_SORTA, 190#, 80#, TEST_TIP_AMB, 0, 0, "II")

    If Len(Trim$(prjI)) = 0 Or Len(Trim$(prjII)) = 0 Then     ' ? ovde
        Err.Raise vbObjectError + 9304, "CreateSEFLiveDummyFaktura", _
              "SavePrijemnica_TX failed."
    End If

    AutoLinkOtkupOtpremnica_TX

    Dim stavke As Collection
    Set stavke = New Collection

    stavke.Add Array(prjI, 990#, 120#, "I", brojPrij)
    stavke.Add Array(prjII, 190#, 80#, "II", brojPrij)

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(TEST_KUP_ID, stavke)

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise vbObjectError + 9300, "CreateSEFLiveDummyFaktura", _
                  "CreateFaktura_TX returned empty FakturaID."
    End If

    LogInfo "Created SEF live dummy faktura=" & fakturaID
    LogInfo "Otkup=" & otkupResult
    LogInfo "Otpremnica=" & otpI & "/" & otpII
    LogInfo "Zbirna=" & zbrI & "/" & zbrII
    LogInfo "Prijemnica=" & prjI & "/" & prjII

    CreateSEFLiveDummyFaktura = fakturaID

    EndRun
    Exit Function

EH:
    LogFatal "CreateSEFLiveDummyFaktura", Err.Number, Err.description
    CreateSEFLiveDummyFaktura = ""
    EndRun
End Function



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


