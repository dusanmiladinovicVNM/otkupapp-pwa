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
Private m_Skipped As Long
Private m_RunID As String
Private m_DateSeq As Long

Private Const TEST_LOG_SHEET As String = "BUSINESS_FLOW_PRO_TEST_LOG"

Private Const TEST_ST_ID As String = "ST-90001"
Private Const TEST_KOOP_ID As String = "KOOP-90001"
Private Const TEST_VOZ_ID As String = "VOZ-90001"
Private Const TEST_KUP_ID As String = "KUP-90001"
Private Const TEST_KULTURA_ID As String = "KUL-90001"
Private Const TEST_PAR_ID As String = "PAR-90001"

Private Const TEST_VRSTA As String = "Test Jabuka"
Private Const TEST_SORTA As String = "Test Sorta"
Private Const TEST_TIP_AMB As String = "Test Gajba"

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

    ' RF-05 (frmDokumenta unos + storno set)
    Test_ProsekGajbeExcludesStornirano
    Test_OpenFaktureExcludeStornirano
    Test_ZbirnaKlasaIIGuard
    Test_PrefillBiraPoslednjuGeneraciju
    Test_GeneracijaIDNaSavePutanji
    Test_MalinaAutoZbirnaFailSignal
    Test_ZbirnaRowDataColumnMapped
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

Private Function RowHasTestMarker(ByVal data As Variant, ByVal rowIndex As Long, _
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
    SeedKupac
    SeedKupac2
    SeedKultura
    SeedKooperant
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

    prevMode = GetConfigValue(CFG_KEY_MALINA_MODE)
    SetConfigValue CFG_KEY_MALINA_MODE, "YES"

    ' Posle Ensure vozac mora postojati (kreiran sad ili od ranijeg run-a).
    Call EnsureVozacMirrorForStanica(MIR_ST, "Test Naziv", "Test Mesto", "")
    AssertTrue RowExists(TBL_VOZACI, "VozacID", MIR_ST), _
        "Malina mirror: vozac VozacID==StanicaID postoji posle Ensure"

    ' Idempotencija: ponovni poziv NE sme da kreira nov red.
    AssertFalse EnsureVozacMirrorForStanica(MIR_ST, "Test Naziv", "Test Mesto", ""), _
        "Malina mirror: ponovni poziv ne kreira duplikat (idempotentno)"

    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    Exit Sub

EH:
    On Error Resume Next
    SetConfigValue CFG_KEY_MALINA_MODE, prevMode
    On Error GoTo 0
    LogFatal "Test_MalinaVozacMirror", Err.Number, Err.description
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
    ' 32=I nove). Bez eksplicitne generacije heuristika kontinuiteta bi spojila
    ' novu Kl.I sa starom Kl.II.
    Dim d As Variant
    ReDim d(1 To 3, 1 To 4)
    d(1, 1) = "DOK-1": d(1, 2) = "I":  d(1, 3) = "OTP-00030": d(1, 4) = "GEN-00001"
    d(2, 1) = "DOK-1": d(2, 2) = "II": d(2, 3) = "OTP-00031": d(2, 4) = "GEN-00001"
    d(3, 1) = "DOK-1": d(3, 2) = "I":  d(3, 3) = "OTP-00032": d(3, 4) = "GEN-00002"

    Dim rI As Long, rII As Long
    PickLatestGenerationRows d, 1, 2, 3, 4, "DOK-1", rI, rII
    AssertEquals "3", CStr(rI), "Prefill: Kl.I iz poslednje generacije (uzastopni ID-evi)"
    AssertEquals "0", CStr(rII), _
                 "Prefill: stara Kl.II (ID 31) se NE spaja sa novom Kl.I (ID 32)"

    ' Obe generacije dvoklasne -> I i II iz ISTE (novije).
    Dim g As Variant
    ReDim g(1 To 4, 1 To 4)
    g(1, 1) = "DOK-2": g(1, 2) = "I":  g(1, 3) = "OTP-00040": g(1, 4) = "GEN-00010"
    g(2, 1) = "DOK-2": g(2, 2) = "II": g(2, 3) = "OTP-00041": g(2, 4) = "GEN-00010"
    g(3, 1) = "DOK-2": g(3, 2) = "I":  g(3, 3) = "OTP-00042": g(3, 4) = "GEN-00011"
    g(4, 1) = "DOK-2": g(4, 2) = "II": g(4, 3) = "OTP-00043": g(4, 4) = "GEN-00011"

    PickLatestGenerationRows g, 1, 2, 3, 4, "DOK-2", rI, rII
    AssertEquals "3", CStr(rI), "Prefill: Kl.I iz novije generacije"
    AssertEquals "4", CStr(rII), "Prefill: Kl.II iz ISTE (novije) generacije"

    ' Poslovni datum ne ucestvuje: novija generacija sa RANIJIM datumom pobedjuje.
    ' (Kolona 4 je generacija; datum bi bio zaseban, ovde je namerno van izbora.)
    Dim h As Variant
    ReDim h(1 To 2, 1 To 4)
    h(1, 1) = "DOK-3": h(1, 2) = "I": h(1, 3) = "OTP-00050": h(1, 4) = "GEN-00020"
    h(2, 1) = "DOK-3": h(2, 2) = "I": h(2, 3) = "OTP-00051": h(2, 4) = "GEN-00021"

    PickLatestGenerationRows h, 1, 2, 3, 4, "DOK-3", rI, rII
    AssertEquals "2", CStr(rI), "Prefill: kasnija generacija pobedjuje (nezavisno od datuma)"

    ' Fallback bez generacije: samo poslednji red, druga klasa se NE pogadja.
    Dim f As Variant
    ReDim f(1 To 3, 1 To 4)
    f(1, 1) = "DOK-4": f(1, 2) = "I":  f(1, 3) = "OTP-00060": f(1, 4) = ""
    f(2, 1) = "DOK-4": f(2, 2) = "II": f(2, 3) = "OTP-00061": f(2, 4) = ""
    f(3, 1) = "DOK-4": f(3, 2) = "I":  f(3, 3) = "OTP-00062": f(3, 4) = ""

    PickLatestGenerationRows f, 1, 2, 3, 4, "DOK-4", rI, rII
    AssertEquals "3", CStr(rI), "Prefill fallback: poslednji upisan red"
    AssertEquals "0", CStr(rII), _
                 "Prefill fallback: bez generacije druga klasa ostaje prazna"

    ' Isti fallback i kad kolone uopste nema (cGen = 0).
    PickLatestGenerationRows f, 1, 2, 3, 0, "DOK-4", rI, rII
    AssertEquals "3", CStr(rI), "Prefill fallback: bez GeneracijaID kolone"
    AssertEquals "0", CStr(rII), "Prefill fallback: bez kolone nema druge klase"

    ' Nepostojeci broj -> nista.
    PickLatestGenerationRows g, 1, 2, 3, 4, "DOK-NEMA", rI, rII
    AssertEquals "00", CStr(rI) & CStr(rII), "Prefill: nepoznat broj ne vraca red"

    Exit Sub

EH:
    LogFatal "Test_PrefillBiraPoslednjuGeneraciju", Err.Number, Err.description
End Sub

' End-to-end: save putanja stvarno pise GeneracijaID, i to ISTU za obe klase
' jednog Multi_TX upisa, a NOVU za ispravku istog broja.
Private Sub Test_GeneracijaIDNaSavePutanji()
    On Error GoTo EH

    If GetColumnIndex(TBL_OTPREMNICA, COL_GENERACIJA_ID) = 0 Then
        LogSkip "GeneracijaID na save putanji", _
               "Kolona GeneracijaID ne postoji (pokreni EnsureSledljivostSchema)"
        Exit Sub
    End If

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
    PickLatestGenerationRows d, cBr, cKl, cId, cGen, brojOtp, rI, rII

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

Private Function OtpGeneracija(ByVal brojOtp As String, ByVal klasa As String) As String
    Dim otpID As String
    otpID = FindOtpremnicaIDByBrojAndKlasa(brojOtp, klasa)
    If Len(otpID) = 0 Then Exit Function

    OtpGeneracija = Trim$(CStr(nz(GetValueByKey(TBL_OTPREMNICA, COL_OTP_ID, otpID, _
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

Private Function ArrayContainsKeyValue(ByVal data As Variant, _
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

    If m_Failed > 0 Then
        MsgBox "Business Flow Pro tests finished with failures." & vbCrLf & summary, _
               vbExclamation, APP_NAME
    Else
        MsgBox "Business Flow Pro tests finished." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If
End Sub

Private Sub ResetCounters()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

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
End Sub

Private Sub LogFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendTestLog "TEST", testName, "FAIL", details
End Sub

Private Sub LogSkip(ByVal testName As String, ByVal reason As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1

    Debug.Print "[SKIP] " & testName & " :: " & reason
    AppendTestLog "TEST", testName, "SKIP", reason
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

    deleted = DeleteTestRowsFromTable(TBL_ZBIRNA, Array("BrojZbirne"))
    total = total + deleted
    Debug.Print "tblZbirna: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_OTPREMNICA, Array("BrojOtpremnice", "BrojZbirne"))
    total = total + deleted
    Debug.Print "tblOtpremnica: " & deleted & " obrisano"

    deleted = DeleteTestRowsFromTable(TBL_OTKUP, Array("BrojDokumenta", "BrojZbirne"))
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

Private Function RowHasTestPrefix(ByVal data As Variant, ByVal rowIndex As Long, _
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


