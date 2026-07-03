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
Private Const TEST_TIP_AMB_2KG As String = "TST-PRO Gajba2kg"

Private Const TEST_PREFIX As String = "TST-PRO"

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

    Test_OtkupSamoKlasaII
    Test_OtkupKesIsplataKnjiziNovac
    Test_OtkupAmbalazaDvojniUpis
    Test_OtkupBrutoIVremeUnosa
    Test_OtkupMultiResultFormat
    Test_KulturaIDFallback
    Test_OtkupAutoAvansPriSnimanju
    Test_OtkupRejectLeavesAllTablesUntouched
    Test_GetSaldoByStation
    Test_OtkupReadHelpersDateRange
    Test_ComputeNetoFromBruto
    Test_LinkOtkupIDsToOtpremnica
    Test_SumHelpersByOtp
    Test_PrijemnicaBrojZaZbirnu
    Test_StornoOtkupObeKlase
    Test_LostBlokAdoptFlow

    Test_DokumentaInputValidationHardening
    Test_DokumentaReadHelpersExcludeStornirano
    Test_DualClassDocumentWrappers
    Test_MalinaAutoZbirnaFromOtpremnice
    Test_MalinaVozacMirror

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

' Fokusiran pod-suite: samo otkup sloj (modOtkup + otkup-blok data helperi).
' Za brzu regresiju posle izmena u modOtkup / frmOtkup / modOtkupBlok.
Public Sub RunOtkupCoverageOnly()
    On Error GoTo EH

    BeginRun "OTKUP COVERAGE ONLY"

    Test_CoreTablesAndColumnsExist
    SeedBusinessFlowProMasterData
    Test_SeedMasterDataAvailable

    Test_OtkupAtomicMultiClassSave
    Test_OtkupClassIIAmbalaza
    Test_OtkupInputValidationHardening
    Test_OtkupReadHelpersExcludeStornirano

    Test_OtkupSamoKlasaII
    Test_OtkupKesIsplataKnjiziNovac
    Test_OtkupAmbalazaDvojniUpis
    Test_OtkupBrutoIVremeUnosa
    Test_OtkupMultiResultFormat
    Test_KulturaIDFallback
    Test_OtkupAutoAvansPriSnimanju
    Test_OtkupRejectLeavesAllTablesUntouched
    Test_GetSaldoByStation
    Test_OtkupReadHelpersDateRange
    Test_ComputeNetoFromBruto
    Test_LinkOtkupIDsToOtpremnica
    Test_SumHelpersByOtp
    Test_PrijemnicaBrojZaZbirnu
    Test_StornoOtkupObeKlase
    Test_LostBlokAdoptFlow

    EndRun
    Exit Sub

EH:
    LogFatal "RunOtkupCoverageOnly", Err.Number, Err.description
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

    Dim result As String
    result = SavePrijemnica_TX( _
        NextTestDate(), TEST_KUP_ID, TEST_VOZ_ID, _
        TEST_PREFIX & "-BAD-PRJ-" & NewScenarioCode("NEGAMB"), _
        TEST_PREFIX & "-BAD-ZBR-" & NewScenarioCode("NEGAMB"), _
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

' ============================================================
' OTKUP COVERAGE (modOtkup + otkup-blok data sloj)
' ============================================================

' Samo Klasa II (kolicinaI = 0): kes i izdata ambalaza MORAJU na red Klase II
' (SaveOtkupMulti_TX preusmerava novac/kolAmbIzdata kad Klase I nema).
Private Sub Test_OtkupSamoKlasaII()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKII")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojDok As String
    Dim brojZbirne As String
    brojDok = TEST_PREFIX & "-OTK-II-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim hasIzdataCol As Boolean
    hasIzdataCol = (GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA) > 0)

    Dim izdata As Long
    If hasIzdataCol Then izdata = 5 Else izdata = 0

    Dim beforeOtkup As Long
    Dim beforeNovac As Long
    beforeOtkup = CountRows(TBL_OTKUP)
    beforeNovac = CountRows(TBL_NOVAC)

    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        0#, 0#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojDok, _
        500#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        hasKlasaII:=True, kolicinaII:=150#, cenaII:=70#, kolAmbIzdata:=izdata)

    AssertTrue Len(Trim$(result)) > 0, "Samo Klasa II: save returns ID"
    AssertEquals "0", CStr(InStr(result, " + ")), "Samo Klasa II: single ID (no plus separator)"
    AssertEquals CStr(beforeOtkup + 1), CStr(CountRows(TBL_OTKUP)), _
                 "Samo Klasa II: exactly one otkup row appended"

    Dim otkII As String
    otkII = FindOtkupIDByBrojAndKlasa(brojDok, "II")

    AssertEquals otkII, result, "Samo Klasa II: result is the class II row"
    AssertEquals "", FindOtkupIDByBrojAndKlasa(brojDok, "I"), "Samo Klasa II: no class I row"

    AssertEquals "500", CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkII, COL_OTK_NOVAC)), _
                 "Samo Klasa II: kes zabelezen na redu Klase II"

    If hasIzdataCol Then
        AssertEquals "5", CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkII, COL_OTK_KOL_AMB_IZDATA)), _
                     "Samo Klasa II: izdata ambalaza zabelezena na redu Klase II"
    Else
        LogSkip "Samo Klasa II: izdata ambalaza", "Nema kolone " & COL_OTK_KOL_AMB_IZDATA
    End If

    AssertEquals CStr(beforeNovac + 1), CStr(CountRows(TBL_NOVAC)), _
                 "Samo Klasa II: exactly one novac row appended"
    AssertEquals NOV_KES_OTKUPAC_KOOP, _
                 CStr(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, otkII, COL_NOV_TIP)), _
                 "Samo Klasa II: novac red vezan na OtkupID Klase II"

    Exit Sub

EH:
    LogFail "Otkup samo Klasa II", Err.description
End Sub

' Kes isplata na otkupu (novac > 0) -> tacno jedan tblNovac red
' (NOV_KES_OTKUPAC_KOOP) vezan na primarni OtkupID, sa imenom kooperanta;
' nepoznat kooperant -> Partner fallback na KooperantID; novac = 0 -> bez reda.
Private Sub Test_OtkupKesIsplataKnjiziNovac()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKKES")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim beforeNovac As Long
    beforeNovac = CountRows(TBL_NOVAC)

    Dim brojDok As String
    brojDok = TEST_PREFIX & "-OTK-KES-" & scenario

    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojDok, _
        300#, "TEST PRIMALAC", GetTestParcelaID(), brojZbirne)

    AssertTrue Len(Trim$(result)) > 0, "Kes isplata: save returns ID"
    AssertEquals CStr(beforeNovac + 1), CStr(CountRows(TBL_NOVAC)), _
                 "Kes isplata: exactly one novac row appended"

    AssertEquals NOV_KES_OTKUPAC_KOOP, _
                 CStr(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result, COL_NOV_TIP)), _
                 "Kes isplata: novac tip je KesOtkupacKoop"
    AssertDoubleNear 300#, TestNumVal(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result, COL_NOV_ISPLATA)), _
                     0.001, "Kes isplata: iznos isplate = novac"
    AssertDoubleNear 0#, TestNumVal(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result, COL_NOV_UPLATA)), _
                     0.001, "Kes isplata: uplata je nula"
    AssertEquals "Test Kooperant", _
                 CStr(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result, COL_NOV_PARTNER)), _
                 "Kes isplata: partner je Ime + Prezime kooperanta"
    AssertEquals TEST_KOOP_ID, _
                 CStr(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result, COL_NOV_KOOP_ID)), _
                 "Kes isplata: novac red nosi KooperantID"
    AssertEquals "TEST PRIMALAC", _
                 CStr(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result, COL_NOV_NAPOMENA)), _
                 "Kes isplata: primalac zabelezen u napomeni"
    AssertEquals brojDok, _
                 CStr(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result, COL_NOV_BROJ_DOK)), _
                 "Kes isplata: novac red nosi broj otkupnog bloka"

    ' Fallback: kooperant koji NE postoji u tblKooperanti -> Partner = KooperantID.
    Dim koopX As String
    koopX = "KOOP-TSTX-" & scenario

    Dim brojDok2 As String
    brojDok2 = TEST_PREFIX & "-OTK-KESX-" & scenario

    Dim result2 As String
    result2 = SaveOtkupMulti_TX( _
        testDate, koopX, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojDok2, _
        100#, "TEST OPERATOR", "", brojZbirne)

    AssertTrue Len(Trim$(result2)) > 0, "Kes isplata: fallback save returns ID"
    AssertEquals koopX, _
                 CStr(GetValueByKey(TBL_NOVAC, COL_NOV_OTKUP_ID, result2, COL_NOV_PARTNER)), _
                 "Kes isplata: nepoznat kooperant -> Partner = KooperantID"

    ' Bez kesa (novac = 0) -> nema novog novac reda.
    Dim beforeNoCash As Long
    beforeNoCash = CountRows(TBL_NOVAC)

    Dim brojDok3 As String
    brojDok3 = TEST_PREFIX & "-OTK-KES0-" & scenario

    Dim result3 As String
    result3 = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojDok3, _
        0#, "TEST OPERATOR", "", brojZbirne)

    AssertTrue Len(Trim$(result3)) > 0, "Kes isplata: zero-cash save returns ID"
    AssertEquals CStr(beforeNoCash), CStr(CountRows(TBL_NOVAC)), _
                 "Kes isplata: novac = 0 ne knjizi novac red"

    Exit Sub

EH:
    LogFail "Otkup kes isplata knjizi novac", Err.description
End Sub

' Ambalaza dvojni upis: primljena (kolAmb) = Kooperant Izlaz + Stanica Ulaz
' (DOK_TIP_OTKUP); izdata (kolAmbIzdata) = Kooperant Ulaz + Stanica Izlaz
' (DOK_TIP_OM_IZLAZ_KOOP). Tacno po jedna noga, ispravne kolicine.
Private Sub Test_OtkupAmbalazaDvojniUpis()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKAMB")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojDok As String
    Dim brojZbirne As String
    brojDok = TEST_PREFIX & "-OTK-AMB-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim hasIzdataCol As Boolean
    hasIzdataCol = (GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA) > 0)

    Dim izdata As Long
    If hasIzdataCol Then izdata = 4 Else izdata = 0

    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 10, TEST_VOZ_ID, brojDok, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        kolAmbIzdata:=izdata)

    AssertTrue Len(Trim$(result)) > 0, "Amb dvojni upis: save returns ID"

    Dim legs As Long
    Dim kol As Double

    kol = SumAmbalazaLegKolicina(result, TEST_KOOP_ID, "Izlaz", DOK_TIP_OTKUP, legs)
    AssertEquals "1", CStr(legs), "Amb dvojni upis: jedna noga Kooperant-Izlaz (otkup)"
    AssertDoubleNear 10#, kol, 0.001, "Amb dvojni upis: Kooperant-Izlaz kolicina"

    kol = SumAmbalazaLegKolicina(result, TEST_ST_ID, "Ulaz", DOK_TIP_OTKUP, legs)
    AssertEquals "1", CStr(legs), "Amb dvojni upis: jedna noga Stanica-Ulaz (otkup)"
    AssertDoubleNear 10#, kol, 0.001, "Amb dvojni upis: Stanica-Ulaz kolicina"

    If hasIzdataCol Then
        kol = SumAmbalazaLegKolicina(result, TEST_KOOP_ID, "Ulaz", DOK_TIP_OM_IZLAZ_KOOP, legs)
        AssertEquals "1", CStr(legs), "Amb dvojni upis: jedna noga Kooperant-Ulaz (izdata)"
        AssertDoubleNear 4#, kol, 0.001, "Amb dvojni upis: Kooperant-Ulaz izdata kolicina"

        kol = SumAmbalazaLegKolicina(result, TEST_ST_ID, "Izlaz", DOK_TIP_OM_IZLAZ_KOOP, legs)
        AssertEquals "1", CStr(legs), "Amb dvojni upis: jedna noga Stanica-Izlaz (izdata)"
        AssertDoubleNear 4#, kol, 0.001, "Amb dvojni upis: Stanica-Izlaz izdata kolicina"

        AssertEquals "4", CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, result, COL_OTK_KOL_AMB_IZDATA)), _
                     "Amb dvojni upis: KolAmbIzdata upisana po imenu kolone"
    Else
        LogSkip "Amb dvojni upis: izdata ambalaza", "Nema kolone " & COL_OTK_KOL_AMB_IZDATA
    End If

    Exit Sub

EH:
    LogFail "Otkup ambalaza dvojni upis", Err.description
End Sub

' BrutoKg: upis po imenu kolone kad je bruto rezim dao neto (brutoKgI > 0);
' bez bruto unosa kolona ostaje prazna (prazno = unet neto). VremeUnosa se
' puni pri svakom snimanju.
Private Sub Test_OtkupBrutoIVremeUnosa()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKBRT")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    If GetColumnIndex(TBL_OTKUP, COL_OTK_BRUTO) = 0 Then
        LogSkip "Otkup BrutoKg upis", "Nema kolone " & COL_OTK_BRUTO
        Exit Sub
    End If

    Dim brojDok1 As String
    brojDok1 = TEST_PREFIX & "-OTK-BRT-" & scenario

    Dim result1 As String
    result1 = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 10, TEST_VOZ_ID, brojDok1, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        brutoKgI:=120#)

    AssertTrue Len(Trim$(result1)) > 0, "BrutoKg: save returns ID"
    AssertEquals "120", CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, result1, COL_OTK_BRUTO)), _
                 "BrutoKg: bruto zamrznut u koloni BrutoKg"
    AssertEquals "100", CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, result1, COL_OTK_KOLICINA)), _
                 "BrutoKg: Kolicina nosi neto"

    If GetColumnIndex(TBL_OTKUP, COL_OTK_VREME_UNOSA) > 0 Then
        AssertTrue Len(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, result1, COL_OTK_VREME_UNOSA))) > 0, _
                   "BrutoKg: VremeUnosa popunjeno pri snimanju"
    Else
        LogSkip "BrutoKg: VremeUnosa", "Nema kolone " & COL_OTK_VREME_UNOSA
    End If

    Dim brojDok2 As String
    brojDok2 = TEST_PREFIX & "-OTK-NETO-" & scenario

    Dim result2 As String
    result2 = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 10, TEST_VOZ_ID, brojDok2, _
        0#, "TEST OPERATOR", "", brojZbirne)

    AssertTrue Len(Trim$(result2)) > 0, "BrutoKg: neto save returns ID"
    AssertEquals "", CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, result2, COL_OTK_BRUTO)), _
                 "BrutoKg: bez bruto unosa kolona ostaje prazna"

    Exit Sub

EH:
    LogFail "Otkup BrutoKg i VremeUnosa", Err.description
End Sub

' Format rezultata za dve klase: "OTK-x + OTK-y". Ugovor sa
' modOtkupBlok.LinkOtkupIDsToOtpremnica (Split na " + ").
Private Sub Test_OtkupMultiResultFormat()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKFMT")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojDok As String
    Dim brojZbirne As String
    brojDok = TEST_PREFIX & "-OTK-FMT-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojDok, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        hasKlasaII:=True, kolicinaII:=50#, cenaII:=5#)

    Dim otkI As String
    Dim otkII As String
    otkI = FindOtkupIDByBrojAndKlasa(brojDok, "I")
    otkII = FindOtkupIDByBrojAndKlasa(brojDok, "II")

    AssertTrue Len(otkI) > 0 And Len(otkII) > 0, "Result format: obe klase snimljene"
    AssertEquals otkI & " + " & otkII, result, _
                 "Result format: 'ID-I + ID-II' (ugovor za LinkOtkupIDsToOtpremnica)"

    Exit Sub

EH:
    LogFail "Otkup multi result format", Err.description
End Sub

' KulturaID fallback: vrsta koje nema u tblKulture -> KulturaID = "vrsta-sorta".
Private Sub Test_KulturaIDFallback()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKKUL")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim vrsta As String
    vrsta = TEST_PREFIX & "-VOCE-" & scenario

    Dim brojDok As String
    brojDok = TEST_PREFIX & "-OTK-KUL-" & scenario

    Dim result As String
    result = SaveOtkup_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, vrsta, "SortaX", _
        10#, 5#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojDok, _
        0#, "TEST OPERATOR", KLASA_I, "", TEST_PREFIX & "-ZBR-" & scenario)

    AssertTrue Len(Trim$(result)) > 0, "KulturaID fallback: save returns ID"
    AssertEquals vrsta & "-SortaX", _
                 CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, result, COL_OTK_KULTURA)), _
                 "KulturaID fallback: nepoznata vrsta -> vrsta-sorta"

    Exit Sub

EH:
    LogFail "KulturaID fallback", Err.description
End Sub

' Auto-avans pri snimanju: otvoren virman avans kooperanta se automatski
' vezuje na novi otkup unutar SaveOtkupMulti_TX (ApplyAvansToOtkup poziv),
' otkup se markira kao isplacen, ostatak avansa ostaje na originalnom redu.
Private Sub Test_OtkupAutoAvansPriSnimanju()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKAVN")

    Dim testDate As Date
    testDate = NextTestDate()

    ' Zaseban (sintetski) kooperant: izoluje avans od ostalih testova koji
    ' koriste TEST_KOOP_ID (da im auto-avans ne "pojede" fixture).
    Dim koopAv As String
    koopAv = "KOOP-TSTAV-" & scenario

    Dim avansID As String
    avansID = SaveNovac( _
        TEST_PREFIX & "-AVANS-" & scenario, testDate, _
        "TEST AVANS KOOP", koopAv, "Kooperant", _
        "", koopAv, "", "", _
        NOV_VIRMAN_AVANS_KOOP, _
        0#, 150#, _
        "TST avans fixture")

    If Len(Trim$(avansID)) = 0 Then
        LogFail "Otkup auto-avans pri snimanju", "Setup: SaveNovac avans nije uspeo."
        Exit Sub
    End If

    Dim brojDok As String
    brojDok = TEST_PREFIX & "-OTK-AVN-" & scenario

    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, koopAv, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 1#, TEST_TIP_AMB, 0, TEST_VOZ_ID, brojDok, _
        0#, "TEST OPERATOR", "", TEST_PREFIX & "-ZBR-" & scenario)

    AssertTrue Len(Trim$(result)) > 0, "Auto-avans: save returns ID"

    AssertDoubleNear 100#, GetIsplataForOtkup(result), 0.01, _
                     "Auto-avans: iznos otkupa vezan iz avansa pri snimanju"
    AssertEquals STATUS_ISPLACENO, _
                 CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, result, COL_OTK_ISPLACENO)), _
                 "Auto-avans: otkup markiran kao isplacen"
    AssertDoubleNear 50#, TestNumVal(LookupValue(TBL_NOVAC, COL_NOV_ID, avansID, COL_NOV_ISPLATA)), _
                     0.01, "Auto-avans: ostatak avansa na originalnom redu"

    Exit Sub

EH:
    LogFail "Otkup auto-avans pri snimanju", Err.description
End Sub

' Sve validacione grane SaveOtkupMulti_TX: odbijen unos NE sme da dira
' tblOtkup, tblAmbalaza NI tblNovac (atomicnost / rollback).
Private Sub Test_OtkupRejectLeavesAllTablesUntouched()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKREJ")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim beforeOtkup As Long
    Dim beforeAmb As Long
    Dim beforeNovac As Long
    beforeOtkup = CountRows(TBL_OTKUP)
    beforeAmb = CountRows(TBL_AMBALAZA)
    beforeNovac = CountRows(TBL_NOVAC)

    Dim r As String

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, "", TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ1-" & scenario, _
        0#, "TEST", "", brojZbirne)
    AssertEquals "", r, "Reject: prazna stanica"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        0#, 0#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ2-" & scenario, _
        0#, "TEST", "", brojZbirne)
    AssertEquals "", r, "Reject: nijedna klasa (kolicinaI = 0, bez Klase II)"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 0#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ3-" & scenario, _
        0#, "TEST", "", brojZbirne)
    AssertEquals "", r, "Reject: cena I = 0 uz kolicinu I"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ4-" & scenario, _
        0#, "TEST", "", brojZbirne, hasKlasaII:=True, kolicinaII:=0#, cenaII:=50#)
    AssertEquals "", r, "Reject: Klasa II bez kolicine"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ5-" & scenario, _
        0#, "TEST", "", brojZbirne, hasKlasaII:=True, kolicinaII:=50#, cenaII:=0#)
    AssertEquals "", r, "Reject: Klasa II bez cene"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, -1, TEST_VOZ_ID, TEST_PREFIX & "-REJ6-" & scenario, _
        100#, "TEST", "", brojZbirne)
    AssertEquals "", r, "Reject: negativna ambalaza (uz kes koji ne sme da se knjizi)"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ7-" & scenario, _
        0#, "TEST", "", brojZbirne, kolAmbIzdata:=-1)
    AssertEquals "", r, "Reject: negativna izdata ambalaza"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ8-" & scenario, _
        0#, "TEST", "", brojZbirne, hasKlasaII:=True, kolicinaII:=50#, cenaII:=5#, kolAmbII:=-1)
    AssertEquals "", r, "Reject: negativna ambalaza Klase II"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, TEST_PREFIX & "-REJ9-" & scenario, _
        -5#, "TEST", "", brojZbirne)
    AssertEquals "", r, "Reject: negativan novac"

    r = SaveOtkupMulti_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, "", 10, TEST_VOZ_ID, TEST_PREFIX & "-REJ10-" & scenario, _
        0#, "TEST", "", brojZbirne)
    AssertEquals "", r, "Reject: ambalaza bez tipa"

    AssertEquals CStr(beforeOtkup), CStr(CountRows(TBL_OTKUP)), _
                 "Reject: tblOtkup netaknut posle svih odbijanja"
    AssertEquals CStr(beforeAmb), CStr(CountRows(TBL_AMBALAZA)), _
                 "Reject: tblAmbalaza netaknuta posle svih odbijanja"
    AssertEquals CStr(beforeNovac), CStr(CountRows(TBL_NOVAC)), _
                 "Reject: tblNovac netaknut posle svih odbijanja"

    Exit Sub

EH:
    LogFail "Otkup reject leaves tables untouched", Err.description
End Sub

' GetSaldoByStation: agregat po kooperantu (kolicina / novac / ambalaza),
' stornirani redovi iskljuceni, nepoznata stanica -> Empty.
Private Sub Test_GetSaldoByStation()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKSAL")

    Dim testDate As Date
    testDate = NextTestDate()

    ' Sintetska stanica -> potpuna izolacija agregata od ostalih test redova.
    Dim st As String
    st = "ST-TSTSAL-" & scenario

    Dim koopA As String
    Dim koopB As String
    koopA = "KOOP-TSTSA-" & scenario
    koopB = "KOOP-TSTSB-" & scenario

    Dim idA1 As String
    Dim idA2 As String
    Dim idB1 As String
    Dim idStorno As String

    idA1 = SaveOtkup_TX(testDate, koopA, st, TEST_VRSTA, TEST_SORTA, _
                        100#, 10#, TEST_TIP_AMB, 2, TEST_VOZ_ID, _
                        TEST_PREFIX & "-SAL-A1-" & scenario, 50#, "TEST")
    idA2 = SaveOtkup_TX(testDate, koopA, st, TEST_VRSTA, TEST_SORTA, _
                        50#, 10#, TEST_TIP_AMB, 1, TEST_VOZ_ID, _
                        TEST_PREFIX & "-SAL-A2-" & scenario, 0#, "TEST")
    idB1 = SaveOtkup_TX(testDate, koopB, st, TEST_VRSTA, TEST_SORTA, _
                        30#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, _
                        TEST_PREFIX & "-SAL-B1-" & scenario, 10#, "TEST")
    idStorno = SaveOtkup_TX(testDate, koopA, st, TEST_VRSTA, TEST_SORTA, _
                            999#, 10#, TEST_TIP_AMB, 7, TEST_VOZ_ID, _
                            TEST_PREFIX & "-SAL-S1-" & scenario, 0#, "TEST")

    AssertTrue Len(idA1) > 0 And Len(idA2) > 0 And Len(idB1) > 0 And Len(idStorno) > 0, _
               "Saldo: fixture redovi snimljeni"

    MarkTestRowStornirano TBL_OTKUP, COL_OTK_ID, idStorno

    Dim saldo As Variant
    saldo = GetSaldoByStation(st)

    AssertTrue Not IsEmpty(saldo), "Saldo: rezultat nije Empty"

    If Not IsEmpty(saldo) Then
        AssertEquals "2", CStr(UBound(saldo, 1)), "Saldo: jedan red po kooperantu"

        AssertDoubleNear 150#, FindSaldoRowValue(saldo, koopA, 2), 0.001, _
                         "Saldo: kolicina kooperanta A (storno iskljucen)"
        AssertDoubleNear 50#, FindSaldoRowValue(saldo, koopA, 3), 0.001, _
                         "Saldo: novac kooperanta A"
        AssertDoubleNear 3#, FindSaldoRowValue(saldo, koopA, 4), 0.001, _
                         "Saldo: ambalaza kooperanta A"

        AssertDoubleNear 30#, FindSaldoRowValue(saldo, koopB, 2), 0.001, _
                         "Saldo: kolicina kooperanta B"
        AssertDoubleNear 10#, FindSaldoRowValue(saldo, koopB, 3), 0.001, _
                         "Saldo: novac kooperanta B"
        AssertDoubleNear 0#, FindSaldoRowValue(saldo, koopB, 4), 0.001, _
                         "Saldo: ambalaza kooperanta B"
    End If

    AssertTrue IsEmpty(GetSaldoByStation("ST-TSTSAL-EMPTY-" & scenario)), _
               "Saldo: nepoznata stanica -> Empty"

    Exit Sub

EH:
    LogFail "GetSaldoByStation", Err.description
End Sub

' Read helperi: BETWEEN granice datumskog filtera (ukljucive) i poziv bez
' datuma (bez filtera).
Private Sub Test_OtkupReadHelpersDateRange()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKDAT")

    Dim d1 As Date
    d1 = NextTestDate()

    Dim d2 As Date
    Dim d3 As Date
    d2 = d1 + 1
    d3 = d1 + 2

    Dim st As String
    Dim koopD As String
    st = "ST-TSTDAT-" & scenario
    koopD = "KOOP-TSTDT-" & scenario

    Dim id1 As String
    Dim id2 As String
    Dim id3 As String

    id1 = SaveOtkup_TX(d1, koopD, st, TEST_VRSTA, TEST_SORTA, 10#, 5#, TEST_TIP_AMB, 0, _
                       TEST_VOZ_ID, TEST_PREFIX & "-DAT1-" & scenario, 0#, "TEST")
    id2 = SaveOtkup_TX(d2, koopD, st, TEST_VRSTA, TEST_SORTA, 10#, 5#, TEST_TIP_AMB, 0, _
                       TEST_VOZ_ID, TEST_PREFIX & "-DAT2-" & scenario, 0#, "TEST")
    id3 = SaveOtkup_TX(d3, koopD, st, TEST_VRSTA, TEST_SORTA, 10#, 5#, TEST_TIP_AMB, 0, _
                       TEST_VOZ_ID, TEST_PREFIX & "-DAT3-" & scenario, 0#, "TEST")

    AssertTrue Len(id1) > 0 And Len(id2) > 0 And Len(id3) > 0, _
               "Date range: fixture redovi snimljeni"

    Dim data As Variant
    data = GetOtkupByStation(st, d1, d2)

    AssertTrue ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id1), _
               "Date range: donja granica ukljucena (stanica)"
    AssertTrue ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id2), _
               "Date range: gornja granica ukljucena (stanica)"
    AssertFalse ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id3), _
                "Date range: red van opsega iskljucen (stanica)"

    data = GetOtkupByStation(st)
    AssertTrue ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id1) _
               And ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id3), _
               "Date range: bez datuma vraca sve redove (stanica)"

    data = GetOtkupByKooperant(koopD, d2, d3)
    AssertFalse ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id1), _
                "Date range: red pre opsega iskljucen (kooperant)"
    AssertTrue ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id2) _
               And ArrayContainsKeyValue(data, TBL_OTKUP, COL_OTK_ID, id3), _
               "Date range: opseg ukljucuje granice (kooperant)"

    Exit Sub

EH:
    LogFail "Otkup read helpers date range", Err.description
End Sub

' modOtkup.ComputeNetoFromBruto: deljena bruto->neto tara logika
' (frmOtkup.btnUnos, zivi prikaz i modOtkupBlok.OtkupBlok_ConfirmUnos).
Private Sub Test_ComputeNetoFromBruto()
    On Error GoTo EH

    If GetTable(TBL_TIP_AMBALAZE) Is Nothing Then
        LogSkip "ComputeNetoFromBruto", "Nema tabele " & TBL_TIP_AMBALAZE
        Exit Sub
    End If

    SeedTipAmbalaze2kg

    Dim neto As Double
    Dim tara As Double

    AssertTrue ComputeNetoFromBruto(100#, 10, TEST_TIP_AMB_2KG, neto, tara), _
               "ComputeNetoFromBruto: validan bruto se konvertuje"
    AssertDoubleNear 20#, tara, 0.001, "ComputeNetoFromBruto: tara = kolAmb x tezina"
    AssertDoubleNear 80#, neto, 0.001, "ComputeNetoFromBruto: neto = bruto - tara"

    AssertFalse ComputeNetoFromBruto(100#, 50, TEST_TIP_AMB_2KG, neto, tara), _
                "ComputeNetoFromBruto: tara >= bruto odbijena"
    AssertDoubleNear 100#, tara, 0.001, _
                     "ComputeNetoFromBruto: tara vracena i na odbijanje"

    AssertFalse ComputeNetoFromBruto(100#, 10, TEST_PREFIX & "-NEMA-TEZINU", neto, tara), _
                "ComputeNetoFromBruto: tip bez tezine odbijen"
    AssertDoubleNear 0#, tara, 0.001, _
                     "ComputeNetoFromBruto: tip bez tezine -> tara 0 (signal pozivaocu)"

    AssertFalse ComputeNetoFromBruto(100#, 0, TEST_TIP_AMB_2KG, neto, tara), _
                "ComputeNetoFromBruto: nula gajbi odbijena"
    AssertFalse ComputeNetoFromBruto(0#, 10, TEST_TIP_AMB_2KG, neto, tara), _
                "ComputeNetoFromBruto: nulti bruto odbijen"
    AssertFalse ComputeNetoFromBruto(100#, 10, "", neto, tara), _
                "ComputeNetoFromBruto: prazan tip odbijen"

    Exit Sub

EH:
    LogFail "ComputeNetoFromBruto", Err.description
End Sub

' modOtkupBlok.LinkOtkupIDsToOtpremnica: parsira "A + B" rezultat i upisuje
' OtpremnicaID na sve redove; prazni/nepostojeci ID-jevi su bezbedan no-op.
Private Sub Test_LinkOtkupIDsToOtpremnica()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKLNK")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim idA As String
    Dim idB As String

    idA = SaveOtkup_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
                       10#, 5#, TEST_TIP_AMB, 0, TEST_VOZ_ID, _
                       TEST_PREFIX & "-LNK-A-" & scenario, 0#, "TEST")
    idB = SaveOtkup_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
                       10#, 5#, TEST_TIP_AMB, 0, TEST_VOZ_ID, _
                       TEST_PREFIX & "-LNK-B-" & scenario, 0#, "TEST")

    AssertTrue Len(idA) > 0 And Len(idB) > 0, "Link: fixture redovi snimljeni"

    Dim otp1 As String
    Dim otp2 As String
    otp1 = "OTP-TSTLNK1-" & scenario
    otp2 = "OTP-TSTLNK2-" & scenario

    LinkOtkupIDsToOtpremnica idA & " + " & idB, otp1

    AssertEquals otp1, CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, idA, COL_OTK_OTPREMNICA_ID)), _
                 "Link: prvi ID iz 'A + B' vezan"
    AssertEquals otp1, CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, idB, COL_OTK_OTPREMNICA_ID)), _
                 "Link: drugi ID iz 'A + B' vezan"

    ' Prazni argumenti = no-op (bez greske, bez promene).
    LinkOtkupIDsToOtpremnica "", otp2
    LinkOtkupIDsToOtpremnica idA, ""

    AssertEquals otp1, CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, idA, COL_OTK_OTPREMNICA_ID)), _
                 "Link: prazan otpID ne menja postojecu vezu"

    ' Nepostojeci ID u listi ne obara upis postojeceg.
    LinkOtkupIDsToOtpremnica "OTK-NEPOSTOJI-" & scenario & " + " & idA, otp2

    AssertEquals otp2, CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, idA, COL_OTK_OTPREMNICA_ID)), _
                 "Link: validan ID vezan i pored nepostojeceg u listi"

    Exit Sub

EH:
    LogFail "LinkOtkupIDsToOtpremnica", Err.description
End Sub

' modOtkupBlok Sum helperi: kolicina / bruto (fallback na neto) / ambalaza
' po otpremnici; stornirani blok iskljucen.
Private Sub Test_SumHelpersByOtp()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKSUM")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim otpID As String
    otpID = "OTP-TSTSUM-" & scenario

    Dim hasBrutoCol As Boolean
    hasBrutoCol = (GetColumnIndex(TBL_OTKUP, COL_OTK_BRUTO) > 0)

    ' r1: sa BrutoKg 120 (neto 100); r2: bez bruto (fallback bruto = neto 50);
    ' r3: storniran (999) - ne sme u sume.
    Dim r1 As String
    If hasBrutoCol Then
        r1 = SaveOtkupMulti_TX( _
            testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
            100#, 10#, TEST_TIP_AMB, 10, TEST_VOZ_ID, _
            TEST_PREFIX & "-SUM1-" & scenario, 0#, "TEST", "", brojZbirne, _
            brutoKgI:=120#)
    Else
        r1 = SaveOtkup_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
                          100#, 10#, TEST_TIP_AMB, 10, TEST_VOZ_ID, _
                          TEST_PREFIX & "-SUM1-" & scenario, 0#, "TEST")
    End If

    Dim r2 As String
    r2 = SaveOtkup_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
                      50#, 10#, TEST_TIP_AMB, 5, TEST_VOZ_ID, _
                      TEST_PREFIX & "-SUM2-" & scenario, 0#, "TEST")

    Dim r3 As String
    r3 = SaveOtkup_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
                      999#, 10#, TEST_TIP_AMB, 9, TEST_VOZ_ID, _
                      TEST_PREFIX & "-SUM3-" & scenario, 0#, "TEST")

    AssertTrue Len(r1) > 0 And Len(r2) > 0 And Len(r3) > 0, "Sum helperi: fixture redovi snimljeni"

    LinkOtkupIDsToOtpremnica r1 & " + " & r2 & " + " & r3, otpID
    MarkTestRowStornirano TBL_OTKUP, COL_OTK_ID, r3

    AssertDoubleNear 150#, SumKolByOtp(otpID), 0.001, _
                     "SumKolByOtp: neto suma bez storniranog"
    AssertDoubleNear 15#, SumAmbByOtp(otpID), 0.001, _
                     "SumAmbByOtp: ambalaza suma bez storniranog"

    If hasBrutoCol Then
        AssertDoubleNear 170#, SumBrutoByOtp(otpID), 0.001, _
                         "SumBrutoByOtp: bruto 120 + fallback neto 50"
    Else
        LogSkip "SumBrutoByOtp", "Nema kolone " & COL_OTK_BRUTO
    End If

    Exit Sub

EH:
    LogFail "Sum helperi po otpremnici", Err.description
End Sub

' modOtkupBlok.PrijemnicaBrojZaZbirnu: dedup po broju (Klasa I+II dele broj),
' vise prijemnica spojeno zarezom, storno preskocen, prazan ulaz -> prazno.
Private Sub Test_PrijemnicaBrojZaZbirnu()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKPRJ")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbirne As String
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim brojP1 As String
    Dim brojP2 As String
    brojP1 = TEST_PREFIX & "-PRJ1-" & scenario
    brojP2 = TEST_PREFIX & "-PRJ2-" & scenario

    Dim p1I As String
    Dim p1II As String
    Dim p2I As String

    p1I = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojP1, brojZbirne, _
                            TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 0, 0, "I")
    p1II = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojP1, brojZbirne, _
                             TEST_VRSTA, TEST_SORTA, 50#, 8#, TEST_TIP_AMB, 0, 0, "II")
    p2I = SavePrijemnica_TX(testDate, TEST_KUP_ID, TEST_VOZ_ID, brojP2, brojZbirne, _
                            TEST_VRSTA, TEST_SORTA, 30#, 10#, TEST_TIP_AMB, 0, 0, "I")

    AssertTrue Len(p1I) > 0 And Len(p1II) > 0 And Len(p2I) > 0, _
               "PrijemnicaBrojZaZbirnu: fixture prijemnice snimljene"

    AssertEquals brojP1 & ", " & brojP2, PrijemnicaBrojZaZbirnu(brojZbirne), _
                 "PrijemnicaBrojZaZbirnu: dedup klasa + spajanje zarezom"

    MarkTestRowStornirano TBL_PRIJEMNICA, "PrijemnicaID", p2I

    AssertEquals brojP1, PrijemnicaBrojZaZbirnu(brojZbirne), _
                 "PrijemnicaBrojZaZbirnu: stornirana prijemnica preskocena"

    AssertEquals "", PrijemnicaBrojZaZbirnu(""), _
                 "PrijemnicaBrojZaZbirnu: prazan BrojZbirne -> prazno"

    Exit Sub

EH:
    LogFail "PrijemnicaBrojZaZbirnu", Err.description
End Sub

' StornoOtkupByBrDok_TX (engine iza "Storniraj blok" u otkup-blok panelu):
' jedan BrDok = obe klase stornirane atomicno; ponovni storno = False.
Private Sub Test_StornoOtkupObeKlase()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKSTB")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojDok As String
    Dim brojZbirne As String
    brojDok = TEST_PREFIX & "-OTK-STB-" & scenario
    brojZbirne = TEST_PREFIX & "-ZBR-" & scenario

    Dim result As String
    result = SaveOtkupMulti_TX( _
        testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
        100#, 10#, TEST_TIP_AMB, 10, TEST_VOZ_ID, brojDok, _
        0#, "TEST OPERATOR", GetTestParcelaID(), brojZbirne, _
        hasKlasaII:=True, kolicinaII:=50#, cenaII:=5#)

    Dim otkI As String
    Dim otkII As String
    otkI = FindOtkupIDByBrojAndKlasa(brojDok, "I")
    otkII = FindOtkupIDByBrojAndKlasa(brojDok, "II")

    AssertTrue Len(otkI) > 0 And Len(otkII) > 0, "Storno obe klase: fixture snimljen"

    AssertTrue StornoOtkupByBrDok_TX(brojDok), "Storno obe klase: storno uspesan"

    AssertEquals "DA", UCase$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkI, COL_OTK_STORNIRANO))), _
                 "Storno obe klase: Klasa I stornirana"
    AssertEquals "DA", UCase$(CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otkII, COL_OTK_STORNIRANO))), _
                 "Storno obe klase: Klasa II stornirana"

    AssertFalse StornoOtkupByBrDok_TX(brojDok), _
                "Storno obe klase: ponovni storno bez aktivnih redova = False"

    Exit Sub

EH:
    LogFail "Storno otkupa obe klase", Err.description
End Sub

' Izgubljeni blokovi + preuzimanje (engine iza otkup-blok panela):
' storno otpremnice -> blok u GetLostOtkupBlokovi; ReassignOtkupToOtpremnica_TX
' re-pointuje vezu (OtpremnicaID + BrojZbirne); storniran cilj se odbija.
Private Sub Test_LostBlokAdoptFlow()
    On Error GoTo EH

    Dim scenario As String
    scenario = NewScenarioCode("OTKLST")

    Dim testDate As Date
    testDate = NextTestDate()

    Dim brojZbrA As String
    Dim brojZbrB As String
    brojZbrA = TEST_PREFIX & "-ZBRA-" & scenario
    brojZbrB = TEST_PREFIX & "-ZBRB-" & scenario

    Dim otpA As String
    otpA = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTPA-" & scenario, brojZbrA, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 0, "I")

    Dim otk As String
    otk = SaveOtkup_TX(testDate, TEST_KOOP_ID, TEST_ST_ID, TEST_VRSTA, TEST_SORTA, _
                       100#, 10#, TEST_TIP_AMB, 0, TEST_VOZ_ID, _
                       TEST_PREFIX & "-LST-" & scenario, 0#, "TEST", KLASA_I, "", brojZbrA)

    AssertTrue Len(otpA) > 0 And Len(otk) > 0, "Lost blok: fixture snimljen"

    AssertTrue ReassignOtkupToOtpremnica_TX(otk, otpA), _
               "Lost blok: vezivanje na aktivnu otpremnicu uspesno"
    AssertEquals otpA, CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otk, COL_OTK_OTPREMNICA_ID)), _
                 "Lost blok: OtpremnicaID postavljen"
    AssertFalse LostBlokoviContains(otk), "Lost blok: uredno vezan blok nije 'izgubljen'"

    AssertTrue StornoOtpremnica_TX(otpA), "Lost blok: storno otpremnice uspesan"
    AssertTrue LostBlokoviContains(otk), _
               "Lost blok: blok stornirane otpremnice u listi izgubljenih"

    Dim otpB As String
    otpB = SaveOtpremnica_TX(testDate, TEST_ST_ID, TEST_VOZ_ID, _
                             TEST_PREFIX & "-OTPB-" & scenario, brojZbrB, _
                             TEST_VRSTA, TEST_SORTA, 100#, 10#, TEST_TIP_AMB, 0, "I")

    AssertTrue Len(otpB) > 0, "Lost blok: ciljna otpremnica B snimljena"

    AssertTrue ReassignOtkupToOtpremnica_TX(otk, otpB), _
               "Lost blok: preuzimanje na novu otpremnicu uspesno"
    AssertEquals otpB, CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otk, COL_OTK_OTPREMNICA_ID)), _
                 "Lost blok: OtpremnicaID re-pointovan na cilj"
    AssertEquals brojZbrB, CStr(GetValueByKey(TBL_OTKUP, COL_OTK_ID, otk, COL_OTK_BROJ_ZBIRNE)), _
                 "Lost blok: BrojZbirne preuzet sa ciljne otpremnice"
    AssertFalse LostBlokoviContains(otk), "Lost blok: posle preuzimanja vise nije 'izgubljen'"

    AssertFalse ReassignOtkupToOtpremnica_TX(otk, otpA), _
                "Lost blok: storniran cilj se odbija"

    Exit Sub

EH:
    LogFail "Lost blok adopt flow", Err.description
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
    SeedVozac
    SeedKupac
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
    created = AutoCreateZbirnaFromOtpremnice()
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
    Call AutoCreateZbirnaFromOtpremnice
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

' Tip ambalaze sa poznatom tezinom (2 kg) za ComputeNetoFromBruto test.
Private Sub SeedTipAmbalaze2kg()
    If GetTable(TBL_TIP_AMBALAZE) Is Nothing Then Exit Sub
    If RowExists(TBL_TIP_AMBALAZE, COL_TAMB_TIP, TEST_TIP_AMB_2KG) Then Exit Sub

    Dim rowData As Variant
    rowData = BlankRow(TBL_TIP_AMBALAZE)

    SetRequiredField rowData, TBL_TIP_AMBALAZE, COL_TAMB_TIP, TEST_TIP_AMB_2KG
    SetRequiredField rowData, TBL_TIP_AMBALAZE, COL_TAMB_TEZINA, 2#

    RequireAppend TBL_TIP_AMBALAZE, rowData, "SeedTipAmbalaze2kg"
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

Private Function TestNumVal(ByVal v As Variant) As Double
    If IsNumeric(v) Then TestNumVal = CDbl(v)
End Function

' Suma kolicine + broj nogu u tblAmbalaza za dati dokument/entitet/smer/tip
' dokumenta (provera dvojnog upisa ambalaze iz otkupa).
Private Function SumAmbalazaLegKolicina(ByVal dokumentID As String, _
                                        ByVal entitetID As String, _
                                        ByVal smer As String, _
                                        ByVal dokTip As String, _
                                        ByRef legCount As Long) As Double
    legCount = 0

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cDok As Long, cEnt As Long, cSmer As Long, cTip As Long, cKol As Long
    cDok = RequireCol(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cEnt = RequireCol(TBL_AMBALAZA, COL_AMB_ENTITET)
    cSmer = RequireCol(TBL_AMBALAZA, COL_AMB_SMER)
    cTip = RequireCol(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cKol = RequireCol(TBL_AMBALAZA, COL_AMB_KOLICINA)

    Dim i As Long
    Dim s As Double
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cDok))) = dokumentID _
           And Trim$(CStr(data(i, cEnt))) = entitetID _
           And Trim$(CStr(data(i, cSmer))) = smer _
           And Trim$(CStr(data(i, cTip))) = dokTip Then
            legCount = legCount + 1
            If IsNumeric(data(i, cKol)) Then s = s + CDbl(data(i, cKol))
        End If
    Next i

    SumAmbalazaLegKolicina = s
End Function

' Vrednost kolone (2=kolicina, 3=novac, 4=ambalaza) iz GetSaldoByStation
' rezultata za datog kooperanta; nema reda -> 0.
Private Function FindSaldoRowValue(ByVal saldo As Variant, _
                                   ByVal koopID As String, _
                                   ByVal colIndex As Long) As Double
    If IsEmpty(saldo) Then Exit Function
    If Not IsArray(saldo) Then Exit Function

    Dim i As Long
    For i = 1 To UBound(saldo, 1)
        If CStr(saldo(i, 1)) = koopID Then
            If IsNumeric(saldo(i, colIndex)) Then
                FindSaldoRowValue = CDbl(saldo(i, colIndex))
            End If
            Exit Function
        End If
    Next i
End Function

' Da li GetLostOtkupBlokovi (modDokumenta) sadrzi dati OtkupID (kolona 1).
Private Function LostBlokoviContains(ByVal otkupID As String) As Boolean
    Dim lost As Variant
    lost = GetLostOtkupBlokovi()
    If Not IsArray(lost) Then Exit Function

    Dim i As Long
    For i = 1 To UBound(lost, 1)
        If Trim$(CStr(lost(i, 1))) = otkupID Then
            LostBlokoviContains = True
            Exit Function
        End If
    Next i
End Function

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

    AppendTestLog "SUITE", suiteName, "START", "RunID=" & m_RunID
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


