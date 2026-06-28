Attribute VB_Name = "modDokumentaTests"
Option Explicit

' ============================================================
' modDokumentaTests
'
' Professional regression suite for the DOKUMENTA core (modDokumenta):
' the Otpremnica -> Zbirna -> Prijemnica document chain.
'
' Covers:
'   - SaveOtpremnica_TX / SaveOtpremnicaMulti_TX   (Klasa I, I+II)
'   - SaveZbirna_TX / SaveZbirnaMulti_TX
'   - SavePrijemnica_TX / SavePrijemnicaMulti_TX
'   - Validation hardening (negative cena, missing amb type, bad class,
'     negative ambalaza) -> empty result + no append
'   - GetOtpremniceByZbirna / GetOtpremniceByStation
'   - GetZbirnaByKupac / GetPrijemniceByKupac (storno exclusion + date)
'   - ValidateZbirna (otpremnica vs zbirna kg reconciliation)
'   - CalculateManjak (zbirna vs prijemnica shrinkage)
'   - Storno exclusion (StornoOtpremnicaByBroj_TX / StornoZbirna_TX /
'     StornoPrijemnicaByBroj_TX)
'   - Orphan/recovery read helpers smoke
'
' Design (matches modFakturaTests / modOtkupTests house style):
'   - Self-contained: own counters, own log sheet, own assert/log
'     helpers prefixed "Dok" to avoid Ambiguous-name on compile.
'   - Standalone runnable on a BLANK workbook: seeds its own master
'     data in an isolated ID namespace (ST-DOK-*, KUP-DOK-*, ...).
'   - Tests create TST-DOK-* rows; cleanup is soft-storno.
'
' Run from Alt+F8:
'   RunDokumentaSuite             -> full suite
'   RunDokumentaSeedOnly          -> only seed master data
'   SoftStornoDokumentaTestRows   -> soft-storno all TST-DOK-* rows
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long
Private m_RunID As String
Private m_DateSeq As Long

Private Const DOK_TEST_LOG As String = "DOKUMENTA_TEST_LOG"

' Isolated master-data namespace.
Private Const DK_ST_ID As String = "ST-DOK-90001"
Private Const DK_KOOP_ID As String = "KOOP-DOK-90001"
Private Const DK_VOZ_ID As String = "VOZ-DOK-90001"
Private Const DK_KUP_ID As String = "KUP-DOK-90001"
Private Const DK_KUL_ID As String = "KUL-DOK-90001"

Private Const DK_VRSTA As String = "DOK Test Jabuka"
Private Const DK_SORTA As String = "DOK Test Sorta"
Private Const DK_TIP_AMB As String = "DOK Test Gajba"
Private Const DK_HLAD As String = "DOK Test Hladnjaca"
Private Const DK_POGON As String = "DOK Test Pogon"

Private Const DOC_PREFIX As String = "TST-DOK"

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub RunDokumentaSuite()
    On Error GoTo EH

    BeginDokRun "DOKUMENTA CHAIN SUITE"

    Test_DokSchemaColumns
    SeedDokMaster
    Test_DokSeedAvailable

    ' ---- Otpremnica ----
    Test_OtpremnicaHappyPath
    Test_OtpremnicaMultiBothClasses
    Test_OtpremnicaInvalidNegativeCena
    Test_OtpremnicaInvalidMissingAmbType
    Test_GetOtpremniceByZbirna
    Test_GetOtpremniceByStationExcludesStornirano

    ' ---- Zbirna ----
    Test_ZbirnaHappyPath
    Test_ZbirnaMultiBothClasses
    Test_ZbirnaInvalidClass
    Test_GetZbirnaByKupacExcludesStornirano
    Test_ValidateZbirnaReconciliation

    ' ---- Prijemnica ----
    Test_PrijemnicaHappyPath
    Test_PrijemnicaMultiBothClasses
    Test_PrijemnicaInvalidNegativeAmbalaza
    Test_GetPrijemniceByKupacExcludesStornirano

    ' ---- chain + calculations ----
    Test_FullChainLinkByZbirna
    Test_CalculateManjakShrinkage
    Test_CalculateManjakNoShrinkage

    ' ---- storno cascade exclusion ----
    Test_StornoOtpremnicaExcludesFromReads
    Test_StornoZbirnaExcludesFromReads
    Test_StornoPrijemnicaExcludesFromReads

    ' ---- orphan/recovery helpers ----
    Test_OrphanHelpersDoNotCrash

    EndDokRun
    Exit Sub

EH:
    LogDokFatal "RunDokumentaSuite", Err.Number, Err.description
    EndDokRun
End Sub

Public Sub RunDokumentaSeedOnly()
    On Error GoTo EH

    BeginDokRun "DOKUMENTA SEED ONLY"

    Test_DokSchemaColumns
    SeedDokMaster
    Test_DokSeedAvailable

    EndDokRun
    Exit Sub

EH:
    LogDokFatal "RunDokumentaSeedOnly", Err.Number, Err.description
    EndDokRun
End Sub

' ============================================================
' SCHEMA + SEED
' ============================================================

Private Sub Test_DokSchemaColumns()
    On Error GoTo EH

    AssertDokColumns TBL_OTPREMNICA, Array(COL_OTP_ID, COL_OTP_DATUM, COL_OTP_STANICA, _
        COL_OTP_VOZAC, COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE, COL_OTP_VRSTA, COL_OTP_SORTA, _
        COL_OTP_KOLICINA, COL_OTP_CENA, COL_OTP_TIP_AMB, COL_OTP_KOL_AMB, COL_OTP_KLASA)

    AssertDokColumns TBL_ZBIRNA, Array(COL_ZBR_ID, COL_ZBR_DATUM, COL_ZBR_VOZAC, COL_ZBR_BROJ, _
        COL_ZBR_KUPAC, COL_ZBR_KOLICINA, COL_ZBR_TIP_AMB, COL_ZBR_KOL_AMB, COL_ZBR_KLASA)

    AssertDokColumns TBL_PRIJEMNICA, Array(COL_PRJ_ID, COL_PRJ_DATUM, COL_PRJ_KUPAC, COL_PRJ_BROJ, _
        COL_PRJ_BROJ_ZBIRNE, COL_PRJ_KOLICINA, COL_PRJ_CENA, COL_PRJ_TIP_AMB, COL_PRJ_KOL_AMB, _
        COL_PRJ_KOL_AMB_VRACENA, COL_PRJ_KLASA, COL_PRJ_FAKTURISANO)
    Exit Sub

EH:
    LogDokFatal "Test_DokSchemaColumns", Err.Number, Err.description
End Sub

Private Sub AssertDokColumns(ByVal tableName As String, ByVal cols As Variant)
    If GetTable(tableName) Is Nothing Then
        LogDokFail "Schema: " & tableName & " exists", "Table missing"
        Exit Sub
    End If

    Dim c As Variant, missing As String
    For Each c In cols
        If GetColumnIndex(tableName, CStr(c)) = 0 Then missing = missing & CStr(c) & " "
    Next c

    AssertDokTrue (missing = ""), "Schema: " & tableName & " has all required columns"
    If missing <> "" Then LogDokInfo "Missing " & tableName & " columns: " & missing
End Sub

Private Sub Test_DokSeedAvailable()
    On Error GoTo EH
    AssertDokTrue DokRowExists(TBL_STANICE, "StanicaID", DK_ST_ID), "Seed: stanica available"
    AssertDokTrue DokRowExists(TBL_VOZACI, "VozacID", DK_VOZ_ID), "Seed: vozac available"
    AssertDokTrue DokRowExists(TBL_KUPCI, "KupacID", DK_KUP_ID), "Seed: kupac available"
    AssertDokTrue DokRowExists(TBL_KULTURE, "KulturaID", DK_KUL_ID), "Seed: kultura available"
    Exit Sub
EH:
    LogDokFatal "Test_DokSeedAvailable", Err.Number, Err.description
End Sub

Private Sub SeedDokMaster()
    On Error GoTo EH
    SeedDokStanica
    SeedDokVozac
    SeedDokKupac
    SeedDokKultura
    SeedDokKooperant
    LogDokPass "Seed master data ready"
    Exit Sub
EH:
    LogDokFail "Seed master data", Err.description
End Sub

Private Sub SeedDokStanica()
    If DokRowExists(TBL_STANICE, "StanicaID", DK_ST_ID) Then Exit Sub
    Dim r As Variant
    r = DokBlankRow(TBL_STANICE)
    DokSetReq r, TBL_STANICE, "StanicaID", DK_ST_ID
    DokSetReq r, TBL_STANICE, "Naziv", "DOK TEST STANICA"
    DokSetOpt r, TBL_STANICE, "Aktivan", "Aktivan"
    DokAppend TBL_STANICE, r, "SeedDokStanica"
End Sub

Private Sub SeedDokVozac()
    If DokRowExists(TBL_VOZACI, "VozacID", DK_VOZ_ID) Then Exit Sub
    Dim r As Variant
    r = DokBlankRow(TBL_VOZACI)
    DokSetReq r, TBL_VOZACI, "VozacID", DK_VOZ_ID
    DokSetReq r, TBL_VOZACI, "Ime", "DOK"
    DokSetReq r, TBL_VOZACI, "Prezime", "Vozac"
    DokSetOpt r, TBL_VOZACI, "Aktivan", "Aktivan"
    DokAppend TBL_VOZACI, r, "SeedDokVozac"
End Sub

Private Sub SeedDokKupac()
    If DokRowExists(TBL_KUPCI, "KupacID", DK_KUP_ID) Then Exit Sub
    Dim r As Variant
    r = DokBlankRow(TBL_KUPCI)
    DokSetReq r, TBL_KUPCI, "KupacID", DK_KUP_ID
    DokSetReq r, TBL_KUPCI, "Naziv", "DOK TEST KUPAC DOO"
    DokSetReq r, TBL_KUPCI, "PIB", "109000777"
    DokSetOpt r, TBL_KUPCI, "Mesto", "Test Grad"
    DokSetOpt r, TBL_KUPCI, "Aktivan", "Aktivan"
    DokAppend TBL_KUPCI, r, "SeedDokKupac"
End Sub

Private Sub SeedDokKultura()
    If DokRowExists(TBL_KULTURE, "KulturaID", DK_KUL_ID) Then Exit Sub
    Dim r As Variant
    r = DokBlankRow(TBL_KULTURE)
    DokSetReq r, TBL_KULTURE, "KulturaID", DK_KUL_ID
    DokSetReq r, TBL_KULTURE, "VrstaVoca", DK_VRSTA
    DokSetReq r, TBL_KULTURE, "SortaVoca", DK_SORTA
    DokSetOpt r, TBL_KULTURE, "Aktivan", "Aktivan"
    DokAppend TBL_KULTURE, r, "SeedDokKultura"
End Sub

Private Sub SeedDokKooperant()
    If DokRowExists(TBL_KOOPERANTI, COL_KOOP_ID, DK_KOOP_ID) Then Exit Sub
    Dim r As Variant
    r = DokBlankRow(TBL_KOOPERANTI)
    DokSetReq r, TBL_KOOPERANTI, "KooperantID", DK_KOOP_ID
    DokSetReq r, TBL_KOOPERANTI, "Ime", "DOK"
    DokSetReq r, TBL_KOOPERANTI, "Prezime", "Kooperant"
    DokSetReq r, TBL_KOOPERANTI, "StanicaID", DK_ST_ID
    DokSetOpt r, TBL_KOOPERANTI, "Aktivan", "Da"
    DokAppend TBL_KOOPERANTI, r, "SeedDokKooperant"
End Sub

' ============================================================
' OTPREMNICA
' ============================================================

Private Sub Test_OtpremnicaHappyPath()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("OTPHAPPY")
    Dim brojOtp As String: brojOtp = DOC_PREFIX & "-OTP-" & sc
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    Dim id As String
    id = SaveOtpremnica_TX(NextDokDate(), DK_ST_ID, DK_VOZ_ID, brojOtp, brojZbr, _
                           DK_VRSTA, DK_SORTA, 500#, 100#, DK_TIP_AMB, 25, KLASA_I)

    AssertDokTrue (Left$(id, 4) = "OTP-"), "Otpremnica: returns OTP- id"
    AssertDokEquals DK_ST_ID, DokLookup(TBL_OTPREMNICA, COL_OTP_ID, id, COL_OTP_STANICA), "Otpremnica: stanica persisted"
    AssertDokEquals brojZbr, DokLookup(TBL_OTPREMNICA, COL_OTP_ID, id, COL_OTP_BROJ_ZBIRNE), "Otpremnica: BrojZbirne persisted"
    AssertDokEquals "500", DokLookup(TBL_OTPREMNICA, COL_OTP_ID, id, COL_OTP_KOLICINA), "Otpremnica: kolicina persisted"
    Exit Sub

EH:
    LogDokFatal "Test_OtpremnicaHappyPath", Err.Number, Err.description
End Sub

Private Sub Test_OtpremnicaMultiBothClasses()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("OTPM12")
    Dim brojOtp As String: brojOtp = DOC_PREFIX & "-OTP-" & sc
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    Dim res As String
    res = SaveOtpremnicaMulti_TX(NextDokDate(), DK_ST_ID, DK_VOZ_ID, brojOtp, brojZbr, _
                                 DK_VRSTA, DK_SORTA, 400#, 100#, DK_TIP_AMB, 20, _
                                 True, 150#, 60#, 0#, 8, 0#)

    Dim idI As String, idII As String
    idI = DokFindIDByBrojKlasa(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ, brojOtp, COL_OTP_KLASA, KLASA_I)
    idII = DokFindIDByBrojKlasa(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ, brojOtp, COL_OTP_KLASA, KLASA_II)

    AssertDokTrue (Len(idI) > 0 And Len(idII) > 0), "Otpremnica I+II: both class rows present"
    AssertDokTrue (idI <> idII), "Otpremnica I+II: distinct OtpremnicaID per class"
    AssertDokEquals "2", CStr(DokCountForBroj(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp)), _
                    "Otpremnica I+II: shares BrojOtpremnice across two rows"
    Exit Sub

EH:
    LogDokFatal "Test_OtpremnicaMultiBothClasses", Err.Number, Err.description
End Sub

Private Sub Test_OtpremnicaInvalidNegativeCena()
    On Error GoTo EH

    Dim before As Long: before = DokCountRows(TBL_OTPREMNICA)
    Dim sc As String: sc = NewDokScenario("OTPNEGC")

    Dim res As String
    res = SaveOtpremnica_TX(NextDokDate(), DK_ST_ID, DK_VOZ_ID, _
                            DOC_PREFIX & "-OTP-" & sc, DOC_PREFIX & "-ZBR-" & sc, _
                            DK_VRSTA, DK_SORTA, 100#, -1#, DK_TIP_AMB, 1, KLASA_I)

    AssertDokEquals "", res, "Otpremnica: negative cena returns empty"
    AssertDokEquals CStr(before), CStr(DokCountRows(TBL_OTPREMNICA)), "Otpremnica: negative cena did not append"
    Exit Sub

EH:
    LogDokFatal "Test_OtpremnicaInvalidNegativeCena", Err.Number, Err.description
End Sub

Private Sub Test_OtpremnicaInvalidMissingAmbType()
    On Error GoTo EH

    Dim beforeOtp As Long: beforeOtp = DokCountRows(TBL_OTPREMNICA)
    Dim beforeAmb As Long: beforeAmb = DokCountRows(TBL_AMBALAZA)
    Dim sc As String: sc = NewDokScenario("OTPNOTIP")

    Dim res As String
    res = SaveOtpremnica_TX(NextDokDate(), DK_ST_ID, DK_VOZ_ID, _
                            DOC_PREFIX & "-OTP-" & sc, DOC_PREFIX & "-ZBR-" & sc, _
                            DK_VRSTA, DK_SORTA, 100#, 10#, "", 1, KLASA_I)

    AssertDokEquals "", res, "Otpremnica: missing amb type returns empty"
    AssertDokEquals CStr(beforeOtp), CStr(DokCountRows(TBL_OTPREMNICA)), "Otpremnica: missing amb type did not append otpremnica"
    AssertDokEquals CStr(beforeAmb), CStr(DokCountRows(TBL_AMBALAZA)), "Otpremnica: missing amb type did not append ambalaza"
    Exit Sub

EH:
    LogDokFatal "Test_OtpremnicaInvalidMissingAmbType", Err.Number, Err.description
End Sub

Private Sub Test_GetOtpremniceByZbirna()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("GETOTPZ")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    SaveOtpremnica_TX NextDokDate(), DK_ST_ID, DK_VOZ_ID, DOC_PREFIX & "-OTP-A-" & sc, brojZbr, _
                      DK_VRSTA, DK_SORTA, 100#, 10#, DK_TIP_AMB, 5, KLASA_I
    SaveOtpremnica_TX NextDokDate(), DK_ST_ID, DK_VOZ_ID, DOC_PREFIX & "-OTP-B-" & sc, brojZbr, _
                      DK_VRSTA, DK_SORTA, 200#, 10#, DK_TIP_AMB, 5, KLASA_I

    AssertDokEquals "2", CStr(RowsOf(GetOtpremniceByZbirna(brojZbr))), _
                    "GetOtpremniceByZbirna: returns all otpremnice for BrojZbirne"
    Exit Sub

EH:
    LogDokFatal "Test_GetOtpremniceByZbirna", Err.Number, Err.description
End Sub

Private Sub Test_GetOtpremniceByStationExcludesStornirano()
    On Error GoTo EH

    Dim stId As String: stId = "ST-DOK-READ1"
    EnsureDokStanica stId, "DOK READ STANICA"

    Dim sc As String: sc = NewDokScenario("OTPSTA")
    Dim d As Date: d = DateSerial(2092, 5, 10)
    Dim brA As String: brA = DOC_PREFIX & "-OTP-A-" & sc

    SaveOtpremnica_TX d, stId, DK_VOZ_ID, brA, DOC_PREFIX & "-ZBR-A-" & sc, _
                      DK_VRSTA, DK_SORTA, 100#, 10#, DK_TIP_AMB, 5, KLASA_I
    SaveOtpremnica_TX d, stId, DK_VOZ_ID, DOC_PREFIX & "-OTP-B-" & sc, DOC_PREFIX & "-ZBR-B-" & sc, _
                      DK_VRSTA, DK_SORTA, 200#, 10#, DK_TIP_AMB, 5, KLASA_I

    AssertDokEquals "2", CStr(RowsOf(GetOtpremniceByStation(stId, d, d))), "GetOtpremniceByStation: both active rows in range"

    StornoOtpremnicaByBroj_TX brA
    AssertDokEquals "1", CStr(RowsOf(GetOtpremniceByStation(stId, d, d))), "GetOtpremniceByStation: excludes stornirano"
    Exit Sub

EH:
    LogDokFatal "Test_GetOtpremniceByStationExcludesStornirano", Err.Number, Err.description
End Sub

' ============================================================
' ZBIRNA
' ============================================================

Private Sub Test_ZbirnaHappyPath()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("ZBRHAPPY")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    Dim id As String
    id = SaveZbirna_TX(NextDokDate(), DK_VOZ_ID, brojZbr, DK_KUP_ID, DK_HLAD, DK_POGON, _
                       DK_VRSTA, DK_SORTA, 500#, DK_TIP_AMB, 25, KLASA_I)

    AssertDokTrue (Left$(id, 4) = "ZBR-"), "Zbirna: returns ZBR- id"
    AssertDokEquals DK_KUP_ID, DokLookup(TBL_ZBIRNA, COL_ZBR_ID, id, COL_ZBR_KUPAC), "Zbirna: kupac persisted"
    AssertDokEquals "500", DokLookup(TBL_ZBIRNA, COL_ZBR_ID, id, COL_ZBR_KOLICINA), "Zbirna: UkupnoKolicina persisted"
    Exit Sub

EH:
    LogDokFatal "Test_ZbirnaHappyPath", Err.Number, Err.description
End Sub

Private Sub Test_ZbirnaMultiBothClasses()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("ZBRM12")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    Dim res As String
    res = SaveZbirnaMulti_TX(NextDokDate(), DK_VOZ_ID, brojZbr, DK_KUP_ID, DK_HLAD, DK_POGON, _
                             DK_VRSTA, DK_SORTA, 400#, DK_TIP_AMB, 20, _
                             True, 150#, 8)

    Dim idI As String, idII As String
    idI = DokFindIDByBrojKlasa(TBL_ZBIRNA, COL_ZBR_ID, COL_ZBR_BROJ, brojZbr, COL_ZBR_KLASA, KLASA_I)
    idII = DokFindIDByBrojKlasa(TBL_ZBIRNA, COL_ZBR_ID, COL_ZBR_BROJ, brojZbr, COL_ZBR_KLASA, KLASA_II)

    AssertDokTrue (Len(idI) > 0 And Len(idII) > 0), "Zbirna I+II: both class rows present"
    AssertDokTrue (idI <> idII), "Zbirna I+II: distinct ZbirnaID per class"
    Exit Sub

EH:
    LogDokFatal "Test_ZbirnaMultiBothClasses", Err.Number, Err.description
End Sub

Private Sub Test_ZbirnaInvalidClass()
    On Error GoTo EH

    Dim before As Long: before = DokCountRows(TBL_ZBIRNA)
    Dim sc As String: sc = NewDokScenario("ZBRBAD")

    Dim res As String
    res = SaveZbirna_TX(NextDokDate(), DK_VOZ_ID, DOC_PREFIX & "-ZBR-" & sc, DK_KUP_ID, _
                        DK_HLAD, DK_POGON, DK_VRSTA, DK_SORTA, 100#, DK_TIP_AMB, 1, "BAD")

    AssertDokEquals "", res, "Zbirna: invalid class returns empty"
    AssertDokEquals CStr(before), CStr(DokCountRows(TBL_ZBIRNA)), "Zbirna: invalid class did not append"
    Exit Sub

EH:
    LogDokFatal "Test_ZbirnaInvalidClass", Err.Number, Err.description
End Sub

Private Sub Test_GetZbirnaByKupacExcludesStornirano()
    On Error GoTo EH

    Dim kupId As String: kupId = "KUP-DOK-READ1"
    EnsureDokKupac kupId, "DOK READ KUPAC"

    Dim sc As String: sc = NewDokScenario("ZBRKUP")
    Dim d As Date: d = DateSerial(2092, 7, 12)
    Dim brZbrA As String: brZbrA = DOC_PREFIX & "-ZBR-A-" & sc

    SaveZbirna_TX d, DK_VOZ_ID, brZbrA, kupId, DK_HLAD, DK_POGON, DK_VRSTA, DK_SORTA, 100#, DK_TIP_AMB, 5, KLASA_I
    SaveZbirna_TX d, DK_VOZ_ID, DOC_PREFIX & "-ZBR-B-" & sc, kupId, DK_HLAD, DK_POGON, DK_VRSTA, DK_SORTA, 200#, DK_TIP_AMB, 5, KLASA_I

    AssertDokEquals "2", CStr(RowsOf(GetZbirnaByKupac(kupId, d, d))), "GetZbirnaByKupac: both active rows in range"

    StornoZbirna_TX brZbrA
    AssertDokEquals "1", CStr(RowsOf(GetZbirnaByKupac(kupId, d, d))), "GetZbirnaByKupac: excludes stornirano"
    Exit Sub

EH:
    LogDokFatal "Test_GetZbirnaByKupacExcludesStornirano", Err.Number, Err.description
End Sub

Private Sub Test_ValidateZbirnaReconciliation()
    On Error GoTo EH

    ' One otpremnica (100kg) + matching zbirna (100kg) under same BrojZbirne
    ' -> ValidateZbirna reports equal sums and ValidKg = True.
    Dim sc As String: sc = NewDokScenario("ZBRVAL")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    SaveOtpremnica_TX NextDokDate(), DK_ST_ID, DK_VOZ_ID, DOC_PREFIX & "-OTP-" & sc, brojZbr, _
                      DK_VRSTA, DK_SORTA, 100#, 10#, DK_TIP_AMB, 5, KLASA_I
    SaveZbirna_TX NextDokDate(), DK_VOZ_ID, brojZbr, DK_KUP_ID, DK_HLAD, DK_POGON, _
                  DK_VRSTA, DK_SORTA, 100#, DK_TIP_AMB, 5, KLASA_I

    Dim v As Variant
    v = ValidateZbirna(brojZbr)
    ' Array(SumaOtpKg, ZbirnaKg, RazlikaKg, ValidKg, SumaOtpAmb, ZbirnaAmb, RazlikaAmb)
    AssertDokEquals "100", CStr(v(0)), "ValidateZbirna: suma otpremnica kg = 100"
    AssertDokEquals "100", CStr(v(1)), "ValidateZbirna: zbirna kg = 100"
    AssertDokTrue CBool(v(3)), "ValidateZbirna: ValidKg True when sums match"
    Exit Sub

EH:
    LogDokFatal "Test_ValidateZbirnaReconciliation", Err.Number, Err.description
End Sub

' ============================================================
' PRIJEMNICA
' ============================================================

Private Sub Test_PrijemnicaHappyPath()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("PRJHAPPY")
    Dim brojPrj As String: brojPrj = DOC_PREFIX & "-PRJ-" & sc
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    Dim id As String
    id = SavePrijemnica_TX(NextDokDate(), DK_KUP_ID, DK_VOZ_ID, brojPrj, brojZbr, _
                           DK_VRSTA, DK_SORTA, 480#, 100#, DK_TIP_AMB, 24, 0, KLASA_I)

    AssertDokTrue (Left$(id, 4) = "PRJ-"), "Prijemnica: returns PRJ- id"
    AssertDokEquals DK_KUP_ID, DokLookup(TBL_PRIJEMNICA, COL_PRJ_ID, id, COL_PRJ_KUPAC), "Prijemnica: kupac persisted"
    AssertDokEquals "480", DokLookup(TBL_PRIJEMNICA, COL_PRJ_ID, id, COL_PRJ_KOLICINA), "Prijemnica: kolicina persisted"
    AssertDokTrue (CStr(DokLookup(TBL_PRIJEMNICA, COL_PRJ_ID, id, COL_PRJ_FAKTURISANO)) <> "Da"), _
                  "Prijemnica: not flagged Fakturisano on create"
    Exit Sub

EH:
    LogDokFatal "Test_PrijemnicaHappyPath", Err.Number, Err.description
End Sub

Private Sub Test_PrijemnicaMultiBothClasses()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("PRJM12")
    Dim brojPrj As String: brojPrj = DOC_PREFIX & "-PRJ-" & sc
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    Dim res As String
    res = SavePrijemnicaMulti_TX(NextDokDate(), DK_KUP_ID, DK_VOZ_ID, brojPrj, brojZbr, _
                                 DK_VRSTA, DK_SORTA, 400#, 100#, DK_TIP_AMB, 20, 0, _
                                 True, 150#, 60#, 8, 0#, 0#)

    Dim idI As String, idII As String
    idI = DokFindIDByBrojKlasa(TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_BROJ, brojPrj, COL_PRJ_KLASA, KLASA_I)
    idII = DokFindIDByBrojKlasa(TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_BROJ, brojPrj, COL_PRJ_KLASA, KLASA_II)

    AssertDokTrue (Len(idI) > 0 And Len(idII) > 0), "Prijemnica I+II: both class rows present"
    AssertDokTrue (idI <> idII), "Prijemnica I+II: distinct PrijemnicaID per class"
    Exit Sub

EH:
    LogDokFatal "Test_PrijemnicaMultiBothClasses", Err.Number, Err.description
End Sub

Private Sub Test_PrijemnicaInvalidNegativeAmbalaza()
    On Error GoTo EH

    Dim beforePrj As Long: beforePrj = DokCountRows(TBL_PRIJEMNICA)
    Dim beforeAmb As Long: beforeAmb = DokCountRows(TBL_AMBALAZA)
    Dim sc As String: sc = NewDokScenario("PRJNEG")

    Dim res As String
    res = SavePrijemnica_TX(NextDokDate(), DK_KUP_ID, DK_VOZ_ID, _
                            DOC_PREFIX & "-PRJ-" & sc, DOC_PREFIX & "-ZBR-" & sc, _
                            DK_VRSTA, DK_SORTA, 100#, 10#, DK_TIP_AMB, -1, 0, KLASA_I)

    AssertDokEquals "", res, "Prijemnica: negative ambalaza returns empty"
    AssertDokEquals CStr(beforePrj), CStr(DokCountRows(TBL_PRIJEMNICA)), "Prijemnica: negative ambalaza did not append prijemnica"
    AssertDokEquals CStr(beforeAmb), CStr(DokCountRows(TBL_AMBALAZA)), "Prijemnica: negative ambalaza did not append ambalaza"
    Exit Sub

EH:
    LogDokFatal "Test_PrijemnicaInvalidNegativeAmbalaza", Err.Number, Err.description
End Sub

Private Sub Test_GetPrijemniceByKupacExcludesStornirano()
    On Error GoTo EH

    Dim kupId As String: kupId = "KUP-DOK-READ2"
    EnsureDokKupac kupId, "DOK READ KUPAC 2"

    Dim sc As String: sc = NewDokScenario("PRJKUP")
    Dim d As Date: d = DateSerial(2092, 8, 14)
    Dim brPrjA As String: brPrjA = DOC_PREFIX & "-PRJ-A-" & sc

    SavePrijemnica_TX d, kupId, DK_VOZ_ID, brPrjA, DOC_PREFIX & "-ZBR-A-" & sc, _
                      DK_VRSTA, DK_SORTA, 100#, 10#, DK_TIP_AMB, 5, 0, KLASA_I
    SavePrijemnica_TX d, kupId, DK_VOZ_ID, DOC_PREFIX & "-PRJ-B-" & sc, DOC_PREFIX & "-ZBR-B-" & sc, _
                      DK_VRSTA, DK_SORTA, 200#, 10#, DK_TIP_AMB, 5, 0, KLASA_I

    AssertDokEquals "2", CStr(RowsOf(GetPrijemniceByKupac(kupId, d, d))), "GetPrijemniceByKupac: both active rows in range"

    StornoPrijemnicaByBroj_TX brPrjA
    AssertDokEquals "1", CStr(RowsOf(GetPrijemniceByKupac(kupId, d, d))), "GetPrijemniceByKupac: excludes stornirano"
    Exit Sub

EH:
    LogDokFatal "Test_GetPrijemniceByKupacExcludesStornirano", Err.Number, Err.description
End Sub

' ============================================================
' CHAIN + CALCULATIONS
' ============================================================

Private Sub Test_FullChainLinkByZbirna()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("CHAIN")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    Dim otp As String, zbr As String, prj As String
    otp = SaveOtpremnica_TX(NextDokDate(), DK_ST_ID, DK_VOZ_ID, DOC_PREFIX & "-OTP-" & sc, brojZbr, _
                            DK_VRSTA, DK_SORTA, 200#, 100#, DK_TIP_AMB, 10, KLASA_I)
    zbr = SaveZbirna_TX(NextDokDate(), DK_VOZ_ID, brojZbr, DK_KUP_ID, DK_HLAD, DK_POGON, _
                        DK_VRSTA, DK_SORTA, 200#, DK_TIP_AMB, 10, KLASA_I)
    prj = SavePrijemnica_TX(NextDokDate(), DK_KUP_ID, DK_VOZ_ID, DOC_PREFIX & "-PRJ-" & sc, brojZbr, _
                            DK_VRSTA, DK_SORTA, 190#, 100#, DK_TIP_AMB, 10, 0, KLASA_I)

    AssertDokTrue (Len(otp) > 0 And Len(zbr) > 0 And Len(prj) > 0), "Chain: otpremnica+zbirna+prijemnica created"
    AssertDokEquals "1", CStr(RowsOf(GetOtpremniceByZbirna(brojZbr))), "Chain: otpremnica reachable by BrojZbirne"
    AssertDokTrue (RowsOf(GetZbirnaByKupac(DK_KUP_ID)) >= 1), "Chain: zbirna reachable by kupac"
    Exit Sub

EH:
    LogDokFatal "Test_FullChainLinkByZbirna", Err.Number, Err.description
End Sub

Private Sub Test_CalculateManjakShrinkage()
    On Error GoTo EH

    ' Zbirna 200kg vs Prijemnica 150kg -> manjak 50kg / 25%.
    Dim sc As String: sc = NewDokScenario("MANJAK")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    SaveZbirna_TX NextDokDate(), DK_VOZ_ID, brojZbr, DK_KUP_ID, DK_HLAD, DK_POGON, _
                  DK_VRSTA, DK_SORTA, 200#, DK_TIP_AMB, 10, KLASA_I
    SavePrijemnica_TX NextDokDate(), DK_KUP_ID, DK_VOZ_ID, DOC_PREFIX & "-PRJ-" & sc, brojZbr, _
                      DK_VRSTA, DK_SORTA, 150#, 100#, DK_TIP_AMB, 10, 0, KLASA_I

    Dim m As Variant
    m = CalculateManjak(brojZbr)
    ' Array(zbirnaKg, prijKg, manjakKg, manjakPct)
    AssertDokEquals "200", CStr(m(0)), "CalculateManjak: zbirna kg = 200"
    AssertDokEquals "150", CStr(m(1)), "CalculateManjak: prijemnica kg = 150"
    AssertDokEquals "50", CStr(m(2)), "CalculateManjak: manjak kg = 50"
    AssertDokTrue (Abs(CDbl(m(3)) - 25#) < 0.01), "CalculateManjak: manjak pct = 25"
    Exit Sub

EH:
    LogDokFatal "Test_CalculateManjakShrinkage", Err.Number, Err.description
End Sub

Private Sub Test_CalculateManjakNoShrinkage()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("NOMANJAK")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    SaveZbirna_TX NextDokDate(), DK_VOZ_ID, brojZbr, DK_KUP_ID, DK_HLAD, DK_POGON, _
                  DK_VRSTA, DK_SORTA, 120#, DK_TIP_AMB, 6, KLASA_I
    SavePrijemnica_TX NextDokDate(), DK_KUP_ID, DK_VOZ_ID, DOC_PREFIX & "-PRJ-" & sc, brojZbr, _
                      DK_VRSTA, DK_SORTA, 120#, 100#, DK_TIP_AMB, 6, 0, KLASA_I

    Dim m As Variant
    m = CalculateManjak(brojZbr)
    AssertDokEquals "0", CStr(m(2)), "CalculateManjak: zero manjak when prijemnica = zbirna"
    Exit Sub

EH:
    LogDokFatal "Test_CalculateManjakNoShrinkage", Err.Number, Err.description
End Sub

' ============================================================
' STORNO CASCADE EXCLUSION
' ============================================================

Private Sub Test_StornoOtpremnicaExcludesFromReads()
    On Error GoTo EH

    Dim sc As String: sc = NewDokScenario("STOOTP")
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc
    Dim brojOtp As String: brojOtp = DOC_PREFIX & "-OTP-" & sc

    Dim id As String
    id = SaveOtpremnica_TX(NextDokDate(), DK_ST_ID, DK_VOZ_ID, brojOtp, brojZbr, _
                           DK_VRSTA, DK_SORTA, 100#, 10#, DK_TIP_AMB, 5, KLASA_I)

    AssertDokEquals "1", CStr(RowsOf(GetOtpremniceByZbirna(brojZbr))), "Storno otpremnica: visible before storno"
    AssertDokTrue StornoOtpremnicaByBroj_TX(brojOtp), "Storno otpremnica: StornoOtpremnicaByBroj_TX returns True"
    AssertDokEquals "Da", CStr(DokLookup(TBL_OTPREMNICA, COL_OTP_ID, id, COL_STORNIRANO)), "Storno otpremnica: flagged Stornirano=Da"
    AssertDokEquals "0", CStr(RowsOf(GetOtpremniceByZbirna(brojZbr))), "Storno otpremnica: hidden from read after storno"
    Exit Sub

EH:
    LogDokFatal "Test_StornoOtpremnicaExcludesFromReads", Err.Number, Err.description
End Sub

Private Sub Test_StornoZbirnaExcludesFromReads()
    On Error GoTo EH

    Dim kupId As String: kupId = "KUP-DOK-STO1"
    EnsureDokKupac kupId, "DOK STORNO KUPAC"

    Dim sc As String: sc = NewDokScenario("STOZBR")
    Dim d As Date: d = DateSerial(2092, 10, 5)
    Dim brojZbr As String: brojZbr = DOC_PREFIX & "-ZBR-" & sc

    SaveZbirna_TX d, DK_VOZ_ID, brojZbr, kupId, DK_HLAD, DK_POGON, DK_VRSTA, DK_SORTA, 100#, DK_TIP_AMB, 5, KLASA_I

    AssertDokEquals "1", CStr(RowsOf(GetZbirnaByKupac(kupId, d, d))), "Storno zbirna: visible before storno"
    AssertDokTrue StornoZbirna_TX(brojZbr), "Storno zbirna: StornoZbirna_TX returns True"
    AssertDokEquals "0", CStr(RowsOf(GetZbirnaByKupac(kupId, d, d))), "Storno zbirna: hidden from read after storno"
    Exit Sub

EH:
    LogDokFatal "Test_StornoZbirnaExcludesFromReads", Err.Number, Err.description
End Sub

Private Sub Test_StornoPrijemnicaExcludesFromReads()
    On Error GoTo EH

    Dim kupId As String: kupId = "KUP-DOK-STO2"
    EnsureDokKupac kupId, "DOK STORNO KUPAC 2"

    Dim sc As String: sc = NewDokScenario("STOPRJ")
    Dim d As Date: d = DateSerial(2092, 11, 6)
    Dim brojPrj As String: brojPrj = DOC_PREFIX & "-PRJ-" & sc

    Dim id As String
    id = SavePrijemnica_TX(d, kupId, DK_VOZ_ID, brojPrj, DOC_PREFIX & "-ZBR-" & sc, _
                           DK_VRSTA, DK_SORTA, 100#, 10#, DK_TIP_AMB, 5, 0, KLASA_I)

    AssertDokEquals "1", CStr(RowsOf(GetPrijemniceByKupac(kupId, d, d))), "Storno prijemnica: visible before storno"
    AssertDokTrue StornoPrijemnicaByBroj_TX(brojPrj), "Storno prijemnica: StornoPrijemnicaByBroj_TX returns True"
    AssertDokEquals "Da", CStr(DokLookup(TBL_PRIJEMNICA, COL_PRJ_ID, id, COL_STORNIRANO)), "Storno prijemnica: flagged Stornirano=Da"
    AssertDokEquals "0", CStr(RowsOf(GetPrijemniceByKupac(kupId, d, d))), "Storno prijemnica: hidden from read after storno"
    Exit Sub

EH:
    LogDokFatal "Test_StornoPrijemnicaExcludesFromReads", Err.Number, Err.description
End Sub

' ============================================================
' ORPHAN / RECOVERY HELPERS (smoke)
' ============================================================

Private Sub Test_OrphanHelpersDoNotCrash()
    On Error GoTo EH

    Dim a As Variant, b As Variant, c As Variant
    a = GetOsirocenePrijemnice()
    b = GetAktivneZbirne()
    c = GetAktivnePrijemnice()

    AssertDokTrue (IsEmpty(a) Or IsArray(a)), "Recovery: GetOsirocenePrijemnice returns array-or-empty"
    AssertDokTrue (IsEmpty(b) Or IsArray(b)), "Recovery: GetAktivneZbirne returns array-or-empty"
    AssertDokTrue (IsEmpty(c) Or IsArray(c)), "Recovery: GetAktivnePrijemnice returns array-or-empty"
    Exit Sub

EH:
    LogDokFatal "Test_OrphanHelpersDoNotCrash", Err.Number, Err.description
End Sub

' ============================================================
' CLEANUP
' ============================================================

Public Sub SoftStornoDokumentaTestRows()
    On Error GoTo EH

    BeginDokRun "SOFT STORNO DOKUMENTA TEST ROWS"

    SoftStornoByMarker TBL_OTPREMNICA, Array(COL_OTP_BROJ, COL_OTP_BROJ_ZBIRNE)
    SoftStornoByMarker TBL_ZBIRNA, Array(COL_ZBR_BROJ)
    SoftStornoByMarker TBL_PRIJEMNICA, Array(COL_PRJ_BROJ, COL_PRJ_BROJ_ZBIRNE)

    EndDokRun
    Exit Sub

EH:
    LogDokFatal "SoftStornoDokumentaTestRows", Err.Number, Err.description
    EndDokRun
End Sub

Private Sub SoftStornoByMarker(ByVal tableName As String, ByVal markerColumns As Variant)
    On Error GoTo EH

    If GetTable(tableName) Is Nothing Then
        LogDokSkip "Soft-storno " & tableName, "Table not found"
        Exit Sub
    End If
    If GetColumnIndex(tableName, COL_STORNIRANO) = 0 Then
        LogDokSkip "Soft-storno " & tableName, "No Stornirano column"
        Exit Sub
    End If

    Dim data As Variant
    data = GetTableData(tableName)
    If IsEmpty(data) Then
        LogDokSkip "Soft-storno " & tableName, "No rows"
        Exit Sub
    End If

    Dim changed As Long, i As Long
    For i = 1 To UBound(data, 1)
        If RowHasMarker(data, i, tableName, markerColumns) Then
            UpdateCell tableName, i, COL_STORNIRANO, "Da"
            changed = changed + 1
        End If
    Next i

    LogDokPass "Soft-storno " & tableName & " changed " & changed & " row(s)"
    Exit Sub

EH:
    LogDokFail "Soft-storno " & tableName, Err.description
End Sub

Private Function RowHasMarker(ByVal data As Variant, ByVal rowIndex As Long, _
                              ByVal tableName As String, ByVal markerColumns As Variant) As Boolean
    Dim c As Variant, colIdx As Long
    For Each c In markerColumns
        colIdx = GetColumnIndex(tableName, CStr(c))
        If colIdx > 0 Then
            If InStr(1, CStr(data(rowIndex, colIdx)), DOC_PREFIX, vbTextCompare) > 0 Then
                RowHasMarker = True
                Exit Function
            End If
        End If
    Next c
End Function

' ============================================================
' DOMAIN HELPERS
' ============================================================

Private Function DokCountForBroj(ByVal tableName As String, ByVal brojColumn As String, _
                                 ByVal brojValue As String) As Long
    On Error GoTo EH
    Dim data As Variant
    data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function
    Dim colBroj As Long: colBroj = GetColumnIndex(tableName, brojColumn)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colBroj)) = brojValue Then DokCountForBroj = DokCountForBroj + 1
    Next i
    Exit Function
EH:
    DokCountForBroj = 0
End Function

Private Function DokFindIDByBrojKlasa(ByVal tableName As String, ByVal idColumn As String, _
                                      ByVal brojColumn As String, ByVal brojValue As String, _
                                      ByVal klasaColumn As String, ByVal klasa As String) As String
    On Error GoTo EH
    Dim data As Variant
    data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function

    Dim colID As Long: colID = GetColumnIndex(tableName, idColumn)
    Dim colBroj As Long: colBroj = GetColumnIndex(tableName, brojColumn)
    Dim colKlasa As Long: colKlasa = GetColumnIndex(tableName, klasaColumn)
    Dim i As Long
    For i = UBound(data, 1) To 1 Step -1
        If CStr(data(i, colBroj)) = brojValue And CStr(data(i, colKlasa)) = klasa Then
            DokFindIDByBrojKlasa = CStr(data(i, colID))
            Exit Function
        End If
    Next i
    Exit Function
EH:
    DokFindIDByBrojKlasa = ""
End Function

Private Function RowsOf(ByVal data As Variant) As Long
    On Error GoTo EH
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function
    RowsOf = UBound(data, 1)
    Exit Function
EH:
    RowsOf = 0
End Function

Private Sub EnsureDokStanica(ByVal stId As String, ByVal naziv As String)
    If DokRowExists(TBL_STANICE, "StanicaID", stId) Then Exit Sub
    Dim r As Variant
    r = DokBlankRow(TBL_STANICE)
    DokSetReq r, TBL_STANICE, "StanicaID", stId
    DokSetReq r, TBL_STANICE, "Naziv", naziv
    DokSetOpt r, TBL_STANICE, "Aktivan", "Aktivan"
    DokAppend TBL_STANICE, r, "EnsureDokStanica"
End Sub

Private Sub EnsureDokKupac(ByVal kupId As String, ByVal naziv As String)
    If DokRowExists(TBL_KUPCI, "KupacID", kupId) Then Exit Sub
    Dim r As Variant
    r = DokBlankRow(TBL_KUPCI)
    DokSetReq r, TBL_KUPCI, "KupacID", kupId
    DokSetReq r, TBL_KUPCI, "Naziv", naziv
    DokSetReq r, TBL_KUPCI, "PIB", "109000778"
    DokSetOpt r, TBL_KUPCI, "Aktivan", "Aktivan"
    DokAppend TBL_KUPCI, r, "EnsureDokKupac"
End Sub

' ============================================================
' GENERIC TABLE HELPERS
' ============================================================

Private Function DokBlankRow(ByVal tableName As String) As Variant
    Dim lo As ListObject
    Set lo = GetTable(tableName)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 9502, "modDokumentaTests.DokBlankRow", "Table not found: " & tableName
    End If
    Dim arr() As Variant
    ReDim arr(1 To lo.ListColumns.count)
    DokBlankRow = arr
End Function

Private Sub DokSetReq(ByRef rowData As Variant, ByVal tableName As String, _
                      ByVal columnName As String, ByVal value As Variant)
    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, columnName)
    If colIdx = 0 Then
        Err.Raise vbObjectError + 9503, "modDokumentaTests.DokSetReq", _
                  "Missing column: " & tableName & "." & columnName
    End If
    rowData(colIdx) = value
End Sub

Private Sub DokSetOpt(ByRef rowData As Variant, ByVal tableName As String, _
                      ByVal columnName As String, ByVal value As Variant)
    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, columnName)
    If colIdx > 0 Then rowData(colIdx) = value
End Sub

Private Sub DokAppend(ByVal tableName As String, ByVal rowData As Variant, ByVal sourceName As String)
    If AppendRow(tableName, rowData) <= 0 Then
        Err.Raise vbObjectError + 9504, "modDokumentaTests." & sourceName, "AppendRow failed for " & tableName
    End If
End Sub

Private Function DokRowExists(ByVal tableName As String, ByVal keyColumn As String, ByVal keyValue As String) As Boolean
    On Error GoTo EH
    If GetTable(tableName) Is Nothing Then Exit Function
    Dim colIdx As Long: colIdx = GetColumnIndex(tableName, keyColumn)
    If colIdx = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colIdx)) = CStr(keyValue) Then
            DokRowExists = True
            Exit Function
        End If
    Next i
    Exit Function
EH:
    DokRowExists = False
End Function

Private Function DokCountRows(ByVal tableName As String) As Long
    On Error GoTo EH
    Dim lo As ListObject
    Set lo = GetTable(tableName)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    DokCountRows = lo.DataBodyRange.rows.count
    Exit Function
EH:
    DokCountRows = 0
End Function

Private Function DokLookup(ByVal tableName As String, ByVal keyColumn As String, _
                           ByVal keyValue As String, ByVal returnColumn As String) As Variant
    On Error GoTo EH
    DokLookup = LookupValue(tableName, keyColumn, keyValue, returnColumn)
    Exit Function
EH:
    DokLookup = Empty
End Function

' ============================================================
' RUN STATE / ASSERTIONS / LOGGING (self-contained)
' ============================================================

Private Sub BeginDokRun(ByVal suiteName As String)
    ResetDokCounters
    InitDokLog

    Randomize
    m_RunID = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int((9999 - 1000 + 1) * Rnd + 1000))
    m_DateSeq = 0

    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "RunID=" & m_RunID
    Debug.Print String$(70, "=")

    AppendDokLog "SUITE", suiteName, "START", "RunID=" & m_RunID
End Sub

Private Sub EndDokRun()
    Dim summary As String
    summary = "RunID=" & m_RunID & _
              " | Total=" & m_Total & _
              " | Passed=" & m_Passed & _
              " | Failed=" & m_Failed & _
              " | Skipped=" & m_Skipped

    Debug.Print String$(70, "-")
    Debug.Print "DOKUMENTA TEST SUMMARY: " & summary
    Debug.Print String$(70, "-")

    AppendDokLog "SUITE", "SUMMARY", "INFO", summary

    If m_Failed > 0 Then
        MsgBox "Dokumenta tests finished with failures." & vbCrLf & summary, vbExclamation, APP_NAME
    Else
        MsgBox "Dokumenta tests finished." & vbCrLf & summary, vbInformation, APP_NAME
    End If
End Sub

Private Sub ResetDokCounters()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

Private Function NewDokScenario(ByVal scenarioName As String) As String
    NewDokScenario = scenarioName & "-" & m_RunID & "-" & CStr(m_Total + 1)
End Function

Private Function NextDokDate() As Date
    m_DateSeq = m_DateSeq + 1
    NextDokDate = DateSerial(2092, 1, 1) + m_DateSeq
End Function

Private Sub AssertDokTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        LogDokPass testName
    Else
        LogDokFail testName, "Assertion failed."
    End If
End Sub

Private Sub AssertDokEquals(ByVal expected As String, ByVal actual As String, ByVal testName As String)
    If CStr(expected) = CStr(actual) Then
        LogDokPass testName
    Else
        LogDokFail testName, "Expected [" & CStr(expected) & "], got [" & CStr(actual) & "]."
    End If
End Sub

Private Sub LogDokPass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1
    Debug.Print "[PASS] " & testName
    AppendDokLog "TEST", testName, "PASS", ""
End Sub

Private Sub LogDokFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendDokLog "TEST", testName, "FAIL", details
End Sub

Private Sub LogDokSkip(ByVal testName As String, ByVal reason As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1
    Debug.Print "[SKIP] " & testName & " :: " & reason
    AppendDokLog "TEST", testName, "SKIP", reason
End Sub

Private Sub LogDokInfo(ByVal message As String)
    Debug.Print "[INFO] " & message
    AppendDokLog "INFO", "", "INFO", message
End Sub

Private Sub LogDokFatal(ByVal sourceName As String, ByVal errNum As Long, ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
    AppendDokLog "FATAL", sourceName, "FAIL", CStr(errNum) & " - " & errDesc
End Sub

Private Sub InitDokLog()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DOK_TEST_LOG)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = DOK_TEST_LOG
        ws.Range("A1:G1").value = Array("Timestamp", "RunID", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendDokLog(ByVal kindText As String, ByVal nameText As String, _
                         ByVal statusText As String, ByVal detailsText As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DOK_TEST_LOG)
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
