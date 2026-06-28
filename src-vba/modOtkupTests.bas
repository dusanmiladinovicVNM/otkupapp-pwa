Attribute VB_Name = "modOtkupTests"
Option Explicit

' ============================================================
' modOtkupTests
'
' Professional regression suite for the OTKUP core (modOtkup).
'
' Covers:
'   - SaveOtkup            (inner save + validation that RAISES)
'   - SaveOtkup_TX         (transactional wrapper, returns "" on fail)
'   - SaveOtkupMulti_TX    (Klasa I / Klasa II / I+II, novac, ambalaza)
'   - GetOtkupByStation / GetOtkupByKooperant / GetSaldoByStation
'   - Ambalaza ledger side effects (TrackAmbalaza double-entry)
'   - Novac integration (KesOtkupacKoop isplata)
'   - Storno exclusion from read helpers
'
' Design (matches modFakturaTests / modSEFTests house style):
'   - Self-contained: own counters, own log sheet, own assert/log
'     helpers prefixed "Otk" to avoid Ambiguous-name on compile.
'   - Standalone runnable on a BLANK workbook: seeds its own master
'     data in an isolated ID namespace (ST-OT-*, KOOP-OT-*, ...).
'   - Tests create TST-OTK-* document rows; cleanup is soft-storno.
'
' Run from Alt+F8:
'   RunOtkupSuite              -> full suite
'   RunOtkupSeedOnly           -> only seed master data
'   SoftStornoOtkupTestRows    -> soft-storno all TST-OTK-* rows
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long
Private m_RunID As String
Private m_DateSeq As Long

Private Const OTKUP_TEST_LOG As String = "OTKUP_TEST_LOG"

' Isolated master-data namespace (independent of BusinessFlowPro seeds).
Private Const OT_ST_ID As String = "ST-OT-90001"
Private Const OT_KOOP_ID As String = "KOOP-OT-90001"
Private Const OT_VOZ_ID As String = "VOZ-OT-90001"
Private Const OT_KUL_ID As String = "KUL-OT-90001"
Private Const OT_PAR_ID As String = "PAR-OT-90001"

Private Const OT_VRSTA As String = "OT Test Jabuka"
Private Const OT_SORTA As String = "OT Test Sorta"
Private Const OT_TIP_AMB As String = "OT Test Gajba"

Private Const DOC_PREFIX As String = "TST-OTK"

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub RunOtkupSuite()
    On Error GoTo EH

    BeginOtkupRun "OTKUP CORE SUITE"

    Test_OtkupSchemaColumns
    SeedOtkupMaster
    Test_SeedAvailable

    ' ---- happy path persistence ----
    Test_SaveOtkupHappyKlasaI
    Test_SaveOtkupKulturaLookup
    Test_SaveOtkupVremeUnosaWritten
    Test_SaveOtkupBrutoOptional
    Test_SaveOtkupTXReturnsIdAndAppends
    Test_OtkupIdsAreSequential

    ' ---- multi-class wrappers ----
    Test_OtkupMultiKlasaIOnly
    Test_OtkupMultiKlasaIIOnly
    Test_OtkupMultiBothClasses

    ' ---- validation hardening (both layers) ----
    Test_OtkupMultiRequiresAtLeastOneClass
    Test_OtkupMultiKlasaIIIncompleteRejected
    Test_OtkupValidationEmptyKooperant
    Test_OtkupValidationEmptyStanica
    Test_OtkupValidationEmptyVrsta
    Test_OtkupValidationNonPositiveKolicina
    Test_OtkupValidationNonPositiveCena
    Test_OtkupValidationNegativeAmbalaza
    Test_OtkupValidationAmbWithoutTip
    Test_OtkupValidationInvalidKlasa
    Test_OtkupInvalidDoesNotMutateTables

    ' ---- ledger + novac side effects ----
    Test_OtkupAmbalazaDoubleLedger
    Test_OtkupAmbalazaIzdataLedger
    Test_OtkupNovacIntegration
    Test_OtkupNovacZeroNoRow

    ' ---- read helpers ----
    Test_GetOtkupByStationExcludesStornirano
    Test_GetOtkupByStationDateFilter
    Test_GetOtkupByKooperantExcludesStornirano
    Test_GetSaldoByStationAggregates
    Test_StornoOtkupRemovesFromReads

    EndOtkupRun
    Exit Sub

EH:
    LogOtkupFatal "RunOtkupSuite", Err.Number, Err.description
    EndOtkupRun
End Sub

Public Sub RunOtkupSeedOnly()
    On Error GoTo EH

    BeginOtkupRun "OTKUP SEED ONLY"

    Test_OtkupSchemaColumns
    SeedOtkupMaster
    Test_SeedAvailable

    EndOtkupRun
    Exit Sub

EH:
    LogOtkupFatal "RunOtkupSeedOnly", Err.Number, Err.description
    EndOtkupRun
End Sub

' ============================================================
' SCHEMA + SEED
' ============================================================

Private Sub Test_OtkupSchemaColumns()
    On Error GoTo EH

    If GetTable(TBL_OTKUP) Is Nothing Then
        LogOtkupFail "Schema: tblOtkup exists", "Table missing"
        Exit Sub
    End If

    Dim required As Variant
    required = Array(COL_OTK_ID, COL_OTK_DATUM, COL_OTK_KOOPERANT, COL_OTK_STANICA, _
                     COL_OTK_KULTURA, COL_OTK_VRSTA, COL_OTK_SORTA, COL_OTK_KOLICINA, _
                     COL_OTK_CENA, COL_OTK_TIP_AMB, COL_OTK_KOL_AMB, COL_OTK_KOL_AMB_IZDATA, _
                     COL_OTK_VREME_UNOSA, COL_OTK_BRUTO, COL_OTK_VOZAC, COL_OTK_BR_DOK, _
                     COL_OTK_NOVAC, COL_OTK_PRIMALAC, COL_OTK_KLASA, COL_OTK_STORNIRANO, _
                     COL_OTK_BROJ_ZBIRNE, COL_OTK_OTPREMNICA_ID, COL_OTK_PARCELA)

    Dim c As Variant
    Dim missing As String
    For Each c In required
        If GetColumnIndex(TBL_OTKUP, CStr(c)) = 0 Then missing = missing & CStr(c) & " "
    Next c

    AssertOtkupTrue (missing = ""), "Schema: tblOtkup has all required columns"
    If missing <> "" Then LogOtkupInfo "Missing tblOtkup columns: " & missing
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupSchemaColumns", Err.Number, Err.description
End Sub

Private Sub Test_SeedAvailable()
    On Error GoTo EH

    AssertOtkupTrue OtkRowExists(TBL_STANICE, "StanicaID", OT_ST_ID), "Seed: stanica available"
    AssertOtkupTrue OtkRowExists(TBL_VOZACI, "VozacID", OT_VOZ_ID), "Seed: vozac available"
    AssertOtkupTrue OtkRowExists(TBL_KOOPERANTI, COL_KOOP_ID, OT_KOOP_ID), "Seed: kooperant available"
    AssertOtkupTrue OtkRowExists(TBL_KULTURE, "KulturaID", OT_KUL_ID), "Seed: kultura available"
    Exit Sub

EH:
    LogOtkupFatal "Test_SeedAvailable", Err.Number, Err.description
End Sub

Private Sub SeedOtkupMaster()
    On Error GoTo EH

    SeedStanicaRow
    SeedVozacRow
    SeedKooperantRow
    SeedKulturaRow
    SeedParcelaRow

    LogOtkupPass "Seed master data ready"
    Exit Sub

EH:
    LogOtkupFail "Seed master data", Err.description
End Sub

Private Sub SeedStanicaRow()
    If OtkRowExists(TBL_STANICE, "StanicaID", OT_ST_ID) Then Exit Sub

    Dim r As Variant
    r = OtkBlankRow(TBL_STANICE)
    OtkSetReq r, TBL_STANICE, "StanicaID", OT_ST_ID
    OtkSetReq r, TBL_STANICE, "Naziv", "OT TEST STANICA"
    OtkSetOpt r, TBL_STANICE, "Mesto", "Test Mesto"
    OtkSetOpt r, TBL_STANICE, "Kontakt", "060111111"
    OtkSetOpt r, TBL_STANICE, "Ime", "OT"
    OtkSetOpt r, TBL_STANICE, "Prezime", "Stanica"
    OtkSetOpt r, TBL_STANICE, "PIN", "9101"
    OtkSetOpt r, TBL_STANICE, "Aktivan", "Aktivan"
    OtkAppend TBL_STANICE, r, "SeedStanicaRow"
End Sub

Private Sub SeedVozacRow()
    If OtkRowExists(TBL_VOZACI, "VozacID", OT_VOZ_ID) Then Exit Sub

    Dim r As Variant
    r = OtkBlankRow(TBL_VOZACI)
    OtkSetReq r, TBL_VOZACI, "VozacID", OT_VOZ_ID
    OtkSetReq r, TBL_VOZACI, "Ime", "OT"
    OtkSetReq r, TBL_VOZACI, "Prezime", "Vozac"
    OtkSetOpt r, TBL_VOZACI, "Telefon", "060111112"
    OtkSetOpt r, TBL_VOZACI, "Aktivan", "Aktivan"
    OtkAppend TBL_VOZACI, r, "SeedVozacRow"
End Sub

Private Sub SeedKooperantRow()
    If OtkRowExists(TBL_KOOPERANTI, COL_KOOP_ID, OT_KOOP_ID) Then Exit Sub

    Dim r As Variant
    r = OtkBlankRow(TBL_KOOPERANTI)
    OtkSetReq r, TBL_KOOPERANTI, "KooperantID", OT_KOOP_ID
    OtkSetReq r, TBL_KOOPERANTI, "Ime", "OT"
    OtkSetReq r, TBL_KOOPERANTI, "Prezime", "Kooperant"
    OtkSetReq r, TBL_KOOPERANTI, "StanicaID", OT_ST_ID
    OtkSetOpt r, TBL_KOOPERANTI, "Mesto", "Test Selo"
    OtkSetOpt r, TBL_KOOPERANTI, "Telefon", "060111113"
    OtkSetOpt r, TBL_KOOPERANTI, "Aktivan", "Da"
    OtkAppend TBL_KOOPERANTI, r, "SeedKooperantRow"
End Sub

Private Sub SeedKulturaRow()
    If OtkRowExists(TBL_KULTURE, "KulturaID", OT_KUL_ID) Then Exit Sub

    Dim r As Variant
    r = OtkBlankRow(TBL_KULTURE)
    OtkSetReq r, TBL_KULTURE, "KulturaID", OT_KUL_ID
    OtkSetReq r, TBL_KULTURE, "VrstaVoca", OT_VRSTA
    OtkSetReq r, TBL_KULTURE, "SortaVoca", OT_SORTA
    OtkSetOpt r, TBL_KULTURE, "Aktivan", "Aktivan"
    OtkAppend TBL_KULTURE, r, "SeedKulturaRow"
End Sub

Private Sub SeedParcelaRow()
    If GetTable(TBL_PARCELE) Is Nothing Then Exit Sub
    If OtkRowExists(TBL_PARCELE, "ParcelaID", OT_PAR_ID) Then Exit Sub

    Dim r As Variant
    r = OtkBlankRow(TBL_PARCELE)
    OtkSetReq r, TBL_PARCELE, "ParcelaID", OT_PAR_ID
    OtkSetReq r, TBL_PARCELE, "KooperantID", OT_KOOP_ID
    OtkSetReq r, TBL_PARCELE, "KatBroj", "OT-TEST-1"
    OtkSetOpt r, TBL_PARCELE, "Kultura", OT_SORTA
    OtkSetOpt r, TBL_PARCELE, "Aktivna", "Da"
    OtkSetOpt r, TBL_PARCELE, "Aktivan", "Aktivan"
    OtkAppend TBL_PARCELE, r, "SeedParcelaRow"
End Sub

' ============================================================
' HAPPY PATH PERSISTENCE
' ============================================================

Private Sub Test_SaveOtkupHappyKlasaI()
    On Error GoTo EH

    Dim brDok As String: brDok = NewOtkupScenario("HAPPY")
    Dim d As Date: d = NextOtkupDate()

    Dim id As String
    id = SaveOtkup(d, OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                   500#, 100#, OT_TIP_AMB, 25, OT_VOZ_ID, brDok, 0#, "")

    AssertOtkupTrue (Left$(id, 4) = "OTK-"), "SaveOtkup: returns OTK- id"
    AssertOtkupEquals OT_KOOP_ID, OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_KOOPERANT), "SaveOtkup: kooperant persisted"
    AssertOtkupEquals OT_ST_ID, OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_STANICA), "SaveOtkup: stanica persisted"
    AssertOtkupEquals OT_VRSTA, OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_VRSTA), "SaveOtkup: vrsta persisted"
    AssertOtkupEquals "500", OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_KOLICINA), "SaveOtkup: kolicina persisted"
    AssertOtkupEquals "100", OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_CENA), "SaveOtkup: cena persisted"
    AssertOtkupEquals KLASA_I, OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_KLASA), "SaveOtkup: klasa defaults to I"
    Exit Sub

EH:
    LogOtkupFatal "Test_SaveOtkupHappyKlasaI", Err.Number, Err.description
End Sub

Private Sub Test_SaveOtkupKulturaLookup()
    On Error GoTo EH

    Dim brDok As String: brDok = NewOtkupScenario("KULT")
    Dim id As String
    id = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                   100#, 80#, OT_TIP_AMB, 5, OT_VOZ_ID, brDok, 0#, "")

    AssertOtkupTrue Len(id) > 0, "SaveOtkup: kultura lookup row saved"
    AssertOtkupEquals OT_KUL_ID, OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_KULTURA), _
                      "SaveOtkup: KulturaID resolved from tblKulture by VrstaVoca"
    Exit Sub

EH:
    LogOtkupFatal "Test_SaveOtkupKulturaLookup", Err.Number, Err.description
End Sub

Private Sub Test_SaveOtkupVremeUnosaWritten()
    On Error GoTo EH

    Dim brDok As String: brDok = NewOtkupScenario("VREME")
    Dim id As String
    id = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                   100#, 80#, OT_TIP_AMB, 5, OT_VOZ_ID, brDok, 0#, "")

    AssertOtkupTrue Len(Trim$(CStr(OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_VREME_UNOSA)))) > 0, _
                    "SaveOtkup: VremeUnosa timestamp written"
    Exit Sub

EH:
    LogOtkupFatal "Test_SaveOtkupVremeUnosaWritten", Err.Number, Err.description
End Sub

Private Sub Test_SaveOtkupBrutoOptional()
    On Error GoTo EH

    ' brutoKg = 0 -> column stays blank (neto)
    Dim brNeto As String: brNeto = NewOtkupScenario("NETO")
    Dim idNeto As String
    idNeto = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                       100#, 80#, OT_TIP_AMB, 5, OT_VOZ_ID, brNeto, 0#, "", _
                       KLASA_I, "", "", 0, 0#)
    AssertOtkupEquals "", Trim$(CStr(OtkLookup(TBL_OTKUP, COL_OTK_ID, idNeto, COL_OTK_BRUTO))), _
                      "SaveOtkup: brutoKg=0 leaves BrutoKg blank"

    ' brutoKg > 0 -> persisted
    Dim brBruto As String: brBruto = NewOtkupScenario("BRUTO")
    Dim idBruto As String
    idBruto = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                        100#, 80#, OT_TIP_AMB, 5, OT_VOZ_ID, brBruto, 0#, "", _
                        KLASA_I, "", "", 0, 112.5)
    AssertOtkupTrue IsNumeric(OtkLookup(TBL_OTKUP, COL_OTK_ID, idBruto, COL_OTK_BRUTO)), _
                    "SaveOtkup: brutoKg>0 persisted as number"
    AssertOtkupTrue (Abs(CDbl(OtkLookup(TBL_OTKUP, COL_OTK_ID, idBruto, COL_OTK_BRUTO)) - 112.5) < 0.001), _
                    "SaveOtkup: brutoKg>0 value equals 112.5"
    Exit Sub

EH:
    LogOtkupFatal "Test_SaveOtkupBrutoOptional", Err.Number, Err.description
End Sub

Private Sub Test_SaveOtkupTXReturnsIdAndAppends()
    On Error GoTo EH

    Dim before As Long: before = OtkCountRows(TBL_OTKUP)
    Dim brDok As String: brDok = NewOtkupScenario("TX")

    Dim id As String
    id = SaveOtkup_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                      300#, 90#, OT_TIP_AMB, 15, OT_VOZ_ID, brDok, 0#, "")

    AssertOtkupTrue (Left$(id, 4) = "OTK-"), "SaveOtkup_TX: returns OTK- id"
    AssertOtkupEquals CStr(before + 1), CStr(OtkCountRows(TBL_OTKUP)), "SaveOtkup_TX: appended exactly one row"
    Exit Sub

EH:
    LogOtkupFatal "Test_SaveOtkupTXReturnsIdAndAppends", Err.Number, Err.description
End Sub

Private Sub Test_OtkupIdsAreSequential()
    On Error GoTo EH

    Dim id1 As String, id2 As String
    id1 = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                    50#, 70#, OT_TIP_AMB, 2, OT_VOZ_ID, NewOtkupScenario("SEQ1"), 0#, "")
    id2 = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                    50#, 70#, OT_TIP_AMB, 2, OT_VOZ_ID, NewOtkupScenario("SEQ2"), 0#, "")

    AssertOtkupTrue (Len(id1) > 0 And Len(id2) > 0 And id1 <> id2), _
                    "GetNextID: consecutive otkup ids are unique"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupIdsAreSequential", Err.Number, Err.description
End Sub

' ============================================================
' MULTI-CLASS WRAPPER
' ============================================================

Private Sub Test_OtkupMultiKlasaIOnly()
    On Error GoTo EH

    Dim brDok As String: brDok = NewOtkupScenario("M1")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            400#, 100#, OT_TIP_AMB, 20, OT_VOZ_ID, brDok, 0#, "", "", "", _
                            False)

    AssertOtkupTrue Len(res) > 0, "Multi Klasa I: save returns id"
    AssertOtkupEquals "1", CStr(CountOtkupForBrDok(brDok)), "Multi Klasa I: exactly one otkup row"
    AssertOtkupTrue Len(FindOtkupIDByBrojKlasa(brDok, KLASA_I)) > 0, "Multi Klasa I: Klasa I row present"
    AssertOtkupEquals "", FindOtkupIDByBrojKlasa(brDok, KLASA_II), "Multi Klasa I: no Klasa II row"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupMultiKlasaIOnly", Err.Number, Err.description
End Sub

Private Sub Test_OtkupMultiKlasaIIOnly()
    On Error GoTo EH

    ' kolicinaI = 0 -> Klasa I skipped; only Klasa II saved.
    Dim brDok As String: brDok = NewOtkupScenario("M2")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            0#, 0#, OT_TIP_AMB, 0, OT_VOZ_ID, brDok, 0#, "", "", "", _
                            True, 150#, 60#, 0, 0#, 8)

    AssertOtkupTrue Len(res) > 0, "Multi Klasa II only: save returns id"
    AssertOtkupEquals "", FindOtkupIDByBrojKlasa(brDok, KLASA_I), "Multi Klasa II only: no Klasa I row"
    AssertOtkupTrue Len(FindOtkupIDByBrojKlasa(brDok, KLASA_II)) > 0, "Multi Klasa II only: Klasa II row present"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupMultiKlasaIIOnly", Err.Number, Err.description
End Sub

Private Sub Test_OtkupMultiBothClasses()
    On Error GoTo EH

    Dim brDok As String: brDok = NewOtkupScenario("M12")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            400#, 100#, OT_TIP_AMB, 20, OT_VOZ_ID, brDok, 0#, "", "", "", _
                            True, 150#, 60#, 0, 0#, 8)

    Dim idI As String, idII As String
    idI = FindOtkupIDByBrojKlasa(brDok, KLASA_I)
    idII = FindOtkupIDByBrojKlasa(brDok, KLASA_II)

    AssertOtkupTrue (Len(idI) > 0 And Len(idII) > 0), "Multi I+II: both class rows present"
    AssertOtkupTrue (idI <> idII), "Multi I+II: distinct OtkupID per class"
    AssertOtkupEquals "2", CStr(CountOtkupForBrDok(brDok)), "Multi I+II: shares BrojDokumenta across two rows"
    AssertOtkupEquals (idI & " + " & idII), res, "Multi I+II: result is 'I + II' composite id"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupMultiBothClasses", Err.Number, Err.description
End Sub

' ============================================================
' VALIDATION HARDENING
' ============================================================

Private Sub Test_OtkupMultiRequiresAtLeastOneClass()
    On Error GoTo EH

    Dim brDok As String: brDok = NewOtkupScenario("NOCLASS")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            0#, 0#, OT_TIP_AMB, 0, OT_VOZ_ID, brDok, 0#, "", "", "", _
                            False)

    AssertOtkupEquals "", res, "Multi: rejects save with no class (I=0, II=off)"
    AssertOtkupEquals "0", CStr(CountOtkupForBrDok(brDok)), "Multi: no row appended when no class"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupMultiRequiresAtLeastOneClass", Err.Number, Err.description
End Sub

Private Sub Test_OtkupMultiKlasaIIIncompleteRejected()
    On Error GoTo EH

    ' hasKlasaII = True but kolicinaII/cenaII <= 0 -> rejected.
    Dim brDok As String: brDok = NewOtkupScenario("BADII")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, brDok, 0#, "", "", "", _
                            True, 0#, 0#)

    AssertOtkupEquals "", res, "Multi: rejects Klasa II with zero kolicina/cena"
    AssertOtkupEquals "0", CStr(CountOtkupForBrDok(brDok)), "Multi: nothing appended on bad Klasa II"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupMultiKlasaIIIncompleteRejected", Err.Number, Err.description
End Sub

Private Sub Test_OtkupValidationEmptyKooperant()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), "", OT_ST_ID, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("NOKOOP"), 0#, "", KLASA_I), _
        "Validation: empty kooperant rejected"
End Sub

Private Sub Test_OtkupValidationEmptyStanica()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), OT_KOOP_ID, "", OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("NOST"), 0#, "", KLASA_I), _
        "Validation: empty stanica rejected"
End Sub

Private Sub Test_OtkupValidationEmptyVrsta()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, "", OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("NOVRSTA"), 0#, "", KLASA_I), _
        "Validation: empty vrsta rejected"
End Sub

Private Sub Test_OtkupValidationNonPositiveKolicina()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, 0#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("ZEROKOL"), 0#, "", KLASA_I), _
        "Validation: kolicina <= 0 rejected"
End Sub

Private Sub Test_OtkupValidationNonPositiveCena()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, 100#, 0#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("ZEROCENA"), 0#, "", KLASA_I), _
        "Validation: cena <= 0 rejected"
End Sub

Private Sub Test_OtkupValidationNegativeAmbalaza()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, -3, OT_VOZ_ID, NewOtkupScenario("NEGAMB"), 0#, "", KLASA_I), _
        "Validation: negative ambalaza rejected"
End Sub

Private Sub Test_OtkupValidationAmbWithoutTip()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, 100#, 90#, "", 5, OT_VOZ_ID, NewOtkupScenario("NOTIP"), 0#, "", KLASA_I), _
        "Validation: ambalaza without tip rejected"
End Sub

Private Sub Test_OtkupValidationInvalidKlasa()
    AssertOtkupRejected _
        TrySaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("BADKLASA"), 0#, "", "III"), _
        "Validation: invalid klasa rejected"
End Sub

Private Sub Test_OtkupInvalidDoesNotMutateTables()
    On Error GoTo EH

    Dim beforeOtk As Long: beforeOtk = OtkCountRows(TBL_OTKUP)
    Dim beforeNov As Long: beforeNov = OtkCountRows(TBL_NOVAC)
    Dim beforeAmb As Long: beforeAmb = OtkCountRows(TBL_AMBALAZA)

    ' Negative ambalaza fails inside the transaction -> full rollback expected.
    Dim brDok As String: brDok = NewOtkupScenario("ATOMIC")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            100#, 90#, OT_TIP_AMB, -5, OT_VOZ_ID, brDok, 250#, "Avans", "", "", _
                            False)

    AssertOtkupEquals "", res, "Atomicity: invalid multi-save returns empty id"
    AssertOtkupEquals CStr(beforeOtk), CStr(OtkCountRows(TBL_OTKUP)), "Atomicity: tblOtkup unchanged"
    AssertOtkupEquals CStr(beforeNov), CStr(OtkCountRows(TBL_NOVAC)), "Atomicity: tblNovac unchanged"
    AssertOtkupEquals CStr(beforeAmb), CStr(OtkCountRows(TBL_AMBALAZA)), "Atomicity: tblAmbalaza unchanged"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupInvalidDoesNotMutateTables", Err.Number, Err.description
End Sub

' ============================================================
' AMBALAZA + NOVAC SIDE EFFECTS
' ============================================================

Private Sub Test_OtkupAmbalazaDoubleLedger()
    On Error GoTo EH

    Dim brDok As String: brDok = NewOtkupScenario("AMB")
    Dim id As String
    id = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                   100#, 90#, OT_TIP_AMB, 10, OT_VOZ_ID, brDok, 0#, "")

    AssertOtkupTrue Len(id) > 0, "Ledger: otkup with ambalaza saved"
    AssertOtkupEquals "2", CStr(CountAmbForDok(id, DOK_TIP_OTKUP)), _
                      "Ledger: kolAmb>0 creates two Otkup ledger entries (Kooperant Izlaz + Stanica Ulaz)"
    AssertOtkupTrue HasAmbEntry(id, DOK_TIP_OTKUP, "Izlaz", "Kooperant"), "Ledger: Kooperant Izlaz entry present"
    AssertOtkupTrue HasAmbEntry(id, DOK_TIP_OTKUP, "Ulaz", "Stanica"), "Ledger: Stanica Ulaz entry present"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupAmbalazaDoubleLedger", Err.Number, Err.description
End Sub

Private Sub Test_OtkupAmbalazaIzdataLedger()
    On Error GoTo EH

    ' kolAmbIzdata > 0 -> OM hands out empty crates to kooperant (OM-Izlaz-Koop).
    Dim brDok As String: brDok = NewOtkupScenario("AMBIZD")
    Dim id As String
    id = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                   100#, 90#, OT_TIP_AMB, 0, OT_VOZ_ID, brDok, 0#, "", _
                   KLASA_I, "", "", 6, 0#)

    AssertOtkupTrue Len(id) > 0, "Ledger izdata: otkup saved"
    AssertOtkupEquals "2", CStr(CountAmbForDok(id, DOK_TIP_OM_IZLAZ_KOOP)), _
                      "Ledger izdata: kolAmbIzdata>0 creates two OM-Izlaz-Koop entries"
    AssertOtkupTrue HasAmbEntry(id, DOK_TIP_OM_IZLAZ_KOOP, "Ulaz", "Kooperant"), "Ledger izdata: Kooperant Ulaz entry present"
    AssertOtkupTrue HasAmbEntry(id, DOK_TIP_OM_IZLAZ_KOOP, "Izlaz", "Stanica"), "Ledger izdata: Stanica Izlaz entry present"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupAmbalazaIzdataLedger", Err.Number, Err.description
End Sub

Private Sub Test_OtkupNovacIntegration()
    On Error GoTo EH

    Dim before As Long: before = OtkCountRows(TBL_NOVAC)
    Dim brDok As String: brDok = NewOtkupScenario("NOVAC")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, brDok, 250#, "Avans gotovina", "", "", _
                            False)

    Dim primaryID As String: primaryID = FindOtkupIDByBrojKlasa(brDok, KLASA_I)

    AssertOtkupTrue Len(res) > 0, "Novac: otkup with cash saved"
    AssertOtkupEquals CStr(before + 1), CStr(OtkCountRows(TBL_NOVAC)), "Novac: exactly one tblNovac row created"
    AssertOtkupEquals "250", CStr(OtkLookup(TBL_NOVAC, COL_NOV_OTKUP_ID, primaryID, COL_NOV_ISPLATA)), _
                      "Novac: isplata linked to primary OtkupID equals cash amount"
    AssertOtkupEquals NOV_KES_OTKUPAC_KOOP, CStr(OtkLookup(TBL_NOVAC, COL_NOV_OTKUP_ID, primaryID, COL_NOV_TIP)), _
                      "Novac: tip is KesOtkupacKoop"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupNovacIntegration", Err.Number, Err.description
End Sub

Private Sub Test_OtkupNovacZeroNoRow()
    On Error GoTo EH

    Dim before As Long: before = OtkCountRows(TBL_NOVAC)
    Dim brDok As String: brDok = NewOtkupScenario("NONOV")
    Dim res As String
    res = SaveOtkupMulti_TX(NextOtkupDate(), OT_KOOP_ID, OT_ST_ID, OT_VRSTA, OT_SORTA, _
                            100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, brDok, 0#, "", "", "", _
                            False)

    AssertOtkupTrue Len(res) > 0, "Novac zero: otkup saved"
    AssertOtkupEquals CStr(before), CStr(OtkCountRows(TBL_NOVAC)), "Novac zero: no tblNovac row when novac=0"
    Exit Sub

EH:
    LogOtkupFatal "Test_OtkupNovacZeroNoRow", Err.Number, Err.description
End Sub

' ============================================================
' READ HELPERS
' ============================================================

Private Sub Test_GetOtkupByStationExcludesStornirano()
    On Error GoTo EH

    Dim stId As String: stId = "ST-OT-READ1"
    EnsureStanica stId, "OT READ STANICA 1"

    Dim brA As String: brA = NewOtkupScenario("RDA")
    Dim brB As String: brB = NewOtkupScenario("RDB")

    Dim idA As String, idB As String
    idA = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, stId, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, brA, 0#, "")
    idB = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, stId, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, brB, 0#, "")

    Dim beforeCnt As Long: beforeCnt = RowsOf(GetOtkupByStation(stId))
    AssertOtkupEquals "2", CStr(beforeCnt), "GetOtkupByStation: returns both active rows"

    StornoOtkupByBrDok_TX brA

    Dim afterCnt As Long: afterCnt = RowsOf(GetOtkupByStation(stId))
    AssertOtkupEquals "1", CStr(afterCnt), "GetOtkupByStation: excludes stornirano row"
    Exit Sub

EH:
    LogOtkupFatal "Test_GetOtkupByStationExcludesStornirano", Err.Number, Err.description
End Sub

Private Sub Test_GetOtkupByStationDateFilter()
    On Error GoTo EH

    Dim stId As String: stId = "ST-OT-READ2"
    EnsureStanica stId, "OT READ STANICA 2"

    Dim dIn As Date: dIn = DateSerial(2091, 6, 15)
    Dim dOut As Date: dOut = DateSerial(2091, 9, 15)

    SaveOtkup dIn, OT_KOOP_ID, stId, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("DIN"), 0#, ""
    SaveOtkup dOut, OT_KOOP_ID, stId, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("DOUT"), 0#, ""

    Dim cnt As Long
    cnt = RowsOf(GetOtkupByStation(stId, DateSerial(2091, 6, 1), DateSerial(2091, 6, 30)))
    AssertOtkupEquals "1", CStr(cnt), "GetOtkupByStation: BETWEEN date filter returns only in-range row"
    Exit Sub

EH:
    LogOtkupFatal "Test_GetOtkupByStationDateFilter", Err.Number, Err.description
End Sub

Private Sub Test_GetOtkupByKooperantExcludesStornirano()
    On Error GoTo EH

    Dim koopId As String: koopId = "KOOP-OT-READ1"
    EnsureKooperant koopId, "OT", "ReadKoop"

    Dim brA As String: brA = NewOtkupScenario("KRDA")
    SaveOtkup NextOtkupDate(), koopId, OT_ST_ID, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, brA, 0#, ""
    SaveOtkup NextOtkupDate(), koopId, OT_ST_ID, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("KRDB"), 0#, ""

    AssertOtkupEquals "2", CStr(RowsOf(GetOtkupByKooperant(koopId))), "GetOtkupByKooperant: returns both active rows"

    StornoOtkupByBrDok_TX brA
    AssertOtkupEquals "1", CStr(RowsOf(GetOtkupByKooperant(koopId))), "GetOtkupByKooperant: excludes stornirano row"
    Exit Sub

EH:
    LogOtkupFatal "Test_GetOtkupByKooperantExcludesStornirano", Err.Number, Err.description
End Sub

Private Sub Test_GetSaldoByStationAggregates()
    On Error GoTo EH

    Dim stId As String: stId = "ST-OT-SALDO1"
    EnsureStanica stId, "OT SALDO STANICA"
    Dim koopId As String: koopId = "KOOP-OT-SALDO1"
    EnsureKooperant koopId, "OT", "SaldoKoop"

    SaveOtkup NextOtkupDate(), koopId, stId, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 10, OT_VOZ_ID, NewOtkupScenario("SAL1"), 0#, ""
    SaveOtkup NextOtkupDate(), koopId, stId, OT_VRSTA, OT_SORTA, 200#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, NewOtkupScenario("SAL2"), 0#, ""

    Dim saldo As Variant
    saldo = GetSaldoByStation(stId)

    Dim kg As Double
    kg = SaldoKolicinaForKoop(saldo, koopId)
    AssertOtkupTrue (Abs(kg - 300#) < 0.01), "GetSaldoByStation: aggregates kolicina per kooperant (100+200=300)"
    Exit Sub

EH:
    LogOtkupFatal "Test_GetSaldoByStationAggregates", Err.Number, Err.description
End Sub

Private Sub Test_StornoOtkupRemovesFromReads()
    On Error GoTo EH

    Dim stId As String: stId = "ST-OT-STORNO1"
    EnsureStanica stId, "OT STORNO STANICA"

    Dim brDok As String: brDok = NewOtkupScenario("STO")
    Dim id As String
    id = SaveOtkup(NextOtkupDate(), OT_KOOP_ID, stId, OT_VRSTA, OT_SORTA, 100#, 90#, OT_TIP_AMB, 5, OT_VOZ_ID, brDok, 0#, "")

    AssertOtkupEquals "1", CStr(RowsOf(GetOtkupByStation(stId))), "Storno: row visible before storno"

    Dim ok As Boolean
    ok = StornoOtkupByBrDok_TX(brDok)
    AssertOtkupTrue ok, "Storno: StornoOtkupByBrDok_TX returns True"
    AssertOtkupEquals "Da", CStr(OtkLookup(TBL_OTKUP, COL_OTK_ID, id, COL_OTK_STORNIRANO)), "Storno: row flagged Stornirano=Da"
    AssertOtkupEquals "0", CStr(RowsOf(GetOtkupByStation(stId))), "Storno: row hidden from read helper after storno"
    Exit Sub

EH:
    LogOtkupFatal "Test_StornoOtkupRemovesFromReads", Err.Number, Err.description
End Sub

' ============================================================
' CLEANUP
' ============================================================

Public Sub SoftStornoOtkupTestRows()
    On Error GoTo EH

    BeginOtkupRun "SOFT STORNO OTKUP TEST ROWS"

    SoftStornoByMarker TBL_OTKUP, Array(COL_OTK_BR_DOK)
    SoftStornoByMarker TBL_NOVAC, Array(COL_NOV_BROJ_DOK)

    EndOtkupRun
    Exit Sub

EH:
    LogOtkupFatal "SoftStornoOtkupTestRows", Err.Number, Err.description
    EndOtkupRun
End Sub

Private Sub SoftStornoByMarker(ByVal tableName As String, ByVal markerColumns As Variant)
    On Error GoTo EH

    If GetTable(tableName) Is Nothing Then
        LogOtkupSkip "Soft-storno " & tableName, "Table not found"
        Exit Sub
    End If
    If GetColumnIndex(tableName, COL_STORNIRANO) = 0 Then
        LogOtkupSkip "Soft-storno " & tableName, "No Stornirano column"
        Exit Sub
    End If

    Dim data As Variant
    data = GetTableData(tableName)
    If IsEmpty(data) Then
        LogOtkupSkip "Soft-storno " & tableName, "No rows"
        Exit Sub
    End If

    Dim changed As Long, i As Long
    For i = 1 To UBound(data, 1)
        If RowHasMarker(data, i, tableName, markerColumns) Then
            UpdateCell tableName, i, COL_STORNIRANO, "Da"
            changed = changed + 1
        End If
    Next i

    LogOtkupPass "Soft-storno " & tableName & " changed " & changed & " row(s)"
    Exit Sub

EH:
    LogOtkupFail "Soft-storno " & tableName, Err.description
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

' Calls SaveOtkup and converts a RAISED validation error into "" so the
' caller can assert rejection uniformly (SaveOtkup raises; _TX returns "").
Private Function TrySaveOtkup(ByVal datum As Date, ByVal kooperantID As String, _
                              ByVal stanicaID As String, ByVal vrstaVoca As String, _
                              ByVal sortaVoca As String, ByVal kolicina As Double, _
                              ByVal cena As Double, ByVal tipAmb As String, _
                              ByVal kolAmb As Long, ByVal vozacID As String, _
                              ByVal brDok As String, ByVal novac As Double, _
                              ByVal primalac As String, ByVal klasa As String) As String
    On Error GoTo EH
    TrySaveOtkup = SaveOtkup(datum, kooperantID, stanicaID, vrstaVoca, sortaVoca, _
                             kolicina, cena, tipAmb, kolAmb, vozacID, brDok, novac, _
                             primalac, klasa)
    Exit Function
EH:
    TrySaveOtkup = ""
End Function

Private Sub AssertOtkupRejected(ByVal result As String, ByVal testName As String)
    AssertOtkupEquals "", result, testName
End Sub

Private Function CountOtkupForBrDok(ByVal brDok As String) As Long
    On Error GoTo EH
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim colBroj As Long: colBroj = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colBroj)) = brDok Then CountOtkupForBrDok = CountOtkupForBrDok + 1
    Next i
    Exit Function
EH:
    CountOtkupForBrDok = 0
End Function

Private Function FindOtkupIDByBrojKlasa(ByVal brDok As String, ByVal klasa As String) As String
    On Error GoTo EH
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim colID As Long: colID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    Dim colBroj As Long: colBroj = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    Dim colKlasa As Long: colKlasa = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)

    Dim i As Long
    For i = UBound(data, 1) To 1 Step -1
        If CStr(data(i, colBroj)) = brDok And CStr(data(i, colKlasa)) = klasa Then
            FindOtkupIDByBrojKlasa = CStr(data(i, colID))
            Exit Function
        End If
    Next i
    Exit Function
EH:
    FindOtkupIDByBrojKlasa = ""
End Function

Private Function CountAmbForDok(ByVal dokID As String, ByVal dokTip As String) As Long
    On Error GoTo EH
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim colDok As Long: colDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    Dim colTip As Long: colTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colDok)) = dokID And CStr(data(i, colTip)) = dokTip Then
            CountAmbForDok = CountAmbForDok + 1
        End If
    Next i
    Exit Function
EH:
    CountAmbForDok = 0
End Function

Private Function HasAmbEntry(ByVal dokID As String, ByVal dokTip As String, _
                             ByVal smer As String, ByVal entitetTip As String) As Boolean
    On Error GoTo EH
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim colDok As Long: colDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    Dim colTip As Long: colTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    Dim colSmer As Long: colSmer = GetColumnIndex(TBL_AMBALAZA, COL_AMB_SMER)
    Dim colEnt As Long: colEnt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colDok)) = dokID And CStr(data(i, colTip)) = dokTip _
           And CStr(data(i, colSmer)) = smer And CStr(data(i, colEnt)) = entitetTip Then
            HasAmbEntry = True
            Exit Function
        End If
    Next i
    Exit Function
EH:
    HasAmbEntry = False
End Function

Private Function SaldoKolicinaForKoop(ByVal saldo As Variant, ByVal koopId As String) As Double
    On Error GoTo EH
    If IsEmpty(saldo) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(saldo, 1)
        If CStr(saldo(i, 1)) = koopId Then
            If IsNumeric(saldo(i, 2)) Then SaldoKolicinaForKoop = CDbl(saldo(i, 2))
            Exit Function
        End If
    Next i
    Exit Function
EH:
    SaldoKolicinaForKoop = 0
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

Private Sub EnsureStanica(ByVal stId As String, ByVal naziv As String)
    If OtkRowExists(TBL_STANICE, "StanicaID", stId) Then Exit Sub
    Dim r As Variant
    r = OtkBlankRow(TBL_STANICE)
    OtkSetReq r, TBL_STANICE, "StanicaID", stId
    OtkSetReq r, TBL_STANICE, "Naziv", naziv
    OtkSetOpt r, TBL_STANICE, "Aktivan", "Aktivan"
    OtkAppend TBL_STANICE, r, "EnsureStanica"
End Sub

Private Sub EnsureKooperant(ByVal koopId As String, ByVal ime As String, ByVal prezime As String)
    If OtkRowExists(TBL_KOOPERANTI, COL_KOOP_ID, koopId) Then Exit Sub
    Dim r As Variant
    r = OtkBlankRow(TBL_KOOPERANTI)
    OtkSetReq r, TBL_KOOPERANTI, "KooperantID", koopId
    OtkSetReq r, TBL_KOOPERANTI, "Ime", ime
    OtkSetReq r, TBL_KOOPERANTI, "Prezime", prezime
    OtkSetReq r, TBL_KOOPERANTI, "StanicaID", OT_ST_ID
    OtkSetOpt r, TBL_KOOPERANTI, "Aktivan", "Da"
    OtkAppend TBL_KOOPERANTI, r, "EnsureKooperant"
End Sub

' ============================================================
' GENERIC TABLE HELPERS
' ============================================================

Private Function OtkBlankRow(ByVal tableName As String) As Variant
    Dim lo As ListObject
    Set lo = GetTable(tableName)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 9402, "modOtkupTests.OtkBlankRow", "Table not found: " & tableName
    End If
    Dim arr() As Variant
    ReDim arr(1 To lo.ListColumns.count)
    OtkBlankRow = arr
End Function

Private Sub OtkSetReq(ByRef rowData As Variant, ByVal tableName As String, _
                      ByVal columnName As String, ByVal value As Variant)
    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, columnName)
    If colIdx = 0 Then
        Err.Raise vbObjectError + 9403, "modOtkupTests.OtkSetReq", _
                  "Missing column: " & tableName & "." & columnName
    End If
    rowData(colIdx) = value
End Sub

Private Sub OtkSetOpt(ByRef rowData As Variant, ByVal tableName As String, _
                      ByVal columnName As String, ByVal value As Variant)
    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, columnName)
    If colIdx > 0 Then rowData(colIdx) = value
End Sub

Private Sub OtkAppend(ByVal tableName As String, ByVal rowData As Variant, ByVal sourceName As String)
    If AppendRow(tableName, rowData) <= 0 Then
        Err.Raise vbObjectError + 9404, "modOtkupTests." & sourceName, "AppendRow failed for " & tableName
    End If
End Sub

Private Function OtkRowExists(ByVal tableName As String, ByVal keyColumn As String, ByVal keyValue As String) As Boolean
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
            OtkRowExists = True
            Exit Function
        End If
    Next i
    Exit Function
EH:
    OtkRowExists = False
End Function

Private Function OtkCountRows(ByVal tableName As String) As Long
    On Error GoTo EH
    Dim lo As ListObject
    Set lo = GetTable(tableName)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    OtkCountRows = lo.DataBodyRange.rows.count
    Exit Function
EH:
    OtkCountRows = 0
End Function

Private Function OtkLookup(ByVal tableName As String, ByVal keyColumn As String, _
                           ByVal keyValue As String, ByVal returnColumn As String) As Variant
    On Error GoTo EH
    OtkLookup = LookupValue(tableName, keyColumn, keyValue, returnColumn)
    Exit Function
EH:
    OtkLookup = Empty
End Function

' ============================================================
' RUN STATE / ASSERTIONS / LOGGING (self-contained)
' ============================================================

Private Sub BeginOtkupRun(ByVal suiteName As String)
    ResetOtkupCounters
    InitOtkupLog

    Randomize
    m_RunID = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int((9999 - 1000 + 1) * Rnd + 1000))
    m_DateSeq = 0

    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "RunID=" & m_RunID
    Debug.Print String$(70, "=")

    AppendOtkupLog "SUITE", suiteName, "START", "RunID=" & m_RunID
End Sub

Private Sub EndOtkupRun()
    Dim summary As String
    summary = "RunID=" & m_RunID & _
              " | Total=" & m_Total & _
              " | Passed=" & m_Passed & _
              " | Failed=" & m_Failed & _
              " | Skipped=" & m_Skipped

    Debug.Print String$(70, "-")
    Debug.Print "OTKUP TEST SUMMARY: " & summary
    Debug.Print String$(70, "-")

    AppendOtkupLog "SUITE", "SUMMARY", "INFO", summary

    If m_Failed > 0 Then
        MsgBox "Otkup tests finished with failures." & vbCrLf & summary, vbExclamation, APP_NAME
    Else
        MsgBox "Otkup tests finished." & vbCrLf & summary, vbInformation, APP_NAME
    End If
End Sub

Private Sub ResetOtkupCounters()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

Private Function NewOtkupScenario(ByVal scenarioName As String) As String
    NewOtkupScenario = DOC_PREFIX & "-" & scenarioName & "-" & m_RunID & "-" & CStr(m_Total + 1)
End Function

Private Function NextOtkupDate() As Date
    m_DateSeq = m_DateSeq + 1
    NextOtkupDate = DateSerial(2091, 1, 1) + m_DateSeq
End Function

Private Sub AssertOtkupTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        LogOtkupPass testName
    Else
        LogOtkupFail testName, "Assertion failed."
    End If
End Sub

Private Sub AssertOtkupEquals(ByVal expected As String, ByVal actual As String, ByVal testName As String)
    If CStr(expected) = CStr(actual) Then
        LogOtkupPass testName
    Else
        LogOtkupFail testName, "Expected [" & CStr(expected) & "], got [" & CStr(actual) & "]."
    End If
End Sub

Private Sub LogOtkupPass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1
    Debug.Print "[PASS] " & testName
    AppendOtkupLog "TEST", testName, "PASS", ""
End Sub

Private Sub LogOtkupFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendOtkupLog "TEST", testName, "FAIL", details
End Sub

Private Sub LogOtkupSkip(ByVal testName As String, ByVal reason As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1
    Debug.Print "[SKIP] " & testName & " :: " & reason
    AppendOtkupLog "TEST", testName, "SKIP", reason
End Sub

Private Sub LogOtkupInfo(ByVal message As String)
    Debug.Print "[INFO] " & message
    AppendOtkupLog "INFO", "", "INFO", message
End Sub

Private Sub LogOtkupFatal(ByVal sourceName As String, ByVal errNum As Long, ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
    AppendOtkupLog "FATAL", sourceName, "FAIL", CStr(errNum) & " - " & errDesc
End Sub

Private Sub InitOtkupLog()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(OTKUP_TEST_LOG)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = OTKUP_TEST_LOG
        ws.Range("A1:G1").value = Array("Timestamp", "RunID", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendOtkupLog(ByVal kindText As String, ByVal nameText As String, _
                           ByVal statusText As String, ByVal detailsText As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(OTKUP_TEST_LOG)
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
