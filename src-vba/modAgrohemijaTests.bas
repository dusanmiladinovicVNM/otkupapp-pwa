Attribute VB_Name = "modAgrohemijaTests"
Option Explicit

' ============================================================
' modAgrohemijaTests
' Dev-only smoke suite za modAgrohemija (RF-27 / AUD-040).
'
' IZOLACIJA (bez ijednog traga posle pokretanja):
'   * dev-guard: potvrda pre bilo kakve mutacije (default dugme = Ne);
'   * test-mode `modJournaling.SetTestModeQuiet True` -> WriteJournalRow i
'     MarkDirtyAndSchedule su no-op (nema CSV journal redova, nema zakazanog
'     AutoSave posle rollback-a); vraca se + StopAutoSaveTimer na kraju;
'   * ceo suite unutar JEDNE clsTransaction -> RollbackTx vraca
'     tblMagacin/tblArtikli/tblKooperanti/tblParcele;
'   * NE zove se BookPocetniDug (koji radi svoj CommitTx) -> ART_POCETNI_DUG se
'     testira direktno preko SaveMagacinCore(allowNoStock:=True);
'   * log ide SAMO u Immediate (Debug.Print) + zavrsni MsgBox -- ne pravi se
'     nikakav worksheet.
'
' Fixture: ART-TEST (cena 100) + stanje (ULAZ 1000), ART_POCETNI_DUG, KOOP-TEST,
' i parcele: PAR-OK (KOOP-TEST, aktivna), PAR-OTHER (drugi koop), PAR-INACTIVE.
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

Private Const ART_TEST As String = "ART-TEST-AGRO"
Private Const KOOP_TEST As String = "KOOP-TEST-AGRO"
Private Const KOOP_OTHER As String = "KOOP-OTHER-AGRO"
Private Const PAR_OK As String = "PAR-TEST-OK"
Private Const PAR_OTHER As String = "PAR-TEST-OTHER"
Private Const PAR_INACTIVE As String = "PAR-TEST-INACTIVE"
Private Const ART_TEST_CENA As Double = 100#

Public Sub RunAgrohemijaSmokeSuite()
    Dim tx As clsTransaction
    Dim txStarted As Boolean
    Dim quietSet As Boolean

    ' Dev-guard: sprecava slucajno pokretanje na klijentu (default = Ne).
    If MsgBox("DEV smoke test za Agrohemiju (RF-27)." & vbCrLf & _
              "Privremeno menja tblMagacin/Artikli/Kooperanti/Parcele i vraca ih " & _
              "rollback-om; journaling i AutoSave su ugaseni za vreme testa." & vbCrLf & vbCrLf & _
              "Pokrenuti?", vbYesNo + vbQuestion + vbDefaultButton2, APP_NAME) <> vbYes Then
        Exit Sub
    End If

    On Error GoTo EH

    ResetAgroTestCounters
    StartAgroSuite "AGROHEMIJA SMOKE SUITE"

    modJournaling.SetTestModeQuiet True
    quietSet = True

    Set tx = New clsTransaction
    tx.BeginTx
    txStarted = True
    tx.AddTableSnapshot TBL_MAGACIN
    tx.AddTableSnapshot TBL_ARTIKLI
    tx.AddTableSnapshot TBL_KOOPERANTI
    tx.AddTableSnapshot TBL_PARCELE

    SeedAgroFixtures

    Test_IzlazBooksSnapshotPrice
    Test_IzlazZeroPriceBlocked
    Test_IzlazNonNumericPriceBlocked
    Test_IzlazNegativePriceBlocked
    Test_NonexistentArtikalBlocked
    Test_NonexistentKooperantBlocked
    Test_IzlazParcelaOK
    Test_IzlazParcelaNonexistent
    Test_IzlazParcelaOtherKoop
    Test_IzlazParcelaInactive
    Test_IzlazParcelaMultiList
    Test_IzlazEmptyParcelaFollowsFlag
    Test_MultiItemRollback
    Test_PocetniDugArtikalExempt
    Test_UlazPositivePriceWorks
    Test_UlazZeroBlockedWithoutFlag
    Test_UlazZeroAllowedWithFlag
    Test_IzlazZeroBlockedEvenWithFlag
    Test_FormExitWiresBasketPrice

    tx.RollbackTx
    txStarted = False
    Set tx = Nothing

    CleanupAgroTestMode quietSet
    quietSet = False

    FinishAgroSuite
    Exit Sub

EH:
    LogAgroFatal "RunAgrohemijaSmokeSuite", Err.Number, Err.description
    On Error Resume Next
    If txStarted And Not tx Is Nothing Then tx.RollbackTx
    If quietSet Then CleanupAgroTestMode True
    On Error GoTo 0
    FinishAgroSuite
End Sub

Private Sub CleanupAgroTestMode(ByVal wasSet As Boolean)
    On Error Resume Next
    If wasSet Then modJournaling.SetTestModeQuiet False
    modJournaling.StopAutoSaveTimer          ' otkazi eventualno zakazan tick
    On Error GoTo 0
End Sub

' ---------- Fixtures ----------

Private Sub SeedAgroFixtures()
    SeedArtikal ART_TEST, "TEST artikal (agro)", ART_TEST_CENA
    SeedArtikal ART_POCETNI_DUG, "Pocetni dug (test)", 1#
    SeedKooperant KOOP_TEST
    SeedParcela PAR_OK, KOOP_TEST, "Da"
    SeedParcela PAR_OTHER, KOOP_OTHER, "Da"
    SeedParcela PAR_INACTIVE, KOOP_TEST, "Ne"

    ' Pocetno stanje: ULAZ 1000 @ 100 da izlazni testovi imaju pokrice.
    SaveMagacinCore Date, ART_TEST, MAG_ULAZ, 1000#, "", "", "T-SEED-ULAZ", _
                    "", "", ART_TEST_CENA
End Sub

Private Sub SeedArtikal(ByVal artID As String, ByVal naziv As String, ByVal cena As Double)
    If FindRows(TBL_ARTIKLI, COL_ART_ID, artID).count > 0 Then Exit Sub
    Dim lo As ListObject: Set lo = GetTable(TBL_ARTIKLI)
    Dim r() As Variant: ReDim r(1 To lo.ListColumns.count)
    SeedCell r, TBL_ARTIKLI, COL_ART_ID, artID
    SeedCell r, TBL_ARTIKLI, COL_ART_NAZIV, naziv
    SeedCell r, TBL_ARTIKLI, COL_ART_TIP, "Test"
    SeedCell r, TBL_ARTIKLI, COL_ART_JM, "kg"
    SeedCell r, TBL_ARTIKLI, COL_ART_CENA, cena
    SeedCell r, TBL_ARTIKLI, COL_ART_PAKOVANJE, 1
    AppendRow TBL_ARTIKLI, r
End Sub

Private Sub SeedKooperant(ByVal koopID As String)
    If FindRows(TBL_KOOPERANTI, COL_KOOP_ID, koopID).count > 0 Then Exit Sub
    Dim lo As ListObject: Set lo = GetTable(TBL_KOOPERANTI)
    Dim r() As Variant: ReDim r(1 To lo.ListColumns.count)
    SeedCell r, TBL_KOOPERANTI, COL_KOOP_ID, koopID
    AppendRow TBL_KOOPERANTI, r
End Sub

Private Sub SeedParcela(ByVal parID As String, ByVal koopID As String, ByVal aktivna As String)
    If FindRows(TBL_PARCELE, COL_PAR_ID, parID).count > 0 Then Exit Sub
    Dim lo As ListObject: Set lo = GetTable(TBL_PARCELE)
    Dim r() As Variant: ReDim r(1 To lo.ListColumns.count)
    SeedCell r, TBL_PARCELE, COL_PAR_ID, parID
    SeedCell r, TBL_PARCELE, COL_PAR_KOOP, koopID
    SeedCell r, TBL_PARCELE, COL_PAR_AKTIVNA, aktivna
    AppendRow TBL_PARCELE, r
End Sub

Private Sub SeedCell(ByRef rowData As Variant, ByVal tbl As String, _
                     ByVal colName As String, ByVal val As Variant)
    Dim idx As Long
    idx = GetColumnIndex(tbl, colName)
    If idx > 0 Then rowData(idx) = val
End Sub

Private Function CountMagRows() As Long
    Dim d As Variant
    d = GetTableData(TBL_MAGACIN)
    If IsEmpty(d) Then CountMagRows = 0 Else CountMagRows = UBound(d, 1)
End Function

' Vraca Err.Number posle poziva koji treba da padne (koristi se sa On Error Resume Next).
Private Function ErrCodeEquals(ByVal errNum As Long, ByVal code As Long) As Boolean
    ErrCodeEquals = (errNum = (vbObjectError + code))
End Function

' ---------- Testovi: cena ----------

Private Sub Test_IzlazBooksSnapshotPrice()
    ' #1 + #2: master cena = 100, snapshot korpe = 80 -> knjizi se 80.
    On Error GoTo EH
    Dim id As String
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 2#, KOOP_TEST, PAR_OK, "T-SNAP", _
                         overrideCena:=80#)
    If Len(Trim$(id)) = 0 Then
        LogAgroFail "Izlaz knjizi snapshot cenu", "Nema ID-a (upis nije uspeo)"
        Exit Sub
    End If
    AssertAgroDoubleEquals 80#, _
        CDbl(LookupValue(TBL_MAGACIN, COL_MAG_ID, id, COL_MAG_CENA)), _
        "Izlaz Cena = snapshot korpe (80, ne master 100)"
    AssertAgroDoubleEquals 160#, _
        CDbl(LookupValue(TBL_MAGACIN, COL_MAG_ID, id, COL_MAG_VREDNOST)), _
        "Izlaz Vrednost = kolicina * snapshot (2 * 80 = 160)"
    Exit Sub
EH:
    LogAgroFail "Izlaz knjizi snapshot cenu", Err.description
End Sub

Private Sub Test_IzlazZeroPriceBlocked()
    ' #3: overrideCena = 0 -> fail-closed (Err 4206), bez upisa.
    Dim id As String, en As Long
    On Error Resume Next
    Err.Clear
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-IZ0", _
                         overrideCena:=0#)
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4206) And Len(Trim$(id)) = 0, _
        "Izlaz cena=0 -> blokirano (Err 4206, bez upisa)"
End Sub

Private Sub Test_IzlazNonNumericPriceBlocked()
    ' #4: nenumericka cena -> resolves 0 -> Err 4206.
    Dim id As String, en As Long
    On Error Resume Next
    Err.Clear
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-IZNAN", _
                         overrideCena:="nije broj")
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4206) And Len(Trim$(id)) = 0, _
        "Izlaz nenumericka cena -> blokirano (Err 4206)"
End Sub

Private Sub Test_IzlazNegativePriceBlocked()
    ' #5: negativna cena -> Err 4206.
    Dim id As String, en As Long
    On Error Resume Next
    Err.Clear
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-IZNEG", _
                         overrideCena:=-5#)
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4206) And Len(Trim$(id)) = 0, _
        "Izlaz negativna cena -> blokirano (Err 4206)"
End Sub

' ---------- Testovi: referencijalni integritet ----------

Private Sub Test_NonexistentArtikalBlocked()
    ' #6: nepostojeci artikal -> Err 4207.
    Dim en As Long
    On Error Resume Next
    Err.Clear
    SaveMagacinCore Date, "ART-NE-POSTOJI", MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-NOART", _
                    overrideCena:=50#
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4207), "Nepostojeci artikal -> blokirano (Err 4207)"
End Sub

Private Sub Test_NonexistentKooperantBlocked()
    ' #7: nepostojeci kooperant (izlaz) -> Err 4208.
    Dim en As Long
    On Error Resume Next
    Err.Clear
    SaveMagacinCore Date, ART_TEST, MAG_IZLAZ, 1#, "KOOP-NE-POSTOJI", "", "T-NOKOOP", _
                    overrideCena:=50#
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4208), "Nepostojeci kooperant (izlaz) -> blokirano (Err 4208)"
End Sub

Private Sub Test_IzlazParcelaOK()
    ' Parcela pripada kooperantu i aktivna -> prolazi.
    On Error GoTo EH
    Dim id As String
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-PAROK", _
                         overrideCena:=100#)
    AssertAgroTrue Len(Trim$(id)) > 0, "Izlaz sa validnom parcelom (koop+aktivna) -> prolazi"
    Exit Sub
EH:
    LogAgroFail "Izlaz sa validnom parcelom", Err.description
End Sub

Private Sub Test_IzlazParcelaNonexistent()
    Dim en As Long
    On Error Resume Next
    Err.Clear
    SaveMagacinCore Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "PAR-NE-POSTOJI", "T-PARNE", _
                    overrideCena:=100#
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4212), "Nepostojeca parcela -> blokirano (Err 4212)"
End Sub

Private Sub Test_IzlazParcelaOtherKoop()
    ' Parcela drugog kooperanta -> blokirano (jezgro nalaza).
    Dim en As Long
    On Error Resume Next
    Err.Clear
    SaveMagacinCore Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OTHER, "T-PAROTH", _
                    overrideCena:=100#
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4213), _
        "Parcela drugog kooperanta -> blokirano (Err 4213)"
End Sub

Private Sub Test_IzlazParcelaInactive()
    ' Neaktivna parcela -> blokirano.
    Dim en As Long
    On Error Resume Next
    Err.Clear
    SaveMagacinCore Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_INACTIVE, "T-PARIN", _
                    overrideCena:=100#
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4214), "Neaktivna parcela -> blokirano (Err 4214)"
End Sub

Private Sub Test_IzlazParcelaMultiList()
    ' ;-lista: prva OK, druga drugog koop -> blokirano na drugoj (Err 4213).
    Dim en As Long
    On Error Resume Next
    Err.Clear
    SaveMagacinCore Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK & ";" & PAR_OTHER, _
                    "T-PARML", overrideCena:=100#
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4213), _
        "Parcela ;-lista sa jednom tudjom -> blokirano (Err 4213)"
End Sub

Private Sub Test_IzlazEmptyParcelaFollowsFlag()
    ' Prazna parcela: PRACENJE ON -> blokirano (Err 4215); OFF -> dozvoljeno.
    Dim id As String, en As Long
    On Error Resume Next
    Err.Clear
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-IZEMPTY", _
                         overrideCena:=100#)
    en = Err.Number
    On Error GoTo 0
    If IsPracenjeParcela() Then
        AssertAgroTrue ErrCodeEquals(en, 4215) And Len(Trim$(id)) = 0, _
            "Izlaz bez parcele + PRACENJE ON -> blokirano (Err 4215)"
    Else
        AssertAgroTrue (en = 0) And Len(Trim$(id)) > 0, _
            "Izlaz bez parcele + PRACENJE OFF -> dozvoljeno"
    End If
End Sub

' ---------- Testovi: transakcija / migracija / ulaz ----------

Private Sub Test_MultiItemRollback()
    ' #8: emulira formu -- outer TX oko vise stavki; pad druge vraca prvu.
    On Error GoTo EH
    Dim tx As clsTransaction
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_MAGACIN

    Dim before As Long: before = CountMagRows()

    Dim id1 As String
    id1 = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-RB1", _
                          overrideCena:=100#)          ' stavka 1 OK

    Dim en As Long
    On Error Resume Next
    Err.Clear
    SaveMagacinCore Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-RB2", _
                    overrideCena:=0#                   ' stavka 2 pada (cena 0)
    en = Err.Number
    On Error GoTo 0

    If en <> 0 Then tx.RollbackTx                      ' forma bi ovde rollback-ovala

    Dim afterCnt As Long: afterCnt = CountMagRows()
    AssertAgroTrue (en <> 0) And (Len(Trim$(id1)) > 0) And (afterCnt = before), _
        "Multi-stavka rollback: pad druge stavke vraca i prvu (nema delimicnog upisa)"
    Set tx = Nothing
    Exit Sub
EH:
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogAgroFail "Multi-stavka rollback", Err.description
End Sub

Private Sub Test_PocetniDugArtikalExempt()
    ' #9: ART_POCETNI_DUG izuzet iz cena/parcela/ref provera. Testira se DIREKTNO
    ' preko SaveMagacinCore (BEZ BookPocetniDug -> bez unutrasnjeg CommitTx/autosave).
    On Error GoTo EH
    Dim id As String
    id = SaveMagacinCore(Date, ART_POCETNI_DUG, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-POCDUG", _
                         "Pocetni dug (test)", "", 500#, allowNoStock:=True)
    AssertAgroTrue Len(Trim$(id)) > 0, _
        "ART_POCETNI_DUG se knjizi (izuzet iz fail-closed/parcela/ref provera)"
    If Len(Trim$(id)) > 0 Then
        AssertAgroDoubleEquals 500#, _
            CDbl(LookupValue(TBL_MAGACIN, COL_MAG_ID, id, COL_MAG_VREDNOST)), _
            "Pocetni dug Vrednost = iznos (500)"
    End If
    Exit Sub
EH:
    LogAgroFail "Pocetni dug artikal izuzet", Err.description
End Sub

Private Sub Test_UlazPositivePriceWorks()
    ' #10: legitiman ulaz sa cenom > 0.
    On Error GoTo EH
    Dim id As String
    id = SaveMagacinCore(Date, ART_TEST, MAG_ULAZ, 10#, "", "", "T-ULOK", _
                         "", "DOB-TEST", 100#)
    AssertAgroTrue Len(Trim$(id)) > 0, "Ulaz sa cenom > 0 radi"
    If Len(Trim$(id)) > 0 Then
        AssertAgroDoubleEquals 100#, _
            CDbl(LookupValue(TBL_MAGACIN, COL_MAG_ID, id, COL_MAG_CENA)), _
            "Ulaz Cena = uneta (100)"
    End If
    Exit Sub
EH:
    LogAgroFail "Ulaz cena>0 radi", Err.description
End Sub

Private Sub Test_UlazZeroBlockedWithoutFlag()
    ' Zero-value ULAZ bez flag-a ostaje blokiran (Err 4206).
    Dim id As String, en As Long
    On Error Resume Next
    Err.Clear
    id = SaveMagacinCore(Date, ART_TEST, MAG_ULAZ, 1#, "", "", "T-UL0NF", _
                         "", "", 0#)
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4206) And Len(Trim$(id)) = 0, _
        "Ulaz cena=0 bez allowZeroValue -> blokirano (Err 4206)"
End Sub

Private Sub Test_UlazZeroAllowedWithFlag()
    ' Nova staza: dokumentovan besplatan/korektivni ULAZ (cena 0) uz flag.
    On Error GoTo EH
    Dim id As String
    id = SaveMagacinCore(Date, ART_TEST, MAG_ULAZ, 5#, "", "", "T-UL0F", _
                         "", "", 0#, allowZeroValue:=True)
    AssertAgroTrue Len(Trim$(id)) > 0, _
        "Ulaz cena=0 uz allowZeroValue -> dozvoljeno (besplatan prijem)"
    If Len(Trim$(id)) > 0 Then
        AssertAgroDoubleEquals 0#, _
            CDbl(LookupValue(TBL_MAGACIN, COL_MAG_ID, id, COL_MAG_VREDNOST)), _
            "Besplatan ulaz Vrednost = 0"
    End If
    Exit Sub
EH:
    LogAgroFail "Ulaz cena=0 uz flag radi", Err.description
End Sub

Private Sub Test_IzlazZeroBlockedEvenWithFlag()
    ' allowZeroValue NE vazi za izlaz -- izlaz sa cenom 0 ostaje blokiran (Err 4206).
    Dim id As String, en As Long
    On Error Resume Next
    Err.Clear
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, PAR_OK, "T-IZ0F", _
                         "", "", 0#, allowZeroValue:=True)
    en = Err.Number
    On Error GoTo 0
    AssertAgroTrue ErrCodeEquals(en, 4206) And Len(Trim$(id)) = 0, _
        "Izlaz cena=0 ostaje blokiran i uz allowZeroValue (Err 4206)"
End Sub

Private Sub Test_FormExitWiresBasketPrice()
    ' Cuva AUD-040 na UI granici: forma MORA proslediti korpa cenu kao overrideCena.
    ' SaveMagacinCore-unit test to NE hvata (bug je bio u formi, ne u jezgru). Cita
    ' izvor frmAgrohemija preko VBProject-a (treba "Trust access to VBA project object
    ' model"); ako pristup nije dozvoljen -> SKIP (ne FAIL).
    Dim src As String
    On Error GoTo NoAccess
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents("frmAgrohemija").CodeModule
    src = cm.Lines(1, cm.CountOfLines)
    On Error GoTo 0

    AssertAgroTrue InStr(src, "overrideCena:=m_KorpaIzlaz(i).cena") > 0, _
        "Forma izlaz prosledjuje korpa cenu kao overrideCena (AUD-040 wiring)"
    AssertAgroTrue InStr(src, "allowZeroValue:=m_KorpaUlaz(i).allowZeroValue") > 0, _
        "Forma ulaz prosledjuje allowZeroValue korpe"
    AssertAgroTrue InStr(src, "SaveMagacinCore(") > 0, _
        "Forma zove SaveMagacinCore (typed greske stizu do UI)"
    Exit Sub
NoAccess:
    LogAgroSkip "Form wiring provera (nema 'Trust access to VBA project object model')"
End Sub

' ---------- Assert / log infrastruktura (bez worksheet-a) ----------

Private Sub AssertAgroTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then LogAgroPass testName Else LogAgroFail testName, "Assertion failed."
End Sub

Private Sub AssertAgroDoubleEquals(ByVal expected As Double, ByVal actual As Double, _
                                   ByVal testName As String)
    If Abs(expected - actual) < 0.000001 Then
        LogAgroPass testName
    Else
        LogAgroFail testName, "Expected=" & CStr(expected) & " Actual=" & CStr(actual)
    End If
End Sub

Private Sub ResetAgroTestCounters()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

Private Sub StartAgroSuite(ByVal suiteName As String)
    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(70, "=")
End Sub

Private Sub FinishAgroSuite()
    Dim summary As String
    summary = "Total=" & m_Total & _
              " | Passed=" & m_Passed & _
              " | Failed=" & m_Failed & _
              " | Skipped=" & m_Skipped

    Debug.Print String$(70, "-")
    Debug.Print "AGRO TEST SUMMARY: " & summary
    Debug.Print String$(70, "-")

    If m_Failed > 0 Then
        MsgBox "Agrohemija tests finished with failures." & vbCrLf & summary, _
               vbExclamation, APP_NAME
    Else
        MsgBox "Agrohemija tests finished." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If
End Sub

Private Sub LogAgroPass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1
    Debug.Print "[PASS] " & testName
End Sub

Private Sub LogAgroFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FAIL] " & testName & " :: " & details
End Sub

Private Sub LogAgroSkip(ByVal testName As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1
    Debug.Print "[SKIP] " & testName
End Sub

Private Sub LogAgroFatal(ByVal sourceName As String, _
                         ByVal errNum As Long, _
                         ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
End Sub
