Attribute VB_Name = "modAgrohemijaTests"
Option Explicit

' ============================================================
' modAgrohemijaTests
' Dev/test-only smoke suite for modAgrohemija (RF-27 / AUD-040).
'
' Izolacija: ceo suite radi unutar JEDNE clsTransaction koja snima
' tblMagacin/tblArtikli/tblKooperanti na pocetku i RADI ROLLBACK na kraju,
' pa testovi NE ostavljaju trag u ledgeru. Fixture: ART-TEST (cena 100) +
' KOOP-TEST + pocetno stanje (ULAZ 1000) za pokrice izlaza.
'
' Pokriva nalaze iz review-a RF-27:
'   1  izlaz knjizi snapshot cenu (ne master)      -> Test_IzlazBooksSnapshotPrice
'   2  master se promeni posle korpe -> snapshot    -> isti (overrideCena pobedjuje master)
'   3  overrideCena = 0 -> nema reda (izlaz)        -> Test_IzlazZeroPriceBlocked
'   4  overrideCena nije broj -> nema reda          -> Test_IzlazNonNumericPriceBlocked
'   5  negativna cena -> blokirano                  -> Test_IzlazNegativePriceBlocked
'   6  nepostojeci artikal -> blokirano             -> Test_NonexistentArtikalBlocked
'   7  nepostojeci kooperant -> blokirano           -> Test_NonexistentKooperantBlocked
'   8  pad druge stavke -> rollback prve            -> Test_MultiItemRollback
'   9  ART_POCETNI_DUG i dalje radi                 -> Test_PocetniDugStillWorks
'   10 legitiman ulaz sa cenom > 0 radi             -> Test_UlazPositivePriceWorks
'   11 deaktiviran artikal/koop = N/A: tblArtikli/tblKooperanti NEMAJU
'      "aktivan" kolonu u semi (samo tblParcele/tblKorisnici je imaju), pa
'      "postoji" je jedini lifecycle koji model danas ima.
'   12 parcela drugog kooperanta = dokumentovan gap (AUD-049): parcela<->koop
'      veza jos nije proverena; ostaje otvoreno dok se AUD-049 ne uradi.
' + nova zero-value ULAZ staza (allowZeroValue): dozvoljen uz flag, izlaz strog.
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

Private Const AGRO_TEST_LOG_SHEET As String = "AGRO_TEST_LOG"
Private Const ART_TEST As String = "ART-TEST-AGRO"
Private Const KOOP_TEST As String = "KOOP-TEST-AGRO"
Private Const ART_TEST_CENA As Double = 100#

Public Sub RunAgrohemijaSmokeSuite()
    Dim tx As clsTransaction
    Dim txStarted As Boolean

    On Error GoTo EH

    ResetAgroTestCounters
    InitAgroTestLog
    StartAgroSuite "AGROHEMIJA SMOKE SUITE"

    Set tx = New clsTransaction
    tx.BeginTx
    txStarted = True
    tx.AddTableSnapshot TBL_MAGACIN
    tx.AddTableSnapshot TBL_ARTIKLI
    tx.AddTableSnapshot TBL_KOOPERANTI

    SeedAgroFixtures

    Test_IzlazBooksSnapshotPrice
    Test_IzlazZeroPriceBlocked
    Test_IzlazNonNumericPriceBlocked
    Test_IzlazNegativePriceBlocked
    Test_NonexistentArtikalBlocked
    Test_NonexistentKooperantBlocked
    Test_MultiItemRollback
    Test_PocetniDugStillWorks
    Test_UlazPositivePriceWorks
    Test_UlazZeroBlockedWithoutFlag
    Test_UlazZeroAllowedWithFlag
    Test_IzlazZeroBlockedEvenWithFlag

    ' Vrati sve -- testovi ne ostavljaju trag u ledgeru.
    tx.RollbackTx
    txStarted = False
    Set tx = Nothing

    FinishAgroSuite
    Exit Sub

EH:
    LogAgroFatal "RunAgrohemijaSmokeSuite", Err.Number, Err.description
    On Error Resume Next
    If txStarted And Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    FinishAgroSuite
End Sub

' ---------- Fixtures ----------

Private Sub SeedAgroFixtures()
    ' Artikal ART-TEST (cena 100) -- upis po imenu kolone (redosled nije siguran).
    If FindRows(TBL_ARTIKLI, COL_ART_ID, ART_TEST).count = 0 Then
        Dim aLo As ListObject: Set aLo = GetTable(TBL_ARTIKLI)
        Dim aRow() As Variant: ReDim aRow(1 To aLo.ListColumns.count)
        SeedCell aRow, TBL_ARTIKLI, COL_ART_ID, ART_TEST
        SeedCell aRow, TBL_ARTIKLI, COL_ART_NAZIV, "TEST artikal (agro)"
        SeedCell aRow, TBL_ARTIKLI, COL_ART_TIP, "Test"
        SeedCell aRow, TBL_ARTIKLI, COL_ART_JM, "kg"
        SeedCell aRow, TBL_ARTIKLI, COL_ART_CENA, ART_TEST_CENA
        SeedCell aRow, TBL_ARTIKLI, COL_ART_PAKOVANJE, 1
        AppendRow TBL_ARTIKLI, aRow
    End If

    ' Kooperant KOOP-TEST -- za referencijalnu proveru dovoljan je ID.
    If FindRows(TBL_KOOPERANTI, COL_KOOP_ID, KOOP_TEST).count = 0 Then
        Dim kLo As ListObject: Set kLo = GetTable(TBL_KOOPERANTI)
        Dim kRow() As Variant: ReDim kRow(1 To kLo.ListColumns.count)
        SeedCell kRow, TBL_KOOPERANTI, COL_KOOP_ID, KOOP_TEST
        AppendRow TBL_KOOPERANTI, kRow
    End If

    ' Pocetno stanje: ULAZ 1000 @ 100 da izlazni testovi imaju pokrice.
    SaveMagacinCore Date, ART_TEST, MAG_ULAZ, 1000#, "", "", "T-SEED-ULAZ", _
                    "", "", ART_TEST_CENA
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

' ---------- Testovi ----------

Private Sub Test_IzlazBooksSnapshotPrice()
    ' #1 + #2: master cena = 100, snapshot korpe = 80 -> knjizi se 80.
    On Error GoTo EH
    Dim id As String
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 2#, KOOP_TEST, "", "T-SNAP", _
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
    ' #3: overrideCena = 0 -> fail-closed, nema reda.
    Dim id As String, raised As Boolean
    On Error Resume Next
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-IZ0", _
                         overrideCena:=0#)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertAgroTrue raised And Len(Trim$(id)) = 0, _
        "Izlaz cena=0 -> blokirano (Err.Raise, bez upisa)"
End Sub

Private Sub Test_IzlazNonNumericPriceBlocked()
    ' #4: nenumericka cena -> resolves 0 -> blokirano.
    Dim id As String, raised As Boolean
    On Error Resume Next
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-IZNAN", _
                         overrideCena:="nije broj")
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertAgroTrue raised And Len(Trim$(id)) = 0, _
        "Izlaz nenumericka cena -> blokirano (bez upisa)"
End Sub

Private Sub Test_IzlazNegativePriceBlocked()
    ' #5: negativna cena -> blokirano.
    Dim id As String, raised As Boolean
    On Error Resume Next
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-IZNEG", _
                         overrideCena:=-5#)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertAgroTrue raised And Len(Trim$(id)) = 0, _
        "Izlaz negativna cena -> blokirano"
End Sub

Private Sub Test_NonexistentArtikalBlocked()
    ' #6: nepostojeci artikal -> blokirano (4207).
    Dim id As String, raised As Boolean
    On Error Resume Next
    id = SaveMagacinCore(Date, "ART-NE-POSTOJI", MAG_IZLAZ, 1#, KOOP_TEST, "", "T-NOART", _
                         overrideCena:=50#)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertAgroTrue raised And Len(Trim$(id)) = 0, _
        "Nepostojeci artikal -> blokirano (4207)"
End Sub

Private Sub Test_NonexistentKooperantBlocked()
    ' #7: nepostojeci kooperant (izlaz) -> blokirano (4208).
    Dim id As String, raised As Boolean
    On Error Resume Next
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, "KOOP-NE-POSTOJI", "", "T-NOKOOP", _
                         overrideCena:=50#)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertAgroTrue raised And Len(Trim$(id)) = 0, _
        "Nepostojeci kooperant (izlaz) -> blokirano (4208)"
End Sub

Private Sub Test_MultiItemRollback()
    ' #8: emulira formu -- outer TX oko vise stavki; pad druge vraca prvu.
    On Error GoTo EH
    Dim tx As clsTransaction
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_MAGACIN

    Dim before As Long: before = CountMagRows()

    Dim id1 As String
    id1 = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-RB1", _
                          overrideCena:=100#)               ' stavka 1 OK

    Dim raised As Boolean
    On Error Resume Next
    SaveMagacinCore Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-RB2", _
                    overrideCena:=0#                        ' stavka 2 pada (cena 0)
    raised = (Err.Number <> 0)
    On Error GoTo 0

    If raised Then tx.RollbackTx                            ' forma bi ovde rollback-ovala

    Dim afterCnt As Long: afterCnt = CountMagRows()
    AssertAgroTrue raised And (Len(Trim$(id1)) > 0) And (afterCnt = before), _
        "Multi-stavka rollback: pad druge stavke vraca i prvu (nema delimicnog upisa)"
    Set tx = Nothing
    Exit Sub
EH:
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    LogAgroFail "Multi-stavka rollback", Err.description
End Sub

Private Sub Test_PocetniDugStillWorks()
    ' #9: ART_POCETNI_DUG migracija -- nije blokiran ni cenom ni referencijalno.
    On Error GoTo EH
    Dim id As String
    id = BookPocetniDug(KOOP_TEST, 500#, "T-POCDUG", Date)
    AssertAgroTrue Len(Trim$(id)) > 0, "Pocetni dug (ART_POCETNI_DUG) i dalje radi"
    If Len(Trim$(id)) > 0 Then
        AssertAgroDoubleEquals 500#, _
            CDbl(LookupValue(TBL_MAGACIN, COL_MAG_ID, id, COL_MAG_VREDNOST)), _
            "Pocetni dug Vrednost = iznos (500)"
    End If
    Exit Sub
EH:
    LogAgroFail "Pocetni dug radi", Err.description
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
    ' Zero-value ULAZ bez flag-a ostaje blokiran (default fail-closed).
    Dim id As String, raised As Boolean
    On Error Resume Next
    id = SaveMagacinCore(Date, ART_TEST, MAG_ULAZ, 1#, "", "", "T-UL0NF", _
                         "", "", 0#)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertAgroTrue raised And Len(Trim$(id)) = 0, _
        "Ulaz cena=0 bez allowZeroValue -> blokirano"
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
    ' allowZeroValue NE vazi za izlaz -- izlaz sa cenom 0 ostaje blokiran.
    Dim id As String, raised As Boolean
    On Error Resume Next
    id = SaveMagacinCore(Date, ART_TEST, MAG_IZLAZ, 1#, KOOP_TEST, "", "T-IZ0F", _
                         "", "", 0#, allowZeroValue:=True)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertAgroTrue raised And Len(Trim$(id)) = 0, _
        "Izlaz cena=0 ostaje blokiran i uz allowZeroValue (izlaz strog)"
End Sub

' ---------- Assert / log infrastruktura ----------

Private Sub AssertAgroTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        LogAgroPass testName
    Else
        LogAgroFail testName, "Assertion failed."
    End If
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
    AppendAgroTestLog "SUITE", suiteName, "START", ""
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

    AppendAgroTestLog "SUITE", "SUMMARY", "INFO", summary

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
    AppendAgroTestLog "TEST", testName, "PASS", ""
End Sub

Private Sub LogAgroFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendAgroTestLog "TEST", testName, "FAIL", details
End Sub

Private Sub LogAgroFatal(ByVal sourceName As String, _
                         ByVal errNum As Long, _
                         ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1
    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
    AppendAgroTestLog "FATAL", sourceName, "FAIL", CStr(errNum) & " - " & errDesc
End Sub

Private Sub InitAgroTestLog()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(AGRO_TEST_LOG_SHEET)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = AGRO_TEST_LOG_SHEET
        ws.Range("A1:F1").value = Array("Timestamp", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendAgroTestLog(ByVal kindText As String, _
                              ByVal nameText As String, _
                              ByVal statusText As String, _
                              ByVal detailsText As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(AGRO_TEST_LOG_SHEET)
    If ws Is Nothing Then Exit Sub

    Dim r As Long
    r = ws.cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.cells(r, 1).value = Now
    ws.cells(r, 2).value = kindText
    ws.cells(r, 3).value = nameText
    ws.cells(r, 4).value = statusText
    ws.cells(r, 5).value = Left$(detailsText, 2000)
    ws.cells(r, 6).value = Environ$("Username")
End Sub
