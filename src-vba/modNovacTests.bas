Option Explicit

' ============================================================
' modNovacTests
' Dev/test-only smoke suite for modNovac.
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

Private Const NOVAC_TEST_LOG_SHEET As String = "NOVAC_TEST_LOG"

Public Sub RunNovacSmokeSuite()
    On Error GoTo EH

    ResetNovacTestCounters
    InitNovacTestLog

    StartNovacSuite "NOVAC SMOKE SUITE"

    Test_SaveNovacRejectsInvalidAmounts
    Test_SaveNovacAcceptsValidUplata
    Test_SaveNovacAcceptsValidIsplata
    Test_StorniranoNovacExcludedFromFakturaUplata
    Test_PartnerMapConflictBlocked
    
    Test_PartialBuyerAvansSplitFaktura
    Test_PartialOtkupAvansSplit
    Test_ResetNovacOtkupLinkRecomputesStatus

    FinishNovacSuite
    Exit Sub

EH:
    LogNovacFatal "RunNovacSmokeSuite", Err.Number, Err.description
    FinishNovacSuite
End Sub

Private Sub Test_SaveNovacRejectsInvalidAmounts()
    On Error GoTo ExpectedError

    Dim id As String
    id = SaveNovac( _
        "TST-NOV-INVALID", Date, _
        "TEST PARTNER", "KUP-TEST", "Kupac", _
        "", "", "", "", _
        NOV_KUPCI_UPLATA, _
        100#, 50#, _
        "Invalid both uplata and isplata")

    LogNovacFail "SaveNovac rejects both uplata and isplata", _
                 "Expected validation error, got ID=" & id
    Exit Sub

ExpectedError:
    LogNovacPass "SaveNovac rejects both uplata and isplata"
End Sub

Private Sub Test_SaveNovacAcceptsValidUplata()
    On Error GoTo EH

    Dim id As String
    id = SaveNovac( _
        "TST-NOV-UPLATA-" & Format$(Now, "yyyymmddhhnnss"), Date, _
        "TEST PARTNER", "KUP-TEST", "Kupac", _
        "", "", "", "", _
        NOV_KUPCI_UPLATA, _
        10#, 0#, _
        "Test valid uplata")

    AssertNovacTrue Len(Trim$(id)) > 0, "SaveNovac accepts valid uplata"
    Exit Sub

EH:
    LogNovacFail "SaveNovac accepts valid uplata", Err.description
End Sub

Private Sub Test_SaveNovacAcceptsValidIsplata()
    On Error GoTo EH

    Dim id As String
    id = SaveNovac( _
        "TST-NOV-ISPLATA-" & Format$(Now, "yyyymmddhhnnss"), Date, _
        "TEST KOOPERANT", "KOOP-TEST", "Kooperant", _
        "", "KOOP-TEST", "", "", _
        NOV_VIRMAN_FIRMA_KOOP, _
        0#, 10#, _
        "Test valid isplata")

    AssertNovacTrue Len(Trim$(id)) > 0, "SaveNovac accepts valid isplata"
    Exit Sub

EH:
    LogNovacFail "SaveNovac accepts valid isplata", Err.description
End Sub

Private Sub Test_StorniranoNovacExcludedFromFakturaUplata()
    On Error GoTo EH

    Const SRC As String = "Test_StorniranoNovacExcludedFromFakturaUplata"

    Dim fakturaID As String
    fakturaID = "TST-FAK-NOVAC-" & Format$(Now, "yyyymmddhhnnss")

    Dim activeID As String
    Dim stornID As String

    activeID = SaveNovac( _
        "TST-UPL-ACT-" & Format$(Now, "yyyymmddhhnnss"), Date, _
        "TEST KUPAC", "KUP-TEST", "Kupac", _
        "", "", fakturaID, "", _
        NOV_KUPCI_UPLATA, _
        100#, 0#, _
        "Active test uplata")

    stornID = SaveNovac( _
        "TST-UPL-STO-" & Format$(Now, "yyyymmddhhnnss"), Date, _
        "TEST KUPAC", "KUP-TEST", "Kupac", _
        "", "", fakturaID, "", _
        NOV_KUPCI_UPLATA, _
        999#, 0#, _
        "Stornirano test uplata")

    If Len(activeID) = 0 Or Len(stornID) = 0 Then
        LogNovacFail "Stornirano novac excluded from faktura uplata", _
                     "Failed to create test novac rows."
        Exit Sub
    End If

    Dim rows As Collection
    Set rows = FindRows(TBL_NOVAC, COL_NOV_ID, stornID)

    If rows Is Nothing Or rows.count = 0 Then
        LogNovacFail "Stornirano novac excluded from faktura uplata", _
                     "Could not find stornirano test row."
        Exit Sub
    End If

    RequireUpdateCell TBL_NOVAC, rows(1), COL_STORNIRANO, "Da", SRC

    Dim total As Double
    total = GetUplataForFaktura(fakturaID)

    If Abs(total - 100#) < 0.0001 Then
        LogNovacPass "Stornirano novac excluded from faktura uplata"
    Else
        LogNovacFail "Stornirano novac excluded from faktura uplata", _
                     "Expected 100, got " & CStr(total)
    End If

    Exit Sub

EH:
    LogNovacFail "Stornirano novac excluded from faktura uplata", _
                 "Err.Number=" & CStr(Err.Number) & _
                 " Source=" & Err.SOURCE & _
                 " Description=" & Err.description
End Sub

Private Sub Test_PartnerMapConflictBlocked()
    On Error GoTo EH

    Dim bankaName As String
    bankaName = "TST BANKA MAP " & Format$(Now, "yyyymmddhhnnss")

    Dim ok As Boolean
    ok = savePartnerMap(bankaName, "KUP-TEST-1", "Kupac", "")

    If Not ok Then
        LogNovacFail "PartnerMap initial save", "savePartnerMap returned False."
        Exit Sub
    End If

    On Error GoTo ExpectedConflict

    ok = savePartnerMap(bankaName, "KUP-TEST-2", "Kupac", "")

    LogNovacFail "PartnerMap conflict blocked", _
                 "Expected conflict error, but savePartnerMap returned " & CStr(ok)
    Exit Sub

ExpectedConflict:
    LogNovacPass "PartnerMap conflict blocked"
    Exit Sub

EH:
    LogNovacFail "PartnerMap conflict blocked", Err.description
End Sub

Private Sub AssertNovacTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        LogNovacPass testName
    Else
        LogNovacFail testName, "Assertion failed."
    End If
End Sub

Private Sub ResetNovacTestCounters()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

Private Sub StartNovacSuite(ByVal suiteName As String)
    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(70, "=")

    AppendNovacTestLog "SUITE", suiteName, "START", ""
End Sub

Private Sub FinishNovacSuite()
    Dim summary As String

    summary = "Total=" & m_Total & _
              " | Passed=" & m_Passed & _
              " | Failed=" & m_Failed & _
              " | Skipped=" & m_Skipped

    Debug.Print String$(70, "-")
    Debug.Print "NOVAC TEST SUMMARY: " & summary
    Debug.Print String$(70, "-")

    AppendNovacTestLog "SUITE", "SUMMARY", "INFO", summary

    If m_Failed > 0 Then
        MsgBox "Novac tests finished with failures." & vbCrLf & summary, _
               vbExclamation, APP_NAME
    Else
        MsgBox "Novac tests finished." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If
End Sub

Private Sub LogNovacPass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1

    Debug.Print "[PASS] " & testName
    AppendNovacTestLog "TEST", testName, "PASS", ""
End Sub

Private Sub LogNovacFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendNovacTestLog "TEST", testName, "FAIL", details
End Sub

Private Sub LogNovacFatal(ByVal sourceName As String, _
                          ByVal errNum As Long, _
                          ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
    AppendNovacTestLog "FATAL", sourceName, "FAIL", CStr(errNum) & " - " & errDesc
End Sub

Private Sub InitNovacTestLog()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NOVAC_TEST_LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = NOVAC_TEST_LOG_SHEET
        ws.Range("A1:F1").value = Array("Timestamp", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendNovacTestLog(ByVal kindText As String, _
                               ByVal nameText As String, _
                               ByVal statusText As String, _
                               ByVal detailsText As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NOVAC_TEST_LOG_SHEET)

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


Private Sub Test_PartialBuyerAvansSplitFaktura()
    On Error GoTo EH

    Dim code As String
    code = NewNovacTestCode("BUYAVSPLIT")

    Dim kupacID As String
    Dim fakturaID As String
    Dim avansID As String

    kupacID = "KUP-TST-" & code
    fakturaID = "FAK-TST-" & code

    AppendTestFakturaRow fakturaID, "TST-FAK-" & code, kupacID, 100#

    avansID = SaveNovac( _
        "TST-AVANS-" & code, Date, _
        "TEST KUPAC " & code, kupacID, "Kupac", _
        "", "", "", "", _
        NOV_KUPCI_AVANS, _
        150#, 0#, _
        "Test buyer partial avans")

    If Len(Trim$(avansID)) = 0 Then
        LogNovacFail "Partial buyer avans split faktura", _
                     "Failed to create avans row."
        Exit Sub
    End If

    If Not ApplyAvansToFaktura_TX(kupacID, fakturaID) Then
        LogNovacFail "Partial buyer avans split faktura", _
                     "ApplyAvansToFaktura_TX returned False."
        Exit Sub
    End If

    Dim originalLeft As Double
    originalLeft = CDbl(LookupValue(TBL_NOVAC, COL_NOV_ID, avansID, COL_NOV_UPLATA))

    AssertNovacDoubleEquals 50#, originalLeft, _
                            "Partial buyer avans original row reduced to remainder"

    Dim totalApplied As Double
    totalApplied = GetUplataForFaktura(fakturaID)

    AssertNovacDoubleEquals 100#, totalApplied, _
                            "Partial buyer avans applied amount linked to faktura"

    Dim splitCount As Long
    splitCount = CountNovacRowsForFaktura(fakturaID)

    AssertNovacTrue splitCount >= 1, _
                    "Partial buyer avans created linked split row"

    Exit Sub

EH:
    LogNovacFail "Partial buyer avans split faktura", _
                 "Err.Number=" & CStr(Err.Number) & _
                 " Source=" & Err.SOURCE & _
                 " Description=" & Err.description
End Sub

Private Sub Test_PartialOtkupAvansSplit()
    On Error GoTo EH

    Dim code As String
    code = NewNovacTestCode("OTKAVSPLIT")

    Dim kooperantID As String
    Dim otkupID As String
    Dim avansID As String

    kooperantID = "KOOP-TST-" & code
    otkupID = "OTK-TST-" & code

    AppendTestOtkupRow otkupID, kooperantID, 100#, 1#

    avansID = SaveNovac( _
        "TST-OTK-AVANS-" & code, Date, _
        "TEST KOOPERANT " & code, kooperantID, "Kooperant", _
        "", kooperantID, "", "", _
        NOV_VIRMAN_AVANS_KOOP, _
        0#, 150#, _
        "Test otkup partial avans")

    If Len(Trim$(avansID)) = 0 Then
        LogNovacFail "Partial otkup avans split", _
                     "Failed to create avans row."
        Exit Sub
    End If

    If Not ApplyAvansToOtkup_TX(kooperantID, otkupID) Then
        LogNovacFail "Partial otkup avans split", _
                     "ApplyAvansToOtkup_TX returned False."
        Exit Sub
    End If

    Dim originalLeft As Double
    originalLeft = CDbl(LookupValue(TBL_NOVAC, COL_NOV_ID, avansID, COL_NOV_ISPLATA))

    AssertNovacDoubleEquals 50#, originalLeft, _
                            "Partial otkup avans original row reduced to remainder"

    Dim totalApplied As Double
    totalApplied = GetIsplataForOtkup(otkupID)

    AssertNovacDoubleEquals 100#, totalApplied, _
                            "Partial otkup avans applied amount linked to otkup"

    Dim isplaceno As String
    isplaceno = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_ISPLACENO))

    AssertNovacTextEquals STATUS_ISPLACENO, isplaceno, _
                          "Partial otkup avans recomputes otkup as paid"

    Exit Sub

EH:
    LogNovacFail "Partial otkup avans split", _
                 "Err.Number=" & CStr(Err.Number) & _
                 " Source=" & Err.SOURCE & _
                 " Description=" & Err.description
End Sub

Private Sub Test_ResetNovacOtkupLinkRecomputesStatus()
    On Error GoTo EH

    Dim code As String
    code = NewNovacTestCode("RESETOTK")

    Dim kooperantID As String
    Dim otkupID As String
    Dim novacID As String

    kooperantID = "KOOP-TST-" & code
    otkupID = "OTK-TST-" & code

    AppendTestOtkupRow otkupID, kooperantID, 100#, 1#

    novacID = SaveNovac( _
        "TST-RESET-NOVAC-" & code, Date, _
        "TEST KOOPERANT " & code, kooperantID, "Kooperant", _
        "", kooperantID, "", "", _
        NOV_VIRMAN_FIRMA_KOOP, _
        0#, 100#, _
        "Test linked otkup payment", _
        otkupID)

    If Len(Trim$(novacID)) = 0 Then
        LogNovacFail "ResetNovacOtkupLink recomputes status", _
                     "Failed to create linked payment row."
        Exit Sub
    End If

    Call UpdateOtkupStatus(otkupID)

    Dim beforeStatus As String
    beforeStatus = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_ISPLACENO))

    AssertNovacTextEquals STATUS_ISPLACENO, beforeStatus, _
                          "Setup: otkup is paid before reset"

    If Not ResetNovacOtkupLink_TX(otkupID) Then
        LogNovacFail "ResetNovacOtkupLink recomputes status", _
                     "ResetNovacOtkupLink_TX returned False."
        Exit Sub
    End If

    Dim afterLinkedAmount As Double
    afterLinkedAmount = GetIsplataForOtkup(otkupID)

    AssertNovacDoubleEquals 0#, afterLinkedAmount, _
                            "ResetNovacOtkupLink removes linked payment amount"

    Dim afterStatus As String
    afterStatus = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_ISPLACENO))

    AssertNovacTextEquals "", afterStatus, _
                          "ResetNovacOtkupLink recomputes otkup as unpaid"

    Dim afterDate As String
    afterDate = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_DATUM_ISPLATE))

    AssertNovacTextEquals "", afterDate, _
                          "ResetNovacOtkupLink clears DatumIsplate"

    Exit Sub

EH:
    LogNovacFail "ResetNovacOtkupLink recomputes status", _
                 "Err.Number=" & CStr(Err.Number) & _
                 " Source=" & Err.SOURCE & _
                 " Description=" & Err.description
End Sub

Private Function NewNovacTestCode(ByVal prefixText As String) As String
    Randomize
    NewNovacTestCode = prefixText & "-" & _
                       Format$(Now, "yyyymmddhhnnss") & "-" & _
                       CStr(Int((9000 * Rnd) + 1000))
End Function

Private Sub AppendTestFakturaRow(ByVal fakturaID As String, _
                                 ByVal brojFakture As String, _
                                 ByVal kupacID As String, _
                                 ByVal iznos As Double)
    Const SRC As String = "AppendTestFakturaRow"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    values.Add COL_FAK_ID, fakturaID
    values.Add COL_FAK_BROJ, brojFakture
    values.Add COL_FAK_DATUM, Date
    values.Add COL_FAK_KUPAC, kupacID
    values.Add COL_FAK_IZNOS, iznos
    values.Add COL_FAK_STATUS, STATUS_NEPLACENO
    values.Add COL_STORNIRANO, ""

    AppendTestRowByColumnMap TBL_FAKTURE, values, SRC
End Sub

Private Sub AppendTestOtkupRow(ByVal otkupID As String, _
                               ByVal kooperantID As String, _
                               ByVal kolicina As Double, _
                               ByVal cena As Double)
    Const SRC As String = "AppendTestOtkupRow"

    Dim values As Object
    Set values = CreateObject("Scripting.Dictionary")

    values.Add COL_OTK_ID, otkupID
    values.Add COL_OTK_DATUM, Date
    values.Add COL_OTK_KOOPERANT, kooperantID
    values.Add COL_OTK_STANICA, "ST-TST"
    values.Add COL_OTK_KULTURA, "KUL-TST"
    values.Add COL_OTK_VRSTA, "Test Vrsta"
    values.Add COL_OTK_SORTA, "Test Sorta"
    values.Add COL_OTK_KOLICINA, kolicina
    values.Add COL_OTK_CENA, cena
    values.Add COL_OTK_TIP_AMB, "Test Amb"
    values.Add COL_OTK_KOL_AMB, 0
    values.Add COL_OTK_VOZAC, "VOZ-TST"
    values.Add COL_OTK_BR_DOK, "TST-OTK-" & otkupID
    values.Add COL_OTK_NOVAC, 0
    values.Add COL_OTK_PRIMALAC, "TEST"
    values.Add COL_OTK_KLASA, "I"
    values.Add COL_STORNIRANO, ""
    values.Add COL_OTK_ISPLACENO, ""
    values.Add COL_OTK_DATUM_ISPLATE, ""

    AppendTestRowByColumnMap TBL_OTKUP, values, SRC
End Sub

Private Sub AppendTestRowByColumnMap(ByVal tableName As String, _
                                     ByVal values As Object, _
                                     ByVal sourceName As String)
    Dim colCount As Long
    colCount = GetTestTableColumnCount(tableName)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 9101, sourceName, _
                  "Could not resolve table column count. Table=" & tableName
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    Dim key As Variant
    Dim colIndex As Long

    For Each key In values.keys
        colIndex = GetColumnIndex(tableName, CStr(key))

        If colIndex > 0 Then
            rowData(colIndex - 1) = values(key)
        End If
    Next key

    If AppendRow(tableName, rowData) <= 0 Then
        Err.Raise vbObjectError + 9102, sourceName, _
                  "Failed to append test row. Table=" & tableName
    End If
End Sub

Private Function GetTestTableColumnCount(ByVal tableName As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                GetTestTableColumnCount = lo.ListColumns.count
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Sub AssertNovacDoubleEquals(ByVal expectedValue As Double, _
                                    ByVal actualValue As Double, _
                                    ByVal testName As String)
    If Abs(expectedValue - actualValue) <= 0.0001 Then
        LogNovacPass testName
    Else
        LogNovacFail testName, _
                     "Expected=" & CStr(expectedValue) & _
                     " Actual=" & CStr(actualValue)
    End If
End Sub

Private Sub AssertNovacTextEquals(ByVal expectedValue As String, _
                                  ByVal actualValue As String, _
                                  ByVal testName As String)
    If Trim$(CStr(expectedValue)) = Trim$(CStr(actualValue)) Then
        LogNovacPass testName
    Else
        LogNovacFail testName, _
                     "Expected=[" & expectedValue & "] Actual=[" & actualValue & "]"
    End If
End Sub

Private Function CountNovacRowsForFaktura(ByVal fakturaID As String) As Long
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    Dim colFakID As Long
    colFakID = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)

    If colFakID = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakID))) = Trim$(fakturaID) Then
            CountNovacRowsForFaktura = CountNovacRowsForFaktura + 1
        End If
    Next i
End Function


