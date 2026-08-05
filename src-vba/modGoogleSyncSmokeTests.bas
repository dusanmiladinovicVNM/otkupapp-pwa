Attribute VB_Name = "modGoogleSyncSmokeTests"

Option Explicit

' ============================================================
' modGoogleSyncSmokeTests
'
' Smoke suite for VBA Google integration:
'   - tblSEFConfig Google keys
'   - OAuth token availability / refresh path
'   - Google Sheets create/write/read/find/add-tab
'   - Drive cleanup by trashing smoke spreadsheet
'
' Does NOT import real PWA OTK records.
' Does NOT write to tblOtkup.
'
' Entry:
'   RunGoogleSyncSmokeSuite
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_RunID As String

Private m_TestSpreadsheetID As String
Private m_TestSpreadsheetName As String

Private Const GOOGLE_DRIVE_API_BASE As String = "https://www.googleapis.com/drive/v3"

Public Sub RunGoogleSyncSmokeSuite()
    On Error GoTo EH

    BeginGoogleSmokeRun

    Test_GoogleConfigTableSchema
    Test_GoogleConfigRequiredKeys
    Test_GoogleAuthAccessToken
    Test_GoogleSheetsCreateWriteReadFind
    Test_GoogleSheetsAddTabWriteRead

    CleanupGoogleSmokeSpreadsheet

    EndGoogleSmokeRun
    Exit Sub

EH:
    LogGoogleSmokeFail "RunGoogleSyncSmokeSuite fatal", Err.description
    On Error Resume Next
    CleanupGoogleSmokeSpreadsheet
    EndGoogleSmokeRun
End Sub

' ============================================================
' TESTS
' ============================================================

Private Sub Test_GoogleConfigTableSchema()
    On Error GoTo EH

    AssertTrue Not GetTable(TBL_SEF_CONFIG) Is Nothing, "tblSEFConfig exists"

    Call RequireColumnIndex(TBL_SEF_CONFIG, "ConfigKey", "Test_GoogleConfigTableSchema")
    Call RequireColumnIndex(TBL_SEF_CONFIG, "ConfigValue", "Test_GoogleConfigTableSchema")

    LogGoogleSmokePass "tblSEFConfig has ConfigKey/ConfigValue"
    Exit Sub

EH:
    LogGoogleSmokeFail "tblSEFConfig schema", Err.description
End Sub

Private Sub Test_GoogleConfigRequiredKeys()
    On Error GoTo EH

    AssertConfigKeyPresent "GOOGLE_CLIENT_ID", True
    AssertConfigKeyPresent "GOOGLE_CLIENT_SECRET", True
    AssertConfigKeyPresent "GOOGLE_PWA_FOLDER_ID", True
    AssertConfigKeyPresent "GOOGLE_REFRESH_TOKEN", True
    AssertConfigKeyPresent "GOOGLE_ACCESS_TOKEN", False
    AssertConfigKeyPresent "GOOGLE_TOKEN_EXPIRES_AT", False

    LogGoogleSmokePass "Google config required keys checked"
    Exit Sub

EH:
    LogGoogleSmokeFail "Google config required keys", Err.description
End Sub

Private Sub Test_GoogleAuthAccessToken()
    On Error GoTo EH

    Dim token As String

    AssertTrue IsGoogleAuthConfigured(), "Google auth configured by refresh token"

    token = GetAccessToken()

    AssertTrue Len(Trim$(token)) > 0, "GetAccessToken returns token"

    ' Never print token.
    LogGoogleSmokePass "Google access token available"
    Exit Sub

EH:
    LogGoogleSmokeFail "Google auth access token", Err.description
End Sub

Private Sub Test_GoogleSheetsCreateWriteReadFind()
    On Error GoTo EH

    Dim folderID As String
    Dim headers(1 To 1, 1 To 4) As Variant
    Dim data(1 To 3, 1 To 4) As Variant
    Dim readBack As Variant
    Dim foundID As String

    folderID = Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))
    AssertTrue Len(folderID) > 0, "GOOGLE_PWA_FOLDER_ID present"

    m_TestSpreadsheetName = "TST-GOOGLE-SMOKE-" & m_RunID

    m_TestSpreadsheetID = CreateSpreadsheet(m_TestSpreadsheetName, folderID)
    AssertTrue Len(Trim$(m_TestSpreadsheetID)) > 0, "CreateSpreadsheet returns ID"

    headers(1, 1) = "ClientRecordID"
    headers(1, 2) = "SyncStatus"
    headers(1, 3) = "Value"
    headers(1, 4) = "CreatedAt"

    data(1, 1) = headers(1, 1)
    data(1, 2) = headers(1, 2)
    data(1, 3) = headers(1, 3)
    data(1, 4) = headers(1, 4)

    data(2, 1) = "TST-CRID-" & m_RunID & "-1"
    data(2, 2) = "Synced"
    data(2, 3) = "SmokeValue1"
    data(2, 4) = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    data(3, 1) = "TST-CRID-" & m_RunID & "-2"
    data(3, 2) = "Synced"
    data(3, 3) = "SmokeValue2"
    data(3, 4) = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    AssertTrue WriteSheetData(m_TestSpreadsheetID, "Sheet1", data), "WriteSheetData Sheet1"

    readBack = ReadSheetData(m_TestSpreadsheetID, "Sheet1")
    AssertTrue Not IsEmpty(readBack), "ReadSheetData returns data"
    AssertEquals "ClientRecordID", CStr(readBack(1, 1)), "Read header col A"
    AssertEquals "SyncStatus", CStr(readBack(1, 2)), "Read header col B"
    AssertEquals "SmokeValue1", CStr(readBack(2, 3)), "Read row value"

    foundID = GetSpreadsheetID(m_TestSpreadsheetName, folderID)
    AssertEquals m_TestSpreadsheetID, foundID, "GetSpreadsheetID exact-name returns created ID"

    LogGoogleSmokePass "Google Sheets create/write/read/find"
    Exit Sub

EH:
    LogGoogleSmokeFail "Google Sheets create/write/read/find", Err.description
End Sub

Private Sub Test_GoogleSheetsAddTabWriteRead()
    On Error GoTo EH

    Dim data(1 To 2, 1 To 3) As Variant
    Dim readBack As Variant

    AssertTrue Len(Trim$(m_TestSpreadsheetID)) > 0, "Smoke spreadsheet exists before AddTab"

    AssertTrue AddSheetTab(m_TestSpreadsheetID, "SmokeTab"), "AddSheetTab SmokeTab"

    data(1, 1) = "Key"
    data(1, 2) = "Value"
    data(1, 3) = "RunID"

    data(2, 1) = "Smoke"
    data(2, 2) = "PASS"
    data(2, 3) = m_RunID

    AssertTrue WriteSheetData(m_TestSpreadsheetID, "SmokeTab", data), "WriteSheetData SmokeTab"

    readBack = ReadSheetData(m_TestSpreadsheetID, "SmokeTab")
    AssertTrue Not IsEmpty(readBack), "Read SmokeTab returns data"
    AssertEquals "Key", CStr(readBack(1, 1)), "SmokeTab header"
    AssertEquals "PASS", CStr(readBack(2, 2)), "SmokeTab value"

    LogGoogleSmokePass "Google Sheets add-tab/write/read"
    Exit Sub

EH:
    LogGoogleSmokeFail "Google Sheets add-tab/write/read", Err.description
End Sub

' ============================================================
' CLEANUP
' ============================================================

Private Sub CleanupGoogleSmokeSpreadsheet()
    On Error GoTo EH

    If Len(Trim$(m_TestSpreadsheetID)) = 0 Then Exit Sub

    If TrashGoogleDriveFile(m_TestSpreadsheetID) Then
        LogGoogleSmokePass "Cleanup trashed smoke spreadsheet"
    Else
        LogGoogleSmokeFail "Cleanup smoke spreadsheet", _
                           "Could not trash spreadsheet ID=" & m_TestSpreadsheetID
    End If

    Exit Sub

EH:
    LogGoogleSmokeFail "Cleanup smoke spreadsheet", Err.description
End Sub

Private Function TrashGoogleDriveFile(ByVal fileID As String) As Boolean
    Dim accessToken As String
    Dim http As Object
    Dim url As String
    Dim body As String

    On Error GoTo EH

    If Len(Trim$(fileID)) = 0 Then
        TrashGoogleDriveFile = False
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        TrashGoogleDriveFile = False
        Exit Function
    End If

    url = GOOGLE_DRIVE_API_BASE & "/files/" & fileID
    body = "{""trashed"":true}"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000

    http.Open "PATCH", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body

    TrashGoogleDriveFile = (http.status >= 200 And http.status < 300)
    Exit Function

EH:
    TrashGoogleDriveFile = False
End Function

' ============================================================
' ASSERT / RUN LOGGING
' ============================================================

Private Sub BeginGoogleSmokeRun()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_RunID = Format$(Now, "yyyymmddhhnnss")

    Debug.Print String$(60, "-")
    Debug.Print "GOOGLE SYNC SMOKE SUITE START " & m_RunID
    Debug.Print String$(60, "-")
End Sub

Private Sub EndGoogleSmokeRun()
    Debug.Print String$(60, "-")
    Debug.Print "GOOGLE SYNC SMOKE SUITE END"
    Debug.Print "TOTAL=" & m_Total & " PASS=" & m_Passed & " FAIL=" & m_Failed
    Debug.Print String$(60, "-")

    If m_Failed = 0 Then
        MsgBox "Google Sync Smoke Suite PASS" & vbCrLf & _
               "Total: " & m_Total & vbCrLf & _
               "Pass: " & m_Passed, vbInformation, APP_NAME
    Else
        MsgBox "Google Sync Smoke Suite FAIL" & vbCrLf & _
               "Total: " & m_Total & vbCrLf & _
               "Pass: " & m_Passed & vbCrLf & _
               "Fail: " & m_Failed & vbCrLf & _
               "Pogledaj Immediate Window / log.", vbCritical, APP_NAME
    End If
End Sub

Private Sub LogGoogleSmokePass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1

    Debug.Print "PASS | " & testName
    LogInfo "GoogleSmoke", "PASS | " & testName
End Sub

Private Sub LogGoogleSmokeFail(ByVal testName As String, ByVal message As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "FAIL | " & testName & " | " & message
    LogError "GoogleSmoke", "FAIL | " & testName & " | " & message
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If condition Then
        LogGoogleSmokePass message
    Else
        Err.Raise vbObjectError + 7800, "AssertTrue", message
    End If
End Sub

Private Sub AssertEquals(ByVal expected As String, _
                         ByVal actual As String, _
                         ByVal message As String)
    If CStr(expected) = CStr(actual) Then
        LogGoogleSmokePass message
    Else
        Err.Raise vbObjectError + 7801, "AssertEquals", _
                  message & " | Expected='" & expected & "' Actual='" & actual & "'"
    End If
End Sub

Private Sub AssertConfigKeyPresent(ByVal configKey As String, _
                                   ByVal valueRequired As Boolean)
    Dim value As String

    value = GetConfigValue(configKey)

    If valueRequired Then
        AssertTrue Len(Trim$(value)) > 0, "Config key has value: " & configKey
    Else
        If Len(Trim$(value)) > 0 Then
            LogGoogleSmokePass "Config key exists/has value: " & configKey
        Else
            LogGoogleSmokePass "Config key optional/empty allowed: " & configKey
        End If
    End If
End Sub

Private Function NormalizeGoogleCellText(ByVal value As Variant) As String
    Dim s As String

    s = CStr(value)
    s = Replace(s, "\u003e", ">")
    s = Replace(s, "\u003c", "<")
    s = Replace(s, "\u0026", "&")
    s = Replace(s, "\u0027", "'")
    s = Replace(s, "\u0022", """")

    NormalizeGoogleCellText = s
End Function

' ============================================================
' MASTER SYNC FIXTURE SUITE
'
' Tests real OTK Google Sheet -> VBA import path using a temporary
' Google Sheet fixture and Excel transaction rollback.
'
' Entry:
'   RunMasterSyncSmokeSuite
' ============================================================

Public Sub RunMasterSyncSmokeSuite()
    Dim tx As clsTransaction

    On Error GoTo EH

    BeginMasterSyncSmokeRun

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA

    Test_MasterSyncFixtureImportAndWriteBack
    Test_MasterSyncDuplicateClientRecordID
    Test_MasterSyncMissingClientRecordID

    tx.RollbackTx
    CleanupGoogleSmokeSpreadsheet

    EndGoogleSmokeRun
    Exit Sub

EH:
    LogGoogleSmokeFail "RunMasterSyncSmokeSuite fatal", Err.description

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    CleanupGoogleSmokeSpreadsheet
    EndGoogleSmokeRun
End Sub

Private Sub BeginMasterSyncSmokeRun()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_RunID = Format$(Now, "yyyymmddhhnnss")

    m_TestSpreadsheetID = ""
    m_TestSpreadsheetName = ""

    Debug.Print String$(60, "-")
    Debug.Print "MASTER SYNC SMOKE SUITE START " & m_RunID
    Debug.Print String$(60, "-")
End Sub

' ============================================================
' MASTER SYNC TESTS
' ============================================================

Private Sub Test_MasterSyncFixtureImportAndWriteBack()
    On Error GoTo EH

    Dim folderID As String
    Dim fixtureData As Variant
    Dim readBack As Variant
    Dim imported As Long
    Dim skipped As Long
    Dim errors As Long
    Dim clientRecordID As String
    Dim newOtkupID As String

    folderID = Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))
    AssertTrue Len(folderID) > 0, "MASTER SYNC: GOOGLE_PWA_FOLDER_ID present"

    clientRecordID = "TST-MSYNC-" & m_RunID & "-IMPORT"

    m_TestSpreadsheetName = "OTK-TST-SYNC-" & m_RunID
    m_TestSpreadsheetID = CreateSpreadsheet(m_TestSpreadsheetName, folderID)
    AssertTrue Len(Trim$(m_TestSpreadsheetID)) > 0, "MASTER SYNC: fixture spreadsheet created"

    fixtureData = BuildOTKFixtureData(clientRecordID, "Synced")
    AssertTrue WriteSheetData(m_TestSpreadsheetID, "Sheet1", fixtureData), _
               "MASTER SYNC: fixture OTK sheet written"

    Call TestHook_ImportOneOTKSheet(m_TestSpreadsheetID, m_TestSpreadsheetName, imported, skipped, errors)

    AssertEquals "1", CStr(imported), "MASTER SYNC: one OTK row imported"
    AssertEquals "0", CStr(errors), "MASTER SYNC: no import errors"

    newOtkupID = FindOtkupIDByClientRecordID(clientRecordID)
    AssertTrue Len(newOtkupID) > 0, "MASTER SYNC: tblOtkup contains imported ClientRecordID"

    readBack = ReadSheetData(m_TestSpreadsheetID, "Sheet1")
    AssertTrue Not IsEmpty(readBack), "MASTER SYNC: read back fixture sheet after import"

    AssertEquals "Synced>Master", NormalizeGoogleCellText(readBack(2, 6)), _
                "MASTER SYNC: Google row SyncStatus written back"

    AssertTrue Len(Trim$(CStr(readBack(2, 2)))) > 0, _
               "MASTER SYNC: Google row ServerRecordID written back"

    LogGoogleSmokePass "MASTER SYNC: fixture import/write-back PASS"
    Exit Sub

EH:
    LogGoogleSmokeFail "MASTER SYNC: fixture import/write-back", Err.description
End Sub

Private Sub Test_MasterSyncDuplicateClientRecordID()
    On Error GoTo EH

    Dim fixtureData As Variant
    Dim readBack As Variant
    Dim imported As Long
    Dim skipped As Long
    Dim errors As Long
    Dim clientRecordID As String

    AssertTrue Len(Trim$(m_TestSpreadsheetID)) > 0, _
               "MASTER SYNC DUP: fixture spreadsheet exists"

    clientRecordID = "TST-MSYNC-" & m_RunID & "-IMPORT"

    ' Reset same ClientRecordID back to Synced to simulate PWA retry / duplicate.
    fixtureData = BuildOTKFixtureData(clientRecordID, "Synced")
    AssertTrue WriteSheetData(m_TestSpreadsheetID, "Sheet1", fixtureData), _
               "MASTER SYNC DUP: duplicate fixture written"

    Call TestHook_ImportOneOTKSheet(m_TestSpreadsheetID, m_TestSpreadsheetName, imported, skipped, errors)

    AssertEquals "0", CStr(imported), "MASTER SYNC DUP: duplicate not imported"
    AssertTrue skipped >= 1, "MASTER SYNC DUP: duplicate skipped"
    AssertEquals "0", CStr(errors), "MASTER SYNC DUP: duplicate is not import error"

    readBack = ReadSheetData(m_TestSpreadsheetID, "Sheet1")
    AssertTrue Not IsEmpty(readBack), "MASTER SYNC DUP: read back fixture sheet"

    AssertEquals "Duplicate", NormalizeGoogleCellText(readBack(2, 6)), _
             "MASTER SYNC DUP: Google row marked Duplicate"

    LogGoogleSmokePass "MASTER SYNC: duplicate ClientRecordID PASS"
    Exit Sub

EH:
    LogGoogleSmokeFail "MASTER SYNC: duplicate ClientRecordID", Err.description
End Sub

Private Sub Test_MasterSyncMissingClientRecordID()
    On Error GoTo EH

    Dim fixtureData As Variant
    Dim readBack As Variant
    Dim imported As Long
    Dim skipped As Long
    Dim errors As Long

    AssertTrue Len(Trim$(m_TestSpreadsheetID)) > 0, _
               "MASTER SYNC MISSING CRID: fixture spreadsheet exists"

    fixtureData = BuildOTKFixtureData("", "Synced")
    AssertTrue WriteSheetData(m_TestSpreadsheetID, "Sheet1", fixtureData), _
               "MASTER SYNC MISSING CRID: fixture written"

    Call TestHook_ImportOneOTKSheet(m_TestSpreadsheetID, m_TestSpreadsheetName, imported, skipped, errors)

    AssertEquals "0", CStr(imported), "MASTER SYNC MISSING CRID: no import"
    AssertTrue errors >= 1, "MASTER SYNC MISSING CRID: error counted"

    readBack = ReadSheetData(m_TestSpreadsheetID, "Sheet1")
    AssertTrue Not IsEmpty(readBack), "MASTER SYNC MISSING CRID: read back fixture sheet"

    AssertTrue Left$(NormalizeGoogleCellText(readBack(2, 6)), Len("SyncError")) = "SyncError", _
           "MASTER SYNC MISSING CRID: Google row marked SyncError"

    LogGoogleSmokePass "MASTER SYNC: missing ClientRecordID PASS"
    Exit Sub

EH:
    LogGoogleSmokeFail "MASTER SYNC: missing ClientRecordID", Err.description
End Sub

' ============================================================
' MASTER SYNC TEST HELPERS
' ============================================================

Private Function BuildOTKFixtureData(ByVal clientRecordID As String, _
                                     ByVal syncStatus As String) As Variant
    Dim data(1 To 2, 1 To 22) As Variant
    Dim kooperantID As String
    Dim kooperantName As String
    Dim vrstaVoca As String
    Dim sortaVoca As String
    Dim otkupacID As String

    kooperantID = GetFirstRequiredTableValue(TBL_KOOPERANTI, "KooperantID")
    kooperantName = GetFirstOptionalTableValue(TBL_KOOPERANTI, "Ime", "Test Kooperant")
    vrstaVoca = GetFirstOptionalTableValue(TBL_KULTURE, "VrstaVoca", "Visnja")
    sortaVoca = GetConfigValue("DefaultSorta")
    If Len(Trim$(sortaVoca)) = 0 Then sortaVoca = "Default"

    otkupacID = GetFirstOptionalTableValue(TBL_STANICE, "StanicaID", "ST-00001")

    ' Header row -- must match GAS/VBA OTK schema exactly.
    data(1, 1) = "ClientRecordID"
    data(1, 2) = "ServerRecordID"
    data(1, 3) = "CreatedAtClient"
    data(1, 4) = "UpdatedAtClient"
    data(1, 5) = "UpdatedAtServer"
    data(1, 6) = "SyncStatus"
    data(1, 7) = "DeviceID"
    data(1, 8) = "OtkupacID"
    data(1, 9) = "Datum"
    data(1, 10) = "KooperantID"
    data(1, 11) = "KooperantName"
    data(1, 12) = "VrstaVoca"
    data(1, 13) = "SortaVoca"
    data(1, 14) = "Klasa"
    data(1, 15) = "Koli" & ChrW(269) & "ina"
    data(1, 16) = "Cena"
    data(1, 17) = "TipAmbalaze"
    data(1, 18) = "KolAmbalaze"
    data(1, 19) = "ParcelaID"
    data(1, 20) = "VozacID"
    data(1, 21) = "Napomena"
    data(1, 22) = "ReceivedAt"

    ' Data row.
    data(2, 1) = clientRecordID
    data(2, 2) = ""
    data(2, 3) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    data(2, 4) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    data(2, 5) = ""
    data(2, 6) = syncStatus
    data(2, 7) = "VBA-SMOKE"
    data(2, 8) = otkupacID
    data(2, 9) = Format$(Date, "yyyy-mm-dd")
    data(2, 10) = kooperantID
    data(2, 11) = kooperantName
    data(2, 12) = vrstaVoca
    data(2, 13) = sortaVoca
    data(2, 14) = "I"
    data(2, 15) = 12.5
    data(2, 16) = 100
    data(2, 17) = "12/1"
    data(2, 18) = 1
    data(2, 19) = ""
    data(2, 20) = ""
    data(2, 21) = "MasterSyncSmoke " & m_RunID
    data(2, 22) = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    BuildOTKFixtureData = data
End Function

Private Function FindOtkupIDByClientRecordID(ByVal clientRecordID As String) As String
    Dim data As Variant
    Dim colCRID As Long
    Dim colOtkupID As Long
    Dim i As Long

    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    colCRID = RequireColumnIndex(TBL_OTKUP, "ClientRecordID", "FindOtkupIDByClientRecordID")
    colOtkupID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "FindOtkupIDByClientRecordID")

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(nz(data(i, colCRID), ""))) = Trim$(clientRecordID) Then
            FindOtkupIDByClientRecordID = CStr(nz(data(i, colOtkupID), ""))
            Exit Function
        End If
    Next i
End Function

Private Function GetFirstRequiredTableValue(ByVal tableName As String, _
                                            ByVal columnName As String) As String
    Dim data As Variant
    Dim colIndex As Long

    data = GetTableData(tableName)
    If IsEmpty(data) Then
        Err.Raise vbObjectError + 7900, "GetFirstRequiredTableValue", _
                  "Tabela je prazna: " & tableName
    End If

    colIndex = RequireColumnIndex(tableName, columnName, "GetFirstRequiredTableValue")

    GetFirstRequiredTableValue = Trim$(CStr(nz(data(1, colIndex), "")))

    If Len(GetFirstRequiredTableValue) = 0 Then
        Err.Raise vbObjectError + 7901, "GetFirstRequiredTableValue", _
                  "Prva vrednost je prazna: " & tableName & "." & columnName
    End If
End Function

Private Function GetFirstOptionalTableValue(ByVal tableName As String, _
                                            ByVal columnName As String, _
                                            ByVal fallbackValue As String) As String
    Dim data As Variant
    Dim colIndex As Long

    On Error GoTo UseFallback

    data = GetTableData(tableName)
    If IsEmpty(data) Then GoTo UseFallback

    colIndex = GetColumnIndex(tableName, columnName)
    If colIndex = 0 Then GoTo UseFallback

    GetFirstOptionalTableValue = Trim$(CStr(nz(data(1, colIndex), "")))
    If Len(GetFirstOptionalTableValue) = 0 Then GoTo UseFallback

    Exit Function

UseFallback:
    GetFirstOptionalTableValue = fallbackValue
End Function

' ============================================================
' SHEETS JSON PARSER SUITE (AUD-001)
'
' Offline unit tests za modGoogleSheets.ParseValuesJson.
' Ne zovu Google, ne pisu u tabele -- mogu se pustiti bilo kada.
'
' Entry:
'   RunSheetsJsonParserTests
' ============================================================

Public Sub RunSheetsJsonParserTests()
    On Error GoTo EH

    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_RunID = Format$(Now, "yyyymmddhhnnss")

    Debug.Print String$(60, "-")
    Debug.Print "SHEETS JSON PARSER SUITE START " & m_RunID
    Debug.Print String$(60, "-")

    Test_ParseValuesJson_SpaceAfterCommaPreserved
    Test_ParseValuesJson_EscapedQuote
    Test_ParseValuesJson_UnicodeEscape
    Test_ParseValuesJson_EmbeddedRowSeparator
    Test_ParseValuesJson_WhitespaceAndEscapes
    Test_ParseValuesJson_NoValues

    EndSheetsJsonParserRun
    Exit Sub

EH:
    LogGoogleSmokeFail "RunSheetsJsonParserTests fatal", Err.description
    EndSheetsJsonParserRun
End Sub

Private Sub EndSheetsJsonParserRun()
    Debug.Print String$(60, "-")
    Debug.Print "SHEETS JSON PARSER SUITE END"
    Debug.Print "TOTAL=" & m_Total & " PASS=" & m_Passed & " FAIL=" & m_Failed
    Debug.Print String$(60, "-")

    If m_Failed = 0 Then
        MsgBox "Sheets JSON Parser Suite PASS" & vbCrLf & _
               "Total: " & m_Total & vbCrLf & _
               "Pass: " & m_Passed, vbInformation, APP_NAME
    Else
        MsgBox "Sheets JSON Parser Suite FAIL" & vbCrLf & _
               "Total: " & m_Total & vbCrLf & _
               "Pass: " & m_Passed & vbCrLf & _
               "Fail: " & m_Failed & vbCrLf & _
               "Pogledaj Immediate Window / log.", vbCritical, APP_NAME
    End If
End Sub

Private Sub Test_ParseValuesJson_SpaceAfterCommaPreserved()
    ' AUD-001 (a): globalni Replace(", " -> ",") je brisao razmak u SVAKOJ
    ' tekst-celiji (adrese, imena, napomene).
    On Error GoTo EH

    Dim json As String
    Dim data As Variant

    json = "{""range"":""Sheet1!A1:B2"",""majorDimension"":""ROWS""," & _
           """values"":[[""Adresa"",""Napomena""]," & _
           "[""Ulica 5, Beograd"",""prvo, drugo, trece""]]}"

    data = ParseValuesJson(json)

    AssertTrue Not IsEmpty(data), "PARSER: space-after-comma sheet parsed"
    AssertEquals "2", CStr(UBound(data, 1)), "PARSER: 2 rows"
    AssertEquals "2", CStr(UBound(data, 2)), "PARSER: 2 cols"
    AssertEquals "Ulica 5, Beograd", CStr(data(2, 1)), _
                 "PARSER: razmak posle zapete ocuvan"
    AssertEquals "prvo, drugo, trece", CStr(data(2, 2)), _
                 "PARSER: vise zapeta u celiji ocuvano"
    Exit Sub

EH:
    LogGoogleSmokeFail "PARSER: space after comma", Err.description
End Sub

Private Sub Test_ParseValuesJson_EscapedQuote()
    ' AUD-001 (b): \" je lomio quote tracking -> mis-split reda.
    On Error GoTo EH

    Dim json As String
    Dim data As Variant

    json = "{""values"":[[""On je rekao \""zdravo\"" tiho"",""druga"",""treca""]]}"

    data = ParseValuesJson(json)

    AssertTrue Not IsEmpty(data), "PARSER: escaped-quote sheet parsed"
    AssertEquals "3", CStr(UBound(data, 2)), "PARSER: escaped quote ne lomi split (3 celije)"
    AssertEquals "On je rekao ""zdravo"" tiho", CStr(data(1, 1)), _
                 "PARSER: navodnik dekodovan i ocuvan"
    AssertEquals "druga", CStr(data(1, 2)), "PARSER: druga celija netaknuta"
    Exit Sub

EH:
    LogGoogleSmokeFail "PARSER: escaped quote", Err.description
End Sub

Private Sub Test_ParseValuesJson_UnicodeEscape()
    ' AUD-001 (c): \uXXXX je prolazio doslovno (s/c/c/z/dj, & < > ' ").
    On Error GoTo EH

    Dim json As String
    Dim data As Variant
    Dim expectedText As String

    ' \u0161 = s-caron, \u010d = c-caron, \u0111 = d-stroke
    json = "{""values"":[[""\u0161e\u010de\u0111a"",""A\u0026B"",""Synced\u003eMaster""]]}"

    data = ParseValuesJson(json)

    ' Ocekivana vrednost se gradi u runtime-u (ChrW) -- izvor ostaje ASCII.
    expectedText = ChrW$(353) & "e" & ChrW$(269) & "e" & ChrW$(273) & "a"

    AssertTrue Not IsEmpty(data), "PARSER: unicode-escape sheet parsed"
    AssertEquals expectedText, CStr(data(1, 1)), _
                 "PARSER: \uXXXX dekodovan u srpska slova"
    AssertEquals "A&B", CStr(data(1, 2)), "PARSER: \u0026 dekodovan u &"
    AssertEquals "Synced>Master", CStr(data(1, 3)), "PARSER: \u003e dekodovan u >"
    Exit Sub

EH:
    LogGoogleSmokeFail "PARSER: unicode escape", Err.description
End Sub

Private Sub Test_ParseValuesJson_EmbeddedRowSeparator()
    ' AUD-001 (d): Split(block, "],[") je lomio red na tekstu koji sadrzi "],[".
    On Error GoTo EH

    Dim json As String
    Dim data As Variant

    json = "{""values"":[[""niz ],[ u tekstu"",""druga""],[""r2c1"",""r2c2""]]}"

    data = ParseValuesJson(json)

    AssertTrue Not IsEmpty(data), "PARSER: embedded ],[ sheet parsed"
    AssertEquals "2", CStr(UBound(data, 1)), "PARSER: embedded ],[ ne pravi lazni red"
    AssertEquals "niz ],[ u tekstu", CStr(data(1, 1)), _
                 "PARSER: embedded ],[ ocuvan u celiji"
    AssertEquals "r2c2", CStr(data(2, 2)), "PARSER: drugi red netaknut"
    Exit Sub

EH:
    LogGoogleSmokeFail "PARSER: embedded row separator", Err.description
End Sub

Private Sub Test_ParseValuesJson_WhitespaceAndEscapes()
    ' prettyPrint JSON (razmaci/novi redovi izmedju tokena) + \\ \/ \t \n
    ' + necitirane brojcane vrednosti.
    On Error GoTo EH

    Dim json As String
    Dim data As Variant

    json = "{" & vbCrLf & _
           "  ""majorDimension"": ""ROWS""," & vbCrLf & _
           "  ""values"": [" & vbCrLf & _
           "    [ ""C:\\putanja\\fajl"", ""a\tb"", ""red1\nred2"", ""1/2"" ]," & vbCrLf & _
           "    [ 12345, 2.5, """", ""x"" ]" & vbCrLf & _
           "  ]" & vbCrLf & _
           "}"

    data = ParseValuesJson(json)

    AssertTrue Not IsEmpty(data), "PARSER: prettyPrint JSON parsed"
    AssertEquals "2", CStr(UBound(data, 1)), "PARSER: prettyPrint 2 rows"
    AssertEquals "4", CStr(UBound(data, 2)), "PARSER: prettyPrint 4 cols"
    AssertEquals "C:\putanja\fajl", CStr(data(1, 1)), "PARSER: \\ dekodovan"
    AssertEquals "a" & vbTab & "b", CStr(data(1, 2)), "PARSER: \t dekodovan"
    AssertEquals "red1" & vbLf & "red2", CStr(data(1, 3)), "PARSER: \n dekodovan"
    AssertEquals "12345", CStr(data(2, 1)), "PARSER: broj bez navodnika"
    AssertEquals "2.5", CStr(data(2, 2)), "PARSER: decimalni broj bez navodnika"
    AssertEquals "", CStr(data(2, 3)), "PARSER: prazna celija ostaje prazna"
    Exit Sub

EH:
    LogGoogleSmokeFail "PARSER: whitespace and escapes", Err.description
End Sub

Private Sub Test_ParseValuesJson_NoValues()
    ' Prazan/nevalidan odgovor -> Empty (nepromenjeno ponasanje).
    On Error GoTo EH

    AssertTrue IsEmpty(ParseValuesJson("{""range"":""Sheet1!A1:Z1000""}")), _
               "PARSER: bez values -> Empty"
    AssertTrue IsEmpty(ParseValuesJson("{""values"":[]}")), _
               "PARSER: prazan values niz -> Empty"
    Exit Sub

EH:
    LogGoogleSmokeFail "PARSER: no values", Err.description
End Sub

