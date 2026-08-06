Attribute VB_Name = "modGoogleSheets"
Option Explicit

' ============================================================
' modGoogleSheets - Google Sheets API v4 Wrapper
'
' Liest und schreibt Google Sheets via REST API.
' Auth via modGoogleAuth.GetAccessToken()
'
' Hauptfunktionen:
'   WriteSheetData   -- schreibt 2D-Array in ein Sheet-Tab
'   ReadSheetData    -- liest Sheet-Tab als 2D-Array
'   ClearSheet       -- loescht alle Daten in einem Tab
'   CreateSpreadsheet -- erstellt neues Google Sheet
'   GetSpreadsheetID -- sucht Sheet-ID nach Name in einem Folder
' ============================================================

' ============================================================
' PATCH: P1 -- modGoogleSheets.WriteSheetData staging/verify/replace
' File: src-vba/modGoogleSheets.bas
'
' Problem:
'   Current WriteSheetData does:
'     1) ClearSheet(target)
'     2) PUT values(target)
'
'   If ClearSheet succeeds and PUT values fails, the target tab can remain empty.
'
' Required minimum:
'   Every WriteSheetData=False must make the full sync fail.
'   Existing callers already propagate False through SyncStammdatenToGoogle_Core,
'   ExportKarticeToGoogle_Core, ExportMgmtReports_Core and orchestrator booleans.
'
' Better implementation in this patch:
'   Write to staging tab -> verify staging -> replace target tab by batchUpdate.
'
' Notes:
'   - Does not delete any project file or unrelated code.
'   - Replaces only WriteSheetData behavior and adds private helpers.
'   - Target tab sheetId changes after replace. This is acceptable for exported
'     data tabs that are consumed by tab name. If a future tab must preserve
'     sheetId due to charts/protected ranges, use a separate preserve-sheetId path.
' ============================================================

Private Const SHEETS_API_BASE As String = "https://sheets.googleapis.com/v4/spreadsheets"
Private Const DRIVE_API_BASE As String = "https://www.googleapis.com/drive/v3"

Private Const GOOGLE_HTTP_MAX_RETRIES As Long = 3
Private Const GOOGLE_HTTP_RETRY_BASE_MS As Long = 1500
Private Const GOOGLE_HTTP_RETRY_MAX_MS As Long = 8000

Private Const GOOGLE_WRITE_MIN_INTERVAL_MS As Long = 1250
Private Const GOOGLE_HTTP_429_WAIT_MS As Long = 65000

Private mLastGoogleWriteTimer As Double

Private mSheetIdCacheBySpreadsheet As Object

Private Function CreateGoogleHttpRequest(ByVal sourceName As String) As Object
    Dim http As Object

    On Error GoTo EH

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000

    Set CreateGoogleHttpRequest = http
    Exit Function

EH:
    LogErr sourceName & ".CreateGoogleHttpRequest"
    Err.Raise Err.Number, sourceName, Err.description
End Function

Private Function SendGoogleHttpWithRetry(ByVal http As Object, _
                                         ByVal sourceName As String, _
                                         Optional ByVal body As String = vbNullString, _
                                         Optional ByVal hasBody As Boolean = False, _
                                         Optional ByVal isWriteRequest As Boolean = False) As Boolean
    Dim attempt As Long
    Dim waitMs As Long

    On Error GoTo EH

    For attempt = 0 To GOOGLE_HTTP_MAX_RETRIES
        If isWriteRequest Then ThrottleGoogleWriteRequest sourceName

        If hasBody Then
            http.Send body
        Else
            http.Send
        End If

        If Not IsRetryableGoogleHttpStatus(CLng(http.status)) Then
            SendGoogleHttpWithRetry = True
            Exit Function
        End If

        If attempt >= GOOGLE_HTTP_MAX_RETRIES Then
            SendGoogleHttpWithRetry = True
            Exit Function
        End If

        waitMs = GoogleRetryAfterMs(http, CLng(http.status), attempt)

        LogWarn sourceName, _
                "Google HTTP " & CStr(http.status) & _
                " retry " & CStr(attempt + 1) & "/" & CStr(GOOGLE_HTTP_MAX_RETRIES) & _
                " after " & CStr(waitMs) & "ms"

        SleepMilliseconds waitMs
    Next attempt

    SendGoogleHttpWithRetry = True
    Exit Function

EH:
    LogErr sourceName & ".SendGoogleHttpWithRetry"
    SendGoogleHttpWithRetry = False
End Function

Private Function IsRetryableGoogleHttpStatus(ByVal statusCode As Long) As Boolean
    Select Case statusCode
        Case 429, 500, 502, 503, 504
            IsRetryableGoogleHttpStatus = True
        Case Else
            IsRetryableGoogleHttpStatus = False
    End Select
End Function

Private Function GoogleRetryDelayMs(ByVal statusCode As Long, ByVal attempt As Long) As Long
    Dim delayMs As Long

    delayMs = GOOGLE_HTTP_RETRY_BASE_MS * (2 ^ attempt)

    If delayMs > GOOGLE_HTTP_RETRY_MAX_MS Then delayMs = GOOGLE_HTTP_RETRY_MAX_MS

    ' Small jitter to avoid repeated synchronized retries.
    Randomize
    delayMs = delayMs + CLng(Int(Rnd() * 500))

    GoogleRetryDelayMs = delayMs
End Function

Private Sub SleepMilliseconds(ByVal milliseconds As Long)
    Dim untilTime As Date

    If milliseconds <= 0 Then Exit Sub

    untilTime = DateAdd("s", CDbl(milliseconds) / 1000#, Now)

    Do While Now < untilTime
        DoEvents
    Loop
End Sub

Private Sub ThrottleGoogleWriteRequest(ByVal sourceName As String)
    Dim nowTimer As Double
    Dim elapsedMs As Long
    Dim waitMs As Long

    nowTimer = Timer

    If mLastGoogleWriteTimer > 0 Then
        If nowTimer >= mLastGoogleWriteTimer Then
            elapsedMs = CLng((nowTimer - mLastGoogleWriteTimer) * 1000#)
        Else
            ' Timer wrapped at midnight.
            elapsedMs = CLng(((86400# - mLastGoogleWriteTimer) + nowTimer) * 1000#)
        End If

        If elapsedMs < GOOGLE_WRITE_MIN_INTERVAL_MS Then
            waitMs = GOOGLE_WRITE_MIN_INTERVAL_MS - elapsedMs
            SleepMilliseconds waitMs
        End If
    End If

    mLastGoogleWriteTimer = Timer
End Sub

Private Function GoogleRetryAfterMs(ByVal http As Object, _
                                    ByVal statusCode As Long, _
                                    ByVal attempt As Long) As Long
    Dim retryAfter As String
    Dim delayMs As Long

    On Error Resume Next
    retryAfter = CStr(http.GetResponseHeader("Retry-After"))
    On Error GoTo 0

    If Len(Trim$(retryAfter)) > 0 Then
        If IsNumeric(retryAfter) Then
            GoogleRetryAfterMs = CLng(CDbl(retryAfter) * 1000#)
            If GoogleRetryAfterMs < 1000 Then GoogleRetryAfterMs = 1000
            Exit Function
        End If
    End If

    If statusCode = 429 Then
        GoogleRetryAfterMs = GOOGLE_HTTP_429_WAIT_MS
        Exit Function
    End If

    delayMs = GOOGLE_HTTP_RETRY_BASE_MS * (2 ^ attempt)
    If delayMs > GOOGLE_HTTP_RETRY_MAX_MS Then delayMs = GOOGLE_HTTP_RETRY_MAX_MS

    Randomize
    GoogleRetryAfterMs = delayMs + CLng(Int(Rnd() * 500))
End Function

Private Function RequireGoogleTextArg(ByVal value As String, _
                                      ByVal argName As String, _
                                      ByVal sourceName As String) As Boolean
    If Len(Trim$(value)) = 0 Then
        LogError sourceName, argName & " je prazan."
        RequireGoogleTextArg = False
    Else
        RequireGoogleTextArg = True
    End If
End Function

Private Function GoogleHttpBodyForLog(ByVal responseText As String) As String
    GoogleHttpBodyForLog = Left$(CStr(responseText), 1000)
End Function

Private Function IsGoogleSheetAlreadyExistsError(ByVal httpStatus As Long, _
                                                 ByVal responseText As String) As Boolean
    Dim s As String

    If httpStatus <> 400 Then Exit Function

    s = LCase$(CStr(responseText))

    IsGoogleSheetAlreadyExistsError = _
        (InStr(1, s, "a sheet with the name", vbTextCompare) > 0 And _
         InStr(1, s, "already exists", vbTextCompare) > 0)
End Function

Private Function IsTwoDimArray(ByVal value As Variant) As Boolean
    On Error GoTo EH

    Dim lb1 As Long
    Dim ub1 As Long
    Dim lb2 As Long
    Dim ub2 As Long

    lb1 = LBound(value, 1)
    ub1 = UBound(value, 1)
    lb2 = LBound(value, 2)
    ub2 = UBound(value, 2)

    IsTwoDimArray = (ub1 >= lb1 And ub2 >= lb2)
    Exit Function

EH:
    IsTwoDimArray = False
End Function

' ============================================================
' PUBLIC -- Write
' ============================================================

Public Function WriteSheetData(ByVal spreadsheetID As String, _
                               ByVal tabName As String, _
                               ByVal data As Variant) As Boolean
    Const SRC As String = "WriteSheetData"

    Dim stagingTab As String
    Dim stagingSheetId As Long
    Dim targetSheetId As Long

    On Error GoTo EH

    WriteSheetData = False

    If Len(Trim$(spreadsheetID)) = 0 Then
        LogError SRC, "spreadsheetID je prazan."
        Exit Function
    End If

    If Len(Trim$(tabName)) = 0 Then
        LogError SRC, "tabName je prazan."
        Exit Function
    End If

    If IsEmpty(data) Then
        LogError SRC, "data je Empty."
        Exit Function
    End If

    If Not IsTwoDimArray(data) Then
        LogError SRC, "data mora biti 2D array."
        Exit Function
    End If

    If Len(GetAccessToken()) = 0 Then
        LogError SRC, "Kein Access Token"
        Exit Function
    End If

    stagingTab = BuildStagingTabName(tabName)
    
    If Not AddSheetTab(spreadsheetID, stagingTab, False) Then
        LogError SRC, "Could not create staging tab. Target=" & tabName & _
                      "; Staging=" & stagingTab
        Exit Function
    End If

    If Not WriteSheetValuesNoClear(spreadsheetID, stagingTab, data, SRC & ".staging") Then
        LogError SRC, "Write to staging tab failed. Target=" & tabName & _
                      "; Staging=" & stagingTab
        SafeDeleteSheetByTitle spreadsheetID, stagingTab
        Exit Function
    End If

    If Not VerifyWrittenSheetData(spreadsheetID, stagingTab, data) Then
        LogError SRC, "Verify staging tab failed. Target=" & tabName & _
                      "; Staging=" & stagingTab
        SafeDeleteSheetByTitle spreadsheetID, stagingTab
        Exit Function
    End If

    stagingSheetId = GetSheetIdByTitle(spreadsheetID, stagingTab)
    If stagingSheetId <= 0 Then
        LogError SRC, "Could not resolve staging sheetId. Staging=" & stagingTab
        SafeDeleteSheetByTitle spreadsheetID, stagingTab
        Exit Function
    End If

    targetSheetId = GetSheetIdByTitle(spreadsheetID, tabName, True)
    
    If Not ReplaceSheetTabWithStaging(spreadsheetID, tabName, targetSheetId, stagingTab, stagingSheetId) Then
        LogError SRC, "Replacing target with staging failed. Target=" & tabName & _
                      "; Staging=" & stagingTab
        SafeDeleteSheetByTitle spreadsheetID, stagingTab
        Exit Function
    End If

    If Not VerifyWrittenSheetData(spreadsheetID, tabName, data) Then
        LogError SRC, "Post-replace verify failed. Target=" & tabName
        WriteSheetData = False
        Exit Function
    End If

    LogInfo SRC, tabName & ": " & CStr(UBound(data, 1)) & _
                 " rows written through staging replace"

    WriteSheetData = True
    Exit Function

EH:
    LogErr SRC
    On Error Resume Next
    If Len(Trim$(stagingTab)) > 0 Then SafeDeleteSheetByTitle spreadsheetID, stagingTab
    On Error GoTo 0
    WriteSheetData = False
End Function


Private Function WriteSheetValuesNoClear(ByVal spreadsheetID As String, _
                                         ByVal tabName As String, _
                                         ByVal data As Variant, _
                                         ByVal sourceName As String) As Boolean
    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object

    On Error GoTo EH

    WriteSheetValuesNoClear = False

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError sourceName, "Kein Access Token"
        Exit Function
    End If

    body = BuildValuesJson(data)

    url = SHEETS_API_BASE & "/" & spreadsheetID & _
          "/values/" & UrlEncode(tabName) & "!A1" & _
          "?valueInputOption=RAW"

    Set http = CreateGoogleHttpRequest(sourceName)

    http.Open "PUT", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    If Not SendGoogleHttpWithRetry(http, sourceName, body, True, True) Then
        WriteSheetValuesNoClear = False
        Exit Function
    End If


    If http.status >= 200 And http.status < 300 Then
        WriteSheetValuesNoClear = True
    Else
        LogError sourceName, _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
    End If


    Exit Function

EH:
    LogErr sourceName
    WriteSheetValuesNoClear = False
End Function

Private Function BuildStagingTabName(ByVal targetTabName As String) As String
    Dim baseName As String

    baseName = SanitizeSheetTabName("__stage_" & targetTabName)

    If Len(baseName) > 70 Then baseName = Left$(baseName, 70)

    BuildStagingTabName = baseName & "_" & Format$(Now, "hhnnss")
End Function

Private Function BuildBackupTabName(ByVal targetTabName As String) As String
    Dim baseName As String

    baseName = SanitizeSheetTabName("__old_" & targetTabName)

    If Len(baseName) > 70 Then baseName = Left$(baseName, 70)

    BuildBackupTabName = baseName & "_" & Format$(Now, "hhnnss")
End Function

Private Function SanitizeSheetTabName(ByVal s As String) As String
    s = Trim$(CStr(s))
    s = Replace(s, "'", "_")
    s = Replace(s, "[", "_")
    s = Replace(s, "]", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "\", "_")

    If Len(s) = 0 Then s = "__stage"
    If Len(s) > 90 Then s = Left$(s, 90)

    SanitizeSheetTabName = s
End Function

Private Function VerifyWrittenSheetData(ByVal spreadsheetID As String, _
                                        ByVal tabName As String, _
                                        ByVal expectedData As Variant) As Boolean
    Const SRC As String = "VerifyWrittenSheetData"

    Dim actual As Variant
    Dim r As Long
    Dim c As Long
    Dim expectedRows As Long
    Dim expectedCols As Long

    On Error GoTo EH

    VerifyWrittenSheetData = False

    actual = ReadSheetData(spreadsheetID, tabName)
    If IsEmpty(actual) Then
        LogError SRC, "Readback is Empty. Tab=" & tabName
        Exit Function
    End If

    expectedRows = UBound(expectedData, 1)
    expectedCols = UBound(expectedData, 2)

    If UBound(actual, 1) < expectedRows Then
        LogError SRC, "Row count mismatch. Tab=" & tabName & _
                      "; Actual=" & CStr(UBound(actual, 1)) & _
                      "; Expected=" & CStr(expectedRows)
        Exit Function
    End If

    If UBound(actual, 2) < expectedCols Then
        LogError SRC, "Column count mismatch. Tab=" & tabName & _
                      "; Actual=" & CStr(UBound(actual, 2)) & _
                      "; Expected=" & CStr(expectedCols)
        Exit Function
    End If
    For r = 1 To expectedRows
        For c = 1 To expectedCols
            If GoogleSheetComparableValue(actual(r, c)) <> _
               GoogleSheetComparableValue(expectedData(r, c)) Then

                LogError SRC, "Value mismatch. Tab=" & tabName & _
                              "; Row=" & CStr(r) & _
                              "; Col=" & CStr(c) & _
                              "; Actual=" & GoogleSheetComparableValue(actual(r, c)) & _
                              "; Expected=" & GoogleSheetComparableValue(expectedData(r, c))
                Exit Function
            End If
        Next c
    Next r

    VerifyWrittenSheetData = True
    Exit Function

EH:
    LogErr SRC
    VerifyWrittenSheetData = False
End Function

Private Function GoogleSheetComparableValue(ByVal value As Variant) As String
    Dim s As String

    If IsEmpty(value) Or IsNull(value) Then
        GoogleSheetComparableValue = ""
        Exit Function
    End If

    If VarType(value) = vbDate Then
        GoogleSheetComparableValue = Format$(CDate(value), "yyyy-mm-dd")
        Exit Function
    End If

    s = CStr(value)

    ' Google / JSON readback ponekad vrati HTML-sensitive karaktere kao unicode escape.
    ' Ovo ne menja podatak u sheet-u, samo normalizuje verify poredenje.
    s = Replace(s, "\u003e", ">")
    s = Replace(s, "\u003E", ">")
    s = Replace(s, "\u003c", "<")
    s = Replace(s, "\u003C", "<")
    s = Replace(s, "\u0026", "&")
    s = Replace(s, "\u0027", "'")
    s = Replace(s, "\u003d", "=")
    s = Replace(s, "\u003D", "=")

    GoogleSheetComparableValue = s
End Function

Private Function SheetCacheKey(ByVal spreadsheetID As String) As String
    SheetCacheKey = Trim$(CStr(spreadsheetID))
End Function

Private Function GetSheetIdCache(ByVal spreadsheetID As String, _
                                 Optional ByVal forceRefresh As Boolean = False) As Object
    Const SRC As String = "GetSheetIdCache"

    Dim cacheKey As String
    Dim sheetDict As Object
    Dim metadataJson As String

    On Error GoTo EH

    cacheKey = SheetCacheKey(spreadsheetID)

    If mSheetIdCacheBySpreadsheet Is Nothing Then
        Set mSheetIdCacheBySpreadsheet = CreateObject("Scripting.Dictionary")
    End If

    If Not forceRefresh Then
        If mSheetIdCacheBySpreadsheet.Exists(cacheKey) Then
            Set GetSheetIdCache = mSheetIdCacheBySpreadsheet(cacheKey)
            Exit Function
        End If
    End If

    metadataJson = FetchSheetMetadataJson(spreadsheetID)

    If Len(Trim$(metadataJson)) = 0 Then
        Set GetSheetIdCache = Nothing
        Exit Function
    End If

    Set sheetDict = BuildSheetIdDictionaryFromMetadataJson(metadataJson)

    If mSheetIdCacheBySpreadsheet.Exists(cacheKey) Then
        Set mSheetIdCacheBySpreadsheet(cacheKey) = sheetDict
    Else
        mSheetIdCacheBySpreadsheet.Add cacheKey, sheetDict
    End If

    Set GetSheetIdCache = sheetDict
    Exit Function

EH:
    LogErr SRC
    Set GetSheetIdCache = Nothing
End Function

Private Function GetSheetIdByTitle(ByVal spreadsheetID As String, _
                                   ByVal tabName As String, _
                                   Optional ByVal forceRefresh As Boolean = False) As Long
    Const SRC As String = "GetSheetIdByTitle"

    Dim dict As Object

    On Error GoTo EH

    GetSheetIdByTitle = 0

    If Len(Trim$(spreadsheetID)) = 0 Or Len(Trim$(tabName)) = 0 Then Exit Function

    Set dict = GetSheetIdCache(spreadsheetID, forceRefresh)

    If Not dict Is Nothing Then
        If dict.Exists(tabName) Then
            GetSheetIdByTitle = CLng(dict(tabName))
            Exit Function
        End If
    End If

    If Not forceRefresh Then
        Set dict = GetSheetIdCache(spreadsheetID, True)

        If Not dict Is Nothing Then
            If dict.Exists(tabName) Then
                GetSheetIdByTitle = CLng(dict(tabName))
                Exit Function
            End If
        End If
    End If

    LogWarn SRC, "Sheet title not found. Tab=" & tabName
    Exit Function

EH:
    LogErr SRC
    GetSheetIdByTitle = 0
End Function

Private Function ExtractSheetIdByTitle(ByVal json As String, _
                                       ByVal expectedTitle As String) As Long
    Dim pos As Long
    Dim idPos As Long
    Dim titlePos As Long
    Dim nextBlock As Long
    Dim sheetIdText As String
    Dim sheetTitle As String

    pos = 1

    Do
        titlePos = InStr(pos, json, """title""", vbTextCompare)
        If titlePos = 0 Then Exit Do

        idPos = InStrRev(Left$(json, titlePos), """sheetId""", -1, vbTextCompare)
        If idPos = 0 Then Exit Do

        nextBlock = InStr(titlePos + 1, json, """title""", vbTextCompare)

        sheetTitle = ExtractJsonSimpleValueAt(json, titlePos)
        sheetIdText = ExtractJsonNumberAt(json, idPos)

        If StrComp(sheetTitle, expectedTitle, vbBinaryCompare) = 0 Then
            If IsNumeric(sheetIdText) Then
                ExtractSheetIdByTitle = CLng(sheetIdText)
                Exit Function
            End If
        End If

        If nextBlock > 0 Then
            pos = nextBlock
        Else
            Exit Do
        End If
      Loop

    ExtractSheetIdByTitle = 0
End Function

Private Function ExtractJsonNumberAt(ByVal json As String, _
                                     ByVal keyPos As Long) As String
    Dim p As Long
    Dim q As Long
    Dim ch As String

    p = InStr(keyPos, json, ":")
    If p = 0 Then Exit Function

    p = p + 1

    Do While p <= Len(json)
        ch = Mid$(json, p, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        p = p + 1
    Loop

    q = p
    Do While q <= Len(json)
        ch = Mid$(json, q, 1)
        If InStr(1, "0123456789-", ch, vbBinaryCompare) = 0 Then Exit Do
        q = q + 1
    Loop

    If q > p Then ExtractJsonNumberAt = Mid$(json, p, q - p)
End Function

Private Function FetchSheetMetadataJson(ByVal spreadsheetID As String) As String
    Const SRC As String = "FetchSheetMetadataJson"

    Dim accessToken As String
    Dim url As String
    Dim http As Object

    On Error GoTo EH

    FetchSheetMetadataJson = vbNullString

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError SRC, "Kein Access Token"
        Exit Function
    End If

    url = SHEETS_API_BASE & "/" & spreadsheetID & _
          "?fields=sheets.properties(sheetId,title)"

    Set http = CreateGoogleHttpRequest(SRC)

    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken

    If Not SendGoogleHttpWithRetry(http, SRC) Then Exit Function

    If http.status <> 200 Then
        LogError SRC, _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
        Exit Function
    End If

    FetchSheetMetadataJson = CStr(http.responseText)
    Exit Function

EH:
    LogErr SRC
    FetchSheetMetadataJson = vbNullString
End Function

Private Function BuildSheetIdDictionaryFromMetadataJson(ByVal json As String) As Object
    Dim dict As Object
    Dim pos As Long
    Dim idPos As Long
    Dim titlePos As Long
    Dim nextTitlePos As Long
    Dim sheetIdText As String
    Dim sheetTitle As String

    Set dict = CreateObject("Scripting.Dictionary")

    pos = 1

    Do
        titlePos = InStr(pos, json, """title""", vbTextCompare)
        If titlePos = 0 Then Exit Do

        idPos = InStrRev(Left$(json, titlePos), """sheetId""", -1, vbTextCompare)
        If idPos = 0 Then Exit Do

        sheetTitle = ExtractJsonSimpleValueAt(json, titlePos)
        sheetIdText = ExtractJsonNumberAt(json, idPos)

        If Len(sheetTitle) > 0 And IsNumeric(sheetIdText) Then
            If dict.Exists(sheetTitle) Then
                dict(sheetTitle) = CLng(sheetIdText)
            Else
                dict.Add sheetTitle, CLng(sheetIdText)
            End If
        End If

        nextTitlePos = InStr(titlePos + 1, json, """title""", vbTextCompare)
        If nextTitlePos > 0 Then
            pos = nextTitlePos
        Else
            Exit Do
        End If
    Loop

    Set BuildSheetIdDictionaryFromMetadataJson = dict
End Function

Private Sub CacheSheetId(ByVal spreadsheetID As String, _
                         ByVal tabName As String, _
                         ByVal sheetID As Long)
    Dim cacheKey As String
    Dim dict As Object

    If sheetID <= 0 Then Exit Sub
    If Len(Trim$(spreadsheetID)) = 0 Then Exit Sub
    If Len(Trim$(tabName)) = 0 Then Exit Sub

    cacheKey = SheetCacheKey(spreadsheetID)

    If mSheetIdCacheBySpreadsheet Is Nothing Then
        Set mSheetIdCacheBySpreadsheet = CreateObject("Scripting.Dictionary")
    End If

    If mSheetIdCacheBySpreadsheet.Exists(cacheKey) Then
        Set dict = mSheetIdCacheBySpreadsheet(cacheKey)
    Else
        Set dict = CreateObject("Scripting.Dictionary")
        mSheetIdCacheBySpreadsheet.Add cacheKey, dict
    End If

    If dict.Exists(tabName) Then
        dict(tabName) = sheetID
    Else
        dict.Add tabName, sheetID
    End If
End Sub

Private Sub RemoveSheetIdFromCache(ByVal spreadsheetID As String, ByVal tabName As String)
    Dim cacheKey As String
    Dim dict As Object

    If mSheetIdCacheBySpreadsheet Is Nothing Then Exit Sub

    cacheKey = SheetCacheKey(spreadsheetID)
    If Not mSheetIdCacheBySpreadsheet.Exists(cacheKey) Then Exit Sub

    Set dict = mSheetIdCacheBySpreadsheet(cacheKey)
    If dict.Exists(tabName) Then dict.Remove tabName
End Sub

Private Function ExtractFirstSheetIdFromResponse(ByVal responseText As String) As Long
    Dim p As Long
    Dim n As String

    p = InStr(1, responseText, """sheetId""", vbTextCompare)
    If p = 0 Then Exit Function

    n = ExtractJsonNumberAt(responseText, p)
    If IsNumeric(n) Then ExtractFirstSheetIdFromResponse = CLng(n)
End Function

Private Function ReplaceSheetTabWithStaging(ByVal spreadsheetID As String, _
                                            ByVal targetTabName As String, _
                                            ByVal targetSheetId As Long, _
                                            ByVal stagingTabName As String, _
                                            ByVal stagingSheetId As Long) As Boolean
    Const SRC As String = "ReplaceSheetTabWithStaging"

    Dim backupTabName As String

    On Error GoTo EH

    ReplaceSheetTabWithStaging = False

    If Len(Trim$(spreadsheetID)) = 0 Then
        LogError SRC, "spreadsheetID je prazan."
        Exit Function
    End If

    If Len(Trim$(targetTabName)) = 0 Then
        LogError SRC, "targetTabName je prazan."
        Exit Function
    End If

    If Len(Trim$(stagingTabName)) = 0 Then
        LogError SRC, "stagingTabName je prazan."
        Exit Function
    End If

    If stagingSheetId <= 0 Then
        LogError SRC, "Invalid stagingSheetId. Staging=" & stagingTabName
        Exit Function
    End If
       If targetSheetId = stagingSheetId Then
        LogError SRC, "Invalid replace: targetSheetId equals stagingSheetId. Target=" & _
                      targetTabName & "; Staging=" & stagingTabName & _
                      "; SheetId=" & CStr(stagingSheetId)
        Exit Function
    End If

    If targetSheetId <= 0 Then
        targetSheetId = GetSheetIdByTitle(spreadsheetID, targetTabName, True)
    End If

    If targetSheetId <= 0 Then
        LogError SRC, "Target sheetId could not be resolved after forced refresh. " & _
                      "Refusing direct rename to avoid title collision. Target=" & targetTabName & _
                      "; Staging=" & stagingTabName
        Exit Function
    End If

    backupTabName = BuildBackupTabName(targetTabName)

    ' Phase 1: old target keeps data, but gets temporary backup title.
    If Not RenameSheetById(spreadsheetID, targetSheetId, backupTabName, SRC & ".renameTargetToBackup") Then
        LogError SRC, "Could not rename target to backup. Target=" & targetTabName & _
                      "; Backup=" & backupTabName
        Exit Function
    End If

    RemoveSheetIdFromCache spreadsheetID, targetTabName
    CacheSheetId spreadsheetID, backupTabName, targetSheetId
    
    ' Phase 2: staging becomes real target.
    If Not RenameSheetById(spreadsheetID, stagingSheetId, targetTabName, SRC & ".renameStagingToTarget") Then
        LogError SRC, "Could not rename staging to target. Attempting recovery. Target=" & _
                      targetTabName & "; Staging=" & stagingTabName & "; Backup=" & backupTabName

        ' Best-effort recovery: restore old target title.
        If RenameSheetById(spreadsheetID, targetSheetId, targetTabName, SRC & ".recoverBackupToTarget") Then
            CacheSheetId spreadsheetID, targetTabName, targetSheetId
            RemoveSheetIdFromCache spreadsheetID, backupTabName
        Else
            LogError SRC, "Recovery failed. Old target still exists as backup name: " & backupTabName
        End If

        Exit Function
    End If

    CacheSheetId spreadsheetID, targetTabName, stagingSheetId
    RemoveSheetIdFromCache spreadsheetID, stagingTabName

    ' Phase 3: delete old target backup. If this fails, data is safe:
    ' new target is already live, old backup tab may remain and can be cleaned later.
    If Not DeleteSheetById(spreadsheetID, targetSheetId) Then
        LogWarn SRC, "Target replaced, but backup delete failed. Manual cleanup may be needed. Backup=" & backupTabName
        ' Success is still true because the target tab now contains the verified staging data.
        ReplaceSheetTabWithStaging = True
        Exit Function
    End If

    RemoveSheetIdFromCache spreadsheetID, backupTabName

    ReplaceSheetTabWithStaging = True
    Exit Function

EH:
    LogErr SRC
    ReplaceSheetTabWithStaging = False
End Function

Private Function ExecuteSheetsBatchUpdate(ByVal spreadsheetID As String, _
                                          ByVal body As String, _
                                          ByVal sourceName As String) As Boolean
    Dim accessToken As String
    Dim url As String
    Dim http As Object

    On Error GoTo EH

    ExecuteSheetsBatchUpdate = False

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError sourceName, "Kein Access Token"
        Exit Function
    End If

    url = SHEETS_API_BASE & "/" & spreadsheetID & ":batchUpdate"

    Set http = CreateGoogleHttpRequest(sourceName)

    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"

    If Not SendGoogleHttpWithRetry(http, sourceName, body, True, True) Then Exit Function

    If http.status >= 200 And http.status < 300 Then
        ExecuteSheetsBatchUpdate = True
    Else
        LogError sourceName, _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
    End If

    Exit Function
EH:
    LogErr sourceName
    ExecuteSheetsBatchUpdate = False
End Function

Private Sub SafeDeleteSheetByTitle(ByVal spreadsheetID As String, _
                                   ByVal tabName As String)
    On Error Resume Next

    Dim sheetID As Long
    sheetID = GetSheetIdByTitle(spreadsheetID, tabName)

    If sheetID > 0 Then
        If DeleteSheetById(spreadsheetID, sheetID) Then
            RemoveSheetIdFromCache spreadsheetID, tabName
        End If
    End If

    On Error GoTo 0
End Sub

Private Function RenameSheetById(ByVal spreadsheetID As String, _
                                 ByVal sheetID As Long, _
                                 ByVal newTitle As String, _
                                 ByVal sourceName As String) As Boolean
    Dim body As String

    RenameSheetById = False

    If sheetID <= 0 Then
        LogError sourceName, "Invalid sheetId for rename. NewTitle=" & newTitle
        Exit Function
    End If

    If Len(Trim$(newTitle)) = 0 Then
        LogError sourceName, "newTitle je prazan. SheetId=" & CStr(sheetID)
        Exit Function
    End If

    body = "{""requests"":[" & _
           "{""updateSheetProperties"":{""properties"":{""sheetId"":" & CStr(sheetID) & _
           ",""title"":""" & JsonEscape(newTitle) & """},""fields"":""title""}}" & _
           "]}"

    RenameSheetById = ExecuteSheetsBatchUpdate(spreadsheetID, body, sourceName)
End Function

Private Function DeleteSheetById(ByVal spreadsheetID As String, _
                                 ByVal sheetID As Long) As Boolean
    Dim body As String

    If sheetID <= 0 Then Exit Function

    body = "{""requests"":[{""deleteSheet"":{""sheetId"":" & CStr(sheetID) & "}}]}"
    DeleteSheetById = ExecuteSheetsBatchUpdate(spreadsheetID, body, "DeleteSheetById")
End Function

Public Sub CleanupGoogleSheetBackupTabs(ByVal spreadsheetID As String)
    Const SRC As String = "CleanupGoogleSheetBackupTabs"

    On Error GoTo EH

    LogWarn SRC, "Manual cleanup helper is intentionally not implemented as automatic delete. " & _
                 "Use Google UI or add explicit allowlist before deleting __old_* tabs. SpreadsheetID=" & spreadsheetID
    Exit Sub

EH:
    LogErr SRC
End Sub

' ============================================================
' Public -- Append single row to existing tab.
' Koristi Sheets API values:append endpoint sa INSERT_ROWS.
'
' Razlika u odnosu na WriteSheetData (full-tab replace via staging+swap):
'   - ne pravi staging, ne brise postojece redove
'   - append-only, atomska operacija ~300-500ms
'   - rowData je 1D Array (jedna vrsta), interno se wrap-uje u 2D
'
' Caller (modStanicaLock.BulkPushPendingForStanica) zove ovo u petlji
' kad se stanica unlock-uje -- svaki red se append-uje pojedinacno.
' Pojedinacni fail ostavlja red sa praznim ClientRecordID u tblOtkup-u,
' pa retry pri sledecem unlock event-u pokupi.
'
' Vraca True samo na HTTP 2xx; ostale slucajeve loguje i vraca False.
' ============================================================
Public Function AppendRowToSheet(ByVal spreadsheetID As String, _
                                 ByVal tabName As String, _
                                 ByVal rowData As Variant) As Boolean
    Const SRC As String = "AppendRowToSheet"

    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object
    Dim wrappedData As Variant
    Dim n As Long
    Dim i As Long

    On Error GoTo EH

    AppendRowToSheet = False

    If Len(Trim$(spreadsheetID)) = 0 Then
        LogError SRC, "spreadsheetID je prazan."
        Exit Function
    End If

    If Len(Trim$(tabName)) = 0 Then
        LogError SRC, "tabName je prazan."
        Exit Function
    End If

    If IsEmpty(rowData) Then
        LogError SRC, "rowData je Empty."
        Exit Function
    End If

    If Not IsArray(rowData) Then
        LogError SRC, "rowData mora biti Array (1D)."
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError SRC, "Kein Access Token"
        Exit Function
    End If

    ' Wrap 1D rowData u 2D (1 x N) jer BuildValuesJson ocekuje 2D shape.
    n = UBound(rowData) - LBound(rowData) + 1
    If n <= 0 Then
        LogError SRC, "rowData je prazan."
        Exit Function
    End If

    ReDim wrappedData(1 To 1, 1 To n)
    For i = 0 To n - 1
        wrappedData(1, i + 1) = rowData(LBound(rowData) + i)
    Next i

    body = BuildValuesJson(wrappedData)

    ' Sheets API values:append -- INSERT_ROWS pravi NOV red (ne overwrite
    ' praznih redova ispod date). RAW cuva format brojeva sa "/" znakovima
    ' (12/1 ambalaza, 1/220526-2 broj dokumenta).
    url = SHEETS_API_BASE & "/" & spreadsheetID & _
          "/values/" & UrlEncode(tabName) & "!A1:append" & _
          "?valueInputOption=RAW" & _
          "&insertDataOption=INSERT_ROWS"

    Set http = CreateGoogleHttpRequest(SRC)
    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"

    If Not SendGoogleHttpWithRetry(http, SRC, body, True, True) Then
        AppendRowToSheet = False
        Exit Function
    End If

    If http.status >= 200 And http.status < 300 Then
        AppendRowToSheet = True
    Else
        LogError SRC, _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
        AppendRowToSheet = False
    End If

    Exit Function

EH:
    LogErr SRC
    AppendRowToSheet = False
End Function

' ============================================================
' 8) Full-sync gate note
' ============================================================
'
' No broad orchestrator rewrite is needed for the minimum acceptance because:
' - SyncStammdatenToGoogle_Core counts a tab only if ExportX returns True.
' - ExportKarticeToGoogle_Core returns WriteSheetData result.
' - ExportMgmtReports_Core counts a report tab only if ExportX returns True.
' - SyncPWAFullCycle_Core computes final result from okStammdaten, okKartice,
'   okMgmt and other gates.
'
' Therefore any WriteSheetData=False must propagate as partial/fail full sync,
' assuming callers do not swallow the return value.
'
' Recommended smoke:
' - Force WriteSheetValuesNoClear to return False after staging tab creation.
'   Expected: target tab remains unchanged; staging tab is deleted; full sync False.
' - Force ReplaceSheetTabWithStaging to return False.
'   Expected: target tab remains unchanged; staging tab cleanup attempted; full sync False.
' - Normal export.
'   Expected: staging tab disappears, target tab contains verified new data.
' ============================================================

    

' ============================================================
' PUBLIC -- Read
' ============================================================

Public Function ReadSheetData(ByVal spreadsheetID As String, _
                              ByVal tabName As String) As Variant
    ' Kompatibilitaets-Wrapper: Fehler und "leer" sind hier NICHT
    ' unterscheidbar (beides Empty). Fuer Sync-Importe MUSS
    ' TryReadSheetData verwendet werden.
    Dim data As Variant

    If TryReadSheetData(spreadsheetID, tabName, data) Then
        ReadSheetData = data
    Else
        ReadSheetData = Empty
    End If
End Function

Public Function TryReadSheetData(ByVal spreadsheetID As String, _
                                 ByVal tabName As String, _
                                 ByRef outData As Variant) As Boolean
    ' FAIL-CLOSED Read.
    '
    '   True  + outData = Array  -> Tab gelesen, Zeilen vorhanden
    '   True  + outData = Empty  -> Tab gelesen, aber leer
    '   False                    -> Argument-/Token-/HTTP-/Parse-Fehler.
    '                               Aufrufer MUSS abbrechen: kein lokaler
    '                               Import und kein Google Writeback,
    '                               sonst gehen Zeilen dauerhaft verloren.
    Dim accessToken As String
    Dim url As String
    Dim http As Object

    outData = Empty
    TryReadSheetData = False

    On Error GoTo EH

    If Not RequireGoogleTextArg(spreadsheetID, "spreadsheetID", "TryReadSheetData") Then Exit Function
    If Not RequireGoogleTextArg(tabName, "tabName", "TryReadSheetData") Then Exit Function

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "TryReadSheetData", "Kein Access Token"
        Exit Function
    End If

    url = SHEETS_API_BASE & "/" & spreadsheetID & _
          "/values/" & UrlEncode(tabName)

    Set http = CreateGoogleHttpRequest("TryReadSheetData")

    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    If Not SendGoogleHttpWithRetry(http, "TryReadSheetData") Then Exit Function

    If http.status <> 200 Then
        LogError "TryReadSheetData", _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
        Exit Function
    End If

    If Not TryParseValuesJson(http.responseText, outData) Then
        outData = Empty
        LogError "TryReadSheetData", _
                 "Defektes/abgeschnittenes JSON von Sheets API. Tab=" & tabName & _
                 "; Bytes=" & CStr(Len(http.responseText)) & _
                 "; Body=" & GoogleHttpBodyForLog(http.responseText)
        Exit Function
    End If

    TryReadSheetData = True
    Exit Function

EH:
    outData = Empty
    LogErr "TryReadSheetData"
    TryReadSheetData = False
End Function

' ============================================================
' PUBLIC -- Clear
' ============================================================

Public Function ClearSheet(ByVal spreadsheetID As String, _
                           ByVal tabName As String) As Boolean
    Dim accessToken As String
    Dim url As String
    Dim http As Object

    On Error GoTo EH

    If Not RequireGoogleTextArg(spreadsheetID, "spreadsheetID", "ClearSheet") Then
        ClearSheet = False
        Exit Function
    End If

    If Not RequireGoogleTextArg(tabName, "tabName", "ClearSheet") Then
        ClearSheet = False
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "ClearSheet", "Kein Access Token"
        ClearSheet = False
        Exit Function
    End If

    url = SHEETS_API_BASE & "/" & spreadsheetID & _
          "/values/" & UrlEncode(tabName) & ":clear"

    Set http = CreateGoogleHttpRequest("ClearSheet")

    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    If Not SendGoogleHttpWithRetry(http, "ClearSheet", "{}", True, True) Then
        ClearSheet = False
        Exit Function
    End If

    If http.status >= 200 And http.status < 300 Then
        ClearSheet = True
    Else
        LogError "ClearSheet", _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
        ClearSheet = False
    End If

    Exit Function

EH:
    LogErr "ClearSheet"
    ClearSheet = False
End Function

' ============================================================
' PUBLIC -- Create Spreadsheet
' ============================================================

Public Function CreateSpreadsheet(ByVal title As String, _
                                  Optional ByVal folderID As String = "") As String
    ' Erstellt ein neues Google Sheet, gibt SpreadsheetID zurueck
    ' Wenn folderID angegeben, wird es in den Folder verschoben
    
    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object
    Dim newID As String
    
    On Error GoTo EH
    
    If Not RequireGoogleTextArg(title, "title", "CreateSpreadsheet") Then
        CreateSpreadsheet = ""
        Exit Function
    End If
    
    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        CreateSpreadsheet = ""
        Exit Function
    End If
    
    url = SHEETS_API_BASE
    body = "{""properties"":{""title"":""" & JsonEscape(title) & """}}"
    
    Set http = CreateGoogleHttpRequest("CreateSpreadsheet")
    
    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body
    
    If http.status <> 200 Then
        LogError "CreateSpreadsheet", _
            "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
            http.status
        CreateSpreadsheet = ""
        Exit Function
    End If
    
    newID = ExtractJsonStringGoogle(http.responseText, "spreadsheetId")
    
    If Len(Trim$(newID)) = 0 Then
        LogError "CreateSpreadsheet", _
             "Google response did not contain spreadsheetId: " & GoogleHttpBodyForLog(http.responseText)
        CreateSpreadsheet = ""
        Exit Function
    End If
    
    ' In Folder verschieben wenn angegeben
    If Len(Trim$(folderID)) > 0 And Len(newID) > 0 Then
        If Not MoveFileToFolder(newID, folderID) Then
            LogWarn "CreateSpreadsheet", _
                "Spreadsheet created but move to folder failed. Title=" & title & _
                ", SpreadsheetID=" & newID & _
                ", FolderID=" & folderID
        End If
    End If

    LogInfo "CreateSpreadsheet", "Created: " & title & " (" & newID & ")"
    CreateSpreadsheet = newID
    Exit Function

EH:
    LogErr "CreateSpreadsheet"
    CreateSpreadsheet = ""
End Function

' ============================================================
' PUBLIC -- Find Spreadsheet by Name in Folder
' ============================================================

Public Function GetSpreadsheetID(ByVal title As String, _
                                 Optional ByVal folderID As String = "") As String
    ' Kompatibilitaets-Wrapper: "nicht gefunden" und "Lookup fehlgeschlagen"
    ' sind hier NICHT unterscheidbar (beides ""). Wer daraus eine Entscheidung
    ' ableitet (z.B. Nummernvergabe), MUSS TryGetSpreadsheetID nutzen.
    Dim foundID As String

    If TryGetSpreadsheetID(title, folderID, foundID) Then
        GetSpreadsheetID = foundID
    Else
        GetSpreadsheetID = ""
    End If
End Function

Public Function TryGetSpreadsheetID(ByVal title As String, _
                                    ByVal folderID As String, _
                                    ByRef outID As String) As Boolean
    ' FAIL-CLOSED Drive Lookup.
    '
    '   True  + outID = ID  -> Spreadsheet gefunden
    '   True  + outID = ""  -> Lookup ok, Spreadsheet existiert wirklich nicht
    '   False               -> Argument-/Token-/HTTP-/Parse-Fehler oder Exception.
    '                          Aufrufer darf daraus NICHT "existiert nicht"
    '                          schliessen (AUD-001: sonst wird ein bereits
    '                          vergebener Beleg-Nummernbereich uebersehen).
    Dim accessToken As String
    Dim url As String
    Dim http As Object
    Dim query As String
    Dim responseText As String

    outID = ""
    TryGetSpreadsheetID = False

    On Error GoTo EH

    If Not RequireGoogleTextArg(title, "title", "TryGetSpreadsheetID") Then Exit Function

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "TryGetSpreadsheetID", "Kein Access Token"
        Exit Function
    End If

    query = "name='" & EscapeDriveQueryValue(title) & _
            "' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"

    If Len(Trim$(folderID)) > 0 Then
        query = query & " and '" & EscapeDriveQueryValue(folderID) & "' in parents"
    End If

    url = DRIVE_API_BASE & "/files?q=" & UrlEncode(query) & _
          "&fields=files(id,name)&pageSize=10"

    Set http = CreateGoogleHttpRequest("TryGetSpreadsheetID")

    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.Send

    responseText = CStr(http.responseText)

    If http.status <> 200 Then
        LogError "TryGetSpreadsheetID", _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(responseText), _
                 http.status
        Exit Function
    End If

    ' Unlesbare/abgeschnittene Drive-Antwort darf nicht als
    ' "nicht gefunden" durchgehen.
    If Not IsWellFormedJsonDocument(responseText) Then
        LogError "TryGetSpreadsheetID", _
                 "Defektes/abgeschnittenes JSON von Drive API. Title=" & title & _
                 "; Bytes=" & CStr(Len(responseText))
        Exit Function
    End If

    outID = ExtractSpreadsheetIDByExactName(responseText, title)

    If Len(Trim$(outID)) = 0 Then
        outID = ""
        LogInfo "TryGetSpreadsheetID", "Spreadsheet not found by exact name: " & title
    End If

    TryGetSpreadsheetID = True
    Exit Function

EH:
    outID = ""
    LogErr "TryGetSpreadsheetID"
    TryGetSpreadsheetID = False
End Function

Public Function TryGetOrCreateSpreadsheetID(ByVal title As String, _
                                            ByVal folderID As String, _
                                            ByRef outID As String) As Boolean
    ' FAIL-CLOSED get-or-create.
    '
    '   True  + outID -> Spreadsheet gefunden oder neu erstellt
    '   False         -> Lookup fehlgeschlagen (dann wird NICHTS erstellt)
    '                    oder Erstellen fehlgeschlagen
    '
    ' AUD-001: Der alte Ablauf war
    '     id = GetSpreadsheetID(...) : If id = "" Then id = CreateSpreadsheet(...)
    ' und hat einen HTTP-/JSON-Fehler beim Lookup als "existiert nicht"
    ' gelesen -> zweites Spreadsheet mit gleichem Namen (Duplikat), im
    ' schlimmsten Fall fuer viele Stanice gleichzeitig.
    Dim foundID As String

    outID = ""
    TryGetOrCreateSpreadsheetID = False

    If Not TryGetSpreadsheetID(title, folderID, foundID) Then
        LogError "TryGetOrCreateSpreadsheetID", _
                 "Drive lookup nije uspeo -- sheet se NE kreira (rizik od duplikata). Title=" & title
        Exit Function
    End If

    If Len(Trim$(foundID)) > 0 Then
        outID = foundID
        TryGetOrCreateSpreadsheetID = True
        Exit Function
    End If

    ' Lookup je prosao i sheet stvarno ne postoji -> bezbedno je kreirati.
    foundID = CreateSpreadsheet(title, folderID)

    If Len(Trim$(foundID)) = 0 Then Exit Function   ' CreateSpreadsheet loguje interno

    outID = foundID
    TryGetOrCreateSpreadsheetID = True
End Function

Public Function IsWellFormedJsonDocument(ByVal json As String) As Boolean
    ' Strukturpruefung (Klammern/Strings/Escapes) fuer beliebige
    ' Google-API-Antworten -- erkennt abgeschnittene und beschaedigte Bodies.
    Dim rowsColl As Collection
    Dim foundValues As Boolean

    If FirstNonWhitespaceChar(json) <> "{" Then Exit Function

    IsWellFormedJsonDocument = TryScanJsonDocument(json, rowsColl, foundValues)
End Function

Private Function EscapeDriveQueryValue(ByVal value As String) As String
    Dim result As String

    result = CStr(value)
    result = Replace(result, "\", "\\")
    result = Replace(result, "'", "\'")

    EscapeDriveQueryValue = result
End Function

Private Function ExtractSpreadsheetIDByExactName(ByVal json As String, _
                                                 ByVal expectedName As String) As String
    Dim pos As Long
    Dim idPos As Long
    Dim namePos As Long
    Dim fileID As String
    Dim fileName As String

    pos = 1

    Do
        idPos = InStr(pos, json, """id""", vbTextCompare)
        If idPos = 0 Then Exit Do

        fileID = ExtractJsonSimpleValueAt(json, idPos)

        namePos = InStr(idPos, json, """name""", vbTextCompare)
        If namePos = 0 Then Exit Do

        fileName = ExtractJsonSimpleValueAt(json, namePos)

        If Len(fileID) > 0 And StrComp(fileName, expectedName, vbBinaryCompare) = 0 Then
            ExtractSpreadsheetIDByExactName = fileID
            Exit Function
        End If

        pos = namePos + 1
    Loop

    ExtractSpreadsheetIDByExactName = ""
End Function

Private Function ExtractJsonSimpleValueAt(ByVal json As String, _
                                          ByVal keyPos As Long) As String
    Dim p As Long
    Dim q As Long

    p = InStr(keyPos, json, ":")
    If p = 0 Then Exit Function

    p = InStr(p, json, """")
    If p = 0 Then Exit Function

    p = p + 1
    q = InStr(p, json, """")

    If q > p Then
        ExtractJsonSimpleValueAt = Mid$(json, p, q - p)
    Else
        ExtractJsonSimpleValueAt = ""
    End If
End Function

' ============================================================
' PUBLIC -- Add Tab to existing Spreadsheet
' ============================================================

Public Function AddSheetTab(ByVal spreadsheetID As String, _
                            ByVal tabName As String, _
                            Optional ByVal checkExisting As Boolean = True) As Boolean
    Const SRC As String = "AddSheetTab"

    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object
    Dim existingSheetId As Long
    Dim newSheetId As Long

    On Error GoTo EH

    If Not RequireGoogleTextArg(spreadsheetID, "spreadsheetID", SRC) Then
        AddSheetTab = False
        Exit Function
    End If

    If Not RequireGoogleTextArg(tabName, "tabName", SRC) Then
        AddSheetTab = False
        Exit Function
    End If

    If checkExisting Then
        existingSheetId = GetSheetIdByTitle(spreadsheetID, tabName)
        If existingSheetId > 0 Then
            LogInfo SRC, "Tab already exists, treated as OK: " & tabName
            AddSheetTab = True
            Exit Function
        End If
    End If

   accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError SRC, "Kein Access Token"
        AddSheetTab = False
        Exit Function
    End If

    url = SHEETS_API_BASE & "/" & spreadsheetID & ":batchUpdate"
    body = "{""requests"":[{""addSheet"":{""properties"":{""title"":""" & _
           JsonEscape(tabName) & """}}}]}"

    Set http = CreateGoogleHttpRequest(SRC)

    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"

    If Not SendGoogleHttpWithRetry(http, SRC, body, True, True) Then
        AddSheetTab = False
        Exit Function
    End If

    If http.status >= 200 And http.status < 300 Then
        newSheetId = ExtractFirstSheetIdFromResponse(CStr(http.responseText))
        If newSheetId > 0 Then CacheSheetId spreadsheetID, tabName, newSheetId

        AddSheetTab = True
        Exit Function
    End If

    If IsGoogleSheetAlreadyExistsError(CLng(http.status), CStr(http.responseText)) Then
        LogInfo SRC, "Tab already exists, treated as OK: " & tabName
        ' Force refresh because this tab exists but cache did not know it.
        existingSheetId = GetSheetIdByTitle(spreadsheetID, tabName, True)
        AddSheetTab = True
        Exit Function
    End If

    LogError SRC, _
             "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
             http.status

    AddSheetTab = False
    Exit Function

EH:
    LogErr SRC
    AddSheetTab = False
End Function

' ============================================================
' PRIVATE -- Move file to folder (Drive API)
' ============================================================

Private Function MoveFileToFolder(ByVal fileID As String, ByVal folderID As String) As Boolean
    Dim accessToken As String
    Dim url As String
    Dim http As Object
    Dim parentsJson As String
    Dim oldParent As String

    On Error GoTo EH

    If Not RequireGoogleTextArg(fileID, "fileID", "MoveFileToFolder") Then
        MoveFileToFolder = False
        Exit Function
    End If

    If Not RequireGoogleTextArg(folderID, "folderID", "MoveFileToFolder") Then
        MoveFileToFolder = False
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "MoveFileToFolder", "Kein Access Token"
        MoveFileToFolder = False
        Exit Function
    End If

    ' Get current parents
    url = DRIVE_API_BASE & "/files/" & fileID & "?fields=parents"

    Set http = CreateGoogleHttpRequest("MoveFileToFolder.GetParents")
    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.Send

    parentsJson = CStr(http.responseText)

    If http.status <> 200 Then
        LogError "MoveFileToFolder", _
                 "Get parents failed. HTTP " & http.status & ": " & GoogleHttpBodyForLog(parentsJson), _
                 http.status
        MoveFileToFolder = False
        Exit Function
    End If

    oldParent = GetFirstParent(parentsJson)

    If Len(Trim$(oldParent)) = 0 Then
        LogWarn "MoveFileToFolder", _
                "No current parent detected for fileID=" & fileID & ". Adding new parent without removeParents."
        url = DRIVE_API_BASE & "/files/" & fileID & _
              "?addParents=" & UrlEncode(folderID) & _
              "&fields=id,parents"
    Else
        url = DRIVE_API_BASE & "/files/" & fileID & _
              "?addParents=" & UrlEncode(folderID) & _
              "&removeParents=" & UrlEncode(oldParent) & _
              "&fields=id,parents"
    End If

    Set http = CreateGoogleHttpRequest("MoveFileToFolder.PatchParents")
    http.Open "PATCH", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send "{}"

    If http.status >= 200 And http.status < 300 Then
        MoveFileToFolder = True
    Else
        LogError "MoveFileToFolder", _
                 "Patch parents failed. HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
        MoveFileToFolder = False
    End If

    Exit Function

EH:
    LogErr "MoveFileToFolder"
    MoveFileToFolder = False
End Function

Private Function GetFirstParent(ByVal json As String) As String
    Dim p As Long
    Dim bracketPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    p = InStr(1, json, """parents""", vbTextCompare)
    If p = 0 Then Exit Function

    bracketPos = InStr(p, json, "[")
    If bracketPos = 0 Then Exit Function

    quoteStart = InStr(bracketPos, json, """")
    If quoteStart = 0 Then Exit Function

    quoteStart = quoteStart + 1
    quoteEnd = InStr(quoteStart, json, """")

    If quoteEnd > quoteStart Then
        GetFirstParent = Mid$(json, quoteStart, quoteEnd - quoteStart)
    End If
End Function

' ============================================================
' PRIVATE -- JSON Builder fuer Sheets API
' ============================================================

Private Function BuildValuesJson(ByVal data As Variant) As String
    ' Baut JSON body fuer values:update API
    ' {"values":[["a","b"],["c","d"]]}
    ' ALLES als String schreiben uz valueInputOption=RAW.
    ' Ovo cuva vodece nule i ID vrednosti kao tekst.
    ' Ako treba Google parsiranje brojeva/datuma, koristiti USER_ENTERED
    ' i selektivno emitovati numericke JSON vrednosti.
    
    Dim sb As String
    Dim i As Long, j As Long
    Dim val As Variant
    Dim sVal As String
    
    sb = "{""values"":["
    
    For i = LBound(data, 1) To UBound(data, 1)
        If i > LBound(data, 1) Then sb = sb & ","
        sb = sb & "["
        
        For j = LBound(data, 2) To UBound(data, 2)
            If j > LBound(data, 2) Then sb = sb & ","
            
            val = data(i, j)
            
            If IsEmpty(val) Or IsNull(val) Then
                sVal = ""
            ElseIf VarType(val) = vbDate Then
                sVal = Format$(CDate(val), "yyyy-mm-dd")
            Else
                sVal = CStr(val)
            End If
            
            sb = sb & """" & JsonEscape(sVal) & """"
        Next j
        
        sb = sb & "]"
    Next i
    
    sb = sb & "]}"
    BuildValuesJson = sb
End Function

Public Function ParseValuesJson(ByVal json As String) As Variant
    ' Parst {"range":...,"values":[[...],[...]]} in ein 1-basiertes 2D Array.
    '
    ' AUD-001: Frueher wurde der JSON-Text vor dem Parsen global normalisiert
    ' (Replace ", " -> ",", "[ " -> "[", " ]" -> "]") und danach an "],["
    ' bzw. an Kommas gesplittet. Das war NICHT quote-aware und hat jede
    ' Textzelle beschaedigt, die ", " enthielt (Adressen, Namen, Napomena),
    ' escaped Quotes falsch behandelt und \uXXXX nie dekodiert.
    '
    ' Jetzt: EIN stateful Scanner (ScanJsonValuesArray) fuer Zeilen UND Zellen,
    ' der Quotes/Backslash-Escapes korrekt verfolgt und alle JSON-Escapes
    ' in einem Durchgang dekodiert.
    Dim data As Variant

    If TryParseValuesJson(json, data) Then
        ParseValuesJson = data
    Else
        ParseValuesJson = Empty
    End If
End Function

Public Function TryParseValuesJson(ByVal json As String, _
                                   ByRef outData As Variant) As Boolean
    ' FAIL-CLOSED Variante von ParseValuesJson.
    '
    '   True  + outData = Array  -> gueltige Antwort mit Zeilen
    '   True  + outData = Empty  -> gueltige, aber leere Antwort
    '                               (kein "values" Key = leerer Range)
    '   False                    -> defektes / abgeschnittenes JSON.
    '                               Aufrufer MUSS das als Fehler behandeln
    '                               (kein Import, kein Google Writeback).
    '
    ' AUD-001: Frueher wurden Teilzeilen aus abgeschnittenem JSON als
    ' gueltige Daten zurueckgegeben. Ein Import haette Zeilen mit
    ' fehlenden Endfeldern lokal committen und die Google-Zeile als
    ' Synced>Master markieren koennen -> Felder dauerhaft verloren.
    Dim rowsColl As Collection
    Dim cellsColl As Collection
    Dim foundValues As Boolean
    Dim rowCount As Long
    Dim colCount As Long
    Dim result() As Variant
    Dim i As Long, j As Long

    outData = Empty
    TryParseValuesJson = False

    ' Sanity: Sheets values.get liefert immer ein JSON-Objekt.
    ' (HTML/Proxy-Fehlerseite mit HTTP 200 darf nicht als "leer" durchgehen.)
    If FirstNonWhitespaceChar(json) <> "{" Then Exit Function

    ' Das GANZE Dokument muss balanciert sein -- nicht nur das values-Array.
    ' Sonst gilt z.B. ein abgeschnittenes {"range":"Sheet1 als "leerer Sheet".
    If Not TryScanJsonDocument(json, rowsColl, foundValues) Then Exit Function

    If Not foundValues Then
        ' Gueltiges Objekt ohne "values" -> leerer Range (Google laesst den
        ' Key bei leerem Bereich komplett weg).
        TryParseValuesJson = True
        Exit Function
    End If

    rowCount = rowsColl.count
    If rowCount = 0 Then
        ' "values":[] -> gueltig und leer
        TryParseValuesJson = True
        Exit Function
    End If

    ' Breite wie bisher aus der ersten Zeile (Header) -- unveraendertes Verhalten.
    colCount = rowsColl(1).count
    If colCount < 1 Then colCount = 1

    ReDim result(1 To rowCount, 1 To colCount)

    For i = 1 To rowCount
        Set cellsColl = rowsColl(i)
        For j = 1 To cellsColl.count
            If j <= colCount Then
                result(i, j) = cellsColl(j)
            End If
        Next j
    Next i

    outData = result
    TryParseValuesJson = True
End Function

Private Function TryScanJsonDocument(ByRef json As String, _
                                     ByRef outRows As Collection, _
                                     ByRef outFoundValues As Boolean) As Boolean
    ' Stateful Scanner ueber das GANZE JSON-Dokument.
    ' Liefert nebenbei die Zeilen des "values"-Arrays
    ' (Collection von Collections, Zellen als String).
    '
    ' Regeln:
    '   - "\" escaped das naechste Zeichen, "\""" toggelt inQuote also NICHT
    '   - Whitespace ausserhalb von Strings wird ignoriert (prettyPrint JSON)
    '   - "," / "]" trennen nur ausserhalb von Strings
    '   - "values" wird nur als TOP-LEVEL KEY erkannt (Stack = "{" und
    '     direkt gefolgt von ":"), nie als Zellinhalt
    '
    ' FAIL-CLOSED: True nur wenn
    '   - der Klammer-Stack sauber leer ist (jede "]" schliesst ein "[",
    '     jede "}" ein "{"),
    '   - das Root-Objekt geschlossen wurde und danach nur Whitespace folgt,
    '   - kein String offen ist,
    '   - keine Zeile offen ist,
    '   - ein gefundenes values-Array auch geschlossen wurde,
    '   - jeder Escape gueltig ist (\uXXXX mit genau 4 Hex-Ziffern).
    ' Eine angefangene Zeile wird NIE uebernommen.
    '
    ' Das ist bewusst KEIN vollstaendiger JSON-Validator (Kommas/Typen werden
    ' nicht geprueft) -- Ziel ist das Erkennen von abgeschnittenen und
    ' beschaedigten Antworten.
    Dim rowsColl As Collection
    Dim cellsColl As Collection
    Dim i As Long
    Dim n As Long
    Dim stack As String
    Dim inQuote As Boolean
    Dim inRow As Boolean
    Dim cellStarted As Boolean
    Dim rootClosed As Boolean
    Dim inValues As Boolean
    Dim valuesClosed As Boolean
    Dim expectValues As Boolean
    Dim valuesRoot As Long
    Dim lastString As String
    Dim ch As String
    Dim esc As String
    Dim hex4 As String
    Dim buf As String

    Set outRows = Nothing
    outFoundValues = False

    Set rowsColl = New Collection

    n = Len(json)
    i = 1

    Do While i <= n
        ch = Mid$(json, i, 1)

        If inQuote Then
            If ch = "\" Then
                If i + 1 > n Then Exit Function     ' Escape am Ende abgeschnitten

                i = i + 1
                esc = Mid$(json, i, 1)

                Select Case esc
                    Case """"
                        buf = buf & """"
                    Case "\"
                        buf = buf & "\"
                    Case "/"
                        buf = buf & "/"
                    Case "b"
                        buf = buf & Chr$(8)
                    Case "f"
                        buf = buf & Chr$(12)
                    Case "n"
                        buf = buf & vbLf
                    Case "r"
                        buf = buf & vbCr
                    Case "t"
                        buf = buf & vbTab
                    Case "u"
                        hex4 = Mid$(json, i + 1, 4)
                        If Not IsHex4Digits(hex4) Then Exit Function
                        buf = buf & JsonUnicodeChar(hex4)
                        i = i + 4
                    Case Else
                        ' Unbekannter Escape (z.B. \x) -> defektes JSON
                        Exit Function
                End Select
            ElseIf ch = """" Then
                inQuote = False
                If Not inRow Then
                    lastString = buf
                    buf = ""
                End If
            Else
                buf = buf & ch
            End If
        Else
            ' Nach dem Root-Objekt darf nur noch Whitespace kommen.
            If rootClosed Then
                If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Function
            End If

            ' Wenn "values" als Top-Level-Key da ist, MUSS ein Array folgen.
            ' {"values":null} / {"values":"x"} / {"values":{}} sind defekt,
            ' nicht "leerer Sheet".
            If expectValues Then
                If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf And ch <> "[" Then
                    Exit Function
                End If
            End If

            Select Case ch
                Case """"
                    inQuote = True
                    If inRow Then
                        cellStarted = True
                    Else
                        buf = ""
                    End If
                Case "{"
                    stack = stack & "{"
                    expectValues = False
                Case "["
                    If Len(stack) = 0 Then Exit Function    ' Root muss ein Objekt sein
                    stack = stack & "["

                    If expectValues And Not outFoundValues Then
                        valuesRoot = Len(stack)
                        outFoundValues = True
                        inValues = True
                    ElseIf inValues And Len(stack) = valuesRoot + 1 Then
                        Set cellsColl = New Collection
                        inRow = True
                        buf = ""
                        cellStarted = False
                    End If

                    expectValues = False
                Case "}"
                    If Right$(stack, 1) <> "{" Then Exit Function
                    stack = Left$(stack, Len(stack) - 1)
                    If Len(stack) = 0 Then rootClosed = True
                    lastString = ""
                Case "]"
                    If Right$(stack, 1) <> "[" Then Exit Function

                    If inValues And inRow And Len(stack) = valuesRoot + 1 Then
                        If cellStarted Then cellsColl.Add buf
                        rowsColl.Add cellsColl
                        Set cellsColl = Nothing
                        inRow = False
                        buf = ""
                        cellStarted = False
                    End If

                    stack = Left$(stack, Len(stack) - 1)

                    If inValues And Len(stack) = valuesRoot - 1 Then
                        inValues = False
                        valuesClosed = True
                    End If

                    lastString = ""
                Case ":"
                    If Not inRow Then
                        If lastString = "values" And stack = "{" Then expectValues = True
                        lastString = ""
                    End If
                Case ","
                    If inRow Then
                        cellsColl.Add buf
                        buf = ""
                        cellStarted = False
                    Else
                        lastString = ""
                        expectValues = False
                    End If
                Case Else
                    If inRow Then
                        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then
                            buf = buf & ch
                            cellStarted = True
                        End If
                    End If
            End Select
        End If

        i = i + 1
    Loop

    ' Abgeschnittenes JSON: offener String, offene Zeile, offene Klammer
    ' oder nicht geschlossenes values-Array -> KEINE Teildaten zurueckgeben.
    If inQuote Then Exit Function
    If Len(stack) <> 0 Then Exit Function
    If Not rootClosed Then Exit Function
    If inRow Then Exit Function
    If Not cellsColl Is Nothing Then Exit Function
    If outFoundValues And Not valuesClosed Then Exit Function

    Set outRows = rowsColl
    TryScanJsonDocument = True
End Function

Private Function FirstNonWhitespaceChar(ByRef s As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then
            FirstNonWhitespaceChar = ch
            Exit Function
        End If
    Next i
End Function

Private Function IsHex4Digits(ByVal s As String) As Boolean
    Dim i As Long
    Dim c As String

    If Len(s) <> 4 Then Exit Function

    For i = 1 To 4
        c = UCase$(Mid$(s, i, 1))
        If Not ((c >= "0" And c <= "9") Or (c >= "A" And c <= "F")) Then Exit Function
    Next i

    IsHex4Digits = True
End Function

Private Function JsonUnicodeChar(ByVal hex4 As String) As String
    ' \uXXXX -> Zeichen. Aufbau zur Laufzeit (ChrW), Quelltext bleibt ASCII.
    ' Werte > &H7FFF muessen fuer ChrW negativ uebergeben werden.
    Dim code As Long

    code = CLng("&H" & hex4)
    If code > 32767 Then code = code - 65536

    JsonUnicodeChar = ChrW$(code)
End Function

