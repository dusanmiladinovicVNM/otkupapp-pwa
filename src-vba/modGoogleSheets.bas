Attribute VB_Name = "modGoogleSheets"
Option Explicit

' ============================================================
' modGoogleSheets – Google Sheets API v4 Wrapper
'
' Liest und schreibt Google Sheets via REST API.
' Auth via modGoogleAuth.GetAccessToken()
'
' Hauptfunktionen:
'   WriteSheetData   — schreibt 2D-Array in ein Sheet-Tab
'   ReadSheetData    — liest Sheet-Tab als 2D-Array
'   ClearSheet       — löscht alle Daten in einem Tab
'   CreateSpreadsheet — erstellt neues Google Sheet
'   GetSpreadsheetID — sucht Sheet-ID nach Name in einem Folder
' ============================================================

Private Const SHEETS_API_BASE As String = "https://sheets.googleapis.com/v4/spreadsheets"
Private Const DRIVE_API_BASE As String = "https://www.googleapis.com/drive/v3"

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

' ============================================================
' PUBLIC — Write
' ============================================================

Public Function WriteSheetData(ByVal spreadsheetID As String, _
                               ByVal tabName As String, _
                               ByVal data As Variant) As Boolean
    ' Schreibt ein 2D-Array (1-based) in ein Google Sheet Tab
    ' Überschreibt vorhandene Daten ab A1
    
    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object
    
    On Error GoTo EH
    
    If Len(Trim$(spreadsheetID)) = 0 Then
        LogError "WriteSheetData", "spreadsheetID je prazan."
        WriteSheetData = False
        Exit Function
    End If

    If Len(Trim$(tabName)) = 0 Then
        LogError "WriteSheetData", "tabName je prazan."
        WriteSheetData = False
        Exit Function
    End If

    If IsEmpty(data) Then
        LogError "WriteSheetData", "data je Empty."
        WriteSheetData = False
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "WriteSheetData", "Kein Access Token"
        WriteSheetData = False
        Exit Function
    End If
    
    ' Erst Sheet leeren
    If Not ClearSheet(spreadsheetID, tabName) Then
        LogError "WriteSheetData", _
             "ClearSheet failed before write. SpreadsheetID=" & spreadsheetID & _
             ", Tab=" & tabName
        WriteSheetData = False
        Exit Function
    End If

    ' Daten als JSON-Body bauen
    body = BuildValuesJson(data)
    
    url = SHEETS_API_BASE & "/" & spreadsheetID & _
          "/values/" & UrlEncode(tabName) & "!A1" & _
          "?valueInputOption=RAW"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000
    
    http.Open "PUT", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body
    
    If http.status >= 200 And http.status < 300 Then
        LogInfo "WriteSheetData", tabName & ": " & UBound(data, 1) & " rows written"
        WriteSheetData = True
    Else
        LogError "WriteSheetData", "HTTP " & http.status & ": " & http.responseText, http.status
        WriteSheetData = False
    End If
    
    Exit Function

EH:
    LogErr "WriteSheetData"
    WriteSheetData = False
End Function

' ============================================================
' PUBLIC — Read
' ============================================================

Public Function ReadSheetData(ByVal spreadsheetID As String, _
                              ByVal tabName As String) As Variant
    Dim accessToken As String
    Dim url As String
    Dim http As Object

    On Error GoTo EH

    If Not RequireGoogleTextArg(spreadsheetID, "spreadsheetID", "ReadSheetData") Then
        ReadSheetData = Empty
        Exit Function
    End If

    If Not RequireGoogleTextArg(tabName, "tabName", "ReadSheetData") Then
        ReadSheetData = Empty
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "ReadSheetData", "Kein Access Token"
        ReadSheetData = Empty
        Exit Function
    End If

    url = SHEETS_API_BASE & "/" & spreadsheetID & _
          "/values/" & UrlEncode(tabName)

    Set http = CreateGoogleHttpRequest("ReadSheetData")

    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.Send

    If http.status <> 200 Then
        LogError "ReadSheetData", _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
        ReadSheetData = Empty
        Exit Function
    End If

    ReadSheetData = ParseValuesJson(http.responseText)
    Exit Function

EH:
    LogErr "ReadSheetData"
    ReadSheetData = Empty
End Function

' ============================================================
' PUBLIC — Clear
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
    http.Send "{}"

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
' PUBLIC — Create Spreadsheet
' ============================================================

Public Function CreateSpreadsheet(ByVal title As String, _
                                  Optional ByVal folderID As String = "") As String
    ' Erstellt ein neues Google Sheet, gibt SpreadsheetID zurück
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
' PUBLIC — Find Spreadsheet by Name in Folder
' ============================================================

Public Function GetSpreadsheetID(ByVal title As String, _
                                 Optional ByVal folderID As String = "") As String
    Dim accessToken As String
    Dim url As String
    Dim http As Object
    Dim query As String
    Dim responseText As String
    Dim foundID As String

    On Error GoTo EH

    If Not RequireGoogleTextArg(title, "title", "GetSpreadsheetID") Then
        GetSpreadsheetID = ""
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "GetSpreadsheetID", "Kein Access Token"
        GetSpreadsheetID = ""
        Exit Function
    End If

    query = "name='" & EscapeDriveQueryValue(title) & _
            "' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"

    If Len(Trim$(folderID)) > 0 Then
        query = query & " and '" & EscapeDriveQueryValue(folderID) & "' in parents"
    End If

    url = DRIVE_API_BASE & "/files?q=" & UrlEncode(query) & _
          "&fields=files(id,name)&pageSize=10"

    Set http = CreateGoogleHttpRequest("GetSpreadsheetID")

    http.Open "GET", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.Send

    responseText = CStr(http.responseText)

    If http.status <> 200 Then
        LogError "GetSpreadsheetID", _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(responseText), _
                 http.status
        GetSpreadsheetID = ""
        Exit Function
    End If

    foundID = ExtractSpreadsheetIDByExactName(responseText, title)

    If Len(Trim$(foundID)) = 0 Then
        LogInfo "GetSpreadsheetID", "Spreadsheet not found by exact name: " & title
        GetSpreadsheetID = ""
        Exit Function
    End If

    GetSpreadsheetID = foundID
    Exit Function

EH:
    LogErr "GetSpreadsheetID"
    GetSpreadsheetID = ""
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
' PUBLIC — Add Tab to existing Spreadsheet
' ============================================================

Public Function AddSheetTab(ByVal spreadsheetID As String, _
                            ByVal tabName As String) As Boolean
    Dim accessToken As String
    Dim url As String
    Dim body As String
    Dim http As Object

    On Error GoTo EH

    If Not RequireGoogleTextArg(spreadsheetID, "spreadsheetID", "AddSheetTab") Then
        AddSheetTab = False
        Exit Function
    End If

    If Not RequireGoogleTextArg(tabName, "tabName", "AddSheetTab") Then
        AddSheetTab = False
        Exit Function
    End If

    accessToken = GetAccessToken()
    If Len(accessToken) = 0 Then
        LogError "AddSheetTab", "Kein Access Token"
        AddSheetTab = False
        Exit Function
    End If

    url = SHEETS_API_BASE & "/" & spreadsheetID & ":batchUpdate"
    body = "{""requests"":[{""addSheet"":{""properties"":{""title"":""" & JsonEscape(tabName) & """}}}]}"

    Set http = CreateGoogleHttpRequest("AddSheetTab")

    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send body

    If http.status >= 200 And http.status < 300 Then
        AddSheetTab = True
    Else
        LogError "AddSheetTab", _
                 "HTTP " & http.status & ": " & GoogleHttpBodyForLog(http.responseText), _
                 http.status
        AddSheetTab = False
    End If

    Exit Function

EH:
    LogErr "AddSheetTab"
    AddSheetTab = False
End Function

' ============================================================
' PRIVATE — Move file to folder (Drive API)
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
    ' Extrahiert erste Parent-ID aus "parents":["xxx"]
    Dim p As Long, startPos As Long, endPos As Long
    p = InStr(json, """parents""")
    If p = 0 Then Exit Function
    startPos = InStr(p, json, """") + 1
    startPos = InStr(startPos, json, "[") + 1
    startPos = InStr(startPos, json, """") + 1
    endPos = InStr(startPos, json, """")
    If endPos > startPos Then GetFirstParent = Mid$(json, startPos, endPos - startPos)
End Function

' ============================================================
' PRIVATE — JSON Builder für Sheets API
' ============================================================

Private Function BuildValuesJson(ByVal data As Variant) As String
    ' Baut JSON body für values:update API
    ' {"values":[["a","b"],["c","d"]]}
    ' ALLES als String schreiben — Google Sheets erkennt Zahlen/Daten automatisch
    ' Verhindert Oktal-Problem bei führenden Nullen (Telefonnummern, BPG etc.)
    
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
    Dim p As Long
    Dim valuesStart As Long
    Dim valuesEnd As Long
    Dim block As String
    Dim rowList() As String
    Dim rowCount As Long
    Dim colCount As Long
    Dim result() As Variant
    Dim i As Long, j As Long
    Dim cells() As String
    
    json = Replace(json, vbCrLf, "")
    json = Replace(json, vbLf, "")
    json = Replace(json, vbCr, "")
    
    ' Spaces zwischen Klammern entfernen
    Do While InStr(json, "[ ") > 0
        json = Replace(json, "[ ", "[")
    Loop
    Do While InStr(json, " ]") > 0
        json = Replace(json, " ]", "]")
    Loop
    Do While InStr(json, ", ") > 0
        json = Replace(json, ", ", ",")
    Loop
    
    p = InStr(json, """values""")
    If p = 0 Then
        ParseValuesJson = Empty
        Exit Function
    End If
    
    valuesStart = InStr(p, json, "[[")
    If valuesStart = 0 Then
        ParseValuesJson = Empty
        Exit Function
    End If
    
    valuesEnd = InStrRev(json, "]]")
    If valuesEnd = 0 Or valuesEnd <= valuesStart Then
        ParseValuesJson = Empty
        Exit Function
    End If
    
    block = Mid$(json, valuesStart + 1, valuesEnd - valuesStart)
    
    rowList = Split(block, "],[")
    rowCount = UBound(rowList) + 1
    
    rowList(0) = Mid$(rowList(0), 2)
    rowList(UBound(rowList)) = Left$(rowList(UBound(rowList)), Len(rowList(UBound(rowList))) - 1)
    
    cells = SplitCsvJson(rowList(0))
    colCount = UBound(cells) + 1
    
    ReDim result(1 To rowCount, 1 To colCount)
    
    For i = 0 To rowCount - 1
        cells = SplitCsvJson(rowList(i))
        For j = 0 To UBound(cells)
            If j < colCount Then
                result(i + 1, j + 1) = CleanJsonValue(cells(j))
            End If
        Next j
    Next i
    
    ParseValuesJson = result
End Function
Private Function SplitCsvJson(ByVal s As String) As String()
    ' Split auf Komma, aber nicht innerhalb von Anführungszeichen
    Dim result() As String
    Dim count As Long, i As Long
    Dim inQuote As Boolean
    Dim current As String
    
    ReDim result(0 To 0)
    
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        
        If ch = """" Then
            inQuote = Not inQuote
        ElseIf ch = "," And Not inQuote Then
            result(count) = current
            count = count + 1
            ReDim Preserve result(0 To count)
            current = ""
        Else
            current = current & ch
        End If
    Next i
    
    result(count) = current
    SplitCsvJson = result
End Function

Private Function CleanJsonValue(ByVal s As String) As String
    s = Trim$(s)
    If Left$(s, 1) = """" And Right$(s, 1) = """" Then
        s = Mid$(s, 2, Len(s) - 2)
    End If
    s = Replace(s, "\""", """")
    s = Replace(s, "\\", "\")
    s = Replace(s, "\n", vbLf)
    CleanJsonValue = s
End Function

