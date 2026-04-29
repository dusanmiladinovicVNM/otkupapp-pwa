Attribute VB_Name = "modGoogleAuth"
Option Explicit

' ============================================================
' modGoogleAuth – OAuth2 für Google APIs aus VBA
'
' Flow:
' 1. Erster Start: GetAuthorizationCode ? Browser öffnet Google Login
' 2. User kopiert Authorization Code ? ExchangeCodeForTokens
' 3. Access Token + Refresh Token werden in tblSEFConfig gespeichert
' 4. Bei jedem API-Call: GetAccessToken prüft Ablauf, refresht wenn nötig
'
' Config-Keys in central ConfigKey/ConfigValue table: tblSEFConfig
'   GOOGLE_CLIENT_ID
'   GOOGLE_CLIENT_SECRET
'   GOOGLE_ACCESS_TOKEN
'   GOOGLE_REFRESH_TOKEN
'   GOOGLE_TOKEN_EXPIRES_AT     (ISO Timestamp)
'
' Setup:
' 1. Google Cloud Console ? OAuth 2.0 Client ID (Desktop App) erstellen
' 2. client_id und client_secret in tblSEFConfig eintragen
' 3. Einmalig: RunGoogleAuthSetup ausführen
' ============================================================

Private Const GOOGLE_AUTH_URL As String = "https://accounts.google.com/o/oauth2/v2/auth"
Private Const GOOGLE_TOKEN_URL As String = "https://oauth2.googleapis.com/token"
Private Const GOOGLE_REDIRECT_URI As String = "urn:ietf:wg:oauth:2.0:oob"
Private Const GOOGLE_SCOPE As String = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive"

' ============================================================
' PUBLIC — Setup (einmalig)
' ============================================================

Public Sub RunGoogleAuthSetup()
    ' Schritt 1: Browser öffnet Google Login
    ' Schritt 2: User kopiert Code
    ' Schritt 3: Code ? Tokens
    
    Dim clientID As String
    Dim clientSecret As String
    
    clientID = GetConfigValue("GOOGLE_CLIENT_ID")
    clientSecret = GetConfigValue("GOOGLE_CLIENT_SECRET")
    
    If Len(Trim$(clientID)) = 0 Or Len(Trim$(clientSecret)) = 0 Then
        MsgBox "GOOGLE_CLIENT_ID i GOOGLE_CLIENT_SECRET moraju biti upisani u tblSEFConfig.", _
                vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' Browser öffnen
    Dim authUrl As String
    authUrl = GOOGLE_AUTH_URL & _
              "?client_id=" & UrlEncodeGoogle(clientID) & _
              "&redirect_uri=" & UrlEncodeGoogle(GOOGLE_REDIRECT_URI) & _
              "&response_type=code" & _
              "&scope=" & UrlEncodeGoogle(GOOGLE_SCOPE) & _
              "&access_type=offline" & _
              "&prompt=consent"
    
    Shell "cmd /c start """" """ & authUrl & """", vbNormalFocus
    
    ' Code abfragen
    Dim authCode As String
    authCode = InputBox("Google Login oeffnen sich im Browser." & vbCrLf & vbCrLf & _
                        "1. Melde dich an und erlaube den Zugriff" & vbCrLf & _
                        "2. Kopiere den Authorization Code" & vbCrLf & _
                        "3. Fuege ihn hier ein:", _
                        "Google OAuth2 Setup")
    
    If Len(Trim$(authCode)) = 0 Then
        MsgBox "Setup abgebrochen.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Code ? Tokens
    If ExchangeCodeForTokens(authCode) Then
        MsgBox "Google OAuth2 erfolgreich eingerichtet!", vbInformation, APP_NAME
    Else
        MsgBox "Token-Austausch fehlgeschlagen. Prüfe Client ID/Secret.", vbCritical, APP_NAME
    End If
End Sub

' ============================================================
' PUBLIC — Access Token holen (für jeden API-Call)
' ============================================================

Public Function GetAccessToken() As String
    ' Gibt gültigen Access Token zurück
    ' Refresht automatisch wenn abgelaufen
    ' Returns "" bei Fehler
    
    Dim accessToken As String
    Dim expiresAt As String
    
    accessToken = GetConfigValue("GOOGLE_ACCESS_TOKEN")
    expiresAt = GetConfigValue("GOOGLE_TOKEN_EXPIRES_AT")
    
    If Len(Trim$(accessToken)) = 0 Then
        LogWarn "modGoogleAuth", "Kein Access Token vorhanden. RunGoogleAuthSetup ausfuehren."
        GetAccessToken = ""
        Exit Function
    End If
    
    ' Prüfe ob abgelaufen (mit 60s Puffer)
    If IsTokenExpired(expiresAt) Then
        If RefreshAccessToken() Then
            GetAccessToken = GetConfigValue("GOOGLE_ACCESS_TOKEN")
        Else
            GetAccessToken = ""
        End If
    Else
        GetAccessToken = accessToken
    End If
End Function

Public Function IsGoogleAuthConfigured() As Boolean
    IsGoogleAuthConfigured = Len(Trim$(GetConfigValue("GOOGLE_REFRESH_TOKEN"))) > 0
End Function

' ============================================================
' PRIVATE — Token Exchange
' ============================================================

Private Function ExchangeCodeForTokens(ByVal authCode As String) As Boolean
    Dim http As Object
    Dim body As String
    Dim responseText As String
    
    On Error GoTo EH
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000
    
    body = "code=" & UrlEncodeGoogle(Trim$(authCode)) & _
           "&client_id=" & UrlEncodeGoogle(GetConfigValue("GOOGLE_CLIENT_ID")) & _
           "&client_secret=" & UrlEncodeGoogle(GetConfigValue("GOOGLE_CLIENT_SECRET")) & _
           "&redirect_uri=" & UrlEncodeGoogle(GOOGLE_REDIRECT_URI) & _
           "&grant_type=authorization_code"
    
    http.Open "POST", GOOGLE_TOKEN_URL, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send body
    
    responseText = http.responseText
    
    If http.status <> 200 Then
        LogError "ExchangeCodeForTokens", _
             "HTTP " & http.status & ": " & RedactGoogleSecrets(responseText), _
             http.status
        ExchangeCodeForTokens = False
        Exit Function
    End If
    
    ' Token parsen und speichern
    Dim accessToken As String
    Dim refreshToken As String
    Dim expiresIn As Long
    
    accessToken = ExtractJsonStringGoogle(responseText, "access_token")
    refreshToken = ExtractJsonStringGoogle(responseText, "refresh_token")
    expiresIn = ParseGoogleExpiresIn(responseText, "ExchangeCodeForTokens")
    
    If Len(accessToken) = 0 Then
        LogError "ExchangeCodeForTokens", "Kein access_token in Response"
        ExchangeCodeForTokens = False
        Exit Function
    End If
    
    ' Speichern
    Call SetConfigValue("GOOGLE_ACCESS_TOKEN", accessToken)
    If Len(refreshToken) > 0 Then
        Call SetConfigValue("GOOGLE_REFRESH_TOKEN", refreshToken)
    End If
    Call SetConfigValue("GOOGLE_TOKEN_EXPIRES_AT", CalculateExpiryTimestamp(expiresIn))
    
    LogInfo "modGoogleAuth", "Tokens erfolgreich gespeichert"
    ExchangeCodeForTokens = True
    Exit Function

EH:
    LogErr "ExchangeCodeForTokens"
    ExchangeCodeForTokens = False
End Function

Private Function RefreshAccessToken() As Boolean
    Dim http As Object
    Dim body As String
    Dim responseText As String
    Dim refreshToken As String
    
    On Error GoTo EH
    
    refreshToken = GetConfigValue("GOOGLE_REFRESH_TOKEN")
    If Len(Trim$(refreshToken)) = 0 Then
        LogError "RefreshAccessToken", "Kein Refresh Token vorhanden"
        RefreshAccessToken = False
        Exit Function
    End If
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 10000, 10000, 30000, 30000
    
    body = "refresh_token=" & UrlEncodeGoogle(refreshToken) & _
           "&client_id=" & UrlEncodeGoogle(GetConfigValue("GOOGLE_CLIENT_ID")) & _
           "&client_secret=" & UrlEncodeGoogle(GetConfigValue("GOOGLE_CLIENT_SECRET")) & _
           "&grant_type=refresh_token"
    
    http.Open "POST", GOOGLE_TOKEN_URL, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send body
    
    responseText = http.responseText
    
    If http.status <> 200 Then
        LogError "RefreshAccessToken", _
             "HTTP " & http.status & ": " & RedactGoogleSecrets(responseText), _
             http.status
        RefreshAccessToken = False
        Exit Function
    End If
    
    Dim accessToken As String
    Dim expiresIn As Long
    
    accessToken = ExtractJsonStringGoogle(responseText, "access_token")
    expiresIn = ParseGoogleExpiresIn(responseText, "RefreshAccessToken")
    
    If Len(accessToken) = 0 Then
        LogError "RefreshAccessToken", "Kein access_token in Refresh-Response"
        RefreshAccessToken = False
        Exit Function
    End If
    
    Call SetConfigValue("GOOGLE_ACCESS_TOKEN", accessToken)
    Call SetConfigValue("GOOGLE_TOKEN_EXPIRES_AT", CalculateExpiryTimestamp(expiresIn))
    
    ' Neuer Refresh Token falls mitgeliefert (selten)
    Dim newRefresh As String
    newRefresh = ExtractJsonStringGoogle(responseText, "refresh_token")
    If Len(newRefresh) > 0 Then
        Call SetConfigValue("GOOGLE_REFRESH_TOKEN", newRefresh)
    End If
    
    LogInfo "modGoogleAuth", "Access Token refreshed"
    RefreshAccessToken = True
    Exit Function

EH:
    LogErr "RefreshAccessToken"
    RefreshAccessToken = False
End Function

' ============================================================
' PRIVATE — Helpers
' ============================================================

Private Function RedactGoogleSecrets(ByVal textValue As String) As String
    Dim result As String
    Dim secretValue As String

    result = CStr(textValue)

    secretValue = Trim$(GetConfigValue("GOOGLE_CLIENT_SECRET"))
    If Len(secretValue) > 0 Then result = Replace(result, secretValue, "***GOOGLE_CLIENT_SECRET***")

    secretValue = Trim$(GetConfigValue("GOOGLE_ACCESS_TOKEN"))
    If Len(secretValue) > 0 Then result = Replace(result, secretValue, "***GOOGLE_ACCESS_TOKEN***")

    secretValue = Trim$(GetConfigValue("GOOGLE_REFRESH_TOKEN"))
    If Len(secretValue) > 0 Then result = Replace(result, secretValue, "***GOOGLE_REFRESH_TOKEN***")

    RedactGoogleSecrets = Left$(result, 1000)
End Function

Private Function ParseGoogleExpiresIn(ByVal responseText As String, ByVal sourceName As String) As Long
    Dim expiresIn As Long

    expiresIn = CLng(val(ExtractJsonStringGoogle(responseText, "expires_in")))

    If expiresIn <= 0 Then
        LogWarn sourceName, "expires_in missing/invalid in Google token response. Using 3600 seconds fallback."
        expiresIn = 3600
    End If

    ParseGoogleExpiresIn = expiresIn
End Function

Private Function IsTokenExpired(ByVal expiresAt As String) As Boolean
    If Len(Trim$(expiresAt)) = 0 Then
        IsTokenExpired = True
        Exit Function
    End If
    
    On Error GoTo Expired
    
    ' Format: "2026-03-18T14:35:00"
    Dim expiryDate As Date
    expiryDate = CDate(Replace(expiresAt, "T", " "))
    
    ' 60 Sekunden Puffer
    IsTokenExpired = (Now >= DateAdd("s", -60, expiryDate))
    Exit Function

Expired:
    IsTokenExpired = True
End Function

Private Function CalculateExpiryTimestamp(ByVal expiresInSeconds As Long) As String
    CalculateExpiryTimestamp = Format$(DateAdd("s", expiresInSeconds, Now), "yyyy-mm-dd\Thh:nn:ss")
End Function

Public Function ExtractJsonStringGoogle(ByVal json As String, ByVal key As String) As String
    ' Einfacher JSON-String-Extraktor (kein vollständiger Parser)
    Dim pattern As String
    Dim p As Long, startPos As Long, endPos As Long
    
    ' Suche "key" : "value" oder "key": value (für Zahlen)
    pattern = """" & key & """"
    p = InStr(1, json, pattern, vbTextCompare)
    If p = 0 Then Exit Function
    
    startPos = p + Len(pattern)
    
    ' Überspringe : und Whitespace
    Do While startPos <= Len(json)
        Select Case Mid$(json, startPos, 1)
            Case ":", " ", vbTab: startPos = startPos + 1
            Case Else: Exit Do
        End Select
    Loop
    
    ' String-Wert (in Anführungszeichen)
    If Mid$(json, startPos, 1) = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, json, """")
        If endPos > startPos Then
            ExtractJsonStringGoogle = Mid$(json, startPos, endPos - startPos)
        End If
    Else
        ' Numerischer Wert (ohne Anführungszeichen)
        endPos = startPos
        Do While endPos <= Len(json)
            Select Case Mid$(json, endPos, 1)
                Case "0" To "9", ".": endPos = endPos + 1
                Case Else: Exit Do
            End Select
        Loop
        If endPos > startPos Then
            ExtractJsonStringGoogle = Mid$(json, startPos, endPos - startPos)
        End If
    End If
End Function

Public Function UrlEncodeGoogle(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, result As String
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = Asc(ch)
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & ch
            Case Else
                result = result & "%" & Right$("0" & Hex$(code), 2)
        End Select
    Next i
    
    UrlEncodeGoogle = result
End Function


