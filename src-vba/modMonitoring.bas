' ============================================================
' OtkupApp / AgriX - modMonitoring.bas
' ============================================================
'
' Svrha:
' - VBA šalje monitoring evente u GAS endpoint preko action="monitorPublic".
' - Monitoring je best-effort: nikada ne sme da obori business operaciju.
' - Ne logovati tokene, SEF API key, pun XML/PDF, pun JMBG ili bankovne podatke.
'
' Potrebno u Excel fajlu:
' 1) Named range ili Config sheet vrednost:
'    MONITORING_ENDPOINT = GAS Web App URL
'    MONITORING_SECRET   = isti string kao GAS Script Property MONITORING_INGEST_SECRET
'    APP_VERSION         = npr. VBA-1.0.0
' 2) GAS mora imati public monitoring hook:
'    action = monitorPublic
'
' Preporuka:
' - Secret drži u hidden/very hidden Config sheet-u ili Named Range-u.
' - Ne stavljaj secret u PWA frontend. Ovo je samo za VBA/desktop.
'
' ============================================================

Option Explicit

Private Const MONITORING_SOURCE As String = "VBA"
Private Const DEFAULT_MONITORING_ENV As String = "DEV"
Private Const DEFAULT_APP_VERSION As String = APP_VERSION
Private Const HTTP_TIMEOUT_MS As Long = 1200
Private Const HTTP_DEBUG_TIMEOUT_MS As Long = 10000

' Modulska referenca: drži in-flight async zahteve da WinHttp ne otkaže slanje pre vremena
Private m_inFlight As Collection

' ============================================================
' PUBLIC API
' ============================================================

Public Sub Monitor_AppOpen(Optional ByVal userId As String = "Operator")
    On Error Resume Next

    Monitor_Event _
        eventType:="VBA_APP_OPEN", _
        severity:="INFO", _
        message:="Excel/VBA client opened", _
        userId:=userId, _
        moduleName:="ThisWorkbook", _
        procedureName:="Workbook_Open", _
        payloadJson:="{""workbook"":" & JsonString(DataBook.name) & "}"
End Sub

Public Sub Monitor_Event( _
    ByVal eventType As String, _
    Optional ByVal severity As String = "INFO", _
    Optional ByVal message As String = "", _
    Optional ByVal userId As String = "Operator", _
    Optional ByVal role As String = "Admin", _
    Optional ByVal moduleName As String = "", _
    Optional ByVal procedureName As String = "", _
    Optional ByVal entityType As String = "", _
    Optional ByVal entityID As String = "", _
    Optional ByVal correlationId As String = "", _
    Optional ByVal payloadJson As String = "{}" _
)
    On Error Resume Next

    If Not Monitoring_IsEnabled() Then Exit Sub

    Dim json As String
    json = BuildBaseJson( _
        eventType:=eventType, _
        severity:=severity, _
        message:=message, _
        userId:=userId, _
        role:=role, _
        moduleName:=moduleName, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityID:=entityID, _
        correlationId:=correlationId, _
        payloadJson:=payloadJson _
    )

    Call Monitoring_PostJson(Monitoring_Endpoint(), json)
End Sub

Public Sub Monitor_Error( _
    ByVal moduleName As String, _
    ByVal procedureName As String, _
    Optional ByVal entityType As String = "", _
    Optional ByVal entityID As String = "", _
    Optional ByVal correlationId As String = "", _
    Optional ByVal userId As String = "Operator", _
    Optional ByVal role As String = "Admin", _
    Optional ByVal errorNumber As Long = 0, _
    Optional ByVal errorDescription As String = "", _
    Optional ByVal errorSource As String = "" _
)
    On Error Resume Next

    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = errorNumber
    errDesc = errorDescription
    errSrc = errorSource

    If errNo = 0 Then errNo = Err.Number
    If Len(errDesc) = 0 Then errDesc = Err.description
    If Len(errSrc) = 0 Then errSrc = Err.SOURCE

    Dim payload As String
    payload = "{" & _
        """errorNumber"":" & CStr(errNo) & "," & _
        """errorDescription"":" & JsonString(Monitoring_SanitizeText(errDesc)) & "," & _
        """errorSource"":" & JsonString(Monitoring_SanitizeText(errSrc)) & "," & _
        """workbook"":" & JsonString(DataBook.name) & _
    "}"

    Monitor_Event _
        eventType:="VBA_ERROR", _
        severity:="ERROR", _
        message:=errDesc, _
        userId:=userId, _
        role:=role, _
        moduleName:=moduleName, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityID:=entityID, _
        correlationId:=correlationId, _
        payloadJson:=payload
End Sub

Public Sub Monitor_Critical( _
    ByVal moduleName As String, _
    ByVal procedureName As String, _
    ByVal message As String, _
    Optional ByVal entityType As String = "", _
    Optional ByVal entityID As String = "", _
    Optional ByVal correlationId As String = "", _
    Optional ByVal payloadJson As String = "{}" _
)
    On Error Resume Next

    Monitor_Event _
        eventType:="VBA_CRITICAL", _
        severity:="CRITICAL", _
        message:=message, _
        moduleName:=moduleName, _
        procedureName:=procedureName, _
        entityType:=entityType, _
        entityID:=entityID, _
        correlationId:=correlationId, _
        payloadJson:=payloadJson
End Sub

Public Sub Monitor_SEF( _
    ByVal eventType As String, _
    ByVal severity As String, _
    ByVal invoiceLocalId As String, _
    ByVal businessInvoiceNo As String, _
    ByVal sefStatus As String, _
    ByVal localStatus As String, _
    Optional ByVal sefRequestId As String = "", _
    Optional ByVal sefInvoiceId As String = "", _
    Optional ByVal attemptCount As Long = 0, _
    Optional ByVal lastHttpCode As String = "", _
    Optional ByVal lastError As String = "", _
    Optional ByVal nextAction As String = "WAIT", _
    Optional ByVal needsManualReview As Boolean = False _
)
    On Error Resume Next

    Dim corr As String
    corr = invoiceLocalId
    If Len(corr) = 0 Then corr = businessInvoiceNo

    Dim payload As String
    payload = "{" & _
        """invoiceLocalId"":" & JsonString(invoiceLocalId) & "," & _
        """businessInvoiceNo"":" & JsonString(businessInvoiceNo) & "," & _
        """sefRequestId"":" & JsonString(sefRequestId) & "," & _
        """sefInvoiceId"":" & JsonString(sefInvoiceId) & "," & _
        """sefStatus"":" & JsonString(sefStatus) & "," & _
        """localStatus"":" & JsonString(localStatus) & "," & _
        """lastAttemptAt"":" & JsonString(IsoNow()) & "," & _
        """attemptCount"":" & CStr(attemptCount) & "," & _
        """lastHttpCode"":" & JsonString(lastHttpCode) & "," & _
        """lastError"":" & JsonString(Monitoring_SanitizeText(lastError)) & "," & _
        """nextAction"":" & JsonString(nextAction) & "," & _
        """needsManualReview"":" & LCase$(CStr(needsManualReview)) & _
    "}"

    Monitor_Event _
        eventType:=eventType, _
        severity:=severity, _
        message:=lastError, _
        moduleName:="modSEF", _
        procedureName:=eventType, _
        entityType:="Faktura", _
        entityID:=invoiceLocalId, _
        correlationId:=corr, _
        payloadJson:=payload
End Sub

Public Sub Monitor_Backup( _
    ByVal backupType As String, _
    ByVal status As String, _
    Optional ByVal backupFileId As String = "", _
    Optional ByVal backupLocation As String = "", _
    Optional ByVal rowsCount As Long = 0, _
    Optional ByVal checksum As String = "", _
    Optional ByVal durationMs As Long = 0, _
    Optional ByVal errorMessage As String = "" _
)
    On Error Resume Next

    Dim sev As String
    Dim eventType As String

    If UCase$(status) = "SUCCESS" Then
        sev = "INFO"
        eventType = "BACKUP_SUCCESS"
    Else
        sev = "CRITICAL"
        eventType = "BACKUP_FAIL"
    End If

    Dim payload As String
    payload = "{" & _
        """backupType"":" & JsonString(backupType) & "," & _
        """sourceSpreadsheetId"":" & JsonString(DataBook.name) & "," & _
        """backupFileId"":" & JsonString(Monitoring_SanitizeText(backupFileId)) & "," & _
        """backupLocation"":" & JsonString(Monitoring_SanitizeText(backupLocation)) & "," & _
        """rowsCount"":" & CStr(rowsCount) & "," & _
        """checksum"":" & JsonString(checksum) & "," & _
        """status"":" & JsonString(status) & "," & _
        """durationMs"":" & CStr(durationMs) & "," & _
        """errorMessage"":" & JsonString(Monitoring_SanitizeText(errorMessage)) & _
    "}"

    Monitor_Event _
        eventType:=eventType, _
        severity:=sev, _
        message:=errorMessage, _
        moduleName:="Backup", _
        procedureName:="Monitor_Backup", _
        entityType:="Backup", _
        entityID:=backupType, _
        correlationId:="BACKUP", _
        payloadJson:=payload
End Sub

Public Function Monitor_Test() As Boolean
    On Error GoTo EH

    Dim json As String
    json = BuildBaseJson( _
        eventType:="VBA_MONITORING_TEST", _
        severity:="INFO", _
        message:="Manual VBA monitoring test", _
        userId:="Operator", _
        role:="Admin", _
        moduleName:="modMonitoring", _
        procedureName:="Monitor_Test", _
        entityType:="", _
        entityID:="", _
        correlationId:="VBA-MONITORING-TEST", _
        payloadJson:="{""test"":true}" _
    )

    Monitor_Test = Monitoring_PostJsonDebug(Monitoring_Endpoint(), json)
    Exit Function

EH:
    Debug.Print "Monitor_Test VBA error: " & Err.Number & " - " & Err.description
    Monitor_Test = False
End Function

Public Function Monitoring_DiagnoseConfig() As Boolean
    On Error Resume Next

    Dim endpoint As String
    Dim secret As String

    endpoint = Monitoring_Endpoint()
    secret = Monitoring_Secret()

    Debug.Print "--- Monitoring config diagnose ---"
    Debug.Print "Endpoint length: " & Len(endpoint)
    Debug.Print "Endpoint ends with /exec: " & (Right$(Trim$(endpoint), 5) = "/exec")
    Debug.Print "Secret length: " & Len(secret)
    Debug.Print "Env: " & Monitoring_Env()
    Debug.Print "AppVersion: " & AppVersion()
    Debug.Print "DeviceId: " & DeviceId()

    Monitoring_DiagnoseConfig = (Len(endpoint) > 0 And Len(secret) > 0)
End Function

Public Function Monitoring_PostJsonDebug(ByVal url As String, ByVal jsonBody As String) As Boolean
    On Error GoTo EH

    Debug.Print "--- Monitoring HTTP diagnose ---"
    Debug.Print "URL: " & url
    Debug.Print "Body length: " & Len(jsonBody)
    Debug.Print "Body preview redacted; eventType present: " & (InStr(1, jsonBody, """eventType""", vbTextCompare) > 0)

    If Len(Trim$(url)) = 0 Then
        Debug.Print "ERROR: MONITORING_ENDPOINT is empty"
        Exit Function
    End If

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.SetTimeouts HTTP_DEBUG_TIMEOUT_MS, HTTP_DEBUG_TIMEOUT_MS, HTTP_DEBUG_TIMEOUT_MS, HTTP_DEBUG_TIMEOUT_MS
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    http.Send jsonBody

    Debug.Print "HTTP Status: " & http.status
    Debug.Print "HTTP Response: " & Left$(CStr(http.responseText), 1000)

    Monitoring_PostJsonDebug = (http.status >= 200 And http.status < 300 And InStr(1, http.responseText, """success"":true", vbTextCompare) > 0)
    Exit Function

EH:
    Debug.Print "HTTP VBA error: " & Err.Number & " - " & Err.description
    Monitoring_PostJsonDebug = False
End Function

' ============================================================
' PRIVATE HTTP
' ============================================================

Private Function Monitoring_PostJson(ByVal url As String, ByVal jsonBody As String) As Boolean
    On Error GoTo EH

    If Len(Trim$(url)) = 0 Then Exit Function

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, True                 ' True = ASINHRONO (ne blokira)
    http.SetTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "Accept", "application/json"
    http.Send jsonBody
    ' NE pozivamo WaitForResponse -> vraca se odmah; round-trip ide u pozadini

    ' Zadrži referencu da se zahtev ne otkaže kad lokalni objekat izade iz opsega.
    If m_inFlight Is Nothing Then Set m_inFlight = New Collection
    m_inFlight.Add http
    Do While m_inFlight.count > 12              ' ogranici; stariji su odavno poslati
        m_inFlight.Remove 1
    Loop

    Monitoring_PostJson = True
    Exit Function

EH:
    Monitoring_PostJson = False
End Function

Private Function BuildBaseJson( _
    ByVal eventType As String, _
    ByVal severity As String, _
    ByVal message As String, _
    ByVal userId As String, _
    ByVal role As String, _
    ByVal moduleName As String, _
    ByVal procedureName As String, _
    ByVal entityType As String, _
    ByVal entityID As String, _
    ByVal correlationId As String, _
    ByVal payloadJson As String _
) As String
    On Error Resume Next

    If Len(Trim$(payloadJson)) = 0 Then payloadJson = "{}"
    If Left$(Trim$(payloadJson), 1) <> "{" Then payloadJson = "{}"
    payloadJson = Monitoring_SanitizePayloadJson(payloadJson)

    BuildBaseJson = "{" & _
        """action"":""monitorPublic""," & _
        """monitoringSecret"":" & JsonString(Monitoring_Secret()) & "," & _
        """environment"":" & JsonString(Monitoring_Env()) & "," & _
        """source"":" & JsonString(MONITORING_SOURCE) & "," & _
        """severity"":" & JsonString(UCase$(severity)) & "," & _
        """eventType"":" & JsonString(eventType) & "," & _
        """userId"":" & JsonString(userId) & "," & _
        """role"":" & JsonString(role) & "," & _
        """deviceId"":" & JsonString(DeviceId()) & "," & _
        """appVersion"":" & JsonString(AppVersion()) & "," & _
        """module"":" & JsonString(moduleName) & "," & _
        """functionName"":" & JsonString(procedureName) & "," & _
        """entityType"":" & JsonString(entityType) & "," & _
        """entityId"":" & JsonString(entityID) & "," & _
        """correlationId"":" & JsonString(correlationId) & "," & _
        """message"":" & JsonString(Monitoring_SanitizeText(message)) & "," & _
        """payload"":" & payloadJson & _
    "}"
End Function

' ============================================================
' CONFIG
' ============================================================
Private Function SafeConfigValue(ByVal key As String) As String
    On Error Resume Next
    SafeConfigValue = ConfigValue(key)
End Function

Private Function Monitoring_IsEnabled() As Boolean
    On Error Resume Next
    Monitoring_IsEnabled = (Len(Monitoring_Endpoint()) > 0 And Len(Monitoring_Secret()) > 0)
End Function

Private Function Monitoring_Endpoint() As String
    Monitoring_Endpoint = SafeConfigValue("MONITORING_ENDPOINT")
End Function

Private Function Monitoring_Secret() As String
    Monitoring_Secret = SafeConfigValue("MONITORING_SECRET")
End Function

Private Function AppVersion() As String
    On Error GoTo Fallback

    AppVersion = APP_VERSION
    Exit Function

Fallback:
    AppVersion = DEFAULT_APP_VERSION
End Function

Private Function Monitoring_Env() As String
    Dim v As String
    v = SafeConfigValue("MONITORING_ENV")
    If Len(v) = 0 Then v = DEFAULT_MONITORING_ENV
    Monitoring_Env = v
End Function

Private Function ConfigValue(ByVal key As String) As String
    On Error Resume Next
    ' Config se uvek cita iz klijentskog DataBook-a (u COMBINED rezimu = ThisWorkbook).
    ConfigValue = ConfigValueFromWorkbook(DataBook, key)
End Function

Private Function ConfigValueFromWorkbook(ByVal wb As Workbook, ByVal key As String) As String
    On Error GoTo Done

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim r As Long
    Dim keyCol As Long
    Dim valueCol As Long
    Dim currentKey As String

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(Trim$(lo.name), "tblSEFConfig", vbTextCompare) = 0 Then

                keyCol = FindListColumnIndex(lo, "ConfigKey")
                valueCol = FindListColumnIndex(lo, "ConfigValue")

                If keyCol = 0 Or valueCol = 0 Then
                    Debug.Print "tblSEFConfig found but ConfigKey/ConfigValue headers not found"
                    Exit Function
                End If

                If lo.DataBodyRange Is Nothing Then
                    Debug.Print "tblSEFConfig has no rows"
                    Exit Function
                End If

                For r = 1 To lo.DataBodyRange.rows.count
                    currentKey = Trim$(CStr(lo.DataBodyRange.cells(r, keyCol).value))

                    If StrComp(currentKey, Trim$(key), vbTextCompare) = 0 Then
                        ConfigValueFromWorkbook = Trim$(CStr(lo.DataBodyRange.cells(r, valueCol).value))
                        Exit Function
                    End If
                Next r

            End If
        Next lo
    Next ws

Done:
End Function

Private Function FindListColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    On Error GoTo Done

    Dim c As ListColumn

    For Each c In lo.ListColumns
        If StrComp(Trim$(c.name), Trim$(headerName), vbTextCompare) = 0 Then
            FindListColumnIndex = c.index
            Exit Function
        End If
    Next c

Done:
End Function

Private Function DeviceId() As String
    On Error Resume Next

    Dim userName As String
    Dim computerName As String

    userName = Environ$("USERNAME")
    computerName = Environ$("COMPUTERNAME")

    DeviceId = computerName & "|" & userName
End Function

' ============================================================
' JSON / TIME HELPERS
' ============================================================

Private Function Monitoring_SanitizeText(ByVal s As String) As String
    On Error Resume Next

    Dim t As String
    t = CStr(s)

    Dim monitoringSecret As String
    Dim googleAccessToken As String
    Dim googleRefreshToken As String
    Dim sefApiKey As String

    monitoringSecret = SafeConfigValue("MONITORING_SECRET")
    googleAccessToken = SafeConfigValue("GOOGLE_ACCESS_TOKEN")
    googleRefreshToken = SafeConfigValue("GOOGLE_REFRESH_TOKEN")
    sefApiKey = SafeConfigValue("SEF_API_KEY")

    If Len(monitoringSecret) > 0 Then t = Replace(t, monitoringSecret, "[REDACTED_MONITORING_SECRET]")
    If Len(googleAccessToken) > 0 Then t = Replace(t, googleAccessToken, "[REDACTED_GOOGLE_ACCESS_TOKEN]")
    If Len(googleRefreshToken) > 0 Then t = Replace(t, googleRefreshToken, "[REDACTED_GOOGLE_REFRESH_TOKEN]")
    If Len(sefApiKey) > 0 Then t = Replace(t, sefApiKey, "[REDACTED_SEF_API_KEY]")

    If Len(t) > 1000 Then t = Left$(t, 1000) & "...[TRUNCATED]"

    Monitoring_SanitizeText = t
End Function

Private Function Monitoring_SanitizePayloadJson(ByVal payloadJson As String) As String
    On Error GoTo Fallback

    Dim t As String
    t = CStr(payloadJson)

    If Len(t) > 3000 Then
        Monitoring_SanitizePayloadJson = "{""truncated"":true,""originalLength"":" & CStr(Len(t)) & "}"
        Exit Function
    End If

    t = Monitoring_SanitizeText(t)
    If Left$(Trim$(t), 1) <> "{" Then t = "{}"

    Monitoring_SanitizePayloadJson = t
    Exit Function

Fallback:
    Monitoring_SanitizePayloadJson = "{}"
End Function

Private Function JsonString(ByVal value As String) As String
    On Error Resume Next

    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    value = Replace(value, vbTab, "\t")

    JsonString = """" & value & """"
End Function

Private Function IsoNow() As String
    On Error Resume Next
    IsoNow = Format$(Now, "yyyy-mm-dd\THH:nn:ss")
End Function


