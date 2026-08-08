Attribute VB_Name = "modSEFClient"
Option Explicit

Public Function SubmitUBLInvoice(ByVal ublXml As String, ByVal requestId As String) As clsSEFResponse

    Dim resp As clsSEFResponse
    Dim http As Object

    Dim baseUrl As String
    Dim apiKey As String
    Dim envName As String
    Dim submitUrl As String

    On Error GoTo EH

    Set resp = New clsSEFResponse

    If Len(Trim$(ublXml)) = 0 Then
        Err.Raise ERR_SEF_HTTP, "SubmitUBLInvoice", "UBL XML is empty."
    End If

    If Len(Trim$(requestId)) = 0 Then
        Err.Raise ERR_SEF_HTTP, "SubmitUBLInvoice", "requestId is empty."
    End If

    GetSEFClientConfig baseUrl, apiKey, envName, "SubmitUBLInvoice"

    submitUrl = BuildSubmitUBLUrl(baseUrl, requestId)

    Set http = CreateSEFHttpRequest()

    http.Open "POST", submitUrl, False
    ApplySEFHeaders http, apiKey, envName, "application/xml; charset=utf-8"

    http.Send ublXml

    resp.httpStatus = CLng(http.status)
    resp.rawBody = CStr(http.responseText)

    If resp.httpStatus = 429 Then
        ApplyRateLimitResponse resp, http
    Else
        ParseSubmitResponse resp
    End If

    DebugSEFHttp "SEF submit response", requestId, resp.httpStatus, _
                 resp.rawBody, ExtractTagValue(ublXml, "cbc:ID")

    Set SubmitUBLInvoice = resp
    Exit Function

EH:
    LogErr "SubmitUBLInvoice"

    Set resp = New clsSEFResponse
    resp.httpStatus = 0
    resp.Success = False
    resp.Accepted = False
    resp.Rejected = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.description
    resp.rawBody = ""

    Set SubmitUBLInvoice = resp
End Function

Public Function GetInvoiceStatus(ByVal sefDocumentId As String) As clsSEFResponse

    Dim resp As clsSEFResponse
    Dim http As Object

    Dim baseUrl As String
    Dim apiKey As String
    Dim envName As String
    Dim statusUrl As String

    On Error GoTo EH

    Set resp = New clsSEFResponse

    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_HTTP, "GetInvoiceStatus", "SEF document ID is empty."
    End If

    GetSEFClientConfig baseUrl, apiKey, envName, "GetInvoiceStatus"

    statusUrl = BuildStatusUrl(baseUrl, sefDocumentId)

    Set http = CreateSEFHttpRequest()

    http.Open "GET", statusUrl, False
    ApplySEFHeaders http, apiKey, envName, ""

    http.Send

    resp.httpStatus = CLng(http.status)
    resp.rawBody = CStr(http.responseText)
    resp.sefDocumentId = sefDocumentId

    If resp.httpStatus = 429 Then
        ApplyRateLimitResponse resp, http
    Else
    ParseStatusResponse resp
    End If
    DebugSEFHttp "SEF status response", sefDocumentId, resp.httpStatus, resp.rawBody

    Set GetInvoiceStatus = resp
    Exit Function

EH:
    LogErr "GetInvoiceStatus"

    Set resp = New clsSEFResponse
    resp.httpStatus = 0
    resp.Success = False
    resp.Accepted = False
    resp.Rejected = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.description
    resp.rawBody = ""
    resp.sefDocumentId = sefDocumentId

    Set GetInvoiceStatus = resp
End Function

Public Function CancelInvoiceOnSEF(ByVal sefDocumentId As String, ByVal cancelComment As String) As clsSEFResponse

    Dim resp As clsSEFResponse
    Dim http As Object

    Dim baseUrl As String
    Dim apiKey As String
    Dim envName As String
    Dim cancelUrl As String
    Dim body As String

    On Error GoTo EH

    Set resp = New clsSEFResponse

    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_HTTP, "CancelInvoiceOnSEF", "SEF document ID is empty."
    End If

    If Len(Trim$(cancelComment)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "CancelInvoiceOnSEF", "Cancel comment is required."
    End If

    GetSEFClientConfig baseUrl, apiKey, envName, "CancelInvoiceOnSEF"

    cancelUrl = BuildCancelUrl(baseUrl)
    body = "{""invoiceId"":" & GetJsonNumericIdLiteral(sefDocumentId, "CancelInvoiceOnSEF") & _
       ",""cancelComments"":""" & JsonEscape(cancelComment) & """}"

    Set http = CreateSEFHttpRequest()

    http.Open "POST", cancelUrl, False
    ApplySEFHeaders http, apiKey, envName, "application/json; charset=utf-8"

    http.Send body

    resp.httpStatus = CLng(http.status)
    resp.rawBody = CStr(http.responseText)
    resp.sefDocumentId = sefDocumentId

    If resp.httpStatus = 429 Then
        ApplyRateLimitResponse resp, http

    ElseIf resp.httpStatus >= 200 And resp.httpStatus < 300 Then
        resp.Success = True
        resp.apiStatus = UCase$(FirstNonEmpty( _
            ExtractJsonString(resp.rawBody, "Status"), _
            "CANCELLED"))
    Else
        resp.Success = False
        resp.apiStatus = "FAILED"
        resp.errorCode = CStr(resp.httpStatus)
        resp.errorMessage = BuildHttpErrorMessage( _
            "HTTP error during SEF cancel.", resp.rawBody)
    End If

    DebugSEFHttp "SEF cancel response", sefDocumentId, resp.httpStatus, resp.rawBody

    Set CancelInvoiceOnSEF = resp
    Exit Function

EH:
    LogErr "CancelInvoiceOnSEF"

    Set resp = New clsSEFResponse
    resp.httpStatus = 0
    resp.Success = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.description
    resp.sefDocumentId = sefDocumentId

    Set CancelInvoiceOnSEF = resp
End Function

Public Function StornoInvoiceOnSEF(ByVal sefDocumentId As String, ByVal stornoComment As String, Optional ByVal stornoNumber As String = "") As clsSEFResponse

    Dim resp As clsSEFResponse
    Dim http As Object

    Dim baseUrl As String
    Dim apiKey As String
    Dim envName As String
    Dim stornoUrl As String
    Dim body As String

    On Error GoTo EH

    Set resp = New clsSEFResponse

    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_HTTP, "StornoInvoiceOnSEF", "SEF document ID is empty."
    End If

    If Len(Trim$(stornoComment)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "StornoInvoiceOnSEF", "Storno comment is required."
    End If

    GetSEFClientConfig baseUrl, apiKey, envName, "StornoInvoiceOnSEF"

    stornoUrl = BuildStornoUrl(baseUrl)
    body = "{""invoiceId"":" & GetJsonNumericIdLiteral(sefDocumentId, "StornoInvoiceOnSEF") & _
            ",""stornoNumber"":""" & JsonEscape(stornoNumber) & _
            """,""stornoComment"":""" & JsonEscape(stornoComment) & """}"

    Set http = CreateSEFHttpRequest()

    http.Open "POST", stornoUrl, False
    ApplySEFHeaders http, apiKey, envName, "application/json; charset=utf-8"

    http.Send body

    resp.httpStatus = CLng(http.status)
    resp.rawBody = CStr(http.responseText)
    resp.sefDocumentId = sefDocumentId

    If resp.httpStatus = 429 Then
        ApplyRateLimitResponse resp, http

    ElseIf resp.httpStatus >= 200 And resp.httpStatus < 300 Then
        resp.Success = True
        resp.apiStatus = UCase$(FirstNonEmpty( _
            ExtractJsonString(resp.rawBody, "Status"), _
            "STORNO"))
    Else
        resp.Success = False
        resp.apiStatus = "FAILED"
        resp.errorCode = CStr(resp.httpStatus)
        resp.errorMessage = BuildHttpErrorMessage( _
            "HTTP error during SEF storno.", resp.rawBody)
    End If

    DebugSEFHttp "SEF storno response", sefDocumentId, resp.httpStatus, resp.rawBody

    Set StornoInvoiceOnSEF = resp
    Exit Function

EH:
    LogErr "StornoInvoiceOnSEF"

    Set resp = New clsSEFResponse
    resp.httpStatus = 0
    resp.Success = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.description
    resp.sefDocumentId = sefDocumentId

    Set StornoInvoiceOnSEF = resp
End Function
Private Sub GetSEFClientConfig(ByRef baseUrl As String, _
                               ByRef apiKey As String, _
                               ByRef envName As String, _
                               ByVal sourceName As String)
    On Error GoTo EH

    baseUrl = Trim$(GetConfigValue("SEF_BASE_URL"))
    apiKey = Trim$(GetConfigValue("SEF_API_KEY"))
    envName = Trim$(GetConfigValue("SEF_ENV"))

    If Len(baseUrl) = 0 Then
        Err.Raise ERR_SEF_CONFIG, sourceName, _
                  "SEF_BASE_URL missing in tblSEFConfig."
    End If

    If Len(apiKey) = 0 Then
        Err.Raise ERR_SEF_CONFIG, sourceName, _
                  "SEF_API_KEY missing in tblSEFConfig."
    End If

    If LCase$(Left$(baseUrl, 8)) <> "https://" Then
        Err.Raise ERR_SEF_CONFIG, sourceName, _
              "SEF_BASE_URL must start with https://. Plain HTTP is not allowed for SEF."
    End If

    Exit Sub

EH:
    LogErr "modSEFClient.GetSEFClientConfig"
    Err.Raise Err.Number, sourceName, Err.description
End Sub

Private Function CreateSEFHttpRequest() As Object
    On Error GoTo EH

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, _
                     HTTP_TIMEOUT_CONNECT_MS, _
                     HTTP_TIMEOUT_SEND_MS, _
                     HTTP_TIMEOUT_RECEIVE_MS

    Set CreateSEFHttpRequest = http
    Exit Function

EH:
    LogErr "modSEFClient.CreateSEFHttpRequest"
    Err.Raise Err.Number, "modSEFClient.CreateSEFHttpRequest", Err.description
End Function

Private Sub ApplySEFHeaders(ByVal http As Object, _
                            ByVal apiKey As String, _
                            ByVal envName As String, _
                            ByVal contentType As String)
    On Error GoTo EH

    If http Is Nothing Then
        Err.Raise ERR_SEF_HTTP, "modSEFClient.ApplySEFHeaders", _
                  "HTTP object is Nothing."
    End If

    http.SetRequestHeader "Accept", "application/json"

    If Len(Trim$(contentType)) > 0 Then
        http.SetRequestHeader "Content-Type", contentType
    End If

    http.SetRequestHeader "ApiKey", apiKey

    If Len(Trim$(envName)) > 0 Then
        http.SetRequestHeader "X-SEF-ENV", envName
    End If

    Exit Sub

EH:
    LogErr "modSEFClient.ApplySEFHeaders"
    Err.Raise Err.Number, "modSEFClient.ApplySEFHeaders", Err.description
End Sub

Private Function IsSEFDebugEnabled() As Boolean
    IsSEFDebugEnabled = (UCase$(Trim$(GetConfigValue("SEF_DEBUG_LOG"))) = "DA")
End Function

Private Sub DebugSEFHttp(ByVal caption As String, _
                         ByVal requestId As String, _
                         ByVal httpStatus As Long, _
                         ByVal responseText As String, _
                         Optional ByVal xmlIdMarker As String = "")
    On Error Resume Next

    If Not IsSEFDebugEnabled() Then Exit Sub

    Debug.Print "--------------------------------"
    Debug.Print caption

    If Len(Trim$(requestId)) > 0 Then
        Debug.Print "RequestId: " & requestId
    End If

    If Len(Trim$(xmlIdMarker)) > 0 Then
        Debug.Print "Invoice XML ID marker: " & xmlIdMarker
    End If

    Debug.Print "HTTP Status: " & CStr(httpStatus)
    Debug.Print "ResponseText: " & Left$(responseText, 2000)
    Debug.Print "--------------------------------"
End Sub

Private Function GetHeaderSafe(ByVal http As Object, ByVal headerName As String) As String
    On Error Resume Next
    GetHeaderSafe = Trim$(CStr(http.GetResponseHeader(headerName)))
End Function

Private Function BuildHttpErrorMessage(ByVal defaultMessage As String, _
                                       ByVal rawBody As String) As String
    BuildHttpErrorMessage = FirstNonEmpty( _
        ExtractJsonString(rawBody, "Message"), _
        ExtractJsonString(rawBody, "message"), _
        ExtractJsonString(rawBody, "error"), _
        defaultMessage)
End Function

Private Function BuildSubmitUBLUrl(ByVal baseUrl As String, ByVal requestId As String) As String
    
    Dim s As String
    s = Trim$(baseUrl)
    
    If Right$(s, 1) = "/" Then
        s = Left$(s, Len(s) - 1)
    End If
    
    BuildSubmitUBLUrl = s & "/api/publicApi/sales-invoice/ubl?requestId=" & UrlEncode(requestId)
End Function

Private Function BuildStatusUrl(ByVal baseUrl As String, ByVal sefDocumentId As String) As String
    
    Dim s As String
    s = Trim$(baseUrl)
    
    If Right$(s, 1) = "/" Then
        s = Left$(s, Len(s) - 1)
    End If
    
    BuildStatusUrl = s & "/api/publicApi/sales-invoice?invoiceId=" & UrlEncode(sefDocumentId)
End Function

Private Function BuildCancelUrl(ByVal baseUrl As String) As String
    
    Dim s As String
    s = Trim$(baseUrl)
    
    If Right$(s, 1) = "/" Then
        s = Left$(s, Len(s) - 1)
    End If
    
    BuildCancelUrl = s & "/api/publicApi/sales-invoice/cancel"
End Function

Private Function BuildStornoUrl(ByVal baseUrl As String) As String
    
    Dim s As String
    s = Trim$(baseUrl)
    
    If Right$(s, 1) = "/" Then
        s = Left$(s, Len(s) - 1)
    End If
    
    BuildStornoUrl = s & "/api/publicApi/sales-invoice/storno"
End Function

Private Sub ParseSubmitResponse(ByRef resp As clsSEFResponse)
    
    Dim body As String
    body = resp.rawBody
    
    Select Case resp.httpStatus
        
        Case 200, 201, 202
            resp.Success = True
            resp.apiStatus = "SENT"
            resp.sefDocumentId = FirstNonEmpty( _
                ExtractJsonNumberOrStringAsString(body, "SalesInvoiceId"), _
                ExtractJsonNumberOrStringAsString(body, "InvoiceId"), _
                ExtractJsonNumberOrStringAsString(body, "PurchaseInvoiceId"))

            resp.SEFInvoiceNumber = ""
            resp.correlationId = ""
            
            If ExtractJsonBoolean(body, "accepted", False) Then
                resp.Accepted = True
                resp.apiStatus = "ACCEPTED"
            End If
        
        Case 400, 422
            resp.Success = False
            resp.Rejected = True
            resp.apiStatus = "REJECTED"
            resp.errorCode = ExtractJsonString(body, "errorCode")
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "message"), _
                ExtractJsonString(body, "error"), _
                "SEF rejected request.")

        ' AUD-030: HTTP 409 = the invoice already exists / conflicts on SEF.
        ' It must NOT be marked REJECTED. REJECTED would enable a corrective
        ' resubmit with a fresh requestId (PrepareRejectedInvoiceForResubmit),
        ' risking a duplicate/incorrect invoice toward the tax authority while
        ' the original document still exists on SEF. Treat 409 as a non-final
        ' CONFLICT: Success=False and Rejected=False, so the send pipeline
        ' routes it to SEF_TECH_FAILED / manual review, where a retry reuses
        ' the SAME requestId (idempotent) instead of auto-rejecting.
        Case 409
            resp.Success = False
            resp.Rejected = False
            resp.apiStatus = "CONFLICT"
            resp.errorCode = FirstNonEmpty( _
                ExtractJsonString(body, "errorCode"), _
                "409")
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "message"), _
                ExtractJsonString(body, "error"), _
                "SEF conflict (409): invoice may already exist. Manual reconciliation required.")

        ' In ParseStatusResponse / ParseSubmitResponse:
        ' Fallback only. Normal 429 handling is done before parser.
        Case 429
            resp.Success = False
            resp.apiStatus = "RATE_LIMITED"
            resp.errorCode = "429"
            resp.errorMessage = "Rate limit exceeded. Retry after delay."
            ' + Retry-After Header auslesen wenn vorhanden
        
        Case Else
            resp.Success = False
            resp.apiStatus = "FAILED"
            resp.errorCode = CStr(resp.httpStatus)
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "message"), _
                ExtractJsonString(body, "error"), _
                "HTTP error during SEF submit.")
    End Select
End Sub

Private Sub ParseStatusResponse(ByRef resp As clsSEFResponse)
    
    Dim body As String
    Dim statusValue As String
    
    body = resp.rawBody
    
    If resp.httpStatus < 200 Or resp.httpStatus >= 300 Then
        resp.Success = False
        resp.Accepted = False
        resp.Rejected = False
        resp.apiStatus = "FAILED"
        resp.errorCode = CStr(resp.httpStatus)
        resp.errorMessage = FirstNonEmpty( _
            ExtractJsonString(body, "Message"), _
            ExtractJsonString(body, "message"), _
            ExtractJsonString(body, "error"), _
            "HTTP error during SEF status query.")
        Exit Sub
    End If
    
    resp.Success = True
    
    statusValue = UCase$(Trim$(FirstNonEmpty( _
        ExtractJsonString(body, "Status"), _
        ExtractJsonString(body, "status"), _
        ExtractJsonString(body, "invoiceStatus"))))
    
    resp.sefDocumentId = FirstNonEmpty( _
        ExtractJsonNumberOrStringAsString(body, "InvoiceId"), _
        resp.sefDocumentId)
    
    resp.correlationId = ExtractJsonString(body, "GlobUniqId")
    
    ' ApiStatus is the exact external SEF status, stored verbatim.
    ' The MEANING of that status (approved / rejected / pending / terminal /
    ' informational / unknown) is decided in ONE place --
    ' modSEFStatusSync.ClassifySEFExternalStatus -- so the client and the
    ' workflow layer cannot drift apart.
    '
    ' AUD-032b: the old hand-written list here knew only "ACCEPTED", while the
    ' official SEF enum (SalesInvoiceStatus) calls approval **Approved**, and
    ' also emits Seen / Paid / OverDue / Archived / Mistake / Deleted / Sending.
    ' Everything outside the old list fell into a "SENT" fallback, which is how
    ' an unconfirmed invoice could look delivered.
    resp.apiStatus = statusValue
    If Len(Trim$(resp.apiStatus)) = 0 Then resp.apiStatus = SEF_STATUS_UNKNOWN

    resp.Accepted = False
    resp.Rejected = False

    Select Case ClassifySEFExternalStatus(resp.apiStatus)

        Case SEF_CLS_ACCEPTED
            resp.Accepted = True

        Case SEF_CLS_REJECTED
            resp.Rejected = True
            resp.errorCode = FirstNonEmpty( _
                ExtractJsonString(body, "ErrorCode"), _
                ExtractJsonString(body, "errorCode"))
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "Message"), _
                ExtractJsonString(body, "message"), _
                "SEF rejected invoice.")

        Case SEF_CLS_ERROR
            ' Document-level ERROR status: the call worked, the document did not.
            resp.Success = False
            resp.errorCode = FirstNonEmpty( _
                ExtractJsonString(body, "ErrorCode"), _
                ExtractJsonString(body, "errorCode"), _
                "SEF_STATUS_ERROR")
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "Message"), _
                ExtractJsonString(body, "message"), _
                "SEF returned ERROR status.")

        Case Else
            ' PENDING / TERMINAL / INFO / UNKNOWN: flags stay False and the
            ' verbatim status travels up; the workflow layer decides.

    End Select

End Sub

Private Function ExtractJsonString(ByVal json As String, ByVal key As String) As String
    
    Dim p As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim pattern As String
    
    pattern = """" & key & """"
    p = InStr(1, json, pattern, vbTextCompare)
    
    If p = 0 Then Exit Function
    
    startPos = p + Len(pattern)
    
    Do While startPos <= Len(json)
        Select Case Mid$(json, startPos, 1)
            Case " ", vbTab, vbCr, vbLf
                startPos = startPos + 1
            Case ":"
                startPos = startPos + 1
                Exit Do
            Case Else
                Exit Function
        End Select
    Loop
    
    Do While startPos <= Len(json)
        Select Case Mid$(json, startPos, 1)
            Case " ", vbTab, vbCr, vbLf
                startPos = startPos + 1
            Case """"
                startPos = startPos + 1
                Exit Do
            Case Else
                Exit Function
        End Select
    Loop
    
    endPos = startPos
    
    Do While endPos <= Len(json)
        If Mid$(json, endPos, 1) = """" Then Exit Do
        endPos = endPos + 1
    Loop
    
    If endPos > startPos Then
        ExtractJsonString = Mid$(json, startPos, endPos - startPos)
    End If

End Function

Private Function ExtractJsonNumberAsString(ByVal json As String, ByVal key As String) As String
    
    Dim pattern As String
    Dim p As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String
    Dim result As String
    
    pattern = """" & key & """:"
    p = InStr(1, json, pattern, vbTextCompare)
    
    If p = 0 Then
        ExtractJsonNumberAsString = ""
        Exit Function
    End If
    
    startPos = p + Len(pattern)
    
    Do While startPos <= Len(json)
        ch = Mid$(json, startPos, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        startPos = startPos + 1
    Loop
    
    endPos = startPos
    
    Do While endPos <= Len(json)
        ch = Mid$(json, endPos, 1)
        If (ch < "0" Or ch > "9") Then Exit Do
        result = result & ch
        endPos = endPos + 1
    Loop
    
    ExtractJsonNumberAsString = Trim$(result)

End Function

Private Function ExtractJsonNumberOrStringAsString(ByVal json As String, _
                                                   ByVal key As String) As String
    Dim asNumber As String
    Dim asString As String

    asNumber = ExtractJsonNumberAsString(json, key)

    If Len(Trim$(asNumber)) > 0 Then
        ExtractJsonNumberOrStringAsString = asNumber
        Exit Function
    End If

    asString = ExtractJsonString(json, key)

    If Len(Trim$(asString)) > 0 Then
        ExtractJsonNumberOrStringAsString = Trim$(asString)
    End If
End Function

Private Function ExtractJsonBoolean(ByVal json As String, _
                                    ByVal key As String, _
                                    Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim pattern As String
    Dim p As Long
    Dim startPos As Long
    Dim token As String
    Dim ch As String

    pattern = """" & key & """"
    p = InStr(1, json, pattern, vbTextCompare)

    If p = 0 Then
        ExtractJsonBoolean = defaultValue
        Exit Function
    End If

    startPos = p + Len(pattern)

    Do While startPos <= Len(json)
        ch = Mid$(json, startPos, 1)

        Select Case ch
            Case " ", vbTab, vbCr, vbLf
                startPos = startPos + 1
            Case ":"
                startPos = startPos + 1
                Exit Do
            Case Else
                ExtractJsonBoolean = defaultValue
                Exit Function
        End Select
    Loop

    Do While startPos <= Len(json)
        ch = Mid$(json, startPos, 1)

        Select Case ch
            Case " ", vbTab, vbCr, vbLf
                startPos = startPos + 1
            Case Else
                Exit Do
        End Select
    Loop

    token = UCase$(Mid$(json, startPos, 5))

    If Left$(token, 4) = "TRUE" Then
        ExtractJsonBoolean = True
    ElseIf Left$(token, 5) = "FALSE" Then
        ExtractJsonBoolean = False
    Else
        ExtractJsonBoolean = defaultValue
    End If
End Function

Private Function FirstNonEmpty(ParamArray values() As Variant) As String
    
    Dim i As Long
    
    For i = LBound(values) To UBound(values)
        If Len(Trim$(CStr(values(i)))) > 0 Then
            FirstNonEmpty = Trim$(CStr(values(i)))
            Exit Function
        End If
    Next i
    
    FirstNonEmpty = ""
End Function

Private Function ExtractTagValue(ByVal xml As String, ByVal tagName As String) As String
    
    Dim openTag As String
    Dim closeTag As String
    Dim p1 As Long
    Dim p2 As Long
    
    openTag = "<" & tagName & ">"
    closeTag = "</" & tagName & ">"
    
    p1 = InStr(1, xml, openTag, vbTextCompare)
    If p1 = 0 Then Exit Function
    
    p1 = p1 + Len(openTag)
    p2 = InStr(p1, xml, closeTag, vbTextCompare)
    If p2 = 0 Then Exit Function
    
    ExtractTagValue = Mid$(xml, p1, p2 - p1)

End Function

Private Function GetJsonNumericIdLiteral(ByVal rawID As String, _
                                         ByVal sourceName As String) As String
    ' Returns a JSON-ready literal for SEF document IDs.
    ' - Pure numeric digits: returns raw digits for embedding as JSON number.
    ' - GUID-like (hex+hyphens): returns quoted string literal with quotes
    '   already applied, so caller-side concatenation produces valid JSON.
    ' - Empty or unrecognized shape: raises ERR_SEF_VALIDATION.
    
    On Error GoTo EH
    
    Dim s As String
    s = Trim$(rawID)
    
    If s = "" Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, _
                  "SEF document ID is empty."
    End If
    
    If IsAllDigits(s) Then
        ' JSON numeric literal - no quotes, exact digits preserved.
        ' Works for IDs beyond VBA Long range because we never CLng the value.
        GetJsonNumericIdLiteral = s
        Exit Function
    End If
    
    If IsGuidLike(s) Then
        ' JSON string literal - quoted. GUID characters do not need escaping.
        GetJsonNumericIdLiteral = """" & s & """"
        Exit Function
    End If
    
    Err.Raise ERR_SEF_VALIDATION, sourceName, _
              "SEF document ID has unrecognized shape: " & s
    Exit Function

EH:
    LogErr "modSEFClient.GetJsonNumericIdLiteral"
    Err.Raise Err.Number, sourceName, Err.description
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    
    If Len(s) = 0 Then
        IsAllDigits = False
        Exit Function
    End If
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then
            IsAllDigits = False
            Exit Function
        End If
    Next i
    
    IsAllDigits = True
End Function

Private Function IsGuidLike(ByVal s As String) As Boolean
    ' Accepts:
    '   - hyphenated GUID: 8-4-4-4-12 hex digits (with optional curly braces)
    '   - bare 32-hex (no hyphens)
    '   - vendor-specific opaque IDs that mix hex digits and hyphens
    ' Rejects:
    '   - empty/whitespace
    '   - strings with non-hex non-hyphen characters
    '   - all-numeric strings (those are NUMERIC shape, not GUID)
    
    Dim s2 As String
    Dim i As Long
    Dim ch As String
    Dim hasNonDigit As Boolean
    
    s2 = s
    
    ' Strip optional curly braces
    If Left$(s2, 1) = "{" And Right$(s2, 1) = "}" Then
        s2 = Mid$(s2, 2, Len(s2) - 2)
    End If
    
    If Len(s2) < 8 Then
        IsGuidLike = False
        Exit Function
    End If
    
    For i = 1 To Len(s2)
        ch = LCase$(Mid$(s2, i, 1))
        Select Case ch
            Case "0" To "9", "a" To "f", "-"
                If (ch >= "a" And ch <= "f") Or ch = "-" Then
                    hasNonDigit = True
                End If
            Case Else
                IsGuidLike = False
                Exit Function
        End Select
    Next i
    
    ' Must contain at least one hex letter or hyphen, otherwise IsAllDigits
    ' would already have matched and we would not be in IsGuidLike check.
    IsGuidLike = hasNonDigit
End Function



Private Sub ApplyRateLimitResponse(ByVal resp As clsSEFResponse, _
                                   ByVal http As Object)
    On Error Resume Next

    resp.Success = False
    resp.Accepted = False
    resp.Rejected = False
    resp.apiStatus = "RATE_LIMITED"
    resp.errorCode = "429"

    Dim retryAfter As String
    retryAfter = GetHeaderSafe(http, "Retry-After")

    If Len(retryAfter) > 0 Then
        resp.errorMessage = "Rate limit exceeded. Retry-After: " & retryAfter
    Else
        resp.errorMessage = "Rate limit exceeded. Retry after delay."
    End If
End Sub

Public Sub Test_SubmitUBLInvoice()

    On Error GoTo EH
    
    Dim dto As clsSEFInvoiceSnapshot
    Dim xml As String
    Dim resp As clsSEFResponse
    Dim requestId As String
    
    Set dto = BuildSEFInvoiceDto("FAK-00001")
    xml = SerializeUBLInvoice(dto)
    
    requestId = "TEST-" & Format$(Now, "yyyymmddhhnnss")
    
    Set resp = SubmitUBLInvoice(xml, requestId)
    
    Debug.Print "RequestId: "; requestId
    Debug.Print "HTTP: "; resp.httpStatus
    Debug.Print "Success: "; resp.Success
    Debug.Print "Accepted: "; resp.Accepted
    Debug.Print "Rejected: "; resp.Rejected
    Debug.Print "ApiStatus: "; resp.apiStatus
    Debug.Print "SEFDocumentId: "; resp.sefDocumentId
    Debug.Print "ErrorCode: "; resp.errorCode
    Debug.Print "ErrorMessage: "; resp.errorMessage
    Debug.Print "RawBody: "; resp.rawBody
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub

Public Sub RunSEFClientParserSmokeSuite()
    Dim failed As Long

    If ExtractJsonString("{""message"":""OK""}", "message") <> "OK" Then failed = failed + 1
    If ExtractJsonNumberOrStringAsString("{""InvoiceId"":12345}", "InvoiceId") <> "12345" Then failed = failed + 1
    If ExtractJsonNumberOrStringAsString("{""InvoiceId"":""12345""}", "InvoiceId") <> "12345" Then failed = failed + 1
    If ExtractJsonBoolean("{""accepted"": true}", "accepted", False) <> True Then failed = failed + 1
    If ExtractJsonBoolean("{""accepted"": false}", "accepted", True) <> False Then failed = failed + 1

    ' Known weakness documentation: escaped strings are not fully decoded by manual parser.
    Debug.Print "Escaped message parser result: "; ExtractJsonString("{""message"":""A \""quoted\"" value""}", "message")

    If failed > 0 Then
        Err.Raise ERR_SEF_VALIDATION, "RunSEFClientParserSmokeSuite", _
                  "SEF client parser smoke failed: " & failed
    End If

    Debug.Print "RunSEFClientParserSmokeSuite PASS"
End Sub

Public Function TestProxyForGetJsonNumericIdLiteral(ByVal rawID As String) As String
    ' Test-only proxy. Forwards to the private GetJsonNumericIdLiteral
    ' so RunSEFDocumentIdShapeSuite in modSEFTests can verify behavior.
    ' Not for production use - business code calls the private helper
    ' directly within modSEFClient (CancelInvoiceOnSEF, StornoInvoiceOnSEF).
    
    TestProxyForGetJsonNumericIdLiteral = _
        GetJsonNumericIdLiteral(rawID, "TestProxyForGetJsonNumericIdLiteral")
End Function

Public Function TestProxyForParseSubmitResponse(ByVal httpStatus As Long, _
                                                ByVal rawBody As String) As clsSEFResponse
    ' Test-only proxy. Forwards to the private ParseSubmitResponse so
    ' RunSEFOfflineSuite can verify HTTP status -> apiStatus classification
    ' (notably 409 -> CONFLICT, not REJECTED). Not used by production code.

    Dim resp As clsSEFResponse
    Set resp = New clsSEFResponse

    resp.httpStatus = httpStatus
    resp.rawBody = rawBody

    ParseSubmitResponse resp

    Set TestProxyForParseSubmitResponse = resp
End Function

Public Function TestProxyForParseStatusResponse(ByVal httpStatus As Long, _
                                                ByVal rawBody As String) As clsSEFResponse
    ' Test-only proxy. Forwards to the private ParseStatusResponse so
    ' RunSEFTestSuite can verify status -> apiStatus classification without an
    ' HTTP call (notably: blank status -> UNKNOWN_STATUS, never SENT).
    ' Not used by production code.

    Dim resp As clsSEFResponse
    Set resp = New clsSEFResponse

    resp.httpStatus = httpStatus
    resp.rawBody = rawBody

    ParseStatusResponse resp

    Set TestProxyForParseStatusResponse = resp
End Function

