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
    
    baseUrl = GetConfigValue("SEF_BASE_URL")
    apiKey = GetConfigValue("SEF_API_KEY")
    envName = GetConfigValue("SEF_ENV")
    
    If Len(Trim$(baseUrl)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "SubmitUBLInvoice", "SEF_BASE_URL missing in tblSEFConfig."
    End If
    
    If Len(Trim$(apiKey)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "SubmitUBLInvoice", "SEF_API_KEY missing in tblSEFConfig."
    End If
    
    submitUrl = BuildSubmitUBLUrl(baseUrl, requestId)
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, _
                     HTTP_TIMEOUT_CONNECT_MS, _
                     HTTP_TIMEOUT_SEND_MS, _
                     HTTP_TIMEOUT_RECEIVE_MS
    
    http.Open "POST", submitUrl, False
    
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "Content-Type", "application/xml; charset=utf-8"
    http.SetRequestHeader "ApiKey", apiKey
    
    If Len(Trim$(envName)) > 0 Then
        http.SetRequestHeader "X-SEF-ENV", envName
    End If
    
    http.Send ublXml
    
    Debug.Print "--------------------------------"
    Debug.Print "RequestId: " & requestId
    Debug.Print "Invoice XML ID marker: " & ExtractTagValue(ublXml, "cbc:ID")
    Debug.Print "HTTP Status: " & http.Status
    Debug.Print "ResponseText: " & CStr(http.responseText)
    Debug.Print "--------------------------------"
     
    resp.HttpStatus = CLng(http.Status)
    resp.RawBody = CStr(http.responseText)
    
    ParseSubmitResponse resp
    
    Set SubmitUBLInvoice = resp
    Exit Function

EH:
    LogErr "SubmitUBLInvoice"
    Set resp = New clsSEFResponse
    resp.HttpStatus = 0
    resp.Success = False
    resp.Accepted = False
    resp.Rejected = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.Description
    resp.RawBody = ""
    
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
    
    baseUrl = GetConfigValue("SEF_BASE_URL")
    apiKey = GetConfigValue("SEF_API_KEY")
    envName = GetConfigValue("SEF_ENV")
    
    If Len(Trim$(baseUrl)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "GetInvoiceStatus", "SEF_BASE_URL missing in tblSEFConfig."
    End If
    
    If Len(Trim$(apiKey)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "GetInvoiceStatus", "SEF_API_KEY missing in tblSEFConfig."
    End If
    
    statusUrl = BuildStatusUrl(baseUrl, sefDocumentId)
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, _
                     HTTP_TIMEOUT_CONNECT_MS, _
                     HTTP_TIMEOUT_SEND_MS, _
                     HTTP_TIMEOUT_RECEIVE_MS
    
    http.Open "GET", statusUrl, False
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "ApiKey", apiKey
    
    If Len(Trim$(envName)) > 0 Then
        http.SetRequestHeader "X-SEF-ENV", envName
    End If
    
    http.Send
    
    resp.HttpStatus = CLng(http.Status)
    resp.RawBody = CStr(http.responseText)
    resp.sefDocumentId = sefDocumentId
    
    ParseStatusResponse resp
    
    Set GetInvoiceStatus = resp
    Exit Function

EH:
    LogErr "GetInvoiceStatus"
    Set resp = New clsSEFResponse
    resp.HttpStatus = 0
    resp.Success = False
    resp.Accepted = False
    resp.Rejected = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.Description
    resp.RawBody = ""
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
    
    baseUrl = GetConfigValue("SEF_BASE_URL")
    apiKey = GetConfigValue("SEF_API_KEY")
    envName = GetConfigValue("SEF_ENV")
    
    If Len(Trim$(baseUrl)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "CancelInvoiceOnSEF", "SEF_BASE_URL missing in tblSEFConfig."
    End If
    
    If Len(Trim$(apiKey)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "CancelInvoiceOnSEF", "SEF_API_KEY missing in tblSEFConfig."
    End If
    
    cancelUrl = BuildCancelUrl(baseUrl)
    
    body = "{""invoiceId"":" & CLng(sefDocumentId) & ",""cancelComments"":""" & JsonEscape(cancelComment) & """}"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, _
                     HTTP_TIMEOUT_CONNECT_MS, _
                     HTTP_TIMEOUT_SEND_MS, _
                     HTTP_TIMEOUT_RECEIVE_MS
    
    http.Open "POST", cancelUrl, False
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "ApiKey", apiKey
    
    If Len(Trim$(envName)) > 0 Then
        http.SetRequestHeader "X-SEF-ENV", envName
    End If
    
    http.Send body
    
    resp.HttpStatus = CLng(http.Status)
    resp.RawBody = CStr(http.responseText)
    resp.sefDocumentId = sefDocumentId
    
    If resp.HttpStatus >= 200 And resp.HttpStatus < 300 Then
        resp.Success = True
        resp.apiStatus = UCase$(FirstNonEmpty( _
            ExtractJsonString(resp.RawBody, "Status"), _
            "CANCELLED"))
    Else
        resp.Success = False
        resp.apiStatus = "FAILED"
        resp.errorCode = CStr(resp.HttpStatus)
        resp.errorMessage = FirstNonEmpty( _
            ExtractJsonString(resp.RawBody, "Message"), _
            ExtractJsonString(resp.RawBody, "message"), _
            ExtractJsonString(resp.RawBody, "error"), _
            "HTTP error during SEF cancel.")
    End If
    
    Set CancelInvoiceOnSEF = resp
    Exit Function

EH:
    LogErr "CancelInvoiceOnSef"
    Set resp = New clsSEFResponse
    resp.HttpStatus = 0
    resp.Success = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.Description
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
    
    baseUrl = GetConfigValue("SEF_BASE_URL")
    apiKey = GetConfigValue("SEF_API_KEY")
    envName = GetConfigValue("SEF_ENV")
    
    If Len(Trim$(baseUrl)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "StornoInvoiceOnSEF", "SEF_BASE_URL missing in tblSEFConfig."
    End If
    
    If Len(Trim$(apiKey)) = 0 Then
        Err.Raise ERR_SEF_CONFIG, "StornoInvoiceOnSEF", "SEF_API_KEY missing in tblSEFConfig."
    End If
    
    stornoUrl = BuildStornoUrl(baseUrl)
    
    body = "{""invoiceId"":" & CLng(sefDocumentId) & ",""stornoNumber"":""" & JsonEscape(stornoNumber) & """,""stornoComment"":""" & JsonEscape(stornoComment) & """}"
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.SetTimeouts HTTP_TIMEOUT_RESOLVE_MS, _
                     HTTP_TIMEOUT_CONNECT_MS, _
                     HTTP_TIMEOUT_SEND_MS, _
                     HTTP_TIMEOUT_RECEIVE_MS
    
    http.Open "POST", stornoUrl, False
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "ApiKey", apiKey
    
    If Len(Trim$(envName)) > 0 Then
        http.SetRequestHeader "X-SEF-ENV", envName
    End If
    
    http.Send body
    
    resp.HttpStatus = CLng(http.Status)
    resp.RawBody = CStr(http.responseText)
    resp.sefDocumentId = sefDocumentId
    
    If resp.HttpStatus >= 200 And resp.HttpStatus < 300 Then
        resp.Success = True
        resp.apiStatus = UCase$(FirstNonEmpty( _
            ExtractJsonString(resp.RawBody, "Status"), _
            "STORNO"))
    Else
        resp.Success = False
        resp.apiStatus = "FAILED"
        resp.errorCode = CStr(resp.HttpStatus)
        resp.errorMessage = FirstNonEmpty( _
            ExtractJsonString(resp.RawBody, "Message"), _
            ExtractJsonString(resp.RawBody, "message"), _
            ExtractJsonString(resp.RawBody, "error"), _
            "HTTP error during SEF storno.")
    End If
    
    Set StornoInvoiceOnSEF = resp
    Exit Function

EH:
    LogErr "StornoInvoiceOnSef"
    Set resp = New clsSEFResponse
    resp.HttpStatus = 0
    resp.Success = False
    resp.apiStatus = "HTTP_ERROR"
    resp.errorCode = "HTTP_EXCEPTION"
    resp.errorMessage = Err.Description
    resp.sefDocumentId = sefDocumentId
    Set StornoInvoiceOnSEF = resp
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
    body = resp.RawBody
    
    Select Case resp.HttpStatus
        
        Case 200, 201, 202
            resp.Success = True
            resp.apiStatus = "SENT"
            resp.sefDocumentId = FirstNonEmpty( _
                ExtractJsonNumberAsString(body, "SalesInvoiceId"), _
                ExtractJsonNumberAsString(body, "InvoiceId"), _
                ExtractJsonNumberAsString(body, "PurchaseInvoiceId"))

            resp.SEFInvoiceNumber = ""
            resp.CorrelationId = ""
            
            If InStr(1, body, """accepted"":true", vbTextCompare) > 0 Then
                resp.Accepted = True
                resp.apiStatus = "ACCEPTED"
            End If
        
        Case 400, 409, 422
            resp.Success = False
            resp.Rejected = True
            resp.apiStatus = "REJECTED"
            resp.errorCode = ExtractJsonString(body, "errorCode")
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "message"), _
                ExtractJsonString(body, "error"), _
                "SEF rejected request.")
        
        ' In ParseStatusResponse / ParseSubmitResponse:
        Case 429
            resp.Success = False
            resp.apiStatus = "RATE_LIMITED"
            resp.errorCode = "429"
            resp.errorMessage = "Rate limit exceeded. Retry after delay."
            ' + Retry-After Header auslesen wenn vorhanden
        
        Case Else
            resp.Success = False
            resp.apiStatus = "FAILED"
            resp.errorCode = CStr(resp.HttpStatus)
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "message"), _
                ExtractJsonString(body, "error"), _
                "HTTP error during SEF submit.")
    End Select
End Sub

Private Sub ParseStatusResponse(ByRef resp As clsSEFResponse)
    
    Dim body As String
    Dim statusValue As String
    
    body = resp.RawBody
    
    If resp.HttpStatus < 200 Or resp.HttpStatus >= 300 Then
        resp.Success = False
        resp.Accepted = False
        resp.Rejected = False
        resp.apiStatus = "FAILED"
        resp.errorCode = CStr(resp.HttpStatus)
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
        ExtractJsonNumberAsString(body, "InvoiceId"), _
        resp.sefDocumentId)
    
    resp.CorrelationId = ExtractJsonString(body, "GlobUniqId")
    
    ' ApiStatus is the exact external SEF status.
    ' Higher layers decide whether that status changes internal workflow.
    Select Case statusValue
        
        Case "ACCEPTED"
            resp.Accepted = True
            resp.Rejected = False
            resp.apiStatus = "ACCEPTED"
        
        Case "REJECTED"
            resp.Accepted = False
            resp.Rejected = True
            resp.apiStatus = "REJECTED"
            resp.errorCode = FirstNonEmpty( _
                ExtractJsonString(body, "ErrorCode"), _
                ExtractJsonString(body, "errorCode"))
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "Message"), _
                ExtractJsonString(body, "message"), _
                "SEF rejected invoice.")
        
        Case "SENT"
            resp.Accepted = False
            resp.Rejected = False
            resp.apiStatus = "SENT"
        
        Case "NEW"
            resp.Accepted = False
            resp.Rejected = False
            resp.apiStatus = "NEW"
        
        Case "DRAFT"
            resp.Accepted = False
            resp.Rejected = False
            resp.apiStatus = "DRAFT"
        
        Case "CANCELLED"
            resp.Accepted = False
            resp.Rejected = False
            resp.apiStatus = "CANCELLED"
        
        Case "STORNO"
            resp.Accepted = False
            resp.Rejected = False
            resp.apiStatus = "STORNO"
        
        Case "ERROR"
            resp.Success = False
            resp.Accepted = False
            resp.Rejected = False
            resp.apiStatus = "ERROR"
            resp.errorCode = FirstNonEmpty( _
                ExtractJsonString(body, "ErrorCode"), _
                ExtractJsonString(body, "errorCode"), _
                "SEF_STATUS_ERROR")
            resp.errorMessage = FirstNonEmpty( _
                ExtractJsonString(body, "Message"), _
                ExtractJsonString(body, "message"), _
                "SEF returned ERROR status.")
        
        Case Else
            resp.Accepted = False
            resp.Rejected = False
            resp.apiStatus = FirstNonEmpty(statusValue, "SENT")
    
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

Private Function JsonEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    JsonEscape = t
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

Private Function UrlEncode(ByVal s As String) As String
    
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = Asc(ch)
        
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122
                result = result & ch
            Case Else
                result = result & "%" & Right$("0" & Hex$(code), 2)
        End Select
    Next i
    
    UrlEncode = result
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
    Debug.Print "HTTP: "; resp.HttpStatus
    Debug.Print "Success: "; resp.Success
    Debug.Print "Accepted: "; resp.Accepted
    Debug.Print "Rejected: "; resp.Rejected
    Debug.Print "ApiStatus: "; resp.apiStatus
    Debug.Print "SEFDocumentId: "; resp.sefDocumentId
    Debug.Print "ErrorCode: "; resp.errorCode
    Debug.Print "ErrorMessage: "; resp.errorMessage
    Debug.Print "RawBody: "; resp.RawBody
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub
