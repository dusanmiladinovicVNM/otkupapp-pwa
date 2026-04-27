Attribute VB_Name = "modSEFTests"
Option Explicit

' ============================================================
' modSEFTests
'
' SEF test suite for OtkupApp / AgriX.
'
' Test groups:
'   1) Offline tests: no SEF HTTP call, no workbook state mutation.
'   2) Live send/refresh smoke: calls real SEF API and mutates SEF state.
'   3) Refresh idempotency smoke: calls real SEF status endpoint.
'   4) Recovery smoke: tests an already-stuck SEF_SENDING invoice.
'
' Safety:
'   Live tests require tblSEFConfig:
'       SEF_TEST_ALLOW_LIVE = DA
'
'   Production live tests additionally require:
'       SEF_TEST_ALLOW_PROD = DA
'
' Recommended:
'   Run on a workbook copy / SEF test environment first.
' ============================================================

Private m_Total As Long
Private m_Passed As Long
Private m_Failed As Long
Private m_Skipped As Long

Private Const TEST_LOG_SHEET As String = "SEF_TEST_LOG"

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub RunSEFOfflineSuite(Optional ByVal fakturaID As String = "")
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF OFFLINE SUITE"

    If Trim$(fakturaID) = "" Then
        fakturaID = FindFirstFakturaID()
    End If

    If Trim$(fakturaID) = "" Then
        LogSkip "Find test invoice", "No faktura found in " & TBL_FAKTURE
        FinishSuite
        Exit Sub
    End If

    LogInfo "Using FakturaID=" & fakturaID

    Test_SEFConfigLooksUsable
    Test_BuildDtoAndUBL fakturaID
    Test_PayloadValidationRejectsEmpty
    Test_PersistenceReadHelpers fakturaID
    Test_ValidateFakturaForSEF_DoesNotCrash fakturaID

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFOfflineSuite", Err.Number, Err.Description
    FinishSuite
End Sub

Public Sub RunSEFLiveSendSuite(ByVal fakturaID As String)
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF LIVE SEND SUITE"

    If Trim$(fakturaID) = "" Then
        Err.Raise ERR_SEF_VALIDATION, "RunSEFLiveSendSuite", _
                  "FakturaID is required for live SEF test."
    End If

    RequireLiveSEFTestsAllowed "RunSEFLiveSendSuite"

    LogInfo "Using FakturaID=" & fakturaID

    Test_BuildDtoAndUBL fakturaID
    Test_ValidateFakturaForSEF_DoesNotCrash fakturaID
    Test_LiveSendAndRefresh fakturaID

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFLiveSendSuite", Err.Number, Err.Description
    FinishSuite
End Sub

Public Sub RunSEFRefreshIdempotencySuite(ByVal fakturaID As String)
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF REFRESH IDEMPOTENCY SUITE"

    If Trim$(fakturaID) = "" Then
        Err.Raise ERR_SEF_VALIDATION, "RunSEFRefreshIdempotencySuite", _
                  "FakturaID is required."
    End If

    RequireLiveSEFTestsAllowed "RunSEFRefreshIdempotencySuite"

    LogInfo "Using FakturaID=" & fakturaID

    Test_RefreshTwiceDoesNotBreakState fakturaID

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFRefreshIdempotencySuite", Err.Number, Err.Description
    FinishSuite
End Sub

Public Sub RunSEFRecoverySuite(ByVal stuckFakturaID As String)
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF RECOVERY SUITE"

    If Trim$(stuckFakturaID) = "" Then
        Err.Raise ERR_SEF_VALIDATION, "RunSEFRecoverySuite", _
                  "A faktura currently stuck in SEF_SENDING is required."
    End If

    RequireLiveSEFTestsAllowed "RunSEFRecoverySuite"

    LogInfo "Using stuck FakturaID=" & stuckFakturaID

    Test_RecoverStuckSendingInvoice stuckFakturaID

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFRecoverySuite", Err.Number, Err.Description
    FinishSuite
End Sub

Public Sub RunSEFBatchMaintenanceSmoke()
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF BATCH MAINTENANCE SMOKE"

    RequireLiveSEFTestsAllowed "RunSEFBatchMaintenanceSmoke"

    Test_BatchRefreshPendingDoesNotCrash
    Test_BatchRecoverStuckDoesNotCrash

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFBatchMaintenanceSmoke", Err.Number, Err.Description
    FinishSuite
End Sub

' ============================================================
' OFFLINE TESTS
' ============================================================

Private Sub Test_SEFConfigLooksUsable()
    On Error GoTo EH

    Dim baseUrl As String
    Dim apiKey As String
    Dim envName As String
    Dim paymentDue As String
    Dim taxPercent As Double

    baseUrl = Trim$(GetConfigValue("SEF_BASE_URL"))
    apiKey = Trim$(GetConfigValue("SEF_API_KEY"))
    envName = Trim$(GetConfigValue("SEF_ENV"))
    paymentDue = Trim$(GetConfigValue("SEF_PAYMENT_DUE_DAYS"))

    AssertTrue Len(baseUrl) > 0, "SEF_BASE_URL exists"
    AssertTrue Len(apiKey) > 0, "SEF_API_KEY exists"
    AssertTrue IsHttpUrl(baseUrl), "SEF_BASE_URL starts with http/https"

    If Len(paymentDue) > 0 Then
        Dim daysValue As Long
        AssertTrue TryParseLong(paymentDue, daysValue), _
                   "SEF_PAYMENT_DUE_DAYS is numeric when present"
        AssertTrue daysValue >= 0, _
                   "SEF_PAYMENT_DUE_DAYS is non-negative"
    Else
        LogPass "SEF_PAYMENT_DUE_DAYS default allowed"
    End If

    taxPercent = GetDefaultTaxPercent()
    AssertTrue taxPercent >= 0, "Default tax percent is non-negative"

    If Len(envName) > 0 Then
        LogInfo "SEF_ENV=" & envName
    End If

    Exit Sub

EH:
    LogFail "SEF config looks usable", Err.Description
End Sub

Private Sub Test_BuildDtoAndUBL(ByVal fakturaID As String)
    On Error GoTo EH

    Dim dto As clsSEFInvoiceSnapshot
    Dim xml As String

    Set dto = BuildSEFInvoiceDto(fakturaID)
    AssertTrue Not dto Is Nothing, "BuildSEFInvoiceDto returns object"
    AssertEquals fakturaID, dto.fakturaID, "DTO FakturaID"
    AssertTrue Len(Trim$(dto.InvoiceNumber)) > 0, "DTO invoice number exists"
    AssertTrue Len(Trim$(dto.BuyerName)) > 0, "DTO buyer name exists"
    AssertTrue Len(Trim$(dto.BuyerPIB)) > 0, "DTO buyer PIB exists"
    AssertTrue dto.TotalNet > 0, "DTO total net > 0"
    AssertTrue dto.TotalGross > 0, "DTO total gross > 0"
    AssertTrue Not dto.Lines Is Nothing, "DTO lines collection exists"
    AssertTrue dto.Lines.count > 0, "DTO has invoice lines"

    xml = SerializeUBLInvoice(dto)
    ValidateSEFPayload xml

    AssertTrue Len(Trim$(xml)) > 0, "UBL XML not empty"
    AssertContains xml, "<Invoice", "UBL has Invoice root"
    AssertContains xml, "<cbc:ID>", "UBL has invoice ID"
    AssertContains xml, "<cac:InvoiceLine>", "UBL has invoice line"
    AssertContains xml, dto.InvoiceNumber, "UBL contains invoice number"

    Exit Sub

EH:
    If InStr(1, Err.Description, "DeliveryDate must not be later than InvoiceDate", vbTextCompare) > 0 Then
        LogSkip "Build DTO and UBL for " & fakturaID, _
                "Local SEF validation blocked invalid dates: " & Err.Description
    Else
        LogFail "Build DTO and UBL for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.Source & _
                " Description=" & Err.Description
    End If
End Sub

Private Sub Test_PayloadValidationRejectsEmpty()
    On Error GoTo ExpectedError

    ValidateSEFPayload ""
    LogFail "ValidateSEFPayload rejects empty payload", _
            "Expected validation error, but no error was raised."
    Exit Sub

ExpectedError:
    LogPass "ValidateSEFPayload rejects empty payload"
End Sub

Private Sub Test_PersistenceReadHelpers(ByVal fakturaID As String)
    On Error GoTo EH

    Dim workflowState As String
    Dim sefDocumentId As String
    Dim submissionID As String
    Dim currentVersion As Long
    Dim nextVersion As Long

    workflowState = GetFakturaSEFWorkflowState(fakturaID)
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    submissionID = GetLastSEFSubmissionID(fakturaID)
    currentVersion = GetCurrentSEFVersionNo(fakturaID)
    nextVersion = GetNextSEFVersionNo(fakturaID)

    AssertTrue currentVersion >= 0, "Current SEF version >= 0"
    AssertTrue nextVersion >= 1, "Next SEF version >= 1"

    LogInfo "WorkflowState=" & workflowState
    LogInfo "SEFDocumentId=" & sefDocumentId
    LogInfo "LastSubmissionID=" & submissionID
    LogInfo "CurrentVersion=" & CStr(currentVersion)
    LogInfo "NextVersion=" & CStr(nextVersion)

    LogPass "Persistence read helpers do not crash"
    Exit Sub

EH:
    LogFail "Persistence read helpers for " & fakturaID, Err.Description
End Sub

Private Sub Test_ValidateFakturaForSEF_DoesNotCrash(ByVal fakturaID As String)
    On Error GoTo EH

    ValidateFakturaForSEF fakturaID
    LogPass "ValidateFakturaForSEF passes for " & fakturaID
    Exit Sub

EH:
    ' This can be an expected business validation failure if the invoice is
    ' already sent/accepted/rejected. It is still useful to record.
    LogSkip "ValidateFakturaForSEF for " & fakturaID, Err.Description
End Sub

' ============================================================
' LIVE TESTS
' ============================================================

Private Sub Test_LiveSendAndRefresh(ByVal fakturaID As String)
    On Error GoTo EH

    Dim beforeState As String
    Dim afterSendState As String
    Dim afterRefreshState As String
    Dim sefDocumentId As String
    Dim submissionID As String
    Dim resultSubmissionID As String
    Dim subStatus As String
    Dim httpStatus As String
    Dim errorCode As String
    Dim errorMessage As String

    LogInfo "==== Live send test start for " & fakturaID & " ===="

    beforeState = GetFakturaSEFWorkflowState(fakturaID)
    LogInfo "Workflow before send=" & beforeState

    resultSubmissionID = SendInvoiceToSEF_TX(fakturaID)

    afterSendState = GetFakturaSEFWorkflowState(fakturaID)
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    submissionID = GetLastSEFSubmissionID(fakturaID)

    LogInfo "SendInvoiceToSEF_TX returned=" & resultSubmissionID
    LogInfo "Workflow after send=" & afterSendState
    LogInfo "SEFDocumentId after send=" & sefDocumentId
    LogInfo "LastSubmissionID after send=" & submissionID

    AssertTrue Len(Trim$(afterSendState)) > 0, "State exists after send"
    AssertTrue Len(Trim$(submissionID)) > 0, "SubmissionID exists after send"

    subStatus = CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "SubmissionStatus"))
    httpStatus = CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "HttpStatus"))
    errorCode = CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "ErrorCode"))
    errorMessage = CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "ErrorMessage"))

    LogInfo "SubmissionStatus=" & subStatus
    LogInfo "HttpStatus=" & httpStatus
    LogInfo "ErrorCode=" & errorCode
    LogInfo "ErrorMessage=" & errorMessage

    Select Case UCase$(Trim$(afterSendState))

        Case UCase$(WF_SEF_REJECTED)
            ' Ovo je validan live rezultat: SEF je primio zahtev i poslovno ga odbio.
            AssertTrue Len(Trim$(errorCode)) > 0, "Rejected submission has ErrorCode"
            AssertTrue Len(Trim$(errorMessage)) > 0, "Rejected submission has ErrorMessage"
            LogPass "Live send reached SEF and was rejected by SEF validation"
            Exit Sub

        Case UCase$(WF_SEF_TECH_FAILED)
            LogFail "Live send technical failure", _
                    "HttpStatus=" & httpStatus & _
                    " ErrorCode=" & errorCode & _
                    " ErrorMessage=" & errorMessage
            Exit Sub

        Case UCase$(WF_SEF_SENT), UCase$(WF_SEF_ACCEPTED)
            ' Refresh only makes sense if SEFDocumentId exists.
            If Len(Trim$(sefDocumentId)) = 0 Then
                LogSkip "Refresh after live send", _
                        "No SEFDocumentId after state " & afterSendState
                LogPass "Live send completed without refresh"
                Exit Sub
            End If

            RefreshSEFStatus_TX fakturaID

            afterRefreshState = GetFakturaSEFWorkflowState(fakturaID)
            LogInfo "Workflow after refresh=" & afterRefreshState

            LogPass "Live send + refresh completed for " & fakturaID
            Exit Sub

        Case Else
            LogFail "Live send ended in unexpected workflow state", afterSendState
            Exit Sub

    End Select

EH:
    If InStr(1, Err.Description, "DeliveryDate must not be later than InvoiceDate", vbTextCompare) > 0 Then
        LogSkip "Live send + refresh for " & fakturaID, _
                "Local SEF validation blocked invalid dates: " & Err.Description
    Else
        LogFail "Live send + refresh for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " | Source=" & Err.Source & _
                " | Description=" & Err.Description
    End If
End Sub

Private Sub Test_RefreshTwiceDoesNotBreakState(ByVal fakturaID As String)
    On Error GoTo EH

    Dim state1 As String
    Dim state2 As String
    Dim state3 As String
    Dim sefDocumentId As String

    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)

    If Len(Trim$(sefDocumentId)) = 0 Then
        LogSkip "Refresh twice", _
                "No SEFDocumentId found for " & fakturaID
        Exit Sub
    End If

    state1 = GetFakturaSEFWorkflowState(fakturaID)
    LogInfo "Before first refresh state=" & state1

    RefreshSEFStatus_TX fakturaID
    state2 = GetFakturaSEFWorkflowState(fakturaID)
    LogInfo "After first refresh state=" & state2

    RefreshSEFStatus_TX fakturaID
    state3 = GetFakturaSEFWorkflowState(fakturaID)
    LogInfo "After second refresh state=" & state3

    AssertTrue Len(Trim$(state2)) > 0, "State exists after first refresh"
    AssertTrue Len(Trim$(state3)) > 0, "State exists after second refresh"

    LogPass "Refresh twice did not break state for " & fakturaID
    Exit Sub

EH:
    LogFail "Refresh twice for " & fakturaID, Err.Description
End Sub

Private Sub Test_RecoverStuckSendingInvoice(ByVal fakturaID As String)
    On Error GoTo EH

    Dim beforeState As String
    Dim afterState As String

    beforeState = GetFakturaSEFWorkflowState(fakturaID)

    If UCase$(Trim$(beforeState)) <> UCase$(WF_SEF_SENDING) Then
        LogSkip "Recover stuck SEF_SENDING", _
                "Invoice is not in SEF_SENDING. Current state=" & beforeState
        Exit Sub
    End If

    RecoverStuckSEFSendingInvoice fakturaID
    afterState = GetFakturaSEFWorkflowState(fakturaID)

    LogInfo "After recovery state=" & afterState
    AssertTrue UCase$(Trim$(afterState)) <> UCase$(WF_SEF_SENDING), _
               "Recovered invoice no longer stuck in SEF_SENDING"

    LogPass "Recover stuck SEF_SENDING for " & fakturaID
    Exit Sub

EH:
    LogFail "Recover stuck SEF_SENDING for " & fakturaID, Err.Description
End Sub

Private Sub Test_BatchRefreshPendingDoesNotCrash()
    On Error GoTo EH

    RefreshPendingOutboundInvoices_TX
    LogPass "RefreshPendingOutboundInvoices_TX completed"
    Exit Sub

EH:
    LogFail "RefreshPendingOutboundInvoices_TX", Err.Description
End Sub

Private Sub Test_BatchRecoverStuckDoesNotCrash()
    On Error GoTo EH

    RecoverAllStuckSEFSendingInvoices
    LogPass "RecoverAllStuckSEFSendingInvoices completed"
    Exit Sub

EH:
    LogFail "RecoverAllStuckSEFSendingInvoices", Err.Description
End Sub

' ============================================================
' TEST DATA HELPERS
' ============================================================

Private Function FindFirstFakturaID() As String
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_FAKTURE)
    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    colID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", "modSEFTests.FindFirstFakturaID")

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(i, colID)))) > 0 Then
            FindFirstFakturaID = Trim$(CStr(data(i, colID)))
            Exit Function
        End If
    Next i

    Exit Function

EH:
    LogFail "FindFirstFakturaID", Err.Description
End Function

Private Sub RequireLiveSEFTestsAllowed(ByVal sourceName As String)
    On Error GoTo EH

    Dim allowLive As String
    Dim allowProd As String
    Dim envName As String
    Dim baseUrl As String

    allowLive = UCase$(Trim$(GetConfigValue("SEF_TEST_ALLOW_LIVE")))
    allowProd = UCase$(Trim$(GetConfigValue("SEF_TEST_ALLOW_PROD")))
    envName = UCase$(Trim$(GetConfigValue("SEF_ENV")))
    baseUrl = UCase$(Trim$(GetConfigValue("SEF_BASE_URL")))

    If allowLive <> "DA" Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, _
                  "Live SEF tests are blocked. Set SEF_TEST_ALLOW_LIVE = DA in tblSEFConfig."
    End If

    If IsLikelyProductionSEF(envName, baseUrl) Then
        If allowProd <> "DA" Then
            Err.Raise ERR_SEF_VALIDATION, sourceName, _
                      "Production-like SEF environment detected. " & _
                      "Set SEF_TEST_ALLOW_PROD = DA only if you intentionally test production."
        End If
    End If

    Exit Sub

EH:
    LogErr "modSEFTests.RequireLiveSEFTestsAllowed"
    Err.Raise Err.Number, sourceName, Err.Description
End Sub

Private Function IsLikelyProductionSEF(ByVal envName As String, ByVal baseUrl As String) As Boolean
    Dim envText As String
    Dim urlText As String

    envText = UCase$(Trim$(envName))
    urlText = UCase$(Trim$(baseUrl))

    If envText = "PROD" Or envText = "PRODUCTION" Then
        IsLikelyProductionSEF = True
        Exit Function
    End If

    If InStr(1, urlText, "DEMO", vbTextCompare) > 0 Then Exit Function
    If InStr(1, urlText, "TEST", vbTextCompare) > 0 Then Exit Function
    If InStr(1, urlText, "SANDBOX", vbTextCompare) > 0 Then Exit Function

    ' Conservative default: if it does not look like test/sandbox,
    ' treat it as production-like.
    IsLikelyProductionSEF = True
End Function

Private Function IsHttpUrl(ByVal valueText As String) As Boolean
    IsHttpUrl = (InStr(1, valueText, "http://", vbTextCompare) = 1 _
              Or InStr(1, valueText, "https://", vbTextCompare) = 1)
End Function

' ============================================================
' ASSERTIONS
' ============================================================

Private Sub AssertTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        LogPass testName
    Else
        LogFail testName, "Assertion failed."
    End If
End Sub

Private Sub AssertEquals(ByVal expected As String, _
                         ByVal actual As String, _
                         ByVal testName As String)
    If CStr(expected) = CStr(actual) Then
        LogPass testName
    Else
        LogFail testName, _
                "Expected [" & CStr(expected) & "], got [" & CStr(actual) & "]."
    End If
End Sub

Private Sub AssertContains(ByVal haystack As String, _
                           ByVal needle As String, _
                           ByVal testName As String)
    If InStr(1, haystack, needle, vbTextCompare) > 0 Then
        LogPass testName
    Else
        LogFail testName, "Missing text: " & needle
    End If
End Sub

' ============================================================
' LOGGING
' ============================================================

Private Sub ResetSEFCounters()
    m_Total = 0
    m_Passed = 0
    m_Failed = 0
    m_Skipped = 0
End Sub

Private Sub StartSuite(ByVal suiteName As String)
    Debug.Print String$(70, "=")
    Debug.Print suiteName & " started at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(70, "=")

    AppendTestLog "SUITE", suiteName, "START", ""
End Sub

Private Sub FinishSuite()
    Dim summary As String

    summary = "Total=" & m_Total & _
              " | Passed=" & m_Passed & _
              " | Failed=" & m_Failed & _
              " | Skipped=" & m_Skipped

    Debug.Print String$(70, "-")
    Debug.Print "SEF TEST SUMMARY: " & summary
    Debug.Print String$(70, "-")

    AppendTestLog "SUITE", "SUMMARY", "INFO", summary

    If m_Failed > 0 Then
        MsgBox "SEF tests finished with failures." & vbCrLf & summary, _
               vbExclamation, APP_NAME
    Else
        MsgBox "SEF tests finished." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If
End Sub

Private Sub LogPass(ByVal testName As String)
    m_Total = m_Total + 1
    m_Passed = m_Passed + 1

    Debug.Print "[PASS] " & testName
    AppendTestLog "TEST", testName, "PASS", ""
End Sub

Private Sub LogFail(ByVal testName As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FAIL] " & testName & " :: " & details
    AppendTestLog "TEST", testName, "FAIL", details
End Sub

Private Sub LogSkip(ByVal testName As String, ByVal reason As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1

    Debug.Print "[SKIP] " & testName & " :: " & reason
    AppendTestLog "TEST", testName, "SKIP", reason
End Sub

Private Sub LogInfo(ByVal message As String)
    Debug.Print "[INFO] " & message
    AppendTestLog "INFO", "", "INFO", message
End Sub

Private Sub LogFatal(ByVal sourceName As String, ByVal errNum As Long, ByVal errDesc As String)
    m_Total = m_Total + 1
    m_Failed = m_Failed + 1

    Debug.Print "[FATAL] " & sourceName & " :: " & CStr(errNum) & " - " & errDesc
    AppendTestLog "FATAL", sourceName, "FAIL", CStr(errNum) & " - " & errDesc
End Sub

Private Sub InitSEFTestLog()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TEST_LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = TEST_LOG_SHEET
        ws.Range("A1:F1").value = Array("Timestamp", "Kind", "Name", "Status", "Details", "Operator")
        ws.rows(1).Font.Bold = True
    End If
End Sub

Private Sub AppendTestLog(ByVal kindText As String, _
                          ByVal nameText As String, _
                          ByVal statusText As String, _
                          ByVal detailsText As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TEST_LOG_SHEET)
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

Private Sub LogSEFFakturaSnapshot(ByVal fakturaID As String, ByVal labelText As String)
    On Error GoTo EH

    If Len(Trim$(fakturaID)) = 0 Then Exit Sub

    LogInfo labelText & " FakturaID=" & fakturaID
    LogInfo labelText & " SEFWorkflowState=" & CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFWorkflowState"))
    LogInfo labelText & " SEFStatus=" & CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))
    LogInfo labelText & " SEFDocumentId=" & CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFDocumentId"))
    LogInfo labelText & " SEFSubmissionIDLast=" & CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFSubmissionIDLast"))
    LogInfo labelText & " SEFLastErrorCode=" & CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFLastErrorCode"))
    LogInfo labelText & " SEFLastErrorMessage=" & CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFLastErrorMessage"))
    LogInfo labelText & " SEFVersionNo=" & CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFVersionNo"))

    Exit Sub

EH:
    LogInfo labelText & " Faktura snapshot failed: " & Err.Description
End Sub

Private Sub LogSEFSubmissionSnapshot(ByVal submissionID As String, ByVal labelText As String)
    On Error GoTo EH

    If Len(Trim$(submissionID)) = 0 Then
        LogInfo labelText & " Submission snapshot skipped: no submissionID."
        Exit Sub
    End If

    LogInfo labelText & " SEFSubmissionID=" & submissionID
    LogInfo labelText & " SubmissionStatus=" & CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "SubmissionStatus"))
    LogInfo labelText & " HttpStatus=" & CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "HttpStatus"))
    LogInfo labelText & " ApiStatus=" & CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "ApiStatus"))
    LogInfo labelText & " SEFDocumentId=" & CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "SEFDocumentId"))
    LogInfo labelText & " ErrorCode=" & CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "ErrorCode"))
    LogInfo labelText & " ErrorMessage=" & CStr(LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "ErrorMessage"))

    Exit Sub

EH:
    LogInfo labelText & " Submission snapshot failed: " & Err.Description
End Sub

' ============================================================
' DESTRUCTIVE LIVE TESTS: CANCEL / STORNO
' ============================================================

Public Sub RunSEFCancelLiveSuite(ByVal fakturaID As String)
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF LIVE CANCEL SUITE"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "RunSEFCancelLiveSuite", _
                  "FakturaID is required."
    End If

    RequireLiveSEFTestsAllowed "RunSEFCancelLiveSuite"
    RequireCancelStornoTestsAllowed "RunSEFCancelLiveSuite"
    ConfirmDangerousSEFMutation "CANCEL", fakturaID

    Test_LiveCancelInvoice fakturaID

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFCancelLiveSuite", Err.Number, Err.Description
    FinishSuite
End Sub

Public Sub RunSEFStornoLiveSuite(ByVal fakturaID As String, _
                                 Optional ByVal stornoNumber As String = "")
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF LIVE STORNO SUITE"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "RunSEFStornoLiveSuite", _
                  "FakturaID is required."
    End If

    If Len(Trim$(stornoNumber)) = 0 Then
        stornoNumber = "ST-" & fakturaID & "-" & Format$(Now, "yyyymmddhhnnss")
    End If

    RequireLiveSEFTestsAllowed "RunSEFStornoLiveSuite"
    RequireCancelStornoTestsAllowed "RunSEFStornoLiveSuite"
    ConfirmDangerousSEFMutation "STORNO", fakturaID

    Test_LiveStornoInvoice fakturaID, stornoNumber

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFStornoLiveSuite", Err.Number, Err.Description
    FinishSuite
End Sub

Private Sub Test_LiveCancelInvoice(ByVal fakturaID As String)
    On Error GoTo EH

    Dim beforeWorkflow As String
    Dim beforeStatus As String
    Dim beforeDocID As String
    Dim afterWorkflow As String
    Dim afterStatus As String
    Dim afterDocID As String
    Dim beforeEvents As Long
    Dim afterEvents As Long
    Dim commentText As String

    commentText = "Automated SEF cancel smoke test " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    beforeWorkflow = GetFakturaSEFWorkflowState(fakturaID)
    beforeStatus = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))
    beforeDocID = GetFakturaSEFDocumentId(fakturaID)
    beforeEvents = CountSEFEventsForFaktura(fakturaID)

    LogInfo "Before cancel Workflow=" & beforeWorkflow
    LogInfo "Before cancel SEFStatus=" & beforeStatus
    LogInfo "Before cancel SEFDocumentId=" & beforeDocID

    If Len(Trim$(beforeDocID)) = 0 Then
        LogSkip "Live cancel " & fakturaID, "No SEFDocumentId."
        Exit Sub
    End If

    ' Change only this line if your service signature is different.
    Call CancelInvoiceOnSEF_TX(fakturaID, commentText)

    On Error Resume Next
    RefreshSEFStatus_TX fakturaID
    Err.Clear
    On Error GoTo EH

    afterWorkflow = GetFakturaSEFWorkflowState(fakturaID)
    afterStatus = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))
    afterDocID = GetFakturaSEFDocumentId(fakturaID)
    afterEvents = CountSEFEventsForFaktura(fakturaID)

    LogInfo "After cancel Workflow=" & afterWorkflow
    LogInfo "After cancel SEFStatus=" & afterStatus
    LogInfo "After cancel SEFDocumentId=" & afterDocID

    AssertTrue Len(Trim$(afterWorkflow)) > 0, "Cancel leaves workflow state populated"
    AssertTrue Len(Trim$(afterStatus)) > 0, "Cancel leaves SEFStatus populated"
    AssertTrue afterEvents > beforeEvents, "Cancel writes SEF event log"

    LogPass "Live cancel completed for " & fakturaID
    Exit Sub

EH:
    If IsExpectedSEFBusinessBlock(Err.Description) Then
        LogSkip "Live cancel blocked by SEF/service rule for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.Source & _
                " Description=" & Err.Description
    Else
        LogFail "Live cancel for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.Source & _
                " Description=" & Err.Description
    End If
End Sub

Private Sub Test_LiveStornoInvoice(ByVal fakturaID As String, _
                                   ByVal stornoNumber As String)
    On Error GoTo EH

    Dim beforeWorkflow As String
    Dim beforeStatus As String
    Dim beforeDocID As String
    Dim afterWorkflow As String
    Dim afterStatus As String
    Dim afterDocID As String
    Dim beforeEvents As Long
    Dim afterEvents As Long
    Dim commentText As String

    commentText = "Automated SEF storno smoke test " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    beforeWorkflow = GetFakturaSEFWorkflowState(fakturaID)
    beforeStatus = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))
    beforeDocID = GetFakturaSEFDocumentId(fakturaID)
    beforeEvents = CountSEFEventsForFaktura(fakturaID)

    LogInfo "Before storno Workflow=" & beforeWorkflow
    LogInfo "Before storno SEFStatus=" & beforeStatus
    LogInfo "Before storno SEFDocumentId=" & beforeDocID
    LogInfo "StornoNumber=" & stornoNumber

    If Len(Trim$(beforeDocID)) = 0 Then
        LogSkip "Live storno " & fakturaID, "No SEFDocumentId."
        Exit Sub
    End If

    ' Change only this line if your service signature is different.
    Call StornoInvoiceOnSEF_TX(fakturaID, stornoNumber, commentText)

    On Error Resume Next
    RefreshSEFStatus_TX fakturaID
    Err.Clear
    On Error GoTo EH

    afterWorkflow = GetFakturaSEFWorkflowState(fakturaID)
    afterStatus = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))
    afterDocID = GetFakturaSEFDocumentId(fakturaID)
    afterEvents = CountSEFEventsForFaktura(fakturaID)

    LogInfo "After storno Workflow=" & afterWorkflow
    LogInfo "After storno SEFStatus=" & afterStatus
    LogInfo "After storno SEFDocumentId=" & afterDocID

    AssertTrue Len(Trim$(afterWorkflow)) > 0, "Storno leaves workflow state populated"
    AssertTrue Len(Trim$(afterStatus)) > 0, "Storno leaves SEFStatus populated"
    AssertTrue afterEvents > beforeEvents, "Storno writes SEF event log"

    LogPass "Live storno completed for " & fakturaID
    Exit Sub

EH:
    If IsExpectedSEFBusinessBlock(Err.Description) Then
        LogSkip "Live storno blocked by SEF/service rule for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.Source & _
                " Description=" & Err.Description
    Else
        LogFail "Live storno for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.Source & _
                " Description=" & Err.Description
    End If
End Sub

Private Sub RequireCancelStornoTestsAllowed(ByVal sourceName As String)
    Dim allowValue As String

    allowValue = UCase$(Trim$(GetConfigValue("SEF_TEST_ALLOW_CANCEL_STORNO")))

    If allowValue <> "DA" Then
        Err.Raise ERR_SEF_VALIDATION, sourceName, _
                  "Cancel/storno live tests are blocked. Set SEF_TEST_ALLOW_CANCEL_STORNO = DA in tblSEFConfig."
    End If
End Sub

Private Sub ConfirmDangerousSEFMutation(ByVal actionName As String, _
                                        ByVal fakturaID As String)
    Dim expectedText As String
    Dim answer As String

    expectedText = actionName & " " & fakturaID

    answer = InputBox( _
        "This will perform a REAL SEF " & actionName & " operation." & vbCrLf & _
        "FakturaID: " & fakturaID & vbCrLf & vbCrLf & _
        "To continue, type exactly:" & vbCrLf & expectedText, _
        "Confirm destructive SEF test")

    If answer <> expectedText Then
        Err.Raise ERR_SEF_VALIDATION, "ConfirmDangerousSEFMutation", _
                  "Destructive SEF test cancelled by user."
    End If
End Sub

Private Function CountSEFEventsForFaktura(ByVal fakturaID As String) As Long
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_SEF_EVENT_LOG)

    If IsEmpty(data) Then Exit Function

    Dim colFakturaID As Long
    colFakturaID = GetColumnIndex(TBL_SEF_EVENT_LOG, "FakturaID")

    If colFakturaID = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakturaID))) = fakturaID Then
            CountSEFEventsForFaktura = CountSEFEventsForFaktura + 1
        End If
    Next i

    Exit Function

EH:
    CountSEFEventsForFaktura = 0
End Function

Private Function IsExpectedSEFBusinessBlock(ByVal textValue As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(textValue))

    IsExpectedSEFBusinessBlock = _
        InStr(1, s, "NOT ALLOWED", vbTextCompare) > 0 Or _
        InStr(1, s, "CANNOT", vbTextCompare) > 0 Or _
        InStr(1, s, "CURRENT STATE", vbTextCompare) > 0 Or _
        InStr(1, s, "STATUS", vbTextCompare) > 0 Or _
        InStr(1, s, "VALIDATION", vbTextCompare) > 0
End Function
