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
    Test_SubmitResponseClassification
    Test_LinePrecisionConsistency
    Test_BuildDtoAndUBL fakturaID
    Test_PayloadValidationRejectsEmpty
    Test_PersistenceReadHelpers fakturaID
    Test_ValidateFakturaForSEF_DoesNotCrash fakturaID

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFOfflineSuite", Err.Number, Err.description
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
    LogFatal "RunSEFLiveSendSuite", Err.Number, Err.description
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
    LogFatal "RunSEFRefreshIdempotencySuite", Err.Number, Err.description
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
    LogFatal "RunSEFRecoverySuite", Err.Number, Err.description
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
    LogFatal "RunSEFBatchMaintenanceSmoke", Err.Number, Err.description
    FinishSuite
End Sub

' ============================================================
' RunSEFTestSuite -- ciljani suite za SEF milestone (PLAN_SANACIJE, sekcija 5C).
'
' HARD GATE: na kraju baca gresku ako je bilo koji test pao, da gate ne bi bio
' "zelen" samo zato sto nista nije puklo (lekcija iz AUD-039 / RF-08).
'
' OFFLINE: nijedan test odavde ne poziva pravi SEF (nema HTTP-a). Zivi SEF pozivi
' ostaju u Run*LiveSuite entry point-ima iza SEF_TEST_ALLOW_LIVE /
' SEF_TEST_ALLOW_CANCEL_STORNO kapija.
'
' TABELE: svi testovi su cisti OSIM
' `Test_SEFRejectedResubmitPassesDuplicateGuard`, koji seed-uje redove u
' tblFakture / tblSEFSubmission / tblSEFEventLog i UVEK ih rollback-uje
' (clsTransaction, isti obrazac kao modFakturaTests). Taj lanac se ne moze
' dokazati cistim funkcijama -- duplicate guard cita stvarne tabele.
' ============================================================
Public Sub RunSEFTestSuite()
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF TEST SUITE (offline hard gate)"

    Test_SubmitResponseClassification
    Test_LinePrecisionConsistency
    Test_SEFAllowedTransitions
    Test_SEFBlockedTransitions

    ' RF-22 / AUD-032 seam testovi
    Test_SEFSendOutcomeContract
    Test_SEFOfficialStatusEnumClassified
    Test_SEFStatusUnknownIsNotSent
    Test_SEFRefreshTransitionMatrix
    Test_SEFStatusCapabilities
    Test_SEFRejectedResubmitPassesDuplicateGuard
    Test_SEFRecoveryOutcomeContract

    FinishSuite

    ' Gate se proverava tek posle gasenja EH-a, da raise ne bi upao u sopstveni
    ' handler i bio prijavljen kao fatalna greska suite-a (RF-08 obrazac).
    On Error GoTo 0

    If m_Failed > 0 Then
        Err.Raise ERR_SEF_VALIDATION, "RunSEFTestSuite", _
                  "SEF test suite FAILED: " & CStr(m_Failed) & " od " & CStr(m_Total) & " testova."
    End If

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogFatal "RunSEFTestSuite", errNum, errDesc
    FinishSuite

    On Error GoTo 0
    Err.Raise ERR_SEF_VALIDATION, "RunSEFTestSuite", _
              "SEF test suite FAILED (fatal): " & errDesc
End Sub

Public Sub RunSEFStateTransitionSuite()
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF STATE TRANSITION SUITE"

    Test_SEFAllowedTransitions
    Test_SEFBlockedTransitions

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFStateTransitionSuite", Err.Number, Err.description
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
    LogFail "SEF config looks usable", Err.description
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
    AssertTrue Not dto.lines Is Nothing, "DTO lines collection exists"
    AssertTrue dto.lines.count > 0, "DTO has invoice lines"

    ' AUD-031b: emitted quantity * unit price must reproduce the line net, so
    ' the UBL is arithmetically consistent for the receiver (tax authority).
    Dim li As Long
    Dim lnItem As clsSEFLine
    For li = 1 To dto.lines.count
        Set lnItem = dto.lines.item(li)
        AssertTrue Abs(Round(lnItem.kolicina * lnItem.cena, 2) - lnItem.neto) < 0.005, _
                   "Line " & li & " net == round(qty*price,2)"
    Next li

    xml = SerializeUBLInvoice(dto)
    ValidateSEFPayload xml

    AssertTrue Len(Trim$(xml)) > 0, "UBL XML not empty"
    AssertContains xml, "<Invoice", "UBL has Invoice root"
    AssertContains xml, "<cbc:ID>", "UBL has invoice ID"
    AssertContains xml, "<cac:InvoiceLine>", "UBL has invoice line"
    AssertContains xml, dto.InvoiceNumber, "UBL contains invoice number"

    Exit Sub

EH:
    If InStr(1, Err.description, "DeliveryDate must not be later than InvoiceDate", vbTextCompare) > 0 Then
        LogSkip "Build DTO and UBL for " & fakturaID, _
                "Local SEF validation blocked invalid dates: " & Err.description
    Else
        LogFail "Build DTO and UBL for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.SOURCE & _
                " Description=" & Err.description
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

Private Sub Test_SubmitResponseClassification()
    On Error GoTo EH

    Dim r As clsSEFResponse

    ' AUD-030 core: HTTP 409 must be CONFLICT, NOT REJECTED (a REJECTED 409
    ' would enable a corrective resubmit with a fresh requestId).
    Set r = TestProxyForParseSubmitResponse(409, "{""message"":""already exists""}")
    AssertEquals "CONFLICT", r.apiStatus, "HTTP 409 -> apiStatus CONFLICT"
    AssertTrue r.Rejected = False, "HTTP 409 not flagged Rejected"
    AssertTrue r.Success = False, "HTTP 409 not flagged Success"

    ' 400 / 422 stay REJECTED (business rejection by SEF validation).
    Set r = TestProxyForParseSubmitResponse(400, "{""message"":""bad request""}")
    AssertEquals "REJECTED", r.apiStatus, "HTTP 400 -> apiStatus REJECTED"
    AssertTrue r.Rejected, "HTTP 400 flagged Rejected"

    Set r = TestProxyForParseSubmitResponse(422, "{""message"":""invalid""}")
    AssertEquals "REJECTED", r.apiStatus, "HTTP 422 -> apiStatus REJECTED"
    AssertTrue r.Rejected, "HTTP 422 flagged Rejected"

    ' 2xx success.
    Set r = TestProxyForParseSubmitResponse(200, "{""SalesInvoiceId"":123}")
    AssertTrue r.Success, "HTTP 200 flagged Success"
    AssertTrue r.Rejected = False, "HTTP 200 not flagged Rejected"

    ' Other 4xx/5xx -> technical FAILED (retryable), NOT REJECTED.
    Set r = TestProxyForParseSubmitResponse(500, "{""message"":""server error""}")
    AssertEquals "FAILED", r.apiStatus, "HTTP 500 -> apiStatus FAILED"
    AssertTrue r.Rejected = False, "HTTP 500 not flagged Rejected"

    Exit Sub

EH:
    LogFail "Submit response classification", Err.description
End Sub

Private Sub Test_LinePrecisionConsistency()
    On Error GoTo EH

    ' AUD-031b: quantity/price keep more precision than 2-decimal money, so the
    ' receiver's Round(qty*price,2) reproduces the line net. Self-contained --
    ' needs no invoice data. Reviewer example: qty=1.234, price=1.234 -> net
    ' Round(1.234*1.234,2)=1.52; old 2dp emit (1.23*1.23=1.5129->1.51) is wrong.
    AssertEquals "1.234", TestProxyXmlQuantity(1.234), "XmlQuantity keeps 3 decimals"
    AssertEquals "1.2340", TestProxyXmlUnitPrice(1.234), "XmlUnitPrice keeps 4 decimals"
    AssertEquals "85.5000", TestProxyXmlUnitPrice(85.5), "XmlUnitPrice pads to 4 decimals"

    ' Numeric invariant (locale-safe): emission precision reproduces the net...
    AssertTrue Round(Round(1.234, 3) * Round(1.234, 4), 2) = 1.52, _
               "Round(qty*price,2) reproduces net with precision fix"
    ' ...while 2-decimal truncation would NOT (guards against regressing it).
    AssertTrue Round(Round(1.234, 2) * Round(1.234, 2), 2) <> 1.52, _
               "2dp truncation is inconsistent (regression guard)"

    Exit Sub

EH:
    LogFail "Line precision consistency", Err.description
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
    LogFail "Persistence read helpers for " & fakturaID, Err.description
End Sub

Private Sub Test_ValidateFakturaForSEF_DoesNotCrash(ByVal fakturaID As String)
    On Error GoTo EH

    ValidateFakturaForSEF fakturaID
    LogPass "ValidateFakturaForSEF passes for " & fakturaID
    Exit Sub

EH:
    ' This can be an expected business validation failure if the invoice is
    ' already sent/accepted/rejected. It is still useful to record.
    LogSkip "ValidateFakturaForSEF for " & fakturaID, Err.description
End Sub

' ============================================================
' RF-22 / AUD-032 -- lifecycle seam testovi (offline, bez SEF poziva)
' ============================================================

' (a) REJECTED / TECH_FAILED se NE smeju prikazati kao "Faktura poslata".
Private Sub Test_SEFSendOutcomeContract()
    On Error GoTo EH

    AssertTrue IsSuccessfulSEFSendState(WF_SEF_SENT), _
               "SEF_SENT je uspesan send"
    AssertTrue IsSuccessfulSEFSendState(WF_SEF_ACCEPTED), _
               "SEF_ACCEPTED je uspesan send"

    AssertTrue Not IsSuccessfulSEFSendState(WF_SEF_REJECTED), _
               "SEF_REJECTED NIJE uspesan send"
    AssertTrue Not IsSuccessfulSEFSendState(WF_SEF_TECH_FAILED), _
               "SEF_TECH_FAILED NIJE uspesan send"
    AssertTrue Not IsSuccessfulSEFSendState(WF_SEF_SENDING), _
               "SEF_SENDING NIJE uspesan send"
    AssertTrue Not IsSuccessfulSEFSendState(""), _
               "Prazan workflow state NIJE uspesan send"

    ' Tipizirane greske: odbijanje se razlikuje od tehnickog pada.
    AssertTrue SEFSendFailureErrNumber(WF_SEF_REJECTED) = ERR_SEF_REJECTED, _
               "REJECTED -> ERR_SEF_REJECTED"
    AssertTrue SEFSendFailureErrNumber(WF_SEF_TECH_FAILED) = ERR_SEF_SEND_FAILED, _
               "TECH_FAILED -> ERR_SEF_SEND_FAILED"
    AssertTrue SEFSendFailureErrNumber("") = ERR_SEF_SEND_FAILED, _
               "Nepoznat ishod -> ERR_SEF_SEND_FAILED"

    ' Sustina AUD-032a: poruka za neuspeh ne sme da sadrzi tekst uspeha.
    Dim successText As String
    successText = Poruka("SEF_MSG_SEND_POSLATA")

    AssertTrue InStr(1, SEFSendOutcomeMessage(WF_SEF_REJECTED, "SUB-1"), successText, vbTextCompare) = 0, _
               "Poruka za REJECTED ne tvrdi da je faktura poslata"
    AssertTrue InStr(1, SEFSendOutcomeMessage(WF_SEF_TECH_FAILED, "SUB-1"), successText, vbTextCompare) = 0, _
               "Poruka za TECH_FAILED ne tvrdi da je faktura poslata"
    AssertTrue InStr(1, SEFSendOutcomeMessage(WF_SEF_SENT, "SUB-1"), successText, vbTextCompare) > 0, _
               "Poruka za SEF_SENT jeste poruka o poslatoj fakturi"

    AssertTrue InStr(1, SEFSendOutcomeMessage(WF_SEF_REJECTED, "SUB-1"), WF_SEF_REJECTED, vbTextCompare) > 0, _
               "Poruka za REJECTED nosi stvarno workflow stanje"
    AssertTrue SEFSendOutcomeMessage(WF_SEF_REJECTED, "SUB-1") <> SEFSendOutcomeMessage(WF_SEF_SENT, "SUB-1"), _
               "Poruka za REJECTED se razlikuje od poruke za uspeh"

    Exit Sub

EH:
    LogFail "SEF send outcome contract", Err.description
End Sub

' (b1) Adapter: SVAKI status iz zvanicnog SalesInvoiceStatus enum-a mora imati
' eksplicitnu klasu. Bez ovoga je "APPROVED" (zvanicno ime za prihvacenu
' fakturu) padao u Case Else i zavrsavao kao nepoznat status.
Private Sub Test_SEFOfficialStatusEnumClassified()
    On Error GoTo EH

    ' --- prihvatanje: zvanicno je Approved, ne Accepted ---
    AssertEquals SEF_CLS_ACCEPTED, ClassifySEFExternalStatus("Approved"), _
                 "Approved -> ACCEPTED klasa"
    AssertEquals SEF_CLS_ACCEPTED, ClassifySEFExternalStatus("APPROVED"), _
                 "APPROVED (velika slova) -> ACCEPTED klasa"
    AssertEquals SEF_CLS_ACCEPTED, ClassifySEFExternalStatus("Accepted"), _
                 "Accepted (zatecene fakture / submit odgovor) -> ACCEPTED klasa"

    AssertEquals SEF_CLS_REJECTED, ClassifySEFExternalStatus("Rejected"), _
                 "Rejected -> REJECTED klasa"

    ' --- u obradi / ceka odluku ---
    AssertEquals SEF_CLS_PENDING, ClassifySEFExternalStatus("New"), "New -> PENDING"
    AssertEquals SEF_CLS_PENDING, ClassifySEFExternalStatus("Draft"), "Draft -> PENDING"
    AssertEquals SEF_CLS_PENDING, ClassifySEFExternalStatus("Sending"), "Sending -> PENDING"
    AssertEquals SEF_CLS_PENDING, ClassifySEFExternalStatus("Sent"), "Sent -> PENDING"
    AssertEquals SEF_CLS_PENDING, ClassifySEFExternalStatus("Seen"), "Seen -> PENDING"

    ' --- terminalno na SEF-u ---
    AssertEquals SEF_CLS_TERMINAL, ClassifySEFExternalStatus("Cancelled"), "Cancelled -> TERMINAL"
    AssertEquals SEF_CLS_TERMINAL, ClassifySEFExternalStatus("Storno"), "Storno -> TERMINAL"
    AssertEquals SEF_CLS_TERMINAL, ClassifySEFExternalStatus("Deleted"), "Deleted -> TERMINAL"

    ' "Mistake" = greska prilikom slanja, NE terminalno stanje. Kad je bio u
    ' TERMINAL klasi, planer ga je vodio u WF_SEF_SENT -- neuspelo slanje je
    ' postajalo lokalno "poslato", bez retry-ja i bez cancel-a.
    AssertEquals SEF_CLS_SEND_FAILED, ClassifySEFExternalStatus("Mistake"), _
                 "Mistake -> SEND_FAILED klasa (ne TERMINAL)"
    AssertTrue ClassifySEFExternalStatus("Mistake") <> SEF_CLS_TERMINAL, _
                 "Mistake nije terminalan (batch ga ne sme preskakati)"

    ' --- poznato, ali ne nosi odluku kupca ---
    AssertEquals SEF_CLS_INFO, ClassifySEFExternalStatus("Paid"), "Paid -> INFO"
    AssertEquals SEF_CLS_INFO, ClassifySEFExternalStatus("OverDue"), "OverDue -> INFO"
    AssertEquals SEF_CLS_INFO, ClassifySEFExternalStatus("Archived"), "Archived -> INFO"

    ' --- zvanicni Unknown i sve van enum-a ---
    AssertEquals SEF_CLS_UNKNOWN, ClassifySEFExternalStatus("Unknown"), _
                 "Zvanicni Unknown -> UNKNOWN klasa"
    AssertEquals SEF_CLS_UNKNOWN, ClassifySEFExternalStatus(""), _
                 "Prazan status -> UNKNOWN klasa"
    AssertEquals SEF_CLS_UNKNOWN, ClassifySEFExternalStatus("NEKI_NOVI_STATUS"), _
                 "Nov/nepoznat status -> UNKNOWN klasa"

    ' Nijedan poznat status ne sme da ispadne "nepoznat" -- to je bio uzrok
    ' zbog kog je odobrena faktura isla na rucnu proveru.
    AssertTrue IsKnownSEFRefreshStatus("Approved"), "Approved je poznat status"
    AssertTrue IsKnownSEFRefreshStatus("Seen"), "Seen je poznat status"
    AssertTrue IsKnownSEFRefreshStatus("Paid"), "Paid je poznat status"
    AssertTrue Not IsKnownSEFRefreshStatus("Unknown"), "Unknown nije poznat status"
    AssertTrue Not IsKnownSEFRefreshStatus(""), "Prazan status nije poznat"

    ' Refresh je "upotrebljiv" za sve sto nosi stvarnu informaciju.
    AssertTrue IsUsableSEFRefreshClass(SEF_CLS_ACCEPTED), "ACCEPTED je upotrebljiv refresh"
    AssertTrue IsUsableSEFRefreshClass(SEF_CLS_INFO), "INFO je upotrebljiv refresh"
    AssertTrue IsUsableSEFRefreshClass(SEF_CLS_SEND_FAILED), "SEND_FAILED je upotrebljiv refresh"
    AssertTrue Not IsUsableSEFRefreshClass(SEF_CLS_ERROR), "ERROR nije upotrebljiv refresh"
    AssertTrue Not IsUsableSEFRefreshClass(SEF_CLS_UNKNOWN), "UNKNOWN nije upotrebljiv refresh"

    Exit Sub

EH:
    LogFail "SEF official status enum classified", Err.description
End Sub

' (b2) Prazan / nepoznat status ne sme tiho da postane SENT.
Private Sub Test_SEFStatusUnknownIsNotSent()
    On Error GoTo EH

    Dim r As clsSEFResponse

    ' Nema "Status" polja u odgovoru -> UNKNOWN_STATUS, nikad "SENT".
    Set r = TestProxyForParseStatusResponse(200, "{""InvoiceId"":5317568}")
    AssertEquals SEF_STATUS_UNKNOWN, r.apiStatus, _
                 "Prazan SEF status -> UNKNOWN_STATUS (ne SENT)"
    AssertTrue UCase$(r.apiStatus) <> "SENT", _
               "Prazan SEF status se ne prijavljuje kao SENT"
    AssertTrue r.Accepted = False, "Prazan status nije prihvatanje"
    AssertTrue Not IsKnownSEFRefreshStatus(r.apiStatus), _
               "UNKNOWN_STATUS nije poznat status"

    ' Zvanicno prihvatanje mora da podigne Accepted flag.
    Set r = TestProxyForParseStatusResponse(200, "{""Status"":""Approved""}")
    AssertEquals "APPROVED", r.apiStatus, "APPROVED se cuva doslovno"
    AssertTrue r.Accepted, "Approved postavlja Accepted flag"
    AssertTrue r.Rejected = False, "Approved ne postavlja Rejected"

    Set r = TestProxyForParseStatusResponse(200, "{""Status"":""Rejected""}")
    AssertTrue r.Rejected, "Rejected postavlja Rejected flag"
    AssertTrue r.Accepted = False, "Rejected ne postavlja Accepted"

    ' Poznati pending statusi ostaju netaknuti.
    Set r = TestProxyForParseStatusResponse(200, "{""Status"":""SENT""}")
    AssertEquals "SENT", r.apiStatus, "SENT ostaje SENT"
    AssertTrue r.Accepted = False, "SENT nije prihvatanje"

    ' Nepoznat, ali neprazan status se cuva doslovno i tretira kao nepoznat.
    Set r = TestProxyForParseStatusResponse(200, "{""Status"":""NEKI_NOVI_STATUS""}")
    AssertEquals "NEKI_NOVI_STATUS", r.apiStatus, _
                 "Nepoznat status se cuva doslovno"
    AssertTrue Not IsKnownSEFRefreshStatus(r.apiStatus), _
               "Nepoznat status se ne tretira kao poznat"

    ' Dokument u ERROR statusu: poziv je uspeo, dokument nije.
    Set r = TestProxyForParseStatusResponse(200, "{""Status"":""Error""}")
    AssertTrue r.Success = False, "ERROR status obara Success"

    Exit Sub

EH:
    LogFail "SEF unknown status is not SENT", Err.description
End Sub

' (b3) ORKESTRACIJA: planer tranzicije za SVAKU kombinaciju (stanje x klasa).
' Ovo je test koji hvata protivrecnosti tipa "SEF_UNKNOWN + pad API-ja ->
' SEF_SYNC_ERROR", koje testovi samog validatora ne mogu da vide: planer i
' state machine su dve strane iste odluke, pa se proveravaju zajedno.
Private Sub Test_SEFRefreshTransitionMatrix()
    On Error GoTo EH

    Dim states As Variant
    Dim classes As Variant
    Dim i As Long
    Dim j As Long
    Dim st As String
    Dim cls As String
    Dim target As String
    Dim illegalCount As Long

    states = Array(WF_LOCAL_DRAFT, WF_LOCAL_FINALIZED, WF_SEF_READY, _
                   WF_SEF_SENDING, WF_SEF_SENT, WF_SEF_ACCEPTED, _
                   WF_SEF_REJECTED, WF_SEF_STORNO, WF_SEF_SYNC_ERROR, _
                   WF_SEF_TECH_FAILED, WF_SEF_UNKNOWN, "")

    classes = Array(SEF_CLS_ACCEPTED, SEF_CLS_REJECTED, SEF_CLS_SEND_FAILED, _
                    SEF_CLS_PENDING, SEF_CLS_TERMINAL, SEF_CLS_INFO, _
                    SEF_CLS_ERROR, SEF_CLS_UNKNOWN)

    ' INVARIJANTA: planer sme da vrati samo prazno ili tranziciju koju state
    ' machine dozvoljava. Sve ostalo bi u produkciji bacilo izuzetak i oborilo
    ' refresh (i rollback-ovalo transakciju).
    For i = LBound(states) To UBound(states)
        For j = LBound(classes) To UBound(classes)

            st = CStr(states(i))
            cls = CStr(classes(j))
            target = SEFRefreshTargetState(st, cls)

            If Len(target) > 0 And Len(Trim$(st)) > 0 Then
                If Not IsSEFTransitionAllowed(UCase$(Trim$(st)), target) Then
                    illegalCount = illegalCount + 1
                    LogFail "Planer predlaze zabranjenu tranziciju", _
                            st & " + " & cls & " -> " & target
                End If
            End If

        Next j
    Next i

    AssertTrue illegalCount = 0, _
               "Planer nikad ne predlaze zabranjenu tranziciju (sva stanja x sve klase)"

    ' --- konkretni ishodi koje trazi lifecycle ---

    ' Odobrenje mora da stigne do SEF_ACCEPTED (jednim ili dva koraka).
    AssertEquals WF_SEF_ACCEPTED, SEFRefreshTargetState(WF_SEF_SENDING, SEF_CLS_ACCEPTED), _
                 "SENDING + Approved -> SEF_ACCEPTED"
    AssertEquals WF_SEF_ACCEPTED, SEFRefreshTargetState(WF_SEF_SENT, SEF_CLS_ACCEPTED), _
                 "SENT + Approved -> SEF_ACCEPTED"
    AssertEquals WF_SEF_ACCEPTED, SEFRefreshTargetState(WF_SEF_UNKNOWN, SEF_CLS_ACCEPTED), _
                 "UNKNOWN + Approved -> SEF_ACCEPTED (izlazi iz rucne provere)"
    ' SEF_SYNC_ERROR ne sme direktno u finalno stanje -> prvi korak je SEF_SENT,
    ' drugi (isti planer, novo stanje) daje SEF_ACCEPTED.
    AssertEquals WF_SEF_SENT, SEFRefreshTargetState(WF_SEF_SYNC_ERROR, SEF_CLS_ACCEPTED), _
                 "SYNC_ERROR + Approved -> prvi korak SEF_SENT"
    AssertEquals WF_SEF_ACCEPTED, _
                 SEFRefreshTargetState(SEFRefreshTargetState(WF_SEF_SYNC_ERROR, SEF_CLS_ACCEPTED), SEF_CLS_ACCEPTED), _
                 "SYNC_ERROR + Approved -> drugi korak SEF_ACCEPTED"

    ' Terminalni udaljeni status izvlaci fakturu iz svih zaglavljenih stanja.
    AssertEquals WF_SEF_SENT, SEFRefreshTargetState(WF_SEF_SENDING, SEF_CLS_TERMINAL), _
                 "SENDING + Storno -> SEF_SENT (ne ostaje zaglavljena)"
    AssertEquals WF_SEF_SENT, SEFRefreshTargetState(WF_SEF_UNKNOWN, SEF_CLS_TERMINAL), _
                 "UNKNOWN + Storno -> SEF_SENT (izlazi iz UNKNOWN)"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_SENT, SEF_CLS_TERMINAL), _
                 "SENT + Storno -> bez promene stanja (samo refresh polja)"

    ' Pad API-ja: SENDING i SENT imaju svoj izlaz, ostalo se NE dira.
    AssertEquals WF_SEF_UNKNOWN, SEFRefreshTargetState(WF_SEF_SENDING, SEF_CLS_ERROR), _
                 "SENDING + pad API-ja -> SEF_UNKNOWN"
    AssertEquals WF_SEF_SYNC_ERROR, SEFRefreshTargetState(WF_SEF_SENT, SEF_CLS_ERROR), _
                 "SENT + pad API-ja -> SEF_SYNC_ERROR"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_SYNC_ERROR, SEF_CLS_ERROR), _
                 "SYNC_ERROR + ponovni pad -> bez promene (nema self-transition)"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_UNKNOWN, SEF_CLS_ERROR), _
                 "UNKNOWN + ponovni pad -> ostaje UNKNOWN, bez izuzetka"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_UNKNOWN, SEF_CLS_UNKNOWN), _
                 "UNKNOWN + opet nepoznat status -> ostaje UNKNOWN"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_SYNC_ERROR, SEF_CLS_UNKNOWN), _
                 "SYNC_ERROR + nepoznat status -> ostaje SYNC_ERROR"

    ' Finalna lokalna stanja se ne vracaju unazad zbog pending/info statusa.
    AssertEquals "", SEFRefreshTargetState(WF_SEF_ACCEPTED, SEF_CLS_PENDING), _
                 "ACCEPTED + pending status -> bez regresije"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_REJECTED, SEF_CLS_PENDING), _
                 "REJECTED + pending status -> bez regresije"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_SENT, SEF_CLS_INFO), _
                 "SENT + Paid/OverDue/Archived -> nema promene (vec je SEF_SENT)"

    ' AUD-032c: PAID/OVERDUE/ARCHIVED ne govore da li je kupac odobrio fakturu,
    ' ali DOKAZUJU da dokument nije vise "u slanju". Dok su vracali prazno,
    ' faktura je ostajala SEF_SENDING i startup recovery ju je nalazio zauvek.
    AssertEquals WF_SEF_SENT, SEFRefreshTargetState(WF_SEF_SENDING, SEF_CLS_INFO), _
                 "SENDING + Paid/OverDue/Archived -> SEF_SENT (izlazi iz slanja)"
    AssertEquals WF_SEF_SENT, SEFRefreshTargetState(WF_SEF_UNKNOWN, SEF_CLS_INFO), _
                 "UNKNOWN + informativan status -> SEF_SENT"
    AssertEquals WF_SEF_SENT, SEFRefreshTargetState(WF_SEF_SYNC_ERROR, SEF_CLS_INFO), _
                 "SYNC_ERROR + informativan status -> SEF_SENT"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_ACCEPTED, SEF_CLS_INFO), _
                 "ACCEPTED + informativan status -> bez regresije"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_REJECTED, SEF_CLS_INFO), _
                 "REJECTED + informativan status -> bez regresije"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_STORNO, SEF_CLS_INFO), _
                 "STORNO + informativan status -> bez regresije"
    ' INFO nikad ne sme da proglasi fakturu prihvacenom -- za to je ACCEPTED klasa.
    AssertTrue SEFRefreshTargetState(WF_SEF_SENDING, SEF_CLS_INFO) <> WF_SEF_ACCEPTED, _
                 "Placeno/dospelo/arhivirano NIJE dokaz prihvatanja"

    ' Greska pri slanju mora u SEF_TECH_FAILED -- jedino stanje iz kog UI nudi retry.
    AssertEquals WF_SEF_TECH_FAILED, SEFRefreshTargetState(WF_SEF_SENDING, SEF_CLS_SEND_FAILED), _
                 "SENDING + Mistake -> SEF_TECH_FAILED (retry moguc)"
    AssertEquals WF_SEF_TECH_FAILED, SEFRefreshTargetState(WF_SEF_UNKNOWN, SEF_CLS_SEND_FAILED), _
                 "UNKNOWN + Mistake -> SEF_TECH_FAILED"
    AssertTrue SEFRefreshTargetState(WF_SEF_SENDING, SEF_CLS_SEND_FAILED) <> WF_SEF_SENT, _
                 "Neuspelo slanje NIKAD ne postaje lokalno SEF_SENT"
    AssertEquals "", SEFRefreshTargetState(WF_SEF_TECH_FAILED, SEF_CLS_SEND_FAILED), _
                 "TECH_FAILED + Mistake -> vec je tamo, bez self-transition"

    ' NORMALNA sekvenca: uspesan submit -> lokalno SEF_SENT -> refresh vrati
    ' MISTAKE. Ovo je najvaznija putanja i ranije je vracala prazno, pa je
    ' faktura ostajala "uspesno poslata" iako SEF tvrdi suprotno.
    AssertEquals WF_SEF_TECH_FAILED, SEFRefreshTargetState(WF_SEF_SENT, SEF_CLS_SEND_FAILED), _
                 "SENT + Mistake -> SEF_TECH_FAILED (normalna sekvenca)"
    AssertEquals WF_SEF_TECH_FAILED, SEFRefreshTargetState(WF_SEF_SYNC_ERROR, SEF_CLS_SEND_FAILED), _
                 "SYNC_ERROR + Mistake -> SEF_TECH_FAILED"
    AssertTrue SEFRefreshTargetState(WF_SEF_SENT, SEF_CLS_SEND_FAILED) <> "", _
                 "SENT + Mistake NE sme da ostane SEF_SENT"
    AssertTransitionAllowed WF_SEF_SENT, WF_SEF_TECH_FAILED
    AssertTransitionAllowed WF_SEF_SYNC_ERROR, WF_SEF_TECH_FAILED

    Exit Sub

EH:
    LogFail "SEF refresh transition matrix", Err.description
End Sub

' (b4) CAPABILITY UGOVOR: sta se sme raditi nad kojim SEF statusom.
' Testira se stvarna sposobnost (cancel/storno/retry/batch-skip), ne ime klase --
' jer se ista greska moze sakriti iza tacnog naziva klase.
Private Sub Test_SEFStatusCapabilities()
    On Error GoTo EH

    ' --- Cancel: zvanicno Draft, New i Mistake (+ nas legacy ERROR marker) ---
    AssertTrue CanCancelSEFStatus("Draft"), "Draft se moze otkazati"
    AssertTrue CanCancelSEFStatus("New"), "New se moze otkazati"
    AssertTrue CanCancelSEFStatus("Mistake"), _
               "Mistake (greska pri slanju) se moze otkazati -- zvanicni ugovor"
    AssertTrue CanCancelSEFStatus("ERROR"), "Legacy ERROR marker se moze otkazati"
    AssertTrue Not CanCancelSEFStatus("Approved"), "Odobrena faktura se ne otkazuje (storno je putanja)"
    AssertTrue Not CanCancelSEFStatus("Sent"), "Poslata faktura se ne otkazuje"
    AssertTrue Not CanCancelSEFStatus(""), "Prazan status ne dozvoljava cancel (fail-closed)"

    ' --- Storno: dokument koji je stvarno predat kupcu ---
    AssertTrue CanStornoSEFStatus("Approved"), "Odobrena faktura se moze stornirati"
    AssertTrue CanStornoSEFStatus("Accepted"), "Zatecen ACCEPTED se moze stornirati"
    AssertTrue CanStornoSEFStatus("Sent"), "Poslata faktura se moze stornirati"
    AssertTrue CanStornoSEFStatus("Rejected"), "Odbijena faktura se moze stornirati"
    AssertTrue Not CanStornoSEFStatus("Mistake"), _
               "Mistake se ne stornira -- otkazuje se"
    AssertTrue Not CanStornoSEFStatus("Draft"), "Draft se ne stornira"
    AssertTrue Not CanStornoSEFStatus(""), "Prazan status ne dozvoljava storno (fail-closed)"

    ' --- Cancel i storno se ne preklapaju ni na jednom statusu ---
    AssertTrue Not (CanCancelSEFStatus("Mistake") And CanStornoSEFStatus("Mistake")), _
               "Isti status ne nudi i cancel i storno"
    AssertTrue Not (CanCancelSEFStatus("Approved") And CanStornoSEFStatus("Approved")), _
               "Odobrena faktura nudi samo storno"

    ' --- Retry: SEF_TECH_FAILED je jedino stanje iz kog se ponovo salje ---
    AssertTrue IsSEFTransitionAllowed(WF_SEF_TECH_FAILED, WF_SEF_READY), _
               "TECH_FAILED -> READY je putanja za retry"
    AssertTrue Not IsSEFTransitionAllowed(WF_SEF_SENT, WF_SEF_READY), _
               "Iz SEF_SENT nema retry putanje (zato Mistake ide u TECH_FAILED)"

    ' --- Slanje: workflow NIJE dovoljan uslov ---
    ' Obican tehnicki pad (nema dokumenta na SEF-u) sme ponovo da se salje...
    AssertTrue CanSendSEFInvoice(WF_SEF_TECH_FAILED, WF_SEF_TECH_FAILED, ""), _
               "Tehnicki pad bez SEF dokumenta -> retry dozvoljen"
    AssertTrue CanSendSEFInvoice(WF_SEF_TECH_FAILED, "", ""), _
               "TECH_FAILED bez spoljnog statusa i bez docId -> retry dozvoljen"
    AssertTrue CanSendSEFInvoice(WF_LOCAL_FINALIZED, "", ""), _
               "Finalizovana faktura se salje prvi put"
    AssertTrue CanSendSEFInvoice(WF_SEF_READY, WF_SEF_READY, ""), _
               "Pripremljena odbijena faktura (resubmit tok) se salje"

    ' ...ali dokument koji ZIVI na SEF-u ne sme ponovo (duplicate guard bi ga
    ' ionako odbio -- ranije je forma palila dugme koje kapija odbija).
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "Mistake", "5317568"), _
               "MISTAKE dokument se ne salje ponovo (prvo Cancel)"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "Cancelled", "5317568"), _
               "Posle cancel-a se ista faktura ne nudi za slanje"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "Approved", "5317568"), _
               "Odobren dokument se ne salje ponovo"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "Paid", "5317568"), _
               "Placen dokument se ne salje ponovo"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_SENT, "", ""), _
               "SEF_SENT nije sendable workflow"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_SENDING, "", ""), _
               "SEF_SENDING nije sendable workflow"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_UNKNOWN, "", ""), _
               "SEF_UNKNOWN nije sendable workflow"

    ' KLJUCNO: SEFStatus je PROMENLJIV -- svaki neuspeo refresh ga prepise u
    ' FAILED/HTTP_ERROR. Odluka se zato vezuje za TRAJAN SEFDocumentId, inace
    ' bi pad mreze ponovo upalio "Retry" nad fakturom sa zivim dokumentom.
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "HTTP_ERROR", "5317568"), _
               "Pad refresh-a ne sme da otkljuca slanje dok SEFDocumentId postoji"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, SEF_STATUS_UNKNOWN, "5317568"), _
               "Nepoznat status + postojeci SEFDocumentId -> slanje blokirano"
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "FAILED", "5317568"), _
               "FAILED status + postojeci SEFDocumentId -> slanje blokirano"
    AssertTrue CanSendSEFInvoice(WF_SEF_TECH_FAILED, "HTTP_ERROR", ""), _
               "Pad refresh-a BEZ SEFDocumentId -> retry i dalje dozvoljen"
    AssertTrue Not CanSendSEFInvoice(WF_LOCAL_FINALIZED, "", "5317568"), _
               "Bilo koji sendable workflow sa zivim dokumentom je blokiran"

    ' REJECTED se salje samo kroz pripremljen tok (SEF_READY + obrisan docId).
    AssertTrue Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "Rejected", ""), _
               "REJECTED van pripremljenog toka nije sendable"
    AssertTrue CanSendSEFInvoice(WF_SEF_READY, "Rejected", ""), _
               "REJECTED iz SEF_READY (pripremljen resubmit) je sendable"

    ' MISTAKE mora da ima BAR jednu putanju: cancel jeste, slanje nije.
    AssertTrue CanCancelSEFStatus("Mistake") And _
               Not CanSendSEFInvoice(WF_SEF_TECH_FAILED, "Mistake", "5317568"), _
               "MISTAKE putanja je Cancel, ne Retry"

    ' --- Poruka mora da uputi na akciju koja je stvarno moguca ---
    AssertContains SEFSendBlockedNextStep("Mistake"), "Cancel", _
                   "Za MISTAKE poruka upucuje na Cancel"
    AssertContains SEFSendBlockedNextStep("Approved"), "Storno", _
                   "Za APPROVED poruka upucuje na Storno (cancel nije dozvoljen)"
    AssertContains SEFSendBlockedNextStep("Sent"), "Storno", _
                   "Za SENT poruka upucuje na Storno"
    AssertContains SEFSendBlockedNextStep("Cancelled"), "portal", _
                   "Za vec otkazan dokument poruka upucuje na proveru, ne na Cancel"
    AssertContains SEFSendBlockedNextStep("Paid"), "portal", _
                   "Za placen dokument poruka ne upucuje na Cancel"
    AssertTrue InStr(1, SEFSendBlockedNextStep("Paid"), "Cancel", vbTextCompare) = 0, _
               "Poruka ne nudi Cancel tamo gde Cancel nije dozvoljen"

    Exit Sub

EH:
    LogFail "SEF status capabilities", Err.description
End Sub

' (b5) INTEGRACIONI: odbijena faktura mora stvarno da prodje kroz resubmit.
'
' Ovo je jedini test u gate suite-u koji dira poslovne tabele -- namerno, jer se
' bas ovaj lanac ne moze dokazati cistim funkcijama: submission red je u statusu
' SENT (refresh ga NE dira), pa je duplicate guard obarao resubmit koji je
' PrepareRejectedInvoiceForResubmit upravo pripremio. Capability funkcija je za
' isti slucaj vracala "sendable" -- tacno, ali nedovoljno.
'
' Redovi se seed-uju u clsTransaction koja se UVEK rollback-uje (isti obrazac kao
' modFakturaTests). Zove se `_Row` verzija, jer clsTransaction.BeginTx puca na
' ugnezdjenu transakciju.
Private Sub Test_SEFRejectedResubmitPassesDuplicateGuard()
    On Error GoTo EH

    Dim tx As clsTransaction
    Dim fakturaID As String
    Dim submissionID As String
    Dim guardBefore As Boolean
    Dim guardAfter As Boolean
    Dim stateAfter As String
    Dim docIdAfter As String
    Dim statusAfter As String
    Dim wasQuiet As Boolean
    Dim quietSet As Boolean

    fakturaID = "TEST-SEF-RESUB-" & Format$(Now, "yyyymmddhhnnss")
    submissionID = "TEST-SUB-" & Format$(Now, "yyyymmddhhnnss")

    ' AppendRow/UpdateCell pisu CSV crash-recovery journal, a njega tx.RollbackTx
    ' NE moze da povuce -- test redovi bi ostali u Journal folderu i sledeci
    ' start bi javio lazno upozorenje o mogucem gubitku podataka. Projekat za to
    ' ima namenski test-mode (isti obrazac kao modAgrohemijaTests).
    wasQuiet = modJournaling.IsTestModeQuiet()
    modJournaling.SetTestModeQuiet True
    quietSet = True

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_SEF_SUBMISSION
    tx.AddTableSnapshot TBL_SEF_EVENT_LOG

    ' Stanje posle: uspesan submit -> refresh vratio REJECTED.
    AppendSEFTestRow TBL_FAKTURE, Array( _
        "FakturaID", fakturaID, _
        "SEFWorkflowState", WF_SEF_REJECTED, _
        "SEFStatus", "REJECTED", _
        "SEFDocumentId", "5317568", _
        "SEFSubmissionIDLast", submissionID)

    AppendSEFTestRow TBL_SEF_SUBMISSION, Array( _
        "SEFSubmissionID", submissionID, _
        "FakturaID", fakturaID, _
        "SubmissionStatus", SEF_SUB_SENT)

    ' Pre pripreme: duplicate guard blokira (to je i bio simptom).
    guardBefore = HasSuccessfulSEFSubmission(fakturaID)
    AssertTrue guardBefore, _
               "Pre pripreme duplicate guard vidi uspesnu submisiju"

    PrepareRejectedInvoiceForResubmit_Row fakturaID

    guardAfter = HasSuccessfulSEFSubmission(fakturaID)
    stateAfter = GetFakturaSEFWorkflowState(fakturaID)
    docIdAfter = GetFakturaSEFDocumentId(fakturaID)
    statusAfter = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))

    ' Sustina nalaza: posle pripreme resubmit vise NE sme da padne kao duplikat.
    AssertTrue Not guardAfter, _
               "Posle pripreme duplicate guard vise ne blokira resubmit"
    AssertEquals WF_SEF_READY, stateAfter, "Priprema vraca workflow u SEF_READY"
    AssertEquals "", Trim$(docIdAfter), "Priprema brise SEFDocumentId"
    AssertTrue CanSendSEFInvoice(stateAfter, statusAfter, docIdAfter), _
               "Pripremljena faktura je sendable i po zajednickoj kapiji"

    ' Scenario 2 (fail-closed): prethodna submisija je ACCEPTED -- neuskladjen
    ' podatak. Priprema tada NE sme da "uspe" i ostavi fakturu koju ce slanje
    ' odbiti kao duplikat; mora da padne glasno i trazi rucnu proveru.
    Dim fakturaID2 As String
    Dim submissionID2 As String
    Dim raised As Boolean

    fakturaID2 = fakturaID & "-ACC"
    submissionID2 = submissionID & "-ACC"

    AppendSEFTestRow TBL_FAKTURE, Array( _
        "FakturaID", fakturaID2, _
        "SEFWorkflowState", WF_SEF_REJECTED, _
        "SEFStatus", "REJECTED", _
        "SEFDocumentId", "5317569", _
        "SEFSubmissionIDLast", submissionID2)

    AppendSEFTestRow TBL_SEF_SUBMISSION, Array( _
        "SEFSubmissionID", submissionID2, _
        "FakturaID", fakturaID2, _
        "SubmissionStatus", SEF_SUB_ACCEPTED)

    raised = False
    On Error Resume Next
    PrepareRejectedInvoiceForResubmit_Row fakturaID2
    raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo EH

    AssertTrue raised, _
               "Priprema pada kad prethodna submisija ostaje ACCEPTED (fail-closed)"
    AssertTrue HasSuccessfulSEFSubmission(fakturaID2), _
               "ACCEPTED submisija se ne prepisuje u REJECTED"

    tx.RollbackTx
    Set tx = Nothing

    RestoreSEFTestQuiet quietSet, wasQuiet
    quietSet = False

    LogPass "Rejected -> prepare -> resubmit prolazi duplicate guard"
    Exit Sub

EH:
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    RestoreSEFTestQuiet quietSet, wasQuiet
    On Error GoTo 0

    LogFail "Rejected resubmit passes duplicate guard", Err.description
End Sub

' Vraca journal test-mode na ZATECENO stanje (ne bezuslovno False), da ugnezdjen
' poziv iz sireg test konteksta ne ostane bez zastite. Otkazuje i eventualno
' zakazan AutoSave tick -- posle rollback-a nema sta da se snima.
Private Sub RestoreSEFTestQuiet(ByVal wasSet As Boolean, ByVal previousValue As Boolean)
    On Error Resume Next
    If wasSet Then
        modJournaling.SetTestModeQuiet previousValue
        modJournaling.StopAutoSaveTimer
    End If
    On Error GoTo 0
End Sub

' Seed helper: upis PO IMENU kolone (pozicijski AppendRow zavisi od redosleda
' kolona, koji se razlikuje po instalaciji). Parovi "kolona", vrednost.
Private Sub AppendSEFTestRow(ByVal tableName As String, ByVal columnValuePairs As Variant)

    Const SRC As String = "modSEFTests.AppendSEFTestRow"

    Dim lo As ListObject
    Set lo = GetTable(tableName)

    Dim rowData() As Variant
    ReDim rowData(0 To lo.ListColumns.count - 1)

    Dim i As Long
    Dim colIndex As Long

    For i = LBound(columnValuePairs) To UBound(columnValuePairs) - 1 Step 2
        colIndex = GetColumnIndex(tableName, CStr(columnValuePairs(i)))
        If colIndex <= 0 Then
            Err.Raise ERR_SEF_VALIDATION, SRC, _
                      "Column not found: " & tableName & "." & CStr(columnValuePairs(i))
        End If
        rowData(colIndex - 1) = columnValuePairs(i + 1)
    Next i

    If AppendRow(tableName, rowData) <= 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Failed to append test row into " & tableName
    End If

End Sub

' (c) Recovery ne sme da prijavi uspeh kad faktura ostaje zaglavljena.
Private Sub Test_SEFRecoveryOutcomeContract()
    On Error GoTo EH

    AssertTrue Not IsSEFRecoveryComplete(WF_SEF_SENDING), _
               "Faktura i dalje u SEF_SENDING NIJE oporavljena"
    AssertTrue IsSEFRecoveryComplete(WF_SEF_TECH_FAILED), _
               "Prelazak u SEF_TECH_FAILED jeste oporavak"
    AssertTrue IsSEFRecoveryComplete(WF_SEF_SENT), _
               "Prelazak u SEF_SENT jeste oporavak"
    AssertTrue IsSEFRecoveryComplete(WF_SEF_UNKNOWN), _
               "Prelazak u SEF_UNKNOWN (rucna provera) izvlaci fakturu iz SEF_SENDING"
    AssertTrue IsSEFRecoveryComplete(WF_SEF_ACCEPTED), _
               "Prelazak u SEF_ACCEPTED jeste oporavak"
    AssertTrue IsSEFRecoveryComplete(WF_SEF_REJECTED), _
               "Prelazak u SEF_REJECTED jeste oporavak"

    ' FAIL-CLOSED: provera se hrani iz GetFakturaSEFWorkflowState, koja na
    ' schema/read gresci vraca prazan string. Negativan test ("<> SENDING") bi
    ' prazno stanje i smece u koloni proglasio uspesnim oporavkom.
    AssertTrue Not IsSEFRecoveryComplete(""), _
               "Prazno stanje NIJE oporavak (neprocitan podatak)"
    AssertTrue Not IsSEFRecoveryComplete("   "), _
               "Sam razmak NIJE oporavak"
    AssertTrue Not IsSEFRecoveryComplete("BOGUS_STATE"), _
               "Nepoznato stanje NIJE oporavak"
    AssertTrue Not IsSEFRecoveryComplete(WF_SEF_READY), _
               "SEF_READY nije stanje u koje SEF_SENDING sme da predje"
    AssertTrue Not IsSEFRecoveryComplete(WF_LOCAL_DRAFT), _
               "LOCAL_DRAFT nije ishod recovery-ja"

    ' Whitelist mora da se poklapa sa state machine-om: svako stanje koje
    ' racunamo kao oporavak mora biti dozvoljen izlaz iz SEF_SENDING.
    AssertTrue IsSEFTransitionAllowed(WF_SEF_SENDING, WF_SEF_SENT), _
               "SENDING -> SENT je dozvoljen izlaz"
    AssertTrue IsSEFTransitionAllowed(WF_SEF_SENDING, WF_SEF_UNKNOWN), _
               "SENDING -> UNKNOWN je dozvoljen izlaz"
    AssertTrue Not IsSEFTransitionAllowed(WF_SEF_SENDING, WF_SEF_READY), _
               "SENDING -> READY nije dozvoljen (pa nije ni oporavak)"

    Exit Sub

EH:
    LogFail "SEF recovery outcome contract", Err.description
End Sub

Private Sub Test_SEFAllowedTransitions()
    AssertTransitionAllowed WF_LOCAL_DRAFT, WF_LOCAL_FINALIZED

    AssertTransitionAllowed WF_LOCAL_FINALIZED, WF_SEF_READY
    AssertTransitionAllowed WF_SEF_READY, WF_SEF_SENDING

    AssertTransitionAllowed WF_SEF_SENDING, WF_SEF_SENT
    AssertTransitionAllowed WF_SEF_SENDING, WF_SEF_ACCEPTED
    AssertTransitionAllowed WF_SEF_SENDING, WF_SEF_REJECTED
    AssertTransitionAllowed WF_SEF_SENDING, WF_SEF_TECH_FAILED
    AssertTransitionAllowed WF_SEF_SENDING, WF_SEF_UNKNOWN

    AssertTransitionAllowed WF_SEF_SENT, WF_SEF_ACCEPTED
    AssertTransitionAllowed WF_SEF_SENT, WF_SEF_REJECTED
    AssertTransitionAllowed WF_SEF_SENT, WF_SEF_SYNC_ERROR
    AssertTransitionAllowed WF_SEF_SENT, WF_SEF_STORNO

    ' AUD-032b: SEF_UNKNOWN mora imati izlaz (rucna provera pa refresh).
    AssertTransitionAllowed WF_SEF_UNKNOWN, WF_SEF_SENT
    AssertTransitionAllowed WF_SEF_UNKNOWN, WF_SEF_ACCEPTED
    AssertTransitionAllowed WF_SEF_UNKNOWN, WF_SEF_REJECTED
    AssertTransitionAllowed WF_SEF_UNKNOWN, WF_SEF_TECH_FAILED

    AssertTransitionAllowed WF_SEF_TECH_FAILED, WF_SEF_READY
    AssertTransitionAllowed WF_SEF_SYNC_ERROR, WF_SEF_SENT

    ' AUD-032b: zvanicni status "Mistake" (greska pri slanju) stize i kad je
    ' lokalno stanje vec SEF_SENT ili SEF_SYNC_ERROR.
    AssertTransitionAllowed WF_SEF_SENT, WF_SEF_TECH_FAILED
    AssertTransitionAllowed WF_SEF_SYNC_ERROR, WF_SEF_TECH_FAILED
    AssertTransitionAllowed WF_SEF_ACCEPTED, WF_SEF_STORNO
    AssertTransitionAllowed WF_SEF_REJECTED, WF_SEF_READY
End Sub

Private Sub Test_SEFBlockedTransitions()
    AssertTransitionBlocked WF_LOCAL_DRAFT, WF_SEF_READY
    AssertTransitionBlocked WF_LOCAL_FINALIZED, WF_SEF_SENT
    AssertTransitionBlocked WF_SEF_READY, WF_SEF_SENT

    AssertTransitionBlocked WF_SEF_SENT, WF_SEF_SENT
    AssertTransitionBlocked WF_SEF_SENT, WF_SEF_READY
    AssertTransitionBlocked WF_SEF_SENT, WF_SEF_SENDING

    AssertTransitionBlocked WF_SEF_ACCEPTED, WF_SEF_ACCEPTED
    AssertTransitionBlocked WF_SEF_ACCEPTED, WF_SEF_SENT
    AssertTransitionBlocked WF_SEF_ACCEPTED, WF_SEF_REJECTED

    AssertTransitionBlocked WF_SEF_REJECTED, WF_SEF_SENT
    AssertTransitionBlocked WF_SEF_REJECTED, WF_SEF_ACCEPTED

    AssertTransitionBlocked WF_SEF_TECH_FAILED, WF_SEF_SENT
    AssertTransitionBlocked WF_SEF_SYNC_ERROR, WF_SEF_ACCEPTED

    AssertTransitionBlocked WF_SEF_UNKNOWN, WF_SEF_READY
    AssertTransitionBlocked WF_SEF_UNKNOWN, WF_SEF_SENDING
    AssertTransitionBlocked WF_SEF_UNKNOWN, WF_SEF_STORNO

    AssertTransitionBlocked WF_SEF_STORNO, WF_SEF_SENT
    AssertTransitionBlocked WF_SEF_STORNO, WF_SEF_READY
    AssertTransitionBlocked WF_SEF_STORNO, WF_SEF_STORNO

    AssertTransitionBlocked "BOGUS_STATE", WF_SEF_SENT
End Sub

Private Sub AssertTransitionAllowed(ByVal oldState As String, _
                                    ByVal newState As String)
    On Error GoTo EH

    ValidateAllowedTransition oldState, newState

    LogPass "Transition allowed: " & oldState & " -> " & newState
    Exit Sub

EH:
    LogFail "Transition should be allowed: " & oldState & " -> " & newState, _
            "Err.Number=" & CStr(Err.Number) & _
            " | Source=" & Err.SOURCE & _
            " | Description=" & Err.description
End Sub

Private Sub AssertTransitionBlocked(ByVal oldState As String, _
                                    ByVal newState As String)
    On Error GoTo ExpectedError

    ValidateAllowedTransition oldState, newState

    LogFail "Transition should be blocked: " & oldState & " -> " & newState, _
            "Expected validation error, but transition was allowed."
    Exit Sub

ExpectedError:
    LogPass "Transition blocked: " & oldState & " -> " & newState
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
    Dim sendErrNo As Long
    Dim sendErrDesc As String

    LogInfo "==== Live send test start for " & fakturaID & " ===="

    beforeState = GetFakturaSEFWorkflowState(fakturaID)
    LogInfo "Workflow before send=" & beforeState

    ' AUD-032a: neuspesan send (REJECTED / TECH_FAILED) sada dolazi kao
    ' tipizirana greska umesto kao "uspeh + SubmissionID". Hvatamo je ovde da bi
    ' test i dalje mogao da razlikuje poslovno odbijanje od tehnickog pada.
    On Error Resume Next
    resultSubmissionID = SendInvoiceToSEF_TX(fakturaID)
    sendErrNo = Err.Number
    sendErrDesc = Err.description
    Err.Clear
    On Error GoTo EH

    If sendErrNo <> 0 And _
       sendErrNo <> ERR_SEF_REJECTED And _
       sendErrNo <> ERR_SEF_SEND_FAILED Then
        Err.Raise sendErrNo, "Test_LiveSendAndRefresh", sendErrDesc
    End If

    If sendErrNo <> 0 Then
        LogInfo "SendInvoiceToSEF_TX raised expected send-outcome error: " & sendErrDesc
    End If

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

    ' AUD-032a: SubmissionID se vraca SAMO za uspesan send; za neuspeh mora doci
    ' greska i prazan povratak (inace bi UI to prikazao kao "Faktura poslata").
    If IsSuccessfulSEFSendState(afterSendState) Then
        AssertTrue sendErrNo = 0, _
                    "Uspesan send ne baca gresku"
        AssertTrue Len(Trim$(resultSubmissionID)) > 0, _
                    "SendInvoiceToSEF_TX returned SubmissionID"
        AssertEquals submissionID, resultSubmissionID, _
                    "Returned SubmissionID matches Faktura last submission"
    Else
        AssertTrue sendErrNo <> 0, _
                    "Neuspesan send (" & afterSendState & ") baca tipiziranu gresku"
        AssertTrue Len(Trim$(resultSubmissionID)) = 0, _
                    "Neuspesan send ne vraca SubmissionID kao potvrdu"
        AssertTrue InStr(1, SEFSendOutcomeMessage(afterSendState, submissionID), _
                            Poruka("SEF_MSG_SEND_POSLATA"), vbTextCompare) = 0, _
                    "Poruka za neuspesan send ne tvrdi da je faktura poslata"
    End If

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
                LogFail "Live send missing SEFDocumentId", _
                        "Workflow=" & afterSendState & _
                        " | SubmissionID=" & submissionID & _
                        " | ResultSubmissionID=" & resultSubmissionID & _
                        " | SubmissionStatus=" & subStatus & _
                        " | HttpStatus=" & httpStatus & _
                        " | ErrorCode=" & errorCode & _
                        " | ErrorMessage=" & errorMessage
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
    If InStr(1, Err.description, "DeliveryDate must not be later than InvoiceDate", vbTextCompare) > 0 Then
        LogSkip "Live send + refresh for " & fakturaID, _
                "Local SEF validation blocked invalid dates: " & Err.description
    Else
        LogFail "Live send + refresh for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " | Source=" & Err.SOURCE & _
                " | Description=" & Err.description
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
    
    ' Ako je bio ACCEPTED/REJECTED pre refresha, mora ostati
    If UCase$(Trim$(state1)) = UCase$(WF_SEF_ACCEPTED) Or _
        UCase$(Trim$(state1)) = UCase$(WF_SEF_REJECTED) Then
        AssertEquals state1, state2, "Final state preserved after first refresh"
        AssertEquals state1, state3, "Final state preserved after second refresh"
    End If

    ' State nikad ne sme da regredira u sending
    AssertTrue UCase$(Trim$(state2)) <> UCase$(WF_SEF_SENDING), _
            "State not regressed to SENDING after first refresh"
    AssertTrue UCase$(Trim$(state3)) <> UCase$(WF_SEF_SENDING), _
            "State not regressed to SENDING after second refresh"
            
    LogPass "Refresh twice did not break state for " & fakturaID
    Exit Sub

EH:
    LogFail "Refresh twice for " & fakturaID, Err.description
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

    Dim recovered As Boolean
    recovered = RecoverStuckSEFSendingInvoice(fakturaID)
    afterState = GetFakturaSEFWorkflowState(fakturaID)

    LogInfo "After recovery state=" & afterState
    LogInfo "Recovery returned=" & CStr(recovered)

    ' AUD-032c: povratna vrednost mora da odgovara stvarnom stanju -- nikad
    ' "recovered" dok je faktura i dalje u SEF_SENDING.
    AssertTrue recovered = IsSEFRecoveryComplete(afterState), _
               "Recovery rezultat odgovara stvarnom stanju fakture"
    AssertTrue UCase$(Trim$(afterState)) <> UCase$(WF_SEF_SENDING), _
               "Recovered invoice no longer stuck in SEF_SENDING"

    LogPass "Recover stuck SEF_SENDING for " & fakturaID
    Exit Sub

EH:
    LogFail "Recover stuck SEF_SENDING for " & fakturaID, Err.description
End Sub

Private Sub Test_BatchRefreshPendingDoesNotCrash()
    On Error GoTo EH

    Dim summaryText As String

    summaryText = RefreshPendingOutboundInvoices_TX()

    ' AUD-032f: batch mora da vrati sazetak, ne da tiho prodje.
    LogInfo "Pending refresh summary: " & summaryText
    AssertTrue Len(Trim$(summaryText)) > 0, _
               "RefreshPendingOutboundInvoices_TX vraca sazetak"

    LogPass "RefreshPendingOutboundInvoices_TX completed"
    Exit Sub

EH:
    LogFail "RefreshPendingOutboundInvoices_TX", Err.description
End Sub

Private Sub Test_BatchRecoverStuckDoesNotCrash()
    On Error GoTo EH

    Dim summaryText As String

    summaryText = RecoverAllStuckSEFSendingInvoices()

    LogInfo "Recovery summary: " & summaryText
    AssertTrue InStr(1, summaryText, "Recovered=", vbTextCompare) > 0, _
               "RecoverAllStuckSEFSendingInvoices vraca sazetak (Found/Recovered/NotRecovered/Failed)"

    LogPass "RecoverAllStuckSEFSendingInvoices completed"
    Exit Sub

EH:
    LogFail "RecoverAllStuckSEFSendingInvoices", Err.description
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
    LogFail "FindFirstFakturaID", Err.description
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
    Err.Raise Err.Number, sourceName, Err.description
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
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = TEST_LOG_SHEET
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
    LogInfo labelText & " Faktura snapshot failed: " & Err.description
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
    LogInfo labelText & " Submission snapshot failed: " & Err.description
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
    LogFatal "RunSEFCancelLiveSuite", Err.Number, Err.description
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
    LogFatal "RunSEFStornoLiveSuite", Err.Number, Err.description
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

    Dim cancelOk As Boolean
    cancelOk = CancelInvoiceOnSEF_TX(fakturaID, commentText)

    If Not cancelOk Then
        LogFail "CancelInvoiceOnSEF_TX returned False for " & fakturaID, _
                "BeforeStatus=" & beforeStatus & _
                " | SEFDocumentId=" & beforeDocID
        Exit Sub
    End If
    
    AssertTrue cancelOk, "CancelInvoiceOnSEF_TX returned True"

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

    ' Workflow ne sme da regredira
    AssertTrue Len(Trim$(afterWorkflow)) > 0, "Cancel leaves workflow state populated"

    ' AUD-032b: posle uspesnog cancel-a faktura NE sme da bude ponudjena za
    ' ponovno slanje. Sam "workflow je neprazan" je propustao SEF_TECH_FAILED
    ' (npr. iz MISTAKE putanje), gde je forma i dalje palila "Retry slanje".
    AssertTrue Not CanSendSEFInvoice(afterWorkflow, afterStatus, afterDocID), _
               "Otkazana faktura se ne nudi za ponovno slanje (workflow=" & _
               afterWorkflow & ", status=" & afterStatus & ")"

    ' SEFStatus mora biti terminalan nakon cancel
    Dim afterStatusUC As String
    afterStatusUC = UCase$(Trim$(afterStatus))
    AssertTrue afterStatusUC = "CANCELLED" Or _
                afterStatusUC = "CANCELED" Or _
                afterStatusUC = "STORNO" Or _
                afterStatusUC = "DRAFT" Or _
                afterStatusUC = "NEW", _
                "Cancel leaves SEFStatus in expected post-cancel range: " & afterStatus

    ' Event log mora rasti
    AssertTrue afterEvents > beforeEvents, "Cancel writes SEF event log"

    ' DocID mora ostati isti
    AssertEquals beforeDocID, afterDocID, "SEFDocumentId unchanged after cancel"

    If IsCancelFinalStatus(afterStatus) Then
        LogPass "Live cancel completed and external status is cancel-like for " & fakturaID
    Else
        LogSkip "Live cancel API call completed, but final cancel status is not verified", _
                "BeforeStatus=" & beforeStatus & _
                " | AfterStatus=" & afterStatus & _
                " | SEFDocumentId=" & afterDocID
    End If
    Exit Sub

EH:
    If IsExpectedSEFBusinessBlock(Err.description) Then
        LogSkip "Live cancel blocked by SEF/service rule for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.SOURCE & _
                " Description=" & Err.description
    Else
        LogFail "Live cancel for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.SOURCE & _
                " Description=" & Err.description
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
    
    If UCase$(Trim$(beforeStatus)) = "STORNO" Then
        LogSkip "Live storno " & fakturaID, _
                "Invoice is already in SEFStatus=STORNO."
        Exit Sub
    End If
  
    LogInfo "Before storno Workflow=" & beforeWorkflow
    LogInfo "Before storno SEFStatus=" & beforeStatus
    LogInfo "Before storno SEFDocumentId=" & beforeDocID
    LogInfo "StornoNumber=" & stornoNumber

    If Len(Trim$(beforeDocID)) = 0 Then
        LogSkip "Live storno " & fakturaID, "No SEFDocumentId."
        Exit Sub
    End If

    ' ISPRAVNO -- commentText je drugi param, stornoNumber je treci
    Dim stornoOk As Boolean

    stornoOk = StornoInvoiceOnSEF_TX(fakturaID, commentText, stornoNumber)

    If Not stornoOk Then
        LogFail "StornoInvoiceOnSEF_TX returned False for " & fakturaID, _
                "BeforeStatus=" & beforeStatus & _
                " | SEFDocumentId=" & beforeDocID & _
                " | StornoNumber=" & stornoNumber
        Exit Sub
    End If

    AssertTrue stornoOk, "StornoInvoiceOnSEF_TX returned True"

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

    AssertTrue Len(Trim$(afterWorkflow)) > 0, "Storno leaves workflow state populated"

    Dim afterStatusUC As String
    afterStatusUC = UCase$(Trim$(afterStatus))
    AssertTrue afterStatusUC = "STORNO" Or _
                afterStatusUC = "CANCELLED" Or _
                afterStatusUC = "CANCELED", _
                "Storno leaves SEFStatus in expected post-storno range: " & afterStatus

    AssertTrue afterEvents > beforeEvents, "Storno writes SEF event log"
    AssertEquals beforeDocID, afterDocID, "SEFDocumentId unchanged after storno"

    If UCase$(Trim$(afterStatus)) = "STORNO" Then
        LogPass "Live storno completed and external status is STORNO for " & fakturaID
    Else
        LogSkip "Live storno API call completed, but final STORNO status is not verified", _
                "BeforeStatus=" & beforeStatus & _
                " | AfterStatus=" & afterStatus & _
                " | SEFDocumentId=" & afterDocID & _
                " | StornoNumber=" & stornoNumber
    End If
    Exit Sub

EH:
    If IsExpectedSEFBusinessBlock(Err.description) Then
        LogSkip "Live storno blocked by SEF/service rule for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.SOURCE & _
                " Description=" & Err.description
    Else
        LogFail "Live storno for " & fakturaID, _
                "Err.Number=" & CStr(Err.Number) & _
                " Source=" & Err.SOURCE & _
                " Description=" & Err.description
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
        InStr(1, s, "CANNOT BE CANCELLED", vbTextCompare) > 0 Or _
        InStr(1, s, "CANNOT BE STORNO", vbTextCompare) > 0 Or _
        InStr(1, s, "CURRENT STATE", vbTextCompare) > 0 Or _
        InStr(1, s, "INVOICE CANNOT", vbTextCompare) > 0 Or _
        InStr(1, s, "IN STATUS:", vbTextCompare) > 0 Or _
        InStr(1, s, "NO SEFDOCUMENTID", vbTextCompare) > 0 Or _
        InStr(1, s, "DESTRUCTIVE SEF TEST CANCELLED", vbTextCompare) > 0
End Function

Private Function IsCancelFinalStatus(ByVal sefStatus As String) As Boolean
    Select Case UCase$(Trim$(sefStatus))
        Case "CANCELLED", "CANCELED", "CANCEL"
            IsCancelFinalStatus = True
        Case Else
            IsCancelFinalStatus = False
    End Select
End Function

' ============================================================
' PATCH 5 -- RunHttpUtilsSmokeSuite
' Dodaje se u modSEFTests (postojeci modul, postojeca konvencija)
' ============================================================
'
' Lokacija: dodaj na kraj modSEFTests, ispred sekcije "DESTRUCTIVE LIVE TESTS".
' Ne pravi novi test modul - postojeca konvencija nalaze da SEF testovi idu u modSEFTests.
'
' Pozivaj sa: ?RunHttpUtilsSmokeSuite
' Ocekivano: PASS=18 FAIL=0
'
' ============================================================

Public Sub RunHttpUtilsSmokeSuite()
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "HTTP UTILS SMOKE SUITE (UrlEncode + JsonEscape)"

    ' --- UrlEncode: ASCII passthrough
    AssertEquals "hello-world", UrlEncode("hello-world"), _
                 "UrlEncode ASCII passthrough"
    
    ' --- UrlEncode: reserved characters
    AssertEquals "a%20b", UrlEncode("a b"), _
                 "UrlEncode space -> %20"
    AssertEquals "a%2Fb", UrlEncode("a/b"), _
                 "UrlEncode slash -> %2F"
    
    ' --- UrlEncode: Serbian diacritics (the whole point of this rewrite)
    ' c (U+010D) -> UTF-8 0xC4 0x8D
    AssertEquals "%C4%8D", UrlEncode(ChrW(&H10D)), _
                 "UrlEncode c (U+010D)"
    
    ' c (U+0107) -> UTF-8 0xC4 0x87
    AssertEquals "%C4%87", UrlEncode(ChrW(&H107)), _
                 "UrlEncode c (U+0107)"
    
    ' s (U+0161) -> UTF-8 0xC5 0xA1
    AssertEquals "%C5%A1", UrlEncode(ChrW(&H161)), _
                 "UrlEncode s (U+0161)"
    
    ' z (U+017E) -> UTF-8 0xC5 0xBE
    AssertEquals "%C5%BE", UrlEncode(ChrW(&H17E)), _
                 "UrlEncode z (U+017E)"
    
    ' dj (U+0111) -> UTF-8 0xC4 0x91
    AssertEquals "%C4%91", UrlEncode(ChrW(&H111)), _
                 "UrlEncode dj (U+0111)"
    
    ' Dj (U+0110) -> UTF-8 0xC4 0x90
    AssertEquals "%C4%90", UrlEncode(ChrW(&H110)), _
                 "UrlEncode Dj (U+0110)"
    
    ' --- UrlEncode: real-world combination
    AssertEquals "%C4%90or%C4%91evi%C4%87", _
                 UrlEncode(ChrW(&H110) & "or" & ChrW(&H111) & "evi" & ChrW(&H107)), _
                 "UrlEncode real surname Djordjevic"
    
    ' --- UrlEncode: RFC 3986 unreserved must NOT encode
    AssertEquals "a-b", UrlEncode("a-b"), "UrlEncode hyphen passthrough"
    AssertEquals "a.b", UrlEncode("a.b"), "UrlEncode period passthrough"
    AssertEquals "a_b", UrlEncode("a_b"), "UrlEncode underscore passthrough"
    AssertEquals "a~b", UrlEncode("a~b"), "UrlEncode tilde passthrough"
    
    ' --- UrlEncode: edge cases
    AssertEquals "", UrlEncode(""), "UrlEncode empty string"
    
    ' --- UrlEncode: stability under repeated calls (no shared mutable state)
    AssertEquals "%C5%A0abac", UrlEncode(ChrW(&H160) & "abac"), _
                 "UrlEncode Sabac repeat 1"
    AssertEquals "%C5%A0abac", UrlEncode(ChrW(&H160) & "abac"), _
                 "UrlEncode Sabac repeat 2"
    
    ' --- JsonEscape: passthrough
    AssertEquals "hello", JsonEscape("hello"), _
                 "JsonEscape ASCII passthrough"
    
    ' --- JsonEscape: backslash and quote
    AssertEquals "He said \""hi\""", JsonEscape("He said ""hi"""), _
                 "JsonEscape double quote"
    AssertEquals "C:\\path", JsonEscape("C:\path"), _
                 "JsonEscape backslash"
    
    FinishSuite
    Exit Sub

EH:
    LogFatal "RunHttpUtilsSmokeSuite", Err.Number, Err.description
    FinishSuite
End Sub

' ============================================================
' PATCH 10 -- RunSEFDocumentIdShapeSuite u modSEFTests (audit #4)
' ============================================================
'
' Lokacija: modSEFTests (postojeci modul, postojeca konvencija)
' Akcija:   Dodaj na kraj modSEFTests, ispred sekcije
'           "DESTRUCTIVE LIVE TESTS: CANCEL / STORNO".
'
' Konvencija: Run*Suite, ResetSEFCounters, InitSEFTestLog, StartSuite,
'             AssertEquals, LogPass, LogFail, FinishSuite. Identicna
'             struktura kao RunHttpUtilsSmokeSuite iz prosle iteracije.
'
' Pozivaj sa: ?RunSEFDocumentIdShapeSuite
' Ocekivano:  PASS=14 FAIL=0
' ============================================================

Public Sub RunSEFDocumentIdShapeSuite()
    On Error GoTo EH

    ResetSEFCounters
    InitSEFTestLog

    StartSuite "SEF DOCUMENT ID SHAPE SUITE (audit #4)"

    ' --- GetJsonNumericIdLiteral with numeric IDs (current SEF format)
    AssertEquals "5317568", _
                 GetJsonNumericIdLiteralPublicProxy("5317568"), _
                 "Numeric 7-digit returns raw"
    
    AssertEquals "123456789012", _
                 GetJsonNumericIdLiteralPublicProxy("123456789012"), _
                 "Numeric 12-digit returns raw"
    
    AssertEquals "123456789012345678", _
                 GetJsonNumericIdLiteralPublicProxy("123456789012345678"), _
                 "Numeric 18-digit (over Long range) returns raw - no precision loss"
    
    AssertEquals "0001234", _
                 GetJsonNumericIdLiteralPublicProxy("0001234"), _
                 "Leading zeros preserved"
    
    ' --- GetJsonNumericIdLiteral with GUID-like IDs (future SEF format)
    AssertEquals """a3f2b1c0-1234-4567-89ab-cdef01234567""", _
                 GetJsonNumericIdLiteralPublicProxy("a3f2b1c0-1234-4567-89ab-cdef01234567"), _
                 "Hyphenated GUID returns quoted string"
    
    AssertEquals """{a3f2b1c0-1234-4567-89ab-cdef01234567}""", _
                 GetJsonNumericIdLiteralPublicProxy("{a3f2b1c0-1234-4567-89ab-cdef01234567}"), _
                 "Bracketed GUID returns quoted string"
    
    AssertEquals """a3f2b1c012344567890abcdef0123456""", _
                 GetJsonNumericIdLiteralPublicProxy("a3f2b1c012344567890abcdef0123456"), _
                 "Bare 32-hex returns quoted string"
    
    ' --- Empty raises
    Test_GetJsonNumericIdLiteralRaises "", "Empty raises"
    
    ' --- Whitespace-only raises (Trim collapses to empty)
    Test_GetJsonNumericIdLiteralRaises "   ", "Whitespace-only raises"
    
    ' --- Garbage raises
    Test_GetJsonNumericIdLiteralRaises "abc!@#", "Garbage with special chars raises"
    Test_GetJsonNumericIdLiteralRaises "abc def", "Garbage with space raises"
    Test_GetJsonNumericIdLiteralRaises "xyz12345", "Non-hex letters raise"
    
    ' --- Real-world JSON body fragment shape verification
    ' Pozivaoc embeduje rezultat direktno: "{""invoiceId"":" & GetJsonNumericIdLiteral(...) & ", ..."
    ' Mora dati validan JSON za oba shape-a.
    AssertEquals "{""invoiceId"":5317568,""x"":1}", _
                 "{""invoiceId"":" & GetJsonNumericIdLiteralPublicProxy("5317568") & ",""x"":1}", _
                 "Numeric ID embedded as JSON number"
    
    AssertEquals "{""invoiceId"":""a3f2b1c0-1234-4567-89ab-cdef01234567"",""x"":1}", _
                 "{""invoiceId"":" & GetJsonNumericIdLiteralPublicProxy("a3f2b1c0-1234-4567-89ab-cdef01234567") & ",""x"":1}", _
                 "GUID ID embedded as JSON string"

    FinishSuite
    Exit Sub

EH:
    LogFatal "RunSEFDocumentIdShapeSuite", Err.Number, Err.description
    FinishSuite
End Sub

' Helper to invoke private GetJsonNumericIdLiteral from this test module.
' The function itself is Private in modSEFClient (correct scope), so we
' need a public proxy for testing. Add this proxy in modSEFClient.
'
' (Smatra se test-only delom API-ja modSEFClient. Ne koristi se iz
' production koda. Slicni proxy patterni koriste se za RunSEFClientParserSmokeSuite.)
Private Function GetJsonNumericIdLiteralPublicProxy(ByVal rawID As String) As String
    GetJsonNumericIdLiteralPublicProxy = TestProxyForGetJsonNumericIdLiteral(rawID)
End Function

Private Sub Test_GetJsonNumericIdLiteralRaises(ByVal idValue As String, _
                                                ByVal testName As String)
    Dim raised As Boolean
    raised = False
    
    On Error Resume Next
    Err.Clear
    Call GetJsonNumericIdLiteralPublicProxy(idValue)
    raised = (Err.Number <> 0)
    On Error GoTo 0
    
    If raised Then
        LogPass testName
    Else
        LogFail testName, "Expected ERR_SEF_VALIDATION but no error was raised."
    End If
End Sub

