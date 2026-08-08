Attribute VB_Name = "modSEFService"
Option Explicit
'Call UpdateFakturaSEFState_Row("FAK-00006", WF_SEF_READY, WF_SEF_READY)

' NOTE:
' WF_SEF_SENT means the outbound send pipeline succeeded locally.
' The exact external SEF status is tracked separately in SEFStatus
' and may later become SENT, DRAFT, STORNO, CANCELLED, ACCEPTED, etc.
Public Function SendInvoiceToSEF_TX(ByVal fakturaID As String) As String
    
    Dim tx As clsTransaction
    Dim txPrep As clsTransaction
    
    Dim dto As clsSEFInvoiceSnapshot
    Dim ublXml As String
    Dim payloadHash As String
    Dim submissionID As String
    Dim versionNo As Long
    Dim currentState As String
    Dim response As clsSEFResponse
    
    Dim reuseLastSubmission As Boolean
    Dim localSendingCommitted As Boolean
    Dim finalState As String

    On Error GoTo EH
    
    currentState = GetFakturaSEFWorkflowState(fakturaID)
    reuseLastSubmission = ShouldReuseLastSubmission(fakturaID)
    
    ' =========================
    ' Validate + build or reuse payload
    ' =========================
    Call ValidateFakturaForSEF(fakturaID)
    
    If reuseLastSubmission Then
        
        submissionID = GetLastSEFSubmissionID(fakturaID)
        ublXml = GetSubmissionRequestBody(submissionID)
        payloadHash = GetSubmissionPayloadHash(submissionID)
        versionNo = GetCurrentSEFVersionNo(fakturaID)
        
        If Len(Trim$(submissionID)) = 0 Then
            Err.Raise ERR_SEF_STATE, "SendInvoiceToSEF_TX", _
                "Retry requested but no previous submission ID exists."
        End If
        
        If Len(Trim$(ublXml)) = 0 Then
            Err.Raise ERR_SEF_STATE, "SendInvoiceToSEF_TX", _
                "Retry requested but previous request body is empty."
        End If
        
        If Len(Trim$(payloadHash)) = 0 Then
            payloadHash = ComputePayloadHash(ublXml)
        End If
        
    Else
        
        Set dto = BuildSEFInvoiceDto(fakturaID)
        ublXml = SerializeUBLInvoice(dto)
        
        If Len(Trim$(ublXml)) = 0 Then
            Err.Raise ERR_SEF_VALIDATION, "SendInvoiceToSEF_TX", _
                "Generated UBL XML is empty."
        End If
        
        payloadHash = ComputePayloadHash(ublXml)
        versionNo = GetNextSEFVersionNo(fakturaID)
        
    End If
    
    ' =========================
    ' PREP TX: move to SEF_READY if needed
    ' =========================
    If currentState = WF_LOCAL_FINALIZED Or currentState = WF_SEF_TECH_FAILED Then
        
        Set txPrep = New clsTransaction
        txPrep.BeginTx
        txPrep.AddTableSnapshot "tblFakture"
        txPrep.AddTableSnapshot "tblSEFEventLog"
        
        Call UpdateFakturaSEFState_Row( _
            fakturaID:=fakturaID, _
            newState:=WF_SEF_READY, _
            sefStatus:=WF_SEF_READY, _
            payloadHash:=payloadHash, _
            versionNo:=versionNo)
        
        If reuseLastSubmission Then
            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_STATE_CHANGED, _
                message:="Retrying previous SEF technical failed submission.", _
                details:="RequestId=" & submissionID)
        Else
            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:="", _
                eventType:=SEF_EVT_STATE_CHANGED, _
                message:="Invoice moved to SEF_READY before submit.", _
                details:="PreviousState=" & currentState)
        End If
        
        txPrep.CommitTx
        Set txPrep = Nothing
        
    End If
    
    ' =========================
    ' TX 1: create submission if new, then set SENDING
    ' =========================
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot "tblFakture"
    tx.AddTableSnapshot "tblSEFSubmission"
    tx.AddTableSnapshot "tblSEFEventLog"
    
    If Not reuseLastSubmission Then
        submissionID = CreateSEFSubmission_Row( _
            fakturaID:=fakturaID, _
            versionNo:=versionNo, _
            workflowState:=WF_SEF_SENDING, _
            payloadHash:=payloadHash, _
            requestBody:=ublXml, _
            requestFormat:="XML")
    End If
    
    Call UpdateFakturaSEFState_Row( _
        fakturaID:=fakturaID, _
        newState:=WF_SEF_SENDING, _
        sefStatus:=WF_SEF_SENDING, _
        payloadHash:=payloadHash, _
        submissionID:=submissionID, _
        versionNo:=versionNo)
    
    Call AppendSEFEvent_Row( _
        fakturaID:=fakturaID, _
        submissionID:=submissionID, _
        eventType:=SEF_EVT_HTTP_SENT, _
        message:="SEF UBL submission started.", _
        details:="PayloadHash=" & payloadHash & "; RequestId=" & submissionID)
    
    tx.CommitTx
    Set tx = Nothing
    
    localSendingCommitted = True

    On Error Resume Next
    Monitor_SEF _
        eventType:="SEF_SEND_START", _
        severity:="INFO", _
        invoiceLocalId:=fakturaID, _
        businessInvoiceNo:=fakturaID, _
        sefStatus:=WF_SEF_SENDING, _
        localStatus:=WF_SEF_SENDING, _
        sefRequestId:=submissionID, _
        attemptCount:=0, _
        lastError:="", _
        nextAction:="WAIT", _
        needsManualReview:=False
    On Error GoTo EH
    
    ' =========================
    ' HTTP call outside TX
    ' requestId = submissionID
    ' =========================
    Set response = SubmitUBLInvoice(ublXml, submissionID)
    
    ' =========================
    ' Guard: successful SEF submit must return SEFDocumentId
    ' =========================
    If Not response Is Nothing Then
        If response.Success Then
            If Len(Trim$(response.sefDocumentId)) = 0 Then
                response.Success = False
                response.Accepted = False
                response.Rejected = False
                response.apiStatus = "FAILED"
                response.errorCode = "MISSING_SEF_DOCUMENT_ID"
                response.errorMessage = _
                    "SEF submit returned success but SEFDocumentId is missing."
            End If
        End If
    End If
    
    ' =========================
    ' TX 2: save result + set final state
    ' =========================
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot "tblFakture"
    tx.AddTableSnapshot "tblSEFSubmission"
    tx.AddTableSnapshot "tblSEFEventLog"
    
    Call SaveSEFSubmissionResult_Row(submissionID, response)
    
    If response.Success Then
        
        If response.Accepted Then
            
            Call UpdateFakturaSEFState_Row( _
                fakturaID:=fakturaID, _
                newState:=WF_SEF_ACCEPTED, _
                sefStatus:=response.apiStatus, _
                sefDocumentId:=response.sefDocumentId, _
                errorCode:="", _
                errorMessage:="", _
                payloadHash:=payloadHash, _
                submissionID:=submissionID, _
                versionNo:=versionNo)
            
            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:="SEF_ACCEPTED", _
                message:="Invoice accepted by SEF.", _
                details:="SEFDocumentId=" & response.sefDocumentId)
        
        Else
            
            Call UpdateFakturaSEFState_Row( _
                fakturaID:=fakturaID, _
                newState:=WF_SEF_SENT, _
                sefStatus:=response.apiStatus, _
                sefDocumentId:=response.sefDocumentId, _
                errorCode:="", _
                errorMessage:="", _
                payloadHash:=payloadHash, _
                submissionID:=submissionID, _
                versionNo:=versionNo)
            
            Call AppendSEFEvent_Row( _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_STATE_CHANGED, _
                message:="Invoice submitted to SEF and SEFDocumentId received.", _
                details:="ApiStatus=" & response.apiStatus & _
                         "; SEFDocumentId=" & response.sefDocumentId)
        
        End If
    
    ElseIf response.Rejected Then
        
        Call UpdateFakturaSEFState_Row( _
            fakturaID:=fakturaID, _
            newState:=WF_SEF_REJECTED, _
            sefStatus:=response.apiStatus, _
            errorCode:=response.errorCode, _
            errorMessage:=response.errorMessage, _
            payloadHash:=payloadHash, _
            submissionID:=submissionID, _
            versionNo:=versionNo)
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_VALIDATION_FAILED, _
            message:="SEF rejected invoice.", _
            details:=response.errorCode & " | " & response.errorMessage)
    
    Else
        
        Call UpdateFakturaSEFState_Row( _
            fakturaID:=fakturaID, _
            newState:=WF_SEF_TECH_FAILED, _
            sefStatus:=response.apiStatus, _
            errorCode:=response.errorCode, _
            errorMessage:=response.errorMessage, _
            payloadHash:=payloadHash, _
            submissionID:=submissionID, _
            versionNo:=versionNo)
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_SYNC_FAILED, _
            message:="Technical failure during SEF submit.", _
            details:=response.errorCode & " | " & response.errorMessage)
    
    End If
    
    Call AppendSEFEvent_Row( _
        fakturaID:=fakturaID, _
        submissionID:=submissionID, _
        eventType:=SEF_EVT_HTTP_RESPONSE, _
        message:="SEF response saved.", _
        details:="HTTP=" & CStr(response.httpStatus) & _
                 "; ApiStatus=" & response.apiStatus & _
                 "; SEFDocumentId=" & response.sefDocumentId)
    
    tx.CommitTx

    On Error Resume Next

    If response Is Nothing Then

        Monitor_SEF _
            eventType:="SEF_SEND_FAIL", _
            severity:="ERROR", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:="FAILED", _
            localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
            sefRequestId:=submissionID, _
            sefInvoiceId:="", _
            attemptCount:=0, _
            lastHttpCode:="0", _
            lastError:="SEF response object is Nothing after submit.", _
            nextAction:="RETRY", _
            needsManualReview:=False

    ElseIf response.Accepted Then

        Monitor_SEF _
            eventType:="SEF_SEND_ACCEPTED", _
            severity:="INFO", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:=response.apiStatus, _
            localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
            sefRequestId:=submissionID, _
            sefInvoiceId:=response.sefDocumentId, _
            attemptCount:=0, _
            lastHttpCode:=CStr(response.httpStatus), _
            lastError:="", _
            nextAction:="WAIT", _
            needsManualReview:=False

    ElseIf response.Success Then

        Monitor_SEF _
            eventType:="SEF_SEND_SUCCESS", _
            severity:="INFO", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:=response.apiStatus, _
            localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
            sefRequestId:=submissionID, _
            sefInvoiceId:=response.sefDocumentId, _
            attemptCount:=0, _
            lastHttpCode:=CStr(response.httpStatus), _
            lastError:="", _
            nextAction:="WAIT", _
            needsManualReview:=False

    ElseIf response.Rejected Then

        Monitor_SEF _
            eventType:="SEF_SEND_REJECTED", _
            severity:="WARN", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:=response.apiStatus, _
            localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
            sefRequestId:=submissionID, _
            sefInvoiceId:=response.sefDocumentId, _
            attemptCount:=0, _
            lastHttpCode:=CStr(response.httpStatus), _
            lastError:=response.errorCode & " | " & response.errorMessage, _
            nextAction:="MANUAL_REVIEW", _
            needsManualReview:=True

    ElseIf UCase$(Trim$(response.apiStatus)) = "CONFLICT" Then

        ' AUD-030: 409 CONFLICT is NOT a rejection. The invoice may already
        ' exist on SEF; a human must reconcile by requestId/docId before any
        ' retry. Flag manual review instead of the default RETRY signal.
        Monitor_SEF _
            eventType:="SEF_SEND_CONFLICT", _
            severity:="WARN", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:=response.apiStatus, _
            localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
            sefRequestId:=submissionID, _
            sefInvoiceId:=response.sefDocumentId, _
            attemptCount:=0, _
            lastHttpCode:=CStr(response.httpStatus), _
            lastError:=response.errorCode & " | " & response.errorMessage, _
            nextAction:="MANUAL_REVIEW", _
            needsManualReview:=True

    Else

        Monitor_SEF _
            eventType:="SEF_SEND_FAIL", _
            severity:="ERROR", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:=response.apiStatus, _
            localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
            sefRequestId:=submissionID, _
            sefInvoiceId:=response.sefDocumentId, _
            attemptCount:=0, _
            lastHttpCode:=CStr(response.httpStatus), _
            lastError:=response.errorCode & " | " & response.errorMessage, _
            nextAction:="RETRY", _
            needsManualReview:=False

    End If

    On Error GoTo 0

    ' =========================
    ' AUD-032a: ishod slanja
    ' =========================
    ' TX 2 je vec komitovan, pa je stanje fakture (REJECTED / TECH_FAILED)
    ' sacuvano i pre ovog raise-a. Ranije se SubmissionID vracao bezuslovno, pa
    ' je frmSEF i za odbijenu fakturu prikazivao "Faktura poslata".
    finalState = GetFakturaSEFWorkflowState(fakturaID)

    If Not IsSuccessfulSEFSendState(finalState) Then
        Err.Raise SEFSendFailureErrNumber(finalState), "SendInvoiceToSEF_TX", _
                  SEFSendOutcomeMessage(finalState, submissionID)
    End If

    SendInvoiceToSEF_TX = submissionID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "SendInvoiceToSEF_TX"

    Monitor_Error _
        moduleName:="modSEFService", _
        procedureName:="SendInvoiceToSEF_TX", _
        entityType:="Faktura", _
        entityID:=fakturaID, _
        correlationId:=fakturaID, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    If localSendingCommitted Then
        Monitor_SEF _
            eventType:="SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING", _
            severity:="CRITICAL", _
            invoiceLocalId:=fakturaID, _
            businessInvoiceNo:=fakturaID, _
            sefStatus:="UNKNOWN", _
            localStatus:=WF_SEF_SENDING, _
            sefRequestId:=submissionID, _
            sefInvoiceId:="", _
            attemptCount:=0, _
            lastHttpCode:="vba-exception", _
            lastError:=errDesc, _
            nextAction:="CHECK_SEF_PORTAL", _
            needsManualReview:=True
    End If

    If Not tx Is Nothing Then tx.RollbackTx
    If Not txPrep Is Nothing Then txPrep.RollbackTx

    On Error GoTo 0
    
    LogErr "SendInvoiceToSEF_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    If Not txPrep Is Nothing Then txPrep.RollbackTx
    On Error GoTo 0

    If errNum <> 0 Then
        Err.Raise errNum, "SendInvoiceToSEF_TX", _
                  "Source=" & errSrc & " | " & errDesc
    Else
        Err.Raise ERR_SEF_STATE, "SendInvoiceToSEF_TX", _
                  "Unexpected SEF send error; original Err was lost before EH capture."
    End If
End Function

' =========================================================
' AUD-032a: ugovor o ishodu slanja
'
' Uspesan send = faktura je stigla na SEF (SEF_SENT ili SEF_ACCEPTED).
' Sve ostalo (REJECTED, TECH_FAILED, bilo sta trece) je neuspeh i NE sme da
' vrati SubmissionID kao potvrdu, niti da se korisniku prikaze kao "poslata".
'
' Funkcije su ciste (bez I/O) da bi RunSEFTestSuite mogao da ih proveri bez
' ijednog poziva ka pravom SEF-u.
' =========================================================
Public Function IsSuccessfulSEFSendState(ByVal workflowState As String) As Boolean
    Select Case UCase$(Trim$(workflowState))
        Case UCase$(WF_SEF_SENT), UCase$(WF_SEF_ACCEPTED)
            IsSuccessfulSEFSendState = True
        Case Else
            IsSuccessfulSEFSendState = False
    End Select
End Function

Public Function SEFSendFailureErrNumber(ByVal workflowState As String) As Long
    If UCase$(Trim$(workflowState)) = UCase$(WF_SEF_REJECTED) Then
        SEFSendFailureErrNumber = ERR_SEF_REJECTED
    Else
        SEFSendFailureErrNumber = ERR_SEF_SEND_FAILED
    End If
End Function

Public Function SEFSendOutcomeMessage(ByVal workflowState As String, _
                                      ByVal submissionID As String) As String

    Dim stateText As String
    Dim messageText As String

    stateText = Trim$(workflowState)
    If Len(stateText) = 0 Then stateText = "?"

    Select Case UCase$(stateText)
        Case UCase$(WF_SEF_ACCEPTED)
            messageText = Poruka("SEF_MSG_SEND_PRIHVACENA")
        Case UCase$(WF_SEF_SENT)
            messageText = Poruka("SEF_MSG_SEND_POSLATA")
        Case UCase$(WF_SEF_REJECTED)
            messageText = Poruka("SEF_MSG_SEND_ODBIJENA")
        Case UCase$(WF_SEF_TECH_FAILED)
            messageText = Poruka("SEF_MSG_SEND_TEH_GRESKA")
        Case Else
            messageText = Poruka("SEF_MSG_SEND_NEPOZNAT_ISHOD")
    End Select

    messageText = messageText & vbCrLf & "Workflow: " & stateText

    If Len(Trim$(submissionID)) > 0 Then
        messageText = messageText & " | SubmissionID: " & submissionID
    End If

    SEFSendOutcomeMessage = messageText
End Function

' AUD-032c: recovery je uspeo samo ako faktura vise NIJE u SEF_SENDING.
' Ranije se "Recovered" event upisivao bezuslovno, pa je zaglavljena faktura pri
' svakom startup-u davala lazan izvestaj o oporavku.
Public Function IsSEFRecoveryComplete(ByVal stateAfterRecovery As String) As Boolean
    IsSEFRecoveryComplete = _
        (UCase$(Trim$(stateAfterRecovery)) <> UCase$(WF_SEF_SENDING))
End Function

Public Function CancelInvoiceOnSEF_TX(ByVal fakturaID As String, ByVal cancelComment As String) As Boolean
    
    Dim tx As clsTransaction
    Dim sefDocumentId As String
    Dim submissionID As String
    Dim response As clsSEFResponse
    
    On Error GoTo EH
    
    Call ValidateFakturaCanBeCancelledOnSEF(fakturaID)
    
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    submissionID = GetLastSEFSubmissionID(fakturaID)
    
    Set response = CancelInvoiceOnSEF(sefDocumentId, cancelComment)
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot "tblSEFEventLog"
    
    If response.Success Then
        Call UpdateFakturaSEFRefreshFields_Row( _
            fakturaID:=fakturaID, _
            sefStatus:=response.apiStatus, _
            sefDocumentId:=response.sefDocumentId, _
            errorCode:="", _
            errorMessage:="")
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_SYNC_OK, _
            message:="Invoice cancelled on SEF.", _
            details:=response.apiStatus & " | " & cancelComment)
    Else
        Call UpdateFakturaSEFRefreshFields_Row( _
            fakturaID:=fakturaID, _
            errorCode:=response.errorCode, _
            errorMessage:=response.errorMessage)
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_SYNC_FAILED, _
            message:="SEF cancel failed.", _
            details:=response.errorCode & " | " & response.errorMessage)
    End If
    
    Call UpdateSEFLastSyncAt_Row(fakturaID)
    
    tx.CommitTx
    CancelInvoiceOnSEF_TX = response.Success
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "CancelInvoiceOnSEF_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    If errNum <> 0 Then
        Err.Raise errNum, "CancelInvoiceOnSEF_TX", _
                  "Source=" & errSrc & " | " & errDesc
    Else
        Err.Raise ERR_SEF_STATE, "CancelInvoiceOnSEF_TX", _
                  "Unexpected error during SEF cancel; original Err was lost before EH capture."
    End If
End Function

Public Function StornoInvoiceOnSEF_TX(ByVal fakturaID As String, ByVal stornoComment As String, Optional ByVal stornoNumber As String = "") As Boolean
    
    Dim tx As clsTransaction
    Dim sefDocumentId As String
    Dim submissionID As String
    Dim response As clsSEFResponse
    
    On Error GoTo EH
    
    Call ValidateFakturaCanBeStorniranoOnSEF(fakturaID)
    
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    submissionID = GetLastSEFSubmissionID(fakturaID)
    
    Set response = StornoInvoiceOnSEF(sefDocumentId, stornoComment, stornoNumber)
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot "tblSEFEventLog"
    
    If response.Success Then
        Call UpdateFakturaSEFRefreshFields_Row( _
            fakturaID:=fakturaID, _
            sefStatus:=response.apiStatus, _
            sefDocumentId:=response.sefDocumentId, _
            errorCode:="", _
            errorMessage:="")
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_SYNC_OK, _
            message:="Invoice storno created on SEF.", _
            details:=response.apiStatus & " | " & stornoComment)
    Else
        Call UpdateFakturaSEFRefreshFields_Row( _
            fakturaID:=fakturaID, _
            errorCode:=response.errorCode, _
            errorMessage:=response.errorMessage)
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_SYNC_FAILED, _
            message:="SEF storno failed.", _
            details:=response.errorCode & " | " & response.errorMessage)
    End If
    
    Call UpdateSEFLastSyncAt_Row(fakturaID)
    
    tx.CommitTx
    StornoInvoiceOnSEF_TX = response.Success
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "StornoInvoiceOnSEF_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    If errNum <> 0 Then
        Err.Raise errNum, "StornoInvoiceOnSEF_TX", _
                  "Source=" & errSrc & " | " & errDesc
    Else
        Err.Raise ERR_SEF_STATE, "StornoInvoiceOnSEF_TX", _
                  "Unexpected error during SEF storno; original Err was lost before EH capture."
    End If
End Function

Private Function ShouldReuseLastSubmission(ByVal fakturaID As String) As Boolean
    
    Dim lastSubmissionID As String
    Dim lastSubmissionStatus As String
    Dim workflowState As String
    
    lastSubmissionID = GetLastSEFSubmissionID(fakturaID)
    If Len(Trim$(lastSubmissionID)) = 0 Then Exit Function
    
    lastSubmissionStatus = UCase$(Trim$(GetLastSEFSubmissionStatus(fakturaID)))
    workflowState = UCase$(Trim$(GetFakturaSEFWorkflowState(fakturaID)))
    
    Select Case workflowState
        Case UCase$(WF_SEF_TECH_FAILED)
            Select Case lastSubmissionStatus
                Case UCase$(SEF_SUB_FAILED), UCase$(SEF_SUB_CREATED)
                    ShouldReuseLastSubmission = True
            End Select
    End Select

End Function

' AUD-032c: vraca True samo ako faktura vise nije zaglavljena u SEF_SENDING.
' Statement-pozivi (`Call RecoverStuckSEFSendingInvoice(id)`) rade i dalje.
Public Function RecoverStuckSEFSendingInvoice(ByVal fakturaID As String) As Boolean

    Dim tx As clsTransaction
    Dim currentState As String
    Dim submissionID As String
    Dim sefDocumentId As String
    Dim refreshOk As Boolean
    Dim stateAfterRefresh As String
    Dim recovered As Boolean

    On Error GoTo EH

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, "RecoverStuckSEFSendingInvoice", _
                  "FakturaID is required."
    End If

    currentState = GetFakturaSEFWorkflowState(fakturaID)

    If currentState <> WF_SEF_SENDING Then
        Err.Raise ERR_SEF_STATE, "RecoverStuckSEFSendingInvoice", _
                  "Invoice is not in SEF_SENDING state: " & currentState
    End If

    submissionID = GetLastSEFSubmissionID(fakturaID)
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)

    ' If SEF already knows the document, prefer refresh.
    If Len(Trim$(sefDocumentId)) > 0 Then

        refreshOk = RefreshSEFStatus_TX(fakturaID)
        stateAfterRefresh = GetFakturaSEFWorkflowState(fakturaID)
        recovered = IsSEFRecoveryComplete(stateAfterRefresh)

        ' AUD-032c: "Recovered" se upisuje SAMO kad je refresh zaista vratio
        ' upotrebljiv status I kad je faktura izasla iz SEF_SENDING. Ranije se
        ' taj event upisivao bezuslovno, pri svakom startup-u.
        If Not recovered Then
            AppendSEFEvent_Row _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_FAILED, _
                message:="Recovery via status refresh did not clear SEF_SENDING.", _
                details:="SEFDocumentId=" & sefDocumentId & _
                         "; RefreshOk=" & CStr(refreshOk) & _
                         "; State=" & stateAfterRefresh
        ElseIf refreshOk Then
            AppendSEFEvent_Row _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_STATE_CHANGED, _
                message:="Recovered stuck SEF_SENDING invoice via status refresh.", _
                details:="SEFDocumentId=" & sefDocumentId & _
                         "; NewState=" & stateAfterRefresh
        Else
            AppendSEFEvent_Row _
                fakturaID:=fakturaID, _
                submissionID:=submissionID, _
                eventType:=SEF_EVT_SYNC_FAILED, _
                message:="Stuck SEF_SENDING invoice moved out of SEF_SENDING for manual review; SEF returned no usable status.", _
                details:="SEFDocumentId=" & sefDocumentId & _
                         "; NewState=" & stateAfterRefresh
        End If

        RecoverStuckSEFSendingInvoice = recovered
        Exit Function
    End If

    ' Otherwise mark as technical failure so normal retry logic can reuse same submission.
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot "tblSEFEventLog"

    UpdateFakturaSEFState_Row _
        fakturaID:=fakturaID, _
        newState:=WF_SEF_TECH_FAILED, _
        sefStatus:=WF_SEF_TECH_FAILED, _
        errorCode:="SEF_SENDING_RECOVERY", _
        errorMessage:="Recovered from stuck SEF_SENDING state.", _
        submissionID:=submissionID

    AppendSEFEvent_Row _
        fakturaID:=fakturaID, _
        submissionID:=submissionID, _
        eventType:=SEF_EVT_STATE_CHANGED, _
        message:="Recovered stuck SEF_SENDING invoice to SEF_TECH_FAILED.", _
        details:="SubmissionID=" & submissionID

    tx.CommitTx
    Set tx = Nothing

    RecoverStuckSEFSendingInvoice = _
        IsSEFRecoveryComplete(GetFakturaSEFWorkflowState(fakturaID))
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "RecoverStuckSEFSendingInvoice"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    If errNum <> 0 Then
        Err.Raise errNum, "RecoverStuckSEFSendingInvoice", _
                  "Source=" & errSrc & " | " & errDesc
    Else
        Err.Raise ERR_SEF_STATE, "RecoverStuckSEFSendingInvoice", _
                  "Unexpected error recovering stuck SEF_SENDING invoice; original Err was lost before EH capture."
    End If
End Function

' AUD-032f: vraca sazetak prolaza (Found/Recovered/NotRecovered/Failed) da bi
' frmSEF mogao da ga pokaze operateru umesto tihog "recovery zavrsen".
' Statement-pozivi (`Call RecoverAllStuckSEFSendingInvoices`, modMain startup)
' rade i dalje.
Public Function RecoverAllStuckSEFSendingInvoices() As String

    On Error GoTo EH

    Const SRC As String = "modSEFService.RecoverAllStuckSEFSendingInvoices"

    Dim foundCount As Long
    Dim recoveredCount As Long
    Dim notRecoveredCount As Long
    Dim failedCount As Long
    Dim itemRecovered As Boolean
    Dim summaryText As String

    On Error Resume Next
    Monitor_Event _
        eventType:="SEF_STARTUP_RECOVERY_START", _
        severity:="INFO", _
        message:="Started recovery for stuck SEF_SENDING invoices", _
        userId:="Operator", _
        moduleName:="modSEFService", _
        procedureName:="RecoverAllStuckSEFSendingInvoices", _
        entityType:="SEF", _
        entityID:="StartupRecovery", _
        correlationId:="SEF-STARTUP-RECOVERY"
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then GoTo SuccessExit

    Dim colFakturaID As Long
    Dim colWorkflow As Long

    colFakturaID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", SRC)
    colWorkflow = RequireColumnIndex(TBL_FAKTURE, "SEFWorkflowState", SRC)

    Dim i As Long
    Dim fakturaID As String
    Dim workflowState As String

    For i = 1 To UBound(data, 1)

        fakturaID = Trim$(CStr(data(i, colFakturaID)))
        workflowState = UCase$(Trim$(CStr(data(i, colWorkflow))))

        If workflowState = UCase$(WF_SEF_SENDING) Then

            foundCount = foundCount + 1

            On Error Resume Next
            Monitor_SEF _
                eventType:="SEF_RECOVERY_INVOICE_FOUND", _
                severity:="WARN", _
                invoiceLocalId:=fakturaID, _
                businessInvoiceNo:=fakturaID, _
                sefStatus:=WF_SEF_SENDING, _
                localStatus:=WF_SEF_SENDING, _
                sefRequestId:=GetLastSEFSubmissionID(fakturaID), _
                attemptCount:=0, _
                lastError:="Stuck SEF_SENDING invoice found during startup recovery", _
                nextAction:="WAIT", _
                needsManualReview:=False
            On Error GoTo EH

            On Error Resume Next
            itemRecovered = False
            itemRecovered = RecoverStuckSEFSendingInvoice(fakturaID)

            If Err.Number <> 0 Then
                failedCount = failedCount + 1

                Dim itemErrNo As Long
                Dim itemErrDesc As String
                Dim itemErrSrc As String

                itemErrNo = Err.Number
                itemErrDesc = Err.description
                itemErrSrc = Err.SOURCE

                LogErr SRC & ".Invoice." & fakturaID

                Monitor_SEF _
                    eventType:="SEF_RECOVERY_INVOICE_FAIL", _
                    severity:="CRITICAL", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:="UNKNOWN", _
                    localStatus:=WF_SEF_SENDING, _
                    sefRequestId:=GetLastSEFSubmissionID(fakturaID), _
                    attemptCount:=0, _
                    lastHttpCode:="vba-error", _
                    lastError:=itemErrDesc, _
                    nextAction:="CHECK_SEF_PORTAL", _
                    needsManualReview:=True

                Monitor_Error _
                    moduleName:="modSEFService", _
                    procedureName:="RecoverAllStuckSEFSendingInvoices.Invoice", _
                    entityType:="Faktura", _
                    entityID:=fakturaID, _
                    correlationId:=fakturaID, _
                    errorNumber:=itemErrNo, _
                    errorDescription:=itemErrDesc, _
                    errorSource:=itemErrSrc

                Err.Clear
            ElseIf itemRecovered Then
                recoveredCount = recoveredCount + 1

                Monitor_SEF _
                    eventType:="SEF_RECOVERY_INVOICE_SUCCESS", _
                    severity:="INFO", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=GetLastSEFSubmissionID(fakturaID), _
                    sefInvoiceId:=GetFakturaSEFDocumentId(fakturaID), _
                    attemptCount:=0, _
                    lastError:="", _
                    nextAction:="WAIT", _
                    needsManualReview:=False
            Else
                ' AUD-032c: poziv nije pukao, ali faktura je i dalje u
                ' SEF_SENDING. To se vise ne prijavljuje kao uspesan oporavak.
                notRecoveredCount = notRecoveredCount + 1

                Monitor_SEF _
                    eventType:="SEF_RECOVERY_INVOICE_NOT_RECOVERED", _
                    severity:="WARN", _
                    invoiceLocalId:=fakturaID, _
                    businessInvoiceNo:=fakturaID, _
                    sefStatus:=WF_SEF_SENDING, _
                    localStatus:=GetFakturaSEFWorkflowState(fakturaID), _
                    sefRequestId:=GetLastSEFSubmissionID(fakturaID), _
                    sefInvoiceId:=GetFakturaSEFDocumentId(fakturaID), _
                    attemptCount:=0, _
                    lastError:="Invoice is still stuck in SEF_SENDING after recovery attempt.", _
                    nextAction:="CHECK_SEF_PORTAL", _
                    needsManualReview:=True
            End If

            On Error GoTo EH

        End If

    Next i

SuccessExit:
    summaryText = "Found=" & foundCount & _
                  "; Recovered=" & recoveredCount & _
                  "; NotRecovered=" & notRecoveredCount & _
                  "; Failed=" & failedCount

    RecoverAllStuckSEFSendingInvoices = summaryText

    On Error Resume Next
    Monitor_Event _
        eventType:="SEF_STARTUP_RECOVERY_SUCCESS", _
        severity:="INFO", _
        message:="Startup SEF recovery completed. " & summaryText, _
        userId:="Operator", _
        moduleName:="modSEFService", _
        procedureName:="RecoverAllStuckSEFSendingInvoices", _
        entityType:="SEF", _
        entityID:="StartupRecovery", _
        correlationId:="SEF-STARTUP-RECOVERY"
    On Error GoTo 0

    Exit Function

EH:
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr SRC

    Monitor_Error _
        moduleName:="modSEFService", _
        procedureName:="RecoverAllStuckSEFSendingInvoices", _
        entityType:="SEF", _
        entityID:="StartupRecovery", _
        correlationId:="SEF-STARTUP-RECOVERY", _
        errorNumber:=errNo, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="SEF_STARTUP_RECOVERY_FAIL", _
        severity:="CRITICAL", _
        message:=errDesc, _
        userId:="Operator", _
        moduleName:="modSEFService", _
        procedureName:="RecoverAllStuckSEFSendingInvoices", _
        entityType:="SEF", _
        entityID:="StartupRecovery", _
        correlationId:="SEF-STARTUP-RECOVERY"

    On Error GoTo 0
    Err.Raise errNo, SRC, errDesc
End Function


' =========================================================
' DEV MAKROI (AUD-032e)
'
' Sve ispod je Private -- vise se NE vidi u Alt+F8 listi. Ovi makroi gadjaju
' ZIVI SEF (slanje fakture je pravni cin) sa hardkodovanim FakturaID-jem i bez
' potvrde, pa im nije mesto u operaterskoj listi makroa. Iz VBE-a se i dalje
' pokrecu (kursor u proceduri -> F5).
'
' Test_CancelInvoiceOnSEF_TX / Test_StornoInvoiceOnSEF_TX su UKLONJENI: cancel i
' storno na SEF-u su pravni cin, a gadjani ekvivalenti vec postoje u modSEFTests
' (RunSEFCancelLiveSuite / RunSEFStornoLiveSuite -- SEF_TEST_ALLOW_LIVE +
' SEF_TEST_ALLOW_CANCEL_STORNO + tipkana potvrda).
' =========================================================

Private Sub Test_SendInvoiceToSEF_TX()

    On Error GoTo EH
    
    Dim submissionID As String
    
    submissionID = SendInvoiceToSEF_TX("FAK-00001")
    
    Debug.Print "SubmissionID: "; submissionID
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00001")
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId("FAK-00001")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00001")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub

Private Sub Test_SendInvoiceToSEF_TX_Debug()

    On Error GoTo EH
    
    Debug.Print "A Start"
    
    Dim submissionID As String
    submissionID = SendInvoiceToSEF_TX("FAK-00008")
    
    Debug.Print "B SubmissionID=" & submissionID
    Debug.Print "C WorkflowState=" & GetFakturaSEFWorkflowState("FAK-00008")
    Debug.Print "D SEFDocumentId=" & GetFakturaSEFDocumentId("FAK-00008")
    Debug.Print "E LastSubmissionID=" & GetLastSEFSubmissionID("FAK-00008")
    Debug.Print "F Done"
    
    Exit Sub

EH:
    Debug.Print "ERR.Number=" & Err.Number
    Debug.Print "ERR.Description=" & Err.description
End Sub

Private Sub Test_SendInvoiceToSEF_TX_RetryCheck()

    On Error GoTo EH
    
    Dim fakturaID As String
    Dim submissionBefore As String
    Dim submissionAfter As String
    Dim workflowBefore As String
    Dim workflowAfter As String
    Dim statusBefore As String
    Dim statusAfter As String
    
    fakturaID = "FAK-00008"   ' change as needed
    
    Debug.Print "======================================"
    Debug.Print "RETRY TEST START"
    Debug.Print "FakturaID=" & fakturaID
    
    workflowBefore = GetFakturaSEFWorkflowState(fakturaID)
    submissionBefore = GetLastSEFSubmissionID(fakturaID)
    statusBefore = GetLastSEFSubmissionStatus(fakturaID)
    
    Debug.Print "Before WorkflowState=" & workflowBefore
    Debug.Print "Before LastSubmissionID=" & submissionBefore
    Debug.Print "Before LastSubmissionStatus=" & statusBefore
    
    submissionAfter = SendInvoiceToSEF_TX(fakturaID)
    
    workflowAfter = GetFakturaSEFWorkflowState(fakturaID)
    statusAfter = GetLastSEFSubmissionStatus(fakturaID)
    
    Debug.Print "After WorkflowState=" & workflowAfter
    Debug.Print "After LastSubmissionID=" & submissionAfter
    Debug.Print "After LastSubmissionStatus=" & statusAfter
    
    If Len(Trim$(submissionBefore)) > 0 Then
        If submissionBefore = submissionAfter Then
            Debug.Print "RESULT: SAME submission reused (retry path)."
        Else
            Debug.Print "RESULT: NEW submission created."
        End If
    Else
        Debug.Print "RESULT: No previous submission existed."
    End If
    
    Debug.Print "======================================"
    Exit Sub

EH:
    Debug.Print "ERR.Number=" & Err.Number
    Debug.Print "ERR.Description=" & Err.description
End Sub

Private Sub Test_SendInvoiceToSEF_TX_RetryCheck_One(ByVal fakturaID As String)

    On Error GoTo EH
    
    Dim submissionBefore As String
    Dim submissionAfter As String
    Dim workflowBefore As String
    Dim workflowAfter As String
    Dim statusBefore As String
    Dim statusAfter As String
    
    Debug.Print "======================================"
    Debug.Print "RETRY TEST START"
    Debug.Print "FakturaID=" & fakturaID
    
    workflowBefore = GetFakturaSEFWorkflowState(fakturaID)
    submissionBefore = GetLastSEFSubmissionID(fakturaID)
    statusBefore = GetLastSEFSubmissionStatus(fakturaID)
    
    Debug.Print "Before WorkflowState=" & workflowBefore
    Debug.Print "Before LastSubmissionID=" & submissionBefore
    Debug.Print "Before LastSubmissionStatus=" & statusBefore
    
    submissionAfter = SendInvoiceToSEF_TX(fakturaID)
    
    workflowAfter = GetFakturaSEFWorkflowState(fakturaID)
    statusAfter = GetLastSEFSubmissionStatus(fakturaID)
    
    Debug.Print "After WorkflowState=" & workflowAfter
    Debug.Print "After LastSubmissionID=" & submissionAfter
    Debug.Print "After LastSubmissionStatus=" & statusAfter
    
    If Len(Trim$(submissionBefore)) > 0 Then
        If submissionBefore = submissionAfter Then
            Debug.Print "RESULT: SAME submission reused (retry path)."
        Else
            Debug.Print "RESULT: NEW submission created."
        End If
    Else
        Debug.Print "RESULT: No previous submission existed."
    End If
    
    Debug.Print "======================================"
    Exit Sub

EH:
    Debug.Print "ERR.Number=" & Err.Number
    Debug.Print "ERR.Description=" & Err.description
End Sub

Private Sub Test_PrepareRejectedInvoiceForResubmit()

    On Error GoTo EH
    
    Call PrepareRejectedInvoiceForResubmit("FAK-00008")
    
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00008")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00008")
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFStatus")
    Debug.Print "SEFLastErrorCode: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFLastErrorCode")
    Debug.Print "SEFLastErrorMessage: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFLastErrorMessage")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub

Private Sub Test_RecoverStuckSEFSendingInvoice()

    On Error GoTo EH
    
    Call RecoverStuckSEFSendingInvoice("FAK-00008")
    
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00008")
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFStatus")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00008")
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId("FAK-00008")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub

Private Sub Test_RefreshPendingOutboundInvoices_TX()

    On Error GoTo EH
    
    Call RefreshPendingOutboundInvoices_TX
    Debug.Print "Pending outbound refresh completed."
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub

Private Sub Test_RecoverAllStuckSEFSendingInvoices()

    On Error GoTo EH
    
    Call RecoverAllStuckSEFSendingInvoices
    Debug.Print "SEF_SENDING recovery completed."
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.description
End Sub
