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
    
    ' =========================
    ' HTTP call outside TX
    ' requestId = submissionID
    ' =========================
    Set response = SubmitUBLInvoice(ublXml, submissionID)
    
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
                sefStatus:=WF_SEF_ACCEPTED, _
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
                sefStatus:=WF_SEF_SENT, _
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
                message:="Invoice sent to SEF.", _
                details:="SEFDocumentId=" & response.sefDocumentId)
        
        End If
    
    ElseIf response.Rejected Then
        
        Call UpdateFakturaSEFState_Row( _
            fakturaID:=fakturaID, _
            newState:=WF_SEF_REJECTED, _
            sefStatus:=WF_SEF_REJECTED, _
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
            sefStatus:=WF_SEF_TECH_FAILED, _
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
        details:="HTTP=" & CStr(response.HttpStatus) & _
                 "; ApiStatus=" & response.apiStatus & _
                 "; SEFDocumentId=" & response.sefDocumentId)
    
    tx.CommitTx
    
    SendInvoiceToSEF_TX = submissionID
    Exit Function

EH:
    LogErr "SendInvoiceToSEF_TX"
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    If Not txPrep Is Nothing Then txPrep.RollbackTx
    On Error GoTo 0
    
    Err.Raise Err.Number, "SendInvoiceToSEF_TX", Err.Description
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
    LogErr "CancelInvoiceOnSEF_TX"
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    Err.Raise Err.Number, "CancelInvoiceOnSEF_TX", Err.Description
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
    LogErr "StornoInvoiceOnSEF_TX"
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    Err.Raise Err.Number, "StornoInvoiceOnSEF_TX", Err.Description
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

Public Sub RecoverStuckSEFSendingInvoice(ByVal fakturaID As String)
    
    Dim tx As clsTransaction
    Dim currentState As String
    Dim submissionID As String
    Dim sefDocumentId As String
    
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
    
    ' If SEF already knows the document, prefer refresh
    If Len(Trim$(sefDocumentId)) > 0 Then
        Call RefreshSEFStatus_TX(fakturaID)
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_STATE_CHANGED, _
            message:="Recovered stuck SEF_SENDING invoice via status refresh.", _
            details:="SEFDocumentId=" & sefDocumentId)
        
        Exit Sub
    End If
    
    ' Otherwise mark as technical failure so normal retry logic can reuse same submission
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot "tblSEFEventLog"
    
    Call UpdateFakturaSEFState_Row( _
        fakturaID:=fakturaID, _
        newState:=WF_SEF_TECH_FAILED, _
        sefStatus:=WF_SEF_TECH_FAILED, _
        errorCode:="SEF_SENDING_RECOVERY", _
        errorMessage:="Recovered from stuck SEF_SENDING state.", _
        submissionID:=submissionID)
    
    Call AppendSEFEvent_Row( _
        fakturaID:=fakturaID, _
        submissionID:=submissionID, _
        eventType:=SEF_EVT_STATE_CHANGED, _
        message:="Recovered stuck SEF_SENDING invoice to SEF_TECH_FAILED.", _
        details:="SubmissionID=" & submissionID)
    
    tx.CommitTx
    Exit Sub

EH:
    LogErr "RecoverStuckSEFSendingInvoice_TX"
    Dim errNum As Long
    Dim errDesc As String
    
    errNum = Err.Number
    errDesc = Err.Description
    
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    
    If errNum <> 0 Then
        Err.Raise errNum, "RecoverStuckSEFSendingInvoice", errDesc
    Else
        Err.Raise ERR_SEF_STATE, "RecoverStuckSEFSendingInvoice", _
            "Unexpected error recovering stuck SEF_SENDING invoice."
    End If
End Sub

Public Sub RecoverAllStuckSEFSendingInvoices()
    
    Dim data As Variant
    Dim colFakturaID As Long
    Dim colWorkflow As Long
    Dim i As Long
    Dim fakturaID As String
    Dim workflowState As String
    
    On Error GoTo EH
    
    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then Exit Sub
    
    colFakturaID = GetColumnIndex(TBL_FAKTURE, "FakturaID")
    colWorkflow = GetColumnIndex(TBL_FAKTURE, "SEFWorkflowState")
    
    If colFakturaID = 0 Or colWorkflow = 0 Then
        Err.Raise ERR_SEF_STATE, "RecoverAllStuckSEFSendingInvoices", _
            "Required columns missing in tblFakture."
    End If
    
    For i = 1 To UBound(data, 1)
        
        fakturaID = Trim$(CStr(data(i, colFakturaID)))
        workflowState = UCase$(Trim$(CStr(data(i, colWorkflow))))
        
        If workflowState = UCase$(WF_SEF_SENDING) Then
            On Error Resume Next
            Call RecoverStuckSEFSendingInvoice(fakturaID)
            On Error GoTo EH
        End If
        
    Next i
    
    Exit Sub

EH:
    LogErr "RecoverAllStuckSEFSendingInvoice_TX"
    Err.Raise Err.Number, "RecoverAllStuckSEFSendingInvoices", Err.Description
End Sub






Public Sub Test_SendInvoiceToSEF_TX()

    On Error GoTo EH
    
    Dim submissionID As String
    
    submissionID = SendInvoiceToSEF_TX("FAK-00001")
    
    Debug.Print "SubmissionID: "; submissionID
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00001")
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId("FAK-00001")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00001")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test_SendInvoiceToSEF_TX_Debug()

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
    Debug.Print "ERR.Description=" & Err.Description
End Sub

Public Sub Test_CancelInvoiceOnSEF_TX()

    On Error GoTo EH
    
    Debug.Print CancelInvoiceOnSEF_TX("FAK-00007", "Test cancel from VBA")
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test_StornoInvoiceOnSEF_TX()

    On Error GoTo EH
    
    Debug.Print StornoInvoiceOnSEF_TX("FAK-00007", "Test storno from VBA", "STORNO-0002")
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test_SendInvoiceToSEF_TX_RetryCheck()

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
    Debug.Print "ERR.Description=" & Err.Description
End Sub

Public Sub Test_SendInvoiceToSEF_TX_RetryCheck_One(ByVal fakturaID As String)

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
    Debug.Print "ERR.Description=" & Err.Description
End Sub

Public Sub Test_PrepareRejectedInvoiceForResubmit()

    On Error GoTo EH
    
    Call PrepareRejectedInvoiceForResubmit("FAK-00008")
    
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00008")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00008")
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFStatus")
    Debug.Print "SEFLastErrorCode: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFLastErrorCode")
    Debug.Print "SEFLastErrorMessage: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFLastErrorMessage")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test_RecoverStuckSEFSendingInvoice()

    On Error GoTo EH
    
    Call RecoverStuckSEFSendingInvoice("FAK-00008")
    
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00008")
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFStatus")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00008")
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId("FAK-00008")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test_RefreshPendingOutboundInvoices_TX()

    On Error GoTo EH
    
    Call RefreshPendingOutboundInvoices_TX
    Debug.Print "Pending outbound refresh completed."
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test_RecoverAllStuckSEFSendingInvoices()

    On Error GoTo EH
    
    Call RecoverAllStuckSEFSendingInvoices
    Debug.Print "SEF_SENDING recovery completed."
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub
