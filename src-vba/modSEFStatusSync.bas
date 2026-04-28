Attribute VB_Name = "modSEFStatusSync"
 Option Explicit

' =========================================================
' OUTBOUND STATUS MODEL
'
' SEFWorkflowState = internal/local process control state
' SEFStatus        = exact latest external status returned by SEF API
'
' These two fields are related but do NOT have to be identical.
'
' Examples:
'   SEFWorkflowState = SEF_SENT,     SEFStatus = SENT
'   SEFWorkflowState = SEF_SENT,     SEFStatus = DRAFT
'   SEFWorkflowState = SEF_SENT,     SEFStatus = STORNO
'   SEFWorkflowState = SEF_ACCEPTED, SEFStatus = ACCEPTED
'   SEFWorkflowState = SEF_REJECTED, SEFStatus = REJECTED
'
' WorkflowState changes only when the LOCAL state machine changes.
' SEFStatus is updated on every successful refresh from SEF.
' =========================================================

Public Function RefreshSubmissionStatus(ByVal fakturaID As String) As Boolean
    RefreshSubmissionStatus = RefreshSEFStatus_TX(fakturaID)
End Function

Public Function RefreshSEFStatus_TX(ByVal fakturaID As String) As Boolean
    
    Dim tx As clsTransaction
    Dim sefDocumentId As String
    Dim submissionID As String
    Dim response As clsSEFResponse
    Dim apiStatus As String
    Dim currentState As String

    
    On Error GoTo EH
    
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, "RefreshSEFStatus_TX", _
            "No SEFDocumentId found for faktura " & fakturaID
    End If
    
    submissionID = GetLastSEFSubmissionID(fakturaID)
    currentState = GetFakturaSEFWorkflowState(fakturaID)
    
    Set response = GetInvoiceStatus(sefDocumentId)
    apiStatus = UCase$(Trim$(response.apiStatus))
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot "tblSEFSubmission"
    tx.AddTableSnapshot "tblSEFEventLog"
    
    ' Optional:
    ' keep if you intentionally want latest API snapshot on submission row too
    If Len(Trim$(submissionID)) > 0 Then
        'Call SaveSEFSubmissionResult_Row(submissionID, response)
    End If
    
    If response.Accepted Then
        
        Call ApplySEFStateOrRefreshOnly( _
            fakturaID:=fakturaID, _
            targetWorkflowState:=WF_SEF_ACCEPTED, _
            sefStatus:="ACCEPTED", _
            sefDocumentId:=response.sefDocumentId, _
            errorCode:="", _
            errorMessage:="")
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_SYNC_OK, _
            message:="SEF status refreshed: ACCEPTED.", _
            details:="SEFDocumentId=" & response.sefDocumentId)
    
    ElseIf response.Rejected Then
        
        Call ApplySEFStateOrRefreshOnly( _
            fakturaID:=fakturaID, _
            targetWorkflowState:=WF_SEF_REJECTED, _
            sefStatus:="REJECTED", _
            sefDocumentId:=response.sefDocumentId, _
            errorCode:=response.errorCode, _
            errorMessage:=response.errorMessage)
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_VALIDATION_FAILED, _
            message:="SEF status refreshed: REJECTED.", _
            details:=response.errorCode & " | " & response.errorMessage)
    
    ElseIf response.Success Then
        
        Select Case apiStatus
            
            Case "SENT", "NEW", "DRAFT"
                
                Call ApplySEFStateOrRefreshOnly( _
                    fakturaID:=fakturaID, _
                    targetWorkflowState:=WF_SEF_SENT, _
                    sefStatus:=apiStatus, _
                    sefDocumentId:=response.sefDocumentId, _
                    errorCode:="", _
                    errorMessage:="")
                
                Call AppendSEFEvent_Row( _
                    fakturaID:=fakturaID, _
                    submissionID:=submissionID, _
                    eventType:=SEF_EVT_SYNC_OK, _
                    message:="SEF status unchanged (pending).", _
                    details:=apiStatus)
            
            Case "STORNO", "CANCELLED", "CANCELED"
    
                If UCase$(Trim$(currentState)) = UCase$(WF_SEF_SYNC_ERROR) Then
                    Call UpdateFakturaSEFState_Row( _
                        fakturaID:=fakturaID, _
                        newState:=WF_SEF_SENT, _
                        sefStatus:=apiStatus, _
                        sefDocumentId:=response.sefDocumentId, _
                        errorCode:="", _
                        errorMessage:="")
                Else
                    Call UpdateFakturaSEFRefreshFields_Row( _
                        fakturaID:=fakturaID, _
                        sefStatus:=apiStatus, _
                        sefDocumentId:=response.sefDocumentId, _
                        errorCode:="", _
                        errorMessage:="")
                End If
                
                Call AppendSEFEvent_Row( _
                    fakturaID:=fakturaID, _
                    submissionID:=submissionID, _
                    eventType:=SEF_EVT_SYNC_OK, _
                    message:="SEF status refreshed: " & apiStatus & ".", _
                    details:=apiStatus)
            
            Case Else
                
                Call ApplySEFStateOrRefreshOnly( _
                    fakturaID:=fakturaID, _
                    targetWorkflowState:=WF_SEF_SENT, _
                    sefStatus:=apiStatus, _
                    sefDocumentId:=response.sefDocumentId, _
                    errorCode:="", _
                    errorMessage:="")
                
                Call AppendSEFEvent_Row( _
                    fakturaID:=fakturaID, _
                    submissionID:=submissionID, _
                    eventType:=SEF_EVT_SYNC_OK, _
                    message:="SEF returned non-final status.", _
                    details:=apiStatus)
        
        End Select
    
    Else
        
        Call UpdateFakturaSEFState_Row( _
            fakturaID:=fakturaID, _
            newState:=WF_SEF_SYNC_ERROR, _
            sefStatus:=WF_SEF_SYNC_ERROR, _
            errorCode:=response.errorCode, _
            errorMessage:=response.errorMessage)
        
        Call AppendSEFEvent_Row( _
            fakturaID:=fakturaID, _
            submissionID:=submissionID, _
            eventType:=SEF_EVT_SYNC_FAILED, _
            message:="SEF status refresh failed.", _
            details:=response.errorCode & " | " & response.errorMessage)
    
    End If
    
    Call UpdateSEFLastSyncAt_Row(fakturaID)
    
    tx.CommitTx
    
    RefreshSEFStatus_TX = True
    Exit Function

EH:
    LogErr "RefreshSEFStatus_TX"
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    
    Err.Raise Err.Number, "RefreshSEFStatus_TX", Err.Description
End Function

Private Sub ApplySEFStateOrRefreshOnly(ByVal fakturaID As String, _
                                       ByVal targetWorkflowState As String, _
                                       ByVal sefStatus As String, _
                                       Optional ByVal sefDocumentId As String = "", _
                                       Optional ByVal errorCode As String = "", _
                                       Optional ByVal errorMessage As String = "")
    On Error GoTo EH

    Dim currentState As String
    currentState = UCase$(Trim$(GetFakturaSEFWorkflowState(fakturaID)))

    Dim targetState As String
    targetState = UCase$(Trim$(targetWorkflowState))

    If currentState = "" Then
        UpdateFakturaSEFState_Row _
            fakturaID:=fakturaID, _
            newState:=targetWorkflowState, _
            sefStatus:=sefStatus, _
            sefDocumentId:=sefDocumentId, _
            errorCode:=errorCode, _
            errorMessage:=errorMessage
        Exit Sub
    End If

    ' Idempotent refresh:
    ' ako je workflow vec u ciljnom stanju, ne radimo transition sam u sebe.
    If currentState = targetState Then
        UpdateFakturaSEFRefreshFields_Row _
            fakturaID:=fakturaID, _
            sefStatus:=sefStatus, _
            sefDocumentId:=sefDocumentId, _
            errorCode:=errorCode, _
            errorMessage:=errorMessage
        Exit Sub
    End If

    ' Ne vracamo finalne lokalne state-ove nazad u SEF_SENT samo zato
    ' što eksterni API vrati pending/non-final status.
    If targetState = UCase$(WF_SEF_SENT) Then
        If IsFinalLocalSEFWorkflowState(currentState) Then
            UpdateFakturaSEFRefreshFields_Row _
                fakturaID:=fakturaID, _
                sefStatus:=sefStatus, _
                sefDocumentId:=sefDocumentId, _
                errorCode:=errorCode, _
                errorMessage:=errorMessage
            Exit Sub
        End If
    End If

    ' Ako je prethodni refresh pao i faktura je u SEF_SYNC_ERROR,
    ' a novi refresh sada vraca finalni status, prvo je vratimo u SEF_SENT,
    ' jer state machine dozvoljava SEF_SYNC_ERROR -> SEF_SENT,
    ' pa zatim SEF_SENT -> finalni state.
    If currentState = UCase$(WF_SEF_SYNC_ERROR) Then
        Select Case targetState
            Case UCase$(WF_SEF_ACCEPTED), UCase$(WF_SEF_REJECTED)
                UpdateFakturaSEFState_Row _
                    fakturaID:=fakturaID, _
                    newState:=WF_SEF_SENT, _
                    sefStatus:=sefStatus, _
                    sefDocumentId:=sefDocumentId, _
                    errorCode:=errorCode, _
                    errorMessage:=errorMessage
        End Select
    End If

    UpdateFakturaSEFState_Row _
        fakturaID:=fakturaID, _
        newState:=targetWorkflowState, _
        sefStatus:=sefStatus, _
        sefDocumentId:=sefDocumentId, _
        errorCode:=errorCode, _
        errorMessage:=errorMessage

    Exit Sub

EH:
    LogErr "modSEFStatusSync.ApplySEFStateOrRefreshOnly"
    Err.Raise Err.Number, "modSEFStatusSync.ApplySEFStateOrRefreshOnly", Err.Description
End Sub

Private Function IsFinalLocalSEFWorkflowState(ByVal workflowState As String) As Boolean
    Select Case UCase$(Trim$(workflowState))
        Case UCase$(WF_SEF_ACCEPTED), _
             UCase$(WF_SEF_REJECTED), _
             UCase$(WF_SEF_STORNO)
            IsFinalLocalSEFWorkflowState = True
        Case Else
            IsFinalLocalSEFWorkflowState = False
    End Select
End Function

Private Function IsTerminalExternalRefreshStatus(ByVal sefStatus As String) As Boolean
    Select Case UCase$(Trim$(sefStatus))
        Case "STORNO", "CANCELLED", "CANCELED"
            IsTerminalExternalRefreshStatus = True
        Case Else
            IsTerminalExternalRefreshStatus = False
    End Select
End Function

Public Sub RefreshPendingOutboundInvoices_TX()

    On Error GoTo EH

    Const SRC As String = "modSEFStatusSync.RefreshPendingOutboundInvoices_TX"

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then Exit Sub

    Dim colFakturaID As Long
    Dim colWorkflow As Long
    Dim colSEFStatus As Long

    colFakturaID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", SRC)
    colWorkflow = RequireColumnIndex(TBL_FAKTURE, "SEFWorkflowState", SRC)
    colSEFStatus = RequireColumnIndex(TBL_FAKTURE, "SEFStatus", SRC)

    Dim i As Long
    Dim fakturaID As String
    Dim workflowState As String
    Dim sefStatus As String

    For i = 1 To UBound(data, 1)

        fakturaID = Trim$(CStr(data(i, colFakturaID)))
        workflowState = UCase$(Trim$(CStr(data(i, colWorkflow))))
        sefStatus = UCase$(Trim$(CStr(data(i, colSEFStatus))))

        Select Case workflowState

            Case UCase$(WF_SEF_SENT), UCase$(WF_SEF_SYNC_ERROR)
                
                If IsTerminalExternalRefreshStatus(sefStatus) Then GoTo NextInvoice

                On Error Resume Next
                RefreshSEFStatus_TX fakturaID

                If Err.Number <> 0 Then
                    LogErr SRC & ".Invoice." & fakturaID
                    Err.Clear
                End If

                On Error GoTo EH

                Application.Wait Now + TimeSerial(0, 0, 2)

        End Select

NextInvoice:
    Next i

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Sub


Public Sub Test2_RefreshSEFStatus_TX()

    On Error GoTo EH
    
    Dim ok As Boolean
    Dim fakturaID As String
    
    fakturaID = "FAK-00008"
    
    ok = RefreshSEFStatus_TX(fakturaID)
    
    Debug.Print "Refresh OK: "; ok
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState(fakturaID)
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId(fakturaID)
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID(fakturaID)
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus")
    Debug.Print "SEFLastErrorCode: "; LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFLastErrorCode")
    Debug.Print "SEFLastErrorMessage: "; LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFLastErrorMessage")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub

Public Sub Test1_RefreshSEFStatus_TX()

    On Error GoTo EH
    
    Dim ok As Boolean
    
    ok = RefreshSEFStatus_TX("FAK-00008")
    
    Debug.Print "Refresh OK: "; ok
    Debug.Print "WorkflowState: "; GetFakturaSEFWorkflowState("FAK-00008")
    Debug.Print "SEFDocumentId: "; GetFakturaSEFDocumentId("FAK-00008")
    Debug.Print "LastSubmissionID: "; GetLastSEFSubmissionID("FAK-00008")
    Debug.Print "SEFStatus: "; LookupValue(TBL_FAKTURE, "FakturaID", "FAK-00008", "SEFStatus")
    
    Exit Sub

EH:
    Debug.Print "ERR " & Err.Number & " - " & Err.Description
End Sub
