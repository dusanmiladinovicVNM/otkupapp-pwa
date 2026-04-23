Attribute VB_Name = "modSEFPersistance"
Option Explicit

' =========================================================
' modSEFPersistence
' Alle SEF-Reads/Writes laufen über modDataAccess
' =========================================================

' =========================
' READ HELPERS
' =========================

Public Function GetFakturaSEFWorkflowState(ByVal fakturaID As String) As String
    Dim v As Variant
    
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFWorkflowState")
    
    If IsEmpty(v) Then
        GetFakturaSEFWorkflowState = ""
    Else
        GetFakturaSEFWorkflowState = Trim$(CStr(v))
    End If
End Function

Public Function GetFakturaSEFDocumentId(ByVal fakturaID As String) As String
    Dim v As Variant
    
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFDocumentId")
    
    If IsEmpty(v) Then
        GetFakturaSEFDocumentId = ""
    Else
        GetFakturaSEFDocumentId = Trim$(CStr(v))
    End If
End Function

Public Function GetLastSEFSubmissionID(ByVal fakturaID As String) As String
    Dim v As Variant
    
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFSubmissionIDLast")
    
    If IsEmpty(v) Then
        GetLastSEFSubmissionID = ""
    Else
        GetLastSEFSubmissionID = Trim$(CStr(v))
    End If
End Function

Public Function GetNextSEFVersionNo(ByVal fakturaID As String) As Long
    Dim v As Variant
    
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFVersionNo")
    
    If IsEmpty(v) Then
        GetNextSEFVersionNo = 1
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        GetNextSEFVersionNo = 1
    ElseIf Not IsNumeric(v) Then
        GetNextSEFVersionNo = 1
    Else
        GetNextSEFVersionNo = CLng(v) + 1
    End If
End Function

Public Function GetCurrentSEFVersionNo(ByVal fakturaID As String) As Long
    Dim v As Variant
    
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFVersionNo")
    
    If IsEmpty(v) Then
        GetCurrentSEFVersionNo = 0
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        GetCurrentSEFVersionNo = 0
    ElseIf Not IsNumeric(v) Then
        GetCurrentSEFVersionNo = 0
    Else
        GetCurrentSEFVersionNo = CLng(v)
    End If
End Function

' =========================
' WRITE HELPERS
' =========================
' NOTE:
' newState controls internal workflow transitions only.
' sefStatus stores the latest known external SEF API status.
' Do not assume newState and sefStatus must be identical.

Public Sub UpdateFakturaSEFState_Row( _
    ByVal fakturaID As String, _
    ByVal newState As String, _
    Optional ByVal sefStatus As String = "", _
    Optional ByVal sefDocumentId As String = "", _
    Optional ByVal errorCode As String = "", _
    Optional ByVal errorMessage As String = "", _
    Optional ByVal payloadHash As String = "", _
    Optional ByVal submissionID As String = "", _
    Optional ByVal versionNo As Long = 0)
    
    Dim rowIndex As Long
    Dim oldState As String
    Dim ok As Boolean
    
    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, "UpdateFakturaSEFState_Row", "FakturaID is required."
    End If
    
    If Len(Trim$(newState)) = 0 Then
        Err.Raise ERR_SEF_STATE, "UpdateFakturaSEFState_Row", "newState is required."
    End If
    
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)
    
    oldState = GetFakturaSEFWorkflowState(fakturaID)
    If Len(oldState) > 0 Then
        Call ValidateAllowedTransition(oldState, newState)
    End If
    
    ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFWorkflowState", newState)
    If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFWorkflowState", fakturaID
    
    If Len(sefStatus) > 0 Then
        ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFStatus", sefStatus)
        If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFStatus", fakturaID
    End If
    
    If Len(sefDocumentId) > 0 Then
        ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFDocumentId", sefDocumentId)
        If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFDocumentId", fakturaID
    End If
    
    ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFLastErrorCode", errorCode)
    If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFLastErrorCode", fakturaID
    
    ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFLastErrorMessage", errorMessage)
    If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFLastErrorMessage", fakturaID
    
    If Len(payloadHash) > 0 Then
        ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFPayloadHash", payloadHash)
        If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFPayloadHash", fakturaID
    End If
    
    If Len(submissionID) > 0 Then
        ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFSubmissionIDLast", submissionID)
        If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFSubmissionIDLast", fakturaID
    End If
    
    If versionNo > 0 Then
        ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFVersionNo", versionNo)
        If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFVersionNo", fakturaID
    End If
    
    Select Case newState
        Case WF_SEF_SENT, WF_SEF_ACCEPTED
            ok = UpdateCell(TBL_FAKTURE, rowIndex, "PoslatNaSEF", "Da")
            If Not ok Then RaiseUpdateError TBL_FAKTURE, "PoslatNaSEF", fakturaID
            
            If Len(Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFSentAt")))) = 0 Then
                ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFSentAt", Now)
                If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFSentAt", fakturaID
            End If
        
        Case WF_SEF_SENDING
            ok = UpdateCell(TBL_FAKTURE, rowIndex, "PoslatNaSEF", "Ne")
            If Not ok Then RaiseUpdateError TBL_FAKTURE, "PoslatNaSEF", fakturaID
    End Select
    
    Select Case newState
        Case WF_SEF_SENT, WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_SYNC_ERROR
            ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFLastSyncAt", Now)
            If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFLastSyncAt", fakturaID
    End Select
    
End Sub

' NOTE:
' This helper updates refresh-related fields without performing
' a workflow transition. Use it when SEFStatus changes but the
' internal workflow state should remain unchanged.

Public Sub UpdateFakturaSEFRefreshFields_Row( _
    ByVal fakturaID As String, _
    Optional ByVal sefStatus As String = "", _
    Optional ByVal sefDocumentId As String = "", _
    Optional ByVal errorCode As String = "", _
    Optional ByVal errorMessage As String = "")
    
    Dim rowIndex As Long
    Dim ok As Boolean
    
    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, "UpdateFakturaSEFRefreshFields_Row", "FakturaID is required."
    End If
    
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)
    
    If Len(sefStatus) > 0 Then
        ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFStatus", sefStatus)
        If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFStatus", fakturaID
    End If
    
    If Len(sefDocumentId) > 0 Then
        ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFDocumentId", sefDocumentId)
        If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFDocumentId", fakturaID
    End If
    
    ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFLastErrorCode", errorCode)
    If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFLastErrorCode", fakturaID
    
    ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFLastErrorMessage", errorMessage)
    If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFLastErrorMessage", fakturaID
End Sub

Public Function CreateSEFSubmission_Row( _
    ByVal fakturaID As String, _
    ByVal versionNo As Long, _
    ByVal workflowState As String, _
    ByVal payloadHash As String, _
    ByVal requestBody As String, _
    ByVal requestFormat As String) As String
    
    Dim submissionID As String
    Dim rowData(1 To 20) As Variant
    Dim newRowIndex As Long
    
    submissionID = GetNextID("tblSEFSubmission", "SEFSubmissionID", "SFS-")
    
    rowData(1) = submissionID            ' SEFSubmissionID
    rowData(2) = fakturaID               ' FakturaID
    rowData(3) = versionNo               ' VersionNo
    rowData(4) = workflowState           ' WorkflowStateAtSubmit
    rowData(5) = Now                     ' CreatedAt
    rowData(6) = Empty                   ' SubmittedAt
    rowData(7) = SEF_SUB_CREATED         ' SubmissionStatus
    rowData(8) = payloadHash             ' PayloadHash
    rowData(9) = requestFormat           ' RequestFormat
    rowData(10) = requestBody            ' RequestBody
    rowData(11) = Empty                  ' ResponseBody
    rowData(12) = Empty                  ' HttpStatus
    rowData(13) = Empty                  ' ApiStatus
    rowData(14) = Empty                  ' CorrelationId
    rowData(15) = Empty                  ' SEFDocumentId
    rowData(16) = Empty                  ' ErrorCode
    rowData(17) = Empty                  ' ErrorMessage
    rowData(18) = GetCurrentOperatorName() ' OperatorName
    rowData(19) = "Ne"                   ' Stornirano
    rowData(20) = Empty                  ' FinishedAt
    
    newRowIndex = AppendRow("tblSEFSubmission", rowData)
    
    If newRowIndex <= 0 Then
        Err.Raise ERR_SEF_STATE, "CreateSEFSubmission_Row", _
            "Could not append row to tblSEFSubmission."
    End If
    
    CreateSEFSubmission_Row = submissionID
End Function
Public Sub SaveSEFSubmissionResult_Row( _
    ByVal submissionID As String, _
    ByVal response As clsSEFResponse)
    
    Dim rowIndex As Long
    Dim ok As Boolean
    Dim subStatus As String
    
    If Len(Trim$(submissionID)) = 0 Then Exit Sub
    
    If response Is Nothing Then
        Err.Raise ERR_SEF_RESPONSE_PARSE, "SaveSEFSubmissionResult_Row", _
            "Response object is Nothing."
    End If
    
    rowIndex = GetSingleRowIndexByKey("tblSEFSubmission", "SEFSubmissionID", submissionID, True)
    
    If response.Accepted Then
        subStatus = SEF_SUB_ACCEPTED
    ElseIf response.Rejected Then
        subStatus = SEF_SUB_REJECTED
    ElseIf response.Success Then
        subStatus = SEF_SUB_SENT
    Else
        subStatus = SEF_SUB_FAILED
    End If
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "SubmittedAt", Now)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "SubmittedAt", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "FinishedAt", Now)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "FinishedAt", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "SubmissionStatus", subStatus)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "SubmissionStatus", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "HttpStatus", response.HttpStatus)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "HttpStatus", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "ApiStatus", response.apiStatus)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "ApiStatus", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "CorrelationId", response.CorrelationId)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "CorrelationId", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "ResponseBody", response.RawBody)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "ResponseBody", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "SEFDocumentId", response.sefDocumentId)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "SEFDocumentId", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "ErrorCode", response.errorCode)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "ErrorCode", submissionID
    
    ok = UpdateCell("tblSEFSubmission", rowIndex, "ErrorMessage", response.errorMessage)
    If Not ok Then RaiseUpdateError "tblSEFSubmission", "ErrorMessage", submissionID
End Sub

Public Sub AppendSEFEvent_Row( _
    ByVal fakturaID As String, _
    ByVal submissionID As String, _
    ByVal eventType As String, _
    ByVal message As String, _
    Optional ByVal details As String = "")
    
    Dim eventID As String
    Dim rowData(1 To 9) As Variant
    Dim newRowIndex As Long
    
    eventID = GetNextID("tblSEFEventLog", "SEFEventID", "SFE-")
    
    rowData(1) = eventID                         ' SEFEventID
    rowData(2) = fakturaID                       ' FakturaID
    rowData(3) = submissionID                    ' SEFSubmissionID
    rowData(4) = Now                             ' EventTime
    rowData(5) = eventType                       ' EventType
    rowData(6) = message                         ' Message
    rowData(7) = details                         ' Details
    rowData(8) = GetCurrentOperatorName()        ' OperatorName
    rowData(9) = "Ne"                            ' Stornirano
    
    newRowIndex = AppendRow("tblSEFEventLog", rowData)
    
    If newRowIndex <= 0 Then
        Err.Raise ERR_SEF_STATE, "AppendSEFEvent_Row", _
            "Could not append row to tblSEFEventLog."
    End If
End Sub

' =========================
' INTERNAL HELPERS
' =========================

Private Function GetSingleRowIndexByKey( _
    ByVal tblName As String, _
    ByVal keyColName As String, _
    ByVal keyValue As Variant, _
    Optional ByVal raiseIfNotFound As Boolean = False) As Long
    
    Dim rowsFound As Collection
    
    Set rowsFound = FindRows(tblName, keyColName, keyValue)
    
    If rowsFound.count = 0 Then
        If raiseIfNotFound Then
            Err.Raise ERR_SEF_STATE, "GetSingleRowIndexByKey", _
                "Row not found in " & tblName & " for " & keyColName & "=" & CStr(keyValue)
        End If
        GetSingleRowIndexByKey = 0
        Exit Function
    End If
    
    If rowsFound.count > 1 Then
        Err.Raise ERR_SEF_DUPLICATE, "GetSingleRowIndexByKey", _
            "Multiple rows found in " & tblName & " for " & keyColName & "=" & CStr(keyValue)
    End If
    
    GetSingleRowIndexByKey = CLng(rowsFound(1))
End Function
Public Sub UpdateSEFLastSyncAt_Row(ByVal fakturaID As String)

    Dim rowIndex As Long
    Dim ok As Boolean
    
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)
    
    ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFLastSyncAt", Now)
    
    If Not ok Then
        RaiseUpdateError TBL_FAKTURE, "SEFLastSyncAt", fakturaID
    End If

End Sub

Private Sub RaiseUpdateError(ByVal tblName As String, ByVal colName As String, ByVal keyValue As String)
    Err.Raise ERR_SEF_STATE, "modSEFPersistence", _
        "Failed to update " & tblName & "." & colName & " for key " & keyValue
End Sub

Private Function GetCurrentOperatorName() As String
    On Error Resume Next
    
    GetCurrentOperatorName = Environ$("Username")
    
    If Len(Trim$(GetCurrentOperatorName)) = 0 Then
        GetCurrentOperatorName = Application.UserName
    End If
    
    If Len(Trim$(GetCurrentOperatorName)) = 0 Then
        GetCurrentOperatorName = "UNKNOWN"
    End If
End Function


Public Function GetSEFSubmissionsForFaktura(ByVal fakturaID As String) As Variant
    
    Dim data As Variant
    Dim filters As Collection
    Dim fp As clsFilterParam
    
    data = GetTableData("tblSEFSubmission")
    If IsEmpty(data) Then
        GetSEFSubmissionsForFaktura = Empty
        Exit Function
    End If
    
    Set filters = New Collection
    
    Set fp = New clsFilterParam
    filters.Add fp.Init(GetColumnIndex("tblSEFSubmission", "FakturaID"), "=", fakturaID)
    
    GetSEFSubmissionsForFaktura = FilterArray(data, filters)

End Function

Public Function GetSEFEventsForFaktura(ByVal fakturaID As String) As Variant
    
    Dim data As Variant
    Dim filters As Collection
    Dim fp As clsFilterParam
    
    data = GetTableData("tblSEFEventLog")
    If IsEmpty(data) Then
        GetSEFEventsForFaktura = Empty
        Exit Function
    End If
    
    Set filters = New Collection
    
    Set fp = New clsFilterParam
    filters.Add fp.Init(GetColumnIndex("tblSEFEventLog", "FakturaID"), "=", fakturaID)
    
    GetSEFEventsForFaktura = FilterArray(data, filters)

End Function

Public Function HasSuccessfulSEFSubmission(ByVal fakturaID As String) As Boolean
    
    Dim data As Variant
    Dim filters As Collection
    Dim fp As clsFilterParam
    Dim colStatus As Long
    Dim i As Long
    
    data = GetSEFSubmissionsForFaktura(fakturaID)
    If IsEmpty(data) Then
        HasSuccessfulSEFSubmission = False
        Exit Function
    End If
    
    colStatus = GetColumnIndex("tblSEFSubmission", "SubmissionStatus")
    
    For i = 1 To UBound(data, 1)
        Select Case Trim$(CStr(data(i, colStatus)))
            Case SEF_SUB_SENT, SEF_SUB_ACCEPTED
                HasSuccessfulSEFSubmission = True
                Exit Function
        End Select
    Next i
    
    HasSuccessfulSEFSubmission = False

End Function

Public Function GetLastSEFSubmissionStatus(ByVal fakturaID As String) As String
    
    Dim data As Variant
    Dim colCreatedAt As Long
    Dim colStatus As Long
    
    data = GetSEFSubmissionsForFaktura(fakturaID)
    If IsEmpty(data) Then
        GetLastSEFSubmissionStatus = ""
        Exit Function
    End If
    
    colCreatedAt = GetColumnIndex("tblSEFSubmission", "CreatedAt")
    colStatus = GetColumnIndex("tblSEFSubmission", "SubmissionStatus")
    
    data = SortArray(data, colCreatedAt, False)
    
    GetLastSEFSubmissionStatus = Trim$(CStr(data(1, colStatus)))

End Function

Public Function GetSubmissionRequestBody(ByVal submissionID As String) As String
    
    Dim v As Variant
    v = LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "RequestBody")
    
    If IsEmpty(v) Then
        GetSubmissionRequestBody = ""
    Else
        GetSubmissionRequestBody = CStr(v)
    End If
End Function

Public Function GetSubmissionPayloadHash(ByVal submissionID As String) As String
    
    Dim v As Variant
    v = LookupValue("tblSEFSubmission", "SEFSubmissionID", submissionID, "PayloadHash")
    
    If IsEmpty(v) Then
        GetSubmissionPayloadHash = ""
    Else
        GetSubmissionPayloadHash = Trim$(CStr(v))
    End If
End Function

Public Sub ClearFakturaLastSubmission_Row(ByVal fakturaID As String)
    
    Dim rowIndex As Long
    Dim ok As Boolean
    
    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, "ClearFakturaLastSubmission_Row", _
            "FakturaID is required."
    End If
    
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)
    
    ok = UpdateCell(TBL_FAKTURE, rowIndex, "SEFSubmissionIDLast", "")
    If Not ok Then RaiseUpdateError TBL_FAKTURE, "SEFSubmissionIDLast", fakturaID

End Sub

