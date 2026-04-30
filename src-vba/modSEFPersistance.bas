Attribute VB_Name = "modSEFPersistance"
Option Explicit

' =========================================================
' modSEFPersistence
' Alle SEF-Reads/Writes laufen über modDataAccess
' =========================================================
Private Const TBL_SEF_SUBMISSION As String = "tblSEFSubmission"
Private Const TBL_SEF_EVENT_LOG As String = "tblSEFEventLog"

' =========================
' READ HELPERS
' =========================

Public Function GetFakturaSEFWorkflowState(ByVal fakturaID As String) As String
    GetFakturaSEFWorkflowState = GetFakturaSEFFieldText( _
        fakturaID, "SEFWorkflowState", "modSEFPersistance.GetFakturaSEFWorkflowState")
End Function

Public Function GetFakturaSEFDocumentId(ByVal fakturaID As String) As String
    GetFakturaSEFDocumentId = GetFakturaSEFFieldText( _
        fakturaID, "SEFDocumentId", "modSEFPersistance.GetFakturaSEFDocumentId")
End Function

Public Function GetLastSEFSubmissionID(ByVal fakturaID As String) As String
    GetLastSEFSubmissionID = GetFakturaSEFFieldText( _
        fakturaID, "SEFSubmissionIDLast", "modSEFPersistance.GetLastSEFSubmissionID")
End Function

Public Function GetNextSEFVersionNo(ByVal fakturaID As String) As Long
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetNextSEFVersionNo"

    Dim rawValue As String
    Dim currentVersion As Long

    rawValue = GetFakturaSEFFieldText(fakturaID, "SEFVersionNo", SRC)

    If rawValue = "" Then
        GetNextSEFVersionNo = 1
    ElseIf Not TryParseLong(rawValue, currentVersion) Then
        GetNextSEFVersionNo = 1
    Else
        GetNextSEFVersionNo = currentVersion + 1
    End If

    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Public Function GetCurrentSEFVersionNo(ByVal fakturaID As String) As Long
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetCurrentSEFVersionNo"

    Dim rawValue As String
    Dim currentVersion As Long

    rawValue = GetFakturaSEFFieldText(fakturaID, "SEFVersionNo", SRC)

    If rawValue = "" Then
        GetCurrentSEFVersionNo = 0
    ElseIf Not TryParseLong(rawValue, currentVersion) Then
        GetCurrentSEFVersionNo = 0
    Else
        GetCurrentSEFVersionNo = currentVersion
    End If

    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
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

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.UpdateFakturaSEFState_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    If Len(Trim$(newState)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "newState is required."
    End If

    RequireFaktureSEFSchema SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    Dim oldState As String
    oldState = GetFakturaSEFWorkflowState(fakturaID)

    If Len(oldState) > 0 Then
        ValidateAllowedTransition oldState, newState
    End If

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFWorkflowState", newState, SRC

    If Len(sefStatus) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFStatus", sefStatus, SRC
    End If

    If Len(sefDocumentId) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFDocumentId", sefDocumentId, SRC
    End If

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorCode", errorCode, SRC
    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorMessage", errorMessage, SRC

    If Len(payloadHash) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFPayloadHash", payloadHash, SRC
    End If

    If Len(submissionID) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFSubmissionIDLast", submissionID, SRC
    End If

    If versionNo > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFVersionNo", versionNo, SRC
    End If

    Select Case newState
        Case WF_SEF_SENT, WF_SEF_ACCEPTED
            RequireUpdateCell TBL_FAKTURE, rowIndex, "PoslatNaSEF", "Da", SRC

            If Len(Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFSentAt")))) = 0 Then
                RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFSentAt", Now, SRC
            End If

        Case WF_SEF_SENDING
            RequireUpdateCell TBL_FAKTURE, rowIndex, "PoslatNaSEF", "Ne", SRC
    End Select

    Select Case newState
        Case WF_SEF_SENT, WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_SYNC_ERROR
            RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastSyncAt", Now, SRC
    End Select

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
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

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.UpdateFakturaSEFRefreshFields_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    RequireFaktureSEFSchema SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    If Len(sefStatus) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFStatus", sefStatus, SRC
    End If

    If Len(sefDocumentId) > 0 Then
        RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFDocumentId", sefDocumentId, SRC
    End If

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorCode", errorCode, SRC
    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastErrorMessage", errorMessage, SRC

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Public Function CreateSEFSubmission_Row( _
    ByVal fakturaID As String, _
    ByVal versionNo As Long, _
    ByVal workflowState As String, _
    ByVal payloadHash As String, _
    ByVal requestBody As String, _
    ByVal requestFormat As String) As String

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.CreateSEFSubmission_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    If versionNo <= 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "VersionNo must be > 0."
    End If

    If Len(Trim$(workflowState)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "WorkflowStateAtSubmit is required."
    End If

    If Len(Trim$(requestBody)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "RequestBody is required."
    End If

    If Len(Trim$(requestFormat)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "RequestFormat is required."
    End If

    RequireSEFSubmissionSchema SRC

    Dim submissionID As String
    submissionID = GetNextID(TBL_SEF_SUBMISSION, "SEFSubmissionID", "SFS-")

    If Len(Trim$(submissionID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "GetNextID did not return SEFSubmissionID."
    End If

    Dim rowData(1 To 20) As Variant

    rowData(1) = submissionID
    rowData(2) = fakturaID
    rowData(3) = versionNo
    rowData(4) = workflowState
    rowData(5) = Now
    rowData(6) = Empty
    rowData(7) = SEF_SUB_CREATED
    rowData(8) = payloadHash
    rowData(9) = requestFormat
    rowData(10) = requestBody
    rowData(11) = Empty
    rowData(12) = Empty
    rowData(13) = Empty
    rowData(14) = Empty
    rowData(15) = Empty
    rowData(16) = Empty
    rowData(17) = Empty
    rowData(18) = GetCurrentOperatorName()
    rowData(19) = "Ne"
    rowData(20) = Empty

    Dim newRowIndex As Long
    newRowIndex = AppendRow(TBL_SEF_SUBMISSION, rowData)

    If newRowIndex <= 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Could not append row to tblSEFSubmission."
    End If

    CreateSEFSubmission_Row = submissionID
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Public Sub SaveSEFSubmissionResult_Row( _
    ByVal submissionID As String, _
    ByVal response As clsSEFResponse)

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.SaveSEFSubmissionResult_Row"

    If Len(Trim$(submissionID)) = 0 Then Exit Sub

    If response Is Nothing Then
        Err.Raise ERR_SEF_RESPONSE_PARSE, SRC, _
                  "Response object is Nothing."
    End If

    RequireSEFSubmissionSchema SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, True)

    Dim subStatus As String

    If response.Accepted Then
        subStatus = SEF_SUB_ACCEPTED
    ElseIf response.Rejected Then
        subStatus = SEF_SUB_REJECTED
    ElseIf response.Success Then
        subStatus = SEF_SUB_SENT
    Else
        subStatus = SEF_SUB_FAILED
    End If

    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "SubmittedAt", Now, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "FinishedAt", Now, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "SubmissionStatus", subStatus, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "HttpStatus", response.httpStatus, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ApiStatus", response.apiStatus, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "CorrelationId", response.CorrelationId, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ResponseBody", response.rawBody, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "SEFDocumentId", response.sefDocumentId, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ErrorCode", response.errorCode, SRC
    RequireUpdateCell TBL_SEF_SUBMISSION, rowIndex, "ErrorMessage", response.errorMessage, SRC

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Public Sub AppendSEFEvent_Row( _
    ByVal fakturaID As String, _
    ByVal submissionID As String, _
    ByVal eventType As String, _
    ByVal message As String, _
    Optional ByVal details As String = "")

    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.AppendSEFEvent_Row"

    RequireSEFEventLogSchema SRC

    Dim eventID As String
    eventID = GetNextID(TBL_SEF_EVENT_LOG, "SEFEventID", "SFE-")

    If Len(Trim$(eventID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "GetNextID did not return SEFEventID."
    End If

    Dim rowData(1 To 9) As Variant

    rowData(1) = eventID
    rowData(2) = fakturaID
    rowData(3) = submissionID
    rowData(4) = Now
    rowData(5) = eventType
    rowData(6) = message
    rowData(7) = details
    rowData(8) = GetCurrentOperatorName()
    rowData(9) = "Ne"

    Dim newRowIndex As Long
    newRowIndex = AppendRow(TBL_SEF_EVENT_LOG, rowData)

    If newRowIndex <= 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "Could not append row to tblSEFEventLog."
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
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
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.UpdateSEFLastSyncAt_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", SRC
    RequireColumnIndex TBL_FAKTURE, "SEFLastSyncAt", SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFLastSyncAt", Now, SRC

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
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
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSEFSubmissionsForFaktura"

    Dim data As Variant
    data = GetTableData(TBL_SEF_SUBMISSION)

    If IsEmpty(data) Then
        GetSEFSubmissionsForFaktura = Empty
        Exit Function
    End If

    Dim filters As Collection
    Dim fp As clsFilterParam

    Set filters = New Collection
    Set fp = New clsFilterParam

    filters.Add fp.Init(RequireColumnIndex(TBL_SEF_SUBMISSION, "FakturaID", SRC), "=", fakturaID)

    GetSEFSubmissionsForFaktura = FilterArray(data, filters)
    Exit Function

EH:
    LogErr SRC
    GetSEFSubmissionsForFaktura = Empty
End Function

Public Function GetSEFEventsForFaktura(ByVal fakturaID As String) As Variant
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSEFEventsForFaktura"

    Dim data As Variant
    data = GetTableData(TBL_SEF_EVENT_LOG)

    If IsEmpty(data) Then
        GetSEFEventsForFaktura = Empty
        Exit Function
    End If

    Dim filters As Collection
    Dim fp As clsFilterParam

    Set filters = New Collection
    Set fp = New clsFilterParam

    filters.Add fp.Init(RequireColumnIndex(TBL_SEF_EVENT_LOG, "FakturaID", SRC), "=", fakturaID)

    GetSEFEventsForFaktura = FilterArray(data, filters)
    Exit Function

EH:
    LogErr SRC
    GetSEFEventsForFaktura = Empty
End Function

Public Function HasSuccessfulSEFSubmission(ByVal fakturaID As String) As Boolean
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.HasSuccessfulSEFSubmission"

    Dim data As Variant
    data = GetSEFSubmissionsForFaktura(fakturaID)

    If IsEmpty(data) Then
        HasSuccessfulSEFSubmission = False
        Exit Function
    End If

    Dim colStatus As Long
    colStatus = RequireColumnIndex(TBL_SEF_SUBMISSION, "SubmissionStatus", SRC)

    Dim i As Long

    For i = 1 To UBound(data, 1)
        Select Case Trim$(CStr(data(i, colStatus)))
            Case SEF_SUB_SENT, SEF_SUB_ACCEPTED
                HasSuccessfulSEFSubmission = True
                Exit Function
        End Select
    Next i

    HasSuccessfulSEFSubmission = False
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Public Function GetLastSEFSubmissionStatus(ByVal fakturaID As String) As String
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetLastSEFSubmissionStatus"

    Dim data As Variant
    data = GetSEFSubmissionsForFaktura(fakturaID)

    If IsEmpty(data) Then
        GetLastSEFSubmissionStatus = ""
        Exit Function
    End If

    Dim colCreatedAt As Long
    Dim colStatus As Long

    colCreatedAt = RequireColumnIndex(TBL_SEF_SUBMISSION, "CreatedAt", SRC)
    colStatus = RequireColumnIndex(TBL_SEF_SUBMISSION, "SubmissionStatus", SRC)

    data = SortArray(data, colCreatedAt, False)

    GetLastSEFSubmissionStatus = Trim$(CStr(data(1, colStatus)))
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Public Function GetSubmissionRequestBody(ByVal submissionID As String) As String
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSubmissionRequestBody"

    If Len(Trim$(submissionID)) = 0 Then Exit Function

    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFSubmissionID", SRC
    RequireColumnIndex TBL_SEF_SUBMISSION, "RequestBody", SRC

    Dim v As Variant
    v = LookupValue(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, "RequestBody")

    If IsEmpty(v) Or IsNull(v) Then
        GetSubmissionRequestBody = ""
    Else
        GetSubmissionRequestBody = CStr(v)
    End If

    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Public Function GetSubmissionPayloadHash(ByVal submissionID As String) As String
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.GetSubmissionPayloadHash"

    If Len(Trim$(submissionID)) = 0 Then Exit Function

    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFSubmissionID", SRC
    RequireColumnIndex TBL_SEF_SUBMISSION, "PayloadHash", SRC

    Dim v As Variant
    v = LookupValue(TBL_SEF_SUBMISSION, "SEFSubmissionID", submissionID, "PayloadHash")

    If IsEmpty(v) Or IsNull(v) Then
        GetSubmissionPayloadHash = ""
    Else
        GetSubmissionPayloadHash = Trim$(CStr(v))
    End If

    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Public Sub ClearFakturaLastSubmission_Row(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFPersistance.ClearFakturaLastSubmission_Row"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", SRC
    RequireColumnIndex TBL_FAKTURE, "SEFSubmissionIDLast", SRC

    Dim rowIndex As Long
    rowIndex = GetSingleRowIndexByKey(TBL_FAKTURE, "FakturaID", fakturaID, True)

    RequireUpdateCell TBL_FAKTURE, rowIndex, "SEFSubmissionIDLast", "", SRC

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Private Sub RequireFaktureSEFSchema(ByVal sourceName As String)
    RequireColumnIndex TBL_FAKTURE, "FakturaID", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFWorkflowState", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFStatus", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFDocumentId", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFLastErrorCode", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFLastErrorMessage", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFPayloadHash", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFSubmissionIDLast", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFVersionNo", sourceName
    RequireColumnIndex TBL_FAKTURE, "PoslatNaSEF", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFSentAt", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFLastSyncAt", sourceName
End Sub

Private Sub RequireSEFSubmissionSchema(ByVal sourceName As String)
    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFSubmissionID", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "FakturaID", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "VersionNo", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "WorkflowStateAtSubmit", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "CreatedAt", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "SubmittedAt", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "SubmissionStatus", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "PayloadHash", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "RequestFormat", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "RequestBody", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ResponseBody", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "HttpStatus", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ApiStatus", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "CorrelationId", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "SEFDocumentId", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ErrorCode", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "ErrorMessage", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "OperatorName", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "Stornirano", sourceName
    RequireColumnIndex TBL_SEF_SUBMISSION, "FinishedAt", sourceName
End Sub

Private Sub RequireSEFEventLogSchema(ByVal sourceName As String)
    RequireColumnIndex TBL_SEF_EVENT_LOG, "SEFEventID", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "FakturaID", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "SEFSubmissionID", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "EventTime", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "EventType", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "Message", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "Details", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "OperatorName", sourceName
    RequireColumnIndex TBL_SEF_EVENT_LOG, "Stornirano", sourceName
End Sub

Private Function GetFakturaSEFFieldText(ByVal fakturaID As String, _
                                        ByVal fieldName As String, _
                                        ByVal sourceName As String) As String
    On Error GoTo EH

    If Len(Trim$(fakturaID)) = 0 Then Exit Function

    RequireColumnIndex TBL_FAKTURE, "FakturaID", sourceName
    RequireColumnIndex TBL_FAKTURE, fieldName, sourceName

    Dim v As Variant
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, fieldName)

    If IsEmpty(v) Or IsNull(v) Then
        GetFakturaSEFFieldText = ""
    Else
        GetFakturaSEFFieldText = Trim$(CStr(v))
    End If

    Exit Function

EH:
    LogErr sourceName
    Err.Raise Err.Number, sourceName, Err.description
End Function

