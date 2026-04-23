Attribute VB_Name = "modSEFValidator"
Option Explicit

Public Sub ValidateAllowedTransition(ByVal oldState As String, ByVal newState As String)
    
    Select Case oldState
        
        Case WF_LOCAL_DRAFT
            If newState <> WF_LOCAL_FINALIZED Then GoTo InvalidTransition
        
        Case WF_LOCAL_FINALIZED
            If newState <> WF_SEF_READY Then GoTo InvalidTransition
        
        Case WF_SEF_READY
            If newState <> WF_SEF_SENDING Then GoTo InvalidTransition
            
        Case WF_SEF_SENDING
            Select Case newState
                Case WF_SEF_SENT, WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_TECH_FAILED, WF_SEF_UNKNOWN
            Case Else
                GoTo InvalidTransition
        End Select
        
        Case WF_SEF_SENT
            Select Case newState
                Case WF_SEF_ACCEPTED, WF_SEF_REJECTED, WF_SEF_SYNC_ERROR, WF_SEF_STORNO
                Case Else
                    GoTo InvalidTransition
            End Select
        
        Case WF_SEF_TECH_FAILED
            If newState <> WF_SEF_READY Then GoTo InvalidTransition
        
        Case WF_SEF_SYNC_ERROR
            If newState <> WF_SEF_SENT Then GoTo InvalidTransition
        
        Case WF_SEF_ACCEPTED
            If newState <> WF_SEF_STORNO Then GoTo InvalidTransition

        Case WF_SEF_REJECTED
            If newState <> WF_SEF_READY Then GoTo InvalidTransition
        
        Case Else
            Err.Raise ERR_SEF_STATE, "ValidateAllowedTransition", _
                "Unknown current workflow state: " & oldState
    End Select
    
    Exit Sub

InvalidTransition:
    Err.Raise ERR_SEF_STATE, "ValidateAllowedTransition", _
        "Illegal SEF state transition: " & oldState & " -> " & newState
End Sub

Public Sub ValidateFakturaForSEF(ByVal fakturaID As String)
    
    Dim fakture As Variant
    Dim i As Long
    
    Dim colFakturaID As Long
    Dim colKupacID As Long
    Dim colWorkflow As Long
    Dim colBrojFakture As Long
    Dim colIznos As Long
    
    Dim found As Boolean
    Dim kupacID As String
    Dim workflowState As String
    Dim brojFakture As String
    Dim iznos As Variant
    
    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", "FakturaID is required."
    End If
    
    fakture = GetTableData(TBL_FAKTURE)
    If IsEmpty(fakture) Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", "TBL_FAKTURE is empty."
    End If
    
    colFakturaID = GetColumnIndex(TBL_FAKTURE, "FakturaID")
    colKupacID = GetColumnIndex(TBL_FAKTURE, "KupacID")
    colWorkflow = GetColumnIndex(TBL_FAKTURE, "SEFWorkflowState")
    colBrojFakture = GetColumnIndex(TBL_FAKTURE, "BrojFakture")
    colIznos = GetColumnIndex(TBL_FAKTURE, "Iznos")
    
    If colFakturaID = 0 Or colKupacID = 0 Or colWorkflow = 0 Or colBrojFakture = 0 Or colIznos = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", _
            "Required columns missing in tblFakture."
    End If
    
    For i = 1 To UBound(fakture, 1)
        If CStr(fakture(i, colFakturaID)) = fakturaID Then
            found = True
            kupacID = Trim$(CStr(fakture(i, colKupacID)))
            workflowState = Trim$(CStr(fakture(i, colWorkflow)))
            brojFakture = Trim$(CStr(fakture(i, colBrojFakture)))
            iznos = fakture(i, colIznos)
            Exit For
        End If
    Next i
    
    If Not found Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", _
            "Faktura not found: " & fakturaID
    End If
    
    If Len(kupacID) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", _
            "KupacID is missing for faktura " & fakturaID
    End If
    
    If Len(brojFakture) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", _
            "BrojFakture is missing for faktura " & fakturaID
    End If
    
    If Not IsNumeric(iznos) Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", _
            "UkupanIznos is not numeric for faktura " & fakturaID
    End If
    
    If CDbl(iznos) <= 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaForSEF", _
            "UkupanIznos must be > 0 for faktura " & fakturaID
    End If
    
    Select Case workflowState
        Case WF_LOCAL_FINALIZED, WF_SEF_READY, WF_SEF_TECH_FAILED
            ' allowed
        
        Case WF_SEF_ACCEPTED
            Err.Raise ERR_SEF_STATE, "ValidateFakturaForSEF", _
                "Faktura already accepted on SEF."
        
        Case WF_SEF_SENDING
            Err.Raise ERR_SEF_STATE, "ValidateFakturaForSEF", _
                "Faktura is already in SEF_SENDING state."
        
        Case WF_SEF_SENT
            Err.Raise ERR_SEF_STATE, "ValidateFakturaForSEF", _
                "Faktura already sent. Refresh status first."
        
        Case WF_SEF_REJECTED
            Err.Raise ERR_SEF_STATE, "ValidateFakturaForSEF", _
                "Faktura was rejected. Correction flow required."
        
        Case Else
            Err.Raise ERR_SEF_STATE, "ValidateFakturaForSEF", _
                "Faktura is not in a sendable state: " & workflowState
    End Select
    
    If HasSuccessfulSEFSubmission(fakturaID) Then
        Err.Raise ERR_SEF_DUPLICATE, "ValidateFakturaForSEF", _
            "Faktura already has a successful SEF submission."
    End If
    
    Call ValidateFakturaHasStavke(fakturaID)
    Call ValidateKupacForSEF(kupacID)
    Call ValidateSEFConfig

End Sub


Private Sub ValidateFakturaHasStavke(ByVal fakturaID As String)
    
    Dim rowsFound As Collection
    Set rowsFound = FindRows("tblFakturaStavke", "FakturaID", fakturaID)
    
    If rowsFound Is Nothing Or rowsFound.count = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateFakturaHasStavke", _
            "Faktura has no stavke: " & fakturaID
    End If

End Sub

Public Sub ValidateSEFPayload(ByVal payload As String)
    
    If Len(Trim$(payload)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateSEFPayload", "SEF payload is empty."
    End If
    
    If InStr(1, payload, "InvoiceNumber", vbTextCompare) = 0 _
       And InStr(1, payload, "<cbc:ID>", vbTextCompare) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateSEFPayload", _
            "SEF payload does not contain an invoice identifier."
    End If

End Sub

Private Sub ValidateKupacForSEF(ByVal kupacID As String)
    
    Dim Naziv As Variant
    Dim pib As Variant
    
    Naziv = LookupValue("tblKupci", "KupacID", kupacID, "Naziv")
    pib = LookupValue("tblKupci", "KupacID", kupacID, "PIB")
    
    If IsEmpty(Naziv) Or Len(Trim$(CStr(Naziv))) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateKupacForSEF", _
            "Kupac naziv is missing for kupac " & kupacID
    End If
    
    If IsEmpty(pib) Or Len(Trim$(CStr(pib))) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, "ValidateKupacForSEF", _
            "Kupac PIB is missing for kupac " & kupacID
    End If

End Sub

Private Sub ValidateSEFConfig()
    
    ' Später gegen echte Config-Keys austauschen, sobald dein modConfig-Zugriff final ist.
    ' Hier nur Struktur:
    
    ' Beispiel:
    ' If Len(Trim$(GetConfigValue("SEF_BASE_URL"))) = 0 Then
    '     Err.Raise ERR_SEF_CONFIG, "ValidateSEFConfig", "SEF_BASE_URL missing."
    ' End If
    
End Sub

Public Sub ValidateFakturaCanBeCancelledOnSEF(ByVal fakturaID As String)
    
    Dim sefDocumentId As String
    Dim sefStatus As String
    
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, "ValidateFakturaCanBeCancelledOnSEF", _
            "No SEFDocumentId found for faktura " & fakturaID
    End If
    
    sefStatus = UCase$(Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))))
    
    Select Case sefStatus
        Case "DRAFT", "NEW", "ERROR"
            ' allowed
        Case Else
            Err.Raise ERR_SEF_STATE, "ValidateFakturaCanBeCancelledOnSEF", _
                "Invoice cannot be cancelled on SEF in status: " & sefStatus
    End Select
End Sub

Public Sub ValidateFakturaCanBeStorniranoOnSEF(ByVal fakturaID As String)
    
    Dim sefDocumentId As String
    Dim sefStatus As String
    
    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)
    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, "ValidateFakturaCanBeStorniranoOnSEF", _
            "No SEFDocumentId found for faktura " & fakturaID
    End If
    
    sefStatus = UCase$(Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))))
    
    Select Case sefStatus
        Case "SENT", "ACCEPTED", "REJECTED"
            ' allowed
        Case Else
            Err.Raise ERR_SEF_STATE, "ValidateFakturaCanBeStorniranoOnSEF", _
                "Invoice cannot be storno on SEF in status: " & sefStatus
    End Select
End Sub

Public Sub PrepareRejectedInvoiceForResubmit(ByVal fakturaID As String)
    
    Dim tx As clsTransaction
    Dim currentState As String
    
    On Error GoTo EH
    
    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, "PrepareRejectedInvoiceForResubmit", _
            "FakturaID is required."
    End If
    
    currentState = GetFakturaSEFWorkflowState(fakturaID)
    
    If currentState <> WF_SEF_REJECTED Then
        Err.Raise ERR_SEF_STATE, "PrepareRejectedInvoiceForResubmit", _
            "Invoice is not in SEF_REJECTED state: " & currentState
    End If
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot "tblSEFEventLog"
    
    Call UpdateFakturaSEFState_Row( _
        fakturaID:=fakturaID, _
        newState:=WF_SEF_READY, _
        sefStatus:=WF_SEF_READY, _
        errorCode:="", _
        errorMessage:="", _
        submissionID:="")
    
    Call AppendSEFEvent_Row( _
        fakturaID:=fakturaID, _
        submissionID:="", _
        eventType:=SEF_EVT_STATE_CHANGED, _
        message:="Rejected invoice prepared for corrected resubmission.", _
        details:="PreviousState=" & currentState)
    
    tx.CommitTx
    Exit Sub

EH:
    LogErr "PrepareRejectedInvoiceForResubmit"
    Dim errNum As Long
    Dim errDesc As String
    
    errNum = Err.Number
    errDesc = Err.Description
    
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    
    If errNum <> 0 Then
        Err.Raise errNum, "PrepareRejectedInvoiceForResubmit", errDesc
    Else
        Err.Raise ERR_SEF_STATE, "PrepareRejectedInvoiceForResubmit", _
            "Unexpected error preparing rejected invoice."
    End If
End Sub

Public Function IsFinalSEFStatus(ByVal sefStatus As String) As Boolean
    
    Select Case UCase$(Trim$(sefStatus))
        Case "ACCEPTED", "REJECTED", "STORNO", "CANCELLED"
            IsFinalSEFStatus = True
        Case Else
            IsFinalSEFStatus = False
    End Select
    
End Function

Public Function IsPendingSEFStatus(ByVal sefStatus As String) As Boolean
    
    Select Case UCase$(Trim$(sefStatus))
        Case "SENT", "NEW", "DRAFT"
            IsPendingSEFStatus = True
        Case Else
            IsPendingSEFStatus = False
    End Select
    
End Function

Public Function GetSEFDisplayStatus(ByVal workflowState As String, ByVal sefStatus As String) As String
    
    If Len(Trim$(sefStatus)) > 0 Then
        GetSEFDisplayStatus = Trim$(sefStatus)
    Else
        GetSEFDisplayStatus = Trim$(workflowState)
    End If
    
End Function
