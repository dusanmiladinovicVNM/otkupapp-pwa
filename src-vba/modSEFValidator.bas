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
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaForSEF"

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
    Dim iznosRaw As String
    Dim iznosValue As Double

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is required."
    End If

    fakture = GetTableData(TBL_FAKTURE)

    If IsEmpty(fakture) Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "TBL_FAKTURE is empty."
    End If

    colFakturaID = RequireColumnIndex(TBL_FAKTURE, "FakturaID", SRC)
    colKupacID = RequireColumnIndex(TBL_FAKTURE, "KupacID", SRC)
    colWorkflow = RequireColumnIndex(TBL_FAKTURE, "SEFWorkflowState", SRC)
    colBrojFakture = RequireColumnIndex(TBL_FAKTURE, "BrojFakture", SRC)
    colIznos = RequireColumnIndex(TBL_FAKTURE, "Iznos", SRC)

    For i = 1 To UBound(fakture, 1)
        If CStr(fakture(i, colFakturaID)) = fakturaID Then
            found = True
            kupacID = Trim$(CStr(fakture(i, colKupacID)))
            workflowState = Trim$(CStr(fakture(i, colWorkflow)))
            brojFakture = Trim$(CStr(fakture(i, colBrojFakture)))
            iznosRaw = Trim$(CStr(fakture(i, colIznos)))
            Exit For
        End If
    Next i

    If Not found Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "Faktura not found: " & fakturaID
    End If

    If Len(kupacID) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "KupacID is missing for faktura " & fakturaID
    End If

    If Len(brojFakture) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "BrojFakture is missing for faktura " & fakturaID
    End If

    If Not TryParseDouble(iznosRaw, iznosValue) Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "UkupanIznos is not numeric for faktura " & fakturaID
    End If

    If iznosValue <= 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "UkupanIznos must be > 0 for faktura " & fakturaID
    End If

    Select Case workflowState

        Case WF_LOCAL_FINALIZED, WF_SEF_READY, WF_SEF_TECH_FAILED
            ' allowed

        Case WF_SEF_ACCEPTED
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura already accepted on SEF."

        Case WF_SEF_SENDING
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura is already in SEF_SENDING state."

        Case WF_SEF_SENT
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura already sent. Refresh status first."

        Case WF_SEF_REJECTED
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura was rejected. Correction flow required."

        Case Else
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Faktura is not in a sendable state: " & workflowState
    End Select

    If HasSuccessfulSEFSubmission(fakturaID) Then
        Err.Raise ERR_SEF_DUPLICATE, SRC, _
                  "Faktura already has a successful SEF submission."
    End If

    ValidateFakturaHasStavke fakturaID
    ValidateKupacForSEF kupacID
    ValidateSEFConfig

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Sub


Private Sub ValidateFakturaHasStavke(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaHasStavke"

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURA_STAVKE, "FakturaID", SRC

    Dim rowsFound As Collection
    Set rowsFound = FindRows(TBL_FAKTURA_STAVKE, "FakturaID", fakturaID)

    If rowsFound Is Nothing Or rowsFound.count = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Faktura has no stavke: " & fakturaID
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Sub

Public Sub ValidateSEFPayload(ByVal payload As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateSEFPayload"

    If Len(Trim$(payload)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "SEF payload is empty."
    End If

    If InStr(1, payload, "InvoiceNumber", vbTextCompare) = 0 _
       And InStr(1, payload, "<cbc:ID>", vbTextCompare) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "SEF payload does not contain an invoice identifier."
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Sub
Private Sub ValidateKupacForSEF(ByVal kupacID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateKupacForSEF"

    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, "KupacID is required."
    End If

    RequireColumnIndex TBL_KUPCI, "KupacID", SRC
    RequireColumnIndex TBL_KUPCI, "Naziv", SRC
    RequireColumnIndex TBL_KUPCI, "PIB", SRC

    Dim naziv As Variant
    Dim pib As Variant

    naziv = LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")
    pib = LookupValue(TBL_KUPCI, "KupacID", kupacID, "PIB")

    If IsEmpty(naziv) Or IsNull(naziv) Or Len(Trim$(CStr(naziv))) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Kupac naziv is missing for kupac " & kupacID
    End If

    If IsEmpty(pib) Or IsNull(pib) Or Len(Trim$(CStr(pib))) = 0 Then
        Err.Raise ERR_SEF_VALIDATION, SRC, _
                  "Kupac PIB is missing for kupac " & kupacID
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Sub

Private Sub ValidateSEFConfig()
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateSEFConfig"

    Dim baseUrl As String
    Dim apiKey As String

    baseUrl = Trim$(GetConfigValue("SEF_BASE_URL"))
    apiKey = Trim$(GetConfigValue("SEF_API_KEY"))

    If Len(baseUrl) = 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
                  "SEF_BASE_URL missing in tblSEFConfig."
    End If

    If Len(apiKey) = 0 Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
                  "SEF_API_KEY missing in tblSEFConfig."
    End If

    If InStr(1, baseUrl, "http://", vbTextCompare) <> 1 _
       And InStr(1, baseUrl, "https://", vbTextCompare) <> 1 Then
        Err.Raise ERR_SEF_CONFIG, SRC, _
                  "SEF_BASE_URL must start with http:// or https://."
    End If

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Sub

Private Function GetFakturaSEFStatusText(ByVal fakturaID As String, _
                                         ByVal sourceName As String) As String
    On Error GoTo EH

    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise ERR_SEF_STATE, sourceName, "FakturaID is required."
    End If

    RequireColumnIndex TBL_FAKTURE, "FakturaID", sourceName
    RequireColumnIndex TBL_FAKTURE, "SEFStatus", sourceName

    Dim v As Variant
    v = LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus")

    If IsEmpty(v) Or IsNull(v) Then
        GetFakturaSEFStatusText = ""
    Else
        GetFakturaSEFStatusText = UCase$(Trim$(CStr(v)))
    End If

    Exit Function

EH:
    LogErr sourceName
    Err.Raise Err.Number, sourceName, Err.Description
End Function

Public Sub ValidateFakturaCanBeCancelledOnSEF(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaCanBeCancelledOnSEF"

    Dim sefDocumentId As String
    Dim sefStatus As String

    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)

    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "No SEFDocumentId found for faktura " & fakturaID
    End If

    sefStatus = GetFakturaSEFStatusText(fakturaID, SRC)

    Select Case sefStatus
        Case "DRAFT", "NEW", "ERROR"
            ' allowed

        Case Else
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Invoice cannot be cancelled on SEF in status: " & sefStatus
    End Select

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Sub

Public Sub ValidateFakturaCanBeStorniranoOnSEF(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "modSEFValidator.ValidateFakturaCanBeStorniranoOnSEF"

    Dim sefDocumentId As String
    Dim sefStatus As String

    sefDocumentId = GetFakturaSEFDocumentId(fakturaID)

    If Len(Trim$(sefDocumentId)) = 0 Then
        Err.Raise ERR_SEF_STATE, SRC, _
                  "No SEFDocumentId found for faktura " & fakturaID
    End If

    sefStatus = GetFakturaSEFStatusText(fakturaID, SRC)

    Select Case sefStatus
        Case "SENT", "ACCEPTED", "REJECTED"
            ' allowed

        Case Else
            Err.Raise ERR_SEF_STATE, SRC, _
                      "Invoice cannot be storno on SEF in status: " & sefStatus
    End Select

    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
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
    
    Call ClearFakturaLastSubmission_Row(fakturaID)
    
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
