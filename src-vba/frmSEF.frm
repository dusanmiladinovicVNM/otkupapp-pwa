VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSEF 
   Caption         =   "UserForm1"
   ClientHeight    =   9960.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15465
   OleObjectBlob   =   "frmSEF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_SetupDone As Boolean
Private mChromeRemoved As Boolean

Private Sub RemoveTitleBar()
    Dim hwnd As LongPtr
    Dim style As Long

    hwnd = FindWindow("ThunderDFrame", Me.caption)

    If hwnd <> 0 Then
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH

    If Not mChromeRemoved Then
        Me.caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If

    ApplyTheme Me, BG_MAIN

    If m_SetupDone Then Exit Sub
    m_SetupDone = True

    Me.caption = "SEF upravljanje"

    Call SetupSEFEventList
    Call LoadFaktureIntoCombo
    Call ClearSEFInfo
    
    Call SetupHelpPage

    Exit Sub

EH:
    LogErr "frmSEF.UserForm_Activate"
    MsgBox "Greška pri otvaranju SEF forme: " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub SetupSEFEventList()
    
    With Me.lstSEFEvents
        .ColumnCount = 4
        .ColumnWidths = "95;80;220;260"
        .MultiSelect = fmMultiSelectSingle
    End With
    
End Sub

Private Sub LoadFaktureIntoCombo()
    
    Dim data As Variant
    Dim colFakturaID As Long
    Dim colBroj As Long
    Dim i As Long
    
    Me.cmbFaktura.Clear
    
    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then Exit Sub
    
    colFakturaID = GetColumnIndex(TBL_FAKTURE, "FakturaID")
    colBroj = GetColumnIndex(TBL_FAKTURE, "BrojFakture")
    
    If colFakturaID = 0 Or colBroj = 0 Then Exit Sub
    
    For i = 1 To UBound(data, 1)
        Me.cmbFaktura.AddItem CStr(data(i, colFakturaID))
        Me.cmbFaktura.List(Me.cmbFaktura.ListCount - 1, 1) = CStr(data(i, colBroj))
    Next i
    
End Sub

Private Function GetSelectedFakturaID() As String
    
    GetSelectedFakturaID = Trim$(CStr(Me.cmbFaktura.value))
    
End Function

Private Sub ClearSEFInfo()
    
    Me.lblFakturaID.caption = ""
    Me.lblBrojFakture.caption = ""
    Me.lblKupacNaziv.caption = ""
    Me.lblWorkflow.caption = ""
    Me.lblSEFStatus.caption = ""
    Me.lblSEFDocumentID.caption = ""
    Me.lblVersion.caption = ""
    Me.lblLastError.caption = ""
    
    Me.lstSEFEvents.Clear
    Call UpdateSEFButtonStates
    
End Sub

Private Sub LoadSelectedFakturaInfo()
    
    Dim fakturaID As String
    Dim kupacID As String
    
    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        Call ClearSEFInfo
        Exit Sub
    End If
    
    kupacID = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "KupacID"))
    
    Me.lblFakturaID.caption = fakturaID
    Me.lblBrojFakture.caption = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "BrojFakture"))
    Me.lblKupacNaziv.caption = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv"))
    Me.lblWorkflow.caption = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFWorkflowState"))
    Me.lblSEFStatus.caption = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))
    Me.lblSEFDocumentID.caption = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFDocumentId"))
    Me.lblVersion.caption = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFVersionNo"))
    Me.lblLastError.caption = CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFLastErrorMessage"))
    
    Call LoadSEFEventsForSelectedFaktura
    Call UpdateSEFButtonStates
    
    Select Case UCase$(Me.lblSEFStatus.caption)
    Case "SENT"
        Me.lblSEFStatus.foreColor = vbBlue
    Case "ACCEPTED"
        Me.lblSEFStatus.foreColor = vbGreen
    Case "REJECTED"
        Me.lblSEFStatus.foreColor = vbRed
    Case "CANCELLED", "STORNO"
        Me.lblSEFStatus.foreColor = RGB(128, 0, 128)
    Case Else
        Me.lblSEFStatus.foreColor = vbBlack
    End Select
    
End Sub

Private Sub LoadSEFEventsForSelectedFaktura()
    
    Dim fakturaID As String
    Dim data As Variant
    Dim colTime As Long
    Dim colType As Long
    Dim colMsg As Long
    Dim colDetails As Long
    Dim i As Long
    
    Me.lstSEFEvents.Clear
    
    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then Exit Sub
    
    data = GetSEFEventsForFaktura(fakturaID)
    If IsEmpty(data) Then Exit Sub
    
    colTime = GetColumnIndex("tblSEFEventLog", "EventTime")
    colType = GetColumnIndex("tblSEFEventLog", "EventType")
    colMsg = GetColumnIndex("tblSEFEventLog", "Message")
    colDetails = GetColumnIndex("tblSEFEventLog", "Details")
    
    For i = 1 To UBound(data, 1)
        Me.lstSEFEvents.AddItem CStr(data(i, colTime))
        Me.lstSEFEvents.List(Me.lstSEFEvents.ListCount - 1, 1) = CStr(data(i, colType))
        Me.lstSEFEvents.List(Me.lstSEFEvents.ListCount - 1, 2) = CStr(data(i, colMsg))
        Me.lstSEFEvents.List(Me.lstSEFEvents.ListCount - 1, 3) = CStr(data(i, colDetails))
    Next i
    
End Sub

Private Sub UpdateSEFButtonStates()
    
    Dim fakturaID As String
    Dim workflowState As String
    Dim sefStatus As String
    
    fakturaID = GetSelectedFakturaID()
    
    If Len(fakturaID) = 0 Then
        Me.btnPosalji.enabled = False
        Me.btnOsvezi.enabled = False
        Me.btnPrepareResubmit.enabled = False
        Me.btnCancel.enabled = False
        Me.btnStorno.enabled = False
        Me.btnRecoverSending.enabled = False
        Exit Sub
    End If
    
    workflowState = UCase$(Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFWorkflowState"))))
    sefStatus = UCase$(Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFStatus"))))
    
    Me.btnPosalji.enabled = (workflowState = UCase$(WF_LOCAL_FINALIZED) Or _
                             workflowState = UCase$(WF_SEF_READY) Or _
                             workflowState = UCase$(WF_SEF_TECH_FAILED))
    
    If workflowState = UCase$(WF_SEF_TECH_FAILED) Then
        Me.btnPosalji.caption = "Retry slanje na SEF"
    Else
        Me.btnPosalji.caption = "Pošalji na SEF"
    End If
    
    If Not Me.btnPosalji.enabled Then
        Me.btnPosalji.caption = "Pošalji na SEF"
    End If
    
    Me.btnOsvezi.enabled = (workflowState = UCase$(WF_SEF_SENT) Or _
                            workflowState = UCase$(WF_SEF_SYNC_ERROR))
    
    Me.btnPrepareResubmit.enabled = (workflowState = UCase$(WF_SEF_REJECTED))
    
    Me.btnCancel.enabled = (sefStatus = "DRAFT" Or sefStatus = "NEW" Or sefStatus = "ERROR")
    
    Me.btnStorno.enabled = (sefStatus = "SENT" Or sefStatus = "ACCEPTED" Or sefStatus = "REJECTED")
    
    Me.btnRecoverSending.enabled = (workflowState = UCase$(WF_SEF_SENDING))
    
End Sub

Private Sub btnUcitaj_Click()
    On Error GoTo EH
    
    Call LoadSelectedFakturaInfo
    Exit Sub

EH:
    LogErr "frmSEF.btnUcitaj"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnPosalji_Click()
    On Error GoTo EH

    Dim fakturaID As String
    Dim submissionID As String

    Me.btnPosalji.enabled = False
    DoEvents

    fakturaID = GetSelectedFakturaID()

    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        GoTo CleanExit
    End If

    If MsgBox("Poslati fakturu " & fakturaID & " na SEF?", _
              vbQuestion + vbYesNo, APP_NAME) = vbNo Then
        GoTo CleanExit
    End If

    submissionID = SendInvoiceToSEF_TX(fakturaID)

    Call LoadSelectedFakturaInfo

    MsgBox "Faktura poslata. SubmissionID: " & submissionID, vbInformation, APP_NAME

CleanExit:
    Me.btnPosalji.enabled = True
    Call UpdateSEFButtonStates
    Exit Sub

EH:
    LogErr "frmSEF.btnPosalji"
    MsgBox Err.Description, vbCritical, APP_NAME
    Resume CleanExit
End Sub

Private Sub btnOsvezi_Click()
    On Error GoTo EH
    
    Dim fakturaID As String
    
    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Call RefreshSEFStatus_TX(fakturaID)
    Call LoadSelectedFakturaInfo
    
    MsgBox "SEF status osvežen.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnOsvezi"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnPrepareResubmit_Click()
    On Error GoTo EH
    
    Dim fakturaID As String
    
    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If MsgBox("Pripremiti odbijenu fakturu za ponovno slanje?", vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Sub
    
    Call PrepareRejectedInvoiceForResubmit(fakturaID)
    Call LoadSelectedFakturaInfo
    
    MsgBox "Faktura je pripremljena za ponovno slanje.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnPrepareResubmit"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnCancel_Click()
    On Error GoTo EH
    
    Dim fakturaID As String
    Dim commentText As String
    Dim ok As Boolean
    
    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    commentText = InputBox("Unesite komentar za cancel:", "SEF cancel")
    If Len(Trim$(commentText)) = 0 Then Exit Sub
    
    If MsgBox("Otkazati fakturu " & fakturaID & " na SEF?", _
          vbExclamation + vbYesNo, APP_NAME) = vbNo Then Exit Sub
    
    ok = CancelInvoiceOnSEF_TX(fakturaID, commentText)
    
    Call LoadSelectedFakturaInfo
    
    If ok Then
        MsgBox "Cancel uspešno poslat.", vbInformation, APP_NAME
    Else
        MsgBox "Cancel nije uspeo.", vbExclamation, APP_NAME
    End If
    Exit Sub

EH:
    LogErr "frmSEF.btnCancel"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnStorno_Click()
    On Error GoTo EH
    
    Dim fakturaID As String
    Dim stornoComment As String
    Dim stornoNumber As String
    Dim ok As Boolean
    
    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    stornoComment = InputBox("Unesite komentar za storno:", "SEF storno")
    If Len(Trim$(stornoComment)) = 0 Then Exit Sub
    
    stornoNumber = InputBox("Unesite storno broj (opciono):", "SEF storno")
    
    If MsgBox("Stornirati fakturu " & fakturaID & " na SEF?", _
          vbExclamation + vbYesNo, APP_NAME) = vbNo Then Exit Sub
    
    ok = StornoInvoiceOnSEF_TX(fakturaID, stornoComment, stornoNumber)
    
    Call LoadSelectedFakturaInfo
    
    If ok Then
        MsgBox "Storno uspešno poslat.", vbInformation, APP_NAME
    Else
        MsgBox "Storno nije uspeo.", vbExclamation, APP_NAME
    End If
    Exit Sub

EH:
    LogErr "frmSEF.btnStorno"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnRecoverSending_Click()
    On Error GoTo EH
    
    Dim fakturaID As String
    
    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Call RecoverStuckSEFSendingInvoice(fakturaID)
    Call LoadSelectedFakturaInfo
    
    MsgBox "Recovery završen.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnRecoverSending"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnRefreshPending_Click()
    On Error GoTo EH
    
    Call RefreshPendingOutboundInvoices_TX
    Call LoadSelectedFakturaInfo
    
    MsgBox "Pending fakture osvežene.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnRefreshPending"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnRecoverAllSending_Click()
    On Error GoTo EH
    
    Call RecoverAllStuckSEFSendingInvoices
    Call LoadSelectedFakturaInfo
    
    MsgBox "SEF_SENDING recovery završen.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnRecoverAllSending"
    MsgBox Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnZatvori_Click()
    Unload Me
End Sub


Private Sub SetupHelpPage()
    Dim helpText As String
    
    helpText = "UPUTSTVO ZA SEF UPRAVLJANJE" & vbCrLf & _
               "============================" & vbCrLf & vbCrLf & _
               "1. STATUSI FAKTURE:" & vbCrLf & _
               "- READY: Faktura je spremna." & vbCrLf & _
               "- SENDING: Faktura se trenutno šalje." & vbCrLf & _
               "- SENT: Faktura uspešno primljena na SEF." & vbCrLf & _
               "- ACCEPTED: Faktura potvrdena." & vbCrLf & _
               "- REJECTED: Greška! Proveri 'Poslednja greška'." & vbCrLf & vbCrLf & _
               "2. PROCEDURA SLANJA:" & vbCrLf & _
               "Izaberi fakturu iz liste -> Klikni 'Pošalji na SEF'." & vbCrLf & _
               "Ako se pojavi status REJECTED, klikni 'Pripremi za ponovno slanje'." & vbCrLf & vbCrLf & _
               "3. TEHNICKA PODRŠKA:" & vbCrLf & _
               "Za sve probleme koji se ne rešavaju sa 'Osveži status'," & vbCrLf & _
               "kontaktiraj administratora i pošalji SEF Event Log (donja tabela)."

    Me.txtHelpBox.value = helpText
End Sub
