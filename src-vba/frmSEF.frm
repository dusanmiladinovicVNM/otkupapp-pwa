VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSEF 
   Caption         =   "UserForm1"
   ClientHeight    =   13815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
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

Private Sub UserForm_Activate()
    On Error GoTo EH

    EnsureUserFormChromeRemoved Me, mChromeRemoved

    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me

    If m_SetupDone Then Exit Sub
    m_SetupDone = True

    ' Header zone
    On Error Resume Next
    StyleFrameTitleLabel lblKopf, "SEF upravljanje"
    StyleSubtitle lblSubtitle, "Slanje, status i recovery elektronskih faktura"
    On Error GoTo EH

    ' Section headers za Frame-ove
    StyleSectionHeader fraInfo, "Info o fakturi"
    StyleSectionHeader fraAkcije, "Akcije nad fakturom"
    StyleSectionHeader fraBatch, "Batch operacije"
    StyleSectionHeader fraEvents, "Event log"

    ' Static naslovne labele (muted, small)
    On Error Resume Next
    StyleLabel lblFakturaIDLabel, TXT_MUTED(), False
    lblFakturaIDLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblBrojFaktureLabel, TXT_MUTED(), False
    lblBrojFaktureLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblKupacLabel, TXT_MUTED(), False
    lblKupacLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblWorkflowLabel, TXT_MUTED(), False
    lblWorkflowLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblSEFStatusLabel, TXT_MUTED(), False
    lblSEFStatusLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblSEFDocumentIDLabel, TXT_MUTED(), False
    lblSEFDocumentIDLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblVersionLabel, TXT_MUTED(), False
    lblVersionLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblLastErrorLabel, TXT_MUTED(), False
    lblLastErrorLabel.Font.Size = FONT_SIZE_SMALL

    StyleLabel lblFakturaSelectorLabel, TXT_MUTED(), False
    lblFakturaSelectorLabel.Font.Size = FONT_SIZE_SMALL
    On Error GoTo EH

    ' Dinamicke value labele (light, bold za ID-jeve)
    On Error Resume Next
    StyleLabel lblFakturaID, TXT_LIGHT(), True
    StyleLabel lblBrojFakture, TXT_LIGHT(), True
    StyleLabel lblKupacNaziv, TXT_LIGHT(), False
    StyleLabel lblWorkflow, TXT_LIGHT(), True
    StyleLabel lblSEFStatus, TXT_LIGHT(), True
    StyleLabel lblSEFDocumentID, TXT_LIGHT(), False
    StyleLabel lblVersion, TXT_LIGHT(), False
    StyleLabel lblLastError, CLR_ERROR(), False
    On Error GoTo EH

    ' Action buttons (single faktura)
    StylePrimaryButton btnUcitaj, "Ucitaj fakturu"
    StylePrimaryButton btnPosalji, "Pošalji na SEF"
    StylePrimaryButton btnOsvezi, "Osveži status"
    StylePrimaryButton btnPrepareResubmit, "Pripremi za ponovno slanje"
    StylePrimaryButton btnCancel, "Otkaži slanje na SEF"
    StylePrimaryButton btnStorno, "Storniraj u SEFu"
    StylePrimaryButton btnRecoverSending, "Recover sending"

    ' Batch buttons
    StylePrimaryButton btnRefreshPending, "Osveži sve Pending"
    StylePrimaryButton btnRecoverAllSending, "Recover sve sending"

    ' Exit
    StyleExitButton btnPovratak, "Zatvori"

    ' Help textbox styling
    On Error Resume Next
    txtHelpBox.BackColor = BG_PANEL()
    txtHelpBox.ForeColor = TXT_LIGHT()
    txtHelpBox.Font.name = "Segoe UI"
    txtHelpBox.Font.Size = 9
    On Error GoTo EH

    ' Column headers iznad Events listbox-a
    SetupAllColumnHeaders

    Call SetupSEFEventList
    Call LoadFaktureIntoCombo
    Call ClearSEFInfo
    Call SetupHelpPage

    ' Force dark Pages u MultiPage1
    ForceDarkAllPages

    ' === Force Z-order ===
    On Error Resume Next

    ' Background sasvim dole
    fraBackground.ZOrder 1

    ' Sve ostale glavne kontrole na vrh
    cmbFaktura.ZOrder 0
    btnUcitaj.ZOrder 0
    fraInfo.ZOrder 0
    fraAkcije.ZOrder 0
    fraBatch.ZOrder 0
    fraEvents.ZOrder 0
    btnPovratak.ZOrder 0

    ' Action buttons unutar frame-ova
    btnPosalji.ZOrder 0
    btnOsvezi.ZOrder 0
    btnPrepareResubmit.ZOrder 0
    btnCancel.ZOrder 0
    btnStorno.ZOrder 0
    btnRecoverSending.ZOrder 0
    btnRefreshPending.ZOrder 0
    btnRecoverAllSending.ZOrder 0

    ' Header
    lblKopf.ZOrder 0
    lblSubtitle.ZOrder 0

    fraBackground.BackColor = BG_MAIN()

On Error GoTo 0

    Exit Sub

EH:
    LogErr "frmSEF.UserForm_Activate"
    MsgBox "Greška pri otvaranju SEF forme: " & Err.description, vbExclamation, APP_NAME
End Sub

' === Column headers setup ===
Private Sub SetupAllColumnHeaders()
    On Error Resume Next
    
    SetColumnHeader lbl_H_SEF1, "Vreme"
    SetColumnHeader lbl_H_SEF2, "Tip"
    SetColumnHeader lbl_H_SEF3, "Poruka"
    SetColumnHeader lbl_H_SEF4, "Detalji"
    
    On Error GoTo 0
End Sub

Private Sub SetColumnHeader(ByVal lbl As MSForms.label, ByVal txt As String)
    On Error Resume Next
    StyleListHeaderLabel lbl
    lbl.caption = txt
    On Error GoTo 0
End Sub

' === MultiPage1 dark background fix ===
Private Sub ForceDarkAllPages()
    On Error Resume Next
    
    Dim i As Long
    For i = 0 To MultiPage1.Pages.count - 1
        MultiPage1.Pages(i).BackColor = BG_MAIN()
    Next i
    
    MultiPage1.BackColor = BG_MAIN()
    
    On Error GoTo 0
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnUcitaj, "Ucitaj fakturu"
    ' btnPosalji caption se menja dinamicki, pa ne forsiramo
    StylePrimaryButton btnPosalji, btnPosalji.caption
    StylePrimaryButton btnOsvezi, "Osveži status"
    StylePrimaryButton btnPrepareResubmit, "Pripremi za ponovno slanje"
    StylePrimaryButton btnCancel, "Otkaži slanje na SEF"
    StylePrimaryButton btnStorno, "Storniraj u SEFu"
    StylePrimaryButton btnRecoverSending, "Recover sending"
    StylePrimaryButton btnRefreshPending, "Osveži sve Pending"
    StylePrimaryButton btnRecoverAllSending, "Recover sve sending"
    StyleExitButton btnPovratak, "Zatvori"
End Sub

Private Sub btnUcitaj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnUcitaj
End Sub
Private Sub btnPosalji_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnPosalji
End Sub
Private Sub btnOsvezi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnOsvezi
End Sub
Private Sub btnPrepareResubmit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnPrepareResubmit
End Sub
Private Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnCancel
End Sub
Private Sub btnStorno_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnStorno
End Sub
Private Sub btnRecoverSending_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnRecoverSending
End Sub
Private Sub btnRefreshPending_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnRefreshPending
End Sub
Private Sub btnRecoverAllSending_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnRecoverAllSending
End Sub
Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
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
    
    ' NOVO - koristi theme constants
    Select Case UCase$(Me.lblSEFStatus.caption)
    Case "SENT"
        Me.lblSEFStatus.ForeColor = RGB(80, 130, 200)       ' info plava
    Case "ACCEPTED"
        Me.lblSEFStatus.ForeColor = CLR_SUCCESS()
    Case "REJECTED"
        Me.lblSEFStatus.ForeColor = CLR_ERROR()
    Case "CANCELLED", "STORNO"
        Me.lblSEFStatus.ForeColor = CLR_WARNING()
    Case Else
        Me.lblSEFStatus.ForeColor = TXT_LIGHT()
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
    MsgBox Err.description, vbCritical, APP_NAME
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
    MsgBox Err.description, vbCritical, APP_NAME
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
    MsgBox Err.description, vbCritical, APP_NAME
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
    MsgBox Err.description, vbCritical, APP_NAME
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
    MsgBox Err.description, vbCritical, APP_NAME
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
    MsgBox Err.description, vbCritical, APP_NAME
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
    MsgBox Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnRefreshPending_Click()
    On Error GoTo EH
    
    Call RefreshPendingOutboundInvoices_TX
    Call LoadSelectedFakturaInfo
    
    MsgBox "Pending fakture osvežene.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnRefreshPending"
    MsgBox Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnRecoverAllSending_Click()
    On Error GoTo EH
    
    Call RecoverAllStuckSEFSendingInvoices
    Call LoadSelectedFakturaInfo
    
    MsgBox "SEF_SENDING recovery završen.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnRecoverAllSending"
    MsgBox Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnPovratak_Click()
    On Error GoTo EH

    ButtonActive btnPovratak
    
    frmOtkupAPP.ReturnToDashboard "SEF zatvoren."
    Unload Me
    
    Exit Sub

EH:
    LogErr "frmSEF.btnPovratak_Click"
    On Error Resume Next
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If

    On Error GoTo 0
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
