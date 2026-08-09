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
    MouseWheel_Attach Me

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
    StylePrimaryButton btnPosalji, Poruka("SEF_LBL_POSALJI_SEF")
    StylePrimaryButton btnOsvezi, Poruka("SEF_LBL_OSVEZI_STATUS")
    StylePrimaryButton btnPrepareResubmit, "Pripremi za ponovno slanje"
    StylePrimaryButton btnCancel, Poruka("SEF_LBL_OTKAZI_SLANJE_SEF")
    StylePrimaryButton btnStorno, "Storniraj u SEFu"
    StylePrimaryButton btnRecoverSending, "Recover sending"

    ' Batch buttons
    StylePrimaryButton btnRefreshPending, Poruka("SEF_LBL_OSVEZI_SVE_PENDING")
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
    MsgBox Poruka("SEF_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbExclamation, APP_NAME
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
    StylePrimaryButton btnOsvezi, Poruka("SEF_LBL_OSVEZI_STATUS")
    StylePrimaryButton btnPrepareResubmit, "Pripremi za ponovno slanje"
    StylePrimaryButton btnCancel, Poruka("SEF_LBL_OTKAZI_SLANJE_SEF")
    StylePrimaryButton btnStorno, "Storniraj u SEFu"
    StylePrimaryButton btnRecoverSending, "Recover sending"
    StylePrimaryButton btnRefreshPending, Poruka("SEF_LBL_OSVEZI_SVE_PENDING")
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

    ' AUD-031a: stornirane fakture must not enter the SEF send combo.
    data = ExcludeStornirano(data, TBL_FAKTURE)
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

' AUD-032d: promena izabrane fakture mora da obrise prikazan SEF kontekst.
' Bez ovoga je na ekranu ostajao status/event log PRETHODNE fakture sve dok
' operater ne klikne "Ucitaj fakturu" -- pa je tudji status izgledao kao njen.
Private Sub cmbFaktura_Change()
    On Error Resume Next
    Call ClearSEFInfo
End Sub

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
    ' AUD-032b: boja se vodi po klasi statusa, ne po rucnoj listi imena --
    ' inace zvanicni SEF statusi (APPROVED, SEEN, PAID, MISTAKE...) ispadnu
    ' neobojeni, tj. izgledaju kao "nista posebno".
    Select Case ClassifySEFExternalStatus(Me.lblSEFStatus.caption)
    Case SEF_CLS_ACCEPTED
        Me.lblSEFStatus.ForeColor = CLR_SUCCESS()
    Case SEF_CLS_REJECTED, SEF_CLS_SEND_FAILED
        ' Greska pri slanju je neuspeh, ne "obavestenje".
        Me.lblSEFStatus.ForeColor = CLR_ERROR()
    Case SEF_CLS_PENDING
        Me.lblSEFStatus.ForeColor = RGB(80, 130, 200)       ' info plava
    Case SEF_CLS_STORNO, SEF_CLS_TERMINAL, SEF_CLS_INFO
        Me.lblSEFStatus.ForeColor = CLR_WARNING()
    Case SEF_CLS_ERROR, SEF_CLS_UNKNOWN
        ' Nepoznat status = rucna provera, ne "sve u redu".
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
    Dim sefDocumentId As String

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
    sefDocumentId = Trim$(CStr(LookupValue(TBL_FAKTURE, "FakturaID", fakturaID, "SEFDocumentId")))
    
    ' AUD-032b: isti spisak koji propusta `ValidateFakturaForSEF`. Sam workflow
    ' nije dovoljan -- SEF_TECH_FAILED nastao iz statusa MISTAKE ima ZIV dokument
    ' na SEF-u, pa bi novo slanje kapija odbila kao duplikat. Ranije je forma
    ' palila "Retry slanje na SEF" koji nije mogao da prodje.
    Me.btnPosalji.enabled = CanSendSEFInvoice(workflowState, sefStatus, sefDocumentId)

    If Me.btnPosalji.enabled And workflowState = UCase$(WF_SEF_TECH_FAILED) Then
        Me.btnPosalji.caption = "Retry slanje na SEF"
    Else
        Me.btnPosalji.caption = Poruka("SEF_LBL_POSALJI_SEF")
    End If
    
    If Not Me.btnPosalji.enabled Then
        Me.btnPosalji.caption = Poruka("SEF_LBL_POSALJI_SEF")
    End If
    
    ' AUD-032b: SEF_UNKNOWN (SEF vratio nepoznat status) je stanje za rucnu
    ' proveru -- operater mora da moze da ponovi "Osvezi status". SEF_TECH_FAILED
    ' se osvezava samo kad dokument stvarno postoji na SEF-u (MISTAKE putanja);
    ' obican tehnicki pad nema SEFDocumentId, pa bi refresh pukao.
    Me.btnOsvezi.enabled = (workflowState = UCase$(WF_SEF_SENT) Or _
                            workflowState = UCase$(WF_SEF_SYNC_ERROR) Or _
                            workflowState = UCase$(WF_SEF_UNKNOWN) Or _
                            (workflowState = UCase$(WF_SEF_TECH_FAILED) And _
                             Len(sefDocumentId) > 0))
    
    Me.btnPrepareResubmit.enabled = (workflowState = UCase$(WF_SEF_REJECTED))
    
    ' AUD-032b: dugmad se pale po ISTOM spisku po kom validator propusta akciju
    ' (`CanCancelSEFStatus` / `CanStornoSEFStatus`), da forma ne bi nudila akciju
    ' koju kapija odbija -- ni obrnuto. Time je i zvanicni "MISTAKE" (greska pri
    ' slanju) dobio Cancel, a "APPROVED" Storno.
    Me.btnCancel.enabled = CanCancelSEFStatus(sefStatus)

    Me.btnStorno.enabled = CanStornoSEFStatus(sefStatus)
    
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
    Dim sendErrNo As Long
    Dim sendErrDesc As String

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

    ' AUD-032a: SendInvoiceToSEF_TX sada baca tipiziranu gresku kad SEF odbije
    ' fakturu (ERR_SEF_REJECTED) ili kad slanje tehnicki padne
    ' (ERR_SEF_SEND_FAILED). Stanje je u oba slucaja vec sacuvano, pa se ishod
    ' hvata ovde i prikazuje kao ishod, a ne kao rusenje forme.
    On Error Resume Next
    submissionID = SendInvoiceToSEF_TX(fakturaID)
    sendErrNo = Err.Number
    sendErrDesc = Err.description
    Err.Clear
    On Error GoTo EH

    Call LoadSelectedFakturaInfo

    If sendErrNo = 0 Then
        ' Poruka po stvarnom SEFWorkflowState-u, ne bezuslovno "Faktura poslata".
        MsgBox SEFSendOutcomeMessage(GetFakturaSEFWorkflowState(fakturaID), submissionID), _
               vbInformation, APP_NAME
    ElseIf sendErrNo = ERR_SEF_REJECTED Or sendErrNo = ERR_SEF_SEND_FAILED Then
        MsgBox sendErrDesc, vbExclamation, APP_NAME
    Else
        Err.Raise sendErrNo, "frmSEF.btnPosalji", sendErrDesc
    End If

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
    Dim ok As Boolean

    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        Exit Sub
    End If

    ok = RefreshSEFStatus_TX(fakturaID)
    Call LoadSelectedFakturaInfo

    ' AUD-032c: prikazuje se stvaran rezultat. Ranije je poruka uvek tvrdila da
    ' je status osvezen, i kad je SEF pao ili vratio nepoznat status.
    If ok Then
        MsgBox Poruka("SEF_MSG_SEF_STATUS_OSVEZEN"), vbInformation, APP_NAME
    Else
        MsgBox Poruka("SEF_MSG_STATUS_NIJE_OSVEZEN") & vbCrLf & _
               "SEFStatus: " & Me.lblSEFStatus.caption, vbExclamation, APP_NAME
    End If
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
        MsgBox Poruka("SEF_MSG_CANCEL_USPESNO_POSLAT"), vbInformation, APP_NAME
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
        MsgBox Poruka("SEF_MSG_STORNO_USPESNO_POSLAT"), vbInformation, APP_NAME
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
    Dim recovered As Boolean

    fakturaID = GetSelectedFakturaID()
    If Len(fakturaID) = 0 Then
        MsgBox "Izaberite fakturu.", vbExclamation, APP_NAME
        Exit Sub
    End If

    recovered = RecoverStuckSEFSendingInvoice(fakturaID)
    Call LoadSelectedFakturaInfo

    ' AUD-032c: "Recovery zavrsen" samo ako faktura vise nije u SEF_SENDING.
    If recovered Then
        MsgBox Poruka("SEF_MSG_RECOVERY_ZAVRSEN") & vbCrLf & _
               "Workflow: " & Me.lblWorkflow.caption, vbInformation, APP_NAME
    Else
        MsgBox Poruka("SEF_MSG_RECOVERY_NEUSPESAN") & vbCrLf & _
               "Workflow: " & Me.lblWorkflow.caption, vbExclamation, APP_NAME
    End If
    Exit Sub

EH:
    LogErr "frmSEF.btnRecoverSending"
    MsgBox Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnRefreshPending_Click()
    On Error GoTo EH

    Dim summaryText As String

    summaryText = RefreshPendingOutboundInvoices_TX()
    Call LoadSelectedFakturaInfo

    ' AUD-032f: batch prolaz prijavljuje sazetak, ne tihi "gotovo".
    MsgBox Poruka("SEF_MSG_PENDING_FAKTURE_OSVEZENE") & vbCrLf & summaryText, _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmSEF.btnRefreshPending"
    MsgBox Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnRecoverAllSending_Click()
    On Error GoTo EH

    Dim summaryText As String

    summaryText = RecoverAllStuckSEFSendingInvoices()
    Call LoadSelectedFakturaInfo

    ' AUD-032f: sazetak (Found/Recovered/NotRecovered/Failed) umesto tihog prolaza.
    MsgBox Poruka("SEF_MSG_SEF_SENDING_RECOVERY") & vbCrLf & summaryText, _
           vbInformation, APP_NAME
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

Private Sub UserForm_Deactivate()
    On Error Resume Next
    MouseWheel_Detach
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    MouseWheel_Detach
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    MouseWheel_Detach

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
               Poruka("SEF_MSG_SENDING_FAKTURA_TRENUTNO") & vbCrLf & _
               Poruka("SEF_MSG_SENT_FAKTURA_USPESNO") & vbCrLf & _
               "- ACCEPTED: Faktura potvrdena." & vbCrLf & _
               Poruka("SEF_MSG_REJECTED_GRESKA_PROVERI") & vbCrLf & vbCrLf & _
               "2. PROCEDURA SLANJA:" & vbCrLf & _
               Poruka("SEF_MSG_IZABERI_FAKTURU_LISTE") & vbCrLf & _
               "Ako se pojavi status REJECTED, klikni 'Pripremi za ponovno slanje'." & vbCrLf & vbCrLf & _
               Poruka("SEF_MSG_TEHNICKA_PODRSKA") & vbCrLf & _
               Poruka("SEF_MSG_SVE_PROBLEME_KOJI") & vbCrLf & _
               Poruka("SEF_MSG_KONTAKTIRAJ_ADMINISTRATORA_POSALJI")

    Me.txtHelpBox.value = helpText
End Sub
