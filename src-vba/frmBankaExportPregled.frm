Option Explicit

Private m_Blokovi As Collection
Private m_SetupDone As Boolean
Private mChromeRemoved As Boolean

Private Sub UserForm_Activate()
    On Error GoTo EH

    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
    If m_SetupDone Then
        LoadBlokovi
        Exit Sub
    End If
    m_SetupDone = True
    
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    
    ' Action buttons
    StylePrimaryButton btnOsvezi, "Osvezi"
    StylePrimaryButton btnExport, "Export u clipboard"
    StyleExitButton btnPovratak, "Povratak"
    
    ' Status / summary labels
    StyleLabel lblStatus, TXT_MUTED(), True
    StyleSubtitle lblSubtitle, "Pregled otvorenih blokova za isplatu"
    StyleFrameTitleLabel lblKopf, "Filteri"
    
    ' Section header
    StyleSectionHeader fraFilter, "Filteri"
    
    SetupList
    PopulateStanicaCombo
    LoadBlokovi
    
    Exit Sub

EH:
    LogErr "frmIsplatePregled.UserForm_Activate"
    MsgBox "Greska pri otvaranju pregleda: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub SetupList()
    With lstBlokovi
        .ColumnCount = 8
        ' Datum | Kooperant | Stanica | BrojDok | Ukupan | Isplaceno | Otvoreno | TR
        .ColumnWidths = "60;140;50;60;75;75;75;30"
    End With
End Sub

Private Sub PopulateStanicaCombo()
    cmbStanica.Clear
    cmbStanica.AddItem ""    ' empty = "Sve stanice"
    FillCmb cmbStanica, GetLookupList(TBL_STANICE, "Naziv")
End Sub

Private Sub LoadBlokovi()
    On Error GoTo EH
    
    Dim datumOd As Date, datumDo As Date
    Dim stanicaID As String
    
    On Error Resume Next
    If Len(Trim$(txtDatumOd.value)) > 0 Then datumOd = CDate(txtDatumOd.value)
    If Len(Trim$(txtDatumDo.value)) > 0 Then datumDo = CDate(txtDatumDo.value)
    On Error GoTo EH
    
    If Len(Trim$(cmbStanica.value)) > 0 Then
        stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbStanica.value, "StanicaID"))
    End If
    
    Set m_Blokovi = BuildBlokIsplataList(datumOd, datumDo, stanicaID)
    
    RenderListbox
    lblStatus.caption = SummarizeBlokList(m_Blokovi)
    Exit Sub

EH:
    LogErr "frmIsplatePregled.LoadBlokovi"
    lblStatus.caption = "Greska pri ucitavanju."
End Sub

Private Sub RenderListbox()
    lstBlokovi.Clear
    
    If m_Blokovi Is Nothing Then Exit Sub
    
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In m_Blokovi
        Set blk = v
        
        lstBlokovi.AddItem Format$(blk.Datum, "d.m.yyyy")
        Dim row As Long
        row = lstBlokovi.ListCount - 1
        
        lstBlokovi.List(row, 1) = blk.kooperantNaziv
        lstBlokovi.List(row, 2) = blk.stanicaID
        lstBlokovi.List(row, 3) = blk.BrojDokumenta
        lstBlokovi.List(row, 4) = Format$(blk.UkupanIznos, "#,##0.00")
        lstBlokovi.List(row, 5) = Format$(blk.VecIsplaceno, "#,##0.00")
        lstBlokovi.List(row, 6) = Format$(blk.OtvorenIznos, "#,##0.00")
        lstBlokovi.List(row, 7) = IIf(blk.HasTekuciRacun, "OK", "—")
    Next v
End Sub

Private Sub btnOsvezi_Click()
    LoadBlokovi
End Sub

Private Sub cmbStanica_Change()
    LoadBlokovi
End Sub

Private Sub txtDatumOd_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RemoveFocusBorder txtDatumOd
    LoadBlokovi
End Sub

Private Sub txtDatumDo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    RemoveFocusBorder txtDatumDo
    LoadBlokovi
End Sub

Private Sub txtDatumOd_Enter():    ApplyFocusBorder txtDatumOd:    End Sub
Private Sub txtDatumDo_Enter():    ApplyFocusBorder txtDatumDo:    End Sub

Private Sub btnExport_Click()
    On Error GoTo EH
    
    If m_Blokovi Is Nothing Then
        MsgBox "Nema podataka za export.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    If m_Blokovi.count = 0 Then
        MsgBox "Nema podataka za export.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    Dim tsv As String
    tsv = ExportBlokListAsTSV(m_Blokovi)
    
    Dim dataObj As Object
    Set dataObj = CreateObject("New:1C3B4210-F441-11CE-B9EA-00AA006B1A69")
    dataObj.SetText tsv
    dataObj.PutInClipboard
    
    MsgBox m_Blokovi.count & " redova kopirano u clipboard. Paste u Excel.", _
           vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmIsplatePregled.btnExport_Click"
    MsgBox "Greska pri export-u: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnPovratak_Click()
    On Error GoTo EH
    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me
    Exit Sub
EH:
    LogErr "frmIsplatePregled.btnPovratak_Click"
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If
    On Error GoTo 0
End Sub

' Mouse hover pattern (kao frmDokumenta)
Private Sub ResetActionButtons()
    StylePrimaryButton btnOsvezi, "Osvezi"
    StylePrimaryButton btnExport, "Export u clipboard"
    StyleExitButton btnPovratak, "Povratak"
End Sub

Private Sub btnOsvezi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnOsvezi
End Sub

Private Sub btnExport_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnExport
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

