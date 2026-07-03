VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBankaExportPregled 
   Caption         =   "UserForm1"
   ClientHeight    =   12930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmBankaExportPregled.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBankaExportPregled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_Blokovi As Collection
Private m_OverrideAmounts As Object        ' Dictionary OtkupID -> custom amount
Private m_SetupDone As Boolean
Private mChromeRemoved As Boolean

Private Sub UserForm_Activate()
    On Error GoTo EH
    MouseWheel_Attach Me

    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
    If m_SetupDone Then
        LoadBlokovi
        Exit Sub
    End If
    m_SetupDone = True
    
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    
    StylePrimaryButton btnOsvezi, "Osve" & ChrW(382) & "i"
    StylePrimaryButton btnExport, "Export u clipboard"
    StylePrimaryButton btnPostaviFull, "Postavi na otvoreno"
    StylePrimaryButton btnGenerisiCSV, Poruka("BANKA_LBL_GENERISI_CSV_COMMIT")
    btnGenerisiCSV.enabled = True
    StyleExitButton btnPovratak, "Povratak"

    StyleLabel lblStatus, TXT_MUTED(), True
    StyleLabel lblSelectionSummary, TXT_MUTED(), True
    StyleSubtitle lblSubtitle, "Pregled otvorenih blokova za isplatu"
    
    StyleSectionHeader fraFilter, "Filteri"
    StyleSectionHeader fraDetail, "Detalji izabranog bloka"
    
    StyleListHeaderLabel lblColDatum
    StyleListHeaderLabel lblColKooperant
    StyleListHeaderLabel lblColStanica
    StyleListHeaderLabel lblColBrojDok
    StyleListHeaderLabel lblColUkupan
    StyleListHeaderLabel lblColIsplaceno
    StyleListHeaderLabel lblColOtvoren
    StyleListHeaderLabel lblColTR
    StyleListHeaderLabel lblColIsplatiti
    
    LayoutTopKpis
    RefreshTopKpis
    
    Set m_OverrideAmounts = CreateObject("Scripting.Dictionary")
    
    SetupList
    PopulateStanicaCombo
    LoadBlokovi
    ClearDetailPanel
    
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.UserForm_Activate"
    MsgBox "Gre" & ChrW(353) & "ka pri otvaranju pregleda: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub SetupList()
    With lstBlokovi
        .ColumnCount = 9   ' jedna vise za "Isplatiti"
        ' Datum | Kooperant | Stanica | BrojDok | Ukupan | Isplaceno | Otvoreno | TR | Isplatiti
        .ColumnWidths = "60;140;50;60;75;75;75;30;75"
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption    ' native checkboxovi
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
    
    ' Pre re-render-a, ocisti overrides koji vise nisu u listi
    PruneStaleOverrides
    
    RenderListbox
    UpdateEmptyState
    lblStatus.caption = SummarizeBlokList(m_Blokovi)
    UpdateSelectionSummary
    RefreshTopKpis
    ClearDetailPanel
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.LoadBlokovi"
    lblStatus.caption = "Gre" & ChrW(353) & "ka pri u" & ChrW(269) & "itavanju."
End Sub

Private Sub PruneStaleOverrides()
    If m_OverrideAmounts Is Nothing Then Exit Sub
    If m_Blokovi Is Nothing Then
        m_OverrideAmounts.RemoveAll
        Exit Sub
    End If
    
    Dim currentSet As Object
    Set currentSet = CreateObject("Scripting.Dictionary")
    
    Dim v As Variant
    For Each v In m_Blokovi
        Dim blk As clsBlokIsplata
        Set blk = v
        currentSet.Add blk.otkupID, True
    Next v
    
    Dim toRemove As Collection
    Set toRemove = New Collection
    
    Dim k As Variant
    For Each k In m_OverrideAmounts.keys
        If Not currentSet.Exists(k) Then
            toRemove.Add CStr(k)
        End If
    Next k
    
    Dim s As Variant
    For Each s In toRemove
        m_OverrideAmounts.Remove CStr(s)
    Next s
End Sub

Private Sub RenderListbox()
    lstBlokovi.Clear
    
    If m_Blokovi Is Nothing Then Exit Sub
    
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In m_Blokovi
        Set blk = v
        
        Dim isplatitiAmount As Double
        isplatitiAmount = GetIsplatitiAmount(blk)
        
        lstBlokovi.AddItem Format$(blk.datum, "d.m.yyyy")
        Dim row As Long
        row = lstBlokovi.ListCount - 1
        
        lstBlokovi.List(row, 1) = blk.kooperantNaziv
        lstBlokovi.List(row, 2) = blk.stanicaID
        lstBlokovi.List(row, 3) = blk.brojDokumenta
        lstBlokovi.List(row, 4) = Format$(blk.UkupanIznos, "#,##0.00")
        lstBlokovi.List(row, 5) = Format$(blk.VecIsplaceno, "#,##0.00")
        lstBlokovi.List(row, 6) = Format$(blk.OtvorenIznos, "#,##0.00")
        lstBlokovi.List(row, 7) = IIf(blk.HasTekuciRacun, "OK", "--")
        lstBlokovi.List(row, 8) = Format$(isplatitiAmount, "#,##0.00")
    Next v
End Sub

'======================================================================
' GetIsplatitiAmount
' Vraca current "Isplatiti" iznos za blok:
'   - override iz m_OverrideAmounts ako postoji
'   - inace OtvorenIznos
'======================================================================
Private Function GetIsplatitiAmount(ByVal blk As clsBlokIsplata) As Double
    If m_OverrideAmounts Is Nothing Then
        GetIsplatitiAmount = blk.OtvorenIznos
        Exit Function
    End If
    
    If m_OverrideAmounts.Exists(blk.otkupID) Then
        GetIsplatitiAmount = CDbl(m_OverrideAmounts(blk.otkupID))
    Else
        GetIsplatitiAmount = blk.OtvorenIznos
    End If
End Function

'======================================================================
' lstBlokovi_Click - pokazi detail panel za clicked row
'======================================================================
Private Sub lstBlokovi_Click()
    HandleListSelectionChange
End Sub

Private Sub lstBlokovi_Change()
    HandleListSelectionChange
End Sub

Private Sub HandleListSelectionChange()
    On Error GoTo EH
    
    ' Selection summary mora uvek da se update-uje
    UpdateSelectionSummary
    
    ' Detail panel samo ako ima fokusiran row (ListIndex >= 0)
    If lstBlokovi.ListIndex < 0 Then
        ClearDetailPanel
        Exit Sub
    End If
    
    Dim blk As clsBlokIsplata
    Set blk = GetBlokByListIndex(lstBlokovi.ListIndex)
    
    If blk Is Nothing Then
        ClearDetailPanel
        Exit Sub
    End If
    
    ' Bez TR redovi: skini check, prikazi gresku u detail panelu
    If Not blk.HasTekuciRacun Then
        lstBlokovi.Selected(lstBlokovi.ListIndex) = False
        lblDetailValidacija.caption = "Ovaj kooperant nema TekuciRacun. Ne mo" & ChrW(382) & "e biti u paketu."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
    Else
        lblDetailValidacija.Visible = False
    End If
    
    PopulateDetailPanel blk
    RefreshTopKpis
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.HandleListSelectionChange"
End Sub

'======================================================================
' Empty state
'======================================================================

Private Sub UpdateEmptyState()
    If m_Blokovi Is Nothing Then
        lblEmptyState.caption = "Nema podataka."
        lblEmptyState.Visible = True
        lstBlokovi.Visible = False
        Exit Sub
    End If
    
    If m_Blokovi.count = 0 Then
        If Len(Trim$(cmbStanica.value)) > 0 Or _
           Len(Trim$(txtDatumOd.value)) > 0 Or _
           Len(Trim$(txtDatumDo.value)) > 0 Then
            lblEmptyState.caption = "Nema rezultata za izabran filter." & vbCrLf & _
                                    "Probaj sira pravila ili klikni Osve" & ChrW(382) & "i."
        Else
            lblEmptyState.caption = "Sve otvorene stavke su zatvorene. ?" & vbCrLf & _
                                    "Nema blokova za isplatu."
        End If
        lblEmptyState.Visible = True
        lstBlokovi.Visible = False
    Else
        lblEmptyState.Visible = False
        lstBlokovi.Visible = True
    End If
End Sub

'======================================================================
' GetBlokByListIndex - mapiranje ListBox index -> clsBlokIsplata
'======================================================================
Private Function GetBlokByListIndex(ByVal idx As Long) As clsBlokIsplata
    If m_Blokovi Is Nothing Then Exit Function
    If idx < 0 Then Exit Function
    If idx >= m_Blokovi.count Then Exit Function
    
    ' Collection index = 1-based; ListBox index = 0-based
    Set GetBlokByListIndex = m_Blokovi(idx + 1)
End Function

'======================================================================
' PopulateDetailPanel - pokazi info izabranog bloka
'======================================================================
Private Sub PopulateDetailPanel(ByVal blk As clsBlokIsplata)
    lblDetailBlok.caption = blk.brojDokumenta & " -- " & blk.kooperantNaziv
    lblDetailOtvoreno.caption = "Otvoreno: " & Format$(blk.OtvorenIznos, "#,##0.00") & " RSD"
    
    Dim currentAmount As Double
    currentAmount = GetIsplatitiAmount(blk)
    txtIsplatiti.value = Format$(currentAmount, "0.00")
    
    lblDetailAvans.caption = "Kooperant avans: " & Format$(blk.KooperantAvansSaldo, "#,##0.00") & " RSD"
    If blk.KooperantAvansSaldo > 0 Then
        lblDetailAvans.ForeColor = CLR_WARNING()    ' info, ne error
    Else
        lblDetailAvans.ForeColor = TXT_MUTED()
    End If
    
    lblDetailTR.caption = "Tek. ra" & ChrW(269) & "un:" & IIf(LenB(blk.TekuciRacun) > 0, blk.TekuciRacun, "--nedostaje--")
    
    If blk.KooperantAvansSaldo > 0 Then
        lblDetailAvansHint.caption = "Primeni avans kroz Dokumenta pre isplate"
        lblDetailAvansHint.Visible = True
    Else
        lblDetailAvansHint.Visible = False
    End If
    
    EnableField txtIsplatiti
    btnPostaviFull.enabled = True
End Sub

Private Sub ClearDetailPanel()
    lblDetailBlok.caption = "(izaberi blok klikom na red)"
    lblDetailOtvoreno.caption = ""
    txtIsplatiti.value = ""
    lblDetailAvans.caption = ""
    lblDetailTR.caption = ""
    lblDetailValidacija.Visible = False
    DisableField txtIsplatiti
    btnPostaviFull.enabled = False
    lblDetailAvansHint.Visible = False
End Sub

'======================================================================
' txtIsplatiti_Exit - apply custom amount sa validacijom
'======================================================================
Private Sub txtIsplatiti_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo EH
    RemoveFocusBorder txtIsplatiti
    
    If lstBlokovi.ListIndex < 0 Then Exit Sub
    
    Dim blk As clsBlokIsplata
    Set blk = GetBlokByListIndex(lstBlokovi.ListIndex)
    If blk Is Nothing Then Exit Sub
    
    Dim newAmount As Double
    If Not TryParseDouble(txtIsplatiti.value, newAmount) Then
        lblDetailValidacija.caption = "Neispravan iznos."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
        txtIsplatiti.value = Format$(GetIsplatitiAmount(blk), "0.00")
        Exit Sub
    End If
    
    If newAmount <= 0 Then
        lblDetailValidacija.caption = "Iznos mora biti veci od 0."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
        txtIsplatiti.value = Format$(GetIsplatitiAmount(blk), "0.00")
        Exit Sub
    End If
    
    If newAmount > blk.OtvorenIznos + 0.01 Then
        lblDetailValidacija.caption = "Iznos veci od otvorenog (" & _
                                       Format$(blk.OtvorenIznos, "#,##0.00") & ")."
        lblDetailValidacija.ForeColor = CLR_ERROR()
        lblDetailValidacija.Visible = True
        txtIsplatiti.value = Format$(GetIsplatitiAmount(blk), "0.00")
        Exit Sub
    End If
    
    ' Validacija OK, zapamti override
    If Abs(newAmount - blk.OtvorenIznos) < 0.01 Then
        ' Vraceno na otvoreno = ne treba override
        If m_OverrideAmounts.Exists(blk.otkupID) Then
            m_OverrideAmounts.Remove blk.otkupID
        End If
    Else
        m_OverrideAmounts(blk.otkupID) = newAmount
    End If
    
    lblDetailValidacija.Visible = False
    
    ' Re-render samo te kolone u listbox-u
    lstBlokovi.List(lstBlokovi.ListIndex, 8) = Format$(newAmount, "#,##0.00")
    
    UpdateSelectionSummary
    RefreshTopKpis
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.txtIsplatiti_Exit"
End Sub

Private Sub txtIsplatiti_Enter()
    ApplyFocusBorder txtIsplatiti
End Sub

Private Sub btnPostaviFull_Click()
    On Error GoTo EH
    
    If lstBlokovi.ListIndex < 0 Then Exit Sub
    
    Dim blk As clsBlokIsplata
    Set blk = GetBlokByListIndex(lstBlokovi.ListIndex)
    If blk Is Nothing Then Exit Sub
    
    ' Skloni override, vrati na full
    If m_OverrideAmounts.Exists(blk.otkupID) Then
        m_OverrideAmounts.Remove blk.otkupID
    End If
    
    txtIsplatiti.value = Format$(blk.OtvorenIznos, "0.00")
    lstBlokovi.List(lstBlokovi.ListIndex, 8) = Format$(blk.OtvorenIznos, "#,##0.00")
    
    lblDetailValidacija.Visible = False
    UpdateSelectionSummary
    RefreshTopKpis
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.btnPostaviFull_Click"
End Sub

'======================================================================
' UpdateSelectionSummary - bottom toolbar live count + sum
'======================================================================
Private Sub UpdateSelectionSummary()
    If m_Blokovi Is Nothing Then
        lblSelectionSummary.caption = ""
        Exit Sub
    End If
    
    Dim selectedCount As Long
    Dim selectedSum As Double
    Dim missingTR As Long
    
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If lstBlokovi.Selected(i) Then
            Dim blk As clsBlokIsplata
            Set blk = GetBlokByListIndex(i)
            If Not blk Is Nothing Then
                If Not blk.HasTekuciRacun Then
                    missingTR = missingTR + 1
                Else
                    selectedCount = selectedCount + 1
                    selectedSum = selectedSum + GetIsplatitiAmount(blk)
                End If
            End If
        End If
    Next i
    
    Dim msg As String
    If selectedCount = 0 And missingTR = 0 Then
        msg = "Nista nije selektovano"
        lblSelectionSummary.ForeColor = TXT_MUTED()
    Else
        msg = "Selektovano: " & selectedCount & " blokova | Suma: " & _
              Format$(selectedSum, "#,##0.00") & " RSD"
        If missingTR > 0 Then
            msg = msg & "  ? Preskoceno (bez TR): " & missingTR
        End If
        lblSelectionSummary.ForeColor = CLR_SUCCESS()
    End If
    
    lblSelectionSummary.caption = msg
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
    tsv = ExportSelectionAsTSV()
    
    Dim dataObj As Object
    Set dataObj = CreateObject("New:1C3B4210-F441-11CE-B9EA-00AA006B1A69")
    dataObj.SetText tsv
    dataObj.PutInClipboard
    
    Dim selCount As Long
    selCount = CountSelected()
    
    If selCount > 0 Then
        MsgBox "Export selektovanih: " & selCount & " redova kopirano. Paste u Excel.", _
               vbInformation, APP_NAME
    Else
        MsgBox "Nista nije selektovano. Export svih: " & m_Blokovi.count & " redova.", _
               vbInformation, APP_NAME
    End If
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.btnExport_Click"
    MsgBox "Gre" & ChrW(353) & "ka pri export-u: " & Err.description, vbCritical, APP_NAME
End Sub

'======================================================================
' ExportSelectionAsTSV - selected only ako ima selection, inace sve
'======================================================================
Private Function ExportSelectionAsTSV() As String
    Dim s As String
    s = "Datum" & vbTab & "Kooperant" & vbTab & "StanicaID" & vbTab & _
        "BrojDok" & vbTab & "Ukupan" & vbTab & "Isplaceno" & vbTab & _
        "Otvoren" & vbTab & "Isplatiti" & vbTab & "TekuciRacun" & vbCrLf
    
    If m_Blokovi Is Nothing Then
        ExportSelectionAsTSV = s
        Exit Function
    End If
    
    Dim hasSelection As Boolean
    hasSelection = (CountSelected() > 0)
    
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If hasSelection And Not lstBlokovi.Selected(i) Then GoTo NextRow
        
        Dim blk As clsBlokIsplata
        Set blk = GetBlokByListIndex(i)
        If blk Is Nothing Then GoTo NextRow
        
        Dim isplatitiAmount As Double
        isplatitiAmount = GetIsplatitiAmount(blk)
        
        s = s & Format$(blk.datum, "yyyy-mm-dd") & vbTab & _
                blk.kooperantNaziv & vbTab & _
                blk.stanicaID & vbTab & _
                blk.brojDokumenta & vbTab & _
                Format$(blk.UkupanIznos, "0.00") & vbTab & _
                Format$(blk.VecIsplaceno, "0.00") & vbTab & _
                Format$(blk.OtvorenIznos, "0.00") & vbTab & _
                Format$(isplatitiAmount, "0.00") & vbTab & _
                blk.TekuciRacun & vbCrLf
NextRow:
    Next i
    
    ExportSelectionAsTSV = s
End Function

Private Function CountSelected() As Long
    Dim n As Long
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If lstBlokovi.Selected(i) Then n = n + 1
    Next i
    CountSelected = n
End Function

'======================================================================
' CollectIsplataBlokovi - blokovi za CSV naloge / specifikaciju isplata:
'   - selektovani redovi ako selekcije ima, inace svi prikazani
'   - preskace blokove bez tekuceg racuna (broji ih u outMissingTR)
'   - IsplatitiIznos = operater unos (override) ili OtvorenIznos
'======================================================================
Private Function CollectIsplataBlokovi(ByRef outMissingTR As Long) As Collection
    Dim result As New Collection
    outMissingTR = 0

    Dim hasSelection As Boolean
    hasSelection = (CountSelected() > 0)

    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If hasSelection And Not lstBlokovi.Selected(i) Then GoTo NextRow

        Dim blk As clsBlokIsplata
        Set blk = GetBlokByListIndex(i)
        If blk Is Nothing Then GoTo NextRow

        If Not blk.HasTekuciRacun Then
            outMissingTR = outMissingTR + 1
            GoTo NextRow
        End If

        blk.IsplatitiIznos = GetIsplatitiAmount(blk)
        If blk.IsplatitiIznos > 0 Then result.Add blk
NextRow:
    Next i

    Set CollectIsplataBlokovi = result
End Function

' Suma IsplatitiIznos preko kolekcije (za potvrdu i status liniju).
Private Function SumIsplatiti(ByVal blokovi As Collection) As Double
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        SumIsplatiti = SumIsplatiti + blk.IsplatitiIznos
    Next v
End Function

Private Sub btnPovratak_Click()
    On Error GoTo EH
    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.btnPovratak_Click"
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

' Mouse hover pattern
Private Sub ResetActionButtons()
    StylePrimaryButton btnOsvezi, "Osve" & ChrW(382) & "i"
    StylePrimaryButton btnExport, "Export u clipboard"
    StylePrimaryButton btnPostaviFull, "Postavi na otvoreno"
    StylePrimaryButton btnGenerisiCSV, Poruka("BANKA_LBL_GENERISI_CSV_COMMIT")
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

Private Sub btnPostaviFull_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPostaviFull
End Sub

Private Sub btnGenerisiCSV_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnGenerisiCSV
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

'======================================================================
' btnGenerisiCSV - CSV naloga za prenos (uvoz u e-banking).
' Selektovani blokovi (ili svi ako selekcije nema), iznos po bloku =
' "Isplatiti" (operater unos ili otvoreno). Potvrda pre upisa fajla.
'======================================================================
Private Sub btnGenerisiCSV_Click()
    On Error GoTo EH

    If m_Blokovi Is Nothing Then
        MsgBox "Nema podataka za naloge.", vbInformation, APP_NAME
        Exit Sub
    End If
    If m_Blokovi.count = 0 Then
        MsgBox "Nema podataka za naloge.", vbInformation, APP_NAME
        Exit Sub
    End If

    ' Platilac (firma) mora biti podesen pre generisanja
    If LenB(Trim$(DocConfigOr("SELLER_ACCOUNT", ""))) = 0 Then
        MsgBox "Nije unet teku" & ChrW(263) & "i ra" & ChrW(269) & "un firme (platilac)." & vbCrLf & _
               "Unesite ga u Pode" & ChrW(353) & "avanja -> Prodavac (firma) -> Teku" & ChrW(263) & "i ra" & ChrW(269) & "un.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim missingTR As Long
    Dim blokovi As Collection
    Set blokovi = CollectIsplataBlokovi(missingTR)

    If blokovi.count = 0 Then
        MsgBox "Nema blokova za naloge: izabrani blokovi nemaju teku" & ChrW(263) & "i ra" & ChrW(269) & "un.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim total As Double
    total = SumIsplatiti(blokovi)

    Dim msg As String
    msg = "Generisati " & blokovi.count & " naloga za prenos?" & vbCrLf & vbCrLf & _
          "Ukupan iznos: " & Format$(total, "#,##0.00") & " RSD" & vbCrLf & _
          "Datum valute: " & Format$(Date, "d.m.yyyy")
    If missingTR > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Presko" & ChrW(269) & "eno (bez TR): " & missingTR & " blokova"
    End If

    If MsgBox(msg, vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub

    Dim csvPath As String
    csvPath = GenerisiNalogeCSV(blokovi)

    If LenB(csvPath) = 0 Then
        MsgBox "Gre" & ChrW(353) & "ka pri generisanju CSV fajla. Pogledajte log.", vbCritical, APP_NAME
        Exit Sub
    End If

    lblStatus.caption = "CSV: " & blokovi.count & " naloga | " & _
                        Format$(total, "#,##0.00") & " RSD"

    MsgBox "Generisano " & blokovi.count & " naloga." & vbCrLf & vbCrLf & csvPath, _
           vbInformation, APP_NAME

    ' Otvori folder sa oznacenim fajlom (operater ga odatle uvozi u e-banking)
    On Error Resume Next
    Shell "explorer.exe /select,""" & csvPath & """", vbNormalFocus
    On Error GoTo 0
    Exit Sub

EH:
    LogErr "frmBankaExportPregled.btnGenerisiCSV_Click"
    MsgBox "Gre" & ChrW(353) & "ka pri generisanju naloga: " & Err.description, vbCritical, APP_NAME
End Sub

'======================================================================
' Top KPI Strip
'======================================================================
Private Sub LayoutTopKpis()
    On Error GoTo EH
    
    LayoutTopKpiInternals fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent
    LayoutTopKpiInternals fraKpiSelected, lblKpiSelTitle, lblKpiSelValue, lblKpiSelAccent
    LayoutTopKpiInternals fraKpiBezTR, lblKpiTRTitle, lblKpiTRValue, lblKpiTRAccent
    LayoutTopKpiInternals fraKpiAvansPool, lblKpiAvTitle, lblKpiAvValue, lblKpiAvAccent
    
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.LayoutTopKpis"
End Sub

Private Sub RefreshTopKpis()
    On Error GoTo EH
    
    ' --- 1) OTVORENO total ---
    Dim totalOpen As Double
    Dim koopCount As Long
    Dim missingTR As Long
    Dim avansPool As Double
    Dim selSum As Double
    Dim selCount As Long
    
    Dim koopSet As Object
    Set koopSet = CreateObject("Scripting.Dictionary")
    Dim koopAvansSet As Object
    Set koopAvansSet = CreateObject("Scripting.Dictionary")
    
    If Not m_Blokovi Is Nothing Then
        Dim v As Variant
        For Each v In m_Blokovi
            Dim blk As clsBlokIsplata
            Set blk = v
            totalOpen = totalOpen + blk.OtvorenIznos
            If Not koopSet.Exists(blk.kooperantID) Then koopSet.Add blk.kooperantID, True
            If Not blk.HasTekuciRacun Then missingTR = missingTR + 1
            
            ' Avans pool po kooperantu (ne duplicirati istog kooperanta)
            If Not koopAvansSet.Exists(blk.kooperantID) Then
                koopAvansSet.Add blk.kooperantID, True
                avansPool = avansPool + blk.KooperantAvansSaldo
            End If
        Next v
    End If
    
    ' Selected
    Dim i As Long
    For i = 0 To lstBlokovi.ListCount - 1
        If lstBlokovi.Selected(i) Then
            Dim sblk As clsBlokIsplata
            Set sblk = GetBlokByListIndex(i)
            If Not sblk Is Nothing Then
                If sblk.HasTekuciRacun Then
                    selCount = selCount + 1
                    selSum = selSum + GetIsplatitiAmount(sblk)
                End If
            End If
        End If
    Next i
    
    ' Card 1: OTVORENO (neutral, info)
    StyleTopKpi fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent, "neutral"
    lblKpiOtvTitle.caption = "Otvoreno ukupno"
    lblKpiOtvValue.caption = Format$(totalOpen, "#,##0") & " RSD"
    
    ' Card 2: SELECTED (ok kad >0)
    Dim selKind As String
    If selCount > 0 Then selKind = "ok" Else selKind = "neutral"
    StyleTopKpi fraKpiSelected, lblKpiSelTitle, lblKpiSelValue, lblKpiSelAccent, selKind
    lblKpiSelTitle.caption = "Selektovano"
    lblKpiSelValue.caption = selCount & " bl / " & Format$(selSum, "#,##0") & " RSD"
    
    ' Card 3: BEZ TR (warn kad >0)
    Dim trKind As String
    If missingTR > 0 Then trKind = "warn" Else trKind = "ok"
    StyleTopKpi fraKpiBezTR, lblKpiTRTitle, lblKpiTRValue, lblKpiTRAccent, trKind
    lblKpiTRTitle.caption = "Bez TR"
    If missingTR = 0 Then
        lblKpiTRValue.caption = "0 (svi imaju)"
    Else
        lblKpiTRValue.caption = missingTR & " blokova"
    End If
    
    ' Card 4: AVANS POOL (ok kad >0, info)
    Dim avKind As String
    If avansPool > 0 Then avKind = "ok" Else avKind = "neutral"
    StyleTopKpi fraKpiAvansPool, lblKpiAvTitle, lblKpiAvValue, lblKpiAvAccent, avKind
    lblKpiAvTitle.caption = "Avans pool"
    lblKpiAvValue.caption = Format$(avansPool, "#,##0") & " RSD"
    
    Exit Sub
EH:
    LogErr "frmBankaExportPregled.RefreshTopKpis"
End Sub
