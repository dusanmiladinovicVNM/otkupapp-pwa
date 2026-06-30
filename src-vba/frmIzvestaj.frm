VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIzvestaj 
   Caption         =   "UserForm1"
   ClientHeight    =   13815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmIzvestaj.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIzvestaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ============================================================
' frmIzvestaj v2.1 - Reporting
' GEAeNDERT: tblIsporuka ? tblPrijemnica
' NEU: Manjak-Tab (Page 6)
' Tabs: Saldo | Otkupljena roba | Primljena ambalaza |
'       Isplata | Zbirni po OM | Prosecna cena | Manjak
' ============================================================

Private m_SetupDone As Boolean
Private m_IsInitializing As Boolean
Private m_IsChangingToggle As Boolean
Private m_IsChangingTipToggle As Boolean
Private m_IsRefreshing As Boolean
Private m_IsLoadingEntiteti As Boolean

Private mChromeRemoved As Boolean

' Runtime tab "Pregled ambalaze" (kooperant) - MultiPage page se dodaje u runtime-u
' (.frx se ne dira). Geometrija kopirana sa lstKartica/lbl_H_KK1 (page-relativne
' koordinate su iste preko stranica MultiPage-a).
Private m_ambBuilt As Boolean
Private m_ambPageIdx As Long
Private m_lstAmb As MSForms.ListBox
Private m_lblAmbTitle As MSForms.label

' Runtime tab "Otkupni listovi" (Otkupna mesta)
Private m_otkBuilt As Boolean
Private m_otkPageIdx As Long
Private WithEvents m_lstOtk As MSForms.ListBox
Attribute m_lstOtk.VB_VarHelpID = -1
Private m_lblOtkTitle As MSForms.label
Private WithEvents m_btnStampajOtk As MSForms.CommandButton
Attribute m_btnStampajOtk.VB_VarHelpID = -1
' Dugmad za stampu uz deljeni "Detalji" panel: Otkupljena roba / Ambalaza.
Private WithEvents m_btnStampajOtkRoba As MSForms.CommandButton
Attribute m_btnStampajOtkRoba.VB_VarHelpID = -1
Private WithEvents m_btnStampajAmb As MSForms.CommandButton
Attribute m_btnStampajAmb.VB_VarHelpID = -1
' OtpremnicaID po redu lstOtkupRoba (ListBox podrzava max 10 kolona -> bez
' skrivene ref-kljuc kolone; kljuc = CStr(ListIndex)).
Private m_otkOtpID As Object

Private Sub UserForm_Activate()
    On Error GoTo EH
    
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    
    SetupAllColumnHeaders
    
    If m_SetupDone Then GoTo Continue_Setup
    m_SetupDone = True

    ' Nova instanca forme -> resetuj modul-stanje panela "Detalji otkupa"
    ' (dinamicke kontrole prethodne instance vise ne postoje posle Unload-a).
    KarticaDetalji_Reset
    m_ambPageIdx = -1   ' dok se runtime tab "Pregled ambala" & ChrW(382) & "e" ne kreira
    m_otkPageIdx = -1   ' dok se runtime tab "Otkupni listovi" ne kreira

    m_IsInitializing = True
    
    ' Header zone (dodati u Designer ako nemas)
    On Error Resume Next
    StyleFrameTitleLabel lblKopf, Poruka("RPT_LBL_IZVESTAJI")
    StyleSubtitle lblSubtitle, Poruka("RPT_LBL_PREGLED_SALDO_OTKUPLJENE")
    On Error GoTo EH
    
    ' Style action buttons
    StylePrimaryButton btnUnos, Poruka("RPT_LBL_PRIKAZI")
    StylePrimaryButton btnStampaj, Poruka("FAK_LBL_STAMPAJ")
    StyleExitButton btnPovratak, "Povratak"
    
    ' Style entity toggle buttons (default: inactive look)
    StyleToggleInactive tglOM, "Otkupna mesta"
    StyleToggleInactive tglKupci, "Kupci"
    StyleToggleInactive tglVozaci, "Vozaci"
    StyleToggleInactive tglKooperanti, "Kooperanti"
    
    ' Tip toggle (pojedinacni/zbirni)
    StyleToggleInactive tglPojedinacni, "Pojedinacni"
    StyleToggleInactive tglZbirni, "Zbirni"
    
    ' Status label
    On Error Resume Next
    StyleLabel lblStatus, TXT_MUTED(), False
    On Error GoTo EH
    
    ' Kartica print button (specijalan za kooperante)
    On Error Resume Next
    StylePrimaryButton btnStampajKarticu, Poruka("RPT_LBL_STAMPAJ_KARTICU")
    On Error GoTo EH
    
    StyleListHeaderLabel lblSidebarEntitet
    lblSidebarEntitet.caption = "ENTITET"

    StyleListHeaderLabel lblSidebarTip
    lblSidebarTip.caption = Poruka("RPT_LBL_TIP_IZVESTAJA")

    StyleListHeaderLabel lblSidebarFilter
    lblSidebarFilter.caption = "FILTER"

    ' Akcent linije (gold tanka linija ispod section header-a)
    lblAccentEntitet.BackColor = BTN_ACTIVE()
    lblAccentEntitet.BackStyle = fmBackStyleOpaque
    lblAccentEntitet.BorderStyle = fmBorderStyleNone

    lblAccentTip.BackColor = BTN_ACTIVE()
    lblAccentTip.BackStyle = fmBackStyleOpaque
    lblAccentTip.BorderStyle = fmBorderStyleNone

    lblAccentFilter.BackColor = BTN_ACTIVE()
    lblAccentFilter.BackStyle = fmBackStyleOpaque
    lblAccentFilter.BorderStyle = fmBorderStyleNone
    
    StyleLabel lblFilterStanica, TXT_MUTED(), False
    lblFilterStanica.Font.Size = FONT_SIZE_SMALL    ' 9pt

    StyleLabel lblFilterOd, TXT_MUTED(), False
    lblFilterOd.Font.Size = FONT_SIZE_SMALL
    
    ' defaults
    SetTipToggle "tglPojedinacni"
    SetEntitetToggle "tglOM"
    
    LoadEntiteti
    
    cmbVrstaRobe.Clear
    cmbVrstaRobe.AddItem ""
    Dim kulture As Variant, i As Long
    kulture = GetLookupList(TBL_KULTURE, "VrstaVoca")
    If IsArray(kulture) Then
        For i = LBound(kulture) To UBound(kulture)
            cmbVrstaRobe.AddItem CStr(kulture(i))
        Next i
    End If
    
    txtDatumOd.value = "1.1." & Year(Date)
    txtDatumDo.value = Format$(Date, "d.m.yyyy")
    
    SetupListBoxes
    On Error Resume Next
    mpReports.Pages(3).caption = "Ambala" & ChrW(382) & "a"   ' "Primljena ambalaza" -> "Ambala" & ChrW(382) & "a"
    On Error GoTo EH
    EnsureKarticaAmbPage          ' runtime tab "Pregled ambala" & ChrW(382) & "e" (pre UpdateReportMode)
    EnsureOtkupListePage          ' runtime tab "Otkupni listovi" (pre UpdateReportMode)
    EnsureDetaljiButtons          ' dugmad za stampu uz "Detalji" (Otk.roba / Ambalaza)
    UpdateReportMode
    ForceDarkAllPages
    
    m_IsInitializing = False
    AutoRefresh

Continue_Setup:
    Exit Sub

EH:
    LogErr "frmIzvestaj.UserForm_Activate"
End Sub
Private Sub AutoRefresh()

    If m_IsInitializing Then Exit Sub
    If m_IsRefreshing Then Exit Sub

    If tglPojedinacni.value Then
        If cmbEntitet.ListIndex < 0 Or cmbEntitet.value = "" Then Exit Sub
    End If

    m_IsRefreshing = True
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    btnUnos_Click

CleanExit:
    Application.ScreenUpdating = True
    m_IsRefreshing = False
    Exit Sub

CleanFail:
    ' optional: Debug.Print Err.Number, Err.Description
    Resume CleanExit
End Sub

' === Toggle styling -- inline u frmIzvestaj ===

Private Sub StyleToggleActive(ByVal tgl As MSForms.ToggleButton, ByVal lbl As String)
    On Error Resume Next
    With tgl
        .caption = lbl
        .BackColor = BTN_ACTIVE()             ' gold
        .ForeColor = RGB(20, 20, 20)            ' dark text on gold
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .BorderStyle = fmBorderStyleNone
    End With
    On Error GoTo 0
End Sub

Private Sub StyleToggleInactive(ByVal tgl As MSForms.ToggleButton, ByVal lbl As String)
    On Error Resume Next
    With tgl
        .caption = lbl
        .BackColor = BG_PANEL()                ' or fallback BG_MAIN()
        .ForeColor = TXT_MUTED()
        .Font.name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = False
        .BorderStyle = fmBorderStyleSingle
    End With
    On Error GoTo 0
End Sub
' ============================================================
' KASKADIERUNG
' ============================================================

Private Function GetActiveEntitetTip() As String
    If tglOM.value Then
        GetActiveEntitetTip = "Otkupna mesta"
    ElseIf tglKupci.value Then
        GetActiveEntitetTip = "Kupci"
    ElseIf tglVozaci.value Then
        GetActiveEntitetTip = "Vozaci"
    ElseIf tglKooperanti.value Then
        GetActiveEntitetTip = "Kooperanti"
    Else
        GetActiveEntitetTip = "Otkupna mesta" ' fallback
    End If
End Function

Private Sub SetEntitetToggle(ByVal activeName As String)

    If m_IsChangingToggle Then Exit Sub

    m_IsChangingToggle = True

    tglOM.value = (activeName = "tglOM")
    tglKupci.value = (activeName = "tglKupci")
    tglVozaci.value = (activeName = "tglVozaci")
    tglKooperanti.value = (activeName = "tglKooperanti")

    m_IsChangingToggle = False

    UpdateEntitetToggleUI

End Sub

Private Sub UpdateEntitetToggleUI()
    If tglOM.value Then StyleToggleActive tglOM, "Otkupna mesta" Else StyleToggleInactive tglOM, "Otkupna mesta"
    If tglKupci.value Then StyleToggleActive tglKupci, "Kupci" Else StyleToggleInactive tglKupci, "Kupci"
    If tglVozaci.value Then StyleToggleActive tglVozaci, "Vozaci" Else StyleToggleInactive tglVozaci, "Vozaci"
    If tglKooperanti.value Then StyleToggleActive tglKooperanti, "Kooperanti" Else StyleToggleInactive tglKooperanti, "Kooperanti"
End Sub

Private Sub tglOM_Click()
    If m_IsChangingToggle Or m_IsInitializing Then Exit Sub
    SetEntitetToggle "tglOM"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub tglKupci_Click()
    If m_IsChangingToggle Or m_IsInitializing Then Exit Sub
    SetEntitetToggle "tglKupci"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub tglVozaci_Click()
    If m_IsChangingToggle Or m_IsInitializing Then Exit Sub
    SetEntitetToggle "tglVozaci"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub
Private Sub tglKooperanti_Click()
    If m_IsChangingToggle Then Exit Sub
    SetEntitetToggle "tglKooperanti"
    LoadEntiteti
    UpdateReportMode
    AutoRefresh
End Sub

Private Function IsZbirniMode() As Boolean
    IsZbirniMode = tglZbirni.value
End Function

Private Sub SetTipToggle(ByVal activeName As String)

    If m_IsChangingTipToggle Then Exit Sub
    m_IsChangingTipToggle = True

    tglPojedinacni.value = (activeName = "tglPojedinacni")
    tglZbirni.value = (activeName = "tglZbirni")

    m_IsChangingTipToggle = False

    UpdateTipToggleUI
End Sub

Private Sub UpdateTipToggleUI()
    If tglPojedinacni.value Then StyleToggleActive tglPojedinacni, "Pojedinacni" Else StyleToggleInactive tglPojedinacni, "Pojedinacni"
    If tglZbirni.value Then StyleToggleActive tglZbirni, "Zbirni" Else StyleToggleInactive tglZbirni, "Zbirni"
End Sub

Private Sub tglPojedinacni_Click()
    If m_IsChangingTipToggle Or m_IsInitializing Then Exit Sub
    SetTipToggle "tglPojedinacni"
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub tglZbirni_Click()
    If m_IsChangingTipToggle Or m_IsInitializing Then Exit Sub
    SetTipToggle "tglZbirni"
    UpdateReportMode
    AutoRefresh
End Sub

Private Sub LoadEntiteti()
    ' Guard: programsko punjenje + ListIndex=0 ne sme da okine cmbEntitet_Change
    ' (inace AutoRefresh generise sve izvestaje jos jednom). On Error GoTo done
    ' garantuje da se guard uvek vrati na False.
    m_IsLoadingEntiteti = True
    On Error GoTo done
    cmbEntitet.Clear

    Select Case GetActiveEntitetTip()
        Case "Otkupna mesta"
            FillCmb cmbEntitet, GetLookupList(TBL_STANICE, "Naziv")
        Case "Kupci"
            FillCmb cmbEntitet, GetLookupList(TBL_KUPCI, "Naziv")
        Case "Vozaci"
            FillCmb cmbEntitet, GetVozacDisplayList()
        Case "Kooperanti"
            Dim koopData As Variant
            koopData = GetTableData(TBL_KOOPERANTI)
            If IsArray(koopData) Then
                Dim colID As Long, colIme As Long, colPrezime As Long, colAktivan As Long
                colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
                colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
                colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
                colAktivan = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
                Dim k As Long
                For k = 1 To UBound(koopData, 1)
                    If CStr(koopData(k, colAktivan)) = STATUS_AKTIVAN Then
                        cmbEntitet.AddItem CStr(koopData(k, colIme)) & " " & _
                                          CStr(koopData(k, colPrezime)) & " (" & _
                                          CStr(koopData(k, colID)) & ")"
                    End If
                Next k
            End If
    End Select

    If cmbEntitet.ListCount > 0 Then cmbEntitet.ListIndex = 0

done:
    m_IsLoadingEntiteti = False
End Sub
Private Sub cmbEntitet_Change()
    ' LoadEntiteti postavlja ListIndex=0 sto bi inace okinulo jos jedan AutoRefresh
    ' (duplo generisanje svih izvestaja). Guard preskace taj programski Change;
    ' rucni izbor kooperanta i dalje refreshuje.
    If m_IsLoadingEntiteti Then Exit Sub
    AutoRefresh
End Sub


Private Sub ResetActionButtons()
    StylePrimaryButton btnUnos, Poruka("RPT_LBL_PRIKAZI")
    StylePrimaryButton btnStampaj, Poruka("FAK_LBL_STAMPAJ")
    StyleExitButton btnPovratak, "Povratak"
    
    On Error Resume Next
    StylePrimaryButton btnStampajKarticu, Poruka("RPT_LBL_STAMPAJ_KARTICU")
    On Error GoTo 0
End Sub

Private Sub btnUnos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnos
End Sub

Private Sub btnStampaj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnStampaj
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub btnStampajKarticu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    On Error Resume Next
    ButtonHover btnStampajKarticu
    On Error GoTo 0
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

' ============================================================
' ZBIRNI/POJEDINACNI
' ============================================================

Private Sub UpdateReportMode()
    Dim isPojed As Boolean
    isPojed = tglPojedinacni.value
    cmbEntitet.enabled = isPojed
    
    ' Alle erstmal ausblenden
    Dim pg As Long
    For pg = 0 To mpReports.Pages.count - 1
        mpReports.Pages(pg).Visible = False
    Next pg
    
    If isPojed Then
        Select Case GetActiveEntitetTip()
            Case "Otkupna mesta"
                mpReports.Pages(0).Visible = True  ' Saldo OM
                mpReports.Pages(2).Visible = True  ' Otkupljena roba
                mpReports.Pages(3).Visible = True  ' Ambalaza
                mpReports.Pages(4).Visible = True  ' Isplata
                mpReports.Pages(6).Visible = True  ' Prosecna cena
                If m_otkPageIdx >= 0 Then mpReports.Pages(m_otkPageIdx).Visible = True  ' Otkupni listovi
            Case "Kupci"
                mpReports.Pages(1).Visible = True  ' Saldo Kupci
                mpReports.Pages(2).Visible = True  ' Otkupljena roba
                mpReports.Pages(3).Visible = True  ' Ambalaza
                mpReports.Pages(6).Visible = True  ' Prosecna cena
                mpReports.Pages(7).Visible = True  ' Manjak
            Case "Vozaci"
                mpReports.Pages(3).Visible = True  ' Ambalaza
                mpReports.Pages(7).Visible = True  ' Manjak
            Case "Kooperanti"
                mpReports.Pages(8).Visible = True  ' Kartica
                If m_ambPageIdx >= 0 Then mpReports.Pages(m_ambPageIdx).Visible = True  ' Pregled ambalaze
        End Select
    Else
        mpReports.Pages(5).Visible = True  ' Zbirni
        mpReports.Pages(6).Visible = True  ' Prosecna cena
        mpReports.Pages(7).Visible = True  ' Manjak
    End If
    
    ' posle Visible=True podesavanja:
    Dim firstVisible As Long
    firstVisible = -1
    For pg = 0 To mpReports.Pages.count - 1
        If mpReports.Pages(pg).Visible Then
            firstVisible = pg
            Exit For
        End If
    Next pg
    If firstVisible >= 0 Then mpReports.value = firstVisible

End Sub

' ============================================================
' LISTBOX SETUP
' ============================================================

Private Sub SetupListBoxes()
    
    With lstSaldoOM
        .ColumnCount = 7
        .ColumnWidths = "120;65;80;80;80;80;50"
    End With
    
    ' Saldo Kupci: Vrsta | Kolicina | Cena | Vrednost |  Novac | Saldo | Ambalaza
    With lstSaldoKupci
        .ColumnCount = 7
        .ColumnWidths = "80;80;70;80;80;80;60"
    End With
    
    With lstOtkupRoba
        .ColumnCount = 10          ' MSForms ListBox limit; Manjak kg+% spojeni, refkey u dict
        .ColumnWidths = "45;70;45;25;75;50;50;50;55;80"   ' fallback; LayoutOtkupRobaHeaders prepisuje
    End With
    
    With lstAmbalaza
        .ColumnCount = 7          ' 6 vidljivih + skrivena ref-kljuc kolona (idx 6)
        .ColumnWidths = "55;90;45;80;50;50;0"
    End With
    
    With lstIsplata
        .ColumnCount = 5
        .ColumnWidths = "120;80;80;80;80"
    End With
    
    With lstZbirni
        .ColumnCount = 5
        .ColumnWidths = "100;80;80;80;80"
    End With
    
    With lstProsecnaCena
        .ColumnCount = 4
        .ColumnWidths = "100;80;80;80"
    End With
    
    With lstManjak
        .ColumnCount = 6
        .ColumnWidths = "70;70;70;60;55;60"
    End With
    
    With lstKartica
        .ColumnCount = 8
        .ColumnWidths = "60;70;140;80;80;80;60;0"   ' fallback; LayoutKarticaHeaders prepisuje po stvarnoj sirini
    End With

    ' Sirine + header pozicije kartice se racunaju iz stvarne lstKartica.width
    ' (auto-fit, da "Saldo amb." kolona i njen header stanu i budu poravnati).
    LayoutKarticaHeaders
    LayoutOtkupRobaHeaders
    LayoutSaldoOMHeaders
End Sub

Private Sub UpdateStatusLabel()
    On Error Resume Next
    
    Dim msg As String
    Dim activeTab As Long
    activeTab = mpReports.value
    
    Dim activeList As MSForms.ListBox
    
    Select Case activeTab
        Case 0:    Set activeList = lstSaldoOM
        Case 1:    Set activeList = lstSaldoKupci
        Case 2:    Set activeList = lstOtkupRoba
        Case 3:    Set activeList = lstAmbalaza
        Case 4:    Set activeList = lstIsplata
        Case 5:    Set activeList = lstZbirni
        Case 6:    Set activeList = lstProsecnaCena
        Case 7:    Set activeList = lstManjak
        Case 8:    Set activeList = lstKartica
    End Select

    If activeList Is Nothing And m_ambPageIdx >= 0 And activeTab = m_ambPageIdx Then
        Set activeList = m_lstAmb
    End If

    If activeList Is Nothing And m_otkPageIdx >= 0 And activeTab = m_otkPageIdx Then
        Set activeList = m_lstOtk
    End If

    If activeList Is Nothing Then
        lblStatus.caption = "Spreman za izve" & ChrW(353) & "taj"
        lblStatus.ForeColor = TXT_MUTED()
    ElseIf activeList.ListCount = 0 Then
        lblStatus.caption = "Nema podataka za izabran filter"
        lblStatus.ForeColor = TXT_MUTED()
    Else
        lblStatus.caption = activeList.ListCount & " redova ucitano | Period: " & _
                            txtDatumOd.value & " - " & txtDatumDo.value
        lblStatus.ForeColor = CLR_SUCCESS()
    End If
    
    On Error GoTo 0
End Sub

' ============================================================
' HAUPTAKTION
' ============================================================

Private Sub btnUnos_Click()
    On Error GoTo EH

    Dim datumOd As Date, datumDo As Date
    datumOd = CDate(txtDatumOd.value)
    datumDo = CDate(txtDatumDo.value)

    Dim zbirni As Boolean
    zbirni = IsZbirniMode()

    Dim entitetID As String
    Dim entitetTip As String

    ' Entitet tip je uvek potreban (OM/Kupac/Vozac),
    ' entitetID je potreban samo za POJEDINACNI
    Select Case GetActiveEntitetTip()
        Case "Otkupna mesta": entitetTip = "OM"
        Case "Kupci":         entitetTip = "Kupac"
        Case "Vozaci":        entitetTip = "Vozac"
        Case "Kooperanti":    entitetTip = "Kooperant"
        Case Else:            entitetTip = "OM"
    End Select

    If Not zbirni Then
        If cmbEntitet.value = "" Then
            MsgBox "Izaberite entitet!", vbExclamation, APP_NAME
            Exit Sub
        End If

        Select Case entitetTip
            Case "OM"
                entitetID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbEntitet.value, "StanicaID"))
            Case "Kupac"
                entitetID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbEntitet.value, "KupacID"))
            Case "Vozac"
                entitetID = ExtractIDFromDisplay(cmbEntitet.value)
           Case "Kooperant"
                entitetID = ExtractIDFromDisplay(cmbEntitet.value)
        End Select
    Else
        entitetID = ""
    End If

    Application.ScreenUpdating = False
    BeginTableCache   ' citaj svaku tabelu jednom za sve Generate* (read-only blok)

    If zbirni Then
        GenerateZbirniReport datumOd, datumDo, entitetTip
        GenerateProsecnaCenaReport entitetTip, "", datumOd, datumDo
        GenerateManjakReport entitetTip, "", datumOd, datumDo
    Else
        GenerateSaldoReport entitetTip, entitetID, datumOd, datumDo
        GenerateOtkupRobaReport entitetTip, entitetID, datumOd, datumDo
        GenerateAmbalazeReport entitetTip, entitetID, datumOd, datumDo
        GenerateIsplataReport entitetTip, entitetID, datumOd, datumDo
        GenerateProsecnaCenaReport entitetTip, entitetID, datumOd, datumDo
        GenerateManjakReport entitetTip, entitetID, datumOd, datumDo
        If entitetTip = "OM" Then GenerateOtkupListeReport entitetID, datumOd, datumDo
        If entitetTip = "Kooperant" Then
            GenerateKarticaReport entitetID, datumOd, datumDo
            GenerateKarticaAmbReport entitetID, datumOd, datumDo
        End If
    End If

    UpdateStatusLabel
    
    EndTableCache
    Application.ScreenUpdating = True
    Exit Sub

EH:
    LogErr "frmIzvestaj.btnUnos"
    EndTableCache
    Application.ScreenUpdating = True
    MsgBox "Gre" & ChrW(353) & "ka pri u" & ChrW(269) & "itavanju izve" & ChrW(353) & "taja: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub mpReports_Change()
    UpdateStatusLabel
    ' Deljeni "Detalji" panel: Kartica (8), "Otkupni listovi", Otkupljena roba (2),
    ' Ambalaza (3).
    Dim onOtk As Boolean
    onOtk = (m_otkPageIdx >= 0 And mpReports.value = m_otkPageIdx)
    Dim onRoba As Boolean: onRoba = (mpReports.value = 2)
    Dim onAmb As Boolean: onAmb = (mpReports.value = 3)
    KarticaDetalji_SetVisible (mpReports.value = 8 Or onOtk Or onRoba Or onAmb)
    ' Dugmad za stampu vidljiva samo na svom tabu.
    If Not m_btnStampajOtk Is Nothing Then m_btnStampajOtk.Visible = onOtk
    If Not m_btnStampajOtkRoba Is Nothing Then m_btnStampajOtkRoba.Visible = onRoba
    If Not m_btnStampajAmb Is Nothing Then m_btnStampajAmb.Visible = onAmb
End Sub
Private Sub UpdateUnosButtonState()

    ' If Zbirni ? no entitet required
    If tglZbirni.value Then
        btnUnos.enabled = True
        Exit Sub
    End If

    ' If Pojedinacni ? entitet must be selected
    btnUnos.enabled = (cmbEntitet.ListIndex >= 0 And cmbEntitet.value <> "")

End Sub

' ============================================================
' SALDO
' ============================================================

Private Sub GenerateSaldoReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                ByVal datumOd As Date, ByVal datumDo As Date)
    If entitetTip = "OM" Then
        GenerateSaldoOM entitetID, datumOd, datumDo
    ElseIf entitetTip = "Kupac" Then
        GenerateSaldoKupci entitetID, datumOd, datumDo
    End If
End Sub

Private Sub GenerateSaldoOM(ByVal stanicaID As String, _
                            ByVal datumOd As Date, ByVal datumDo As Date)
    lstSaldoOM.Clear
    
    Dim data As Variant
    data = ReportSaldoOM(stanicaID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        lstSaldoOM.AddItem CStr(data(i, 1))
        For j = 2 To UBound(data, 2)
            If IsNumeric(data(i, j)) Then
                If j = 7 Then  ' Ambalaza
                    lstSaldoOM.List(lstSaldoOM.ListCount - 1, j - 1) = Format$(data(i, j), "#,##0")
                Else
                    lstSaldoOM.List(lstSaldoOM.ListCount - 1, j - 1) = Format$(data(i, j), "#,##0.00")
                End If
            Else
                lstSaldoOM.List(lstSaldoOM.ListCount - 1, j - 1) = CStr(data(i, j))
            End If
        Next j
    Next i
End Sub

Private Sub GenerateSaldoKupci(ByVal kupacID As String, _
                               ByVal datumOd As Date, ByVal datumDo As Date)
    lstSaldoKupci.Clear
    
    Dim data As Variant
    data = ReportSaldoKupci(kupacID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim fmts As Variant
    fmts = Array("", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0")
    
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        lstSaldoKupci.AddItem CStr(data(i, 1))
        For j = 2 To UBound(data, 2)
            If IsNumeric(data(i, j)) And data(i, j) <> "" Then
                lstSaldoKupci.List(lstSaldoKupci.ListCount - 1, j - 1) = _
                    Format$(CDbl(data(i, j)), CStr(fmts(j - 1)))
            ElseIf CStr(data(i, j)) <> "" Then
                lstSaldoKupci.List(lstSaldoKupci.ListCount - 1, j - 1) = CStr(data(i, j))
            End If
        Next j
    Next i
End Sub

' ============================================================
' OTKUPLJENA ROBA
' ============================================================

Private Sub GenerateOtkupRobaReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                    ByVal datumOd As Date, ByVal datumDo As Date)
    lstOtkupRoba.Clear
    KarticaDetalji_Clear
    Set m_otkOtpID = CreateObject("Scripting.Dictionary")

    Dim data As Variant
    data = ReportOtkupRoba(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub

    Dim isOM As Boolean: isOM = (entitetTip = "OM")
    Dim i As Long, j As Long, li As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, 1)) Then
            lstOtkupRoba.AddItem Format$(CDate(data(i, 1)), "d.m.yy")
        Else
            lstOtkupRoba.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        End If
        li = lstOtkupRoba.ListCount - 1

        If isOM Then
            ' OM: 12 logickih kolona -> 10 listbox kolona (Manjak kg+% spojeni;
            ' OtpremnicaID u m_otkOtpID jer ListBox podrzava max 10 kolona).
            lstOtkupRoba.List(li, 1) = CStr(IIf(IsEmpty(data(i, 2)), "", data(i, 2)))
            lstOtkupRoba.List(li, 2) = CStr(IIf(IsEmpty(data(i, 3)), "", data(i, 3)))
            lstOtkupRoba.List(li, 3) = CStr(IIf(IsEmpty(data(i, 4)), "", data(i, 4)))
            lstOtkupRoba.List(li, 4) = CStr(IIf(IsEmpty(data(i, 5)), "", data(i, 5)))
            If IsNumeric(data(i, 6)) Then lstOtkupRoba.List(li, 5) = FmtKolicina(CDbl(data(i, 6)))
            If IsNumeric(data(i, 7)) Then lstOtkupRoba.List(li, 6) = FmtKolicina(CDbl(data(i, 7)))
            If IsNumeric(data(i, 8)) Then lstOtkupRoba.List(li, 7) = FmtKolicina(CDbl(data(i, 8)))
            If IsNumeric(data(i, 9)) Then lstOtkupRoba.List(li, 8) = FmtKolicina(CDbl(data(i, 9)))
            Dim mStr As String: mStr = ""
            If IsNumeric(data(i, 10)) And Not IsEmpty(data(i, 10)) Then mStr = FmtKolicina(CDbl(data(i, 10)))
            If IsNumeric(data(i, 11)) And Not IsEmpty(data(i, 11)) Then _
                mStr = Trim$(mStr & " / " & Format$(CDbl(data(i, 11)), "0.0") & "%")
            lstOtkupRoba.List(li, 9) = mStr
            Dim rk As String: rk = CStr(IIf(IsEmpty(data(i, 12)), "", data(i, 12)))
            If Left$(rk, 4) = "OTP|" Then m_otkOtpID(CStr(li)) = Mid$(rk, 5)
        Else
            For j = 2 To UBound(data, 2)
                If IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
                    If GetActiveEntitetTip() = "Kupci" And j = UBound(data, 2) Then
                        lstOtkupRoba.List(li, j - 1) = Format$(CDbl(data(i, j)), "#,##0.00")
                    Else
                        lstOtkupRoba.List(li, j - 1) = FmtKolicina(CDbl(data(i, j)))
                    End If
                Else
                    lstOtkupRoba.List(li, j - 1) = CStr(IIf(IsEmpty(data(i, j)), "", data(i, j)))
                End If
            Next j
        End If
    Next i
End Sub

' ============================================================
' AMBALAZA (unveraendert)
' ============================================================

Private Sub GenerateAmbalazeReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                   ByVal datumOd As Date, ByVal datumDo As Date)
    lstAmbalaza.Clear
    KarticaDetalji_Clear
    
    Dim isZb As Boolean
    isZb = IsZbirniMode()
    
    Dim data As Variant
    data = ReportAmbalaza(entitetTip, entitetID, datumOd, datumDo, isZb)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If IsDate(data(i, 1)) Then
            lstAmbalaza.AddItem Format$(CDate(data(i, 1)), "d.m.yy")
        Else
            lstAmbalaza.AddItem CStr(data(i, 1))
        End If
        lstAmbalaza.List(lstAmbalaza.ListCount - 1, 1) = CStr(data(i, 2))
        lstAmbalaza.List(lstAmbalaza.ListCount - 1, 2) = CStr(data(i, 3))
        lstAmbalaza.List(lstAmbalaza.ListCount - 1, 3) = CStr(data(i, 4))
        If IsNumeric(data(i, 5)) And data(i, 5) <> "" Then
            lstAmbalaza.List(lstAmbalaza.ListCount - 1, 4) = Format$(CLng(data(i, 5)), "#,##0")
        End If
        If IsNumeric(data(i, 6)) And data(i, 6) <> "" Then
            lstAmbalaza.List(lstAmbalaza.ListCount - 1, 5) = Format$(CLng(data(i, 6)), "#,##0")
        End If
        ' Skrivena kol. 6 = ref-kljuc (AMB|<DokTip>|<DokID>); samo pojedinacni (7 kol).
        If UBound(data, 2) >= 7 Then
            lstAmbalaza.List(lstAmbalaza.ListCount - 1, 6) = CStr(IIf(IsEmpty(data(i, 7)), "", data(i, 7)))
        End If
    Next i
End Sub

' ============================================================
' ISPLATA (unveraendert)
' ============================================================

Private Sub GenerateIsplataReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                  ByVal datumOd As Date, ByVal datumDo As Date)
    lstIsplata.Clear
    
    Dim data As Variant
    data = ReportIsplata(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long, j As Long
    For i = 1 To UBound(data, 1)
        lstIsplata.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        For j = 2 To UBound(data, 2)
            If IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
                lstIsplata.List(lstIsplata.ListCount - 1, j - 1) = Format$(CDbl(data(i, j)), "#,##0.00")
            Else
                lstIsplata.List(lstIsplata.ListCount - 1, j - 1) = ""
            End If
        Next j
    Next i
End Sub

' ============================================================
' ZBIRNI PO OM (unveraendert)
' ============================================================

Private Sub GenerateZbirniReport(ByVal datumOd As Date, ByVal datumDo As Date, _
                                 ByVal entitetTip As String)
    lstZbirni.Clear
    
    Dim data As Variant
    data = ReportZbirni(entitetTip, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim isVozac As Boolean
    isVozac = (entitetTip = "Vozac")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstZbirni.AddItem CStr(data(i, 1))
        
        If isVozac Then
            ' Amb Izlaz, Amb Vracena, Manjak kg, Manjak %
            If IsNumeric(data(i, 2)) Then lstZbirni.List(lstZbirni.ListCount - 1, 1) = Format$(CLng(data(i, 2)), "#,##0")
            If IsNumeric(data(i, 3)) Then lstZbirni.List(lstZbirni.ListCount - 1, 2) = Format$(CLng(data(i, 3)), "#,##0")
            If IsNumeric(data(i, 4)) Then lstZbirni.List(lstZbirni.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.0") & " kg"
            If IsNumeric(data(i, 5)) Then lstZbirni.List(lstZbirni.ListCount - 1, 4) = Format$(CDbl(data(i, 5)), "#,##0.00") & "%"
        Else
            ' Vrsta, Kolicina, Vrednost, Prosek
            lstZbirni.List(lstZbirni.ListCount - 1, 1) = CStr(data(i, 2))
            If IsNumeric(data(i, 3)) Then lstZbirni.List(lstZbirni.ListCount - 1, 2) = Format$(CDbl(data(i, 3)), "#,##0.00")
            If IsNumeric(data(i, 4)) Then lstZbirni.List(lstZbirni.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.00")
            If IsNumeric(data(i, 5)) Then
                If CDbl(data(i, 5)) > 0 Then
                    lstZbirni.List(lstZbirni.ListCount - 1, 4) = Format$(CDbl(data(i, 5)), "#,##0.00")
                End If
            End If
        End If
    Next i
End Sub

' ============================================================
' PROSECNA CENA
' ============================================================

Private Sub GenerateProsecnaCenaReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                       ByVal datumOd As Date, ByVal datumDo As Date)
    lstProsecnaCena.Clear
    
    Dim data As Variant
    data = ReportProsecnaCena(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstProsecnaCena.AddItem CStr(data(i, 1))
        If IsNumeric(data(i, 2)) Then lstProsecnaCena.List(lstProsecnaCena.ListCount - 1, 1) = Format$(CDbl(data(i, 2)), "#,##0.00")
        If IsNumeric(data(i, 3)) Then lstProsecnaCena.List(lstProsecnaCena.ListCount - 1, 2) = Format$(CDbl(data(i, 3)), "#,##0.00")
        If IsNumeric(data(i, 4)) Then
            If CDbl(data(i, 4)) > 0 Then
                lstProsecnaCena.List(lstProsecnaCena.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.00")
            End If
        End If
    Next i
End Sub


' ============================================================
' MANJAK (NEU)
' Header: BrojZbirne | Zbirna kg | Prijemnica kg | Manjak kg | Manjak % | Prosek gajbe
' ============================================================

Private Sub GenerateManjakReport(ByVal entitetTip As String, ByVal entitetID As String, _
                                 ByVal datumOd As Date, ByVal datumDo As Date)
    lstManjak.Clear
    If lstManjak Is Nothing Then Exit Sub
    
    Dim data As Variant
    data = ReportManjak(entitetTip, entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        lstManjak.AddItem CStr(data(i, 1))
        If IsNumeric(data(i, 2)) Then lstManjak.List(lstManjak.ListCount - 1, 1) = Format$(CDbl(data(i, 2)), "#,##0.0")
        If IsNumeric(data(i, 3)) Then lstManjak.List(lstManjak.ListCount - 1, 2) = Format$(CDbl(data(i, 3)), "#,##0.0")
        If IsNumeric(data(i, 4)) Then lstManjak.List(lstManjak.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0.0")
        If IsNumeric(data(i, 5)) Then lstManjak.List(lstManjak.ListCount - 1, 4) = Format$(CDbl(data(i, 5)), "#,##0.00") & "%"
        If IsNumeric(data(i, 6)) Then
            If CDbl(data(i, 6)) > 0 Then
                lstManjak.List(lstManjak.ListCount - 1, 5) = Format$(CDbl(data(i, 6)), "#,##0.00")
            End If
        End If
    Next i
End Sub

Private Sub GenerateKarticaReport(ByVal entitetID As String, _
                                  ByVal datumOd As Date, ByVal datumDo As Date)
    lstKartica.Clear
    
    Dim ime As String, prezime As String
    ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Ime"))
    prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Prezime"))
    lblKarticaKoop.caption = "KARTICA: " & ime & " " & prezime & " (" & entitetID & ")"
    lblKarticaPeriod.caption = "Period: " & txtDatumOd.value & " - " & txtDatumDo.value
    
    Dim data As Variant
    data = ReportKarticaKooperanta(entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' Spalte 1: Datum
        If IsDate(data(i, 1)) Then
            lstKartica.AddItem Format$(CDate(data(i, 1)), "d.m.yyyy")
        Else
            lstKartica.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        End If
        
        ' Spalte 2: BrojDok
        lstKartica.List(lstKartica.ListCount - 1, 1) = _
            CStr(IIf(IsEmpty(data(i, 2)), "", data(i, 2)))
        
        ' Spalte 3: Opis (+ parcela ako postoji)
        Dim opisKartice As String
        Dim parcelaKartice As String
        
        opisKartice = CStr(IIf(IsEmpty(data(i, 4)), "", data(i, 4)))
        parcelaKartice = CStr(IIf(IsEmpty(data(i, 3)), "", data(i, 3)))
        
        If Trim$(parcelaKartice) <> "" Then
            opisKartice = opisKartice & " / Parcela: " & parcelaKartice
        End If
        
        lstKartica.List(lstKartica.ListCount - 1, 2) = opisKartice
        
        ' Spalten 4-6: Zaduzenje, Razduzenje, Saldo (novac -> uvek 2 decimale)
        Dim j As Long
        For j = 5 To 7
            If IsNumeric(data(i, j)) And Not IsEmpty(data(i, j)) Then
                lstKartica.List(lstKartica.ListCount - 1, j - 2) = Format$(CDbl(data(i, j)), "#,##0.00")
            End If
        Next j

        ' Spalte 7: Saldo ambalaze (gajbe -> ceo broj, sa znakom)
        If IsNumeric(data(i, 8)) And Not IsEmpty(data(i, 8)) Then
            lstKartica.List(lstKartica.ListCount - 1, 6) = Format$(CDbl(data(i, 8)), "#,##0")
        End If

        ' Skrivena kolona 7: ref-kljuc reda (OTK|<id> / NOV / MAG) za "Detalji otkupa"
        lstKartica.List(lstKartica.ListCount - 1, 7) = _
            CStr(IIf(IsEmpty(data(i, 9)), "", data(i, 9)))
    Next i

    ' Novi izvestaj kartice -> ocisti panel detalja (prethodni kooperant)
    KarticaDetalji_Clear
End Sub

' Tab "Pregled ambalaze": tok gajbi kooperanta (Ulaz/Izlaz/Saldo).
Private Sub GenerateKarticaAmbReport(ByVal entitetID As String, _
                                     ByVal datumOd As Date, ByVal datumDo As Date)
    If m_lstAmb Is Nothing Then Exit Sub
    m_lstAmb.Clear

    On Error Resume Next
    Dim ime As String, prezime As String
    ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Ime"))
    prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Prezime"))
    If Not m_lblAmbTitle Is Nothing Then
        m_lblAmbTitle.caption = "PREGLED AMBALA" & ChrW(381) & "E:" & ime & " " & prezime & " (" & entitetID & ")"
    End If
    On Error GoTo 0

    Dim data As Variant
    data = ReportKarticaAmbalaze(entitetID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub

    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' Datum
        If IsDate(data(i, 1)) Then
            m_lstAmb.AddItem Format$(CDate(data(i, 1)), "d.m.yyyy")
        Else
            m_lstAmb.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        End If
        ' Broj dok + Opis
        m_lstAmb.List(m_lstAmb.ListCount - 1, 1) = CStr(IIf(IsEmpty(data(i, 2)), "", data(i, 2)))
        m_lstAmb.List(m_lstAmb.ListCount - 1, 2) = CStr(IIf(IsEmpty(data(i, 3)), "", data(i, 3)))
        ' Ulaz / Izlaz (gajbe -> ceo broj; prazno kad je 0)
        If IsNumeric(data(i, 4)) And data(i, 4) <> "" Then
            If CDbl(data(i, 4)) <> 0 Then m_lstAmb.List(m_lstAmb.ListCount - 1, 3) = Format$(CDbl(data(i, 4)), "#,##0")
        End If
        If IsNumeric(data(i, 5)) And data(i, 5) <> "" Then
            If CDbl(data(i, 5)) <> 0 Then m_lstAmb.List(m_lstAmb.ListCount - 1, 4) = Format$(CDbl(data(i, 5)), "#,##0")
        End If
        ' Saldo (running)
        If IsNumeric(data(i, 6)) And data(i, 6) <> "" Then
            m_lstAmb.List(m_lstAmb.ListCount - 1, 5) = Format$(CDbl(data(i, 6)), "#,##0")
        End If
    Next i
End Sub

' Klik na red kartice -> read-only panel "Detalji otkupa" desno (sve stavke
' otkupnog lista koje se unose u frmOtkup). Samo pregled (ne za izmenu).
Private Sub lstKartica_Click()
    On Error Resume Next
    KarticaDetalji_ShowForRow Me, lstKartica
End Sub

Private Sub lstKartica_Change()
    On Error Resume Next
    KarticaDetalji_ShowForRow Me, lstKartica
End Sub

Private Sub btnStampajKarticu_Click()
    On Error GoTo EH

    If cmbEntitet.ListIndex < 0 Then
        MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim koopID As String
    koopID = ExtractIDFromDisplay(cmbEntitet.value)

    ' Tab-aware: na runtime tabu "Pregled ambalaze" -> kartica ambalaze;
    ' inace (Kartica) -> finansijska kartica kooperanta (kao do sada).
    If m_ambPageIdx >= 0 And mpReports.value = m_ambPageIdx Then
        PrintKarticaAmbalazePDF koopID, CDate(txtDatumOd.value), CDate(txtDatumDo.value)
    Else
        PrintKarticaPDF koopID, CDate(txtDatumOd.value), CDate(txtDatumDo.value)
    End If
    Exit Sub

EH:
    LogErr "frmIzvestaj.btnStampajKarticu_Click"
    
    Select Case Err.Number
        Case vbObjectError + 7501, vbObjectError + 7502
            MsgBox Err.description, vbExclamation, APP_NAME
        Case Else
            MsgBox "Gre" & ChrW(353) & "ka pri " & ChrW(353) & "tampanju kartice: " & Err.description, vbCritical, APP_NAME
    End Select
End Sub

' (REPORT-TABELLEN sekcija uklonjena -- tblRpt* se nigde ne citaju)


' ============================================================
' DRUCKEN
' ============================================================
Private Sub btnStampaj_Click()
    On Error GoTo EH
    Dim activeTab As Long
    activeTab = mpReports.value
    
    Dim lst As MSForms.ListBox
    Dim title As String
    Dim headers As Variant
    
    Select Case activeTab
        Case 0
            Set lst = lstSaldoOM
            title = "Saldo OM"
            headers = Array( _
                "Kooperant", _
                "Koli" & ChrW(269) & "ina", _
                "Vrednost", _
                "Isplaceno", _
                "Agro zadu" & ChrW(382) & "enje", _
                "Saldo", _
                "Ambala" & ChrW(382) & "a")
                
        Case 1
            Set lst = lstSaldoKupci
            title = "Saldo Kupci"
            headers = Array( _
                "Vrsta", _
                "Koli" & ChrW(269) & "ina", _
                "Cena", _
                "Vrednost", _
                "Novac", _
                "Saldo", _
                "Ambala" & ChrW(382) & "a")
                
        Case 2
            Set lst = lstOtkupRoba
            title = "Otkupljena roba"
            
            If GetActiveEntitetTip() = "Otkupna mesta" Then
                headers = Array( _
                    "Datum", _
                    "Otpremnica", _
                    "Vrsta", _
                    "Klasa", _
                    "Vozac", _
                    "Otp kg", _
                    "Blokovi kg", _
                    "Razlika kg (Otk. listovi - Otprem.)", _
                    "Prijemnica kg", _
                    "Manjak kg / %")
            Else
                headers = Array( _
                    "Nr", _
                    "Vrsta", _
                    "Koli" & ChrW(269) & "ina kg", _
                    "Vrednost", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    "")
            End If
                
        Case 3
            Set lst = lstAmbalaza
            title = "Ambala" & ChrW(382) & "a"
            headers = Array( _
                "Datum", _
                "Mesto", _
                "Tip", _
                "Dokument", _
                "Ulaz", _
                "Izlaz")
                
        Case 4
            Set lst = lstIsplata
            title = "Isplata"
            headers = Array( _
                "Kooperant", _
                "Ke" & ChrW(353) & " otkupac", _
                "Virman firma", _
                "Virman avans", _
                "Ukupno")
                
        Case 5
            Set lst = lstZbirni
            title = "Zbirni izve" & ChrW(353) & "taj"
            
            If GetActiveEntitetTip() = "Vozaci" Then
                headers = Array( _
                    "Vozac", _
                    "Amb izlaz", _
                    "Amb vracena", _
                    "Manjak kg", _
                    "Manjak %")
            Else
                headers = Array( _
                    "Entitet", _
                    "Vrsta", _
                    "Koli" & ChrW(269) & "ina", _
                    "Vrednost", _
                    "Prosek")
            End If
                
        Case 6
            Set lst = lstProsecnaCena
            title = "Prosecna cena"
            headers = Array( _
                "Vrsta", _
                "Koli" & ChrW(269) & "ina", _
                "Vrednost", _
                "Prosecna cena")
                
        Case 7
            Set lst = lstManjak
            title = "Manjak"
            headers = Array( _
                "Br. zbirne", _
                "Zbirna kg", _
                "Prijemnica kg", _
                "Manjak kg", _
                "Manjak %", _
                "Prosek gajbe")
                
        Case 8
            Set lst = lstKartica
            title = "Kartica kooperanta"
            headers = Array( _
                "Datum", _
                "Broj dok.", _
                "Opis", _
                "Zadu" & ChrW(382) & "enje", _
                "Razdu" & ChrW(382) & "enje", _
                "Saldo", _
                "Saldo amb.")
    End Select

    ' Runtime tab "Pregled ambalaze" (dinamicki index, nije u Select Case 0-8)
    If lst Is Nothing And m_ambPageIdx >= 0 And activeTab = m_ambPageIdx Then
        Set lst = m_lstAmb
        title = "Pregled ambala" & ChrW(382) & "e"
        headers = Array("Datum", "Broj dok.", "Opis", "Ulaz", "Izlaz", "Saldo")
    End If

    ' Runtime tab "Otkupni listovi" (dinamicki index; skrivena kol. 7 = ref-kljuc
    ' se NE stampa jer headers ima 7 stavki).
    If lst Is Nothing And m_otkPageIdx >= 0 And activeTab = m_otkPageIdx Then
        Set lst = m_lstOtk
        title = "Otkupni listovi"
        headers = Array("Datum", "Broj dok.", "Kooperant", "Vrsta", "Klasa", "Koli" & ChrW(269) & "ina", "Vrednost")
    End If

    If lst Is Nothing Then Exit Sub
    If lst.ListCount = 0 Then
        MsgBox "Nema podataka za " & ChrW(353) & "tampu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Stampaj samo kolone koje imaju zaglavlje (kartica ima skrivenu kolonu
    ' sa ref-kljucem koja se NE stampa); ostali tabovi: headers == ColumnCount.
    Dim colsToPrint As Long
    colsToPrint = UBound(headers) - LBound(headers) + 1
    If colsToPrint > lst.ColumnCount Then colsToPrint = lst.ColumnCount

    Dim data() As Variant
    ReDim data(1 To lst.ListCount, 1 To colsToPrint)
    Dim i As Long, j As Long
    For i = 0 To lst.ListCount - 1
        For j = 0 To colsToPrint - 1
            data(i + 1, j + 1) = lst.List(i, j)
        Next j
    Next i
    
    Dim entLabel As String
    entLabel = GetActiveEntitetTip()   ' "Otkupna mesta" / "Kupci" / "Vozaci"

    Dim entName As String
    If IsZbirniMode() Then
        entName = "Svi"
    Else
        entName = cmbEntitet.value
    End If
    Dim fullTitle As String
    fullTitle = title & " - " & entLabel & ": " & entName & _
            " (" & txtDatumOd.value & " - " & txtDatumDo.value & ")"
    
    PrintIzvestaj data, fullTitle, headers
    Exit Sub
EH:
    LogErr "frmIzvestaj.btnStampaj"
    MsgBox "Gre" & ChrW(353) & "ka pri " & ChrW(353) & "tampanju: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' NAVIGATION & HELPER
' ============================================================

Private Sub btnPovratak_Click()
    Me.Hide
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        frmOtkupAPP.Show
    End If
End Sub




Private Sub SetupAllColumnHeaders()
    On Error Resume Next
    
    ' === SaldoOM Page ===
    SetColumnHeader lbl_H_SOM1, "Kooperant"
    SetColumnHeader lbl_H_SOM2, "Koli" & ChrW(269) & "ina"
    SetColumnHeader lbl_H_SOM3, "Vrednost"
    SetColumnHeader lbl_H_SOM4, "Isplaceno"
    SetColumnHeader lbl_H_SOM5, Poruka("RPT_LBL_AGRO_ZADUZENJE")
    SetColumnHeader lbl_H_SOM6, "Saldo"
    SetColumnHeader lbl_H_SOM7, Poruka("RPT_LBL_AMBALAZA")
    
    LayoutSaldoOMHeaders

    ' === SaldoKupci Page ===
    SetColumnHeader lbl_H_SK1, "Vrsta"
    SetColumnHeader lbl_H_SK2, "Koli" & ChrW(269) & "ina"
    SetColumnHeader lbl_H_SK3, "Cena"
    SetColumnHeader lbl_H_SK4, "Vrednost"
    SetColumnHeader lbl_H_SK5, "Novac"
    SetColumnHeader lbl_H_SK6, "Saldo"
    SetColumnHeader lbl_H_SK7, Poruka("RPT_LBL_AMBALAZA")
    
    ' === OtkupRoba Page === (naslovi + sirine + OR11 -> LayoutOtkupRobaHeaders)
    LayoutOtkupRobaHeaders
    
    ' === Ambalaza Page ===
    SetColumnHeader lbl_H_AMB1, "Datum"
    SetColumnHeader lbl_H_AMB2, "Mesto"
    SetColumnHeader lbl_H_AMB3, "Tip"
    SetColumnHeader lbl_H_AMB4, "Dokument"
    SetColumnHeader lbl_H_AMB5, "Ulaz"
    SetColumnHeader lbl_H_AMB6, "Izlaz"
    
    ' === Isplata Page ===
    SetColumnHeader lbl_H_ISP1, "Kooperant"
    SetColumnHeader lbl_H_ISP2, Poruka("RPT_LBL_KES_OTKUPAC")
    SetColumnHeader lbl_H_ISP3, "Virman firma"
    SetColumnHeader lbl_H_ISP4, "Virman avans"
    SetColumnHeader lbl_H_ISP5, "Ukupno"
    
    ' === Zbirni Page ===
    SetColumnHeader lbl_H_ZB1, "Entitet"
    SetColumnHeader lbl_H_ZB2, "Vrsta"
    SetColumnHeader lbl_H_ZB3, "Koli" & ChrW(269) & "ina"
    SetColumnHeader lbl_H_ZB4, "Vrednost"
    SetColumnHeader lbl_H_ZB5, "Prosek"
    
    ' === ProsecnaCena Page ===
    SetColumnHeader lbl_H_PC1, "Vrsta"
    SetColumnHeader lbl_H_PC2, "Koli" & ChrW(269) & "ina"
    SetColumnHeader lbl_H_PC3, "Vrednost"
    SetColumnHeader lbl_H_PC4, "Prosecna cena"
    
    ' === Manjak Page ===
    SetColumnHeader lbl_H_MAN1, "Br. zbirne"
    SetColumnHeader lbl_H_MAN2, "Zbirna kg"
    SetColumnHeader lbl_H_MAN3, "Prijemnica kg"
    SetColumnHeader lbl_H_MAN4, "Manjak kg"
    SetColumnHeader lbl_H_MAN5, "Manjak %"
    SetColumnHeader lbl_H_MAN6, "Prosek gajbe"
    
    ' === KarticaKooperanta Page ===
    SetColumnHeader lbl_H_KK1, "Datum"
    SetColumnHeader lbl_H_KK2, "Broj dok."
    SetColumnHeader lbl_H_KK3, "Opis"
    SetColumnHeader lbl_H_KK4, Poruka("RPT_LBL_ZADUZENJE")
    SetColumnHeader lbl_H_KK5, Poruka("RPT_LBL_RAZDUZENJE")
    SetColumnHeader lbl_H_KK6, "Saldo"
    LayoutKarticaHeaders            ' kreira lbl_H_KK7 + auto-fit sirine/pozicije (mora poslednje)

    On Error GoTo 0
End Sub

' Kartica headeri + sirine kolona: racunaju se iz STVARNE lstKartica.width (auto-fit),
' jer header "Saldo amb." (lbl_H_KK7) ne postoji u binarnom .frx i mora se kreirati u
' runtime-u (Controls.Add, kao modKarticaDetalji; .frx se ne dira). I ColumnWidths i
' header Left/Width idu iz istog izvora (udeo sirine) -> uvek poravnato i sve kolone
' staju (ranije je amb-kolona prelivala desnu ivicu pa je header bezao na "S").
' Mora da se zove POSLE SetColumnHeader poziva (pozicioniranje je poslednje).
Private Sub LayoutKarticaHeaders()
    On Error Resume Next

    ' KK7 ("Saldo amb.") - kreiraj idempotentno (lookup po imenu).
    Dim lblAmb As MSForms.label
    Set lblAmb = lstKartica.parent.Controls("lbl_H_KK7")
    If lblAmb Is Nothing Then
        Set lblAmb = lstKartica.parent.Controls.Add("Forms.Label.1", "lbl_H_KK7", True)
        If Not lblAmb Is Nothing Then
            lblAmb.top = lbl_H_KK6.top
            lblAmb.Height = lbl_H_KK6.Height
        End If
    End If
    SetColumnHeader lblAmb, "Saldo amb."

    Dim availW As Double
    availW = lstKartica.width - 16          ' rezerva za scrollbar/border
    If availW < 140 Then Exit Sub

    ' Udeli 7 vidljivih kolona (zbir = 1.0): Datum, Br.dok, Opis, Zad, Razd, Saldo, Saldo amb.
    Dim prop As Variant
    prop = Array(0.1, 0.12, 0.2, 0.145, 0.145, 0.16, 0.13)
    Dim names As Variant
    names = Array("lbl_H_KK1", "lbl_H_KK2", "lbl_H_KK3", "lbl_H_KK4", _
                  "lbl_H_KK5", "lbl_H_KK6", "lbl_H_KK7")

    Dim X As Double: X = lstKartica.Left
    Dim cw As String: cw = ""
    Dim k As Long
    For k = 0 To 6
        Dim wCol As Long
        wCol = CLng(Int(availW * CDbl(prop(k))))
        Dim lbl As MSForms.label
        Set lbl = Nothing
        Set lbl = lstKartica.parent.Controls(CStr(names(k)))
        If Not lbl Is Nothing Then
            lbl.Left = X
            lbl.width = wCol
        End If
        cw = cw & CStr(wCol) & ";"
        X = X + wCol
    Next k
    cw = cw & "0"                            ' skrivena ref-kljuc kolona (idx 7)
    lstKartica.ColumnCount = 8
    lstKartica.ColumnWidths = cw
End Sub

' Poravnaj header labele Saldo OM tacno nad kolonama (auto-fit iz lstSaldoOM.width,
' kao LayoutKarticaHeaders) - .frx pozicije su se razilazile sa kolonama. I header
' Left/Width i lstSaldoOM.ColumnWidths idu iz istog izvora -> uvek poravnato.
Private Sub LayoutSaldoOMHeaders()
    On Error Resume Next
    Dim prop As Variant
    prop = Array(0.216, 0.117, 0.144, 0.144, 0.144, 0.144, 0.091)
    Dim names As Variant
    names = Array("lbl_H_SOM1", "lbl_H_SOM2", "lbl_H_SOM3", "lbl_H_SOM4", _
                  "lbl_H_SOM5", "lbl_H_SOM6", "lbl_H_SOM7")
    Dim availW As Double: availW = lstSaldoOM.width - 16
    If availW < 120 Then Exit Sub
    Dim X As Double: X = lstSaldoOM.Left
    Dim cw As String: cw = ""
    Dim k As Long
    For k = 0 To 6
        Dim wCol As Long: wCol = CLng(Int(availW * CDbl(prop(k))))
        Dim lbl As MSForms.label
        Set lbl = Nothing
        Set lbl = lstSaldoOM.parent.Controls(CStr(names(k)))
        If Not lbl Is Nothing Then
            lbl.Left = X
            lbl.width = wCol
        End If
        cw = cw & CStr(wCol)
        If k < 6 Then cw = cw & ";"
        X = X + wCol
    Next k
    lstSaldoOM.ColumnCount = 7
    lstSaldoOM.ColumnWidths = cw
End Sub

' Naslovi + sirine kolona za "Otkupljena roba" (OM varijanta): 10 vidljivih kolona
' (ListBox limit; Manjak kg+% spojeni u jednu kolonu, OtpremnicaID u m_otkOtpID).
' OR8 ima podnaslov u zagradi. Auto-fit iz lstOtkupRoba.width (kao
' LayoutKarticaHeaders). .frx se ne dira (CLAUDE.md).
Private Sub LayoutOtkupRobaHeaders()
    On Error Resume Next

    Dim caps As Variant
    caps = Array("Datum", "Br. Otp.", "Vrsta", "Klasa", "Vozac", "Otp kg", _
                 "Blokovi kg", "Razlika kg", _
                 "Prijemnica kg", "Manjak kg / %")
    Dim prop As Variant
    prop = Array(0.085, 0.12, 0.09, 0.05, 0.135, 0.09, 0.09, 0.085, 0.09, 0.165)
    Dim names As Variant
    names = Array("lbl_H_OR1", "lbl_H_OR2", "lbl_H_OR3", "lbl_H_OR4", "lbl_H_OR5", _
                  "lbl_H_OR6", "lbl_H_OR7", "lbl_H_OR8", "lbl_H_OR9", "lbl_H_OR10")

    Dim availW As Double
    availW = lstOtkupRoba.width - 16
    If availW < 140 Then Exit Sub

    Dim X As Double: X = lstOtkupRoba.Left
    Dim cw As String: cw = ""
    Dim k As Long
    For k = 0 To 9
        Dim wCol As Long
        wCol = CLng(Int(availW * CDbl(prop(k))))
        Dim lbl As MSForms.label
        Set lbl = Nothing
        Set lbl = lstOtkupRoba.parent.Controls(CStr(names(k)))
        If Not lbl Is Nothing Then
            lbl.Left = X
            lbl.width = wCol
            StyleListHeaderLabel lbl
            lbl.caption = CStr(caps(k))
        End If
        cw = cw & CStr(wCol)
        If k < 9 Then cw = cw & ";"
        X = X + wCol
    Next k
    lstOtkupRoba.ColumnCount = 10
    lstOtkupRoba.ColumnWidths = cw
End Sub

Private Sub SetColumnHeader(ByVal lbl As MSForms.label, ByVal txt As String)
    On Error Resume Next
    StyleListHeaderLabel lbl
    lbl.caption = txt
    On Error GoTo 0
End Sub

' Runtime tab "Pregled ambalaze" (kooperant): MultiPage page + naslov + lista.
' Geometrija kopirana sa lstKartica/lbl_H_KK1 (page-relativne koord. su iste preko
' stranica). Idempotentno (m_ambBuilt); nova instanca forme -> vars default-uju i
' page se ponovo gradi. .frx se ne dira (CLAUDE.md).
Private Sub EnsureKarticaAmbPage()
    On Error GoTo EH
    If m_ambBuilt Then Exit Sub
    If mpReports Is Nothing Then Exit Sub

    Dim pg As MSForms.Page
    Set pg = mpReports.Pages.Add("pgKarticaAmb", "Pregled ambala" & ChrW(382) & "e")
    If pg Is Nothing Then Exit Sub
    m_ambPageIdx = pg.index
    On Error Resume Next
    pg.BackColor = BG_MAIN()
    On Error GoTo EH

    Set m_lblAmbTitle = pg.Controls.Add("Forms.Label.1", "lblAmbTitle", True)
    With m_lblAmbTitle
        .Left = lstKartica.Left
        .top = lbl_H_KK1.top - 18
        If .top < 2 Then .top = 2
        .width = lstKartica.width
        .Height = 14
    End With
    StyleListHeaderLabel m_lblAmbTitle
    m_lblAmbTitle.caption = "PREGLED AMBALA" & ChrW(381) & "E"

    Set m_lstAmb = pg.Controls.Add("Forms.ListBox.1", "lstKarticaAmb", True)
    With m_lstAmb
        .Left = lstKartica.Left
        .top = lstKartica.top
        .width = lstKartica.width
        .Height = lstKartica.Height
        .ColumnCount = 6
    End With
    StyleListBox m_lstAmb

    LayoutAmbHeaders pg

    m_ambBuilt = True
    Exit Sub
EH:
    m_ambBuilt = True   ' ne pokusavaj ponovo u istoj instanci (izbegni dupli page)
    LogErr "frmIzvestaj.EnsureKarticaAmbPage"
End Sub

' Headeri (6) + sirine kolona za "Pregled ambalaze", auto-fit iz m_lstAmb.width.
Private Sub LayoutAmbHeaders(ByVal pg As MSForms.Page)
    On Error Resume Next
    Dim caps As Variant
    caps = Array("Datum", "Broj dok.", "Opis", "Ulaz", "Izlaz", "Saldo")
    Dim prop As Variant
    prop = Array(0.13, 0.18, 0.29, 0.13, 0.13, 0.14)   ' zbir = 1.0

    Dim availW As Double
    availW = m_lstAmb.width - 16
    If availW < 120 Then availW = m_lstAmb.width

    Dim X As Double: X = m_lstAmb.Left
    Dim topY As Double: topY = lbl_H_KK1.top
    Dim hh As Double: hh = lbl_H_KK1.Height
    Dim cw As String: cw = ""
    Dim k As Long
    For k = 0 To 5
        Dim wCol As Long
        wCol = CLng(Int(availW * CDbl(prop(k))))
        Dim lbl As MSForms.label
        Set lbl = Nothing
        Set lbl = pg.Controls("lblAmbH" & CStr(k + 1))
        If lbl Is Nothing Then
            Set lbl = pg.Controls.Add("Forms.Label.1", "lblAmbH" & CStr(k + 1), True)
        End If
        If Not lbl Is Nothing Then
            lbl.top = topY
            lbl.Height = hh
            lbl.Left = X
            lbl.width = wCol
            StyleListHeaderLabel lbl
            lbl.caption = CStr(caps(k))
        End If
        cw = cw & CStr(wCol) & ";"
        X = X + wCol
    Next k
    If Len(cw) > 0 Then cw = Left$(cw, Len(cw) - 1)
    m_lstAmb.ColumnWidths = cw
End Sub

' ============================================================
' Runtime tab "Otkupni listovi" (Otkupna mesta): MultiPage page + lista svih
' otkup linija stanice. Desni panel "Detalji otkupa" (modKarticaDetalji) se deli
' sa Karticom; ISPOD njega dugme "Stampaj otkupni list" (stampa ceo BrDok).
' Geometrija kopirana sa lstKartica (page-rel. koord. su iste preko stranica).
' Idempotentno (m_otkBuilt); .frx se ne dira (CLAUDE.md).
' ============================================================
Private Sub EnsureOtkupListePage()
    On Error GoTo EH
    If m_otkBuilt Then Exit Sub
    If mpReports Is Nothing Then Exit Sub

    Dim pg As MSForms.Page
    Set pg = mpReports.Pages.Add("pgOtkupListe", "Otkupni listovi")
    If pg Is Nothing Then Exit Sub
    m_otkPageIdx = pg.index
    On Error Resume Next
    pg.BackColor = BG_MAIN()
    On Error GoTo EH

    Set m_lblOtkTitle = pg.Controls.Add("Forms.Label.1", "lblOtkTitle", True)
    With m_lblOtkTitle
        .Left = lstKartica.Left
        .top = lbl_H_KK1.top - 18
        If .top < 2 Then .top = 2
        .width = lstKartica.width
        .Height = 14
    End With
    StyleListHeaderLabel m_lblOtkTitle
    m_lblOtkTitle.caption = "OTKUPNI LISTOVI"

    Set m_lstOtk = pg.Controls.Add("Forms.ListBox.1", "lstOtkupListe", True)
    With m_lstOtk
        .Left = lstKartica.Left
        .top = lstKartica.top
        .width = lstKartica.width
        .Height = lstKartica.Height
        .ColumnCount = 8          ' 7 vidljivih + skrivena ref-kljuc kolona (idx 7)
    End With
    StyleListBox m_lstOtk

    LayoutOtkListeHeaders pg

    ' Desni panel "Detalji otkupa" (deljen sa Karticom) + dugme ispod njega.
    KarticaDetalji_Ensure Me, lstKartica
    Dim pl As Double, pt As Double, pw As Double, ph As Double
    KarticaDetalji_PanelRect pl, pt, pw, ph
    If pw > 0 Then
        Set m_btnStampajOtk = Me.Controls.Add("Forms.CommandButton.1", "btnStampajOtkList", True)
        With m_btnStampajOtk
            .Left = pl
            .top = pt + ph + 6
            .width = pw
            .Height = 26
        End With
        ' Caption preko ChrW(352) = S-caron (Sh); ASCII izvor -> nezavisno od
        ' kodne strane fajla (izbegava mesanje UTF-8/CP1250 pri merge-u).
        StylePrimaryButton m_btnStampajOtk
        m_btnStampajOtk.caption = ChrW(352) & "tampaj otkupni list"
        On Error Resume Next
        m_btnStampajOtk.ZOrder 0
        On Error GoTo EH
        m_btnStampajOtk.Visible = False
    End If

    KarticaDetalji_SetVisible False   ' panel sakriven dok se ne udje na tab

    m_otkBuilt = True
    Exit Sub
EH:
    m_otkBuilt = True   ' ne pokusavaj ponovo u istoj instanci (izbegni dupli page)
    LogErr "frmIzvestaj.EnsureOtkupListePage"
End Sub

' Headeri (7 vidljivih) + sirine kolona za "Otkupni listovi" (kol. 8 = skriveni
' ref-kljuc, sirina 0). Auto-fit iz m_lstOtk.width (kao LayoutAmbHeaders).
Private Sub LayoutOtkListeHeaders(ByVal pg As MSForms.Page)
    On Error Resume Next
    Dim caps As Variant
    caps = Array("Datum", "Broj dok.", "Kooperant", "Vrsta", "Klasa", "Koli" & ChrW(269) & "ina", "Vrednost")
    Dim prop As Variant
    prop = Array(0.11, 0.15, 0.24, 0.13, 0.08, 0.13, 0.16)   ' zbir = 1.0

    Dim availW As Double
    availW = m_lstOtk.width - 16
    If availW < 120 Then availW = m_lstOtk.width

    Dim X As Double: X = m_lstOtk.Left
    Dim topY As Double: topY = lbl_H_KK1.top
    Dim hh As Double: hh = lbl_H_KK1.Height
    Dim cw As String: cw = ""
    Dim k As Long
    For k = 0 To 6
        Dim wCol As Long
        wCol = CLng(Int(availW * CDbl(prop(k))))
        Dim lbl As MSForms.label
        Set lbl = Nothing
        Set lbl = pg.Controls("lblOtkH" & CStr(k + 1))
        If lbl Is Nothing Then
            Set lbl = pg.Controls.Add("Forms.Label.1", "lblOtkH" & CStr(k + 1), True)
        End If
        If Not lbl Is Nothing Then
            lbl.top = topY
            lbl.Height = hh
            lbl.Left = X
            lbl.width = wCol
            StyleListHeaderLabel lbl
            lbl.caption = CStr(caps(k))
        End If
        cw = cw & CStr(wCol) & ";"
        X = X + wCol
    Next k
    cw = cw & "0"   ' skrivena ref-kljuc kolona (idx 7)
    m_lstOtk.ColumnWidths = cw
End Sub

' Puni "Otkupni listovi": sve otkup linije stanice (kol. 7 = ref-kljuc OTK|<id>).
Private Sub GenerateOtkupListeReport(ByVal stanicaID As String, _
                                     ByVal datumOd As Date, ByVal datumDo As Date)
    If m_lstOtk Is Nothing Then Exit Sub
    m_lstOtk.Clear
    KarticaDetalji_Clear

    On Error Resume Next
    Dim naziv As String
    naziv = CStr(LookupValue(TBL_STANICE, "StanicaID", stanicaID, "Naziv"))
    If Not m_lblOtkTitle Is Nothing Then
        m_lblOtkTitle.caption = "OTKUPNI LISTOVI: " & naziv & " (" & stanicaID & ")"
    End If
    On Error GoTo 0

    Dim data As Variant
    data = ReportOtkupListe(stanicaID, datumOd, datumDo)
    If IsEmpty(data) Then Exit Sub

    Dim i As Long
    For i = 1 To UBound(data, 1)
        ' Datum
        If IsDate(data(i, 1)) Then
            m_lstOtk.AddItem Format$(CDate(data(i, 1)), "d.m.yyyy")
        Else
            m_lstOtk.AddItem CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        End If
        m_lstOtk.List(m_lstOtk.ListCount - 1, 1) = CStr(IIf(IsEmpty(data(i, 2)), "", data(i, 2)))   ' BrDok
        m_lstOtk.List(m_lstOtk.ListCount - 1, 2) = CStr(IIf(IsEmpty(data(i, 3)), "", data(i, 3)))   ' Kooperant
        m_lstOtk.List(m_lstOtk.ListCount - 1, 3) = CStr(IIf(IsEmpty(data(i, 4)), "", data(i, 4)))   ' Vrsta
        m_lstOtk.List(m_lstOtk.ListCount - 1, 4) = CStr(IIf(IsEmpty(data(i, 5)), "", data(i, 5)))   ' Klasa
        ' Kolicina (kg)
        If IsNumeric(data(i, 6)) Then m_lstOtk.List(m_lstOtk.ListCount - 1, 5) = FmtKolicina(CDbl(data(i, 6)))
        ' Vrednost
        If IsNumeric(data(i, 7)) Then m_lstOtk.List(m_lstOtk.ListCount - 1, 6) = Format$(CDbl(data(i, 7)), "#,##0.00")
        ' Skrivena ref-kljuc kolona (idx 7) = OTK|<OtkupID>
        m_lstOtk.List(m_lstOtk.ListCount - 1, 7) = CStr(IIf(IsEmpty(data(i, 8)), "", data(i, 8)))
    Next i
End Sub

' Klik/izbor reda "Otkupni listovi" -> isti panel "Detalji otkupa" desno.
Private Sub m_lstOtk_Click()
    On Error Resume Next
    KarticaDetalji_ShowForRow Me, m_lstOtk
End Sub

Private Sub m_lstOtk_Change()
    On Error Resume Next
    KarticaDetalji_ShowForRow Me, m_lstOtk
End Sub

' Stampa otkupni list za trenutno izabran red (ceo BrDok: Klasa I + II).
Private Sub m_btnStampajOtk_Click()
    On Error GoTo EH
    Dim otkID As String
    otkID = KarticaDetalji_CurrentOtkupID()
    If Len(Trim$(otkID)) = 0 Then
        MsgBox "Izaberite otkupni list (klik na red u listi).", vbExclamation, APP_NAME
        Exit Sub
    End If
    ReprintOtkupniListByOtkupID otkID
    Exit Sub
EH:
    LogErr "frmIzvestaj.m_btnStampajOtk_Click"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi otkupnog lista: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' "Detalji" panel: klik na red Otkupljena roba / Ambalaza -> deljeni panel
' (modKarticaDetalji). Ispod panela dugme za stampu (po tabu).
' ============================================================
Private Sub lstOtkupRoba_Click()
    On Error Resume Next
    ShowOtkupRobaDetalji
End Sub
Private Sub lstOtkupRoba_Change()
    On Error Resume Next
    ShowOtkupRobaDetalji
End Sub
Private Sub ShowOtkupRobaDetalji()
    On Error Resume Next
    Dim idx As Long: idx = lstOtkupRoba.ListIndex
    Dim oid As String
    If Not m_otkOtpID Is Nothing Then
        If m_otkOtpID.Exists(CStr(idx)) Then oid = CStr(m_otkOtpID(CStr(idx)))
    End If
    KarticaDetalji_ShowOtpremnica Me, lstOtkupRoba, idx, oid
End Sub
Private Sub lstAmbalaza_Click()
    On Error Resume Next
    KarticaDetalji_ShowForRow Me, lstAmbalaza, 6     ' skrivena kol. 6 = AMB|<tip>|<id>
End Sub
Private Sub lstAmbalaza_Change()
    On Error Resume Next
    KarticaDetalji_ShowForRow Me, lstAmbalaza, 6
End Sub

' Dugmad za stampu uz deljeni panel (kreirana forma-level kao m_btnStampajOtk).
Private Sub EnsureDetaljiButtons()
    On Error GoTo EH
    KarticaDetalji_Ensure Me, lstKartica
    Dim pl As Double, pt As Double, pw As Double, ph As Double
    KarticaDetalji_PanelRect pl, pt, pw, ph
    If pw <= 0 Then Exit Sub
    If m_btnStampajOtkRoba Is Nothing Then
        Set m_btnStampajOtkRoba = Me.Controls.Add("Forms.CommandButton.1", "btnStampajOtkRoba", True)
        With m_btnStampajOtkRoba
            .Left = pl: .top = pt + ph + 6: .width = pw: .Height = 26
        End With
        StylePrimaryButton m_btnStampajOtkRoba
        m_btnStampajOtkRoba.caption = ChrW(352) & "tampaj otpremnicu"
        On Error Resume Next
        m_btnStampajOtkRoba.ZOrder 0
        On Error GoTo EH
        m_btnStampajOtkRoba.Visible = False
    End If
    If m_btnStampajAmb Is Nothing Then
        Set m_btnStampajAmb = Me.Controls.Add("Forms.CommandButton.1", "btnStampajAmbDok", True)
        With m_btnStampajAmb
            .Left = pl: .top = pt + ph + 6: .width = pw: .Height = 26
        End With
        StylePrimaryButton m_btnStampajAmb
        m_btnStampajAmb.caption = ChrW(352) & "tampaj dokument"
        On Error Resume Next
        m_btnStampajAmb.ZOrder 0
        On Error GoTo EH
        m_btnStampajAmb.Visible = False
    End If
    Exit Sub
EH:
    LogErr "frmIzvestaj.EnsureDetaljiButtons"
End Sub

' Stampa grupni otkupni list za otpremnicu izabranu u panelu "Detalji".
Private Sub m_btnStampajOtkRoba_Click()
    On Error GoTo EH
    Dim otpID As String
    otpID = KarticaDetalji_CurrentOtpremnicaID()
    If Len(Trim$(otpID)) = 0 Then
        MsgBox "Izaberite otpremnicu (klik na red u listi).", vbExclamation, APP_NAME
        Exit Sub
    End If
    OutputOtpremnicaPDF otpID
    Exit Sub
EH:
    LogErr "frmIzvestaj.m_btnStampajOtkRoba_Click"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi otpremnice: " & Err.description, vbCritical, APP_NAME
End Sub

' Stampa dokument izabranog ambalaza reda, ruta po tipu dokumenta.
Private Sub m_btnStampajAmb_Click()
    On Error GoTo EH
    Dim dokID As String, dokTip As String
    KarticaDetalji_CurrentAmbDok dokID, dokTip
    If Len(Trim$(dokID)) = 0 Then
        MsgBox "Izaberite dokument (klik na red u listi).", vbExclamation, APP_NAME
        Exit Sub
    End If
    Select Case dokTip
        Case DOK_TIP_PRIJEMNICA
            PrintPrijemnica dokID
        Case DOK_TIP_OTKUP
            ReprintOtkupniListByOtkupID dokID
        Case DOK_TIP_OTPREMNICA
            OutputOtpremnicaPDF dokID
        Case DOK_TIP_OM_IZLAZ_KOOP, DOK_TIP_OM_ULAZ_KOOP, _
             DOK_TIP_OM_IZLAZ_FIRMA, DOK_TIP_OM_ULAZ_FIRMA
            ' OM<->kooperant / OM<->firma kretanje (prazne gajbe) -> revers.
            StampajReversAmbDok dokID, dokTip
        Case Else
            MsgBox "Za tip dokumenta '" & dokTip & "' " & ChrW(353) & "tampa nije dostupna iz ovog pregleda.", _
                   vbInformation, APP_NAME
    End Select
    Exit Sub
EH:
    LogErr "frmIzvestaj.m_btnStampajAmb_Click"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi dokumenta: " & Err.description, vbCritical, APP_NAME
End Sub

' Revers (OM<->kooperant kretanje ambalaze) za izabrani ambalaza red. Argumenti se
' rekonstruisu iz dve noge ledger-a (Kooperant + Stanica) koje dele DokumentID
' (vazi i za samostalni revers i za nogu knjizenu uz otkup). prijem=True za povrat.
Private Sub StampajReversAmbDok(ByVal dokID As String, ByVal dokTip As String)
    On Error GoTo EH
    Dim d As Variant: d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Sub
    Dim cDat As Long, cTip As Long, cKol As Long, cEnt As Long
    Dim cEntTip As Long, cDok As Long, cDokTip As Long, cVoz As Long
    cDat = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    cTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    cKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    cEnt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    cEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    cDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cVoz = GetColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC)
    If cDok = 0 Or cDokTip = 0 Then Exit Sub

    Dim isFirma As Boolean
    isFirma = (dokTip = DOK_TIP_OM_IZLAZ_FIRMA Or dokTip = DOK_TIP_OM_ULAZ_FIRMA)

    Dim datum As Date, haveDatum As Boolean
    Dim tipAmb As String, omID As String, koopID As String
    Dim kolAmb As Long
    Dim revVozacID As String
    Dim i As Long
    For i = 1 To UBound(d, 1)
        If CStr(d(i, cDok)) = dokID And CStr(d(i, cDokTip)) = dokTip Then
            If Not haveDatum And IsDate(d(i, cDat)) Then
                datum = CDate(d(i, cDat)): haveDatum = True
            End If
            If Len(tipAmb) = 0 Then tipAmb = CStr(d(i, cTip))
            Dim et As String: et = CStr(d(i, cEntTip))
            If et = "Stanica" Then
                omID = CStr(d(i, cEnt))
                If isFirma Then
                    If IsNumeric(d(i, cKol)) Then kolAmb = kolAmb + CLng(d(i, cKol))
                    If cVoz > 0 And Len(revVozacID) = 0 Then revVozacID = CStr(d(i, cVoz))
                End If
            ElseIf et = "Kooperant" Then
                koopID = CStr(d(i, cEnt))
                If IsNumeric(d(i, cKol)) Then kolAmb = kolAmb + CLng(d(i, cKol))
            End If
        End If
    Next i

    If isFirma Then
        If Len(Trim$(omID)) = 0 Then
            MsgBox "Revers (firma) nije moguce rekonstruisati (nedostaje OM noga).", _
                   vbExclamation, APP_NAME
            Exit Sub
        End If
    ElseIf Len(Trim$(koopID)) = 0 Or Len(Trim$(omID)) = 0 Then
        MsgBox "Revers nije moguce rekonstruisati (nedostaje OM ili kooperant noga).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not haveDatum Then datum = Date

    Dim omNaziv As String, koopNaziv As String, vrsta As String
    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    ' Uz-otkup revers: DokumentID = otkupID -> vrsta iz otkupa; standalone -> prazno.
    vrsta = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, dokID, COL_OTK_VRSTA))

    If isFirma Then
        Dim prijemF As Boolean: prijemF = (dokTip = DOK_TIP_OM_ULAZ_FIRMA)
        Dim revVozacNaziv As String
        revVozacNaziv = Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", revVozacID, "Ime")) & " " & _
                              CStr(LookupValue(TBL_VOZACI, "VozacID", revVozacID, "Prezime")))
        OutputIzdavanjeAmbalaze datum, dokID, omNaziv, omID, _
                                revVozacNaziv, "", _
                                tipAmb, kolAmb, vrsta, prijemF, "FIRMA"
        Exit Sub
    End If

    koopNaziv = Trim$(CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime")) & " " & _
                      CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime")))
    Dim prijem As Boolean: prijem = (dokTip = DOK_TIP_OM_ULAZ_KOOP)
    OutputIzdavanjeAmbalaze datum, dokID, omNaziv, omID, koopNaziv, koopID, _
                            tipAmb, kolAmb, vrsta, prijem
    Exit Sub
EH:
    LogErr "frmIzvestaj.StampajReversAmbDok"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi reversa: " & Err.description, vbCritical, APP_NAME
End Sub

' Prijemnice (" + "-spojeni ID-jevi) za zbirnu kojoj pripada otpremnica.
Private Function PrijemniceZaOtpremnicu(ByVal otpID As String) As String
    On Error Resume Next
    Dim brZbirne As String
    brZbirne = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE))
    If Len(Trim$(brZbirne)) = 0 Then Exit Function
    Dim d As Variant: d = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(d) Then Exit Function
    d = ExcludeStornirano(d, TBL_PRIJEMNICA)
    If Not IsArray(d) Then Exit Function
    Dim cId As Long, cZb As Long
    cId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    cZb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    If cId = 0 Or cZb = 0 Then Exit Function
    Dim res As String, i As Long
    For i = 1 To UBound(d, 1)
        If CStr(d(i, cZb)) = brZbirne Then
            If Len(res) > 0 Then res = res & " + "
            res = res & CStr(d(i, cId))
        End If
    Next i
    PrijemniceZaOtpremnicu = res
End Function

Private Sub ForceDarkAllPages()
    On Error Resume Next
    
    Dim i As Long
    For i = 0 To mpReports.Pages.count - 1
        mpReports.Pages(i).BackColor = BG_MAIN()
    Next i
    
    mpReports.BackColor = BG_MAIN()
    
    On Error GoTo 0
End Sub


