VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtkupAPP 
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17760
   OleObjectBlob   =   "frmOtkupAPP.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmOtkupAPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private navButtons As Collection
Private isDragging As Boolean
Private dragOffsetX As Double
Private dragOffsetY As Double

Private mChromeRemoved As Boolean

Private mActiveContent As Object
' v6.11 UI
Private Const SIDEBAR_KPI_H As Single = 92
Private mIsSwitchingContent As Boolean
Private Const CONTENT_PAD As Single = 0

Private mPWASyncLog As Collection
Private mPWASyncLogActive As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo EH

    ResizeMainForm
    
    SetupShell
    SetupHeader
    SetupSidebar
    SetupButtons
    SetupCards
    
    SetupShellResponsive
    SetupTopBarStatus
    Exit Sub

EH:
    LogErr "frmOtkupAPP.UserForm_Initialize"
    MsgBox "Greška pri pokretanju glavnog ekrana: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH
    
    EnsureUserFormChromeRemoved Me, mChromeRemoved

    ' Ako je otvorena neka sekcija u desnom delu,
    ' ne prikazuj dashboard/card/status pri aktiviranju shell-a.
    If mIsSwitchingContent Or Not mActiveContent Is Nothing Then
        HideDashboardArea
        Exit Sub
    End If

    ' Dnevni kontrolni pregled (kontrolni zbirovi + problemi). Uvek vidljiv;
    ' crveno samo kada ima problema (neuskladjene kolicine / orphani / nevalidno).
    Dim imaProblema As Boolean
    Dim cardText As String
    cardText = GetKontrolaPregled(imaProblema)

    lblStatus.Visible = True
    lblStatus.caption = cardText
    If imaProblema Then
        lblStatus.ForeColor = RGB(255, 80, 80)
        lblStatus.Font.Bold = True
    Else
        lblStatus.ForeColor = RGB(220, 220, 220)
        lblStatus.Font.Bold = False
    End If
    
    ' v6.11 UI refresh
    RefreshSidebarKpi
    RefreshBankaBadge
    PositionAccentBar
    ' v6.11 UI: nakon final layout-a, repozicioniraj accent bar na aktivno dugme
    If Not navButtons Is Nothing Then
        HighlightActive btnBlocks
    End If
    Exit Sub
    
EH:
    LogErr "frmOtkupAPP.UserForm_Activate"

    On Error Resume Next
    lblStatus.Visible = True
    lblStatus.caption = "Greška pri proveri dokumenata. Pogledajte log."
    lblStatus.ForeColor = RGB(255, 80, 80)
    lblStatus.Font.Bold = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        ShutdownApp
    End If
End Sub

' =========================
' UI
' =========================

Private Sub SetupShell()
    Me.BackColor = BG_MAIN()
End Sub

Private Sub ResizeMainForm()
    On Error GoTo EH

    Dim targetW As Single
    Dim targetH As Single

    Me.StartUpPosition = 0
    Me.Left = 0
    Me.Top = 0

    ' Ne oslanjamo se na Application.Width/Height jer je Excel sakriven.
    targetW = ScreenWidthPoints()
    targetH = ScreenHeightPoints()

    ' Mali odbitak da forma ne ode ispod taskbara / ivica.
    targetW = targetW
    targetH = targetH - 30

    If targetW < 900 Then targetW = 900
    If targetH < 600 Then targetH = 600

    Me.width = targetW
    Me.Height = targetH

    Exit Sub

EH:
    LogErr "frmOtkupAPP.ResizeMainForm"

    ' fallback, ali tek ako WinAPI ne uspe
    Me.StartUpPosition = 0
    Me.Left = 0
    Me.Top = 0
    Me.width = 1200
    Me.Height = 750
End Sub

Private Sub SetupShellResponsive()
    Dim padOuter As Single
    Dim padInner As Single
    Dim headerH As Single
    Dim sidebarW As Single
    Dim contentTop As Single
    Dim rightLeft As Single
    Dim rightW As Single
    Dim rightH As Single
    Dim summaryH As Single

    padOuter = 12
    padInner = 18
    headerH = 44
    sidebarW = 240
    summaryH = 30

    ' header
    lblTitleBar.Left = 0
    lblTitleBar.Top = 0
    lblTitleBar.width = Me.InsideWidth
    lblTitleBar.Height = headerH

    imgLogo.Left = 12
    imgLogo.Top = 6
    imgLogo.width = 220
    imgLogo.Height = 34

    btnMaticni.Left = Me.InsideWidth - btnMaticni.width - 54
    btnMaticni.Top = 7

    lblClose.Left = Me.InsideWidth - 26
    lblClose.Top = 10

    ' sidebar
    contentTop = headerH + padOuter

    fraSidebar.Left = padOuter
    fraSidebar.Top = contentTop
    fraSidebar.width = sidebarW
    fraSidebar.Height = Me.InsideHeight - contentTop - padOuter

    ' right side layout
    rightLeft = fraSidebar.Left + fraSidebar.width + padInner
    rightW = Me.InsideWidth - rightLeft - padOuter
    rightH = Me.InsideHeight - contentTop - padOuter
    If rightW < 100 Then rightW = 100
    If rightH < 100 Then rightH = 100

    ' big card
    lblCardAlerts.Left = rightLeft
    lblCardAlerts.Top = contentTop + 4
    lblCardAlerts.width = rightW
    lblCardAlerts.Height = rightH - summaryH - 14
    If lblCardAlerts.Height < 50 Then lblCardAlerts.Height = 50

    ' red status text inside big card
    lblStatus.Left = lblCardAlerts.Left + 15
    lblStatus.Top = lblCardAlerts.Top + 15
    lblStatus.width = lblCardAlerts.width - 30
    lblStatus.Height = lblCardAlerts.Height - 20

    ' bottom summary card
    lblCardSummary.Left = rightLeft
    lblCardSummary.Top = lblCardAlerts.Top + lblCardAlerts.Height + 10
    lblCardSummary.width = rightW
    lblCardSummary.Height = summaryH

    LayoutSidebarButtons
    FitActiveContent
End Sub

Private Sub LayoutSidebarButtons()
    Dim btnTop As Single
    Dim btnH As Single
    Dim gap As Single
    Dim leftPos As Single
    Dim btnW As Single

    leftPos = 16
    btnW = fraSidebar.width - 32
    btnH = 34
    gap = 10
    btnTop = 18

    btnBlocks.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnPurchase.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnAgro.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnReports.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnInvoicing.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap
    
    btnBanka.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap
    
    btnBankaIsplate.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap
    
    btnSyncPWA.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnMargin.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnTrace.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnOpenExcel.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnSnapshot.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnExit.Move leftPos, btnTop, btnW, btnH
    
    ' v6.11 UI: pomeri accent bar na aktivno nav dugme
    PositionAccentBar

    ' v6.11 UI: layout KPI kartice na dnu sidebara
    LayoutSidebarKpiCard
End Sub

Private Sub SetSidebarEnabled(ByVal enabled As Boolean)
    Dim btn As Object

    If navButtons Is Nothing Then Exit Sub

    On Error Resume Next

    For Each btn In navButtons
        btn.enabled = enabled
    Next btn

    btnMaticni.enabled = enabled

    On Error GoTo 0
End Sub

Private Sub SetupHeader()
    With lblTitleBar
        .BackColor = BG_TOP
        .caption = vbNullString
        .Left = 0
        .Top = 0
        .width = Me.InsideWidth
        .Height = 42
    End With

    With lblAppTitle
        '.Caption = APP_NAME & " v" & APP_VERSION
        .caption = ""
        .ForeColor = TXT_LIGHT
        .BackStyle = fmBackStyleTransparent
        .Font.name = "Segoe UI Semibold"
        .Font.Size = 14
        .Left = 18
        .Top = 10
        .width = 220
        .Height = 22
    End With
    
    With btnMaticni
        .caption = "Maticni podaci"
        .Left = Me.InsideWidth - 170
        .Top = 7
        .width = 125
        .Height = 28
        .BackColor = BTN_ACTIVE
        .ForeColor = vbWhite
        .Font.name = "Segoe UI Semibold"
        .Font.Size = 9
        .TakeFocusOnClick = False
    End With

    With lblClose
        .caption = ChrW(&H2715)
        .ForeColor = TXT_MUTED
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .Font.name = "Segoe UI Symbol"
        .Font.Size = 13
        .Left = Me.InsideWidth - 34
        .Top = 12
        .width = 20
        .Height = 20
    End With
End Sub

Private Sub SetupSidebar()
    With fraSidebar
        .caption = vbNullString
        .BackColor = BG_PANEL
        .BorderStyle = fmBorderStyleSingle
        .Left = 18
        .Top = 58
        .width = 235
        .Height = Me.InsideHeight - 76
    End With
End Sub

Private Sub SetupButtons()
    Dim topPos As Double
    topPos = 20

    StyleNavButton btnBlocks, "Otkupni blokovi", topPos
    topPos = topPos + 42

    StyleNavButton btnPurchase, "Otkup i prodaja", topPos
    topPos = topPos + 42

    StyleNavButton btnAgro, "Agrohemija", topPos
    topPos = topPos + 42

    StyleNavButton btnReports, "Izveštaj", topPos
    topPos = topPos + 42

    StyleNavButton btnInvoicing, "Fakturisanje", topPos
    topPos = topPos + 42
    
    StyleNavButton btnBanka, "Banka uvoz izvoda", topPos
    topPos = topPos + 42
    
    StyleNavButton btnBankaIsplate, "Banka platni nalozi", topPos
    topPos = topPos + 42
    
    StyleNavButton btnSyncPWA, "Sinhronizuj PWA", topPos
    topPos = topPos + 42

    StyleNavButton btnMargin, "Marža", topPos
    topPos = topPos + 42

    StyleNavButton btnTrace, "Izveštaj o sledljivosti", topPos
    topPos = topPos + 42

    StyleNavButton btnOpenExcel, "Otvori Excel", topPos
    topPos = topPos + 42

    StyleNavButton btnSnapshot, "Snimi", topPos
    topPos = topPos + 42

    StyleNavButton btnExit, "Izlaz", topPos

    Set navButtons = New Collection
    navButtons.Add btnBlocks
    navButtons.Add btnPurchase
    navButtons.Add btnAgro
    navButtons.Add btnReports
    navButtons.Add btnInvoicing
    navButtons.Add btnBanka
    navButtons.Add btnBankaIsplate
    navButtons.Add btnSyncPWA
    navButtons.Add btnMargin
    navButtons.Add btnTrace
    navButtons.Add btnOpenExcel
    navButtons.Add btnSnapshot
    navButtons.Add btnExit

    HighlightActive btnBlocks
End Sub

Private Sub StyleNavButton(btn As MSForms.CommandButton, txt As String, topPos As Double)
    With btn
        .caption = "   " & txt
        .Left = 16
        .Top = topPos
        .width = 200
        .Height = 34
        .BackColor = BTN_BG
        .ForeColor = TXT_ON_DARK
        .Font.name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = False
        .TakeFocusOnClick = False

        On Error Resume Next
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
        On Error GoTo 0
    End With
End Sub

Private Sub SetupCards()
    With lblCardAlerts
        .Left = 280
        .Top = 70
        .width = 520
        .Height = 380
        .BackColor = RGB(36, 42, 54)
        .BorderStyle = fmBorderStyleSingle
        .caption = ""
    End With

    With lblStatus
        .Left = 295
        .Top = 85
        .width = 490
        .Height = 370
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(255, 120, 120)
        .Font.name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ZOrder 0   ' bring to front
    End With

    With lblCardSummary
        .Left = 280
        .Top = 456
        .width = 520
        .Height = 30
        .BackColor = RGB(36, 42, 54)
        .BorderStyle = fmBorderStyleSingle
        .caption = ""
    End With
End Sub

Private Sub HighlightActive(activeBtn As MSForms.CommandButton)
    Dim btn As MSForms.CommandButton
    If navButtons Is Nothing Then Exit Sub

    For Each btn In navButtons
        btn.BackColor = BTN_BG
        btn.ForeColor = TXT_ON_DARK
    Next btn

    activeBtn.BackColor = BTN_ACTIVE
    activeBtn.ForeColor = vbWhite

    ' Accent bar — koristi koordinate iste kao activeBtn (oba su child of fraSidebar)
    StyleAccentBar lblNavAccent
    With lblNavAccent
        .Visible = True
        .width = 4
        .Height = activeBtn.Height
        .Left = activeBtn.Left - 8
        .Top = activeBtn.Top
        .ZOrder 0
    End With
End Sub

Private Sub SetHover(btn As MSForms.CommandButton, ByVal hovered As Boolean)
    If btn.BackColor <> BTN_ACTIVE Then
        If hovered Then
            btn.BackColor = BTN_HOVER
            btn.ForeColor = vbWhite
        Else
            btn.BackColor = BTN_BG
            btn.ForeColor = TXT_ON_DARK
        End If
    End If
End Sub

' =========================
' Header actions
' =========================

Private Sub lblClose_Click()
    On Error GoTo EH

    ShutdownApp
    Exit Sub

EH:
    LogErr "frmOtkupAPP.lblClose_Click"
    Application.Visible = True
End Sub

Private Sub lblClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub lblTitleBar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.ForeColor = TXT_MUTED
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.ForeColor = TXT_MUTED
    btnMaticni.BackColor = BTN_ACTIVE
    ResetHover
End Sub

Private Sub ResetHover()
    On Error GoTo EH

    If navButtons Is Nothing Then Exit Sub

    Dim btn As MSForms.CommandButton

    For Each btn In navButtons
        If btn.BackColor <> BTN_ACTIVE Then
            btn.BackColor = BTN_BG
        End If
    Next btn

    Exit Sub

EH:
    LogErr "frmOtkupAPP.ResetHover"
End Sub

' =========================
' Navigation button clicks
' =========================

Private Sub btnBlocks_Click()
    OpenContentForm frmOtkup, btnBlocks, "Otkupni blokovi"
End Sub

Private Sub btnPurchase_Click()
    OpenContentForm frmDokumenta, btnPurchase, "Otkup i prodaja"
End Sub

Private Sub btnAgro_Click()
    OpenContentForm frmAgrohemija, btnAgro, "Agrohemija"
End Sub

Private Sub btnReports_Click()
    OpenContentForm frmIzvestaj, btnReports, "Izvestaji"
End Sub

Private Sub btnInvoicing_Click()
    OpenContentForm frmFakturisanje, btnInvoicing, "Fakturisanje"
End Sub

Private Sub btnBanka_Click()
    On Error GoTo EH

    Dim oldPointer As Integer
    oldPointer = Me.MousePointer

    HighlightActive btnBanka

    Me.MousePointer = fmMousePointerHourGlass
    btnBanka.enabled = False

    lblStatus.Visible = True
    lblStatus.caption = "Uvozim nove bankovne izvode..."
    lblStatus.ForeColor = RGB(220, 220, 220)
    lblStatus.Font.Bold = False

    ImportBankaInbox_TX

    lblStatus.caption = "Banka uvezena. Otvaram mapiranje..."

    OpenContentForm frmBankaImport, btnBanka, "Banka uvoz izvoda"

CleanExit:
    btnBanka.enabled = True
    Me.MousePointer = oldPointer
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description

    LogErr "frmOtkupAPP.btnBanka_Click"

    On Error Resume Next
    lblStatus.Visible = True
    lblStatus.caption = "Greška pri uvozu banke. Otvaram mapiranje za postojece stavke."
    lblStatus.ForeColor = RGB(255, 80, 80)
    lblStatus.Font.Bold = True

    MsgBox "Greška pri uvozu bankovnih izvoda:" & vbCrLf & errDesc & vbCrLf & vbCrLf & _
           "Mapiranje ce biti otvoreno za postojece neuparene stavke.", vbExclamation

    OpenContentForm frmBankaImport, btnBanka, "Banka uvoz izvoda"

    Resume CleanExit
End Sub

Private Sub btnbankaisplate_Click()
    OpenContentForm frmBankaExportPregled, btnBankaIsplate, "Banke platni nalozi"
End Sub

Private Sub btnSyncPWA_Click()
    On Error GoTo EH

    Dim oldPointer As Integer
    Dim ok As Boolean

    oldPointer = Me.MousePointer

    HighlightActive btnSyncPWA

    Me.MousePointer = fmMousePointerHourGlass
    SetSidebarEnabled False

    BeginPWASyncLog
    DoEvents

    ok = SyncPWAFullCycle_ForButton()

    FinishPWASyncLog ok

    On Error Resume Next
    RefreshSidebarKpi
    RefreshBankaBadge
    On Error GoTo 0

CleanExit:
    SetSidebarEnabled True
    Me.MousePointer = oldPointer
    Exit Sub

EH:
    LogErr "frmOtkupAPP.btnSyncPWA_Click"

    On Error Resume Next
    FinishPWASyncLog False
    Resume CleanExit
End Sub

Private Sub btnMargin_Click()
    OpenContentForm frmMarza, btnMargin, "Marža"
End Sub

Private Sub btnTrace_Click()
    HighlightActive btnTrace
    lblStatus.caption = "Sekcija: Izveštaj o sledljivosti"
    OpenContentForm frmSledljivost, btnTrace, "Izveštaj o sledljivosti"
End Sub

Private Sub btnOpenExcel_Click()
    HighlightActive btnOpenExcel
    lblStatus.caption = "Sekcija: Otvori Excel"

    Me.Hide
    Application.Visible = True
    frmExcelMini.Show vbModeless
End Sub

Private Sub btnSnapshot_Click()
    On Error GoTo EH

    HighlightActive btnSnapshot

    lblStatus.Visible = True
    lblStatus.caption = "Snimam podatke..."
    lblStatus.ForeColor = RGB(220, 220, 220)
    lblStatus.Font.Bold = False

    SaveApp

    lblStatus.caption = "Sacuvano."
    lblStatus.ForeColor = RGB(120, 220, 140)
    lblStatus.Font.Bold = True

    MsgBox "Sacuvano!", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmOtkupAPP.btnSnapshot_Click"

    lblStatus.Visible = True
    lblStatus.caption = "Greška pri snimanju. Pogledajte log."
    lblStatus.ForeColor = RGB(255, 80, 80)
    lblStatus.Font.Bold = True

    MsgBox "Greška pri snimanju: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnExit_Click()
    On Error GoTo EH

    ThisWorkbook.Save
    ShutdownApp

    Application.DisplayAlerts = False
    Application.Quit

CleanExit:
    On Error Resume Next
    Application.DisplayAlerts = True
    On Error GoTo 0
    Exit Sub

EH:
    LogErr "frmOtkupAPP.btnExit_Click"

    On Error Resume Next
    Application.DisplayAlerts = True
    On Error GoTo 0

    MsgBox "Fajl nije uspešno sacuvan ili aplikacija nije zatvorena: " & Err.description, vbExclamation, APP_NAME
End Sub

Private Sub btnMaticni_Click()
    On Error GoTo EH

    HighlightActive btnMaticni
    lblStatus.caption = "Sekcija: Maticni podaci"
    OpenMaticniForm

    Exit Sub

EH:
    LogErr "frmOtkupAPP.btnMaticni_Click"
    MsgBox "Greška pri otvaranju maticnih podataka: " & Err.description, vbCritical, APP_NAME
End Sub

Public Sub OpenMaticniForm()
    On Error GoTo EH

    Load frmMaticniPodaci

    With frmMaticniPodaci
        .StartUpPosition = 0
        .Left = Me.Left + btnMaticni.Left + 2
        .Top = Me.Top + btnMaticni.Top + btnMaticni.Height
        .Show vbModeless
    End With

    Exit Sub

EH:
    LogErr "frmOtkupAPP.OpenMaticniForm"
    MsgBox "Greška pri otvaranju maticnih podataka: " & Err.description, vbCritical, APP_NAME
End Sub

' =========================
' Hover states
' =========================

Private Sub btnBlocks_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnBlocks, True
End Sub

Private Sub btnPurchase_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnPurchase, True
End Sub

Private Sub btnAgro_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnAgro, True
End Sub

Private Sub btnReports_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnReports, True
End Sub

Private Sub btnInvoicing_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnInvoicing, True
End Sub

Private Sub btnBanka_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnBanka, True
End Sub

Private Sub btnBankaIsplate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnBankaIsplate, True
End Sub

Private Sub btnSyncPWA_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)
    ResetHover
    SetHover btnSyncPWA, True
End Sub

Private Sub btnMargin_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnMargin, True
End Sub

Private Sub btnTrace_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnTrace, True
End Sub

Private Sub btnOpenExcel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnOpenExcel, True
End Sub

Private Sub btnSnapshot_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnSnapshot, True
End Sub

Private Sub btnExit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetHover
    SetHover btnExit, True
End Sub

Private Sub btnMaticni_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnMaticni.BackColor = BTN_HOVER
End Sub

Private Sub CloseActiveContent()
    On Error Resume Next

    If Not mActiveContent Is Nothing Then
        Unload mActiveContent
    End If

    Set mActiveContent = Nothing

    On Error GoTo 0
End Sub

Private Sub FitActiveContent()
    On Error GoTo EH

    If mActiveContent Is Nothing Then Exit Sub

    With mActiveContent
        .StartUpPosition = 0

        ' Poravnanje sa desnim content delom
        .Left = Me.Left + lblCardAlerts.Left
        .Top = Me.Top + fraSidebar.Top

        ' Pokriva i lblCardAlerts i lblCardSummary,
        ' i ravna se gore/dole sa sidebarom
        .width = lblCardAlerts.width
        .Height = fraSidebar.Height
    End With

    Exit Sub

EH:
    LogErr "frmOtkupAPP.FitActiveContent"
End Sub

Public Sub OpenContentFormPublic(ByVal contentForm As Object, _
                                  ByVal sectionTitle As String)
    On Error GoTo EH
    
    ' Highlight nije moguc bez parent nav button-a (jer SEF nema sidebar dugme)
    ' Zato koristimo trenutni nav highlight koji je vec aktivan (btnInvoicing)
    OpenContentForm contentForm, btnInvoicing, sectionTitle
    
    Exit Sub
    
EH:
    LogErr "frmOtkupAPP.OpenContentFormPublic"
End Sub

Private Sub OpenContentForm(ByVal contentForm As Object, _
                            ByVal activeBtn As MSForms.CommandButton, _
                            ByVal sectionTitle As String)
    On Error GoTo EH

    Dim oldContent As Object
    Dim oldPointer As Integer

    oldPointer = Me.MousePointer
    Me.MousePointer = fmMousePointerHourGlass

    mIsSwitchingContent = True

    Set oldContent = mActiveContent

    HideDashboardArea
    HighlightActive activeBtn

    If Not oldContent Is Nothing Then
        If oldContent Is contentForm Then
            FitActiveContent
            GoTo CleanExit
        End If
    End If

    Set mActiveContent = contentForm

    FitActiveContent
    mActiveContent.Show vbModeless
    FitActiveContent

    If Not oldContent Is Nothing Then
        modJournaling.FlushNow "section-switch"   ' <-- DODATO: persistuj staru sekciju
        Unload oldContent
    End If

    HideDashboardArea
    FitActiveContent

CleanExit:
    mIsSwitchingContent = False
    Me.MousePointer = oldPointer
    Exit Sub

EH:
    mIsSwitchingContent = False
    Me.MousePointer = oldPointer

    LogErr "frmOtkupAPP.OpenContentForm"

    ShowDashboardArea "Greška pri otvaranju sekcije: " & sectionTitle

    lblStatus.ForeColor = RGB(255, 80, 80)
    lblStatus.Font.Bold = True

    MsgBox "Greška pri otvaranju sekcije '" & sectionTitle & "': " & Err.description, vbCritical, APP_NAME
End Sub


Public Sub ReturnToDashboard(Optional ByVal statusText As String = "")
    On Error Resume Next

    If mIsSwitchingContent Then Exit Sub
    
    modJournaling.FlushNow "return-to-dashboard"   ' <-- DODATO

    Set mActiveContent = Nothing

    ShowDashboardArea statusText

    On Error GoTo 0
End Sub

Private Sub HideDashboardArea()
    On Error Resume Next

    lblStatus.Visible = False
    lblCardAlerts.Visible = False
    lblCardSummary.Visible = False

    On Error GoTo 0
End Sub

Private Sub ShowDashboardArea(Optional ByVal statusText As String = "")
    On Error Resume Next

    lblCardAlerts.Visible = True
    lblCardSummary.Visible = True
    lblStatus.Visible = True

    If Len(statusText) > 0 Then
        lblStatus.caption = statusText
    Else
        lblStatus.caption = ""
    End If

    lblStatus.ForeColor = RGB(220, 220, 220)
    lblStatus.Font.Bold = False

    On Error GoTo 0
End Sub

' ============================================================
' v6.11 UI: ACCENT BAR (gold strip on active nav button)
' ============================================================

Private Sub PositionAccentBar()
    On Error GoTo EH

    ' Trazi aktivno dugme (BackColor = BTN_ACTIVE)
    Dim btn As MSForms.CommandButton
    Dim activeBtn As MSForms.CommandButton

    If navButtons Is Nothing Then Exit Sub

    For Each btn In navButtons
        If btn.BackColor = BTN_ACTIVE() Then
            Set activeBtn = btn
            Exit For
        End If
    Next btn

    If activeBtn Is Nothing Then
        On Error Resume Next
        lblNavAccent.Visible = False
        On Error GoTo 0
        Exit Sub
    End If

    On Error Resume Next
    With lblNavAccent
        .Visible = True
        StyleAccentBar lblNavAccent
        .width = 4
        .Height = activeBtn.Height
        .Left = fraSidebar.Left + activeBtn.Left - 8
        .Top = fraSidebar.Top + activeBtn.Top
        .ZOrder 0
    End With
    On Error GoTo 0

    Exit Sub

EH:
    LogErr "frmOtkupAPP.PositionAccentBar"
End Sub

' ============================================================
' v6.11 UI: KPI CARD u dnu sidebara (Današnji otkup kg)
' ============================================================

Private Sub LayoutSidebarKpiCard()
    On Error GoTo EH

    ' fraSidebarKpi je novi Frame (vidi Designer instrukcije)
    Dim leftPos As Single
    Dim btnW As Single

    leftPos = 16
    btnW = fraSidebar.width - 32

    With fraSidebarKPI
        .Move leftPos, fraSidebar.Height - SIDEBAR_KPI_H - 12, btnW, SIDEBAR_KPI_H
        StyleKpiCardFrame fraSidebarKPI
    End With

    ' Layout label-a unutar kartice
    With lblKPITitle
        .Move 10, 8, fraSidebarKPI.width - 20, 14
        StyleKpiLabel lblKPITitle, "title"
        .caption = "Današnji otkup"
    End With

    With lblKPIValue
        .Move 10, 24, fraSidebarKPI.width - 20, 28
        StyleKpiLabel lblKPIValue, "value"
    End With

    With lblKPIDelta
        .Move 10, 54, fraSidebarKPI.width - 20, 14
        StyleKpiLabel lblKPIDelta, "delta_up"
    End With

    With lblKPISub
        .Move 10, 70, fraSidebarKPI.width - 20, 14
        StyleKpiLabel lblKPISub, "muted"
    End With

    RefreshSidebarKpi

    Exit Sub

EH:
    LogErr "frmOtkupAPP.LayoutSidebarKpiCard"
End Sub

Public Sub RefreshSidebarKpi()
    On Error GoTo EH

    Dim today As Date
    today = Date

    Dim totalKgToday As Double
    Dim totalKgYesterday As Double
    Dim cntOtp As Long
    Dim cntPrij As Long

    totalKgToday = SumOtkupKgForDate(today)
    totalKgYesterday = SumOtkupKgForDate(today - 1)

    cntOtp = CountDocsForDate(TBL_OTPREMNICA, COL_OTP_DATUM, today)
    cntPrij = CountDocsForDate(TBL_PRIJEMNICA, COL_PRJ_DATUM, today)

    lblKPIValue.caption = Format$(totalKgToday, "#,##0") & " kg"

    If totalKgYesterday > 0 Then
        Dim deltaPct As Double
        deltaPct = ((totalKgToday - totalKgYesterday) / totalKgYesterday) * 100

        If deltaPct >= 0 Then
            lblKPIDelta.caption = ChrW(&H25B2) & " " & Format$(deltaPct, "0.0") & "%  vs juce"
            StyleKpiLabel lblKPIDelta, "delta_up"
        Else
            lblKPIDelta.caption = ChrW(&H25BC) & " " & Format$(Abs(deltaPct), "0.0") & "%  vs juce"
            StyleKpiLabel lblKPIDelta, "delta_down"
        End If
    Else
        lblKPIDelta.caption = ""
    End If

    lblKPISub.caption = cntOtp & " otpremnica  •  " & cntPrij & " prijemnica"

    Exit Sub

EH:
    LogErr "frmOtkupAPP.RefreshSidebarKpi"
    lblKPIValue.caption = "—"
    lblKPIDelta.caption = ""
    lblKPISub.caption = ""
End Sub

' --- helperi za KPI ---
Private Function SafeDateKey(ByVal v As Variant) As String
    On Error GoTo BadValue

    If isError(v) Then GoTo BadValue
    If IsEmpty(v) Then GoTo BadValue
    If Len(Trim$(CStr(v))) = 0 Then GoTo BadValue

    If IsDate(v) Then
        SafeDateKey = Format$(CDate(v), "yyyy-mm-dd")
    Else
        SafeDateKey = vbNullString
    End If

    Exit Function

BadValue:
    SafeDateKey = vbNullString
End Function

Private Function SafeKpiDouble(ByVal v As Variant) As Double
    On Error GoTo BadValue

    If isError(v) Then GoTo BadValue
    If IsEmpty(v) Then GoTo BadValue
    If Len(Trim$(CStr(v))) = 0 Then GoTo BadValue

    If IsNumeric(v) Then
        SafeKpiDouble = CDbl(v)
    Else
        SafeKpiDouble = 0
    End If

    Exit Function

BadValue:
    SafeKpiDouble = 0
End Function
Private Function SumOtkupKgForDate(ByVal targetDate As Date) As Double
    On Error GoTo EH

    Dim data As Variant
    Dim colDate As Long
    Dim colKg As Long
    Dim i As Long
    Dim targetKey As String
    Dim rowKey As String

    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    colDate = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    colKg = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)

    If colDate <= 0 Or colKg <= 0 Then Exit Function

    targetKey = Format$(targetDate, "yyyy-mm-dd")

    For i = LBound(data, 1) To UBound(data, 1)
        rowKey = SafeDateKey(data(i, colDate))

        If rowKey = targetKey Then
            SumOtkupKgForDate = SumOtkupKgForDate + SafeKpiDouble(data(i, colKg))
        End If
    Next i

    Exit Function

EH:
    LogErr "frmOtkupAPP.SumOtkupKgForDate"
    SumOtkupKgForDate = 0
End Function

Private Function CountDocsForDate(ByVal tableName As String, _
                                  ByVal dateColName As String, _
                                  ByVal targetDate As Date) As Long
    On Error GoTo EH

    Dim data As Variant
    Dim colDate As Long
    Dim i As Long
    Dim targetKey As String
    Dim rowKey As String

    data = GetTableData(tableName)
    If IsEmpty(data) Then Exit Function

    colDate = GetColumnIndex(tableName, dateColName)
    If colDate <= 0 Then Exit Function

    targetKey = Format$(targetDate, "yyyy-mm-dd")

    For i = LBound(data, 1) To UBound(data, 1)
        rowKey = SafeDateKey(data(i, colDate))
        If rowKey = targetKey Then
            CountDocsForDate = CountDocsForDate + 1
        End If
    Next i

    Exit Function

EH:
    LogErr "frmOtkupAPP.CountDocsForDate"
    CountDocsForDate = 0
End Function

' ============================================================
' v6.11 UI: BADGE brojac za neobradene Banka redove
' ============================================================

Public Sub RefreshBankaBadge()
    On Error GoTo EH

    Dim openCount As Long
    openCount = CountOpenBankaImport()

    StyleBadge lblBankaBadge, openCount

    If openCount > 0 Then
        ' Pozicija — desno gornji ugao btnBanka
        With lblBankaBadge
            .width = 22
            .Height = 14
            .Left = fraSidebar.Left + btnBanka.Left + btnBanka.width - 36
            .Top = fraSidebar.Top + btnBanka.Top - 46
            .ZOrder 0
        End With
    End If

    Exit Sub

EH:
    LogErr "frmOtkupAPP.RefreshBankaBadge"
End Sub

Private Function CountOpenBankaImport() As Long
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Function

    Dim colObr As Long
    On Error Resume Next
    colObr = RequireColumnIndex(TBL_BANKA_IMPORT, "Obradjeno", "CountOpenBankaImport")
    On Error GoTo EH

    If colObr <= 0 Then Exit Function

    Dim i As Long, cnt As Long

    For i = 1 To UBound(data, 1)
        If LCase$(NzToText(data(i, colObr))) <> "da" Then
            cnt = cnt + 1
        End If
    Next i

    CountOpenBankaImport = cnt
    Exit Function

EH:
    LogErr "CountOpenBankaImport"
    CountOpenBankaImport = 0
End Function

' ============================================================
' v6.11 UI: TOP BAR STATUS (Online dot + Operator + Sezona)
' ============================================================

Private Sub SetupTopBarStatus()
    On Error GoTo EH

    ' Status dot
    StyleStatusDot lblStatusDot, True
    With lblStatusDot
        .width = 14
        .Height = 18
        .Left = imgLogo.Left + imgLogo.width + 12
        .Top = 12
    End With

    ' "Online" text
    With lblOnlineText
        .caption = "Online"
        .BackStyle = fmBackStyleTransparent
        .ForeColor = CLR_SUCCESS()
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_SMALL
        .Left = lblStatusDot.Left + lblStatusDot.width + 2
        .Top = 14
        .width = 50
        .Height = 16
    End With

    ' Operator
    With lblOperator
        .caption = "Operator: " & GetCurrentOperatorName()
        .BackStyle = fmBackStyleTransparent
        .ForeColor = TXT_MUTED()
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_SMALL
        .Left = lblOnlineText.Left + lblOnlineText.width + 14
        .Top = 14
        .width = 200
        .Height = 16
    End With

    ' Sezona
    With lblSezona
        .caption = "Sezona " & Year(Date)
        .BackStyle = fmBackStyleTransparent
        .ForeColor = TXT_MUTED()
        .Font.name = APP_FONT
        .Font.Size = FONT_SIZE_SMALL
        .Left = lblOperator.Left + lblOperator.width + 14
        .Top = 14
        .width = 100
        .Height = 16
    End With

    Exit Sub

EH:
    LogErr "frmOtkupAPP.SetupTopBarStatus"
End Sub

Private Function GetCurrentOperatorName() As String
    On Error Resume Next
    GetCurrentOperatorName = Environ$("USERNAME")
    If GetCurrentOperatorName = "" Then GetCurrentOperatorName = "—"
    On Error GoTo 0
End Function


Public Sub BeginPWASyncLog()
    Set mPWASyncLog = New Collection
    mPWASyncLogActive = True

    lblStatus.Visible = True
    lblStatus.caption = ""
    lblStatus.ForeColor = RGB(220, 220, 220)
    lblStatus.Font.Bold = False

    On Error Resume Next
    lblStatus.WordWrap = True
    lblStatus.AutoSize = False
    On Error GoTo 0

    AppendPWASyncLog "Pokrecem PWA / Google sync..."
End Sub

Public Sub AppendPWASyncLog(ByVal message As String)
    If Not mPWASyncLogActive Then Exit Sub

    Dim lineText As String
    Dim txt As String
    Dim i As Long

    If mPWASyncLog Is Nothing Then Set mPWASyncLog = New Collection

    lineText = Format$(Now, "hh:nn:ss") & "  " & CStr(message)
    mPWASyncLog.Add lineText

    Do While mPWASyncLog.count > 8
        mPWASyncLog.Remove 1
    Loop

    For i = 1 To mPWASyncLog.count
        If Len(txt) > 0 Then txt = txt & vbCrLf
        txt = txt & CStr(mPWASyncLog(i))
    Next i

    lblStatus.caption = txt
    DoEvents
End Sub

Public Sub FinishPWASyncLog(ByVal ok As Boolean)
    If ok Then
        AppendPWASyncLog "Sync zavrsen uspesno."
        lblStatus.ForeColor = RGB(120, 220, 140)
        lblStatus.Font.Bold = True
    Else
        AppendPWASyncLog "Sync zavrsen sa greskom. Proveri log."
        lblStatus.ForeColor = RGB(255, 90, 90)
        lblStatus.Font.Bold = True
    End If

    DoEvents
    mPWASyncLogActive = False
End Sub
