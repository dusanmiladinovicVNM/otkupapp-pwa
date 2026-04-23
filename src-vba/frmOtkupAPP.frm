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

Private Sub RemoveTitleBar()
    Dim hwnd As LongPtr
    Dim style As Long

    hwnd = FindWindow("ThunderDFrame", Me.Caption)

    If hwnd <> 0 Then
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Initialize()
    ResizeMainForm
    
    'SetupShell
    SetupHeader
    SetupSidebar
    SetupButtons
    SetupCards
    
    SetupShellResponsive
End Sub

Private Sub UserForm_Activate()

    If Not mChromeRemoved Then
        Me.Caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If

    Dim warnText As String
    warnText = CheckVerwaisteDokumente()

    If warnText <> "" Then
        lblStatus.Visible = True
        lblStatus.Caption = warnText
        lblStatus.foreColor = RGB(255, 80, 80)
        lblStatus.Font.Bold = True
    Else
        lblStatus.Visible = False
    End If
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
    Me.StartUpPosition = 0
    Me.Left = 0
    Me.Top = 0
    Me.Width = Application.Width - 10
    Me.Height = Application.Height - 10
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
    lblTitleBar.Width = Me.InsideWidth
    lblTitleBar.Height = headerH

    imgLogo.Left = 12
    imgLogo.Top = 6
    imgLogo.Width = 220
    imgLogo.Height = 34

    btnMaticni.Left = Me.InsideWidth - btnMaticni.Width - 54
    btnMaticni.Top = 7

    lblClose.Left = Me.InsideWidth - 26
    lblClose.Top = 10

    ' sidebar
    contentTop = headerH + padOuter

    fraSidebar.Left = padOuter
    fraSidebar.Top = contentTop
    fraSidebar.Width = sidebarW
    fraSidebar.Height = Me.InsideHeight - contentTop - padOuter

    ' right side layout
    rightLeft = fraSidebar.Left + fraSidebar.Width + padInner
    rightW = Me.InsideWidth - rightLeft - padOuter
    rightH = Me.InsideHeight - contentTop - padOuter

    ' big card
    lblCardAlerts.Left = rightLeft
    lblCardAlerts.Top = contentTop + 4
    lblCardAlerts.Width = rightW
    lblCardAlerts.Height = rightH - summaryH - 14

    ' red status text inside big card
    lblStatus.Left = lblCardAlerts.Left + 15
    lblStatus.Top = lblCardAlerts.Top + 15
    lblStatus.Width = lblCardAlerts.Width - 30
    lblStatus.Height = lblCardAlerts.Height - 20

    ' bottom summary card
    lblCardSummary.Left = rightLeft
    lblCardSummary.Top = lblCardAlerts.Top + lblCardAlerts.Height + 10
    lblCardSummary.Width = rightW
    lblCardSummary.Height = summaryH

    LayoutSidebarButtons
End Sub

Private Sub LayoutSidebarButtons()
    Dim btnTop As Single
    Dim btnH As Single
    Dim gap As Single
    Dim leftPos As Single
    Dim btnW As Single

    leftPos = 16
    btnW = fraSidebar.Width - 32
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

    btnMargin.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnTrace.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnOpenExcel.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnSnapshot.Move leftPos, btnTop, btnW, btnH
    btnTop = btnTop + btnH + gap

    btnExit.Move leftPos, btnTop, btnW, btnH
End Sub

Private Sub SetupHeader()
    With lblTitleBar
        .BackColor = BG_TOP
        .Caption = vbNullString
        .Left = 0
        .Top = 0
        .Width = Me.InsideWidth
        .Height = 42
    End With

    With lblAppTitle
        '.Caption = APP_NAME & " v" & APP_VERSION
        .Caption = ""
        .foreColor = TXT_LIGHT
        .BackStyle = fmBackStyleTransparent
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 14
        .Left = 18
        .Top = 10
        .Width = 220
        .Height = 22
    End With
    
    With btnMaticni
        .Caption = "Maticni podaci"
        .Left = Me.InsideWidth - 170
        .Top = 7
        .Width = 125
        .Height = 28
        .BackColor = BTN_ACTIVE
        .foreColor = vbWhite
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 9
        .TakeFocusOnClick = False
    End With

    With lblClose
        .Caption = ChrW(&H2715)
        .foreColor = TXT_MUTED
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 13
        .Left = Me.InsideWidth - 34
        .Top = 12
        .Width = 20
        .Height = 20
    End With
End Sub

Private Sub SetupSidebar()
    With fraSidebar
        .Caption = vbNullString
        .BackColor = BG_PANEL
        .BorderStyle = fmBorderStyleSingle
        .Left = 18
        .Top = 58
        .Width = 235
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
    navButtons.Add btnMargin
    navButtons.Add btnTrace
    navButtons.Add btnOpenExcel
    navButtons.Add btnSnapshot
    navButtons.Add btnExit

    HighlightActive btnBlocks
End Sub

Private Sub StyleNavButton(btn As MSForms.CommandButton, txt As String, topPos As Double)
    With btn
        .Caption = "   " & txt
        .Left = 16
        .Top = topPos
        .Width = 200
        .Height = 34
        .BackColor = BTN_BG
        .foreColor = TXT_LIGHT
        .Font.Name = "Segoe UI"
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
        .Width = 520
        .Height = 380
        .BackColor = RGB(36, 42, 54)
        .BorderStyle = fmBorderStyleSingle
        .Caption = ""
    End With

    With lblStatus
        .Left = 295
        .Top = 85
        .Width = 490
        .Height = 370
        .BackStyle = fmBackStyleTransparent
        .foreColor = RGB(255, 120, 120)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ZOrder 0   ' bring to front
    End With

    With lblCardSummary
        .Left = 280
        .Top = 456
        .Width = 520
        .Height = 30
        .BackColor = RGB(36, 42, 54)
        .BorderStyle = fmBorderStyleSingle
        .Caption = ""
    End With
End Sub

Private Sub HighlightActive(activeBtn As MSForms.CommandButton)
    Dim btn As MSForms.CommandButton

    If navButtons Is Nothing Then Exit Sub

    For Each btn In navButtons
        btn.BackColor = BTN_BG
        btn.foreColor = TXT_LIGHT
    Next btn

    activeBtn.BackColor = BTN_ACTIVE
    activeBtn.foreColor = vbWhite

    lblNavAccent.Visible = True
    lblNavAccent.Top = activeBtn.Top
    lblNavAccent.Left = activeBtn.Left - 6
    lblNavAccent.Height = activeBtn.Height
End Sub

Private Sub SetHover(btn As MSForms.CommandButton, ByVal hovered As Boolean)
    If btn.BackColor <> BTN_ACTIVE Then
        If hovered Then
            btn.BackColor = BTN_HOVER
            btn.foreColor = vbWhite
        Else
            btn.BackColor = BTN_BG
            btn.foreColor = TXT_LIGHT
        End If
    End If
End Sub

' =========================
' Header actions
' =========================

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.foreColor = RGB(255, 255, 255)
End Sub

Private Sub lblTitleBar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.foreColor = TXT_MUTED
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblClose.foreColor = TXT_MUTED
    btnMaticni.BackColor = BTN_ACTIVE
    ResetHover
End Sub

Private Sub ResetHover()
    Dim btn As MSForms.CommandButton

    For Each btn In navButtons
        If btn.BackColor <> BTN_ACTIVE Then
            btn.BackColor = BTN_BG
        End If
    Next btn
End Sub

' =========================
' Navigation button clicks
' =========================

Private Sub btnBlocks_Click()
    HighlightActive btnBlocks
    lblStatus.Caption = "Sekcija: Otkupni blokovi"
    Me.Hide
    frmOtkup.Show
End Sub

Private Sub btnPurchase_Click()
    HighlightActive btnPurchase
    lblStatus.Caption = "Sekcija: Otkup i prodaja"
    Me.Hide
    frmDokumenta.Show
End Sub

Private Sub btnAgro_Click()
    HighlightActive btnAgro
    lblStatus.Caption = "Sekcija: Agrohemija"
    Me.Hide
    frmAgrohemija.Show
End Sub

Private Sub btnReports_Click()
    HighlightActive btnReports
    lblStatus.Caption = "Sekcija: Izvestaj"
    Me.Hide
    frmIzvestaj.Show
    Me.Show
End Sub

Private Sub btnInvoicing_Click()
    HighlightActive btnInvoicing
    lblStatus.Caption = "Sekcija: Fakturisanje"
    Me.Hide
    frmFakturisanje.Show
End Sub

Private Sub btnMargin_Click()
    HighlightActive btnMargin
    lblStatus.Caption = "Sekcija: Marža"
    Me.Hide
    frmMarza.Show
End Sub

Private Sub btnTrace_Click()
    HighlightActive btnTrace
    lblStatus.Caption = "Sekcija: Izveštaj o sledljivosti"
    Me.Hide
    frmSledljivost.Show
End Sub

Private Sub btnOpenExcel_Click()
    HighlightActive btnOpenExcel
    lblStatus.Caption = "Sekcija: Otvori Excel"

    Me.Hide
    Application.Visible = True
    frmExcelMini.Show vbModeless
End Sub

Private Sub btnSnapshot_Click()
    HighlightActive btnSnapshot
    lblStatus.Caption = "Sekcija: Snimi"
End Sub

Private Sub btnExit_Click()
    On Error GoTo EH

    ThisWorkbook.Save
    Application.Visible = True
    Application.Quit
    Exit Sub

EH:
    MsgBox "Fajl nije uspešno sacuvan: " & Err.Description, vbExclamation
End Sub

Private Sub btnMaticni_Click()
    OpenMaticniForm
    HighlightActive btnMaticni
    lblStatus.Caption = "Sekcija: Maticni podaci"
    'frmMaticniPodaci.Show
End Sub

Public Sub OpenMaticniForm()
    Load frmMaticniPodaci

    With frmMaticniPodaci
        .StartUpPosition = 0
        .Left = Me.Left + btnMaticni.Left + 2
        .Top = Me.Top + btnMaticni.Top + btnMaticni.Height
        
        .Show vbModeless
    End With
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
