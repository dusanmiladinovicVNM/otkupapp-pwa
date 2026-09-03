VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Prijava"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' frmLogin / prijava operatera (modAuth.Login je prikazuje modalno)
' Ugovor sa modAuth: Public LoginOK; validacija ide kroz
' modAuth.ValidateLogin; tri neuspela pokusaja zatvaraju formu.
'
' Izgled: PREKO CELOG EKRANA -- ista forest podloga kao splash, a prijava
' je KARTICA u sredini. Razlog nije ukras: prijava je prvo sto operater
' vidi posle splash-a, a mali dijalog nad sakrivenim Excelom je izgledao
' kao da aplikacija nije ni startovala.
'
' Kartica prati jezik ljuske: krem povrsina, zlatna nit, natpisi verzalom,
' polja u shell-u (ivica + ispuna; focus/error kroz modUiKit.ShellState),
' zeleno primarno dugme. .frx nema logotip, pa znak stoji kao tekst -- i to
' je JEDINI znak na ekranu.
'
' Kontrole iz dizajnera se samo stilizuju i premestaju; podloga, kartica i
' shell-ovi su runtime (modUiKit), pa se .frx ne dira. Nema module-level
' MSForms deklaracija (meka forma). Sve mere su u tackama.
'
' Stil kontrola iz dizajnera ide kroz POSTOJECE primitive ljuske
' (PanelStilNatpis / PanelStilNapomena / PanelStilDugme) -- iste one koje
' oblace panele Podesavanja i Admin. Ovde su samo GEOMETRIJA (Postavi) i
' polje u shell-u (PostaviUnos), jer prijava nije panel nego forma za sebe.
' ============================================================

Private Const BANDS   As Long = 40         ' trake vertikalnog gradijenta
Private Const CARD_W  As Single = 344      ' kartica prijave
Private Const CARD_H  As Single = 290
Private Const FLD_H   As Single = 28       ' = FIELD_H ljuske
Private Const FLD_PAD As Single = 9        ' = INPUT_PAD ljuske
Private Const MAX_ATT As Long = 3

Public LoginOK As Boolean
Private mAttempts As Long
Private mChromeRemoved As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo EH

    Me.caption = APP_NAME & " - Prijava"
    mChromeRemoved = False
    Me.LoginOK = False
    mAttempts = 0

    BuildLogin

    Exit Sub

EH:
    LogErr "frmLogin.UserForm_Initialize"
End Sub

Private Sub BuildLogin()
    Dim i As Long, fnt As String, w As Single, h As Single, bh As Single
    Dim cx As Single, cy As Single, iw As Single
    fnt = DisplayFont()

    ' Ceo ekran (isti racun kao splash; Excel je sakriven pa Application.Width
    ' ne vredi), kartica u sredini.
    w = ScreenWidthPoints()
    h = ScreenHeightPoints()
    If w < 600 Then w = 600
    If h < 400 Then h = 400

    Me.StartUpPosition = 0
    Me.Left = 0
    Me.top = 0
    Me.width = w
    Me.Height = h
    Me.BackColor = C_FOREST

    ' vertikalni gradijent -- iza svega
    bh = h / BANDS
    For i = 0 To BANDS - 1
        NewLbl Me, "lgGr" & i, "", 0, i * bh, w, bh + 1, 8, False, 0, _
               Lerp(C_FOREST, C_FOREST_DK, i / (BANDS - 1))
        Me.Controls("lgGr" & i).ZOrder 1
    Next i

    cx = (w - CARD_W) / 2
    cy = (h - CARD_H) / 2 - 20          ' malo iznad sredine -- optican centar
    iw = CARD_W - 2 * PAD               ' unutrasnja sirina kartice

    ' kartica: ivica + krem povrsina + zlatna nit na vrhu
    NewLbl Me, "lgB", "", cx, cy, CARD_W, CARD_H, 8, False, 0, C_BORDER
    NewLbl Me, "lgF", "", cx + 1, cy + 1, CARD_W - 2, CARD_H - 2, 8, False, 0, C_CREAM
    NewLbl Me, "lgTop", "", cx + 1, cy + 1, CARD_W - 2, 3, 8, False, 0, C_GOLD

    ' znak: "AX" zlatno + "OtkupApp" forest, u display fontu
    NewLbl Me, "lgAX", "AX", cx + PAD, cy + 26, 34, TxtH(20), 20, True, C_GOLD, -1, _
           fmTextAlignLeft, fnt
    NewLbl Me, "lgName", "OtkupApp", cx + PAD + 32, cy + 28, 180, TxtH(18), 18, True, _
           C_FOREST, -1, fmTextAlignLeft, fnt

    ' naslov u display fontu (PanelStilNaslov je TS_H1 -- ovde je naslov ekrana,
    ' pa ide TS_DISPLAY, isto kao naslov u ljusci)
    lblTitle.caption = Poruka("OTKUI_LOGIN_NASLOV")
    lblTitle.ForeColor = C_FOREST
    lblTitle.Font.name = fnt
    lblTitle.Font.Size = TS_DISPLAY
    lblTitle.Font.bold = True
    Postavi lblTitle, cx + PAD, cy + 66, iw, TS_DISPLAY

    lblSubtitle.caption = Poruka("OTKUI_LOGIN_PODNASLOV")
    PanelStilNapomena lblSubtitle
    Postavi lblSubtitle, cx + PAD, cy + 92, iw, TS_META

    ' polja: natpis iznad (verzal, kao u formi), shell + TextBox uvucen za FLD_PAD
    lblUser.caption = Poruka("KOR_LBL_KORISNICKO_IME")
    PanelStilNatpis lblUser                      ' sam podize u verzal
    Postavi lblUser, cx + PAD, cy + 116, iw, TS_LABEL
    NewShell Me, "shUser", cx + PAD, cy + 130, iw, FLD_H, C_INPUT_BORDER, C_WHITE
    PostaviUnos txtUser, cx + PAD, cy + 130, iw

    lblPin.caption = Poruka("OTKUI_LOGIN_PIN")
    PanelStilNatpis lblPin
    Postavi lblPin, cx + PAD, cy + 168, iw, TS_LABEL
    NewShell Me, "shPin", cx + PAD, cy + 182, iw, FLD_H, C_INPUT_BORDER, C_WHITE
    PostaviUnos txtPin, cx + PAD, cy + 182, iw
    txtPin.PasswordChar = ChrW(8226)   ' bullet -> maskiran PIN

    ' greska ispod polja, rust kao pilula greske u mrezi
    lblErr.caption = ""
    PanelStilNapomena lblErr
    lblErr.ForeColor = C_RUST
    Postavi lblErr, cx + PAD, cy + 218, iw, TS_META

    ' dugmad: primarno levo, "Otkazi" tiho desno; Enter = prijava, Esc = otkaz
    With btnOK
        .caption = Poruka("OTKUI_LOGIN_PRIJAVA")
        .Left = cx + PAD: .top = cy + 240: .width = 160: .Height = 30
        .Default = True
        .ZOrder 0
    End With
    PanelStilDugme btnOK, "primary"
    With btnCancel
        .caption = Poruka("OTKUI_LOGIN_OTKAZI")
        .Left = cx + CARD_W - PAD - 104: .top = cy + 240: .width = 104: .Height = 30
        .Cancel = True
        .ZOrder 0
    End With
    PanelStilDugme btnCancel, "ghost"

    ' tiha linija ispod kartice -- isti kredit kao na splash-u
    NewLbl Me, "lgBy", "Powered by AgriX", cx, cy + CARD_H + 14, CARD_W, TxtH(TS_MICRO), _
           TS_MICRO, False, RGB(178, 190, 172), -1, fmTextAlignCenter
End Sub

' Geometrija natpisa iz dizajnera. Stil je vec postavljen (PanelStil*); ovde je
' samo mesto, sirina i visina linije, plus ZOrder iznad podloge kartice.
Private Sub Postavi(lbl As Object, ByVal X As Single, ByVal Y As Single, _
                    ByVal w As Single, ByVal fs As Single)
    With lbl
        .TextAlign = fmTextAlignLeft
        .WordWrap = False
        .Left = X: .top = Y: .width = w: .Height = TxtH(fs)
        .ZOrder 0
    End With
End Sub

' TextBox iz dizajnera UNUTAR shell-a -- iste mere kao NewTxt u NewFieldG
' (uvucen za FLD_PAD, 4pt od vrha shell-a, visina FLD_H - 8). Zato NE ide kroz
' PanelStilUnos: taj primitiv crta SVOJU ivicu, a ovde ivicu nosi shell -- dve
' ivice jedna u drugoj su tacno ono sto shell resava.
Private Sub PostaviUnos(t As Object, ByVal shellX As Single, ByVal shellY As Single, ByVal shellW As Single)
    With t
        .Left = shellX + FLD_PAD
        .top = shellY + 4
        .width = shellW - 2 * FLD_PAD
        .Height = FLD_H - 8
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .BackStyle = fmBackStyleOpaque
        .BackColor = C_WHITE
        .ForeColor = C_FOREST
        .Font.name = F_UI
        .Font.Size = TS_BODY
        .Font.bold = False
        .TextAlign = fmTextAlignLeft
        .ZOrder 0
    End With
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    txtUser.SetFocus
End Sub

' focus/error stanje shell-a, kao u formi ljuske
Private Sub txtUser_Enter()
    ShellState Me, "shUser", "focus"
End Sub

Private Sub txtUser_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ShellState Me, "shUser", "normal"
End Sub

Private Sub txtPin_Enter()
    ShellState Me, "shPin", "focus"
End Sub

Private Sub txtPin_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ShellState Me, "shPin", "normal"
End Sub

' hover dugmadi (kontrole iz dizajnera imaju svoje evente; WithEvents nema)
Private Sub btnOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PanelStilDugmeHover btnOK, "primary", True
    PanelStilDugmeHover btnCancel, "ghost", False
End Sub

Private Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PanelStilDugmeHover btnCancel, "ghost", True
    PanelStilDugmeHover btnOK, "primary", False
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PanelStilDugmeHover btnOK, "primary", False
    PanelStilDugmeHover btnCancel, "ghost", False
End Sub

Private Sub btnOK_Click()
    On Error GoTo EH
    lblErr.caption = ""

    If modAuth.ValidateLogin(txtUser.text, txtPin.text) Then
        Me.LoginOK = True
        Me.Hide
        Exit Sub
    End If

    mAttempts = mAttempts + 1
    txtPin.text = ""
    If mAttempts >= MAX_ATT Then
        Me.LoginOK = False
        Me.Hide
        Exit Sub
    End If

    lblErr.caption = Poruka("AUTH_LBL_PRIJAVA_GRESKA") & " (" & mAttempts & "/" & MAX_ATT & ")"
    ShellState Me, "shPin", "error"
    txtUser.SetFocus
    Exit Sub
EH:
    LogErr "frmLogin.btnOK_Click"
End Sub

Private Sub btnCancel_Click()
    Me.LoginOK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Me.LoginOK = False
End Sub
