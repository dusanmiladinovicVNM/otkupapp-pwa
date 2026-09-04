Attribute VB_Name = "modUiFaze"
Option Explicit

' ============================================================
' modUiFaze - splash, prijava i "Excel je otvoren" kao FAZE JEDNE LJUSKE
'
' Do v6-ui-213 su to bile tri zasebne forme (frmSplash, frmLogin,
' frmExcelMini). Svaka je nosila svoju kopiju istog jezika: gradijent od 40
' traka, zlatnu nit, znak "AX OtkupApp", shell polja, primarno dugme -- a
' kopija istog pravila na cetiri mesta se razidje prvom doradom (isto pravilo
' kao PanelStilDugmeHover u par.26 kataloga).
'
' Sada postoji JEDAN prozor (frmOtkupUI) i cetiri faze u njemu:
'
'   BOOT   ceo ekran, znak + verzija + "Pokrecem aplikaciju..."   (bivsi frmSplash)
'   LOGIN  ista podloga, kartica prijave 344x290 u sredini        (bivsi frmLogin)
'   MINI   kartica 232x78 gore desno dok je Excel otkriven        (bivsi frmExcelMini)
'   APP    ljuska (modOtkupUI) -- podrazumevana faza
'
' ZASTO FAZA, A NE EKRAN REGISTRA: ekran iz modUiScreens ima stavku u
' sidebaru, oblast i pravo, i zivi UNUTAR hroma ljuske. Ove tri pokrivaju CEO
' prozor -- i zaglavlje i sidebar -- i nijedna nije navigaciona. Registar bi
' im dao stavku menija koja ne sme da postoji.
'
' SVE KONTROLE SU RUNTIME (modUiKit), u okviru "zFaza" preko celog prozora.
' Nijedna nova WithEvents deklaracija: dugmad i polja se zicaju kroz
' WireBtn / WireInput -> clsFlatBtn, a tagovi krecu sa "fz", pa ih
' modOtkupUI.UiEvent prosledjuje ovamo (FazaEvent) pre svega ostalog.
'
' ZNAK JE TEKSTUALAN, ne slika. Rasterski logotip je ziveo u .frx-u splash-a i
' mini kartice, a .frx se ne pravi iz koda (CLAUDE.md par.3): kontrole se
' dodaju runtime-om. Isti tekstualni znak vec nosi zaglavlje ljuske (hdrAX /
' hdrName) i kartica prijave, pa je ovo jedini znak koji aplikacija ima -- ne
' cetvrti pored tri.
'
' Fajl mora ostati 100% ASCII.
' ============================================================

Public Const FAZA_APP   As String = "APP"
Public Const FAZA_BOOT  As String = "BOOT"
Public Const FAZA_LOGIN As String = "LOGIN"
Public Const FAZA_MINI  As String = "MINI"

Private Const BANDS    As Long = 40       ' trake vertikalnog gradijenta
Private Const CARD_W   As Single = 344    ' kartica prijave
Private Const CARD_H   As Single = 290
Private Const FLD_H    As Single = 28     ' = FIELD_H ljuske
Private Const FLD_PAD  As Single = 9      ' = INPUT_PAD ljuske
Private Const MAX_ATT  As Long = 3        ' tri pokusaja, kao u frmLogin
Private Const MINI_W   As Single = 232
Private Const MINI_H   As Single = 78
Private Const FOOT_H   As Single = 52     ' podnozje splash-a: linija + dva reda
Private Const LOGO_FS  As Single = 30     ' znak na punom ekranu

Private mFaza As String
Private mForma As Object          ' ziva instanca ljuske (i u testu, gde nije default)
Private mZiva As Boolean          ' forma jos postoji -- prekida DoEvents petlje

Private mPrijavaOK As Boolean
Private mPrijavaCeka As Boolean
Private mUPrijavi As Boolean      ' brana od ugnjezdene prijave
Private mPokusaji As Long

' Prigusen tekst na forest podlozi -- ista vrednost kao hdrStat u zaglavlju.
Private Function MutedOnForest() As Long
    MutedOnForest = RGB(178, 190, 172)
End Function

'------------------------------------------------------------- stanje ----

Public Function Faza() As String
    If Len(mFaza) = 0 Then mFaza = FAZA_APP
    Faza = mFaza
End Function

' Da li je prozor zauzet necim sto NIJE ljuska. Citaju je globalni tasteri i
' zatvaranje: F1-F8 ne smeju da rade dok je na ekranu prijava ili splash.
Public Function FazaAktivna() As Boolean
    FazaAktivna = (Faza() <> FAZA_APP)
End Function

' Sme li UserForm_Initialize da izgradi ljusku.
'
' SAMO U FAZI APP. U prijavi gradnja ne sme jer cita registar ekrana i prava
' operatera (ScrAktivan, OsveziAlatke), a operatera u tom trenutku JOS NEMA --
' izgradjena ljuska bi dobila prava prazne sesije i zapamtila ih.
'
' Ni u splash-u ne gradi, iako bi smela: red ostaje isti koji je frmSplash
' imao -- prvo znak na ekranu, pa gradnja iza njega (ShowOtkupUI ->
' OtkupUI_EnsureShellBuilt). Gradnja pre prvog piksela bi splash-u oduzela
' bas ono zbog cega postoji.
Public Function FazaGradiLjusku() As Boolean
    FazaGradiLjusku = (Faza() = FAZA_APP)
End Function

' Forma je ozivela. Zove je modOtkupUI.OtkupUI_Init iz UserForm_Initialize.
Public Sub FazaOtvorena(frm As Object)
    Set mForma = frm
    mZiva = True
End Sub

' Forma je oborena (Terminate). Sve petlje moraju da stanu -- inace DoEvents
' vrti prazan hod nad prozorom koga vise nema.
Public Sub FazaOtpusti()
    On Error Resume Next
    mZiva = False
    mPrijavaCeka = False
    mFaza = FAZA_APP
    Set mForma = Nothing
End Sub

'-------------------------------------------------------------- ulazi ----

' Splash: pokazi znak i cekaj zadato vreme. Ljuska se gradi TEK posle ovoga
' (modOtkupUI.ShowOtkupUI), isto kao sto je frmSplash radio.
Public Sub FazaBoot(ByVal sekundi As Double)
    Dim kraj As Date
    On Error GoTo EH
    mFaza = FAZA_BOOT
    Uvedi
    Postavi
    kraj = DateAdd("s", sekundi, Now)
    Do While Now < kraj
        DoEvents
        If Not mZiva Then Exit Do
    Loop
    Exit Sub
EH:
    LogErr "modUiFaze.FazaBoot"
End Sub

' Prijava. Vraca True kad je operater prijavljen -- ugovor je isti koji je
' modAuth.Login imao sa frmLogin (frmLogin.LoginOK).
'
' Ljuska je bez modalnosti (ShowModal = False), pa se cekanje pravi rukom:
' kartica pokriva prozor, zone ljuske se GASE (ne samo prekriju -- prekriven
' Frame i dalje prima Tab), i vrti se DoEvents dok dugme ne postavi ishod.
' Isti postupak koji je frmSplash koristio za svoje dve sekunde.
Public Function FazaPrijava() As Boolean
    Dim prethodna As String, bilaPrikazana As Boolean
    On Error GoTo EH

    ' Druga prijava preko prve nije zamena operatera nego dva cekanja na istom
    ' steku. Pozivalac dobija "nije prijavljen", sto je i tacno.
    If mUPrijavi Then Exit Function
    mUPrijavi = True

    prethodna = Faza()
    bilaPrikazana = Vidljiva()

    mPokusaji = 0
    mPrijavaOK = False
    mPrijavaCeka = True
    mFaza = FAZA_LOGIN
    Uvedi
    Postavi
    Fokus "fzlUser"

    Do While mPrijavaCeka
        DoEvents
        If Not mZiva Then Exit Do
    Loop

    FazaPrijava = mPrijavaOK
    mFaza = FAZA_APP

    ' Odakle je prijava pozvana, tamo se i vraca:
    '  - zamena operatera u ljusci (prozor je vec bio na ekranu) -> nazad u APP;
    '  - start (prozor smo otvorili mi) -> SAKRIJ, tacno kao "Unload frmLogin".
    '    Iza prijave u modMain.StartApp ide first-run kapija, koja trazi VIDLJIV
    '    i upotrebljiv Excel; prazna forest podloga preko celog ekrana bi je
    '    pokrila.
    If bilaPrikazana And prethodna = FAZA_APP Then
        Postavi
    Else
        FazaSakrij
    End If

    mUPrijavi = False
    Exit Function
EH:
    LogErr "modUiFaze.FazaPrijava"
    mPrijavaCeka = False
    mUPrijavi = False
    mFaza = FAZA_APP
    FazaPrijava = False
End Function

' "Otvori Excel": prozor se SKUPLJA na karticu umesto da se sakrije. Ljuska
' ostaje izgradjena, pa je povratak trenutan -- isto obrazlozenje po kome se
' forma na X sakriva a ne unload-uje.
Public Sub FazaMini()
    On Error GoTo EH
    mFaza = FAZA_MINI
    Uvedi
    Postavi
    Exit Sub
EH:
    LogErr "modUiFaze.FazaMini"
End Sub

' Nazad u aplikaciju: puna velicina, zone ljuske vracene, ljuska izgradjena.
'
' AktivirajLjusku se zove RUKOM: forma se izmedju faza ne prikazuje ponovo
' (nikad nije ni sakrivena), pa UserForm_Activate ne puca -- a bez njega mreza
' ostaje prazna i fokus van polja.
Public Sub FazaApp()
    On Error GoTo EH
    mFaza = FAZA_APP
    Uvedi
    modOtkupUI.OtkupUI_EnsureShellBuilt
    Postavi
    modOtkupUI.OtkupUI_AktivirajLjusku mForma
    Exit Sub
EH:
    LogErr "modUiFaze.FazaApp"
End Sub

' Sakrij prozor i vrati fazu na APP. Zove je i ShowOtkupUI kad nijedan ekran
' nije dozvoljen: splash ne sme da ostane na ekranu iznad poruke o odbijanju.
Public Sub FazaSakrij()
    On Error Resume Next
    mFaza = FAZA_APP
    If mForma Is Nothing Then Exit Sub
    mForma.Hide
End Sub

'------------------------------------------------------- forma javlja ----

' UserForm_Activate. True = faza je preuzela aktivaciju (ljuska ne radi nista).
Public Function FazaAktiviraj(frm As Object) As Boolean
    On Error Resume Next
    Set mForma = frm
    mZiva = True
    If Not FazaAktivna() Then Exit Function
    ' Mini kartica NE ide kroz GoFullScreen: on prvo razvuce prozor na ceo
    ' ekran, pa bi svaki klik na karticu bljesnuo punim ekranom pre nego sto se
    ' vrati na 232x78. Naslovna traka je vec skinuta pri prvom punom prikazu.
    If Faza() <> FAZA_MINI Then modOtkupUI.GoFullScreen frm
    Postavi
    FazaAktiviraj = True
End Function

' UserForm_Resize. True = faza je rasporedila prozor sama.
Public Function FazaRaspored(frm As Object) As Boolean
    On Error Resume Next
    If Not FazaAktivna() Then Exit Function
    Set mForma = frm
    Postavi
    FazaRaspored = True
End Function

' Sistemski X / Alt+F4. True = faza je preuzela zatvaranje.
Public Function FazaZatvori() As Boolean
    Select Case Faza()
        Case FAZA_BOOT
            FazaZatvori = True            ' splash se ne prekida
        Case FAZA_LOGIN
            mPrijavaOK = False            ' isto sto radi "Otkazi"
            mPrijavaCeka = False
            FazaZatvori = True
        Case FAZA_MINI
            NazadUAplikaciju
            FazaZatvori = True
    End Select
End Function

' Taster dok faza drzi prozor. True = potrosen.
Public Function FazaTaster(ByVal KeyCode As Long) As Boolean
    Select Case Faza()
        Case FAZA_LOGIN
            Select Case KeyCode
                Case vbKeyReturn: Prijavi: FazaTaster = True
                Case vbKeyEscape: Otkazi: FazaTaster = True
            End Select
        Case FAZA_BOOT, FAZA_MINI
            ' Nijedan taster ljuske (F1-F8, Esc) ne sme da radi ispod zavese.
            FazaTaster = True
    End Select
End Function

' Dogadjaj kontrole faze. Prosledjuje ga modOtkupUI.UiEvent za svaki tag koji
' pocinje sa "fz" -- pre nego sto dodirne ijednu kontrolu ljuske, koje u fazi
' prijave jos ni nema.
Public Sub FazaEvent(ByVal tag As String, ByVal ev As String, ByVal arg As Variant)
    On Error Resume Next
    Select Case ev
        Case "Click":   FazaKlik tag
        Case "Focus":   Shell_ tag, "focus"
        Case "Blur":    Shell_ tag, "normal"
        Case "KeyDown": FazaTaster CLng(arg)
    End Select
End Sub

Public Sub FazaKlik(ByVal tag As String)
    On Error GoTo EH
    Select Case tag
        Case "fzlOK":     Prijavi
        Case "fzlCancel": Otkazi
        Case "fzmBack":   NazadUAplikaciju
    End Select
    Exit Sub
EH:
    LogErr "modUiFaze.FazaKlik"
End Sub

'------------------------------------------------------------ radnje ----

Private Sub Prijavi()
    Dim z As Object
    On Error GoTo EH
    If Faza() <> FAZA_LOGIN Then Exit Sub
    Set z = Zona()
    If z Is Nothing Then Exit Sub

    z.Controls("fzlErr").caption = ""

    If modAuth.ValidateLogin(z.Controls("fzlUser").text, z.Controls("fzlPin").text) Then
        mPrijavaOK = True
        mPrijavaCeka = False
        Exit Sub
    End If

    mPokusaji = mPokusaji + 1
    z.Controls("fzlPin").text = ""
    If mPokusaji >= MAX_ATT Then
        mPrijavaOK = False
        mPrijavaCeka = False
        Exit Sub
    End If

    z.Controls("fzlErr").caption = Poruka("AUTH_LBL_PRIJAVA_GRESKA") & _
                                   " (" & mPokusaji & "/" & MAX_ATT & ")"
    ' REDOSLED: fokus PA stanje. SetFocus okida Blur PIN polja, koji shell vraca
    ' na "normal" -- rust okvir bi nestao istog trena. Isto kao u frmLogin.
    Fokus "fzlUser"
    modUiKit.ShellState z, "fzlShP", "error"
    Exit Sub
EH:
    LogErr "modUiFaze.Prijavi"
    mPrijavaOK = False
    mPrijavaCeka = False
End Sub

Private Sub Otkazi()
    mPrijavaOK = False
    mPrijavaCeka = False
End Sub

Private Sub NazadUAplikaciju()
    On Error GoTo EH
    Application.Visible = False
    FazaApp
    Exit Sub
EH:
    LogErr "modUiFaze.NazadUAplikaciju"
    On Error Resume Next
    Application.Visible = False
    FazaApp
End Sub

'------------------------------------------------------------ gradnja ----

Private Function Zona() As Object
    On Error Resume Next
    If mForma Is Nothing Then Exit Function
    Set Zona = mForma.Controls("zFaza")
    Err.Clear
End Function

' Ziva i VIDLJIVA forma + izgradjene kontrole faze.
'
' Redosled je bitan: .Show okida UserForm_Initialize, koje kroz OtkupUI_Init
' postavlja mForma -- pa se zFaza pravi TEK posle toga.
'
' Vidljivost se cita sa same
' forme, ne iz zastavice: X ljuske je sakriva mimo ovog modula
' (OtkupUI_Sakrij), pa bi zapamceno "vec je prikazana" ostavilo Alt+F8 bez
' ijednog prozora.
Private Sub Uvedi()
    On Error Resume Next
    If Not Vidljiva() Then frmOtkupUI.show vbModeless
    Gradi
End Sub

Private Function Vidljiva() As Boolean
    On Error Resume Next
    If mForma Is Nothing Then Exit Function
    Vidljiva = mForma.Visible
    Err.Clear
End Function

' Kontrole se prave JEDNOM po zivotu forme i posle se samo pale, gase i
' pomeraju -- isti razlog zbog koga BuildNav gradi obe sekcije sidebara
' odjednom: ziv event-sink (clsFlatBtn) se u ovom projektu ne rusi i ne pravi
' u hodu. "Vec izgradjeno" se cita sa SAME zone, ne iz zastavice: zastavica
' preziveli Unload/Release i drugi poziv bi pao na duplo ime kontrole.
Private Sub Gradi()
    Dim z As Object
    On Error GoTo EH
    If mForma Is Nothing Then Exit Sub

    Set z = Zona()
    If z Is Nothing Then
        ' NewZone, ne NewFrame: zona se pravi UGASENA. Kad se ljuska izgradi
        ' prva (start bez prijave), zFaza nastaje POSLE nje i time iznad nje --
        ' vidljiv prazan okvir bi na trenutak prekrio ceo ekran.
        Set z = modUiKit.NewZone(mForma, "zFaza", 0, 0, 400, 300, C_FOREST)
    ElseIf z.Controls.count > 0 Then
        Exit Sub
    End If

    GradiPodlogu z
    GradiBoot z
    GradiPrijavu z
    GradiMini z
    Exit Sub
EH:
    LogErr "modUiFaze.Gradi"
End Sub

' Forest gradijent + zlatna nit. Deli je BOOT i LOGIN; MINI je gasi.
Private Sub GradiPodlogu(z As Object)
    Dim i As Long
    For i = 0 To BANDS - 1
        modUiKit.NewLbl z, "fzpGr" & i, "", 0, i * 10, 100, 11, 8, False, 0, _
                        modUiKit.Lerp(C_FOREST, C_FOREST_DK, i / (BANDS - 1))
    Next i
    modUiKit.NewLbl z, "fzpLn", "", 0, 0, 100, 3, 8, False, 0, C_GOLD
End Sub

Private Sub GradiBoot(z As Object)
    Dim fnt As String
    fnt = modUiKit.DisplayFont()
    modUiKit.NewLbl z, "fzbAX", "AX", 0, 0, 60, modUiKit.TxtH(LOGO_FS), LOGO_FS, True, _
                    C_GOLD, -1, fmTextAlignRight, fnt
    modUiKit.NewLbl z, "fzbName", "OtkupApp", 0, 0, 240, modUiKit.TxtH(LOGO_FS), LOGO_FS, True, _
                    C_CREAM, -1, fmTextAlignLeft, fnt
    modUiKit.NewLbl z, "fzbVer", "v" & APP_VERSION, 0, 0, 300, modUiKit.TxtH(TS_META), TS_META, _
                    False, MutedOnForest(), -1, fmTextAlignCenter
    modUiKit.NewLbl z, "fzbDiv", "", 0, 0, 100, 1, 8, False, 0, C_HDR_EDGE
    modUiKit.NewLbl z, "fzbBy", "Powered by AgriX", 0, 0, 200, modUiKit.TxtH(TS_MICRO), TS_MICRO, _
                    False, MutedOnForest(), -1, fmTextAlignLeft
    modUiKit.NewLbl z, "fzbDot", "", 0, 0, 6, 6, 8, False, 0, C_GOLD
    modUiKit.NewLbl z, "fzbStat", Poruka("OTKUI_SPLASH_POKRECEM"), 0, 0, 158, _
                    modUiKit.TxtH(TS_META), TS_META, False, MutedOnForest(), -1, fmTextAlignRight
End Sub

' Kartica prijave: ivica + krem povrsina + zlatna nit, znak, dva shell polja,
' greska i dva dugmeta. Mere su iste kao u frmLogin -- ovo je ista kartica,
' samo bez svoje forme.
Private Sub GradiPrijavu(z As Object)
    Dim fnt As String, iw As Single
    fnt = modUiKit.DisplayFont()
    iw = CARD_W - 2 * PAD

    modUiKit.NewLbl z, "fzlCardB", "", 0, 0, CARD_W, CARD_H, 8, False, 0, C_BORDER
    modUiKit.NewLbl z, "fzlCardF", "", 0, 0, CARD_W - 2, CARD_H - 2, 8, False, 0, C_CREAM
    modUiKit.NewLbl z, "fzlTop", "", 0, 0, CARD_W - 2, 3, 8, False, 0, C_GOLD

    modUiKit.NewLbl z, "fzlAX", "AX", 0, 0, 34, modUiKit.TxtH(20), 20, True, C_GOLD, -1, _
                    fmTextAlignLeft, fnt
    modUiKit.NewLbl z, "fzlName", "OtkupApp", 0, 0, 180, modUiKit.TxtH(18), 18, True, _
                    C_FOREST, -1, fmTextAlignLeft, fnt

    modUiKit.NewLbl z, "fzlTitle", Poruka("OTKUI_LOGIN_NASLOV"), 0, 0, iw, _
                    modUiKit.TxtH(TS_DISPLAY), TS_DISPLAY, True, C_FOREST, -1, _
                    fmTextAlignLeft, fnt
    modUiKit.NewLbl z, "fzlSub", Poruka("OTKUI_LOGIN_PODNASLOV"), 0, 0, iw, _
                    modUiKit.TxtH(TS_META), TS_META, False, C_MUTED, -1

    modUiKit.NewLbl z, "fzlCapU", UCase$(Poruka("KOR_LBL_KORISNICKO_IME")), 0, 0, iw, _
                    modUiKit.TxtH(TS_LABEL), TS_LABEL, False, C_MUTED, -1
    modUiKit.NewShell z, "fzlShU", 0, 0, iw, FLD_H, C_INPUT_BORDER, C_WHITE
    modUiKit.NewTxt z, "fzlUser", "", 0, 0, iw - 2 * FLD_PAD, FLD_H - 8, False

    modUiKit.NewLbl z, "fzlCapP", UCase$(Poruka("OTKUI_LOGIN_PIN")), 0, 0, iw, _
                    modUiKit.TxtH(TS_LABEL), TS_LABEL, False, C_MUTED, -1
    modUiKit.NewShell z, "fzlShP", 0, 0, iw, FLD_H, C_INPUT_BORDER, C_WHITE
    modUiKit.NewTxt z, "fzlPin", "", 0, 0, iw - 2 * FLD_PAD, FLD_H - 8, False
    z.Controls("fzlPin").PasswordChar = ChrW(8226)      ' bullet -> maskiran PIN

    modUiKit.NewLbl z, "fzlErr", "", 0, 0, iw, modUiKit.TxtH(TS_META), TS_META, False, C_RUST, -1

    modUiKit.BtnV z, "fzlOK", Poruka("OTKUI_LOGIN_PRIJAVA"), 0, 0, 160, 30, "primary"
    modUiKit.BtnV z, "fzlCancel", Poruka("OTKUI_LOGIN_OTKAZI"), 0, 0, 104, 30, "ghost"

    modUiKit.NewLbl z, "fzlBy", "Powered by AgriX", 0, 0, CARD_W, modUiKit.TxtH(TS_MICRO), _
                    TS_MICRO, False, MutedOnForest(), -1, fmTextAlignCenter
End Sub

Private Sub GradiMini(z As Object)
    Dim fnt As String
    fnt = modUiKit.DisplayFont()
    modUiKit.NewLbl z, "fzmCardB", "", 0, 0, MINI_W, MINI_H, 8, False, 0, C_BORDER
    modUiKit.NewLbl z, "fzmCardF", "", 1, 1, MINI_W - 2, MINI_H - 2, 8, False, 0, C_CREAM
    modUiKit.NewLbl z, "fzmBar", "", 1, 1, 5, MINI_H - 2, 8, False, 0, C_FOREST
    modUiKit.NewLbl z, "fzmAX", "AX", 16, 8, 22, modUiKit.TxtH(TS_H1 + 3), TS_H1 + 3, True, _
                    C_GOLD, -1, fmTextAlignLeft, fnt
    modUiKit.NewLbl z, "fzmName", "OtkupApp", 38, 9, 80, modUiKit.TxtH(TS_H1 + 1), TS_H1 + 1, _
                    True, C_FOREST, -1, fmTextAlignLeft, fnt
    modUiKit.NewLbl z, "fzmSub", Poruka("OTKUI_MINI_EXCEL"), 118, 11, MINI_W - 132, _
                    modUiKit.TxtH(TS_META), TS_META, False, C_MUTED, -1, fmTextAlignRight
    modUiKit.BtnV z, "fzmBack", Poruka("OTKUI_MINI_NAZAD"), 16, 38, MINI_W - 32, 28, "primary"
End Sub

'---------------------------------------------------------- raspored ----

' Jedno mesto koje zna kako izgleda svaka faza: velicina prozora, sta se vidi i
' gde stoji. Zove se i pri prikazu i pri promeni velicine.
Private Sub Postavi()
    Dim z As Object, f As Object
    On Error GoTo EH
    Set f = mForma
    If f Is Nothing Then Exit Sub
    Set z = Zona()
    If z Is Nothing Then Exit Sub

    Select Case Faza()
        Case FAZA_APP
            z.Visible = False
            PunProzor f
            modOtkupUI.OtkupUI_ZoneUstupi False
        Case FAZA_MINI
            modOtkupUI.OtkupUI_ZoneUstupi True
            MiniProzor f
            z.Visible = True
            z.ZOrder 0
            PostaviMini z
        Case Else
            modOtkupUI.OtkupUI_ZoneUstupi True
            z.Visible = True
            z.ZOrder 0
            PostaviPunEkran z, f
    End Select
    Exit Sub
EH:
    LogErr "modUiFaze.Postavi"
End Sub

Private Sub PostaviPunEkran(z As Object, f As Object)
    Dim w As Single, h As Single, bh As Single, i As Long
    Dim cx As Single, cy As Single, Y As Single, znakW As Single

    w = f.InsideWidth: h = f.InsideHeight
    If w < 400 Then w = 400
    If h < 300 Then h = 300
    z.Left = 0: z.top = 0: z.width = w: z.Height = h
    z.BackColor = C_FOREST

    ' podloga: gradijent preko cele zone + zlatna nit
    bh = h / BANDS
    For i = 0 To BANDS - 1
        Vidi z, "fzpGr" & i, True
        Mesto z, "fzpGr" & i, 0, i * bh, w, bh + 1
    Next i
    Vidi z, "fzpLn", True
    Mesto z, "fzpLn", 0, 0, w, 3

    If Faza() = FAZA_BOOT Then
        PrikaziGrupu z, "fzb", True
        PrikaziGrupu z, "fzl", False
        PrikaziGrupu z, "fzm", False

        ' Znak centriran u gornjoj trecini: "AX" desno poravnat do sredine,
        ' "OtkupApp" levo od nje -- par se sam centrira bez merenja teksta.
        Y = h * 0.34 - modUiKit.TxtH(LOGO_FS) / 2
        znakW = 150
        Mesto z, "fzbAX", w / 2 - 60 - znakW / 2, Y, 60, modUiKit.TxtH(LOGO_FS)
        Mesto z, "fzbName", w / 2 - znakW / 2 + 4, Y, 240, modUiKit.TxtH(LOGO_FS)
        Mesto z, "fzbVer", (w - 300) / 2, Y + modUiKit.TxtH(LOGO_FS) + 10, 300, _
              modUiKit.TxtH(TS_META)

        Mesto z, "fzbDiv", PAD, h - FOOT_H, w - 2 * PAD, 1
        Mesto z, "fzbBy", PAD, h - FOOT_H + 16, 200, modUiKit.TxtH(TS_MICRO)
        Mesto z, "fzbDot", w - PAD - 168, h - FOOT_H + 19, 6, 6
        Mesto z, "fzbStat", w - PAD - 158, h - FOOT_H + 15, 158, modUiKit.TxtH(TS_META)
        Exit Sub
    End If

    ' LOGIN: kartica u sredini, malo iznad optickog centra
    PrikaziGrupu z, "fzb", False
    PrikaziGrupu z, "fzm", False
    PrikaziGrupu z, "fzl", True

    cx = (w - CARD_W) / 2
    cy = (h - CARD_H) / 2 - 20
    PostaviKarticu z, cx, cy
End Sub

Private Sub PostaviKarticu(z As Object, ByVal cx As Single, ByVal cy As Single)
    Dim iw As Single
    iw = CARD_W - 2 * PAD

    Mesto z, "fzlCardB", cx, cy, CARD_W, CARD_H
    Mesto z, "fzlCardF", cx + 1, cy + 1, CARD_W - 2, CARD_H - 2
    Mesto z, "fzlTop", cx + 1, cy + 1, CARD_W - 2, 3

    Mesto z, "fzlAX", cx + PAD, cy + 26, 34, modUiKit.TxtH(20)
    Mesto z, "fzlName", cx + PAD + 32, cy + 28, 180, modUiKit.TxtH(18)
    Mesto z, "fzlTitle", cx + PAD, cy + 66, iw, modUiKit.TxtH(TS_DISPLAY)
    Mesto z, "fzlSub", cx + PAD, cy + 92, iw, modUiKit.TxtH(TS_META)

    Mesto z, "fzlCapU", cx + PAD, cy + 116, iw, modUiKit.TxtH(TS_LABEL)
    modUiKit.MoveShell z, "fzlShU", cx + PAD, cy + 130, iw
    Mesto z, "fzlUser", cx + PAD + FLD_PAD, cy + 134, iw - 2 * FLD_PAD, FLD_H - 8

    Mesto z, "fzlCapP", cx + PAD, cy + 168, iw, modUiKit.TxtH(TS_LABEL)
    modUiKit.MoveShell z, "fzlShP", cx + PAD, cy + 182, iw
    Mesto z, "fzlPin", cx + PAD + FLD_PAD, cy + 186, iw - 2 * FLD_PAD, FLD_H - 8

    Mesto z, "fzlErr", cx + PAD, cy + 218, iw, modUiKit.TxtH(TS_META)

    modUiKit.MoveBtn z, "fzlOK", cx + PAD, cy + 240
    modUiKit.MoveBtn z, "fzlCancel", cx + CARD_W - PAD - 104, cy + 240
    Mesto z, "fzlBy", cx, cy + CARD_H + 14, CARD_W, modUiKit.TxtH(TS_MICRO)

    ' Kartica mora da bude IZNAD gradijenta: trake su napravljene prve, pa bi
    ' inace pokrile sve na sebi.
    PodigniGrupu z, "fzl"
End Sub

Private Sub PostaviMini(z As Object)
    PrikaziGrupu z, "fzp", False
    PrikaziGrupu z, "fzb", False
    PrikaziGrupu z, "fzl", False
    PrikaziGrupu z, "fzm", True
    z.Left = 0: z.top = 0
    z.width = MINI_W: z.Height = MINI_H
    z.BackColor = C_CREAM
    PodigniGrupu z, "fzm"
End Sub

' Prozor se skuplja na karticu i seda gore desno u Excelu -- isto mesto koje je
' frmExcelMini birao. Naslovna traka je vec skinuta (GoFullScreen).
'
' Mera se postavlja DVA PUTA: Width je spoljasnja, a kartica je crtana po
' unutrasnjoj (InsideWidth). Razlika je debljina okvira, koju WS_THICKFRAME
' menja -- pa fiksan dodatak ("+6") odsece kartici desnu ivicu na delu masina.
Private Sub MiniProzor(f As Object)
    On Error Resume Next
    f.StartUpPosition = 0
    f.width = MINI_W
    f.Height = MINI_H
    f.width = MINI_W + (f.width - f.InsideWidth)
    f.Height = MINI_H + (f.Height - f.InsideHeight)
    If Application.Visible Then
        f.Left = Application.Left + Application.width - f.width - 20
        f.top = Application.top + 40
    End If
End Sub

' Povratak iz mini kartice u ljusku. Kad prozor vec ima punu meru, ne dira se
' nista -- inace bi svaki povratak na ekran ponistio velicinu koju je operater
' rucno namestio (forma se i dalje hvata za ivicu).
Private Sub PunProzor(f As Object)
    On Error Resume Next
    If f Is Nothing Then Exit Sub
    If f.width > MINI_W + 40 Then Exit Sub
    modOtkupUI.GoFullScreen f
End Sub

'--------------------------------------------------------- sitni alat ----

Private Sub Mesto(z As Object, ByVal nm As String, ByVal X As Single, ByVal Y As Single, _
                  ByVal w As Single, ByVal h As Single)
    On Error Resume Next
    With z.Controls(nm)
        .Left = X: .top = Y: .width = w: .Height = h
        .Visible = True
    End With
    Err.Clear
End Sub

Private Sub Vidi(z As Object, ByVal nm As String, ByVal vis As Boolean)
    On Error Resume Next
    z.Controls(nm).Visible = vis
    Err.Clear
End Sub

' Kontrole faze nose prefiks od tri slova ("fzp", "fzb", "fzl", "fzm"), pa se
' cela grupa pali i gasi bez spiska imena koji bi zastareo pri prvoj doradi.
Private Sub PrikaziGrupu(z As Object, ByVal pref As String, ByVal vis As Boolean)
    Dim c As Object
    On Error Resume Next
    For Each c In z.Controls
        If Left$(c.name, 3) = pref Then c.Visible = vis
    Next c
    Err.Clear
End Sub

Private Sub PodigniGrupu(z As Object, ByVal pref As String)
    Dim c As Object
    On Error Resume Next
    For Each c In z.Controls
        If Left$(c.name, 3) = pref Then c.ZOrder 0
    Next c
    Err.Clear
End Sub

Private Sub Fokus(ByVal nm As String)
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    z.Controls(nm).SetFocus
    Err.Clear
End Sub

' Focus/error stanje shell-a oko polja prijave.
Private Sub Shell_(ByVal tag As String, ByVal state As String)
    Dim z As Object
    On Error Resume Next
    Set z = Zona()
    If z Is Nothing Then Exit Sub
    Select Case tag
        Case "fzlUser": modUiKit.ShellState z, "fzlShU", state
        Case "fzlPin":  modUiKit.ShellState z, "fzlShP", state
    End Select
    Err.Clear
End Sub

'-------------------------------------------------------------- test ----

' Seam za suite: izgradi fazu nad DATOM formom, bez .Show (u harnessu se forma
' ne prikazuje). Bez ovoga se kartica prijave ne moze izmeriti, jer je jedini
' proizvodni ulaz u nju petlja koja ceka klik operatera.
Public Sub FazaGradiTest(frm As Object, ByVal faza As String)
    Set mForma = frm
    mZiva = True
    mFaza = faza
    mPokusaji = 0
    mPrijavaOK = False
    mPrijavaCeka = (faza = FAZA_LOGIN)
    Gradi
    Postavi
End Sub

' Faza se mora moci postaviti PRE nego sto forma nastane: bas to meri tvrdnja
' da prijava ne gradi ljusku (UserForm_Initialize je jedini trenutak u kome
' se ta odluka donosi).
Public Sub FazaRezimTest(ByVal faza As String)
    mFaza = faza
End Sub

Public Function FazaPokusajiTest() As Long
    FazaPokusajiTest = mPokusaji
End Function

Public Function FazaCekaTest() As Boolean
    FazaCekaTest = mPrijavaCeka
End Function

Public Function FazaIshodTest() As Boolean
    FazaIshodTest = mPrijavaOK
End Function
