VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFlatBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' clsFlatBtn - genericki WithEvents omotac za SVE runtime kontrole
' aplikacije. WithEvents ZIVI OVDE, nikad u formi (CLAUDE.md;
' docs/SELF_UPDATE.md zamka #11).
'
'---------------------------------------------------------------------
' KLASA JE ZATVORENA ZA IZMENE (faza S0, v6-ui-64).
'
' Zasto: klasa je "tvrd" modul (module-level WithEvents As MSForms.*),
' pa je self-update NE moze osveziti kroz AddFromString - svaka njena
' izmena ide u fazu 2 (Remove + Import), a to je najrizicniji put koji
' updater ima. Modul je "mek" (faza 1) i menja se slobodno.
'
' Zato ovde NEMA nijednog imena kontrole i nijednog komada UI logike:
' svi eventi idu u modOtkupUI.UiEvent / .UiAsk / .HandleKeyFrom, koji
' odlucuju sta se desava. Nov ekran, dugme, cip, kolona ili panel NE
' traze izmenu ovog fajla.
'
' Ako se ucini da novom ekranu treba dogadjaj kog ovde nema - to je
' skoro uvek znak da se kontrola crta pogresnom vrstom kontrole.
' Dizajn sistem koristi SAMO cetiri: Label, Frame, TextBox, ComboBox.
' Prekidaci, cipovi, segmenti, pilule i celije mreze su Label-i.
'
' Pokriveni dogadjaji (tag ide uz svaki):
'   Label     Click, DblClick, MouseMove -> Hover, desni MouseUp -> Menu
'   Frame     Click, MouseMove, KeyDown, Scroll, desni MouseUp -> Menu
'   TextBox   Change, Enter -> Focus, Exit -> Blur, KeyDown,
'             KeyPress (moze da odbije znak), AfterUpdate -> Commit,
'             desni MouseUp -> Menu
'   ComboBox  isto kao TextBox + DropButtonClick -> Drop
'
' Vrste (kind): btn | chip | seg | hdr | nav | card | chev | row |
'               zone | input
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Private WithEvents L As MSForms.Label
Attribute L.VB_VarHelpID = -1
Private WithEvents FZ As MSForms.Frame
Attribute FZ.VB_VarHelpID = -1
Private WithEvents TX As MSForms.TextBox
Attribute TX.VB_VarHelpID = -1
Private WithEvents CB As MSForms.ComboBox
Attribute CB.VB_VarHelpID = -1

Private mKind As String
Private mTag As String
Private mBg As Long
Private mFg As Long

'---------------------------------------------------------------- bind
Public Sub Bind(ByVal ctl As Object, ByVal kind As String, ByVal tagName As String)
    mKind = kind
    mTag = tagName
    Select Case kind
        Case "zone"
            Set FZ = ctl
        Case "input"
            Select Case TypeName(ctl)
                Case "TextBox":  Set TX = ctl
                Case "ComboBox": Set CB = ctl
            End Select
        Case Else
            Set L = ctl
            mBg = ctl.BackColor
            mFg = ctl.ForeColor
    End Select
End Sub

Public Property Get SinkTag() As String
    SinkTag = mTag
End Property

' Otisak verzije. ImportAllVBA prvo uklanja istoimenu komponentu, a kad to ne
' uspe (zive instance klase / kod se izvrsava) uvozi je pod imenom "clsFlatBtn1"
' i dalje radi STARA klasa - tiho, bez ijedne poruke. OtkupUI_SelfCheck cita
' ovo i odmah kaze da li je klasa zaista zamenjena.
Public Property Get Build() As String
    Build = "v6-ui-86"
End Property

'--------------------------------------------------------------- hover
Private Sub L_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    Select Case mKind
        Case "row"
            ' red mreze se boji centralno (jedan render), ne po celiji
            modOtkupUI.UiEvent mTag, "Hover", Empty
            Exit Sub
        Case "chev"
            Exit Sub                      ' strelica mora da zadrzi boju ispune
        Case "hdr"
            ' zaglavlje kolone pokazuje neutralnu strelicu sortiranja SAMO
            ' dok je pod pokazivacem - modul crta, klasa samo javlja
            modOtkupUI.UiEvent mTag, "Hover", Empty
        Case "card"
            ' Kartica rezima se boji CENTRALNO (RefreshModeCards): aktivna ima
            ' tamnu ispunu i krem naslov. Boje se menjaju POSLE Bind-a, pa su
            ' mBg/mFg zapamceni ovde zastareli - ResetVisual bi posle hover-a
            ' vratio staru (tamnu) boju teksta i naslov aktivne kartice bi
            ' pocrneo. Zato kartica ne ulazi ni u hover ni u reset.
            Exit Sub
    End Select
    If Not IsSelected() Then L.BackColor = HoverColor(L.BackColor)
    modOtkupUI.NoteHover Me
End Sub

Private Sub FZ_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modOtkupUI.ResetAllBtnVisuals
End Sub

' Klik na PRAZNU povrsinu zone. Bez ovoga otvoreni panel (nas dropdown ili
' "Filteri") nije imao od cega da se zatvori: Frame nije slao nijedan klik
' dalje, pa je klik van panela padao u prazno. Panel sam nije zona pa se
' klikom u njega ne zatvara.
Private Sub FZ_Click()
    modOtkupUI.UiEvent mTag, "Click", Empty
End Sub

' Frame MOZE da drzi fokus (Label ne moze). Klik na Label unutar zone ostavlja
' fokus upravo na zoni, a bez ovog handler-a F-tasteri odatle nisu imali kuda -
' zato su "prestajali da rade" posle klika na zaglavlje kolone. modOtkupUI uz to
' vraca fokus u polje (KeepKeys); ovo je druga linija odbrane.
Private Sub FZ_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If ForwardKey(KeyCode, Shift) Then KeyCode = 0
End Sub

' Zona sa sopstvenim klizacem (ScrollBars) - ekran moze da pomeri ono sto je
' iscrtao van vidljivog dela. Danas nijedna zona nema klizac; kanal postoji da
' bi visoki ekrani (izvestaji, podesavanja) mogli da ga dobiju bez diranja
' ove klase.
Private Sub FZ_Scroll(ByVal ActionX As MSForms.fmScrollAction, _
                      ByVal ActionY As MSForms.fmScrollAction, _
                      ByVal RequestDx As Single, ByVal RequestDy As Single, _
                      ByVal ActualDx As MSForms.ReturnSingle, _
                      ByVal ActualDy As MSForms.ReturnSingle)
    modOtkupUI.UiEvent mTag, "Scroll", RequestDy
End Sub

Private Sub L_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If mKind = "row" Or mKind = "chev" Or mKind = "card" Then Exit Sub
    ResetVisual
End Sub

'--------------------------------------------------------------- desni
' Desni taster je jedini ulaz za radnje nad izabranom stavkom koji ne trazi
' dugme na ekranu. Levi taster nikad ne ide ovuda - "Menu" znaci ISKLJUCIVO
' desni klik.
Private Sub L_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then modOtkupUI.UiEvent mTag, "Menu", Empty
End Sub

Private Sub FZ_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then modOtkupUI.UiEvent mTag, "Menu", Empty
End Sub

Private Sub TX_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then modOtkupUI.UiEvent mTag, "Menu", Empty
End Sub

' Selektovano stanje nosi Font.Bold - hover ga ne dira.
Private Function IsSelected() As Boolean
    On Error Resume Next
    Select Case mKind
        Case "nav", "chip", "seg": IsSelected = L.Font.Bold
    End Select
End Function

Private Function HoverColor(ByVal bg As Long) As Long
    Select Case bg
        Case modOtkupUI.C_GREEN:    HoverColor = modOtkupUI.C_BTN_HOVER
        Case modOtkupUI.C_WHITE:    HoverColor = RGB(246, 244, 238)
        Case modOtkupUI.C_GOLD:     HoverColor = RGB(216, 186, 92)
        Case modOtkupUI.C_FOREST:   HoverColor = RGB(32, 48, 24)
        Case modOtkupUI.C_HDR_SURF: HoverColor = RGB(62, 84, 52)
        Case Else:                  HoverColor = RGB(232, 228, 216)
    End Select
End Function

Public Sub ResetVisual()
    On Error Resume Next
    If L Is Nothing Then Exit Sub
    ' "card" je izuzet iz istog razloga kao "row"/"chev" - vidi L_MouseMove
    If mKind = "row" Or mKind = "chev" Or mKind = "card" Then Exit Sub
    If IsSelected() Then Exit Sub
    L.BackColor = mBg
    L.ForeColor = mFg
End Sub

' Boje se pamte pri Bind-u. Ako ih centralni render kasnije promeni, zapamcene
' postaju zastarele i ResetVisual bi vratio staru boju. Render koji menja boju
' vezane kontrole treba da javi novu osnovu (modOtkupUI.RebaseSink).
Public Sub Rebase()
    On Error Resume Next
    If L Is Nothing Then Exit Sub
    mBg = L.BackColor
    mFg = L.ForeColor
End Sub

'--------------------------------------------------------------- click
Private Sub L_Click()
    modOtkupUI.UiEvent mTag, "Click", Empty
End Sub

Private Sub L_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    modOtkupUI.UiEvent mTag, "DblClick", Empty
End Sub

'--------------------------------------------------------------- input
' Enter/Exit nose focus state ivice (ShellState) - to MSForms TextBox
' sam po sebi nema.
Private Sub TX_Change()
    modOtkupUI.UiEvent mTag, "Change", Empty
End Sub

Private Sub TX_Enter()
    modOtkupUI.UiEvent mTag, "Focus", Empty
End Sub

Private Sub TX_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    modOtkupUI.UiEvent mTag, "Blur", Empty
End Sub

Private Sub TX_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If ForwardKey(KeyCode, Shift) Then KeyCode = 0
End Sub

' Jedini dogadjaj koji sme da ODBIJE unos. Modul vraca kod znaka koji prolazi,
' izmenjen kod (npr. tacka -> lokalni decimalni zarez) ili 0 = ne propusti.
Private Sub TX_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = modOtkupUI.UiAsk(mTag, "KeyPress", CLng(KeyAscii))
End Sub

' Za razliku od Change (svaki znak), Commit stize jednom - kad je korisnik
' zavrsio unos i vrednost se stvarno promenila. Mesto za normalizaciju
' (zaokruzivanje, format) koja bi usred kucanja smetala.
Private Sub TX_AfterUpdate()
    modOtkupUI.UiEvent mTag, "Commit", Empty
End Sub

Private Sub CB_Change()
    modOtkupUI.UiEvent mTag, "Change", Empty
End Sub

Private Sub CB_Enter()
    modOtkupUI.UiEvent mTag, "Focus", Empty
End Sub

Private Sub CB_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    modOtkupUI.UiEvent mTag, "Blur", Empty
End Sub

Private Sub CB_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If ForwardKey(KeyCode, Shift) Then KeyCode = 0
End Sub

Private Sub CB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = modOtkupUI.UiAsk(mTag, "KeyPress", CLng(KeyAscii))
End Sub

Private Sub CB_AfterUpdate()
    modOtkupUI.UiEvent mTag, "Commit", Empty
End Sub

' Zahtev za otvaranjem liste koji NE dolazi od nase strelice (npr. Alt+dole).
' Nativno dugme je ugaseno (ShowDropButtonWhen = Never), ali dogadjaj i dalje
' postoji - bez ovoga bi se otvorila MSForms lista sa svojom ivicom i klizacem.
Private Sub CB_DropButtonClick()
    modOtkupUI.UiEvent mTag, "Drop", Empty
End Sub

' MSForms UserForm nema KeyPreview - F-tasteri se hvataju i u svakoj
' kontroli koja moze da drzi fokus, pa se prosledjuju istoj rutini.
' Tag ide uz taster: bez njega modul ne zna IZ KOG polja je taster stigao,
' pa ne moze da razlikuje "F4 = rezim Prijemnice" od radnje nad tim poljem.
Private Function ForwardKey(ByVal KeyCode As Long, ByVal Shift As Integer) As Boolean
    On Error Resume Next
    If (Shift And 2) <> 0 Then          ' Ctrl
        Select Case KeyCode
            Case vbKeyK: modOtkupUI.FocusKontekst: ForwardKey = True: Exit Function
            Case vbKeyF: modOtkupUI.FocusPretraga: ForwardKey = True: Exit Function
        End Select
    End If
    ForwardKey = modOtkupUI.HandleKeyFrom(mTag, KeyCode, Shift)
End Function
