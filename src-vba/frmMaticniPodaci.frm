VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMaticniPodaci 
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2265
   OleObjectBlob   =   "frmMaticniPodaci.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMaticniPodaci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ============================================================
' frmStammdatenMenu / popup menu
' Responsibility:
'   - UI navigation only
'   - no business logic
'   - no data writes
' ============================================================

Private m_IsOpeningChild As Boolean
Private m_IsClosing As Boolean

Private Sub RemoveTitleBar()
    Dim hwnd As LongPtr

    ' Pronadi prozor po klasi ThunderDFrame (VBA UserForm)
    hwnd = FindWindow("ThunderDFrame", Me.caption)

    If hwnd <> 0 Then
        Dim style As Long
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo EH

    Me.StartUpPosition = 0

    m_IsOpeningChild = False
    m_IsClosing = False

    ApplyFormTheme Me, BG_PANEL

    ' Staticna dugmad ostaju stilizovana kao fallback; ceo meni se
    ' (sve sekcije, postojece + nove) gradi dinamicki preko modMaticniLookups,
    ' pa se .frx ne dira. Ako dinamicka izgradnja ne uspe, staticna dugmad
    ' ostaju vidljiva i rade po starom.
    StyleMenuButton btnKooperanti, "Kooperanti"
    StyleMenuButton btnStanice, "Stanice"
    StyleMenuButton btnKupci, "Kupci"
    StyleMenuButton btnVozaci, "Vozaci"
    StyleMenuButton btnArtikli, "Artikli"
    StyleMenuButton btnParcele, "Parcele"
    StyleExitButton btnExit, "Izadji"

    AttachMaticniMenu Me

    Exit Sub

EH:
    LogErr "frmStammdatenMenu.UserForm_Initialize"
    MsgBox Poruka("MATICNI_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbCritical, APP_NAME
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH

    ' Vrati na stari working pattern.
    ' Bitno: ne koristiti mChromeRemoved ovde.
    Me.caption = ""
    RemoveTitleBar

    Exit Sub

EH:
    LogErr "frmStammdatenMenu.UserForm_Activate"
End Sub

Private Sub UserForm_Deactivate()
    On Error GoTo EH

    If m_IsOpeningChild Then Exit Sub
    If m_IsClosing Then Exit Sub

    CloseMenuAndReturn
    Exit Sub

EH:
    LogErr "frmStammdatenMenu.UserForm_Deactivate"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        CloseMenuAndReturn
    End If
End Sub

Private Sub CloseMenuAndReturn()
    On Error GoTo EH

    If m_IsClosing Then Exit Sub
    m_IsClosing = True

    On Error Resume Next
    frmOtkupAPP.Show
    On Error GoTo EH

    Unload Me
    Exit Sub

EH:
    LogErr "frmStammdatenMenu.CloseMenuAndReturn"
    Unload Me
End Sub

Private Sub ResetMenuButtons()
    ResetButtonGroupWithExit btnExit, _
                             btnKooperanti, _
                             btnStanice, _
                             btnKupci, _
                             btnVozaci, _
                             btnArtikli, _
                             btnParcele
End Sub

Private Sub HoverMenuButton(ByVal btn As Object)
    On Error GoTo EH

    ResetMenuButtons
    ButtonHover btn

    Exit Sub

EH:
    LogErr "frmStammdatenMenu.HoverMenuButton"
End Sub

' Javna ulazna tacka — otvara frmStammdaten za datu sekciju.
' sekTag      = Tag za frmStammdaten (npr. "TipAmbalaze")
' sekCaption  = naziv u naslovu (npr. "Ambalaza"); fallback na sekTag.
' Poziva je i staticni handler (OpenStammdatenForm) i dinamicki meni
' (modMaticniLookups). Drzi m_IsOpeningChild da Deactivate ne zatvori meni.
Public Sub OpenSekcija(ByVal sekTag As String, Optional ByVal sekCaption As String = "")
    On Error GoTo EH

    If Trim$(sekTag) = "" Then
        Err.Raise vbObjectError + 7601, "frmMaticniPodaci.OpenSekcija", _
                  "Naziv sekcije je obavezan."
    End If

    If Trim$(sekCaption) = "" Then sekCaption = sekTag

    ' Force unload prethodnog frmStammdaten ako jos uvek postoji
    On Error Resume Next
    Unload frmStammdaten
    On Error GoTo EH

    ' Postavi Tag PRE OpenContentFormPublic poziva
    frmStammdaten.Tag = sekTag

    ' Sakri menu pre prebacivanja na frmOtkupAPP shell
    m_IsOpeningChild = True
    Me.Hide

    ' frmOtkupAPP otvara frmStammdaten kao mActiveContent (modeless child)
    On Error Resume Next
    frmOtkupAPP.Show
    frmOtkupAPP.OpenContentFormPublic frmStammdaten, "Maticni podaci: " & sekCaption
    On Error GoTo EH

    Unload Me
    Exit Sub

EH:
    LogErr "frmMaticniPodaci.OpenSekcija"
    m_IsOpeningChild = False
    On Error Resume Next
    Me.Show
    On Error GoTo 0
    MsgBox Poruka("MATICNI_ERR_GRESKA_PRI_OTVARANJU_2") & Err.description, vbCritical, APP_NAME
End Sub

' Staticni handleri (fallback dugmad) idu kroz isti put.
Private Sub OpenStammdatenForm(ByVal nazivSekcije As String)
    OpenSekcija nazivSekcije, nazivSekcije
End Sub

' ============================================================
' HOVER
' ============================================================

Private Sub btnKooperanti_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HoverMenuButton btnKooperanti
End Sub

Private Sub btnStanice_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HoverMenuButton btnStanice
End Sub

Private Sub btnKupci_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HoverMenuButton btnKupci
End Sub

Private Sub btnVozaci_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HoverMenuButton btnVozaci
End Sub

Private Sub btnArtikli_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HoverMenuButton btnArtikli
End Sub

Private Sub btnParcele_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HoverMenuButton btnParcele
End Sub

Private Sub btnExit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HoverMenuButton btnExit
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    ResetMenuButtons
End Sub

' ============================================================
' CLICKS
' ============================================================

Private Sub btnKooperanti_Click()
    ButtonActive btnKooperanti
    OpenStammdatenForm "Kooperanti"
End Sub

Private Sub btnStanice_Click()
    ButtonActive btnStanice
    OpenStammdatenForm "Stanice"
End Sub

Private Sub btnKupci_Click()
    ButtonActive btnKupci
    OpenStammdatenForm "Kupci"
End Sub

Private Sub btnVozaci_Click()
    ButtonActive btnVozaci
    OpenStammdatenForm "Vozaci"
End Sub

Private Sub btnArtikli_Click()
    ButtonActive btnArtikli
    OpenStammdatenForm "Artikli"
End Sub

Private Sub btnParcele_Click()
    ButtonActive btnParcele
    OpenStammdatenForm "Parcele"
End Sub

Private Sub btnExit_Click()
    CloseMenuAndReturn
End Sub

