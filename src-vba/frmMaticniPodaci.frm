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

Private isOpeningChild As Boolean

Private Sub RemoveTitleBar()
    Dim hwnd As LongPtr

    ' pronadi prozor po klasi ThunderDFrame (VBA UserForm)
    hwnd = FindWindow("ThunderDFrame", Me.Caption)

    If hwnd <> 0 Then
        Dim style As Long
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    isOpeningChild = False
    
    ApplyFormTheme Me, BG_PANEL

    StyleMenuButton btnKooperanti, "Kooperanti"
    StyleMenuButton btnStanice, "Stanice"
    StyleMenuButton btnKupci, "Kupci"
    StyleMenuButton btnVozaci, "Vozaci"
    StyleMenuButton btnArtikli, "Artikli"
    StyleMenuButton btnParcele, "Parcele"
    StyleExitButton btnExit, "Izadji"

End Sub

Private Sub UserForm_Activate()
    Me.Caption = ""
    RemoveTitleBar
End Sub

Private Sub UserForm_Deactivate()
    If Not isOpeningChild Then
        Unload Me
    End If
End Sub



Private Sub btnKooperanti_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
    ButtonHover btnKooperanti
End Sub

Private Sub btnStanice_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
    ButtonHover btnStanice
End Sub

Private Sub btnKupci_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
    ButtonHover btnKupci
End Sub

Private Sub btnVozaci_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
    ButtonHover btnVozaci
End Sub

Private Sub btnArtikli_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
    ButtonHover btnArtikli
End Sub

Private Sub btnParcele_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
    ButtonHover btnParcele
End Sub

Private Sub btnExit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
    ButtonHover btnExit
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetButtonGroupWithExit btnExit, btnKooperanti, btnStanice, btnKupci, btnVozaci, btnArtikli, btnParcele
End Sub

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
    Unload Me
End Sub

Private Sub OpenStammdatenForm(ByVal nazivSekcije As String)
    Dim frm As New frmStammdaten

    isOpeningChild = True

    Me.Hide
    frmOtkupAPP.Hide

    frm.Tag = nazivSekcije
    frm.StartUpPosition = 2
    frm.Show

    Unload Me
End Sub
