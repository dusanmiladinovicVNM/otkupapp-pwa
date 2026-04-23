VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "UserForm1"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16470
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmMain – Hauptmenü
' ============================================================

Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME & " v" & APP_VERSION
End Sub
Private Sub UserForm_Activate()
    Dim warnText As String
    warnText = CheckVerwaisteDokumente()
    
    If warnText <> "" Then
        lblWarning.Visible = True
        lblWarning.Caption = warnText
        lblWarning.foreColor = RGB(255, 0, 0)
        lblWarning.Font.Bold = True
    Else
        lblWarning.Visible = False
    End If
End Sub

' --- Otkup (Kooperant zu Station) ---
Private Sub btnUnosPodataka_Click()
    Me.Hide
    frmOtkup.Show
End Sub

' --- Dokumenta (Otpremnica/Zbirna/Prijemnica) ---
Private Sub btnDokumenta_Click()
    Me.Hide
    frmDokumenta.Show
End Sub

Private Sub btnAgrohemija_Click()
    Me.Hide
    frmAgrohemija.Show
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
    TestPdfTextParser
End Sub

' --- Stammdaten ---
Private Sub btnKooperanti_Click()
    Me.Hide
    Dim frm As New frmStammdaten
    frm.Tag = "Kooperanti"
    frm.Show
End Sub

Private Sub btnStanice_Click()
    Me.Hide
    Dim frm As New frmStammdaten
    frm.Tag = "Stanice"
    frm.Show
End Sub

Private Sub btnKupci_Click()
    Me.Hide
    Dim frm As New frmStammdaten
    frm.Tag = "Kupci"
    frm.Show
End Sub

Private Sub btnVozaci_Click()
    Me.Hide
    Dim frm As New frmStammdaten
    frm.Tag = "Vozaci"
    frm.Show
End Sub

Private Sub btnMagacin_Click()
    Me.Hide
    Dim frm As New frmStammdaten
    frm.Tag = "Magacin"
    frm.Show
End Sub

Private Sub btnArtikli_Click()
    Me.Hide
    Dim frm As New frmStammdaten
    frm.Tag = "Artikli"
    frm.Show
End Sub

Private Sub btnParcele_Click()
    Me.Hide
    Dim frm As New frmStammdaten
    frm.Tag = "Parcele"
    frm.Show
End Sub

' --- Berichte ---
Private Sub btnIzvestaji_Click()
    Me.Hide
    frmIzvestaj.Show
End Sub

' --- Fakturierung ---
Private Sub btnFakturisanje_Click()
    Me.Hide
    frmFakturisanje.Show
End Sub

' --- Marža ---
Private Sub btnMarza_Click()
    Me.Hide
    frmMarza.Show
End Sub

' --- Otkupni blokovi ---
Private Sub btnOtkupniBlokovi_Click()
    Me.Hide
    frmSledljivost.Show
End Sub

' --- System ---
Private Sub btnOtvoriExcel_Click()
    OpenExcel
End Sub

Private Sub btnZatvoriExcel_Click()
    CloseExcel
End Sub

Private Sub btnSnimi_Click()
    SaveApp
    MsgBox "Sacuvano!", vbInformation, APP_NAME
End Sub

Private Sub btnIzadji_Click()
    ShutdownApp
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        ShutdownApp
    End If
End Sub
