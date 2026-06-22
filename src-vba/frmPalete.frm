VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPalete 
   Caption         =   "UserForm1"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   OleObjectBlob   =   "frmPalete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPalete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mChromeRemoved As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo EH

    Me.cmbFilterStatus.Clear
    Me.cmbFilterStatus.AddItem ""            ' Sve
    Me.cmbFilterStatus.AddItem "Otvorena"
    Me.cmbFilterStatus.AddItem "Zatvorena"

    Me.cmbFilterPre.Clear
    Me.cmbFilterPre.AddItem ""               ' Sve
    Me.cmbFilterPre.AddItem "Ne"
    Me.cmbFilterPre.AddItem "Da"

    Me.txtFilterGod.value = Year(Date)
    
    FillCmb Me.cmbFilterVrsta, GetLookupList(TBL_KULTURE, "VrstaVoca")
    Me.cmbFilterVrsta.AddItem "", 0

    Me.lstPalete.ColumnCount = 13
    Me.lstPalete.ColumnWidths = "0;30;32;50;50;30;40;40;48;48;50;50;50"
    
    Me.lstPaleteHdr.ColumnCount = 13
    Me.lstPaleteHdr.ColumnWidths = Me.lstPalete.ColumnWidths
    Dim hdr(0 To 0, 0 To 12) As Variant
    hdr(0, 1) = "Broj":   hdr(0, 2) = "God":     hdr(0, 3) = "Vrsta"
    hdr(0, 4) = "Sorta":  hdr(0, 5) = "Klasa":   hdr(0, 6) = "TipAmb"
    hdr(0, 7) = "Gajb":   hdr(0, 8) = "Kap":     hdr(0, 9) = "Neto"
    hdr(0, 10) = "Bruto": hdr(0, 11) = "Status": hdr(0, 12) = "Prer."
    Me.lstPaleteHdr.List = hdr
    Me.lstPaleteHdr.locked = True
    
    Me.lstStavke.ColumnCount = 5
    Me.lstStavke.ColumnWidths = "60;38;54;38;42"
    
    Me.lstStavkeHdr.ColumnCount = 5
    Me.lstStavkeHdr.ColumnWidths = Me.lstStavke.ColumnWidths
    Dim hdrS(0 To 0, 0 To 4) As Variant
    hdrS(0, 0) = "PrijemID": hdrS(0, 1) = "BrPrij": hdrS(0, 2) = "Zbirna"
    hdrS(0, 3) = "Gajb":     hdrS(0, 4) = "Neto"
    Me.lstStavkeHdr.List = hdrS
    Me.lstStavkeHdr.locked = True
    


    RefreshGrid
    Exit Sub
EH:
    MsgBox "Greska pri otvaranju: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next

    EnsureUserFormChromeRemoved Me, mChromeRemoved
    ApplyThemeToControls Me

    StyleFrameTitleLabel lblKopf, "Palete"
    StyleSubtitle lblSubtitle, "Pregled paleta, štampa i prerada"

    Me.lstPaleteHdr.Font.Bold = True
    Me.lstPaleteHdr.BackColor = BG_TOP()
    Me.lstStavkeHdr.Font.Bold = True
    Me.lstStavkeHdr.BackColor = BG_TOP()

    StylePrimaryButton btnPreradi, "Preradi izabrane"
    StyleStornoButton btnStorniraj, "Storniraj"
End Sub

Private Sub RefreshGrid()
    On Error GoTo EH
    Dim god As Long
    If IsNumeric(Me.txtFilterGod.value) Then god = CLng(Me.txtFilterGod.value)

    Dim data As Variant
    data = GetPaleteForGrid(god, Trim$(Me.cmbFilterVrsta.value), _
                            Trim$(Me.cmbFilterStatus.value), Trim$(Me.cmbFilterPre.value))

    Me.lstStavke.Clear
    If IsEmpty(data) Then
        Me.lstPalete.Clear
    Else
        Me.lstPalete.List = data
    End If
    Exit Sub
EH:
    MsgBox "Greska pri osvezavanju: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnOsvezi_Click()
    RefreshGrid
End Sub

' MultiSelect ListBox NE okida Click pouzdano -> Change + ListIndex (red sa
' fokusom = poslednji kliknut) za prikaz stavki desno.
Private Sub lstPalete_Change()
    Dim ids As Collection: Set ids = SelectedPaletaIDs()
    If ids.count = 0 Then
        Dim i As Long: i = Me.lstPalete.ListIndex
        If i >= 0 Then ids.Add CStr(Me.lstPalete.List(i, 0))
    End If
    Dim s As Variant: s = GetPaletaStavkeForGridMulti(ids)
    If IsEmpty(s) Then
        Me.lstStavke.Clear
    Else
        Me.lstStavke.List = s
    End If
End Sub

' Red sa fokusom (akcije nad jednom paletom).
Private Function CurrentPaletaID() As String
    Dim i As Long: i = Me.lstPalete.ListIndex
    If i >= 0 Then CurrentPaletaID = CStr(Me.lstPalete.List(i, 0))
End Function

Private Function SelectedPaletaIDs() As Collection
    Dim c As Collection: Set c = New Collection
    Dim i As Long
    For i = 0 To Me.lstPalete.ListCount - 1
        If Me.lstPalete.Selected(i) Then c.Add CStr(Me.lstPalete.List(i, 0))
    Next i
    Set SelectedPaletaIDs = c
End Function

Private Sub btnStampaj_Click()
    Dim pid As String: pid = CurrentPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    PrintPaletniList pid
End Sub

Private Sub btnPDF_Click()
    Dim pid As String: pid = CurrentPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    ExportPaletniListPDF pid, True
End Sub

Private Sub btnStampajNepotpune_Click()
    On Error GoTo EH
    Dim n As Long: n = PrintNepotpunePalete()
    MsgBox n & " nepotpunih paleta poslato na izlaz (po PALETA_PRINT_MODE).", _
           vbInformation, APP_NAME
    Exit Sub
EH:
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnZatvori_Click()
    On Error GoTo EH
    Dim pid As String: pid = CurrentPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    ClosePaletaManual_TX pid
    RefreshGrid
    MsgBox "Paleta je zatvorena.", vbInformation, APP_NAME
    Exit Sub
EH:
    MsgBox "Paleta nije zatvorena: " & Err.description, vbExclamation, APP_NAME
End Sub

Private Sub btnPreradi_Click()
    On Error GoTo EH
    Dim ids As Collection: Set ids = SelectedPaletaIDs()
    If ids.count = 0 Then
        MsgBox "Izaberite bar jednu paletu (Ctrl/Shift za vise).", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim preID As String
    preID = SavePrerada_TX(ids, _
                CLng(val(Me.txtKutije.value)), _
                CLng(val(Me.txtKese.value)), _
                CDbl(val(Replace(Me.txtNeto.value, ",", "."))), _
                Trim$(Me.txtNapomena.value))

    If preID <> "" Then ExportPreradaPDF preID, True

    Me.txtKutije.value = ""
    Me.txtKese.value = ""
    Me.txtNeto.value = ""
    Me.txtNapomena.value = ""
    RefreshGrid
    MsgBox "Prerada je sacuvana.", vbInformation, APP_NAME
    Exit Sub
EH:
    MsgBox "Prerada nije sacuvana: " & Err.description, vbExclamation, APP_NAME
End Sub

Private Sub btnStorniraj_Click()
    On Error GoTo EH
    Dim pid As String: pid = CurrentPaletaID()
    If pid = "" Then
        MsgBox "Izaberite paletu.", vbInformation, APP_NAME
        Exit Sub
    End If
    If MsgBox("Stornirati izabranu paletu?", vbYesNo + vbQuestion, APP_NAME) <> vbYes Then Exit Sub

    If StornoPaleta_TX(pid) Then
        RefreshGrid
        MsgBox "Paleta je stornirana.", vbInformation, APP_NAME
    Else
        MsgBox "Storno nije uspeo (vidi log).", vbExclamation, APP_NAME
    End If
    Exit Sub
EH:
    MsgBox "Greska pri stornu: " & Err.description, vbExclamation, APP_NAME
End Sub

Private Sub btnPovratak_Click()
    Unload Me
End Sub

