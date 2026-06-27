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

' --- Prerada panel: dinamicke kontrole (Controls.Add; .frx se ne dira) ---
Private WithEvents mTxtTezinaPalete As MSForms.TextBox
Private WithEvents mTxtBruto As MSForms.TextBox
Private WithEvents mDdTipKutije As MSForms.ComboBox
Private WithEvents mCmbTipKese As MSForms.ComboBox
Private mCmbFilterSorta As MSForms.ComboBox
Private mCmbFilterTipGP As MSForms.ComboBox
Private mLblTezinaPalete As MSForms.label
Private mLblBruto As MSForms.label
Private mLblFilterSorta As MSForms.label
Private mLblFilterTipGP As MSForms.label
Private mBuilt As Boolean

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

    BuildPreradaControls
    LayoutDynamic
End Sub

Private Sub RefreshGrid()
    On Error GoTo EH
    Dim god As Long
    If IsNumeric(Me.txtFilterGod.value) Then god = CLng(Me.txtFilterGod.value)

    Dim data As Variant
    data = GetPaleteForGrid(god, Trim$(Me.cmbFilterVrsta.value), _
                            Trim$(Me.cmbFilterStatus.value), Trim$(Me.cmbFilterPre.value), _
                            SortaFilterVal())

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

    Dim brKut As Long: brKut = CLng(ToNum(Me.txtKutije.value))
    Dim brKes As Long: brKes = CLng(ToNum(Me.txtKese.value))
    Dim tezPal As Double, bruto As Double
    Dim tipKut As String, tipKes As String, tipGP As String
    If Not mTxtTezinaPalete Is Nothing Then tezPal = ToNum(mTxtTezinaPalete.value)
    If Not mTxtBruto Is Nothing Then bruto = ToNum(mTxtBruto.value)
    If Not mDdTipKutije Is Nothing Then tipKut = Trim$(mDdTipKutije.value)
    If Not mCmbTipKese Is Nothing Then tipKes = Trim$(mCmbTipKese.value)
    If Not mCmbFilterTipGP Is Nothing Then tipGP = Trim$(mCmbFilterTipGP.value)

    Dim amb As Double: amb = brKut * GetTezinaKutije(tipKut) + brKes * GetTezinaKese(tipKes)
    Dim neto As Double: neto = bruto - tezPal - amb
    If neto < 0 Then neto = 0

    Dim preID As String
    preID = SavePrerada_TX(ids, brKut, brKes, neto, Trim$(Me.txtNapomena.value), _
                tezPal, bruto, amb, tipKut, tipKes, tipGP)

    If preID <> "" Then ExportPreradaPDF preID, True

    ClearPreradaInputs
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

' ============================================================
' Prerada panel - dinamicke kontrole + preracun neta + raspored.
' .frx se ne dira (isti pristup kao modOtkupBlok/clsBlokUI).
' ============================================================
Private Sub BuildPreradaControls()
    On Error GoTo EH
    If mBuilt Then Exit Sub

    Set mTxtTezinaPalete = Me.Controls.Add("Forms.TextBox.1", "txtTezinaPalete", True)
    Set mTxtBruto = Me.Controls.Add("Forms.TextBox.1", "txtBruto", True)
    Set mDdTipKutije = Me.Controls.Add("Forms.ComboBox.1", "ddTipKutije", True)
    Set mCmbTipKese = Me.Controls.Add("Forms.ComboBox.1", "cmbTipKese", True)
    Set mCmbFilterSorta = Me.Controls.Add("Forms.ComboBox.1", "cmbFilterSorta", True)
    Set mCmbFilterTipGP = Me.Controls.Add("Forms.ComboBox.1", "cmbFilterTipGP", True)
    Set mLblTezinaPalete = Me.Controls.Add("Forms.Label.1", "lblTezPal", True)
    Set mLblBruto = Me.Controls.Add("Forms.Label.1", "lblBrutoPal", True)
    Set mLblFilterSorta = Me.Controls.Add("Forms.Label.1", "lblFiltSorta", True)
    Set mLblFilterTipGP = Me.Controls.Add("Forms.Label.1", "lblFiltTipGP", True)

    mLblTezinaPalete.caption = "Tez. palete:"
    mLblBruto.caption = "Bruto:"
    mLblFilterSorta.caption = "Sorta:"
    mLblFilterTipGP.caption = "Gotov proizvod:"

    mDdTipKutije.style = fmStyleDropDownList
    mCmbTipKese.style = fmStyleDropDownCombo
    mCmbFilterSorta.style = fmStyleDropDownCombo
    mCmbFilterTipGP.style = fmStyleDropDownCombo

    ' Reuse postojecih polja: kutije/kese = broj; neto = izracunat (locked).
    On Error Resume Next
    Me.lblKutije.caption = "Kutije:"
    Me.lblKese.caption = "Kese:"
    Me.lblNeto.caption = "Neto (izr.):"
    Me.txtNeto.locked = True
    Me.txtNeto.TabStop = False
    On Error GoTo EH

    FillCmb mDdTipKutije, GetKutijeOptions()
    FillCmb mCmbTipKese, GetKeseOptions()
    FillCmb mCmbFilterTipGP, GetVrstaGPOptions()
    mCmbFilterTipGP.AddItem "", 0
    RefreshSortaFilter
    SetPreradaTabOrder

    mBuilt = True
    Exit Sub
EH:
    LogErr "frmPalete.BuildPreradaControls"
End Sub

' Tab redosled u panelu prerade. Dinamicke kontrole se po defaultu dodaju na
' kraj tab-order-a; ovde ih ubacujemo u vizuelni niz (odozgo nadole). Neto je
' zakljucan (TabStop=False) pa ispada iz tabovanja, ali ostaje u nizu.
Private Sub SetPreradaTabOrder()
    On Error Resume Next
    Dim b As Integer: b = Me.txtKutije.TabIndex
    mTxtTezinaPalete.TabIndex = b
    mDdTipKutije.TabIndex = b + 1
    Me.txtKutije.TabIndex = b + 2
    mCmbTipKese.TabIndex = b + 3
    Me.txtKese.TabIndex = b + 4
    mTxtBruto.TabIndex = b + 5
    Me.txtNeto.TabIndex = b + 6
    Me.txtNapomena.TabIndex = b + 7
End Sub

' Idempotentan raspored (apsolutne pozicije iz stabilnih sidrista) - moze
' se zvati vise puta (Activate / posle teme) bez gomilanja pomeraja.
Private Sub LayoutDynamic()
    On Error Resume Next
    If Not mBuilt Then Exit Sub

    ' Filter red 2: Sorta + Tip gotovog proizvoda (ispod reda 1).
    ' Sekvencijalni raspored sa sopstvenim razmacima -> duza labela
    ' "Gotov proizvod" se ne preklapa sa combo-om (red 1 koordinate ne valjaju).
    Dim f1 As Double: f1 = Me.txtFilterGod.Top
    Dim f2 As Double: f2 = f1 + 26
    Dim rl As Double: rl = Me.lblFilterGod.Left
    Dim sCmbW As Double: sCmbW = Me.cmbFilterVrsta.width
    Dim sCmbX As Double: sCmbX = rl + 40
    PutLbl mLblFilterSorta, rl, f2, 36
    PutCtl mCmbFilterSorta, sCmbX, f2, sCmbW
    Dim tLblX As Double: tLblX = sCmbX + sCmbW + 14
    PutLbl mLblFilterTipGP, tLblX, f2, 88
    PutCtl mCmbFilterTipGP, tLblX + 90, f2, 150

    ' Pomeri grid zonu ispod reda 2 (apsolutno -> idempotentno).
    Dim gTop As Double: gTop = f2 + 26
    ShiftTop Me.lblPalete, gTop
    ShiftTop Me.lstPaleteHdr, gTop + 13
    ShiftTop Me.lstPalete, gTop + 27
    ShiftTop Me.lblStavke, gTop
    ShiftTop Me.lstStavkeHdr, gTop + 13
    ShiftTop Me.lstStavke, gTop + 27

    ' Desni panel inputa ispod stavki (anchored: lstStavke + dugmad).
    Dim px As Double: px = Me.lstStavke.Left
    Dim pw As Double: pw = Me.lstStavke.width
    Dim btnTop As Double: btnTop = Me.btnPreradi.Top
    Const ROWH As Double = 24
    Dim pTop As Double: pTop = btnTop - 12 - 6 * ROWH

    Me.lstPalete.Height = btnTop - 12 - Me.lstPalete.Top
    Me.lstStavke.Height = pTop - 8 - Me.lstStavke.Top
    If Me.lstStavke.Height < 48 Then Me.lstStavke.Height = 48

    LayoutPreradaRows px, pw, pTop, ROWH
End Sub

Private Sub LayoutPreradaRows(ByVal x As Double, ByVal w As Double, _
                              ByVal y0 As Double, ByVal rowH As Double)
    On Error Resume Next
    Dim labelW As Double: labelW = 76
    Dim inX As Double: inX = x + labelW + 4
    Dim brojW As Double: brojW = 42
    Dim inW As Double: inW = w - labelW - 4
    Dim comboW As Double: comboW = inW - brojW - 6
    If comboW < 50 Then comboW = 50
    Dim brojX As Double: brojX = x + w - brojW

    ' 0 Tezina palete
    PutLbl mLblTezinaPalete, x, y0, labelW
    PutCtl mTxtTezinaPalete, inX, y0, inW

    ' 1 Kutije: tip (dd) + broj
    PutLbl Me.lblKutije, x, y0 + rowH, labelW
    PutCtl mDdTipKutije, inX, y0 + rowH, comboW
    PutCtl Me.txtKutije, brojX, y0 + rowH, brojW

    ' 2 Kese: tip (cmb) + broj
    PutLbl Me.lblKese, x, y0 + 2 * rowH, labelW
    PutCtl mCmbTipKese, inX, y0 + 2 * rowH, comboW
    PutCtl Me.txtKese, brojX, y0 + 2 * rowH, brojW

    ' 3 Bruto
    PutLbl mLblBruto, x, y0 + 3 * rowH, labelW
    PutCtl mTxtBruto, inX, y0 + 3 * rowH, inW

    ' 4 Neto (izracunato)
    PutLbl Me.lblNeto, x, y0 + 4 * rowH, labelW
    PutCtl Me.txtNeto, inX, y0 + 4 * rowH, inW

    ' 5 Napomena
    PutLbl Me.lblNapomena, x, y0 + 5 * rowH, labelW
    PutCtl Me.txtNapomena, inX, y0 + 5 * rowH, inW
End Sub

Private Sub PutLbl(ByVal ctl As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double)
    On Error Resume Next
    ctl.Left = x: ctl.Top = y + 2: ctl.width = w: ctl.Height = 14: ctl.Visible = True
End Sub

Private Sub PutCtl(ByVal ctl As Object, ByVal x As Double, ByVal y As Double, ByVal w As Double)
    On Error Resume Next
    ctl.Left = x: ctl.Top = y: ctl.width = w: ctl.Height = 18: ctl.Visible = True
End Sub

Private Sub ShiftTop(ByVal ctl As Object, ByVal newTop As Double)
    On Error Resume Next
    ctl.Top = newTop
End Sub

' Sorta filter zavisi od izabrane vrste (kaskada kao na Cenovniku).
Private Sub RefreshSortaFilter()
    On Error Resume Next
    If mCmbFilterSorta Is Nothing Then Exit Sub
    Dim cur As String: cur = Trim$(mCmbFilterSorta.value)
    mCmbFilterSorta.Clear
    mCmbFilterSorta.AddItem ""
    Dim vr As String: vr = Trim$(Me.cmbFilterVrsta.value)
    Dim arr As Variant
    If vr = "" Then
        arr = GetLookupList(TBL_KULTURE, "SortaVoca")
    Else
        arr = GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", vr)
    End If
    If IsArray(arr) Then
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            mCmbFilterSorta.AddItem CStr(arr(i))
        Next i
    End If
    If cur <> "" Then mCmbFilterSorta.value = cur
End Sub

Private Function SortaFilterVal() As String
    On Error Resume Next
    If Not mCmbFilterSorta Is Nothing Then SortaFilterVal = Trim$(mCmbFilterSorta.value)
End Function

' Neto = Bruto - tezina palete - (broj_kutija*tez_tipa + broj_kesa*tez_tipa).
Private Sub RecomputeNeto()
    On Error Resume Next
    If Not mBuilt Then Exit Sub
    Dim bruto As Double: bruto = ToNum(mTxtBruto.value)
    Dim tezPal As Double: tezPal = ToNum(mTxtTezinaPalete.value)
    Dim brKut As Long: brKut = CLng(ToNum(Me.txtKutije.value))
    Dim brKes As Long: brKes = CLng(ToNum(Me.txtKese.value))
    Dim amb As Double
    amb = brKut * GetTezinaKutije(Trim$(mDdTipKutije.value)) + _
          brKes * GetTezinaKese(Trim$(mCmbTipKese.value))
    Dim neto As Double: neto = bruto - tezPal - amb
    If neto < 0 Then neto = 0
    Me.txtNeto.value = Format$(neto, "0.00")
End Sub

' Lokalno-nezavisno citanje broja (zapeta -> tacka; Val koristi tacku).
Private Function ToNum(ByVal s As Variant) As Double
    ToNum = val(Replace(Trim$(CStr(s)), ",", "."))
End Function

Private Sub ClearPreradaInputs()
    On Error Resume Next
    Me.txtKutije.value = ""
    Me.txtKese.value = ""
    Me.txtNeto.value = ""
    Me.txtNapomena.value = ""
    mTxtTezinaPalete.value = ""
    mTxtBruto.value = ""
    mDdTipKutije.value = ""
    mCmbTipKese.value = ""
    ' tipGP (izlaz) ostaje izabran za sledecu preradu
End Sub

' --- preracun na promenu bilo kog ulaza koji utice na neto ---
Private Sub mTxtTezinaPalete_Change()
    RecomputeNeto
End Sub

Private Sub mTxtBruto_Change()
    RecomputeNeto
End Sub

Private Sub mDdTipKutije_Change()
    RecomputeNeto
End Sub

Private Sub mCmbTipKese_Change()
    RecomputeNeto
End Sub

Private Sub txtKutije_Change()
    RecomputeNeto
End Sub

Private Sub txtKese_Change()
    RecomputeNeto
End Sub

' Vrsta -> osvezi listu sorti (kaskada); grid se osvezava na Osvezi.
Private Sub cmbFilterVrsta_Change()
    RefreshSortaFilter
End Sub

