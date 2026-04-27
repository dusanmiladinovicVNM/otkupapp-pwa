VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtkup 
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   OleObjectBlob   =   "frmOtkup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' frmOtkup v2.1 – NUR Otkup (Kooperant ? Station)
' Rechte Seite (Isporuka) wurde entfernt.
' Otpremnica/Zbirna/Prijemnica sind jetzt in frmDokumenta.
' ============================================================
Private mChromeRemoved As Boolean

Private Sub RemoveTitleBar()
    Dim hwnd As LongPtr
    Dim style As Long

    hwnd = FindWindow("ThunderDFrame", Me.caption)

    If hwnd <> 0 Then
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Activate()
    If Not mChromeRemoved Then
        Me.caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If
End Sub

Private Sub UserForm_Initialize()

    ApplyFormTheme Me, BG_MAIN
    ApplyThemeToControls Me

    ' bitne kontrole - eksplicitni stil
    StylePrimaryButton btnUnos, "Unos"
    StyleExitButton btnPovratak, "Povratak"
    StyleStornoButton btnStornoOtkup, "Storno"
    StyleLabel lblUkupnoKG, TXT_MUTED, True

    StyleComboBox cmbVrstaVoca
    StyleComboBox cmbSortaVoca
    StyleComboBox cmbOtkupnoMesto
    StyleComboBox cmbKooperant
    StyleComboBox cmbParcela
    StyleComboBox cmbVozac
    StyleComboBox cmbTipAmbalaze

    StyleTextBox txtDatum
    StyleTextBox txtKolicina
    StyleTextBox txtCena
    StyleTextBox txtKolAmbalaze
    StyleTextBox txtNovac
    StyleTextBox txtPrimalac
    StyleTextBox txtBrojDokumenta
    StyleTextBox txtBrojZbirne
    StyleTextBox txtKolicinaKLII
    StyleTextBox txtCenaKLII
    
    ' Datum defaults
    txtDatum.value = Format$(Date, "d.m.yyyy")
    
    ' ComboBoxen füllen
    FillCmb cmbVrstaVoca, GetLookupList(TBL_KULTURE, "VrstaVoca")
    FillComboDisplayID cmbOtkupnoMesto, TBL_STANICE, "Naziv", "StanicaID"
    FillCmb cmbVozac, GetVozacDisplayList()
    FillCmb cmbTipAmbalaze, Array(AMB_12_1, AMB_6_1)
    
    ' Numerische Felder auf 0 setzen
    txtKolicina.value = ""
    txtCena.value = ""
    txtKolAmbalaze.value = ""
    txtNovac.value = "0"
    
    ' Klasa II – initial disabled
    DisableField txtKolicinaKLII
    DisableField txtCenaKLII
    chkDveKlase.value = False
    lblUkupnoKG.caption = ""
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnUnos, "Unos"
    StyleExitButton btnPovratak, "Povratak"
    StyleStornoButton btnStornoOtkup, "Storno"
End Sub

Private Sub btnUnos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnos
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub btnStornootkup_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnStornoOtkup
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

Private Sub chkDveKlase_Click()
    If chkDveKlase.value Then
        EnableField txtKolicinaKLII
        EnableField txtCenaKLII
        StyleTextBox txtKolicinaKLII
        StyleTextBox txtCenaKLII
    Else
        DisableField txtKolicinaKLII
        DisableField txtCenaKLII
        lblUkupnoKG.caption = ""
    End If
End Sub

Private Sub txtKolicinaKlII_Change()
    UpdateUkupnoKg
End Sub

Private Sub txtKolicina_Change()
    UpdateUkupnoKg
End Sub

Private Sub UpdateUkupnoKg()
    On Error GoTo EH

    If Not chkDveKlase.value Then
        lblUkupnoKG.caption = ""
        Exit Sub
    End If

    Dim kl1 As Double
    Dim kl2 As Double

    If Trim$(txtKolicina.value) <> "" Then
        TryParseDouble txtKolicina.value, kl1
    End If

    If Trim$(txtKolicinaKLII.value) <> "" Then
        TryParseDouble txtKolicinaKLII.value, kl2
    End If

    lblUkupnoKG.caption = "Ukupno: " & Format$(kl1 + kl2, "#,##0.00") & " kg"
    Exit Sub

EH:
    LogErr "frmOtkup.UpdateUkupnoKg"
    lblUkupnoKG.caption = ""
End Sub

' ============================================================
' KASKADIERUNG – VrstaVoca ? SortaVoca
' ============================================================

Private Sub cmbVrstaVoca_Change()
    ' Wenn VrstaVoca gewählt wird, SortaVoca-Liste filtern
    cmbSortaVoca.Clear
    If cmbVrstaVoca.value <> "" Then
        FillCmb cmbSortaVoca, _
            GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", cmbVrstaVoca.value)
    End If
End Sub

Private Sub cmbOtkupnoMesto_Change()
    On Error GoTo EH

    cmbKooperant.Clear
    cmbParcela.Clear

    If cmbOtkupnoMesto.value = "" Then Exit Sub

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)

    If stanicaID = "" Then Exit Sub

    FillComboKooperantiByStanica cmbKooperant, stanicaID
    Exit Sub

EH:
    LogErr "frmOtkup.cmbOtkupnoMesto_Change"
    cmbKooperant.Clear
    cmbParcela.Clear
End Sub

Private Sub cmbKooperant_Change()
    On Error GoTo EH

    cmbParcela.Clear

    If cmbKooperant.ListIndex < 0 Then Exit Sub

    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbKooperant.value)

    If kooperantID = "" Then Exit Sub

    Dim parData As Variant
    parData = GetTableData(TBL_PARCELE)

    If IsEmpty(parData) Then Exit Sub

    Dim colKoop As Long
    Dim colID As Long
    Dim colKat As Long
    Dim colKultura As Long
    Dim colPovrsina As Long

    colID = RequireColumnIndex(TBL_PARCELE, COL_PAR_ID, _
                               "frmOtkup.cmbKooperant_Change")
    colKoop = RequireColumnIndex(TBL_PARCELE, COL_PAR_KOOP, _
                                 "frmOtkup.cmbKooperant_Change")
    colKat = RequireColumnIndex(TBL_PARCELE, COL_PAR_KAT_BROJ, _
                                "frmOtkup.cmbKooperant_Change")
    colKultura = RequireColumnIndex(TBL_PARCELE, COL_PAR_KULTURA, _
                                    "frmOtkup.cmbKooperant_Change")
    colPovrsina = RequireColumnIndex(TBL_PARCELE, COL_PAR_POVRSINA, _
                                     "frmOtkup.cmbKooperant_Change")

    Dim i As Long
    Dim povrsina As Double

    For i = 1 To UBound(parData, 1)
        If CStr(parData(i, colKoop)) = kooperantID Then

            povrsina = 0
            If IsNumeric(parData(i, colPovrsina)) Then povrsina = CDbl(parData(i, colPovrsina))

            cmbParcela.AddItem CStr(parData(i, colKat)) & " | " & _
                               CStr(parData(i, colKultura)) & " | " & _
                               Format$(povrsina, "#,##0.00") & " ha (" & _
                               CStr(parData(i, colID)) & ")"
        End If
    Next i

    Exit Sub

EH:
    LogErr "frmOtkup.cmbKooperant_Change"
    cmbParcela.Clear
End Sub

Private Sub cmbParcela_Change()
    On Error GoTo EH

    If cmbParcela.ListIndex < 0 Then Exit Sub

    Dim parcelaID As String
    parcelaID = ExtractParcelaID(cmbParcela.value)

    If parcelaID = "" Then Exit Sub

    Dim kultura As String
    kultura = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KULTURA))

    If kultura = "" Then Exit Sub

    Dim vrsta As String
    vrsta = CStr(LookupValue(TBL_KULTURE, "SortaVoca", kultura, "VrstaVoca"))

    If vrsta <> "" Then
        cmbVrstaVoca.value = vrsta

        On Error Resume Next
        cmbSortaVoca.value = kultura
        On Error GoTo EH
    End If

    Exit Sub

EH:
    LogErr "frmOtkup.cmbParcela_Change"
End Sub

' Helper: ParcelaID aus Display-String extrahieren
' Format: "KatBroj | Kultura | 2.50 ha (PAR001)"
Private Function ExtractParcelaID(ByVal display As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStrRev(display, "(")
    p2 = InStrRev(display, ")")
    If p1 > 0 And p2 > p1 Then
        ExtractParcelaID = Mid$(display, p1 + 1, p2 - p1 - 1)
    End If
End Function

' ============================================================
' OTKUP
' ============================================================

Private Sub btnUnos_Click()
    On Error GoTo EH

    ButtonActive btnUnos

    If cmbOtkupnoMesto.value = "" Then
        MsgBox "Izaberite otkupno mesto!", vbExclamation, APP_NAME
        cmbOtkupnoMesto.SetFocus
        Exit Sub
    End If

    If cmbKooperant.ListIndex < 0 Then
        MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
        cmbKooperant.SetFocus
        Exit Sub
    End If

    If cmbVrstaVoca.value = "" Then
        MsgBox "Izaberite vrstu voca!", vbExclamation, APP_NAME
        cmbVrstaVoca.SetFocus
        Exit Sub
    End If

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        MsgBox "Unesite ispravan datum!", vbExclamation, APP_NAME
        txtDatum.SetFocus
        Exit Sub
    End If

    Dim kolicinaI As Double
    If Not TryParseDouble(txtKolicina.value, kolicinaI) Or kolicinaI <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtKolicina.SetFocus
        Exit Sub
    End If

    Dim cenaI As Double
    If Not TryParseDouble(txtCena.value, cenaI) Or cenaI <= 0 Then
        MsgBox "Unesite ispravnu cenu!", vbExclamation, APP_NAME
        txtCena.SetFocus
        Exit Sub
    End If

    Dim kolicinaII As Double
    Dim cenaII As Double

    If chkDveKlase.value Then
        If Not TryParseDouble(txtKolicinaKLII.value, kolicinaII) Or kolicinaII <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKLII.SetFocus
            Exit Sub
        End If

        If Not TryParseDouble(txtCenaKLII.value, cenaII) Or cenaII <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKLII.SetFocus
            Exit Sub
        End If
    End If

    Dim kolAmb As Long
    If Trim$(txtKolAmbalaze.value) <> "" Then
        If Not TryParseLong(txtKolAmbalaze.value, kolAmb) Then
            MsgBox "Unesite ispravnu kolicinu ambalaže!", vbExclamation, APP_NAME
            txtKolAmbalaze.SetFocus
            Exit Sub
        End If
    End If

    If kolAmb > 0 And cmbTipAmbalaze.value = "" Then
        MsgBox "Izaberite tip ambalaže!", vbExclamation, APP_NAME
        cmbTipAmbalaze.SetFocus
        Exit Sub
    End If

    Dim novac As Double
    If Trim$(txtNovac.value) <> "" Then
        If Not TryParseDouble(txtNovac.value, novac) Or novac < 0 Then
            MsgBox "Unesite ispravan iznos novca!", vbExclamation, APP_NAME
            txtNovac.SetFocus
            Exit Sub
        End If
    End If

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)

    If stanicaID = "" Then
        MsgBox "Nije pronaden ID otkupnog mesta!", vbExclamation, APP_NAME
        cmbOtkupnoMesto.SetFocus
        Exit Sub
    End If

    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbKooperant.value)

    If kooperantID = "" Then
        MsgBox "Nije pronaden ID kooperanta!", vbExclamation, APP_NAME
        cmbKooperant.SetFocus
        Exit Sub
    End If

    Dim vozacID As String
    If cmbVozac.value <> "" Then
        vozacID = ExtractIDFromDisplay(cmbVozac.value)
    End If

    If Trim$(txtBrojDokumenta.value) <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_OTKUP, COL_OTK_BR_DOK, _
                                Trim$(txtBrojDokumenta.value), COL_OTK_DATUM)

        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    Dim parcelaID As String

    If cmbParcela.ListIndex >= 0 Then
        parcelaID = ExtractParcelaID(cmbParcela.value)

        If parcelaID <> "" Then
            Dim parKultura As String
            parKultura = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KULTURA))

            If parKultura <> "" Then
                Dim selectedKultura As String
                selectedKultura = Trim$(cmbSortaVoca.value)

                If selectedKultura <> "" Then
                    If StrComp(parKultura, selectedKultura, vbTextCompare) <> 0 Then
                        Dim ans As VbMsgBoxResult

                        ans = MsgBox("Kultura parcele (" & parKultura & _
                                     ") ne odgovara izabranoj sorti/kulturi (" & _
                                     selectedKultura & ")!" & vbCrLf & vbCrLf & _
                                     "Želite li ipak da nastavite?", _
                                     vbExclamation + vbYesNo, APP_NAME)

                        If ans = vbNo Then Exit Sub
                    End If
                End If
            End If
        End If
    End If

    Dim result As String

    result = SaveOtkupMulti_TX( _
        datum:=datumDok, _
        kooperantID:=kooperantID, _
        stanicaID:=stanicaID, _
        vrstaVoca:=cmbVrstaVoca.value, _
        sortaVoca:=cmbSortaVoca.value, _
        kolicinaI:=kolicinaI, _
        cenaI:=cenaI, _
        tipAmb:=cmbTipAmbalaze.value, _
        kolAmb:=kolAmb, _
        vozacID:=vozacID, _
        brDok:=Trim$(txtBrojDokumenta.value), _
        novac:=novac, _
        primalac:=txtPrimalac.value, _
        parcelaID:=parcelaID, _
        brojZbirne:=Trim$(txtBrojZbirne.value), _
        hasKlasaII:=chkDveKlase.value, _
        kolicinaII:=kolicinaII, _
        cenaII:=cenaII)

    If result = "" Then
        MsgBox "Greška pri cuvanju otkupa. Promene su vracene.", vbCritical, APP_NAME
        Exit Sub
    End If

    MsgBox "Otkup sacuvan: " & result, vbInformation, APP_NAME

    ClearOtkupFields
    Exit Sub

EH:
    LogErr "frmOtkup.btnUnos"
    MsgBox "Greška pri unosu otkupa: " & Err.Description, vbCritical, APP_NAME
End Sub
Private Sub ClearOtkupFields()
    txtBrojDokumenta.value = ""
    txtKolicina.value = ""
    txtCena.value = ""
    txtKolAmbalaze.value = ""
    txtNovac.value = "0"
    txtPrimalac.value = ""
    cmbParcela.Clear
    
    ' Klasa II zurücksetzen
    chkDveKlase.value = False
    DisableField txtKolicinaKLII
    DisableField txtCenaKLII
    lblUkupnoKG.caption = ""
    
    txtKolicina.SetFocus
End Sub

Private Sub btnStornoOtkup_Click()
    ButtonActive btnStornoOtkup
End Sub
' ============================================================
' NAVIGATION
' ============================================================

Private Sub btnPovratak_Click()
    ButtonActive btnPovratak
    Me.Hide
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        frmOtkupAPP.Show
    End If
End Sub

