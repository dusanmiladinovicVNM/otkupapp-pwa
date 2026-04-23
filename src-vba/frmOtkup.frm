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

    hwnd = FindWindow("ThunderDFrame", Me.Caption)

    If hwnd <> 0 Then
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style And Not WS_CAPTION
        SetWindowLong hwnd, GWL_STYLE, style
        DrawMenuBar hwnd
    End If
End Sub

Private Sub UserForm_Activate()
    If Not mChromeRemoved Then
        Me.Caption = ""
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
    txtDatum.Value = Format$(Date, "d.m.yyyy")
    
    ' ComboBoxen füllen
    FillCmb cmbVrstaVoca, GetLookupList(TBL_KULTURE, "VrstaVoca")
    FillCmb cmbOtkupnoMesto, GetLookupList(TBL_STANICE, "Naziv")
    FillCmb cmbVozac, GetVozacDisplayList()
    FillCmb cmbTipAmbalaze, Array(AMB_12_1, AMB_6_1)
    
    ' Numerische Felder auf 0 setzen
    txtKolicina.Value = ""
    txtCena.Value = ""
    txtKolAmbalaze.Value = ""
    txtNovac.Value = "0"
    
    ' Klasa II – initial disabled
    DisableField txtKolicinaKLII
    DisableField txtCenaKLII
    chkDveKlase.Value = False
    lblUkupnoKG.Caption = ""
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
    If chkDveKlase.Value Then
        EnableField txtKolicinaKLII
        EnableField txtCenaKLII
        StyleTextBox txtKolicinaKLII
        StyleTextBox txtCenaKLII
    Else
        DisableField txtKolicinaKLII
        DisableField txtCenaKLII
        lblUkupnoKG.Caption = ""
    End If
End Sub

Private Sub txtKolicinaKlII_Change()
    UpdateUkupnoKg
End Sub

Private Sub txtKolicina_Change()
    UpdateUkupnoKg
End Sub

Private Sub UpdateUkupnoKg()
    If Not chkDveKlase.Value Then
        lblUkupnoKG.Caption = ""
        Exit Sub
    End If
    Dim kl1 As Double, kl2 As Double
    If IsNumeric(txtKolicina.Value) Then kl1 = CDbl(txtKolicina.Value)
    If IsNumeric(txtKolicinaKLII.Value) Then kl2 = CDbl(txtKolicinaKLII.Value)
    lblUkupnoKG.Caption = "Ukupno: " & Format$(kl1 + kl2, "#,##0.00") & " kg"
End Sub

' ============================================================
' KASKADIERUNG – VrstaVoca ? SortaVoca
' ============================================================

Private Sub cmbVrstaVoca_Change()
    ' Wenn VrstaVoca gewählt wird, SortaVoca-Liste filtern
    cmbSortaVoca.Clear
    If cmbVrstaVoca.Value <> "" Then
        FillCmb cmbSortaVoca, _
            GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", cmbVrstaVoca.Value)
    End If
End Sub

Private Sub cmbOtkupnoMesto_Change()
    cmbKooperant.Clear
    If cmbOtkupnoMesto.Value = "" Then Exit Sub
    
    Dim stanicaID As String
    stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbOtkupnoMesto.Value, "StanicaID"))
    FillComboKooperantiByStanica cmbKooperant, stanicaID
End Sub

Private Sub cmbKooperant_Change()
    cmbParcela.Clear
    If cmbKooperant.ListIndex < 0 Then Exit Sub
    
    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbKooperant.Value)
    
    ' Parcele dieses Kooperanten laden
    Dim parData As Variant
    parData = GetTableData(TBL_PARCELE)
    If IsEmpty(parData) Then Exit Sub
    
    Dim colKoop As Long, colID As Long, colKat As Long
    Dim colKultura As Long, colPovrsina As Long
    colID = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
    colKoop = GetColumnIndex(TBL_PARCELE, COL_PAR_KOOP)
    colKat = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_BROJ)
    colKultura = GetColumnIndex(TBL_PARCELE, COL_PAR_KULTURA)
    colPovrsina = GetColumnIndex(TBL_PARCELE, COL_PAR_POVRSINA)
    
    Dim i As Long
    For i = 1 To UBound(parData, 1)
        If CStr(parData(i, colKoop)) = kooperantID Then
            ' Display: "KatBroj | Kultura | 2.5 ha (PAR001)"
            cmbParcela.AddItem CStr(parData(i, colKat)) & " | " & _
                               CStr(parData(i, colKultura)) & " | " & _
                               Format$(parData(i, colPovrsina), "#,##0.00") & " ha (" & _
                               CStr(parData(i, colID)) & ")"
        End If
    Next i
End Sub

Private Sub cmbParcela_Change()
    If cmbParcela.ListIndex < 0 Then Exit Sub
    
    Dim parcelaID As String
    parcelaID = ExtractParcelaID(cmbParcela.Value)
    
    ' Kultura der Parcela holen
    Dim kultura As String
    kultura = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KULTURA))
    
    If kultura = "" Then Exit Sub
    
    ' Auto-Vorfüllung: Kultura ? VrstaVoca/SortaVoca
    ' Kultura in tblKulture nachschlagen
    Dim vrsta As String
    vrsta = CStr(LookupValue(TBL_KULTURE, "SortaVoca", kultura, "VrstaVoca"))
    
    If vrsta <> "" Then
        ' VrstaVoca setzen (löst cmbVrstaVoca_Change aus ? füllt SortaVoca)
        cmbVrstaVoca.Value = vrsta
        ' SortaVoca direkt setzen
        On Error Resume Next
        cmbSortaVoca.Value = kultura
        On Error GoTo 0
    End If
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
    ' Validierung
    If cmbOtkupnoMesto.Value = "" Then
        MsgBox "Izaberite otkupno mesto!", vbExclamation, APP_NAME
        cmbOtkupnoMesto.SetFocus
        Exit Sub
    End If
    If cmbKooperant.ListIndex < 0 Then
        MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
        cmbKooperant.SetFocus
        Exit Sub
    End If
    If cmbVrstaVoca.Value = "" Then
        MsgBox "Izaberite vrstu voca!", vbExclamation, APP_NAME
        cmbVrstaVoca.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtKolicina.Value) Or val(txtKolicina.Value) <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtKolicina.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtCena.Value) Or val(txtCena.Value) <= 0 Then
        MsgBox "Unesite ispravnu cenu!", vbExclamation, APP_NAME
        txtCena.SetFocus
        Exit Sub
    End If
    
    ' Klasa II Validierung
    If chkDveKlase.Value Then
        If Not IsNumeric(txtKolicinaKLII.Value) Or val(txtKolicinaKLII.Value) <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKLII.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtCenaKLII.Value) Or val(txtCenaKLII.Value) <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKLII.SetFocus
            Exit Sub
        End If
    End If
    
    Dim stanicaID As String
    stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbOtkupnoMesto.Value, "StanicaID"))
    
    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbKooperant.Value)
    
    Dim vozacID As String
    vozacID = ExtractIDFromDisplay(cmbVozac.Value)
    
    Dim kolAmb As Long
    If IsNumeric(txtKolAmbalaze.Value) Then kolAmb = CLng(txtKolAmbalaze.Value)
    
    Dim novac As Double
    If IsNumeric(txtNovac.Value) Then novac = CDbl(txtNovac.Value)
    
    ' Duplikat-Check
    If txtBrojDokumenta.Value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_OTKUP, COL_OTK_BR_DOK, txtBrojDokumenta.Value, COL_OTK_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    
    Dim parcelaID As String
    If cmbParcela.ListIndex >= 0 Then
        parcelaID = ExtractParcelaID(cmbParcela.Value)
        
        Dim parKultura As String
        parKultura = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, parcelaID, COL_PAR_KULTURA))
        
        If parKultura <> "" And cmbVrstaVoca.Value <> "" Then
            If StrComp(parKultura, cmbVrstaVoca.Value, vbTextCompare) <> 0 Then
                Dim ans As VbMsgBoxResult
                ans = MsgBox("Kultura parcele (" & parKultura & ") ne odgovara vrsti voca (" & _
                             cmbVrstaVoca.Value & ")!" & vbCrLf & vbCrLf & _
                             "Zelite li ipak da nastavite?", _
                             vbExclamation + vbYesNo, APP_NAME)
                If ans = vbNo Then Exit Sub
            End If
        End If
    End If
    
    
    ' --- Klasa I speichern (immer) ---
    Dim result As String
    result = SaveOtkup_TX( _
        datum:=CDate(txtDatum.Value), _
        kooperantID:=kooperantID, _
        stanicaID:=stanicaID, _
        vrstaVoca:=cmbVrstaVoca.Value, _
        sortaVoca:=cmbSortaVoca.Value, _
        kolicina:=CDbl(txtKolicina.Value), _
        cena:=CDbl(txtCena.Value), _
        tipAmb:=cmbTipAmbalaze.Value, _
        kolAmb:=kolAmb, _
        vozacID:=vozacID, _
        brDok:=txtBrojDokumenta.Value, _
        novac:=novac, _
        primalac:=txtPrimalac.Value, _
        klasa:=KLASA_I, _
        parcelaID:=parcelaID, _
        brojZbirne:=txtBrojZbirne.Value)
    
    If result = "" Then
        MsgBox "Greska pri cuvanju!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' --- Klasa II speichern (wenn aktiv) ---
    Dim resultII As String
    If chkDveKlase.Value Then
        resultII = SaveOtkup_TX( _
            datum:=CDate(txtDatum.Value), _
            kooperantID:=kooperantID, _
            stanicaID:=stanicaID, _
            vrstaVoca:=cmbVrstaVoca.Value, _
            sortaVoca:=cmbSortaVoca.Value, _
            kolicina:=CDbl(txtKolicinaKLII.Value), _
            cena:=CDbl(txtCenaKLII.Value), _
            tipAmb:=cmbTipAmbalaze.Value, _
            kolAmb:=0, _
            vozacID:=vozacID, _
            brDok:=txtBrojDokumenta.Value, _
            novac:=0, _
            primalac:=txtPrimalac.Value, _
            klasa:=KLASA_II, _
            parcelaID:=parcelaID, _
            brojZbirne:=txtBrojZbirne.Value)
        
        If resultII <> "" Then
            MsgBox "Otkup sacuvan: " & result & " + " & resultII, vbInformation, APP_NAME
        Else
            MsgBox "Klasa I sacuvana (" & result & "), greska pri Klasi II!", vbExclamation, APP_NAME
        End If
    Else
        MsgBox "Otkup sacuvan: " & result, vbInformation, APP_NAME
    End If
    
    If novac > 0 Then
        Dim koopNaziv As String
        koopNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime")) & " " & _
                    CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"))
        
        SaveNovac_TX txtBrojDokumenta.Value, CDate(txtDatum.Value), _
                     koopNaziv, kooperantID, "Kooperant", stanicaID, _
                     kooperantID, "", "", "", 0, novac, _
                     txtPrimalac.Value
    End If
    
    ' Koop-Avans verrechnen
    ApplyAvansToOtkup_TX kooperantID, result
    If chkDveKlase.Value And resultII <> "" Then
        ApplyAvansToOtkup_TX kooperantID, resultII
    End If
    
    ClearOtkupFields
    Exit Sub
EH:
    LogErr "frmOtkup.btnUnos"
    MsgBox "Greska pri unosu otkupa: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub ClearOtkupFields()
    txtBrojDokumenta.Value = ""
    txtKolicina.Value = ""
    txtCena.Value = ""
    txtKolAmbalaze.Value = ""
    txtNovac.Value = "0"
    txtPrimalac.Value = ""
    cmbParcela.Clear
    
    ' Klasa II zurücksetzen
    chkDveKlase.Value = False
    DisableField txtKolicinaKLII
    DisableField txtCenaKLII
    lblUkupnoKG.Caption = ""
    
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
        frmMain.Show
    End If
End Sub
