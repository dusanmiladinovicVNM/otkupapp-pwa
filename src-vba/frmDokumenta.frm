VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDokumenta 
   ClientHeight    =   13755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15600
   OleObjectBlob   =   "frmDokumenta.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDokumenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ============================================================
' frmDokumenta v2.1 – Otpremnica + Zbirna + Prijemnica
' Ein Form, 6 Frames, kein MultiPage
' ============================================================

Private m_SetupDone As Boolean
Private m_OtkupIDs() As String

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
    
    If m_SetupDone Then Exit Sub
    m_SetupDone = True
    
    'ModernUI_ApplyTheme Me
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    
    ' action buttons
    StylePrimaryButton btnUnosOtp, "Unos otpremnice"
    StylePrimaryButton btnUnosZbr, "Unos zbirne"
    StylePrimaryButton btnUnosPrij, "Unos prijemnice"
    StylePrimaryButton btnUnosOMUlaz, "Unos OM ulaz"
    StylePrimaryButton btnUnosIzlaz, "Unos izlaz"
    StyleExitButton btnPovratak, "Povratak"
    StyleExitButton btnStorno, "Storno"

    ' status labels
    StyleLabel lblValidacijaKG, TXT_MUTED(), True
    StyleLabel lblValidacijaAmb, TXT_MUTED(), True
    StyleLabel lblManjak, TXT_MUTED(), True
    StyleLabel lblProsekGajbe, TXT_MUTED(), False
    StyleLabel lblOMAvansSaldo, TXT_MUTED(), True
    StyleLabel lblStornoWarning, CLR_WARNING(), True

    ' important frames
    StyleFrame fraOtpremnica
    StyleFrame fraZbirna
    StyleFrame fraPrijemnica
    StyleFrame fraOMUlaz
    StyleFrame fraIzlazKupci
    StyleFrame fraStorno
        
    'Me.Caption = "Dokumenta - Otpremnica / Zbirna / Prijemnica"
    
    txtDatum.Value = Format$(Date, "d.m.yyyy")
    
    FillCmb cmbVrstaVoca, GetLookupList(TBL_KULTURE, "VrstaVoca")
    FillCmb cmbOtkupnoMesto, GetLookupList(TBL_STANICE, "Naziv")
    FillCmb cmbKupac, GetLookupList(TBL_KUPCI, "Naziv")
    FillCmb cmbVozac, GetVozacDisplayList()
    FillCmb cmbTipAmbOtp, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbZbr, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbPrij, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbOMUlaz, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbIzlaz, Array(AMB_12_1, AMB_6_1)

    
    lblValidacijaKG.Caption = ""
    lblValidacijaAmb.Caption = ""
    lblManjak.Caption = ""
    
    ' Druga klasa:
    DisableField txtKolicinaKlIIOtp
    DisableField txtCenaKlIIOtp
    chkDveKlaseOtp.Value = False
    lblUkupnoKgOtp.Caption = ""
    
    DisableField txtUkupnoKgKlIIZbr
    chkDveKlaseZbr.Value = False
    lblUkupnoKgZbr.Caption = ""
    
    DisableField txtKolicinaKlIIPrij
    DisableField txtCenaKlIIPrij
    chkDveKlasePrij.Value = False
    lblUkupnoKgPrij.Caption = ""
    lblProsekGajbe.Caption = ""
    
        ' Storno ComboBox füllen
    With cmbStornoDokument
        .Clear
        .AddItem "Otkup"
        .AddItem "Otpremnica"
        .AddItem "Zbirna"
        .AddItem "Prijemnica"
        .AddItem "Faktura"
        .AddItem "Novac"
    End With
    
    CheckVerwaisteDokumente
End Sub

Private Sub chkDveKlaseOtp_Click()
    If chkDveKlaseOtp.Value Then
        EnableField txtKolicinaKlIIOtp
        EnableField txtCenaKlIIOtp
        StyleTextBox txtKolicinaKlIIOtp
        StyleTextBox txtCenaKlIIOtp
    Else
        DisableField txtKolicinaKlIIOtp
        DisableField txtCenaKlIIOtp
        lblUkupnoKgOtp.Caption = ""
    End If
End Sub

Private Sub chkDveKlaseZbr_Click()
    If chkDveKlaseZbr.Value Then
        EnableField txtUkupnoKgKlIIZbr
        StyleTextBox txtUkupnoKgKlIIZbr
    Else
        DisableField txtUkupnoKgKlIIZbr
    End If
End Sub

Private Sub chkDveKlasePrij_Click()
    If chkDveKlasePrij.Value Then
        EnableField txtKolicinaKlIIPrij
        EnableField txtCenaKlIIPrij
        StyleTextBox txtKolicinaKlIIPrij
        StyleTextBox txtCenaKlIIPrij
    Else
        DisableField txtKolicinaKlIIPrij
        DisableField txtCenaKlIIPrij
    End If
End Sub

Private Sub txtKolicinaOtp_Change()
    UpdateUkupnoKgOtp
End Sub

Private Sub txtKolicinaKlIIOtp_Change()
    UpdateUkupnoKgOtp
End Sub

Private Sub UpdateUkupnoKgOtp()
    If Not chkDveKlaseOtp.Value Then
        lblUkupnoKgOtp.Caption = ""
        Exit Sub
    End If
    Dim kl1 As Double, kl2 As Double
    If IsNumeric(txtKolicinaOtp.Value) Then kl1 = CDbl(txtKolicinaOtp.Value)
    If IsNumeric(txtKolicinaKlIIOtp.Value) Then kl2 = CDbl(txtKolicinaKlIIOtp.Value)
    lblUkupnoKgOtp.Caption = "Ukupno: " & Format$(kl1 + kl2, "#,##0.00") & " kg"
End Sub

Private Sub txtKolicinaPrij_Change()
    If txtBrojZbirnePrij.Value <> "" Then
        UpdateManjak txtBrojZbirnePrij.Value
    End If
    ' Prosek Gajbe auch updaten
    txtKolAmbPrij_Change
End Sub

Private Sub txtKolicinaKlIIPrij_Change()
    If txtBrojZbirnePrij.Value <> "" Then
        UpdateManjak txtBrojZbirnePrij.Value
    End If
End Sub

Private Sub txtKolAmbPrij_Change()
    If IsNumeric(txtKolicinaPrij.Value) And IsNumeric(txtKolAmbPrij.Value) Then
        If CLng(txtKolAmbPrij.Value) > 0 Then
            lblProsekGajbe.Caption = "Prosek gajbe: " & _
                Format$(CDbl(txtKolicinaPrij.Value) / CLng(txtKolAmbPrij.Value), "#,##0.00") & " kg"
        End If
    Else
        lblProsekGajbe.Caption = ""
    End If
End Sub

Private Sub cmbOtkupnoMesto_Change()
    Dim stanicaID As String
    If cmbOtkupnoMesto.Value = "" Then
        cmbPrimalacOMUlaz.Clear
        Exit Sub
    End If
    stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbOtkupnoMesto.Value, "StanicaID"))
    FillComboKooperantiByStanica cmbPrimalacOMUlaz, stanicaID
    UpdateOMAvansSaldo
End Sub

' ============================================================
' KASKADIERUNG
' ============================================================

Private Sub cmbVrstaVoca_Change()
    cmbSortaVoca.Clear
    If cmbVrstaVoca.Value <> "" Then
        FillCmb cmbSortaVoca, _
            GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", cmbVrstaVoca.Value)
    End If
End Sub

Private Sub cmbKupac_Change()
    cmbHladnjaca.Clear
    cmbPogon.Clear
    
    If cmbKupac.Value <> "" Then
        Dim kupacID As String
        kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbKupac.Value, "KupacID"))
        
        Dim hlad As String
        hlad = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Hladnjaca"))
        If hlad <> "" Then cmbHladnjaca.AddItem hlad
    End If
    
    FillOpenFakture
End Sub

' ============================================================
' OTPREMNICA UNOS (Izlaz OM)
' ============================================================

Private Sub btnUnosOtp_Click()
    On Error GoTo EH
    ' Validierung
    If cmbOtkupnoMesto.Value = "" Then
        MsgBox "Izaberite otkupno mesto!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If cmbVozac.Value = "" Then
        MsgBox "Izaberite vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not IsNumeric(txtKolicinaOtp.Value) Or val(txtKolicinaOtp.Value) <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtKolicinaOtp.SetFocus
        Exit Sub
    End If
    
    If chkDveKlaseOtp.Value Then
        If Not IsNumeric(txtKolicinaKlIIOtp.Value) Or val(txtKolicinaKlIIOtp.Value) <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKlIIOtp.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtCenaKlIIOtp.Value) Or val(txtCenaKlIIOtp.Value) <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKlIIOtp.SetFocus
            Exit Sub
        End If
    End If
    
    Dim stanicaID As String
    stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbOtkupnoMesto.Value, "StanicaID"))
    
    Dim vozacID As String
    vozacID = ExtractIDFromDisplay(cmbVozac.Value)
    
    Dim cena As Double
    If IsNumeric(txtCenaOtp.Value) Then cena = CDbl(txtCenaOtp.Value)
    
    Dim kolAmb As Long
    If IsNumeric(txtKolAmbOtp.Value) Then kolAmb = CLng(txtKolAmbOtp.Value)
    
    'Duplikat Check
    If txtBrojOtp.Value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_OTPREMNICA, COL_OTP_BROJ, txtBrojOtp.Value, COL_OTP_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    
    ' --- Klasa I ---
    Dim result As String
    result = SaveOtpremnica_TX( _
        datum:=CDate(txtDatum.Value), _
        stanicaID:=stanicaID, _
        vozacID:=vozacID, _
        brojOtp:=txtBrojOtp.Value, _
        brojZbirne:=txtBrojZbirneOtp.Value, _
        vrsta:=cmbVrstaVoca.Value, _
        sorta:=cmbSortaVoca.Value, _
        kolicina:=CDbl(txtKolicinaOtp.Value), _
        cena:=cena, _
        tipAmb:=cmbTipAmbOtp.Value, _
        kolAmb:=kolAmb, _
        klasa:=KLASA_I)
    
    If result = "" Then
        MsgBox "Greska!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' --- Klasa II ---
    If chkDveKlaseOtp.Value Then
        Dim resultII As String
        resultII = SaveOtpremnica_TX( _
            datum:=CDate(txtDatum.Value), _
            stanicaID:=stanicaID, _
            vozacID:=vozacID, _
            brojOtp:=txtBrojOtp.Value, _
            brojZbirne:=txtBrojZbirneOtp.Value, _
            vrsta:=cmbVrstaVoca.Value, _
            sorta:=cmbSortaVoca.Value, _
            kolicina:=CDbl(txtKolicinaKlIIOtp.Value), _
            cena:=CDbl(txtCenaKlIIOtp.Value), _
            tipAmb:=cmbTipAmbOtp.Value, _
            kolAmb:=0, _
            klasa:=KLASA_II)
        
        If resultII <> "" Then
            MsgBox "Otpremnica sacuvana: " & result & " + " & resultII, vbInformation, APP_NAME
        Else
            MsgBox "Klasa I sacuvana, greska pri Klasi II!", vbExclamation, APP_NAME
        End If
    Else
        MsgBox "Otpremnica sacuvana: " & result, vbInformation, APP_NAME
    End If
    
    ClearOtpremnicaFields
    If txtBrojZbirne.Value <> "" Then UpdateValidacija
    Exit Sub
EH:
    LogErr "frmDokumenta.btnUnosOtp"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub ClearOtpremnicaFields()
    txtBrojOtp.Value = ""
    txtKolicinaOtp.Value = ""
    txtCenaOtp.Value = ""
    txtKolAmbOtp.Value = ""
    chkDveKlaseOtp.Value = False
    DisableField txtKolicinaKlIIOtp
    DisableField txtCenaKlIIOtp
    lblUkupnoKgOtp.Caption = ""
End Sub

Private Sub btnUnosOMUlaz_Click()
    On Error GoTo EH
    If cmbOtkupnoMesto.Value = "" Then
        MsgBox "Izaberite otkupno mesto!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim stanicaID As String
    stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbOtkupnoMesto.Value, "StanicaID"))
    
    Dim vozacID As String
    If cmbVozac.Value <> "" Then vozacID = ExtractIDFromDisplay(cmbVozac.Value)
    
    ' Duplikat-Check
    If txtBrojDokOMUlaz.Value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_AMBALAZA, COL_AMB_DOK_ID, txtBrojDokOMUlaz.Value, COL_AMB_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    
    ' Ambalaza
    Dim kolAmb As Long
    If IsNumeric(txtKolAmbOMUlaz.Value) Then kolAmb = CLng(txtKolAmbOMUlaz.Value)
    
    If kolAmb > 0 Then
        TrackAmbalaza CDate(txtDatum.Value), cmbTipAmbOMUlaz.Value, kolAmb, _
                      "Ulaz", stanicaID, "Stanica", _
                      vozacID, txtBrojDokOMUlaz.Value, DOK_TIP_OM_ULAZ
    End If
    
    ' Novac
    Dim novac As Double
    Dim kooperantID As String
    Dim otkupID As String
    Dim tip As String
    
    If IsNumeric(txtNovacOMUlaz.Value) Then novac = CDbl(txtNovacOMUlaz.Value)
    If novac > 0 Then
        ' Primalac = Kooperant der das Geld bekommt

        If cmbPrimalacOMUlaz.Value <> "" Then
            kooperantID = ExtractIDFromDisplay(cmbPrimalacOMUlaz.Value)
            
            If cmbOtkupBlok.ListIndex >= 0 Then
                otkupID = m_OtkupIDs(cmbOtkupBlok.ListIndex)
                
                ' Validierung: nicht mehr als Restbetrag
                Dim otkupi As Variant
                otkupi = GetOpenOtkupi(kooperantID)
                If Not IsEmpty(otkupi) Then
                    Dim preostalo As Double
                    preostalo = CDbl(otkupi(cmbOtkupBlok.ListIndex + 1, 5))
                    If novac > preostalo Then
                        MsgBox "Iznos (" & Format$(novac, "#,##0") & ") veci od preostalog (" & _
                               Format$(preostalo, "#,##0") & ")!", vbExclamation, APP_NAME
                        Exit Sub
                    End If
                End If
                
                If tglIzOMAvansa.Value Then
                    Dim omSaldo As Double
                    omSaldo = GetOMAvansSaldo(stanicaID)
                    If novac > omSaldo Then
                        MsgBox "Nedovoljno OM avansa! Raspolozivo: " & _
                               Format$(omSaldo, "#,##0") & " RSD", vbExclamation, APP_NAME
                        Exit Sub
                    End If
                    tip = NOV_KES_OTKUPAC_KOOP
                Else
                    tip = NOV_VIRMAN_FIRMA_KOOP
                End If
            Else
                tip = NOV_VIRMAN_AVANS_KOOP
            End If
        Else
            kooperantID = ""
            otkupID = ""
            tip = NOV_KES_FIRMA_OTKUPAC
        End If

        SaveNovac_TX txtBrojDokOMUlaz.Value, CDate(txtDatum.Value), _
                     cmbOtkupnoMesto.Value, stanicaID, "OM", stanicaID, _
                     kooperantID, "", cmbVrstaVoca.Value, tip, 0, novac, _
                     cmbPrimalacOMUlaz.Value, otkupID

        ' Otkup Status aktualisieren
        If otkupID <> "" Then UpdateOtkupStatus otkupID
        UpdateOMAvansSaldo
    End If
    
    MsgBox "Sacuvano!", vbInformation, APP_NAME
    txtBrojDokOMUlaz.Value = ""
    txtKolAmbOMUlaz.Value = ""
    txtNovacOMUlaz.Value = ""
    cmbPrimalacOMUlaz.Value = ""
    cmbOtkupBlok.Clear
    Exit Sub
EH:
    LogErr "frmDokumenta.btnUnosOMUlaz"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub cmbPrimalacOMUlaz_Change()
    cmbOtkupBlok.Clear
    tglIzOMAvansa.Enabled = (cmbPrimalacOMUlaz.Value <> "")
    
    If cmbPrimalacOMUlaz.Value = "" Then
        tglIzOMAvansa.Value = False
        Exit Sub
    End If
    
    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbPrimalacOMUlaz.Value)
    FillOpenOtkupi
End Sub

Private Sub UpdateOMAvansSaldo()
    If cmbOtkupnoMesto.Value = "" Then
        lblOMAvansSaldo.Caption = ""
        Exit Sub
    End If
    Dim stanicaID As String
    stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbOtkupnoMesto.Value, "StanicaID"))
    
    Dim saldo As Double
    saldo = GetOMAvansSaldo(stanicaID)
    
    If saldo > 0 Then
        lblOMAvansSaldo.Caption = "OM Avans: " & Format$(saldo, "#,##0") & " RSD"
    Else
        lblOMAvansSaldo.Caption = ""
    End If
End Sub

' ============================================================
' ZBIRNA UNOS
' ============================================================

Private Sub btnUnosZbr_Click()
On Error GoTo EH
    If cmbVozac.Value = "" Then
        MsgBox "Izaberite vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If cmbKupac.Value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If txtBrojZbirne.Value = "" Then
        MsgBox "Unesite broj zbirne!", vbExclamation, APP_NAME
        txtBrojZbirne.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtUkupnoKGZbr.Value) Or val(txtUkupnoKGZbr.Value) <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtUkupnoKGZbr.SetFocus
        Exit Sub
    End If
    
    If chkDveKlaseZbr.Value Then
        If Not IsNumeric(txtUkupnoKgKlIIZbr.Value) Or val(txtUkupnoKgKlIIZbr.Value) <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtUkupnoKgKlIIZbr.SetFocus
            Exit Sub
        End If
    End If
    
    Dim kupacID As String
    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbKupac.Value, "KupacID"))
    
    Dim vozacID As String
    vozacID = ExtractIDFromDisplay(cmbVozac.Value)
    
    Dim hladnjaca As String
    hladnjaca = cmbHladnjaca.Value
    
    Dim pogon As String
    pogon = cmbPogon.Value
    
    Dim ukupnoAmb As Long
    If IsNumeric(txtUkupnoAmbZbr.Value) Then ukupnoAmb = CLng(txtUkupnoAmbZbr.Value)
    
    If Not UpdateValidacija() Then
        MsgBox "Validacija nije zelena. Proverite kg i ambalažu (razlika mora biti 0).", vbExclamation, APP_NAME
        Exit Sub
    End If

    'Duplikat Check
    If txtBrojZbirne.Value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_ZBIRNA, COL_ZBR_BROJ, txtBrojZbirne.Value, COL_ZBR_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    
    ' --- Klasa I ---
    Dim result As String
    result = SaveZbirna_TX( _
        datum:=CDate(txtDatum.Value), _
        vozacID:=ExtractIDFromDisplay(cmbVozac.Value), _
        brojZbirne:=txtBrojZbirne.Value, _
        kupacID:=kupacID, _
        hladnjaca:=cmbHladnjaca.Value, _
        pogon:=cmbPogon.Value, _
        vrstaVoca:=cmbVrstaVoca.Value, _
        sortaVoca:=cmbSortaVoca.Value, _
        ukupnoKol:=CDbl(txtUkupnoKGZbr.Value), _
        tipAmb:=cmbTipAmbZbr.Value, _
        ukupnoAmb:=ukupnoAmb, _
        klasa:=KLASA_I)
    
    If result = "" Then
        MsgBox "Greska!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' --- Klasa II ---
    If chkDveKlaseZbr.Value Then
        Dim resultII As String
        resultII = SaveZbirna_TX( _
            datum:=CDate(txtDatum.Value), _
            vozacID:=ExtractIDFromDisplay(cmbVozac.Value), _
            brojZbirne:=txtBrojZbirne.Value, _
            kupacID:=kupacID, _
            hladnjaca:=cmbHladnjaca.Value, _
            pogon:=cmbPogon.Value, _
            vrstaVoca:=cmbVrstaVoca.Value, _
            sortaVoca:=cmbSortaVoca.Value, _
            ukupnoKol:=CDbl(txtUkupnoKgKlIIZbr.Value), _
            tipAmb:=cmbTipAmbZbr.Value, _
            ukupnoAmb:=0, _
            klasa:=KLASA_II)
        
        If resultII <> "" Then
            MsgBox "Zbirna sacuvana: " & result & " + " & resultII, vbInformation, APP_NAME
        Else
            MsgBox "Klasa I sacuvana, greska pri Klasi II!", vbExclamation, APP_NAME
        End If
    Else
        MsgBox "Zbirna sacuvana: " & result, vbInformation, APP_NAME
    End If
    
    UpdateValidacija
    Exit Sub
EH:
    LogErr "frmDokumenta.btnUnosZbr"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' ZBIRNA VALIDIERUNG – Live-Update
' ============================================================

Private Sub txtBrojZbirne_AfterUpdate()
    ' BrojZbirne auch in Otpremnica-Feld setzen
    txtBrojZbirneOtp.Value = txtBrojZbirne.Value
    UpdateValidacija
End Sub

Private Function UpdateValidacija() As Boolean
    UpdateValidacija = False
    
    If txtBrojZbirne.Value = "" Then
        lblValidacijaKG.Caption = ""
        lblValidacijaAmb.Caption = ""
        Exit Function
    End If
    
    Dim inputKgKlI As Double, inputKgKlII As Double, inputAmb As Long
    If IsNumeric(txtUkupnoKGZbr.Value) Then inputKgKlI = CDbl(txtUkupnoKGZbr.Value)
    If chkDveKlaseZbr.Value Then
        If IsNumeric(txtUkupnoKgKlIIZbr.Value) Then inputKgKlII = CDbl(txtUkupnoKgKlIIZbr.Value)
    End If
    If IsNumeric(txtUkupnoAmbZbr.Value) Then inputAmb = CLng(txtUkupnoAmbZbr.Value)
    
    Dim val As Variant
    val = ValidateZbirnaPreUnosa(txtBrojZbirne.Value, inputKgKlI, inputKgKlII, inputAmb)
    
    ' val(0-3): KlI  |  val(4-7): KlII  |  val(8-10): Amb
    Dim sumaKgI As Double: sumaKgI = CDbl(val(0))
    Dim zbrKgI As Double: zbrKgI = CDbl(val(1))
    Dim razKgI As Double: razKgI = CDbl(val(2))
    Dim validKgI As Boolean: validKgI = CBool(val(3))
    
    Dim sumaKgII As Double: sumaKgII = CDbl(val(4))
    Dim zbrKgII As Double: zbrKgII = CDbl(val(5))
    Dim razKgII As Double: razKgII = CDbl(val(6))
    Dim validKgII As Boolean: validKgII = CBool(val(7))
    
    ' KlI Anzeige
    Dim kgCaption As String
    kgCaption = "Kl.I - Otp: " & Format$(sumaKgI, "#,##0.0") & " kg | " & _
                "Zbr: " & Format$(zbrKgI, "#,##0.0") & " kg | " & _
                "Raz: " & Format$(razKgI, "#,##0.0") & " kg"
    
    ' KlII hinzufügen wenn aktiv
    If chkDveKlaseZbr.Value Then
        kgCaption = kgCaption & "  ||  Kl.II - Otp: " & Format$(sumaKgII, "#,##0.0") & _
                    " | Zbr: " & Format$(zbrKgII, "#,##0.0") & _
                    " | Raz: " & Format$(razKgII, "#,##0.0")
    End If
    
    lblValidacijaKG.Caption = kgCaption
    
    ' Farbe: beide Klassen müssen stimmen
    Dim kgValid As Boolean
    If chkDveKlaseZbr.Value Then
        kgValid = validKgI And validKgII
    Else
        kgValid = validKgI
    End If
    
    If kgValid Then
        lblValidacijaKG.foreColor = CLR_SUCCESS()
    ElseIf zbrKgI = 0 Then
        lblValidacijaKG.foreColor = TXT_MUTED()
    Else
        lblValidacijaKG.foreColor = CLR_ERROR()
    End If
    
    ' Ambalaza
    Dim sumaAmb As Long: sumaAmb = CLng(val(8))
    Dim zbrAmb As Long: zbrAmb = CLng(val(9))
    Dim razAmb As Long: razAmb = CLng(val(10))
    
    lblValidacijaAmb.Caption = "Amb Otp: " & sumaAmb & " | " & _
                                "Amb Zbr: " & zbrAmb & " | " & _
                                "Raz: " & razAmb
    
    If razAmb = 0 And zbrAmb > 0 Then
        lblValidacijaAmb.foreColor = CLR_SUCCESS()
    ElseIf zbrAmb = 0 Then
        lblValidacijaAmb.foreColor = TXT_MUTED()
    Else
        lblValidacijaAmb.foreColor = CLR_ERROR()
    End If
    
    UpdateValidacija = (kgValid And (razAmb = 0) And (zbrAmb > 0))
End Function

' ============================================================
' PRIJEMNICA UNOS (Ulaz Kupci)
' ============================================================

Private Sub btnUnosPrij_Click()
    On Error GoTo EH
    If cmbKupac.Value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If txtBrojZbirnePrij.Value = "" Then
        MsgBox "Unesite broj zbirne!", vbExclamation, APP_NAME
        txtBrojZbirnePrij.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtKolicinaPrij.Value) Or val(txtKolicinaPrij.Value) <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtKolicinaPrij.SetFocus
        Exit Sub
    End If
    
    If chkDveKlasePrij.Value Then
        If Not IsNumeric(txtKolicinaKlIIPrij.Value) Or val(txtKolicinaKlIIPrij.Value) <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKlIIPrij.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtCenaKlIIPrij.Value) Or val(txtCenaKlIIPrij.Value) <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKlIIPrij.SetFocus
            Exit Sub
        End If
    End If
    
    Dim kupacID As String
    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbKupac.Value, "KupacID"))
    
    Dim cena As Double
    If IsNumeric(txtCenaPrij.Value) Then cena = CDbl(txtCenaPrij.Value)
    
    Dim kolAmb As Long
    If IsNumeric(txtKolAmbPrij.Value) Then kolAmb = CLng(txtKolAmbPrij.Value)
    
    Dim kolAmbVracena As Long
    If IsNumeric(txtKolAmbVracena.Value) Then kolAmbVracena = CLng(txtKolAmbVracena.Value)
    
    'Duplikat Check
    If txtBrojPrij.Value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_PRIJEMNICA, COL_PRJ_BROJ, txtBrojPrij.Value, COL_PRJ_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    
    ' --- Klasa I ---
    Dim result As String
    result = SavePrijemnica_TX( _
        datum:=CDate(txtDatum.Value), _
        kupacID:=kupacID, _
        vozacID:=ExtractIDFromDisplay(cmbVozac.Value), _
        brojPrij:=txtBrojPrij.Value, _
        brojZbirne:=txtBrojZbirnePrij.Value, _
        vrstaVoca:=cmbVrstaVoca.Value, _
        sortaVoca:=cmbSortaVoca.Value, _
        kolicina:=CDbl(txtKolicinaPrij.Value), _
        cena:=cena, _
        tipAmb:=cmbTipAmbPrij.Value, _
        kolAmb:=kolAmb, _
        kolAmbVracena:=kolAmbVracena, _
        klasa:=KLASA_I)
    
    If result = "" Then
        MsgBox "Greska!", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' --- Klasa II ---
    If chkDveKlasePrij.Value Then
        Dim resultII As String
        resultII = SavePrijemnica_TX( _
            datum:=CDate(txtDatum.Value), _
            kupacID:=kupacID, _
            vozacID:=ExtractIDFromDisplay(cmbVozac.Value), _
            brojPrij:=txtBrojPrij.Value, _
            brojZbirne:=txtBrojZbirnePrij.Value, _
            vrstaVoca:=cmbVrstaVoca.Value, _
            sortaVoca:=cmbSortaVoca.Value, _
            kolicina:=CDbl(txtKolicinaKlIIPrij.Value), _
            cena:=CDbl(txtCenaKlIIPrij.Value), _
            tipAmb:=cmbTipAmbPrij.Value, _
            kolAmb:=0, _
            kolAmbVracena:=0, _
            klasa:=KLASA_II)
        
        If resultII <> "" Then
            MsgBox "Prijemnica sacuvana: " & result & " + " & resultII, vbInformation, APP_NAME
        Else
            MsgBox "Klasa I sacuvana, greska pri Klasi II!", vbExclamation, APP_NAME
        End If
    Else
        MsgBox "Prijemnica sacuvana: " & result, vbInformation, APP_NAME
    End If
    
    UpdateManjak txtBrojZbirnePrij.Value
    ClearPrijemnicaFields
    Exit Sub
EH:
    LogErr "frmDokumenta.btnUnosPrij"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub
Private Sub ClearPrijemnicaFields()
    txtBrojPrij.Value = ""
    txtKolicinaPrij.Value = ""
    txtCenaPrij.Value = ""
    txtKolAmbPrij.Value = ""
    txtKolAmbVracena.Value = ""
    chkDveKlasePrij.Value = False
    DisableField txtKolicinaKlIIPrij
    DisableField txtCenaKlIIPrij
End Sub

Private Sub txtBrojZbirnePrij_AfterUpdate()
    If txtBrojZbirnePrij.Value <> "" Then
        UpdateManjak txtBrojZbirnePrij.Value
    End If
End Sub

Private Sub UpdateManjak(ByVal brojZbirne As String)

    Dim pendingKlI As Double, pendingKlII As Double
    If IsNumeric(txtKolicinaPrij.Value) Then pendingKlI = CDbl(txtKolicinaPrij.Value)
    If chkDveKlasePrij.Value Then
        If IsNumeric(txtKolicinaKlIIPrij.Value) Then pendingKlII = CDbl(txtKolicinaKlIIPrij.Value)
    End If
    
    Dim manjak As Variant
    manjak = CalculateManjakPreview(brojZbirne, pendingKlI, pendingKlII)

    If Not IsArray(manjak) Or UBound(manjak) < 3 Then Exit Sub
    
    ' manjak = Array(ZbirnaKg, PrijemnicaKg, ManjakKg, ManjakPct)
    Dim zbrKg As Double: zbrKg = CDbl(manjak(0))
    Dim prijKg As Double: prijKg = CDbl(manjak(1))
    Dim manjakKg As Double: manjakKg = CDbl(manjak(2))
    Dim manjakPct As Double: manjakPct = CDbl(manjak(3))
    
    lblManjak.Caption = "Zbirna: " & Format$(zbrKg, "#,##0.0") & " kg | " & _
                         "Prijemnica: " & Format$(prijKg, "#,##0.0") & " kg | " & _
                         "Manjak: " & Format$(manjakKg, "#,##0.0") & " kg (" & _
                         Format$(manjakPct, "#,##0.00") & "%)"
    
    If Abs(manjakPct) < 0.5 Then
        lblManjak.foreColor = CLR_SUCCESS()
    ElseIf Abs(manjakPct) < 2 Then
        lblManjak.foreColor = CLR_WARNING()
    Else
        lblManjak.foreColor = CLR_ERROR()
    End If
    
    ' Prosek Gajbe
    Dim prosek As Double
    prosek = CalculateProsekGajbeByZbirna(brojZbirne)
    If prosek > 0 Then
        lblProsekGajbe.Caption = "Prosek gajbe: " & Format$(prosek, "#,##0.00") & " kg"
    Else
        lblProsekGajbe.Caption = ""
    End If
End Sub

' ============================================================
' IZLAZ KUPCI (Banka-Zahlung + Ambalaza)
' ============================================================

Private Sub btnUnosIzlaz_Click()
    On Error GoTo EH
    If cmbKupac.Value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim kupacNaziv As String
    kupacNaziv = cmbKupac.Value
    
    Dim kupacID As String
    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", kupacNaziv, "KupacID"))
    
    ' Duplikat-Check
    If txtBrojIzlaz.Value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_NOVAC, COL_NOV_BROJ_DOK, txtBrojIzlaz.Value, COL_NOV_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    
    ' FakturaID aus ComboBox
    Dim fakturaID As String
    If cmbFakturaIzlaz.Value <> "" Then
        ' BrojFakture aus ComboBox zu FakturaID nachschlagen
        Dim brojFakture As String
        brojFakture = Left$(cmbFakturaIzlaz.Value, InStr(cmbFakturaIzlaz.Value, " - ") - 1)
        fakturaID = CStr(LookupValue(TBL_FAKTURE, COL_FAK_BROJ, brojFakture, COL_FAK_ID))
    End If
    
    ' Novac
    Dim novac As Double
    If IsNumeric(txtNovacIzlaz.Value) Then novac = CDbl(txtNovacIzlaz.Value)
    
    If novac > 0 Then
        Dim tip As String
        If cmbFakturaIzlaz.Value <> "" Then
            tip = NOV_KUPCI_UPLATA
        Else
            tip = NOV_KUPCI_AVANS
        End If
        SaveNovac_TX txtBrojIzlaz.Value, CDate(txtDatum.Value), _
                     cmbKupac.Value, kupacID, "Kupac", "", _
                     "", fakturaID, "", tip, novac, 0, ""
        
        ' Faktura Status aktualisieren wenn voll bezahlt
        If fakturaID <> "" Then
            UpdateFakturaStatus fakturaID
        End If
    End If
    
    ' Ambalaza zurück vom Kunden
    Dim kolAmbIzlaz As Long
    If IsNumeric(txtKolAmbIzlaz.Value) Then kolAmbIzlaz = CLng(txtKolAmbIzlaz.Value)
    
    If kolAmbIzlaz > 0 Then
        Dim vozacID As String
        If cmbVozac.Value <> "" Then vozacID = ExtractIDFromDisplay(cmbVozac.Value)
        
        TrackAmbalaza CDate(txtDatum.Value), cmbTipAmbIzlaz.Value, kolAmbIzlaz, "Ulaz", kupacID, "Kupac", vozacID, txtBrojIzlaz.Value, DOK_TIP_IZLAZ_KUPCI
    End If
    
    
    MsgBox "Sacuvano!", vbInformation, APP_NAME
    txtBrojIzlaz.Value = ""
    txtNovacIzlaz.Value = ""
    txtKolAmbIzlaz.Value = ""
    cmbFakturaIzlaz.Value = ""
    FillOpenFakture    ' ? fehlt
    Exit Sub
EH:
    LogErr "frmDokumenta.btnUnosIzlaz"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub FillOpenFakture()
    cmbFakturaIzlaz.Clear
    If cmbKupac.Value = "" Then Exit Sub
    
    Dim kupacID As String
    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbKupac.Value, "KupacID"))
    
    Dim fakture As Variant
    fakture = GetOpenFakture(kupacID)
    If IsEmpty(fakture) Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(fakture, 1)
        cmbFakturaIzlaz.AddItem CStr(fakture(i, 1)) & " - " & _
            Format$(CDbl(fakture(i, 5)), "#,##0.00") & " od " & _
            Format$(CDbl(fakture(i, 3)), "#,##0.00") & " RSD"
    Next i
End Sub

' ============================================================
' STORNO-BEREICH in frmDokumenta
' Eine TextBox + ComboBox + Button
' ============================================================

Private Sub btnStorno_Click()
    On Error GoTo EH
    Dim tipDok As String
    tipDok = cmbStornoDokument.Value
    
    Dim brDok As String
    brDok = Trim$(txtStornoBroj.Value)
    
    If tipDok = "" Then
        MsgBox "Izaberite tip dokumenta!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If brDok = "" Then
        MsgBox "Unesite broj dokumenta!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim Success As Boolean
    
    Select Case tipDok
        Case "Otkup"
            ' Otkup hat OtkupID, nicht BrojOtp
            Dim otkupID As String
            otkupID = LookupActiveID(TBL_OTKUP, COL_OTK_BR_DOK, brDok, COL_OTK_ID)
            If otkupID = "" Then
                MsgBox "Otkup '" & brDok & "' nije pronadjen!", vbExclamation, APP_NAME
                Exit Sub
            End If
            If ConfirmStorno("otkup", brDok) Then Success = StornoOtkup_TX(otkupID)
            
        Case "Otpremnica"
            Dim otpID As String
            otpID = LookupActiveID(TBL_OTPREMNICA, COL_OTP_BROJ, brDok, COL_OTP_ID)
            If otpID = "" Then
                MsgBox "Otpremnica '" & brDok & "' nije pronadjena!", vbExclamation, APP_NAME
                Exit Sub
            End If
            If ConfirmStorno("otpremnicu", brDok) Then Success = StornoOtpremnica_TX(otpID)
            
        Case "Zbirna"
            If ConfirmStorno("zbirnu", brDok) Then Success = StornoZbirna_TX(brDok)
            
        Case "Prijemnica"
            Dim prijID As String
            prijID = LookupActiveID(TBL_PRIJEMNICA, COL_PRJ_BROJ, brDok, COL_PRJ_ID)
            If prijID = "" Then
                MsgBox "Prijemnica '" & brDok & "' nije pronadjena!", vbExclamation, APP_NAME
                Exit Sub
            End If
            If ConfirmStorno("prijemnicu", brDok) Then Success = StornoPrijemnica_TX(prijID)
            
        Case "Faktura"
            Dim fakID As String
            fakID = LookupActiveID(TBL_FAKTURE, COL_FAK_BROJ, brDok, COL_FAK_ID)
            If fakID = "" Then
                MsgBox "Faktura '" & brDok & "' nije pronadjena!", vbExclamation, APP_NAME
                Exit Sub
            End If
            If ConfirmStorno("fakturu", brDok) Then Success = StornoFaktura_TX(fakID)
            
        Case "Novac"
            If ConfirmStorno("novac stavku", brDok) Then Success = StornoNovac_TX(brDok)
    End Select
    
    If Success Then
        MsgBox "Stornirano!", vbInformation, APP_NAME
        txtStornoBroj.Value = ""
        CheckVerwaisteDokumente
    End If
    Exit Sub
EH:
    LogErr "frmDokumenta.btnStorno"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Function ConfirmStorno(ByVal tipText As String, ByVal broj As String) As Boolean
    ConfirmStorno = (MsgBox("Stornirati " & tipText & " " & broj & "?", _
                            vbQuestion + vbYesNo, APP_NAME) = vbYes)
End Function

Private Sub CheckVerwaisteDokumente()
    Dim verwOtp As Variant
    verwOtp = GetVerwaisteDokumente("Otpremnica")
    
    Dim verwPrij As Variant
    verwPrij = GetVerwaisteDokumente("Prijemnica")
    
    If IsEmpty(verwOtp) And IsEmpty(verwPrij) Then
        lblStornoWarning.Visible = False
        Exit Sub
    End If
    
    Dim msg As String
    Dim i As Long
    
    If Not IsEmpty(verwOtp) Then
        msg = UBound(verwOtp, 1) & " otp. ceka novu zbirnu: "
        For i = 1 To UBound(verwOtp, 1)
            If i > 1 Then msg = msg & ", "
            msg = msg & CStr(verwOtp(i, 2))
        Next i
        msg = msg & vbCrLf
    End If
    
    If Not IsEmpty(verwPrij) Then
        msg = msg & UBound(verwPrij, 1) & " prij. ceka novu zbirnu: "
        For i = 1 To UBound(verwPrij, 1)
            If i > 1 Then msg = msg & ", "
            msg = msg & CStr(verwPrij(i, 2))
        Next i
    End If
    
    lblStornoWarning.Caption = msg
    lblStornoWarning.foreColor = CLR_WARNING()
    lblStornoWarning.Visible = True
End Sub

' ============================================================
' NAVIGATION
' ============================================================

Private Sub btnPovratak_Click()
    Me.Hide
    frmOtkupAPP.Show
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnUnosOtp, "Unos otpremnice"
    StylePrimaryButton btnUnosZbr, "Unos zbirne"
    StylePrimaryButton btnUnosPrij, "Unos prijemnice"
    StylePrimaryButton btnUnosOMUlaz, "Unos OM ulaz"
    StylePrimaryButton btnUnosIzlaz, "Unos izlaz"
    StyleExitButton btnPovratak, "Povratak"
    StyleExitButton btnStorno, "Storno"
End Sub

Private Sub btnUnosOtp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnosOtp
End Sub

Private Sub btnUnosZbr_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnosZbr
End Sub

Private Sub btnUnosPrij_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnosPrij
End Sub

Private Sub btnUnosOMUlaz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnosOMUlaz
End Sub

Private Sub btnUnosIzlaz_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnUnosIzlaz
End Sub

Private Sub btnStorno_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnStorno
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        frmOtkupAPP.Show
    End If
End Sub


Private Sub FillOpenOtkupi()
    cmbOtkupBlok.Clear
    Erase m_OtkupIDs
    
    If cmbPrimalacOMUlaz.Value = "" Then Exit Sub
    
    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbPrimalacOMUlaz.Value)
    
    Dim otkupi As Variant
    otkupi = GetOpenOtkupi(kooperantID)
    If IsEmpty(otkupi) Then Exit Sub
    
    ReDim m_OtkupIDs(0 To UBound(otkupi, 1) - 1)
    
    Dim i As Long
    For i = 1 To UBound(otkupi, 1)
        m_OtkupIDs(i - 1) = CStr(otkupi(i, 2))
        cmbOtkupBlok.AddItem CStr(otkupi(i, 1)) & " - " & _
            Format$(CDbl(otkupi(i, 5)), "#,##0.00") & " od " & _
            Format$(CDbl(otkupi(i, 3)), "#,##0.00") & " RSD"
    Next i
End Sub
