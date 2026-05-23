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

Private m_FocusableInputs As Collection

Private mChromeRemoved As Boolean

Private Sub UserForm_Activate()
    On Error GoTo EH

    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
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

    ' important frames — section headers (UPPERCASE + gold)
    StyleSectionHeader fraOtpremnica, "Izlaz OM  (Otpremnica)"
    StyleSectionHeader fraZbirna, "Zbirna otpremnica"
    StyleSectionHeader fraPrijemnica, "Ulaz Kupci  (Prijemnica)"
    StyleSectionHeader fraOMUlaz, "Ulaz OM  (Novac kooperantu)"
    StyleSectionHeader fraIzlazKupci, "Izlaz Kupci  (Novac od kupca)"
    StyleSectionHeader fraStorno, "Storno"
        
    txtDatum.value = Format$(Date, "d.m.yyyy")
    
    FillCmb cmbVrstaVoca, GetLookupList(TBL_KULTURE, "VrstaVoca")
    FillComboDisplayID cmbOtkupnoMesto, TBL_STANICE, "Naziv", "StanicaID"
    FillComboDisplayID cmbKupac, TBL_KUPCI, COL_KUP_NAZIV, COL_KUP_ID
    FillCmb cmbVozac, GetVozacDisplayList()
    FillCmb cmbTipAmbOtp, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbZbr, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbPrij, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbOMUlaz, Array(AMB_12_1, AMB_6_1)
    FillCmb cmbTipAmbIzlaz, Array(AMB_12_1, AMB_6_1)
    
    lblValidacijaKG.caption = ""
    lblValidacijaAmb.caption = ""
    lblManjak.caption = ""
    
    ' Druga klasa:
    DisableField txtKolicinaKlIIOtp
    DisableField txtCenaKlIIOtp
    chkDveKlaseOtp.value = False
    lblUkupnoKgOtp.caption = ""
    
    DisableField txtUkupnoKgKlIIZbr
    chkDveKlaseZbr.value = False
    lblUkupnoKgZbr.caption = ""
    
    DisableField txtKolicinaKlIIPrij
    DisableField txtCenaKlIIPrij
    chkDveKlasePrij.value = False
    lblUkupnoKgPrij.caption = ""
    lblProsekGajbe.caption = ""
    
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

    ' v6.11 UI: register F-key shortcut targets
    SetupFkeyAccelerators
    SetupFkeyHints

    ' v6.11 UI: enable form-level KeyDown for F-keys
    'Me.KeyPreview = True
    
    ' v6.11 UI: subtitle, KPI traka, akcent linije
    StyleSubtitle lblSubtitle, "Operativni unos dokumenata, ambalaže i novca"
    StyleFrameTitleLabel lblKopf, "Osnovni podaci"
    LayoutTopKpis
    RefreshTopKpis

    StyleSectionAccent lblAccentOtp, fraOtpremnica, "primary"
    StyleSectionAccent lblAccentZbr, fraZbirna, "info"
    StyleSectionAccent lblAccentPrij, fraPrijemnica, "primary"
    StyleSectionAccent lblAccentOMUlaz, fraOMUlaz, "primary"
    StyleSectionAccent lblAccentIzlaz, fraIzlazKupci, "primary"
    StyleSectionAccent lblAccentStorno, fraStorno, "warn"
    
    ' Frame captions empty — naslovi idu kao zasebni Label-i ispod akcent linije
    fraOtpremnica.caption = ""
    fraZbirna.caption = ""
    fraPrijemnica.caption = ""
    fraOMUlaz.caption = ""
    fraIzlazKupci.caption = ""
    fraStorno.caption = ""

    StyleFrameTitleLabel lblTitleOtp, "Izlaz OM  (Otpremnica)"
    StyleFrameTitleLabel lblTitleZbr, "Zbirna otpremnica"
    StyleFrameTitleLabel lblTitlePrij, "Ulaz Kupci  (Prijemnica)"
    StyleFrameTitleLabel lblTitleOMUlaz, "Ulaz OM  (Novac kooperantu)"
    StyleFrameTitleLabel lblTitleIzlaz, "Izlaz Kupci  (Novac od kupca)"
    StyleFrameTitleLabel lblTitleStorno, "Storno"
    Exit Sub

EH:
    LogErr "frmDokumenta.UserForm_Activate"
    MsgBox "Greška pri otvaranju dokumenata: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub txtCenaOtp_Change():       UpdateUkupnoKgOtp: End Sub
Private Sub txtCenaKlIIOtp_Change():   UpdateUkupnoKgOtp: End Sub

Private Sub cmbFakturaIzlaz_Change()
    On Error GoTo EH

    If cmbFakturaIzlaz.value = "" Then
        lblManjak.caption = ""        ' kupacSaldo dummy — ako imaš zaseban label, zameni
        Exit Sub
    End If

    Dim fakturaID As String
    fakturaID = GetComboID(cmbFakturaIzlaz)

    If fakturaID = "" Then Exit Sub

    Dim iznos As Double
    Dim uplaceno As Double
    Dim preostalo As Double

    If TryParseDouble(NzToText(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS)), iznos) Then
        uplaceno = GetUplataForFaktura(fakturaID)
        preostalo = iznos - uplaceno

        ' Koristi lblOMAvansSaldo kao multipurpose status bar — ili dodaj novi lblSaldoKupca
        lblOMAvansSaldo.caption = "SALDO KUPCA   Uplaceno: " & Format$(uplaceno, "#,##0") & _
                                  "   Ostatak: " & Format$(preostalo, "#,##0") & " RSD"
        If preostalo <= 0 Then
            lblOMAvansSaldo.ForeColor = CLR_SUCCESS()
        ElseIf uplaceno > 0 Then
            lblOMAvansSaldo.ForeColor = CLR_WARNING()
        Else
            lblOMAvansSaldo.ForeColor = TXT_MUTED()
        End If
    End If

    Exit Sub

EH:
    LogErr "frmDokumenta.cmbFakturaIzlaz_Change"
End Sub

Private Sub chkDveKlaseOtp_Click()
    If chkDveKlaseOtp.value Then
        EnableField txtKolicinaKlIIOtp
        EnableField txtCenaKlIIOtp
        StyleTextBox txtKolicinaKlIIOtp
        StyleTextBox txtCenaKlIIOtp
    Else
        DisableField txtKolicinaKlIIOtp
        DisableField txtCenaKlIIOtp
        lblUkupnoKgOtp.caption = ""
    End If
End Sub

Private Sub chkDveKlaseZbr_Click()
    If chkDveKlaseZbr.value Then
        EnableField txtUkupnoKgKlIIZbr
        StyleTextBox txtUkupnoKgKlIIZbr
    Else
        DisableField txtUkupnoKgKlIIZbr
    End If
End Sub

Private Sub chkDveKlasePrij_Click()
    If chkDveKlasePrij.value Then
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
    Dim kl1 As Double, kl2 As Double, cena1 As Double, cena2 As Double
    Dim ukupnoKg As Double, ukupnoRsd As Double

    If IsNumeric(txtKolicinaOtp.value) Then kl1 = CDbl(txtKolicinaOtp.value)
    If IsNumeric(txtCenaOtp.value) Then cena1 = CDbl(txtCenaOtp.value)

    If chkDveKlaseOtp.value Then
        If IsNumeric(txtKolicinaKlIIOtp.value) Then kl2 = CDbl(txtKolicinaKlIIOtp.value)
        If IsNumeric(txtCenaKlIIOtp.value) Then cena2 = CDbl(txtCenaKlIIOtp.value)
    End If

    ukupnoKg = kl1 + kl2
    ukupnoRsd = (kl1 * cena1) + (kl2 * cena2)

    If ukupnoKg <= 0 And ukupnoRsd <= 0 Then
        lblUkupnoKgOtp.caption = ""
        Exit Sub
    End If

    If chkDveKlaseOtp.value Then
        lblUkupnoKgOtp.caption = "UKUPNO  " & Format$(ukupnoKg, "#,##0.0") & " kg"
        If ukupnoRsd > 0 Then
            lblUkupnoKgOtp.caption = lblUkupnoKgOtp.caption & "  •  " & Format$(ukupnoRsd, "#,##0") & " RSD"
        End If
    Else
        If ukupnoRsd > 0 Then
            lblUkupnoKgOtp.caption = "UKUPNO  " & Format$(ukupnoRsd, "#,##0") & " RSD"
        Else
            lblUkupnoKgOtp.caption = ""
        End If
    End If

    StylePreviewBox lblUkupnoKgOtp, "ok"
End Sub

Private Sub txtKolicinaPrij_Change()
    If txtBrojZbirnePrij.value <> "" Then
        UpdateManjak txtBrojZbirnePrij.value
    End If
    txtKolAmbPrij_Change
    UpdateUkupnoKgPrij        ' <-- DODATO
End Sub

Private Sub txtKolicinaKlIIPrij_Change()
    If txtBrojZbirnePrij.value <> "" Then
        UpdateManjak txtBrojZbirnePrij.value
    End If
    UpdateUkupnoKgPrij        ' <-- DODATO
End Sub

Private Sub txtCenaPrij_Change():       UpdateUkupnoKgPrij: End Sub
Private Sub txtCenaKlIIPrij_Change():   UpdateUkupnoKgPrij: End Sub

Private Sub txtKolAmbPrij_Change()
    If IsNumeric(txtKolicinaPrij.value) And IsNumeric(txtKolAmbPrij.value) Then
        If CLng(txtKolAmbPrij.value) > 0 Then
            lblProsekGajbe.caption = "Prosek gajbe: " & _
                Format$(CDbl(txtKolicinaPrij.value) / CLng(txtKolAmbPrij.value), "#,##0.00") & " kg"
        End If
    Else
        lblProsekGajbe.caption = ""
    End If
End Sub

Private Sub cmbOtkupnoMesto_Change()
    On Error GoTo EH

    Dim stanicaID As String

    If cmbOtkupnoMesto.value = "" Then
        cmbPrimalacOMUlaz.Clear
        lblOMAvansSaldo.caption = ""
        txtBrojOtp.value = ""   ' ? DODATO: ocisti predlog otpremnice
        Exit Sub
    End If

    stanicaID = GetComboID(cmbOtkupnoMesto)

    If stanicaID = "" Then
        cmbPrimalacOMUlaz.Clear
        lblOMAvansSaldo.caption = ""
        Exit Sub
    End If

    FillComboKooperantiByStanica cmbPrimalacOMUlaz, stanicaID
    UpdateOMAvansSaldo
    RefreshBrojOtpSuggestion
    
    Exit Sub

EH:
    LogErr "frmDokumenta.cmbOtkupnoMesto_Change"
End Sub

Private Sub cmbVozac_Change()
    On Error GoTo EH

    If cmbVozac.value = "" Then
        txtBrojZbirne.value = ""
        txtBrojZbirneOtp.value = ""   ' postojeca sinhronizacija (linija 870)
        Exit Sub
    End If

    RefreshBrojZbirneSuggestion
    Exit Sub

EH:
    LogErr "frmDokumenta.cmbVozac_Change"
End Sub

Private Sub txtDatum_AfterUpdate()
    ' AfterUpdate puca samo na commit (Tab/Enter/focus exit), ne na svaki keystroke
    RefreshBrojOtpSuggestion
    RefreshBrojZbirneSuggestion
End Sub
' ============================================================
' KASKADIERUNG
' ============================================================

Private Sub cmbVrstaVoca_Change()
    cmbSortaVoca.Clear
    If cmbVrstaVoca.value <> "" Then
        FillCmb cmbSortaVoca, _
            GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", cmbVrstaVoca.value)
    End If
End Sub

Private Sub cmbKupac_Change()
    On Error GoTo EH

    cmbHladnjaca.Clear
    cmbPogon.Clear

    If cmbKupac.value <> "" Then
        Dim kupacID As String
        kupacID = GetComboID(cmbKupac)

        If kupacID <> "" Then
            Dim hlad As String
            hlad = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, "Hladnjaca"))

            If hlad <> "" Then cmbHladnjaca.AddItem hlad
        End If
    End If

    FillOpenFakture
    Exit Sub

EH:
    LogErr "frmDokumenta.cmbKupac_Change"
End Sub

' ============================================================
' OTPREMNICA UNOS (Izlaz OM)
' ============================================================

Private Sub btnUnosOtp_Click()
    On Error GoTo EH

    ' Validacija
    If cmbOtkupnoMesto.value = "" Then
        MsgBox "Izaberite otkupno mesto!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If cmbVozac.value = "" Then
        MsgBox "Izaberite vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If Not IsNumeric(txtKolicinaOtp.value) Or val(txtKolicinaOtp.value) <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtKolicinaOtp.SetFocus
        Exit Sub
    End If

    If chkDveKlaseOtp.value Then
        If Not IsNumeric(txtKolicinaKlIIOtp.value) Or val(txtKolicinaKlIIOtp.value) <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKlIIOtp.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(txtCenaKlIIOtp.value) Or val(txtCenaKlIIOtp.value) <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKlIIOtp.SetFocus
            Exit Sub
        End If
    End If

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        MsgBox "Unesite ispravan datum!", vbExclamation, APP_NAME
        txtDatum.SetFocus
        Exit Sub
    End If

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)

    If stanicaID = "" Then
        MsgBox "Nije pronaden ID otkupnog mesta!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim vozacID As String
    vozacID = ExtractIDFromDisplay(cmbVozac.value)

    If vozacID = "" Then
        MsgBox "Nije pronaden ID vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim kolicinaI As Double
    If Not TryParseDouble(txtKolicinaOtp.value, kolicinaI) Or kolicinaI <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtKolicinaOtp.SetFocus
        Exit Sub
    End If

    Dim cenaI As Double
    If txtCenaOtp.value <> "" Then
        If Not TryParseDouble(txtCenaOtp.value, cenaI) Or cenaI < 0 Then
            MsgBox "Unesite ispravnu cenu!", vbExclamation, APP_NAME
            txtCenaOtp.SetFocus
            Exit Sub
        End If
    End If

    Dim kolAmb As Long
    If txtKolAmbOtp.value <> "" Then
        If Not TryParseLong(txtKolAmbOtp.value, kolAmb) Then
            MsgBox "Unesite ispravnu kolicinu ambalaže!", vbExclamation, APP_NAME
            txtKolAmbOtp.SetFocus
            Exit Sub
        End If
    End If

    ' Duplikat check
    If txtBrojOtp.value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_OTPREMNICA, COL_OTP_BROJ, txtBrojOtp.value, COL_OTP_DATUM)

        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    ' Klasa II vrednosti pripremiti bez IIf,
    ' jer VBA IIf evaluira obe grane.
    Dim kolicinaII As Double
    Dim cenaII As Double

    If chkDveKlaseOtp.value Then
        If Not TryParseDouble(txtKolicinaKlIIOtp.value, kolicinaII) Or kolicinaII <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKlIIOtp.SetFocus
            Exit Sub
        End If

        If Not TryParseDouble(txtCenaKlIIOtp.value, cenaII) Or cenaII <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKlIIOtp.SetFocus
            Exit Sub
        End If
    End If

    ' Atomicni save kroz modDokumenta wrapper
    Dim result As String

    result = SaveOtpremnicaMulti_TX( _
        datum:=datumDok, _
        stanicaID:=stanicaID, _
        vozacID:=vozacID, _
        brojOtp:=txtBrojOtp.value, _
        brojZbirne:=txtBrojZbirneOtp.value, _
        vrsta:=cmbVrstaVoca.value, _
        sorta:=cmbSortaVoca.value, _
        kolicinaI:=kolicinaI, _
        cenaI:=cenaI, _
        tipAmb:=cmbTipAmbOtp.value, _
        kolAmb:=kolAmb, _
        hasKlasaII:=chkDveKlaseOtp.value, _
        kolicinaII:=kolicinaII, _
        cenaII:=cenaII)

    If result = "" Then
        MsgBox "Greška pri cuvanju otpremnice. Promene su vracene.", vbCritical, APP_NAME
        Exit Sub
    End If

    MsgBox "Otpremnica sacuvana: " & result, vbInformation, APP_NAME

    ClearOtpremnicaFields

    If txtBrojZbirne.value <> "" Then
        UpdateValidacija
    End If

    Exit Sub

EH:
    LogErr "frmDokumenta.btnUnosOtp"
    MsgBox "Greška: " & Err.description, vbCritical, APP_NAME
End Sub
Private Sub ClearOtpremnicaFields()
    txtBrojOtp.value = ""
    txtKolicinaOtp.value = ""
    txtCenaOtp.value = ""
    txtKolAmbOtp.value = ""
    chkDveKlaseOtp.value = False
    DisableField txtKolicinaKlIIOtp
    DisableField txtCenaKlIIOtp
    lblUkupnoKgOtp.caption = ""
    
    RefreshBrojOtpSuggestion   ' ? DODATO: predloži sledeci za istu stanicu/datum
End Sub

Private Sub RefreshBrojOtpSuggestion()
    ' Predlaže BrojOtpremnice. OTP je VBA-only — scan samo tblOtpremnica.
    On Error GoTo EH

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)
    If Len(stanicaID) = 0 Then Exit Sub

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then Exit Sub

    Dim suggested As String
    suggested = SuggestNextBroj(KIND_OTP, stanicaID, datumDok)

    If Len(suggested) > 0 Then
        txtBrojOtp.value = suggested
    End If
    Exit Sub

EH:
    LogErr "frmDokumenta.RefreshBrojOtpSuggestion"
End Sub

Private Sub RefreshBrojZbirneSuggestion()
    ' Predlaže BrojZbirne. ZBR scan = tblZbirna + VOZ-{vozacID} sheet.
    On Error GoTo EH

    If cmbVozac.value = "" Then Exit Sub

    Dim vozacID As String
    vozacID = ExtractIDFromDisplay(cmbVozac.value)
    If Len(vozacID) = 0 Then Exit Sub

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then Exit Sub

    Dim suggested As String
    suggested = SuggestNextBroj(KIND_ZBR, vozacID, datumDok)

    If Len(suggested) > 0 Then
        txtBrojZbirne.value = suggested
        ' Sinhronizuj u OTP sekciji (postojeca txtBrojZbirne_AfterUpdate logika
        ' na liniji 868 puca samo na manual edit; ovde radimo programatsko
        ' setovanje, _AfterUpdate ne puca, pa eksplicitno sinhronizujemo).
        txtBrojZbirneOtp.value = suggested
    End If
    Exit Sub

EH:
    LogErr "frmDokumenta.RefreshBrojZbirneSuggestion"
End Sub

Public Function SaveOMUlaz_TX(ByVal datum As Date, _
                              ByVal brojDok As String, _
                              ByVal stanicaNaziv As String, _
                              ByVal stanicaID As String, _
                              ByVal vozacID As String, _
                              ByVal tipAmb As String, _
                              ByVal kolAmb As Long, _
                              ByVal vrstaVoca As String, _
                              ByVal novac As Double, _
                              ByVal kooperantID As String, _
                              ByVal primalacDisplay As String, _
                              ByVal otkupID As String, _
                              ByVal tipNovca As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If kolAmb <= 0 And novac <= 0 Then
        Err.Raise vbObjectError + 1501, "SaveOMUlaz_TX", _
                  "Nema ambalaže ni novca za cuvanje."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP

    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, _
                      "Ulaz", stanicaID, "Stanica", _
                      vozacID, brojDok, DOK_TIP_OM_ULAZ
    End If

    If novac > 0 Then
        Dim novacID As String

        novacID = SaveNovac( _
            brojDok:=brojDok, _
            datum:=datum, _
            partner:=stanicaNaziv, _
            partnerId:=stanicaID, _
            entitetTip:="OM", _
            omID:=stanicaID, _
            kooperantID:=kooperantID, _
            fakturaID:="", _
            vrstaVoca:=vrstaVoca, _
            tip:=tipNovca, _
            uplata:=0, _
            isplata:=novac, _
            napomena:=primalacDisplay, _
            otkupID:=otkupID)

        If novacID = "" Then
            Err.Raise vbObjectError + 1502, "SaveOMUlaz_TX", _
                      "SaveNovac fehlgeschlagen"
        End If

        If otkupID <> "" Then
            UpdateOtkupStatus otkupID
        End If
    End If

    tx.CommitTx

    Set tx = Nothing

    SaveOMUlaz_TX = True
    Exit Function

EH:
    LogErr "SaveOMUlaz_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOMUlaz_TX = False
End Function

Private Sub btnUnosOMUlaz_Click()
    On Error GoTo EH

    If cmbOtkupnoMesto.value = "" Then
        MsgBox "Izaberite otkupno mesto!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        MsgBox "Unesite ispravan datum!", vbExclamation, APP_NAME
        txtDatum.SetFocus
        Exit Sub
    End If

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)

    If stanicaID = "" Then
        MsgBox "Nije pronaden ID otkupnog mesta!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim vozacID As String
    If cmbVozac.value <> "" Then
        vozacID = ExtractIDFromDisplay(cmbVozac.value)
    End If

    Dim brojDok As String
    brojDok = Trim$(txtBrojDokOMUlaz.value)

    If brojDok <> "" Then
        Dim dupMsg As String

        dupMsg = CheckDuplicate(TBL_AMBALAZA, COL_AMB_DOK_ID, brojDok, COL_AMB_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If

        dupMsg = CheckDuplicate(TBL_NOVAC, COL_NOV_BROJ_DOK, brojDok, COL_NOV_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    Dim kolAmb As Long
    If Trim$(txtKolAmbOMUlaz.value) <> "" Then
        If Not TryParseLong(txtKolAmbOMUlaz.value, kolAmb) Then
            MsgBox "Unesite ispravnu kolicinu ambalaže!", vbExclamation, APP_NAME
            txtKolAmbOMUlaz.SetFocus
            Exit Sub
        End If
    End If

    Dim novac As Double
    If Trim$(txtNovacOMUlaz.value) <> "" Then
        If Not TryParseDouble(txtNovacOMUlaz.value, novac) Or novac < 0 Then
            MsgBox "Unesite ispravan iznos novca!", vbExclamation, APP_NAME
            txtNovacOMUlaz.SetFocus
            Exit Sub
        End If
    End If

    If kolAmb <= 0 And novac <= 0 Then
        MsgBox "Unesite ambalažu ili novac za OM ulaz.", vbExclamation, APP_NAME
        Exit Sub
    End If

    If kolAmb > 0 And cmbTipAmbOMUlaz.value = "" Then
        MsgBox "Izaberite tip ambalaže!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim kooperantID As String
    Dim otkupID As String
    Dim tipNovca As String

    If novac > 0 Then
        If cmbPrimalacOMUlaz.value <> "" Then
            kooperantID = ExtractIDFromDisplay(cmbPrimalacOMUlaz.value)

            If kooperantID = "" Then
                MsgBox "Nije pronaden ID primaoca.", vbExclamation, APP_NAME
                Exit Sub
            End If

            If cmbOtkupBlok.ListIndex >= 0 Then
                otkupID = m_OtkupIDs(cmbOtkupBlok.ListIndex)

                Dim otkupi As Variant
                otkupi = GetOpenOtkupi(kooperantID)

                If Not IsEmpty(otkupi) Then
                    Dim preostalo As Double
                    If TryParseDouble(NzToText(otkupi(cmbOtkupBlok.ListIndex + 1, 5)), preostalo) Then
                        If novac > preostalo Then
                            MsgBox "Iznos (" & Format$(novac, "#,##0") & _
                                   ") veci od preostalog (" & _
                                   Format$(preostalo, "#,##0") & ")!", _
                                   vbExclamation, APP_NAME
                            Exit Sub
                        End If
                    End If
                End If

                If tglIzOMAvansa.value Then
                    Dim omSaldo As Double
                    omSaldo = GetOMAvansSaldo(stanicaID)

                    If novac > omSaldo Then
                        MsgBox "Nedovoljno OM avansa! Raspoloživo: " & _
                               Format$(omSaldo, "#,##0") & " RSD", _
                               vbExclamation, APP_NAME
                        Exit Sub
                    End If

                    tipNovca = NOV_KES_OTKUPAC_KOOP
                Else
                    tipNovca = NOV_VIRMAN_FIRMA_KOOP
                End If
            Else
                tipNovca = NOV_VIRMAN_AVANS_KOOP
            End If
        Else
            kooperantID = ""
            otkupID = ""
            tipNovca = NOV_KES_FIRMA_OTKUPAC
        End If
    End If

    If Not SaveOMUlaz_TX( _
        datum:=datumDok, _
        brojDok:=brojDok, _
        stanicaNaziv:=cmbOtkupnoMesto.value, _
        stanicaID:=stanicaID, _
        vozacID:=vozacID, _
        tipAmb:=cmbTipAmbOMUlaz.value, _
        kolAmb:=kolAmb, _
        vrstaVoca:=cmbVrstaVoca.value, _
        novac:=novac, _
        kooperantID:=kooperantID, _
        primalacDisplay:=cmbPrimalacOMUlaz.value, _
        otkupID:=otkupID, _
        tipNovca:=tipNovca) Then

        MsgBox "Greška pri cuvanju OM ulaza. Promene su vracene.", vbCritical, APP_NAME
        Exit Sub
    End If

    UpdateOMAvansSaldo

    MsgBox "Sacuvano!", vbInformation, APP_NAME

    txtBrojDokOMUlaz.value = ""
    txtKolAmbOMUlaz.value = ""
    txtNovacOMUlaz.value = ""
    cmbPrimalacOMUlaz.value = ""
    cmbOtkupBlok.Clear
    tglIzOMAvansa.value = False

    Exit Sub

EH:
    LogErr "frmDokumenta.btnUnosOMUlaz"
    MsgBox "Greška: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub cmbPrimalacOMUlaz_Change()
    cmbOtkupBlok.Clear
    tglIzOMAvansa.enabled = (cmbPrimalacOMUlaz.value <> "")
    
    If cmbPrimalacOMUlaz.value = "" Then
        tglIzOMAvansa.value = False
        Exit Sub
    End If
    
    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbPrimalacOMUlaz.value)
    FillOpenOtkupi
End Sub

Private Sub UpdateOMAvansSaldo()
    On Error GoTo EH

    If cmbOtkupnoMesto.value = "" Then
        lblOMAvansSaldo.caption = ""
        Exit Sub
    End If

    Dim stanicaID As String
    stanicaID = GetComboID(cmbOtkupnoMesto)

    If stanicaID = "" Then
        lblOMAvansSaldo.caption = ""
        Exit Sub
    End If

    Dim saldo As Double
    saldo = GetOMAvansSaldo(stanicaID)

    If saldo > 0 Then
        lblOMAvansSaldo.caption = "OM Avans: " & Format$(saldo, "#,##0") & " RSD"
    Else
        lblOMAvansSaldo.caption = ""
    End If

    Exit Sub

EH:
    LogErr "frmDokumenta.UpdateOMAvansSaldo"
    lblOMAvansSaldo.caption = ""
End Sub

' ============================================================
' ZBIRNA UNOS
' ============================================================

Private Sub btnUnosZbr_Click()
    On Error GoTo EH

    If cmbVozac.value = "" Then
        MsgBox "Izaberite vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If cmbKupac.value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If Trim$(txtBrojZbirne.value) = "" Then
        MsgBox "Unesite broj zbirne!", vbExclamation, APP_NAME
        txtBrojZbirne.SetFocus
        Exit Sub
    End If
    
    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        MsgBox "Unesite ispravan datum!", vbExclamation, APP_NAME
        txtDatum.SetFocus
        Exit Sub
    End If

    Dim ukupnoKolI As Double
    If Not TryParseDouble(txtUkupnoKGZbr.value, ukupnoKolI) Or ukupnoKolI <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtUkupnoKGZbr.SetFocus
        Exit Sub
    End If
    
    Dim ukupnoKolII As Double
    If chkDveKlaseZbr.value Then
        If Not TryParseDouble(txtUkupnoKgKlIIZbr.value, ukupnoKolII) Or ukupnoKolII <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtUkupnoKgKlIIZbr.SetFocus
            Exit Sub
        End If
    End If

    Dim ukupnoAmb As Long
    If Trim$(txtUkupnoAmbZbr.value) <> "" Then
        If Not TryParseLong(txtUkupnoAmbZbr.value, ukupnoAmb) Then
            MsgBox "Unesite ispravnu kolicinu ambalaže!", vbExclamation, APP_NAME
            txtUkupnoAmbZbr.SetFocus
            Exit Sub
        End If
    End If

    Dim kupacID As String
    kupacID = GetComboID(cmbKupac)

    If kupacID = "" Then
        MsgBox "Nije pronaden ID kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If kupacID = "" Then
        MsgBox "Nije pronaden ID kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim vozacID As String
    vozacID = ExtractIDFromDisplay(cmbVozac.value)

    If vozacID = "" Then
        MsgBox "Nije pronaden ID vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If Not UpdateValidacija() Then
        MsgBox "Validacija nije prošla. Proverite kg i ambalažu (razlika mora biti 0).", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Duplikat check
    If txtBrojZbirne.value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_ZBIRNA, COL_ZBR_BROJ, txtBrojZbirne.value, COL_ZBR_DATUM)

        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    ' Atomicni save kroz modDokumenta wrapper
    Dim result As String

    result = SaveZbirnaMulti_TX( _
        datum:=datumDok, _
        vozacID:=vozacID, _
        brojZbirne:=txtBrojZbirne.value, _
        kupacID:=kupacID, _
        hladnjaca:=cmbHladnjaca.value, _
        pogon:=cmbPogon.value, _
        vrstaVoca:=cmbVrstaVoca.value, _
        sortaVoca:=cmbSortaVoca.value, _
         ukupnoKolI:=ukupnoKolI, _
        tipAmb:=cmbTipAmbZbr.value, _
        ukupnoAmb:=ukupnoAmb, _
        hasKlasaII:=chkDveKlaseZbr.value, _
        ukupnoKolII:=ukupnoKolII)

    If result = "" Then
        MsgBox "Greška pri cuvanju zbirne. Promene su vracene.", vbCritical, APP_NAME
        Exit Sub
    End If

    MsgBox "Zbirna sacuvana: " & result, vbInformation, APP_NAME

    UpdateValidacija
    
    ' Predloži sledeci broj za istog vozaca/datum
    RefreshBrojZbirneSuggestion
    Exit Sub

EH:
    LogErr "frmDokumenta.btnUnosZbr"
    MsgBox "Greška: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' ZBIRNA VALIDIERUNG – Live-Update
' ============================================================

Private Sub txtBrojZbirne_AfterUpdate()
    ' BrojZbirne auch in Otpremnica-Feld setzen
    txtBrojZbirneOtp.value = txtBrojZbirne.value
    UpdateValidacija
End Sub

Private Function UpdateValidacija() As Boolean
    UpdateValidacija = False
    
    If txtBrojZbirne.value = "" Then
        lblValidacijaKG.caption = ""
        lblValidacijaAmb.caption = ""
        Exit Function
    End If
    
    Dim inputKgKlI As Double
    Dim inputKgKlII As Double
    Dim inputAmb As Long
    If Trim$(txtUkupnoKGZbr.value) <> "" Then
        If Not TryParseDouble(txtUkupnoKGZbr.value, inputKgKlI) Then
            lblValidacijaKG.caption = "Neispravna kolicina Kl.I"
            lblValidacijaKG.ForeColor = CLR_ERROR()
            Exit Function
        End If
    End If

    If chkDveKlaseZbr.value Then
        If Trim$(txtUkupnoKgKlIIZbr.value) <> "" Then
            If Not TryParseDouble(txtUkupnoKgKlIIZbr.value, inputKgKlII) Then
                lblValidacijaKG.caption = "Neispravna kolicina Kl.II"
                lblValidacijaKG.ForeColor = CLR_ERROR()
                Exit Function
            End If
        End If
    End If

    If Trim$(txtUkupnoAmbZbr.value) <> "" Then
        If Not TryParseLong(txtUkupnoAmbZbr.value, inputAmb) Then
            lblValidacijaAmb.caption = "Neispravna kolicina ambalaže"
            lblValidacijaAmb.ForeColor = CLR_ERROR()
            Exit Function
        End If
    End If
    
    Dim val As Variant
    val = ValidateZbirnaPreUnosa(txtBrojZbirne.value, inputKgKlI, inputKgKlII, inputAmb)
    
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
    If chkDveKlaseZbr.value Then
        kgCaption = kgCaption & "  ||  Kl.II - Otp: " & Format$(sumaKgII, "#,##0.0") & _
                    " | Zbr: " & Format$(zbrKgII, "#,##0.0") & _
                    " | Raz: " & Format$(razKgII, "#,##0.0")
    End If
    
    lblValidacijaKG.caption = kgCaption
    
    ' Farbe: beide Klassen müssen stimmen
    Dim kgValid As Boolean
    If chkDveKlaseZbr.value Then
        kgValid = validKgI And validKgII
    Else
        kgValid = validKgI
    End If
    
    If kgValid Then
        lblValidacijaKG.ForeColor = CLR_SUCCESS()
    ElseIf zbrKgI = 0 Then
        lblValidacijaKG.ForeColor = TXT_MUTED()
    Else
        lblValidacijaKG.ForeColor = CLR_ERROR()
    End If
    
    ' Ambalaza
    Dim sumaAmb As Long: sumaAmb = CLng(val(8))
    Dim zbrAmb As Long: zbrAmb = CLng(val(9))
    Dim razAmb As Long: razAmb = CLng(val(10))
    
    lblValidacijaAmb.caption = "Amb Otp: " & sumaAmb & " | " & _
                                "Amb Zbr: " & zbrAmb & " | " & _
                                "Raz: " & razAmb
    
    If razAmb = 0 And zbrAmb > 0 Then
        lblValidacijaAmb.ForeColor = CLR_SUCCESS()
    ElseIf zbrAmb = 0 Then
        lblValidacijaAmb.ForeColor = TXT_MUTED()
    Else
        lblValidacijaAmb.ForeColor = CLR_ERROR()
    End If
    
    UpdateValidacija = (kgValid And (razAmb = 0) And (zbrAmb > 0))
End Function

' ============================================================
' PRIJEMNICA UNOS (Ulaz Kupci)
' ============================================================

Private Sub btnUnosPrij_Click()
    On Error GoTo EH

    If cmbKupac.value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If cmbVozac.value = "" Then
        MsgBox "Izaberite vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If Trim$(txtBrojPrij.value) = "" Then
        MsgBox "Unesite broj prijemnice!", vbExclamation, APP_NAME
        txtBrojPrij.SetFocus
        Exit Sub
    End If

    If Trim$(txtBrojZbirnePrij.value) = "" Then
        MsgBox "Unesite broj zbirne!", vbExclamation, APP_NAME
        txtBrojZbirnePrij.SetFocus
        Exit Sub
    End If

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        MsgBox "Unesite ispravan datum!", vbExclamation, APP_NAME
        txtDatum.SetFocus
        Exit Sub
    End If

    Dim kolicinaI As Double
    If Not TryParseDouble(txtKolicinaPrij.value, kolicinaI) Or kolicinaI <= 0 Then
        MsgBox "Unesite ispravnu kolicinu!", vbExclamation, APP_NAME
        txtKolicinaPrij.SetFocus
        Exit Sub
    End If

    Dim cenaI As Double
    If Trim$(txtCenaPrij.value) <> "" Then
        If Not TryParseDouble(txtCenaPrij.value, cenaI) Or cenaI < 0 Then
            MsgBox "Unesite ispravnu cenu!", vbExclamation, APP_NAME
            txtCenaPrij.SetFocus
            Exit Sub
        End If
    End If

    Dim kolAmb As Long
    If Trim$(txtKolAmbPrij.value) <> "" Then
        If Not TryParseLong(txtKolAmbPrij.value, kolAmb) Then
            MsgBox "Unesite ispravnu kolicinu ambalaže!", vbExclamation, APP_NAME
            txtKolAmbPrij.SetFocus
            Exit Sub
        End If
    End If

    Dim kolAmbVracena As Long
    If Trim$(txtKolAmbVracena.value) <> "" Then
        If Not TryParseLong(txtKolAmbVracena.value, kolAmbVracena) Then
            MsgBox "Unesite ispravnu kolicinu vracene ambalaže!", vbExclamation, APP_NAME
            txtKolAmbVracena.SetFocus
            Exit Sub
        End If
    End If

    Dim kolicinaII As Double
    Dim cenaII As Double

    If chkDveKlasePrij.value Then
        If Not TryParseDouble(txtKolicinaKlIIPrij.value, kolicinaII) Or kolicinaII <= 0 Then
            MsgBox "Unesite kolicinu za II klasu!", vbExclamation, APP_NAME
            txtKolicinaKlIIPrij.SetFocus
            Exit Sub
        End If

        If Not TryParseDouble(txtCenaKlIIPrij.value, cenaII) Or cenaII <= 0 Then
            MsgBox "Unesite cenu za II klasu!", vbExclamation, APP_NAME
            txtCenaKlIIPrij.SetFocus
            Exit Sub
        End If
    End If

    Dim kupacID As String
    kupacID = GetComboID(cmbKupac)

    If kupacID = "" Then
        MsgBox "Nije pronaden ID kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If kupacID = "" Then
        MsgBox "Nije pronaden ID kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim vozacID As String
    vozacID = ExtractIDFromDisplay(cmbVozac.value)

    If vozacID = "" Then
        MsgBox "Nije pronaden ID vozaca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If txtBrojPrij.value <> "" Then
        Dim dupMsg As String
        dupMsg = CheckDuplicate(TBL_PRIJEMNICA, COL_PRJ_BROJ, txtBrojPrij.value, COL_PRJ_DATUM)

        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    Dim result As String

    result = SavePrijemnicaMulti_TX( _
        datum:=datumDok, _
        kupacID:=kupacID, _
        vozacID:=vozacID, _
        brojPrij:=txtBrojPrij.value, _
        brojZbirne:=txtBrojZbirnePrij.value, _
        vrstaVoca:=cmbVrstaVoca.value, _
        sortaVoca:=cmbSortaVoca.value, _
        kolicinaI:=kolicinaI, _
        cenaI:=cenaI, _
        tipAmb:=cmbTipAmbPrij.value, _
        kolAmb:=kolAmb, _
        kolAmbVracena:=kolAmbVracena, _
        hasKlasaII:=chkDveKlasePrij.value, _
        kolicinaII:=kolicinaII, _
        cenaII:=cenaII)

    If result = "" Then
        MsgBox "Greška pri cuvanju prijemnice. Promene su vracene.", vbCritical, APP_NAME
        Exit Sub
    End If

    MsgBox "Prijemnica sacuvana: " & result, vbInformation, APP_NAME

    txtBrojPrij.value = ""
    txtKolicinaPrij.value = ""
    txtCenaPrij.value = ""
    txtKolAmbPrij.value = ""
    txtKolAmbVracena.value = ""

    chkDveKlasePrij.value = False
    DisableField txtKolicinaKlIIPrij
    DisableField txtCenaKlIIPrij
    txtKolicinaKlIIPrij.value = ""
    txtCenaKlIIPrij.value = ""

    lblUkupnoKgPrij.caption = ""
    lblProsekGajbe.caption = ""

    If txtBrojZbirnePrij.value <> "" Then
        UpdateManjak txtBrojZbirnePrij.value
    End If

    FillOpenFakture

    Exit Sub

EH:
    LogErr "frmDokumenta.btnUnosPrij"
    MsgBox "Greška: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub ClearPrijemnicaFields()
    txtBrojPrij.value = ""
    txtKolicinaPrij.value = ""
    txtCenaPrij.value = ""
    txtKolAmbPrij.value = ""
    txtKolAmbVracena.value = ""
    chkDveKlasePrij.value = False
    DisableField txtKolicinaKlIIPrij
    DisableField txtCenaKlIIPrij
End Sub

Private Sub txtBrojZbirnePrij_AfterUpdate()
    If txtBrojZbirnePrij.value <> "" Then
        UpdateManjak txtBrojZbirnePrij.value
    End If
End Sub

Private Sub UpdateManjak(ByVal brojZbirne As String)

    Dim pendingKlI As Double, pendingKlII As Double
    If IsNumeric(txtKolicinaPrij.value) Then pendingKlI = CDbl(txtKolicinaPrij.value)
    If chkDveKlasePrij.value Then
        If IsNumeric(txtKolicinaKlIIPrij.value) Then pendingKlII = CDbl(txtKolicinaKlIIPrij.value)
    End If
    
    Dim manjak As Variant
    manjak = CalculateManjakPreview(brojZbirne, pendingKlI, pendingKlII)

    If Not IsArray(manjak) Or UBound(manjak) < 3 Then Exit Sub
    
    ' manjak = Array(ZbirnaKg, PrijemnicaKg, ManjakKg, ManjakPct)
    Dim zbrKg As Double: zbrKg = CDbl(manjak(0))
    Dim prijKg As Double: prijKg = CDbl(manjak(1))
    Dim manjakKg As Double: manjakKg = CDbl(manjak(2))
    Dim manjakPct As Double: manjakPct = CDbl(manjak(3))
    
    lblManjak.caption = "Zbirna: " & Format$(zbrKg, "#,##0.0") & " kg | " & _
                         "Prijemnica: " & Format$(prijKg, "#,##0.0") & " kg | " & _
                         "Manjak: " & Format$(manjakKg, "#,##0.0") & " kg (" & _
                         Format$(manjakPct, "#,##0.00") & "%)"
    
    If Abs(manjakPct) < 0.5 Then
        lblManjak.ForeColor = CLR_SUCCESS()
    ElseIf Abs(manjakPct) < 2 Then
        lblManjak.ForeColor = CLR_WARNING()
    Else
        lblManjak.ForeColor = CLR_ERROR()
    End If
    
    ' Prosek Gajbe
    Dim prosek As Double
    prosek = CalculateProsekGajbeByZbirna(brojZbirne)
    If prosek > 0 Then
        lblProsekGajbe.caption = "Prosek gajbe: " & Format$(prosek, "#,##0.00") & " kg"
    Else
        lblProsekGajbe.caption = ""
    End If
End Sub

Private Sub UpdateUkupnoKgPrij()
    Dim kl1 As Double, kl2 As Double, cena1 As Double, cena2 As Double
    Dim ukupnoKg As Double, ukupnoRsd As Double

    If IsNumeric(txtKolicinaPrij.value) Then kl1 = CDbl(txtKolicinaPrij.value)
    If IsNumeric(txtCenaPrij.value) Then cena1 = CDbl(txtCenaPrij.value)

    If chkDveKlasePrij.value Then
        If IsNumeric(txtKolicinaKlIIPrij.value) Then kl2 = CDbl(txtKolicinaKlIIPrij.value)
        If IsNumeric(txtCenaKlIIPrij.value) Then cena2 = CDbl(txtCenaKlIIPrij.value)
    End If

    ukupnoKg = kl1 + kl2
    ukupnoRsd = (kl1 * cena1) + (kl2 * cena2)

    If ukupnoKg <= 0 And ukupnoRsd <= 0 Then
        lblUkupnoKgPrij.caption = ""
        Exit Sub
    End If

    If chkDveKlasePrij.value Then
        lblUkupnoKgPrij.caption = "UKUPNO  " & Format$(ukupnoKg, "#,##0.0") & " kg"
        If ukupnoRsd > 0 Then
            lblUkupnoKgPrij.caption = lblUkupnoKgPrij.caption & "  •  " & Format$(ukupnoRsd, "#,##0") & " RSD"
        End If
    Else
        If ukupnoRsd > 0 Then
            lblUkupnoKgPrij.caption = "UKUPNO  " & Format$(ukupnoRsd, "#,##0") & " RSD"
        Else
            lblUkupnoKgPrij.caption = ""
        End If
    End If

    StylePreviewBox lblUkupnoKgPrij, "ok"
End Sub

' ============================================================
' IZLAZ KUPCI (Banka-Zahlung + Ambalaza)
' ============================================================

Private Sub btnUnosIzlaz_Click()
    On Error GoTo EH

    If cmbKupac.value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim datumDok As Date
    If Not TryParseDateValue(txtDatum.value, datumDok) Then
        MsgBox "Unesite ispravan datum!", vbExclamation, APP_NAME
        txtDatum.SetFocus
        Exit Sub
    End If

    Dim kupacID As String
    kupacID = GetComboID(cmbKupac)

    If kupacID = "" Then
        MsgBox "Nije pronaden ID kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim vozacID As String
    If cmbVozac.value <> "" Then
        vozacID = ExtractIDFromDisplay(cmbVozac.value)
    End If

    Dim brojDok As String
    brojDok = Trim$(txtBrojDokIzlaz.value)

    If brojDok <> "" Then
        Dim dupMsg As String

        dupMsg = CheckDuplicate(TBL_AMBALAZA, COL_AMB_DOK_ID, brojDok, COL_AMB_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If

        dupMsg = CheckDuplicate(TBL_NOVAC, COL_NOV_BROJ_DOK, brojDok, COL_NOV_DATUM)
        If dupMsg <> "" Then
            MsgBox dupMsg, vbExclamation, APP_NAME
            Exit Sub
        End If
    End If

    Dim kolAmb As Long
    If Trim$(txtKolAmbIzlaz.value) <> "" Then
        If Not TryParseLong(txtKolAmbIzlaz.value, kolAmb) Then
            MsgBox "Unesite ispravnu kolicinu ambalaže!", vbExclamation, APP_NAME
            txtKolAmbIzlaz.SetFocus
            Exit Sub
        End If
    End If

    Dim novac As Double
    If Trim$(txtNovacIzlaz.value) <> "" Then
        If Not TryParseDouble(txtNovacIzlaz.value, novac) Or novac < 0 Then
            MsgBox "Unesite ispravan iznos novca!", vbExclamation, APP_NAME
            txtNovacIzlaz.SetFocus
            Exit Sub
        End If
    End If

    If kolAmb <= 0 And novac <= 0 Then
        MsgBox "Unesite ambalažu ili uplatu kupca.", vbExclamation, APP_NAME
        Exit Sub
    End If

    If kolAmb > 0 And cmbTipAmbIzlaz.value = "" Then
        MsgBox "Izaberite tip ambalaže!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim fakturaID As String
    Dim tipNovca As String
    Dim napomena As String

    If novac > 0 Then
        If cmbFakturaIzlaz.value <> "" Then
            fakturaID = GetComboID(cmbFakturaIzlaz)

            If fakturaID = "" Then
                MsgBox "Nije pronaden ID izabrane fakture.", vbExclamation, APP_NAME
                Exit Sub
            End If

            Dim fakIznos As Double
            Dim uplaceno As Double
            Dim preostalo As Double

            If TryParseDouble(NzToText(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS)), fakIznos) Then
                uplaceno = GetUplataForFaktura(fakturaID)
                preostalo = fakIznos - uplaceno

                If preostalo > 0 And novac > preostalo Then
                    MsgBox "Uplata (" & Format$(novac, "#,##0") & _
                           ") je veca od preostalog iznosa fakture (" & _
                           Format$(preostalo, "#,##0") & ").", _
                           vbExclamation, APP_NAME
                    Exit Sub
                End If
            End If

            tipNovca = NOV_KUPCI_UPLATA
            napomena = "Uplata po fakturi: " & cmbFakturaIzlaz.value
        Else
            fakturaID = ""
            tipNovca = NOV_KUPCI_AVANS
            napomena = "Avans kupca"
        End If
    End If

    If Not SaveKupciIzlaz_TX( _
        datum:=datumDok, _
        brojDok:=brojDok, _
        kupacNaziv:=cmbKupac.value, _
        kupacID:=kupacID, _
        vozacID:=vozacID, _
        tipAmb:=cmbTipAmbIzlaz.value, _
        kolAmb:=kolAmb, _
        vrstaVoca:=cmbVrstaVoca.value, _
        novac:=novac, _
        fakturaID:=fakturaID, _
        napomena:=napomena, _
        tipNovca:=tipNovca) Then

        MsgBox "Greška pri cuvanju izlaza kupca. Promene su vracene.", vbCritical, APP_NAME
        Exit Sub
    End If

    MsgBox "Sacuvano!", vbInformation, APP_NAME

    txtBrojDokIzlaz.value = ""
    txtKolAmbIzlaz.value = ""
    txtNovacIzlaz.value = ""
    cmbFakturaIzlaz.value = ""

    FillOpenFakture

    Exit Sub

EH:
    LogErr "frmDokumenta.btnUnosIzlaz"
    MsgBox "Greška: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub FillOpenFakture()
    On Error GoTo EH

    cmbFakturaIzlaz.Clear
    cmbFakturaIzlaz.ColumnCount = 2
    cmbFakturaIzlaz.ColumnWidths = "300 pt;0 pt"
    cmbFakturaIzlaz.BoundColumn = 1
    cmbFakturaIzlaz.TextColumn = 1

    If cmbKupac.value = "" Then Exit Sub

    Dim kupacID As String
    kupacID = GetComboID(cmbKupac)

    If kupacID = "" Then Exit Sub

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then Exit Sub

    Dim colID As Long
    Dim colBroj As Long
    Dim colDatum As Long
    Dim colKupac As Long
    Dim colIznos As Long
    Dim colStatus As Long

    colID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, "frmDokumenta.FillOpenFakture")
    colBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, "frmDokumenta.FillOpenFakture")
    colDatum = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM, "frmDokumenta.FillOpenFakture")
    colKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, "frmDokumenta.FillOpenFakture")
    colIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, "frmDokumenta.FillOpenFakture")
    colStatus = RequireColumnIndex(TBL_FAKTURE, COL_FAK_STATUS, "frmDokumenta.FillOpenFakture")

    Dim i As Long
    Dim fakturaID As String
    Dim brojFakture As String
    Dim status As String
    Dim iznos As Double
    Dim uplaceno As Double
    Dim preostalo As Double
    Dim datumTxt As String
    Dim displayText As String

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colKupac))) = kupacID Then

            status = Trim$(NzToText(data(i, colStatus)))

            If status <> STATUS_PLACENO Then
                fakturaID = Trim$(NzToText(data(i, colID)))
                brojFakture = Trim$(NzToText(data(i, colBroj)))

                iznos = 0
                If TryParseDouble(NzToText(data(i, colIznos)), iznos) Then
                    uplaceno = GetUplataForFaktura(fakturaID)
                    preostalo = iznos - uplaceno
                Else
                    uplaceno = 0
                    preostalo = 0
                End If

                datumTxt = ""
                If IsDate(data(i, colDatum)) Then
                    datumTxt = Format$(CDate(data(i, colDatum)), "dd.mm.yyyy")
                End If

                displayText = brojFakture

                If datumTxt <> "" Then
                    displayText = displayText & " | " & datumTxt
                End If

                displayText = displayText & " | iznos " & Format$(iznos, "#,##0") & _
                              " | preostalo " & Format$(preostalo, "#,##0")

                cmbFakturaIzlaz.AddItem displayText
                cmbFakturaIzlaz.List(cmbFakturaIzlaz.ListCount - 1, 1) = fakturaID
            End If
        End If
    Next i

    Exit Sub

EH:
    LogErr "frmDokumenta.FillOpenFakture"
    cmbFakturaIzlaz.Clear
End Sub

' ============================================================
' STORNO-BEREICH in frmDokumenta
' Eine TextBox + ComboBox + Button
' ============================================================

Private Sub btnStorno_Click()
    On Error GoTo EH
    Dim tipDok As String
    tipDok = cmbStornoDokument.value
    
    Dim brDok As String
    brDok = Trim$(txtStornoBroj.value)
    
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
        txtStornoBroj.value = ""
        CheckVerwaisteDokumente
    End If
    Exit Sub
EH:
    LogErr "frmDokumenta.btnStorno"
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
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
    
    lblStornoWarning.caption = msg
    lblStornoWarning.ForeColor = CLR_WARNING()
    lblStornoWarning.Visible = True
End Sub

' ============================================================
' NAVIGATION
' ============================================================

Private Sub btnPovratak_Click()
    On Error GoTo EH

    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me

    Exit Sub

EH:
    LogErr "frmDokumenta.btnPovratak_Click"
    Unload Me
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
    On Error Resume Next

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If

    On Error GoTo 0
End Sub


Private Sub FillOpenOtkupi()
    cmbOtkupBlok.Clear
    Erase m_OtkupIDs
    
    If cmbPrimalacOMUlaz.value = "" Then Exit Sub
    
    Dim kooperantID As String
    kooperantID = ExtractIDFromDisplay(cmbPrimalacOMUlaz.value)
    
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

' ============================================================
' v6.11 UI: FOCUS BORDER (gold accent on active input)
' ============================================================

Private Sub txtKolicinaOtp_Enter():     ApplyFocusBorder txtKolicinaOtp:     End Sub
Private Sub txtKolicinaOtp_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtKolicinaOtp: End Sub

Private Sub txtCenaOtp_Enter():         ApplyFocusBorder txtCenaOtp:         End Sub
Private Sub txtCenaOtp_Exit(ByVal Cancel As MSForms.ReturnBoolean):     RemoveFocusBorder txtCenaOtp: End Sub

Private Sub txtKolAmbOtp_Enter():       ApplyFocusBorder txtKolAmbOtp:       End Sub
Private Sub txtKolAmbOtp_Exit(ByVal Cancel As MSForms.ReturnBoolean):   RemoveFocusBorder txtKolAmbOtp: End Sub

Private Sub txtKolicinaKlIIOtp_Enter(): ApplyFocusBorder txtKolicinaKlIIOtp: End Sub
Private Sub txtKolicinaKlIIOtp_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtKolicinaKlIIOtp: End Sub

Private Sub txtCenaKlIIOtp_Enter():     ApplyFocusBorder txtCenaKlIIOtp:     End Sub
Private Sub txtCenaKlIIOtp_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtCenaKlIIOtp: End Sub

Private Sub txtBrojOtp_Enter():         ApplyFocusBorder txtBrojOtp:         End Sub
Private Sub txtBrojOtp_Exit(ByVal Cancel As MSForms.ReturnBoolean):     RemoveFocusBorder txtBrojOtp: End Sub

' Zbirna
Private Sub txtBrojZbirne_Enter():      ApplyFocusBorder txtBrojZbirne:      End Sub
Private Sub txtBrojZbirne_Exit(ByVal Cancel As MSForms.ReturnBoolean):  RemoveFocusBorder txtBrojZbirne: End Sub

Private Sub txtUkupnoKGZbr_Enter():     ApplyFocusBorder txtUkupnoKGZbr:     End Sub
Private Sub txtUkupnoKGZbr_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtUkupnoKGZbr: End Sub

Private Sub txtUkupnoAmbZbr_Enter():    ApplyFocusBorder txtUkupnoAmbZbr:    End Sub
Private Sub txtUkupnoAmbZbr_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtUkupnoAmbZbr: End Sub

Private Sub txtUkupnoKgKlIIZbr_Enter(): ApplyFocusBorder txtUkupnoKgKlIIZbr: End Sub
Private Sub txtUkupnoKgKlIIZbr_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtUkupnoKgKlIIZbr: End Sub

' Prijemnica
Private Sub txtBrojPrij_Enter():        ApplyFocusBorder txtBrojPrij:        End Sub
Private Sub txtBrojPrij_Exit(ByVal Cancel As MSForms.ReturnBoolean):    RemoveFocusBorder txtBrojPrij: End Sub

Private Sub txtBrojZbirnePrij_Enter():  ApplyFocusBorder txtBrojZbirnePrij:  End Sub
Private Sub txtBrojZbirnePrij_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtBrojZbirnePrij: End Sub

Private Sub txtKolicinaPrij_Enter():    ApplyFocusBorder txtKolicinaPrij:    End Sub
Private Sub txtKolicinaPrij_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtKolicinaPrij: End Sub

Private Sub txtCenaPrij_Enter():        ApplyFocusBorder txtCenaPrij:        End Sub
Private Sub txtCenaPrij_Exit(ByVal Cancel As MSForms.ReturnBoolean):    RemoveFocusBorder txtCenaPrij: End Sub

Private Sub txtKolAmbPrij_Enter():      ApplyFocusBorder txtKolAmbPrij:      End Sub
Private Sub txtKolAmbPrij_Exit(ByVal Cancel As MSForms.ReturnBoolean):  RemoveFocusBorder txtKolAmbPrij: End Sub

Private Sub txtKolAmbVracena_Enter():   ApplyFocusBorder txtKolAmbVracena:   End Sub
Private Sub txtKolAmbVracena_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtKolAmbVracena: End Sub

Private Sub txtKolicinaKlIIPrij_Enter(): ApplyFocusBorder txtKolicinaKlIIPrij: End Sub
Private Sub txtKolicinaKlIIPrij_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtKolicinaKlIIPrij: End Sub

Private Sub txtCenaKlIIPrij_Enter():    ApplyFocusBorder txtCenaKlIIPrij:    End Sub
Private Sub txtCenaKlIIPrij_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtCenaKlIIPrij: End Sub

' OM Ulaz
Private Sub txtBrojDokOMUlaz_Enter():   ApplyFocusBorder txtBrojDokOMUlaz:   End Sub
Private Sub txtBrojDokOMUlaz_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtBrojDokOMUlaz: End Sub

Private Sub txtKolAmbOMUlaz_Enter():    ApplyFocusBorder txtKolAmbOMUlaz:    End Sub
Private Sub txtKolAmbOMUlaz_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtKolAmbOMUlaz: End Sub

Private Sub txtNovacOMUlaz_Enter():     ApplyFocusBorder txtNovacOMUlaz:     End Sub
Private Sub txtNovacOMUlaz_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtNovacOMUlaz: End Sub

' Izlaz Kupci
Private Sub txtBrojDokIzlaz_Enter():    ApplyFocusBorder txtBrojDokIzlaz:    End Sub
Private Sub txtBrojDokIzlaz_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtBrojDokIzlaz: End Sub

Private Sub txtKolAmbIzlaz_Enter():     ApplyFocusBorder txtKolAmbIzlaz:     End Sub
Private Sub txtKolAmbIzlaz_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtKolAmbIzlaz: End Sub

Private Sub txtNovacIzlaz_Enter():      ApplyFocusBorder txtNovacIzlaz:      End Sub
Private Sub txtNovacIzlaz_Exit(ByVal Cancel As MSForms.ReturnBoolean): RemoveFocusBorder txtNovacIzlaz: End Sub

' Datum
Private Sub txtDatum_Enter():           ApplyFocusBorder txtDatum:           End Sub
Private Sub txtDatum_Exit(ByVal Cancel As MSForms.ReturnBoolean):       RemoveFocusBorder txtDatum: End Sub

' Storno
Private Sub txtStornoBroj_Enter():      ApplyFocusBorder txtStornoBroj:      End Sub
Private Sub txtStornoBroj_Exit(ByVal Cancel As MSForms.ReturnBoolean):  RemoveFocusBorder txtStornoBroj: End Sub

' ============================================================
' v6.11 UI: F-KEY ACCELERATORS
' F2 = fokus na Otpremnica (Izlaz OM)
' F3 = fokus na Zbirna
' F4 = fokus na Prijemnica (Ulaz Kupci)
' F5 = fokus na OM Ulaz
' F6 = fokus na Izlaz Kupci
' F8 = fokus na Storno
' ============================================================

Private Sub SetupFkeyAccelerators()
    ' Prazno — KeyDown radi rutiranje. Zovem proceduru da ostane explicit hook tacka
    ' za buducu konfiguraciju (npr. role-based shortcuts).
End Sub

Private Sub SetupFkeyHints()
    On Error Resume Next
    ' Pokušaj da nade i stilizuje F-key hint label-e ako ih dodaš u Designer:
    ' lblFkeyOtp, lblFkeyZbr, lblFkeyPrij, lblFkeyOMUlaz, lblFkeyIzlaz, lblFkeyStorno
    StyleFkeyHint Me.Controls("lblFkeyOtp"), "F2"
    StyleFkeyHint Me.Controls("lblFkeyZbr"), "F3"
    StyleFkeyHint Me.Controls("lblFkeyPrij"), "F4"
    StyleFkeyHint Me.Controls("lblFkeyOMUlaz"), "F5"
    StyleFkeyHint Me.Controls("lblFkeyIzlaz"), "F6"
    StyleFkeyHint Me.Controls("lblFkeyStorno"), "F8"
    On Error GoTo 0
End Sub
' ============================================================
' v6.11 UI: F-KEY ACCELERATORS
' VBA UserForm nema KeyPreview, pa rutiramo KeyDown sa svake kontrole
' kroz centralni HandleFkey helper.
' ============================================================

Private Sub HandleFkey(ByVal KeyCode As MSForms.ReturnInteger)
    On Error Resume Next

    Select Case KeyCode
        Case vbKeyF2
            txtBrojOtp.SetFocus
            KeyCode = 0
        Case vbKeyF3
            txtBrojZbirne.SetFocus
            KeyCode = 0
        Case vbKeyF4
            txtBrojPrij.SetFocus
            KeyCode = 0
        Case vbKeyF5
            txtBrojDokOMUlaz.SetFocus
            KeyCode = 0
        Case vbKeyF6
            txtBrojDokIzlaz.SetFocus
            KeyCode = 0
        Case vbKeyF8
            cmbStornoDokument.SetFocus
            KeyCode = 0
    End Select

    On Error GoTo 0
End Sub

' --- KeyDown handleri po kontroli ---
' Cilj: F-tasteri rade bez obzira gde je fokus.
' Svaka input kontrola koja ce primati fokus mora imati svoj _KeyDown.

Private Sub txtBrojOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):       HandleFkey KeyCode: End Sub
Private Sub txtKolicinaOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    HandleFkey KeyCode: End Sub
Private Sub txtCenaOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):        HandleFkey KeyCode: End Sub
Private Sub txtKolAmbOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      HandleFkey KeyCode: End Sub
Private Sub txtKolicinaKlIIOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): HandleFkey KeyCode: End Sub
Private Sub txtCenaKlIIOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    HandleFkey KeyCode: End Sub
Private Sub txtBrojZbirneOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):  HandleFkey KeyCode: End Sub

Private Sub txtBrojZbirne_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     HandleFkey KeyCode: End Sub
Private Sub txtUkupnoKGZbr_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    HandleFkey KeyCode: End Sub
Private Sub txtUkupnoAmbZbr_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub
Private Sub txtUkupnoKgKlIIZbr_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): HandleFkey KeyCode: End Sub

Private Sub txtBrojPrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):       HandleFkey KeyCode: End Sub
Private Sub txtBrojZbirnePrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): HandleFkey KeyCode: End Sub
Private Sub txtKolicinaPrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub
Private Sub txtCenaPrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):       HandleFkey KeyCode: End Sub
Private Sub txtKolAmbPrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     HandleFkey KeyCode: End Sub
Private Sub txtKolAmbVracena_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):  HandleFkey KeyCode: End Sub
Private Sub txtKolicinaKlIIPrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): HandleFkey KeyCode: End Sub
Private Sub txtCenaKlIIPrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub

Private Sub txtBrojDokOMUlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):  HandleFkey KeyCode: End Sub
Private Sub txtKolAmbOMUlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub
Private Sub txtNovacOMUlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    HandleFkey KeyCode: End Sub

Private Sub txtBrojDokIzlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub
Private Sub txtKolAmbIzlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    HandleFkey KeyCode: End Sub
Private Sub txtNovacIzlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     HandleFkey KeyCode: End Sub

Private Sub txtDatum_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):          HandleFkey KeyCode: End Sub
Private Sub txtStornoBroj_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     HandleFkey KeyCode: End Sub

' ComboBox-ovi takodje moraju
Private Sub cmbVrstaVoca_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      HandleFkey KeyCode: End Sub
Private Sub cmbSortaVoca_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      HandleFkey KeyCode: End Sub
Private Sub cmbOtkupnoMesto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub
Private Sub cmbKupac_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):          HandleFkey KeyCode: End Sub
Private Sub cmbHladnjaca_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      HandleFkey KeyCode: End Sub
Private Sub cmbPogon_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):          HandleFkey KeyCode: End Sub
Private Sub cmbVozac_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):          HandleFkey KeyCode: End Sub
Private Sub cmbTipAmbOtp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      HandleFkey KeyCode: End Sub
Private Sub cmbTipAmbZbr_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      HandleFkey KeyCode: End Sub
Private Sub cmbTipAmbPrij_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):     HandleFkey KeyCode: End Sub
Private Sub cmbTipAmbOMUlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub
Private Sub cmbTipAmbIzlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):    HandleFkey KeyCode: End Sub
Private Sub cmbPrimalacOMUlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): HandleFkey KeyCode: End Sub
Private Sub cmbOtkupBlok_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):      HandleFkey KeyCode: End Sub
Private Sub cmbFakturaIzlaz_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer):   HandleFkey KeyCode: End Sub
Private Sub cmbStornoDokument_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer): HandleFkey KeyCode: End Sub

' ============================================================
' v6.11 UI: TOP KPI CARDS
' Danas | OM saldo | Otvoreno kg | Validacija
' ============================================================

Private Sub LayoutTopKpis()
    On Error GoTo EH

    LayoutTopKpiInternals fraKPIDanas, lblKpiDanasTitle, lblKpiDanasValue, lblKpiDanasAccent
    LayoutTopKpiInternals fraKpiOM, lblKpiOMTitle, lblKpiOMValue, lblKpiOMAccent
    LayoutTopKpiInternals fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent
    LayoutTopKpiInternals fraKpiValidacija, lblKpiValTitle, lblKpiValValue, lblKpiValAccent

    Exit Sub

EH:
    LogErr "frmDokumenta.LayoutTopKpis"
End Sub

Public Sub RefreshTopKpis()
    On Error GoTo EH

    ' --- 1) Danas (datum) ---
    StyleTopKpi fraKPIDanas, lblKpiDanasTitle, lblKpiDanasValue, lblKpiDanasAccent, "neutral"
    lblKpiDanasTitle.caption = "Danas"
    lblKpiDanasValue.caption = Format$(Date, "d.m.yyyy")

    ' --- 2) OM saldo (avans po izabranom OM, ili 0) ---
    Dim omSaldo As Double
    omSaldo = 0
    If cmbOtkupnoMesto.value <> "" Then
        Dim sid As String
        sid = GetComboID(cmbOtkupnoMesto)
        If sid <> "" Then omSaldo = GetOMAvansSaldo(sid)
    End If

    StyleTopKpi fraKpiOM, lblKpiOMTitle, lblKpiOMValue, lblKpiOMAccent, _
                IIf(omSaldo > 0, "ok", "neutral")
    lblKpiOMTitle.caption = "OM saldo"
    lblKpiOMValue.caption = Format$(omSaldo, "#,##0")

    ' --- 3) Otvoreno kg (zbir kg za danas iz tblOtkup) ---
    Dim kgToday As Double
    kgToday = SumOtkupKgToday()

    StyleTopKpi fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent, _
                IIf(kgToday > 0, "ok", "neutral")
    lblKpiOtvTitle.caption = "Otvoreno kg"
    lblKpiOtvValue.caption = Format$(kgToday, "#,##0")

    ' --- 4) Validacija (status zbirne ako je popunjeno, inace neutral) ---
    Dim valKind As String
    Dim valText As String

    If txtBrojZbirne.value = "" Then
        valKind = "neutral"
        valText = "—"
    Else
        ' Zovem postojecu UpdateValidacija (vraca True ako je 0/0)
        If UpdateValidacija() Then
            valKind = "ok"
            valText = "OK"
        Else
            valKind = "warn"
            valText = "Razlika"
        End If
    End If

    StyleTopKpi fraKpiValidacija, lblKpiValTitle, lblKpiValValue, lblKpiValAccent, valKind
    lblKpiValTitle.caption = "Validacija"
    lblKpiValValue.caption = valText

    Exit Sub

EH:
    LogErr "frmDokumenta.RefreshTopKpis"
End Sub

Private Function SumOtkupKgToday() As Double
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim colDatum As Long, colKg As Long, colStorno As Long
    colDatum = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, "SumOtkupKgToday")
    colKg = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "SumOtkupKgToday")

    On Error Resume Next
    colStorno = RequireColumnIndex(TBL_OTKUP, "Stornirano", "SumOtkupKgToday")
    On Error GoTo EH

    Dim i As Long, total As Double, rowKg As Double

    For i = 1 To UBound(data, 1)
        If colStorno > 0 Then
            If LCase$(NzToText(data(i, colStorno))) = "da" Then GoTo NextRow
        End If

        If IsDate(data(i, colDatum)) Then
            If DateValue(CDate(data(i, colDatum))) = DateValue(Date) Then
                If TryParseDouble(NzToText(data(i, colKg)), rowKg) Then
                    total = total + rowKg
                End If
            End If
        End If
NextRow:
    Next i

    SumOtkupKgToday = total
    Exit Function

EH:
    LogErr "SumOtkupKgToday"
    SumOtkupKgToday = 0
End Function

