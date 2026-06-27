Attribute VB_Name = "modOtkupBlok"
Option Explicit

' ============================================================
' modOtkupBlok - Panel "Otkupni blokovi" u frmOtkup.
'
' Panel NE unosi sam u tblOtkup - vodi POSTOJECU levu frmOtkup formu:
'   - Klik na otpremnicu (sredina) popuni levu formu: otkupno mesto,
'     vrsta, sorta, vozac, datum, broj zbirne i cenu; broj otkupnog
'     lista racuna kanonski SuggestNextBroj (OM iz polja + datum otpr.).
'   - Gore: "Cena po otpremnici" (DEFAULT za nove blokove; rucni override u
'     txtCena se postuje i ne pregazuje) + sazetak Preostalo za unos.
'   - Korisnik unese kooperanta + kolicinu i klikne postojeci "Unos".
'   - frmOtkup.btnUnos_Click pre snimanja zove OtkupBlok_ConfirmUnos
'     (upozorenje na prekoracenje), a posle uspeha OtkupBlok_AfterUnos
'     (vezivanje OtkupID->OtpremnicaID + osvezavanje + auto-deselekcija
'     kad Preostalo padne na 0).
'   - Desno: blokovi izabrane otpremnice (+ zbirni red), sa dugmadima
'     "Storniraj blok" (StornoOtkup_TX) i ChrW(352) & "tampaj list" (PrintOtkupniList).
'   - Lista otpremnica: kolona "Ostatak", filter "samo nezavrsene",
'     sort po datumu (najnovije gore).
'
' Sve kontrole panela su dinamicke (Controls.Add) - frmOtkup.frx se ne
' menja. Cena se cuva kao BRUTO (sa PDV nadoknadom); neto/PDV iz nje.
'
' Integracija u frmOtkup:
'   UserForm_Initialize:                 AttachOtkupBlokPanel Me
'   btnUnos_Click (pre snimanja):        If Not OtkupBlok_ConfirmUnos() Then Exit Sub
'   btnUnos_Click (po uspehu, posle ClearOtkupFields): OtkupBlok_AfterUnos result
' ============================================================

' --- Layout (tacke) ---
Private Const PANEL_LEFT  As Double = 302
Private Const OTP_W       As Double = 346
Private Const BLOK_LEFT   As Double = 652       ' PANEL_LEFT + OTP_W + 4 (manji razmak)
Private Const BLOK_W      As Double = 504
Private Const GRID_TOP    As Double = 120       ' spusteno za akcioni red (dugmad + datum spec)
Private Const EXP_WIDTH   As Double = 1164
Private Const TOGGLE_W    As Double = 130

Private Const OTP_COLW  As String = "0;0;58;38;56;58;44;42;36"
Private Const BLOK_COLW As String = "0;62;104;58;44;46;66;58;66"

' Const ne moze da sadrzi ChrW() poziv -- koristimo funkcije umesto konstanti
Private Function OTP_CAPS() As String
    OTP_CAPS = ";;Otkupno mesto;Koli" & ChrW(269) & "ina;Datum;Kupac;Prodajna;Cena za;Ostatak"
End Function
Private Function BLOK_CAPS() As String
    BLOK_CAPS = ";br. bloka;Ime i Prezime;Datum;Koli" & ChrW(269) & "ina;Cena bez PDV;Vrednost;Iznos PDV;Ukupna vrednost"
End Function

' --- Stanje (modul-level; jedna frmOtkup instanca po sekciji) ---
Private mForm As Object
Private mWrappers As Collection
Private mPanelCtls As Collection
Private mCenaBlok As Object           ' OtpremnicaID -> cena (bruto)
Private mBuilt As Boolean
Private mVisible As Boolean
Private mOrigWidth As Double
Private mActiveOtpID As String
Private mFilterNezavrsene As Boolean
Private mMultiMode As Boolean

Private mBtnToggle As MSForms.CommandButton
Private mBtnStorno As MSForms.CommandButton
Private mBtnPrint As MSForms.CommandButton
Private mBtnFilter As MSForms.CommandButton
Private mBtnBiraj As MSForms.CommandButton
Private mBtnSpecDatum As MSForms.CommandButton
Private mBtnLost As MSForms.CommandButton        ' sekcija "Izgubljeni blokovi"
Private mBtnPreuzmi As MSForms.CommandButton
Private mLostMode As Boolean
Private mTxtSpecOd As MSForms.TextBox
Private mTxtSpecDo As MSForms.TextBox
Private mLstOtp As MSForms.ListBox
Private mLstBlok As MSForms.ListBox
Private mTxtCenaOtp As MSForms.TextBox
Private mLblUkupno As MSForms.label
Private mLblNapisano As MSForms.label
Private mLblPreostalo As MSForms.label
Private mLblUkupnoAmb As MSForms.label
Private mLblNapisanoAmb As MSForms.label
Private mLblPreostaloAmb As MSForms.label
Private mLblZbirna As MSForms.label

' ============================================================
' PUBLIC - ulazna tacka + event ruteri + frmOtkup hooks
' ============================================================

Public Sub AttachOtkupBlokPanel(ByVal frm As Object)
    On Error GoTo EH

    Set mForm = frm
    Set mWrappers = New Collection
    Set mCenaBlok = CreateObject("Scripting.Dictionary")
    mBuilt = False
    mVisible = False
    mActiveOtpID = ""
    mFilterNezavrsene = False
    mMultiMode = False
    mOrigWidth = mForm.width

    If UCase$(Trim$(GetConfigValue("OTKUP_BLOK_PANEL"))) = "NO" Then Exit Sub

    Set mBtnToggle = mForm.Controls.Add("Forms.CommandButton.1", "btnOtkBlokToggle", True)
    mBtnToggle.width = TOGGLE_W
    mBtnToggle.Height = 24
    mBtnToggle.top = 6
    mBtnToggle.Left = mForm.InsideWidth - TOGGLE_W - 6
    mBtnToggle.caption = Poruka("OTKUP_LBL_OTKUPNI_BLOKOVI")
    On Error Resume Next
    StylePrimaryButton mBtnToggle, Poruka("OTKUP_LBL_OTKUPNI_BLOKOVI")
    On Error GoTo EH

    WireBtn mBtnToggle, "TOGGLE"
    Exit Sub
EH:
    LogErr "modOtkupBlok.AttachOtkupBlokPanel"
End Sub

Public Sub OtkupBlok_OnButton(ByVal action As String)
    On Error GoTo EH
    Select Case action
        Case "TOGGLE": TogglePanel
        Case "STORNO": StornoSelectedBlok
        Case "PRINT": PrintSelectedBlok
        Case "FILTER": ToggleFilter
        Case "BIRAJ": BirajOrPrint
        Case "SPECDATUM": PrintSpecOdDo
        Case "LOST": ToggleLostMode
        Case "ADOPT": AdoptSelectedLostBlok
    End Select
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_OnButton"
End Sub

Public Sub OtkupBlok_OnText(ByVal action As String)
    On Error GoTo EH
    If action = "CENA" Then OnCenaTyping
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_OnText"
End Sub

Public Sub OtkupBlok_OnTextAfter(ByVal action As String)
    On Error GoTo EH
    If action = "CENA" Then OnCenaChanged
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_OnTextAfter"
End Sub

Public Sub OtkupBlok_OnListClick(ByVal action As String)
    On Error GoTo EH
    If action = "OTP" Then SelectOtpFromList
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_OnListClick"
End Sub

' Pre snimanja (frmOtkup.btnUnos_Click): upozorenje na prekoracenje preostale
' kolicine otpremnice. Vraca False samo ako operater odustane. Panel neaktivan
' / greska -> True (nikad ne blokira normalan unos).
Public Function OtkupBlok_ConfirmUnos() As Boolean
    OtkupBlok_ConfirmUnos = True
    On Error GoTo EH

    If Not mVisible Then Exit Function
    If Len(mActiveOtpID) = 0 Then Exit Function

    Dim kol As Double, k2 As Double
    If Not TryParseDouble(CStr(mForm.Controls("txtKolicina").value), kol) Then kol = 0
    On Error Resume Next
    If mForm.Controls("chkDveKlase").value = True Then
        TryParseDouble CStr(mForm.Controls("txtKolicinaKLII").value), k2
    End If
    On Error GoTo EH

    ' Bruto unos: preostalo otpremnice je u NETO -> oduzmi taru ambalaze sa Klase I
    ' (i Klase II, ako postoji) da poredjenje bude nedvosmisleno (neto vs neto).
    If OtkupBrutoUnos() Then
        Dim ka As Long, tw As Double
        Dim tip As String
        On Error Resume Next
        tip = CStr(mForm.Controls("cmbTipAmbalaze").value)
        If TryParseLong(CStr(mForm.Controls("txtKolAmbalaze").value), ka) Then
            If ka > 0 And Len(Trim$(tip)) > 0 Then
                tw = ka * GetTezinaGajbice(tip)
                If tw > 0 And tw < kol Then kol = kol - tw
            End If
        End If
        ' Klasa II: zasebne gajbe (runtime polje txtKolAmbalazeIIRT).
        Dim ka2 As Long, tw2 As Double
        If k2 > 0 And TryParseLong(CStr(mForm.Controls("txtKolAmbalazeIIRT").value), ka2) Then
            If ka2 > 0 And Len(Trim$(tip)) > 0 Then
                tw2 = ka2 * GetTezinaGajbice(tip)
                If tw2 > 0 And tw2 < k2 Then k2 = k2 - tw2
            End If
        End If
        On Error GoTo EH
    End If

    Dim total As Double: total = kol + k2
    If total <= 0 Then Exit Function

    Dim ukupno As Double
    ukupno = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_KOLICINA))
    Dim preost As Double: preost = ukupno - SumKolByOtp(mActiveOtpID)

    If total > preost + 0.0001 Then
        If MsgBox("Unos (" & FmtKgDec(total) & " kg) premasuje preostalih " & _
                  FmtKgDec(preost) & " kg za ovu otpremnicu." & vbCrLf & _
                  "Nastaviti?", vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            OtkupBlok_ConfirmUnos = False
        End If
    End If
    Exit Function
EH:
    LogErr "modOtkupBlok.OtkupBlok_ConfirmUnos"
    OtkupBlok_ConfirmUnos = True
End Function

' Posle uspesnog "Unos": vezi sacuvane redove za otpremnicu, ujednaci cenu,
' osvezi i auto-deselektuj kad je otpremnica popunjena.
Public Sub OtkupBlok_AfterUnos(ByVal otkupIDs As String)
    On Error GoTo EH
    If Not mVisible Then Exit Sub
    If Len(mActiveOtpID) = 0 Then Exit Sub

    LinkOtkupIDsToOtpremnica otkupIDs, mActiveOtpID

    ' "Cena po otpremnici" je DEFAULT za SLEDECI blok: ClearOtkupFields je obrisao
    ' txtCena, pa ga vracamo na default radi brzog unosa. Cenu UPRAVO unetog bloka
    ' (txtCena u trenutku snimanja -- moze biti rucni override) NE diramo: vec je
    ' sacuvana sa blokom i default ne sme da je pregazi.
    If mCenaBlok.Exists(mActiveOtpID) Then
        SetLeftCtl "txtCena", Format$(CDbl(mCenaBlok(mActiveOtpID)), "0.00")
    End If

    LoadBlokovi
    LoadOtpremnice
    RefreshSummary

    ' auto-deselekcija kad je otpremnica popunjena (zastita od pogresnog vezivanja)
    Dim ukupno As Double
    ukupno = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_KOLICINA))
    If ukupno - SumKolByOtp(mActiveOtpID) <= 0.0001 Then
        MsgBox "Otpremnica je popunjena (Preostalo = 0). Izaberite drugu otpremnicu za sledeci blok.", _
               vbInformation, APP_NAME
        mActiveOtpID = ""
        LoadBlokovi
        RefreshSummary
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_AfterUnos"
End Sub

' ============================================================
' TOGGLE + BUILD
' ============================================================

Private Sub TogglePanel()
    On Error GoTo EH

    If Not mBuilt Then
        BuildPanel
        mBuilt = True
    End If

    mVisible = Not mVisible
    SetPanelVisible mVisible

    If mVisible Then
        LoadOtpremnice
        LoadBlokovi
        mForm.width = EXP_WIDTH
        mBtnToggle.caption = Poruka("OTKUP_LBL_SAKRIJ_BLOKOVE")
    Else
        mForm.width = mOrigWidth
        mBtnToggle.caption = Poruka("OTKUP_LBL_OTKUPNI_BLOKOVI")
    End If

    mBtnToggle.Left = mForm.InsideWidth - mBtnToggle.width - 6
    Exit Sub
EH:
    LogErr "modOtkupBlok.TogglePanel"
End Sub

Private Sub BuildPanel()
    Set mPanelCtls = New Collection

    Dim gridH As Double
    gridH = mForm.InsideHeight - GRID_TOP - 14
    If gridH < 120 Then gridH = 120

    ' Gornji red: cena po otpremnici + sazetak
    Dim lblc As Object
    Set lblc = AddCtl("Label", "lblOtkBlokCenaL", PANEL_LEFT, 7, 108, 14)
    lblc.caption = "Cena po otpremnici:": StyleHdr lblc
    Set mTxtCenaOtp = AddCtl("TextBox", "txtOtkBlokCenaOtp", PANEL_LEFT + 112, 5, 60, 18)
    On Error Resume Next
    StyleTextBox mTxtCenaOtp
    On Error GoTo 0

    Set mLblUkupno = AddCtl("Label", "lblOtkBlokUk", PANEL_LEFT + 190, 7, 150, 14)
    Set mLblNapisano = AddCtl("Label", "lblOtkBlokNap", PANEL_LEFT + 346, 7, 150, 14)
    Set mLblPreostalo = AddCtl("Label", "lblOtkBlokPre", PANEL_LEFT + 502, 7, 150, 14)
    mLblUkupno.caption = Poruka("OTKUP_LBL_UKUPNO")
    mLblNapisano.caption = Poruka("OTKUP_LBL_BLOKOVIMA")
    mLblPreostalo.caption = Poruka("OTKUP_LBL_OSTATAK")

    ' Drugi red sazetka: ambalaza (ispod kg)
    Set mLblUkupnoAmb = AddCtl("Label", "lblOtkBlokUkAmb", PANEL_LEFT + 190, 22, 150, 14)
    Set mLblNapisanoAmb = AddCtl("Label", "lblOtkBlokNapAmb", PANEL_LEFT + 346, 22, 150, 14)
    Set mLblPreostaloAmb = AddCtl("Label", "lblOtkBlokPreAmb", PANEL_LEFT + 502, 22, 150, 14)
    ' Info (#4): broj zbirne za izabranu otpremnicu (azurira RefreshSummary).
    Set mLblZbirna = AddCtl("Label", "lblOtkBlokZbirna", BLOK_LEFT + 70, 44, 380, 16)
    On Error Resume Next
    mLblZbirna.WordWrap = True
    On Error GoTo 0
    StyleHdr mLblZbirna
    mLblZbirna.caption = "Zbirna: -"
    mLblUkupnoAmb.caption = Poruka("OTKUP_LBL_UKUPNO_AMB")
    mLblNapisanoAmb.caption = Poruka("OTKUP_LBL_BLOKOVIMA_AMB")
    mLblPreostaloAmb.caption = Poruka("OTKUP_LBL_OSTATAK_AMB")

    ' Naslovi (red 44) + filter nad listom otpremnica.
    Dim t1 As Object: Set t1 = AddCtl("Label", "lblOtkBlokT1", PANEL_LEFT, 44, 226, 14)
    t1.caption = "OTPREMNICE  (klik = izbor)": StyleHdr t1
    Set mBtnFilter = AddCtl("CommandButton", "btnOtkBlokFilter", PANEL_LEFT + OTP_W - 120, 42, 120, 22)
    mBtnFilter.caption = "Prikaz: Sve"

    Dim t2 As Object: Set t2 = AddCtl("Label", "lblOtkBlokT2", BLOK_LEFT, 44, 56, 14)
    t2.caption = "BLOKOVI": StyleHdr t2

    ' --- Akcioni red (red 66), iznad listboxova ---
    ' Levo (nad otpremnicama): dnevna / periodicna specifikacija (Od/Do + dugme).
    ' Datum Od/Do pretpopunjeni na danas -> klik daje dnevnu specifikaciju;
    ' promenom datuma dobija se period (od-do). Reuse istog renderera (RenderSpec).
    Dim lblOd As Object: Set lblOd = AddCtl("Label", "lblOtkBlokSpecOd", PANEL_LEFT, 69, 22, 14)
    lblOd.caption = "Od:": StyleHdr lblOd
    Set mTxtSpecOd = AddCtl("TextBox", "txtOtkBlokSpecOd", PANEL_LEFT + 24, 66, 64, 18)
    Dim lblDo As Object: Set lblDo = AddCtl("Label", "lblOtkBlokSpecDo", PANEL_LEFT + 94, 69, 22, 14)
    lblDo.caption = "Do:": StyleHdr lblDo
    Set mTxtSpecDo = AddCtl("TextBox", "txtOtkBlokSpecDo", PANEL_LEFT + 118, 66, 64, 18)
    Set mBtnSpecDatum = AddCtl("CommandButton", "btnOtkBlokSpecDatum", PANEL_LEFT + 190, 66, 150, 22)
    mBtnSpecDatum.caption = ChrW(352) & "tampaj po datumu"

    ' Desno (nad blokovima): storno / stampa lista / biranje otpremnica za spec.
    Set mBtnStorno = AddCtl("CommandButton", "btnOtkBlokStorno", BLOK_LEFT, 66, 78, 22)
    mBtnStorno.caption = "Storniraj"
    Set mBtnPrint = AddCtl("CommandButton", "btnOtkBlokPrint", BLOK_LEFT + 82, 66, 84, 22)
    mBtnPrint.caption = ChrW(352) & "tampaj list"
    Set mBtnBiraj = AddCtl("CommandButton", "btnOtkBlokBiraj", BLOK_LEFT + 170, 66, 124, 22)
    mBtnBiraj.caption = "Biraj otpremnice"
    ' Sekcija "Izgubljeni blokovi" (slobodan prostor desno na akcionom redu).
    Set mBtnLost = AddCtl("CommandButton", "btnOtkBlokLost", BLOK_LEFT + 300, 66, 116, 22)
    mBtnLost.caption = "Izgubljeni"
    Set mBtnPreuzmi = AddCtl("CommandButton", "btnOtkBlokPreuzmi", BLOK_LEFT + 420, 66, 82, 22)
    mBtnPreuzmi.caption = "Preuzmi"

    On Error Resume Next
    StyleExitButton mBtnFilter, "Prikaz: Sve"
    StyleTextBox mTxtSpecOd
    StyleTextBox mTxtSpecDo
    StylePrimaryButton mBtnSpecDatum, ChrW(352) & "tampaj po datumu"
    StyleExitButton mBtnStorno, "Storniraj"
    StylePrimaryButton mBtnPrint, ChrW(352) & "tampaj list"
    StyleExitButton mBtnBiraj, "Biraj otpremnice"
    StyleExitButton mBtnLost, "Izgubljeni"
    StylePrimaryButton mBtnPreuzmi, "Preuzmi"
    On Error GoTo 0
    mTxtSpecOd.value = Format$(Date, "d.m.yyyy")
    mTxtSpecDo.value = Format$(Date, "d.m.yyyy")

    ' Zaglavlja kolona (red 90; listbox tik ispod)
    AddHeaders "hOtp", PANEL_LEFT, 90, OTP_COLW, OTP_CAPS
    AddHeaders "hBlok", BLOK_LEFT, 90, BLOK_COLW, BLOK_CAPS

    ' Grid-ovi
    Set mLstOtp = AddCtl("ListBox", "lstOtkBlokOtp", PANEL_LEFT, GRID_TOP, OTP_W, gridH)
    mLstOtp.ColumnCount = 9
    mLstOtp.ColumnWidths = OTP_COLW

    Set mLstBlok = AddCtl("ListBox", "lstOtkBlokBlok", BLOK_LEFT, GRID_TOP, BLOK_W, gridH)
    mLstBlok.ColumnCount = 9
    mLstBlok.ColumnWidths = BLOK_COLW

    ' Eventi
    WireTxt mTxtCenaOtp, "CENA"
    WireLst mLstOtp, "OTP"
    WireBtn mBtnStorno, "STORNO"
    WireBtn mBtnPrint, "PRINT"
    WireBtn mBtnFilter, "FILTER"
    WireBtn mBtnBiraj, "BIRAJ"
    WireBtn mBtnSpecDatum, "SPECDATUM"
    WireBtn mBtnLost, "LOST"
    WireBtn mBtnPreuzmi, "ADOPT"
End Sub

Private Sub SetPanelVisible(ByVal b As Boolean)
    Dim c As Object
    For Each c In mPanelCtls
        On Error Resume Next
        c.Visible = b
        On Error GoTo 0
    Next c
End Sub

' ============================================================
' LOAD - pregled otpremnica (sredina) + blokovi izabrane otpremnice (desno)
' ============================================================

Private Sub LoadOtpremnice()
    On Error GoTo EH
    mLstOtp.Clear

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Sub
    data = ExcludeStornirano(data, TBL_OTPREMNICA)
    If IsEmpty(data) Then Exit Sub

    Dim cId As Long, cBroj As Long, cSt As Long, cDat As Long
    Dim cZbr As Long, cKol As Long, cCena As Long
    cId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    cBroj = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    cSt = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA)
    cDat = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM)
    cZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    cKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    cCena = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA)

    Dim dSt As Object: Set dSt = BuildLookup(TBL_STANICE, "StanicaID", "Naziv")
    ' Kolona "Kupac" (firma): BrojZbirne -> KupacID (zbirna) -> Naziv (kupci).
    Dim dKupId As Object: Set dKupId = BuildLookup(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_KUPAC)
    Dim dKupNaziv As Object: Set dKupNaziv = BuildLookup(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
    Dim dCe As Object: Set dCe = BuildFirstBlokCena()
    Dim dNap As Object: Set dNap = BuildNapisanoByOtp()

    ' Indeksi za prikaz (uz filter "samo nezavrsene") + sort po datumu DESC
    Dim n As Long: n = UBound(data, 1)
    Dim idx() As Long: ReDim idx(1 To n)
    Dim keyd() As Double: ReDim keyd(1 To n)
    Dim m As Long: m = 0
    Dim i As Long
    For i = 1 To n
        Dim id0 As String: id0 = CStr(data(i, cId))
        Dim nap0 As Double: nap0 = 0
        If dNap.Exists(id0) Then nap0 = dNap(id0)
        Dim pre0 As Double: pre0 = NumVal(data(i, cKol)) - nap0
        If (Not mFilterNezavrsene) Or (pre0 > 0.0001) Then
            m = m + 1
            idx(m) = i
            If IsDate(data(i, cDat)) Then keyd(m) = CDbl(CDate(data(i, cDat))) Else keyd(m) = 0
        End If
    Next i

    Dim a As Long, b As Long, ti As Long, tk As Double
    For a = 2 To m
        ti = idx(a): tk = keyd(a): b = a - 1
        Do While b >= 1
            If keyd(b) >= tk Then Exit Do
            idx(b + 1) = idx(b): keyd(b + 1) = keyd(b): b = b - 1
        Loop
        idx(b + 1) = ti: keyd(b + 1) = tk
    Next a

    Dim r As Long, j As Long
    For j = 1 To m
        i = idx(j)
        Dim otpID As String: otpID = CStr(data(i, cId))
        Dim prodajna As Double: prodajna = NumVal(data(i, cCena))
        Dim cenaBlok As Double
        If mCenaBlok.Exists(otpID) Then
            cenaBlok = mCenaBlok(otpID)
        ElseIf dCe.Exists(otpID) Then
            cenaBlok = dCe(otpID)
        Else
            cenaBlok = prodajna
        End If
        Dim ukupno As Double: ukupno = NumVal(data(i, cKol))
        Dim nap As Double: nap = 0
        If dNap.Exists(otpID) Then nap = dNap(otpID)

        mLstOtp.AddItem otpID
        r = mLstOtp.ListCount - 1
        mLstOtp.List(r, 1) = CStr(data(i, cBroj))
        mLstOtp.List(r, 2) = DictVal(dSt, CStr(data(i, cSt)))
        mLstOtp.List(r, 3) = FmtKgDec(ukupno)
        mLstOtp.List(r, 4) = FmtDate(data(i, cDat))
        mLstOtp.List(r, 5) = KupacNazivZaZbirnu(dKupId, dKupNaziv, CStr(data(i, cZbr)))
        mLstOtp.List(r, 6) = FmtKg(prodajna)
        mLstOtp.List(r, 7) = FmtKg(cenaBlok)
        mLstOtp.List(r, 8) = FmtKgDec(ukupno - nap)
    Next j
    Exit Sub
EH:
    LogErr "modOtkupBlok.LoadOtpremnice"
End Sub

Private Sub LoadBlokovi()
    On Error GoTo EH
    If mLostMode Then
        LoadLostBlokovi
        Exit Sub
    End If
    mLstBlok.Clear
    If Len(mActiveOtpID) = 0 Then Exit Sub

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub

    Dim cId As Long, cOtp As Long, cKoop As Long, cKol As Long
    Dim cCena As Long, cBr As Long, cDat As Long, cAmb As Long
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    cAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)

    Dim dKo As Object: Set dKo = BuildKoopNames()
    Dim stopa As Double: stopa = PdvStopa()

    Dim sumKol As Double, sumVred As Double, sumPdv As Double, sumUk As Double, sumAmb As Double
    Dim i As Long, r As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOtp))) <> mActiveOtpID Then GoTo NextRow

        Dim kol As Double: kol = NumVal(data(i, cKol))
        Dim amb As Double: amb = NumVal(data(i, cAmb))
        Dim bruto As Double: bruto = NumVal(data(i, cCena))
        Dim neto As Double: neto = bruto / (1 + stopa / 100)
        Dim vred As Double: vred = kol * neto
        Dim pdv As Double: pdv = vred * stopa / 100
        Dim uk As Double: uk = kol * bruto

        mLstBlok.AddItem CStr(data(i, cId))      ' col0 (hidden) = OtkupID
        r = mLstBlok.ListCount - 1
        mLstBlok.List(r, 1) = CStr(data(i, cBr))
        mLstBlok.List(r, 2) = DictVal(dKo, Trim$(CStr(data(i, cKoop))))
        mLstBlok.List(r, 3) = FmtDate(data(i, cDat))
        mLstBlok.List(r, 4) = FmtKgDec(kol)
        mLstBlok.List(r, 5) = FmtRsd(neto)
        mLstBlok.List(r, 6) = FmtRsd(vred)
        mLstBlok.List(r, 7) = FmtRsd(pdv)
        mLstBlok.List(r, 8) = FmtRsd(uk)

        sumKol = sumKol + kol: sumVred = sumVred + vred
        sumPdv = sumPdv + pdv: sumUk = sumUk + uk
        sumAmb = sumAmb + amb
NextRow:
    Next i

    ' Zbirni red
    If mLstBlok.ListCount > 0 Then
        mLstBlok.AddItem ""                       ' col0 prazan -> storno/print ga preskace
        r = mLstBlok.ListCount - 1
        mLstBlok.List(r, 1) = "UKUPNO"
        mLstBlok.List(r, 2) = "amb: " & FmtKg(sumAmb)
        mLstBlok.List(r, 4) = FmtKgDec(sumKol)
        mLstBlok.List(r, 6) = FmtRsd(sumVred)
        mLstBlok.List(r, 7) = FmtRsd(sumPdv)
        mLstBlok.List(r, 8) = FmtRsd(sumUk)
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.LoadBlokovi"
End Sub

' ============================================================
' INTERAKCIJA
' ============================================================

Private Sub SelectOtpFromList()
    On Error GoTo EH
    If mMultiMode Then Exit Sub          ' multiselect za stampu - ne vodi levu formu
    If mLstOtp.ListIndex < 0 Then Exit Sub

    mActiveOtpID = CStr(mLstOtp.List(mLstOtp.ListIndex, 0))
    If Len(mActiveOtpID) = 0 Then Exit Sub

    Dim cena As Double
    If mCenaBlok.Exists(mActiveOtpID) Then
        cena = mCenaBlok(mActiveOtpID)
    Else
        cena = ExistingBlokCena(mActiveOtpID)
        If cena <= 0 Then cena = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_CENA))
        mCenaBlok(mActiveOtpID) = cena
    End If
    mTxtCenaOtp.value = Format$(cena, "0.00")

    PrefillLeftForm mActiveOtpID, cena
    LoadBlokovi
    RefreshSummary

    ' brzi unos: fokus odmah na Kooperanta
    On Error Resume Next
    mForm.Controls("cmbKooperant").SetFocus
    On Error GoTo 0
    Exit Sub
EH:
    LogErr "modOtkupBlok.SelectOtpFromList"
End Sub

Private Sub PrefillLeftForm(ByVal otpID As String, ByVal cena As Double)
    On Error Resume Next

    ' Datum (i broj zbirne) PRE otkupnog mesta.
    Dim vDat As Variant: vDat = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_DATUM)
    If IsDate(vDat) Then SetLeftCtl "txtDatum", Format$(CDate(vDat), "d.m.yyyy")
    SetLeftCtl "txtBrojZbirne", CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE))

    SetComboByIdAny mForm.Controls("cmbOtkupnoMesto"), _
                    CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA))
    mForm.Controls("cmbVrstaVoca").value = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_VRSTA))
    mForm.Controls("cmbSortaVoca").value = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_SORTA))
    SetComboByIdAny mForm.Controls("cmbVozac"), _
                    CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_VOZAC))

    SetLeftCtl "txtCena", Format$(cena, "0.00")

    ' Broj otkupnog lista: izricito (cmbOtkupnoMesto_Change ne okida ako je OM isti
    ' kao kod prethodne otpremnice). OM iz polja + datum otpremnice, kanonski generator.
    Dim stanicaID As String: stanicaID = GetComboID(mForm.Controls("cmbOtkupnoMesto"))
    If Len(stanicaID) > 0 And IsDate(vDat) Then
        Dim brDok As String: brDok = SuggestNextBroj(KIND_OTK, stanicaID, CDate(vDat), False)
        If Len(brDok) > 0 Then SetLeftCtl "txtBrojDokumenta", brDok
    End If
End Sub

' Kucanje u "Cena po otpremnici" (Change) -> uzivo u levu txtCena + kolona "Cena za".
Private Sub OnCenaTyping()
    On Error GoTo EH
    If Len(mActiveOtpID) = 0 Then Exit Sub
    Dim cena As Double
    If Not TryParseDouble(mTxtCenaOtp.value, cena) Or cena <= 0 Then Exit Sub
    mCenaBlok(mActiveOtpID) = cena
    SetLeftCtl "txtCena", Format$(cena, "0.00")

    Dim li As Long: li = mLstOtp.ListIndex
    If li >= 0 Then
        If CStr(mLstOtp.List(li, 0)) = mActiveOtpID Then mLstOtp.List(li, 7) = FmtKg(cena)
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.OnCenaTyping"
End Sub

' Promena cene gore (AfterUpdate) -> azurira DEFAULT cenu otpremnice (za sledeci
' blok) + osvezava prikaz. Vec uneti blokovi se NE preracunavaju: "Cena po
' otpremnici" je default, a rucni override po bloku (txtCena) se postuje.
Private Sub OnCenaChanged()
    On Error GoTo EH
    If Len(mActiveOtpID) = 0 Then Exit Sub
    Dim cena As Double
    If Not TryParseDouble(mTxtCenaOtp.value, cena) Or cena <= 0 Then Exit Sub

    mCenaBlok(mActiveOtpID) = cena
    mTxtCenaOtp.value = Format$(cena, "0.00")
    SetLeftCtl "txtCena", Format$(cena, "0.00")
    LoadOtpremnice
    RefreshSummary
    Exit Sub
EH:
    LogErr "modOtkupBlok.OnCenaChanged"
End Sub

Private Sub ToggleFilter()
    On Error GoTo EH
    mFilterNezavrsene = Not mFilterNezavrsene
    If mFilterNezavrsene Then
        mBtnFilter.caption = "Prikaz: Nezavrsene"
    Else
        mBtnFilter.caption = "Prikaz: Sve"
    End If
    LoadOtpremnice
    Exit Sub
EH:
    LogErr "modOtkupBlok.ToggleFilter"
End Sub

Private Sub StornoSelectedBlok()
    On Error GoTo EH
    Dim li As Long: li = mLstBlok.ListIndex
    If li < 0 Then
        MsgBox "Izaberite blok za storno.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim otkupID As String: otkupID = Trim$(CStr(mLstBlok.List(li, 0)))
    If Len(otkupID) = 0 Then Exit Sub      ' zbirni red

    ' Klasa I i II dele isti broj dokumenta (zaseban OtkupID po klasi) -> storno
    ' obuhvata CEO otkup (obe klase), ne samo izabrani blok. (col1 = BrDok)
    Dim brDok As String: brDok = Trim$(CStr(mLstBlok.List(li, 1)))

    If MsgBox("Stornirati ceo otkup br. " & brDok & " (" & CStr(mLstBlok.List(li, 2)) & _
              ", sve klase)?", vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Sub

    Dim ok As Boolean
    If Len(brDok) > 0 Then
        ok = StornoOtkupByBrDok_TX(brDok)
    Else
        ok = StornoOtkup_TX(otkupID)       ' fallback: blok bez broja dokumenta
    End If

    If ok Then
        LoadBlokovi
        LoadOtpremnice
        RefreshSummary
        MsgBox "Otkup storniran: " & brDok, vbInformation, APP_NAME
    Else
        MsgBox "Storno nije uspeo za " & brDok & ".", vbCritical, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.StornoSelectedBlok"
    MsgBox "Gre" & ChrW(353) & "ka pri storno bloka: " & Err.description, vbCritical, APP_NAME
End Sub

' Sekcija "Izgubljeni blokovi": blokovi cija je otpremnica stornirana/nestala.
Private Sub LoadLostBlokovi()
    On Error GoTo EH
    mLstBlok.Clear
    Dim lost As Variant: lost = GetLostOtkupBlokovi()
    If Not IsArray(lost) Then Exit Sub

    Dim dKo As Object: Set dKo = BuildKoopNames()
    Dim i As Long, r As Long
    For i = 1 To UBound(lost, 1)
        mLstBlok.AddItem CStr(lost(i, 1))                  ' col0 (skriveno) = OtkupID
        r = mLstBlok.ListCount - 1
        mLstBlok.List(r, 1) = CStr(lost(i, 2))             ' br. bloka
        mLstBlok.List(r, 2) = DictVal(dKo, Trim$(CStr(lost(i, 3))))  ' kooperant
        mLstBlok.List(r, 3) = FmtDate(lost(i, 4))          ' datum
        mLstBlok.List(r, 4) = FmtKgDec(NumVal(lost(i, 5))) ' kolicina
        mLstBlok.List(r, 6) = "stara otp: " & CStr(lost(i, 7))       ' kol. "Vrednost"
    Next i
    Exit Sub
EH:
    LogErr "modOtkupBlok.LoadLostBlokovi"
End Sub

' Toggle prikaza izgubljenih blokova u listi BLOKOVI.
Private Sub ToggleLostMode()
    On Error GoTo EH
    mLostMode = Not mLostMode
    If mLostMode Then mBtnLost.caption = "Nazad" Else mBtnLost.caption = "Izgubljeni"
    LoadBlokovi
    Exit Sub
EH:
    LogErr "modOtkupBlok.ToggleLostMode"
End Sub

' Preuzmi izabrani izgubljeni blok na trenutno izabranu (aktivnu) otpremnicu.
' Re-point veze (OtpremnicaID + BrojZbirne); OtkupID/uplate/ambalaza ostaju.
Private Sub AdoptSelectedLostBlok()
    On Error GoTo EH
    If Not mLostMode Then
        MsgBox "Klikni 'Izgubljeni', pa izaberi blok za preuzimanje.", vbInformation, APP_NAME
        Exit Sub
    End If
    Dim li As Long: li = mLstBlok.ListIndex
    If li < 0 Then MsgBox "Izaberite izgubljeni blok.", vbExclamation, APP_NAME: Exit Sub
    Dim otkupID As String: otkupID = Trim$(CStr(mLstBlok.List(li, 0)))
    If Len(otkupID) = 0 Then Exit Sub
    Dim brDok As String: brDok = Trim$(CStr(mLstBlok.List(li, 1)))

    If Len(mActiveOtpID) = 0 Then
        MsgBox "Prvo izaberi CILJNU otpremnicu (leva lista), pa Preuzmi.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim tBroj As String
    tBroj = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_BROJ))

    If MsgBox("Preuzeti blok br. " & brDok & " na otpremnicu " & tBroj & "?" & vbCrLf & _
              "(menja se samo veza; OtkupID, uplate i ambala" & ChrW(382) & "a ostaju.)", _
              vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Sub

    If ReassignOtkupToOtpremnica_TX(otkupID, mActiveOtpID) Then
        mLostMode = False
        mBtnLost.caption = "Izgubljeni"
        LoadOtpremnice
        LoadBlokovi
        RefreshSummary
        MsgBox "Blok preuzet na otpremnicu " & tBroj & ".", vbInformation, APP_NAME
    Else
        MsgBox "Preuzimanje nije uspelo (cilj mozda storniran).", vbCritical, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.AdoptSelectedLostBlok"
    MsgBox "Gre" & ChrW(353) & "ka pri preuzimanju: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub PrintSelectedBlok()
    On Error GoTo EH
    Dim li As Long: li = mLstBlok.ListIndex
    If li < 0 Then
        MsgBox "Izaberite blok za " & ChrW(353) & "tampu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim otkupID As String: otkupID = Trim$(CStr(mLstBlok.List(li, 0)))
    If Len(otkupID) = 0 Then Exit Sub      ' zbirni red

    PrintOtkupniList otkupID
    Exit Sub
EH:
    LogErr "modOtkupBlok.PrintSelectedBlok"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi otkupnog lista: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' MULTISELECT + SPECIFIKACIJA (batch stampa otpremnica)
' ============================================================

' Dugme "Biraj otpremnice" <-> ChrW(352) & "tampaj specifikaciju".
Private Sub BirajOrPrint()
    On Error GoTo EH
    If Not mMultiMode Then
        ' udji u multiselect rezim za stampu
        mMultiMode = True
        mActiveOtpID = ""
        On Error Resume Next
        mLstOtp.MultiSelect = fmMultiSelectMulti
        On Error GoTo EH
        mBtnBiraj.caption = ChrW(352) & "tampaj specifikaciju"
        LoadBlokovi
        RefreshSummary
    Else
        ' skupi izabrane otpremnice -> stampa -> vrati u normalan rezim
        Dim sel As Collection: Set sel = New Collection
        Dim i As Long
        For i = 0 To mLstOtp.ListCount - 1
            If mLstOtp.Selected(i) Then sel.Add CStr(mLstOtp.List(i, 0))
        Next i
        If sel.count = 0 Then
            MsgBox "Izaberite bar jednu otpremnicu (klik na redove).", vbExclamation, APP_NAME
            Exit Sub
        End If

        PrintSpecifikacija sel

        mMultiMode = False
        On Error Resume Next
        mLstOtp.MultiSelect = fmMultiSelectSingle
        On Error GoTo EH
        mBtnBiraj.caption = "Biraj otpremnice"
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.BirajOrPrint"
End Sub

' Specifikacija RUCNO izabranih otpremnica (postojeci tok: dugme "Biraj
' otpremnice" -> multiselect -> ChrW(352) & "tampaj specifikaciju"). Tanak omotac oko
' zajednickog renderera RenderSpec (filter po skupu OtpremnicaID).
Private Sub PrintSpecifikacija(ByVal otpIDs As Collection)
    On Error GoTo EH
    Dim selSet As Object: Set selSet = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    For Each v In otpIDs
        Dim oid0 As String: oid0 = CStr(v)
        If Not selSet.Exists(oid0) Then selSet.Add oid0, True
    Next v

    Dim subtitle As String
    subtitle = "Datum stampe: " & Format$(Date, "d.m.yyyy") & "     Otpremnica: " & otpIDs.count
    RenderSpec selSet, False, Date, Date, subtitle
    Exit Sub
EH:
    LogErr "modOtkupBlok.PrintSpecifikacija"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi specifikacije: " & Err.description, vbCritical, APP_NAME
End Sub

' Handler dugmeta ChrW(352) & "tampaj po datumu" (footer strip). Cita Od/Do polja
' (TryParseDateValue) i zove renderer u rezimu filtera po datumu.
Private Sub PrintSpecOdDo()
    On Error GoTo EH
    Dim dOd As Date, dDo As Date
    If Not TryParseDateValue(CStr(mTxtSpecOd.value), dOd) Then
        MsgBox "Unesite ispravan datum 'Od' (npr. " & Format$(Date, "d.m.yyyy") & ").", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not TryParseDateValue(CStr(mTxtSpecDo.value), dDo) Then
        MsgBox "Unesite ispravan datum 'Do' (npr. " & Format$(Date, "d.m.yyyy") & ").", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    If dDo < dOd Then
        MsgBox "'Do' datum ne sme biti pre 'Od' datuma.", vbExclamation, APP_NAME
        Exit Sub
    End If
    PrintSpecifikacijaPoDatumu dOd, dDo
    Exit Sub
EH:
    LogErr "modOtkupBlok.PrintSpecOdDo"
End Sub

' Dnevna / periodicna specifikacija: svi (ne-stornirani) otkup blokovi cija je
' kolona Datum u opsegu [datumOd, datumDo]. Tanak omotac oko RenderSpec.
Private Sub PrintSpecifikacijaPoDatumu(ByVal datumOd As Date, ByVal datumDo As Date)
    On Error GoTo EH
    Dim subtitle As String
    If Int(CDbl(datumOd)) = Int(CDbl(datumDo)) Then
        subtitle = "Dnevna specifikacija     Datum: " & Format$(datumOd, "d.m.yyyy")
    Else
        subtitle = "Specifikacija     Period: " & Format$(datumOd, "d.m.yyyy") & _
                   " - " & Format$(datumDo, "d.m.yyyy")
    End If
    RenderSpec Nothing, True, datumOd, datumDo, subtitle
    Exit Sub
EH:
    LogErr "modOtkupBlok.PrintSpecifikacijaPoDatumu"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi specifikacije: " & Err.description, vbCritical, APP_NAME
End Sub

' Jezgro: ispisuje specifikaciju otkupnih blokova (tabela sa okvirima, A4
' landscape) i exportuje u PDF. Filter po redu:
'   byDate=True  -> kolona Datum u [datumOd, datumDo]  (selSet sme biti Nothing)
'   byDate=False -> OtpremnicaID u selSet              (rucna selekcija)
' Izlaz je sortiran po (Otkupno mesto, Datum) radi grupisanja.
Private Sub RenderSpec(ByVal selSet As Object, ByVal byDate As Boolean, _
                       ByVal datumOd As Date, ByVal datumDo As Date, _
                       ByVal subtitle As String)
    On Error GoTo EH

    Dim dKo As Object: Set dKo = BuildKoopNames()
    Dim dSt As Object: Set dSt = BuildLookup(TBL_STANICE, "StanicaID", "Naziv")
    Dim dZbr As Object: Set dZbr = BuildLookup(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ_ZBIRNE)
    Dim dOtp As Object: Set dOtp = BuildLookup(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ)
    Dim stopa As Double: stopa = PdvStopa()

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then MsgBox "Nema podataka u otkupu.", vbInformation, APP_NAME: Exit Sub
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then MsgBox "Nema blokova.", vbInformation, APP_NAME: Exit Sub

    Dim cOtp As Long, cKoop As Long, cKol As Long, cCena As Long, cBr As Long, cDat As Long, cSt As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    cSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)

    ' --- 1) skupi indekse redova koji prolaze filter + sort-kljuc (OM | datum) ---
    Dim n As Long: n = UBound(data, 1)
    Dim idx() As Long: ReDim idx(1 To n)
    Dim keys() As String: ReDim keys(1 To n)
    Dim m As Long: m = 0
    Dim i As Long
    For i = 1 To n
        Dim oid As String: oid = Trim$(CStr(data(i, cOtp)))
        Dim pass As Boolean: pass = False
        If byDate Then
            If IsDate(data(i, cDat)) Then
                Dim dd As Double: dd = Int(CDbl(CDate(data(i, cDat))))
                pass = (dd >= Int(CDbl(datumOd)) And dd <= Int(CDbl(datumDo)))
            End If
        ElseIf Not selSet Is Nothing Then
            pass = selSet.Exists(oid)
        End If
        If pass Then
            m = m + 1
            idx(m) = i
            Dim dkey As String: dkey = "00000000"
            If IsDate(data(i, cDat)) Then dkey = Format$(CDate(data(i, cDat)), "yyyymmdd")
            keys(m) = DictVal(dSt, CStr(data(i, cSt))) & "|" & dkey
        End If
    Next i

    If m = 0 Then
        If byDate Then
            MsgBox "Nema otkupnih blokova u izabranom periodu.", vbInformation, APP_NAME
        Else
            MsgBox "Izabrane otpremnice nemaju blokova.", vbInformation, APP_NAME
        End If
        Exit Sub
    End If

    ' --- 2) insertion sort po (Otkupno mesto, Datum) ASC (kao u LoadOtpremnice) ---
    Dim a As Long, b As Long, ti As Long, tk As String
    For a = 2 To m
        ti = idx(a): tk = keys(a): b = a - 1
        Do While b >= 1
            If keys(b) <= tk Then Exit Do
            idx(b + 1) = idx(b): keys(b + 1) = keys(b): b = b - 1
        Loop
        idx(b + 1) = ti: keys(b + 1) = tk
    Next a

    ' --- 3) ispis ---
    Dim ws As Worksheet: Set ws = EnsureSpecSheet()
    Application.ScreenUpdating = False
    ws.cells.Font.name = "Calibri"
    ws.cells.Font.Size = 9

    ws.cells(1, 1).value = "SPECIFIKACIJA OTKUPNIH BLOKOVA"
    ws.cells(1, 1).Font.Size = 12
    ws.cells(1, 1).Font.Bold = True

    Dim hdr As Variant
    hdr = Array("Broj zbirne", "Broj otpremnice", "Otkupno mesto", "br. bloka", "Ime i Prezime", _
                "Datum", "Koli" & ChrW(269) & "ina", "Cena bez PDV", "Vrednost", "Iznos PDV", "Ukupna vrednost")
    Const R0 As Long = 4
    Const NC As Long = 11
    Dim cc As Long
    For cc = 0 To UBound(hdr)
        ws.cells(R0, cc + 1).value = hdr(cc)
        ws.cells(R0, cc + 1).Font.Bold = True
    Next cc

    Dim r As Long: r = R0 + 1
    Dim sumKol As Double, sumVred As Double, sumPdv As Double, sumUk As Double, cnt As Long
    Dim j As Long
    For j = 1 To m
        i = idx(j)
        Dim oid2 As String: oid2 = Trim$(CStr(data(i, cOtp)))
        Dim kol As Double: kol = NumVal(data(i, cKol))
        Dim bruto As Double: bruto = NumVal(data(i, cCena))
        Dim neto As Double: neto = bruto / (1 + stopa / 100)
        Dim vred As Double: vred = kol * neto
        Dim pdv As Double: pdv = vred * stopa / 100
        Dim uk As Double: uk = kol * bruto

        ws.cells(r, 1).value = DictVal(dZbr, oid2)
        ws.cells(r, 2).value = DictVal(dOtp, oid2)
        ws.cells(r, 3).value = DictVal(dSt, CStr(data(i, cSt)))
        ws.cells(r, 4).value = CStr(data(i, cBr))
        ws.cells(r, 5).value = DictVal(dKo, Trim$(CStr(data(i, cKoop))))
        ws.cells(r, 6).value = FmtDate(data(i, cDat))
        ws.cells(r, 7).value = kol
        ws.cells(r, 8).value = neto
        ws.cells(r, 9).value = vred
        ws.cells(r, 10).value = pdv
        ws.cells(r, 11).value = uk
        sumKol = sumKol + kol: sumVred = sumVred + vred
        sumPdv = sumPdv + pdv: sumUk = sumUk + uk
        cnt = cnt + 1
        r = r + 1
    Next j

    ' UKUPNO red
    ws.cells(r, 5).value = "UKUPNO"
    ws.cells(r, 7).value = sumKol
    ws.cells(r, 9).value = sumVred
    ws.cells(r, 10).value = sumPdv
    ws.cells(r, 11).value = sumUk
    ws.Range(ws.cells(r, 1), ws.cells(r, NC)).Font.Bold = True

    ws.cells(2, 1).value = subtitle & "     Blokova: " & cnt

    ' formati: kolicina sa decimalama samo ako su unete, novac 2 decimale
    ws.Range(ws.cells(R0 + 1, 7), ws.cells(r, 7)).NumberFormat = "#,##0.###"
    ws.Range(ws.cells(R0 + 1, 8), ws.cells(r, NC)).NumberFormat = "#,##0.00"

    ' iscrtana polja
    Dim tbl As Range: Set tbl = ws.Range(ws.cells(R0, 1), ws.cells(r, NC))
    tbl.Borders.LineStyle = xlContinuous
    tbl.Borders.Weight = xlThin
    tbl.rows.RowHeight = 13

    ' sirine kolona
    ws.columns("A").ColumnWidth = 12   ' Broj zbirne
    ws.columns("B").ColumnWidth = 14   ' Broj otpremnice
    ws.columns("C").ColumnWidth = 18   ' Otkupno mesto
    ws.columns("D").ColumnWidth = 12   ' br. bloka
    ws.columns("E").ColumnWidth = 24   ' Ime i Prezime
    ws.columns("F").ColumnWidth = 11   ' Datum
    ws.columns("G").ColumnWidth = 9    ' Kolicina
    ws.columns("H").ColumnWidth = 11   ' Cena bez PDV
    ws.columns("I").ColumnWidth = 13   ' Vrednost
    ws.columns("J").ColumnWidth = 12   ' Iznos PDV
    ws.columns("K").ColumnWidth = 14   ' Ukupna vrednost

    ' PageSetup (Orientation/PaperSize/FitToPages) trazi drajver stampaca.
    ' Na racunaru bez podrazumevanog stampaca ti property-ji bacaju gresku 1004,
    ' pa ih stitimo: PDF mora da izadje i kad stampaca nema (best-effort izgled).
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$" & R0 & ":$" & R0
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(r, NC)).Address
    End With
    On Error GoTo EH

    ' Direktno u PDF pored radne sveske i otvori odmah (bez preview-a).
    ' Vremenski pecat u imenu -> nema "file in use" ako je prethodni PDF otvoren.
    Dim pdfPath As String
    pdfPath = ThisWorkbook.path & "\Specifikacija_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    Dim wasHidden As Boolean: wasHidden = (ws.Visible <> xlSheetVisible)
    ws.Visible = xlSheetVisible
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, OpenAfterPublish:=True
    If wasHidden Then ws.Visible = xlSheetHidden

    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    LogErr "modOtkupBlok.RenderSpec"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi specifikacije: " & Err.description, vbCritical, APP_NAME
End Sub

Private Function EnsureSpecSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SpecifikacijaSablon")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "SpecifikacijaSablon"
        ws.Visible = xlSheetHidden
    End If
    ws.cells.Clear
    Set EnsureSpecSheet = ws
End Function

Private Sub RefreshSummary()
    On Error GoTo EH

    If Not mLblZbirna Is Nothing Then
        If Len(mActiveOtpID) = 0 Then
            mLblZbirna.caption = "Zbirna: -"
        Else
            Dim zbrBroj As String, prjBroj As String
            zbrBroj = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_BROJ_ZBIRNE))
            If Len(Trim$(zbrBroj)) > 0 Then prjBroj = PrijemnicaBrojZaZbirnu(zbrBroj)
            If Len(Trim$(zbrBroj)) = 0 Then zbrBroj = "(nije vezana)"
            If Len(prjBroj) > 0 Then
                mLblZbirna.caption = "Zbirna: " & zbrBroj & " | Prijemnica: " & prjBroj
            Else
                mLblZbirna.caption = "Zbirna: " & zbrBroj
            End If
        End If
    End If

    If Len(mActiveOtpID) = 0 Then
        mLblUkupno.caption = Poruka("OTKUP_LBL_UKUPNO")
        mLblNapisano.caption = Poruka("OTKUP_LBL_BLOKOVIMA")
        mLblPreostalo.caption = Poruka("OTKUP_LBL_OSTATAK")
        mLblUkupnoAmb.caption = Poruka("OTKUP_LBL_UKUPNO_AMB")
        mLblNapisanoAmb.caption = Poruka("OTKUP_LBL_BLOKOVIMA_AMB")
        mLblPreostaloAmb.caption = Poruka("OTKUP_LBL_OSTATAK_AMB")
        Exit Sub
    End If

    Dim ukupno As Double
    ukupno = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_KOLICINA))
    Dim ukupnoBruto As Double
    ukupnoBruto = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_BRUTO))
    If ukupnoBruto <= 0 Then ukupnoBruto = ukupno
    Dim napisano As Double: napisano = SumKolByOtp(mActiveOtpID)
    Dim napisanoBruto As Double: napisanoBruto = SumBrutoByOtp(mActiveOtpID)
    Dim preostalo As Double: preostalo = ukupno - napisano

    Dim ukupnoAmb As Double
    ukupnoAmb = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_KOL_AMB))
    Dim napisanoAmb As Double: napisanoAmb = SumAmbByOtp(mActiveOtpID)
    Dim preostaloAmb As Double: preostaloAmb = ukupnoAmb - napisanoAmb

    mLblUkupno.caption = "Ukupno kg: " & FmtKgBrutoNeto(ukupnoBruto, ukupno)
    mLblNapisano.caption = "U blokovima: " & FmtKgBrutoNeto(napisanoBruto, napisano)
    mLblPreostalo.caption = "Ostatak: " & FmtKgBrutoNeto(ukupnoBruto - napisanoBruto, preostalo)

    mLblUkupnoAmb.caption = "Ukupno amb: " & FmtKg(ukupnoAmb)
    mLblNapisanoAmb.caption = "U blokovima amb: " & FmtKg(napisanoAmb)
    mLblPreostaloAmb.caption = "Ostatak amb: " & FmtKg(preostaloAmb)

    On Error Resume Next
    If preostalo < -0.0001 Then
        mLblPreostalo.ForeColor = RGB(200, 0, 0)
    Else
        mLblPreostalo.ForeColor = RGB(0, 120, 0)
    End If
    If preostaloAmb < -0.0001 Then
        mLblPreostaloAmb.ForeColor = RGB(200, 0, 0)
    Else
        mLblPreostaloAmb.ForeColor = RGB(0, 120, 0)
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.RefreshSummary"
End Sub

' ============================================================
' VEZIVANJE + CENA
' ============================================================

Private Sub LinkOtkupIDsToOtpremnica(ByVal otkupIDs As String, ByVal otpID As String)
    Dim tx As clsTransaction
    On Error GoTo EH
    If Len(otpID) = 0 Or Len(Trim$(otkupIDs)) = 0 Then Exit Sub

    Dim ids() As String: ids = Split(otkupIDs, " + ")

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim j As Long
    For j = LBound(ids) To UBound(ids)
        Dim id As String: id = Trim$(ids(j))
        If Len(id) > 0 Then
            Dim rows As Collection: Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, id)
            Dim k As Long
            For k = 1 To rows.count
                RequireUpdateCell TBL_OTKUP, rows(k), COL_OTK_OTPREMNICA_ID, otpID, _
                                  "modOtkupBlok.LinkOtkupIDsToOtpremnica"
            Next k
        End If
    Next j

    tx.CommitTx
    Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "modOtkupBlok.LinkOtkupIDsToOtpremnica"
End Sub

' ============================================================
' HELPERS
' ============================================================

' BrojZbirne -> broj(evi) prijemnice (veza otpremnica<->prijemnica je preko
' BrojZbirne). Vise prijemnica iste zbirne -> spojeno zarezom; storno preskace.
Private Function PrijemnicaBrojZaZbirnu(ByVal brojZbirne As String) As String
    On Error GoTo EH
    If Len(Trim$(brojZbirne)) = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function

    Dim cZbr As Long, cBroj As Long
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cBroj = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    If cZbr = 0 Or cBroj = 0 Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long, out As String, b As String
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cZbr))) = Trim$(brojZbirne) Then
            b = Trim$(CStr(data(i, cBroj)))
            If Len(b) > 0 And Not seen.Exists(b) Then
                seen.Add b, True
                If Len(out) > 0 Then out = out & ", "
                out = out & b
            End If
        End If
    Next i
    PrijemnicaBrojZaZbirnu = out
    Exit Function
EH:
    LogErr "modOtkupBlok.PrijemnicaBrojZaZbirnu"
End Function

Private Sub SetLeftCtl(ByVal nm As String, ByVal val As String)
    On Error Resume Next
    mForm.Controls(nm).value = val
    On Error GoTo 0
End Sub

' Postavi combo po ID-u; radi i za 2-kolonske (bound) i 1-kolonske "... (ID)".
Private Sub SetComboByIdAny(ByVal cmb As Object, ByVal idValue As String)
    On Error Resume Next
    idValue = Trim$(idValue)
    If Len(idValue) = 0 Then Exit Sub

    If SetComboByID(cmb, idValue) Then Exit Sub
    If cmb.ListIndex >= 0 Then Exit Sub

    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If ExtractIDFromDisplay(CStr(cmb.List(i))) = idValue Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Function SumKolByOtp(ByVal otpID As String) As Double
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim cOtp As Long, cKol As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)

    Dim i As Long, s As Double
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOtp))) = otpID Then s = s + NumVal(data(i, cKol))
    Next i
    SumKolByOtp = s
End Function

' Zbir BRUTO kg otkup blokova za otpremnicu (BrutoKg po redu; ako je prazno -> neto).
Private Function SumBrutoByOtp(ByVal otpID As String) As Double
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim cOtp As Long, cKol As Long, cBruto As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cBruto = GetColumnIndex(TBL_OTKUP, COL_OTK_BRUTO)

    Dim i As Long, s As Double, b As Double
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOtp))) = otpID Then
            b = 0
            If cBruto > 0 Then b = NumVal(data(i, cBruto))
            If b <= 0 Then b = NumVal(data(i, cKol))   ' red bez bruto -> bruto = neto
            s = s + b
        End If
    Next i
    SumBrutoByOtp = s
End Function

Private Function SumAmbByOtp(ByVal otpID As String) As Double
    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function

    Dim cOtp As Long, cAmb As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)

    Dim i As Long, s As Double
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOtp))) = otpID Then s = s + NumVal(data(i, cAmb))
    Next i
    SumAmbByOtp = s
End Function

Private Function ExistingBlokCena(ByVal otpID As String) As Double
    If Len(otpID) = 0 Then Exit Function
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, otpID)
    If rows.count = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cCena As Long: cCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    ExistingBlokCena = NumVal(data(rows(1), cCena))
End Function

' OtpremnicaID -> cena prvog povezanog bloka.
Private Function BuildFirstBlokCena() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildFirstBlokCena = d

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cOtp As Long, cCena As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cOtp)))
        If Len(k) > 0 And Not d.Exists(k) Then d.Add k, NumVal(data(i, cCena))
    Next i
End Function

' OtpremnicaID -> ukupna kolicina svih (ne-storniranih) blokova.
Private Function BuildNapisanoByOtp() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildNapisanoByOtp = d

    Dim data As Variant: data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    Dim cOtp As Long, cKol As Long
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cOtp)))
        If Len(k) > 0 Then
            If d.Exists(k) Then
                d(k) = d(k) + NumVal(data(i, cKol))
            Else
                d.Add k, NumVal(data(i, cKol))
            End If
        End If
    Next i
End Function

Private Function BuildLookup(ByVal tbl As String, ByVal keyName As String, _
                             ByVal valName As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildLookup = d

    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim ck As Long, cv As Long
    ck = GetColumnIndex(tbl, keyName)
    cv = GetColumnIndex(tbl, valName)
    If ck = 0 Or cv = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, ck)))
        If Len(k) > 0 And Not d.Exists(k) Then d.Add k, CStr(data(i, cv))
    Next i
End Function

Private Function BuildKoopNames() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set BuildKoopNames = d

    Dim data As Variant: data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function

    Dim cId As Long, cIme As Long, cPr As Long
    cId = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    cIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    cPr = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    If cId = 0 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim k As String: k = Trim$(CStr(data(i, cId)))
        If Len(k) > 0 And Not d.Exists(k) Then
            d.Add k, Trim$(CStr(data(i, cIme)) & " " & CStr(data(i, cPr)))
        End If
    Next i
End Function

Private Function DictVal(ByVal d As Object, ByVal k As String) As String
    If d Is Nothing Then Exit Function
    k = Trim$(k)
    If d.Exists(k) Then DictVal = CStr(d(k))
End Function

' BrojZbirne -> Kupac (firma) naziv; fallback na KupacID ako naziv fali.
Private Function KupacNazivZaZbirnu(ByVal dKupId As Object, ByVal dKupNaziv As Object, _
                                    ByVal brojZbirne As String) As String
    Dim kid As String: kid = DictVal(dKupId, brojZbirne)
    Dim nm As String: nm = DictVal(dKupNaziv, kid)
    If Len(nm) > 0 Then KupacNazivZaZbirnu = nm Else KupacNazivZaZbirnu = kid
End Function

Private Function PdvStopa() As Double
    Dim s As Double
    If Not TryParseDouble(GetConfigValue(CFG_PDV_NADOKNADA_STOPA), s) Then s = 0
    If s <= 0 Then s = PDV_NADOKNADA_DEFAULT
    PdvStopa = s
End Function

Private Function NumVal(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumVal = CDbl(v)
End Function

' Bruto rezim: "bruto (neto X)"; neto rezim ili bruto==neto: samo vrednost.
Private Function FmtKgBrutoNeto(ByVal brutoVal As Double, ByVal netoVal As Double) As String
    If OtkupBrutoUnos() And Abs(brutoVal - netoVal) > 0.0001 Then
        FmtKgBrutoNeto = FmtKgDec(brutoVal) & " (neto " & FmtKgDec(netoVal) & ")"
    Else
        FmtKgBrutoNeto = FmtKgDec(netoVal)
    End If
End Function

Private Function FmtKg(ByVal X As Double) As String
    FmtKg = Format$(X, "#,##0")
End Function

' Kolicina (kg): uvek 2 decimale (npr. 1234.00) -- panel + liste otpremnica/blokova.
' Konvencija ista kao zivi prikaz u frmOtkup.UpdateUkupnoKg ("#,##0.00").
Private Function FmtKgDec(ByVal X As Double) As String
    FmtKgDec = Format$(X, "#,##0.00")
End Function

Private Function FmtRsd(ByVal X As Double) As String
    FmtRsd = Format$(X, "#,##0.00")
End Function

Private Function FmtDate(ByVal v As Variant) As String
    If IsDate(v) Then FmtDate = Format$(CDate(v), "d.m.yyyy")
End Function

' --- dinamicke kontrole + event wiring ---

Private Function AddCtl(ByVal kind As String, ByVal nm As String, _
                        ByVal l As Double, ByVal t As Double, _
                        ByVal w As Double, ByVal h As Double) As Object
    Dim c As Object
    Set c = mForm.Controls.Add("Forms." & kind & ".1", nm, True)
    c.Left = l: c.top = t: c.width = w: c.Height = h
    mPanelCtls.Add c
    Set AddCtl = c
End Function

Private Sub AddHeaders(ByVal prefix As String, ByVal baseLeft As Double, _
                       ByVal top As Double, ByVal widths As String, ByVal caps As String)
    Dim wArr() As String: wArr = Split(widths, ";")
    Dim cArr() As String: cArr = Split(caps, ";")
    Dim X As Double: X = baseLeft
    Dim k As Long
    For k = 0 To UBound(wArr)
        Dim wv As Double: wv = val(wArr(k))
        If wv > 0 Then
            Dim cap As String: cap = ""
            If k <= UBound(cArr) Then cap = cArr(k)
            Dim c As Object
            Set c = AddCtl("Label", prefix & "_" & k, X, top, wv, 26)
            c.caption = cap
            On Error Resume Next
            StyleListHeaderLabel c
            c.WordWrap = True
            On Error GoTo 0
        End If
        X = X + wv
    Next k
End Sub

Private Sub StyleHdr(ByVal c As Object)
    On Error Resume Next
    c.Font.Bold = True
    c.Font.Size = 9
    c.ForeColor = RGB(40, 40, 40)
    On Error GoTo 0
End Sub

Private Sub WireBtn(ByVal b As Object, ByVal act As String)
    Dim w As clsBlokUI: Set w = New clsBlokUI
    w.action = act
    Set w.btn = b
    mWrappers.Add w
End Sub

Private Sub WireTxt(ByVal t As Object, ByVal act As String)
    Dim w As clsBlokUI: Set w = New clsBlokUI
    w.action = act
    Set w.txt = t
    mWrappers.Add w
End Sub

Private Sub WireLst(ByVal l As Object, ByVal act As String)
    Dim w As clsBlokUI: Set w = New clsBlokUI
    w.action = act
    Set w.lst = l
    mWrappers.Add w
End Sub
