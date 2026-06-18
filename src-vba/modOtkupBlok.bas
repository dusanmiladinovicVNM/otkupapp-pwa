Attribute VB_Name = "modOtkupBlok"
Option Explicit

' ============================================================
' modOtkupBlok – Opcioni panel "Otkupni blokovi" u frmOtkup.
'
' Aktivira se na dugme (toggle). Default je skriven, pa frmOtkup
' radi 100% kao do sada. Kada se ukljuci, forma se prosiri i
' pojave se: pregled OTPREMNICA (levo) + tabela OTKUPNIH BLOKOVA
' (desno) za izabranu otpremnicu + red za unos + mini-sazetak kg.
'
' Sve kontrole se prave dinamicki (Controls.Add) – frmOtkup.frx
' se NE menja. Eventi dinamickih kontrola idu preko clsBlokUI.
'
' Integracija: u frmOtkup.UserForm_Initialize: AttachOtkupBlokPanel Me
'
' Model:
'  - Svaki blok = tblOtkup red vezan na otpremnicu (OtpremnicaID).
'  - Klik na otpremnicu upise njen broj, broj zbirne, datum i cenu.
'  - CENA je po otpremnici: jednom uneta vazi za SVE blokove te
'    otpremnice (propagira se na sve povezane tblOtkup redove).
'  - Posle "Dodaj blok" automatski se pokrece AutoLink (= rucno
'    "automatski povezi" iz Sledljivosti).
'  - Cena se cuva kao BRUTO (sa PDV nadoknadom) – isto kao postojeci
'    otkup; neto / PDV se racunaju iz nje.
' ============================================================

' --- Layout (tacke; doteraj po ekranu) ---
Private Const PANEL_LEFT  As Double = 312
Private Const OTP_W       As Double = 360
Private Const BLOK_LEFT   As Double = 680       ' PANEL_LEFT + OTP_W + 8
Private Const BLOK_W      As Double = 460
Private Const GRID_TOP    As Double = 92
Private Const EXP_WIDTH   As Double = 1155       ' BLOK_LEFT + BLOK_W + 15
Private Const TOGGLE_W    As Double = 130

Private Const OTP_COLW  As String = "0;0;72;40;48;100;48;52"
Private Const OTP_CAPS  As String = ";;Otkupno mesto;Kolicina;Datum;Hladnjaca;Prodajna;Cena za blok"
Private Const BLOK_COLW As String = "0;42;98;48;44;52;56;50;58"
Private Const BLOK_CAPS As String = ";br. bloka;Ime i Prezime;Datum;Kolicina;Cena bez PDV;Vrednost;Iznos PDV;Ukupna vrednost"

' --- Stanje (modul-level; jedna frmOtkup instanca po sekciji) ---
Private mForm As Object
Private mWrappers As Collection
Private mPanelCtls As Collection
Private mCenaBlok As Object           ' OtpremnicaID -> cena (bruto)
Private mBuilt As Boolean
Private mVisible As Boolean
Private mSuppress As Boolean
Private mOrigWidth As Double
Private mActiveOtpID As String

Private mBtnToggle As MSForms.CommandButton
Private mBtnDodaj As MSForms.CommandButton
Private mLstOtp As MSForms.ListBox
Private mLstBlok As MSForms.ListBox
Private mTxtId As MSForms.TextBox
Private mTxtBrZbirne As MSForms.TextBox
Private mTxtDatum As MSForms.TextBox
Private mCmbKoop As MSForms.ComboBox
Private mTxtKol As MSForms.TextBox
Private mTxtCena As MSForms.TextBox
Private mTxtBrBlok As MSForms.TextBox
Private mLblUkupno As MSForms.label
Private mLblNapisano As MSForms.label
Private mLblPreostalo As MSForms.label

' ============================================================
' PUBLIC – ulazna tacka + event ruteri (zove ih clsBlokUI)
' ============================================================

Public Sub AttachOtkupBlokPanel(ByVal frm As Object)
    On Error GoTo EH

    Set mForm = frm
    Set mWrappers = New Collection
    Set mCenaBlok = CreateObject("Scripting.Dictionary")
    mBuilt = False
    mVisible = False
    mSuppress = False
    mActiveOtpID = ""
    mOrigWidth = mForm.width

    If UCase$(Trim$(GetConfigValue("OTKUP_BLOK_PANEL"))) = "NO" Then Exit Sub

    Set mBtnToggle = mForm.Controls.Add("Forms.CommandButton.1", "btnOtkBlokToggle", True)
    mBtnToggle.width = TOGGLE_W
    mBtnToggle.Height = 20
    mBtnToggle.Top = 6
    mBtnToggle.Left = mForm.InsideWidth - TOGGLE_W - 6
    mBtnToggle.caption = "Otkupni blokovi  »"
    On Error Resume Next
    StylePrimaryButton mBtnToggle, "Otkupni blokovi  »"
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
        Case "DODAJ": DodajBlok
    End Select
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_OnButton"
End Sub

Public Sub OtkupBlok_OnText(ByVal action As String)
    On Error GoTo EH
    If action = "IDBROJ" Then RefreshFromIdBroj
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
        mBtnToggle.caption = "«  Sakrij blokove"
    Else
        mForm.width = mOrigWidth
        mBtnToggle.caption = "Otkupni blokovi  »"
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

    ' Mini-sazetak (gore)
    Set mLblUkupno = AddCtl("Label", "lblOtkBlokUk", PANEL_LEFT, 4, 160, 14)
    Set mLblNapisano = AddCtl("Label", "lblOtkBlokNap", PANEL_LEFT + 166, 4, 160, 14)
    Set mLblPreostalo = AddCtl("Label", "lblOtkBlokPre", PANEL_LEFT + 332, 4, 160, 14)
    mLblUkupno.caption = "Ukupno kg: —"
    mLblNapisano.caption = "U blokovima: —"
    mLblPreostalo.caption = "Preostalo: —"

    ' Red za unos (sve u jednom redu, preko cele sirine panela)
    AddMicro "lblOtkE1", PANEL_LEFT, 22, 46, "Otpremnica br."
    AddMicro "lblOtkE2", PANEL_LEFT + 50, 22, 74, "Broj zbirne"
    AddMicro "lblOtkE3", PANEL_LEFT + 128, 22, 60, "Datum"
    AddMicro "lblOtkE4", PANEL_LEFT + 192, 22, 150, "Ime i Prezime (kooperant)"
    AddMicro "lblOtkE5", PANEL_LEFT + 346, 22, 46, "Kolicina"
    AddMicro "lblOtkE6", PANEL_LEFT + 396, 22, 50, "Cena (otpremnica)"
    AddMicro "lblOtkE7", PANEL_LEFT + 450, 22, 56, "Br. bloka"

    Set mTxtId = AddCtl("TextBox", "txtOtkBlokId", PANEL_LEFT, 38, 46, 18)
    Set mTxtBrZbirne = AddCtl("TextBox", "txtOtkBlokZbr", PANEL_LEFT + 50, 38, 74, 18)
    Set mTxtDatum = AddCtl("TextBox", "txtOtkBlokDat", PANEL_LEFT + 128, 38, 60, 18)
    Set mCmbKoop = AddCtl("ComboBox", "cmbOtkBlokKoop", PANEL_LEFT + 192, 38, 150, 18)
    Set mTxtKol = AddCtl("TextBox", "txtOtkBlokKol", PANEL_LEFT + 346, 38, 46, 18)
    Set mTxtCena = AddCtl("TextBox", "txtOtkBlokCena", PANEL_LEFT + 396, 38, 50, 18)
    Set mTxtBrBlok = AddCtl("TextBox", "txtOtkBlokBr", PANEL_LEFT + 450, 38, 56, 18)
    Set mBtnDodaj = AddCtl("CommandButton", "btnOtkBlokDodaj", PANEL_LEFT + 510, 37, 84, 20)
    mBtnDodaj.caption = "Dodaj blok"

    On Error Resume Next
    mCmbKoop.MatchEntry = fmMatchEntryComplete
    StyleComboBox mCmbKoop
    StyleTextBox mTxtId: StyleTextBox mTxtBrZbirne: StyleTextBox mTxtDatum
    StyleTextBox mTxtKol: StyleTextBox mTxtCena: StyleTextBox mTxtBrBlok
    StylePrimaryButton mBtnDodaj, "Dodaj blok"
    On Error GoTo 0
    LockField mTxtBrZbirne          ' izvedeno iz otpremnice – samo prikaz
    LockField mTxtDatum

    ' Naslovi
    Dim t1 As Object, t2 As Object
    Set t1 = AddCtl("Label", "lblOtkBlokT1", PANEL_LEFT, 60, OTP_W, 14)
    t1.caption = "OTPREMNICE  (klik = izbor)": StyleHdr t1
    Set t2 = AddCtl("Label", "lblOtkBlokT2", BLOK_LEFT, 60, BLOK_W, 14)
    t2.caption = "OTKUPNI BLOKOVI  (za izabranu otpremnicu)": StyleHdr t2

    ' Zaglavlja kolona
    AddHeaders "hOtp", PANEL_LEFT, 76, OTP_COLW, OTP_CAPS
    AddHeaders "hBlok", BLOK_LEFT, 76, BLOK_COLW, BLOK_CAPS

    ' Grid-ovi
    Set mLstOtp = AddCtl("ListBox", "lstOtkBlokOtp", PANEL_LEFT, GRID_TOP, OTP_W, gridH)
    mLstOtp.ColumnCount = 8
    mLstOtp.ColumnWidths = OTP_COLW

    Set mLstBlok = AddCtl("ListBox", "lstOtkBlokBlok", BLOK_LEFT, GRID_TOP, BLOK_W, gridH)
    mLstBlok.ColumnCount = 9
    mLstBlok.ColumnWidths = BLOK_COLW

    ' Eventi
    WireTxt mTxtId, "IDBROJ"
    WireTxt mTxtCena, "CENA"
    WireLst mLstOtp, "OTP"
    WireBtn mBtnDodaj, "DODAJ"
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
' LOAD – pregled otpremnica (levo) + blokovi izabrane otpremnice (desno)
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
    Dim dHl As Object: Set dHl = BuildLookup(TBL_ZBIRNA, COL_ZBR_BROJ, COL_ZBR_HLADNJACA)
    Dim dCe As Object: Set dCe = BuildFirstBlokCena()      ' OtpremnicaID -> cena prvog bloka

    Dim i As Long, r As Long
    For i = 1 To UBound(data, 1)
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

        mLstOtp.AddItem otpID
        r = mLstOtp.ListCount - 1
        mLstOtp.List(r, 1) = CStr(data(i, cBroj))
        mLstOtp.List(r, 2) = DictVal(dSt, CStr(data(i, cSt)))
        mLstOtp.List(r, 3) = FmtKg(NumVal(data(i, cKol)))
        mLstOtp.List(r, 4) = FmtDate(data(i, cDat))
        mLstOtp.List(r, 5) = DictVal(dHl, CStr(data(i, cZbr)))
        mLstOtp.List(r, 6) = FmtKg(prodajna)
        mLstOtp.List(r, 7) = FmtKg(cenaBlok)
    Next i
    Exit Sub
EH:
    LogErr "modOtkupBlok.LoadOtpremnice"
End Sub

Private Sub LoadBlokovi()
    On Error GoTo EH
    mLstBlok.Clear
    If Len(mActiveOtpID) = 0 Then Exit Sub      ' prikaz po izabranoj otpremnici

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub

    Dim cId As Long, cOtp As Long, cKoop As Long, cKol As Long
    Dim cCena As Long, cBr As Long, cDat As Long
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)

    Dim dKo As Object: Set dKo = BuildKoopNames()
    Dim stopa As Double: stopa = PdvStopa()

    Dim i As Long, r As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cOtp))) <> mActiveOtpID Then GoTo NextRow

        Dim kol As Double: kol = NumVal(data(i, cKol))
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
        mLstBlok.List(r, 4) = FmtKg(kol)
        mLstBlok.List(r, 5) = FmtRsd(neto)
        mLstBlok.List(r, 6) = FmtRsd(vred)
        mLstBlok.List(r, 7) = FmtRsd(pdv)
        mLstBlok.List(r, 8) = FmtRsd(uk)
NextRow:
    Next i
    Exit Sub
EH:
    LogErr "modOtkupBlok.LoadBlokovi"
End Sub

' ============================================================
' INTERAKCIJA
' ============================================================

Private Sub SelectOtpFromList()
    On Error GoTo EH
    If mLstOtp.ListIndex < 0 Then Exit Sub

    Dim r As Long: r = mLstOtp.ListIndex
    mActiveOtpID = CStr(mLstOtp.List(r, 0))
    Dim broj As String: broj = CStr(mLstOtp.List(r, 1))

    mSuppress = True
    mTxtId.value = broj
    mSuppress = False

    FillOtpDisplayFields mActiveOtpID
    LoadBlokovi
    RefreshSummary broj
    Exit Sub
EH:
    LogErr "modOtkupBlok.SelectOtpFromList"
End Sub

Private Sub RefreshFromIdBroj()
    On Error GoTo EH
    If mSuppress Then Exit Sub

    Dim broj As String: broj = Trim$(mTxtId.value)
    mActiveOtpID = OtpIdFromBroj(broj)
    If Len(mActiveOtpID) > 0 Then FillOtpDisplayFields mActiveOtpID

    LoadBlokovi
    RefreshSummary broj
    Exit Sub
EH:
    LogErr "modOtkupBlok.RefreshFromIdBroj"
End Sub

' Klik/izbor otpremnice upisuje broj zbirne, datum i cenu u formu bloka,
' i puni kooperante te stanice. Cena = cena vec uneta za otpremnicu
' (iz postojecih blokova), inace prodajna cena otpremnice.
Private Sub FillOtpDisplayFields(ByVal otpID As String)
    On Error GoTo EH
    If Len(otpID) = 0 Then Exit Sub

    Dim stanicaID As String
    stanicaID = Trim$(CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA)))
    FillComboKooperantiByStanica mCmbKoop, stanicaID

    mTxtBrZbirne.value = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE))
    mTxtDatum.value = FmtDate(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_DATUM))

    Dim cena As Double
    If mCenaBlok.Exists(otpID) Then
        cena = mCenaBlok(otpID)
    Else
        cena = ExistingBlokCena(otpID)
        If cena <= 0 Then cena = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_CENA))
        mCenaBlok(otpID) = cena
    End If
    mTxtCena.value = Format$(cena, "0.####")
    Exit Sub
EH:
    LogErr "modOtkupBlok.FillOtpDisplayFields"
End Sub

Private Sub RefreshSummary(ByVal brojOtp As String)
    On Error GoTo EH

    Dim otpID As String: otpID = OtpIdFromBroj(brojOtp)
    If Len(otpID) = 0 Then
        mLblUkupno.caption = "Ukupno kg: —"
        mLblNapisano.caption = "U blokovima: —"
        mLblPreostalo.caption = "Preostalo: —"
        Exit Sub
    End If

    Dim ukupno As Double
    ukupno = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_KOLICINA))
    Dim napisano As Double: napisano = SumKolByOtp(otpID)
    Dim preostalo As Double: preostalo = ukupno - napisano

    mLblUkupno.caption = "Ukupno kg: " & FmtKg(ukupno)
    mLblNapisano.caption = "U blokovima: " & FmtKg(napisano)
    mLblPreostalo.caption = "Preostalo: " & FmtKg(preostalo)

    On Error Resume Next
    If preostalo < -0.0001 Then
        mLblPreostalo.ForeColor = RGB(200, 0, 0)
    Else
        mLblPreostalo.ForeColor = RGB(0, 120, 0)
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.RefreshSummary"
End Sub

' Promena cene u polju -> cena vazi za celu otpremnicu (sve blokove)
Private Sub OnCenaChanged()
    On Error GoTo EH
    If Len(mActiveOtpID) = 0 Then Exit Sub
    Dim cena As Double
    If Not TryParseDouble(mTxtCena.value, cena) Or cena <= 0 Then Exit Sub

    mCenaBlok(mActiveOtpID) = cena
    ApplyCenaToOtpremnica mActiveOtpID, cena
    LoadBlokovi
    LoadOtpremnice
    RefreshSummary Trim$(mTxtId.value)
    Exit Sub
EH:
    LogErr "modOtkupBlok.OnCenaChanged"
End Sub

Private Sub DodajBlok()
    On Error GoTo EH

    Dim broj As String: broj = Trim$(mTxtId.value)
    Dim otpID As String: otpID = OtpIdFromBroj(broj)
    If Len(otpID) = 0 Then
        MsgBox "Unesite ili izaberite ispravan broj otpremnice.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim koopID As String: koopID = ExtractIDFromDisplay(Trim$(mCmbKoop.value))
    If Len(koopID) = 0 Then
        MsgBox "Izaberite kooperanta (ime i prezime).", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim kol As Double
    If Not TryParseDouble(mTxtKol.value, kol) Or kol <= 0 Then
        MsgBox "Unesite ispravnu kolicinu.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim cena As Double
    If Not TryParseDouble(mTxtCena.value, cena) Or cena <= 0 Then
        MsgBox "Unesite ispravnu cenu za otpremnicu.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim brBlok As String: brBlok = Trim$(mTxtBrBlok.value)

    Dim ukupno As Double
    ukupno = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_KOLICINA))
    Dim preostalo As Double: preostalo = ukupno - SumKolByOtp(otpID)
    If kol > preostalo + 0.0001 Then
        If MsgBox("Kolicina (" & FmtKg(kol) & " kg) je veca od preostalih " & _
                  FmtKg(preostalo) & " kg za ovu otpremnicu." & vbCrLf & _
                  "Nastaviti?", vbExclamation + vbYesNo, APP_NAME) = vbNo Then Exit Sub
    End If

    mCenaBlok(otpID) = cena

    Dim newID As String: newID = SaveOtkupBlok(otpID, koopID, kol, cena, brBlok)
    If Len(newID) = 0 Then Exit Sub      ' greska je vec prijavljena

    ApplyCenaToOtpremnica otpID, cena    ' cena vazi za SVE blokove otpremnice

    On Error Resume Next
    AutoLinkOtkupOtpremnica_TX           ' = rucno "automatski povezi" iz Sledljivosti
    On Error GoTo EH

    mActiveOtpID = otpID
    LoadBlokovi
    LoadOtpremnice
    RefreshSummary broj
    mTxtKol.value = ""
    mTxtBrBlok.value = ""

    MsgBox "Otkupni blok dodat: " & newID, vbInformation, APP_NAME
    On Error Resume Next
    mTxtKol.SetFocus
    Exit Sub
EH:
    LogErr "modOtkupBlok.DodajBlok"
    MsgBox "Greska pri dodavanju bloka: " & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' UPIS – lean tblOtkup red + OtpremnicaID
' ============================================================

Private Function SaveOtkupBlok(ByVal otpID As String, ByVal koopID As String, _
                               ByVal kolicina As Double, ByVal cenaBruto As Double, _
                               ByVal brBlok As String) As String
    Dim tx As clsTransaction
    On Error GoTo EH

    Dim datum As Date
    Dim vDat As Variant: vDat = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_DATUM)
    If IsDate(vDat) Then datum = CDate(vDat) Else datum = Date

    Dim stanicaID As String: stanicaID = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA))
    Dim vrsta As String: vrsta = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_VRSTA))
    Dim sorta As String: sorta = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_SORTA))
    Dim brZbr As String: brZbr = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE))
    Dim klasa As String: klasa = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_KLASA))
    If Len(Trim$(klasa)) = 0 Or InStr(klasa, "/") > 0 Then klasa = KLASA_I

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim newID As String
    newID = SaveOtkup(datum:=datum, kooperantID:=koopID, stanicaID:=stanicaID, _
                      vrstaVoca:=vrsta, sortaVoca:=sorta, kolicina:=kolicina, _
                      cena:=cenaBruto, tipAmb:="", kolAmb:=0, vozacID:="", _
                      brDok:=brBlok, novac:=0, primalac:="", klasa:=klasa, _
                      parcelaID:="", brojZbirne:=brZbr)

    If Len(Trim$(newID)) = 0 Then
        Err.Raise vbObjectError + 2010, "modOtkupBlok.SaveOtkupBlok", "SaveOtkup nije vratio ID."
    End If

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, newID)
    If rows.count = 0 Then
        Err.Raise vbObjectError + 2011, "modOtkupBlok.SaveOtkupBlok", _
                  "Otkup red nije pronaden: " & newID
    End If

    RequireUpdateCell TBL_OTKUP, rows(1), COL_OTK_OTPREMNICA_ID, otpID, "modOtkupBlok.SaveOtkupBlok"

    tx.CommitTx
    Set tx = Nothing

    SaveOtkupBlok = newID
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "modOtkupBlok.SaveOtkupBlok"
    MsgBox "Greska pri upisu bloka: " & Err.description, vbCritical, APP_NAME
    SaveOtkupBlok = ""
End Function

' Cena po otpremnici: postavi istu cenu na SVE tblOtkup redove te otpremnice.
Private Sub ApplyCenaToOtpremnica(ByVal otpID As String, ByVal cena As Double)
    Dim tx As clsTransaction
    On Error GoTo EH
    If Len(otpID) = 0 Then Exit Sub

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, otpID)
    If rows.count = 0 Then Exit Sub

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim k As Long
    For k = 1 To rows.count
        RequireUpdateCell TBL_OTKUP, rows(k), COL_OTK_CENA, cena, "modOtkupBlok.ApplyCenaToOtpremnica"
    Next k

    tx.CommitTx
    Set tx = Nothing
    Exit Sub
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr "modOtkupBlok.ApplyCenaToOtpremnica"
End Sub

' ============================================================
' HELPERS
' ============================================================

Private Function OtpIdFromBroj(ByVal broj As String) As String
    broj = Trim$(broj)
    If Len(broj) = 0 Then Exit Function
    OtpIdFromBroj = Trim$(CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, broj, COL_OTP_ID)))
End Function

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

' Cena vec upotrebljena za otpremnicu (iz prvog povezanog bloka), inace 0.
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

' Jedan prolaz kroz tblOtkup: OtpremnicaID -> cena prvog povezanog bloka.
' (Da LoadOtpremnice ne radi FindRows po svakom redu.)
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

Private Function PdvStopa() As Double
    Dim s As Double
    If Not TryParseDouble(GetConfigValue(CFG_PDV_NADOKNADA_STOPA), s) Then s = 0
    If s <= 0 Then s = PDV_NADOKNADA_DEFAULT
    PdvStopa = s
End Function

Private Function NumVal(ByVal v As Variant) As Double
    If IsNumeric(v) Then NumVal = CDbl(v)
End Function

Private Function FmtKg(ByVal x As Double) As String
    FmtKg = Format$(x, "#,##0")
End Function

Private Function FmtRsd(ByVal x As Double) As String
    FmtRsd = Format$(x, "#,##0.00")
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
    c.Left = l: c.Top = t: c.width = w: c.Height = h
    mPanelCtls.Add c
    Set AddCtl = c
End Function

Private Sub AddMicro(ByVal nm As String, ByVal l As Double, ByVal t As Double, _
                     ByVal w As Double, ByVal cap As String)
    Dim c As Object: Set c = AddCtl("Label", nm, l, t, w, 12)
    c.caption = cap
    On Error Resume Next
    c.Font.Size = 8
    c.ForeColor = RGB(90, 90, 90)
    On Error GoTo 0
End Sub

Private Sub AddHeaders(ByVal prefix As String, ByVal baseLeft As Double, _
                       ByVal top As Double, ByVal widths As String, ByVal caps As String)
    Dim wArr() As String: wArr = Split(widths, ";")
    Dim cArr() As String: cArr = Split(caps, ";")
    Dim x As Double: x = baseLeft
    Dim k As Long
    For k = 0 To UBound(wArr)
        Dim wv As Double: wv = Val(wArr(k))
        If wv > 0 Then
            Dim cap As String: cap = ""
            If k <= UBound(cArr) Then cap = cArr(k)
            Dim c As Object
            Set c = AddCtl("Label", prefix & "_" & k, x, top, wv, 14)
            c.caption = cap
            On Error Resume Next
            StyleListHeaderLabel c
            On Error GoTo 0
        End If
        x = x + wv
    Next k
End Sub

Private Sub StyleHdr(ByVal c As Object)
    On Error Resume Next
    c.Font.Bold = True
    c.Font.Size = 9
    c.ForeColor = RGB(40, 40, 40)
    On Error GoTo 0
End Sub

Private Sub LockField(ByVal t As Object)
    On Error Resume Next
    t.Locked = True
    t.TabStop = False
    t.BackColor = RGB(238, 238, 238)
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
