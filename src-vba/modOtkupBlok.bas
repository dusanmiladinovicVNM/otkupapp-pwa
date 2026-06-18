Attribute VB_Name = "modOtkupBlok"
Option Explicit

' ============================================================
' modOtkupBlok – Panel "Otkupni blokovi" u frmOtkup.
'
' Panel NE unosi sam u tblOtkup. Umesto toga vodi POSTOJECU levu
' frmOtkup formu:
'   - Klik na otpremnicu (sredina) popuni levu formu: otkupno mesto,
'     vrsta, sorta, vozac, datum, broj zbirne i cenu.
'   - Gore se unosi samo CENA po otpremnici (vazi za sve blokove te
'     otpremnice) + sazetak Preostalo za unos.
'   - Korisnik unese kooperanta + kolicinu i klikne postojeci "Unos".
'   - Posle "Unos" (frmOtkup poziva OtkupBlok_AfterUnos) upravo
'     sacuvani red(ovi) se VEZU za izabranu otpremnicu (OtpremnicaID),
'     pa se Preostalo i desna lista odmah azuriraju – bez rucnog
'     "automatski povezi" iz Sledljivosti.
'   - Skroz desno: lista uradjenih blokova izabrane otpremnice.
'
' Sve kontrole panela su dinamicke (Controls.Add) – frmOtkup.frx se
' ne menja. Cena se cuva kao BRUTO (sa PDV nadoknadom); neto/PDV se
' racunaju iz nje.
'
' Integracija u frmOtkup:
'   UserForm_Initialize:  AttachOtkupBlokPanel Me
'   btnUnos_Click (po uspehu, posle ClearOtkupFields):
'                         OtkupBlok_AfterUnos result
' + importovati modOtkupBlok.bas i clsBlokUI.cls
' ============================================================

' --- Layout (tacke; doteraj po ekranu) ---
Private Const PANEL_LEFT  As Double = 312
Private Const OTP_W       As Double = 360
Private Const BLOK_LEFT   As Double = 680       ' PANEL_LEFT + OTP_W + 8
Private Const BLOK_W      As Double = 460
Private Const GRID_TOP    As Double = 62
Private Const EXP_WIDTH   As Double = 1155
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
Private mOrigWidth As Double
Private mActiveOtpID As String

Private mBtnToggle As MSForms.CommandButton
Private mLstOtp As MSForms.ListBox
Private mLstBlok As MSForms.ListBox
Private mTxtCenaOtp As MSForms.TextBox
Private mLblUkupno As MSForms.label
Private mLblNapisano As MSForms.label
Private mLblPreostalo As MSForms.label

' ============================================================
' PUBLIC – ulazna tacka + event ruteri (zove ih clsBlokUI) + AfterUnos
' ============================================================

Public Sub AttachOtkupBlokPanel(ByVal frm As Object)
    On Error GoTo EH

    Set mForm = frm
    Set mWrappers = New Collection
    Set mCenaBlok = CreateObject("Scripting.Dictionary")
    mBuilt = False
    mVisible = False
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
    If action = "TOGGLE" Then TogglePanel
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_OnButton"
End Sub

Public Sub OtkupBlok_OnText(ByVal action As String)
    ' rezervisano (trenutno nema Change-vezanih polja)
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

Public Sub OtkupBlok_OnComboChange(ByVal action As String)
    On Error GoTo EH
    If action = "OMCHANGE" Then OnOmChanged
    Exit Sub
EH:
    LogErr "modOtkupBlok.OtkupBlok_OnComboChange"
End Sub

' Poziva frmOtkup.btnUnos_Click po uspesnom cuvanju (result = OtkupID-jevi
' spojeni sa " + "). Vezuje ih za izabranu otpremnicu i osvezava panel.
Public Sub OtkupBlok_AfterUnos(ByVal otkupIDs As String)
    On Error GoTo EH
    If Not mVisible Then Exit Sub
    If Len(mActiveOtpID) = 0 Then Exit Sub

    LinkOtkupIDsToOtpremnica otkupIDs, mActiveOtpID

    ' ClearOtkupFields je obrisao txtCena – cena je po otpremnici, vrati je
    If mCenaBlok.Exists(mActiveOtpID) Then
        SetLeftCtl "txtCena", Format$(CDbl(mCenaBlok(mActiveOtpID)), "0.00")
    End If

    ' Broj otkupnog lista: sledeci redni broj za OM (iz polja) + datum otpremnice
    Dim brDok As String: brDok = OtpBrojDok(mActiveOtpID)
    If Len(brDok) > 0 Then SetLeftCtl "txtBrojDokumenta", brDok

    LoadBlokovi
    LoadOtpremnice
    RefreshSummary
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
    mLblUkupno.caption = "Ukupno kg: —"
    mLblNapisano.caption = "U blokovima: —"
    mLblPreostalo.caption = "Preostalo: —"

    ' Naslovi
    Dim t1 As Object, t2 As Object
    Set t1 = AddCtl("Label", "lblOtkBlokT1", PANEL_LEFT, 30, OTP_W, 14)
    t1.caption = "OTPREMNICE  (klik = izbor; puni levu formu)": StyleHdr t1
    Set t2 = AddCtl("Label", "lblOtkBlokT2", BLOK_LEFT, 30, BLOK_W, 14)
    t2.caption = "OTKUPNI BLOKOVI  (izabrane otpremnice)": StyleHdr t2

    ' Zaglavlja kolona
    AddHeaders "hOtp", PANEL_LEFT, 46, OTP_COLW, OTP_CAPS
    AddHeaders "hBlok", BLOK_LEFT, 46, BLOK_COLW, BLOK_CAPS

    ' Grid-ovi
    Set mLstOtp = AddCtl("ListBox", "lstOtkBlokOtp", PANEL_LEFT, GRID_TOP, OTP_W, gridH)
    mLstOtp.ColumnCount = 8
    mLstOtp.ColumnWidths = OTP_COLW

    Set mLstBlok = AddCtl("ListBox", "lstOtkBlokBlok", BLOK_LEFT, GRID_TOP, BLOK_W, gridH)
    mLstBlok.ColumnCount = 9
    mLstBlok.ColumnWidths = BLOK_COLW

    ' Eventi
    WireTxt mTxtCenaOtp, "CENA"
    WireLst mLstOtp, "OTP"

    ' osvezi broj otkupnog lista kad se promeni "Otkupno mesto" u levoj formi
    On Error Resume Next
    WireCmb mForm.Controls("cmbOtkupnoMesto"), "OMCHANGE"
    On Error GoTo 0
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
' LOAD – pregled otpremnica (sredina) + blokovi izabrane otpremnice (desno)
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
    Dim dCe As Object: Set dCe = BuildFirstBlokCena()

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
    If Len(mActiveOtpID) = 0 Then Exit Sub

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

' Klik na otpremnicu -> popuni levu frmOtkup formu + cenu i sazetak.
Private Sub SelectOtpFromList()
    On Error GoTo EH
    If mLstOtp.ListIndex < 0 Then Exit Sub

    mActiveOtpID = CStr(mLstOtp.List(mLstOtp.ListIndex, 0))

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
    Exit Sub
EH:
    LogErr "modOtkupBlok.SelectOtpFromList"
End Sub

' Popunjava POSTOJECU levu formu podacima otpremnice. cmbOtkupnoMesto i
' cmbVrstaVoca preko svojih _Change dogadjaja pune kooperante / sorte.
Private Sub PrefillLeftForm(ByVal otpID As String, ByVal cena As Double)
    On Error Resume Next

    SetComboByIdAny mForm.Controls("cmbOtkupnoMesto"), _
                    CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA))
    mForm.Controls("cmbVrstaVoca").value = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_VRSTA))
    mForm.Controls("cmbSortaVoca").value = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_SORTA))
    SetComboByIdAny mForm.Controls("cmbVozac"), _
                    CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_VOZAC))

    SetLeftCtl "txtBrojZbirne", CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ_ZBIRNE))
    Dim vDat As Variant: vDat = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_DATUM)
    If IsDate(vDat) Then SetLeftCtl "txtDatum", Format$(CDate(vDat), "d.m.yyyy")
    SetLeftCtl "txtCena", Format$(cena, "0.00")

    ' Broj otkupnog lista po OM + datumu OTPREMNICE (ne danasnji datum)
    Dim brDok As String: brDok = OtpBrojDok(otpID)
    If Len(brDok) > 0 Then SetLeftCtl "txtBrojDokumenta", brDok
End Sub

' Promena cene gore -> vazi za celu otpremnicu (sve blokove) + leva forma.
Private Sub OnCenaChanged()
    On Error GoTo EH
    If Len(mActiveOtpID) = 0 Then Exit Sub
    Dim cena As Double
    If Not TryParseDouble(mTxtCenaOtp.value, cena) Or cena <= 0 Then Exit Sub

    mCenaBlok(mActiveOtpID) = cena
    mTxtCenaOtp.value = Format$(cena, "0.00")
    SetLeftCtl "txtCena", Format$(cena, "0.00")
    ApplyCenaToOtpremnica mActiveOtpID, cena
    LoadBlokovi
    LoadOtpremnice
    RefreshSummary
    Exit Sub
EH:
    LogErr "modOtkupBlok.OnCenaChanged"
End Sub

' Promena "Otkupno mesto" u levoj formi -> osvezi broj otkupnog lista
' (OM prati polje; datum ostaje iz otpremnice).
Private Sub OnOmChanged()
    On Error GoTo EH
    If Not mVisible Then Exit Sub
    If Len(mActiveOtpID) = 0 Then Exit Sub
    Dim brDok As String: brDok = OtpBrojDok(mActiveOtpID)
    If Len(brDok) > 0 Then SetLeftCtl "txtBrojDokumenta", brDok
    Exit Sub
EH:
    LogErr "modOtkupBlok.OnOmChanged"
End Sub

Private Sub RefreshSummary()
    On Error GoTo EH

    If Len(mActiveOtpID) = 0 Then
        mLblUkupno.caption = "Ukupno kg: —"
        mLblNapisano.caption = "U blokovima: —"
        mLblPreostalo.caption = "Preostalo: —"
        Exit Sub
    End If

    Dim ukupno As Double
    ukupno = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_KOLICINA))
    Dim napisano As Double: napisano = SumKolByOtp(mActiveOtpID)
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

' ============================================================
' VEZIVANJE + CENA
' ============================================================

' Vezi tacno date OtkupID-jeve (iz btnUnos result-a) za otpremnicu.
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

' Cena po otpremnici: postavi istu cenu na SVE tblOtkup redove otpremnice.
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

Private Sub SetLeftCtl(ByVal nm As String, ByVal val As String)
    On Error Resume Next
    mForm.Controls(nm).value = val
    On Error GoTo 0
End Sub

' Postavi combo na vrednost po ID-u, radi i za 2-kolonske (bound ID) i za
' 1-kolonske combo-e oblika "Ime Prezime (ID)" / "ID - Ime" (npr. cmbVozac).
Private Sub SetComboByIdAny(ByVal cmb As Object, ByVal idValue As String)
    On Error Resume Next
    idValue = Trim$(idValue)
    If Len(idValue) = 0 Then Exit Sub

    If SetComboByID(cmb, idValue) Then Exit Sub      ' 2-kolonski combo
    If cmb.ListIndex >= 0 Then Exit Sub

    Dim i As Long                                     ' 1-kolonski "... (ID)"
    For i = 0 To cmb.ListCount - 1
        If ExtractIDFromDisplay(CStr(cmb.List(i))) = idValue Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

' Broj otkupnog lista za izabranu otpremnicu: OM iz polja "Otkupno mesto"
' (uskladjen sa onim sto ce se sacuvati), DATUM iz otpremnice. Redni broj
' (prvi bez sufiksa, ostali -N) racuna GenerateBrojDokumenta iz tblOtkup.
Private Function OtpBrojDok(ByVal otpID As String) As String
    On Error Resume Next

    Dim stanicaID As String
    stanicaID = GetComboID(mForm.Controls("cmbOtkupnoMesto"))
    If Len(Trim$(stanicaID)) = 0 Then _
        stanicaID = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_STANICA))

    Dim vDat As Variant: vDat = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_DATUM)
    If Len(Trim$(stanicaID)) = 0 Or Not IsDate(vDat) Then Exit Function

    OtpBrojDok = GenerateBrojDokumenta(stanicaID, CDate(vDat))
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

' Jedan prolaz: OtpremnicaID -> cena prvog povezanog bloka.
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

Private Sub WireCmb(ByVal c As Object, ByVal act As String)
    Dim w As clsBlokUI: Set w = New clsBlokUI
    w.action = act
    Set w.cmb = c
    mWrappers.Add w
End Sub
