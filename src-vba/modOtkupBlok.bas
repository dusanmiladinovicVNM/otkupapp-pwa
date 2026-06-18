Attribute VB_Name = "modOtkupBlok"
Option Explicit

' ============================================================
' modOtkupBlok – Panel "Otkupni blokovi" u frmOtkup.
'
' Panel NE unosi sam u tblOtkup – vodi POSTOJECU levu frmOtkup formu:
'   - Klik na otpremnicu (sredina) popuni levu formu: otkupno mesto,
'     vrsta, sorta, vozac, datum, broj zbirne i cenu; broj otkupnog
'     lista racuna kanonski SuggestNextBroj (OM iz polja + datum otpr.).
'   - Gore: "Cena po otpremnici" (override, vazi za sve blokove) + sazetak
'     Preostalo za unos.
'   - Korisnik unese kooperanta + kolicinu i klikne postojeci "Unos".
'   - frmOtkup.btnUnos_Click pre snimanja zove OtkupBlok_ConfirmUnos
'     (upozorenje na prekoracenje), a posle uspeha OtkupBlok_AfterUnos
'     (vezivanje OtkupID->OtpremnicaID + osvezavanje + auto-deselekcija
'     kad Preostalo padne na 0).
'   - Desno: blokovi izabrane otpremnice (+ zbirni red), sa dugmadima
'     "Storniraj blok" (StornoOtkup_TX) i "Stampaj list" (PrintOtkupniList).
'   - Lista otpremnica: kolona "Preostalo", filter "samo nezavrsene",
'     sort po datumu (najnovije gore).
'
' Sve kontrole panela su dinamicke (Controls.Add) – frmOtkup.frx se ne
' menja. Cena se cuva kao BRUTO (sa PDV nadoknadom); neto/PDV iz nje.
'
' Integracija u frmOtkup:
'   UserForm_Initialize:                 AttachOtkupBlokPanel Me
'   btnUnos_Click (pre snimanja):        If Not OtkupBlok_ConfirmUnos() Then Exit Sub
'   btnUnos_Click (po uspehu, posle ClearOtkupFields): OtkupBlok_AfterUnos result
' ============================================================

' --- Layout (tacke) ---
Private Const PANEL_LEFT  As Double = 312
Private Const OTP_W       As Double = 360
Private Const BLOK_LEFT   As Double = 680       ' PANEL_LEFT + OTP_W + 8
Private Const BLOK_W      As Double = 460
Private Const GRID_TOP    As Double = 88
Private Const EXP_WIDTH   As Double = 1155
Private Const TOGGLE_W    As Double = 130

Private Const OTP_COLW  As String = "0;0;58;38;56;68;44;42;46"
Private Const OTP_CAPS  As String = ";;Otkupno mesto;Kolicina;Datum;Hladnjaca;Prodajna;Cena za;Preostalo"
Private Const BLOK_COLW As String = "0;42;70;58;44;46;66;58;66"
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
Private mFilterNezavrsene As Boolean

Private mBtnToggle As MSForms.CommandButton
Private mBtnStorno As MSForms.CommandButton
Private mBtnPrint As MSForms.CommandButton
Private mBtnFilter As MSForms.CommandButton
Private mLstOtp As MSForms.ListBox
Private mLstBlok As MSForms.ListBox
Private mTxtCenaOtp As MSForms.TextBox
Private mLblUkupno As MSForms.label
Private mLblNapisano As MSForms.label
Private mLblPreostalo As MSForms.label

' ============================================================
' PUBLIC – ulazna tacka + event ruteri + frmOtkup hooks
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
    mOrigWidth = mForm.width

    If UCase$(Trim$(GetConfigValue("OTKUP_BLOK_PANEL"))) = "NO" Then Exit Sub

    Set mBtnToggle = mForm.Controls.Add("Forms.CommandButton.1", "btnOtkBlokToggle", True)
    mBtnToggle.width = TOGGLE_W
    mBtnToggle.Height = 24
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
        Case "STORNO": StornoSelectedBlok
        Case "PRINT": PrintSelectedBlok
        Case "FILTER": ToggleFilter
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

    Dim total As Double: total = kol + k2
    If total <= 0 Then Exit Function

    Dim ukupno As Double
    ukupno = NumVal(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, mActiveOtpID, COL_OTP_KOLICINA))
    Dim preost As Double: preost = ukupno - SumKolByOtp(mActiveOtpID)

    If total > preost + 0.0001 Then
        If MsgBox("Unos (" & FmtKg(total) & " kg) premasuje preostalih " & _
                  FmtKg(preost) & " kg za ovu otpremnicu." & vbCrLf & _
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

    ' cena po otpremnici vazi za SVE blokove; txtCena je obrisao ClearOtkupFields
    If mCenaBlok.Exists(mActiveOtpID) Then
        ApplyCenaToOtpremnica mActiveOtpID, CDbl(mCenaBlok(mActiveOtpID))
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

    ' Naslovi + dugmad (dugmad levo od toggle-a da se ne preklapaju)
    Dim t1 As Object: Set t1 = AddCtl("Label", "lblOtkBlokT1", PANEL_LEFT, 28, 226, 14)
    t1.caption = "OTPREMNICE  (klik = izbor)": StyleHdr t1
    Set mBtnFilter = AddCtl("CommandButton", "btnOtkBlokFilter", PANEL_LEFT + OTP_W - 120, 26, 120, 22)
    mBtnFilter.caption = "Prikaz: Sve"

    Dim t2 As Object: Set t2 = AddCtl("Label", "lblOtkBlokT2", BLOK_LEFT, 28, 118, 14)
    t2.caption = "OTKUPNI BLOKOVI": StyleHdr t2
    Set mBtnStorno = AddCtl("CommandButton", "btnOtkBlokStorno", BLOK_LEFT + 124, 26, 92, 22)
    mBtnStorno.caption = "Storniraj blok"
    Set mBtnPrint = AddCtl("CommandButton", "btnOtkBlokPrint", BLOK_LEFT + 220, 26, 92, 22)
    mBtnPrint.caption = "Stampaj list"

    On Error Resume Next
    StyleExitButton mBtnFilter, "Prikaz: Sve"
    StyleExitButton mBtnStorno, "Storniraj blok"
    StylePrimaryButton mBtnPrint, "Stampaj list"
    On Error GoTo 0

    ' Zaglavlja kolona (spustena dalje od dugmadi; listbox tik ispod)
    AddHeaders "hOtp", PANEL_LEFT, 58, OTP_COLW, OTP_CAPS
    AddHeaders "hBlok", BLOK_LEFT, 58, BLOK_COLW, BLOK_CAPS

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
        mLstOtp.List(r, 3) = FmtKg(ukupno)
        mLstOtp.List(r, 4) = FmtDate(data(i, cDat))
        mLstOtp.List(r, 5) = DictVal(dHl, CStr(data(i, cZbr)))
        mLstOtp.List(r, 6) = FmtKg(prodajna)
        mLstOtp.List(r, 7) = FmtKg(cenaBlok)
        mLstOtp.List(r, 8) = FmtKg(ukupno - nap)
    Next j
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

    Dim sumKol As Double, sumVred As Double, sumPdv As Double, sumUk As Double
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

        sumKol = sumKol + kol: sumVred = sumVred + vred
        sumPdv = sumPdv + pdv: sumUk = sumUk + uk
NextRow:
    Next i

    ' Zbirni red
    If mLstBlok.ListCount > 0 Then
        mLstBlok.AddItem ""                       ' col0 prazan -> storno/print ga preskace
        r = mLstBlok.ListCount - 1
        mLstBlok.List(r, 1) = "UKUPNO"
        mLstBlok.List(r, 4) = FmtKg(sumKol)
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

' Promena cene gore (AfterUpdate) -> propagacija na sve blokove + osvezavanje.
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

    If MsgBox("Stornirati blok " & otkupID & " (" & CStr(mLstBlok.List(li, 2)) & ", " & _
              CStr(mLstBlok.List(li, 4)) & " kg)?", vbQuestion + vbYesNo, APP_NAME) = vbNo Then Exit Sub

    If StornoOtkup_TX(otkupID) Then
        LoadBlokovi
        LoadOtpremnice
        RefreshSummary
        MsgBox "Blok storniran: " & otkupID, vbInformation, APP_NAME
    Else
        MsgBox "Storno nije uspeo za " & otkupID & ".", vbCritical, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "modOtkupBlok.StornoSelectedBlok"
    MsgBox "Greska pri storno bloka: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub PrintSelectedBlok()
    On Error GoTo EH
    Dim li As Long: li = mLstBlok.ListIndex
    If li < 0 Then
        MsgBox "Izaberite blok za stampu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim otkupID As String: otkupID = Trim$(CStr(mLstBlok.List(li, 0)))
    If Len(otkupID) = 0 Then Exit Sub      ' zbirni red

    PrintOtkupniList otkupID
    Exit Sub
EH:
    LogErr "modOtkupBlok.PrintSelectedBlok"
    MsgBox "Greska pri stampi otkupnog lista: " & Err.description, vbCritical, APP_NAME
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
            Set c = AddCtl("Label", prefix & "_" & k, x, top, wv, 26)
            c.caption = cap
            On Error Resume Next
            StyleListHeaderLabel c
            c.WordWrap = True
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
