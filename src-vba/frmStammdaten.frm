VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStammdaten 
   Caption         =   "UserForm1"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmStammdaten.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStammdaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ============================================================
' frmStammdaten - Universelles Stammdaten-Form
' Wird ueber .Tag gesteuert: "Kooperanti", "Stanice", "Kupci", "Vozaci"
' ============================================================

Private m_TableName As String
Private m_Headers As Variant
Private m_FieldCount As Long
Private m_SelectedRow As Long
Private m_SetupDone As Boolean

Private m_RowMap() As Long
Private m_RowMapCount As Long

Private mChromeRemoved As Boolean

Private mGeoClearConfirmPending As Boolean

' Runtime dugme "Deaktiviraj/Aktiviraj" (soft-delete) -- WithEvents omotac.
Private m_softWrap As clsStmBtn

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

Private Sub UserForm_Initialize()

    ' Nichts hier - Tag ist noch nicht verfuegbar
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH
    MouseWheel_Attach Me

    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me

    If m_SetupDone Then Exit Sub
    m_SetupDone = True

    ' Podesavanja: ne ucitavamo maticni-podaci listu -- gradimo config editor
    ' u runtime-u (modPodesavanja; isti runtime-controls pristup kao
    ' modOtkupBlok/clsBlokUI). frmStammdaten.frx se NE dira.
    If Me.Tag = "Pode" & ChrW(353) & "avanja" Then
        modPodesavanja.BuildConfigEditor Me
        If Not mChromeRemoved Then
            Me.caption = ""             ' brisi caption
            RemoveTitleBar               ' onda ga sakri
            mChromeRemoved = True
        End If
        Exit Sub
    End If

    ' Admin: runtime panel (modAdmin) -- isti runtime-controls pristup kao
    ' Podesavanja; frmStammdaten.frx se NE dira.
    If Me.Tag = "Admin" Then
        modAdmin.BuildAdminPanel Me
        If Not mChromeRemoved Then
            Me.caption = ""             ' brisi caption
            RemoveTitleBar               ' onda ga sakri
            mChromeRemoved = True
        End If
        Exit Sub
    End If

    ' Style staticnih elemenata koji ne menjaju za Setup
    On Error Resume Next
    StylePrimaryButton btnDodaj, "Dodaj"
    StylePrimaryButton btnIzmeni, "Izmeni"
    StyleExitButton btnPovratak, "Povratak"

    ' GEO buttons (samo Parcele tab, ali style ih unapred)
    StylePrimaryButton btnGeoOpen, "Otvori GeoSrbija"
    StylePrimaryButton btnPasteCoords, "Paste koordinata"
    StylePrimaryButton btnGeoSave, "Sacuvaj geo"
    StyleExitButton btnGeoClear, Poruka("STM_LBL_OBRISI_GEO")
    StylePrimaryButton btnOpenMap, "Google Maps"
    StylePrimaryButton btnOpenPolygonEditor, "Polygon editor"
    On Error GoTo EH

    ' Labeli maticnih podataka: prosiri 50% (duzi natpisi se ne skracuju; desno
    ' od njih ima slobodnog prostora). Jednokratno (m_SetupDone guard iznad).
    On Error Resume Next
    Dim liLbl As Long
    For liLbl = 1 To 10
        Me.Controls("lblField" & liLbl).width = Me.Controls("lblField" & liLbl).width * 1.5
    Next liLbl
    On Error GoTo EH

    PokreniSetup

    LoadList
    ClearFields
    SetupColumnHeaders            ' DODATO
    StyleAllFields                 ' DODATO

    ' Cenovnik je append-only: izmena ne menja istoriju, vec se dodaje
    ' novi (vazeci) red. Zato se dugme "Izmeni" sakriva za Cenovnik.
    On Error Resume Next
    btnIzmeni.Visible = (Me.Tag <> "Cenovnik")
    EnsureSoftDeleteButton
    On Error GoTo EH

    m_SelectedRow = 0

    ' RemoveTitleBar SAMO POSLE Setup-a (caption je vec postavljen)
    If Not mChromeRemoved Then
        Me.caption = ""             ' brisi caption
        RemoveTitleBar               ' onda ga sakri
        mChromeRemoved = True
    End If
    Exit Sub

EH:
    LogErr "frmStammdaten.UserForm_Activate"
    MsgBox Poruka("OTKUP_ERR_GRESKA_PRI_OTVARANJU") & Err.description, vbCritical, APP_NAME
End Sub

' Izbor Setup* procedure po Tag-u. Izdvojeno iz UserForm_Activate zato sto ga
' zove i test seam (StmTestLista): dva spiska sekcija bi se razisla, pa bi test
' merio drugu formu od one koju operater vidi.
Private Sub PokreniSetup()
    Select Case Me.Tag
        Case "Kooperanti": SetupKooperanti
        Case "Stanice": SetupStanice
        Case "Korisnici": SetupKorisnici
        Case "Kupci": SetupKupci
        Case "Vozaci": SetupVozaci
        Case "Parcele": SetupParcele
        Case "Artikli": SetupArtikli
        Case "Kulture": SetupKulture
        Case "TipAmbalaze": SetupTipAmbalaze
        Case "TipPalete": SetupTipPalete
        Case "Cenovnik": SetupCenovnik
        Case "Kutije": SetupKutije
        Case "Kese": SetupKese
        Case "VrstaGP": SetupVrstaGP
        Case Else: SetupKooperanti
    End Select
End Sub

' ============================================================
' TEST SEAM -- PRIVREMEN, umire zajedno sa ovom formom.
'
' Postoji da bi se u JEDNOM trenutku izmerilo ono sto se posle brisanja
' frmStammdaten vise ne moze izmeriti: da li novi citac (modMaticniIzvor)
' vraca isti skup zapisa koji je LoadList vracao. Kad forma ode, ode i ovaj
' seam i test 161 -- to je svrha, ne propust.
'
' Vraca Array(redovi2D, n, brKolona). redovi2D je 1-based (red, kolona);
' kolona 1 je ono sto legacy lista drzi kao identitet reda.
' ============================================================
Public Function StmTestLista(ByVal sekTag As String) As Variant
    Dim i As Long, j As Long, nc As Long, outA() As Variant
    On Error GoTo EH
    Me.tag = sekTag
    PokreniSetup
    LoadList
    nc = lstData.ColumnCount
    If nc < 1 Then nc = 1
    If lstData.ListCount = 0 Then
        StmTestLista = Array(Empty, 0, nc)
        Exit Function
    End If
    ReDim outA(1 To lstData.ListCount, 1 To nc)
    For i = 0 To lstData.ListCount - 1
        For j = 0 To nc - 1
            outA(i + 1, j + 1) = NzToText(lstData.List(i, j))
        Next j
    Next i
    StmTestLista = Array(outA, lstData.ListCount, nc)
    Exit Function
EH:
    Err.Raise Err.Number, "frmStammdaten.StmTestLista[" & sekTag & "]", Err.description
End Function

Private Sub StyleAllFields()
    On Error Resume Next

    Dim i As Long

    ' lblField1..10 -- naslovne labele, muted small
    For i = 1 To 10
        Dim lbl As MSForms.label
        Set lbl = Me.Controls("lblField" & i)
        StyleLabel lbl, TXT_MUTED(), False
        lbl.Font.Size = FONT_SIZE_SMALL
    Next i

    ' GEO koordinate labele
    StyleLabel lblNCoord, TXT_MUTED(), False
    lblNCoord.Font.Size = FONT_SIZE_SMALL
    StyleLabel lblECoord, TXT_MUTED(), False
    lblECoord.Font.Size = FONT_SIZE_SMALL

    On Error GoTo 0
End Sub

Private Sub SetupColumnHeaders()
    On Error Resume Next

    ' Default: sakriti sve
    Dim i As Long
    For i = 1 To 10
        Dim lblH As MSForms.label
        Set lblH = Me.Controls("lbl_H_STM" & i)
        If Not lblH Is Nothing Then
            StyleListHeaderLabel lblH
            lblH.caption = ""
            lblH.Visible = False
        End If
    Next i

    ' Entity-specific column headers
    Select Case Me.Tag
        Case "Kooperanti"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Ime i Prezime", True
            ShowHeader 3, "Telefon", True
            ShowHeader 4, "Stanica", True
            ShowHeader 5, "BPG", True
            ShowHeader 6, "Ra" & ChrW(269) & "un", True
            ShowHeader 7, "Pin", True
            ShowHeader 8, "Adresa", True
            ShowHeader 9, "JMBG", True
            ShowHeader 10, "Aktivan", True

        Case "Stanice"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Naziv", True
            ShowHeader 3, "Mesto", True
            ShowHeader 4, "Telefon", True
            ShowHeader 5, "Aktivan", True
            ShowHeader 6, "Kontakt Ime", True
            ShowHeader 7, "Kontakt Prezime", True
            ShowHeader 8, "Pin", True
            ShowHeader 9, "Hladnjaca", True

        Case "Korisnici"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Korisnik", True
            ShowHeader 3, "Ime i prezime", True
            ShowHeader 4, "Uloga", True
            ShowHeader 5, "Aktivan", True
            ShowHeader 6, "Oblasti (DA)", True

        Case "Kupci"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Naziv", True
            ShowHeader 3, "Adresa", True
            ShowHeader 4, "Dr" & ChrW(382) & "ava", True
            ShowHeader 5, "PIB", True
            ShowHeader 6, "MB", True
            ShowHeader 7, "Email", True
            ShowHeader 8, "Hladnjaca", True
            ShowHeader 9, "Aktivan", True
            ShowHeader 10, "Ra" & ChrW(269) & "un", True

        Case "Vozaci"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Ime", True
            ShowHeader 3, "Prezime", True
            ShowHeader 4, "Telefon", True
            ShowHeader 5, "Aktivan", True
            ShowHeader 6, "PIN", True

        Case "Parcele"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Kooperant", True
            ShowHeader 3, "Kat. broj", True
            ShowHeader 4, "Kat. opstina", True
            ShowHeader 5, "Kultura", True
            ShowHeader 6, "Povrsina (ha)", True
            ShowHeader 7, "GGAP", True
            ShowHeader 8, "Geo", True
            ShowHeader 9, "Rizik", True
            ShowHeader 10, "Napomena", True

        Case "Artikli"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Naziv", True
            ShowHeader 3, "Tip", True
            ShowHeader 4, "Jed. mere", True
            ShowHeader 5, "Cena", True
            ShowHeader 6, "Doza/ha", True
            ShowHeader 7, "Kultura", True
            ShowHeader 8, "Pakovanje", True

        Case "Kulture"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Vrsta vo" & ChrW(263) & "a", True
            ShowHeader 3, "Sorta vo" & ChrW(263) & "a", True
            ShowHeader 4, "Gajbica/paleti", True
            ShowHeader 5, "Aktivan", True
            ShowHeader 6, "Tip amb.", True
            ShowHeader 7, "Prag upoz.", True
            ShowHeader 8, "Prag blok.", True

        Case "TipAmbalaze"
            ShowHeader 1, "Tip ambala" & ChrW(382) & "e", True
            ShowHeader 2, "Te" & ChrW(382) & "ina gajbice (kg)", True
            ShowHeader 3, "Aktivan", True

        Case "TipPalete"
            ShowHeader 1, "Tip palete", True
            ShowHeader 2, "Te" & ChrW(382) & "ina (kg)", True
            ShowHeader 3, "Aktivan", True

        Case "Cenovnik"
            ShowHeader 1, "ID", True
            ShowHeader 2, "Datum", True
            ShowHeader 3, "Vrsta", True
            ShowHeader 4, "Sorta", True
            ShowHeader 5, "Klasa", True
            ShowHeader 6, "Cena", True

        Case "Kutije"
            ShowHeader 1, "Tip kutije", True
            ShowHeader 2, "Te" & ChrW(382) & "ina (kg)", True
            ShowHeader 3, "Aktivan", True

        Case "Kese"
            ShowHeader 1, "Tip kese", True
            ShowHeader 2, "Te" & ChrW(382) & "ina (kg)", True
            ShowHeader 3, "Aktivan", True

        Case "VrstaGP"
            ShowHeader 1, "Tip gotovog proizvoda", True
            ShowHeader 2, "Aktivan", True
    End Select

    On Error GoTo 0
End Sub

Private Sub ShowHeader(ByVal index As Long, ByVal txt As String, ByVal isVisible As Boolean)
    On Error Resume Next
    Dim lbl As MSForms.label
    Set lbl = Me.Controls("lbl_H_STM" & index)
    If Not lbl Is Nothing Then
        lbl.caption = txt
        lbl.Visible = isVisible
    End If
    On Error GoTo 0
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnDodaj, "Dodaj"
    StylePrimaryButton btnIzmeni, "Izmeni"
    StyleExitButton btnPovratak, "Povratak"
End Sub

Private Sub btnDodaj_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnDodaj
End Sub

Private Sub btnIzmeni_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnIzmeni
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
ResetActionButtons:     ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub

' ============================================================
' SETUP - Konfiguriert das Form je nach Entitaet
' ============================================================

Private Sub SetupKooperanti()
    ResetFieldVisibility

    Me.caption = "Kooperanti"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "KOOPERANTI"
    StyleSubtitle lblSubtitle, "Mati" & ChrW(269) & "ni podaci o kooperantima i njihovim stanicama"
    On Error GoTo 0

    m_TableName = TBL_KOOPERANTI

    m_Headers = Array( _
        "KooperantID", _
        "Ime i Prezime", _
        "Telefon", _
        "StanicaID", _
        "BPGBroj", _
        "TekuciRacun", _
        "Pin", _
        "Adresa", _
        "JMBG", _
        "Aktivan" _
    )

    m_FieldCount = 10

    lblField1.caption = "Ime": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Prezime": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Mesto": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Telefon": lblField4.Visible = True: txtField4.Visible = True

    lblField5.caption = "Stanica": lblField5.Visible = True: cmbField1.Visible = True
    lblField6.caption = "BPG Broj": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "Teku" & ChrW(263) & "i Ra" & ChrW(269) & "un": lblField7.Visible = True: txtField7.Visible = True
    lblField8.caption = "Pin": lblField8.Visible = True: txtField8.Visible = True
    lblField9.caption = "Adresa": lblField9.Visible = True: txtField9.Visible = True
    lblField10.caption = "JMBG": lblField10.Visible = True: txtField10.Visible = True

    LoadStaniceIntoCombo
End Sub

Private Sub SetupStanice()
    ResetFieldVisibility

    Me.caption = "Otkupna Mesta"

        On Error Resume Next
    StyleFrameTitleLabel lblTitle, "OTKUPNE STANICE"
    StyleSubtitle lblSubtitle, "Mati" & ChrW(269) & "ni podaci o otkupnim stanicama"
    On Error GoTo 0

    m_TableName = TBL_STANICE

    ' Redosled mora pratiti tblStanice (pozicijski prikaz u listi).
    ' Napomena: realna kol. 4 je "Kontakt" (telefon); JeHladnjaca je dodata kolona.
    m_Headers = Array( _
        "StanicaID", _
        "Naziv", _
        "Mesto", _
        "Telefon", _
        "Aktivan", _
        "Ime", _
        "Prezime", _
        "PIN", _
        "JeHladnjaca" _
    )

    m_FieldCount = 7

    lblField1.caption = "Naziv": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Mesto": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Telefon": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Kontakt Ime": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "Kontakt Prezime": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "Pin": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "Hladnjaca?": lblField7.Visible = True: txtField7.Visible = False
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False

    ' Hladnjaca flag (Da/Ne) -- auto-lanac otpremnica+zbirna+prijemnica.
    cmbField1.Visible = True
    cmbField1.Clear
    cmbField1.AddItem "Ne"
    cmbField1.AddItem "Da"
    AlignControlToRow cmbField1, txtField7
End Sub

Private Sub SetupKorisnici()
    ResetFieldVisibility
    RemoveKorisniciOblasti          ' ukloni eventualne stare dinamicke oblasti kontrole

    Me.caption = "Korisnici"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "KORISNICI"
    StyleSubtitle lblSubtitle, "Prijava i prava pristupa po oblasti (admin + korisnici)"
    On Error GoTo 0

    m_TableName = TBL_KORISNICI

    ' Display-headeri (lista). CRUD se radi po IMENU kolone (drift-safe),
    ' ne pozicijski - jer tblKorisnici ima i 9 kolona oblasti (DA/NE).
    m_Headers = Array( _
        COL_KOR_ID, _
        COL_KOR_USERNAME, _
        COL_KOR_IME, _
        COL_KOR_ULOGA, _
        COL_KOR_AKTIVAN, _
        "Oblasti" _
    )

    m_FieldCount = 6

    ' Leva kolona: tekstualna polja (ime/PIN) + 3 dropdowna (Uloga/Aktivan/Stanica).
    lblField1.caption = Poruka("KOR_LBL_KORISNICKO_IME"): lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Ime i prezime": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "PIN (izmena: prazno=isti)": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Uloga": lblField4.Visible = True: cmbField1.Visible = True
    lblField5.caption = "Aktivan": lblField5.Visible = True: cmbField2.Visible = True
    lblField6.caption = "Stanica": lblField6.Visible = True: cmbField3.Visible = True

    ' Labele bez preloma u dva reda.
    Dim li As Long
    For li = 1 To 6
        Me.Controls("lblField" & li).WordWrap = False
    Next li

    ' Uloga (Admin/Korisnik) -- dropdown na redu txtField4.
    cmbField1.Clear
    cmbField1.ColumnCount = 1
    cmbField1.AddItem ULOGA_ADMIN
    cmbField1.AddItem ULOGA_KORISNIK
    cmbField1.style = fmStyleDropDownList
    AlignControlToRow cmbField1, txtField4

    ' Aktivan (DA/NE) -- dropdown na redu txtField5.
    cmbField2.Clear
    cmbField2.ColumnCount = 1
    cmbField2.AddItem "DA"
    cmbField2.AddItem "NE"
    cmbField2.style = fmStyleDropDownList
    AlignControlToRow cmbField2, txtField5

    ' Stanica (Naziv + skriveni StanicaID, opciono) -- dropdown na redu txtField6.
    LoadStaniceIntoCombo cmbField3
    cmbField3.AddItem "", 0          ' prazno = bez stanice (opciono)
    cmbField3.style = fmStyleDropDownList
    AlignControlToRow cmbField3, txtField6

    ' Desna kolona: oblasti kao DA/NE dropdown (Model A, eksplicitno).
    BuildKorisniciOblasti

    ' Podrazumevano za NOVOG korisnika: Korisnik + Aktivan, sve oblasti NE.
    ' (Admin se bira rucno -> ApplyAdminOblastiLock postavi sve na DA i zakljuca.)
    KorisniciSetDefaults
End Sub

' Podrazumevano stanje editora za NOVOG korisnika: uloga Korisnik, Aktivan DA,
' sve oblasti NE (i otkljucane). Admin se bira rucno (cmbField1_Change -> DA+lock).
Private Sub KorisniciSetDefaults()
    On Error GoTo EH
    cmbField1.value = ULOGA_KORISNIK
    cmbField2.value = "DA"

    Dim o As Variant, cb As MSForms.ComboBox
    For Each o In modAuth.OblastiList()
        Set cb = Nothing
        On Error Resume Next                       ' kontrola mozda jos nije izgradjena
        Set cb = Me.Controls("cmbObl_" & CStr(o))
        On Error GoTo EH
        If Not cb Is Nothing Then
            cb.Locked = False
            cb.value = "NE"
        End If
    Next o
    Exit Sub
EH:
    LogErr "frmStammdaten.KorisniciSetDefaults"
End Sub

' --- Korisnici helperi (prava po oblasti, Model A) ---
' "DA"/"NE" za datu oblast iz desne kolone dropdowna. Normalizaciju uloge i
' pravilo "admin dobija sve" NE radi forma nego pisac (modMaticniKorisnici):
' dva mesta koja isto pravilo pisu bila bi dva mesta koja se razidju.
Private Function OblComboVal(ByVal oblast As String) As String
    Dim cb As MSForms.ComboBox
    On Error Resume Next
    Set cb = Me.Controls("cmbObl_" & oblast)
    On Error GoTo 0

    If cb Is Nothing Then
        OblComboVal = "NE"
    ElseIf StrComp(Trim$(cb.value), "DA", vbTextCompare) = 0 Then
        OblComboVal = "DA"
    Else
        OblComboVal = "NE"
    End If
End Function

' "DA"/"NE" za datu oblast iz reda tabele (za punjenje dropdowna pri izboru).
' ByRef: citac po celiji -- ByVal bi kopirao ceo niz po pozivu (v. KOPIJA_NIZA).
Private Function OblastValueFromRow(ByRef data As Variant, ByVal rowIdx As Long, ByVal oblast As String) As String
    Dim ci As Long
    ci = GetColumnIndex(TBL_KORISNICI, oblast)
    If ci > 0 Then
        If StrComp(Trim$(NzToText(data(rowIdx, ci))), "DA", vbTextCompare) = 0 Then
            OblastValueFromRow = "DA"
        Else
            OblastValueFromRow = "NE"
        End If
    Else
        OblastValueFromRow = "NE"
    End If
End Function

' Prikaz naziva oblasti (dijakritika kroz ChrW -- izvor ostaje ASCII; kolone
' u tblKorisnici su ASCII konstante OBL_*). Prikaz != naziv kolone.
Private Function OblastCaption(ByVal oblast As String) As String
    Select Case oblast
        Case OBL_IZVESTAJI:    OblastCaption = "Izve" & ChrW(353) & "taji"
        Case OBL_MARZA:        OblastCaption = "Mar" & ChrW(382) & "a"
        Case OBL_MATICNI:      OblastCaption = "Mati" & ChrW(269) & "ni podaci"
        Case OBL_OTVORI_EXCEL: OblastCaption = "Otvori Excel"
        Case OBL_SYNC_PWA:     OblastCaption = "Sinhronizuj PWA"
        Case Else:             OblastCaption = oblast
    End Select
End Function

' Desna kolona Korisnici editora: po jedan DA/NE dropdown za svaku oblast
' (Model A). Kontrole su dinamicke (Controls.Add) jer .frx ima samo cmbField1..6;
' forma se po sekciji Unload-uje (frmMaticniPodaci.OpenSekcija) pa ne cure dalje.
Private Sub BuildKorisniciOblasti()
    On Error GoTo EH

    Dim oblasti As Variant
    oblasti = modAuth.OblastiList()

    ' Geometrija reda iz txtField1 (pouzdano poravnat u .frx).
    Dim rowH As Single
    rowH = txtField1.Height
    If rowH < 15 Then rowH = 18

    Const ROWGAP As Single = 4
    Const LBLW As Single = 124      ' prostor za naziv oblasti
    Const LBLGAP As Single = 16     ' razmak labela -> combo (da cmb ne pada preko labele)
    Const CMBW As Single = 86       ' siri combo (citljiviji DA/NE)
    Const COLGAP As Single = 190    ' jos vise desno od leve kolone polja

    ' Desna kolona pocinje desno od leve kolone polja.
    Dim colX As Single
    colX = txtField1.Left + txtField1.width + COLGAP

    ' Prvi combo se ravna sa poljem "Korisnicko ime" (txtField1); header tik iznad.
    Dim firstTop As Single
    firstTop = txtField1.top

    ' Header tik iznad prvog reda; clamp da ne isklizne iznad forme.
    Dim hdrTop As Single
    hdrTop = firstTop - rowH - 2
    If hdrTop < 4 Then hdrTop = 4

    Dim hdr As MSForms.label
    Set hdr = Me.Controls.Add("Forms.Label.1", "lblOblHdr", True)
    hdr.Left = colX
    hdr.top = hdrTop
    hdr.width = LBLW + LBLGAP + CMBW
    hdr.Height = rowH
    hdr.caption = "OBLASTI (pristup)"
    hdr.WordWrap = False
    StyleLabel hdr, TXT_MUTED(), True
    hdr.Font.Size = FONT_SIZE_SMALL
    hdr.Font.Bold = True

    Dim yy As Single
    yy = firstTop

    Dim k As Long, oname As String
    Dim lb As MSForms.label, cb As MSForms.ComboBox
    For k = LBound(oblasti) To UBound(oblasti)
        oname = CStr(oblasti(k))

        Set lb = Me.Controls.Add("Forms.Label.1", "lblObl_" & oname, True)
        lb.Left = colX
        lb.top = yy + 2
        lb.width = LBLW
        lb.Height = rowH
        lb.caption = OblastCaption(oname)
        lb.WordWrap = False
        StyleLabel lb, TXT_MUTED(), False
        lb.Font.Size = FONT_SIZE_SMALL

        Set cb = Me.Controls.Add("Forms.ComboBox.1", "cmbObl_" & oname, True)
        cb.Left = colX + LBLW + LBLGAP
        cb.top = yy
        cb.width = CMBW
        cb.Height = rowH
        cb.Clear
        cb.AddItem "DA"
        cb.AddItem "NE"
        cb.style = fmStyleDropDownList
        cb.value = "NE"

        yy = yy + rowH + ROWGAP
    Next k

    ' Prosiri formu ako desna kolona prelazi vidljivu sirinu.
    Dim needW As Single
    needW = colX + LBLW + LBLGAP + CMBW + 24
    If Me.InsideWidth < needW Then
        Me.width = Me.width + (needW - Me.InsideWidth)
    End If
    Exit Sub

EH:
    LogErr "frmStammdaten.BuildKorisniciOblasti"
End Sub

' Ukloni dinamicke oblasti kontrole (idempotentno).
Private Sub RemoveKorisniciOblasti()
    On Error Resume Next
    Dim oblasti As Variant
    oblasti = modAuth.OblastiList()
    Dim k As Long
    For k = LBound(oblasti) To UBound(oblasti)
        Me.Controls.Remove "lblObl_" & CStr(oblasti(k))
        Me.Controls.Remove "cmbObl_" & CStr(oblasti(k))
    Next k
    Me.Controls.Remove "lblOblHdr"
    On Error GoTo 0
End Sub

' CSV dozvoljenih oblasti iz reda tabele (kolone gde je vrednost "DA").
Private Function OblastiCsvFromRow(ByVal data As Variant, ByVal rowIdx As Long) As String
    Dim res As String
    Dim obl As Variant
    Dim ci As Long
    For Each obl In modAuth.OblastiList()
        ci = GetColumnIndex(TBL_KORISNICI, CStr(obl))
        If ci > 0 Then
            If StrComp(Trim$(NzToText(data(rowIdx, ci))), "DA", vbTextCompare) = 0 Then
                If Len(res) > 0 Then res = res & ", "
                res = res & OblastCaption(CStr(obl))   ' prikaz sa dijakritikom
            End If
        End If
    Next obl
    OblastiCsvFromRow = res
End Function

Private Sub SetupKupci()
    ResetFieldVisibility

    Me.caption = "Kupci"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "KUPCI"
    StyleSubtitle lblSubtitle, "Mati" & ChrW(269) & "ni podaci o kupcima"
    On Error GoTo 0

    m_TableName = TBL_KUPCI

    m_Headers = Array( _
        "KupacID", _
        "Naziv", _
        "Adresa", _
        "Dr" & ChrW(382) & "ava", _
        "PIB", _
        "MaticniBroj", _
        "Email", _
        "Hladnjaca", _
        "Aktivan", _
        "TekuciRacun" _
    )

    m_FieldCount = 11

    lblField1.caption = "Naziv": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Ulica": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Mesto": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Po" & ChrW(353) & "tanski Broj": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "Dr" & ChrW(382) & "ava": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "PIB": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "Mati" & ChrW(269) & "ni Broj": lblField7.Visible = True: txtField7.Visible = True
    lblField8.caption = "Email": lblField8.Visible = True: txtField8.Visible = True
    lblField9.caption = "Hladnjaca": lblField9.Visible = True: txtField9.Visible = True
    lblField10.caption = "Teku" & ChrW(263) & "i Ra" & ChrW(269) & "un": lblField10.Visible = True: txtField10.Visible = True
End Sub

Private Sub SetupVozaci()
    ResetFieldVisibility

    Me.caption = "Vozaci"
        On Error Resume Next
    StyleFrameTitleLabel lblTitle, "VOZACI"
    StyleSubtitle lblSubtitle, "Mati" & ChrW(269) & "ni podaci o vozacima"
    On Error GoTo 0
    m_TableName = TBL_VOZACI

    m_Headers = Array( _
        "VozacID", _
        "Ime", _
        "Prezime", _
        "Telefon", _
        "Aktivan", _
        "PIN" _
    )

    m_FieldCount = 5

    lblField1.caption = "Ime": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Prezime": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Telefon": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "PIN": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "": lblField5.Visible = False: txtField5.Visible = False
    lblField6.caption = "": lblField6.Visible = False: txtField6.Visible = False
    lblField7.caption = "": lblField7.Visible = False: txtField7.Visible = False
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False
End Sub

Private Sub SetupParcele()
    ResetFieldVisibility

    Me.caption = "Parcele"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "KATASTARSKE PARCELE"
    StyleSubtitle lblSubtitle, "Mati" & ChrW(269) & "ni podaci o parcelama kooperanata"
    On Error GoTo 0

    m_TableName = TBL_PARCELE

    ' Display headers for ListBox only
    m_Headers = Array( _
        "ParcelaID", _
        "Kooperant", _
        "KatBroj", _
        "KatOpstina", _
        "Kultura", _
        "PovrsinaHa", _
        "GGAPStatus", _
        "Geo", _
        "Rizik", _
        "Napomena" _
    )

    m_FieldCount = 19

    cmbField1.Visible = True     ' Kooperant
    cmbField2.Visible = True     ' Kultura
    cmbField3.Visible = True     ' GGAPStatus

    cmbField4.Visible = False
    cmbField5.Visible = False
    cmbField6.Visible = False

    lblField1.caption = "Kooperant": lblField1.Visible = True: txtField1.Visible = False
    lblField2.caption = "Kat. Broj": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Kat. Op" & ChrW(353) & "tina": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Kultura": lblField4.Visible = True: txtField4.Visible = False
    lblField5.caption = "Povrsina (ha)": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "GGAP Status": lblField6.Visible = True: txtField6.Visible = False
    lblField7.caption = "Napomena": lblField7.Visible = True: txtField7.Visible = True
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False

    cmbField1.Clear

    Dim data As Variant
    Dim i As Long
    Dim colID As Long, colIme As Long, colPrez As Long

    data = GetTableData(TBL_KOOPERANTI)
    If Not IsEmpty(data) Then
        colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
        colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
        colPrez = GetColumnIndex(TBL_KOOPERANTI, "Prezime")

        For i = 1 To UBound(data, 1)
            cmbField1.AddItem CStr(data(i, colIme)) & " " & _
                              CStr(data(i, colPrez)) & " (" & _
                              CStr(data(i, colID)) & ")"
        Next i
    End If

    cmbField2.Clear
    Dim kulture As Variant
    kulture = GetLookupList(TBL_KULTURE, "VrstaVoca")
    If IsArray(kulture) Then
        For i = LBound(kulture) To UBound(kulture)
            cmbField2.AddItem CStr(kulture(i))
        Next i
    End If

    cmbField3.Clear
    cmbField3.AddItem "Da"
    cmbField3.AddItem "Ne"
    cmbField3.AddItem "U postupku"

    SetGeoControlsVisible False

End Sub
Private Sub SetupArtikli()
    ResetFieldVisibility
    Me.caption = "Artikli"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "ARTIKLI"
    StyleSubtitle lblSubtitle, "Mati" & ChrW(269) & "ni podaci o artiklima"
    On Error GoTo 0

    m_TableName = TBL_ARTIKLI

    m_Headers = Array( _
        "ArtikalID", _
        "Naziv", _
        "Tip", _
        "JedinicaMere", _
        "CenaPoJedinici", _
        "DozaPoHa", _
        "Kultura", _
        "Pakovanje" _
    )

    m_FieldCount = 7

    cmbField2.Visible = False: cmbField3.Visible = False: cmbField4.Visible = False
    lblField1.caption = "Naziv": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Tip": lblField2.Visible = True: txtField2.Visible = False
    lblField3.caption = "Jedinica Mere": lblField3.Visible = True: txtField3.Visible = False
    lblField4.caption = "Cena po jed.": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "Kultura": lblField5.Visible = True: txtField5.Visible = False
    lblField6.caption = "Doza po ha": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "Pakovanje": lblField7.Visible = True: txtField7.Visible = True
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False

    ' Tip
    cmbField5.Visible = True
    cmbField5.Clear
    cmbField5.AddItem "Pesticid"
    cmbField5.AddItem "Djubrivo"
    cmbField5.AddItem "SadniMaterijal"

    ' Jedinica mere
    cmbField6.Visible = True
    cmbField6.Clear
    cmbField6.AddItem "kg"
    cmbField6.AddItem "l"
    cmbField6.AddItem "kom"

    ' Kultura
    cmbField1.Visible = True
    cmbField1.Clear
    cmbField1.AddItem ""

    Dim kulture As Variant
    Dim i As Long

    kulture = GetLookupList(TBL_KULTURE, "VrstaVoca")
    If IsArray(kulture) Then
        For i = LBound(kulture) To UBound(kulture)
            cmbField1.AddItem CStr(kulture(i))
        Next i
    End If
End Sub

Private Sub SetupKulture()
    ResetFieldVisibility
    Me.caption = "Kulture"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "KULTURE"
    StyleSubtitle lblSubtitle, "Mati" & ChrW(269) & "ni podaci o vrstama i sortama vo" & ChrW(263) & "a"
    On Error GoTo 0

    m_TableName = TBL_KULTURE

    ' Redosled mora pratiti tblKulture (pozicijski prikaz u listi):
    ' KulturaID | VrstaVoca | SortaVoca | GajbicaPoPaleti | Aktivan | TipAmbalaze
    m_Headers = Array( _
        "KulturaID", _
        "VrstaVoca", _
        "SortaVoca", _
        "GajbicaPoPaleti", _
        "Aktivan", _
        "TipAmbalaze" _
    )

    m_FieldCount = 6

    lblField1.caption = "Vrsta vo" & ChrW(263) & "a": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Sorta vo" & ChrW(263) & "a": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Gajbica po paleti": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Tip ambala" & ChrW(382) & "e (podraz.)": lblField4.Visible = True: txtField4.Visible = False
    lblField5.caption = "Prag upozorenja (kg/gajb.)": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "Prag blokade (kg/gajb.)": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "": lblField7.Visible = False: txtField7.Visible = False
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False

    ' Tip ambalaze (podrazumevani za kulturu) -- combo iz tblTipAmbalaze.
    cmbField1.Visible = True
    cmbField1.Clear
    cmbField1.AddItem ""
    Dim taOpt As Variant, ti As Long
    taOpt = GetTipAmbalazeOptions()
    If IsArray(taOpt) Then
        For ti = LBound(taOpt) To UBound(taOpt)
            cmbField1.AddItem CStr(taOpt(ti))
        Next ti
    End If
    AlignControlToRow cmbField1, txtField4
End Sub

Private Sub SetupTipAmbalaze()
    ResetFieldVisibility
    Me.caption = "Tip ambala" & ChrW(382) & "e"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "TIP AMBALA" & ChrW(381) & "E"
    StyleSubtitle lblSubtitle, ChrW(352) & "ifarnik ambala" & ChrW(382) & "e (tip i te" & ChrW(382) & "ina prazne gajbice)"
    On Error GoTo 0

    m_TableName = TBL_TIP_AMBALAZE

    m_Headers = Array(COL_TAMB_TIP, COL_TAMB_TEZINA)
    m_FieldCount = 2

    lblField1.caption = "Tip ambala" & ChrW(382) & "e": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Te" & ChrW(382) & "ina gajbice (kg)": lblField2.Visible = True: txtField2.Visible = True
End Sub

Private Sub SetupTipPalete()
    ResetFieldVisibility
    Me.caption = "Tip palete"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "TIP PALETE"
    StyleSubtitle lblSubtitle, ChrW(352) & "ifarnik paleta (tip i te" & ChrW(382) & "ina prazne palete)"
    On Error GoTo 0

    m_TableName = TBL_TIP_PALETE

    m_Headers = Array(COL_TPAL_TIP, COL_TPAL_TEZINA)
    m_FieldCount = 2

    lblField1.caption = "Tip palete": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Te" & ChrW(382) & "ina (kg)": lblField2.Visible = True: txtField2.Visible = True
End Sub

Private Sub SetupKutije()
    ResetFieldVisibility
    Me.caption = "Kutije"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "KUTIJE"
    StyleSubtitle lblSubtitle, ChrW(352) & "ifarnik kutija (tip i te" & ChrW(382) & "ina prazne kutije)"
    On Error GoTo 0

    m_TableName = TBL_KUTIJE

    m_Headers = Array(COL_KUT_TIP, COL_KUT_TEZINA, "Aktivan")
    m_FieldCount = 2

    lblField1.caption = "Tip kutije": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Te" & ChrW(382) & "ina (kg)": lblField2.Visible = True: txtField2.Visible = True
End Sub

Private Sub SetupKese()
    ResetFieldVisibility
    Me.caption = "Kese"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "KESE"
    StyleSubtitle lblSubtitle, ChrW(352) & "ifarnik kesa (tip i te" & ChrW(382) & "ina prazne kese)"
    On Error GoTo 0

    m_TableName = TBL_KESE

    m_Headers = Array(COL_KES_TIP, COL_KES_TEZINA, "Aktivan")
    m_FieldCount = 2

    lblField1.caption = "Tip kese": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Te" & ChrW(382) & "ina (kg)": lblField2.Visible = True: txtField2.Visible = True
End Sub

Private Sub SetupVrstaGP()
    ResetFieldVisibility
    Me.caption = "Vrsta gotovog proizvoda"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "VRSTA GOTOVOG PROIZVODA"
    StyleSubtitle lblSubtitle, ChrW(352) & "ifarnik tipova gotovog proizvoda"
    On Error GoTo 0

    m_TableName = TBL_VRSTA_GP

    m_Headers = Array(COL_VGP_TIP, "Aktivan")
    m_FieldCount = 1

    lblField1.caption = "Tip gotovog proizvoda": lblField1.Visible = True: txtField1.Visible = True
End Sub

Private Sub SetupCenovnik()
    ResetFieldVisibility
    Me.caption = "Cenovnik"

    On Error Resume Next
    StyleFrameTitleLabel lblTitle, "CENOVNIK"
    StyleSubtitle lblSubtitle, Poruka("STM_LBL_CENE_PROIZVODU_SVAKA")
    On Error GoTo 0

    m_TableName = TBL_CENOVNIK

    ' Prikaz (lista) -- samo kljucne kolone:
    m_Headers = Array( _
        COL_CEN_ID, _
        COL_CEN_DATUM, _
        COL_CEN_VRSTA, _
        COL_CEN_SORTA, _
        COL_CEN_KLASA, _
        COL_CEN_CENA _
    )

    m_FieldCount = 5

    lblField1.caption = "Vrsta vo" & ChrW(263) & "a": lblField1.Visible = True: txtField1.Visible = False
    lblField2.caption = "Sorta vo" & ChrW(263) & "a": lblField2.Visible = True: txtField2.Visible = False
    lblField3.caption = "Klasa": lblField3.Visible = True: txtField3.Visible = False
    lblField4.caption = "Datum": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "Cena": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "": lblField6.Visible = False: txtField6.Visible = False
    lblField7.caption = "": lblField7.Visible = False: txtField7.Visible = False
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False

    ' Vrsta voca
    cmbField1.Visible = True
    cmbField1.Clear
    Dim kulture As Variant
    Dim i As Long
    kulture = GetLookupList(TBL_KULTURE, "VrstaVoca")
    If IsArray(kulture) Then
        For i = LBound(kulture) To UBound(kulture)
            cmbField1.AddItem CStr(kulture(i))
        Next i
    End If

    ' Sorta voca (kaskada se puni u cmbField1_Change)
    cmbField2.Visible = True
    cmbField2.Clear

    ' Klasa
    cmbField3.Visible = True
    cmbField3.Clear
    cmbField3.AddItem KLASA_I
    cmbField3.AddItem KLASA_II

    ' Poravnaj combo-e (Vrsta/Sorta/Klasa) sa redovima 1/2/3 (lblField1..3).
    ' U .frx su cmbField1..3 u zasebnom klasteru nize pa bi prekrivali
    ' txtField4/5 (Datum/Cena). Forma se otvara sveza po sekciji (OpenSekcija
    ' radi Unload), pa repozicioniranje ne utice na druge tabove.
    AlignControlToRow cmbField1, txtField1
    AlignControlToRow cmbField2, txtField2
    AlignControlToRow cmbField3, txtField3

    txtField4.value = Format$(Date, "d.m.yyyy")
End Sub

' Kopira geometriju reda (lblFieldN <-> txtFieldN su pouzdano poravnati u .frx)
' na drugu kontrolu istog reda. Koristi se da combo sedne na red svoje labele.
Private Sub AlignControlToRow(ByVal ctl As MSForms.Control, ByVal refCtl As MSForms.Control)
    On Error Resume Next
    ctl.Left = refCtl.Left
    ctl.top = refCtl.top
    ctl.width = refCtl.width
    ctl.Height = refCtl.Height
End Sub

' ============================================================
' SOFT-DELETE (#1): runtime dugme "Deaktiviraj/Aktiviraj"
' Vidljivo samo za tabele koje imaju kolonu Aktivan/Aktivna.
' Klik flipuje status izabranog reda (ne brise ga).
' ============================================================
Private Function AktivanColName() As String
    ' Vraca naziv kolone statusa ("Aktivan" ili "Aktivna") ili "" ako ne postoji.
    On Error Resume Next
    If GetColumnIndex(m_TableName, "Aktivan") > 0 Then
        AktivanColName = "Aktivan"
    ElseIf GetColumnIndex(m_TableName, "Aktivna") > 0 Then
        AktivanColName = "Aktivna"
    End If
End Function

Private Sub EnsureSoftDeleteButton()
    On Error Resume Next

    ' Napravi dugme jednom (runtime), desno od btnIzmeni.
    If m_softWrap Is Nothing Then
        Dim c As MSForms.CommandButton
        Set c = Me.Controls.Add("Forms.CommandButton.1", "btnSoftDelete", True)
        c.Left = btnIzmeni.Left + btnIzmeni.width + 8
        c.top = btnIzmeni.top
        c.width = btnIzmeni.width
        c.Height = btnIzmeni.Height
        StyleStornoButton c, "Deaktiviraj/Aktiviraj"

        Set m_softWrap = New clsStmBtn
        Set m_softWrap.btn = c
    End If

    ' Vidljivo samo gde ima status kolona (i nije Cenovnik/Podesavanja).
    m_softWrap.btn.Visible = (Len(AktivanColName()) > 0)
End Sub

' Public -- poziva ga clsStmBtn na klik. Flipuje status izabranog reda.
Public Sub OnSoftDeleteClick()
    Dim kljuc As String, odgovor As String, noviStatus As String
    On Error GoTo EH

    If lstData.ListIndex >= 0 Then m_SelectedRow = GetMappedSelectedRow()
    If m_SelectedRow = 0 Then
        MsgBox Poruka("MATU_ERR_NEMA_REDA"), vbExclamation, APP_NAME
        Exit Sub
    End If

    ' Od M4 SVE sekcije idu kroz modMaticniUnos, Korisnici ukljuceno. Time je
    ' zatvoren nalaz iz M2a: ovo dugme je korisniku upisivalo "Neaktivan", a
    ' modAuth neaktivnim smatra samo "NE" -- deaktivirani korisnik se i dalje
    ' prijavljivao. Sada pisac za tu sekciju pise "NE" (v. modMaticniKorisnici
    ' i UI_MIGRACIJA_KATALOG 26.18).
    '
    ' Nepoznat Tag se ODBIJA, isto kao u btnDodaj/btnIzmeni: zatecena putanja
    ' koja bi tiho upisala pogresan recnik je upravo ono sto se ovde zatvara.
    kljuc = SekcijaKljuc()
    If Len(kljuc) = 0 Then
        MsgBox Poruka("MATU_ERR_NEPOZNATA_SEKCIJA") & " " & Me.Tag, vbCritical, APP_NAME
        Exit Sub
    End If

    odgovor = modMaticniUnos.MatPromeniStatus(kljuc, m_SelectedRow, noviStatus)
    If Len(odgovor) > 0 Then
        MsgBox odgovor, vbExclamation, APP_NAME
        Exit Sub
    End If

    LoadList
    ClearFields
    m_SelectedRow = 0
    MsgBox Poruka("MATU_OK_STATUS") & " " & noviStatus, vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmStammdaten.OnSoftDeleteClick"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_PROMENI") & Err.description, vbCritical, APP_NAME
End Sub

' Azurira prvu kolonu iz liste alias-a koja stvarno postoji u tabeli.
' Tolerantno na schema drift: ako nijedan alias ne postoji, tiho preskace
' (ne rusi transakciju). Vraca True ako je nesto azurirano.
' UpdateFirstExistingCol je PRESELJENO u modMaticniIzvor.MatKolonaPolja
' ("@alias:A,B" u opisu polja) -- isti probe, ali sada ga koriste i unos i
' izmena, a ne samo izmena.
' ============================================================
' LISTE LADEN
' ============================================================
Private Sub LoadList()
    On Error GoTo EH

    lstData.RowSource = ""
    lstData.Clear
    ResetRowMap

    Dim data As Variant
    data = GetTableData(m_TableName)
    If IsEmpty(data) Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim maxCols As Long

    Select Case Me.Tag

        Case "Kooperanti"
            lstData.ColumnCount = 10

            Dim kID As Long, kIme As Long, kPrez As Long, kMesto As Long
            Dim kTel As Long, kStanica As Long, kAktivan As Long
            Dim kBPG As Long, kRacun As Long, kPin As Long
            Dim kAdresa As Long, kJMBG As Long

            kID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
            kIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
            kPrez = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
            kMesto = GetColumnIndex(TBL_KOOPERANTI, "Mesto")
            kTel = GetColumnIndex(TBL_KOOPERANTI, "Telefon")
            kStanica = GetColumnIndex(TBL_KOOPERANTI, "StanicaID")
            kAktivan = GetColumnIndex(TBL_KOOPERANTI, "Aktivan")
            kBPG = GetColumnIndex(TBL_KOOPERANTI, "BPGBroj")
            kRacun = GetColumnIndex(TBL_KOOPERANTI, "TekuciRacun")
            kPin = GetColumnIndex(TBL_KOOPERANTI, "Pin")
            kAdresa = GetColumnIndex(TBL_KOOPERANTI, "Adresa")
            kJMBG = GetColumnIndex(TBL_KOOPERANTI, "JMBG")

            If kID = 0 Or kIme = 0 Or kPrez = 0 Or kMesto = 0 Or kTel = 0 Or _
               kStanica = 0 Or kAktivan = 0 Or kBPG = 0 Or kRacun = 0 Or _
               kPin = 0 Or kAdresa = 0 Or kJMBG = 0 Then
                Err.Raise vbObjectError + 7201, "frmStammdaten.LoadList", _
                          "Nedostaju kolone u tblKooperanti."
            End If

            Dim punoIme As String
            Dim punaAdresa As String
            Dim stanicaNaziv As String

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, kID))) <> "" Then
                    AddRowMap i

                    punoIme = Trim$(NzToText(data(i, kIme)) & " " & NzToText(data(i, kPrez)))

                    punaAdresa = Trim$(NzToText(data(i, kAdresa)))
                    If Len(Trim$(NzToText(data(i, kMesto)))) > 0 Then
                        If Len(punaAdresa) > 0 Then punaAdresa = punaAdresa & ", "
                        punaAdresa = punaAdresa & NzToText(data(i, kMesto))
                    End If

                    stanicaNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", _
                                                    NzToText(data(i, kStanica)), "Naziv"))

                    lstData.AddItem NzToText(data(i, kID))
                    lstData.List(lstData.ListCount - 1, 1) = punoIme
                    lstData.List(lstData.ListCount - 1, 2) = NzToText(data(i, kTel))
                    lstData.List(lstData.ListCount - 1, 3) = stanicaNaziv
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, kBPG))
                    lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, kRacun))
                    lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, kPin))
                    lstData.List(lstData.ListCount - 1, 7) = punaAdresa
                    lstData.List(lstData.ListCount - 1, 8) = NzToText(data(i, kJMBG))
                    lstData.List(lstData.ListCount - 1, 9) = NzToText(data(i, kAktivan))
                End If
            Next i

        Case "Kupci"
            lstData.ColumnCount = 10

            Dim kupID As Long, kupNaziv As Long, kupUlica As Long, kupMesto As Long
            Dim kupPosta As Long, kupDrzava As Long, kupPIB As Long, kupMB As Long
            Dim kupEmail As Long, kupHlad As Long, kupAktivan As Long, kupRacun As Long

            kupID = GetColumnIndex(TBL_KUPCI, "KupacID")
            kupNaziv = GetColumnIndex(TBL_KUPCI, "Naziv")
            kupUlica = GetColumnIndex(TBL_KUPCI, "Ulica")
            kupMesto = GetColumnIndex(TBL_KUPCI, "Mesto")
            kupPosta = GetColumnIndex(TBL_KUPCI, "PostanskiBroj")
            kupDrzava = GetColumnIndex(TBL_KUPCI, "Dr" & ChrW(382) & "ava")
            If kupDrzava = 0 Then kupDrzava = GetColumnIndex(TBL_KUPCI, "Drzava")
            kupPIB = GetColumnIndex(TBL_KUPCI, "PIB")
            kupMB = GetColumnIndex(TBL_KUPCI, "MaticniBroj")
            kupEmail = GetColumnIndex(TBL_KUPCI, "Email")
            kupHlad = GetColumnIndex(TBL_KUPCI, "Hladnjaca")
            kupAktivan = GetColumnIndex(TBL_KUPCI, "Aktivan")
            kupRacun = GetColumnIndex(TBL_KUPCI, "TekuciRacun")

            ' Tolerantno na schema-drift: obavezan je samo PK (KupacID); kolone
            ' koje fale ostaju prazne u listi umesto da obore ceo tab. Ranije je
            ' tvrdi Err.Raise na bilo koju od 12 kolona rusio otvaranje "Kupci"
            ' kad se sema instalacije razlikuje (npr. "Drzava" bez dijakritike).
            If kupID = 0 Then
                Err.Raise vbObjectError + 7202, "frmStammdaten.LoadList", _
                          "Nedostaje kolona KupacID u tblKupci."
            End If

            Dim kupacAdresa As String

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, kupID))) <> "" Then
                    AddRowMap i

                    kupacAdresa = ""
                    If kupUlica > 0 Then kupacAdresa = Trim$(NzToText(data(i, kupUlica)))

                    If kupPosta > 0 Then
                        If Len(Trim$(NzToText(data(i, kupPosta)))) > 0 Then
                            If Len(kupacAdresa) > 0 Then kupacAdresa = kupacAdresa & ", "
                            kupacAdresa = kupacAdresa & NzToText(data(i, kupPosta))
                        End If
                    End If

                    If kupMesto > 0 Then
                        If Len(Trim$(NzToText(data(i, kupMesto)))) > 0 Then
                            If Len(kupacAdresa) > 0 Then kupacAdresa = kupacAdresa & " "
                            kupacAdresa = kupacAdresa & NzToText(data(i, kupMesto))
                        End If
                    End If

                    lstData.AddItem NzToText(data(i, kupID))
                    If kupNaziv > 0 Then lstData.List(lstData.ListCount - 1, 1) = NzToText(data(i, kupNaziv))
                    lstData.List(lstData.ListCount - 1, 2) = kupacAdresa
                    If kupDrzava > 0 Then lstData.List(lstData.ListCount - 1, 3) = NzToText(data(i, kupDrzava))
                    If kupPIB > 0 Then lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, kupPIB))
                    If kupMB > 0 Then lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, kupMB))
                    If kupEmail > 0 Then lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, kupEmail))
                    If kupHlad > 0 Then lstData.List(lstData.ListCount - 1, 7) = NzToText(data(i, kupHlad))
                    If kupAktivan > 0 Then lstData.List(lstData.ListCount - 1, 8) = NzToText(data(i, kupAktivan))
                    If kupRacun > 0 Then lstData.List(lstData.ListCount - 1, 9) = NzToText(data(i, kupRacun))
                End If
            Next i

        Case "Parcele"
            lstData.ColumnCount = 10

            Dim pid As Long, pKoop As Long, pKat As Long, pOpstina As Long
            Dim pKultura As Long, pPov As Long, pGGAP As Long
            Dim pGeoStatus As Long, pGeoSource As Long, pRizik As Long, pNapomena As Long

            pid = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
            pKoop = GetColumnIndex(TBL_PARCELE, COL_PAR_KOOP)
            pKat = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_BROJ)
            pOpstina = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_OPSTINA)
            pKultura = GetColumnIndex(TBL_PARCELE, COL_PAR_KULTURA)
            pPov = GetColumnIndex(TBL_PARCELE, COL_PAR_POVRSINA)
            pGGAP = GetColumnIndex(TBL_PARCELE, COL_PAR_GGAP)
            pGeoStatus = GetColumnIndex(TBL_PARCELE, COL_PAR_GEO_STATUS)
            pGeoSource = GetColumnIndex(TBL_PARCELE, COL_PAR_GEO_SOURCE)
            pRizik = GetColumnIndex(TBL_PARCELE, COL_PAR_RIZIK)
            pNapomena = GetColumnIndex(TBL_PARCELE, COL_PAR_NAPOMENA)

            If pid = 0 Or pKoop = 0 Or pKat = 0 Or pOpstina = 0 Or _
               pKultura = 0 Or pPov = 0 Or pGGAP = 0 Or pGeoStatus = 0 Or _
               pGeoSource = 0 Or pRizik = 0 Or pNapomena = 0 Then
                Err.Raise vbObjectError + 7203, "frmStammdaten.LoadList", _
                          "Nedostaju kolone u tblParcele."
            End If

            Dim koopID As String
            Dim koopNaziv As String
            Dim geoInfo As String
            Dim rizikInfo As String

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, pid))) <> "" Then
                    AddRowMap i

                    koopID = NzToText(data(i, pKoop))
                    koopNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime")) & " " & _
                                CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime")) & _
                                " (" & koopID & ")"

                    geoInfo = NzToText(data(i, pGeoStatus))
                    If Len(NzToText(data(i, pGeoSource))) > 0 Then
                        If Len(geoInfo) > 0 Then geoInfo = geoInfo & " / "
                        geoInfo = geoInfo & NzToText(data(i, pGeoSource))
                    End If

                    rizikInfo = NzToText(data(i, pRizik))

                    lstData.AddItem NzToText(data(i, pid))
                    lstData.List(lstData.ListCount - 1, 1) = koopNaziv
                    lstData.List(lstData.ListCount - 1, 2) = NzToText(data(i, pKat))
                    lstData.List(lstData.ListCount - 1, 3) = NzToText(data(i, pOpstina))
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, pKultura))
                    lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, pPov))
                    lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, pGGAP))
                    lstData.List(lstData.ListCount - 1, 7) = geoInfo
                    lstData.List(lstData.ListCount - 1, 8) = rizikInfo
                    lstData.List(lstData.ListCount - 1, 9) = NzToText(data(i, pNapomena))
                End If
            Next i

        Case "Cenovnik"
            lstData.ColumnCount = UBound(m_Headers) - LBound(m_Headers) + 1

            Dim cCenId As Long, cCenDat As Long, cCenVr As Long
            Dim cCenSo As Long, cCenKl As Long, cCenCe As Long
            cCenId = GetColumnIndex(TBL_CENOVNIK, COL_CEN_ID)
            cCenDat = GetColumnIndex(TBL_CENOVNIK, COL_CEN_DATUM)
            cCenVr = GetColumnIndex(TBL_CENOVNIK, COL_CEN_VRSTA)
            cCenSo = GetColumnIndex(TBL_CENOVNIK, COL_CEN_SORTA)
            cCenKl = GetColumnIndex(TBL_CENOVNIK, COL_CEN_KLASA)
            cCenCe = GetColumnIndex(TBL_CENOVNIK, COL_CEN_CENA)

            If cCenId = 0 Or cCenVr = 0 Or cCenKl = 0 Or cCenCe = 0 Then
                Err.Raise vbObjectError + 7210, "frmStammdaten.LoadList", _
                          "Nedostaju kolone u tblCenovnik."
            End If

            ' Reuse postojecih helpera (modArrayUtils / modHelpers): nema rucnog sorta.
            ' ExcludeStornirano -> SortArray (datum opadajuce, tie-break CenaID -> kasniji unos gore).
            Dim cenData As Variant
            cenData = ExcludeStornirano(data, TBL_CENOVNIK)
            If IsEmpty(cenData) Then Exit Sub

            Dim cenSortCol As Long
            cenSortCol = cCenDat
            If cenSortCol = 0 Then cenSortCol = cCenId
            cenData = SortArray(cenData, cenSortCol, False, cCenId)
            If IsEmpty(cenData) Then Exit Sub

            Dim dStr As String
            For i = 1 To UBound(cenData, 1)
                If Trim$(NzToText(cenData(i, cCenId))) <> "" Then
                    AddRowMap i      ' Cenovnik je append-only (bez izmene) -> map samo za klik

                    dStr = ""
                    If cCenDat > 0 Then
                        If IsDate(cenData(i, cCenDat)) Then
                            dStr = Format$(CDate(cenData(i, cCenDat)), "d.m.yyyy")
                        Else
                            dStr = NzToText(cenData(i, cCenDat))
                        End If
                    End If

                    lstData.AddItem NzToText(cenData(i, cCenId))
                    lstData.List(lstData.ListCount - 1, 1) = dStr
                    lstData.List(lstData.ListCount - 1, 2) = NzToText(cenData(i, cCenVr))
                    lstData.List(lstData.ListCount - 1, 3) = NzToText(cenData(i, cCenSo))
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(cenData(i, cCenKl))
                    lstData.List(lstData.ListCount - 1, 5) = NzToText(cenData(i, cCenCe))
                End If
            Next i

        Case "Korisnici"
            lstData.ColumnCount = 6

            Dim koID As Long, koU As Long, koIme As Long, koUl As Long, koAk As Long
            koID = GetColumnIndex(TBL_KORISNICI, COL_KOR_ID)
            koU = GetColumnIndex(TBL_KORISNICI, COL_KOR_USERNAME)
            koIme = GetColumnIndex(TBL_KORISNICI, COL_KOR_IME)
            koUl = GetColumnIndex(TBL_KORISNICI, COL_KOR_ULOGA)
            koAk = GetColumnIndex(TBL_KORISNICI, COL_KOR_AKTIVAN)

            If koID = 0 Or koU = 0 Then
                Err.Raise vbObjectError + 7211, "frmStammdaten.LoadList", _
                          "Nedostaju kolone u tblKorisnici."
            End If

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, koID))) <> "" Then
                    AddRowMap i

                    lstData.AddItem NzToText(data(i, koID))
                    lstData.List(lstData.ListCount - 1, 1) = NzToText(data(i, koU))
                    lstData.List(lstData.ListCount - 1, 2) = NzToText(data(i, koIme))
                    lstData.List(lstData.ListCount - 1, 3) = NzToText(data(i, koUl))
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, koAk))

                    If StrComp(Trim$(NzToText(data(i, koUl))), ULOGA_ADMIN, vbTextCompare) = 0 Then
                        lstData.List(lstData.ListCount - 1, 5) = "SVE (admin)"
                    Else
                        lstData.List(lstData.ListCount - 1, 5) = OblastiCsvFromRow(data, i)
                    End If
                End If
            Next i

        Case "Kulture"
            ' Po IMENU (ne pozicijski): tblKulture ima audit kolone izmedju
            ' TipAmbalaze i Pragova, pa fiksne pozicije ne vaze na svim instalacijama.
            ' Kolone liste: 0=ID 1=Vrsta 2=Sorta 3=Gajbica 4=Aktivan 5=TipAmb
            '               6=PragUpoz 7=PragBlok (klik-load u lstData_Click prati ovo).
            lstData.ColumnCount = 8

            Dim cKulID As Long, cKulVr As Long, cKulSo As Long, cKulGa As Long
            Dim cKulAk As Long, cKulTa As Long, cKulPu As Long, cKulPb As Long
            cKulID = GetColumnIndex(TBL_KULTURE, "KulturaID")
            cKulVr = GetColumnIndex(TBL_KULTURE, "VrstaVoca")
            cKulSo = GetColumnIndex(TBL_KULTURE, "SortaVoca")
            cKulGa = GetColumnIndex(TBL_KULTURE, COL_KUL_GAJBICA_PALETA)
            cKulAk = GetColumnIndex(TBL_KULTURE, "Aktivan")
            cKulTa = GetColumnIndex(TBL_KULTURE, COL_KUL_TIP_AMBALAZE)
            cKulPu = GetColumnIndex(TBL_KULTURE, COL_KUL_PRAG_PROSEK_UPOZ)
            cKulPb = GetColumnIndex(TBL_KULTURE, COL_KUL_PRAG_PROSEK_BLOK)

            If cKulID = 0 Then
                Err.Raise vbObjectError + 7212, "frmStammdaten.LoadList", _
                          "Nedostaje kolona KulturaID u tblKulture."
            End If

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, cKulID))) <> "" Then
                    AddRowMap i
                    lstData.AddItem NzToText(data(i, cKulID))
                    If cKulVr > 0 Then lstData.List(lstData.ListCount - 1, 1) = NzToText(data(i, cKulVr))
                    If cKulSo > 0 Then lstData.List(lstData.ListCount - 1, 2) = NzToText(data(i, cKulSo))
                    If cKulGa > 0 Then lstData.List(lstData.ListCount - 1, 3) = NzToText(data(i, cKulGa))
                    If cKulAk > 0 Then lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, cKulAk))
                    If cKulTa > 0 Then lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, cKulTa))
                    If cKulPu > 0 Then lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, cKulPu))
                    If cKulPb > 0 Then lstData.List(lstData.ListCount - 1, 7) = NzToText(data(i, cKulPb))
                End If
            Next i

        Case Else
            lstData.ColumnCount = UBound(m_Headers) - LBound(m_Headers) + 1
            maxCols = lstData.ColumnCount

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, 1))) <> "" Then
                    AddRowMap i

                    lstData.AddItem NzToText(data(i, 1))

                    For j = 2 To Application.Min(UBound(data, 2), maxCols)
                        lstData.List(lstData.ListCount - 1, j - 1) = NzToText(data(i, j))
                    Next j
                End If
            Next i

    End Select

    Exit Sub

EH:
    LogErr "frmStammdaten.LoadList"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_UCITAVANJU") & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' AUSWAHL IN LISTE ? Felder fuellen
' ============================================================

Private Sub lstData_Click()
    If lstData.ListIndex < 0 Then Exit Sub

    ResetGeoClearConfirm
    ClearGeoStatus

    m_SelectedRow = GetMappedSelectedRow()
    If m_SelectedRow = 0 Then Exit Sub

    Dim data As Variant
    Dim koopNaziv As String
    Dim kID As String

    Select Case Me.Tag

        Case "Kooperanti"
            data = GetTableData(m_TableName)
            If IsEmpty(data) Then Exit Sub

            txtField1.value = NzToText(data(m_SelectedRow, 2))   ' Ime
            txtField2.value = NzToText(data(m_SelectedRow, 3))   ' Prezime
            txtField3.value = NzToText(data(m_SelectedRow, 4))   ' Mesto
            txtField4.value = NzToText(data(m_SelectedRow, 5))   ' Telefon

            SafeSetCombo cmbField1, CStr(LookupValue(TBL_STANICE, "StanicaID", NzToText(data(m_SelectedRow, 6)), "Naziv"))

            txtField6.value = NzToText(data(m_SelectedRow, 8))   ' BPGBroj
            txtField7.value = NzToText(data(m_SelectedRow, 9))   ' TekuciRacun
            txtField8.value = NzToText(data(m_SelectedRow, 10))  ' Pin
            txtField9.value = NzToText(data(m_SelectedRow, 11))  ' Adresa
            txtField10.value = NzToText(data(m_SelectedRow, 12)) ' JMBG

        Case "Stanice"
            txtField1.value = lstData.List(lstData.ListIndex, 1) ' Naziv
            txtField2.value = lstData.List(lstData.ListIndex, 2) ' Mesto/Telefon (Kontakt)
            txtField3.value = lstData.List(lstData.ListIndex, 3) ' Telefon (Kontakt)
            txtField4.value = lstData.List(lstData.ListIndex, 5) ' Ime
            txtField5.value = lstData.List(lstData.ListIndex, 6) ' Prezime
            txtField6.value = lstData.List(lstData.ListIndex, 7) ' Pin
            SafeSetCombo cmbField1, lstData.List(lstData.ListIndex, 8) ' JeHladnjaca

        Case "Korisnici"
            data = GetTableData(m_TableName)
            If IsEmpty(data) Then Exit Sub

            Dim ciU As Long, ciIme As Long, ciPin As Long, ciUl As Long, ciAk As Long, ciSt As Long
            ciU = GetColumnIndex(TBL_KORISNICI, COL_KOR_USERNAME)
            ciIme = GetColumnIndex(TBL_KORISNICI, COL_KOR_IME)
            ciPin = GetColumnIndex(TBL_KORISNICI, COL_KOR_PIN)
            ciUl = GetColumnIndex(TBL_KORISNICI, COL_KOR_ULOGA)
            ciAk = GetColumnIndex(TBL_KORISNICI, COL_KOR_AKTIVAN)
            ciSt = GetColumnIndex(TBL_KORISNICI, COL_KOR_STANICA)

            txtField1.value = NzToText(data(m_SelectedRow, ciU))
            txtField2.value = NzToText(data(m_SelectedRow, ciIme))
            txtField3.value = ""   ' PIN se ne prikazuje; prazno pri izmeni = bez promene
            SafeSetCombo cmbField1, NzToText(data(m_SelectedRow, ciUl))                ' Uloga
            SafeSetCombo cmbField2, IIf(UCase$(Trim$(NzToText(data(m_SelectedRow, ciAk)))) = "NE", "NE", "DA")  ' Aktivan
            ' Stanica: prikazi Naziv iz sacuvanog StanicaID (skrivena kolona = ID).
            SafeSetCombo cmbField3, CStr(LookupValue(TBL_STANICE, "StanicaID", NzToText(data(m_SelectedRow, ciSt)), "Naziv"))

            ' Oblasti -- desni DA/NE dropdowni iz reda tabele.
            Dim oblSel As Variant, cbSel As MSForms.ComboBox
            For Each oblSel In modAuth.OblastiList()
                Set cbSel = Nothing
                On Error Resume Next
                Set cbSel = Me.Controls("cmbObl_" & CStr(oblSel))
                On Error GoTo 0
                If Not cbSel Is Nothing Then
                    SafeSetCombo cbSel, OblastValueFromRow(data, m_SelectedRow, CStr(oblSel))
                End If
            Next oblSel

        Case "Kupci"
            data = GetTableData(m_TableName)
            If IsEmpty(data) Then Exit Sub

            txtField1.value = NzToText(data(m_SelectedRow, 2))   ' Naziv
            txtField2.value = NzToText(data(m_SelectedRow, 3))   ' Ulica
            txtField3.value = NzToText(data(m_SelectedRow, 4))   ' Mesto
            txtField4.value = NzToText(data(m_SelectedRow, 5))   ' PostanskiBroj
            txtField5.value = NzToText(data(m_SelectedRow, 6))   ' Drzava
            txtField6.value = NzToText(data(m_SelectedRow, 7))   ' PIB
            txtField7.value = NzToText(data(m_SelectedRow, 8))   ' MaticniBroj
            txtField8.value = NzToText(data(m_SelectedRow, 9))   ' Email
            txtField9.value = NzToText(data(m_SelectedRow, 10))  ' Hladnjaca
            txtField10.value = NzToText(data(m_SelectedRow, 12)) ' TekuciRacun

        Case "Vozaci"
            txtField1.value = lstData.List(lstData.ListIndex, 1) ' Ime
            txtField2.value = lstData.List(lstData.ListIndex, 2) ' Prezime
            txtField3.value = lstData.List(lstData.ListIndex, 3) ' Telefon
            txtField4.value = lstData.List(lstData.ListIndex, 5) ' PIN

        Case "Parcele"
            data = GetTableData(m_TableName)
            If IsEmpty(data) Then Exit Sub

            Dim koopID As String
            koopID = NzToText(data(m_SelectedRow, 2))

            koopNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime")) & " " & _
                        CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime")) & _
                        " (" & koopID & ")"

            SafeSetCombo cmbField1, koopNaziv                    ' Kooperant
            txtField2.value = NzToText(data(m_SelectedRow, 3))  ' KatBroj
            txtField3.value = NzToText(data(m_SelectedRow, 4))  ' KatOpstina
            SafeSetCombo cmbField2, NzToText(data(m_SelectedRow, 5))  ' Kultura
            txtField5.value = NzToText(data(m_SelectedRow, 6))  ' PovrsinaHa
            SafeSetCombo cmbField3, NzToText(data(m_SelectedRow, 7))  ' GGAPStatus
            txtField7.value = NzToText(data(m_SelectedRow, 20)) ' Napomena

        Case "Artikli"
            txtField1.value = lstData.List(lstData.ListIndex, 1)   ' Naziv
            SafeSetCombo cmbField5, lstData.List(lstData.ListIndex, 2)   ' Tip
            SafeSetCombo cmbField6, lstData.List(lstData.ListIndex, 3)   ' JedinicaMere
            txtField4.value = lstData.List(lstData.ListIndex, 4)   ' CenaPoJedinici
            txtField6.value = lstData.List(lstData.ListIndex, 5)   ' DozaPoHa
            SafeSetCombo cmbField1, lstData.List(lstData.ListIndex, 6)   ' Kultura
            txtField7.value = lstData.List(lstData.ListIndex, 7)   ' Pakovanje

        Case "Kulture"
            txtField1.value = lstData.List(lstData.ListIndex, 1)   ' VrstaVoca
            txtField2.value = lstData.List(lstData.ListIndex, 2)   ' SortaVoca
            txtField3.value = lstData.List(lstData.ListIndex, 3)   ' GajbicaPoPaleti
            SafeSetCombo cmbField1, lstData.List(lstData.ListIndex, 5)   ' TipAmbalaze (podraz.)
            txtField5.value = lstData.List(lstData.ListIndex, 6)   ' PragProsekUpoz
            txtField6.value = lstData.List(lstData.ListIndex, 7)   ' PragProsekBlok

        Case "TipAmbalaze", "TipPalete"
            txtField1.value = lstData.List(lstData.ListIndex, 0)   ' Tip (PK)
            txtField2.value = lstData.List(lstData.ListIndex, 1)   ' Tezina

        Case "Kutije", "Kese"
            txtField1.value = lstData.List(lstData.ListIndex, 0)   ' Tip (PK)
            txtField2.value = lstData.List(lstData.ListIndex, 1)   ' Tezina

        Case "VrstaGP"
            txtField1.value = lstData.List(lstData.ListIndex, 0)   ' Tip (PK)

        Case "Cenovnik"
            ' Append-only istorija -- klik samo prikazuje (bez izmene).
            SafeSetCombo cmbField1, lstData.List(lstData.ListIndex, 2)   ' Vrsta
            cmbField2.Clear
            If Trim$(cmbField1.value) <> "" Then
                FillCmb cmbField2, GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", cmbField1.value)
            End If
            SafeSetCombo cmbField2, lstData.List(lstData.ListIndex, 3)   ' Sorta
            SafeSetCombo cmbField3, lstData.List(lstData.ListIndex, 4)   ' Klasa
            txtField5.value = lstData.List(lstData.ListIndex, 5)         ' Cena
    End Select

    UpdateGeoControlsVisibility

End Sub

' Kaskada Vrsta -> Sorta (samo Cenovnik). Za ostale tabove cmbField1
' ima drugu ulogu, pa tu ne radimo nista.
Private Sub cmbField1_Change()
    On Error Resume Next

    ' Korisnici: uloga Admin -> sve oblasti DA i zakljucane (admin vidi sve).
    If Me.Tag = "Korisnici" Then
        ApplyAdminOblastiLock
        Exit Sub
    End If

    If Me.Tag <> "Cenovnik" Then Exit Sub

    Dim trenutnaSorta As String
    trenutnaSorta = Trim$(cmbField2.value)

    cmbField2.Clear
    If Trim$(cmbField1.value) <> "" Then
        FillCmb cmbField2, GetLookupList(TBL_KULTURE, "SortaVoca", "VrstaVoca", cmbField1.value)
    End If

    ' Zadrzi izbor ako sorta i dalje postoji u novoj listi.
    If Len(trenutnaSorta) > 0 Then SafeSetCombo cmbField2, trenutnaSorta
End Sub

' Uloga Admin -> sve oblasti DA i zakljucane (bypass kao modAuth.KorisnikImaPravo);
' Korisnik -> otkljucaj da admin moze birati po oblasti. Idempotentno; bez efekta
' dok desni dropdowni jos nisu izgradjeni (BuildKorisniciOblasti).
Private Sub ApplyAdminOblastiLock()
    On Error GoTo EH
    Dim isAdmin As Boolean
    isAdmin = (StrComp(Trim$(cmbField1.value), ULOGA_ADMIN, vbTextCompare) = 0)

    Dim oblLock As Variant, cbL As MSForms.ComboBox
    For Each oblLock In modAuth.OblastiList()
        Set cbL = Nothing
        On Error Resume Next                       ' kontrola mozda jos nije izgradjena
        Set cbL = Me.Controls("cmbObl_" & CStr(oblLock))
        On Error GoTo EH
        If Not cbL Is Nothing Then
            If isAdmin Then
                cbL.value = "DA"
                cbL.Locked = True
            Else
                cbL.Locked = False
            End If
        End If
    Next oblLock
    Exit Sub
EH:
    LogErr "frmStammdaten.ApplyAdminOblastiLock"
End Sub

' ============================================================
' HINZUFUeGEN
' ============================================================

' btnDodaj_Click i btnIzmeni_Click su preseljeni nize, uz ostale unosne
' procedure (v. odeljak UNOS I IZMENA).

' ============================================================
' AeNDERN
' ============================================================

' ============================================================
' UNOS I IZMENA -- forma vise NE nosi provere ni upis.
'
' Sve sto je do v6-ui-189 stajalo u btnDodaj_Click (544 linije) i
' btnIzmeni_Click (463) preseljeno je u modMaticniUnos, koji od sada zovu OBE
' strane: i ova forma i novi ekrani. Razlog i zasto ovde NIJE ponovljena odluka
' iz Faze B (dve kopije namerno): docs/UI_MIGRACIJA_KATALOG.md 26.5.
'
' Formi je ostalo samo ono sto forma jeste: koja kontrola nosi koje polje.
' ============================================================

' Kljuc sekcije koji razume modMaticniIzvor / modMaticniUnos. Prazno = sekcija
' koju modul ne pokriva (Korisnici, do M4).
Private Function SekcijaKljuc() As String
    SekcijaKljuc = modMaticniIzvor.MatKljucIzLegacyTag(Me.Tag)
End Function

' Ime kontrole koja nosi dato polje. JEDINO mesto na kom forma jos zna svoj
' raspored -- preslikano iz Setup* procedura, red po red.
Private Function KontrolaZaPolje(ByVal p As String) As String
    Select Case Me.Tag
        Case "Kooperanti"
            Select Case p
                Case "ime": KontrolaZaPolje = "txtField1"
                Case "prezime": KontrolaZaPolje = "txtField2"
                Case "mesto": KontrolaZaPolje = "txtField3"
                Case "telefon": KontrolaZaPolje = "txtField4"
                Case "stanica": KontrolaZaPolje = "cmbField1"
                Case "bpg": KontrolaZaPolje = "txtField6"
                Case "racun": KontrolaZaPolje = "txtField7"
                Case "pin": KontrolaZaPolje = "txtField8"
                Case "adresa": KontrolaZaPolje = "txtField9"
                Case "jmbg": KontrolaZaPolje = "txtField10"
            End Select
        Case "Stanice"
            Select Case p
                Case "naziv": KontrolaZaPolje = "txtField1"
                Case "mesto": KontrolaZaPolje = "txtField2"
                Case "telefon": KontrolaZaPolje = "txtField3"
                Case "kime": KontrolaZaPolje = "txtField4"
                Case "kprezime": KontrolaZaPolje = "txtField5"
                Case "pin": KontrolaZaPolje = "txtField6"
                Case "hladnjaca": KontrolaZaPolje = "cmbField1"
            End Select
        Case "Kupci"
            Select Case p
                Case "naziv": KontrolaZaPolje = "txtField1"
                Case "ulica": KontrolaZaPolje = "txtField2"
                Case "mesto": KontrolaZaPolje = "txtField3"
                Case "posta": KontrolaZaPolje = "txtField4"
                Case "drzava": KontrolaZaPolje = "txtField5"
                Case "pib": KontrolaZaPolje = "txtField6"
                Case "mb": KontrolaZaPolje = "txtField7"
                Case "email": KontrolaZaPolje = "txtField8"
                Case "hladnjaca": KontrolaZaPolje = "txtField9"
                Case "racun": KontrolaZaPolje = "txtField10"
            End Select
        Case "Korisnici"
            Select Case p
                Case "korime": KontrolaZaPolje = "txtField1"
                Case "ime": KontrolaZaPolje = "txtField2"
                Case "pin": KontrolaZaPolje = "txtField3"
                Case "uloga": KontrolaZaPolje = "cmbField1"
                Case "aktivan": KontrolaZaPolje = "cmbField2"
                Case "stanica": KontrolaZaPolje = "cmbField3"
            End Select
        Case "Vozaci"
            Select Case p
                Case "ime": KontrolaZaPolje = "txtField1"
                Case "prezime": KontrolaZaPolje = "txtField2"
                Case "telefon": KontrolaZaPolje = "txtField3"
                Case "pin": KontrolaZaPolje = "txtField4"
            End Select
        Case "Parcele"
            Select Case p
                Case "kooperant": KontrolaZaPolje = "cmbField1"
                Case "katbroj": KontrolaZaPolje = "txtField2"
                Case "katopstina": KontrolaZaPolje = "txtField3"
                Case "kultura": KontrolaZaPolje = "cmbField2"
                Case "povrsina": KontrolaZaPolje = "txtField5"
                Case "ggap": KontrolaZaPolje = "cmbField3"
                Case "napomena": KontrolaZaPolje = "txtField7"
            End Select
        Case "Artikli"
            Select Case p
                Case "naziv": KontrolaZaPolje = "txtField1"
                Case "tip": KontrolaZaPolje = "cmbField5"
                Case "jm": KontrolaZaPolje = "cmbField6"
                Case "cena": KontrolaZaPolje = "txtField4"
                Case "doza": KontrolaZaPolje = "txtField6"
                Case "kultura": KontrolaZaPolje = "cmbField1"
                Case "pakovanje": KontrolaZaPolje = "txtField7"
            End Select
        Case "Kulture"
            Select Case p
                Case "vrsta": KontrolaZaPolje = "txtField1"
                Case "sorta": KontrolaZaPolje = "txtField2"
                Case "gajbica": KontrolaZaPolje = "txtField3"
                Case "tipamb": KontrolaZaPolje = "cmbField1"
                Case "pragupoz": KontrolaZaPolje = "txtField5"
                Case "pragblok": KontrolaZaPolje = "txtField6"
            End Select
        Case "Cenovnik"
            Select Case p
                Case "vrsta": KontrolaZaPolje = "cmbField1"
                Case "sorta": KontrolaZaPolje = "cmbField2"
                Case "klasa": KontrolaZaPolje = "cmbField3"
                Case "datum": KontrolaZaPolje = "txtField4"
                Case "cena": KontrolaZaPolje = "txtField5"
            End Select
        Case "VrstaGP"
            If p = "tip" Then KontrolaZaPolje = "txtField1"
        Case "TipAmbalaze", "TipPalete", "Kutije", "Kese"
            Select Case p
                Case "tip": KontrolaZaPolje = "txtField1"
                Case "tezina": KontrolaZaPolje = "txtField2"
            End Select
    End Select
End Function

' Recnik "kljuc polja -> vrednost kontrole", u obliku koji modMaticniUnos ocekuje.
'
' Combo koji nosi strani kljuc daje SKRIVENI ID, ne prikaz: dve stanice istog
' naziva su moguce, pa bi trazenje po nazivu pogodilo prvu. Isti razlog zbog kog
' se red bira po PK, a ne po poziciji u listi.
Private Function PokupiPolja(ByVal kljuc As String) As Object
    Dim d As Object, a As Variant, r As Variant, nm As String, pk As String
    Set d = CreateObject("Scripting.Dictionary")
    Set PokupiPolja = d
    a = modMaticniIzvor.MatPolja(kljuc)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        pk = modMaticniIzvor.PoljeF(CStr(r), 0)
        nm = KontrolaZaPolje(pk)
        If Len(nm) > 0 Then d(pk) = VrednostKontrole(nm, CStr(r))
    Next r
End Function

Private Function VrednostKontrole(ByVal nm As String, ByVal spec As String) As String
    Dim ctl As Object
    On Error GoTo Fallback
    Set ctl = Me.Controls(nm)
    Select Case modMaticniIzvor.PoljeF(spec, 5)
        Case "@stanice", "@kooperanti"
            VrednostKontrole = GetSelectedComboHiddenID(ctl)
        Case Else
            VrednostKontrole = Trim$(NzToText(ctl.value))
    End Select
    Exit Function
Fallback:
    VrednostKontrole = ""
End Function

' Vrati fokus na polje koje je pisac odbio -- forma je to radila sa SetFocus
' odmah uz svaku proveru.
Private Sub FokusNaPolje(ByVal polja As Object)
    Dim nm As String
    On Error Resume Next
    If polja Is Nothing Then Exit Sub
    If Not polja.Exists(modMaticniUnos.MAT_FOKUS) Then Exit Sub
    nm = KontrolaZaPolje(CStr(polja(modMaticniUnos.MAT_FOKUS)))
    If Len(nm) > 0 Then Me.Controls(nm).SetFocus
End Sub

Private Sub btnDodaj_Click()
    Dim kljuc As String, polja As Object, odgovor As String, noviID As String
    On Error GoTo EH

    If Me.Tag = "Korisnici" Then
        DodajKorisnika
        Exit Sub
    End If

    kljuc = SekcijaKljuc()
    If Len(kljuc) = 0 Then
        MsgBox Poruka("MATU_ERR_NEPOZNATA_SEKCIJA") & " " & Me.Tag, vbCritical, APP_NAME
        Exit Sub
    End If

    Set polja = PokupiPolja(kljuc)
    odgovor = modMaticniUnos.MatDodaj(kljuc, polja, noviID)
    If Len(odgovor) > 0 Then
        MsgBox odgovor, vbExclamation, APP_NAME
        FokusNaPolje polja
        Exit Sub
    End If

    MsgBox Poruka("MATU_OK_DODATO") & " " & noviID, vbInformation, APP_NAME
    LoadList
    If kljuc = "CENOVNIK" Then
        txtField5.value = ""        ' spremno za sledecu cenu; proizvod ostaje
    Else
        ClearFields
    End If
    Exit Sub

EH:
    LogErr "frmStammdaten.btnDodaj_Click"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_DODAVANJU") & Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnIzmeni_Click()
    Dim kljuc As String, polja As Object, odgovor As String
    On Error GoTo EH

    If lstData.ListIndex >= 0 Then m_SelectedRow = GetMappedSelectedRow()
    If m_SelectedRow = 0 Then
        MsgBox Poruka("MATU_ERR_NEMA_REDA"), vbExclamation, APP_NAME
        Exit Sub
    End If

    If Me.Tag = "Korisnici" Then
        IzmeniKorisnika
        Exit Sub
    End If

    kljuc = SekcijaKljuc()
    If Len(kljuc) = 0 Then
        MsgBox Poruka("MATU_ERR_NEPOZNATA_SEKCIJA") & " " & Me.Tag, vbCritical, APP_NAME
        Exit Sub
    End If

    Set polja = PokupiPolja(kljuc)
    odgovor = modMaticniUnos.MatIzmeni(kljuc, m_SelectedRow, polja)
    If Len(odgovor) > 0 Then
        MsgBox odgovor, vbExclamation, APP_NAME
        FokusNaPolje polja
        Exit Sub
    End If

    MsgBox Poruka("MATU_OK_IZMENJENO"), vbInformation, APP_NAME
    LoadList
    ClearFields
    m_SelectedRow = 0
    Exit Sub

EH:
    LogErr "frmStammdaten.btnIzmeni_Click"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_IZMENI") & Err.description, vbCritical, APP_NAME
End Sub

' Korisnici od M4 idu kroz modMaticniKorisnici -- istog pisca koga zove i
' ekran. Forma vise ne zna ni za PreparePin ni za recnik "DA"/"NE" ni za
' redosled kolona prava; samo pokupi ono sto je operater otkucao.
'
' Dvanaest combo-a prava se predaje pod prefiksom "obl:" -- to je jedini deo
' koji ekran NE salje, jer prava tamo imaju svoju listu. Pisac zato pise samo
' one oblasti koje je stvarno dobio.
Private Function PokupiKorisnika() As Object
    Dim d As Object, obl As Variant
    Set d = PokupiPolja("KORISNICI")
    Set PokupiKorisnika = d
    ' Stanicu je PokupiPolja vec procitalo kao skriveni ID (izvor "@stanice").
    For Each obl In modAuth.OblastiList()
        d(modMaticniKorisnici.KOR_OBL_PREFIKS & CStr(obl)) = _
            OblComboVal(CStr(obl))
    Next obl
End Function

Private Sub DodajKorisnika()
    Dim polja As Object, odgovor As String, noviID As String
    On Error GoTo EH
    Set polja = PokupiKorisnika()
    odgovor = modMaticniUnos.MatDodaj("KORISNICI", polja, noviID)
    If Len(odgovor) > 0 Then
        MsgBox odgovor, vbExclamation, APP_NAME
        FokusNaPolje polja
        Exit Sub
    End If
    MsgBox Poruka("MATU_OK_DODATO") & " " & noviID, vbInformation, APP_NAME
    LoadList
    ClearFields
    KorisniciSetDefaults
    Exit Sub
EH:
    LogErr "frmStammdaten.DodajKorisnika"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_DODAVANJU") & Err.description, vbCritical, APP_NAME
End Sub

Private Sub IzmeniKorisnika()
    Dim polja As Object, odgovor As String
    On Error GoTo EH
    Set polja = PokupiKorisnika()
    odgovor = modMaticniUnos.MatIzmeni("KORISNICI", m_SelectedRow, polja)
    If Len(odgovor) > 0 Then
        MsgBox odgovor, vbExclamation, APP_NAME
        FokusNaPolje polja
        Exit Sub
    End If
    MsgBox Poruka("MATU_OK_IZMENJENO"), vbInformation, APP_NAME
    LoadList
    ClearFields
    m_SelectedRow = 0
    Exit Sub
EH:
    LogErr "frmStammdaten.IzmeniKorisnika"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_IZMENI") & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' NAVIGATION & HELPER
' ============================================================

Private Sub btnPovratak_Click()
    On Error GoTo EH

    ButtonActive btnPovratak

    frmOtkupAPP.ReturnToDashboard "Mati" & ChrW(269) & "ni podaci zatvoreni."
    Unload Me

    Exit Sub

EH:
    LogErr "frmStammdaten.btnPovratak_Click"
    On Error Resume Next
    Unload Me
    On Error GoTo 0
End Sub

Private Sub UserForm_Deactivate()
    On Error Resume Next
    MouseWheel_Detach
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    MouseWheel_Detach
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    MouseWheel_Detach

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Mati" & ChrW(269) & "ni podaci zatvoreni."
    End If

    On Error GoTo 0
End Sub

Private Sub ClearFields()
    txtField1.value = ""
    txtField2.value = ""
    txtField3.value = ""
    txtField4.value = ""
    txtField5.value = ""
    txtField6.value = ""
    txtField7.value = ""
    txtField8.value = ""
    txtField9.value = ""
    txtField10.value = ""

    cmbField1.value = ""
    cmbField2.value = ""
    cmbField3.value = ""
    cmbField4.value = ""
    cmbField5.value = ""
    cmbField6.value = ""

    m_SelectedRow = 0

    ' Cenovnik: zadrzi podrazumevani datum (danas) i posle ciscenja.
    On Error Resume Next
    If Me.Tag = "Cenovnik" Then txtField4.value = Format$(Date, "d.m.yyyy")
    On Error GoTo 0

    UpdateGeoControlsVisibility
    ResetGeoClearConfirm
    ClearGeoStatus

End Sub

' Validira pragove proseka (upoz/blok) iz txtField5/txtField6 za "Kulture" tab.
' Prazno -> 0 (provera iskljucena za tu kulturu). Blok mora biti >= upoz kad su
' oba > 0 (inace bi blokada gasila upozorenje). Vraca False (uz MsgBox + fokus)
' na nevalidan unos; inace puni upozOut/blokOut.
' ValidatePragoviKulture je PRESELJENO u modMaticniUnos (ukrstena provera
' pragova kulture). Ostavljena kopija bi bila drugo mesto na kom isto pravilo
' zivi -- tacno ono sto M2 uklanja.

Private Sub ResetFieldVisibility()
    On Error Resume Next

    Dim i As Long

    For i = 1 To 10
        With Me.Controls("lblField" & i)
            .caption = ""
            .Visible = False
        End With

        With Me.Controls("txtField" & i)
            .value = ""
            .Visible = False
        End With
    Next i

    For i = 1 To 6
        With Me.Controls("cmbField" & i)
            .value = ""
            .Clear
            .Visible = False
            .ColumnCount = 1
            .ColumnWidths = ""
        End With
    Next i

    SetGeoControlsVisible False

    On Error GoTo 0
End Sub

Private Sub ResetRowMap()
    Erase m_RowMap
    m_RowMapCount = 0
End Sub

Private Sub AddRowMap(ByVal tableRow As Long)
    m_RowMapCount = m_RowMapCount + 1
    ReDim Preserve m_RowMap(0 To m_RowMapCount - 1)
    m_RowMap(m_RowMapCount - 1) = tableRow
End Sub

Private Function GetMappedSelectedRow() As Long
    If lstData.ListIndex < 0 Then Exit Function
    If m_RowMapCount = 0 Then Exit Function
    If lstData.ListIndex > UBound(m_RowMap) Then Exit Function

    GetMappedSelectedRow = m_RowMap(lstData.ListIndex)
End Function

Private Function GetSelectedComboHiddenID(cmb As MSForms.ComboBox) As String
    On Error GoTo Fallback

    If cmb.ListIndex >= 0 Then
        If cmb.ColumnCount >= 2 Then
            GetSelectedComboHiddenID = Trim$(CStr(cmb.List(cmb.ListIndex, 1)))
            Exit Function
        End If
    End If

Fallback:
    GetSelectedComboHiddenID = Trim$(cmb.value)
End Function

'========================
'GEO MODULE
'========================

Private Sub SetGeoControlsVisible(ByVal isVisible As Boolean)
    On Error Resume Next

    btnGeoOpen.Visible = isVisible
    btnPasteCoords.Visible = isVisible
    btnGeoClear.Visible = isVisible
    btnGeoSave.Visible = isVisible
    btnOpenMap.Visible = isVisible
    btnOpenPolygonEditor.Visible = isVisible

    txtNCoord.Visible = isVisible
    txtECoord.Visible = isVisible

    lblNCoord.Visible = isVisible
    lblECoord.Visible = isVisible

    Dim geoStatus As Object
    Set geoStatus = Me.Controls("lblGeoStatus")
    geoStatus.Visible = isVisible And Len(Trim$(geoStatus.caption)) > 0

    On Error GoTo 0
End Sub

Private Sub UpdateGeoControlsVisibility()
    SetGeoControlsVisible (Me.Tag = "Parcele" And _
                           m_SelectedRow > 0 And _
                           lstData.ListIndex >= 0)
End Sub
Private Sub btnGeoOpen_Click()
    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    ResetGeoClearConfirm
    ClearGeoStatus

    If Not HasSelectedParcelaForGeo() Then Exit Sub

    Dim katBroj As String
    Dim katOpstina As String
    Dim searchText As String

    katBroj = Trim$(NzToText(lstData.List(lstData.ListIndex, 2)))
    katOpstina = Trim$(NzToText(lstData.List(lstData.ListIndex, 3)))

    If Len(katBroj) = 0 Or Len(katOpstina) = 0 Then
        SetGeoStatus Poruka("STM_MSG_PARCELA_NEMA_KATASTARSKI"), True
        Exit Sub
    End If

    searchText = katBroj & " " & Replace(katOpstina, "KO ", "")

    CopyToClipboard searchText
    ThisWorkbook.FollowHyperlink modMaticniGeo.GEO_URL_SRBIJA

    SetGeoStatus "GeoSrbija otvorena. Pretraga je kopirana: " & searchText, False
    Exit Sub

EH:
    LogErr "frmStammdaten.btnGeoOpen_Click"
    SetGeoStatus Poruka("STM_MSG_GRESKA_PRI_OTVARANJU"), True
End Sub

Private Sub btnGeoSave_Click()
    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    ResetGeoClearConfirm
    ClearGeoStatus

    If Not HasSelectedParcelaForGeo() Then Exit Sub

    Dim nVal As Double
    Dim eVal As Double

    If Not TryParseDouble(txtNCoord.value, nVal) Then
        SetGeoStatus "Unesi validnu N koordinatu.", True
        txtNCoord.SetFocus
        Exit Sub
    End If

    If Not TryParseDouble(txtECoord.value, eVal) Then
        SetGeoStatus "Unesi validnu E koordinatu.", True
        txtECoord.SetFocus
        Exit Sub
    End If

    If nVal <= 0 Or eVal <= 0 Then
        SetGeoStatus "Koordinate moraju biti pozitivne vrednosti.", True
        Exit Sub
    End If

    Dim parcelaID As String
    parcelaID = GetSelectedParcelaID()

    If Len(parcelaID) = 0 Then
        SetGeoStatus "Izabrana parcela nema ParcelaID.", True
        Exit Sub
    End If

    SaveParcelGeoPoint m_SelectedRow, nVal, eVal

    LoadList
    ReselectParcelaInList parcelaID
    ClearGeoFields

    SetGeoStatus "Geo podaci su sacuvani lokalno.", False
    Exit Sub

EH:
    LogErr "frmStammdaten.btnGeoSave_Click"
    SetGeoStatus Poruka("STM_MSG_GRESKA_PRI_CUVANJU"), True
End Sub

Private Sub btnGeoClear_Click()
    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    ClearGeoStatus

    If Not HasSelectedParcelaForGeo() Then Exit Sub

    If Not mGeoClearConfirmPending Then
        mGeoClearConfirmPending = True
        btnGeoClear.caption = "Potvrdi brisanje"
        SetGeoStatus Poruka("STM_MSG_KLIKNI_JOS_JEDNOM"), True
        Exit Sub
    End If

    Dim parcelaID As String
    parcelaID = GetSelectedParcelaID()

    If Len(parcelaID) = 0 Then
        SetGeoStatus "Izabrana parcela nema ParcelaID.", True
        Exit Sub
    End If

    ClearParcelGeo m_SelectedRow

    ResetGeoClearConfirm
    LoadList
    ReselectParcelaInList parcelaID
    ClearGeoFields

    SetGeoStatus "Geo podaci su obrisani.", False
    Exit Sub

EH:
    ResetGeoClearConfirm
    LogErr "frmStammdaten.btnGeoClear_Click"
    SetGeoStatus Poruka("STM_MSG_GRESKA_PRI_BRISANJU"), True
End Sub

Private Sub btnPasteCoords_Click()
    On Error GoTo EH

    ResetGeoClearConfirm
    ClearGeoStatus

    If Me.Tag <> "Parcele" Then Exit Sub
    If Not HasSelectedParcelaForGeo() Then Exit Sub

    Dim txt As String
    txt = Trim$(GetClipboardText())

    If txt = "" Then
        SetGeoStatus "Clipboard je prazan.", True
        Exit Sub
    End If

    Dim nVal As Double
    Dim eVal As Double

    If Not modMaticniGeo.GeoIzTeksta(txt, nVal, eVal) Then
        SetGeoStatus "Nisu pronadene validne koordinate u clipboard-u.", True
        Exit Sub
    End If

    txtNCoord.value = FormatCoordForTextBox(nVal)
    txtECoord.value = FormatCoordForTextBox(eVal)

    SetGeoStatus "Koordinate su ucitane iz clipboard-a.", False
    Exit Sub

EH:
    LogErr "frmStammdaten.btnPasteCoords_Click"
    SetGeoStatus Poruka("STM_MSG_GRESKA_PRI_UCITAVANJU"), True
End Sub

Private Sub btnOpenMap_Click()
    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    ResetGeoClearConfirm
    ClearGeoStatus

    If Not HasSelectedParcelaForGeo() Then Exit Sub

    Dim data As Variant
    data = GetTableData(TBL_PARCELE)

    If IsEmpty(data) Then
        SetGeoStatus "Tabela parcela je prazna.", True
        Exit Sub
    End If

    Dim latIdx As Long
    Dim lngIdx As Long

    latIdx = GetColumnIndex(TBL_PARCELE, COL_PAR_LAT)
    lngIdx = GetColumnIndex(TBL_PARCELE, COL_PAR_LNG)

    If latIdx = 0 Or lngIdx = 0 Then
        SetGeoStatus "Lat/Lng kolone nisu pronadene.", True
        Exit Sub
    End If

    Dim lat As Double
    Dim lng As Double

    If Not TryParseDouble(NzToText(data(m_SelectedRow, latIdx)), lat) Or _
       Not TryParseDouble(NzToText(data(m_SelectedRow, lngIdx)), lng) Then
        SetGeoStatus "Parcela nema validne Lat/Lng geo podatke.", True
        Exit Sub
    End If

    If OpenGoogleMaps(lat, lng) Then
        SetGeoStatus "Google Maps otvoren.", False
    Else
        SetGeoStatus "Google Maps nije mogao biti otvoren. Pogledaj log.", True
    End If

    Exit Sub

EH:
    LogErr "frmStammdaten.btnOpenMap_Click"
    SetGeoStatus Poruka("STM_MSG_GRESKA_PRI_OTVARANJU_2"), True
End Sub

Private Sub btnOpenPolygonEditor_Click()
    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    ResetGeoClearConfirm
    ClearGeoStatus

    If Not HasSelectedParcelaForGeo() Then Exit Sub

    Dim data As Variant
    data = GetTableData(TBL_PARCELE)

    If IsEmpty(data) Then
        SetGeoStatus "Tabela parcela je prazna.", True
        Exit Sub
    End If

    Dim idIdx As Long
    idIdx = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)

    If idIdx = 0 Then
        SetGeoStatus "ParcelaID kolona nije pronadena.", True
        Exit Sub
    End If

    Dim parcelaID As String
    parcelaID = Trim$(NzToText(data(m_SelectedRow, idIdx)))

    If parcelaID = "" Then
        SetGeoStatus "Izabrana parcela nema ParcelaID.", True
        Exit Sub
    End If

    Me.MousePointer = fmMousePointerHourGlass
    SetGeoStatus "Sinhronizujem parcelu u Google...", False
    DoEvents

    If Not SyncSelectedParcelaToGoogle(parcelaID) Then
        Me.MousePointer = fmMousePointerDefault
        SetGeoStatus "Parcela nije sinhronizovana. Editor nije otvoren.", True
        Exit Sub
    End If

    If OpenParcelPolygonEditor(parcelaID) Then
        SetGeoStatus "Polygon editor otvoren.", False
    Else
        SetGeoStatus "Polygon editor nije mogao biti otvoren. Pogledaj log.", True
    End If

    Me.MousePointer = fmMousePointerDefault
    Exit Sub

EH:
    ' LogErr PRE On Error naredbi -- one resetuju Err.
    LogErr "frmStammdaten.btnOpenPolygonEditor_Click"
    On Error Resume Next
    Me.MousePointer = fmMousePointerDefault
    On Error GoTo 0

    SetGeoStatus Poruka("STM_MSG_GRESKA_PRI_OTVARANJU_3"), True
End Sub
Private Sub SetGeoStatus(ByVal message As String, Optional ByVal isError As Boolean = False)
    On Error Resume Next

    Dim ctl As Object
    Set ctl = Me.Controls("lblGeoStatus")

    If Not ctl Is Nothing Then
        ctl.caption = message
        ctl.Visible = (Len(Trim$(message)) > 0)

        If isError Then
            ctl.ForeColor = CLR_ERROR()       ' bilo RGB(255, 80, 80)
            ctl.Font.Bold = True
        Else
            ctl.ForeColor = CLR_SUCCESS()     ' bilo RGB(120, 220, 140)
            ctl.Font.Bold = False
        End If
    End If

    On Error GoTo 0
End Sub

Private Sub ClearGeoStatus()
    SetGeoStatus vbNullString, False
End Sub

Private Sub ResetGeoClearConfirm()
    On Error Resume Next

    mGeoClearConfirmPending = False
    btnGeoClear.caption = Poruka("STM_LBL_OBRISI_GEO")

    On Error GoTo 0
End Sub

Private Function HasSelectedParcelaForGeo() As Boolean
    If Me.Tag <> "Parcele" Then Exit Function

    If m_SelectedRow = 0 Or lstData.ListIndex < 0 Then
        SetGeoStatus "Izaberi parcelu iz liste.", True
        Exit Function
    End If

    HasSelectedParcelaForGeo = True
End Function

'========================
'HELPERS
'========================
Private Sub SafeSetCombo(cmb As MSForms.ComboBox, ByVal v As String)
    On Error GoTo EH

    Dim i As Long
    Dim wanted As String

    wanted = Trim$(v)

    If Len(wanted) = 0 Then
        cmb.ListIndex = -1
        cmb.value = ""
        Exit Sub
    End If

    ' 1) First try exact visible value match.
    For i = 0 To cmb.ListCount - 1
        If Trim$(CStr(cmb.List(i, 0))) = wanted Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i

    ' 2) If ComboBox has hidden ID column, try matching against column 1.
    If cmb.ColumnCount >= 2 Then
        For i = 0 To cmb.ListCount - 1
            If Trim$(CStr(cmb.List(i, 1))) = wanted Then
                cmb.ListIndex = i
                Exit Sub
            End If
        Next i
    End If

    ' 3) If wanted looks like "Name (ID)", try extracting ID and matching hidden column.
    Dim extractedID As String
    extractedID = ExtractIDFromDisplaySafe(wanted)

    If Len(extractedID) > 0 And cmb.ColumnCount >= 2 Then
        For i = 0 To cmb.ListCount - 1
            If Trim$(CStr(cmb.List(i, 1))) = extractedID Then
                cmb.ListIndex = i
                Exit Sub
            End If
        Next i
    End If

    ' 4) No match.
    cmb.ListIndex = -1
    cmb.value = ""

    Exit Sub

EH:
    LogErr "frmStammdaten.SafeSetCombo"

    On Error Resume Next
    cmb.ListIndex = -1
    cmb.value = ""
End Sub

Private Function ExtractIDFromDisplaySafe(ByVal displayText As String) As String
    On Error GoTo EH

    Dim p1 As Long
    Dim p2 As Long

    p1 = InStrRev(displayText, "(")
    p2 = InStrRev(displayText, ")")

    If p1 > 0 And p2 > p1 Then
        ExtractIDFromDisplaySafe = Trim$(Mid$(displayText, p1 + 1, p2 - p1 - 1))
    Else
        ExtractIDFromDisplaySafe = ""
    End If

    Exit Function

EH:
    ExtractIDFromDisplaySafe = ""
End Function

Private Sub LoadStaniceIntoCombo(Optional ByRef cmb As MSForms.ComboBox)
    On Error GoTo EH

    ' Podrazumevano cmbField1 (Kooperanti); Korisnici prosledjuje cmbField3.
    If cmb Is Nothing Then Set cmb = cmbField1

    Dim data As Variant
    Dim i As Long

    cmb.Clear
    cmb.ColumnCount = 2
    cmb.ColumnWidths = "150 pt;0 pt"

    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then Exit Sub

    Dim colID As Long
    Dim colNaziv As Long

    colID = GetColumnIndex(TBL_STANICE, "StanicaID")
    colNaziv = GetColumnIndex(TBL_STANICE, "Naziv")

    If colID = 0 Or colNaziv = 0 Then
        MsgBox "Nedostaju kolone StanicaID/Naziv u tabeli stanica.", vbCritical, APP_NAME
        Exit Sub
    End If

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colID))) <> "" Then
            cmb.AddItem NzToText(data(i, colNaziv))
            cmb.List(cmb.ListCount - 1, 1) = NzToText(data(i, colID))
        End If
    Next i

    Exit Sub

EH:
    LogErr "frmStammdaten.LoadStaniceIntoCombo"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_UCITAVANJU_2") & Err.description, vbCritical, APP_NAME
End Sub

Private Function FormatCoordForTextBox(ByVal v As Double) As String
    FormatCoordForTextBox = Replace(Format$(v, "0.############"), ",", ".")
End Function

' TryExtractTwoCoordinates i CleanCoordToken su PRESELJENI u modMaticniGeo
' (GeoIzTeksta / OcistiToken). Isto pravilo je trebalo i novom ekranu, a dve
' kopije praga "|d| > 1000" bi se razisle prvom doradom.
Private Sub ClearGeoFields()
    On Error Resume Next

    txtNCoord.value = ""
    txtECoord.value = ""

    On Error GoTo 0
End Sub



Private Function GetSelectedParcelaID() As String
    Const SRC As String = "frmStammdaten.GetSelectedParcelaID"

    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Function
    If m_SelectedRow = 0 Then Exit Function

    Dim data As Variant
    Dim idIdx As Long

    data = GetTableData(TBL_PARCELE)
    If IsEmpty(data) Then Exit Function

    idIdx = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
    If idIdx = 0 Then Exit Function

    GetSelectedParcelaID = Trim$(NzToText(data(m_SelectedRow, idIdx)))
    Exit Function

EH:
    LogErr SRC
    GetSelectedParcelaID = ""
End Function

Private Sub ReselectParcelaInList(ByVal parcelaID As String)
    Const SRC As String = "frmStammdaten.ReselectParcelaInList"

    On Error GoTo EH

    Dim i As Long
    Dim cleanID As String

    cleanID = Trim$(parcelaID)

    If Me.Tag <> "Parcele" Or Len(cleanID) = 0 Then
        m_SelectedRow = 0
        UpdateGeoControlsVisibility
        Exit Sub
    End If

    For i = 0 To lstData.ListCount - 1
        If Trim$(NzToText(lstData.List(i, 0))) = cleanID Then
            lstData.ListIndex = i
            m_SelectedRow = GetMappedSelectedRow()
            UpdateGeoControlsVisibility
            Exit Sub
        End If
    Next i

    m_SelectedRow = 0
    UpdateGeoControlsVisibility
    Exit Sub

EH:
    LogErr SRC
    m_SelectedRow = 0
    UpdateGeoControlsVisibility
End Sub

Public Function OpenGoogleMaps(ByVal lat As Double, ByVal lng As Double) As Boolean
    On Error GoTo EH

    ' Adresa se gradi u modMaticniGeo -- isti oblik koristi i novi ekran.
    ThisWorkbook.FollowHyperlink modMaticniGeo.GeoUrlMape(lat, lng)

    OpenGoogleMaps = True
    Exit Function

EH:
    LogErr "frmStammdaten.OpenGoogleMaps"
    OpenGoogleMaps = False
End Function

Public Function OpenParcelPolygonEditor(ByVal parcelaID As String) As Boolean
    On Error GoTo EH

    Dim url As String

    url = modMaticniGeo.GeoUrlPoligon(parcelaID)
    If Len(url) = 0 Then
        LogError "frmStammdaten.OpenParcelPolygonEditor", "ParcelaID nije prosleden."
        Exit Function
    End If

    ThisWorkbook.FollowHyperlink url

    OpenParcelPolygonEditor = True
    Exit Function

EH:
    LogErr "frmStammdaten.OpenParcelPolygonEditor"
    OpenParcelPolygonEditor = False
End Function

