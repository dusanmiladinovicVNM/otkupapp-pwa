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

    m_FieldCount = 7

    lblField1.caption = Poruka("KOR_LBL_KORISNICKO_IME"): lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Ime i prezime": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "PIN (izmena: prazno=isti)": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Uloga (Admin/Korisnik)": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "Aktivan (DA/NE)": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "StanicaID (opciono)": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "Oblasti (DA, zarezom)": lblField7.Visible = True: txtField7.Visible = True
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False
End Sub

' --- Korisnici helperi (prava po oblasti, Model A) ---
Private Function NormalizeUloga(ByVal v As String) As String
    If StrComp(Trim$(v), ULOGA_ADMIN, vbTextCompare) = 0 Then
        NormalizeUloga = ULOGA_ADMIN
    Else
        NormalizeUloga = ULOGA_KORISNIK
    End If
End Function

Private Function CsvSadrzi(ByVal csv As String, ByVal oblast As String) As Boolean
    Dim parts() As String
    Dim p As Variant
    parts = Split(csv, ",")
    For Each p In parts
        If StrComp(Trim$(CStr(p)), Trim$(oblast), vbTextCompare) = 0 Then
            CsvSadrzi = True
            Exit Function
        End If
    Next p
End Function

' "DA"/"NE" za datu oblast: Admin -> uvek DA; inace prema CSV listi dozvoljenih.
Private Function OblDa(ByVal isAdmin As Boolean, ByVal csv As String, ByVal oblast As String) As String
    If isAdmin Then
        OblDa = "DA"
    ElseIf CsvSadrzi(csv, oblast) Then
        OblDa = "DA"
    Else
        OblDa = "NE"
    End If
End Function

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
                res = res & CStr(obl)
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

    m_FieldCount = 4

    lblField1.caption = "Vrsta vo" & ChrW(263) & "a": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Sorta vo" & ChrW(263) & "a": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Gajbica po paleti": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Tip ambala" & ChrW(382) & "e (podraz.)": lblField4.Visible = True: txtField4.Visible = False
    lblField5.caption = "": lblField5.Visible = False: txtField5.Visible = False
    lblField6.caption = "": lblField6.Visible = False: txtField6.Visible = False
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
    On Error GoTo EH

    Dim colName As String
    colName = AktivanColName()
    If Len(colName) = 0 Then Exit Sub

    If lstData.ListIndex >= 0 Then m_SelectedRow = GetMappedSelectedRow()
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite stavku iz liste!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim data As Variant
    data = GetTableData(m_TableName)
    If IsEmpty(data) Then Exit Sub

    Dim colIdx As Long
    colIdx = GetColumnIndex(m_TableName, colName)

    Dim cur As String, newVal As String
    cur = Trim$(NzToText(data(m_SelectedRow, colIdx)))
    If StrComp(cur, STATUS_NEAKTIVAN, vbTextCompare) = 0 Then
        newVal = STATUS_AKTIVAN
    Else
        newVal = STATUS_NEAKTIVAN
    End If

    Dim tx As clsTransaction
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot m_TableName
    RequireUpdateCell m_TableName, m_SelectedRow, colName, newVal, "frmStammdaten.OnSoftDeleteClick"
    tx.CommitTx
    Set tx = Nothing

    LoadList
    ClearFields
    m_SelectedRow = 0
    MsgBox "Status promenjen u: " & newVal, vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "frmStammdaten.OnSoftDeleteClick"
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    MsgBox Poruka("STM_ERR_GRESKA_PRI_PROMENI") & Err.description, vbCritical, APP_NAME
End Sub

' Azurira prvu kolonu iz liste alias-a koja stvarno postoji u tabeli.
' Tolerantno na schema drift: ako nijedan alias ne postoji, tiho preskace
' (ne rusi transakciju). Vraca True ako je nesto azurirano.
Private Function UpdateFirstExistingCol(ByVal SRC As String, ByVal newValue As String, _
                                        ParamArray colAliases() As Variant) As Boolean
    Dim k As Long
    For k = LBound(colAliases) To UBound(colAliases)
        If GetColumnIndex(m_TableName, CStr(colAliases(k))) > 0 Then
            RequireUpdateCell m_TableName, m_SelectedRow, CStr(colAliases(k)), newValue, SRC
            UpdateFirstExistingCol = True
            Exit Function
        End If
    Next k
End Function
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
            kupPIB = GetColumnIndex(TBL_KUPCI, "PIB")
            kupMB = GetColumnIndex(TBL_KUPCI, "MaticniBroj")
            kupEmail = GetColumnIndex(TBL_KUPCI, "Email")
            kupHlad = GetColumnIndex(TBL_KUPCI, "Hladnjaca")
            kupAktivan = GetColumnIndex(TBL_KUPCI, "Aktivan")
            kupRacun = GetColumnIndex(TBL_KUPCI, "TekuciRacun")

            If kupID = 0 Or kupNaziv = 0 Or kupUlica = 0 Or kupMesto = 0 Or _
               kupPosta = 0 Or kupDrzava = 0 Or kupPIB = 0 Or kupMB = 0 Or _
               kupEmail = 0 Or kupHlad = 0 Or kupAktivan = 0 Or kupRacun = 0 Then
                Err.Raise vbObjectError + 7202, "frmStammdaten.LoadList", _
                          "Nedostaju kolone u tblKupci."
            End If

            Dim kupacAdresa As String

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, kupID))) <> "" Then
                    AddRowMap i

                    kupacAdresa = Trim$(NzToText(data(i, kupUlica)))

                    If Len(Trim$(NzToText(data(i, kupPosta)))) > 0 Then
                        If Len(kupacAdresa) > 0 Then kupacAdresa = kupacAdresa & ", "
                        kupacAdresa = kupacAdresa & NzToText(data(i, kupPosta))
                    End If

                    If Len(Trim$(NzToText(data(i, kupMesto)))) > 0 Then
                        If Len(kupacAdresa) > 0 Then kupacAdresa = kupacAdresa & " "
                        kupacAdresa = kupacAdresa & NzToText(data(i, kupMesto))
                    End If

                    lstData.AddItem NzToText(data(i, kupID))
                    lstData.List(lstData.ListCount - 1, 1) = NzToText(data(i, kupNaziv))
                    lstData.List(lstData.ListCount - 1, 2) = kupacAdresa
                    lstData.List(lstData.ListCount - 1, 3) = NzToText(data(i, kupDrzava))
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, kupPIB))
                    lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, kupMB))
                    lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, kupEmail))
                    lstData.List(lstData.ListCount - 1, 7) = NzToText(data(i, kupHlad))
                    lstData.List(lstData.ListCount - 1, 8) = NzToText(data(i, kupAktivan))
                    lstData.List(lstData.ListCount - 1, 9) = NzToText(data(i, kupRacun))
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
            txtField4.value = NzToText(data(m_SelectedRow, ciUl))
            txtField5.value = NzToText(data(m_SelectedRow, ciAk))
            txtField6.value = NzToText(data(m_SelectedRow, ciSt))
            txtField7.value = OblastiCsvFromRow(data, m_SelectedRow)

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

' ============================================================
' HINZUFUeGEN
' ============================================================

Private Sub btnDodaj_Click()
    On Error GoTo EH

    Dim rowData As Variant
    Dim newID As String
    Dim stanicaID As String

    Select Case Me.Tag

        Case "Kooperanti"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite ime!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite prezime!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField1.value) = "" Then
                MsgBox "Izaberite stanicu!", vbExclamation, APP_NAME
                cmbField1.SetFocus
                Exit Sub
            End If

            stanicaID = GetSelectedComboHiddenID(cmbField1)

            If stanicaID = "" Or InStr(1, stanicaID, "ST-", vbTextCompare) = 0 Then
                stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbField1.value, "StanicaID"))
            End If

            If stanicaID = "" Then
                MsgBox "Nije prona" & ChrW(273) & "en StanicaID za izabranu stanicu.", vbExclamation, APP_NAME
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "KooperantID", "KOOP-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.value), _
                Trim$(txtField2.value), _
                Trim$(txtField3.value), _
                Trim$(txtField4.value), _
                stanicaID, _
                STATUS_AKTIVAN, _
                Trim$(txtField6.value), _
                Trim$(txtField7.value), _
                Trim$(txtField8.value), _
                Trim$(txtField9.value), _
                Trim$(txtField10.value) _
            )

        Case "Stanice"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite naziv!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite mesto!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "StanicaID", "ST-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.value), _
                Trim$(txtField2.value), _
                Trim$(txtField3.value), _
                STATUS_AKTIVAN, _
                Trim$(txtField4.value), _
                Trim$(txtField5.value), _
                Trim$(txtField6.value), _
                IIf(Trim$(cmbField1.value) = "", "Ne", Trim$(cmbField1.value)) _
            )

        Case "Korisnici"
            If Trim$(txtField1.value) = "" Then
                MsgBox Poruka("KOR_MSG_UNESITE_IME"), vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField3.value) = "" Then
                MsgBox "Unesite PIN!", vbExclamation, APP_NAME
                txtField3.SetFocus
                Exit Sub
            End If

            Dim postojiKor As Variant
            postojiKor = LookupValue(m_TableName, COL_KOR_USERNAME, Trim$(txtField1.value), COL_KOR_ID)
            If Not IsEmpty(postojiKor) Then
                If Len(Trim$(CStr(postojiKor))) > 0 Then
                    MsgBox "Korisnik '" & Trim$(txtField1.value) & Poruka("KOR_MSG_VEC_POSTOJI"), vbExclamation, APP_NAME
                    Exit Sub
                End If
            End If

            Dim ulogaDodaj As String
            ulogaDodaj = NormalizeUloga(txtField4.value)

            Dim admDodaj As Boolean
            admDodaj = (StrComp(ulogaDodaj, ULOGA_ADMIN, vbTextCompare) = 0)

            Dim aktivDodaj As String
            If UCase$(Trim$(txtField5.value)) = "NE" Then aktivDodaj = "NE" Else aktivDodaj = "DA"

            Dim csvDodaj As String
            csvDodaj = Trim$(txtField7.value)

            newID = GetNextID(m_TableName, COL_KOR_ID, "KOR-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.value), _
                Trim$(txtField2.value), _
                modAuth.PreparePin(Trim$(txtField3.value)), _
                ulogaDodaj, _
                aktivDodaj, _
                Trim$(txtField6.value), _
                Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
                OblDa(admDodaj, csvDodaj, OBL_OTKUP), _
                OblDa(admDodaj, csvDodaj, OBL_DOKUMENTA), _
                OblDa(admDodaj, csvDodaj, OBL_AGROHEMIJA), _
                OblDa(admDodaj, csvDodaj, OBL_IZVESTAJI), _
                OblDa(admDodaj, csvDodaj, OBL_FAKTURISANJE), _
                OblDa(admDodaj, csvDodaj, OBL_BANKA), _
                OblDa(admDodaj, csvDodaj, OBL_MARZA), _
                OblDa(admDodaj, csvDodaj, OBL_SLEDLJIVOST), _
                OblDa(admDodaj, csvDodaj, OBL_MATICNI) _
            )

        Case "Kupci"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite naziv kupca!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "KupacID", "KUP-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.value), _
                Trim$(txtField2.value), _
                Trim$(txtField3.value), _
                Trim$(txtField4.value), _
                Trim$(txtField5.value), _
                Trim$(txtField6.value), _
                Trim$(txtField7.value), _
                Trim$(txtField8.value), _
                Trim$(txtField9.value), _
                STATUS_AKTIVAN, _
                Trim$(txtField10.value) _
            )

        Case "Vozaci"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite ime vozaca!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite prezime vozaca!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "VozacID", "VOZ-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.value), _
                Trim$(txtField2.value), _
                Trim$(txtField3.value), _
                STATUS_AKTIVAN, _
                Trim$(txtField4.value) _
            )

        Case "Parcele"
            If Trim$(cmbField1.value) = "" Then
                MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
                cmbField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite katastarski broj!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Trim$(txtField3.value) = "" Then
                MsgBox Poruka("STM_MSG_UNESITE_KATASTARSKU_OPSTINU"), vbExclamation, APP_NAME
                txtField3.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField2.value) = "" Then
                MsgBox "Izaberite kulturu!", vbExclamation, APP_NAME
                cmbField2.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField3.value) = "" Then
                MsgBox "Izaberite GGAP status!", vbExclamation, APP_NAME
                cmbField3.SetFocus
                Exit Sub
            End If

            Dim koopID As String
            Dim povrsina As Double

            koopID = ExtractIDFromDisplay(cmbField1.value)

            If Not TryParseDouble(txtField5.value, povrsina) Or povrsina <= 0 Then
                MsgBox Poruka("STM_MSG_UNESITE_VALIDNU_POVRSINU"), vbExclamation, APP_NAME
                txtField5.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "ParcelaID", "PAR-")

            rowData = Array( _
                newID, _
                koopID, _
                Trim$(txtField2.value), _
                Trim$(txtField3.value), _
                Trim$(cmbField2.value), _
                povrsina, _
                Trim$(cmbField3.value), _
                "Da", _
                "", _
                "", _
                "", _
                "", _
                "", _
                "", _
                "", _
                "", _
                "", _
                "", _
                "", _
                Trim$(txtField7.value) _
            )

        Case "Artikli"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite naziv artikla!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField5.value) = "" Then
                MsgBox "Izaberite tip artikla!", vbExclamation, APP_NAME
                cmbField5.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField6.value) = "" Then
                MsgBox "Izaberite jedinicu mere!", vbExclamation, APP_NAME
                cmbField6.SetFocus
                Exit Sub
            End If

            Dim cenaArt As Double
            Dim dozaArt As Double
            Dim pakovanjeArt As Double

            If Not TryParseDouble(txtField4.value, cenaArt) Or cenaArt < 0 Then
                MsgBox "Unesite validnu cenu!", vbExclamation, APP_NAME
                txtField4.SetFocus
                Exit Sub
            End If

            If Not TryParseDouble(txtField6.value, dozaArt) Or dozaArt < 0 Then
                MsgBox "Unesite validnu dozu!", vbExclamation, APP_NAME
                txtField6.SetFocus
                Exit Sub
            End If

            If Not TryParseDouble(txtField7.value, pakovanjeArt) Or pakovanjeArt < 0 Then
                MsgBox "Unesite validno pakovanje!", vbExclamation, APP_NAME
                txtField7.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "ArtikalID", "ART-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.value), _
                Trim$(cmbField5.value), _
                Trim$(cmbField6.value), _
                cenaArt, _
                dozaArt, _
                Trim$(cmbField1.value), _
                pakovanjeArt _
            )

        Case "Kulture"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite vrstu vo" & ChrW(263) & "a!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite sortu vo" & ChrW(263) & "a!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            Dim gajbicaKul As Double
            If Trim$(txtField3.value) <> "" Then
                If Not TryParseDouble(txtField3.value, gajbicaKul) Or gajbicaKul < 0 Then
                    MsgBox "Unesite validan broj gajbica po paleti!", vbExclamation, APP_NAME
                    txtField3.SetFocus
                    Exit Sub
                End If
            End If

            newID = GetNextID(m_TableName, "KulturaID", "KUL-")

            ' tblKulture: jezgro (ID/Vrsta/Sorta) pozicijski; Gajbica/Aktivan/
            ' TipAmbalaze PO IMENU (dodate kolone na kraju, redosled nesiguran).
            Dim newRowKul As Long
            newRowKul = AppendRow(m_TableName, _
                Array(newID, Trim$(txtField1.value), Trim$(txtField2.value)))

            If newRowKul > 0 Then
                If Trim$(txtField3.value) <> "" Then _
                    UpdateCell m_TableName, newRowKul, COL_KUL_GAJBICA_PALETA, gajbicaKul
                UpdateCell m_TableName, newRowKul, "Aktivan", STATUS_AKTIVAN
                UpdateCell m_TableName, newRowKul, COL_KUL_TIP_AMBALAZE, Trim$(cmbField1.value)
                MsgBox "Dodato: " & newID, vbInformation, APP_NAME
                LoadList
                ClearFields
            Else
                MsgBox Poruka("STM_MSG_GRESKA_PRI_DODAVANJU"), vbCritical, APP_NAME
            End If
            Exit Sub

        Case "TipAmbalaze"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip ambala" & ChrW(382) & "e!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaAmb As Double
            If Not TryParseDouble(txtField2.value, tezinaAmb) Or tezinaAmb < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu gajbice (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Len(Trim$(NzToText(LookupValue(m_TableName, COL_TAMB_TIP, Trim$(txtField1.value), COL_TAMB_TIP)))) > 0 Then
                MsgBox "Tip ambala" & ChrW(382) & "e '" & Trim$(txtField1.value) & "' ve" & ChrW(263) & " postoji!", vbExclamation, APP_NAME
                Exit Sub
            End If

            newID = Trim$(txtField1.value)   ' PK je sam naziv tipa
            rowData = Array(Trim$(txtField1.value), tezinaAmb, STATUS_AKTIVAN)

        Case "TipPalete"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip palete!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaPal As Double
            If Not TryParseDouble(txtField2.value, tezinaPal) Or tezinaPal < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu palete (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Len(Trim$(NzToText(LookupValue(m_TableName, COL_TPAL_TIP, Trim$(txtField1.value), COL_TPAL_TIP)))) > 0 Then
                MsgBox "Tip palete '" & Trim$(txtField1.value) & "' ve" & ChrW(263) & " postoji!", vbExclamation, APP_NAME
                Exit Sub
            End If

            newID = Trim$(txtField1.value)
            rowData = Array(Trim$(txtField1.value), tezinaPal, STATUS_AKTIVAN)

        Case "Kutije"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip kutije!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaKut As Double
            If Not TryParseDouble(txtField2.value, tezinaKut) Or tezinaKut < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu kutije (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Len(Trim$(NzToText(LookupValue(m_TableName, COL_KUT_TIP, Trim$(txtField1.value), COL_KUT_TIP)))) > 0 Then
                MsgBox "Tip kutije '" & Trim$(txtField1.value) & "' ve" & ChrW(263) & " postoji!", vbExclamation, APP_NAME
                Exit Sub
            End If

            newID = Trim$(txtField1.value)
            rowData = Array(Trim$(txtField1.value), tezinaKut, STATUS_AKTIVAN)

        Case "Kese"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip kese!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaKes As Double
            If Not TryParseDouble(txtField2.value, tezinaKes) Or tezinaKes < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu kese (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Len(Trim$(NzToText(LookupValue(m_TableName, COL_KES_TIP, Trim$(txtField1.value), COL_KES_TIP)))) > 0 Then
                MsgBox "Tip kese '" & Trim$(txtField1.value) & "' ve" & ChrW(263) & " postoji!", vbExclamation, APP_NAME
                Exit Sub
            End If

            newID = Trim$(txtField1.value)
            rowData = Array(Trim$(txtField1.value), tezinaKes, STATUS_AKTIVAN)

        Case "VrstaGP"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip gotovog proizvoda!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Len(Trim$(NzToText(LookupValue(m_TableName, COL_VGP_TIP, Trim$(txtField1.value), COL_VGP_TIP)))) > 0 Then
                MsgBox "Tip gotovog proizvoda '" & Trim$(txtField1.value) & "' ve" & ChrW(263) & " postoji!", vbExclamation, APP_NAME
                Exit Sub
            End If

            newID = Trim$(txtField1.value)
            rowData = Array(Trim$(txtField1.value), STATUS_AKTIVAN)

        Case "Cenovnik"
            ' Append-only: dodaje novi (vazeci) red preko modCenovnik.
            If Trim$(cmbField1.value) = "" Then
                MsgBox "Izaberite vrstu vo" & ChrW(263) & "a!", vbExclamation, APP_NAME
                cmbField1.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField3.value) = "" Then
                MsgBox "Izaberite klasu!", vbExclamation, APP_NAME
                cmbField3.SetFocus
                Exit Sub
            End If

            Dim cenaCen As Double
            If Not TryParseDouble(txtField5.value, cenaCen) Or cenaCen <= 0 Then
                MsgBox "Unesite validnu cenu!", vbExclamation, APP_NAME
                txtField5.SetFocus
                Exit Sub
            End If

            Dim datCen As Date
            If Not TryParseDateValue(txtField4.value, datCen) Then datCen = Date

            Dim resCen As String
            resCen = AddCena(datCen, cmbField1.value, cmbField2.value, cmbField3.value, cenaCen)

            If resCen <> "" Then
                MsgBox "Dodata cena: " & resCen, vbInformation, APP_NAME
                LoadList
                txtField5.value = ""        ' spremno za sledecu cenu; proizvod ostaje
            Else
                MsgBox Poruka("STM_MSG_GRESKA_PRI_DODAVANJU_2"), vbCritical, APP_NAME
            End If
            Exit Sub

        Case Else
            MsgBox "Nepoznat tip mati" & ChrW(269) & "nih podataka: " & Me.Tag, vbCritical, APP_NAME
            Exit Sub

    End Select

    If AppendRow(m_TableName, rowData) > 0 Then
        MsgBox "Dodato: " & newID, vbInformation, APP_NAME

        ' MALINA: nova stanica -> par-vozac sa istim ID-em (izvestaji/ambalaza).
        ' Idempotentno + self-gated u modMalina; ne sme da obori unos stanice.
        If Me.Tag = "Stanice" And IsMalinaMode() Then
            On Error Resume Next
            EnsureVozacMirrorForStanica newID, Trim$(txtField1.value), _
                Trim$(txtField2.value), ""
            On Error GoTo EH
        End If

        LoadList
        ClearFields
    Else
        MsgBox Poruka("STM_MSG_GRESKA_PRI_DODAVANJU"), vbCritical, APP_NAME
    End If

    Exit Sub

EH:
    LogErr "frmStammdaten.btnDodaj_Click"
    MsgBox Poruka("STM_ERR_GRESKA_PRI_DODAVANJU") & Err.description, vbCritical, APP_NAME
End Sub

' ============================================================
' AeNDERN
' ============================================================

Private Sub btnIzmeni_Click()
    On Error GoTo EH

    Const SRC As String = "frmStammdaten.btnIzmeni_Click"

    If lstData.ListIndex >= 0 Then
        m_SelectedRow = GetMappedSelectedRow()
    End If

    If m_SelectedRow = 0 Then
        MsgBox "Izaberite stavku iz liste!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    Dim stID As String
    Dim koopIDEdit As String
    Dim povrsinaEdit As Double
    Dim cenaEdit As Double
    Dim dozaEdit As Double
    Dim pakovanjeEdit As Double

    Select Case Me.Tag

        Case "Kooperanti"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite ime!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite prezime!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            stID = GetSelectedComboHiddenID(cmbField1)

            If stID = "" Or InStr(1, stID, "ST-", vbTextCompare) = 0 Then
                stID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbField1.value, "StanicaID"))
            End If

            If stID = "" Then
                MsgBox "Nije prona" & ChrW(273) & "en StanicaID za izabranu stanicu.", vbExclamation, APP_NAME
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, "Ime", Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Prezime", Trim$(txtField2.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Mesto", Trim$(txtField3.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Telefon", Trim$(txtField4.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "StanicaID", stID, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "BPGBroj", Trim$(txtField6.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "TekuciRacun", Trim$(txtField7.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Pin", Trim$(txtField8.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Adresa", Trim$(txtField9.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "JMBG", Trim$(txtField10.value), SRC

            tx.CommitTx
            

        Case "Stanice"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite naziv!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            ' tblStanice kod razlicitih instalacija ima drift u nazivima kolona
            ' (npr. Ime vs KontaktIme, PIN vs Pin). Da izmena NE pukne na
            ' nepostojecoj koloni (rollback cele izmene), azuriramo prvu kolonu
            ' koja STVARNO postoji (alias-probe). Reuse: GetColumnIndex.
            UpdateFirstExistingCol SRC, Trim$(txtField1.value), "Naziv"
            UpdateFirstExistingCol SRC, Trim$(txtField2.value), "Mesto"
            UpdateFirstExistingCol SRC, Trim$(txtField3.value), "Kontakt", "Telefon"
            UpdateFirstExistingCol SRC, Trim$(txtField4.value), "Ime", "KontaktIme"
            UpdateFirstExistingCol SRC, Trim$(txtField5.value), "Prezime", "KontaktPrezime"
            UpdateFirstExistingCol SRC, Trim$(txtField6.value), "PIN", "Pin"
            UpdateFirstExistingCol SRC, Trim$(cmbField1.value), COL_STA_JE_HLADNJACA

            tx.CommitTx

        Case "Korisnici"
            If Trim$(txtField1.value) = "" Then
                MsgBox Poruka("KOR_MSG_UNESITE_IME"), vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            Dim ulogaIzm As String
            ulogaIzm = NormalizeUloga(txtField4.value)

            Dim admIzm As Boolean
            admIzm = (StrComp(ulogaIzm, ULOGA_ADMIN, vbTextCompare) = 0)

            Dim aktivIzm As String
            If UCase$(Trim$(txtField5.value)) = "NE" Then aktivIzm = "NE" Else aktivIzm = "DA"

            RequireUpdateCell m_TableName, m_SelectedRow, COL_KOR_USERNAME, Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KOR_IME, Trim$(txtField2.value), SRC
            If Len(Trim$(txtField3.value)) > 0 Then
                RequireUpdateCell m_TableName, m_SelectedRow, COL_KOR_PIN, modAuth.PreparePin(Trim$(txtField3.value)), SRC
            End If
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KOR_ULOGA, ulogaIzm, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KOR_AKTIVAN, aktivIzm, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KOR_STANICA, Trim$(txtField6.value), SRC

            Dim oblIzm As Variant
            For Each oblIzm In modAuth.OblastiList()
                RequireUpdateCell m_TableName, m_SelectedRow, CStr(oblIzm), _
                                  OblDa(admIzm, Trim$(txtField7.value), CStr(oblIzm)), SRC
            Next oblIzm

            tx.CommitTx

        Case "Kupci"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite naziv kupca!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, "Naziv", Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Ulica", Trim$(txtField2.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Mesto", Trim$(txtField3.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "PostanskiBroj", Trim$(txtField4.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Dr" & ChrW(382) & "ava", Trim$(txtField5.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "PIB", Trim$(txtField6.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "MaticniBroj", Trim$(txtField7.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Email", Trim$(txtField8.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Hladnjaca", Trim$(txtField9.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "TekuciRacun", Trim$(txtField10.value), SRC

            tx.CommitTx

        Case "Vozaci"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite ime vozaca!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite prezime vozaca!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, "Ime", Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Prezime", Trim$(txtField2.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Telefon", Trim$(txtField3.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "PIN", Trim$(txtField4.value), SRC

            tx.CommitTx

        Case "Parcele"
            If Trim$(cmbField1.value) = "" Then
                MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
                cmbField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite katastarski broj!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Trim$(txtField3.value) = "" Then
                MsgBox Poruka("STM_MSG_UNESITE_KATASTARSKU_OPSTINU"), vbExclamation, APP_NAME
                txtField3.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField2.value) = "" Then
                MsgBox "Izaberite kulturu!", vbExclamation, APP_NAME
                cmbField2.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField3.value) = "" Then
                MsgBox "Izaberite GGAP status!", vbExclamation, APP_NAME
                cmbField3.SetFocus
                Exit Sub
            End If

            If Not TryParseDouble(txtField5.value, povrsinaEdit) Or povrsinaEdit <= 0 Then
                MsgBox Poruka("STM_MSG_UNESITE_VALIDNU_POVRSINU"), vbExclamation, APP_NAME
                txtField5.SetFocus
                Exit Sub
            End If

            koopIDEdit = ExtractIDFromDisplay(cmbField1.value)

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, "KooperantID", koopIDEdit, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "KatBroj", Trim$(txtField2.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "KatOpstina", Trim$(txtField3.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Kultura", Trim$(cmbField2.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "PovrsinaHa", povrsinaEdit, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "GGAPStatus", Trim$(cmbField3.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Napomena", Trim$(txtField7.value), SRC

            tx.CommitTx

        Case "Artikli"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite naziv artikla!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField5.value) = "" Then
                MsgBox "Izaberite tip artikla!", vbExclamation, APP_NAME
                cmbField5.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField6.value) = "" Then
                MsgBox "Izaberite jedinicu mere!", vbExclamation, APP_NAME
                cmbField6.SetFocus
                Exit Sub
            End If

            If Not TryParseDouble(txtField4.value, cenaEdit) Or cenaEdit < 0 Then
                MsgBox "Unesite validnu cenu!", vbExclamation, APP_NAME
                txtField4.SetFocus
                Exit Sub
            End If

            If Not TryParseDouble(txtField6.value, dozaEdit) Or dozaEdit < 0 Then
                MsgBox "Unesite validnu dozu!", vbExclamation, APP_NAME
                txtField6.SetFocus
                Exit Sub
            End If

            If Not TryParseDouble(txtField7.value, pakovanjeEdit) Or pakovanjeEdit < 0 Then
                MsgBox "Unesite validno pakovanje!", vbExclamation, APP_NAME
                txtField7.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, "Naziv", Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Tip", Trim$(cmbField5.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "JedinicaMere", Trim$(cmbField6.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "CenaPoJedinici", cenaEdit, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "DozaPoHa", dozaEdit, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Kultura", Trim$(cmbField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Pakovanje", pakovanjeEdit, SRC

            tx.CommitTx

        Case "Kulture"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite vrstu vo" & ChrW(263) & "a!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.value) = "" Then
                MsgBox "Unesite sortu vo" & ChrW(263) & "a!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            Dim gajbicaEdit As Double
            If Trim$(txtField3.value) <> "" Then
                If Not TryParseDouble(txtField3.value, gajbicaEdit) Or gajbicaEdit < 0 Then
                    MsgBox "Unesite validan broj gajbica po paleti!", vbExclamation, APP_NAME
                    txtField3.SetFocus
                    Exit Sub
                End If
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, "VrstaVoca", Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "SortaVoca", Trim$(txtField2.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KUL_GAJBICA_PALETA, gajbicaEdit, SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KUL_TIP_AMBALAZE, Trim$(cmbField1.value), SRC

            tx.CommitTx

        Case "TipAmbalaze"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip ambala" & ChrW(382) & "e!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaAmbEdit As Double
            If Not TryParseDouble(txtField2.value, tezinaAmbEdit) Or tezinaAmbEdit < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu gajbice (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, COL_TAMB_TIP, Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_TAMB_TEZINA, tezinaAmbEdit, SRC

            tx.CommitTx

        Case "TipPalete"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip palete!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaPalEdit As Double
            If Not TryParseDouble(txtField2.value, tezinaPalEdit) Or tezinaPalEdit < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu palete (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, COL_TPAL_TIP, Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_TPAL_TEZINA, tezinaPalEdit, SRC

            tx.CommitTx

        Case "Kutije"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip kutije!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaKutEdit As Double
            If Not TryParseDouble(txtField2.value, tezinaKutEdit) Or tezinaKutEdit < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu kutije (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, COL_KUT_TIP, Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KUT_TEZINA, tezinaKutEdit, SRC

            tx.CommitTx

        Case "Kese"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip kese!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            Dim tezinaKesEdit As Double
            If Not TryParseDouble(txtField2.value, tezinaKesEdit) Or tezinaKesEdit < 0 Then
                MsgBox "Unesite validnu te" & ChrW(382) & "inu kese (kg)!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, COL_KES_TIP, Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, COL_KES_TEZINA, tezinaKesEdit, SRC

            tx.CommitTx

        Case "VrstaGP"
            If Trim$(txtField1.value) = "" Then
                MsgBox "Unesite tip gotovog proizvoda!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            tx.BeginTx
            tx.AddTableSnapshot m_TableName

            RequireUpdateCell m_TableName, m_SelectedRow, COL_VGP_TIP, Trim$(txtField1.value), SRC

            tx.CommitTx

        Case "Cenovnik"
            ' Append-only: istorija cena se ne menja. Nova cena = Dodaj.
            MsgBox "Cenovnik je append-only." & vbCrLf & _
                   Poruka("STM_MSG_NOVU_CENU_KORISTITE"), _
                   vbInformation, APP_NAME
            Exit Sub

        Case Else
            MsgBox "Nepoznat tip mati" & ChrW(269) & "nih podataka: " & Me.Tag, vbCritical, APP_NAME
            Exit Sub

    End Select

    Set tx = Nothing

    MsgBox "Izmenjeno!", vbInformation, APP_NAME

    LoadList
    ClearFields
    m_SelectedRow = 0

    Exit Sub

EH:
    LogErr SRC

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

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

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next

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
    ThisWorkbook.FollowHyperlink "https://a3.geosrbija.rs/"

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

    If Not TryExtractTwoCoordinates(txt, nVal, eVal) Then
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
    On Error Resume Next
    Me.MousePointer = fmMousePointerDefault
    On Error GoTo 0

    LogErr "frmStammdaten.btnOpenPolygonEditor_Click"
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

Private Sub LoadStaniceIntoCombo()
    On Error GoTo EH

    Dim data As Variant
    Dim i As Long

    cmbField1.Clear
    cmbField1.ColumnCount = 2
    cmbField1.ColumnWidths = "150 pt;0 pt"

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
            cmbField1.AddItem NzToText(data(i, colNaziv))
            cmbField1.List(cmbField1.ListCount - 1, 1) = NzToText(data(i, colID))
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

Private Sub ClearGeoFields()
    On Error Resume Next

    txtNCoord.value = ""
    txtECoord.value = ""

    On Error GoTo 0
End Sub

Private Function TryExtractTwoCoordinates(ByVal rawText As String, _
                                          ByRef firstCoord As Double, _
                                          ByRef secondCoord As Double) As Boolean
    On Error GoTo EH

    Dim txt As String
    txt = Trim$(rawText)

    If txt = "" Then Exit Function

    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, vbTab, " ")
    txt = Replace(txt, ";", " ")

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    Dim tokens() As String
    tokens = Split(txt, " ")

    Dim vals(0 To 1) As Double
    Dim count As Long
    Dim i As Long
    Dim d As Double
    Dim candidate As String

    For i = LBound(tokens) To UBound(tokens)
        candidate = CleanCoordToken(tokens(i))

        If TryParseDouble(candidate, d) Then
            If Abs(d) > 1000 Then
                vals(count) = d
                count = count + 1

                If count = 2 Then Exit For
            End If
        End If
    Next i

    If count < 2 Then Exit Function

    firstCoord = vals(0)
    secondCoord = vals(1)

    TryExtractTwoCoordinates = True
    Exit Function

EH:
    TryExtractTwoCoordinates = False
End Function

Private Function CleanCoordToken(ByVal token As String) As String
    Dim s As String
    s = Trim$(token)

    s = Replace(s, "N=", "")
    s = Replace(s, "E=", "")
    s = Replace(s, "N:", "")
    s = Replace(s, "E:", "")
    s = Replace(s, "n=", "")
    s = Replace(s, "e=", "")
    s = Replace(s, "n:", "")
    s = Replace(s, "e:", "")

    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")

    CleanCoordToken = s
End Function

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

    Dim url As String

    url = "https://www.google.com/maps?q=" & _
          Replace(CStr(lat), ",", ".") & "," & _
          Replace(CStr(lng), ",", ".")

    ThisWorkbook.FollowHyperlink url

    OpenGoogleMaps = True
    Exit Function

EH:
    LogErr "frmStammdaten.OpenGoogleMaps"
    OpenGoogleMaps = False
End Function

Public Function OpenParcelPolygonEditor(ByVal parcelaID As String) As Boolean
    On Error GoTo EH

    Dim url As String

    If Trim$(parcelaID) = "" Then
        LogError "frmStammdaten.OpenParcelPolygonEditor", "ParcelaID nije prosleden."
        Exit Function
    End If

    url = "https://dusanmiladinovicvnm.github.io/otkupapp-pwa/parcel-draw.html?parcelaId=" & _
          WorksheetFunction.EncodeURL(parcelaID)

    ThisWorkbook.FollowHyperlink url

    OpenParcelPolygonEditor = True
    Exit Function

EH:
    LogErr "frmStammdaten.OpenParcelPolygonEditor"
    OpenParcelPolygonEditor = False
End Function

