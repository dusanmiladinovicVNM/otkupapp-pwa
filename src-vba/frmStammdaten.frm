VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStammdaten 
   Caption         =   "UserForm1"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19515
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
' frmStammdaten – Universelles Stammdaten-Form
' Wird über .Tag gesteuert: "Kooperanti", "Stanice", "Kupci", "Vozaci"
' ============================================================

Private m_TableName As String
Private m_Headers As Variant
Private m_FieldCount As Long
Private m_SelectedRow As Long
Private m_SetupDone As Boolean

Private m_RowMap() As Long
Private m_RowMapCount As Long

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

Private Sub UserForm_Initialize()
    
    ' Nichts hier – Tag ist noch nicht verfügbar
End Sub

Private Sub UserForm_Activate()
    On Error GoTo EH

    If Not mChromeRemoved Then
        RemoveTitleBar
        Me.caption = ""
        mChromeRemoved = True
    End If

    ApplyTheme Me, BG_MAIN

    If m_SetupDone Then Exit Sub
    m_SetupDone = True

    Select Case Me.Tag
        Case "Kooperanti": SetupKooperanti
        Case "Stanice": SetupStanice
        Case "Kupci": SetupKupci
        Case "Vozaci": SetupVozaci
        Case "Parcele": SetupParcele
        Case "Artikli": SetupArtikli
        Case Else: SetupKooperanti
    End Select

    LoadList
    ClearFields
    m_SelectedRow = 0
    Exit Sub

EH:
    LogErr "frmStammdaten.UserForm_Activate"
    MsgBox "Greška pri otvaranju maticnih podataka: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' SETUP – Konfiguriert das Form je nach Entität
' ============================================================

Private Sub SetupKooperanti()
    ResetFieldVisibility

    Me.caption = "Kooperanti"
    lblTitle.caption = "Kooperanti"
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
    lblField7.caption = "Tekuci Racun": lblField7.Visible = True: txtField7.Visible = True
    lblField8.caption = "Pin": lblField8.Visible = True: txtField8.Visible = True
    lblField9.caption = "Adresa": lblField9.Visible = True: txtField9.Visible = True
    lblField10.caption = "JMBG": lblField10.Visible = True: txtField10.Visible = True

    LoadStaniceIntoCombo
End Sub

Private Sub SetupStanice()
    ResetFieldVisibility

    Me.caption = "Otkupna Mesta"
    lblTitle.caption = "Otkupna Mesta"
    m_TableName = TBL_STANICE

    m_Headers = Array( _
        "StanicaID", _
        "Naziv", _
        "Mesto", _
        "Telefon", _
        "Aktivan", _
        "KontaktIme", _
        "KontaktPrezime", _
        "Pin" _
    )

    m_FieldCount = 7

    lblField1.caption = "Naziv": lblField1.Visible = True: txtField1.Visible = True
    lblField2.caption = "Mesto": lblField2.Visible = True: txtField2.Visible = True
    lblField3.caption = "Telefon": lblField3.Visible = True: txtField3.Visible = True
    lblField4.caption = "Kontakt Ime": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "Kontakt Prezime": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "Pin": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "": lblField7.Visible = False: txtField7.Visible = False
    lblField8.caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.caption = "": lblField10.Visible = False: txtField10.Visible = False
End Sub

Private Sub SetupKupci()
    ResetFieldVisibility

    Me.caption = "Kupci"
    lblTitle.caption = "Kupci"
    m_TableName = TBL_KUPCI

    m_Headers = Array( _
        "KupacID", _
        "Naziv", _
        "Adresa", _
        "Drzava", _
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
    lblField4.caption = "Postanski Broj": lblField4.Visible = True: txtField4.Visible = True
    lblField5.caption = "Drzava": lblField5.Visible = True: txtField5.Visible = True
    lblField6.caption = "PIB": lblField6.Visible = True: txtField6.Visible = True
    lblField7.caption = "Maticni Broj": lblField7.Visible = True: txtField7.Visible = True
    lblField8.caption = "Email": lblField8.Visible = True: txtField8.Visible = True
    lblField9.caption = "Hladnjaca": lblField9.Visible = True: txtField9.Visible = True
    lblField10.caption = "Tekuci Racun": lblField10.Visible = True: txtField10.Visible = True
End Sub

Private Sub SetupVozaci()
    ResetFieldVisibility

    Me.caption = "Vozaci"
    lblTitle.caption = "Vozaci"
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
    lblTitle.caption = "Katastarske Parcele"
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
    lblField3.caption = "Kat. Opstina": lblField3.Visible = True: txtField3.Visible = True
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
End Sub
Private Sub SetupArtikli()
    ResetFieldVisibility
    Me.caption = "Artikli"
    lblTitle.caption = "Artikli Agrohemija"
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
            kupDrzava = GetColumnIndex(TBL_KUPCI, "Drzava")
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

            Dim pID As Long, pKoop As Long, pKat As Long, pOpstina As Long
            Dim pKultura As Long, pPov As Long, pGGAP As Long
            Dim pGeoStatus As Long, pGeoSource As Long, pRizik As Long, pNapomena As Long

            pID = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
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

            If pID = 0 Or pKoop = 0 Or pKat = 0 Or pOpstina = 0 Or _
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
                If Trim$(NzToText(data(i, pID))) <> "" Then
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

                    lstData.AddItem NzToText(data(i, pID))
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
    MsgBox "Greška pri ucitavanju liste: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' AUSWAHL IN LISTE ? Felder füllen
' ============================================================

Private Sub lstData_Click()
    If lstData.ListIndex < 0 Then Exit Sub
    
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
            txtField2.value = lstData.List(lstData.ListIndex, 2) ' Mesto
            txtField3.value = lstData.List(lstData.ListIndex, 3) ' Telefon
            txtField4.value = lstData.List(lstData.ListIndex, 5) ' KontaktIme
            txtField5.value = lstData.List(lstData.ListIndex, 6) ' KontaktPrezime
            txtField6.value = lstData.List(lstData.ListIndex, 7) ' Pin

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
    End Select
End Sub

' ============================================================
' HINZUFÜGEN
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
                MsgBox "Nije pronaden StanicaID za izabranu stanicu.", vbExclamation, APP_NAME
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
                Trim$(txtField6.value) _
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
                MsgBox "Unesite katastarsku opštinu!", vbExclamation, APP_NAME
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
                MsgBox "Unesite validnu površinu!", vbExclamation, APP_NAME
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

        Case Else
            MsgBox "Nepoznat tip maticnih podataka: " & Me.Tag, vbCritical, APP_NAME
            Exit Sub

    End Select

    If AppendRow(m_TableName, rowData) > 0 Then
        MsgBox "Dodato: " & newID, vbInformation, APP_NAME
        LoadList
        ClearFields
    Else
        MsgBox "Greška pri dodavanju!", vbCritical, APP_NAME
    End If

    Exit Sub

EH:
    LogErr "frmStammdaten.btnDodaj_Click"
    MsgBox "Greška pri dodavanju: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' ÄNDERN
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
                MsgBox "Nije pronaden StanicaID za izabranu stanicu.", vbExclamation, APP_NAME
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

            RequireUpdateCell m_TableName, m_SelectedRow, "Naziv", Trim$(txtField1.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Mesto", Trim$(txtField2.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Telefon", Trim$(txtField3.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "KontaktIme", Trim$(txtField4.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "KontaktPrezime", Trim$(txtField5.value), SRC
            RequireUpdateCell m_TableName, m_SelectedRow, "Pin", Trim$(txtField6.value), SRC

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
            RequireUpdateCell m_TableName, m_SelectedRow, "Drzava", Trim$(txtField5.value), SRC
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
                MsgBox "Unesite katastarsku opštinu!", vbExclamation, APP_NAME
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
                MsgBox "Unesite validnu površinu!", vbExclamation, APP_NAME
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

        Case Else
            MsgBox "Nepoznat tip maticnih podataka: " & Me.Tag, vbCritical, APP_NAME
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

    MsgBox "Greška pri izmeni: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' NAVIGATION & HELPER
' ============================================================

Private Sub btnPovratak_Click()
    Unload Me
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        frmOtkupAPP.Show
    End If
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
Private Sub btnGeoOpen_Click()

    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If lstData.ListIndex < 0 Then
        MsgBox "Izaberite parcelu iz liste!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim katBroj As String
    Dim katOpstina As String
    Dim searchText As String

    katBroj = lstData.List(lstData.ListIndex, 2)
    katOpstina = lstData.List(lstData.ListIndex, 3)

    If Len(katBroj) = 0 Or Len(katOpstina) = 0 Then
        MsgBox "Parcela nema katastarski broj ili katastarsku opštinu.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    searchText = katBroj & " " & Replace(katOpstina, "KO ", "")

    ' ?? Kopiraj u clipboard
    CopyToClipboard searchText

    ' ?? Otvori GeoSrbiju
    ThisWorkbook.FollowHyperlink "https://a3.geosrbija.rs/"

    'MsgBox "GeoSrbija otvorena." & vbCrLf & vbCrLf & _
           "Pretraga je kopirana:" & vbCrLf & _
           searchText & vbCrLf & vbCrLf & _
           "Samo CTRL + V i Enter", _
           vbInformation, APP_NAME

    Exit Sub

EH:
    LogErr "frmStammdaten.btnGeoOpen_Click"
    MsgBox "Greška pri otvaranju GeoSrbije: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnGeoSave_Click()
    On Error GoTo EH
    
    If Me.Tag <> "Parcele" Then Exit Sub
    
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim nVal As Double
    Dim eVal As Double
    
    If Not TryParseDouble(txtNCoord.value, nVal) Then
        MsgBox "Unesite validnu N koordinatu.", vbExclamation, APP_NAME
        txtNCoord.SetFocus
        Exit Sub
    End If

    If Not TryParseDouble(txtECoord.value, eVal) Then
        MsgBox "Unesite validnu E koordinatu.", vbExclamation, APP_NAME
        txtECoord.SetFocus
        Exit Sub
    End If

    If nVal <= 0 Or eVal <= 0 Then
        MsgBox "Koordinate moraju biti pozitivne vrednosti.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    SaveParcelGeoPoint m_SelectedRow, nVal, eVal
    
    MsgBox "Geo podaci su sacuvani.", vbInformation, APP_NAME
    
    LoadList
    ClearGeoFields
    Exit Sub
    
EH:
    LogErr "frmStammdaten.btnGeoSave_Click"
    MsgBox "Greška pri cuvanju geo podataka: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnGeoClear_Click()
    On Error GoTo EH
    
    If Me.Tag <> "Parcele" Then Exit Sub
    
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu iz liste!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If MsgBox("Da li sigurno želite da obrišete geo podatke za izabranu parcelu?", _
              vbQuestion + vbYesNo, APP_NAME) <> vbYes Then
        Exit Sub
    End If
    
    ClearParcelGeo m_SelectedRow
    
    MsgBox "Geo podaci su obrisani.", vbInformation, APP_NAME

    LoadList
    ClearGeoFields

    Exit Sub

EH:
    LogErr "frmStammdaten.btnGeoClear_Click"
    MsgBox "Greška pri brisanju geo podataka: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnPasteCoords_Click()
    On Error GoTo EH

    Dim txt As String
    txt = Trim$(GetClipboardText())

    If txt = "" Then
        MsgBox "Clipboard je prazan.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim nVal As Double
    Dim eVal As Double

    If Not TryExtractTwoCoordinates(txt, nVal, eVal) Then
        MsgBox "Nisu pronadene validne koordinate u clipboard-u.", vbExclamation, APP_NAME
        Exit Sub
    End If

    txtNCoord.value = FormatCoordForTextBox(nVal)
    txtECoord.value = FormatCoordForTextBox(eVal)

    MsgBox "Koordinate su ucitane.", vbInformation, APP_NAME

    Exit Sub

EH:
    LogErr "frmStammdaten.btnPasteCoords_Click"
    MsgBox "Greška pri ucitavanju koordinata: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnOpenMap_Click()
    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim data As Variant
    data = GetTableData(TBL_PARCELE)

    If IsEmpty(data) Then
        MsgBox "Tabela parcela je prazna.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim latIdx As Long
    Dim lngIdx As Long

    latIdx = GetColumnIndex(TBL_PARCELE, COL_PAR_LAT)
    lngIdx = GetColumnIndex(TBL_PARCELE, COL_PAR_LNG)

    If latIdx = 0 Or lngIdx = 0 Then
        MsgBox "Lat/Lng kolone nisu pronadene.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim lat As Double
    Dim lng As Double

    If Not TryParseDouble(NzToText(data(m_SelectedRow, latIdx)), lat) Or _
       Not TryParseDouble(NzToText(data(m_SelectedRow, lngIdx)), lng) Then
        MsgBox "Parcela nema validne Lat/Lng geo podatke.", vbExclamation, APP_NAME
        Exit Sub
    End If

    OpenGoogleMaps lat, lng

    Exit Sub

EH:
    LogErr "frmStammdaten.btnOpenMap_Click"
    MsgBox "Greška pri otvaranju Google Maps: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnOpenPolygonEditor_Click()
    On Error GoTo EH

    If Me.Tag <> "Parcele" Then Exit Sub

    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim data As Variant
    data = GetTableData(TBL_PARCELE)

    If IsEmpty(data) Then
        MsgBox "Tabela parcela je prazna.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim idIdx As Long
    idIdx = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)

    If idIdx = 0 Then
        MsgBox "ParcelaID kolona nije pronadena.", vbCritical, APP_NAME
        Exit Sub
    End If

    Dim parcelaID As String
    parcelaID = NzToText(data(m_SelectedRow, idIdx))

    If parcelaID = "" Then
        MsgBox "Izabrana parcela nema ParcelaID.", vbExclamation, APP_NAME
        Exit Sub
    End If

    OpenParcelPolygonEditor parcelaID

    Exit Sub

EH:
    LogErr "frmStammdaten.btnOpenPolygonEditor_Click"
    MsgBox "Greška pri otvaranju polygon editora: " & Err.Description, vbCritical, APP_NAME
End Sub

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
    MsgBox "Greška pri ucitavanju stanica: " & Err.Description, vbCritical, APP_NAME
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

Public Sub OpenGoogleMaps(ByVal lat As Double, ByVal lng As Double)
    On Error GoTo EH

    Dim url As String

    url = "https://www.google.com/maps?q=" & _
          Replace(CStr(lat), ",", ".") & "," & _
          Replace(CStr(lng), ",", ".")

    ThisWorkbook.FollowHyperlink url
    Exit Sub

EH:
    LogErr "frmStammdaten.OpenGoogleMaps"
    MsgBox "Greška pri otvaranju Google Maps: " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub OpenParcelPolygonEditor(ByVal parcelaID As String)
    On Error GoTo EH

    Dim url As String

    If Trim$(parcelaID) = "" Then
        MsgBox "ParcelaID nije prosleden.", vbExclamation, APP_NAME
        Exit Sub
    End If

    url = "https://dusanmiladinovicvnm.github.io/otkupapp-pwa/parcel-draw.html?parcelaId=" & _
          WorksheetFunction.EncodeURL(parcelaID)

    ThisWorkbook.FollowHyperlink url
    Exit Sub

EH:
    LogErr "frmStammdaten.OpenParcelPolygonEditor"
    MsgBox "Greška pri otvaranju polygon editora: " & Err.Description, vbCritical, APP_NAME
End Sub

