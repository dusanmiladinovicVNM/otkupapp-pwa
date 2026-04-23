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

Private Sub UserForm_Initialize()
    
    ' Nichts hier – Tag ist noch nicht verfügbar
End Sub

Private Sub UserForm_Activate()
    If Not mChromeRemoved Then
        Me.Caption = ""
        RemoveTitleBar
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
End Sub

' ============================================================
' SETUP – Konfiguriert das Form je nach Entität
' ============================================================

Private Sub SetupKooperanti()
    cmbField1.Visible = True
    cmbField2.Visible = False
    cmbField3.Visible = False
    cmbField4.Visible = False
    cmbField5.Visible = False
    cmbField6.Visible = False

    Me.Caption = "Kooperanti"
    lblTitle.Caption = "Kooperanti"
    m_TableName = TBL_KOOPERANTI

    ' display headers for ListBox
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

    lblField1.Caption = "Ime": lblField1.Visible = True: txtField1.Visible = True
    lblField2.Caption = "Prezime": lblField2.Visible = True: txtField2.Visible = True
    lblField3.Caption = "Mesto": lblField3.Visible = True: txtField3.Visible = True
    lblField4.Caption = "Telefon": lblField4.Visible = True: txtField4.Visible = True

    lblField5.Caption = "Stanica": lblField5.Visible = True: cmbField1.Visible = True
    lblField6.Caption = "BPG Broj": lblField6.Visible = True: txtField6.Visible = True
    lblField7.Caption = "Tekuci Racun": lblField7.Visible = True: txtField7.Visible = True
    lblField8.Caption = "Pin": lblField8.Visible = True: txtField8.Visible = True
    lblField9.Caption = "Adresa": lblField9.Visible = True: txtField9.Visible = True
    lblField10.Caption = "JMBG": lblField10.Visible = True: txtField10.Visible = True

    LoadStaniceIntoCombo
End Sub

Private Sub SetupStanice()
    cmbField1.Visible = False
    cmbField2.Visible = False
    cmbField3.Visible = False
    cmbField4.Visible = False
    cmbField5.Visible = False
    cmbField6.Visible = False

    Me.Caption = "Otkupna Mesta"
    lblTitle.Caption = "Otkupna Mesta"
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

    lblField1.Caption = "Naziv": lblField1.Visible = True: txtField1.Visible = True
    lblField2.Caption = "Mesto": lblField2.Visible = True: txtField2.Visible = True
    lblField3.Caption = "Telefon": lblField3.Visible = True: txtField3.Visible = True
    lblField4.Caption = "Kontakt Ime": lblField4.Visible = True: txtField4.Visible = True
    lblField5.Caption = "Kontakt Prezime": lblField5.Visible = True: txtField5.Visible = True
    lblField6.Caption = "Pin": lblField6.Visible = True: txtField6.Visible = True
    lblField7.Caption = "": lblField7.Visible = False: txtField7.Visible = False
    lblField8.Caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.Caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.Caption = "": lblField10.Visible = False: txtField10.Visible = False
End Sub

Private Sub SetupKupci()
    cmbField1.Visible = False
    cmbField2.Visible = False
    cmbField3.Visible = False
    cmbField4.Visible = False
    cmbField5.Visible = False
    cmbField6.Visible = False

    Me.Caption = "Kupci"
    lblTitle.Caption = "Kupci"
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

    lblField1.Caption = "Naziv": lblField1.Visible = True: txtField1.Visible = True
    lblField2.Caption = "Ulica": lblField2.Visible = True: txtField2.Visible = True
    lblField3.Caption = "Mesto": lblField3.Visible = True: txtField3.Visible = True
    lblField4.Caption = "Postanski Broj": lblField4.Visible = True: txtField4.Visible = True
    lblField5.Caption = "Drzava": lblField5.Visible = True: txtField5.Visible = True
    lblField6.Caption = "PIB": lblField6.Visible = True: txtField6.Visible = True
    lblField7.Caption = "Maticni Broj": lblField7.Visible = True: txtField7.Visible = True
    lblField8.Caption = "Email": lblField8.Visible = True: txtField8.Visible = True
    lblField9.Caption = "Hladnjaca": lblField9.Visible = True: txtField9.Visible = True
    lblField10.Caption = "Tekuci Racun": lblField10.Visible = True: txtField10.Visible = True
End Sub

Private Sub SetupVozaci()
    cmbField1.Visible = False
    cmbField2.Visible = False
    cmbField3.Visible = False
    cmbField4.Visible = False
    cmbField5.Visible = False
    cmbField6.Visible = False

    Me.Caption = "Vozaci"
    lblTitle.Caption = "Vozaci"
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

    lblField1.Caption = "Ime": lblField1.Visible = True: txtField1.Visible = True
    lblField2.Caption = "Prezime": lblField2.Visible = True: txtField2.Visible = True
    lblField3.Caption = "Telefon": lblField3.Visible = True: txtField3.Visible = True
    lblField4.Caption = "PIN": lblField4.Visible = True: txtField4.Visible = True
    lblField5.Caption = "": lblField5.Visible = False: txtField5.Visible = False
    lblField6.Caption = "": lblField6.Visible = False: txtField6.Visible = False
    lblField7.Caption = "": lblField7.Visible = False: txtField7.Visible = False
    lblField8.Caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.Caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.Caption = "": lblField10.Visible = False: txtField10.Visible = False
End Sub

Private Sub SetupParcele()
    Me.Caption = "Parcele"
    lblTitle.Caption = "Katastarske Parcele"
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

    lblField1.Caption = "Kooperant": lblField1.Visible = True: txtField1.Visible = False
    lblField2.Caption = "Kat. Broj": lblField2.Visible = True: txtField2.Visible = True
    lblField3.Caption = "Kat. Opstina": lblField3.Visible = True: txtField3.Visible = True
    lblField4.Caption = "Kultura": lblField4.Visible = True: txtField4.Visible = False
    lblField5.Caption = "Povrsina (ha)": lblField5.Visible = True: txtField5.Visible = True
    lblField6.Caption = "GGAP Status": lblField6.Visible = True: txtField6.Visible = False
    lblField7.Caption = "Napomena": lblField7.Visible = True: txtField7.Visible = True
    lblField8.Caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.Caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.Caption = "": lblField10.Visible = False: txtField10.Visible = False

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
    Me.Caption = "Artikli"
    lblTitle.Caption = "Artikli Agrohemija"
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
    lblField1.Caption = "Naziv": lblField1.Visible = True: txtField1.Visible = True
    lblField2.Caption = "Tip": lblField2.Visible = True: txtField2.Visible = False
    lblField3.Caption = "Jedinica Mere": lblField3.Visible = True: txtField3.Visible = False
    lblField4.Caption = "Cena po jed.": lblField4.Visible = True: txtField4.Visible = True
    lblField5.Caption = "Kultura": lblField5.Visible = True: txtField5.Visible = False
    lblField6.Caption = "Doza po ha": lblField6.Visible = True: txtField6.Visible = True
    lblField7.Caption = "Pakovanje": lblField7.Visible = True: txtField7.Visible = True
    lblField8.Caption = "": lblField8.Visible = False: txtField8.Visible = False
    lblField9.Caption = "": lblField9.Visible = False: txtField9.Visible = False
    lblField10.Caption = "": lblField10.Visible = False: txtField10.Visible = False

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
    lstData.RowSource = ""
    lstData.Clear

    Dim data As Variant
    data = GetTableData(m_TableName)
    If IsEmpty(data) Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim maxCols As Long

    Dim stanicaNaziv As String
    Dim punoIme As String
    Dim punaAdresa As String
    Dim kupacAdresa As String

    Select Case Me.Tag

        Case "Kooperanti"
            lstData.ColumnCount = 10

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, 1))) <> "" Then
                    ' 1 KooperantID
                    ' 2 Ime
                    ' 3 Prezime
                    ' 4 Mesto
                    ' 5 Telefon
                    ' 6 StanicaID
                    ' 7 Aktivan
                    ' 8 BPGBroj
                    ' 9 TekuciRacun
                    ' 10 Pin
                    ' 11 Adresa
                    ' 12 JMBG

                    punoIme = Trim$(NzToText(data(i, 2)) & " " & NzToText(data(i, 3)))

                    punaAdresa = Trim$(NzToText(data(i, 11)))
                    If Len(Trim$(NzToText(data(i, 4)))) > 0 Then
                        If Len(punaAdresa) > 0 Then punaAdresa = punaAdresa & ", "
                        punaAdresa = punaAdresa & NzToText(data(i, 4))
                    End If

                    stanicaNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", NzToText(data(i, 6)), "Naziv"))

                    lstData.AddItem NzToText(data(i, 1))                           ' 0 KooperantID
                    lstData.List(lstData.ListCount - 1, 1) = punoIme               ' 1 Ime + Prezime
                    lstData.List(lstData.ListCount - 1, 2) = NzToText(data(i, 5))  ' 2 Telefon
                    lstData.List(lstData.ListCount - 1, 3) = stanicaNaziv          ' 3 Stanica
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, 8))  ' 4 BPGBroj
                    lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, 9))  ' 5 TekuciRacun
                    lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, 10)) ' 6 Pin
                    lstData.List(lstData.ListCount - 1, 7) = punaAdresa            ' 7 Adresa + Mesto
                    lstData.List(lstData.ListCount - 1, 8) = NzToText(data(i, 12)) ' 8 JMBG
                    lstData.List(lstData.ListCount - 1, 9) = NzToText(data(i, 7))  ' 9 Aktivan
                End If
            Next i

        Case "Kupci"
            lstData.ColumnCount = 10

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, 1))) <> "" Then
                    ' 1 KupacID
                    ' 2 Naziv
                    ' 3 Ulica
                    ' 4 Mesto
                    ' 5 PostanskiBroj
                    ' 6 Drzava
                    ' 7 PIB
                    ' 8 MaticniBroj
                    ' 9 Email
                    ' 10 Hladnjaca
                    ' 11 Aktivan
                    ' 12 TekuciRacun

                    kupacAdresa = Trim$(NzToText(data(i, 3))) ' Ulica

                    If Len(Trim$(NzToText(data(i, 5)))) > 0 Then
                        If Len(kupacAdresa) > 0 Then kupacAdresa = kupacAdresa & ", "
                        kupacAdresa = kupacAdresa & NzToText(data(i, 5)) ' PostanskiBroj
                    End If

                    If Len(Trim$(NzToText(data(i, 4)))) > 0 Then
                        If Len(kupacAdresa) > 0 Then kupacAdresa = kupacAdresa & " "
                        kupacAdresa = kupacAdresa & NzToText(data(i, 4)) ' Mesto
                    End If

                    lstData.AddItem NzToText(data(i, 1))                           ' 0 KupacID
                    lstData.List(lstData.ListCount - 1, 1) = NzToText(data(i, 2))  ' 1 Naziv
                    lstData.List(lstData.ListCount - 1, 2) = kupacAdresa           ' 2 Adresa
                    lstData.List(lstData.ListCount - 1, 3) = NzToText(data(i, 6))  ' 3 Drzava
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, 7))  ' 4 PIB
                    lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, 8))  ' 5 MaticniBroj
                    lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, 9))  ' 6 Email
                    lstData.List(lstData.ListCount - 1, 7) = NzToText(data(i, 10)) ' 7 Hladnjaca
                    lstData.List(lstData.ListCount - 1, 8) = NzToText(data(i, 11)) ' 8 Aktivan
                    lstData.List(lstData.ListCount - 1, 9) = NzToText(data(i, 12)) ' 9 TekuciRacun
                End If
            Next i

        Case "Parcele"
            lstData.ColumnCount = 10

            Dim koopID As String
            Dim koopNaziv As String
            Dim geoInfo As String
            Dim rizikInfo As String

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, 1))) <> "" Then
                    ' 1  ParcelaID
                    ' 2  KooperantID
                    ' 3  KatBroj
                    ' 4  KatOpstina
                    ' 5  Kultura
                    ' 6  PovrsinaHa
                    ' 7  GGAPStatus
                    ' 8  Aktivna
                    ' 9  GeoStatus
                    ' 10 GeoSource
                    ' 11 N_Coord
                    ' 12 E_Coord
                    ' 13 Lat
                    ' 14 Lng
                    ' 15 PolygonGeoJSON
                    ' 16 MeteoEnabled
                    ' 17 RizikStatus
                    ' 18 DatumGeoUnosa
                    ' 19 DatumAzuriranja
                    ' 20 Napomena

                    koopID = NzToText(data(i, 2))
                    koopNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime")) & " " & _
                                CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime")) & _
                                " (" & koopID & ")"

                    geoInfo = NzToText(data(i, 9))
                    If Len(NzToText(data(i, 10))) > 0 Then
                        If Len(geoInfo) > 0 Then geoInfo = geoInfo & " / "
                        geoInfo = geoInfo & NzToText(data(i, 10))
                    End If

                    rizikInfo = NzToText(data(i, 17))

                    lstData.AddItem NzToText(data(i, 1))                           ' 0 ParcelaID
                    lstData.List(lstData.ListCount - 1, 1) = koopNaziv             ' 1 Kooperant
                    lstData.List(lstData.ListCount - 1, 2) = NzToText(data(i, 3))  ' 2 KatBroj
                    lstData.List(lstData.ListCount - 1, 3) = NzToText(data(i, 4))  ' 3 KatOpstina
                    lstData.List(lstData.ListCount - 1, 4) = NzToText(data(i, 5))  ' 4 Kultura
                    lstData.List(lstData.ListCount - 1, 5) = NzToText(data(i, 6))  ' 5 PovrsinaHa
                    lstData.List(lstData.ListCount - 1, 6) = NzToText(data(i, 7))  ' 6 GGAPStatus
                    lstData.List(lstData.ListCount - 1, 7) = geoInfo               ' 7 Geo
                    lstData.List(lstData.ListCount - 1, 8) = rizikInfo             ' 8 Rizik
                    lstData.List(lstData.ListCount - 1, 9) = NzToText(data(i, 20)) ' 9 Napomena
                End If
            Next i
            
        Case Else
            lstData.ColumnCount = UBound(m_Headers) - LBound(m_Headers) + 1
            maxCols = lstData.ColumnCount

            For i = 1 To UBound(data, 1)
                If Trim$(NzToText(data(i, 1))) <> "" Then
                    lstData.AddItem NzToText(data(i, 1))

                    For j = 2 To WorksheetFunction.Min(UBound(data, 2), maxCols)
                        lstData.List(lstData.ListCount - 1, j - 1) = NzToText(data(i, j))
                    Next j
                End If
            Next i
    End Select
End Sub

Private Function NzToText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzToText = ""
    Else
        NzToText = Trim$(CStr(v))
    End If
End Function

' ============================================================
' AUSWAHL IN LISTE ? Felder füllen
' ============================================================

Private Sub lstData_Click()
    If lstData.ListIndex < 0 Then Exit Sub
    m_SelectedRow = lstData.ListIndex + 1

    Dim data As Variant
    Dim koopNaziv As String
    Dim kID As String

    Select Case Me.Tag

        Case "Kooperanti"
            data = GetTableData(m_TableName)
            If IsEmpty(data) Then Exit Sub

            txtField1.Value = NzToText(data(m_SelectedRow, 2))   ' Ime
            txtField2.Value = NzToText(data(m_SelectedRow, 3))   ' Prezime
            txtField3.Value = NzToText(data(m_SelectedRow, 4))   ' Mesto
            txtField4.Value = NzToText(data(m_SelectedRow, 5))   ' Telefon

            SafeSetCombo cmbField1, CStr(LookupValue(TBL_STANICE, "StanicaID", NzToText(data(m_SelectedRow, 6)), "Naziv"))

            txtField6.Value = NzToText(data(m_SelectedRow, 8))   ' BPGBroj
            txtField7.Value = NzToText(data(m_SelectedRow, 9))   ' TekuciRacun
            txtField8.Value = NzToText(data(m_SelectedRow, 10))  ' Pin
            txtField9.Value = NzToText(data(m_SelectedRow, 11))  ' Adresa
            txtField10.Value = NzToText(data(m_SelectedRow, 12)) ' JMBG

        Case "Stanice"
            txtField1.Value = lstData.List(lstData.ListIndex, 1) ' Naziv
            txtField2.Value = lstData.List(lstData.ListIndex, 2) ' Mesto
            txtField3.Value = lstData.List(lstData.ListIndex, 3) ' Telefon
            txtField4.Value = lstData.List(lstData.ListIndex, 5) ' KontaktIme
            txtField5.Value = lstData.List(lstData.ListIndex, 6) ' KontaktPrezime
            txtField6.Value = lstData.List(lstData.ListIndex, 7) ' Pin

        Case "Kupci"
            data = GetTableData(m_TableName)
            If IsEmpty(data) Then Exit Sub

            txtField1.Value = NzToText(data(m_SelectedRow, 2))   ' Naziv
            txtField2.Value = NzToText(data(m_SelectedRow, 3))   ' Ulica
            txtField3.Value = NzToText(data(m_SelectedRow, 4))   ' Mesto
            txtField4.Value = NzToText(data(m_SelectedRow, 5))   ' PostanskiBroj
            txtField5.Value = NzToText(data(m_SelectedRow, 6))   ' Drzava
            txtField6.Value = NzToText(data(m_SelectedRow, 7))   ' PIB
            txtField7.Value = NzToText(data(m_SelectedRow, 8))   ' MaticniBroj
            txtField8.Value = NzToText(data(m_SelectedRow, 9))   ' Email
            txtField9.Value = NzToText(data(m_SelectedRow, 10))  ' Hladnjaca
            txtField10.Value = NzToText(data(m_SelectedRow, 12)) ' TekuciRacun

        Case "Vozaci"
            txtField1.Value = lstData.List(lstData.ListIndex, 1) ' Ime
            txtField2.Value = lstData.List(lstData.ListIndex, 2) ' Prezime
            txtField3.Value = lstData.List(lstData.ListIndex, 3) ' Telefon
            txtField4.Value = lstData.List(lstData.ListIndex, 5) ' PIN

        Case "Parcele"
            data = GetTableData(m_TableName)
            If IsEmpty(data) Then Exit Sub

            Dim koopID As String
            koopID = NzToText(data(m_SelectedRow, 2))

            koopNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime")) & " " & _
                        CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime")) & _
                        " (" & koopID & ")"

            SafeSetCombo cmbField1, koopNaziv                    ' Kooperant
            txtField2.Value = NzToText(data(m_SelectedRow, 3))  ' KatBroj
            txtField3.Value = NzToText(data(m_SelectedRow, 4))  ' KatOpstina
            SafeSetCombo cmbField2, NzToText(data(m_SelectedRow, 5))  ' Kultura
            txtField5.Value = NzToText(data(m_SelectedRow, 6))  ' PovrsinaHa
            SafeSetCombo cmbField3, NzToText(data(m_SelectedRow, 7))  ' GGAPStatus
            txtField7.Value = NzToText(data(m_SelectedRow, 20)) ' Napomena

        Case "Artikli"
            txtField1.Value = lstData.List(lstData.ListIndex, 1)   ' Naziv
            SafeSetCombo cmbField5, lstData.List(lstData.ListIndex, 2)   ' Tip
            SafeSetCombo cmbField6, lstData.List(lstData.ListIndex, 3)   ' JedinicaMere
            txtField4.Value = lstData.List(lstData.ListIndex, 4)   ' CenaPoJedinici
            txtField6.Value = lstData.List(lstData.ListIndex, 5)   ' DozaPoHa
            SafeSetCombo cmbField1, lstData.List(lstData.ListIndex, 6)   ' Kultura
            txtField7.Value = lstData.List(lstData.ListIndex, 7)   ' Pakovanje
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
            If Trim$(txtField1.Value) = "" Then
                MsgBox "Unesite ime!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.Value) = "" Then
                MsgBox "Unesite prezime!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField1.Value) = "" Then
                MsgBox "Izaberite stanicu!", vbExclamation, APP_NAME
                cmbField1.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "KooperantID", "KOOP-")
            stanicaID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbField1.Value, "StanicaID"))

            rowData = Array( _
                newID, _
                Trim$(txtField1.Value), _
                Trim$(txtField2.Value), _
                Trim$(txtField3.Value), _
                Trim$(txtField4.Value), _
                stanicaID, _
                STATUS_AKTIVAN, _
                Trim$(txtField6.Value), _
                Trim$(txtField7.Value), _
                Trim$(txtField8.Value), _
                Trim$(txtField9.Value), _
                Trim$(txtField10.Value) _
            )
            
        Case "Stanice"
            If Trim$(txtField1.Value) = "" Then
                MsgBox "Unesite naziv!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.Value) = "" Then
                MsgBox "Unesite mesto!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "StanicaID", "ST-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.Value), _
                Trim$(txtField2.Value), _
                Trim$(txtField3.Value), _
                STATUS_AKTIVAN, _
                Trim$(txtField4.Value), _
                Trim$(txtField5.Value), _
                Trim$(txtField6.Value) _
            )
            
        Case "Kupci"
            If Trim$(txtField1.Value) = "" Then
                MsgBox "Unesite naziv kupca!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "KupacID", "KUP-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.Value), _
                Trim$(txtField2.Value), _
                Trim$(txtField3.Value), _
                Trim$(txtField4.Value), _
                Trim$(txtField5.Value), _
                Trim$(txtField6.Value), _
                Trim$(txtField7.Value), _
                Trim$(txtField8.Value), _
                Trim$(txtField9.Value), _
                STATUS_AKTIVAN, _
                Trim$(txtField10.Value) _
            )
            
        Case "Vozaci"
            If Trim$(txtField1.Value) = "" Then
                MsgBox "Unesite ime vozaca!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.Value) = "" Then
                MsgBox "Unesite prezime vozaca!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "VozacID", "VOZ-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.Value), _
                Trim$(txtField2.Value), _
                Trim$(txtField3.Value), _
                STATUS_AKTIVAN, _
                Trim$(txtField4.Value) _
            )
                           
        Case "Parcele"
            If Trim$(cmbField1.Value) = "" Then
                MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
                cmbField1.SetFocus
                Exit Sub
            End If

            If Trim$(txtField2.Value) = "" Then
                MsgBox "Unesite katastarski broj!", vbExclamation, APP_NAME
                txtField2.SetFocus
                Exit Sub
            End If

            If Trim$(txtField3.Value) = "" Then
                MsgBox "Unesite katastarsku opstinu!", vbExclamation, APP_NAME
                txtField3.SetFocus
                Exit Sub
            End If

            If Not IsNumeric(Replace(txtField5.Value, ",", ".")) Then
                MsgBox "Unesite validnu povrsinu!", vbExclamation, APP_NAME
                txtField5.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "ParcelaID", "PAR-")

            Dim koopID As String
            Dim povrsina As Double

            koopID = ExtractIDFromDisplay(cmbField1.Value)
            povrsina = CDbl(Replace(txtField5.Value, ",", "."))

            rowData = Array( _
                newID, _
                koopID, _
                Trim$(txtField2.Value), _
                Trim$(txtField3.Value), _
                Trim$(cmbField2.Value), _
                povrsina, _
                Trim$(cmbField3.Value), _
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
                Trim$(txtField7.Value) _
            )
            
        Case "Artikli"
            If Trim$(txtField1.Value) = "" Then
                MsgBox "Unesite naziv artikla!", vbExclamation, APP_NAME
                txtField1.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField5.Value) = "" Then
                MsgBox "Izaberite tip artikla!", vbExclamation, APP_NAME
                cmbField5.SetFocus
                Exit Sub
            End If

            If Trim$(cmbField6.Value) = "" Then
                MsgBox "Izaberite jedinicu mere!", vbExclamation, APP_NAME
                cmbField6.SetFocus
                Exit Sub
            End If

            Dim cenaArt As Double
            Dim dozaArt As Double
            Dim pakovanjeArt As Double

            If IsNumeric(Replace(txtField4.Value, ",", ".")) Then
                cenaArt = CDbl(Replace(txtField4.Value, ",", "."))
            Else
                MsgBox "Unesite validnu cenu!", vbExclamation, APP_NAME
                txtField4.SetFocus
                Exit Sub
            End If

            If IsNumeric(Replace(txtField6.Value, ",", ".")) Then
                dozaArt = CDbl(Replace(txtField6.Value, ",", "."))
            Else
                MsgBox "Unesite validnu dozu!", vbExclamation, APP_NAME
                txtField6.SetFocus
                Exit Sub
            End If

            If IsNumeric(Replace(txtField7.Value, ",", ".")) Then
                pakovanjeArt = CDbl(Replace(txtField7.Value, ",", "."))
            Else
                MsgBox "Unesite validno pakovanje!", vbExclamation, APP_NAME
                txtField7.SetFocus
                Exit Sub
            End If

            newID = GetNextID(m_TableName, "ArtikalID", "ART-")

            rowData = Array( _
                newID, _
                Trim$(txtField1.Value), _
                Trim$(cmbField5.Value), _
                Trim$(cmbField6.Value), _
                cenaArt, _
                dozaArt, _
                Trim$(cmbField1.Value), _
                pakovanjeArt _
            )
        End Select

    If AppendRow(m_TableName, rowData) > 0 Then
        MsgBox "Dodato: " & newID, vbInformation, APP_NAME
        LoadList
        ClearFields
    Else
        MsgBox "Greska!", vbCritical, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "frmStammdaten.btnDodaj"
    MsgBox "Greska pri dodavanju: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' ÄNDERN
' ============================================================

Private Sub btnIzmeni_Click()
    On Error GoTo EH
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite stavku iz liste!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Select Case Me.Tag
            
        Case "Kooperanti"
            UpdateCell m_TableName, m_SelectedRow, "Ime", Trim$(txtField1.Value)
            UpdateCell m_TableName, m_SelectedRow, "Prezime", Trim$(txtField2.Value)
            UpdateCell m_TableName, m_SelectedRow, "Mesto", Trim$(txtField3.Value)
            UpdateCell m_TableName, m_SelectedRow, "Telefon", Trim$(txtField4.Value)

            Dim stID As String
            stID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbField1.Value, "StanicaID"))
            UpdateCell m_TableName, m_SelectedRow, "StanicaID", stID

            UpdateCell m_TableName, m_SelectedRow, "BPGBroj", Trim$(txtField6.Value)
            UpdateCell m_TableName, m_SelectedRow, "TekuciRacun", Trim$(txtField7.Value)
            UpdateCell m_TableName, m_SelectedRow, "Pin", Trim$(txtField8.Value)
            UpdateCell m_TableName, m_SelectedRow, "Adresa", Trim$(txtField9.Value)
            UpdateCell m_TableName, m_SelectedRow, "JMBG", Trim$(txtField10.Value)
            
        Case "Stanice"
            UpdateCell m_TableName, m_SelectedRow, "Naziv", Trim$(txtField1.Value)
            UpdateCell m_TableName, m_SelectedRow, "Mesto", Trim$(txtField2.Value)
            UpdateCell m_TableName, m_SelectedRow, "Telefon", Trim$(txtField3.Value)
            UpdateCell m_TableName, m_SelectedRow, "KontaktIme", Trim$(txtField4.Value)
            UpdateCell m_TableName, m_SelectedRow, "KontaktPrezime", Trim$(txtField5.Value)
            UpdateCell m_TableName, m_SelectedRow, "Pin", Trim$(txtField6.Value)
            
        Case "Kupci"
            UpdateCell m_TableName, m_SelectedRow, "Naziv", Trim$(txtField1.Value)
            UpdateCell m_TableName, m_SelectedRow, "Ulica", Trim$(txtField2.Value)
            UpdateCell m_TableName, m_SelectedRow, "Mesto", Trim$(txtField3.Value)
            UpdateCell m_TableName, m_SelectedRow, "PostanskiBroj", Trim$(txtField4.Value)
            UpdateCell m_TableName, m_SelectedRow, "Drzava", Trim$(txtField5.Value)
            UpdateCell m_TableName, m_SelectedRow, "PIB", Trim$(txtField6.Value)
            UpdateCell m_TableName, m_SelectedRow, "MaticniBroj", Trim$(txtField7.Value)
            UpdateCell m_TableName, m_SelectedRow, "Email", Trim$(txtField8.Value)
            UpdateCell m_TableName, m_SelectedRow, "Hladnjaca", Trim$(txtField9.Value)
            UpdateCell m_TableName, m_SelectedRow, "TekuciRacun", Trim$(txtField10.Value)
            
        Case "Vozaci"
            UpdateCell m_TableName, m_SelectedRow, "Ime", Trim$(txtField1.Value)
            UpdateCell m_TableName, m_SelectedRow, "Prezime", Trim$(txtField2.Value)
            UpdateCell m_TableName, m_SelectedRow, "Telefon", Trim$(txtField3.Value)
            UpdateCell m_TableName, m_SelectedRow, "PIN", Trim$(txtField4.Value)
            
        Case "Parcele"
            Dim koopIDEdit As String
            koopIDEdit = ExtractIDFromDisplay(cmbField1.Value)

            UpdateCell m_TableName, m_SelectedRow, "KooperantID", koopIDEdit
            UpdateCell m_TableName, m_SelectedRow, "KatBroj", Trim$(txtField2.Value)
            UpdateCell m_TableName, m_SelectedRow, "KatOpstina", Trim$(txtField3.Value)
            UpdateCell m_TableName, m_SelectedRow, "Kultura", Trim$(cmbField2.Value)
            UpdateCell m_TableName, m_SelectedRow, "PovrsinaHa", CDbl(txtField5.Value)
            UpdateCell m_TableName, m_SelectedRow, "GGAPStatus", Trim$(cmbField3.Value)
            UpdateCell m_TableName, m_SelectedRow, "Napomena", Trim$(txtField7.Value)
            
        Case "Artikli"
            UpdateCell m_TableName, m_SelectedRow, "Naziv", Trim$(txtField1.Value)
            UpdateCell m_TableName, m_SelectedRow, "Tip", Trim$(cmbField5.Value)
            UpdateCell m_TableName, m_SelectedRow, "JedinicaMere", Trim$(cmbField6.Value)
            UpdateCell m_TableName, m_SelectedRow, "CenaPoJedinici", CDbl(Replace(txtField4.Value, ",", "."))
            UpdateCell m_TableName, m_SelectedRow, "DozaPoHa", CDbl(Replace(txtField6.Value, ",", "."))
            UpdateCell m_TableName, m_SelectedRow, "Kultura", Trim$(cmbField1.Value)
            UpdateCell m_TableName, m_SelectedRow, "Pakovanje", CDbl(Replace(txtField7.Value, ",", "."))
        End Select
    
    MsgBox "Izmenjeno!", vbInformation, APP_NAME
    LoadList
    Exit Sub
EH:
    LogErr "frmSledljivost.btnIzmeni"
    MsgBox "Greska pri izmeni: " & Err.Description, vbCritical, APP_NAME
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
        frmMain.Show
    End If
End Sub

Private Sub ClearFields()
    txtField1.Value = ""
    txtField2.Value = ""
    txtField3.Value = ""
    txtField4.Value = ""
    txtField5.Value = ""
    txtField6.Value = ""
    txtField7.Value = ""
    txtField8.Value = ""
    txtField9.Value = ""
    txtField10.Value = ""

    cmbField1.Value = ""
    cmbField2.Value = ""
    cmbField3.Value = ""
    cmbField4.Value = ""
    cmbField5.Value = ""
    cmbField6.Value = ""

    m_SelectedRow = 0
End Sub

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

    Dim katBroj As String
    Dim katOpstina As String
    Dim searchText As String

    katBroj = lstData.List(lstData.ListIndex, 2)
    katOpstina = lstData.List(lstData.ListIndex, 3)

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
    MsgBox "Greška: " & Err.Description, vbCritical

End Sub

Private Sub btnGeoSave_Click()
    On Error GoTo EH
    
    If Me.Tag <> "Parcele" Then Exit Sub
    
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If txtNCoord.Value = "" Or txtECoord.Value = "" Then
        MsgBox "Unesite N i E!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim nVal As Double
    Dim eVal As Double
    
    nVal = CDbl(txtNCoord.Value)
    eVal = CDbl(txtECoord.Value)
    
    Debug.Print "txtNCoord=" & txtNCoord.Value
    Debug.Print "txtECoord=" & txtECoord.Value
    Debug.Print "nVal=" & nVal
    Debug.Print "eVal=" & eVal
    
    SaveParcelGeoPoint m_SelectedRow, nVal, eVal
    
    MsgBox "Geo sacuvan ?", vbInformation, APP_NAME
    LoadList
    Exit Sub
    
EH:
    MsgBox "Greška: " & Err.Description, vbCritical
End Sub

Private Sub btnGeoClear_Click()
    On Error GoTo EH
    
    If Me.Tag <> "Parcele" Then Exit Sub
    
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu iz liste!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ClearParcelGeo m_SelectedRow
    MsgBox "Geo podaci su obrisani.", vbInformation, APP_NAME
    LoadList
    Exit Sub
    
EH:
    MsgBox "Greška pri brisanju geo podataka: " & Err.Description, vbCritical, APP_NAME
End Sub


Private Sub btnPasteCoords_Click()

    Dim txt As String
    Dim tokens() As String
    Dim cleanVals() As String
    Dim count As Long
    Dim i As Long
    
    txt = Trim$(GetClipboardText())
    
    If txt = "" Then
        MsgBox "Clipboard je prazan.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, ",", " ")
    txt = Replace(txt, ";", " ")
    
    tokens = Split(txt, " ")
    
    ReDim cleanVals(0 To 1)
    count = 0
    
    For i = LBound(tokens) To UBound(tokens)
        If IsNumeric(tokens(i)) Then
            If Len(tokens(i)) >= 5 Then
                cleanVals(count) = tokens(i)
                count = count + 1
                If count = 2 Then Exit For
            End If
        End If
    Next i
    
    If count < 2 Then
        MsgBox "Nisu pronadene validne koordinate.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' GeoSrbija obicno daje: N pa E
    txtNCoord.Value = cleanVals(0)
    txtECoord.Value = cleanVals(1)
    
    MsgBox "Koordinate ucitane ?", vbInformation, APP_NAME

End Sub

Private Sub btnOpenMap_Click()

    On Error GoTo EH
    
    If Me.Tag <> "Parcele" Then Exit Sub
    
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim data As Variant
    Dim latIdx As Long
    Dim lngIdx As Long
    Dim lat As Double
    Dim lng As Double
    
    data = GetTableData(TBL_PARCELE)
    If IsEmpty(data) Then Exit Sub
    
    latIdx = GetColumnIndex(TBL_PARCELE, "Lat")
    lngIdx = GetColumnIndex(TBL_PARCELE, "Lng")
    
    If latIdx = 0 Or lngIdx = 0 Then
        MsgBox "Lat/Lng kolone nisu pronadene.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    If Trim$(CStr(data(m_SelectedRow, latIdx))) = "" Or Trim$(CStr(data(m_SelectedRow, lngIdx))) = "" Then
        MsgBox "Parcela nema geo podatke.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    lat = CDbl(data(m_SelectedRow, latIdx))
    lng = CDbl(data(m_SelectedRow, lngIdx))
    
    OpenGoogleMaps lat, lng
    Exit Sub

EH:
    MsgBox "Greška: " & Err.Description, vbCritical, APP_NAME

End Sub

Private Sub btnOpenPolygonEditor_Click()

    If Me.Tag <> "Parcele" Then Exit Sub
    
    If m_SelectedRow = 0 Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim data As Variant
    Dim idIdx As Long
    Dim parcelaID As String
    
    data = GetTableData(TBL_PARCELE)
    idIdx = GetColumnIndex(TBL_PARCELE, "ParcelaID")
    
    parcelaID = CStr(data(m_SelectedRow, idIdx))
    
    OpenParcelPolygonEditor parcelaID

End Sub

'========================
'HELPERS
'========================
Private Sub SafeSetCombo(cmb As MSForms.ComboBox, ByVal v As String)
    Dim i As Long

    If Len(Trim$(v)) = 0 Then
        cmb.ListIndex = -1
        Exit Sub
    End If

    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = v Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i

    cmb.ListIndex = -1
End Sub

Private Sub LoadStaniceIntoCombo()
    Dim data As Variant
    Dim i As Long

    cmbField1.Clear

    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then Exit Sub

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, 1))) <> "" Then
            ' display name
            cmbField1.AddItem NzToText(data(i, 2)) ' Naziv

            ' store ID in hidden column
            cmbField1.List(cmbField1.ListCount - 1, 1) = NzToText(data(i, 1)) ' StanicaID
        End If
    Next i

    cmbField1.ColumnCount = 2
    cmbField1.ColumnWidths = "150 pt;0 pt" ' hide ID column
End Sub

Public Sub OpenGoogleMaps(ByVal lat As Double, ByVal lng As Double)

    Dim url As String
    url = "https://www.google.com/maps?q=" & _
          Replace(CStr(lat), ",", ".") & "," & _
          Replace(CStr(lng), ",", ".")
    
    ThisWorkbook.FollowHyperlink url

End Sub

Public Sub OpenParcelPolygonEditor(ByVal parcelaID As String)
    Dim url As String
    url = "https://dusanmiladinovicvnm.github.io/otkupapp-pwa/parcel-draw.html?parcelaId=" & parcelaID
    ThisWorkbook.FollowHyperlink url
End Sub
