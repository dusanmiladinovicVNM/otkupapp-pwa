VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFakturisanje 
   Caption         =   "UserForm1"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21495
   OleObjectBlob   =   "frmFakturisanje.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFakturisanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ============================================================
' frmFakturisanje v2.1 – Rechnungserstellung
' GEÄNDERT: Basiert auf tblPrijemnica statt tblIsporuka
' Faktura = Prijemnica.Kolicina × Prijemnica.Cena
' ============================================================
'
' CONTROLS:
'   cmbKupac (ComboBox) – Label: "Izaberi Kupca/Hladnjacu"
'   btnUnesi (CommandButton) – Caption: "Unesi"
'   lstPrijemnice (ListBox) – ColumnCount=8, MultiSelect
'   btnIzradiFakturu (CommandButton) – Caption: "Izradi Fakturu"
'   btnStampaj (CommandButton) – Caption: "Stampaj"
'   btnPovratak (CommandButton) – Caption: "Povratak u glavni meni"
'   cmbFaktura (ComboBox) – Faktura za štampu, display text + hidden FakturaID
'
' Header-Labels über lstPrijemnice:
'   BrojPrij | BrojZbirne | Datum | Klasa | Kolicina | Cena | Vrednost | Fakturisano

Private m_SetupDone As Boolean
Private m_IsLoading As Boolean
Private m_PrijemniceData As Variant
Private m_DataIndices() As Long

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

Private Sub UserForm_Activate()
    On Error GoTo EH

    If Not mChromeRemoved Then
        RemoveTitleBar
        Me.caption = ""
        mChromeRemoved = True
    End If

    If m_SetupDone Then Exit Sub
    m_SetupDone = True

    ApplyTheme Me, BG_MAIN

    ' Kupac combo: display = Naziv, hidden ID = KupacID
    FillComboDisplayID cmbKupac, TBL_KUPCI, COL_KUP_NAZIV, COL_KUP_ID

    With lstPrijemnice
        .Clear
        .ColumnCount = 8
        .ColumnWidths = "70;70;65;35;65;55;80;140"
        .MultiSelect = fmMultiSelectMulti
    End With
    
    With cmbFaktura
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "260 pt;0 pt"
        .BoundColumn = 1
        .TextColumn = 1
    End With

    Exit Sub

EH:
    LogErr "frmFakturisanje.UserForm_Activate"
    MsgBox "Greška pri otvaranju fakturisanja: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub chkPrikaziFakturisane_Click()
    If cmbKupac.value <> "" Then btnUnesi_Click
End Sub

Private Sub cmbKupac_Change()
    On Error GoTo EH

    If m_IsLoading Then Exit Sub

    lstPrijemnice.Clear
    cmbFaktura.Clear
    Erase m_DataIndices
    m_PrijemniceData = Empty

    If cmbKupac.value <> "" Then
        FillFaktureZaKupca
    End If

    Exit Sub

EH:
    LogErr "frmFakturisanje.cmbKupac_Change"
End Sub

' ============================================================
' PRIJEMNICE LADEN
' ============================================================

Private Sub btnUnesi_Click()
    On Error GoTo EH

    If cmbKupac.value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    lstPrijemnice.Clear
    Erase m_DataIndices
    m_PrijemniceData = Empty

    Dim kupacID As String
    kupacID = GetComboID(cmbKupac)

    If kupacID = "" Then
        MsgBox "Nije pronaden ID kupca.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    FillFaktureZaKupca

    m_PrijemniceData = GetPrijemniceByKupac(kupacID)

    If IsEmpty(m_PrijemniceData) Then
        MsgBox "Nema prijemnica za ovog kupca!", vbInformation, APP_NAME
        Exit Sub
    End If

    m_PrijemniceData = ExcludeStornirano(m_PrijemniceData, TBL_PRIJEMNICA)

    If IsEmpty(m_PrijemniceData) Then
        MsgBox "Nema aktivnih prijemnica za ovog kupca!", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim uplataDict As Object
    Set uplataDict = BuildUplataDictByFaktura()

    Dim colBroj As Long
    Dim colBrZbr As Long
    Dim colDatum As Long
    Dim colKol As Long
    Dim colCena As Long
    Dim colKlasa As Long
    Dim colFakturisano As Long
    Dim colFakturaID As Long

    colBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, _
                                 "frmFakturisanje.btnUnesi")
    colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                  "frmFakturisanje.btnUnesi")
    colDatum = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, _
                                  "frmFakturisanje.btnUnesi")
    colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                "frmFakturisanje.btnUnesi")
    colCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, _
                                 "frmFakturisanje.btnUnesi")
    colKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, _
                                  "frmFakturisanje.btnUnesi")
    colFakturisano = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, _
                                        "frmFakturisanje.btnUnesi")
    colFakturaID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, _
                                      "frmFakturisanje.btnUnesi")

    ReDim m_DataIndices(0 To UBound(m_PrijemniceData, 1))

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(m_PrijemniceData, 1)

        If Trim$(CStr(m_PrijemniceData(i, colFakturisano))) = "Da" Then
            If Not chkPrikaziFakturisane.value Then GoTo NextPrij
        End If

        m_DataIndices(count) = i
        count = count + 1

        lstPrijemnice.AddItem CStr(m_PrijemniceData(i, colBroj))
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 1) = CStr(m_PrijemniceData(i, colBrZbr))

        If IsDate(m_PrijemniceData(i, colDatum)) Then
            lstPrijemnice.List(lstPrijemnice.ListCount - 1, 2) = _
                Format$(CDate(m_PrijemniceData(i, colDatum)), "d.m.yyyy")
        Else
            lstPrijemnice.List(lstPrijemnice.ListCount - 1, 2) = ""
        End If

        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 3) = CStr(m_PrijemniceData(i, colKlasa))

        Dim kol As Double
        Dim cena As Double
        Dim vrednost As Double

        kol = 0
        cena = 0
        vrednost = 0

        If IsNumeric(m_PrijemniceData(i, colKol)) Then kol = CDbl(m_PrijemniceData(i, colKol))
        If IsNumeric(m_PrijemniceData(i, colCena)) Then cena = CDbl(m_PrijemniceData(i, colCena))

        vrednost = kol * cena

        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 4) = Format$(kol, "#,##0.00")
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 5) = Format$(cena, "#,##0.00")
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 6) = Format$(vrednost, "#,##0.00")
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 7) = ""

        If Trim$(CStr(m_PrijemniceData(i, colFakturisano))) = "Da" Then
            Dim fakID As String
            fakID = Trim$(CStr(m_PrijemniceData(i, colFakturaID)))

            If fakID <> "" Then
                Dim brojFakture As String
                brojFakture = CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakID, COL_FAK_BROJ))

                Dim fakIznos As Double
                Dim fakIznosVal As Variant
                fakIznosVal = LookupValue(TBL_FAKTURE, COL_FAK_ID, fakID, COL_FAK_IZNOS)

                If IsNumeric(fakIznosVal) Then fakIznos = CDbl(fakIznosVal)

                Dim fakUplaceno As Double
                fakUplaceno = 0

                If Not uplataDict Is Nothing Then
                    If uplataDict.Exists(fakID) Then fakUplaceno = CDbl(uplataDict(fakID))
                End If

                lstPrijemnice.List(lstPrijemnice.ListCount - 1, 7) = _
                    brojFakture & " (" & Format$(fakUplaceno, "#,##0") & _
                    "/" & Format$(fakIznos, "#,##0") & ")"
            Else
                lstPrijemnice.List(lstPrijemnice.ListCount - 1, 7) = "Fakturisano"
            End If
        End If

NextPrij:
    Next i

    If count = 0 Then
        MsgBox "Sve prijemnice su vec fakturisane!", vbInformation, APP_NAME
    End If

    Exit Sub

EH:
    LogErr "frmFakturisanje.btnUnesi"
    MsgBox "Greška pri ucitavanju prijemnica: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub FillFaktureZaKupca()
    On Error GoTo EH

    cmbFaktura.Clear
    cmbFaktura.ColumnCount = 2
    cmbFaktura.ColumnWidths = "260 pt;0 pt"
    cmbFaktura.BoundColumn = 1
    cmbFaktura.TextColumn = 1

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
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, _
                               "frmFakturisanje.FillFaktureZaKupca")
    colBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, _
                                 "frmFakturisanje.FillFaktureZaKupca")
    colDatum = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM, _
                                  "frmFakturisanje.FillFaktureZaKupca")
    colKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, _
                                  "frmFakturisanje.FillFaktureZaKupca")
    colIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, _
                                  "frmFakturisanje.FillFaktureZaKupca")
    colStatus = RequireColumnIndex(TBL_FAKTURE, COL_FAK_STATUS, _
                                   "frmFakturisanje.FillFaktureZaKupca")
    colStorno = RequireColumnIndex(TBL_FAKTURE, COL_STORNIRANO, _
                                   "frmFakturisanje.FillFaktureZaKupca")

    Dim i As Long
    Dim fakturaID As String
    Dim brojFakture As String
    Dim datumTxt As String
    Dim iznos As Double
    Dim status As String
    Dim displayText As String

    For i = UBound(data, 1) To 1 Step -1
        If Trim$(CStr(data(i, colStorno))) = "Da" Then GoTo NextFak

        If Trim$(CStr(data(i, colKupac))) = kupacID Then
            fakturaID = Trim$(CStr(data(i, colID)))
            brojFakture = Trim$(CStr(data(i, colBroj)))
            status = Trim$(CStr(data(i, colStatus)))

            datumTxt = ""
            If IsDate(data(i, colDatum)) Then
                datumTxt = Format$(CDate(data(i, colDatum)), "d.m.yyyy")
            End If

            iznos = 0
            If IsNumeric(data(i, colIznos)) Then iznos = CDbl(data(i, colIznos))

            displayText = brojFakture

            If datumTxt <> "" Then
                displayText = displayText & " | " & datumTxt
            End If

            displayText = displayText & " | " & Format$(iznos, "#,##0.00") & " RSD"

            If status <> "" Then
                displayText = displayText & " | " & status
            End If

            cmbFaktura.AddItem displayText
            cmbFaktura.List(cmbFaktura.ListCount - 1, 1) = fakturaID
        End If

NextFak:
    Next i

    If cmbFaktura.ListCount > 0 Then
        cmbFaktura.ListIndex = 0
    End If

    Exit Sub

EH:
    LogErr "frmFakturisanje.FillFaktureZaKupca"
    cmbFaktura.Clear
End Sub

' ============================================================
' FAKTURA ERSTELLEN
' ============================================================

Private Sub btnIzradiFakturu_Click()
    On Error GoTo EH
    
    If cmbKupac.value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If lstPrijemnice.ListCount = 0 Or IsEmpty(m_PrijemniceData) Then
        MsgBox "Prvo ucitajte prijemnice za kupca.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim selectedCount As Long
    Dim i As Long

    For i = 0 To lstPrijemnice.ListCount - 1
        If lstPrijemnice.Selected(i) Then selectedCount = selectedCount + 1
    Next i

    If selectedCount = 0 Then
        MsgBox "Izaberite stavke za fakturisanje!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim kupacID As String
    kupacID = GetComboID(cmbKupac)

    If kupacID = "" Then
        MsgBox "Nije pronaden ID kupca.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim colID As Long
    Dim colKol As Long
    Dim colCena As Long
    Dim colBroj As Long
    Dim colKlasa As Long
    Dim colFakturisano As Long
    Dim colFakturaID As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, _
                               "frmFakturisanje.btnIzradiFakturu")
    colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                "frmFakturisanje.btnIzradiFakturu")
    colCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, _
                                 "frmFakturisanje.btnIzradiFakturu")
    colBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, _
                                 "frmFakturisanje.btnIzradiFakturu")
    colKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, _
                                  "frmFakturisanje.btnIzradiFakturu")
    colFakturisano = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, _
                                        "frmFakturisanje.btnIzradiFakturu")
    colFakturaID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, _
                                      "frmFakturisanje.btnIzradiFakturu")
    colStorno = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, _
                                   "frmFakturisanje.btnIzradiFakturu")

    Dim stavke As New Collection
    Dim seenPrijemnice As Object
    Set seenPrijemnice = CreateObject("Scripting.Dictionary")

    For i = 0 To lstPrijemnice.ListCount - 1
        If lstPrijemnice.Selected(i) Then

            Dim dataRow As Long
            dataRow = m_DataIndices(i)

            If dataRow <= 0 Or dataRow > UBound(m_PrijemniceData, 1) Then
                MsgBox "Neispravno mapiranje izabrane prijemnice.", vbCritical, APP_NAME
                Exit Sub
            End If

            Dim prijemnicaID As String
            prijemnicaID = Trim$(CStr(m_PrijemniceData(dataRow, colID)))

            If prijemnicaID = "" Then
                MsgBox "Izabrana prijemnica nema ID.", vbCritical, APP_NAME
                Exit Sub
            End If

            If seenPrijemnice.Exists(prijemnicaID) Then
                MsgBox "Ista prijemnica je izabrana više puta: " & prijemnicaID, _
                       vbExclamation, APP_NAME
                Exit Sub
            End If

            seenPrijemnice.Add prijemnicaID, True

            If Trim$(CStr(m_PrijemniceData(dataRow, colStorno))) = "Da" Then
                MsgBox "Izabrana prijemnica je stornirana i ne može se fakturisati.", _
                       vbExclamation, APP_NAME
                Exit Sub
            End If

            If Trim$(CStr(m_PrijemniceData(dataRow, colFakturisano))) = "Da" Or _
               Trim$(CStr(m_PrijemniceData(dataRow, colFakturaID))) <> "" Then

                MsgBox "Izabrana prijemnica je vec fakturisana i ne može biti ukljucena u novu fakturu.", _
                       vbExclamation, APP_NAME
                Exit Sub
            End If

            If Not IsNumeric(m_PrijemniceData(dataRow, colKol)) Then
                MsgBox "Kolicina nije ispravna za prijemnicu: " & prijemnicaID, _
                       vbExclamation, APP_NAME
                Exit Sub
            End If

            If Not IsNumeric(m_PrijemniceData(dataRow, colCena)) Then
                MsgBox "Cena nije ispravna za prijemnicu: " & prijemnicaID, _
                       vbExclamation, APP_NAME
                Exit Sub
            End If

            Dim kolicina As Double
            Dim cena As Double

            kolicina = CDbl(m_PrijemniceData(dataRow, colKol))
            cena = CDbl(m_PrijemniceData(dataRow, colCena))

            If kolicina <= 0 Then
                MsgBox "Kolicina mora biti veca od nule za prijemnicu: " & prijemnicaID, _
                       vbExclamation, APP_NAME
                Exit Sub
            End If

            If cena < 0 Then
                MsgBox "Cena ne sme biti negativna za prijemnicu: " & prijemnicaID, _
                       vbExclamation, APP_NAME
                Exit Sub
            End If

            Dim stavka As Variant
            stavka = Array( _
                prijemnicaID, _
                kolicina, _
                cena, _
                CStr(m_PrijemniceData(dataRow, colKlasa)), _
                CStr(m_PrijemniceData(dataRow, colBroj)) _
            )

            stavke.Add stavka
        End If
    Next i

    If stavke.count = 0 Then
        MsgBox "Nema validnih stavki za fakturisanje.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim total As Double
    total = CalculateTotal(stavke)

    If total <= 0 Then
        MsgBox "Ukupan iznos fakture mora biti veci od nule.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim msg As String
    msg = "Kreirati fakturu za " & cmbKupac.value & "?" & vbCrLf & _
          "Broj stavki: " & stavke.count & vbCrLf & _
          "Ukupan iznos: " & Format$(total, "#,##0.00") & " RSD"

    If MsgBox(msg, vbQuestion + vbYesNo, APP_NAME) = vbNo Then
        Exit Sub
    End If

    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)

    If fakturaID <> "" Then
        MsgBox "Faktura kreirana: " & fakturaID, vbInformation, APP_NAME

        btnUnesi_Click
        FillFaktureZaKupca
        SetComboByID cmbFaktura, fakturaID

    Else
        MsgBox "Greška pri kreiranju fakture. Promene su vracene.", vbCritical, APP_NAME
    End If

    Exit Sub

EH:
    LogErr "frmFakturisanje.btnIzradiFakturu"
    MsgBox "Greška pri izradi fakture: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Function CalculateTotal(ByVal stavke As Collection) As Double
    On Error GoTo EH

    If stavke Is Nothing Then Exit Function

    Dim s As Variant
    Dim total As Double

    For Each s In stavke
        If IsNumeric(s(1)) And IsNumeric(s(2)) Then
            total = total + (CDbl(s(1)) * CDbl(s(2)))
        End If
    Next s

    CalculateTotal = total
    Exit Function

EH:
    LogErr "frmFakturisanje.CalculateTotal"
    CalculateTotal = 0
End Function

' ============================================================
' DRUCKEN
' ============================================================

Private Sub btnStampaj_Click()
    On Error GoTo EH

    If cmbKupac.value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If cmbFaktura.value = "" Then
        FillFaktureZaKupca
    End If

    If cmbFaktura.value = "" Then
        MsgBox "Nema faktura za ovog kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim fakturaID As String
    fakturaID = GetComboID(cmbFaktura)

    If fakturaID = "" Then
        MsgBox "Nije pronaden ID fakture za štampu.", vbExclamation, APP_NAME
        Exit Sub
    End If

    PrintFaktura fakturaID
    Exit Sub

EH:
    LogErr "frmFakturisanje.btnStampaj"
    MsgBox "Greška pri štampanju: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' NAVIGATION
' ============================================================


Private Sub btnSEF_Click()
    frmSEF.Show
End Sub

Private Sub btnPovratak_Click()
    Me.Hide
    frmOtkupAPP.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        frmOtkupAPP.Show
    End If
End Sub



