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
'
' Header-Labels über lstPrijemnice:
'   BrojPrij | BrojZbirne | Datum | Kolicina | Cena | Vrednost | TipAmb | KolAmb

Private m_SetupDone As Boolean
Private m_PrijemniceData As Variant
Private m_DataIndices() As Long

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

Private Sub UserForm_Activate()
    
    If Not mChromeRemoved Then
        Me.Caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If
    ApplyTheme Me, BG_MAIN
    If m_SetupDone Then Exit Sub
    m_SetupDone = True
    
    Me.Caption = "Fakturisanje"
    
    cmbKupac.Clear
    Dim kupci As Variant
    kupci = GetLookupList(TBL_KUPCI, "Naziv")
    If IsArray(kupci) Then
        Dim i As Long
        For i = LBound(kupci) To UBound(kupci)
            cmbKupac.AddItem CStr(kupci(i))
        Next i
    End If

    With lstPrijemnice
        .ColumnCount = 8
        .ColumnWidths = "70;70;65;35;65;55;80;140"
        .MultiSelect = fmMultiSelectMulti
    End With
End Sub

Private Sub chkPrikaziFakturisane_Click()
    If cmbKupac.Value <> "" Then btnUnesi_Click
End Sub

' ============================================================
' PRIJEMNICE LADEN
' ============================================================

Private Sub btnUnesi_Click()
    On Error GoTo EH
    If cmbKupac.Value = "" Then
        MsgBox "Izaberite kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    lstPrijemnice.Clear
    
    Dim kupacID As String
    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbKupac.Value, "KupacID"))
    
    ' Alle Prijemnice für diesen Kupac
    m_PrijemniceData = GetPrijemniceByKupac(kupacID)
    m_PrijemniceData = ExcludeStornirano(m_PrijemniceData, TBL_PRIJEMNICA)
    
    ' Uplata-Dict einmal bauen
    Dim uplataDict As Object
    Set uplataDict = BuildUplataDictByFaktura()
    If IsEmpty(m_PrijemniceData) Then
        MsgBox "Nema prijemnica za ovog kupca!", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ' TODO: Filter auf nefakturisane (wenn Fakturisano-Spalte in tblPrijemnica existiert)
    ' Vorerst alle anzeigen
    
    Dim colBroj As Long, colBrZbr As Long, colDatum As Long
    Dim colKol As Long, colCena As Long, colKlasa As Long, colFakturisano As Long, colFakturaID As Long
    colBroj = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    colDatum = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    colKlasa = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA)
    colFakturisano = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO)
    colFakturaID = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID)
    
    ' Zähler für gefilterte Daten-Indizes

    ReDim m_DataIndices(0 To UBound(m_PrijemniceData, 1))
    Dim count As Long
    
    Dim i As Long
    For i = 1 To UBound(m_PrijemniceData, 1)
    
        ' Nur nefakturisane anzeigen
        If colFakturisano > 0 Then
            If CStr(m_PrijemniceData(i, colFakturisano)) = "Da" Then
                If Not chkPrikaziFakturisane.Value Then GoTo NextPrij
            End If
        End If
        
        m_DataIndices(count) = i   ' ListBox-Index count ? Array-Index i
        count = count + 1
        
        lstPrijemnice.AddItem CStr(m_PrijemniceData(i, colBroj))
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 1) = CStr(m_PrijemniceData(i, colBrZbr))
        
        If IsDate(m_PrijemniceData(i, colDatum)) Then
            lstPrijemnice.List(lstPrijemnice.ListCount - 1, 2) = Format$(CDate(m_PrijemniceData(i, colDatum)), "d.m.yyyy")
        End If
        ' Klasa
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 3) = CStr(m_PrijemniceData(i, colKlasa))
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 4) = Format$(m_PrijemniceData(i, colKol), "#,##0.00")
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 5) = Format$(m_PrijemniceData(i, colCena), "#,##0.00")
        
        Dim vrednost As Double
        vrednost = 0
        If IsNumeric(m_PrijemniceData(i, colKol)) And IsNumeric(m_PrijemniceData(i, colCena)) Then
            vrednost = CDbl(m_PrijemniceData(i, colKol)) * CDbl(m_PrijemniceData(i, colCena))
        End If
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 6) = Format$(vrednost, "#,##0.00")
        
        ' Spalte 7: Fakturisano
        lstPrijemnice.List(lstPrijemnice.ListCount - 1, 7) = ""
        
        If colFakturisano > 0 Then
            If CStr(m_PrijemniceData(i, colFakturisano)) = "Da" Then
                Dim fakID As String
                fakID = CStr(m_PrijemniceData(i, colFakturaID))
                Dim brojFakture As String
                brojFakture = CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakID, COL_FAK_BROJ))
                Dim fakIznos As Double
                fakIznos = CDbl(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakID, COL_FAK_IZNOS))
                
                Dim fakUplaceno As Double
                fakUplaceno = 0
                If uplataDict.Exists(fakID) Then fakUplaceno = uplataDict(fakID)
                
                lstPrijemnice.List(lstPrijemnice.ListCount - 1, 7) = _
                    brojFakture & " (" & Format$(fakUplaceno, "#,##0") & "/" & Format$(fakIznos, "#,##0") & ")"
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
    MsgBox "Greska pri ucitavanju prijemnica: " & Err.Description, vbCritical, APP_NAME
End Sub

' ============================================================
' FAKTURA ERSTELLEN
' ============================================================

Private Sub btnIzradiFakturu_Click()
    On Error GoTo EH
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
    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbKupac.Value, "KupacID"))
    
    ' Stavke sammeln
    Dim stavke As New Collection
    Dim colID As Long, colKol As Long, colCena As Long, colBroj As Long, colKlasa As Long
    colID = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    colBroj = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    colKlasa = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA)
    
    For i = 0 To lstPrijemnice.ListCount - 1
        If lstPrijemnice.Selected(i) Then
            Dim dataRow As Long
            dataRow = m_DataIndices(i)   ' ? korrektes Mapping
            
            ' Array: PrijemnicaID, Kolicina, Cena, Klasa, BrojPrijemnice
            Dim stavka As Variant
            stavka = Array( _
                CStr(m_PrijemniceData(dataRow, colID)), _
                CDbl(m_PrijemniceData(dataRow, colKol)), _
                CDbl(m_PrijemniceData(dataRow, colCena)), _
                CStr(m_PrijemniceData(dataRow, colKlasa)), _
                CStr(m_PrijemniceData(dataRow, colBroj)))
            stavke.Add stavka
        End If
    Next i
    
    ' Bestätigung
    Dim msg As String
    msg = "Kreirati fakturu za " & cmbKupac.Value & "?" & vbCrLf & _
          "Broj stavki: " & stavke.count & vbCrLf & _
          "Ukupan iznos: " & Format$(CalculateTotal(stavke), "#,##0.00") & " RSD"
    
    If MsgBox(msg, vbQuestion + vbYesNo, APP_NAME) = vbNo Then
        Exit Sub
    End If
    
    Dim fakturaID As String
    fakturaID = CreateFaktura_TX(kupacID, stavke)
    
    If fakturaID <> "" Then
        MsgBox "Faktura kreirana: " & fakturaID, vbInformation, APP_NAME
        btnUnesi_Click
    Else
        MsgBox "Greska pri kreiranju fakture!", vbCritical, APP_NAME
    End If
    Exit Sub
EH:
    LogErr "frmFakturisanje.btnIzradiFakturu"
    MsgBox "Greska pri izradi fakture: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Function CalculateTotal(ByVal stavke As Collection) As Double
    Dim s As Variant
    Dim total As Double
    For Each s In stavke
        total = total + (CDbl(s(1)) * CDbl(s(2)))
    Next s
    CalculateTotal = total
End Function

' ============================================================
' DRUCKEN
' ============================================================

Private Sub btnStampaj_Click()
    On Error GoTo EH
    Dim kupacID As String
    kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbKupac.Value, "KupacID"))
    
    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then
        MsgBox "Nema faktura!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim colKupac As Long, colID As Long
    colKupac = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)
    colID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    
    Dim colStorno As Long
    colStorno = GetColumnIndex(TBL_FAKTURE, COL_STORNIRANO)
    
    Dim lastFakturaID As String
    Dim i As Long
    For i = UBound(data, 1) To 1 Step -1
        If colStorno > 0 Then
            If CStr(data(i, colStorno)) = "Da" Then GoTo NextFak
        End If
        If CStr(data(i, colKupac)) = kupacID Then
            lastFakturaID = CStr(data(i, colID))
            Exit For
        End If
NextFak:
    Next i
    
    If lastFakturaID = "" Then
        MsgBox "Nema faktura za ovog kupca!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    PrintFaktura lastFakturaID
    Exit Sub
EH:
    LogErr "frmFakturisanje.btnStampaj"
    MsgBox "Greska pri stampanju: " & Err.Description, vbCritical, APP_NAME
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


