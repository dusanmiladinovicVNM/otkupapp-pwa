VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAgrohemija 
   Caption         =   "UserForm1"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14400
   OleObjectBlob   =   "frmAgrohemija.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAgrohemija"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tKorpaItem
    artikalID As String
    ArtikalNaziv As String
    kolicina As Double
    cena As Double
    vrednost As Double
    parcelaID As String
    jm As String
End Type

Private m_KorpaIzlaz() As tKorpaItem
Private m_KorpaIzlazCount As Long
Private m_KorpaUlaz() As tKorpaItem
Private m_KorpaUlazCount As Long

Private m_ParcelaIDs() As String
Private m_ParcelaHa() As Double

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
    ApplyTheme Me, BG_MAIN()
    LoadKooperanti
    LoadArtikli
    LoadArtikliUlaz
End Sub

Private Sub UserForm_Activate()
    If Not mChromeRemoved Then
        Me.caption = ""
        RemoveTitleBar
        mChromeRemoved = True
    End If
End Sub

Private Sub LoadKooperanti()
    cmbKooperant.Clear
    Dim data As Variant
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Sub
    
    Dim colID As Long, colIme As Long, colPrez As Long
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrez = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 1)) <> "" Then
            cmbKooperant.AddItem CStr(data(i, colIme)) & " " & _
                CStr(data(i, colPrez)) & " (" & CStr(data(i, colID)) & ")"
        End If
    Next i
End Sub

Private Sub LoadArtikli()
    cmbArtikal.Clear
    Dim data As Variant
    data = GetTableData(TBL_ARTIKLI)
    If IsEmpty(data) Then Exit Sub
    
    Dim colID As Long, colNaziv As Long, colJM As Long
    colID = GetColumnIndex(TBL_ARTIKLI, COL_ART_ID)
    colNaziv = GetColumnIndex(TBL_ARTIKLI, COL_ART_NAZIV)
    colJM = GetColumnIndex(TBL_ARTIKLI, COL_ART_JM)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 1)) <> "" Then
            cmbArtikal.AddItem CStr(data(i, colNaziv)) & " [" & CStr(data(i, colJM)) & "] (" & CStr(data(i, colID)) & ")"
        End If
    Next i
End Sub

Private Sub cmbKooperant_Change()
    lstParcele.Clear
    Erase m_ParcelaIDs
    Erase m_ParcelaHa
    lblPreporuka.caption = ""
    lblVrednost.caption = ""
    txtKolicina.value = ""
    
    If cmbKooperant.value = "" Then Exit Sub
    
    Dim koopID As String
    koopID = ExtractIDFromDisplay(cmbKooperant.value)
    
    Dim parcele As Variant
    parcele = GetParceleByKooperant(koopID)
    If IsEmpty(parcele) Then Exit Sub
    
    ReDim m_ParcelaIDs(0 To UBound(parcele, 1) - 1)
    ReDim m_ParcelaHa(0 To UBound(parcele, 1) - 1)
    
    Dim i As Long
    For i = 1 To UBound(parcele, 1)
        m_ParcelaIDs(i - 1) = CStr(parcele(i, 1))
        m_ParcelaHa(i - 1) = CDbl(parcele(i, 5))
        lstParcele.AddItem CStr(parcele(i, 6))  ' Display string
    Next i
    
    ' Dug anzeigen
    Dim dug As Double
    dug = GetAgrohemijaDug(koopID) - GetAgroAbzug(koopID)
    If dug > 0 Then
        lblDug.caption = "Dug: " & Format$(dug, "#,##0") & " RSD"
    Else
        lblDug.caption = ""
    End If
End Sub

Private Sub cmbArtikal_Change()
    UpdatePreporuka
End Sub

Private Sub lstParcele_Click()
    UpdatePreporuka
End Sub

Private Sub UpdatePreporuka()
    lblPreporuka.caption = ""
    lblVrednost.caption = ""
    txtKolicina.value = ""
    
    If cmbArtikal.value = "" Then Exit Sub
    
    ' Summiere ha aller ausgewählten Parcele
    Dim totalHa As Double
    Dim selectedParcele As String
    Dim i As Long
    For i = 0 To lstParcele.ListCount - 1
        If lstParcele.Selected(i) Then
            totalHa = totalHa + m_ParcelaHa(i)
            If selectedParcele <> "" Then selectedParcele = selectedParcele & ","
            selectedParcele = selectedParcele & m_ParcelaIDs(i)
        End If
    Next i
    
    If totalHa = 0 Then Exit Sub
    
    Dim artikalID As String
    artikalID = ExtractIDFromDisplay(cmbArtikal.value)
    
    Dim preporuka As Double
    preporuka = CalculatePreporuka(artikalID, totalHa)
    
    Dim jm As String
    jm = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_JM))
    
    ' Nach Berechnung der Preporuka:
    Dim pakovanje As Double
    Dim pakStr As String
    pakStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_PAKOVANJE))
    If IsNumeric(pakStr) And CDbl(pakStr) > 0 Then
        pakovanje = CDbl(pakStr)
        
        ' Aufrunden auf ganze Verpackungen
        Dim brojPakovanja As Long
        brojPakovanja = -Int(-preporuka / pakovanje)  ' Ceiling
        Dim izdajKol As Double
        izdajKol = brojPakovanja * pakovanje
        
        lblPreporuka.caption = "Doza: " & Format$(preporuka, "0.00") & " " & jm & " za " & Format$(totalHa, "0.00") & " ha" & vbCrLf & _
                               "Izdavanje: " & brojPakovanja & " x " & FormatKol(pakovanje) & " " & jm & " = " & FormatKol(izdajKol) & " " & jm
        txtKolicina.value = FormatKol(izdajKol)
    Else
        lblPreporuka.caption = "Doza: " & Format$(preporuka, "0.00") & " " & jm & " za " & Format$(totalHa, "0.00") & " ha"
        txtKolicina.value = Format$(preporuka, "0.00")
    End If
    
    UpdateVrednost
End Sub

Private Sub txtKolicina_Change()
    UpdateVrednost
End Sub

Private Sub UpdateVrednost()
    lblVrednost.caption = ""
    If cmbArtikal.value = "" Then Exit Sub
    If Not IsNumeric(txtKolicina.value) Then Exit Sub
    
    Dim artikalID As String
    artikalID = ExtractIDFromDisplay(cmbArtikal.value)
    
    Dim cena As Double
    Dim cenaStr As String
    cenaStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_CENA))
    If IsNumeric(cenaStr) Then cena = CDbl(cenaStr)
    
    Dim vrednost As Double
    vrednost = CDbl(txtKolicina.value) * cena
    
    lblVrednost.caption = "Vrednost: " & Format$(vrednost, "#,##0") & " RSD"
End Sub

Private Sub btnDodajIzlaz_Click()
    On Error GoTo EH
    
    If cmbArtikal.value = "" Then
        MsgBox "Izaberite artikal!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not IsNumeric(txtKolicina.value) Or CDbl(txtKolicina.value) <= 0 Then
        MsgBox "Unesite validnu kolicinu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Mindestens eine Parcela ausgewählt
    Dim hasSelection As Boolean
    Dim i As Long
    For i = 0 To lstParcele.ListCount - 1
        If lstParcele.Selected(i) Then hasSelection = True: Exit For
    Next i
    If Not hasSelection Then
        MsgBox "Izaberite parcelu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim artikalID As String
    artikalID = ExtractIDFromDisplay(cmbArtikal.value)
    
    Dim artNaziv As String
    artNaziv = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_NAZIV))
    Dim jm As String
    jm = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_JM))
    Dim cena As Double
    Dim cenaStr As String
    cenaStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_CENA))
    If IsNumeric(cenaStr) Then cena = CDbl(cenaStr)
    
    Dim kol As Double
    kol = CDbl(txtKolicina.value)
    
    ' Nach Kolicina-Prüfung:
    Dim pakStr2 As String
    pakStr2 = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_PAKOVANJE))
    If IsNumeric(pakStr2) And CDbl(pakStr2) > 0 Then
        Dim pak As Double: pak = CDbl(pakStr2)
        If Abs(kol / pak - Int(kol / pak + 0.001)) > 0.01 Then
            MsgBox "Kolicina mora biti u celim pakovanjima (" & Format$(pak, "0.##") & " " & jm & ")!", vbExclamation, APP_NAME
            Exit Sub
        End If
    End If
    
    ' Parcela IDs sammeln
    Dim parcelaIDs As String
    For i = 0 To lstParcele.ListCount - 1
        If lstParcele.Selected(i) Then
            If parcelaIDs <> "" Then parcelaIDs = parcelaIDs & ";"
            parcelaIDs = parcelaIDs & m_ParcelaIDs(i)
        End If
    Next i
    
    ' In Korpa Array
    m_KorpaIzlazCount = m_KorpaIzlazCount + 1
    ReDim Preserve m_KorpaIzlaz(1 To m_KorpaIzlazCount)
    With m_KorpaIzlaz(m_KorpaIzlazCount)
        .artikalID = artikalID
        .ArtikalNaziv = artNaziv
        .kolicina = kol
        .cena = cena
        .vrednost = kol * cena
        .parcelaID = parcelaIDs
        .jm = jm
    End With
    
    ' ListBox aktualisieren
    lstKorpa.AddItem artNaziv & " - " & Format$(kol, "0.00") & " " & jm & " = " & Format$(kol * cena, "#,##0") & " RSD"
    
    ' Felder zurücksetzen
    cmbArtikal.value = ""
    txtKolicina.value = ""
    lblPreporuka.caption = ""
    lblVrednost.caption = ""
    
    Exit Sub
EH:
    LogErr "frmAgrohemija.btnDodajIzlaz"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnZavrsiIzlaz_Click()
    On Error GoTo EH
    
    If m_KorpaIzlazCount = 0 Then
        MsgBox "Korpa je prazna!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If cmbKooperant.value = "" Then
        MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim koopID As String
    koopID = ExtractIDFromDisplay(cmbKooperant.value)
    
    ' BrojDokumenta generieren
    If txtBrojDokIzlaz.value = "" Then
        MsgBox "Unesite broj dokumenta!", vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim brojDok As String
    brojDok = txtBrojDokIzlaz.value
    
    ' TX starten
    Dim tx As New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_MAGACIN
    
    Dim i As Long
    For i = 1 To m_KorpaIzlazCount
        Dim result As String
        result = SaveMagacin(Date, m_KorpaIzlaz(i).artikalID, MAG_IZLAZ, _
                              m_KorpaIzlaz(i).kolicina, koopID, _
                              m_KorpaIzlaz(i).parcelaID, brojDok)
        If result = "" Then
            tx.RollbackTx
            MsgBox "Greska pri cuvanju!", vbCritical, APP_NAME
            Exit Sub
        End If
    Next i
    
    tx.CommitTx
    
    MsgBox "Izdavanje zavrseno: " & brojDok & vbCrLf & _
           m_KorpaIzlazCount & " stavki", vbInformation, APP_NAME
    
    ' Reset
    ClearKorpaIzlaz
    cmbKooperant.value = ""
    txtBrojDokIzlaz.value = ""
    
    Exit Sub
EH:
    LogErr "frmAgrohemija.btnZavrsiIzlaz"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub ClearKorpaIzlaz()
    Erase m_KorpaIzlaz
    m_KorpaIzlazCount = 0
    lstKorpa.Clear
End Sub

'==========================================
' ARTIKLI ULAZ
'==========================================
Private Sub LoadArtikliUlaz()
    cmbArtikalUlaz.Clear
    Dim data As Variant
    data = GetTableData(TBL_ARTIKLI)
    If IsEmpty(data) Then Exit Sub
    
    Dim colID As Long, colNaziv As Long, colJM As Long, colDoza As Long, colCena As Long
    colID = GetColumnIndex(TBL_ARTIKLI, COL_ART_ID)
    colNaziv = GetColumnIndex(TBL_ARTIKLI, COL_ART_NAZIV)
    colJM = GetColumnIndex(TBL_ARTIKLI, COL_ART_JM)
    colDoza = GetColumnIndex(TBL_ARTIKLI, COL_ART_DOZA)
    colCena = GetColumnIndex(TBL_ARTIKLI, COL_ART_CENA)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 1)) <> "" Then
            cmbArtikalUlaz.AddItem CStr(data(i, colNaziv)) & " [" & _
                CStr(data(i, colJM)) & "] (" & CStr(data(i, colID)) & ")"
        End If
    Next i
End Sub

Private Sub cmbArtikalUlaz_Change()
    lblUlazDoza.caption = ""
    txtCenaUlaz.value = ""
    lblUlazVrednost.caption = ""
    
    If cmbArtikalUlaz.value = "" Then Exit Sub
    
    Dim artID As String
    artID = ExtractIDFromDisplay(cmbArtikalUlaz.value)
    
    ' Doza anzeigen
    Dim dozaStr As String
    dozaStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artID, COL_ART_DOZA))
    Dim jm As String
    jm = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artID, COL_ART_JM))
    lblUlazDoza.caption = "Doza: " & dozaStr & " " & jm & "/ha"
    
    ' Cena vorausfüllen
    Dim cenaStr As String
    cenaStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artID, COL_ART_CENA))
    If IsNumeric(cenaStr) Then txtCenaUlaz.value = cenaStr
End Sub

Private Sub txtKolicinaUlaz_Change()
    UpdateUlazVrednost
End Sub

Private Sub txtCenaUlaz_Change()
    UpdateUlazVrednost
End Sub

Private Sub UpdateUlazVrednost()
    lblUlazVrednost.caption = ""
    If Not IsNumeric(txtKolicinaUlaz.value) Then Exit Sub
    If Not IsNumeric(txtCenaUlaz.value) Then Exit Sub
    
    Dim vrednost As Double
    vrednost = CDbl(txtKolicinaUlaz.value) * CDbl(txtCenaUlaz.value)
    lblUlazVrednost.caption = "Vrednost: " & Format$(vrednost, "#,##0") & " RSD"
End Sub

Private Sub btnDodajUlaz_Click()
    On Error GoTo EH
    
    If cmbArtikalUlaz.value = "" Then
        MsgBox "Izaberite artikal!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not IsNumeric(txtKolicinaUlaz.value) Or CDbl(txtKolicinaUlaz.value) <= 0 Then
        MsgBox "Unesite validnu kolicinu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not IsNumeric(txtCenaUlaz.value) Or CDbl(txtCenaUlaz.value) <= 0 Then
        MsgBox "Unesite validnu cenu!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim artikalID As String
    artikalID = ExtractIDFromDisplay(cmbArtikalUlaz.value)
    
    Dim artNaziv As String
    artNaziv = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_NAZIV))
    Dim jm As String
    jm = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_JM))
    
    Dim kol As Double: kol = CDbl(txtKolicinaUlaz.value)
    Dim cena As Double: cena = CDbl(txtCenaUlaz.value)
    
    m_KorpaUlazCount = m_KorpaUlazCount + 1
    ReDim Preserve m_KorpaUlaz(1 To m_KorpaUlazCount)
    With m_KorpaUlaz(m_KorpaUlazCount)
        .artikalID = artikalID
        .ArtikalNaziv = artNaziv
        .kolicina = kol
        .cena = cena
        .vrednost = kol * cena
        .parcelaID = ""
        .jm = jm
    End With
    
    lstKorpaUlaz.AddItem artNaziv & " - " & Format$(kol, "0.00") & " " & jm & _
                         " x " & Format$(cena, "#,##0") & " = " & Format$(kol * cena, "#,##0") & " RSD"
    
    cmbArtikalUlaz.value = ""
    txtKolicinaUlaz.value = ""
    txtCenaUlaz.value = ""
    lblUlazVrednost.caption = ""
    lblUlazDoza.caption = ""
    
    Exit Sub
EH:
    LogErr "frmAgrohemija.btnDodajUlaz"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub btnZavrsiUlaz_Click()
    On Error GoTo EH
    
    If m_KorpaUlazCount = 0 Then
        MsgBox "Korpa je prazna!", vbExclamation, APP_NAME
        Exit Sub
    End If
    If txtBrojDokUlaz.value = "" Then
        MsgBox "Unesite broj dokumenta!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim brojDok As String
    brojDok = txtBrojDokUlaz.value
    
    Dim dobavljacID As String
    dobavljacID = cmbDobavljac.value
    
    Dim tx As New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_MAGACIN
    
    Dim i As Long
    For i = 1 To m_KorpaUlazCount
        Dim result As String
        result = SaveMagacin(Date, m_KorpaUlaz(i).artikalID, MAG_ULAZ, _
                              m_KorpaUlaz(i).kolicina, "", "", brojDok, "", _
                              dobavljacID, m_KorpaUlaz(i).cena)
        If result = "" Then
            tx.RollbackTx
            MsgBox "Greska pri cuvanju!", vbCritical, APP_NAME
            Exit Sub
        End If
    Next i
    
    tx.CommitTx
    
    MsgBox "Prijem zavrsen: " & brojDok & vbCrLf & _
           m_KorpaUlazCount & " stavki", vbInformation, APP_NAME
    
    ClearKorpaUlaz
    txtBrojDokUlaz.value = ""
    
    Exit Sub
EH:
    LogErr "frmAgrohemija.btnZavrsiUlaz"
    MsgBox "Greska: " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub ClearKorpaUlaz()
    Erase m_KorpaUlaz
    m_KorpaUlazCount = 0
    lstKorpaUlaz.Clear
End Sub

Private Sub btnPovratak_Click()
    Unload Me
    frmOtkupAPP.Show
End Sub

Private Function FormatKol(ByVal val As Double) As String
    If val = Int(val) Then
        FormatKol = CStr(CLng(val))
    Else
        FormatKol = Format$(val, "0.##")
    End If
End Function

