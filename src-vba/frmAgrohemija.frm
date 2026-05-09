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

Private Sub UserForm_Initialize()
    ApplyTheme Me, BG_MAIN()
    LoadKooperanti
    LoadArtikli
    LoadArtikliUlaz
End Sub

Private Sub UserForm_Activate()
    EnsureUserFormChromeRemoved Me, mChromeRemoved
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
    
    Dim trenutnoUKorpi As Double
    trenutnoUKorpi = GetKorpaIzlazKolicinaZaArtikal(artikalID)

    Dim stanjeDict As Object
    Set stanjeDict = BuildArtikalStanjeDict()

    Dim dostupno As Double
    dostupno = 0#
    If stanjeDict.Exists(artikalID) Then dostupno = CDbl(stanjeDict(artikalID))

    If trenutnoUKorpi + kol > dostupno Then
        MsgBox "Nedovoljno stanje za artikal!" & vbCrLf & _
            "Na stanju: " & FormatKol(dostupno) & " " & jm & vbCrLf & _
            "Vec u korpi: " & FormatKol(trenutnoUKorpi) & " " & jm & vbCrLf & _
            "Pokušavate dodati: " & FormatKol(kol) & " " & jm, _
            vbExclamation, APP_NAME
        Exit Sub
    End If
    
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
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnZavrsiIzlaz_Click()
    Const SRC As String = "frmAgrohemija.btnZavrsiIzlaz"

    Dim tx As clsTransaction
    Dim txStarted As Boolean

    On Error GoTo EH

    If m_KorpaIzlazCount = 0 Then
        MsgBox "Korpa je prazna!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If cmbKooperant.value = "" Then
        MsgBox "Izaberite kooperanta!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If Trim$(txtBrojDokIzlaz.value) = "" Then
        MsgBox "Unesite broj dokumenta!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim koopID As String
    koopID = ExtractIDFromDisplay(cmbKooperant.value)

    Dim brojDok As String
    brojDok = Trim$(txtBrojDokIzlaz.value)

    ' Pre-TX UX/business check: agregirano po artiklu.
    ' SaveMagacin takode radi finalni check, ali ovde dobijamo bolju poruku
    ' i sprecavamo delimican loop pre rollback-a.
    ValidateKorpaIzlazStanje

    Set tx = New clsTransaction
    tx.BeginTx
    txStarted = True
    tx.AddTableSnapshot TBL_MAGACIN

    Dim i As Long
    For i = 1 To m_KorpaIzlazCount
        Dim result As String

        result = SaveMagacin( _
            Date, _
            m_KorpaIzlaz(i).artikalID, _
            MAG_IZLAZ, _
            m_KorpaIzlaz(i).kolicina, _
            koopID, _
            m_KorpaIzlaz(i).parcelaID, _
            brojDok)

        If Len(Trim$(result)) = 0 Then
            Err.Raise vbObjectError + 4301, SRC, _
                      "Greška pri cuvanju izlaza. ArtikalID=" & _
                      m_KorpaIzlaz(i).artikalID & _
                      "; Kolicina=" & CStr(m_KorpaIzlaz(i).kolicina)
        End If
    Next i

    tx.CommitTx
    txStarted = False

    MsgBox "Izdavanje završeno: " & brojDok & vbCrLf & _
           m_KorpaIzlazCount & " stavki", vbInformation, APP_NAME

    ClearKorpaIzlaz
    cmbKooperant.value = ""
    txtBrojDokIzlaz.value = ""

    Set tx = Nothing
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description

    LogErr SRC

    On Error Resume Next
    If txStarted And Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    MsgBox "Greška pri cuvanju izdavanja, promene vracene: " & errDesc, _
           vbCritical, APP_NAME

    Set tx = Nothing
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
    MsgBox "Greska: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub btnZavrsiUlaz_Click()
    Const SRC As String = "frmAgrohemija.btnZavrsiUlaz"

    Dim tx As clsTransaction
    Dim txStarted As Boolean

    On Error GoTo EH

    If m_KorpaUlazCount = 0 Then
        MsgBox "Korpa je prazna!", vbExclamation, APP_NAME
        Exit Sub
    End If

    If Trim$(txtBrojDokUlaz.value) = "" Then
        MsgBox "Unesite broj dokumenta!", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim brojDok As String
    brojDok = Trim$(txtBrojDokUlaz.value)

    Dim dobavljacID As String
    dobavljacID = Trim$(cmbDobavljac.value)

    Set tx = New clsTransaction
    tx.BeginTx
    txStarted = True
    tx.AddTableSnapshot TBL_MAGACIN

    Dim i As Long
    For i = 1 To m_KorpaUlazCount
        Dim result As String

        result = SaveMagacin( _
            Date, _
            m_KorpaUlaz(i).artikalID, _
            MAG_ULAZ, _
            m_KorpaUlaz(i).kolicina, _
            "", _
            "", _
            brojDok, _
            "", _
            dobavljacID, _
            m_KorpaUlaz(i).cena)

        If Len(Trim$(result)) = 0 Then
            Err.Raise vbObjectError + 4311, SRC, _
                      "Greška pri cuvanju ulaza. ArtikalID=" & _
                      m_KorpaUlaz(i).artikalID & _
                      "; Kolicina=" & CStr(m_KorpaUlaz(i).kolicina)
        End If
    Next i

    tx.CommitTx
    txStarted = False

    MsgBox "Prijem završen: " & brojDok & vbCrLf & _
           m_KorpaUlazCount & " stavki", vbInformation, APP_NAME

    ClearKorpaUlaz
    txtBrojDokUlaz.value = ""

    Set tx = Nothing
    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.description

    LogErr SRC

    On Error Resume Next
    If txStarted And Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    MsgBox "Greška pri cuvanju prijema, promene vracene: " & errDesc, _
           vbCritical, APP_NAME

    Set tx = Nothing
End Sub

Private Sub ClearKorpaUlaz()
    Erase m_KorpaUlaz
    m_KorpaUlazCount = 0
    lstKorpaUlaz.Clear
End Sub

Private Sub btnPovratak_Click()
    On Error GoTo EH

    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me

    Exit Sub

EH:
    LogErr "frmAgrohemija.btnPovratak_Click"
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If

    On Error GoTo 0
End Sub

Private Function FormatKol(ByVal val As Double) As String
    If val = Int(val) Then
        FormatKol = CStr(CLng(val))
    Else
        FormatKol = Format$(val, "0.##")
    End If
End Function

Private Sub ValidateKorpaIzlazStanje()
    Const SRC As String = "frmAgrohemija.ValidateKorpaIzlazStanje"

    If m_KorpaIzlazCount = 0 Then Exit Sub

    Dim stanjeDict As Object
    Set stanjeDict = BuildArtikalStanjeDict()

    Dim needDict As Object
    Set needDict = CreateObject("Scripting.Dictionary")

    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")

    Dim jmDict As Object
    Set jmDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To m_KorpaIzlazCount
        Dim artID As String
        artID = Trim$(m_KorpaIzlaz(i).artikalID)

        If Len(artID) = 0 Then
            Err.Raise vbObjectError + 4320, SRC, "Korpa sadrži stavku bez ArtikalID."
        End If

        If Not needDict.Exists(artID) Then
            needDict.Add artID, 0#
            nameDict.Add artID, m_KorpaIzlaz(i).ArtikalNaziv
            jmDict.Add artID, m_KorpaIzlaz(i).jm
        End If

        needDict(artID) = CDbl(needDict(artID)) + CDbl(m_KorpaIzlaz(i).kolicina)
    Next i

    Dim key As Variant
    For Each key In needDict.keys
        Dim available As Double
        available = 0#

        If stanjeDict.Exists(CStr(key)) Then
            available = CDbl(stanjeDict(CStr(key)))
        End If

        Dim needed As Double
        needed = CDbl(needDict(CStr(key)))

        If available < needed Then
            Err.Raise vbObjectError + 4321, SRC, _
                      "Nedovoljno stanje za artikal: " & _
                      CStr(nameDict(CStr(key))) & " (" & CStr(key) & ")." & vbCrLf & _
                      "Na stanju: " & FormatKol(available) & " " & CStr(jmDict(CStr(key))) & vbCrLf & _
                      "U korpi: " & FormatKol(needed) & " " & CStr(jmDict(CStr(key)))
        End If
    Next key
End Sub

Private Function BuildArtikalStanjeDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim stanje As Variant
    stanje = GetMagacinStanje()

    If IsEmpty(stanje) Then
        Set BuildArtikalStanjeDict = dict
        Exit Function
    End If

    Dim i As Long
    For i = 1 To UBound(stanje, 1)
        Dim artID As String
        artID = Trim$(CStr(stanje(i, 1)))

        If Len(artID) > 0 Then
            If IsNumeric(stanje(i, 7)) Then
                dict(artID) = CDbl(stanje(i, 7))
            Else
                dict(artID) = 0#
            End If
        End If
    Next i

    Set BuildArtikalStanjeDict = dict
End Function

Private Function GetKorpaIzlazKolicinaZaArtikal(ByVal artikalID As String) As Double
    Dim i As Long
    For i = 1 To m_KorpaIzlazCount
        If Trim$(m_KorpaIzlaz(i).artikalID) = Trim$(artikalID) Then
            GetKorpaIzlazKolicinaZaArtikal = _
                GetKorpaIzlazKolicinaZaArtikal + CDbl(m_KorpaIzlaz(i).kolicina)
        End If
    Next i
End Function

