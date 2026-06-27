VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBankaImport 
   Caption         =   "UserForm1"
   ClientHeight    =   13590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "frmBankaImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBankaImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_Data As Variant
Private m_BimIDs() As String
Private mChromeRemoved As Boolean

Private Sub UserForm_Activate()
    On Error GoTo EH
    
    EnsureUserFormChromeRemoved Me, mChromeRemoved
    
    ApplyTheme Me, BG_MAIN()
    ApplyThemeToControls Me
    
    ' Headers
    StyleFrameTitleLabel lblKopf, "Bank Import"
    StyleSubtitle lblSubtitle, Poruka("BANKA_LBL_UVOZ_TRANSAKCIJA_BANKARSKIH")
    
    ' Section headers (ako koristis frames)
    ' StyleSectionHeader fraDetail, "Detalji selektovane stavke"
    ' StyleSectionHeader fraPreview, "Pregled automatskog mapiranja"
    
    ' Action buttons
    StylePrimaryButton btnAutoJedan, "Automatski mapiraj red"
    StylePrimaryButton btnAutoSve, "Automatski mapiraj sve"
    StylePrimaryButton btnSacuvajRucno, "Rucno mapiraj red"
    StylePrimaryButton btnSkip, "Preskoci red"
    StylePrimaryButton btnOsvezi, "Osvezi"
    StyleExitButton btnPovratak, "Zatvori"     ' ili "Povratak"
    
    ' Status labels
    StyleLabel lblStatus, TXT_MUTED(), True
    StyleLabel lblIzvodSummary, TXT_MUTED(), True
    
    ' Detail labels (suptilno)
    StyleLabel lblBimID, TXT_MUTED(), False
    StyleLabel lblPartner, TXT_MUTED(), False
    StyleLabel lblPozivNaBroj, TXT_MUTED(), False
    StyleLabel lblOpis, TXT_MUTED(), False
    StyleLabel lblSvrha, TXT_MUTED(), False
    StyleLabel lblIznos, TXT_MUTED(), True       ' bold jer je money
    StyleLabel lblPreview, TXT_MUTED(), False
    
    ' Initialize cmbMapTip
    If cmbMapTip.ListCount = 0 Then
        cmbMapTip.AddItem "Kupac"
        cmbMapTip.AddItem "Kooperant"
        cmbMapTip.AddItem "OM"
    End If
    
    SetupList
    LoadBankaRows
    
    ' KPI strip (opciono — vidi Izmena 2)
     LayoutTopKpis
     RefreshTopKpis
    
    lstBanka.SetFocus
    
    Exit Sub
    
EH:
    LogErr "frmBankaImport.UserForm_Activate"
    MsgBox "Greska pri otvaranju forme: " & Err.description, vbCritical, APP_NAME
End Sub

Private Sub SetupList()
    With lstBanka
        .ColumnCount = 7
        .ColumnWidths = "70;70;140;80;70;70;60"
    End With
End Sub

Private Sub LoadBankaRows()
    Dim i As Long
    Dim colID As Long, colDatum As Long, colPartner As Long
    Dim colPoziv As Long, colUplata As Long, colIsplata As Long, colObr As Long
    
    lstBanka.Clear
    Erase m_BimIDs
    
    m_Data = GetBankaImportOpen()
    If IsEmpty(m_Data) Then
        lblStatus.caption = "Nema otvorenih stavki."
        UpdateIzvodSummaryLabel
        Exit Sub
    End If
    
    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    colDatum = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE)
    colPartner = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER)
    colPoziv = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POZIV_NA_BROJ)
    colUplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
    colIsplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)
    colObr = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO)
    
    ReDim m_BimIDs(0 To UBound(m_Data, 1) - 1)
    
    For i = 1 To UBound(m_Data, 1)
        lstBanka.AddItem CStr(m_Data(i, colID))
        lstBanka.List(lstBanka.ListCount - 1, 1) = Format$(m_Data(i, colDatum), "d.m.yyyy")
        lstBanka.List(lstBanka.ListCount - 1, 2) = CStr(m_Data(i, colPartner))
        lstBanka.List(lstBanka.ListCount - 1, 3) = CStr(m_Data(i, colPoziv))
        lstBanka.List(lstBanka.ListCount - 1, 4) = Format$(CDbl(nz(m_Data(i, colUplata), "0")), "#,##0.00")
        lstBanka.List(lstBanka.ListCount - 1, 5) = Format$(CDbl(nz(m_Data(i, colIsplata), "0")), "#,##0.00")
        lstBanka.List(lstBanka.ListCount - 1, 6) = CStr(nz(m_Data(i, colObr), ""))
        
        m_BimIDs(i - 1) = CStr(m_Data(i, colID))
    Next i
    
    lblStatus.caption = lstBanka.ListCount & " otvorenih stavki"
    
    UpdateIzvodSummaryLabel
    RefreshTopKpis
    
End Sub

Private Sub lstBanka_Click()
    If lstBanka.ListIndex < 0 Then Exit Sub
    ShowSelectedRow
    UpdateAutoPreview
End Sub

Private Sub ShowSelectedRow()
    Dim bimID As String
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    
    lblBimID.caption = bimID
    lblPartner.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_PARTNER))
    lblPozivNaBroj.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_POZIV_NA_BROJ))
    lblOpis.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_OPIS))
    lblSvrha.caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_SVRHA_PLACANJA))
    
    Dim uplata As Double, isplata As Double
    uplata = CDbl(nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_UPLATA), "0"))
    isplata = CDbl(nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_ISPLATA), "0"))
    
    lblIznos.caption = "Uplata: " & Format$(uplata, "#,##0.00") & _
                       " | Isplata: " & Format$(isplata, "#,##0.00")
    
    If uplata > 0 Then
        cmbMapTip.value = "Kupac"
    ElseIf isplata > 0 Then
        cmbMapTip.value = "Kooperant"
    End If
    
    LoadManualTargets
End Sub

Private Sub cmbMapTip_Change()
    LoadManualTargets
    UpdateAutoPreview
End Sub

Private Sub LoadManualTargets()
    cmbPartner.Clear
    cmbFaktura.Clear
    
    Select Case cmbMapTip.value
        Case "Kupac"
            FillCmb cmbPartner, GetLookupList(TBL_KUPCI, "Naziv")
            
        Case "Kooperant"
            Dim data As Variant
            Dim i As Long
            Dim colID As Long, colIme As Long, colPrezime As Long
            
            data = GetTableData(TBL_KOOPERANTI)
            If IsEmpty(data) Then Exit Sub
            
            colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
            colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
            colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
            
            For i = 1 To UBound(data, 1)
                cmbPartner.AddItem CStr(data(i, colID)) & " - " & _
                                   CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime))
            Next i
            
        Case "OM"
            FillCmb cmbPartner, GetLookupList(TBL_STANICE, "Naziv")
    End Select
End Sub
Private Sub cmbPartner_Change()
    If cmbMapTip.value = "Kooperant" Then
        LoadOtkupBlokoviForSelectedKooperant
    End If
    UpdateAutoPreview
End Sub
Private Sub cmbOtkupBlok_Change()
    UpdateAutoPreview
End Sub
Private Sub LoadOtkupBlokoviForSelectedKooperant()
    Dim kooperantID As String
    Dim data As Variant
    Dim colKoop As Long, colBrDok As Long
    Dim dict As Object
    Dim i As Long
    
    cmbOtkupBlok.Clear
    
    If cmbPartner.value = "" Then Exit Sub
    
    kooperantID = ExtractIDFromDisplay(cmbPartner.value)
    
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub
    
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Sub
    
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID Then
            If Trim$(CStr(data(i, colBrDok))) <> "" Then
                If Not dict.Exists(CStr(data(i, colBrDok))) Then
                    dict.Add CStr(data(i, colBrDok)), True
                End If
            End If
        End If
    Next i
    
    Dim k As Variant
    For Each k In dict.keys
        cmbOtkupBlok.AddItem CStr(k)
    Next k
End Sub

Private Sub btnAutoJedan_Click()
    Dim bimID As String
    Dim result As String
    
    If lstBanka.ListIndex < 0 Then
        MsgBox "Izaberite stavku!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    result = AutoMapBankaImportRow_TX(bimID)
    
    If result <> "" Then
        MsgBox "Automatski mapirano.", vbInformation, APP_NAME
    End If
    
    LoadBankaRows
End Sub

Private Sub btnAutoSve_Click()
    Dim n As Long
    n = AutoMapAllBankaImport_TX()
    MsgBox "Automatski mapirano: " & n, vbInformation, APP_NAME
    LoadBankaRows
End Sub

Private Sub btnSacuvajRucno_Click()
    Dim bimID As String
    
    If lstBanka.ListIndex < 0 Then
        MsgBox "Izaberite stavku!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    
    Select Case cmbMapTip.value
        Case "Kupac"
            Dim kupacID As String
            kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbPartner.value, "KupacID"))
            Call MapBankaImportAsKupac_TX(bimID, kupacID, "", True)
            
    Case "Kooperant"
        Dim kooperantID As String
        Dim brojBloka As String
        Dim n As Long
    
        kooperantID = ExtractIDFromDisplay(cmbPartner.value)
        brojBloka = Trim$(cmbOtkupBlok.value)
    
        If brojBloka <> "" Then
            n = MapBankaImportAsKooperantBlockManual_TX(bimID, kooperantID, brojBloka, True)
        Else
            n = MapBankaImportAsKooperantBlock_TX(bimID, kooperantID, True)
        End If
            
        Case "OM"
            Dim omID As String
            omID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbPartner.value, "StanicaID"))
            Call MapBankaImportAsOM_TX(bimID, omID, "", True)
    End Select
    
    LoadBankaRows
End Sub

Private Sub btnSkip_Click()
    Dim bimID As String
    
    If lstBanka.ListIndex < 0 Then
        MsgBox "Izaberite stavku!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    
    If SkipBankaImportRow_TX(bimID) Then
        LoadBankaRows
    End If
End Sub

Private Sub btnOsvezi_Click()
    LoadBankaRows
End Sub

Private Sub btnPovratak_Click()
    On Error GoTo EH

    frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    Unload Me

    Exit Sub

EH:
    LogErr "frmBankaImport.btnPovratak_Click"
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next

    If CloseMode = vbFormControlMenu Then
        frmOtkupAPP.ReturnToDashboard "Sekcija zatvorena."
    End If

    On Error GoTo 0
End Sub

'PREVIEW
Private Sub UpdateAutoPreview()
    Dim bimID As String
    
    lblPreview.caption = ""
    
    If lstBanka.ListIndex < 0 Then Exit Sub
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    lblPreview.caption = BuildAutoPreviewText(bimID)
End Sub

Private Function BuildAutoPreviewText(ByVal bankaImportID As String) As String
    Dim bim As Variant
    Dim partnerName As String
    Dim uplata As Double
    Dim isplata As Double
    Dim mapped As Variant
    Dim s As String
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        BuildAutoPreviewText = "Preview nije dostupan."
        Exit Function
    End If
    
    partnerName = CStr(bim(1, 3))
    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    isplata = CDbl(NzBIM(bim(1, 6), 0#))
    
    s = "BIM ID: " & bankaImportID & vbCrLf
    s = s & "Partner: " & partnerName & vbCrLf
    
    If Trim$(CStr(bim(1, 10))) <> "" Then
        s = s & "Poziv na broj: " & CStr(bim(1, 10)) & vbCrLf
    End If
    
    If uplata > 0 And isplata = 0 Then
        s = s & BuildIncomingPreview(bankaImportID, partnerName)
    ElseIf isplata > 0 And uplata = 0 Then
        s = s & BuildOutgoingPreview(bankaImportID, partnerName)
    Else
        s = s & "Status: Nije cist smer uplata/isplata"
    End If
    
    BuildAutoPreviewText = s
End Function

Private Function BuildIncomingPreview(ByVal bankaImportID As String, ByVal partnerName As String) As String
    Dim mapped As Variant
    Dim kupacID As String
    Dim fakturaID As String
    Dim s As String
    
    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        If CStr(mapped(1)) = "Kupac" Then
            kupacID = CStr(mapped(0))
            fakturaID = TryResolveFakturaForKupac(bankaImportID, kupacID)
            
            s = "Smer: Uplata" & vbCrLf
            s = s & "Auto match: Kupac" & vbCrLf
            s = s & "KupacID: " & kupacID & vbCrLf
            s = s & "Kupac: " & CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")) & vbCrLf
            
            If fakturaID <> "" Then
                s = s & "FakturaID: " & fakturaID & vbCrLf
                s = s & "Broj fakture: " & CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_BROJ)) & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KUPCI_UPLATA
            Else
                s = s & "Faktura: nije jednoznacno nadjena" & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KUPCI_AVANS
            End If
            
            BuildIncomingPreview = s
            Exit Function
        End If
    End If
    
    mapped = TryResolveKupacBIM(partnerName)
    If Not IsEmpty(mapped) Then
        kupacID = CStr(mapped(0))
        fakturaID = TryResolveFakturaForKupac(bankaImportID, kupacID)
        
        s = "Smer: Uplata" & vbCrLf
        s = s & "Auto match: Kupac (heuristika)" & vbCrLf
        s = s & "KupacID: " & kupacID & vbCrLf
        s = s & "Kupac: " & CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv")) & vbCrLf
        
        If fakturaID <> "" Then
            s = s & "FakturaID: " & fakturaID & vbCrLf
            s = s & "Broj fakture: " & CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_BROJ)) & vbCrLf
            s = s & "Tip knjizenja: " & NOV_KUPCI_UPLATA
        Else
            s = s & "Faktura: nije jednoznacno nadjena" & vbCrLf
            s = s & "Tip knjizenja: " & NOV_KUPCI_AVANS
        End If
        
        BuildIncomingPreview = s
        Exit Function
    End If
    
    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        s = "Smer: Uplata" & vbCrLf
        s = s & "Auto match: OM" & vbCrLf
        s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
        s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
        s = s & "Tip knjizenja: " & NOV_KES_FIRMA_OTKUPAC
        BuildIncomingPreview = s
        Exit Function
    End If
    
    BuildIncomingPreview = "Smer: Uplata" & vbCrLf & "Auto match: Nije pronadjen"
End Function

Private Function BuildOutgoingPreview(ByVal bankaImportID As String, ByVal partnerName As String) As String
    Dim mapped As Variant
    Dim kooperantID As String
    Dim kandidati As Variant
    Dim s As String
    Dim blockNo As String
    Dim i As Long
    
    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        Select Case CStr(mapped(1))
            Case "Kooperant"
                kooperantID = CStr(mapped(0))
                blockNo = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ))
                kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blockNo)
                
                s = "Smer: Isplata" & vbCrLf
                s = s & "Auto match: Kooperant" & vbCrLf
                s = s & "KooperantID: " & kooperantID & vbCrLf
                s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf
                
                If Trim$(blockNo) <> "" Then
                    s = s & "Blok: " & blockNo & vbCrLf
                End If
                
                If IsEmpty(kandidati) Then
                    s = s & "Otkup kandidati: nema otvorenih stavki" & vbCrLf
                    s = s & "Tip knjizenja: " & NOV_VIRMAN_AVANS_KOOP
                Else
                    s = s & "Otkup kandidati:" & vbCrLf
                    For i = 1 To UBound(kandidati, 1)
                        s = s & " - " & CStr(kandidati(i, 1)) & " | otvoreno: " & _
                            Format$(CDbl(kandidati(i, 2)), "#,##0.00") & " | " & _
                            CStr(kandidati(i, 3)) & vbCrLf
                    Next i
                    s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_KOOP
                End If
                
                BuildOutgoingPreview = s
                Exit Function
                
            Case "OM"
                s = "Smer: Isplata" & vbCrLf
                s = s & "Auto match: OM" & vbCrLf
                s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
                s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
                s = s & "Tip knjizenja: " & NOV_KES_FIRMA_OTKUPAC
                BuildOutgoingPreview = s
                Exit Function
        End Select
    End If
    
    mapped = TryResolveKooperantBIM(partnerName)
    If Not IsEmpty(mapped) Then
        kooperantID = CStr(mapped(0))
        blockNo = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ))
        kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blockNo)
        
        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: Kooperant (heuristika)" & vbCrLf
        s = s & "KooperantID: " & kooperantID & vbCrLf
        s = s & "Kooperant: " & GetKooperantNaziv(kooperantID) & vbCrLf
        
        If Trim$(blockNo) <> "" Then
            s = s & "Blok: " & blockNo & vbCrLf
        End If
        
        If IsEmpty(kandidati) Then
            s = s & "Otkup kandidati: nema otvorenih stavki" & vbCrLf
            s = s & "Tip knjizenja: " & NOV_VIRMAN_AVANS_KOOP
        Else
            s = s & "Otkup kandidati:" & vbCrLf
            For i = 1 To UBound(kandidati, 1)
                s = s & " - " & CStr(kandidati(i, 1)) & " | otvoreno: " & _
                    Format$(CDbl(kandidati(i, 2)), "#,##0.00") & " | " & _
                    CStr(kandidati(i, 3)) & vbCrLf
            Next i
            s = s & "Tip knjizenja: " & NOV_VIRMAN_FIRMA_KOOP
        End If
        
        BuildOutgoingPreview = s
        Exit Function
    End If
    
    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: OM" & vbCrLf
        s = s & "OMID: " & CStr(mapped(0)) & vbCrLf
        s = s & "OM: " & CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(mapped(0)), "Naziv")) & vbCrLf
        s = s & "Tip knjizenja: " & NOV_KES_FIRMA_OTKUPAC
        BuildOutgoingPreview = s
        Exit Function
    End If
    
    BuildOutgoingPreview = "Smer: Isplata" & vbCrLf & "Auto match: Nije pronadjen"
End Function

'======================================================================
' v6.18+: Statement Summary Display
'
' Prikazuje saldo info iz najnovijeg aktivnog (non-stornirano) izvoda
' u tblBankaImport. Citaje BrojIzvoda + DatumIzvoda i 4 saldo polja,
' pa prikazuje math integrity status (Level 1) read-side kao vizualnu
' potvrdu operateru.
'
' Pozvana iz LoadBankaRows() na kraju, sto pokriva sve refresh tacke
' (Activate, btnOsvezi_Click, i posle svake Auto*/Sacuvaj/Skip akcije).
'======================================================================
Private Sub UpdateIzvodSummaryLabel()
    On Error GoTo EH
    
    Dim data As Variant
    Dim colBrojIzvoda As Long
    Dim colDatumIzvoda As Long
    Dim colPocetno As Long
    Dim colZavrsno As Long
    Dim colDuguje As Long
    Dim colPotrazuje As Long
    
    Dim i As Long
    Dim maxRow As Long
    Dim maxDatum As Date
    
    data = GetTableData(TBL_BANKA_IMPORT)
    
    If IsEmpty(data) Then
        Me.lblIzvodSummary.caption = "Nema importovanih izvoda."
        Exit Sub
    End If
    
    If UBound(data, 1) < 1 Then
        Me.lblIzvodSummary.caption = "Nema importovanih izvoda."
        Exit Sub
    End If
    
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    
    If IsEmpty(data) Then
        Me.lblIzvodSummary.caption = "Nema aktivnih izvoda."
        Exit Sub
    End If
    
    colBrojIzvoda = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA)
    colDatumIzvoda = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_IZVODA)
    colPocetno = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POCETNO_STANJE)
    colZavrsno = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ZAVRSNO_STANJE)
    colDuguje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_DUGUJE)
    colPotrazuje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_POTRAZUJE)
    
    maxRow = -1
    maxDatum = #1/1/1900#
    
    For i = 1 To UBound(data, 1)
        Dim rowDatumStr As String
        rowDatumStr = Trim$(CStr(data(i, colDatumIzvoda)))
        
        If LenB(rowDatumStr) > 0 Then
            Dim rowDatum As Date
            On Error Resume Next
            rowDatum = CDate(rowDatumStr)
            If Err.Number = 0 Then
                On Error GoTo EH
                If rowDatum >= maxDatum Then
                    maxDatum = rowDatum
                    maxRow = i
                End If
            Else
                Err.Clear
                On Error GoTo EH
            End If
        End If
    Next i
    
    If maxRow < 0 Then
        Me.lblIzvodSummary.caption = "Nema aktivnih izvoda sa validnim datumom."
        Exit Sub
    End If
    
    Dim brIzv As String
    Dim datumIzv As String
    Dim pocetno As Double
    Dim zavrsno As Double
    Dim duguje As Double
    Dim potraz As Double
    
    brIzv = CStr(data(maxRow, colBrojIzvoda))
    datumIzv = CStr(data(maxRow, colDatumIzvoda))
    pocetno = CDbl(nz(data(maxRow, colPocetno), "0"))
    zavrsno = CDbl(nz(data(maxRow, colZavrsno), "0"))
    duguje = CDbl(nz(data(maxRow, colDuguje), "0"))
    potraz = CDbl(nz(data(maxRow, colPotrazuje), "0"))
    
    ' Level 1 read-side math check
    Dim expected As Double
    Dim diff As Double
    expected = pocetno + potraz - duguje
    diff = Abs(expected - zavrsno)
    
    Dim statusTxt As String
    Dim statusColor As Long
    
    If pocetno = 0 And zavrsno = 0 And duguje = 0 And potraz = 0 Then
        ' Legacy row (pre-v6.18 import, bez saldo metapodataka)
        statusTxt = "(legacy - bez saldo metapodataka)"
        statusColor = RGB(128, 128, 128)
    ElseIf diff <= 0.01 Then
        statusTxt = "OK"
        statusColor = RGB(0, 128, 0)
    Else
        statusTxt = "DIFF " & Format$(diff, "#,##0.00")
        statusColor = RGB(192, 0, 0)
    End If
    
    Me.lblIzvodSummary.caption = _
        "Izvod " & brIzv & " (" & datumIzv & ")  |  " & _
        "Pocetno: " & Format$(pocetno, "#,##0.00") & "  |  " & _
        "Uplate: " & Format$(potraz, "#,##0.00") & "  |  " & _
        "Isplate: " & Format$(duguje, "#,##0.00") & "  |  " & _
        "Zavrsno: " & Format$(zavrsno, "#,##0.00") & "  |  " & _
        statusTxt
    Me.lblIzvodSummary.ForeColor = statusColor
    
    Exit Sub

EH:
    On Error Resume Next
    Me.lblIzvodSummary.caption = "(greska pri citanju saldo info-a)"
    Me.lblIzvodSummary.ForeColor = RGB(128, 128, 128)
    
    LogErr "frmBankaImport.UpdateIzvodSummaryLabel"
End Sub

Private Sub LayoutTopKpis()
    On Error GoTo EH
    
    LayoutTopKpiInternals fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent
    LayoutTopKpiInternals fraKpiAutoMatch, lblKpiAutoTitle, lblKpiAutoValue, lblKpiAutoAccent
    LayoutTopKpiInternals fraKpiUplate, lblKpiUplTitle, lblKpiUplValue, lblKpiUplAccent
    LayoutTopKpiInternals fraKpiIsplate, lblKpiIspTitle, lblKpiIspValue, lblKpiIspAccent
    
    Exit Sub
EH:
    LogErr "frmBankaImport.LayoutTopKpis"
End Sub

Private Sub RefreshTopKpis()
    On Error GoTo EH
    
    Dim totalCount As Long
    Dim totalUplata As Double
    Dim totalIsplata As Double
    Dim autoMatchCount As Long
    
    If IsArray(m_Data) Then
        totalCount = UBound(m_Data, 1)
        
        Dim colUplata As Long, colIsplata As Long, colPartner As Long
        colUplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
        colIsplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)
        colPartner = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER)
        
        Dim i As Long
        For i = 1 To UBound(m_Data, 1)
            totalUplata = totalUplata + CDbl(nz(m_Data(i, colUplata), "0"))
            totalIsplata = totalIsplata + CDbl(nz(m_Data(i, colIsplata), "0"))
            
            ' Check if auto-match exists
            Dim partnerName As String
            partnerName = CStr(m_Data(i, colPartner))
            Dim mapped As Variant
            mapped = LookupPartnerMap(partnerName)
            If Not IsEmpty(mapped) Then
                autoMatchCount = autoMatchCount + 1
            End If
        Next i
    End If
    
    ' Card 1: Otvoreno
    StyleTopKpi fraKpiOtvoreno, lblKpiOtvTitle, lblKpiOtvValue, lblKpiOtvAccent, "neutral"
    lblKpiOtvTitle.caption = "Otvoreno"
    lblKpiOtvValue.caption = totalCount & " stavki"
    
    ' Card 2: Auto match (ok ako >0)
    Dim autoKind As String
    If autoMatchCount > 0 Then autoKind = "ok" Else autoKind = "neutral"
    StyleTopKpi fraKpiAutoMatch, lblKpiAutoTitle, lblKpiAutoValue, lblKpiAutoAccent, autoKind
    lblKpiAutoTitle.caption = "Auto match"
    lblKpiAutoValue.caption = autoMatchCount & " / " & totalCount
    
    ' Card 3: Uplate ukupno
    StyleTopKpi fraKpiUplate, lblKpiUplTitle, lblKpiUplValue, lblKpiUplAccent, "neutral"
    lblKpiUplTitle.caption = "Uplate"
    lblKpiUplValue.caption = Format$(totalUplata, "#,##0") & " RSD"
    
    ' Card 4: Isplate ukupno
    StyleTopKpi fraKpiIsplate, lblKpiIspTitle, lblKpiIspValue, lblKpiIspAccent, "neutral"
    lblKpiIspTitle.caption = "Isplate"
    lblKpiIspValue.caption = Format$(totalIsplata, "#,##0") & " RSD"
    
    Exit Sub
EH:
    LogErr "frmBankaImport.RefreshTopKpis"
End Sub

Private Sub ResetActionButtons()
    StylePrimaryButton btnAutoJedan, "Automatski mapiraj red"
    StylePrimaryButton btnAutoSve, "Automatski mapiraj sve"
    StylePrimaryButton btnSacuvajRucno, "Rucno mapiraj red"
    StylePrimaryButton btnSkip, "Preskoci red"
    StylePrimaryButton btnOsvezi, "Osvezi"
    StyleExitButton btnPovratak, "Zatvori"
End Sub

Private Sub btnAutoJedan_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnAutoJedan
End Sub

Private Sub btnAutoSve_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnAutoSve
End Sub

Private Sub btnSacuvajRucno_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnSacuvajRucno
End Sub

Private Sub btnSkip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnSkip
End Sub

Private Sub btnOsvezi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnOsvezi
End Sub

Private Sub btnPovratak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
    ButtonHover btnPovratak
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ResetActionButtons
End Sub
