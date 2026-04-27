VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBankaImport 
   Caption         =   "UserForm1"
   ClientHeight    =   13590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21855
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

Private Sub UserForm_Activate()
    Me.caption = "Banka import"
    
    If cmbMapTip.ListCount = 0 Then
        cmbMapTip.AddItem "Kupac"
        cmbMapTip.AddItem "Kooperant"
        cmbMapTip.AddItem "OM"
    End If
    
    SetupList
    LoadBankaRows
    lstBanka.SetFocus
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
        lstBanka.List(lstBanka.ListCount - 1, 4) = Format$(CDbl(Nz(m_Data(i, colUplata), "0")), "#,##0.00")
        lstBanka.List(lstBanka.ListCount - 1, 5) = Format$(CDbl(Nz(m_Data(i, colIsplata), "0")), "#,##0.00")
        lstBanka.List(lstBanka.ListCount - 1, 6) = CStr(Nz(m_Data(i, colObr), ""))
        
        m_BimIDs(i - 1) = CStr(m_Data(i, colID))
    Next i
    
    lblStatus.caption = lstBanka.ListCount & " otvorenih stavki"
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
    uplata = CDbl(Nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_UPLATA), "0"))
    isplata = CDbl(Nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_ISPLATA), "0"))
    
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

Private Sub btnZatvori_Click()
    Unload Me
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

