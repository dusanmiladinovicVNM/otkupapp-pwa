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
    Me.Caption = "Banka import"
    
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
        lblStatus.Caption = "Nema otvorenih stavki."
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
    
    lblStatus.Caption = lstBanka.ListCount & " otvorenih stavki"
End Sub

Private Sub lstBanka_Click()
    If lstBanka.ListIndex < 0 Then Exit Sub
    ShowSelectedRow
    UpdateAutoPreview
End Sub

Private Sub ShowSelectedRow()
    Dim bimID As String
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    
    lblBimID.Caption = bimID
    lblPartner.Caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_PARTNER))
    lblPozivNaBroj.Caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_POZIV_NA_BROJ))
    lblOpis.Caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_OPIS))
    lblSvrha.Caption = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_SVRHA_PLACANJA))
    
    Dim uplata As Double, isplata As Double
    uplata = CDbl(Nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_UPLATA), "0"))
    isplata = CDbl(Nz(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bimID, COL_BIM_ISPLATA), "0"))
    
    lblIznos.Caption = "Uplata: " & Format$(uplata, "#,##0.00") & _
                       " | Isplata: " & Format$(isplata, "#,##0.00")
    
    If uplata > 0 Then
        cmbMapTip.Value = "Kupac"
    ElseIf isplata > 0 Then
        cmbMapTip.Value = "Kooperant"
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
    
    Select Case cmbMapTip.Value
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
    If cmbMapTip.Value = "Kooperant" Then
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
    
    If cmbPartner.Value = "" Then Exit Sub
    
    kooperantID = ExtractIDFromDisplay(cmbPartner.Value)
    
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
    
    Select Case cmbMapTip.Value
        Case "Kupac"
            Dim kupacID As String
            kupacID = CStr(LookupValue(TBL_KUPCI, "Naziv", cmbPartner.Value, "KupacID"))
            Call MapBankaImportAsKupac_TX(bimID, kupacID, "", True)
            
    Case "Kooperant"
        Dim kooperantID As String
        Dim brojBloka As String
        Dim n As Long
    
        kooperantID = ExtractIDFromDisplay(cmbPartner.Value)
        brojBloka = Trim$(cmbOtkupBlok.Value)
    
        If brojBloka <> "" Then
            n = MapBankaImportAsKooperantBlockManual_TX(bimID, kooperantID, brojBloka, True)
        Else
            n = MapBankaImportAsKooperantBlock_TX(bimID, kooperantID, True)
        End If
            
        Case "OM"
            Dim omID As String
            omID = CStr(LookupValue(TBL_STANICE, "Naziv", cmbPartner.Value, "StanicaID"))
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
    
    lblPreview.Caption = ""
    
    If lstBanka.ListIndex < 0 Then Exit Sub
    
    bimID = m_BimIDs(lstBanka.ListIndex)
    lblPreview.Caption = BuildAutoPreviewText(bimID)
End Sub

Private Function BuildAutoPreviewText(ByVal bankaImportID As String) As String
    Dim bim As Variant
    Dim partnerName As String
    Dim uplata As Double
    Dim isplata As Double
    Dim mapped As Variant
    Dim s As String
    
    bim = GetBankaImportRowByID_Preview(bankaImportID)
    If IsEmpty(bim) Then
        BuildAutoPreviewText = "Preview nije dostupan."
        Exit Function
    End If
    
    partnerName = CStr(bim(1, 3))
    uplata = CDbl(NzBankaPreview(bim(1, 5), 0#))
    isplata = CDbl(NzBankaPreview(bim(1, 6), 0#))
    
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
            fakturaID = TryResolveFakturaForKupac_Preview(bankaImportID, kupacID)
            
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
    
    mapped = TryResolveKupacBIM_Preview(partnerName)
    If Not IsEmpty(mapped) Then
        kupacID = CStr(mapped(0))
        fakturaID = TryResolveFakturaForKupac_Preview(bankaImportID, kupacID)
        
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
    
    mapped = TryResolveOMBIM_Preview(partnerName)
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
                kandidati = GetOtkupCandidatesForKooperantBlock_Preview(kooperantID, blockNo)
                
                s = "Smer: Isplata" & vbCrLf
                s = s & "Auto match: Kooperant" & vbCrLf
                s = s & "KooperantID: " & kooperantID & vbCrLf
                s = s & "Kooperant: " & GetKooperantNazivPreview(kooperantID) & vbCrLf
                
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
    
    mapped = TryResolveKooperantBIM_Preview(partnerName)
    If Not IsEmpty(mapped) Then
        kooperantID = CStr(mapped(0))
        blockNo = CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ))
        kandidati = GetOtkupCandidatesForKooperantBlock_Preview(kooperantID, blockNo)
        
        s = "Smer: Isplata" & vbCrLf
        s = s & "Auto match: Kooperant (heuristika)" & vbCrLf
        s = s & "KooperantID: " & kooperantID & vbCrLf
        s = s & "Kooperant: " & GetKooperantNazivPreview(kooperantID) & vbCrLf
        
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
    
    mapped = TryResolveOMBIM_Preview(partnerName)
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

Private Function GetBankaImportRowByID_Preview(ByVal bankaImportID As String) As Variant
    Dim data As Variant
    Dim result(1 To 1, 1 To 10) As Variant
    Dim colID As Long
    Dim colBrojDok As Long
    Dim colDatumTx As Long
    Dim colPartner As Long
    Dim colPartnerKonto As Long
    Dim colUplata As Long
    Dim colIsplata As Long
    Dim colOpis As Long
    Dim colSvrha As Long
    Dim colRef As Long
    Dim colPoziv As Long
    Dim i As Long
    
    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Function
    
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    If IsEmpty(data) Then Exit Function
    
    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    colBrojDok = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA)
    colDatumTx = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE)
    colPartner = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER)
    colPartnerKonto = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER_KONTO)
    colUplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
    colIsplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)
    colOpis = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OPIS)
    colSvrha = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_SVRHA_PLACANJA)
    colRef = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BANKA_REFERENZ)
    colPoziv = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POZIV_NA_BROJ)
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colID)) = bankaImportID Then
            result(1, 1) = CStr(data(i, colBrojDok))
            result(1, 2) = data(i, colDatumTx)
            result(1, 3) = CStr(data(i, colPartner))
            result(1, 4) = CStr(data(i, colPartnerKonto))
            result(1, 5) = CDbl(NzBankaPreview(data(i, colUplata), 0#))
            result(1, 6) = CDbl(NzBankaPreview(data(i, colIsplata), 0#))
            result(1, 7) = CStr(data(i, colOpis))
            result(1, 8) = CStr(data(i, colSvrha))
            result(1, 9) = CStr(data(i, colRef))
            result(1, 10) = CStr(data(i, colPoziv))
            GetBankaImportRowByID_Preview = result
            Exit Function
        End If
    Next i
End Function

Private Function TryResolveKupacBIM_Preview(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long, colNaziv As Long
    Dim i As Long, hits As Long
    Dim hitID As String
    
    data = GetTableData(TBL_KUPCI)
    If IsEmpty(data) Then Exit Function
    
    colID = GetColumnIndex(TBL_KUPCI, "KupacID")
    colNaziv = GetColumnIndex(TBL_KUPCI, "Naziv")
    
    For i = 1 To UBound(data, 1)
        If NormalizeLoosePreview(CStr(data(i, colNaziv))) = NormalizeLoosePreview(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i
    
    If hits = 1 Then TryResolveKupacBIM_Preview = Array(hitID, "Kupac", "")
End Function

Private Function TryResolveKooperantBIM_Preview(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long, colIme As Long, colPrezime As Long
    Dim i As Long, hits As Long
    Dim hitID As String, hitOMID As String
    Dim fullName As String
    
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then Exit Function
    
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    
    For i = 1 To UBound(data, 1)
        fullName = NormalizeLoosePreview(CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime)))
        If fullName = NormalizeLoosePreview(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
            hitOMID = CStr(NzBankaPreview(LookupValue(TBL_KOOPERANTI, "KooperantID", hitID, COL_KOOP_STANICA), ""))
        End If
    Next i
    
    If hits = 1 Then TryResolveKooperantBIM_Preview = Array(hitID, "Kooperant", hitOMID)
End Function

Private Function TryResolveOMBIM_Preview(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long, colNaziv As Long
    Dim i As Long, hits As Long
    Dim hitID As String
    
    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then Exit Function
    
    colID = GetColumnIndex(TBL_STANICE, "StanicaID")
    colNaziv = GetColumnIndex(TBL_STANICE, "Naziv")
    
    For i = 1 To UBound(data, 1)
        If NormalizeLoosePreview(CStr(data(i, colNaziv))) = NormalizeLoosePreview(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i
    
    If hits = 1 Then TryResolveOMBIM_Preview = Array(hitID, "OM", hitID)
End Function

Private Function TryResolveFakturaForKupac_Preview(ByVal bankaImportID As String, ByVal kupacID As String) As String
    Dim bim As Variant
    Dim faktData As Variant
    Dim colFID As Long, colBroj As Long, colKup As Long, colIznos As Long
    Dim poziv As String, svrha As String, uplata As Double
    Dim i As Long, hitCount As Long, hitID As String
    
    bim = GetBankaImportRowByID_Preview(bankaImportID)
    If IsEmpty(bim) Then Exit Function
    
    poziv = NormalizeLoosePreview(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    svrha = NormalizeLoosePreview(CStr(bim(1, 8)))
    uplata = CDbl(NzBankaPreview(bim(1, 5), 0#))
    
    faktData = GetTableData(TBL_FAKTURE)
    If IsEmpty(faktData) Then Exit Function
    
    faktData = ExcludeStornirano(faktData, TBL_FAKTURE)
    If IsEmpty(faktData) Then Exit Function
    
    colFID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
    colKup = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)
    colIznos = GetColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS)
    
    For i = 1 To UBound(faktData, 1)
        If CStr(faktData(i, colKup)) <> kupacID Then GoTo NextI
        
        Dim brojFak As String
        brojFak = NormalizeLoosePreview(CStr(faktData(i, colBroj)))
        
        If brojFak <> "" Then
            If poziv <> "" Then
                If InStr(1, poziv, brojFak, vbTextCompare) > 0 Or InStr(1, brojFak, poziv, vbTextCompare) > 0 Then
                    hitCount = hitCount + 1
                    hitID = CStr(faktData(i, colFID))
                    GoTo NextI
                End If
            End If
            
            If svrha <> "" Then
                If InStr(1, svrha, brojFak, vbTextCompare) > 0 Then
                    hitCount = hitCount + 1
                    hitID = CStr(faktData(i, colFID))
                    GoTo NextI
                End If
            End If
        End If
        
        If Abs(CDbl(NzBankaPreview(faktData(i, colIznos), 0#)) - uplata) < 0.01 Then
            hitCount = hitCount + 1
            hitID = CStr(faktData(i, colFID))
        End If
        
NextI:
    Next i
    
    If hitCount = 1 Then TryResolveFakturaForKupac_Preview = hitID
End Function

Private Function GetOtkupCandidatesForKooperantBlock_Preview(ByVal kooperantID As String, _
                                                             ByVal brojBloka As String) As Variant
    Dim data As Variant
    Dim tempResult() As Variant
    Dim finalResult() As Variant
    Dim colOtkID As Long, colKoop As Long, colBrDok As Long
    Dim colKol As Long, colCena As Long, colVrsta As Long
    Dim i As Long, count As Long, r As Long, c As Long
    
    If Trim$(kooperantID) = "" Then Exit Function
    If Trim$(brojBloka) = "" Then Exit Function
    
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then Exit Function
    
    colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    
    ReDim tempResult(1 To 2, 1 To 3)
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) <> kooperantID Then GoTo NextI
        
        If NormalizeLoosePreview(CStr(data(i, colBrDok))) = NormalizeLoosePreview(brojBloka) Then
            Dim vrednost As Double
            Dim uplaceno As Double
            Dim otvoreno As Double
            
            If IsNumeric(data(i, colKol)) And IsNumeric(data(i, colCena)) Then
                vrednost = CDbl(data(i, colKol)) * CDbl(data(i, colCena))
            End If
            
            uplaceno = GetUplataForOtkup(CStr(data(i, colOtkID)))
            otvoreno = vrednost - uplaceno
            
            If otvoreno > 0.009 Then
                count = count + 1
                If count > 2 Then Exit For
                
                tempResult(count, 1) = CStr(data(i, colOtkID))
                tempResult(count, 2) = otvoreno
                tempResult(count, 3) = CStr(data(i, colVrsta))
            End If
        End If
NextI:
    Next i
    
    If count = 0 Then Exit Function
    
    ReDim finalResult(1 To count, 1 To 3)
    
    For r = 1 To count
        For c = 1 To 3
            finalResult(r, c) = tempResult(r, c)
        Next c
    Next r
    
    If count = 2 Then
        If CDbl(finalResult(2, 2)) > CDbl(finalResult(1, 2)) Then
            Dim T1 As Variant, t2 As Variant, t3 As Variant
            T1 = finalResult(1, 1): t2 = finalResult(1, 2): t3 = finalResult(1, 3)
            finalResult(1, 1) = finalResult(2, 1)
            finalResult(1, 2) = finalResult(2, 2)
            finalResult(1, 3) = finalResult(2, 3)
            finalResult(2, 1) = T1
            finalResult(2, 2) = t2
            finalResult(2, 3) = t3
        End If
    End If
    
    GetOtkupCandidatesForKooperantBlock_Preview = finalResult
End Function

Private Function GetKooperantNazivPreview(ByVal kooperantID As String) As String
    GetKooperantNazivPreview = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime")) & _
                               " " & _
                               CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"))
End Function

Private Function NormalizeLoosePreview(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, vbTab, " ")
    s = Replace(s, ",", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "/", " ")
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeLoosePreview = Trim$(s)
End Function

Private Function NzBankaPreview(ByVal v As Variant, Optional ByVal fallback As Variant = "") As Variant
    If IsError(v) Then
        NzBankaPreview = fallback
    ElseIf IsNull(v) Then
        NzBankaPreview = fallback
    ElseIf IsEmpty(v) Then
        NzBankaPreview = fallback
    ElseIf Trim$(CStr(v)) = "" Then
        NzBankaPreview = fallback
    Else
        NzBankaPreview = v
    End If
End Function
