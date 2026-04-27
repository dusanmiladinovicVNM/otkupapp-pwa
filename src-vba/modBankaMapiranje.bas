Attribute VB_Name = "modBankaMapiranje"
Option Explicit

' ============================================================
' modBankaMapiranje
'
' Mapiranje iz tblBankaImport -> tblNovac
'
' Logika uskladjena sa stvarnim unosom iz frmDokumenta:
'
' 1) Kupac uplata:
'    - Partner = Naziv kupca
'    - PartnerID = KupacID
'    - EntitetTip = "Kupac"
'    - Tip = NOV_KUPCI_UPLATA ili NOV_KUPCI_AVANS
'
' 2) Isplata kooperantu preko banke:
'    - Partner = Naziv stanice / OM
'    - PartnerID = StanicaID
'    - EntitetTip = "OM"
'    - OMID = StanicaID
'    - KooperantID = KooperantID
'    - Tip = NOV_VIRMAN_FIRMA_KOOP ili NOV_VIRMAN_AVANS_KOOP
'
' 3) Uplata OM / dopuna OM:
'    - Partner = Naziv stanice
'    - PartnerID = StanicaID
'    - EntitetTip = "OM"
'    - OMID = StanicaID
'    - Tip = NOV_KES_FIRMA_OTKUPAC
'
' Obradjeno u tblBankaImport:
'   ""      = nije obradjeno
'   "Da"    = obradjeno
'   "Skip"  = preskoceno
'   "Error" = nije moguce automatski mapirati
'
' Napomena u tblNovac:
'   BIM:<id>; Ref:<...>; Konto:<...>; Opis:<...>; Svrha:<...>
' ============================================================


' ============================================================
' PUBLIC
' ============================================================

Public Function GetBankaImportOpen() As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim colObr As Long
    Dim i As Long, j As Long
    Dim outRow As Long
    
    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then
        GetBankaImportOpen = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    If IsEmpty(data) Then
        GetBankaImportOpen = Empty
        Exit Function
    End If
    
    colObr = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO)
    
    ReDim result(1 To UBound(data, 1), 1 To UBound(data, 2))
    
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(NzBIM(data(i, colObr), ""))) <> "Da" _
           And Trim$(CStr(NzBIM(data(i, colObr), ""))) <> "Skip" Then
            outRow = outRow + 1
            For j = 1 To UBound(data, 2)
                result(outRow, j) = data(i, j)
            Next j
        End If
    Next i
    
    If outRow = 0 Then
        GetBankaImportOpen = Empty
        Exit Function
    End If

    Dim finalResult() As Variant
    Dim r As Long, c As Long

    ReDim finalResult(1 To outRow, 1 To UBound(data, 2))

    For r = 1 To outRow
        For c = 1 To UBound(data, 2)
            finalResult(r, c) = result(r, c)
        Next c
    Next r

    GetBankaImportOpen = finalResult
End Function

Public Function AutoMapBankaImportRow(ByVal bankaImportID As String) As String
    Dim bim As Variant
    Dim uplata As Double
    Dim isplata As Double
    Dim partnerName As String
    
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        AutoMapBankaImportRow = ""
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        AutoMapBankaImportRow = ""
        Exit Function
    End If
    
    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    isplata = CDbl(NzBIM(bim(1, 6), 0#))
    partnerName = CStr(bim(1, 3))
    
    If uplata > 0 And isplata = 0 Then
        AutoMapBankaImportRow = AutoMapIncomingKupac(bankaImportID)
        If AutoMapBankaImportRow = "" Then UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If
    
    If isplata > 0 And uplata = 0 Then
        AutoMapBankaImportRow = AutoMapOutgoingKooperantOrOM(bankaImportID)
        If AutoMapBankaImportRow = "" Then UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If
    
    MsgBox "Stavka nema cist smer uplata/isplata: " & partnerName, vbExclamation, APP_NAME
    UpdateBankaImportStatus bankaImportID, "Error"
    AutoMapBankaImportRow = ""
End Function

Public Function AutoMapBankaImportRow_TX(ByVal bankaImportID As String) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_PARTNER_MAP
    tx.AddTableSnapshot TBL_OTKUP
    
    AutoMapBankaImportRow_TX = AutoMapBankaImportRow(bankaImportID)
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "AutoMapBankaImportRow_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri automatskom mapiranju banke, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    AutoMapBankaImportRow_TX = ""
End Function

Public Function MapBankaImportAsKupac(ByVal bankaImportID As String, _
                                      ByVal kupacID As String, _
                                      Optional ByVal fakturaID As String = "", _
                                      Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim bim As Variant
    Dim kupacNaziv As String
    Dim tip As String
    
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        MapBankaImportAsKupac = ""
        Exit Function
    End If
    
    If Trim$(kupacID) = "" Then
        MsgBox "KupacID je obavezan!", vbExclamation, APP_NAME
        MapBankaImportAsKupac = ""
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        MapBankaImportAsKupac = ""
        Exit Function
    End If
    
    kupacNaziv = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv"))
    If kupacNaziv = "" Then kupacNaziv = CStr(bim(1, 3))
    
    If fakturaID <> "" Then
        tip = NOV_KUPCI_UPLATA
    Else
        tip = NOV_KUPCI_AVANS
    End If
    
    MapBankaImportAsKupac = SaveNovac( _
        CStr(IIf(Trim$(CStr(bim(1, 1))) = "", "IZVOD", CStr(bim(1, 1)))), _
        CDate(bim(1, 2)), _
        kupacNaziv, _
        kupacID, _
        "Kupac", _
        "", _
        "", _
        fakturaID, _
        "", _
        tip, _
        CDbl(NzBIM(bim(1, 5), 0#)), _
        CDbl(NzBIM(bim(1, 6), 0#)), _
        BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kupac") _
    )
    
    If MapBankaImportAsKupac = "" Then
        UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If
    
    UpdateBankaImportStatus bankaImportID, "Da"
    
    If savePartnerMapFlag Then
        Call savePartnerMap(CStr(bim(1, 3)), kupacID, "Kupac", "")
    End If
    
    If fakturaID <> "" Then UpdateFakturaStatus fakturaID
End Function

Public Function MapBankaImportAsKupac_TX(ByVal bankaImportID As String, _
                                         ByVal kupacID As String, _
                                         Optional ByVal fakturaID As String = "", _
                                         Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKupac_TX = MapBankaImportAsKupac(bankaImportID, kupacID, fakturaID, savePartnerMapFlag)
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "MapBankaImportasKupac_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri mapiranju kupca, promene vracene: " & Err.Description, vbCritical, APP_NAME
    MapBankaImportAsKupac_TX = ""
End Function

Public Function MapBankaImportAsKooperant(ByVal bankaImportID As String, _
                                          ByVal kooperantID As String, _
                                          Optional ByVal otkupID As String = "", _
                                          Optional ByVal vrstaVoca As String = "", _
                                          Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim bim As Variant
    Dim omID As String
    Dim omNaziv As String
    Dim tip As String
    
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    If Trim$(kooperantID) = "" Then
        MsgBox "KooperantID je obavezan!", vbExclamation, APP_NAME
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    omID = CStr(NzBIM(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_STANICA), ""))
    If omID = "" Then
        MsgBox "Kooperant nema StanicaID!", vbExclamation, APP_NAME
        MapBankaImportAsKooperant = ""
        Exit Function
    End If
    
    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    If omNaziv = "" Then omNaziv = omID
    
    If otkupID <> "" Then
        tip = NOV_VIRMAN_FIRMA_KOOP
    Else
        tip = NOV_VIRMAN_AVANS_KOOP
    End If
    
    MapBankaImportAsKooperant = SaveNovac( _
        CStr(IIf(Trim$(CStr(bim(1, 1))) = "", "IZVOD", CStr(bim(1, 1)))), _
        CDate(bim(1, 2)), _
        omNaziv, _
        omID, _
        "OM", _
        omID, _
        kooperantID, _
        "", _
        vrstaVoca, _
        tip, _
        CDbl(NzBIM(bim(1, 5), 0#)), _
        CDbl(NzBIM(bim(1, 6), 0#)), _
        BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant") _
    )
    
    If MapBankaImportAsKooperant = "" Then
        UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If
    
    UpdateBankaImportStatus bankaImportID, "Da"
    
    If savePartnerMapFlag Then
        Call savePartnerMap(CStr(bim(1, 3)), kooperantID, "Kooperant", omID)
    End If
    
    If otkupID <> "" Then
        Dim novRows As Collection
        Set novRows = FindRows(TBL_NOVAC, COL_NOV_ID, MapBankaImportAsKooperant)
        If novRows.count > 0 Then
            UpdateCell TBL_NOVAC, novRows(1), COL_NOV_OTKUP_ID, otkupID
            UpdateOtkupStatus otkupID
        End If
    End If
End Function

Public Function MapBankaImportAsKooperant_TX(ByVal bankaImportID As String, _
                                             ByVal kooperantID As String, _
                                             Optional ByVal otkupID As String = "", _
                                             Optional ByVal vrstaVoca As String = "", _
                                             Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKooperant_TX = MapBankaImportAsKooperant(bankaImportID, kooperantID, otkupID, vrstaVoca, savePartnerMapFlag)
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "MapBankaImportAsKooperant_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri mapiranju kooperanta, promene vracene: " & Err.Description, vbCritical, APP_NAME
    MapBankaImportAsKooperant_TX = ""
End Function

Public Function MapBankaImportAsOM(ByVal bankaImportID As String, _
                                   ByVal omID As String, _
                                   Optional ByVal vrstaVoca As String = "", _
                                   Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim bim As Variant
    Dim omNaziv As String
    
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        MapBankaImportAsOM = ""
        Exit Function
    End If
    
    If Trim$(omID) = "" Then
        MsgBox "OMID je obavezan!", vbExclamation, APP_NAME
        MapBankaImportAsOM = ""
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        MapBankaImportAsOM = ""
        Exit Function
    End If
    
    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    If omNaziv = "" Then omNaziv = omID
    
    MapBankaImportAsOM = SaveNovac( _
        CStr(IIf(Trim$(CStr(bim(1, 1))) = "", "IZVOD", CStr(bim(1, 1)))), _
        CDate(bim(1, 2)), _
        omNaziv, _
        omID, _
        "OM", _
        omID, _
        "", _
        "", _
        vrstaVoca, _
        NOV_KES_FIRMA_OTKUPAC, _
        CDbl(NzBIM(bim(1, 5), 0#)), _
        CDbl(NzBIM(bim(1, 6), 0#)), _
        BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "OM") _
    )
    
    If MapBankaImportAsOM = "" Then
        UpdateBankaImportStatus bankaImportID, "Error"
        Exit Function
    End If
    
    UpdateBankaImportStatus bankaImportID, "Da"
    
    If savePartnerMapFlag Then
        Call savePartnerMap(CStr(bim(1, 3)), omID, "OM", omID)
    End If
End Function

Public Function MapBankaImportAsOM_TX(ByVal bankaImportID As String, _
                                      ByVal omID As String, _
                                      Optional ByVal vrstaVoca As String = "", _
                                      Optional ByVal savePartnerMapFlag As Boolean = True) As String
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsOM_TX = MapBankaImportAsOM(bankaImportID, omID, vrstaVoca, savePartnerMapFlag)
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "MapBankaImportAsOM_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri mapiranju OM, promene vracene: " & Err.Description, vbCritical, APP_NAME
    MapBankaImportAsOM_TX = ""
End Function

Public Function MapBankaImportAsKooperantBlock(ByVal bankaImportID As String, _
                                               ByVal kooperantID As String, _
                                               Optional ByVal savePartnerMapFlag As Boolean = True) As Long
    Dim blockNo As String
    
    blockNo = Trim$(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    MapBankaImportAsKooperantBlock = MapBankaImportAsKooperantBlockCore( _
        bankaImportID, kooperantID, blockNo, savePartnerMapFlag)
End Function

Public Function MapBankaImportAsKooperantBlock_TX(ByVal bankaImportID As String, _
                                                  ByVal kooperantID As String, _
                                                  Optional ByVal savePartnerMapFlag As Boolean = True) As Long
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKooperantBlock_TX = MapBankaImportAsKooperantBlock(bankaImportID, kooperantID, savePartnerMapFlag)
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "MapBankaImportAsKooperantBlock_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri mapiranju kooperanta po bloku, promene vracene: " & Err.Description, vbCritical, APP_NAME
    MapBankaImportAsKooperantBlock_TX = 0
End Function

Public Function MapBankaImportAsKooperantBlockManual(ByVal bankaImportID As String, _
                                                     ByVal kooperantID As String, _
                                                     ByVal brojBloka As String, _
                                                     Optional ByVal savePartnerMapFlag As Boolean = True) As Long
    MapBankaImportAsKooperantBlockManual = MapBankaImportAsKooperantBlockCore( _
        bankaImportID, kooperantID, brojBloka, savePartnerMapFlag)
End Function

Public Function MapBankaImportAsKooperantBlockManual_TX(ByVal bankaImportID As String, _
                                                        ByVal kooperantID As String, _
                                                        ByVal brojBloka As String, _
                                                        Optional ByVal savePartnerMapFlag As Boolean = True) As Long
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_PARTNER_MAP
    
    MapBankaImportAsKooperantBlockManual_TX = MapBankaImportAsKooperantBlockManual( _
        bankaImportID, kooperantID, brojBloka, savePartnerMapFlag)
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "MapBankaImportAsKooperantBlockManual_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri rucnom mapiranju kooperanta po bloku, promene vracene: " & Err.Description, vbCritical, APP_NAME
    MapBankaImportAsKooperantBlockManual_TX = 0
End Function

Private Function MapBankaImportAsKooperantBlockCore(ByVal bankaImportID As String, _
                                                    ByVal kooperantID As String, _
                                                    ByVal blockNo As String, _
                                                    Optional ByVal savePartnerMapFlag As Boolean = True) As Long
    Dim bim As Variant
    Dim omID As String
    Dim omNaziv As String
    Dim isplataUkupno As Double
    Dim preostaloZaRaspodelu As Double
    Dim kandidati As Variant
    Dim i As Long
    
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then Exit Function
    
    If Trim$(kooperantID) = "" Then
        MsgBox "KooperantID je obavezan!", vbExclamation, APP_NAME
        Exit Function
    End If
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        Exit Function
    End If
    
    omID = CStr(NzBIM(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_STANICA), ""))
    If omID = "" Then
        MsgBox "Kooperant nema StanicaID!", vbExclamation, APP_NAME
        Exit Function
    End If
    
    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    If omNaziv = "" Then omNaziv = omID
    
    isplataUkupno = CDbl(NzBIM(bim(1, 6), 0#))
    If isplataUkupno <= 0 Then Exit Function
    
    kandidati = GetOtkupCandidatesForKooperantBlock(kooperantID, blockNo)
    preostaloZaRaspodelu = isplataUkupno
    
    If IsEmpty(kandidati) Then
        If SaveNovac( _
            CStr(IIf(Trim$(CStr(bim(1, 1))) = "", "IZVOD", CStr(bim(1, 1)))), _
            CDate(bim(1, 2)), _
            omNaziv, _
            omID, _
            "OM", _
            omID, _
            kooperantID, _
            "", _
            "", _
            NOV_VIRMAN_AVANS_KOOP, _
            0, _
            preostaloZaRaspodelu, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant") _
        ) <> "" Then
            MapBankaImportAsKooperantBlockCore = 1
            UpdateBankaImportStatus bankaImportID, "Da"
            If savePartnerMapFlag Then
                Call savePartnerMap(CStr(bim(1, 3)), kooperantID, "Kooperant", omID)
            End If
        Else
            UpdateBankaImportStatus bankaImportID, "Error"
        End If
        Exit Function
    End If
    
    For i = 1 To UBound(kandidati, 1)
        If preostaloZaRaspodelu <= 0 Then Exit For
        
        Dim otkupID As String
        Dim otvoreno As Double
        Dim iznosZaRed As Double
        Dim vrstaVoca As String
        Dim novID As String
        
        otkupID = CStr(kandidati(i, 1))
        otvoreno = CDbl(NzBIM(kandidati(i, 2), 0#))
        vrstaVoca = CStr(kandidati(i, 3))
        
        If otvoreno <= 0 Then GoTo NextCandidate
        
        If preostaloZaRaspodelu >= otvoreno Then
            iznosZaRed = otvoreno
        Else
            iznosZaRed = preostaloZaRaspodelu
        End If
        
        novID = SaveNovac( _
            CStr(IIf(Trim$(CStr(bim(1, 1))) = "", "IZVOD", CStr(bim(1, 1)))), _
            CDate(bim(1, 2)), _
            omNaziv, _
            omID, _
            "OM", _
            omID, _
            kooperantID, _
            "", _
            vrstaVoca, _
            NOV_VIRMAN_FIRMA_KOOP, _
            0, _
            iznosZaRed, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant") _
        )
        
        If novID <> "" Then
            Dim novRows As Collection
            Set novRows = FindRows(TBL_NOVAC, COL_NOV_ID, novID)
            If novRows.count > 0 Then
                UpdateCell TBL_NOVAC, novRows(1), COL_NOV_OTKUP_ID, otkupID
            End If
            
            UpdateOtkupStatus otkupID
            MapBankaImportAsKooperantBlockCore = MapBankaImportAsKooperantBlockCore + 1
            preostaloZaRaspodelu = preostaloZaRaspodelu - iznosZaRed
        End If
        
NextCandidate:
    Next i
    
    If preostaloZaRaspodelu > 0 Then
        If SaveNovac( _
            CStr(IIf(Trim$(CStr(bim(1, 1))) = "", "IZVOD", CStr(bim(1, 1)))), _
            CDate(bim(1, 2)), _
            omNaziv, _
            omID, _
            "OM", _
            omID, _
            kooperantID, _
            "", _
            "", _
            NOV_VIRMAN_AVANS_KOOP, _
            0, _
            preostaloZaRaspodelu, _
            BuildBIMNapomena(bankaImportID, CStr(bim(1, 9)), CStr(bim(1, 4)), CStr(bim(1, 7)), CStr(bim(1, 8)), "Kooperant-visak") _
        ) <> "" Then
            MapBankaImportAsKooperantBlockCore = MapBankaImportAsKooperantBlockCore + 1
        End If
    End If
    
    If MapBankaImportAsKooperantBlockCore > 0 Then
        UpdateBankaImportStatus bankaImportID, "Da"
        If savePartnerMapFlag Then
            Call savePartnerMap(CStr(bim(1, 3)), kooperantID, "Kooperant", omID)
        End If
    Else
        UpdateBankaImportStatus bankaImportID, "Error"
    End If
End Function


Public Function SkipBankaImportRow(ByVal bankaImportID As String) As Boolean
    If Not ValidateBankaImportNotProcessed(bankaImportID) Then
        SkipBankaImportRow = False
        Exit Function
    End If
    
    UpdateBankaImportStatus bankaImportID, "Skip"
    SkipBankaImportRow = True
End Function

Public Function SkipBankaImportRow_TX(ByVal bankaImportID As String) As Boolean
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    
    SkipBankaImportRow_TX = SkipBankaImportRow(bankaImportID)
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "SkipBankaImportRow_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri preskakanju bank stavke, promene vracene: " & Err.Description, vbCritical, APP_NAME
    SkipBankaImportRow_TX = False
End Function

Public Function AutoMapAllBankaImport_TX() As Long
    Dim tx As clsTransaction
    Dim data As Variant
    Dim colID As Long
    Dim i As Long
    Dim novID As String
    
    On Error GoTo EH
    
    data = GetBankaImportOpen()
    If IsEmpty(data) Then
        AutoMapAllBankaImport_TX = 0
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_PARTNER_MAP
    tx.AddTableSnapshot TBL_OTKUP
    
    For i = 1 To UBound(data, 1)
        novID = AutoMapBankaImportRow(CStr(data(i, colID)))
        If novID <> "" Then AutoMapAllBankaImport_TX = AutoMapAllBankaImport_TX + 1
    Next i
    
    tx.CommitTx
    Exit Function
    
EH:
    LogErr "AutoMapAllBankaImport_TX"
    If Not tx Is Nothing Then tx.RollbackTx
    MsgBox "Greska pri automatskom mapiranju svih bank stavki, promene vracene: " & Err.Description, vbCritical, APP_NAME
    AutoMapAllBankaImport_TX = 0
End Function


' ============================================================
' PRIVATE - AUTO MAP
' ============================================================

Private Function AutoMapIncomingKupac(ByVal bankaImportID As String) As String
    Dim bim As Variant
    Dim partnerName As String
    Dim mapped As Variant
    Dim fakturaID As String
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function
    
    partnerName = CStr(bim(1, 3))
    
    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        If CStr(mapped(1)) = "Kupac" Then
            fakturaID = TryResolveFakturaForKupac(bankaImportID, CStr(mapped(0)))
            AutoMapIncomingKupac = MapBankaImportAsKupac(bankaImportID, CStr(mapped(0)), fakturaID, False)
            Exit Function
        End If
    End If
    
    mapped = TryResolveKupacBIM(partnerName)
    If Not IsEmpty(mapped) Then
        fakturaID = TryResolveFakturaForKupac(bankaImportID, CStr(mapped(0)))
        AutoMapIncomingKupac = MapBankaImportAsKupac(bankaImportID, CStr(mapped(0)), fakturaID, False)
        Exit Function
    End If
    
    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        AutoMapIncomingKupac = MapBankaImportAsOM(bankaImportID, CStr(mapped(0)), "", False)
        Exit Function
    End If
    
    AutoMapIncomingKupac = ""
End Function

Private Function AutoMapOutgoingKooperantOrOM(ByVal bankaImportID As String) As String
    Dim bim As Variant
    Dim partnerName As String
    Dim mapped As Variant
    Dim createdCount As Long
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function
    
    partnerName = CStr(bim(1, 3))
    
    mapped = LookupPartnerMap(partnerName)
    If Not IsEmpty(mapped) Then
        Select Case CStr(mapped(1))
            Case "Kooperant"
                createdCount = MapBankaImportAsKooperantBlock(bankaImportID, CStr(mapped(0)), False)
                If createdCount > 0 Then AutoMapOutgoingKooperantOrOM = "OK"
                Exit Function
            Case "OM"
                AutoMapOutgoingKooperantOrOM = MapBankaImportAsOM(bankaImportID, CStr(mapped(0)), "", False)
                Exit Function
        End Select
    End If
    
    mapped = TryResolveKooperantBIM(partnerName)
    If Not IsEmpty(mapped) Then
        createdCount = MapBankaImportAsKooperantBlock(bankaImportID, CStr(mapped(0)), False)
        If createdCount > 0 Then AutoMapOutgoingKooperantOrOM = "OK"
        Exit Function
    End If
    
    mapped = TryResolveOMBIM(partnerName)
    If Not IsEmpty(mapped) Then
        AutoMapOutgoingKooperantOrOM = MapBankaImportAsOM(bankaImportID, CStr(mapped(0)), "", False)
        Exit Function
    End If
    
    AutoMapOutgoingKooperantOrOM = ""
End Function



' ============================================================
' PUBLIC - FACTURA MATCH
' ============================================================

Public Function TryResolveFakturaForKupac(ByVal bankaImportID As String, ByVal kupacID As String) As String
    Dim bim As Variant
    Dim faktData As Variant
    Dim colFID As Long, colBroj As Long, colKup As Long, colIznos As Long, colStatus As Long
    Dim poziv As String, svrha As String, uplata As Double
    Dim i As Long, hitCount As Long, hitID As String
    
    bim = GetBankaImportRowByID(bankaImportID)
    If IsEmpty(bim) Then Exit Function
    
    poziv = NormalizeLooseBIM(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    svrha = NormalizeLooseBIM(CStr(bim(1, 8)))
    uplata = CDbl(NzBIM(bim(1, 5), 0#))
    
    faktData = GetTableData(TBL_FAKTURE)
    If IsEmpty(faktData) Then Exit Function
    
    faktData = ExcludeStornirano(faktData, TBL_FAKTURE)
    If IsEmpty(faktData) Then Exit Function
    
    colFID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
    colKup = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)
    colIznos = GetColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS)
    colStatus = GetColumnIndex(TBL_FAKTURE, COL_FAK_STATUS)
    
    For i = 1 To UBound(faktData, 1)
        If CStr(faktData(i, colKup)) <> kupacID Then GoTo NextI
        
        Dim brojFak As String
        brojFak = NormalizeLooseBIM(CStr(faktData(i, colBroj)))
        
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
        
        If Abs(CDbl(NzBIM(faktData(i, colIznos), 0#)) - uplata) < 0.01 Then
            hitCount = hitCount + 1
            hitID = CStr(faktData(i, colFID))
        End If
        
NextI:
    Next i
    
    If hitCount = 1 Then TryResolveFakturaForKupac = hitID
End Function

Private Function TryResolveOtkupForKooperant(ByVal bankaImportID As String, _
                                             ByVal kooperantID As String) As String
    Dim pozivNaBroj As String
    Dim otkData As Variant
    Dim colBrDok As Long, colOtkID As Long, colKoop As Long
    Dim i As Long
    Dim hitCount As Long
    Dim hitID As String
    
    pozivNaBroj = Trim$(CStr(LookupValue(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID, COL_BIM_POZIV_NA_BROJ)))
    If pozivNaBroj = "" Then Exit Function
    
    otkData = GetTableData(TBL_OTKUP)
    If IsEmpty(otkData) Then Exit Function
    
    otkData = ExcludeStornirano(otkData, TBL_OTKUP)
    If IsEmpty(otkData) Then Exit Function
    
    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    
    For i = 1 To UBound(otkData, 1)
        If CStr(otkData(i, colKoop)) <> kooperantID Then GoTo NextI
        
        If NormalizeLooseBIM(CStr(otkData(i, colBrDok))) = NormalizeLooseBIM(pozivNaBroj) Then
            hitCount = hitCount + 1
            hitID = CStr(otkData(i, colOtkID))
        End If
        
NextI:
    Next i
    
    If hitCount = 1 Then
        TryResolveOtkupForKooperant = hitID
    End If
End Function

Public Function GetOtkupCandidatesForKooperantBlock(ByVal kooperantID As String, _
                                                     ByVal brojBloka As String) As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim colOtkID As Long
    Dim colKoop As Long
    Dim colBrDok As Long
    Dim colKol As Long
    Dim colCena As Long
    Dim colVrsta As Long
    Dim i As Long
    Dim count As Long
    
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
    
    ReDim result(1 To 2, 1 To 3)
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) <> kooperantID Then GoTo NextI
        
        If NormalizeLooseBIM(CStr(data(i, colBrDok))) = NormalizeLooseBIM(brojBloka) Then
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
                
                result(count, 1) = CStr(data(i, colOtkID))
                result(count, 2) = otvoreno
                result(count, 3) = CStr(data(i, colVrsta))
            End If
        End If
        
NextI:
    Next i
    
    If count = 0 Then Exit Function

    Dim finalResult() As Variant
    Dim r As Long, c As Long

    ReDim finalResult(1 To count, 1 To 3)

    For r = 1 To count
        For c = 1 To 3
            finalResult(r, c) = result(r, c)
        Next c
    Next r
    
    ' Max 2 reda -> jednostavan swap, veca otvorena prva
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
    
    GetOtkupCandidatesForKooperantBlock = finalResult
End Function

' ============================================================
' PUBLIC - ACCESS
' ============================================================

Public Function GetBankaImportRowByID(ByVal bankaImportID As String) As Variant
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
    If IsEmpty(data) Then
        GetBankaImportRowByID = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    If IsEmpty(data) Then
        GetBankaImportRowByID = Empty
        Exit Function
    End If
    
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
            result(1, 5) = CDbl(NzBIM(data(i, colUplata), 0#))
            result(1, 6) = CDbl(NzBIM(data(i, colIsplata), 0#))
            result(1, 7) = CStr(data(i, colOpis))
            result(1, 8) = CStr(data(i, colSvrha))
            result(1, 9) = CStr(data(i, colRef))
            result(1, 10) = CStr(data(i, colPoziv))
            GetBankaImportRowByID = result
            Exit Function
        End If
    Next i
    
    GetBankaImportRowByID = Empty
End Function

Private Function FindBankaImportRowIndex(ByVal bankaImportID As String) As Long
    Dim rows As Collection
    
    Set rows = FindRows(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID)
    If rows Is Nothing Then Exit Function
    If rows.count = 0 Then Exit Function
    
    FindBankaImportRowIndex = CLng(rows(1))
End Function

Private Function ValidateBankaImportNotProcessed(ByVal bankaImportID As String) As Boolean
    Dim rows As Collection
    Dim data As Variant
    Dim colObr As Long
    Dim colStorno As Long
    Dim r As Long
    
    Set rows = FindRows(TBL_BANKA_IMPORT, COL_BIM_ID, bankaImportID)
    If rows Is Nothing Or rows.count = 0 Then
        MsgBox "BankaImport red nije pronadjen: " & bankaImportID, vbExclamation, APP_NAME
        ValidateBankaImportNotProcessed = False
        Exit Function
    End If
    
    r = CLng(rows(1))
    
    data = GetTableData(TBL_BANKA_IMPORT)
    If IsEmpty(data) Then
        MsgBox "tblBankaImport je prazan!", vbExclamation, APP_NAME
        ValidateBankaImportNotProcessed = False
        Exit Function
    End If
    
    colObr = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO)
    colStorno = GetColumnIndex(TBL_BANKA_IMPORT, COL_STORNIRANO)
    
    If colStorno > 0 Then
        If CStr(data(r, colStorno)) = "Da" Then
            MsgBox "BankaImport red je storniran!", vbExclamation, APP_NAME
            ValidateBankaImportNotProcessed = False
            Exit Function
        End If
    End If
    
    If Trim$(CStr(NzBIM(data(r, colObr), ""))) = "Da" Then
        MsgBox "BankaImport red je vec obradjen!", vbExclamation, APP_NAME
        ValidateBankaImportNotProcessed = False
        Exit Function
    End If
    
    If Trim$(CStr(NzBIM(data(r, colObr), ""))) = "Skip" Then
        MsgBox "BankaImport red je vec preskocen!", vbExclamation, APP_NAME
        ValidateBankaImportNotProcessed = False
        Exit Function
    End If
    
    ValidateBankaImportNotProcessed = True
End Function

Private Sub UpdateBankaImportStatus(ByVal bankaImportID As String, ByVal newStatus As String)
    Dim rowIdx As Long
    
    rowIdx = FindBankaImportRowIndex(bankaImportID)
    If rowIdx <= 0 Then Exit Sub
    
    UpdateCell TBL_BANKA_IMPORT, rowIdx, COL_BIM_OBRADJENO, newStatus
End Sub


' ============================================================
' PUBLIC - ENTITY RESOLUTION
' ============================================================

Public Function TryResolveKupacBIM(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colNaziv As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    
    data = GetTableData(TBL_KUPCI)
    If IsEmpty(data) Then
        TryResolveKupacBIM = Empty
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_KUPCI, "KupacID")
    colNaziv = GetColumnIndex(TBL_KUPCI, "Naziv")
    
    For i = 1 To UBound(data, 1)
        If NormalizeLooseBIM(CStr(data(i, colNaziv))) = NormalizeLooseBIM(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i
    
    If hits = 1 Then TryResolveKupacBIM = Array(hitID, "Kupac", "")
End Function

Public Function TryResolveKooperantBIM(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colIme As Long
    Dim colPrezime As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    Dim hitOMID As String
    Dim fullName As String
    
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then
        TryResolveKooperantBIM = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_KOOPERANTI)
    If IsEmpty(data) Then
        TryResolveKooperantBIM = Empty
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_KOOPERANTI, "KooperantID")
    colIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    colPrezime = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    
    For i = 1 To UBound(data, 1)
        fullName = NormalizeLooseBIM(CStr(data(i, colIme)) & " " & CStr(data(i, colPrezime)))
        If fullName = NormalizeLooseBIM(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
            hitOMID = CStr(NzBIM(LookupValue(TBL_KOOPERANTI, "KooperantID", hitID, COL_KOOP_STANICA), ""))
        End If
    Next i
    
    If hits = 1 Then TryResolveKooperantBIM = Array(hitID, "Kooperant", hitOMID)
End Function

Public Function TryResolveOMBIM(ByVal partnerName As String) As Variant
    Dim data As Variant
    Dim colID As Long
    Dim colNaziv As Long
    Dim i As Long
    Dim hits As Long
    Dim hitID As String
    
    data = GetTableData(TBL_STANICE)
    If IsEmpty(data) Then
        TryResolveOMBIM = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_STANICE)
    If IsEmpty(data) Then
        TryResolveOMBIM = Empty
        Exit Function
    End If
    
    colID = GetColumnIndex(TBL_STANICE, "StanicaID")
    colNaziv = GetColumnIndex(TBL_STANICE, "Naziv")
    
    For i = 1 To UBound(data, 1)
        If NormalizeLooseBIM(CStr(data(i, colNaziv))) = NormalizeLooseBIM(partnerName) Then
            hits = hits + 1
            hitID = CStr(data(i, colID))
        End If
    Next i
    
    If hits = 1 Then TryResolveOMBIM = Array(hitID, "OM", hitID)
End Function


' ============================================================
' PRIVATE/PUBLIC - HELPERS
' ============================================================

Private Function BuildBIMNapomena(ByVal bankaImportID As String, _
                                  ByVal bankaRef As String, _
                                  ByVal partnerKonto As String, _
                                  ByVal opis As String, _
                                  ByVal svrha As String, _
                                  ByVal reason As String) As String
    Dim s As String
    
    s = "BIM:" & bankaImportID
    If Trim$(bankaRef) <> "" Then s = s & "; Ref:" & bankaRef
    If Trim$(partnerKonto) <> "" Then s = s & "; Konto:" & partnerKonto
    If Trim$(opis) <> "" Then s = s & "; Opis:" & Left$(opis, 80)
    If Trim$(svrha) <> "" Then s = s & "; Svrha:" & Left$(svrha, 120)
    If Trim$(reason) <> "" Then s = s & "; Match:" & Left$(reason, 50)
    
    BuildBIMNapomena = s
End Function

Public Function NormalizeLooseBIM(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, vbTab, " ")
    s = Replace(s, ",", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "/", " ")
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeLooseBIM = Trim$(s)
End Function

Public Function NzBIM(ByVal v As Variant, Optional ByVal Fallback As Variant = "") As Variant
    If IsError(v) Then
        NzBIM = Fallback
    ElseIf IsNull(v) Then
        NzBIM = Fallback
    ElseIf IsEmpty(v) Then
        NzBIM = Fallback
    ElseIf Trim$(CStr(v)) = "" Then
        NzBIM = Fallback
    Else
        NzBIM = v
    End If
End Function

Public Function GetKooperantNaziv(ByVal kooperantID As String) As String
    GetKooperantNaziv = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime")) & _
                        " " & _
                        CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"))
End Function

' ============================================================
' TESTS
' ============================================================

Public Sub Test_GetBankaImportOpen()
    Dim data As Variant
    Dim i As Long
    Dim colID As Long
    
    data = GetBankaImportOpen()
    If IsEmpty(data) Then
        Debug.Print "No open rows."
        Exit Sub
    End If
    
    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    
    For i = 1 To UBound(data, 1)
        Debug.Print data(i, colID)
    Next i
End Sub

Public Sub Test_MapBankaImportAsKooperantBlock_TX()
    Dim bankaImportID As String
    Dim kooperantID As String
    Dim n As Long
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    kooperantID = InputBox$("KooperantID:")
    If Trim$(kooperantID) = "" Then Exit Sub
    
    n = MapBankaImportAsKooperantBlock_TX(bankaImportID, kooperantID, True)
    Debug.Print "Created rows: " & n
End Sub

Public Sub Test_AutoMapBankaImportRow_TX()
    Dim bankaImportID As String
    Dim novacID As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    novacID = AutoMapBankaImportRow_TX(bankaImportID)
    Debug.Print "Created NovacID: " & novacID
End Sub

Public Sub Test_MapBankaImportAsKupac_TX()
    Dim bankaImportID As String
    Dim kupacID As String
    Dim fakturaID As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    kupacID = InputBox$("KupacID:")
    If Trim$(kupacID) = "" Then Exit Sub
    
    fakturaID = InputBox$("FakturaID (optional):")
    
    Debug.Print MapBankaImportAsKupac_TX(bankaImportID, kupacID, fakturaID, True)
End Sub

Public Sub Test_MapBankaImportAsKooperant_TX()
    Dim bankaImportID As String
    Dim kooperantID As String
    Dim otkupID As String
    Dim vrsta As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    kooperantID = InputBox$("KooperantID:")
    If Trim$(kooperantID) = "" Then Exit Sub
    
    otkupID = InputBox$("OtkupID (optional):")
    vrsta = InputBox$("VrstaVoca (optional):")
    
    Debug.Print MapBankaImportAsKooperant_TX(bankaImportID, kooperantID, otkupID, vrsta, True)
End Sub

Public Sub Test_MapBankaImportAsOM_TX()
    Dim bankaImportID As String
    Dim omID As String
    Dim vrsta As String
    
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    omID = InputBox$("OMID:")
    If Trim$(omID) = "" Then Exit Sub
    
    vrsta = InputBox$("VrstaVoca (optional):")
    
    Debug.Print MapBankaImportAsOM_TX(bankaImportID, omID, vrsta, True)
End Sub

Public Sub Test_SkipBankaImportRow_TX()
    Dim bankaImportID As String
     
    bankaImportID = InputBox$("BankaImportID:")
    If Trim$(bankaImportID) = "" Then Exit Sub
    
    If SkipBankaImportRow_TX(bankaImportID) Then
        Debug.Print "Skipped: " & bankaImportID
    End If
End Sub

