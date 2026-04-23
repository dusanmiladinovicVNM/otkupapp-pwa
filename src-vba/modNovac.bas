Attribute VB_Name = "modNovac"
Option Explicit

' ============================================================
' modNovac
' ============================================================

Public Function GetBankaByPartner(ByVal partnerNaziv As String, _
                                  ByVal datumOd As Date, _
                                  ByVal datumDo As Date, _
                                  Optional ByVal omID As String = "") As Variant
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then
        GetBankaByPartner = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then
        GetBankaByPartner = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER), "=", partnerNaziv
    filters.Add fp
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    If omID <> "" Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_NOVAC, COL_NOV_OM_ID), "=", omID
        filters.Add fp
    End If
    
    GetBankaByPartner = FilterArray(data, filters)
End Function

Public Function SaveNovac_TX(ByVal brojDok As String, ByVal datum As Date, _
                              ByVal partner As String, ByVal partnerID As String, _
                              ByVal entitetTip As String, ByVal omID As String, _
                              ByVal kooperantID As String, ByVal fakturaID As String, _
                              ByVal vrstaVoca As String, ByVal tip As String, _
                              ByVal uplata As Double, ByVal isplata As Double, _
                              Optional ByVal napomena As String = "", _
                              Optional ByVal otkupID As String = "") As String
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_OTKUP    ' ? NEU (UpdateOtkupStatus)
    
    SaveNovac_TX = SaveNovac(brojDok, datum, partner, partnerID, _
                              entitetTip, omID, kooperantID, fakturaID, _
                              vrstaVoca, tip, uplata, isplata, napomena, otkupID)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "SaveNovac_TX"
    tx.RollbackTx
    MsgBox "Greska pri unosu novca, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    SaveNovac_TX = ""
End Function

Public Function SaveNovac(ByVal brojDok As String, ByVal datum As Date, _
                          ByVal partner As String, ByVal partnerID As String, _
                          ByVal entitetTip As String, ByVal omID As String, _
                          ByVal kooperantID As String, ByVal fakturaID As String, _
                          ByVal vrstaVoca As String, ByVal tip As String, _
                          ByVal uplata As Double, ByVal isplata As Double, _
                          Optional ByVal napomena As String = "", _
                          Optional ByVal otkupID As String = "") As String
    
    Dim newID As String
    newID = GetNextID(TBL_NOVAC, COL_NOV_ID, "NOV-")
    
    Dim rowData As Variant
    rowData = Array(newID, brojDok, datum, partner, partnerID, _
                    entitetTip, omID, kooperantID, fakturaID, _
                    vrstaVoca, tip, uplata, isplata, napomena, _
                    "", otkupID, "") ' Stornirano, OtkupID, OsirocenoOD
    
    If AppendRow(TBL_NOVAC, rowData) > 0 Then
        SaveNovac = newID
    Else
        SaveNovac = ""
    End If
End Function

Public Function LookupPartnerMap(ByVal bankaName As String) As Variant
    ' Returns: Array(PartnerID, EntitetTip, OMID) oder Empty
    
    Dim data As Variant
    data = GetTableData(TBL_PARTNER_MAP)
    If IsEmpty(data) Then
        LookupPartnerMap = Empty
        Exit Function
    End If
    
    Dim colName As Long, colPID As Long, colTip As Long, colOM As Long
    colName = GetColumnIndex(TBL_PARTNER_MAP, COL_PM_BANKA_NAME)
    colPID = GetColumnIndex(TBL_PARTNER_MAP, COL_PM_PARTNER_ID)
    colTip = GetColumnIndex(TBL_PARTNER_MAP, COL_PM_ENTITET_TIP)
    colOM = GetColumnIndex(TBL_PARTNER_MAP, COL_PM_OM_ID)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(CStr(data(i, colName)))) = UCase$(Trim$(bankaName)) Then
            LookupPartnerMap = Array( _
                CStr(data(i, colPID)), _
                CStr(data(i, colTip)), _
                CStr(data(i, colOM)))
            Exit Function
        End If
    Next i
    
    LookupPartnerMap = Empty
End Function

Public Function savePartnerMap(ByVal bankaName As String, ByVal partnerID As String, _
                               ByVal entitetTip As String, ByVal omID As String) As Boolean
    ' Prüfe ob schon existiert
    Dim existing As Variant
    existing = LookupPartnerMap(bankaName)
    If Not IsEmpty(existing) Then
        savePartnerMap = True
        Exit Function
    End If
    
    Dim rowData As Variant
    rowData = Array(bankaName, partnerID, entitetTip, omID)
    
    savePartnerMap = (AppendRow(TBL_PARTNER_MAP, rowData) > 0)
End Function

Private Function GetVrstaFromFaktura(ByVal fakturaID As String) As String
    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If IsEmpty(stavkeData) Then
        GetVrstaFromFaktura = "(Nepoznato)"
        Exit Function
    End If
    
    Dim colFID As Long, colPrijemnicaID As Long
    colFID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID)
    colPrijemnicaID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID)
    
    Dim i As Long
    For i = 1 To UBound(stavkeData, 1)
        If CStr(stavkeData(i, colFID)) = fakturaID Then
            Dim prijID As String
            prijID = CStr(stavkeData(i, colPrijemnicaID))
            GetVrstaFromFaktura = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijID, COL_PRJ_VRSTA))
            If GetVrstaFromFaktura = "" Then GetVrstaFromFaktura = "(Nepoznato)"
            Exit Function
        End If
    Next i
    
    GetVrstaFromFaktura = "(Nepoznato)"
End Function

Public Function GetUplataByVrsta(ByVal kupacID As String, _
                                 ByVal datumOd As Date, _
                                 ByVal datumDo As Date) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim novacData As Variant
    novacData = GetTableData(TBL_NOVAC)
    If IsEmpty(novacData) Then
        Set GetUplataByVrsta = dict
        Exit Function
    End If
    
    ' Cache: FakturaID ? VrstaVoca
    Dim vrstaFakCache As Object
    Set vrstaFakCache = BuildVrstaFakturaCache()
    
    Dim colPID As Long, colUplata As Long, colDatum As Long, colFakID As Long, colVrsta As Long
    colPID = GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    colDatum = GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM)
    colFakID = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)
    colVrsta = GetColumnIndex(TBL_NOVAC, COL_NOV_VRSTA)
    
    Dim n As Long
    For n = 1 To UBound(novacData, 1)
        If CStr(novacData(n, colPID)) <> kupacID Then GoTo NextRow
        If Not IsDate(novacData(n, colDatum)) Then GoTo NextRow
        If CDate(novacData(n, colDatum)) < datumOd Or CDate(novacData(n, colDatum)) > datumDo Then GoTo NextRow
        If Not IsNumeric(novacData(n, colUplata)) Then GoTo NextRow
        If CDbl(novacData(n, colUplata)) <= 0 Then GoTo NextRow
        
        Dim uplata As Double
        uplata = CDbl(novacData(n, colUplata))
        
        Dim vrsta As String
        Dim fakturaID As String
        fakturaID = CStr(novacData(n, colFakID))
        
        If fakturaID <> "" Then
            If vrstaFakCache.Exists(fakturaID) Then
                vrsta = vrstaFakCache(fakturaID)
            Else
                vrsta = "(Nepoznato)"
            End If
        Else
            vrsta = CStr(novacData(n, colVrsta))
            If vrsta = "" Then vrsta = "(Nerasporedeno)"
        End If
        
        If Not dict.Exists(vrsta) Then dict.Add vrsta, 0#
        dict(vrsta) = dict(vrsta) + uplata
NextRow:
    Next n
    
    Set GetUplataByVrsta = dict
End Function

Public Function GetUplataForFaktura(ByVal fakturaID As String) As Double
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then
        GetUplataForFaktura = 0
        Exit Function
    End If
    
    Dim colFakID As Long, colUplata As Long
    colFakID = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    
    Dim total As Double
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colFakID)) = fakturaID Then
            If IsNumeric(data(i, colUplata)) Then
                total = total + CDbl(data(i, colUplata))
            End If
        End If
    Next i
    
    GetUplataForFaktura = total
End Function

Public Function ApplyAvansToFaktura_TX(ByVal kupacID As String, _
                                        ByVal fakturaID As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE
    
    ApplyAvansToFaktura kupacID, fakturaID
    ApplyAvansToFaktura_TX = True
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "ApplyAvansToFaktura_TX"
    tx.RollbackTx
    MsgBox "Greska pri raspodeli avansa, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    ApplyAvansToFaktura_TX = False
End Function

Public Sub ApplyAvansToFaktura(ByVal kupacID As String, ByVal fakturaID As String)
    ' Suche alle unverbrauchten Avans-Zahlungen für diesen Kupac
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Sub
    
    Dim colID As Long, colPID As Long, colTip As Long, colUplata As Long, colFakID As Long
    colID = GetColumnIndex(TBL_NOVAC, COL_NOV_ID)
    colPID = GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID)
    colTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    colFakID = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)
    
    ' Faktura-Iznos und bereits bezahlt
    Dim fakIznos As Double
    fakIznos = CDbl(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS))
    Dim fakUplaceno As Double
    fakUplaceno = GetUplataForFaktura(fakturaID)
    Dim preostalo As Double
    preostalo = fakIznos - fakUplaceno
    
    If preostalo <= 0 Then Exit Sub
    
    ' Alle Avans-Zeilen für diesen Kupac sammeln (chronologisch)
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If preostalo <= 0 Then Exit For
        If CStr(data(i, colPID)) <> kupacID Then GoTo NextAvans
        If CStr(data(i, colTip)) <> NOV_KUPCI_AVANS Then GoTo NextAvans
        If CStr(data(i, colFakID)) <> "" Then GoTo NextAvans  ' bereits verrechnet
        If Not IsNumeric(data(i, colUplata)) Then GoTo NextAvans
        
        Dim avansIznos As Double
        avansIznos = CDbl(data(i, colUplata))
        If avansIznos <= 0 Then GoTo NextAvans
        
        ' Wie viel von diesem Avans verrechnen?
        Dim apply As Double
        If avansIznos <= preostalo Then
            ' Ganzer Avans wird verbraucht
            apply = avansIznos
        Else
            ' Avans ist größer als Restbetrag ? aufteilen
            apply = preostalo
        End If
        
        If avansIznos <= preostalo Then
            ' Ganzer Avans ? FakturaID setzen
            Dim rows As Collection
            Set rows = FindRows(TBL_NOVAC, COL_NOV_ID, CStr(data(i, colID)))
            If rows.count > 0 Then
                UpdateCell TBL_NOVAC, rows(1), COL_NOV_FAKTURA_ID, fakturaID
            End If
        Else
            ' Avans aufteilen: Original reduzieren, neue Zeile für den verrechneten Teil
            Dim origRows As Collection
            Set origRows = FindRows(TBL_NOVAC, COL_NOV_ID, CStr(data(i, colID)))
            If origRows.count > 0 Then
                ' Original auf Rest reduzieren
                UpdateCell TBL_NOVAC, origRows(1), COL_NOV_UPLATA, avansIznos - apply
            End If
            
            ' Neue Zeile für verrechneten Teil
            SaveNovac CStr(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_BROJ_DOK))), _
                      CDate(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM))), _
                      CStr(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER))), _
                      kupacID, "Kupac", "", "", fakturaID, "", _
                      NOV_KUPCI_AVANS, apply, 0, "Avans raspodela"
        End If
        
        preostalo = preostalo - apply
NextAvans:
    Next i
    
    ' Faktura-Status prüfen
    If preostalo <= 0 Then
        UpdateFakturaStatus fakturaID
    End If
End Sub

Public Function GetOpenFakture(ByVal kupacID As String) As Variant
    ' Returns: 2D Array (BrojFakture, FakturaID, Iznos, Uplaceno, Preostalo)
    ' oder Empty wenn nichts offen
    
    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)
    If IsEmpty(data) Then
        GetOpenFakture = Empty
        Exit Function
    End If
    
    ' Uplata-Dict vorberechnen
    Dim uplataDict As Object
    Set uplataDict = BuildUplataDictByFaktura()
    
    Dim colID As Long, colBroj As Long, colKupac As Long, colIznos As Long, colStatus As Long
    colID = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
    colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
    colKupac = GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)
    colIznos = GetColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS)
    colStatus = GetColumnIndex(TBL_FAKTURE, COL_FAK_STATUS)
    
    ' Erst zählen
    Dim count As Long
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKupac)) = kupacID And CStr(data(i, colStatus)) = STATUS_NEPLACENO Then
            Dim iznos As Double: iznos = CDbl(data(i, colIznos))
            Dim uplaceno As Double: uplaceno = 0
            If uplataDict.Exists(CStr(data(i, colID))) Then uplaceno = uplataDict(CStr(data(i, colID)))
            If iznos - uplaceno > 0 Then count = count + 1
        End If
    Next i
    
    If count = 0 Then
        GetOpenFakture = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)
    Dim idx As Long
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKupac)) = kupacID And CStr(data(i, colStatus)) = STATUS_NEPLACENO Then
            iznos = CDbl(data(i, colIznos))
            uplaceno = 0
            If uplataDict.Exists(CStr(data(i, colID))) Then uplaceno = uplataDict(CStr(data(i, colID)))
            Dim preostalo As Double: preostalo = iznos - uplaceno
            If preostalo > 0 Then
                idx = idx + 1
                result(idx, 1) = CStr(data(i, colBroj))
                result(idx, 2) = CStr(data(i, colID))
                result(idx, 3) = iznos
                result(idx, 4) = uplaceno
                result(idx, 5) = preostalo
            End If
        End If
    Next i
    
    GetOpenFakture = result
End Function

Public Function BuildUplataDictByFaktura() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If Not IsArray(data) Then
        Set BuildUplataDictByFaktura = dict
        Exit Function
    End If
    
    Dim colFakID As Long, colUplata As Long
    colFakID = GetColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim fID As String
        fID = CStr(data(i, colFakID))
        If fID <> "" Then
            If Not dict.Exists(fID) Then dict.Add fID, 0#
            If IsNumeric(data(i, colUplata)) Then
                dict(fID) = dict(fID) + CDbl(data(i, colUplata))
            End If
        End If
    Next i
    
    Set BuildUplataDictByFaktura = dict
End Function

Private Function BuildVrstaFakturaCache() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If IsEmpty(stavkeData) Then
        Set BuildVrstaFakturaCache = dict
        Exit Function
    End If
    
    Dim colFID As Long, colPrijID As Long
    colFID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID)
    colPrijID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID)
    
    Dim i As Long
    For i = 1 To UBound(stavkeData, 1)
        Dim fID As String
        fID = CStr(stavkeData(i, colFID))
        If Not dict.Exists(fID) Then
            Dim vrsta As String
            vrsta = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, _
                         CStr(stavkeData(i, colPrijID)), COL_PRJ_VRSTA))
            If vrsta = "" Then vrsta = "(Nepoznato)"
            dict.Add fID, vrsta
        End If
    Next i
    
    Set BuildVrstaFakturaCache = dict
End Function


Public Function GetUplataForOtkup(ByVal otkupID As String) As Double
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    
    Dim colOtkID As Long, colIsplata As Long
    colOtkID = GetColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID)
    colIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colOtkID)) = otkupID Then
            If IsNumeric(data(i, colIsplata)) Then
                GetUplataForOtkup = GetUplataForOtkup + CDbl(data(i, colIsplata))
            End If
        End If
    Next i
End Function

Public Sub UpdateOtkupStatus(ByVal otkupID As String)
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    
    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If rows.count = 0 Then Exit Sub
    
    Dim r As Long: r = rows(1)
    Dim colKol As Long, colCena As Long
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    
    Dim vrednost As Double
    If IsNumeric(otkupData(r, colKol)) And IsNumeric(otkupData(r, colCena)) Then
        vrednost = CDbl(otkupData(r, colKol)) * CDbl(otkupData(r, colCena))
    End If
    
    If GetUplataForOtkup(otkupID) >= vrednost And vrednost > 0 Then
        UpdateCell TBL_OTKUP, r, COL_OTK_ISPLACENO, STATUS_ISPLACENO
        UpdateCell TBL_OTKUP, r, COL_OTK_DATUM_ISPLATE, Date
    End If
End Sub

Public Function BuildIsplataDictByOtkup() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If Not IsArray(data) Then
        Set BuildIsplataDictByOtkup = dict
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_NOVAC)
    If Not IsArray(data) Then
        Set BuildIsplataDictByOtkup = dict
        Exit Function
    End If
    
    Dim colOtkID As Long, colIsplata As Long
    colOtkID = GetColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID)
    colIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim oID As String
        oID = CStr(data(i, colOtkID))
        If oID <> "" Then
            If Not dict.Exists(oID) Then dict.Add oID, 0#
            If IsNumeric(data(i, colIsplata)) Then
                dict(oID) = dict(oID) + CDbl(data(i, colIsplata))
            End If
        End If
    Next i
    
    Set BuildIsplataDictByOtkup = dict
End Function

Public Function GetOpenOtkupi(ByVal kooperantID As String) As Variant
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        GetOpenOtkupi = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then
        GetOpenOtkupi = Empty
        Exit Function
    End If
    
    Dim colID As Long, colBrDok As Long, colKoop As Long
    Dim colKol As Long, colCena As Long, colIspl As Long
    colID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    colIspl = GetColumnIndex(TBL_OTKUP, COL_OTK_ISPLACENO)
    
    Dim isplataDict As Object
    Set isplataDict = BuildIsplataDictByOtkup()
    
    ' Zählen
    Dim count As Long, i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID And _
           CStr(data(i, colIspl)) <> STATUS_ISPLACENO Then
            Dim vrednost As Double: vrednost = 0
            If IsNumeric(data(i, colKol)) And IsNumeric(data(i, colCena)) Then
                vrednost = CDbl(data(i, colKol)) * CDbl(data(i, colCena))
            End If
            Dim isplaceno As Double: isplaceno = 0
            If isplataDict.Exists(CStr(data(i, colID))) Then isplaceno = isplataDict(CStr(data(i, colID)))
            If vrednost - isplaceno > 0 Then count = count + 1
        End If
    Next i
    
    If count = 0 Then
        GetOpenOtkupi = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)
    Dim idx As Long
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID And _
           CStr(data(i, colIspl)) <> STATUS_ISPLACENO Then
            vrednost = 0
            If IsNumeric(data(i, colKol)) And IsNumeric(data(i, colCena)) Then
                vrednost = CDbl(data(i, colKol)) * CDbl(data(i, colCena))
            End If
            isplaceno = 0
            If isplataDict.Exists(CStr(data(i, colID))) Then isplaceno = isplataDict(CStr(data(i, colID)))
            If vrednost - isplaceno > 0 Then
                idx = idx + 1
                result(idx, 1) = CStr(data(i, colBrDok))
                result(idx, 2) = CStr(data(i, colID))
                result(idx, 3) = vrednost
                result(idx, 4) = isplaceno
                result(idx, 5) = vrednost - isplaceno
            End If
        End If
    Next i
    
    GetOpenOtkupi = result
End Function

Public Function GetOMAvansSaldo(ByVal omID As String) As Double
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    
    Dim colOMID As Long, colTip As Long, colIsplata As Long
    colOMID = GetColumnIndex(TBL_NOVAC, COL_NOV_OM_ID)
    colTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    colIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
    
    Dim avansTotal As Double, isplataTotal As Double
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colOMID)) <> omID Then GoTo NextRow
        If Not IsNumeric(data(i, colIsplata)) Then GoTo NextRow
        
        Select Case CStr(data(i, colTip))
            Case NOV_KES_FIRMA_OTKUPAC
                avansTotal = avansTotal + CDbl(data(i, colIsplata))
            Case NOV_KES_OTKUPAC_KOOP
                isplataTotal = isplataTotal + CDbl(data(i, colIsplata))
        End Select
NextRow:
    Next i
    
    GetOMAvansSaldo = avansTotal - isplataTotal
End Function

Public Sub ApplyAvansToOtkup(ByVal kooperantID As String, ByVal otkupID As String)
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Sub
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Sub
    
    Dim colID As Long, colKoopID As Long, colTip As Long
    Dim colIsplata As Long, colOtkID As Long
    colID = GetColumnIndex(TBL_NOVAC, COL_NOV_ID)
    colKoopID = GetColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID)
    colTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    colIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
    colOtkID = GetColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID)
    
    ' Otkup-Vrednost
    Dim otkData As Variant
    otkData = GetTableData(TBL_OTKUP)
    Dim otkRows As Collection
    Set otkRows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If otkRows.count = 0 Then Exit Sub
    
    Dim r As Long: r = otkRows(1)
    Dim colKol As Long, colCena As Long
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    
    Dim otkVrednost As Double
    If IsNumeric(otkData(r, colKol)) And IsNumeric(otkData(r, colCena)) Then
        otkVrednost = CDbl(otkData(r, colKol)) * CDbl(otkData(r, colCena))
    End If
    
    Dim preostalo As Double
    preostalo = otkVrednost - GetUplataForOtkup(otkupID)
    If preostalo <= 0 Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If preostalo <= 0 Then Exit For
        If CStr(data(i, colKoopID)) <> kooperantID Then GoTo NextAvans
        If CStr(data(i, colTip)) <> NOV_VIRMAN_AVANS_KOOP Then GoTo NextAvans
        If CStr(data(i, colOtkID)) <> "" Then GoTo NextAvans
        If Not IsNumeric(data(i, colIsplata)) Then GoTo NextAvans
        
        Dim avansIznos As Double
        avansIznos = CDbl(data(i, colIsplata))
        If avansIznos <= 0 Then GoTo NextAvans
        
        Dim applyAmt As Double
        Dim avansRows As Collection
        
        If avansIznos <= preostalo Then
            applyAmt = avansIznos
            Set avansRows = FindRows(TBL_NOVAC, COL_NOV_ID, CStr(data(i, colID)))
            If avansRows.count > 0 Then
                UpdateCell TBL_NOVAC, avansRows(1), COL_NOV_OTKUP_ID, otkupID
            End If
        Else
            applyAmt = preostalo
            Set avansRows = FindRows(TBL_NOVAC, COL_NOV_ID, CStr(data(i, colID)))
            If avansRows.count > 0 Then
                UpdateCell TBL_NOVAC, avansRows(1), COL_NOV_ISPLATA, avansIznos - applyAmt
            End If
            SaveNovac CStr(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_BROJ_DOK))), _
                      CDate(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM))), _
                      CStr(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER))), _
                      CStr(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID))), _
                      "Kooperant", _
                      CStr(data(i, GetColumnIndex(TBL_NOVAC, COL_NOV_OM_ID))), _
                      kooperantID, "", "", _
                      NOV_VIRMAN_AVANS_KOOP, 0, applyAmt, "Avans raspodela", otkupID
        End If
        
        preostalo = preostalo - applyAmt
NextAvans:
    Next i
    
    If preostalo <= 0 Then UpdateOtkupStatus otkupID
End Sub

Public Function ApplyAvansToOtkup_TX(ByVal kooperantID As String, _
                                      ByVal otkupID As String) As Boolean
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP
    
    ApplyAvansToOtkup kooperantID, otkupID
    ApplyAvansToOtkup_TX = True
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "ApplyAvansToOtkup_TX"
    tx.RollbackTx
    MsgBox "Greska pri raspodeli avansa, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    ApplyAvansToOtkup_TX = False
End Function

Public Sub ResetNovacOtkupLink(ByVal otkupID As String)
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Sub
    
    Dim colOtkID As Long
    colOtkID = GetColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colOtkID)) = otkupID Then
            UpdateCell TBL_NOVAC, i, COL_NOV_OTKUP_ID, ""
        End If
    Next i
End Sub

Public Function GetAgroAbzug(ByVal kooperantID As String) As Double
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    
    Dim colKoop As Long, colTip As Long, colUplata As Long
    colKoop = GetColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID)
    colTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    colUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID And _
           CStr(data(i, colTip)) = "AgroAbzug" Then
            If IsNumeric(data(i, colUplata)) Then
                GetAgroAbzug = GetAgroAbzug + CDbl(data(i, colUplata))
            End If
        End If
    Next i
End Function

