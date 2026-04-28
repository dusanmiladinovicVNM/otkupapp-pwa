Attribute VB_Name = "modNovac"
Option Explicit

' ============================================================
' modNovac
' ============================================================

Public Function GetBankaByPartner(ByVal partnerNaziv As String, _
                                  ByVal datumOd As Date, _
                                  ByVal datumDo As Date, _
                                  Optional ByVal omID As String = "") As Variant
    Const SRC As String = "GetBankaByPartner"

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

    Dim colPartner As Long
    Dim colDatum As Long
    Dim colOMID As Long

    colPartner = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER, SRC)
    colDatum = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, SRC)

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init colPartner, "=", partnerNaziv
    filters.Add fp

    Set fp = New clsFilterParam
    fp.Init colDatum, "BETWEEN", datumOd, datumDo
    filters.Add fp

    If Len(Trim$(omID)) > 0 Then
        colOMID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OM_ID, SRC)

        Set fp = New clsFilterParam
        fp.Init colOMID, "=", omID
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
    tx.AddTableSnapshot TBL_OTKUP

    SaveNovac_TX = SaveNovac(brojDok, datum, partner, partnerID, _
                              entitetTip, omID, kooperantID, fakturaID, _
                              vrstaVoca, tip, uplata, isplata, napomena, otkupID)

    If SaveNovac_TX = "" Then
        Err.Raise vbObjectError + 1015, "SaveNovac_TX", _
                  "SaveNovac fehlgeschlagen"
    End If

    tx.CommitTx
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr "SaveNovac_TX"
    tx.RollbackTx
    On Error GoTo 0

    SaveNovac_TX = ""

    Debug.Print "SaveNovac_TX failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function
Public Function SaveNovac(ByVal brojDok As String, ByVal datum As Date, _
                          ByVal partner As String, ByVal partnerID As String, _
                          ByVal entitetTip As String, ByVal omID As String, _
                          ByVal kooperantID As String, ByVal fakturaID As String, _
                          ByVal vrstaVoca As String, ByVal tip As String, _
                          ByVal uplata As Double, ByVal isplata As Double, _
                          Optional ByVal napomena As String = "", _
                          Optional ByVal otkupID As String = "") As String
                          
    Const SRC As String = "SaveNovac"

    Call ValidateNovacInput( _
        brojDok:=brojDok, _
        datum:=datum, _
        partner:=partner, _
        partnerID:=partnerID, _
        entitetTip:=entitetTip, _
        tip:=tip, _
        uplata:=uplata, _
        isplata:=isplata, _
        sourceName:=SRC)
    
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
    Const SRC As String = "LookupPartnerMap"

    Dim data As Variant
    data = GetTableData(TBL_PARTNER_MAP)

    If IsEmpty(data) Then
        LookupPartnerMap = Empty
        Exit Function
    End If

    Dim colName As Long
    Dim colPID As Long
    Dim colTip As Long
    Dim colOM As Long

    colName = RequireColumnIndex(TBL_PARTNER_MAP, COL_PM_BANKA_NAME, SRC)
    colPID = RequireColumnIndex(TBL_PARTNER_MAP, COL_PM_PARTNER_ID, SRC)
    colTip = RequireColumnIndex(TBL_PARTNER_MAP, COL_PM_ENTITET_TIP, SRC)
    colOM = RequireColumnIndex(TBL_PARTNER_MAP, COL_PM_OM_ID, SRC)

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

Public Function savePartnerMap(ByVal bankaName As String, _
                               ByVal partnerID As String, _
                               ByVal entitetTip As String, _
                               ByVal omID As String) As Boolean
    Const SRC As String = "savePartnerMap"

    If Len(Trim$(bankaName)) = 0 Then
        Err.Raise vbObjectError + 1036, SRC, _
                  "BankaName je obavezan za partner mapu."
    End If

    If Len(Trim$(partnerID)) = 0 Then
        Err.Raise vbObjectError + 1037, SRC, _
                  "PartnerID je obavezan za partner mapu."
    End If

    If Len(Trim$(entitetTip)) = 0 Then
        Err.Raise vbObjectError + 1038, SRC, _
                  "EntitetTip je obavezan za partner mapu."
    End If

    Dim existing As Variant
    existing = LookupPartnerMap(bankaName)

    If Not IsEmpty(existing) Then

        If UCase$(Trim$(CStr(existing(0)))) = UCase$(Trim$(partnerID)) And _
           UCase$(Trim$(CStr(existing(1)))) = UCase$(Trim$(entitetTip)) And _
           UCase$(Trim$(CStr(existing(2)))) = UCase$(Trim$(omID)) Then

            savePartnerMap = True
            Exit Function

        End If

        Err.Raise vbObjectError + 1039, SRC, _
                  "BankaName already mapped to a different partner. " & _
                  "BankaName=" & bankaName & _
                  " ExistingPartnerID=" & CStr(existing(0)) & _
                  " ExistingEntitetTip=" & CStr(existing(1)) & _
                  " ExistingOMID=" & CStr(existing(2)) & _
                  " NewPartnerID=" & partnerID & _
                  " NewEntitetTip=" & entitetTip & _
                  " NewOMID=" & omID
    End If

    Dim rowData As Variant
    rowData = Array(bankaName, partnerID, entitetTip, omID)

    If AppendRow(TBL_PARTNER_MAP, rowData) <= 0 Then
        Err.Raise vbObjectError + 1040, SRC, _
                  "Failed to append partner map row. BankaName=" & bankaName
    End If

    savePartnerMap = True
End Function

Private Function GetVrstaFromFaktura(ByVal fakturaID As String) As String
    Const SRC As String = "GetVrstaFromFaktura"

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then
        GetVrstaFromFaktura = "(Nepoznato)"
        Exit Function
    End If

    stavkeData = ExcludeStornirano(stavkeData, TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then
        GetVrstaFromFaktura = "(Nepoznato)"
        Exit Function
    End If

    Dim colFID As Long
    Dim colPrijemnicaID As Long

    colFID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC)
    colPrijemnicaID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC)

    Dim i As Long
    For i = 1 To UBound(stavkeData, 1)

        If Trim$(CStr(stavkeData(i, colFID))) = Trim$(fakturaID) Then
            Dim prijID As String
            prijID = Trim$(CStr(stavkeData(i, colPrijemnicaID)))

            GetVrstaFromFaktura = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, prijID, COL_PRJ_VRSTA))

            If Len(Trim$(GetVrstaFromFaktura)) = 0 Then
                GetVrstaFromFaktura = "(Nepoznato)"
            End If

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
    
    novacData = ExcludeStornirano(novacData, TBL_NOVAC)
    
    If IsEmpty(novacData) Then
        Set GetUplataByVrsta = dict
        Exit Function
    End If
    
    ' Cache: FakturaID ? VrstaVoca
    Dim vrstaFakCache As Object
    Set vrstaFakCache = BuildVrstaFakturaCache()
    
    Const SRC As String = "GetUplataByVrsta"

    Dim colPID As Long, colUplata As Long, colDatum As Long, colFakID As Long, colVrsta As Long
    colPID = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID, SRC)
    colUplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_UPLATA, SRC)
    colDatum = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, SRC)
    colFakID = RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC)
    colVrsta = RequireColumnIndex(TBL_NOVAC, COL_NOV_VRSTA, SRC)
    
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

    data = ExcludeStornirano(data, TBL_NOVAC)

    If IsEmpty(data) Then
        GetUplataForFaktura = 0
        Exit Function
    End If

    Const SRC As String = "GetUplataForFaktura"

    Dim colFakID As Long, colUplata As Long
    colFakID = RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC)
    colUplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_UPLATA, SRC)

    Dim total As Double
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colFakID))) = Trim$(fakturaID) Then
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

    If kupacID = "" Or fakturaID = "" Then
        Err.Raise vbObjectError + 1016, "ApplyAvansToFaktura_TX", _
                  "KupacID i FakturaID su obavezni."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE

    ApplyAvansToFaktura kupacID, fakturaID

    tx.CommitTx

    ApplyAvansToFaktura_TX = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr "ApplyAvansToFaktura_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    ApplyAvansToFaktura_TX = False

    Debug.Print "ApplyAvansToFaktura_TX failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function

Public Sub ApplyAvansToFaktura(ByVal kupacID As String, ByVal fakturaID As String)
    ' Suche alle unverbrauchten Avans-Zahlungen für diesen Kupac
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Sub
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Sub
    
    Const SRC As String = "ApplyAvansToFaktura"

    Dim colID As Long, colPID As Long, colTip As Long, colUplata As Long, colFakID As Long
    Dim colBrojDok As Long, colDatum As Long, colPartner As Long

    colID = RequireColumnIndex(TBL_NOVAC, COL_NOV_ID, SRC)
    colPID = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID, SRC)
    colTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, SRC)
    colUplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_UPLATA, SRC)
    colFakID = RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC)

    colBrojDok = RequireColumnIndex(TBL_NOVAC, COL_NOV_BROJ_DOK, SRC)
    colDatum = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, SRC)
    colPartner = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER, SRC)
    
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
            ' Full avans consumption: link existing avans row to faktura.
            Dim rows As Collection
            Set rows = FindRows(TBL_NOVAC, COL_NOV_ID, CStr(data(i, colID)))

            If rows Is Nothing Or rows.count = 0 Then
                Err.Raise vbObjectError + 1024, "ApplyAvansToFaktura", _
                    "Avans row not found for NovacID=" & CStr(data(i, colID))
            End If

            RequireUpdateCell TBL_NOVAC, rows(1), COL_NOV_FAKTURA_ID, fakturaID, _
                        "ApplyAvansToFaktura"

        Else
            ' Partial avans consumption: reduce original row and create consumed split row.
            Dim origRows As Collection
            Set origRows = FindRows(TBL_NOVAC, COL_NOV_ID, CStr(data(i, colID)))

            If origRows Is Nothing Or origRows.count = 0 Then
                    Err.Raise vbObjectError + 1025, "ApplyAvansToFaktura", _
                        "Avans row not found for split. NovacID=" & CStr(data(i, colID))
            End If

            RequireUpdateCell TBL_NOVAC, origRows(1), COL_NOV_UPLATA, avansIznos - apply, _
                            "ApplyAvansToFaktura"

            Dim splitNovacID As String
            splitNovacID = SaveNovac( _
                CStr(data(i, colBrojDok)), _
                CDate(data(i, colDatum)), _
                CStr(data(i, colPartner)), _
                kupacID, _
                "Kupac", _
                "", _
                "", _
                fakturaID, _
                "", _
                NOV_KUPCI_AVANS, _
                apply, _
                0, _
                "Avans raspodela")

            If Len(Trim$(splitNovacID)) = 0 Then
                Err.Raise vbObjectError + 1026, "ApplyAvansToFaktura", _
                        "Failed to create split avans row for FakturaID=" & fakturaID
            End If
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
    data = ExcludeStornirano(data, TBL_FAKTURE)

    If IsEmpty(data) Then
        GetOpenFakture = Empty
        Exit Function
    End If
    
    ' Uplata-Dict vorberechnen
    Dim uplataDict As Object
    Set uplataDict = BuildUplataDictByFaktura()
    
    Const SRC As String = "GetOpenFakture"

    Dim colID As Long, colBroj As Long, colKupac As Long, colIznos As Long, colStatus As Long
    colID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, SRC)
    colBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, SRC)
    colKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, SRC)
    colIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, SRC)
    colStatus = RequireColumnIndex(TBL_FAKTURE, COL_FAK_STATUS, SRC)
    
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

    data = ExcludeStornirano(data, TBL_NOVAC)

    If Not IsArray(data) Then
        Set BuildUplataDictByFaktura = dict
        Exit Function
    End If

    Const SRC As String = "BuildUplataDictByFaktura"

    Dim colFakID As Long, colUplata As Long
    colFakID = RequireColumnIndex(TBL_NOVAC, COL_NOV_FAKTURA_ID, SRC)
    colUplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_UPLATA, SRC)

    Dim i As Long
    For i = 1 To UBound(data, 1)

        Dim fID As String
        fID = Trim$(CStr(data(i, colFakID)))

        If Len(fID) > 0 Then
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
    stavkeData = ExcludeStornirano(stavkeData, TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then
        Set BuildVrstaFakturaCache = dict
        Exit Function
    End If
    
    Const SRC As String = "BuildVrstaFakturaCache"

    Dim colFID As Long
    Dim colPrijID As Long

    colFID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC)
    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC)
    
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
    ' Historical name: returns total Isplata linked to OtkupID.
    Const SRC As String = "GetUplataForOtkup"

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    Dim colOtkID As Long
    Dim colIsplata As Long

    colOtkID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC)
    colIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, SRC)

    Dim i As Long
    For i = 1 To UBound(data, 1)

        If Trim$(CStr(data(i, colOtkID))) = Trim$(otkupID) Then
            If IsNumeric(data(i, colIsplata)) Then
                GetUplataForOtkup = GetUplataForOtkup + CDbl(data(i, colIsplata))
            End If
        End If

    Next i
End Function

Public Sub UpdateOtkupStatus(ByVal otkupID As String)
    Const SRC As String = "UpdateOtkupStatus"

    If Len(Trim$(otkupID)) = 0 Then
        Err.Raise vbObjectError + 1043, SRC, _
                  "OtkupID je obavezan."
    End If

    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)

    If IsEmpty(otkupData) Then Exit Sub

    Dim rows As Collection
    Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)

    If rows Is Nothing Or rows.count = 0 Then
        Err.Raise vbObjectError + 1044, SRC, _
                  "Otkup row not found. OtkupID=" & otkupID
    End If

    Dim r As Long
    r = CLng(rows(1))

    Dim colKol As Long
    Dim colCena As Long
    Dim colDatumIsplate As Long
    Dim colStornirano As Long

    colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, SRC)
    colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, SRC)
    colDatumIsplate = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM_ISPLATE, SRC)

    colStornirano = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)

    If colStornirano > 0 Then
        If UCase$(Trim$(CStr(otkupData(r, colStornirano)))) = "DA" Then
            Exit Sub
        End If
    End If

    Dim vrednost As Double
    vrednost = 0#

    If IsNumeric(otkupData(r, colKol)) And IsNumeric(otkupData(r, colCena)) Then
        vrednost = CDbl(otkupData(r, colKol)) * CDbl(otkupData(r, colCena))
    End If

    Dim placeno As Double
    placeno = GetIsplataForOtkup(otkupID)

    If vrednost > 0 And placeno >= vrednost Then

        RequireUpdateCell TBL_OTKUP, r, COL_OTK_ISPLACENO, STATUS_ISPLACENO, SRC

        If Len(Trim$(CStr(otkupData(r, colDatumIsplate)))) = 0 Then
            RequireUpdateCell TBL_OTKUP, r, COL_OTK_DATUM_ISPLATE, Date, SRC
        End If

    Else

        RequireUpdateCell TBL_OTKUP, r, COL_OTK_ISPLACENO, "", SRC
        RequireUpdateCell TBL_OTKUP, r, COL_OTK_DATUM_ISPLATE, "", SRC

    End If
End Sub

Public Function GetIsplataForOtkup(ByVal otkupID As String) As Double
    GetIsplataForOtkup = GetUplataForOtkup(otkupID)
End Function

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
    
    Const SRC As String = "BuildIsplataDictByOtkup"

    Dim colOtkID As Long, colIsplata As Long
    colOtkID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC)
    colIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, SRC)
    
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
    
    Const SRC As String = "GetOpenOtkupi"

    Dim colID As Long, colBrDok As Long, colKoop As Long
    Dim colKol As Long, colCena As Long, colIspl As Long
    colID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    colBrDok = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, SRC)
    colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, SRC)
    colIspl = RequireColumnIndex(TBL_OTKUP, COL_OTK_ISPLACENO, SRC)
    
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
    Const SRC As String = "GetOMAvansSaldo"

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    Dim colOMID As Long
    Dim colTip As Long
    Dim colIsplata As Long

    colOMID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OM_ID, SRC)
    colTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, SRC)
    colIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, SRC)

    Dim avansTotal As Double
    Dim isplataTotal As Double

    Dim i As Long
    For i = 1 To UBound(data, 1)

        If Trim$(CStr(data(i, colOMID))) <> Trim$(omID) Then GoTo NextRow
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
    
    Const SRC As String = "ApplyAvansToOtkup"

    Dim colID As Long, colKoopID As Long, colTip As Long
    Dim colIsplata As Long, colOtkID As Long
    Dim colBrojDok As Long, colDatum As Long, colPartner As Long
    Dim colPartnerID As Long, colOMID As Long

    colID = RequireColumnIndex(TBL_NOVAC, COL_NOV_ID, SRC)
    colKoopID = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, SRC)
    colTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, SRC)
    colIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, SRC)
    colOtkID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC)

    colBrojDok = RequireColumnIndex(TBL_NOVAC, COL_NOV_BROJ_DOK, SRC)
    colDatum = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, SRC)
    colPartner = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER, SRC)
    colPartnerID = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID, SRC)
    colOMID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OM_ID, SRC)
    
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

            If avansRows Is Nothing Or avansRows.count = 0 Then
                Err.Raise vbObjectError + 1027, SRC, _
                        "Avans row not found for NovacID=" & CStr(data(i, colID))
            End If

            RequireUpdateCell TBL_NOVAC, avansRows(1), COL_NOV_OTKUP_ID, otkupID, SRC

        Else
            applyAmt = preostalo

            Set avansRows = FindRows(TBL_NOVAC, COL_NOV_ID, CStr(data(i, colID)))

            If avansRows Is Nothing Or avansRows.count = 0 Then
                Err.Raise vbObjectError + 1028, SRC, _
                        "Avans row not found for split. NovacID=" & CStr(data(i, colID))
            End If

            RequireUpdateCell TBL_NOVAC, avansRows(1), COL_NOV_ISPLATA, avansIznos - applyAmt, SRC

            Dim splitNovacID As String
            splitNovacID = SaveNovac( _
                CStr(data(i, colBrojDok)), _
                CDate(data(i, colDatum)), _
                CStr(data(i, colPartner)), _
                CStr(data(i, colPartnerID)), _
                "Kooperant", _
                CStr(data(i, colOMID)), _
                kooperantID, _
                "", _
                "", _
                NOV_VIRMAN_AVANS_KOOP, _
                0, _
                applyAmt, _
                "Avans raspodela", _
                otkupID)

            If Len(Trim$(splitNovacID)) = 0 Then
                Err.Raise vbObjectError + 1029, SRC, _
                        "Failed to create split avans row for OtkupID=" & otkupID
            End If
        End If
        
        preostalo = preostalo - applyAmt
NextAvans:
    Next i
    
    If preostalo <= 0 Then UpdateOtkupStatus otkupID
End Sub

Public Function ApplyAvansToOtkup_TX(ByVal kooperantID As String, _
                                      ByVal otkupID As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If Trim$(kooperantID) = "" Or Trim$(otkupID) = "" Then
        Err.Raise vbObjectError + 1017, "ApplyAvansToOtkup_TX", _
                  "KooperantID i OtkupID su obavezni."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP

    ApplyAvansToOtkup kooperantID, otkupID

    tx.CommitTx
    Set tx = Nothing

    ApplyAvansToOtkup_TX = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr "ApplyAvansToOtkup_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    ApplyAvansToOtkup_TX = False

    Debug.Print "ApplyAvansToOtkup_TX failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function
Public Sub ResetNovacOtkupLink(ByVal otkupID As String)
    Const SRC As String = "ResetNovacOtkupLink"

    If Len(Trim$(otkupID)) = 0 Then
        Err.Raise vbObjectError + 1041, SRC, _
                  "OtkupID je obavezan."
    End If

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Sub

    Dim colOtkID As Long
    Dim colStornirano As Long

    colOtkID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC)
    colStornirano = GetColumnIndex(TBL_NOVAC, COL_STORNIRANO)

    Dim i As Long
    For i = 1 To UBound(data, 1)

        If colStornirano > 0 Then
            If UCase$(Trim$(CStr(data(i, colStornirano)))) = "DA" Then
                GoTo NextRow
            End If
        End If

        If Trim$(CStr(data(i, colOtkID))) = Trim$(otkupID) Then
            RequireUpdateCell TBL_NOVAC, i, COL_NOV_OTKUP_ID, "", SRC
        End If

NextRow:
    Next i
End Sub

Public Function ResetNovacOtkupLink_TX(ByVal otkupID As String) As Boolean
    Const SRC As String = "ResetNovacOtkupLink_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If Len(Trim$(otkupID)) = 0 Then
        Err.Raise vbObjectError + 1042, SRC, _
                  "OtkupID je obavezan."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_OTKUP

    Call ResetNovacOtkupLink(otkupID)
    Call UpdateOtkupStatus(otkupID)

    tx.CommitTx
    Set tx = Nothing

    ResetNovacOtkupLink_TX = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr SRC

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    ResetNovacOtkupLink_TX = False

    Debug.Print SRC & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function

Public Function GetAgroAbzug(ByVal kooperantID As String) As Double
    Const SRC As String = "GetAgroAbzug"

    Dim data As Variant
    data = GetTableData(TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    data = ExcludeStornirano(data, TBL_NOVAC)

    If IsEmpty(data) Then Exit Function

    Dim colKoop As Long
    Dim colTip As Long
    Dim colUplata As Long

    colKoop = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, SRC)
    colTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, SRC)
    colUplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_UPLATA, SRC)

    Dim i As Long
    For i = 1 To UBound(data, 1)

        If Trim$(CStr(data(i, colKoop))) = Trim$(kooperantID) And _
           CStr(data(i, colTip)) = "AgroAbzug" Then

            If IsNumeric(data(i, colUplata)) Then
                GetAgroAbzug = GetAgroAbzug + CDbl(data(i, colUplata))
            End If

        End If

    Next i
End Function

Private Sub ValidateNovacInput(ByVal brojDok As String, _
                               ByVal datum As Date, _
                               ByVal partner As String, _
                               ByVal partnerID As String, _
                               ByVal entitetTip As String, _
                               ByVal tip As String, _
                               ByVal uplata As Double, _
                               ByVal isplata As Double, _
                               ByVal sourceName As String)

    If Len(Trim$(tip)) = 0 Then
        Err.Raise vbObjectError + 1030, sourceName, _
                  "Tip novca je obavezan."
    End If

    If uplata < 0 Or isplata < 0 Then
        Err.Raise vbObjectError + 1031, sourceName, _
                  "Uplata/Isplata ne sme biti negativna."
    End If

    If uplata > 0 And isplata > 0 Then
        Err.Raise vbObjectError + 1032, sourceName, _
                  "Novac red ne sme imati i uplatu i isplatu."
    End If

    If uplata = 0 And isplata = 0 Then
        Err.Raise vbObjectError + 1033, sourceName, _
                  "Novac red mora imati uplatu ili isplatu."
    End If

    If Len(Trim$(partnerID)) = 0 And Len(Trim$(partner)) = 0 Then
        Err.Raise vbObjectError + 1034, sourceName, _
                  "Partner ili PartnerID je obavezan."
    End If

    If Len(Trim$(entitetTip)) = 0 Then
        Err.Raise vbObjectError + 1035, sourceName, _
                  "EntitetTip je obavezan."
    End If
End Sub
