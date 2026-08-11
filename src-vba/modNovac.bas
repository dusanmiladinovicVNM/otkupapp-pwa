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
                              ByVal partner As String, ByVal partnerId As String, _
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

    SaveNovac_TX = SaveNovac(brojDok, datum, partner, partnerId, _
                              entitetTip, omID, kooperantID, fakturaID, _
                              vrstaVoca, tip, uplata, isplata, napomena, otkupID)

    If SaveNovac_TX = "" Then
        Err.Raise vbObjectError + 1015, "SaveNovac_TX", _
                  "SaveNovac fehlgeschlagen"
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="NOVAC_SAVE_SUCCESS", _
        severity:="INFO", _
        message:="Novac transaction saved. Tip=" & tip & _
                 "; Uplata=" & CStr(uplata) & _
                 "; Isplata=" & CStr(isplata), _
        userId:="Operator", _
        moduleName:="modNovac", _
        procedureName:="SaveNovac_TX", _
        entityType:="Novac", _
        entityID:=SaveNovac_TX, _
        correlationId:=SaveNovac_TX
    On Error GoTo 0

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE
    
    Dim corrId As String
    If Len(Trim$(fakturaID)) > 0 Then
        corrId = fakturaID
    ElseIf Len(Trim$(otkupID)) > 0 Then
        corrId = otkupID
    Else
        corrId = partnerId
    End If
    
    On Error Resume Next

    LogErr "SaveNovac_TX"

    Monitor_Error _
        moduleName:="modNovac", _
        procedureName:="SaveNovac_TX", _
        entityType:="Novac", _
        entityID:=SaveNovac_TX, _
        correlationId:=corrId, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="NOVAC_SAVE_FAIL", _
        severity:="ERROR", _
        message:="Novac transaction failed. Tip=" & tip & _
                 "; EntitetTip=" & entitetTip & _
                 "; FakturaID=" & fakturaID & _
                 "; OtkupID=" & otkupID & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modNovac", _
        procedureName:="SaveNovac_TX", _
        entityType:="Novac", _
        entityID:=SaveNovac_TX, _
        correlationId:=IIf(Len(Trim$(fakturaID)) > 0, fakturaID, otkupID)

    If Not tx Is Nothing Then tx.RollbackTx

    On Error GoTo 0

    SaveNovac_TX = ""

    Debug.Print "SaveNovac_TX failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function
Public Function SaveNovac(ByVal brojDok As String, ByVal datum As Date, _
                          ByVal partner As String, ByVal partnerId As String, _
                          ByVal entitetTip As String, ByVal omID As String, _
                          ByVal kooperantID As String, ByVal fakturaID As String, _
                          ByVal vrstaVoca As String, ByVal tip As String, _
                          ByVal uplata As Double, ByVal isplata As Double, _
                          Optional ByVal napomena As String = "", _
                          Optional ByVal otkupID As String = "") As String

    Const SRC As String = "SaveNovac"

    On Error GoTo EH

    Call ValidateNovacInput( _
        brojDok:=brojDok, _
        datum:=datum, _
        partner:=partner, _
        partnerId:=partnerId, _
        entitetTip:=entitetTip, _
        tip:=tip, _
        uplata:=uplata, _
        isplata:=isplata, _
        sourceName:=SRC)

    Dim newID As String
    newID = GetNextID(TBL_NOVAC, COL_NOV_ID, "NOV-")

    If Len(Trim$(newID)) = 0 Then
        Err.Raise vbObjectError + 1050, SRC, _
                  "GetNextID nije vratio NovacID. " & _
                  "BrojDok=" & brojDok & _
                  "; PartnerID=" & partnerId & _
                  "; Tip=" & tip & _
                  "; FakturaID=" & fakturaID & _
                  "; OtkupID=" & otkupID
    End If

    ' AUD-003 / FM-0019 #1: schema-presence guard pre pozicionog AppendRow-a.
    ' Ako je sema tblNovac driftovala (nedostaje/preimenovana kolona), fail-fast
    ' umesto tihe korupcije ledger reda. Redosled prati Array(...) ispod.
    RequireColumns TBL_NOVAC, SRC, _
                   COL_NOV_ID, _
                   COL_NOV_BROJ_DOK, _
                   COL_NOV_DATUM, _
                   COL_NOV_PARTNER, _
                   COL_NOV_PARTNER_ID, _
                   COL_NOV_ENTITET_TIP, _
                   COL_NOV_OM_ID, _
                   COL_NOV_KOOP_ID, _
                   COL_NOV_FAKTURA_ID, _
                   COL_NOV_VRSTA, _
                   COL_NOV_TIP, _
                   COL_NOV_UPLATA, _
                   COL_NOV_ISPLATA, _
                   COL_NOV_NAPOMENA, _
                   COL_STORNIRANO, _
                   COL_NOV_OTKUP_ID, _
                   COL_OSIROCENO_OD

    Dim rowData As Variant
    rowData = Array(newID, brojDok, datum, partner, partnerId, _
                    entitetTip, omID, kooperantID, fakturaID, _
                    vrstaVoca, tip, uplata, isplata, napomena, _
                    "", otkupID, "") ' Stornirano, OtkupID, OsirocenoOD

    Dim rowIdx As Long
    rowIdx = AppendRow(TBL_NOVAC, rowData)

    If rowIdx <= 0 Then
        Err.Raise vbObjectError + 1051, SRC, _
                  "AppendRow failed for " & TBL_NOVAC & ". " & _
                  "NovacID=" & newID & _
                  "; BrojDok=" & brojDok & _
                  "; PartnerID=" & partnerId & _
                  "; EntitetTip=" & entitetTip & _
                  "; Tip=" & tip & _
                  "; Uplata=" & CStr(uplata) & _
                  "; Isplata=" & CStr(isplata) & _
                  "; FakturaID=" & fakturaID & _
                  "; OtkupID=" & otkupID
    End If

    SaveNovac = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
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
                               ByVal partnerId As String, _
                               ByVal entitetTip As String, _
                               ByVal omID As String) As Boolean
    Const SRC As String = "savePartnerMap"

    If Len(Trim$(bankaName)) = 0 Then
        Err.Raise vbObjectError + 1036, SRC, _
                  "BankaName je obavezan za partner mapu."
    End If

    If Len(Trim$(partnerId)) = 0 Then
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

        If UCase$(Trim$(CStr(existing(0)))) = UCase$(Trim$(partnerId)) And _
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
                  " NewPartnerID=" & partnerId & _
                  " NewEntitetTip=" & entitetTip & _
                  " NewOMID=" & omID
    End If

    Dim rowData As Variant
    rowData = Array(bankaName, partnerId, entitetTip, omID)

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
    
    ' Cache: FakturaID -> Dictionary(VrstaVoca -> vrednost stavki) za srazmernu
    ' raspodelu uplate (RF-06 / FM-0028 #6).
    Dim vrstaFakCache As Object
    Set vrstaFakCache = BuildFakturaVrstaUdeoCache()

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
        fakturaID = Trim$(CStr(novacData(n, colFakID)))

        If fakturaID <> "" Then
            ' Uplata po fakturi se deli po vrstama SRAZMERNO stavkama fakture.
            ' Pre RF-06 je cela uplata isla na vrstu PRVE stavke, pa je izvestaj
            ' salda kupca pokazivao dug na jednoj vrsti i visak na drugoj.
            Dim udeli As Object
            Set udeli = Nothing
            If vrstaFakCache.Exists(fakturaID) Then Set udeli = vrstaFakCache(fakturaID)

            Dim podela As Object
            Set podela = RaspodeliPoUdelima(uplata, udeli)

            Dim pk As Variant
            For Each pk In podela.keys
                If Not dict.Exists(CStr(pk)) Then dict.Add CStr(pk), 0#
                dict(CStr(pk)) = dict(CStr(pk)) + CDbl(podela(pk))
            Next pk
        Else
            vrsta = CStr(novacData(n, colVrsta))
            If vrsta = "" Then vrsta = "(Nerasporedeno)"

            If Not dict.Exists(vrsta) Then dict.Add vrsta, 0#
            dict(vrsta) = dict(vrsta) + uplata
        End If
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
                                        ByVal fakturaID As String, _
                                        Optional ByRef appliedAmount As Double) As Boolean
    appliedAmount = 0

    Dim tx As New clsTransaction

    On Error GoTo EH

    If kupacID = "" Or fakturaID = "" Then
        Err.Raise vbObjectError + 1016, "ApplyAvansToFaktura_TX", _
                  "KupacID i FakturaID su obavezni."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE

    ' AUD-010 / FM-0019 #11: vrati stvarno primenjeni iznos (ByRef) uz Boolean.
    ApplyAvansToFaktura kupacID, fakturaID, appliedAmount

    tx.CommitTx
   
    ApplyAvansToFaktura_TX = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "ApplyAvansToFaktura_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    ApplyAvansToFaktura_TX = False
    appliedAmount = 0

    Debug.Print "ApplyAvansToFaktura_TX failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function

Public Sub ApplyAvansToFaktura(ByVal kupacID As String, ByVal fakturaID As String, _
                               Optional ByRef appliedAmount As Double)
    appliedAmount = 0

    ' Suche alle unverbrauchten Avans-Zahlungen fuer diesen Kupac
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
    ' Napomena roditelja -> split nasledjuje BIM marker (poreklo se ne gubi).
    Dim colNapomena As Long
    colNapomena = RequireColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA, SRC)

    ' AUD-010 / FM-0019 #4,#6: target-owner + target-active guard.
    ' Ne primeni avans na fakturu drugog kupca (ili nepostojecu) niti na storniranu.
    Dim fakKupac As String
    fakKupac = Trim$(CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_KUPAC)))
    If StrComp(fakKupac, Trim$(kupacID), vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1020, SRC, _
                  "Avans se ne moze primeniti na fakturu drugog kupca ili nepostojecu fakturu. " & _
                  "FakturaID=" & fakturaID & "; FakturaKupac=" & fakKupac & "; TrazeniKupac=" & kupacID
    End If

    Dim fakStorno As String
    fakStorno = Trim$(CStr(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_STORNIRANO)))
    If UCase$(fakStorno) = "DA" Then
        Err.Raise vbObjectError + 1021, SRC, _
                  "Avans se ne moze primeniti na storniranu fakturu. FakturaID=" & fakturaID
    End If

    ' Faktura-Iznos und bereits bezahlt
    Dim fakIznos As Double
    fakIznos = CDbl(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS))
    Dim fakUplaceno As Double
    fakUplaceno = GetUplataForFaktura(fakturaID)
    Dim preostalo As Double
    preostalo = fakIznos - fakUplaceno

    If preostalo <= 0 Then Exit Sub
    
    ' Alle Avans-Zeilen fuer diesen Kupac sammeln (chronologisch)
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
            ' Avans ist groesser als Restbetrag ? aufteilen
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
                BuildAvansSplitNapomena(CStr(data(i, colNapomena)), "Avans raspodela"))

            If Len(Trim$(splitNovacID)) = 0 Then
                Err.Raise vbObjectError + 1026, "ApplyAvansToFaktura", _
                        "Failed to create split avans row for FakturaID=" & fakturaID
            End If
        End If
        
        appliedAmount = appliedAmount + apply
        preostalo = preostalo - apply
NextAvans:
    Next i

    ' Faktura-Status pruefen
    If preostalo <= 0 Then
        UpdateFakturaStatus fakturaID
    End If
End Sub

Public Function GetOpenFakture(ByVal kupacID As String) As Variant
    ' Returns: 2D Array (BrojFakture, FakturaID, Iznos, Uplaceno, Preostalo, Datum)
    ' oder Empty wenn nichts offen
    '
    ' Jedini read-model otvorenih faktura kupca (docs/AgriX_Functional_Map_v142.md
    ' 20.12): izbacuje stornirane, trazi status Neplaceno, racuna stvarno uplaceno i
    ' vraca samo preostalo > 0. frmDokumenta.FillOpenFakture zove OVO -- ne sme se
    ' duplirati filter u formi (stara forma je imala slabiji: Status <> Placeno).
    ' Datum je 6. kolona (dodata za prikaz u formi), pa stari 5-kolonski pozivaoci
    ' ostaju nepromenjeni.

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
    Dim colDatum As Long
    colID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, SRC)
    colBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, SRC)
    colKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, SRC)
    colIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, SRC)
    colStatus = RequireColumnIndex(TBL_FAKTURE, COL_FAK_STATUS, SRC)
    colDatum = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM, SRC)
    
    ' Erst zaehlen
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
    ReDim result(1 To count, 1 To 6)
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
                result(idx, 6) = data(i, colDatum)
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

' RF-06 (AUD-023 / FM-0028 #6): FakturaID -> Dictionary(VrstaVoca -> vrednost stavki).
' Ranije je ovaj kes cuvao SAMO vrstu PRVE stavke fakture, pa je cela uplata po
' fakturi sa vise vrsta voca zavrsavala na jednoj vrsti. Sada nosi tezine svih
' stavki (Kolicina * Cena), pa se uplata deli srazmerno (RaspodeliPoUdelima).
Private Function BuildFakturaVrstaUdeoCache() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If IsEmpty(stavkeData) Then
        Set BuildFakturaVrstaUdeoCache = dict
        Exit Function
    End If
    stavkeData = ExcludeStornirano(stavkeData, TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then
        Set BuildFakturaVrstaUdeoCache = dict
        Exit Function
    End If

    Const SRC As String = "BuildFakturaVrstaUdeoCache"

    Dim colFID As Long
    Dim colPrijID As Long
    Dim colKol As Long
    Dim colCena As Long

    colFID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, SRC)
    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, SRC)
    ' Kolicina/Cena su nosioci tezine; kod stare seme bez njih sve stavke dobijaju
    ' istu tezinu (1) -> raspodela postaje ravnomerna umesto "sve na prvu".
    colKol = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KOLICINA)
    colCena = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_CENA)

    ' PrijemnicaID -> VrstaVoca, jednim prolazom (umesto LookupValue po stavci).
    Dim prijVrstaDict As Object
    Set prijVrstaDict = BuildLookupDict(TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_VRSTA)

    Dim i As Long
    For i = 1 To UBound(stavkeData, 1)
        Dim fID As String
        fID = Trim$(CStr(stavkeData(i, colFID)))
        If Len(fID) > 0 Then
            Dim prijID As String
            prijID = Trim$(CStr(stavkeData(i, colPrijID)))

            Dim vrsta As String
            If prijVrstaDict.Exists(prijID) Then vrsta = CStr(prijVrstaDict(prijID)) Else vrsta = ""
            If Trim$(vrsta) = "" Then vrsta = "(Nepoznato)"

            Dim tezina As Double
            tezina = 1
            If colKol > 0 And colCena > 0 Then
                If IsNumeric(stavkeData(i, colKol)) And IsNumeric(stavkeData(i, colCena)) Then
                    tezina = CDbl(stavkeData(i, colKol)) * CDbl(stavkeData(i, colCena))
                End If
            End If
            If tezina < 0 Then tezina = 0

            If Not dict.Exists(fID) Then dict.Add fID, CreateObject("Scripting.Dictionary")
            Dim udeli As Object
            Set udeli = dict(fID)
            If Not udeli.Exists(vrsta) Then udeli.Add vrsta, 0#
            udeli(vrsta) = udeli(vrsta) + tezina
        End If
    Next i

    Set BuildFakturaVrstaUdeoCache = dict
End Function

' Podeli iznos po tezinama iz `udeli` (kljuc -> tezina). Cist racun, bez tabela:
' testira ga modIzvestajTests.RunIzvestajTests.
'   - zbir tezina <= 0 (ili prazan dict) -> ceo iznos na "(Nepoznato)"
'   - racuna se u CELIM PARAMA metodom najvecih ostataka (largest remainder):
'     svaki deo dobije floor svog idealnog udela, pa se visak para deli redom
'     po najvecem ostatku.
'
' Dve invarijante koje metod garantuje (obe su bile potrebne):
'   1. zbir delova == ZaokruziNovac(iznos)  -- prikaz po vrstama se slaze sa
'      UKUPNO (100/3 daje 33,34 + 33,33 + 33,33, ne 33,33 x 3 = 99,99);
'   2. nijedan deo NIJE negativan. Raniji oblik je poslednjem kljucu davao
'      ostatak POSLE zaokruzivanja, pa kad se prethodni delovi zaokruze navise
'      preko cilja poslednji ode u minus (0,03 na 5 jednakih vrsta -> poslednja
'      vrsta -0,01). Clamp na nulu se NE sme koristiti kao popravka: razbio bi
'      invarijantu 1 (zbir bi postao 0,04).
'
' Napomena o tipu: pare se drze u Double (celobrojne vrednosti do 2^53 su tacne),
' a ne u Long -- Long bi pukao Overflow-om na iznosu preko ~21,4 miliona.
' Returns: Dictionary(kljuc -> deo iznosa, tacno na 2 decimale)
Public Function RaspodeliPoUdelima(ByVal iznos As Double, ByVal udeli As Object) As Object
    Dim outDict As Object
    Set outDict = CreateObject("Scripting.Dictionary")

    Dim ukupno As Double
    ukupno = 0

    Dim k As Variant
    If Not udeli Is Nothing Then
        For Each k In udeli.keys
            If IsNumeric(udeli(k)) Then ukupno = ukupno + CDbl(udeli(k))
        Next k
    End If

    If ukupno <= 0 Then
        outDict.Add "(Nepoznato)", ZaokruziNovac(iznos)
        Set RaspodeliPoUdelima = outDict
        Exit Function
    End If

    Dim keys As Variant
    keys = udeli.keys

    Dim n As Long
    n = UBound(keys) + 1

    Dim ciljPara As Double
    ciljPara = ZaokruziNovac(iznos) * 100
    ciljPara = Int(ciljPara + 0.5)          ' celobrojno, bez FP repa

    Dim para() As Double
    ReDim para(0 To n - 1)
    Dim ostatak() As Double
    ReDim ostatak(0 To n - 1)

    Dim sumPara As Double
    sumPara = 0

    Dim i As Long
    For i = 0 To n - 1
        Dim tezina As Double
        tezina = 0
        If IsNumeric(udeli(keys(i))) Then tezina = CDbl(udeli(keys(i)))
        If tezina < 0 Then tezina = 0       ' negativna tezina nema smisla za udeo

        Dim ideal As Double
        ideal = (iznos * 100) * (tezina / ukupno)
        If ideal < 0 Then ideal = 0         ' iznos je zbir uplata -> nikad < 0

        para(i) = Int(ideal)                ' floor nad ne-negativnim
        ostatak(i) = ideal - para(i)
        sumPara = sumPara + para(i)
    Next i

    ' Visak para (0 <= visak <= n) ide redom na najveci ostatak.
    Dim visak As Long
    visak = CLng(ciljPara - sumPara)

    Dim j As Long
    For j = 1 To visak
        Dim best As Long
        best = -1
        Dim bestOst As Double
        bestOst = -1
        For i = 0 To n - 1
            If ostatak(i) > bestOst Then
                bestOst = ostatak(i)
                best = i
            End If
        Next i
        If best < 0 Then Exit For           ' odbrana: nema vise kandidata
        para(best) = para(best) + 1
        ostatak(best) = -1                  ' iskoriscen
    Next j

    For i = 0 To n - 1
        outDict.Add CStr(keys(i)), para(i) / 100#
    Next i

    Set RaspodeliPoUdelima = outDict
End Function

' Finansijsko zaokruzivanje na 2 decimale (half-up). VBA `Round` je banker's
' rounding (2,345 -> 2,34), sto na novcu daje sistematsko odstupanje.
Public Function ZaokruziNovac(ByVal value As Double) As Double
    Dim znak As Double
    znak = 1
    If value < 0 Then znak = -1
    ZaokruziNovac = znak * Int(Abs(value) * 100 + 0.5) / 100
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
        Dim oid As String
        oid = CStr(data(i, colOtkID))
        If oid <> "" Then
            If Not dict.Exists(oid) Then dict.Add oid, 0#
            If IsNumeric(data(i, colIsplata)) Then
                dict(oid) = dict(oid) + CDbl(data(i, colIsplata))
            End If
        End If
    Next i
    
    Set BuildIsplataDictByOtkup = dict
End Function

Public Function GetOpenOtkupi(Optional ByVal kooperantID As String = "") As Variant
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
    Dim colDatum As Long, colStanica As Long
    colID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    colBrDok = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, SRC)
    colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, SRC)
    colIspl = RequireColumnIndex(TBL_OTKUP, COL_OTK_ISPLACENO, SRC)
    colDatum = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)        ' v6.18+
    colStanica = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)    ' v6.18+
    
    Dim isplataDict As Object
    Set isplataDict = BuildIsplataDictByOtkup()
    
    Dim filterByKoop As Boolean
    filterByKoop = (LenB(Trim$(kooperantID)) > 0)
    
    ' Zaehlen
    Dim count As Long, i As Long
    For i = 1 To UBound(data, 1)
        If filterByKoop Then
            If CStr(data(i, colKoop)) <> kooperantID Then GoTo NextCount
        End If
        If CStr(data(i, colIspl)) = STATUS_ISPLACENO Then GoTo NextCount
        
        Dim vrednost As Double: vrednost = 0
        If IsNumeric(data(i, colKol)) And IsNumeric(data(i, colCena)) Then
            vrednost = CDbl(data(i, colKol)) * CDbl(data(i, colCena))
        End If
        Dim isplaceno As Double: isplaceno = 0
        If isplataDict.Exists(CStr(data(i, colID))) Then isplaceno = isplataDict(CStr(data(i, colID)))
        If vrednost - isplaceno > 0 Then count = count + 1
NextCount:
    Next i
    
    If count = 0 Then
        GetOpenOtkupi = Empty
        Exit Function
    End If
    
    ' v6.18+: shape extended from 5 to 7 cols (backward-compatible)
    ' Col 1-5 stays as before; Col 6 = Datum, Col 7 = StanicaID.
    Dim result() As Variant
    ReDim result(1 To count, 1 To 7)
    Dim idx As Long
    
    For i = 1 To UBound(data, 1)
        If filterByKoop Then
            If CStr(data(i, colKoop)) <> kooperantID Then GoTo NextRow
        End If
        If CStr(data(i, colIspl)) = STATUS_ISPLACENO Then GoTo NextRow
        
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
            result(idx, 6) = data(i, colDatum)              ' v6.18+
            result(idx, 7) = CStr(data(i, colStanica))      ' v6.18+
        End If
NextRow:
    Next i
    
    GetOpenOtkupi = result
End Function
' ============================================================
' KLASIFIKACIJA TIPA (kanal placanja). Jedna definicija za sve citaoce - da se
' konstante ne nabrajaju po modulima (modIzvestaj/modStammdatenSync/modStorno).
' ============================================================

' Gotovina (blagajna). Sve ostalo je bezgotovinsko (virman/banka).
Public Function IsKesNovacTip(ByVal tip As String) As Boolean
    Select Case Trim$(tip)
        Case NOV_KES_FIRMA_OTKUPAC, NOV_KES_OTKUPAC_KOOP
            IsKesNovacTip = True
    End Select
End Function

' Avans Firma -> Otkupac (OM), nezavisno od kanala. Racuna OBA tipa jer redovi
' uvezeni iz izvoda PRE razdvajanja kanala nose KES tip - bez toga bi OM avans
' saldo, izvestaji i PWA export izgubili te iznose.
Public Function IsFirmaOtkupacAvansTip(ByVal tip As String) As Boolean
    Select Case Trim$(tip)
        Case NOV_KES_FIRMA_OTKUPAC, NOV_VIRMAN_FIRMA_OTKUPAC
            IsFirmaOtkupacAvansTip = True
    End Select
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

        ' Avans Firma->Otkupac ulazi oba kanala (kes + virman iz izvoda).
        If IsFirmaOtkupacAvansTip(CStr(data(i, colTip))) Then
            avansTotal = avansTotal + CDbl(data(i, colIsplata))
        ElseIf CStr(data(i, colTip)) = NOV_KES_OTKUPAC_KOOP Then
            isplataTotal = isplataTotal + CDbl(data(i, colIsplata))
        End If

NextRow:
    Next i

    GetOMAvansSaldo = avansTotal - isplataTotal
End Function

Public Sub ApplyAvansToOtkup(ByVal kooperantID As String, ByVal otkupID As String, _
                             Optional ByRef appliedAmount As Double)
    appliedAmount = 0

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
    ' Napomena roditelja -> split nasledjuje BIM marker (poreklo se ne gubi).
    Dim colNapomenaO As Long
    colNapomenaO = RequireColumnIndex(TBL_NOVAC, COL_NOV_NAPOMENA, SRC)

    ' Otkup-Vrednost
    Dim otkData As Variant
    otkData = GetTableData(TBL_OTKUP)
    Dim otkRows As Collection
    Set otkRows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If otkRows Is Nothing Then Exit Sub
    If otkRows.count = 0 Then Exit Sub

    ' AUD-026: target mora biti JEDNOZNACAN. OtkupID je kanonski kljuc, ali
    ' ako je u podacima dupliran, "prvi red pobedjuje" znaci da se vrednost i
    ' vlasnik citaju sa JEDNOG reda, a avans se u ledgeru vezuje samo za
    ' dvosmislen OtkupID -- akcija nad prikazanim redom B moze da se izvrsi
    ' nad redom A. Isto pravilo kao fail-closed kapije u modPrint/modFaktura
    ' (RF-08): duplikat nije "uzmi jedan" nego "ne znam target".
    ' Guard ide PRE svakog citanja target reda i PRE ijednog upisa.
    If otkRows.count > 1 Then
        Err.Raise vbObjectError + 1045, SRC, _
                  "Dupli OtkupID: " & otkupID & " (" & CStr(otkRows.count) & " redova). " & _
                  "Avans nije primenjen jer target nije jednoznacan -- pokrenite proveru " & _
                  "integriteta podataka."
    End If

    Dim r As Long: r = otkRows(1)

    ' AUD-010 / FM-0019 #5,#6: target-owner + target-active guard.
    ' Ne primeni avans na otkup drugog kooperanta niti na stornirani otkup.
    Dim colOtkKoop As Long
    colOtkKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    If StrComp(Trim$(CStr(otkData(r, colOtkKoop))), Trim$(kooperantID), vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1018, SRC, _
                  "Avans se ne moze primeniti na otkup drugog kooperanta. OtkupID=" & otkupID & _
                  "; OtkupKooperant=" & CStr(otkData(r, colOtkKoop)) & _
                  "; TrazeniKooperant=" & kooperantID
    End If

    Dim colOtkStorno As Long
    colOtkStorno = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    If colOtkStorno > 0 Then
        If UCase$(Trim$(CStr(otkData(r, colOtkStorno)))) = "DA" Then
            Err.Raise vbObjectError + 1019, SRC, _
                      "Avans se ne moze primeniti na stornirani otkup. OtkupID=" & otkupID
        End If
    End If

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
                BuildAvansSplitNapomena(CStr(data(i, colNapomenaO)), "Avans raspodela"), _
                otkupID)

            If Len(Trim$(splitNovacID)) = 0 Then
                Err.Raise vbObjectError + 1029, SRC, _
                        "Failed to create split avans row for OtkupID=" & otkupID
            End If
        End If

        appliedAmount = appliedAmount + applyAmt
        preostalo = preostalo - applyAmt
NextAvans:
    Next i

    If preostalo <= 0 Then UpdateOtkupStatus otkupID
End Sub
Public Function ApplyAvansToOtkup_TX(ByVal kooperantID As String, _
                                      ByVal otkupID As String, _
                                      Optional ByRef appliedAmount As Double) As Boolean
    appliedAmount = 0

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

    ' AUD-010 / FM-0019 #11: vrati stvarno primenjeni iznos (ByRef) uz Boolean.
    ApplyAvansToOtkup kooperantID, otkupID, appliedAmount

    tx.CommitTx
 
    Set tx = Nothing

    ApplyAvansToOtkup_TX = True
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "ApplyAvansToOtkup_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    ApplyAvansToOtkup_TX = False
    appliedAmount = 0

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
    errDesc = Err.description
    errSrc = Err.SOURCE

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
                               ByVal partnerId As String, _
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

    If Len(Trim$(partnerId)) = 0 And Len(Trim$(partner)) = 0 Then
        Err.Raise vbObjectError + 1034, sourceName, _
                  "Partner ili PartnerID je obavezan."
    End If

    If Len(Trim$(entitetTip)) = 0 Then
        Err.Raise vbObjectError + 1035, sourceName, _
                  "EntitetTip je obavezan."
    End If
End Sub


'======================================================================
' GetKooperantUnallocatedAvans
'
' Vraca sumu nedodeljenog NOV_VIRMAN_AVANS_KOOP-a za datog kooperanta.
' Nedodeljen = OtkupID je prazan (nije linkovan na konkretan blok).
'
' Read-only helper, stornirano-safe. Koristi se za info display u
' frmIsplatePregled-u (operator vidi koliko avansa kooperant ima
' pre nego sto izabere blok za isplatu).
'
' Apply Avans logika je u ApplyAvansToOtkup_TX (ide kroz frmDokumenta).
'======================================================================
Public Function GetKooperantUnallocatedAvans(ByVal kooperantID As String) As Double
    Const SRC As String = "GetKooperantUnallocatedAvans"
    
    If Len(Trim$(kooperantID)) = 0 Then Exit Function
    
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Then Exit Function
    
    Dim colKoop As Long
    Dim colTip As Long
    Dim colIsplata As Long
    Dim colOtkID As Long
    
    colKoop = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, SRC)
    colTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, SRC)
    colIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, SRC)
    colOtkID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC)
    
    Dim total As Double
    Dim i As Long
    
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colKoop))) <> Trim$(kooperantID) Then GoTo NextRow
        If CStr(data(i, colTip)) <> NOV_VIRMAN_AVANS_KOOP Then GoTo NextRow
        If Trim$(CStr(data(i, colOtkID))) <> "" Then GoTo NextRow
        
        If IsNumeric(data(i, colIsplata)) Then
            total = total + CDbl(data(i, colIsplata))
        End If
NextRow:
    Next i
    
    GetKooperantUnallocatedAvans = total
End Function

'======================================================================
' BuildKooperantUnallocatedAvansDict
'
' Single-pass dict KooperantID -> nedodeljeni avans saldo.
' Koristi se kao cache u frmIsplatePregled da izbegnemo NxM pozive.
'======================================================================
Public Function BuildKooperantUnallocatedAvansDict() As Object
    Const SRC As String = "BuildKooperantUnallocatedAvansDict"
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If Not IsArray(data) Then
        Set BuildKooperantUnallocatedAvansDict = dict
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_NOVAC)
    If Not IsArray(data) Then
        Set BuildKooperantUnallocatedAvansDict = dict
        Exit Function
    End If
    
    Dim colKoop As Long
    Dim colTip As Long
    Dim colIsplata As Long
    Dim colOtkID As Long
    
    colKoop = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, SRC)
    colTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, SRC)
    colIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, SRC)
    colOtkID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OTKUP_ID, SRC)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colTip)) <> NOV_VIRMAN_AVANS_KOOP Then GoTo NextRow
        If Trim$(CStr(data(i, colOtkID))) <> "" Then GoTo NextRow
        
        Dim kID As String
        kID = Trim$(CStr(data(i, colKoop)))
        If LenB(kID) = 0 Then GoTo NextRow
        
        If Not dict.Exists(kID) Then dict.Add kID, 0#
        
        If IsNumeric(data(i, colIsplata)) Then
            dict(kID) = dict(kID) + CDbl(data(i, colIsplata))
        End If
NextRow:
    Next i
    
    Set BuildKooperantUnallocatedAvansDict = dict
End Function
