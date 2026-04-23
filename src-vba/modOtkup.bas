Attribute VB_Name = "modOtkup"
Option Explicit

' ============================================================
' modOtkup – Aufkauf-Geschäftslogik
' Kernmodul: Erfassung Lieferant zu Station
' ============================================================

Public Function SaveOtkup_TX(ByVal datum As Date, ByVal kooperantID As String, _
                              ByVal stanicaID As String, ByVal vrstaVoca As String, _
                              ByVal sortaVoca As String, ByVal kolicina As Double, _
                              ByVal cena As Double, ByVal tipAmb As String, _
                              ByVal kolAmb As Long, ByVal vozacID As String, _
                              ByVal brDok As String, ByVal novac As Double, _
                              ByVal primalac As String, _
                              Optional ByVal klasa As String = "I", _
                              Optional ByVal parcelaID As String = "", _
                              Optional ByVal brojZbirne As String = "") As String
                              
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    
    SaveOtkup_TX = SaveOtkup(datum, kooperantID, stanicaID, vrstaVoca, _
                              sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                              vozacID, brDok, novac, primalac, klasa, _
                              parcelaID, brojZbirne)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "SaveOtkup_TX"
    tx.RollbackTx
    MsgBox "Greska pri unosu otkupa, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    SaveOtkup_TX = ""
End Function

Public Function SaveOtkup(ByVal datum As Date, ByVal kooperantID As String, _
                          ByVal stanicaID As String, ByVal vrstaVoca As String, _
                          ByVal sortaVoca As String, ByVal kolicina As Double, _
                          ByVal cena As Double, ByVal tipAmb As String, _
                          ByVal kolAmb As Long, ByVal vozacID As String, _
                          ByVal brDok As String, ByVal novac As Double, _
                          ByVal primalac As String, _
                          Optional ByVal klasa As String = "I", _
                          Optional ByVal parcelaID As String = "", _
                          Optional ByVal brojZbirne As String = "") As String
    ' Speichert einen neuen Otkup-Datensatz
    ' Returns: OtkupID oder "" bei Fehler
    
    ' Validierung
    If kooperantID = "" Then
        MsgBox "Kooperant mora biti izabran!", vbExclamation, APP_NAME
        SaveOtkup = ""
        Exit Function
    End If
    If kolicina <= 0 Then
        MsgBox "Kolicina mora biti veca od 0!", vbExclamation, APP_NAME
        SaveOtkup = ""
        Exit Function
    End If
    
    ' ID generieren
    Dim newID As String
    newID = GetNextID(TBL_OTKUP, COL_OTK_ID, "OTK-")
    
    ' Kultura-Lookup
    Dim kulturaID As String
    kulturaID = LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, "KulturaID")
    If kulturaID = "" Then kulturaID = vrstaVoca & "-" & sortaVoca
    
    ' Zeile erstellen
    Dim rowData As Variant
    rowData = Array(newID, datum, kooperantID, stanicaID, kulturaID, _
                    vrstaVoca, sortaVoca, kolicina, cena, tipAmb, _
                    kolAmb, vozacID, brDok, novac, primalac, klasa, _
                    "", brojZbirne, "", Empty, "", parcelaID)
    
    Dim result As Long
    result = AppendRow(TBL_OTKUP, rowData)
    
    If result > 0 Then
        If kolAmb > 0 Then
            TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", kooperantID, "Kooperant", , newID, DOK_TIP_OTKUP
        End If
        SaveOtkup = newID
    Else
        SaveOtkup = ""
    End If
End Function

Public Function GetOtkupByStation(ByVal stanicaID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    ' Holt alle Aufkäufe einer Station, optional nach Zeitraum gefiltert
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        GetOtkupByStation = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    ' Station-Filter
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA), "=", stanicaID
    filters.Add fp
    
    ' Datum-Filter
    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If
    
    GetOtkupByStation = FilterArray(data, filters)
End Function

Public Function GetOtkupByKooperant(ByVal kooperantID As String, _
                                    Optional ByVal datumOd As Date = 0, _
                                    Optional ByVal datumDo As Date = 0) As Variant
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        GetOtkupByKooperant = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT), "=", kooperantID
    filters.Add fp
    
    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If
    
    GetOtkupByKooperant = FilterArray(data, filters)
End Function

Public Function GetSaldoByStation(ByVal stanicaID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    ' Berechnet Saldo: Otkup - Isporuka - Banka pro Kooperant
    ' Ersetzt den alten "Saldo" Tab aus den Izveštaji
    
    Dim otkupData As Variant
    otkupData = GetOtkupByStation(stanicaID, datumOd, datumDo)
    
    ' Aggregation: Kooperant ? Roba(kg), Novac, Ambalaža
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If Not IsEmpty(otkupData) Then
        Dim colKoop As Long, colKol As Long, colNovac As Long, colAmb As Long
        colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
        colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
        colNovac = GetColumnIndex(TBL_OTKUP, COL_OTK_NOVAC)
        colAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
        
        Dim i As Long
        For i = 1 To UBound(otkupData, 1)
            Dim key As String
            key = CStr(otkupData(i, colKoop))
            If Not dict.Exists(key) Then
                dict.Add key, Array(0#, 0#, 0#) ' Kolicina, Novac, Ambalaza
            End If
            Dim vals As Variant
            vals = dict(key)
            If IsNumeric(otkupData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkupData(i, colKol))
            If IsNumeric(otkupData(i, colNovac)) Then vals(1) = vals(1) + CDbl(otkupData(i, colNovac))
            If IsNumeric(otkupData(i, colAmb)) Then vals(2) = vals(2) + CLng(otkupData(i, colAmb))
            dict(key) = vals
        Next i
    End If
    
    ' TODO: Minus Banka-Zahlungen und Isporuka abziehen
    
    ' Result als 2D-Array
    If dict.count = 0 Then
        GetSaldoByStation = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 4)
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        result(i + 1, 1) = keys(i)
        vals = dict(keys(i))
        result(i + 1, 2) = vals(0)
        result(i + 1, 3) = vals(1)
        result(i + 1, 4) = vals(2)
    Next i
    
    GetSaldoByStation = result
End Function

