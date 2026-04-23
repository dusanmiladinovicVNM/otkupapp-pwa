Attribute VB_Name = "modAgrohemija"
Option Explicit

Public Function GetParceleByKooperant(ByVal kooperantID As String) As Variant
    Dim data As Variant
    data = GetTableData(TBL_PARCELE)
    If IsEmpty(data) Then
        GetParceleByKooperant = Empty
        Exit Function
    End If
    
    Dim colKoop As Long, colID As Long, colKat As Long
    Dim colOpstina As Long, colKultura As Long, colPovrsina As Long
    colID = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
    colKoop = GetColumnIndex(TBL_PARCELE, COL_PAR_KOOP)
    colKat = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_BROJ)
    colOpstina = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_OPSTINA)
    colKultura = GetColumnIndex(TBL_PARCELE, COL_PAR_KULTURA)
    colPovrsina = GetColumnIndex(TBL_PARCELE, COL_PAR_POVRSINA)
    
    Dim count As Long, i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID Then count = count + 1
    Next i
    
    If count = 0 Then
        GetParceleByKooperant = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 6)
    Dim idx As Long
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID Then
            idx = idx + 1
            result(idx, 1) = CStr(data(i, colID))
            result(idx, 2) = CStr(data(i, colKat))
            result(idx, 3) = CStr(data(i, colOpstina))
            result(idx, 4) = CStr(data(i, colKultura))
            result(idx, 5) = CDbl(data(i, colPovrsina))
            result(idx, 6) = CStr(data(i, colKat)) & " " & CStr(data(i, colOpstina)) & _
                             " (" & CStr(data(i, colKultura)) & ", " & _
                             Format$(CDbl(data(i, colPovrsina)), "0.00") & " ha)"
        End If
    Next i
    
    GetParceleByKooperant = result
End Function

Public Function CalculatePreporuka(ByVal artikalID As String, _
                                   ByVal povrsinaHa As Double) As Double
    Dim raw As Variant
    raw = LookupValue(TBL_ARTIKLI, COL_ART_ID, Trim$(artikalID), COL_ART_DOZA)

    If IsError(raw) Or IsEmpty(raw) Or IsNull(raw) Or Not IsNumeric(raw) Then
        CalculatePreporuka = 0
        Exit Function
    End If

    CalculatePreporuka = CDbl(raw) * povrsinaHa
End Function

Public Function SaveMagacin(ByVal datum As Date, ByVal artikalID As String, _
                             ByVal tip As String, ByVal kolicina As Double, _
                             Optional ByVal kooperantID As String = "", _
                             Optional ByVal parcelaID As String = "", _
                             Optional ByVal brojDok As String = "", _
                             Optional ByVal napomena As String = "", _
                             Optional ByVal dobavljacID As String = "") As String
    Dim newID As String
    newID = GetNextID(TBL_MAGACIN, COL_MAG_ID, "MAG-")
    
    ' Cena aus Artikli
    Dim cena As Double
    Dim cenaStr As String
    cenaStr = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, artikalID, COL_ART_CENA))
    If IsNumeric(cenaStr) Then cena = CDbl(cenaStr)
    If IsNumeric(kolicina) Then kolicina = CDbl(kolicina)
    Dim vrednost As Double
    vrednost = kolicina * cena
    
    Dim rowData As Variant
    rowData = Array(newID, datum, artikalID, tip, kolicina, _
                    kooperantID, parcelaID, brojDok, cena, vrednost, _
                    napomena, "", dobavljacID)
    
    If AppendRow(TBL_MAGACIN, rowData) > 0 Then
        SaveMagacin = newID
    Else
        SaveMagacin = ""
    End If
End Function

Public Function SaveMagacin_TX(ByVal datum As Date, ByVal artikalID As String, _
                                ByVal tip As String, ByVal kolicina As Double, _
                                Optional ByVal kooperantID As String = "", _
                                Optional ByVal parcelaID As String = "", _
                                Optional ByVal brojDok As String = "", _
                                Optional ByVal napomena As String = "", _
                                Optional ByVal dobavljacID As String = "") As String
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_MAGACIN
    
    SaveMagacin_TX = SaveMagacin(datum, artikalID, tip, kolicina, kooperantID, parcelaID, brojDok, napomena, dobavljacID)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "SaveMagacin_TX"
    tx.RollbackTx
    MsgBox "Greska pri unosu magacina, promene vracene: " & Err.Description, vbCritical, APP_NAME
    SaveMagacin_TX = ""
End Function

Public Function GetMagacinStanje() As Variant
    ' Returns: 2D Array (ArtikalID, Naziv, Tip, JM, Ulaz, Izlaz, Stanje)
    Dim data As Variant
    data = GetTableData(TBL_MAGACIN)
    If IsEmpty(data) Then
        GetMagacinStanje = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_MAGACIN)
    If IsEmpty(data) Then
        GetMagacinStanje = Empty
        Exit Function
    End If
    
    Dim colArt As Long, colTip As Long, colKol As Long
    colArt = GetColumnIndex(TBL_MAGACIN, COL_MAG_ARTIKAL)
    colTip = GetColumnIndex(TBL_MAGACIN, COL_MAG_TIP)
    colKol = GetColumnIndex(TBL_MAGACIN, COL_MAG_KOLICINA)
    
    ' Dict: ArtikalID ? Array(Ulaz, Izlaz)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim artID As String
        artID = CStr(data(i, colArt))
        If Not dict.Exists(artID) Then dict.Add artID, Array(0#, 0#)
        
        Dim vals As Variant: vals = dict(artID)
        If IsNumeric(data(i, colKol)) Then
            If CStr(data(i, colTip)) = MAG_ULAZ Then
                vals(0) = vals(0) + CDbl(data(i, colKol))
            ElseIf CStr(data(i, colTip)) = MAG_IZLAZ Then
                vals(1) = vals(1) + CDbl(data(i, colKol))
            End If
        End If
        dict(artID) = vals
    Next i
    
    If dict.count = 0 Then
        GetMagacinStanje = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 7)
    Dim keys As Variant: keys = dict.keys
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, keys(i), COL_ART_NAZIV))
        result(i + 1, 3) = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, keys(i), COL_ART_TIP))
        result(i + 1, 4) = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, keys(i), COL_ART_JM))
        result(i + 1, 5) = vals(0)          ' Ulaz
        result(i + 1, 6) = vals(1)          ' Izlaz
        result(i + 1, 7) = vals(0) - vals(1) ' Stanje
    Next i
    
    GetMagacinStanje = result
End Function

Public Function ReportIzdavanjePoKooperantu(Optional ByVal datumOd As Date = 0, _
                                             Optional ByVal datumDo As Date = 0) As Variant
    Dim data As Variant
    data = GetTableData(TBL_MAGACIN)
    If IsEmpty(data) Then
        ReportIzdavanjePoKooperantu = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_MAGACIN)
    If IsEmpty(data) Then
        ReportIzdavanjePoKooperantu = Empty
        Exit Function
    End If
    
    Dim colKoop As Long, colTip As Long, colVrednost As Long, colDat As Long
    colKoop = GetColumnIndex(TBL_MAGACIN, COL_MAG_KOOP)
    colTip = GetColumnIndex(TBL_MAGACIN, COL_MAG_TIP)
    colVrednost = GetColumnIndex(TBL_MAGACIN, COL_MAG_VREDNOST)
    colDat = GetColumnIndex(TBL_MAGACIN, COL_MAG_DATUM)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colTip)) <> MAG_IZLAZ Then GoTo NextRow
        If CStr(data(i, colKoop)) = "" Then GoTo NextRow
        
        If datumOd > 0 Then
            If IsDate(data(i, colDat)) Then
                If CDate(data(i, colDat)) < datumOd Or CDate(data(i, colDat)) > datumDo Then GoTo NextRow
            End If
        End If
        
        Dim koopID As String
        koopID = CStr(data(i, colKoop))
        If Not dict.Exists(koopID) Then dict.Add koopID, 0#
        If IsNumeric(data(i, colVrednost)) Then
            dict(koopID) = dict(koopID) + CDbl(data(i, colVrednost))
        End If
NextRow:
    Next i
    
    If dict.count = 0 Then
        ReportIzdavanjePoKooperantu = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 3)
    Dim keys As Variant: keys = dict.keys
    Dim totalVrednost As Double
    
    For i = 0 To dict.count - 1
        Dim ime As String, prezime As String
        ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", keys(i), "Ime"))
        prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", keys(i), "Prezime"))
        
        result(i + 1, 1) = ime & " " & prezime
        result(i + 1, 2) = keys(i)
        result(i + 1, 3) = dict(keys(i))
        totalVrednost = totalVrednost + dict(keys(i))
    Next i
    
    result(dict.count + 1, 1) = "UKUPNO"
    result(dict.count + 1, 2) = ""
    result(dict.count + 1, 3) = totalVrednost
    
    ReportIzdavanjePoKooperantu = result
End Function

Public Function ReportStanjePoDoabvljacu() As Variant
    Dim data As Variant
    data = GetTableData(TBL_MAGACIN)
    If IsEmpty(data) Then
        ReportStanjePoDoabvljacu = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_MAGACIN)
    If IsEmpty(data) Then
        ReportStanjePoDoabvljacu = Empty
        Exit Function
    End If
    
    Dim colDobavljac As Long, colTip As Long, colVrednost As Long
    colDobavljac = GetColumnIndex(TBL_MAGACIN, COL_MAG_DOBAVLJAC)
    colTip = GetColumnIndex(TBL_MAGACIN, COL_MAG_TIP)
    colVrednost = GetColumnIndex(TBL_MAGACIN, COL_MAG_VREDNOST)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colTip)) <> MAG_ULAZ Then GoTo NextDob
        Dim dobID As String
        dobID = CStr(data(i, colDobavljac))
        If dobID = "" Then dobID = "(Nepoznat)"
        
        If Not dict.Exists(dobID) Then dict.Add dobID, 0#
        If IsNumeric(data(i, colVrednost)) Then
            dict(dobID) = dict(dobID) + CDbl(data(i, colVrednost))
        End If
NextDob:
    Next i
    
    If dict.count = 0 Then
        ReportStanjePoDoabvljacu = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 2)
    Dim keys As Variant: keys = dict.keys
    Dim total As Double
    
    For i = 0 To dict.count - 1
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = dict(keys(i))
        total = total + dict(keys(i))
    Next i
    
    result(dict.count + 1, 1) = "UKUPNO"
    result(dict.count + 1, 2) = total
    
    ReportStanjePoDoabvljacu = result
End Function

Public Function GetAgrohemijaDug(ByVal kooperantID As String) As Double
    ' Summe aller Izlaz-Vrednosti für diesen Kooperant
    Dim data As Variant
    data = GetTableData(TBL_MAGACIN)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_MAGACIN)
    If IsEmpty(data) Then Exit Function
    
    Dim colKoop As Long, colTip As Long, colVrednost As Long
    colKoop = GetColumnIndex(TBL_MAGACIN, COL_MAG_KOOP)
    colTip = GetColumnIndex(TBL_MAGACIN, COL_MAG_TIP)
    colVrednost = GetColumnIndex(TBL_MAGACIN, COL_MAG_VREDNOST)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colKoop)) = kooperantID And _
           CStr(data(i, colTip)) = MAG_IZLAZ Then
            If IsNumeric(data(i, colVrednost)) Then
                GetAgrohemijaDug = GetAgrohemijaDug + CDbl(data(i, colVrednost))
            End If
        End If
    Next i
End Function

