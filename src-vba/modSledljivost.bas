Attribute VB_Name = "modSledljivost"
Option Explicit

' ============================================================
' modSledljivost v1.0 – Automatische Zuordnung Otkup ? Otpremnica
' ============================================================

Public Function AutoLinkOtkupOtpremnica() As Long
    ' Verknüpft alle Otkupi ohne OtpremnicaID automatisch
    ' Returns: Anzahl zugeordneter Otkupi
    
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If IsEmpty(otkupData) Then Exit Function
    otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
    If IsEmpty(otkupData) Then Exit Function
    
    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(otpData) Then Exit Function
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    If IsEmpty(otpData) Then Exit Function
    
    ' Otkup Spalten
    Dim colOtkID As Long, colOtkSt As Long, colOtkDat As Long
    Dim colOtkVoz As Long, colOtkOtpID As Long
    colOtkID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colOtkSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    colOtkDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    colOtkVoz = GetColumnIndex(TBL_OTKUP, COL_OTK_VOZAC)
    colOtkOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    
    ' Otpremnica Spalten
    Dim colOtpID As Long, colOtpSt As Long, colOtpDat As Long, colOtpVoz As Long
    colOtpID = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    colOtpSt = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA)
    colOtpDat = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM)
    colOtpVoz = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC)
    
    ' Otpremnica-Index: Key = "StanicaID|Datum|VozacID" ? Collection of OtpremnicaIDs
    Dim otpIndex As Object
    Set otpIndex = CreateObject("Scripting.Dictionary")
    
    Dim colOtpKlasa As Long
    colOtpKlasa = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA)
    
    Dim j As Long
    For j = 1 To UBound(otpData, 1)
        Dim otpKey As String
        otpKey = CStr(otpData(j, colOtpSt)) & "|" & _
                 Format$(CDate(otpData(j, colOtpDat)), "YYYY-MM-DD") & "|" & _
                 CStr(otpData(j, colOtpVoz)) & "|" & _
                 CStr(otpData(j, colOtpKlasa))
        
        If Not otpIndex.Exists(otpKey) Then
            Set otpIndex(otpKey) = New Collection
        End If
        otpIndex(otpKey).Add CStr(otpData(j, colOtpID))
    Next j
    
    Dim colOtkKlasa As Long
    colOtkKlasa = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    
    ' Otkupi zuordnen
    Dim linked As Long
    Dim i As Long
    For i = 1 To UBound(otkupData, 1)
        ' Schon zugeordnet ? skip
        If CStr(otkupData(i, colOtkOtpID)) <> "" Then GoTo NextOtkup
        
        Dim otkKey As String
        otkKey = CStr(otkupData(i, colOtkSt)) & "|" & _
                 Format$(CDate(otkupData(i, colOtkDat)), "YYYY-MM-DD") & "|" & _
                 CStr(otkupData(i, colOtkVoz)) & "|" & _
                 CStr(otkupData(i, colOtkKlasa))
        
        If otpIndex.Exists(otkKey) Then
            If otpIndex(otkKey).count = 1 Then
                ' Eindeutig ? automatisch zuordnen
                Dim otkRows As Collection
                Set otkRows = FindRows(TBL_OTKUP, COL_OTK_ID, CStr(otkupData(i, colOtkID)))
                If otkRows.count > 0 Then
                    UpdateCell TBL_OTKUP, otkRows(1), COL_OTK_OTPREMNICA_ID, otpIndex(otkKey)(1)
                    linked = linked + 1
                End If
            End If
            ' Count > 1 ? mehrdeutig, manuell
        End If
NextOtkup:
    Next i
    
    AutoLinkOtkupOtpremnica = linked
End Function

Public Function GetUnlinkedOtkupi() As Variant
    ' Returns: 2D Array aller Otkupi ohne OtpremnicaID
    ' Spalten: OtkupID, Datum, StanicaID, VozacID, Kooperant, Kolicina, VrstaVoca
    
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        GetUnlinkedOtkupi = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_OTKUP)
    If IsEmpty(data) Then
        GetUnlinkedOtkupi = Empty
        Exit Function
    End If
    
    Dim colID As Long, colDat As Long, colSt As Long, colVoz As Long
    Dim colKoop As Long, colKol As Long, colVrsta As Long, colOtpID As Long
    colID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    colSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    colVoz = GetColumnIndex(TBL_OTKUP, COL_OTK_VOZAC)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    colOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    
    ' Zählen
    Dim count As Long, i As Long
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colOtpID)) = "" Then count = count + 1
    Next i
    
    If count = 0 Then
        GetUnlinkedOtkupi = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 7)
    Dim idx As Long
    
    For i = 1 To UBound(data, 1)
        If CStr(data(i, colOtpID)) = "" Then
            idx = idx + 1
            result(idx, 1) = CStr(data(i, colID))
            result(idx, 2) = data(i, colDat)
            result(idx, 3) = CStr(data(i, colSt))
            result(idx, 4) = CStr(data(i, colVoz))
            result(idx, 5) = CStr(data(i, colKoop))
            result(idx, 6) = data(i, colKol)
            result(idx, 7) = CStr(data(i, colVrsta))
        End If
    Next i
    
    GetUnlinkedOtkupi = result
End Function

Public Function TraceByZbirna(ByVal brojZbirne As String) As Variant
    ' Komplette Rückverfolgung: Zbirna ? Otpremnice ? Otkupi ? Kooperanti
    ' Returns: 2D Array (Kooperant, Kolicina, VrstaVoca, StanicaID, OtkupDatum, OtkupID, OtpremnicaID)
    
    ' 1. Otpremnice dieser Zbirna finden
    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(otpData) Then
        TraceByZbirna = Empty
        Exit Function
    End If
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    If IsEmpty(otpData) Then
        TraceByZbirna = Empty
        Exit Function
    End If
    
    Dim colOtpID As Long, colOtpZbr As Long
    colOtpID = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    colOtpZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    
    Dim otpIDs As Object
    Set otpIDs = CreateObject("Scripting.Dictionary")
    
    Dim j As Long
    For j = 1 To UBound(otpData, 1)
        If CStr(otpData(j, colOtpZbr)) = brojZbirne Then
            otpIDs.Add CStr(otpData(j, colOtpID)), True
        End If
    Next j
    
    If otpIDs.count = 0 Then
        TraceByZbirna = Empty
        Exit Function
    End If
    
    ' 2. Otkupi mit diesen OtpremnicaIDs finden
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If IsEmpty(otkupData) Then
        TraceByZbirna = Empty
        Exit Function
    End If
    otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
    If IsEmpty(otkupData) Then
        TraceByZbirna = Empty
        Exit Function
    End If
    
    Dim colID As Long, colDat As Long, colSt As Long, colKoop As Long
    Dim colKol As Long, colVrsta As Long, colKlasa As Long
    colID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    colSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    colOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    colKlasa = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    
    ' Zählen
    Dim count As Long, i As Long
    For i = 1 To UBound(otkupData, 1)
        If otpIDs.Exists(CStr(otkupData(i, colOtpID))) Then count = count + 1
    Next i
    
    If count = 0 Then
        TraceByZbirna = Empty
        Exit Function
    End If
    
    ' Parcela-Spalte aus tblOtkup
    Dim colOtkParcela As Long
    colOtkParcela = GetColumnIndex(TBL_OTKUP, COL_OTK_PARCELA)
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 14)
    Dim idx As Long
    
    For i = 1 To UBound(otkupData, 1)
        If otpIDs.Exists(CStr(otkupData(i, colOtpID))) Then
            idx = idx + 1
            
            Dim ime As String, prezime As String
            ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", CStr(otkupData(i, colKoop)), "Ime"))
            prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", CStr(otkupData(i, colKoop)), "Prezime"))
            
            result(idx, 1) = ime & " " & prezime
            result(idx, 2) = otkupData(i, colKol)
            result(idx, 3) = CStr(otkupData(i, colVrsta))
            result(idx, 4) = CStr(otkupData(i, colSt))
            result(idx, 5) = otkupData(i, colDat)
            result(idx, 6) = CStr(otkupData(i, colID))
            result(idx, 7) = CStr(otkupData(i, colOtpID))
            result(idx, 8) = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", _
                             CStr(otkupData(i, colKoop)), COL_KOOP_BPG))
            
            ' Spalten 9-11: alles aus tblParcele
            Dim pID As String
            pID = CStr(otkupData(i, colOtkParcela))
            If pID <> "" Then
                result(idx, 9) = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, pID, COL_PAR_KAT_BROJ))
                result(idx, 10) = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, pID, COL_PAR_GGAP))
                result(idx, 13) = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, pID, COL_PAR_KULTURA))
                result(idx, 14) = CStr(LookupValue(TBL_PARCELE, COL_PAR_ID, pID, COL_PAR_POVRSINA))
            Else
                result(idx, 9) = ""
                result(idx, 10) = ""
                result(idx, 13) = ""
                result(idx, 14) = ""
            End If
            
            result(idx, 11) = CStr(otkupData(i, colKlasa))
            result(idx, 12) = pID

        End If
    Next i
    
    TraceByZbirna = result
End Function

