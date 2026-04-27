Attribute VB_Name = "modSledljivost"
Option Explicit

' ============================================================
' modSledljivost v1.1 – Automatische Zuordnung Otkup ? Otpremnica
'
' Hardening only:
'   - same matching logic
'   - same return shapes
'   - no PWA/trace flow change
'   - column guards
'   - checked updates
'   - optional TX wrapper for batch auto-link
' ============================================================

Public Function AutoLinkOtkupOtpremnica_TX() As Long
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    AutoLinkOtkupOtpremnica_TX = AutoLinkOtkupOtpremnica()

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    LogErr "modSledljivost.AutoLinkOtkupOtpremnica_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    AutoLinkOtkupOtpremnica_TX = 0
End Function

Public Function AutoLinkOtkupOtpremnica() As Long
    On Error GoTo EH

    Const SRC As String = "modSledljivost.AutoLinkOtkupOtpremnica"

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

    ' Otkup columns
    Dim colOtkID As Long
    Dim colOtkSt As Long
    Dim colOtkDat As Long
    Dim colOtkVoz As Long
    Dim colOtkOtpID As Long
    Dim colOtkKlasa As Long
    Dim colOtkZbirna As Long

    colOtkID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    colOtkSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    colOtkDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    colOtkVoz = RequireColumnIndex(TBL_OTKUP, COL_OTK_VOZAC, SRC)
    colOtkOtpID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, SRC)
    colOtkKlasa = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, SRC)
    colOtkZbirna = RequireColumnIndex(TBL_OTKUP, "BrojZbirne", SRC)

    ' Otpremnica columns
    Dim colOtpID As Long
    Dim colOtpSt As Long
    Dim colOtpDat As Long
    Dim colOtpVoz As Long
    Dim colOtpKlasa As Long
    Dim colOtpZbirna As Long

    colOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    colOtpSt = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, SRC)
    colOtpDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    colOtpVoz = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    colOtpKlasa = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, SRC)
    colOtpZbirna = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)

    Dim strictIndex As Object
    Dim fallbackIndex As Object

    Set strictIndex = CreateObject("Scripting.Dictionary")
    Set fallbackIndex = CreateObject("Scripting.Dictionary")

    Dim j As Long

    For j = 1 To UBound(otpData, 1)

        Dim otpID As String
        Dim otpZbirna As String

        otpID = Trim$(CStr(otpData(j, colOtpID)))
        otpZbirna = Trim$(CStr(otpData(j, colOtpZbirna)))

        If Len(otpID) > 0 Then

            If Len(otpZbirna) > 0 Then
                AddAutoLinkCandidate _
                    strictIndex, _
                    BuildAutoLinkKey( _
                        otpData(j, colOtpSt), _
                        otpData(j, colOtpDat), _
                        otpData(j, colOtpVoz), _
                        otpData(j, colOtpKlasa), _
                        otpZbirna, _
                        True), _
                    otpID
            Else
                ' Legacy fallback only when Otpremnica has no BrojZbirne.
                AddAutoLinkCandidate _
                    fallbackIndex, _
                    BuildAutoLinkKey( _
                        otpData(j, colOtpSt), _
                        otpData(j, colOtpDat), _
                        otpData(j, colOtpVoz), _
                        otpData(j, colOtpKlasa), _
                        "", _
                        False), _
                    otpID
            End If

        End If

    Next j

    Dim linked As Long
    Dim i As Long

    For i = 1 To UBound(otkupData, 1)

        If Len(Trim$(CStr(otkupData(i, colOtkOtpID)))) > 0 Then
            GoTo NextOtkup
        End If

        Dim otkupID As String
        Dim otkupZbirna As String
        Dim key As String
        Dim targetOtpID As String

        otkupID = Trim$(CStr(otkupData(i, colOtkID)))
        otkupZbirna = Trim$(CStr(otkupData(i, colOtkZbirna)))

        If Len(otkupID) = 0 Then GoTo NextOtkup

        If Len(otkupZbirna) > 0 Then

            ' New/canonical path:
            ' BrojZbirne exists, so link only to same BrojZbirne.
            key = BuildAutoLinkKey( _
                    otkupData(i, colOtkSt), _
                    otkupData(i, colOtkDat), _
                    otkupData(i, colOtkVoz), _
                    otkupData(i, colOtkKlasa), _
                    otkupZbirna, _
                    True)

            targetOtpID = GetUniqueAutoLinkTarget(strictIndex, key)

        Else

            ' Legacy fallback:
            ' only when Otkup has no BrojZbirne.
            key = BuildAutoLinkKey( _
                    otkupData(i, colOtkSt), _
                    otkupData(i, colOtkDat), _
                    otkupData(i, colOtkVoz), _
                    otkupData(i, colOtkKlasa), _
                    "", _
                    False)

            targetOtpID = GetUniqueAutoLinkTarget(fallbackIndex, key)

        End If

        If Len(targetOtpID) > 0 Then
            Dim otkRows As Collection
            Set otkRows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)

            If Not otkRows Is Nothing Then
                If otkRows.count > 0 Then
                    RequireUpdateCell TBL_OTKUP, otkRows(1), _
                                      COL_OTK_OTPREMNICA_ID, targetOtpID, SRC
                    linked = linked + 1
                End If
            End If
        End If

NextOtkup:
    Next i

    AutoLinkOtkupOtpremnica = linked
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.Description
End Function

Private Function BuildAutoLinkKey(ByVal stanicaID As Variant, _
                                  ByVal datumValue As Variant, _
                                  ByVal vozacID As Variant, _
                                  ByVal klasa As Variant, _
                                  ByVal brojZbirne As String, _
                                  ByVal includeZbirna As Boolean) As String
    If Not IsDate(datumValue) Then Exit Function

    BuildAutoLinkKey = _
        UCase$(Trim$(CStr(stanicaID))) & "|" & _
        Format$(CDate(datumValue), "yyyy-mm-dd") & "|" & _
        UCase$(Trim$(CStr(vozacID))) & "|" & _
        UCase$(Trim$(CStr(klasa)))

    If includeZbirna Then
        BuildAutoLinkKey = BuildAutoLinkKey & "|" & _
                           UCase$(Trim$(brojZbirne))
    End If
End Function

Private Sub AddAutoLinkCandidate(ByVal index As Object, _
                                 ByVal key As String, _
                                 ByVal otpID As String)
    If Len(Trim$(key)) = 0 Then Exit Sub
    If Len(Trim$(otpID)) = 0 Then Exit Sub

    Dim bucket As Collection

    If Not index.Exists(key) Then
        Set bucket = New Collection
        index.Add key, bucket
    End If

    index(key).Add otpID
End Sub

Private Function GetUniqueAutoLinkTarget(ByVal index As Object, _
                                         ByVal key As String) As String
    If Len(Trim$(key)) = 0 Then Exit Function

    If Not index.Exists(key) Then Exit Function

    If index(key).count = 1 Then
        GetUniqueAutoLinkTarget = CStr(index(key)(1))
    End If
End Function

Public Function GetUnlinkedOtkupi() As Variant
    On Error GoTo EH

    ' Returns: 2D Array aller Otkupi ohne OtpremnicaID
    ' Spalten ostaju iste:
    '   1 OtkupID
    '   2 Datum
    '   3 StanicaID
    '   4 VozacID
    '   5 Kooperant
    '   6 Kolicina
    '   7 VrstaVoca

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

    Dim colID As Long
    Dim colDat As Long
    Dim colSt As Long
    Dim colVoz As Long
    Dim colKoop As Long
    Dim colKol As Long
    Dim colVrsta As Long
    Dim colOtpID As Long

    colID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, _
                               "modSledljivost.GetUnlinkedOtkupi")
    colDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                                "modSledljivost.GetUnlinkedOtkupi")
    colSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, _
                               "modSledljivost.GetUnlinkedOtkupi")
    colVoz = RequireColumnIndex(TBL_OTKUP, COL_OTK_VOZAC, _
                                "modSledljivost.GetUnlinkedOtkupi")
    colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, _
                                 "modSledljivost.GetUnlinkedOtkupi")
    colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, _
                                "modSledljivost.GetUnlinkedOtkupi")
    colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, _
                                  "modSledljivost.GetUnlinkedOtkupi")
    colOtpID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, _
                                  "modSledljivost.GetUnlinkedOtkupi")

    Dim count As Long
    Dim i As Long

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
    Exit Function

EH:
    LogErr "modSledljivost.GetUnlinkedOtkupi"
    GetUnlinkedOtkupi = Empty
End Function

Public Function TraceByZbirna(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    ' Komplette Rückverfolgung:
    '   Zbirna ? Otpremnice ? Otkupi ? Kooperanti
    '
    ' Return shape ostaje isti:
    '   1  Kooperant
    '   2  Kolicina
    '   3  VrstaVoca
    '   4  StanicaID
    '   5  OtkupDatum
    '   6  OtkupID
    '   7  OtpremnicaID
    '   8  BPG
    '   9  KatBroj
    '   10 GGAP
    '   11 Klasa
    '   12 ParcelaID
    '   13 Kultura
    '   14 Povrsina

    If Trim$(brojZbirne) = "" Then
        TraceByZbirna = Empty
        Exit Function
    End If

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

    Dim colOtpID As Long
    Dim colOtpZbr As Long

    colOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, _
                                  "modSledljivost.TraceByZbirna")
    colOtpZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                   "modSledljivost.TraceByZbirna")

    Dim otpIDs As Object
    Set otpIDs = CreateObject("Scripting.Dictionary")

    Dim j As Long
    Dim otpID As String

    For j = 1 To UBound(otpData, 1)
        If CStr(otpData(j, colOtpZbr)) = brojZbirne Then
            otpID = CStr(otpData(j, colOtpID))

            If otpID <> "" Then
                If Not otpIDs.Exists(otpID) Then
                    otpIDs.Add otpID, True
                End If
            End If
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

    Dim colID As Long
    Dim colDat As Long
    Dim colSt As Long
    Dim colKoop As Long
    Dim colKol As Long
    Dim colVrsta As Long
    Dim colOtkOtpID As Long
    Dim colKlasa As Long
    Dim colOtkParcela As Long

    colID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, _
                               "modSledljivost.TraceByZbirna")
    colDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                                "modSledljivost.TraceByZbirna")
    colSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, _
                               "modSledljivost.TraceByZbirna")
    colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, _
                                 "modSledljivost.TraceByZbirna")
    colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, _
                                "modSledljivost.TraceByZbirna")
    colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, _
                                  "modSledljivost.TraceByZbirna")
    colOtkOtpID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, _
                                     "modSledljivost.TraceByZbirna")
    colKlasa = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, _
                                  "modSledljivost.TraceByZbirna")
    colOtkParcela = RequireColumnIndex(TBL_OTKUP, COL_OTK_PARCELA, _
                                       "modSledljivost.TraceByZbirna")

    ' Lookup table guards — ne menjaju logiku, samo fail-fast za schema drift
    RequireColumnIndex TBL_KOOPERANTI, "KooperantID", "modSledljivost.TraceByZbirna"
    RequireColumnIndex TBL_KOOPERANTI, "Ime", "modSledljivost.TraceByZbirna"
    RequireColumnIndex TBL_KOOPERANTI, "Prezime", "modSledljivost.TraceByZbirna"
    RequireColumnIndex TBL_KOOPERANTI, COL_KOOP_BPG, "modSledljivost.TraceByZbirna"

    RequireColumnIndex TBL_PARCELE, COL_PAR_ID, "modSledljivost.TraceByZbirna"
    RequireColumnIndex TBL_PARCELE, COL_PAR_KAT_BROJ, "modSledljivost.TraceByZbirna"
    RequireColumnIndex TBL_PARCELE, COL_PAR_GGAP, "modSledljivost.TraceByZbirna"
    RequireColumnIndex TBL_PARCELE, COL_PAR_KULTURA, "modSledljivost.TraceByZbirna"
    RequireColumnIndex TBL_PARCELE, COL_PAR_POVRSINA, "modSledljivost.TraceByZbirna"

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(otkupData, 1)
        If otpIDs.Exists(CStr(otkupData(i, colOtkOtpID))) Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        TraceByZbirna = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To 14)

    Dim idx As Long
    Dim ime As String
    Dim prezime As String
    Dim koopID As String
    Dim pID As String

    For i = 1 To UBound(otkupData, 1)
        If otpIDs.Exists(CStr(otkupData(i, colOtkOtpID))) Then
            idx = idx + 1

            koopID = CStr(otkupData(i, colKoop))

            ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime"))
            prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime"))

            result(idx, 1) = ime & " " & prezime
            result(idx, 2) = otkupData(i, colKol)
            result(idx, 3) = CStr(otkupData(i, colVrsta))
            result(idx, 4) = CStr(otkupData(i, colSt))
            result(idx, 5) = otkupData(i, colDat)
            result(idx, 6) = CStr(otkupData(i, colID))
            result(idx, 7) = CStr(otkupData(i, colOtkOtpID))
            result(idx, 8) = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", _
                                             koopID, COL_KOOP_BPG))

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
    Exit Function

EH:
    LogErr "modSledljivost.TraceByZbirna"
    TraceByZbirna = Empty
End Function

