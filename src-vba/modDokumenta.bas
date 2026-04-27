Attribute VB_Name = "modDokumenta"

Option Explicit

' ============================================================
' modDokumenta – Otpremnica, Zbirna, Prijemnica
' Dokumentenfluss: Otkup zu Otpremnica zu Zbirna zu Prijemnica zu Faktura
' ============================================================

' ============================================================
' OTPREMNICA – Station gibt Ware an Fahrer
' ============================================================
Public Function SaveOtpremnicaMulti_TX(ByVal datum As Date, _
                                       ByVal stanicaID As String, _
                                       ByVal vozacID As String, _
                                       ByVal brojOtp As String, _
                                       ByVal brojZbirne As String, _
                                       ByVal vrsta As String, _
                                       ByVal sorta As String, _
                                       ByVal kolicinaI As Double, _
                                       ByVal cenaI As Double, _
                                       ByVal tipAmb As String, _
                                       ByVal kolAmb As Long, _
                                       Optional ByVal hasKlasaII As Boolean = False, _
                                       Optional ByVal kolicinaII As Double = 0, _
                                       Optional ByVal cenaII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    Dim resultI As String
    resultI = SaveOtpremnica( _
        datum, _
        stanicaID, _
        vozacID, _
        brojOtp, _
        brojZbirne, _
        vrsta, _
        sorta, _
        kolicinaI, _
        cenaI, _
        tipAmb, _
        kolAmb, _
        KLASA_I)

    If resultI = "" Then
        Err.Raise vbObjectError + 1101, "SaveOtpremnicaMulti_TX", _
                  "SaveOtpremnica Klasa I fehlgeschlagen"
    End If

    Dim resultII As String

    If hasKlasaII Then
        resultII = SaveOtpremnica( _
            datum, _
            stanicaID, _
            vozacID, _
            brojOtp, _
            brojZbirne, _
            vrsta, _
            sorta, _
            kolicinaII, _
            cenaII, _
            tipAmb, _
            0, _
            KLASA_II)

        If resultII = "" Then
            Err.Raise vbObjectError + 1102, "SaveOtpremnicaMulti_TX", _
                      "SaveOtpremnica Klasa II fehlgeschlagen"
        End If

        SaveOtpremnicaMulti_TX = resultI & " + " & resultII
    Else
        SaveOtpremnicaMulti_TX = resultI
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    LogErr "SaveOtpremnicaMulti_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtpremnicaMulti_TX = ""
End Function

Public Function SaveOtpremnica_TX(ByVal datum As Date, ByVal stanicaID As String, _
                                   ByVal vozacID As String, ByVal brojOtp As String, _
                                   ByVal brojZbirne As String, ByVal vrsta As String, _
                                   ByVal sorta As String, ByVal kolicina As Double, _
                                   ByVal cena As Double, ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, _
                                   Optional ByVal klasa As String = "I") As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    SaveOtpremnica_TX = SaveOtpremnica(datum, stanicaID, vozacID, brojOtp, _
                                        brojZbirne, vrsta, sorta, kolicina, _
                                        cena, tipAmb, kolAmb, klasa)

    If SaveOtpremnica_TX = "" Then
        Err.Raise vbObjectError + 1001, "SaveOtpremnica_TX", _
                  "SaveOtpremnica fehlgeschlagen"
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    LogErr "SaveOtpremnica_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtpremnica_TX = ""
End Function

Public Function SaveOtpremnica(ByVal datum As Date, ByVal stanicaID As String, _
                               ByVal vozacID As String, ByVal brojOtp As String, _
                               ByVal brojZbirne As String, ByVal vrsta As String, _
                               ByVal sorta As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, _
                               Optional ByVal klasa As String = "I") As String
    
    If stanicaID = "" Or vozacID = "" Or kolicina <= 0 Then
        Err.Raise vbObjectError + 1002, "SaveOtpremnica", _
                  "Stanica, vozac i kolicina su obavezni!"
    End If
    
    Dim newID As String
    newID = GetNextID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-")
    
    Dim rowData As Variant
    rowData = Array(newID, datum, stanicaID, vozacID, brojOtp, _
                    brojZbirne, vrsta, sorta, kolicina, cena, tipAmb, kolAmb, klasa)
    
    If AppendRow(TBL_OTPREMNICA, rowData) > 0 Then
        If kolAmb > 0 Then
            TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", stanicaID, "Stanica", vozacID, newID, DOK_TIP_OTPREMNICA
        End If
        SaveOtpremnica = newID
    Else
        Err.Raise vbObjectError + 1003, "SaveOtpremnica", _
                  "AppendRow fehlgeschlagen für tblOtpremnica"
    End If
End Function

Public Function GetOtpremniceByZbirna(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByZbirna = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
            "modDokumenta.GetOtpremniceByZbirna"), "=", brojZbirne
    filters.Add fp

    GetOtpremniceByZbirna = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetOtpremniceByZbirna"
    GetOtpremniceByZbirna = Empty
End Function

Public Function GetOtpremniceByStation(ByVal stanicaID As String, _
                                       Optional ByVal datumOd As Date = 0, _
                                       Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByStation = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, _
            "modDokumenta.GetOtpremniceByStation"), "=", stanicaID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, _
                "modDokumenta.GetOtpremniceByStation"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtpremniceByStation = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetOtpremniceByStation"
    GetOtpremniceByStation = Empty
End Function

' ============================================================
' ZBIRNA – Gesamtdokument Fahrer
' ============================================================
Public Function SaveZbirnaMulti_TX(ByVal datum As Date, _
                                   ByVal vozacID As String, _
                                   ByVal brojZbirne As String, _
                                   ByVal kupacID As String, _
                                   ByVal hladnjaca As String, _
                                   ByVal pogon As String, _
                                   ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, _
                                   ByVal ukupnoKolI As Double, _
                                   ByVal tipAmb As String, _
                                   ByVal ukupnoAmb As Long, _
                                   Optional ByVal hasKlasaII As Boolean = False, _
                                   Optional ByVal ukupnoKolII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    Dim resultI As String
    resultI = SaveZbirna( _
        datum, _
        vozacID, _
        brojZbirne, _
        kupacID, _
        hladnjaca, _
        pogon, _
        vrstaVoca, _
        sortaVoca, _
        ukupnoKolI, _
        tipAmb, _
        ukupnoAmb, _
        KLASA_I)

    If resultI = "" Then
        Err.Raise vbObjectError + 1201, "SaveZbirnaMulti_TX", _
                  "SaveZbirna Klasa I fehlgeschlagen"
    End If

    Dim resultII As String

    If hasKlasaII Then
        resultII = SaveZbirna( _
            datum, _
            vozacID, _
            brojZbirne, _
            kupacID, _
            hladnjaca, _
            pogon, _
            vrstaVoca, _
            sortaVoca, _
            ukupnoKolII, _
            tipAmb, _
            0, _
            KLASA_II)

        If resultII = "" Then
            Err.Raise vbObjectError + 1202, "SaveZbirnaMulti_TX", _
                      "SaveZbirna Klasa II fehlgeschlagen"
        End If

        SaveZbirnaMulti_TX = resultI & " + " & resultII
    Else
        SaveZbirnaMulti_TX = resultI
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    LogErr "SaveZbirnaMulti_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveZbirnaMulti_TX = ""
End Function

Public Function SaveZbirna_TX(ByVal datum As Date, ByVal vozacID As String, _
                               ByVal brojZbirne As String, ByVal kupacID As String, _
                               ByVal hladnjaca As String, ByVal pogon As String, _
                               ByVal vrstaVoca As String, ByVal sortaVoca As String, _
                               ByVal ukupnoKol As Double, ByVal tipAmb As String, _
                               ByVal ukupnoAmb As Long, _
                               Optional ByVal klasa As String = "I") As String
    Dim tx As New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    SaveZbirna_TX = SaveZbirna(datum, vozacID, brojZbirne, kupacID, _
                                hladnjaca, pogon, vrstaVoca, sortaVoca, _
                                ukupnoKol, tipAmb, ukupnoAmb, klasa)

    If SaveZbirna_TX = "" Then
        Err.Raise vbObjectError + 1004, "SaveZbirna_TX", _
                  "SaveZbirna fehlgeschlagen"
    End If

    tx.CommitTx
    Exit Function

EH:
    LogErr "SaveZbirna_TX"

    On Error Resume Next
    tx.RollbackTx
    On Error GoTo 0

    SaveZbirna_TX = ""
End Function

Public Function SaveZbirna(ByVal datum As Date, ByVal vozacID As String, _
                           ByVal brojZbirne As String, ByVal kupacID As String, _
                           ByVal hladnjaca As String, ByVal pogon As String, _
                           ByVal vrstaVoca As String, ByVal sortaVoca As String, _
                           ByVal ukupnoKol As Double, ByVal tipAmb As String, _
                           ByVal ukupnoAmb As Long, _
                           Optional ByVal klasa As String = "I") As String
    On Error GoTo EH

    If vozacID = "" Or brojZbirne = "" Then
        Err.Raise vbObjectError + 1007, "SaveZbirna", _
                  "Vozac i broj zbirne su obavezni."
    End If

    If ukupnoKol <= 0 Then
        Err.Raise vbObjectError + 1008, "SaveZbirna", _
                  "Ukupna kolicina mora biti veca od nule."
    End If

    Dim newID As String
    newID = GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-")

    If newID = "" Then
        Err.Raise vbObjectError + 1009, "SaveZbirna", _
                  "GetNextID nije vratio ZbirnaID."
    End If

    Dim rowData As Variant
    rowData = Array(newID, datum, vozacID, brojZbirne, kupacID, _
                    hladnjaca, pogon, vrstaVoca, sortaVoca, _
                    ukupnoKol, tipAmb, ukupnoAmb, klasa)

    If AppendRow(TBL_ZBIRNA, rowData) > 0 Then
        SaveZbirna = newID
    Else
        Err.Raise vbObjectError + 1010, "SaveZbirna", _
                  "AppendRow fehlgeschlagen für tblZbirna."
    End If

    Exit Function

EH:
    LogErr "SaveZbirna"
    SaveZbirna = ""
End Function

Public Function GetZbirnaByKupac(ByVal kupacID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)

    If IsEmpty(data) Then
        GetZbirnaByKupac = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, _
            "modDokumenta.GetZbirnaByKupac"), "=", kupacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, _
                "modDokumenta.GetZbirnaByKupac"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetZbirnaByKupac = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetZbirnaByKupac"
    GetZbirnaByKupac = Empty
End Function

' ============================================================
' ZBIRNA VALIDIERUNG
' ============================================================
Public Function ValidateZbirna(ByVal brojZbirne As String) As Variant
    ' Prüft Summe Otpremnice vs Zbirna
    ' Returns: Array(SumaOtpKg, ZbirnaKg, RazlikaKg, ValidKg,
    '                SumaOtpAmb, ZbirnaAmb, RazlikaAmb)
    On Error GoTo EH

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    Dim sumaOtpKg As Double
    Dim sumaOtpAmb As Long

    If Not IsEmpty(otpData) Then
        Dim colKol As Long
        Dim colAmb As Long

        colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                    "modDokumenta.ValidateZbirna")
        colAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, _
                                    "modDokumenta.ValidateZbirna")

        Dim i As Long
        For i = 1 To UBound(otpData, 1)
            If IsNumeric(otpData(i, colKol)) Then sumaOtpKg = sumaOtpKg + CDbl(otpData(i, colKol))
            If IsNumeric(otpData(i, colAmb)) Then sumaOtpAmb = sumaOtpAmb + CLng(otpData(i, colAmb))
        Next i
    End If

    Dim zbirnaKg As Double
    Dim zbirnaAmb As Long

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

    If Not IsEmpty(zbrData) Then
        Dim colZbrBroj As Long
        Dim colZbrKol As Long
        Dim colZbrAmb As Long

        colZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                        "modDokumenta.ValidateZbirna")
        colZbrKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                       "modDokumenta.ValidateZbirna")
        colZbrAmb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB, _
                                       "modDokumenta.ValidateZbirna")

        For i = 1 To UBound(zbrData, 1)
            If CStr(zbrData(i, colZbrBroj)) = brojZbirne Then
                If IsNumeric(zbrData(i, colZbrKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(i, colZbrKol))
                If IsNumeric(zbrData(i, colZbrAmb)) Then zbirnaAmb = zbirnaAmb + CLng(zbrData(i, colZbrAmb))
            End If
        Next i
    End If

    ValidateZbirna = Array(sumaOtpKg, zbirnaKg, sumaOtpKg - zbirnaKg, _
                           (Abs(sumaOtpKg - zbirnaKg) < 0.01), _
                           sumaOtpAmb, zbirnaAmb, sumaOtpAmb - zbirnaAmb)

    Exit Function

EH:
    LogErr "modDokumenta.ValidateZbirna"
    ValidateZbirna = Array(0#, 0#, 0#, False, 0&, 0&, 0&)
End Function

Public Function ValidateZbirnaPreUnosa(ByVal brojZbirne As String, _
                                      ByVal inputKgKlI As Double, _
                                      ByVal inputKgKlII As Double, _
                                      ByVal inputAmb As Long) As Variant
    On Error GoTo EH

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    Dim sumaKgKlI As Double
    Dim sumaKgKlII As Double
    Dim sumaAmb As Long

    If IsArray(otpData) Then
        Dim colKol As Long
        Dim colAmb As Long
        Dim colKlasa As Long
        Dim i As Long

        colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                    "modDokumenta.ValidateZbirnaPreUnosa")
        colAmb = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB, _
                                    "modDokumenta.ValidateZbirnaPreUnosa")
        colKlasa = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, _
                                      "modDokumenta.ValidateZbirnaPreUnosa")

        Dim rowKlasa As String

        For i = 1 To UBound(otpData, 1)
            rowKlasa = Trim$(CStr(otpData(i, colKlasa)))

            If rowKlasa = KLASA_I Then
                If IsNumeric(otpData(i, colKol)) Then sumaKgKlI = sumaKgKlI + CDbl(otpData(i, colKol))
            ElseIf rowKlasa = KLASA_II Then
                If IsNumeric(otpData(i, colKol)) Then sumaKgKlII = sumaKgKlII + CDbl(otpData(i, colKol))
            End If

            If IsNumeric(otpData(i, colAmb)) Then sumaAmb = sumaAmb + CLng(otpData(i, colAmb))
        Next i
    End If

    ValidateZbirnaPreUnosa = Array( _
        sumaKgKlI, inputKgKlI, sumaKgKlI - inputKgKlI, (Abs(sumaKgKlI - inputKgKlI) < 0.01), _
        sumaKgKlII, inputKgKlII, sumaKgKlII - inputKgKlII, (Abs(sumaKgKlII - inputKgKlII) < 0.01), _
        sumaAmb, inputAmb, sumaAmb - inputAmb _
    )

    Exit Function

EH:
    LogErr "modDokumenta.ValidateZbirnaPreUnosa"
    ValidateZbirnaPreUnosa = Array(0#, inputKgKlI, -inputKgKlI, False, _
                                   0#, inputKgKlII, -inputKgKlII, False, _
                                   0&, inputAmb, -inputAmb)
End Function

' ============================================================
' PRIJEMNICA – Kunde wiegt bei Annahme
' ============================================================

Public Function SavePrijemnicaMulti_TX(ByVal datum As Date, _
                                       ByVal kupacID As String, _
                                       ByVal vozacID As String, _
                                       ByVal brojPrij As String, _
                                       ByVal brojZbirne As String, _
                                       ByVal vrstaVoca As String, _
                                       ByVal sortaVoca As String, _
                                       ByVal kolicinaI As Double, _
                                       ByVal cenaI As Double, _
                                       ByVal tipAmb As String, _
                                       ByVal kolAmb As Long, _
                                       ByVal kolAmbVracena As Long, _
                                       Optional ByVal hasKlasaII As Boolean = False, _
                                       Optional ByVal kolicinaII As Double = 0, _
                                       Optional ByVal cenaII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE

    Dim resultI As String
    resultI = SavePrijemnica( _
        datum, _
        kupacID, _
        vozacID, _
        brojPrij, _
        brojZbirne, _
        vrstaVoca, _
        sortaVoca, _
        kolicinaI, _
        cenaI, _
        tipAmb, _
        kolAmb, _
        kolAmbVracena, _
        KLASA_I)

    If resultI = "" Then
        Err.Raise vbObjectError + 1301, "SavePrijemnicaMulti_TX", _
                  "SavePrijemnica Klasa I fehlgeschlagen"
    End If

    Dim resultII As String

    If hasKlasaII Then
        resultII = SavePrijemnica( _
            datum, _
            kupacID, _
            vozacID, _
            brojPrij, _
            brojZbirne, _
            vrstaVoca, _
            sortaVoca, _
            kolicinaII, _
            cenaII, _
            tipAmb, _
            0, _
            0, _
            KLASA_II)

        If resultII = "" Then
            Err.Raise vbObjectError + 1302, "SavePrijemnicaMulti_TX", _
                      "SavePrijemnica Klasa II fehlgeschlagen"
        End If

        SavePrijemnicaMulti_TX = resultI & " + " & resultII
    Else
        SavePrijemnicaMulti_TX = resultI
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    LogErr "SavePrijemnicaMulti_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SavePrijemnicaMulti_TX = ""
End Function

Public Function SavePrijemnica_TX(ByVal datum As Date, ByVal kupacID As String, _
                                   ByVal vozacID As String, ByVal brojPrij As String, _
                                   ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, ByVal kolicina As Double, _
                                   ByVal cena As Double, ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                                   Optional ByVal klasa As String = "I") As String
    Dim tx As New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE

    SavePrijemnica_TX = SavePrijemnica(datum, kupacID, vozacID, brojPrij, _
                                        brojZbirne, vrstaVoca, sortaVoca, _
                                        kolicina, cena, tipAmb, kolAmb, _
                                        kolAmbVracena, klasa)

    If SavePrijemnica_TX = "" Then
        Err.Raise vbObjectError + 1011, "SavePrijemnica_TX", _
                  "SavePrijemnica fehlgeschlagen"
    End If

    tx.CommitTx
    Exit Function

EH:
    LogErr "SavePrijemnica_TX"

    On Error Resume Next
    tx.RollbackTx
    On Error GoTo 0

    SavePrijemnica_TX = ""
End Function
    
Public Function SavePrijemnica(ByVal datum As Date, ByVal kupacID As String, _
                               ByVal vozacID As String, ByVal brojPrij As String, _
                               ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                               ByVal sortaVoca As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                               Optional ByVal klasa As String = "I") As String
    On Error GoTo EH

    If kupacID = "" Or brojZbirne = "" Or kolicina <= 0 Then
        Err.Raise vbObjectError + 1012, "SavePrijemnica", _
                  "Kupac, broj zbirne i kolicina su obavezni."
    End If

    Dim newID As String
    newID = GetNextID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-")

    If newID = "" Then
        Err.Raise vbObjectError + 1013, "SavePrijemnica", _
                  "GetNextID nije vratio PrijemnicaID."
    End If

    Dim rowData As Variant
    rowData = Array(newID, datum, kupacID, vozacID, brojPrij, brojZbirne, _
                    vrstaVoca, sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                    kolAmbVracena, klasa, "Ne", "")

    If AppendRow(TBL_PRIJEMNICA, rowData) <= 0 Then
        Err.Raise vbObjectError + 1014, "SavePrijemnica", _
                  "AppendRow fehlgeschlagen für tblPrijemnica."
    End If

    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    If kolAmbVracena > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmbVracena, "Ulaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    RelinkFakturaStavke newID, brojPrij

    SavePrijemnica = newID
    Exit Function

EH:
    LogErr "SavePrijemnica"
    SavePrijemnica = ""
End Function

Public Function GetPrijemniceByKupac(ByVal kupacID As String, _
                                      Optional ByVal datumOd As Date = 0, _
                                      Optional ByVal datumDo As Date = 0, _
                                      Optional ByVal samoNefakturisano As Boolean = False) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetPrijemniceByKupac = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, _
            "modDokumenta.GetPrijemniceByKupac"), "=", kupacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, _
                "modDokumenta.GetPrijemniceByKupac"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    If samoNefakturisano Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, _
                "modDokumenta.GetPrijemniceByKupac"), "<>", "Da"
        filters.Add fp
    End If

    GetPrijemniceByKupac = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetPrijemniceByKupac"
    GetPrijemniceByKupac = Empty
End Function

Public Function SaveKupciIzlaz_TX(ByVal datum As Date, _
                                  ByVal brojDok As String, _
                                  ByVal kupacNaziv As String, _
                                  ByVal kupacID As String, _
                                  ByVal vozacID As String, _
                                  ByVal tipAmb As String, _
                                  ByVal kolAmb As Long, _
                                  ByVal vrstaVoca As String, _
                                  ByVal novac As Double, _
                                  ByVal fakturaID As String, _
                                  ByVal napomena As String, _
                                  ByVal tipNovca As String) As Boolean
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If kupacID = "" Then
        Err.Raise vbObjectError + 1601, "SaveKupciIzlaz_TX", _
                  "KupacID je obavezan."
    End If

    If kolAmb <= 0 And novac <= 0 Then
        Err.Raise vbObjectError + 1602, "SaveKupciIzlaz_TX", _
                  "Nema ambalaže ni novca za cuvanje."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC
    tx.AddTableSnapshot TBL_FAKTURE

    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, _
                      "Izlaz", kupacID, "Kupac", _
                      vozacID, brojDok, DOK_TIP_IZLAZ_KUPCI
    End If

    If novac > 0 Then
        Dim novacID As String

        novacID = SaveNovac( _
            brojDok:=brojDok, _
            datum:=datum, _
            partner:=kupacNaziv, _
            partnerID:=kupacID, _
            entitetTip:="Kupac", _
            omID:="", _
            kooperantID:="", _
            fakturaID:=fakturaID, _
            vrstaVoca:=vrstaVoca, _
            tip:=tipNovca, _
            uplata:=novac, _
            isplata:=0, _
            napomena:=napomena, _
            otkupID:="")

        If novacID = "" Then
            Err.Raise vbObjectError + 1603, "SaveKupciIzlaz_TX", _
                      "SaveNovac fehlgeschlagen"
        End If

        If fakturaID <> "" Then
            UpdateFakturaStatus fakturaID
        End If
    End If

    tx.CommitTx
    Set tx = Nothing

    SaveKupciIzlaz_TX = True
    Exit Function

EH:
    LogErr "SaveKupciIzlaz_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveKupciIzlaz_TX = False
End Function

' ============================================================
' MANJAK – Schwundberechnung
' ============================================================

Public Function CalculateManjak(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim zbirnaKg As Double
    Dim zbrData As Variant

    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then
        zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

        Dim colBroj As Long
        Dim colZbrKol As Long
        Dim j As Long

        colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                     "modDokumenta.CalculateManjak")
        colZbrKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                       "modDokumenta.CalculateManjak")

        For j = 1 To UBound(zbrData, 1)
            If CStr(zbrData(j, colBroj)) = brojZbirne Then
                If IsNumeric(zbrData(j, colZbrKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(j, colZbrKol))
            End If
        Next j
    End If

    Dim prijKg As Double
    Dim prijData As Variant

    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)

    If Not IsEmpty(prijData) Then
        Dim colBrZbr As Long
        Dim colKol As Long
        Dim i As Long

        colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                      "modDokumenta.CalculateManjak")
        colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                    "modDokumenta.CalculateManjak")

        For i = 1 To UBound(prijData, 1)
            If CStr(prijData(i, colBrZbr)) = brojZbirne Then
                If IsNumeric(prijData(i, colKol)) Then prijKg = prijKg + CDbl(prijData(i, colKol))
            End If
        Next i
    End If

    Dim manjakKg As Double
    Dim manjakPct As Double

    manjakKg = zbirnaKg - prijKg

    If zbirnaKg > 0 Then manjakPct = manjakKg / zbirnaKg * 100

    CalculateManjak = Array(zbirnaKg, prijKg, manjakKg, manjakPct)
    Exit Function

EH:
    LogErr "modDokumenta.CalculateManjak"
    CalculateManjak = Array(0#, 0#, 0#, 0#)
End Function

Public Function CalculateManjakPreview(ByVal brojZbirne As String, _
                                      ByVal pendingKgKlI As Double, _
                                      ByVal pendingKgKlII As Double) As Variant
    On Error GoTo EH

    Dim zbirnaKg As Double

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)

    If IsArray(zbrData) Then
        Dim colBroj As Long
        Dim colKol As Long
        Dim i As Long

        colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                     "modDokumenta.CalculateManjakPreview")
        colKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                                    "modDokumenta.CalculateManjakPreview")

        For i = 1 To UBound(zbrData, 1)
            If CStr(zbrData(i, colBroj)) = brojZbirne Then
                If IsNumeric(zbrData(i, colKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(i, colKol))
            End If
        Next i
    End If

    Dim prijKg As Double
    Dim prijData As Variant

    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)

    If IsArray(prijData) Then
        Dim colBrZbr As Long
        Dim colPrijKol As Long

        colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                      "modDokumenta.CalculateManjakPreview")
        colPrijKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                        "modDokumenta.CalculateManjakPreview")

        For i = 1 To UBound(prijData, 1)
            If CStr(prijData(i, colBrZbr)) = brojZbirne Then
                If IsNumeric(prijData(i, colPrijKol)) Then prijKg = prijKg + CDbl(prijData(i, colPrijKol))
            End If
        Next i
    End If

    prijKg = prijKg + pendingKgKlI + pendingKgKlII

    Dim manjakKg As Double
    Dim manjakPct As Double

    manjakKg = zbirnaKg - prijKg

    If zbirnaKg > 0 Then manjakPct = manjakKg / zbirnaKg * 100

    CalculateManjakPreview = Array(zbirnaKg, prijKg, manjakKg, manjakPct)
    Exit Function

EH:
    LogErr "modDokumenta.CalculateManjakPreview"
    CalculateManjakPreview = Array(0#, pendingKgKlI + pendingKgKlII, 0#, 0#)
End Function

Public Function CalculateManjakByOtpremnica(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim manjak As Variant
    manjak = CalculateManjak(brojZbirne)

    Dim zbirnaKg As Double
    Dim manjakKg As Double
    Dim manjakPct As Double

    zbirnaKg = CDbl(manjak(0))
    manjakKg = CDbl(manjak(2))
    manjakPct = CDbl(manjak(3))

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)

    If IsEmpty(otpData) Then
        CalculateManjakByOtpremnica = Empty
        Exit Function
    End If

    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    If IsEmpty(otpData) Then
        CalculateManjakByOtpremnica = Empty
        Exit Function
    End If

    Dim colBroj As Long
    Dim colKol As Long
    Dim colCena As Long
    Dim colStan As Long

    colBroj = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, _
                                 "modDokumenta.CalculateManjakByOtpremnica")
    colStan = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA, _
                                 "modDokumenta.CalculateManjakByOtpremnica")
    colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                "modDokumenta.CalculateManjakByOtpremnica")
    colCena = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA, _
                                 "modDokumenta.CalculateManjakByOtpremnica")

    Dim rowCount As Long
    rowCount = UBound(otpData, 1)

    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 7)

    Dim i As Long
    For i = 1 To rowCount
        Dim kol As Double
        Dim udeo As Double
        Dim cena As Double

        kol = 0
        udeo = 0
        cena = 0

        If IsNumeric(otpData(i, colKol)) Then kol = CDbl(otpData(i, colKol))
        If zbirnaKg > 0 Then udeo = kol / zbirnaKg
        If IsNumeric(otpData(i, colCena)) Then cena = CDbl(otpData(i, colCena))

        result(i, 1) = CStr(otpData(i, colBroj))
        result(i, 2) = CStr(otpData(i, colStan))
        result(i, 3) = kol
        result(i, 4) = udeo
        result(i, 5) = udeo * manjakKg
        result(i, 6) = manjakPct
        result(i, 7) = udeo * manjakKg * cena
    Next i

    CalculateManjakByOtpremnica = result
    Exit Function

EH:
    LogErr "modDokumenta.CalculateManjakByOtpremnica"
    CalculateManjakByOtpremnica = Empty
End Function

' ============================================================
' PROSEK GAJBE – Durchschnittsgewicht pro Kästchen
' ============================================================

Public Function CalculateProsekGajbe(ByVal brojOtp As String) As Double
    On Error GoTo EH

    If Trim$(brojOtp) = "" Then Exit Function

    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_BROJ, _
                       "modDokumenta.CalculateProsekGajbe"
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                       "modDokumenta.CalculateProsekGajbe"
    RequireColumnIndex TBL_OTPREMNICA, COL_OTP_KOL_AMB, _
                       "modDokumenta.CalculateProsekGajbe"

    Dim kolVal As Variant
    Dim ambVal As Variant

    kolVal = LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOLICINA)
    ambVal = LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOL_AMB)

    Dim kol As Double
    Dim amb As Long

    If IsNumeric(kolVal) Then kol = CDbl(kolVal)
    If IsNumeric(ambVal) Then amb = CLng(ambVal)

    If amb > 0 Then
        CalculateProsekGajbe = kol / amb
    Else
        CalculateProsekGajbe = 0
    End If

    Exit Function

EH:
    LogErr "modDokumenta.CalculateProsekGajbe"
    CalculateProsekGajbe = 0
End Function

Public Function CalculateProsekGajbeByZbirna(ByVal brojZbirne As String) As Double
    On Error GoTo EH

    If Trim$(brojZbirne) = "" Then Exit Function

    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_BROJ, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KOLICINA, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"
    RequireColumnIndex TBL_ZBIRNA, COL_ZBR_KOL_AMB, _
                       "modDokumenta.CalculateProsekGajbeByZbirna"

    Dim zbrKol As Variant
    Dim zbrAmb As Variant

    zbrKol = LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOLICINA)
    zbrAmb = LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOL_AMB)

    Dim kol As Double
    Dim amb As Long

    If IsNumeric(zbrKol) Then kol = CDbl(zbrKol)
    If IsNumeric(zbrAmb) Then amb = CLng(zbrAmb)

    If amb > 0 Then
        CalculateProsekGajbeByZbirna = kol / amb
    Else
        CalculateProsekGajbeByZbirna = 0
    End If

    Exit Function

EH:
    LogErr "modDokumenta.CalculateProsekGajbeByZbirna"
    CalculateProsekGajbeByZbirna = 0
End Function

' ============================================================
' Storno Verweiste
' ============================================================

Public Function GetVerwaisteDokumente(ByVal dokumentTip As String) As Variant
    ' Returns: 2D Array der Dokumente deren BrojZbirne auf eine
    '          stornierte Zbirna zeigt, die selbst aber NICHT storniert sind.
    '
    ' dokumentTip: "Otpremnica" oder "Prijemnica"
    '
    ' Otpremnica Returns: (OtpremnicaID, BrojOtp, BrojZbirne, VrstaVoca, Kolicina)
    ' Prijemnica Returns: (PrijemnicaID, BrojPrij, BrojZbirne, KupacNaziv, Kolicina)
    ' oder Empty
    On Error GoTo EH

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)

    If IsEmpty(zbrData) Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If

    Dim colZbrBroj As Long
    Dim colZbrStorno As Long

    colZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, _
                                    "modDokumenta.GetVerwaisteDokumente")
    colZbrStorno = RequireColumnIndex(TBL_ZBIRNA, COL_STORNIRANO, _
                                      "modDokumenta.GetVerwaisteDokumente")

    Dim storniraneBrojevi As Object
    Set storniraneBrojevi = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim brz As String

    For i = 1 To UBound(zbrData, 1)
        If Trim$(NzToText(zbrData(i, colZbrStorno))) = "Da" Then
            brz = Trim$(NzToText(zbrData(i, colZbrBroj)))

            If brz <> "" Then
                If Not storniraneBrojevi.Exists(brz) Then
                    storniraneBrojevi.Add brz, True
                End If
            End If
        End If
    Next i

    For i = 1 To UBound(zbrData, 1)
        If Trim$(NzToText(zbrData(i, colZbrStorno))) <> "Da" Then
            brz = Trim$(NzToText(zbrData(i, colZbrBroj)))

            If storniraneBrojevi.Exists(brz) Then
                storniraneBrojevi.Remove brz
            End If
        End If
    Next i

    If storniraneBrojevi.count = 0 Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If

    Select Case dokumentTip
        Case "Otpremnica"
            GetVerwaisteDokumente = GetVerwaisteOtpremnice(storniraneBrojevi)

        Case "Prijemnica"
            GetVerwaisteDokumente = GetVerwaistePrijemnice(storniraneBrojevi)

        Case Else
            GetVerwaisteDokumente = Empty
    End Select

    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaisteDokumente"
    GetVerwaisteDokumente = Empty
End Function

Private Function GetVerwaisteOtpremnice(ByVal storniraneBrojevi As Object) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVerwaisteOtpremnice = Empty
        Exit Function
    End If

    Dim colID As Long
    Dim colBrOtp As Long
    Dim colBrZbr As Long
    Dim colVrsta As Long
    Dim colKol As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, _
                               "modDokumenta.GetVerwaisteOtpremnice")
    colBrOtp = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colBrZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, _
                                  "modDokumenta.GetVerwaisteOtpremnice")
    colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, _
                                "modDokumenta.GetVerwaisteOtpremnice")
    colStorno = RequireColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO, _
                                   "modDokumenta.GetVerwaisteOtpremnice")

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextCount

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            count = count + 1
        End If

NextCount:
    Next i

    If count = 0 Then
        GetVerwaisteOtpremnice = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)

    Dim idx As Long
    Dim kol As Double

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextRow

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            idx = idx + 1

            kol = 0
            If IsNumeric(data(i, colKol)) Then kol = CDbl(data(i, colKol))

            result(idx, 1) = NzToText(data(i, colID))
            result(idx, 2) = NzToText(data(i, colBrOtp))
            result(idx, 3) = NzToText(data(i, colBrZbr))
            result(idx, 4) = NzToText(data(i, colVrsta))
            result(idx, 5) = kol
        End If

NextRow:
    Next i

    GetVerwaisteOtpremnice = result
    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaisteOtpremnice"
    GetVerwaisteOtpremnice = Empty
End Function

Private Function GetVerwaistePrijemnice(ByVal storniraneBrojevi As Object) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If

    Dim colID As Long
    Dim colBrPrij As Long
    Dim colBrZbr As Long
    Dim colKupac As Long
    Dim colKol As Long
    Dim colStorno As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, _
                               "modDokumenta.GetVerwaistePrijemnice")
    colBrPrij = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, _
                                   "modDokumenta.GetVerwaistePrijemnice")
    colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, _
                                  "modDokumenta.GetVerwaistePrijemnice")
    colKupac = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, _
                                  "modDokumenta.GetVerwaistePrijemnice")
    colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, _
                                "modDokumenta.GetVerwaistePrijemnice")
    colStorno = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, _
                                   "modDokumenta.GetVerwaistePrijemnice")

    Dim count As Long
    Dim i As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextCount

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            count = count + 1
        End If

NextCount:
    Next i

    If count = 0 Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)

    Dim idx As Long
    Dim kupacNaziv As String
    Dim kol As Double

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colStorno))) = "Da" Then GoTo NextRow

        If storniraneBrojevi.Exists(Trim$(NzToText(data(i, colBrZbr)))) Then
            idx = idx + 1

            kupacNaziv = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, _
                                          NzToText(data(i, colKupac)), COL_KUP_NAZIV))

            kol = 0
            If IsNumeric(data(i, colKol)) Then kol = CDbl(data(i, colKol))

            result(idx, 1) = NzToText(data(i, colID))
            result(idx, 2) = NzToText(data(i, colBrPrij))
            result(idx, 3) = NzToText(data(i, colBrZbr))
            result(idx, 4) = kupacNaziv
            result(idx, 5) = kol
        End If

NextRow:
    Next i

    GetVerwaistePrijemnice = result
    Exit Function

EH:
    LogErr "modDokumenta.GetVerwaistePrijemnice"
    GetVerwaistePrijemnice = Empty
End Function

Public Sub RelinkFakturaStavke(ByVal newPrijemnicaID As String, _
                               ByVal brojPrijemnice As String)
    ' Sucht verwaiste FakturaStavke die auf eine stornierte Prijemnica
    ' mit gleichem BrojPrijemnice zeigen, und verlinkt sie auf die neue.
    On Error GoTo EH

    If Trim$(newPrijemnicaID) = "" Or Trim$(brojPrijemnice) = "" Then Exit Sub

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then Exit Sub

    Dim colPrijID As Long
    Dim colOsir As Long
    Dim colFakID As Long

    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, _
                                   "modDokumenta.RelinkFakturaStavke")
    colOsir = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD, _
                                 "modDokumenta.RelinkFakturaStavke")
    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, _
                                  "modDokumenta.RelinkFakturaStavke")

    Dim i As Long
    Dim oldPrijID As String
    Dim oldBroj As String
    Dim fakID As String
    Dim newPrijRows As Collection

    For i = 1 To UBound(stavkeData, 1)

        If Trim$(NzToText(stavkeData(i, colOsir))) = "" Then GoTo NextStavka

        oldPrijID = Trim$(NzToText(stavkeData(i, colPrijID)))

        If oldPrijID = "" Then GoTo NextStavka

        oldBroj = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, oldPrijID, COL_PRJ_BROJ))

        If oldBroj = brojPrijemnice Then
            fakID = Trim$(NzToText(stavkeData(i, colFakID)))

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_FS_PRIJEMNICA_ID, _
                              newPrijemnicaID, "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_OSIROCENO_OD, _
                              "", "modDokumenta.RelinkFakturaStavke"

            Set newPrijRows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, newPrijemnicaID)

            If newPrijRows.count = 0 Then
                Err.Raise vbObjectError + 7410, "modDokumenta.RelinkFakturaStavke", _
                          "Nova prijemnica nije pronadena za relink: " & newPrijemnicaID
            End If

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRows(1), COL_PRJ_FAKTURISANO, _
                              "Da", "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRows(1), COL_PRJ_FAKTURA_ID, _
                              fakID, "modDokumenta.RelinkFakturaStavke"
        End If

NextStavka:
    Next i

    Exit Sub

EH:
    LogErr "modDokumenta.RelinkFakturaStavke"
    Err.Raise Err.Number, "modDokumenta.RelinkFakturaStavke", Err.Description
End Sub

' ============================================================
' HELPER – Vozac-Report (ersetzt alten modTransport)
' ============================================================

Public Function GetVozacDokumenta(ByVal vozacID As String, _
                                   Optional ByVal datumOd As Date = 0, _
                                   Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetVozacDokumenta = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, _
            "modDokumenta.GetVozacDokumenta"), "=", vozacID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, _
                "modDokumenta.GetVozacDokumenta"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetVozacDokumenta = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modDokumenta.GetVozacDokumenta"
    GetVozacDokumenta = Empty
End Function

Public Function BuildZbirnaVrstaCache() As Object
    On Error GoTo EH

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(otpData) Then
        Set BuildZbirnaVrstaCache = dict
        Exit Function
    End If

    Dim colBrZbr As Long
    Dim colVrsta As Long

    colBrZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, _
                                  "modDokumenta.BuildZbirnaVrstaCache")
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, _
                                  "modDokumenta.BuildZbirnaVrstaCache")

    Dim i As Long
    Dim brz As String
    Dim vrsta As String

    For i = 1 To UBound(otpData, 1)
        brz = Trim$(NzToText(otpData(i, colBrZbr)))
        vrsta = Trim$(NzToText(otpData(i, colVrsta)))

        If brz <> "" Then
            If Not dict.Exists(brz) Then
                dict.Add brz, vrsta
            End If
        End If
    Next i

    Set BuildZbirnaVrstaCache = dict
    Exit Function

EH:
    LogErr "modDokumenta.BuildZbirnaVrstaCache"

    Dim emptyDict As Object
    Set emptyDict = CreateObject("Scripting.Dictionary")
    Set BuildZbirnaVrstaCache = emptyDict
End Function

Public Function GetVrstaFromCache(ByVal dict As Object, _
                                  ByVal brojZbirne As String) As String
    On Error GoTo EH

    If dict Is Nothing Then
        GetVrstaFromCache = ""
    ElseIf dict.Exists(brojZbirne) Then
        GetVrstaFromCache = CStr(dict(brojZbirne))
    Else
        GetVrstaFromCache = ""
    End If

    Exit Function

EH:
    LogErr "modDokumenta.GetVrstaFromCache"
    GetVrstaFromCache = ""
End Function

