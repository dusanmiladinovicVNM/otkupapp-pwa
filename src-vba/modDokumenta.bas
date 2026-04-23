Attribute VB_Name = "modDokumenta"

Option Explicit

' ============================================================
' modDokumenta – Otpremnica, Zbirna, Prijemnica
' Dokumentenfluss: Otkup zu Otpremnica zu Zbirna zu Prijemnica zu Faktura
' ============================================================

' ============================================================
' OTPREMNICA – Station gibt Ware an Fahrer
' ============================================================

Public Function SaveOtpremnica_TX(ByVal datum As Date, ByVal stanicaID As String, _
                                   ByVal vozacID As String, ByVal brojOtp As String, _
                                   ByVal brojZbirne As String, ByVal vrsta As String, _
                                   ByVal sorta As String, ByVal kolicina As Double, _
                                   ByVal cena As Double, ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, _
                                   Optional ByVal klasa As String = "I") As String
    Dim tx As New clsTransaction
    
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
    Exit Function
EH:
    LogErr "SaveOtpremnica_TX"
    tx.RollbackTx
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
    ' Alle Otpremnice einer Zbirna
    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then
        GetOtpremniceByZbirna = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE), "=", brojZbirne
    filters.Add fp
    
    GetOtpremniceByZbirna = FilterArray(data, filters)
End Function

Public Function GetOtpremniceByStation(ByVal stanicaID As String, _
                                       Optional ByVal datumOd As Date = 0, _
                                       Optional ByVal datumDo As Date = 0) As Variant
    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then
        GetOtpremniceByStation = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA), "=", stanicaID
    filters.Add fp
    
    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If
    
    GetOtpremniceByStation = FilterArray(data, filters)
End Function

' ============================================================
' ZBIRNA – Gesamtdokument Fahrer
' ============================================================

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
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "SaveZbirna_TX"
    tx.RollbackTx
    MsgBox "Greska pri unosu zbirne, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    SaveZbirna_TX = ""
End Function

Public Function SaveZbirna(ByVal datum As Date, ByVal vozacID As String, _
                           ByVal brojZbirne As String, ByVal kupacID As String, _
                           ByVal hladnjaca As String, ByVal pogon As String, _
                           ByVal vrstaVoca As String, ByVal sortaVoca As String, _
                           ByVal ukupnoKol As Double, ByVal tipAmb As String, _
                           ByVal ukupnoAmb As Long, _
                           Optional ByVal klasa As String = "I") As String
    
    If vozacID = "" Or brojZbirne = "" Then
        MsgBox "Vozac i broj zbirne su obavezni!", vbExclamation, APP_NAME
        SaveZbirna = ""
        Exit Function
    End If
    
    Dim newID As String
    newID = GetNextID(TBL_ZBIRNA, COL_ZBR_ID, "ZBR-")
    
    Dim rowData As Variant
    rowData = Array(newID, datum, vozacID, brojZbirne, kupacID, _
                    hladnjaca, pogon, vrstaVoca, sortaVoca, _
                    ukupnoKol, tipAmb, ukupnoAmb, klasa)
    
    If AppendRow(TBL_ZBIRNA, rowData) > 0 Then
        SaveZbirna = newID
    Else
        SaveZbirna = ""
    End If
End Function

Public Function GetZbirnaByKupac(ByVal kupacID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then
        GetZbirnaByKupac = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC), "=", kupacID
    filters.Add fp
    
    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If
    
    GetZbirnaByKupac = FilterArray(data, filters)
End Function

' ============================================================
' ZBIRNA VALIDIERUNG
' ============================================================

Public Function ValidateZbirna(ByVal brojZbirne As String) As Variant
    ' Prüft Summe Otpremnice vs Zbirna
    ' Returns: Array(SumaOtpKg, ZbirnaKg, RazlikaKg, ValidKg,
    '                SumaOtpAmb, ZbirnaAmb, RazlikaAmb)
    
    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)
    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    
    Dim sumaOtpKg As Double
    Dim sumaOtpAmb As Long
    
    If Not IsEmpty(otpData) Then
        Dim colKol As Long, colAmb As Long
        colKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
        colAmb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB)
        
        Dim i As Long
        For i = 1 To UBound(otpData, 1)
            If IsNumeric(otpData(i, colKol)) Then sumaOtpKg = sumaOtpKg + CDbl(otpData(i, colKol))
            If IsNumeric(otpData(i, colAmb)) Then sumaOtpAmb = sumaOtpAmb + CLng(otpData(i, colAmb))
        Next i
    End If
    
    ' Zbirna-Daten
    Dim zbirnaKg As Double
    Dim zbirnaAmb As Long
    
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
    
    If Not IsEmpty(zbrData) Then
        Dim colZbrBroj As Long, colZbrKol As Long, colZbrAmb As Long
        colZbrBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
        colZbrKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
        colZbrAmb = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB)
        
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
End Function

Public Function ValidateZbirnaPreUnosa(ByVal brojZbirne As String, _
                                      ByVal inputKgKlI As Double, _
                                      ByVal inputKgKlII As Double, _
                                      ByVal inputAmb As Long) As Variant
    ' Returns: Array(SumaOtpKgKlI, InputKgKlI, RazlikaKgKlI, ValidKgKlI,
    '                SumaOtpKgKlII, InputKgKlII, RazlikaKgKlII, ValidKgKlII,
    '                SumaOtpAmb, InputAmb, RazlikaAmb)

    Dim otpData As Variant
    otpData = GetOtpremniceByZbirna(brojZbirne)
    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

    Dim sumaKgKlI As Double
    Dim sumaKgKlII As Double
    Dim sumaAmb As Long

    If IsArray(otpData) Then
        Dim colKol As Long, colAmb As Long, colKlasa As Long, i As Long
        colKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
        colAmb = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOL_AMB)
        colKlasa = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA)
        
        If colKol > 0 And colAmb > 0 And colKlasa > 0 Then
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
    End If

    ValidateZbirnaPreUnosa = Array( _
        sumaKgKlI, inputKgKlI, sumaKgKlI - inputKgKlI, (Abs(sumaKgKlI - inputKgKlI) < 0.01), _
        sumaKgKlII, inputKgKlII, sumaKgKlII - inputKgKlII, (Abs(sumaKgKlII - inputKgKlII) < 0.01), _
        sumaAmb, inputAmb, sumaAmb - inputAmb _
    )
End Function

' ============================================================
' PRIJEMNICA – Kunde wiegt bei Annahme
' ============================================================

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
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE  ' ? NEU (RelinkFakturaStavke)
    tx.AddTableSnapshot TBL_FAKTURE          ' ? NEU (falls Relink)

    SavePrijemnica_TX = SavePrijemnica(datum, kupacID, vozacID, brojPrij, _
                                        brojZbirne, vrstaVoca, sortaVoca, _
                                        kolicina, cena, tipAmb, kolAmb, _
                                        kolAmbVracena, klasa)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "SavePrijemnica_TX"
    tx.RollbackTx
    MsgBox "Greska pri unosu prijemnice, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    SavePrijemnica_TX = ""
End Function
    
Public Function SavePrijemnica(ByVal datum As Date, ByVal kupacID As String, _
                               ByVal vozacID As String, ByVal brojPrij As String, _
                               ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                               ByVal sortaVoca As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                               Optional ByVal klasa As String = "I") As String
    
    If kupacID = "" Or brojZbirne = "" Or kolicina <= 0 Then
        MsgBox "Kupac, broj zbirne i kolicina su obavezni!", vbExclamation, APP_NAME
        SavePrijemnica = ""
        Exit Function
    End If
    
    Dim newID As String
    newID = GetNextID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-")
    
    Dim rowData As Variant
    rowData = Array(newID, datum, kupacID, vozacID, brojPrij, brojZbirne, _
                    vrstaVoca, sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                    kolAmbVracena, klasa, "Ne", "")
    
    If AppendRow(TBL_PRIJEMNICA, rowData) > 0 Then
        If kolAmb > 0 Then
            TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", kupacID, "Kupac", vozacID, newID, DOK_TIP_PRIJEMNICA
        End If
        If kolAmbVracena > 0 Then
            TrackAmbalaza datum, tipAmb, kolAmbVracena, "Ulaz", kupacID, "Kupac", vozacID, newID, DOK_TIP_PRIJEMNICA
        End If
        RelinkFakturaStavke newID, brojPrij
        SavePrijemnica = newID
    Else
        SavePrijemnica = ""
    End If
End Function
Public Function GetPrijemniceByKupac(ByVal kupacID As String, _
                                      Optional ByVal datumOd As Date = 0, _
                                      Optional ByVal datumDo As Date = 0, _
                                      Optional ByVal samoNefakturisano As Boolean = False) As Variant
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then
        GetPrijemniceByKupac = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC), "=", kupacID
    filters.Add fp
    
    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If
    
    ' TODO: Fakturisano-Status wenn in tblPrijemnica hinzugefügt
    
    GetPrijemniceByKupac = FilterArray(data, filters)
End Function

' ============================================================
' MANJAK – Schwundberechnung
' ============================================================

Public Function CalculateManjak(ByVal brojZbirne As String) As Variant
    ' Returns: Array(ZbirnaKg, PrijemnicaKg, ManjakKg, ManjakPct)
    
    ' Zbirna-Gewicht
    Dim zbirnaKg As Double
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsArray(zbrData) Then
        zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
        Dim colBroj As Long, colZbrKol As Long
        colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
        colZbrKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
        Dim j As Long
        For j = 1 To UBound(zbrData, 1)
            If CStr(zbrData(j, colBroj)) = brojZbirne Then
                If IsNumeric(zbrData(j, colZbrKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(j, colZbrKol))
            End If
        Next j
    End If
    
    ' Prijemnica-Gewicht (Summe aller Prijemnice mit diesem BrojZbirne)
    Dim prijKg As Double
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If Not IsEmpty(prijData) Then
        Dim colBrZbr As Long, colKol As Long
        colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
        
        Dim i As Long
        For i = 1 To UBound(prijData, 1)
            If CStr(prijData(i, colBrZbr)) = brojZbirne Then
                If IsNumeric(prijData(i, colKol)) Then prijKg = prijKg + CDbl(prijData(i, colKol))
            End If
        Next i
    End If
    
    Dim manjakKg As Double
    manjakKg = zbirnaKg - prijKg
    
    Dim manjakPct As Double
    If zbirnaKg > 0 Then manjakPct = manjakKg / zbirnaKg * 100
    
    CalculateManjak = Array(zbirnaKg, prijKg, manjakKg, manjakPct)
End Function

Public Function CalculateManjakPreview(ByVal brojZbirne As String, _
                                      ByVal pendingKgKlI As Double, _
                                      ByVal pendingKgKlII As Double) As Variant
    ' Returns: Array(ZbirnaKgGesamt, PrijemnicaKgGesamt, ManjakKg, ManjakPct)

    Dim zbirnaKg As Double
    
    ' Zbirna: Summe beider Klassen
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
    If IsArray(zbrData) Then
        Dim colBroj As Long, colKol As Long, i As Long
        colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
        colKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
        If colBroj > 0 And colKol > 0 Then
            For i = 1 To UBound(zbrData, 1)
                If CStr(zbrData(i, colBroj)) = brojZbirne Then
                    If IsNumeric(zbrData(i, colKol)) Then zbirnaKg = zbirnaKg + CDbl(zbrData(i, colKol))
                End If
            Next i
        End If
    End If

    ' Prijemnica: Summe aller bestehenden
    Dim prijKg As Double
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If IsArray(prijData) Then
        Dim colBrZbr As Long, colPrijKol As Long
        colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        colPrijKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
        If colBrZbr > 0 And colPrijKol > 0 Then
            For i = 1 To UBound(prijData, 1)
                If CStr(prijData(i, colBrZbr)) = brojZbirne Then
                    If IsNumeric(prijData(i, colPrijKol)) Then prijKg = prijKg + CDbl(prijData(i, colPrijKol))
                End If
            Next i
        End If
    End If

    ' Pending addieren (noch nicht gespeichert)
    prijKg = prijKg + pendingKgKlI + pendingKgKlII

    Dim manjakKg As Double, manjakPct As Double
    manjakKg = zbirnaKg - prijKg
    If zbirnaKg > 0 Then manjakPct = manjakKg / zbirnaKg * 100

    CalculateManjakPreview = Array(zbirnaKg, prijKg, manjakKg, manjakPct)
End Function

Public Function CalculateManjakByOtpremnica(ByVal brojZbirne As String) As Variant
    ' Manjak proportional aufgeteilt pro Otpremnica
    ' Returns: 2D Array (BrojOtp, StanicaID, Kolicina, Udeo, ManjakKg, ManjakPct, ManjakRSD)
    
    Dim manjak As Variant
    manjak = CalculateManjak(brojZbirne)
    Dim zbirnaKg As Double: zbirnaKg = CDbl(manjak(0))
    Dim manjakKg As Double: manjakKg = CDbl(manjak(2))
    Dim manjakPct As Double: manjakPct = CDbl(manjak(3))
    
    ' Alle Otpremnice dieser Zbirna
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
    
    Dim colBroj As Long, colKol As Long, colCena As Long, colStan As Long
    colBroj = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    colStan = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_STANICA)
    colKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    colCena = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA)
    
    Dim rowCount As Long
    rowCount = UBound(otpData, 1)
    
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 7)
    
    Dim i As Long
    For i = 1 To rowCount
        Dim kol As Double, udeo As Double, cena As Double
        kol = 0: udeo = 0: cena = 0
        
        If IsNumeric(otpData(i, colKol)) Then kol = CDbl(otpData(i, colKol))
        If zbirnaKg > 0 Then udeo = kol / zbirnaKg
        If IsNumeric(otpData(i, colCena)) Then cena = CDbl(otpData(i, colCena))
        
        result(i, 1) = CStr(otpData(i, colBroj))    ' BrojOtpremnice
        result(i, 2) = CStr(otpData(i, colStan))     ' StanicaID
        result(i, 3) = kol                            ' Kolicina
        result(i, 4) = udeo                           ' Udeo u zbirnoj
        result(i, 5) = udeo * manjakKg                ' ManjakKg
        result(i, 6) = manjakPct                      ' ManjakPct
        result(i, 7) = udeo * manjakKg * cena          ' ManjakRSD
    Next i
    
    CalculateManjakByOtpremnica = result
End Function

' ============================================================
' PROSEK GAJBE – Durchschnittsgewicht pro Kästchen
' ============================================================

Public Function CalculateProsekGajbe(ByVal brojOtp As String) As Double
    Dim kolVal As Variant, ambVal As Variant
    kolVal = LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOLICINA)
    ambVal = LookupValue(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOL_AMB)
    
    Dim kol As Double, amb As Long
    If Not IsEmpty(kolVal) Then
        If IsNumeric(kolVal) Then kol = CDbl(kolVal)
    End If
    If Not IsEmpty(ambVal) Then
        If IsNumeric(ambVal) Then amb = CLng(ambVal)
    End If
    
    If amb > 0 Then
        CalculateProsekGajbe = kol / amb
    Else
        CalculateProsekGajbe = 0
    End If
End Function

Public Function CalculateProsekGajbeByZbirna(ByVal brojZbirne As String) As Double
    ' Durchschnittsgewicht für ganze Zbirna
    Dim zbrKol As Variant, zbrAmb As Variant
    zbrKol = LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOLICINA)
    zbrAmb = LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOL_AMB)
    
    Dim kol As Double, amb As Long
    If Not IsEmpty(zbrKol) Then
        If IsNumeric(zbrKol) Then kol = CDbl(zbrKol)
    End If
    If Not IsEmpty(zbrAmb) Then
        If IsNumeric(zbrAmb) Then amb = CLng(zbrAmb)
    End If
    
    If amb > 0 Then
        CalculateProsekGajbeByZbirna = kol / amb
    Else
        CalculateProsekGajbeByZbirna = 0
    End If
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
    
    ' 1. Alle stornierten BrojZbirne sammeln
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If
    
    Dim storniraneBrojevi As Object
    Set storniraneBrojevi = CreateObject("Scripting.Dictionary")
    
    Dim colZbrBroj As Long, colZbrStorno As Long
    colZbrBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    colZbrStorno = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
    
    If colZbrStorno = 0 Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(zbrData, 1)
        If CStr(zbrData(i, colZbrStorno)) = "Da" Then
            Dim brz As String
            brz = CStr(zbrData(i, colZbrBroj))
            If Not storniraneBrojevi.Exists(brz) Then
                storniraneBrojevi.Add brz, True
            End If
        End If
    Next i
    
    ' Prüfe ob es für diese BrojZbirne eine NEUE (nicht-stornierte) Zbirna gibt
    For i = 1 To UBound(zbrData, 1)
        If CStr(zbrData(i, colZbrStorno)) <> "Da" Then
            brz = CStr(zbrData(i, colZbrBroj))
            If storniraneBrojevi.Exists(brz) Then
                ' Neue Zbirna existiert ? nicht mehr verwaist
                storniraneBrojevi.Remove brz
            End If
        End If
    Next i
    
    If storniraneBrojevi.count = 0 Then
        GetVerwaisteDokumente = Empty
        Exit Function
    End If
    
    ' 2. Dokumente finden die auf stornierte BrojZbirne zeigen
    If dokumentTip = "Otpremnica" Then
        GetVerwaisteDokumente = GetVerwaisteOtpremnice(storniraneBrojevi)
    ElseIf dokumentTip = "Prijemnica" Then
        GetVerwaisteDokumente = GetVerwaistePrijemnice(storniraneBrojevi)
    Else
        GetVerwaisteDokumente = Empty
    End If
End Function

Private Function GetVerwaisteOtpremnice(ByVal storniraneBrojevi As Object) As Variant
    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then
        GetVerwaisteOtpremnice = Empty
        Exit Function
    End If
    
    Dim colID As Long, colBrOtp As Long, colBrZbr As Long
    Dim colVrsta As Long, colKol As Long, colStorno As Long
    colID = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    colBrOtp = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    colBrZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    colVrsta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA)
    colKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    colStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)
    
    ' Zählen
    Dim count As Long
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If colStorno > 0 Then
            If CStr(data(i, colStorno)) = "Da" Then GoTo NextCount
        End If
        If storniraneBrojevi.Exists(CStr(data(i, colBrZbr))) Then
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
    
    For i = 1 To UBound(data, 1)
        If colStorno > 0 Then
            If CStr(data(i, colStorno)) = "Da" Then GoTo NextRow
        End If
        If storniraneBrojevi.Exists(CStr(data(i, colBrZbr))) Then
            idx = idx + 1
            result(idx, 1) = CStr(data(i, colID))
            result(idx, 2) = CStr(data(i, colBrOtp))
            result(idx, 3) = CStr(data(i, colBrZbr))
            result(idx, 4) = CStr(data(i, colVrsta))
            result(idx, 5) = CDbl(data(i, colKol))
        End If
NextRow:
    Next i
    
    GetVerwaisteOtpremnice = result
End Function

Private Function GetVerwaistePrijemnice(ByVal storniraneBrojevi As Object) As Variant
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If
    
    Dim colID As Long, colBrPrij As Long, colBrZbr As Long
    Dim colKupac As Long, colKol As Long, colStorno As Long
    colID = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    colBrPrij = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    colKupac = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC)
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    
    ' Zählen
    Dim count As Long
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If colStorno > 0 Then
            If CStr(data(i, colStorno)) = "Da" Then GoTo NextCountP
        End If
        If storniraneBrojevi.Exists(CStr(data(i, colBrZbr))) Then
            count = count + 1
        End If
NextCountP:
    Next i
    
    If count = 0 Then
        GetVerwaistePrijemnice = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To count, 1 To 5)
    Dim idx As Long
    
    For i = 1 To UBound(data, 1)
        If colStorno > 0 Then
            If CStr(data(i, colStorno)) = "Da" Then GoTo NextRowP
        End If
        If storniraneBrojevi.Exists(CStr(data(i, colBrZbr))) Then
            idx = idx + 1
            result(idx, 1) = CStr(data(i, colID))
            result(idx, 2) = CStr(data(i, colBrPrij))
            result(idx, 3) = CStr(data(i, colBrZbr))
            result(idx, 4) = CStr(LookupValue(TBL_KUPCI, "KupacID", _
                                  CStr(data(i, colKupac)), "Naziv"))
            result(idx, 5) = CDbl(data(i, colKol))
        End If
NextRowP:
    Next i
    
    GetVerwaistePrijemnice = result
End Function

Public Sub RelinkFakturaStavke(ByVal newPrijemnicaID As String, _
                                ByVal brojPrijemnice As String)
    ' Sucht verwaiste FakturaStavke die auf eine stornierte Prijemnica
    ' mit gleichem BrojPrijemnice zeigen, und verlinkt sie auf die neue.
    
    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If IsEmpty(stavkeData) Then Exit Sub
    
    Dim colPrijID As Long, colOsir As Long
    colPrijID = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID)
    colOsir = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD)
    If colOsir = 0 Then Exit Sub
    
    Dim i As Long
    For i = 1 To UBound(stavkeData, 1)
        ' Nur verwaiste Stavke
        If CStr(stavkeData(i, colOsir)) = "" Then GoTo NextStavka
        
        ' Alte PrijemnicaID holen
        Dim oldPrijID As String
        oldPrijID = CStr(stavkeData(i, colPrijID))
        
        ' BrojPrijemnice der alten (stornierten) Prijemnica prüfen
        Dim oldBroj As String
        oldBroj = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, oldPrijID, COL_PRJ_BROJ))
        
        If oldBroj = brojPrijemnice Then
            ' Re-link auf neue Prijemnica
            UpdateCell TBL_FAKTURA_STAVKE, i, colPrijID, newPrijemnicaID
            UpdateCell TBL_FAKTURA_STAVKE, i, COL_OSIROCENO_OD, ""
            
            ' Neue Prijemnica als fakturisano markieren
            Dim fakID As String
            fakID = CStr(stavkeData(i, GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID)))
            Dim newPrijRows As Collection
            Set newPrijRows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, newPrijemnicaID)
            If newPrijRows.count > 0 Then
                UpdateCell TBL_PRIJEMNICA, newPrijRows(1), COL_PRJ_FAKTURISANO, "Da"
                UpdateCell TBL_PRIJEMNICA, newPrijRows(1), COL_PRJ_FAKTURA_ID, fakID
            End If
        End If
NextStavka:
    Next i
End Sub

' ============================================================
' HELPER – Vozac-Report (ersetzt alten modTransport)
' ============================================================

Public Function GetVozacDokumenta(ByVal vozacID As String, _
                                   Optional ByVal datumOd As Date = 0, _
                                   Optional ByVal datumDo As Date = 0) As Variant
    ' Alle Otpremnice eines Fahrers
    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(data) Then
        GetVozacDokumenta = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC), "=", vozacID
    filters.Add fp
    
    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If
    
    GetVozacDokumenta = FilterArray(data, filters)
End Function

Public Function BuildZbirnaVrstaCache() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsEmpty(otpData) Then
        Set BuildZbirnaVrstaCache = dict
        Exit Function
    End If
    
    Dim colBrZbr As Long, colVrsta As Long
    colBrZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    colVrsta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA)
    
    Dim i As Long
    For i = 1 To UBound(otpData, 1)
        Dim brz As String
        brz = CStr(otpData(i, colBrZbr))
        If Not dict.Exists(brz) Then
            dict.Add brz, CStr(otpData(i, colVrsta))
        End If
    Next i
    
    Set BuildZbirnaVrstaCache = dict
End Function

Public Function GetVrstaFromCache(ByVal dict As Object, _
                                  ByVal brojZbirne As String) As String
    If dict Is Nothing Then
        GetVrstaFromCache = ""
    ElseIf dict.Exists(brojZbirne) Then
        GetVrstaFromCache = dict(brojZbirne)
    Else
        GetVrstaFromCache = ""
    End If
End Function
