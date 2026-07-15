Attribute VB_Name = "modDokumenta"
'Attribute VB_Name = "modDokumenta"

Option Explicit

' ============================================================
' modDokumenta - Otpremnica, Zbirna, Prijemnica
' Dokumentenfluss: Otkup zu Otpremnica zu Zbirna zu Prijemnica zu Faktura
' ============================================================

' ============================================================
' OTPREMNICA - Station gibt Ware an Fahrer
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
                                       Optional ByVal cenaII As Double = 0, _
                                       Optional ByVal brutoKgI As Double = 0, _
                                       Optional ByVal kolAmbII As Long = 0, _
                                       Optional ByVal brutoKgII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    ' Klasa I je opciona (kolicinaI = 0 -> snima se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)
    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1103, "SaveOtpremnicaMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    Dim resultI As String
    If hasKlasaI Then
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
            KLASA_I, _
            brutoKgI)

        If resultI = "" Then
            Err.Raise vbObjectError + 1101, "SaveOtpremnicaMulti_TX", _
                      "SaveOtpremnica Klasa I fehlgeschlagen"
        End If
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
            kolAmbII, _
            KLASA_II, _
            brutoKgII)

        If resultII = "" Then
            Err.Raise vbObjectError + 1102, "SaveOtpremnicaMulti_TX", _
                      "SaveOtpremnica Klasa II fehlgeschlagen"
        End If
    End If

    If hasKlasaI And hasKlasaII Then
        SaveOtpremnicaMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SaveOtpremnicaMulti_TX = resultI
    Else
        SaveOtpremnicaMulti_TX = resultII
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "SaveOtpremnicaMulti_TX"

    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnicaMulti_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnicaMulti_TX, _
        correlationId:=brojOtp, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveOtpremnicaMulti_TX failed. BrojOtp=" & brojOtp & _
                 "; BrojZbirne=" & brojZbirne & _
                 "; StanicaID=" & stanicaID & _
                 "; VozacID=" & vozacID & _
                 "; HasKlasaII=" & CStr(hasKlasaII) & _
                 "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnicaMulti_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnicaMulti_TX, _
        correlationId:=brojOtp

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
                                   Optional ByVal klasa As String = "I", _
                                   Optional ByVal brutoKg As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTPREMNICA
    tx.AddTableSnapshot TBL_AMBALAZA

    SaveOtpremnica_TX = SaveOtpremnica(datum, stanicaID, vozacID, brojOtp, _
                                        brojZbirne, vrsta, sorta, kolicina, _
                                        cena, tipAmb, kolAmb, klasa, brutoKg)

    If SaveOtpremnica_TX = "" Then
        Err.Raise vbObjectError + 1001, "SaveOtpremnica_TX", _
                  "SaveOtpremnica fehlgeschlagen"
    End If

    tx.CommitTx

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SaveOtpremnica_TX"
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnica_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnica_TX, _
        correlationId:=brojOtp, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveOtpremnica_TX failed. BrojOtp=" & brojOtp & _
                "; BrojZbirne=" & brojZbirne & _
                "; StanicaID=" & stanicaID & _
                "; VozacID=" & vozacID & _
                "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveOtpremnica_TX", _
        entityType:="Otpremnica", _
        entityID:=SaveOtpremnica_TX, _
        correlationId:=brojOtp

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtpremnica_TX = ""

    PrintTxFailure "SaveOtpremnica_TX", errSrc, errNum, errDesc
End Function

Public Function SaveOtpremnica(ByVal datum As Date, ByVal stanicaID As String, _
                               ByVal vozacID As String, ByVal brojOtp As String, _
                               ByVal brojZbirne As String, ByVal vrsta As String, _
                               ByVal sorta As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, _
                               Optional ByVal klasa As String = "I", _
                               Optional ByVal brutoKg As Double = 0) As String

    On Error GoTo EH

    Call ValidateOtpremnicaInput(stanicaID, vozacID, brojOtp, brojZbirne, _
                             kolicina, cena, tipAmb, kolAmb, klasa)

    Dim newID As String
    newID = GetNextID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-")

    If Len(Trim$(newID)) = 0 Then
        Err.Raise vbObjectError + 1002, "SaveOtpremnica", _
                "GetNextID nije vratio OtpremnicaID."
    End If

    Dim rowData As Variant
    rowData = Array(newID, datum, stanicaID, vozacID, brojOtp, _
                    brojZbirne, vrsta, sorta, kolicina, cena, tipAmb, kolAmb, klasa)

    Dim newRow As Long
    newRow = AppendRow(TBL_OTPREMNICA, rowData)
    If newRow > 0 Then
        If kolAmb > 0 Then
            TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", stanicaID, "Stanica", vozacID, newID, DOK_TIP_OTPREMNICA
        End If
        ' Bruto (kad je unet bruto pa oduzeta ambalaza) -> upis po imenu; prazno = neto.
        If brutoKg > 0 Then UpdateCell TBL_OTPREMNICA, newRow, COL_OTP_BRUTO, brutoKg
        SaveOtpremnica = newID
    Else
        Err.Raise vbObjectError + 1003, "SaveOtpremnica", _
                  "AppendRow fehlgeschlagen fuer tblOtpremnica"
    End If
    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SaveOtpremnica"
    On Error GoTo 0

    Err.Raise errNum, "SaveOtpremnica", _
              "Source=" & errSrc & " | " & errDesc
End Function

Public Function GetOtpremniceByZbirna(ByVal brojZbirne As String) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTPREMNICA)

    If IsEmpty(data) Then
        GetOtpremniceByZbirna = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_OTPREMNICA)

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

' True ako zbirna sa datim brojem postoji u tblZbirna (iskljucujuci stornirane).
' Storno-aware namerno: prijemnica ne sme da referencira storniranu zbirnu.
' (modBrojevi.BrojZbirneExists je Private i broji i stornirane -> drugi scenario.)
Public Function ZbirnaPostoji(ByVal brojZbirne As String) As Boolean
    On Error GoTo EH

    Dim b As String: b = Trim$(brojZbirne)
    If Len(b) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function
    data = ExcludeStornirano(data, TBL_ZBIRNA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long
    iBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modDokumenta.ZbirnaPostoji")

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(r, iBroj))), b, vbTextCompare) = 0 Then
            ZbirnaPostoji = True
            Exit Function
        End If
    Next r
    Exit Function

EH:
    LogErr "modDokumenta.ZbirnaPostoji", "broj=" & brojZbirne
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
    
    data = ExcludeStornirano(data, TBL_OTPREMNICA)

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
' ZBIRNA - Gesamtdokument Fahrer
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
                                   Optional ByVal ukupnoKolII As Double = 0, _
                                   Optional ByVal ukupnoAmbII As Long = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_ZBIRNA

    ' Klasa I je opciona (ukupnoKolI = 0 -> snima se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (ukupnoKolI > 0)
    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1203, "SaveZbirnaMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    Dim resultI As String
    If hasKlasaI Then
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
            ukupnoAmbII, _
            KLASA_II)

        If resultII = "" Then
            Err.Raise vbObjectError + 1202, "SaveZbirnaMulti_TX", _
                      "SaveZbirna Klasa II fehlgeschlagen"
        End If
    End If

    If hasKlasaI And hasKlasaII Then
        SaveZbirnaMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SaveZbirnaMulti_TX = resultI
    Else
        SaveZbirnaMulti_TX = resultII
    End If

    tx.CommitTx

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SaveZbirnaMulti_TX"
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirnaMulti_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirnaMulti_TX, _
        correlationId:=brojZbirne, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveZbirnaMulti_TX failed. BrojZbirne=" & brojZbirne & _
             "; VozacID=" & vozacID & _
             "; KupacID=" & kupacID & _
             "; HasKlasaII=" & CStr(hasKlasaII) & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirnaMulti_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirnaMulti_TX, _
        correlationId:=brojZbirne

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveZbirnaMulti_TX = ""

    PrintTxFailure "SaveZbirnaMulti_TX", errSrc, errNum, errDesc
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
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SaveZbirna_TX"
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirna_TX, _
        correlationId:=brojZbirne, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SaveZbirna_TX failed. BrojZbirne=" & brojZbirne & _
             "; VozacID=" & vozacID & _
             "; KupacID=" & kupacID & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SaveZbirna_TX", _
        entityType:="Zbirna", _
        entityID:=SaveZbirna_TX, _
        correlationId:=brojZbirne

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveZbirna_TX = ""

    PrintTxFailure "SaveZbirna_TX", errSrc, errNum, errDesc
End Function

Public Function SaveZbirna(ByVal datum As Date, ByVal vozacID As String, _
                           ByVal brojZbirne As String, ByVal kupacID As String, _
                           ByVal hladnjaca As String, ByVal pogon As String, _
                           ByVal vrstaVoca As String, ByVal sortaVoca As String, _
                           ByVal ukupnoKol As Double, ByVal tipAmb As String, _
                           ByVal ukupnoAmb As Long, _
                           Optional ByVal klasa As String = "I") As String
    On Error GoTo EH

    Call ValidateZbirnaInput(vozacID, brojZbirne, kupacID, ukupnoKol, _
                         tipAmb, ukupnoAmb, klasa)

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
                  "AppendRow fehlgeschlagen fuer tblZbirna."
    End If

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SaveZbirna"
    On Error GoTo 0

    Err.Raise errNum, "SaveZbirna", _
              "Source=" & errSrc & " | " & errDesc
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
    
    data = ExcludeStornirano(data, TBL_ZBIRNA)

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
    ' Prueft Summe Otpremnice vs Zbirna
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
' PRIJEMNICA - Kunde wiegt bei Annahme
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
                                       Optional ByVal cenaII As Double = 0, _
                                       Optional ByVal kolAmbII As Long = 0, _
                                       Optional ByVal brutoKgI As Double = 0, _
                                       Optional ByVal brutoKgII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    ' Klasa I je opciona (kolicinaI = 0 -> snima se samo Klasa II). Bar jedna klasa.
    Dim hasKlasaI As Boolean: hasKlasaI = (kolicinaI > 0)
    If Not hasKlasaI And Not hasKlasaII Then
        Err.Raise vbObjectError + 1303, "SavePrijemnicaMulti_TX", _
                  "Mora postojati bar jedna klasa (I ili II)."
    End If

    Dim resultI As String
    If hasKlasaI Then
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
            KLASA_I, _
            brutoKgI)

        If resultI = "" Then
            Err.Raise vbObjectError + 1301, "SavePrijemnicaMulti_TX", _
                      "SavePrijemnica Klasa I fehlgeschlagen"
        End If
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
            kolAmbII, _
            0, _
            KLASA_II, _
            brutoKgII)

        If resultII = "" Then
            Err.Raise vbObjectError + 1302, "SavePrijemnicaMulti_TX", _
                      "SavePrijemnica Klasa II fehlgeschlagen"
        End If
    End If

    If hasKlasaI And hasKlasaII Then
        SavePrijemnicaMulti_TX = resultI & " + " & resultII
    ElseIf hasKlasaI Then
        SavePrijemnicaMulti_TX = resultI
    Else
        SavePrijemnicaMulti_TX = resultII
    End If

    ' Paletizacija Klase I UNUTAR transakcije -> atomicno sa prijemnicom.
    Dim closedPal As Collection
    Set closedPal = New Collection
    If hasKlasaI And kolAmb > 0 Then
        PaletizePrijemnica prijemnicaID:=resultI, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=KLASA_I, netoKg:=kolicinaI, brGajbica:=kolAmb, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    ' Paletizacija Klase II (zasebne gajbe) -> u istu kolekciju zatvorenih paleta.
    If hasKlasaII And kolAmbII > 0 Then
        PaletizePrijemnica prijemnicaID:=resultII, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=KLASA_II, netoKg:=kolicinaII, brGajbica:=kolAmbII, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    tx.CommitTx

    ' Print/PDF zatvorenih paleta = post-commit side effect (bez rollback rizika).
    PaletniListOutputClosed closedPal

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SavePrijemnicaMulti_TX"
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnicaMulti_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnicaMulti_TX, _
        correlationId:=brojPrij, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SavePrijemnicaMulti_TX failed. BrojPrij=" & brojPrij & _
             "; BrojZbirne=" & brojZbirne & _
             "; KupacID=" & kupacID & _
             "; VozacID=" & vozacID & _
             "; HasKlasaII=" & CStr(hasKlasaII) & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnicaMulti_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnicaMulti_TX, _
        correlationId:=brojPrij

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SavePrijemnicaMulti_TX = ""

    PrintTxFailure "SavePrijemnicaMulti_TX", errSrc, errNum, errDesc
End Function

Public Function SavePrijemnica_TX(ByVal datum As Date, ByVal kupacID As String, _
                                   ByVal vozacID As String, ByVal brojPrij As String, _
                                   ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, ByVal kolicina As Double, _
                                   ByVal cena As Double, ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                                   Optional ByVal klasa As String = "I", _
                                   Optional ByVal brutoKg As Double = 0) As String
    Dim tx As New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    SavePrijemnica_TX = SavePrijemnica(datum, kupacID, vozacID, brojPrij, _
                                        brojZbirne, vrstaVoca, sortaVoca, _
                                        kolicina, cena, tipAmb, kolAmb, _
                                        kolAmbVracena, klasa, brutoKg)

    If SavePrijemnica_TX = "" Then
        Err.Raise vbObjectError + 1011, "SavePrijemnica_TX", _
                  "SavePrijemnica fehlgeschlagen"
    End If

    ' Paletizacija UNUTAR transakcije -> atomicno sa prijemnicom (gajbice bilo koje klase).
    Dim closedPal As Collection
    Set closedPal = New Collection
    If kolAmb > 0 Then
        PaletizePrijemnica prijemnicaID:=SavePrijemnica_TX, brojPrij:=brojPrij, _
            brojZbirne:=brojZbirne, vrstaVoca:=vrstaVoca, sortaVoca:=sortaVoca, _
            klasa:=klasa, netoKg:=kolicina, brGajbica:=kolAmb, tipAmb:=tipAmb, _
            closedPalIDs:=closedPal
    End If

    tx.CommitTx

    ' Print/PDF zatvorenih paleta = post-commit side effect (bez rollback rizika).
    PaletniListOutputClosed closedPal

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SavePrijemnica_TX"
    Monitor_Error _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnica_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnica_TX, _
        correlationId:=brojPrij, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="DOKUMENT_SAVE_FAIL", _
        severity:="ERROR", _
        message:="SavePrijemnica_TX failed. BrojPrij=" & brojPrij & _
             "; BrojZbirne=" & brojZbirne & _
             "; KupacID=" & kupacID & _
             "; VozacID=" & vozacID & _
             "; Error=" & errDesc, _
        userId:="Operator", _
        moduleName:="modDokumenta", _
        procedureName:="SavePrijemnica_TX", _
        entityType:="Prijemnica", _
        entityID:=SavePrijemnica_TX, _
        correlationId:=brojPrij

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SavePrijemnica_TX = ""

    PrintTxFailure "SavePrijemnica_TX", errSrc, errNum, errDesc
End Function
    
Public Function SavePrijemnica(ByVal datum As Date, ByVal kupacID As String, _
                               ByVal vozacID As String, ByVal brojPrij As String, _
                               ByVal brojZbirne As String, ByVal vrstaVoca As String, _
                               ByVal sortaVoca As String, ByVal kolicina As Double, _
                               ByVal cena As Double, ByVal tipAmb As String, _
                               ByVal kolAmb As Long, ByVal kolAmbVracena As Long, _
                               Optional ByVal klasa As String = "I", _
                               Optional ByVal brutoKg As Double = 0) As String
    On Error GoTo EH

    Call ValidatePrijemnicaInput(kupacID, vozacID, brojPrij, brojZbirne, _
                             kolicina, cena, tipAmb, kolAmb, _
                             kolAmbVracena, klasa)

    Dim newID As String
    newID = GetNextID(TBL_PRIJEMNICA, COL_PRJ_ID, "PRJ-")

    If newID = "" Then
        Err.Raise vbObjectError + 1013, "SavePrijemnica", _
                  "GetNextID nije vratio PrijemnicaID."
    End If

    Dim rowData As Variant
    rowData = BuildPrijemnicaRowData( _
                newID, datum, kupacID, vozacID, brojPrij, brojZbirne, _
                vrstaVoca, sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                kolAmbVracena, klasa)


    Dim appendedRow As Long
    appendedRow = AppendRow(TBL_PRIJEMNICA, rowData)

    If appendedRow <= 0 Then
        Err.Raise vbObjectError + 1014, "SavePrijemnica", _
                "AppendRow fehlgeschlagen fuer tblPrijemnica."
    End If

    ' Bruto tezina (preneto iz otkupa kad je OTKUP_BRUTO_UNOS) -> upis po imenu;
    ' prazno = neto. Kolona postoji posle EnsureDoradeSchema (na kraju tblPrijemnica).
    If brutoKg > 0 Then UpdateCell TBL_PRIJEMNICA, appendedRow, COL_PRJ_BRUTO, brutoKg

    ' Ambalaza je ENTITETSKI-relativna (smer iz ugla hladnjace / Kupca):
    ' 1. txt = pune gajbe koje hladnjaca PRIMA od zbirne -> Kupac ULAZ.
    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, "Ulaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    ' 2. txt = zamena: prazne gajbe koje hladnjaca VRACA (daje vozacu) -> Kupac IZLAZ.
    If kolAmbVracena > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmbVracena, "Izlaz", kupacID, "Kupac", _
                      vozacID, newID, DOK_TIP_PRIJEMNICA
    End If

    RelinkFakturaStavke newID, brojPrij, klasa

    SavePrijemnica = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "SavePrijemnica"
    On Error GoTo 0

    Err.Raise errNum, "SavePrijemnica", _
              "Source=" & errSrc & " | " & errDesc
End Function

Private Function BuildPrijemnicaRowData(ByVal prijemnicaID As String, _
                                        ByVal datum As Date, _
                                        ByVal kupacID As String, _
                                        ByVal vozacID As String, _
                                        ByVal brojPrij As String, _
                                        ByVal brojZbirne As String, _
                                        ByVal vrstaVoca As String, _
                                        ByVal sortaVoca As String, _
                                        ByVal kolicina As Double, _
                                        ByVal cena As Double, _
                                        ByVal tipAmb As String, _
                                        ByVal kolAmb As Long, _
                                        ByVal kolAmbVracena As Long, _
                                        ByVal klasa As String) As Variant
    Const SRC As String = "BuildPrijemnicaRowData"

    Dim colCount As Long
    colCount = GetDokumentaTableColumnCount(TBL_PRIJEMNICA)

    If colCount <= 0 Then
        Err.Raise vbObjectError + 1430, SRC, _
                  "Could not resolve tblPrijemnica column count."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KUPAC, kupacID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_VOZAC, vozacID, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_BROJ, brojPrij, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, brojZbirne, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_VRSTA, vrstaVoca, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_SORTA, sortaVoca, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOLICINA, kolicina, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_CENA, cena, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_TIP_AMB, tipAmb, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOL_AMB, kolAmb, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA, kolAmbVracena, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_KLASA, klasa, SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, "Ne", SRC
    SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, "", SRC

    If GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO) > 0 Then
        SetRowValueByColumn rowData, TBL_PRIJEMNICA, COL_STORNIRANO, "", SRC
    End If

    BuildPrijemnicaRowData = rowData
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
    
    data = ExcludeStornirano(data, TBL_PRIJEMNICA)

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
                  Poruka("DOK_ERR_NEMA_AMBALAZE_NOVCA")
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
            partnerId:=kupacID, _
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
' MANJAK - Schwundberechnung
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
' PROSEK GAJBE - Durchschnittsgewicht pro Kaestchen
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

    ' Dvoklasni doc: sumiraj kol i amb preko SVIH redova broja (ne samo prvi red).
    Dim kol As Double, amb As Double
    kol = SumByBroj(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOLICINA)
    amb = SumByBroj(TBL_OTPREMNICA, COL_OTP_BROJ, brojOtp, COL_OTP_KOL_AMB)

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

    ' Dvoklasni doc: sumiraj kol i amb preko SVIH redova broja (ne samo prvi red).
    Dim kol As Double, amb As Double
    kol = SumByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOLICINA)
    amb = SumByBroj(TBL_ZBIRNA, COL_ZBR_BROJ, brojZbirne, COL_ZBR_KOL_AMB)

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

' Sumira valCol preko SVIH redova gde brojCol = broj (obe klase dvorednog doc-a).
Private Function SumByBroj(ByVal tbl As String, ByVal brojCol As String, _
                           ByVal broj As String, ByVal valCol As String) As Double
    Dim data As Variant: data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function
    Dim cB As Long, cV As Long
    cB = GetColumnIndex(tbl, brojCol)
    cV = GetColumnIndex(tbl, valCol)
    If cB = 0 Or cV = 0 Then Exit Function
    Dim i As Long, s As Double
    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, cB))) = Trim$(broj) Then
            If IsNumeric(data(i, cV)) Then s = s + CDbl(data(i, cV))
        End If
    Next i
    SumByBroj = s
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
                               ByVal brojPrijemnice As String, _
                               Optional ByVal klasaFilter As String = "")
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
    Dim colFsKlasa As Long

    colPrijID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, _
                                   "modDokumenta.RelinkFakturaStavke")
    colOsir = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD, _
                                 "modDokumenta.RelinkFakturaStavke")
    colFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, _
                                  "modDokumenta.RelinkFakturaStavke")
    colFsKlasa = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KLASA, _
                                "modDokumenta.RelinkFakturaStavke")

    Dim i As Long
    Dim oldPrijID As String
    Dim oldBroj As String
    Dim fakID As String

    For i = 1 To UBound(stavkeData, 1)
    
        Dim stavkaKlasa As String
        stavkaKlasa = Trim$(NzToText(stavkeData(i, colFsKlasa)))
        If Len(Trim$(klasaFilter)) > 0 Then
            If UCase$(stavkaKlasa) <> UCase$(Trim$(klasaFilter)) Then GoTo NextStavka
        End If

        If Trim$(NzToText(stavkeData(i, colOsir))) = "" Then GoTo NextStavka

        oldPrijID = Trim$(NzToText(stavkeData(i, colPrijID)))

        If oldPrijID = "" Then GoTo NextStavka

        oldBroj = CStr(LookupValue(TBL_PRIJEMNICA, COL_PRJ_ID, oldPrijID, COL_PRJ_BROJ))

        If oldBroj = brojPrijemnice Then
            fakID = Trim$(NzToText(stavkeData(i, colFakID)))
            
            If Len(Trim$(fakID)) = 0 Then
            Err.Raise vbObjectError + 7413, "modDokumenta.RelinkFakturaStavke", _
                    "FakturaStavka nema FakturaID za relink. BrojPrijemnice=" & _
                    brojPrijemnice & "; Klasa=" & stavkaKlasa
            End If

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_FS_PRIJEMNICA_ID, _
                              newPrijemnicaID, "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_FAKTURA_STAVKE, i, COL_OSIROCENO_OD, _
                              "", "modDokumenta.RelinkFakturaStavke"

            Dim newPrijRow As Long
            newPrijRow = FindPrijemnicaRowByIDAndKlasa(newPrijemnicaID, stavkaKlasa, _
                                           "modDokumenta.RelinkFakturaStavke")

            If newPrijRow <= 0 Then
                Err.Raise vbObjectError + 7410, "modDokumenta.RelinkFakturaStavke", _
                        "Nova prijemnica nije pronadena za relink. PrijemnicaID=" & _
                        newPrijemnicaID & "; Klasa=" & stavkaKlasa
            End If

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRow, COL_PRJ_FAKTURISANO, _
                            "Da", "modDokumenta.RelinkFakturaStavke"

            RequireUpdateCell TBL_PRIJEMNICA, newPrijRow, COL_PRJ_FAKTURA_ID, _
                            fakID, "modDokumenta.RelinkFakturaStavke"


            If Len(Trim$(fakID)) > 0 Then
                UpdateFakturaStatus fakID
            End If

        End If

NextStavka:
    Next i

    Exit Sub

EH:
    LogErr "modDokumenta.RelinkFakturaStavke"
    Err.Raise Err.Number, "modDokumenta.RelinkFakturaStavke", Err.description
End Sub

' ============================================================
' HELPER - Vozac-Report (ersetzt alten modTransport)
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

    data = ExcludeStornirano(data, TBL_OTPREMNICA)

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
    
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)

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

Private Sub PrintTxFailure(ByVal sourceName As String, _
                           ByVal errSrc As String, _
                           ByVal errNum As Long, _
                           ByVal errDesc As String)
    Debug.Print sourceName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Sub

Private Sub RequireValidDocumentClass(ByVal klasa As String, _
                                      ByVal sourceName As String)

    Select Case Trim$(CStr(klasa))
        Case KLASA_I, KLASA_II
            Exit Sub
    End Select

    Err.Raise vbObjectError + 1400, sourceName, _
              "Neispravna klasa dokumenta: " & klasa
End Sub

Private Sub ValidateOtpremnicaInput(ByVal stanicaID As String, _
                                    ByVal vozacID As String, _
                                    ByVal brojOtp As String, _
                                    ByVal brojZbirne As String, _
                                    ByVal kolicina As Double, _
                                    ByVal cena As Double, _
                                    ByVal tipAmb As String, _
                                    ByVal kolAmb As Long, _
                                    ByVal klasa As String)

    Const SRC As String = "ValidateOtpremnicaInput"

    If Len(Trim$(stanicaID)) = 0 Then
        Err.Raise vbObjectError + 1401, SRC, "StanicaID je obavezan."
    End If

    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise vbObjectError + 1402, SRC, "VozacID je obavezan."
    End If

    If Len(Trim$(brojOtp)) = 0 Then
        Err.Raise vbObjectError + 1403, SRC, "Broj otpremnice je obavezan."
    End If

    'If Len(Trim$(brojZbirne)) = 0 Then
    '    Err.Raise vbObjectError + 1404, SRC, "Broj zbirne je obavezan."
    'End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1405, SRC, "Koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If cena < 0 Then
        Err.Raise vbObjectError + 1406, SRC, "Cena ne sme biti negativna."
    End If

    If kolAmb < 0 Then
        Err.Raise vbObjectError + 1407, SRC, "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If kolAmb > 0 And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 1408, SRC, "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    RequireValidDocumentClass klasa, SRC
End Sub

Private Sub ValidateZbirnaInput(ByVal vozacID As String, _
                                ByVal brojZbirne As String, _
                                ByVal kupacID As String, _
                                ByVal ukupnoKol As Double, _
                                ByVal tipAmb As String, _
                                ByVal ukupnoAmb As Long, _
                                ByVal klasa As String)

    Const SRC As String = "ValidateZbirnaInput"

    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise vbObjectError + 1410, SRC, "VozacID je obavezan."
    End If

    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise vbObjectError + 1411, SRC, "Broj zbirne je obavezan."
    End If

    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise vbObjectError + 1412, SRC, "KupacID je obavezan."
    End If

    If ukupnoKol <= 0 Then
        Err.Raise vbObjectError + 1413, SRC, "Ukupna koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If ukupnoAmb < 0 Then
        Err.Raise vbObjectError + 1414, SRC, "Ukupna ambala" & ChrW(382) & "a ne sme biti negativna."
    End If

    If ukupnoAmb > 0 And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 1415, SRC, "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    RequireValidDocumentClass klasa, SRC
End Sub

Private Sub ValidatePrijemnicaInput(ByVal kupacID As String, _
                                    ByVal vozacID As String, _
                                    ByVal brojPrij As String, _
                                    ByVal brojZbirne As String, _
                                    ByVal kolicina As Double, _
                                    ByVal cena As Double, _
                                    ByVal tipAmb As String, _
                                    ByVal kolAmb As Long, _
                                    ByVal kolAmbVracena As Long, _
                                    ByVal klasa As String)

    Const SRC As String = "ValidatePrijemnicaInput"

    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise vbObjectError + 1420, SRC, "KupacID je obavezan."
    End If

    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise vbObjectError + 1421, SRC, "VozacID je obavezan."
    End If

    If Len(Trim$(brojPrij)) = 0 Then
        Err.Raise vbObjectError + 1422, SRC, "Broj prijemnice je obavezan."
    End If

    If Len(Trim$(brojZbirne)) = 0 Then
        Err.Raise vbObjectError + 1423, SRC, "Broj zbirne je obavezan."
    End If

    ' Referencijalni integritet: broj zbirne mora da postoji u tblZbirna.
    ' Samo u BLOK modu -- u UPOZORENJE modu forma trazi potvrdu, pa backend ne
    ' sme tvrdo da padne. Ujedno je bezbedna mreza i za ne-form pozivaoce.
    If PrijemnicaZbirnaBlokira() Then
        If Not ZbirnaPostoji(brojZbirne) Then
            Err.Raise vbObjectError + 1428, SRC, _
                "Zbirna '" & brojZbirne & "' ne postoji u sistemu."
        End If
    End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1424, SRC, "Koli" & ChrW(269) & "ina mora biti veca od nule."
    End If

    If cena < 0 Then
        Err.Raise vbObjectError + 1425, SRC, "Cena ne sme biti negativna."
    End If

    If kolAmb < 0 Or kolAmbVracena < 0 Then
        Err.Raise vbObjectError + 1426, SRC, "Koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e ne sme biti negativna."
    End If

    If (kolAmb > 0 Or kolAmbVracena > 0) And Len(Trim$(tipAmb)) = 0 Then
        Err.Raise vbObjectError + 1427, SRC, "Tip ambala" & ChrW(382) & "e je obavezan kada postoji ambala" & ChrW(382) & "a."
    End If

    RequireValidDocumentClass klasa, SRC
End Sub

Private Sub SetRowValueByColumn(ByRef rowData() As Variant, _
                                ByVal tableName As String, _
                                ByVal columnName As String, _
                                ByVal value As Variant, _
                                ByVal sourceName As String)
    Dim colIndex As Long
    colIndex = RequireColumnIndex(tableName, columnName, sourceName)

    rowData(colIndex - 1) = value
End Sub

Private Function GetDokumentaTableColumnCount(ByVal tableName As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                GetDokumentaTableColumnCount = lo.ListColumns.count
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function FindPrijemnicaRowByIDAndKlasa(ByVal prijemnicaID As String, _
                                               ByVal klasa As String, _
                                               ByVal sourceName As String) As Long
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then Exit Function

    Dim colID As Long
    Dim colKlasa As Long

    colID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, sourceName)
    colKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, sourceName)

    Dim i As Long
    Dim foundRow As Long
    Dim foundCount As Long

    For i = 1 To UBound(data, 1)
        If Trim$(NzToText(data(i, colID))) = Trim$(prijemnicaID) And _
           Trim$(NzToText(data(i, colKlasa))) = Trim$(klasa) Then

            foundRow = i
            foundCount = foundCount + 1
        End If
    Next i

    If foundCount > 1 Then
        Err.Raise vbObjectError + 7412, sourceName, _
                  "Dupli redovi za PrijemnicaID + Klasa. PrijemnicaID=" & _
                  prijemnicaID & "; Klasa=" & klasa
    End If

    FindPrijemnicaRowByIDAndKlasa = foundRow
End Function

' ============================================================
' STORNO PREGLED (read-only) -- agregira stornirane dokumente po tipu za
' prikaz u panelu unutar frmDokumenta (dugme "Pregled storniranih").
' Soft-delete: red je storniran kad je COL_STORNIRANO = "Da" (modStorno).
' Jedinstven (unifikovan) skup korisnih kolona za sve tipove:
'   Broj | Datum | Partner | Vrsta | Sorta | Klasa | Kolicina | Cena | Iznos
'   | Zbirna | Otpremnica | Faktura  (poslednje 3 = lanac zavisnih dokumenata)
' Partner se razresava na naziv/ime (best-effort; fallback = ID).
' ============================================================

' Tipovi storno dokumenata u redosledu prikaza (isti nazivi kao cmbStornoDokument).
Public Function StorniraniTipovi() As Variant
    StorniraniTipovi = Array("Otkup", "Otpremnica", "Zbirna", "Prijemnica", "Faktura", "Novac", _
                         "Revers izdavanje koop.", "Revers povrat koop.", _
                         "Revers izdato OM (firma).", "Revers prijem od OM (firma).")
End Function

' Zaglavlja unifikovanih kolona (0-bazni niz, 12 kolona).
Public Function StorniraniHeaders() As Variant
    StorniraniHeaders = Array("Broj", "Datum", "Partner", "Vrsta", "Sorta", _
                              "Klasa", "Koli" & ChrW(269) & "ina", "Cena", "Iznos (RSD)", _
                              "Zbirna", "Otpremnica", "Faktura")
End Function

' Vrati 2D niz (1..n, 1..12) storniranih redova za dati tip u unifikovanom
' rasporedu kolona (osnovne + lanac: Zbirna/Otpremnica/Faktura). Empty ako nema.
' chainIdx: pre-izgradjen indeks lanca (BuildChainIndex); ako fali, gradi se sam.
Public Function GetStorniraniByTip(ByVal tip As String, _
                                   Optional ByVal chainIdx As Object = Nothing) As Variant
    On Error GoTo EH
    If chainIdx Is Nothing Then Set chainIdx = BuildChainIndex()

    ' OM<->koop revers (ambalaza) ima drugaciju strukturu (2 noge, EntitetID/Tip) ->
    ' poseban scan tblAmbalaze.
    Select Case tip
        Case "Revers izdavanje koop."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_IZLAZ_KOOP)
            Exit Function
        Case "Revers povrat koop."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_ULAZ_KOOP)
            Exit Function
        Case "Revers izdato OM (firma)."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_ULAZ_FIRMA, "Stanica")
            Exit Function
        Case "Revers prijem od OM (firma)."
            GetStorniraniByTip = GetStorniraniRevers(DOK_TIP_OM_IZLAZ_FIRMA, "Stanica")
            Exit Function
    End Select

    Dim tbl As String
    Dim iBroj As Long, iDat As Long, iPart As Long, iVrsta As Long, iSorta As Long
    Dim iKlasa As Long, iKol As Long, iCena As Long, iIzn As Long, iIzn2 As Long
    Dim iZbr As Long, iFak As Long
    Dim partnerIsName As Boolean
    Dim pdict As Object

    Select Case tip
        Case "Otkup"
            tbl = TBL_OTKUP
            Set pdict = BuildIdNameDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
            iBroj = GetColumnIndex(tbl, COL_OTK_BR_DOK)
            iDat = GetColumnIndex(tbl, COL_OTK_DATUM)
            iPart = GetColumnIndex(tbl, COL_OTK_KOOPERANT)
            iVrsta = GetColumnIndex(tbl, COL_OTK_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_OTK_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_OTK_KLASA)
            iKol = GetColumnIndex(tbl, COL_OTK_KOLICINA)
            iCena = GetColumnIndex(tbl, COL_OTK_CENA)
            iIzn = GetColumnIndex(tbl, COL_OTK_NOVAC)
            iZbr = GetColumnIndex(tbl, COL_OTK_BROJ_ZBIRNE)

        Case "Otpremnica"
            tbl = TBL_OTPREMNICA
            Set pdict = BuildIdNameDict(TBL_STANICE, "StanicaID", "Naziv")
            iBroj = GetColumnIndex(tbl, COL_OTP_BROJ)
            iDat = GetColumnIndex(tbl, COL_OTP_DATUM)
            iPart = GetColumnIndex(tbl, COL_OTP_STANICA)
            iVrsta = GetColumnIndex(tbl, COL_OTP_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_OTP_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_OTP_KLASA)
            iKol = GetColumnIndex(tbl, COL_OTP_KOLICINA)
            iCena = GetColumnIndex(tbl, COL_OTP_CENA)
            iZbr = GetColumnIndex(tbl, COL_OTP_BROJ_ZBIRNE)

        Case "Zbirna"
            tbl = TBL_ZBIRNA
            Set pdict = BuildIdNameDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
            iBroj = GetColumnIndex(tbl, COL_ZBR_BROJ)
            iDat = GetColumnIndex(tbl, COL_ZBR_DATUM)
            iPart = GetColumnIndex(tbl, COL_ZBR_KUPAC)
            iVrsta = GetColumnIndex(tbl, COL_ZBR_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_ZBR_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_ZBR_KLASA)
            iKol = GetColumnIndex(tbl, COL_ZBR_KOLICINA)
            iZbr = GetColumnIndex(tbl, COL_ZBR_BROJ)        ' sopstveni broj = pivot zbirne

        Case "Prijemnica"
            tbl = TBL_PRIJEMNICA
            Set pdict = BuildIdNameDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
            iBroj = GetColumnIndex(tbl, COL_PRJ_BROJ)
            iDat = GetColumnIndex(tbl, COL_PRJ_DATUM)
            iPart = GetColumnIndex(tbl, COL_PRJ_KUPAC)
            iVrsta = GetColumnIndex(tbl, COL_PRJ_VRSTA)
            iSorta = GetColumnIndex(tbl, COL_PRJ_SORTA)
            iKlasa = GetColumnIndex(tbl, COL_PRJ_KLASA)
            iKol = GetColumnIndex(tbl, COL_PRJ_KOLICINA)
            iCena = GetColumnIndex(tbl, COL_PRJ_CENA)
            iZbr = GetColumnIndex(tbl, COL_PRJ_BROJ_ZBIRNE)
            iFak = GetColumnIndex(tbl, COL_PRJ_FAKTURA_ID)

        Case "Faktura"
            tbl = TBL_FAKTURE
            Set pdict = BuildIdNameDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
            iBroj = GetColumnIndex(tbl, COL_FAK_BROJ)
            iDat = GetColumnIndex(tbl, COL_FAK_DATUM)
            iPart = GetColumnIndex(tbl, COL_FAK_KUPAC)
            iIzn = GetColumnIndex(tbl, COL_FAK_IZNOS)
            iFak = GetColumnIndex(tbl, COL_FAK_ID)          ' sopstveni FakturaID

        Case "Novac"
            tbl = TBL_NOVAC
            partnerIsName = True       ' Partner je vec naziv (ne ID)
            iBroj = GetColumnIndex(tbl, COL_NOV_BROJ_DOK)
            iDat = GetColumnIndex(tbl, COL_NOV_DATUM)
            iPart = GetColumnIndex(tbl, COL_NOV_PARTNER)
            iVrsta = GetColumnIndex(tbl, COL_NOV_VRSTA)
            iIzn = GetColumnIndex(tbl, COL_NOV_UPLATA)
            iIzn2 = GetColumnIndex(tbl, COL_NOV_ISPLATA)    ' Iznos = Uplata - Isplata
            iFak = GetColumnIndex(tbl, COL_NOV_FAKTURA_ID)

        Case Else
            Exit Function
    End Select

    Dim data As Variant
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim colStorno As Long
    colStorno = GetColumnIndex(tbl, COL_STORNIRANO)
    If colStorno = 0 Then Exit Function        ' tabela bez storno markera -> nista

    Dim otpByZbr As Object: Set otpByZbr = chainIdx("otpByZbr")
    Dim fakByZbr As Object: Set fakByZbr = chainIdx("fakByZbr")
    Dim fakById As Object:  Set fakById = chainIdx("fakById")
    Dim zbrByFak As Object: Set zbrByFak = chainIdx("zbrByFak")

    Dim rows As Collection
    Set rows = New Collection

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(NzToText(data(i, colStorno)))) = "DA" Then
            Dim part As String
            If partnerIsName Then
                part = StornoCellText(data, i, iPart)
            Else
                part = ResolveNameFromDict(pdict, StornoCellRaw(data, i, iPart))
            End If

            ' Lanac zavisnih dokumenata preko BrojZbirne / FakturaID.
            Dim zbr As String, fakId As String, otp As String, fak As String
            zbr = StornoCellText(data, i, iZbr)
            fakId = StornoCellText(data, i, iFak)
            If Len(zbr) = 0 And Len(fakId) > 0 Then zbr = DictGet(zbrByFak, fakId)
            otp = DictGet(otpByZbr, zbr)
            If Len(fakId) > 0 Then fak = DictGet(fakById, fakId) Else fak = DictGet(fakByZbr, zbr)

            ' Iznos: stored (Faktura=Iznos, Novac=Uplata-Isplata) ili izracunat
            ' Kolicina x Cena (Otkup/Otpremnica/Prijemnica nemaju zaseban iznos).
            Dim iznos As String
            iznos = StornoIznosText(StornoCellRaw(data, i, iIzn), StornoCellRaw(data, i, iIzn2))
            If Len(iznos) = 0 Then _
                iznos = StornoMnozi(StornoCellRaw(data, i, iKol), StornoCellRaw(data, i, iCena))

            rows.Add Array( _
                StornoCellText(data, i, iBroj), _
                StornoDateText(StornoCellRaw(data, i, iDat)), _
                part, _
                StornoCellText(data, i, iVrsta), _
                StornoCellText(data, i, iSorta), _
                StornoCellText(data, i, iKlasa), _
                StornoNumText(StornoCellRaw(data, i, iKol), "#,##0.00"), _
                StornoNumText(StornoCellRaw(data, i, iCena), "#,##0.00"), _
                iznos, _
                zbr, otp, fak)
        End If
    Next i

    GetStorniraniByTip = StornoRowsTo2D(rows, 12)
    Exit Function

EH:
    LogErr "modDokumenta.GetStorniraniByTip(" & tip & ")"
    GetStorniraniByTip = Empty
End Function

' Stornirani OM<->koop revers (dokTip) iz tblAmbalaza -> unifikovane 12 kolona.
' Jedan red po dokumentu: Kooperant noga (nosi partnera); Stanica noga se preskace.
' Mapiranje: Vrsta=TipAmbalaze, Kolicina=broj gajbica; ostalo prazno.
' Preskace OM-Izlaz-Koop noge knjizene UZ otkup (DokumentID = otkupID "OTK-...";
' vec se vide pod grupom 'Otkup') -> ostaju samo standalone reversi (broj x/ddmmyy).
Private Function GetStorniraniRevers(ByVal dokTip As String, _
                                     Optional ByVal entType As String = "Kooperant") As Variant
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim iBroj As Long, iDat As Long, iEntID As Long, iEntTip As Long
    Dim iDokTip As Long, iTipAmb As Long, iKol As Long, iStorno As Long
    iBroj = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    iDat = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    iEntID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    iEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    iDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    iTipAmb = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    iKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    iStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    If iBroj = 0 Or iDokTip = 0 Or iStorno = 0 Then Exit Function

    Dim pdict As Object
    If entType = "Stanica" Then
        Set pdict = BuildIdNameDict(TBL_STANICE, "StanicaID", "Naziv")
    Else
        Set pdict = BuildIdNameDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    End If
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If UCase$(Trim$(NzToText(data(i, iStorno)))) = "DA" Then
            If Trim$(CStr(data(i, iDokTip))) = dokTip And _
               Trim$(CStr(data(i, iEntTip))) = entType And _
               UCase$(Left$(Trim$(CStr(data(i, iBroj))), 3)) <> "OTK" Then
                rows.Add Array( _
                    StornoCellText(data, i, iBroj), _
                    StornoDateText(StornoCellRaw(data, i, iDat)), _
                    ResolveNameFromDict(pdict, StornoCellRaw(data, i, iEntID)), _
                    StornoCellText(data, i, iTipAmb), _
                    "", "", _
                    StornoNumText(StornoCellRaw(data, i, iKol), "#,##0"), _
                    "", "", "", "", "")
            End If
        End If
    Next i
    GetStorniraniRevers = StornoRowsTo2D(rows, 12)
    Exit Function
EH:
    LogErr "modDokumenta.GetStorniraniRevers(" & dokTip & ")"
    GetStorniraniRevers = Empty
End Function

' Svi stornirani grupisano: Collection of Array(tipNaziv, rows2D(1..n,1..12)).
' Indeks lanca se gradi JEDNOM (reverzni cross-table lookup-i) i deli svim tipovima.
Public Function GetStorniraniGrupisano() As Collection
    On Error GoTo EH
    Dim res As Collection
    Set res = New Collection
    Set GetStorniraniGrupisano = res

    Dim idx As Object: Set idx = BuildChainIndex()
    Dim tipovi As Variant: tipovi = StorniraniTipovi()
    Dim k As Long
    For k = LBound(tipovi) To UBound(tipovi)
        Dim tip As String: tip = CStr(tipovi(k))
        Dim rowsv As Variant: rowsv = GetStorniraniByTip(tip, idx)
        If Not IsEmpty(rowsv) Then res.Add Array(tip, rowsv)
    Next k
    Exit Function
EH:
    LogErr "modDokumenta.GetStorniraniGrupisano"
End Function

' id -> "Naziv" (ili "Ime Prezime") recnik; prazan recnik ako tabela/kolone fale.
' Public: reuse i iz modIntegritet (OtkupnoMestoByZbirna).
Public Function BuildIdNameDict(ByVal tbl As String, ByVal idCol As String, _
                                ByVal nameCol1 As String, _
                                Optional ByVal nameCol2 As String = "") As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set BuildIdNameDict = d

    Dim data As Variant
    data = GetTableData(tbl)
    If IsEmpty(data) Then Exit Function

    Dim ci As Long, c1 As Long, c2 As Long
    ci = GetColumnIndex(tbl, idCol)
    c1 = GetColumnIndex(tbl, nameCol1)
    If Len(nameCol2) > 0 Then c2 = GetColumnIndex(tbl, nameCol2)
    If ci = 0 Or c1 = 0 Then Exit Function

    Dim i As Long, idv As String, nm As String
    For i = 1 To UBound(data, 1)
        idv = Trim$(NzToText(data(i, ci)))
        If Len(idv) > 0 Then
            nm = Trim$(NzToText(data(i, c1)))
            If c2 > 0 Then nm = Trim$(nm & " " & NzToText(data(i, c2)))
            If Not d.Exists(idv) Then d.Add idv, nm
        End If
    Next i
End Function

Private Function ResolveNameFromDict(ByVal d As Object, ByVal idRaw As Variant) As String
    Dim s As String
    s = Trim$(NzToText(idRaw))
    If Len(s) = 0 Then Exit Function
    If Not d Is Nothing Then
        If d.Exists(s) Then
            If Len(Trim$(NzToText(d(s)))) > 0 Then
                ResolveNameFromDict = d(s)
                Exit Function
            End If
        End If
    End If
    ResolveNameFromDict = s          ' fallback: prikazi ID ako nema imena
End Function

' Sirova vrednost celije (Empty ako kolona ne postoji -> idx=0).
Private Function StornoCellRaw(ByVal data As Variant, ByVal r As Long, ByVal idx As Long) As Variant
    If idx = 0 Then Exit Function
    StornoCellRaw = data(r, idx)
End Function

Private Function StornoCellText(ByVal data As Variant, ByVal r As Long, ByVal idx As Long) As String
    If idx = 0 Then Exit Function
    StornoCellText = Trim$(NzToText(data(r, idx)))
End Function

Private Function StornoDateText(ByVal v As Variant) As String
    If IsDate(v) Then
        StornoDateText = Format$(CDate(v), "d.m.yyyy")
    Else
        StornoDateText = Trim$(NzToText(v))
    End If
End Function

Private Function StornoNumText(ByVal v As Variant, ByVal fmt As String) As String
    Dim d As Double
    If TryParseDouble(Trim$(NzToText(v)), d) Then
        If d <> 0 Then StornoNumText = Format$(d, fmt)
    End If
End Function

' Iznos = v1 - v2 (Novac: Uplata - Isplata; ostali: v2 prazno -> samo v1).
Private Function StornoIznosText(ByVal v1 As Variant, ByVal v2 As Variant) As String
    Dim u As Double, s As Double
    If Not TryParseDouble(Trim$(NzToText(v1)), u) Then u = 0
    If Not TryParseDouble(Trim$(NzToText(v2)), s) Then s = 0
    Dim net As Double: net = u - s
    If net <> 0 Then StornoIznosText = Format$(net, "#,##0.00")
End Function

' Iznos = Kolicina x Cena (prazno ako je proizvod 0).
Private Function StornoMnozi(ByVal vKol As Variant, ByVal vCena As Variant) As String
    Dim kol As Double, cena As Double
    If Not TryParseDouble(Trim$(NzToText(vKol)), kol) Then kol = 0
    If Not TryParseDouble(Trim$(NzToText(vCena)), cena) Then cena = 0
    Dim p As Double: p = kol * cena
    If p <> 0 Then StornoMnozi = Format$(p, "#,##0.00")
End Function

' --- Indeks lanca dokumenata (reverzni lookup-i preko BrojZbirne / FakturaID) ---

' Prazan recnik sa case-insensitive poredjenjem kljuceva.
Private Function NewDict() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NewDict = d
End Function

Private Function DictGet(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If Len(key) = 0 Then Exit Function
    If d.Exists(key) Then DictGet = CStr(d(key))
End Function

' Dodaj val u listu pod key (", "-spojeno, bez duplikata).
Private Sub DictAppend(ByVal d As Object, ByVal key As String, ByVal val As String)
    If d Is Nothing Then Exit Sub
    If Len(key) = 0 Or Len(val) = 0 Then Exit Sub
    If Not d.Exists(key) Then
        d.Add key, val
    Else
        Dim cur As String: cur = CStr(d(key))
        If InStr(1, ", " & cur & ", ", ", " & val & ", ", vbTextCompare) = 0 Then
            d(key) = cur & ", " & val
        End If
    End If
End Sub

' Izgradi indeks lanca: otpByZbr, fakByZbr, fakById, zbrByFak (sve case-insensitive).
Private Function BuildChainIndex() As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Set BuildChainIndex = idx

    Dim fakById As Object:  Set fakById = NewDict()
    Dim otpByZbr As Object: Set otpByZbr = NewDict()
    Dim fakByZbr As Object: Set fakByZbr = NewDict()
    Dim zbrByFak As Object: Set zbrByFak = NewDict()
    idx.Add "fakById", fakById
    idx.Add "otpByZbr", otpByZbr
    idx.Add "fakByZbr", fakByZbr
    idx.Add "zbrByFak", zbrByFak

    Dim d As Variant, i As Long, k As String, v As String

    ' Faktura: FakturaID -> BrojFakture
    d = GetTableData(TBL_FAKTURE)
    If Not IsEmpty(d) Then
        Dim cfi As Long, cfb As Long
        cfi = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
        cfb = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
        If cfi > 0 And cfb > 0 Then
            For i = 1 To UBound(d, 1)
                k = Trim$(NzToText(d(i, cfi))): v = Trim$(NzToText(d(i, cfb)))
                If Len(k) > 0 And Not fakById.Exists(k) Then fakById.Add k, v
            Next i
        End If
    End If

    ' Otpremnica: BrojZbirne -> lista BrojOtpremnice
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsEmpty(d) Then
        Dim coz As Long, cob As Long
        coz = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
        cob = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
        If coz > 0 And cob > 0 Then
            For i = 1 To UBound(d, 1)
                DictAppend otpByZbr, Trim$(NzToText(d(i, coz))), Trim$(NzToText(d(i, cob)))
            Next i
        End If
    End If

    ' Prijemnica: BrojZbirne <-> FakturaID  (i BrojZbirne -> BrojFakture preko fakById)
    d = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(d) Then
        Dim cpz As Long, cpf As Long
        cpz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        cpf = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID)
        If cpz > 0 And cpf > 0 Then
            For i = 1 To UBound(d, 1)
                Dim z As String, fid As String
                z = Trim$(NzToText(d(i, cpz))): fid = Trim$(NzToText(d(i, cpf)))
                If Len(fid) > 0 Then
                    DictAppend zbrByFak, fid, z
                    If fakById.Exists(fid) Then DictAppend fakByZbr, z, CStr(fakById(fid))
                End If
            Next i
        End If
    End If
End Function

' Kolekcija 0-baznih 1D nizova -> 2D (1..n, 1..ncol). Empty ako je prazna.
Private Function StornoRowsTo2D(ByVal rows As Collection, ByVal ncol As Long) As Variant
    If rows Is Nothing Then Exit Function
    If rows.count = 0 Then Exit Function

    Dim arr() As Variant
    ReDim arr(1 To rows.count, 1 To ncol)
    Dim r As Long, c As Long, one As Variant
    For r = 1 To rows.count
        one = rows(r)
        For c = 1 To ncol
            arr(r, c) = one(c - 1)
        Next c
    Next r
    StornoRowsTo2D = arr
End Function

' ============================================================
' IZGUBLJENI OTKUP BLOKOVI (read + bezbedan re-point)
' Blok je "izgubljen" kad mu OtpremnicaID pokazuje na storniranu ili
' nepostojecu otpremnicu (a sam blok nije storniran). Koriste: dashboard
' dijagnostika (modHelpers.CheckVerwaisteDokumente) i panel Otkupni blokovi.
' ============================================================

' Vrati 2D (1..n, 1..7): OtkupID | BrojDokumenta | KooperantID | Datum |
' Kolicina | StaraOtpremnicaID | StaraOtpremnicaBroj. Empty ako nema.
Public Function GetLostOtkupBlokovi() As Variant
    On Error GoTo EH

    Dim otk As Variant: otk = GetTableData(TBL_OTKUP)
    If IsEmpty(otk) Then Exit Function

    ' mape otpremnica po ID-u: stornirane i aktivne (ID -> BrojOtpremnice)
    Dim storMap As Object: Set storMap = CreateObject("Scripting.Dictionary")
    Dim aktMap As Object: Set aktMap = CreateObject("Scripting.Dictionary")
    storMap.CompareMode = vbTextCompare: aktMap.CompareMode = vbTextCompare

    Dim otp As Variant: otp = GetTableData(TBL_OTPREMNICA)
    If Not IsEmpty(otp) Then
        Dim oId As Long, oBr As Long, oSt As Long
        oId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
        oBr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
        oSt = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)
        Dim j As Long, idv As String
        For j = 1 To UBound(otp, 1)
            idv = Trim$(NzToText(otp(j, oId)))
            If Len(idv) > 0 Then
                If UCase$(Trim$(NzToText(otp(j, oSt)))) = "DA" Then
                    If Not storMap.Exists(idv) Then storMap.Add idv, Trim$(NzToText(otp(j, oBr)))
                ElseIf Not aktMap.Exists(idv) Then
                    aktMap.Add idv, Trim$(NzToText(otp(j, oBr)))
                End If
            End If
        Next j
    End If

    Dim cId As Long, cBr As Long, cKoop As Long, cDat As Long
    Dim cKol As Long, cOtp As Long, cStor As Long
    cId = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    cBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    cKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    cDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    cKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    cOtp = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
    cStor = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(otk, 1)
        If UCase$(Trim$(NzToText(otk(i, cStor)))) = "DA" Then GoTo NextRow
        Dim otpId As String: otpId = Trim$(NzToText(otk(i, cOtp)))
        If Len(otpId) = 0 Then GoTo NextRow            ' nije vezan -> nije "izgubljen"

        Dim staleBroj As String: staleBroj = ""
        Dim lost As Boolean: lost = False
        If storMap.Exists(otpId) Then
            lost = True: staleBroj = CStr(storMap(otpId))
        ElseIf Not aktMap.Exists(otpId) Then
            lost = True: staleBroj = "(nepostojeca)"
        End If

        If lost Then
            rows.Add Array( _
                Trim$(NzToText(otk(i, cId))), _
                Trim$(NzToText(otk(i, cBr))), _
                Trim$(NzToText(otk(i, cKoop))), _
                otk(i, cDat), _
                otk(i, cKol), _
                otpId, _
                staleBroj)
        End If
NextRow:
    Next i

    GetLostOtkupBlokovi = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetLostOtkupBlokovi"
    GetLostOtkupBlokovi = Empty
End Function

' Osirocene prijemnice: AKTIVNE prijemnice cija zbirna (BrojZbirne) vise nije
' aktivna (stornirana ili ne postoji). Jedan red po BrojPrijemnice (Klasa I+II
' dele broj). Za re-point UI (frmDokumenta recovery panel).
' Kolone 1..7: BrojPrijemnice|Datum|Vrsta|Sorta|Kolicina|StaraZbirna|Status
Public Function GetOsirocenePrijemnice() As Variant
    On Error GoTo EH

    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function

    ' BrojZbirne -> ima li bar jednu AKTIVNU zbirnu; i postoji li uopste.
    Dim aktZbr As Object: Set aktZbr = CreateObject("Scripting.Dictionary")
    Dim allZbr As Object: Set allZbr = CreateObject("Scripting.Dictionary")
    aktZbr.CompareMode = vbTextCompare: allZbr.CompareMode = vbTextCompare
    Dim zd As Variant: zd = GetTableData(TBL_ZBIRNA)
    If Not IsEmpty(zd) Then
        Dim zBr As Long, zSt As Long, zr As Long, zk As String
        zBr = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
        zSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
        For zr = 1 To UBound(zd, 1)
            zk = Trim$(NzToText(zd(zr, zBr)))
            If Len(zk) > 0 Then
                allZbr(zk) = True
                If UCase$(Trim$(NzToText(zd(zr, zSt)))) <> "DA" Then aktZbr(zk) = True
            End If
        Next zr
    End If

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long
    Dim cKol As Long, cZbr As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    cVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    cSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    cKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBr = 0 Or cZbr = 0 Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(prj, 1)
        Dim isStor As Boolean: isStor = False
        If cSt > 0 Then isStor = (UCase$(Trim$(NzToText(prj(i, cSt)))) = "DA")
        If Not isStor Then
            Dim bz As String: bz = Trim$(NzToText(prj(i, cZbr)))
            Dim brp As String: brp = Trim$(NzToText(prj(i, cBr)))
            If Len(bz) > 0 And Len(brp) > 0 Then
                If Not aktZbr.Exists(bz) Then
                    If Not seen.Exists(brp) Then
                        seen(brp) = True
                        Dim st As String
                        If allZbr.Exists(bz) Then st = "zbirna stornirana" Else st = "zbirna ne postoji"
                        rows.Add Array(brp, prj(i, cDat), Trim$(NzToText(prj(i, cVr))), _
                            Trim$(NzToText(prj(i, cSo))), StornoNumText(prj(i, cKol), "#,##0.00"), bz, st)
                    End If
                End If
            End If
        End If
    Next i

    GetOsirocenePrijemnice = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetOsirocenePrijemnice"
    GetOsirocenePrijemnice = Empty
End Function

' Aktivne (ne-stornirane) zbirne, jedan red po BrojZbirne (Klasa I+II dele broj).
' Za izbor cilja u re-point UI. Kolone 1..5: BrojZbirne|Datum|Vrsta|Sorta|Kolicina
Public Function GetAktivneZbirne() As Variant
    On Error GoTo EH

    Dim zd As Variant: zd = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zd) Then Exit Function

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long, cKol As Long, cSt As Long
    cBr = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    cDat = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM)
    cVr = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VRSTA)
    cSo = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_SORTA)
    cKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
    cSt = GetColumnIndex(TBL_ZBIRNA, COL_STORNIRANO)
    If cBr = 0 Then Exit Function

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Dim rows As Collection: Set rows = New Collection
    Dim i As Long
    For i = 1 To UBound(zd, 1)
        Dim isStor As Boolean: isStor = False
        If cSt > 0 Then isStor = (UCase$(Trim$(NzToText(zd(i, cSt)))) = "DA")
        If Not isStor Then
            Dim bz As String: bz = Trim$(NzToText(zd(i, cBr)))
            If Len(bz) > 0 Then
                If Not seen.Exists(bz) Then
                    seen(bz) = True
                    rows.Add Array(bz, zd(i, cDat), Trim$(NzToText(zd(i, cVr))), _
                        Trim$(NzToText(zd(i, cSo))), StornoNumText(zd(i, cKol), "#,##0.00"))
                End If
            End If
        End If
    Next i

    GetAktivneZbirne = StornoRowsTo2D(rows, 5)
    Exit Function
EH:
    LogErr "modDokumenta.GetAktivneZbirne"
    GetAktivneZbirne = Empty
End Function

' Stornirane prijemnice koje imaju AKTIVNE (osirocene) paleta-stavke. Za P1 pallet
' re-point. Kolone 1..6: BrojPrijemnice|Datum|Vrsta|Sorta|Gajbica|StavkiAktivnih
Public Function GetPrijemniceSaOsirocenimPaletama() As Variant
    On Error GoTo EH
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(ps) Then Exit Function

    Dim stPrij As Object: Set stPrij = CreateObject("Scripting.Dictionary"): stPrij.CompareMode = vbTextCompare
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prj) Then
        Dim qId As Long, qSt As Long, q As Long
        qId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
        qSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
        For q = 1 To UBound(prj, 1)
            If qSt > 0 Then
                If UCase$(Trim$(NzToText(prj(q, qSt)))) = "DA" Then stPrij(Trim$(NzToText(prj(q, qId)))) = True
            End If
        Next q
    End If

    Dim cBr As Long, cPid As Long, cVr As Long, cSo As Long, cGajb As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    cPid = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    cVr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_VRSTA)
    cSo = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_SORTA)
    cGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If cBr = 0 Or cPid = 0 Then Exit Function

    Dim gSum As Object: Set gSum = CreateObject("Scripting.Dictionary"): gSum.CompareMode = vbTextCompare
    Dim sCnt As Object: Set sCnt = CreateObject("Scripting.Dictionary"): sCnt.CompareMode = vbTextCompare
    Dim vrD As Object: Set vrD = CreateObject("Scripting.Dictionary"): vrD.CompareMode = vbTextCompare
    Dim soD As Object: Set soD = CreateObject("Scripting.Dictionary"): soD.CompareMode = vbTextCompare
    Dim order As Collection: Set order = New Collection
    Dim i As Long
    For i = 1 To UBound(ps, 1)
        Dim stx As Boolean: stx = False
        If cSt > 0 Then stx = (UCase$(Trim$(NzToText(ps(i, cSt)))) = "DA")
        If Not stx Then
            Dim pid As String: pid = Trim$(NzToText(ps(i, cPid)))
            Dim br As String: br = Trim$(NzToText(ps(i, cBr)))
            If stPrij.Exists(pid) And Len(br) > 0 Then
                If Not gSum.Exists(br) Then
                    gSum(br) = 0&: sCnt(br) = 0&: order.Add br
                    vrD(br) = Trim$(NzToText(ps(i, cVr))): soD(br) = Trim$(NzToText(ps(i, cSo)))
                End If
                If IsNumeric(ps(i, cGajb)) Then gSum(br) = CLng(gSum(br)) + CLng(ps(i, cGajb))
                sCnt(br) = CLng(sCnt(br)) + 1
            End If
        End If
    Next i

    Dim rows As Collection: Set rows = New Collection
    Dim v As Variant
    For Each v In order
        Dim br2 As String: br2 = CStr(v)
        rows.Add Array(br2, LookupValue(TBL_PRIJEMNICA, COL_PRJ_BROJ, br2, COL_PRJ_DATUM), _
                       CStr(vrD(br2)), CStr(soD(br2)), CLng(gSum(br2)), CLng(sCnt(br2)))
    Next v
    GetPrijemniceSaOsirocenimPaletama = StornoRowsTo2D(rows, 6)
    Exit Function
EH:
    LogErr "modDokumenta.GetPrijemniceSaOsirocenimPaletama"
    GetPrijemniceSaOsirocenimPaletama = Empty
End Function

' Aktivne (ne-stornirane) prijemnice = ciljevi za pallet re-point. Ukljucuje i
' paletizovane i NEpaletizovane (motor ReassignPaleteToPrijemnica_TX podrzava cilj
' bez svezih stavki: STEP1 tada ne undo-uje nista). Kolone 1..7:
'   BrojPrijemnice|Datum|Vrsta|Sorta|Gajbica(KolAmb)|Zbirna|Paletizovana(Da/Ne)
Public Function GetAktivnePrijemnice() As Variant
    On Error GoTo EH
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function

    ' BrojPrijemnice koje imaju bar jednu AKTIVNU paleta-stavku (-> Paletizovana=Da).
    Dim palBr As Object: Set palBr = CreateObject("Scripting.Dictionary"): palBr.CompareMode = vbTextCompare
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(ps) Then
        Dim xBr As Long, xSt As Long, x As Long
        xBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        xSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        For x = 1 To UBound(ps, 1)
            Dim sx As Boolean: sx = False
            If xSt > 0 Then sx = (UCase$(Trim$(NzToText(ps(x, xSt)))) = "DA")
            If Not sx Then palBr(Trim$(NzToText(ps(x, xBr)))) = True
        Next x
    End If

    Dim cBr As Long, cDat As Long, cVr As Long, cSo As Long, cAmb As Long, cZbr As Long, cSt As Long
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    cVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    cSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    cAmb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBr = 0 Then Exit Function

    Dim gSum As Object: Set gSum = CreateObject("Scripting.Dictionary"): gSum.CompareMode = vbTextCompare
    Dim datD As Object: Set datD = CreateObject("Scripting.Dictionary"): datD.CompareMode = vbTextCompare
    Dim vrD As Object: Set vrD = CreateObject("Scripting.Dictionary"): vrD.CompareMode = vbTextCompare
    Dim soD As Object: Set soD = CreateObject("Scripting.Dictionary"): soD.CompareMode = vbTextCompare
    Dim zbD As Object: Set zbD = CreateObject("Scripting.Dictionary"): zbD.CompareMode = vbTextCompare
    Dim order As Collection: Set order = New Collection
    Dim i As Long
    For i = 1 To UBound(prj, 1)
        Dim stp As Boolean: stp = False
        If cSt > 0 Then stp = (UCase$(Trim$(NzToText(prj(i, cSt)))) = "DA")
        If Not stp Then
            Dim br As String: br = Trim$(NzToText(prj(i, cBr)))
            If Len(br) > 0 Then
                If Not gSum.Exists(br) Then
                    gSum(br) = 0&: order.Add br
                    datD(br) = prj(i, cDat): vrD(br) = Trim$(NzToText(prj(i, cVr)))
                    soD(br) = Trim$(NzToText(prj(i, cSo))): zbD(br) = Trim$(NzToText(prj(i, cZbr)))
                End If
                If cAmb > 0 Then If IsNumeric(prj(i, cAmb)) Then gSum(br) = CLng(gSum(br)) + CLng(prj(i, cAmb))
            End If
        End If
    Next i

    Dim rows As Collection: Set rows = New Collection
    Dim v As Variant
    For Each v In order
        Dim br2 As String: br2 = CStr(v)
        rows.Add Array(br2, datD(br2), CStr(vrD(br2)), CStr(soD(br2)), CLng(gSum(br2)), _
                       CStr(zbD(br2)), IIf(palBr.Exists(br2), "Da", "Ne"))
    Next v
    GetAktivnePrijemnice = StornoRowsTo2D(rows, 7)
    Exit Function
EH:
    LogErr "modDokumenta.GetAktivnePrijemnice"
    GetAktivnePrijemnice = Empty
End Function

' Bezbedan re-point izgubljenog bloka na ciljnu (aktivnu) otpremnicu:
' menja SAMO OtpremnicaID + BrojZbirne (cuva OtkupID -> uplate/ambalaza ostaju).
' Transakciono. Vraca False ako cilj ne postoji/storniran ili upis padne.
' Faza 7 korak 5: dual-write BrojOtpremnice na blok (denorm poslovni kljuc, stabilan
' kroz re-verziju otpremnice). otpID prazan -> ocisti (unbind). Guarded na kolonu
' (schema-drift safe) -> non-breaking dok se citanje ne prebaci sa OtpremnicaID.
Public Sub SetOtkupBrojOtpremnice(ByVal rowIndex As Long, ByVal otpID As String)
    On Error Resume Next
    If GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_OTPREMNICE) = 0 Then Exit Sub
    Dim broj As String: broj = ""
    If Len(Trim$(otpID)) > 0 Then _
        broj = Trim$(CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ)))
    UpdateCell TBL_OTKUP, rowIndex, COL_OTK_BROJ_OTPREMNICE, broj
End Sub

Public Function ReassignOtkupToOtpremnica_TX(ByVal otkupID As String, _
                                             ByVal targetOtpID As String) As Boolean
    Const SRC As String = "modDokumenta.ReassignOtkupToOtpremnica_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    If Len(Trim$(otkupID)) = 0 Or Len(Trim$(targetOtpID)) = 0 Then Exit Function

    ' cilj mora postojati (BrojOtpremnice nikad nije blank) i biti aktivan.
    ' NB: Stornirano JE blank za aktivnu otpremnicu -> ne sme se koristiti
    '     za proveru postojanja (LookupValue bi vratio Empty = lazno "ne postoji").
    Dim tBroj As Variant
    tBroj = LookupValue(TBL_OTPREMNICA, COL_OTP_ID, targetOtpID, COL_OTP_BROJ)
    If IsEmpty(tBroj) Then Exit Function                        ' cilj stvarno ne postoji
    Dim tStor As String
    tStor = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, targetOtpID, COL_STORNIRANO))
    If UCase$(Trim$(tStor)) = "DA" Then Exit Function           ' cilj storniran

    Dim tZbr As String
    tZbr = NzToText(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, targetOtpID, COL_OTP_BROJ_ZBIRNE))

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP

    Dim rows As Collection: Set rows = FindRows(TBL_OTKUP, COL_OTK_ID, otkupID)
    If rows.count = 0 Then Err.Raise vbObjectError + 2600, SRC, "Otkup nije nadjen: " & otkupID

    Dim hasZbr As Boolean: hasZbr = (GetColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE) > 0)
    Dim k As Long
    For k = 1 To rows.count
        RequireUpdateCell TBL_OTKUP, rows(k), COL_OTK_OTPREMNICA_ID, targetOtpID, SRC
        If hasZbr Then RequireUpdateCell TBL_OTKUP, rows(k), COL_OTK_BROJ_ZBIRNE, tZbr, SRC
        SetOtkupBrojOtpremnice rows(k), targetOtpID
    Next k

    tx.CommitTx
    Set tx = Nothing
    ReassignOtkupToOtpremnica_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    ReassignOtkupToOtpremnica_TX = False
End Function

' Re-point prijemnice na drugu (aktivnu) zbirnu po BrojZbirne. Koristi se posle
' storna otpremnice+zbirne (van autohladnjace): operater napravi novu otpremnicu+
' zbirnu pa "osirocenu" prijemnicu prevezuje na nju. Handle = BrojPrijemnice
' (citljiv, eksteran kad nije autohladnjaca); stvarna veza lanca = BrojZbirne.
' PrijemnicaID se NE dira -> faktura-stavke i paleta-stavke ostaju ispravne.
Public Function ReassignPrijemnicaToZbirna_TX(ByVal brPrijemnice As String, _
                                              ByVal targetBrZbirne As String) As Boolean
    Const SRC As String = "modDokumenta.ReassignPrijemnicaToZbirna_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    brPrijemnice = Trim$(brPrijemnice)
    targetBrZbirne = Trim$(targetBrZbirne)
    If Len(brPrijemnice) = 0 Or Len(targetBrZbirne) = 0 Then Exit Function

    ' Cilj mora biti AKTIVNA zbirna (postoji + nije stornirana). ZbirnaID je uvek
    ' popunjen -> dokaz postojanja; Stornirano je blank za aktivnu pa se NE sme
    ' koristiti kao dokaz postojanja.
    Dim tId As Variant
    tId = LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, COL_ZBR_ID)
    If IsEmpty(tId) Then Exit Function                          ' zbirna ne postoji
    Dim tStor As String
    tStor = NzToText(LookupValue(TBL_ZBIRNA, COL_ZBR_BROJ, targetBrZbirne, COL_STORNIRANO))
    If UCase$(Trim$(tStor)) = "DA" Then Exit Function           ' cilj-zbirna stornirana

    Dim cBrPrij As Long, cBrZbr As Long, cStorno As Long
    cBrPrij = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If cBrPrij = 0 Or cBrZbr = 0 Then Exit Function

    Dim data As Variant: data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then Exit Function

    ' Aktivni redovi prijemnice za dati BrojPrijemnice (Klasa I i II dele broj).
    Dim targetRows As Collection: Set targetRows = New Collection
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, cBrPrij))) = brPrijemnice Then
            If cStorno = 0 Or UCase$(Trim$(CStr(data(i, cStorno)))) <> "DA" Then
                targetRows.Add i
            End If
        End If
    Next i
    If targetRows.count = 0 Then Exit Function                  ' nema aktivne prijemnice

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    Dim k As Long
    For k = 1 To targetRows.count
        RequireUpdateCell TBL_PRIJEMNICA, targetRows(k), COL_PRJ_BROJ_ZBIRNE, targetBrZbirne, SRC
    Next k

    ' Sledljivost: paletne stavke te prijemnice moraju dobiti NOVU BrojZbirne, inace
    ' ostaju sa mrtvom zbirnom (paleta -> zbirna -> kooperanti pukne). Menja se samo
    ' veza (BrojZbirne); roba/pripadnost prijemnici ostaje.
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If Not IsEmpty(ps) Then
        Dim pBr As Long, pZb As Long, pSt As Long
        pBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        pZb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
        pSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        If pBr > 0 And pZb > 0 Then
            Dim r2 As Long
            For r2 = 1 To UBound(ps, 1)
                If Trim$(CStr(ps(r2, pBr))) = brPrijemnice _
                   And (pSt = 0 Or UCase$(Trim$(CStr(ps(r2, pSt)))) <> "DA") Then
                    RequireUpdateCell TBL_PALETA_STAVKA, r2, COL_PALS_BROJ_ZBIRNE, targetBrZbirne, SRC
                End If
            Next r2
        End If
    End If

    tx.CommitTx
    Set tx = Nothing
    ReassignPrijemnicaToZbirna_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    ReassignPrijemnicaToZbirna_TX = False
End Function
