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

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA

    SaveOtkup_TX = SaveOtkup(datum, kooperantID, stanicaID, vrstaVoca, _
                              sortaVoca, kolicina, cena, tipAmb, kolAmb, _
                              vozacID, brDok, novac, primalac, klasa, _
                              parcelaID, brojZbirne)

    If SaveOtkup_TX = "" Then
        Err.Raise vbObjectError + 1801, "SaveOtkup_TX", _
                  "SaveOtkup fehlgeschlagen"
    End If

    tx.CommitTx
    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr "SaveOtkup_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtkup_TX = ""

    PrintOtkupTxFailure "SaveOtkup_TX", errSrc, errNum, errDesc
End Function

Public Function SaveOtkupMulti_TX(ByVal datum As Date, _
                                   ByVal kooperantID As String, _
                                   ByVal stanicaID As String, _
                                   ByVal vrstaVoca As String, _
                                   ByVal sortaVoca As String, _
                                   ByVal kolicinaI As Double, _
                                   ByVal cenaI As Double, _
                                   ByVal tipAmb As String, _
                                   ByVal kolAmb As Long, _
                                   ByVal vozacID As String, _
                                   ByVal brDok As String, _
                                   ByVal novac As Double, _
                                   ByVal primalac As String, _
                                   ByVal parcelaID As String, _
                                   ByVal brojZbirne As String, _
                                   Optional ByVal hasKlasaII As Boolean = False, _
                                   Optional ByVal kolicinaII As Double = 0, _
                                   Optional ByVal cenaII As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    If Trim$(kooperantID) = "" Then
        Err.Raise vbObjectError + 1810, "SaveOtkupMulti_TX", _
                  "KooperantID je obavezan."
    End If

    If Trim$(stanicaID) = "" Then
        Err.Raise vbObjectError + 1811, "SaveOtkupMulti_TX", _
                  "StanicaID je obavezan."
    End If

    If kolicinaI <= 0 Or cenaI <= 0 Then
        Err.Raise vbObjectError + 1812, "SaveOtkupMulti_TX", _
                  "Kolicina i cena za Klasu I moraju biti vece od nule."
    End If

    If hasKlasaII Then
        If kolicinaII <= 0 Or cenaII <= 0 Then
            Err.Raise vbObjectError + 1813, "SaveOtkupMulti_TX", _
                      "Kolicina i cena za Klasu II moraju biti vece od nule."
        End If
    End If

    If kolAmb < 0 Then
        Err.Raise vbObjectError + 1814, "SaveOtkupMulti_TX", _
                  "Kolicina ambalaze ne sme biti negativna."
    End If

    If novac < 0 Then
        Err.Raise vbObjectError + 1815, "SaveOtkupMulti_TX", _
                  "Iznos novca ne sme biti negativan."
    End If

    If kolAmb > 0 And Trim$(tipAmb) = "" Then
        Err.Raise vbObjectError + 1816, "SaveOtkupMulti_TX", _
                  "Tip ambalaze je obavezan kada postoji ambalaza."
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
    tx.AddTableSnapshot TBL_NOVAC

    Dim resultI As String
    resultI = SaveOtkup( _
        datum:=datum, _
        kooperantID:=kooperantID, _
        stanicaID:=stanicaID, _
        vrstaVoca:=vrstaVoca, _
        sortaVoca:=sortaVoca, _
        kolicina:=kolicinaI, _
        cena:=cenaI, _
        tipAmb:=tipAmb, _
        kolAmb:=kolAmb, _
        vozacID:=vozacID, _
        brDok:=brDok, _
        novac:=novac, _
        primalac:=primalac, _
        klasa:=KLASA_I, _
        parcelaID:=parcelaID, _
        brojZbirne:=brojZbirne)

    If resultI = "" Then
        Err.Raise vbObjectError + 1817, "SaveOtkupMulti_TX", _
                  "SaveOtkup Klasa I fehlgeschlagen"
    End If

    Dim resultII As String

    If hasKlasaII Then
        resultII = SaveOtkup( _
            datum:=datum, _
            kooperantID:=kooperantID, _
            stanicaID:=stanicaID, _
            vrstaVoca:=vrstaVoca, _
            sortaVoca:=sortaVoca, _
            kolicina:=kolicinaII, _
            cena:=cenaII, _
            tipAmb:=tipAmb, _
            kolAmb:=0, _
            vozacID:=vozacID, _
            brDok:=brDok, _
            novac:=0, _
            primalac:=primalac, _
            klasa:=KLASA_II, _
            parcelaID:=parcelaID, _
            brojZbirne:=brojZbirne)

        If resultII = "" Then
            Err.Raise vbObjectError + 1818, "SaveOtkupMulti_TX", _
                      "SaveOtkup Klasa II fehlgeschlagen"
        End If
    End If

    If novac > 0 Then
        Dim koopNaziv As String
        koopNaziv = GetKooperantNazivForNovac(kooperantID)

        Dim novacID As String
        novacID = SaveNovac( _
            brojDok:=brDok, _
            datum:=datum, _
            partner:=koopNaziv, _
            partnerID:=kooperantID, _
            entitetTip:="Kooperant", _
            omID:=stanicaID, _
            kooperantID:=kooperantID, _
            fakturaID:="", _
            vrstaVoca:=vrstaVoca, _
            tip:=NOV_KES_OTKUPAC_KOOP, _
            uplata:=0, _
            isplata:=novac, _
            napomena:=primalac, _
            otkupID:=resultI)

        If novacID = "" Then
            Err.Raise vbObjectError + 1819, "SaveOtkupMulti_TX", _
                      "SaveNovac fehlgeschlagen"
        End If
    End If

    ApplyAvansToOtkup kooperantID, resultI

    If hasKlasaII Then
        ApplyAvansToOtkup kooperantID, resultII
    End If

    tx.CommitTx
    Set tx = Nothing

    If hasKlasaII Then
        SaveOtkupMulti_TX = resultI & " + " & resultII
    Else
        SaveOtkupMulti_TX = resultI
    End If

    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr "SaveOtkupMulti_TX"

    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    SaveOtkupMulti_TX = ""

    PrintOtkupTxFailure "SaveOtkupMulti_TX", errSrc, errNum, errDesc
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
    On Error GoTo EH

    If Trim$(kooperantID) = "" Then
        Err.Raise vbObjectError + 1820, "SaveOtkup", _
                  "Kooperant mora biti izabran."
    End If

    If Trim$(stanicaID) = "" Then
        Err.Raise vbObjectError + 1821, "SaveOtkup", _
                  "Stanica mora biti izabrana."
    End If

    If Trim$(vrstaVoca) = "" Then
        Err.Raise vbObjectError + 1822, "SaveOtkup", _
                  "Vrsta voca je obavezna."
    End If

    If kolicina <= 0 Then
        Err.Raise vbObjectError + 1823, "SaveOtkup", _
                  "Kolicina mora biti veca od nule."
    End If

    If cena <= 0 Then
        Err.Raise vbObjectError + 1824, "SaveOtkup", _
                  "Cena mora biti veca od nule."
    End If

    If kolAmb < 0 Then
        Err.Raise vbObjectError + 1825, "SaveOtkup", _
                  "Kolicina ambalaze ne sme biti negativna."
    End If

    If novac < 0 Then
        Err.Raise vbObjectError + 1826, "SaveOtkup", _
                  "Novac ne sme biti negativan."
    End If

    If kolAmb > 0 And Trim$(tipAmb) = "" Then
        Err.Raise vbObjectError + 1827, "SaveOtkup", _
                  "Tip ambalaze je obavezan kada postoji ambalaza."
    End If
    
    Call RequireValidOtkupClass(klasa, "SaveOtkup")

    RequireColumns TBL_OTKUP, "SaveOtkup", _
                   COL_OTK_ID, _
                   COL_OTK_DATUM, _
                   COL_OTK_KOOPERANT, _
                   COL_OTK_STANICA, _
                   COL_OTK_KULTURA, _
                   COL_OTK_VRSTA, _
                   COL_OTK_SORTA, _
                   COL_OTK_KOLICINA, _
                   COL_OTK_CENA, _
                   COL_OTK_TIP_AMB, _
                   COL_OTK_KOL_AMB, _
                   COL_OTK_VOZAC, _
                   COL_OTK_BR_DOK, _
                   COL_OTK_NOVAC, _
                   COL_OTK_PRIMALAC, _
                   COL_OTK_KLASA, _
                   COL_OTK_STORNIRANO, _
                   COL_OTK_BROJ_ZBIRNE, _
                   COL_OTK_ISPLACENO, _
                   COL_OTK_DATUM_ISPLATE, _
                   COL_OTK_OTPREMNICA_ID, _
                   COL_OTK_PARCELA

    Dim newID As String
    newID = GetNextID(TBL_OTKUP, COL_OTK_ID, "OTK-")

    If newID = "" Then
        Err.Raise vbObjectError + 1828, "SaveOtkup", _
                  "GetNextID nije vratio OtkupID."
    End If

    Dim kulturaID As String
    kulturaID = CStr(LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, "KulturaID"))

    If kulturaID = "" Then
        kulturaID = vrstaVoca & "-" & sortaVoca
    End If

    Dim rowData As Variant
    rowData = Array( _
        newID, _
        datum, _
        kooperantID, _
        stanicaID, _
        kulturaID, _
        vrstaVoca, _
        sortaVoca, _
        kolicina, _
        cena, _
        tipAmb, _
        kolAmb, _
        vozacID, _
        brDok, _
        novac, _
        primalac, _
        klasa, _
        "", _
        brojZbirne, _
        "", _
        Empty, _
        "", _
        parcelaID _
    )

    If AppendRow(TBL_OTKUP, rowData) <= 0 Then
        Err.Raise vbObjectError + 1829, "SaveOtkup", _
                  "AppendRow fehlgeschlagen für tblOtkup."
    End If

    If kolAmb > 0 Then
        TrackAmbalaza datum, tipAmb, kolAmb, "Izlaz", _
                      kooperantID, "Kooperant", vozacID, _
                      newID, DOK_TIP_OTKUP
    End If

    SaveOtkup = newID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    On Error Resume Next
    LogErr "SaveOtkup"
    On Error GoTo 0

    Err.Raise errNum, "SaveOtkup", _
              "Source=" & errSrc & " | " & errDesc
End Function

Private Function GetKooperantNazivForNovac(ByVal kooperantID As String) As String
    On Error GoTo EH

    Dim ime As String
    Dim prezime As String

    ime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, kooperantID, "Ime")))
    prezime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, kooperantID, "Prezime")))

    GetKooperantNazivForNovac = Trim$(ime & " " & prezime)

    If GetKooperantNazivForNovac = "" Then
        GetKooperantNazivForNovac = kooperantID
    End If

    Exit Function

EH:
    LogErr "GetKooperantNazivForNovac"
    GetKooperantNazivForNovac = kooperantID
End Function

Public Function GetOtkupByStation(ByVal stanicaID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByStation = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByStation = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, _
            "modOtkup.GetOtkupByStation"), "=", stanicaID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                "modOtkup.GetOtkupByStation"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtkupByStation = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modOtkup.GetOtkupByStation"
    GetOtkupByStation = Empty
End Function

Public Function GetOtkupByKooperant(ByVal kooperantID As String, _
                                    Optional ByVal datumOd As Date = 0, _
                                    Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByKooperant = Empty
        Exit Function
    End If
    
    data = ExcludeStornirano(data, TBL_OTKUP)

    If IsEmpty(data) Then
        GetOtkupByKooperant = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, _
            "modOtkup.GetOtkupByKooperant"), "=", kooperantID
    filters.Add fp

    If datumOd > 0 And datumDo > 0 Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, _
                "modOtkup.GetOtkupByKooperant"), "BETWEEN", datumOd, datumDo
        filters.Add fp
    End If

    GetOtkupByKooperant = FilterArray(data, filters)
    Exit Function

EH:
    LogErr "modOtkup.GetOtkupByKooperant"
    GetOtkupByKooperant = Empty
End Function
Public Function GetSaldoByStation(ByVal stanicaID As String, _
                                  Optional ByVal datumOd As Date = 0, _
                                  Optional ByVal datumDo As Date = 0) As Variant
    On Error GoTo EH

    Dim otkupData As Variant
    otkupData = GetOtkupByStation(stanicaID, datumOd, datumDo)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(otkupData) Then
        Dim colKoop As Long
        Dim colKol As Long
        Dim colNovac As Long
        Dim colAmb As Long

        colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, _
                                     "modOtkup.GetSaldoByStation")
        colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, _
                                    "modOtkup.GetSaldoByStation")
        colNovac = RequireColumnIndex(TBL_OTKUP, COL_OTK_NOVAC, _
                                      "modOtkup.GetSaldoByStation")
        colAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB, _
                                    "modOtkup.GetSaldoByStation")

        Dim i As Long
        Dim key As String
        Dim vals As Variant

        For i = 1 To UBound(otkupData, 1)
            key = CStr(otkupData(i, colKoop))

            If key <> "" Then
                If Not dict.Exists(key) Then
                    dict.Add key, Array(0#, 0#, 0#)
                End If

                vals = dict(key)

                If IsNumeric(otkupData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkupData(i, colKol))
                If IsNumeric(otkupData(i, colNovac)) Then vals(1) = vals(1) + CDbl(otkupData(i, colNovac))
                If IsNumeric(otkupData(i, colAmb)) Then vals(2) = vals(2) + CLng(otkupData(i, colAmb))

                dict(key) = vals
            End If
        Next i
    End If

    ' TODO:
    ' Ovaj helper trenutno racuna samo bruto saldo iz tblOtkup.
    ' Banka/Novac/Isporuka korekcije treba rešiti u posebnom report modulu,
    ' ne širiti ovaj core save modul bez jasnog accounting pravila.

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
    Exit Function

EH:
    LogErr "modOtkup.GetSaldoByStation"
    GetSaldoByStation = Empty
End Function


Private Sub PrintOtkupTxFailure(ByVal sourceName As String, _
                                ByVal errSrc As String, _
                                ByVal errNum As Long, _
                                ByVal errDesc As String)
    Debug.Print sourceName & " failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Sub

Private Sub RequireValidOtkupClass(ByVal klasa As String, _
                                   ByVal sourceName As String)

    Select Case Trim$(CStr(klasa))
        Case KLASA_I, KLASA_II
            Exit Sub
    End Select

    Err.Raise vbObjectError + 1830, sourceName, _
              "Neispravna klasa otkupa: " & klasa
End Sub

