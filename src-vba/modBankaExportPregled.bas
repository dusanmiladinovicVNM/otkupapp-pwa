Attribute VB_Name = "modBankaExportPregled"
Option Explicit

' ============================================================
' modIsplatePregled v6.18+
'
' Read-only helper sloj za pregled otvorenih otkup blokova
' (per-kooperant aggregation + already-paid awareness + TR check)
' + izlazi za pripremu isplata:
'   - GenerisiNalogeCSV: CSV naloga za prenos (uvoz u e-banking);
'     iznos po bloku = clsBlokIsplata.IsplatitiIznos (operater unos).
'   - PrintIsplataSpecifikacija: specifikacija isplata (PDF/stampa,
'     ISPLATA_SPEC_PRINT_MODE), isti blokovi i iznosi kao CSV.
'
' Bazira se na postojecim helperima:
'   - GetOpenOtkupi (extended 7-col shape)
'   - GetKooperantNaziv (modBankaMapiranje)
'   - LookupValue za TekuciRacun
'
' NE pise nista u tblNovac: isplata se knjizi tek kroz uvoz izvoda
' (modBankaImport/modBankaMapiranje, auto-map poziv na broj = broj bloka).
' ============================================================

'======================================================================
' BuildBlokIsplataList
'
' Vraca Collection of BlokIsplata, sve sistemski-otvoreni blokovi
' enrich-ovani sa kooperant TR + KooperantNaziv.
'
' Filteri:
'   - datumOd / datumDo: opciono date range (#1/1/1900# = no filter)
'   - stanicaIDFilter: opciono stanica filter ("" = sve stanice)
'
' Performance: jedan poziv GetOpenOtkupi() (single-pass),
' jedan poziv BuildIsplataDictByOtkup() (sadrzi se u GetOpenOtkupi),
' N lookup-a za TR i KooperantNaziv (cached preko Dictionary-ja).
'======================================================================
Public Function BuildBlokIsplataList( _
    Optional ByVal datumOd As Date, _
    Optional ByVal datumDo As Date, _
    Optional ByVal stanicaIDFilter As String = "" _
) As Collection
    
    Dim result As New Collection
    Dim openOtkupi As Variant
    
    openOtkupi = GetOpenOtkupi("")
    
    If IsEmpty(openOtkupi) Then
        Set BuildBlokIsplataList = result
        Exit Function
    End If
    
    Dim trCache As Object
    Dim nazivCache As Object
    Dim avansCache As Object
    Set trCache = BuildKooperantTekuciRacunCache()
    Set nazivCache = CreateObject("Scripting.Dictionary")
    Set avansCache = BuildKooperantUnallocatedAvansDict()
    
    Dim i As Long
    For i = 1 To UBound(openOtkupi, 1)
        Dim brojDok As String
        Dim otkupID As String
        Dim ukupan As Double
        Dim isplaceno As Double
        Dim otvoren As Double
        Dim datumVal As Date
        Dim stanicaID As String
        
        brojDok = CStr(openOtkupi(i, 1))
        otkupID = CStr(openOtkupi(i, 2))
        ukupan = CDbl(openOtkupi(i, 3))
        isplaceno = CDbl(openOtkupi(i, 4))
        otvoren = CDbl(openOtkupi(i, 5))
        
        If IsDate(openOtkupi(i, 6)) Then
            datumVal = CDate(openOtkupi(i, 6))
        Else
            On Error Resume Next
            datumVal = CDate(CStr(openOtkupi(i, 6)))
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextRow
            End If
            On Error GoTo 0
        End If
        
        stanicaID = CStr(openOtkupi(i, 7))
        
        If stanicaIDFilter <> "" And stanicaID <> stanicaIDFilter Then GoTo NextRow
        If datumOd > #1/1/1900# And datumVal < datumOd Then GoTo NextRow
        If datumDo > #1/1/1900# And datumVal > datumDo Then GoTo NextRow
        
        Dim kooperantID As String
        kooperantID = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_KOOPERANT))
        If LenB(Trim$(kooperantID)) = 0 Then GoTo NextRow
        
        Dim kooperantNaziv As String
        If nazivCache.Exists(kooperantID) Then
            kooperantNaziv = nazivCache(kooperantID)
        Else
            kooperantNaziv = GetKooperantNaziv(kooperantID)
            nazivCache.Add kooperantID, kooperantNaziv
        End If
        
        Dim tr As String
        If trCache.Exists(kooperantID) Then
            tr = trCache(kooperantID)
        Else
            tr = ""
        End If
        
        ' --- promenjeno: class instance umesto UDT ---
        Dim blk As clsBlokIsplata
        Set blk = New clsBlokIsplata
        blk.otkupID = otkupID
        blk.datum = datumVal
        blk.kooperantID = kooperantID
        blk.kooperantNaziv = kooperantNaziv
        blk.stanicaID = stanicaID
        blk.TekuciRacun = tr
        blk.brojDokumenta = brojDok
        blk.UkupanIznos = ukupan
        blk.VecIsplaceno = isplaceno
        blk.OtvorenIznos = otvoren
        blk.HasTekuciRacun = (LenB(Trim$(tr)) > 0)
        If avansCache.Exists(kooperantID) Then
            blk.KooperantAvansSaldo = CDbl(avansCache(kooperantID))
        Else
            blk.KooperantAvansSaldo = 0
        End If
        
        result.Add blk
        
NextRow:
    Next i
    
    Set BuildBlokIsplataList = result
End Function
'======================================================================
' BuildKooperantTekuciRacunCache
' Single-pass cache KooperantID -> TekuciRacun
'======================================================================
Private Function BuildKooperantTekuciRacunCache() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim data As Variant
    data = GetTableData(TBL_KOOPERANTI)
    If IsEmpty(data) Then
        Set BuildKooperantTekuciRacunCache = dict
        Exit Function
    End If
    
    Const SRC As String = "BuildKooperantTekuciRacunCache"
    Dim colID As Long, colTR As Long
    colID = RequireColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID, SRC)
    colTR = RequireColumnIndex(TBL_KOOPERANTI, COL_KOOP_TEKUCI_RACUN, SRC)
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim kID As String
        kID = Trim$(CStr(data(i, colID)))
        If LenB(kID) > 0 Then
            If Not dict.Exists(kID) Then
                dict.Add kID, Trim$(CStr(data(i, colTR)))
            End If
        End If
    Next i
    
    Set BuildKooperantTekuciRacunCache = dict
End Function

'======================================================================
' SummarizeBlokList - status bar string
'======================================================================
Public Function SummarizeBlokList(ByVal blokovi As Collection) As String
    If blokovi Is Nothing Then
        SummarizeBlokList = "Nema otvorenih blokova."
        Exit Function
    End If
    If blokovi.count = 0 Then
        SummarizeBlokList = "Nema otvorenih blokova."
        Exit Function
    End If
    
    Dim totalOpen As Double
    Dim missingTR As Long
    Dim kooperantSet As Object
    Set kooperantSet = CreateObject("Scripting.Dictionary")
    
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v               ' Set jer je v object reference
        totalOpen = totalOpen + blk.OtvorenIznos
        If Not blk.HasTekuciRacun Then missingTR = missingTR + 1
        If Not kooperantSet.Exists(blk.kooperantID) Then
            kooperantSet.Add blk.kooperantID, True
        End If
    Next v
    
    Dim baseMsg As String
    baseMsg = blokovi.count & " blokova | " & _
              kooperantSet.count & " kooperanata | " & _
              "Otvoreno: " & Format$(totalOpen, "#,##0.00") & " RSD"
    
    If missingTR > 0 Then
        baseMsg = baseMsg & " | Bez TR: " & missingTR
    End If
    
    SummarizeBlokList = baseMsg
End Function

'======================================================================
' PrintIsplataSpecifikacija - specifikacija isplata po blokovima (PDF/
' stampa po ISPLATA_SPEC_PRINT_MODE, default PDF -> otvori se).
' Ocekuje kolekciju blokova SA postavljenim IsplatitiIznos
' (CollectIsplataBlokovi u frmBankaExportPregled). Render u modPrint
' (EnsureIsplataSpecSablon/FillIsplataSpecSablon, house-style).
'======================================================================
Public Sub PrintIsplataSpecifikacija(ByVal blokovi As Collection, _
                                     Optional ByVal platilacRacun As String = "")
    On Error GoTo EH
    If blokovi Is Nothing Then Exit Sub
    If blokovi.count = 0 Then Exit Sub

    platilacRacun = NormalizujRacun(platilacRacun)
    If LenB(platilacRacun) = 0 Then platilacRacun = NormalizujRacun(DocConfigOr("SELLER_ACCOUNT", ""))

    Dim n As Long: n = blokovi.count
    Dim spec() As Variant: ReDim spec(1 To n, 1 To 9)
    Dim sumUkupan As Double, sumIsplaceno As Double
    Dim sumOtvoreno As Double, sumIsplatiti As Double

    Dim i As Long
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        i = i + 1
        spec(i, 1) = i
        spec(i, 2) = Format$(blk.datum, "d.m.yyyy")
        spec(i, 3) = blk.brojDokumenta
        spec(i, 4) = blk.kooperantNaziv
        spec(i, 5) = blk.TekuciRacun
        spec(i, 6) = blk.UkupanIznos
        spec(i, 7) = blk.VecIsplaceno
        spec(i, 8) = blk.OtvorenIznos
        spec(i, 9) = blk.IsplatitiIznos
        sumUkupan = sumUkupan + blk.UkupanIznos
        sumIsplaceno = sumIsplaceno + blk.VecIsplaceno
        sumOtvoreno = sumOtvoreno + blk.OtvorenIznos
        sumIsplatiti = sumIsplatiti + blk.IsplatitiIznos
    Next v

    Dim subtitle As String
    subtitle = "Datum: " & Format$(Date, "d.m.yyyy") & _
               "     Platilac: " & DocConfigOr("SELLER_NAME", "") & _
               " (" & platilacRacun & ")"

    Dim ws As Worksheet
    Set ws = FillIsplataSpecSablon(spec, n, subtitle, _
                                   sumUkupan, sumIsplaceno, sumOtvoreno, sumIsplatiti)
    If ws Is Nothing Then Exit Sub

    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_ISPLATA_SPEC_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            Dim pdfPath As String
            pdfPath = EnsureDocFolder(PDF_DIR_SPECIFIKACIJE) & "\Specifikacija_isplata_" & _
                      Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
            DocExportPdf ws, pdfPath, True
        ' OFF -> bez izlaza
    End Select
    Exit Sub
EH:
    LogErr "modBankaExportPregled.PrintIsplataSpecifikacija"
    MsgBox "Gre" & ChrW(353) & "ka pri izradi specifikacije: " & Err.description, vbCritical, APP_NAME
End Sub

'======================================================================
' GenerisiNalogeCSV - CSV naloga za prenos za uvoz u e-banking.
'
' Jedan red = jedan nalog za prenos po otkupnom bloku:
'   - platilac: SELLER_NAME / SELLER_ACCOUNT (config, grupa "Prodavac (firma)")
'   - primalac: kooperant (naziv + TekuciRacun iz tblKooperanti)
'   - iznos:    blk.IsplatitiIznos (postavlja forma: operater unos ili otvoreno)
'   - poziv na broj (odobrenje) = broj otkupnog bloka -> jaki kljuc za
'     auto-map pri kasnijem uvozu izvoda (frmBankaImport)
'
' Ocekuje blokove SA tekucim racunom (filtrira ih forma). Blokove bez TR
' preskace i tiho broji (defenzivno). Vraca punu putanju CSV fajla,
' "" ako nema nijednog reda ili upis ne uspe.
'
' Format: ";" separator, UTF-8 (BOM - vidi WriteAllTextUtf8), decimalna
' tacka u iznosu (deterministicki, nezavisno od Windows locale-a).
' Kolone drzati stabilnim: e-banking uvozi se mapiraju po pozicijama.
'======================================================================
Public Function GenerisiNalogeCSV(ByVal blokovi As Collection, _
                                  Optional ByVal platilacRacun As String = "") As String
    On Error GoTo EH

    GenerisiNalogeCSV = ""
    If blokovi Is Nothing Then Exit Function
    If blokovi.count = 0 Then Exit Function

    ' Platilac racun: prosledjen iz forme (combo "Sa racuna"); prazno ->
    ' SELLER_ACCOUNT (backward-compatible kada racuna ima samo jedan).
    Dim platilacNaziv As String
    platilacNaziv = DocConfigOr("SELLER_NAME", "")
    platilacRacun = NormalizujRacun(platilacRacun)
    If LenB(platilacRacun) = 0 Then platilacRacun = NormalizujRacun(DocConfigOr("SELLER_ACCOUNT", ""))
    If LenB(platilacRacun) = 0 Then Exit Function   ' forma validira i javlja poruku

    Dim sifra As String, svrhaBase As String
    sifra = DocConfigOr(CFG_BANKA_NALOG_SIFRA, BANKA_NALOG_SIFRA_DEFAULT)
    svrhaBase = DocConfigOr(CFG_BANKA_NALOG_SVRHA, BANKA_NALOG_SVRHA_DEFAULT)

    Dim datumValute As String
    datumValute = Format$(Date, "dd.mm.yyyy")

    Dim s As String
    s = "RacunPlatioca;NazivPlatioca;RacunPrimaoca;NazivPrimaoca;Iznos;Valuta;" & _
        "SifraPlacanja;Model;PozivNaBroj;SvrhaPlacanja;DatumValute" & vbCrLf

    Dim rows As Long
    Dim blk As clsBlokIsplata
    Dim v As Variant
    For Each v In blokovi
        Set blk = v
        If blk.HasTekuciRacun And blk.IsplatitiIznos > 0 Then
            s = s & CsvField(platilacRacun) & ";" & _
                    CsvField(platilacNaziv) & ";" & _
                    CsvField(NormalizujRacun(blk.TekuciRacun)) & ";" & _
                    CsvField(blk.kooperantNaziv) & ";" & _
                    CsvIznos(blk.IsplatitiIznos) & ";" & _
                    "RSD;" & _
                    CsvField(sifra) & ";" & _
                    ";" & _
                    CsvField(blk.brojDokumenta) & ";" & _
                    CsvField(svrhaBase & " " & blk.brojDokumenta) & ";" & _
                    datumValute & vbCrLf
            rows = rows + 1
        End If
    Next v

    If rows = 0 Then Exit Function

    Dim csvPath As String
    csvPath = EnsureDocFolder(CSV_DIR_BANKA_NALOZI) & "\Nalozi_za_prenos_" & _
              Format$(Now, "yyyymmdd_hhnnss") & ".csv"
    WriteAllTextUtf8 csvPath, s

    GenerisiNalogeCSV = csvPath
    Exit Function

EH:
    LogErr "modBankaExportPregled.GenerisiNalogeCSV"
    GenerisiNalogeCSV = ""
End Function

' Iznos za CSV: uvek decimalna TACKA, bez hiljada separatora, 2 decimale.
' Format$ "0.00" na sr locale daje zarez -> normalizuj deterministicki.
Private Function CsvIznos(ByVal amt As Double) As String
    CsvIznos = Replace(Format$(amt, "0.00"), ",", ".")
End Function

' CSV polje: trim + quote ako sadrzi separator/navodnik/novi red.
Private Function CsvField(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If InStr(t, ";") > 0 Or InStr(t, """") > 0 Or _
       InStr(t, vbCr) > 0 Or InStr(t, vbLf) > 0 Then
        t = """" & Replace(t, """", """""") & """"
    End If
    CsvField = t
End Function

' Tekuci racun za nalog: skini razmake (format "160-xxxx-xx" ili 18 cifara
' ostaje kako je unet u maticne podatke / config).
Private Function NormalizujRacun(ByVal racun As String) As String
    NormalizujRacun = Replace(Trim$(racun), " ", "")
End Function

'======================================================================
' BankaNazivZaRacun - ime banke iz vodeceg NBS koda racuna (prve 3
' cifre / deo pre prve crtice). Za prikaz u combo-u "Sa racuna"
' (frmBankaExportPregled) kada firma ima racune u vise banaka.
' Nepoznat kod -> "" (prikaze se samo racun).
'======================================================================
Public Function BankaNazivZaRacun(ByVal racun As String) As String
    Dim r As String
    r = Replace(Trim$(racun), " ", "")

    Dim kod As String
    If InStr(r, "-") > 0 Then
        kod = Left$(r, InStr(r, "-") - 1)
    ElseIf Len(r) >= 3 Then
        kod = Left$(r, 3)
    End If

    Select Case kod
        Case "105": BankaNazivZaRacun = "AIK"
        Case "155": BankaNazivZaRacun = "Halkbank"
        Case "160": BankaNazivZaRacun = "Banca Intesa"
        Case "165": BankaNazivZaRacun = "Addiko"
        Case "170": BankaNazivZaRacun = "UniCredit"
        Case "200": BankaNazivZaRacun = "Po" & ChrW(353) & "tanska " & ChrW(353) & "tedionica"
        Case "205": BankaNazivZaRacun = "NLB Komercijalna"
        Case "220": BankaNazivZaRacun = "ProCredit"
        Case "265": BankaNazivZaRacun = "Raiffeisen"
        Case "275", "325": BankaNazivZaRacun = "OTP"
        Case "340": BankaNazivZaRacun = "Erste"
        Case Else: BankaNazivZaRacun = ""
    End Select
End Function

