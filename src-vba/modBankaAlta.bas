Attribute VB_Name = "modBankaAlta"
Option Explicit

' ============================================================
' modBankaAlta
' Parser za pdftotext izlaz ALTA Banka izvoda (racun 190-...).
'
' Ugovor (isto kao ostali bank-parseri): orkestrator
' modBankaImport.ParseBankaIzvodForImport rutira na 5 funkcija:
'   ExtractIzvodBrojAlta / ExtractIzvodDatumAlta / ExtractIzvodRacunAlta /
'   ExtractIzvodSaldoAlta / ParseBankaIzvodAlta
' Deljeni 4-nivo integrity + 17-kol staging ostaju u orkestratoru.
'
' Output-kolone ParseBankaIzvodAlta (isto kao ostali):
' 1 Datum Izvoda  2 Datum knjizenja  3 Partner  4 Racun
' 5 Zaduzenje(isplata)  6 Odobrenje(uplata)  7 Sifra
' 8 Svrha  9 Poziv na broj  10 Referenca (Podaci za reklamaciju)
'
' Zaglavlje:
'   "IZVOD BR. 110"                                  -> broj izvoda
'   "ZA PROMENU SREDSTAVA NA RA...NU DANA dd.mm.yyyy" -> datum izvoda
'   "<vlasnik> 190-...-.."                            -> racun izvoda
' (naslov "IZVOD BR." razlikuje Altu od Komercijalne/Halk "Izvod broj" i
'  ProCredit "IZVOD NNN"; koristi se u DetectBank.)
'
' STANJE blok (saldo, jedna data-linija: 4 iznosa + 2 cela broja), ispod labela
'   "Prethodno stanje / duguje potrazuje / Novo stanje / zaduzenje odobrenje":
'   "11,346.21 1,654,811.11 2,504,625.00 861,160.10 14 2"
'   = Pocetno  Duguje  Potrazuje  Novo  nZaduzenja  nOdobrenja
'
' PROMENE (transakcije) -- bound na "PROMENE" .. "Ukupno za ra". Po transakciji
' (fiksni redosled), R.B. = standalone ceo broj (1..N):
'   [R.B.]
'   partner            naziv i sediste (1-2 linije)
'   racun              \d{3}-\d+-\d{2}
'   poreklo naloga     banka (1-2 linije, ignorise se)
'   datum knjizenja    dd.mm.yyyy   <-- DatumTransakcije
'   datum prijema      dd.mm.yyyy
'   [zaduzenje]        standalone iznos (SAMO isplata; kod uplate stoji "- : -")
'   "Obr. naknada..."  naknada (ignorise se; iznos na istoj ili sledecoj liniji)
'   sifra-linija:
'      isplata:  "- : - <sifra> <svrha> [ref]"   (odobrenje = 0)
'      uplata:   "<odobrenje> <sifra> <svrha>"    (zaduzenje = 0)
'   svrha (+ [tagovi]/poziv) ... referenca = standalone 15-cifreni broj (kraj bloka)
'
' Broj-format: 1,234.56 (zarez hiljade, tacka decimala) -- isti kao ostali.
' ============================================================


' ---------- 5 ugovornih funkcija ----------

Public Function ExtractIzvodBrojAlta(ByRef lines() As String) As String
    Dim re As Object, m As Object, i As Long, s As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    ' Primarno: "IZVOD BR. 110" (naslov)
    re.pattern = "IZVOD\s+BR\.?\s*(\d+)"
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            ExtractIzvodBrojAlta = Trim$(m.SubMatches(0))
            Exit Function
        End If
    Next i

    ' Fallback: "IZVOD 110" (bez BR.)
    re.pattern = "IZVOD\s+(\d+)"
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            ExtractIzvodBrojAlta = Trim$(m.SubMatches(0))
            Exit Function
        End If
    Next i
End Function

Public Function ExtractIzvodDatumAlta(ByRef lines() As String) As String
    Dim i As Long, s As String, d As String

    ' Primarno: header "ZA PROMENU SREDSTAVA NA RA...NU DANA dd.mm.yyyy"
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If InStr(1, s, "SREDSTAVA NA RA", vbTextCompare) > 0 _
           Or InStr(1, s, "PROMENU SREDSTAVA", vbTextCompare) > 0 Then
            d = FirstDateInStringAlta(s)
            If d <> "" Then
                ExtractIzvodDatumAlta = d
                Exit Function
            End If
        End If
    Next i

    ' Fallback: bilo koja linija sa "DANA" + datum
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If InStr(1, s, "DANA", vbTextCompare) > 0 Then
            d = FirstDateInStringAlta(s)
            If d <> "" Then
                ExtractIzvodDatumAlta = d
                Exit Function
            End If
        End If
    Next i
End Function

Public Function ExtractIzvodRacunAlta(ByRef lines() As String) As String
    Dim i As Long, acc As String
    Dim boundIdx As Long

    ' Racun vlasnika = prvi \d{3}-\d+-\d{2} PRE "STANJE"/"PROMENE" (zaglavlje).
    boundIdx = UBound(lines)
    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), "STANJE", vbTextCompare) > 0 _
           Or InStr(1, lines(i), "PROMENE", vbTextCompare) > 0 Then
            boundIdx = i
            Exit For
        End If
    Next i

    For i = LBound(lines) To boundIdx
        acc = ExtractAccountAnywhereAlta(lines(i))
        If acc <> "" Then
            ExtractIzvodRacunAlta = acc
            Exit Function
        End If
    Next i

    ' Fallback: "Ukupno za ra...un <acct>" (podnozje)
    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), "Ukupno za ra", vbTextCompare) > 0 Then
            acc = ExtractAccountAnywhereAlta(lines(i))
            If acc <> "" Then
                ExtractIzvodRacunAlta = acc
                Exit Function
            End If
        End If
    Next i
End Function

Public Function ExtractIzvodSaldoAlta(ByRef lines() As String) As BankIzvodSaldo
    Dim result As BankIzvodSaldo
    result.parsed = False

    On Error GoTo EH

    Dim i As Long, j As Long, anchorIdx As Long
    anchorIdx = -1

    For i = LBound(lines) To UBound(lines)
        If InStr(1, NormalizeSpacesAlta(lines(i)), "Prethodno stanje", vbTextCompare) > 0 Then
            anchorIdx = i
            Exit For
        End If
    Next i

    If anchorIdx < 0 Then
        ExtractIzvodSaldoAlta = result
        Exit Function
    End If

    ' Labeli su rasporedjeni preko vise linija; data-linija je do ~12 ispod anchora.
    Dim scanEnd As Long
    scanEnd = anchorIdx + 12
    If scanEnd > UBound(lines) Then scanEnd = UBound(lines)

    For j = anchorIdx + 1 To scanEnd
        If TryParseSaldoLineAlta(lines(j), result) Then
            result.parsed = True
            Exit For
        End If
    Next j

    ExtractIzvodSaldoAlta = result
    Exit Function

EH:
    ' Parser nikad ne raise-uje, samo kaze nije parsed.
    result.parsed = False
    ExtractIzvodSaldoAlta = result
End Function

Public Function ParseBankaIzvodAlta(ByVal txt As String) As Variant
    Dim lines() As String
    Dim izvodDatum As String
    Dim i As Long

    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)

    izvodDatum = ExtractIzvodDatumAlta(lines)

    ' Bound na PROMENE sekciju [txStart..txEnd]: sve pre "PROMENE" (zaglavlje,
    ' STANJE) i sve od "Ukupno za ra" (rezime/podnozje) se iskljucuje.
    Dim txStart As Long, txEnd As Long
    txStart = LBound(lines)
    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), "PROMENE", vbTextCompare) > 0 Then
            txStart = i + 1
            Exit For
        End If
    Next i

    txEnd = UBound(lines)
    For i = txStart To UBound(lines)
        If InStr(1, lines(i), "Ukupno za ra", vbTextCompare) > 0 Then
            txEnd = i - 1
            Exit For
        End If
    Next i

    ' R.B. anchori = standalone ceo broj (1-3 cifre) u [txStart..txEnd].
    Dim rbIdx() As Long, nRb As Long
    ReDim rbIdx(0 To (UBound(lines) - LBound(lines)))
    nRb = 0
    For i = txStart To txEnd
        If IsStandaloneRbAlta(lines(i)) Then
            rbIdx(nRb) = i
            nRb = nRb + 1
        End If
    Next i

    If nRb = 0 Then
        ParseBankaIzvodAlta = Empty
        Exit Function
    End If

    Dim tmp() As Variant
    ReDim tmp(1 To nRb, 1 To 10)
    Dim nValid As Long: nValid = 0

    Dim blockEnd As Long, c As Long, txn As Variant
    For i = 0 To nRb - 1
        If i < nRb - 1 Then
            blockEnd = rbIdx(i + 1) - 1
        Else
            blockEnd = txEnd
        End If

        txn = ParseOneAltaTxn(lines, rbIdx(i), blockEnd, izvodDatum)

        If IsRealTxnAlta(txn) Then
            nValid = nValid + 1
            For c = 1 To 10
                tmp(nValid, c) = txn(c - 1)
            Next c
        End If
    Next i

    If nValid = 0 Then
        ParseBankaIzvodAlta = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To nValid, 1 To 10)
    For i = 1 To nValid
        For c = 1 To 10
            result(i, c) = tmp(i, c)
        Next c
    Next i

    ParseBankaIzvodAlta = result
End Function

' Prava transakcija: ima racun (partner konto) ili bar jedan iznos > 0.
' Cisto prazan (0/0 bez racuna) blok je phantom -- odbacuje se.
Private Function IsRealTxnAlta(ByRef txn As Variant) As Boolean
    Dim racun As String
    Dim zad As Double, odo As Double
    racun = Trim$(CStr(txn(3)))
    zad = CDbl(txn(4))
    odo = CDbl(txn(5))
    If racun = "" And zad = 0 And odo = 0 Then Exit Function
    IsRealTxnAlta = True
End Function


' ---------- interna logika transakcije ----------

Private Function ParseOneAltaTxn(ByRef lines() As String, _
                                 ByVal rbIdx As Long, _
                                 ByVal blockEnd As Long, _
                                 ByVal izvodDatum As String) As Variant
    Dim datumKnjiz As String
    Dim partner As String, racun As String
    Dim zaduzenje As Double, odobrenje As Double
    Dim sifra As String, svrha As String, poziv As String, referenca As String
    Dim i As Long, ln As String

    ' --- racun + partner (partner = linije izmedju R.B. i racuna) ---
    Dim accIdx As Long: accIdx = -1
    For i = rbIdx + 1 To blockEnd
        If IsAccountLineAlta(lines(i)) Then
            accIdx = i
            Exit For
        End If
    Next i

    Dim scanFrom As Long
    If accIdx >= 0 Then
        racun = Trim$(lines(accIdx))
        For i = rbIdx + 1 To accIdx - 1
            ln = NormalizeSpacesAlta(lines(i))
            If ln <> "" Then
                If partner <> "" Then partner = partner & " "
                partner = partner & ln
            End If
        Next i
        scanFrom = accIdx + 1
    Else
        scanFrom = rbIdx + 1
    End If
    partner = NormalizeSpacesAlta(partner)

    ' --- dva datuma (knjizenja, prijema); poreklo naloga izmedju racuna i
    '     prvog datuma (naziv banke) se preskace ---
    Dim date1 As Long, date2 As Long
    date1 = -1: date2 = -1
    For i = scanFrom To blockEnd
        If IsDateLineAlta(lines(i)) Then
            If date1 < 0 Then
                date1 = i
            ElseIf date2 < 0 Then
                date2 = i
                Exit For
            End If
        End If
    Next i

    If date1 >= 0 Then datumKnjiz = Trim$(lines(date1))
    Dim afterDates As Long
    If date2 >= 0 Then
        afterDates = date2
    ElseIf date1 >= 0 Then
        afterDates = date1
    Else
        afterDates = scanFrom - 1
    End If

    ' --- naknada anchor razdvaja zaduzenje (pre) od sifra-linije (posle) ---
    Dim feeIdx As Long: feeIdx = -1
    For i = afterDates + 1 To blockEnd
        If InStr(1, lines(i), "Obr. naknada", vbTextCompare) > 0 Then
            feeIdx = i
            Exit For
        End If
    Next i

    ' --- zaduzenje (isplata) = standalone iznos izmedju datuma i naknade ---
    Dim zEnd As Long
    If feeIdx >= 0 Then zEnd = feeIdx - 1 Else zEnd = afterDates
    For i = afterDates + 1 To zEnd
        If IsAmountAlta(lines(i)) Then
            zaduzenje = ToNumberAlta(lines(i))
            Exit For
        End If
    Next i

    ' --- sifra-linija posle naknade:
    '     uplata  "<odobrenje> <sifra> <svrha>"  ILI
    '     isplata "- : - <sifra> <svrha>" (poziv u prefiksu ako postoji) ---
    Dim sifraIdx As Long: sifraIdx = -1
    Dim svrhaInline As String
    Dim startSifra As Long
    If feeIdx >= 0 Then startSifra = feeIdx + 1 Else startSifra = afterDates + 1
    For i = startSifra To blockEnd
        If ParseSifraLineAlta(lines(i), odobrenje, sifra, poziv, svrhaInline) Then
            sifraIdx = i
            Exit For
        End If
    Next i

    ' --- svrha + referenca: svrhaInline + linije posle sifra-linije do kraja bloka ---
    Dim tail As String
    tail = svrhaInline
    If sifraIdx >= 0 Then
        For i = sifraIdx + 1 To blockEnd
            ln = NormalizeSpacesAlta(lines(i))
            If ln <> "" Then
                If tail <> "" Then tail = tail & " "
                tail = tail & ln
            End If
        Next i
    End If
    tail = NormalizeSpacesAlta(tail)

    ' referenca (Podaci za reklamaciju) = poslednji standalone 15-cifreni token
    referenca = ExtractRefAlta(tail)
    If referenca <> "" Then tail = NormalizeSpacesAlta(Replace(tail, referenca, ""))

    ' ukloni [..] sistemske tagove; izvuci eventualni model/slash poziv iz svrhe
    tail = StripBracketsAlta(tail)
    If poziv = "" Then ExtractPozivFromSvrhaAlta tail, poziv
    svrha = NormalizeSpacesAlta(tail)

    ParseOneAltaTxn = Array(izvodDatum, datumKnjiz, partner, racun, _
                            zaduzenje, odobrenje, sifra, svrha, poziv, referenca)
End Function

' Parsira sifra-liniju. Vraca True i puni odobrenje/sifra/poziv/svrhaInline.
Private Function ParseSifraLineAlta(ByVal raw As String, _
                                    ByRef odobrenje As Double, _
                                    ByRef sifra As String, _
                                    ByRef poziv As String, _
                                    ByRef svrhaInline As String) As Boolean
    Dim s As String
    s = NormalizeSpacesAlta(raw)
    If s = "" Then Exit Function

    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False

    ' Uplata: "<odobrenje> <sifra> <svrha...>"
    re.pattern = "^(\d[\d,]*\.\d{2})\s+(\d{3})\b(.*)$"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        odobrenje = ToNumberAlta(m.SubMatches(0))
        sifra = m.SubMatches(1)
        svrhaInline = Trim$(m.SubMatches(2))
        ParseSifraLineAlta = True
        Exit Function
    End If

    ' Isplata: "<pozivZad> : <pozivOdob> <sifra> <svrha...>" (prazno polje = "-")
    re.pattern = "^(\S+)\s*:\s*(\S+)\s+(\d{3})\b(.*)$"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        Dim pz As String, po As String
        pz = Trim$(m.SubMatches(0))
        po = Trim$(m.SubMatches(1))
        sifra = m.SubMatches(2)
        svrhaInline = Trim$(m.SubMatches(3))
        If pz <> "-" And pz <> "" Then
            poziv = pz
        ElseIf po <> "-" And po <> "" Then
            poziv = po
        End If
        ParseSifraLineAlta = True
        Exit Function
    End If
End Function


' ---------- format helperi (Alta-lokalni; privatni) ----------

Private Function TryParseSaldoLineAlta(ByVal raw As String, _
                                       ByRef result As BankIzvodSaldo) As Boolean
    TryParseSaldoLineAlta = False

    Dim norm As String
    norm = NormalizeSpacesAlta(raw)
    If norm = "" Then Exit Function

    Dim parts() As String
    parts = Split(norm, " ")
    If (UBound(parts) - LBound(parts) + 1) <> 6 Then Exit Function

    Dim b As Long: b = LBound(parts)

    ' Prva 4 tokena = iznosi, zadnja 2 = celi brojevi (broj naloga Z/O).
    If Not IsAmountAlta(parts(b)) Then Exit Function
    If Not IsAmountAlta(parts(b + 1)) Then Exit Function
    If Not IsAmountAlta(parts(b + 2)) Then Exit Function
    If Not IsAmountAlta(parts(b + 3)) Then Exit Function
    If Not AllDigitsAlta(parts(b + 4)) Then Exit Function
    If Not AllDigitsAlta(parts(b + 5)) Then Exit Function

    result.PocetnoStanje = ToNumberAlta(parts(b))
    result.UkupanDuguje = ToNumberAlta(parts(b + 1))
    result.UkupanPotrazuje = ToNumberAlta(parts(b + 2))
    result.ZavrsnoStanje = ToNumberAlta(parts(b + 3))
    result.BrojNalogaZaduzenje = CLng(parts(b + 4))
    result.BrojNalogaOdobrenje = CLng(parts(b + 5))

    TryParseSaldoLineAlta = True
End Function

Private Function IsStandaloneRbAlta(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) >= 1 And Len(s) <= 3 And AllDigitsAlta(s) Then
        If val(s) >= 1 Then IsStandaloneRbAlta = True
    End If
End Function

Private Function IsDateLineAlta(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 10 Then IsDateLineAlta = (s Like "##.##.####")
End Function

Private Function IsAmountAlta(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d+(,\d{3})*\.\d{2}$"
    IsAmountAlta = re.Test(Trim$(s))
End Function

Private Function ToNumberAlta(ByVal s As String) As Double
    s = Replace(Trim$(s), ",", "")
    ToNumberAlta = val(s)
End Function

Private Function AllDigitsAlta(ByVal s As String) As Boolean
    Dim i As Long, ch As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next i
    AllDigitsAlta = True
End Function

Private Function IsAccountLineAlta(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d{3}-\d{5,20}-\d{2}$"
    IsAccountLineAlta = re.Test(Trim$(s))
End Function

Private Function ExtractAccountAnywhereAlta(ByVal s As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "(\d{3}-\d{5,20}-\d{2})"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ExtractAccountAnywhereAlta = m.SubMatches(0)
    End If
End Function

Private Function FirstDateInStringAlta(ByVal s As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "(\d{2}\.\d{2}\.\d{4})"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        FirstDateInStringAlta = m.SubMatches(0)
    End If
End Function

' referenca = poslednji token duzine 15 koji je ceo broj (Podaci za reklamaciju).
Private Function ExtractRefAlta(ByVal s As String) As String
    Dim toks() As String, i As Long, t As String, ref As String
    toks = Split(NormalizeSpacesAlta(s), " ")
    For i = LBound(toks) To UBound(toks)
        t = Trim$(toks(i))
        If Len(t) = 15 And AllDigitsAlta(t) Then ref = t
    Next i
    ExtractRefAlta = ref
End Function

Private Function StripBracketsAlta(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.pattern = "\[[^\]]*\]"
    StripBracketsAlta = NormalizeSpacesAlta(re.Replace(s, ""))
End Function

Private Sub ExtractPozivFromSvrhaAlta(ByRef svrha As String, ByRef poziv As String)
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False

    ' Model poziv "(NN) <broj>"
    re.pattern = "\(\d{2}\)\s*\d+"
    If re.Test(svrha) Then
        Set m = re.Execute(svrha)(0)
        poziv = NormalizeSpacesAlta(m.value)
        svrha = NormalizeSpacesAlta(Replace(svrha, m.value, ""))
        Exit Sub
    End If

    ' Slash poziv "003/26"
    re.pattern = "\d+/\d+"
    If re.Test(svrha) Then
        Set m = re.Execute(svrha)(0)
        poziv = m.value
        svrha = NormalizeSpacesAlta(Replace(svrha, poziv, ""))
    End If
End Sub

Private Function NormalizeSpacesAlta(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeSpacesAlta = s
End Function


' ---------- test / dijagnostika ----------

' Alt+F8: izaberi ALTA PDF -> detekcija + saldo + pun parse + per-red dump.
Public Sub Test_AltaParse()
    Dim pdfPath As String, txt As String, lines() As String
    Dim saldo As BankIzvodSaldo, parsed As Variant, i As Long

    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub

    txt = ExtractTextFromPdf(pdfPath)
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)

    Debug.Print "=== DetectBank: " & DetectBank(lines) & " ==="
    Debug.Print "Broj izvoda:  " & ExtractIzvodBrojAlta(lines)
    Debug.Print "Datum izvoda: " & ExtractIzvodDatumAlta(lines)
    Debug.Print "Racun izvoda: " & ExtractIzvodRacunAlta(lines)

    saldo = ExtractIzvodSaldoAlta(lines)
    Debug.Print "--- Saldo (parsed=" & saldo.parsed & ") ---"
    Debug.Print "Pocetno=" & saldo.PocetnoStanje & " Duguje=" & saldo.UkupanDuguje & _
                " Potrazuje=" & saldo.UkupanPotrazuje & " Novo=" & saldo.ZavrsnoStanje & _
                " nZ=" & saldo.BrojNalogaZaduzenje & " nO=" & saldo.BrojNalogaOdobrenje

    On Error Resume Next
    parsed = ParseBankaIzvodForImport(txt, "test.pdf")
    If Err.Number <> 0 Then
        Debug.Print "FULL PARSE FAIL: [" & Err.Number & "] " & Err.description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If IsEmpty(parsed) Then
        Debug.Print "FULL PARSE: Empty (nema transakcija)"
        Exit Sub
    End If

    Debug.Print "--- OK: " & UBound(parsed, 1) & " transakcija ---"
    For i = 1 To UBound(parsed, 1)
        Debug.Print i & " | " & parsed(i, 4) & " | " & parsed(i, 5) & _
                    " | racun=" & parsed(i, 6) & _
                    " | Isl=" & parsed(i, 8) & " Upl=" & parsed(i, 7) & _
                    " | sif=" & parsed(i, 9) & " | " & parsed(i, 10) & _
                    " | poz=" & parsed(i, 11) & " | ref=" & parsed(i, 12)
    Next i
End Sub
