Attribute VB_Name = "modBankaProCredit"
Option Explicit

' ============================================================
' modBankaProCredit
' Parser za pdftotext izlaz ProCredit Bank izvoda (racun 220-...).
'
' Ugovor (isti kao Komercijalna parser): orkestrator
' modBankaImport.ParseBankaIzvodForImport rutira na 5 funkcija:
'   ExtractIzvodBrojProCredit / ExtractIzvodDatumProCredit /
'   ExtractIzvodRacunProCredit / ExtractIzvodSaldoProCredit /
'   ParseBankaIzvodProCredit
' Deljeni 4-nivo integrity + 17-kol staging ostaju u orkestratoru.
'
' Output-kolone ParseBankaIzvodProCredit (isto kao Komercijalna):
' 1 Datum Izvoda   2 Datum Izvrs   3 Partner   4 Racun
' 5 Zaduzenje(isplata)   6 Odobrenje(uplata)   7 Sifra
' 8 Svrha   9 Poziv na broj   10 Referenca (Doznake)
'
' Layout ProCredit izvoda (pdftotext -raw), po transakciji (fiksni redosled):
'   [R.B.]            standalone ceo broj (1..N)
'   partner           naziv i sediste primaoca (1-2 linije)
'   racun             205-.../105-... (\d{3}-\d+-\d{2})
'   poreklo naloga    naziv banke (ignorise se; nema kolonu)
'   [datum prijema]   standalone dd.mm.yyyy  <-- PIVOT transakcije
'   iznos Zaduzenje
'   provizija Zad     (ignorise se za saldo)
'   iznos Odobrenje
'   provizija Odo     (ignorise se za saldo)
'   sifra             npr. 221
'   svrha             1-2 linije
'   [poziv na broj]   opciono ("003/26" ili standalone "2026")
'   referenca         Doznake, standalone ceo broj >= 6 cifara
'
' Broj-format: 1,234.56 (zarez hiljade, tacka decimala) - isti kao Komercijalna.
'
' STANJE blok (saldo, jedna data-linija: 4 iznosa + 2 cela broja):
'   "Prethodno stanje" (anchor) ... "22,601.38 1,491,942.25 2,009,910.30 540,569.43 14 2"
'   = Pocetno  Duguje  Potrazuje  Novo  nZaduzenja  nOdobrenja
' ============================================================


' ---------- 5 ugovornih funkcija ----------

Public Function ExtractIzvodBrojProCredit(ByRef lines() As String) As String
    Dim re As Object, m As Object, i As Long, s As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "IZVOD\s+(\d+)"

    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            ExtractIzvodBrojProCredit = Trim$(m.SubMatches(0))
            Exit Function
        End If
    Next i
End Function

Public Function ExtractIzvodDatumProCredit(ByRef lines() As String) As String
    Dim i As Long, s As String, tmp As String, p As Long

    ' Primarno: "Izvod za datum: dd.mm.yyyy" (footer)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        p = InStr(1, s, "Izvod za datum:", vbTextCompare)
        If p > 0 Then
            tmp = Trim$(Mid$(s, p + Len("Izvod za datum:")))
            If IsDateLinePC(Left$(tmp, 10)) Then
                ExtractIzvodDatumProCredit = Left$(tmp, 10)
                Exit Function
            End If
        End If
    Next i

    ' Fallback: "... NA DAN dd.mm.yyyy" (naslov)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        p = InStr(1, s, "NA DAN", vbTextCompare)
        If p > 0 Then
            tmp = Trim$(Mid$(s, p + Len("NA DAN")))
            If IsDateLinePC(Left$(tmp, 10)) Then
                ExtractIzvodDatumProCredit = Left$(tmp, 10)
                Exit Function
            End If
        End If
    Next i
End Function

Public Function ExtractIzvodRacunProCredit(ByRef lines() As String) As String
    Dim i As Long, s As String, acc As String

    ' Header/footer: "IZVOD 162 za racun 220-... Izvod za datum: ..."
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If InStr(1, s, "IZVOD", vbTextCompare) > 0 And InStr(1, s, "za ra", vbTextCompare) > 0 Then
            acc = ExtractAccountAnywherePC(s)
            If acc <> "" Then
                ExtractIzvodRacunProCredit = acc
                Exit Function
            End If
        End If
    Next i

    ' Fallback: footer linija sa "Izvod za datum:" takodje sadrzi racun
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If InStr(1, s, "Izvod za datum:", vbTextCompare) > 0 Then
            acc = ExtractAccountAnywherePC(s)
            If acc <> "" Then
                ExtractIzvodRacunProCredit = acc
                Exit Function
            End If
        End If
    Next i
End Function

Public Function ExtractIzvodSaldoProCredit(ByRef lines() As String) As BankIzvodSaldo
    Dim result As BankIzvodSaldo
    result.parsed = False

    On Error GoTo EH

    Dim i As Long, j As Long, anchorIdx As Long
    anchorIdx = -1

    For i = LBound(lines) To UBound(lines)
        If InStr(1, NormalizeSpacesPC(lines(i)), "Prethodno stanje", vbTextCompare) > 0 Then
            anchorIdx = i
            Exit For
        End If
    Next i

    If anchorIdx < 0 Then
        ExtractIzvodSaldoProCredit = result
        Exit Function
    End If

    ' Labeli su rasporedjeni preko vise linija; data-linija je do ~10 ispod anchora.
    Dim scanEnd As Long
    scanEnd = anchorIdx + 20
    If scanEnd > UBound(lines) Then scanEnd = UBound(lines)

    For j = anchorIdx + 1 To scanEnd
        If TryParseSaldoLinePC(lines(j), result) Then
            result.parsed = True
            Exit For
        End If
    Next j

    ExtractIzvodSaldoProCredit = result
    Exit Function

EH:
    ' Parser nikad ne raise-uje, samo kaze nije parsed.
    result.parsed = False
    ExtractIzvodSaldoProCredit = result
End Function

Public Function ParseBankaIzvodProCredit(ByVal txt As String) As Variant
    Dim lines() As String
    Dim izvodDatum As String
    Dim i As Long

    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)

    izvodDatum = ExtractIzvodDatumProCredit(lines)

    ' 1) Pivoti = standalone dd.mm.yyyy linije (jedan datum prijema po transakciji).
    Dim pivots() As Long, nPiv As Long
    ReDim pivots(0 To (UBound(lines) - LBound(lines)))
    nPiv = 0
    For i = LBound(lines) To UBound(lines)
        If IsDateLinePC(lines(i)) Then
            pivots(nPiv) = i
            nPiv = nPiv + 1
        End If
    Next i

    If nPiv = 0 Then
        ParseBankaIzvodProCredit = Empty
        Exit Function
    End If

    ' 2) R.B. iznad svakog pivota (prvi standalone kratak ceo broj) = gornja granica bloka.
    Dim rbIdx() As Long
    ReDim rbIdx(0 To nPiv - 1)
    For i = 0 To nPiv - 1
        rbIdx(i) = FindRbAbovePC(lines, pivots(i))
    Next i

    ' 3) Parsiraj svaku transakciju; donja granica = R.B. sledece (ili kraj).
    '    Validacija odbacuje phantom transakcije (standalone datum van tabele):
    '    integrity ih NE hvata jer 0/0 red ne menja ni sume ni brojeve naloga.
    Dim tmp() As Variant
    ReDim tmp(1 To nPiv, 1 To 10)
    Dim nValid As Long: nValid = 0

    Dim afterEnd As Long, c As Long, txn As Variant
    For i = 0 To nPiv - 1
        If i < nPiv - 1 Then
            afterEnd = rbIdx(i + 1) - 1
        Else
            afterEnd = UBound(lines)
        End If
        If afterEnd < pivots(i) Then afterEnd = UBound(lines)

        txn = ParseOneProCreditTxn(lines, rbIdx(i), pivots(i), afterEnd, izvodDatum)

        If IsRealTxnPC(txn) Then
            nValid = nValid + 1
            For c = 1 To 10
                tmp(nValid, c) = txn(c - 1)
            Next c
        End If
    Next i

    If nValid = 0 Then
        ParseBankaIzvodProCredit = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To nValid, 1 To 10)
    For i = 1 To nValid
        For c = 1 To 10
            result(i, c) = tmp(i, c)
        Next c
    Next i

    ParseBankaIzvodProCredit = result
End Function

' Prava transakcija: ima datum transakcije + (racun ili bar jedan iznos > 0).
' Odbacuje phantom pivote (standalone datum bez racuna i bez iznosa).
Private Function IsRealTxnPC(ByRef txn As Variant) As Boolean
    Dim datum As String, racun As String
    Dim zad As Double, odo As Double

    datum = Trim$(CStr(txn(1)))
    racun = Trim$(CStr(txn(3)))
    zad = CDbl(txn(4))
    odo = CDbl(txn(5))

    If datum = "" Then Exit Function
    If racun = "" And zad = 0 And odo = 0 Then Exit Function
    IsRealTxnPC = True
End Function


' ---------- interna logika transakcije ----------

Private Function ParseOneProCreditTxn(ByRef lines() As String, _
                                      ByVal rbIdx As Long, _
                                      ByVal pivotIdx As Long, _
                                      ByVal afterEnd As Long, _
                                      ByVal izvodDatum As String) As Variant
    Dim datumIzvrs As String
    Dim partner As String, racun As String
    Dim zaduzenje As Double, odobrenje As Double
    Dim sifra As String, svrha As String, poziv As String, referenca As String
    Dim i As Long, ln As String

    datumIzvrs = Trim$(lines(pivotIdx))

    ' --- Pre datuma: partner (do racuna); poreklo naloga ignorisemo ---
    Dim accIdx As Long: accIdx = -1
    For i = rbIdx + 1 To pivotIdx - 1
        If IsAccountLinePC(lines(i)) Then
            accIdx = i
            Exit For
        End If
    Next i

    Dim partEnd As Long
    If accIdx >= 0 Then
        partEnd = accIdx - 1
        racun = Trim$(lines(accIdx))
    Else
        partEnd = pivotIdx - 1
    End If

    For i = rbIdx + 1 To partEnd
        ln = NormalizeSpacesPC(lines(i))
        If ln <> "" Then
            If partner <> "" Then partner = partner & " "
            partner = partner & ln
        End If
    Next i
    partner = NormalizeSpacesPC(partner)

    ' --- Posle datuma: 4 iznosa (Zad, provZ, Odo, provO), pa sifra ---
    ' Bound na afterEnd (kraj OVE transakcije), ne UBound(lines) -- da se kod
    ' loseg preloma ne procita linija iz sledece transakcije.
    If pivotIdx + 1 <= afterEnd Then
        If IsAmountPC(lines(pivotIdx + 1)) Then zaduzenje = ToNumberPC(lines(pivotIdx + 1))
    End If
    If pivotIdx + 3 <= afterEnd Then
        If IsAmountPC(lines(pivotIdx + 3)) Then odobrenje = ToNumberPC(lines(pivotIdx + 3))
    End If

    Dim sifraIdx As Long: sifraIdx = pivotIdx + 5
    If sifraIdx <= afterEnd Then sifra = Trim$(lines(sifraIdx))

    ' --- Referenca (Doznake) = prvi standalone ceo broj >= 6 cifara posle sifre ---
    Dim refIdx As Long: refIdx = -1
    For i = sifraIdx + 1 To afterEnd
        If IsStandaloneIntMinPC(lines(i), 6) Then
            refIdx = i
            Exit For
        End If
    Next i
    If refIdx >= 0 Then referenca = Trim$(lines(refIdx))

    ' --- Svrha (+ standalone poziv) = linije izmedju sifre i reference ---
    Dim svrhaEnd As Long
    If refIdx >= 0 Then svrhaEnd = refIdx - 1 Else svrhaEnd = afterEnd

    For i = sifraIdx + 1 To svrhaEnd
        ln = NormalizeSpacesPC(lines(i))
        If ln = "" Then GoTo NextSvrha

        If IsPureIntPC(ln) And Len(ln) < 6 Then
            poziv = ln                       ' standalone kratak broj = (Model) poziv na broj
        Else
            If svrha <> "" Then svrha = svrha & " "
            svrha = svrha & ln
        End If
NextSvrha:
    Next i
    svrha = NormalizeSpacesPC(svrha)

    ' Poziv embedovan u svrhu ("... potrosnja 003/26")
    If poziv = "" Then ExtractSlashPozivPC svrha, poziv

    ParseOneProCreditTxn = Array(izvodDatum, datumIzvrs, partner, racun, _
                                 zaduzenje, odobrenje, sifra, svrha, poziv, referenca)
End Function

Private Function FindRbAbovePC(ByRef lines() As String, ByVal pivotIdx As Long) As Long
    Dim i As Long
    For i = pivotIdx - 1 To LBound(lines) Step -1
        If IsStandaloneIntPC(lines(i), 3) Then
            FindRbAbovePC = i
            Exit Function
        End If
    Next i
    FindRbAbovePC = LBound(lines)
End Function

Private Function TryParseSaldoLinePC(ByVal raw As String, _
                                     ByRef result As BankIzvodSaldo) As Boolean
    TryParseSaldoLinePC = False

    Dim norm As String
    norm = NormalizeSpacesPC(raw)
    If norm = "" Then Exit Function

    Dim parts() As String
    parts = Split(norm, " ")
    If (UBound(parts) - LBound(parts) + 1) <> 6 Then Exit Function

    Dim b As Long: b = LBound(parts)

    ' Prva 4 tokena = iznosi, zadnja 2 = celi brojevi (broj naloga Z/O).
    If Not IsAmountPC(parts(b)) Then Exit Function
    If Not IsAmountPC(parts(b + 1)) Then Exit Function
    If Not IsAmountPC(parts(b + 2)) Then Exit Function
    If Not IsAmountPC(parts(b + 3)) Then Exit Function
    If Not IsPureIntPC(parts(b + 4)) Then Exit Function
    If Not IsPureIntPC(parts(b + 5)) Then Exit Function

    result.PocetnoStanje = ToNumberPC(parts(b))
    result.UkupanDuguje = ToNumberPC(parts(b + 1))
    result.UkupanPotrazuje = ToNumberPC(parts(b + 2))
    result.ZavrsnoStanje = ToNumberPC(parts(b + 3))
    result.BrojNalogaZaduzenje = CLng(parts(b + 4))
    result.BrojNalogaOdobrenje = CLng(parts(b + 5))

    TryParseSaldoLinePC = True
End Function


' ---------- format helperi (ProCredit-lokalni; privatni) ----------

Private Function IsDateLinePC(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 10 Then IsDateLinePC = (s Like "##.##.####")
End Function

Private Function IsAmountPC(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d+(,\d{3})*\.\d{2}$"
    IsAmountPC = re.Test(Trim$(s))
End Function

Private Function ToNumberPC(ByVal s As String) As Double
    s = Replace(Trim$(s), ",", "")
    ToNumberPC = val(s)
End Function

Private Function AllDigitsPC(ByVal s As String) As Boolean
    Dim i As Long, ch As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next i
    AllDigitsPC = True
End Function

Private Function IsPureIntPC(ByVal s As String) As Boolean
    IsPureIntPC = AllDigitsPC(s)
End Function

Private Function IsStandaloneIntPC(ByVal s As String, ByVal maxLen As Long) As Boolean
    s = Trim$(s)
    If AllDigitsPC(s) Then
        If Len(s) >= 1 And Len(s) <= maxLen Then IsStandaloneIntPC = True
    End If
End Function

Private Function IsStandaloneIntMinPC(ByVal s As String, ByVal minLen As Long) As Boolean
    s = Trim$(s)
    If AllDigitsPC(s) Then
        If Len(s) >= minLen Then IsStandaloneIntMinPC = True
    End If
End Function

Private Function IsAccountLinePC(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d{3}-\d{5,20}-\d{2}$"
    IsAccountLinePC = re.Test(Trim$(s))
End Function

Private Function ExtractAccountAnywherePC(ByVal s As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "(\d{3}-\d{5,20}-\d{2})"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ExtractAccountAnywherePC = m.SubMatches(0)
    End If
End Function

Private Function NormalizeSpacesPC(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeSpacesPC = s
End Function

Private Sub ExtractSlashPozivPC(ByRef svrha As String, ByRef poziv As String)
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "(\d+/\d+)"
    If re.Test(svrha) Then
        Set m = re.Execute(svrha)(0)
        poziv = m.SubMatches(0)
        svrha = NormalizeSpacesPC(Replace(svrha, poziv, ""))
    End If
End Sub


' ---------- test / dijagnostika ----------

' Alt+F8: izaberi ProCredit PDF -> detekcija + saldo + pun parse + per-red dump.
Public Sub Test_ProCreditParse()
    Dim pdfPath As String, txt As String, lines() As String
    Dim saldo As BankIzvodSaldo, parsed As Variant, i As Long

    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub

    txt = ExtractTextFromPdf(pdfPath)
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)

    Debug.Print "=== DetectBank: " & DetectBank(lines) & " ==="
    Debug.Print "Broj izvoda:  " & ExtractIzvodBrojProCredit(lines)
    Debug.Print "Datum izvoda: " & ExtractIzvodDatumProCredit(lines)
    Debug.Print "Racun izvoda: " & ExtractIzvodRacunProCredit(lines)

    saldo = ExtractIzvodSaldoProCredit(lines)
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
