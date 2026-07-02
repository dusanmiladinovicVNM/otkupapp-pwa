Attribute VB_Name = "modBankaHalk"
Option Explicit

' ============================================================
' modBankaHalk
' Parser za pdftotext izlaz Halkbank izvoda (racun 155-...).
'
' Ugovor (isto kao ostali bank-parseri): orkestrator
' modBankaImport.ParseBankaIzvodForImport rutira na 5 funkcija:
'   ExtractIzvodBrojHalk / ExtractIzvodDatumHalk / ExtractIzvodRacunHalk /
'   ExtractIzvodSaldoHalk / ParseBankaIzvodHalk
' Deljeni 4-nivo integrity + 17-kol staging ostaju u orkestratoru.
'
' Output-kolone ParseBankaIzvodHalk (isto kao ostali):
' 1 Datum Izvoda  2 Datum Izvrs  3 Partner  4 Racun
' 5 Zaduzenje(isplata)  6 Odobrenje(uplata)  7 Sifra
' 8 Svrha  9 Poziv na broj  10 Referenca (Referentna oznaka)
'
' KRITICNO: Halkbank izvod ima ODVOJEN blok "NEIZVRSENI NALOZI" (nalozi na
' cekanju, sopstvena numeracija 1..N). Ti NE ulaze u saldo. Parser se BOUND-uje
' na izvrsenu sekciju: staje na prvi "Ukupno na ra" / "Ukupno RSD"; sve posle
' (NEIZVRSENI + footer) se ignorise.
'
' Layout izvrsene transakcije (fiksni redosled):
'   [R.B.]            "N."  (ceo broj + tacka)
'   partner           Poslovno ime i sediste (1-2 linije)
'   racun             \d{3}-\d+-\d{2}
'   poreklo naloga    naziv banke (ignorise se)
'   datum izvrsenja   dd.mm.yyyy   <-- DatumTransakcije
'   datum prijema     dd.mm.yyyy
'   zaduzenje         iznos (standalone)
'   "Obr. naknada X"  naknada (ignorise se)
'   "<odobrenje> <sifra>"   npr. "0.00 221" ili "64,271.72 270"
'   svrha             1-2 linije, na kraju [poziv] + referenca "0870011..."
'   "N (M)"           trailing (ignorise se)
'
' Broj-format: 1,234.56 (isti kao ostali). SALDO data-linija: 4 iznosa + 2 broja
'   ("0.00 90,891.72 90,891.72 0.00 3 4"), ispod labela "... duguje potrazuje ...".
' ============================================================


' ---------- 5 ugovornih funkcija ----------

Public Function ExtractIzvodBrojHalk(ByRef lines() As String) As String
    Dim re As Object, m As Object, i As Long, s As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    ' "IZVOD 128" (naslov) ili "Izvod broj 128" (footer)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        re.pattern = "IZVOD\s+broj\s+(\d+)"
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            ExtractIzvodBrojHalk = Trim$(m.SubMatches(0))
            Exit Function
        End If
        re.pattern = "IZVOD\s+(\d+)"
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            ExtractIzvodBrojHalk = Trim$(m.SubMatches(0))
            Exit Function
        End If
    Next i
End Function

Public Function ExtractIzvodDatumHalk(ByRef lines() As String) As String
    Dim i As Long, s As String, tmp As String, p As Long

    ' Primarno: "Izvod za datum: dd.mm.yyyy" (footer)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        p = InStr(1, s, "Izvod za datum:", vbTextCompare)
        If p > 0 Then
            tmp = Trim$(Mid$(s, p + Len("Izvod za datum:")))
            If IsDateLineHalk(Left$(tmp, 10)) Then
                ExtractIzvodDatumHalk = Left$(tmp, 10)
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
            If IsDateLineHalk(Left$(tmp, 10)) Then
                ExtractIzvodDatumHalk = Left$(tmp, 10)
                Exit Function
            End If
        End If
    Next i
End Function

Public Function ExtractIzvodRacunHalk(ByRef lines() As String) As String
    Dim i As Long, s As String, acc As String

    ' "Izvod broj 128 Za racun 155-... Izvod za datum: ..." (footer)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If InStr(1, s, "Izvod broj", vbTextCompare) > 0 And InStr(1, s, "za ra", vbTextCompare) > 0 Then
            acc = ExtractAccountAnywhereHalk(s)
            If acc <> "" Then
                ExtractIzvodRacunHalk = acc
                Exit Function
            End If
        End If
    Next i

    ' Fallback: bilo koja "Izvod za datum:" linija (sadrzi racun)
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If InStr(1, s, "Izvod za datum:", vbTextCompare) > 0 Then
            acc = ExtractAccountAnywhereHalk(s)
            If acc <> "" Then
                ExtractIzvodRacunHalk = acc
                Exit Function
            End If
        End If
    Next i
End Function

Public Function ExtractIzvodSaldoHalk(ByRef lines() As String) As BankIzvodSaldo
    Dim result As BankIzvodSaldo
    result.parsed = False

    On Error GoTo EH

    Dim i As Long, j As Long, anchorIdx As Long
    anchorIdx = -1

    ' Anchor: label linija "stanje duguje potrazuje novo stanje ..." (rec "duguje" je ASCII).
    For i = LBound(lines) To UBound(lines)
        If InStr(1, NormalizeSpacesHalk(lines(i)), "duguje", vbTextCompare) > 0 Then
            anchorIdx = i
            Exit For
        End If
    Next i

    If anchorIdx >= 0 Then
        Dim scanEnd As Long
        scanEnd = anchorIdx + 4
        If scanEnd > UBound(lines) Then scanEnd = UBound(lines)
        For j = anchorIdx + 1 To scanEnd
            If TryParseSaldoLineHalk(lines(j), result) Then
                result.parsed = True
                ExtractIzvodSaldoHalk = result
                Exit Function
            End If
        Next j
    End If

    ' Fallback: globalno trazi jedinu 4-iznosa+2-broja liniju.
    For i = LBound(lines) To UBound(lines)
        If TryParseSaldoLineHalk(lines(i), result) Then
            result.parsed = True
            Exit For
        End If
    Next i

    ExtractIzvodSaldoHalk = result
    Exit Function

EH:
    result.parsed = False
    ExtractIzvodSaldoHalk = result
End Function

Public Function ParseBankaIzvodHalk(ByVal txt As String) As Variant
    Dim lines() As String
    Dim izvodDatum As String
    Dim i As Long

    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)

    izvodDatum = ExtractIzvodDatumHalk(lines)

    ' BOUND na izvrsenu sekciju: prvi "Ukupno na ra"/"Ukupno RSD" -> kraj.
    ' Sve posle (NEIZVRSENI NALOZI + footer) se ignorise.
    Dim executedEnd As Long
    executedEnd = UBound(lines)
    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), "Ukupno na ra", vbTextCompare) > 0 _
           Or InStr(1, lines(i), "Ukupno RSD", vbTextCompare) > 0 Then
            executedEnd = i - 1
            Exit For
        End If
    Next i

    ' R.B. anchori ("N.") u [0..executedEnd].
    Dim rbIdx() As Long, nRb As Long
    ReDim rbIdx(0 To (UBound(lines) - LBound(lines)))
    nRb = 0
    For i = LBound(lines) To executedEnd
        If IsRbLineHalk(lines(i)) Then
            rbIdx(nRb) = i
            nRb = nRb + 1
        End If
    Next i

    If nRb = 0 Then
        ParseBankaIzvodHalk = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To nRb, 1 To 10)

    Dim blockEnd As Long, c As Long, txn As Variant
    For i = 0 To nRb - 1
        If i < nRb - 1 Then
            blockEnd = rbIdx(i + 1) - 1
        Else
            blockEnd = executedEnd
        End If

        txn = ParseOneHalkTxn(lines, rbIdx(i), blockEnd, izvodDatum)

        For c = 1 To 10
            result(i + 1, c) = txn(c - 1)
        Next c
    Next i

    ParseBankaIzvodHalk = result
End Function


' ---------- interna logika transakcije ----------

Private Function ParseOneHalkTxn(ByRef lines() As String, _
                                 ByVal rbIdx As Long, _
                                 ByVal blockEnd As Long, _
                                 ByVal izvodDatum As String) As Variant
    Dim datumIzvrs As String
    Dim partner As String, racun As String
    Dim zaduzenje As Double, odobrenje As Double
    Dim sifra As String, svrha As String, poziv As String, referenca As String
    Dim i As Long, ln As String

    ' --- racun + partner ---
    Dim accIdx As Long: accIdx = -1
    For i = rbIdx + 1 To blockEnd
        If IsAccountLineHalk(lines(i)) Then
            accIdx = i
            Exit For
        End If
    Next i

    Dim scanFrom As Long
    If accIdx >= 0 Then
        racun = Trim$(lines(accIdx))
        For i = rbIdx + 1 To accIdx - 1
            ln = NormalizeSpacesHalk(lines(i))
            If ln <> "" Then
                If partner <> "" Then partner = partner & " "
                partner = partner & ln
            End If
        Next i
        scanFrom = accIdx + 1
    Else
        scanFrom = rbIdx + 1
    End If
    partner = NormalizeSpacesHalk(partner)

    ' --- dva datuma (izvrsenja, prijema) ---
    Dim date1 As Long, date2 As Long
    date1 = -1: date2 = -1
    For i = scanFrom To blockEnd
        If IsDateLineHalk(lines(i)) Then
            If date1 < 0 Then
                date1 = i
            ElseIf date2 < 0 Then
                date2 = i
                Exit For
            End If
        End If
    Next i

    If date1 >= 0 Then datumIzvrs = Trim$(lines(date1))
    Dim afterDates As Long
    If date2 >= 0 Then
        afterDates = date2
    ElseIf date1 >= 0 Then
        afterDates = date1
    Else
        afterDates = scanFrom - 1
    End If

    ' --- zaduzenje = iznos odmah posle datuma prijema ---
    If afterDates + 1 <= blockEnd Then
        If IsAmountHalk(lines(afterDates + 1)) Then zaduzenje = ToNumberHalk(lines(afterDates + 1))
    End If

    ' --- odobrenje + sifra = prva linija "<iznos> <3 cifre>" posle datuma ---
    Dim odoSifIdx As Long: odoSifIdx = -1
    For i = afterDates + 1 To blockEnd
        If ParseOdoSifHalk(lines(i), odobrenje, sifra) Then
            odoSifIdx = i
            Exit For
        End If
    Next i

    ' --- svrha / poziv / referenca (do prve 13-cifrene reference) ---
    Dim startTail As Long
    If odoSifIdx >= 0 Then startTail = odoSifIdx + 1 Else startTail = afterDates + 2

    Dim svrhaRaw As String, toks() As String, t As Long, tok As String
    Dim doneRef As Boolean: doneRef = False
    For i = startTail To blockEnd
        If doneRef Then Exit For
        toks = Split(NormalizeSpacesHalk(lines(i)), " ")
        For t = LBound(toks) To UBound(toks)
            tok = toks(t)
            If tok <> "" Then
                If IsRef13Halk(tok) Then
                    referenca = tok
                    doneRef = True
                    Exit For
                End If
                If svrhaRaw <> "" Then svrhaRaw = svrhaRaw & " "
                svrhaRaw = svrhaRaw & tok
            End If
        Next t
    Next i

    svrhaRaw = NormalizeSpacesHalk(svrhaRaw)
    ExtractPozivHalk svrhaRaw, poziv     ' izvuce model/dashed poziv; ostatak = svrha
    svrha = svrhaRaw

    ParseOneHalkTxn = Array(izvodDatum, datumIzvrs, partner, racun, _
                            zaduzenje, odobrenje, sifra, svrha, poziv, referenca)
End Function


' ---------- format helperi (Halk-lokalni; privatni) ----------

Private Function IsRbLineHalk(ByVal s As String) As Boolean
    s = Trim$(s)
    ' "1." "12." itd. (ceo broj + tacka)
    If Len(s) >= 2 And Right$(s, 1) = "." Then
        IsRbLineHalk = AllDigitsHalk(Left$(s, Len(s) - 1))
    End If
End Function

Private Function IsDateLineHalk(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 10 Then IsDateLineHalk = (s Like "##.##.####")
End Function

Private Function IsAmountHalk(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d+(,\d{3})*\.\d{2}$"
    IsAmountHalk = re.Test(Trim$(s))
End Function

Private Function ParseOdoSifHalk(ByVal s As String, _
                                 ByRef odobrenje As Double, _
                                 ByRef sifra As String) As Boolean
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^(\d+(,\d{3})*\.\d{2})\s+(\d{3})\b"
    If re.Test(NormalizeSpacesHalk(s)) Then
        Set m = re.Execute(NormalizeSpacesHalk(s))(0)
        odobrenje = ToNumberHalk(m.SubMatches(0))
        sifra = m.SubMatches(2)
        ParseOdoSifHalk = True
    End If
End Function

Private Function ToNumberHalk(ByVal s As String) As Double
    s = Replace(Trim$(s), ",", "")
    ToNumberHalk = val(s)
End Function

Private Function AllDigitsHalk(ByVal s As String) As Boolean
    Dim i As Long, ch As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then Exit Function
    Next i
    AllDigitsHalk = True
End Function

Private Function IsRef13Halk(ByVal s As String) As Boolean
    ' Referentna oznaka = 13-cifreni ceo broj (u uzorku prefiks 0870011).
    IsRef13Halk = (Len(s) = 13 And AllDigitsHalk(s))
End Function

Private Function IsAccountLineHalk(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d{3}-\d{5,20}-\d{2}$"
    IsAccountLineHalk = re.Test(Trim$(s))
End Function

Private Function ExtractAccountAnywhereHalk(ByVal s As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "(\d{3}-\d{5,20}-\d{2})"
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        ExtractAccountAnywhereHalk = m.SubMatches(0)
    End If
End Function

Private Function TryParseSaldoLineHalk(ByVal raw As String, _
                                       ByRef result As BankIzvodSaldo) As Boolean
    TryParseSaldoLineHalk = False

    Dim norm As String
    norm = NormalizeSpacesHalk(raw)
    If norm = "" Then Exit Function

    Dim parts() As String
    parts = Split(norm, " ")
    If (UBound(parts) - LBound(parts) + 1) <> 6 Then Exit Function

    Dim b As Long: b = LBound(parts)
    If Not IsAmountHalk(parts(b)) Then Exit Function
    If Not IsAmountHalk(parts(b + 1)) Then Exit Function
    If Not IsAmountHalk(parts(b + 2)) Then Exit Function
    If Not IsAmountHalk(parts(b + 3)) Then Exit Function
    If Not AllDigitsHalk(parts(b + 4)) Then Exit Function
    If Not AllDigitsHalk(parts(b + 5)) Then Exit Function

    result.PocetnoStanje = ToNumberHalk(parts(b))
    result.UkupanDuguje = ToNumberHalk(parts(b + 1))
    result.UkupanPotrazuje = ToNumberHalk(parts(b + 2))
    result.ZavrsnoStanje = ToNumberHalk(parts(b + 3))
    result.BrojNalogaZaduzenje = CLng(parts(b + 4))
    result.BrojNalogaOdobrenje = CLng(parts(b + 5))

    TryParseSaldoLineHalk = True
End Function

Private Function NormalizeSpacesHalk(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeSpacesHalk = s
End Function

Private Sub ExtractPozivHalk(ByRef svrha As String, ByRef poziv As String)
    Dim re As Object, m As Object

    ' 1) Model poziv "(NN) <broj>"
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "\(\d{2}\)\s*\d+"
    If re.Test(svrha) Then
        Set m = re.Execute(svrha)(0)
        poziv = NormalizeSpacesHalk(m.value)
        svrha = NormalizeSpacesHalk(Replace(svrha, m.value, ""))
        Exit Sub
    End If

    ' 2) Dashed poziv "NNNNNN-NNNNN..."
    re.pattern = "\d{5,}-\d{5,}"
    If re.Test(svrha) Then
        Set m = re.Execute(svrha)(0)
        poziv = m.value
        svrha = NormalizeSpacesHalk(Replace(svrha, m.value, ""))
    End If
End Sub
