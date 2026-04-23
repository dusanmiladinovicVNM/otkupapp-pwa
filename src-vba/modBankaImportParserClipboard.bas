Attribute VB_Name = "modBankaImportParserClipboard"
Option Explicit

' ============================================================
' modBankaImport – Parser für Komercijalna Banka Izvod
' Input: Clipboard-Text (Strg+A, Strg+C aus PDF)
' Output: 2D Array der Transaktionen
'
' Spalten:
' 1  Datum Izvoda
' 2  Datum Izvrsenja
' 3  Partner
' 4  Racun
' 5  Zaduzenje
' 6  Odobrenje
' 7  Sifra
' 8  Svrha
' 9  Poziv na broj
' 10 Referenca
' ============================================================

Public Function ParseBankaIzvod(ByVal txt As String) As Variant
    Dim Lines() As String
    txt = Replace(txt, Chr$(12), vbLf)
    Lines = Split(Replace(txt, vbCr, ""), vbLf)
    
    ' Header parsen
    Dim izvodBroj As String, izvodDatum As String
    Dim i As Long
    
    For i = LBound(Lines) To UBound(Lines)
        If Left$(Trim$(Lines(i)), 11) = "Izvod broj" Then
            izvodBroj = ExtractAfter(Lines(i), "Izvod broj ")
            Dim pZa As Long
            pZa = InStr(izvodBroj, " Za")
            If pZa > 0 Then izvodBroj = Left$(izvodBroj, pZa - 1)
        End If
        
        If InStr(Lines(i), "Izvod za datum:") > 0 Then
            izvodDatum = Trim$(ExtractAfter(Lines(i), "Izvod za datum: "))
            Dim pSt As Long
            pSt = InStr(izvodDatum, " ")
            If pSt > 0 Then izvodDatum = Left$(izvodDatum, pSt - 1)
        End If
    Next i
    
' Transaktions-Blöcke sammeln
Dim blocks As New Collection
Dim currBlock As String
Dim inTxn As Boolean

For i = LBound(Lines) To UBound(Lines)
    If IsTxnStart(Lines(i)) Then
        If Len(Trim$(currBlock)) > 0 Then
            blocks.Add currBlock
        End If
        
        currBlock = NormalizeTxnStartLine(Lines(i))
        inTxn = True
    
    ElseIf inTxn Then
        If InStr(1, Lines(i), "Ukupno za racun", vbTextCompare) > 0 Or _
           InStr(1, Lines(i), "Ukupno za racun", vbTextCompare) > 0 Then
            If Len(Trim$(currBlock)) > 0 Then
                blocks.Add currBlock
            End If
            
            currBlock = ""
            inTxn = False
        
        ElseIf Trim$(Lines(i)) <> "" Then
            currBlock = currBlock & vbLf & Trim$(Lines(i))
        End If
    End If
Next i
    
    If Len(Trim$(currBlock)) > 0 Then blocks.Add currBlock
    
    If blocks.count = 0 Then
        ParseBankaIzvod = Empty
        Exit Function
    End If
    
    ' Blöcke parsen
    Dim result() As Variant
    ReDim result(1 To blocks.count, 1 To 10)
    
    Dim b As Long
    For b = 1 To blocks.count
        Dim txn As Variant
        txn = ParseTxnBlock(blocks(b))
        
        result(b, 1) = izvodDatum
        result(b, 2) = txn(0)
        result(b, 3) = txn(1)
        result(b, 4) = txn(2)
        result(b, 5) = txn(3)
        result(b, 6) = txn(4)
        result(b, 7) = txn(5)
        result(b, 8) = txn(6)
        result(b, 9) = txn(7)
        result(b, 10) = txn(8)
    Next b
    
    ParseBankaIzvod = result
End Function

Private Function IsTxnStart(ByVal s As String) As Boolean
    Dim re As Object
    
    s = Trim$(s)
    If s = "" Then Exit Function
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "^\d{1,3}(\s+.*)?$"
    
    If re.Test(s) Then
        If Not IsDateLine(s) Then
            IsTxnStart = True
        End If
    End If
End Function

Private Function ParseTxnBlock(ByVal blockText As String) As Variant
    ' Returns:
    ' Array(DatumIzvrsenja, Partner, Racun, Zaduzenje, Odobrenje, Sifra, Svrha, PozivNaBroj, Referenca)
    
    Dim Lines() As String
    Lines = Split(blockText, vbLf)
    
    Dim racun As String
    racun = FindAccountInBlock(Lines)
    
    ' Phase 1: Datums-Zeilen finden
    Dim datumLines() As Long
    ReDim datumLines(0)
    
    Dim datumCount As Long
    Dim i As Long
    
    For i = 1 To UBound(Lines) ' Zeile 0 = Redni broj
        If IsDateLine(Trim$(Lines(i))) Then
            datumCount = datumCount + 1
            ReDim Preserve datumLines(datumCount)
            datumLines(datumCount) = i
        End If
    Next i
    
    Dim datumIzvrsenja As String
    If datumCount >= 1 Then datumIzvrsenja = Trim$(Lines(datumLines(1)))
    
    ' Phase 2: Partner = alles zwischen Redni broj und erstem Datum
    Dim partner As String
    Dim firstDatumLine As Long
    If datumCount >= 1 Then
        firstDatumLine = datumLines(1)
    Else
        firstDatumLine = UBound(Lines)
    End If
    
    Dim partnerEnd As Long
    partnerEnd = firstDatumLine - 1
    
    Dim porekloStart As Long
    porekloStart = partnerEnd
    
    For i = partnerEnd To 2 Step -1
        Dim testLine As String
        testLine = Trim$(Lines(i))
        
        If InStr(1, testLine, "EKSPOZITURA", vbTextCompare) > 0 Or _
           InStr(1, testLine, "CENTRALA", vbTextCompare) > 0 Or _
           (Len(testLine) >= 4 And IsNumeric(Left$(testLine, 4))) Then
            porekloStart = i
        Else
            Exit For
        End If
    Next i
    
    For i = 1 To porekloStart - 1
        Dim ln As String
        ln = Trim$(Lines(i))
        
        If Not IsAccountLine(ln) Then
            If partner <> "" Then partner = partner & " "
            partner = partner & ln
        End If
    Next i
    
    partner = NormalizeSpaces(partner)
    
    ' Phase 3: Beträge + Sifra
    Dim zaduzenje As Double, odobrenje As Double, sifra As String
    Dim naknada As Double
    Dim amountRest As String
    Dim amountLineStart As Long
    
    If datumCount >= 2 Then
        amountLineStart = datumLines(2) + 1
    ElseIf datumCount >= 1 Then
        amountLineStart = datumLines(1) + 1
    Else
        amountLineStart = UBound(Lines)
    End If
    
    If amountLineStart <= UBound(Lines) Then
        Dim zaduzenjeStr As String
        zaduzenjeStr = Trim$(Lines(amountLineStart))
        If IsAmount(zaduzenjeStr) Then
            zaduzenje = ToNumber(zaduzenjeStr)
        End If
    End If
    
    If amountLineStart + 1 <= UBound(Lines) Then
        Dim amtLine As String
        amtLine = Trim$(Lines(amountLineStart + 1))
        ParseAmountLine amtLine, naknada, odobrenje, sifra, amountRest
    End If
    
' Phase 4: Svrha + Referenca
Dim svrha As String, referenca As String
Dim pozivNaBroj As String
Dim svrhaStart As Long

If Trim$(amountRest) <> "" Then
    svrha = Trim$(amountRest)
End If

svrhaStart = amountLineStart + 2

For i = svrhaStart To UBound(Lines)
    ln = Trim$(Lines(i))
    If ln = "" Then GoTo NextSvrha
    
    If Not IsDateOnlyText(ln) Then
        If svrha <> "" Then svrha = svrha & " "
        svrha = svrha & ln
    End If
    
NextSvrha:
Next i

svrha = NormalizeSpaces(svrha)

ExtractReferenceFromSvrha svrha, referenca
ExtractPozivNaBroj svrha, pozivNaBroj
svrha = CleanSvrha(svrha)

ParseTxnBlock = Array(datumIzvrsenja, partner, racun, zaduzenje, _
                      odobrenje, sifra, svrha, pozivNaBroj, referenca)
End Function

Private Function IsDateLine(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 10 Then
        IsDateLine = (s Like "##.##.####")
    End If
End Function

Private Function IsDateOnlyText(ByVal s As String) As Boolean
    s = Trim$(s)
    If s = "" Then Exit Function
    
    If Right$(s, 1) = "." Then
        s = Left$(s, Len(s) - 1)
    End If
    
    IsDateOnlyText = (s Like "##.##.####")
End Function

Private Function IsAmount(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.pattern = "^\d+(,\d{3})*\.\d{2}$"
    IsAmount = re.Test(Trim$(s))
End Function

Private Sub ParseAmountLine(ByVal s As String, ByRef naknada As Double, _
                            ByRef odobrenje As Double, ByRef sifra As String, _
                            ByRef restText As String)
    ' Beispiel:
    ' "55.00 0.00 254 PID-3-2022 [97]0591000000037327647 087000494763671"
    ' "100.00 0.00 253 Uplata javnih prihoda izuzev poreza i"
    
    Dim parts() As String
    Dim i As Long
    Dim tok As String
    Dim countFound As Long
    Dim posAfterSifra As Long
    
    parts = Split(Trim$(s), " ")
    
    For i = LBound(parts) To UBound(parts)
        tok = Trim$(parts(i))
        If tok <> "" Then
            countFound = countFound + 1
            
            Select Case countFound
                Case 1
                    If IsAmount(tok) Then naknada = ToNumber(tok)
                Case 2
                    If IsAmount(tok) Then odobrenje = ToNumber(tok)
                Case 3
                    If Len(tok) = 3 And IsNumeric(tok) Then
                        sifra = tok
                        posAfterSifra = InStr(1, s, sifra, vbTextCompare)
                        If posAfterSifra > 0 Then
                            restText = Trim$(Mid$(s, posAfterSifra + Len(sifra)))
                        End If
                        Exit For
                    End If
            End Select
        End If
    Next i
End Sub

Private Sub ExtractReferenceFromSvrha(ByRef svrha As String, ByRef referenca As String)
    Dim parts() As String
    Dim lastTok As String
    Dim i As Long
    Dim rebuilt As String
    
    If Trim$(svrha) = "" Then Exit Sub
    
    parts = Split(Trim$(svrha), " ")
    If UBound(parts) < 0 Then Exit Sub
    
    lastTok = Trim$(parts(UBound(parts)))
    
    If IsReference(lastTok) Then
        referenca = lastTok
        
        For i = LBound(parts) To UBound(parts) - 1
            If Trim$(parts(i)) <> "" Then
                If rebuilt <> "" Then rebuilt = rebuilt & " "
                rebuilt = rebuilt & parts(i)
            End If
        Next i
        
        svrha = rebuilt
    End If
End Sub

Private Sub ExtractPozivNaBroj(ByRef svrha As String, ByRef pozivNaBroj As String)
    Dim p As Long
    Dim i As Long
    Dim ch As String
    Dim digits As String
    Dim fullMatch As String
    
    svrha = NormalizeSpaces(svrha)
    
    ' Suche nach [NN]
    p = InStr(1, svrha, "[")
    Do While p > 0
        If p + 3 <= Len(svrha) Then
            If Mid$(svrha, p, 1) = "[" _
               And IsNumeric(Mid$(svrha, p + 1, 2)) _
               And Mid$(svrha, p + 3, 1) = "]" Then
                
                ' Ziffern nach ] einsammeln, Whitespaces ignorieren
                digits = ""
                For i = p + 4 To Len(svrha)
                    ch = Mid$(svrha, i, 1)
                    
                    If ch Like "[0-9]" Then
                        digits = digits & ch
                    ElseIf ch = " " Or ch = vbTab Then
                        ' ignorieren
                    Else
                        Exit For
                    End If
                Next i
                
                If Len(digits) >= 6 Then
                    fullMatch = Mid$(svrha, p, 4) & digits
                    pozivNaBroj = fullMatch
                    
                    ' Originaltextbereich entfernen:
                    ' von [ bis Ende der gefundenen Ziffern/Spaces
                    Dim j As Long
                    j = p + 4
                    Do While j <= Len(svrha)
                        ch = Mid$(svrha, j, 1)
                        If (ch Like "[0-9]") Or ch = " " Or ch = vbTab Then
                            j = j + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    svrha = Left$(svrha, p - 1) & Mid$(svrha, j)
                    svrha = NormalizeSpaces(svrha)
                    Exit Sub
                End If
            End If
        End If
        
        p = InStr(p + 1, svrha, "[")
    Loop
End Sub
Private Function CleanSvrha(ByVal s As String) As String
    s = NormalizeSpaces(s)
    
    If Right$(s, 4) = "[97]" Then
        s = Trim$(Left$(s, Len(s) - 4))
    End If
    
    CleanSvrha = s
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    s = Trim$(s)
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeSpaces = s
End Function

Private Function IsReference(ByVal s As String) As Boolean
    Dim clean As String
    clean = Replace(Trim$(s), " ", "")
    
    If Len(clean) >= 12 And IsNumeric(clean) Then
        IsReference = True
    End If
End Function

Private Function ExtractLongNumber(ByVal s As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Global = False
    re.pattern = "(\d{12,})"
    
    If re.Test(s) Then
        Set matches = re.Execute(s)
        ExtractLongNumber = matches(0).SubMatches(0)
    End If
End Function

Private Function ToNumber(ByVal s As String) As Double
    s = Trim$(s)
    If s = "" Then Exit Function
    
    s = Replace(s, ",", "")
    ToNumber = val(s)
End Function

Private Function ExtractAfter(ByVal s As String, ByVal marker As String) As String
    Dim p As Long
    p = InStr(1, s, marker, vbTextCompare)
    If p > 0 Then
        ExtractAfter = Mid$(s, p + Len(marker))
    End If
End Function

Private Function IsAccountLine(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    s = Trim$(s)
    re.pattern = "^\d{3}-\d{5,20}-\d{2}$"
    re.Global = False
    
    IsAccountLine = re.Test(s)
End Function

Private Function ExtractAccountFromText(ByVal s As String) As String
    Dim re As Object, matches As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Global = False
    re.pattern = "(\d{3}-\d{5,20}-\d{2})"
    
    If re.Test(s) Then
        Set matches = re.Execute(s)
        ExtractAccountFromText = matches(0).SubMatches(0)
    End If
End Function

Private Function FindAccountInBlock(ByRef Lines() As String) As String
    Dim i As Long
    Dim s As String
    
    For i = LBound(Lines) To UBound(Lines)
        s = ExtractAccountFromText(Trim$(Lines(i)))
        If s <> "" Then
            FindAccountInBlock = s
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeTxnStartLine(ByVal s As String) As String
    Dim re As Object
    Dim matches As Object
    Dim rest As String
    
    s = Trim$(s)
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "^(\d{1,3})\s*(.*)$"
    
    If re.Test(s) Then
        Set matches = re.Execute(s)
        rest = Trim$(matches(0).SubMatches(1))
        
        NormalizeTxnStartLine = matches(0).SubMatches(0)
        If rest <> "" Then
            NormalizeTxnStartLine = NormalizeTxnStartLine & vbLf & rest
        End If
    Else
        NormalizeTxnStartLine = s
    End If
End Function
Sub TestParser()
    Dim txt As String
    Dim cb As MSForms.DataObject
    Set cb = New MSForms.DataObject
    
    cb.GetFromClipboard
    txt = cb.GetText
    
    Dim result As Variant
    result = ParseBankaIzvod(txt)
    
    If IsEmpty(result) Then
        Debug.Print "Keine Transaktionen gefunden"
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To UBound(result, 1)
        Debug.Print "--- Txn " & i & " ---"
        Debug.Print "Datum Izvoda: " & result(i, 1)
        Debug.Print "Datum Izvrš: " & result(i, 2)
        Debug.Print "Partner: " & result(i, 3)
        Debug.Print "Racun: " & result(i, 4)
        Debug.Print "Zaduzenje: " & result(i, 5)
        Debug.Print "Odobrenje: " & result(i, 6)
        Debug.Print "Sifra: " & result(i, 7)
        Debug.Print "Svrha: " & result(i, 8)
        Debug.Print "Poziv na broj: " & result(i, 9)
        Debug.Print "Referenca: " & result(i, 10)
    Next i
End Sub

