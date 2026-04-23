Attribute VB_Name = "modBankaImportParserPdfToText"
Option Explicit

' ============================================================
' modBankaImport_PdfText
' Parser für pdftotext-Ausgabe von Komercijalna Banka Izvod
'
' Output-Spalten:
' 1  Datum Izvoda
' 2  Datum Izvrš
' 3  Partner
' 4  Racun
' 5  Zaduzenje
' 6  Odobrenje
' 7  Sifra
' 8  Svrha
' 9  Poziv na broj
' 10 Referenca
' ============================================================

Public Function ExtractTextFromPdf(ByVal pdfPath As String) As String

    Dim exePath As String
    Dim tempTxt As String
    Dim cmd As String
    Dim sh As Object
    
    exePath = "C:\Users\Dusan\Desktop\OtkupAPP\poppler-25.12.0\Library\bin\pdftotext.exe"
    
    tempTxt = Environ$("TEMP") & "\pdf_extract.txt"
    
    cmd = """" & exePath & """ -raw -nopgbrk -enc UTF-8 """ & pdfPath & """ """ & tempTxt & """"
    
    Set sh = CreateObject("WScript.Shell")
    
    ' True = warten bis fertig
    sh.Run cmd, 0, True
    
    ExtractTextFromPdf = ReadAllText(tempTxt)

End Function

Private Function ReadAllText(ByVal filePath As String) As String
    Dim stm As Object
    
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2              ' Text
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    
    ReadAllText = stm.ReadText
    
    stm.Close
    Set stm = Nothing
End Function


Function PickPdf() As String

    With Application.FileDialog(msoFileDialogFilePicker)
    
        .title = "PDF auswählen"
        .filters.Clear
        .filters.Add "PDF", "*.pdf"
        
        If .Show = -1 Then
            PickPdf = .SelectedItems(1)
        End If
        
    End With

End Function

Public Function ParseBankaIzvodPdfText(ByVal txt As String) As Variant
    Dim Lines() As String
    Dim blocks As Collection
    Dim result() As Variant
    Dim izvodDatum As String
    Dim i As Long
    
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    Lines = Split(txt, vbLf)
    
    izvodDatum = ExtractIzvodDatumPdfText(Lines)
    Set blocks = CollectPdfTextTxnBlocks(Lines)
    
    If blocks Is Nothing Then
        ParseBankaIzvodPdfText = Empty
        Exit Function
    End If
    
    If blocks.count = 0 Then
        ParseBankaIzvodPdfText = Empty
        Exit Function
    End If
    
    ReDim result(1 To blocks.count, 1 To 10)
    
    For i = 1 To blocks.count
        Dim txn As Variant
        txn = ParsePdfTextTxnBlock(CStr(blocks(i)), izvodDatum)
        
        result(i, 1) = txn(0)
        result(i, 2) = txn(1)
        result(i, 3) = txn(2)
        result(i, 4) = txn(3)
        result(i, 5) = txn(4)
        result(i, 6) = txn(5)
        result(i, 7) = txn(6)
        result(i, 8) = txn(7)
        result(i, 9) = txn(8)
        result(i, 10) = txn(9)
    Next i
    
    ParseBankaIzvodPdfText = result
End Function

Public Function ExtractIzvodDatumPdfText(ByRef Lines() As String) As String
    Dim i As Long
    Dim s As String
    Dim tmp As String
    Dim p As Long
    
    For i = LBound(Lines) To UBound(Lines)
        s = Trim$(Lines(i))
        If InStr(1, s, "Izvod za datum:", vbTextCompare) > 0 Then
            tmp = ExtractAfterPdf(s, "Izvod za datum:")
            tmp = Trim$(tmp)
            p = InStr(tmp, " ")
            If p > 0 Then tmp = Left$(tmp, p - 1)
            ExtractIzvodDatumPdfText = tmp
            Exit Function
        End If
    Next i
End Function

Private Function CollectPdfTextTxnBlocks(ByRef Lines() As String) As Collection
    Dim blocks As New Collection
    Dim i As Long
    Dim s As String
    Dim currBlock As String
    Dim inTxn As Boolean
    
    For i = LBound(Lines) To UBound(Lines)
        s = NormalizeSpacesPdf(Lines(i))
        If s = "" Then GoTo NextLine
        
        If IsPdfTextTxnStart(s) Then
            If Len(Trim$(currBlock)) > 0 Then
                blocks.Add currBlock
            End If
            
            currBlock = s
            inTxn = True
        
        ElseIf inTxn Then
            If InStr(1, s, "Ukupno za racun", vbTextCompare) > 0 Or _
               InStr(1, s, "Ukupno za racun", vbTextCompare) > 0 Or _
               InStr(1, s, "(postoji", vbTextCompare) > 0 Or _
               InStr(1, s, "Ukupno RSD", vbTextCompare) > 0 Or _
               InStr(1, s, "Iznos ukupno naplacene naknade", vbTextCompare) > 0 Or _
               InStr(1, s, "Iznos ukupno naplacene naknade", vbTextCompare) > 0 Or _
               Left$(s, 11) = "Izvod broj " Then
                
                If Len(Trim$(currBlock)) > 0 Then
                    blocks.Add currBlock
                End If
                
                currBlock = ""
                inTxn = False
            Else
                currBlock = currBlock & vbLf & s
            End If
        End If
        
NextLine:
    Next i
    
    If Len(Trim$(currBlock)) > 0 Then blocks.Add currBlock
    Set CollectPdfTextTxnBlocks = blocks
End Function
Private Function IsPdfTextTxnStart(ByVal s As String) As Boolean
    s = Trim$(s)
    
    If Len(s) <= 3 And IsNumeric(s) Then
        IsPdfTextTxnStart = True
    End If
End Function

Private Function NormalizePdfTextTxnStart(ByVal s As String) As String
    NormalizePdfTextTxnStart = Trim$(s)
End Function

Private Function ParsePdfTextTxnBlock(ByVal blockText As String, ByVal izvodDatum As String) As Variant
    Dim Lines() As String
    Dim i As Long
    Dim ln As String
    
    Dim datumIzvrsenja As String
    Dim partner As String
    Dim racun As String
    Dim zaduzenje As Double
    Dim odobrenje As Double
    Dim sifra As String
    Dim svrha As String
    Dim pozivNaBroj As String
    Dim referenca As String
    
    Dim firstDateIdx As Long
    Dim secondDateIdx As Long
    Dim idxZad As Long
    Dim idxNak As Long
    Dim idxOdoSif As Long
    
    Lines = Split(blockText, vbLf)
    
    For i = 1 To UBound(Lines)
        ln = NormalizeSpacesPdf(Lines(i))
        If IsDateLinePdf(ln) Then
            If firstDateIdx = 0 Then
                firstDateIdx = i
                datumIzvrsenja = ln
            ElseIf secondDateIdx = 0 Then
                secondDateIdx = i
                Exit For
            End If
        End If
    Next i
    
    For i = 1 To firstDateIdx - 1
        ln = NormalizeSpacesPdf(Lines(i))
        If ln = "" Then GoTo NextPartner
        
        If IsAccountLinePdf(ln) Then
            racun = ln
            GoTo NextPartner
        End If
        
        If InStr(1, ln, "CENTRALA", vbTextCompare) > 0 Then GoTo NextPartner
        If InStr(1, ln, "EKSPOZITURA", vbTextCompare) > 0 Then GoTo NextPartner
        
        If partner <> "" Then partner = partner & " "
        partner = partner & ln
        
NextPartner:
    Next i
    
    partner = NormalizeSpacesPdf(partner)
    
    idxZad = secondDateIdx + 1
    idxNak = secondDateIdx + 2
    idxOdoSif = secondDateIdx + 3
    
    If idxZad <= UBound(Lines) Then
        If IsAmountPdf(Lines(idxZad)) Then
            zaduzenje = ToNumberPdf(Lines(idxZad))
        End If
    End If
    
    If idxOdoSif <= UBound(Lines) Then
        ParsePdfOdobrenjeSifraLineStrict NormalizeSpacesPdf(Lines(idxOdoSif)), odobrenje, sifra, svrha, referenca
    End If
    
    For i = idxOdoSif + 1 To UBound(Lines)
        ln = NormalizeSpacesPdf(Lines(i))
        If ln = "" Then GoTo NextSvrha
        
        If IsReferencePdf(ln) Then
            If referenca = "" Then referenca = ln
            GoTo NextSvrha
        End If
        
        If IsDateLinePdf(ln) Or IsDateOnlyTextPdf(ln) Then GoTo NextSvrha
        If IsAmountPdf(ln) Then GoTo NextSvrha
        If IsAccountLinePdf(ln) Then GoTo NextSvrha
        
        If InStr(1, ln, "Ukupno za racun", vbTextCompare) > 0 Or _
        InStr(1, ln, "Ukupno za racun", vbTextCompare) > 0 Then
    
        ' Falls schon Text davor in derselben Zeile steht, nur den Teil vor "Ukupno..." behalten
        Dim pUk As Long
        pUk = InStr(1, ln, "Ukupno za", vbTextCompare)
    
        If pUk > 1 Then
                ln = Trim$(Left$(ln, pUk - 1))
                If ln <> "" Then
                    If svrha <> "" Then svrha = svrha & " "
                        svrha = svrha & ln
                    End If
                End If
    
            Exit For
        End If
        If InStr(1, ln, "(postoji", vbTextCompare) > 0 Then Exit For
        If InStr(1, ln, "Ukupno RSD", vbTextCompare) > 0 Then Exit For
        If InStr(1, ln, "Iznos ukupno naplacene naknade", vbTextCompare) > 0 Then Exit For
        If InStr(1, ln, "Iznos ukupno naplacene naknade", vbTextCompare) > 0 Then Exit For
        If Left$(ln, 11) = "Izvod broj " Then Exit For
        
        If svrha <> "" Then svrha = svrha & " "
        svrha = svrha & ln
        
NextSvrha:
    Next i
    
    svrha = NormalizeSpacesPdf(svrha)
    ExtractReferenceFromSvrhaPdf svrha, referenca
    ExtractPozivNaBrojPdf svrha, pozivNaBroj
    svrha = CleanSvrhaPdf(svrha)
    
    ParsePdfTextTxnBlock = Array(izvodDatum, datumIzvrsenja, partner, racun, _
                                 zaduzenje, odobrenje, sifra, svrha, pozivNaBroj, referenca)
End Function
  
Private Function FindAmountSifraLineIndexPdf(ByRef Lines() As String) As Long
    Dim i As Long
    For i = LBound(Lines) To UBound(Lines)
        If IsPdfAmountSifraLine(Trim$(Lines(i))) Then
            FindAmountSifraLineIndexPdf = i
            Exit Function
        End If
    Next i
End Function

Private Function FindStandaloneAmountNearAmountLinePdf(ByRef Lines() As String, ByVal amountLineIdx As Long) As Long
    Dim i As Long
    
    If amountLineIdx <= 0 Then Exit Function
    
    ' zuerst rückwärts suchen
    For i = amountLineIdx - 1 To LBound(Lines) Step -1
        If IsAmountPdf(Trim$(Lines(i))) Then
            FindStandaloneAmountNearAmountLinePdf = i
            Exit Function
        End If
        
        If IsDateLinePdf(Trim$(Lines(i))) Then Exit For
    Next i
    
    ' dann vorwärts suchen
    For i = amountLineIdx + 1 To UBound(Lines)
        If IsAmountPdf(Trim$(Lines(i))) Then
            FindStandaloneAmountNearAmountLinePdf = i
            Exit Function
        End If
        
        If IsReferencePdf(Trim$(Lines(i))) Then Exit For
    Next i
End Function

Private Function FindSecondStandaloneAmountIndexPdf(ByRef Lines() As String, ByVal amountLineIdx As Long) As Long
    Dim i As Long
    Dim foundCount As Long
    
    For i = LBound(Lines) To UBound(Lines)
        If i <> amountLineIdx Then
            If IsAmountPdf(Trim$(Lines(i))) Then
                foundCount = foundCount + 1
                ' erste reine Betragzeile im Block
                FindSecondStandaloneAmountIndexPdf = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IsPdfAmountSifraLine(ByVal s As String) As Boolean
    Dim re As Object
    
    s = NormalizeSpacesPdf(s)
    If s = "" Then Exit Function
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "^\d+(,\d{3})*\.\d{2}\s+\d{3}(\s+.*)?$"
    
    IsPdfAmountSifraLine = re.Test(s)
End Function


Private Sub ParsePdfAmountSifraLine(ByVal s As String, _
                                    ByRef amountFromLine As Double, _
                                    ByRef sifra As String, _
                                    ByRef svrha As String)
    Dim re As Object
    Dim m As Object
    
    s = NormalizeSpacesPdf(s)
    If s = "" Then Exit Sub
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "^(\d+(,\d{3})*\.\d{2})\s+(\d{3})(?:\s+(.*))?$"
    
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        
        amountFromLine = ToNumberPdf(m.SubMatches(0))
        sifra = m.SubMatches(2)
        svrha = NormalizeSpacesPdf(m.SubMatches(3))
    End If
End Sub

Private Sub ParsePdfOdobrenjeSifraLineStrict(ByVal s As String, _
                                             ByRef odobrenje As Double, _
                                             ByRef sifra As String, _
                                             ByRef svrha As String, _
                                             ByRef referenca As String)
    Dim re As Object
    Dim m As Object
    Dim pUk As Long
    
    s = NormalizeSpacesPdf(s)
    If s = "" Then Exit Sub
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "^(\d+(,\d{3})*\.\d{2})\s+(\d{3})(?:\s+(.*))?$"
    
    If re.Test(s) Then
        Set m = re.Execute(s)(0)
        
        odobrenje = ToNumberPdf(m.SubMatches(0))
        sifra = m.SubMatches(2)
        svrha = NormalizeSpacesPdf(m.SubMatches(3))
        
        ' Alles ab "Ukupno za ..." abschneiden, falls es in derselben Zeile hängt
        pUk = InStr(1, svrha, "Ukupno za racun", vbTextCompare)
        If pUk = 0 Then pUk = InStr(1, svrha, "Ukupno za racun", vbTextCompare)
        If pUk > 0 Then
            svrha = Trim$(Left$(svrha, pUk - 1))
        End If
        
        ' Referenz am Ende derselben Zeile mitnehmen
        ExtractReferenceFromSvrhaPdf svrha, referenca
    End If
End Sub

Private Function IsDateLinePdf(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 10 Then IsDateLinePdf = (s Like "##.##.####")
End Function

Private Function IsDateOnlyTextPdf(ByVal s As String) As Boolean
    s = Trim$(s)
    If s = "" Then Exit Function
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    IsDateOnlyTextPdf = (s Like "##.##.####")
End Function

Private Function IsAccountLinePdf(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d{3}-\d{5,20}-\d{2}$"
    IsAccountLinePdf = re.Test(Trim$(s))
End Function

Private Function FindAccountInLinesPdf(ByRef Lines() As String) As String
    Dim i As Long
    Dim s As String
    
    For i = LBound(Lines) To UBound(Lines)
        s = NormalizeSpacesPdf(Lines(i))
        
        If InStr(1, s, "Ukupno za racun", vbTextCompare) > 0 Or _
           InStr(1, s, "Ukupno za racun", vbTextCompare) > 0 Then
            Exit Function
        End If
        
        If IsAccountLinePdf(s) Then
            FindAccountInLinesPdf = s
            Exit Function
        End If
    Next i
End Function
Private Function ExtractAccountFromTextPdf(ByVal s As String) As String
    Dim re As Object
    Dim matches As Object
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "(\d{3}-\d{5,20}-\d{2})"
    
    If re.Test(s) Then
        Set matches = re.Execute(s)
        ExtractAccountFromTextPdf = matches(0).SubMatches(0)
    End If
End Function

Private Function FindFirstDateInLinesPdf(ByRef Lines() As String) As String
    Dim i As Long
    For i = LBound(Lines) To UBound(Lines)
        If IsDateLinePdf(Lines(i)) Then
            FindFirstDateInLinesPdf = Trim$(Lines(i))
            Exit Function
        End If
    Next i
End Function

Private Function IsReferencePdf(ByVal s As String) As Boolean
    s = Replace(Trim$(s), " ", "")
    If Len(s) >= 12 And IsNumeric(s) Then IsReferencePdf = True
End Function

Private Function FindReferenceInLinesPdf(ByRef Lines() As String) As String
    Dim i As Long
    For i = LBound(Lines) To UBound(Lines)
        If IsReferencePdf(Lines(i)) Then
            FindReferenceInLinesPdf = Trim$(Lines(i))
            Exit Function
        End If
    Next i
End Function

Private Function IsAmountPdf(ByVal s As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.pattern = "^\d+(,\d{3})*\.\d{2}$"
    IsAmountPdf = re.Test(Trim$(s))
End Function

Private Function ToNumberPdf(ByVal s As String) As Double
    s = Replace(Trim$(s), ",", "")
    ToNumberPdf = val(s)
End Function

Private Function NormalizeSpacesPdf(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, Chr$(160), " ")
    s = Trim$(s)
    
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    
    NormalizeSpacesPdf = s
End Function

Private Sub ExtractPozivNaBrojPdf(ByRef svrha As String, ByRef pozivNaBroj As String)
    Dim p As Long
    Dim i As Long
    Dim ch As String
    Dim digits As String
    
    svrha = NormalizeSpacesPdf(svrha)
    p = InStr(1, svrha, "[")
    
    Do While p > 0
        If p + 3 <= Len(svrha) Then
            If Mid$(svrha, p, 1) = "[" _
               And IsNumeric(Mid$(svrha, p + 1, 2)) _
               And Mid$(svrha, p + 3, 1) = "]" Then
                
                digits = ""
                For i = p + 4 To Len(svrha)
                    ch = Mid$(svrha, i, 1)
                    If ch Like "[0-9]" Then
                        digits = digits & ch
                    ElseIf ch = " " Or ch = vbTab Then
                    Else
                        Exit For
                    End If
                Next i
                
                If Len(digits) >= 6 Then
                    pozivNaBroj = Mid$(svrha, p, 4) & digits
                    svrha = Trim$(Left$(svrha, p - 1) & Mid$(svrha, i))
                    svrha = NormalizeSpacesPdf(svrha)
                    Exit Sub
                End If
            End If
        End If
        
        p = InStr(p + 1, svrha, "[")
    Loop
End Sub

Private Sub ExtractReferenceFromSvrhaPdf(ByRef svrha As String, ByRef referenca As String)
    Dim parts() As String
    Dim lastTok As String
    Dim i As Long
    Dim rebuilt As String
    
    If Trim$(svrha) = "" Then Exit Sub
    
    parts = Split(Trim$(svrha), " ")
    lastTok = Trim$(parts(UBound(parts)))
    
    If referenca = "" And IsReferencePdf(lastTok) Then referenca = lastTok
    
    If IsReferencePdf(lastTok) Then
        For i = LBound(parts) To UBound(parts) - 1
            If Trim$(parts(i)) <> "" Then
                If rebuilt <> "" Then rebuilt = rebuilt & " "
                rebuilt = rebuilt & parts(i)
            End If
        Next i
        svrha = rebuilt
    End If
End Sub

Private Function CleanSvrhaPdf(ByVal s As String) As String
    Dim parts() As String
    Dim rebuilt As String
    Dim i As Long
    Dim lastTok As String
    Dim pUk As Long
    
    s = NormalizeSpacesPdf(s)
    
    ' Hängendes [97] entfernen
    If Right$(s, 4) = "[97]" Then
        s = Trim$(Left$(s, Len(s) - 4))
    End If
    
    ' Alles ab "Ukupno za ..." abschneiden
    pUk = InStr(1, s, "Ukupno za", vbTextCompare)
    If pUk > 0 Then
        s = Trim$(Left$(s, pUk - 1))
    End If
    
    s = NormalizeSpacesPdf(s)
    If s = "" Then
        CleanSvrhaPdf = s
        Exit Function
    End If
    
    ' Falls letztes Token nur Datum ist, entfernen
    parts = Split(s, " ")
    If UBound(parts) >= 0 Then
        lastTok = Trim$(parts(UBound(parts)))
        If IsDateOnlyTextPdf(lastTok) Then
            For i = LBound(parts) To UBound(parts) - 1
                If Trim$(parts(i)) <> "" Then
                    If rebuilt <> "" Then rebuilt = rebuilt & " "
                    rebuilt = rebuilt & parts(i)
                End If
            Next i
            s = rebuilt
        End If
    End If
    
    CleanSvrhaPdf = NormalizeSpacesPdf(s)
End Function

Private Function ExtractAfterPdf(ByVal s As String, ByVal marker As String) As String
    Dim p As Long
    p = InStr(1, s, marker, vbTextCompare)
    If p > 0 Then
        ExtractAfterPdf = Mid$(s, p + Len(marker))
    End If
End Function


Public Function ExtractIzvodBrojPdfText(ByRef Lines() As String) As String
    Dim i As Long
    Dim s As String
    Dim re As Object
    Dim m As Object
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "Izvod broj\s+([0-9]+)"
    
    For i = LBound(Lines) To UBound(Lines)
        s = Trim$(Lines(i))
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            ExtractIzvodBrojPdfText = Trim$(m.SubMatches(0))
            Exit Function
        End If
    Next i
End Function

Public Function ExtractIzvodRacunPdfText(ByRef Lines() As String) As String
    Dim i As Long
    Dim s As String
    Dim acc As String
    
    For i = LBound(Lines) To UBound(Lines)
        s = NormalizeSpacesPdf(Lines(i))
        
        ' nur Kopfbereich durchsuchen, bevor Transaktionen starten
        If IsPdfTextTxnStart(s) Then Exit For
        
        acc = ExtractAccountFromTextPdf(s)
        If acc <> "" Then
            ExtractIzvodRacunPdfText = acc
            Exit Function
        End If
    Next i
End Function








Sub TestPdfTextParser()
    Dim pdfPath As String
    Dim txt As String
    Dim result As Variant
    Dim i As Long
    
    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub
    
    txt = ExtractTextFromPdf(pdfPath)
    If Trim$(txt) = "" Then
        Debug.Print "Kein Text gelesen."
        Exit Sub
    End If
    
    result = ParseBankaIzvodPdfText(txt)
    
    If IsEmpty(result) Then
        Debug.Print "Keine Transaktionen gefunden."
        Exit Sub
    End If
    
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

Sub TestPdfTextParser123()
    Dim pdfPath As String
    Dim txt As String
    Dim result As Variant
    Dim Lines() As String
    Dim brojIzvoda As String
    Dim datumIzvoda As String
    Dim brojRacuna As String
    Dim i As Long
    
    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub
    
    txt = ExtractTextFromPdf(pdfPath)
    If Trim$(txt) = "" Then
        Debug.Print "Kein Text gelesen."
        Exit Sub
    End If
    
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    Lines = Split(txt, vbLf)
    
    brojIzvoda = ExtractIzvodBrojPdfText(Lines)
    datumIzvoda = ExtractIzvodDatumPdfText(Lines)
    brojRacuna = ExtractIzvodRacunPdfText(Lines)
    
    Debug.Print "========================================"
    Debug.Print "PDF: " & pdfPath
    Debug.Print "Broj Izvoda: " & brojIzvoda
    Debug.Print "Datum Izvoda: " & datumIzvoda
    Debug.Print "Broj Racuna: " & brojRacuna
    Debug.Print "========================================"
    
    result = ParseBankaIzvodPdfText(txt)
    
    If IsEmpty(result) Then
        Debug.Print "Keine Transaktionen gefunden."
        Exit Sub
    End If
    
    For i = 1 To UBound(result, 1)
        Debug.Print "--- Txn " & i & " ---"
        Debug.Print "Broj Izvoda: " & brojIzvoda
        Debug.Print "Datum Izvoda: " & datumIzvoda
        Debug.Print "Broj Racuna: " & brojRacuna
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
