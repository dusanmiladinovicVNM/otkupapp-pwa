Attribute VB_Name = "modBankaImportParserPdfToText"
Option Explicit

' ============================================================
' modBankaImport_PdfText
' Parser fuer pdftotext-Ausgabe von Komercijalna Banka Izvod
'
' Output-Spalten:
' 1  Datum Izvoda
' 2  Datum Izvrs
' 3  Partner
' 4  Racun
' 5  Zaduzenje
' 6  Odobrenje
' 7  Sifra
' 8  Svrha
' 9  Poziv na broj
' 10 Referenca
' ============================================================

Public Type BankIzvodSaldo
    PocetnoStanje As Double           ' "Prethodno stanje" iz STANJE bloka
    UkupanDuguje As Double            ' "Duguje" (suma isplata reported)
    UkupanPotrazuje As Double         ' "Potrazivuje" (suma uplata reported)
    ZavrsnoStanje As Double           ' "Novo stanje"
    BrojNalogaZaduzenje As Long       ' Broj naloga - Zaduzenje
    BrojNalogaOdobrenje As Long       ' Broj naloga - Odobrenje
    parsed As Boolean                  ' True iff sva polja uspesno parsed
End Type

Public Function ExtractTextFromPdf(ByVal pdfPath As String) As String
    Const SRC As String = "ExtractTextFromPdf"

    Dim exePath As String
    Dim tempTxt As String
    Dim cmd As String
    Dim sh As Object
    Dim exitCode As Long

    On Error GoTo EH

    If Len(Trim$(pdfPath)) = 0 Then
        Err.Raise vbObjectError + 2600, SRC, "PDF path je obavezan."
    End If

    If Dir$(pdfPath) = "" Then
        Err.Raise vbObjectError + 2601, SRC, "PDF fajl ne postoji: " & pdfPath
    End If

    exePath = ResolvePdfToTextExePath()
    tempTxt = BuildUniquePdfTextTempPath(pdfPath)

    ' Defensive: never allow stale temp content from a previous failed run.
    DeleteFileIfExists tempTxt

    cmd = QuoteArg(exePath) & " -raw -nopgbrk -enc UTF-8 " & _
          QuoteArg(pdfPath) & " " & QuoteArg(tempTxt)

    Set sh = CreateObject("WScript.Shell")

    ' True = wait until finished. Return value is process exit code.
    exitCode = CLng(sh.Run(cmd, 0, True))

    If exitCode <> 0 Then
        Err.Raise vbObjectError + 2602, SRC, _
                  "pdftotext nije uspeo. ExitCode=" & CStr(exitCode) & _
                  " Pdf=" & pdfPath & _
                  " Cmd=" & cmd
    End If

    If Dir$(tempTxt) = "" Then
        Err.Raise vbObjectError + 2603, SRC, _
                  "pdftotext nije napravio temp txt fajl: " & tempTxt
    End If

    ExtractTextFromPdf = ReadAllText(tempTxt)

CleanUp:
    On Error Resume Next
    DeleteFileIfExists tempTxt
    Set sh = Nothing
    On Error GoTo 0
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    DeleteFileIfExists tempTxt
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

Private Function ResolvePdfToTextExePath() As String
    Const SRC As String = "ResolvePdfToTextExePath"

    Dim rootPath As String
    Dim defaultExePath As String
    Dim configuredPath As String

    rootPath = GetLocalConfigValue("APP_ROOT_PATH", "")

    If Len(Trim$(rootPath)) = 0 Then
        If Len(Trim$(ThisWorkbook.path)) > 0 Then
            rootPath = ThisWorkbook.path
        Else
            rootPath = "C:\OtkupApp"
        End If
    End If

    defaultExePath = rootPath & "\" & APP_PDFTOTEXT_RELATIVE_EXE_PATH

    configuredPath = GetLocalConfigValue(CONFIG_KEY_PDFTOTEXT_EXE_PATH, defaultExePath)
    configuredPath = Trim$(configuredPath)

    If Len(configuredPath) = 0 Then
        Err.Raise vbObjectError + 2610, SRC, _
                  CONFIG_KEY_PDFTOTEXT_EXE_PATH & " nije pode" & ChrW(353) & "en. " & _
                  "Podesi putanju do pdftotext.exe u tblLocalConfig."
    End If

    ' Plain executable name, e.g. pdftotext.exe, is allowed for PATH deployments.
    ' Any path-like value must exist.
    If InStr(1, configuredPath, "\", vbTextCompare) > 0 Or _
       InStr(1, configuredPath, ":", vbTextCompare) > 0 Then

        If Dir$(configuredPath) = "" Then
            Err.Raise vbObjectError + 2611, SRC, _
                      "pdftotext.exe nije pronadjen: " & configuredPath
        End If
    End If

    ResolvePdfToTextExePath = configuredPath
End Function

Private Function BuildUniquePdfTextTempPath(ByVal pdfPath As String) As String
    Const SRC As String = "BuildUniquePdfTextTempPath"

    Dim tempFolder As String
    Dim baseName As String
    Dim n As Long
    Dim candidate As String

    tempFolder = GetPdfExtractTempFolder()

    baseName = GetBaseFileNameNoExt(pdfPath)
    If Len(Trim$(baseName)) = 0 Then baseName = "pdf"

    Randomize

    Do
        n = n + 1
        candidate = tempFolder & "\pdf_extract_" & _
                    Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                    Replace(CStr(Timer), ".", "") & "_" & _
                    CStr(Int(Rnd() * 1000000)) & "_" & _
                    CStr(n) & "_" & baseName & ".txt"

        If Dir$(candidate) = "" Then
            BuildUniquePdfTextTempPath = candidate
            Exit Function
        End If
    Loop While n < 20

    Err.Raise vbObjectError + 2613, SRC, _
              "Nije moguce napraviti unique temp txt path."
End Function

Private Function GetBaseFileNameNoExt(ByVal filePath As String) As String
    Dim p As Long
    Dim fileName As String

    fileName = Trim$(filePath)

    p = InStrRev(fileName, "\")
    If p > 0 Then fileName = Mid$(fileName, p + 1)

    p = InStrRev(fileName, "/")
    If p > 0 Then fileName = Mid$(fileName, p + 1)

    p = InStrRev(fileName, ".")
    If p > 0 Then fileName = Left$(fileName, p - 1)

    fileName = Replace(fileName, " ", "_")
    fileName = Replace(fileName, ":", "_")
    fileName = Replace(fileName, "/", "_")
    fileName = Replace(fileName, "\", "_")
    fileName = Replace(fileName, "*", "_")
    fileName = Replace(fileName, "?", "_")
    fileName = Replace(fileName, """", "_")
    fileName = Replace(fileName, "<", "_")
    fileName = Replace(fileName, ">", "_")
    fileName = Replace(fileName, "|", "_")

    GetBaseFileNameNoExt = fileName
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = """" & Replace(value, """", """""") & """"
End Function

Private Sub DeleteFileIfExists(ByVal filePath As String)
    On Error Resume Next

    If Len(Trim$(filePath)) > 0 Then
        If Dir$(filePath) <> "" Then Kill filePath
    End If

    On Error GoTo 0
End Sub

Private Function GetPdfExtractTempFolder() As String
    Const SRC As String = "GetPdfExtractTempFolder"

    Dim tempFolder As String

    tempFolder = GetLocalConfigValue("APP_TEMP_PATH", "")

    If Len(Trim$(tempFolder)) = 0 Then
        tempFolder = Environ$("TEMP")
    End If

    If Len(Trim$(tempFolder)) = 0 Then
        tempFolder = Environ$("TMP")
    End If

    If Len(Trim$(tempFolder)) = 0 Then
        Err.Raise vbObjectError + 2612, SRC, _
                  "APP_TEMP_PATH/TEMP/TMP folder nije dostupan."
    End If

    If Dir$(tempFolder, vbDirectory) = "" Then
        MkDir tempFolder
    End If

    GetPdfExtractTempFolder = tempFolder
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
    
        .title = Poruka("BANKA_LBL_PDF_AUSWAHLEN")
        .filters.Clear
        .filters.Add "PDF", "*.pdf"
        
        If .Show = -1 Then
            PickPdf = .SelectedItems(1)
        End If
        
    End With

End Function

Public Function ParseBankaIzvodPdfText(ByVal txt As String) As Variant
    Dim lines() As String
    Dim blocks As Collection
    Dim result() As Variant
    Dim izvodDatum As String
    Dim i As Long
    
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)
    
    izvodDatum = ExtractIzvodDatumPdfText(lines)
    Set blocks = CollectPdfTextTxnBlocks(lines)
    
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

Public Function ExtractIzvodDatumPdfText(ByRef lines() As String) As String
    Dim i As Long
    Dim s As String
    Dim tmp As String
    Dim p As Long
    
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
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

Private Function CollectPdfTextTxnBlocks(ByRef lines() As String) As Collection
    Dim blocks As New Collection
    Dim i As Long
    Dim s As String
    Dim currBlock As String
    Dim inTxn As Boolean
    
    For i = LBound(lines) To UBound(lines)
        s = NormalizeSpacesPdf(lines(i))
        If s = "" Then GoTo NextLine
        
        If IsPdfTextTxnStart(s) Then
            If Len(Trim$(currBlock)) > 0 Then
                blocks.Add currBlock
            End If
            
            currBlock = s
            inTxn = True
        
        ElseIf inTxn Then
            If InStr(1, s, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare) > 0 Or _
               InStr(1, s, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare) > 0 Or _
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
    Dim lines() As String
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
    
    lines = Split(blockText, vbLf)
    
    For i = 1 To UBound(lines)
        ln = NormalizeSpacesPdf(lines(i))
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
        ln = NormalizeSpacesPdf(lines(i))
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
    
    If idxZad <= UBound(lines) Then
        If IsAmountPdf(lines(idxZad)) Then
            zaduzenje = ToNumberPdf(lines(idxZad))
        End If
    End If
    
    If idxOdoSif <= UBound(lines) Then
        ParsePdfOdobrenjeSifraLineStrict NormalizeSpacesPdf(lines(idxOdoSif)), odobrenje, sifra, svrha, referenca
    End If
    
    For i = idxOdoSif + 1 To UBound(lines)
        ln = NormalizeSpacesPdf(lines(i))
        If ln = "" Then GoTo NextSvrha
        
        If IsReferencePdf(ln) Then
            If referenca = "" Then referenca = ln
            GoTo NextSvrha
        End If
        
        If IsDateLinePdf(ln) Or IsDateOnlyTextPdf(ln) Then GoTo NextSvrha
        If IsAmountPdf(ln) Then GoTo NextSvrha
        If IsAccountLinePdf(ln) Then GoTo NextSvrha
        
        If InStr(1, ln, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare) > 0 Or _
        InStr(1, ln, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare) > 0 Then
    
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
  
Private Function FindAmountSifraLineIndexPdf(ByRef lines() As String) As Long
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If IsPdfAmountSifraLine(Trim$(lines(i))) Then
            FindAmountSifraLineIndexPdf = i
            Exit Function
        End If
    Next i
End Function

Private Function FindStandaloneAmountNearAmountLinePdf(ByRef lines() As String, ByVal amountLineIdx As Long) As Long
    Dim i As Long
    
    If amountLineIdx <= 0 Then Exit Function
    
    ' zuerst rueckwaerts suchen
    For i = amountLineIdx - 1 To LBound(lines) Step -1
        If IsAmountPdf(Trim$(lines(i))) Then
            FindStandaloneAmountNearAmountLinePdf = i
            Exit Function
        End If
        
        If IsDateLinePdf(Trim$(lines(i))) Then Exit For
    Next i
    
    ' dann vorwaerts suchen
    For i = amountLineIdx + 1 To UBound(lines)
        If IsAmountPdf(Trim$(lines(i))) Then
            FindStandaloneAmountNearAmountLinePdf = i
            Exit Function
        End If
        
        If IsReferencePdf(Trim$(lines(i))) Then Exit For
    Next i
End Function

Private Function FindSecondStandaloneAmountIndexPdf(ByRef lines() As String, ByVal amountLineIdx As Long) As Long
    Dim i As Long
    Dim foundCount As Long
    
    For i = LBound(lines) To UBound(lines)
        If i <> amountLineIdx Then
            If IsAmountPdf(Trim$(lines(i))) Then
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
        
        ' Alles ab "Ukupno za ..." abschneiden, falls es in derselben Zeile haengt
        pUk = InStr(1, svrha, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare)
        If pUk = 0 Then pUk = InStr(1, svrha, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare)
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

Private Function FindAccountInLinesPdf(ByRef lines() As String) As String
    Dim i As Long
    Dim s As String
    
    For i = LBound(lines) To UBound(lines)
        s = NormalizeSpacesPdf(lines(i))
        
        If InStr(1, s, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare) > 0 Or _
           InStr(1, s, "Ukupno za ra" & ChrW(269) & "un", vbTextCompare) > 0 Then
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

Private Function FindFirstDateInLinesPdf(ByRef lines() As String) As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If IsDateLinePdf(lines(i)) Then
            FindFirstDateInLinesPdf = Trim$(lines(i))
            Exit Function
        End If
    Next i
End Function

Private Function IsReferencePdf(ByVal s As String) As Boolean
    s = Replace(Trim$(s), " ", "")
    If Len(s) >= 12 And IsNumeric(s) Then IsReferencePdf = True
End Function

Private Function FindReferenceInLinesPdf(ByRef lines() As String) As String
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If IsReferencePdf(lines(i)) Then
            FindReferenceInLinesPdf = Trim$(lines(i))
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
    
    ' Haengendes [97] entfernen
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


Public Function ExtractIzvodBrojPdfText(ByRef lines() As String) As String
    Dim i As Long
    Dim s As String
    Dim re As Object
    Dim m As Object
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.pattern = "Izvod broj\s+([0-9]+)"
    
    For i = LBound(lines) To UBound(lines)
        s = Trim$(lines(i))
        If re.Test(s) Then
            Set m = re.Execute(s)(0)
            ExtractIzvodBrojPdfText = Trim$(m.SubMatches(0))
            Exit Function
        End If
    Next i
End Function

Public Function ExtractIzvodRacunPdfText(ByRef lines() As String) As String
    Dim i As Long
    Dim s As String
    Dim acc As String
    
    For i = LBound(lines) To UBound(lines)
        s = NormalizeSpacesPdf(lines(i))
        
        ' nur Kopfbereich durchsuchen, bevor Transaktionen starten
        If IsPdfTextTxnStart(s) Then Exit For
        
        acc = ExtractAccountFromTextPdf(s)
        If acc <> "" Then
            ExtractIzvodRacunPdfText = acc
            Exit Function
        End If
    Next i
End Function


'======================================================================
' v6.18+: Statement-level saldo extraction for integrity verification
'======================================================================

'----------------------------------------------------------------------
' ExtractIzvodSaldoPdfText
'
' Vadi statement-level saldo polja iz STANJE bloka Komercijalna Banka
' PDF izvoda (pdftotext -raw -nopgbrk -enc UTF-8 output).
'
' Format koji parser ocekuje (svi labeli na zasebnim linijama,
' zatim 6 tokenova na jednoj data liniji):
'   "Prethodno stanje"
'   "Duguje Potrazivuje"      (ili "Duguje   Potrazivuje")
'   "Novo stanje"
'   "Zaduzenje Odobrenje"
'   "1,775.16 5,230.00 6,000.00 2,545.16 3 1"
'
' Lokator: "Prethodno stanje" label-a se nalazi unutar STANJE sekcije
' i pojavljuje se tacno jednom po izvodu.
'
' Pristup: lociraj "Prethodno stanje" line, pa skeniraj sledecih 6 linija
' za prvu koja matchuje "double double double double long long" pattern.
' Tolerantan na red-redosled labela jer pdftotext moze blago varirati.
'----------------------------------------------------------------------
Public Function ExtractIzvodSaldoPdfText(ByRef lines() As String) As BankIzvodSaldo
    Dim result As BankIzvodSaldo
    result.parsed = False
    
    On Error GoTo EH
    
    Dim i As Long, j As Long
    Dim lineNorm As String
    Dim anchorIdx As Long
    anchorIdx = -1
    
    ' Faza 1: nadji "Prethodno stanje" anchor
    For i = LBound(lines) To UBound(lines)
        lineNorm = NormalizeSpacesPdf(lines(i))
        If InStr(1, lineNorm, "Prethodno stanje", vbTextCompare) > 0 Then
            anchorIdx = i
            Exit For
        End If
    Next i
    
    If anchorIdx < 0 Then
        ExtractIzvodSaldoPdfText = result
        Exit Function
    End If
    
    ' Faza 2: u sledecih 8 linija od anchora trazi data liniju
    '         (4 double + 2 long)
    Dim scanEnd As Long
    scanEnd = anchorIdx + 8
    If scanEnd > UBound(lines) Then scanEnd = UBound(lines)
    
    For j = anchorIdx + 1 To scanEnd
        If TryParseSaldoDataLine(lines(j), result) Then
            result.parsed = True
            Exit For
        End If
    Next j
    
    ExtractIzvodSaldoPdfText = result
    Exit Function

EH:
    ' Defensive: parser nikad ne sme da raise-uje, samo da kaze nije parsed
    result.parsed = False
    ExtractIzvodSaldoPdfText = result
End Function

'----------------------------------------------------------------------
' TryParseSaldoDataLine
'
' Komercijalna Banka pdftotext output collapse-uje nulte vrednosti iz
' Duguje / Potrazuje kolona. Stoga data linija ima 4-6 tokena zavisno
' od broja naloga zaduzenja/odobrenja:
'
'   Z>0, O>0: [Pocetno, Duguje, Potrazuje, Novo, Z, O]   (6 tokena)
'   Z=0, O>0: [Pocetno,         Potrazuje, Novo, Z, O]   (5 tokena, Duguje=0)
'   Z>0, O=0: [Pocetno, Duguje,            Novo, Z, O]   (5 tokena, Potrazuje=0)
'   Z=0, O=0: [Pocetno,                    Novo, Z, O]   (4 tokena, oba=0)
'
' Algoritam: poslednja dva tokena su uvek integeri Z/O. Pretposlednja
' grupa double-ova ima 2-4 elementa, koji se mapiraju koristeci Z/O
' brojeve kao indikator da li je Duguje/Potrazuje collapse-ovan.
'----------------------------------------------------------------------
Private Function TryParseSaldoDataLine(ByVal raw As String, _
                                        ByRef result As BankIzvodSaldo) As Boolean
    TryParseSaldoDataLine = False
    
    Dim norm As String
    norm = NormalizeSpacesPdf(raw)
    If LenB(norm) = 0 Then Exit Function
    
    Dim parts() As String
    parts = Split(norm, " ")
    
    Dim n As Long
    n = UBound(parts) - LBound(parts) + 1
    If n < 4 Or n > 6 Then Exit Function
    
    Dim base As Long
    base = LBound(parts)
    
    ' Poslednja dva tokena: Z_count, O_count (uvek integeri)
    Dim zCount As Long, oCount As Long
    If Not modParse.TryParseLong(parts(base + n - 2), zCount) Then Exit Function
    If Not modParse.TryParseLong(parts(base + n - 1), oCount) Then Exit Function
    
    If zCount < 0 Or zCount > 9999 Then Exit Function
    If oCount < 0 Or oCount > 9999 Then Exit Function
    
    ' Broj double-ova u toj liniji = n - 2
    Dim doubleCount As Long
    doubleCount = n - 2
    
    ' Validacija konzistentnosti: broj double-ova mora odgovarati Z/O pattern-u
    Dim expectedDoubles As Long
    expectedDoubles = 2  ' uvek imamo Pocetno + Novo
    If zCount > 0 Then expectedDoubles = expectedDoubles + 1   ' Duguje
    If oCount > 0 Then expectedDoubles = expectedDoubles + 1   ' Potrazuje
    
    If doubleCount <> expectedDoubles Then Exit Function
    
    ' Parse double-ove u redosledu
    Dim d() As Double
    ReDim d(0 To doubleCount - 1)
    Dim k As Long
    For k = 0 To doubleCount - 1
        If Not modParse.TryParseDouble(parts(base + k), d(k)) Then Exit Function
    Next k
    
    ' Mapiranje zavisno od kombinacije
    ' Index u d() ide redom: pocetno je uvek d(0), novo je uvek d(doubleCount-1)
    Dim pocetno As Double, novo As Double
    Dim duguje As Double, potraz As Double
    
    pocetno = d(0)
    novo = d(doubleCount - 1)
    duguje = 0
    potraz = 0
    
    Select Case doubleCount
        Case 2
            ' [Pocetno, Novo] - oba count = 0
            ' duguje = 0, potraz = 0 vec
            
        Case 3
            ' [Pocetno, X, Novo] gde X je Duguje ako Z>0, inace Potrazuje
            If zCount > 0 And oCount = 0 Then
                duguje = d(1)
            ElseIf zCount = 0 And oCount > 0 Then
                potraz = d(1)
            Else
                ' inkonzistentno: 3 double-a ali Z i O oba > 0 ili oba = 0
                Exit Function
            End If
            
        Case 4
            ' [Pocetno, Duguje, Potrazuje, Novo] - oba count > 0
            If zCount > 0 And oCount > 0 Then
                duguje = d(1)
                potraz = d(2)
            Else
                Exit Function
            End If
            
        Case Else
            Exit Function
    End Select
    
    result.PocetnoStanje = pocetno
    result.UkupanDuguje = duguje
    result.UkupanPotrazuje = potraz
    result.ZavrsnoStanje = novo
    result.BrojNalogaZaduzenje = zCount
    result.BrojNalogaOdobrenje = oCount
    
    TryParseSaldoDataLine = True
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
        Debug.Print "Datum Izvrs: " & result(i, 2)
        Debug.Print "Partner: " & result(i, 3)
        Debug.Print "Ra" & ChrW(269) & "un: " & result(i, 4)
        Debug.Print "Zadu" & ChrW(382) & "enje: " & result(i, 5)
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
    Dim lines() As String
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
    lines = Split(txt, vbLf)
    
    brojIzvoda = ExtractIzvodBrojPdfText(lines)
    datumIzvoda = ExtractIzvodDatumPdfText(lines)
    brojRacuna = ExtractIzvodRacunPdfText(lines)
    
    Debug.Print "========================================"
    Debug.Print "PDF: " & pdfPath
    Debug.Print "Broj Izvoda: " & brojIzvoda
    Debug.Print "Datum Izvoda: " & datumIzvoda
    Debug.Print "Broj Ra" & ChrW(269) & "una: " & brojRacuna
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
        Debug.Print "Broj Ra" & ChrW(269) & "una: " & brojRacuna
        Debug.Print "Datum Izvrs: " & result(i, 2)
        Debug.Print "Partner: " & result(i, 3)
        Debug.Print "Ra" & ChrW(269) & "un: " & result(i, 4)
        Debug.Print "Zadu" & ChrW(382) & "enje: " & result(i, 5)
        Debug.Print "Odobrenje: " & result(i, 6)
        Debug.Print "Sifra: " & result(i, 7)
        Debug.Print "Svrha: " & result(i, 8)
        Debug.Print "Poziv na broj: " & result(i, 9)
        Debug.Print "Referenca: " & result(i, 10)
    Next i
End Sub
