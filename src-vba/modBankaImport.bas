Option Explicit

Public Sub ImportBankaInbox_TX()
    Dim tx As clsTransaction

    On Error GoTo EH

    EnsureFolderExists GetBankaInboxPath()
    EnsureFolderExists GetBankaProcessedPath()
    EnsureFolderExists GetBankaErrorPath()

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT

    ImportBankaInbox

    tx.CommitTx

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE
    
    LogErr "ImportBankaInbox_TX"
    
    If Not tx Is Nothing Then
        On Error Resume Next
        tx.RollbackTx
        On Error GoTo 0
    End If
    
    Err.Raise errNum, "ImportBankaInbox_TX", errDesc
End Sub

Public Sub ImportBankaInbox()
    Dim files As Collection
    Dim fileName As Variant
    Dim fullPath As String
    Dim inboxPath As String

    inboxPath = GetBankaInboxPath()

    Set files = New Collection

    fileName = Dir$(inboxPath & "\*.pdf")
    Do While fileName <> ""
        files.Add CStr(fileName)
        fileName = Dir$
    Loop

    For Each fileName In files
        fullPath = inboxPath & "\" & CStr(fileName)
        ImportOnePdfIntoBankaImport fullPath
    Next fileName
End Sub

Public Sub ImportOnePdfIntoBankaImport(ByVal pdfPath As String)
    Dim txt As String
    Dim Parsed As Variant
    Dim savedCount As Long
    Dim fileName As String
    
    On Error GoTo EH

    fileName = GetFileNameFromPath(pdfPath)
    txt = ExtractTextFromPdf(pdfPath)
    
    If Trim$(txt) = "" Then
        MoveFileSafe pdfPath, GetBankaErrorPath() & "\" & fileName
        Exit Sub
    End If
    
    Parsed = ParseBankaIzvodForImport(txt, fileName)
    
    If IsEmpty(Parsed) Then
        MoveFileSafe pdfPath, GetBankaErrorPath() & "\" & fileName
        Exit Sub
    End If
    
    savedCount = SaveBankaImportRows(Parsed)
    
    MoveFileSafe pdfPath, GetBankaProcessedPath() & "\" & fileName
    Exit Sub
    
EH:
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.description

    LogErr "ImportOnePdfIntoBankaImport"

    On Error Resume Next
    If Len(Trim$(fileName)) > 0 Then
        MoveFileSafe pdfPath, GetBankaErrorPath() & "\" & fileName
    End If
    On Error GoTo 0

    Err.Raise errNum, "ImportOnePdfIntoBankaImport", errDesc
End Sub

Public Function ParseBankaIzvodForImport(ByVal txt As String, ByVal sourceFile As String) As Variant
    Dim lines() As String
    Dim txData As Variant
    Dim result() As Variant
    Dim brojIzvoda As String
    Dim datumIzvoda As String
    Dim brojRacuna As String
    Dim saldo As BankIzvodSaldo
    Dim i As Long
    Dim sumUplata As Double
    Dim sumIsplata As Double
    Dim countUplata As Long
    Dim countIsplata As Long
    Dim expectedZavrsno As Double
    Dim diff As Double
    
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)

    brojIzvoda = ExtractIzvodBrojPdfText(lines)
    datumIzvoda = ExtractIzvodDatumPdfText(lines)
    brojRacuna = ExtractIzvodRacunPdfText(lines)
    
    If Trim$(brojIzvoda) = "" Then
        Err.Raise vbObjectError + 1000, "ParseBankaIzvodForImport", "Broj izvoda nije pronadjen."
    End If
    
    If Trim$(datumIzvoda) = "" Then
        Err.Raise vbObjectError + 1001, "ParseBankaIzvodForImport", "Datum izvoda nije pronadjen."
    End If
    
    If Trim$(brojRacuna) = "" Then
        Err.Raise vbObjectError + 1002, "ParseBankaIzvodForImport", "Broj racuna izvoda nije pronadjen."
    End If
    
    ' v6.18+: extract saldo block
    saldo = ExtractIzvodSaldoPdfText(lines)
    If Not saldo.Parsed Then
        Err.Raise vbObjectError + 1003, "ParseBankaIzvodForImport", _
            "STANJE blok izvoda " & brojIzvoda & " nije pronadjen ili ne sadrzi " & _
            "ocekivana saldo polja (Prethodno stanje, Duguje, Potrazuje, Novo stanje, Zaduzenje, Odobrenje)."
    End If
    
    txData = ParseBankaIzvodPdfText(txt)
    If IsEmpty(txData) Then
        ParseBankaIzvodForImport = Empty
        Exit Function
    End If
    
    ' v6.18+: integrity check (3 nivoa) PRE staging-a
    For i = 1 To UBound(txData, 1)
        Dim uplataVal As Double, isplataVal As Double
        uplataVal = CDbl(NzBIM(txData(i, 6), 0#))    ' Odobrenje = uplata
        isplataVal = CDbl(NzBIM(txData(i, 5), 0#))   ' Zaduzenje = isplata
        
        If uplataVal > 0 Then
            sumUplata = sumUplata + uplataVal
            countUplata = countUplata + 1
        End If
        If isplataVal > 0 Then
            sumIsplata = sumIsplata + isplataVal
            countIsplata = countIsplata + 1
        End If
    Next i
    
    ' Level 1: Pocetno + Uplate - Isplate == Zavrsno
    expectedZavrsno = saldo.PocetnoStanje + sumUplata - sumIsplata
    diff = Abs(expectedZavrsno - saldo.ZavrsnoStanje)
    If diff > 0.01 Then
        Err.Raise vbObjectError + 1004, "ParseBankaIzvodForImport", _
            "INTEGRITY FAIL izvod " & brojIzvoda & ": " & _
            "Pocetno=" & Format$(saldo.PocetnoStanje, "#,##0.00") & _
            " + Uplate=" & Format$(sumUplata, "#,##0.00") & _
            " - Isplate=" & Format$(sumIsplata, "#,##0.00") & _
            " = " & Format$(expectedZavrsno, "#,##0.00") & _
            " (ocekivano Novo stanje: " & Format$(saldo.ZavrsnoStanje, "#,##0.00") & "). " & _
            "Diff=" & Format$(diff, "#,##0.00")
    End If
    
    ' Level 2: parsed uplate sumiraju u banka-reported Potrazuje
    If Abs(sumUplata - saldo.UkupanPotrazuje) > 0.01 Then
        Err.Raise vbObjectError + 1005, "ParseBankaIzvodForImport", _
            "PARSER MISMATCH izvod " & brojIzvoda & ": " & _
            "Parsed uplate=" & Format$(sumUplata, "#,##0.00") & _
            ", banka reported Potrazuje=" & Format$(saldo.UkupanPotrazuje, "#,##0.00") & ". " & _
            "Parser je propustio uplatu."
    End If
    
    ' Level 3: parsed isplate sumiraju u banka-reported Duguje
    If Abs(sumIsplata - saldo.UkupanDuguje) > 0.01 Then
        Err.Raise vbObjectError + 1006, "ParseBankaIzvodForImport", _
            "PARSER MISMATCH izvod " & brojIzvoda & ": " & _
            "Parsed isplate=" & Format$(sumIsplata, "#,##0.00") & _
            ", banka reported Duguje=" & Format$(saldo.UkupanDuguje, "#,##0.00") & ". " & _
            "Parser je propustio isplatu."
    End If
    
    ' Level 4: broj naloga match (jeftino, dodatni signal)
    If countUplata <> saldo.BrojNalogaOdobrenje Then
        Err.Raise vbObjectError + 1007, "ParseBankaIzvodForImport", _
            "PARSER COUNT MISMATCH izvod " & brojIzvoda & ": " & _
            "Parsed uplata=" & countUplata & _
            ", banka reported Odobrenje=" & saldo.BrojNalogaOdobrenje
    End If
    
    If countIsplata <> saldo.BrojNalogaZaduzenje Then
        Err.Raise vbObjectError + 1008, "ParseBankaIzvodForImport", _
            "PARSER COUNT MISMATCH izvod " & brojIzvoda & ": " & _
            "Parsed isplata=" & countIsplata & _
            ", banka reported Zaduzenje=" & saldo.BrojNalogaZaduzenje
    End If
    
    ' v6.18+: result shape 13 -> 17 kolona (4 nova saldo polja)
    ReDim result(1 To UBound(txData, 1), 1 To 17)

    For i = 1 To UBound(txData, 1)
        result(i, 1) = brojIzvoda          ' BrojDokumenta
        result(i, 2) = datumIzvoda         ' DatumIzvoda
        result(i, 3) = brojRacuna          ' BrojRacuna
        result(i, 4) = txData(i, 2)        ' DatumTransakcije
        result(i, 5) = txData(i, 3)        ' Partner
        result(i, 6) = txData(i, 4)        ' PartnerKonto
        result(i, 7) = txData(i, 6)        ' Uplata / Odobrenje
        result(i, 8) = txData(i, 5)        ' Isplata / Zaduzenje
        result(i, 9) = txData(i, 7)        ' Sifra
        result(i, 10) = txData(i, 8)       ' SvrhaPlacanja
        result(i, 11) = txData(i, 9)       ' PozivNaBroj
        result(i, 12) = txData(i, 10)      ' BankaReferenz
        result(i, 13) = sourceFile         ' IzvorFajl
        result(i, 14) = saldo.PocetnoStanje      ' v6.18+ PocetnoStanje
        result(i, 15) = saldo.ZavrsnoStanje      ' v6.18+ ZavrsnoStanje
        result(i, 16) = saldo.UkupanDuguje       ' v6.18+ UkupanDuguje
        result(i, 17) = saldo.UkupanPotrazuje    ' v6.18+ UkupanPotrazuje
    Next i

    ParseBankaIzvodForImport = result
End Function

Public Function SaveBankaImportRows(ByRef data As Variant) As Long
    Dim colID As Long
    Dim colBrojDok As Long
    Dim colDatumIzvoda As Long
    Dim colBrojRacuna As Long
    Dim colDatumTx As Long
    Dim colPartner As Long
    Dim colPartnerKonto As Long
    Dim colOpis As Long
    Dim colUplata As Long
    Dim colIsplata As Long
    Dim colValuta As Long
    Dim colPozivNaBroj As Long
    Dim colSvrha As Long
    Dim colRef As Long
    Dim colIzvorFajl As Long
    Dim colImportVreme As Long
    Dim colObradjeno As Long
    Dim colStornirano As Long
    ' v6.18+ saldo kolone
    Dim colPocetnoStanje As Long
    Dim colZavrsnoStanje As Long
    Dim colUkupanDuguje As Long
    Dim colUkupanPotrazuje As Long
    
    Dim rowData() As Variant
    Dim colCount As Long
    Dim i As Long
    Dim rowIdx As Long
    Dim savedCount As Long
    Dim newID As String
    
    If IsEmpty(data) Then Exit Function
    If Not IsArray(data) Then Exit Function
    If UBound(data, 1) < 1 Then Exit Function
    
    colID = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID)
    colBrojDok = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA)
    colDatumIzvoda = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_IZVODA)
    colBrojRacuna = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_RACUNA)
    colDatumTx = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE)
    colPartner = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER)
    colPartnerKonto = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER_KONTO)
    colOpis = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OPIS)
    colUplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
    colIsplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)
    colValuta = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_VALUTA)
    colPozivNaBroj = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POZIV_NA_BROJ)
    colSvrha = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_SVRHA_PLACANJA)
    colRef = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BANKA_REFERENZ)
    colIzvorFajl = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_IZVOR_FAJL)
    colImportVreme = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_IMPORT_VREME)
    colObradjeno = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO)
    colStornirano = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_STORNIRANO)
    ' v6.18+ saldo lookups
    colPocetnoStanje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POCETNO_STANJE)
    colZavrsnoStanje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ZAVRSNO_STANJE)
    colUkupanDuguje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_DUGUJE)
    colUkupanPotrazuje = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_POTRAZUJE)
    
    colCount = GetTable(TBL_BANKA_IMPORT).ListColumns.count
    
    For i = 1 To UBound(data, 1)
        If Not IsDuplicateBankaImport( _
            CStr(data(i, 1)), _
            data(i, 4), _
            CDbl(NzBIM(data(i, 7), 0#)), _
            CDbl(NzBIM(data(i, 8), 0#)), _
            CStr(data(i, 5)), _
            CStr(data(i, 12)) _
        ) Then
            
            newID = GetNextID(TBL_BANKA_IMPORT, COL_BIM_ID, PREFIX_BANKA_IMPORT)
            
            ReDim rowData(1 To colCount)
            
            rowData(colID) = newID
            rowData(colBrojDok) = CStr(data(i, 1))
            rowData(colDatumIzvoda) = CStr(data(i, 2))
            rowData(colBrojRacuna) = CStr(data(i, 3))
            rowData(colDatumTx) = CStr(data(i, 4))
            rowData(colPartner) = CStr(data(i, 5))
            rowData(colPartnerKonto) = CStr(data(i, 6))
            rowData(colOpis) = CStr(data(i, 10))
            rowData(colUplata) = CDbl(NzBIM(data(i, 7), 0#))
            rowData(colIsplata) = CDbl(NzBIM(data(i, 8), 0#))
            rowData(colValuta) = "RSD"
            rowData(colPozivNaBroj) = CStr(data(i, 11))
            rowData(colSvrha) = CStr(data(i, 10))
            rowData(colRef) = CStr(data(i, 12))
            rowData(colIzvorFajl) = CStr(data(i, 13))
            rowData(colImportVreme) = Now
            rowData(colObradjeno) = vbNullString
            rowData(colStornirano) = vbNullString
            ' v6.18+ saldo metadata (copied to every row from same izvod)
            rowData(colPocetnoStanje) = CDbl(NzBIM(data(i, 14), 0#))
            rowData(colZavrsnoStanje) = CDbl(NzBIM(data(i, 15), 0#))
            rowData(colUkupanDuguje) = CDbl(NzBIM(data(i, 16), 0#))
            rowData(colUkupanPotrazuje) = CDbl(NzBIM(data(i, 17), 0#))
            
            rowIdx = AppendRow(TBL_BANKA_IMPORT, rowData)
            If rowIdx > 0 Then savedCount = savedCount + 1
        End If
    Next i
    
    SaveBankaImportRows = savedCount
End Function

Public Function IsDuplicateBankaImport(ByVal BrojDokumenta As String, _
                                       ByVal datumTransakcije As Variant, _
                                       ByVal uplata As Double, _
                                       ByVal isplata As Double, _
                                       ByVal partner As String, _
                                       ByVal bankaReferenz As String) As Boolean
    Dim data As Variant
    Dim i As Long
    
    Dim colBrojDok As Long
    Dim colDatumTx As Long
    Dim colUplata As Long
    Dim colIsplata As Long
    Dim colPartner As Long
    Dim colRef As Long
    
    data = GetTableData(TBL_BANKA_IMPORT)
    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)
    
    If IsEmpty(data) Then Exit Function
    
    colBrojDok = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA)
    colDatumTx = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE)
    colUplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA)
    colIsplata = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA)
    colPartner = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER)
    colRef = GetColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BANKA_REFERENZ)
    
    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colBrojDok))) = Trim$(BrojDokumenta) Then
            
            If Trim$(bankaReferenz) <> "" Then
                If Trim$(CStr(data(i, colRef))) = Trim$(bankaReferenz) Then
                    IsDuplicateBankaImport = True
                    Exit Function
                End If
            Else
                If Trim$(CStr(data(i, colDatumTx))) = Trim$(CStr(datumTransakcije)) _
                   And CDbl(NzBIM(data(i, colUplata), 0#)) = uplata _
                   And CDbl(NzBIM(data(i, colIsplata), 0#)) = isplata _
                   And Trim$(CStr(data(i, colPartner))) = Trim$(partner) Then
                    IsDuplicateBankaImport = True
                    Exit Function
                End If
            End If
            
        End If
    Next i
End Function

Public Sub Test_SaldoIntegrityOnSamplePDF()
    ' Pokreni rucno na sample PDF-u (21.09.2021 Komercijalna Banka).
    ' Ocekivano: saldo polja parsiraju (1775.16, 5230.00, 6000.00, 2545.16, 3, 1),
    ' integrity check prolazi.
    
    Dim pdfPath As String
    Dim txt As String
    Dim lines() As String
    Dim saldo As BankIzvodSaldo
    Dim Parsed As Variant
    Dim ok As Boolean
    
    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub
    
    txt = ExtractTextFromPdf(pdfPath)
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)
    
    ' Test 1: saldo extraction direktno
    saldo = ExtractIzvodSaldoPdfText(lines)
    Debug.Print "--- Saldo Extraction Test ---"
    Debug.Print "Parsed: " & saldo.Parsed
    Debug.Print "PocetnoStanje: " & saldo.PocetnoStanje
    Debug.Print "UkupanDuguje: " & saldo.UkupanDuguje
    Debug.Print "UkupanPotrazuje: " & saldo.UkupanPotrazuje
    Debug.Print "ZavrsnoStanje: " & saldo.ZavrsnoStanje
    Debug.Print "BrojNalogaZaduzenje: " & saldo.BrojNalogaZaduzenje
    Debug.Print "BrojNalogaOdobrenje: " & saldo.BrojNalogaOdobrenje
    Debug.Print "Math check: " & saldo.PocetnoStanje & " + " & saldo.UkupanPotrazuje & _
                " - " & saldo.UkupanDuguje & " = " & _
                (saldo.PocetnoStanje + saldo.UkupanPotrazuje - saldo.UkupanDuguje) & _
                " (expected " & saldo.ZavrsnoStanje & ")"
    
    ' Test 2: full parse + integrity check
    Debug.Print ""
    Debug.Print "--- Full Parse Test ---"
    On Error Resume Next
    Parsed = ParseBankaIzvodForImport(txt, GetFileNameFromPath2(pdfPath))
    If Err.Number <> 0 Then
        Debug.Print "FAIL: " & Err.Number & " - " & Err.description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If IsEmpty(Parsed) Then
        Debug.Print "FAIL: parse returned Empty"
        Exit Sub
    End If
    
    Debug.Print "OK: parsed " & UBound(Parsed, 1) & " transakcija sa saldo metadata"
    Debug.Print "First row saldo: Pocetno=" & Parsed(1, 14) & _
                " Zavrsno=" & Parsed(1, 15) & _
                " Duguje=" & Parsed(1, 16) & _
                " Potrazuje=" & Parsed(1, 17)
End Sub

' helper jer je GetFileNameFromPath u modBankaImport private — kopija
Private Function GetFileNameFromPath2(ByVal filePath As String) As String
    Dim p As Long
    p = InStrRev(filePath, "\")
    If p > 0 Then
        GetFileNameFromPath2 = Mid$(filePath, p + 1)
    Else
        GetFileNameFromPath2 = filePath
    End If
End Function

'HELPERS

Private Function GetFileNameFromPath(ByVal filePath As String) As String
    Dim p As Long
    
    p = InStrRev(filePath, "\")
    If p > 0 Then
        GetFileNameFromPath = Mid$(filePath, p + 1)
    Else
        GetFileNameFromPath = filePath
    End If
End Function

Private Function NzBIM(ByVal v As Variant, Optional ByVal Fallback As Variant = "") As Variant
    If isError(v) Then
        NzBIM = Fallback
    ElseIf IsNull(v) Then
        NzBIM = Fallback
    ElseIf IsEmpty(v) Then
        NzBIM = Fallback
    ElseIf Trim$(CStr(v)) = "" Then
        NzBIM = Fallback
    Else
        NzBIM = v
    End If
End Function

Private Sub ClearRowBuffer(ByRef rowData() As Variant)
    Dim j As Long
    For j = LBound(rowData, 2) To UBound(rowData, 2)
        rowData(1, j) = vbNullString
    Next j
End Sub

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Dir$(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

Private Sub MoveFileSafe(ByVal sourcePath As String, ByVal targetPath As String)
    Dim finalTarget As String
    
    finalTarget = GetUniqueTargetPath(targetPath)
    Name sourcePath As finalTarget
End Sub

Private Function GetUniqueTargetPath(ByVal targetPath As String) As String
    Dim baseName As String
    Dim ext As String
    Dim folderPath As String
    Dim p As Long
    Dim n As Long
    Dim candidate As String
    
    If Dir$(targetPath) = "" Then
        GetUniqueTargetPath = targetPath
        Exit Function
    End If
    
    p = InStrRev(targetPath, "\")
    folderPath = Left$(targetPath, p - 1)
    
    baseName = Mid$(targetPath, p + 1)
    p = InStrRev(baseName, ".")
    
    If p > 0 Then
        ext = Mid$(baseName, p)
        baseName = Left$(baseName, p - 1)
    Else
        ext = ""
    End If
    
    n = 1
    Do
        candidate = folderPath & "\" & baseName & "_" & Format$(n, "000") & ext
        If Dir$(candidate) = "" Then
            GetUniqueTargetPath = candidate
            Exit Function
        End If
        n = n + 1
    Loop
End Function

Private Function GetBankaInboxPath() As String
    GetBankaInboxPath = GetLocalConfigValue("BANKA_INBOX_PATH", APP_BANKA_INBOX)
End Function

Private Function GetBankaProcessedPath() As String
    GetBankaProcessedPath = GetLocalConfigValue("BANKA_PROCESSED_PATH", APP_BANKA_PROCESSED)
End Function

Private Function GetBankaErrorPath() As String
    GetBankaErrorPath = GetLocalConfigValue("BANKA_ERROR_PATH", APP_BANKA_ERROR)
End Function


Public Sub Diag_DumpPdfTextAroundStanje()
    Dim pdfPath As String
    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim anchorIdx As Long
    Dim lineNorm As String
    
    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub
    
    txt = ExtractTextFromPdf(pdfPath)
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    lines = Split(txt, vbLf)
    
    Debug.Print "=== TOTAL LINES: " & (UBound(lines) - LBound(lines) + 1) & " ==="
    Debug.Print ""
    
    ' Probaj naci STANJE i Prethodno stanje
    anchorIdx = -1
    For i = LBound(lines) To UBound(lines)
        lineNorm = Trim$(lines(i))
        If InStr(1, lineNorm, "STANJE", vbTextCompare) > 0 _
           Or InStr(1, lineNorm, "Prethodno stanje", vbTextCompare) > 0 _
           Or InStr(1, lineNorm, "Novo stanje", vbTextCompare) > 0 _
           Or InStr(1, lineNorm, "Prethodni saldo", vbTextCompare) > 0 _
           Or InStr(1, lineNorm, "Novi saldo", vbTextCompare) > 0 Then
            Debug.Print "[ANCHOR HIT @ line " & i & "]: " & lineNorm
            If anchorIdx < 0 Then anchorIdx = i
        End If
    Next i
    
    If anchorIdx < 0 Then
        Debug.Print ""
        Debug.Print "*** NIJE NADJEN NIJEDAN ANCHOR ***"
        Debug.Print "Prva 60 linija fajla:"
        Dim maxL As Long
        maxL = 60
        If UBound(lines) < maxL Then maxL = UBound(lines)
        For i = LBound(lines) To maxL
            Debug.Print "[" & i & "]: " & lines(i)
        Next i
        Exit Sub
    End If
    
    Debug.Print ""
    Debug.Print "=== KONTEKST OKO PRVOG ANCHOR-A (line " & anchorIdx & ") ==="
    Debug.Print "[Lines " & Max2(0, anchorIdx - 5) & " do " & Min2(UBound(lines), anchorIdx + 25) & "]"
    Debug.Print ""
    
    Dim startL As Long, endL As Long
    startL = anchorIdx - 5
    If startL < LBound(lines) Then startL = LBound(lines)
    endL = anchorIdx + 25
    If endL > UBound(lines) Then endL = UBound(lines)
    
    For i = startL To endL
        Debug.Print "[" & i & "]: " & lines(i)
    Next i
End Sub

Private Function Max2(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then Max2 = a Else Max2 = b
End Function

Private Function Min2(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min2 = a Else Min2 = b
End Function

