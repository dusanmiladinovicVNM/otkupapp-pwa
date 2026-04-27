Attribute VB_Name = "modBankaImport"
Option Explicit
'TODO: Matematische prüfen der Korrektheit der Auszug selber und prüfung ps+saldo=konacno stanje, und falls einige Izvodi vermisst sind in getverwaisteDokumente
Public Sub ImportBankaInbox_TX()
    Dim tx As clsTransaction
    
    On Error GoTo EH
    
    EnsureFolderExists APP_BANKA_INBOX
    EnsureFolderExists APP_BANKA_PROCESSED
    EnsureFolderExists APP_BANKA_ERROR
    
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT

    ImportBankaInbox

    tx.CommitTx
    Exit Sub
    
EH:
    LogErr "ImportBankaInbox_TX"
    If Not tx Is Nothing Then
        On Error Resume Next
        tx.RollbackTx
        On Error GoTo 0
    End If
    Err.Raise Err.Number, "ImportBankaInbox_TX", Err.Description
End Sub

Public Sub ImportBankaInbox()
    Dim files As Collection
    Dim fileName As Variant
    Dim fullPath As String
    
    Set files = New Collection
    
    fileName = Dir$(APP_BANKA_INBOX & "\*.pdf")
    Do While fileName <> ""
        files.Add CStr(fileName)
        fileName = Dir$
    Loop
    
    For Each fileName In files
        fullPath = APP_BANKA_INBOX & "\" & CStr(fileName)
        ImportOnePdfIntoBankaImport fullPath
    Next fileName
End Sub

Public Sub ImportOnePdfIntoBankaImport(ByVal pdfPath As String)
    Dim txt As String
    Dim parsed As Variant
    Dim savedCount As Long
    Dim fileName As String
    
    On Error GoTo EH

    fileName = GetFileNameFromPath(pdfPath)
    txt = ExtractTextFromPdf(pdfPath)
    
    If Trim$(txt) = "" Then
        MoveFileSafe pdfPath, APP_BANKA_ERROR & "\" & fileName
        Exit Sub
    End If
    
    parsed = ParseBankaIzvodForImport(txt, fileName)
    
    If IsEmpty(parsed) Then
        MoveFileSafe pdfPath, APP_BANKA_ERROR & "\" & fileName
        Exit Sub
    End If
    
    savedCount = SaveBankaImportRows(parsed)
    
    MoveFileSafe pdfPath, APP_BANKA_PROCESSED & "\" & fileName
    Exit Sub
    
EH:
    LogErr "ImportOnePdfIntoBankaImport_TX"
    On Error Resume Next
    MoveFileSafe pdfPath, APP_BANKA_ERROR & "\" & fileName
    On Error GoTo 0
    
    Err.Raise Err.Number, "ImportOnePdfIntoBankaImport", Err.Description
End Sub

Public Function ParseBankaIzvodForImport(ByVal txt As String, ByVal sourceFile As String) As Variant
    Dim Lines() As String
    Dim txData As Variant
    Dim result() As Variant
    Dim brojIzvoda As String
    Dim datumIzvoda As String
    Dim brojRacuna As String
    Dim i As Long
    
    txt = Replace(txt, Chr$(12), vbLf)
    txt = Replace(txt, vbCr, "")
    Lines = Split(txt, vbLf)

    brojIzvoda = ExtractIzvodBrojPdfText(Lines)
    datumIzvoda = ExtractIzvodDatumPdfText(Lines)
    brojRacuna = ExtractIzvodRacunPdfText(Lines)
    
    If Trim$(brojIzvoda) = "" Then
        Err.Raise vbObjectError + 1000, "ParseBankaIzvodForImport", "Broj izvoda nije pronadjen."
    End If
    
    If Trim$(datumIzvoda) = "" Then
        Err.Raise vbObjectError + 1001, "ParseBankaIzvodForImport", "Datum izvoda nije pronadjen."
    End If
    
    If Trim$(brojRacuna) = "" Then
        Err.Raise vbObjectError + 1002, "ParseBankaIzvodForImport", "Broj racuna izvoda nije pronadjen."
    End If
    
    txData = ParseBankaIzvodPdfText(txt)
    If IsEmpty(txData) Then
        ParseBankaIzvodForImport = Empty
        Exit Function
    End If
    
    ReDim result(1 To UBound(txData, 1), 1 To 13)

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
            
            rowIdx = AppendRow(TBL_BANKA_IMPORT, rowData)
            If rowIdx > 0 Then savedCount = savedCount + 1
        End If
    Next i
    
    SaveBankaImportRows = savedCount
End Function

Public Function IsDuplicateBankaImport(ByVal brojDokumenta As String, _
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
        If Trim$(CStr(data(i, colBrojDok))) = Trim$(brojDokumenta) Then
            
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




Public Sub Test_SaveBankaImportRows()
    Dim pdfPath As String
    Dim txt As String
    Dim parsed As Variant
    Dim savedCount As Long
    
    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub
    
    txt = ExtractTextFromPdf(pdfPath)
    If Trim$(txt) = "" Then
        Debug.Print "Kein Text gelesen."
        Exit Sub
    End If
    
    parsed = ParseBankaIzvodForImport(txt, Dir$(pdfPath))
    
    If IsEmpty(parsed) Then
        Debug.Print "Keine Daten für Import."
        Exit Sub
    End If
    
    savedCount = SaveBankaImportRows(parsed)
    
    Debug.Print "Gespeicherte Zeilen: " & savedCount
End Sub


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
    If IsError(v) Then
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
