Attribute VB_Name = "modBankaImport"
Option Explicit

' ============================================================
' PATCH: Option B -- deferred file moves after batch commit
' File: src-vba/modBankaImport.bas
'
' Intent:
' - DB staging and file-system moves are separated.
' - ImportOnePdfIntoBankaImport does NOT move PDF files immediately.
' - Successful files are added to a pending move list.
' - ImportBankaInbox_TX commits the transaction first.
' - Only after successful CommitTx are successful PDFs moved to Processed.
' - On rollback/error, no successful PDFs are moved to Processed.
'
' This patch is designed to sit on top of P1-3 fail-fast SaveBankaImportRows.
' ============================================================


' ============================================================
' PATCH: P1-3 BankaImport fail-fast staging
' File: src-vba/modBankaImport.bas
'
' Goals:
' - SaveBankaImportRows uses RequireColumnIndex for all tblBankaImport columns.
' - AppendRow <= 0 is a hard failure.
' - ImportOnePdfIntoBankaImport distinguishes:
'     imported, duplicate-only, parse error, integrity error, append error.
' - PDF goes to Processed only after staging is reliable.
' - Append/schema failures bubble to ImportBankaInbox_TX rollback.
' ============================================================

' ============================================================
' 1) Add these constants near top of modBankaImport.bas
' ============================================================

Private Const BIM_STATUS_IMPORTED As String = "imported"
Private Const BIM_STATUS_DUPLICATE_ONLY As String = "duplicate-only"
Private Const BIM_STATUS_PARSE_ERROR As String = "parse error"
Private Const BIM_STATUS_INTEGRITY_ERROR As String = "integrity error"
Private Const BIM_STATUS_APPEND_ERROR As String = "append error"
Private Const BIM_STATUS_SCHEMA_ERROR As String = "schema error"
Private Const BIM_STATUS_EXTRACT_ERROR As String = "extract error"
Private Const BIM_STATUS_UNKNOWN_ERROR As String = "unknown error"

Private Const ERR_BIM_IMPORT_BASE As Long = vbObjectError + 2700
Private Const ERR_BIM_SAVE_BASE As Long = vbObjectError + 2800

Public Sub ImportBankaInbox_TX()
    Const SRC As String = "ImportBankaInbox_TX"

    Dim tx As clsTransaction
    Dim successMoves As Collection
    Dim errorMoves As Collection
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    On Error GoTo EH

    EnsureFolderExists GetBankaInboxPath()
    EnsureFolderExists GetBankaProcessedPath()
    EnsureFolderExists GetBankaErrorPath()

    Set successMoves = New Collection
    Set errorMoves = New Collection

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT

    ImportBankaInboxToPendingMoves successMoves, errorMoves

    tx.CommitTx
    Set tx = Nothing

    ExecutePendingBankaFileMoves successMoves
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr SRC

    If Not tx Is Nothing Then tx.RollbackTx

    ' Only PDF-level failures are moved to Error after rollback.
    ' Success moves are intentionally NOT executed on rollback.
    ExecutePendingBankaFileMoves errorMoves

    Debug.Print SRC & " failed. DB rolled back. Success moves NOT executed. " & _
                "Error moves executed only for PDF-level errors. " & _
                "Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc

    On Error GoTo 0

    Err.Raise errNum, SRC, errDesc
End Sub

Public Sub ImportBankaInbox()
    ImportBankaInbox_WithDrivePull
End Sub

Public Sub ImportBankaInbox_WithDrivePull()
    Const SRC As String = "ImportBankaInbox_WithDrivePull"

    On Error GoTo EH

    ' Drive povlacenje je BEST-EFFORT: ako Drive putanja nije dostupna (offline,
    ' pogresan BANKA_DRIVE_SOURCE_PATH, nepristupacan folder) NE obaraj uvoz --
    ' zabelezi WARN i uvezi lokalni Inbox svejedno. Sam uvoz (_TX) ostaje hard.
    If BankaDrivePullConfigured() Then
        On Error Resume Next
        PullBankPdfsFromDriveProduction
        If Err.Number <> 0 Then
            Err.Clear
            LogError SRC, "Drive pull preskocen -- nastavljam sa lokalnim Inboxom " & _
                          "(detalji: PullBankPdfsFromDriveProduction u dnevnom logu).", 0, "WARN"
        End If
        On Error GoTo EH
    End If

    ImportBankaInbox_TX
    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Private Sub ImportBankaInboxToPendingMoves(ByRef successMoves As Collection, _
                                           ByRef errorMoves As Collection)
    Const SRC As String = "ImportBankaInboxToPendingMoves"

    Dim files As Collection
    Dim fileName As Variant
    Dim fullPath As String
    Dim inboxPath As String
    Dim statusText As String

    On Error GoTo EH

    If successMoves Is Nothing Then Set successMoves = New Collection
    If errorMoves Is Nothing Then Set errorMoves = New Collection

    inboxPath = GetBankaInboxPath()

    Set files = New Collection

    fileName = Dir$(inboxPath & "\*.pdf")
    Do While fileName <> ""
        files.Add CStr(fileName)
        fileName = Dir$
    Loop

    For Each fileName In files
        fullPath = inboxPath & "\" & CStr(fileName)
        statusText = ImportOnePdfIntoBankaImport_Core(fullPath, successMoves, errorMoves)

        Debug.Print SRC & " staged. File=" & CStr(fileName) & _
                    " Status=" & statusText
    Next fileName

    Exit Sub
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Sub

Private Function ImportOnePdfIntoBankaImport_Core(ByVal pdfPath As String, _
                                                  ByRef successMoves As Collection, _
                                                  ByRef errorMoves As Collection) As String
    Const SRC As String = "ImportOnePdfIntoBankaImport_Core"

    Dim txt As String
    Dim parsed As Variant
    Dim savedCount As Long
    Dim fileName As String
    Dim targetPath As String

    On Error GoTo EH

    If successMoves Is Nothing Then Set successMoves = New Collection
    If errorMoves Is Nothing Then Set errorMoves = New Collection

    fileName = GetFileNameFromPath(pdfPath)

    If Len(Trim$(pdfPath)) = 0 Then
        Err.Raise ERR_BIM_IMPORT_BASE + 1, SRC, "PDF path je obavezan."
    End If

    If Dir$(pdfPath) = "" Then
        Err.Raise ERR_BIM_IMPORT_BASE + 2, SRC, "PDF fajl ne postoji: " & pdfPath
    End If

    txt = ExtractTextFromPdf(pdfPath)

    If Len(Trim$(txt)) = 0 Then
        Err.Raise ERR_BIM_IMPORT_BASE + 3, SRC, _
                  "PDF extract je vratio prazan tekst. File=" & fileName
    End If

    parsed = ParseBankaIzvodForImport(txt, fileName)
    
    If IsEmpty(parsed) Then
        Err.Raise ERR_BIM_IMPORT_BASE + 4, SRC, _
                  "Parser nije vratio nijednu transakciju. File=" & fileName
    End If

    savedCount = SaveBankaImportRows(parsed)

    If savedCount > 0 Then
        ImportOnePdfIntoBankaImport_Core = BIM_STATUS_IMPORTED
    Else
        ImportOnePdfIntoBankaImport_Core = BIM_STATUS_DUPLICATE_ONLY
    End If

    targetPath = GetBankaProcessedPath() & "\" & fileName
    AddPendingBankaFileMove successMoves, pdfPath, targetPath, ImportOnePdfIntoBankaImport_Core

    Exit Function
    
EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim errCategory As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE
    errCategory = ClassifyBankaImportError(errNum, errSrc, errDesc)

    On Error Resume Next
    LogErr SRC

    fileName = GetFileNameFromPath(pdfPath)

    If ShouldMoveBankaImportFailureToError(errCategory) Then
        If Len(Trim$(pdfPath)) > 0 And Dir$(pdfPath) <> "" Then
            AddPendingBankaFileMove errorMoves, pdfPath, _
                                    GetBankaErrorPath() & "\" & fileName, _
                                    errCategory
        End If
    End If

    Debug.Print SRC & " failed. Status=" & errCategory & _
                " Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc & _
                " File=" & fileName

    On Error GoTo 0

    Err.Raise errNum, SRC, _
              "[" & errCategory & "] Source=" & errSrc & " | " & errDesc
End Function

Private Function ShouldMoveBankaImportFailureToError(ByVal errorCategory As String) As Boolean
    Select Case LCase$(Trim$(errorCategory))
        Case BIM_STATUS_PARSE_ERROR, _
             BIM_STATUS_INTEGRITY_ERROR, _
             BIM_STATUS_EXTRACT_ERROR
            ShouldMoveBankaImportFailureToError = True

        Case Else
            ' Schema/append/unknown failures are not necessarily bad PDFs.
            ' Keep file in Inbox for operator/dev investigation.
            ShouldMoveBankaImportFailureToError = False
    End Select
End Function


Public Sub ImportOnePdfIntoBankaImport(ByVal pdfPath As String)
    Const SRC As String = "ImportOnePdfIntoBankaImport"

    Dim tx As clsTransaction
    Dim successMoves As Collection
    Dim errorMoves As Collection
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    On Error GoTo EH

    Set successMoves = New Collection
    Set errorMoves = New Collection

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_BANKA_IMPORT

    ImportOnePdfIntoBankaImport_Core pdfPath, successMoves, errorMoves

    tx.CommitTx
    Set tx = Nothing

    ExecutePendingBankaFileMoves successMoves
    Exit Sub
    
EH:
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    If Not tx Is Nothing Then tx.RollbackTx
    ExecutePendingBankaFileMoves errorMoves
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Sub

' Multi-bank dispatch: prepoznaj banku iz pdftotext teksta.
' Default (nema fingerprint-a) = "KOMERC" (postojeci Komercijalna parser).
Public Function DetectBank(ByRef lines() As String) As String
    Dim i As Long, s As String
    Dim hasProCreditHeader As Boolean, hasProCreditAccount As Boolean
    Dim hasHalkHeader As Boolean, hasHalkAccount As Boolean

    ' Fingerprint = NASLOV izvoda + PREFIKS racuna banke (ne naziv banke).
    ' Nazivi "ProCredit"/"HALKBANK" se javljaju kao PARTNER u tudjim izvodima, pa
    ' bi labava detekcija po nazivu pogresno preusmerila ceo izvod (regresija na
    ' Komercijalna putu). Racun-prefiks 220-/155- je stabilan kod banke izvoda.
    For i = LBound(lines) To UBound(lines)
        s = lines(i)
        If InStr(1, s, "STANJE I PROMENE SREDSTAVA", vbTextCompare) > 0 Then hasProCreditHeader = True
        If InStr(1, s, "220-", vbTextCompare) > 0 Then hasProCreditAccount = True
        If InStr(1, s, "INFORMACIJE O PLATNIM TRANSAKCIJAMA", vbTextCompare) > 0 Then hasHalkHeader = True
        If InStr(1, s, "155-", vbTextCompare) > 0 Then hasHalkAccount = True
    Next i

    If hasProCreditHeader And hasProCreditAccount Then
        DetectBank = "PROCREDIT"
    ElseIf hasHalkHeader And hasHalkAccount Then
        DetectBank = "HALK"
    Else
        DetectBank = "KOMERC"
    End If
End Function

' Alt+F8: bank-agnostic test -- DetectBank + pun parse + per-red dump. Radi za sve banke.
Public Sub Test_BankParse()
    Dim pdfPath As String, txt As String, tmp As String, lines() As String
    Dim parsed As Variant, i As Long

    pdfPath = PickPdf()
    If pdfPath = "" Then Exit Sub

    txt = ExtractTextFromPdf(pdfPath)
    tmp = Replace(Replace(txt, Chr$(12), vbLf), vbCr, "")
    lines = Split(tmp, vbLf)

    Debug.Print "=== DetectBank: " & DetectBank(lines) & " ==="

    On Error Resume Next
    parsed = ParseBankaIzvodForImport(txt, "test.pdf")
    If Err.Number <> 0 Then
        Debug.Print "PARSE FAIL: [" & Err.Number & "] " & Err.description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If IsEmpty(parsed) Then
        Debug.Print "PARSE: Empty (nema transakcija)"
        Exit Sub
    End If

    Debug.Print "Izvod=" & parsed(1, 1) & "  Datum=" & parsed(1, 2) & "  Racun=" & parsed(1, 3)
    Debug.Print "Saldo: Pocetno=" & parsed(1, 14) & " Novo=" & parsed(1, 15) & _
                " Duguje=" & parsed(1, 16) & " Potrazuje=" & parsed(1, 17)
    Debug.Print "--- OK: " & UBound(parsed, 1) & " transakcija ---"
    For i = 1 To UBound(parsed, 1)
        Debug.Print i & " | " & parsed(i, 4) & " | " & parsed(i, 5) & _
                    " | racun=" & parsed(i, 6) & _
                    " | Isl=" & parsed(i, 8) & " Upl=" & parsed(i, 7) & _
                    " | sif=" & parsed(i, 9) & " | " & parsed(i, 10) & _
                    " | poz=" & parsed(i, 11) & " | ref=" & parsed(i, 12)
    Next i
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

    ' Multi-bank dispatch (Case Else = Komercijalna, backward-compatible).
    Dim bankId As String
    bankId = DetectBank(lines)

    Select Case bankId
        Case "PROCREDIT"
            brojIzvoda = ExtractIzvodBrojProCredit(lines)
            datumIzvoda = ExtractIzvodDatumProCredit(lines)
            brojRacuna = ExtractIzvodRacunProCredit(lines)
            saldo = ExtractIzvodSaldoProCredit(lines)
            txData = ParseBankaIzvodProCredit(txt)
        Case "HALK"
            brojIzvoda = ExtractIzvodBrojHalk(lines)
            datumIzvoda = ExtractIzvodDatumHalk(lines)
            brojRacuna = ExtractIzvodRacunHalk(lines)
            saldo = ExtractIzvodSaldoHalk(lines)
            txData = ParseBankaIzvodHalk(txt)
        Case Else
            brojIzvoda = ExtractIzvodBrojPdfText(lines)
            datumIzvoda = ExtractIzvodDatumPdfText(lines)
            brojRacuna = ExtractIzvodRacunPdfText(lines)
            saldo = ExtractIzvodSaldoPdfText(lines)
            txData = ParseBankaIzvodPdfText(txt)
    End Select
    
    If Trim$(brojIzvoda) = "" Then
        Err.Raise vbObjectError + 1000, "ParseBankaIzvodForImport", "Broj izvoda nije pronadjen."
    End If
    
    If Trim$(datumIzvoda) = "" Then
        Err.Raise vbObjectError + 1001, "ParseBankaIzvodForImport", "Datum izvoda nije pronadjen."
    End If
    
    If Trim$(brojRacuna) = "" Then
        Err.Raise vbObjectError + 1002, "ParseBankaIzvodForImport", "Broj ra" & ChrW(269) & "una izvoda nije pronadjen."
    End If
    
    ' v6.18+: saldo block je izvucen po banci u dispatch-u gore.
    If Not saldo.parsed Then
        Err.Raise vbObjectError + 1003, "ParseBankaIzvodForImport", _
            "STANJE blok izvoda " & brojIzvoda & " nije pronadjen ili ne sadrzi " & _
            "ocekivana saldo polja (Prethodno stanje, Duguje, Potrazuje, Novo stanje, Zadu" & ChrW(382) & "enje, Odobrenje)."
    End If
    
    ' txData je izvucen po banci u dispatch-u gore.
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
            ", banka reported Zadu" & ChrW(382) & "enje=" & saldo.BrojNalogaZaduzenje
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
    Const SRC As String = "SaveBankaImportRows"

    On Error GoTo EH

    Dim lo As ListObject
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
    Dim duplicateCount As Long

    If IsEmpty(data) Then
        Err.Raise ERR_BIM_SAVE_BASE + 1, SRC, "Nema data array za staging."
    End If

    If Not IsArray(data) Then
        Err.Raise ERR_BIM_SAVE_BASE + 2, SRC, "Staging data nije array."
    End If

    If UBound(data, 1) < 1 Then
        Err.Raise ERR_BIM_SAVE_BASE + 3, SRC, "Staging data nema redove."
    End If

    Set lo = GetTable(TBL_BANKA_IMPORT)
    If lo Is Nothing Then
        Err.Raise ERR_BIM_SAVE_BASE + 4, SRC, _
                  "Ne postoji tabela: " & TBL_BANKA_IMPORT
    End If

    ' Fail-fast schema validation. Missing column must stop import immediately.
    colID = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ID, SRC)
    colBrojDok = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, SRC)
    colDatumIzvoda = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_IZVODA, SRC)
    colBrojRacuna = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_RACUNA, SRC)
    colDatumTx = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE, SRC)
    colPartner = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER, SRC)
    colPartnerKonto = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER_KONTO, SRC)
    colOpis = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OPIS, SRC)
    colUplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA, SRC)
    colIsplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA, SRC)
    colValuta = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_VALUTA, SRC)
    colPozivNaBroj = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POZIV_NA_BROJ, SRC)
    colSvrha = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_SVRHA_PLACANJA, SRC)
    colRef = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BANKA_REFERENZ, SRC)
    colIzvorFajl = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_IZVOR_FAJL, SRC)
    colImportVreme = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_IMPORT_VREME, SRC)
    colObradjeno = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_OBRADJENO, SRC)
    colStornirano = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_STORNIRANO, SRC)

    colPocetnoStanje = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_POCETNO_STANJE, SRC)
    colZavrsnoStanje = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ZAVRSNO_STANJE, SRC)
    colUkupanDuguje = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_DUGUJE, SRC)
    colUkupanPotrazuje = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UKUPAN_POTRAZUJE, SRC)

    colCount = lo.ListColumns.count
    
    For i = 1 To UBound(data, 1)
        If IsDuplicateBankaImport( _
            CStr(data(i, 1)), _
            data(i, 4), _
            CDbl(NzBIM(data(i, 7), 0#)), _
            CDbl(NzBIM(data(i, 8), 0#)), _
            CStr(data(i, 5)), _
            CStr(data(i, 12)) _
        ) Then
            duplicateCount = duplicateCount + 1
        Else
            newID = GetNextID(TBL_BANKA_IMPORT, COL_BIM_ID, PREFIX_BANKA_IMPORT)

            If Len(Trim$(newID)) = 0 Then
                Err.Raise ERR_BIM_SAVE_BASE + 5, SRC, _
                          "GetNextID nije vratio BankaImportID. Row=" & CStr(i)
            End If

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

            ' v6.18+ saldo metadata copied to every row from same izvod.
            rowData(colPocetnoStanje) = CDbl(NzBIM(data(i, 14), 0#))
            rowData(colZavrsnoStanje) = CDbl(NzBIM(data(i, 15), 0#))
            rowData(colUkupanDuguje) = CDbl(NzBIM(data(i, 16), 0#))
            rowData(colUkupanPotrazuje) = CDbl(NzBIM(data(i, 17), 0#))

            rowIdx = AppendRow(TBL_BANKA_IMPORT, rowData)

            If rowIdx <= 0 Then
                Err.Raise ERR_BIM_SAVE_BASE + 6, SRC, _
                          "AppendRow failed for " & TBL_BANKA_IMPORT & _
                          ". Row=" & CStr(i) & _
                          " BrojDokumenta=" & CStr(data(i, 1)) & _
                          " Partner=" & CStr(data(i, 5)) & _
                          " Referenz=" & CStr(data(i, 12))
            End If

            savedCount = savedCount + 1
        End If
    Next i
    
    Debug.Print SRC & " completed. Saved=" & CStr(savedCount) & _
                " Duplicates=" & CStr(duplicateCount)

    SaveBankaImportRows = savedCount
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
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function


Public Function IsDuplicateBankaImport(ByVal brojDokumenta As String, _
                                       ByVal datumTransakcije As Variant, _
                                       ByVal uplata As Double, _
                                       ByVal isplata As Double, _
                                       ByVal partner As String, _
                                       ByVal bankaReferenz As String) As Boolean
    Const SRC As String = "IsDuplicateBankaImport"

    On Error GoTo EH

    Dim data As Variant
    Dim i As Long

    Dim colBrojDok As Long
    Dim colDatumTx As Long
    Dim colUplata As Long
    Dim colIsplata As Long
    Dim colPartner As Long
    Dim colRef As Long

    data = GetTableData(TBL_BANKA_IMPORT)

    If IsEmpty(data) Then
        IsDuplicateBankaImport = False
        Exit Function
    End If

    data = ExcludeStornirano(data, TBL_BANKA_IMPORT)

    If IsEmpty(data) Then
        IsDuplicateBankaImport = False
        Exit Function
    End If
    
    colBrojDok = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BROJ_DOKUMENTA, SRC)
    colDatumTx = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_DATUM_TRANSAKCIJE, SRC)
    colUplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_UPLATA, SRC)
    colIsplata = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_ISPLATA, SRC)
    colPartner = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_PARTNER, SRC)
    colRef = RequireColumnIndex(TBL_BANKA_IMPORT, COL_BIM_BANKA_REFERENZ, SRC)

    For i = 1 To UBound(data, 1)
        If Trim$(CStr(data(i, colBrojDok))) = Trim$(brojDokumenta) Then

            If Len(Trim$(bankaReferenz)) > 0 Then
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

    IsDuplicateBankaImport = False
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
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function

' helper jer je GetFileNameFromPath u modBankaImport private -- kopija
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

Private Function ClassifyBankaImportError(ByVal errNumber As Long, _
                                          ByVal errSource As String, _
                                          ByVal errDescription As String) As String
    Dim s As String
    s = UCase$(Trim$(errSource & " " & errDescription))

    If InStr(1, s, "PDFTOTEXT", vbTextCompare) > 0 Or _
       InStr(1, s, "PDF EXTRACT", vbTextCompare) > 0 Or _
       InStr(1, s, "EXTRACTTEXTFROMPDF", vbTextCompare) > 0 Then
        ClassifyBankaImportError = BIM_STATUS_EXTRACT_ERROR
        Exit Function
    End If

    If InStr(1, s, "INTEGRITY FAIL", vbTextCompare) > 0 Or _
       InStr(1, s, "PARSER MISMATCH", vbTextCompare) > 0 Or _
       InStr(1, s, "PARSER COUNT MISMATCH", vbTextCompare) > 0 Or _
       InStr(1, s, "STANJE BLOK", vbTextCompare) > 0 Then
        ClassifyBankaImportError = BIM_STATUS_INTEGRITY_ERROR
        Exit Function
    End If

    If InStr(1, s, "APPENDROW", vbTextCompare) > 0 Or _
       InStr(1, s, "SAVEBANKAIMPORTROWS", vbTextCompare) > 0 Then
        ClassifyBankaImportError = BIM_STATUS_APPEND_ERROR
        Exit Function
    End If

    If InStr(1, s, "REQUIRECOLUMNINDEX", vbTextCompare) > 0 Or _
       InStr(1, s, "KOLONA", vbTextCompare) > 0 Or _
       InStr(1, s, "COLUMN", vbTextCompare) > 0 Or _
       InStr(1, s, "SCHEMA", vbTextCompare) > 0 Then
        ClassifyBankaImportError = BIM_STATUS_SCHEMA_ERROR
        Exit Function
    End If

    If InStr(1, s, "PARSE", vbTextCompare) > 0 Or _
       InStr(1, s, "PARSER", vbTextCompare) > 0 Or _
       InStr(1, s, "BROJ IZVODA", vbTextCompare) > 0 Or _
       InStr(1, s, "DATUM IZVODA", vbTextCompare) > 0 Or _
       InStr(1, s, "BROJ RA" & ChrW(268) & "UNA", vbTextCompare) > 0 Then
        ClassifyBankaImportError = BIM_STATUS_PARSE_ERROR
        Exit Function
    End If

    ClassifyBankaImportError = BIM_STATUS_UNKNOWN_ERROR
End Function

Private Sub AddPendingBankaFileMove(ByRef pendingMoves As Collection, _
                                    ByVal sourcePath As String, _
                                    ByVal targetPath As String, _
                                    ByVal statusText As String)
    If pendingMoves Is Nothing Then Set pendingMoves = New Collection

    ' Array indexes:
    '   0 = source path
    '   1 = target path
    '   2 = status
    pendingMoves.Add Array(sourcePath, targetPath, statusText)
End Sub

Private Sub ExecutePendingBankaFileMoves(ByVal pendingMoves As Collection)
    Const SRC As String = "ExecutePendingBankaFileMoves"

    On Error GoTo EH

    If pendingMoves Is Nothing Then Exit Sub

    Dim i As Long
    Dim moveData As Variant
    Dim sourcePath As String
    Dim targetPath As String
    Dim statusText As String

    For i = 1 To pendingMoves.count
        moveData = pendingMoves(i)

        sourcePath = CStr(moveData(0))
        targetPath = CStr(moveData(1))
        statusText = CStr(moveData(2))

        If Len(Trim$(sourcePath)) = 0 Then GoTo NextMove
        If Len(Trim$(targetPath)) = 0 Then GoTo NextMove

        If Dir$(sourcePath) = "" Then
            Debug.Print SRC & ": source missing, skip. Source=" & sourcePath
            GoTo NextMove
        End If

        MoveFileSafe sourcePath, targetPath

        Debug.Print SRC & ": moved. Status=" & statusText & _
                    " Source=" & sourcePath & _
                    " Target=" & targetPath

NextMove:
    Next i

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, _
              "Pomeranje PDF fajla nije uspelo. Potrebna je rucna provera foldera. " & _
              "Source=" & errSrc & " | " & errDesc
End Sub

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


Public Sub Test_SaldoIntegrityOnSamplePDF()
    ' Pokreni rucno na sample PDF-u (21.09.2021 Komercijalna Banka).
    ' Ocekivano: saldo polja parsiraju (1775.16, 5230.00, 6000.00, 2545.16, 3, 1),
    ' integrity check prolazi.
    
    Dim pdfPath As String
    Dim txt As String
    Dim lines() As String
    Dim saldo As BankIzvodSaldo
    Dim parsed As Variant
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
    Debug.Print "Parsed: " & saldo.parsed
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
    parsed = ParseBankaIzvodForImport(txt, GetFileNameFromPath2(pdfPath))
    If Err.Number <> 0 Then
        Debug.Print "FAIL: " & Err.Number & " - " & Err.description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    If IsEmpty(parsed) Then
        Debug.Print "FAIL: parse returned Empty"
        Exit Sub
    End If
    
    Debug.Print "OK: parsed " & UBound(parsed, 1) & " transakcija sa saldo metadata"
    Debug.Print "First row saldo: Pocetno=" & parsed(1, 14) & _
                " Zavrsno=" & parsed(1, 15) & _
                " Duguje=" & parsed(1, 16) & _
                " Potrazuje=" & parsed(1, 17)
End Sub

' ============================================================
' BANKA DRIVE -> LOCAL INBOX PRODUCTION PULL
'
' Purpose:
' - Pull bank statement PDFs from Google Drive for Desktop folder
'   into local C:\AgriX\Bank\Inbox.
' - Move Drive original to Drive\Downloaded only after local copy
'   is verified.
' - Then existing ImportBankaInbox_TX can parse/stage/move local PDFs.
'
' Required config:
'   BANKA_DRIVE_SOURCE_PATH
'   BANKA_DRIVE_DOWNLOADED_PATH          optional
'   BANKA_DRIVE_MAX_FILES                optional, default 50
'   BANKA_DRIVE_MIN_FILE_AGE_SECONDS     optional, default 15
'
' Existing config reused:
'   BANKA_INBOX_PATH
'   BANKA_PROCESSED_PATH
'   BANKA_ERROR_PATH
' ============================================================

Private Function BankaDrivePullConfigured() As Boolean
    BankaDrivePullConfigured = (Len(Trim$(GetLocalConfigValue("BANKA_DRIVE_SOURCE_PATH", ""))) > 0)
End Function

Public Function PullBankPdfsFromDriveProduction() As Long
    Const SRC As String = "PullBankPdfsFromDriveProduction"

    On Error GoTo EH

    Dim driveSourcePath As String
    Dim driveDownloadedPath As String
    Dim localInboxPath As String
    Dim maxFiles As Long
    Dim minAgeSeconds As Long
    Dim files As Collection
    Dim item As Variant
    Dim pulledCount As Long

    driveSourcePath = BankaNormalizeFolderPath(GetLocalConfigValue("BANKA_DRIVE_SOURCE_PATH", ""))
    driveDownloadedPath = BankaNormalizeFolderPath(GetLocalConfigValue("BANKA_DRIVE_DOWNLOADED_PATH", ""))
    localInboxPath = BankaNormalizeFolderPath(GetBankaInboxPath())

    maxFiles = CLng(val(GetLocalConfigValue("BANKA_DRIVE_MAX_FILES", "50")))
    If maxFiles <= 0 Then maxFiles = 50

    minAgeSeconds = CLng(val(GetLocalConfigValue("BANKA_DRIVE_MIN_FILE_AGE_SECONDS", "15")))
    If minAgeSeconds < 0 Then minAgeSeconds = 15

    If Len(driveSourcePath) = 0 Then Exit Function

    If Len(driveDownloadedPath) = 0 Then
        driveDownloadedPath = BankaParentFolderPath(driveSourcePath) & "\Downloaded"
    End If

    If Dir$(driveSourcePath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 9501, SRC, _
            "Drive source folder ne postoji ili nije dostupan: " & driveSourcePath
    End If

    If StrComp(driveSourcePath, localInboxPath, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 9502, SRC, _
            "Drive source i lokalni inbox ne smeju biti isti folder."
    End If

    BankaEnsureFolderExistsRecursive localInboxPath
    BankaEnsureFolderExistsRecursive driveDownloadedPath

    Set files = BankaCollectPdfFiles(driveSourcePath)

    For Each item In files
        If pulledCount >= maxFiles Then Exit For

        If BankaIsFileReadyForPull(CStr(item), minAgeSeconds) Then
            BankaPullOnePdfFromDrive CStr(item), localInboxPath, driveDownloadedPath
            pulledCount = pulledCount + 1
        Else
            Debug.Print SRC & ": skip not-ready file: " & CStr(item)
        End If
    Next item

    Debug.Print SRC & ": completed. Pulled=" & CStr(pulledCount)
    PullBankPdfsFromDriveProduction = pulledCount
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Private Sub BankaPullOnePdfFromDrive(ByVal sourcePdfPath As String, _
                                     ByVal localInboxPath As String, _
                                     ByVal driveDownloadedPath As String)
    Const SRC As String = "BankaPullOnePdfFromDrive"

    On Error GoTo EH

    Dim fileName As String
    Dim localFinalPath As String
    Dim localTempPath As String
    Dim driveDownloadedTargetPath As String
    Dim sourceSize As Long
    Dim copiedSize As Long
    Dim movedOk As Boolean

    If Dir$(sourcePdfPath) = "" Then
        Err.Raise vbObjectError + 9510, SRC, "PDF ne postoji: " & sourcePdfPath
    End If

    fileName = BankaSafeFileName(BankaFileNameFromPath(sourcePdfPath))

    If LCase$(Right$(fileName, 4)) <> ".pdf" Then
        Err.Raise vbObjectError + 9511, SRC, "Fajl nije PDF: " & fileName
    End If

    sourceSize = FileLen(sourcePdfPath)
    If sourceSize <= 0 Then
        Err.Raise vbObjectError + 9512, SRC, "PDF je prazan: " & sourcePdfPath
    End If

    localFinalPath = GetUniqueTargetPath(localInboxPath & "\" & fileName)
    localTempPath = localFinalPath & ".part"
    driveDownloadedTargetPath = GetUniqueTargetPath(driveDownloadedPath & "\" & fileName)

    If Dir$(localTempPath) <> "" Then Kill localTempPath

    FileCopy sourcePdfPath, localTempPath

    If Dir$(localTempPath) = "" Then
        Err.Raise vbObjectError + 9513, SRC, "Temp lokalni PDF nije kreiran."
    End If

    copiedSize = FileLen(localTempPath)
    If copiedSize <> sourceSize Then
        Err.Raise vbObjectError + 9514, SRC, _
            "Kopirani PDF nema istu velicinu. Source=" & CStr(sourceSize) & _
            " Local=" & CStr(copiedSize) & " File=" & fileName
    End If

    Name localTempPath As localFinalPath

    If Dir$(localFinalPath) = "" Then
        Err.Raise vbObjectError + 9515, SRC, "Final lokalni PDF nije kreiran."
    End If

    MoveFileSafe sourcePdfPath, driveDownloadedTargetPath
    movedOk = True

    Debug.Print SRC & ": pulled. Source=" & sourcePdfPath & _
                " Local=" & localFinalPath & _
                " DriveDownloaded=" & driveDownloadedTargetPath

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    If Not movedOk Then
        If Len(localTempPath) > 0 And Dir$(localTempPath) <> "" Then Kill localTempPath
        If Len(localFinalPath) > 0 And Dir$(localFinalPath) <> "" Then Kill localFinalPath
    End If

    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Sub

Private Function BankaCollectPdfFiles(ByVal folderPath As String) As Collection
    Dim result As Collection
    Dim f As String

    Set result = New Collection
    folderPath = BankaNormalizeFolderPath(folderPath)

    f = Dir$(folderPath & "\*.pdf")
    Do While Len(f) > 0
        result.Add folderPath & "\" & f
        f = Dir$
    Loop

    Set BankaCollectPdfFiles = result
End Function

Private Function BankaIsFileReadyForPull(ByVal filePath As String, _
                                         ByVal minAgeSeconds As Long) As Boolean
    On Error GoTo NotReady

    Dim ageSeconds As Long
    Dim s1 As Long
    Dim s2 As Long

    If Dir$(filePath) = "" Then GoTo NotReady

    ageSeconds = DateDiff("s", FileDateTime(filePath), Now)
    If ageSeconds < minAgeSeconds Then GoTo NotReady

    s1 = FileLen(filePath)
    If s1 <= 0 Then GoTo NotReady

    DoEvents

    s2 = FileLen(filePath)
    If s1 <> s2 Then GoTo NotReady

    BankaIsFileReadyForPull = True
    Exit Function

NotReady:
    BankaIsFileReadyForPull = False
End Function

Private Function BankaNormalizeFolderPath(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)

    Do While Len(folderPath) > 1 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    BankaNormalizeFolderPath = folderPath
End Function

Private Function BankaParentFolderPath(ByVal folderPath As String) As String
    Dim p As Long

    folderPath = BankaNormalizeFolderPath(folderPath)
    p = InStrRev(folderPath, "\")

    If p <= 0 Then
        BankaParentFolderPath = folderPath
    Else
        BankaParentFolderPath = Left$(folderPath, p - 1)
    End If
End Function

Private Function BankaFileNameFromPath(ByVal filePath As String) As String
    Dim p As Long

    p = InStrRev(filePath, "\")
    If p > 0 Then
        BankaFileNameFromPath = Mid$(filePath, p + 1)
    Else
        BankaFileNameFromPath = filePath
    End If
End Function

Private Function BankaSafeFileName(ByVal fileName As String) As String
    Dim badChars As Variant
    Dim i As Long

    fileName = Trim$(fileName)
    If Len(fileName) = 0 Then fileName = "bank.pdf"

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(badChars) To UBound(badChars)
        fileName = Replace(fileName, CStr(badChars(i)), "_")
    Next i

    If Len(fileName) > 180 Then fileName = Left$(fileName, 180)

    BankaSafeFileName = fileName
End Function

Private Sub BankaEnsureFolderExistsRecursive(ByVal folderPath As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long

    folderPath = BankaNormalizeFolderPath(folderPath)

    If Len(folderPath) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) <> "" Then Exit Sub

    parts = Split(folderPath, "\")
    currentPath = parts(0)

    If Right$(currentPath, 1) = ":" Then currentPath = currentPath & "\"

    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Right$(currentPath, 1) <> "\" Then currentPath = currentPath & "\"
            currentPath = currentPath & parts(i)

            If Dir$(currentPath, vbDirectory) = "" Then MkDir currentPath
        End If
    Next i
End Sub

