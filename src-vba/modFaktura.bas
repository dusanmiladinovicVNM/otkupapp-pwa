Attribute VB_Name = "modFaktura"

Option Explicit

' ============================================================
' modFaktura v2.1 - Rechnungserstellung
' GEAeNDERT: Basiert auf tblPrijemnica statt tblIsporuka
' Faktura-Betrag = Prijemnica.Kolicina x Prijemnica.Cena
' ============================================================

Public Function CreateFaktura_TX(ByVal kupacID As String, _
                                  ByVal stavke As Collection) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_NOVAC

    CreateFaktura_TX = CreateFaktura(kupacID, stavke)

    If CreateFaktura_TX = "" Then
        Err.Raise vbObjectError + 1701, "CreateFaktura_TX", _
                  "CreateFaktura fehlgeschlagen"
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="FAKTURA_CREATE_SUCCESS", _
        severity:="INFO", _
        message:="Faktura created successfully", _
        userId:="Operator", _
        moduleName:="modFaktura", _
        procedureName:="CreateFaktura_TX", _
        entityType:="Faktura", _
        entityID:=CreateFaktura_TX, _
        correlationId:=CreateFaktura_TX
    On Error GoTo 0

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    LogErr "CreateFaktura_TX"

    Monitor_Error _
        moduleName:="modFaktura", _
        procedureName:="CreateFaktura_TX", _
        entityType:="Faktura", _
        entityID:=CreateFaktura_TX, _
        correlationId:=CreateFaktura_TX, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    Monitor_Event _
        eventType:="FAKTURA_CREATE_FAIL", _
        severity:="ERROR", _
        message:=errDesc, _
        userId:="Operator", _
        moduleName:="modFaktura", _
        procedureName:="CreateFaktura_TX", _
        entityType:="Faktura", _
        entityID:=CreateFaktura_TX, _
        correlationId:=CreateFaktura_TX

    If Not tx Is Nothing Then tx.RollbackTx

    On Error GoTo 0

    CreateFaktura_TX = ""

    Debug.Print " CreateFaktura_TX failed. Source=" & errSrc & _
                " Err=" & CStr(errNum) & _
                " Desc=" & errDesc
End Function

' Base funkcija -- NE zovi je spolja. Jedini ulaz je CreateFaktura_TX, koji drzi
' snapshot transakciju; direktan poziv bi kod greske ostavio pola upisa
' (faktura header bez stavki, prijemnice markirane bez fakture) -- AUD-011 /
' FM-0034 #3.
Private Function CreateFaktura(ByVal kupacID As String, _
                               ByVal stavke As Collection) As String
    On Error GoTo EH

    If Trim$(kupacID) = "" Then
        Err.Raise vbObjectError + 1702, "CreateFaktura", _
                  "KupacID je obavezan."
    End If

    If stavke Is Nothing Then
        Err.Raise vbObjectError + 1703, "CreateFaktura", _
                  "Stavke nisu prosledjene."
    End If

    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1704, "CreateFaktura", _
                  "Faktura mora imati bar jednu stavku."
    End If

    ' Fail-fast schema guards
    RequireColumnIndex TBL_FAKTURE, COL_FAK_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_BROJ, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_DATUM, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_KUPAC, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_IZNOS, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_STATUS, "CreateFaktura"

    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRIJEMNICA_ID, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_KOLICINA, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_CENA, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_KLASA, "CreateFaktura"
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRIJEMNICE, "CreateFaktura"

    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_ID, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_KUPAC, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_STORNIRANO, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_CENA, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_KLASA, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_BROJ, "CreateFaktura"

    Dim fakturaID As String
    fakturaID = GetNextID(TBL_FAKTURE, COL_FAK_ID, "FAK-")

    If fakturaID = "" Then
        Err.Raise vbObjectError + 1705, "CreateFaktura", _
                  "GetNextID nije vratio FakturaID."
    End If

    Dim brojFakture As String
    brojFakture = GenerateBrojFakture()

    If brojFakture = "" Then
        Err.Raise vbObjectError + 1706, "CreateFaktura", _
                  "GenerateBrojFakture nije vratio broj fakture."
    End If
    
        Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(prijData) Then
        Err.Raise vbObjectError + 1717, "CreateFaktura", _
                  "Tabela prijemnica je prazna."
    End If

    Dim colPrjKol As Long
    Dim colPrjCena As Long
    Dim colPrjKlasa As Long
    Dim colPrjBroj As Long
    Dim colPrjKupac As Long

    colPrjKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "CreateFaktura")
    colPrjCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, "CreateFaktura")
    colPrjKlasa = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, "CreateFaktura")
    colPrjBroj = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, "CreateFaktura")
    colPrjKupac = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, "CreateFaktura")

    ' Pre-validacija svih prijemnica pre bilo kog upisa.
    ' Business module trusts only PrijemnicaID from caller.
    ' Kolicina/Cena/Klasa/BrojPrijemnice are derived from tblPrijemnica.
    Dim s As Variant
    Dim prijemnicaID As String
    Dim rows As Collection
    Dim prijRows As Object
    Dim prijValues As Object

    Set prijRows = CreateObject("Scripting.Dictionary")
    Set prijValues = CreateObject("Scripting.Dictionary")

    For Each s In stavke

        prijemnicaID = GetPrijemnicaIDFromFakturaStavka(s, "CreateFaktura")

        Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID)

        If rows Is Nothing Or rows.count = 0 Then
            Err.Raise vbObjectError + 1708, "CreateFaktura", _
                      "Prijemnica nije pronadena: " & prijemnicaID
        End If

        If prijRows.Exists(prijemnicaID) Then
            Err.Raise vbObjectError + 1710, "CreateFaktura", _
                      "Dupla prijemnica u izboru: " & prijemnicaID
        End If

        ' Fail-closed kod duplog PrijemnicaID: bez ovoga se tiho uzima prvi
        ' pogodak, pa se fakturise kolicina/cena pogresnog reda (AUD-011 /
        ' FM-0034 #2). Isti obrazac kao RequireSingleFakturaRow nad tblFakture.
        If rows.count > 1 Then
            Err.Raise vbObjectError + 1707, "CreateFaktura", _
                      "Duplikat PrijemnicaID=" & prijemnicaID & _
                      "; Count=" & CStr(rows.count)
        End If

        Dim rowPrijValidate As Long
        rowPrijValidate = CLng(rows(1))

        ' Vlasnistvo: prijemnica mora da pripada kupcu fakture. Bez ove provere
        ' se prijemnica drugog kupca moze zavuci u fakturu (UI filter po kupcu
        ' nije sigurnosna granica) -- AUD-011 / FM-0034 #1.
        Dim prjKupac As String
        prjKupac = Trim$(CStr(prijData(rowPrijValidate, colPrjKupac)))

        If prjKupac <> Trim$(kupacID) Then
            Err.Raise vbObjectError + 1721, "CreateFaktura", _
                      "Prijemnica pripada drugom kupcu. PrijemnicaID=" & prijemnicaID & _
                      "; Prijemnica.KupacID=" & prjKupac & _
                      "; Faktura.KupacID=" & Trim$(kupacID)
        End If

        If Not IsPrijemnicaAvailableForFaktura(rowPrijValidate, prijemnicaID) Then
            Err.Raise vbObjectError + 1709, "CreateFaktura", _
                      "Prijemnica je ve" & ChrW(263) & " fakturisana ili stornirana: " & prijemnicaID
        End If

        If Not IsNumeric(prijData(rowPrijValidate, colPrjKol)) Then
            Err.Raise vbObjectError + 1711, "CreateFaktura", _
                      "Koli" & ChrW(269) & "ina nije numericka za prijemnicu: " & prijemnicaID
        End If

        If Not IsNumeric(prijData(rowPrijValidate, colPrjCena)) Then
            Err.Raise vbObjectError + 1711, "CreateFaktura", _
                      "Cena nije numericka za prijemnicu: " & prijemnicaID
        End If

        Dim prjKolicina As Double
        Dim prjCena As Double
        Dim prjKlasa As String
        Dim prjBroj As String

        prjKolicina = CDbl(prijData(rowPrijValidate, colPrjKol))
        prjCena = CDbl(prijData(rowPrijValidate, colPrjCena))
        prjKlasa = CStr(prijData(rowPrijValidate, colPrjKlasa))
        prjBroj = CStr(prijData(rowPrijValidate, colPrjBroj))

        If prjKolicina <= 0 Then
            Err.Raise vbObjectError + 1712, "CreateFaktura", _
                      "Koli" & ChrW(269) & "ina mora biti veca od nule. PrijemnicaID=" & prijemnicaID
        End If

        If prjCena < 0 Then
            Err.Raise vbObjectError + 1713, "CreateFaktura", _
                      "Cena ne sme biti negativna. PrijemnicaID=" & prijemnicaID
        End If

        prijRows.Add prijemnicaID, rowPrijValidate
        prijValues.Add prijemnicaID, Array(prjKolicina, prjCena, prjKlasa, prjBroj)

    Next s

    ' Ukupan iznos se racuna iz canonical tblPrijemnica vrednosti.
    Dim ukupno As Double
    Dim key As Variant
    Dim prjVals As Variant

    For Each key In prijValues.keys
        prjVals = prijValues(CStr(key))
        ukupno = ukupno + (CDbl(prjVals(0)) * CDbl(prjVals(1)))
    Next key

    If ukupno <= 0 Then
        Err.Raise vbObjectError + 1714, "CreateFaktura", _
                  "Ukupan iznos fakture mora biti veci od nule."
    End If

    ' Faktura header
    Dim fakturaRow As Variant
    fakturaRow = Array( _
        fakturaID, _
        brojFakture, _
        Date, _
        kupacID, _
        ukupno, _
        STATUS_NEPLACENO, _
        Empty, _
        "", _
        "", _
        WF_LOCAL_FINALIZED, _
        "", _
        "", _
        "", _
        Empty, _
        Empty, _
        "", _
        "", _
        "", _
        0, _
        "Ne", _
        "" _
    )

    If AppendRow(TBL_FAKTURE, fakturaRow) <= 0 Then
        Err.Raise vbObjectError + 1715, "CreateFaktura", _
                  "AppendRow fehlgeschlagen fuer tblFakture."
    End If

    ' Faktura stavke + markiranje prijemnica
    Dim stavkaID As String
    Dim stavkaNum As Long
    Dim stavkaRow As Variant
    Dim rowPrij As Long

    For Each s In stavke
        stavkaNum = stavkaNum + 1
        stavkaID = fakturaID & "-" & Format$(stavkaNum, "00")

        prijemnicaID = GetPrijemnicaIDFromFakturaStavka(s, "CreateFaktura")
        rowPrij = CLng(prijRows(prijemnicaID))

        prjVals = prijValues(prijemnicaID)

        stavkaRow = Array( _
            stavkaID, _
            fakturaID, _
            prijemnicaID, _
            CDbl(prjVals(0)), _
            CDbl(prjVals(1)), _
            CStr(prjVals(2)), _
            CStr(prjVals(3)), _
            "", _
            "" _
        )

        If AppendRow(TBL_FAKTURA_STAVKE, stavkaRow) <= 0 Then
            Err.Raise vbObjectError + 1716, "CreateFaktura", _
                      "AppendRow fehlgeschlagen fuer tblFakturaStavke."
        End If

        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURISANO, _
                          "Da", "CreateFaktura"

        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURA_ID, _
                          fakturaID, "CreateFaktura"
    Next s

    ' Avans automatisch verrechnen.
    ' Ovo mora biti base funkcija, ne ApplyAvansToFaktura_TX,
    ' jer CreateFaktura_TX vec drzi siru transakciju.
    ApplyAvansToFaktura kupacID, fakturaID

    CreateFaktura = fakturaID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    LogErr "CreateFaktura"
    On Error GoTo 0

    Err.Raise errNum, "CreateFaktura", _
              "Source=" & errSrc & " | " & errDesc
End Function

Private Function GenerateBrojFakture() As String
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    Dim currentYear As Long
    currentYear = Year(Date)

    Dim maxNum As Long

    If Not IsEmpty(data) Then
        Dim colBroj As Long
        colBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, _
                                     "GenerateBrojFakture")

        Dim i As Long
        Dim broj As String
        Dim parts As Variant
        Dim num As Long
        Dim yr As Long

        For i = 1 To UBound(data, 1)
            broj = Trim$(CStr(data(i, colBroj)))

            If InStr(broj, "/") > 0 Then
                parts = Split(broj, "/")

                If UBound(parts) >= 1 Then
                    num = 0
                    yr = 0

                    If IsNumeric(Trim$(parts(0))) Then num = CLng(Trim$(parts(0)))
                    If IsNumeric(Trim$(parts(1))) Then yr = CLng(Trim$(parts(1)))

                    If yr = currentYear Then
                        If num > maxNum Then maxNum = num
                    End If
                End If
            End If
        Next i
    End If

    GenerateBrojFakture = CStr(maxNum + 1) & "/" & CStr(currentYear)
    Exit Function

EH:
    LogErr "GenerateBrojFakture"
    GenerateBrojFakture = ""
End Function

Public Sub PrintFaktura(ByVal fakturaID As String)
    On Error GoTo EH

    Dim fRow As Long
    fRow = RequireSingleFakturaRow(fakturaID, "PrintFaktura")

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then
        Err.Raise vbObjectError + 1732, "PrintFaktura", _
                  "Tabela faktura je prazna."
    End If

    Dim colFakBroj As Long
    Dim colFakDatum As Long
    Dim colFakKupac As Long
    Dim colFakIznos As Long
    Dim colFakStornirano As Long

    colFakBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, _
                                    "PrintFaktura")
    colFakDatum = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM, _
                                     "PrintFaktura")
    colFakKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, _
                                     "PrintFaktura")
    colFakIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, _
                                     "PrintFaktura")
    colFakStornirano = RequireColumnIndex(TBL_FAKTURE, COL_STORNIRANO, _
                                     "PrintFaktura")

    If UCase$(Trim$(CStr(data(fRow, colFakStornirano)))) = "DA" Then
        Err.Raise vbObjectError + 1736, "PrintFaktura", _
              Poruka("FAK_ERR_STORNIRANA_FAKTURA_MOZE") & fakturaID
    End If

    Dim kupacID As String
    kupacID = Trim$(CStr(data(fRow, colFakKupac)))

    Dim kupacNaziv As String
    kupacNaziv = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, COL_KUP_NAZIV))

    If kupacNaziv = "" Then kupacNaziv = kupacID

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If IsEmpty(stavkeData) Then
        Err.Raise vbObjectError + 1734, "PrintFaktura", _
                  "Faktura nema stavke: " & fakturaID
    End If

    Dim colStFakID As Long, colStBrojPrij As Long, colStKlasa As Long
    Dim colStKol As Long, colStCena As Long
    colStFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, "PrintFaktura")
    colStBrojPrij = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRIJEMNICE, "PrintFaktura")
    colStKlasa = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KLASA, "PrintFaktura")
    colStKol = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KOLICINA, "PrintFaktura")
    colStCena = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_CENA, "PrintFaktura")

    Dim stavke() As Variant
    ReDim stavke(1 To UBound(stavkeData, 1), 1 To 5)
    Dim outRow As Long, j As Long
    Dim kolicina As Double, cena As Double
    For j = 1 To UBound(stavkeData, 1)
        If Trim$(CStr(stavkeData(j, colStFakID))) = fakturaID Then
            outRow = outRow + 1
            kolicina = 0: cena = 0
            If IsNumeric(stavkeData(j, colStKol)) Then kolicina = CDbl(stavkeData(j, colStKol))
            If IsNumeric(stavkeData(j, colStCena)) Then cena = CDbl(stavkeData(j, colStCena))
            stavke(outRow, 1) = stavkeData(j, colStBrojPrij)
            stavke(outRow, 2) = stavkeData(j, colStKlasa)
            stavke(outRow, 3) = kolicina
            stavke(outRow, 4) = cena
            stavke(outRow, 5) = kolicina * cena
        End If
    Next j

    If outRow = 0 Then
        Err.Raise vbObjectError + 1735, "PrintFaktura", _
                  "Nisu pronadene stavke za fakturu: " & fakturaID
    End If

    Dim ukupno As Double
    If IsNumeric(data(fRow, colFakIznos)) Then ukupno = CDbl(data(fRow, colFakIznos))

    Dim ws As Worksheet
    Set ws = FillFakturaSablon(CStr(data(fRow, colFakBroj)), data(fRow, colFakDatum), _
                               kupacNaziv, stavke, outRow, ukupno)
    If ws Is Nothing Then Exit Sub

    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_FAKTURA_PRINT_MODE), "PRINT")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            DocExportPdf ws, ThisWorkbook.path & "\Faktura_" & _
                         Replace(CStr(data(fRow, colFakBroj)), "/", "-") & ".pdf", True
        ' OFF -> bez izlaza
    End Select
    Exit Sub

EH:
    LogErr "PrintFaktura"
    Err.Raise Err.Number, "PrintFaktura", Err.description
End Sub

Public Sub UpdateFakturaStatus(ByVal fakturaID As String)
    On Error GoTo EH

    Const SRC As String = "UpdateFakturaStatus"

    If Trim$(fakturaID) = "" Then
        Err.Raise vbObjectError + 1723, SRC, _
                "FakturaID je obavezan."
    End If

    Dim colID As Long
    Dim colIznos As Long
    Dim colStatus As Long
    Dim colDatumPlacanja As Long
    Dim colStornirano As Long

    colID = RequireColumnIndex(TBL_FAKTURE, COL_FAK_ID, SRC)
    colIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, SRC)
    colStatus = RequireColumnIndex(TBL_FAKTURE, COL_FAK_STATUS, SRC)
    colDatumPlacanja = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM_PLACANJA, SRC)
    colStornirano = RequireColumnIndex(TBL_FAKTURE, COL_STORNIRANO, SRC)

    Dim r As Long
    r = RequireSingleFakturaRow(fakturaID, SRC)

    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)

    If IsEmpty(data) Then
        Err.Raise vbObjectError + 1722, SRC, _
                  "Tabela faktura je prazna."
    End If

    If UCase$(Trim$(CStr(data(r, colStornirano)))) = "DA" Then
        Exit Sub
    End If

    If Not IsNumeric(data(r, colIznos)) Then
        Err.Raise vbObjectError + 1720, SRC, _
                  "Iznos fakture nije numericki: " & fakturaID
    End If

    Dim fakturaIznos As Double
    fakturaIznos = CDbl(data(r, colIznos))

    Dim uplaceno As Double
    uplaceno = GetUplataForFaktura(fakturaID)

    Dim currentStatus As String
    Dim currentDatumPlacanja As String

    currentStatus = Trim$(CStr(data(r, colStatus)))
    currentDatumPlacanja = Trim$(CStr(data(r, colDatumPlacanja)))

    If uplaceno >= fakturaIznos And fakturaIznos > 0 Then

        If currentStatus <> STATUS_PLACENO Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_STATUS, _
                              STATUS_PLACENO, SRC
        End If

        If Len(currentDatumPlacanja) = 0 Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_DATUM_PLACANJA, _
                              Date, SRC
        End If

    Else

        If currentStatus <> STATUS_NEPLACENO Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_STATUS, _
                              STATUS_NEPLACENO, SRC
        End If

        If Len(currentDatumPlacanja) > 0 Then
            RequireUpdateCell TBL_FAKTURE, r, COL_FAK_DATUM_PLACANJA, _
                              Empty, SRC
        End If

    End If

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

Private Function RequireSingleFakturaRow(ByVal fakturaID As String, _
                                         ByVal sourceName As String) As Long
    If Len(Trim$(fakturaID)) = 0 Then
        Err.Raise vbObjectError + 1740, sourceName, _
                  "FakturaID je obavezan."
    End If

    RequireColumnIndex TBL_FAKTURE, COL_FAK_ID, sourceName

    Dim rows As Collection
    Set rows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)

    If rows Is Nothing Then
        Err.Raise vbObjectError + 1741, sourceName, _
                  "FindRows je vratio Nothing za FakturaID=" & fakturaID
    End If

    If rows.count = 0 Then
        Err.Raise vbObjectError + 1742, sourceName, _
                  "Faktura nije pronadena: " & fakturaID
    End If

    If rows.count > 1 Then
        Err.Raise vbObjectError + 1743, sourceName, _
                  "Duplicate FakturaID. FakturaID=" & fakturaID & _
                  "; Count=" & CStr(rows.count)
    End If

    RequireSingleFakturaRow = CLng(rows(1))
End Function

Private Function IsPrijemnicaAvailableForFaktura(ByVal rowIndex As Long, _
                                                 ByVal prijemnicaID As String) As Boolean
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)

    If IsEmpty(data) Then Exit Function

    Dim colFakturisano As Long
    Dim colFakturaID As Long
    Dim colStorno As Long

    colFakturisano = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, _
                                        "IsPrijemnicaAvailableForFaktura")
    colFakturaID = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, _
                                      "IsPrijemnicaAvailableForFaktura")
    colStorno = RequireColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO, _
                                   "IsPrijemnicaAvailableForFaktura")

    If rowIndex <= 0 Or rowIndex > UBound(data, 1) Then Exit Function

    If Trim$(CStr(data(rowIndex, colStorno))) = "Da" Then Exit Function
    If Trim$(CStr(data(rowIndex, colFakturisano))) = "Da" Then Exit Function
    If Trim$(CStr(data(rowIndex, colFakturaID))) <> "" Then Exit Function

    IsPrijemnicaAvailableForFaktura = True
    Exit Function

EH:
    LogErr "IsPrijemnicaAvailableForFaktura"
    IsPrijemnicaAvailableForFaktura = False
End Function

Private Function GetPrijemnicaIDFromFakturaStavka(ByVal stavka As Variant, _
                                                  ByVal sourceName As String) As String
    On Error GoTo EH

    GetPrijemnicaIDFromFakturaStavka = Trim$(CStr(stavka(0)))

    If Len(GetPrijemnicaIDFromFakturaStavka) = 0 Then
        Err.Raise vbObjectError + 1718, sourceName, _
                  "Stavka nema PrijemnicaID."
    End If

    Exit Function

EH:
    Err.Raise vbObjectError + 1719, sourceName, _
              "Neispravan oblik stavke fakture. Ocekuje se da stavka(0) bude PrijemnicaID."
End Function

