Attribute VB_Name = "modFaktura"

Option Explicit

' ============================================================
' modFaktura v2.1 – Rechnungserstellung
' GEÄNDERT: Basiert auf tblPrijemnica statt tblIsporuka
' Faktura-Betrag = Prijemnica.Kolicina × Prijemnica.Cena
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
    Set tx = Nothing
    Exit Function

EH:
    LogErr "CreateFaktura_TX"

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateFaktura_TX = ""
End Function

Public Function CreateFaktura(ByVal kupacID As String, _
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
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, "CreateFaktura"
    RequireColumnIndex TBL_PRIJEMNICA, COL_STORNIRANO, "CreateFaktura"

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

    ' Pre-validacija svih prijemnica pre bilo kog upisa.
    Dim s As Variant
    Dim prijemnicaID As String
    Dim rows As Collection
    Dim prijRows As Object
    Set prijRows = CreateObject("Scripting.Dictionary")

    For Each s In stavke
        prijemnicaID = Trim$(CStr(s(0)))

        If prijemnicaID = "" Then
            Err.Raise vbObjectError + 1707, "CreateFaktura", _
                      "Stavka nema PrijemnicaID."
        End If

        Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID)

        If rows.count = 0 Then
            Err.Raise vbObjectError + 1708, "CreateFaktura", _
                      "Prijemnica nije pronadena: " & prijemnicaID
        End If

        If Not IsPrijemnicaAvailableForFaktura(rows(1), prijemnicaID) Then
            Err.Raise vbObjectError + 1709, "CreateFaktura", _
                      "Prijemnica je vec fakturisana ili stornirana: " & prijemnicaID
        End If

        If prijRows.Exists(prijemnicaID) Then
            Err.Raise vbObjectError + 1710, "CreateFaktura", _
                      "Dupla prijemnica u izboru: " & prijemnicaID
        End If

        prijRows.Add prijemnicaID, rows(1)
    Next s

    ' Ukupan iznos
    Dim ukupno As Double
    Dim kolicina As Double
    Dim cena As Double

    For Each s In stavke
        If Not IsNumeric(s(1)) Or Not IsNumeric(s(2)) Then
            Err.Raise vbObjectError + 1711, "CreateFaktura", _
                      "Kolicina ili cena nisu numericke vrednosti."
        End If

        kolicina = CDbl(s(1))
        cena = CDbl(s(2))

        If kolicina <= 0 Then
            Err.Raise vbObjectError + 1712, "CreateFaktura", _
                      "Kolicina mora biti veca od nule."
        End If

        If cena < 0 Then
            Err.Raise vbObjectError + 1713, "CreateFaktura", _
                      "Cena ne sme biti negativna."
        End If

        ukupno = ukupno + (kolicina * cena)
    Next s

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
                  "AppendRow fehlgeschlagen für tblFakture."
    End If

    ' Faktura stavke + markiranje prijemnica
    Dim stavkaID As String
    Dim stavkaNum As Long
    Dim stavkaRow As Variant
    Dim rowPrij As Long

    For Each s In stavke
        stavkaNum = stavkaNum + 1
        stavkaID = fakturaID & "-" & Format$(stavkaNum, "00")

        prijemnicaID = Trim$(CStr(s(0)))
        rowPrij = CLng(prijRows(prijemnicaID))

        stavkaRow = Array( _
            stavkaID, _
            fakturaID, _
            prijemnicaID, _
            CDbl(s(1)), _
            CDbl(s(2)), _
            CStr(s(3)), _
            CStr(s(4)), _
            "", _
            "" _
        )

        If AppendRow(TBL_FAKTURA_STAVKE, stavkaRow) <= 0 Then
            Err.Raise vbObjectError + 1716, "CreateFaktura", _
                      "AppendRow fehlgeschlagen für tblFakturaStavke."
        End If

        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURISANO, _
                          "Da", "CreateFaktura"

        RequireUpdateCell TBL_PRIJEMNICA, rowPrij, COL_PRJ_FAKTURA_ID, _
                          fakturaID, "CreateFaktura"
    Next s

    ' Avans automatisch verrechnen.
    ' Ovo mora biti base funkcija, ne ApplyAvansToFaktura_TX,
    ' jer CreateFaktura_TX vec drži širu transakciju.
    ApplyAvansToFaktura kupacID, fakturaID

    CreateFaktura = fakturaID
    Exit Function

EH:
    LogErr "CreateFaktura"
    CreateFaktura = ""
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

    If Trim$(fakturaID) = "" Then
        Err.Raise vbObjectError + 1730, "PrintFaktura", _
                  "FakturaID je obavezan."
    End If

    Dim rows As Collection
    Set rows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)

    If rows.count = 0 Then
        Err.Raise vbObjectError + 1731, "PrintFaktura", _
                  "Faktura nije pronadena: " & fakturaID
    End If

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

    colFakBroj = RequireColumnIndex(TBL_FAKTURE, COL_FAK_BROJ, _
                                    "PrintFaktura")
    colFakDatum = RequireColumnIndex(TBL_FAKTURE, COL_FAK_DATUM, _
                                     "PrintFaktura")
    colFakKupac = RequireColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC, _
                                     "PrintFaktura")
    colFakIznos = RequireColumnIndex(TBL_FAKTURE, COL_FAK_IZNOS, _
                                     "PrintFaktura")

    Dim fRow As Long
    fRow = rows(1)

    Dim kupacID As String
    kupacID = Trim$(CStr(data(fRow, colFakKupac)))

    Dim kupacNaziv As String
    kupacNaziv = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, COL_KUP_NAZIV))

    If kupacNaziv = "" Then kupacNaziv = kupacID

    Dim wsSablon As Worksheet
    Set wsSablon = Nothing

    On Error Resume Next
    Set wsSablon = ThisWorkbook.Worksheets("FakturaSablon")
    On Error GoTo EH

    If wsSablon Is Nothing Then
        Err.Raise vbObjectError + 1733, "PrintFaktura", _
                  "FakturaSablon sheet ne postoji."
    End If

    With wsSablon
        .Range("BrojFakture").value = data(fRow, colFakBroj)
        .Range("DatumFakture").value = data(fRow, colFakDatum)
        .Range("KupacNaziv").value = kupacNaziv
    End With

    ClearFakturaStavkeArea wsSablon

    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)

    If IsEmpty(stavkeData) Then
        Err.Raise vbObjectError + 1734, "PrintFaktura", _
                  "Faktura nema stavke: " & fakturaID
    End If

    Dim colStFakID As Long
    Dim colStBrojPrij As Long
    Dim colStKlasa As Long
    Dim colStKol As Long
    Dim colStCena As Long

    colStFakID = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_FAKTURA_ID, _
                                    "PrintFaktura")
    colStBrojPrij = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRIJEMNICE, _
                                       "PrintFaktura")
    colStKlasa = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KLASA, _
                                    "PrintFaktura")
    colStKol = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_KOLICINA, _
                                  "PrintFaktura")
    colStCena = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_CENA, _
                                   "PrintFaktura")

    Dim startCell As Range
    Set startCell = wsSablon.Range("StavkaStart")

    Dim outRow As Long
    Dim j As Long
    Dim kolicina As Double
    Dim cena As Double
    Dim vrednost As Double

    For j = 1 To UBound(stavkeData, 1)
        If Trim$(CStr(stavkeData(j, colStFakID))) = fakturaID Then
            outRow = outRow + 1

            kolicina = 0
            cena = 0
            vrednost = 0

            If IsNumeric(stavkeData(j, colStKol)) Then kolicina = CDbl(stavkeData(j, colStKol))
            If IsNumeric(stavkeData(j, colStCena)) Then cena = CDbl(stavkeData(j, colStCena))

            vrednost = kolicina * cena

            ' Pretpostavljeni layout od StavkaStart:
            ' Col 0 = R.br.
            ' Col 1 = Broj prijemnice
            ' Col 2 = Klasa
            ' Col 3 = Kolicina
            ' Col 4 = Cena
            ' Col 5 = Vrednost
            startCell.Offset(outRow - 1, 0).value = outRow
            startCell.Offset(outRow - 1, 1).value = stavkeData(j, colStBrojPrij)
            startCell.Offset(outRow - 1, 2).value = stavkeData(j, colStKlasa)
            startCell.Offset(outRow - 1, 3).value = kolicina
            startCell.Offset(outRow - 1, 4).value = cena
            startCell.Offset(outRow - 1, 5).value = vrednost
        End If
    Next j

    If outRow = 0 Then
        Err.Raise vbObjectError + 1735, "PrintFaktura", _
                  "Nisu pronadene stavke za fakturu: " & fakturaID
    End If

    On Error Resume Next
    wsSablon.Range("UkupnoFaktura").value = data(fRow, colFakIznos)
    On Error GoTo EH

    wsSablon.PrintOut Copies:=1
    Exit Sub

EH:
    LogErr "PrintFaktura"
    Err.Raise Err.Number, "PrintFaktura", Err.Description
End Sub

Private Sub ClearFakturaStavkeArea(ByVal ws As Worksheet)
    On Error GoTo EH

    Dim startCell As Range
    Set startCell = ws.Range("StavkaStart")

    ' Cisti 50 redova × 6 kolona:
    ' R.br | BrojPrij | Klasa | Kolicina | Cena | Vrednost
    startCell.Resize(50, 6).ClearContents

    Exit Sub

EH:
    LogErr "ClearFakturaStavkeArea"
    Err.Raise Err.Number, "ClearFakturaStavkeArea", Err.Description
End Sub

Public Sub UpdateFakturaStatus(ByVal fakturaID As String)
    On Error GoTo EH

    If Trim$(fakturaID) = "" Then Exit Sub

    RequireColumnIndex TBL_FAKTURE, COL_FAK_ID, "UpdateFakturaStatus"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_IZNOS, "UpdateFakturaStatus"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_STATUS, "UpdateFakturaStatus"
    RequireColumnIndex TBL_FAKTURE, COL_FAK_DATUM_PLACANJA, "UpdateFakturaStatus"

    Dim fakturaIznosVal As Variant
    fakturaIznosVal = LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS)

    If Not IsNumeric(fakturaIznosVal) Then
        Err.Raise vbObjectError + 1720, "UpdateFakturaStatus", _
                  "Iznos fakture nije numericki: " & fakturaID
    End If

    Dim fakturaIznos As Double
    fakturaIznos = CDbl(fakturaIznosVal)

    Dim uplaceno As Double
    uplaceno = GetUplataForFaktura(fakturaID)

    Dim rows As Collection
    Set rows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)

    If rows.count = 0 Then
        Err.Raise vbObjectError + 1721, "UpdateFakturaStatus", _
                  "Faktura nije pronadena: " & fakturaID
    End If

    If uplaceno >= fakturaIznos And fakturaIznos > 0 Then
        RequireUpdateCell TBL_FAKTURE, rows(1), COL_FAK_STATUS, _
                          STATUS_PLACENO, "UpdateFakturaStatus"

        RequireUpdateCell TBL_FAKTURE, rows(1), COL_FAK_DATUM_PLACANJA, _
                          Date, "UpdateFakturaStatus"
    Else
        RequireUpdateCell TBL_FAKTURE, rows(1), COL_FAK_STATUS, _
                          STATUS_NEPLACENO, "UpdateFakturaStatus"

        RequireUpdateCell TBL_FAKTURE, rows(1), COL_FAK_DATUM_PLACANJA, _
                          Empty, "UpdateFakturaStatus"
    End If

    Exit Sub

EH:
    LogErr "UpdateFakturaStatus"
    Err.Raise Err.Number, "UpdateFakturaStatus", Err.Description
End Sub

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
