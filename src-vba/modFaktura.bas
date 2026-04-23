Attribute VB_Name = "modFaktura"

Option Explicit

' ============================================================
' modFaktura v2.1 – Rechnungserstellung
' GEÄNDERT: Basiert auf tblPrijemnica statt tblIsporuka
' Faktura-Betrag = Prijemnica.Kolicina × Prijemnica.Cena
' ============================================================

Public Function CreateFaktura_TX(ByVal kupacID As String, _
                                  ByVal stavke As Collection) As String
    Dim tx As New clsTransaction
    
    On Error GoTo EH
    tx.BeginTx
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_PRIJEMNICA
    tx.AddTableSnapshot TBL_NOVAC       ' ApplyAvansToFaktura
    
    CreateFaktura_TX = CreateFaktura(kupacID, stavke)
    
    tx.CommitTx
    Exit Function
EH:
    LogErr "CreateFaktura_TX"
    tx.RollbackTx
    MsgBox "Greska pri kreiranju fakture, promene vracene: " & Err.Description, _
           vbCritical, APP_NAME
    CreateFaktura_TX = ""
End Function

Public Function CreateFaktura(ByVal kupacID As String, _
                              ByVal stavke As Collection) As String
    ' stavke = Collection of Arrays (PrijemnicaID, Kolicina, Cena, Klasa, BrojPrijemnice)
    ' Returns: FakturaID oder ""
    
    If kupacID = "" Or stavke.count = 0 Then
        MsgBox "Kupac i stavke su obavezni!", vbExclamation, APP_NAME
        CreateFaktura = ""
        Exit Function
    End If
    
    Dim fakturaID As String
    fakturaID = GetNextID(TBL_FAKTURE, COL_FAK_ID, "FAK-")
    
    Dim brojFakture As String
    brojFakture = GenerateBrojFakture()
    
    ' Gesamtbetrag
    Dim ukupno As Double
    Dim s As Variant
    For Each s In stavke
        ukupno = ukupno + (CDbl(s(1)) * CDbl(s(2)))
    Next s
    
    ' Faktura-Kopf
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
    
    If AppendRow(TBL_FAKTURE, fakturaRow) = 0 Then
        CreateFaktura = ""
        Exit Function
    End If
    
    ' Stavke
    Dim stavkaID As String
    Dim stavkaNum As Long
    For Each s In stavke
        stavkaNum = stavkaNum + 1
        stavkaID = fakturaID & "-" & Format$(stavkaNum, "00")
        
        ' KulturaID aus Prijemnica über Zbirna/Otpremnica lookup
        ' Vereinfacht: Vrsta/Sorta direkt in Stavke
        Dim stavkaRow As Variant
        stavkaRow = Array(stavkaID, fakturaID, s(0), s(1), s(2), s(3), s(4))
        AppendRow TBL_FAKTURA_STAVKE, stavkaRow
        
        ' Prijemnica als fakturisano markieren
        Dim rows As Collection
        Set rows = FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, CStr(s(0)))
        If rows.count > 0 Then
            UpdateCell TBL_PRIJEMNICA, rows(1), COL_PRJ_FAKTURISANO, "Da"
            UpdateCell TBL_PRIJEMNICA, rows(1), COL_PRJ_FAKTURA_ID, fakturaID
        End If
    Next s
    ' Avans automatisch verrechnen
    ApplyAvansToFaktura kupacID, fakturaID
    
    CreateFaktura = fakturaID
End Function

Private Function GenerateBrojFakture() As String
    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)
    
    Dim maxNum As Long
    If Not IsEmpty(data) Then
        Dim colBroj As Long
        colBroj = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
        Dim i As Long
        For i = 1 To UBound(data, 1)
            If InStr(CStr(data(i, colBroj)), "/") > 0 Then
                Dim parts As Variant
                parts = Split(CStr(data(i, colBroj)), "/")
                Dim num As Long
                On Error Resume Next
                num = CLng(parts(0))
                On Error GoTo 0
                If num > maxNum Then maxNum = num
            End If
        Next i
    End If
    
    GenerateBrojFakture = CStr(maxNum + 1) & "/" & Year(Date)
End Function

Public Sub PrintFaktura(ByVal fakturaID As String)
    Dim rows As Collection
    Set rows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)
    If rows.count = 0 Then Exit Sub
    
    Dim data As Variant
    data = GetTableData(TBL_FAKTURE)
    Dim fRow As Long
    fRow = rows(1)
    
    Dim kupacID As String
    kupacID = CStr(data(fRow, GetColumnIndex(TBL_FAKTURE, COL_FAK_KUPAC)))
    Dim kupacNaziv As String
    kupacNaziv = CStr(LookupValue(TBL_KUPCI, "KupacID", kupacID, "Naziv"))
    
    ' Template füllen
    Dim wsSablon As Worksheet
    Set wsSablon = Nothing
    On Error Resume Next
    Set wsSablon = ThisWorkbook.Sheets("FakturaSablon")
    On Error GoTo 0                          ' ? OK, Sheet-Existenzprüfung
    If wsSablon Is Nothing Then
        MsgBox "FakturaSablon sheet ne postoji!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    With wsSablon
        .Range("BrojFakture").Value = data(fRow, GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ))
        .Range("DatumFakture").Value = data(fRow, GetColumnIndex(TBL_FAKTURE, COL_FAK_DATUM))
        .Range("KupacNaziv").Value = kupacNaziv
    End With
    
    ' Stavke eintragen
    Dim stavkeData As Variant
    stavkeData = GetTableData(TBL_FAKTURA_STAVKE)
    If Not IsEmpty(stavkeData) Then
        Dim colFakID As Long
        colFakID = GetColumnIndex(TBL_FAKTURA_STAVKE, "FakturaID")
        Dim outRow As Long
        
        Dim j As Long
        For j = 1 To UBound(stavkeData, 1)
            If CStr(stavkeData(j, colFakID)) = fakturaID Then
                outRow = outRow + 1
                wsSablon.Range("StavkaStart").Offset(outRow, 0).Value = _
                    stavkeData(j, GetColumnIndex(TBL_FAKTURA_STAVKE, "KulturaID"))
            End If
        Next j
    End If
    
    wsSablon.PrintOut Copies:=1
End Sub

Public Sub UpdateFakturaStatus(ByVal fakturaID As String)
    Dim uplaceno As Double
    uplaceno = GetUplataForFaktura(fakturaID)
    
    Dim fakturaIznos As Double
    fakturaIznos = CDbl(LookupValue(TBL_FAKTURE, COL_FAK_ID, fakturaID, COL_FAK_IZNOS))
    
    If uplaceno >= fakturaIznos Then
        Dim rows As Collection
        Set rows = FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)
        If rows.count > 0 Then
            UpdateCell TBL_FAKTURE, rows(1), COL_FAK_STATUS, STATUS_PLACENO
            UpdateCell TBL_FAKTURE, rows(1), COL_FAK_DATUM_PLACANJA, Date
        End If
    End If
End Sub
