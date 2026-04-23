Attribute VB_Name = "modIzvestaj"
Option Explicit

' ============================================================
' modIzvestaj v3.0 – Report Business Logic
' Alle Funktionen geben 2D-Arrays zurück
' Form ist nur noch für UI-Darstellung zuständig
' ============================================================

Public Enum IzvestajTip
    izvSaldo = 1
    izvOtkupljenaRoba = 2
    izvPrimljenaAmbalaza = 3
    izvIsplata = 4
    izvZbirniPoOM = 5
    izvManjak = 6
    izvProsecnaCena = 7
End Enum

Public Function ReportSaldoOM(ByVal stanicaID As String, _
                              ByVal datumOd As Date, _
                              ByVal datumDo As Date) As Variant
    ' Returns: 2D Array (Name, Kolicina, Vrednost, Novac, Saldo, Ambalaza)
    ' Letzte Zeile = UKUPNO
    
    Dim otkupData As Variant
    otkupData = GetOtkupByStation(stanicaID, datumOd, datumDo)
    If IsEmpty(otkupData) Then
        ReportSaldoOM = Empty
        Exit Function
    End If
    
    otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)  ' ? eine Zeile

    ' --- Otkup pro Kooperant aggregieren ---
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colKoop As Long, colKol As Long, colCena As Long, colAmb As Long
    colKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    colAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
    
    Dim i As Long
    For i = 1 To UBound(otkupData, 1)
        Dim key As String
        key = CStr(otkupData(i, colKoop))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(otkupData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkupData(i, colKol))
        If IsNumeric(otkupData(i, colKol)) And IsNumeric(otkupData(i, colCena)) Then
            vals(1) = vals(1) + CDbl(otkupData(i, colKol)) * CDbl(otkupData(i, colCena))
        End If
        If IsNumeric(otkupData(i, colAmb)) Then vals(2) = vals(2) + CLng(otkupData(i, colAmb))
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportSaldoOM = Empty
        Exit Function
    End If
    
    ' --- Novac pro Kooperant aus tblNovac ---
    Dim novacDict As Object
    Set novacDict = CreateObject("Scripting.Dictionary")
    
    Dim novacData As Variant
    novacData = GetTableData(TBL_NOVAC)
    novacData = ExcludeStornirano(novacData, TBL_NOVAC)
    
    If IsArray(novacData) Then
        Dim colNovKoop As Long, colNovIsplata As Long, colNovDatum As Long
        colNovKoop = GetColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID)
        colNovIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
        colNovDatum = GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM)
        
        Dim n As Long
        For n = 1 To UBound(novacData, 1)
            Dim koopID As String
            koopID = CStr(novacData(n, colNovKoop))
            If koopID <> "" Then
                If IsDate(novacData(n, colNovDatum)) Then
                    If CDate(novacData(n, colNovDatum)) >= datumOd And _
                       CDate(novacData(n, colNovDatum)) <= datumDo Then
                        ' Kooperant ins dict aufnehmen falls noch nicht vorhanden
                        If Not dict.Exists(koopID) Then
                            ' Prüfen ob Kooperant zu dieser Station gehört
                            Dim koopStation As String
                            koopStation = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "StanicaID"))
                            If koopStation = stanicaID Then
                                dict.Add koopID, Array(0#, 0#, 0#)
                            End If
                        End If
                        
                        If dict.Exists(koopID) Then
                            If Not novacDict.Exists(koopID) Then novacDict.Add koopID, 0#
                            If IsNumeric(novacData(n, colNovIsplata)) Then
                                novacDict(koopID) = novacDict(koopID) + CDbl(novacData(n, colNovIsplata))
                            End If
                        End If
                    End If
                End If
            End If
        Next n
    End If
    
        ' --- OM Avans berechnen (VOR dem ReDim) ---
    Dim omAvans As Double
    omAvans = 0
    
    If IsArray(novacData) Then
        Dim colNovTip As Long
        colNovTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
        Dim colNovOMID As Long
        colNovOMID = GetColumnIndex(TBL_NOVAC, COL_NOV_OM_ID)
        
        For n = 1 To UBound(novacData, 1)
            If CStr(novacData(n, colNovOMID)) = stanicaID Then
                If IsDate(novacData(n, colNovDatum)) Then
                    If CDate(novacData(n, colNovDatum)) >= datumOd And _
                       CDate(novacData(n, colNovDatum)) <= datumDo Then
                        If IsNumeric(novacData(n, colNovIsplata)) Then
                            Select Case CStr(novacData(n, colNovTip))
                                Case NOV_KES_FIRMA_OTKUPAC
                                    omAvans = omAvans + CDbl(novacData(n, colNovIsplata))
                                Case NOV_KES_OTKUPAC_KOOP
                                    omAvans = omAvans - CDbl(novacData(n, colNovIsplata))
                            End Select
                        End If
                    End If
                End If
            End If
        Next n
    End If
    
    Dim hasOMAvans As Boolean
    hasOMAvans = (omAvans > 0)
    
    ' --- Agrohemija pro Kooperant (Dict) ---
    Dim magData As Variant
    magData = GetTableData(TBL_MAGACIN)
    If Not IsEmpty(magData) Then magData = ExcludeStornirano(magData, TBL_MAGACIN)
    
    Dim colMagKoop As Long, colMagTip As Long, colMagVrednost As Long, colMagDat As Long
    If IsArray(magData) Then
        colMagKoop = GetColumnIndex(TBL_MAGACIN, COL_MAG_KOOP)
        colMagTip = GetColumnIndex(TBL_MAGACIN, COL_MAG_TIP)
        colMagVrednost = GetColumnIndex(TBL_MAGACIN, COL_MAG_VREDNOST)
        colMagDat = GetColumnIndex(TBL_MAGACIN, COL_MAG_DATUM)
    End If
    
    Dim agroKoopDict As Object
    Set agroKoopDict = CreateObject("Scripting.Dictionary")
    Dim agroBezStanica As Double  ' nerasporedjena Agrohemija (kein Kooperant)
    agroBezStanica = 0
    
    If Not IsEmpty(magData) Then
        If IsArray(magData) Then
            Dim m As Long
            For m = 1 To UBound(magData, 1)
                If CStr(magData(m, colMagTip)) = MAG_IZLAZ Then
                    If IsDate(magData(m, colMagDat)) Then
                        If CDate(magData(m, colMagDat)) >= datumOd And _
                           CDate(magData(m, colMagDat)) <= datumDo Then
                            If IsNumeric(magData(m, colMagVrednost)) Then
                                Dim magKoopID As String
                                magKoopID = CStr(magData(m, colMagKoop))
                                
                                If magKoopID <> "" And dict.Exists(magKoopID) Then
                                    If Not agroKoopDict.Exists(magKoopID) Then agroKoopDict.Add magKoopID, 0#
                                    agroKoopDict(magKoopID) = agroKoopDict(magKoopID) + CDbl(magData(m, colMagVrednost))
                                ElseIf magKoopID = "" Then
                                    agroBezStanica = agroBezStanica + CDbl(magData(m, colMagVrednost))
                                End If
                            End If
                        End If
                    End If
                End If
            Next m
        End If
    End If
    
    ' --- Ergebnis-Array: 7 Spalten ---
    ' Kooperant | Kolicina | Vrednost | Isplaceno | AgroZaduzenje | Saldo | Ambalaza
    
    Dim rowCount As Long
    rowCount = dict.count + 1  ' +UKUPNO
    If hasOMAvans Then rowCount = rowCount + 1
    If agroBezStanica > 0 Then rowCount = rowCount + 1
    
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 7)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totKol As Double, totVr As Double, totNov As Double
    Dim totAgro As Double, totAmb As Long
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        
        Dim novacSum As Double
        novacSum = 0
        If novacDict.Exists(keys(i)) Then novacSum = novacDict(keys(i))
        
        Dim agroSum As Double
        agroSum = 0
        If agroKoopDict.Exists(keys(i)) Then agroSum = agroKoopDict(keys(i))
        
        Dim ime As String, prezime As String
        ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", keys(i), "Ime"))
        prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", keys(i), "Prezime"))
        
        result(i + 1, 1) = ime & " " & prezime
        result(i + 1, 2) = vals(0)                          ' Kolicina
        result(i + 1, 3) = vals(1)                          ' Vrednost
        result(i + 1, 4) = novacSum                         ' Isplaceno
        result(i + 1, 5) = agroSum                          ' AgroZaduzenje
        result(i + 1, 6) = vals(1) - novacSum - agroSum     ' Saldo
        result(i + 1, 7) = vals(2)                          ' Ambalaza
        
        totKol = totKol + vals(0)
        totVr = totVr + vals(1)
        totNov = totNov + novacSum
        totAgro = totAgro + agroSum
        totAmb = totAmb + vals(2)
    Next i
    
    ' OM Avans (nerasporedjen)
    If hasOMAvans Then
        Dim omAvansRow As Long
        omAvansRow = dict.count + 1
        result(omAvansRow, 1) = "OM AVANS (nerasporedjen)"
        result(omAvansRow, 4) = omAvans
        totNov = totNov + omAvans
    End If
    
    ' Agrohemija (nerasporedjena — ohne Kooperant)
    If agroBezStanica > 0 Then
        Dim agroRow As Long
        agroRow = rowCount - 1
        result(agroRow, 1) = "AGROHEMIJA (nerasporedjena)"
        result(agroRow, 5) = agroBezStanica
        totAgro = totAgro + agroBezStanica
    End If
    
    ' UKUPNO
    result(rowCount, 1) = "UKUPNO"
    result(rowCount, 2) = totKol
    result(rowCount, 3) = totVr
    result(rowCount, 4) = totNov
    result(rowCount, 5) = totAgro
    result(rowCount, 6) = totVr - totNov - totAgro
    result(rowCount, 7) = totAmb
    
    ReportSaldoOM = result
End Function

Public Function ReportKarticaKooperanta(ByVal kooperantID As String, _
                                        ByVal datumOd As Date, _
                                        ByVal datumDo As Date) As Variant
    ' Returns: 2D Array
    ' (1)=Datum, (2)=BrojDok, (3)=BrojParcele, (4)=Opis,
    ' (5)=Zaduzenje, (6)=Razduzenje, (7)=Saldo
    
    Dim moves As New Collection
    
    Dim i As Long
    
    ' 1. Otkup = Zaduzenje
    Dim otkData As Variant
    otkData = GetTableData(TBL_OTKUP)
    If IsArray(otkData) Then
        otkData = ExcludeStornirano(otkData, TBL_OTKUP)
        If IsArray(otkData) Then
            Dim colOtkDat As Long, colOtkKoop As Long
            Dim colOtkKol As Long, colOtkCena As Long
            Dim colOtkVrsta As Long, colOtkKlasa As Long
            Dim colOtkBrDok As Long, colParcela As Long
            
            colOtkDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
            colOtkKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
            colOtkKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
            colOtkCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
            colOtkVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
            colOtkKlasa = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
            colOtkBrDok = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
            colParcela = GetColumnIndex(TBL_OTKUP, COL_OTK_PARCELA)
            
            For i = 1 To UBound(otkData, 1)
                If CStr(otkData(i, colOtkKoop)) = kooperantID Then
                    If IsDate(otkData(i, colOtkDat)) Then
                        Dim otkDatum As Date
                        otkDatum = CDate(otkData(i, colOtkDat))
                        
                        If otkDatum >= datumOd And otkDatum <= datumDo Then
                            Dim vr As Double
                            vr = 0
                            If IsNumeric(otkData(i, colOtkKol)) And IsNumeric(otkData(i, colOtkCena)) Then
                                vr = CDbl(otkData(i, colOtkKol)) * CDbl(otkData(i, colOtkCena))
                            End If
                            
                            Dim opis As String
                            opis = "Otkup " & CStr(otkData(i, colOtkVrsta)) & " " & _
                                   CStr(otkData(i, colOtkKlasa)) & " " & _
                                   Format$(val(otkData(i, colOtkKol)), "#,##0") & "kg"
                            
                            moves.Add Array( _
                                otkDatum, _
                                CStr(otkData(i, colOtkBrDok)), _
                                CStr(otkData(i, colParcela)), _
                                opis, _
                                vr, _
                                0#)
                        End If
                    End If
                End If
            Next i
        End If
    End If
    
    ' 2. Novac = Razduzenje
    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)
    If IsArray(novData) Then
        novData = ExcludeStornirano(novData, TBL_NOVAC)
        If IsArray(novData) Then
            Dim colNovDat As Long, colNovKoop As Long
            Dim colNovIsplata As Long, colNovTip As Long, colNovBrDok As Long
            
            colNovDat = GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM)
            colNovKoop = GetColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID)
            colNovIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
            colNovTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
            colNovBrDok = GetColumnIndex(TBL_NOVAC, COL_NOV_BROJ_DOK)
            
            Dim n As Long
            For n = 1 To UBound(novData, 1)
                If CStr(novData(n, colNovKoop)) = kooperantID Then
                    If IsDate(novData(n, colNovDat)) Then
                        Dim novDatum As Date
                        novDatum = CDate(novData(n, colNovDat))
                        
                        If novDatum >= datumOd And novDatum <= datumDo Then
                            Dim iznos As Double
                            iznos = 0
                            If IsNumeric(novData(n, colNovIsplata)) Then
                                iznos = CDbl(novData(n, colNovIsplata))
                            End If
                            
                            If iznos > 0 Then
                                Dim tipNovca As String
                                Dim novOpis As String
                                
                                tipNovca = CStr(novData(n, colNovTip))
                                Select Case tipNovca
                                    Case NOV_KES_OTKUPAC_KOOP: novOpis = "Kes Otkupac"
                                    Case NOV_VIRMAN_FIRMA_KOOP: novOpis = "Virman Firma"
                                    Case NOV_VIRMAN_AVANS_KOOP: novOpis = "Virman Avans"
                                    Case Else: novOpis = tipNovca
                                End Select
                                
                                moves.Add Array( _
                                    novDatum, _
                                    CStr(novData(n, colNovBrDok)), _
                                    "", _
                                    novOpis, _
                                    0#, _
                                    iznos)
                            End If
                        End If
                    End If
                End If
            Next n
        End If
    End If
    
    ' 3. Agrohemija = Razduzenje
    Dim magData As Variant
    magData = GetTableData(TBL_MAGACIN)
    If IsArray(magData) Then
        magData = ExcludeStornirano(magData, TBL_MAGACIN)
        If IsArray(magData) Then
            Dim colMagDat As Long, colMagKoop As Long, colMagTip As Long
            Dim colMagVrednost As Long, colMagArtikal As Long, colMagBrDok As Long
            
            colMagDat = GetColumnIndex(TBL_MAGACIN, COL_MAG_DATUM)
            colMagKoop = GetColumnIndex(TBL_MAGACIN, COL_MAG_KOOP)
            colMagTip = GetColumnIndex(TBL_MAGACIN, COL_MAG_TIP)
            colMagVrednost = GetColumnIndex(TBL_MAGACIN, COL_MAG_VREDNOST)
            colMagArtikal = GetColumnIndex(TBL_MAGACIN, COL_MAG_ARTIKAL)
            colMagBrDok = GetColumnIndex(TBL_MAGACIN, COL_MAG_BR_DOK)
            
            Dim m As Long
            For m = 1 To UBound(magData, 1)
                If CStr(magData(m, colMagKoop)) = kooperantID Then
                    If CStr(magData(m, colMagTip)) = MAG_IZLAZ Then
                        If IsDate(magData(m, colMagDat)) Then
                            Dim magDatum As Date
                            magDatum = CDate(magData(m, colMagDat))
                            
                            If magDatum >= datumOd And magDatum <= datumDo Then
                                Dim magVr As Double
                                magVr = 0
                                If IsNumeric(magData(m, colMagVrednost)) Then
                                    magVr = CDbl(magData(m, colMagVrednost))
                                End If
                                
                                If magVr > 0 Then
                                    Dim artNaziv As String
                                    artNaziv = CStr(LookupValue(TBL_ARTIKLI, COL_ART_ID, _
                                                  CStr(magData(m, colMagArtikal)), COL_ART_NAZIV))
                                    
                                    moves.Add Array( _
                                        magDatum, _
                                        CStr(magData(m, colMagBrDok)), _
                                        "", _
                                        "Agrohemija " & artNaziv, _
                                        0#, _
                                        magVr)
                                End If
                            End If
                        End If
                    End If
                End If
            Next m
        End If
    End If
    
    If moves.count = 0 Then
        ReportKarticaKooperanta = Empty
        Exit Function
    End If
    
    ' Prebaci u niz za sortiranje:
    ' 1 Datum, 2 BrojDok, 3 BrojParcele, 4 Opis, 5 Zaduzenje, 6 Razduzenje
    Dim arr() As Variant
    ReDim arr(1 To moves.count, 1 To 6)
    
    For i = 1 To moves.count
        Dim mv As Variant
        mv = moves(i)
        arr(i, 1) = mv(0)
        arr(i, 2) = mv(1)
        arr(i, 3) = mv(2)
        arr(i, 4) = mv(3)
        arr(i, 5) = mv(4)
        arr(i, 6) = mv(5)
    Next i
    
    ' Sort po datumu
    Dim swapped As Boolean
    Dim temp As Variant
    Dim s As Long, c As Long
    
    Do
        swapped = False
        For s = 1 To UBound(arr, 1) - 1
            If CDate(arr(s, 1)) > CDate(arr(s + 1, 1)) Then
                For c = 1 To 6
                    temp = arr(s, c)
                    arr(s, c) = arr(s + 1, c)
                    arr(s + 1, c) = temp
                Next c
                swapped = True
            End If
        Next s
    Loop While swapped
    
    ' Rezultat: + saldo
    Dim result() As Variant
    ReDim result(1 To UBound(arr, 1) + 1, 1 To 7)
    
    Dim runSaldo As Double
    Dim totZad As Double, totRaz As Double
    
    For i = 1 To UBound(arr, 1)
        result(i, 1) = arr(i, 1) ' Datum
        result(i, 2) = arr(i, 2) ' BrojDok
        result(i, 3) = arr(i, 3) ' BrojParcele
        result(i, 4) = arr(i, 4) ' Opis
        result(i, 5) = arr(i, 5) ' Zaduzenje
        result(i, 6) = arr(i, 6) ' Razduzenje
        
        runSaldo = runSaldo + arr(i, 5) - arr(i, 6)
        result(i, 7) = runSaldo
        
        totZad = totZad + arr(i, 5)
        totRaz = totRaz + arr(i, 6)
    Next i
    
    Dim ukRow As Long
    ukRow = UBound(arr, 1) + 1
    result(ukRow, 4) = "UKUPNO"
    result(ukRow, 5) = totZad
    result(ukRow, 6) = totRaz
    result(ukRow, 7) = totZad - totRaz
    
    ReportKarticaKooperanta = result
End Function

Public Sub PrintKarticaPDF(ByVal kooperantID As String, _
                           ByVal datumOd As Date, ByVal datumDo As Date)
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("KarticaSablon")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "KarticaSablon sheet ne postoji!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Kartica-Daten holen
    Dim data As Variant
    data = ReportKarticaKooperanta(kooperantID, datumOd, datumDo)
    If IsEmpty(data) Then
        MsgBox "Nema podataka za ovog kooperanta!", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Header
    Dim ime As String, prezime As String
    ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime"))
    prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"))
    Dim bpg As String
    bpg = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_BPG))
    
    ws.Range("KartKoop").Value = ime & " " & prezime & " (" & kooperantID & ")"
    ws.Range("KartBPG").Value = bpg
    ws.Range("KartPeriod").Value = Format$(datumOd, "DD.MM.YYYY") & " - " & Format$(datumDo, "DD.MM.YYYY")
    
    Const NUM_COLS As Long = 6
    
    ' Alte Daten löschen
    Dim startRow As Long
    startRow = ws.Range("KartStart").row
    Dim lastRow As Long
    lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= startRow Then
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, NUM_COLS)).ClearContents
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, NUM_COLS)).ClearFormats
    End If
    
    ' Daten einfügen
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim outRow As Long
        outRow = startRow + i - 1
        
        ' Datum
        If IsDate(data(i, 1)) Then
            ws.cells(outRow, 1).Value = Format$(CDate(data(i, 1)), "DD.MM.YYYY")
        Else
            ws.cells(outRow, 1).Value = CStr(IIf(IsEmpty(data(i, 1)), "", data(i, 1)))
        End If
        
        ' BrojDok
        ws.cells(outRow, 2).Value = CStr(IIf(IsEmpty(data(i, 2)), "", data(i, 2)))
        
        ' Opis
        ws.cells(outRow, 3).Value = CStr(IIf(IsEmpty(data(i, 3)), "", data(i, 3)))
        
        ' Zaduzenje
        If IsNumeric(data(i, 4)) And Not IsEmpty(data(i, 4)) Then
            If CDbl(data(i, 4)) > 0 Then ws.cells(outRow, 4).Value = CDbl(data(i, 4))
        End If
        
        ' Razduzenje
        If IsNumeric(data(i, 5)) And Not IsEmpty(data(i, 5)) Then
            If CDbl(data(i, 5)) > 0 Then ws.cells(outRow, 5).Value = CDbl(data(i, 5))
        End If
        
        ' Saldo
        If IsNumeric(data(i, 6)) And Not IsEmpty(data(i, 6)) Then
            ws.cells(outRow, 6).Value = CDbl(data(i, 6))
        End If
    Next i
    
    ' Formatierung Datenbereich
    Dim dataRows As Long
    dataRows = UBound(data, 1)
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.cells(startRow, 1), ws.cells(startRow + dataRows - 1, NUM_COLS))
    
    ' Rahmen
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Zahlenformat D-F
    ws.Range(ws.cells(startRow, 4), ws.cells(startRow + dataRows - 1, 6)).NumberFormat = "#,##0.00"
    
    ' Alternierende Farbe
    Dim r As Long
    For r = 0 To dataRows - 1
        If r Mod 2 = 1 Then
            ws.Range(ws.cells(startRow + r, 1), _
                     ws.cells(startRow + r, NUM_COLS)).Interior.Color = RGB(217, 225, 242)
        End If
    Next r
    
    ' UKUPNO Zeile fett
    Dim ukRow As Long
    ukRow = startRow + dataRows - 1
    ws.Range(ws.cells(ukRow, 1), ws.cells(ukRow, NUM_COLS)).Font.Bold = True
    ws.Range(ws.cells(ukRow, 1), ws.cells(ukRow, NUM_COLS)).Interior.Color = RGB(68, 114, 196)
    ws.Range(ws.cells(ukRow, 1), ws.cells(ukRow, NUM_COLS)).Font.Color = RGB(255, 255, 255)
    
    ' Fusszeile
    Dim footRow As Long
    footRow = ukRow + 2
    ws.cells(footRow, 1).Value = "Datum stampe: " & Format$(Date, "DD.MM.YYYY")
    ws.cells(footRow + 1, 1).Value = "Potpis kooperanta: ___________"
    ws.cells(footRow + 1, 4).Value = "Potpis firme: ___________"
    
    ' PDF Export
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Kartica_" & Replace(kooperantID, "-", "") & "_" & _
              Format$(datumOd, "YYYYMMDD") & "-" & Format$(datumDo, "YYYYMMDD") & ".pdf"
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, _
                           OpenAfterPublish:=True
    
    Application.ScreenUpdating = True
End Sub

' ============================================================
' KUPCI
' ============================================================
Public Function ReportSaldoKupci(ByVal kupacID As String, _
                                 ByVal datumOd As Date, _
                                 ByVal datumDo As Date) As Variant
    ' Returns: 2D Array (Vrsta, Kolicina, Cena, Vrednost, Novac, Saldo, Ambalaza)
    ' Vorletzte Zeile = (Nerasporedeno) wenn Avans existiert
    ' Letzte Zeile = UKUPNO
    
    Dim prijData As Variant
    prijData = GetPrijemniceByKupac(kupacID, datumOd, datumDo)
    If IsEmpty(prijData) Then
        ReportSaldoKupci = Empty
        Exit Function
    End If
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    
    ' --- Prijemnice pro VrstaVoca aggregieren ---
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colVrsta As Long, colKol As Long, colCena As Long, colAmb As Long
    colVrsta = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    colAmb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    
    Dim i As Long
    For i = 1 To UBound(prijData, 1)
        Dim key As String
        key = CStr(prijData(i, colVrsta))
        If key = "" Then key = "(Nepoznato)"
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(prijData(i, colKol)) Then vals(0) = vals(0) + CDbl(prijData(i, colKol))
        If IsNumeric(prijData(i, colCena)) Then vals(1) = CDbl(prijData(i, colCena))
        If IsNumeric(prijData(i, colKol)) And IsNumeric(prijData(i, colCena)) Then
            vals(2) = vals(2) + CDbl(prijData(i, colKol)) * CDbl(prijData(i, colCena))
        End If
        If IsNumeric(prijData(i, colAmb)) Then vals(3) = vals(3) + CLng(prijData(i, colAmb))
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportSaldoKupci = Empty
        Exit Function
    End If
    
    ' --- Novac pro Vrsta ---
    Dim novacDict As Object
    Set novacDict = GetUplataByVrsta(kupacID, datumOd, datumDo)
    
    ' --- Gesamt-Novac (für UKUPNO Saldo) ---
    Dim novacTotal As Double
    Dim novacData As Variant
    novacData = GetTableData(TBL_NOVAC)
    novacData = ExcludeStornirano(novacData, TBL_NOVAC)
    
    If IsArray(novacData) Then
        Dim colNovPartnerID As Long, colNovUplata As Long, colNovDatum As Long
        colNovPartnerID = GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID)
        colNovUplata = GetColumnIndex(TBL_NOVAC, COL_NOV_UPLATA)
        colNovDatum = GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM)
        
        Dim n As Long
        For n = 1 To UBound(novacData, 1)
            If CStr(novacData(n, colNovPartnerID)) = kupacID Then
                If IsDate(novacData(n, colNovDatum)) Then
                    If CDate(novacData(n, colNovDatum)) >= datumOd And _
                       CDate(novacData(n, colNovDatum)) <= datumDo Then
                        If IsNumeric(novacData(n, colNovUplata)) Then
                            novacTotal = novacTotal + CDbl(novacData(n, colNovUplata))
                        End If
                    End If
                End If
            End If
        Next n
    End If
    
    ' --- Ergebnis-Array ---
    Dim hasNerasporedeno As Boolean
    hasNerasporedeno = novacDict.Exists("(Nerasporedeno)")
    
    Dim rowCount As Long
    rowCount = dict.count + 1  ' +1 UKUPNO
    If hasNerasporedeno Then rowCount = rowCount + 1
    
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 7)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totKol As Double, totVr As Double, totNov As Double, totAmb As Long
    Dim idx As Long
    
    For i = 0 To dict.count - 1
        idx = i + 1
        vals = dict(keys(i))
        
        Dim novacVrsta As Double
        novacVrsta = 0
        If novacDict.Exists(keys(i)) Then novacVrsta = novacDict(keys(i))
        
        result(idx, 1) = keys(i)              ' Vrsta
        result(idx, 2) = vals(0)              ' Kolicina
        result(idx, 3) = vals(1)              ' Cena (letzte)
        result(idx, 4) = vals(2)              ' Vrednost
        result(idx, 5) = novacVrsta           ' Novac pro Vrsta
        result(idx, 6) = vals(2) - novacVrsta ' Saldo pro Vrsta
        result(idx, 7) = vals(3)              ' Ambalaza
        
        totKol = totKol + vals(0)
        totVr = totVr + vals(2)
        totNov = totNov + novacVrsta
        totAmb = totAmb + vals(3)
    Next i
    
    ' Nerasporedeno
    If hasNerasporedeno Then
        idx = dict.count + 1
        result(idx, 1) = "(Nerasporedeno)"
        result(idx, 5) = novacDict("(Nerasporedeno)")
        totNov = totNov + novacDict("(Nerasporedeno)")
    End If
    
    ' UKUPNO
    result(rowCount, 1) = "UKUPNO"
    result(rowCount, 2) = totKol
    result(rowCount, 3) = ""       ' Keine Durchschnittscena
    result(rowCount, 4) = totVr
    result(rowCount, 5) = novacTotal
    result(rowCount, 6) = totVr - novacTotal
    result(rowCount, 7) = totAmb
    
    ReportSaldoKupci = result
End Function


Public Function ReportIsplata(ByVal entitetTip As String, _
                              ByVal entitetID As String, _
                              ByVal datumOd As Date, _
                              ByVal datumDo As Date) As Variant
    ' Returns: 2D Array pro Kooperant
    ' Spalten: Kooperant | KesOtkupac | VirmanFirma | VirmanAvans | Ukupno
    ' + Summary: OM Avans primljeno | OM Avans podeljeno | Kod Otkupca
    
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then
        ReportIsplata = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_NOVAC)
    If Not IsArray(data) Then
        ReportIsplata = Empty
        Exit Function
    End If
    
    Dim colDatum As Long, colOMID As Long, colTip As Long
    Dim colIsplata As Long, colKoopID As Long, colPartnerID As Long
    colDatum = GetColumnIndex(TBL_NOVAC, COL_NOV_DATUM)
    colOMID = GetColumnIndex(TBL_NOVAC, COL_NOV_OM_ID)
    colTip = GetColumnIndex(TBL_NOVAC, COL_NOV_TIP)
    colIsplata = GetColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA)
    colKoopID = GetColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID)
    colPartnerID = GetColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID)
    
    ' Dicts: KooperantID ? Array(KesOtkupac, VirmanFirma, VirmanAvans)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim totalOMAvans As Double
    Dim totalKesOtkupac As Double
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Not IsDate(data(i, colDatum)) Then GoTo NextRow
        If CDate(data(i, colDatum)) < datumOd Or CDate(data(i, colDatum)) > datumDo Then GoTo NextRow
        
        Dim match As Boolean: match = False
        Select Case entitetTip
            Case "OM":    match = (CStr(data(i, colOMID)) = entitetID)
            Case "Kupac": match = (CStr(data(i, colPartnerID)) = entitetID)
        End Select
        If Not match Then GoTo NextRow
        
        Dim tipNovca As String
        tipNovca = CStr(data(i, colTip))
        Dim iznos As Double: iznos = 0
        If IsNumeric(data(i, colIsplata)) Then iznos = CDbl(data(i, colIsplata))
        If iznos <= 0 Then GoTo NextRow
        
        Dim koopID As String
        koopID = CStr(data(i, colKoopID))
        
        ' OM Avans (Firma ? Otkupac) — kein Kooperant
        If tipNovca = NOV_KES_FIRMA_OTKUPAC Then
            totalOMAvans = totalOMAvans + iznos
            GoTo NextRow
        End If
        
        ' Kooperant-bezogene Isplate
        If koopID = "" Then GoTo NextRow
        
        If Not dict.Exists(koopID) Then dict.Add koopID, Array(0#, 0#, 0#)
        Dim vals As Variant
        vals = dict(koopID)
        
        Select Case tipNovca
            Case NOV_KES_OTKUPAC_KOOP
                vals(0) = vals(0) + iznos
                totalKesOtkupac = totalKesOtkupac + iznos
            Case NOV_VIRMAN_FIRMA_KOOP
                vals(1) = vals(1) + iznos
            Case NOV_VIRMAN_AVANS_KOOP
                vals(2) = vals(2) + iznos
        End Select
        
        dict(koopID) = vals
NextRow:
    Next i
    
    If dict.count = 0 And totalOMAvans = 0 Then
        ReportIsplata = Empty
        Exit Function
    End If
    
    ' Ergebnis: Kooperanten + UKUPNO + 3 Summary-Zeilen
    Dim rowCount As Long
    rowCount = dict.count + 4  ' UKUPNO + 3 Kontrolle
    
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 5)
    
    Dim keys As Variant
    If dict.count > 0 Then keys = dict.keys
    Dim totKes As Double, totVirman As Double, totAvans As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        
        Dim ime As String, prezime As String
        ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", keys(i), "Ime"))
        prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", keys(i), "Prezime"))
        
        result(i + 1, 1) = ime & " " & prezime
        result(i + 1, 2) = vals(0)                          ' KesOtkupac
        result(i + 1, 3) = vals(1)                          ' VirmanFirma
        result(i + 1, 4) = vals(2)                          ' VirmanAvans
        result(i + 1, 5) = vals(0) + vals(1) + vals(2)      ' Ukupno
        
        totKes = totKes + vals(0)
        totVirman = totVirman + vals(1)
        totAvans = totAvans + vals(2)
    Next i
    
    ' UKUPNO
    Dim ukRow As Long
    ukRow = dict.count + 1
    result(ukRow, 1) = "UKUPNO"
    result(ukRow, 2) = totKes
    result(ukRow, 3) = totVirman
    result(ukRow, 4) = totAvans
    result(ukRow, 5) = totKes + totVirman + totAvans
    
    ' Kontrolle
    result(ukRow + 1, 1) = "OM Avans (primljeno)"
    result(ukRow + 1, 5) = totalOMAvans
    
    result(ukRow + 2, 1) = "OM Avans (podeljeno)"
    result(ukRow + 2, 5) = totalKesOtkupac
    
    result(ukRow + 3, 1) = "Kod Otkupca"
    result(ukRow + 3, 5) = totalOMAvans - totalKesOtkupac
    
    ReportIsplata = result
End Function

Public Function ReportOtkupRoba(ByVal entitetTip As String, _
                                ByVal entitetID As String, _
                                ByVal datumOd As Date, _
                                ByVal datumDo As Date) As Variant
    ' Returns: 2D Array (Col1, Col2, Kolicina, Vrednost)
    '   OM:    Datum, BrojOtp+Vrsta, Kg, RSD
    '   Kupac: Nr, Vrsta, Kg, RSD
    '   Vozac: Nr, Vrsta, Kg, RSD
    ' Letzte Zeile = UKUPNO
    
    If entitetTip = "OM" Then
        ReportOtkupRoba = ReportOtkupRobaOM(entitetID, datumOd, datumDo)
    ElseIf entitetTip = "Kupac" Then
        ReportOtkupRoba = ReportOtkupRobaKupac(entitetID, datumOd, datumDo)
    ElseIf entitetTip = "Vozac" Then
        ReportOtkupRoba = ReportOtkupRobaVozac(entitetID, datumOd, datumDo)
    Else
        ReportOtkupRoba = Empty
    End If
End Function

Private Function ReportOtkupRobaOM(ByVal stanicaID As String, _
                                   ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    Dim otpData As Variant
    otpData = GetOtpremniceByStation(stanicaID, datumOd, datumDo)
    If IsEmpty(otpData) Then
        ReportOtkupRobaOM = Empty
        Exit Function
    End If
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    If Not IsArray(otpData) Then
        ReportOtkupRobaOM = Empty
        Exit Function
    End If
    
    Dim colVrsta As Long, colKol As Long, colBrOtp As Long
    Dim colDatum As Long, colKlasa As Long, colVozac As Long
    Dim colOtpID As Long, colBrZbirne As Long
    colVrsta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA)
    colKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    colBrOtp = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    colDatum = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM)
    colKlasa = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA)
    colVozac = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC)
    colOtpID = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    colBrZbirne = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    
    ' --- Otkup-Summen pro OtpremnicaID ---
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    Dim otkupDict As Object
    Set otkupDict = CreateObject("Scripting.Dictionary")
    
    If IsArray(otkupData) Then
        otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
        If IsArray(otkupData) Then
            Dim colOtkOtpID As Long, colOtkKol As Long
            colOtkOtpID = GetColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID)
            colOtkKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
            
            Dim j As Long
            For j = 1 To UBound(otkupData, 1)
                Dim otpKey As String
                otpKey = CStr(otkupData(j, colOtkOtpID))
                If otpKey <> "" Then
                    If Not otkupDict.Exists(otpKey) Then otkupDict.Add otpKey, 0#
                    If IsNumeric(otkupData(j, colOtkKol)) Then
                        otkupDict(otpKey) = otkupDict(otpKey) + CDbl(otkupData(j, colOtkKol))
                    End If
                End If
            Next j
        End If
    End If
    
    ' --- Manjak pro Zbirna ---
    Dim manjakDict As Object
    Set manjakDict = BuildManjakDict()
    
    ' --- Ergebnis ---
    Dim rowCount As Long
    rowCount = UBound(otpData, 1)
    
    Dim result() As Variant
    ReDim result(1 To rowCount + 1, 1 To 10)
    
    Dim totOtp As Double, totBlokovi As Double
    Dim totRazlika As Double, totManjak As Double
    Dim i As Long
    
    For i = 1 To rowCount
        Dim kgOtp As Double: kgOtp = 0
        If IsNumeric(otpData(i, colKol)) Then kgOtp = CDbl(otpData(i, colKol))
        
        Dim thisOtpID As String
        thisOtpID = CStr(otpData(i, colOtpID))
        
        Dim kgBlokovi As Double: kgBlokovi = 0
        If otkupDict.Exists(thisOtpID) Then kgBlokovi = otkupDict(thisOtpID)
        
        Dim razlika As Double
        razlika = kgBlokovi - kgOtp
        
        ' Manjak proportional berechnen
        Dim thisBrZbirne As String
        thisBrZbirne = CStr(otpData(i, colBrZbirne))
        
        Dim manjak As Double: manjak = 0
        Dim manjakPct As Double: manjakPct = 0
        Dim mVals As Variant
        If manjakDict.Exists(thisBrZbirne) Then
            mVals = manjakDict(thisBrZbirne)
            Dim zbirnaTotal As Double: zbirnaTotal = mVals(0)
            Dim prijTotal As Double: prijTotal = mVals(1)
            
            If zbirnaTotal > 0 And prijTotal > 0 Then
                ' Proportionaler Manjak: (Zbirna - Prijemnica) × (OtpKg / ZbirnaKg)
                Dim ukupnoManjak As Double
                ukupnoManjak = zbirnaTotal - prijTotal
                manjak = ukupnoManjak * (kgOtp / zbirnaTotal)
                manjakPct = ukupnoManjak / zbirnaTotal * 100
            End If
        End If
        
        ' Vozac Name
        Dim vozID As String
        vozID = CStr(otpData(i, colVozac))
        Dim vozNaziv As String
        If vozID <> "" Then
            vozNaziv = CStr(LookupValue(TBL_VOZACI, "VozacID", vozID, "Ime")) & " " & _
                       CStr(LookupValue(TBL_VOZACI, "VozacID", vozID, "Prezime"))
        Else
            vozNaziv = ""
        End If
        
        result(i, 1) = CDate(otpData(i, colDatum))
        result(i, 2) = CStr(otpData(i, colBrOtp))
        result(i, 3) = CStr(otpData(i, colVrsta))
        result(i, 4) = CStr(otpData(i, colKlasa))
        result(i, 5) = vozNaziv
        result(i, 6) = kgOtp
        result(i, 7) = kgBlokovi
        result(i, 8) = razlika
        result(i, 9) = manjak
        result(i, 10) = manjakPct
        
        totOtp = totOtp + kgOtp
        totBlokovi = totBlokovi + kgBlokovi
        totRazlika = totRazlika + razlika
        totManjak = totManjak + manjak
    Next i
    
    ' UKUPNO
    result(rowCount + 1, 2) = "UKUPNO"
    result(rowCount + 1, 6) = totOtp
    result(rowCount + 1, 7) = totBlokovi
    result(rowCount + 1, 8) = totRazlika
    result(rowCount + 1, 9) = totManjak
    If totOtp > 0 Then result(rowCount + 1, 10) = totManjak / totOtp * 100
    
    ReportOtkupRobaOM = result
End Function

Private Function ReportOtkupRobaKupac(ByVal kupacID As String, _
                                      ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    ' Aggregiert pro VrstaVoca
    Dim prijData As Variant
    prijData = GetPrijemniceByKupac(kupacID, datumOd, datumDo)
    If IsEmpty(prijData) Then
        ReportOtkupRobaKupac = Empty
        Exit Function
    End If
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colVrsta As Long, colKol As Long, colCena As Long
    colVrsta = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    
    Dim i As Long
    For i = 1 To UBound(prijData, 1)
        Dim key As String
        key = CStr(prijData(i, colVrsta))
        If key = "" Then key = "(Nepoznato)"
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(prijData(i, colKol)) Then vals(0) = vals(0) + CDbl(prijData(i, colKol))
        If IsNumeric(prijData(i, colKol)) And IsNumeric(prijData(i, colCena)) Then
            vals(1) = vals(1) + CDbl(prijData(i, colKol)) * CDbl(prijData(i, colCena))
        End If
        dict(key) = vals
    Next i
    
    ReportOtkupRobaKupac = DictToResultArray(dict)
End Function

Private Function ReportOtkupRobaVozac(ByVal vozacID As String, _
                                      ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    ' Aggregiert pro VrstaVoca aus Otpremnice
    Dim otpData As Variant
    otpData = GetVozacDokumenta(vozacID, datumOd, datumDo)
    If IsEmpty(otpData) Then
        ReportOtkupRobaVozac = Empty
        Exit Function
    End If
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colVrsta As Long, colKol As Long, colCena As Long
    colVrsta = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA)
    colKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    colCena = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA)
    
    Dim i As Long
    For i = 1 To UBound(otpData, 1)
        Dim key As String
        key = CStr(otpData(i, colVrsta))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(otpData(i, colKol)) Then vals(0) = vals(0) + CDbl(otpData(i, colKol))
        If IsNumeric(otpData(i, colKol)) And IsNumeric(otpData(i, colCena)) Then
            vals(1) = vals(1) + CDbl(otpData(i, colKol)) * CDbl(otpData(i, colCena))
        End If
        dict(key) = vals
    Next i
    
    ReportOtkupRobaVozac = DictToResultArray(dict)
End Function

' ============================================================
' AMBALAZA REPORT
' ============================================================

Public Function ReportAmbalaza(ByVal entitetTip As String, _
                               ByVal entitetID As String, _
                               ByVal datumOd As Date, _
                               ByVal datumDo As Date, _
                               ByVal zbirni As Boolean) As Variant
    ' Zbirni Returns: 2D Array (Tip, "", "", "", Ulaz, Izlaz)
    ' Einzeln Returns: 2D Array (Datum, Mesto, Tip, DokID, Ulaz, Izlaz)
    ' Letzte Zeile = UKUPNO
    
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then
        ReportAmbalaza = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_AMBALAZA)  ' ? NEU
    If IsEmpty(data) Then
        ReportAmbalaza = Empty
        Exit Function
    End If
    ' --- Filter aufbauen ---
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    If entitetTip = "OM" Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET), "=", entitetID
        filters.Add fp
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP), "=", "Stanica"
        filters.Add fp
    ElseIf entitetTip = "Kupac" Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET), "=", entitetID
        filters.Add fp
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP), "=", "Kupac"
        filters.Add fp
    ElseIf entitetTip = "Vozac" Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC), "=", entitetID
        filters.Add fp
    End If
    
    Dim filtered As Variant
    filtered = FilterArray(data, filters)
    If IsEmpty(filtered) Then
        ReportAmbalaza = Empty
        Exit Function
    End If
    
    Dim colTip As Long, colKol As Long, colSmer As Long
    Dim colDokID As Long, colDatum As Long
    Dim colEntitet As Long, colEntTip As Long
    colTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    colKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    colSmer = GetColumnIndex(TBL_AMBALAZA, COL_AMB_SMER)
    colDokID = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    colDatum = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    colEntitet = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    colEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    
    If zbirni Then
        ReportAmbalaza = ReportAmbalazeZbirni(filtered, colTip, colKol, colSmer)
    Else
        ReportAmbalaza = ReportAmbalazePojedinacni(filtered, colDatum, colEntitet, colEntTip, _
                                                    colTip, colDokID, colKol, colSmer)
    End If
End Function

Private Function ReportAmbalazeZbirni(ByVal filtered As Variant, _
                                      ByVal colTip As Long, ByVal colKol As Long, _
                                      ByVal colSmer As Long) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim key As String
        key = CStr(filtered(i, colTip))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        Dim kol As Long: kol = 0
        If IsNumeric(filtered(i, colKol)) Then kol = CLng(filtered(i, colKol))
        If CStr(filtered(i, colSmer)) = "Ulaz" Then
            vals(0) = vals(0) + kol
        Else
            vals(1) = vals(1) + kol
        End If
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportAmbalazeZbirni = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 6)
    
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = ""
        result(i + 1, 3) = ""
        result(i + 1, 4) = ""
        result(i + 1, 5) = vals(0)
        result(i + 1, 6) = vals(1)
    Next i
    
    ReportAmbalazeZbirni = result
End Function

Private Function ReportAmbalazePojedinacni(ByVal filtered As Variant, _
                                            ByVal colDatum As Long, ByVal colEntitet As Long, _
                                            ByVal colEntTip As Long, ByVal colTip As Long, _
                                            ByVal colDokID As Long, ByVal colKol As Long, _
                                            ByVal colSmer As Long) As Variant
    Dim rowCount As Long
    rowCount = UBound(filtered, 1)
    
    Dim result() As Variant
    ReDim result(1 To rowCount + 1, 1 To 6)  ' +1 UKUPNO
    
    Dim totalUlaz As Long, totalIzlaz As Long
    Dim i As Long
    
    For i = 1 To rowCount
        Dim kol As Long: kol = 0
        If IsNumeric(filtered(i, colKol)) Then kol = CLng(filtered(i, colKol))
        
        ' Mesto-Name auflösen
        Dim entID As String: entID = CStr(filtered(i, colEntitet))
        Dim entTipVal As String: entTipVal = CStr(filtered(i, colEntTip))
        
        result(i, 1) = CDate(filtered(i, colDatum))
        result(i, 2) = ResolveEntitetName(entID, entTipVal)
        result(i, 3) = CStr(filtered(i, colTip))
        result(i, 4) = CStr(filtered(i, colDokID))
        
        If CStr(filtered(i, colSmer)) = "Ulaz" Then
            result(i, 5) = kol
            result(i, 6) = ""
            totalUlaz = totalUlaz + kol
        Else
            result(i, 5) = ""
            result(i, 6) = kol
            totalIzlaz = totalIzlaz + kol
        End If
    Next i
    
    ' UKUPNO
    result(rowCount + 1, 1) = "UKUPNO"
    result(rowCount + 1, 2) = ""
    result(rowCount + 1, 3) = ""
    result(rowCount + 1, 4) = "Saldo: " & Format$(totalUlaz - totalIzlaz, "#,##0")
    result(rowCount + 1, 5) = totalUlaz
    result(rowCount + 1, 6) = totalIzlaz
    
    ReportAmbalazePojedinacni = result
End Function

' ============================================================
' PROSECNA CENA i MANJAK
' ============================================================

Public Function ReportProsecnaCena(ByVal entitetTip As String, _
                                   ByVal entitetID As String, _
                                   ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    ' Returns: 2D Array (Vrsta, Kolicina, Vrednost, ProsecnaCena)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    
    If entitetTip = "Kupac" Then
        Dim prijData As Variant
        prijData = GetPrijemniceByKupac(entitetID, datumOd, datumDo)
        If IsEmpty(prijData) Then
            ReportProsecnaCena = Empty
            Exit Function
        End If
        prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
        
        Dim vrstaCache As Object
        Set vrstaCache = BuildZbirnaVrstaCache()
        
        Dim colBrZbr As Long, colPrijKol As Long, colPrijCena As Long
        colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        colPrijKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
        colPrijCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
        
        For i = 1 To UBound(prijData, 1)
            Dim vrsta As String
            vrsta = GetVrstaFromCache(vrstaCache, CStr(prijData(i, colBrZbr)))
            If vrsta = "" Then vrsta = "(Nepoznato)"
            
            If Not dict.Exists(vrsta) Then dict.Add vrsta, Array(0#, 0#)
            Dim vals As Variant
            vals = dict(vrsta)
            If IsNumeric(prijData(i, colPrijKol)) Then vals(0) = vals(0) + CDbl(prijData(i, colPrijKol))
            If IsNumeric(prijData(i, colPrijKol)) And IsNumeric(prijData(i, colPrijCena)) Then
                vals(1) = vals(1) + CDbl(prijData(i, colPrijKol)) * CDbl(prijData(i, colPrijCena))
            End If
            dict(vrsta) = vals
        Next i
    Else
        ' OM einzeln oder Zbirni (alle)
        Dim otkData As Variant
        If entitetID <> "" Then
            otkData = GetOtkupByStation(entitetID, datumOd, datumDo)
        Else
            otkData = GetTableData(TBL_OTKUP)
            If Not IsEmpty(otkData) Then
                Dim filters As New Collection
                Dim fp As clsFilterParam
                Set fp = New clsFilterParam
                fp.Init GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM), "BETWEEN", datumOd, datumDo
                filters.Add fp
                otkData = FilterArray(otkData, filters)
            End If
        End If
        If IsEmpty(otkData) Then
            ReportProsecnaCena = Empty
            Exit Function
        End If
        otkData = ExcludeStornirano(otkData, TBL_OTKUP)
        
        Dim colVrsta As Long, colKol As Long, colCena As Long
        colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
        colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
        colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
        
        For i = 1 To UBound(otkData, 1)
            Dim key As String
            key = CStr(otkData(i, colVrsta))
            If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
            vals = dict(key)
            If IsNumeric(otkData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkData(i, colKol))
            If IsNumeric(otkData(i, colKol)) And IsNumeric(otkData(i, colCena)) Then
                vals(1) = vals(1) + CDbl(otkData(i, colKol)) * CDbl(otkData(i, colCena))
            End If
            dict(key) = vals
        Next i
    End If
    
    If dict.count = 0 Then
        ReportProsecnaCena = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 4)
    
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = vals(0)
        result(i + 1, 3) = vals(1)
        If vals(0) > 0 Then
            result(i + 1, 4) = vals(1) / vals(0)
        Else
            result(i + 1, 4) = 0
        End If
    Next i
    
    ReportProsecnaCena = result
End Function

Public Function ReportManjak(ByVal entitetTip As String, _
                             ByVal entitetID As String, _
                             ByVal datumOd As Date, _
                             ByVal datumDo As Date) As Variant
    ' Returns: 2D Array (BrojZbirne, ZbirnaKg, PrijKg, ManjakKg, ManjakPct, ProsekGajbe)
    ' Letzte Zeile = UKUPNO
    
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then
        ReportManjak = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    If entitetTip = "Kupac" And entitetID <> "" Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC), "=", entitetID
        filters.Add fp
    ElseIf entitetTip = "Vozac" And entitetID <> "" Then
        Set fp = New clsFilterParam
        fp.Init GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC), "=", entitetID
        filters.Add fp
    End If
    
    Dim filtered As Variant
    filtered = FilterArray(zbrData, filters)
    If IsEmpty(filtered) Then
        ReportManjak = Empty
        Exit Function
    End If
    filtered = ExcludeStornirano(filtered, TBL_ZBIRNA)
    
    ' Prijemnica EINMAL laden für Performance
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    Dim colPBrZbr As Long, colPKol As Long
    If IsArray(prijData) Then
        colPBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        colPKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    End If
    
    ' Zbirna-Daten
    Dim colBroj As Long, colZbrKol As Long, colZbrAmb As Long
    colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    colZbrKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
    colZbrAmb = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB)
    
    ' Zbirne aggregieren (mehrere Zeilen pro BrojZbirne bei Klasa I + II)
    Dim zbrDict As Object
    Set zbrDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim brZbr As String
        brZbr = CStr(filtered(i, colBroj))
        If Not zbrDict.Exists(brZbr) Then zbrDict.Add brZbr, Array(0#, 0#)
        Dim zv As Variant
        zv = zbrDict(brZbr)
        If IsNumeric(filtered(i, colZbrKol)) Then zv(0) = zv(0) + CDbl(filtered(i, colZbrKol))
        If IsNumeric(filtered(i, colZbrAmb)) Then zv(1) = zv(1) + CLng(filtered(i, colZbrAmb))
        zbrDict(brZbr) = zv
    Next i
    
    ' Prijemnica pro BrojZbirne aggregieren
    Dim prijDict As Object
    Set prijDict = CreateObject("Scripting.Dictionary")
    
    If IsArray(prijData) Then
        For i = 1 To UBound(prijData, 1)
            Dim pBrZbr As String
            pBrZbr = CStr(prijData(i, colPBrZbr))
            If zbrDict.Exists(pBrZbr) Then
                If Not prijDict.Exists(pBrZbr) Then prijDict.Add pBrZbr, 0#
                If IsNumeric(prijData(i, colPKol)) Then
                    prijDict(pBrZbr) = prijDict(pBrZbr) + CDbl(prijData(i, colPKol))
                End If
            End If
        Next i
    End If
    
    ' Ergebnis
    Dim rowCount As Long
    rowCount = zbrDict.count
    If rowCount = 0 Then
        ReportManjak = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To rowCount + 1, 1 To 6)  ' +1 UKUPNO
    
    Dim keys As Variant
    keys = zbrDict.keys
    Dim totalZbrKg As Double, totalPrijKg As Double
    
    For i = 0 To zbrDict.count - 1
        zv = zbrDict(keys(i))
        Dim zbrKg As Double: zbrKg = zv(0)
        Dim zbrAmb As Long: zbrAmb = CLng(zv(1))
        
        Dim prijKg As Double: prijKg = 0
        If prijDict.Exists(keys(i)) Then prijKg = prijDict(keys(i))
        
        Dim manjakKg As Double: manjakKg = zbrKg - prijKg
        Dim manjakPct As Double
        If zbrKg > 0 Then manjakPct = manjakKg / zbrKg * 100 Else manjakPct = 0
        
        Dim prosek As Double: prosek = 0
        If zbrAmb > 0 Then prosek = zbrKg / zbrAmb
        
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = zbrKg
        result(i + 1, 3) = prijKg
        result(i + 1, 4) = manjakKg
        result(i + 1, 5) = manjakPct
        result(i + 1, 6) = prosek
        
        totalZbrKg = totalZbrKg + zbrKg
        totalPrijKg = totalPrijKg + prijKg
    Next i
    
    ' UKUPNO
    result(rowCount + 1, 1) = "UKUPNO"
    result(rowCount + 1, 2) = totalZbrKg
    result(rowCount + 1, 3) = totalPrijKg
    result(rowCount + 1, 4) = totalZbrKg - totalPrijKg
    If totalZbrKg > 0 Then
        result(rowCount + 1, 5) = (totalZbrKg - totalPrijKg) / totalZbrKg * 100
    Else
        result(rowCount + 1, 5) = 0
    End If
    result(rowCount + 1, 6) = ""
    
    ReportManjak = result
End Function

Public Function ReportZbirni(ByVal entitetTip As String, _
                             ByVal datumOd As Date, _
                             ByVal datumDo As Date) As Variant
    ' Returns: 2D Array (Entitet, Info, Col3, Col4, Col5)
    '   OM:    StanicaNaziv, Vrsta, Kolicina, Vrednost, ProsekCena
    '   Kupac: KupacNaziv, Vrsta, Kolicina, Vrednost, ProsekCena
    '   Vozac: VozacIme, AmbIzlaz, AmbVracena, ManjakKg, ManjakPct
    ' Letzte Zeile = UKUPNO
    
    Select Case entitetTip
        Case "OM":    ReportZbirni = ReportZbirniOM(datumOd, datumDo)
        Case "Kupac": ReportZbirni = ReportZbirniKupac(datumOd, datumDo)
        Case "Vozac": ReportZbirni = ReportZbirniVozac(datumOd, datumDo)
        Case Else:    ReportZbirni = Empty
    End Select
End Function

Private Function ReportZbirniOM(ByVal datumOd As Date, _
                                ByVal datumDo As Date) As Variant
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        ReportZbirniOM = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    Dim filtered As Variant
    filtered = FilterArray(data, filters)
    If IsEmpty(filtered) Then
        ReportZbirniOM = Empty
        Exit Function
    End If
    filtered = ExcludeStornirano(filtered, TBL_OTKUP)
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colStation As Long, colVrsta As Long, colKol As Long, colCena As Long
    colStation = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    colVrsta = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    colKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    colCena = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    
    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim key As String
        key = CStr(filtered(i, colStation)) & "|" & CStr(filtered(i, colVrsta))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(filtered(i, colKol)) Then vals(0) = vals(0) + CDbl(filtered(i, colKol))
        If IsNumeric(filtered(i, colKol)) And IsNumeric(filtered(i, colCena)) Then
            vals(1) = vals(1) + CDbl(filtered(i, colKol)) * CDbl(filtered(i, colCena))
        End If
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportZbirniOM = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 5)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totalKg As Double, totalRSD As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        Dim parts As Variant
        parts = Split(keys(i), "|")
        
        result(i + 1, 1) = CStr(LookupValue(TBL_STANICE, "StanicaID", parts(0), "Naziv"))
        result(i + 1, 2) = parts(1)
        result(i + 1, 3) = vals(0)
        result(i + 1, 4) = vals(1)
        If vals(0) > 0 Then result(i + 1, 5) = vals(1) / vals(0) Else result(i + 1, 5) = 0
        
        totalKg = totalKg + vals(0)
        totalRSD = totalRSD + vals(1)
    Next i
    
    result(dict.count + 1, 1) = ""
    result(dict.count + 1, 2) = "UKUPNO"
    result(dict.count + 1, 3) = totalKg
    result(dict.count + 1, 4) = totalRSD
    result(dict.count + 1, 5) = ""
    
    ReportZbirniOM = result
End Function

Private Function ReportZbirniKupac(ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then
        ReportZbirniKupac = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    Dim filtered As Variant
    filtered = FilterArray(data, filters)
    If IsEmpty(filtered) Then
        ReportZbirniKupac = Empty
        Exit Function
    End If
    filtered = ExcludeStornirano(filtered, TBL_PRIJEMNICA)
    
    ' Cache für Vrsta-Lookup
    Dim vrstaCache As Object
    Set vrstaCache = BuildZbirnaVrstaCache()
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colKupac As Long, colKol As Long, colCena As Long, colBrZbr As Long
    colKupac = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC)
    colKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    colCena = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    colBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    
    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim vrsta As String
        vrsta = GetVrstaFromCache(vrstaCache, CStr(filtered(i, colBrZbr)))
        If vrsta = "" Then vrsta = "(Nepoznato)"
        
        Dim key As String
        key = CStr(filtered(i, colKupac)) & "|" & vrsta
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(filtered(i, colKol)) Then vals(0) = vals(0) + CDbl(filtered(i, colKol))
        If IsNumeric(filtered(i, colKol)) And IsNumeric(filtered(i, colCena)) Then
            vals(1) = vals(1) + CDbl(filtered(i, colKol)) * CDbl(filtered(i, colCena))
        End If
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportZbirniKupac = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 5)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totalKg As Double, totalRSD As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        Dim parts As Variant
        parts = Split(keys(i), "|")
        
        result(i + 1, 1) = CStr(LookupValue(TBL_KUPCI, "KupacID", parts(0), "Naziv"))
        result(i + 1, 2) = parts(1)
        result(i + 1, 3) = vals(0)
        result(i + 1, 4) = vals(1)
        If vals(0) > 0 Then result(i + 1, 5) = vals(1) / vals(0) Else result(i + 1, 5) = 0
        
        totalKg = totalKg + vals(0)
        totalRSD = totalRSD + vals(1)
    Next i
    
    result(dict.count + 1, 1) = ""
    result(dict.count + 1, 2) = "UKUPNO"
    result(dict.count + 1, 3) = totalKg
    result(dict.count + 1, 4) = totalRSD
    result(dict.count + 1, 5) = ""
    
    ReportZbirniKupac = result
End Function

Private Function ReportZbirniVozac(ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then
        ReportZbirniVozac = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init GetColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    Dim zbrFiltered As Variant
    zbrFiltered = FilterArray(zbrData, filters)
    If IsEmpty(zbrFiltered) Then
        ReportZbirniVozac = Empty
        Exit Function
    End If
    zbrFiltered = ExcludeStornirano(zbrFiltered, TBL_ZBIRNA)
    
    ' Prijemnica-Daten EINMAL laden (Performance-Fix)
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    
    Dim colPBrZbr As Long, colPAmbVr As Long, colPKol As Long
    If IsArray(prijData) Then
        colPBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
        colPAmbVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA)
        colPKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colVozac As Long, colBroj As Long, colKol As Long, colAmb As Long
    colVozac = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC)
    colBroj = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ)
    colKol = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA)
    colAmb = GetColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB)
    
    Dim i As Long, j As Long
    For i = 1 To UBound(zbrFiltered, 1)
        Dim vozacID As String
        vozacID = CStr(zbrFiltered(i, colVozac))
        If Not dict.Exists(vozacID) Then dict.Add vozacID, Array(0#, 0#, 0#, 0#)
        ' (0)=AmbIzlaz, (1)=AmbVracena, (2)=ZbirnaKg, (3)=PrijKg
        
        Dim vals As Variant
        vals = dict(vozacID)
        
        If IsNumeric(zbrFiltered(i, colAmb)) Then vals(0) = vals(0) + CLng(zbrFiltered(i, colAmb))
        If IsNumeric(zbrFiltered(i, colKol)) Then vals(2) = vals(2) + CDbl(zbrFiltered(i, colKol))
        
        ' Prijemnica-Daten für diese Zbirna aus vorgeladenem Array
        Dim brZbr As String
        brZbr = CStr(zbrFiltered(i, colBroj))
        
        If IsArray(prijData) Then
            For j = 1 To UBound(prijData, 1)
                If CStr(prijData(j, colPBrZbr)) = brZbr Then
                    If IsNumeric(prijData(j, colPKol)) Then vals(3) = vals(3) + CDbl(prijData(j, colPKol))
                    If IsNumeric(prijData(j, colPAmbVr)) Then vals(1) = vals(1) + CLng(prijData(j, colPAmbVr))
                End If
            Next j
        End If
        
        dict(vozacID) = vals
    Next i
    
    If dict.count = 0 Then
        ReportZbirniVozac = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 5)
    
    Dim keys As Variant
    keys = dict.keys
    Dim tAmbIzl As Long, tAmbVr As Long, tZbrKg As Double, tPrijKg As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        
        Dim manjakKg As Double
        manjakKg = vals(2) - vals(3)
        Dim manjakPct As Double
        If vals(2) > 0 Then manjakPct = manjakKg / vals(2) * 100 Else manjakPct = 0
        
        result(i + 1, 1) = CStr(LookupValue(TBL_VOZACI, "VozacID", keys(i), "Ime")) & " " & _
                           CStr(LookupValue(TBL_VOZACI, "VozacID", keys(i), "Prezime"))
        result(i + 1, 2) = vals(0)       ' AmbIzlaz
        result(i + 1, 3) = vals(1)       ' AmbVracena
        result(i + 1, 4) = manjakKg      ' ManjakKg
        result(i + 1, 5) = manjakPct     ' ManjakPct
        
        tAmbIzl = tAmbIzl + vals(0)
        tAmbVr = tAmbVr + vals(1)
        tZbrKg = tZbrKg + vals(2)
        tPrijKg = tPrijKg + vals(3)
    Next i
    
    ' UKUPNO
    result(dict.count + 1, 1) = "UKUPNO"
    result(dict.count + 1, 2) = tAmbIzl
    result(dict.count + 1, 3) = tAmbVr
    result(dict.count + 1, 4) = tZbrKg - tPrijKg
    If tZbrKg > 0 Then
        result(dict.count + 1, 5) = (tZbrKg - tPrijKg) / tZbrKg * 100
    Else
        result(dict.count + 1, 5) = 0
    End If
    
    ReportZbirniVozac = result
End Function

' ============================================================
' AUSGABE (unverändert)
' ============================================================

Public Sub OutputToSheet(ByVal data As Variant, ByVal targetRange As Range, _
                         Optional ByVal headers As Variant)
    If IsEmpty(data) Then
        targetRange.Value = "Nema podataka"
        Exit Sub
    End If
    
    Dim startRow As Long
    startRow = 0
    
    If Not IsMissing(headers) Then
        Dim h As Long
        For h = LBound(headers) To UBound(headers)
            targetRange.Offset(0, h - LBound(headers)).Value = headers(h)
            targetRange.Offset(0, h - LBound(headers)).Font.Bold = True
        Next h
        startRow = 1
    End If
    
    Dim r As Long, c As Long
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            targetRange.Offset(startRow + r - 1, c - 1).Value = data(r, c)
        Next c
    Next r
End Sub


' ============================================================
' SHARED HELPER – Dict(Key zu Array(Kg, RSD)) zu 2D Result
' ============================================================

Private Function DictToResultArray(ByVal dict As Object) As Variant
    ' Konvertiert Dictionary(String ? Array(Double, Double))
    ' zu 2D Array (Nr, Key, Kg, RSD) + UKUPNO
    
    If dict.count = 0 Then
        DictToResultArray = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 4)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totalKg As Double, totalRSD As Double
    
    Dim i As Long
    For i = 0 To dict.count - 1
        Dim vals As Variant
        vals = dict(keys(i))
        result(i + 1, 1) = CStr(i + 1)
        result(i + 1, 2) = keys(i)
        result(i + 1, 3) = vals(0)
        result(i + 1, 4) = vals(1)
        totalKg = totalKg + vals(0)
        totalRSD = totalRSD + vals(1)
    Next i
    
    result(dict.count + 1, 1) = ""
    result(dict.count + 1, 2) = "UKUPNO"
    result(dict.count + 1, 3) = totalKg
    result(dict.count + 1, 4) = totalRSD
    
    DictToResultArray = result
End Function

' ============================================================
' SHARED HELPER – Entitet-Name auflösen
' ============================================================

Private Function ResolveEntitetName(ByVal entitetID As String, _
                                    ByVal entitetTip As String) As String
    Select Case entitetTip
        Case "Stanica"
            ResolveEntitetName = CStr(LookupValue(TBL_STANICE, "StanicaID", entitetID, "Naziv"))
        Case "Kupac"
            ResolveEntitetName = CStr(LookupValue(TBL_KUPCI, "KupacID", entitetID, "Naziv"))
        Case "Kooperant"
            ResolveEntitetName = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Ime")) & " " & _
                                 CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Prezime"))
        Case Else
            ResolveEntitetName = entitetID
    End Select
End Function
