Attribute VB_Name = "modMigracija"
'Attribute VB_Name = "modMigracija"
' ============================================================
' modMigracija - jednokratna migracija CISTIH PODATAKA iz starog
' OtkupApp fajla u novi (prazan). Mapiranje PO IMENU kolone,
' samo vrednosti (bez formula i koda). Ne treba "Trust access to
' VBA project object model" (koristi obican Excel objektni model).
'
' Upotreba: u NOVOM fajlu  ->  Alt+F8  ->  MigrirajPodatkeIzStarog
'
' Provere integriteta (da "success" ne sakrije izgubljene podatke):
'   - tabele koje su SAMO u starom (nema ih u novom)      -> glasno prijavljeno
'   - stare kolone sa podatkom bez cilja u novom (rename)  -> prijavljeno
'   - zbir CISTO numerickih kolona (kolicine/vrednosti): staro vs novo
'   Bilo koji problem: naslov "PROBLEMI: N" + upozoravajuca ikonica.
' ============================================================
Option Explicit

Public Sub MigrirajPodatkeIzStarog()
    Dim putanja As Variant
    putanja = Application.GetOpenFilename( _
        "OtkupApp fajlovi (*.xlsm;*.xlsb),*.xlsm;*.xlsb", , _
        "Izaberi STARI OtkupApp fajl (sa podacima)")
    If VarType(putanja) = vbBoolean Then Exit Sub      ' Cancel

    ' --- Zastita od slucajnog destruktivnog rerun-a ---
    Dim chk As ListObject
    Set chk = NadjiListObject(ThisWorkbook, "tblOtkup")
    If Not chk Is Nothing Then
        If chk.ListRows.count > 0 Then
            If MsgBox("Ovaj fajl VE" & ChrW(262) & " ima podatke (tblOtkup: " & chk.ListRows.count & _
                      " redova). Migracija ce ih PREPISATI." & vbCrLf & vbCrLf & _
                      "Nastaviti?", vbExclamation + vbYesNo + vbDefaultButton2, _
                      "Migracija podataka") = vbNo Then Exit Sub
        End If
    End If

    Dim novi As Workbook: Set novi = ThisWorkbook
    Dim stari As Workbook
    Dim prevEvents As Boolean, prevCalc As XlCalculation, prevSec As Long, prevSU As Boolean
    Dim summary As String, total As Long, tbls As Long, problems As Long

    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevSec = Application.AutomationSecurity
    prevSU = Application.ScreenUpdating

    ' Foolproof: novi fajl MORA imati tblKorisnici (+ audit kolone) PRE kopiranja,
    ' jer migracija prolazi kroz tabele NOVOG fajla pa povlaci istoimene iz starog.
    ' Bez ovoga, ako Ensure nije rucno pokrenut, korisnici se ne bi preneli.
    ' Idempotentno; best-effort (greska u semi ne sme da obori migraciju).
    On Error Resume Next
    EnsureKorisniciSchema
    EnsureAuditColumnsCore
    On Error GoTo 0

    On Error GoTo CLEAN
    Application.EnableEvents = False                    ' ne pokreci Workbook_Open starog
    Application.AutomationSecurity = 3                  ' msoAutomationSecurityForceDisable (bez makroa)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set stari = Workbooks.Open(fileName:=CStr(putanja), ReadOnly:=True, UpdateLinks:=0)

    Dim ws As Worksheet, loNovi As ListObject, n As Long
    Dim ckey As String, cval As String, isCfg As Boolean, eDesc As String, errNum As Long
    Dim warn As String
    For Each ws In novi.Worksheets
        For Each loNovi In ws.ListObjects
            If Not SkipTabela(loNovi.name) Then
                isCfg = ConfigKolone(loNovi.name, ckey, cval)
                warn = ""

                On Error Resume Next                       ' jedna losa tabela ne prekida ceo prolaz
                Err.Clear
                If isCfg Then
                    n = MergeConfigTabelu(stari, loNovi, ckey, cval)
                Else
                    n = KopirajTabelu(stari, loNovi, warn, problems)
                End If
                errNum = Err.Number: eDesc = Err.description
                On Error GoTo CLEAN

                If errNum <> 0 Then
                    problems = problems + 1
                    summary = summary & "  " & loNovi.name & " - GRE" & ChrW(352) & "KA: " & eDesc & vbCrLf
                ElseIf n = -1 Then
                    summary = summary & "  " & loNovi.name & " - (nema u starom)" & vbCrLf
                ElseIf n = -2 Then
                    problems = problems + 1
                    summary = summary & "  " & loNovi.name & " - (config kolone nenadjene!)" & vbCrLf
                ElseIf isCfg Then
                    tbls = tbls + 1
                    summary = summary & "  " & loNovi.name & " - merge, " & n & " kljuceva" & vbCrLf
                Else
                    total = total + n: tbls = tbls + 1
                    summary = summary & "  " & loNovi.name & " - " & n & " red." & vbCrLf
                    If Len(warn) > 0 Then summary = summary & warn
                End If
            End If
        Next loNovi
    Next ws

    ' --- A) Obrnuti prolaz kroz STARI: tabele koje NISU u novom (redovi bi tiho ostali) ---
    Dim wsS As Worksheet, loS As ListObject, cntOld As Long
    For Each wsS In stari.Worksheets
        For Each loS In wsS.ListObjects
            If Not SkipTabela(loS.name) Then
                If NadjiListObject(novi, loS.name) Is Nothing Then
                    cntOld = 0
                    On Error Resume Next
                    If Not loS.DataBodyRange Is Nothing Then cntOld = loS.ListRows.count
                    On Error GoTo CLEAN
                    If cntOld > 0 Then
                        problems = problems + 1
                        summary = summary & "  !! " & loS.name & " - U STAROM " & cntOld & _
                                  " red., a tabele NEMA u novom -> NIJE preneto!" & vbCrLf
                    End If
                End If
            End If
        Next loS
    Next wsS

    stari.Close SaveChanges:=False: Set stari = Nothing

CLEAN:
    Dim em As String
    If Err.Number <> 0 Then em = "GRE" & ChrW(352) & "KA: " & Err.description & vbCrLf & vbCrLf
    On Error Resume Next
    If Not stari Is Nothing Then stari.Close SaveChanges:=False
    Application.Calculation = prevCalc
    Application.AutomationSecurity = prevSec
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
    On Error GoTo 0

    Dim hdr As String
    If problems > 0 Then hdr = "!! PROBLEMI: " & problems & " -> vidi '!!' i '~' redove dole !!" & vbCrLf & vbCrLf
    MsgBox em & hdr & "Tabela preneto: " & tbls & "   |   redova ukupno: " & total & _
           vbCrLf & vbCrLf & summary & vbCrLf & _
           "Sada SNIMI novi fajl (Ctrl+S) i proveri par tabela.", _
           IIf(Len(em) > 0 Or problems > 0, vbExclamation, vbInformation), "Migracija podataka"
End Sub

' Vrati broj prenetih redova; -1 ako tabele nema u starom fajlu.
Private Function KopirajTabelu(ByVal stari As Workbook, ByVal loNovi As ListObject, _
                               ByRef warn As String, ByRef prob As Long) As Long
    Dim loStari As ListObject
    Set loStari = NadjiListObject(stari, loNovi.name)
    If loStari Is Nothing Then KopirajTabelu = -1: Exit Function

    ' stari prazan (nema body) -> nista za prenos
    If loStari.DataBodyRange Is Nothing Then KopirajTabelu = 0: Exit Function
    If loStari.ListRows.count = 0 Then KopirajTabelu = 0: Exit Function
    ' ocisti eventualne postojece redove u novom (idempotentno)
    If Not loNovi.DataBodyRange Is Nothing Then loNovi.DataBodyRange.ClearContents

    ' mapiranje: novi col -> stari col, PO IMENU
    Dim nNew As Long: nNew = loNovi.ListColumns.count
    Dim mapCol() As Long: ReDim mapCol(1 To nNew)
    Dim j As Long, k As Long, staroIme As String
    For j = 1 To nNew
        staroIme = StaroImeKolone(loNovi.name, loNovi.ListColumns(j).name)
        For k = 1 To loStari.ListColumns.count
            If StrComp(Trim$(loStari.ListColumns(k).name), Trim$(staroIme), vbTextCompare) = 0 Then
                mapCol(j) = k: Exit For
            End If
        Next k
    Next j

    ' telo starog kao 2D niz (vrednosti, ne formule)
    Dim SRC As Variant: SRC = loStari.DataBodyRange.value
    If Not IsArray(SRC) Then
        Dim one(1 To 1, 1 To 1) As Variant: one(1, 1) = SRC: SRC = one
    End If
    Dim nRows As Long: nRows = UBound(SRC, 1)

    ' osiguraj TACNO nRows redova u novoj tabeli - robustno preko ListRows.Add/Delete
    ' (radi i kad tabela nema prikazan header ili deli sheet sa drugom tabelom;
    '  bez Resize-a koji ume da padne na Error 91 / koliziju)
    Do While loNovi.ListRows.count > nRows
        loNovi.ListRows(loNovi.ListRows.count).Delete
    Loop
    Do While loNovi.ListRows.count < nRows
        loNovi.ListRows.Add
    Loop

    ' upisi SAMO mapirane kolone (nove/kalkulisane kolone ostaju netaknute)
    Dim colArr() As Variant, r As Long
    For j = 1 To nNew
        If mapCol(j) > 0 Then
            ReDim colArr(1 To nRows, 1 To 1)
            For r = 1 To nRows
                colArr(r, 1) = SRC(r, mapCol(j))
            Next r
            loNovi.ListColumns(j).DataBodyRange.value = colArr
        End If
    Next j
    KopirajTabelu = nRows

    ' --- Provere integriteta posle kopiranja (best-effort; ne obaraju uspeh) ---
    On Error Resume Next
    ProveriKoloneIZbir loStari, loNovi, SRC, mapCol, nNew, warn, prob
    Err.Clear
    On Error GoTo 0
End Function

' True ako je tabela key/value config; vrati imena key/value kolona.
Private Function ConfigKolone(ByVal naziv As String, ByRef keyCol As String, _
                              ByRef valCol As String) As Boolean
    Select Case LCase$(naziv)
        Case "tblconfig", "tbllocalconfig"
            keyCol = "Kljuc": valCol = "Vrednost": ConfigKolone = True
        Case "tblsefconfig"
            keyCol = "ConfigKey": valCol = "ConfigValue": ConfigKolone = True
    End Select
End Function

' Config merge (key/value): ako novi red VEC ima vrednost -> ostaje (ne prepisuje
' se iz starog); ako je prazno -> uzima se iz starog. Kljucevi kojih nema u novom
' se dodaju iz starog (sve kolone po imenu). Vrati broj kljuceva u novom.
'   -1 = tabele nema u starom;  -2 = key/value kolone nisu nadjene.
Private Function MergeConfigTabelu(ByVal stari As Workbook, ByVal loNovi As ListObject, _
                                   ByVal keyCol As String, ByVal valCol As String) As Long
    Dim loStari As ListObject
    Set loStari = NadjiListObject(stari, loNovi.name)
    If loStari Is Nothing Then MergeConfigTabelu = -1: Exit Function

    Dim sKey As Long, sVal As Long, nKey As Long, nVal As Long
    sKey = ColIndexByName(loStari, keyCol): sVal = ColIndexByName(loStari, valCol)
    nKey = ColIndexByName(loNovi, keyCol): nVal = ColIndexByName(loNovi, valCol)
    If sKey = 0 Or sVal = 0 Or nKey = 0 Or nVal = 0 Then MergeConfigTabelu = -2: Exit Function

    ' stari -> dict: kljuc -> vrednost (+ pamti red radi dodavanja ostalih kolona)
    Dim dVal As Object: Set dVal = CreateObject("Scripting.Dictionary"): dVal.CompareMode = vbTextCompare
    Dim dRow As Object: Set dRow = CreateObject("Scripting.Dictionary"): dRow.CompareMode = vbTextCompare
    Dim i As Long, kkey As String
    For i = 1 To loStari.ListRows.count
        kkey = Trim$(CStr(loStari.DataBodyRange.cells(i, sKey).value))
        If Len(kkey) > 0 And Not dVal.Exists(kkey) Then
            dVal(kkey) = loStari.DataBodyRange.cells(i, sVal).value
            dRow(kkey) = i
        End If
    Next i

    ' novi: prazne vrednosti popuni iz starog; oznaci postojece kljuceve
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For i = 1 To loNovi.ListRows.count
        kkey = Trim$(CStr(loNovi.DataBodyRange.cells(i, nKey).value))
        If Len(kkey) > 0 Then
            seen(kkey) = True
            If Len(Trim$(CStr(loNovi.DataBodyRange.cells(i, nVal).value))) = 0 Then
                If dVal.Exists(kkey) Then loNovi.DataBodyRange.cells(i, nVal).value = dVal(kkey)
            End If
        End If
    Next i

    ' kljucevi iz starog kojih nema u novom -> dodaj red
    Dim k As Variant, lr As ListRow
    For Each k In dVal.keys
        If Not seen.Exists(CStr(k)) Then
            Set lr = loNovi.ListRows.Add
            KopirajRedPoImenu loStari, CLng(dRow(CStr(k))), loNovi, lr.index
        End If
    Next k

    MergeConfigTabelu = loNovi.ListRows.count
End Function

Private Function ColIndexByName(ByVal lo As ListObject, ByVal naziv As String) As Long
    Dim c As Long
    For c = 1 To lo.ListColumns.count
        If StrComp(Trim$(lo.ListColumns(c).name), Trim$(naziv), vbTextCompare) = 0 Then
            ColIndexByName = c: Exit Function
        End If
    Next c
End Function

' Kopira jedan red iz starog u red novog (vrednosti), mapiranje po imenu kolone.
Private Sub KopirajRedPoImenu(ByVal loStari As ListObject, ByVal staroRed As Long, _
                              ByVal loNovi As ListObject, ByVal noviRed As Long)
    Dim j As Long, sIdx As Long
    For j = 1 To loNovi.ListColumns.count
        sIdx = ColIndexByName(loStari, loNovi.ListColumns(j).name)
        If sIdx > 0 Then
            loNovi.DataBodyRange.cells(noviRed, j).value = loStari.DataBodyRange.cells(staroRed, sIdx).value
        End If
    Next j
End Sub

Private Function NadjiListObject(ByVal wb As Workbook, ByVal naziv As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, naziv, vbTextCompare) = 0 Then
                Set NadjiListObject = lo: Exit Function
            End If
        Next lo
    Next ws
End Function

' Preskace se SAMO tblRpt* (izvedeni izvestaji - regenerisu se iz podataka).
' SVE ostalo se prenosi, ukljucujuci config tabele (tblConfig, tblSEFConfig,
' tblLocalConfig) jer cuvaju aktuelna podesavanja (OAuth / SEF / bank putanje /
' licenca). Config tabele su key/value; novi prazan fajl ih ima prazne do
' SetupNewPC, pa pun copy starih redova donosi sva podesavanja bez gubitka.
Private Function SkipTabela(ByVal naziv As String) As Boolean
    SkipTabela = (LCase$(Left$(naziv, 6)) = "tblrpt")
End Function

' Override za PREIMENOVANE kolone izmedju verzija. Default = isto ime.
' Primer (ako je novo "Kontakt" u staroj verziji bilo "Telefon"):
'   If tabela = "tblStanice" And novoIme = "Kontakt" Then StaroImeKolone = "Telefon": Exit Function
Private Function StaroImeKolone(ByVal tabela As String, ByVal novoIme As String) As String
    StaroImeKolone = novoIme
End Function

' Provere posle kopiranja jedne tabele (best-effort; NIKAD ne baca gresku dalje):
'   B) stare kolone SA PODATKOM koje nemaju cilj u novom (preimenovane/izbacene)
'   D) zbir CISTO numerickih kolona: staro (izvor SRC) vs novo (citano nazad DST)
' Nadje li problem: doda '!!'/'~' red u warn i uveca prob.
Private Sub ProveriKoloneIZbir(ByVal loStari As ListObject, ByVal loNovi As ListObject, _
                               ByRef SRC As Variant, ByRef mapCol() As Long, ByVal nNew As Long, _
                               ByRef warn As String, ByRef prob As Long)
    On Error Resume Next
    Dim nRows As Long: nRows = UBound(SRC, 1)
    Dim j As Long, k As Long, rr As Long, hasData As Boolean
    Dim oc As Long, ob As Long, nc As Long, nb As Long
    Dim oldSum As Double, newSum As Double

    ' B) koje su stare kolone iskoriscene (imaju cilj u novom)?
    Dim usedOld() As Boolean
    ReDim usedOld(1 To loStari.ListColumns.count)
    For j = 1 To nNew
        If mapCol(j) > 0 Then usedOld(mapCol(j)) = True
    Next j
    For k = 1 To loStari.ListColumns.count
        If Not usedOld(k) Then
            hasData = False
            For rr = 1 To nRows
                Select Case VarType(SRC(rr, k))
                    Case vbEmpty, vbNull, vbError
                        ' nema podatka
                    Case vbString
                        If Len(Trim$(CStr(SRC(rr, k)))) > 0 Then hasData = True: Exit For
                    Case Else
                        hasData = True: Exit For
                End Select
            Next rr
            If hasData Then
                prob = prob + 1
                warn = warn & "     !! kolona '" & loStari.ListColumns(k).name & _
                       "' ima podatke u starom a NEMA cilj u novom -> ta kolona NIJE preneta" & vbCrLf
            End If
        End If
    Next k

    ' D) zbir CISTO numerickih kolona: staro (SRC) vs novo (citano nazad)
    Dim DST As Variant
    DST = loNovi.DataBodyRange.value
    If Not IsArray(DST) Then
        Dim one(1 To 1, 1 To 1) As Variant: one(1, 1) = DST: DST = one
    End If
    For j = 1 To nNew
        If mapCol(j) > 0 Then
            oc = 0: ob = 0
            oldSum = SumNumeric(SRC, mapCol(j), oc, ob)
            ' cisto numericka kolona = bar 1 broj i 0 ne-praznih ne-brojeva na strani starog
            If oc > 0 And ob = 0 Then
                nc = 0: nb = 0
                newSum = SumNumeric(DST, j, nc, nb)
                If Round(oldSum, 2) <> Round(newSum, 2) Then
                    prob = prob + 1
                    warn = warn & "     ~ kolona '" & loNovi.ListColumns(j).name & _
                           "': zbir staro=" & Format$(oldSum, "0.00") & _
                           " novo=" & Format$(newSum, "0.00") & " -> RAZLIKA, proveri!" & vbCrLf
                End If
            End If
        End If
    Next j
End Sub

' Zbir SAMO stvarno numerickih celija (broj/datum); text-brojevi se NE sabiraju.
'   cnt = koliko numerickih;  bad = koliko NE-praznih NE-numerickih (tekst/greska/bool).
Private Function SumNumeric(ByRef arr As Variant, ByVal col As Long, _
                            ByRef cnt As Long, ByRef bad As Long) As Double
    Dim r As Long, s As Double
    cnt = 0: bad = 0: s = 0#
    For r = 1 To UBound(arr, 1)
        Select Case VarType(arr(r, col))
            Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDate, vbDecimal
                s = s + CDbl(arr(r, col)): cnt = cnt + 1
            Case vbEmpty, vbNull
                ' prazno - ignorisi
            Case vbString
                If Len(Trim$(CStr(arr(r, col)))) > 0 Then bad = bad + 1
            Case Else
                bad = bad + 1
        End Select
    Next r
    SumNumeric = s
End Function

