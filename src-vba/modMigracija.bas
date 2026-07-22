Attribute VB_Name = "modMigracija"
'Attribute VB_Name = "modMigracija"
' ============================================================
' modMigracija - jednokratna migracija CISTIH PODATAKA iz starog
' OtkupApp fajla u novi (prazan). Mapiranje PO IMENU kolone,
' vrednosti + format kolone (bez formula i koda). Ne treba "Trust access to
' VBA project object model" (koristi obican Excel objektni model).
'
' Upotreba: u NOVOM fajlu  ->  Alt+F8  ->  MigrirajPodatkeIzStarog
'
' Provere integriteta (da "success" ne sakrije izgubljene podatke):
'   - tabele koje su SAMO u starom (nema ih u novom)      -> glasno prijavljeno
'   - stare kolone sa podatkom bez cilja u novom (rename)  -> prijavljeno
'   - NOVE kolone bez izvora u starom: vezne (*ID / Broj*) -> PROBLEM (red stize
'     razvezan); audit se pune unapred pa se preskacu; kalkulisane (formula) tako-
'     dje; ostale (kozmetika) -> samo info
'   - zbir CISTO numerickih kolona: staro vs novo (citano nazad) = da li je UPIS
'     legao (kalkulisana kolona/validacija/koercija); NE hvata pogresno MAPIRANJE
'   - fail-closed: provera koja NIJE izvedena se prijavi (nije isto sto i "prosla")
'   Bilo koji problem: naslov "PROBLEMI: N" + upozoravajuca ikonica.
'
' Format: ako je kolona u NOVOM "General", preuzme se NumberFormat iz starog
' (datumi/iznosi inace posle array-upisa ostaju goli brojevi). Namerni format
' novog sablona (ne-General) se NE dira.
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

    ' Pre-migracija backup (rollback ako nesto podje po zlu; best-effort, kao pre-update).
    Dim bkPath As String
    If Len(ThisWorkbook.path) = 0 Then
        summary = summary & "  (backup preskocen: fajl nije snimljen na disk)" & vbCrLf
    Else
        bkPath = BackupPreMigracije()
        If Len(bkPath) > 0 Then
            summary = summary & "  (backup: " & bkPath & ")" & vbCrLf
        Else
            problems = problems + 1
            summary = summary & "  !! backup NIJE uspeo (disk/zakljucan?) - rollback nece biti trivijalan" & vbCrLf
        End If
    End If

    ' Foolproof: novi fajl MORA imati tblKorisnici (+ audit kolone) PRE kopiranja,
    ' jer migracija prolazi kroz tabele NOVOG fajla pa povlaci istoimene iz starog.
    ' Bez ovoga, ako Ensure nije rucno pokrenut, korisnici se ne bi preneli.
    ' Best-effort ALI fail-closed: greska u semi ne obara migraciju, ali se PRIJAVI.
    On Error Resume Next
    Err.Clear: EnsureKorisniciSchema
    If Err.Number <> 0 Then problems = problems + 1: summary = summary & "  !! EnsureKorisniciSchema NIJE izvedena (" & Err.description & ") -> korisnici/kolone mozda nepotpuni" & vbCrLf
    Err.Clear: EnsureAuditColumnsCore
    If Err.Number <> 0 Then problems = problems + 1: summary = summary & "  !! EnsureAuditColumnsCore NIJE izvedena (" & Err.description & ") -> audit kolone mozda nedostaju" & vbCrLf
    Err.Clear
    On Error GoTo 0

    On Error GoTo CLEAN
    Application.EnableEvents = False                    ' ne pokreci Workbook_Open starog
    Application.AutomationSecurity = 3                  ' msoAutomationSecurityForceDisable (bez makroa)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set stari = Workbooks.Open(fileName:=CStr(putanja), ReadOnly:=True, UpdateLinks:=0)

    ' Sanity: da li je izabran BAS OtkupApp fajl? Inace svaka tabela vrati -1,
    ' problems ostane 0, i "0 redova" izgleda kao uspeh. tblOtkup = potpis baze.
    If NadjiListObject(stari, "tblOtkup") Is Nothing Then
        problems = problems + 1
        summary = summary & "  !! izabrani fajl NEMA tblOtkup -> nije OtkupApp baza? (nista nije preneto)" & vbCrLf
        GoTo CLEAN
    End If

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

    Err.Clear                                   ' ocisti zaostali per-tabela Err (inace se u CLEAN prijavi 2x)
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

    ' Ne dozvoli da tih AutoSave / refleksni Ctrl+S upise POLU-migriran ili proble-
    ' matican fajl. Na problem/gresku oznaci "snimljeno" -> operater mora SVESNO da
    ' snimi posle citanja upozorenja (ista logika kao self-update, ne apel u MsgBox).
    If problems > 0 Or Len(em) > 0 Then ThisWorkbook.Saved = True

    Dim hdr As String
    If problems > 0 Then hdr = "!! PROBLEMI: " & problems & " -> vidi '!!' i '~' redove dole !!" & vbCrLf & vbCrLf
    ' Footer objasnjava sta Saved=True znaci (inace operater zatvori bez prompta i
    ' izgubi ceo posao - podatak nije unisten, ali sat vremena rada jeste).
    Dim foot As String
    If problems > 0 Or Len(em) > 0 Then
        foot = "PAZI: fajl NIJE snimljen i Excel NECE pitati pri zatvaranju." & vbCrLf & _
               "Ako nastavljas, snimi SVESNO (Ctrl+S). Stari fajl je netaknut."
    Else
        foot = "Sada SNIMI novi fajl (Ctrl+S) i proveri par tabela."
    End If
    MsgBox em & hdr & "Tabela preneto: " & tbls & "   |   redova ukupno: " & total & _
           vbCrLf & vbCrLf & summary & vbCrLf & foot, _
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

    ' osiguraj TACNO nRows redova u novoj tabeli.
    ' Smanjivanje (redak rerun nad manjim) -> robustno Delete red po red.
    Do While loNovi.ListRows.count > nRows
        loNovi.ListRows(loNovi.ListRows.count).Delete
    Loop
    ' Povecavanje -> prvo BRZI bulk-resize (jedan poziv). Add-po-red je bio usko
    ' grlo na velikim bazama (tblOtkup, desetine hiljada = po jedan COM poziv/red).
    ' Resize ume da padne na Error 91 / koliziju (deljen sheet) -> tad fallback na
    ' robustnu Add petlju (isto ponasanje kao ranije).
    If loNovi.ListRows.count < nRows Then
        Dim okBrzo As Boolean: okBrzo = False
        On Error Resume Next
        Dim hRng As Range: Set hRng = loNovi.HeaderRowRange
        ' Resize BEZBEDNO samo ako je ciljna zona ispod header-a PRAZNA. Inace bi
        ' Resize upio zaostale celije (napomena/rucni unos/ostatak testa) kao redove
        ' tabele; nemapirane kolone bi zadrzale tudji sadrzaj, a provera C bi ga
        ' smatrala "ocekivano praznim" -> strani podatak, a izvestaj kaze OK.
        ' Ako nije prazno -> fallback na (sporu ali bezbednu) Add petlju.
        If Not hRng Is Nothing Then
            If Application.CountA(hRng.Offset(1).Resize(nRows, hRng.columns.count)) = 0 Then
                loNovi.Resize hRng.Resize(nRows + 1, hRng.columns.count)
                okBrzo = (Err.Number = 0 And loNovi.ListRows.count = nRows)
            End If
        End If
        Err.Clear
        On Error GoTo 0
        If Not okBrzo Then
            Do While loNovi.ListRows.count < nRows
                loNovi.ListRows.Add
            Loop
        End If
    End If

    ' upisi SAMO mapirane kolone (nove/kalkulisane kolone ostaju netaknute)
    Dim colArr() As Variant, r As Long
    For j = 1 To nNew
        If mapCol(j) > 0 Then
            ReDim colArr(1 To nRows, 1 To 1)
            For r = 1 To nRows
                colArr(r, 1) = SRC(r, mapCol(j))
            Next r
            loNovi.ListColumns(j).DataBodyRange.value = colArr
            PreuzmiFormatKolone loStari, mapCol(j), loNovi, j   ' datumi/iznosi: format iz starog ako je novi General
        End If
    Next j
    KopirajTabelu = nRows

    ' --- Provere integriteta posle kopiranja (best-effort; ne obaraju uspeh) ---
    On Error Resume Next
    Err.Clear
    ProveriKoloneIZbir loStari, loNovi, SRC, mapCol, nNew, warn, prob
    If Err.Number <> 0 Then                       ' fail-closed: provera koja nije dovrsena != prosla
        prob = prob + 1
        warn = warn & "     !! provere integriteta za ovu tabelu NISU dovrsene (" & Err.description & ")" & vbCrLf
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Preuzmi NumberFormat iz starog u novi za jednu kolonu, ali SAMO ako je novi
' "General" (da se ne pregazi namerni format novog sablona). Datumi/iznosi bi
' inace posle array-upisa (.Value = niz) ostali goli brojevi. Best-effort.
Private Sub PreuzmiFormatKolone(ByVal loStari As ListObject, ByVal sCol As Long, _
                                ByVal loNovi As ListObject, ByVal nCol As Long)
    On Error Resume Next
    Dim novaBody As Range, staraBody As Range
    Set novaBody = loNovi.ListColumns(nCol).DataBodyRange
    Set staraBody = loStari.ListColumns(sCol).DataBodyRange
    If Not (novaBody Is Nothing Or staraBody Is Nothing) Then
        ' format iz PRVE NEPRAZNE celije starog (prazna/General prva celija ne sme
        ' da sakrije datumsku/iznosnu kolonu sa vodecim praznim redovima)
        Dim srcCell As Range, dstFmt As String, srcFmt As String
        If staraBody.cells.count = 1 Then
            ' JEDNA celija -> SpecialCells bi radio nad CELIM sheet-om (Excel zamka);
            ' uzmi bas tu celiju, ne tudji format sa lista.
            srcFmt = staraBody.cells(1).NumberFormat
        Else
            Set srcCell = Nothing
            Set srcCell = staraBody.SpecialCells(xlCellTypeConstants)
            If srcCell Is Nothing Then
                srcFmt = staraBody.cells(1).NumberFormat
            Else
                srcFmt = srcCell.Areas(1).cells(1).NumberFormat
            End If
        End If
        dstFmt = novaBody.cells(1).NumberFormat
        If dstFmt = "General" And srcFmt <> "General" And Len(srcFmt) > 0 Then
            novaBody.NumberFormat = srcFmt
        End If
    End If
    Err.Clear
End Sub

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
        ' aktivacija/trial se NE migrira (vezana za masinu) -> ni ne udje u merge
        If Len(kkey) > 0 And Not JeLicencaKljuc(kkey) And Not dVal.Exists(kkey) Then
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
' tblLocalConfig) jer cuvaju aktuelna podesavanja (OAuth / SEF / bank putanje).
' Config tabele su key/value; novi prazan fajl ih ima prazne do SetupNewPC, pa pun
' copy starih redova donosi podesavanja bez gubitka -- OSIM AKTIVACIJE: LICENSE_* i
' TRIAL_* kljucevi se NE migriraju (vezani za masinu; nova masina re-aktivira).
Private Function SkipTabela(ByVal naziv As String) As Boolean
    SkipTabela = (LCase$(Left$(naziv, 6)) = "tblrpt")
End Function

' Override za PREIMENOVANE kolone izmedju verzija. Default = isto ime.
' Primer (ako je novo "Kontakt" u staroj verziji bilo "Telefon"):
'   If tabela = "tblStanice" And novoIme = "Kontakt" Then StaroImeKolone = "Telefon": Exit Function
Private Function StaroImeKolone(ByVal tabela As String, ByVal novoIme As String) As String
    StaroImeKolone = novoIme
End Function

' Provere posle kopiranja jedne tabele (best-effort; NIKAD ne baca gresku dalje).
' Svaka jedinica cisti Err pre i proverava ga posle -> provera koja NIJE izvedena
' se PRIJAVI (fail-closed), ne proguta se kao "sve u redu".
'   B) stare kolone SA PODATKOM koje nemaju cilj u novom (preimenovane/izbacene)
'   C) NOVE kolone bez izvora u starom (mapCol=0): vezne -> PROBLEM, audit/formula
'      se preskacu, ostale -> info
'   D) zbir CISTO numerickih kolona: staro vs novo (citano nazad) = da li je UPIS
'      legao; NE hvata pogresno mapiranje (obe strane bi citale istu kolonu)
' Nadje li problem: doda '!!'/'~' red u warn i uveca prob.
Private Sub ProveriKoloneIZbir(ByVal loStari As ListObject, ByVal loNovi As ListObject, _
                               ByRef SRC As Variant, ByRef mapCol() As Long, ByVal nNew As Long, _
                               ByRef warn As String, ByRef prob As Long)
    On Error Resume Next
    Dim nRows As Long: nRows = UBound(SRC, 1)
    Dim j As Long, k As Long
    Dim oc As Long, ob As Long, nc As Long, nb As Long
    Dim oldSum As Double, newSum As Double
    Dim imeStare As String, imeNove As String, imaPod As Boolean, hf As Variant
    Dim usedOld() As Boolean
    Dim DST As Variant

    ' B) stare kolone bez cilja u novom (rename/izbaceno), a imaju podatak
    ReDim usedOld(1 To loStari.ListColumns.count)
    For j = 1 To nNew
        If mapCol(j) > 0 Then usedOld(mapCol(j)) = True
    Next j
    For k = 1 To loStari.ListColumns.count
        If Not usedOld(k) Then
            Err.Clear
            imeStare = loStari.ListColumns(k).name
            imaPod = KolonaImaPodatak(SRC, k, nRows)
            If Err.Number <> 0 Then
                prob = prob + 1
                warn = warn & "     !! provera stare kolone #" & k & " NIJE izvedena (" & Err.description & ")" & vbCrLf
            ElseIf imaPod Then
                prob = prob + 1
                warn = warn & "     !! kolona '" & imeStare & _
                       "' ima podatke u starom a NEMA cilj u novom -> NIJE preneta" & vbCrLf
            End If
        End If
    Next k

    ' C) NOVE kolone bez izvora u starom. Vezne (*ID / Broj*) prazne = razvezan red
    '    (PROBLEM). Audit se pune unapred; kalkulisane (formula) nisu prazne -> preskoci.
    For j = 1 To nNew
        If mapCol(j) = 0 Then
            Err.Clear
            imeNove = loNovi.ListColumns(j).name
            If JeAuditKolona(imeNove) Then
                ' audit bez izvora = migracija sa pre-audit verzije; ocekivano, tiho
            Else
                hf = loNovi.ListColumns(j).DataBodyRange.HasFormula
                If Err.Number <> 0 Then
                    prob = prob + 1
                    warn = warn & "     !! provera nove kolone '" & imeNove & "' NIJE izvedena (" & Err.description & ")" & vbCrLf
                ElseIf hf = True Then
                    ' kalkulisana kolona - ocekivano, nije prazna
                ElseIf JeKriticnaKolona(loNovi.name, imeNove) Then
                    prob = prob + 1
                    warn = warn & "     !! VEZNA kolona '" & imeNove & _
                           "' nema izvor u starom -> red stize RAZVEZAN (proveri dok. lanac)" & vbCrLf
                Else
                    warn = warn & "     ~ nova kolona '" & imeNove & _
                           "' nema izvor u starom -> ostaje prazna" & vbCrLf
                End If
            End If
        End If
    Next j

    ' D) zbir CISTO numerickih kolona: staro (SRC) vs novo (citano nazad)
    Err.Clear
    DST = loNovi.DataBodyRange.value
    If Err.Number <> 0 Then
        prob = prob + 1
        warn = warn & "     !! provera zbira NIJE izvedena (citanje novog: " & Err.description & ")" & vbCrLf
        Exit Sub
    End If
    If Not IsArray(DST) Then
        Dim one(1 To 1, 1 To 1) As Variant: one(1, 1) = DST: DST = one
    End If
    For j = 1 To nNew
        If mapCol(j) > 0 Then
            Err.Clear
            oc = 0: ob = 0: nc = 0: nb = 0: oldSum = 0#: newSum = 0#
            oldSum = SumNumeric(SRC, mapCol(j), oc, ob)
            ' cisto numericka kolona = bar 1 broj i 0 ne-praznih ne-brojeva na strani starog
            If oc > 0 And ob = 0 Then
                newSum = SumNumeric(DST, j, nc, nb)
                If Err.Number <> 0 Then
                    prob = prob + 1
                    warn = warn & "     !! provera zbira za '" & loNovi.ListColumns(j).name & "' NIJE izvedena (" & Err.description & ")" & vbCrLf
                ElseIf Round(oldSum, 2) <> Round(newSum, 2) Then
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

' True ako kolona (SRC niz, indeks col) ima bar jednu ne-praznu celiju.
Private Function KolonaImaPodatak(ByRef arr As Variant, ByVal col As Long, ByVal nRows As Long) As Boolean
    Dim rr As Long
    For rr = 1 To nRows
        Select Case VarType(arr(rr, col))
            Case vbEmpty, vbNull, vbError
                ' prazno
            Case vbString
                If Len(Trim$(CStr(arr(rr, col)))) > 0 Then KolonaImaPodatak = True: Exit Function
            Case Else
                KolonaImaPodatak = True: Exit Function
        End Select
    Next rr
End Function

' Vezne (dokumentni lanac) kolone: nova takva kolona prazna = red stize razvezan
' (GetVerwaisteDokumente ga posle prijavi). NE ukljucuje audit (te se pune unapred).
Private Function JeKriticnaKolona(ByVal tabela As String, ByVal kolona As String) As Boolean
    Dim s As String: s = LCase$(Trim$(kolona))
    ' identifikacione/vezne kolone: zavrsavaju se na "id"
    ' (OtkupID, OtpremnicaID, PrijemnicaID, ZbirnaID, FakturaID, CorrectionID, SEFDocumentId...)
    If Len(s) >= 2 Then
        If Right$(s, 2) = "id" Then JeKriticnaKolona = True: Exit Function
    End If
    ' poslovni vezni kljucevi bez "id" sufiksa (dokumentni lanac + kompozitni ID-jevi:
    ' tblOtkup identitet = BrojDokumenta + Klasa; tblAmbalaza = DokumentID + DokumentTip)
    Select Case s
        Case "brojzbirne", "brojotpremnice", "brojprijemnice", "brojfakture", _
             "brojdokumenta", "brojbloka", "klasa", "dokumenttip", "doktip"
            JeKriticnaKolona = True
    End Select
End Function

' Audit kolone (EnsureAuditColumnsCore): pune se unapred iz sloja podataka, pa
' prazne posle migracije sa pre-audit verzije NISU problem -> preskacu se tiho.
Private Function JeAuditKolona(ByVal kolona As String) As Boolean
    Select Case LCase$(Trim$(kolona))
        Case "createdat", "createdby", "modifiedat", "modifiedby"
            JeAuditKolona = True
    End Select
End Function

' Aktivacija/trial: vezano za masinu, ne migrira se (nova masina re-aktivira).
Private Function JeLicencaKljuc(ByVal kljuc As String) As Boolean
    Dim s As String: s = UCase$(Trim$(kljuc))
    JeLicencaKljuc = (Left$(s, 8) = "LICENSE_" Or Left$(s, 6) = "TRIAL_")
End Function

' Snimi kopiju NOVOG fajla pre migracije u <putanja>\Backup\ (rollback). "" na neuspeh.
Private Function BackupPreMigracije() As String
    On Error GoTo EH
    If Len(ThisWorkbook.path) = 0 Then Exit Function        ' nije snimljen -> nema gde
    Dim bkDir As String: bkDir = ThisWorkbook.path & "\Backup"
    If Dir(bkDir, vbDirectory) = "" Then MkDir bkDir
    Dim nm As String
    nm = "AgriX_pre-migracija_" & APP_VERSION & "_" & Format$(Now, "yyyy-mm-dd_hhmm") & ".xlsm"
    ThisWorkbook.SaveCopyAs bkDir & "\" & nm
    BackupPreMigracije = bkDir & "\" & nm
    Exit Function
EH:
    BackupPreMigracije = ""
End Function

