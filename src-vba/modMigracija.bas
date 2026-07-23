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

    ' Pre-migracija backup je OBAVEZAN: ovo je destruktivan alat, bez potvrdjenog
    ' backup-a nema sigurnog povratka -> radije PREKINI (fail-closed) nego na slepo.
    If Len(ThisWorkbook.path) = 0 Then
        MsgBox "Fajl nije snimljen na disk pa ne mogu da napravim backup." & vbCrLf & _
               "Prvo snimi novi fajl (Save As) pa ponovo pokreni migraciju.", _
               vbExclamation, "Migracija - prekinuto (nema backup-a)"
        Exit Sub
    End If
    Dim bkPath As String: bkPath = BackupPreMigracije()
    If Len(bkPath) = 0 Then
        MsgBox "Backup NIJE uspeo (disk pun / zakljucan / nema prava upisa)." & vbCrLf & _
               "Migracija je PREKINUTA - bez backup-a nema sigurnog povratka.", _
               vbCritical, "Migracija - prekinuto (backup neuspeo)"
        Exit Sub
    End If
    summary = summary & "  (backup: " & bkPath & ")" & vbCrLf

    ' Otkazi eventualni ZAKAZAN AutoSave (modJournaling OnTime) da posle migracije
    ' NE upise automatski (mozda problematican) rezultat. Migracija ne ide kroz TX
    ' sloj pa ga sama ne zakazuje; ostaje samo da otkazemo zatecen tajmer. Best-effort.
    On Error Resume Next
    StopAutoSaveTimer
    On Error GoTo 0

    ' Excelov (OneDrive/SharePoint) AutoSave je ZASEBAN mehanizam od modJournaling
    ' tajmera i pise NEZAVISNO od VBA -> ugasi ga za ovu sesiju, inace "fajl ostaje
    ' prljav = svesna kapija" NE vazi za cloud klijente (Excel sam upise rezultat pre
    ' nego operater procita PROBLEMI). Vraca se SAMO na cistom uspehu (vidi CLEAN);
    ' na problem/gresku ostaje ugasen da auto-save ne persistuje pre svesne odluke.
    ' LATE-BIND (Object): AutoSaveOn postoji tek u Excel 2016+/365. Na starijem Excel-u
    ' early-bound "ThisWorkbook.AutoSaveOn" je COMPILE greska ("member not found") koju
    ' On Error NE hvata i koja obori CEO projekat; late-bound baca runtime 438 koji se
    ' ovde uhvati i preskoci -> kompajlira se svuda, gasi cloud AutoSave gde postoji.
    Dim wbLate As Object: Set wbLate = ThisWorkbook
    Dim prevAutoSave As Boolean, hadAutoSave As Boolean
    On Error Resume Next
    prevAutoSave = wbLate.AutoSaveOn              ' 438 na starom Excel-u / greska = nije cloud
    hadAutoSave = (Err.Number = 0)
    If hadAutoSave And prevAutoSave Then wbLate.AutoSaveOn = False
    Err.Clear
    On Error GoTo 0

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
    Application.Cursor = xlWait                          ' vidljiv signal da radi (ScreenUpdating je off)

    Set stari = Workbooks.Open(fileName:=CStr(putanja), ReadOnly:=True, UpdateLinks:=0)

    ' Sanity: da li je izabran BAS OtkupApp fajl? Inace svaka tabela vrati -1,
    ' problems ostane 0, i "0 redova" izgleda kao uspeh. tblOtkup = potpis baze.
    If NadjiListObject(stari, "tblOtkup") Is Nothing Then
        problems = problems + 1
        summary = summary & "  !! izabrani fajl NEMA tblOtkup -> nije OtkupApp baza? (nista nije preneto)" & vbCrLf
        GoTo CLEAN
    End If

    ' --- Licenca: SAME-MACHINE gate ---
    ' Aktivacija (LICENSE_KEY/BOUND_PARTS/...) prelazi SAMO ako je stara licenca
    ' vezana za OVU masinu (isti prag kao modLicense: LicPartsMatch >= LIC_MIN_MATCH=2).
    ' Na drugoj/neutvrdjenoj masini se preskace (BOUND_PARTS nosi otisak STARE masine
    ' -> nova bi se zakljucala). Config licence (ENABLED/ENDPOINT) prelazi uvek.
    Application.StatusBar = "Migracija: provera licence (masina) ..."
    Dim oldBound As String
    oldBound = StaroConfigVrednost(stari, "tblSEFConfig", "ConfigKey", "ConfigValue", "LICENSE_BOUND_PARTS")
    Dim sameMachine As Boolean
    If Len(oldBound) > 0 Then
        sameMachine = JeIstaMasina(oldBound)
        If sameMachine Then
            summary = summary & "  (licenca: ISTA masina -> aktivacija se prenosi)" & vbCrLf
        Else
            summary = summary & "  (licenca: druga/neutvrdjena masina -> aktivacija NE prelazi; re-aktivacija)" & vbCrLf
        End If
    End If

    Dim ws As Worksheet, loNovi As ListObject, n As Long
    Dim ckey As String, cval As String, isCfg As Boolean, eDesc As String, errNum As Long
    Dim warn As String
    For Each ws In novi.Worksheets
        For Each loNovi In ws.ListObjects
            If Not SkipTabela(loNovi.name) Then
                isCfg = ConfigKolone(loNovi.name, ckey, cval)
                warn = ""
                Application.StatusBar = "Migracija: " & loNovi.name & " (" & (tbls + 1) & ") ..."

                On Error Resume Next                       ' jedna losa tabela ne prekida ceo prolaz
                Err.Clear
                If isCfg Then
                    n = MergeConfigTabelu(stari, loNovi, ckey, cval, warn, problems, sameMachine)
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
                    If Len(warn) > 0 Then summary = summary & warn
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
    Application.StatusBar = False
    Application.Cursor = xlDefault
    ' cloud AutoSave: vrati SAMO na cistom uspehu; na problem/gresku ostaje UGASEN da
    ' auto-save ne persistuje pre nego operater svesno odluci (snimi ili odbaci).
    ' Sledeci put kad se fajl otvori AutoSave se sam vrati (per-sesija property).
    If hadAutoSave And prevAutoSave And problems = 0 And Len(em) = 0 Then wbLate.AutoSaveOn = True
    On Error GoTo 0

    ' NE diramo ThisWorkbook.Saved. (Saved=True bi Excelu reklo "nema izmena" pa bi
    ' zatvaranje PROSLO bez pitanja i TIHO izgubilo migraciju - suprotno od namere.)
    ' Radni fajl ostaje "prljav" -> Excel pri zatvaranju NORMALNO pita "Snimi?", sto
    ' je svesna kapija; AutoSave je otkazan na pocetku, a backup je napravljen pre
    ' izmena, pa je i tih upis oporaviv.
    Dim hdr As String
    If problems > 0 Then hdr = "!! PROBLEMI: " & problems & " -> vidi '!!' i '~' redove dole !!" & vbCrLf & vbCrLf
    Dim foot As String
    If problems > 0 Or Len(em) > 0 Then
        foot = "Rezultat NIJE snimljen. PRE snimanja proveri gornje probleme." & vbCrLf & _
               "Ako nesto nije u redu: zatvori BEZ snimanja (stari fajl je netaknut)." & vbCrLf & _
               "Backup pre migracije: " & bkPath
    Else
        foot = "Sada SNIMI novi fajl (Ctrl+S) i proveri par tabela." & vbCrLf & _
               "Backup pre migracije: " & bkPath
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

    ' osiguraj TACNO nRows redova u novoj tabeli - robustno preko ListRows.Add/Delete.
    ' NAMERNO NE koristimo ListObject.Resize: on ume TIHO da prosiri tabelu preko
    ' obicnog sadrzaja ispod nje (napomene/rucni unos/formule/merged/shapes) koji
    ' CountA ne pokriva u potpunosti; nemapirane kolone bi zadrzale tudji sadrzaj a
    ' provera C bi ga smatrala "ocekivano praznim". Za destruktivan alat robusnost >
    ' brzina (isti razlog kao u originalu). Ako Add postane usko grlo na ogromnoj
    ' bazi, resenje je mereno + posebno guardovano, ne slepi Resize.
    Do While loNovi.ListRows.count > nRows
        loNovi.ListRows(loNovi.ListRows.count).Delete
    Loop
    Dim addedCnt As Long
    Do While loNovi.ListRows.count < nRows
        loNovi.ListRows.Add
        addedCnt = addedCnt + 1
        If addedCnt Mod 1000 = 0 Then _
            Application.StatusBar = "Migracija: " & loNovi.name & " - red " & addedCnt & "/" & nRows
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
        ' Format CELE stare kolone (ne pojedinacne celije). Range.NumberFormat vrati
        ' zajednicki format ako je uniforman (obicno jeste za datum/iznos kolonu),
        ' ili Null ako je mesan. Ovako se hvata i format-samo-formula kolona i kolona
        ' sa vodecim praznim redovima -- sto SpecialCells(xlCellTypeConstants) NE bi
        ' (promasi formule i, nad jednoceliskim opsegom, radi nad celim sheet-om).
        Dim rawSrc As Variant, rawDst As Variant, srcFmt As String, dstFmt As String
        rawSrc = staraBody.NumberFormat
        If IsNull(rawSrc) Then srcFmt = staraBody.cells(1).NumberFormat Else srcFmt = CStr(rawSrc)
        rawDst = novaBody.NumberFormat
        If IsNull(rawDst) Then dstFmt = "" Else dstFmt = CStr(rawDst)   ' mesan novi -> ne diraj
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
                                   ByVal keyCol As String, ByVal valCol As String, _
                                   ByRef warn As String, ByRef prob As Long, _
                                   ByVal sameMachine As Boolean) As Long
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
        ' aktivacija/trial: na ISTOJ masini prelazi; na drugoj se NE migrira (vezana
        ' za masinu). sameMachine=True zaobilazi JeLicencaKljuc filter.
        If Len(kkey) > 0 And (sameMachine Or Not JeLicencaKljuc(kkey)) And Not dVal.Exists(kkey) Then
            dVal(kkey) = loStari.DataBodyRange.cells(i, sVal).value
            dRow(kkey) = i
            ' (samo DRUGA masina) nov license-slican kljuc koji NIJE u aktivacionoj
            ' listi ni poznat config -> PRENET, ali PRIJAVLJEN (da nov aktivacioni
            ' kljuc ne otputuje tiho na novu masinu). Na istoj masini sve i onako prelazi.
            If (Not sameMachine) And JeLicencaPrefiks(kkey) And Not JeLicencaConfigPoznat(kkey) Then
                prob = prob + 1
                warn = warn & "     ~ config kljuc '" & kkey & "' lici na licencni a nije u listi -> PRENET, proveri (mozda dodati u JeLicencaKljuc)" & vbCrLf
            End If
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
' copy starih redova donosi podesavanja bez gubitka. AKTIVACIJA licence/trial-a
' (machine-bound kljucevi, vidi JeLicencaKljuc) prelazi SAMO na ISTOJ masini
' (sameMachine gate); config licence (ENABLED/ENDPOINT/DANI) prelazi uvek.
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

' AKTIVACIJA/binding/stanje licence: migrira se SAMO na ISTOJ masini (sameMachine
' gate u MergeConfigTabelu preko JeIstaMasina). Na drugoj/neutvrdjenoj masini se
' preskace (nova bi se zakljucala -- BOUND_PARTS nosi otisak stare masine).
' EKSPLICITNA lista (ne prefiks) -- da NE pokupi legitiman config koji se sme uvek
' preneti (LICENSE_ENABLED/ENDPOINT, TRIAL_ENABLED/DAYS). Kljucevi potvrdjeni u
' modLicense/modTrial (CFG_LIC_*, TRIAL_*). Nov, JOS NEPOZNAT LICENSE_*/TRIAL_*
' kljuc se (na drugoj masini) prenese ali PRIJAVI (JeLicencaPrefiks +
' JeLicencaConfigPoznat) -- da tiho ne otputuje na novu masinu.
' NAPOMENA (ostaje otvoreno): ovo ne CISTI eventualno zatecenu aktivaciju u NOVOM
' sablonu (target-clean).
Private Function JeLicencaKljuc(ByVal kljuc As String) As Boolean
    Select Case UCase$(Trim$(kljuc))
        Case "LICENSE_KEY", "LICENSE_TOKEN", "LICENSE_BOUND_PARTS", _
             "LICENSE_HWM", "LICENSE_STATUS", "LICENSE_NEXT_CHECK", _
             "TRIAL_START", "TRIAL_HWM"
            JeLicencaKljuc = True
    End Select
End Function

' Lici na licencni/trial kljuc (prefiks). Koristi se da nov, jos nepoznat
' LICENSE_*/TRIAL_* kljuc (npr. buduci LICENSE_DEVICE_ID) ne otputuje TIHO:
' prenese se (bez lazne blokade) ali se PRIJAVI da se doda u odgovarajucu listu.
Private Function JeLicencaPrefiks(ByVal kljuc As String) As Boolean
    Dim s As String: s = UCase$(Trim$(kljuc))
    JeLicencaPrefiks = (Left$(s, 8) = "LICENSE_" Or Left$(s, 6) = "TRIAL_")
End Function

' POZNAT config licence/trial-a koji se SME preneti (nije aktivacija/binding).
Private Function JeLicencaConfigPoznat(ByVal kljuc As String) As Boolean
    Select Case UCase$(Trim$(kljuc))
        Case "LICENSE_ENABLED", "LICENSE_ENDPOINT", "TRIAL_ENABLED", "TRIAL_DAYS"
            JeLicencaConfigPoznat = True
    End Select
End Function

' True ako je STARA vezana licenca vezana za OVU masinu -> aktivacija se sme preneti.
' Reuse modLicense: isti otisak (GetDeviceParts) i isti prag (LicPartsMatch >= 2 =
' LIC_MIN_MATCH) kao sto sama provera licence koristi. Greska / ne moze da utvrdi ->
' False (bezbedan default: ne prenosi aktivaciju kad nismo sigurni).
Private Function JeIstaMasina(ByVal oldBound As String) As Boolean
    On Error GoTo done
    If Len(Trim$(oldBound)) = 0 Then Exit Function
    JeIstaMasina = (LicPartsMatch(GetDeviceParts(), oldBound) >= 2)   ' 2 = modLicense.LIC_MIN_MATCH
done:
End Function

' Vrednost kljuca iz key/value config tabele STAROG fajla ("" ako nema/greska).
Private Function StaroConfigVrednost(ByVal stari As Workbook, ByVal tabela As String, _
                                     ByVal keyCol As String, ByVal valCol As String, _
                                     ByVal kljuc As String) As String
    On Error GoTo done
    Dim lo As ListObject: Set lo = NadjiListObject(stari, tabela)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim kc As Long: kc = ColIndexByName(lo, keyCol)
    Dim vc As Long: vc = ColIndexByName(lo, valCol)
    If kc = 0 Or vc = 0 Then Exit Function
    Dim i As Long
    For i = 1 To lo.ListRows.count
        If StrComp(Trim$(CStr(lo.DataBodyRange.cells(i, kc).value)), kljuc, vbTextCompare) = 0 Then
            StaroConfigVrednost = Trim$(CStr(lo.DataBodyRange.cells(i, vc).value))
            Exit Function
        End If
    Next i
done:
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

