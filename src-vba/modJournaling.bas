Attribute VB_Name = "modJournaling"
Option Explicit

' ============================================================
' modJournal - TX-Level CSV Journaling
'
' Schreibt jede AppendRow-Operation sofort als CSV-Zeile.
' Zweck: Crash-Recovery. Wenn Excel abstuerzt bevor gespeichert
' wird, koennen alle Transaktionen aus dem Journal reimportiert
' werden.
'
' Journal-Pfad: ThisWorkbook.Path & "\Journal\"
' Dateiname:    tblName_YYYY-MM-DD.csv (eine pro Tabelle pro Tag)
' Rotation:     Dateien aelter als 30 Tage werden bei App-Start geloescht
'
' WICHTIG: Journal-Write darf NIEMALS die eigentliche Operation
' blockieren. Daher: On Error Resume Next um den Write.
' ============================================================

Private Const JOURNAL_FOLDER As String = "Journal"
Private Const JOURNAL_MAX_DAYS As Long = 30
Private Const BACKUP_FOLDER As String = "Backup"
Private Const BACKUP_MAX_DAYS As Long = 30
' Starost sama nije dovoljna: backup se pravi na SVAKI start, pa 30 dana rada
' znaci desetine kopija po 10 MB. Mereno 12.09.2026: disk je pao na 1 GB od 233,
' a pad backupa je tog jutra oborio pokretanje aplikacije. Zato i gornja granica
' broja kopija -- brise se sto je starije OD BILO KOG od dva pravila.
Private Const BACKUP_MAX_KEEP As Long = 20

' ============================================================
' AutoSave state -- AR-002
' ============================================================
Private m_LastAutoSaveAt As Date
Private m_HasAutoSaved As Boolean
Private m_AutoSaveInProgress As Boolean

Private Const AUTOSAVE_DEBOUNCE_SECONDS As Long = 3

' AR-002a: deferred AutoSave scheduling (Application.OnTime)
Private Const AUTOSAVE_IDLE_SECONDS    As Long = 60    ' snimi 60s posle poslednje aktivnosti
Private Const AUTOSAVE_MAX_AGE_SECONDS As Long = 600   ' kapica za neprekidan unos
Private m_NextSaveTime  As Date
Private m_SaveScheduled As Boolean

' Test-mode: dok je TRUE, WriteJournalRow i MarkDirtyAndSchedule su no-op --
' da dev smoke-suite (npr. modAgrohemijaTests), koji mutira tabele pa radi
' rollback, ne ostavi CSV journal redove niti zakaze AutoSave posle rollback-a.
' Postavlja se ISKLJUCIVO iz test modula; produkcioni tok ga nikad ne dira.
Private m_TestModeQuiet As Boolean

Public Sub SetTestModeQuiet(ByVal onOff As Boolean)
    m_TestModeQuiet = onOff
End Sub

Public Function IsTestModeQuiet() As Boolean
    IsTestModeQuiet = m_TestModeQuiet
End Function

' ============================================================
' PUBLIC - Aufgerufen aus modDataAccess.AppendRow
' ============================================================

Public Sub WriteJournalRow(ByVal tblName As String, ByVal rowData As Variant)
    ' Schreibt eine komplette rowData-Zeile als CSV-Append
    ' Fehlschlag ist still - Journal darf nie die App blockieren

    If m_TestModeQuiet Then Exit Sub            ' test-mode: bez journal traga

    Dim journalPath As String
    Dim fileName As String
    Dim filePath As String
    Dim line As String
    Dim ff As Integer
    Dim i As Long
    
    On Error Resume Next
    
    ' Pfad bauen
    journalPath = ThisWorkbook.path & "\" & JOURNAL_FOLDER
    
    ' Ordner erstellen falls nicht vorhanden
    If Dir(journalPath, vbDirectory) = "" Then
        MkDir journalPath
    End If
    
    ' Dateiname: tblOtkup_2026-03-18.csv
    fileName = tblName & "_" & Format$(Date, "yyyy-mm-dd") & ".csv"
    filePath = journalPath & "\" & fileName
    
    ' Header schreiben wenn Datei neu ist
    If Dir(filePath) = "" Then
        ff = FreeFile
        Open filePath For Output As #ff
        
        ' Header: Timestamp + alle Spaltennamen der Tabelle
        Dim headers As Variant
        headers = GetTableHeaders(tblName)
        
        If Not IsEmpty(headers) Then
            line = "JournalTime"
            For i = LBound(headers) To UBound(headers)
                line = line & ";" & CStr(headers(i))
            Next i
            Print #ff, line
        End If
        
        Close #ff
    End If
    
    ' Datenzeile bauen: Timestamp + alle Werte
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    
    For i = LBound(rowData) To UBound(rowData)
        line = line & ";" & EscapeCSV(CStr(NzJournal(rowData(i), "")))
    Next i
    
    ' Append an Datei
    ff = FreeFile
    Open filePath For Append As #ff
    Print #ff, line
    Close #ff
    
    On Error GoTo 0
End Sub

' ============================================================
' PUBLIC - Rotation (aufgerufen aus modMain.StartApp)
' ============================================================

Public Sub PurgeOldJournals()
    ' Loescht Journal-Dateien die aelter als JOURNAL_MAX_DAYS sind
    
    Dim journalPath As String
    Dim fileName As String
    Dim filePath As String
    Dim fileDate As Date
    Dim datePart As String
    Dim pos As Long
    
    On Error Resume Next
    
    journalPath = ThisWorkbook.path & "\" & JOURNAL_FOLDER
    
    If Dir(journalPath, vbDirectory) = "" Then Exit Sub
    
    fileName = Dir(journalPath & "\*.csv")
    
    Do While fileName <> ""
        ' Datum aus Dateiname extrahieren: tblName_2026-03-18.csv
        pos = InStrRev(fileName, "_")
        If pos > 0 Then
            datePart = Mid$(fileName, pos + 1)
            datePart = Left$(datePart, 10)  ' "2026-03-18"
            
            If IsDate(datePart) Then
                fileDate = CDate(datePart)
                
                If DateDiff("d", fileDate, Date) > JOURNAL_MAX_DAYS Then
                    filePath = journalPath & "\" & fileName
                    Kill filePath
                End If
            End If
        End If
        
        fileName = Dir()
    Loop
    
    On Error GoTo 0
End Sub

' ============================================================
' PUBLIC - Recovery Check (aufgerufen aus modMain.StartApp)
' ============================================================

Public Function CheckJournalForRecovery() As String
    ' Prueft ob heute Journal-Eintraege existieren die nicht in Excel sind
    ' Returns: "" wenn alles OK, oder Warn-String mit Details
    
    Dim journalPath As String
    Dim fileName As String
    Dim filePath As String
    Dim ff As Integer
    Dim line As String
    Dim parts() As String
    Dim tblName As String
    Dim journalCount As Long
    Dim excelCount As Long
    Dim warnings As String
    Dim pos As Long
    Dim lo As ListObject
    
    On Error Resume Next
    
    journalPath = ThisWorkbook.path & "\" & JOURNAL_FOLDER
    
    If Dir(journalPath, vbDirectory) = "" Then
        CheckJournalForRecovery = ""
        Exit Function
    End If
    
    ' Nur heutige Dateien pruefen
    fileName = Dir(journalPath & "\*_" & Format$(Date, "yyyy-mm-dd") & ".csv")
    
    Do While fileName <> ""
        ' Tabellenname aus Dateiname extrahieren
        pos = InStrRev(fileName, "_")
        If pos > 0 Then
            tblName = Left$(fileName, pos - 1)
        Else
            GoTo NextFile
        End If
        
        filePath = journalPath & "\" & fileName
        
        ' Journal-Zeilen zaehlen (minus Header)
        journalCount = 0
        ff = FreeFile
        Open filePath For Input As #ff
        Do While Not EOF(ff)
            Line Input #ff, line
            journalCount = journalCount + 1
        Loop
        Close #ff
        journalCount = journalCount - 1  ' Header abziehen
        
        If journalCount < 0 Then journalCount = 0
        
        ' Excel-Zeilen zaehlen
        Set lo = GetTable(tblName)
        If lo Is Nothing Then
            excelCount = 0
        ElseIf lo.DataBodyRange Is Nothing Then
            excelCount = 0
        Else
            excelCount = lo.DataBodyRange.rows.count
        End If
        
        ' Wenn Journal mehr Zeilen hat als Excel ? potentieller Datenverlust
        If journalCount > excelCount Then
            If warnings <> "" Then warnings = warnings & vbCrLf
            warnings = warnings & tblName & ": Journal hat " & journalCount & _
                       " Eintraege, Excel hat " & excelCount & " Zeilen. " & _
                       "Moeglicher Datenverlust nach Absturz!"
        End If
        
NextFile:
        fileName = Dir()
    Loop
    
    On Error GoTo 0
    
    CheckJournalForRecovery = warnings
End Function

' ============================================================
' PUBLIC - File Backup (aufgerufen aus modMain.StartApp)
' ============================================================

' Vraca True ako backup postoji posle poziva, False ako nije napravljen.
'
' NIKAD NE PODIZE GRESKU. Do 12.09.2026. je zavrsavala sa Err.Raise, a zove se
' iz StartApp -- pa je pun disk oborio CELO pokretanje aplikacije: operater je
' dobio "Greska pri pokretanju" i nije mogao da radi nista. Backup je sigurnosna
' mreza, ne preduslov ispravnosti; njegov izostanak sme da smanji zastitu, ne da
' oduzme alat. (Suprotno vazi za MakePreImportBackup u modVbaTools: on stoji PRED
' destruktivnom operacijom, pa je tamo fail-closed tacan izbor.)
'
' Neuspeh se NE gubi: LogErr + Monitor_Backup FAILED su i dosad tu, a pozivalac
' (modMain.StartApp) sada na False javi operateru toast-om.
'
' ciljniFolder je test seam -- prazan znaci normalan "<sveska>\Backup".
Public Function BackupFileOnStart(Optional ByVal ciljniFolder As String = "") As Boolean

    Dim t0 As Single
    t0 = Timer

    Dim backupPath As String
    Dim srcPath As String
    Dim destName As String
    Dim destPath As String
    Dim baseName As String
    Dim ext As String
    Dim dotPos As Long
    
    On Error GoTo EH
    
    srcPath = ThisWorkbook.fullName
    If Len(ciljniFolder) > 0 Then
        backupPath = ciljniFolder
    Else
        backupPath = ThisWorkbook.path & "\" & BACKUP_FOLDER
    End If
    
    ' Ordner erstellen falls nicht vorhanden
    If Dir(backupPath, vbDirectory) = "" Then
        MkDir backupPath
    End If
    
    ' Basisname + Extension trennen
    baseName = ThisWorkbook.name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then
        ext = Mid$(baseName, dotPos)
        baseName = Left$(baseName, dotPos - 1)
    Else
        ext = ".xlsm"
    End If
    
    ' Zielname
    destName = baseName & "_" & Format$(Now, "yyyy-mm-dd\_hhmm") & ext
    destPath = backupPath & "\" & destName
    
    ' Nicht doppelt kopieren
    On Error Resume Next
    Dim existCheck As String
    existCheck = Dir(destPath)
    On Error GoTo EH
    
    If existCheck <> "" Then
        BackupFileOnStart = True      ' kopija za ovaj minut vec postoji
        Exit Function
    End If
    
    ' Kopieren
    ThisWorkbook.SaveCopyAs destPath

    ' Uspeh se TVRDI tek kad fajl stvarno postoji. Bez ovog reda je funkcija
    ' padala na podrazumevano False i StartApp je na svakom normalnom startu
    ' javljao operateru da backup nije napravljen -- laz u suprotnom smeru.
    If Len(Dir(destPath)) = 0 Then
        Err.Raise 53, "BackupFileOnStart", "SaveCopyAs nije prijavio gresku, ali fajl ne postoji: " & destPath
    End If
    BackupFileOnStart = True
    
    LogInfo "BackupFileOnStart", "Backup erstellt: " & destName
    ' Monitoring ide na mrezu. Iz testa se NE salje: bez ovog garda je RunAllTests
    ' visio 9 minuta na cekanju endpointa umesto da prijavi rezultat. Isti razlog
    ' zbog koga postoji IsTestModeQuiet za journal.
    On Error Resume Next
    If Not IsTestMode() Then
        Monitor_Backup _
            backupType:="STARTUP_BACKUP", _
            status:="SUCCESS", _
            backupLocation:="Startup backup completed", _
            durationMs:=CLng((Timer - t0) * 1000), _
            errorMessage:=""
    End If
    On Error GoTo 0

    Exit Function
EH:
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "modJournaling.BackupFileOnStart"
    On Error Resume Next
    If IsTestMode() Then GoTo BezMonitoringa      ' v. gard na uspesnom putu

    Monitor_Backup _
        backupType:="STARTUP_BACKUP", _
        status:="FAILED", _
        backupLocation:="Startup backup failed", _
        durationMs:=CLng((Timer - t0) * 1000), _
        errorMessage:=errDesc

    Monitor_Error _
        moduleName:="modMain", _
        procedureName:="BackupFileOnStart", _
        entityType:="Backup", _
        entityID:="STARTUP_BACKUP", _
        correlationId:="BACKUP-STARTUP", _
        errorNumber:=errNo, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

BezMonitoringa:
    ' Namerno BEZ Err.Raise -- v. zaglavlje procedure.
    BackupFileOnStart = False
End Function

' Vreme nastanka iz imena backup kopije -- i ujedno provera VLASNISTVA.
' Vraca 0 za sve sto nije jedan od dva kanonska oblika OVE sveske:
'
'   <baza>_YYYY-MM-DD_HHMM.xls*                     redovni startup backup
'   <baza>_pre-vba-import_YYYY-MM-DD_HHMMSS.xls*    kopija pred import (modVbaTools)
'
' Zasto tacan oblik, a ne prefiks: "<baza>*" bi u istom folderu pokupio i
' AgriX_DEV2_..., AgriX_DEV_old_..., AgriX_DEV_test_... To nije egzotika nego
' zatecen nacin rada -- na disku stoji desetak slicno imenovanih DEV kopija.
' Retention koji brise sme da bude samo NAJUZI moguci.
'
' Vreme, ne samo datum: backup se pravi na svaki start, pa dvadeset kopija istog
' dana ima isti datum. Sortiranje po datumu bi tada zavisilo od redosleda koji
' vrati Dir(), pa bi "sacuvaj najnovijih 20" cuvalo proizvoljnih 20.
'
' DateSerial/TimeSerial umesto CDate: ISO zapis "2026-03-18" kroz CDate zavisi od
' locale-a masine, a po ovoj vrednosti se BRISE fajl.
Public Function BackupVremeIzImena(ByVal ime As String, ByVal baza As String) As Date
    Dim s As String, tacka As Long, rep As String
    Dim g As Long, m As Long, d As Long, h As Long, mi As Long, sek As Long

    BackupVremeIzImena = 0
    If Len(baza) = 0 Then Exit Function

    tacka = InStrRev(ime, ".")
    If tacka <= 0 Then Exit Function
    If LCase$(Left$(Mid$(ime, tacka), 4)) <> ".xls" Then Exit Function
    s = Left$(ime, tacka - 1)

    If StrComp(Left$(s, Len(baza) + 1), baza & "_", vbTextCompare) <> 0 Then Exit Function
    rep = Mid$(s, Len(baza) + 2)

    If StrComp(Left$(rep, 15), "pre-vba-import_", vbTextCompare) = 0 Then
        rep = Mid$(rep, 16)
        If Len(rep) <> 17 Then Exit Function
        sek = DeoBroj(rep, 16, 2)
    ElseIf Len(rep) = 15 Then
        sek = 0
    Else
        Exit Function
    End If

    If Mid$(rep, 5, 1) <> "-" Or Mid$(rep, 8, 1) <> "-" Or Mid$(rep, 11, 1) <> "_" Then Exit Function
    g = DeoBroj(rep, 1, 4)
    m = DeoBroj(rep, 6, 2)
    d = DeoBroj(rep, 9, 2)
    h = DeoBroj(rep, 12, 2)
    mi = DeoBroj(rep, 14, 2)

    If g < 2000 Or g > 2999 Then Exit Function
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    If h > 23 Or mi > 59 Or sek > 59 Then Exit Function

    On Error GoTo EH
    BackupVremeIzImena = DateSerial(g, m, d) + TimeSerial(h, mi, sek)
    Exit Function
EH:
    BackupVremeIzImena = 0
End Function

' Ceo isecak mora biti cifra; -1 ako nije, pa provere opsega odbiju ime.
Private Function DeoBroj(ByVal s As String, ByVal od As Long, ByVal duz As Long) As Long
    Dim t As String, i As Long
    t = Mid$(s, od, duz)
    If Len(t) <> duz Then
        DeoBroj = -1
        Exit Function
    End If
    For i = 1 To duz
        If Mid$(t, i, 1) < "0" Or Mid$(t, i, 1) > "9" Then
            DeoBroj = -1
            Exit Function
        End If
    Next i
    DeoBroj = CLng(t)
End Function

' Koje kopije idu na brisanje. Cista odluka nad SPISKOM imena -- bez fajl-sistema,
' pa je merljiva testom.
'
' Dva pravila, brise se po BILO KOM:
'   starost  > BACKUP_MAX_DAYS
'   pozicija > BACKUP_MAX_KEEP  (od najnovije, po PUNOM vremenu)
'
' cuvajPreImport = True PINUJE sve pre-vba-import kopije, bez obzira na starost i
' broj. Zove se sa modImportState.ImportNijeDovrsen(): dok recovery nije zatvoren,
' modVbaTools u registru drzi pokazivac "poslednji siguran backup" (prevbackup)
' bas na jednu od njih, a RecoverImportState ga prikazuje operateru. Retention koji
' bi je obrisao ostavio bi poruku koja pokazuje na fajl kog nema -- i ponistio bas
' onu zastitu zbog koje #312 i #313 postoje. Dok traje opasnost, prostor je
' jeftiniji od oporavka.
'
' Vraca imena razdvojena sa vbLf ("" ako nema sta).
Public Function BackupZaBrisanje(ByVal imena As Variant, ByVal baza As String, _
                                 ByVal sada As Date, ByVal cuvajPreImport As Boolean) As String
    Dim i As Long, j As Long, n As Long
    Dim ime As Variant
    Dim spisak() As String, vremena() As Date
    Dim tS As String, tD As Date
    Dim out As String

    If Not IsArray(imena) Then Exit Function
    On Error GoTo EH

    ReDim spisak(0 To UBound(imena) - LBound(imena))
    ReDim vremena(0 To UBound(imena) - LBound(imena))
    n = 0
    For Each ime In imena
        tD = BackupVremeIzImena(CStr(ime), baza)
        If tD > 0 Then
            spisak(n) = CStr(ime)
            vremena(n) = tD
            n = n + 1
        End If
    Next ime
    If n = 0 Then Exit Function

    ' opadajuce po PUNOM vremenu; spisak je kratak pa je prosto umetanje dosta
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If vremena(j) > vremena(i) Then
                tD = vremena(i)
                vremena(i) = vremena(j)
                vremena(j) = tD
                tS = spisak(i)
                spisak(i) = spisak(j)
                spisak(j) = tS
            End If
        Next j
    Next i

    For i = 0 To n - 1
        If cuvajPreImport And InStr(1, spisak(i), "_pre-vba-import_", vbTextCompare) > 0 Then
            ' pinovano: aktivan recovery artefakt
        ElseIf DateDiff("d", vremena(i), sada) > BACKUP_MAX_DAYS Or (i + 1) > BACKUP_MAX_KEEP Then
            If Len(out) > 0 Then out = out & vbLf
            out = out & spisak(i)
        End If
    Next i

    BackupZaBrisanje = out
    Exit Function
EH:
    LogErr "modJournaling.BackupZaBrisanje"
End Function

' Obrise stare kopije OVE sveske iz Backup foldera.
Public Sub PurgeOldBackups()
    Dim backupPath As String, baseName As String
    Dim fileName As String, spisak As Collection
    Dim zaBrisanje As Variant, x As Variant
    Dim niz() As String, i As Long
    Dim obrisano As Long, pali As Long
    Dim dotPos As Long

    On Error GoTo EH

    backupPath = ThisWorkbook.path & "\" & BACKUP_FOLDER
    If Dir(backupPath, vbDirectory) = "" Then Exit Sub

    baseName = ThisWorkbook.name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then baseName = Left$(baseName, dotPos - 1)

    Set spisak = New Collection
    fileName = Dir(backupPath & "\*.xls*")
    Do While fileName <> ""
        spisak.Add fileName
        fileName = Dir()
    Loop
    If spisak.count = 0 Then Exit Sub

    ReDim niz(0 To spisak.count - 1)
    For i = 1 To spisak.count
        niz(i - 1) = spisak(i)
    Next i

    ' Filtar vlasnistva i oblika je u BackupVremeIzImena; ovde se prosledjuje samo
    ' ime sveske i stanje recovery markera.
    zaBrisanje = Split(BackupZaBrisanje(niz, baseName, Now, modImportState.ImportNijeDovrsen()), vbLf)
    For Each x In zaBrisanje
        If Len(Trim$(CStr(x))) > 0 Then
            On Error Resume Next
            Err.Clear
            Kill backupPath & "\" & CStr(x)
            If Err.Number = 0 Then
                obrisano = obrisano + 1
            Else
                pali = pali + 1
            End If
            On Error GoTo EH
        End If
    Next x

    ' Neuspelo brisanje se ne precutkuje: folder koji raste bez reci je i doveo do
    ' punog diska. Jedan zbirni red, ne red po fajlu.
    If pali > 0 Then
        LogError "modJournaling.PurgeOldBackups", _
                 "Backup retention: obrisano " & obrisano & ", NIJE uspelo " & pali & _
                 " (fajl zauzet ili nema prava).", 0, "WARN"
    ElseIf obrisano > 0 Then
        LogInfo "PurgeOldBackups", "Backup retention: obrisano " & obrisano & " kopija."
    End If
    Exit Sub
EH:
    LogErr "modJournaling.PurgeOldBackups"
End Sub

' ============================================================
' PRIVATE HELPERS
' ============================================================

Private Function EscapeCSV(ByVal s As String) As String
    ' CSV-Escape: Wenn Semikolon, Anfuehrungszeichen oder Newline enthalten
    If InStr(s, ";") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCrLf) > 0 Or InStr(s, vbLf) > 0 Then
        s = Replace(s, """", """""")
        EscapeCSV = """" & s & """"
    Else
        EscapeCSV = s
    End If
End Function

Private Function NzJournal(ByVal v As Variant, Optional ByVal Fallback As Variant = "") As Variant
    If isError(v) Then
        NzJournal = Fallback
    ElseIf IsNull(v) Then
        NzJournal = Fallback
    ElseIf IsEmpty(v) Then
        NzJournal = Fallback
    Else
        NzJournal = v
    End If
End Function

' ============================================================
' PUBLIC - AutoSave after TX commit
'
' Pozvano iz clsTransaction.CommitTx posle uspesnog commit-a.
' Best-effort save: greska ne sme da propaga jer je TX vec commit-ovan
' u memoriji i operator ne sme da vidi failure za save koji je tehnicki
' uspeo na nivou poslovne logike.
'
' Debounce sprecava rapid-fire saves u istom rafalu (npr. tri sukcesivna
' otkupa u 5 sekundi). Globalni state znaci da debounce vazi kroz ceo
' Excel session bez obzira koja clsTransaction ga okida.
' ============================================================

Public Sub AutoSaveAfterCommit(ByVal sourceName As String, _
                               Optional ByVal force As Boolean = False)
    Dim prevAlerts As Boolean
    Dim alertsTouched As Boolean
    
    On Error GoTo EH
    
    ' Reentrancy guard -- set BEFORE any other check.
    ' Guards against Excel events firing during Save that might re-enter here.
    If m_AutoSaveInProgress Then Exit Sub
    m_AutoSaveInProgress = True
    
    If (Not force) And (Not ShouldAutoSaveNow()) Then
        ' Silent skip when debounce active. Log on INFO level for traceability.
        LogInfo "AutoSaveAfterCommit", _
                "Skipped (debounce). Source=" & sourceName
        GoTo CleanExit
    End If
    
    If ThisWorkbook.ReadOnly Then
        LogWarn "AutoSaveAfterCommit", _
                "Workbook read-only. AutoSave skipped. Source=" & sourceName
        GoTo CleanExit
    End If

    ' Prekinut VBA import -> projekat je mozda NEPOTPUN, a Save ga betonira.
    ' Provera ide POSLE debounce-a i read-only kapije, ali PRE svakog dodira
    ' sveske. Vlasnik markera i razlog: modImportState.ImportNijeDovrsen.
    '
    ' Cena je svesna i ide u log, ne u tisinu: CommitTx prepusta stvarni upis
    ' ovoj proceduri (clsTransaction.cls), pa dok marker stoji commitovani
    ' podaci zive samo u memoriji. Nepotpun VBA projekat je gora steta --
    ' on prezivi zatvaranje fajla, a nesnimljen red se moze uneti ponovo.
    If modImportState.ImportNijeDovrsen() Then
        LogWarn "AutoSaveAfterCommit", _
                "PREKINUT VBA IMPORT -- AutoSave preskocen da ne bi snimio " & _
                "nepotpun projekat. Dovrsi ImportAllVBA ili vrati backup. " & _
                "Source=" & sourceName
        GoTo CleanExit
    End If
    
    If Len(Trim$(ThisWorkbook.path)) = 0 Then
        LogWarn "AutoSaveAfterCommit", _
                "Workbook has no path. AutoSave skipped. Source=" & sourceName
        GoTo CleanExit
    End If
    
    ' Suppress Compatibility Checker / external link prompts during Save.
    ' Must be restored on every exit path including EH.
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    alertsTouched = True
    
    ThisWorkbook.Save
    
    Application.DisplayAlerts = prevAlerts
    alertsTouched = False
    
    m_LastAutoSaveAt = Now
    m_HasAutoSaved = True
    
    LogInfo "AutoSaveAfterCommit", _
            "Saved after TX commit. Source=" & sourceName

CleanExit:
    If alertsTouched Then Application.DisplayAlerts = prevAlerts
    m_AutoSaveInProgress = False
    Exit Sub
    
EH:
    ' Critical: AutoSave failure must NEVER propagate. The TX is already
    ' committed in memory; operator must not see save failure for a save
    ' that succeeded at the business-logic level.
    LogErr "AutoSaveAfterCommit"
    
    If alertsTouched Then Application.DisplayAlerts = prevAlerts
    m_AutoSaveInProgress = False
    ' Intentionally no Err.Raise.
End Sub

' ============================================================
' AR-002a: Deferred AutoSave (Application.OnTime)
'
' CommitTx vise ne zove AutoSaveAfterCommit sinhrono -> MsgBox ne ceka
' ThisWorkbook.Save. Stvarni save i dalje radi AutoSaveAfterCommit
' (jedino mesto sa log stringovima/gardama -> runbook ostaje isti).
' Dirty-flag je ugradjeni ThisWorkbook.Saved -> bez dodatnog brojaca.
' Schedule/Cancel obrazac preslikan iz modStanicaLock heartbeat-a.
' ============================================================

Public Sub MarkDirtyAndSchedule(ByVal sourceName As String)
    On Error Resume Next

    If m_TestModeQuiet Then Exit Sub            ' test-mode: ne zakazuj AutoSave

    Dim delaySec As Long
    delaySec = AUTOSAVE_IDLE_SECONDS               ' uvek 60s posle poslednje aktivnosti

    ' Force-save SAMO ako neprekidan unos traje duze od MAX_AGE od poslednjeg save-a.
    If m_HasAutoSaved And _
       DateDiff("s", m_LastAutoSaveAt, Now) >= AUTOSAVE_MAX_AGE_SECONDS Then
        delaySec = 0
    End If

    ScheduleAutoSaveTimer delaySec
    LogInfo "AutoSaveAfterCommit", _
            "Scheduled deferred save in " & delaySec & "s. Source=" & sourceName
End Sub

Public Sub AutoSaveTick()
    ' Application.OnTime callback. MORA biti Public u standardnom modulu.
    m_SaveScheduled = False
    If ThisWorkbook.Saved Then Exit Sub        ' nista nesnimljeno -> ne snimaj
    AutoSaveAfterCommit "autosave-timer"       ' force=False -> postojeci debounce/log
End Sub

Public Sub FlushNow(ByVal sourceName As String)
    ' Trenutni flush na granici: section switch / dashboard / shutdown.
    ' Hvata i ne-TX izmene (maticni/SEF/config) jer gleda Saved, ne TX-trigger.
    CancelAutoSaveTimer
    If ThisWorkbook.Saved Then Exit Sub
    AutoSaveAfterCommit sourceName, True        ' force=True -> bypass debounce
End Sub

Public Sub StopAutoSaveTimer()
    CancelAutoSaveTimer
End Sub

Private Sub ScheduleAutoSaveTimer(ByVal delaySec As Long)
    CancelAutoSaveTimer                          ' uvek otkazi prethodni pre novog
    If delaySec < 0 Then delaySec = 0
    m_NextSaveTime = Now + TimeSerial(0, 0, delaySec)
    On Error Resume Next
    Application.OnTime m_NextSaveTime, "modJournaling.AutoSaveTick"
    On Error GoTo 0
    m_SaveScheduled = True
End Sub

Private Sub CancelAutoSaveTimer()
    If Not m_SaveScheduled Then Exit Sub
    On Error Resume Next
    Application.OnTime m_NextSaveTime, "modJournaling.AutoSaveTick", , False
    On Error GoTo 0
    m_SaveScheduled = False
End Sub


' ============================================================
' PRIVATE - Debounce check
' ============================================================

Private Function ShouldAutoSaveNow() As Boolean
    ' First call always saves. Subsequent calls within debounce window skip.
    
    If Not m_HasAutoSaved Then
        ShouldAutoSaveNow = True
        Exit Function
    End If
    
    ShouldAutoSaveNow = _
        (DateDiff("s", m_LastAutoSaveAt, Now) >= AUTOSAVE_DEBOUNCE_SECONDS)
End Function


' ============================================================
' PUBLIC - Test/diagnostic accessors
'
' Used by the smoke test to verify behavior without exposing
' internal state to business modules.
' ============================================================

Public Function HasAutoSavedAtLeastOnce() As Boolean
    HasAutoSavedAtLeastOnce = m_HasAutoSaved
End Function

Public Function GetLastAutoSaveAt() As Date
    GetLastAutoSaveAt = m_LastAutoSaveAt
End Function

Public Sub ResetAutoSaveStateForTests()
    ' Dev-only. Resets debounce state so tests can verify first-save behavior.
    m_LastAutoSaveAt = 0
    m_HasAutoSaved = False
    m_AutoSaveInProgress = False
    CancelAutoSaveTimer                          ' <-- DODATO
End Sub


' ============================================================
' PUBLIC - Smoke test
'
' Run via:
'   ?TestAutoSaveSmoke
'
' Returns string report. Follows the Test_* pattern already in use
' (e.g. modSEFClient.Test_SubmitUBLInvoice).
'
' This is an integration smoke - it actually saves the workbook
' once to verify the save path works. Run it on a workbook that
' is already in a clean savable state.
' ============================================================

Public Function TestAutoSaveSmoke() As String
    Dim report As String
    Dim p As Long, f As Long
    Dim before As Date
    Dim after As Date
    
    report = "AutoSave Smoke Test - " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
             String$(60, "-") & vbCrLf
    
    ' Reset state for deterministic test run
    Call ResetAutoSaveStateForTests
    
    ' --- Test 1: First call saves
    Call AutoSaveAfterCommit("TestAutoSaveSmoke.Test1")
    Call AssertSmoke("First call saves", _
                     HasAutoSavedAtLeastOnce(), True, report, p, f)
    
    ' --- Test 2: Immediate second call is debounced
    before = GetLastAutoSaveAt()
    Call AutoSaveAfterCommit("TestAutoSaveSmoke.Test2")
    after = GetLastAutoSaveAt()
    Call AssertSmoke("Immediate second call debounced", _
                     (after = before), True, report, p, f)
    
    ' --- Test 3: After debounce window, save fires again
    Application.Wait Now + TimeSerial(0, 0, 4)
    before = GetLastAutoSaveAt()
    Call AutoSaveAfterCommit("TestAutoSaveSmoke.Test3")
    after = GetLastAutoSaveAt()
    Call AssertSmoke("Save after debounce window", _
                     (after > before), True, report, p, f)
    
    ' --- Test 4: Reentrancy flag clears after success
    Call AutoSaveAfterCommit("TestAutoSaveSmoke.Test4a")
    Application.Wait Now + TimeSerial(0, 0, 4)
    before = GetLastAutoSaveAt()
    Call AutoSaveAfterCommit("TestAutoSaveSmoke.Test4b")
    after = GetLastAutoSaveAt()
    Call AssertSmoke("Reentrancy flag clears between calls", _
                     (after > before), True, report, p, f)
    
    report = report & String$(60, "-") & vbCrLf
    report = report & "PASS: " & p & "  FAIL: " & f & vbCrLf
    
    TestAutoSaveSmoke = report
    Debug.Print report
End Function

Private Sub AssertSmoke(ByVal testName As String, _
                       ByVal actual As Boolean, _
                       ByVal expected As Boolean, _
                       ByRef report As String, _
                       ByRef p As Long, _
                       ByRef f As Long)
    If actual = expected Then
        p = p + 1
        report = report & "  [PASS] " & testName & vbCrLf
    Else
        f = f + 1
        report = report & "  [FAIL] " & testName & _
                 " - expected=" & expected & " actual=" & actual & vbCrLf
    End If
End Sub

