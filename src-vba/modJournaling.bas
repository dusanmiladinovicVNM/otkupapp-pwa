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

' ============================================================
' PUBLIC - Aufgerufen aus modDataAccess.AppendRow
' ============================================================

Public Sub WriteJournalRow(ByVal tblName As String, ByVal rowData As Variant)
    ' Schreibt eine komplette rowData-Zeile als CSV-Append
    ' Fehlschlag ist still - Journal darf nie die App blockieren
    
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

Public Sub BackupFileOnStart()

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
    backupPath = ThisWorkbook.path & "\" & BACKUP_FOLDER
    
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
        Exit Sub
    End If
    
    ' Kopieren
    ThisWorkbook.SaveCopyAs destPath
    
    LogInfo "BackupFileOnStart", "Backup erstellt: " & destName
    On Error Resume Next
        Monitor_Backup _
            backupType:="STARTUP_BACKUP", _
            status:="SUCCESS", _
            backupLocation:="Startup backup completed", _
            durationMs:=CLng((Timer - t0) * 1000), _
            errorMessage:=""
    On Error GoTo 0

Exit Sub
    Exit Sub
EH:
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

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

    LogErr "modJournaling.BackupFileOnStart"

    On Error GoTo 0
    Err.Raise errNo, errSrc, errDesc
End Sub

Public Sub PurgeOldBackups()
    ' Loescht Backup-Dateien die aelter als BACKUP_MAX_DAYS sind
    ' Basiert auf Dateiname-Datum, nicht File-System-Datum
    
    Dim backupPath As String
    Dim fileName As String
    Dim filePath As String
    Dim datePart As String
    Dim fileDate As Date
    Dim pos As Long
    
    On Error Resume Next
    
    backupPath = ThisWorkbook.path & "\" & BACKUP_FOLDER
    
    If Dir(backupPath, vbDirectory) = "" Then Exit Sub
    
    fileName = Dir(backupPath & "\*.xls*")
    
    Do While fileName <> ""
        ' Datum aus Dateiname extrahieren: ..._2026-03-18_0845.xlsm
        ' Suche das Muster _YYYY-MM-DD_ (11 Zeichen vor der letzten _HHMM)
        pos = InStrRev(fileName, ".")
        If pos > 5 Then
            ' 5 Zeichen vor dem Punkt: _0845
            ' 11 Zeichen davor: _2026-03-18
            datePart = Mid$(fileName, pos - 15, 10)  ' "2026-03-18"
            
            If IsDate(datePart) Then
                fileDate = CDate(datePart)
                
                If DateDiff("d", fileDate, Date) > BACKUP_MAX_DAYS Then
                    filePath = backupPath & "\" & fileName
                    Kill filePath
                End If
            End If
        End If
        
        fileName = Dir()
    Loop
    
    On Error GoTo 0
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

