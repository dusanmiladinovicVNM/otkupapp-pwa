Option Explicit

' ============================================================
' modMain v2.1 – ValidateAllTables aktualisiert
' ============================================================

Private m_Initialized As Boolean
Private mIsShuttingDown As Boolean

Public Sub StartApp()
    On Error GoTo EH

    On Error Resume Next
    Monitor_Event _
        eventType:="VBA_STARTAPP_START", _
        severity:="INFO", _
        message:="StartApp started", _
        userId:="Operator", _
        moduleName:="modMain", _
        procedureName:="StartApp", _
        entityType:="App", _
        entityID:="Startup", _
        correlationId:="VBA-STARTUP"
    On Error GoTo EH

    If Not m_Initialized Then InitApp

    ' --- Pristup: licenca + trial ("trial samo ako NIJE licenciran") ---
    ' Licencirana masina propusta; nelicencirana dobija trial (ako je ukljucen)
    ' ili pada na license gate. Opt-in: LICENSE_ENABLED / TRIAL_ENABLED.
    ' Detalji: modLicense.AccessGateOrQuit.
    If Not AccessGateOrQuit() Then Exit Sub

    Application.Visible = False

    frmSplash.Show             ' <-- splash pre main forme

    Call BackupFileOnStart
    Call PurgeOldBackups
    Call PurgeOldJournals
    Call PurgeOldLogs
    Call LogAppStart

    ' SEF recovery ostaje non-blocking za startup.
    ' Sama procedura RecoverAllStuckSEFSendingInvoices sada šalje monitoring.
    On Error Resume Next
    Call RecoverAllStuckSEFSendingInvoices
    On Error GoTo EH

    Dim journalWarning As String
    journalWarning = CheckJournalForRecovery()

    If journalWarning <> "" Then
        On Error Resume Next
        Monitor_Event _
            eventType:="JOURNAL_RECOVERY_WARN", _
            severity:="WARN", _
            message:=journalWarning, _
            userId:="Operator", _
            moduleName:="modMain", _
            procedureName:="StartApp", _
            entityType:="Journal", _
            entityID:="Recovery", _
            correlationId:="JOURNAL-STARTUP-RECOVERY"
        On Error GoTo EH

        MsgBox "UPOZORENJE - Moguc gubitak podataka!" & vbCrLf & vbCrLf & _
               journalWarning & vbCrLf & vbCrLf & _
               "Proverite Journal folder i reimportujte ako je potrebno.", _
               vbExclamation, APP_NAME
    End If

    On Error Resume Next
    Monitor_Event _
        eventType:="VBA_STARTAPP_SUCCESS", _
        severity:="INFO", _
        message:="StartApp completed successfully", _
        userId:="Operator", _
        moduleName:="modMain", _
        procedureName:="StartApp", _
        entityType:="App", _
        entityID:="Startup", _
        correlationId:="VBA-STARTUP"
    On Error GoTo 0

    ' NEW: Auto-sync scheduler.
    ' Ako SYNC_AUTO_INTERVAL_MIN nije postavljen ili je 0,
    ' StartScheduledSync samo loguje OFF i izlazi.
    On Error Resume Next
    StartScheduledSync
    If Err.Number <> 0 Then
        LogErr "modMain.StartApp.StartScheduledSync"
        Err.Clear
    End If
    On Error GoTo 0

    ' frmSplash sam sebe Unloaduje i pokrece frmOtkupAPP
    Exit Sub

EH:
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    Monitor_Error _
        moduleName:="modMain", _
        procedureName:="StartApp", _
        entityType:="App", _
        entityID:="Startup", _
        correlationId:="VBA-STARTUP", _
        errorNumber:=errNo, _
        errorDescription:=errDesc, _
        errorSource:=errSrc

    LogErr "modMain.StartApp"

    On Error GoTo 0
    Err.Raise errNo, errSrc, errDesc
End Sub

Public Sub InitApp()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrHandler
    
    ValidateAllTables
    m_Initialized = True
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Greska pri inicijalizaciji: " & Err.description, vbCritical, APP_NAME
    Resume CleanUp
End Sub

Public Sub ShutdownApp()
    On Error GoTo EH

    If mIsShuttingDown Then Exit Sub
    mIsShuttingDown = True

    ' Otkazi pending scheduled-sync OnTime pre gasenja.
    On Error Resume Next
    StopScheduledSync
    If Err.Number <> 0 Then
        LogErr "modMain.ShutdownApp.StopScheduledSync"
        Err.Clear
    End If
    On Error GoTo EH

    Application.Visible = True

    UnloadAllUserForms

    LogAppShutdown

    Exit Sub

EH:
    Application.Visible = True
    LogErr "modMain.ShutdownApp"
End Sub

Private Sub UnloadAllUserForms()
    On Error Resume Next

    Do While VBA.UserForms.count > 0
        Unload VBA.UserForms(0)
    Loop

    On Error GoTo 0
End Sub

Public Sub OpenExcel()
    Application.Visible = True
End Sub

Public Sub CloseExcel()
    Application.Visible = False
End Sub

Public Sub SaveApp()
    Application.ScreenUpdating = False
    ThisWorkbook.Save
    Application.ScreenUpdating = True
End Sub

Private Sub ValidateAllTables()
    Dim tblNames As Variant
    
    tblNames = Array(TBL_KOOPERANTI, TBL_STANICE, TBL_VOZACI, _
                     TBL_KUPCI, TBL_KULTURE, TBL_OTKUP, _
                     TBL_OTPREMNICA, TBL_ZBIRNA, TBL_PRIJEMNICA, _
                     TBL_FAKTURE, TBL_FAKTURA_STAVKE, _
                     TBL_NOVAC, TBL_AMBALAZA, TBL_CONFIG)
    
    Dim i As Long
    Dim missing As String
    
    For i = LBound(tblNames) To UBound(tblNames)
        If GetTable(CStr(tblNames(i))) Is Nothing Then
            missing = missing & CStr(tblNames(i)) & vbCrLf
        End If
    Next i
    
    If missing <> "" Then
        MsgBox "Sledece tabele ne postoje:" & vbCrLf & vbCrLf & missing & _
               vbCrLf & "Pokrenite Setup.", vbExclamation, APP_NAME
    End If
End Sub

