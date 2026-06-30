Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
' modMain v2.1 - ValidateAllTables aktualisiert
' ============================================================

Private m_Initialized As Boolean
Private mIsShuttingDown As Boolean

' KPI sidebar je "dirty" kad se doda red u TBL_OTKUP/OTPREMNICA/PRIJEMNICA
' (postavlja modDataAccess.AppendRow) ili posle PWA importa. frmOtkupAPP.UserForm_Activate
' osvezava KPI samo kad je dirty -- ne pri svakom povratku na dashboard.
Public gKpiDirty As Boolean

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

    ' --- Licenca (per-uredjaj / node-locked) ---
    ' Blokira pokretanje ako licenca nije vazeca za OVAJ racunar.
    ' Opt-in: radi samo ako je LICENSE_ENABLED = YES u tblSEFConfig
    ' (inace fail-open, ne dira postojece instalacije). Detalji: modLicense.
    If Not AccessGateOrQuit() Then Exit Sub

    ' --- Min-version gate (flota) ---
    ' Server (GAS action "checkVersion") javlja minimalnu dozvoljenu verziju;
    ' zastarela verzija dobija upozorenje, a uz enforce=YES i blok pokretanja.
    ' Opt-in na MONITORING_ENDPOINT+SECRET; fail-open offline. Vidi modUpdateGate.
    If Not UpdateGateOrQuit() Then Exit Sub

    ' --- Self-update (povuci novu verziju koda iz AgriX_Release) ---
    ' Opt-in na REL_FOLDER_ID; fail-soft. Na "Da" -> import na praznom
    ' stack-u pa Exit (ne pokrecemo ostatak; restart aktivira nov kod).
    ' Vidi modSelfUpdate.
    If CheckForUpdateOnOpen() Then
        Application.OnTime Now, "RunSelfUpdate"
        Exit Sub
    End If

    ' --- Per-user prijava (opt-in: AUTH_ENABLED u tblSEFConfig) ---
    ' Dok AUTH_ENABLED != YES -> sve radi kao pre (bez prijave).
    ' Neuspela prijava -> Excel vidljiv + zakazano gasenje (mirror license gate;
    ' zatvaranje se ne radi unutar Workbook_Open lanca, vec na sledeci tick).
    If modAuth.AuthEnabled() Then
        If Not modAuth.Login() Then
            Application.Visible = True
            Application.OnTime Now + TimeSerial(0, 0, 1), "QuitAfterFailedLogin"
            Exit Sub
        End If
    End If

    Application.Visible = False

    frmSplash.Show             ' <-- splash pre main forme

    Call BackupFileOnStart
    Call PurgeOldBackups
    Call PurgeOldJournals
    Call PurgeOldLogs
    Call LogAppStart

    ' SEF recovery ostaje non-blocking za startup.
    ' Sama procedura RecoverAllStuckSEFSendingInvoices sada salje monitoring.
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

    On Error Resume Next
    EnsurePoruke
    If Err.Number <> 0 Then
        LogErr "modMain.InitApp.EnsurePoruke"
        Err.Clear
    End If
    On Error GoTo ErrHandler

    ValidateAllTables
    m_Initialized = True
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Gre" & ChrW(353) & "ka pri inicijalizaciji: " & Err.description, vbCritical, APP_NAME
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
    Application.ScreenUpdating = False

    UnloadAllUserForms

    Application.ScreenUpdating = True
    LogAppShutdown

    Exit Sub

EH:
    Application.ScreenUpdating = True
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
                     TBL_NOVAC, TBL_AMBALAZA, TBL_CONFIG, TBL_PORUKE)
    
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

