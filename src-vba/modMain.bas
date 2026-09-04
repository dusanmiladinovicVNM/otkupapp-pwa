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

    ' --- Self-update (povuci novu verziju koda iz AgriX_Release) ---
    ' Startup watchdog (RecoverPendingSelfUpdate) se zove IZNUTRA CheckForUpdateOnOpen
    ' (isti modul). NE zovi ga odavde direktno: modSelfUpdate je u SKIP_MODULES (frozen),
    ' pa star klijent posle self-update-a ima NOV modMain + STAR modSelfUpdate; direktan
    ' (early-bound) poziv NOVOG simbola = COMPILE error ("Sub or Function not defined")
    ' koji obori ceo StartApp. modMain sme early-bind SAMO stabilne modSelfUpdate simbole
    ' (docs/SELF_UPDATE.md zamka #19).
    ' VAZNO: ide PRE min-version gate-a. Inace bi enforce=YES blokirao pokretanje
    ' bas onog klijenta kome je update najpotrebniji (zastarela verzija dobije
    ' zakazano gasenje pre nego sto stigne do ove provere). Sad: prvo ponudi
    ' self-update; ako ga korisnik odbije, min-version gate ispod moze da blokira.
    ' Opt-in na REL_FOLDER_ID; fail-soft. Na "Da" -> import na praznom stack-u pa
    ' Exit. OnTime workbook-qualified (dve otvorene kopije -> pravi workbook).
    If CheckForUpdateOnOpen() Then
        Application.OnTime Now, "'" & Replace$(ThisWorkbook.name, "'", "''") & "'!RunSelfUpdate"
        Exit Sub
    End If

    ' --- Min-version gate (flota) ---
    ' Server (GAS action "checkVersion") javlja minimalnu dozvoljenu verziju;
    ' zastarela verzija dobija upozorenje, a uz enforce=YES i blok pokretanja.
    ' Opt-in na MONITORING_ENDPOINT+SECRET; fail-open offline. Vidi modUpdateGate.
    If Not UpdateGateOrQuit() Then Exit Sub

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

    ' --- First-run setup gate (per-masina) ---
    ' Ako ovaj racunar jos nije prosao SetupNewPC (APP_SETUP_COMPLETED != "DA" u
    ' tblLocalConfig), ponudi podesavanje odmah -- pre skrivanja Excela i splash-a,
    ' dok je prozor jos vidljiv i interaktivan. Jednokratno: cim SetupNewPC prodje
    ' zeleno i upise "DA", ova kapija se vise ne javlja. Fail-soft (ne obara start).
    On Error Resume Next
    If UCase$(Trim$(GetLocalConfigValue("APP_SETUP_COMPLETED", ""))) <> "DA" Then
        If MsgBox(Poruka("SETUP_MSG_FIRSTRUN_PONUDA"), vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            SetupNewPC
        End If
    End If
    On Error GoTo EH

    ' --- Tockic misa nad listama (per-masina; Podesavanja -> "Interfejs / lokalno") ---
    ' MOUSEWHEEL_SCROLL: DA/NE, prazno = DA (ukljuceno). Brane u modMouseWheel
    ' (VBE-guard, lenjivi hook samo nad listom) vaze i kad je ukljuceno.
    On Error Resume Next
    If UCase$(Trim$(GetLocalConfigValue("MOUSEWHEEL_SCROLL", "DA"))) <> "NE" Then MouseWheel_On Else MouseWheel_Off
    On Error GoTo EH

    Application.Visible = False

    ' Splash je od v6-ui-213 FAZA ljuske, ne svoja forma: isti prozor
    ' (frmOtkupUI) drzi splash, prijavu, mini karticu i aplikaciju. Redosled je
    ' isti koji je frmSplash imao -- dve sekunde znaka, pa ulaz u ljusku --
    ' samo sto ulaz sada zove StartApp, a ne forma sama sebe.
    modUiFaze.FazaBoot 2
    modOtkupUI.ShowOtkupUI

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

    ' Idle pre-warm Storno cockpit keza (prvi warm ~60s posle pokretanja); posle
    ' svake TX se prezakazuje iz CommitTx. Fail-soft. Vidi modStornoWarm.
    On Error Resume Next
    ScheduleStornoWarm
    On Error GoTo 0

    ' Ljuska je vec na ekranu (splash faza + ShowOtkupUI iznad).
    ' Stari meni (frmOtkupAPP) je obrisan u koraku 7 -- ljuska je jedini ulaz.
    Exit Sub

EH:
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    errNo = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "modMain.StartApp"
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

    ' Schema self-heal: kolone dodate kroz self-update KODA nastanu automatski
    ' posle restarta (silent, idempotentno; isti obrazac kao EnsurePoruke).
    On Error Resume Next
    EnsureRuntimeSchema
    If Err.Number <> 0 Then
        LogErr "modMain.InitApp.EnsureRuntimeSchema"
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
    ' Otkazi pending storno-warm OnTime (inace moze da reotvori workbook posle close).
    StopStornoWarm
    On Error GoTo EH

    ' Skini mouse hook pre gasenja (higijena; inace ga skida i QueryClose formi).
    On Error Resume Next
    MouseWheel_Off
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

' Zatvaranje aplikacije sa sledeceg tick-a. Postoji zbog OnTime: zatvaranje
' sveske IZ event-a kontrole (clsFlatBtn sink je tada na steku) ili iz
' Workbook_Open lanca nije bezbedno -- isti razlog zbog kojeg licencna kapija
' i neuspela prijava zakazuju gasenje umesto da ga izvrse odmah.
' Snima, kao i legacy "Izlaz" (frmOtkupAPP.btnExit_Click).
Public Sub ZatvoriAplikaciju()
    On Error Resume Next
    ThisWorkbook.Close SaveChanges:=True
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

