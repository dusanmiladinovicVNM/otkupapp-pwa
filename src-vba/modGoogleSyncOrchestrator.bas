Attribute VB_Name = "modGoogleSyncOrchestrator"
Option Explicit

' ============================================================
' modGoogleSyncOrchestrator
'
' Jedno dugme za ceo PWA / Google sync ciklus:
'
'   1. INBOUND / MASTER:
'      - Geo pull Google -> tblParcele
'      - Import OTK/PWA -> tblOtkup
'      - Auto-create Otpremnice
'      - Import VOZ/Zbirne -> tblZbirna
'
'   2. OUTBOUND / STAMMDATEN:
'      - Export Stammdaten
'      - Export Kartice
'      - Export MgmtReports
'
'   3. PWA LOCK:
'      - SyncControl tab u Stammdaten
'      - MASTER_SYNC_LOCK = YES/NO
' ============================================================

Private Const ORCH_MODULE As String = "modGoogleSyncOrchestrator"
Private Const REQUIRE_GEO_PULL_BEFORE_OUTBOUND As Boolean = True

Private Const SYNC_AUTO_INTERVAL_MIN_KEY As String = "SYNC_AUTO_INTERVAL_MIN"
Private Const SYNC_AUTO_MIN_ALLOWED_MIN As Long = 15

Private Const SYNC_CONTROL_TAB As String = "SyncControl"
Private Const MASTER_SYNC_LOCK_KEY As String = "MASTER_SYNC_LOCK"

Private Const MASTER_SYNC_LOCK_TTL_MIN As Long = 10
Private Const EVENT_PWA_UNLOCK_FAIL As String = "PWA_MASTER_SYNC_UNLOCK_FAIL"
Private Const EVENT_PWA_UNLOCK_FAIL_SEVERITY As String = "ERROR"

Private mNextScheduledRun As Date
Private mIsScheduled As Boolean
Private mSyncInProgress As Boolean

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub SyncPWAFullCycle()
    Call SyncPWAFullCycle_Core(True)
End Sub

Public Function SyncPWAFullCycle_ForButton() As Boolean
    SyncPWAFullCycle_ForButton = SyncPWAFullCycle_Core(False)
End Function

' ============================================================
' CORE
' ============================================================

Private Function SyncPWAFullCycle_Core(ByVal showMessages As Boolean) As Boolean
    Dim summary As String

    Dim okGeo As Boolean
    Dim okOtkup As Boolean
    Dim okOtpremnice As Boolean
    Dim okZbirne As Boolean
    Dim okStammdaten As Boolean
    Dim okKartice As Boolean
    Dim okMgmt As Boolean

    Dim createdOtp As Long
    Dim createdZbr As Long
    Dim errNum As Long
    Dim errDesc As String
    Dim pwaLockAcquired As Boolean
    
    Dim pwaUnlockOk As Boolean
    Dim unlockErrDesc As String
    Dim finalMessageAlreadyShown As Boolean

    On Error GoTo EH

    SyncPWAFullCycle_Core = False
    
    pwaUnlockOk = True
    finalMessageAlreadyShown = False

    If mSyncInProgress Then
        LogWarn ORCH_MODULE, "SyncPWAFullCycle_Core re-entered. Aborting."

        If showMessages Then
            MsgBox "Sync je vec u toku. Sacekaj da se zavrsi pa pokusaj ponovo.", _
                   vbExclamation, APP_NAME
        End If

        Exit Function
    End If

    mSyncInProgress = True

    LogInfo ORCH_MODULE, "Full PWA / Google sync cycle started."

    summary = "PWA / Google sync ciklus:" & vbCrLf & vbCrLf

    SyncProgress "Proveravam Google konfiguraciju..."

    If Not IsGoogleAuthConfigured() Then
        LogError ORCH_MODULE, "Google OAuth2 nije konfigurisan."

        summary = summary & "GRESKA - Google OAuth2 nije konfigurisan." & vbCrLf

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then
            MsgBox summary & vbCrLf & _
                   "Pokreni RunGoogleAuthSetup iz modGoogleAuth.", _
                   vbCritical, APP_NAME
        End If

        GoTo CleanExit
    End If

    If Len(Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))) = 0 Then
        LogError ORCH_MODULE, "GOOGLE_PWA_FOLDER_ID nije postavljen."

        summary = summary & "GRESKA - GOOGLE_PWA_FOLDER_ID nije postavljen." & vbCrLf

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then
            MsgBox summary, vbCritical, APP_NAME
        End If

        GoTo CleanExit
    End If

    SyncProgress "Zakljucavam PWA upis dok traje master sync..."

    If Not SetPWAMasterSyncLock(True, "Master sync je u toku. Sacekajte zavrsetak.") Then
        LogError ORCH_MODULE, "PWA lock could not be acquired. Cycle aborted."

        summary = summary & "GRESKA - PWA lock nije postavljen. Sync prekinut." & vbCrLf

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then
            MsgBox summary, vbCritical, APP_NAME
        End If

        GoTo CleanExit
    End If

    pwaLockAcquired = True
    SyncProgress "PWA upis je privremeno zakljucan."

    summary = summary & "[MASTER / INBOUND]" & vbCrLf

    ' 1. Geo pull je hard gate da Stammdaten ne pregazi poligone.
    SyncProgress "Povlacim geo/poligone iz Google Sheets..."
    okGeo = ImportParcelGeoFromGoogleToMaster()
    AppendStep summary, okGeo, "Geo/Polygon pull Google -> tblParcele"

    If REQUIRE_GEO_PULL_BEFORE_OUTBOUND And Not okGeo Then
        LogError ORCH_MODULE, _
                 "Geo pull failed. Outbound Stammdaten sync aborted to avoid overwriting PolygonGeoJSON."

        summary = summary & vbCrLf & _
                  "Outbound Stammdaten sync je prekinut da ne pregazi poligone." & vbCrLf

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then MsgBox summary, vbExclamation, APP_NAME
        GoTo CleanExit
    End If

    ' 2. OTK import
    SyncProgress "Uvozim OTK zapise iz PWA..."
    okOtkup = ImportOtkupFromPWA_Core(False)
    AppendStep summary, okOtkup, "Import OTK/PWA -> tblOtkup"

    If Not okOtkup Then
        LogError ORCH_MODULE, "ImportOtkupFromPWA_Core failed. Cycle aborted before outbound sync."

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then MsgBox summary, vbExclamation, APP_NAME
        GoTo CleanExit
    End If

    ' 2b. MALINA: VozacID := StanicaID (pre auto-otpremnice; pali okidac)
    If IsMalinaMode() Then
        SyncProgress "Malina: popunjavam VozacID iz StanicaID..."

        On Error Resume Next
        Err.Clear
        Call StampVozacFromStanicaForMalina_TX
        errNum = Err.Number
        errDesc = Err.description
        On Error GoTo EH

        If errNum <> 0 Then
            AppendStep summary, False, "Malina VozacID:=StanicaID | Error=" & errDesc
            LogError ORCH_MODULE, "StampVozacFromStanicaForMalina_TX failed: " & errDesc

            Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                                 okStammdaten, okKartice, okMgmt, False

            If showMessages Then MsgBox summary, vbExclamation, APP_NAME
            GoTo CleanExit
        End If

        AppendStep summary, True, "Malina: VozacID:=StanicaID"
    End If

    ' 3. Auto-create Otpremnice
    SyncProgress "Kreiram / povezujem otpremnice..."

    On Error Resume Next
    Err.Clear
    createdOtp = AutoCreateOtpremniceFromPWA_TX()
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo EH

    okOtpremnice = (errNum = 0)

    If okOtpremnice Then
        AppendStep summary, True, _
            "Auto-create Otpremnice from PWA Otkup (" & CStr(createdOtp) & " kreirano)"
    Else
        AppendStep summary, False, _
            "Auto-create Otpremnice from PWA Otkup | Error=" & errDesc

        LogError ORCH_MODULE, "AutoCreateOtpremniceFromPWA_TX failed: " & errDesc

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then MsgBox summary, vbExclamation, APP_NAME
        GoTo CleanExit
    End If

    ' 3b. MALINA: auto-zbirna iz otpremnice (1:1; u malini zamenjuje korak 4)
    If IsMalinaMode() Then
        SyncProgress "Malina: kreiram zbirne iz otpremnica..."

        On Error Resume Next
        Err.Clear
        createdZbr = AutoCreateZbirnaFromOtpremnice_TX()
        errNum = Err.Number
        errDesc = Err.description
        On Error GoTo EH

        If errNum <> 0 Then
            AppendStep summary, False, "Malina auto-zbirna | Error=" & errDesc
            LogError ORCH_MODULE, "AutoCreateZbirnaFromOtpremnice_TX failed: " & errDesc

            Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                                 okStammdaten, okKartice, okMgmt, False

            If showMessages Then MsgBox summary, vbExclamation, APP_NAME
            GoTo CleanExit
        End If

        AppendStep summary, True, "Malina auto-zbirna (" & CStr(createdZbr) & " kreirano)"
    End If

    ' 4. VOZ/Zbirne import
    SyncProgress "Uvozim zbirne vozaca..."
    okZbirne = ImportZbirneFromPWA_Core(False)
    AppendStep summary, okZbirne, "Import VOZ/Zbirne -> tblZbirna"

    If Not okZbirne Then
        LogError ORCH_MODULE, "ImportZbirneFromPWA_Core failed. Cycle aborted before outbound sync."

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then MsgBox summary, vbExclamation, APP_NAME
        GoTo CleanExit
    End If

    summary = summary & vbCrLf & "[STAMMDATEN / OUTBOUND]" & vbCrLf

    ' 5. Outbound exports
    SyncProgress "Eksportujem Stammdaten..."
    okStammdaten = SyncStammdatenToGoogle_Core(False)
    AppendStep summary, okStammdaten, "Export Stammdaten tbl* -> Google"

    SyncProgress "Eksportujem kartice..."
    okKartice = ExportKarticeToGoogle_Core(False)
    AppendStep summary, okKartice, "Export Kartice -> Google"

    SyncProgress "Eksportujem MgmtReports..."
    okMgmt = ExportMgmtReports_Core(False)
    AppendStep summary, okMgmt, "Export MgmtReports -> Google"

    SyncPWAFullCycle_Core = _
        okGeo And _
        okOtkup And _
        okOtpremnice And _
        okZbirne And _
        okStammdaten And _
        okKartice And _
        okMgmt

    Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                         okStammdaten, okKartice, okMgmt, _
                         SyncPWAFullCycle_Core

    LogInfo ORCH_MODULE, _
        "Full PWA / Google sync cycle completed. " & _
        "Geo=" & CStr(okGeo) & _
        "; Otkup=" & CStr(okOtkup) & _
        "; Otpremnice=" & CStr(okOtpremnice) & _
        "; Zbirne=" & CStr(okZbirne) & _
        "; Stammdaten=" & CStr(okStammdaten) & _
        "; Kartice=" & CStr(okKartice) & _
        "; MgmtReports=" & CStr(okMgmt)

    If SyncPWAFullCycle_Core Then
        SyncProgress "Full sync uspesno zavrsen."
    Else
        SyncProgress "Full sync zavrsen sa greskom / partial statusom."
    End If

    If showMessages Then
        If Not pwaLockAcquired Then
            If SyncPWAFullCycle_Core Then
                MsgBox summary & vbCrLf & "Status: OK", vbInformation, APP_NAME
            Else
                MsgBox summary & vbCrLf & "Status: GRESKA / PARTIAL", vbExclamation, APP_NAME
            End If
            finalMessageAlreadyShown = True
        End If
    End If

CleanExit:
    If pwaLockAcquired Then
        SyncProgress "Otkljucavam PWA upis..."

        If Not SetPWAMasterSyncLock(False, "Master sync zavrsen.") Then
            pwaUnlockOk = False
            unlockErrDesc = _
                "PWA master sync lock nije skinut iz VBA ciklusa. " & _
                "PWA moze ostati privremeno zakljucan do isteka GAS/PWA lock TTL-a " & _
                "(~" & CStr(MASTER_SYNC_LOCK_TTL_MIN) & " min). " & _
                "Ako se stanje ne oporavi automatski, proveri Stammdaten/SyncControl i MASTER_SYNC_LOCK."

            LogError ORCH_MODULE, unlockErrDesc
            SyncProgress "GRESKA: PWA lock nije skinut iz VBA ciklusa; ocekuje se automatski TTL oporavak."

            ' Not permanent lockout, but the operator must not see a green cycle.
            SyncPWAFullCycle_Core = False

            Monitor_PWAMasterUnlockFail unlockErrDesc

            Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                                 okStammdaten, okKartice, okMgmt, False, _
                                 EVENT_PWA_UNLOCK_FAIL_SEVERITY, _
                                 "unlock-failed-ttl-" & CStr(MASTER_SYNC_LOCK_TTL_MIN) & "min"

            If showMessages Then
                MsgBox summary & vbCrLf & _
                       "Status: GRESKA / PARTIAL" & vbCrLf & vbCrLf & _
                       "Sync koraci su mozda zavrseni, ali VBA nije uspeo da skine PWA lock." & vbCrLf & _
                       "PWA se po GAS/PWA pravilu automatski oporavlja posle ~" & _
                       CStr(MASTER_SYNC_LOCK_TTL_MIN) & " minuta." & vbCrLf & _
                       "Ako se ne oporavi, proveri Stammdaten / SyncControl.", _
                       vbExclamation, APP_NAME
                finalMessageAlreadyShown = True
            End If
        Else
            SyncProgress "PWA upis je ponovo dozvoljen."

            If showMessages And Not finalMessageAlreadyShown Then
                If SyncPWAFullCycle_Core Then
                    MsgBox summary & vbCrLf & "Status: OK", vbInformation, APP_NAME
                Else
                    MsgBox summary & vbCrLf & "Status: GRESKA / PARTIAL", vbExclamation, APP_NAME
                End If
                finalMessageAlreadyShown = True
            End If
        End If
    End If

    mSyncInProgress = False
    Exit Function

EH:
    Dim fatalDesc As String
    fatalDesc = Err.description

    LogErr ORCH_MODULE & ".SyncPWAFullCycle_Core"

    Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                         okStammdaten, okKartice, okMgmt, False

    If showMessages Then
        MsgBox "Greska u PWA / Google sync ciklusu: " & fatalDesc, _
               vbCritical, APP_NAME
    End If

    SyncPWAFullCycle_Core = False
    Resume CleanExit
End Function

Private Sub AppendStep(ByRef summary As String, _
                       ByVal ok As Boolean, _
                       ByVal stepName As String)
    If ok Then
        summary = summary & "OK - " & stepName & vbCrLf
        LogInfo ORCH_MODULE, "OK - " & stepName
    Else
        summary = summary & "GRESKA - " & stepName & vbCrLf
        LogError ORCH_MODULE, "FAIL - " & stepName
    End If
End Sub

Private Sub SyncProgress(ByVal message As String)
    On Error Resume Next

    Dim uf As Object

    LogInfo ORCH_MODULE, "PROGRESS - " & CStr(message)

    For Each uf In VBA.UserForms
        If TypeName(uf) = "frmOtkupAPP" Then
            uf.AppendPWASyncLog CStr(message)
            Exit For
        End If
    Next uf

    DoEvents
End Sub

Private Sub Monitor_PWAFullCycle(ByVal okGeo As Boolean, _
                                 ByVal okOtkup As Boolean, _
                                 ByVal okOtpremnice As Boolean, _
                                 ByVal okZbirne As Boolean, _
                                 ByVal okStammdaten As Boolean, _
                                 ByVal okKartice As Boolean, _
                                 ByVal okMgmt As Boolean, _
                                 ByVal cycleOk As Boolean, _
                                 Optional ByVal severityOverride As String = "", _
                                 Optional ByVal reason As String = "")
    On Error Resume Next

    Dim corrId As String
    Dim msg As String
    Dim sev As String

    corrId = "PWA-FULL-CYCLE-" & Format$(Now, "yyyymmddhhnnss")

    msg = "Geo=" & CStr(okGeo) & _
          "; Otkup=" & CStr(okOtkup) & _
          "; Otpremnice=" & CStr(okOtpremnice) & _
          "; Zbirne=" & CStr(okZbirne) & _
          "; Stammdaten=" & CStr(okStammdaten) & _
          "; Kartice=" & CStr(okKartice) & _
          "; MgmtReports=" & CStr(okMgmt)

    If Len(Trim$(reason)) > 0 Then
        msg = msg & "; Reason=" & reason
    End If

    If Len(Trim$(severityOverride)) > 0 Then
        sev = UCase$(Trim$(severityOverride))
    Else
        sev = IIf(cycleOk, "INFO", "CRITICAL")
    End If

    Monitor_Event _
        eventType:=IIf(cycleOk, "PWA_FULL_CYCLE_SUCCESS", "PWA_FULL_CYCLE_FAIL"), _
        severity:=sev, _
        message:=msg, _
        userId:="Operator", _
        moduleName:=ORCH_MODULE, _
        procedureName:="SyncPWAFullCycle_Core", _
        entityType:="MasterData", _
        entityID:="PWA-FULL-CYCLE", _
        correlationId:=corrId

    If Not cycleOk Then
        Monitor_Error _
            moduleName:=ORCH_MODULE, _
            procedureName:="SyncPWAFullCycle_Core", _
            entityType:="MasterData", _
            entityID:="PWA-FULL-CYCLE", _
            correlationId:=corrId, _
            errorNumber:=0, _
            errorDescription:="PWA full sync cycle failed or completed degraded. " & msg, _
            errorSource:=ORCH_MODULE
    End If

    On Error GoTo 0
End Sub

Private Sub Monitor_PWAMasterUnlockFail(ByVal errorDescription As String)
    On Error Resume Next

    Dim corrId As String
    corrId = "PWA-UNLOCK-FAIL-" & Format$(Now, "yyyymmddhhnnss")

    Monitor_Error _
        moduleName:=ORCH_MODULE, _
        procedureName:="SyncPWAFullCycle_Core", _
        entityType:="MasterData", _
        entityID:="PWA-MASTER-SYNC-LOCK", _
        correlationId:=corrId, _
        errorNumber:=0, _
        errorDescription:=errorDescription, _
        errorSource:=ORCH_MODULE

    Monitor_Event _
        eventType:=EVENT_PWA_UNLOCK_FAIL, _
        severity:=EVENT_PWA_UNLOCK_FAIL_SEVERITY, _
        message:=errorDescription, _
        userId:="Operator", _
        moduleName:=ORCH_MODULE, _
        procedureName:="SyncPWAFullCycle_Core", _
        entityType:="MasterData", _
        entityID:="PWA-MASTER-SYNC-LOCK", _
        correlationId:=corrId

    On Error GoTo 0
End Sub

' ============================================================
' PWA MASTER SYNC LOCK
' ============================================================

Private Function GetOrCreateStammdatenSheetIDForSyncLock() As String
    Const SRC As String = "GetOrCreateStammdatenSheetIDForSyncLock"

    Dim folderID As String
    Dim sheetID As String

    On Error GoTo EH

    folderID = GetConfigValue("GOOGLE_PWA_FOLDER_ID")
    If Len(Trim$(folderID)) = 0 Then Exit Function

    sheetID = GetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID")

    If Len(Trim$(sheetID)) = 0 Then
        sheetID = GetSpreadsheetID("Stammdaten", folderID)
    End If

    If Len(Trim$(sheetID)) = 0 Then
        sheetID = CreateSpreadsheet("Stammdaten", folderID)
    End If

    If Len(Trim$(sheetID)) > 0 Then
        Call SetConfigValue("GOOGLE_STAMMDATEN_SHEET_ID", sheetID)

        On Error Resume Next
        Call AddSheetTab(sheetID, SYNC_CONTROL_TAB)
        On Error GoTo EH
    End If

    GetOrCreateStammdatenSheetIDForSyncLock = sheetID
    Exit Function

EH:
    LogErr SRC
    GetOrCreateStammdatenSheetIDForSyncLock = ""
End Function

Private Function SetPWAMasterSyncLock(ByVal locked As Boolean, _
                                      Optional ByVal message As String = "") As Boolean
    Const SRC As String = "SetPWAMasterSyncLock"

    Dim sheetID As String
    Dim rows(1 To 5, 1 To 2) As Variant

    On Error GoTo EH

    sheetID = GetOrCreateStammdatenSheetIDForSyncLock()

    If Len(Trim$(sheetID)) = 0 Then
        LogError SRC, "Stammdaten sheetID nije dostupan."
        SetPWAMasterSyncLock = False
        Exit Function
    End If

    rows(1, 1) = "Parameter"
    rows(1, 2) = "Vrednost"

    rows(2, 1) = MASTER_SYNC_LOCK_KEY
    rows(2, 2) = IIf(locked, "YES", "NO")

    rows(3, 1) = "MASTER_SYNC_UPDATED_AT"
    rows(3, 2) = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")

    rows(4, 1) = "MASTER_SYNC_MESSAGE"
    rows(4, 2) = CStr(message)

    rows(5, 1) = "MASTER_SYNC_OWNER"
    rows(5, 2) = "VBA"

    SetPWAMasterSyncLock = WriteSheetData(sheetID, SYNC_CONTROL_TAB, rows)

    If SetPWAMasterSyncLock Then
        LogInfo SRC, "PWA master sync lock=" & IIf(locked, "YES", "NO")
    Else
        LogError SRC, "WriteSheetData failed for SyncControl."
    End If

    Exit Function

EH:
    LogErr SRC
    SetPWAMasterSyncLock = False
End Function

' ============================================================
' SCHEDULED AUTO-SYNC
' ============================================================

Public Sub StartScheduledSync()
    Dim intervalMin As Long

    intervalMin = GetScheduledSyncIntervalMin()

    If intervalMin <= 0 Then
        LogInfo ORCH_MODULE, _
            "Scheduled auto-sync disabled (" & SYNC_AUTO_INTERVAL_MIN_KEY & " <= 0)."
        Exit Sub
    End If

    If intervalMin < SYNC_AUTO_MIN_ALLOWED_MIN Then
        LogWarn ORCH_MODULE, _
            SYNC_AUTO_INTERVAL_MIN_KEY & "=" & CStr(intervalMin) & _
            " is below min allowed " & CStr(SYNC_AUTO_MIN_ALLOWED_MIN) & _
            ". Auto-sync disabled."
        Exit Sub
    End If

    ScheduleNextSyncTick intervalMin

    LogInfo ORCH_MODULE, _
        "Scheduled auto-sync started. Interval=" & CStr(intervalMin) & " min. " & _
        "Next run at " & Format$(mNextScheduledRun, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub StopScheduledSync()
    If Not mIsScheduled Then Exit Sub

    On Error Resume Next
    Application.OnTime EarliestTime:=mNextScheduledRun, _
                       Procedure:="ScheduledSyncTick", _
                       Schedule:=False
    On Error GoTo 0

    mIsScheduled = False
    LogInfo ORCH_MODULE, "Scheduled auto-sync stopped."
End Sub

Public Sub ScheduledSyncTick()
    On Error GoTo EH

    If mSyncInProgress Then
        LogWarn ORCH_MODULE, _
            "ScheduledSyncTick fired but sync already in progress. Skipping."
        GoTo Reschedule
    End If

    LogInfo ORCH_MODULE, "ScheduledSyncTick firing automatic SyncPWAFullCycle_Core."

    Call SyncPWAFullCycle_Core(False)

Reschedule:
    Dim intervalMin As Long
    intervalMin = GetScheduledSyncIntervalMin()

    If intervalMin >= SYNC_AUTO_MIN_ALLOWED_MIN Then
        ScheduleNextSyncTick intervalMin
        LogInfo ORCH_MODULE, _
            "Next scheduled tick at " & Format$(mNextScheduledRun, "yyyy-mm-dd hh:nn:ss")
    Else
        mIsScheduled = False
        LogInfo ORCH_MODULE, _
            "Scheduled auto-sync stopping because interval is disabled or below threshold."
    End If

    Exit Sub

EH:
    LogErr ORCH_MODULE & ".ScheduledSyncTick"
    Err.Clear

    On Error Resume Next

    Dim fallbackInterval As Long
    fallbackInterval = GetScheduledSyncIntervalMin()

    If fallbackInterval >= SYNC_AUTO_MIN_ALLOWED_MIN Then
        ScheduleNextSyncTick fallbackInterval
    Else
        mIsScheduled = False
    End If
End Sub

Private Sub ScheduleNextSyncTick(ByVal intervalMin As Long)
    On Error Resume Next

    If mIsScheduled Then
        Application.OnTime EarliestTime:=mNextScheduledRun, _
                           Procedure:="ScheduledSyncTick", _
                           Schedule:=False
    End If

    On Error GoTo 0

    mNextScheduledRun = Now + (CDbl(intervalMin) / 1440#)

    On Error Resume Next
    Application.OnTime EarliestTime:=mNextScheduledRun, _
                       Procedure:="ScheduledSyncTick"
    If Err.Number <> 0 Then
        LogErr ORCH_MODULE & ".ScheduleNextSyncTick"
        mIsScheduled = False
        Err.Clear
    Else
        mIsScheduled = True
    End If
    On Error GoTo 0
End Sub

Private Function GetScheduledSyncIntervalMin() As Long
    Dim raw As String
    Dim n As Long

    On Error Resume Next
    raw = Trim$(CStr(GetConfigValue(SYNC_AUTO_INTERVAL_MIN_KEY)))
    On Error GoTo 0

    If Len(raw) = 0 Then
        GetScheduledSyncIntervalMin = 0
        Exit Function
    End If

    On Error Resume Next
    n = CLng(raw)

    If Err.Number <> 0 Then
        GetScheduledSyncIntervalMin = 0
        Err.Clear
        Exit Function
    End If

    On Error GoTo 0

    GetScheduledSyncIntervalMin = n
End Function

' ============================================================
' TEST
' ============================================================

Public Sub Test_SyncPWAFullCycle()
    Call SyncPWAFullCycle
End Sub

Public Sub Test_StartScheduledSync()
    Call StartScheduledSync
End Sub

Public Sub Test_StopScheduledSync()
    Call StopScheduledSync
End Sub

