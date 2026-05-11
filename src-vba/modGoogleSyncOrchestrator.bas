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
' UI dugme poziva:
'   SyncPWAFullCycle
' ============================================================

Private Const ORCH_MODULE As String = "modGoogleSyncOrchestrator"
Private Const REQUIRE_GEO_PULL_BEFORE_OUTBOUND As Boolean = True

Public Sub SyncPWAFullCycle()
    Dim ok As Boolean

    ok = SyncPWAFullCycle_Core(True)

    If ok Then
        MsgBox "PWA / Google sync ciklus je uspešno završen.", _
               vbInformation, APP_NAME
    Else
        MsgBox "PWA / Google sync ciklus je završen sa greškom." & vbCrLf & _
               "Proveri log pre sledeceg Stammdaten exporta.", _
               vbExclamation, APP_NAME
    End If
End Sub

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
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo EH

    SyncPWAFullCycle_Core = False

    LogInfo ORCH_MODULE, "Full PWA / Google sync cycle started."

    If Not IsGoogleAuthConfigured() Then
        LogError ORCH_MODULE, "Google OAuth2 nije konfigurisan."
        If showMessages Then
            MsgBox "Google OAuth2 nije konfigurisan!" & vbCrLf & _
                   "Pokreni RunGoogleAuthSetup iz modGoogleAuth.", _
                   vbCritical, APP_NAME
        End If
        Exit Function
    End If

    If Len(Trim$(GetConfigValue("GOOGLE_PWA_FOLDER_ID"))) = 0 Then
        LogError ORCH_MODULE, "GOOGLE_PWA_FOLDER_ID nije postavljen."
        If showMessages Then
            MsgBox "GOOGLE_PWA_FOLDER_ID nije postavljen u tblConfig.", _
                   vbCritical, APP_NAME
        End If
        Exit Function
    End If

    summary = "PWA / Google sync ciklus:" & vbCrLf & vbCrLf
    summary = summary & "[MASTER / INBOUND]" & vbCrLf

    ' 1. Geo pull je hard gate da Stammdaten ne pregazi poligone.
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
        Exit Function
    End If

    ' 2. OTK import
    okOtkup = ImportOtkupFromPWA_Core(False)
    AppendStep summary, okOtkup, "Import OTK/PWA -> tblOtkup"

    If Not okOtkup Then
        LogError ORCH_MODULE, "ImportOtkupFromPWA_Core failed. Cycle aborted before outbound sync."

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then MsgBox summary, vbExclamation, APP_NAME
        Exit Function
    End If

    ' 3. Auto-create Otpremnice
    On Error Resume Next
    Err.Clear
    createdOtp = AutoCreateOtpremniceFromPWA_TX()
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo EH

    okOtpremnice = (errNum = 0)

    If okOtpremnice Then
        AppendStep summary, True, "Auto-create Otpremnice from PWA Otkup (" & CStr(createdOtp) & " kreirano)"
    Else
        AppendStep summary, False, "Auto-create Otpremnice from PWA Otkup | Error=" & errDesc
        LogError ORCH_MODULE, "AutoCreateOtpremniceFromPWA failed: " & errDesc

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then MsgBox summary, vbExclamation, APP_NAME
        Exit Function
    End If

    ' 4. VOZ/Zbirne import
    okZbirne = ImportZbirneFromPWA_Core(False)
    AppendStep summary, okZbirne, "Import VOZ/Zbirne -> tblZbirna"

    If Not okZbirne Then
        LogError ORCH_MODULE, "ImportZbirneFromPWA_Core failed. Cycle aborted before outbound sync."

        Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                             okStammdaten, okKartice, okMgmt, False

        If showMessages Then MsgBox summary, vbExclamation, APP_NAME
        Exit Function
    End If

    summary = summary & vbCrLf & "[STAMMDATEN / OUTBOUND]" & vbCrLf

    ' 5. Outbound exports
    okStammdaten = SyncStammdatenToGoogle_Core(False)
    AppendStep summary, okStammdaten, "Export Stammdaten tbl* -> Google"

    okKartice = ExportKarticeToGoogle_Core(False)
    AppendStep summary, okKartice, "Export Kartice -> Google"

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

    If showMessages Then
        If SyncPWAFullCycle_Core Then
            MsgBox summary & vbCrLf & "Status: OK", vbInformation, APP_NAME
        Else
            MsgBox summary & vbCrLf & "Status: GRESKA / PARTIAL", vbExclamation, APP_NAME
        End If
    End If

    Exit Function

EH:
    LogErr ORCH_MODULE & ".SyncPWAFullCycle_Core"

    Monitor_PWAFullCycle okGeo, okOtkup, okOtpremnice, okZbirne, _
                         okStammdaten, okKartice, okMgmt, False

    If showMessages Then
        MsgBox "Greška u PWA / Google sync ciklusu: " & Err.description, _
               vbCritical, APP_NAME
    End If

    SyncPWAFullCycle_Core = False
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

Private Sub Monitor_PWAFullCycle(ByVal okGeo As Boolean, _
                                 ByVal okOtkup As Boolean, _
                                 ByVal okOtpremnice As Boolean, _
                                 ByVal okZbirne As Boolean, _
                                 ByVal okStammdaten As Boolean, _
                                 ByVal okKartice As Boolean, _
                                 ByVal okMgmt As Boolean, _
                                 ByVal cycleOk As Boolean)
    On Error Resume Next

    Dim corrId As String
    Dim msg As String

    corrId = "PWA-FULL-CYCLE-" & Format$(Now, "yyyymmddhhnnss")

    msg = "Geo=" & CStr(okGeo) & _
          "; Otkup=" & CStr(okOtkup) & _
          "; Otpremnice=" & CStr(okOtpremnice) & _
          "; Zbirne=" & CStr(okZbirne) & _
          "; Stammdaten=" & CStr(okStammdaten) & _
          "; Kartice=" & CStr(okKartice) & _
          "; MgmtReports=" & CStr(okMgmt)

    Monitor_Event _
        eventType:=IIf(cycleOk, "PWA_FULL_CYCLE_SUCCESS", "PWA_FULL_CYCLE_FAIL"), _
        severity:=IIf(cycleOk, "INFO", "CRITICAL"), _
        message:=msg, _
        userId:="Operator", _
        moduleName:=ORCH_MODULE, _
        procedureName:="SyncPWAFullCycle_Core", _
        entityType:="MasterData", _
        entityId:="PWA-FULL-CYCLE", _
        correlationId:=corrId

    If Not cycleOk Then
        Monitor_Error _
            moduleName:=ORCH_MODULE, _
            procedureName:="SyncPWAFullCycle_Core", _
            entityType:="MasterData", _
            entityId:="PWA-FULL-CYCLE", _
            correlationId:=corrId, _
            errorNumber:=0, _
            errorDescription:="PWA full sync cycle failed or partially failed. " & msg, _
            errorSource:=ORCH_MODULE
    End If
End Sub
Public Sub Test_SyncPWAFullCycle()
    Call SyncPWAFullCycle
End Sub

