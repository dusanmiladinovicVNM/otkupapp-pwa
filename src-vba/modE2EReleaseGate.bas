
Option Explicit

' ============================================================
' modE2EReleaseGate
'
' v6.10 release-level E2E gate.
'
' This does not replace individual smoke suites.
' It orchestrates the VBA-side release checks and documents
' the GAS checks that must be run from Apps Script editor.
'
' Entry:
'   RunE2EReleaseGate_v610
' ============================================================

Private m_Total As Long
Private m_Pass As Long
Private m_Warn As Long
Private m_Fail As Long

Public Sub RunE2EReleaseGate_v610()
    On Error GoTo EH

    BeginE2EGate

    E2E_RunVbaSuite "RunGoogleSyncSmokeSuite", _
                    "Google transport/auth/sheets smoke"

    E2E_RunVbaSuite "RunMasterSyncSmokeSuite", _
                    "MasterSync fixture import/write-back/idempotency smoke"

    E2E_RunVbaSuite "RunNovacSmokeSuite", _
                    "Finance / Novac smoke suite"

    E2E_RunVbaSuite "RunFakturaSmokeSuite", _
                    "Faktura smoke suite"

    E2E_RunVbaSuite "RunBusinessFlowProSuite", _
                    "Business flow professional regression suite"

    E2E_RunVbaSuite "RunProductionHealthCheck", _
                    "Production health check"

    E2E_ManualGate "runGasRouteHealthCheck", _
                   "Must be run in Apps Script editor. Expected: PASS 29 handlers."

    E2E_ManualGate "runGasSmokeSuite", _
                   "Must be run in Apps Script editor. Expected: PASS 10/10."

    E2E_Warn "ProductionHealthCheck known data issue", _
             "Current workbook may fail on pre-existing Faktura/Prijemnica test references. Sync section must be PASS."

    EndE2EGate
    Exit Sub

EH:
    E2E_Fail "RunE2EReleaseGate_v610 fatal", _
             "Err.Number=" & CStr(Err.Number) & _
             " Source=" & Err.SOURCE & _
             " Description=" & Err.description

    EndE2EGate
End Sub

' ============================================================
' RUNNERS
' ============================================================

Private Sub E2E_RunVbaSuite(ByVal procName As String, ByVal description As String)
    On Error GoTo EH

    Debug.Print "RUN | " & procName & " | " & description
    Application.Run procName

    E2E_Pass procName, "Completed. Verify suite summary output/log for PASS counters."
    Exit Sub

EH:
    E2E_Fail procName, Err.description
End Sub

Private Sub E2E_ManualGate(ByVal gateName As String, ByVal instruction As String)
    E2E_Warn gateName, instruction
End Sub

' ============================================================
' LOGGING
' ============================================================

Private Sub BeginE2EGate()
    m_Total = 0
    m_Pass = 0
    m_Warn = 0
    m_Fail = 0

    Debug.Print String$(70, "=")
    Debug.Print "AGRIX v6.10 E2E RELEASE GATE START " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(70, "=")

    LogInfo "E2EReleaseGate", "START v6.10"
End Sub

Private Sub EndE2EGate()
    Dim summary As String

    summary = "TOTAL=" & m_Total & _
              " PASS=" & m_Pass & _
              " WARN=" & m_Warn & _
              " FAIL=" & m_Fail

    Debug.Print String$(70, "-")
    Debug.Print "AGRIX v6.10 E2E RELEASE GATE END"
    Debug.Print summary
    Debug.Print String$(70, "-")

    LogInfo "E2EReleaseGate", summary

    If m_Fail > 0 Then
        MsgBox "E2E Release Gate finished with FAILURES." & vbCrLf & summary, _
               vbCritical, APP_NAME
    ElseIf m_Warn > 0 Then
        MsgBox "E2E Release Gate finished with WARNINGS." & vbCrLf & _
               summary & vbCrLf & vbCrLf & _
               "Manual GAS checks and known ProductionHealthCheck notes must be reviewed.", _
               vbExclamation, APP_NAME
    Else
        MsgBox "E2E Release Gate PASS." & vbCrLf & summary, _
               vbInformation, APP_NAME
    End If
End Sub

Private Sub E2E_Pass(ByVal name As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Pass = m_Pass + 1

    Debug.Print "PASS | " & name & " | " & details
    LogInfo "E2EReleaseGate", "PASS | " & name & " | " & details
End Sub

Private Sub E2E_Warn(ByVal name As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Warn = m_Warn + 1

    Debug.Print "WARN | " & name & " | " & details
    LogWarn "E2EReleaseGate", "WARN | " & name & " | " & details
End Sub

Private Sub E2E_Fail(ByVal name As String, ByVal details As String)
    m_Total = m_Total + 1
    m_Fail = m_Fail + 1

    Debug.Print "FAIL | " & name & " | " & details
    LogError "E2EReleaseGate", "FAIL | " & name & " | " & details
End Sub

