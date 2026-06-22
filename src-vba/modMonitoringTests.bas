Attribute VB_Name = "modMonitoringTests"
Option Explicit

Public Sub TestMonitoring_All()
    Debug.Print String(60, "-")
    Debug.Print "MONITORING TEST SUITE START"
    Debug.Print "Time:", Now
    Debug.Print String(60, "-")

    TestMonitoring_Config
    TestMonitoring_HTTP
    TestMonitoring_ErrorEvent
    TestMonitoring_SEFUnknown
    TestMonitoring_BackupSuccess
    TestMonitoring_BackupFail

    Debug.Print String(60, "-")
    Debug.Print "MONITORING TEST SUITE END"
    Debug.Print String(60, "-")
End Sub

Public Sub TestMonitoring_Config()
    Debug.Print vbCrLf & "[1] CONFIG TEST"
    Debug.Print Monitoring_DiagnoseConfig()
    
    Debug.Print "PASS ako vidiš:"
    Debug.Print "- Endpoint length > 0"
    Debug.Print "- Endpoint ends with /exec = True"
    Debug.Print "- Secret length > 0"
    Debug.Print "- AppVersion = tvoja verzija"
End Sub

Public Sub TestMonitoring_HTTP()
    Debug.Print vbCrLf & "[2] HTTP INGEST TEST"
    
    Dim ok As Boolean
    ok = Monitor_Test()
    
    Debug.Print "Monitor_Test result:", ok
    
    If ok Then
        Debug.Print "PASS: GAS je vratio success=true"
        Debug.Print "Proveri sheet tabs: Events, Health"
    Else
        Debug.Print "FAIL: pogledaj HTTP Status / HTTP Response iz debug output-a"
    End If
End Sub

Public Sub TestMonitoring_ErrorEvent()
    Debug.Print vbCrLf & "[3] VBA ERROR EVENT TEST"

    On Error GoTo EH

    Err.Raise 9101, "TestMonitoring_ErrorEvent", "Namerno testirana VBA greška bez sensitive podataka"

    Exit Sub

EH:
    Monitor_Error _
        moduleName:="modMonitoringTests", _
        procedureName:="TestMonitoring_ErrorEvent", _
        entityType:="Test", _
        entityID:="TEST-ERROR-001", _
        correlationId:="TEST-ERROR-001", _
        errorNumber:=Err.Number, _
        errorDescription:=Err.description, _
        errorSource:=Err.SOURCE

    Debug.Print "PASS ako se pojavio red u Events + Errors"
End Sub

Public Sub TestMonitoring_SEFUnknown()
    Debug.Print vbCrLf & "[4] SEF UNKNOWN / CRITICAL TEST"

    Monitor_SEF _
        eventType:="SEF_UNKNOWN", _
        severity:="CRITICAL", _
        invoiceLocalId:="FAK-TEST-001", _
        businessInvoiceNo:="TEST-001/2026", _
        sefStatus:="UNKNOWN", _
        localStatus:="SEF_UNKNOWN", _
        sefRequestId:="REQ-TEST-001", _
        sefInvoiceId:="", _
        attemptCount:=1, _
        lastHttpCode:="timeout", _
        lastError:="Test timeout after SEF send; status unknown", _
        nextAction:="CHECK_SEF_PORTAL", _
        needsManualReview:=True

    Debug.Print "PASS ako se pojavio red u:"
    Debug.Print "- Events"
    Debug.Print "- Errors"
    Debug.Print "- SEFStatus"
    Debug.Print "- Alerts"
    Debug.Print "- Health"
    Debug.Print "- AuditCritical"
End Sub

Public Sub TestMonitoring_BackupSuccess()
    Debug.Print vbCrLf & "[5] BACKUP SUCCESS TEST"

    Monitor_Backup _
        backupType:="TEST_BACKUP", _
        status:="SUCCESS", _
        backupFileId:="test-backup-id", _
        backupLocation:="TEST_LOCATION_REDACTION_SAFE", _
        rowsCount:=123, _
        checksum:="test-checksum", _
        durationMs:=250, _
        errorMessage:=""

    Debug.Print "PASS ako se pojavio red u Events + Backups"
End Sub

Public Sub TestMonitoring_BackupFail()
    Debug.Print vbCrLf & "[6] BACKUP FAIL TEST"

    Monitor_Backup _
        backupType:="TEST_BACKUP_FAIL", _
        status:="FAILED", _
        backupFileId:="test-backup-failed-id", _
        backupLocation:="TEST_LOCATION_REDACTION_SAFE", _
        rowsCount:=0, _
        checksum:="", _
        durationMs:=1200, _
        errorMessage:="Namerno testiran backup fail"

    Debug.Print "PASS ako se pojavio red u:"
    Debug.Print "- Events"
    Debug.Print "- Errors"
    Debug.Print "- Backups"
    Debug.Print "- Alerts"
    Debug.Print "- Health"
End Sub

