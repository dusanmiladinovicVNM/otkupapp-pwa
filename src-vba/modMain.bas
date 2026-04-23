Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
' modMain v2.1 – ValidateAllTables aktualisiert
' ============================================================

Private m_Initialized As Boolean

Public Sub StartApp()
    If Not m_Initialized Then InitApp
    Application.Visible = False
    frmOtkupAPP.Show vbModeless
    
    Call BackupFileOnStart
    Call PurgeOldBackups
    
    Call PurgeOldJournals
    
    Call PurgeOldLogs
    Call LogAppStart
    

    On Error Resume Next
        Call RecoverAllStuckSEFSendingInvoices
    On Error GoTo 0
    
    Dim journalWarning As String
    journalWarning = CheckJournalForRecovery()
    If journalWarning <> "" Then
        MsgBox "UPOZORENJE - Moguc gubitak podataka!" & vbCrLf & vbCrLf & _
               journalWarning & vbCrLf & vbCrLf & _
               "Proverite Journal folder i reimportujte ako je potrebno.", _
               vbExclamation, APP_NAME
    End If
End Sub

Public Sub InitApp()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrHandler
    ValidateAllTables
    m_Initialized = True
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "Greska pri inicijalizaciji: " & Err.Description, vbCritical, APP_NAME
    Resume Cleanup
End Sub

Public Sub ShutdownApp()
    Application.Visible = True
    Unload frmMain
    Call LogAppShutdown
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


