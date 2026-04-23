Attribute VB_Name = "modLogError"
Option Explicit

' ============================================================
' modLogError – Zentralisiertes Error-Logging
'
' Schreibt Fehler in eine Textdatei für Remote-Support.
' Wenn der Kunde anruft, kann er die Log-Datei schicken.
'
' Log-Pfad: ThisWorkbook.Path & "\Log\"
' Dateiname: OtkupApp_2026-03-18.log (eine pro Tag)
' Rotation:  Dateien älter als 30 Tage werden bei App-Start gelöscht
'
' WICHTIG: Log-Write darf NIEMALS die App blockieren.
' ============================================================

Private Const LOG_FOLDER As String = "Log"
Private Const LOG_MAX_DAYS As Long = 30
Private Const LOG_PREFIX As String = "OtkupApp_"

' ============================================================
' Log-Levels
' ============================================================
Public Const LOG_ERROR As String = "ERROR"
Public Const LOG_WARN As String = "WARN"
Public Const LOG_INFO As String = "INFO"

' ============================================================
' PUBLIC - Hauptfunktion
' ============================================================

Public Sub LogError(ByVal Source As String, ByVal message As String, _
                    Optional ByVal errNumber As Long = 0, _
                    Optional ByVal level As String = "ERROR", _
                    Optional ByVal details As String = "")
    ' Schreibt eine Log-Zeile in die Tagesdatei
    '
    ' Aufruf aus EH-Blöcken:
    '   LogError "SaveOtkup", Err.Description, Err.Number
    '
    ' Oder als Info:
    '   LogError "StartApp", "App gestartet", level:=LOG_INFO
    
    Dim logPath As String
    Dim fileName As String
    Dim filePath As String
    Dim line As String
    Dim ff As Integer
    
    On Error Resume Next
    
    logPath = ThisWorkbook.Path & "\" & LOG_FOLDER
    
    ' Ordner erstellen
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    
    ' Dateiname: OtkupApp_2026-03-18.log
    fileName = LOG_PREFIX & Format$(Date, "yyyy-mm-dd") & ".log"
    filePath = logPath & "\" & fileName
    
    ' Log-Zeile bauen
    ' Format: 2026-03-18 14:35:22 | ERROR | SaveOtkup | 5 | Kooperant mora biti izabran! | details
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
           PadRight(level, 5) & " | " & _
           PadRight(Source, 30) & " | "
    
    If errNumber <> 0 Then
        line = line & CStr(errNumber) & " | "
    Else
        line = line & "- | "
    End If
    
    line = line & message
    
    If Len(Trim$(details)) > 0 Then
        line = line & " | " & details
    End If
    
    ' Append
    ff = FreeFile
    Open filePath For Append As #ff
    Print #ff, line
    Close #ff
    
    ' Debug-Fenster auch
    Debug.Print line
    
    On Error GoTo 0
End Sub

' ============================================================
' PUBLIC - Kurzformen
' ============================================================

Public Sub LogErr(ByVal Source As String, Optional ByVal details As String = "")
    ' Kurzform: loggt aktuellen Err direkt
    ' Aufruf: LogErr "SaveOtkup"
    ' Muss im EH-Block aufgerufen werden wo Err noch aktiv ist
    
    If Err.Number <> 0 Then
        LogError Source, Err.Description, Err.Number, LOG_ERROR, details
    End If
End Sub

Public Sub LogWarn(ByVal Source As String, ByVal message As String, _
                   Optional ByVal details As String = "")
    LogError Source, message, 0, LOG_WARN, details
End Sub

Public Sub LogInfo(ByVal Source As String, ByVal message As String, _
                   Optional ByVal details As String = "")
    LogError Source, message, 0, LOG_INFO, details
End Sub

' ============================================================
' PUBLIC - App Lifecycle Logging
' ============================================================

Public Sub LogAppStart()
    LogInfo "APP", "=== OtkupApp " & APP_VERSION & " gestartet ==="
    LogInfo "APP", "File: " & ThisWorkbook.Name
    LogInfo "APP", "User: " & Environ$("Username")
End Sub

Public Sub LogAppShutdown()
    LogInfo "APP", "=== OtkupApp beendet ==="
End Sub

' ============================================================
' PUBLIC - Rotation (aufgerufen aus modMain.StartApp)
' ============================================================

Public Sub PurgeOldLogs()
    ' Löscht Log-Dateien die älter als LOG_MAX_DAYS sind
    
    Dim logPath As String
    Dim fileName As String
    Dim filePath As String
    Dim datePart As String
    Dim fileDate As Date
    
    On Error Resume Next
    
    logPath = ThisWorkbook.Path & "\" & LOG_FOLDER
    
    If Dir(logPath, vbDirectory) = "" Then Exit Sub
    
    fileName = Dir(logPath & "\*.log")
    
    Do While fileName <> ""
        ' Datum aus Dateiname: OtkupApp_2026-03-18.log
        If Len(fileName) >= Len(LOG_PREFIX) + 14 Then
            datePart = Mid$(fileName, Len(LOG_PREFIX) + 1, 10)
            
            If IsDate(datePart) Then
                fileDate = CDate(datePart)
                
                If DateDiff("d", fileDate, Date) > LOG_MAX_DAYS Then
                    filePath = logPath & "\" & fileName
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

Private Function PadRight(ByVal s As String, ByVal totalWidth As Long) As String
    If Len(s) >= totalWidth Then
        PadRight = Left$(s, totalWidth)
    Else
        PadRight = s & Space$(totalWidth - Len(s))
    End If
End Function


