' ============================================================
' Banka Drive Pull - DODATAK za postojeci modBankaImport (sveska)
'
' KAKO: VBE -> otvori modBankaImport -> idi na KRAJ modula (posle
'       poslednjeg End Sub/Function) -> nalepi SVE ispod ove linije.
'       (Mora da bude UNUTAR modBankaImport jer zove Private GetBankaInboxPath.)
'
' Zatim: Debug -> Compile VBAProject (mora cisto).
'
' KONFIGURACIJA (tblLocalConfig, preko Podesavanja ili rucno):
'   BANKA_DRIVE_SOURCE_PATH          = folder koji Google Drive Desktop sinhronizuje (obavezno)
'   BANKA_DRIVE_DOWNLOADED_PATH      = gde se original PDF premesta posle (opc.; default <source>\..\Downloaded)
'   BANKA_DRIVE_MAX_FILES            = max po prolazu (opc., default 50)
'   BANKA_DRIVE_MIN_FILE_AGE_SECONDS = preskoci fajlove mladje od N sek (opc., default 15; da se ne hvata polu-sinhronizovan)
'
' OKIDAC: Alt+F8 -> ImportBankaInbox_WithDrivePull   (povuce sa Drive-a pa uvuce u bazu)
'         ili samo PullBankPdfsFromDriveProduction   (samo povuce u inbox)
'         -> po zelji zakaci na dugme u frmBankaImport.
' ============================================================

Public Sub ImportBankaInbox_WithDrivePull()
    Const SRC As String = "ImportBankaInbox_WithDrivePull"

    On Error GoTo EH

    If BankaDrivePullConfigured() Then
        PullBankPdfsFromDriveProduction
    End If

    ImportBankaInbox_TX
    Exit Sub

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Sub

Private Function BankaDrivePullConfigured() As Boolean
    BankaDrivePullConfigured = (Len(Trim$(GetLocalConfigValue("BANKA_DRIVE_SOURCE_PATH", ""))) > 0)
End Function

Public Function PullBankPdfsFromDriveProduction() As Long
    Const SRC As String = "PullBankPdfsFromDriveProduction"

    On Error GoTo EH

    Dim driveSourcePath As String
    Dim driveDownloadedPath As String
    Dim localInboxPath As String
    Dim maxFiles As Long
    Dim minAgeSeconds As Long
    Dim files As Collection
    Dim item As Variant
    Dim pulledCount As Long

    driveSourcePath = BankaNormalizeFolderPath(GetLocalConfigValue("BANKA_DRIVE_SOURCE_PATH", ""))
    driveDownloadedPath = BankaNormalizeFolderPath(GetLocalConfigValue("BANKA_DRIVE_DOWNLOADED_PATH", ""))
    localInboxPath = BankaNormalizeFolderPath(GetBankaInboxPath())

    maxFiles = CLng(val(GetLocalConfigValue("BANKA_DRIVE_MAX_FILES", "50")))
    If maxFiles <= 0 Then maxFiles = 50

    minAgeSeconds = CLng(val(GetLocalConfigValue("BANKA_DRIVE_MIN_FILE_AGE_SECONDS", "15")))
    If minAgeSeconds < 0 Then minAgeSeconds = 15

    If Len(driveSourcePath) = 0 Then Exit Function

    If Len(driveDownloadedPath) = 0 Then
        driveDownloadedPath = BankaParentFolderPath(driveSourcePath) & "\Downloaded"
    End If

    If Dir$(driveSourcePath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 9501, SRC, _
            "Drive source folder ne postoji ili nije dostupan: " & driveSourcePath
    End If

    If StrComp(driveSourcePath, localInboxPath, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 9502, SRC, _
            "Drive source i lokalni inbox ne smeju biti isti folder."
    End If

    BankaEnsureFolderExistsRecursive localInboxPath
    BankaEnsureFolderExistsRecursive driveDownloadedPath

    Set files = BankaCollectPdfFiles(driveSourcePath)

    For Each item In files
        If pulledCount >= maxFiles Then Exit For

        If BankaIsFileReadyForPull(CStr(item), minAgeSeconds) Then
            BankaPullOnePdfFromDrive CStr(item), localInboxPath, driveDownloadedPath
            pulledCount = pulledCount + 1
        Else
            Debug.Print SRC & ": skip not-ready file: " & CStr(item)
        End If
    Next item

    Debug.Print SRC & ": completed. Pulled=" & CStr(pulledCount)
    PullBankPdfsFromDriveProduction = pulledCount
    Exit Function

EH:
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

Private Sub BankaPullOnePdfFromDrive(ByVal sourcePdfPath As String, _
                                     ByVal localInboxPath As String, _
                                     ByVal driveDownloadedPath As String)
    Const SRC As String = "BankaPullOnePdfFromDrive"

    On Error GoTo EH

    Dim fileName As String
    Dim localFinalPath As String
    Dim localTempPath As String
    Dim driveDownloadedTargetPath As String
    Dim sourceSize As Long
    Dim copiedSize As Long
    Dim movedOk As Boolean

    If Dir$(sourcePdfPath) = "" Then
        Err.Raise vbObjectError + 9510, SRC, "PDF ne postoji: " & sourcePdfPath
    End If

    fileName = BankaSafeFileName(BankaFileNameFromPath(sourcePdfPath))

    If LCase$(Right$(fileName, 4)) <> ".pdf" Then
        Err.Raise vbObjectError + 9511, SRC, "Fajl nije PDF: " & fileName
    End If

    sourceSize = FileLen(sourcePdfPath)
    If sourceSize <= 0 Then
        Err.Raise vbObjectError + 9512, SRC, "PDF je prazan: " & sourcePdfPath
    End If

    localFinalPath = GetUniqueTargetPath(localInboxPath & "\" & fileName)
    localTempPath = localFinalPath & ".part"
    driveDownloadedTargetPath = GetUniqueTargetPath(driveDownloadedPath & "\" & fileName)

    If Dir$(localTempPath) <> "" Then Kill localTempPath

    FileCopy sourcePdfPath, localTempPath

    If Dir$(localTempPath) = "" Then
        Err.Raise vbObjectError + 9513, SRC, "Temp lokalni PDF nije kreiran."
    End If

    copiedSize = FileLen(localTempPath)
    If copiedSize <> sourceSize Then
        Err.Raise vbObjectError + 9514, SRC, _
            "Kopirani PDF nema istu velicinu. Source=" & CStr(sourceSize) & _
            " Local=" & CStr(copiedSize) & " File=" & fileName
    End If

    Name localTempPath As localFinalPath

    If Dir$(localFinalPath) = "" Then
        Err.Raise vbObjectError + 9515, SRC, "Final lokalni PDF nije kreiran."
    End If

    MoveFileSafe sourcePdfPath, driveDownloadedTargetPath
    movedOk = True

    Debug.Print SRC & ": pulled. Source=" & sourcePdfPath & _
                " Local=" & localFinalPath & _
                " DriveDownloaded=" & driveDownloadedTargetPath

    Exit Sub

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next

    If Not movedOk Then
        If Len(localTempPath) > 0 And Dir$(localTempPath) <> "" Then Kill localTempPath
        If Len(localFinalPath) > 0 And Dir$(localFinalPath) <> "" Then Kill localFinalPath
    End If

    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Sub

Private Function BankaCollectPdfFiles(ByVal folderPath As String) As Collection
    Dim result As Collection
    Dim f As String

    Set result = New Collection
    folderPath = BankaNormalizeFolderPath(folderPath)

    f = Dir$(folderPath & "\*.pdf")
    Do While Len(f) > 0
        result.Add folderPath & "\" & f
        f = Dir$
    Loop

    Set BankaCollectPdfFiles = result
End Function

Private Function BankaIsFileReadyForPull(ByVal filePath As String, _
                                         ByVal minAgeSeconds As Long) As Boolean
    On Error GoTo NotReady

    Dim ageSeconds As Long
    Dim s1 As Long
    Dim s2 As Long

    If Dir$(filePath) = "" Then GoTo NotReady

    ageSeconds = DateDiff("s", FileDateTime(filePath), Now)
    If ageSeconds < minAgeSeconds Then GoTo NotReady

    s1 = FileLen(filePath)
    If s1 <= 0 Then GoTo NotReady

    DoEvents

    s2 = FileLen(filePath)
    If s1 <> s2 Then GoTo NotReady

    BankaIsFileReadyForPull = True
    Exit Function

NotReady:
    BankaIsFileReadyForPull = False
End Function

Private Function BankaNormalizeFolderPath(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)

    Do While Len(folderPath) > 1 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    BankaNormalizeFolderPath = folderPath
End Function

Private Function BankaParentFolderPath(ByVal folderPath As String) As String
    Dim p As Long

    folderPath = BankaNormalizeFolderPath(folderPath)
    p = InStrRev(folderPath, "\")

    If p <= 0 Then
        BankaParentFolderPath = folderPath
    Else
        BankaParentFolderPath = Left$(folderPath, p - 1)
    End If
End Function

Private Function BankaFileNameFromPath(ByVal filePath As String) As String
    Dim p As Long

    p = InStrRev(filePath, "\")
    If p > 0 Then
        BankaFileNameFromPath = Mid$(filePath, p + 1)
    Else
        BankaFileNameFromPath = filePath
    End If
End Function

Private Function BankaSafeFileName(ByVal fileName As String) As String
    Dim badChars As Variant
    Dim i As Long

    fileName = Trim$(fileName)
    If Len(fileName) = 0 Then fileName = "bank.pdf"

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(badChars) To UBound(badChars)
        fileName = Replace(fileName, CStr(badChars(i)), "_")
    Next i

    If Len(fileName) > 180 Then fileName = Left$(fileName, 180)

    BankaSafeFileName = fileName
End Function

Private Sub BankaEnsureFolderExistsRecursive(ByVal folderPath As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long

    folderPath = BankaNormalizeFolderPath(folderPath)

    If Len(folderPath) = 0 Then Exit Sub
    If Dir$(folderPath, vbDirectory) <> "" Then Exit Sub

    parts = Split(folderPath, "\")
    currentPath = parts(0)

    If Right$(currentPath, 1) = ":" Then currentPath = currentPath & "\"

    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Right$(currentPath, 1) <> "\" Then currentPath = currentPath & "\"
            currentPath = currentPath & parts(i)

            If Dir$(currentPath, vbDirectory) = "" Then MkDir currentPath
        End If
    Next i
End Sub

