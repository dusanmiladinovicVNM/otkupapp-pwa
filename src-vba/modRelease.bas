Attribute VB_Name = "modRelease"
Option Explicit

' ============================================================
' modRelease - BUILD-ONLY: objavi src-vba kod + version.json u
' AgriX_Release folder na Drive-u (kanal za self-update klijenata).
'
' Pokrece se RUCNO na build masini, posle:
'   ImportAllVBA -> Compile -> AssertBlankBuild,
' a PRE 'git checkout -- modBuildInfo.bas' (sub cita stamp-ovan BUILD_*).
'
' Klijenti imaju ovaj kod ali ga NIKAD ne pozovu (kao modBuildGuard /
' ExportAllVBA). Vidi docs/RELEASE_PROCEDURE.md korak 7b.
' ASCII-only modul (vidi CLAUDE.md, sekcija 4 - encoding).
' ============================================================

' Izvorni folder = isti src-vba put kao modVbaTools.ImportAllVBA. IZMENI po masini.
Private Const SRC_FOLDER As String = "C:\Users\Dusan\Documents\GitHub\otkupapp-pwa\src-vba\"

Public Sub PublishReleaseToDrive()
    Const SRC As String = "modRelease.PublishReleaseToDrive"
    On Error GoTo EH

    If Len(REL_FOLDER_ID) = 0 Then
        MsgBox "REL_FOLDER_ID nije postavljen u modConfig.bas.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(SRC_FOLDER) Then
        MsgBox "Izvorni folder ne postoji: " & SRC_FOLDER, vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim fil As Object, ext As String, uploaded As Long, failed As String

    For Each fil In fso.GetFolder(SRC_FOLDER).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        Select Case ext
            Case "bas", "cls", "frm", "frx", "doccls"
                If Len(DriveUploadFile(REL_FOLDER_ID, fil.path, fil.name)) > 0 Then
                    uploaded = uploaded + 1
                Else
                    failed = failed & "  " & fil.name & vbCrLf
                End If
        End Select
    Next fil

    ' version.json (manifest) - iz stamp-ovanog modBuildInfo + APP_VERSION
    Dim manifest As String, tmpPath As String, okMan As Boolean
    manifest = BuildManifestJson()
    tmpPath = fso.GetSpecialFolder(2) & "\version.json"      ' TemporaryFolder
    WriteReleaseTextFile tmpPath, manifest
    okMan = (Len(DriveUploadFile(REL_FOLDER_ID, tmpPath, "version.json")) > 0)

    MsgBox "Objavljeno u AgriX_Release:" & vbCrLf & _
           "  fajlova (kod): " & uploaded & vbCrLf & _
           "  version.json:  " & IIf(okMan, "OK", "GRESKA") & vbCrLf & vbCrLf & _
           IIf(Len(failed) > 0, "NIJE uspelo:" & vbCrLf & failed & vbCrLf, "Sve OK." & vbCrLf) & _
           "Manifest: " & manifest, _
           IIf(Len(failed) > 0 Or Not okMan, vbExclamation, vbInformation), APP_NAME
    Exit Sub
EH:
    LogErr SRC, Err.description
    MsgBox "Greska pri objavljivanju: " & Err.description, vbCritical, APP_NAME
End Sub

' app_version = SemVer komparator (isti kao modUpdateGate / VERSION_LATEST).
Private Function BuildManifestJson() As String
    BuildManifestJson = _
        "{""app_version"":""" & APP_VERSION & """," & _
        """build_version"":""" & BUILD_VERSION & """," & _
        """build_sha"":""" & BUILD_SHA & """," & _
        """build_date"":""" & BUILD_DATE & """}"
End Function

Private Sub WriteReleaseTextFile(ByVal path As String, ByVal content As String)
    Dim ff As Integer: ff = FreeFile
    Open path For Output As #ff
    Print #ff, content;
    Close #ff
End Sub
