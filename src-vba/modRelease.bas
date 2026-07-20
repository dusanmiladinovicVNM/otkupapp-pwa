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

    ' 1) Upload svih code fajlova. Pamti lokalna imena (za prune) + files JSON
    '    (ime+velicina, za manifest / buducu klijentsku verifikaciju).
    Dim fil As Object, ext As String, uploaded As Long, failed As String
    Dim localNames As Object: Set localNames = CreateObject("Scripting.Dictionary")
    localNames.CompareMode = 1                       ' TextCompare (case-insensitive)
    Dim filesJson As String

    For Each fil In fso.GetFolder(SRC_FOLDER).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        Select Case ext
            Case "bas", "cls", "frm", "frx", "doccls"
                If Not localNames.Exists(fil.name) Then localNames.Add fil.name, True
                If Len(DriveUploadFile(REL_FOLDER_ID, fil.path, fil.name)) > 0 Then
                    uploaded = uploaded + 1
                    If Len(filesJson) > 0 Then filesJson = filesJson & ","
                    filesJson = filesJson & "{""name"":""" & fil.name & """,""size"":" & fil.Size & "}"
                Else
                    failed = failed & "  " & fil.name & vbCrLf
                End If
        End Select
    Next fil

    ' 2) ATOMARNOST OBJAVE: ako je makar JEDAN upload pao, NE objavljuj version.json
    '    i NE prune-uj. Klijent broji fajlove po listingu -> polu-objavljen release
    '    (npr. nov modConfig ali stara/nedostajuca forma) izgledao bi kompletan.
    If Len(failed) > 0 Then
        MsgBox "Objava PREKINUTA - neki code fajlovi nisu uspeli:" & vbCrLf & failed & vbCrLf & _
               "version.json NIJE objavljen (release ostaje na prethodnoj verziji)." & vbCrLf & _
               "Uspesno: " & uploaded & ". Proverite Drive/auth pa ponovite.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    ' 3) Prune: ukloni (Trash) zastarele code fajlove iz AgriX_Release-a koji vise
    '    ne postoje u src-vba (npr. obrisani test moduli) - inace bi ih klijent
    '    ponovo preuzimao (DownloadReleaseFiles skida svaki podrzani fajl u folderu).
    Dim pruned As Long, pruneList As String
    Dim remote As Object: Set remote = DriveListFolder(REL_FOLDER_ID)
    If Not remote Is Nothing Then
        Dim k As Variant, kext As String
        For Each k In remote.Keys
            If StrComp(CStr(k), "version.json", vbTextCompare) <> 0 Then
                kext = LCase$(fso.GetExtensionName(CStr(k)))
                Select Case kext
                    Case "bas", "cls", "frm", "frx", "doccls"
                        If Not localNames.Exists(CStr(k)) Then
                            If DriveTrashFile(CStr(remote(k))) Then
                                pruned = pruned + 1
                                pruneList = pruneList & "  " & CStr(k) & vbCrLf
                            End If
                        End If
                End Select
            End If
        Next k
    End If

    ' 4) version.json (manifest) TEK sada - posle punog uploada + prune-a.
    Dim manifest As String, tmpPath As String, okMan As Boolean
    manifest = BuildManifestJson(filesJson)
    tmpPath = fso.GetSpecialFolder(2) & "\version.json"      ' TemporaryFolder
    WriteReleaseTextFile tmpPath, manifest
    okMan = (Len(DriveUploadFile(REL_FOLDER_ID, tmpPath, "version.json")) > 0)

    MsgBox "Objavljeno u AgriX_Release:" & vbCrLf & _
           "  fajlova (kod): " & uploaded & vbCrLf & _
           "  ocisceno (stale): " & pruned & vbCrLf & _
           IIf(Len(pruneList) > 0, pruneList, "") & _
           "  version.json:  " & IIf(okMan, "OK", "GRESKA") & vbCrLf & vbCrLf & _
           IIf(okMan, "Sve OK.", "PAZNJA: version.json nije objavljen!"), _
           IIf(okMan, vbInformation, vbExclamation), APP_NAME
    Exit Sub
EH:
    LogErr SRC, Err.description
    MsgBox "Greska pri objavljivanju: " & Err.description, vbCritical, APP_NAME
End Sub

' app_version = SemVer komparator (isti kao modUpdateGate / VERSION_LATEST).
' files = JSON niz {name,size} (za manifest; klijent moze da verifikuje kompletnost/
' velicinu). SHA-256 po fajlu je sledeci korak (pravi snapshot); za sada name+size.
Private Function BuildManifestJson(ByVal filesJson As String) As String
    BuildManifestJson = _
        "{""app_version"":""" & APP_VERSION & """," & _
        """build_version"":""" & BUILD_VERSION & """," & _
        """build_sha"":""" & BUILD_SHA & """," & _
        """build_date"":""" & BUILD_DATE & """," & _
        """files"":[" & filesJson & "]}"
End Function

Private Sub WriteReleaseTextFile(ByVal path As String, ByVal content As String)
    Dim ff As Integer: ff = FreeFile
    Open path For Output As #ff
    Print #ff, content;
    Close #ff
End Sub
