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

' Immutability granica: re-objava vec objavljene verzije (releases\<v> ima manifest.json)
' je ODBIJENA (klijent koji tu APP_VERSION vec ima NECE povuci izmenjene bajtove; snapshot
' bi bio prepisan). Za NAMERNU re-objavu - iskljucivo TEST kanal / hitna zamena pokvarenog
' release-a - privremeno postavi True (kao REL_FOLDER_ID=TEST; NE commit-uj True).
Private Const ALLOW_REPUBLISH As Boolean = False

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

    ' 0) F2: obezbedi versioned layout AgriX_Release\releases\<APP_VERSION>\ PRE
    '    uploada. Padne li ovo (Drive/auth), prekid pre ijedne izmene kanala.
    Dim relSub As String: relSub = DriveEnsureFolder(REL_FOLDER_ID, "releases")
    Dim vfid As String
    If Len(relSub) > 0 Then vfid = DriveEnsureFolder(relSub, APP_VERSION)
    If Len(relSub) = 0 Or Len(vfid) = 0 Then
        MsgBox "Ne mogu da pripremim versioned folder (releases\" & APP_VERSION & ")." & vbCrLf & _
               "Objava PREKINUTA - nista nije menjano. Proverite Drive/auth.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' Immutability: ako versioned folder VEC ima manifest.json, verzija je POTPUNO
    ' objavljena. Re-objava = HARD ODBIJENA (prepisala bi snapshot; klijent koji tu
    ' APP_VERSION vec ima NECE povuci izmenjene bajtove). Za namernu re-objavu (TEST /
    ' hitna zamena) -> ALLOW_REPUBLISH=True. (Retry NEUSPELE objave je bezbedan i BEZ
    ' toga: manifest.json se pise tek na kraju, pa ga posle pada jos nema.)
    If Len(DriveFindInFolder(vfid, "manifest.json")) > 0 Then
        If Not ALLOW_REPUBLISH Then
            MsgBox "Verzija " & APP_VERSION & " je VEC objavljena (releases\" & APP_VERSION & _
                   " ima manifest.json). Re-objava je ODBIJENA (immutable snapshot)." & vbCrLf & vbCrLf & _
                   "Klijent koji tu APP_VERSION vec ima NECE povuci izmenjene bajtove." & vbCrLf & _
                   "Za NOV release: bump APP_VERSION u modConfig." & vbCrLf & _
                   "Za NAMERNU re-objavu (TEST kanal / hitna zamena): ALLOW_REPUBLISH=True u modRelease.", _
                   vbCritical, APP_NAME
            Exit Sub
        End If
        LogErr SRC, "REPUBLISH verzije " & APP_VERSION & " (ALLOW_REPUBLISH=True) - snapshot se prepisuje"
    End If

    ' 1) Upload svih code fajlova u OBA kanala (flat REL_FOLDER_ID = stari klijenti;
    '    versioned vfid = sledljivost + F2 lanac). Pamti lokalna imena (za prune) +
    '    files JSON (ime+velicina+sha256, deljen za version.json i manifest.json).
    Dim fil As Object, ext As String, uploaded As Long, failed As String
    Dim localNames As Object: Set localNames = CreateObject("Scripting.Dictionary")
    localNames.CompareMode = 1                       ' TextCompare (case-insensitive)
    Dim filesJson As String

    For Each fil In fso.GetFolder(SRC_FOLDER).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        Select Case ext
            Case "bas", "cls", "frm", "frx", "doccls"
                If Not localNames.Exists(fil.name) Then localNames.Add fil.name, True
                ' SHA-256 sirovih bajtova (reuse modDrive.Sha256File). Build masina
                ' MORA da hesuje; prazno -> ne objavljuj manifest bez hesa (fail).
                Dim sh As String: sh = Sha256File(fil.path)
                If Len(sh) = 0 Then
                    failed = failed & "  " & fil.name & " (SHA-256 nedostupan na build masini)" & vbCrLf
                ElseIf Len(DriveUploadFile(REL_FOLDER_ID, fil.path, fil.name)) = 0 Then
                    failed = failed & "  " & fil.name & " (flat upload)" & vbCrLf
                ElseIf Len(DriveUploadFile(vfid, fil.path, fil.name)) = 0 Then
                    failed = failed & "  " & fil.name & " (versioned upload)" & vbCrLf
                Else
                    uploaded = uploaded + 1
                    If Len(filesJson) > 0 Then filesJson = filesJson & ","
                    filesJson = filesJson & "{""name"":""" & fil.name & """,""size"":" & fil.Size & _
                                ",""sha256"":""" & sh & """}"
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

    ' 4) Manifesti TEK sada - posle punog uploada (oba kanala) + prune-a. Identican
    '    sadrzaj: manifest.json (versioned, koren lanca) i version.json (flat, stari
    '    klijenti). manifest_sha256 (SHA-256 versioned manifest.json) ide u current.json.
    Dim manifest As String, tmpMan As String, tmpVer As String
    manifest = BuildManifestJson(filesJson)

    tmpMan = fso.GetSpecialFolder(2) & "\manifest.json"      ' TemporaryFolder
    WriteReleaseTextFile tmpMan, manifest
    Dim okVMan As Boolean: okVMan = (Len(DriveUploadFile(vfid, tmpMan, "manifest.json")) > 0)
    Dim manSha As String: manSha = Sha256File(tmpMan)

    tmpVer = fso.GetSpecialFolder(2) & "\version.json"
    WriteReleaseTextFile tmpVer, manifest
    Dim okMan As Boolean: okMan = (Len(DriveUploadFile(REL_FOLDER_ID, tmpVer, "version.json")) > 0)

    ' 5) current.json (F2 pokazivac) - POSLEDNJI upis (atomarni commit versioned-a).
    '    Pise se SAMO ako je versioned manifest.json uspesno objavljen (inace bi
    '    pokazivao na folder bez validnog manifesta -> klijent fail-closed).
    Dim okCur As Boolean, curJson As String, tmpCur As String
    ' current.json se pise SAMO ako je versioned manifest OK I manifest_sha256 validan
    ' 64-hex (Sha256File vrati "" na neuspeh). Bez validnog hesa nema versioned lanca
    ' poverenja -> ne objavljuj pokazivac (klijent tada koristi flat; okCur=False).
    If okVMan And Len(manSha) = 64 Then
        curJson = "{""app_version"":""" & APP_VERSION & """," & _
                  """release_folder_id"":""" & vfid & """," & _
                  """manifest_sha256"":""" & manSha & """," & _
                  """published"":""" & Format$(Now, "yyyy-mm-dd hh:nn") & """}"
        tmpCur = fso.GetSpecialFolder(2) & "\current.json"
        WriteReleaseTextFile tmpCur, curJson
        okCur = (Len(DriveUploadFile(REL_FOLDER_ID, tmpCur, "current.json")) > 0)
    End If

    ' 6) Retention: zadrzi najnovijih 10 versioned foldera, ostalo u Trash (meko).
    Dim relPruned As Long, relPruneList As String
    PruneOldReleases relSub, 10, relPruned, relPruneList

    MsgBox "Objavljeno u AgriX_Release:" & vbCrLf & _
           "  fajlova (kod, oba kanala): " & uploaded & vbCrLf & _
           "  versioned:     releases\" & APP_VERSION & vbCrLf & _
           "  ocisceno (flat stale): " & pruned & vbCrLf & _
           IIf(Len(pruneList) > 0, pruneList, "") & _
           "  ocisceno (stare verzije): " & relPruned & vbCrLf & _
           IIf(Len(relPruneList) > 0, relPruneList, "") & _
           "  manifest.json: " & IIf(okVMan, "OK", "GRESKA") & vbCrLf & _
           "  version.json:  " & IIf(okMan, "OK", "GRESKA") & vbCrLf & _
           "  current.json:  " & IIf(okCur, "OK", "GRESKA") & _
               IIf(Not okCur And okVMan And Len(manSha) <> 64, " (manifest_sha256 nije 64-hex -> pokazivac PRESKOCEN)", "") & vbCrLf & vbCrLf & _
           IIf(okVMan And okMan And okCur, "Sve OK.", "PAZNJA: neki manifest/pokazivac NIJE objavljen!"), _
           IIf(okVMan And okMan And okCur, vbInformation, vbExclamation), APP_NAME
    Exit Sub
EH:
    LogErr SRC, Err.description
    MsgBox "Greska pri objavljivanju: " & Err.description, vbCritical, APP_NAME
End Sub

' app_version = SemVer komparator (isti kao modUpdateGate / VERSION_LATEST).
' files = JSON niz {name,size,sha256} (klijent verifikuje kompletnost + SHA-256
' po fajlu). Isti sadrzaj sluzi i za flat version.json i za versioned manifest.json.
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

' F2 retention: zadrzi najnovijih 'keep' versioned foldera u releases\, ostalo u
' Trash (meko). Sortira po VersionCompare (SemVer) opadajuce. Broji/izlistava
' obrisane. Fail-soft (greska u prune-u ne obara objavu).
Private Sub PruneOldReleases(ByVal releasesFolderID As String, ByVal keep As Long, _
                             ByRef prunedOut As Long, ByRef listOut As String)
    On Error Resume Next
    If Len(releasesFolderID) = 0 Then Exit Sub
    Dim d As Object: Set d = DriveListFolder(releasesFolderID)
    If d Is Nothing Then Exit Sub
    If d.count = 0 Then Exit Sub

    Dim names() As String, cnt As Long
    ReDim names(0 To d.count - 1)
    Dim k As Variant
    For Each k In d.Keys
        ' samo verzija-slicna imena (poc. cifrom) - da slucajni ne-verzija folder ne strada
        If Len(CStr(k)) > 0 Then
            If IsNumeric(Left$(CStr(k), 1)) Then names(cnt) = CStr(k): cnt = cnt + 1
        End If
    Next k
    If cnt <= keep Then Exit Sub

    Dim i As Long, j As Long, tmp As String
    For i = 0 To cnt - 2
        For j = i + 1 To cnt - 1
            If VersionCompare(names(i), names(j)) < 0 Then tmp = names(i): names(i) = names(j): names(j) = tmp
        Next j
    Next i

    For i = keep To cnt - 1
        If DriveTrashFile(CStr(d(names(i)))) Then
            prunedOut = prunedOut + 1
            listOut = listOut & "  " & names(i) & vbCrLf
        End If
    Next i
End Sub

' F2 rollback (RUCNO, Alt+F8): prepisi current.json da pokazuje na STARIJU vec
' objavljenu verziju (releases\<ver>). Klijenti koji JOS nisu azurirani povuku tu
' verziju; vec-azurni ostaju (VersionCompare). manifest_sha256 se preracuna iz
' STVARNIH bajtova te verzije (koren lanca poverenja).
Public Sub RollbackReleaseTo()
    Const SRC As String = "modRelease.RollbackReleaseTo"
    On Error GoTo EH
    If Len(REL_FOLDER_ID) = 0 Then
        MsgBox "REL_FOLDER_ID nije postavljen u modConfig.bas.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim ver As String: ver = Trim$(InputBox("Rollback current.json na verziju (npr. 2.20.0):", APP_NAME))
    If Len(ver) = 0 Then Exit Sub

    Dim relSub As String: relSub = DriveEnsureFolder(REL_FOLDER_ID, "releases")
    Dim d As Object
    If Len(relSub) > 0 Then Set d = DriveListFolder(relSub)
    If d Is Nothing Then MsgBox "Ne mogu da procitam releases\ folder.", vbCritical, APP_NAME: Exit Sub
    If Not d.Exists(ver) Then
        MsgBox "Verzija nije objavljena: releases\" & ver, vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim vfid As String: vfid = CStr(d(ver))

    ' skini manifest.json te verzije -> manifest_sha256 STVARNIH bajtova
    Dim manId As String: manId = DriveFindInFolder(vfid, "manifest.json")
    If Len(manId) = 0 Then MsgBox "releases\" & ver & " nema manifest.json.", vbCritical, APP_NAME: Exit Sub
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim mPath As String: mPath = fso.GetSpecialFolder(2) & "\_rollback_manifest.json"
    If Not DriveDownloadToFile(manId, mPath) Then MsgBox "Ne mogu da skinem manifest.json.", vbCritical, APP_NAME: Exit Sub
    Dim manSha As String: manSha = Sha256File(mPath)

    If MsgBox("Prepisati current.json -> releases\" & ver & " ?" & vbCrLf & _
              "Klijenti koji JOS nisu azurirani ce povuci OVU (stariju) verziju.", _
              vbYesNo + vbExclamation, APP_NAME) <> vbYes Then Exit Sub

    Dim curJson As String
    curJson = "{""app_version"":""" & ver & """," & _
              """release_folder_id"":""" & vfid & """," & _
              """manifest_sha256"":""" & manSha & """," & _
              """published"":""" & Format$(Now, "yyyy-mm-dd hh:nn") & """,""rollback"":true}"
    Dim tmpCur As String: tmpCur = fso.GetSpecialFolder(2) & "\current.json"
    WriteReleaseTextFile tmpCur, curJson
    If Len(DriveUploadFile(REL_FOLDER_ID, tmpCur, "current.json")) > 0 Then
        MsgBox "current.json prepisan na releases\" & ver & "." & vbCrLf & _
               "manifest_sha256 = " & Left$(manSha, 12) & "...", vbInformation, APP_NAME
    Else
        MsgBox "GRESKA: current.json nije prepisan.", vbCritical, APP_NAME
    End If
    Exit Sub
EH:
    LogErr SRC, Err.description
    MsgBox "Greska pri rollback-u: " & Err.description, vbCritical, APP_NAME
End Sub

' F2 (RUCNO, Alt+F8): izlistaj objavljene versioned verzije + na koju pokazuje
' current.json (za izbor rollback cilja).
Public Sub ListReleases()
    Const SRC As String = "modRelease.ListReleases"
    On Error GoTo EH
    If Len(REL_FOLDER_ID) = 0 Then
        MsgBox "REL_FOLDER_ID nije postavljen u modConfig.bas.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim relSub As String: relSub = DriveEnsureFolder(REL_FOLDER_ID, "releases")
    Dim d As Object
    If Len(relSub) > 0 Then Set d = DriveListFolder(relSub)

    Dim names() As String, cnt As Long
    Dim i As Long, j As Long, tmp As String, k As Variant
    If Not d Is Nothing Then
        If d.count > 0 Then
            ReDim names(0 To d.count - 1)
            For Each k In d.Keys
                If Len(CStr(k)) > 0 Then
                    If IsNumeric(Left$(CStr(k), 1)) Then names(cnt) = CStr(k): cnt = cnt + 1
                End If
            Next k
            For i = 0 To cnt - 2
                For j = i + 1 To cnt - 1
                    If VersionCompare(names(i), names(j)) < 0 Then tmp = names(i): names(i) = names(j): names(j) = tmp
                Next j
            Next i
        End If
    End If

    Dim curVer As String
    curVer = Trim$(ExtractJsonStringGoogle(RelReadDriveText(REL_FOLDER_ID, "current.json"), "app_version"))

    Dim msg As String
    msg = "current.json -> " & IIf(Len(curVer) > 0, curVer, "(nema/flat)") & vbCrLf & vbCrLf & _
          "Objavljene verzije (releases\), najnovije prvo:" & vbCrLf
    If cnt = 0 Then
        msg = msg & "  (nema versioned foldera)"
    Else
        For i = 0 To cnt - 1
            msg = msg & "  " & names(i) & IIf(StrComp(names(i), curVer, vbTextCompare) = 0, "   <== current", "") & vbCrLf
        Next i
    End If
    MsgBox msg, vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr SRC, Err.description
    MsgBox "Greska (ListReleases): " & Err.description, vbCritical, APP_NAME
End Sub

' Skini imenovani tekstualni fajl iz Drive foldera -> sadrzaj (ili ""). modRelease
' lokalna kopija (modSelfUpdate ima svoju - moduli su nezavisni kanali).
Private Function RelReadDriveText(ByVal folderID As String, ByVal fileName As String) As String
    On Error Resume Next
    Dim id As String: id = DriveFindInFolder(folderID, fileName)
    If Len(id) = 0 Then Exit Function
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tmp As String: tmp = fso.GetSpecialFolder(2) & "\_rel_" & fileName
    If DriveDownloadToFile(id, tmp) Then
        Dim ff As Integer: ff = FreeFile
        Open tmp For Input As #ff
        If LOF(ff) > 0 Then RelReadDriveText = Input$(LOF(ff), ff)
        Close #ff
    End If
End Function
