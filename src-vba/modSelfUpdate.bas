Attribute VB_Name = "modSelfUpdate"
Option Explicit

' ============================================================
' modSelfUpdate - KLIJENT strana self-update-a (Funkcija A).
'
' Tok:
'   Workbook_Open -> StartApp -> CheckForUpdateOnOpen()  (fail-soft)
'     - cita AgriX_Release/version.json (modDrive), poredi app_version sa
'       lokalnim APP_VERSION (modUpdateGate.VersionCompare).
'     - ako je novije -> MsgBox vbYesNo. Na "Da" vrati True.
'   StartApp tada: Application.OnTime Now, "RunSelfUpdate" + Exit Sub.
'   RunSelfUpdate(): backup -> download -> PrepareRuntimeForSelfUpdate ->
'     code-merge import -> save -> "zatvori i otvori".
'
' KLJUCNO (root cause): moduli koji grade dinamicke kontrole sa WithEvents
' (modOtkupBlok/modPodesavanja/modMaticniLookups/modKarticaDetalji) drze ZIVE
' event-sink reference; dok su zive, editovanje njihovog CodeModule-a obara COM
' ([-2147417848] "disconnected"). Zato PrepareRuntimeForSelfUpdate PRE importa
' oslobodi sve te reference (Release/Reset) + unload sve forme + stop scheduled
' sync. Posle toga je CIST code-merge (AddFromString) bezbedan za SVE module -
' bez Remove/Import (nema modX1 duplikata), bez form-load greske.
'
' OPT-IN: radi samo ako je REL_FOLDER_ID postavljen (modConfig).
' FAIL-SOFT: svaka greska u proveri samo se loguje, start ide dalje.
' Reuse: modDrive (download/list), modUpdateGate.VersionCompare,
'        modGoogleAuth.ExtractJsonStringGoogle, StopScheduledSync, *_Release/Reset.
' VBA-pristup helperi su LOKALNI (Self*) - modVbaTools se ne update-uje na
' klijentu (SELF_MODULE skip) pa njegov Public ne stigne.
' Skip pri importu: modSelfUpdate (na stack-u) + modVbaTools (dev tool).
' ASCII-only modul (vidi CLAUDE.md, sekcija 4 - encoding).
' ============================================================

Private Const MANIFEST_NAME As String = "version.json"
Private Const SKIP_MODULES As String = "modSelfUpdate,modVbaTools"

' Vrati True ako je korisnik POTVRDIO azuriranje (tada pozivalac treba da
' zakaze RunSelfUpdate i prekine ostatak startup-a). Inace False (nastavi start).
Public Function CheckForUpdateOnOpen() As Boolean
    Const SRC As String = "modSelfUpdate.CheckForUpdateOnOpen"
    On Error GoTo EH

    If Len(REL_FOLDER_ID) = 0 Then Exit Function          ' opt-in

    Dim remoteVer As String: remoteVer = GetRemoteAppVersion()
    If Len(remoteVer) = 0 Then Exit Function               ' offline / nema manifesta -> tiho
    If VersionCompare(APP_VERSION, remoteVer) >= 0 Then Exit Function   ' vec azurno

    Application.Visible = True
    If MsgBox("Postoji nova verzija programa: " & remoteVer & vbCrLf & _
              "Trenutna verzija: " & APP_VERSION & vbCrLf & vbCrLf & _
              "Azurirati sada? (preporuceno)", _
              vbYesNo + vbQuestion, APP_NAME) = vbYes Then
        CheckForUpdateOnOpen = True
    End If
    Exit Function
EH:
    LogErr SRC, Err.description                            ' bug u checku ne sme da spreci start
End Function

' Pokrece se preko Application.OnTime (prazan stack). Public zbog OnTime.
Public Sub RunSelfUpdate()
    Const SRC As String = "modSelfUpdate.RunSelfUpdate"
    On Error GoTo EH

    If Not SelfVBAAccessible() Then Exit Sub

    ' 1) Backup pre svega (rollback ako import pukne)
    If Not MakePreUpdateBackup() Then
        If MsgBox("Backup pre azuriranja nije uspeo." & vbCrLf & _
                  "Nastaviti ipak?", vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Sub
    End If

    ' 2) Download svih fajlova iz AgriX_Release u temp folder
    Dim tempDir As String: tempDir = MakeTempDir()
    Dim n As Long: n = DownloadReleaseFiles(tempDir)
    If n = 0 Then
        MsgBox "Preuzimanje nije uspelo (0 fajlova). Azuriranje otkazano." & vbCrLf & _
               "Pokusajte ponovo ili preuzmite novu verziju rucno.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' 3) Oslobodi runtime stanje (root cause fix) PRE importa
    PrepareRuntimeForSelfUpdate

    ' 4) Code-merge import svih komponenti
    Dim summary As String: summary = ImportFromFolder(tempDir, SKIP_MODULES)

    ' 5) Snimi (restart aktivira nov kod + schema self-heal kroz InitApp)
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo EH

    MsgBox "Azuriranje zavrseno. Preuzeto fajlova: " & n & vbCrLf & _
           summary & vbCrLf & vbCrLf & _
           "ZATVORITE i ponovo OTVORITE fajl da se promene aktiviraju.", _
           vbInformation, APP_NAME
    Exit Sub
EH:
    LogErr SRC, Err.description
    MsgBox "Greska pri azuriranju: " & Err.description & vbCrLf & vbCrLf & _
           "Ako program ne radi ispravno, vratite kopiju iz 'Backup' foldera " & _
           "(AgriX_pre-update_*.xlsm).", vbCritical, APP_NAME
End Sub

' Oslobodi runtime stanje pre self-update importa: ugasi evente/sync, otpusti
' module-level reference dinamickih kontrola/WithEvents (inace CodeModule edit
' tih modula diskonektuje COM), unload sve forme. Best-effort (On Error Resume
' Next) - ako neki Release fali (nije u toj verziji), preskoci.
Private Sub PrepareRuntimeForSelfUpdate()
    On Error Resume Next

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    StopScheduledSync               ' modGoogleSyncOrchestrator (otkazi pending OnTime sync)

    ' Release module-level WithEvents/kontrole (dinamicki paneli)
    OtkupBlok_Release               ' modOtkupBlok (clsBlokUI)
    Podesavanja_Release             ' modPodesavanja (clsConfigBtn)
    MaticniMenu_Release             ' modMaticniLookups (clsLookupMenuBtn)
    KarticaDetalji_Reset            ' modKarticaDetalji

    ' Unload sve forme (otpusti njihove kontrole / event sink-ove)
    Do While VBA.UserForms.count > 0
        Unload VBA.UserForms(0)
    Loop

    DoEvents
    On Error GoTo 0
End Sub

' ---------------- private ----------------

' Procitaj app_version iz AgriX_Release/version.json (ili "" ako nedostupno).
Private Function GetRemoteAppVersion() As String
    Const SRC As String = "modSelfUpdate.GetRemoteAppVersion"
    On Error GoTo EH

    Dim id As String: id = DriveFindInFolder(REL_FOLDER_ID, MANIFEST_NAME)
    If Len(id) = 0 Then Exit Function

    Dim tmp As String: tmp = Environ$("TEMP") & "\AgriX_version.json"
    If Not DriveDownloadToFile(id, tmp) Then Exit Function

    GetRemoteAppVersion = Trim$(ExtractJsonStringGoogle(ReadAllText(tmp), "app_version"))
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Snimi kopiju trenutnog .xlsm u <putanja>\Backup\AgriX_pre-update_*.xlsm.
Private Function MakePreUpdateBackup() As Boolean
    On Error GoTo EH
    ' NB: ne zovi promenljivu 'dir' - sudara se sa ugradjenom Dir() ("expected array").
    Dim bkDir As String: bkDir = ThisWorkbook.path & "\Backup"
    If Dir(bkDir, vbDirectory) = "" Then MkDir bkDir
    Dim nm As String
    nm = "AgriX_pre-update_" & APP_VERSION & "_" & Format$(Now, "yyyy-mm-dd_hhmm") & ".xlsm"
    ThisWorkbook.SaveCopyAs bkDir & "\" & nm
    MakePreUpdateBackup = True
    Exit Function
EH:
    MakePreUpdateBackup = False
End Function

' Napravi prazan temp folder za preuzimanje (obrise stari ako postoji).
Private Function MakeTempDir() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim d As String: d = fso.GetSpecialFolder(2) & "\AgriX_update"   ' TemporaryFolder
    On Error Resume Next
    If fso.FolderExists(d) Then fso.DeleteFolder d, True
    On Error GoTo 0
    fso.CreateFolder d
    MakeTempDir = d
End Function

' Skini sve kod-fajlove iz AgriX_Release u tempDir. Vrati broj skinutih.
Private Function DownloadReleaseFiles(ByVal tempDir As String) As Long
    Const SRC As String = "modSelfUpdate.DownloadReleaseFiles"
    On Error GoTo EH

    Dim dict As Object: Set dict = DriveListFolder(REL_FOLDER_ID)
    If dict Is Nothing Then Exit Function

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim k As Variant, ext As String, n As Long
    For Each k In dict.Keys
        ext = LCase$(fso.GetExtensionName(CStr(k)))
        Select Case ext
            Case "bas", "cls", "frm", "frx", "doccls"
                If DriveDownloadToFile(CStr(dict(k)), tempDir & "\" & CStr(k)) Then n = n + 1
        End Select
    Next k
    DownloadReleaseFiles = n
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Cist code-merge u mestu (DeleteLines + AddFromString). Radi za SVE jer je
' runtime stanje oslobodjeno (PrepareRuntimeForSelfUpdate) pa CodeModule edit
' vise ne diskonektuje. Bez Remove (nema modX1 duplikata), bez Import (nema
' form-load greske). Dizajn formi (.frx) se NE menja - za to treba pun reinstall.
' Retry (3 prolaza) + per-modul greska kao safety net.
Private Function ImportFromFolder(ByVal folder As String, ByVal skipCsv As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim skip As String: skip = "," & LCase$(skipCsv) & ","
    Dim st As Object: Set st = CreateObject("Scripting.Dictionary")   ' fajl(lower) -> "ok"/"skip"
    Dim er As Object: Set er = CreateObject("Scripting.Dictionary")   ' fajl(lower) -> poslednja greska
    Dim pass As Long, anyLeft As Boolean, usedPass As Long
    Dim fil As Object, ext As String, baseName As String, fkey As String
    Dim body As String, vbc As Object, proj As Object

    For pass = 1 To 3
        usedPass = pass
        anyLeft = False
        Set proj = ThisWorkbook.VBProject
        For Each fil In fso.GetFolder(folder).files
            ext = LCase$(fso.GetExtensionName(fil.name))
            If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "doccls" Then
                baseName = fso.GetBaseName(fil.name)
                fkey = LCase$(fil.name)
                If InStr(skip, "," & LCase$(baseName) & ",") = 0 And Not st.Exists(fkey) Then
                    On Error Resume Next
                    Err.Clear
                    body = ExtractModuleCode(fil.path)
                    Set vbc = Nothing
                    Set vbc = proj.VBComponents(baseName)
                    If Not vbc Is Nothing Then
                        If vbc.CodeModule.CountOfLines > 0 Then _
                            vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
                        If Len(body) > 0 Then vbc.CodeModule.AddFromString body
                        If Err.Number = 0 Then st(fkey) = "ok"
                    ElseIf ext = "bas" Or ext = "cls" Then
                        Set vbc = proj.VBComponents.Add(IIf(ext = "bas", 1, 2))
                        vbc.name = baseName
                        If Len(body) > 0 Then vbc.CodeModule.AddFromString body
                        If Err.Number = 0 Then st(fkey) = "ok"
                    Else
                        st(fkey) = "skip"     ' nova forma/sheet -> reinstall
                    End If
                    If Err.Number <> 0 Then
                        er(fkey) = "[" & Err.Number & "] " & Err.description
                        anyLeft = True
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        Next fil
        If Not anyLeft Then Exit For
    Next pass

    Dim okN As Long, k As Variant, failS As String, skipS As String
    For Each k In st.Keys
        If st(k) = "ok" Then
            okN = okN + 1
        Else
            skipS = skipS & "  " & k & vbCrLf
        End If
    Next k
    For Each k In er.Keys
        If Not st.Exists(CStr(k)) Then failS = failS & "  " & k & " -> " & er(k) & vbCrLf
    Next k

    ImportFromFolder = "Azurirano: " & okN & " (prolaza: " & usedPass & ")" & _
        IIf(Len(skipS) > 0, vbCrLf & "Preskoceno (novo, reinstall):" & vbCrLf & skipS, "") & _
        IIf(Len(failS) > 0, vbCrLf & "GRESKE:" & vbCrLf & failS, "")
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff
    If LOF(ff) > 0 Then ReadAllText = Input$(LOF(ff), ff)
    Close #ff
End Function

' True ako VBA projekat ima programski pristup (inace prijavi sta da ukljuci).
Private Function SelfVBAAccessible() As Boolean
    On Error Resume Next
    Dim c As Long
    c = ThisWorkbook.VBProject.VBComponents.count
    SelfVBAAccessible = (Err.Number = 0)
    On Error GoTo 0
    If Not SelfVBAAccessible Then
        MsgBox "Nema programskog pristupa VBA projektu." & vbCrLf & vbCrLf & _
               "Ukljuci: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
               "Macro Settings > 'Trust access to the VBA project object model'.", _
               vbExclamation, APP_NAME
    End If
End Function

' Izvuci editabilni kod iz izvoznog VBA fajla (.bas/.cls/.frm/.doccls) bezbedan
' za AddFromString:
'  - preskoci header: VERSION, Begin..End dizajn blok (forme/cls, uklj.
'    BeginProperty/EndProperty i ugnezdene kontrole), module Attribute linije;
'  - u kodu STRIP-uj sve "Attribute ..." linije (member atributi tipa
'    Attribute x.VB_VarHelpID = -1) - AddFromString ih ne prima (Syntax error).
' Case-insensitive (cls: BEGIN/END velikim; forme: Begin/End).
Private Function ExtractModuleCode(ByVal path As String) As String
    Dim allTxt As String, arr() As String, i As Long, depth As Long
    Dim inHeader As Boolean, ls As String, u As String, body As String

    allTxt = ReadAllText(path)
    arr = Split(allTxt, vbCrLf)
    If UBound(arr) <= 0 Then arr = Split(allTxt, vbLf)

    inHeader = True
    For i = 0 To UBound(arr)
        ls = LTrim$(arr(i))
        u = UCase$(ls)
        If inHeader Then
            If depth > 0 Then
                If u Like "BEGIN*" Then
                    depth = depth + 1
                ElseIf u Like "END*" Then
                    depth = depth - 1
                End If
            ElseIf u Like "VERSION *" Then
                ' skip
            ElseIf u Like "BEGIN*" Then
                depth = depth + 1
            ElseIf u Like "ATTRIBUTE *" Then
                ' module attribute -> skip
            ElseIf Len(Trim$(arr(i))) = 0 Then
                ' prazna linija u headeru -> skip
            Else
                inHeader = False
                body = arr(i)                ' prva kod linija (nije Attribute)
            End If
        Else
            If Not (u Like "ATTRIBUTE *") Then
                If Len(body) = 0 Then body = arr(i) Else body = body & vbCrLf & arr(i)
            End If
        End If
    Next i

    ExtractModuleCode = body
End Function
