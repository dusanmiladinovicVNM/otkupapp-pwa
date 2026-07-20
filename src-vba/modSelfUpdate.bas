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
'     code-merge (faza 1) -> za "tvrde" Remove+OnTime+Import (faza 2) -> save.
'
' KLJUCNO (potvrdjen root cause): moduli sa MODULE-LEVEL deklaracijom MSForms
' kontrole (Private mBtn As MSForms.CommandButton ...) - modOtkupBlok(20),
' modKarticaDetalji(2), modPodesavanja(1) - NE mogu kroz AddFromString: dodavanje
' tih deklaracija bind-uje MSForms tip-biblioteku u toku COM edita pa diskonektuje
' CodeModule ([-2147417848]). (DeleteLines prodje; pada tek AddFromString.
' Nije stvar zivih instanci - Release nije pomogao.) Resenje: njih ide FAZA 2
' preko Import-a (rekreacija komponente, podnosi MSForms decls; radi u
' ImportAllVBA), uz Remove->flush(OnTime)->Import da nema modX1 duplikata.
' PrepareRuntimeForSelfUpdate (unload formi + release + stop sync/tajmeri) je
' higijena pre Remove-a (da forma ne drzi kontrole tih modula).
'
' HARDENING (posle crash-a na velikim update-ovima, v2.16.1 -> v2.21.0):
'  - DELTA-SKIP: komponenta ciji je kod IDENTICAN novom telu se NE dira
'    (ranije se prepisivao CEO projekat ~90 komponenti na svaki update ->
'    nepotrebni COM editi = nepotreban rizik). Faza 2 se sada desava samo
'    kad je neki "tvrd" modul stvarno izmenjen.
'  - FORME/SHEET NIKAD U FAZU 2: u failed (-> VBComponents.Remove) idu SAMO
'    .bas/.cls. Remove forme u runtime-u = zamka #1 ("Errors during load",
'    korupcija, Document Recovery = crash Excela), a faza 2 ionako uvozi samo
'    .bas/.cls pa bi forma i TRAJNO nestala iz projekta. Forma ciji merge
'    padne: best-effort rollback na STARI kod + poruka "potreban reinstall".
'  - TAJMERI: pre importa se otkazuju SVI poznati Application.OnTime tikovi
'    (sync + AutoSaveTick + StanicaLock heartbeat) - tik koji opali izmedju
'    faza (dok su moduli uklonjeni) forsira compile polomljenog projekta,
'    a AutoSave bi jos i SNIMIO polu-azuriran fajl.
'  - EnableEvents/ScreenUpdating se VRACAJU na svakom izlazu update toka -
'    inace Workbook_Open ne opali pri ponovnom otvaranju u istoj Excel
'    instanci ("zatvori i otvori" izgleda kao da je update ubio aplikaciju).
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
              Poruka("SU_AZURIRATI_SADA"), _
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
        If MsgBox(Poruka("SU_BACKUP_NIJE") & vbCrLf & _
                  "Nastaviti ipak?", vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Sub
    End If

    ' 2) Download svih fajlova iz AgriX_Release u temp folder
    Dim tempDir As String: tempDir = MakeTempDir()
    Dim n As Long: n = DownloadReleaseFiles(tempDir)
    If n = 0 Then
        MsgBox Poruka("SU_PREUZIMANJE_OTKAZANO") & vbCrLf & _
               Poruka("SU_POKUSAJTE"), vbCritical, APP_NAME
        Exit Sub
    End If

    ' 3-6) Zajednicko jezgro (PrepareRuntime -> code-merge -> faza 2 -> save).
    '      Isti put koristi i RunSelfUpdateDev (lokalni folder umesto Drive-a),
    '      pa se "tvrda" faza-2 logika ne duplira.
    RunSelfUpdateCore tempDir, n
    Exit Sub
EH:
    RestoreRuntimeAfterSelfUpdate
    LogErr SRC, Err.description
End Sub

' Zajednicko jezgro update-a: uzmi VEC pripremljen folder (tempDir) pun
' src-vba fajlova - skinut sa Drive-a (RunSelfUpdate) ILI kopiran iz lokalnog
' git klona (RunSelfUpdateDev) - i odradi IDENTICAN merge: PrepareRuntime ->
' faza 1 code-merge -> faza 2 (Remove+Import za "tvrde") -> save. n = broj
' fajlova (za poruku). Ovde su events/screen vec spremni za gasenje.
Private Sub RunSelfUpdateCore(ByVal tempDir As String, ByVal n As Long)
    Const SRC As String = "modSelfUpdate.RunSelfUpdateCore"
    On Error GoTo EH

    ' 3) Oslobodi runtime stanje (root cause fix) PRE importa
    PrepareRuntimeForSelfUpdate

    ' 4) Code-merge import. failed = moduli koje AddFromString ne moze (module-
    '    level MSForms kontrole -> bind na MSForms tip-biblioteku diskonektuje
    '    CodeModule). Njih ide FAZA 2 preko Import-a (koji to podnosi).
    Dim failed As Object: Set failed = CreateObject("Scripting.Dictionary")
    Dim summary As String: summary = ImportFromFolder(tempDir, SKIP_MODULES, failed)

    ' 5) Faza 2: ukloni "tvrde" SADA (Remove se flush-uje kad makro zavrsi), pa
    '    ih uvezi Import-om u sledecem OnTime prolazu (rekreacija komponente -
    '    drugi mehanizam, podnosi MSForms decls; flush -> bez modX1 duplikata).
    If failed.count > 0 Then
        Dim proj As Object: Set proj = ThisWorkbook.VBProject
        Dim fk As Variant, remC As Object
        For Each fk In failed.Keys
            On Error Resume Next
            Set remC = Nothing
            Set remC = proj.VBComponents(CStr(fk))
            ' Bezbednosna brana (zamka #1): Remove SME samo std (1) / class (2)
            ' modul. Formu/dokument modul NIKAD - Remove forme u runtime-u pravi
            ' korupciju (Document Recovery), a faza 2 je ne bi ni vratila.
            If Not remC Is Nothing Then
                If remC.Type = 1 Or remC.Type = 2 Then proj.VBComponents.Remove remC
            End If
            On Error GoTo EH
        Next fk
        SaveSetting "AgriXSelfUpdate", "phase2", "dir", tempDir
        SaveSetting "AgriXSelfUpdate", "phase2", "n", CStr(n)
        ' NAPOMENA: EnableEvents ostaje ISKLJUCEN do kraja faze 2 (zastita
        ' prozora u kome su moduli uklonjeni); vraca ga RunSelfUpdatePhase2.
        Application.OnTime Now + TimeSerial(0, 0, 2), "RunSelfUpdatePhase2"
        Exit Sub                  ' kraj makroa -> Remove se flush-uje -> faza 2
    End If

    ' 6) Snimi (sve proslo code-merge-om, nema faze 2)
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo EH

    RestoreRuntimeAfterSelfUpdate

    MsgBox Poruka("SU_ZAVRSENO_FAJLOVA") & n & vbCrLf & _
           summary & vbCrLf & vbCrLf & _
           "ZATVORITE i ponovo OTVORITE fajl da se promene aktiviraju.", _
           vbInformation, APP_NAME
    Exit Sub
EH:
    Dim errTxt As String: errTxt = Err.description   ' pre RestoreRuntime (On Error resetuje Err)
    RestoreRuntimeAfterSelfUpdate
    LogErr SRC, errTxt
    MsgBox Poruka("SU_GRESKA_AZURIRANJE") & errTxt & vbCrLf & vbCrLf & _
           "Ako program ne radi ispravno, vratite kopiju iz 'Backup' foldera " & _
           "(AgriX_pre-update_*.xlsm).", vbCritical, APP_NAME
End Sub

' ============================================================
' DEV TEST (RUCNO, Alt+F8): RunSelfUpdateDev
'
' Najlaksi nacin da se self-update engine testira NA SVOJOJ masini bez Drive-a:
' code-merge iz LOKALNOG src-vba foldera (git klon), kroz ISTI RunSelfUpdateCore
' kao pravi self-update. Testira ono sto ImportAllVBA NE testira - bas code-merge
' put (DeleteLines+AddFromString), gde su forme pucale. Ne dira flotu (bez
' PublishReleaseToDrive), ne trazi Google auth ni REL_FOLDER_ID.
'
' POSTUPAK:
'   1) Otvori KOPIJU klijentske sveske (ne originalni build-master).
'   2) Alt+F8 -> RunSelfUpdateDev -> izaberi svoj ...\otkupapp-pwa\src-vba\ folder.
'   3) Backup se napravi sam; merge tece; na kraju "zatvori i otvori".
'   4) Posle restarta: Alt+F11 -> nema duplikata (modX1); Debug->Compile cist;
'      probaj forme (Dokumenta "Storno", Integritet overlay...).
'   Rollback po potrebi: Backup\AgriX_pre-update_*.xlsm.
'
' NB: modSelfUpdate je u SKIP_MODULES -> ovaj DEV kod se pri merge-u NE prepisuje
' (harness ostaje), isto kao u produkciji.
' ============================================================
Public Sub RunSelfUpdateDev()
    Const SRC As String = "modSelfUpdate.RunSelfUpdateDev"
    On Error GoTo EH

    If Not SelfVBAAccessible() Then Exit Sub

    If MsgBox("DEV TEST self-update-a iz LOKALNOG foldera (bez Drive-a)." & vbCrLf & vbCrLf & _
              "Code-merge-uje OVU svesku iz izabranog src-vba foldera, isto kao pravi " & _
              "self-update (faza 1 + faza 2). Pokreni na KOPIJI klijenta!" & vbCrLf & vbCrLf & _
              "Nastaviti?", vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    ' 1) Izbor src-vba foldera (git klon)
    Dim srcFolder As String: srcFolder = PickFolderDev()
    If Len(srcFolder) = 0 Then Exit Sub

    ' 2) Backup pre svega (rollback ako merge pukne) - reuse produkcijskog
    If Not MakePreUpdateBackup() Then
        If MsgBox("Backup nije uspeo. Nastaviti ipak?", _
                  vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Sub
    End If

    ' 3) Kopiraj kod fajlove u cist temp (mesto Drive download-a) -> isti ulaz u core
    Dim tempDir As String: tempDir = MakeTempDir()
    Dim n As Long: n = CopyCodeFilesDev(srcFolder, tempDir)
    If n = 0 Then
        MsgBox "U izabranom folderu nema src-vba fajlova (.bas/.cls/.frm/.frx/.doccls)." & vbCrLf & _
               "Folder: " & srcFolder, vbCritical, APP_NAME
        Exit Sub
    End If

    ' 4) Isti core kao pravi self-update (PrepareRuntime -> merge -> faza 2 -> save)
    RunSelfUpdateCore tempDir, n
    Exit Sub
EH:
    RestoreRuntimeAfterSelfUpdate
    LogErr SRC, Err.description
    MsgBox "DEV self-update greska: " & Err.description, vbCritical, APP_NAME
End Sub

' Folder picker za DEV test (izaberi src-vba). "" na Cancel / nedostupno.
Private Function PickFolderDev() As String
    On Error Resume Next
    Dim fd As Object: Set fd = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    If fd Is Nothing Then Exit Function
    fd.Title = "Izaberi src-vba folder (git klon)"
    fd.InitialFileName = ThisWorkbook.path & "\"
    If fd.Show = -1 Then PickFolderDev = fd.SelectedItems(1)
End Function

' Kopiraj SAMO kod fajlove (isti filter kao DownloadReleaseFiles) iz lokalnog
' foldera u tempDir. Vrati broj kopiranih. Izvor se NE dira (samo citanje).
Private Function CopyCodeFilesDev(ByVal srcFolder As String, ByVal tempDir As String) As Long
    Const SRC As String = "modSelfUpdate.CopyCodeFilesDev"
    On Error GoTo EH

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(srcFolder) Then Exit Function

    Dim fil As Object, ext As String, n As Long
    For Each fil In fso.GetFolder(srcFolder).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        Select Case ext
            Case "bas", "cls", "frm", "frx", "doccls"
                fso.CopyFile fil.path, tempDir & "\" & fil.name, True
                n = n + 1
        End Select
    Next fil
    CopyCodeFilesDev = n
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Faza 2 (Application.OnTime; posle flush-a Remove-ova iz faze 1). Uvezi (Import)
' module uklonjene u fazi 1 - sad stvarno obrisane pa Import pravi cist modul
' (bez duplikata). Import podnosi module-level MSForms kontrole koje AddFromString
' ne moze. Public zbog OnTime.
Public Sub RunSelfUpdatePhase2()
    Const SRC As String = "modSelfUpdate.RunSelfUpdatePhase2"
    On Error GoTo EH

    Dim p2dir As String: p2dir = GetSetting("AgriXSelfUpdate", "phase2", "dir", "")
    Dim nTxt As String: nTxt = GetSetting("AgriXSelfUpdate", "phase2", "n", "?")
    DeleteSetting "AgriXSelfUpdate", "phase2"
    If Len(p2dir) = 0 Then
        RestoreRuntimeAfterSelfUpdate       ' faza 1 je ostavila events OFF
        Exit Sub
    End If

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fil As Object, ext As String, baseName As String, imported As Long, stillFail As String
    Dim comExists As Boolean, tmpc As Object

    For Each fil In fso.GetFolder(p2dir).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        If ext = "bas" Or ext = "cls" Then
            baseName = fso.GetBaseName(fil.name)
            comExists = True
            On Error Resume Next
            Set tmpc = Nothing
            Set tmpc = proj.VBComponents(baseName)
            comExists = Not (tmpc Is Nothing)
            On Error GoTo 0
            If Not comExists Then              ' uvezi samo uklonjene (faza 1)
                On Error Resume Next
                Err.Clear
                proj.VBComponents.Import fil.path
                If Err.Number = 0 Then
                    imported = imported + 1
                Else
                    stillFail = stillFail & "  " & fil.name & " -> [" & Err.Number & "] " & Err.description & vbCrLf
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next fil

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo EH

    RestoreRuntimeAfterSelfUpdate

    MsgBox Poruka("SU_ZAVRSENO_PREUZETO") & nTxt & ", 2. faza uvezeno: " & imported & _
           IIf(Len(stillFail) > 0, vbCrLf & vbCrLf & "I DALJE NIJE USPELO:" & vbCrLf & stillFail, "") & vbCrLf & vbCrLf & _
           "ZATVORITE i ponovo OTVORITE fajl da se promene aktiviraju.", _
           IIf(Len(stillFail) > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub
EH:
    Dim errTxt As String: errTxt = Err.description   ' pre RestoreRuntime (On Error resetuje Err)
    RestoreRuntimeAfterSelfUpdate
    LogErr SRC, errTxt
    MsgBox Poruka("SU_GRESKA_2FAZA") & errTxt & vbCrLf & _
           "Vratite kopiju iz 'Backup' foldera (AgriX_pre-update_*.xlsm).", vbCritical, APP_NAME
End Sub

' Oslobodi runtime stanje pre self-update importa: ugasi evente/sync, otpusti
' module-level reference dinamickih kontrola/WithEvents (inace CodeModule edit
' tih modula diskonektuje COM), unload sve forme. Best-effort (On Error Resume
' Next) - ako neki Release fali (nije u toj verziji), preskoci.
Private Sub PrepareRuntimeForSelfUpdate()
    On Error Resume Next

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Otkazi SVE poznate Application.OnTime tikove. Tik koji opali usred
    ' importa (ili u prozoru izmedju faze 1 i 2, dok su "tvrdi" moduli
    ' uklonjeni) forsira demand-compile polomljenog projekta; AutoSaveTick
    ' bi uz to jos i SNIMIO polu-azuriran fajl preko radne kopije.
    StopScheduledSync               ' modGoogleSyncOrchestrator (otkazi pending OnTime sync)
    StopAutoSaveTimer               ' modJournaling (otkazi pending AutoSaveTick)
    StopHeartbeatTimer              ' modStanicaLock (otkazi 90s heartbeat)

    ' Release module-level WithEvents/kontrole (dinamicki paneli)
    OtkupBlok_Release               ' modOtkupBlok (clsBlokUI)
    Podesavanja_Release             ' modPodesavanja (clsConfigBtn)
    MaticniMenu_Release             ' modMaticniLookups (clsLookupMenuBtn)
    KarticaDetalji_Reset            ' modKarticaDetalji
    MouseWheel_Off                  ' modMouseWheel (skini LL mouse hook pre izmene koda)

    ' Unload sve forme (otpusti njihove kontrole / event sink-ove)
    Do While VBA.UserForms.count > 0
        Unload VBA.UserForms(0)
    Loop

    DoEvents
    On Error GoTo 0
End Sub

' Vrati aplikaciona podesavanja koja PrepareRuntimeForSelfUpdate gasi. MORA se
' pozvati na SVAKOM izlazu update toka (uspeh, greska, prekid) - inace
' Workbook_Open ne opali pri sledecem otvaranju fajla u ISTOJ Excel instanci
' ("zatvori i otvori" tada izgleda kao da je update ubio aplikaciju), a
' Workbook_BeforeClose higijena se preskoci.
Private Sub RestoreRuntimeAfterSelfUpdate()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
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

' Cist code-merge u mestu (DeleteLines + AddFromString), SAMO za komponente cije
' se telo stvarno razlikuje od novog (DELTA-SKIP: identican kod se NE dira ->
' drasticno manje COM edita po update-u, manji rizik, brze). Radi za SVE jer je
' runtime stanje oslobodjeno (PrepareRuntimeForSelfUpdate) pa CodeModule edit
' vise ne diskonektuje. Bez Remove (nema modX1 duplikata), bez Import (nema
' form-load greske). Dizajn formi (.frx) se NE menja - za to treba pun reinstall.
' failedOut (kandidati za fazu 2 Remove+Import): SAMO .bas/.cls! Forme/sheet
' komponente NIKAD (Remove forme = korupcija/crash, zamka #1; faza 2 ih ne bi ni
' vratila jer uvozi samo .bas/.cls). Forma ciji merge padne posle svih prolaza:
' best-effort rollback na stari kod + "potreban reinstall" u izvestaju.
' Retry (3 prolaza) + per-modul greska kao safety net.
Private Function ImportFromFolder(ByVal folder As String, ByVal skipCsv As String, ByRef failedOut As Object) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim skip As String: skip = "," & LCase$(skipCsv) & ","
    Dim st As Object: Set st = CreateObject("Scripting.Dictionary")   ' fajl(lower) -> "ok"/"same"/"skip"
    Dim er As Object: Set er = CreateObject("Scripting.Dictionary")   ' fajl(lower) -> poslednja greska
    Dim orig As Object: Set orig = CreateObject("Scripting.Dictionary") ' fajl(lower) -> kod PRE 1. izmene (rollback formi)
    Dim pass As Long, anyLeft As Boolean, usedPass As Long
    Dim fil As Object, ext As String, baseName As String, fkey As String
    Dim body As String, cur As String, readOk As Boolean, extractOk As Boolean
    Dim vbc As Object, proj As Object

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
                    ' Ekstrakcija mora biti cista I ne-prazna (sem sheet modula,
                    ' koji su legitimno prazni) PRE nego sto se komponenta takne -
                    ' inace bi se DeleteLines uradio bez validnog novog tela.
                    extractOk = (Err.Number = 0 And (Len(body) > 0 Or ext = "doccls"))
                    If Not extractOk Then
                        If Err.Number = 0 Then _
                            Err.Raise vbObjectError + 2801, , "prazno telo posle ekstrakcije (" & fil.name & ")"
                        ' (Resume Next: Err ostaje -> zavrsi dole u er + retry)
                    Else
                        Set vbc = Nothing
                        Set vbc = proj.VBComponents(baseName)
                        If Not vbc Is Nothing Then
                            ' Delta-skip: procitaj postojeci kod i uporedi sa novim.
                            cur = ""
                            If vbc.CodeModule.CountOfLines > 0 Then _
                                cur = vbc.CodeModule.Lines(1, vbc.CodeModule.CountOfLines)
                            readOk = (Err.Number = 0)
                            If readOk And SameCode(cur, body) Then
                                st(fkey) = "same"              ' identican kod -> NE diraj
                            Else
                                Err.Clear                      ' greska citanja nije greska merge-a
                                ' Zapamti original SAMO iz prvog (jos netaknutog) pokusaja,
                                ' za rollback formi ciji merge trajno padne.
                                If readOk And Not orig.Exists(fkey) Then orig(fkey) = cur
                                If vbc.CodeModule.CountOfLines > 0 Then _
                                    vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
                                If Len(body) > 0 Then vbc.CodeModule.AddFromString body
                                If Err.Number = 0 Then st(fkey) = "ok"
                            End If
                        ElseIf ext = "bas" Or ext = "cls" Then
                            Err.Clear     ' Err 9 od lookup-a iznad ne sme da "oboji" Add put
                            Set vbc = proj.VBComponents.Add(IIf(ext = "bas", 1, 2))
                            vbc.name = baseName
                            If Len(body) > 0 Then vbc.CodeModule.AddFromString body
                            If Err.Number = 0 Then st(fkey) = "ok"
                        Else
                            st(fkey) = "skip"     ' nova forma/sheet -> reinstall
                        End If
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

    Dim okN As Long, sameN As Long, k As Variant, failS As String, skipS As String
    For Each k In st.Keys
        Select Case st(k)
            Case "ok":   okN = okN + 1
            Case "same": sameN = sameN + 1
            Case Else:   skipS = skipS & "  " & k & vbCrLf
        End Select
    Next k
    For Each k In er.Keys
        If Not st.Exists(CStr(k)) Then
            ext = LCase$(fso.GetExtensionName(CStr(k)))
            If ext = "bas" Or ext = "cls" Then
                ' Kandidat za fazu 2 (Remove+Import podnosi MSForms decls).
                failS = failS & "  " & k & " -> " & er(k) & vbCrLf
                If Not failedOut Is Nothing Then failedOut(fso.GetBaseName(CStr(k))) = folder & "\" & CStr(k)
            Else
                ' Forma/sheet: NIKAD u fazu 2. Vrati stari (radni) kod da klijent
                ' ostane upotrebljiv; nova verzija te forme stize reinstall-om.
                RestoreComponentCode orig, CStr(k)
                failS = failS & "  " & k & " -> " & er(k) & _
                        " (forma/sheet - vracen stari kod, potreban reinstall)" & vbCrLf
            End If
        End If
    Next k

    ImportFromFolder = "Azurirano: " & okN & ", bez izmene: " & sameN & " (prolaza: " & usedPass & ")" & _
        IIf(Len(skipS) > 0, vbCrLf & "Preskoceno (novo, reinstall):" & vbCrLf & skipS, "") & _
        IIf(Len(failS) > 0, vbCrLf & "GRESKE:" & vbCrLf & failS, "")
End Function

' Da li su dva tela koda identicna. Ignorise SAMO zavrsne CR/LF (fajl moze da
' se zavrsava praznim redom koji CodeModule.Lines ne vraca); sve ostalo mora
' biti bajt-za-bajt isto (izvori su ASCII, exporti iz istog VBE).
Private Function SameCode(ByVal a As String, ByVal b As String) As Boolean
    Do While Len(a) > 0
        Select Case Right$(a, 1)
            Case vbCr, vbLf: a = Left$(a, Len(a) - 1)
            Case Else: Exit Do
        End Select
    Loop
    Do While Len(b) > 0
        Select Case Right$(b, 1)
            Case vbCr, vbLf: b = Left$(b, Len(b) - 1)
            Case Else: Exit Do
        End Select
    Loop
    SameCode = (StrComp(a, b, vbBinaryCompare) = 0)
End Function

' Best-effort rollback koda JEDNE komponente na sadrzaj pre merge pokusaja
' (orig snimljen u ImportFromFolder pre prvog DeleteLines). Za forme/sheet ciji
' AddFromString ne prodje ni u 3 prolaza: bolje staro radno telo nego prazno /
' polu-upisano. Tiho (Resume Next) - rollback ne sme da obori ostatak update-a.
Private Sub RestoreComponentCode(ByVal orig As Object, ByVal fkey As String)
    On Error Resume Next
    If orig Is Nothing Then Exit Sub
    If Not orig.Exists(fkey) Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim vbc As Object
    Set vbc = ThisWorkbook.VBProject.VBComponents(fso.GetBaseName(fkey))
    If vbc Is Nothing Then Exit Sub

    If vbc.CodeModule.CountOfLines > 0 Then _
        vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
    If Len(orig(fkey)) > 0 Then vbc.CodeModule.AddFromString CStr(orig(fkey))
End Sub

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
        MsgBox Poruka("SU_TRUST"), vbExclamation, APP_NAME
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
