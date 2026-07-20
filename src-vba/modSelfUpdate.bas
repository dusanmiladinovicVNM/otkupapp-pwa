Attribute VB_Name = "modSelfUpdate"
Option Explicit

' ============================================================
' modSelfUpdate - KLIJENT strana self-update-a (Funkcija A).
'
' Tok:
'   Workbook_Open -> StartApp -> CheckForUpdateOnOpen() (koji prvo pozove
'     RecoverPendingSelfUpdate watchdog; fail-soft) -> na "Da" OnTime RunSelfUpdate.
'   RunSelfUpdate(): backup -> download (KOMPLETAN) -> RunSelfUpdateCore.
'   RunSelfUpdateCore: PrepareRuntime -> faza 1 code-merge -> faza 2 (tvrdi) -> save.
'
' KLJUCNO (potvrdjen root cause): moduli sa MODULE-LEVEL deklaracijom MSForms
' kontrole / WithEvents (modOtkupBlok, modKarticaDetalji, modPodesavanja,
' modMouseWheel, clsWheelList, clsUiSink) NE mogu kroz AddFromString: dodavanje
' tih deklaracija bind-uje MSForms tip-biblioteku u toku COM edita pa diskonektuje
' CodeModule ([-2147417848]). Ti moduli se PREPOZNAJU UNAPRED (IsHardModuleBody:
' module-level "WithEvents" ili "As MSForms.") i idu PRAVO u fazu 2 (Remove +
' OnTime flush + Import) - AddFromString se nad njima NIKAD ne poziva.
'
' ATOMARNOST (naucena zamka): update se snima SAMO pri PUNOM uspehu. Bilo koji
' fatalni ishod (forma se ne moze azurirati, faza 2 Import padne, Save ne uspe,
' download nepotpun) -> NISTA se ne snima. Posto se pre pune uspesnosti nikad ne
' snima, fajl na disku ostaje STARA ISPRAVNA verzija; korisnik zatvara bez
' snimanja i radna verzija je ocuvana. modConfig/APP_VERSION se dize samo kad se
' snimi -> nema tihe "polu-nove" verzije koja bi prestala da nudi update.
'
' STARTUP RECOVERY: ako faza 2 nikad ne opali (sesija prekinuta), na sledecem
' startu RecoverPendingSelfUpdate cisti stale stanje i obavesti (disk je stara
' verzija; NE pokusava da dovrsi fazu 2 nad starim projektom - to bi mesalo
' verzije).
'
' OPT-IN: radi samo ako je REL_FOLDER_ID postavljen (modConfig).
' FAIL-SOFT: svaka greska u proveri samo se loguje, start ide dalje.
' Reuse: modDrive (download/list), modUpdateGate.VersionCompare,
'        modGoogleAuth.ExtractJsonStringGoogle, Stop*Timer, *_Release/Reset.
' Skip pri importu: modSelfUpdate (na stack-u) + modVbaTools (dev tool).
' NAPOMENA: posto je modSelfUpdate u SKIP_MODULES, ispravke SAMOG updatera ne
' stizu self-update-om - klijent ih dobija jednokratno kroz ImportAllVBA / nov
' .xlsm. Tek posle toga self-update opet vazi. (docs/SELF_UPDATE.md)
' ASCII-only modul (vidi CLAUDE.md, sekcija 4 - encoding).
' ============================================================

Private Const MANIFEST_NAME As String = "version.json"
Private Const SKIP_MODULES As String = "modSelfUpdate,modVbaTools"

' Vrati True ako je korisnik POTVRDIO azuriranje (tada pozivalac treba da
' zakaze RunSelfUpdate i prekine ostatak startup-a). Inace False (nastavi start).
Public Function CheckForUpdateOnOpen() As Boolean
    Const SRC As String = "modSelfUpdate.CheckForUpdateOnOpen"
    On Error GoTo EH

    ' Startup watchdog (prekinuta prethodna faza 2) - zvan ODAVDE (isti modul), NE iz
    ' modMain: modMain je updatable, modSelfUpdate frozen (SKIP_MODULES); direktan poziv
    ' NOVOG simbola iz modMain-a obori compile na starom klijentu (zamka #19). Ide PRE
    ' opt-in provere - recovery ne zavisi od toga da li je remote dostupan.
    RecoverPendingSelfUpdate

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

' STARTUP WATCHDOG (zvan iz CheckForUpdateOnOpen, NE iz modMain - compile-safety,
' zamka #19). Ako je faza 2 iz prethodne sesije ostala "pending" (nikad nije opalila
' - Excel zatvoren,
' OnTime otkazan), ocisti stale stanje. Posto se pre pune uspesnosti nikad ne
' snima, fajl na disku je STARA ISPRAVNA verzija -> bezbedno je samo obrisati
' pending marker + temp i obavestiti. NE pokrece se faza 2 nad starim projektom
' (mesao bi staru i novu verziju). Fail-soft.
Public Sub RecoverPendingSelfUpdate()
    On Error Resume Next
    Dim sec As String: sec = P2Section()
    If Len(GetSetting("AgriXSelfUpdate", sec, "pending", "")) = 0 Then Exit Sub

    Dim d As String: d = GetSetting("AgriXSelfUpdate", sec, "dir", "")
    DeleteSetting "AgriXSelfUpdate", sec
    If Len(d) > 0 Then
        Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(d) Then fso.DeleteFolder d, True
    End If
    MsgBox "Prethodni self-update nije zavrsen do kraja." & vbCrLf & _
           "Radna verzija je ocuvana (nije snimljena delimicna verzija)." & vbCrLf & _
           "Mozete ponoviti azuriranje po zelji.", vbExclamation, APP_NAME
End Sub

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

    ' 2) Download svih fajlova iz AgriX_Release u temp folder. ATOMARNOST:
    '    mora stici SVAKI fajl iz listinga (inace bi se snimila polu-nova verzija).
    Dim tempDir As String: tempDir = MakeTempDir()
    Dim expected As Long, note As String
    Dim n As Long: n = DownloadReleaseFiles(tempDir, expected, note)
    If expected = 0 Then
        MsgBox Poruka("SU_PREUZIMANJE_OTKAZANO") & vbCrLf & _
               IIf(Len(note) > 0, note & vbCrLf, "") & _
               Poruka("SU_POKUSAJTE"), vbCritical, APP_NAME
        Exit Sub
    End If
    If n <> expected Then
        MsgBox "Nepotpuno ili NEVERIFIKOVANO preuzimanje release-a (" & n & "/" & expected & " fajlova)." & _
               IIf(Len(note) > 0, vbCrLf & note, "") & vbCrLf & _
               "Azuriranje je PREKINUTO - nista nije menjano, radna verzija je ocuvana." & vbCrLf & _
               "Proverite internet/Drive pristup (i SHA-256: Alt+F8 -> Test_Sha256File) pa ponovite.", _
               vbCritical, APP_NAME
        Exit Sub
    End If

    ' 3-6) Zajednicko jezgro (PrepareRuntime -> code-merge -> faza 2 -> save).
    '      Isti put koristi i RunSelfUpdateDev (lokalni folder umesto Drive-a).
    RunSelfUpdateCore tempDir, n
    Exit Sub
EH:
    RestoreRuntimeAfterSelfUpdate
    LogErr SRC, Err.description
End Sub

' Zajednicko jezgro update-a: uzmi VEC pripremljen (i KOMPLETAN) folder pun
' src-vba fajlova - skinut sa Drive-a (RunSelfUpdate) ILI kopiran iz lokalnog
' git klona (RunSelfUpdateDev) - i odradi ATOMARAN merge. Snima SAMO pri punom
' uspehu; svaki fatalni ishod -> nista se ne snima, radna verzija ocuvana.
Private Sub RunSelfUpdateCore(ByVal tempDir As String, ByVal n As Long)
    Const SRC As String = "modSelfUpdate.RunSelfUpdateCore"
    On Error GoTo EH

    ' 0) NO-OP GUARD: ako delta-skip kaze da NEMA promena, ne diraj NISTA - bez
    '    teardown-a zivog app-a, bez importa, bez save-a. (Ponovljeni update /
    '    vec-azurna verzija; eliminise rizik Remove+Import modMouseWheel-a.)
    If Not AnyUpdatePending(tempDir) Then
        RestoreRuntimeAfterSelfUpdate
        On Error Resume Next                          ' no-op: ne ostavljaj skinut temp folder
        Dim fsoTmp As Object: Set fsoTmp = CreateObject("Scripting.FileSystemObject")
        If fsoTmp.FolderExists(tempDir) Then fsoTmp.DeleteFolder tempDir, True
        On Error GoTo EH
        MsgBox "Nema izmena koda - vec ste na najnovijoj verziji." & vbCrLf & _
               "Nista nije menjano (delta-skip: 0 promena).", vbInformation, APP_NAME
        Exit Sub
    End If

    ' 3) Oslobodi runtime stanje (root cause fix) PRE importa
    PrepareRuntimeForSelfUpdate

    ' 4) Faza 1: code-merge (soft) + delta-skip + pre-rutiranje tvrdih u fazu 2.
    Dim failed As Object: Set failed = CreateObject("Scripting.Dictionary")
    Dim needsReinstall As Boolean
    Dim summary As String
    summary = ImportFromFolder(tempDir, SKIP_MODULES, failed, needsReinstall)

    ' ATOMARNOST #1: forma/sheet se ne moze azurirati kroz self-update -> NISTA se
    ' ne snima; auto-close bez snimanja (disk ostaje stara ispravna verzija).
    If needsReinstall Then
        AbortSelfUpdate "Azuriranje NIJE moguce kroz self-update (forma/sheet zahteva reinstall)." & _
               vbCrLf & summary
        Exit Sub
    End If

    ' 5) Faza 2: tvrdi .bas/.cls (module-level MSForms/WithEvents) i eventualni
    '    genuino pali .bas/.cls. Ukloni postojece (type-guard: samo std/class),
    '    zapamti TACNU listu fajlova, zakazi Import. NE SNIMA se ovde.
    If failed.count > 0 Then
        Dim proj As Object: Set proj = ThisWorkbook.VBProject
        Dim fk As Variant, remC As Object, files As String, sec As String
        sec = P2Section()
        For Each fk In failed.Keys
            On Error Resume Next
            Set remC = Nothing
            Set remC = proj.VBComponents(CStr(fk))
            If Not remC Is Nothing Then
                If remC.Type = 1 Or remC.Type = 2 Then proj.VBComponents.Remove remC
            End If
            On Error GoTo EH
            If Len(files) > 0 Then files = files & ","
            files = files & CStr(failed(fk))         ' tacan filename za fazu 2
        Next fk
        SaveSetting "AgriXSelfUpdate", sec, "dir", tempDir
        SaveSetting "AgriXSelfUpdate", sec, "files", files
        SaveSetting "AgriXSelfUpdate", sec, "n", CStr(n)
        SaveSetting "AgriXSelfUpdate", sec, "pending", "1"   ' watchdog marker
        ' Workbook-qualified (dve otvorene kopije -> pravi workbook hvata proc)
        Application.OnTime Now + TimeSerial(0, 0, 2), QualifiedProc("RunSelfUpdatePhase2")
        Exit Sub                  ' kraj makroa -> Remove se flush-uje -> faza 2
    End If

    ' 6) Sve soft proslo -> ATOMARAN, VERIFIKOVAN save (bez potvrdjenog save = neuspeh)
    If Not SaveWorkbookVerified() Then
        AbortSelfUpdate "Azuriranje NIJE snimljeno (fajl je mozda samo-za-citanje ili zakljucan)."
        Exit Sub
    End If

    RestoreRuntimeAfterSelfUpdate
    MsgBox Poruka("SU_ZAVRSENO_FAJLOVA") & n & vbCrLf & _
           summary & vbCrLf & vbCrLf & _
           "ZATVORITE i ponovo OTVORITE fajl da se promene aktiviraju.", _
           vbInformation, APP_NAME
    Exit Sub
EH:
    AbortSelfUpdate Poruka("SU_GRESKA_AZURIRANJE") & Err.description
End Sub

' Faza 2 (Application.OnTime; posle flush-a Remove-ova iz faze 1). Uvezi (Import)
' TACNO one module koje je faza 1 uklonila (lista fajlova iz Settinga) - ne skenira
' ceo temp (inace bi mogao da uveze skip-listu/dev module ili sirovo prosao fajl).
' Import podnosi module-level MSForms deklaracije. ATOMARNOST: snima SAMO pri
' punom uspehu. Public zbog OnTime.
Public Sub RunSelfUpdatePhase2()
    Const SRC As String = "modSelfUpdate.RunSelfUpdatePhase2"
    On Error GoTo EH

    Dim sec As String: sec = P2Section()
    Dim p2dir As String: p2dir = GetSetting("AgriXSelfUpdate", sec, "dir", "")
    Dim filesCsv As String: filesCsv = GetSetting("AgriXSelfUpdate", sec, "files", "")
    Dim nTxt As String: nTxt = GetSetting("AgriXSelfUpdate", sec, "n", "?")
    ' NB: marker se NE brise ovde na pocetku - brise se tek na uspeh ILI kontrolisan
    ' fatalni izlaz, da crash USRED importa ostavi trag za RecoverPendingSelfUpdate.
    If Len(p2dir) = 0 Then
        ' Faza 1 je (ako je stigla dovde) uklonila tvrde module; bez temp foldera ne
        ' mogu da ih uvezem -> memorijski projekat je mozda polovan. Fail-closed:
        ' auto-close bez snimanja (disk ostaje stara ispravna verzija).
        DeleteSetting "AgriXSelfUpdate", sec
        AbortSelfUpdate "2. faza: nedostaje temp folder (izgubljen state posle faze 1)."
        Exit Sub
    End If

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim skip As String: skip = "," & LCase$(SKIP_MODULES) & ","
    Dim arr() As String, i As Long, fn As String, baseName As String, ext As String
    Dim imported As Long, expected As Long, stillFail As String

    arr = Split(filesCsv, ",")
    For i = LBound(arr) To UBound(arr)
        fn = Trim$(arr(i))
        If Len(fn) > 0 Then
            baseName = fso.GetBaseName(fn)
            ext = LCase$(fso.GetExtensionName(fn))
            If (ext = "bas" Or ext = "cls") And InStr(skip, "," & LCase$(baseName) & ",") = 0 Then
                expected = expected + 1
                ' Posle faze-1 Remove komponenta MORA biti nestala. Ako jos postoji,
                ' Remove nije zavrsio -> FATALNO (ranije se tiho preskakala = mesan
                ' build sa starim modulom + novim ostatkom).
                If ComponentExists(proj, baseName) Then
                    stillFail = stillFail & "  " & fn & " -> stara komponenta jos postoji (Remove nedovrsen)" & vbCrLf
                Else
                    Dim newC As Object
                    On Error Resume Next
                    Set newC = Nothing
                    Err.Clear
                    Set newC = proj.VBComponents.Import(p2dir & "\" & fn)
                    Dim impErr As Long: impErr = Err.Number
                    On Error GoTo 0
                    ' verifikuj VRACENU komponentu (autoritativna; ime iz VB_Name
                    ' atributa fajla) - tacnog imena i tipa (bas=std/1, cls=class/2).
                    If impErr = 0 And ImportedOk(newC, baseName, ext) Then
                        imported = imported + 1
                    Else
                        stillFail = stillFail & "  " & fn & " -> Import nije verifikovan [" & impErr & _
                                    "] (ime='" & SafeCompName(newC) & "')" & vbCrLf
                    End If
                End If
            End If
        End If
    Next i

    ' Faza 2 se zakazuje SAMO kad ima tvrdih modula (failed.count>0 -> filesCsv nije
    ' prazan -> expected>0). expected=0 ovde = izgubljen/pokvaren registry state POSLE
    ' faze-1 Remove-ova -> fatalno (inace bi "imported<>expected" bilo 0<>0=False pa bi
    ' se snimio build BEZ uklonjenih tvrdih modula).
    If expected = 0 Then
        DeleteSetting "AgriXSelfUpdate", sec
        AbortSelfUpdate "2. faza: nema modula za uvoz (izgubljen state) - radna verzija ocuvana."
        Exit Sub
    End If

    ' ATOMARNOST: SVE mora uspeti (imported = expected, bez stillFail). Inace se NE
    ' snima i workbook se auto-zatvara bez snimanja (disk ostaje stara verzija).
    If Len(stillFail) > 0 Or imported <> expected Then
        DeleteSetting "AgriXSelfUpdate", sec
        AbortSelfUpdate "2. faza NIJE uspela za sve module (uvezeno " & imported & "/" & expected & "):" & _
               vbCrLf & stillFail
        Exit Sub
    End If

    If Not SaveWorkbookVerified() Then
        DeleteSetting "AgriXSelfUpdate", sec
        AbortSelfUpdate "2. faza uspela ali SAVE nije (read-only/zakljucan)."
        Exit Sub
    End If

    ' Uspeh -> ocisti marker + temp
    DeleteSetting "AgriXSelfUpdate", sec
    On Error Resume Next
    If fso.FolderExists(p2dir) Then fso.DeleteFolder p2dir, True
    On Error GoTo EH

    RestoreRuntimeAfterSelfUpdate
    MsgBox Poruka("SU_ZAVRSENO_PREUZETO") & nTxt & ", 2. faza uvezeno: " & imported & vbCrLf & vbCrLf & _
           "ZATVORITE i ponovo OTVORITE fajl da se promene aktiviraju.", _
           vbInformation, APP_NAME
    Exit Sub
EH:
    DeleteSetting "AgriXSelfUpdate", P2Section()
    AbortSelfUpdate Poruka("SU_GRESKA_2FAZA") & Err.description
End Sub

' Oslobodi runtime stanje pre self-update importa: ugasi evente/sync/tajmere,
' otpusti module-level reference dinamickih kontrola/WithEvents (inace CodeModule
' edit tih modula diskonektuje COM), unload sve forme. Best-effort.
Private Sub PrepareRuntimeForSelfUpdate()
    On Error Resume Next

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Otkazi SVE poznate Application.OnTime tikove. Tik koji opali usred importa
    ' (ili u prozoru izmedju faze 1 i 2, dok su "tvrdi" moduli uklonjeni) forsira
    ' demand-compile polomljenog projekta; AutoSaveTick/StornoWarm bi uz to jos i
    ' pokrenuli izmenu preko radne kopije.
    StopScheduledSync               ' modGoogleSyncOrchestrator (pending OnTime sync)
    StopAutoSaveTimer               ' modJournaling (pending AutoSaveTick)
    StopHeartbeatTimer              ' modStanicaLock (90s heartbeat)
    StopStornoWarm                  ' modStornoWarm (storno browse warm cache tick)

    ' Release module-level WithEvents/kontrole (dinamicki paneli)
    OtkupBlok_Release               ' modOtkupBlok (clsBlokUI)
    Podesavanja_Release             ' modPodesavanja (clsConfigBtn)
    MaticniMenu_Release             ' modMaticniLookups (clsLookupMenuBtn)
    KarticaDetalji_Reset            ' modKarticaDetalji
    MouseWheel_Off                  ' modMouseWheel (skini LL mouse hook pre izmene koda)

    ' Unload sve forme (Terminate svake otpusti svoje clsUiSink sink-ove i kontrole)
    Do While VBA.UserForms.count > 0
        Unload VBA.UserForms(0)
    Loop

    DoEvents
    On Error GoTo 0
End Sub

' Vrati aplikaciona podesavanja koja PrepareRuntimeForSelfUpdate gasi. MORA se
' pozvati na SVAKOM izlazu update toka (uspeh, greska, prekid) - inace
' Workbook_Open ne opali pri sledecem otvaranju fajla u ISTOJ Excel instanci.
Private Sub RestoreRuntimeAfterSelfUpdate()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Snimi radnu svesku i VERIFIKUJ da je stvarno snimljena. True samo ako Save nije
' bacio gresku I ThisWorkbook.Saved = True (read-only / zakljucan / pun disk /
' Drive-sync blok -> False). Bez ovoga bi neuspeo Save izgledao kao uspeh.
Private Function SaveWorkbookVerified() As Boolean
    On Error Resume Next
    Err.Clear
    ThisWorkbook.Save
    SaveWorkbookVerified = (Err.Number = 0) And ThisWorkbook.Saved
    On Error GoTo 0
End Function

' Fatalni izlaz update-a: NISTA se ne snima + workbook se AUTO-ZATVARA bez
' snimanja (tehnicka garancija atomarnosti - ne oslanja se na to da operater
' nece pritisnuti Ctrl+S / da OneDrive AutoSave nece upisati polu-nov projekat).
' Disk ostaje stara ISPRAVNA verzija. Poziva se sa fatalnih grana POSLE
' PrepareRuntime (kad je projekat vec izmenjen u memoriji).
Private Sub AbortSelfUpdate(ByVal msg As String)
    On Error Resume Next
    RestoreRuntimeAfterSelfUpdate
    ThisWorkbook.Saved = True          ' ODMAH: ni AutoSave ni close-prompt ne upisu polu-nov projekat
    LogErr "modSelfUpdate.AbortSelfUpdate", msg
    MsgBox msg & vbCrLf & vbCrLf & _
           "NISTA nije snimljeno - radna verzija je ocuvana." & vbCrLf & _
           "Program ce se sada ZATVORITI bez snimanja. Otvorite ga ponovo; po zelji " & _
           "ponovite azuriranje ili vratite Backup\AgriX_pre-update_*.xlsm.", _
           vbCritical, APP_NAME
    ' Zatvori DIREKTNO iz tekuceg (OnTime-dispatch) stack-a - NE preko novog OnTime
    ' makroa: polomljen/nekompajlabilan projekat mozda ne bi razresio novi name-dispatch
    ' (demand-compile bi pao) pa auto-close nikad ne bi opalio. Direktan poziv radi jer
    ' je modSelfUpdate vec na stack-u (ne trazi re-compile celog projekta).
    AbortSelfUpdateClose
End Sub

' Auto-close bez snimanja. Zove se DIREKTNO iz AbortSelfUpdate (ne vise preko
' Application.OnTime - polomljen projekat ne bi razresio name-dispatch); Public
' zadrzan (istorijski OnTime cilj / rucni fallback). Saved=True pre Close -> ni
' ShutdownApp/FlushNow (koji gledaju .Saved) ni Close ne upisu polu-nov memorijski
' projekat preko stare (dobre) kopije na disku.
Public Sub AbortSelfUpdateClose()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    ThisWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub

' Workbook-qualified naziv procedure za Application.OnTime ("'Ime.xlsm'!Proc").
' Da se, kad su dve kopije AgriX-a otvorene (narocito DEV test na kopiji), OnTime
' razresi u PRAVOM workbooku, a ne u proizvoljnom.
Private Function QualifiedProc(ByVal procName As String) As String
    QualifiedProc = "'" & Replace$(ThisWorkbook.name, "'", "''") & "'!" & procName
End Function

' Sekcija u registru (SaveSetting) scope-ovana po workbook-u -> dve otvorene kopije
' ne dele phase2 stanje (inace bi RecoverPendingSelfUpdate jedne obrisao pending
' druge). Ime sanitizovano na alfanumerik.
Private Function P2Section() As String
    Dim s As String, i As Long, ch As String, out As String
    s = ThisWorkbook.name
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Then out = out & ch
    Next i
    P2Section = "phase2_" & out
End Function

' True ako komponenta datog imena postoji u projektu.
Private Function ComponentExists(ByVal proj As Object, ByVal baseName As String) As Boolean
    On Error Resume Next
    Dim c As Object
    Set c = proj.VBComponents(baseName)
    ComponentExists = Not (c Is Nothing)
End Function

' Verifikuj VRACENU komponentu iz Import-a (autoritativna - ne re-lookup po imenu):
' nije Nothing, tacnog imena (VB_Name), i ocekivanog tipa (bas -> std (1),
' cls -> class (2)). Fail-closed verifikacija faze 2. Hvata i .bas/.cls bez
' Attribute VB_Name (Import ga tada nazove pogresno -> ime != baseName).
Private Function ImportedOk(ByVal c As Object, ByVal baseName As String, ByVal ext As String) As Boolean
    On Error Resume Next
    If c Is Nothing Then Exit Function
    If StrComp(c.name, baseName, vbTextCompare) <> 0 Then Exit Function
    If ext = "bas" Then
        ImportedOk = (c.Type = 1)
    Else
        ImportedOk = (c.Type = 2)
    End If
End Function

' Ime komponente za dijagnostiku (ili "?" ako Nothing). Da izvestaj pokaze KAKO
' je Import pogresno nazvao modul (npr. "Module1" kad fali Attribute VB_Name).
Private Function SafeCompName(ByVal c As Object) As String
    On Error Resume Next
    If c Is Nothing Then SafeCompName = "?" Else SafeCompName = c.name
End Function

' ---------------- private ----------------

' Procitaj app_version. F2: prvo current.json (versioned pokazivac); fallback
' version.json (flat/legacy). "" ako nedostupno (offline -> fail-soft).
Private Function GetRemoteAppVersion() As String
    Const SRC As String = "modSelfUpdate.GetRemoteAppVersion"
    On Error GoTo EH

    Dim cur As String: cur = DownloadNamedText(REL_FOLDER_ID, "current.json")
    If Len(cur) > 0 Then
        GetRemoteAppVersion = Trim$(ExtractJsonStringGoogle(cur, "app_version"))
        If Len(GetRemoteAppVersion) > 0 Then Exit Function
    End If
    GetRemoteAppVersion = Trim$(ExtractJsonStringGoogle(DownloadNamedText(REL_FOLDER_ID, MANIFEST_NAME), "app_version"))
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

' Napravi JEDINSTVEN prazan temp folder za preuzimanje. Jedinstven (timestamp) da
' drugi workbook / drugi update ne obrise folder koji ceka fazu 2.
Private Function MakeTempDir() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base As String: base = fso.GetSpecialFolder(2) & "\AgriX_update_" & Format$(Now, "yyyymmdd_hhnnss")
    Dim d As String: d = base
    Dim k As Long
    Do While fso.FolderExists(d)
        k = k + 1
        d = base & "_" & k
    Loop
    fso.CreateFolder d
    MakeTempDir = d
End Function

' Skini sve kod-fajlove iz AgriX_Release u tempDir. Vrati broj USPESNO skinutih,
' a kroz expectedOut broj koji je TREBALO da stigne (podrzane ekstenzije iz
' listinga). Pozivalac prekida update ako downloaded <> expected (atomarnost).
' Skini kod-fajlove iz AgriX_Release u tempDir. Vrati broj USPESNO skinutih (i
' VERIFIKOVANIH), expectedOut = broj koji je TREBALO. Pozivalac prekida ako
' downloaded <> expected -> nesklad hesa / nedostajuci fajl = fatalno (atomarno).
'  - Ako manifest ima files[] (F1+): download SAMO tih fajlova + verifikacija
'    SHA-256 (fallback na prisustvo/broj ako SHA nedostupan na masini).
'  - Ako nema files[] (stari publisher): legacy listing-download (kao pre F1).
Private Function DownloadReleaseFiles(ByVal tempDir As String, ByRef expectedOut As Long, ByRef note As String) As Long
    Const SRC As String = "modSelfUpdate.DownloadReleaseFiles"
    On Error GoTo EH

    expectedOut = 0
    note = ""

    ' F2: odredi IZVOR (versioned preko current.json + provera manifest_sha256,
    ' inace flat/legacy fallback). manifest_sha256 nesklad = fail-closed -> False
    ' -> expected ostaje 0 -> pozivalac javi prekid (nista se ne skida).
    Dim srcFolder As String, manifestText As String
    If Not ResolveReleaseSource(tempDir, srcFolder, manifestText, note) Then Exit Function

    Dim dict As Object: Set dict = DriveListFolder(srcFolder)
    If dict Is Nothing Then Exit Function

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim files As Object: Set files = ParseManifestFiles(manifestText)
    Dim n As Long, expected As Long

    If files.count = 0 Then
        ' LEGACY (stari publisher bez files[]) -> listing-driven, bez hash provere.
        AppendNote note, "legacy manifest (bez files[]) - bez SHA verifikacije"
        Dim k As Variant, ext As String
        For Each k In dict.Keys
            ext = LCase$(fso.GetExtensionName(CStr(k)))
            Select Case ext
                Case "bas", "cls", "frm", "frx", "doccls"
                    expected = expected + 1
                    If DriveDownloadToFile(CStr(dict(k)), tempDir & "\" & CStr(k)) Then n = n + 1
            End Select
        Next k
    Else
        ' MANIFEST-DRIVEN + SHA-256 verifikacija (F1). Manifest je merodavan izvor
        ' liste; stale fajlovi iz foldera se ignorisu.
        Dim shaOk As Boolean: shaOk = Sha256Available(tempDir)
        If Not shaOk Then AppendNote note, "SHA-256 nedostupan na ovoj masini - fallback na prisustvo/broj (bez hesa)"
        Dim nm As Variant, dest As String, wantSha As String
        For Each nm In files.Keys
            expected = expected + 1
            wantSha = CStr(files(nm))
            If Not dict.Exists(CStr(nm)) Then
                LogErr SRC, "manifest fajl nije na Drive-u: " & CStr(nm)     ' -> n<expected -> fatalno
            ElseIf shaOk And Not IsSha256Hex(wantSha) Then
                ' Hashed manifest na SHA-sposobnoj masini MORA imati validan 64-hex sha
                ' po fajlu; prazan/nevalidan sha = pokvaren/podmetnut manifest -> ne broj
                ' (fail-closed; n<expected -> pozivalac prekida update).
                LogErr SRC, "manifest bez validnog sha256 (64 hex): " & CStr(nm)
            ElseIf DriveDownloadToFile(CStr(dict(CStr(nm))), tempDir & "\" & CStr(nm)) Then
                dest = tempDir & "\" & CStr(nm)
                If shaOk Then
                    If LCase$(Sha256File(dest)) = LCase$(wantSha) Then
                        n = n + 1
                    Else
                        LogErr SRC, "SHA-256 NESKLAD (korupcija/stale): " & CStr(nm)   ' -> n<expected -> fatalno
                    End If
                Else
                    n = n + 1        ' SHA nedostupan NA KLIJENTU -> prisustvo/broj (dokumentovano)
                End If
            End If
        Next nm
    End If

    expectedOut = expected
    DownloadReleaseFiles = n
    Exit Function
EH:
    LogErr SRC, Err.description
End Function

' Skini imenovani tekstualni fajl iz datog Drive foldera -> sadrzaj (ili "").
' Generalizacija (F2): current.json / version.json (poznata tacka = REL_FOLDER_ID);
' manifest.json iz versioned foldera ide preko ResolveReleaseSource. Razdvojen
' temp po imenu da current.json i version.json ne gaze isti fajl.
Private Function DownloadNamedText(ByVal folderID As String, ByVal fileName As String) As String
    On Error Resume Next
    Dim id As String: id = DriveFindInFolder(folderID, fileName)
    If Len(id) = 0 Then Exit Function
    Dim tmp As String: tmp = Environ$("TEMP") & "\AgriX_dl_" & fileName
    If DriveDownloadToFile(id, tmp) Then DownloadNamedText = ReadAllText(tmp)
End Function

' F2: odredi IZVOR release-a (folder + tekst manifesta). Lanac poverenja:
'   current.json (u REL_FOLDER_ID - poznata, stabilna tacka) -> release_folder_id
'   + manifest_sha256. manifest.json iz versioned foldera se skine i NJEGOV
'   SHA-256 proveri protiv manifest_sha256 iz current.json PRE ijednog download-a
'   koda. Nesklad = fail-closed (False, prekid - moguca korupcija/podvala).
' Fallback (nema current.json / nema release_folder_id) = FLAT/legacy:
'   folder = REL_FOLDER_ID, manifest = version.json (kao pre F2).
' True + folderOut + manifestOut (+ note). False = fatalno; note = razlog.
Private Function ResolveReleaseSource(ByVal tempDir As String, ByRef folderOut As String, _
                                      ByRef manifestOut As String, ByRef note As String) As Boolean
    Const SRC As String = "modSelfUpdate.ResolveReleaseSource"
    On Error GoTo EH

    Dim cur As String: cur = DownloadNamedText(REL_FOLDER_ID, "current.json")
    Dim vfid As String
    If Len(cur) > 0 Then vfid = Trim$(ExtractJsonStringGoogle(cur, "release_folder_id"))

    If Len(vfid) = 0 Then
        ' FLAT/legacy fallback - nema versioned layout-a.
        folderOut = REL_FOLDER_ID
        manifestOut = DownloadNamedText(REL_FOLDER_ID, MANIFEST_NAME)
        ResolveReleaseSource = True
        Exit Function
    End If

    ' VERSIONED (F2). Skini manifest.json iz versioned foldera pa proveri njegov
    ' SHA-256 protiv manifest_sha256 iz current.json (koren poverenja) PRE koda.
    Dim mId As String: mId = DriveFindInFolder(vfid, "manifest.json")
    If Len(mId) = 0 Then
        note = "F2: manifest.json nije nadjen u versioned folderu (prekid)"
        Exit Function
    End If
    Dim mPath As String: mPath = tempDir & "\_manifest.json"
    If Not DriveDownloadToFile(mId, mPath) Then
        note = "F2: manifest.json se ne moze skinuti (prekid)"
        Exit Function
    End If

    Dim wantSha As String: wantSha = Trim$(ExtractJsonStringGoogle(cur, "manifest_sha256"))
    If Len(wantSha) > 0 Then
        If Sha256Available(tempDir) Then
            If LCase$(Sha256File(mPath)) <> LCase$(wantSha) Then
                note = "F2: manifest_sha256 NESKLAD - moguca korupcija/podvala (fail-closed)"
                Exit Function
            End If
        Else
            AppendNote note, "F2: SHA-256 nedostupan - manifest_sha256 nije proveren"
        End If
    Else
        AppendNote note, "F2: current.json bez manifest_sha256 (manifest kriptografski nevezan)"
    End If

    folderOut = vfid
    manifestOut = ReadAllText(mPath)
    ResolveReleaseSource = True
    Exit Function
EH:
    note = "ResolveReleaseSource greska: " & Err.description
    LogErr SRC, Err.description
End Function

' Dodaj liniju napomene (visestruke se gomilaju, ne gaze prethodne).
Private Sub AppendNote(ByRef note As String, ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    If Len(note) > 0 Then note = note & vbCrLf & s Else note = s
End Sub

' Parsiraj files[] iz manifesta -> Dictionary(ime -> sha256hex). Prazno ako nema
' files[] (stari publisher). size se ne parsira (broj; sha256 subsumira duzinu).
' Reuse ExtractJsonStringGoogle + split-po-"}" (isti obrazac kao DriveListFolder).
Private Function ParseManifestFiles(ByVal json As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                    ' TextCompare
    Set ParseManifestFiles = d
    If Len(json) = 0 Then Exit Function

    Dim p As Long: p = InStr(1, json, """files""", vbTextCompare)
    If p = 0 Then Exit Function           ' nema files[] -> legacy

    Dim parts() As String, i As Long, nm As String, sh As String
    parts = Split(Mid$(json, p), "}")
    For i = LBound(parts) To UBound(parts)
        nm = ExtractJsonStringGoogle(parts(i), "name")
        If Len(nm) > 0 Then
            sh = ExtractJsonStringGoogle(parts(i), "sha256")
            If Not d.Exists(nm) Then d.Add nm, sh
        End If
    Next i
End Function

' Da li SHA-256 radi na ovoj masini (testira BAS Sha256File put nad "abc" fajlom).
Private Function Sha256Available(ByVal workDir As String) As Boolean
    On Error Resume Next
    Dim p As String: p = workDir & "\_shatest.tmp"
    Dim ff As Integer: ff = FreeFile
    Open p For Output As #ff
    Print #ff, "abc";                    ' tacno 3 bajta
    Close #ff
    Sha256Available = (LCase$(Sha256File(p)) = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
    Kill p
End Function

' True ako je s tacno 64 hex znaka (validan SHA-256 otisak). Prazan/kraci/duzi/
' ne-hex = nevalidan -> manifest entry se ne prihvata na SHA-sposobnoj masini.
Private Function IsSha256Hex(ByVal s As String) As Boolean
    If Len(s) <> 64 Then Exit Function
    Dim i As Long, c As Long
    For i = 1 To 64
        c = AscW(Mid$(s, i, 1))
        If Not ((c >= 48 And c <= 57) Or (c >= 65 And c <= 70) Or (c >= 97 And c <= 102)) Then Exit Function
    Next i
    IsSha256Hex = True
End Function

' Faza 1: delta-skip + pre-rutiranje tvrdih + soft code-merge. Radi za soft module
' jer je runtime oslobodjen (PrepareRuntime) pa CodeModule edit ne diskonektuje.
'  - identican kod (SameCode)        -> "same", NE diraj
'  - tvrd .bas/.cls (WithEvents/MSForms module-level) -> "phase2" (failedOut), kod
'    se NE dira; Remove+Import ide u fazi 2 (AddFromString ga ne bi podneo)
'  - obican .bas/.cls                 -> DeleteLines+AddFromString (ili Add za nov)
'  - forma/doccls                     -> code-merge; NIKAD Remove; pad -> rollback
'    starog koda + needsReinstall (fatalno; pozivalac ne snima)
'  - nova forma/sheet                 -> "skip" + needsReinstall (reinstall)
' Genuino pali .bas/.cls -> failedOut (recovery kroz fazu 2). failedOut: baseName
' -> filename (tacna lista za fazu 2). Retry 3 prolaza + per-modul greska safety.
Private Function ImportFromFolder(ByVal folder As String, ByVal skipCsv As String, _
                                  ByRef failedOut As Object, ByRef needsReinstall As Boolean) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim skip As String: skip = "," & LCase$(skipCsv) & ","
    Dim st As Object: Set st = CreateObject("Scripting.Dictionary")   ' fkey -> ok/same/phase2/skip
    Dim er As Object: Set er = CreateObject("Scripting.Dictionary")   ' fkey -> genuine greska
    Dim orig As Object: Set orig = CreateObject("Scripting.Dictionary") ' fkey -> stari kod (rollback formi)
    Dim pass As Long, anyLeft As Boolean, usedPass As Long
    Dim fil As Object, ext As String, baseName As String, fkey As String
    Dim body As String, cur As String, readOk As Boolean, extractOk As Boolean, compExists As Boolean
    Dim vbc As Object, proj As Object
    Dim okN As Long, sameN As Long, phase2N As Long, k As Variant, failS As String, skipS As String, hardS As String

    needsReinstall = False

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
                    extractOk = (Err.Number = 0 And (Len(body) > 0 Or ext = "doccls"))
                    If Not extractOk Then
                        ' ekstrakcija pala / prazno telo za ne-sheet -> ne diraj komponentu
                        If Err.Number = 0 Then Err.Raise vbObjectError + 2801, , "prazno telo (" & fil.name & ")"
                    Else
                        Set vbc = Nothing
                        Set vbc = proj.VBComponents(baseName)
                        compExists = Not (vbc Is Nothing)
                        Err.Clear                         ' Err9 od lookup-a ne boji dalje

                        cur = ""
                        readOk = True
                        If compExists Then
                            If vbc.CodeModule.CountOfLines > 0 Then _
                                cur = vbc.CodeModule.Lines(1, vbc.CodeModule.CountOfLines)
                            readOk = (Err.Number = 0)
                            Err.Clear
                        End If

                        If compExists And readOk And SameCode(cur, body) Then
                            st(fkey) = "same"             ' delta-skip: identican -> ne diraj
                        ElseIf (ext = "bas" Or ext = "cls") And IsHardModuleBody(body) Then
                            ' TVRD modul -> NIKAD AddFromString; pravo u fazu 2 (Remove+Import).
                            st(fkey) = "phase2"
                            If Not failedOut Is Nothing Then failedOut(baseName) = fil.name
                        ElseIf (ext = "frm" Or ext = "doccls") And _
                               (IsHardModuleBody(body) Or (compExists And readOk And IsHardModuleBody(cur))) Then
                            ' Forma/sheet sa MODULE-LEVEL WithEvents/MSForms se NE sme menjati
                            ' kroz AddFromString (diskonektuje CodeModule; zamka #3) NITI
                            ' Remove+Import u runtime-u (zamka #1). Ni novo telo tvrdo, ni
                            ' zatecena forma tvrda (zamrznute forme) -> jedini bezbedan ishod
                            ' = reinstall (fail-closed; NE polu-merge -> NE crash na wire-up-u).
                            st(fkey) = "skip"
                            needsReinstall = True
                        ElseIf compExists Then
                            If readOk And Not orig.Exists(fkey) Then orig(fkey) = cur
                            If vbc.CodeModule.CountOfLines > 0 Then _
                                vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
                            If Len(body) > 0 Then vbc.CodeModule.AddFromString body
                            If Err.Number = 0 Then st(fkey) = "ok"
                        ElseIf ext = "bas" Or ext = "cls" Then
                            Err.Clear                     ' Err9 lookup ne sme da oboji Add put
                            Set vbc = proj.VBComponents.Add(IIf(ext = "bas", 1, 2))
                            vbc.name = baseName
                            If Len(body) > 0 Then vbc.CodeModule.AddFromString body
                            If Err.Number = 0 Then st(fkey) = "ok"
                        Else
                            st(fkey) = "skip"             ' nova forma/sheet -> reinstall
                            needsReinstall = True
                        End If
                    End If
                    If Err.Number <> 0 And Not st.Exists(fkey) Then
                        er(fkey) = "[" & Err.Number & "] " & Err.description
                        anyLeft = True
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        Next fil
        If Not anyLeft Then Exit For
    Next pass

    For Each k In st.Keys
        Select Case st(k)
            Case "ok":     okN = okN + 1
            Case "same":   sameN = sameN + 1
            Case "phase2": phase2N = phase2N + 1: hardS = hardS & "  " & k & vbCrLf
            Case Else:     skipS = skipS & "  " & k & vbCrLf     ' skip (nova forma/sheet)
        End Select
    Next k
    For Each k In er.Keys
        If Not st.Exists(CStr(k)) Then
            ext = LCase$(fso.GetExtensionName(CStr(k)))
            If ext = "bas" Or ext = "cls" Then
                ' genuino pao .bas/.cls -> pokusaj fazu 2 (Remove+Import podnosi vise)
                If Not failedOut Is Nothing Then failedOut(fso.GetBaseName(CStr(k))) = CStr(k)
                failS = failS & "  " & k & " -> " & er(k) & " (faza 2)" & vbCrLf
            Else
                ' forma/sheet pala -> vrati stari kod + reinstall (fatalno za save)
                RestoreComponentCode orig, CStr(k)
                needsReinstall = True
                failS = failS & "  " & k & " -> " & er(k) & " (forma - vracen stari kod, reinstall)" & vbCrLf
            End If
        End If
    Next k

    ImportFromFolder = "Azurirano: " & okN & ", bez izmene: " & sameN & _
        ", faza 2 (tvrdi): " & phase2N & " (prolaza: " & usedPass & ")" & _
        IIf(Len(hardS) > 0, vbCrLf & "Faza 2 (Remove+Import):" & vbCrLf & hardS, "") & _
        IIf(Len(skipS) > 0, vbCrLf & "Preskoceno (novo, reinstall):" & vbCrLf & skipS, "") & _
        IIf(Len(failS) > 0, vbCrLf & "GRESKE:" & vbCrLf & failS, "")
End Function

' Da li telo (izvuceno ExtractModuleCode) ima MODULE-LEVEL "WithEvents" ili
' "As MSForms." deklaraciju (pre prve procedure). Takvi .bas/.cls NE smeju kroz
' AddFromString (diskonektuje CodeModule) -> pravo u fazu 2. Zamena za hardkodiranu
' listu: detekcija radi i za nove module (npr. clsUiSink).
Private Function IsHardModuleBody(ByVal body As String) As Boolean
    Dim arr() As String, i As Long, u As String, w As String
    ' Normalizuj SVE prekide reda (vbCrLf iz fajla; lone vbCr iz CodeModule.Lines) ->
    ' vbLf, da detekcija radi i nad telom iz fajla i nad kodom procitanim iz projekta.
    Dim t As String: t = Replace$(Replace$(body, vbCrLf, vbLf), vbCr, vbLf)
    arr = Split(t, vbLf)
    For i = 0 To UBound(arr)
        u = CodeLineUpper(arr(i))          ' bez stringa/komentara (da "WithEvents"
        If Len(u) > 0 Then                 ' u komentaru ne da lazni pozitiv)
            ' kraj module-level dela = prva procedura
            w = u
            If Left$(w, 7) = "PUBLIC " Then w = Mid$(w, 8)
            If Left$(w, 8) = "PRIVATE " Then w = Mid$(w, 9)
            If Left$(w, 7) = "FRIEND " Then w = Mid$(w, 8)
            If Left$(w, 7) = "STATIC " Then w = Mid$(w, 8)
            If w Like "SUB *" Or w Like "FUNCTION *" Or w Like "PROPERTY *" Then Exit For
            If InStr(1, u, "WITHEVENTS ") > 0 Then IsHardModuleBody = True: Exit Function
            If InStr(1, u, " AS MSFORMS.") > 0 Then IsHardModuleBody = True: Exit Function
        End If
    Next i
End Function

' Kodni deo linije (bez trailing komentara i sadrzaja stringova ne dira - samo
' odseca komentar posle apostrofa koji NIJE u stringu), UPPER + LTrim. Da rec
' "WithEvents"/"MSForms." u KOMENTARU ne da lazni "tvrd modul".
Private Function CodeLineUpper(ByVal s As String) As String
    Dim i As Long, ch As String, inQ As Boolean, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then inQ = Not inQ
        If ch = "'" And Not inQ Then Exit For
        out = out & ch
    Next i
    CodeLineUpper = UCase$(LTrim$(out))
End Function

' Da li su dva tela koda identicna. Ignorise SAMO zavrsne CR/LF; ostalo mora biti
' bajt-za-bajt isto (izvori su ASCII, exporti iz istog VBE).
Private Function SameCode(ByVal a As String, ByVal b As String) As Boolean
    ' BINARNO poredjenje kanonizovanih tela. CanonCode lowercase-uje SAMO kod IZVAN
    ' stringova/komentara -> apsorbuje VBE re-casing identifikatora (VBE unifikuje
    ' casing istoimenih identifikatora projekt-wide; potvrdjeno AscW proj=120 'x' vs
    ' file=88 'X'), ali CUVA case unutar string-literala i komentara. Tako izmena
    ' SAMO case-a u stringu (npr. "DA" -> "da": JSON kljuc, ID prefiks, API vrednost)
    ' i DALJE menja sadrzaj i detektuje se (vbTextCompare bi je slepo progutao).
    SameCode = (StrComp(CanonCode(a), CanonCode(b), vbBinaryCompare) = 0)
End Function

' Kanonizuj kod za poredjenje: SVE vrste prekida reda (CRLF/CR/LF) -> LF, pa skini
' zavrsne prazne redove. KRITICNO: VBE CodeModule.Lines vraca redove razdvojene
' sa vbCr, a ExtractModuleCode gradi telo sa vbCrLf. Bez ove normalizacije
' SameCode uvek vraca False (svaki red se razlikuje po separatoru) -> delta-skip
' NIKAD ne prijavi "isto" -> svaki update nepotrebno re-merge/re-import-uje SVE
' komponente, uklj. rizican modMouseWheel Remove+Import (AddressOf hook) -> crash
' pri ponovljenom update-u nad vec-azurnim, punim app-om.
Private Function CanonCode(ByVal s As String) As String
    ' 1) sve vrste prekida reda -> LF
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    ' 1b) VBE CodeModule.Lines ume da vrati NON-BREAKING SPACE (ChrW 160) tamo gde
    '     fajl ima obican space (0x20) -> ista duzina, nevidljiva razlika (uzrok
    '     preostalih ~64 laznih "razlicito" posle vodeci-blank fixa). Normalizuj.
    s = Replace$(s, ChrW$(160), " ")
    ' 2) RTrim svaki red + lowercase kod IZVAN stringova/komentara (LowerOutsideStrings):
    '    apsorbuje VBE identifier re-casing bez gubljenja case-a u stringovima/komentarima
    '    (SameCode posle poredi BINARNO). Trailing whitespace nebitan (hvata "ista
    '    duzina ali razlicit" slucaj: red sa/bez zavrsnog space-a, npr. clsWheelList).
    Dim arr() As String, i As Long
    arr = Split(s, vbLf)
    For i = LBound(arr) To UBound(arr)
        arr(i) = RTrim$(LowerOutsideStrings(arr(i)))
    Next i
    s = Join(arr, vbLf)
    ' 3) skini VODECE i zavrsne prazne redove - VBE CodeModule.Lines cesto vraca
    '    vodeci prazan red koga ExtractModuleCode (iz fajla) nema (uzrok #78 lazno
    '    "razlicito" -> nepotreban re-import -> crash na modMouseWheel Remove).
    Do While Len(s) > 0
        If Left$(s, 1) = vbLf Then s = Mid$(s, 2) Else Exit Do
    Loop
    Do While Len(s) > 0
        If Right$(s, 1) = vbLf Then s = Left$(s, Len(s) - 1) Else Exit Do
    Loop
    CanonCode = s
End Function

' Lowercase kod IZVAN string-literala i komentara; case unutar "..." i posle ' se
' CUVA. Koristi SameCode da absorbuje VBE identifier re-casing (identifikatori se
' lowercase-uju na obe strane -> match) a da NE proguta case-only izmenu u stringu.
' VBA string-literal nikad ne prelazi fizicki red (line-continuation je van navodnika),
' pa je per-red obrada tacna. "" unutar stringa = escaped navodnik (ostaje u stringu).
Private Function LowerOutsideStrings(ByVal s As String) As String
    Dim i As Long, n As Long, c As String, out As String, inStr As Boolean
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If inStr Then
            If c = """" Then
                If i < n And Mid$(s, i + 1, 1) = """" Then
                    out = out & """"""            ' escaped "" -> ostaje u stringu
                    i = i + 2
                Else
                    inStr = False
                    out = out & c
                    i = i + 1
                End If
            Else
                out = out & c                     ' cuvaj case unutar stringa
                i = i + 1
            End If
        Else
            If c = """" Then
                inStr = True
                out = out & c
                i = i + 1
            ElseIf c = "'" Then
                out = out & Mid$(s, i)            ' komentar do kraja reda -> cuvaj
                Exit Do
            Else
                out = out & LCase$(c)             ' kod van stringa/komentara -> lower
                i = i + 1
            End If
        End If
    Loop
    LowerOutsideStrings = out
End Function

' READ-ONLY: da li ima ista da se azurira (nova komponenta ILI izmenjen kod).
' Pozvati PRE PrepareRuntime -> ako vrati False, update NE dira nista (bez
' teardown-a zivog app-a, bez importa, bez save-a). Ponovljeni / vec-azuran
' update = pravi no-op -> nema Remove+Import (pa ni crash na modMouseWheel).
Private Function AnyUpdatePending(ByVal folder As String) As Boolean
    ' FAIL-SAFE: bilo koja greska u proveri -> tretiraj kao "ima izmena" (True) ->
    ' pun atomarni update put (koji je i sam fail-closed). NIKAD ne prijavljuj lazni
    ' no-op ("vec ste azurni") na gresku - to bi tiho preskocilo stvaran update.
    On Error GoTo EH_SAFE
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim skip As String: skip = "," & LCase$(SKIP_MODULES) & ","
    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim fil As Object, ext As String, baseName As String, body As String, cur As String
    Dim vbc As Object
    For Each fil In fso.GetFolder(folder).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "doccls" Then
            baseName = fso.GetBaseName(fil.name)
            If InStr(skip, "," & LCase$(baseName) & ",") = 0 Then
                If Not ComponentExists(proj, baseName) Then
                    AnyUpdatePending = True: Exit Function        ' nova komponenta
                End If
                Set vbc = proj.VBComponents(baseName)
                body = ExtractModuleCode(fil.path)
                cur = ""
                If vbc.CodeModule.CountOfLines > 0 Then cur = vbc.CodeModule.Lines(1, vbc.CodeModule.CountOfLines)
                If Not SameCode(cur, body) Then
                    AnyUpdatePending = True: Exit Function     ' izmenjen kod
                End If
            End If
        End If
    Next fil
    Exit Function
EH_SAFE:
    AnyUpdatePending = True    ' nesigurni smo -> ne preskaci; pun put je atoman/fail-closed
End Function

' Best-effort rollback koda JEDNE forme/sheet komponente na sadrzaj pre merge
' pokusaja (orig snimljen pre prvog DeleteLines). Tiho (Resume Next).
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

' Izvuci editabilni kod iz izvoznog VBA fajla (.bas/.cls/.frm/.doccls) bezbedan za
' AddFromString: preskoci header (VERSION, Begin..End dizajn blok, module Attribute
' linije) i strip-uj sve "Attribute ..." linije u kodu.
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

' ============================================================
' DEV TEST (RUCNO, Alt+F8): RunSelfUpdateDev
'
' Najlaksi nacin da se self-update ENGINE testira na svojoj masini bez Drive-a:
' code-merge iz LOKALNOG src-vba foldera (git klon), kroz ISTI RunSelfUpdateCore
' kao pravi self-update (atomarno). Testira ono sto ImportAllVBA NE testira - bas
' code-merge put (DeleteLines/AddFromString + delta-skip + hard->faza2), gde su
' forme pucale.
'
' GUARD (dev-only, fail-closed na klijentu): pokrece se SAMO ako je izabran folder
' pravi git klon "src-vba" (IsDevCloneFolder: ime "src-vba" + <parent>\.git). Na
' obicnoj klijent masini (nema git klona) odbija se - ne moze slucajno da se
' pokrene iz Alt+F8 nad proizvoljnim folderom.
'
' POSTUPAK: otvori KOPIJU klijenta -> Alt+F8 RunSelfUpdateDev -> izaberi
' ...\otkupapp-pwa\src-vba\ -> backup se pravi sam -> restart. Rollback:
' Backup\AgriX_pre-update_*.xlsm.
' ============================================================
Public Sub RunSelfUpdateDev()
    Const SRC As String = "modSelfUpdate.RunSelfUpdateDev"
    On Error GoTo EH

    If Not SelfVBAAccessible() Then Exit Sub

    If MsgBox("DEV TEST self-update-a iz LOKALNOG git klona (bez Drive-a)." & vbCrLf & vbCrLf & _
              "Code-merge-uje OVU svesku iz izabranog src-vba foldera, isto kao pravi " & _
              "self-update (atomarno). Pokreni na KOPIJI klijenta!" & vbCrLf & vbCrLf & _
              "Nastaviti?", vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Sub

    ' 1) Izbor src-vba foldera + DEV GUARD (mora biti git klon)
    Dim srcFolder As String: srcFolder = PickFolderDev()
    If Len(srcFolder) = 0 Then Exit Sub
    If Not IsDevCloneFolder(srcFolder) Then
        MsgBox "Odbijeno: izabrani folder nije dev git klon." & vbCrLf & _
               "Ocekujem ...\otkupapp-pwa\src-vba\ (ime 'src-vba' + '.git' u roditeljskom folderu)." & vbCrLf & _
               "DEV test se koristi samo na build/dev masini.", vbCritical, APP_NAME
        Exit Sub
    End If

    ' 2) Backup pre svega (rollback) - reuse produkcijskog
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

    ' 4) Isti ATOMARAN core kao pravi self-update
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

' DEV GUARD: izabrani folder mora biti git klon "src-vba" (<parent>\.git postoji).
' Fail-closed na klijentu (nema git klona -> False -> DEV test se ne pokrece).
Private Function IsDevCloneFolder(ByVal folder As String) As Boolean
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Function
    If LCase$(fso.GetBaseName(folder)) <> "src-vba" Then Exit Function
    Dim parent As String: parent = fso.GetParentFolderName(folder)
    IsDevCloneFolder = fso.FolderExists(parent & "\.git")
End Function

' DIJAGNOSTIKA (RUCNO, Alt+F8): READ-ONLY provera sta delta-skip odlucuje. NE
' menja projekat (nema Remove/Import/AddFromString/PrepareRuntime) -> NE MOZE da
' srusi Excel. Poredi kod svake komponente u projektu sa telom fajla iz izabranog
' src-vba (isti SameCode kao update). Pokazuje koliko je "isto" vs "razlicito" i
' PRVU razliku (poziciju + isecak) - da vidimo da li delta-skip radi i, ako ne,
' zasto (line-ending? whitespace? struktura?).
Public Sub Test_DeltaSkip()
    Const SRC As String = "modSelfUpdate.Test_DeltaSkip"
    On Error GoTo EH

    Dim srcFolder As String: srcFolder = PickFolderDev()
    If Len(srcFolder) = 0 Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim skip As String: skip = "," & LCase$(SKIP_MODULES) & ","
    Dim fil As Object, ext As String, baseName As String
    Dim body As String, cur As String, ca As String, cb As String
    Dim vbc As Object
    Dim sameN As Long, diffN As Long, missN As Long, diffList As String, firstDiff As String

    For Each fil In fso.GetFolder(srcFolder).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "doccls" Then
            baseName = fso.GetBaseName(fil.name)
            If InStr(skip, "," & LCase$(baseName) & ",") = 0 Then
                Set vbc = Nothing
                On Error Resume Next
                Set vbc = proj.VBComponents(baseName)
                On Error GoTo EH
                If vbc Is Nothing Then
                    missN = missN + 1
                Else
                    body = ExtractModuleCode(fil.path)
                    cur = ""
                    On Error Resume Next
                    If vbc.CodeModule.CountOfLines > 0 Then cur = vbc.CodeModule.Lines(1, vbc.CodeModule.CountOfLines)
                    On Error GoTo EH
                    If SameCode(cur, body) Then
                        sameN = sameN + 1
                    Else
                        diffN = diffN + 1
                        ca = CanonCode(cur): cb = CanonCode(body)
                        If Len(diffList) < 600 Then diffList = diffList & "  " & fil.name & _
                            " (proj=" & Len(ca) & "b, file=" & Len(cb) & "b)" & vbCrLf
                        If Len(firstDiff) = 0 Then firstDiff = DiffSnippet(fil.name, ca, cb)
                    End If
                End If
            End If
        End If
    Next fil

    MsgBox "Delta-skip provera (READ-ONLY, bez izmena):" & vbCrLf & _
           "  isto:      " & sameN & vbCrLf & _
           "  RAZLICITO: " & diffN & vbCrLf & _
           "  nema komp:  " & missN & vbCrLf & vbCrLf & _
           IIf(diffN = 0, _
               "SVE ISTO -> delta-skip RADI. (Crash je onda u PrepareRuntime/zivi app, ne u importu.)", _
               "Razliciti:" & vbCrLf & diffList & vbCrLf & firstDiff), _
           IIf(diffN = 0, vbInformation, vbExclamation), APP_NAME
    Exit Sub
EH:
    MsgBox "Test_DeltaSkip greska: " & Err.description, vbCritical, APP_NAME
End Sub

' Prva pozicija razlike dve kanonizovane verzije + isecak (prekidi reda vidljivi
' kao \n). Za dijagnostiku zasto SameCode ne prijavi "isto".
Private Function DiffSnippet(ByVal name As String, ByVal a As String, ByVal b As String) As String
    Dim i As Long, n As Long, ascInfo As String
    n = IIf(Len(a) < Len(b), Len(a), Len(b))
    For i = 1 To n
        If Mid$(a, i, 1) <> Mid$(b, i, 1) Then Exit For
    Next i
    If i <= Len(a) And i <= Len(b) Then
        ascInfo = vbCrLf & "  AscW: proj=" & AscW(Mid$(a, i, 1)) & " file=" & AscW(Mid$(b, i, 1))
    End If
    DiffSnippet = "Prva razlika (" & name & ") @ poz " & i & " od " & n & ":" & vbCrLf & _
                  "  proj: [" & Vis(Mid$(a, IIf(i > 15, i - 15, 1), 40)) & "]" & vbCrLf & _
                  "  file: [" & Vis(Mid$(b, IIf(i > 15, i - 15, 1), 40)) & "]" & ascInfo
End Function

' Ucini nevidljive znake vidljivim za MsgBox (prekidi reda / tab). Ne dira
' strukturne vbCrLf u samoj poruci (primenjuje se samo na isecak koda).
Private Function Vis(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, vbLf, "\n")
    s = Replace$(s, vbTab, "\t")
    s = Replace$(s, ChrW$(160), "<NBSP>")
    Vis = s
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
