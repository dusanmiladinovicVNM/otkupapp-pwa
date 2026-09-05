Attribute VB_Name = "modVbaTools"
Option Explicit

' ============================================================
' modVbaTools - DEV alat: round-trip VBA koda src-vba <-> ThisWorkbook.VBProject.
' Pokrece se RUCNO (Alt+F8): ExportAllVBA / ImportAllVBA.
'
' PREDUSLOV: File > Options > Trust Center > Trust Center Settings >
'            Macro Settings > "Trust access to the VBA project object model".
'
' ------------------------------------------------------------
' STA JE ImportAllVBA (v2)
' ------------------------------------------------------------
' NIJE vise "ukloni istoimenu komponentu pa je odmah uvezi". To je bio uzrok
' "The Form or MDIForm name frmX is already in use" (.log fajlovi u src-vba) i
' modX1 duplikata: VBComponents.Remove je u runtime-u ODLOZEN i izvrsi se tek
' kad se makro zavrsi, pa Remove+Import u ISTOM prolazu radi nad imenom koje jos
' postoji. Zato je alat "radio iz drugog pokusaja".
'
' ImportAllVBA sada znaci: USAGLASI projekat sa canonical izvorom (src-vba), uz
' najmanju mogucu izmenu, fail-closed, bez ijedne poznate VBIDE trke.
'
' ------------------------------------------------------------
' ARHITEKTURA (tvrd invarijant, PR #271 / UI katalog 27.18)
' ------------------------------------------------------------
' Projekat ima TACNO JEDNU UserForm - frmOtkupUI. Splash, prijava i mini kartica
' vise nisu forme nego FAZE istog prozora (modUiFaze: BOOT/LOGIN/MINI/APP).
' frmOtkupUI nema kontrole u dizajneru, nema module-level WithEvents ni
' "As MSForms.*", a .frx je zapecacena ljuska (logo zivi u modLogo.bas).
' Alat tu arhitekturu AKTIVNO BRANI: druga .frm u izvoru = odbijen import.
'
' ------------------------------------------------------------
' POLITIKA PO VRSTI KOMPONENTE
' ------------------------------------------------------------
'   frmOtkupUI (postoji)  code merge (DeleteLines+AddFromString), NIKAD Remove
'   frmOtkupUI (nema je)  clean import u fazi 2 (nema sudara imena) + verifikacija
'   .doccls               code merge u postojecu komponentu; nema je = FATALNO
'                         (Document modul se NIKAD ne uklanja i ne dodaje)
'   soft .bas/.cls        code merge; nova komponenta = Add + AddFromString
'   tvrd .bas/.cls        Remove (faza 1) -> OnTime -> Import (faza 2)
'   komponenta van izvora Remove (stale), samo tip 1/2/3, nikad tip 100
'
' "Tvrd" = MODULE-LEVEL (pre prve procedure) "WithEvents" ili "As MSForms.".
' Dodavanje takve deklaracije kroz AddFromString bind-uje MSForms tip-biblioteku
' usred COM edita i diskonektuje CodeModule ([-2147417848]); zato takav modul ide
' iskljucivo kroz Import. Procedure-level "Dim x As MSForms.ComboBox" NIJE tvrd.
' Detekcija je ista kao modSelfUpdate.IsHardModuleBody (dokazan obrazac).
'
' ------------------------------------------------------------
' DVE FAZE - i zasto DoEvents nije resenje
' ------------------------------------------------------------
' Izmedju Remove-a i Import-a mora da se zavrsi PRVA VBA procedura i vrati
' kontrola Excel/VBE message loop-u. DoEvents to NE radi (stack ostaje). Zato:
' faza 1 uradi sve Remove-ove, upise durable stanje (SaveSetting), zakaze
' Application.OnTime i IZADJE; faza 2 (Public, workbook-kvalifikovana) uvozi i
' verifikuje. Faza 1 zbog toga NE sme da prikaze MsgBox kad je faza 2 zakazana -
' modalni dijalog drzi makro na stack-u i Remove se ne bi flush-ovao.
'
' ------------------------------------------------------------
' STA ALAT NE RADI
' ------------------------------------------------------------
' - ne snima radnu svesku (rucni Compile je zavrsna kapija),
' - ne kompajlira preko SendKeys-a i ne dira UI automation,
' - ne dira modSelfUpdate ni release format,
' - ne uvozi sam sebe: modVbaTools se izvrsava dok import traje, pa se preskace;
'   razlika prema izvoru se PRIJAVLJUJE u izvestaju (ne cuti se).
'
' Poredjenje koda (SameCode/CanonCode/LowerOutsideStrings), detekcija tvrdog tela
' i ekstrakcija tela iz export fajla su NAMERNO duplirane iz modSelfUpdate: tamo
' su Private, a modSelfUpdate se u ovom zadatku ne dira.
'
' ASCII-only modul (vidi CLAUDE.md, sekcija 3).
' ============================================================

' ---------- podesavanja po masini (IZMENI; "\" na kraju) ----------
Private Const SRC_FOLDER As String = "C:\Users\Dusan\Documents\GitHub\otkupapp-pwa\src-vba\"
Private Const EXPORT_FOLDER As String = "C:\Users\Dusan\Desktop\AgriX\src-vbaExport23.06.2026_najnoviji\"

' ---------- invarijanti arhitekture ----------
Private Const SELF_MODULE As String = "modVbaTools"      ' ovaj modul (preskace se)
Private Const ONLY_FORM As String = "frmOtkupUI"         ' jedina UserForm projekta

' ---------- faza 2 ----------
Private Const REG_APP As String = "AgriXVbaTools"
Private Const PHASE2_PROC As String = "ImportAllVBA_Phase2"
Private Const PHASE2_SEC As Long = 2

' ---------- plan tekuceg importa ----------
' Jedan import u jednom trenutku (mBusy brani re-entry), pa plan zivi kao stanje
' modula umesto kao deset ByRef parametara kroz svaku proceduru. Ovde NEMA nijedne
' WithEvents / As MSForms. deklaracije - modul mora ostati "soft".
Private mBusy As Boolean
Private mAct As Object          ' lname -> akcija (same/doc/form/formnew/soft/softnew/hard/self)
Private mBody As Object         ' lname -> novo telo iz izvora
Private mSrcMap As Object       ' lname -> ekstenzija izvora
Private mSrcFile As Object      ' lname -> pun put do izvornog fajla
Private mSrcName As Object      ' lname -> ime komponente (case sa diska)
Private mStale As String        ' csv zatecenih komponenti kojih nema u izvoru
Private mHard As String         ' csv IMENA FAJLOVA za fazu 2
Private mGone As String         ' csv imena stvarno uklonjenih u fazi 1
Private mFormNew As Boolean     ' frmOtkupUI ne postoji -> clean import u fazi 2
Private mSelfNote As String     ' izvestaj o samom modVbaTools
Private mSum As String          ' izvestaj faze 1
Private mRbFail As String       ' komponente kojima ROLLBACK NIJE uspeo

' ============================================================
' EXPORT
' ============================================================

' Eksportuj sve komponente projekta u EXPORT_FOLDER (dev snapshot).
Public Sub ExportAllVBA()
    If Not VBAProjectAccessible() Then Exit Sub

    Dim folder As String
    folder = ResolveFolder(EXPORT_FOLDER, "Izaberi folder za EXPORT VBA koda")
    If Len(folder) = 0 Then Exit Sub

    Dim vbc As Object, ext As String, n As Long
    On Error GoTo EH
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        ext = ExtForType(vbc.Type)
        If Len(ext) > 0 Then
            vbc.Export folder & vbc.name & ext
            n = n + 1
        End If
    Next vbc
    MsgBox "Eksportovano komponenti: " & n & vbCrLf & folder, vbInformation, "ExportAllVBA"
    Exit Sub
EH:
    MsgBox "ExportAllVBA nije zavrsen: [" & Err.Number & "] " & Err.description, _
           vbCritical, "ExportAllVBA"
End Sub

' ============================================================
' IMPORT - faza 1
' ============================================================

' Usaglasi VBAProject sa canonical izvorom u src-vba. Jedan korisnicki potez:
' Alt+F8 -> ImportAllVBA. Ako je potrebna faza 2, ona se nastavlja sama.
Public Sub ImportAllVBA()
    Dim problem As String, bkPath As String, folder As String
    Dim fatal As String, recNote As String

    ' 0) re-entry brana. mBusy ostaje True kroz prozor faze 1 -> faze 2; ako je
    '    faza 2 prekinuta (Esc, greska u dispatch-u), operater sme da je odblokira.
    If mBusy Then
        If MsgBox("ImportAllVBA je vec u toku (2. faza je zakazana)." & vbCrLf & vbCrLf & _
                  "Ako je prethodni prolaz prekinut, potvrdi da se stanje ocisti " & _
                  "i krene iz pocetka.", vbYesNo + vbExclamation, "ImportAllVBA") <> vbYes Then Exit Sub
        mBusy = False
    End If

    If Not VBAProjectAccessible() Then Exit Sub

    ' 0b) zaostalo stanje prekinute faze 2 se NE nastavlja slepo - cisti se, a
    '     nedostajuci moduli se ionako vide u novom planu (nema komponente = novo).
    recNote = RecoverImportState()

    folder = ResolveFolder(SRC_FOLDER, "Izaberi src-vba folder")
    If Len(folder) = 0 Then Exit Sub

    ' 1) PREFLIGHT + PLAN - citanje, bez ijedne izmene projekta
    problem = BuildSourceMap(folder)
    If Len(problem) = 0 Then problem = ValidateSourceLayout(folder)
    If Len(problem) = 0 Then problem = ValidateTargetShape()
    If Len(problem) = 0 Then problem = BuildPlan()
    If Len(problem) > 0 Then
        MsgBox "PREFLIGHT je odbio import." & vbCrLf & _
               "NISTA nije menjano u VBA projektu." & vbCrLf & vbCrLf & problem, _
               vbCritical, "ImportAllVBA"
        Exit Sub
    End If

    ' 2) nema sta da se radi -> izlaz bez backup-a, teardown-a i ijedne mutacije
    If Len(mStale) = 0 And Len(mHard) = 0 And Not mFormNew And Not PlanHasMerges() Then
        MsgBox "Projekat je vec usaglasen sa izvorom - 0 izmena." & vbCrLf & vbCrLf & _
               mSum & vbCrLf & mSelfNote & IIf(Len(recNote) > 0, vbCrLf & recNote, ""), _
               vbInformation, "ImportAllVBA"
        Exit Sub
    End If

    If MsgBox("ImportAllVBA ce uskladiti projekat sa izvorom:" & vbCrLf & vbCrLf & _
              Cap(mSum, 900) & vbCrLf & vbCrLf & "Nastaviti?", _
              vbYesNo + vbQuestion, "ImportAllVBA") <> vbYes Then Exit Sub

    mBusy = True

    ' 3) BACKUP pre prve destruktivne operacije. Bez backup-a nema importa.
    bkPath = MakePreImportBackup()
    If Len(bkPath) = 0 Then
        mBusy = False
        MsgBox "Backup NIJE uspeo - import nije ni pokrenut (fail-closed)." & vbCrLf & _
               "Proveri da li je " & ThisWorkbook.path & "\Backup upisiv.", _
               vbCritical, "ImportAllVBA"
        Exit Sub
    End If

    ' 4) zaostali .log iz ranijih neuspelih importa (da nov .log bude nov nalaz)
    DeleteImportLogs folder

    ' 5) oslobodi runtime PRE ijedne VBIDE izmene
    PrepareRuntimeForImport

    ' --- od ove tacke je projekat u izmeni: svaki izlaz ide kroz FAIL/EH ---
    On Error GoTo EH

    ' 6) STALE PRVO: zaostala forma/modul moze da referencira ono cega u novom
    '    izvoru vise nema i da obori compile; ako izvor kaze da ne postoji, nema
    '    razloga da je nosimo kroz ostatak importa.
    RemoveStaleComponents fatal
    If Len(fatal) > 0 Then GoTo FAIL

    ' 7) .doccls (ThisWorkbook / Sheet) - samo code merge, nikad Remove/Add
    MergeDocumentModules fatal
    If Len(fatal) > 0 Then GoTo FAIL

    ' 8) frmOtkupUI - code merge nad postojecom ljuskom
    MergeOnlyForm fatal
    If Len(fatal) > 0 Then GoTo FAIL

    ' 9) soft .bas/.cls
    MergeBasCls fatal
    If Len(fatal) > 0 Then GoTo FAIL

    ' 10) tvrdi .bas/.cls - Remove sada, Import u fazi 2
    RemovePhase2Components fatal
    If Len(fatal) > 0 Then GoTo FAIL

    ' 11) ima li uklanjanja ili clean importa? Onda faza 2 - i zato NEMA MsgBox-a:
    '     modalni dijalog bi zadrzao makro na stack-u i Remove se ne bi flush-ovao.
    '     Zavrsna verifikacija (broj formi, prisustvo komponenti) ima smisla tek
    '     POSLE flush-a, pa je i ona u fazi 2.
    If Len(mHard) > 0 Or Len(mGone) > 0 Or mFormNew Then
        If Not SaveImportPhase2State(folder, bkPath) Then
            fatal = "Nije uspeo upis stanja za 2. fazu (SaveSetting)."
            GoTo FAIL
        End If
        Application.ScreenUpdating = True          ' ekran radi; eventi ostaju off do faze 2
        Application.StatusBar = "ImportAllVBA: 2. faza krece za " & PHASE2_SEC & " s - ne diraj Excel..."
        Application.OnTime Now + TimeSerial(0, 0, PHASE2_SEC), QualifiedProc(PHASE2_PROC)
        Exit Sub                                   ' KRAJ makroa -> VBIDE flush-uje Remove
    End If

    On Error GoTo 0

    ' 12) bez uklanjanja: verifikuj odmah i zavrsi
    problem = VerifyFinalProject(folder)
    RestoreRuntimeAfterImport
    DeleteImportLogs folder
    mBusy = False
    If Len(problem) > 0 Then
        ShowImportFailure "Zavrsna provera projekta NIJE prosla:" & vbCrLf & problem, bkPath
    Else
        ShowImportSuccess mSum & vbCrLf & mSelfNote & IIf(Len(recNote) > 0, vbCrLf & recNote, ""), bkPath
    End If
    Exit Sub

FAIL:
    On Error Resume Next          ' greska u samoj FAIL grani ne sme da vrti EH -> FAIL
    RestoreRuntimeAfterImport
    ClearImportPhase2State
    mBusy = False
    ShowImportFailure fatal, bkPath
    Exit Sub
EH:
    fatal = "Neocekivana greska u 1. fazi: [" & Err.Number & "] " & Err.description
    Resume FAIL
End Sub

' ============================================================
' IMPORT - faza 2 (Application.OnTime; posle flush-a Remove-ova)
' ============================================================

' Public zbog Application.OnTime. Uvozi TACNO ono sto je faza 1 uklonila (lista
' fajlova iz durable stanja) i sprovodi zavrsnu verifikaciju projekta.
Public Sub ImportAllVBA_Phase2()
    Dim sec As String, folder As String, bkPath As String
    Dim hardCsv As String, goneCsv As String, formNew As Boolean
    Dim savedN As Long, phase1Sum As String, problem As String, fatal As String
    Dim arr() As String, i As Long, fn As String, baseName As String, ext As String
    Dim imported As Long, expected As Long, report As String

    sec = P2Section()
    If GetSetting(REG_APP, sec, "pending", "") <> "1" Then
        ' nista zakazano (vec obradjeno ili ocisceno) - tiho, bez dijaloga
        mBusy = False
        RestoreRuntimeAfterImport
        Exit Sub
    End If

    On Error GoTo EH

    folder = GetSetting(REG_APP, sec, "dir", "")
    hardCsv = GetSetting(REG_APP, sec, "hard", "")
    goneCsv = GetSetting(REG_APP, sec, "gone", "")
    formNew = (GetSetting(REG_APP, sec, "formnew", "0") = "1")
    savedN = CLng("0" & GetSetting(REG_APP, sec, "hardn", "0"))
    bkPath = GetSetting(REG_APP, sec, "backup", "")
    phase1Sum = GetSetting(REG_APP, sec, "sum", "")

    If Len(folder) = 0 Then
        fatal = "2. faza: izgubljen je put do izvora (stanje posle 1. faze nije citljivo)." & vbCrLf & _
                "Faza 1 je vec uklonila komponente - projekat je nepotpun."
        GoTo FAIL
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim proj As Object: Set proj = ThisWorkbook.VBProject

    ' 1) sve sto je faza 1 uklonila MORA stvarno biti nestalo. Ako nije, Remove
    '    nije flush-ovan i svaki Import bi opet dao "name already in use" / modX1.
    arr = Split(goneCsv, ",")
    For i = LBound(arr) To UBound(arr)
        baseName = Trim$(arr(i))
        If Len(baseName) > 0 Then
            If ComponentExists(proj, baseName) Then
                fatal = fatal & "  " & baseName & " -> jos postoji (Remove nije dovrsen)" & vbCrLf
            End If
        End If
    Next i
    If Len(fatal) > 0 Then
        fatal = "2. faza: uklanjanje iz 1. faze nije flush-ovano:" & vbCrLf & fatal
        GoTo FAIL
    End If

    ' 2) Import tvrdih .bas/.cls
    arr = Split(hardCsv, ",")
    For i = LBound(arr) To UBound(arr)
        fn = Trim$(arr(i))
        If Len(fn) > 0 Then
            expected = expected + 1
            baseName = BaseNameOf(fn)
            ext = LCase$(ExtOf(fn))
            If StrComp(baseName, SELF_MODULE, vbTextCompare) = 0 Then
                fatal = fatal & "  " & fn & " -> " & SELF_MODULE & " se ne uvozi" & vbCrLf
            ElseIf ComponentExists(proj, baseName) Then
                fatal = fatal & "  " & fn & " -> istoimena komponenta jos postoji" & vbCrLf
            ElseIf ImportOne(proj, folder & fn, baseName, ext, report) Then
                imported = imported + 1
            Else
                fatal = fatal & report
            End If
        End If
    Next i

    ' 3) integritet handoff-a: broj fajlova mora da odgovara onome sto je 1. faza
    '    upisala (skraceno/pokvareno stanje = uklonjeni moduli koji se ne vracaju)
    If Len(fatal) = 0 And expected <> savedN Then
        fatal = "2. faza: lista modula je izmenjena (" & expected & " != " & savedN & ")." & vbCrLf
    End If

    ' 4) clean import jedine forme - samo kad je faza 1 utvrdila da je NEMA
    If Len(fatal) = 0 And formNew Then
        If ComponentExists(proj, ONLY_FORM) Then
            fatal = "2. faza: " & ONLY_FORM & " vec postoji - clean import je otkazan." & vbCrLf
        ElseIf Not FileIsCrLf(folder & ONLY_FORM & ".frm") Then
            fatal = "2. faza: " & ONLY_FORM & ".frm nema CRLF krajeve reda." & vbCrLf & _
                    "Import bi ga uveo kao STANDARDNI modul sa zaglavljem kao kodom." & vbCrLf
        ElseIf Not ImportOne(proj, folder & ONLY_FORM & ".frm", ONLY_FORM, "frm", report) Then
            fatal = report
        End If
    End If

    If Len(fatal) > 0 Then
        fatal = "2. faza NIJE uspela (uvezeno " & imported & "/" & expected & "):" & vbCrLf & fatal
        GoTo FAIL
    End If

    ' 5) zavrsna verifikacija nad flush-ovanim projektom
    problem = VerifyFinalProject(folder)
    DeleteImportLogs folder
    ClearImportPhase2State
    RestoreRuntimeAfterImport
    mBusy = False

    If Len(problem) > 0 Then
        ShowImportFailure "Zavrsna provera projekta NIJE prosla:" & vbCrLf & problem & vbCrLf & _
                          vbCrLf & phase1Sum, bkPath
    Else
        ShowImportSuccess phase1Sum & vbCrLf & "2. faza: uvezeno " & imported & " tvrdih modula" & _
                          IIf(formNew, " + " & ONLY_FORM, ""), bkPath
    End If
    Exit Sub

FAIL:
    On Error Resume Next          ' greska u samoj FAIL grani ne sme da vrti EH -> FAIL
    ' .log fajlovi se brisu TEK posle hvatanja greske - brisanje loga nikad ne sme
    ' da maskira razlog pada.
    DeleteImportLogs folder
    ClearImportPhase2State
    RestoreRuntimeAfterImport
    mBusy = False
    ShowImportFailure fatal, bkPath
    Exit Sub
EH:
    fatal = "Neocekivana greska u 2. fazi: [" & Err.Number & "] " & Err.description
    Resume FAIL
End Sub

' ============================================================
' PREFLIGHT / PLAN
' ============================================================

' Napuni mSrcMap / mSrcFile / mSrcName iz foldera. "" = uredu.
Private Function BuildSourceMap(ByVal folder As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fil As Object, ext As String, baseName As String, lname As String

    Set mSrcMap = CreateObject("Scripting.Dictionary")
    Set mSrcFile = CreateObject("Scripting.Dictionary")
    Set mSrcName = CreateObject("Scripting.Dictionary")

    On Error GoTo EH
    For Each fil In fso.GetFolder(folder).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Or ext = "doccls" Then
            baseName = fso.GetBaseName(fil.name)
            lname = LCase$(baseName)
            If mSrcMap.Exists(lname) Then
                BuildSourceMap = "Dva izvorna fajla za istu komponentu: " & _
                                 mSrcName(lname) & "." & mSrcMap(lname) & " i " & fil.name
                Exit Function
            End If
            mSrcMap(lname) = ext
            mSrcFile(lname) = fil.path
            mSrcName(lname) = baseName
        End If
    Next fil

    If mSrcMap.count = 0 Then
        BuildSourceMap = "U folderu nema nijednog .bas/.cls/.frm/.doccls fajla:" & vbCrLf & folder
    End If
    Exit Function
EH:
    BuildSourceMap = "Citanje izvornog foldera nije uspelo: [" & Err.Number & "] " & Err.description
End Function

' Canonical source invarijant: TACNO jedna .frm i TACNO jedna .frx, obe frmOtkupUI.
Private Function ValidateSourceLayout(ByVal folder As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fil As Object, ext As String
    Dim frmN As Long, frxN As Long, frmList As String, frxList As String

    On Error GoTo EH
    For Each fil In fso.GetFolder(folder).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        If ext = "frm" Then
            frmN = frmN + 1
            frmList = frmList & "  " & fil.name & vbCrLf
        ElseIf ext = "frx" Then
            frxN = frxN + 1
            frxList = frxList & "  " & fil.name & vbCrLf
        End If
    Next fil

    If frmN <> 1 Or Not mSrcMap.Exists(LCase$(ONLY_FORM)) Then
        ValidateSourceLayout = _
            "Izvor mora imati TACNO JEDNU UserForm: " & ONLY_FORM & ".frm" & vbCrLf & _
            "Nadjeno .frm fajlova: " & frmN & vbCrLf & frmList & vbCrLf & _
            "Nova arhitektura (PR #271) NE dozvoljava dodatne UserForm-e - splash," & vbCrLf & _
            "prijava i mini kartica su FAZE istog prozora (modUiFaze), ne forme." & vbCrLf & _
            "Nov ekran je modScr*, nov pun-ekran sadrzaj je faza. Trece ne postoji."
        Exit Function
    End If

    If frxN <> 1 Or Not fso.FileExists(folder & ONLY_FORM & ".frx") Then
        ValidateSourceLayout = _
            "Izvor mora imati TACNO JEDAN .frx: " & ONLY_FORM & ".frx" & vbCrLf & _
            "Nadjeno .frx fajlova: " & frxN & vbCrLf & frxList
        Exit Function
    End If
    Exit Function
EH:
    ValidateSourceLayout = "Provera izvornog rasporeda nije uspela: [" & Err.Number & "] " & Err.description
End Function

' Zatecen projekat mora da bude oblika koji plan ume da odrzi. Sve provere su
' PRE ijedne izmene (fail-closed sa 0 mutacija).
Private Function ValidateTargetShape() As String
    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim k As Variant, lname As String, ext As String, want As Long
    Dim vbc As Object, ctlN As Long, okOut As Boolean, cur As String

    On Error GoTo EH
    For Each k In mSrcMap.Keys
        lname = CStr(k)
        If lname <> LCase$(SELF_MODULE) Then
            ext = mSrcMap(lname)
            want = TypeForExt(ext)
            If ComponentExists(proj, CStr(mSrcName(lname))) Then
                Set vbc = proj.VBComponents(CStr(mSrcName(lname)))
                If vbc.Type <> want Then
                    ValidateTargetShape = "Komponenta '" & mSrcName(lname) & "' je tipa " & vbc.Type & _
                        ", a izvor (" & ext & ") trazi tip " & want & "." & vbCrLf & _
                        "Alat ne menja tip komponente - resi rucno u VBE."
                    Exit Function
                End If
            End If
        End If
    Next k

    ' frmOtkupUI, ako postoji, mora ostati prazna zapecacena ljuska
    If ComponentExists(proj, ONLY_FORM) Then
        Set vbc = proj.VBComponents(ONLY_FORM)
        If vbc.Type <> 3 Then
            ValidateTargetShape = ONLY_FORM & " u projektu nije UserForm (tip " & vbc.Type & ")."
            Exit Function
        End If
        ctlN = DesignerControlCount(vbc, okOut)
        If Not okOut Then
            ValidateTargetShape = "Ne moze da se procita dizajner forme " & ONLY_FORM & _
                " - import se ne pokrece (fail-closed)."
            Exit Function
        End If
        If ctlN <> 0 Then
            ValidateTargetShape = ONLY_FORM & " ima " & ctlN & " kontrolu(e) u DIZAJNERU." & vbCrLf & _
                "Ljuska mora ostati prazna (sve kontrole nastaju u runtime-u)." & vbCrLf & _
                "Alat ne radi automatski popravak dizajnera - resi rucno u VBE."
            Exit Function
        End If
        cur = ComponentCode(vbc, okOut)
        If Not okOut Then
            ValidateTargetShape = "Ne moze da se procita kod forme " & ONLY_FORM & "."
            Exit Function
        End If
        If IsHardModuleBody(cur) Then
            ValidateTargetShape = "Zatecena " & ONLY_FORM & " ima MODULE-LEVEL WithEvents ili" & vbCrLf & _
                "'As MSForms.' deklaraciju. Code merge nad takvim telom diskonektuje" & vbCrLf & _
                "CodeModule, a forma se ne sme Remove+Import-ovati u runtime-u." & vbCrLf & _
                "Resi rucno u VBE (ocisti deklaracije) pa ponovi import."
            Exit Function
        End If
    End If
    Exit Function
EH:
    ValidateTargetShape = "Provera zatecenog projekta nije uspela: [" & Err.Number & "] " & Err.description
End Function

' Odluci sta se radi sa svakom komponentom. Citanje, bez izmene. "" = uredu.
Private Function BuildPlan() As String
    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim k As Variant, lname As String, ext As String, baseName As String
    Dim body As String, cur As String, okOut As Boolean, exists As Boolean, vbc As Object
    Dim nSame As Long, nDoc As Long, nForm As Long, nSoft As Long, nNew As Long
    Dim nHard As Long, nStale As Long, nEmpty As Long, emptyS As String, c As Object

    Set mAct = CreateObject("Scripting.Dictionary")
    Set mBody = CreateObject("Scripting.Dictionary")
    mStale = "": mHard = "": mGone = "": mSum = "": mSelfNote = "": mRbFail = ""
    mFormNew = False

    On Error GoTo EH
    For Each k In mSrcMap.Keys
        lname = CStr(k)
        ext = mSrcMap(lname)
        baseName = mSrcName(lname)

        body = ExtractModuleCode(CStr(mSrcFile(lname)))

        exists = ComponentExists(proj, baseName)
        cur = ""
        If exists Then
            Set vbc = proj.VBComponents(baseName)
            cur = ComponentCode(vbc, okOut)
            If Not okOut Then
                BuildPlan = "Ne moze da se procita kod komponente '" & baseName & "'."
                Exit Function
            End If
        End If

        If lname = LCase$(SELF_MODULE) Then
            ' Sam sebe ne dira, ali cutati ne sme.
            mAct(lname) = "self"
            If Not exists Then
                mSelfNote = SELF_MODULE & ": nema ga u projektu (a izvrsava se?) - proveri rucno."
            ElseIf SameCode(cur, body) Then
                mSelfNote = SELF_MODULE & ": isti kao izvor."
            Else
                mSelfNote = SELF_MODULE & ": SOURCE RAZLICIT -- nije primenjen jer modul izvrsava import." & vbCrLf & _
                            "  Zameni ga rucno u VBE (import fajla " & baseName & ".bas) pa ponovi."
            End If
        ElseIf exists And SameCode(cur, body) Then
            mAct(lname) = "same"
            nSame = nSame + 1
        ElseIf Len(body) = 0 Then
            ' PRAZAN IZVOR = NO-OP, za svaku ekstenziju. Prazan izvozni fajl ne
            ' opisuje komponentu: 42 od 43 .doccls u src-vba nemaju nijednu
            ' liniju koda (samo ThisWorkbook ima telo) - to su artefakti prvog
            ' punog eksporta, a ne tvrdnja "ovaj modul mora biti prazan".
            ' Zato prazan izvor NIKAD ne brise zatecen kod (fail-safe i protiv
            ' lose ekstrakcije nad ne-praznim fajlom) i NIKAD ne obara import
            ' zbog lista koga u ovoj svesci nema.
            mAct(lname) = "same"
            nEmpty = nEmpty + 1
            If Not exists Then emptyS = emptyS & "  " & baseName & "." & ext & vbCrLf
        ElseIf ext = "doccls" Then
            ' .doccls SA KODOM mora imati svoju komponentu - document modul se
            ' ne kreira ni Add-om ni Import-om, pa je jedini ishod fail-closed.
            If Not exists Then
                BuildPlan = "Izvor ima " & baseName & ".doccls SA KODOM, a u projektu NEMA" & vbCrLf & _
                    "odgovarajuci document modul (list ili ThisWorkbook)." & vbCrLf & _
                    "Document moduli se ne kreiraju importom - fali list u radnoj svesci."
                Exit Function
            End If
            mAct(lname) = "doc"
            mBody(lname) = body
            nDoc = nDoc + 1
        ElseIf ext = "frm" Then
            If IsHardModuleBody(body) Then
                BuildPlan = "Izvor " & baseName & ".frm ima MODULE-LEVEL WithEvents ili 'As MSForms.'" & vbCrLf & _
                    "deklaraciju. To je zabranjeno u jedinoj formi projekta (vidi" & vbCrLf & _
                    ".claude/rules/forme-i-kontrole.md) - sinkovi idu kroz clsUiSink / clsFlatBtn."
                Exit Function
            End If
            If exists Then
                mAct(lname) = "form"
                mBody(lname) = body
                nForm = nForm + 1
            Else
                mAct(lname) = "formnew"
                mFormNew = True
                nForm = nForm + 1
            End If
        ElseIf IsHardModuleBody(body) Then
            ' tvrd .bas/.cls -> nikad AddFromString; Remove (faza 1) + Import (faza 2)
            mAct(lname) = "hard"
            AddCsv mHard, baseName & "." & ext
            nHard = nHard + 1
        ElseIf exists Then
            mAct(lname) = "soft"
            mBody(lname) = body
            nSoft = nSoft + 1
        Else
            mAct(lname) = "softnew"
            mBody(lname) = body
            nNew = nNew + 1
        End If
    Next k

    ' NEGATIVE DELTA: sve sto je u projektu a nema ga u izvoru je stale.
    ' Samo tip 1/2/3 - document moduli (tip 100) nisu kandidati za brisanje.
    For Each c In proj.VBComponents
        If c.Type = 1 Or c.Type = 2 Or c.Type = 3 Then
            If StrComp(c.name, SELF_MODULE, vbTextCompare) <> 0 Then
                If Not mSrcMap.Exists(LCase$(c.name)) Then
                    AddCsv mStale, c.name
                    nStale = nStale + 1
                End If
            End If
        End If
    Next c

    mSum = "Plan:" & vbCrLf & _
           "  bez izmene (delta-skip): " & nSame & vbCrLf & _
           "  document moduli (merge): " & nDoc & vbCrLf & _
           "  " & ONLY_FORM & ": " & IIf(mFormNew, "clean import (nema je)", IIf(nForm > 0, "code merge", "bez izmene")) & vbCrLf & _
           "  soft .bas/.cls merge: " & nSoft & ", novih: " & nNew & vbCrLf & _
           "  tvrdi (Remove + 2. faza Import): " & nHard & vbCrLf & _
           "  STALE (van izvora, Remove): " & nStale & vbCrLf & _
           IIf(nStale > 0, "    " & mStale & vbCrLf, "") & _
           IIf(nHard > 0, "    tvrdi: " & mHard & vbCrLf, "") & _
           "  prazan izvor (preskocen): " & nEmpty & vbCrLf & _
           IIf(Len(emptyS) > 0, "    bez komponente u svesci:" & vbCrLf & Cap(emptyS, 300), "")
    Exit Function
EH:
    BuildPlan = "Pravljenje plana nije uspelo: [" & Err.Number & "] " & Err.description
End Function

' Ima li plan ijedan code merge / Add?
Private Function PlanHasMerges() As Boolean
    Dim k As Variant
    For Each k In mAct.Keys
        Select Case mAct(k)
            Case "doc", "form", "soft", "softnew": PlanHasMerges = True: Exit Function
        End Select
    Next k
End Function

' ============================================================
' PRIMENA
' ============================================================

' Ukloni komponente kojih nema u izvoru (tip 1/2/3). Imena se skupljaju unapred -
' Remove usred For Each nad VBComponents pomera kolekciju.
Private Sub RemoveStaleComponents(ByRef fatal As String)
    If Len(mStale) = 0 Then Exit Sub

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim arr() As String, i As Long, nm As String, c As Object
    arr = Split(mStale, ",")
    For i = LBound(arr) To UBound(arr)
        nm = Trim$(arr(i))
        If Len(nm) > 0 Then
            Set c = Nothing
            On Error Resume Next
            Set c = proj.VBComponents(nm)
            On Error GoTo 0
            If c Is Nothing Then
                mSum = mSum & "  stale " & nm & ": vise ga nema" & vbCrLf
            ElseIf c.Type = 100 Then
                fatal = fatal & "Odbijeno uklanjanje document modula '" & nm & "'." & vbCrLf
            Else
                On Error Resume Next
                Err.Clear
                proj.VBComponents.Remove c
                If Err.Number <> 0 Then
                    fatal = fatal & "Remove '" & nm & "' nije uspeo: [" & Err.Number & "] " & Err.description & vbCrLf
                Else
                    AddCsv mGone, nm
                End If
                On Error GoTo 0
            End If
        End If
    Next i
    If Len(mGone) > 0 Then mSum = mSum & "Uklonjeno (van izvora): " & mGone & vbCrLf
End Sub

' .doccls: samo code merge u postojecu komponentu. Pad = fatalno (uz rollback).
Private Sub MergeDocumentModules(ByRef fatal As String)
    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim k As Variant, nm As String, errS As String, rbOk As Boolean, n As Long

    For Each k In mAct.Keys
        If mAct(k) = "doc" Then
            nm = mSrcName(k)
            If Not ComponentExists(proj, nm) Then
                fatal = fatal & "Document modul '" & nm & "' je nestao tokom importa." & vbCrLf
                Exit Sub
            End If
            If ReplaceCodeWithRollback(proj.VBComponents(nm), CStr(mBody(k)), errS, rbOk) Then
                n = n + 1
            Else
                If Not rbOk Then mRbFail = mRbFail & "  " & nm & vbCrLf
                fatal = fatal & "Code merge document modula '" & nm & "' nije uspeo: " & errS & vbCrLf
                Exit Sub
            End If
        End If
    Next k
    If n > 0 Then mSum = mSum & "Document moduli azurirani: " & n & vbCrLf
End Sub

' frmOtkupUI: iskljucivo code merge nad postojecom komponentom. NIKAD Remove.
Private Sub MergeOnlyForm(ByRef fatal As String)
    Dim lname As String: lname = LCase$(ONLY_FORM)
    If Not mAct.Exists(lname) Then Exit Sub
    If mAct(lname) <> "form" Then Exit Sub

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim vbc As Object, errS As String, rbOk As Boolean, ctlN As Long, okOut As Boolean

    If Not ComponentExists(proj, ONLY_FORM) Then
        fatal = "Forma " & ONLY_FORM & " je nestala tokom importa."
        Exit Sub
    End If
    Set vbc = proj.VBComponents(ONLY_FORM)

    ' ponovi kapije neposredno pre izmene (stanje se moglo promeniti)
    If vbc.Type <> 3 Then
        fatal = ONLY_FORM & " nije UserForm (tip " & vbc.Type & ") - merge otkazan."
        Exit Sub
    End If
    ctlN = DesignerControlCount(vbc, okOut)
    If Not okOut Or ctlN <> 0 Then
        fatal = ONLY_FORM & ": dizajner nije prazan (kontrola: " & ctlN & ") ili nije citljiv - merge otkazan."
        Exit Sub
    End If

    If ReplaceCodeWithRollback(vbc, CStr(mBody(lname)), errS, rbOk) Then
        mSum = mSum & ONLY_FORM & ": code merge OK (forma NIJE uklanjana)" & vbCrLf
    Else
        If Not rbOk Then mRbFail = mRbFail & "  " & ONLY_FORM & vbCrLf
        fatal = "Code merge forme " & ONLY_FORM & " nije uspeo: " & errS
    End If
End Sub

' Soft .bas/.cls: merge u postojecu ili Add za novu. Pad -> rollback pa fallback
' na fazu 2 (Remove + Import podnosi vise od AddFromString-a).
Private Sub MergeBasCls(ByRef fatal As String)
    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim k As Variant, lname As String, nm As String, ext As String
    Dim errS As String, rbOk As Boolean, n As Long, nAdd As Long
    Dim vbc As Object, addedName As String, addErr As Long

    For Each k In mAct.Keys
        lname = CStr(k)
        Select Case mAct(lname)

        Case "soft"
            nm = mSrcName(lname)
            ext = mSrcMap(lname)
            If Not ComponentExists(proj, nm) Then
                fatal = "Komponenta '" & nm & "' je nestala tokom importa."
                Exit Sub
            End If
            If ReplaceCodeWithRollback(proj.VBComponents(nm), CStr(mBody(lname)), errS, rbOk) Then
                n = n + 1
            ElseIf rbOk Then
                ' stari kod je vracen -> bezbedno je pokusati fazu 2
                mSum = mSum & "  " & nm & ": merge pao (" & errS & ") -> 2. faza" & vbCrLf
                On Error Resume Next
                Err.Clear
                proj.VBComponents.Remove proj.VBComponents(nm)
                If Err.Number = 0 Then
                    AddCsv mGone, nm
                    AddCsv mHard, nm & "." & ext
                Else
                    fatal = "Fallback Remove '" & nm & "' nije uspeo: [" & Err.Number & "] " & Err.description
                End If
                On Error GoTo 0
                If Len(fatal) > 0 Then Exit Sub
            Else
                mRbFail = mRbFail & "  " & nm & vbCrLf
                fatal = "Code merge '" & nm & "' nije uspeo: " & errS
                Exit Sub
            End If

        Case "softnew"
            nm = mSrcName(lname)
            ext = mSrcMap(lname)
            Set vbc = Nothing
            addedName = ""
            On Error Resume Next
            Err.Clear
            Set vbc = proj.VBComponents.Add(TypeForExt(ext))
            If Err.Number = 0 Then
                vbc.name = nm
                If Len(CStr(mBody(lname))) > 0 Then vbc.CodeModule.AddFromString CStr(mBody(lname))
                addedName = vbc.name
            End If
            addErr = Err.Number
            errS = "[" & Err.Number & "] " & Err.description
            On Error GoTo 0

            If addErr = 0 And StrComp(addedName, nm, vbTextCompare) = 0 Then
                nAdd = nAdd + 1
            Else
                ' nedovrsena komponenta se uklanja pa se fajl uvozi u fazi 2
                mSum = mSum & "  " & nm & ": Add pao (" & errS & ", ime='" & addedName & "') -> 2. faza" & vbCrLf
                On Error Resume Next
                If Not vbc Is Nothing Then proj.VBComponents.Remove vbc
                On Error GoTo 0
                If Len(addedName) > 0 Then AddCsv mGone, addedName
                AddCsv mHard, nm & "." & ext
            End If

        End Select
    Next k

    If n > 0 Or nAdd > 0 Then _
        mSum = mSum & "Soft moduli: azurirano " & n & ", novih " & nAdd & vbCrLf
End Sub

' Ukloni tvrde .bas/.cls (faza 1). Import ide u fazi 2, posle flush-a.
Private Sub RemovePhase2Components(ByRef fatal As String)
    If Len(mHard) = 0 Then Exit Sub

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim arr() As String, i As Long, nm As String, c As Object
    arr = Split(mHard, ",")
    For i = LBound(arr) To UBound(arr)
        nm = BaseNameOf(Trim$(arr(i)))
        If Len(nm) > 0 Then
            If Not CsvHas(mGone, nm) Then
                Set c = Nothing
                On Error Resume Next
                Set c = proj.VBComponents(nm)
                On Error GoTo 0
                If Not c Is Nothing Then
                    If c.Type <> 1 And c.Type <> 2 Then
                        fatal = fatal & "Tvrd modul '" & nm & "' nije std/class (tip " & c.Type & ")." & vbCrLf
                    Else
                        On Error Resume Next
                        Err.Clear
                        proj.VBComponents.Remove c
                        If Err.Number <> 0 Then
                            fatal = fatal & "Remove tvrdog '" & nm & "' nije uspeo: [" & _
                                    Err.Number & "] " & Err.description & vbCrLf
                        Else
                            AddCsv mGone, nm
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Zameni kod komponente uz rollback. rollbackOk je True kad rollback nije ni bio
' potreban ILI je uspeo; False znaci da je komponenta ostala BEZ ispravnog koda.
Private Function ReplaceCodeWithRollback(ByVal vbc As Object, ByVal newBody As String, _
                                         ByRef errOut As String, ByRef rollbackOk As Boolean) As Boolean
    Dim old As String, okOut As Boolean, errNum As Long, errDesc As String
    rollbackOk = True
    errOut = ""

    old = ComponentCode(vbc, okOut)
    If Not okOut Then
        errOut = "stari kod nije citljiv"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    If vbc.CodeModule.CountOfLines > 0 Then vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
    If Len(newBody) > 0 Then vbc.CodeModule.AddFromString newBody
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo 0

    If errNum = 0 Then
        ReplaceCodeWithRollback = True
        Exit Function
    End If

    errOut = "[" & errNum & "] " & errDesc
    ' pokusaj povratka na staro telo
    On Error Resume Next
    Err.Clear
    If vbc.CodeModule.CountOfLines > 0 Then vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
    If Len(old) > 0 Then vbc.CodeModule.AddFromString old
    If Err.Number <> 0 Then
        rollbackOk = False
        errOut = errOut & " + ROLLBACK PAO [" & Err.Number & "] " & Err.description
    End If
    On Error GoTo 0
End Function

' Uvezi jedan fajl i verifikuj VRACENU komponentu (ime + tip). report dobija opis
' greske (prazan na uspeh).
Private Function ImportOne(ByVal proj As Object, ByVal fpath As String, ByVal baseName As String, _
                           ByVal ext As String, ByRef report As String) As Boolean
    Dim c As Object, impErr As Long, impDesc As String, gotName As String
    report = ""

    On Error Resume Next
    Err.Clear
    Set c = proj.VBComponents.Import(fpath)
    impErr = Err.Number
    impDesc = Err.description
    gotName = "?"
    If Not c Is Nothing Then gotName = c.name
    On Error GoTo 0

    If impErr = 0 And Not c Is Nothing Then
        If StrComp(gotName, baseName, vbTextCompare) = 0 And c.Type = TypeForExt(ext) Then
            ImportOne = True
            Exit Function
        End If
    End If

    report = "  " & baseName & "." & ext & " -> Import nije verifikovan [" & impErr & "] " & _
             impDesc & " (ime='" & gotName & "')" & vbCrLf
End Function

' ============================================================
' ZAVRSNA VERIFIKACIJA
' ============================================================

' "" = projekat odgovara izvoru. Zove se tek nad flush-ovanim projektom.
Private Function VerifyFinalProject(ByVal folder As String) As String
    Dim problem As String
    problem = BuildSourceMap(folder)
    If Len(problem) > 0 Then
        VerifyFinalProject = "ponovno citanje izvora: " & problem
        Exit Function
    End If

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim c As Object, k As Variant, lname As String, nm As String, want As Long
    Dim formN As Long, formNames As String, extra As String, missing As String
    Dim ctlN As Long, okOut As Boolean, bad As String

    On Error GoTo EH

    ' 1) forme: tacno jedna, i to frmOtkupUI, prazna u dizajneru
    For Each c In proj.VBComponents
        If c.Type = 3 Then
            formN = formN + 1
            formNames = formNames & " " & c.name
        End If
        If c.Type = 1 Or c.Type = 2 Or c.Type = 3 Then
            If StrComp(c.name, SELF_MODULE, vbTextCompare) <> 0 Then
                If Not mSrcMap.Exists(LCase$(c.name)) Then extra = extra & " " & c.name
            End If
        End If
    Next c

    If formN <> 1 Or Trim$(formNames) <> ONLY_FORM Then
        bad = bad & "  UserForm-i: " & formN & " (" & Trim$(formNames) & "), ocekivano 1 (" & ONLY_FORM & ")" & vbCrLf
    Else
        ctlN = DesignerControlCount(proj.VBComponents(ONLY_FORM), okOut)
        If Not okOut Then
            bad = bad & "  " & ONLY_FORM & ": dizajner nije citljiv" & vbCrLf
        ElseIf ctlN <> 0 Then
            bad = bad & "  " & ONLY_FORM & ": " & ctlN & " kontrola u dizajneru (ocekivano 0)" & vbCrLf
        End If
    End If

    ' 2) svaka izvorna komponenta postoji sa tacnim tipom
    For Each k In mSrcMap.Keys
        lname = CStr(k)
        If lname <> LCase$(SELF_MODULE) Then
            nm = mSrcName(lname)
            want = TypeForExt(CStr(mSrcMap(lname)))
            If Not ComponentExists(proj, nm) Then
                ' prazan izvorni fajl ne opisuje komponentu (v. BuildPlan) -
                ' isto pravilo mora vaziti i ovde, inace plan preskoci a
                ' provera trazi pa oborila bi uspesan import
                If Len(ExtractModuleCode(CStr(mSrcFile(lname)))) > 0 Then missing = missing & " " & nm
            ElseIf proj.VBComponents(nm).Type <> want Then
                bad = bad & "  " & nm & ": tip " & proj.VBComponents(nm).Type & ", ocekivano " & want & vbCrLf
            End If
        End If
    Next k

    If Len(missing) > 0 Then bad = bad & "  nedostaju komponente:" & missing & vbCrLf
    ' 3) visak = zaostala stale komponenta ILI artefakt neuspelog importa (modX1)
    If Len(extra) > 0 Then bad = bad & "  komponente van izvora (visak):" & extra & vbCrLf

    VerifyFinalProject = bad
    Exit Function
EH:
    VerifyFinalProject = "  provera je pukla: [" & Err.Number & "] " & Err.description
End Function

' ============================================================
' RUNTIME
' ============================================================

' Oslobodi runtime stanje pre VBIDE izmene: ugasi tajmere, otpusti module-level
' reference dinamickih kontrola/WithEvents, skini mouse hook, unload sve forme.
' SVI pozivi su KASNO VEZANI i fail-soft: modVbaTools nije u self-update kanalu,
' pa ne sme da obori compile na klijentu koji neku od tih procedura nema.
Private Sub PrepareRuntimeForImport()
    On Error Resume Next

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Application.OnTime tikovi: tik koji opali usred importa (ili u prozoru
    ' izmedju faza) forsira demand-compile nepotpunog projekta.
    CallOptional "StopScheduledSync"          ' modGoogleSyncOrchestrator
    CallOptional "StopAutoSaveTimer"          ' modJournaling
    CallOptional "StopHeartbeatTimer"         ' modStanicaLock
    CallOptional "StopStornoWarm"             ' modStornoWarm
    CallOptional "StopOtkupUITimers"          ' modOtkupUI (toast)

    ' Module-level WithEvents / kontrole. OtkupUI_Release mora da odradi svoje pre
    ' unload-a: panel radne povrsine drzi OKVIR unutar forme, a modul panela njega.
    CallOptional "OtkupBlok_Release"
    CallOptional "Podesavanja_Release"
    CallOptional "MaticniMenu_Release"
    CallOptional "Admin_Release"
    CallOptional "KarticaDetalji_Reset"
    CallOptional "MouseWheel_Off"
    CallOptional "OtkupUI_Release"

    Do While VBA.UserForms.count > 0
        Unload VBA.UserForms(0)
    Loop

    ' NB: ovaj DoEvents pusta unload-e da se sleknu. NIJE sinhronizacija za
    ' VBComponents.Remove - to radi iskljucivo izlazak iz makroa + Application.OnTime.
    DoEvents
    On Error GoTo 0
End Sub

' Vrati aplikaciona podesavanja. MORA se pozvati na SVAKOM izlazu (uspeh, greska,
' prekid, oporavak) - inace Workbook_Open ne opali pri sledecem otvaranju fajla u
' ISTOJ Excel instanci.
Private Sub RestoreRuntimeAfterImport()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' Fail-soft kasno vezan poziv opcione procedure u OVOJ radnoj svesci.
Private Sub CallOptional(ByVal procName As String)
    On Error Resume Next
    Application.Run QualifiedProc(procName)
    Err.Clear
End Sub

' ============================================================
' DURABLE STANJE FAZE 2
' ============================================================

' pending="1" se upisuje POSLEDNJI - polovicno stanje se ne racuna kao zakazano.
Private Function SaveImportPhase2State(ByVal folder As String, ByVal bkPath As String) As Boolean
    Dim sec As String: sec = P2Section()
    On Error GoTo EH
    SaveSetting REG_APP, sec, "dir", folder
    SaveSetting REG_APP, sec, "hard", mHard
    SaveSetting REG_APP, sec, "hardn", CStr(CsvCount(mHard))
    SaveSetting REG_APP, sec, "gone", mGone
    SaveSetting REG_APP, sec, "formnew", IIf(mFormNew, "1", "0")
    SaveSetting REG_APP, sec, "backup", bkPath
    SaveSetting REG_APP, sec, "sum", Cap(mSum, 900)
    SaveSetting REG_APP, sec, "pending", "1"
    SaveImportPhase2State = True
    Exit Function
EH:
    SaveImportPhase2State = False
End Function

Private Sub ClearImportPhase2State()
    On Error Resume Next
    DeleteSetting REG_APP, P2Section()
    Err.Clear
End Sub

' Zaostalo stanje prekinute faze 2 se NE nastavlja - cisti se i prijavljuje.
' Vraca napomenu za izvestaj ("" ako nije bilo nicega).
Private Function RecoverImportState() As String
    Dim sec As String: sec = P2Section()
    On Error Resume Next
    If GetSetting(REG_APP, sec, "pending", "") = "1" Then
        RecoverImportState = "NAPOMENA: zatecena je prekinuta 2. faza ranijeg importa - stanje je" & vbCrLf & _
            "ocisceno i krenulo se iz pocetka (backup: " & GetSetting(REG_APP, sec, "backup", "?") & ")."
        DeleteSetting REG_APP, sec
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
    End If
    Err.Clear
End Function

' Sekcija u registru scope-ovana po radnoj svesci - dve otvorene kopije ne dele
' stanje faze 2.
Private Function P2Section() As String
    Dim s As String, i As Long, ch As String, out As String
    s = ThisWorkbook.name
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Then out = out & ch
    Next i
    P2Section = "import_" & out
End Function

' Workbook-kvalifikovano ime procedure ("'Ime.xlsm'!Proc") - kad su dve kopije
' otvorene, OnTime/Run moraju da pogode PRAVU svesku.
Private Function QualifiedProc(ByVal procName As String) As String
    QualifiedProc = "'" & Replace$(ThisWorkbook.name, "'", "''") & "'!" & procName
End Function

' ============================================================
' BACKUP / LOG / FAJLOVI
' ============================================================

' Kopija radne sveske pre prve destruktivne operacije. "" = neuspeh (fail-closed).
Private Function MakePreImportBackup() As String
    On Error GoTo EH
    ' NB: promenljiva se ne zove 'dir' - sudara se sa ugradjenom Dir().
    Dim bkDir As String: bkDir = ThisWorkbook.path & "\Backup"
    If Dir(bkDir, vbDirectory) = "" Then MkDir bkDir

    Dim baseName As String: baseName = BaseNameOf(ThisWorkbook.name)
    Dim nm As String
    nm = baseName & "_pre-vba-import_" & Format$(Now, "yyyy-mm-dd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs bkDir & "\" & nm
    If Dir(bkDir & "\" & nm) = "" Then Exit Function      ' kopija nije stvarno nastala
    MakePreImportBackup = bkDir & "\" & nm
    Exit Function
EH:
    MakePreImportBackup = ""
End Function

' VBIDE ostavlja frmX.log / modX.log kad Import padne. Cisti se PRE rada (da nov
' log bude nov nalaz) i posle faze 2 - ali TEK posle hvatanja Err.Number, nikad
' kao maskiranje greske.
Private Sub DeleteImportLogs(ByVal folder As String)
    On Error Resume Next
    Dim f As String, doomed As String
    f = Dir(folder & "*.log")
    Do While Len(f) > 0
        doomed = doomed & f & "|"
        f = Dir()
    Loop
    Dim arr() As String, i As Long
    arr = Split(doomed, "|")
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) > 0 Then Kill folder & arr(i)
    Next i
    Err.Clear
End Sub

' Da li fajl ima CRLF krajeve reda. .frm sa LF-om Import NE prepozna kao formu
' nego kao standardni modul sa zaglavljem kao kodom (v. .gitattributes).
Private Function FileIsCrLf(ByVal fpath As String) As Boolean
    On Error GoTo EH
    Dim ff As Integer, s As String, n As Long
    ff = FreeFile
    Open fpath For Binary As #ff
    n = LOF(ff)
    If n > 4096 Then n = 4096
    If n > 0 Then
        s = Space$(n)
        Get #ff, 1, s
    End If
    Close #ff
    FileIsCrLf = (InStr(1, s, vbCrLf) > 0)
    Exit Function
EH:
    On Error Resume Next
    Close #ff
    FileIsCrLf = False
End Function

Private Function ReadAllText(ByVal fpath As String) As String
    Dim ff As Integer: ff = FreeFile
    Open fpath For Input As #ff
    If LOF(ff) > 0 Then ReadAllText = Input$(LOF(ff), ff)
    Close #ff
End Function

' Izvuci editabilno telo iz VBA export fajla (.bas/.cls/.frm/.doccls) bezbedno za
' AddFromString: preskoci VERSION, Begin..End dizajn blok i SVE "Attribute" linije
' (ni modulske ni clanske ne smeju u CodeModule).
Private Function ExtractModuleCode(ByVal fpath As String) As String
    Dim allTxt As String, arr() As String, i As Long, depth As Long
    Dim inHeader As Boolean, ls As String, u As String, body As String

    allTxt = ReadAllText(fpath)
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
                ' modulski atribut -> skip
            ElseIf Len(Trim$(arr(i))) = 0 Then
                ' prazna linija u zaglavlju -> skip
            Else
                inHeader = False
                body = arr(i)
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
' POREDJENJE KODA I DETEKCIJA TVRDOG TELA
' (isti algoritam kao modSelfUpdate - tamo je Private, a taj modul se ne dira)
' ============================================================

' Ima li telo MODULE-LEVEL (pre prve procedure) "WithEvents" ili "As MSForms."
' deklaraciju. Procedure-level "Dim x As MSForms.ComboBox" NE cini modul tvrdim.
Private Function IsHardModuleBody(ByVal body As String) As Boolean
    Dim arr() As String, i As Long, u As String, w As String
    ' fajl daje vbCrLf, CodeModule.Lines daje lone vbCr -> normalizuj oba
    Dim t As String: t = Replace$(Replace$(body, vbCrLf, vbLf), vbCr, vbLf)
    arr = Split(t, vbLf)
    For i = 0 To UBound(arr)
        u = CodeLineUpper(arr(i))
        If Len(u) > 0 Then
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

' Kodni deo linije (bez trailing komentara), UPPER + LTrim - da "WithEvents" u
' komentaru ne da lazni pozitiv.
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

' Binarno poredjenje kanonizovanih tela.
Private Function SameCode(ByVal a As String, ByVal b As String) As Boolean
    SameCode = (StrComp(CanonCode(a), CanonCode(b), vbBinaryCompare) = 0)
End Function

' Kanonizacija za delta-skip: sve vrste prekida reda -> LF, VBE NBSP -> space,
' RTrim po redu, lowercase kod IZVAN stringova/komentara (VBE re-casing
' identifikatora), pa skini vodece i zavrsne prazne redove.
Private Function CanonCode(ByVal s As String) As String
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    s = Replace$(s, ChrW$(160), " ")
    Dim arr() As String, i As Long
    arr = Split(s, vbLf)
    For i = LBound(arr) To UBound(arr)
        arr(i) = RTrim$(LowerOutsideStrings(arr(i)))
    Next i
    s = Join(arr, vbLf)
    Do While Len(s) > 0
        If Left$(s, 1) = vbLf Then s = Mid$(s, 2) Else Exit Do
    Loop
    Do While Len(s) > 0
        If Right$(s, 1) = vbLf Then s = Left$(s, Len(s) - 1) Else Exit Do
    Loop
    CanonCode = s
End Function

' Lowercase kod izvan string-literala i komentara; case unutar "..." i posle '
' se CUVA (inace bi case-only izmena u stringu prosla kao "isto").
Private Function LowerOutsideStrings(ByVal s As String) As String
    ' NB: promenljiva se zove inQ - "inStr" bi se sudarilo sa ugradjenom InStr().
    Dim i As Long, n As Long, c As String, out As String, inQ As Boolean
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If inQ Then
            If c = """" Then
                If i < n And Mid$(s, i + 1, 1) = """" Then
                    out = out & """"""
                    i = i + 2
                Else
                    inQ = False
                    out = out & c
                    i = i + 1
                End If
            Else
                out = out & c
                i = i + 1
            End If
        Else
            If c = """" Then
                inQ = True
                out = out & c
                i = i + 1
            ElseIf c = "'" Then
                out = out & Mid$(s, i)
                Exit Do
            Else
                out = out & LCase$(c)
                i = i + 1
            End If
        End If
    Loop
    LowerOutsideStrings = out
End Function

' ============================================================
' SITNI HELPERI
' ============================================================

Private Function VBAProjectAccessible() As Boolean
    On Error Resume Next
    Dim n As Long
    n = ThisWorkbook.VBProject.VBComponents.count
    VBAProjectAccessible = (Err.Number = 0)
    On Error GoTo 0
    If Not VBAProjectAccessible Then
        MsgBox "Nema programskog pristupa VBA projektu." & vbCrLf & vbCrLf & _
               "Ukljuci: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
               "Macro Settings > 'Trust access to the VBA project object model'.", _
               vbExclamation, "VBA pristup"
    End If
End Function

' Fiksna putanja ako postoji; inace izbor foldera (masina bez te putanje).
Private Function ResolveFolder(ByVal fixedPath As String, ByVal caption As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(fixedPath) Then
        ResolveFolder = fixedPath
        Exit Function
    End If

    Dim fd As Object, p As String
    On Error Resume Next
    Set fd = Application.FileDialog(4)            ' msoFileDialogFolderPicker
    On Error GoTo 0
    If fd Is Nothing Then
        MsgBox "Folder ne postoji: " & fixedPath, vbExclamation, "modVbaTools"
        Exit Function
    End If
    fd.title = caption
    If fd.Show <> -1 Then Exit Function
    p = fd.SelectedItems(1)
    If Right$(p, 1) <> "\" Then p = p & "\"
    ResolveFolder = p
End Function

Private Function ComponentExists(ByVal proj As Object, ByVal baseName As String) As Boolean
    On Error Resume Next
    Dim c As Object
    Set c = proj.VBComponents(baseName)
    ComponentExists = Not (c Is Nothing)
    Err.Clear
End Function

' Kod komponente; okOut=False znaci da citanje nije uspelo (ne "prazan modul").
Private Function ComponentCode(ByVal vbc As Object, ByRef okOut As Boolean) As String
    Dim s As String
    okOut = False
    On Error Resume Next
    Err.Clear
    If vbc.CodeModule.CountOfLines > 0 Then s = vbc.CodeModule.Lines(1, vbc.CodeModule.CountOfLines)
    okOut = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    ComponentCode = s
End Function

' Broj kontrola u DIZAJNERU forme (ne u zivoj instanci). okOut=False = necitljivo.
Private Function DesignerControlCount(ByVal vbc As Object, ByRef okOut As Boolean) As Long
    Dim n As Long
    okOut = False
    On Error Resume Next
    Err.Clear
    n = vbc.Designer.Controls.count
    okOut = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    DesignerControlCount = n
End Function

Private Function TypeForExt(ByVal ext As String) As Long
    Select Case LCase$(ext)
        Case "bas":    TypeForExt = 1
        Case "cls":    TypeForExt = 2
        Case "frm":    TypeForExt = 3
        Case "doccls": TypeForExt = 100
        Case Else:     TypeForExt = 0
    End Select
End Function

Private Function ExtForType(ByVal t As Long) As String
    Select Case t
        Case 1:    ExtForType = ".bas"
        Case 2:    ExtForType = ".cls"
        Case 3:    ExtForType = ".frm"
        Case 100:  ExtForType = ".doccls"
        Case Else: ExtForType = ""
    End Select
End Function

Private Function BaseNameOf(ByVal fileName As String) As String
    Dim p As Long: p = InStrRev(fileName, ".")
    If p > 1 Then BaseNameOf = Left$(fileName, p - 1) Else BaseNameOf = fileName
End Function

Private Function ExtOf(ByVal fileName As String) As String
    Dim p As Long: p = InStrRev(fileName, ".")
    If p > 0 Then ExtOf = Mid$(fileName, p + 1)
End Function

Private Sub AddCsv(ByRef csv As String, ByVal item As String)
    If Len(csv) > 0 Then csv = csv & ","
    csv = csv & item
End Sub

Private Function CsvHas(ByVal csv As String, ByVal item As String) As Boolean
    CsvHas = (InStr(1, "," & csv & ",", "," & item & ",", vbTextCompare) > 0)
End Function

Private Function CsvCount(ByVal csv As String) As Long
    If Len(csv) = 0 Then Exit Function
    CsvCount = UBound(Split(csv, ",")) + 1
End Function

Private Function Cap(ByVal s As String, ByVal n As Long) As String
    If Len(s) <= n Then Cap = s Else Cap = Left$(s, n) & "..."
End Function

' ============================================================
' IZVESTAJI
' ============================================================

Private Sub ShowImportSuccess(ByVal summary As String, ByVal bkPath As String)
    MsgBox "ImportAllVBA je zavrsen." & vbCrLf & vbCrLf & _
           Cap(summary, 1400) & vbCrLf & vbCrLf & _
           "Backup: " & bkPath & vbCrLf & vbCrLf & _
           "SLEDECE (rucno, tim redom):" & vbCrLf & _
           "  1. Alt+F11 -> Debug -> Compile VBAProject" & vbCrLf & _
           "  2. ako je cisto -> Save" & vbCrLf & _
           "  3. zatvori pa ponovo otvori fajl" & vbCrLf & vbCrLf & _
           "Alat NE snima svesku - rucni Compile je zavrsna kapija.", _
           vbInformation, "ImportAllVBA"
End Sub

Private Sub ShowImportFailure(ByVal msg As String, ByVal bkPath As String)
    Dim rb As String
    If Len(mRbFail) > 0 Then
        rb = vbCrLf & "ROLLBACK NIJE USPEO za:" & vbCrLf & mRbFail & _
             "Te komponente su ostale BEZ ispravnog koda." & vbCrLf
    End If
    MsgBox "ImportAllVBA NIJE uspeo." & vbCrLf & vbCrLf & _
           Cap(msg, 1200) & vbCrLf & rb & vbCrLf & _
           "NE SNIMAJ RADNU SVESKU." & vbCrLf & _
           "Zatvori je BEZ snimanja; na disku je ispravna verzija." & vbCrLf & _
           IIf(Len(bkPath) > 0, "Backup: " & bkPath, "Backup nije napravljen."), _
           vbCritical, "ImportAllVBA"
End Sub
