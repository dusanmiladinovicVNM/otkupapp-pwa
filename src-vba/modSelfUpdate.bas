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
'   StartApp tada: Application.OnTime Now, "RunSelfUpdate" + Exit Sub
'     (ne importuje se u toku Workbook_Open/StartApp stack-a; tek na praznom
'      stack-u i bez otvorene glavne forme).
'   RunSelfUpdate(): backup -> download svih fajlova -> import -> save ->
'     poruka "zatvori i otvori" (restart pokrece InitApp/ValidateAllTables sa
'     novim kodom -> schema self-heal kroz postojeci put, bez rucne migracije).
'
' OPT-IN: radi samo ako je REL_FOLDER_ID postavljen (modConfig).
' FAIL-SOFT: svaka greska u proveri samo se loguje, start ide dalje.
' Reuse: modDrive (download/list), modUpdateGate.VersionCompare,
'        modGoogleAuth.ExtractJsonStringGoogle.
' VBA-pristup / import helperi su LOKALNI (Self*) - NE zovu modVbaTools, jer
' ImportAllVBA preskace modVbaTools (SELF_MODULE) pa njegov Public ne stigne na
' klijenta ("sub or function not defined"). Zato je modSelfUpdate samodovoljan.
'
' Skip pri importu: moduli koji su NA call stack-u (modSelfUpdate) + dev tool
' (modVbaTools, klijent ga nikad ne pokrece) -> njih ne diramo da se kod koji
' se izvrsava ne obrise. ThisWorkbook se MERGE-uje (ne brise).
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

    ' Preduslov: programski pristup VBA projektu (helper sam prijavi ako fali).
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

    ' 3) Import (skip moduli na stack-u; ThisWorkbook se merge-uje)
    Dim summary As String: summary = ImportFromFolder(tempDir, SKIP_MODULES)

    ' 4) Snimi (restart aktivira nov kod + schema self-heal kroz InitApp)
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

' Zameni KOD postojecih komponenti - NE Remove/Import. Remove+Import formi u
' runtime-u rusi projekat ("Errors during load" / korupcija). Kod svake
' komponente se vadi kroz ExtractModuleCode (strip header + sve Attribute
' linije, da AddFromString ne baci "Syntax error" na member atributima).
' Dizajn formi (.frx/kontrole) se NE menja - za izmene dizajna formi treba
' pun reinstall .xlsm.
Private Function ImportFromFolder(ByVal folder As String, ByVal skipCsv As String) As String
    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fil As Object, ext As String, baseName As String, vbc As Object
    Dim body As String, merged As Long, added As Long, skippedNew As String, t As Long
    Dim skip As String: skip = "," & LCase$(skipCsv) & ","

    For Each fil In fso.GetFolder(folder).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        baseName = fso.GetBaseName(fil.name)
        Select Case ext
            Case "bas", "cls", "frm", "doccls"
                If InStr(skip, "," & LCase$(baseName) & ",") = 0 Then
                    body = ExtractModuleCode(fil.path)

                    Set vbc = Nothing
                    On Error Resume Next
                    Set vbc = proj.VBComponents(baseName)
                    On Error GoTo 0

                    If Not vbc Is Nothing Then
                        ' postoji -> zameni SAMO kod (bez Remove/Import)
                        With vbc.CodeModule
                            If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                            If Len(body) > 0 Then .AddFromString body
                        End With
                        merged = merged + 1
                    ElseIf ext = "bas" Or ext = "cls" Then
                        ' nov standardni/klasni modul -> dodaj (bez Remove)
                        t = IIf(ext = "bas", 1, 2)
                        Set vbc = proj.VBComponents.Add(t)
                        vbc.name = baseName
                        If Len(body) > 0 Then vbc.CodeModule.AddFromString body
                        added = added + 1
                    Else
                        ' nova forma / novi sheet -> ne pravimo u runtime-u (reinstall)
                        skippedNew = skippedNew & "  " & fil.name & vbCrLf
                    End If
                End If
        End Select
    Next fil

    ImportFromFolder = "Azurirano (kod): " & merged & ", novih modula: " & added & _
        IIf(Len(skippedNew) > 0, vbCrLf & "Preskoceno (novo, treba reinstall):" & vbCrLf & skippedNew, "")
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff
    If LOF(ff) > 0 Then ReadAllText = Input$(LOF(ff), ff)
    Close #ff
End Function

' --- lokalni helperi (kopija logike iz modVbaTools; samodovoljno jer se
'     modVbaTools ne update-uje na klijentu - SELF_MODULE skip u ImportAllVBA) ---

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

