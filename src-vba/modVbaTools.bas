Attribute VB_Name = "modVbaTools"

' ============================================================
' modVbaTools - dev alat za EXPORT/IMPORT celog VBA koda iz/u src-vba folder.
' Round-trip: git source (src-vba) <-> Excel VBA projekat. Pokrece se RUCNO (Alt+F8).
'
' PREDUSLOV: File > Options > Trust Center > Trust Center Settings >
'            Macro Settings > "Trust access to the VBA project object model".
'
' Mapiranje ekstenzija:
'   1 StdModule -> .bas    2 ClassModule -> .cls
'   3 MSForm -> .frm(+.frx)    100 Document -> .doccls (ThisWorkbook / Sheet)
' ============================================================
Option Explicit

' Modul koji sadrzi OVAJ kod - preskace se pri importu (ne sme da se obrise
' dok se izvrsava). Ako preimenujes modul, promeni i ovu konstantu.
Private Const SELF_MODULE As String = "modVbaTools"

' Koliko duplikata najvise po jednom prolazu. Iz jednog pokretanja ne prodje
' proizvoljno mnogo brisanja (od 1506 trazenih ostalo je 125 neobrisanih), pa se
' radi u serijama -- pusti RemoveDuplicateModules dok ne javi 0 kandidata.
Private Const MAX_BRISANJA As Long = 200

' Radni folder sa izvorom (src-vba). Koriste ga ImportAllVBA i alat za duplikate.
Private Const SRC_FOLDER As String = "C:\Users\Dusan\Documents\GitHub\otkupapp-pwa\src-vba\"

Public Sub ExportAllVBA()
    Const folder As String = "C:\Users\Dusan\Desktop\AgriX\src-vbaExport23.06.2026_najnoviji\"        ' <-- IZMENI ("\" na kraju)
    If Not VBAProjectAccessible() Then Exit Sub
    If Dir(folder, vbDirectory) = "" Then
        MsgBox "Folder ne postoji: " & folder, vbExclamation, "ExportAllVBA": Exit Sub
    End If

    Dim vbc As Object, ext As String, n As Long
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        Select Case vbc.Type
            Case 1:    ext = ".bas"
            Case 2:    ext = ".cls"
            Case 3:    ext = ".frm"
            Case 100:  ext = ".doccls"
            Case Else: ext = ""
        End Select
        If Len(ext) > 0 Then
            vbc.Export folder & vbc.name & ext
            n = n + 1
        End If
    Next vbc
    MsgBox "Eksportovano komponenti: " & n & vbCrLf & folder, vbInformation, "ExportAllVBA"
End Sub

Public Sub ImportAllVBA()
    Dim folder As String
    folder = SRC_FOLDER                                    ' <-- IZMENI gore, u SRC_FOLDER
    If Not VBAProjectAccessible() Then Exit Sub

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then
        MsgBox "Folder ne postoji: " & folder, vbExclamation, "ImportAllVBA": Exit Sub
    End If

    Dim fil As Object, ext As String, baseName As String
    Dim vbc As Object, t As Long, imported As Long, skipped As String

    ' 1) document moduli (.doccls): kod se MERGE-uje u postojecu komponentu
    For Each fil In fso.GetFolder(folder).files
        If LCase$(fso.GetExtensionName(fil.name)) = "doccls" Then
            baseName = fso.GetBaseName(fil.name)
            Set vbc = Nothing
            On Error Resume Next
            Set vbc = proj.VBComponents(baseName)
            On Error GoTo 0
            If vbc Is Nothing Then
                skipped = skipped & "  " & fil.name & " (nema komponente '" & baseName & "')" & vbCrLf
            Else
                With vbc.CodeModule
                    If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                    .AddFromString ReadCodeBody(fil.path)
                End With
                imported = imported + 1
            End If
        End If
    Next fil

    ' 2) standardni / klasni / forme
    For Each fil In fso.GetFolder(folder).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        baseName = fso.GetBaseName(fil.name)
        If (ext = "bas" Or ext = "cls" Or ext = "frm") _
           And StrComp(baseName, SELF_MODULE, vbTextCompare) <> 0 Then

            If ext = "frm" And Not fso.FileExists(folder & baseName & ".frx") Then
                skipped = skipped & "  " & fil.name & " (nema .frx para)" & vbCrLf
            Else
                On Error Resume Next                         ' ukloni istoimenu (izbegni 'modX1')
                proj.VBComponents.Remove proj.VBComponents(baseName)
                On Error GoTo 0

                If FileHasVBHeader(fil.path) Then
                    proj.VBComponents.Import fil.path        ' header nosi ime
                    imported = imported + 1
                Else
                    t = 0
                    If ext = "bas" Then
                        t = 1
                    ElseIf ext = "cls" Then
                        t = 2
                    End If
                    If t > 0 Then
                        Set vbc = proj.VBComponents.Add(t)
                        vbc.name = baseName
                        vbc.CodeModule.AddFromFile fil.path
                        imported = imported + 1
                    Else
                        skipped = skipped & "  " & fil.name & " (forma bez headera)" & vbCrLf
                    End If
                End If
            End If
        End If
    Next fil

    MsgBox "Uvezeno komponenti: " & imported & vbCrLf & vbCrLf & _
           IIf(Len(skipped) > 0, "Preskoceno (rucno):" & vbCrLf & skipped, "Bez preskocenih.") & _
           vbCrLf & SELF_MODULE & " se ne uvozi (izvrsava se).", _
           vbInformation, "ImportAllVBA"
End Sub

' ============================================================
' CISCENJE DUPLIKATA (clsStmBtn1..125, clsSEFResponse1..99, frmX2 ...)
'
' Uzrok: u ImportAllVBA se prvo radi VBComponents.Remove pa Import. Za KLASE i
' FORME brisanje nije trenutno - ime ostaje zauzeto dok makro traje - pa Import
' napravi kopiju sa brojem na kraju. Svaki sledeci ImportAllVBA doda jos jednu.
'
' Pravilo brisanja je namerno usko, da ne obrise nista pravo. Komponenta X<broj>
' se brise SAMO ako:
'   1) NE postoji fajl X<broj>.bas/.cls/.frm u src-vba (nije praceni modul),
'   2) postoji fajl X.bas/.cls/.frm u src-vba,
'   3) komponenta X postoji u projektu (original je tu, kopija je visak).
' Document moduli (ThisWorkbook, Sheet1..Sheet25) se NIKAD ne diraju.
'
' Redosled: ListDuplicateModules (pregled) -> RemoveDuplicateModules (brisanje)
'           -> snimi, zatvori i otvori radnu svesku -> ListDuplicateModules opet.
' ============================================================

' Samo prebroji i ispisi, ne brise nista.
Public Sub ListDuplicateModules()
    ScanDuplicateModules False
End Sub

' Obrise duplikate (uz potvrdu).
Public Sub RemoveDuplicateModules()
    ScanDuplicateModules True
End Sub

Private Sub ScanDuplicateModules(ByVal doRemove As Boolean)
    If Not VBAProjectAccessible() Then Exit Sub

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(SRC_FOLDER) Then
        MsgBox "Folder ne postoji: " & SRC_FOLDER, vbExclamation, "Duplikati": Exit Sub
    End If

    ' 1) imena koja src-vba prati (bez ekstenzije)
    Dim tracked As Object: Set tracked = CreateObject("Scripting.Dictionary")
    tracked.CompareMode = vbTextCompare
    Dim fil As Object, ext As String
    For Each fil In fso.GetFolder(SRC_FOLDER).files
        ext = LCase$(fso.GetExtensionName(fil.name))
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            tracked(fso.GetBaseName(fil.name)) = True
        End If
    Next fil

    ' 2) klasifikacija komponenti u projektu
    Dim dupNames() As String, dupCount As Long
    ReDim dupNames(0 To proj.VBComponents.count)

    Dim dupList As String, orphanList As String, unknownList As String
    Dim orphanCount As Long, unknownCount As Long
    Dim vbc As Object, nm As String, baseName As String

    For Each vbc In proj.VBComponents
        nm = vbc.name
        baseName = TrimTrailingDigits(nm)
        ' modVbaTools i njegove kopije (modVbaTools1...) se NE diraju: odavde se
        ' kod izvrsava, a brisanje modula koji radi rusi Excel. Kopije tog modula
        ' obrisi rucno u VBE (desni klik > Remove), tek kad si van njega.
        If vbc.Type <> 100 And StrComp(baseName, SELF_MODULE, vbTextCompare) <> 0 _
           And Not tracked.Exists(nm) Then

            If Len(baseName) > 0 And Len(baseName) < Len(nm) And tracked.Exists(baseName) Then
                If ComponentExists(proj, baseName) Then
                    dupNames(dupCount) = nm
                    dupCount = dupCount + 1
                    dupList = dupList & nm & vbCrLf
                Else
                    orphanCount = orphanCount + 1
                    orphanList = orphanList & nm & " -> nema '" & baseName & "' u projektu" & vbCrLf
                End If
            Else
                unknownCount = unknownCount + 1
                unknownList = unknownList & nm & vbCrLf
            End If
        End If
    Next vbc

    ' 3) izvestaj u fajl pored radne sveske + rezime
    Dim reportPath As String
    reportPath = WriteDuplicateReport(fso, dupList, orphanList, unknownList)

    Dim summary As String
    summary = "Ukupno komponenti: " & proj.VBComponents.count & vbCrLf & _
              "Duplikati za brisanje: " & dupCount & vbCrLf & _
              "Bez originala (RUCNO proveriti): " & orphanCount & vbCrLf & _
              "Van src-vba, nije duplikat (ne dira se): " & unknownCount & vbCrLf & vbCrLf & _
              IIf(Len(reportPath) > 0, "Spisak: " & reportPath, "Spisak nije mogao da se upise.")
    Debug.Print summary

    If Not doRemove Then
        MsgBox summary, vbInformation, "Duplikati - samo pregled"
        Exit Sub
    End If

    If dupCount = 0 Then
        MsgBox "Nema duplikata za brisanje." & vbCrLf & vbCrLf & summary, vbInformation, "Duplikati"
        Exit Sub
    End If

    If MsgBox("Brisem " & dupCount & " komponenti iz VBA projekta." & vbCrLf & _
              "Izvor istine ostaje src-vba; originali se ne diraju." & vbCrLf & vbCrLf & _
              "SNIMI radnu svesku pre nego sto potvrdis." & vbCrLf & vbCrLf & _
              "Nastaviti?" & vbCrLf & vbCrLf & summary, _
              vbYesNo + vbDefaultButton2 + vbExclamation, "Brisanje duplikata") <> vbYes Then Exit Sub

    Dim countBefore As Long, countAfter As Long
    countBefore = proj.VBComponents.count

    Dim i As Long, removed As Long, failed As String, failedCount As Long
    For i = 0 To dupCount - 1
        On Error Resume Next
        Err.Clear
        proj.VBComponents.Remove proj.VBComponents(dupNames(i))
        If Err.Number <> 0 Then
            failedCount = failedCount + 1
            failed = failed & "  " & dupNames(i) & " (" & Err.Number & " " & Err.Description & ")" & vbCrLf
        Else
            removed = removed + 1
        End If
        On Error GoTo 0
        If removed >= MAX_BRISANJA Then Exit For
    Next i

    countAfter = proj.VBComponents.count
    Debug.Print "Obrisano: " & removed & ", neuspesno: " & failedCount & _
                "; komponenti pre/posle: " & countBefore & "/" & countAfter
    If Len(failed) > 0 Then Debug.Print failed

    Dim tail As String
    ' VBE stablo se ne osvezava posle brisanja -- komponente ostaju iscrtane iako
    ' ih u projektu nema. Merodavan je broj komponenti, ne Project Explorer.
    tail = "Komponenti pre/posle: " & countBefore & "/" & countAfter & vbCrLf & _
           "Ostalo kandidata: " & (dupCount - removed) & vbCrLf & vbCrLf & _
           "Project Explorer i dalje crta obrisane -- to je zastareo prikaz," & vbCrLf & _
           "veruj broju. Pusti RemoveDuplicateModules ponovo dok ne javi 0" & vbCrLf & _
           "kandidata, pa SNIMI svesku (bez snimanja se sve vraca)." & vbCrLf & vbCrLf & _
           "Ako broj komponenti stoji u mestu kroz dva prolaza, brisanje iznutra" & vbCrLf & _
           "ne prolazi -- zatvori Excel i pusti:" & vbCrLf & _
           "  powershell -File tools\clean_vba_duplicates.ps1 -Workbook ""<putanja.xlsm>"" -Apply"

    MsgBox "Obrisano (prijavljeno): " & removed & vbCrLf & _
           "Neuspesno: " & failedCount & _
           IIf(failedCount > 0, vbCrLf & vbCrLf & Left$(failed, 700) & _
               "(ceo spisak: Immediate prozor)", "") & vbCrLf & vbCrLf & tail, _
           vbInformation, "Brisanje duplikata"
End Sub

' "clsStmBtn125" -> "clsStmBtn"; "modOtkup" -> "modOtkup" (nista da skine).
Private Function TrimTrailingDigits(ByVal s As String) As String
    Dim i As Long
    i = Len(s)
    Do While i > 0
        If InStr("0123456789", Mid$(s, i, 1)) = 0 Then Exit Do
        i = i - 1
    Loop
    TrimTrailingDigits = Left$(s, i)
End Function

Private Function ComponentExists(ByVal proj As Object, ByVal compName As String) As Boolean
    Dim vbc As Object
    On Error Resume Next
    Set vbc = proj.VBComponents(compName)
    On Error GoTo 0
    ComponentExists = Not (vbc Is Nothing)
End Function

' Vrati putanju izvestaja, ili "" ako upis nije uspeo.
Private Function WriteDuplicateReport(ByVal fso As Object, ByVal dupList As String, _
                                      ByVal orphanList As String, ByVal unknownList As String) As String
    Dim path As String, ts As Object
    path = ThisWorkbook.path & "\vba_duplikati.txt"
    On Error GoTo Fail
    Set ts = fso.CreateTextFile(path, True)
    ts.WriteLine "Duplikati - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ts.WriteLine ""
    ts.WriteLine "[ZA BRISANJE]"
    ts.WriteLine IIf(Len(dupList) > 0, dupList, "  (nema)")
    ts.WriteLine "[BEZ ORIGINALA - rucna provera]"
    ts.WriteLine IIf(Len(orphanList) > 0, orphanList, "  (nema)")
    ts.WriteLine "[VAN src-vba - ne dira se]"
    ts.WriteLine IIf(Len(unknownList) > 0, unknownList, "  (nema)")
    ts.Close
    WriteDuplicateReport = path
    Exit Function
Fail:
    WriteDuplicateReport = ""
End Function

' ---------------- helperi ----------------

Private Function VBAProjectAccessible() As Boolean
    On Error Resume Next
    Dim c As Long
    c = ThisWorkbook.VBProject.VBComponents.count
    VBAProjectAccessible = (Err.Number = 0)
    On Error GoTo 0
    If Not VBAProjectAccessible Then
        MsgBox "Nema programskog pristupa VBA projektu." & vbCrLf & vbCrLf & _
               "Ukljuci: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
               "Macro Settings > 'Trust access to the VBA project object model'.", _
               vbExclamation, "VBA pristup"
    End If
End Function

Private Function FileHasVBHeader(ByVal path As String) As Boolean
    Dim ff As Integer, s As String
    ff = FreeFile
    Open path For Input As #ff
    If Not EOF(ff) Then Line Input #ff, s
    Close #ff
    FileHasVBHeader = (InStr(1, s, "Attribute VB_Name", vbTextCompare) > 0) _
                   Or (UCase$(Left$(LTrim$(s), 7)) = "VERSION")
End Function

' Vrati samo kod (preskoci VBA header blok ako postoji) - za .doccls merge.
Private Function ReadCodeBody(ByVal path As String) As String
    Dim ff As Integer, s As String, ls As String, body As String, started As Boolean
    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, s
        If started Then
            body = body & vbCrLf & s
        Else
            ls = LTrim$(s)
            If Not (ls Like "VERSION*" Or ls = "BEGIN" Or ls = "END" _
                Or ls Like "MultiUse =*" Or ls Like "Attribute VB_*" Or Len(Trim$(s)) = 0) Then
                started = True: body = s
            End If
        End If
    Loop
    Close #ff
    ReadCodeBody = body
End Function
