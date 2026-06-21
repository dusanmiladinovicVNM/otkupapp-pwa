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

Public Sub ExportAllVBA()
    Const FOLDER As String = "C:\put\do\src-vba\"        ' <-- IZMENI ("\" na kraju)
    If Not VBAProjectAccessible() Then Exit Sub
    If Dir(FOLDER, vbDirectory) = "" Then
        MsgBox "Folder ne postoji: " & FOLDER, vbExclamation, "ExportAllVBA": Exit Sub
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
            vbc.Export FOLDER & vbc.name & ext
            n = n + 1
        End If
    Next vbc
    MsgBox "Eksportovano komponenti: " & n & vbCrLf & FOLDER, vbInformation, "ExportAllVBA"
End Sub

Public Sub ImportAllVBA()
    Const FOLDER As String = "C:\put\do\src-vba\"        ' <-- IZMENI
    If Not VBAProjectAccessible() Then Exit Sub

    Dim proj As Object: Set proj = ThisWorkbook.VBProject
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FOLDER) Then
        MsgBox "Folder ne postoji: " & FOLDER, vbExclamation, "ImportAllVBA": Exit Sub
    End If

    Dim fil As Object, ext As String, baseName As String
    Dim vbc As Object, t As Long, imported As Long, skipped As String

    ' 1) document moduli (.doccls): kod se MERGE-uje u postojecu komponentu
    For Each fil In fso.GetFolder(FOLDER).Files
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
    For Each fil In fso.GetFolder(FOLDER).Files
        ext = LCase$(fso.GetExtensionName(fil.name))
        baseName = fso.GetBaseName(fil.name)
        If (ext = "bas" Or ext = "cls" Or ext = "frm") _
           And StrComp(baseName, SELF_MODULE, vbTextCompare) <> 0 Then

            If ext = "frm" And Not fso.FileExists(FOLDER & baseName & ".frx") Then
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
