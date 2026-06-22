Attribute VB_Name = "modPregledListova"
' ============================================================
' modPregledListova - pravi / azurira list "Pregled listova":
'   kolona A = naziv lista, kolona B = klikabilan link ka tom listu.
' Pokrece se RUCNO (Alt+F8 -> NapraviPregledListova). Bezbedno se moze
' pokretati vise puta - sadrzaj se svaki put iznova generise.
' ============================================================
Option Explicit

Public Const PREGLED_SHEET As String = "Pregled listova"

' Glavni ulaz: regenerise "Pregled listova" iz svih radnih listova.
Public Sub NapraviPregledListova()
    Dim wsPregled As Worksheet
    Set wsPregled = GetOrCreatePregledSheet()

    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' Ocisti prethodni sadrzaj (sam list ostaje).
    wsPregled.Cells.Clear

    ' Zaglavlje.
    wsPregled.Range("A1").Value = "Naziv lista"
    wsPregled.Range("B1").Value = "Link"
    wsPregled.Range("A1:B1").Font.Bold = True

    ' Jedan red po listu (preskoci sam "Pregled listova").
    Dim ws As Worksheet, r As Long
    r = 2
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.name, PREGLED_SHEET, vbTextCompare) <> 0 Then
            ' Kolona A: naziv lista.
            wsPregled.Cells(r, 1).Value = ws.name
            ' Kolona B: klikabilan link -> celija A1 ciljnog lista.
            wsPregled.Hyperlinks.Add _
                Anchor:=wsPregled.Cells(r, 2), _
                Address:="", _
                SubAddress:=SubAdresaLista(ws.name), _
                TextToDisplay:=ws.name
            r = r + 1
        End If
    Next ws

    ' Kozmetika: zamrznut header + autofit kolona.
    wsPregled.Columns("A:B").AutoFit
    wsPregled.Activate
    wsPregled.Range("A2").Select
    ActiveWindow.FreezePanes = False
    wsPregled.Range("A2").Select
    ActiveWindow.FreezePanes = True

    Application.ScreenUpdating = True
    MsgBox "Pregled listova azuriran. Listova u pregledu: " & (r - 2), _
           vbInformation, "Pregled listova"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Greska pri pravljenju pregleda: " & Err.Description, _
           vbExclamation, "Pregled listova"
End Sub

' Vrati postojeci "Pregled listova" ili ga napravi kao prvi list u workbook-u.
Private Function GetOrCreatePregledSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreatePregledSheet = ThisWorkbook.Worksheets(PREGLED_SHEET)
    On Error GoTo 0

    If GetOrCreatePregledSheet Is Nothing Then
        Set GetOrCreatePregledSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        GetOrCreatePregledSheet.name = PREGLED_SHEET
    End If
End Function

' SubAddress za interni hyperlink: naziv lista u apostrofima (radi i kad ima
' razmak), a apostrof u nazivu se udvaja po Excel pravilu.
Private Function SubAdresaLista(ByVal sheetName As String) As String
    SubAdresaLista = "'" & Replace(sheetName, "'", "''") & "'!A1"
End Function
