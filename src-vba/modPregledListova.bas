Attribute VB_Name = "modPregledListova"
' ============================================================
' modPregledListova - pravi / azurira list "Pregled listova":
'   kolona A = naziv lista, kolona B = klikabilan link ka tom listu.
' Iznad tabele su dva dugmeta:
'   "Pokreni program" -> otvara glavni ekran (frmOtkupAPP.Show)
'   "Otvori VBA"      -> otvara VBA editor
' Pokrece se RUCNO (Alt+F8 -> NapraviPregledListova). Bezbedno se moze
' pokretati vise puta - sadrzaj i dugmad se svaki put iznova generisu.
' ============================================================
Option Explicit

Public Const PREGLED_SHEET As String = "Pregled listova"

' Red zaglavlja i prvi red podataka (redovi 1-2 su rezervisani za dugmad).
Private Const HDR_ROW As Long = 3
Private Const DATA_ROW As Long = 4

' Glavni ulaz: regenerise "Pregled listova" iz svih radnih listova.
Public Sub NapraviPregledListova()
    Dim wsPregled As Worksheet
    Set wsPregled = GetOrCreatePregledSheet()

    Application.ScreenUpdating = False
    On Error GoTo Fail

    ' Ocisti prethodni sadrzaj i stara dugmad (Cells.Clear ne brise dugmad).
    wsPregled.Cells.Clear
    ObrisiDugmad wsPregled

    ' Zaglavlje tabele.
    wsPregled.Cells(HDR_ROW, 1).Value = "Naziv lista"
    wsPregled.Cells(HDR_ROW, 2).Value = "Link"
    wsPregled.Range(wsPregled.Cells(HDR_ROW, 1), wsPregled.Cells(HDR_ROW, 2)).Font.Bold = True

    ' Jedan red po listu (preskoci sam "Pregled listova").
    Dim ws As Worksheet, r As Long
    r = DATA_ROW
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

    ' Dugmad iznad tabele + kozmetika (autofit, zamrznut band sa dugmadima/headerom).
    wsPregled.Columns("A:B").AutoFit
    DodajDugmad wsPregled
    wsPregled.Activate
    wsPregled.Range("A" & DATA_ROW).Select
    ActiveWindow.FreezePanes = False
    wsPregled.Range("A" & DATA_ROW).Select
    ActiveWindow.FreezePanes = True

    Application.ScreenUpdating = True
    MsgBox "Pregled listova azuriran. Listova u pregledu: " & (r - DATA_ROW), _
           vbInformation, "Pregled listova"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Greska pri pravljenju pregleda: " & Err.description, _
           vbExclamation, "Pregled listova"
End Sub

' --- Dugmad: akcije (Public, da ih OnAction moze pozvati) -------------------

' "Pokreni program" -> udji u glavni ekran. Init/licenca/splash su vec
' odradjeni pri otvaranju fajla (StartApp), pa ovde samo prikazujemo glavnu
' formu, isto kao i ostali pozivaci (frmMarza, frmIzvestaj, frmMaticniPodaci...).
Public Sub PokreniProgram()
    On Error GoTo Fail
    frmOtkupAPP.Show
    Exit Sub
Fail:
    MsgBox "Ne mogu da otvorim program: " & Err.description, _
           vbExclamation, "Pregled listova"
End Sub

' "Otvori VBA" -> otvori VBA editor. Prvo cist nacin (zahteva 'Trust access to
' the VBA project object model'); ako nije dozvoljen, fallback na Alt+F11.
Public Sub OtvoriVBA()
    On Error Resume Next
    Application.VBE.MainWindow.Visible = True
    If Err.Number <> 0 Then
        Err.Clear
        Application.SendKeys "%{F11}", True
    End If
    On Error GoTo 0
End Sub

' --- Helperi ----------------------------------------------------------------

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

' Ukloni sva (stara) dugmad sa lista - poziva se pre ponovnog crtanja.
Private Sub ObrisiDugmad(ByVal ws As Worksheet)
    On Error Resume Next
    Do While ws.Buttons.count > 0
        ws.Buttons(1).Delete
    Loop
    On Error GoTo 0
End Sub

' Nacrtaj dva dugmeta u "bandu" iznad tabele (redovi 1-2).
Private Sub DodajDugmad(ByVal ws As Worksheet)
    Dim topPt As Double, hPt As Double, wPt As Double, leftPt As Double
    leftPt = ws.Range("A1").Left
    topPt = ws.Range("A1").Top + 2
    wPt = 120
    hPt = 26

    Dim btn As Button
    Set btn = ws.Buttons.Add(leftPt, topPt, wPt, hPt)
    btn.Caption = "Pokreni program"
    btn.OnAction = "PokreniProgram"

    Set btn = ws.Buttons.Add(leftPt + wPt + 8, topPt, wPt, hPt)
    btn.Caption = "Otvori VBA"
    btn.OnAction = "OtvoriVBA"
End Sub
