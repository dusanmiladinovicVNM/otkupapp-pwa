Attribute VB_Name = "modPregledListova"
'Attribute VB_Name = "modPregledListova"
' ============================================================
' modPregledListova - pravi / azurira list "Pregled listova":
'   kolona A = naziv lista, kolona B = klikabilan link ka tom listu.
' Iznad tabele su dugmici:
'   "Pokreni program" -> otvara ljusku (modOtkupUI.ShowOtkupUI)
'   "Otvori VBA"      -> otvara VBA editor
'   "Migracije"       -> MigrirajPodatkeIzStarog (uvoz iz starog fajla)
'   "Ocisti tabele"   -> brise SAMO unose iz definisanih tabela (header+tbl ostaju)
'   "Uvezi VBA"       -> modVbaTools.ImportAllVBA (usaglasi projekat sa src-vba)
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
    On Error GoTo fail

    ' Ocisti prethodni sadrzaj i stara dugmad (Cells.Clear ne brise dugmad).
    wsPregled.cells.Clear
    ObrisiDugmad wsPregled

    ' Zaglavlje tabele.
    wsPregled.cells(HDR_ROW, 1).value = "Naziv lista"
    wsPregled.cells(HDR_ROW, 2).value = "Link"
    wsPregled.Range(wsPregled.cells(HDR_ROW, 1), wsPregled.cells(HDR_ROW, 2)).Font.Bold = True

    ' Jedan red po listu (preskoci sam "Pregled listova").
    Dim ws As Worksheet, r As Long
    r = DATA_ROW
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.name, PREGLED_SHEET, vbTextCompare) <> 0 Then
            ' Kolona A: naziv lista.
            wsPregled.cells(r, 1).value = ws.name
            ' Kolona B: klikabilan link -> celija A1 ciljnog lista.
            wsPregled.Hyperlinks.Add _
                anchor:=wsPregled.cells(r, 2), _
                Address:="", _
                SubAddress:=SubAdresaLista(ws.name), _
                TextToDisplay:=ws.name
            r = r + 1
        End If
    Next ws

    ' Dugmad iznad tabele + kozmetika (autofit, zamrznut band sa dugmadima/headerom).
    wsPregled.columns("A:B").AutoFit
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

fail:
    Application.ScreenUpdating = True
    MsgBox "Gre" & ChrW(353) & "ka pri pravljenju pregleda: " & Err.description, _
           vbExclamation, "Pregled listova"
End Sub

' --- Dugmad: akcije (Public, da ih OnAction moze pozvati) -------------------

' "Pokreni program" -> udji u glavni ekran. Init/licenca/splash su vec
' odradjeni pri otvaranju fajla (StartApp), pa ovde samo prikazujemo ljusku,
' isto sto StartApp radi posle splash faze (modUiFaze.FazaBoot, v6-ui-214).
Public Sub PokreniProgram()
    On Error GoTo fail
    modOtkupUI.ShowOtkupUI
    Exit Sub
fail:
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

' "Migracije" -> jednokratni uvoz cistih podataka iz starog OtkupApp fajla.
' Reuse modMigracija.MigrirajPodatkeIzStarog (sam bira fajl i ima svoju potvrdu).
Public Sub PokreniMigraciju()
    On Error GoTo fail
    MigrirajPodatkeIzStarog
    Exit Sub
fail:
    MsgBox "Gre" & ChrW(353) & "ka pri migraciji: " & Err.description, _
           vbExclamation, "Migracije"
End Sub

' "Uvezi VBA" -> usaglasi VBA projekat sa src-vba folderom (dev alat).
'
' Namerno BEZ dodatne potvrde: ImportAllVBA sam prikazuje PLAN (sta se menja, sta
' se uklanja) i pita "Nastaviti?", pravi backup pre prve izmene i odbija posao ako
' izvor nije ispravan. Jos jedan MsgBox bi dodao klik bez ijedne nove garancije.
'
' Na masini bez dev putanje (SRC_FOLDER ne postoji) alat otvara izbor foldera, a
' preflight odbija sve sto nije pravi src-vba (mora imati modVbaTools.bas, tacno
' jednu .frm i njen .frx). Zato dugme ne moze da naskodi pogresnim klikom.
Public Sub UveziVBA()
    ImportAllVBA
End Sub

' "Ocisti tabele" -> brise SAMO unose (DataBodyRange) iz dole navedenih tabela;
' zaglavlja i sami ListObject-i ostaju. Trazi da operater ukuca "OBRISI" (NE moze
' undo) -- InputBox umesto Da/Ne da se izbegne slucajno brisanje jednim klikom.
' Napomena: "otkupna mesta" = tblStanice (nema zasebne tabele).
Public Sub OcistiTabele()
    Dim tbls As Variant
    tbls = Array( _
        TBL_OTKUP, TBL_OTPREMNICA, TBL_ZBIRNA, TBL_PRIJEMNICA, _
        TBL_FAKTURE, TBL_FAKTURA_STAVKE, TBL_AMBALAZA, _
        TBL_PALETA, TBL_PALETA_STAVKA, TBL_PRERADA, TBL_PRERADA_STAVKA, _
        TBL_ARTIKLI, TBL_MAGACIN, TBL_NOVAC, TBL_KUPCI, TBL_STANICE, _
        TBL_KULTURE, TBL_TIP_AMBALAZE, TBL_TIP_PALETE, TBL_KOOPERANTI, _
        TBL_VOZACI, TBL_CENOVNIK)

    Dim n As Long
    n = UBound(tbls) - LBound(tbls) + 1

    Dim odg As String
    odg = InputBox("Obrisati SVE unose iz " & n & " tabela?" & vbCrLf & _
                   "(zaglavlja i tabele ostaju; radnja se NE mo" & ChrW(382) & "e opozvati)" & vbCrLf & vbCrLf & _
                   "Za potvrdu ukucajte: OBRI" & ChrW(352) & "I", "Ocisti tabele")
    If Not PotvrdaObrisi(odg) Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo fail

    Dim i As Long, ocisceno As Long, nedostaju As String
    For i = LBound(tbls) To UBound(tbls)
        If ObrisiUnoseTabele(CStr(tbls(i))) Then
            ocisceno = ocisceno + 1
        Else
            nedostaju = nedostaju & "  - " & CStr(tbls(i)) & vbCrLf
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Ocisceno tabela: " & ocisceno & " / " & n & _
           IIf(Len(nedostaju) > 0, vbCrLf & vbCrLf & "Nisu pronadjene:" & vbCrLf & nedostaju, ""), _
           vbInformation, "Ocisti tabele"
    Exit Sub

fail:
    Application.ScreenUpdating = True
    MsgBox "Gre" & ChrW(353) & "ka pri ciscenju tabela: " & Err.description, _
           vbExclamation, "Ocisti tabele"
End Sub

' Potvrda brisanja: tacno "OBRISI" (sa ili bez dijakritike na S, nezavisno od
' velicine slova). Prazno / Cancel / bilo sta drugo = False.
Private Function PotvrdaObrisi(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    t = Replace$(t, ChrW(353), "s")   ' s-caron -> s
    t = Replace$(t, ChrW(352), "S")   ' S-caron -> S
    PotvrdaObrisi = (UCase$(t) = "OBRISI")
End Function

' --- Helperi ----------------------------------------------------------------

' Vrati postojeci "Pregled listova" ili ga napravi kao prvi list u workbook-u.
Private Function GetOrCreatePregledSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreatePregledSheet = ThisWorkbook.Worksheets(PREGLED_SHEET)
    On Error GoTo 0

    If GetOrCreatePregledSheet Is Nothing Then
        Set GetOrCreatePregledSheet = ThisWorkbook.Worksheets.Add(before:=ThisWorkbook.Worksheets(1))
        GetOrCreatePregledSheet.name = PREGLED_SHEET
    End If
End Function

' SubAddress za interni hyperlink: naziv lista u apostrofima (radi i kad ima
' razmak), a apostrof u nazivu se udvaja po Excel pravilu.
Private Function SubAdresaLista(ByVal sheetName As String) As String
    SubAdresaLista = "'" & Replace(sheetName, "'", "''") & "'!A1"
End Function

' Obrisi SAMO unose iz jedne tabele (header + ListObject ostaju). Vrati True
' ako je tabela pronadjena. Reuse: GetTable + ustaljeni DataBodyRange.Delete.
Private Function ObrisiUnoseTabele(ByVal tblName As String) As Boolean
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then Exit Function

    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    On Error GoTo 0
    ObrisiUnoseTabele = True
End Function

' Ukloni sva (stara) dugmad sa lista - poziva se pre ponovnog crtanja.
Private Sub ObrisiDugmad(ByVal ws As Worksheet)
    On Error Resume Next
    Do While ws.buttons.count > 0
        ws.buttons(1).Delete
    Loop
    On Error GoTo 0
End Sub

' Nacrtaj dugmice u "bandu" iznad tabele (redovi 1-2), jedan pored drugog.
Private Sub DodajDugmad(ByVal ws As Worksheet)
    Dim specs As Variant
    specs = Array( _
        Array("Pokreni program", "PokreniProgram"), _
        Array("Otvori VBA", "OtvoriVBA"), _
        Array("Migracije", "PokreniMigraciju"), _
        Array("Ocisti tabele", "OcistiTabele"), _
        Array("Uvezi VBA", "UveziVBA"))

    Dim leftPt As Double, topPt As Double, wPt As Double, hPt As Double, gap As Double
    leftPt = ws.Range("A1").Left
    topPt = ws.Range("A1").top + 2
    wPt = 120: hPt = 26: gap = 8

    Dim i As Long, btn As Button
    For i = LBound(specs) To UBound(specs)
        Set btn = ws.buttons.Add(leftPt + i * (wPt + gap), topPt, wPt, hPt)
        btn.caption = specs(i)(0)
        btn.OnAction = specs(i)(1)
    Next i
End Sub

