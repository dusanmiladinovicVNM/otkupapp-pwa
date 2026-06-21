Attribute VB_Name = "modMigracija"
' ============================================================
' modMigracija - jednokratna migracija CISTIH PODATAKA iz starog
' OtkupApp fajla u novi (prazan). Mapiranje PO IMENU kolone,
' samo vrednosti (bez formula i koda). Ne treba "Trust access to
' VBA project object model" (koristi obican Excel objektni model).
'
' Upotreba: u NOVOM fajlu  ->  Alt+F8  ->  MigrirajPodatkeIzStarog
' ============================================================
Option Explicit

Public Sub MigrirajPodatkeIzStarog()
    Dim putanja As Variant
    putanja = Application.GetOpenFilename( _
        "OtkupApp fajlovi (*.xlsm;*.xlsb),*.xlsm;*.xlsb", , _
        "Izaberi STARI OtkupApp fajl (sa podacima)")
    If VarType(putanja) = vbBoolean Then Exit Sub      ' Cancel

    ' --- Zastita od slucajnog destruktivnog rerun-a ---
    Dim chk As ListObject
    Set chk = NadjiListObject(ThisWorkbook, "tblOtkup")
    If Not chk Is Nothing Then
        If chk.ListRows.count > 0 Then
            If MsgBox("Ovaj fajl VEC ima podatke (tblOtkup: " & chk.ListRows.count & _
                      " redova). Migracija ce ih PREPISATI." & vbCrLf & vbCrLf & _
                      "Nastaviti?", vbExclamation + vbYesNo + vbDefaultButton2, _
                      "Migracija podataka") = vbNo Then Exit Sub
        End If
    End If

    Dim novi As Workbook: Set novi = ThisWorkbook
    Dim stari As Workbook
    Dim prevEvents As Boolean, prevCalc As XlCalculation, prevSec As Long, prevSU As Boolean
    Dim summary As String, total As Long, tbls As Long

    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevSec = Application.AutomationSecurity
    prevSU = Application.ScreenUpdating

    On Error GoTo CLEAN
    Application.EnableEvents = False                    ' ne pokreci Workbook_Open starog
    Application.AutomationSecurity = 3                  ' msoAutomationSecurityForceDisable (bez makroa)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set stari = Workbooks.Open(Filename:=CStr(putanja), ReadOnly:=True, UpdateLinks:=0)

    Dim ws As Worksheet, loNovi As ListObject, n As Long
    For Each ws In novi.Worksheets
        For Each loNovi In ws.ListObjects
            If Not SkipTabela(loNovi.name) Then
                n = KopirajTabelu(stari, loNovi)
                If n < 0 Then
                    summary = summary & "  " & loNovi.name & " - (nema u starom)" & vbCrLf
                Else
                    total = total + n: tbls = tbls + 1
                    summary = summary & "  " & loNovi.name & " - " & n & " red." & vbCrLf
                End If
            End If
        Next loNovi
    Next ws

    stari.Close SaveChanges:=False: Set stari = Nothing

CLEAN:
    Dim em As String
    If Err.Number <> 0 Then em = "GRESKA: " & Err.description & vbCrLf & vbCrLf
    On Error Resume Next
    If Not stari Is Nothing Then stari.Close SaveChanges:=False
    Application.Calculation = prevCalc
    Application.AutomationSecurity = prevSec
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevSU
    On Error GoTo 0

    MsgBox em & "Tabela preneto: " & tbls & "   |   redova ukupno: " & total & _
           vbCrLf & vbCrLf & summary & vbCrLf & _
           "Sada SNIMI novi fajl (Ctrl+S) i proveri par tabela.", _
           IIf(Len(em) > 0, vbExclamation, vbInformation), "Migracija podataka"
End Sub

' Vrati broj prenetih redova; -1 ako tabele nema u starom fajlu.
Private Function KopirajTabelu(ByVal stari As Workbook, ByVal loNovi As ListObject) As Long
    Dim loStari As ListObject
    Set loStari = NadjiListObject(stari, loNovi.name)
    If loStari Is Nothing Then KopirajTabelu = -1: Exit Function

    ' ocisti eventualne postojece redove u novom (idempotentno)
    If Not loNovi.DataBodyRange Is Nothing Then loNovi.DataBodyRange.ClearContents
    If loStari.ListRows.count = 0 Then KopirajTabelu = 0: Exit Function

    ' mapiranje: novi col -> stari col, PO IMENU
    Dim nNew As Long: nNew = loNovi.ListColumns.count
    Dim mapCol() As Long: ReDim mapCol(1 To nNew)
    Dim j As Long, k As Long, staroIme As String
    For j = 1 To nNew
        staroIme = StaroImeKolone(loNovi.name, loNovi.ListColumns(j).name)
        For k = 1 To loStari.ListColumns.count
            If StrComp(Trim$(loStari.ListColumns(k).name), Trim$(staroIme), vbTextCompare) = 0 Then
                mapCol(j) = k: Exit For
            End If
        Next k
    Next j

    ' telo starog kao 2D niz (vrednosti, ne formule)
    Dim src As Variant: src = loStari.DataBodyRange.Value
    If Not IsArray(src) Then
        Dim one(1 To 1, 1 To 1) As Variant: one(1, 1) = src: src = one
    End If
    Dim nRows As Long: nRows = UBound(src, 1)

    ' prosiri novu tabelu na header + nRows, pa upisi SAMO mapirane kolone
    ' (nove/kalkulisane kolone u novom fajlu ostaju netaknute)
    loNovi.Resize loNovi.HeaderRowRange.Resize(nRows + 1, nNew)
    Dim colArr() As Variant, r As Long
    For j = 1 To nNew
        If mapCol(j) > 0 Then
            ReDim colArr(1 To nRows, 1 To 1)
            For r = 1 To nRows
                colArr(r, 1) = src(r, mapCol(j))
            Next r
            loNovi.ListColumns(j).DataBodyRange.Value = colArr
        End If
    Next j
    KopirajTabelu = nRows
End Function

Private Function NadjiListObject(ByVal wb As Workbook, ByVal naziv As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, naziv, vbTextCompare) = 0 Then
                Set NadjiListObject = lo: Exit Function
            End If
        Next lo
    Next ws
End Function

' Tabele koje se NE prenose:
'  - tblRpt*  : izvedeni izvestaji (regenerisu se)
'  - tblConfig / tblSEFConfig : masinski-specificni (OAuth/SEF/licenca) - OPT-IN
'  - tblMeteo : kes
' Ako zelis i config da prenese, izbaci ga iz liste.
Private Function SkipTabela(ByVal naziv As String) As Boolean
    If LCase$(Left$(naziv, 6)) = "tblrpt" Then SkipTabela = True: Exit Function
    Dim skip As Variant, i As Long
    skip = Array("tblConfig", "tblSEFConfig", "tblMeteo")
    For i = LBound(skip) To UBound(skip)
        If StrComp(naziv, CStr(skip(i)), vbTextCompare) = 0 Then SkipTabela = True: Exit Function
    Next i
End Function

' Override za PREIMENOVANE kolone izmedju verzija. Default = isto ime.
' Primer (ako je novo "Kontakt" u staroj verziji bilo "Telefon"):
'   If tabela = "tblStanice" And novoIme = "Kontakt" Then StaroImeKolone = "Telefon": Exit Function
Private Function StaroImeKolone(ByVal tabela As String, ByVal novoIme As String) As String
    StaroImeKolone = novoIme
End Function
