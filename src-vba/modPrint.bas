Attribute VB_Name = "modPrint"
Option Explicit

' ============================================================
' modPrint – Druckausgabe (ersetzt direkte PrintOut-Aufrufe)
' ============================================================

Public Sub PrintIzvestaj(ByVal data As Variant, ByVal reportTitle As String, _
                         ByVal headers As Variant)
    ' Generischer Report-Druck
    ' Schreibt in ein temporäres Print-Sheet und druckt
    
    Dim wsPrint As Worksheet
    On Error Resume Next
    Set wsPrint = ThisWorkbook.Sheets("_Print")
    On Error GoTo 0
    
    If wsPrint Is Nothing Then
        Set wsPrint = ThisWorkbook.Sheets.Add
        wsPrint.name = "_Print"
    End If
    
    wsPrint.cells.Clear
    
    ' Titel
    wsPrint.Range("A1").value = reportTitle
    wsPrint.Range("A1").Font.Size = 14
    wsPrint.Range("A1").Font.Bold = True
    
    ' Daten ausgeben
    OutputToSheet data, wsPrint.Range("A3"), headers
    
    ' Drucken
    wsPrint.PrintOut Copies:=1
    
    ' Aufräumen
    wsPrint.Visible = xlSheetVeryHidden
End Sub

' ============================================================
' OTKUPNI LIST (zakonski) — OtkupSablon, dva primerka jedan iznad drugog,
' A4 portrait. PDV nadoknada se racuna (CFG_PDV_NADOKNADA_STOPA, default 8%).
' Izlaz po CFG_OTKUP_PRINT_MODE: (prazno/PDF) | PRINT | PREVIEW | OFF.
' otkupIDs = rezultat SaveOtkupMulti_TX (npr. "OTK-1 + OTK-2" ili "OTK-1").
' ============================================================

' Implementira stari stub: pojedinacni otkupni list -> izlaz po modu.
Public Sub PrintOtkupniList(ByVal otkupID As String)
    OutputOtkupniList otkupID
End Sub

' Glavni ulaz (zove se posle SaveOtkupMulti_TX). Best-effort: greska se loguje.
Public Sub OutputOtkupniList(ByVal otkupIDs As String)
    On Error GoTo EH
    Dim mode As String
    mode = UCase$(Trim$(GetConfigValue(CFG_OTKUP_PRINT_MODE)))

    Select Case mode
        Case "OFF"
            ' bez izlaza
        Case "PRINT"
            Dim ws As Worksheet: Set ws = FillOtkupSablon(otkupIDs)
            If Not ws Is Nothing Then ws.PrintOut Copies:=1
        Case "PREVIEW"
            Dim wp As Worksheet: Set wp = FillOtkupSablon(otkupIDs)
            If Not wp Is Nothing Then wp.PrintPreview
        Case Else
            ExportOtkupniListPDF otkupIDs, True   ' default: tihi PDF
    End Select
    Exit Sub
EH:
    LogErr "modPrint.OutputOtkupniList"
End Sub

' PDF otkupnog lista -> <workbook>\OtkupniList_<brDok>.pdf
Public Function ExportOtkupniListPDF(ByVal otkupIDs As String, _
                                     Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = FillOtkupSablon(otkupIDs)
    If ws Is Nothing Then Exit Function

    Dim suff As String: suff = Replace(Replace(otkupIDs, " + ", "_"), "/", "-")
    Dim pdfPath As String: pdfPath = ThisWorkbook.Path & "\OtkupniList_" & suff & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, OpenAfterPublish:=openAfter
    ExportOtkupniListPDF = pdfPath
    Exit Function
EH:
    LogErr "modPrint.ExportOtkupniListPDF"
End Function

' Popuni OtkupSablon sa dva primerka. Vraca sheet (ili Nothing).
Private Function FillOtkupSablon(ByVal otkupIDs As String) As Worksheet
    On Error GoTo EH
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsureOtkupSablon
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("OtkupSablon")
    On Error GoTo EH
    If ws Is Nothing Then Exit Function

    Dim d As Variant: d = GetTableData(TBL_OTKUP)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long, iVr As Long, iSo As Long, iKl As Long, iKol As Long, iCe As Long
    Dim iKoop As Long, iSt As Long, iBr As Long, iDat As Long, iTip As Long, iKolAmb As Long
    iID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    iVr = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    iSo = GetColumnIndex(TBL_OTKUP, COL_OTK_SORTA)
    iKl = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    iKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    iCe = GetColumnIndex(TBL_OTKUP, COL_OTK_CENA)
    iKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    iSt = GetColumnIndex(TBL_OTKUP, COL_OTK_STANICA)
    iBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    iDat = GetColumnIndex(TBL_OTKUP, COL_OTK_DATUM)
    iTip = GetColumnIndex(TBL_OTKUP, COL_OTK_TIP_AMB)
    iKolAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)

    Dim ids() As String: ids = Split(otkupIDs, " + ")
    Dim stavke() As Variant: ReDim stavke(0 To UBound(ids), 0 To 4)
    Dim cnt As Long: cnt = 0
    Dim osnovica As Double: osnovica = 0
    ' Cena uneta u formi je BRUTO (vec sadrzi PDV nadoknadu). U otkupnom listu
    ' prikazujemo NETO cenu/vrednost, a PDV nadoknadu dodajemo kao posebnu stavku.
    Dim stopa As Double: stopa = PrNz(GetConfigValue(CFG_PDV_NADOKNADA_STOPA))
    If stopa <= 0 Then stopa = PDV_NADOKNADA_DEFAULT
    Dim koopID As String, stID As String, brDok As String, datum As String
    Dim tipAmb As String, kolAmb As Double
    Dim j As Long, r As Long
    For j = 0 To UBound(ids)
        Dim wantID As String: wantID = Trim$(ids(j))
        If wantID <> "" Then
            For r = 1 To UBound(d, 1)
                If CStr(d(r, iID)) = wantID Then
                    Dim kol As Double: kol = PrNz(d(r, iKol))
                    Dim cenBruto As Double: cenBruto = PrNz(d(r, iCe))
                    Dim cenNeto As Double: cenNeto = cenBruto / (1 + stopa / 100)
                    stavke(cnt, 0) = Trim$(CStr(d(r, iVr)) & " " & CStr(d(r, iSo)))
                    stavke(cnt, 1) = CStr(d(r, iKl))
                    stavke(cnt, 2) = kol
                    stavke(cnt, 3) = cenNeto
                    stavke(cnt, 4) = kol * cenNeto
                    osnovica = osnovica + kol * cenNeto
                    If koopID = "" Then
                        koopID = CStr(d(r, iKoop)): stID = CStr(d(r, iSt))
                        brDok = CStr(d(r, iBr)): datum = Format$(d(r, iDat), "dd.mm.yyyy")
                        tipAmb = CStr(d(r, iTip)): kolAmb = PrNz(d(r, iKolAmb))
                    End If
                    cnt = cnt + 1
                    Exit For
                End If
            Next r
        End If
    Next j
    If cnt = 0 Then Exit Function

    Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
    h("name") = GetConfigValue("SELLER_NAME")
    h("pib") = GetConfigValue("SELLER_PIB")
    h("mb") = GetConfigValue("SELLER_MATICNI_BROJ")
    h("addr") = Trim$(GetConfigValue("SELLER_STREET") & ", " & _
                GetConfigValue("SELLER_POSTAL_CODE") & " " & GetConfigValue("SELLER_CITY"))
    h("acct") = GetConfigValue("SELLER_ACCOUNT")
    h("brDok") = brDok
    h("datum") = datum
    h("stanica") = CStr(LookupValue(TBL_STANICE, "StanicaID", stID, "Naziv"))
    h("koop") = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, "Ime")) & " " & _
                CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, "Prezime")))
    h("bpg") = CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, COL_KOOP_BPG))
    h("racun") = CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, COL_KOOP_TEKUCI_RACUN))
    h("amb") = tipAmb & " x " & CStr(CLng(kolAmb))
    h("stopa") = stopa
    h("osnovica") = osnovica
    h("nadoknada") = osnovica * stopa / 100
    h("ukupno") = osnovica + osnovica * stopa / 100

    Application.ScreenUpdating = False
    On Error Resume Next
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ws.cells.UnMerge
    On Error GoTo EH
    ws.cells.Clear
    ws.cells.Font.name = "Calibri"
    ws.cells.Font.Size = 10
    ws.columns("A").ColumnWidth = 5
    ws.columns("B").ColumnWidth = 22
    ws.columns("C").ColumnWidth = 8
    ws.columns("D").ColumnWidth = 15
    ws.columns("E").ColumnWidth = 14
    ws.columns("F").ColumnWidth = 16

    Dim r0 As Long, lastRow As Long
    r0 = WriteOtkupCopy(ws, 2, "Primerak za poljoprivrednika", h, stavke, cnt)

    ' Razmak + isprekidana linija za secenje izmedju dva primerka.
    Dim cutRow As Long: cutRow = r0 + 1
    ws.Range(ws.cells(cutRow, 1), ws.cells(cutRow, 6)).Merge
    With ws.cells(cutRow, 1)
        .value = ChrW(9986) & " " & String$(80, "-")
        .Font.Color = RGB(150, 150, 150)
        .HorizontalAlignment = xlCenter
    End With

    lastRow = WriteOtkupCopy(ws, cutRow + 2, "Primerak za otkupljivaca", h, stavke, cnt)

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .CenterHorizontally = True
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(lastRow, 6)).Address
    End With
    Application.PrintCommunication = True
    On Error GoTo 0

    Application.ScreenUpdating = oldScreen
    Set FillOtkupSablon = ws
    Exit Function
EH:
    Application.ScreenUpdating = oldScreen
    LogErr "modPrint.FillOtkupSablon"
End Function

' Ispisuje jedan stilizovani primerak od reda r0; vraca prvi slobodan red.
Private Function WriteOtkupCopy(ByVal ws As Worksheet, ByVal r0 As Long, _
                                ByVal copyLbl As String, ByVal h As Object, _
                                ByVal stavke As Variant, ByVal nStavke As Long) As Long
    Dim rr As Long: rr = r0
    Dim grayClr As Long: grayClr = RGB(90, 90, 90)
    Dim ruleClr As Long: ruleClr = RGB(110, 110, 110)
    Dim fillClr As Long: fillClr = RGB(217, 225, 242)

    ' --- zaglavlje firme (+ logo gore desno ako postoji) ---
    With ws.cells(rr, 1)
        .value = h("name")
        .Font.Bold = True
        .Font.Size = 12
    End With
    rr = rr + 1
    ws.cells(rr, 1).value = h("addr")
    rr = rr + 1
    With ws.cells(rr, 1)
        .value = "PIB: " & h("pib") & "    MB: " & h("mb") & "    Ziro: " & h("acct")
        .Font.Size = 9
        .Font.Color = grayClr
    End With
    DrawOtkupLogo ws, r0
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 6)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = ruleClr
    End With
    rr = rr + 1

    ' --- naslov (opis + veliki naslov, centrirano) ---
    ws.Range(ws.cells(rr, 1), ws.cells(rr, 6)).Merge
    With ws.cells(rr, 1)
        .value = "Otkup poljoprivrednih proizvoda"
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = grayClr
        .HorizontalAlignment = xlCenter
    End With
    rr = rr + 1
    ws.Range(ws.cells(rr, 1), ws.cells(rr, 6)).Merge
    With ws.cells(rr, 1)
        .value = "OTKUPNI LIST  br. " & h("brDok")
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 6)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = ruleClr
    End With
    rr = rr + 1

    ' --- datum / otkupno mesto / oznaka primerka ---
    WriteLabelVal ws, rr, 1, "Datum:", CStr(h("datum"))
    WriteLabelVal ws, rr, 4, "Otkupno mesto:", CStr(h("stanica"))
    rr = rr + 1
    With ws.cells(rr, 1)
        .value = copyLbl
        .Font.Italic = True
        .Font.Color = grayClr
    End With
    rr = rr + 1

    ' --- poljoprivrednik ---
    WriteLabelVal ws, rr, 1, "Poljoprivrednik:", CStr(h("koop"))
    rr = rr + 1
    WriteLabelVal ws, rr, 1, "BPG:", CStr(h("bpg"))
    WriteLabelVal ws, rr, 4, "Tekuci racun:", CStr(h("racun"))
    rr = rr + 1

    ' --- stavke ---
    Dim hdr As Long: hdr = rr
    ws.cells(rr, 1).value = "Rb"
    ws.cells(rr, 2).value = "Proizvod"
    ws.cells(rr, 3).value = "Klasa"
    ws.cells(rr, 4).value = "Kolicina kg"
    ws.cells(rr, 5).value = "Cena bez PDV"
    ws.cells(rr, 6).value = "Vrednost bez PDV"
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 6))
        .Font.Bold = True
        .Interior.Color = fillClr
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    rr = rr + 1
    Dim k As Long
    For k = 0 To nStavke - 1
        ws.cells(rr, 1).value = k + 1
        ws.cells(rr, 2).value = stavke(k, 0)
        ws.cells(rr, 3).value = stavke(k, 1)
        ws.cells(rr, 4).value = stavke(k, 2)
        ws.cells(rr, 5).value = stavke(k, 3)
        ws.cells(rr, 6).value = stavke(k, 4)
        ws.cells(rr, 1).HorizontalAlignment = xlCenter
        ws.cells(rr, 3).HorizontalAlignment = xlCenter
        rr = rr + 1
    Next k
    With ws.Range(ws.cells(hdr, 1), ws.cells(rr - 1, 6)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range(ws.cells(hdr + 1, 4), ws.cells(rr - 1, 6)).NumberFormat = "#,##0.00"

    ' --- ambalaza (levo) + obracun PDV nadoknade (desno, uokvireno) ---
    Dim ob As Long: ob = rr
    WriteLabelVal ws, ob, 1, "Ambalaza:", CStr(h("amb"))

    WriteLabelVal ws, ob, 4, "Osnovica (bez PDV):", ""
    ws.cells(ob, 6).value = h("osnovica")
    WriteLabelVal ws, ob + 1, 4, "PDV nadoknada (" & Format$(h("stopa"), "0.##") & "%):", ""
    ws.cells(ob + 1, 6).value = h("nadoknada")
    WriteLabelVal ws, ob + 2, 4, "UKUPNO ZA ISPLATU:", ""
    ws.cells(ob + 2, 6).value = h("ukupno")
    With ws.Range(ws.cells(ob + 2, 4), ws.cells(ob + 2, 6))
        .Font.Bold = True
        .Interior.Color = fillClr
    End With
    ws.Range(ws.cells(ob, 4), ws.cells(ob + 2, 6)).BorderAround Weight:=xlThin
    With ws.Range(ws.cells(ob + 2, 4), ws.cells(ob + 2, 6)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range(ws.cells(ob, 6), ws.cells(ob + 2, 6))
        .NumberFormat = "#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    rr = ob + 4

    ' --- napomena + potpisi ---
    With ws.cells(rr, 1)
        .value = "Poljoprivrednik svojim potpisom potvrdjuje prijem nadoknade."
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = grayClr
    End With
    rr = rr + 1
    ws.cells(rr, 1).value = "Potpis poljoprivrednika:  ________"
    ws.cells(rr, 1).Font.Color = grayClr
    ws.cells(rr, 4).value = "Potpis / pecat otkupljivaca:  ____________"
    ws.cells(rr, 4).Font.Color = grayClr
    rr = rr + 1

    WriteOtkupCopy = rr
End Function

' Upisuje "labela vrednost" u jednu celiju, sa podebljanim delom vrednosti.
Private Sub WriteLabelVal(ByVal ws As Worksheet, ByVal rowIx As Long, ByVal colIx As Long, _
                          ByVal lbl As String, ByVal val As String)
    Dim s As String
    If val = "" Then s = lbl Else s = lbl & " " & val
    With ws.cells(rowIx, colIx)
        .value = s
        .Font.Bold = False
        If Len(val) > 0 Then
            On Error Resume Next
            .Characters(Start:=Len(lbl) + 2, Length:=Len(val)).Font.Bold = True
            On Error GoTo 0
        End If
    End With
End Sub

' Ubacuje logo gore desno (preko zaglavlja) ako je dostupan. Tiho preskace ako ga nema.
Private Sub DrawOtkupLogo(ByVal ws As Worksheet, ByVal topRow As Long)
    On Error GoTo done
    Dim p As String: p = GetOtkupLogoPath()
    If p = "" Then Exit Sub

    Dim w As Double, hgt As Double
    w = 52: hgt = 40
    Dim rcell As Range: Set rcell = ws.cells(topRow, 6)
    Dim L As Double, T As Double
    L = rcell.Left + rcell.Width - w
    If L < ws.cells(topRow, 5).Left Then L = ws.cells(topRow, 5).Left
    T = rcell.Top

    ws.Shapes.AddPicture fileName:=p, LinkToFile:=msoFalse, _
                         SaveWithDocument:=msoTrue, _
                         Left:=L, Top:=T, Width:=w, Height:=hgt
done:
End Sub

' Putanja loga: config SELLER_LOGO_PATH, pa <workbook>\logo.png / logo.jpg. "" ako nema.
Private Function GetOtkupLogoPath() As String
    On Error Resume Next
    Dim p As String
    p = Trim$(CStr(GetConfigValue("SELLER_LOGO_PATH")))
    If p <> "" Then
        If Dir$(p) <> "" Then GetOtkupLogoPath = p: Exit Function
    End If
    Dim cand As String
    cand = ThisWorkbook.Path & "\logo.png"
    If Dir$(cand) <> "" Then GetOtkupLogoPath = cand: Exit Function
    cand = ThisWorkbook.Path & "\logo.jpg"
    If Dir$(cand) <> "" Then GetOtkupLogoPath = cand
End Function

Public Sub EnsureOtkupSablon()
    On Error GoTo EH
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("OtkupSablon")
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "OtkupSablon"
    ws.columns("A").ColumnWidth = 6
    ws.columns("B").ColumnWidth = 24
    ws.columns("C").ColumnWidth = 8
    ws.columns("D").ColumnWidth = 16
    ws.columns("E").ColumnWidth = 12
    ws.columns("F").ColumnWidth = 16
    Exit Sub
EH:
    LogErr "modPrint.EnsureOtkupSablon"
End Sub

Private Function PrNz(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then PrNz = CDbl(v)
End Function

