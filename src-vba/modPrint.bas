Attribute VB_Name = "modPrint"
Option Explicit

' --- Geometrija otkupnog lista (svaki primerak = tacno 1/3 A4) -------------
' Klijent koristi papir sa dve perforacije -> tri jednaka dela po 99mm
' (A4 visina 297mm / 3). Dva primerka idu u gornje dve trecine, donja trecina
' ostaje prazna. Stampa MORA biti 1:1 (Zoom=100, bez "Fit to page" / "Prilagodi").
Private Const OL_THIRD_PT As Double = 280.63      ' 99 mm u tackama (99/25.4*72)
Private Const OL_TOP_SPACER_PT As Double = 18#    ' ~6.3mm prazno iznad sadrzaja primerka
Private Const OL_MIN_FILLER_PT As Double = 4#     ' min. donji razmak do perforacije
' Kalibracija: ako stampac NE postuje TopMargin=0 nego sadrzaj pomeri nadole za
' T mm, prva perforacija nece pasti na granicu primeraka. Tada ovde unesi
' T u tackama (T_mm / 25.4 * 72) - prvi primerak se za toliko skrati i granica
' se vraca na 99mm. Podrazumevano 0 (stampac postuje 0 marginu / borderless).
Private Const OL_TOP_MARGIN_TRIM_PT As Double = 0#

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
    Dim iKolAmbIzd As Long, iVreme As Long
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
    iKolAmbIzd = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA)
    iVreme = GetColumnIndex(TBL_OTKUP, COL_OTK_VREME_UNOSA)

    Dim ids() As String: ids = Split(otkupIDs, " + ")
    Dim stavke() As Variant: ReDim stavke(0 To UBound(ids), 0 To 4)
    Dim cnt As Long: cnt = 0
    Dim osnovica As Double: osnovica = 0
    ' Cena uneta u formi je BRUTO (vec sadrzi PDV nadoknadu). U otkupnom listu
    ' prikazujemo NETO cenu/vrednost, a PDV nadoknadu dodajemo kao posebnu stavku.
    Dim stopa As Double: stopa = PrNz(GetConfigValue(CFG_PDV_NADOKNADA_STOPA))
    If stopa <= 0 Then stopa = PDV_NADOKNADA_DEFAULT
    Dim koopID As String, stID As String, brDok As String, datum As String
    Dim tipAmb As String, kolAmb As Double, kolAmbIzd As Double
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
                    stavke(cnt, 4) = cenBruto   ' Cena s PDV (bruto cena uneta u formi)
                    osnovica = osnovica + kol * cenNeto
                    If koopID = "" Then
                        koopID = CStr(d(r, iKoop)): stID = CStr(d(r, iSt))
                        brDok = CStr(d(r, iBr)): datum = Format$(d(r, iDat), "dd.mm.yyyy")
                        If iVreme > 0 Then If IsDate(d(r, iVreme)) Then datum = datum & "  (sn. " & Format$(d(r, iVreme), "hh:nn") & ")"
                        tipAmb = CStr(d(r, iTip)): kolAmb = PrNz(d(r, iKolAmb))
                        If iKolAmbIzd > 0 Then kolAmbIzd = PrNz(d(r, iKolAmbIzd))
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
    h("ambIzdata") = tipAmb & " x " & CStr(CLng(kolAmbIzd))
    h("stopa") = stopa
    h("osnovica") = osnovica
    h("nadoknada") = osnovica * stopa / 100
    h("ukupno") = osnovica + osnovica * stopa / 100
    h("rok") = DocConfigOr(CFG_OTKUP_ROK, OTKUP_ROK_DEFAULT)
    h("klauzula") = DocConfigOr(CFG_OTKUP_KLAUZULA, OtkupKlauzulaDefault())

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
    ' NAPOMENA: bez AutoFit - visine redova se postavljaju eksplicitno tako da
    ' svaki primerak zauzme tacno 1/3 A4 (99mm). AutoFit bi to ponistio.

    ' Dva primerka, svaki dopunjen filler-redom do tacno 1/3 A4 (99mm). Bez
    ' stampane linije za secenje - papir je vec perforiran; granice izmedju
    ' primeraka padaju tacno na perforacije (99mm i 198mm od vrha lista), a
    ' donja trecina (198-297mm) ostaje prazna.
    Dim r0 As Long, lastRow As Long
    r0 = WriteOtkupCopy(ws, 1, "Primerak za poljoprivrednika", h, stavke, cnt, _
                        OL_THIRD_PT - OL_TOP_MARGIN_TRIM_PT)
    lastRow = WriteOtkupCopy(ws, r0, "Primerak za otkupljivaca", h, stavke, cnt, _
                             OL_THIRD_PT) - 1

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = 100                       ' fiksna skala 1:1 -> visine redova u mm su tacne
        .LeftMargin = Application.InchesToPoints(0.31)
        .RightMargin = Application.InchesToPoints(0.31)
        .TopMargin = 0                    ' primerak 1 pocinje na vrhu lista (gornji red je prazan spacer)
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = False
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

' Ispisuje jedan kompaktan primerak od reda r0 i dopunjava ga filler-redom do
' tacno 1/3 A4 (targetPt tacaka). Vraca prvi slobodan red (pocetak sledeceg
' primerka). Sve visine redova su eksplicitne (bez AutoFit) da bi zbir bio tacan.
' Optimalno 1-2 stavke po primerku; vise stavki smanjuje donji razmak.
Private Function WriteOtkupCopy(ByVal ws As Worksheet, ByVal r0 As Long, _
                                ByVal copyLbl As String, ByVal h As Object, _
                                ByVal stavke As Variant, ByVal nStavke As Long, _
                                ByVal targetPt As Double) As Long
    Dim rr As Long
    Dim grayClr As Long: grayClr = DocColGray()
    Dim fillClr As Long: fillClr = DocColHeaderFill()
    Dim usedPt As Double: usedPt = 0

    ' --- gornji prazan razmak (apsorbuje nestampajucu ivicu stampaca) ---
    ws.rows(r0).RowHeight = OL_TOP_SPACER_PT
    usedPt = usedPt + OL_TOP_SPACER_PT
    rr = r0 + 1

    ' --- zaglavlje prodavca (3 reda) ---
    Dim shTop As Long: shTop = rr
    rr = DocSellerHeader(ws, rr, 6, 6)
    ws.cells(shTop, 1).Font.Size = 11
    ws.rows(shTop).RowHeight = 15#
    ws.rows(shTop + 1).RowHeight = 12#
    ws.rows(shTop + 2).RowHeight = 12#
    usedPt = usedPt + 39#

    ' --- naslov (2 reda) ---
    Dim tbTop As Long: tbTop = rr
    rr = DocTitleBlock(ws, rr, 6, "Otkup poljoprivrednih proizvoda", _
                       "OTKUPNI LIST  br. " & h("brDok"))
    ws.rows(tbTop).RowHeight = 12#
    ws.cells(tbTop + 1, 1).Font.Size = 14
    ws.rows(tbTop + 1).RowHeight = 18#
    usedPt = usedPt + 30#

    ' --- datum / otkupno mesto ---
    DocLabelVal ws, rr, 1, "Datum:", CStr(h("datum"))
    DocLabelVal ws, rr, 4, "Otkupno mesto:", CStr(h("stanica"))
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1

    ' --- oznaka primerka ---
    With ws.cells(rr, 1)
        .value = copyLbl
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = grayClr
    End With
    ws.rows(rr).RowHeight = 11#
    usedPt = usedPt + 11#
    rr = rr + 1

    ' --- poljoprivrednik + BPG + tekuci racun ---
    DocLabelVal ws, rr, 1, "Poljoprivrednik:", CStr(h("koop"))
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1
    DocLabelVal ws, rr, 1, "BPG:", CStr(h("bpg"))
    DocLabelVal ws, rr, 4, "Tekuci racun:", CStr(h("racun"))
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1

    ' --- stavke (kratke oznake kolona, jedan red) ---
    Dim hdr As Long: hdr = rr
    ws.cells(rr, 1).value = "Rb"
    ws.cells(rr, 2).value = "Proizvod"
    ws.cells(rr, 3).value = "Klasa"
    ws.cells(rr, 4).value = "Kol. (kg)"
    ws.cells(rr, 5).value = "Cena neto"
    ws.cells(rr, 6).value = "Cena s PDV"
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 6))
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = fillClr
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.rows(rr).RowHeight = 14#
    usedPt = usedPt + 14#
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
        ws.rows(rr).RowHeight = 13#
        usedPt = usedPt + 13#
        rr = rr + 1
    Next k
    With ws.Range(ws.cells(hdr, 1), ws.cells(rr - 1, 6)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range(ws.cells(hdr + 1, 4), ws.cells(rr - 1, 6)).NumberFormat = "#,##0.00"

    ' --- ambalaza + rok isplate (levo) + obracun PDV nadoknade (desno, uokvireno) ---
    Dim ob As Long: ob = rr
    DocLabelVal ws, ob, 1, "Primljena ambalaza:", CStr(h("amb"))
    DocLabelVal ws, ob + 1, 1, "Izdata ambalaza:", CStr(h("ambIzdata"))
    DocLabelVal ws, ob + 2, 1, "Rok isplate:", CStr(h("rok"))

    DocLabelVal ws, ob, 4, "Osnovica (bez PDV):", ""
    ws.cells(ob, 6).value = h("osnovica")
    DocLabelVal ws, ob + 1, 4, "PDV nadoknada (" & Format$(h("stopa"), "0.##") & "%):", ""
    ws.cells(ob + 1, 6).value = h("nadoknada")
    DocLabelVal ws, ob + 2, 4, "UKUPNO ZA ISPLATU:", ""
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
    ws.rows(ob).RowHeight = 13#
    ws.rows(ob + 1).RowHeight = 13#
    ws.rows(ob + 2).RowHeight = 14#
    usedPt = usedPt + 40#
    rr = ob + 3

    ' --- klauzula (obavezni element otkupnog bloka), sitan font ---
    ws.Range(ws.cells(rr, 1), ws.cells(rr, 6)).Merge
    With ws.cells(rr, 1)
        .value = CStr(h("klauzula"))
        .Font.Size = 7
        .Font.Color = RGB(60, 60, 60)
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
    ws.rows(rr).RowHeight = 30#
    usedPt = usedPt + 30#
    rr = rr + 1

    ' --- napomena + potpisi ---
    With ws.cells(rr, 1)
        .value = "Poljoprivrednik svojim potpisom potvrdjuje prijem nadoknade."
        .Font.Size = 7
        .Font.Italic = True
        .Font.Color = grayClr
    End With
    ws.rows(rr).RowHeight = 10#
    usedPt = usedPt + 10#
    rr = rr + 1
    ws.cells(rr, 1).value = "Potpis poljoprivrednika:  ____________"
    ws.cells(rr, 1).Font.Size = 9
    ws.cells(rr, 1).Font.Color = grayClr
    ws.cells(rr, 4).value = "Potpis / pecat otkupljivaca:  ____________"
    ws.cells(rr, 4).Font.Size = 9
    ws.cells(rr, 4).Font.Color = grayClr
    ws.rows(rr).RowHeight = 16#
    usedPt = usedPt + 16#
    rr = rr + 1

    ' --- filler red: dopuni primerak do tacno 1/3 A4 (targetPt) ---
    Dim fillPt As Double: fillPt = targetPt - usedPt
    If fillPt < OL_MIN_FILLER_PT Then fillPt = OL_MIN_FILLER_PT
    ws.rows(rr).RowHeight = fillPt
    rr = rr + 1

    WriteOtkupCopy = rr
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

