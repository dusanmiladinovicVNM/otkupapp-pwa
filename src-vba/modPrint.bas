Attribute VB_Name = "modPrint"
Option Explicit

' --- Geometrija otkupnog lista (svaki primerak = tacno 1/3 A4) -------------
' Klijent koristi papir sa dve perforacije -> tri jednaka dela po 99mm
' (A4 visina 297mm / 3). Dva primerka idu u gornje dve trecine, donja trecina
' ostaje prazna. Stampa MORA biti 1:1 (Zoom=100, bez "Fit to page" / "Prilagodi").
Private Const OL_THIRD_PT As Double = 280.63      ' 99 mm u tackama (99/25.4*72)
Private Const OL_TOP_SPACER_PT As Double = 9#     ' ~3.2mm prazno iznad sadrzaja primerka (smanjena gornja margina)
Private Const OL_MIN_FILLER_PT As Double = 17#    ' donji razmak do perforacije (vise = potpisi vise iznad reza)
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
    Dim iKolAmbIzd As Long, iVreme As Long, iBruto As Long
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
    iBruto = GetColumnIndex(TBL_OTKUP, COL_OTK_BRUTO)

    Dim ids() As String: ids = Split(otkupIDs, " + ")
    Dim stavke() As Variant: ReDim stavke(0 To UBound(ids), 0 To 6)
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
                    ' Bruto: zamrznut iz unosa (BrutoKg) ako postoji, inace izvedeno iz
                    ' trenutne tare gajbice (fallback za stare/neto redove). Zamrznut bruto
                    ' ostaje tacan i ako se tezina gajbice kasnije promeni u sifarniku.
                    Dim storedBruto As Double: If iBruto > 0 Then storedBruto = PrNz(d(r, iBruto))
                    Dim kolBruto As Double
                    If storedBruto > 0 Then
                        kolBruto = storedBruto
                    Else
                        Dim crateW As Double: crateW = PrNz(LookupValue(TBL_TIP_AMBALAZE, COL_TAMB_TIP, CStr(d(r, iTip)), COL_TAMB_TEZINA))
                        kolBruto = kol + PrNz(d(r, iKolAmb)) * crateW
                    End If
                    stavke(cnt, 0) = Trim$(CStr(d(r, iVr)) & " " & CStr(d(r, iSo)))
                    stavke(cnt, 1) = CStr(d(r, iKl))
                    stavke(cnt, 2) = cenNeto        ' Cena bez PDV
                    stavke(cnt, 3) = cenBruto       ' Cena s PDV
                    stavke(cnt, 4) = kol            ' Kolicina neto
                    stavke(cnt, 5) = kolBruto       ' Kolicina bruto (neto + gajbice * tara)
                    stavke(cnt, 6) = kol * cenNeto  ' Vrednost neto
                    osnovica = osnovica + kol * cenNeto
                    ' Primljena ambalaza = zbir gajbi po SVIM stavkama (Klasa I + II).
                    kolAmb = kolAmb + PrNz(d(r, iKolAmb))
                    If koopID = "" Then
                        koopID = CStr(d(r, iKoop)): stID = CStr(d(r, iSt))
                        brDok = CStr(d(r, iBr)): datum = Format$(d(r, iDat), "dd.mm.yyyy")
                        If iVreme > 0 Then If IsDate(d(r, iVreme)) Then datum = datum & "  Vreme: " & Format$(d(r, iVreme), "hh:nn")
                        tipAmb = CStr(d(r, iTip))
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
    Dim objMesto As String: objMesto = Trim$(CStr(GetConfigValue("SELLER_OBJEKAT_MESTO")))
    Dim objReg As String: objReg = Trim$(CStr(GetConfigValue("SELLER_OBJEKAT_BR_REGISTRA")))
    Dim objLine As String: objLine = ""
    If Len(objMesto) > 0 Then objLine = "Objekat: " & objMesto
    If Len(objReg) > 0 Then
        If Len(objLine) > 0 Then
            objLine = objLine & "    Reg. br: " & objReg
        Else
            objLine = "Objekat reg. br: " & objReg
        End If
    End If
    h("objekat") = objLine
    h("brDok") = brDok
    h("datum") = datum
    h("stanica") = CStr(LookupValue(TBL_STANICE, "StanicaID", stID, "Naziv"))
    h("koop") = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, "Ime")) & " " & _
                CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, "Prezime")))
    h("bpg") = CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, COL_KOOP_BPG))
    h("racun") = CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, COL_KOOP_TEKUCI_RACUN))
    ' Saldo ambalaze = entitetski saldo kooperanta (koliko gajbica drzi/duguje):
    ' pocetno stanje pre bloka (Ulaz +, Izlaz -) + izdato (Kooperant Ulaz)
    ' - primljeno (Kooperant Izlaz). Pocetno se cita iz ledgera po redosledu
    ' upisa, pa je ispravno i kod ponovne stampe starijeg bloka.
    Dim ambPoc As Long: ambPoc = GetKooperantAmbOpening(koopID, tipAmb, ids)
    h("ambPocetno") = tipAmb & " x " & CStr(ambPoc)
    h("ambPrijem") = CStr(CLng(kolAmb))         ' primljeno (broj gajbi, tekuci blok)
    h("ambIzdavanje") = CStr(CLng(kolAmbIzd))   ' izdato (broj gajbi, tekuci blok)
    h("ambSaldo") = tipAmb & " x " & CStr(CLng(ambPoc + kolAmbIzd - kolAmb))
    h("stopa") = stopa
    h("osnovica") = osnovica
    h("nadoknada") = osnovica * stopa / 100
    h("ukupno") = osnovica + osnovica * stopa / 100
    h("rok") = DocConfigOr(CFG_OTKUP_ROK, OTKUP_ROK_DEFAULT)
    ' Klauzula iz podesavanja moze da sadrzi tokene koji se pri stampi zamenjuju
    ' podacima OVOG otkupnog lista (broj gazdinstva se razlikuje po kooperantu):
    '   {BPG} {POLJOPRIVREDNIK} {RACUN} {DATUM} {BROJ}
    ' Zamena je case-insensitive (radi i {bpg}).
    Dim kl As String: kl = DocConfigOr(CFG_OTKUP_KLAUZULA, OtkupKlauzulaDefault())
    kl = Replace(kl, "{BPG}", CStr(h("bpg")), , , vbTextCompare)
    kl = Replace(kl, "{POLJOPRIVREDNIK}", CStr(h("koop")), , , vbTextCompare)
    kl = Replace(kl, "{RACUN}", CStr(h("racun")), , , vbTextCompare)
    kl = Replace(kl, "{DATUM}", CStr(h("datum")), , , vbTextCompare)
    kl = Replace(kl, "{BROJ}", CStr(h("brDok")), , , vbTextCompare)
    h("klauzula") = kl

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
    ws.columns("A").ColumnWidth = 4
    ws.columns("B").ColumnWidth = 17
    ws.columns("C").ColumnWidth = 6
    ws.columns("D").ColumnWidth = 10
    ws.columns("E").ColumnWidth = 10
    ws.columns("F").ColumnWidth = 9
    ws.columns("G").ColumnWidth = 9
    ws.columns("H").ColumnWidth = 12
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
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(lastRow, 8)).Address
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

    ' --- zaglavlje prodavca (2 reda): ime; adresa levo + PIB/MB/Ziro desno ---
    With ws.cells(rr, 1)
        .value = CStr(h("name"))
        .Font.Bold = True
        .Font.Size = 11
    End With
    ws.rows(rr).RowHeight = 15#
    usedPt = usedPt + 15#
    rr = rr + 1
    With ws.cells(rr, 1)
        .value = CStr(h("addr"))
        .Font.Size = 9
        .Font.Color = grayClr
    End With
    ws.Range(ws.cells(rr, 5), ws.cells(rr, 8)).Merge
    With ws.cells(rr, 5)
        .value = "PIB: " & CStr(h("pib")) & "   MB: " & CStr(h("mb")) & "   Ziro: " & CStr(h("acct"))
        .Font.Size = 9
        .Font.Color = grayClr
        .HorizontalAlignment = xlRight
    End With
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 8)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = DocColRule()
    End With
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1

    ' --- objekat (lokacija + br registra), ispod zaglavlja (samo ako je u configu) ---
    If Len(CStr(h("objekat"))) > 0 Then
        With ws.cells(rr, 1)
            .value = CStr(h("objekat"))
            .Font.Size = 9
        End With
        ws.rows(rr).RowHeight = 12#
        usedPt = usedPt + 12#
        rr = rr + 1
    End If

    ' --- naslov (centriran preko cele sirine): mali descriptor + veliki OTKUPNI LIST ---
    ws.Range(ws.cells(rr, 1), ws.cells(rr, 8)).Merge
    With ws.cells(rr, 1)
        .value = "Otkup poljoprivrednih proizvoda"
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = grayClr
        .HorizontalAlignment = xlCenter
    End With
    ws.rows(rr).RowHeight = 12#
    usedPt = usedPt + 12#
    rr = rr + 1
    ws.Range(ws.cells(rr, 1), ws.cells(rr, 8)).Merge
    With ws.cells(rr, 1)
        .value = "OTKUPNI LIST  br. " & h("brDok")
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 8)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = DocColRule()
    End With
    ws.rows(rr).RowHeight = 20#
    usedPt = usedPt + 20#
    rr = rr + 1

    ' --- datum / otkupno mesto / rok isplate (skroz desno) ---
    DocLabelVal ws, rr, 1, "Datum:", CStr(h("datum"))
    DocLabelVal ws, rr, 4, "Otkupno mesto:", CStr(h("stanica"))
    ws.Range(ws.cells(rr, 6), ws.cells(rr, 8)).Merge
    With ws.cells(rr, 6)
        .value = "Rok isplate: " & CStr(h("rok"))
        .Font.Size = 9
        .HorizontalAlignment = xlRight
    End With
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1

    ' --- poljoprivrednik: ime skroz levo, pa BPG, pa tekuci racun (1 red) ---
    With ws.cells(rr, 1)
        .value = CStr(h("koop"))
        .Font.Bold = True
    End With
    DocLabelVal ws, rr, 4, "BPG:", CStr(h("bpg"))
    DocLabelVal ws, rr, 6, "TR:", CStr(h("racun"))
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1

    ' --- stavke (kratke oznake kolona, jedan red) ---
    Dim hdr As Long: hdr = rr
    ws.cells(rr, 1).value = "Rb"
    ws.cells(rr, 2).value = "Proizvod"
    ws.cells(rr, 3).value = "Klasa"
    ws.cells(rr, 4).value = "Cena bez PDV"
    ws.cells(rr, 5).value = "Cena s PDV"
    ws.cells(rr, 6).value = "Kol. neto"
    ws.cells(rr, 7).value = "Kol. bruto"
    ws.cells(rr, 8).value = "Vrednost"
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 8))
        .Font.Bold = True
        .Font.Size = 8
        .Interior.Color = fillClr
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.rows(rr).RowHeight = 20#
    usedPt = usedPt + 20#
    rr = rr + 1
    Dim k As Long
    For k = 0 To nStavke - 1
        ws.cells(rr, 1).value = k + 1
        ws.cells(rr, 2).value = stavke(k, 0)
        ws.cells(rr, 3).value = stavke(k, 1)
        ws.cells(rr, 4).value = stavke(k, 2)
        ws.cells(rr, 5).value = stavke(k, 3)
        ws.cells(rr, 6).value = stavke(k, 4)
        ws.cells(rr, 7).value = stavke(k, 5)
        ws.cells(rr, 8).value = stavke(k, 6)
        ws.cells(rr, 1).HorizontalAlignment = xlCenter
        ws.cells(rr, 3).HorizontalAlignment = xlCenter
        ws.rows(rr).RowHeight = 13#
        usedPt = usedPt + 13#
        rr = rr + 1
    Next k
    With ws.Range(ws.cells(hdr, 1), ws.cells(rr - 1, 8)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range(ws.cells(hdr + 1, 4), ws.cells(rr - 1, 8)).NumberFormat = "#,##0.00"

    ' --- ambalaza tabelica (3 reda, da primerak ostane tacno 1/3 A4): pocetno
    '     stanje | primljeno + izdato (jedan red, dva inline para u kol. 1 i 3) |
    '     saldo. Levo uokvireno; desno obracun PDV nadoknade (isto 3 reda). ---
    Dim ob As Long: ob = rr
    DocLabelVal ws, ob, 1, "Pocetno stanje:", ""
    ws.cells(ob, 4).value = CStr(h("ambPocetno"))
    DocLabelVal ws, ob + 1, 1, "Primljeno:", CStr(h("ambPrijem"))
    DocLabelVal ws, ob + 1, 3, "Izdato:", CStr(h("ambIzdavanje"))
    DocLabelVal ws, ob + 2, 1, "Saldo ambalaze:", ""
    ws.cells(ob + 2, 4).value = CStr(h("ambSaldo"))
    ws.cells(ob + 2, 1).Font.Bold = True
    ws.cells(ob + 2, 4).Font.Bold = True
    With ws.Range(ws.cells(ob, 4), ws.cells(ob + 2, 4))   ' pocetno/saldo desno poravnati
        .HorizontalAlignment = xlRight
    End With
    ws.Range(ws.cells(ob, 1), ws.cells(ob + 2, 4)).BorderAround Weight:=xlThin
    With ws.Range(ws.cells(ob + 2, 1), ws.cells(ob + 2, 4)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    DocLabelVal ws, ob, 5, "Osnovica (bez PDV):", ""
    ws.cells(ob, 8).value = h("osnovica")
    DocLabelVal ws, ob + 1, 5, "PDV nadoknada (" & Format$(h("stopa"), "0.##") & "%):", ""
    ws.cells(ob + 1, 8).value = h("nadoknada")
    DocLabelVal ws, ob + 2, 5, "UKUPNO ZA ISPLATU:", ""
    ws.cells(ob + 2, 8).value = h("ukupno")
    With ws.Range(ws.cells(ob + 2, 5), ws.cells(ob + 2, 8))
        .Font.Bold = True
        .Interior.Color = fillClr
    End With
    ws.Range(ws.cells(ob, 5), ws.cells(ob + 2, 8)).BorderAround Weight:=xlThin
    With ws.Range(ws.cells(ob + 2, 5), ws.cells(ob + 2, 8)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Range(ws.cells(ob, 8), ws.cells(ob + 2, 8))
        .NumberFormat = "#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    ws.rows(ob).RowHeight = 13#
    ws.rows(ob + 1).RowHeight = 13#
    ws.rows(ob + 2).RowHeight = 14#
    usedPt = usedPt + 40#
    rr = ob + 3

    ' --- klauzula (obavezni element): upija sav preostali prostor (sav "dobitak"
    '     od obrisanih redova) tako da potpisi dodju tacno iznad perforacije, a
    '     duga klauzula iz podesavanja ima mesta da se prikaze ---
    Dim reservePt As Double: reservePt = 16# + OL_MIN_FILLER_PT        ' potpisi + donji razmak
    Dim klauzPt As Double: klauzPt = targetPt - usedPt - reservePt
    If klauzPt < 24# Then klauzPt = 24#                                ' minimalna visina klauzule
    ws.Range(ws.cells(rr, 1), ws.cells(rr, 8)).Merge
    With ws.cells(rr, 1)
        .value = CStr(h("klauzula"))
        .Font.Size = 7
        .Font.Color = RGB(60, 60, 60)
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlJustify
    End With
    ws.rows(rr).RowHeight = klauzPt
    usedPt = usedPt + klauzPt
    rr = rr + 1

    ' --- potpisi (napomena uklonjena - klauzula vec sadrzi saglasnost potpisom) ---
    ws.cells(rr, 1).value = "Potpis poljoprivrednika:  ____________"
    ws.cells(rr, 1).Font.Size = 9
    ws.cells(rr, 1).Font.Color = grayClr
    ws.cells(rr, 5).value = "Potpis / pecat otkupljivaca:  ____________"
    ws.cells(rr, 5).Font.Size = 9
    ws.cells(rr, 5).Font.Color = grayClr
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

' ============================================================
' PRIJEMNICA — PrijemnicaSablon, jedan A4 portrait dokument (prijem robe na
' hladnjacu). Izlaz po CFG_PRIJEMNICA_PRINT_MODE: OFF/prazno (DEFAULT, bez
' izlaza — kao do sada) | PDF | PRINT | PREVIEW. Auto-izlaz okida
' frmDokumenta.btnUnosPrij posle snimanja. prijemnicaIDs = rezultat
' SavePrijemnicaMulti_TX ("PRJ-1" ili "PRJ-1 + PRJ-2" za dve klase).
' ============================================================

' Glavni ulaz (zove se posle SavePrijemnicaMulti_TX). Best-effort: greska se loguje.
Public Sub OutputPrijemnica(ByVal prijemnicaIDs As String)
    On Error GoTo EH
    Dim mode As String
    mode = UCase$(Trim$(GetConfigValue(CFG_PRIJEMNICA_PRINT_MODE)))

    Select Case mode
        Case "PRINT"
            Dim ws As Worksheet: Set ws = FillPrijemnicaSablon(prijemnicaIDs)
            If Not ws Is Nothing Then ws.PrintOut Copies:=1
        Case "PREVIEW"
            Dim wp As Worksheet: Set wp = FillPrijemnicaSablon(prijemnicaIDs)
            If Not wp Is Nothing Then wp.PrintPreview
        Case "PDF"
            ExportPrijemnicaPDF prijemnicaIDs, True
        Case Else
            ' OFF ili prazno (DEFAULT) -> bez izlaza; ponasanje kao do sada.
    End Select
    Exit Sub
EH:
    LogErr "modPrint.OutputPrijemnica"
End Sub

' Pojedinacni prijemnica -> fizicka stampa (simetricno PrintOtkupniList/PrintPaletniList).
Public Sub PrintPrijemnica(ByVal prijemnicaID As String)
    Dim ws As Worksheet: Set ws = FillPrijemnicaSablon(prijemnicaID)
    If Not ws Is Nothing Then ws.PrintOut Copies:=1
End Sub

' PDF prijemnice -> <workbook>\Prijemnica_<brPrij>.pdf. Vraca putanju.
Public Function ExportPrijemnicaPDF(ByVal prijemnicaIDs As String, _
                                    Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = FillPrijemnicaSablon(prijemnicaIDs)
    If ws Is Nothing Then
        MsgBox "PDF prijemnice nije napravljen: priprema lista (PrijemnicaSablon) " & _
               "nije uspela. Proveri da li su podaci prijemnice kompletni.", _
               vbExclamation, APP_NAME
        Exit Function
    End If

    Dim folder As String: folder = ThisWorkbook.Path
    If Len(folder) = 0 Then folder = Environ$("TEMP")
    Dim suff As String: suff = Replace(Replace(prijemnicaIDs, " + ", "_"), "/", "-")
    ' Vremenski pecat u imenu -> nema "file in use" (1004) ako je prethodni PDF otvoren.
    Dim pdfPath As String
    pdfPath = folder & "\Prijemnica_" & suff & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, OpenAfterPublish:=openAfter
    ExportPrijemnicaPDF = pdfPath
    Exit Function
EH:
    MsgBox "Greska pri izradi PDF prijemnice:" & vbCrLf & _
           "  [" & Err.Number & "] " & Err.Description, vbExclamation, APP_NAME
    LogErr "modPrint.ExportPrijemnicaPDF"
End Function

' Popuni PrijemnicaSablon (zaglavlje firme + podaci + stavke + potpisi). Vraca sheet
' (ili Nothing). Klasa I [+ Klasa II] = po jedan red u tabeli stavki; zaglavlje
' (datum/kupac/vozac/brojevi) se cita iz prvog reda.
Private Function FillPrijemnicaSablon(ByVal prijemnicaIDs As String) As Worksheet
    On Error GoTo EH
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsurePrijemnicaSablon
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PrijemnicaSablon")
    On Error GoTo EH
    If ws Is Nothing Then Exit Function

    Dim d As Variant: d = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long, iDat As Long, iKup As Long, iVoz As Long, iBr As Long, iBrZbr As Long
    Dim iVr As Long, iSo As Long, iKl As Long, iKol As Long, iCe As Long, iTip As Long
    Dim iKolAmb As Long, iKolAmbV As Long
    iID = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    iDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    iKup = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC)
    iVoz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VOZAC)
    iBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    iBrZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    iVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    iSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    iKl = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA)
    iKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    iCe = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    iTip = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_TIP_AMB)
    iKolAmb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    iKolAmbV = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA)

    Dim ids() As String: ids = Split(prijemnicaIDs, " + ")
    Dim stavke() As Variant: ReDim stavke(0 To UBound(ids), 0 To 5)
    Dim cnt As Long: cnt = 0
    Dim kupID As String, vozID As String, brPrij As String, brZbr As String
    Dim datum As String
    Dim ukKg As Double, ukVred As Double, ukAmb As Double, ambV As Double
    Dim j As Long, r As Long
    For j = 0 To UBound(ids)
        Dim wantID As String: wantID = Trim$(ids(j))
        If wantID <> "" Then
            For r = 1 To UBound(d, 1)
                If CStr(d(r, iID)) = wantID Then
                    Dim pkol As Double: pkol = PrNz(d(r, iKol))
                    Dim pcen As Double: pcen = PrNz(d(r, iCe))
                    stavke(cnt, 0) = Trim$(CStr(d(r, iVr)) & " " & CStr(d(r, iSo)))
                    stavke(cnt, 1) = CStr(d(r, iKl))
                    stavke(cnt, 2) = pkol
                    stavke(cnt, 3) = pcen
                    stavke(cnt, 4) = CStr(d(r, iTip))
                    stavke(cnt, 5) = PrNz(d(r, iKolAmb))
                    ukKg = ukKg + pkol
                    ukVred = ukVred + pkol * pcen
                    ukAmb = ukAmb + PrNz(d(r, iKolAmb))
                    If cnt = 0 Then
                        kupID = CStr(d(r, iKup)): vozID = CStr(d(r, iVoz))
                        brPrij = CStr(d(r, iBr)): brZbr = CStr(d(r, iBrZbr))
                        datum = Format$(d(r, iDat), "dd.mm.yyyy")
                        ambV = PrNz(d(r, iKolAmbV))
                    End If
                    cnt = cnt + 1
                    Exit For
                End If
            Next r
        End If
    Next j
    If cnt = 0 Then Exit Function

    Dim kupNaziv As String, vozNaziv As String
    kupNaziv = CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, kupID, COL_KUP_NAZIV))
    vozNaziv = Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", vozID, "Ime")) & " " & _
                     CStr(LookupValue(TBL_VOZACI, "VozacID", vozID, "Prezime")))

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

    ' --- zaglavlje firme + naslov ---
    Dim rr As Long: rr = 1
    rr = DocSellerHeader(ws, rr, 8, 8)
    rr = rr + 1
    rr = DocTitleBlock(ws, rr, 8, "Prijem robe na hladnjacu", "PRIJEMNICA  br. " & brPrij)
    rr = rr + 1

    ' --- podaci o prijemu ---
    DocLabelVal ws, rr, 1, "Datum:", datum
    DocLabelVal ws, rr, 5, "Broj zbirne:", brZbr
    rr = rr + 1
    DocLabelVal ws, rr, 1, "Kupac / hladnjaca:", kupNaziv
    rr = rr + 1
    DocLabelVal ws, rr, 1, "Vozac:", vozNaziv
    rr = rr + 2

    ' --- stavke (Klasa I [+ Klasa II]) ---
    Dim hdr As Long: hdr = rr
    ws.cells(rr, 1).value = "Rb"
    ws.cells(rr, 2).value = "Proizvod"
    ws.cells(rr, 3).value = "Klasa"
    ws.cells(rr, 4).value = "Kol. (kg)"
    ws.cells(rr, 5).value = "Cena"
    ws.cells(rr, 6).value = "Vrednost"
    ws.cells(rr, 7).value = "Tip amb."
    ws.cells(rr, 8).value = "Br. gajbica"
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 8))
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = DocColHeaderFill()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    rr = rr + 1
    Dim k As Long
    For k = 0 To cnt - 1
        ws.cells(rr, 1).value = k + 1
        ws.cells(rr, 2).value = stavke(k, 0)
        ws.cells(rr, 3).value = stavke(k, 1)
        ws.cells(rr, 4).value = stavke(k, 2)
        ws.cells(rr, 5).value = stavke(k, 3)
        ws.cells(rr, 6).value = CDbl(stavke(k, 2)) * CDbl(stavke(k, 3))
        ws.cells(rr, 7).value = stavke(k, 4)
        ws.cells(rr, 8).value = stavke(k, 5)
        ws.cells(rr, 1).HorizontalAlignment = xlCenter
        ws.cells(rr, 3).HorizontalAlignment = xlCenter
        rr = rr + 1
    Next k
    ' --- ukupno red ---
    ws.cells(rr, 2).value = "UKUPNO"
    ws.cells(rr, 4).value = ukKg
    ws.cells(rr, 6).value = ukVred
    ws.cells(rr, 8).value = ukAmb
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 8))
        .Font.Bold = True
        .Interior.Color = DocColHeaderFill()
    End With
    With ws.Range(ws.cells(hdr, 1), ws.cells(rr, 8)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.Range(ws.cells(hdr + 1, 4), ws.cells(rr, 6)).NumberFormat = "#,##0.00"
    rr = rr + 2

    ' --- vracena ambalaza (ako je uneta) ---
    If ambV > 0 Then
        DocLabelVal ws, rr, 1, "Vracena ambalaza (kom):", CStr(CLng(ambV))
        rr = rr + 2
    End If

    ' --- footer: datum stampe + potpisi ---
    ws.cells(rr, 1).value = "Datum stampe: " & Format$(Date, "dd.mm.yyyy")
    ws.cells(rr, 1).Font.Color = DocColGray()
    rr = rr + 2
    ws.cells(rr, 1).value = "Robu predao (vozac): ____________________"
    ws.cells(rr, 1).Font.Color = DocColGray()
    ws.cells(rr, 5).value = "Robu primio: ____________________"
    ws.cells(rr, 5).Font.Color = DocColGray()
    Dim lastRow As Long: lastRow = rr

    ' --- A4 portrait, sve kolone na jednu stranu po sirini ---
    ' PageSetup property-ji (PaperSize/Orientation/FitToPages) traze drajver
    ' stampaca; na racunaru bez stampaca bacaju 1004. Stitimo ih (On Error Resume
    ' Next) da PDF izadje i bez stampaca. NE koristimo Application.PrintCommunication
    ' jer na takvim masinama zna da blokira (isti hardening kao PrintSpecifikacija).
    On Error Resume Next
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(lastRow, 8)).Address
    End With
    On Error GoTo 0

    Application.ScreenUpdating = oldScreen
    Set FillPrijemnicaSablon = ws
    Exit Function
EH:
    Application.ScreenUpdating = oldScreen
    LogErr "modPrint.FillPrijemnicaSablon"
End Function

Public Sub EnsurePrijemnicaSablon()
    On Error GoTo EH
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PrijemnicaSablon")
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "PrijemnicaSablon"
    ws.columns("A").ColumnWidth = 5
    ws.columns("B").ColumnWidth = 20
    ws.columns("C").ColumnWidth = 7
    ws.columns("D").ColumnWidth = 11
    ws.columns("E").ColumnWidth = 11
    ws.columns("F").ColumnWidth = 13
    ws.columns("G").ColumnWidth = 12
    ws.columns("H").ColumnWidth = 11
    Exit Sub
EH:
    LogErr "modPrint.EnsurePrijemnicaSablon"
End Sub

Private Function PrNz(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then PrNz = CDbl(v)
End Function

