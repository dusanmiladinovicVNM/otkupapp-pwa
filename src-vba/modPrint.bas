Attribute VB_Name = "modPrint"
'Attribute VB_Name = "modPrint"
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

' Smer trenutnog OM<->koop reversa (postavlja OutputIzdavanjeAmbalaze): False =
' izdavanje (OM->koop), True = prijem/povrat (koop->OM). Cita ga Export (ime PDF)
' i WriteIzdavanjeCopy (preko h("prijem")) za naslov/saldo/potpise.
Private m_izdAmbPrijem As Boolean

' ============================================================
' modPrint - Druckausgabe (ersetzt direkte PrintOut-Aufrufe)
' ============================================================

Public Sub PrintIzvestaj(ByVal data As Variant, ByVal reportTitle As String, _
                         ByVal headers As Variant)
    ' Generischer Report-Druck
    ' Schreibt in ein temporaeres Print-Sheet und druckt
    
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
    
    ' Izlaz: PDF koji se otvori (pouzdanije od direktne stampe; bez zavisnosti
    ' od podrazumevanog stampaca). Sheet mora biti vidljiv za ExportAsFixedFormat.
    On Error Resume Next
    wsPrint.Visible = xlSheetVisible
    Dim pdfPath As String
    pdfPath = EnsureDocFolder(PDF_DIR_IZVESTAJI) & "\Izve" & ChrW(353) & "taj_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    wsPrint.UsedRange.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                                          Quality:=xlQualityStandard, OpenAfterPublish:=True
    On Error GoTo 0
    
    ' Aufraeumen
    wsPrint.Visible = xlSheetVeryHidden
End Sub

' ============================================================
' OTPREMNICA (PDF) - isti vizuelni stil kao otkupni / grupni otkupni list
' (WriteOtkupCopy, dva primerka 1/3 A4), podaci iz reda tblOtpremnice.
' ============================================================
Public Sub OutputOtpremnicaPDF(ByVal otpID As String)
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = FillOtpremnicaSablon(otpID)
    If ws Is Nothing Then
        MsgBox "Otpremnica nije pronadjena ili se ne mo" & ChrW(382) & "e pripremiti (" & otpID & ").", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim broj As String: broj = CStr(LookupValue(TBL_OTPREMNICA, COL_OTP_ID, otpID, COL_OTP_BROJ))
    Dim suff As String: suff = Replace(Replace(broj, "/", "-"), "\\", "-")
    If Len(Trim$(suff)) = 0 Then suff = otpID
    Dim pdfPath As String: pdfPath = EnsureDocFolder(PDF_DIR_OTPREMNICE) & "\Otpremnica_" & suff & ".pdf"
    DocExportPdf ws, pdfPath, True
    Exit Sub
EH:
    LogErr "modPrint.OutputOtpremnicaPDF"
    MsgBox "Gre" & ChrW(353) & "ka pri stampi otpremnice: " & Err.description, vbCritical, APP_NAME
End Sub

' Popuni OtpremnicaSablon iz reda tblOtpremnice (isti obrazac kao grupni otkupni
' list: WriteOtkupCopy, dva primerka). Vraca sheet (ili Nothing).
Private Function FillOtpremnicaSablon(ByVal otpID As String) As Worksheet
    On Error GoTo EH
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsureOtpremnicaSablon
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_OTPREMNICA_SABLON)
    On Error GoTo EH
    If ws Is Nothing Then Exit Function

    Dim d As Variant: d = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(d) Then Exit Function
    Dim cId As Long: cId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    If cId = 0 Then Exit Function
    Dim fr As Long, r As Long: fr = 0
    For r = 1 To UBound(d, 1)
        If CStr(d(r, cId)) = otpID Then fr = r: Exit For
    Next r
    If fr = 0 Then Exit Function

    Dim stopa As Double: stopa = PrNz(GetConfigValue(CFG_PDV_NADOKNADA_STOPA))
    If stopa <= 0 Then stopa = PDV_NADOKNADA_DEFAULT

    Dim kol As Double: kol = OtpN(d, fr, COL_OTP_KOLICINA)
    Dim cenBruto As Double: cenBruto = OtpN(d, fr, COL_OTP_CENA)
    Dim cenNeto As Double: cenNeto = cenBruto / (1 + stopa / 100)
    Dim storedBruto As Double: storedBruto = OtpN(d, fr, COL_OTP_BRUTO)
    Dim tipAmb As String: tipAmb = CStr(OtpC(d, fr, COL_OTP_TIP_AMB))
    Dim kolAmb As Double: kolAmb = OtpN(d, fr, COL_OTP_KOL_AMB)
    Dim kolBruto As Double
    If storedBruto > 0 Then
        kolBruto = storedBruto
    Else
        Dim crateW As Double
        crateW = PrNz(LookupValue(TBL_TIP_AMBALAZE, COL_TAMB_TIP, tipAmb, COL_TAMB_TEZINA))
        kolBruto = kol + kolAmb * crateW
    End If

    Dim stavke() As Variant: ReDim stavke(0 To 0, 0 To 6)
    stavke(0, 0) = Trim$(CStr(OtpC(d, fr, COL_OTP_VRSTA)) & " " & CStr(OtpC(d, fr, COL_OTP_SORTA)))
    stavke(0, 1) = CStr(OtpC(d, fr, COL_OTP_KLASA))
    stavke(0, 2) = cenNeto
    stavke(0, 3) = cenBruto
    stavke(0, 4) = kol
    stavke(0, 5) = kolBruto
    stavke(0, 6) = kol * cenNeto
    Dim cnt As Long: cnt = 1
    Dim osnovica As Double: osnovica = kol * cenNeto

    Dim stID As String: stID = CStr(OtpC(d, fr, COL_OTP_STANICA))

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
        If Len(objLine) > 0 Then objLine = objLine & "    Reg. br: " & objReg Else objLine = "Objekat reg. br: " & objReg
    End If
    h("objekat") = objLine
    h("brDok") = CStr(OtpC(d, fr, COL_OTP_BROJ))
    Dim datV As Variant: datV = OtpC(d, fr, COL_OTP_DATUM)
    If IsDate(datV) Then h("datum") = Format$(CDate(datV), "dd.mm.yyyy") Else h("datum") = CStr(datV)
    Dim stanicaNaziv As String: stanicaNaziv = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", stID, "Naziv")))
    If Len(stanicaNaziv) = 0 Then stanicaNaziv = stID
    h("stanica") = stanicaNaziv
    h("koop") = "": h("bpg") = "": h("ra" & ChrW(269) & "un") = ""

    Dim ambPoc As Long: ambPoc = GetStanicaAmbSaldo(stID, tipAmb)
    h("ambPocetno") = tipAmb & " x " & CStr(ambPoc)
    h("ambPrijem") = tipAmb & " x " & CStr(CLng(kolAmb))
    h("ambIzdavanje") = tipAmb & " x 0"
    h("ambSaldo") = tipAmb & " x " & CStr(CLng(ambPoc - kolAmb))

    h("stopa") = stopa
    h("osnovica") = osnovica
    h("nadoknada") = osnovica * stopa / 100
    h("ukupno") = osnovica + osnovica * stopa / 100
    h("rok") = DocConfigOr(CFG_OTKUP_ROK, OTKUP_ROK_DEFAULT)

    Dim kl As String: kl = DocConfigOr(CFG_OTKUP_KLAUZULA, OtkupKlauzulaDefault())
    kl = Replace(kl, "{BPG}", "", , , vbTextCompare)
    kl = Replace(kl, "{POLJOPRIVREDNIK}", "", , , vbTextCompare)
    kl = Replace(kl, "{RA" & ChrW(268) & "UN}", "", , , vbTextCompare)
    kl = Replace(kl, "{DATUM}", CStr(h("datum")), , , vbTextCompare)
    kl = Replace(kl, "{BROJ}", CStr(h("brDok")), , , vbTextCompare)
    h("klauzula") = kl

    h("naslov") = "OTPREMNICA"
    h("grupni") = "1"

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

    Dim R0 As Long, lastRow As Long
    R0 = WriteOtkupCopy(ws, 1, "", h, stavke, cnt, OL_THIRD_PT - OL_TOP_MARGIN_TRIM_PT)
    lastRow = WriteOtkupCopy(ws, R0, "", h, stavke, cnt, OL_THIRD_PT) - 1

    DocPageSetupThirdA4 ws, lastRow

    Application.ScreenUpdating = oldScreen
    Set FillOtpremnicaSablon = ws
    Exit Function
EH:
    Application.ScreenUpdating = oldScreen
    LogErr "modPrint.FillOtpremnicaSablon"
End Function

Public Sub EnsureOtpremnicaSablon()
    On Error GoTo EH
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_OTPREMNICA_SABLON)
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = WS_OTPREMNICA_SABLON
    Exit Sub
EH:
    LogErr "modPrint.EnsureOtpremnicaSablon"
End Sub

Private Function OtpC(ByVal d As Variant, ByVal r As Long, ByVal colName As String) As Variant
    Dim c As Long: c = GetColumnIndex(TBL_OTPREMNICA, colName)
    If c >= 1 Then OtpC = d(r, c) Else OtpC = ""
End Function

Private Function OtpN(ByVal d As Variant, ByVal r As Long, ByVal colName As String) As Double
    Dim v As Variant: v = OtpC(d, r, colName)
    If IsNumeric(v) Then OtpN = CDbl(v)
End Function

' ============================================================
' OTKUPNI LIST (zakonski) -- OtkupSablon, dva primerka jedan iznad drugog,
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
    mode = DocResolveMode(GetConfigValue(CFG_OTKUP_PRINT_MODE), "PDF")

    Select Case mode
        Case "PRINT", "PREVIEW"
            Dim ws As Worksheet: Set ws = FillOtkupSablon(otkupIDs)
            If Not ws Is Nothing Then DocPrintWs ws, mode
        Case "PDF"
            ExportOtkupniListPDF otkupIDs, True   ' default (prazno) -> tihi PDF
        ' OFF -> bez izlaza
    End Select
    Exit Sub
EH:
    LogErr "modPrint.OutputOtkupniList"
End Sub

' PDF otkupnog lista -> <workbook>\Otkupni listovi\OtkupniList_<brDok>.pdf
Public Function ExportOtkupniListPDF(ByVal otkupIDs As String, _
                                     Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = FillOtkupSablon(otkupIDs)
    If ws Is Nothing Then Exit Function

    Dim suff As String: suff = Replace(Replace(otkupIDs, " + ", "_"), "/", "-")
    Dim pdfPath As String: pdfPath = EnsureDocFolder(PDF_DIR_OTKUPNI) & "\OtkupniList_" & suff & ".pdf"

    DocExportPdf ws, pdfPath, openAfter
    ExportOtkupniListPDF = pdfPath
    Exit Function
EH:
    LogErr "modPrint.ExportOtkupniListPDF"
End Function

' Reprint celog otkupnog lista na osnovu jednog OtkupID-a. Klasa I i II dele isti
' BrDok (zasebni OtkupID po klasi) -> skupi SVE aktivne redove tog BrDok-a i
' odstampaj ih zajedno (kao posle unosa), pa list ne bude nepotpun. Izlaz po
' CFG_OTKUP_PRINT_MODE (kao OutputOtkupniList).
Public Sub ReprintOtkupniListByOtkupID(ByVal otkupID As String)
    On Error GoTo EH
    If Trim$(otkupID) = "" Then Exit Sub

    Dim d As Variant: d = GetTableData(TBL_OTKUP)
    If Not IsArray(d) Then Exit Sub
    d = ExcludeStornirano(d, TBL_OTKUP)
    If Not IsArray(d) Then Exit Sub

    Dim iID As Long, iBr As Long
    iID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    iBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    If iID = 0 Or iBr = 0 Then Exit Sub

    Dim brDok As String, r As Long
    For r = 1 To UBound(d, 1)
        If CStr(d(r, iID)) = otkupID Then brDok = CStr(d(r, iBr)): Exit For
    Next r

    Dim ids As String
    If brDok <> "" Then
        For r = 1 To UBound(d, 1)
            If CStr(d(r, iBr)) = brDok Then
                If ids = "" Then ids = CStr(d(r, iID)) Else ids = ids & " + " & CStr(d(r, iID))
            End If
        Next r
    End If
    If ids = "" Then ids = otkupID   ' fallback: bar taj red

    OutputOtkupniList ids
    Exit Sub
EH:
    LogErr "modPrint.ReprintOtkupniListByOtkupID"
End Sub

' Popuni OtkupSablon sa dva primerka. Vraca sheet (ili Nothing).
Private Function FillOtkupSablon(ByVal otkupIDs As String) As Worksheet
    On Error GoTo EH
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsureOtkupSablon
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_OTKUP_SABLON)
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
    h("ambPrijem") = tipAmb & " x " & CStr(CLng(kolAmb))         ' primljeno (tekuci blok)
    h("ambIzdavanje") = tipAmb & " x " & CStr(CLng(kolAmbIzd))   ' izdato (tekuci blok)
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
    kl = Replace(kl, "{RA" & ChrW(268) & "UN}", CStr(h("ra" & ChrW(269) & "un")), , , vbTextCompare)
    kl = Replace(kl, "{DATUM}", CStr(h("datum")), , , vbTextCompare)
    kl = Replace(kl, "{BROJ}", CStr(h("brDok")), , , vbTextCompare)
    h("klauzula") = kl
    h("naslov") = "OTKUPNI LIST"
    h("grupni") = ""

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
    Dim R0 As Long, lastRow As Long
    R0 = WriteOtkupCopy(ws, 1, "Primerak za poljoprivrednika", h, stavke, cnt, _
                        OL_THIRD_PT - OL_TOP_MARGIN_TRIM_PT)
    lastRow = WriteOtkupCopy(ws, R0, "Primerak za otkupljivaca", h, stavke, cnt, _
                             OL_THIRD_PT) - 1

    DocPageSetupThirdA4 ws, lastRow

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
Private Function WriteOtkupCopy(ByVal ws As Worksheet, ByVal R0 As Long, _
                                ByVal copyLbl As String, ByVal h As Object, _
                                ByVal stavke As Variant, ByVal nStavke As Long, _
                                ByVal targetPt As Double) As Long
    Dim rr As Long
    Dim grayClr As Long: grayClr = DocColGray()
    Dim fillClr As Long: fillClr = DocColHeaderFill()
    Dim usedPt As Double: usedPt = 0

    ' --- gornji prazan razmak (apsorbuje nestampajucu ivicu stampaca) ---
    ws.rows(R0).RowHeight = OL_TOP_SPACER_PT
    usedPt = usedPt + OL_TOP_SPACER_PT
    rr = R0 + 1

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
        .value = CStr(h("naslov")) & "  br. " & h("brDok")
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

    ' --- subjekt (1 red): poljoprivrednik (otkup) ili otkupno mesto (grupni/malina) ---
    With ws.cells(rr, 1)
        .value = IIf(CStr(h("grupni")) = "1", CStr(h("stanica")), CStr(h("koop")))
        .Font.Bold = True
    End With
    If CStr(h("grupni")) <> "1" Then
        ' BPG i tekuci racun samo na pojedinacnom otkupnom listu (proizvodjac);
        ' grupni otkupni list nema pojedinacnog proizvodjaca.
        DocLabelVal ws, rr, 4, "BPG:", CStr(h("bpg"))
        DocLabelVal ws, rr, 6, "TR:", CStr(h("ra" & ChrW(269) & "un"))
    End If
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
    DocLabelVal ws, ob + 2, 1, "Saldo ambala" & ChrW(382) & "e:", ""
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
    Set ws = ThisWorkbook.Sheets(WS_OTKUP_SABLON)
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = WS_OTKUP_SABLON
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
' GRUPNI OTKUPNI LIST (malina mod) - isti obrazac kao otkupni list (dva primerka,
' 1/3 A4), ali podaci se citaju iz PRIJEMNICE, a umesto proizvodjaca (kooperant +
' BPG + TR) stoji otkupno mesto (stanica = vozac u malina modu). Cena s prijemnice
' je BRUTO (kao otkup) -> prikaz neto + PDV nadoknada. Ambalaza = saldo na nivou
' stanice (entitet "Stanica"). Izlaz po CFG_OTKUP_PRINT_MODE (kao otkupni list).
' Okida frmDokumenta.btnUnosPrij posle snimanja prijemnice (samo u malina modu).
' prijemnicaIDs = rezultat SavePrijemnicaMulti_TX ("PRJ-1" ili "PRJ-1 + PRJ-2").
' ============================================================

' Glavni ulaz. Best-effort: greska se loguje, ne prekida tok prijemnice.
Public Sub OutputGrupniOtkupniList(ByVal prijemnicaIDs As String)
    On Error GoTo EH
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_OTKUP_PRINT_MODE), "PDF")

    Select Case mode
        Case "PRINT", "PREVIEW"
            Dim ws As Worksheet: Set ws = FillGrupniOtkupSablon(prijemnicaIDs)
            If Not ws Is Nothing Then DocPrintWs ws, mode
        Case "PDF"
            ExportGrupniOtkupniListPDF prijemnicaIDs, True   ' default (prazno) -> tihi PDF
        ' OFF -> bez izlaza
    End Select
    Exit Sub
EH:
    LogErr "modPrint.OutputGrupniOtkupniList"
End Sub

' PDF grupnog otkupnog lista -> <workbook>\Otkupni listovi\GrupniOtkupniList_<brPrij>.pdf
Public Function ExportGrupniOtkupniListPDF(ByVal prijemnicaIDs As String, _
                                           Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = FillGrupniOtkupSablon(prijemnicaIDs)
    If ws Is Nothing Then Exit Function

    Dim suff As String: suff = Replace(Replace(prijemnicaIDs, " + ", "_"), "/", "-")
    Dim pdfPath As String: pdfPath = EnsureDocFolder(PDF_DIR_OTKUPNI) & "\GrupniOtkupniList_" & suff & ".pdf"

    DocExportPdf ws, pdfPath, openAfter
    ExportGrupniOtkupniListPDF = pdfPath
    Exit Function
EH:
    LogErr "modPrint.ExportGrupniOtkupniListPDF"
End Function

' Popuni GrupniOtkupSablon iz podataka prijemnice. Vraca sheet (ili Nothing).
Private Function FillGrupniOtkupSablon(ByVal prijemnicaIDs As String) As Worksheet
    On Error GoTo EH
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsureGrupniOtkupSablon
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_GRUPNI_OTKUP_SABLON)
    On Error GoTo EH
    If ws Is Nothing Then Exit Function

    Dim d As Variant: d = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long, iDat As Long, iVoz As Long, iBr As Long
    Dim iVr As Long, iSo As Long, iKl As Long, iKol As Long, iCe As Long, iTip As Long
    Dim iKolAmb As Long, iKolAmbV As Long, iBruto As Long
    iID = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    iDat = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM)
    iVoz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VOZAC)
    iBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    iVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    iSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    iKl = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA)
    iKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    iCe = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA)
    iTip = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_TIP_AMB)
    iKolAmb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    iKolAmbV = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA)
    iBruto = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BRUTO)

    Dim ids() As String: ids = Split(prijemnicaIDs, " + ")
    Dim stavke() As Variant: ReDim stavke(0 To UBound(ids), 0 To 6)
    Dim cnt As Long: cnt = 0
    Dim osnovica As Double: osnovica = 0
    ' Cena na prijemnici je BRUTO (sadrzi PDV nadoknadu), kao na otkupu: prikazujemo
    ' NETO cenu/vrednost, a PDV nadoknadu kao posebnu stavku.
    Dim stopa As Double: stopa = PrNz(GetConfigValue(CFG_PDV_NADOKNADA_STOPA))
    If stopa <= 0 Then stopa = PDV_NADOKNADA_DEFAULT
    Dim vozID As String, brPrij As String, datum As String
    Dim tipAmb As String, kolAmb As Double, kolAmbVr As Double
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
                    ' tare gajbice (fallback za neto redove). Isto kao otkupni list.
                    Dim storedBruto As Double: storedBruto = 0
                    If iBruto > 0 Then storedBruto = PrNz(d(r, iBruto))
                    Dim kolBruto As Double
                    If storedBruto > 0 Then
                        kolBruto = storedBruto
                    Else
                        Dim crateW As Double
                        crateW = PrNz(LookupValue(TBL_TIP_AMBALAZE, COL_TAMB_TIP, CStr(d(r, iTip)), COL_TAMB_TEZINA))
                        kolBruto = kol + PrNz(d(r, iKolAmb)) * crateW
                    End If
                    stavke(cnt, 0) = Trim$(CStr(d(r, iVr)) & " " & CStr(d(r, iSo)))
                    stavke(cnt, 1) = CStr(d(r, iKl))
                    stavke(cnt, 2) = cenNeto        ' Cena bez PDV
                    stavke(cnt, 3) = cenBruto       ' Cena s PDV
                    stavke(cnt, 4) = kol            ' Kolicina neto
                    stavke(cnt, 5) = kolBruto       ' Kolicina bruto
                    stavke(cnt, 6) = kol * cenNeto  ' Vrednost neto
                    osnovica = osnovica + kol * cenNeto
                    kolAmb = kolAmb + PrNz(d(r, iKolAmb))   ' primljeno (sve klase)
                    If cnt = 0 Then
                        vozID = CStr(d(r, iVoz))
                        brPrij = CStr(d(r, iBr))
                        datum = Format$(d(r, iDat), "dd.mm.yyyy")
                        tipAmb = CStr(d(r, iTip))
                        If iKolAmbV > 0 Then kolAmbVr = PrNz(d(r, iKolAmbV))
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
    h("brDok") = brPrij
    h("datum") = datum
    ' Otkupno mesto = stanica (u malina modu VozacID == StanicaID). Naziv iz tblStanice;
    ' fallback na ime+prezime vozaca ako stanica nije nadjena.
    Dim stanicaNaziv As String
    stanicaNaziv = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", vozID, "Naziv")))
    If Len(stanicaNaziv) = 0 Then
        stanicaNaziv = Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", vozID, "Ime")) & " " & _
                             CStr(LookupValue(TBL_VOZACI, "VozacID", vozID, "Prezime")))
    End If
    h("stanica") = stanicaNaziv
    h("koop") = "": h("bpg") = "": h("ra" & ChrW(269) & "un") = ""

    ' Ambalaza - saldo na nivou stanice (entitet "Stanica"): Pocetno = saldo gajbica
    ' stanice iz tblAmbalaza (prijemnica pise "Kupac" stranu, ne pomera "Stanica");
    ' Primljeno/Izdato sa prijemnice; Saldo = Pocetno + Izdato - Primljeno (ista
    ' formula kao otkupni list).
    Dim ambPoc As Long: ambPoc = GetStanicaAmbSaldo(vozID, tipAmb)
    h("ambPocetno") = tipAmb & " x " & CStr(ambPoc)
    h("ambPrijem") = tipAmb & " x " & CStr(CLng(kolAmb))
    h("ambIzdavanje") = tipAmb & " x " & CStr(CLng(kolAmbVr))
    h("ambSaldo") = tipAmb & " x " & CStr(CLng(ambPoc + kolAmbVr - kolAmb))

    h("stopa") = stopa
    h("osnovica") = osnovica
    h("nadoknada") = osnovica * stopa / 100
    h("ukupno") = osnovica + osnovica * stopa / 100
    h("rok") = DocConfigOr(CFG_OTKUP_ROK, OTKUP_ROK_DEFAULT)

    ' Klauzula ostaje kako jeste; producent-tokeni ({POLJOPRIVREDNIK}/{RACUN}/{BPG})
    ' su prazni jer grupni list nema pojedinacnog proizvodjaca; datum/broj se popune.
    Dim kl As String: kl = DocConfigOr(CFG_OTKUP_KLAUZULA, OtkupKlauzulaDefault())
    kl = Replace(kl, "{BPG}", "", , , vbTextCompare)
    kl = Replace(kl, "{POLJOPRIVREDNIK}", "", , , vbTextCompare)
    kl = Replace(kl, "{RA" & ChrW(268) & "UN}", "", , , vbTextCompare)
    kl = Replace(kl, "{DATUM}", CStr(h("datum")), , , vbTextCompare)
    kl = Replace(kl, "{BROJ}", CStr(h("brDok")), , , vbTextCompare)
    h("klauzula") = kl

    h("naslov") = "GRUPNI OTKUPNI LIST"
    h("grupni") = "1"

    ' --- prep + dva primerka + PageSetup. DRZATI U SINHRONIZACIJI sa FillOtkupSablon
    '     (isti obrazac/geometrija 1/3 A4; renderuje isti WriteOtkupCopy). ---
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

    Dim R0 As Long, lastRow As Long
    R0 = WriteOtkupCopy(ws, 1, "", h, stavke, cnt, OL_THIRD_PT - OL_TOP_MARGIN_TRIM_PT)
    lastRow = WriteOtkupCopy(ws, R0, "", h, stavke, cnt, OL_THIRD_PT) - 1

    DocPageSetupThirdA4 ws, lastRow

    Application.ScreenUpdating = oldScreen
    Set FillGrupniOtkupSablon = ws
    Exit Function
EH:
    Application.ScreenUpdating = oldScreen
    LogErr "modPrint.FillGrupniOtkupSablon"
End Function

Public Sub EnsureGrupniOtkupSablon()
    On Error GoTo EH
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_GRUPNI_OTKUP_SABLON)
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = WS_GRUPNI_OTKUP_SABLON
    ws.columns("A").ColumnWidth = 6
    ws.columns("B").ColumnWidth = 24
    ws.columns("C").ColumnWidth = 8
    ws.columns("D").ColumnWidth = 16
    ws.columns("E").ColumnWidth = 12
    ws.columns("F").ColumnWidth = 16
    Exit Sub
EH:
    LogErr "modPrint.EnsureGrupniOtkupSablon"
End Sub

' ============================================================
' PRIJEMNICA -- PrijemnicaSablon, jedan A4 portrait dokument (prijem robe na
' hladnjacu). Izlaz po CFG_PRIJEMNICA_PRINT_MODE: OFF/prazno (DEFAULT, bez
' izlaza -- kao do sada) | PDF | PRINT | PREVIEW. Auto-izlaz okida
' frmDokumenta.btnUnosPrij posle snimanja. prijemnicaIDs = rezultat
' SavePrijemnicaMulti_TX ("PRJ-1" ili "PRJ-1 + PRJ-2" za dve klase).
' ============================================================

' Glavni ulaz (zove se posle SavePrijemnicaMulti_TX). Best-effort: greska se loguje.
Public Sub OutputPrijemnica(ByVal prijemnicaIDs As String)
    On Error GoTo EH
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_PRIJEMNICA_PRINT_MODE), "OFF")

    Select Case mode
        Case "PRINT", "PREVIEW"
            Dim ws As Worksheet: Set ws = FillPrijemnicaSablon(prijemnicaIDs)
            If Not ws Is Nothing Then DocPrintWs ws, mode
        Case "PDF"
            ExportPrijemnicaPDF prijemnicaIDs, True
        ' OFF / prazno (DEFAULT) -> bez izlaza; ponasanje kao do sada.
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

' PDF prijemnice -> <workbook>\Prijemnice\Prijemnica_<brPrij>.pdf. Vraca putanju.
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

    Dim folder As String: folder = EnsureDocFolder(PDF_DIR_PRIJEMNICE)
    Dim suff As String: suff = Replace(Replace(prijemnicaIDs, " + ", "_"), "/", "-")
    ' Vremenski pecat u imenu -> nema "file in use" (1004) ako je prethodni PDF otvoren.
    Dim pdfPath As String
    pdfPath = folder & "\Prijemnica_" & suff & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    DocExportPdf ws, pdfPath, openAfter
    ExportPrijemnicaPDF = pdfPath
    Exit Function
EH:
    MsgBox "Gre" & ChrW(353) & "ka pri izradi PDF prijemnice:" & vbCrLf & _
           "  [" & Err.Number & "] " & Err.description, vbExclamation, APP_NAME
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
    Set ws = ThisWorkbook.Sheets(WS_PRIJEMNICA_SABLON)
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
        DocLabelVal ws, rr, 1, "Vracena ambala" & ChrW(382) & "a (kom):", CStr(CLng(ambV))
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
    Set ws = ThisWorkbook.Sheets(WS_PRIJEMNICA_SABLON)
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = WS_PRIJEMNICA_SABLON
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

' ============================================================
' OM <-> KOOPERANT: REVERS kretanja prazne ambalaze - IzdAmbSablon, dva primerka
' po 1/3 A4 (kao otkupni list). Okida ga frmDokumenta.btnUnosOMUlaz posle uspesnog
' SaveOMUlaz_TX kad je aktivan "Izdato koop." (OM->koop; DOK_TIP_OM_IZLAZ_KOOP) ili
' "Prijem koop." (koop->OM povrat; DOK_TIP_OM_ULAZ_KOOP). Smer nosi parametar prijem
' (False=izdavanje, True=prijem) -> naslov/saldo/potpisi/ime PDF. Podaci iz forme
' (broj dok. opcioni). Izlaz po CFG_OM_IZDAVANJE_PRINT_MODE: PDF | PRINT | PREVIEW | OFF.
' ============================================================

' Glavni ulaz (zove se posle SaveOMUlaz_TX). Best-effort: greska se loguje.
Public Sub OutputIzdavanjeAmbalaze(ByVal datum As Date, ByVal brojDok As String, _
                                   ByVal omNaziv As String, ByVal omID As String, _
                                   ByVal koopNaziv As String, ByVal koopID As String, _
                                   ByVal tipAmb As String, ByVal kolAmb As Long, _
                                   ByVal vrstaVoca As String, _
                                   ByVal prijem As Boolean)
    On Error GoTo EH
    m_izdAmbPrijem = prijem
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_OM_IZDAVANJE_PRINT_MODE), "PDF")

    Select Case mode
        Case "PRINT", "PREVIEW"
            Dim ws As Worksheet
            Set ws = FillIzdavanjeAmbalazeSablon(datum, brojDok, omNaziv, omID, _
                                                 koopNaziv, koopID, tipAmb, kolAmb, vrstaVoca)
            If Not ws Is Nothing Then DocPrintWs ws, mode
        Case "PDF"
            ExportIzdavanjeAmbalazePDF datum, brojDok, omNaziv, omID, _
                                       koopNaziv, koopID, tipAmb, kolAmb, vrstaVoca, True
        ' OFF -> bez izlaza
    End Select
    Exit Sub
EH:
    LogErr "modPrint.OutputIzdavanjeAmbalaze"
End Sub

' PDF reversa -> <workbook>\Revers ambalaze\IzdavanjeAmbalaze_<suff>_<vreme>.pdf. Vraca putanju.
' Vremenski pecat u imenu -> nema "file in use" (1004) ako je prethodni otvoren.
Public Function ExportIzdavanjeAmbalazePDF(ByVal datum As Date, ByVal brojDok As String, _
                                           ByVal omNaziv As String, ByVal omID As String, _
                                           ByVal koopNaziv As String, ByVal koopID As String, _
                                           ByVal tipAmb As String, ByVal kolAmb As Long, _
                                           ByVal vrstaVoca As String, _
                                           Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = FillIzdavanjeAmbalazeSablon(datum, brojDok, omNaziv, omID, _
                                         koopNaziv, koopID, tipAmb, kolAmb, vrstaVoca)
    If ws Is Nothing Then
        MsgBox "PDF reversa (izdavanje ambala" & ChrW(382) & "e) nije napravljen: priprema lista nije uspela.", _
               vbExclamation, APP_NAME
        Exit Function
    End If

    Dim folder As String: folder = EnsureDocFolder(PDF_DIR_REVERS)
    Dim suff As String: suff = Trim$(brojDok)
    If suff = "" Then suff = koopID
    suff = Replace(Replace(suff, " + ", "_"), "/", "-")
    Dim pref As String: pref = IIf(m_izdAmbPrijem, "PrijemAmbalaze_", "IzdavanjeAmbalaze_")
    Dim pdfPath As String
    pdfPath = folder & "\" & pref & suff & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    DocExportPdf ws, pdfPath, openAfter
    ExportIzdavanjeAmbalazePDF = pdfPath
    Exit Function
EH:
    MsgBox "Gre" & ChrW(353) & "ka pri izradi PDF reversa (izdavanje ambala" & ChrW(382) & "e):" & vbCrLf & _
           "  [" & Err.Number & "] " & Err.description, vbExclamation, APP_NAME
    LogErr "modPrint.ExportIzdavanjeAmbalazePDF"
End Function

' Popuni IzdAmbSablon: dva identicna primerka jedan iznad drugog, svaki tacno
' 1/3 A4 (99mm) - ista geometrija kao otkupni list (papir sa dve perforacije;
' donja trecina ostaje prazna, stampa 1:1). Vraca sheet (ili Nothing). Podaci iz
' forme; saldo kooperanta se cita iz ledgera (POSLE upisa) -> novo stanje za tip.
Private Function FillIzdavanjeAmbalazeSablon(ByVal datum As Date, ByVal brojDok As String, _
                                             ByVal omNaziv As String, ByVal omID As String, _
                                             ByVal koopNaziv As String, ByVal koopID As String, _
                                             ByVal tipAmb As String, ByVal kolAmb As Long, _
                                             ByVal vrstaVoca As String) As Worksheet
    On Error GoTo EH
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsureIzdavanjeAmbalazeSablon
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_IZDAMB_SABLON)
    On Error GoTo EH
    If ws Is Nothing Then Exit Function

    ' --- podaci za list (zaglavlje firme iz configa, kao otkupni list) ---
    Dim h As Object: Set h = CreateObject("Scripting.Dictionary")
    h("name") = GetConfigValue("SELLER_NAME")
    h("pib") = GetConfigValue("SELLER_PIB")
    h("mb") = GetConfigValue("SELLER_MATICNI_BROJ")
    h("addr") = Trim$(GetConfigValue("SELLER_STREET") & ", " & _
                GetConfigValue("SELLER_POSTAL_CODE") & " " & GetConfigValue("SELLER_CITY"))
    h("acct") = GetConfigValue("SELLER_ACCOUNT")
    Dim objMesto As String: objMesto = Trim$(CStr(GetConfigValue("SELLER_OBJEKAT_MESTO")))
    Dim objReg As String: objReg = Trim$(CStr(GetConfigValue("SELLER_OBJEKAT_BR_REGISTRA")))
    Dim objLine As String
    If Len(objMesto) > 0 Then objLine = "Objekat: " & objMesto
    If Len(objReg) > 0 Then
        If Len(objLine) > 0 Then
            objLine = objLine & "    Reg. br: " & objReg
        Else
            objLine = "Objekat reg. br: " & objReg
        End If
    End If
    h("objekat") = objLine
    Dim brLbl As String: brLbl = Trim$(brojDok)
    If brLbl = "" Then brLbl = "-"
    h("brDok") = brLbl
    h("datum") = Format$(datum, "dd.mm.yyyy")
    h("stanica") = omNaziv
    h("koop") = Trim$(koopNaziv)
    Dim bpg As String
    If Trim$(koopID) <> "" Then _
        bpg = Trim$(CStr(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koopID, COL_KOOP_BPG)))
    h("bpg") = bpg
    h("tipAmb") = tipAmb
    h("kolAmb") = kolAmb
    h("vrsta") = Trim$(vrstaVoca)
    h("prijem") = m_izdAmbPrijem
    ' Saldo kod kooperanta: novo stanje (POSLE upisa) i izvedeno pocetno (pre dokumenta).
    ' delta na kooperantu: izdavanje +kolAmb (Ulaz), prijem/povrat -kolAmb (Izlaz).
    Dim saldoNovo As Long
    h("ambHaveSaldo") = IzdAmbSaldoVal(koopID, tipAmb, saldoNovo)
    If h("ambHaveSaldo") Then
        Dim ambDelta As Long
        If m_izdAmbPrijem Then ambDelta = -kolAmb Else ambDelta = kolAmb
        h("ambPocetno") = CStr(saldoNovo - ambDelta) & " kom"
        h("ambSaldo") = CStr(saldoNovo) & " kom"
    End If
    h("ambKretanje") = CStr(kolAmb) & " kom"

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
    ws.columns("C").ColumnWidth = 8
    ws.columns("D").ColumnWidth = 10
    ws.columns("E").ColumnWidth = 10
    ws.columns("F").ColumnWidth = 9
    ws.columns("G").ColumnWidth = 9
    ws.columns("H").ColumnWidth = 12
    ' NAPOMENA: bez AutoFit - visine redova su eksplicitne tako da svaki primerak
    ' zauzme tacno 1/3 A4 (99mm); AutoFit bi to ponistio.

    ' Dva identicna primerka, svaki dopunjen filler-redom do tacno 1/3 A4. Granice
    ' izmedju primeraka padaju na perforacije (99mm i 198mm); donja trecina prazna.
    Dim R0 As Long, lastRow As Long
    R0 = WriteIzdavanjeCopy(ws, 1, "Primerak: OM (otkupno mesto)", h, _
                            OL_THIRD_PT - OL_TOP_MARGIN_TRIM_PT)
    lastRow = WriteIzdavanjeCopy(ws, R0, "Primerak: kooperant", h, OL_THIRD_PT) - 1

    DocPageSetupThirdA4 ws, lastRow

    Application.ScreenUpdating = oldScreen
    Set FillIzdavanjeAmbalazeSablon = ws
    Exit Function
EH:
    Application.ScreenUpdating = oldScreen
    LogErr "modPrint.FillIzdavanjeAmbalazeSablon"
End Function

' Ispisuje jedan primerak reversa od reda R0 i dopunjava ga do tacno 1/3 A4
' (targetPt tacaka). Vraca prvi slobodan red (pocetak sledeceg primerka). Sve
' visine su eksplicitne (bez AutoFit) - isti pristup kao WriteOtkupCopy.
Private Function WriteIzdavanjeCopy(ByVal ws As Worksheet, ByVal R0 As Long, _
                                    ByVal copyLbl As String, ByVal h As Object, _
                                    ByVal targetPt As Double) As Long
    Dim rr As Long
    Dim grayClr As Long: grayClr = DocColGray()
    Dim fillClr As Long: fillClr = DocColHeaderFill()
    Dim usedPt As Double: usedPt = 0

    ' --- gornji prazan razmak (apsorbuje nestampajucu ivicu stampaca) ---
    ws.rows(R0).RowHeight = OL_TOP_SPACER_PT
    usedPt = usedPt + OL_TOP_SPACER_PT
    rr = R0 + 1

    ' --- zaglavlje prodavca (ime; adresa levo + PIB/MB/Ziro desno) ---
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

    ' --- objekat (samo ako je u configu) ---
    If Len(CStr(h("objekat"))) > 0 Then
        With ws.cells(rr, 1)
            .value = CStr(h("objekat"))
            .Font.Size = 9
        End With
        ws.rows(rr).RowHeight = 12#
        usedPt = usedPt + 12#
        rr = rr + 1
    End If

    ' --- naslov: mali descriptor + veliki REVERS ---
    ws.Range(ws.cells(rr, 1), ws.cells(rr, 8)).Merge
    With ws.cells(rr, 1)
        .value = IIf(h("prijem"), "Prijem (povrat) prazne ambala" & ChrW(382) & "e od kooperanta", "Izdavanje prazne ambala" & ChrW(382) & "e kooperantu")
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
        .value = "REVERS  br. " & CStr(h("brDok"))
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

    ' --- datum / otkupno mesto / oznaka primerka (desno) ---
    DocLabelVal ws, rr, 1, "Datum:", CStr(h("datum"))
    DocLabelVal ws, rr, 4, "Otkupno mesto:", CStr(h("stanica"))
    ws.Range(ws.cells(rr, 6), ws.cells(rr, 8)).Merge
    With ws.cells(rr, 6)
        .value = copyLbl
        .Font.Size = 8
        .Font.Color = grayClr
        .HorizontalAlignment = xlRight
    End With
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1

    ' --- kooperant (prima): ime levo, BPG (ako postoji) desno ---
    With ws.cells(rr, 1)
        .value = "Kooperant: " & CStr(h("koop"))
        .Font.Bold = True
    End With
    If Len(CStr(h("bpg"))) > 0 Then DocLabelVal ws, rr, 5, "BPG:", CStr(h("bpg"))
    ws.rows(rr).RowHeight = 13#
    usedPt = usedPt + 13#
    rr = rr + 1

    ' --- vrsta voca (samo ako je uneta) ---
    If Len(CStr(h("vrsta"))) > 0 Then
        DocLabelVal ws, rr, 1, "Vrsta vo" & ChrW(263) & "a:", CStr(h("vrsta"))
        ws.rows(rr).RowHeight = 12#
        usedPt = usedPt + 12#
        rr = rr + 1
    End If

    ' --- stavka: Rb | Tip ambalaze | Broj gajbica (kom) ---
    Dim hdr As Long: hdr = rr
    ws.cells(rr, 1).value = "Rb"
    ws.cells(rr, 2).value = "Tip ambala" & ChrW(382) & "e"
    ws.cells(rr, 5).value = "Broj gajbica (kom)"
    ws.Range(ws.cells(rr, 2), ws.cells(rr, 4)).Merge
    ws.Range(ws.cells(rr, 5), ws.cells(rr, 8)).Merge
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 8))
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Color = fillClr
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.rows(rr).RowHeight = 18#
    usedPt = usedPt + 18#
    rr = rr + 1
    ws.cells(rr, 1).value = 1
    ws.cells(rr, 2).value = CStr(h("tipAmb"))
    ws.cells(rr, 5).value = h("kolAmb")
    ws.Range(ws.cells(rr, 2), ws.cells(rr, 4)).Merge
    ws.Range(ws.cells(rr, 5), ws.cells(rr, 8)).Merge
    ws.cells(rr, 1).HorizontalAlignment = xlCenter
    ws.cells(rr, 5).HorizontalAlignment = xlCenter
    ws.rows(rr).RowHeight = 14#
    usedPt = usedPt + 14#
    rr = rr + 1
    ws.cells(rr, 2).value = "UKUPNO"
    ws.cells(rr, 5).value = h("kolAmb")
    ws.Range(ws.cells(rr, 2), ws.cells(rr, 4)).Merge
    ws.Range(ws.cells(rr, 5), ws.cells(rr, 8)).Merge
    ws.cells(rr, 5).HorizontalAlignment = xlCenter
    With ws.Range(ws.cells(rr, 1), ws.cells(rr, 8))
        .Font.Bold = True
        .Interior.Color = fillClr
    End With
    With ws.Range(ws.cells(hdr, 1), ws.cells(rr, 8)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    ws.rows(rr).RowHeight = 14#
    usedPt = usedPt + 14#
    rr = rr + 1

    ' --- ambalaza kod kooperanta: Pocetno stanje -> kretanje -> Novo saldo ---
    '     (entitetski saldo, ista logika kao otkupni list; novo = posle ovog dokumenta)
    If h("ambHaveSaldo") Then
        Dim kretLbl As String
        kretLbl = IIf(h("prijem"), "Primljeno (od kooperanta):", "Izdato (kooperantu):")
        DocLabelVal ws, rr, 1, "Pocetno stanje (tip " & CStr(h("tipAmb")) & "):", CStr(h("ambPocetno"))
        ws.rows(rr).RowHeight = 13#
        usedPt = usedPt + 13#
        rr = rr + 1
        DocLabelVal ws, rr, 1, kretLbl, CStr(h("ambKretanje"))
        ws.rows(rr).RowHeight = 13#
        usedPt = usedPt + 13#
        rr = rr + 1
        DocLabelVal ws, rr, 1, "Novo stanje (saldo):", CStr(h("ambSaldo"))
        ws.cells(rr, 1).Font.Bold = True
        ws.rows(rr).RowHeight = 13#
        usedPt = usedPt + 13#
        rr = rr + 1
    End If

    ' --- fleksibilni razmak: gura potpise tacno iznad perforacije ---
    Dim reservePt As Double: reservePt = 16# + OL_MIN_FILLER_PT
    Dim gapPt As Double: gapPt = targetPt - usedPt - reservePt
    If gapPt < 6# Then gapPt = 6#
    ws.rows(rr).RowHeight = gapPt
    usedPt = usedPt + gapPt
    rr = rr + 1

    ' --- potpisi ---
    ws.cells(rr, 1).value = IIf(h("prijem"), "Predao (kooperant):  ____________________", "Izdao (OM):  ____________________")
    ws.cells(rr, 1).Font.Size = 9
    ws.cells(rr, 1).Font.Color = grayClr
    ws.cells(rr, 5).value = IIf(h("prijem"), "Primio (OM):  ____________________", "Primio (kooperant):  ____________________")
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

    WriteIzdavanjeCopy = rr
End Function

' Entitetski saldo (kom) datog tipa kod kooperanta -> ByRef saldo; True ako je tip
' nadjen u ledgeru. GetAmbalazeStanje vraca 2D niz: (.,1)=Tip, (.,2)=saldo (Ulaz +,
' Izlaz -). Saldo se cita POSLE upisa pa odrazava novo stanje; pocetno = saldo - delta.
Private Function IzdAmbSaldoVal(ByVal koopID As String, ByVal tipAmb As String, _
                                ByRef saldo As Long) As Boolean
    On Error GoTo EH
    saldo = 0
    If Trim$(koopID) = "" Then Exit Function
    Dim st As Variant: st = GetAmbalazeStanje(koopID, "Kooperant")
    If Not IsArray(st) Then Exit Function
    Dim i As Long
    For i = LBound(st, 1) To UBound(st, 1)
        If Trim$(CStr(st(i, 1))) = Trim$(tipAmb) Then
            saldo = CLng(st(i, 2))
            IzdAmbSaldoVal = True
            Exit Function
        End If
    Next i
    Exit Function
EH:
    LogErr "modPrint.IzdAmbSaldoVal"
End Function

' Kreira IzdAmbSablon (8 kolona, kao OtkupSablon). Sme da ostane vidljiv u
' workbook-u, kao OtkupSablon / PrijemnicaSablon / PaletaSablon.
Public Sub EnsureIzdavanjeAmbalazeSablon()
    On Error GoTo EH
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_IZDAMB_SABLON)
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = WS_IZDAMB_SABLON
    ws.columns("A").ColumnWidth = 4
    ws.columns("B").ColumnWidth = 17
    ws.columns("C").ColumnWidth = 8
    ws.columns("D").ColumnWidth = 10
    ws.columns("E").ColumnWidth = 10
    ws.columns("F").ColumnWidth = 9
    ws.columns("G").ColumnWidth = 9
    ws.columns("H").ColumnWidth = 12
    Exit Sub
EH:
    LogErr "modPrint.EnsureIzdavanjeAmbalazeSablon"
End Sub

