Attribute VB_Name = "modDocStyle"
'Attribute VB_Name = "modDocStyle"
Option Explicit

' ============================================================
' modDocStyle - zajednicki stil za stampane obrasce
' (otkupni / paletni / preradni list): logo, zaglavlje firme,
' naslov u dve linije, "labela + podebljana vrednost".
' Funkcije su Public da bi ih delili modPrint i modPaletniList.
' ============================================================

' TEST SEAM (posmatra): poslednji izvoz u PDF -- "zahtev|predato|grana": da li je
' pozivalac trazio otvaranje (1/0), sta je DocExportPdf predao Excelu (1/0) i
' kojom granom je izvoz otisao (R = Range, W = Worksheet). Pise se samo u test
' rezimu; cita ga DocPdfTest, gasi DocPdfTestReset (i ResetSeamova).
Private mPdfTest As String

Public Function DocColHeaderFill() As Long
    DocColHeaderFill = RGB(217, 225, 242)
End Function

Public Function DocColGray() As Long
    DocColGray = RGB(90, 90, 90)
End Function

Public Function DocColRule() As Long
    DocColRule = RGB(110, 110, 110)
End Function

' GetConfigValue sa folbekom na podrazumevanu vrednost.
Public Function DocConfigOr(ByVal key As String, ByVal dflt As String) As String
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(GetConfigValue(key)))
    On Error GoTo 0
    If v = "" Then DocConfigOr = dflt Else DocConfigOr = v
End Function

' Podrazumevana klauzula o PDV nadoknadi (cl. 34 ZPDV). Override preko
' config kljuca OTKUP_KLAUZULA (vidi modConfig.CFG_OTKUP_KLAUZULA).
Public Function OtkupKlauzulaDefault() As String
    OtkupKlauzulaDefault = _
        "PDV nadoknada obracunata je u skladu sa clanom 34. Zakona o porezu na " & _
        "dodatu vrednost. Otkupljivac se obavezuje da poljoprivredniku isplati " & _
        "vrednost otkupljenih poljoprivrednih proizvoda uvecanu za iznos PDV " & _
        "nadoknade. Pravo na odbitak PDV nadoknade kao prethodnog poreza " & _
        "otkupljivac ostvaruje pod uslovom da je izvrsio isplatu poljoprivredniku " & _
        "na njegov teku" & ChrW(263) & "i ra" & ChrW(269) & "un."
End Function

' Putanja loga: config SELLER_LOGO_PATH, pa <workbook>\logo.png / logo.jpg. "" ako nema.
Public Function DocLogoPath() As String
    On Error Resume Next
    Dim p As String
    p = Trim$(CStr(GetConfigValue("SELLER_LOGO_PATH")))
    If p <> "" Then
        If Dir$(p) <> "" Then DocLogoPath = p: Exit Function
    End If
    Dim cand As String
    cand = ThisWorkbook.path & "\logo.png"
    If Dir$(cand) <> "" Then DocLogoPath = cand: Exit Function
    cand = ThisWorkbook.path & "\logo.jpg"
    If Dir$(cand) <> "" Then DocLogoPath = cand
End Function

' Ubacuje logo gore desno (preko zaglavlja). Tiho preskace ako ga nema.
Public Sub DocDrawLogo(ByVal ws As Worksheet, ByVal topRow As Long, ByVal rightCol As Long)
    On Error GoTo done
    Dim p As String: p = DocLogoPath()
    If p = "" Then Exit Sub

    Dim w As Double, hgt As Double
    w = 52: hgt = 40
    Dim rcell As Range: Set rcell = ws.cells(topRow, rightCol)
    Dim l As Double, t As Double
    l = rcell.Left + rcell.width - w
    If l < ws.cells(topRow, 1).Left Then l = ws.cells(topRow, 1).Left
    t = rcell.top

    ws.Shapes.AddPicture fileName:=p, LinkToFile:=msoFalse, _
                         SaveWithDocument:=msoTrue, _
                         Left:=l, top:=t, width:=w, Height:=hgt
done:
End Sub

' Upisuje "labela vrednost" u jednu celiju, sa podebljanim delom vrednosti.
Public Sub DocLabelVal(ByVal ws As Worksheet, ByVal rowIx As Long, ByVal colIx As Long, _
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

' Zaglavlje firme (naziv/adresa/PIB-MB-ziro) u koloni 1, logo gore desno,
' linija ispod. Cita SELLER_* iz configa. Vraca prvi slobodan red.
Public Function DocSellerHeader(ByVal ws As Worksheet, ByVal atRow As Long, _
                                ByVal LASTCOL As Long, ByVal rightCol As Long) As Long
    With ws.cells(atRow, 1)
        .value = GetConfigValue("SELLER_NAME")
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.cells(atRow + 1, 1).value = Trim$(GetConfigValue("SELLER_STREET") & ", " & _
        GetConfigValue("SELLER_POSTAL_CODE") & " " & GetConfigValue("SELLER_CITY"))
    With ws.cells(atRow + 2, 1)
        .value = "PIB: " & GetConfigValue("SELLER_PIB") & "    MB: " & _
                 GetConfigValue("SELLER_MATICNI_BROJ") & "    Ziro: " & _
                 GetConfigValue("SELLER_ACCOUNT")
        .Font.Size = 9
        .Font.Color = DocColGray()
    End With
    DocDrawLogo ws, atRow, rightCol
    With ws.Range(ws.cells(atRow + 2, 1), ws.cells(atRow + 2, LASTCOL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = DocColRule()
    End With
    DocSellerHeader = atRow + 3
End Function

' Naslov: sitan opis + veliki naslov, centrirano preko 1..lastCol, linija ispod.
' Vraca prvi slobodan red.
Public Function DocTitleBlock(ByVal ws As Worksheet, ByVal atRow As Long, _
                              ByVal LASTCOL As Long, ByVal descriptor As String, _
                              ByVal title As String) As Long
    ws.Range(ws.cells(atRow, 1), ws.cells(atRow, LASTCOL)).Merge
    With ws.cells(atRow, 1)
        .value = descriptor
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = DocColGray()
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.cells(atRow + 1, 1), ws.cells(atRow + 1, LASTCOL)).Merge
    With ws.cells(atRow + 1, 1)
        .value = title
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
    End With
    ws.rows(atRow + 1).RowHeight = 22
    With ws.Range(ws.cells(atRow + 1, 1), ws.cells(atRow + 1, LASTCOL)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = DocColRule()
    End With
    DocTitleBlock = atRow + 2
End Function

' ------------------------------------------------------------
' Zajednicki izlaz u PDF (isti parametri za sve stampane obrasce) --
' JEDINO mesto u kodu koje zove ExportAsFixedFormat. cilj je Worksheet
' (obrasci) ili Range (izvestaji izvoze UsedRange).
'
' openAfter je ZAHTEV pozivaoca; PDF_OTVORI_POSLE = NO ga gasi za sve
' obrasce. Fixture ga pinuje na NO: test prolaz i dalje pravi papir
' (226-228 broje fajlove), ali ne otvara viewer koji operateru uzima
' fokus dok radi na istom racunaru. Test rezim tu NIJE signal: BFP ga
' pali samo u pojedinim testovima, a otkupni listovi iz Test_OTK_*
' nastaju van njega (izmereno 19.10.2026, test 249).
'
' Izvoz ide RANO VEZANIM pozivom u svakoj grani, u istom obliku koji se
' vec kompajlirao (ws. / UsedRange. sa imenovanim argumentima). Poziv
' nad cilj As Object bi bio kasno vezan -- neproveren oblik (review #431,
' P1). Pozicioni poziv nije uzet: jedan pomeren zarez bi OpenAfterPublish
' tiho predao drugom parametru, a smer "otvara" test ne vozi kroz izvoz.
' ------------------------------------------------------------
Public Sub DocExportPdf(ByVal cilj As Object, ByVal pdfPath As String, _
                        ByVal openAfter As Boolean)
    Dim otvori As Boolean, ws As Worksheet, rng As Range
    otvori = openAfter And IsPdfOtvoriPosle()
    If IsTestMode() Then mPdfTest = IIf(openAfter, "1", "0") & "|" & IIf(otvori, "1", "0")
    If TypeOf cilj Is Range Then
        If IsTestMode() Then mPdfTest = mPdfTest & "|R"
        Set rng = cilj
        rng.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                                Quality:=xlQualityStandard, _
                                IncludeDocProperties:=False, OpenAfterPublish:=otvori
    Else
        If IsTestMode() Then mPdfTest = mPdfTest & "|W"
        Set ws = cilj
        ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                               Quality:=xlQualityStandard, _
                               IncludeDocProperties:=False, OpenAfterPublish:=otvori
    End If
End Sub

' TEST SEAM: "zahtev|predato|grana" poslednjeg izvoza ("" = nije bilo izvoza);
' grana je R (Range) ili W (Worksheet). Tvrdo gejtovan.
Public Function DocPdfTest() As String
    If Not IsTestMode() Then Exit Function
    DocPdfTest = mPdfTest
End Function

' TEST SEAM: brise zabelezen izvoz -- test pre poteza, i ResetSeamova.
Public Sub DocPdfTestReset()
    If Not IsTestMode() Then Exit Sub
    mPdfTest = ""
End Sub

' ------------------------------------------------------------
' PageSetup za 1/3-A4 obrasce sa dva primerka (otkupni / grupni /
' otpremnica / revers ambalaze). Stampa 1:1 (Zoom=100, sve margine 0
' osim 0.31" levo/desno) tako da eksplicitne visine redova padaju tacno
' na 99mm/198mm perforacije. PrintArea = A1:H<lastRow> (8 kolona).
' On Error + PrintCommunication=False: brz batch upis, radi i bez stampaca.
' Geometrija je deljena -> svi 1/3-A4 obrasci se menjaju na jednom mestu.
' ------------------------------------------------------------
Public Sub DocPageSetupThirdA4(ByVal ws As Worksheet, ByVal lastRow As Long)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = 100
        .LeftMargin = Application.InchesToPoints(0.31)
        .RightMargin = Application.InchesToPoints(0.31)
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = False
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(lastRow, 8)).Address
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
' Izlazni mod obrasca. Prizna OFF/PRINT/PREVIEW/PDF; prazno ili nepoznato
' -> defMode (default ponasanje tog dokumenta: "PDF" ili "OFF"). Centralizuje
' razliku u defaultu (otkupni/grupni/izdamb=PDF, prijemnica/paleta=OFF) i daje
' jedno mesto za dispatch po modu (DocPrintWs za stampu/pregled).
' ------------------------------------------------------------
Public Function DocResolveMode(ByVal mode As String, ByVal defMode As String) As String
    Dim m As String: m = UCase$(Trim$(mode))
    Select Case m
        Case "OFF", "PRINT", "PREVIEW", "PDF"
            DocResolveMode = m
        Case Else
            DocResolveMode = UCase$(Trim$(defMode))
    End Select
End Function

' Stampa ili pregled vec napunjenog sheeta: PREVIEW -> pregled, inace stampac.
Public Sub DocPrintWs(ByVal ws As Worksheet, ByVal mode As String)
    If UCase$(Trim$(mode)) = "PREVIEW" Then
        ws.PrintPreview
    Else
        ws.PrintOut Copies:=1
    End If
End Sub

