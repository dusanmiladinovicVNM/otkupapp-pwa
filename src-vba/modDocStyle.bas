Attribute VB_Name = "modDocStyle"
'Attribute VB_Name = "modDocStyle"
Option Explicit

' ============================================================
' modDocStyle - zajednicki stil za stampane obrasce
' (otkupni / paletni / preradni list): logo, zaglavlje firme,
' naslov u dve linije, "labela + podebljana vrednost".
' Funkcije su Public da bi ih delili modPrint i modPaletniList.
'
' Uz stil, ovde je i JEDINI izlaz na stampac/PDF (DocPrintWs /
' DocExportPdf) i provera pred stampu (DocEnsurePrintReady). Razlog:
' PageSetup je driver-zavisan - drajver podrazumevanog stampaca (tipicno
' stampac za etikete, koji nema A4) obara i velicinu papira i skaliranje
' na "no scaling", pa dokument izadje prelomljen bez ijedne izmene u
' aplikaciji. Videti DocEnsurePrintReady.
' ============================================================

' --- namera skaliranja poslednje pripremljenog obrasca ---
' Dva obrasca se ISKLJUCUJU i ne smeju se pomesati:
'   1:1  - 1/3-A4 obrasci (otkupni, grupni, otpremnica, revers): Zoom=100,
'          BEZ skaliranja, jer visine redova padaju na perforacije 99/198mm.
'          "Fit to page" bi ovde pomerio tekst sa perforacije = kvar.
'   FITW - tabelarni obrasci (specifikacije, izvestaji, kartice, palete):
'          sve kolone na jednu stranu po sirini.
' Namera se upisuje u DocPageSetupThirdA4 / DocPageSetupFitWide, koje se
' uvek izvrse neposredno pre izlaza, pa je citanje pouzdano. Sheet koji
' nije oznacen (sabloni sa inline PageSetup blokom) tretira se kao FITW -
' svi takvi koriste FitToPagesWide=1; jedini Zoom=100 u projektu je
' DocPageSetupThirdA4, a on nameru upisuje.
Private Const PS_ONE_TO_ONE As String = "1:1"
Private Const PS_FIT_WIDE As String = "FITW"
Private mPsIntentSheet As String
Private mPsIntent As String

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
' Zajednicki izlaz jednog sheeta u PDF (isti parametri za sve
' stampane obrasce). Zamena za rasute ExportAsFixedFormat pozive
' (modPrint / modPaletniList / modIzvestaj / modOtkupBlok / frmSledljivost).
' ------------------------------------------------------------
Public Sub DocExportPdf(ByVal ws As Worksheet, ByVal pdfPath As String, _
                        ByVal openAfter As Boolean)
    If Not DocEnsurePrintReady(ws, False) Then Exit Sub
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, OpenAfterPublish:=openAfter
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
    DocSetPrintIntent ws, PS_ONE_TO_ONE
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
' PageSetup za tabelarne obrasce koji MORAJU stati na jednu stranu po
' sirini (specifikacije, pregledi): A4, "Fit to 1 page wide", visina
' slobodna (vise strana po duzini je u redu).
'
' Zasto helper, a ne inline blok kao ranije - dve zamke koje su obarale
' "fit to page" tako da su poslednje kolone bezale na drugu stranu:
'   1) Application.PrintCommunication = False: PageSetup upisi se keshiraju,
'      a PrintArea/PrintTitleRows/Zoom pod njim znaju da padnu na 1004; uz
'      On Error Resume Next to prodje TIHO i sheet ostane na starom zoom-u
'      (= 100%). Zato ga ovde NE koristimo (isti hardening kao
'      FillPrijemnicaSablon, vidi komentar tamo).
'   2) Rucni page break zaostao na trajnom sablon-sheetu: Excel ga postuje
'      i kad je fit-to-page ukljucen -> ResetAllPageBreaks pre svega.
' Samo skaliranje (upis + provera + rucni folbek) radi DocApplyFitWide,
' koju koristi i provera pred stampu.
' ------------------------------------------------------------
Public Sub DocPageSetupFitWide(ByVal ws As Worksheet, ByVal pgOrient As Long, _
                               ByVal prArea As Range, ByVal titleRows As String, _
                               ByVal lrMargin As Double, ByVal tbMargin As Double)
    On Error Resume Next
    ws.ResetAllPageBreaks
    With ws.PageSetup
        .PrintArea = prArea.Address
        .PaperSize = xlPaperA4
        .Orientation = pgOrient
        .LeftMargin = Application.InchesToPoints(lrMargin)
        .RightMargin = Application.InchesToPoints(lrMargin)
        .TopMargin = Application.InchesToPoints(tbMargin)
        .BottomMargin = Application.InchesToPoints(tbMargin)
        .CenterHorizontally = True
        .CenterVertically = False
        If Len(titleRows) > 0 Then .PrintTitleRows = titleRows
    End With
    DocApplyFitWide ws
    DocSetPrintIntent ws, PS_FIT_WIDE
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
    If Not DocEnsurePrintReady(ws, True) Then Exit Sub
    If UCase$(Trim$(mode)) = "PREVIEW" Then
        ws.PrintPreview
    Else
        ws.PrintOut Copies:=1
    End If
End Sub

' ------------------------------------------------------------
' PROVERA PRED STAMPU - jedina kapija ka stampacu/PDF-u.
'
' Zasto postoji: vecina firmi drzi stampac za etikete i A4 stampac na
' ISTOM racunaru i prebacuje podrazumevani stampac izmedju njih. Excel
' PageSetup ide kroz drajver trenutno podrazumevanog stampaca, pa drajver
' etiketnog stampaca (rolna, nema A4) obara velicinu papira i prebacuje
' skaliranje na "no scaling". Podesavanje se pamti U SHEETU, pa ostaje
' oboreno i posle vracanja na A4 stampac. Rezultat: isti .xlsm, ista
' verzija, a dokument izadje prelomljen - bez ijedne izmene u aplikaciji.
'
' Sta radi (tim redom):
'   1) papir: ako nije A4, pokusa ispravku; ako drajver i dalje odbija A4,
'      to je sam stampac a ne podesavanje dokumenta -> pita operatera
'      (jedina poruka koja iskace, i samo u tom slucaju).
'   2) skaliranje: vrati ga na nameru obrasca (1:1 ili fit-to-width) i
'      PROVERI da je primljeno; ako nije, upise rucno izracunat zoom.
' U normalnom radu ne prikazuje nista.
'
' forPrinter: True za fizicku stampu/pregled, False za PDF (drugaciji
' tekst poruke - PDF ne ide na uredjaj, ali dobija njegovu velicinu strane).
' Vraca False samo ako operater odustane.
' ------------------------------------------------------------
Public Function DocEnsurePrintReady(ByVal ws As Worksheet, _
                                    ByVal forPrinter As Boolean) As Boolean
    On Error Resume Next
    DocEnsurePrintReady = True
    If ws Is Nothing Then Exit Function

    ' --- 1) velicina papira ---
    If ws.PageSetup.PaperSize <> xlPaperA4 Then ws.PageSetup.PaperSize = xlPaperA4
    If ws.PageSetup.PaperSize <> xlPaperA4 Then
        LogWarn "DocEnsurePrintReady", "Drajver podrazumevanog stampaca ne prihvata A4", _
                "sheet=" & ws.name & "; stampac=" & Application.ActivePrinter
        Dim posledica As String
        If forPrinter Then
            posledica = "AgriX dokument bi otisao na taj uredjaj i ne bi izasao ispravno."
        Else
            posledica = "PDF bi dobio velicinu strane tog uredjaja, a ne A4."
        End If
        If MsgBox("Podrazumevani stampac ne podrzava A4:" & vbCrLf & vbCrLf & _
                  "    " & Application.ActivePrinter & vbCrLf & vbCrLf & _
                  "To je najcesce stampac za etikete. " & posledica & vbCrLf & vbCrLf & _
                  "Prebaci podrazumevani stampac u Windows-u na A4 stampac pa " & _
                  "ponovi ovu stampu." & vbCrLf & vbCrLf & _
                  "Ipak nastaviti?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then
            DocEnsurePrintReady = False
            Exit Function
        End If
    End If

    ' --- 2) skaliranje prema nameri obrasca ---
    If DocPrintIntent(ws) = PS_ONE_TO_ONE Then
        DocApplyOneToOne ws
    Else
        DocApplyFitWide ws
    End If
    On Error GoTo 0
End Function

' Namera poslednje pripremljenog obrasca; neoznacen sheet -> FITW (vidi
' obrazlozenje u deklaracionoj sekciji).
Private Function DocPrintIntent(ByVal ws As Worksheet) As String
    If StrComp(mPsIntentSheet, ws.name, vbTextCompare) = 0 Then
        DocPrintIntent = mPsIntent
    Else
        DocPrintIntent = PS_FIT_WIDE
    End If
End Function

Private Sub DocSetPrintIntent(ByVal ws As Worksheet, ByVal intent As String)
    mPsIntentSheet = ws.name
    mPsIntent = intent
End Sub

' Stampa 1:1 - Zoom = 100, nikakvo skaliranje (1/3-A4 obrasci).
Private Sub DocApplyOneToOne(ByVal ws As Worksheet)
    On Error Resume Next
    Dim z As Variant
    z = ws.PageSetup.Zoom
    If VarType(z) <> vbBoolean Then
        If CLng(z) = 100 Then Exit Sub
    End If
    ws.PageSetup.Zoom = 100
End Sub

' Sve kolone na jednu stranu po sirini + PROVERA da je drajver primio.
' Folbek kad nije: procenat racunat iz stvarne sirine print-area
' (Range.Width je u tackama i ne trazi drajver) upisan kao fiksan .Zoom.
' Sve se cita sa sheeta, pa je rutina upotrebljiva i pri pripremi obrasca
' (DocPageSetupFitWide) i pri proveri pred stampu (DocEnsurePrintReady).
Private Sub DocApplyFitWide(ByVal ws As Worksheet)
    On Error Resume Next
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False

    ' "1 strana po sirini" nista ne znaci ako papir nije A4 - tada se
    ' odnosi na tudju velicinu papira (rolna, Letter, custom).
    Dim z As Variant
    z = ws.PageSetup.Zoom
    If ws.PageSetup.PaperSize = xlPaperA4 And VarType(z) = vbBoolean Then
        If z = False And ws.PageSetup.FitToPagesWide = 1 Then Exit Sub
    End If

    Dim prArea As Range
    If Len(ws.PageSetup.PrintArea) > 0 Then Set prArea = ws.Range(ws.PageSetup.PrintArea)
    If prArea Is Nothing Then Set prArea = ws.UsedRange
    If prArea Is Nothing Then Exit Sub

    Dim usableW As Double, needW As Double, pct As Long
    ' A4: portrait 210mm = 595.28pt, landscape 297mm = 841.89pt.
    If ws.PageSetup.Orientation = xlLandscape Then usableW = 841.89 Else usableW = 595.28
    usableW = usableW - ws.PageSetup.LeftMargin - ws.PageSetup.RightMargin
    usableW = usableW * 0.97   ' rezerva za hardversku (neispisivu) marginu drajvera
    needW = prArea.Width
    If usableW <= 0 Or needW <= 0 Then Exit Sub
    pct = 100
    If needW > usableW Then pct = Int(usableW / needW * 100)
    If pct < 10 Then pct = 10
    ws.PageSetup.Zoom = pct
End Sub

