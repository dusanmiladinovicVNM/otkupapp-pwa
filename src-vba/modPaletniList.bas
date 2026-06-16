Attribute VB_Name = "modPaletniList"
Option Explicit

' ============================================================
' modPaletniList — paletni list sveze robe + prerada (Phase 2)
'
' Inkrement 1: numeracija po godini (1..n, reset svake godine).
'   Pattern preuzet iz modFaktura.GenerateBrojFakture (per-year max+1),
'   NE iz modBrojevi (cija je kanon x/ddmmyy po stanici/danu).
'
' Paletizacija (ovde): PaletizePrijemnica (kapacitet-gajbica raspodela + rubna
'   paleta preko vise prijemnica), PrintPaletniList (reuse _Print konvencija iz
'   modPrint), PrintNepotpunePalete, kooperanti preko modSledljivost.TraceByZbirna.
'   Poziva se UNUTAR modDokumenta.SavePrijemnica_TX / SavePrijemnicaMulti_TX,
'   PRE CommitTx -> atomicno sa prijemnicom; print/PDF je post-commit side effect.
'   Stavka se kljuca po PrijemnicaID; ista prijemnica se ne paletizuje dvaput.
'
' Prerada: tblPrerada/tblPreradaStavka preko SavePrerada_TX; operater bira
'   palete (otvorene/zatvorene), funkcija markira Preradjeno=Da.
'
' Reuse: GetTableData / RequireColumnIndex / LogErr (postojeci helperi).
' Sema tabela: modSetup.EnsurePaletniListSchema (pokrenuti jednom).
' ============================================================

' Vraca sledeci redni broj palete za TEKUCU godinu (1 ako jos nema palete
' u ovoj godini). Prikaz na listu: BrojPalete & "/" & Godina.
Public Function GenerateBrojPalete() As Long
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PALETA)

    Dim yr As Long
    yr = Year(Date)

    Dim maxN As Long
    maxN = 0

    If Not IsEmpty(data) Then
        Dim iBroj As Long, iGod As Long
        iBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, "GenerateBrojPalete")
        iGod = RequireColumnIndex(TBL_PALETA, COL_PAL_GODINA, "GenerateBrojPalete")

        Dim r As Long, n As Long
        For r = 1 To UBound(data, 1)
            If CLng(Val(CStr(data(r, iGod)))) = yr Then
                n = CLng(Val(CStr(data(r, iBroj))))
                If n > maxN Then maxN = n
            End If
        Next r
    End If

    GenerateBrojPalete = maxN + 1
    Exit Function

EH:
    LogErr "modPaletniList.GenerateBrojPalete"
    GenerateBrojPalete = 0
End Function

' ============================================================
' PUBLIC — paletizacija iz modDokumenta.SavePrijemnica_TX / Multi_TX.
' Poziva se UNUTAR transakcije, PRE CommitTx -> atomicno sa prijemnicom.
' (TX vec drzi Calculation=manual, pa nema poseban calc-guard ovde.)
'
' Puni otvorenu paletu iste vrste/sorte/klase/tipa ambalaze gajbicama
' prijemnice. Kad paleta dostigne kapacitet -> zatvara je (ID ide u
' closedPalIDs za POST-commit izlaz). Jedna prijemnica moze da pregazi
' granicu (ostatak na novu paletu); jedna paleta skuplja stavke iz vise
' prijemnica (rubna paleta). Stavka se kljuca po PrijemnicaID.
'
' Greske se NE gutaju -> propagiraju u TX wrapper koji radi RollbackTx.
' ============================================================
Public Function PaletizePrijemnica( _
        ByVal prijemnicaID As String, ByVal brojPrij As String, _
        ByVal brojZbirne As String, ByVal vrstaVoca As String, _
        ByVal sortaVoca As String, ByVal klasa As String, _
        ByVal netoKg As Double, ByVal brGajbica As Long, _
        ByVal tipAmb As String, _
        Optional ByRef closedPalIDs As Collection = Nothing) As String

    Const SRC As String = "modPaletniList.PaletizePrijemnica"

    If brGajbica <= 0 Then Exit Function       ' nema gajbica (Klasa II / bez ambalaze)

    RequirePaletaSchema SRC
    RequirePaletaStavkaSchema SRC
    EnsurePrijemnicaNotAlreadyPaletized prijemnicaID, SRC

    Dim crateW As Double: crateW = GetTezinaGajbice(tipAmb)
    Dim defCap As Long: defCap = GetKapacitetPalete(vrstaVoca)

    Dim touched As Object: Set touched = CreateObject("Scripting.Dictionary")
    Dim nClosed As Long
    Dim remaining As Long: remaining = brGajbica

    Do While remaining > 0
        Dim palRow As Long, palID As String
        palID = GetOrCreateOpenPaleta(vrstaVoca, sortaVoca, klasa, tipAmb, defCap, palRow)
        If palID = "" Or palRow = 0 Then
            Err.Raise vbObjectError + 7331, SRC, _
                      "Ne mogu da otvorim/nadjem paletu za: " & vrstaVoca
        End If

        Dim used As Long, curNeto As Double, curAmb As Double
        Dim palKg As Double, cap As Long
        GetPaletaAggregates palRow, used, curNeto, curAmb, palKg, cap
        If cap <= 0 Then cap = defCap

        Dim freeSlots As Long: freeSlots = cap - used
        If freeSlots <= 0 Then
            ' puna a jos otvorena -> zatvori i otvori novu u sledecoj iteraciji
            ClosePaleta palRow, SRC
            If Not closedPalIDs Is Nothing Then closedPalIDs.Add palID
            nClosed = nClosed + 1
            GoTo NextIter
        End If

        Dim take As Long: take = remaining
        If take > freeSlots Then take = freeSlots

        Dim takeNeto As Double, takeAmb As Double
        takeNeto = netoKg * (take / brGajbica)
        takeAmb = take * crateW

        AddStavka palID, prijemnicaID, brojPrij, brojZbirne, klasa, _
                  vrstaVoca, sortaVoca, take, takeNeto, takeAmb

        Dim newGajb As Long: newGajb = used + take
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BR_GAJBICA, newGajb, SRC
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_NETO, curNeto + takeNeto, SRC
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_AMBALAZA, curAmb + takeAmb, SRC
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BRUTO, _
                          (curNeto + takeNeto) + (curAmb + takeAmb) + palKg, SRC

        touched(palID) = True
        remaining = remaining - take

        If newGajb >= cap Then
            ClosePaleta palRow, SRC
            If Not closedPalIDs Is Nothing Then closedPalIDs.Add palID
            nClosed = nClosed + 1
        End If
NextIter:
    Loop

    PaletizePrijemnica = "palete=" & touched.count & "; zatvoreno=" & nClosed & _
                         "; gajbica=" & brGajbica
End Function

' Idempotency guard: ista PrijemnicaID ne sme imati aktivnu (ne-storniranu)
' paletnu stavku -> sprecava dvostruku paletizaciju (retry/re-save).
Private Sub EnsurePrijemnicaNotAlreadyPaletized(ByVal prijemnicaID As String, _
                                                ByVal src As String)
    Dim d As Variant: d = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(d) Then Exit Sub

    Dim iPrij As Long, iStorno As Long
    iPrij = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID, src)
    iStorno = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, src)

    Dim r As Long
    For r = 1 To UBound(d, 1)
        If CStr(d(r, iPrij)) = prijemnicaID _
           And UCase$(Trim$(CStr(d(r, iStorno)))) <> "DA" Then
            Err.Raise vbObjectError + 7330, src, _
                      "Prijemnica " & prijemnicaID & " je vec paletizovana."
        End If
    Next r
End Sub

' POST-commit izlaz (print/PDF po modu) za palete zatvorene u paletizaciji.
' Best-effort: greska se loguje, ali NE rollback-uje iskomitovane podatke.
Public Sub PaletniListOutputClosed(ByVal closedPalIDs As Collection)
    If closedPalIDs Is Nothing Then Exit Sub

    Dim v As Variant
    For Each v In closedPalIDs
        On Error Resume Next
        OutputPaletniListByMode CStr(v)
        If Err.Number <> 0 Then
            LogErr "modPaletniList.PaletniListOutputClosed[" & CStr(v) & "]"
            Err.Clear
        End If
        On Error GoTo 0
    Next v
End Sub

' ============================================================
' PUBLIC — rucna stampa nepotpunih (otvorenih) paleta.
' Kraj smene: Alt+F8 -> PrintNepotpunePalete (kasnije dugme u UI).
' ============================================================
Public Sub PrintNepotpunePalete()
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PALETA)
    If IsEmpty(data) Then
        MsgBox "Nema paleta.", vbInformation, APP_NAME
        Exit Sub
    End If

    Dim iID As Long, iStatus As Long, iStorno As Long
    iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
    iStatus = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)
    iStorno = GetColumnIndex(TBL_PALETA, COL_STORNIRANO)

    Dim r As Long, cnt As Long
    For r = 1 To UBound(data, 1)
        If CStr(data(r, iStatus)) = PAL_STATUS_OTVORENA _
           And UCase$(CStr(data(r, iStorno))) <> "DA" Then
            OutputPaletniListByMode CStr(data(r, iID))
            cnt = cnt + 1
        End If
    Next r

    If cnt = 0 Then MsgBox "Nema otvorenih (nepotpunih) paleta.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "modPaletniList.PrintNepotpunePalete"
End Sub

' ============================================================
' PUBLIC — paletni list dokument preko PaletaSablon (isti pristup kao
' frmSledljivost.PrintTracePDF: Sablon + named-range fill + Export/Print).
' PaletaSablon se auto-kreira (EnsurePaletaSablon) i sme da se stilizuje —
' popunjavanje ide po imenima opsega, ne po poziciji.
' ============================================================

' Fizicka stampa jednog paletnog lista.
Public Sub PrintPaletniList(ByVal palID As String)
    On Error GoTo EH
    Dim broj As String, god As String
    Dim ws As Worksheet
    Set ws = FillPaletaSablon(palID, broj, god)
    If ws Is Nothing Then Exit Sub
    ws.PrintOut Copies:=1
    Exit Sub
EH:
    LogErr "modPaletniList.PrintPaletniList"
End Sub

' PDF jednog paletnog lista -> <workbook>\Paleta_<broj>-<god>.pdf. Vraca putanju.
Public Function ExportPaletniListPDF(ByVal palID As String, _
                                     Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim broj As String, god As String
    Dim ws As Worksheet
    Set ws = FillPaletaSablon(palID, broj, god)
    If ws Is Nothing Then Exit Function

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Paleta_" & broj & "-" & god & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, _
                           OpenAfterPublish:=openAfter

    ExportPaletniListPDF = pdfPath
    Exit Function
EH:
    LogErr "modPaletniList.ExportPaletniListPDF"
End Function

' "U bilo kom trenutku": Alt+F8 -> upises broj palete (tekuca godina) -> PDF.
Public Sub ExportPaletniListPDF_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj palete (godina " & Year(Date) & "):", "Paletni list -> PDF")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub

    Dim broj As Long
    broj = CLng(Val(ans))
    If broj <= 0 Then Exit Sub

    Dim palID As String
    palID = FindPaletaIDByBroj(broj, Year(Date))
    If palID = "" Then
        MsgBox "Nije nadjena paleta br. " & broj & "/" & Year(Date) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If

    ExportPaletniListPDF palID, True
    Exit Sub
EH:
    LogErr "modPaletniList.ExportPaletniListPDF_Prompt"
End Sub

' Auto-izlaz pri zatvaranju palete, po config-u PALETA_PRINT_MODE:
'   (prazno/PDF) -> tihi PDF | PRINT -> stampac | PREVIEW -> pregled | OFF -> nista
Private Sub OutputPaletniListByMode(ByVal palID As String)
    Dim mode As String
    mode = UCase$(Trim$(GetConfigValue(CFG_PALETA_PRINT_MODE)))

    Select Case mode
        Case "PRINT"
            PrintPaletniList palID
        Case "PREVIEW"
            Dim broj As String, god As String
            Dim ws As Worksheet
            Set ws = FillPaletaSablon(palID, broj, god)
            If Not ws Is Nothing Then ws.PrintPreview
        Case "PDF"
            ExportPaletniListPDF palID, False   ' tihi PDF, bez papira
        Case Else
            ' OFF ili prazno (DEFAULT) -> bez izlaza; snimanje ostaje trenutno.
            ' Auto-izlaz pune palete se ukljucuje rucno:
            '   SetConfigValue "PALETA_PRINT_MODE", "PDF" | "PRINT" | "PREVIEW"
    End Select
End Sub

' Popuni PaletaSablon (header + tezine + stavke). Vraca sheet + broj/god (ByRef).
Private Function FillPaletaSablon(ByVal palID As String, _
                                  ByRef brojOut As String, _
                                  ByRef godOut As String) As Worksheet
    EnsurePaletaSablon

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PaletaSablon")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long
    iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)

    Dim hRow As Long, r As Long
    For r = 1 To UBound(d, 1)
        If CStr(d(r, iID)) = palID Then hRow = r: Exit For
    Next r
    If hRow = 0 Then Exit Function

    brojOut = CStr(NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_BROJ))))
    godOut = CStr(NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_GODINA))))

    Application.ScreenUpdating = False

    ws.Range("PalBroj").value = brojOut & "/" & godOut
    ws.Range("PalDatum").value = Format$(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_DATUM)), "dd.mm.yyyy")
    ws.Range("PalVrsta").value = SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA))
    ws.Range("PalTip").value = SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_TIP_PALETE))
    ws.Range("PalStatus").value = SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_STATUS))
    ws.Range("PalGajbica").value = NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)))
    ws.Range("PalNeto").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_NETO)))
    ws.Range("PalAmbalaza").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_AMBALAZA)))
    ws.Range("PalPaleta").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_PALETA_KG)))
    ws.Range("PalBruto").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PALETA, COL_PAL_BRUTO)))

    Dim startRow As Long
    startRow = ws.Range("PalStavkaStart").row

    Dim lastRow As Long
    lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= startRow Then
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, 6)).Clear
    End If

    Dim s As Variant
    s = GetTableData(TBL_PALETA_STAVKA)
    Dim outR As Long, rb As Long
    outR = startRow
    rb = 0
    If Not IsEmpty(s) Then
        Dim sPal As Long, sPrij As Long, sZbr As Long, sGajb As Long, sNeto As Long
        sPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
        sPrij = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        sZbr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
        sGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
        sNeto = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO)

        Dim sr As Long
        For sr = 1 To UBound(s, 1)
            If CStr(SafeCell(s, sr, sPal)) = palID Then
                rb = rb + 1
                ws.cells(outR, 1).value = rb
                ws.cells(outR, 2).value = GetKooperantiZaZbirnu(CStr(SafeCell(s, sr, sZbr)))
                ws.cells(outR, 3).value = SafeCell(s, sr, sPrij)
                ws.cells(outR, 4).value = SafeCell(s, sr, sZbr)
                ws.cells(outR, 5).value = NzL(SafeCell(s, sr, sGajb))
                ws.cells(outR, 6).value = NzD(SafeCell(s, sr, sNeto))
                outR = outR + 1
            End If
        Next sr
    End If

    ' --- stilizacija stavki (kao SledljivostSablon: okviri + naizmenicne boje) ---
    Dim dataEnd As Long
    dataEnd = outR - 1
    If dataEnd >= startRow Then
        With ws.Range(ws.cells(startRow, 1), ws.cells(dataEnd, 6)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        ws.Range(ws.cells(startRow, 5), ws.cells(dataEnd, 5)).NumberFormat = "#,##0"
        ws.Range(ws.cells(startRow, 6), ws.cells(dataEnd, 6)).NumberFormat = "#,##0.00"

        Dim zr As Long
        For zr = 0 To dataEnd - startRow
            If zr Mod 2 = 1 Then
                ws.Range(ws.cells(startRow + zr, 1), _
                         ws.cells(startRow + zr, 6)).Interior.Color = RGB(217, 225, 242)
            End If
        Next zr
    End If

    ' --- footer (kao sledljivost): datum stampe + potpis/pecat ---
    Dim footRow As Long
    footRow = dataEnd + 2
    If footRow <= startRow Then footRow = startRow + 1
    ws.cells(footRow, 1).value = "Datum stampe: " & Format$(Date, "dd.mm.yyyy")
    ws.cells(footRow + 2, 1).value = "Potpis: ____________________"
    ws.cells(footRow + 2, 4).value = "Pecat: ____________________"

    ' --- portrait + sve kolone na JEDNU stranu po sirini (Neto kg se ne gubi) ---
    On Error Resume Next
    Application.PrintCommunication = False   ' batch PageSetup: bez sporog round-trip-a do (mreznog) stampaca
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(footRow + 2, 6)).Address
    End With
    Application.PrintCommunication = True
    On Error GoTo 0

    Application.ScreenUpdating = True
    Set FillPaletaSablon = ws
End Function

' Kreira PaletaSablon (labela + named-range + osnovni format) ako ne postoji.
' Postojeci se NE dira -> mozes ga slobodno stilizovati.
Public Sub EnsurePaletaSablon()
    On Error GoTo EH

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PaletaSablon")
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "PaletaSablon"

    ws.Range("A1:F1").Merge
    ws.Range("A1").value = "PALETNI LIST"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter

    ws.Range("A3").value = "Broj:"
    ws.Range("A4").value = "Datum:"
    ws.Range("A5").value = "Vrsta voca:"
    ws.Range("A6").value = "Tip palete:"
    ws.Range("A7").value = "Status:"
    ws.Range("D3").value = "Broj gajbica:"
    ws.Range("D4").value = "Neto (kg):"
    ws.Range("D5").value = "Ambalaza (kg):"
    ws.Range("D6").value = "Paleta (kg):"
    ws.Range("D7").value = "BRUTO (kg):"

    ws.Range("B3").name = "PalBroj"
    ws.Range("B4").name = "PalDatum"
    ws.Range("B5").name = "PalVrsta"
    ws.Range("B6").name = "PalTip"
    ws.Range("B7").name = "PalStatus"
    ws.Range("E3").name = "PalGajbica"
    ws.Range("E4").name = "PalNeto"
    ws.Range("E5").name = "PalAmbalaza"
    ws.Range("E6").name = "PalPaleta"
    ws.Range("E7").name = "PalBruto"

    ws.Range("A9").value = "Rb"
    ws.Range("B9").value = "Kooperant"
    ws.Range("C9").value = "Br. prijemnice"
    ws.Range("D9").value = "Br. zbirne"
    ws.Range("E9").value = "Gajbica"
    ws.Range("F9").value = "Neto kg"
    With ws.Range("A9:F9")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ws.Range("A10").name = "PalStavkaStart"

    ws.columns("A").ColumnWidth = 6
    ws.columns("B").ColumnWidth = 28
    ws.columns("C").ColumnWidth = 16
    ws.columns("D").ColumnWidth = 16
    ws.columns("E").ColumnWidth = 12
    ws.columns("F").ColumnWidth = 12

    ws.Range("D7,E7").Font.Bold = True
    ws.Range("A3:B7").Borders.LineStyle = xlContinuous
    ws.Range("D3:E7").Borders.LineStyle = xlContinuous
    Exit Sub

EH:
    LogErr "modPaletniList.EnsurePaletaSablon"
End Sub

' Nadji PaletaID po broju palete + godini (za Prompt).
Private Function FindPaletaIDByBroj(ByVal broj As Long, ByVal god As Long) As String
    Dim d As Variant
    d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long, iBroj As Long, iGod As Long
    iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
    iBroj = GetColumnIndex(TBL_PALETA, COL_PAL_BROJ)
    iGod = GetColumnIndex(TBL_PALETA, COL_PAL_GODINA)

    Dim r As Long
    For r = 1 To UBound(d, 1)
        If NzL(SafeCell(d, r, iBroj)) = broj And NzL(SafeCell(d, r, iGod)) = god Then
            FindPaletaIDByBroj = CStr(SafeCell(d, r, iID))
            Exit Function
        End If
    Next r
End Function

' ============================================================
' PRIVATE — paleta lifecycle + lookup + util
' ============================================================

' Vraca PaletaID otvorene palete za vrstu voca (i njen rowIndex preko
' outRow); ako je nema, kreira novu i vraca nju.
' Vraca PaletaID otvorene palete iste vrste/sorte/klase/tipa ambalaze (i njen
' rowIndex preko outRow); ako je nema, kreira novu. Pretraga (ne identitet):
' bira prvu (najstariju) podudarnu otvorenu paletu deterministicki.
Private Function GetOrCreateOpenPaleta(ByVal vrstaVoca As String, _
                                       ByVal sortaVoca As String, _
                                       ByVal klasa As String, _
                                       ByVal tipAmb As String, _
                                       ByVal capacity As Long, _
                                       ByRef outRow As Long) As String
    outRow = 0

    Dim data As Variant
    data = GetTableData(TBL_PALETA)

    If Not IsEmpty(data) Then
        Dim iVrsta As Long, iSorta As Long, iKlasa As Long, iTipAmb As Long
        Dim iStatus As Long, iStorno As Long, iID As Long, iPre As Long
        iVrsta = GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA)
        iSorta = GetColumnIndex(TBL_PALETA, COL_PAL_SORTA)
        iKlasa = GetColumnIndex(TBL_PALETA, COL_PAL_KLASA)
        iTipAmb = GetColumnIndex(TBL_PALETA, COL_PAL_TIP_AMBALAZE)
        iStatus = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)
        iStorno = GetColumnIndex(TBL_PALETA, COL_STORNIRANO)
        iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
        iPre = GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO)

        Dim r As Long
        For r = 1 To UBound(data, 1)
            If CStr(data(r, iVrsta)) = vrstaVoca _
               And CStr(SafeCell(data, r, iSorta)) = sortaVoca _
               And CStr(SafeCell(data, r, iKlasa)) = klasa _
               And CStr(SafeCell(data, r, iTipAmb)) = tipAmb _
               And CStr(data(r, iStatus)) = PAL_STATUS_OTVORENA _
               And UCase$(CStr(data(r, iStorno))) <> "DA" _
               And UCase$(Trim$(CStr(SafeCell(data, r, iPre)))) <> "DA" Then
                outRow = r
                GetOrCreateOpenPaleta = CStr(data(r, iID))
                Exit Function
            End If
        Next r
    End If

    GetOrCreateOpenPaleta = CreateNewPaleta(vrstaVoca, sortaVoca, klasa, _
                                            tipAmb, capacity, outRow)
End Function

Private Function CreateNewPaleta(ByVal vrstaVoca As String, _
                                 ByVal sortaVoca As String, _
                                 ByVal klasa As String, _
                                 ByVal tipAmb As String, _
                                 ByVal capacity As Long, _
                                 ByRef outRow As Long) As String
    Dim newID As String
    newID = GetNextID(TBL_PALETA, COL_PAL_ID, "PAL-")
    If newID = "" Then
        CreateNewPaleta = ""
        Exit Function
    End If

    Dim broj As Long: broj = GenerateBrojPalete()
    Dim tip As String: tip = GetConfigValue(CFG_DEFAULT_TIP_PALETE)
    Dim palKg As Double: palKg = GetTezinaPalete(tip)

    PalAppendRow TBL_PALETA, _
        Array(COL_PAL_ID, COL_PAL_BROJ, COL_PAL_GODINA, COL_PAL_DATUM, _
              COL_PAL_VRSTA, COL_PAL_SORTA, COL_PAL_KLASA, COL_PAL_TIP_AMBALAZE, _
              COL_PAL_TIP_PALETE, COL_PAL_KAPACITET, COL_PAL_BR_GAJBICA, _
              COL_PAL_NETO, COL_PAL_AMBALAZA, COL_PAL_PALETA_KG, COL_PAL_BRUTO, _
              COL_PAL_STATUS, COL_PAL_PRERADJENO, COL_PAL_CREATED, COL_STORNIRANO), _
        Array(newID, broj, Year(Date), Date, _
              vrstaVoca, sortaVoca, klasa, tipAmb, _
              tip, capacity, 0, _
              0, 0, palKg, palKg, _
              PAL_STATUS_OTVORENA, "", Now, "")

    outRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, newID)
    CreateNewPaleta = newID
End Function

Private Sub AddStavka(ByVal palID As String, ByVal prijemnicaID As String, _
                      ByVal brojPrij As String, ByVal brojZbirne As String, _
                      ByVal klasa As String, ByVal vrstaVoca As String, _
                      ByVal sortaVoca As String, ByVal gajbice As Long, _
                      ByVal neto As Double, ByVal amb As Double)
    Dim sid As String
    sid = GetNextID(TBL_PALETA_STAVKA, COL_PALS_ID, "PLS-")

    PalAppendRow TBL_PALETA_STAVKA, _
        Array(COL_PALS_ID, COL_PALS_PALETA_ID, COL_PALS_PRIJEMNICA_ID, _
              COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, COL_PALS_KLASA, _
              COL_PALS_VRSTA, COL_PALS_SORTA, COL_PALS_BR_GAJBICA, _
              COL_PALS_NETO, COL_PALS_AMBALAZA, COL_PALS_CREATED, COL_STORNIRANO), _
        Array(sid, palID, prijemnicaID, _
              brojPrij, brojZbirne, klasa, _
              vrstaVoca, sortaVoca, gajbice, _
              neto, amb, Now, "")
End Sub

Private Sub ClosePaleta(ByVal palRow As Long, ByVal src As String)
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_STATUS, PAL_STATUS_ZATVORENA, src
End Sub

' Generican append po imenu kolone (rowData velicine sa brojem kolona tabele).
Private Sub PalAppendRow(ByVal tblName As String, _
                         ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 9320, "PalAppendRow", "Nema tabele: " & tblName
    End If

    Dim n As Long
    n = lo.ListColumns.count

    Dim rowData() As Variant
    ReDim rowData(0 To n - 1)

    Dim i As Long, idx As Long
    For i = LBound(cols) To UBound(cols)
        idx = GetColumnIndex(tblName, CStr(cols(i)))
        If idx >= 1 And idx <= n Then rowData(idx - 1) = vals(i)
    Next i

    Dim newRow As Long
    newRow = AppendRow(tblName, rowData)
    If newRow = 0 Then
        Err.Raise vbObjectError + 9321, "PalAppendRow", _
                  "AppendRow nije uspeo za tabelu: " & tblName
    End If
End Sub

Private Sub GetPaletaAggregates(ByVal palRow As Long, ByRef used As Long, _
                                ByRef neto As Double, ByRef amb As Double, _
                                ByRef palk As Double, ByRef cap As Long)
    used = 0: neto = 0: amb = 0: palk = 0: cap = 0

    Dim d As Variant
    d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Sub
    If palRow < 1 Or palRow > UBound(d, 1) Then Exit Sub

    used = NzL(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)))
    neto = NzD(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_NETO)))
    amb = NzD(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_AMBALAZA)))
    palk = NzD(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_PALETA_KG)))
    cap = NzL(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_KAPACITET)))
End Sub

Private Function FindRowIndexByID(ByVal tblName As String, _
                                  ByVal colName As String, _
                                  ByVal idVal As String) As Long
    Dim c As Collection
    Set c = FindRows(tblName, colName, idVal)
    If c.count > 0 Then FindRowIndexByID = c.item(1) Else FindRowIndexByID = 0
End Function

Private Function GetKapacitetPalete(ByVal vrstaVoca As String) As Long
    Dim v As Variant
    v = LookupValue(TBL_KULTURE, "VrstaVoca", vrstaVoca, COL_KUL_GAJBICA_PALETA)

    Dim n As Long
    n = NzL(v)
    If n <= 0 Then n = PALETA_DEFAULT_KAPACITET
    GetKapacitetPalete = n
End Function

Private Function GetTezinaGajbice(ByVal tipAmb As String) As Double
    GetTezinaGajbice = NzD(LookupValue(TBL_TIP_AMBALAZE, COL_TAMB_TIP, tipAmb, COL_TAMB_TEZINA))
End Function

Private Function GetTezinaPalete(ByVal tip As String) As Double
    GetTezinaPalete = NzD(LookupValue(TBL_TIP_PALETE, COL_TPAL_TIP, tip, COL_TPAL_TEZINA))
End Function

' Distinct kooperanti za zbirnu (reuse modSledljivost.TraceByZbirna, kol.1).
Private Function GetKooperantiZaZbirnu(ByVal brojZbirne As String) As String
    On Error Resume Next

    Dim t As Variant
    t = TraceByZbirna(brojZbirne)
    If IsEmpty(t) Then Exit Function

    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim r As Long, k As String
    For r = LBound(t, 1) To UBound(t, 1)
        k = Trim$(CStr(t(r, 1)))
        If Len(k) > 0 Then
            If Not dic.Exists(k) Then dic.Add k, True
        End If
    Next r

    If dic.count > 0 Then GetKooperantiZaZbirnu = Join(dic.keys, ", ")
End Function

Private Function NzD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then NzD = CDbl(v)
End Function

Private Function NzL(ByVal v As Variant) As Long
    On Error Resume Next
    If IsNumeric(v) Then NzL = CLng(v)
End Function

' Bezbedno citanje celije iz GetTableData niza: ako kolona ne postoji
' (idx < 1, npr. schema drift), vrati Empty umesto subscript-error.
Private Function SafeCell(ByVal d As Variant, ByVal r As Long, _
                          ByVal idx As Long) As Variant
    If idx >= 1 Then SafeCell = d(r, idx) Else SafeCell = Empty
End Function

' ============================================================
' Schema guards (fail-fast) + exact-row lookup za paletni/prerada domen.
' Kriticne kolone idu preko RequireColumns (modSchemaGuard); ID lookup puca
' na 0 i na >1 -- nikad "prvi od duplikata".
' ============================================================
Private Sub RequirePaletaSchema(ByVal src As String)
    RequireColumns TBL_PALETA, src, _
        COL_PAL_ID, COL_PAL_BROJ, COL_PAL_GODINA, COL_PAL_DATUM, COL_PAL_VRSTA, _
        COL_PAL_SORTA, COL_PAL_KLASA, COL_PAL_TIP_AMBALAZE, COL_PAL_TIP_PALETE, _
        COL_PAL_KAPACITET, COL_PAL_BR_GAJBICA, COL_PAL_NETO, COL_PAL_AMBALAZA, _
        COL_PAL_PALETA_KG, COL_PAL_BRUTO, COL_PAL_STATUS, COL_PAL_PRERADJENO, _
        COL_PAL_CREATED, COL_STORNIRANO
End Sub

Private Sub RequirePaletaStavkaSchema(ByVal src As String)
    RequireColumns TBL_PALETA_STAVKA, src, _
        COL_PALS_ID, COL_PALS_PALETA_ID, COL_PALS_PRIJEMNICA_ID, _
        COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, COL_PALS_KLASA, _
        COL_PALS_VRSTA, COL_PALS_SORTA, COL_PALS_BR_GAJBICA, _
        COL_PALS_NETO, COL_PALS_AMBALAZA, COL_PALS_CREATED, COL_STORNIRANO
End Sub

Private Sub RequirePreradaSchema(ByVal src As String)
    RequireColumns TBL_PRERADA, src, _
        COL_PRE_ID, COL_PRE_BROJ, COL_PRE_GODINA, COL_PRE_NETO_IZLAZ, _
        COL_PRE_KUTIJE, COL_PRE_KESE, COL_STORNIRANO
End Sub

Private Sub RequirePreradaStavkaSchema(ByVal src As String)
    RequireColumns TBL_PRERADA_STAVKA, src, _
        COL_PRES_ID, COL_PRES_PRERADA_ID, COL_PRES_PALETA_ID, _
        COL_PRES_BROJ_PALETE, COL_PRES_NETO, COL_STORNIRANO
End Sub

' Exact-row lookup po kljucu. Puca ako nema reda (0) ili ima vise (>1).
' Za IDENTITET (PaletaID, PrijemnicaID), NE za pretragu otvorenih paleta.
Private Function RequireSingleRowIndexByKey(ByVal tblName As String, _
                                            ByVal keyCol As String, _
                                            ByVal keyValue As String, _
                                            ByVal src As String) As Long
    Dim hits As Collection
    Set hits = FindRows(tblName, keyCol, keyValue)
    If hits.count = 0 Then
        Err.Raise vbObjectError + 7320, src, _
                  "Nema reda u " & tblName & " za " & keyCol & "=" & keyValue & "."
    ElseIf hits.count > 1 Then
        Err.Raise vbObjectError + 7321, src, _
                  "Vise redova (" & hits.count & ") u " & tblName & " za " & _
                  keyCol & "=" & keyValue & "."
    End If
    RequireSingleRowIndexByKey = CLng(hits(1))
End Function

' ============================================================
' PRERADA (preradni list) — palete -> kutije/kese.
' Palete se markiraju Preradjeno=Da i izlaze iz lagera. Sopstveni broj
' 1..n po godini; PDF preko PreradaSablon (isti stil). Bez kalo racunice.
' ============================================================

' Broj prerade za tekucu godinu (1..n), mirror GenerateBrojPalete.
Public Function GenerateBrojPrerade() As Long
    On Error GoTo EH
    Dim data As Variant
    data = GetTableData(TBL_PRERADA)
    Dim yr As Long: yr = Year(Date)
    Dim maxN As Long: maxN = 0
    If Not IsEmpty(data) Then
        Dim iBroj As Long, iGod As Long
        iBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, "GenerateBrojPrerade")
        iGod = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, "GenerateBrojPrerade")
        Dim r As Long, n As Long
        For r = 1 To UBound(data, 1)
            If CLng(Val(CStr(data(r, iGod)))) = yr Then
                n = CLng(Val(CStr(data(r, iBroj))))
                If n > maxN Then maxN = n
            End If
        Next r
    End If
    GenerateBrojPrerade = maxN + 1
    Exit Function
EH:
    LogErr "modPaletniList.GenerateBrojPrerade"
    GenerateBrojPrerade = 0
End Function

' Snimi preradu: brojeviPaleta = niz brojeva paleta (tekuca godina).
' Markira palete Preradjeno=Da. Vraca PreradaID ("" na gresci/otkazu).
Public Function SavePrerada(ByVal datum As Date, ByVal brojeviPaleta As Variant, _
                            ByVal brojKutija As Long, ByVal brojKesa As Long, _
                            ByVal netoKolicina As Double) As String
    On Error GoTo EH

    If Not IsArray(brojeviPaleta) Then Exit Function

    Dim yr As Long: yr = Year(Date)
    Dim pal As Object
    Set pal = CreateObject("Scripting.Dictionary")   ' PaletaID -> NetoKg

    Dim i As Long, pbr As Long, pid As String
    For i = LBound(brojeviPaleta) To UBound(brojeviPaleta)
        pbr = CLng(Val(CStr(brojeviPaleta(i))))
        If pbr > 0 Then
            pid = FindPaletaIDByBroj(pbr, yr)
            If pid = "" Then
                MsgBox "Paleta " & pbr & "/" & yr & " ne postoji.", vbExclamation, APP_NAME
                Exit Function
            End If
            If IsPaletaPreradjena(pid) Then
                MsgBox "Paleta " & pbr & "/" & yr & " je vec preradjena.", vbExclamation, APP_NAME
                Exit Function
            End If
            If Not pal.Exists(pid) Then pal.Add pid, GetPaletaNum(pid, COL_PAL_NETO)
        End If
    Next i

    If pal.count = 0 Then
        MsgBox "Nije izabrana nijedna validna paleta.", vbExclamation, APP_NAME
        Exit Function
    End If

    Dim preID As String: preID = GetNextID(TBL_PRERADA, COL_PRE_ID, "PRE-")
    Dim brPre As Long: brPre = GenerateBrojPrerade()

    PalAppendRow TBL_PRERADA, _
        Array(COL_PRE_ID, COL_PRE_BROJ, COL_PRE_GODINA, COL_PRE_DATUM, _
              COL_PRE_NETO, COL_PRE_KUTIJE, COL_PRE_KESE, COL_STORNIRANO), _
        Array(preID, brPre, yr, datum, netoKolicina, brojKutija, brojKesa, "")

    Dim k As Variant
    For Each k In pal.keys
        Dim sid As String: sid = GetNextID(TBL_PRERADA_STAVKA, COL_PRES_ID, "PRS-")
        PalAppendRow TBL_PRERADA_STAVKA, _
            Array(COL_PRES_ID, COL_PRES_PRERADA_ID, COL_PRES_PALETA_ID, _
                  COL_PRES_BROJ_PALETE, COL_PRES_NETO), _
            Array(sid, preID, CStr(k), CLng(GetPaletaNum(CStr(k), COL_PAL_BROJ)), pal(k))
        MarkPaletaPreradjena CStr(k)
    Next k

    SavePrerada = preID
    Exit Function
EH:
    LogErr "modPaletniList.SavePrerada"
    SavePrerada = ""
End Function

' Alt+F8 ulaz: unos paleta + kutije/kese/neto -> snimi + PDF.
Public Sub SavePrerada_Prompt()
    On Error GoTo EH

    Dim sp As String
    sp = InputBox("Brojevi paleta za preradu (zarezom, npr. 1,2,5):", "Prerada")
    If Trim$(sp) = "" Then Exit Sub

    Dim parts() As String
    parts = Split(sp, ",")
    Dim brojevi() As Long
    ReDim brojevi(LBound(parts) To UBound(parts))
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        brojevi(i) = CLng(Val(Trim$(parts(i))))
    Next i

    Dim sk As String: sk = InputBox("Broj kutija:", "Prerada", "0")
    If StrPtr(sk) = 0 Then Exit Sub
    Dim se As String: se = InputBox("Broj kesa:", "Prerada", "0")
    If StrPtr(se) = 0 Then Exit Sub
    Dim sn As String: sn = InputBox("Neto izlaz (kg):", "Prerada", "0")
    If StrPtr(sn) = 0 Then Exit Sub

    Dim preID As String
    preID = SavePrerada(Date, brojevi, CLng(Val(sk)), CLng(Val(se)), _
                        CDbl(Val(Replace(sn, ",", "."))))
    If preID <> "" Then ExportPreradaPDF preID, True
    Exit Sub
EH:
    LogErr "modPaletniList.SavePrerada_Prompt"
End Sub

' PDF preradnog lista -> <workbook>\Prerada_<broj>-<god>.pdf.
Public Function ExportPreradaPDF(ByVal preID As String, _
                                 Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim broj As String, god As String
    Dim ws As Worksheet
    Set ws = FillPreradaSablon(preID, broj, god)
    If ws Is Nothing Then Exit Function

    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Prerada_" & broj & "-" & god & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, _
                           OpenAfterPublish:=openAfter

    ExportPreradaPDF = pdfPath
    Exit Function
EH:
    LogErr "modPaletniList.ExportPreradaPDF"
End Function

' Alt+F8: upisi broj prerade (tekuca godina) -> PDF.
Public Sub ExportPreradaPDF_Prompt()
    On Error GoTo EH
    Dim ans As String
    ans = InputBox("Broj prerade (godina " & Year(Date) & "):", "Preradni list -> PDF")
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub
    Dim broj As Long: broj = CLng(Val(ans))
    If broj <= 0 Then Exit Sub
    Dim preID As String: preID = FindPreradaIDByBroj(broj, Year(Date))
    If preID = "" Then
        MsgBox "Nije nadjena prerada br. " & broj & "/" & Year(Date) & ".", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    ExportPreradaPDF preID, True
    Exit Sub
EH:
    LogErr "modPaletniList.ExportPreradaPDF_Prompt"
End Sub

Private Function FindPreradaIDByBroj(ByVal broj As Long, ByVal god As Long) As String
    Dim d As Variant
    d = GetTableData(TBL_PRERADA)
    If IsEmpty(d) Then Exit Function
    Dim iID As Long, iBroj As Long, iGod As Long
    iID = GetColumnIndex(TBL_PRERADA, COL_PRE_ID)
    iBroj = GetColumnIndex(TBL_PRERADA, COL_PRE_BROJ)
    iGod = GetColumnIndex(TBL_PRERADA, COL_PRE_GODINA)
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If NzL(SafeCell(d, r, iBroj)) = broj And NzL(SafeCell(d, r, iGod)) = god Then
            FindPreradaIDByBroj = CStr(SafeCell(d, r, iID))
            Exit Function
        End If
    Next r
End Function

' Popuni PreradaSablon (header + izlaz + lista paleta) -> sheet + broj/god.
Private Function FillPreradaSablon(ByVal preID As String, _
                                   ByRef brojOut As String, _
                                   ByRef godOut As String) As Worksheet
    EnsurePreradaSablon

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PreradaSablon")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_PRERADA)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long: iID = GetColumnIndex(TBL_PRERADA, COL_PRE_ID)
    Dim hRow As Long, r As Long
    For r = 1 To UBound(d, 1)
        If CStr(d(r, iID)) = preID Then hRow = r: Exit For
    Next r
    If hRow = 0 Then Exit Function

    brojOut = CStr(NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_BROJ))))
    godOut = CStr(NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_GODINA))))

    Application.ScreenUpdating = False

    ws.Range("PreBroj").value = brojOut & "/" & godOut
    ws.Range("PreDatum").value = Format$(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_DATUM)), "dd.mm.yyyy")
    ws.Range("PreKutije").value = NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE)))
    ws.Range("PreKese").value = NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_KESE)))
    ws.Range("PreNeto").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_NETO)))

    Dim startRow As Long: startRow = ws.Range("PreStavkaStart").row
    Dim lastRow As Long: lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= startRow Then ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, 3)).Clear

    Dim s As Variant: s = GetTableData(TBL_PRERADA_STAVKA)
    Dim outR As Long, rb As Long
    outR = startRow: rb = 0
    If Not IsEmpty(s) Then
        Dim sPre As Long, sBroj As Long, sNeto As Long
        sPre = GetColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PRERADA_ID)
        sBroj = GetColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_BROJ_PALETE)
        sNeto = GetColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_NETO)
        Dim sr As Long
        For sr = 1 To UBound(s, 1)
            If CStr(SafeCell(s, sr, sPre)) = preID Then
                rb = rb + 1
                ws.cells(outR, 1).value = rb
                ws.cells(outR, 2).value = NzL(SafeCell(s, sr, sBroj)) & "/" & godOut
                ws.cells(outR, 3).value = NzD(SafeCell(s, sr, sNeto))
                outR = outR + 1
            End If
        Next sr
    End If

    Dim dataEnd As Long: dataEnd = outR - 1
    If dataEnd >= startRow Then
        With ws.Range(ws.cells(startRow, 1), ws.cells(dataEnd, 3)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        ws.Range(ws.cells(startRow, 3), ws.cells(dataEnd, 3)).NumberFormat = "#,##0.00"
        Dim zr As Long
        For zr = 0 To dataEnd - startRow
            If zr Mod 2 = 1 Then
                ws.Range(ws.cells(startRow + zr, 1), _
                         ws.cells(startRow + zr, 3)).Interior.Color = RGB(217, 225, 242)
            End If
        Next zr
    End If

    Dim footRow As Long: footRow = dataEnd + 2
    If footRow <= startRow Then footRow = startRow + 1
    ws.cells(footRow, 1).value = "Datum stampe: " & Format$(Date, "dd.mm.yyyy")
    ws.cells(footRow + 2, 1).value = "Potpis: ____________________"
    ws.cells(footRow + 2, 3).value = "Pecat: ____________________"

    On Error Resume Next
    Application.PrintCommunication = False   ' batch PageSetup: bez sporog round-trip-a do (mreznog) stampaca
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(footRow + 2, 3)).Address
    End With
    Application.PrintCommunication = True
    On Error GoTo 0

    Application.ScreenUpdating = True
    Set FillPreradaSablon = ws
End Function

' Kreira PreradaSablon ako ne postoji (NE dira postojeci -> stilizuj slobodno).
Public Sub EnsurePreradaSablon()
    On Error GoTo EH
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PreradaSablon")
    On Error GoTo EH
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = "PreradaSablon"

    ws.Range("A1:C1").Merge
    ws.Range("A1").value = "PRERADNI LIST"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").HorizontalAlignment = xlCenter

    ws.Range("A3").value = "Broj:"
    ws.Range("A4").value = "Datum:"
    ws.Range("A6").value = "Broj kutija:"
    ws.Range("A7").value = "Broj kesa:"
    ws.Range("A8").value = "Neto (kg):"
    ws.Range("A3,A4,A6,A7,A8").Font.Bold = True

    ws.Range("B3").name = "PreBroj"
    ws.Range("B4").name = "PreDatum"
    ws.Range("B6").name = "PreKutije"
    ws.Range("B7").name = "PreKese"
    ws.Range("B8").name = "PreNeto"
    ws.Range("A8,B8").Font.Bold = True

    ws.Range("A10").value = "Rb"
    ws.Range("B10").value = "Broj palete"
    ws.Range("C10").value = "Neto kg"
    With ws.Range("A10:C10")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ws.Range("A11").name = "PreStavkaStart"

    ws.columns("A").ColumnWidth = 8
    ws.columns("B").ColumnWidth = 18
    ws.columns("C").ColumnWidth = 14
    ws.Range("A3:B8").Borders.LineStyle = xlContinuous
    Exit Sub
EH:
    LogErr "modPaletniList.EnsurePreradaSablon"
End Sub

' --- tblPaleta helper-i za preradu ---
Private Function GetPaletaNum(ByVal palID As String, ByVal colName As String) As Double
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Function
    Dim iID As Long: iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
    Dim iCol As Long: iCol = GetColumnIndex(TBL_PALETA, colName)
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If CStr(SafeCell(d, r, iID)) = palID Then
            GetPaletaNum = NzD(SafeCell(d, r, iCol))
            Exit Function
        End If
    Next r
End Function

Private Function IsPaletaPreradjena(ByVal palID As String) As Boolean
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Function
    Dim iID As Long: iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
    Dim iP As Long: iP = GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO)
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If CStr(SafeCell(d, r, iID)) = palID Then
            IsPaletaPreradjena = (UCase$(Trim$(CStr(SafeCell(d, r, iP)))) = "DA")
            Exit Function
        End If
    Next r
End Function

Private Sub MarkPaletaPreradjena(ByVal palID As String)
    Dim rowIdx As Long
    rowIdx = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, palID)
    If rowIdx > 0 Then UpdateCell TBL_PALETA, rowIdx, COL_PAL_PRERADJENO, "Da"
End Sub
