Attribute VB_Name = "modPaletniList"
Option Explicit

' ============================================================
' modPaletniList — paletni list sveze robe + prerada (Phase 2)
'
' Inkrement 1: numeracija po godini (1..n, reset svake godine).
'   Pattern preuzet iz modFaktura.GenerateBrojFakture (per-year max+1),
'   NE iz modBrojevi (cija je kanon x/ddmmyy po stanici/danu).
'
' Inkrement 2 (ovde): OnPrijemnicaSaved (kapacitet-gajbica raspodela + rubna
'   paleta preko vise prijemnica/zbirnih), PrintPaletniList (reuse _Print
'   konvencija iz modPrint), PrintNepotpunePalete, kooperanti preko
'   modSledljivost.TraceByZbirna. Hook: modDokumenta.SavePrijemnica_TX /
'   SavePrijemnicaMulti_TX (posle CommitTx -> best-effort).
'
' Inkrement 3 (sledece): prerada (tblPrerada/tblPreradaStavka) + UI dugmad
'   (evidencija paleta, stampa nepotpune, prerada).
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
' PUBLIC — hook iz modDokumenta.SavePrijemnica_TX / Multi_TX
' (poziva se POSLE CommitTx -> bez rollback rizika; best-effort).
'
' Puni otvorenu paletu te vrste voca gajbicama prijemnice. Kad
' paleta dostigne kapacitet (tblKulture.GajbicaPoPaleti, default 240)
' -> zatvara je i stampa. Jedna prijemnica moze da pregazi granicu
' (ostatak ide na novu paletu); jedna paleta moze da skupi stavke iz
' vise prijemnica/zbirnih (rubna paleta) -> tblPaletaStavka.
' ============================================================
Public Sub OnPrijemnicaSaved(ByVal brojPrij As String, _
                             ByVal brojZbirne As String, _
                             ByVal vrstaVoca As String, _
                             ByVal netoKg As Double, _
                             ByVal brGajbica As Long, _
                             ByVal tipAmb As String)
    On Error GoTo EH

    If brGajbica <= 0 Then Exit Sub          ' nema gajbica (npr. Klasa II) -> nije na paleti

    Dim capacity As Long
    capacity = GetKapacitetPalete(vrstaVoca)

    Dim crateW As Double
    crateW = GetTezinaGajbice(tipAmb)        ' kg po gajbici (0 ako sifarnik prazan)

    Dim remaining As Long
    remaining = brGajbica

    Do While remaining > 0
        Dim palRow As Long
        Dim palID As String
        palID = GetOrCreateOpenPaleta(vrstaVoca, palRow)
        If palID = "" Or palRow = 0 Then Exit Sub

        Dim used As Long, curNeto As Double, curAmb As Double, palKg As Double
        GetPaletaAggregates palRow, used, curNeto, curAmb, palKg

        Dim freeSlots As Long
        freeSlots = capacity - used
        If freeSlots <= 0 Then
            ' bezbednosni slucaj: puna ali jos "Otvorena" -> zatvori i nastavi
            ClosePaleta palRow
            PrintPaletniList palID
            GoTo NextIter
        End If

        Dim take As Long
        take = remaining
        If take > freeSlots Then take = freeSlots

        Dim takeNeto As Double, takeAmb As Double
        takeNeto = netoKg * (take / brGajbica)
        takeAmb = take * crateW

        AddStavka palID, brojPrij, brojZbirne, take, takeNeto, takeAmb

        Dim newGajb As Long
        newGajb = used + take
        UpdateCell TBL_PALETA, palRow, COL_PAL_BR_GAJBICA, newGajb
        UpdateCell TBL_PALETA, palRow, COL_PAL_NETO, curNeto + takeNeto
        UpdateCell TBL_PALETA, palRow, COL_PAL_AMBALAZA, curAmb + takeAmb
        UpdateCell TBL_PALETA, palRow, COL_PAL_BRUTO, _
                   (curNeto + takeNeto) + (curAmb + takeAmb) + palKg

        remaining = remaining - take

        If newGajb >= capacity Then
            ClosePaleta palRow
            PrintPaletniList palID            ' auto-stampa pune palete
        End If
NextIter:
    Loop

    Exit Sub

EH:
    LogErr "modPaletniList.OnPrijemnicaSaved"
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
            PrintPaletniList CStr(data(r, iID))
            cnt = cnt + 1
        End If
    Next r

    If cnt = 0 Then MsgBox "Nema otvorenih (nepotpunih) paleta.", vbInformation, APP_NAME
    Exit Sub

EH:
    LogErr "modPaletniList.PrintNepotpunePalete"
End Sub

' ============================================================
' PUBLIC — stampa jednog paletnog lista (reuse _Print konvencija
' iz modPrint.PrintIzvestaj: _Print sheet -> PrintOut -> sakrij).
' ============================================================
Public Sub PrintPaletniList(ByVal palID As String)
    On Error GoTo EH

    Dim d As Variant
    d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Sub

    Dim iID As Long
    iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)

    Dim hRow As Long, r As Long
    hRow = 0
    For r = 1 To UBound(d, 1)
        If CStr(d(r, iID)) = palID Then hRow = r: Exit For
    Next r
    If hRow = 0 Then Exit Sub

    Dim broj As Variant, god As Variant, dat As Variant, vrsta As Variant, tip As Variant
    broj = d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_BROJ))
    god = d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_GODINA))
    dat = d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_DATUM))
    vrsta = d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA))
    tip = d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_TIP_PALETE))

    Dim gajb As Long, neto As Double, amb As Double, palk As Double, bruto As Double
    gajb = NzL(d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)))
    neto = NzD(d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_NETO)))
    amb = NzD(d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_AMBALAZA)))
    palk = NzD(d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_PALETA_KG)))
    bruto = NzD(d(hRow, GetColumnIndex(TBL_PALETA, COL_PAL_BRUTO)))

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("_Print")
    On Error GoTo EH
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "_Print"
    End If

    ws.Visible = xlSheetVisible
    ws.cells.Clear

    ws.Range("A1").value = "PALETNI LIST   br. " & CStr(broj) & "/" & CStr(god)
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True

    ws.Range("A3").value = "Datum:":        ws.Range("B3").value = Format$(dat, "d.m.yyyy")
    ws.Range("A4").value = "Vrsta voca:":   ws.Range("B4").value = vrsta
    ws.Range("A5").value = "Tip palete:":   ws.Range("B5").value = tip
    ws.Range("A6").value = "Broj gajbica:": ws.Range("B6").value = gajb

    ws.Range("A8").value = "Neto (kg):":     ws.Range("B8").value = Format$(neto, "#,##0.00")
    ws.Range("A9").value = "Ambalaza (kg):": ws.Range("B9").value = Format$(amb, "#,##0.00")
    ws.Range("A10").value = "Paleta (kg):":  ws.Range("B10").value = Format$(palk, "#,##0.00")
    ws.Range("A11").value = "BRUTO (kg):":   ws.Range("B11").value = Format$(bruto, "#,##0.00")
    ws.Range("A11:B11").Font.Bold = True

    ws.Range("A13").value = "Proizvodjaci / stavke:"
    ws.Range("A13").Font.Bold = True
    ws.Range("A14").value = "Kooperant"
    ws.Range("B14").value = "Br. prijemnice"
    ws.Range("C14").value = "Br. zbirne"
    ws.Range("D14").value = "Gajbica"
    ws.Range("E14").value = "Neto kg"
    ws.Range("A14:E14").Font.Bold = True

    Dim s As Variant
    s = GetTableData(TBL_PALETA_STAVKA)
    Dim outR As Long
    outR = 15
    If Not IsEmpty(s) Then
        Dim sPal As Long, sPrij As Long, sZbr As Long, sGajb As Long, sNeto As Long
        sPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
        sPrij = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
        sZbr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
        sGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
        sNeto = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO)

        Dim sr As Long
        For sr = 1 To UBound(s, 1)
            If CStr(s(sr, sPal)) = palID Then
                ws.cells(outR, 1).value = GetKooperantiZaZbirnu(CStr(s(sr, sZbr)))
                ws.cells(outR, 2).value = s(sr, sPrij)
                ws.cells(outR, 3).value = s(sr, sZbr)
                ws.cells(outR, 4).value = s(sr, sGajb)
                ws.cells(outR, 5).value = Format$(NzD(s(sr, sNeto)), "#,##0.00")
                outR = outR + 1
            End If
        Next sr
    End If

    ws.columns("A:E").AutoFit
    ws.PrintOut Copies:=1
    ws.Visible = xlSheetVeryHidden
    Exit Sub

EH:
    LogErr "modPaletniList.PrintPaletniList"
End Sub

' ============================================================
' PRIVATE — paleta lifecycle + lookup + util
' ============================================================

' Vraca PaletaID otvorene palete za vrstu voca (i njen rowIndex preko
' outRow); ako je nema, kreira novu i vraca nju.
Private Function GetOrCreateOpenPaleta(ByVal vrstaVoca As String, _
                                       ByRef outRow As Long) As String
    outRow = 0

    Dim data As Variant
    data = GetTableData(TBL_PALETA)

    If Not IsEmpty(data) Then
        Dim iVrsta As Long, iStatus As Long, iStorno As Long, iID As Long
        iVrsta = GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA)
        iStatus = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)
        iStorno = GetColumnIndex(TBL_PALETA, COL_STORNIRANO)
        iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)

        Dim r As Long
        For r = 1 To UBound(data, 1)
            If CStr(data(r, iVrsta)) = vrstaVoca _
               And CStr(data(r, iStatus)) = PAL_STATUS_OTVORENA _
               And UCase$(CStr(data(r, iStorno))) <> "DA" Then
                outRow = r
                GetOrCreateOpenPaleta = CStr(data(r, iID))
                Exit Function
            End If
        Next r
    End If

    GetOrCreateOpenPaleta = CreateNewPaleta(vrstaVoca, outRow)
End Function

Private Function CreateNewPaleta(ByVal vrstaVoca As String, _
                                 ByRef outRow As Long) As String
    Dim newID As String
    newID = GetNextID(TBL_PALETA, COL_PAL_ID, "PAL-")
    If newID = "" Then
        CreateNewPaleta = ""
        Exit Function
    End If

    Dim broj As Long
    broj = GenerateBrojPalete()

    Dim tip As String
    tip = GetConfigValue(CFG_DEFAULT_TIP_PALETE)

    Dim palKg As Double
    palKg = GetTezinaPalete(tip)

    PalAppendRow TBL_PALETA, _
        Array(COL_PAL_ID, COL_PAL_BROJ, COL_PAL_GODINA, COL_PAL_DATUM, _
              COL_PAL_VRSTA, COL_PAL_TIP_PALETE, COL_PAL_BR_GAJBICA, _
              COL_PAL_NETO, COL_PAL_AMBALAZA, COL_PAL_PALETA_KG, _
              COL_PAL_BRUTO, COL_PAL_STATUS, COL_STORNIRANO), _
        Array(newID, broj, Year(Date), Date, vrstaVoca, tip, 0, _
              0, 0, palKg, palKg, PAL_STATUS_OTVORENA, "")

    outRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, newID)
    CreateNewPaleta = newID
End Function

Private Sub AddStavka(ByVal palID As String, ByVal brojPrij As String, _
                      ByVal brojZbirne As String, ByVal gajbice As Long, _
                      ByVal neto As Double, ByVal amb As Double)
    Dim sid As String
    sid = GetNextID(TBL_PALETA_STAVKA, COL_PALS_ID, "PLS-")

    PalAppendRow TBL_PALETA_STAVKA, _
        Array(COL_PALS_ID, COL_PALS_PALETA_ID, COL_PALS_BROJ_PRIJ, _
              COL_PALS_BROJ_ZBIRNE, COL_PALS_BR_GAJBICA, COL_PALS_NETO, _
              COL_PALS_AMBALAZA), _
        Array(sid, palID, brojPrij, brojZbirne, gajbice, neto, amb)
End Sub

Private Sub ClosePaleta(ByVal palRow As Long)
    UpdateCell TBL_PALETA, palRow, COL_PAL_STATUS, PAL_STATUS_ZATVORENA
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

    AppendRow tblName, rowData
End Sub

Private Sub GetPaletaAggregates(ByVal palRow As Long, ByRef used As Long, _
                                ByRef neto As Double, ByRef amb As Double, _
                                ByRef palk As Double)
    used = 0: neto = 0: amb = 0: palk = 0

    Dim d As Variant
    d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Sub
    If palRow < 1 Or palRow > UBound(d, 1) Then Exit Sub

    used = NzL(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)))
    neto = NzD(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_NETO)))
    amb = NzD(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_AMBALAZA)))
    palk = NzD(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_PALETA_KG)))
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
