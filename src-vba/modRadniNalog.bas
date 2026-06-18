Attribute VB_Name = "modRadniNalog"
Option Explicit

' ============================================================
' modRadniNalog — radni nalozi unutar hladnjace (SKICA / skeleton).
'
' Pokriva operacije koje TRANSFORMISU lager, a koje modHladnjaca (stanje) i
' modPaletniList (paletizacija/prerada) ne modeluju:
'   - Prepakivanje  : palete -> nove palete (konsolidacija / zamena ambalaze)
'   - Sortiranje    : paleta -> vise izlaznih paleta po frakciji (Klasa I/II/izvoz)
'   - (Prerada)     : terminalno pakovanje (kutije/kese) — danas u modPaletniList;
'                     moze se KASNIJE unifikovati kao tip naloga (NE sada).
'
' OBLIK (svi tipovi dele isti model):
'   ULAZ  : 1..N izvornih paleta (trose se -> Preradjeno=Da, izlaze iz sirovog lagera)
'   IZLAZ : 0..M paleta (vracaju se u lager preko modPaletniList.CreatePaleta)
'           i/ili pakovanje (kutije/kese) i/ili otpad
'   KALO  = NetoUlaz - NetoIzlaz (nerazknjizen gubitak); otpad je MERENA linija.
'   randman po frakciji = NetoIzlaz(frakcija) / NetoUlaz   (izvestaj, TODO)
'
' KICMA: jedan clsTransaction (snapshot tblPaleta + tblRadniNalog + _Stavka);
'   exact-row validacija ulaza; izlazne palete preko POSTOJECEG CreatePaleta
'   (logika palete se NE duplira); EH rollback + re-raise. Read-modeli bez TX
'   (0-based 2D Variant). Storno preko modStorno.StornoRadniNalog_TX. Stanje u
'   modHladnjaca se NE pise -> computed-on-read (nove palete se pojave same).
'
' SLEDLJIVOST: izlazna paleta -> RN stavka (Izlaz) -> NalogID -> RN stavke (Ulaz)
'   -> ulazne palete -> tblPaletaStavka.BrojZbirne -> TraceByZbirna -> otkup.
'
' Reuse: GetTableData / GetColumnIndex / FindRows / GetNextID / RequireColumns /
'   RequireColumnIndex / RequireUpdateCell / clsTransaction / LogErr /
'   modPaletniList.CreatePaleta.
'
' FINALIZACIJA (vidi dno): EnsureRadniNalogSchema (modSetup) pokrenuti jednom;
'   frmRadniNalog (build-guide kao docs/frmPalete-build.md); promovisati lokalne
'   helpere (NzL/NzD/SafeCell/RnAppendRow/RequireSingleRowIndexByKey) u
'   modHelpers/modDataAccess — isti TODO kao u modPaletniList (review nalaz #1/#2).
' ============================================================

' ============================================================
' READ-MODELI (bez TX) — 0-based 2D Variant za ListBox.List ili Empty.
' ============================================================

' Nalozi za grid. Filteri: god=0 -> sve; tip "" -> sve. Stornirani se ne prikazuju.
' Kolone (0-based): 0 NalogID(skriveno),1 Broj/God,2 Datum,3 Tip,4 Vrsta,
' 5 NetoUlaz,6 NetoIzlaz,7 Kalo,8 Kalo%,9 Izvrsilac,10 Status.
Public Function GetRadniNaloziForGrid(Optional ByVal god As Long = 0, _
                                      Optional ByVal tip As String = "") As Variant
    On Error GoTo EH
    Dim d As Variant: d = GetTableData(TBL_RADNI_NALOG)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long, iBroj As Long, iGod As Long, iDat As Long, iTip As Long
    Dim iVr As Long, iNU As Long, iNI As Long, iKalo As Long, iKaloP As Long
    Dim iIzv As Long, iStat As Long, iStorno As Long
    iID = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_ID)
    iBroj = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_BROJ)
    iGod = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_GODINA)
    iDat = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_DATUM)
    iTip = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_TIP)
    iVr = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_VRSTA)
    iNU = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_NETO_ULAZ)
    iNI = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_NETO_IZLAZ)
    iKalo = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_KALO)
    iKaloP = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_KALO_PROC)
    iIzv = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_IZVRSILAC)
    iStat = GetColumnIndex(TBL_RADNI_NALOG, COL_RN_STATUS)
    iStorno = GetColumnIndex(TBL_RADNI_NALOG, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If UCase$(Trim$(CStr(SafeCell(d, r, iStorno)))) <> "DA" _
           And (god = 0 Or NzL(SafeCell(d, r, iGod)) = god) _
           And (tip = "" Or CStr(SafeCell(d, r, iTip)) = tip) Then
            rows.Add r
        End If
    Next r
    If rows.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To rows.count - 1, 0 To 10)
    Dim k As Long
    For k = 0 To rows.count - 1
        r = rows(k + 1)
        res(k, 0) = CStr(SafeCell(d, r, iID))
        res(k, 1) = NzL(SafeCell(d, r, iBroj)) & "/" & NzL(SafeCell(d, r, iGod))
        res(k, 2) = Format$(SafeCell(d, r, iDat), "dd.mm.yyyy")
        res(k, 3) = CStr(SafeCell(d, r, iTip))
        res(k, 4) = CStr(SafeCell(d, r, iVr))
        res(k, 5) = NzD(SafeCell(d, r, iNU))
        res(k, 6) = NzD(SafeCell(d, r, iNI))
        res(k, 7) = NzD(SafeCell(d, r, iKalo))
        res(k, 8) = NzD(SafeCell(d, r, iKaloP))
        res(k, 9) = CStr(SafeCell(d, r, iIzv))
        res(k, 10) = CStr(SafeCell(d, r, iStat))
    Next k

    GetRadniNaloziForGrid = res
    Exit Function
EH:
    LogErr "modRadniNalog.GetRadniNaloziForGrid"
End Function

' Stavke (ulaz + izlaz) jednog naloga. Kolone (0-based): 0 Smer,1 TipIzlaza,
' 2 Frakcija,3 Vrsta,4 Sorta,5 Gajbice,6 Kutije,7 Kese,8 Neto,9 PaletaID.
Public Function GetNalogStavkeForGrid(ByVal nalogID As String) As Variant
    On Error GoTo EH
    If Trim$(nalogID) = "" Then Exit Function
    Dim d As Variant: d = GetTableData(TBL_RN_STAVKA)
    If IsEmpty(d) Then Exit Function

    Dim iNal As Long, iSmer As Long, iTip As Long, iFr As Long, iVr As Long
    Dim iSo As Long, iGa As Long, iKu As Long, iKe As Long, iNe As Long
    Dim iPal As Long, iStorno As Long
    iNal = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_NALOG_ID)
    iSmer = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_SMER)
    iTip = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_TIP_IZLAZ)
    iFr = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_FRAKCIJA)
    iVr = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_VRSTA)
    iSo = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_SORTA)
    iGa = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_GAJBICE)
    iKu = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_KUTIJE)
    iKe = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_KESE)
    iNe = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_NETO)
    iPal = GetColumnIndex(TBL_RN_STAVKA, COL_RNS_PALETA_ID)
    iStorno = GetColumnIndex(TBL_RN_STAVKA, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If CStr(SafeCell(d, r, iNal)) = nalogID _
           And UCase$(Trim$(CStr(SafeCell(d, r, iStorno)))) <> "DA" Then
            rows.Add r
        End If
    Next r
    If rows.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To rows.count - 1, 0 To 9)
    Dim k As Long
    For k = 0 To rows.count - 1
        r = rows(k + 1)
        res(k, 0) = CStr(SafeCell(d, r, iSmer))
        res(k, 1) = CStr(SafeCell(d, r, iTip))
        res(k, 2) = CStr(SafeCell(d, r, iFr))
        res(k, 3) = CStr(SafeCell(d, r, iVr))
        res(k, 4) = CStr(SafeCell(d, r, iSo))
        res(k, 5) = NzL(SafeCell(d, r, iGa))
        res(k, 6) = NzL(SafeCell(d, r, iKu))
        res(k, 7) = NzL(SafeCell(d, r, iKe))
        res(k, 8) = NzD(SafeCell(d, r, iNe))
        res(k, 9) = CStr(SafeCell(d, r, iPal))
    Next k

    GetNalogStavkeForGrid = res
    Exit Function
EH:
    LogErr "modRadniNalog.GetNalogStavkeForGrid"
End Function

' Broj naloga za tekucu godinu (1..n). Mirror GenerateBrojPalete/Prerade.
Public Function GenerateBrojNaloga() As Long
    On Error GoTo EH
    Dim data As Variant: data = GetTableData(TBL_RADNI_NALOG)
    Dim yr As Long: yr = Year(Date)
    Dim maxN As Long: maxN = 0
    If Not IsEmpty(data) Then
        Dim iBroj As Long, iGod As Long
        iBroj = RequireColumnIndex(TBL_RADNI_NALOG, COL_RN_BROJ, "GenerateBrojNaloga")
        iGod = RequireColumnIndex(TBL_RADNI_NALOG, COL_RN_GODINA, "GenerateBrojNaloga")
        Dim r As Long, n As Long
        For r = 1 To UBound(data, 1)
            If CLng(Val(CStr(data(r, iGod)))) = yr Then
                n = CLng(Val(CStr(data(r, iBroj))))
                If n > maxN Then maxN = n
            End If
        Next r
    End If
    GenerateBrojNaloga = maxN + 1
    Exit Function
EH:
    LogErr "modRadniNalog.GenerateBrojNaloga"
    GenerateBrojNaloga = 0
End Function

' ============================================================
' TX — snimi radni nalog: trosi ulazne palete, proizvodi izlaze (palete /
' pakovanje / otpad), racuna kalo. Pun clsTransaction wrapper po kicmi.
' izlazi = Collection of clsNalogIzlaz. Bez MsgBox -> baca gresku (UI hvata).
' Vraca NalogID.
' ============================================================
Public Function SaveRadniNalog_TX( _
        ByVal tipNaloga As String, _
        ByVal ulazPaletaIDs As Collection, _
        ByVal izlazi As Collection, _
        Optional ByVal izvrsilac As String = "", _
        Optional ByVal smena As String = "", _
        Optional ByVal vremePocetak As Variant, _
        Optional ByVal vremeKraj As Variant, _
        Optional ByVal napomena As String = "") As String

    Const SRC As String = "modRadniNalog.SaveRadniNalog_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    ' --- validacija ulaza ---
    If Trim$(tipNaloga) = "" Then
        Err.Raise vbObjectError + 7380, SRC, "Tip naloga je prazan."
    End If
    If ulazPaletaIDs Is Nothing Then
        Err.Raise vbObjectError + 7381, SRC, "Nije izabrana ulazna paleta."
    End If
    If ulazPaletaIDs.count = 0 Then
        Err.Raise vbObjectError + 7381, SRC, "Nije izabrana ulazna paleta."
    End If
    If izlazi Is Nothing Then
        Err.Raise vbObjectError + 7382, SRC, "Nije definisan izlaz naloga."
    End If
    If izlazi.count = 0 Then
        Err.Raise vbObjectError + 7382, SRC, "Nije definisan izlaz naloga."
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_RADNI_NALOG
    tx.AddTableSnapshot TBL_RN_STAVKA

    RequireRadniNalogSchema SRC
    RequireRnStavkaSchema SRC

    Dim dPal As Variant: dPal = GetTableData(TBL_PALETA)
    Dim iNeto As Long, iPre As Long, iStorno As Long, iVrsta As Long
    iNeto = RequireColumnIndex(TBL_PALETA, COL_PAL_NETO, SRC)
    iPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, SRC)
    iStorno = RequireColumnIndex(TBL_PALETA, COL_STORNIRANO, SRC)
    iVrsta = RequireColumnIndex(TBL_PALETA, COL_PAL_VRSTA, SRC)

    Dim nalogID As String: nalogID = GetNextID(TBL_RADNI_NALOG, COL_RN_ID, "RN-")
    Dim brNalog As Long: brNalog = GenerateBrojNaloga()

    ' --- ULAZ: resolve (tacno jednom), validacija, oznaka potroseno, RN stavka ---
    Dim palRows As Object: Set palRows = CreateObject("Scripting.Dictionary")
    Dim netoUlaz As Double: netoUlaz = 0
    Dim vrstaUlaz As String: vrstaUlaz = ""

    Dim v As Variant
    For Each v In ulazPaletaIDs
        Dim pid As String: pid = Trim$(CStr(v))
        If pid <> "" And Not palRows.Exists(pid) Then
            Dim rIdx As Long
            rIdx = RequireSingleRowIndexByKey(TBL_PALETA, COL_PAL_ID, pid, SRC)
            If UCase$(Trim$(CStr(SafeCell(dPal, rIdx, iStorno)))) = "DA" Then
                Err.Raise vbObjectError + 7383, SRC, "Paleta " & pid & " je stornirana."
            End If
            If UCase$(Trim$(CStr(SafeCell(dPal, rIdx, iPre)))) = "DA" Then
                Err.Raise vbObjectError + 7384, SRC, _
                          "Paleta " & pid & " je vec iskoriscena/preradjena."
            End If

            palRows.Add pid, rIdx
            Dim palNeto As Double: palNeto = NzD(SafeCell(dPal, rIdx, iNeto))
            netoUlaz = netoUlaz + palNeto
            If vrstaUlaz = "" Then vrstaUlaz = CStr(SafeCell(dPal, rIdx, iVrsta))

            ' paleta izlazi iz sirovog lagera (isti marker kao prerada -> lager
            ' read-modeli je vise ne racunaju; storno vraca na "")
            RequireUpdateCell TBL_PALETA, rIdx, COL_PAL_PRERADJENO, "Da", SRC

            RnAppendRow TBL_RN_STAVKA, _
                Array(COL_RNS_ID, COL_RNS_NALOG_ID, COL_RNS_SMER, COL_RNS_TIP_IZLAZ, _
                      COL_RNS_PALETA_ID, COL_RNS_FRAKCIJA, COL_RNS_VRSTA, COL_RNS_SORTA, _
                      COL_RNS_GAJBICE, COL_RNS_KUTIJE, COL_RNS_KESE, COL_RNS_NETO, _
                      COL_RNS_CREATED, COL_STORNIRANO), _
                Array(GetNextID(TBL_RN_STAVKA, COL_RNS_ID, "RNS-"), nalogID, RN_SMER_ULAZ, "", _
                      pid, "", vrstaUlaz, "", _
                      0, 0, 0, palNeto, _
                      Now, "")
        End If
    Next v

    ' --- IZLAZ: nove palete (reuse CreatePaleta) i/ili pakovanje i/ili otpad ---
    Dim netoIzlaz As Double: netoIzlaz = 0
    Dim o As Variant
    For Each o In izlazi
        Dim iz As clsNalogIzlaz
        Set iz = o
        Dim newPal As String: newPal = ""

        Select Case iz.TipIzlaza
            Case RN_IZLAZ_PALETA
                ' logika kreiranja palete se NE duplira -> modPaletniList.CreatePaleta
                newPal = CreatePaleta(iz.VrstaVoca, iz.SortaVoca, iz.Frakcija, _
                                      iz.TipAmbalaze, iz.Kapacitet)
                If newPal = "" Then
                    Err.Raise vbObjectError + 7385, SRC, "Ne mogu da kreiram izlaznu paletu."
                End If
                Dim prow As Long
                prow = RequireSingleRowIndexByKey(TBL_PALETA, COL_PAL_ID, newPal, SRC)
                RequireUpdateCell TBL_PALETA, prow, COL_PAL_BR_GAJBICA, iz.BrojGajbica, SRC
                RequireUpdateCell TBL_PALETA, prow, COL_PAL_NETO, iz.NetoKg, SRC
                ' TODO(business): preracun COL_PAL_AMBALAZA/COL_PAL_BRUTO za izlaznu
                ' paletu (treba tezina gajbice/palete kao u PaletizePrijemnica).

            Case RN_IZLAZ_PAKOVANJE
                ' finalno pakovanje (kutije/kese) — bez nove palete
            Case RN_IZLAZ_OTPAD
                ' merena otpadna frakcija — bez nove palete
            Case Else
                Err.Raise vbObjectError + 7386, SRC, "Nepoznat tip izlaza: " & iz.TipIzlaza
        End Select

        RnAppendRow TBL_RN_STAVKA, _
            Array(COL_RNS_ID, COL_RNS_NALOG_ID, COL_RNS_SMER, COL_RNS_TIP_IZLAZ, _
                  COL_RNS_PALETA_ID, COL_RNS_FRAKCIJA, COL_RNS_VRSTA, COL_RNS_SORTA, _
                  COL_RNS_GAJBICE, COL_RNS_KUTIJE, COL_RNS_KESE, COL_RNS_NETO, _
                  COL_RNS_CREATED, COL_STORNIRANO), _
            Array(GetNextID(TBL_RN_STAVKA, COL_RNS_ID, "RNS-"), nalogID, RN_SMER_IZLAZ, iz.TipIzlaza, _
                  newPal, iz.Frakcija, iz.VrstaVoca, iz.SortaVoca, _
                  iz.BrojGajbica, iz.Kutije, iz.Kese, iz.NetoKg, _
                  Now, "")

        netoIzlaz = netoIzlaz + iz.NetoKg
    Next o

    ' --- kalo (otpad je MERENA izlazna linija; kalo = nerazknjizen ostatak) ---
    Dim kalo As Double: kalo = netoUlaz - netoIzlaz
    Dim kaloProc As Double
    If netoUlaz > 0 Then kaloProc = kalo / netoUlaz * 100

    ' --- zaglavlje naloga ---
    Dim vp As Variant, vk As Variant
    If IsMissing(vremePocetak) Then vp = Empty Else vp = vremePocetak
    If IsMissing(vremeKraj) Then vk = Empty Else vk = vremeKraj

    RnAppendRow TBL_RADNI_NALOG, _
        Array(COL_RN_ID, COL_RN_BROJ, COL_RN_GODINA, COL_RN_DATUM, COL_RN_TIP, _
              COL_RN_VRSTA, COL_RN_NETO_ULAZ, COL_RN_NETO_IZLAZ, COL_RN_KALO, _
              COL_RN_KALO_PROC, COL_RN_IZVRSILAC, COL_RN_SMENA, COL_RN_VREME_POCETAK, _
              COL_RN_VREME_KRAJ, COL_RN_STATUS, COL_RN_NAPOMENA, COL_RN_CREATED, _
              COL_STORNIRANO), _
        Array(nalogID, brNalog, Year(Date), Date, tipNaloga, _
              vrstaUlaz, netoUlaz, netoIzlaz, kalo, _
              kaloProc, izvrsilac, smena, vp, _
              vk, RN_STATUS_ZAVRSEN, napomena, Now, _
              "")

    tx.CommitTx
    SaveRadniNalog_TX = nalogID
    Exit Function

EH:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function

' ============================================================
' PRIVATE — schema guards + helperi.
' TODO(dedup, review #1/#2): RnAppendRow -> modDataAccess.AppendRowByNames,
'   RequireSingleRowIndexByKey + NzL/NzD/SafeCell -> modHelpers. Lokalno za sada
'   (kao u modPaletniList) da modul kompajlira samostalno, bez sirih izmena.
' ============================================================
Private Sub RequireRadniNalogSchema(ByVal src As String)
    RequireColumns TBL_RADNI_NALOG, src, _
        COL_RN_ID, COL_RN_BROJ, COL_RN_GODINA, COL_RN_DATUM, COL_RN_TIP, _
        COL_RN_VRSTA, COL_RN_NETO_ULAZ, COL_RN_NETO_IZLAZ, COL_RN_KALO, _
        COL_RN_KALO_PROC, COL_RN_IZVRSILAC, COL_RN_SMENA, COL_RN_VREME_POCETAK, _
        COL_RN_VREME_KRAJ, COL_RN_STATUS, COL_RN_NAPOMENA, COL_RN_CREATED, COL_STORNIRANO
End Sub

Private Sub RequireRnStavkaSchema(ByVal src As String)
    RequireColumns TBL_RN_STAVKA, src, _
        COL_RNS_ID, COL_RNS_NALOG_ID, COL_RNS_SMER, COL_RNS_TIP_IZLAZ, _
        COL_RNS_PALETA_ID, COL_RNS_FRAKCIJA, COL_RNS_VRSTA, COL_RNS_SORTA, _
        COL_RNS_GAJBICE, COL_RNS_KUTIJE, COL_RNS_KESE, COL_RNS_NETO, _
        COL_RNS_CREATED, COL_STORNIRANO
End Sub

' Exact-row lookup po kljucu. Puca na 0 i na >1 (identitet, ne pretraga).
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

' Generican append po imenu kolone (mirror modPaletniList.PalAppendRow).
Private Sub RnAppendRow(ByVal tblName As String, _
                        ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 9330, "RnAppendRow", "Nema tabele: " & tblName
    End If

    Dim n As Long: n = lo.ListColumns.count
    Dim rowData() As Variant
    ReDim rowData(0 To n - 1)

    Dim i As Long, idx As Long
    For i = LBound(cols) To UBound(cols)
        idx = GetColumnIndex(tblName, CStr(cols(i)))
        If idx >= 1 And idx <= n Then rowData(idx - 1) = vals(i)
    Next i

    If AppendRow(tblName, rowData) = 0 Then
        Err.Raise vbObjectError + 9331, "RnAppendRow", _
                  "AppendRow nije uspeo za tabelu: " & tblName
    End If
End Sub

Private Function NzD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then NzD = CDbl(v)
End Function

Private Function NzL(ByVal v As Variant) As Long
    On Error Resume Next
    If IsNumeric(v) Then NzL = CLng(v)
End Function

Private Function SafeCell(ByVal d As Variant, ByVal r As Long, _
                          ByVal idx As Long) As Variant
    If idx >= 1 Then SafeCell = d(r, idx) Else SafeCell = Empty
End Function

' ============================================================
' FINALIZACIJA — koraci da skica postane gotov modul:
'
' 1) modConfig.bas — konstante (TBL_RADNI_NALOG/_STAVKA, COL_RN_*, COL_RNS_*,
'    RN_TIP_*, RN_SMER_*, RN_IZLAZ_*, RN_STATUS_*) -> DODATO u ovom commit-u.
'
' 2) modSetup.bas — EnsureRadniNalogSchema (reuse EnsureDataTable) -> DODATO.
'    Pokrenuti JEDNOM (Alt+F8). Do tada read-modeli/TX fail-fast pucaju jasno.
'
' 3) modPaletniList.bas — Public CreatePaleta(...) (reuse CreateNewPaleta) ->
'    DODATO. Tako izlazne palete naloga koriste postojecu logiku palete.
'
' 4) modStorno.bas — StornoRadniNalog_TX / StornoRadniNalog -> DODATO:
'    vraca ulazne palete u lager (Preradjeno=""), stornira izlazne palete,
'    markira nalog + stavke. Po istoj kicmi kao StornoPaleta/Prerada.
'
' 5) frmRadniNalog (kao frmPalete; vidi docs/frmPalete-build.md): izbor ulaznih
'    paleta (multi-select), grid izlaza (frakcija/kolicina), izvrsilac/smena/vreme;
'    dugmad zovu SAMO read-modele + SaveRadniNalog_TX (+ StornoRadniNalog_TX).
'
' 6) Sledljivost izvestaj: GetOtkupiZaNalog(nalogID) -> izlaz->nalog->ulazne
'    palete -> reuse modPaletniList trace (promovisati GetOtkupiZaPalete u Public).
'    Ostavljeno kao TODO da se ne duplira trace logika.
' ============================================================
