Attribute VB_Name = "modHladnjaca"
Option Explicit

' ============================================================
' modHladnjaca — Lager / hladnjaca (stanje robe). SKICA (skeleton).
'
' Naslednik stuba "Phase 2: Kuhlhaus-Modul (Lager, Prerada, Monitoring,
' Sledljivost)". Zatvara petlju: otkup -> prijemnica -> paleta -> [LAGER]
' -> prerada -> [LAGER izlaz: kutije/kese] -> prodaja/faktura.
'
' ARHITEKTONSKA ODLUKA (zasto ovako):
'   Stanje lagera se RACUNA NA CITANJU (computed-on-read) iz tblPaleta +
'   tblPrerada. Te tabele SU vec knjiga (Storno/Preradjeno flegovi su revizija),
'   pa NE materijalizujemo tblLager u prvoj iteraciji:
'     + nula novih upisa u postojece TX tokove (SavePrijemnica_TX, SavePrerada_TX,
'       storno) -> najmanji bezbedan diff, nemoguca desinhronizacija;
'     + izvor istine ostaje jedan (palete/prerada), bez paralelne knjige koja
'       moze da odluta.
'   Jedina cinjenica koju palete NEMAJU je KOMORA (lokacija u hladnjaci) ->
'   opciona kolona COL_PAL_KOMORA + jedan TX setter (SetPaletaKomora_TX) po kicmi.
'   Materijalizovani movement-journal u tblLager (ako jednom zatreba audit po
'   pokretu) ostaje OZNACEN TODO hook u postojecim TX-ovima — NE zica se sada.
'
' Reuse (bez duplikacije): GetTableData / GetColumnIndex / FindRows /
'   RequireColumnIndex / RequireUpdateCell / GetNextID / clsTransaction / LogErr.
' Read-modeli (bez TX) vracaju 0-based 2D Variant spreman za ListBox.List ili
' Empty; stornirane i preradjene palete se NE racunaju u sirovi lager.
'
' FINALIZACIJA (vidi blok na dnu): preseliti COL_PAL_KOMORA u modConfig,
'   schema u modSetup.EnsureHladnjacaSchema, helpere u modHelpers, dodati
'   frmHladnjaca (kao frmPalete) i dugme u frmOtkupAPP.
' ============================================================

' SKICA-lokalno: u finalu preseliti u modConfig (pored COL_PAL_*).
Private Const COL_PAL_KOMORA As String = "Komora"

' Status koji se racunaju kao "fizicki u hladnjaci" (sirov lager).
' Otvorena i Zatvorena su obe na stanju; izlaze tek prelaskom u preradu
' (Preradjeno=Da) ili stornom.

' ============================================================
' READ-MODELI — rade ODMAH, bez ikakve izmene seme (citaju postojece tabele).
' ============================================================

' Sirovi lager: ne-stornirane i ne-preradjene palete, grupisane po
' Vrsta|Sorta|Klasa|Komora -> zbir paleta/gajbica/neto/bruto.
' Filteri: god=0 -> sve; vrsta/klasa "" -> sve.
' Kolone (0-based): 0 Vrsta,1 Sorta,2 Klasa,3 Komora,4 BrPaleta,5 Gajbice,
' 6 NetoKg,7 BrutoKg. Empty ako nema.
Public Function GetLagerSirovaForGrid(Optional ByVal god As Long = 0, _
                                      Optional ByVal vrsta As String = "", _
                                      Optional ByVal klasa As String = "") As Variant
    On Error GoTo EH
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Function

    Dim iVrsta As Long, iSorta As Long, iKlasa As Long, iKomora As Long
    Dim iStat As Long, iPre As Long, iStorno As Long, iGod As Long
    Dim iGajb As Long, iNeto As Long, iBruto As Long
    iVrsta = GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA)
    iSorta = GetColumnIndex(TBL_PALETA, COL_PAL_SORTA)
    iKlasa = GetColumnIndex(TBL_PALETA, COL_PAL_KLASA)
    iKomora = GetColumnIndex(TBL_PALETA, COL_PAL_KOMORA)   ' moze biti 0 (kolona jos ne postoji)
    iStat = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)
    iPre = GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO)
    iStorno = GetColumnIndex(TBL_PALETA, COL_STORNIRANO)
    iGod = GetColumnIndex(TBL_PALETA, COL_PAL_GODINA)
    iGajb = GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)
    iNeto = GetColumnIndex(TBL_PALETA, COL_PAL_NETO)
    iBruto = GetColumnIndex(TBL_PALETA, COL_PAL_BRUTO)

    ' agregacija po kompozitnom kljucu; redosled pojavljivanja ocuvan
    Dim order As Collection: Set order = New Collection
    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary") ' key -> Array(vr,so,kl,ko,brP,gaj,neto,bruto)

    Dim r As Long
    For r = 1 To UBound(d, 1)
        If UCase$(Trim$(CStr(SafeCell(d, r, iStorno)))) <> "DA" _
           And UCase$(Trim$(CStr(SafeCell(d, r, iPre)))) <> "DA" _
           And (god = 0 Or NzL(SafeCell(d, r, iGod)) = god) _
           And (vrsta = "" Or CStr(SafeCell(d, r, iVrsta)) = vrsta) _
           And (klasa = "" Or CStr(SafeCell(d, r, iKlasa)) = klasa) Then

            Dim vr As String, so As String, kl As String, ko As String
            vr = CStr(SafeCell(d, r, iVrsta))
            so = CStr(SafeCell(d, r, iSorta))
            kl = CStr(SafeCell(d, r, iKlasa))
            ko = CStr(SafeCell(d, r, iKomora))             ' "" ako kolone nema
            Dim key As String: key = vr & "|" & so & "|" & kl & "|" & ko

            If Not agg.Exists(key) Then
                order.Add key
                agg(key) = Array(vr, so, kl, ko, 0&, 0&, 0#, 0#)
            End If
            Dim a As Variant: a = agg(key)
            a(4) = a(4) + 1                                 ' BrPaleta
            a(5) = a(5) + NzL(SafeCell(d, r, iGajb))        ' Gajbice
            a(6) = a(6) + NzD(SafeCell(d, r, iNeto))        ' NetoKg
            a(7) = a(7) + NzD(SafeCell(d, r, iBruto))       ' BrutoKg
            agg(key) = a
        End If
    Next r
    If order.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To order.count - 1, 0 To 7)
    Dim k As Long, kk As Variant
    For k = 0 To order.count - 1
        Dim a2 As Variant: a2 = agg(order(k + 1))
        res(k, 0) = a2(0): res(k, 1) = a2(1): res(k, 2) = a2(2): res(k, 3) = a2(3)
        res(k, 4) = a2(4): res(k, 5) = a2(5): res(k, 6) = a2(6): res(k, 7) = a2(7)
    Next k

    GetLagerSirovaForGrid = res
    Exit Function
EH:
    LogErr "modHladnjaca.GetLagerSirovaForGrid"
End Function

' Prerada na lageru (izlaz): ne-stornirane prerade.
' Kolone (0-based): 0 Broj/God,1 Datum,2 NetoIzlazKg,3 Kutije,4 Kese,5 Napomena.
' Empty ako nema.
Public Function GetLagerPreradaForGrid(Optional ByVal god As Long = 0) As Variant
    On Error GoTo EH
    Dim d As Variant: d = GetTableData(TBL_PRERADA)
    If IsEmpty(d) Then Exit Function

    Dim iBroj As Long, iGod As Long, iDat As Long, iNeto As Long
    Dim iKut As Long, iKes As Long, iNap As Long, iStorno As Long
    iBroj = GetColumnIndex(TBL_PRERADA, COL_PRE_BROJ)
    iGod = GetColumnIndex(TBL_PRERADA, COL_PRE_GODINA)
    iDat = GetColumnIndex(TBL_PRERADA, COL_PRE_DATUM)
    iNeto = GetColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ)
    iKut = GetColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE)
    iKes = GetColumnIndex(TBL_PRERADA, COL_PRE_KESE)
    iNap = GetColumnIndex(TBL_PRERADA, COL_PRE_NAPOMENA)
    iStorno = GetColumnIndex(TBL_PRERADA, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If UCase$(Trim$(CStr(SafeCell(d, r, iStorno)))) <> "DA" _
           And (god = 0 Or NzL(SafeCell(d, r, iGod)) = god) Then
            rows.Add r
        End If
    Next r
    If rows.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To rows.count - 1, 0 To 5)
    Dim k As Long
    For k = 0 To rows.count - 1
        r = rows(k + 1)
        res(k, 0) = NzL(SafeCell(d, r, iBroj)) & "/" & NzL(SafeCell(d, r, iGod))
        res(k, 1) = Format$(SafeCell(d, r, iDat), "dd.mm.yyyy")
        res(k, 2) = NzD(SafeCell(d, r, iNeto))
        res(k, 3) = NzL(SafeCell(d, r, iKut))
        res(k, 4) = NzL(SafeCell(d, r, iKes))
        res(k, 5) = CStr(SafeCell(d, r, iNap))
    Next k

    GetLagerPreradaForGrid = res
    Exit Function
EH:
    LogErr "modHladnjaca.GetLagerPreradaForGrid"
End Function

' Saldo na jedan pogled (za labelu/statusnu traku). ByRef agregati za godinu
' (god=0 -> sve godine). Sve iz izvora istine, bez TX.
Public Sub GetLagerSummary(Optional ByVal god As Long = 0, _
                           Optional ByRef brOtvorenih As Long, _
                           Optional ByRef brZatvorenih As Long, _
                           Optional ByRef netoNaPaletama As Double, _
                           Optional ByRef brPrerada As Long, _
                           Optional ByRef kutijaUkupno As Long, _
                           Optional ByRef kesaUkupno As Long, _
                           Optional ByRef netoIzlazUkupno As Double)
    brOtvorenih = 0: brZatvorenih = 0: netoNaPaletama = 0
    brPrerada = 0: kutijaUkupno = 0: kesaUkupno = 0: netoIzlazUkupno = 0
    On Error GoTo EH

    ' --- sirovi lager (palete jos u hladnjaci) ---
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If Not IsEmpty(d) Then
        Dim iStat As Long, iPre As Long, iStorno As Long, iGod As Long, iNeto As Long
        iStat = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)
        iPre = GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO)
        iStorno = GetColumnIndex(TBL_PALETA, COL_STORNIRANO)
        iGod = GetColumnIndex(TBL_PALETA, COL_PAL_GODINA)
        iNeto = GetColumnIndex(TBL_PALETA, COL_PAL_NETO)
        Dim r As Long
        For r = 1 To UBound(d, 1)
            If UCase$(Trim$(CStr(SafeCell(d, r, iStorno)))) <> "DA" _
               And UCase$(Trim$(CStr(SafeCell(d, r, iPre)))) <> "DA" _
               And (god = 0 Or NzL(SafeCell(d, r, iGod)) = god) Then
                If CStr(SafeCell(d, r, iStat)) = PAL_STATUS_ZATVORENA Then
                    brZatvorenih = brZatvorenih + 1
                Else
                    brOtvorenih = brOtvorenih + 1
                End If
                netoNaPaletama = netoNaPaletama + NzD(SafeCell(d, r, iNeto))
            End If
        Next r
    End If

    ' --- prerada (izlaz) ---
    Dim p As Variant: p = GetTableData(TBL_PRERADA)
    If Not IsEmpty(p) Then
        Dim pGod As Long, pNeto As Long, pKut As Long, pKes As Long, pStorno As Long
        pGod = GetColumnIndex(TBL_PRERADA, COL_PRE_GODINA)
        pNeto = GetColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ)
        pKut = GetColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE)
        pKes = GetColumnIndex(TBL_PRERADA, COL_PRE_KESE)
        pStorno = GetColumnIndex(TBL_PRERADA, COL_STORNIRANO)
        Dim pr As Long
        For pr = 1 To UBound(p, 1)
            If UCase$(Trim$(CStr(SafeCell(p, pr, pStorno)))) <> "DA" _
               And (god = 0 Or NzL(SafeCell(p, pr, pGod)) = god) Then
                brPrerada = brPrerada + 1
                kutijaUkupno = kutijaUkupno + NzL(SafeCell(p, pr, pKut))
                kesaUkupno = kesaUkupno + NzL(SafeCell(p, pr, pKes))
                netoIzlazUkupno = netoIzlazUkupno + NzD(SafeCell(p, pr, pNeto))
            End If
        Next pr
    End If
    Exit Sub
EH:
    LogErr "modHladnjaca.GetLagerSummary"
End Sub

' Distinct komore (za ComboBox filter). Prazno ako kolona ne postoji.
Public Function GetKomoreList() As Variant
    On Error GoTo EH
    Dim iKom As Long: iKom = GetColumnIndex(TBL_PALETA, COL_PAL_KOMORA)
    If iKom = 0 Then GetKomoreList = Array(): Exit Function
    GetKomoreList = GetLookupList(TBL_PALETA, COL_PAL_KOMORA)
    Exit Function
EH:
    LogErr "modHladnjaca.GetKomoreList"
    GetKomoreList = Array()
End Function

' ============================================================
' TX — jedina nova izmena podataka: dodela/premestaj palete u komoru.
' Pun clsTransaction wrapper po kicmi (snapshot -> validacija -> upis ->
' commit; EH: rollback + re-raise). UI hvata gresku.
' ============================================================
Public Function SetPaletaKomora_TX(ByVal palID As String, _
                                   ByVal komora As String) As String
    Const SRC As String = "modHladnjaca.SetPaletaKomora_TX"

    Dim tx As clsTransaction
    On Error GoTo EH

    If Trim$(palID) = "" Then
        Err.Raise vbObjectError + 7360, SRC, "PaletaID je prazan."
    End If

    ' fail-fast ako schema jos nije migrirana (kolona Komora ne postoji)
    If GetColumnIndex(TBL_PALETA, COL_PAL_KOMORA) = 0 Then
        Err.Raise vbObjectError + 7361, SRC, _
                  "Kolona '" & COL_PAL_KOMORA & "' ne postoji. Pokreni " & _
                  "EnsureHladnjacaSchema (modSetup) jednom."
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA

    ' exact-row (puca na 0 i na >1 — nikad 'prvi od duplikata')
    Dim hits As Collection: Set hits = FindRows(TBL_PALETA, COL_PAL_ID, palID)
    If hits.count = 0 Then
        Err.Raise vbObjectError + 7362, SRC, "Nema palete: " & palID
    ElseIf hits.count > 1 Then
        Err.Raise vbObjectError + 7363, SRC, _
                  "Vise redova (" & hits.count & ") za paletu: " & palID
    End If
    Dim rIdx As Long: rIdx = CLng(hits(1))

    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If UCase$(Trim$(CStr(SafeCell(d, rIdx, GetColumnIndex(TBL_PALETA, COL_STORNIRANO))))) = "DA" Then
        Err.Raise vbObjectError + 7364, SRC, "Paleta je stornirana."
    End If

    RequireUpdateCell TBL_PALETA, rIdx, COL_PAL_KOMORA, komora, SRC

    tx.CommitTx
    SetPaletaKomora_TX = palID
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
' PRIVATE helperi (SKICA): identicni su privatnima u modPaletniList.
' TODO(dedup): promovisati NzL/NzD/SafeCell u modHelpers (Public) i obrisati
' lokalne kopije ovde i u modPaletniList. Vidi review nalaz #1.
' ============================================================
Private Function NzD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then NzD = CDbl(v)
End Function

Private Function NzL(ByVal v As Variant) As Long
    On Error Resume Next
    If IsNumeric(v) Then NzL = CLng(v)
End Function

' Bezbedno citanje celije iz GetTableData niza: ako kolona ne postoji
' (idx < 1, schema drift), vrati Empty umesto subscript-error.
Private Function SafeCell(ByVal d As Variant, ByVal r As Long, _
                          ByVal idx As Long) As Variant
    If idx >= 1 Then SafeCell = d(r, idx) Else SafeCell = Empty
End Function

' ============================================================
' FINALIZACIJA — koraci da skica postane gotov modul (ne diraju ovaj fajl):
'
' 1) modConfig.bas — preseliti konstantu (i dodati sifarnik komora ako se hoce):
'      Public Const COL_PAL_KOMORA As String = "Komora"
'      Public Const TBL_KOMORA     As String = "tblKomora"   ' opciono
'      Public Const COL_KOM_OZNAKA As String = "Oznaka"
'      Public Const COL_KOM_OPIS   As String = "Opis"
'    (i obrisati Private Const COL_PAL_KOMORA s vrha ovog modula)
'
' 2) modSetup.bas — dodati uz EnsurePaletniListSchema (koristi njegove Private
'    EnsureColumnOnTable / EnsureDataTable, bez duplikacije):
'      Public Sub EnsureHladnjacaSchema()
'          On Error GoTo EH
'          EnsureColumnOnTable TBL_PALETA, COL_PAL_KOMORA
'          ' opciono sifarnik komora:
'          ' EnsureDataTable TBL_KOMORA, "Komore", Array(COL_KOM_OZNAKA, COL_KOM_OPIS)
'          LogSetup "OK", "EnsureHladnjacaSchema done"
'          Exit Sub
'      EH:
'          LogSetup "ERROR", "EnsureHladnjacaSchema failed: " & Err.description
'      End Sub
'    Pokrenuti JEDNOM (Alt+F8) na master workbook-u. SetPaletaKomora_TX do tada
'    fail-fast puca s jasnom porukom (po dizajnu).
'
' 3) frmHladnjaca (kao frmPalete, vidi docs/frmPalete-build.md): lstSirovo +
'    lstPrerada + filter (Godina/Vrsta/Klasa/Komora) + summary labela; dugmad
'    zovu SAMO read-modele gore i SetPaletaKomora_TX. Embed preko
'    OpenContentForm iz frmOtkupAPP (btnHladnjaca u SetupButtons).
'
' 4) (BUDUCE, samo ako zatreba audit po pokretu) materijalizovan tblLager kao
'    append-only movement-journal. Hook tacke (post tek tada, UNUTAR tih TX-ova,
'    pre CommitTx, uz tx.AddTableSnapshot TBL_LAGER):
'       - modDokumenta.PaletizePrijemnica  -> +ULAZ (kg/gajbice po vrsti/klasi)
'       - modPaletniList.SavePrerada_TX    -> -IZLAZ (palete) +IZLAZ (kutije/kese)
'       - modStorno.StornoPaleta/Prerada   -> reverzni pokret
'    Dok je stanje computed-on-read (sada), ovo NIJE potrebno i ne sme se zicati.
' ============================================================
