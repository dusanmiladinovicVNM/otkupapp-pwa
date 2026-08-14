Attribute VB_Name = "modPaletniList"
'Attribute VB_Name = "modPaletniList"
Option Explicit

' ============================================================
' modPaletniList -- paletni list sveze robe + prerada (Phase 2)
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
' Sema tabela: modSetup.SetupPaletniListSchema (pokrenuti jednom).
' ============================================================

' Ispravka stornirane prijemnice: kad se ista roba samo prevezuje na ispravljenu
' prijemnicu (ReassignPaleteToPrijemnica_TX re-point originalnih paleta), sveza
' paletizacija u SavePrijemnica*/PaletizePrijemnica se PRESKACE -- inace bi se
' kreirale palete koje se odmah storniraju (prazna otvorena paleta + potrosen broj).
' Forme postavljaju True pre snimanja ispravke i False odmah posle. Default False.
Private mSkipPaletize As Boolean

' Povratne vrednosti AdjustPaletaGajbiceZaPrijemnicu_TX (pored >=0 = broj korigovanih klasa).
Public Const ADJ_NEEDS_CHOICE As Long = -1   ' visak ne staje -> UI pita: prelij ili preko kapaciteta
Public Const ADJ_BLOCKED As Long = -2        ' blokirano (npr. preradjena paleta) -> outInfo nosi razlog

Public Sub SetPaletizeSkip(ByVal b As Boolean)
    mSkipPaletize = b
End Sub

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
            If CLng(val(CStr(data(r, iGod)))) = yr Then
                n = CLng(val(CStr(data(r, iBroj))))
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
' PUBLIC -- paletizacija iz modDokumenta.SavePrijemnica_TX / Multi_TX.
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

    ' Toggle: paletiranje iskljuceno (Podesavanja) -> bez paleta/paletnih listova.
    ' Prijemnica se i dalje snima normalno; samo izostaje paletizacija. Default ON.
    If Not IsPaletiranjeEnabled() Then Exit Function

    ' Ispravka: preskoci svezu paletizaciju -- palete se prevezuju re-pointom
    ' (ReassignPaleteToPrijemnica_TX) da se ista roba ne paletizuje pa odmah stornira.
    If mSkipPaletize Then Exit Function

    If brGajbica <= 0 Then Exit Function       ' nema gajbica (Klasa II / bez ambalaze)

    RequirePaletaSchema SRC
    RequirePaletaStavkaSchema SRC
    RequirePrijemnicaNotPaletized prijemnicaID, SRC

    Dim crateW As Double: crateW = GetTezinaGajbice(tipAmb)

    Dim touched As Object: Set touched = CreateObject("Scripting.Dictionary")
    Dim nClosed As Long

    ' Raspodela = zajednicka petlja SpillGajbice (ista i za korekciju/Adjust).
    SpillGajbice prijemnicaID, brojPrij, brojZbirne, klasa, vrstaVoca, sortaVoca, _
                 tipAmb, brGajbica, netoKg / brGajbica, crateW, touched, SRC, _
                 closedPalIDs, nClosed

    PaletizePrijemnica = "palete=" & touched.count & "; zatvoreno=" & nClosed & _
                         "; gajbica=" & brGajbica
End Function

' Idempotency guard: ista PrijemnicaID ne sme imati aktivnu (ne-storniranu)
' paletnu stavku -> sprecava dvostruku paletizaciju (retry/re-save).
'
' Require*, ne Ensure*: ovo NISTA ne menja -- samo cita i puca. Ime je ranije bilo
' EnsurePrijemnicaNotAlreadyPaletized, sto je pozivaocu obecavalo mutaciju koje
' nema. Require* je zatecena konvencija projekta za provere preduslova
' (RequireSingleRow, RequireColumnIndex, RequireBimSmer...).
Private Sub RequirePrijemnicaNotPaletized(ByVal prijemnicaID As String, _
                                          ByVal SRC As String)
    Dim d As Variant: d = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(d) Then Exit Sub

    Dim iPrij As Long, iStorno As Long
    iPrij = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID, SRC)
    iStorno = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, SRC)

    Dim r As Long
    For r = 1 To UBound(d, 1)
        If CStr(d(r, iPrij)) = prijemnicaID _
           And UCase$(Trim$(CStr(d(r, iStorno)))) <> "DA" Then
            Err.Raise vbObjectError + 7330, SRC, _
                      "Prijemnica " & prijemnicaID & " je ve" & ChrW(263) & " paletizovana."
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

' UI getter (post-commit): kratak status palete(a) za datu prijemnicu, za
' prikaz u frmDokumenta posle snimanja. Cita iskomitovane tabele, bez izmena.
' Vraca "" ako prijemnica nije paletizovana (npr. Klasa II / bez gajbica).
Public Function GetPaletaStatusForPrijemnica(ByVal prijemnicaID As String) As String
    On Error GoTo EH
    If Trim$(prijemnicaID) = "" Then Exit Function

    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function

    Dim iPrij As Long, iPalID As Long, iGajb As Long, iStorno As Long
    iPrij = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    iPalID = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
    iGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    iStorno = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)

    ' PaletaID -> dodato gajbica za ovu prijemnicu (redosled pojavljivanja)
    Dim order As Collection: Set order = New Collection
    Dim added As Object: Set added = CreateObject("Scripting.Dictionary")
    Dim r As Long, pid As String
    For r = 1 To UBound(s, 1)
        If CStr(SafeCell(s, r, iPrij)) = prijemnicaID _
           And UCase$(Trim$(CStr(SafeCell(s, r, iStorno)))) <> "DA" Then
            pid = CStr(SafeCell(s, r, iPalID))
            If Not added.Exists(pid) Then added.Add pid, 0&: order.Add pid
            added(pid) = added(pid) + NzL(SafeCell(s, r, iGajb))
        End If
    Next r
    If order.count = 0 Then Exit Function

    Dim dp As Variant: dp = GetTableData(TBL_PALETA)
    Dim out As String, v As Variant
    For Each v In order
        Dim ri As Long: ri = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, CStr(v))
        If ri > 0 Then
            Dim brj As String, god As String, used As Long, cap As Long, st As String
            brj = CStr(NzL(SafeCell(dp, ri, GetColumnIndex(TBL_PALETA, COL_PAL_BROJ))))
            god = CStr(NzL(SafeCell(dp, ri, GetColumnIndex(TBL_PALETA, COL_PAL_GODINA))))
            used = NzL(SafeCell(dp, ri, GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)))
            cap = NzL(SafeCell(dp, ri, GetColumnIndex(TBL_PALETA, COL_PAL_KAPACITET)))
            st = LCase$(CStr(SafeCell(dp, ri, GetColumnIndex(TBL_PALETA, COL_PAL_STATUS))))
            If out <> "" Then out = out & vbCrLf
            out = out & "Paleta " & brj & "/" & god & ": +" & added(CStr(v)) & _
                  " gajb., stanje " & used & "/" & cap & ", " & st
        End If
    Next v

    GetPaletaStatusForPrijemnica = out
    Exit Function
EH:
    LogErr "modPaletniList.GetPaletaStatusForPrijemnica"
    GetPaletaStatusForPrijemnica = ""
End Function

' Kompaktan spisak fizickih paleta koje nose AKTIVNE paleta-stavke date prijemnice
' (po BrojPrijemnice): npr. "12/2026 (10g), 14/2026 (6g)". "" ako nema. Read-only
' (bez TX). Storno-agnosticno na strani prijemnice: gleda samo aktivne stavke, pa
' radi i posle storna (osirocene stavke i dalje nose stari BrojPrijemnice).
' Za: preview u recovery panelu (izbor leve prijemnice) + upozorenje pri stornu.
Public Function GetPaleteInfoForPrijemnicaBroj(ByVal brojPrij As String) As String
    On Error GoTo EH
    brojPrij = Trim$(brojPrij)
    If Len(brojPrij) = 0 Then Exit Function

    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function

    Dim iBr As Long, iPal As Long, iGajb As Long, iSt As Long
    iBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    iPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
    iGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    iSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If iBr = 0 Or iPal = 0 Then Exit Function

    ' PaletaID -> gajbice (agregat), uz redosled pojavljivanja.
    Dim order As Collection: Set order = New Collection
    Dim gaj As Object: Set gaj = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To UBound(s, 1)
        If Trim$(CStr(SafeCell(s, r, iBr))) = brojPrij _
           And UCase$(Trim$(CStr(SafeCell(s, r, iSt)))) <> "DA" Then
            Dim pid As String: pid = CStr(SafeCell(s, r, iPal))
            If Not gaj.Exists(pid) Then gaj.Add pid, 0&: order.Add pid
            gaj(pid) = CLng(gaj(pid)) + NzL(SafeCell(s, r, iGajb))
        End If
    Next r
    If order.count = 0 Then Exit Function

    Const MAXP As Long = 8
    Dim out As String, n As Long
    Dim v As Variant
    For Each v In order
        n = n + 1
        If n > MAXP Then
            out = out & ", +" & (order.count - MAXP) & " jos"
            Exit For
        End If
        If Len(out) > 0 Then out = out & ", "
        out = out & PaletaLabel(CStr(v)) & " (" & CLng(gaj(CStr(v))) & "g)"
    Next v

    GetPaleteInfoForPrijemnicaBroj = out
    Exit Function
EH:
    LogErr "modPaletniList.GetPaleteInfoForPrijemnicaBroj"
    GetPaleteInfoForPrijemnicaBroj = ""
End Function

' ============================================================
' STORNO CENTAR - uvid u palete (READ-ONLY, ne mutira). Za dati kljuc (fieldCol =
' COL_PALS_BROJ_PRIJ ili COL_PALS_BROJ_ZBIRNE) vrati po jednu stavku dicta po paleti:
'   paletaID, label, used, cap, neto, amb, preradjena (agregati palete)
'   thisGajb, thisNeto, thisAmb (suma AKTIVNIH stavki za ovaj kljuc = detach delta).
' Reuse GetPaletaAggregates/PaletaLabel/IsPaletaPreradjena (isti modul).
' ============================================================
Public Function GetPaleteImpactByField(ByVal fieldCol As String, ByVal value As String) As Collection
    Const SRC As String = "modPaletniList.GetPaleteImpactByField"
    Dim result As New Collection
    Set GetPaleteImpactByField = result
    On Error GoTo EH
    value = Trim$(value)
    If Len(value) = 0 Then Exit Function

    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function
    Dim cKey As Long, cPal As Long, cGa As Long, cNe As Long, cAm As Long, cSt As Long
    cKey = GetColumnIndex(TBL_PALETA_STAVKA, fieldCol)
    cPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
    cGa = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    cNe = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO)
    cAm = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_AMBALAZA)
    cSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    If cKey = 0 Or cPal = 0 Then Exit Function

    Dim order As Collection: Set order = New Collection
    Dim thG As Object: Set thG = CreateObject("Scripting.Dictionary")
    Dim thN As Object: Set thN = CreateObject("Scripting.Dictionary")
    Dim thA As Object: Set thA = CreateObject("Scripting.Dictionary")
    Dim r As Long, pid As String
    For r = 1 To UBound(s, 1)
        If Trim$(CStr(SafeCell(s, r, cKey))) = value _
           And UCase$(Trim$(CStr(SafeCell(s, r, cSt)))) <> "DA" Then
            pid = Trim$(CStr(SafeCell(s, r, cPal)))
            If Len(pid) > 0 Then
                If Not thG.Exists(pid) Then thG.Add pid, 0&: thN.Add pid, 0#: thA.Add pid, 0#: order.Add pid
                thG(pid) = NzL(thG(pid)) + NzL(SafeCell(s, r, cGa))
                thN(pid) = NzD(thN(pid)) + NzD(SafeCell(s, r, cNe))
                thA(pid) = NzD(thA(pid)) + NzD(SafeCell(s, r, cAm))
            End If
        End If
    Next r

    Dim v As Variant
    For Each v In order
        pid = CStr(v)
        Dim palRow As Long: palRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, pid)
        Dim used As Long, pNeto As Double, pAmb As Double, palk As Double, cap As Long
        If palRow > 0 Then GetPaletaAggregates palRow, used, pNeto, pAmb, palk, cap
        Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
        d("paletaID") = pid
        d("label") = PaletaLabel(pid)
        d("used") = used
        d("cap") = cap
        d("neto") = pNeto
        d("amb") = pAmb
        d("preradjena") = IsPaletaPreradjena(pid)
        d("thisGajb") = CLng(thG(pid))
        d("thisNeto") = CDbl(thN(pid))
        d("thisAmb") = CDbl(thA(pid))
        result.Add d
    Next v
    Exit Function
EH:
    LogErr SRC
End Function

' ============================================================
' frmPalete (#44) read-modeli + rucno zatvaranje. Read-modeli su SAMO za
' citanje (bez TX); vracaju 0-based 2D Variant (spreman za ListBox.List) ili
' Empty. Stornirane palete se ne prikazuju. Sva izmena ide preko TX wrappera.
' ============================================================

' Palete za grid. Filteri: god=0 -> sve; vrsta/status/preradjeno "" -> sve.
' status: "Otvorena"|"Zatvorena"|""; preradjeno: "Da"|"Ne"|"".
' Kolone (0-based): 0 PaletaID(skriveno),1 Broj,2 Godina,3 Vrsta,4 Sorta,
' 5 Klasa,6 TipAmb,7 Gajbice,8 Kapacitet,9 Neto,10 Bruto,11 Status,12 Preradjeno.
Public Function GetPaleteForGrid(Optional ByVal god As Long = 0, _
                                 Optional ByVal vrsta As String = "", _
                                 Optional ByVal status As String = "", _
                                 Optional ByVal preradjeno As String = "", _
                                 Optional ByVal sorta As String = "") As Variant
    On Error GoTo EH
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long, iBroj As Long, iGod As Long, iVrsta As Long, iSorta As Long
    Dim iKlasa As Long, iTipA As Long, iGajb As Long, iKap As Long, iNeto As Long
    Dim iBruto As Long, iStat As Long, iPre As Long, iStorno As Long
    iID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
    iBroj = GetColumnIndex(TBL_PALETA, COL_PAL_BROJ)
    iGod = GetColumnIndex(TBL_PALETA, COL_PAL_GODINA)
    iVrsta = GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA)
    iSorta = GetColumnIndex(TBL_PALETA, COL_PAL_SORTA)
    iKlasa = GetColumnIndex(TBL_PALETA, COL_PAL_KLASA)
    iTipA = GetColumnIndex(TBL_PALETA, COL_PAL_TIP_AMBALAZE)
    iGajb = GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)
    iKap = GetColumnIndex(TBL_PALETA, COL_PAL_KAPACITET)
    iNeto = GetColumnIndex(TBL_PALETA, COL_PAL_NETO)
    iBruto = GetColumnIndex(TBL_PALETA, COL_PAL_BRUTO)
    iStat = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)
    iPre = GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO)
    iStorno = GetColumnIndex(TBL_PALETA, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If UCase$(Trim$(CStr(SafeCell(d, r, iStorno)))) <> "DA" _
           And (god = 0 Or NzL(SafeCell(d, r, iGod)) = god) _
           And (vrsta = "" Or CStr(SafeCell(d, r, iVrsta)) = vrsta) _
           And (sorta = "" Or CStr(SafeCell(d, r, iSorta)) = sorta) _
           And (status = "" Or CStr(SafeCell(d, r, iStat)) = status) _
           And PreradMatch(CStr(SafeCell(d, r, iPre)), preradjeno) Then
            rows.Add r
        End If
    Next r
    If rows.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To rows.count - 1, 0 To 12)
    Dim k As Long
    For k = 0 To rows.count - 1
        r = rows(k + 1)
        res(k, 0) = CStr(SafeCell(d, r, iID))
        res(k, 1) = NzL(SafeCell(d, r, iBroj))
        res(k, 2) = NzL(SafeCell(d, r, iGod))
        res(k, 3) = CStr(SafeCell(d, r, iVrsta))
        res(k, 4) = CStr(SafeCell(d, r, iSorta))
        res(k, 5) = CStr(SafeCell(d, r, iKlasa))
        res(k, 6) = CStr(SafeCell(d, r, iTipA))
        res(k, 7) = NzL(SafeCell(d, r, iGajb))
        res(k, 8) = NzL(SafeCell(d, r, iKap))
        res(k, 9) = Format$(NzD(SafeCell(d, r, iNeto)), "0.00")
        res(k, 10) = Format$(NzD(SafeCell(d, r, iBruto)), "0.00")
        res(k, 11) = CStr(SafeCell(d, r, iStat))
        res(k, 12) = CStr(SafeCell(d, r, iPre))
    Next k

    GetPaleteForGrid = res
    Exit Function
EH:
    LogErr "modPaletniList.GetPaleteForGrid"
End Function

Private Function PreradMatch(ByVal CellVal As String, ByVal filter As String) As Boolean
    Select Case UCase$(Trim$(filter))
        Case "": PreradMatch = True
        Case "DA": PreradMatch = (UCase$(Trim$(CellVal)) = "DA")
        Case "NE": PreradMatch = (UCase$(Trim$(CellVal)) <> "DA")
        Case Else: PreradMatch = True
    End Select
End Function

' Prerade za desni pregled. god=0 -> sve. Stornirane se ne prikazuju.
' Kolone (0-based): 0 PreradaID(skriveno),1 Broj,2 Datum,3 Neto,4 Kutije,
' 5 Kese,6 TipGotovogProizvoda.
Public Function GetPreradeForGrid(Optional ByVal god As Long = 0) As Variant
    On Error GoTo EH
    Dim d As Variant: d = GetTableData(TBL_PRERADA)
    If IsEmpty(d) Then Exit Function

    Dim iID As Long, iBroj As Long, iGod As Long, iDat As Long
    Dim iNeto As Long, iKut As Long, iKes As Long, iGP As Long, iStorno As Long
    iID = GetColumnIndex(TBL_PRERADA, COL_PRE_ID)
    iBroj = GetColumnIndex(TBL_PRERADA, COL_PRE_BROJ)
    iGod = GetColumnIndex(TBL_PRERADA, COL_PRE_GODINA)
    iDat = GetColumnIndex(TBL_PRERADA, COL_PRE_DATUM)
    iNeto = GetColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ)
    iKut = GetColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE)
    iKes = GetColumnIndex(TBL_PRERADA, COL_PRE_KESE)
    iGP = GetColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP)
    iStorno = GetColumnIndex(TBL_PRERADA, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim r As Long
    For r = 1 To UBound(d, 1)
        If Trim$(CStr(SafeCell(d, r, iID))) <> "" _
           And UCase$(Trim$(CStr(SafeCell(d, r, iStorno)))) <> "DA" _
           And (god = 0 Or NzL(SafeCell(d, r, iGod)) = god) Then
            rows.Add r
        End If
    Next r
    If rows.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To rows.count - 1, 0 To 6)
    Dim k As Long
    For k = 0 To rows.count - 1
        r = rows(k + 1)
        res(k, 0) = CStr(SafeCell(d, r, iID))
        res(k, 1) = NzL(SafeCell(d, r, iBroj))
        res(k, 2) = ""
        If IsDate(SafeCell(d, r, iDat)) Then res(k, 2) = Format$(CDate(SafeCell(d, r, iDat)), "dd.mm.yyyy")
        res(k, 3) = Format$(NzD(SafeCell(d, r, iNeto)), "0.00")
        res(k, 4) = NzL(SafeCell(d, r, iKut))
        res(k, 5) = NzL(SafeCell(d, r, iKes))
        If iGP > 0 Then res(k, 6) = CStr(SafeCell(d, r, iGP))
    Next k
    GetPreradeForGrid = res
    Exit Function
EH:
    LogErr "modPaletniList.GetPreradeForGrid"
End Function

' Stavke izabrane palete za grid. Kolone (0-based): 0 PrijemnicaID,
' 1 BrojPrijemnice, 2 BrojZbirne, 3 Gajbice, 4 NetoKg. Empty ako nema.
Public Function GetPaletaStavkeForGrid(ByVal palID As String) As Variant
    On Error GoTo EH
    If Trim$(palID) = "" Then Exit Function
    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function

    Dim iPal As Long, iPrij As Long, iBrPrij As Long, iZbir As Long
    Dim iGajb As Long, iNeto As Long, iStorno As Long
    iPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
    iPrij = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    iBrPrij = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    iZbir = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
    iGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    iNeto = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO)
    iStorno = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)

    Dim rows As Collection: Set rows = New Collection
    Dim r As Long
    For r = 1 To UBound(s, 1)
        If CStr(SafeCell(s, r, iPal)) = palID _
           And UCase$(Trim$(CStr(SafeCell(s, r, iStorno)))) <> "DA" Then
            rows.Add r
        End If
    Next r
    If rows.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To rows.count - 1, 0 To 4)
    Dim k As Long
    For k = 0 To rows.count - 1
        r = rows(k + 1)
        res(k, 0) = CStr(SafeCell(s, r, iPrij))
        res(k, 1) = CStr(SafeCell(s, r, iBrPrij))
        res(k, 2) = CStr(SafeCell(s, r, iZbir))
        res(k, 3) = NzL(SafeCell(s, r, iGajb))
        res(k, 4) = Format$(NzD(SafeCell(s, r, iNeto)), "0.00")
    Next k

    GetPaletaStavkeForGrid = res
    Exit Function
EH:
    LogErr "modPaletniList.GetPaletaStavkeForGrid"
End Function

' Stavke za VISE izabranih paleta (agregat). Iste kolone kao
' GetPaletaStavkeForGrid: 0 PrijemnicaID, 1 BrojPrijemnice, 2 BrojZbirne,
' 3 Gajbice, 4 NetoKg. Empty ako nema.
Public Function GetPaletaStavkeForGridMulti(ByVal paletaIDs As Collection) As Variant
    On Error GoTo EH
    If paletaIDs Is Nothing Then Exit Function
    If paletaIDs.count = 0 Then Exit Function

    Dim want As Object: Set want = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    For Each v In paletaIDs
        want(CStr(v)) = True
    Next v

    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function

    Dim iPal As Long, iPrij As Long, iBrPrij As Long, iZbir As Long
    Dim iGajb As Long, iNeto As Long, iStorno As Long
    iPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
    iPrij = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID)
    iBrPrij = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    iZbir = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
    iGajb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    iNeto = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO)
    iStorno = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)

    ' grupisi po PrijemnicaID: ista prijemnica na vise izabranih paleta = 1 red,
    ' Gajbice/Neto = zbir porcija preko izabranih paleta.
    Dim order As Collection: Set order = New Collection
    Dim dBr As Object: Set dBr = CreateObject("Scripting.Dictionary")
    Dim dZb As Object: Set dZb = CreateObject("Scripting.Dictionary")
    Dim dGa As Object: Set dGa = CreateObject("Scripting.Dictionary")
    Dim dNe As Object: Set dNe = CreateObject("Scripting.Dictionary")

    Dim r As Long, prij As String
    For r = 1 To UBound(s, 1)
        If want.Exists(CStr(SafeCell(s, r, iPal))) _
           And UCase$(Trim$(CStr(SafeCell(s, r, iStorno)))) <> "DA" Then
            prij = CStr(SafeCell(s, r, iPrij))
            If Not dGa.Exists(prij) Then
                order.Add prij
                dBr(prij) = CStr(SafeCell(s, r, iBrPrij))
                dZb(prij) = CStr(SafeCell(s, r, iZbir))
                dGa(prij) = 0&
                dNe(prij) = 0#
            End If
            dGa(prij) = dGa(prij) + NzL(SafeCell(s, r, iGajb))
            dNe(prij) = dNe(prij) + NzD(SafeCell(s, r, iNeto))
        End If
    Next r
    If order.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(0 To order.count - 1, 0 To 4)
    Dim k As Long, p As String
    For k = 0 To order.count - 1
        p = order(k + 1)
        res(k, 0) = p
        res(k, 1) = dBr(p)
        res(k, 2) = dZb(p)
        res(k, 3) = dGa(p)
        res(k, 4) = Format$(dNe(p), "0.00")
    Next k

    GetPaletaStavkeForGridMulti = res
    Exit Function
EH:
    LogErr "modPaletniList.GetPaletaStavkeForGridMulti"
End Function

' Rucno zatvaranje otvorene palete (TX). Validira: postoji tacno jednom, nije
' stornirana/preradjena, jeste otvorena. Bez MsgBox -> baca gresku (UI hvata).
' Posle commita pokrece izlaz (po PALETA_PRINT_MODE). Vraca PaletaID.
Public Function ClosePaletaManual_TX(ByVal palID As String) As String
    Const SRC As String = "modPaletniList.ClosePaletaManual_TX"

    Dim tx As clsTransaction
    On Error GoTo EH

    If Trim$(palID) = "" Then
        Err.Raise vbObjectError + 7350, SRC, "PaletaID je prazan."
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA

    RequirePaletaSchema SRC
    Dim rIdx As Long
    rIdx = RequireSingleRowIndexByKey(TBL_PALETA, COL_PAL_ID, palID, SRC)

    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If UCase$(Trim$(CStr(SafeCell(d, rIdx, GetColumnIndex(TBL_PALETA, COL_STORNIRANO))))) = "DA" Then
        Err.Raise vbObjectError + 7351, SRC, "Paleta je stornirana."
    End If
    If UCase$(Trim$(CStr(SafeCell(d, rIdx, GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO))))) = "DA" Then
        Err.Raise vbObjectError + 7352, SRC, "Paleta je ve" & ChrW(263) & " preradjena."
    End If
    If CStr(SafeCell(d, rIdx, GetColumnIndex(TBL_PALETA, COL_PAL_STATUS))) <> PAL_STATUS_OTVORENA Then
        Err.Raise vbObjectError + 7353, SRC, "Paleta nije otvorena."
    End If

    RequireUpdateCell TBL_PALETA, rIdx, COL_PAL_STATUS, PAL_STATUS_ZATVORENA, SRC

    tx.CommitTx

    ' POST-commit izlaz (po modu) -> bez rollback rizika.
    Dim closed As Collection: Set closed = New Collection
    closed.Add palID
    PaletniListOutputClosed closed

    ClosePaletaManual_TX = palID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function

' ============================================================
' PUBLIC -- rucna stampa nepotpunih (otvorenih) paleta.
' Kraj smene: Alt+F8 -> PrintNepotpunePalete (kasnije dugme u UI).
' ============================================================
' Izlaz (po PALETA_PRINT_MODE) za sve otvorene (nepotpune) palete. Vraca broj
' obradjenih paleta; poruke su u UI sloju (modPaletniListUI / frmDokumenta).
Public Function PrintNepotpunePalete() As Long
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_PALETA)
    If IsEmpty(data) Then Exit Function

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

    PrintNepotpunePalete = cnt
    Exit Function

EH:
    LogErr "modPaletniList.PrintNepotpunePalete"
End Function

' ============================================================
' PUBLIC -- paletni list dokument preko PaletaSablon (isti pristup kao
' frmSledljivost.PrintTracePDF: Sablon + named-range fill + Export/Print).
' PaletaSablon se auto-kreira (EnsurePaletaSablon) i sme da se stilizuje --
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

' PDF jednog paletnog lista -> <workbook>\Paletni listovi\Paleta_<broj>-<god>.pdf. Vraca putanju.
Public Function ExportPaletniListPDF(ByVal palID As String, _
                                     Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim broj As String, god As String
    Dim ws As Worksheet
    Set ws = FillPaletaSablon(palID, broj, god)
    If ws Is Nothing Then Exit Function

    Dim pdfPath As String
    pdfPath = EnsureDocFolder(PDF_DIR_PALETNI) & "\Paleta_" & broj & "-" & god & ".pdf"

    DocExportPdf ws, pdfPath, openAfter

    ExportPaletniListPDF = pdfPath
    Exit Function
EH:
    LogErr "modPaletniList.ExportPaletniListPDF"
End Function

' Auto-izlaz pri zatvaranju palete, po config-u PALETA_PRINT_MODE:
'   (prazno/PDF) -> tihi PDF | PRINT -> stampac | PREVIEW -> pregled | OFF -> nista
Private Sub OutputPaletniListByMode(ByVal palID As String)
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_PALETA_PRINT_MODE), "OFF")

    Select Case mode
        Case "PRINT"
            PrintPaletniList palID
        Case "PREVIEW"
            Dim broj As String, god As String
            Dim ws As Worksheet
            Set ws = FillPaletaSablon(palID, broj, god)
            If Not ws Is Nothing Then DocPrintWs ws, mode
        Case "PDF"
            ExportPaletniListPDF palID, True    ' PDF + otvori (kao otkupni list)
        ' OFF / prazno (DEFAULT) -> bez izlaza; auto-izlaz pune palete rucno:
        '   SetConfigValue "PALETA_PRINT_MODE", "PDF" | "PRINT" | "PREVIEW"
    End Select
End Sub

' Popuni PaletaSablon (header + tezine + stavke). Vraca sheet + broj/god (ByRef).
Private Function FillPaletaSablon(ByVal palID As String, _
                                  ByRef brojOut As String, _
                                  ByRef godOut As String) As Worksheet
    Const SRC As String = "modPaletniList.FillPaletaSablon"
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsurePaletaSablon

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_PALETA_SABLON)
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

    On Error GoTo EH
    Application.ScreenUpdating = False

    ws.Range("PalBroj").NumberFormat = "@"
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
        ws.Range(ws.cells(startRow, 4), ws.cells(lastRow, 5)).UnMerge
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, 5)).Clear
    End If

    ' --- stavke: jedan red po OTKUPU (sifra kooperanta, neto, ambalaza) ---
    Dim ids As Collection: Set ids = New Collection
    ids.Add palID
    Dim o As Variant: o = GetOtkupiZaPalete(ids)

    Dim outR As Long, rb As Long
    outR = startRow: rb = 0
    If Not IsEmpty(o) Then
        Dim k As Long
        For k = 1 To UBound(o, 1)
            rb = rb + 1
            ws.cells(outR, 1).value = rb
            ws.cells(outR, 2).value = CStr(o(k, 1))                   ' Kooperant (sifra)
            ws.cells(outR, 3).value = o(k, 3)                         ' Neto kg
            ws.Range(ws.cells(outR, 4), ws.cells(outR, 5)).Merge
            ws.cells(outR, 4).value = o(k, 4) & " x " & CStr(o(k, 5)) ' Ambalaza: kom x tip (D:E)
            outR = outR + 1
        Next k
    End If

    ' --- stilizacija stavki (okviri + naizmenicne boje) ---
    Dim dataEnd As Long
    dataEnd = outR - 1
    If dataEnd >= startRow Then
        With ws.Range(ws.cells(startRow, 1), ws.cells(dataEnd, 5)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        ws.Range(ws.cells(startRow, 3), ws.cells(dataEnd, 3)).NumberFormat = "#,##0.00"
        ws.Range(ws.cells(startRow, 4), ws.cells(dataEnd, 4)).NumberFormat = "@"

        Dim zr As Long
        For zr = 0 To dataEnd - startRow
            If zr Mod 2 = 1 Then
                ws.Range(ws.cells(startRow + zr, 1), _
                         ws.cells(startRow + zr, 5)).Interior.Color = RGB(217, 225, 242)
            End If
        Next zr
    End If

    ' --- footer (kao sledljivost): datum stampe + potpis/pecat ---
    Dim footRow As Long
    footRow = dataEnd + 2
    If footRow <= startRow Then footRow = startRow + 1
    ws.cells(footRow, 1).value = "Datum stampe: " & Format$(Date, "dd.mm.yyyy")
    ws.cells(footRow, 1).Font.Color = DocColGray()
    ws.cells(footRow + 2, 1).value = "Potpis: ____________________"
    ws.cells(footRow + 2, 1).Font.Color = DocColGray()
    ws.cells(footRow + 2, 4).value = "Pecat: ____________________"
    ws.cells(footRow + 2, 4).Font.Color = DocColGray()

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
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(footRow + 2, 5)).Address
    End With
    Application.PrintCommunication = True
    On Error GoTo 0

    Application.ScreenUpdating = oldScreen
    Set FillPaletaSablon = ws
    Exit Function

EH:
    Application.ScreenUpdating = oldScreen
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

' Kreira/obnavlja PaletaSablon u zajednickom stilu (logo, naslov, polja, sazetak).
' Verzija layouta je u H1; na promenu verzije sheet se ponovo izgradi.
Public Sub EnsurePaletaSablon()
    On Error GoTo EH
    Const LAYOUT_VER As String = "3"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_PALETA_SABLON)
    On Error GoTo EH
    If Not ws Is Nothing Then
        If CStr(ws.Range("H1").value) = LAYOUT_VER Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = WS_PALETA_SABLON
    ws.cells.Font.name = "Calibri"
    ws.cells.Font.Size = 10
    ws.columns("A").ColumnWidth = 12
    ws.columns("B").ColumnWidth = 14
    ws.columns("C").ColumnWidth = 14
    ws.columns("D").ColumnWidth = 14
    ws.columns("E").ColumnWidth = 18

    Dim r As Long
    r = DocSellerHeader(ws, 1, 5, 5)
    r = DocTitleBlock(ws, r, 5, "Skladisno poslovanje - formiranje palete", "PALETNI LIST")

    Dim fr As Long: fr = r + 1
    ws.cells(fr, 1).value = "Broj:"
    ws.cells(fr + 1, 1).value = "Datum:"
    ws.cells(fr + 2, 1).value = "Tip palete:"
    ws.cells(fr + 3, 1).value = "Status:"
    ws.cells(fr, 4).value = "Broj gajbica:"
    ws.cells(fr + 1, 4).value = "Neto (kg):"
    ws.cells(fr + 2, 4).value = "Ambala" & ChrW(382) & "a (kg):"
    ws.cells(fr + 3, 4).value = "Paleta (kg):"
    ws.cells(fr + 4, 4).value = "BRUTO (kg):"

    ws.cells(fr, 2).name = "PalBroj"
    ws.cells(fr + 1, 2).name = "PalDatum"
    ws.cells(fr + 2, 2).name = "PalTip"
    ws.cells(fr + 3, 2).name = "PalStatus"
    ws.cells(fr, 5).name = "PalGajbica"
    ws.cells(fr + 1, 5).name = "PalNeto"
    ws.cells(fr + 2, 5).name = "PalAmbalaza"
    ws.cells(fr + 3, 5).name = "PalPaleta"
    ws.cells(fr + 4, 5).name = "PalBruto"

    ws.Range(ws.cells(fr, 2), ws.cells(fr + 3, 2)).Font.Bold = True
    ws.Range(ws.cells(fr, 5), ws.cells(fr + 4, 5)).Font.Bold = True
    ws.cells(fr, 5).NumberFormat = "0"
    ws.Range(ws.cells(fr + 1, 5), ws.cells(fr + 4, 5)).NumberFormat = "#,##0.00"

    ' desni sazetak: uokviri + istakni BRUTO
    ws.Range(ws.cells(fr, 4), ws.cells(fr + 4, 5)).BorderAround Weight:=xlThin
    With ws.Range(ws.cells(fr + 4, 4), ws.cells(fr + 4, 5))
        .Interior.Color = DocColHeaderFill()
        .Font.Bold = True
    End With
    With ws.Range(ws.cells(fr + 4, 4), ws.cells(fr + 4, 5)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' vrsta voca kao podnaslov iznad tabele (na paleti je uvek ista vrsta)
    Dim subRow As Long: subRow = fr + 6
    ws.cells(subRow, 1).value = "Vrsta vo" & ChrW(263) & "a:"
    ws.cells(subRow, 1).Font.Color = DocColGray()
    ws.Range(ws.cells(subRow, 2), ws.cells(subRow, 5)).Merge
    ws.cells(subRow, 2).name = "PalVrsta"
    With ws.cells(subRow, 2)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlLeft
    End With
    ws.rows(subRow).RowHeight = 20

    ' tabela stavki (bez kolone Vrsta; Ambalaza preko D:E)
    Dim hdr As Long: hdr = subRow + 1
    ws.cells(hdr, 1).value = "Rb"
    ws.cells(hdr, 2).value = "Kooperant"
    ws.cells(hdr, 3).value = "Neto kg"
    ws.Range(ws.cells(hdr, 4), ws.cells(hdr, 5)).Merge
    ws.cells(hdr, 4).value = "Ambala" & ChrW(382) & "a"
    With ws.Range(ws.cells(hdr, 1), ws.cells(hdr, 5))
        .Font.Bold = True
        .Interior.Color = DocColHeaderFill()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ws.cells(hdr + 1, 1).name = "PalStavkaStart"

    ws.Range(ws.cells(1, 1), ws.cells(hdr, 5)).EntireRow.AutoFit
    ws.Range("H1").value = LAYOUT_VER
    ws.Range("H1").Font.Color = RGB(255, 255, 255)
    Exit Sub

EH:
    Application.DisplayAlerts = True
    LogErr "modPaletniList.EnsurePaletaSablon"
End Sub

' Nadji PaletaID po broju palete + godini (za Prompt).
Public Function FindPaletaIDByBroj(ByVal broj As Long, ByVal god As Long) As String
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
' PRIVATE -- paleta lifecycle + lookup + util
' ============================================================

' ============================================================
' PUBLIC - info za frmOtkup: koliko gajbi jos treba da se zatvori AKTIVNA
' (otvorena) paleta za dati proizvod/klasu. Read-only: NE kreira paletu.
' Vraca "" ako je paletiranje iskljuceno ili nema otvorene palete.
' tipAmb (opciono) suzava na konkretan tip ambalaze; prazno = bilo koji.
Public Function GajbeDoZatvaranjaPaleteInfo(ByVal vrstaVoca As String, _
                                            ByVal sortaVoca As String, _
                                            ByVal klasa As String, _
                                            Optional ByVal tipAmb As String = "") As String
    On Error GoTo EH
    If Not IsPaletiranjeEnabled() Then Exit Function
    If Len(Trim$(vrstaVoca)) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_PALETA)
    If IsEmpty(data) Then Exit Function

    Dim iVrsta As Long, iSorta As Long, iKlasa As Long, iTipAmb As Long
    Dim iStatus As Long, iStorno As Long, iPre As Long
    Dim iGajb As Long, iKap As Long, iBroj As Long, iGod As Long
    iVrsta = GetColumnIndex(TBL_PALETA, COL_PAL_VRSTA)
    iSorta = GetColumnIndex(TBL_PALETA, COL_PAL_SORTA)
    iKlasa = GetColumnIndex(TBL_PALETA, COL_PAL_KLASA)
    iTipAmb = GetColumnIndex(TBL_PALETA, COL_PAL_TIP_AMBALAZE)
    iStatus = GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)
    iStorno = GetColumnIndex(TBL_PALETA, COL_STORNIRANO)
    iPre = GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO)
    iGajb = GetColumnIndex(TBL_PALETA, COL_PAL_BR_GAJBICA)
    iKap = GetColumnIndex(TBL_PALETA, COL_PAL_KAPACITET)
    iBroj = GetColumnIndex(TBL_PALETA, COL_PAL_BROJ)
    iGod = GetColumnIndex(TBL_PALETA, COL_PAL_GODINA)

    Dim r As Long
    For r = 1 To UBound(data, 1)
        If CStr(data(r, iVrsta)) = vrstaVoca _
           And CStr(SafeCell(data, r, iSorta)) = sortaVoca _
           And CStr(SafeCell(data, r, iKlasa)) = klasa _
           And CStr(data(r, iStatus)) = PAL_STATUS_OTVORENA _
           And UCase$(CStr(data(r, iStorno))) <> "DA" _
           And UCase$(Trim$(CStr(SafeCell(data, r, iPre)))) <> "DA" Then

            If Len(Trim$(tipAmb)) = 0 Or CStr(SafeCell(data, r, iTipAmb)) = tipAmb Then
                Dim used As Long, cap As Long, ostatak As Long
                used = NzL(SafeCell(data, r, iGajb))
                cap = NzL(SafeCell(data, r, iKap))
                If cap <= 0 Then cap = GetKapacitetPalete(vrstaVoca)
                ostatak = cap - used
                If ostatak < 0 Then ostatak = 0

                GajbeDoZatvaranjaPaleteInfo = "Paleta br. " & _
                    CStr(NzL(SafeCell(data, r, iBroj))) & "/" & _
                    CStr(NzL(SafeCell(data, r, iGod))) & ": jo" & ChrW(353) & " " & _
                    CStr(ostatak) & " gajbi do zatvaranja (" & _
                    CStr(used) & "/" & CStr(cap) & ")"
                Exit Function
            End If
        End If
    Next r
    Exit Function
EH:
    LogErr "modPaletniList.GajbeDoZatvaranjaPaleteInfo"
End Function

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

Private Sub ClosePaleta(ByVal palRow As Long, ByVal SRC As String)
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_STATUS, PAL_STATUS_ZATVORENA, SRC
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

' Tezina jedne gajbice (kg) iz sifarnika tblTipAmbalaze. Public: jedinstveni izvor
' tare za paletni list, otkupni list i bruto->neto konverziju u frmOtkup/modOtkupBlok.
Public Function GetTezinaGajbice(ByVal tipAmb As String) As Double
    GetTezinaGajbice = NzD(LookupValue(TBL_TIP_AMBALAZE, COL_TAMB_TIP, tipAmb, COL_TAMB_TEZINA))
End Function

Private Function GetTezinaPalete(ByVal tip As String) As Double
    GetTezinaPalete = NzD(LookupValue(TBL_TIP_PALETE, COL_TPAL_TIP, tip, COL_TPAL_TEZINA))
End Function

' ============================================================
' P1: re-point paleta-stavki sa STORNIRANE prijemnice na NOVU (reenter), za
' autohladnjaca teardown+ponovni unos. Cuva FIZICKE palete (zatvorene/zapecacene),
' menja samo dokument-vlasnika slice-a. Koraci:
'   1) UNDO sveze paletizacije nove prijemnice (storniraj njene stavke + skini sa
'      totala palete; reopen ako padne ispod kapaciteta; prazne palete ostaju).
'   2) RE-POINT osirocenih stavki stare prijemnice -> nova, PO KLASI (stara Kl.X
'      -> PrijemnicaID nove Kl.X; BrojPrijemnice/BrojZbirne nove).
'   3) KG-SYNC: ako se broj gajbica PO KLASI poklapa, skaliraj neto stavki na neto
'      nove prijemnice (+ total palete). Ako se broj gajbica razlikuje -> ne diraj
'      kolicine, vrati outGajbDiff=True (UI zatim zove AdjustPaletaGajbiceZaPrijemnicu_TX).
'   4) RELABEL (STEP 2b, uz allowRelabel): razlika vrsta/sorta/tipAmb -> prepravi
'      etiketu na prevezanim stavkama + njihovim paletama (roba se ne pomera).
' Sve u jednoj transakciji. outWarn nosi poruke za UI (prazno = bez napomena).
' ============================================================
Public Function ReassignPaleteToPrijemnica_TX(ByVal oldBroj As String, _
                                              ByVal newBroj As String, _
                                              Optional ByRef outWarn As String, _
                                              Optional ByVal allowRelabel As Boolean = False, _
                                              Optional ByRef outGajbDiff As Boolean = False) As Boolean
    Const SRC As String = "modPaletniList.ReassignPaleteToPrijemnica_TX"
    Dim tx As clsTransaction
    On Error GoTo EH

    outWarn = ""
    outGajbDiff = False
    oldBroj = Trim$(oldBroj): newBroj = Trim$(newBroj)
    If Len(oldBroj) = 0 Or Len(newBroj) = 0 Then Exit Function
    If StrComp(oldBroj, newBroj, vbTextCompare) = 0 Then
        outWarn = "Stara i nova prijemnica su isti broj."
        Exit Function
    End If

    ' --- Nova prijemnica po klasi: PrijemnicaID, neto(Kolicina), gajbica(KolAmb) + BrojZbirne. ---
    Dim newById As Object: Set newById = CreateObject("Scripting.Dictionary")
    Dim newNeto As Object: Set newNeto = CreateObject("Scripting.Dictionary")
    Dim newGajb As Object: Set newGajb = CreateObject("Scripting.Dictionary")
    newById.CompareMode = vbTextCompare: newNeto.CompareMode = vbTextCompare: newGajb.CompareMode = vbTextCompare
    Dim newBrZbr As String: newBrZbr = ""

    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function
    Dim pcBr As Long, pcId As Long, pcKl As Long, pcKol As Long, pcAmb As Long, pcZbr As Long, pcSt As Long
    pcBr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)
    pcId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
    pcKl = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, SRC)
    pcKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, SRC)
    pcAmb = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB, SRC)
    pcZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
    pcSt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    ' identitet nove prijemnice (za relabel-u-mestu; svi redovi broja dele isti)
    Dim pcVr As Long, pcSo As Long, pcTa As Long
    pcVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    pcSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    pcTa = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_TIP_AMB)
    Dim nVr As String, nSo As String, nTa As String

    Dim r As Long
    For r = 1 To UBound(prj, 1)
        If Trim$(CStr(prj(r, pcBr))) = newBroj Then
            Dim stN As Boolean: stN = False
            If pcSt > 0 Then stN = (UCase$(Trim$(CStr(prj(r, pcSt)))) = "DA")
            If Not stN Then
                Dim kl As String: kl = Trim$(CStr(prj(r, pcKl)))
                If Len(kl) = 0 Then kl = "I"
                newById(kl) = Trim$(CStr(prj(r, pcId)))
                newNeto(kl) = NzD(prj(r, pcKol))
                newGajb(kl) = NzL(prj(r, pcAmb))
                If Len(newBrZbr) = 0 Then
                    newBrZbr = Trim$(CStr(prj(r, pcZbr)))
                    nVr = Trim$(NzToText(SafeCell(prj, r, pcVr)))
                    nSo = Trim$(NzToText(SafeCell(prj, r, pcSo)))
                    If pcTa > 0 Then nTa = Trim$(NzToText(SafeCell(prj, r, pcTa)))
                End If
            End If
        End If
    Next r
    If newById.count = 0 Then
        outWarn = "Nova prijemnica " & newBroj & " nije aktivna / ne postoji."
        Exit Function
    End If

    ' --- Jedan citanje stavki; sakupi fresh(new) i orphan(old) redove + old gajbica po klasi. ---
    Dim ps As Variant: ps = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(ps) Then
        outWarn = "Nema paleta-stavki."
        Exit Function
    End If
    Dim sBr As Long, sPal As Long, sKl As Long, sGajb As Long, sNeto As Long, sAmb As Long, sPid As Long, sZbr As Long, sSt As Long
    sBr = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, SRC)
    sPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, SRC)
    sKl = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_KLASA, SRC)
    sGajb = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA, SRC)
    sNeto = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO, SRC)
    sAmb = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_AMBALAZA, SRC)
    sPid = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PRIJEMNICA_ID, SRC)
    sZbr = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE, SRC)
    sSt = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, SRC)

    Dim freshRows As Collection: Set freshRows = New Collection
    Dim oldRows As Collection: Set oldRows = New Collection
    Dim oldGajbByKl As Object: Set oldGajbByKl = CreateObject("Scripting.Dictionary")
    oldGajbByKl.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To UBound(ps, 1)
        If UCase$(Trim$(CStr(ps(i, sSt)))) <> "DA" Then
            Dim bp As String: bp = Trim$(CStr(ps(i, sBr)))
            If bp = newBroj Then
                freshRows.Add i
            ElseIf bp = oldBroj Then
                oldRows.Add i
                Dim kO As String: kO = Trim$(CStr(ps(i, sKl)))
                If Len(kO) = 0 Then kO = "I"
                oldGajbByKl(kO) = NzL(oldGajbByKl(kO)) + NzL(ps(i, sGajb))
            End If
        End If
    Next i
    If oldRows.count = 0 Then
        outWarn = "Nema osirocenih paleta-stavki za prijemnicu " & oldBroj & "."
        Exit Function
    End If

    ' Identitet-guard: razlika vrsta/sorta/tipAmb -> re-point bez relabela bi ostavio
    ' pogresno oznacenu paletu (paleta jos nosi staru etiketu). Trazi potvrdu (allowRelabel).
    Dim verdict As Variant: verdict = EvaluatePaletaReassign(oldBroj, newBroj)
    Dim relabelNeeded As Boolean: relabelNeeded = (CStr(verdict(0)) = "RELABEL")
    If relabelNeeded And Not allowRelabel Then
        outWarn = "Razlika u identitetu (" & CStr(verdict(2)) & ") - re-point bi ostavio " & _
                  "pogresno oznacenu paletu. Potrebna potvrda relabela."
        Exit Function
    End If

    ' HARD GUARD (co-tenant relabel): fizicka paleta je HOMOGENA (vrsta/sorta/tipAmb).
    ' Ako RELABEL, a neka od paleta koje nose ove stavke deli robu sa DRUGOM prijemnicom
    ' (aktivna stavka bp != oldBroj/newBroj), promena identiteta headera bi TIHO iskvarila
    ' identitet te druge robe. To se NE sme dozvoliti ni uz potvrdu -> blokada (ne warning).
    ' Fizicki ispravno resenje: skini ove stavke sa deljene palete pa unesi kao nov unos
    ' (sveza paletizacija na ispravno oznacenu paletu). Radi i kad je allowRelabel=True.
    Dim k As Long
    If relabelNeeded Then
        Dim tgtPal As Object: Set tgtPal = CreateObject("Scripting.Dictionary")
        For k = 1 To oldRows.count
            tgtPal(CStr(ps(oldRows(k), sPal))) = True
        Next k
        Dim qg As Long, sharedPal As String
        For qg = 1 To UBound(ps, 1)
            If UCase$(Trim$(CStr(ps(qg, sSt)))) <> "DA" Then
                Dim bpg As String: bpg = Trim$(CStr(ps(qg, sBr)))
                If bpg <> oldBroj And bpg <> newBroj Then
                    If tgtPal.Exists(CStr(ps(qg, sPal))) Then
                        sharedPal = PaletaLabel(CStr(ps(qg, sPal)))
                        Exit For
                    End If
                End If
            End If
        Next qg
        If Len(sharedPal) > 0 Then
            outWarn = "BLOKIRANO: paleta " & sharedPal & " deli robu sa drugom prijemnicom (" & bpg & _
                      "). Promena identiteta (" & CStr(verdict(2)) & ") bi iskvarila i tu robu. " & _
                      "Skini stavke ove prijemnice sa palete (Mod: Palete -> Skini stavke), pa unesi " & _
                      "robu kao nov unos na ispravno oznacenu paletu."
            Exit Function
        End If
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    ' ---- STEP 1: undo sveze (new) paletizacije ----
    For k = 1 To freshRows.count
        i = freshRows(k)
        DecrementPaletaForStavka CStr(ps(i, sPal)), NzL(ps(i, sGajb)), NzD(ps(i, sNeto)), NzD(ps(i, sAmb)), SRC
        RequireUpdateCell TBL_PALETA_STAVKA, i, COL_STORNIRANO, "Da", SRC
    Next k

    ' ---- STEP 2: delta-warn + re-point + KG-sync ----
    Dim warnMsg As String: warnMsg = ""
    Dim kk As Variant
    For Each kk In oldGajbByKl.Keys
        If Not newById.Exists(kk) Then
            warnMsg = warnMsg & "Klasa " & kk & ": nova prijemnica nema tu klasu - stavke ostaju " & _
                      "osirocene (skini ih: Osiroceni dokumenti -> Palete -> Skini stavke). "
            outGajbDiff = True
        ElseIf NzL(oldGajbByKl(kk)) <> NzL(newGajb(kk)) Then
            warnMsg = warnMsg & "Klasa " & kk & ": gajbica staro=" & NzL(oldGajbByKl(kk)) & " novo=" & _
                      NzL(newGajb(kk)) & ". "
            outGajbDiff = True
        End If
    Next kk
    ' Nova klasa bez starih stavki (npr. ispravka dodala Klasu II) -> i to je delta.
    For Each kk In newById.Keys
        If Not oldGajbByKl.Exists(kk) And NzL(newGajb(kk)) > 0 Then outGajbDiff = True
    Next kk

    For k = 1 To oldRows.count
        i = oldRows(k)
        Dim kl3 As String: kl3 = Trim$(CStr(ps(i, sKl)))
        If Len(kl3) = 0 Then kl3 = "I"
        If newById.Exists(kl3) Then
            RequireUpdateCell TBL_PALETA_STAVKA, i, COL_PALS_PRIJEMNICA_ID, CStr(newById(kl3)), SRC
            RequireUpdateCell TBL_PALETA_STAVKA, i, COL_PALS_BROJ_PRIJ, newBroj, SRC
            If Len(newBrZbr) > 0 Then RequireUpdateCell TBL_PALETA_STAVKA, i, COL_PALS_BROJ_ZBIRNE, newBrZbr, SRC
            ' KG-sync samo kad se broj gajbica klase poklapa
            If NzL(oldGajbByKl(kl3)) = NzL(newGajb(kl3)) And NzL(newGajb(kl3)) > 0 Then
                Dim perG As Double: perG = NzD(newNeto(kl3)) / NzL(newGajb(kl3))
                Dim oldStNeto As Double: oldStNeto = NzD(ps(i, sNeto))
                Dim newStNeto As Double: newStNeto = NzL(ps(i, sGajb)) * perG
                If Abs(newStNeto - oldStNeto) > 0.0001 Then
                    RequireUpdateCell TBL_PALETA_STAVKA, i, COL_PALS_NETO, newStNeto, SRC
                    AdjustPaletaNeto CStr(ps(i, sPal)), (newStNeto - oldStNeto), SRC
                End If
            End If
        End If
    Next k

    ' ---- STEP 2b: relabel-u-mestu (samo uz allowRelabel + razlika identiteta) ----
    ' Prepravi vrsta/sorta/tipAmb na PREVEZANIM stavkama + njihovim paletama. Roba se
    ' ne pomera; menja se etiketa da odgovara ispravljenoj prijemnici. Log po paleti.
    ' Deljene palete su vec odbijene hard-guardom gore, pa je relabel ovde bezbedan:
    ' svaka dodirnuta paleta nosi SAMO ovu prijemnicu (old->new).
    If relabelNeeded And allowRelabel Then
        Dim doneP As Object: Set doneP = CreateObject("Scripting.Dictionary")
        For k = 1 To oldRows.count
            i = oldRows(k)
            Dim klR As String: klR = Trim$(CStr(ps(i, sKl)))
            If Len(klR) = 0 Then klR = "I"
            If newById.Exists(klR) Then          ' samo stavke koje su stvarno prevezane
                RequireUpdateCell TBL_PALETA_STAVKA, i, COL_PALS_VRSTA, nVr, SRC
                RequireUpdateCell TBL_PALETA_STAVKA, i, COL_PALS_SORTA, nSo, SRC
                Dim pidR As String: pidR = CStr(ps(i, sPal))
                If Not doneP.Exists(pidR) Then
                    doneP(pidR) = True
                    Dim pRow As Long: pRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, pidR)
                    If pRow > 0 Then
                        RequireUpdateCell TBL_PALETA, pRow, COL_PAL_VRSTA, nVr, SRC
                        RequireUpdateCell TBL_PALETA, pRow, COL_PAL_SORTA, nSo, SRC
                        If Len(nTa) > 0 Then RequireUpdateCell TBL_PALETA, pRow, COL_PAL_TIP_AMBALAZE, nTa, SRC
                        PaletaLog pidR, "RELABEL", "prij " & oldBroj & "->" & newBroj & " (" & CStr(verdict(2)) & ")"
                    End If
                End If
            End If
        Next k
    End If

    tx.CommitTx
    Set tx = Nothing
    outWarn = warnMsg
    ReassignPaleteToPrijemnica_TX = True
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    ReassignPaleteToPrijemnica_TX = False
End Function

' Skini gajbica/neto/amb sa palete za jednu ponistenu stavku; reopen ako padne ispod kapaciteta.
Private Sub DecrementPaletaForStavka(ByVal palID As String, ByVal gajb As Long, _
                                     ByVal neto As Double, ByVal amb As Double, ByVal SRC As String)
    Dim palRow As Long: palRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, palID)
    If palRow = 0 Then Exit Sub
    Dim used As Long, pNeto As Double, pAmb As Double, palk As Double, cap As Long
    GetPaletaAggregates palRow, used, pNeto, pAmb, palk, cap
    Dim nUsed As Long: nUsed = used - gajb: If nUsed < 0 Then nUsed = 0
    Dim nNeto As Double: nNeto = pNeto - neto: If nNeto < 0 Then nNeto = 0
    Dim nAmb As Double: nAmb = pAmb - amb: If nAmb < 0 Then nAmb = 0
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BR_GAJBICA, nUsed, SRC
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_NETO, nNeto, SRC
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_AMBALAZA, nAmb, SRC
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BRUTO, nNeto + nAmb + palk, SRC
    If cap > 0 And nUsed < cap Then
        Dim d As Variant: d = GetTableData(TBL_PALETA)
        If CStr(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_STATUS))) = PAL_STATUS_ZATVORENA Then
            RequireUpdateCell TBL_PALETA, palRow, COL_PAL_STATUS, PAL_STATUS_OTVORENA, SRC
        End If
    End If
End Sub

' Uskladi total neto palete za KG-sync (+ recompute Bruto).
Private Sub AdjustPaletaNeto(ByVal palID As String, ByVal deltaNeto As Double, ByVal SRC As String)
    Dim palRow As Long: palRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, palID)
    If palRow = 0 Then Exit Sub
    Dim used As Long, pNeto As Double, pAmb As Double, palk As Double, cap As Long
    GetPaletaAggregates palRow, used, pNeto, pAmb, palk, cap
    Dim nn As Double: nn = pNeto + deltaNeto: If nn < 0 Then nn = 0
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_NETO, nn, SRC
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BRUTO, nn + pAmb + palk, SRC
End Sub

' Audit trag za izmenu palete (relabel/detach/adjust) preko Monitor_Event + vidljiva
' "Istorija" kolona na tblPaleta. Best-effort (nikad ne rusi TX).
Private Sub PaletaLog(ByVal palID As String, ByVal action As String, ByVal detail As String)
    On Error Resume Next
    Monitor_Event eventType:="PALETA_" & action, severity:="INFO", _
        message:=action & " paleta=" & palID & " " & detail, _
        moduleName:="modPaletniList", procedureName:="PaletaLog", _
        entityType:="Paleta", entityID:=palID, correlationId:=palID

    ' Vidljivi trag na paleti: append imenovanih parova (act=..;det=..;t=..),
    ' delimiter-safe (";" u detalju -> ","). Best-effort; samo ako kolona postoji.
    Dim ci As Long: ci = GetColumnIndex(TBL_PALETA, COL_PAL_ISTORIJA)
    If ci > 0 Then
        Dim pr As Long: pr = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, palID)
        If pr > 0 Then
            Dim d As Variant: d = GetTableData(TBL_PALETA)
            Dim prev As String: prev = CStr(SafeCell(d, pr, ci))
            Dim entry As String
            entry = "act=" & action & ";det=" & Replace(detail, ";", ",") & _
                    ";t=" & Format$(Now, "yyyy-mm-dd hh:nn")
            If Len(prev) > 0 Then entry = prev & " | " & entry
            RequireUpdateCell TBL_PALETA, pr, COL_PAL_ISTORIJA, entry, "modPaletniList.PaletaLog"
        End If
    End If
End Sub

' Persistentan business-trag kad ispravka (skip paletizacije) ne uspe da preveze
' palete: nova prijemnica je snimljena ali NEPALETIZOVANA, a stare palete su
' osirocene. MsgBox je prolazan; ovo ide u Monitor (WARN) da stanje ostane vidljivo
' i posle klika. Osirocene palete su i dalje vidljive u recovery panelu (Mod: Palete)
' i hvata ih Integritet C4 (stavka -> stornirana prijemnica). Best-effort.
Public Sub LogRelinkFailure(ByVal oldBroj As String, ByVal newBroj As String, _
                            ByVal razlog As String)
    On Error Resume Next
    Monitor_Event eventType:="PALETA_RELINK_FAIL", severity:="WARN", _
        message:="Ispravka: prijemnica " & newBroj & " snimljena ali NEPALETIZOVANA " & _
                 "(prevezivanje sa " & oldBroj & " nije uspelo). Osirocene palete " & _
                 "stare prijemnice cekaju rucni re-point (Mod: Palete). Razlog: " & razlog, _
        moduleName:="modPaletniList", procedureName:="LogRelinkFailure", _
        entityType:="Prijemnica", entityID:=newBroj, correlationId:=oldBroj
End Sub

' "Broj/Godina" oznaka palete za poruke (npr. "12/2026"); "?" ako ne postoji.
Private Function PaletaLabel(ByVal palID As String) As String
    On Error Resume Next
    PaletaLabel = "?"
    Dim pr As Long: pr = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, palID)
    If pr = 0 Then Exit Function
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    If IsEmpty(d) Then Exit Function
    PaletaLabel = CStr(NzL(SafeCell(d, pr, GetColumnIndex(TBL_PALETA, COL_PAL_BROJ)))) & "/" & _
                  CStr(NzL(SafeCell(d, pr, GetColumnIndex(TBL_PALETA, COL_PAL_GODINA))))
End Function

' ============================================================
' VERDIKT: presudi kako ciljna prijemnica (newBroj) stoji prema osirocenim
' paleta-stavkama stornirane (oldBroj). Read-only, bez TX. Jedan izvor istine
' za UI ocenu + relabel-gejt u ReassignPaleteToPrijemnica_TX.
'   CLEAN   = isti identitet + iste gajbice po klasi -> obican re-point
'   GAJBICA = isti identitet, razlicit broj gajbica (PER KLASA) -> re-point + koriguj
'   RELABEL = razlicit vrsta/sorta/tipAmb (roba ista) -> re-point + prepravi etiketu
'   NONE    = nova nije aktivna / nema osirocenih stavki
' Vraca Array(0..2): kategorija, kratka oznaka (za UI kolonu), detalj.
' ============================================================
Public Function EvaluatePaletaReassign(ByVal oldBroj As String, _
                                       ByVal newBroj As String) As Variant
    On Error GoTo EH
    oldBroj = Trim$(oldBroj): newBroj = Trim$(newBroj)
    EvaluatePaletaReassign = Array("NONE", "", "")
    If Len(oldBroj) = 0 Or Len(newBroj) = 0 Then Exit Function

    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then Exit Function
    Dim pBr As Long, pKl As Long, pVr As Long, pSo As Long, pTa As Long, pAmb As Long, pStt As Long
    pBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    pKl = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA)
    pVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    pSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    pTa = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_TIP_AMB)
    pAmb = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB)
    pStt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    If pBr = 0 Then Exit Function

    ' NOVA: identitet (prvi aktivni red) + gajbice PO KLASI. STARA: identitet (prvi red).
    Dim nVr As String, nSo As String, nTa As String, haveNew As Boolean
    Dim oVr As String, oSo As String, oTa As String, haveOldId As Boolean
    Dim newGajb As Object: Set newGajb = CreateObject("Scripting.Dictionary")
    newGajb.CompareMode = vbTextCompare
    Dim r As Long, kl As String
    For r = 1 To UBound(prj, 1)
        Dim b As String: b = Trim$(CStr(SafeCell(prj, r, pBr)))
        If b = newBroj Then
            If Not (pStt > 0 And UCase$(Trim$(CStr(SafeCell(prj, r, pStt)))) = "DA") Then
                If Not haveNew Then
                    nVr = Trim$(NzToText(SafeCell(prj, r, pVr)))
                    nSo = Trim$(NzToText(SafeCell(prj, r, pSo)))
                    If pTa > 0 Then nTa = Trim$(NzToText(SafeCell(prj, r, pTa)))
                    haveNew = True
                End If
                kl = Trim$(CStr(SafeCell(prj, r, pKl))): If Len(kl) = 0 Then kl = "I"
                newGajb(kl) = NzL(newGajb(kl)) + NzL(SafeCell(prj, r, pAmb))
            End If
        ElseIf b = oldBroj Then
            If Not haveOldId Then
                oVr = Trim$(NzToText(SafeCell(prj, r, pVr)))
                oSo = Trim$(NzToText(SafeCell(prj, r, pSo)))
                If pTa > 0 Then oTa = Trim$(NzToText(SafeCell(prj, r, pTa)))
                haveOldId = True
            End If
        End If
    Next r
    If Not haveNew Then EvaluatePaletaReassign = Array("NONE", "nova nije aktivna", ""): Exit Function

    ' STARA: gajbice PO KLASI iz aktivnih (osirocenih) stavki.
    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function
    Dim sBr As Long, sKl As Long, sGa As Long, sStt As Long
    sBr = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ)
    sKl = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_KLASA)
    sGa = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
    sStt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
    Dim oldGajb As Object: Set oldGajb = CreateObject("Scripting.Dictionary")
    oldGajb.CompareMode = vbTextCompare
    Dim haveOld As Boolean
    For r = 1 To UBound(s, 1)
        If Trim$(CStr(SafeCell(s, r, sBr))) = oldBroj _
           And UCase$(Trim$(CStr(SafeCell(s, r, sStt)))) <> "DA" Then
            kl = Trim$(CStr(SafeCell(s, r, sKl))): If Len(kl) = 0 Then kl = "I"
            oldGajb(kl) = NzL(oldGajb(kl)) + NzL(SafeCell(s, r, sGa))
            haveOld = True
        End If
    Next r
    If Not haveOld Then EvaluatePaletaReassign = Array("NONE", "nema osirocenih stavki", ""): Exit Function

    ' --- presuda: identitet pa gajbice PER KLASA ---
    Dim idDiff As String
    If StrComp(oVr, nVr, vbTextCompare) <> 0 Then idDiff = idDiff & "Vrsta " & oVr & "->" & nVr & " "
    If StrComp(oSo, nSo, vbTextCompare) <> 0 Then idDiff = idDiff & "Sorta " & oSo & "->" & nSo & " "
    If StrComp(oTa, nTa, vbTextCompare) <> 0 Then idDiff = idDiff & "TipAmb " & oTa & "->" & nTa & " "
    If Len(idDiff) > 0 Then
        EvaluatePaletaReassign = Array("RELABEL", "Prevezi + etiketa", Trim$(idDiff))
        Exit Function
    End If

    Dim gDiff As String, v As Variant
    For Each v In oldGajb.Keys
        If Not newGajb.Exists(v) Then
            gDiff = gDiff & "Kl." & CStr(v) & " " & NzL(oldGajb(v)) & "->nema; "
        ElseIf NzL(oldGajb(v)) <> NzL(newGajb(v)) Then
            gDiff = gDiff & "Kl." & CStr(v) & " " & NzL(oldGajb(v)) & "->" & NzL(newGajb(v)) & "; "
        End If
    Next v
    For Each v In newGajb.Keys
        If Not oldGajb.Exists(v) And NzL(newGajb(v)) > 0 Then
            gDiff = gDiff & "Kl." & CStr(v) & " 0->" & NzL(newGajb(v)) & "; "
        End If
    Next v

    If Len(gDiff) > 0 Then
        EvaluatePaletaReassign = Array("GAJBICA", "Prevezi + koriguj", Trim$(gDiff))
    Else
        EvaluatePaletaReassign = Array("CLEAN", "Prevezi", "cisto")
    End If
    Exit Function
EH:
    LogErr "modPaletniList.EvaluatePaletaReassign"
    EvaluatePaletaReassign = Array("NONE", "", "")
End Function

' ============================================================
' DETACH (dupli unos / los utovar): skini osirocene paleta-stavke stornirane
' prijemnice (oldBroj) sa njihovih paleta -- BEZ prevezivanja. Per-stavka ->
' su-stanari (druge prijemnice na istoj paleti) ostaju NETAKNUTI. Paleta koja
' posle skidanja ostane BEZ aktivnih stavki (fantom paleta duplog unosa) se
' automatski stornira ("storno svega sto je pogresna prijemnica stvorila");
' preradjena paleta se NE stornira (samo napomena). Transakciono.
' Vraca broj skinutih stavki; outInfo nosi rezime za UI.
' ============================================================
Public Function DetachOsirocenePaletaStavke_TX(ByVal oldBroj As String, _
                                               Optional ByRef outInfo As String) As Long
    Const SRC As String = "modPaletniList.DetachOsirocenePaletaStavke_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    outInfo = ""
    oldBroj = Trim$(oldBroj)
    If Len(oldBroj) = 0 Then Exit Function

    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function
    Dim sBr As Long, sPal As Long, sGa As Long, sNe As Long, sAmb As Long, sStt As Long
    sBr = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, SRC)
    sPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, SRC)
    sGa = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA, SRC)
    sNe = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO, SRC)
    sAmb = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_AMBALAZA, SRC)
    sStt = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, SRC)

    Dim rowsC As Collection: Set rowsC = New Collection
    Dim r As Long
    For r = 1 To UBound(s, 1)
        If Trim$(CStr(SafeCell(s, r, sBr))) = oldBroj _
           And UCase$(Trim$(CStr(SafeCell(s, r, sStt)))) <> "DA" Then rowsC.Add r
    Next r
    If rowsC.count = 0 Then
        outInfo = "Nema osirocenih stavki za " & oldBroj & "."
        Exit Function
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    Dim touched As Object: Set touched = CreateObject("Scripting.Dictionary")
    Dim k As Long, cnt As Long
    For k = 1 To rowsC.count
        r = rowsC(k)
        DecrementPaletaForStavka CStr(SafeCell(s, r, sPal)), NzL(SafeCell(s, r, sGa)), _
                                 NzD(SafeCell(s, r, sNe)), NzD(SafeCell(s, r, sAmb)), SRC
        RequireUpdateCell TBL_PALETA_STAVKA, r, COL_STORNIRANO, "Da", SRC
        touched(CStr(SafeCell(s, r, sPal))) = True
        PaletaLog CStr(SafeCell(s, r, sPal)), "DETACH", "prij=" & oldBroj & " gajb=" & NzL(SafeCell(s, r, sGa))
        cnt = cnt + 1
    Next k

    ' Fantom palete: dodirnuta paleta bez ijedne preostale aktivne stavke -> storno
    ' kroz kanonski modStorno.StornoPaleta (kompozicija u nasem TX, kao sto
    ' StornoOtkupByBrDok_TX komponuje StornoOtkup). Preradjena se preskace uz napomenu.
    Dim emptied As String
    Dim v As Variant
    For Each v In touched.Keys
        If IsEmpty(GetPaletaStavkeForGrid(CStr(v))) Then      ' nema aktivnih stavki
            If IsPaletaPreradjena(CStr(v)) Then
                emptied = emptied & PaletaLabel(CStr(v)) & " (preradjena - NIJE stornirana!) "
            ElseIf StornoPaleta(CStr(v)) Then
                PaletaLog CStr(v), "STORNO_PRAZNA", "detach prij=" & oldBroj
                emptied = emptied & PaletaLabel(CStr(v)) & " "
            End If
        End If
    Next v

    tx.CommitTx
    Set tx = Nothing
    outInfo = "Skinuto stavki: " & cnt & " (prijemnica " & oldBroj & ")."
    If Len(emptied) > 0 Then outInfo = outInfo & " Ispraznjene palete stornirane: " & Trim$(emptied) & "."
    DetachOsirocenePaletaStavke_TX = cnt
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    DetachOsirocenePaletaStavke_TX = 0
End Function

' ============================================================
' KOREKCIJA KOLICINA U MESTU: uskladi paleta-stavke AKTIVNE prijemnice (brojPrij)
' sa njenim dokument-vrednostima (KolAmb + Kolicina po klasi) -- roba se NE pomera,
' koriguje se evidencija na paletama koje je vec nose.
'   delta < 0: skini sa POSLEDNJE stavke te prijemnice (fizicki vrh); stavka koja
'              padne na 0 se stornira; paleta ispod kapaciteta se reopen-uje.
'   delta > 0, staje na paletu poslednje stavke: dopuni tu stavku.
'   delta > 0, NE staje: spillMode odlucuje:
'       "PRELIJ" -> dopuni do kapaciteta pa visak na sledecu otvorenu/novu paletu
'       "PREKO"  -> sve na istu paletu PREKO kapaciteta (operater svesno slaze vise)
'       ""       -> nista se ne menja; vraca ADJ_NEEDS_CHOICE (UI pita operatera)
'   klasa bez ijedne stavke (ispravka dodala klasu): normalna sveza paletizacija
'   te klase (bez pitanja).
' Neto/ambalaza kg: posle gajbica se SVE stavke klase re-sinhronizuju proporcionalno
' (neto = gajb * dokNeto/dokGajb; amb = gajb * tezina gajbice), pa se headeri
' dodirnutih paleta preracunaju iz aktivnih stavki (ukljucuje su-stanare).
' Preradjena paleta u obuhvatu = ADJ_BLOCKED (prvo storno prerade).
' Vraca: >=0 broj korigovanih klasa; ADJ_NEEDS_CHOICE; ADJ_BLOCKED.
' ============================================================
Public Function AdjustPaletaGajbiceZaPrijemnicu_TX(ByVal brojPrij As String, _
        Optional ByVal spillMode As String = "", _
        Optional ByRef outInfo As String) As Long
    Const SRC As String = "modPaletniList.AdjustPaletaGajbiceZaPrijemnicu_TX"
    Dim tx As clsTransaction
    On Error GoTo EH
    outInfo = ""
    brojPrij = Trim$(brojPrij)
    If Len(brojPrij) = 0 Then AdjustPaletaGajbiceZaPrijemnicu_TX = ADJ_BLOCKED: Exit Function
    spillMode = UCase$(Trim$(spillMode))

    ' --- Dokument (aktivni redovi prijemnice) po klasi ---
    Dim prj As Variant: prj = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(prj) Then AdjustPaletaGajbiceZaPrijemnicu_TX = ADJ_BLOCKED: Exit Function
    Dim pBr As Long, pKl As Long, pId As Long, pKol As Long, pAmb As Long, pStt As Long
    Dim pVr As Long, pSo As Long, pTa As Long, pZbr As Long
    pBr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)
    pKl = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, SRC)
    pId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
    pKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, SRC)
    pAmb = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB, SRC)
    pZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
    pVr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA)
    pSo = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA)
    pTa = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_TIP_AMB)
    pStt = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)

    Dim docId As Object, docNeto As Object, docGajb As Object
    Set docId = CreateObject("Scripting.Dictionary"): docId.CompareMode = vbTextCompare
    Set docNeto = CreateObject("Scripting.Dictionary"): docNeto.CompareMode = vbTextCompare
    Set docGajb = CreateObject("Scripting.Dictionary"): docGajb.CompareMode = vbTextCompare
    Dim dVr As String, dSo As String, dTa As String, dZbr As String
    Dim r As Long, kl As String
    For r = 1 To UBound(prj, 1)
        If Trim$(CStr(SafeCell(prj, r, pBr))) = brojPrij Then
            If Not (pStt > 0 And UCase$(Trim$(CStr(SafeCell(prj, r, pStt)))) = "DA") Then
                kl = Trim$(CStr(SafeCell(prj, r, pKl))): If Len(kl) = 0 Then kl = "I"
                docId(kl) = Trim$(CStr(SafeCell(prj, r, pId)))
                docNeto(kl) = NzD(SafeCell(prj, r, pKol))
                docGajb(kl) = NzL(SafeCell(prj, r, pAmb))
                If Len(dZbr) = 0 Then
                    dZbr = Trim$(NzToText(SafeCell(prj, r, pZbr)))
                    dVr = Trim$(NzToText(SafeCell(prj, r, pVr)))
                    dSo = Trim$(NzToText(SafeCell(prj, r, pSo)))
                    If pTa > 0 Then dTa = Trim$(NzToText(SafeCell(prj, r, pTa)))
                End If
            End If
        End If
    Next r
    If docId.count = 0 Then
        outInfo = "Prijemnica " & brojPrij & " nije aktivna / ne postoji."
        AdjustPaletaGajbiceZaPrijemnicu_TX = ADJ_BLOCKED
        Exit Function
    End If

    ' --- Aktivne stavke prijemnice po klasi (hronoloski = redosled u tabeli) ---
    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    Dim sBr As Long, sPal As Long, sKl As Long, sGa As Long, sNe As Long, sAm As Long, sStt As Long
    Dim stRows As Object: Set stRows = CreateObject("Scripting.Dictionary")   ' kl -> Collection(row)
    stRows.CompareMode = vbTextCompare
    Dim stGajb As Object: Set stGajb = CreateObject("Scripting.Dictionary")   ' kl -> suma gajbica
    stGajb.CompareMode = vbTextCompare
    If Not IsEmpty(s) Then
        sBr = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, SRC)
        sPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, SRC)
        sKl = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_KLASA, SRC)
        sGa = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA, SRC)
        sNe = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO, SRC)
        sAm = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_AMBALAZA, SRC)
        sStt = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, SRC)
        For r = 1 To UBound(s, 1)
            If Trim$(CStr(SafeCell(s, r, sBr))) = brojPrij _
               And UCase$(Trim$(CStr(SafeCell(s, r, sStt)))) <> "DA" Then
                kl = Trim$(CStr(SafeCell(s, r, sKl))): If Len(kl) = 0 Then kl = "I"
                If Not stRows.Exists(kl) Then stRows.Add kl, New Collection
                stRows(kl).Add r
                stGajb(kl) = NzL(stGajb(kl)) + NzL(SafeCell(s, r, sGa))
            End If
        Next r
    End If

    ' --- DRY faza: preradjena blokada + potreba za izborom (visak ne staje) ---
    Dim crateW As Double: crateW = GetTezinaGajbice(dTa)
    Dim needTxt As String, blockTxt As String
    Dim vKl As Variant
    For Each vKl In docId.Keys
        Dim dlt As Long: dlt = NzL(docGajb(vKl)) - NzL(stGajb(vKl))
        Dim netoDiffKl As Boolean: netoDiffKl = ClassNetoOutOfSync(stRows, CStr(vKl), s, sGa, sNe, _
                                                NzL(docGajb(vKl)), NzD(docNeto(vKl)))
        If dlt <> 0 Or netoDiffKl Then
            ' preradjena paleta u obuhvatu klase?
            If stRows.Exists(vKl) Then
                Dim ck As Long
                For ck = 1 To stRows(vKl).count
                    If IsPaletaPreradjena(CStr(SafeCell(s, stRows(vKl)(ck), sPal))) Then
                        blockTxt = blockTxt & "Klasa " & CStr(vKl) & ": paleta " & _
                                   PaletaLabel(CStr(SafeCell(s, stRows(vKl)(ck), sPal))) & " je preradjena. "
                        Exit For
                    End If
                Next ck
            End If
            ' izbor potreban? (dodavanje preko slobodnog na poslednjoj paleti)
            If dlt > 0 And stRows.Exists(vKl) And Len(spillMode) = 0 Then
                Dim lastPalID As String
                lastPalID = CStr(SafeCell(s, stRows(vKl)(stRows(vKl).count), sPal))
                Dim lpRow As Long: lpRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, lastPalID)
                Dim u As Long, nn As Double, aa As Double, pk As Double, cp As Long
                GetPaletaAggregates lpRow, u, nn, aa, pk, cp
                If cp > 0 And dlt > (cp - u) Then
                    needTxt = needTxt & "Klasa " & CStr(vKl) & ": +" & dlt & " gajb. ne staje na paletu " & _
                              PaletaLabel(lastPalID) & " (slobodno " & (cp - u) & " od " & cp & "). "
                End If
            End If
        End If
    Next vKl
    ' Stavke klase koje dokument vise nema -> ne diramo ovde (Detach/Reassign teren).
    For Each vKl In stGajb.Keys
        If Not docId.Exists(vKl) Then
            blockTxt = blockTxt & "Klasa " & CStr(vKl) & ": stavke postoje a dokument nema tu klasu " & _
                       "(resi kroz Prevezi/Skini, ne kroz korekciju). "
        End If
    Next vKl
    If Len(blockTxt) > 0 Then
        outInfo = blockTxt
        AdjustPaletaGajbiceZaPrijemnicu_TX = ADJ_BLOCKED
        Exit Function
    End If
    If Len(needTxt) > 0 Then
        outInfo = needTxt
        AdjustPaletaGajbiceZaPrijemnicu_TX = ADJ_NEEDS_CHOICE
        Exit Function
    End If

    ' --- MUTACIJA (transakciono) ---
    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PALETA
    tx.AddTableSnapshot TBL_PALETA_STAVKA

    ' touched: palID -> akumulirana delta gajbica (0 = samo kg/amb re-sync).
    ' Kljucevi = palete za header-recompute + Istorija/ADJUST log sa vidljivom deltom.
    Dim touched As Object: Set touched = CreateObject("Scripting.Dictionary")
    Dim doneCnt As Long, infoTxt As String

    For Each vKl In docId.Keys
        kl = CStr(vKl)
        Dim delta As Long: delta = NzL(docGajb(kl)) - NzL(stGajb(kl))
        Dim changedKl As Boolean: changedKl = False

        If delta < 0 Then
            ' skidanje: od poslednje stavke ka prvoj
            Dim remM As Long: remM = -delta
            Dim k As Long
            For k = stRows(kl).count To 1 Step -1
                If remM = 0 Then Exit For
                r = stRows(kl)(k)
                Dim g As Long: g = NzL(SafeCell(s, r, sGa))
                Dim take As Long: take = g: If take > remM Then take = remM
                If take > 0 Then
                    If g - take = 0 Then
                        RequireUpdateCell TBL_PALETA_STAVKA, r, COL_STORNIRANO, "Da", SRC
                        RequireUpdateCell TBL_PALETA_STAVKA, r, COL_PALS_BR_GAJBICA, 0, SRC
                        RequireUpdateCell TBL_PALETA_STAVKA, r, COL_PALS_NETO, 0, SRC
                        RequireUpdateCell TBL_PALETA_STAVKA, r, COL_PALS_AMBALAZA, 0, SRC
                    Else
                        RequireUpdateCell TBL_PALETA_STAVKA, r, COL_PALS_BR_GAJBICA, g - take, SRC
                    End If
                    Dim pidD As String: pidD = CStr(SafeCell(s, r, sPal))
                    touched(pidD) = NzL(touched(pidD)) - take
                    remM = remM - take
                    changedKl = True
                End If
            Next k
        ElseIf delta > 0 Then
            If Not stRows.Exists(kl) Then
                ' nova klasa bez stavki -> sveza paletizacija te klase
                SpillGajbice CStr(docId(kl)), brojPrij, dZbr, kl, dVr, dSo, dTa, _
                             delta, NzD(docNeto(kl)) / NzL(docGajb(kl)), crateW, touched, SRC
                changedKl = True
            Else
                Dim lp As String: lp = CStr(SafeCell(s, stRows(kl)(stRows(kl).count), sPal))
                Dim lpR As Long: lpR = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, lp)
                Dim u2 As Long, n2 As Double, a2 As Double, p2 As Double, c2 As Long
                GetPaletaAggregates lpR, u2, n2, a2, p2, c2
                Dim freeSlots As Long: freeSlots = c2 - u2: If freeSlots < 0 Then freeSlots = 0
                Dim lastRow As Long: lastRow = stRows(kl)(stRows(kl).count)
                Dim gLast As Long: gLast = NzL(SafeCell(s, lastRow, sGa))
                If delta <= freeSlots Or spillMode = "PREKO" Or c2 = 0 Then
                    ' sve na istu paletu (staje, ili operater svesno preko kapaciteta)
                    RequireUpdateCell TBL_PALETA_STAVKA, lastRow, COL_PALS_BR_GAJBICA, gLast + delta, SRC
                    touched(lp) = NzL(touched(lp)) + delta
                    If spillMode = "PREKO" And delta > freeSlots Then _
                        infoTxt = infoTxt & "Paleta " & PaletaLabel(lp) & ": " & (u2 + delta) & _
                                  " gajb. (preko kapaciteta " & c2 & "). "
                    changedKl = True
                Else
                    ' PRELIJ: dopuni do kapaciteta pa visak na sledecu/nove palete
                    Dim fillN As Long: fillN = freeSlots
                    If fillN > 0 Then
                        RequireUpdateCell TBL_PALETA_STAVKA, lastRow, COL_PALS_BR_GAJBICA, gLast + fillN, SRC
                        touched(lp) = NzL(touched(lp)) + fillN
                    End If
                    If lpR > 0 Then ClosePaleta lpR, SRC   ' puna -> zatvori pre trazenja sledece
                    SpillGajbice CStr(docId(kl)), brojPrij, dZbr, kl, dVr, dSo, dTa, _
                                 delta - fillN, NzD(docNeto(kl)) / NzL(docGajb(kl)), crateW, touched, SRC
                    infoTxt = infoTxt & "Klasa " & kl & ": +" & fillN & " na " & PaletaLabel(lp) & _
                              ", +" & (delta - fillN) & " preliveno na sledecu paletu. "
                    changedKl = True
                End If
            End If
        End If

        ' KG/amb re-sync svih aktivnih stavki klase (fresh read posle gajbica-izmena)
        If NzL(docGajb(kl)) > 0 Then
            If ResyncStavkeNetoAmb(brojPrij, kl, NzD(docNeto(kl)) / NzL(docGajb(kl)), crateW, touched, SRC) Then changedKl = True
        End If
        If changedKl Then
            doneCnt = doneCnt + 1
            infoTxt = infoTxt & "Klasa " & kl & ": gajbice " & NzL(stGajb(kl)) & " -> " & NzL(docGajb(kl)) & _
                      ", neto " & Format$(NzD(docNeto(kl)), "0.##") & " kg. "
        End If
    Next vKl

    ' Header-i dodirnutih paleta = suma aktivnih stavki (self-healing, ukljucuje
    ' su-stanare). Log nosi VIDLJIVU deltu gajbica po paleti (+3 / -2 / 0 = samo kg).
    Dim vP As Variant
    For Each vP In touched.Keys
        RecomputePaletaFromStavke CStr(vP), SRC
        Dim dG As Long: dG = NzL(touched(vP))
        Dim dTxt As String
        If dG = 0 Then
            dTxt = "0 (samo kg)"
        Else
            dTxt = IIf(dG > 0, "+", "") & dG
        End If
        PaletaLog CStr(vP), "ADJUST", "prij=" & brojPrij & " gajb=" & dTxt
    Next vP

    tx.CommitTx
    Set tx = Nothing
    If doneCnt = 0 Then infoTxt = "Nista za korekciju (kolicine vec usaglasene)."
    outInfo = Trim$(infoTxt)
    AdjustPaletaGajbiceZaPrijemnicu_TX = doneCnt
    Exit Function
EH:
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    outInfo = "Greska: " & Err.description
    AdjustPaletaGajbiceZaPrijemnicu_TX = ADJ_BLOCKED
End Function

' Da li neto stavki klase odstupa od dokument-cilja (za dry proveru "ima li posla").
Private Function ClassNetoOutOfSync(ByVal stRows As Object, ByVal kl As String, _
        ByVal s As Variant, ByVal sGa As Long, ByVal sNe As Long, _
        ByVal docG As Long, ByVal docN As Double) As Boolean
    On Error Resume Next
    If Not stRows.Exists(kl) Then Exit Function
    If docG <= 0 Then Exit Function
    Dim perG As Double: perG = docN / docG
    Dim k As Long
    For k = 1 To stRows(kl).count
        Dim r As Long: r = stRows(kl)(k)
        If Abs(NzD(SafeCell(s, r, sNe)) - NzL(SafeCell(s, r, sGa)) * perG) > 0.0001 Then
            ClassNetoOutOfSync = True
            Exit Function
        End If
    Next k
End Function

' Da li je paleta preradjena (korekcije na njoj su blokirane dok se prerada ne stornira).
Private Function IsPaletaPreradjena(ByVal palID As String) As Boolean
    On Error Resume Next
    Dim pr As Long: pr = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, palID)
    If pr = 0 Then Exit Function
    Dim d As Variant: d = GetTableData(TBL_PALETA)
    IsPaletaPreradjena = (UCase$(Trim$(CStr(SafeCell(d, pr, _
        GetColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO))))) = "DA")
End Function

' ZAJEDNICKA raspodelna petlja: rasporedi "remaining" gajbica na otvorene/nove
' palete kao NOVE stavke date prijemnice + inkrementalno azuriraj header palete.
' Koriste je PaletizePrijemnica (sveza paletizacija; closedPalIDs za post-commit
' stampu) i AdjustPaletaGajbiceZaPrijemnicu_TX (PRELIJ visak / nova klasa).
' Jedna implementacija -> geometrija punjenja/zatvaranja ne moze da drift-uje.
Private Sub SpillGajbice(ByVal prijemnicaID As String, ByVal brojPrij As String, _
        ByVal brojZbirne As String, ByVal klasa As String, ByVal vrsta As String, _
        ByVal sorta As String, ByVal tipAmb As String, ByVal remaining As Long, _
        ByVal perG As Double, ByVal crateW As Double, ByVal touched As Object, _
        ByVal SRC As String, _
        Optional ByVal closedPalIDs As Collection = Nothing, _
        Optional ByRef nClosed As Long)
    Dim defCap As Long: defCap = GetKapacitetPalete(vrsta)
    Do While remaining > 0
        Dim palRow As Long, palID As String
        palID = GetOrCreateOpenPaleta(vrsta, sorta, klasa, tipAmb, defCap, palRow)
        If palID = "" Or palRow = 0 Then
            Err.Raise vbObjectError + 7331, SRC, _
                      "Ne mogu da otvorim/nadjem paletu za: " & vrsta
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
        takeNeto = take * perG
        takeAmb = take * crateW

        AddStavka palID, prijemnicaID, brojPrij, brojZbirne, klasa, _
                  vrsta, sorta, take, takeNeto, takeAmb

        Dim newGajb As Long: newGajb = used + take
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BR_GAJBICA, newGajb, SRC
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_NETO, curNeto + takeNeto, SRC
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_AMBALAZA, curAmb + takeAmb, SRC
        RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BRUTO, _
                          (curNeto + takeNeto) + (curAmb + takeAmb) + palKg, SRC

        touched(palID) = NzL(touched(palID)) + take   ' akumuliraj +delta gajbica po paleti
        remaining = remaining - take

        If newGajb >= cap Then
            ClosePaleta palRow, SRC
            If Not closedPalIDs Is Nothing Then closedPalIDs.Add palID
            nClosed = nClosed + 1
        End If
NextIter:
    Loop
End Sub

' Re-sinhronizuj Neto/Ambalazu SVIH aktivnih stavki (brojPrij, klasa) na
' proporcionalne vrednosti (gajb * perG / gajb * crateW). Fresh read (posle
' gajbica-izmena). Vraca True ako je nesto pisano; dodirnute palete u touched.
Private Function ResyncStavkeNetoAmb(ByVal brojPrij As String, ByVal klasa As String, _
        ByVal perG As Double, ByVal crateW As Double, ByVal touched As Object, _
        ByVal SRC As String) As Boolean
    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    If IsEmpty(s) Then Exit Function
    Dim sBr As Long, sPal As Long, sKl As Long, sGa As Long, sNe As Long, sAm As Long, sStt As Long
    sBr = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_PRIJ, SRC)
    sPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, SRC)
    sKl = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_KLASA, SRC)
    sGa = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA, SRC)
    sNe = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO, SRC)
    sAm = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_AMBALAZA, SRC)
    sStt = RequireColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO, SRC)
    Dim r As Long, kl As String
    For r = 1 To UBound(s, 1)
        If Trim$(CStr(SafeCell(s, r, sBr))) = brojPrij _
           And UCase$(Trim$(CStr(SafeCell(s, r, sStt)))) <> "DA" Then
            kl = Trim$(CStr(SafeCell(s, r, sKl))): If Len(kl) = 0 Then kl = "I"
            If StrComp(kl, klasa, vbTextCompare) = 0 Then
                Dim g As Long: g = NzL(SafeCell(s, r, sGa))
                Dim tgtN As Double: tgtN = g * perG
                Dim tgtA As Double: tgtA = g * crateW
                If Abs(NzD(SafeCell(s, r, sNe)) - tgtN) > 0.0001 Then
                    RequireUpdateCell TBL_PALETA_STAVKA, r, COL_PALS_NETO, tgtN, SRC
                    touched(CStr(SafeCell(s, r, sPal))) = NzL(touched(CStr(SafeCell(s, r, sPal))))
                    ResyncStavkeNetoAmb = True
                End If
                If Abs(NzD(SafeCell(s, r, sAm)) - tgtA) > 0.0001 Then
                    RequireUpdateCell TBL_PALETA_STAVKA, r, COL_PALS_AMBALAZA, tgtA, SRC
                    touched(CStr(SafeCell(s, r, sPal))) = NzL(touched(CStr(SafeCell(s, r, sPal))))
                    ResyncStavkeNetoAmb = True
                End If
            End If
        End If
    Next r
End Function

' Header palete = suma njenih AKTIVNIH stavki (gajbice/neto/amb; bruto = neto+amb+palKg).
' Status: otvorena i puna -> zatvori; zatvorena ispod kapaciteta -> otvori (mirror
' DecrementPaletaForStavka reopen semantike). Self-healing, ukljucuje su-stanare.
Private Sub RecomputePaletaFromStavke(ByVal palID As String, ByVal SRC As String)
    Dim palRow As Long: palRow = FindRowIndexByID(TBL_PALETA, COL_PAL_ID, palID)
    If palRow = 0 Then Exit Sub

    Dim s As Variant: s = GetTableData(TBL_PALETA_STAVKA)
    Dim sumG As Long, sumN As Double, sumA As Double
    If Not IsEmpty(s) Then
        Dim iPal As Long, iGa As Long, iNe As Long, iAm As Long, iSt As Long, r As Long
        iPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
        iGa = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BR_GAJBICA)
        iNe = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_NETO)
        iAm = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_AMBALAZA)
        iSt = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        For r = 1 To UBound(s, 1)
            If CStr(SafeCell(s, r, iPal)) = palID _
               And UCase$(Trim$(CStr(SafeCell(s, r, iSt)))) <> "DA" Then
                sumG = sumG + NzL(SafeCell(s, r, iGa))
                sumN = sumN + NzD(SafeCell(s, r, iNe))
                sumA = sumA + NzD(SafeCell(s, r, iAm))
            End If
        Next r
    End If

    Dim used As Long, pNeto As Double, pAmb As Double, palk As Double, cap As Long
    GetPaletaAggregates palRow, used, pNeto, pAmb, palk, cap
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BR_GAJBICA, sumG, SRC
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_NETO, sumN, SRC
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_AMBALAZA, sumA, SRC
    RequireUpdateCell TBL_PALETA, palRow, COL_PAL_BRUTO, sumN + sumA + palk, SRC

    Dim d As Variant: d = GetTableData(TBL_PALETA)
    Dim st As String: st = CStr(SafeCell(d, palRow, GetColumnIndex(TBL_PALETA, COL_PAL_STATUS)))
    If cap > 0 Then
        If sumG >= cap And st = PAL_STATUS_OTVORENA Then
            RequireUpdateCell TBL_PALETA, palRow, COL_PAL_STATUS, PAL_STATUS_ZATVORENA, SRC
        ElseIf sumG < cap And st = PAL_STATUS_ZATVORENA Then
            RequireUpdateCell TBL_PALETA, palRow, COL_PAL_STATUS, PAL_STATUS_OTVORENA, SRC
        End If
    End If
End Sub

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

' Otkupi za skup paleta: preko njihovih zbirni (tblPaletaStavka.BrojZbirne) ->
' tblOtkup, filtrirano po klasi tih paleta, dedup po OtkupID. Vraca 1-based 2D:
' 1 KooperantID(sifra), 2 VrstaVoca, 3 Kolicina(neto), 4 KolAmbalaze, 5 TipAmbalaze.
Private Function GetOtkupiZaPalete(ByVal paletaIDs As Collection) As Variant
    On Error GoTo EH
    If paletaIDs Is Nothing Then Exit Function
    If paletaIDs.count = 0 Then Exit Function

    Dim palSet As Object: Set palSet = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    For Each v In paletaIDs
        palSet(CStr(v)) = True
    Next v

    ' zbirne tih paleta (iz stavki)
    Dim zbSet As Object: Set zbSet = CreateObject("Scripting.Dictionary")
    Dim sp As Variant: sp = GetTableData(TBL_PALETA_STAVKA)
    Dim r As Long
    If Not IsEmpty(sp) Then
        Dim spPal As Long, spZb As Long, spStorno As Long
        spPal = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID)
        spZb = GetColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE)
        spStorno = GetColumnIndex(TBL_PALETA_STAVKA, COL_STORNIRANO)
        For r = 1 To UBound(sp, 1)
            If palSet.Exists(CStr(SafeCell(sp, r, spPal))) _
               And UCase$(Trim$(CStr(SafeCell(sp, r, spStorno)))) <> "DA" Then
                Dim zb As String: zb = Trim$(CStr(SafeCell(sp, r, spZb)))
                If zb <> "" Then zbSet(zb) = True
            End If
        Next r
    End If
    If zbSet.count = 0 Then Exit Function

    ' klase tih paleta
    Dim klSet As Object: Set klSet = CreateObject("Scripting.Dictionary")
    Dim dp As Variant: dp = GetTableData(TBL_PALETA)
    If Not IsEmpty(dp) Then
        Dim pid As Long, pKl As Long
        pid = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
        pKl = GetColumnIndex(TBL_PALETA, COL_PAL_KLASA)
        For r = 1 To UBound(dp, 1)
            If palSet.Exists(CStr(SafeCell(dp, r, pid))) Then
                klSet(UCase$(Trim$(CStr(SafeCell(dp, r, pKl))))) = True
            End If
        Next r
    End If

    ' klasa filter samo ako paleta ima definisanu klasu (legacy palete bez klase
    ' -> ne filtriraj po klasi, da lista ne bude prazna)
    Dim filterKlasa As Boolean
    Dim kk As Variant
    For Each kk In klSet.keys
        If Trim$(CStr(kk)) <> "" Then filterKlasa = True
    Next kk

    ' OtkupID-jevi preko zbirne: zbirna -> otpremnice -> otkupi (reuse TraceByZbirna).
    ' Otkup nije direktno vezan za BrojZbirne, nego preko OtpremnicaID.
    Dim wantOtk As Object: Set wantOtk = CreateObject("Scripting.Dictionary")
    Dim z As Variant
    For Each z In zbSet.keys
        Dim t As Variant: t = TraceByZbirna(CStr(z))
        If Not IsEmpty(t) Then
            Dim tr As Long
            For tr = LBound(t, 1) To UBound(t, 1)
                Dim tid As String: tid = Trim$(CStr(t(tr, 6)))   ' OtkupID = kolona 6
                If tid <> "" Then wantOtk(tid) = True
            Next tr
        End If
    Next z
    If wantOtk.count = 0 Then Exit Function

    ' detalji otkupa iz tblOtkup po OtkupID (+ klasa filter), dedup
    Dim o As Variant: o = GetTableData(TBL_OTKUP)
    If IsEmpty(o) Then Exit Function
    Dim oid As Long, oKoop As Long, oVr As Long, oKol As Long
    Dim oAmb As Long, oTip As Long, oKl As Long, oStorno As Long
    oid = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    oKoop = GetColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT)
    oVr = GetColumnIndex(TBL_OTKUP, COL_OTK_VRSTA)
    oKol = GetColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA)
    oAmb = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB)
    oTip = GetColumnIndex(TBL_OTKUP, COL_OTK_TIP_AMB)
    oKl = GetColumnIndex(TBL_OTKUP, COL_OTK_KLASA)
    oStorno = GetColumnIndex(TBL_OTKUP, COL_OTK_STORNIRANO)

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim rows As Collection: Set rows = New Collection
    For r = 1 To UBound(o, 1)
        Dim otkID As String: otkID = CStr(SafeCell(o, r, oid))
        If wantOtk.Exists(otkID) _
           And (Not filterKlasa Or klSet.Exists(UCase$(Trim$(CStr(SafeCell(o, r, oKl)))))) _
           And UCase$(Trim$(CStr(SafeCell(o, r, oStorno)))) <> "DA" Then
            If Not seen.Exists(otkID) Then
                seen(otkID) = True
                rows.Add r
            End If
        End If
    Next r
    If rows.count = 0 Then Exit Function

    Dim res As Variant: ReDim res(1 To rows.count, 1 To 5)
    Dim k As Long
    For k = 1 To rows.count
        r = rows(k)
        res(k, 1) = CStr(SafeCell(o, r, oKoop))
        res(k, 2) = CStr(SafeCell(o, r, oVr))
        res(k, 3) = NzD(SafeCell(o, r, oKol))
        res(k, 4) = NzL(SafeCell(o, r, oAmb))
        res(k, 5) = CStr(SafeCell(o, r, oTip))
    Next k

    GetOtkupiZaPalete = res
    Exit Function
EH:
    LogErr "modPaletniList.GetOtkupiZaPalete"
End Function

' Public (reuse iz frmDokumenta recovery panela i modOtkupBlok prefill-a; nema duplikata).
Public Function NzD(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then NzD = CDbl(v)
End Function

Public Function NzL(ByVal v As Variant) As Long
    On Error Resume Next
    If IsNumeric(v) Then NzL = CLng(v)
End Function

' Bezbedno citanje celije iz GetTableData niza: ako kolona ne postoji
' (idx < 1, npr. schema drift), vrati Empty umesto subscript-error.
Public Function SafeCell(ByVal d As Variant, ByVal r As Long, _
                          ByVal idx As Long) As Variant
    If idx >= 1 Then SafeCell = d(r, idx) Else SafeCell = Empty
End Function

' ============================================================
' Schema guards (fail-fast) + exact-row lookup za paletni/prerada domen.
' Kriticne kolone idu preko RequireColumns (modSchemaGuard); ID lookup puca
' na 0 i na >1 -- nikad "prvi od duplikata".
' ============================================================
Private Sub RequirePaletaSchema(ByVal SRC As String)
    RequireColumns TBL_PALETA, SRC, _
        COL_PAL_ID, COL_PAL_BROJ, COL_PAL_GODINA, COL_PAL_DATUM, COL_PAL_VRSTA, _
        COL_PAL_SORTA, COL_PAL_KLASA, COL_PAL_TIP_AMBALAZE, COL_PAL_TIP_PALETE, _
        COL_PAL_KAPACITET, COL_PAL_BR_GAJBICA, COL_PAL_NETO, COL_PAL_AMBALAZA, _
        COL_PAL_PALETA_KG, COL_PAL_BRUTO, COL_PAL_STATUS, COL_PAL_PRERADJENO, _
        COL_PAL_CREATED, COL_STORNIRANO
End Sub

Private Sub RequirePaletaStavkaSchema(ByVal SRC As String)
    RequireColumns TBL_PALETA_STAVKA, SRC, _
        COL_PALS_ID, COL_PALS_PALETA_ID, COL_PALS_PRIJEMNICA_ID, _
        COL_PALS_BROJ_PRIJ, COL_PALS_BROJ_ZBIRNE, COL_PALS_KLASA, _
        COL_PALS_VRSTA, COL_PALS_SORTA, COL_PALS_BR_GAJBICA, _
        COL_PALS_NETO, COL_PALS_AMBALAZA, COL_PALS_CREATED, COL_STORNIRANO
End Sub

Private Sub RequirePreradaSchema(ByVal SRC As String)
    RequireColumns TBL_PRERADA, SRC, _
        COL_PRE_ID, COL_PRE_BROJ, COL_PRE_GODINA, COL_PRE_NETO_IZLAZ, _
        COL_PRE_KUTIJE, COL_PRE_KESE, COL_STORNIRANO
End Sub

Private Sub RequirePreradaStavkaSchema(ByVal SRC As String)
    RequireColumns TBL_PRERADA_STAVKA, SRC, _
        COL_PRES_ID, COL_PRES_PRERADA_ID, COL_PRES_PALETA_ID, _
        COL_PRES_BROJ_PALETE, COL_PRES_NETO, COL_STORNIRANO
End Sub

' Schema-drift: dopuni tblPrerada kolone ako fale. Idempotentno (no-op kad
' postoje). Resava 0 u sazetku paletnog lista kada SetupPaletniListSchema nije
' pokrenut posle nadogradnje.
'
' Spisak kolona se NE drzi ovde. Ranije je bio lokalno hardkodiran (sest
' COL_PRE_*), pa bi izmena registra u modSetup-u ostavila ovu funkciju da leci
' zastarelu semu. Sada se trazi od registra "sredi tblPrerada" -- jedan izvor
' istine, kako i kaze .claude/rules/podaci-i-config.md.
'
' EnsureTableSchema je repair-only: ako tabele nema, ne pravi je (self-heal usred
' poslovnog toka ne sme da otvara nove listove -- to radi Setup* komanda).
Private Sub EnsurePreradaCols()
    On Error Resume Next
    EnsureTableSchema TBL_PRERADA
End Sub

' Exact-row lookup po kljucu. Puca ako nema reda (0) ili ima vise (>1).
' Za IDENTITET (PaletaID, PrijemnicaID), NE za pretragu otvorenih paleta.
Private Function RequireSingleRowIndexByKey(ByVal tblName As String, _
                                            ByVal keyCol As String, _
                                            ByVal keyValue As String, _
                                            ByVal SRC As String) As Long
    Dim hits As Collection
    Set hits = FindRows(tblName, keyCol, keyValue)
    If hits.count = 0 Then
        Err.Raise vbObjectError + 7320, SRC, _
                  "Nema reda u " & tblName & " za " & keyCol & "=" & keyValue & "."
    ElseIf hits.count > 1 Then
        Err.Raise vbObjectError + 7321, SRC, _
                  "Vi" & ChrW(353) & "e redova (" & hits.count & ") u " & tblName & " za " & _
                  keyCol & "=" & keyValue & "."
    End If
    RequireSingleRowIndexByKey = CLng(hits(1))
End Function

' ============================================================
' PRERADA (preradni list) -- palete -> kutije/kese.
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
            If CLng(val(CStr(data(r, iGod)))) = yr Then
                n = CLng(val(CStr(data(r, iBroj))))
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

' Snimi preradu (TX): operater je izabrao palete (kolekcija PaletaID-jeva).
' Snapshot tblPrerada/tblPreradaStavka/tblPaleta; validira (paleta postoji tacno
' jednom, nije stornirana, nije vec preradjena); upisuje preradu + stavke i
' markira palete Preradjeno=Da. Bez MsgBox -> baca gresku (UI je hvata).
' Operater bira otvorene i/ili zatvorene palete. Vraca PreradaID.
Public Function SavePrerada_TX(ByVal paletaIDs As Collection, _
                               ByVal brojKutija As Long, _
                               ByVal brojKesa As Long, _
                               ByVal netoIzlazKg As Double, _
                               Optional ByVal napomena As String = "", _
                               Optional ByVal tezinaPaleteKg As Double = 0, _
                               Optional ByVal brutoKg As Double = 0, _
                               Optional ByVal tezinaAmbalazeKg As Double = 0, _
                               Optional ByVal tipKutije As String = "", _
                               Optional ByVal tipKese As String = "", _
                               Optional ByVal tipGotovogProizvoda As String = "") As String
    Const SRC As String = "modPaletniList.SavePrerada_TX"

    Dim tx As clsTransaction
    On Error GoTo EH

    ' --- validacija ulaza ---
    If paletaIDs Is Nothing Then
        Err.Raise vbObjectError + 7340, SRC, "Nije izabrana nijedna paleta."
    End If
    If paletaIDs.count = 0 Then
        Err.Raise vbObjectError + 7340, SRC, "Nije izabrana nijedna paleta."
    End If
    If netoIzlazKg <= 0 Then
        Err.Raise vbObjectError + 7341, SRC, "Neto izlaz mora biti > 0."
    End If
    If brojKutija <= 0 And brojKesa <= 0 Then
        Err.Raise vbObjectError + 7342, SRC, "Broj kutija ili kesa mora biti > 0."
    End If

    ' Osiguraj nove kolone PRE transakcije (inace PalAppendRow tiho preskoci
    ' upis bruto/paleta/ambalaza -> 0 u sazetku paletnog lista).
    EnsurePreradaCols

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_PRERADA
    tx.AddTableSnapshot TBL_PRERADA_STAVKA
    tx.AddTableSnapshot TBL_PALETA

    RequirePreradaSchema SRC
    RequirePreradaStavkaSchema SRC
    RequirePaletaSchema SRC

    Dim dPal As Variant: dPal = GetTableData(TBL_PALETA)
    Dim iStorno As Long, iPre As Long, iNeto As Long, iBroj As Long
    iStorno = RequireColumnIndex(TBL_PALETA, COL_STORNIRANO, SRC)
    iPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, SRC)
    iNeto = RequireColumnIndex(TBL_PALETA, COL_PAL_NETO, SRC)
    iBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, SRC)

    ' --- resolve + validate svake palete (tacno jednom; dedup) ---
    Dim palRows As Object: Set palRows = CreateObject("Scripting.Dictionary") ' PaletaID -> rowIdx
    Dim netoUlaz As Double: netoUlaz = 0

    Dim v As Variant
    For Each v In paletaIDs
        Dim pid As String: pid = Trim$(CStr(v))
        If pid <> "" And Not palRows.Exists(pid) Then
            Dim rIdx As Long
            rIdx = RequireSingleRowIndexByKey(TBL_PALETA, COL_PAL_ID, pid, SRC)
            If UCase$(Trim$(CStr(SafeCell(dPal, rIdx, iStorno)))) = "DA" Then
                Err.Raise vbObjectError + 7343, SRC, "Paleta " & pid & " je stornirana."
            End If
            If UCase$(Trim$(CStr(SafeCell(dPal, rIdx, iPre)))) = "DA" Then
                Err.Raise vbObjectError + 7344, SRC, "Paleta " & pid & " je ve" & ChrW(263) & " preradjena."
            End If
            palRows.Add pid, rIdx
            netoUlaz = netoUlaz + NzD(SafeCell(dPal, rIdx, iNeto))
        End If
    Next v

    If palRows.count = 0 Then
        Err.Raise vbObjectError + 7340, SRC, "Nije izabrana nijedna validna paleta."
    End If

    ' --- upis prerade + stavki + markiranje paleta ---
    Dim preID As String: preID = GetNextID(TBL_PRERADA, COL_PRE_ID, "PRE-")
    Dim brPre As Long: brPre = GenerateBrojPrerade()

    PalAppendRow TBL_PRERADA, _
        Array(COL_PRE_ID, COL_PRE_BROJ, COL_PRE_GODINA, COL_PRE_DATUM, _
              COL_PRE_NETO_ULAZ, COL_PRE_NETO_IZLAZ, COL_PRE_KUTIJE, COL_PRE_KESE, _
              COL_PRE_TEZINA_PALETE, COL_PRE_BRUTO, COL_PRE_AMBALAZA, _
              COL_PRE_TIP_KUTIJE, COL_PRE_TIP_KESE, COL_PRE_TIP_GP, _
              COL_PRE_NAPOMENA, COL_PRE_CREATED, COL_STORNIRANO), _
        Array(preID, brPre, Year(Date), Date, _
              netoUlaz, netoIzlazKg, brojKutija, brojKesa, _
              tezinaPaleteKg, brutoKg, tezinaAmbalazeKg, _
              tipKutije, tipKese, tipGotovogProizvoda, _
              napomena, Now, "")

    Dim k As Variant
    For Each k In palRows.keys
        Dim sid As String: sid = GetNextID(TBL_PRERADA_STAVKA, COL_PRES_ID, "PRS-")
        PalAppendRow TBL_PRERADA_STAVKA, _
            Array(COL_PRES_ID, COL_PRES_PRERADA_ID, COL_PRES_PALETA_ID, _
                  COL_PRES_BROJ_PALETE, COL_PRES_NETO, COL_PRES_CREATED, COL_STORNIRANO), _
            Array(sid, preID, CStr(k), _
                  NzL(SafeCell(dPal, CLng(palRows(k)), iBroj)), _
                  NzD(SafeCell(dPal, CLng(palRows(k)), iNeto)), Now, "")
        RequireUpdateCell TBL_PALETA, CLng(palRows(k)), COL_PAL_PRERADJENO, "Da", SRC
    Next k

    tx.CommitTx
    SavePrerada_TX = preID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.description
    If Not tx Is Nothing Then tx.RollbackTx
    LogErr SRC
    Err.Raise errNum, SRC, errDesc
End Function

' PDF preradnog lista -> <workbook>\Preradni listovi\Prerada_<broj>-<god>.pdf.
Public Function ExportPreradaPDF(ByVal preID As String, _
                                 Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim broj As String, god As String
    Dim ws As Worksheet
    Set ws = FillPreradaSablon(preID, broj, god)
    If ws Is Nothing Then Exit Function

    Dim pdfPath As String
    pdfPath = EnsureDocFolder(PDF_DIR_PRERADA) & "\Prerada_" & broj & "-" & god & ".pdf"

    DocExportPdf ws, pdfPath, openAfter

    ExportPreradaPDF = pdfPath
    Exit Function
EH:
    LogErr "modPaletniList.ExportPreradaPDF"
End Function

' Auto-izlaz preradnog lista (paletni list got. proizvoda) po CFG_PRERADA_PRINT_MODE:
' PDF (default) | PRINT | PREVIEW | OFF. Eksplicitni dvoklik-reprint i dalje ide
' kroz ExportPreradaPDF (uvek PDF).
Public Sub OutputPreradaList(ByVal preID As String)
    On Error GoTo EH
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_PRERADA_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            Dim broj As String, god As String
            Dim ws As Worksheet: Set ws = FillPreradaSablon(preID, broj, god)
            If Not ws Is Nothing Then DocPrintWs ws, mode
        Case "PDF"
            ExportPreradaPDF preID, True
        ' OFF -> bez izlaza
    End Select
    Exit Sub
EH:
    LogErr "modPaletniList.OutputPreradaList"
End Sub

Public Function FindPreradaIDByBroj(ByVal broj As Long, ByVal god As Long) As String
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
    Const SRC As String = "modPaletniList.FillPreradaSablon"
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsurePreradaSablon

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_PRERADA_SABLON)
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

    On Error GoTo EH
    Application.ScreenUpdating = False

    ' Zaglavlje firme (SELLER_* iz configa) se osvezava na SVAKO punjenje.
    ' PreradaSablon je kesiran (EnsurePreradaSablon ga gradi jednom, pa preskace
    ' dok se LAYOUT_VER ne promeni), pa bi podaci o firmi inace ostali zamrznuti
    ' od trenutka prve izgradnje (prazni ako SELLER_* tada nije bio popunjen).
    ' Logo se prvo skida jer ga DocDrawLogo samo dodaje (ne dedupe-uje) -> inace
    ' bi se gomilao na svaki reprint (dvoklik). Isti pristup kao FillPrijemnicaSablon.
    Dim si As Long
    For si = ws.Shapes.count To 1 Step -1
        ws.Shapes(si).Delete
    Next si
    DocSellerHeader ws, 1, 5, 5

    ws.Range("PreBroj").NumberFormat = "@"
    ws.Range("PreBroj").value = brojOut & "/" & godOut
    ws.Range("PreDatum").value = Format$(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_DATUM)), "dd.mm.yyyy")
    ws.Range("PreKutije").value = NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE)))
    ws.Range("PreKese").value = NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_KESE)))
    ws.Range("PreNeto").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ)))
    ws.Range("PreTezinaPalete").NumberFormat = "0.00"   ' Double -> bez E-notacije/tarabi
    ws.Range("PreTezinaPalete").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_TEZINA_PALETE)))
    ws.Range("PreBruto").NumberFormat = "0.00"
    ws.Range("PreBruto").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_BRUTO)))
    ws.Range("PreAmbalaza").NumberFormat = "0.00"
    ws.Range("PreAmbalaza").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_AMBALAZA)))

    ' palete ove prerade -> otkupi (potrebno za obe varijante lista)
    Dim palIDs As Collection: Set palIDs = New Collection
    Dim s As Variant: s = GetTableData(TBL_PRERADA_STAVKA)
    If Not IsEmpty(s) Then
        Dim sPre As Long, sPalID As Long
        sPre = GetColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PRERADA_ID)
        sPalID = GetColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PALETA_ID)
        Dim sr As Long
        For sr = 1 To UBound(s, 1)
            If CStr(SafeCell(s, sr, sPre)) = preID Then
                palIDs.Add CStr(SafeCell(s, sr, sPalID))
            End If
        Next sr
    End If

    Dim o As Variant: o = GetOtkupiZaPalete(palIDs)

    ' Naslov vrste: SAMO tekst iz comboa "Gotov proizvod:" (tip gotovog
    ' proizvoda sa prerade, COL_PRE_TIP_GP). Vrsta/sorta iz izabranih paleta
    ' sveze robe se vise NE citaju iz tblPaleta (zahtev operatera).
    Dim tipGP As String
    tipGP = NzToText(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP)))
    ws.Range("PreVrsta").value = Trim$(tipGP)
    ' Telo lista zavisi od toggle-a "Detaljni prikaz sledljivosti" (sablon je vec
    ' izgradjen u tom rezimu preko EnsurePreradaSablon):
    '   DA -> puna tabela stavki (jedan red po otkupu)
    '   NE -> samo lista sifri kooperanata (zarezom)
    Dim footRow As Long
    If IsPreradaSledljivostDetalj() Then
        footRow = FillPreradaStavkeDetalj(ws, o)
    Else
        footRow = FillPreradaSifreZbirno(ws, o)
    End If
    ws.cells(footRow, 1).value = "Datum stampe: " & Format$(Date, "dd.mm.yyyy")
    ws.cells(footRow, 1).Font.Color = DocColGray()
    ws.cells(footRow + 2, 1).value = "Potpis: ____________________"
    ws.cells(footRow + 2, 1).Font.Color = DocColGray()
    ws.cells(footRow + 2, 4).value = "Pecat: ____________________"
    ws.cells(footRow + 2, 4).Font.Color = DocColGray()

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
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(footRow + 2, 5)).Address
    End With
    Application.PrintCommunication = True
    On Error GoTo 0

    Application.ScreenUpdating = oldScreen
    Set FillPreradaSablon = ws
    Exit Function

EH:
    Application.ScreenUpdating = oldScreen
    LogErr SRC
    Err.Raise Err.Number, SRC, Err.description
End Function

' Kreira/obnavlja PreradaSablon u zajednickom stilu. Verzija layouta je u H1 i
' nosi rezim (D=detaljno / N=zbirno) -- promena toggle-a "Detaljni prikaz
' sledljivosti" (PRERADA_SLEDLJIVOST_DETALJ) rebuild-uje sablon u drugom rasporedu.
Public Sub EnsurePreradaSablon()
    On Error GoTo EH
    Const LAYOUT_VER As String = "5"

    Dim detalj As Boolean: detalj = IsPreradaSledljivostDetalj()
    Dim verKey As String: verKey = LAYOUT_VER & IIf(detalj, "-D", "-N")

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(WS_PRERADA_SABLON)
    On Error GoTo EH
    If Not ws Is Nothing Then
        If CStr(ws.Range("H1").value) = verKey Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If

    Set ws = ThisWorkbook.Sheets.Add
    ws.name = WS_PRERADA_SABLON
    ws.cells.Font.name = "Calibri"
    ws.cells.Font.Size = 10
    ws.columns("A").ColumnWidth = 12
    ws.columns("B").ColumnWidth = 14
    ws.columns("C").ColumnWidth = 14
    ws.columns("D").ColumnWidth = 14
    ws.columns("E").ColumnWidth = 18

    Dim r As Long
    r = DocSellerHeader(ws, 1, 5, 5)
    r = DocTitleBlock(ws, r, 5, "Prerada i pakovanje", "PALETNI LIST GOTOVIH PROIZVODA")

    ' zajednicko zaglavlje (Broj/Datum) -- ostatak layouta zavisi od rezima
    Dim fr As Long: fr = r + 1
    ws.cells(fr, 1).value = "Broj:"
    ws.cells(fr + 1, 1).value = "Datum:"
    ws.cells(fr, 2).name = "PreBroj"
    ws.cells(fr + 1, 2).name = "PreDatum"
    ws.Range(ws.cells(fr, 2), ws.cells(fr + 1, 2)).Font.Bold = True

    If detalj Then
        BuildPreradaSablonDetalj ws, fr
    Else
        BuildPreradaSablonZbirno ws, fr
    End If

    ws.Range("H1").value = verKey
    ws.Range("H1").Font.Color = RGB(255, 255, 255)
    Exit Sub
EH:
    Application.DisplayAlerts = True
    LogErr "modPaletniList.EnsurePreradaSablon"
End Sub

' DA layout (detaljna sledljivost): desni sazetak (tezine/ambalaza) + "Vrsta voca"
' + zaglavlje tabele stavki (Rb/Kooperant/Neto kg/Ambalaza). Imena opsega:
' PreTezinaPalete/PreBruto/PreKutije/PreKese/PreAmbalaza/PreNeto/PreVrsta/PreStavkaStart.
Private Sub BuildPreradaSablonDetalj(ByVal ws As Worksheet, ByVal fr As Long)
    ws.cells(fr, 4).value = "Te" & ChrW(382) & "ina palete (kg):"
    ws.cells(fr + 1, 4).value = "Bruto (kg):"
    ws.cells(fr + 2, 4).value = "Broj kutija:"
    ws.cells(fr + 3, 4).value = "Broj kesa:"
    ws.cells(fr + 4, 4).value = "Te" & ChrW(382) & "ina ambala" & ChrW(382) & "e (kg):"
    ws.cells(fr + 5, 4).value = "Neto (kg):"

    ws.cells(fr, 5).name = "PreTezinaPalete"
    ws.cells(fr + 1, 5).name = "PreBruto"
    ws.cells(fr + 2, 5).name = "PreKutije"
    ws.cells(fr + 3, 5).name = "PreKese"
    ws.cells(fr + 4, 5).name = "PreAmbalaza"
    ws.cells(fr + 5, 5).name = "PreNeto"

    ws.Range(ws.cells(fr, 5), ws.cells(fr + 5, 5)).Font.Bold = True
    ' broj kutija/kesa = celi; tezine = 2 decimale
    ws.Range(ws.cells(fr + 2, 5), ws.cells(fr + 3, 5)).NumberFormat = "0"
    ws.cells(fr, 5).NumberFormat = "#,##0.00"
    ws.cells(fr + 1, 5).NumberFormat = "#,##0.00"
    ws.cells(fr + 4, 5).NumberFormat = "#,##0.00"
    ws.cells(fr + 5, 5).NumberFormat = "#,##0.00"

    ' desni sazetak: uokviri + istakni Neto (poslednji red)
    ws.Range(ws.cells(fr, 4), ws.cells(fr + 5, 5)).BorderAround Weight:=xlThin
    With ws.Range(ws.cells(fr + 5, 4), ws.cells(fr + 5, 5))
        .Interior.Color = DocColHeaderFill()
        .Font.Bold = True
    End With
    With ws.Range(ws.cells(fr + 5, 4), ws.cells(fr + 5, 5)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' vrsta voca kao podnaslov iznad tabele (na preradi je uvek ista vrsta)
    Dim subRow As Long: subRow = fr + 7
    ws.cells(subRow, 1).value = "Vrsta vo" & ChrW(263) & "a:"
    ws.cells(subRow, 1).Font.Color = DocColGray()
    ws.Range(ws.cells(subRow, 2), ws.cells(subRow, 5)).Merge
    ws.cells(subRow, 2).name = "PreVrsta"
    With ws.cells(subRow, 2)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlLeft
    End With
    ws.rows(subRow).RowHeight = 20

    ' tabela stavki (bez kolone Vrsta; Ambalaza preko D:E)
    Dim hdr As Long: hdr = subRow + 1
    ws.cells(hdr, 1).value = "Rb"
    ws.cells(hdr, 2).value = "Kooperant"
    ws.cells(hdr, 3).value = "Neto kg"
    ws.Range(ws.cells(hdr, 4), ws.cells(hdr, 5)).Merge
    ws.cells(hdr, 4).value = "Ambala" & ChrW(382) & "a"
    With ws.Range(ws.cells(hdr, 1), ws.cells(hdr, 5))
        .Font.Bold = True
        .Interior.Color = DocColHeaderFill()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ws.cells(hdr + 1, 1).name = "PreStavkaStart"

    ws.Range(ws.cells(1, 1), ws.cells(hdr, 5)).EntireRow.AutoFit
End Sub

' NE layout (bez detaljne sledljivosti): "Vrsta voca" + centriran, uvecan sazetak
' tezina/ambalaze (preko B:D, vece slovo) + lista sifri kooperanata (PreSifre,
' puni se zarezom). Bez detaljne tabele stavki. Ista imena opsega za sazetak.
Private Sub BuildPreradaSablonZbirno(ByVal ws As Worksheet, ByVal fr As Long)
    ' "Vrsta voca" odmah ispod Broj/Datum
    Dim subRow As Long: subRow = fr + 3
    ws.cells(subRow, 1).value = "Vrsta vo" & ChrW(263) & "a:"
    ws.cells(subRow, 1).Font.Color = DocColGray()
    ws.Range(ws.cells(subRow, 2), ws.cells(subRow, 5)).Merge
    ws.cells(subRow, 2).name = "PreVrsta"
    With ws.cells(subRow, 2)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlLeft
    End With
    ws.rows(subRow).RowHeight = 20

    ' centriran, uvecan sazetak (popunjava prostor osloboden bez detaljne tabele):
    ' label preko B:C (desno), vrednost u D (levo, bold, vece slovo). Okvir oko B:D.
    Dim lbl(0 To 5) As String
    lbl(0) = "Te" & ChrW(382) & "ina palete (kg):"
    lbl(1) = "Bruto (kg):"
    lbl(2) = "Broj kutija:"
    lbl(3) = "Broj kesa:"
    lbl(4) = "Te" & ChrW(382) & "ina ambala" & ChrW(382) & "e (kg):"
    lbl(5) = "Neto (kg):"
    Dim nm(0 To 5) As String
    nm(0) = "PreTezinaPalete"
    nm(1) = "PreBruto"
    nm(2) = "PreKutije"
    nm(3) = "PreKese"
    nm(4) = "PreAmbalaza"
    nm(5) = "PreNeto"

    Dim bt As Long: bt = subRow + 2
    Dim i As Long, rr As Long
    For i = 0 To 5
        rr = bt + i
        ws.Range(ws.cells(rr, 2), ws.cells(rr, 3)).Merge
        ws.cells(rr, 2).value = lbl(i)
        With ws.cells(rr, 2)
            .HorizontalAlignment = xlRight
            .Font.Size = 12
        End With
        ws.cells(rr, 4).name = nm(i)
        With ws.cells(rr, 4)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlLeft
        End With
        ws.rows(rr).RowHeight = 20
    Next i

    ' broj kutija/kesa = celi; tezine = 2 decimale
    ws.Range(ws.cells(bt + 2, 4), ws.cells(bt + 3, 4)).NumberFormat = "0"
    ws.cells(bt, 4).NumberFormat = "#,##0.00"
    ws.cells(bt + 1, 4).NumberFormat = "#,##0.00"
    ws.cells(bt + 4, 4).NumberFormat = "#,##0.00"
    ws.cells(bt + 5, 4).NumberFormat = "#,##0.00"

    Dim bb As Long: bb = bt + 5
    ws.Range(ws.cells(bt, 2), ws.cells(bb, 4)).BorderAround Weight:=xlMedium
    With ws.Range(ws.cells(bb, 2), ws.cells(bb, 4))       ' istakni Neto
        .Interior.Color = DocColHeaderFill()
        .Font.Bold = True
    End With
    With ws.Range(ws.cells(bb, 2), ws.cells(bb, 4)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' lista sifri kooperanata (puni FillPreradaSifreZbirno, zarezom)
    Dim lblRow As Long: lblRow = bb + 2
    ws.cells(lblRow, 1).value = ChrW(352) & "ifre kooperanata:"
    ws.cells(lblRow, 1).Font.Bold = True
    Dim codesRow As Long: codesRow = lblRow + 1
    ws.Range(ws.cells(codesRow, 1), ws.cells(codesRow, 5)).Merge
    ws.cells(codesRow, 1).name = "PreSifre"
    With ws.cells(codesRow, 1)
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Size = 11
    End With
    ws.rows(codesRow).RowHeight = 48

    ' autofit samo zaglavlje (1..fr+1) da rucne visine sazetka/sifri prezive
    ws.Range(ws.cells(1, 1), ws.cells(fr + 1, 5)).EntireRow.AutoFit
End Sub

' DA fill: puna tabela stavki (jedan red po otkupu). o = GetOtkupiZaPalete.
' Cisti stare redove pre upisa. Vraca footRow (prvi red posle tabele, za footer).
Private Function FillPreradaStavkeDetalj(ByVal ws As Worksheet, ByVal o As Variant) As Long
    Dim startRow As Long: startRow = ws.Range("PreStavkaStart").row
    Dim lastRow As Long: lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= startRow Then
        ws.Range(ws.cells(startRow, 4), ws.cells(lastRow, 5)).UnMerge
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, 5)).Clear
    End If

    Dim outR As Long, rb As Long
    outR = startRow: rb = 0
    If Not IsEmpty(o) Then
        Dim k As Long
        For k = 1 To UBound(o, 1)
            rb = rb + 1
            ws.cells(outR, 1).value = rb
            ws.cells(outR, 2).value = CStr(o(k, 1))                   ' Kooperant (sifra)
            ws.cells(outR, 3).value = o(k, 3)                         ' Neto kg
            ws.Range(ws.cells(outR, 4), ws.cells(outR, 5)).Merge
            ws.cells(outR, 4).value = o(k, 4) & " x " & CStr(o(k, 5)) ' Ambalaza: kom x tip (D:E)
            outR = outR + 1
        Next k
    End If

    Dim dataEnd As Long: dataEnd = outR - 1
    If dataEnd >= startRow Then
        With ws.Range(ws.cells(startRow, 1), ws.cells(dataEnd, 5)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        ' uniforman font na svim redovima kooperanata (fix: prvih redova font)
        With ws.Range(ws.cells(startRow, 1), ws.cells(dataEnd, 5))
            .Font.Size = 10
            .Font.Bold = False
        End With
        ws.Range(ws.cells(startRow, 1), ws.cells(dataEnd, 5)).EntireRow.AutoFit
        ws.Range(ws.cells(startRow, 3), ws.cells(dataEnd, 3)).NumberFormat = "#,##0.00"
        ws.Range(ws.cells(startRow, 4), ws.cells(dataEnd, 4)).NumberFormat = "@"
        Dim zr As Long
        For zr = 0 To dataEnd - startRow
            If zr Mod 2 = 1 Then
                ws.Range(ws.cells(startRow + zr, 1), _
                         ws.cells(startRow + zr, 5)).Interior.Color = RGB(217, 225, 242)
            End If
        Next zr
    End If

    Dim footRow As Long: footRow = dataEnd + 2
    If footRow <= startRow Then footRow = startRow + 1
    FillPreradaStavkeDetalj = footRow
End Function

' NE fill: samo lista sifri kooperanata (distinct, zarezom) u PreSifre.
' o = GetOtkupiZaPalete (kol 1 = sifra kooperanta). Vraca footRow.
Private Function FillPreradaSifreZbirno(ByVal ws As Worksheet, ByVal o As Variant) As Long
    Dim codes As String
    If Not IsEmpty(o) Then
        Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
        Dim k As Long, sif As String
        For k = 1 To UBound(o, 1)
            sif = Trim$(CStr(o(k, 1)))
            If Len(sif) > 0 And Not seen.Exists(sif) Then
                seen.Add sif, True
                If Len(codes) > 0 Then codes = codes & ", "
                codes = codes & sif
            End If
        Next k
    End If
    ws.Range("PreSifre").value = codes

    Dim codesRow As Long: codesRow = ws.Range("PreSifre").row
    ws.rows(codesRow).RowHeight = 48    ' re-assert (merged red se ne AutoFit-uje)
    FillPreradaSifreZbirno = codesRow + 2
End Function

' --- tblPaleta helper-i za preradu uklonjeni: SavePrerada_TX cita tblPaleta
'     inline (snapshot dPal) i pise preko RequireUpdateCell. ---

