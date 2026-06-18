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
                                 Optional ByVal preradjeno As String = "") As Variant
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
        res(k, 9) = NzD(SafeCell(d, r, iNeto))
        res(k, 10) = NzD(SafeCell(d, r, iBruto))
        res(k, 11) = CStr(SafeCell(d, r, iStat))
        res(k, 12) = CStr(SafeCell(d, r, iPre))
    Next k

    GetPaleteForGrid = res
    Exit Function
EH:
    LogErr "modPaletniList.GetPaleteForGrid"
End Function

Private Function PreradMatch(ByVal cellVal As String, ByVal filter As String) As Boolean
    Select Case UCase$(Trim$(filter))
        Case "": PreradMatch = True
        Case "DA": PreradMatch = (UCase$(Trim$(cellVal)) = "DA")
        Case "NE": PreradMatch = (UCase$(Trim$(cellVal)) <> "DA")
        Case Else: PreradMatch = True
    End Select
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
        res(k, 4) = NzD(SafeCell(s, r, iNeto))
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
        res(k, 4) = dNe(p)
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
        Err.Raise vbObjectError + 7352, SRC, "Paleta je vec preradjena."
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
' PUBLIC — rucna stampa nepotpunih (otvorenih) paleta.
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
    pdfPath = DataBook.Path & "\Paleta_" & broj & "-" & god & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, _
                           OpenAfterPublish:=openAfter

    ExportPaletniListPDF = pdfPath
    Exit Function
EH:
    LogErr "modPaletniList.ExportPaletniListPDF"
End Function

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
    Const SRC As String = "modPaletniList.FillPaletaSablon"
    Dim oldScreen As Boolean: oldScreen = Application.ScreenUpdating

    EnsurePaletaSablon

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DataBook.Sheets("PaletaSablon")
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
    Set ws = DataBook.Sheets("PaletaSablon")
    On Error GoTo EH
    If Not ws Is Nothing Then
        If CStr(ws.Range("H1").value) = LAYOUT_VER Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If

    Set ws = DataBook.Sheets.Add
    ws.name = "PaletaSablon"
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
    ws.cells(fr + 2, 4).value = "Ambalaza (kg):"
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
    ws.cells(subRow, 1).value = "Vrsta voca:"
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
    ws.cells(hdr, 4).value = "Ambalaza"
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
        Dim pID As Long, pKl As Long
        pID = GetColumnIndex(TBL_PALETA, COL_PAL_ID)
        pKl = GetColumnIndex(TBL_PALETA, COL_PAL_KLASA)
        For r = 1 To UBound(dp, 1)
            If palSet.Exists(CStr(SafeCell(dp, r, pID))) Then
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
    Dim oID As Long, oKoop As Long, oVr As Long, oKol As Long
    Dim oAmb As Long, oTip As Long, oKl As Long, oStorno As Long
    oID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
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
        Dim otkID As String: otkID = CStr(SafeCell(o, r, oID))
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

' Snimi preradu (TX): operater je izabrao palete (kolekcija PaletaID-jeva).
' Snapshot tblPrerada/tblPreradaStavka/tblPaleta; validira (paleta postoji tacno
' jednom, nije stornirana, nije vec preradjena); upisuje preradu + stavke i
' markira palete Preradjeno=Da. Bez MsgBox -> baca gresku (UI je hvata).
' Operater bira otvorene i/ili zatvorene palete. Vraca PreradaID.
Public Function SavePrerada_TX(ByVal paletaIDs As Collection, _
                               ByVal brojKutija As Long, _
                               ByVal brojKesa As Long, _
                               ByVal netoIzlazKg As Double, _
                               Optional ByVal napomena As String = "") As String
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
                Err.Raise vbObjectError + 7344, SRC, "Paleta " & pid & " je vec preradjena."
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
              COL_PRE_NAPOMENA, COL_PRE_CREATED, COL_STORNIRANO), _
        Array(preID, brPre, Year(Date), Date, _
              netoUlaz, netoIzlazKg, brojKutija, brojKesa, _
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

' PDF preradnog lista -> <workbook>\Prerada_<broj>-<god>.pdf.
Public Function ExportPreradaPDF(ByVal preID As String, _
                                 Optional ByVal openAfter As Boolean = True) As String
    On Error GoTo EH
    Dim broj As String, god As String
    Dim ws As Worksheet
    Set ws = FillPreradaSablon(preID, broj, god)
    If ws Is Nothing Then Exit Function

    Dim pdfPath As String
    pdfPath = DataBook.Path & "\Prerada_" & broj & "-" & god & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, _
                           OpenAfterPublish:=openAfter

    ExportPreradaPDF = pdfPath
    Exit Function
EH:
    LogErr "modPaletniList.ExportPreradaPDF"
End Function

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
    Set ws = DataBook.Sheets("PreradaSablon")
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

    ws.Range("PreBroj").NumberFormat = "@"
    ws.Range("PreBroj").value = brojOut & "/" & godOut
    ws.Range("PreDatum").value = Format$(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_DATUM)), "dd.mm.yyyy")
    ws.Range("PreKutije").value = NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE)))
    ws.Range("PreKese").value = NzL(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_KESE)))
    ws.Range("PreNeto").value = NzD(SafeCell(d, hRow, GetColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ)))

    Dim startRow As Long: startRow = ws.Range("PreStavkaStart").row
    Dim lastRow As Long: lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= startRow Then
        ws.Range(ws.cells(startRow, 4), ws.cells(lastRow, 5)).UnMerge
        ws.Range(ws.cells(startRow, 1), ws.cells(lastRow, 5)).Clear
    End If

    ' prerada -> prerađene palete -> otkupi: jedan red po OTKUPU
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
    If IsEmpty(o) Then
        ws.Range("PreVrsta").value = ""
    Else
        ws.Range("PreVrsta").value = CStr(o(1, 2))
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

' Kreira/obnavlja PreradaSablon u zajednickom stilu. Verzija layouta je u H1.
Public Sub EnsurePreradaSablon()
    On Error GoTo EH
    Const LAYOUT_VER As String = "3"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DataBook.Sheets("PreradaSablon")
    On Error GoTo EH
    If Not ws Is Nothing Then
        If CStr(ws.Range("H1").value) = LAYOUT_VER Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Set ws = Nothing
    End If

    Set ws = DataBook.Sheets.Add
    ws.name = "PreradaSablon"
    ws.cells.Font.name = "Calibri"
    ws.cells.Font.Size = 10
    ws.columns("A").ColumnWidth = 12
    ws.columns("B").ColumnWidth = 14
    ws.columns("C").ColumnWidth = 14
    ws.columns("D").ColumnWidth = 14
    ws.columns("E").ColumnWidth = 18

    Dim r As Long
    r = DocSellerHeader(ws, 1, 5, 5)
    r = DocTitleBlock(ws, r, 5, "Prerada i pakovanje", "PRERADNI LIST")

    Dim fr As Long: fr = r + 1
    ws.cells(fr, 1).value = "Broj:"
    ws.cells(fr + 1, 1).value = "Datum:"
    ws.cells(fr, 4).value = "Broj kutija:"
    ws.cells(fr + 1, 4).value = "Broj kesa:"
    ws.cells(fr + 2, 4).value = "Neto (kg):"

    ws.cells(fr, 2).name = "PreBroj"
    ws.cells(fr + 1, 2).name = "PreDatum"
    ws.cells(fr, 5).name = "PreKutije"
    ws.cells(fr + 1, 5).name = "PreKese"
    ws.cells(fr + 2, 5).name = "PreNeto"

    ws.Range(ws.cells(fr, 2), ws.cells(fr + 1, 2)).Font.Bold = True
    ws.Range(ws.cells(fr, 5), ws.cells(fr + 2, 5)).Font.Bold = True
    ws.Range(ws.cells(fr, 5), ws.cells(fr + 1, 5)).NumberFormat = "0"
    ws.cells(fr + 2, 5).NumberFormat = "#,##0.00"

    ' desni sazetak: uokviri + istakni Neto
    ws.Range(ws.cells(fr, 4), ws.cells(fr + 2, 5)).BorderAround Weight:=xlThin
    With ws.Range(ws.cells(fr + 2, 4), ws.cells(fr + 2, 5))
        .Interior.Color = DocColHeaderFill()
        .Font.Bold = True
    End With
    With ws.Range(ws.cells(fr + 2, 4), ws.cells(fr + 2, 5)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' vrsta voca kao podnaslov iznad tabele (na preradi je uvek ista vrsta)
    Dim subRow As Long: subRow = fr + 4
    ws.cells(subRow, 1).value = "Vrsta voca:"
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
    ws.cells(hdr, 4).value = "Ambalaza"
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
    ws.Range("H1").value = LAYOUT_VER
    ws.Range("H1").Font.Color = RGB(255, 255, 255)
    Exit Sub
EH:
    Application.DisplayAlerts = True
    LogErr "modPaletniList.EnsurePreradaSablon"
End Sub

' --- tblPaleta helper-i za preradu uklonjeni: SavePrerada_TX cita tblPaleta
'     inline (snapshot dPal) i pise preko RequireUpdateCell. ---
