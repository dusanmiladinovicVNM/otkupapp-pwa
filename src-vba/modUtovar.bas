Attribute VB_Name = "modUtovar"
Option Explicit

' ============================================================
' modUtovar (krug 5 revizije #248) -- UTOVARNA LISTA: dokument
' FIZICKE isporuke gotove robe.
'
' Grain: prerada je proizvodni lot (koliko je PROIZVEDENO); utovarna
' stavka (prerada + kg) je prodajna jedinica -- parcijalna prodaja
' (500 kg od 2.000) je legalna, "na stanju" = NetoIzlazKg - SUM
' aktivnih utovarenih kg. Prerada se NIKAD ne zakljucava fakturom.
'
' v1 ugovor: JEDAN utovar = JEDNA GP faktura, prave se u ISTOJ
' transakciji (CreateUtovarSaFakturom_TX); DatumUtovara je datum
' izrade i on ide na SEF kao datum isporuke. Poseban ekran utovara i
' stampani obrazac utovarne liste su sledeci korak (dokument vec
' postoji podatkovno i stampa ce citati ove tabele).
'
' Storno simetrija (modStorno): storno FAKTURE oslobadja utovar
' (roba ostaje utovarena); storno UTOVARA vraca robu na stanje i
' dozvoljen je samo nad nefakturisanim utovarom.
' ============================================================

' Sledeci broj utovara -- maxN+1 unutar date godine (isti obrazac kao
' GenerateBrojPalete). godina=0 -> tekuca. Parametar postoji zbog
' migracije starih GP faktura: utovar iz fakture 2025. mora nositi
' broj i godinu 2025, ne "3/2026 sa datumom 14.08.2025" (revizija #6).
Public Function GenerateBrojUtovara(Optional ByVal godina As Long = 0) As Long
    Const SRC As String = "modUtovar.GenerateBrojUtovara"
    Dim d As Variant, i As Long, maxN As Long
    Dim cBr As Long, cGod As Long
    If godina = 0 Then godina = Year(Date)
    d = GetTableData(TBL_UTOVAR)
    If IsArray(d) Then
        cBr = RequireColumnIndex(TBL_UTOVAR, COL_UT_BROJ, SRC)
        cGod = RequireColumnIndex(TBL_UTOVAR, COL_UT_GODINA, SRC)
        For i = 1 To UBound(d, 1)
            If IsNumeric(d(i, cGod)) And IsNumeric(d(i, cBr)) Then
                If CLng(d(i, cGod)) = godina Then
                    If CLng(d(i, cBr)) > maxN Then maxN = CLng(d(i, cBr))
                End If
            End If
        Next i
    End If
    GenerateBrojUtovara = maxN + 1
End Function

' Mapa AKTIVNO utovarenih kg po preradi -- jedan prolaz (S5 pravilo).
' Aktivna stavka = stavka nije stornirana I njen utovar nije storniran.
' Meko: sveska pre nadogradnje nema tabele -> prazna mapa (sve na
' stanju). Public: dele je grid, writer i storno kapija -- JEDNO
' pravilo, ne tri kopije (pouka kruga 3).
Public Function UtovarenoPoPreradi() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set UtovarenoPoPreradi = d

    If GetTable(TBL_UTOVAR_STAVKE) Is Nothing Then Exit Function
    If GetTable(TBL_UTOVAR) Is Nothing Then Exit Function

    ' Aktivni utovari (dict utovarID -> True).
    Dim ut As Variant, i As Long
    Dim aktivni As Object: Set aktivni = CreateObject("Scripting.Dictionary")
    aktivni.CompareMode = vbTextCompare
    ut = GetTableData(TBL_UTOVAR)
    If IsArray(ut) Then
        ut = ExcludeStornirano(ut, TBL_UTOVAR)
        If IsArray(ut) Then
            Dim cUtId As Long
            cUtId = RequireColumnIndex(TBL_UTOVAR, COL_UT_ID, "modUtovar.UtovarenoPoPreradi")
            For i = 1 To UBound(ut, 1)
                aktivni(Trim$(CStr(nz(ut(i, cUtId))))) = True
            Next i
        End If
    End If

    Dim s As Variant
    s = GetTableData(TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then Exit Function
    s = ExcludeStornirano(s, TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then Exit Function

    Dim cUt As Long, cPre As Long, cKol As Long, k As String
    cUt = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_UTOVAR_ID, "modUtovar.UtovarenoPoPreradi")
    cPre = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_PRERADA_ID, "modUtovar.UtovarenoPoPreradi")
    cKol = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_KOLICINA, "modUtovar.UtovarenoPoPreradi")
    For i = 1 To UBound(s, 1)
        If aktivni.Exists(Trim$(CStr(nz(s(i, cUt))))) Then
            k = Trim$(CStr(nz(s(i, cPre))))
            If Len(k) > 0 And IsNumeric(s(i, cKol)) Then
                If d.Exists(k) Then
                    d(k) = CDbl(d(k)) + CDbl(s(i, cKol))
                Else
                    d.Add k, CDbl(s(i, cKol))
                End If
            End If
        End If
    Next i
End Function

' Upis jednog prevoz polja: prazno ne dira, "-" brise, ostalo upisuje.
Private Sub UtUpisiPolje(ByVal rowUt As Long, ByVal kolona As String, _
                         ByVal vrednost As String, ByVal SRC As String)
    Dim v As String
    v = Trim$(vrednost)
    If Len(v) = 0 Then Exit Sub
    If v = "-" Then v = ""
    RequireUpdateCell TBL_UTOVAR, rowUt, kolona, v, SRC
End Sub

' Tekst polja utovara po imenu kolone -- meko (starija sveska bez
' prevoz kolona daje prazno, ne pad).
Private Function UtPolje(ByRef ut As Variant, ByVal rowUt As Long, _
                         ByVal kolona As String) As String
    Dim c As Long
    c = GetColumnIndex(TBL_UTOVAR, kolona)
    If c > 0 Then UtPolje = Trim$(CStr(nz(ut(rowUt, c))))
End Function

' Auto-ucenje sifarnika prevoznika (smoke 5d): posle "Sacuvaj prevoz"
' kombinacija prevoznik+vozac ulazi u tblPrevoznici da sledeci unos
' bude izbor iz liste, ne kucanje. Nova kombinacija = nov red; poznata
' kombinacija sa novom registracijom = azuriranje registracije.
' MEKO: ucenje ne sme da obori cuvanje prevoza -- greska se loguje.
Public Sub UpsertPrevoznikVozac(ByVal naziv As String, _
                                ByVal vozac As String, _
                                ByVal registracija As String)
    Const SRC As String = "modUtovar.UpsertPrevoznikVozac"
    On Error GoTo EH
    naziv = Trim$(naziv): vozac = Trim$(vozac): registracija = Trim$(registracija)
    If Len(naziv) = 0 And Len(vozac) = 0 Then Exit Sub

    Dim lo As ListObject
    Set lo = GetTable(TBL_PREVOZNICI)
    If lo Is Nothing Then Exit Sub    ' sveska pre nadogradnje seme

    Dim d As Variant, i As Long, cNaz As Long, cVoz As Long, cReg As Long
    cNaz = RequireColumnIndex(TBL_PREVOZNICI, COL_PRV_NAZIV, SRC)
    cVoz = RequireColumnIndex(TBL_PREVOZNICI, COL_PRV_VOZAC, SRC)
    cReg = RequireColumnIndex(TBL_PREVOZNICI, COL_PRV_REG, SRC)
    d = GetTableData(TBL_PREVOZNICI)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If StrComp(Trim$(CStr(nz(d(i, cNaz)))), naziv, vbTextCompare) = 0 And _
               StrComp(Trim$(CStr(nz(d(i, cVoz)))), vozac, vbTextCompare) = 0 Then
                If Len(registracija) > 0 Then _
                    If StrComp(Trim$(CStr(nz(d(i, cReg)))), registracija, vbTextCompare) <> 0 Then _
                        RequireUpdateCell TBL_PREVOZNICI, i, COL_PRV_REG, registracija, SRC
                Exit Sub
            End If
        Next i
    End If

    ' Nova kombinacija -- upis PO IMENU kolone (drift-safe transport
    ' kroz pozicioni AppendRow, isti obrazac kao modMalina).
    Dim rowData() As Variant
    ReDim rowData(1 To lo.ListColumns.count)
    rowData(RequireColumnIndex(TBL_PREVOZNICI, COL_PRV_ID, SRC)) = _
        GetNextID(TBL_PREVOZNICI, COL_PRV_ID, "PRV-")
    rowData(cNaz) = naziv
    rowData(cVoz) = vozac
    rowData(cReg) = registracija
    rowData(RequireColumnIndex(TBL_PREVOZNICI, COL_PRV_AKTIVAN, SRC)) = "Aktivan"
    AppendRow TBL_PREVOZNICI, rowData
    Exit Sub
EH:
    LogErr SRC
End Sub

' Kapacitet jednog pakovanja u NETO kg robe (bez tezine ambalaze --
' ona ulazi samo u bruto). Izvor po prioritetu: POSTAVKA iz
' Podesavanja (operaterov contract: kapaciteti se resavaju tamo) ->
' izvedeno iz samog lota (neto/broj) -> 0 (nepoznato). Sanity opseg
' 0.1-1000 kg stiti od datumski formatirane config celije (ista mina
' kao rok trajanja).
Private Function KapacitetPakovanja(ByVal cfgKljuc As String, _
                                    ByVal lotNeto As Double, _
                                    ByVal lotBroj As Double) As Double
    Const EPS As Double = 0.0001
    Dim v As Variant
    v = GetConfigValue(cfgKljuc)
    If IsNumeric(v) Then
        If CDbl(v) >= 0.1 And CDbl(v) <= 1000 Then
            KapacitetPakovanja = CDbl(v)
            Exit Function
        End If
    End If
    If lotBroj > 0 And lotNeto > EPS Then _
        KapacitetPakovanja = lotNeto / lotBroj
End Function

' Broj pakovanja za datu kolicinu. Kapacitet: Podesavanja (cfgKljuc)
' pa lot fallback -- v. KapacitetPakovanja.
' samoTacno=True (dokument): broj SAMO kad je kg celobrojan umnozak
' kapaciteta, inace Empty -- utovarna lista ne nosi aproksimacije.
' samoTacno=False (grid): broj CELIH pakovanja na stanju (Fix) -- 733
' kg uz 10 kg/kutiji = 73 cele kutije; bez kapaciteta vraca 0.
Public Function PakovanjaZaKg(ByVal kg As Double, ByVal lotNeto As Double, _
                              ByVal lotBroj As Double, _
                              ByVal samoTacno As Boolean, _
                              ByVal cfgKljuc As String) As Variant
    Const EPS As Double = 0.0001
    Dim poKom As Double, n As Double
    If samoTacno Then PakovanjaZaKg = Empty Else PakovanjaZaKg = 0&
    poKom = KapacitetPakovanja(cfgKljuc, lotNeto, lotBroj)
    If poKom <= EPS Then Exit Function
    n = kg / poKom
    If samoTacno Then
        If Abs(n - Round(n)) <= 0.001 And Round(n) > 0 Then _
            PakovanjaZaKg = CLng(Round(n))
    Else
        PakovanjaZaKg = CLng(Fix(n + EPS))
    End If
End Function

' Prikazni tekst pakovanja iz dva (opciona) broja: "50 kut. / 50 kesa",
' "50 kut.", "50 kesa" ili "" -- Empty/0 se ne prikazuje.
Private Function PakTekst(ByVal k As Variant, ByVal s As Variant) As String
    Dim kT As String, sT As String
    If Not IsEmpty(k) Then If CLng(k) > 0 Then kT = CStr(CLng(k)) & " kut."
    If Not IsEmpty(s) Then If CLng(s) > 0 Then sT = CStr(CLng(s)) & " kesa"
    If Len(kT) > 0 And Len(sT) > 0 Then
        PakTekst = kT & " / " & sT
    Else
        PakTekst = kT & sT
    End If
End Function

' Broj pakovanja koji je STVARNO usao u kamion (revizija #6 t.4 +
' #12 smoke): cela kolicina -> puni brojevi iz prerade; parcijala ->
' po vrsti pakovanja NEZAVISNO, iz kapaciteta izvedenog iz lota
' (500 kg / 10 kg = 50 kutija i 50 kesa), samo kad je kolicina
' celobrojan umnozak; inace PRAZNO -- transportni dokument radije
' bez podatka nego sa pogresnim brojem.
Private Sub UtsPakovanja(ByVal kol As Double, ByVal neto As Double, _
                         ByVal kut As Double, ByVal kes As Double, _
                         ByRef outKut As Variant, ByRef outKes As Variant)
    Const EPS As Double = 0.0001
    outKut = Empty: outKes = Empty
    If kol >= neto - EPS Then
        If kut > 0 Then outKut = CLng(kut)
        If kes > 0 Then outKes = CLng(kes)
        Exit Sub
    End If
    outKut = PakovanjaZaKg(kol, neto, kut, True, CFG_GP_KG_KUTIJA)
    outKes = PakovanjaZaKg(kol, neto, kes, True, CFG_GP_KG_KESA)
End Sub

' Broj AKTIVNIH (nestorniranih, neosirocenih) faktura-stavki koje
' tvrde ovaj utovar. KANONSKO pravilo (revizija #10 B2) koje dele
' CreateFakturaIzUtovara (re-fakturisanje kontradiktornog utovara je
' dupla prodaja) i modStorno.StornoUtovar (storno utovara cija roba
' je na aktivnoj fakturi = dupla zaliha) -- header marker sam nije
' dovoljan dokaz "nefakturisanosti".
Public Function AktivnihFstZaUtovar(ByVal utovarID As String) As Long
    Const SRC As String = "modUtovar.AktivnihFstZaUtovar"
    Dim fs As Variant, i As Long
    Dim cUt As Long, cSt As Long, cOs As Long
    fs = GetTableData(TBL_FAKTURA_STAVKE)
    If Not IsArray(fs) Then Exit Function
    cUt = GetColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_UTOVAR_ID)
    If cUt = 0 Then Exit Function    ' sveska pre GP nadogradnje
    cSt = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_STORNIRANO, SRC)
    cOs = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_OSIROCENO_OD, SRC)
    For i = 1 To UBound(fs, 1)
        If Trim$(CStr(nz(fs(i, cUt)))) = Trim$(utovarID) Then
            If UCase$(Trim$(CStr(nz(fs(i, cSt))))) <> "DA" _
               And Len(Trim$(CStr(nz(fs(i, cOs))))) = 0 Then
                AktivnihFstZaUtovar = AktivnihFstZaUtovar + 1
            End If
        End If
    Next i
End Function

' Aktivno utovareno kg jedne prerade (kapija storna prerade).
Public Function UtovarenoKgPrerade(ByVal preradaID As String) As Double
    Dim d As Object: Set d = UtovarenoPoPreradi()
    If d.Exists(Trim$(preradaID)) Then _
        UtovarenoKgPrerade = CDbl(d(Trim$(preradaID)))
End Function

' ============================================================
' READ MODEL liste UTOVARI (ekran Fakturisanje) -- 1..7:
'   1 UtovarID (prazan kod duplog -- IdIliPrazno guard)
'   2 Broj ("1/2026")   3 Datum   4 Kupac (naziv)
'   5 Roba ("51/2026" / "N pre.")   6 Ukupno kg   7 Broj fakture ("")
'   8 Prevoznik   9 Registracija (krug 5d -- vidljivost prevoza)
' Stornirani utovari se ne listaju; stavke storniranog utovara ne.
' ============================================================
Public Function GetUtovariForGrid() As Variant
    Const SRC As String = "modUtovar.GetUtovariForGrid"
    On Error GoTo EH

    If GetTable(TBL_UTOVAR) Is Nothing Then Exit Function
    Dim ut As Variant
    ut = GetTableData(TBL_UTOVAR)
    If Not IsArray(ut) Then Exit Function
    ut = ExcludeStornirano(ut, TBL_UTOVAR)
    If Not IsArray(ut) Then Exit Function

    Dim cId As Long, cBr As Long, cGod As Long, cDat As Long
    Dim cKup As Long, cFid As Long
    cId = RequireColumnIndex(TBL_UTOVAR, COL_UT_ID, SRC)
    cBr = RequireColumnIndex(TBL_UTOVAR, COL_UT_BROJ, SRC)
    cGod = RequireColumnIndex(TBL_UTOVAR, COL_UT_GODINA, SRC)
    cDat = RequireColumnIndex(TBL_UTOVAR, COL_UT_DATUM, SRC)
    cKup = RequireColumnIndex(TBL_UTOVAR, COL_UT_KUPAC, SRC)
    cFid = RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURA_ID, SRC)

    Dim brojac As Object: Set brojac = modFaktura.BrojacIdova(TBL_UTOVAR, COL_UT_ID)
    Dim kupMapa As Object: Set kupMapa = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
    Dim fakBroj As Object: Set fakBroj = CreateObject("Scripting.Dictionary")
    fakBroj.CompareMode = vbTextCompare
    Dim fd As Variant, j As Long
    fd = GetTableData(TBL_FAKTURE)
    If IsArray(fd) Then fd = ExcludeStornirano(fd, TBL_FAKTURE)
    If IsArray(fd) Then
        ' cFkId, NE cFId: u ovoj proceduri vec zivi cFid (utovarova
        ' kolona FakturaID) -- VBA je case-insensitive pa bi cFId bio
        ' "Duplicate declaration" (ista mina kao Const SRC / Dim src).
        Dim cFkId As Long, cFkBr As Long
        cFkId = GetColumnIndex(TBL_FAKTURE, COL_FAK_ID)
        cFkBr = GetColumnIndex(TBL_FAKTURE, COL_FAK_BROJ)
        For j = 1 To UBound(fd, 1)
            If Not fakBroj.Exists(Trim$(CStr(nz(fd(j, cFkId))))) Then _
                fakBroj.Add Trim$(CStr(nz(fd(j, cFkId)))), Trim$(CStr(nz(fd(j, cFkBr))))
        Next j
    End If

    ' Roba i kg po utovaru -- jedan prolaz kroz aktivne stavke.
    Dim roba As Object: Set roba = CreateObject("Scripting.Dictionary")
    roba.CompareMode = vbTextCompare
    Dim kg As Object: Set kg = CreateObject("Scripting.Dictionary")
    kg.CompareMode = vbTextCompare
    Dim cnt As Object: Set cnt = CreateObject("Scripting.Dictionary")
    cnt.CompareMode = vbTextCompare
    Dim s As Variant, k As String
    s = GetTableData(TBL_UTOVAR_STAVKE)
    If IsArray(s) Then s = ExcludeStornirano(s, TBL_UTOVAR_STAVKE)
    If IsArray(s) Then
        Dim cSUt As Long, cSBr As Long, cSKol As Long
        cSUt = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_UTOVAR_ID, SRC)
        cSBr = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_BROJ_PRERADE, SRC)
        cSKol = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_KOLICINA, SRC)
        For j = 1 To UBound(s, 1)
            k = Trim$(CStr(nz(s(j, cSUt))))
            If Len(k) > 0 Then
                roba(k) = Trim$(CStr(nz(s(j, cSBr))))
                If cnt.Exists(k) Then cnt(k) = CLng(cnt(k)) + 1 Else cnt.Add k, 1
                If IsNumeric(s(j, cSKol)) Then
                    If kg.Exists(k) Then
                        kg(k) = CDbl(kg(k)) + CDbl(s(j, cSKol))
                    Else
                        kg.Add k, CDbl(s(j, cSKol))
                    End If
                End If
            End If
        Next j
    End If

    Dim outA() As Variant, i As Long, n As Long, utID As String, fid As String
    ReDim outA(1 To UBound(ut, 1), 1 To 9)
    For i = 1 To UBound(ut, 1)
        utID = Trim$(CStr(nz(ut(i, cId))))
        n = n + 1
        outA(n, 1) = modFaktura.IdIliPrazno(brojac, utID)
        outA(n, 2) = Trim$(CStr(nz(ut(i, cBr)))) & "/" & Trim$(CStr(nz(ut(i, cGod))))
        outA(n, 3) = ut(i, cDat)
        If kupMapa.Exists(Trim$(CStr(nz(ut(i, cKup))))) Then
            outA(n, 4) = Trim$(CStr(kupMapa(Trim$(CStr(nz(ut(i, cKup)))))))
        Else
            outA(n, 4) = Trim$(CStr(nz(ut(i, cKup))))
        End If
        If cnt.Exists(utID) Then
            If CLng(cnt(utID)) = 1 Then
                outA(n, 5) = CStr(roba(utID))
            Else
                outA(n, 5) = CStr(CLng(cnt(utID))) & " pre."
            End If
        Else
            outA(n, 5) = ""
        End If
        If kg.Exists(utID) Then outA(n, 6) = CDbl(kg(utID)) Else outA(n, 6) = 0#
        fid = Trim$(CStr(nz(ut(i, cFid))))
        If fakBroj.Exists(fid) Then
            outA(n, 7) = CStr(fakBroj(fid))
        Else
            outA(n, 7) = ""
        End If
        outA(n, 8) = UtPolje(ut, i, COL_UT_PREVOZNIK)
        outA(n, 9) = UtPolje(ut, i, COL_UT_REGISTRACIJA)
    Next i
    If n = 0 Then Exit Function

    Dim res() As Variant, r As Long, c As Long
    ReDim res(1 To n, 1 To 9)
    For r = 1 To n
        For c = 1 To 9
            res(r, c) = outA(r, c)
        Next c
    Next r
    GetUtovariForGrid = res
    Exit Function

EH:
    LogErr SRC
End Function

' ============================================================
' PODACI PREVOZA (krug 5d): prevoznik/vozac/registracija/plomba/
' temperaturni rezim/mesto istovara/PO broj/napomena -- unose se na
' listi Utovari (radnja "Sacuvaj prevoz") i idu na stampani obrazac.
' Upis PO IMENU (kolone su dopuna na kraj tabele); kapije: utovar
' postoji tacno jednom i nije storniran. Semantika polja: PRAZNO ne
' dira postojecu vrednost (operater dopunjava npr. samo plombu),
' crtica "-" BRISE polje (ispravka pogresnog unosa).
' ============================================================
Public Function UpdateUtovarPrevoz_TX(ByVal utovarID As String, _
        ByVal prevoznik As String, ByVal vozac As String, _
        ByVal registracija As String, ByVal plomba As String, _
        ByVal tempRezim As String, ByVal mestoIstovara As String, _
        ByVal poBroj As String, ByVal napomena As String, _
        Optional ByVal datumUtovara As String = "", _
        Optional ByVal vremeUtovara As String = "") As Boolean
    Const SRC As String = "modUtovar.UpdateUtovarPrevoz_TX"
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_UTOVAR

    Dim rows As Collection
    Set rows = FindRows(TBL_UTOVAR, COL_UT_ID, Trim$(utovarID))
    If rows Is Nothing Then
        Err.Raise vbObjectError + 1758, SRC, "Utovar ne postoji: " & utovarID
    ElseIf rows.count <> 1 Then
        Err.Raise vbObjectError + 1758, SRC, _
                  "Utovar ne postoji jednoznacno: " & utovarID
    End If
    Dim rowUt As Long: rowUt = CLng(rows(1))
    Dim d As Variant: d = GetTableData(TBL_UTOVAR)
    If UCase$(Trim$(CStr(nz(d(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_STORNIRANO, SRC)))))) = "DA" Then
        Err.Raise vbObjectError + 1759, SRC, "Utovar je storniran: " & utovarID
    End If

    UtUpisiPolje rowUt, COL_UT_PREVOZNIK, prevoznik, SRC
    UtUpisiPolje rowUt, COL_UT_VOZAC, vozac, SRC
    UtUpisiPolje rowUt, COL_UT_REGISTRACIJA, registracija, SRC
    UtUpisiPolje rowUt, COL_UT_PLOMBA, plomba, SRC
    UtUpisiPolje rowUt, COL_UT_TEMP_REZIM, tempRezim, SRC
    UtUpisiPolje rowUt, COL_UT_MESTO_ISTOVARA, mestoIstovara, SRC
    UtUpisiPolje rowUt, COL_UT_PO_BROJ, poBroj, SRC
    UtUpisiPolje rowUt, COL_UT_NAPOMENA, napomena, SRC

    ' Datum/vreme utovara su EDITABILNI pre SEF slanja (revizija #6
    ' P1): default nastaje pri izradi (Date/Now), ali stvaran utovar
    ' moze biti ranije/kasnije od klika. Prazno = ne diraj; datum se
    ' NE moze obrisati ("-" nije dozvoljen -- dokument mora imati
    ' datum, on je SEF datum isporuke).
    ' LOCK posle stvarnog SEF slanja (revizija #7 B1): DeliveryDate je
    ' PORESKI podatak -- kad je faktura utovara otisla (ili pokusala da
    ' ode) spolja, lokalna promena datuma bi se razisla sa SEF-om.
    ' Menjanje je dozvoljeno samo u stanjima za koja NIJE dokazano
    ' spoljno slanje: LOCAL_FINALIZED / SEF_READY / SEF_TECH_FAILED
    ' (+ prazno kod starih faktura). Transportni tekst (prevoznik,
    ' plomba...) ostaje editabilan -- nije poreski podatak.
    datumUtovara = Trim$(datumUtovara)
    vremeUtovara = Trim$(vremeUtovara)
    If (Len(datumUtovara) > 0 And datumUtovara <> "-") _
       Or (Len(vremeUtovara) > 0 And vremeUtovara <> "-") Then
        Dim lockFid As String
        lockFid = Trim$(CStr(nz(d(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURA_ID, SRC)))))
        If Len(lockFid) > 0 Then
            Dim wfState As String
            wfState = Trim$(GetFakturaSEFWorkflowState(lockFid))
            If Len(wfState) > 0 _
               And wfState <> WF_LOCAL_FINALIZED _
               And wfState <> WF_SEF_READY _
               And wfState <> WF_SEF_TECH_FAILED Then
                Err.Raise vbObjectError + 1775, SRC, _
                          "Datum/vreme utovara je zakljucan: faktura je u SEF stanju " & _
                          wfState & " -- promena bi se razisla sa poslatim dokumentom."
            End If
        End If
    End If
    If Len(datumUtovara) > 0 And datumUtovara <> "-" Then
        ' Srpski format prikazuje datum sa zavrsnom tackom ("2.6.2026.")
        ' a IsDate bas nju ne prima -- skini je pre provere (operater
        ' unosi/kopira upravo taj oblik).
        If Right$(datumUtovara, 1) = "." Then _
            datumUtovara = Left$(datumUtovara, Len(datumUtovara) - 1)
        If Not IsDate(datumUtovara) Then
            Err.Raise vbObjectError + 1774, SRC, _
                      "Datum utovara nije validan datum: " & datumUtovara
        End If
        RequireUpdateCell TBL_UTOVAR, rowUt, COL_UT_DATUM, _
                          CDate(datumUtovara), SRC
    End If
    vremeUtovara = Trim$(vremeUtovara)
    If Len(vremeUtovara) > 0 And vremeUtovara <> "-" Then
        If Not IsDate(vremeUtovara) Then
            Err.Raise vbObjectError + 1774, SRC, _
                      "Vreme utovara nije validno (hh:mm): " & vremeUtovara
        End If
        RequireUpdateCell TBL_UTOVAR, rowUt, COL_UT_VREME, _
                          Format$(CDate(vremeUtovara), "hh:mm"), SRC
    End If

    tx.CommitTx
    UpdateUtovarPrevoz_TX = True
    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE
    LogErr SRC
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0
    UpdateUtovarPrevoz_TX = False
End Function

' ============================================================
' STAMPA UTOVARNE LISTE -- dokument koji ide sa robom. Rezim iz
' CFG_UTOVAR_PRINT_MODE (default PRINT; OFF u test fixture-u).
' ============================================================
Public Sub PrintUtovar(ByVal utovarID As String)
    Const SRC As String = "modUtovar.PrintUtovar"
    On Error GoTo EH

    Dim rows As Collection
    Set rows = FindRows(TBL_UTOVAR, COL_UT_ID, Trim$(utovarID))
    If rows Is Nothing Then Exit Sub
    If rows.count <> 1 Then
        Err.Raise vbObjectError + 1757, SRC, _
                  "Utovar ne postoji jednoznacno: " & utovarID
    End If

    Dim ut As Variant, rowUt As Long
    ut = GetTableData(TBL_UTOVAR)
    rowUt = CLng(rows(1))
    Dim broj As String, datum As Variant, kupacID As String, kupacNaziv As String
    broj = Trim$(CStr(nz(ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_BROJ, SRC))))) & _
           "/" & Trim$(CStr(nz(ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_GODINA, SRC)))))
    datum = ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_DATUM, SRC))
    kupacID = Trim$(CStr(nz(ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_KUPAC, SRC)))))
    kupacNaziv = Trim$(CStr(nz(LookupValue(TBL_KUPCI, COL_KUP_ID, kupacID, COL_KUP_NAZIV))))
    If kupacNaziv = "" Then kupacNaziv = kupacID

    ' Krug 5d: header podaci profesionalnog obrasca (meko -- starija
    ' sveska bez prevoz kolona daje prazna polja, ne pad).
    Dim vreme As String, mestoIst As String, poBroj As String
    Dim napomena As String, fakBroj As String, fid As String
    Dim prevoz(0 To 4) As String
    ' Vreme moze biti i PRAVA vremenska vrednost: kolona General format
    ' pretvori upisano "11:30" u serijski broj, pa je CStr davao
    ' "0,479166..." na obrascu (smoke 5d). Datum/broj -> "hh:mm".
    Dim cVre As Long
    cVre = GetColumnIndex(TBL_UTOVAR, COL_UT_VREME)
    If cVre > 0 Then
        If IsDate(ut(rowUt, cVre)) Or IsNumeric(ut(rowUt, cVre)) Then
            If Len(Trim$(CStr(nz(ut(rowUt, cVre))))) > 0 Then _
                vreme = Format$(CDate(ut(rowUt, cVre)), "hh:mm")
        Else
            vreme = UtPolje(ut, rowUt, COL_UT_VREME)
        End If
    End If
    mestoIst = UtPolje(ut, rowUt, COL_UT_MESTO_ISTOVARA)
    poBroj = UtPolje(ut, rowUt, COL_UT_PO_BROJ)
    napomena = UtPolje(ut, rowUt, COL_UT_NAPOMENA)
    prevoz(0) = UtPolje(ut, rowUt, COL_UT_PREVOZNIK)
    prevoz(1) = UtPolje(ut, rowUt, COL_UT_VOZAC)
    prevoz(2) = UtPolje(ut, rowUt, COL_UT_REGISTRACIJA)
    prevoz(3) = UtPolje(ut, rowUt, COL_UT_PLOMBA)
    prevoz(4) = UtPolje(ut, rowUt, COL_UT_TEMP_REZIM)
    fid = UtPolje(ut, rowUt, COL_UT_FAKTURA_ID)
    If Len(fid) > 0 Then _
        fakBroj = Trim$(CStr(nz(LookupValue(TBL_FAKTURE, COL_FAK_ID, fid, COL_FAK_BROJ))))

    ' Rok trajanja = datum prerade + N meseci (Podesavanja; default 24
    ' za smrznuto). IZVEDEN podatak, jasno dokumentovan -- posebna
    ' kolona po preradi je buduci korak.
    ' Sanity opseg je obavezan: datumski formatirana celija u configu
    ' vrati datum, CStr na srpskom locale-u da "23.1.1900." a CLng to
    ' parsira kao 2311900 (tacke = hiljade) -- DateAdd preko 9999. god
    ' onda obara celu stampu greskom 5.
    Dim rokMeseci As Long, rokD As Double
    rokMeseci = 24
    If IsNumeric(GetConfigValue(CFG_GP_ROK_MESECI)) Then
        rokD = CDbl(GetConfigValue(CFG_GP_ROK_MESECI))
        If rokD >= 1 And rokD <= 600 And rokD = Fix(rokD) Then _
            rokMeseci = CLng(rokD)
    End If
    ' Rok PO VRSTI proizvoda (revizija #9, potvrdjeno poslovno:
    ' proizvodi imaju RAZLICITE rokove): tblVrstaGotovihProizvoda
    ' RokMeseci ima prednost; prazno/nevalidno = globalni default.
    ' Isti sanity opseg kao global (1-600 celih meseci). Sopstvena
    ' petlja promenljiva (vgI): oslanjanje na Dim i nize u proceduri
    ' obara compile ("Variable not defined") -- smoke #10 nalaz.
    Dim vrstaRok As Object: Set vrstaRok = CreateObject("Scripting.Dictionary")
    vrstaRok.CompareMode = vbTextCompare
    Dim vg As Variant, cVgTip As Long, cVgRok As Long, vgI As Long
    If Not GetTable(TBL_VRSTA_GP) Is Nothing Then
        cVgTip = GetColumnIndex(TBL_VRSTA_GP, COL_VGP_TIP)
        cVgRok = GetColumnIndex(TBL_VRSTA_GP, COL_VGP_ROK)
        If cVgTip > 0 And cVgRok > 0 Then
            vg = GetTableData(TBL_VRSTA_GP)
            If IsArray(vg) Then
                For vgI = 1 To UBound(vg, 1)
                    If IsNumeric(vg(vgI, cVgRok)) Then
                        rokD = CDbl(vg(vgI, cVgRok))
                        If rokD >= 1 And rokD <= 600 And rokD = Fix(rokD) Then _
                            vrstaRok(Trim$(CStr(nz(vg(vgI, cVgTip))))) = CLng(rokD)
                    End If
                Next vgI
            End If
        End If
    End If

    ' Podaci prerade: proizvod, pakovanje, datum proizvodnje, neto
    ' izlaz (za "cela paleta / deo") i BRUTO (srazmerno za parcijalu).
    Dim preInfo As Object: Set preInfo = CreateObject("Scripting.Dictionary")
    preInfo.CompareMode = vbTextCompare
    Dim pd As Variant, i As Long
    pd = GetTableData(TBL_PRERADA)
    If IsArray(pd) Then
        Dim cPId As Long, cPTip As Long, cPKut As Long, cPKes As Long
        Dim cPDat As Long, cPNeto As Long, cPBru As Long
        cPId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, SRC)
        cPTip = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, SRC)
        cPKut = RequireColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE, SRC)
        cPKes = RequireColumnIndex(TBL_PRERADA, COL_PRE_KESE, SRC)
        cPDat = RequireColumnIndex(TBL_PRERADA, COL_PRE_DATUM, SRC)
        cPNeto = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
        cPBru = GetColumnIndex(TBL_PRERADA, COL_PRE_BRUTO)
        For i = 1 To UBound(pd, 1)
            If Not preInfo.Exists(Trim$(CStr(nz(pd(i, cPId))))) Then
                Dim bru As Double, net As Double
                net = 0#: bru = 0#
                If IsNumeric(pd(i, cPNeto)) Then net = CDbl(pd(i, cPNeto))
                If cPBru > 0 Then
                    If IsNumeric(pd(i, cPBru)) Then bru = CDbl(pd(i, cPBru))
                End If
                Dim lotKut As Double, lotKes As Double
                lotKut = 0#: lotKes = 0#
                If IsNumeric(pd(i, cPKut)) Then lotKut = CDbl(pd(i, cPKut))
                If IsNumeric(pd(i, cPKes)) Then lotKes = CDbl(pd(i, cPKes))
                preInfo.Add Trim$(CStr(nz(pd(i, cPId)))), Array( _
                    Trim$(CStr(nz(pd(i, cPTip)))), _
                    Trim$(CStr(nz(pd(i, cPKut)))) & " kut. / " & _
                    Trim$(CStr(nz(pd(i, cPKes)))) & " kesa", _
                    pd(i, cPDat), net, bru, lotKut, lotKes)
            End If
        Next i
    End If

    ' Stavke obrasca (1..8): Lot | Proizvod | Dat. proizv. | Rok |
    ' Pakovanje | Paleta | Neto | Bruto. Bruto: cela paleta = bruto
    ' prerade; parcijala srazmerno neto udelu (aritmetika nad stvarnim
    ' merenjima, ne izmisljanje); bez bruto podatka = neto.
    Dim s As Variant, stavke() As Variant, nSt As Long
    Dim totNeto As Double, totBruto As Double
    Dim palCele As Long, palDelovi As Long
    s = GetTableData(TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then Exit Sub
    s = ExcludeStornirano(s, TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then Exit Sub
    Dim cSUt As Long, cSPre As Long, cSBr As Long, cSKol As Long
    Dim cSKut As Long, cSKes As Long
    cSUt = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_UTOVAR_ID, SRC)
    cSPre = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_PRERADA_ID, SRC)
    cSBr = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_BROJ_PRERADE, SRC)
    cSKol = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_KOLICINA, SRC)
    ' Pakovanja STVARNO utovarena (revizija #6 t.4) -- meko: stara
    ' sveska/stare stavke ih nemaju.
    cSKut = GetColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_KUTIJE)
    cSKes = GetColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_KESE)
    ReDim stavke(1 To UBound(s, 1), 1 To 8)
    For i = 1 To UBound(s, 1)
        If Trim$(CStr(nz(s(i, cSUt)))) = Trim$(utovarID) Then
            nSt = nSt + 1
            stavke(nSt, 1) = Trim$(CStr(nz(s(i, cSBr))))
            Dim pv As Variant, preK As String, kol As Double
            preK = Trim$(CStr(nz(s(i, cSPre))))
            kol = 0#
            If IsNumeric(s(i, cSKol)) Then kol = CDbl(s(i, cSKol))
            stavke(nSt, 7) = kol
            totNeto = totNeto + kol
            ' Pakovanje: iz STAVKE (stvarno utovareno); bez podatka na
            ' stavci fallback na punu preradu SAMO za celu paletu --
            ' parcijala bez podatka ostaje prazna (t.4: transportni
            ' dokument radije prazan nego pogresan).
            Dim pakS As String, kutS As String, kesS As String
            pakS = "": kutS = "": kesS = ""
            If cSKut > 0 Then
                If IsNumeric(s(i, cSKut)) Then _
                    If CDbl(s(i, cSKut)) > 0 Then kutS = CStr(CLng(s(i, cSKut)))
            End If
            If cSKes > 0 Then
                If IsNumeric(s(i, cSKes)) Then _
                    If CDbl(s(i, cSKes)) > 0 Then kesS = CStr(CLng(s(i, cSKes)))
            End If
            If Len(kutS) > 0 And Len(kesS) > 0 Then
                pakS = kutS & " kut. / " & kesS & " kesa"
            ElseIf Len(kutS) > 0 Then
                pakS = kutS & " kut."
            ElseIf Len(kesS) > 0 Then
                pakS = kesS & " kesa"
            End If
            If preInfo.Exists(preK) Then
                pv = preInfo(preK)
                stavke(nSt, 2) = CStr(pv(0))
                If Len(pakS) = 0 Then
                    If kol >= CDbl(pv(3)) - 0.0001 Then
                        pakS = CStr(pv(1))
                    Else
                        ' Stare parcijalne stavke bez upisanih pakovanja
                        ' (pre revizije #12): ista STROGA racunica iz
                        ' kapaciteta lota -- 500 kg / 10 kg = 50.
                        pakS = PakTekst( _
                            PakovanjaZaKg(kol, CDbl(pv(3)), CDbl(pv(5)), _
                                          True, CFG_GP_KG_KUTIJA), _
                            PakovanjaZaKg(kol, CDbl(pv(3)), CDbl(pv(6)), _
                                          True, CFG_GP_KG_KESA))
                    End If
                End If
                stavke(nSt, 5) = pakS
                If IsDate(pv(2)) Then
                    stavke(nSt, 3) = CDate(pv(2))
                    ' Rok po vrsti proizvoda; bez unosa = globalni.
                    Dim rokEf As Long
                    rokEf = rokMeseci
                    If vrstaRok.Exists(CStr(pv(0))) Then _
                        rokEf = CLng(vrstaRok(CStr(pv(0))))
                    stavke(nSt, 4) = DateAdd("m", rokEf, CDate(pv(2)))
                Else
                    stavke(nSt, 3) = ""
                    stavke(nSt, 4) = ""
                End If
                ' Bruto SAMO kad je stvarno izmeren (revizija #11 P1):
                ' cela paleta nosi izmereni bruto prerade; parcijala
                ' NEMA svoj izmeren bruto pa se NE stampa procena --
                ' transportni dokument radije prazan nego priblizan
                ' (isti princip kao pakovanja, t.4 revizije #6).
                If CDbl(pv(3)) > 0.0001 And kol >= CDbl(pv(3)) - 0.0001 Then
                    stavke(nSt, 6) = "1"
                    palCele = palCele + 1
                    If CDbl(pv(4)) > 0.0001 Then
                        stavke(nSt, 8) = CDbl(pv(4))
                        totBruto = totBruto + CDbl(pv(4))
                    Else
                        stavke(nSt, 8) = ""
                    End If
                Else
                    stavke(nSt, 6) = "deo"
                    palDelovi = palDelovi + 1
                    stavke(nSt, 8) = ""
                End If
            Else
                stavke(nSt, 2) = ""
                stavke(nSt, 3) = ""
                stavke(nSt, 4) = ""
                stavke(nSt, 5) = pakS
                stavke(nSt, 6) = ""
                ' Bez podataka prerade nema ni izmerenog bruta.
                stavke(nSt, 8) = ""
            End If
        End If
    Next i
    If nSt = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = FillUtovarSablon(broj, datum, vreme, kupacNaziv, mestoIst, _
                              poBroj, fakBroj, prevoz, napomena, _
                              stavke, nSt, _
                              Array(palCele, palDelovi, totNeto, totBruto))
    If ws Is Nothing Then Exit Sub

    ' Rezim iz Podesavanja (kartica Stampa) -- kao svi dokumenti.
    ' Default PDF, NE PRINT: bez podesenog kljuca dokument ne sme tiho
    ' da ode na stampac (smoke 5c nalaz).
    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_UTOVAR_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            DocExportPdf ws, ThisWorkbook.path & "\Utovar_" & _
                         Replace(broj, "/", "-") & ".pdf", True
        ' OFF -> bez izlaza (sablon je ipak popunjen -- test ga cita)
    End Select
    Exit Sub

EH:
    LogErr SRC
End Sub

' ============================================================
' UTOVAR + GP FAKTURA u jednoj transakciji (v1: 1 utovar = 1 faktura).
' stavke: Collection of Array(preradaID, kolicinaKg, cena).
' Kapije u BASE, pod TX: kupac postoji tacno jednom; prerada postoji
' tacno jednom, nije stornirana, ima IMENOVAN proizvod; kolicina > 0 i
' <= na stanju; dupla prerada u istoj listi zabranjena; cena > 0.
' ============================================================
Public Function CreateUtovarSaFakturom_TX(ByVal kupacID As String, _
                                          ByVal stavke As Collection) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_UTOVAR
    tx.AddTableSnapshot TBL_UTOVAR_STAVKE
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_NOVAC

    CreateUtovarSaFakturom_TX = CreateUtovarSaFakturom(kupacID, stavke)

    If CreateUtovarSaFakturom_TX = "" Then
        Err.Raise vbObjectError + 1730, "CreateUtovarSaFakturom_TX", _
                  "CreateUtovarSaFakturom nije uspeo."
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="UTOVAR_FAKTURA_GP_SUCCESS", _
        severity:="INFO", _
        message:="Utovar + GP faktura created successfully", _
        userId:="Operator", _
        moduleName:="modUtovar", _
        procedureName:="CreateUtovarSaFakturom_TX", _
        entityType:="Faktura", _
        entityID:=CreateUtovarSaFakturom_TX, _
        correlationId:=CreateUtovarSaFakturom_TX
    On Error GoTo 0

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "CreateUtovarSaFakturom_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modUtovar", _
        procedureName:="CreateUtovarSaFakturom_TX", _
        entityType:="Faktura", _
        entityID:="", _
        correlationId:="CreateUtovarSaFakturom", _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateUtovarSaFakturom_TX = ""
End Function

' Base -- NE zovi je spolja (pola upisa bez transakcije).
' Model B (revizija #9): brzi put je tanak kompozit dva core-a --
' utovar (fizicki dokument) + faktura iz njega (finansijski). Isti
' core-ovi rade i razdvojeno: "Napravi utovar" danas, "Fakturisi"
' sutra sa liste Utovari.
Private Function CreateUtovarSaFakturom(ByVal kupacID As String, _
                                        ByVal stavke As Collection) As String
    Dim utovarID As String
    utovarID = CreateUtovarCore(kupacID, stavke)
    CreateUtovarSaFakturom = CreateFakturaIzUtovara(utovarID)
End Function

' ============================================================
' SAMO UTOVAR (model B, revizija #9) -- fizicka isporuka BEZ fakture:
' magacin pravi i stampa utovarnu listu danas, racunovodstvo klikne
' "Fakturisi" na listi Utovari kasnije. Stavka nosi i dogovorenu CENU
' (CenaKg) -- nju kasnije cita CreateFakturaIzUtovara kad utovar nema
' istoriju faktura.
' ============================================================
Public Function CreateUtovar_TX(ByVal kupacID As String, _
                                ByVal stavke As Collection) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_UTOVAR
    tx.AddTableSnapshot TBL_UTOVAR_STAVKE

    CreateUtovar_TX = CreateUtovarCore(kupacID, stavke)

    If CreateUtovar_TX = "" Then
        Err.Raise vbObjectError + 1779, "CreateUtovar_TX", _
                  "CreateUtovarCore nije uspeo."
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="UTOVAR_CREATED", _
        severity:="INFO", _
        message:="Utovar (bez fakture) created", _
        userId:="Operator", _
        moduleName:="modUtovar", _
        procedureName:="CreateUtovar_TX", _
        entityType:="Utovar", _
        entityID:=CreateUtovar_TX, _
        correlationId:=CreateUtovar_TX
    On Error GoTo 0

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "CreateUtovar_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modUtovar", _
        procedureName:="CreateUtovar_TX", _
        entityType:="Utovar", _
        entityID:="", _
        correlationId:="CreateUtovarCore", _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateUtovar_TX = ""
End Function

' Core upisa utovara: kapije + tblUtovar + tblUtovarStavke (kolicina,
' pakovanja, cena). NE zovi spolja bez transakcije.
Private Function CreateUtovarCore(ByVal kupacID As String, _
                                  ByVal stavke As Collection) As String
    Const SRC As String = "CreateUtovarCore"
    On Error GoTo EH

    If Trim$(kupacID) = "" Then
        Err.Raise vbObjectError + 1731, SRC, "KupacID je obavezan."
    End If
    ' Writer je samostalna granica: GP nema prijemnicu za implicitnu
    ' proveru vlasnistva, kupac se proverava ovde.
    Dim kupRows As Collection
    Set kupRows = FindRows(TBL_KUPCI, COL_KUP_ID, Trim$(kupacID))
    If kupRows Is Nothing Then
        Err.Raise vbObjectError + 1751, SRC, _
                  "Kupac ne postoji u tblKupci: " & kupacID
    ElseIf kupRows.count <> 1 Then
        Err.Raise vbObjectError + 1751, SRC, _
                  "Kupac ne postoji jednoznacno u tblKupci: " & kupacID & _
                  "; Count=" & CStr(kupRows.count)
    End If
    If stavke Is Nothing Then
        Err.Raise vbObjectError + 1732, SRC, "Stavke nisu prosledjene."
    End If
    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1733, SRC, _
                  "Utovar mora imati bar jednu stavku."
    End If

    ' Fail-fast schema guards -- bez utovar tabela se staje ODMAH.
    RequireColumnIndex TBL_UTOVAR, COL_UT_ID, SRC
    RequireColumnIndex TBL_UTOVAR_STAVKE, COL_UTS_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_BROJ_PRERADE, SRC
    RequireColumnIndex TBL_FAKTURA_STAVKE, COL_FS_UTOVAR_ID, SRC

    Dim preData As Variant
    preData = GetTableData(TBL_PRERADA)
    If IsEmpty(preData) Then
        Err.Raise vbObjectError + 1734, SRC, "Tabela prerada je prazna."
    End If

    Dim colNetoIzlaz As Long, colBroj As Long, colGodina As Long
    Dim colStorno As Long, colTipGp As Long, colKut As Long, colKes As Long
    colNetoIzlaz = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
    colBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, SRC)
    colGodina = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, SRC)
    colStorno = RequireColumnIndex(TBL_PRERADA, COL_STORNIRANO, SRC)
    colTipGp = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, SRC)
    ' Pakovanja (revizija #6 t.4) -- meko: sluze za izracunljiv broj
    ' pakovanja utovarene kolicine, ne za kapiju.
    colKut = GetColumnIndex(TBL_PRERADA, COL_PRE_KUTIJE)
    colKes = GetColumnIndex(TBL_PRERADA, COL_PRE_KESE)

    ' Na stanju = proizvedeno - vec utovareno (jedno pravilo za sve).
    Dim utovareno As Object: Set utovareno = UtovarenoPoPreradi()

    ' Pre-validacija SVIH stavki pre ijednog upisa.
    Dim s As Variant, preradaID As String, kolicina As Double, cena As Double
    Dim rows As Collection, rowPre As Long, raspolozivo As Double
    Dim preRows As Object, preValues As Object
    Set preRows = CreateObject("Scripting.Dictionary")
    Set preValues = CreateObject("Scripting.Dictionary")

    For Each s In stavke
        preradaID = Trim$(CStr(s(0)))
        If Len(preradaID) = 0 Then
            Err.Raise vbObjectError + 1735, SRC, "PreradaID je obavezan."
        End If
        If Not IsNumeric(s(1)) Then
            Err.Raise vbObjectError + 1752, SRC, _
                      "Kolicina nije numericka. PreradaID=" & preradaID
        End If
        kolicina = CDbl(s(1))
        If kolicina <= 0 Then
            Err.Raise vbObjectError + 1752, SRC, _
                      "Kolicina mora biti veca od nule. PreradaID=" & preradaID
        End If
        If Not IsNumeric(s(2)) Then
            Err.Raise vbObjectError + 1736, SRC, _
                      "Cena nije numericka. PreradaID=" & preradaID
        End If
        cena = CDbl(s(2))
        If cena <= 0 Then
            Err.Raise vbObjectError + 1736, SRC, _
                      "Cena mora biti veca od nule. PreradaID=" & preradaID
        End If
        If preRows.Exists(preradaID) Then
            Err.Raise vbObjectError + 1737, SRC, _
                      "Dupla prerada u izboru: " & preradaID
        End If

        Set rows = FindRows(TBL_PRERADA, COL_PRE_ID, preradaID)
        If rows Is Nothing Then
            Err.Raise vbObjectError + 1738, SRC, _
                      "Prerada nije pronadjena: " & preradaID
        End If
        If rows.count = 0 Then
            Err.Raise vbObjectError + 1738, SRC, _
                      "Prerada nije pronadjena: " & preradaID
        End If
        If rows.count > 1 Then
            Err.Raise vbObjectError + 1739, SRC, _
                      "Duplikat PreradaID=" & preradaID & _
                      "; Count=" & CStr(rows.count)
        End If
        rowPre = CLng(rows(1))

        ' Inline "DA" provera (IsStorniranoValue je Private u modStorno).
        If UCase$(Trim$(CStr(nz(preData(rowPre, colStorno))))) = "DA" Then
            Err.Raise vbObjectError + 1740, SRC, _
                      "Prerada je stornirana: " & preradaID
        End If
        ' Stavka prodajne fakture mora imenovati proizvod.
        If Len(Trim$(CStr(nz(preData(rowPre, colTipGp))))) = 0 Then
            Err.Raise vbObjectError + 1750, SRC, _
                      "TipGotovogProizvoda je prazan -- faktura mora imenovati proizvod: " & preradaID
        End If
        If Not IsNumeric(preData(rowPre, colNetoIzlaz)) Then
            Err.Raise vbObjectError + 1742, SRC, _
                      "NetoIzlazKg nije numericki. PreradaID=" & preradaID
        End If
        ' KLJUCNA kapija graina: kolicina <= na stanju. Parcijalna
        ' prodaja je legalna; prekoracenje stanja nije.
        raspolozivo = CDbl(preData(rowPre, colNetoIzlaz))
        If utovareno.Exists(preradaID) Then _
            raspolozivo = raspolozivo - CDbl(utovareno(preradaID))
        If kolicina > raspolozivo + 0.0001 Then
            Err.Raise vbObjectError + 1753, SRC, _
                      "Kolicina " & CStr(kolicina) & " kg prelazi stanje (" & _
                      CStr(raspolozivo) & " kg). PreradaID=" & preradaID
        End If

        Dim preKut As Double, preKes As Double
        preKut = 0#: preKes = 0#
        If colKut > 0 Then
            If IsNumeric(preData(rowPre, colKut)) Then preKut = CDbl(preData(rowPre, colKut))
        End If
        If colKes > 0 Then
            If IsNumeric(preData(rowPre, colKes)) Then preKes = CDbl(preData(rowPre, colKes))
        End If

        preRows.Add preradaID, rowPre
        preValues.Add preradaID, Array( _
            kolicina, cena, _
            Trim$(CStr(nz(preData(rowPre, colBroj)))) & "/" & _
            Trim$(CStr(nz(preData(rowPre, colGodina)))), _
            CDbl(preData(rowPre, colNetoIzlaz)), preKut, preKes)
    Next s

    Dim ukupno As Double, key As Variant, preVals As Variant
    For Each key In preValues.keys
        preVals = preValues(CStr(key))
        ukupno = ukupno + (CDbl(preVals(0)) * CDbl(preVals(1)))
    Next key
    If ukupno <= 0 Then
        Err.Raise vbObjectError + 1743, SRC, _
                  "Ukupan iznos fakture mora biti veci od nule."
    End If

    ' --- UPIS 1: utovarna lista (fizicka isporuka, danasnji datum).
    Dim utovarID As String
    utovarID = GetNextID(TBL_UTOVAR, COL_UT_ID, "UT-")
    If utovarID = "" Then
        Err.Raise vbObjectError + 1754, SRC, "GetNextID nije vratio UtovarID."
    End If

    ' Positional AppendRow je ovde bezbedan: tblUtovar/tblUtovarStavke
    ' pravi EnsureUtovarSchemaCore pa je redosled kolona nas (v. Array
    ' u modSetup); svaka BUDUCA kolona ide na kraj (EnsureDataTable).
    Dim rowUtNovi As Long
    rowUtNovi = AppendRow(TBL_UTOVAR, Array( _
        utovarID, GenerateBrojUtovara(), Year(Date), Date, kupacID, _
        "", "", "", ""))
    If rowUtNovi <= 0 Then
        Err.Raise vbObjectError + 1755, SRC, "AppendRow nije uspeo za tblUtovar."
    End If
    ' Vreme utovara (krug 5d) -- po imenu, kolona je dopuna na kraju.
    RequireUpdateCell TBL_UTOVAR, rowUtNovi, COL_UT_VREME, _
                      Format$(Now, "hh:mm"), SRC

    Dim stavkaNum As Long, rowUts As Long
    For Each s In stavke
        preradaID = Trim$(CStr(s(0)))
        preVals = preValues(preradaID)
        stavkaNum = stavkaNum + 1
        rowUts = AppendRow(TBL_UTOVAR_STAVKE, Array( _
            utovarID & "-" & Format$(stavkaNum, "00"), utovarID, preradaID, _
            CStr(preVals(2)), CDbl(preVals(0)), ""))
        If rowUts <= 0 Then
            Err.Raise vbObjectError + 1756, SRC, _
                      "AppendRow nije uspeo za tblUtovarStavke."
        End If
        ' Pakovanja STVARNO utovarena (revizija #6 t.4) -- upis po
        ' imenu (kolone su dopuna na kraju); prazno kad nije dokazivo.
        Dim pkKut As Variant, pkKes As Variant
        UtsPakovanja CDbl(preVals(0)), CDbl(preVals(3)), _
                     CDbl(preVals(4)), CDbl(preVals(5)), pkKut, pkKes
        If Not IsEmpty(pkKut) Then _
            RequireUpdateCell TBL_UTOVAR_STAVKE, rowUts, COL_UTS_KUTIJE, pkKut, SRC
        If Not IsEmpty(pkKes) Then _
            RequireUpdateCell TBL_UTOVAR_STAVKE, rowUts, COL_UTS_KESE, pkKes, SRC
        ' Dogovorena cena na STAVCI (model B): fakturisanje kasnije je
        ' cita odavde kad utovar nema istoriju faktura.
        RequireUpdateCell TBL_UTOVAR_STAVKE, rowUts, COL_UTS_CENA, _
                          CDbl(preVals(1)), SRC
    Next s

    CreateUtovarCore = utovarID
    Exit Function

EH:
    Dim errNum2 As Long
    Dim errDesc2 As String
    Dim errSrc2 As String

    errNum2 = Err.Number
    errDesc2 = Err.description
    errSrc2 = Err.SOURCE

    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum2, SRC, "Source=" & errSrc2 & " | " & errDesc2
End Function


' ============================================================
' NOVA FAKTURA IZ POSTOJECEG UTOVARA (revizija #6 t.1 -- lifecycle).
' Posle storna GP fakture utovar ostaje aktivan i nefakturisan, a roba
' je FIZICKI otisla -- nov utovar bi tvrdio da je izasla dva puta.
' Ovaj writer koristi postojeci utovar: NE dira stanje, cita njegove
' aktivne stavke, pravi SAMO novu fakturu + stavke, ponovo markira
' utovar i ostavlja originalni DatumUtovara (SEF datum isporuke).
' Cene: iz POSLEDNJE (stornirane) fakture tog utovara; utovar
' napravljen BEZ fakture (model B) jos nema FST istoriju pa se cita
' dogovorena cena sa utovarne stavke (CenaKg).
' novaCena > 0 = KOREKCIJA CENE (revizija #8): ista roba, isti utovar,
' nova faktura sa novom cenom -- fizicka isporuka se NE falsifikuje
' stornom utovara zbog pogresne cene. v1: jedna cena za ceo utovar,
' pa je za utovar sa VISE prerada korekcija odbijena (razlicite cene
' po stavci jos nisu podrzane).
' ============================================================
Public Function CreateFakturaIzUtovara_TX(ByVal utovarID As String, _
        Optional ByVal novaCena As Double = 0) As String
    Dim tx As clsTransaction
    Set tx = New clsTransaction

    On Error GoTo EH

    tx.BeginTx
    tx.AddTableSnapshot TBL_UTOVAR
    tx.AddTableSnapshot TBL_FAKTURE
    tx.AddTableSnapshot TBL_FAKTURA_STAVKE
    tx.AddTableSnapshot TBL_NOVAC

    CreateFakturaIzUtovara_TX = CreateFakturaIzUtovara(utovarID, novaCena)

    If CreateFakturaIzUtovara_TX = "" Then
        Err.Raise vbObjectError + 1760, "CreateFakturaIzUtovara_TX", _
                  "CreateFakturaIzUtovara nije uspeo."
    End If

    tx.CommitTx

    On Error Resume Next
    Monitor_Event _
        eventType:="FAKTURA_IZ_UTOVARA_SUCCESS", _
        severity:="INFO", _
        message:="GP faktura iz postojeceg utovara", _
        userId:="Operator", _
        moduleName:="modUtovar", _
        procedureName:="CreateFakturaIzUtovara_TX", _
        entityType:="Faktura", _
        entityID:=CreateFakturaIzUtovara_TX, _
        correlationId:=utovarID
    On Error GoTo 0

    Set tx = Nothing
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr "CreateFakturaIzUtovara_TX"
    On Error Resume Next
    Monitor_Error _
        moduleName:="modUtovar", _
        procedureName:="CreateFakturaIzUtovara_TX", _
        entityType:="Faktura", _
        entityID:="", _
        correlationId:=utovarID, _
        errorNumber:=errNum, _
        errorDescription:=errDesc, _
        errorSource:=errSrc
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    CreateFakturaIzUtovara_TX = ""
End Function

' Base -- NE zovi je spolja (pola upisa bez transakcije).
Private Function CreateFakturaIzUtovara(ByVal utovarID As String, _
        Optional ByVal novaCena As Double = 0) As String
    Const SRC As String = "CreateFakturaIzUtovara"
    On Error GoTo EH

    utovarID = Trim$(utovarID)
    If Len(utovarID) = 0 Then
        Err.Raise vbObjectError + 1761, SRC, "UtovarID je obavezan."
    End If
    If novaCena < 0 Then
        Err.Raise vbObjectError + 1777, SRC, "Cena ne moze biti negativna."
    End If

    ' --- Kapije nad utovarom: postoji tacno jednom, aktivan, NEFAKTURISAN.
    Dim utRows As Collection
    Set utRows = FindRows(TBL_UTOVAR, COL_UT_ID, utovarID)
    If utRows Is Nothing Then
        Err.Raise vbObjectError + 1762, SRC, "Utovar ne postoji: " & utovarID
    End If
    If utRows.count <> 1 Then
        Err.Raise vbObjectError + 1762, SRC, _
                  "Utovar ne postoji jednoznacno: " & utovarID & _
                  "; Count=" & CStr(utRows.count)
    End If
    Dim ut As Variant, rowUt As Long
    ut = GetTableData(TBL_UTOVAR)
    rowUt = CLng(utRows(1))
    If UCase$(Trim$(CStr(nz(ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_STORNIRANO, SRC)))))) = "DA" Then
        Err.Raise vbObjectError + 1763, SRC, "Utovar je storniran: " & utovarID
    End If
    If UCase$(Trim$(CStr(nz(ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURISANO, SRC)))))) = "DA" _
       Or Len(Trim$(CStr(nz(ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_FAKTURA_ID, SRC)))))) > 0 Then
        Err.Raise vbObjectError + 1764, SRC, _
                  "Utovar je vec fakturisan: " & utovarID & _
                  " -- prvo storno postojece fakture."
    End If

    Dim kupacID As String
    kupacID = Trim$(CStr(nz(ut(rowUt, RequireColumnIndex(TBL_UTOVAR, COL_UT_KUPAC, SRC)))))
    Dim kupRows As Collection
    Set kupRows = FindRows(TBL_KUPCI, COL_KUP_ID, kupacID)
    If kupRows Is Nothing Then
        Err.Raise vbObjectError + 1765, SRC, "Kupac ne postoji u tblKupci: " & kupacID
    ElseIf kupRows.count <> 1 Then
        Err.Raise vbObjectError + 1765, SRC, _
                  "Kupac ne postoji jednoznacno u tblKupci: " & kupacID
    End If

    ' --- Aktivne stavke utovara (roba koja je fizicki otisla).
    Dim s As Variant, i As Long
    s = GetTableData(TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then
        Err.Raise vbObjectError + 1766, SRC, "Nema utovarnih stavki."
    End If
    s = ExcludeStornirano(s, TBL_UTOVAR_STAVKE)
    If Not IsArray(s) Then
        Err.Raise vbObjectError + 1766, SRC, "Nema aktivnih utovarnih stavki."
    End If
    Dim cSUt As Long, cSPre As Long, cSBr As Long, cSKol As Long
    cSUt = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_UTOVAR_ID, SRC)
    cSPre = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_PRERADA_ID, SRC)
    cSBr = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_BROJ_PRERADE, SRC)
    cSKol = RequireColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_KOLICINA, SRC)

    ' --- Kontradikcija (revizija #7 B2, kanonski helper #10):
    ' "nefakturisan" utovar sa aktivnom FST -- finansijski dokument
    ' tvrdi prodaju koju utovar porice; re-fakturisanje bi napravilo
    ' duplu prodaju.
    Dim aktivnihFst As Long
    aktivnihFst = AktivnihFstZaUtovar(utovarID)
    If aktivnihFst > 0 Then
        Err.Raise vbObjectError + 1776, SRC, _
                  "Utovar " & utovarID & " nije markiran kao fakturisan, a nosi " & _
                  CStr(aktivnihFst) & " aktivnih faktura-stavki -- podaci su " & _
                  "neusaglaseni, re-fakturisanje je blokirano."
    End If

    ' --- Cene iz prethodnih (storniranih) faktura ovog utovara: FST se
    ' NE filtrira po stornu -- bas stornirani redovi nose cene; kasniji
    ' red pobedjuje (poslednja faktura).
    Dim cene As Object: Set cene = CreateObject("Scripting.Dictionary")
    cene.CompareMode = vbTextCompare
    Dim fs As Variant, cFsUt As Long, cFsPre As Long, cFsCena As Long
    fs = GetTableData(TBL_FAKTURA_STAVKE)
    If IsArray(fs) Then
        cFsUt = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_UTOVAR_ID, SRC)
        cFsPre = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_PRERADA_ID, SRC)
        cFsCena = RequireColumnIndex(TBL_FAKTURA_STAVKE, COL_FS_CENA, SRC)
        For i = 1 To UBound(fs, 1)
            If Trim$(CStr(nz(fs(i, cFsUt)))) = utovarID Then
                If IsNumeric(fs(i, cFsCena)) Then
                    If CDbl(fs(i, cFsCena)) > 0 Then _
                        cene(Trim$(CStr(nz(fs(i, cFsPre))))) = CDbl(fs(i, cFsCena))
                End If
            End If
        Next i
    End If

    ' --- Pre-validacija svih stavki pre ijednog upisa.
    Dim stavke As Collection: Set stavke = New Collection
    Dim preradaID As String, kol As Double, ukupno As Double
    For i = 1 To UBound(s, 1)
        If Trim$(CStr(nz(s(i, cSUt)))) = utovarID Then
            preradaID = Trim$(CStr(nz(s(i, cSPre))))
            If Len(preradaID) = 0 Then
                Err.Raise vbObjectError + 1767, SRC, _
                          "Utovarna stavka bez PreradaID: " & utovarID
            End If
            If Not IsNumeric(s(i, cSKol)) Then
                Err.Raise vbObjectError + 1767, SRC, _
                          "Utovarna stavka bez kolicine. PreradaID=" & preradaID
            End If
            kol = CDbl(s(i, cSKol))
            If kol <= 0 Then
                Err.Raise vbObjectError + 1767, SRC, _
                          "Kolicina stavke mora biti > 0. PreradaID=" & preradaID
            End If
            ' Cena: korekcija (novaCena) > poslednja faktura ovog
            ' utovara > dogovorena cena SA STAVKE (model B: utovar
            ' napravljen bez fakture jos nema FST istoriju).
            Dim cenaSt As Double, cSCen As Long
            cSCen = GetColumnIndex(TBL_UTOVAR_STAVKE, COL_UTS_CENA)
            If novaCena > 0 Then
                cenaSt = novaCena
            ElseIf cene.Exists(preradaID) Then
                cenaSt = CDbl(cene(preradaID))
            ElseIf cSCen > 0 And IsNumeric(s(i, cSCen)) Then
                If CDbl(s(i, cSCen)) <= 0 Then
                    Err.Raise vbObjectError + 1768, SRC, _
                              "Cena na utovarnoj stavci nije validna za preradu " & preradaID
                End If
                cenaSt = CDbl(s(i, cSCen))
            Else
                Err.Raise vbObjectError + 1768, SRC, _
                          "Nema cene ni iz prethodne fakture ni sa utovarne stavke za preradu " & _
                          preradaID & " -- unesi cenu pri radnji 'Ponovi fakturu'."
            End If
            ' Naziv proizvoda mora postojati (ista kapija kao izrada).
            If Len(Trim$(CStr(nz(LookupValue(TBL_PRERADA, COL_PRE_ID, preradaID, COL_PRE_TIP_GP))))) = 0 Then
                Err.Raise vbObjectError + 1769, SRC, _
                          "TipGotovogProizvoda je prazan -- faktura mora imenovati proizvod: " & preradaID
            End If
            stavke.Add Array(preradaID, kol, cenaSt, _
                             Trim$(CStr(nz(s(i, cSBr)))))
            ukupno = ukupno + kol * cenaSt
        End If
    Next i
    If stavke.count = 0 Then
        Err.Raise vbObjectError + 1766, SRC, _
                  "Utovar nema aktivnih stavki: " & utovarID
    End If
    ' v1 korekcija cene = JEDNA cena za ceo utovar; utovar sa vise
    ' prerada (razlicite cene po stavci) se odbija umesto da se sve
    ' tiho izravna na istu cenu.
    If novaCena > 0 And stavke.count > 1 Then
        Err.Raise vbObjectError + 1778, SRC, _
                  "Korekcija cene za utovar sa vise prerada jos nije " & _
                  "podrzana (utovar " & utovarID & " ima " & _
                  CStr(stavke.count) & " stavke)."
    End If
    If ukupno <= 0 Then
        Err.Raise vbObjectError + 1770, SRC, _
                  "Ukupan iznos fakture mora biti veci od nule."
    End If

    ' --- Upis fakture (isti pozicioni oblik kao CreateUtovarSaFakturom).
    Dim fakturaID As String
    fakturaID = GetNextID(TBL_FAKTURE, COL_FAK_ID, "FAK-")
    If fakturaID = "" Then
        Err.Raise vbObjectError + 1771, SRC, "GetNextID nije vratio FakturaID."
    End If
    Dim brojFakture As String
    brojFakture = GenerateBrojFakture()
    If brojFakture = "" Then
        Err.Raise vbObjectError + 1771, SRC, _
                  "GenerateBrojFakture nije vratio broj fakture."
    End If

    If AppendRow(TBL_FAKTURE, Array( _
        fakturaID, brojFakture, Date, kupacID, ukupno, STATUS_NEPLACENO, _
        Empty, "", "", WF_LOCAL_FINALIZED, "", "", "", Empty, Empty, _
        "", "", "", 0, "Ne", "")) <= 0 Then
        Err.Raise vbObjectError + 1772, SRC, "AppendRow nije uspeo za tblFakture."
    End If

    Dim sv As Variant, stavkaNum As Long, rowStavke As Long
    For Each sv In stavke
        stavkaNum = stavkaNum + 1
        rowStavke = AppendRow(TBL_FAKTURA_STAVKE, Array( _
            fakturaID & "-" & Format$(stavkaNum, "00"), fakturaID, "", _
            CDbl(sv(1)), CDbl(sv(2)), "", "", "", ""))
        If rowStavke <= 0 Then
            Err.Raise vbObjectError + 1773, SRC, _
                      "AppendRow nije uspeo za tblFakturaStavke."
        End If
        RequireUpdateCell TBL_FAKTURA_STAVKE, rowStavke, COL_FS_PRERADA_ID, _
                          CStr(sv(0)), SRC
        RequireUpdateCell TBL_FAKTURA_STAVKE, rowStavke, COL_FS_BROJ_PRERADE, _
                          CStr(sv(3)), SRC
        RequireUpdateCell TBL_FAKTURA_STAVKE, rowStavke, COL_FS_UTOVAR_ID, _
                          utovarID, SRC
    Next sv

    ' Marker: utovar ponovo fakturisan (1:1), DatumUtovara NETAKNUT.
    RequireUpdateCell TBL_UTOVAR, rowUt, COL_UT_FAKTURISANO, "Da", SRC
    RequireUpdateCell TBL_UTOVAR, rowUt, COL_UT_FAKTURA_ID, fakturaID, SRC

    ApplyAvansToFaktura kupacID, fakturaID

    CreateFakturaIzUtovara = fakturaID
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    LogErr SRC
    On Error GoTo 0

    Err.Raise errNum, SRC, "Source=" & errSrc & " | " & errDesc
End Function
