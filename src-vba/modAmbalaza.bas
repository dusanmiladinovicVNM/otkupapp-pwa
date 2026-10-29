Attribute VB_Name = "modAmbalaza"

'Attribute VB_Name = "modAmbalaza"
Option Explicit

' ============================================================
' modAmbalaza - knjiga ambalaze (AMB-10)
'
' Jedan model: red je PRENOS (OdNalog -> NaNalog, VrstaKretanja), a storno je
' kontra-stav sa StornoOd. Stari model (Smer / EntitetID / EntitetTip / VozacID /
' Stornirano / ReversID i pisac TrackAmbalaza) obrisan je u 10e-3, zajedno sa
' kolonama iz kanona (docs/DOMEN/AMBALAZA.md 6.13e).
'
' Pisci ne otvaraju svoju transakciju: zovu se iz vecih poslovnih TX-ova koji
' vec snapshot-uju tblAmbalaza (AMB-INV-08).
' ============================================================

' --- KNJIGA (AMB-10b): greske koje POZIVALAC mora da razlikuje ---
'
' Potvrda deficita nije kvar nego PITANJE operateru (AMB-10, 6.5). Ekran je mora
' razlikovati od svake druge greske PO BROJU, ne po tekstu poruke: tekst je
' prevodiv i menja se, broj je ugovor.
Public Const AMB_ERR_POTVRDA_DEFICITA As Long = vbObjectError + 4470
Public Const AMB_ERR_DEFICIT_NEPOKRIV As Long = vbObjectError + 4471
Public Const AMB_ERR_IDENTITET As Long = vbObjectError + 4472
Public Const AMB_ERR_JEDAN_PARTNER As Long = vbObjectError + 4475
Public Const AMB_ERR_STORNO As Long = vbObjectError + 4476
Private Const AMB_ERR_KNJIGA_KVAR As Long = vbObjectError + 4473
Private Const AMB_ERR_KNJIGA_ULAZ As Long = vbObjectError + 4474

' TEST SEAM (10d, review #414 P3): pad ispravke reversa NEPOSREDNO pre commit-a,
' kad su storno, nov revers, trag i veza vec napisani. Jedini nacin da se izmeri
' da spoljni snapshot vraca i red koji je ugnjezdeni CreateCorrectionContext vec
' commit-ovao. Tvrdo gejtovan test rezimom; modTest.ResetSeamova ga gasi.
Private mIspravkaPadPreCommit As Boolean

' ============================================================
' Helpers
' ============================================================

Private Function AmbText(ByVal v As Variant) As String
    If isError(v) Or IsNull(v) Or IsEmpty(v) Then
        AmbText = ""
    Else
        AmbText = Trim$(CStr(v))
    End If
End Function

' ============================================================
' READ: saldo ambalaze po entitetu
'
' Returns:
'   result(row, 1) = TipAmbalaze
'   result(row, 2) = Saldo
'
' Semantika:
'   Ulaz  => +Kolicina
'   Izlaz => -Kolicina
'
' BEZ HANDLERA (AMB-10-ODL-31, stavka 135). Ranije je kvar knjige isao u LogErr
' i vracao Empty, pa je papir nosio nulu ili tiho izostavljao redove stanja --
' netacna tvrdnja operateru, ne odsustvo podatka. Kvar ide pozivaocu.
' ============================================================

Public Function GetAmbalazeStanje(ByVal entitetID As String, _
                                  ByVal entitetTip As String) As Variant
    Const SRC As String = "modAmbalaza.GetAmbalazeStanje"

    If Len(Trim$(entitetID)) = 0 Or Len(Trim$(entitetTip)) = 0 Then
        GetAmbalazeStanje = Empty
        Exit Function
    End If

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)

    If IsEmpty(data) Then
        GetAmbalazeStanje = Empty
        Exit Function
    End If

    ' NOV MODEL (AMB-10c): red imenuje OBE strane, pa JEDAN red dotice DVA
    ' naloga -- suprotnim znakom. Stari oblik je nosio jednu stranu i Smer, pa
    ' je filter bio jedan; sada se mora pitati za obe strane, inace polovina
    ' knjige ispada iz salda.
    '
    ' Pravilo znaka se NE prepisuje ovde nego je isto kao u AmbSaldoNaloga
    ' (6.8): +kolicina kad je nalog ODREDISTE, -kolicina kad je IZVOR. Ova
    ' funkcija se od nje razlikuje samo oblikom -- vraca SVE tipove ambalaze
    ' tog naloga odjednom, jer stampa i izvestaj tako citaju.
    '
    ' DVE RAZLICITE STVARI, da se ne pomesaju (review 08.10.2026, P2 #2):
    '   RedDoticeKnjigu bira KANONSKE redove -- i original i kontra-stav su to.
    '   Storno se ponistava ALGEBARSKI: original +N i kontra-stav -N daju 0.
    ' Stari ExcludeStornirano je citao mutabilnu zastavicu koju nov model ne
    ' pise, pa ovde ne bi gasio nista -- zato odlazi, a ne zato sto ga menja
    ' RedDoticeKnjigu.
    RequireKnjigaSchema SRC

    Dim kolIdx As Object, vrsteDok As Object
    Set kolIdx = KnjigaZaCitanje(data, SRC, vrsteDok)

    Dim colTip As Long
    Dim colKol As Long
    Dim colOdTip As Long, colOdID As Long, colNaTip As Long, colNaID As Long

    colTip = kolIdx(COL_AMB_TIP)
    colKol = kolIdx(COL_AMB_KOLICINA)
    colOdTip = kolIdx(COL_AMB_OD_TIP)
    colOdID = kolIdx(COL_AMB_OD_ID)
    colNaTip = kolIdx(COL_AMB_NA_TIP)
    colNaID = kolIdx(COL_AMB_NA_ID)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long, znak As Long
    For i = 1 To UBound(data, 1)
        znak = 0
        If RedDoticeKnjigu(data, i, kolIdx) Then
            If IstiNalog(AmbText(data(i, colNaTip)), AmbText(data(i, colNaID)), _
                         entitetTip, entitetID) Then
                znak = 1
            ElseIf IstiNalog(AmbText(data(i, colOdTip)), AmbText(data(i, colOdID)), _
                             entitetTip, entitetID) Then
                znak = -1
            End If
        End If
        If znak <> 0 Then

            Dim key As String
            key = AmbText(data(i, colTip))

            If Len(key) = 0 Then
                Err.Raise vbObjectError + 4411, SRC, _
                          "Prazan TipAmbalaze u redu " & CStr(i)
            End If

            If Not IsNumeric(data(i, colKol)) Then
                Err.Raise vbObjectError + 4410, SRC, _
                          "Neispravna koli" & ChrW(269) & "ina ambala" & ChrW(382) & "e u redu " & CStr(i)
            End If

            If Not dict.Exists(key) Then dict.Add key, 0&

            dict(key) = CLng(dict(key)) + znak * CLng(data(i, colKol))
        End If
    Next i

    If dict.count = 0 Then
        GetAmbalazeStanje = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 2)

    Dim keys As Variant
    keys = dict.keys

    For i = 0 To dict.count - 1
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = dict(keys(i))
    Next i

    GetAmbalazeStanje = result
End Function

' ============================================================
' READ: pocetno stanje ambalaze kooperanta PRE datog bloka (D-057).
'
' Omotac oko AmbStanjeOkoDokumenta, sa zatecenim ugovorom otkupnog lista: nalog
' je kooperant, a blok je skup DokumentID-eva BEZ tipa dokumenta -- test 223 mu
' predaje i revers kao "blok". Granica, ODL-27 i to da kvar knjige ide pozivaocu
' zive u jezgru, jednom za sve papire (AMB-10-ODL-31).
'
' Stanje je POSLOVNO (AMB-POSL-01): otkupni list kooperantu kaze koliko gajbi
' firme duguje, isto kao kartica -- pokrice deficita su njegove gajbe, ne nase.
' Revers kooperantu cita isto (AMB-POSL-03); papiri stanice i vozaca i zapisnik
' o otpisu citaju FIZICKO stanje istog jezgra, pa se jezgro ne menja nego
' dobija izbor.
'
' blockOtkupIDs: niz (npr. Split rezultat) ili string "OTK-1 + OTK-2".
' Vraca Long (entitetski saldo: Ulaz = +Kolicina, Izlaz = -Kolicina).
' ============================================================
Public Function GetKooperantAmbOpening(ByVal koopID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal blockOtkupIDs As Variant, _
                                       Optional ByRef outPoznato As Boolean) As Long
    GetKooperantAmbOpening = AmbStanjeOkoDokumenta(AMB_NALOG_KOOPERANT, koopID, tipAmb, _
                                                   "", blockOtkupIDs, outPoznato, _
                                                   samoPoslovno:=True)
End Function

' DA LI RED DOTICE NALOG, I SA KOJIM ZNAKOM.
'
' Jedno mesto, jer se isto pitanje u AmbStanjeOkoDokumenta postavlja TRIPUT: pri
' trazenju granice dokumenta, pri sumiranju pre nje i pri sumiranju redova
' dokumenta. Kopije istog pravila bi se razisle prvom doradom -- a razlika se ne
' bi videla kao greska nego kao pogresno stanje na stampi.
Private Function NalogZnakReda(ByRef data As Variant, ByVal i As Long, _
                               ByRef kol As Object, ByVal nalogTip As String, _
                               ByVal nalogID As String) As Long
    If Not RedDoticeKnjigu(data, i, kol) Then Exit Function

    If IstiNalog(AmbText(data(i, kol(COL_AMB_NA_TIP))), _
                 AmbText(data(i, kol(COL_AMB_NA_ID))), _
                 nalogTip, nalogID) Then
        NalogZnakReda = 1
    ElseIf IstiNalog(AmbText(data(i, kol(COL_AMB_OD_TIP))), _
                     AmbText(data(i, kol(COL_AMB_OD_ID))), _
                     nalogTip, nalogID) Then
        NalogZnakReda = -1
    End If
End Function

' RED PRIPADA DOKUMENTU: DokumentID je u skupu, i tip dokumenta je trazeni.
' Identitet dokumenta u knjizi je (DokumentTIP, DokumentID) -- AMB-INV-04. Prazan
' dokTip ne poredi tip: to je zatecen ugovor GetKooperantAmbOpening.
Private Function RedJeDokumenta(ByRef data As Variant, ByVal i As Long, _
                                ByVal colDokTip As Long, ByVal colDokID As Long, _
                                ByVal dokTip As String, ByVal blk As Object) As Boolean
    If Not blk.Exists(AmbText(data(i, colDokID))) Then Exit Function
    If Len(Trim$(dokTip)) > 0 Then
        If StrComp(AmbText(data(i, colDokTip)), Trim$(dokTip), vbTextCompare) <> 0 Then
            Exit Function
        End If
    End If
    RedJeDokumenta = True
End Function

' ISTI TIP AMBALAZE -- bez obzira na velika slova i razmake, isto kao kanonski
' saldo (AmbSaldoNaloga: StrComp vbTextCompare) i izbor reda za papir
' (modIzvestaj.AmbTipKljuc: UCase$ i Trim$). Pisac trazi samo neprazan tip, pa
' knjiga sme da nosi "letvarica" i "LETVARICA"; citalac koji ih razlikuje daje
' papiru drugi odgovor od salda (review #421, P2 #2).
Private Function IstiTipAmb(ByVal a As String, ByVal b As String) As Boolean
    IstiTipAmb = (StrComp(Trim$(a), Trim$(b), vbTextCompare) = 0)
End Function

' TEHNICKI RED -- ne ulazi u POSLOVNI saldo partnera (AMB-POSL-01).
'
' Poslovni dogadjaj je ono sto je pozivalac pisca TRAZIO (AmbVrstaJeZahtev).
' ULAZ_TUDJE nije zahtev nego posledica koju pisac generise sam: pokrice
' deficita (AMB-INV-07) i ulaz pozajmice (AMB-INV-13) -- partnerove SOPSTVENE
' gajbe ulaze u opticaj. On ih fizicki drzi, pa ostaju u fizickom saldu; firma
' ga njima nije zaduzila, pa ne ulaze u poslovni. Sta firma za njih duguje nosi
' obaveza (AmbObavezaPartneru) -- treci, odvojen broj.
'
' Kontra-stav se sudi po vrsti ORIGINALA (vrsteDok: AmbID -> vrsta; integritet
' knjige jemci da StornoOd pokazuje na postojeci red), pa je storno pokrica i
' sam tehnicki, a storno pravog prenosa poslovni.
Private Function AmbRedJeTehnicki(ByRef data As Variant, ByVal i As Long, _
                                  ByVal kol As Object, ByVal vrsteDok As Object) As Boolean
    Dim vrsta As String, orig As String
    vrsta = AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA)))
    orig = AmbText(data(i, kol(COL_AMB_STORNO_OD)))
    If Len(orig) > 0 Then
        If vrsteDok.Exists(orig) Then vrsta = CStr(vrsteDok(orig))
    End If
    AmbRedJeTehnicki = Not modAmbalazaUgovor.AmbVrstaJeZahtev(vrsta)
End Function

' ============================================================
' READ: stanje NALOGA OKO JEDNOG DOKUMENTA, po redosledu upisa (AMB-10-ODL-31).
'
' Vraca stanje naloga za tipAmb PRE dokumenta: zbir redova tog naloga i tipa
' upisanih pre PRVOG reda dokumenta koji dotice nalog (granica). outPosle je to
' stanje plus redovi dokumenta za taj nalog i tip -- stanje POSLE dokumenta, i
' kad su posle njega upisani drugi dokumenti. Ponovna stampa je zato ista kao
' prva: kasniji dokument, i storno ranijeg upisan posle, ne menjaju papir
' (ODL-26). Granica je RED, ne datum: dva dokumenta istog dana su dve tacke.
'
' U outPosle ulaze i redovi dokumenta koji nisu poslovni (pokrice deficita,
' ULAZ_TUDJE), jer su deo stanja posle dokumenta. Kontra-stav (StornoOd) ne
' ulazi: on je kasniji dogadjaj, a ne deo dokumenta.
'
' BEZ GRANICE NEMA STANJA (AMB-10-ODL-27, 11.10.2026). Dokument koji nema
' nijedan svoj red za nalog (otkup bez gajbi, neto unos) nema ni TACKU u knjizi
' od koje bi se znalo "pre njega". Citac tada NE pogadja: outPoznato je False, a
' vrednosti nisu stanje. Ranije je ovde stajao fallback na CELU knjigu ("legacy
' red pre dvojnog upisa") -- prva stampa je bila tacna, a svaka ponovna je
' pokazivala DANASNJE stanje (review #415, P2; reprodukovano
' Test_Amb_PonovnaStampaBezRedaKnjige). Starih redova vise nema.
'
' BEZ HANDLERA, namerno (review #415, P2). "Nepoznato" (outPoznato = False) sme
' da znaci samo jedno: dokument nema svoj red knjige (ODL-27). Kvar knjige, pad
' seme ili bilo koja druga greska ide pozivaocu -- handler koji bi je pretvorio u
' False stampao bi pokvarenu knjigu kao legitiman "-". Stampa ima svoj handler i
' list tada ne popunjava.
'
' dokIDs: niz ili string "ID1 + ID2". dokTip: tip dokumenta u knjizi (prazno =
' ne poredi se, v. RedJeDokumenta).
'
' samoPoslovno (AMB-POSL-01): oba zbira preskacu redove tehnickog pokrica
' (AmbRedJeTehnicki). Granica ostaje ista -- prvi red dokumenta, bilo kog
' oblika -- jer je granica TACKA u knjizi, a ne iznos.
' ============================================================
Public Function AmbStanjeOkoDokumenta(ByVal nalogTip As String, _
                                      ByVal nalogID As String, _
                                      ByVal tipAmb As String, _
                                      ByVal dokTip As String, _
                                      ByVal dokIDs As Variant, _
                                      Optional ByRef outPoznato As Boolean, _
                                      Optional ByRef outPosle As Long, _
                                      Optional ByVal samoPoslovno As Boolean = False) As Long
    Const SRC As String = "modAmbalaza.AmbStanjeOkoDokumenta"

    AmbStanjeOkoDokumenta = 0
    outPoznato = False
    outPosle = 0

    If Len(Trim$(nalogTip)) = 0 Or Len(Trim$(nalogID)) = 0 Then Exit Function
    If Len(Trim$(tipAmb)) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    ' Redovi se NE izbacuju iz niza (ExcludeStornirano) nego se preskacu u
    ' petlji: granica dokumenta je INDEKS u append redosledu, a brisanje redova
    ' bi ga pomerilo. Storno gasi RedDoticeKnjigu kroz kontra-stav, koji je i
    ' sam upisan POSLE dokumenta -- pa u stanje pre njega ionako ne ulazi.
    RequireKnjigaSchema SRC

    Dim kolIdx As Object, vrsteDok As Object
    Set kolIdx = KnjigaZaCitanje(data, SRC, vrsteDok)

    Dim colTip As Long, colKol As Long, colDokID As Long
    Dim colDokTip As Long, colStornoOd As Long
    colTip = kolIdx(COL_AMB_TIP)
    colKol = kolIdx(COL_AMB_KOLICINA)
    colDokID = kolIdx(COL_AMB_DOK_ID)
    colDokTip = kolIdx(COL_AMB_DOK_TIP)
    colStornoOd = kolIdx(COL_AMB_STORNO_OD)

    ' Skup ID-eva dokumenta (niz ili string "ID1 + ID2").
    Dim blk As Object
    Set blk = CreateObject("Scripting.Dictionary")
    blk.CompareMode = vbTextCompare

    If IsArray(dokIDs) Then
        Dim el As Variant
        For Each el In dokIDs
            If Len(Trim$(CStr(el))) > 0 Then blk(Trim$(CStr(el))) = True
        Next el
    Else
        Dim parts() As String
        parts = Split(CStr(dokIDs), " + ")
        Dim p As Long
        For p = LBound(parts) To UBound(parts)
            If Len(Trim$(parts(p))) > 0 Then blk(Trim$(parts(p))) = True
        Next p
    End If

    ' 1) Prvi (najraniji) red dokumenta za ovaj nalog = granica "pre dokumenta".
    Dim i As Long
    Dim minIdx As Long: minIdx = 0
    If blk.count > 0 Then
        For i = 1 To UBound(data, 1)
            If NalogZnakReda(data, i, kolIdx, nalogTip, nalogID) <> 0 Then
                If RedJeDokumenta(data, i, colDokTip, colDokID, dokTip, blk) Then
                    minIdx = i
                    Exit For
                End If
            End If
        Next i
    End If

    ' 2) Saldo svih redova PRE granice. Bez granice nema stanja (ODL-27).
    outPoznato = (minIdx > 0)
    If Not outPoznato Then Exit Function
    Dim cutoff As Long
    cutoff = minIdx - 1

    Dim saldo As Long: saldo = 0
    Dim znak As Long
    For i = 1 To cutoff
        znak = NalogZnakReda(data, i, kolIdx, nalogTip, nalogID)
        If znak <> 0 And IstiTipAmb(AmbText(data(i, colTip)), tipAmb) Then
            If samoPoslovno Then
                If AmbRedJeTehnicki(data, i, kolIdx, vrsteDok) Then znak = 0
            End If
            If znak <> 0 And IsNumeric(data(i, colKol)) Then
                saldo = saldo + znak * CLng(data(i, colKol))
            End If
        End If
    Next i

    ' 3) Stanje POSLE dokumenta: pre + redovi dokumenta za nalog i tip, bez
    ' kontra-stava. Prolaz ide preko cele knjige, ne od granice: redovi dokumenta
    ' su po definiciji granice na njoj ili posle nje, pa je ishod isti.
    Dim posle As Long: posle = saldo
    For i = 1 To UBound(data, 1)
        If Len(AmbText(data(i, colStornoOd))) = 0 Then
            If RedJeDokumenta(data, i, colDokTip, colDokID, dokTip, blk) Then
                znak = NalogZnakReda(data, i, kolIdx, nalogTip, nalogID)
                If znak <> 0 And IstiTipAmb(AmbText(data(i, colTip)), tipAmb) Then
                    If samoPoslovno Then
                        If AmbRedJeTehnicki(data, i, kolIdx, vrsteDok) Then znak = 0
                    End If
                    If znak <> 0 And IsNumeric(data(i, colKol)) Then
                        posle = posle + znak * CLng(data(i, colKol))
                    End If
                End If
            End If
        End If
    Next i

    outPosle = posle
    AmbStanjeOkoDokumenta = saldo
End Function

' GetStanicaAmbSaldo (DANASNJI saldo Stanica:<ID>) je obrisan u S6 koraku 3
' (stavka 178): jedini pozivalac je bio grupni otkupni list, koji sada cita
' nalog VOZAC oko svog reda kroz AmbStanjeOkoDokumenta (AMBALAZA.md 6.13g).

' ============================================================
' READ: opcije za "Tip ambalaze" combo
'
' Vraca AKTIVNE tipove iz tblTipAmbalaze (maticni podaci) -- tacno sifarnik.
' Rezervnih tipova (12/1, 6/1 nad praznim sifarnikom) vise nema (KI-021, odluka
' operatera 19.10.2026): pisac trazi tip iz sifarnika (AMB-INV-12), pa bi combo
' nudio tip koji upis odbija. Prazan sifarnik daje prazan combo; tip se dodaje u
' Maticnim podacima. Greska citanja daje prazan spisak uz ERROR u log -- combo
' ne izmislja tip.
' ============================================================
Public Function GetTipAmbalazeOptions() As Variant
    On Error GoTo EH
    GetTipAmbalazeOptions = GetLookupList(TBL_TIP_AMBALAZE, COL_TAMB_TIP, , , True)
    Exit Function
EH:
    LogErr "modAmbalaza.GetTipAmbalazeOptions"
    GetTipAmbalazeOptions = Array()
End Function

' Podrazumevani tip ambalaze za kulturu (tblKulture.TipAmbalaze).
' Match po VrstaVoca (+ SortaVoca ako je data). Vraca "" ako nema.
' Koristi se za auto-popunjavanje tipa ambalaze u frmOtkup/frmDokumenta.
Public Function GetKulturaTipAmbalaze(ByVal vrsta As String, ByVal sorta As String) As String
    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_KULTURE)
    If IsEmpty(data) Then Exit Function

    Dim cv As Long, cS As Long, cT As Long
    cv = GetColumnIndex(TBL_KULTURE, "VrstaVoca")
    cS = GetColumnIndex(TBL_KULTURE, "SortaVoca")
    cT = GetColumnIndex(TBL_KULTURE, COL_KUL_TIP_AMBALAZE)
    If cv = 0 Or cT = 0 Then Exit Function

    Dim i As Long, hit As String
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(nz(data(i, cv))), Trim$(vrsta), vbTextCompare) = 0 Then
            If Len(Trim$(sorta)) = 0 Or cS = 0 _
               Or StrComp(Trim$(nz(data(i, cS))), Trim$(sorta), vbTextCompare) = 0 Then
                hit = Trim$(nz(data(i, cT)))
                If Len(hit) > 0 Then
                    GetKulturaTipAmbalaze = hit
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function

EH:
    GetKulturaTipAmbalaze = ""
End Function



' ============================================================
' KNJIGA AMBALAZE -- PISAC (AMB-10b)
' ============================================================
'
' Dogadjaj je PRENOS: jedan red imenuje OBE strane (OdNalog -> NaNalog), kolicina
' je UVEK pozitivna, smera kao podatka nema. Knjiga je APPEND-ONLY -- pogresan
' unos se ne popravlja UPDATE-om nego stornom i novim dogadjajem. Pun model:
' docs/DOMEN/AMBALAZA.md 6.1-6.9.
'
' Ugovor (zatvorene liste, matrica strana, razresavanje naloga, doprinos obavezi,
' veza dokument <-> kretanje) je modAmbalazaUgovor. Ovde je UPIS i ono sto se MORA
' procitati da bi upis bio zakonit -- nista vise.
'
' 10b-1 nije dirao pozivna mesta: stari pisac je jos pisao stari oblik, pa je
' tblAmbalaza tokom prelaza nosila DVA oblika reda. Od 10e-3 oblik je jedan, a
' red koji mu ne pripada je KVAR (KnjigaIntegritet) -- i dalje FAIL-CLOSED.
'
' ZASTO PISAC CITA STANJE: AMB-INV-07 (nijedan realan nalog ispod nule) i
' AMB-INV-09 (obaveza >= 0) su granice koje se bez stanja ne mogu proveriti. UI
' sme da ih prikaze unapred, ali racun koji vazi je ovaj, u trenutku upisa --
' izmedju pitanja i odgovora stanje se moglo promeniti drugim unosom (6.5). Isti
' obrazac koji repo vec drzi kod ApplyAvansToOtkup / IsplataBlokProblem.
'
' ZASTO OVDE NEMA "On Error GoTo EH / LogErr" BLOKA (stari pisac ga je imao):
' odbijanje pisca je najcesce POSLOVNI ishod, ne kvar. Potvrda deficita je
' pitanje operateru, a ne greska -- da svaki upit zavrsi u logu gresaka, log bi
' prestao da bude signal. Greske se zato ne gutaju i ne prepakuju nego dizu dalje,
' sa svojim brojem.
'
' STA PISAC NE RADI:
'   STORNO       -- kontra-stav sa StornoOd. PrenesiAmbalazu upisuje samo
'                   ORIGINALE, pa StornoOd ostaje prazan; ulaz za storno je
'                   StornirajAmbalazuDokumenta (AMB-10-ODL-16).
'   AMB-INV-08   -- "upis u istoj transakciji sa izvornim dokumentom" se ne moze
'                   dokazati iznutra: clsTransaction nema globalan registar aktivne
'                   transakcije. To je STATICKA kapija nad pozivnim mestima, a njih
'                   u ovom rezu nema -- ide uz njih, u 10b-2.
'   BROJ          -- numericki niz ambalaznog dokumenta (koji modBrojevi kind, koji
'                   kontekst, i prozor jedinstvenosti) vezan je za pozivna mesta:
'                   stari revers broji po (stanica, dan), a revers kupca stanicu
'                   nema. Odluka ide uz cutover; ovde se broj samo zahteva kao
'                   neprazan, uz vlasnika niza (AMB-10-DOK).

Private Sub RequireKnjigaSchema(ByVal sourceName As String)
    modSchema.SchemaReadyOrFail sourceName, TBL_AMBALAZA

    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_OD_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_OD_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_NA_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_NA_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_VRSTA_KRETANJA, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA, COL_AMB_STORNO_OD, sourceName)
End Sub

Private Sub RequireAmbDokSchema(ByVal sourceName As String)
    modSchema.SchemaReadyOrFail sourceName, TBL_AMBALAZA_DOKUMENT

    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_VRSTA, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_DATUM, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_TIP, sourceName)
    Call RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_ID, sourceName)
End Sub

' Indeksi kolona knjige, JEDNOM po citaocu: ime -> indeks. Svaki citalac je ranije
' nosio svoj blok RequireColumnIndex poziva, a ugovor zapisanog reda trazi SVE
' kolone -- jedanaest parametara po pozivu bilo bi necitljivo i lako se razilazi.
Private Function KnjigaIndeksi(ByVal sourceName As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim imena As Variant, j As Long
    ' Samo kanonske kolone. Kolone starog modela su ovde stajale dok je citalac
    ' morao da proveri i NJEGOV ugovor; otisle su iz kanona u 10e-3.
    imena = Array(COL_AMB_ID, COL_AMB_DATUM, COL_AMB_TIP, COL_AMB_KOLICINA, _
                  COL_AMB_OD_TIP, COL_AMB_OD_ID, COL_AMB_NA_TIP, COL_AMB_NA_ID, _
                  COL_AMB_DOK_TIP, COL_AMB_DOK_ID, COL_AMB_VRSTA_KRETANJA, _
                  COL_AMB_STORNO_OD)

    For j = LBound(imena) To UBound(imena)
        d.Add CStr(imena(j)), RequireColumnIndex(TBL_AMBALAZA, CStr(imena(j)), sourceName)
    Next j

    Set KnjigaIndeksi = d
End Function

' TRI STANJA REDA, NE DVA (review #400, P2).
'
' Red DOTICE knjigu kad je BILO KOJA nova kolona popunjena. Prva verzija je gledala
' samo OdNalogTip, pa je red sa praznim izvorom a popunjenim odredistem prolazio kao
' "stari oblik" -- tiho, i van svakog salda.
Private Function RedDoticeKnjigu(ByRef data As Variant, ByVal i As Long, _
                                 ByRef kol As Object) As Boolean
    Dim imena As Variant, j As Long
    imena = Array(COL_AMB_OD_TIP, COL_AMB_OD_ID, COL_AMB_NA_TIP, COL_AMB_NA_ID, _
                  COL_AMB_VRSTA_KRETANJA, COL_AMB_STORNO_OD)

    For j = LBound(imena) To UBound(imena)
        If Len(AmbText(data(i, kol(CStr(imena(j)))))) > 0 Then
            RedDoticeKnjigu = True
            Exit Function
        End If
    Next j
End Function

' PRAZAN RED NIJE RED. Prazna Excel tabela ima jedan prazan red u DataBodyRange,
' pa bi ga svaka kapija inace prijavila kao kvar na praznoj knjizi. Meri se po SVIM
' kolonama koje knjiga uopste cita -- nijedno polje nije popunjeno.
Private Function RedPrazan(ByRef data As Variant, ByVal i As Long, _
                           ByRef kol As Object) As Boolean
    Dim k As Variant
    For Each k In kol.Keys
        If Len(AmbText(data(i, kol(CStr(k))))) > 0 Then Exit Function
    Next k

    RedPrazan = True
End Function

' UGOVOR ZAPISANOG REDA -- jedno mesto za SVE citaoce (saldo, obaveza,
' idempotencija, identitet, i storno u 10d).
'
' Zasto nad svakim citanjem a ne samo pri upisu: saldo ulazi u kapiju deficita, pa
' pokvaren ZAPISAN red menja odluku SLEDECEG upisa. Red uracunat pola-pola (izvor
' bez ID-a) razbija ocuvanje kolicine bez ijedne poruke -- partner dobije +20, a
' nijedan stvarni nalog ne dobije -20.
'
' Postojanje naloga u maticnoj tabeli se OVDE NE proverava: to je kapija UPISA
' (AmbNalogProblem u RequireAmbPrenos). Da citalac to radi, obrisan maticni red bi
' retroaktivno oborio svako citanje, a saldo bi postao kvadratan nad tabelom.
'
' STORNO SE MERI INVERZNO. Kontra-stav ima zamenjene strane, pa bi matrica za
' njegovu vrstu pala na ISPRAVNOM redu (inverz ULAZ_TUDJE je PARTNER -> GRANICA).
' Zamena argumenata JE provera inverza -- bez druge matrice i bez izuzetka.
Private Function KnjigaRedProblem(ByRef data As Variant, ByVal i As Long, _
                                  ByRef kol As Object) As String
    If Len(AmbText(data(i, kol(COL_AMB_ID)))) = 0 Then
        KnjigaRedProblem = "nema AmbID"
        Exit Function
    End If

    If Not IsDate(data(i, kol(COL_AMB_DATUM))) Then
        KnjigaRedProblem = "nema datum"
        Exit Function
    End If

    If Len(AmbText(data(i, kol(COL_AMB_DOK_TIP)))) = 0 Or _
       Len(AmbText(data(i, kol(COL_AMB_DOK_ID)))) = 0 Then
        KnjigaRedProblem = "nema identitet dokumenta"
        Exit Function
    End If

    If Not IsNumeric(data(i, kol(COL_AMB_KOLICINA))) Then
        KnjigaRedProblem = "kolicina nije broj"
        Exit Function
    End If

    Dim odTip As String, odID As String, naTip As String, naID As String
    Dim kolicina As Double, tipAmb As String, vrsta As String, p As String

    odTip = AmbText(data(i, kol(COL_AMB_OD_TIP)))
    odID = AmbText(data(i, kol(COL_AMB_OD_ID)))
    naTip = AmbText(data(i, kol(COL_AMB_NA_TIP)))
    naID = AmbText(data(i, kol(COL_AMB_NA_ID)))
    kolicina = CDbl(data(i, kol(COL_AMB_KOLICINA)))
    tipAmb = AmbText(data(i, kol(COL_AMB_TIP)))
    vrsta = AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA)))

    If Len(AmbText(data(i, kol(COL_AMB_STORNO_OD)))) > 0 Then
        p = modAmbalazaUgovor.AmbPrenosStrukturaProblem(naTip, naID, odTip, odID, _
                                                        kolicina, tipAmb, vrsta)
        If Len(p) > 0 Then KnjigaRedProblem = "storno: " & p
        Exit Function
    End If

    KnjigaRedProblem = modAmbalazaUgovor.AmbPrenosStrukturaProblem(odTip, odID, naTip, naID, _
                                                                  kolicina, tipAmb, vrsta)
End Function

' DVA REZIMA ISTOG SUDA. Citanje je fail-closed i staje na PRVOM kvaru -- to je
' ispravno za citaoca, jer racunanje dalje nad pokvarenom knjigom nema smisla.
' Audit mora suprotno: da vidi SVE nalaze i da NE obori ceo integritetni prolaz.
'
' Zato jedna implementacija, dva izlaza: bez kolekcije se DIZE (zatecen ugovor,
' nepromenjen), sa kolekcijom se SAKUPLJA. Druga kopija pravila bila bi tacno ono
' sto AMBALAZA.md 6.x zabranjuje -- ugovor je JEDNO mesto za sve citaoce.
Private Sub KnjigaKvar(ByVal nalazi As Collection, ByVal sourceName As String, _
                       ByVal red As Long, ByVal ambID As String, _
                       ByVal razlog As String)
    If nalazi Is Nothing Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, razlog
    End If

    nalazi.Add Array(CStr(red), ambID, razlog)
End Sub

' AUDIT KANONSKE KNJIGE -- isti sud kao citanje, sakupljen umesto dignut.
'
' Postoji jer modIntegritet pre ovoga NIJE gledao kanonsku knjigu: jedina
' ambalazna provera (Chk_B10, obrisana u 10e-1) bila je gejtovana starim tipovima
' dokumenta (ReversTipJe), pa je nad kanonskim redom cutala. Operater je mogao da pokrene
' integritet nad pokvarenom knjigom i dobije "nema nalaza" -- zato sto skener
' tu knjigu nije ni pogledao.
'
' Vraca Collection od Array(red, AmbID, razlog); prazna = knjiga je cista. Ne
' vraca niz: modIntegritet vec ima CollToArray, pa bi treca kopija konvertera
' bila treca stvar koja moze da divergira.
Public Function AmbKnjigaNalazi() As Collection
    Const SRC As String = "modAmbalaza.AmbKnjigaNalazi"

    Dim nalazi As Collection
    Set nalazi = New Collection
    Set AmbKnjigaNalazi = nalazi

    RequireKnjigaSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If Not IsArray(data) Then Exit Function

    Dim kol As Object
    Set kol = KnjigaIndeksi(SRC)
    KnjigaIntegritet data, kol, SRC, nalazi
End Function

' INTEGRITET KNJIGE -- jedan prolaz, pa SVAKI citalac dobija isti odgovor.
'
' Vraca mapu AmbID -> VrstaKretanja; ta mapa je i sama dokaz jedinstvenosti
' identiteta, a citaocu obaveze treba za vrstu ORIGINALA kad naidje na kontra-stav.
'
' JEDINSTVENOST AmbID-a NIJE SVOJSTVO REDA NEGO KNJIGE, pa je KnjigaRedProblem ne
' moze proveriti. Ranije je stajala samo u citaocu obaveze (review #400, drugi
' krug), pa je AmbSaldoNaloga sabirao dva reda sa istim stabilnim identitetom:
'
'   AMB-X  SpoljniSvet -> Stanica  100  NABAVKA
'   AMB-X  SpoljniSvet -> Stanica  100  NABAVKA       -> Stanica = 200
'
' Taj saldo ulazi u kapiju deficita, pa korumpiran identitet otvara izlaz BEZ
' pokrica. Za 10d je gore: StornoOd = AMB-X vise ne pokazuje na jedan original.
' Pravilo je isto kao za nalog (AmbNalogProblem): 0 pada, 1 prolazi, 2+ pada.
Private Function KnjigaIntegritet(ByRef data As Variant, ByRef kol As Object, _
                                  ByVal sourceName As String, _
                                  Optional ByVal nalazi As Collection) As Object
    Dim vrste As Object
    Set vrste = CreateObject("Scripting.Dictionary")
    Set KnjigaIntegritet = vrste

    If IsEmpty(data) Then Exit Function

    Dim i As Long, p As String, ambID As String
    For i = 1 To UBound(data, 1)
        If RedPrazan(data, i, kol) Then
            ' artefakt prazne tabele -- nije red
        ElseIf Not RedDoticeKnjigu(data, i, kol) Then
            ' TRI ISHODA, NE CETIRI (10e-3, AMBALAZA.md 6.13e): prazan, knjiga, KVAR.
            ' Do 10e-3 je red koji ne dotice knjigu smeo da bude VALJAN stari red i
            ' preskakao se. Stari model je obrisan iz kanona, pa takav red nije vise
            ' "drugi oblik" nego kvar: kolicina koju nosi bi tiho nestala iz svakog
            ' salda. Zatecena sveska sa redom starog oblika staje ovde, po imenu.
            KnjigaKvar nalazi, sourceName, i, AmbText(data(i, kol(COL_AMB_ID))), _
                       "Red " & CStr(i) & ": ne dotice knjigu -- nema nalog, vrstu " & _
                       "kretanja ni StornoOd (red starog modela ambalaze?)"
        Else
            p = KnjigaRedProblem(data, i, kol)
            ambID = AmbText(data(i, kol(COL_AMB_ID)))
            If Len(p) > 0 Then
                KnjigaKvar nalazi, sourceName, i, ambID, _
                           "Red knjige " & CStr(i) & ": " & p
            ElseIf vrste.Exists(ambID) Then
                KnjigaKvar nalazi, sourceName, i, ambID, _
                           "Dva reda knjige nose AmbID '" & ambID & "'."
            Else
                ' DODELA, ne .Add: Add bi na postojecem kljucu pukao sam od sebe i time
                ' bio SLUCAJNA druga brana ispred imenovane provere iznad. Dvosmerni
                ' dokaz je to i pokazao -- sa ugasenom imenovanom kapijom test je i
                ' dalje padao, samo sa tudjom porukom. Kapija sme biti samo jedna.
                vrste(ambID) = AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA)))
            End If
        End If
    Next i

    ' StornoOd mora da pokazuje na postojeci red -- DRUGI prolaz, jer original sme
    ' da stoji posle svog kontra-stava (knjiga je append-only, ne sortirana).
    Dim st As String
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            st = AmbText(data(i, kol(COL_AMB_STORNO_OD)))
            If Len(st) > 0 Then
                If Not vrste.Exists(st) Then
                    KnjigaKvar nalazi, sourceName, i, AmbText(data(i, kol(COL_AMB_ID))), _
                               "StornoOd '" & st & "' ne pokazuje na red knjige."
                End If
            End If
        End If
    Next i
End Function

' JEDAN ULAZ ZA CITANJE KNJIGE: indeksi I dokazan integritet, nikad odvojeno.
'
' Spojeni su namerno. Da `KnjigaIndeksi` ostane javno dostupan, nov citalac bi uzeo
' indekse i preskocio proveru -- a tacno taj oblik propusta je i bio P2: citalac
' obaveze je proveravao identitet, citalac salda nije.
'
' CENA JE JEDAN PROLAZ PO POZIVU CITAOCA, i to je svesno. Nad fixture-om se ne meri;
' nad pravom knjigom ce 10c morati da izmeri, a ako zaboli, odgovor je kes po
' transakciji -- ne slabije citanje.
Private Function KnjigaZaCitanje(ByRef data As Variant, ByVal sourceName As String, _
                                 ByRef outVrste As Object) As Object
    Dim kol As Object
    Set kol = KnjigaIndeksi(sourceName)
    Set outVrste = KnjigaIntegritet(data, kol, sourceName)
    Set KnjigaZaCitanje = kol
End Function

' Nalog kao kljuc, i neuredjen PAR naloga kao kljuc.
'
' Par je NEUREDJEN jer jedan revers sme da nosi i izdavanje i povrat prema istom
' partneru: Stanica -> K1 i K1 -> Stanica su isti poslovni par.
Private Function NalogKljuc(ByVal tip As String, ByVal id As String) As String
    NalogKljuc = Trim$(tip) & ":" & Trim$(id)
End Function

Private Function ParKljuc(ByVal tipA As String, ByVal idA As String, _
                          ByVal tipB As String, ByVal idB As String) As String
    Dim a As String, b As String
    a = NalogKljuc(tipA, idA)
    b = NalogKljuc(tipB, idB)

    If StrComp(a, b, vbTextCompare) <= 0 Then
        ParKljuc = a & " <-> " & b
    Else
        ParKljuc = b & " <-> " & a
    End If
End Function

Private Function IstiNalog(ByVal tipA As String, ByVal idA As String, _
                           ByVal tipB As String, ByVal idB As String) As Boolean
    IstiNalog = (StrComp(Trim$(tipA), Trim$(tipB), vbTextCompare) = 0) And _
                (StrComp(Trim$(idA), Trim$(idB), vbTextCompare) = 0)
End Function

' DAN, ne trenutak: knjiga ambalaze je dnevna (broj reversa je niz po danu), a
' celija moze da nosi i vreme. Poredjenje po trenutku bi isti dogadjaj unet dva
' puta istog dana prijavilo kao dva.
Private Function IstiDan(ByVal a As Variant, ByVal b As Date) As Boolean
    If Not IsDate(a) Then Exit Function
    IstiDan = (Int(CDate(a)) = Int(b))
End Function

' SALDO NALOGA IZ KNJIGE: +kolicina kad je nalog ODREDISTE, -kolicina kad je
' IZVOR. Storno je kontra-stav sa zamenjenim stranama, pa se gasi istim pravilom
' -- bez posebnog slucaja i bez kolone Stornirano.
'
' NEMA "On Error -> vrati 0". Nekadasnji GetStanicaAmbSaldo je tako radio i to je
' fail-open koji ovde ne sme da postoji: na ovaj broj se oslanja kapija deficita,
' pa bi progutana greska proizvela upis BEZ pokrica -- tj. negativan saldo.
Public Function AmbSaldoNaloga(ByVal tip As String, ByVal id As String, _
                               ByVal tipAmb As String) As Double
    Const SRC As String = "modAmbalaza.AmbSaldoNaloga"

    Dim p As String
    p = modAmbalazaUgovor.AmbNalogProblem(tip, id)
    If Len(p) > 0 Then Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, p

    If Len(Trim$(tipAmb)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Saldo se vodi PO TIPU ambalaze -- tip je obavezan."
    End If

    RequireKnjigaSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cTipA As Long, cKol As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)

    Dim i As Long, saldo As Double
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 Then
                If IstiNalog(AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), tip, id) Then
                    saldo = saldo + CDbl(data(i, cKol))
                ElseIf IstiNalog(AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), tip, id) Then
                    saldo = saldo - CDbl(data(i, cKol))
                End If
            End If
        End If
    Next i

    AmbSaldoNaloga = saldo
End Function


' SALDO SVIH NALOGA JEDNOG TIPA, U JEDNOM PROLAZU.
'
' AmbSaldoNaloga odgovara na "koliko drzi OVAJ nalog, od OVOG tipa gajbe".
' Izvestaj pita obrnuto: "koliko drzi SVAKI kooperant, preko svih tipova" -- a to
' je N prolaza kroz knjigu kad se zove u petlji. Zato jedan prolaz, ali ISTO
' pravilo znaka (ODREDISTE +, IZVOR -) i ISTI lifecycle (RedDoticeKnjigu).
'
' Postoji zato da izvestaj ne bi nosio svoju kopiju tog pravila. Do 10c ju je
' nosio: ReportSaldoOM je imao svoj prolaz, svoj Select Case po Smer-u i svoj
' ExcludeStornirano -- tri stvari koje su vlasnistvo knjige.
'
' Red sme da ima isti TIP na obe strane (Stanica -> Stanica, dve stanice). Oba
' doprinosa se tada primenjuju, svaki svom ID-u. Isti NALOG na obe strane ne
' postoji -- to zabranjuje AMB-INV-02 -- pa se nijedan nalog ne anulira sam.
'
' Suma je preko SVIH tipova ambalaze, kao i kolona koju izvestaj prikazuje. Kome
' treba po tipu, zove GetAmbalazeStanje (isto pravilo, drugi oblik).
' samoPoslovno: bez tehnickih redova (AMB-POSL-01). datumDo > 0: stanje NA DAN --
' redovi kasnijeg datuma ne ulaze (pregled za proslost ne sme da nosi danasnje
' stanje). protivTip/protivID: samo kretanja sa TIM nalogom -- svaki red imenuje
' obe strane, pa je "saldo kooperanta prema stanici X" podatak, ne alokacija.
' Red prema granici (otpis sa naloga kooperanta, #445) pripada nalogu koji je
' vlasnik ambalaznog dokumenta (RedPripadaProtivnalogu).
' Bez opcija je ovo zatecen fizicki, danasnji, globalni saldo.
Public Function AmbSaldoPoNalogu(ByVal tip As String, _
                                 Optional ByVal samoPoslovno As Boolean = False, _
                                 Optional ByVal datumDo As Date = 0, _
                                 Optional ByVal protivTip As String = "", _
                                 Optional ByVal protivID As String = "") As Object
    Const SRC As String = "modAmbalaza.AmbSaldoPoNalogu"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set AmbSaldoPoNalogu = res

    If Len(Trim$(tip)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, "Tip naloga je obavezan."
    End If

    RequireKnjigaSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim kolIdx As Object, vrsteDok As Object
    Set kolIdx = KnjigaZaCitanje(data, SRC, vrsteDok)

    Dim cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long, cKol As Long
    cOdTip = kolIdx(COL_AMB_OD_TIP)
    cOdID = kolIdx(COL_AMB_OD_ID)
    cNaTip = kolIdx(COL_AMB_NA_TIP)
    cNaID = kolIdx(COL_AMB_NA_ID)
    cKol = kolIdx(COL_AMB_KOLICINA)

    Dim cDat As Long, saProtivnim As Boolean
    cDat = kolIdx(COL_AMB_DATUM)
    saProtivnim = (Len(Trim$(protivTip)) > 0)

    Dim cDokTip As Long, cDokID As Long, vlasnik As Object
    If saProtivnim Then
        cDokTip = kolIdx(COL_AMB_DOK_TIP)
        cDokID = kolIdx(COL_AMB_DOK_ID)
        Set vlasnik = AmbDokVlasnikMapa()
    End If

    ' VBA Or ne prekida izracunavanje: predikat protivnaloga se zove SAMO kad je
    ' protivnalog zadat (bez njega mapa vlasnika i kolone dokumenta ne postoje).
    Dim i As Long, nid As String, pripada As Boolean
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kolIdx) Then
            If samoPoslovno Then
                If AmbRedJeTehnicki(data, i, kolIdx, vrsteDok) Then GoTo SledeciRedSalda
            End If
            If datumDo <> 0 Then
                If Int(CDbl(CDate(data(i, cDat)))) > Int(CDbl(datumDo)) Then GoTo SledeciRedSalda
            End If
            If StrComp(AmbText(data(i, cNaTip)), Trim$(tip), vbTextCompare) = 0 Then
                pripada = True
                If saProtivnim Then pripada = RedPripadaProtivnalogu(data, i, AmbText(data(i, cOdTip)), _
                                                 AmbText(data(i, cOdID)), cDokTip, cDokID, _
                                                 vlasnik, protivTip, protivID)
                If pripada Then
                    nid = AmbText(data(i, cNaID))
                    If Len(nid) > 0 Then
                        If Not res.Exists(nid) Then res.Add nid, 0&
                        res(nid) = CLng(res(nid)) + CLng(data(i, cKol))
                    End If
                End If
            End If
            If StrComp(AmbText(data(i, cOdTip)), Trim$(tip), vbTextCompare) = 0 Then
                pripada = True
                If saProtivnim Then pripada = RedPripadaProtivnalogu(data, i, AmbText(data(i, cNaTip)), _
                                                 AmbText(data(i, cNaID)), cDokTip, cDokID, _
                                                 vlasnik, protivTip, protivID)
                If pripada Then
                    nid = AmbText(data(i, cOdID))
                    If Len(nid) > 0 Then
                        If Not res.Exists(nid) Then res.Add nid, 0&
                        res(nid) = CLng(res(nid)) - CLng(data(i, cKol))
                    End If
                End If
            End If
        End If
SledeciRedSalda:
    Next i
End Function

' DA LI RED PRIPADA PROTIVNALOGU (AMB-POSL-02).
'
' Pripada kad je druga strana reda bas taj nalog. Red prema GRANICI (SpoljniSvet)
' nema drugu realnu stranu: otpis sa naloga kooperanta (#445) ide Kooperant ->
' SpoljniSvet, a stanica je samo VLASNIK dokumenta -- broj iz njenog niza
' (AMB-10-ODL-34 c), isti OM koji stoji i na zapisniku. Takav red pripada nalogu
' koji je vlasnik ambalaznog dokumenta, i nijednom drugom; bez toga bi pregled
' po OM nosio otpisane gajbe kao dug.
Private Function RedPripadaProtivnalogu(ByRef data As Variant, ByVal i As Long, _
                                        ByVal drugaTip As String, ByVal drugaID As String, _
                                        ByVal cDokTip As Long, ByVal cDokID As Long, _
                                        ByVal vlasnik As Object, _
                                        ByVal protivTip As String, _
                                        ByVal protivID As String) As Boolean
    Dim k As String, v As Variant
    If IstiNalog(drugaTip, drugaID, protivTip, protivID) Then
        RedPripadaProtivnalogu = True
    ElseIf StrComp(Trim$(drugaTip), AMB_NALOG_SPOLJNI, vbTextCompare) = 0 Then
        If StrComp(AmbText(data(i, cDokTip)), DOK_TIP_AMBALAZA_DOKUMENT, vbTextCompare) = 0 Then
            k = AmbText(data(i, cDokID))
            If vlasnik.Exists(k) Then
                v = vlasnik(k)
                RedPripadaProtivnalogu = IstiNalog(CStr(v(0)), CStr(v(1)), protivTip, protivID)
            End If
        End If
    End If
End Function

' AmbDokID -> Array(BrojOwnerTip, BrojOwnerID), jednim prolazom kroz zaglavlja.
' Storniran dokument ostaje u mapi: njegov original i kontra-stav pripadaju
' istom vlasniku, pa se medjusobno potiru.
Private Function AmbDokVlasnikMapa() As Object
    Const SRC As String = "modAmbalaza.AmbDokVlasnikMapa"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set AmbDokVlasnikMapa = res

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA_DOKUMENT)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cOwTip As Long, cOwID As Long, i As Long, k As String
    cID = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, SRC)
    cOwTip = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_TIP, SRC)
    cOwID = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_ID, SRC)
    For i = 1 To UBound(data, 1)
        k = AmbText(data(i, cID))
        If Len(k) > 0 Then res(k) = Array(AmbText(data(i, cOwTip)), AmbText(data(i, cOwID)))
    Next i
End Function


' SVA AKTIVNA KRETANJA JEDNOG NALOGA, SA ZNAKOM -- u jednom prolazu.
'
' Kartice su REDNI izvestaji: treba im datum, dokument i doprinos TOG reda, a ne
' saldo. Do 10c su to vadile same, kroz Smer -- pa je svaka nosila SVOJU kopiju
' pravila znaka, i svaka svoj lifecycle (ExcludeStornirano).
'
' Vraca 2D niz (1 To n, 1 To 11):
'   1 = Datum   2 = DokumentID   3 = DokumentTip   4 = TipAmbalaze
'   5 = Kolicina SA ZNAKOM: pozitivna kad nalog PRIMA, negativna kad daje (6.8)
'   6 = Otkazano: True za kontra-stav I za original na koji pokazuje
'   7 = ProtivpartnerTip   8 = ProtivpartnerID
'   9 = VrstaKretanja   10 = AmbID   11 = StornoOd (prazno kod originala)
'
' Kolone 6-8 su IDENTITET, ne prikaz: ko je druga strana i da li par jos
' stoji. Kako se to imenuje coveku ostaje na citaocu.
'
' Kolone 9-11 su dodate na KRAJ (KART-HRON-01): kartica po njima redja i
' imenuje kretanja istog dokumenta -- vrsta kaze ZASTO je kretanje nastalo, a
' StornoOd koji red je kontra-stav. Citaoci kolona 1-8 se ne menjaju.
'
' Prazno kad nalog nema ni jedno kretanje -- isti ugovor kao GetAmbalazeStanje, pa
' pozivaoci koji vrte IsArray ostaju nedirnuti.
'
' Kontra-stav se vraca kao OBICAN red, sa svojim znakom: kartica mora da pokaze i
' storno, a ne da ga sakrije. Saldo se tako sam vraca na pocetno.
'
' samoPoslovno (AMB-POSL-01): bez tehnickih redova (AmbRedJeTehnicki) -- ono sto
' citaju kartice. Bez njega je ovo fizicki tok, koji citaju pregled kretanja i
' provere.
Public Function AmbKretanjaNaloga(ByVal tip As String, ByVal id As String, _
                                  Optional ByVal samoPoslovno As Boolean = False) As Variant
    Const SRC As String = "modAmbalaza.AmbKretanjaNaloga"

    Dim p As String
    p = modAmbalazaUgovor.AmbNalogProblem(tip, id)
    If Len(p) > 0 Then Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, p

    RequireKnjigaSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim kolIdx As Object, vrsteDok As Object
    Set kolIdx = KnjigaZaCitanje(data, SRC, vrsteDok)

    Dim cDat As Long, cDokID As Long, cDokTip As Long, cTipA As Long, cKol As Long
    Dim cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cStorno As Long, cAmbID As Long, cVK As Long
    cStorno = kolIdx(COL_AMB_STORNO_OD)
    cVK = kolIdx(COL_AMB_VRSTA_KRETANJA)
    cAmbID = kolIdx(COL_AMB_ID)
    cDat = kolIdx(COL_AMB_DATUM)
    cDokID = kolIdx(COL_AMB_DOK_ID)
    cDokTip = kolIdx(COL_AMB_DOK_TIP)
    cTipA = kolIdx(COL_AMB_TIP)
    cKol = kolIdx(COL_AMB_KOLICINA)
    cOdTip = kolIdx(COL_AMB_OD_TIP)
    cOdID = kolIdx(COL_AMB_OD_ID)
    cNaTip = kolIdx(COL_AMB_NA_TIP)
    cNaID = kolIdx(COL_AMB_NA_ID)

    ' Prvo se skupe POZICIJE pa se niz dimenzionise jednom: ReDim Preserve nad 2D
    ' nizom menja samo poslednju dimenziju, pa rast po redovima nije moguc.
    Dim i As Long, znak As Long

    ' OTKAZANI PAROVI, jednim prolazom: kontra-stav nosi StornoOd -> AmbID
    ' originala, pa su otkazana OBA reda.
    '
    ' Ovo je KOLONA, a ne filter, jer se dva citaoca iste knjige legitimno
    ' razlikuju (6.8): pregled KRETANJA storniran dokument skriva, a KARTICA ga
    ' prikazuje -- storno je i sam dogadjaj koji operater mora da vidi. Odluku
    ' zato nosi pozivalac, a knjiga samo kaze sta je otkazano.
    Dim otkazani As Object
    Set otkazani = CreateObject("Scripting.Dictionary")
    otkazani.CompareMode = vbTextCompare

    Dim so As String
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kolIdx) Then
            so = AmbText(data(i, cStorno))
            If Len(so) > 0 Then
                otkazani(so) = True
                otkazani(AmbText(data(i, cAmbID))) = True
            End If
        End If
    Next i

    Dim poz As Collection
    Set poz = New Collection

    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kolIdx) Then
            znak = 0
            If IstiNalog(AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), tip, id) Then
                znak = 1
            ElseIf IstiNalog(AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), tip, id) Then
                znak = -1
            End If
            If znak <> 0 And samoPoslovno Then
                If AmbRedJeTehnicki(data, i, kolIdx, vrsteDok) Then znak = 0
            End If
            If znak <> 0 Then poz.Add Array(i, znak)
        End If
    Next i

    If poz.count = 0 Then Exit Function

    Dim res() As Variant
    ReDim res(1 To poz.count, 1 To 11)

    Dim n As Long, ri As Long, par As Variant
    For n = 1 To poz.count
        par = poz(n)
        ri = CLng(par(0))
        znak = CLng(par(1))
        res(n, 1) = data(ri, cDat)
        res(n, 2) = AmbText(data(ri, cDokID))
        res(n, 3) = AmbText(data(ri, cDokTip))
        res(n, 4) = AmbText(data(ri, cTipA))
        res(n, 5) = znak * CDbl(data(ri, cKol))
        res(n, 6) = otkazani.Exists(AmbText(data(ri, cAmbID)))
        ' PROTIVPARTNER: kad nalog PRIMA, druga strana je IZVOR; kad daje,
        ' ODREDISTE. AMB-INV-02 (OdNalog <> NaNalog) garantuje da to nije on sam.
        If znak > 0 Then
            res(n, 7) = AmbText(data(ri, cOdTip))
            res(n, 8) = AmbText(data(ri, cOdID))
        Else
            res(n, 7) = AmbText(data(ri, cNaTip))
            res(n, 8) = AmbText(data(ri, cNaID))
        End If
        res(n, 9) = AmbText(data(ri, cVK))
        res(n, 10) = AmbText(data(ri, cAmbID))
        res(n, 11) = AmbText(data(ri, cStorno))
    Next n

    AmbKretanjaNaloga = res
End Function


' KAKO SE AMBALAZNI DOKUMENT ZOVE: poslovni broj i vrsta, po AmbDokID -- u jednom
' prolazu kroz zaglavlja.
'
' Knjiga nosi (DokumentTip, DokumentID) i to je ispravno: ledger ne sme da zna
' kako se dokument PRIKAZUJE. Ali AmbDokID je opaque "ADK-<hex>", pa bi kartica
' bez ovog prevoda operateru pokazala tehnicki identitet kao broj dokumenta, i
' genericki "AmbalazaDokument" umesto "Revers" (review 08.10.2026, P2 #1).
'
' Mapa, a ne LookupValue po redu: kartica ima N redova, a zaglavlja su jedna
' tabela -- isti razlog zbog koga postoji i AmbDokRedMapa.
'
' VREDNOST JE NIZ (broj, vrsta, datum), ne spojen string.
'
' Prvo izdanje je pakovalo "broj|vrsta|datum" uz komentar da "broj ne sadrzi
' '|'". To je bila PRETPOSTAVKA napisana kao cinjenica: kupcev revers (ODL-23)
' nosi broj sa KUPCEVOG dokumenta, a nijedna kapija ne filtrira karaktere -- pa
' je "KUP|R-17" legalan broj. Parser bi tada procitao broj="KUP",
' vrsta="R-17", datum=vrsta, i to ne bi PUKLO nego tiho promenilo i karticu i
' PAPIR (review 08.10.2026, P2).
'
' Datum se cuva kao SIROVA vrednost celije, ne kao tekst: tekstualni datum je
' vec jednom pojeo stanicu iz opisa, jer je IsDate nad "8.10.2026." False.
' AKTIVNA ZAGLAVLJA -- skup AmbDokID-eva koji NISU stornirani.
'
' Ide kroz modHelpers.ExcludeStornirano, a ne kroz novo poredjenje: semantika
' vrednosti "Da" vec zivi u DVE privatne kopije (modStorno.IsStorniranoValue i
' modProductionHealthCheck.IsStorniranoValue). Treca bi bila treca stvar koja moze
' da divergira -- isti rod greske koji je HARD/SOFT paritet vec naplatio.
'
' ZASTO SKUP, A NE ZASTAVICA U AmbDokPrikazMapa: tu mapu cita i KARTICA, koja po
' AMBALAZA.md 6.8 storniran dokument mora da PRIKAZE. Filtriranje te mape bi joj
' pojelo poslovni broj bas za povucene dokumente -- a ta asimetrija (pregled
' skriva, kartica prikazuje) je tvrdnja, ne navika.
Public Function AmbDokAktivanSkup() As Object
    Const SRC As String = "modAmbalaza.AmbDokAktivanSkup"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set AmbDokAktivanSkup = res

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA_DOKUMENT)
    data = ExcludeStornirano(data, TBL_AMBALAZA_DOKUMENT)
    If Not IsArray(data) Then Exit Function

    Dim cID As Long
    cID = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, SRC)

    Dim i As Long, k As String
    For i = 1 To UBound(data, 1)
        k = AmbText(data(i, cID))
        If Len(k) > 0 Then res(k) = True
    Next i
End Function

Public Function AmbDokPrikazMapa() As Object
    Const SRC As String = "modAmbalaza.AmbDokPrikazMapa"

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res.CompareMode = vbTextCompare
    Set AmbDokPrikazMapa = res

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA_DOKUMENT)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cBroj As Long, cVrsta As Long, cDat As Long
    cID = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, SRC)
    cBroj = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ, SRC)
    cVrsta = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_VRSTA, SRC)
    cDat = RequireColumnIndex(TBL_AMBALAZA_DOKUMENT, COL_AMBD_DATUM, SRC)

    Dim i As Long, k As String
    For i = 1 To UBound(data, 1)
        k = AmbText(data(i, cID))
        If Len(k) > 0 Then
            res(k) = Array(AmbText(data(i, cBroj)), _
                           AmbText(data(i, cVrsta)), _
                           data(i, cDat))
        End If
    Next i
End Function

' OBAVEZA FIRME PREMA PARTNERU -- izvedena iz iste knjige, bez ijedne mutabilne
' kolone, preko DOPRINOSA po dogadjaju (AMB-INV-09).
'
' Prosta razlika dve sume NIJE dovoljna: storno ULAZA upisuje kontra-stav koji
' anulira fizicko stanje, a razlika dve sume ostavila bi obavezu da visi.
'
' Vrsta ORIGINALA se cita iz reda na koji StornoOd pokazuje, a ne iz samog
' kontra-stava. Razlika je bitna: inace bi 10d mogao da ugasi obavezu upisujuci
' storno sa drugom vrstom, i nijedna provera to ne bi videla. StornoOd koji ne
' pokazuje nigde je kvar, ne nula.
Public Function AmbObavezaPartneru(ByVal tip As String, ByVal id As String, _
                                   ByVal tipAmb As String) As Double
    Const SRC As String = "modAmbalaza.AmbObavezaPartneru"

    Dim p As String
    p = modAmbalazaUgovor.AmbNalogProblem(tip, id)
    If Len(p) > 0 Then Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, p

    If Len(Trim$(tipAmb)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Obaveza se vodi PO TIPU ambalaze -- tip je obavezan."
    End If

    ' OBAVEZA POSTOJI SAMO PREMA NALOGU KOJI MOZE DA PRIMI TUDJU AMBALAZU.
    '
    ' Bez ove kapije funkcija vraca broj i za stanicu: njeni VRACANJE_TUDJE redovi
    ' daju -N, pa bi "firma duguje stanici -70" izgledalo kao podatak. Klasa se
    ' CITA iz iste matrice koja definise pokrice (ULAZ_TUDJE: GRANICA -> PARTNER)
    ' -- dug nastaje tacno tim dogadjajem, pa mu je klasa ista po konstrukciji.
    Dim klasaDuga As String
    klasaDuga = modAmbalazaUgovor.AmbPokriceKlasa()
    If Len(klasaDuga) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Klasa duga nije definisana u matrici."
    End If

    If Not modAmbalazaUgovor.AmbNalogUKlasi(klasaDuga, tip) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Obaveza se ne vodi prema nalogu " & Trim$(tip) & ": dug nastaje " & _
                  "ulazom tudje ambalaze, a on ide samo na " & klasaDuga & "."
    End If

    RequireKnjigaSchema SRC

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cTipA As Long, cKol As Long, cSt As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    cID = kol(COL_AMB_ID)
    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)
    cSt = kol(COL_AMB_STORNO_OD)

    ' Mapa AmbID -> vrsta dolazi iz KAPIJE INTEGRITETA, ne iz prolaza ovog citaoca:
    ' jedinstvenost identiteta je svojstvo KNJIGE, pa mora vaziti i za citaoca salda
    ' (review #400, drugi krug).
    Dim i As Long
    Dim obaveza As Double, stornoOd As String, vrstaOrig As String
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 Then
                If IstiNalog(AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), tip, id) Or _
                   IstiNalog(AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), tip, id) Then

                    stornoOd = AmbText(data(i, cSt))
                    vrstaOrig = ""
                    If Len(stornoOd) > 0 Then
                        If Not vrste.Exists(stornoOd) Then
                            Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                                      "StornoOd '" & stornoOd & "' ne pokazuje na red knjige."
                        End If
                        vrstaOrig = CStr(vrste(stornoOd))
                    End If

                    obaveza = obaveza + modAmbalazaUgovor.AmbDoprinosObavezi( _
                                            AmbText(data(i, cVK)), _
                                            CDbl(data(i, cKol)), _
                                            vrstaOrig)
                End If
            End If
        End If
    Next i

    AmbObavezaPartneru = obaveza
End Function

' DEFICIT: koliko bi IZVORNOM nalogu falilo da se prenos upise bez pokrica.
'
' GRANICA (SpoljniSvet) nema fizicko stanje -- ona je izvor i ponor opticaja
' (6.3) -- pa prenos IZ nje nikad nije u manjku. Za svaki realan nalog deficit je
' obican racun, i javan je zato sto UI sme da ga PRIKAZE pre upisa.
Public Function AmbDeficitZaPrenos(ByVal odTip As String, ByVal odID As String, _
                                   ByVal tipAmb As String, _
                                   ByVal kolicina As Double) As Double
    If Not modAmbalazaUgovor.AmbNalogUKlasi(AMB_KLASA_REALAN, odTip) Then Exit Function

    Dim saldo As Double
    saldo = AmbSaldoNaloga(odTip, odID, tipAmb)
    If kolicina > saldo Then AmbDeficitZaPrenos = kolicina - saldo
End Function

' ============================================================
' AMBALAZNI DOKUMENT -- zaglavlje (AMB-10-DOK)
' ============================================================

Public Function NoviAmbDokID() As String
    Const SRC As String = "modAmbalaza.NoviAmbDokID"

    NoviAmbDokID = NewEntityID("ADK-")

    If Len(Trim$(NoviAmbDokID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "NewEntityID nije vratio AmbDokID."
    End If
End Function

' Zaglavlje dokumenta koji nosi dogadjaje bez sopstvenog poslovnog dokumenta
' (revers, nabavka, otpis). Vraca AmbDokID -- on ide u tblAmbalaza.DokumentID,
' uz DokumentTIP = DOK_TIP_AMBALAZA_DOKUMENT.
'
' Red se gradi PO IMENU KOLONE (SetRowValueByColumn), ne golim Array(...):
' AppendRow pise poziciono, pa bi kolona ubacena u sredinu tiho poslala sve iza
' sebe u pogresna polja.
' Zaglavlje trazi SVOJU tabelu u snapshotu, ne knjigu: pozivalac koji kreira
' dokument I redove mora da snapshotuje OBE, inace rollback vraca pola
' dokumenta -- zaglavlje bez redova ili redove bez zaglavlja.
Public Function UpisiAmbDokument(ByVal tx As clsTransaction, _
                                 ByVal vrsta As String, ByVal broj As String, _
                                 ByVal datum As Date, _
                                 ByVal brojOwnerTip As String, _
                                 ByVal brojOwnerID As String, _
                                 Optional ByVal napomena As String = "") As String
    Const SRC As String = "modAmbalaza.UpisiAmbDokument"

    RequireAmbTxVlasnistvo tx, TBL_AMBALAZA_DOKUMENT, SRC
    RequireAmbDokSchema SRC
    modAmbalazaUgovor.RequireAmbDok vrsta, broj, datum, brojOwnerTip, brojOwnerID, SRC

    ' ZAUZETOST BROJA JE KAPIJA PISCA, NE UI-ja. RequireAmbDok sudi OBLIK
    ' (vrsta, neprazan broj, datum, klasa vlasnika); bez ovoga su dva poziva sa
    ' istim rucno prosledjenim brojem davala dva AmbDokID-a i jedan poslovni broj
    ' u istom nizu (review 03.10.2026, P2 #2). Opseg je kanonski:
    ' (BrojOwnerTip, BrojOwnerID, dan) -- isti koji generator koristi.
    Dim zauzeo As String
    zauzeo = modBrojevi.AmbDokBrojZauzet(brojOwnerTip, brojOwnerID, datum, broj)
    If Len(zauzeo) > 0 Then
        Err.Raise AMB_ERR_IDENTITET, SRC, _
                  "Broj '" & Trim$(broj) & "' je u nizu " & Trim$(brojOwnerTip) & _
                  " '" & Trim$(brojOwnerID) & "' tog dana vec zauzet (dokument " & _
                  zauzeo & "). Storno ne oslobadja broj."
    End If

    Dim vrstaK As String, ownerK As String
    vrstaK = modAmbalazaUgovor.AmbDokVrstaKanon(vrsta)
    ownerK = modAmbalazaUgovor.AmbNalogTipKanon(brojOwnerTip)
    If Len(vrstaK) = 0 Or Len(ownerK) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                  "Kanonski zapis nije nadjen posle prosle kapije -- kvar ugovora."
    End If

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_AMBALAZA_DOKUMENT)
    If colCount <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Ne mogu da odredim broj kolona za tblAmbalazaDokument."
    End If

    Dim novID As String
    novID = NoviAmbDokID()

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, novID, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_VRSTA, vrstaK, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ, Trim$(broj), SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_DATUM, datum, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_TIP, ownerK, SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_BROJ_OWNER_ID, Trim$(brojOwnerID), SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_AMBD_NAPOMENA, Trim$(napomena), SRC
    SetRowValueByColumn rowData, TBL_AMBALAZA_DOKUMENT, COL_STORNIRANO, "", SRC

    Dim rowIdx As Long
    rowIdx = AppendRow(TBL_AMBALAZA_DOKUMENT, rowData)
    If rowIdx <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "AppendRow nije uspeo za tblAmbalazaDokument."
    End If

    ' Dokument koji smo upravo napravili pripada OVOJ transakciji, pa ga ona i
    ' vezuje: sledeci PrenesiAmbalazu nad njim ne trazi nista od pozivaoca.
    tx.BindSourceDocument DOK_TIP_AMBALAZA_DOKUMENT, novID

    UpisiAmbDokument = novID
End Function

' Zaglavlje u celini: vrsta + vlasnik numerickog niza.
'
' Postoji jer AMB-10-ODL-10 trazi da se BrojOwner uporedi sa stranom kretanja,
' a to se ne moze iz vrste same. AmbDokVrstaZaID ostaje javan (testovi ga
' koriste) i poziva ovo, da provera storniranog i postojanja stoji na JEDNOM
' mestu.
' Vlasnik broja ROBNOG dokumenta, iz zatvorene mape (AMB-10-ODL-22).
'
' Prazno za tip koji ga ne objavljuje -- i to je fail-closed, jer obrnuta
' kapija ODL-10 bez vlasnika broja odbija povrat od kupca.
'
' Cita se IZ TABELE dokumenta, ne iz argumenta pozivaoca: tabela je cinjenica,
' a argument bi bio tvrdnja pisca o sebi (isti razlog kao u modStornoZurnal).
Public Sub AmbRobniZaglavlje(ByVal dokTip As String, ByVal dokID As String, _
                             ByRef outOwnerTip As String, _
                             ByRef outOwnerID As String)
    Const SRC As String = "modAmbalaza.AmbRobniZaglavlje"

    outOwnerTip = ""
    outOwnerID = ""
    If Len(Trim$(dokID)) = 0 Then Exit Sub

    Dim mapa As Variant, i As Long, red As Variant
    mapa = modAmbalazaUgovor.AmbRobniVlasniciBroja()
    For i = LBound(mapa) To UBound(mapa)
        red = mapa(i)
        If StrComp(Trim$(dokTip), CStr(red(0)), vbTextCompare) = 0 Then
            outOwnerTip = CStr(red(3))
            outOwnerID = Trim$(NzToText(LookupValue(CStr(red(1)), CStr(red(2)), _
                                                    dokID, CStr(red(4)))))
            If Len(outOwnerID) = 0 Then
                Err.Raise AMB_ERR_IDENTITET, SRC, _
                          "Dokument " & Trim$(dokTip) & " " & Trim$(dokID) & _
                          " ne nosi vlasnika svog broja (" & CStr(red(4)) & _
                          "). AMB-10-ODL-22 trazi da ga robni dokument objavi."
            End If
            Exit Sub
        End If
    Next i
End Sub

Public Sub AmbDokZaglavlje(ByVal ambDokID As String, _
                           ByRef vrsta As String, _
                           ByRef brojOwnerTip As String, _
                           ByRef brojOwnerID As String)
    Const SRC As String = "modAmbalaza.AmbDokZaglavlje"

    RequireAmbDokSchema SRC
    RequireTacnoJedan TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID), _
                      "Ambalazni dokument", SRC

    Dim st As Variant
    st = LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID), COL_STORNIRANO)
    If Len(AmbText(st)) > 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Ambalazni dokument '" & Trim$(ambDokID) & "' je storniran."
    End If

    vrsta = AmbText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                Trim$(ambDokID), COL_AMBD_VRSTA))
    brojOwnerTip = AmbText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                       Trim$(ambDokID), COL_AMBD_BROJ_OWNER_TIP))
    brojOwnerID = AmbText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                      Trim$(ambDokID), COL_AMBD_BROJ_OWNER_ID))

    If Len(vrsta) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                  "Ambalazni dokument '" & Trim$(ambDokID) & "' nema vrstu."
    End If
End Sub

' Vrsta posla sa zaglavlja, za kapiju AmbDokDozvoljavaKretanje.
'
' STORNIRAN DOKUMENT NE PRIMA NOVA KRETANJA: zaglavlje sme da nosi Stornirano
' (ono nije knjiga), a dopisivanje na ponisten dokument bilo bi kretanje bez
' ziveg povoda.
Public Function AmbDokVrstaZaID(ByVal ambDokID As String) As String
    Dim ot As String, oi As String
    AmbDokZaglavlje ambDokID, AmbDokVrstaZaID, ot, oi
End Function

' ============================================================
' UPIS JEDNOG REDA KNJIGE -- jedini put do tblAmbalaza u novom modelu
' ============================================================
'
' AMB-INV-04 se proverava OVDE, nad svakim upisanim redom, a ne samo u
' PrenesiAmbalazu: pokrice deficita i ostatak podele pisac generise SAM, pa bi ih
' provera na ulazu promasila.
'
' Kljuc je (DokumentTIP, DokumentID, VrstaKretanja, TipAmbalaze) i vazi za
' ORIGINALE. Kontra-stav (StornoOd <> "") ima isti kljuc kao original po
' konstrukciji, pa se na njega ne primenjuje -- njegovu jedinstvenost drzi
' AMB-INV-06, u 10d.
' AMB-INV-08 ZIVI OVDE, jer kroz ovu funkciju prolazi SVAKI red knjige -- i
' pokrice deficita i ostatak podele, ne samo trazeni prenos. Kopija kapije na
' ulazu u PrenesiAmbalazu bila bi placebo: jezgro bi odbilo isti upis i bez nje,
' pa je nijedna sabotaza ne bi mogla oboriti. Isti razlog je tamo vec zapisan za
' identitet dokumenta.
Private Function UpisiRedKnjige(ByVal tx As clsTransaction, _
                                ByVal datum As Date, ByVal tipAmb As String, _
                                ByVal kolicina As Double, _
                                ByVal odTip As String, ByVal odID As String, _
                                ByVal naTip As String, ByVal naID As String, _
                                ByVal dokTip As String, ByVal dokID As String, _
                                ByVal vrsta As String, ByVal stornoOd As String, _
                                ByVal sourceName As String) As String
    RequireAmbTxVlasnistvo tx, TBL_AMBALAZA, sourceName
    RequireAmbTxIzvorniDokument tx, dokTip, dokID, sourceName
    RequireKnjigaSchema sourceName

    ' KONTRA-STAV SE PROVERAVA U OBRNUTOM SMERU -- isto pravilo koje citalac
    ' vec ima (KnjigaRedProblem). Kontra-stav nosi zamenjene Od i Na, pa bi ga
    ' prava provera odbila na matrici klasa: IZDATA_PRAZNA trazi SOPSTVENI kao
    ' izvor, a kontra-stav tu ima partnera. Pravilo stoji na JEDNOM mestu u
    ' pisacu, da ne moze da se razidje sa citaocem.
    If Len(Trim$(stornoOd)) = 0 Then
        modAmbalazaUgovor.RequireAmbPrenos odTip, odID, naTip, naID, kolicina, tipAmb, vrsta, sourceName
    Else
        modAmbalazaUgovor.RequireAmbPrenos naTip, naID, odTip, odID, kolicina, tipAmb, vrsta, sourceName
    End If

    Dim vrstaK As String, odTipK As String, naTipK As String
    vrstaK = modAmbalazaUgovor.AmbVrstaKanon(vrsta)
    odTipK = modAmbalazaUgovor.AmbNalogTipKanon(odTip)
    naTipK = modAmbalazaUgovor.AmbNalogTipKanon(naTip)
    If Len(vrstaK) = 0 Or Len(odTipK) = 0 Or Len(naTipK) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, _
                  "Kanonski zapis nije nadjen posle prosle kapije -- kvar ugovora."
    End If

    If Len(Trim$(dokTip)) = 0 Or Len(Trim$(dokID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "Red knjige nema identitet dokumenta (AMB-INV-04 i -08)."
    End If

    If Len(Trim$(stornoOd)) = 0 Then
        ' AMB-INV-10 (sprovodjenje AMB-10-ODL-3): dokument ZAKLJUCAVA svoj poslovni
        ' par naloga. AMB-INV-04 to ne pokriva -- njegov kljuc nosi vrstu i tip, pa
        ' dva protivpartnera prolaze cim se razlikuje bilo koje od toga dvoga:
        '
        '   ADK-1  Stanica -> K1  IZDATA_PRAZNA  GAJBA_A
        '   ADK-1  K2 -> Stanica  POVRAT_PRAZNE  GAJBA_A    <- druga vrsta
        '   ADK-1  Stanica -> K2  IZDATA_PRAZNA  GAJBA_B    <- drugi tip
        Dim par As String
        par = DokumentParProblem(dokTip, dokID, odTipK, odID, naTipK, naID, vrstaK, sourceName)
        If Len(par) > 0 Then
            Err.Raise AMB_ERR_JEDAN_PARTNER, sourceName, par
        End If

        Dim sudar As String
        sudar = IdentitetZauzeo(dokTip, dokID, vrstaK, tipAmb, sourceName)
        If Len(sudar) > 0 Then
            Err.Raise AMB_ERR_IDENTITET, sourceName, _
                      "AMB-INV-04: " & Trim$(dokTip) & " '" & Trim$(dokID) & "' je vec knjizio " & _
                      vrstaK & " za '" & Trim$(tipAmb) & "' (red " & sudar & ")."
        End If
    End If

    Dim colCount As Long
    colCount = TabelaBrojKolona(TBL_AMBALAZA)
    If colCount <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, "Ne mogu da odredim broj kolona za tblAmbalaza."
    End If

    ' Transakcioni identitet, ne max+1: StornoOd pokazuje na AmbID, pa on mora biti
    ' stabilan i neprotumaciv (ista odluka kao ReversID, DOCUMENT_HEADER_LINES par. 2).
    Dim novID As String
    novID = NewEntityID("AMB-")
    If Len(Trim$(novID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, "NewEntityID nije vratio AmbID."
    End If

    Dim rowData() As Variant
    ReDim rowData(0 To colCount - 1)

    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_ID, novID, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_DATUM, datum, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_TIP, Trim$(tipAmb), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_KOLICINA, kolicina, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_OD_TIP, odTipK, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_OD_ID, Trim$(odID), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_NA_TIP, naTipK, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_NA_ID, Trim$(naID), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_DOK_TIP, Trim$(dokTip), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_DOK_ID, Trim$(dokID), sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_VRSTA_KRETANJA, vrstaK, sourceName
    SetRowValueByColumn rowData, TBL_AMBALAZA, COL_AMB_STORNO_OD, Trim$(stornoOd), sourceName

    Dim rowIdx As Long
    rowIdx = AppendRow(TBL_AMBALAZA, rowData)
    If rowIdx <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, "AppendRow nije uspeo za tblAmbalaza."
    End If

    UpisiRedKnjige = novID
End Function

' AMB-INV-10: dokument ZAKLJUCAVA svoj poslovni par naloga.
'
' Prva verzija je BROJALA naloge van granice i trazila "najvise dva". To pada tacno
' tamo gde je jedna strana granica (review #400, drugi krug):
'
'   ADK-NAB-1  SpoljniSvet -> Stanica1  NABAVKA  GAJBA_A
'   ADK-NAB-1  SpoljniSvet -> Stanica2  NABAVKA  GAJBA_B
'
' dva realna naloga, broj = 2, PROSLO BI -- a to je jedan nabavni dokument preko DVE
' stanice, tacno ono zbog cega AMB-10-ODL-3 postoji. Ogledalno vazi za OTPIS.
'
' Zato se ne broji nego POREDI: svi ORIGINALNI redovi dokumenta, izuzev generisanog
' pokrica, imaju JEDAN I ISTI neuredjen par {Od, Na}.
'
' POKRICE JE IZUZETO jer ga pisac generise sam (ULAZ_TUDJE nikad nije zahtev), pa
' njegov par ({SpoljniSvet, izvor}) nije poslovni par dokumenta. Ali ne sme da uvede
' trecu stranu, pa njegovo ODREDISTE mora biti clan zakljucanog para.
'
' Vazi za SVE dokumente, ne samo ambalazne: merenje svih devet mesta knjizenja daje
' po dokumentu tacno jedan par (otkup K1<->Stanica, otpremnica Stanica<->Vozac,
' prijemnica Kupac<->Vozac, revers dve strane). Ogranicenje je time na KNJIZI, ne na
' vrsti dokumenta -- isti razlog zbog kog AMB-INV-04 nosi DokumentTIP.
Private Function DokumentParProblem(ByVal dokTip As String, ByVal dokID As String, _
                                    ByVal odTipK As String, ByVal odID As String, _
                                    ByVal naTipK As String, ByVal naID As String, _
                                    ByVal vrstaK As String, _
                                    ByVal sourceName As String) As String
    Dim parovi As Object, clanovi As Object, pokrica As Object
    Set parovi = CreateObject("Scripting.Dictionary")
    Set clanovi = CreateObject("Scripting.Dictionary")
    Set pokrica = CreateObject("Scripting.Dictionary")

    DodajStranu parovi, clanovi, pokrica, odTipK, odID, naTipK, naID, vrstaK

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)

    If Not IsEmpty(data) Then
        Dim kol As Object, vrste As Object
        Set kol = KnjigaZaCitanje(data, sourceName, vrste)

        Dim i As Long
        For i = 1 To UBound(data, 1)
            If RedDoticeKnjigu(data, i, kol) Then
                If Len(AmbText(data(i, kol(COL_AMB_STORNO_OD)))) = 0 Then
                    If StrComp(AmbText(data(i, kol(COL_AMB_DOK_TIP))), Trim$(dokTip), vbTextCompare) = 0 And _
                       StrComp(AmbText(data(i, kol(COL_AMB_DOK_ID))), Trim$(dokID), vbTextCompare) = 0 Then
                        DodajStranu parovi, clanovi, pokrica, _
                                    AmbText(data(i, kol(COL_AMB_OD_TIP))), AmbText(data(i, kol(COL_AMB_OD_ID))), _
                                    AmbText(data(i, kol(COL_AMB_NA_TIP))), AmbText(data(i, kol(COL_AMB_NA_ID))), _
                                    AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA)))
                    End If
                End If
            End If
        Next i
    End If

    If parovi.count > 1 Then
        DokumentParProblem = "AMB-INV-10: " & Trim$(dokTip) & " '" & Trim$(dokID) & _
                             "' bi nosio " & CStr(parovi.count) & " razlicita para naloga (" & _
                             Join(parovi.Keys, " | ") & ") -- jedan dokument, jedan poslovni par."
        Exit Function
    End If

    ' Pokrice bez zakljucanog para nema sa cim da se poredi. Pisac ga ne proizvodi
    ' sam: trazeni prenos se upisuje PRE pokrica, pa par uvek postoji. Ostaje kao
    ' izgovorena granica, ne kao tiha.
    If parovi.count = 0 Then Exit Function

    Dim k As Variant
    For Each k In pokrica.Keys
        If Not clanovi.Exists(CStr(k)) Then
            DokumentParProblem = "AMB-INV-10: pokrice deficita ide na '" & CStr(k) & _
                                 "', a par dokumenta je " & Join(parovi.Keys, "") & "."
            Exit Function
        End If
    Next k
End Function

' Jedan red doprinosi ILI paru dokumenta ILI spisku pokrica -- nikad oboma.
Private Sub DodajStranu(ByRef parovi As Object, ByRef clanovi As Object, _
                        ByRef pokrica As Object, _
                        ByVal odTip As String, ByVal odID As String, _
                        ByVal naTip As String, ByVal naID As String, _
                        ByVal vrsta As String)
    If StrComp(Trim$(vrsta), AMB_VK_ULAZ_TUDJE, vbTextCompare) = 0 Then
        Dim p As String
        p = NalogKljuc(naTip, naID)
        If Not pokrica.Exists(p) Then pokrica.Add p, True
        Exit Sub
    End If

    Dim par As String
    par = ParKljuc(odTip, odID, naTip, naID)
    If Not parovi.Exists(par) Then parovi.Add par, True

    Dim a As String, b As String
    a = NalogKljuc(odTip, odID)
    b = NalogKljuc(naTip, naID)
    If Not clanovi.Exists(a) Then clanovi.Add a, True
    If Not clanovi.Exists(b) Then clanovi.Add b, True
End Sub

' Vraca AmbID reda koji je vec zauzeo isti identitet efekta, inace "".
Private Function IdentitetZauzeo(ByVal dokTip As String, ByVal dokID As String, _
                                 ByVal vrstaK As String, ByVal tipAmb As String, _
                                 ByVal sourceName As String) As String
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cOdTip As Long, cNaTip As Long, cVK As Long
    Dim cTipA As Long, cDokT As Long, cDokI As Long, cSt As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, sourceName, vrste)

    cID = kol(COL_AMB_ID)
    cOdTip = kol(COL_AMB_OD_TIP)
    cNaTip = kol(COL_AMB_NA_TIP)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cDokT = kol(COL_AMB_DOK_TIP)
    cDokI = kol(COL_AMB_DOK_ID)
    cSt = kol(COL_AMB_STORNO_OD)

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If Len(AmbText(data(i, cSt))) = 0 Then
                If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cVK)), Trim$(vrstaK), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 Then
                    IdentitetZauzeo = AmbText(data(i, cID))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ZBIR VEC UPISANOG ZA ISTI ZAHTEV -- osnova idempotencije.
'
' Meri se po PARU NALOGA i po SKUPU vrsta koje taj zahtev sme da proizvede
' (AmbVrsteZahteva), a ne po trazenoj vrsti i kolicini jednog reda: podela
' obaveze razbija zahtev od 20 na 12 + 8, pa bi poredjenje sa jednim redom
' prijavilo sudar nad ispravnim ponavljanjem.
Private Function ZbirZahteva(ByVal dokTip As String, ByVal dokID As String, _
                             ByVal datum As Date, ByVal tipAmb As String, _
                             ByVal odTip As String, ByVal odID As String, _
                             ByVal naTip As String, ByVal naID As String, _
                             ByVal vrstaK As String, _
                             ByRef outAmbID As String, _
                             ByVal sourceName As String) As Double
    outAmbID = ""

    Dim dozvoljene As Variant
    dozvoljene = modAmbalazaUgovor.AmbVrsteZahteva(vrstaK)
    If UBound(dozvoljene) < LBound(dozvoljene) Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cID As Long, cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cTipA As Long, cKol As Long, cDokT As Long, cDokI As Long
    Dim cSt As Long, cDat As Long
    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, sourceName, vrste)

    cID = kol(COL_AMB_ID)
    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)
    cDokT = kol(COL_AMB_DOK_TIP)
    cDokI = kol(COL_AMB_DOK_ID)
    cSt = kol(COL_AMB_STORNO_OD)
    cDat = kol(COL_AMB_DATUM)

    Dim i As Long, j As Long, vr As String, zbir As Double
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If Len(AmbText(data(i, cSt))) = 0 Then
                If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cTipA)), Trim$(tipAmb), vbTextCompare) = 0 And _
                   IstiDan(data(i, cDat), datum) Then
                    If IstiNalog(AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), odTip, odID) And _
                       IstiNalog(AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), naTip, naID) Then

                        vr = AmbText(data(i, cVK))
                        For j = LBound(dozvoljene) To UBound(dozvoljene)
                            If StrComp(vr, CStr(dozvoljene(j)), vbTextCompare) = 0 Then
                                zbir = zbir + CDbl(data(i, cKol))
                                If Len(outAmbID) = 0 Or _
                                   StrComp(vr, Trim$(vrstaK), vbTextCompare) = 0 Then
                                    outAmbID = AmbText(data(i, cID))
                                End If
                                Exit For
                            End If
                        Next j
                    End If
                End If
            End If
        End If
    Next i

    ZbirZahteva = zbir
End Function

' ============================================================
' STORNO U KNJIZI -- KONTRA-STAV, NE ZASTAVICA (AMB-10-ODL-16)
' ============================================================
'
' Zastavica Stornirano je mehanizam STAROG modela i nov citalac je ne gleda:
' RedDoticeKnjigu trazi Od_*/Na_*/VrstaKretanja/StornoOd, a AmbSaldoNaloga
' sumira po tome. Dokument presecen na nov model a storniran zastavicom
' ostavio bi gajbe na saldu TIHO, pa ovaj ulaz ide PRED cutover mesta
' knjizenja, ne posle njega.
'
' DATUM KONTRA-STAVA JE DATUM ORIGINALA, ne danasnji. Stara zastavica je red
' uklanjala iz SVIH perioda; isti datum je jedini oblik koji ne menja nijedan
' periodski saldo. Danasnji datum bi ostavio fantom u starom periodu i visak
' u novom.
'
' IDEMPOTENTNO: original koji VEC ima kontra-stav se preskace, pa drugi poziv
' vraca 0 i ne duplira. Vraca broj upisanih kontra-stavova.
Public Function StornirajAmbalazuDokumenta(ByVal tx As clsTransaction, _
                                           ByVal dokTip As String, _
                                           ByVal dokID As String) As Long
    Const SRC As String = "modAmbalaza.StornirajAmbalazuDokumenta"

    If tx Is Nothing Then
        Err.Raise AMB_ERR_STORNO, SRC, _
                  "Storno knjige trazi aktivnu transakciju (AMB-INV-08)."
    End If
    If Len(Trim$(dokTip)) = 0 Or Len(Trim$(dokID)) = 0 Then
        Err.Raise AMB_ERR_STORNO, SRC, "Storno knjige trazi identitet dokumenta."
    End If

    RequireKnjigaSchema SRC

    ' OVAJ PRIMITIV NE VEZUJE DOKUMENT, I TO JE SUSTINA KAPIJE.
    '
    ' Ranije je zvao BindSourceDocument sam, uz obrazlozenje da to nije
    ' samopotvrda jer ista kapija trazi i izvornu tabelu u snapshotu. Ta odbrana
    ' je falsifikovana (review 03.10.2026, P1 #1): snapshot je jeftin i ne
    ' dokazuje da je dokument PROMENJEN. Pozivalac je mogao da snapshotuje
    ' tblOtkup, anulira efekat AKTIVNOG otkupa i prodje sve kapije.
    '
    ' Ledger-storno NIJE pisac izvornog dokumenta. Vezivanje zato pripada
    ' kanonskom piscu koji je zaglavlje stvarno promenio (modStorno.StornoOtkup:
    ' MarkRowStornirano, pa bind, pa ovaj poziv). Ako nije vezao, ovde pada
    ' fail-closed kroz RequireAmbTxIzvorniDokument -- sto je i jedini nacin da
    ' AMB-INV-08 znaci "izvorni dokument je promenjen u istoj TX", a ne
    ' "transakcija je rekla da ga poseduje".
    '
    ' Posledica: tblAmbalazaDokument (nabavka, revers) jos NEMA kanonskog
    ' storno pisca, pa njegov ledger storno ovde pada -- namerno, dok taj pisac
    ' ne nastane (10d).

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    Dim cID As Long, cDat As Long, cTipA As Long, cKol As Long
    Dim cOdTip As Long, cOdID As Long, cNaTip As Long, cNaID As Long
    Dim cVK As Long, cDokT As Long, cDokI As Long, cSt As Long
    cID = kol(COL_AMB_ID)
    cDat = kol(COL_AMB_DATUM)
    cTipA = kol(COL_AMB_TIP)
    cKol = kol(COL_AMB_KOLICINA)
    cOdTip = kol(COL_AMB_OD_TIP)
    cOdID = kol(COL_AMB_OD_ID)
    cNaTip = kol(COL_AMB_NA_TIP)
    cNaID = kol(COL_AMB_NA_ID)
    cVK = kol(COL_AMB_VRSTA_KRETANJA)
    cDokT = kol(COL_AMB_DOK_TIP)
    cDokI = kol(COL_AMB_DOK_ID)
    cSt = kol(COL_AMB_STORNO_OD)

    ' PRVI PROLAZ: na sta kontra-stavovi vec pokazuju. Bez ovoga drugi poziv
    ' udvaja storno, a saldo prelazi na drugu stranu umesto da stane na nuli.
    '
    ' NEDOSTIZNA ODBRANA -- namerno, i tako imenovana. Primitiv ne vezuje dokument
    ' (P1 #1), pa ga test ne moze pozvati direktno; jedini pozivalac je
    ' modStorno.StornoOtkup, a RequireStornoAllowed odbija DRUGI storno pre njega.
    ' Grana se zato NE MOZE dosegnuti i njena sabotaza je obrisana iz kataloga:
    ' sabotaza koja ne moze da obori nijednu tvrdnju je placebo. Ostaje jer bi
    ' drugi pozivalac (otpremnica, prijemnica) bez nje udvajao kontra-stav.
    Dim vecStornirani As Object
    Set vecStornirani = CreateObject("Scripting.Dictionary")
    Dim i As Long, st As String
    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            st = AmbText(data(i, cSt))
            If Len(st) > 0 Then
                If Not vecStornirani.Exists(st) Then vecStornirani.Add st, True
            End If
        End If
    Next i

    ' DRUGI PROLAZ: aktivni originali ovog dokumenta.
    Dim originali As Collection
    Set originali = New Collection
    ' SVI nalozi koje kontra-stavovi dotice -- iz njih se posle mere DVE
    ' invarijante: AMB-INV-07 nad realnim, AMB-INV-09 nad partnerskim.
    Dim nalozi As Object
    Set nalozi = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(data, 1)
        If RedDoticeKnjigu(data, i, kol) Then
            If Len(AmbText(data(i, cSt))) = 0 Then
                If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 And _
                   StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 Then
                    If Not vecStornirani.Exists(AmbText(data(i, cID))) Then
                        If Not IsDate(data(i, cDat)) Then
                            Err.Raise AMB_ERR_STORNO, SRC, _
                                      "Red knjige " & AmbText(data(i, cID)) & _
                                      " nema datum -- kontra-stav ga nasledjuje, pa bez " & _
                                      "njega storno ne moze da nastane."
                        End If
                        originali.Add Array(AmbText(data(i, cID)), _
                                            CDate(data(i, cDat)), _
                                            AmbText(data(i, cTipA)), _
                                            CDbl(data(i, cKol)), _
                                            AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), _
                                            AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), _
                                            AmbText(data(i, cVK)))
                        ZabeleziNalog nalozi, _
                                      AmbText(data(i, cOdTip)), AmbText(data(i, cOdID)), _
                                      AmbText(data(i, cTipA))
                        ZabeleziNalog nalozi, _
                                      AmbText(data(i, cNaTip)), AmbText(data(i, cNaID)), _
                                      AmbText(data(i, cTipA))
                    End If
                End If
            End If
        End If
    Next i

    If originali.count = 0 Then
        ' FAIL-CLOSED ZA AMBALAZNI DOKUMENT (review 08.10.2026, P3).
        '
        ' On po konstrukciji ima bar jedan poslovni red: zaglavlje i knjiga
        ' nastaju u ISTOJ transakciji. Nula aktivnih originala znaci pokvaren
        ' zapis, a tih izlaz bi ga ostavio sa zaglavljem "Stornirano = Da" i
        ' NULA kontra-stavova -- dokument koji izgleda storniran, a knjiga ga
        ' nikad nije videla.
        '
        ' ROBNI dokument sme da nema nijedan red (prijemnica bez ambalaze je
        ' legitimna), pa za njega ostaje tih izlaz -- inace bi svaki storno
        ' takve prijemnice pao.
        If StrComp(Trim$(dokTip), DOK_TIP_AMBALAZA_DOKUMENT, vbTextCompare) = 0 Then
            Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                      "Ambalazni dokument '" & Trim$(dokID) & "' nema nijedan aktivan " & _
                      "red knjige -- storno bi ostavio zaglavlje bez kontra-stava."
        End If
        Exit Function
    End If

    Dim r As Variant
    For Each r In originali
        ' Od i Na ZAMENJENI, StornoOd pokazuje na original.
        UpisiRedKnjige tx, CDate(r(1)), CStr(r(2)), CDbl(r(3)), _
                       CStr(r(6)), CStr(r(7)), CStr(r(4)), CStr(r(5)), _
                       Trim$(dokTip), Trim$(dokID), CStr(r(8)), CStr(r(0)), SRC
    Next r

    ' DVE INVARIJANTE NAD POSLE-STANJEM.
    '
    ' Kontra-stav ide direktno kroz UpisiRedKnjige, pa ZAOBILAZI sve sto stoji u
    ' PrenesiAmbalazu -- a tamo zivi i AMB-INV-07. Bez ovih provera storno pise
    ' stanje koje normalan pisac eksplicitno zabranjuje (review 03.10.2026, P1 #2):
    '
    '   otkup donese 20 na stanicu -> ta 20 odu dalje (saldo 0) -> storno otkupa
    '   upise kontra-stav -20  =>  saldo -20, a INV-09 je uredan
    '
    ' Mere se POSTOJECIM citaocima nad upisanim stanjem, ne drugom kopijom pravila
    ' o znaku -- druga kopija bi se razisla sa prvom.
    Dim k As Variant, nl As Variant
    Dim saldo As Double, ob As Double

    ' AMB-INV-07: realan nalog ne sme da ostane u minusu. GRANICA (SpoljniSvet)
    ' je izuzeta po konstrukciji -- ona nema fizicko stanje (6.3), pa nije u
    ' klasi REALAN i petlja je i ne vidi.
    For Each k In nalozi.Keys
        nl = nalozi(k)
        If modAmbalazaUgovor.AmbNalogUKlasi(AMB_KLASA_REALAN, CStr(nl(0))) Then
            saldo = AmbSaldoNaloga(CStr(nl(0)), CStr(nl(1)), CStr(nl(2)))
            If saldo < 0 Then
                Err.Raise AMB_ERR_STORNO, SRC, _
                          "AMB-INV-07: storno bi ostavio saldo " & CStr(saldo) & " na " & _
                          CStr(nl(0)) & " '" & CStr(nl(1)) & "' za '" & CStr(nl(2)) & _
                          "' -- ta ambalaza je posle ovog dokumenta otisla dalje, pa " & _
                          "se dokument ne moze stornirati bez njenog vracanja."
            End If
        End If
    Next k

    ' AMB-INV-09: zahtev koji je ovaj ulaz NASLEDIO od pisca (PrenesiAmbalazu,
    ' VRACANJE_TUDJE): storno ULAZA tudje ambalaze cija je obaveza vec zatvorena
    ' vracanjem daje NEGATIVNU obavezu. Klasa duga se CITA iz matrice pokrica, jer
    ' dug nastaje tacno tim dogadjajem -- AmbObavezaPartneru dize gresku za nalog
    ' van te klase, pa se van nje i ne sme pitati.
    Dim klasaDuga As String
    klasaDuga = modAmbalazaUgovor.AmbPokriceKlasa()
    If Len(klasaDuga) > 0 Then
        For Each k In nalozi.Keys
            nl = nalozi(k)
            If modAmbalazaUgovor.AmbNalogUKlasi(klasaDuga, CStr(nl(0))) Then
                ob = AmbObavezaPartneru(CStr(nl(0)), CStr(nl(1)), CStr(nl(2)))
                If ob < 0 Then
                    Err.Raise AMB_ERR_STORNO, SRC, _
                              "Storno bi ostavio obavezu " & CStr(ob) & " prema " & _
                              CStr(nl(0)) & " '" & CStr(nl(1)) & "' za '" & CStr(nl(2)) & _
                              "' -- tudja ambalaza je vec vracena, pa se njen ulaz " & _
                              "ne moze stornirati (AMB-INV-09)."
                End If
            End If
        Next k
    End If

    StornirajAmbalazuDokumenta = originali.count
End Function

' Nalog koji je kontra-stav dotakao. BEZ klasne kapije: jedan popis, a klasu
' bira CITALAC -- INV-07 gleda REALAN, INV-09 klasu pokrica. Dva popisa bi se
' razisla, a prazan popis bi tiho ugasio onu proveru koja ga nema.
Private Sub ZabeleziNalog(ByRef nalozi As Object, _
                          ByVal tip As String, ByVal id As String, _
                          ByVal tipAmb As String)
    If Len(Trim$(tip)) = 0 Then Exit Sub

    Dim kljuc As String
    kljuc = NalogKljuc(tip, id) & "|" & UCase$(Trim$(tipAmb))
    If Not nalozi.Exists(kljuc) Then
        nalozi.Add kljuc, Array(Trim$(tip), Trim$(id), Trim$(tipAmb))
    End If
End Sub

' IMA LI KNJIGA KONTRA-STAV ZA OVAJ DOKUMENT?
'
' Javno jer ga undo garda (modStornoZurnal.UndoGuardReasonZaOp) mora pitati, a
' knjigu ne sme da cita sama -- oblik reda knjige nije njen posao.
'
' KLJUC JE KOMPOZITAN: (DokumentTIP, DokumentID). Prva verzija je trazila samo
' ID, uz obrazlozenje da je ID globalno jedinstven pa tip ne dodaje razlucivost
' -- a AMB-INV-04 nosi DokumentTIP tacno zato sto se jedan globalni namespace
' DokumentID-eva NE SME pretpostaviti (review 03.10.2026, P2 #1). Garda koja
' sluzi svim presecenim dokumentima ne sme da ima slabiji identitet od knjige.
Public Function AmbImaKontraStav(ByVal dokTip As String, _
                                 ByVal dokID As String) As Boolean
    Const SRC As String = "modAmbalaza.AmbImaKontraStav"

    If Len(Trim$(dokTip)) = 0 Or Len(Trim$(dokID)) = 0 Then Exit Function

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim cDokT As Long, cDokI As Long, cSt As Long
    cDokT = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cDokI = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cSt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_STORNO_OD)
    ' Zatecena sveska bez nove kolone nema ni kontra-stavova -- ali odgovor
    ' "nema" bi tada bio pretpostavka. Sema je kanon, pa nedostatak kolone je
    ' kvar, i tako se i dize.
    If cDokT = 0 Then Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                                "tblAmbalaza nema " & COL_AMB_DOK_TIP & "."
    If cDokI = 0 Then Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                                "tblAmbalaza nema " & COL_AMB_DOK_ID & "."
    If cSt = 0 Then Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                              "tblAmbalaza nema " & COL_AMB_STORNO_OD & "."

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Len(AmbText(data(i, cSt))) > 0 Then
            If StrComp(AmbText(data(i, cDokT)), Trim$(dokTip), vbTextCompare) = 0 Then
                If StrComp(AmbText(data(i, cDokI)), Trim$(dokID), vbTextCompare) = 0 Then
                    AmbImaKontraStav = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' STORNO AMBALAZNOG DOKUMENTA -- revers, nabavka, otpis.
'
' Do ovog reza ga NIJE BILO: tblAmbalazaDokument je imao pisca i kolonu
' Stornirano, ali nijedan put da je okrene. Dok je revers bio skup nogu u
' tblAmbalaza, storno je isao po REDU (ekran Storno, AmbID noge). Cim revers
' postane dokument, storno mora da bude PO DOKUMENTU -- inace bi zaglavlje i
' knjiga mogli da se razidju.
'
' Redosled je isti kao kod otkupa, otpremnice i prijemnice: oznaci zaglavlje,
' pa VEZI (AMB-10-ODL-15 trazi bind POSLE stvarne izmene), pa kontra-stav.
' Kontra-stav sam proverava AMB-INV-07 i -09 nad POSLE-stanjem.
Public Function StornirajAmbDokument_TX(ByVal ambDokID As String) As Boolean
    Const SRC As String = "modAmbalaza.StornirajAmbDokument_TX"

    Dim tx As clsTransaction
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    ' Nepostojeci ili vec storniran dokument pada PRE transakcije -- jezgro ga
    ' proverava ponovo, ali ovde se ne otvara tx za dokument koji ne postoji.
    RedAmbDokumenta ambDokID, SRC

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    StornirajAmbDokument tx, ambDokID

    tx.CommitTx
    Set tx = Nothing
    StornirajAmbDokument_TX = True
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogError SRC, errDesc, errNum
    StornirajAmbDokument_TX = False
End Function

' JEZGRO STORNA AMBALAZNOG DOKUMENTA, U TRANSAKCIJI POZIVAOCA.
'
' Izdvojeno za ispravku reversa (10d): storno starog i upis novog moraju da dele
' JEDAN rollback, pa jezgro ne sme samo da otvara i zatvara transakciju. Zato
' trazi snapshot zaglavlja kao i svaki drugi pisac -- ne pretpostavlja ga.
' Knjigu trazi StornirajAmbalazuDokumenta sam.
Private Sub StornirajAmbDokument(ByVal tx As clsTransaction, ByVal ambDokID As String)
    Const SRC As String = "modAmbalaza.StornirajAmbDokument"
    Dim red As Long

    RequireAmbTxVlasnistvo tx, TBL_AMBALAZA_DOKUMENT, SRC
    red = RedAmbDokumenta(ambDokID, SRC)

    ' MarkRowStornirano je PRIVATE u modStorno, pa se iz ovog modula ne vidi --
    ' a ono je ionako samo omotac oko RequireUpdateCell. Zove se primitiv
    ' direktno, umesto da se tudja privatnost otvara zbog jednog poziva.
    '
    ' Kvar se nije video ni u jednoj jeftinoj kapiji: VBA kompajlira NA ZAHTEV,
    ' pa je 'Sub or Function not defined' pukao tek kad je prvi test pozvao ovu
    ' proceduru -- posle 600s i ubijenog Excela.
    ' "Da" je literal jer je STORNO_DA takodje Private u modStorno. Vrednost je
    ' ista koju citaju IsStorniranoValue i svi ostali pisci.
    RequireUpdateCell TBL_AMBALAZA_DOKUMENT, red, COL_STORNIRANO, "Da", SRC
    tx.BindSourceDocument DOK_TIP_AMBALAZA_DOKUMENT, ambDokID
    StornirajAmbalazuDokumenta tx, DOK_TIP_AMBALAZA_DOKUMENT, ambDokID
End Sub

' Red zaglavlja, fail-closed. Nepostojeci ili vec storniran dokument se NE
' stornira drugi put -- druga storno operacija bi upisala drugi kontra-stav nad
' istim redovima, a AmbImaKontraStav bi ga tek posle prijavio.
Private Function RedAmbDokumenta(ByVal ambDokID As String, _
                                 ByVal SRC As String) As Long
    Dim redovi As Collection
    RequireTacnoJedan TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID), _
                      "AmbDokID", SRC
    Set redovi = FindRows(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, Trim$(ambDokID))
    RedAmbDokumenta = CLng(redovi(1))

    If UCase$(Trim$(NzToText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                         Trim$(ambDokID), COL_STORNIRANO)))) = "DA" Then
        Err.Raise AMB_ERR_IDENTITET, SRC, _
                  "Ambalazni dokument '" & Trim$(ambDokID) & "' je vec storniran."
    End If
End Function

' REVERS KAO AMBALAZNI DOKUMENT (10b-2, cetvrto mesto knjizenja).
'
' Stari pisac (modDokumenta.SaveOMUlaz_TX) je knjizio SEST nogu u cetiri smera,
' a vozaca nosio kao ZIG -- pa se njegov saldo dobijao INVERZIJOM smera, sto je
' fail-open: citalac koji inverziju zaboravi dobija pogresan ZNAK, ne gresku.
' Ovde je po smeru JEDAN red koji imenuje obe strane, iz zatvorene mape
' AmbReversSmerovi.
'
' Revers je AMBALAZNI dokument (AMB-10-ODL-5), pa dobija pravo zaglavlje u
' tblAmbalazaDokument umesto brojDok rasutog po nogama. Broj je NAS
' (AmbDokBrojOwnerKlasa = SOPSTVENI, vlasnik stanica), a kapija zauzetosti iz
' AMB-10-ODL-20 radi isti posao kao zateceni RequireBrojSlobodanUNizu --
' nad istim kljucem (vlasnik, dan).
Public Function UpisiReversAmbalaze_TX(ByVal datum As Date, ByVal broj As String, _
                                       ByVal stanicaID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal kolicina As Double, _
                                       ByVal smer As String, _
                                       ByVal kooperantID As String, _
                                       ByVal vozacID As String, _
                                       Optional ByVal napomena As String = "") As String
    Const SRC As String = "modAmbalaza.UpisiReversAmbalaze_TX"

    Dim tx As clsTransaction
    Dim dokID As String, brojK As String, red As Variant
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    red = ReversSmerZaUpis(kolicina, smer, SRC)
    brojK = ReversBrojZaUpis(datum, broj, stanicaID, SRC)

    Set tx = New clsTransaction
    tx.BeginTx
    ' Obe tabele: zaglavlje i knjiga nastaju u ISTOM rollback-u (AMB-INV-08).
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    dokID = UpisiReversUTx(tx, red, datum, brojK, stanicaID, tipAmb, kolicina, _
                           smer, kooperantID, vozacID, napomena, SRC)

    tx.CommitTx
    Set tx = Nothing

    UpisiReversAmbalaze_TX = dokID
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogError SRC, errDesc, errNum
    Err.Raise errNum, SRC, errDesc
End Function

' Ulaz reversa koji se sudi PRE broja: kolicina, pa smer iz zatvorene mape.
' Redosled gresaka je isti kao pre izdvajanja jezgra (10d) -- operater koji
' pogresi i smer i broj dobija poruku o smeru.
Private Function ReversSmerZaUpis(ByVal kolicina As Double, ByVal smer As String, _
                                  ByVal SRC As String) As Variant
    If kolicina <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Revers trazi pozitivnu kolicinu."
    End If
    ReversSmerZaUpis = ReversSmerRed(smer, SRC)
End Function

' Broj reversa koji ide u upis: prosledjen prolazi kapije starog niza, prazan se
' izracunava iz ambalaznog niza stanice. Izdvojeno iz UpisiReversAmbalaze_TX da
' ga ispravka (10d) prolazi ISTIM putem -- druga kopija kapija bi se razisla.
Private Function ReversBrojZaUpis(ByVal datum As Date, ByVal broj As String, _
                                  ByVal stanicaID As String, _
                                  ByVal SRC As String) As String
    Dim brojK As String
    brojK = Trim$(broj)
    If Len(brojK) > 0 Then
        ' DVE KAPIJE BROJA, A NOV MODEL POKRIVA SAMO JEDNU.
        '
        ' Zauzetost (AMB-10-ODL-20) radi UpisiAmbDokument -- to je zamena za
        ' zateceni RequireBrojSlobodanUNizu. Ali OBLIK I KONTEKST broja (pripada li
        ' bas nizu te stanice i tog dana) nov model NE proverava, pa bi prelazak
        ' tiho izgubio tu kapiju.
        '
        ' Zato ostaje, i to samo za broj KOJI JE POZIVALAC ZADAO: generisan broj
        ' dolazi iz ambalaznog niza i nema REV oblik, pa bi ga ova kapija odbila.
        modBrojevi.RequireBrojUKontekstu modBrojevi.KIND_REV, stanicaID, datum, _
                                         brojK, SRC
        ' I ZAUZETOST NAD STARIM NIZOM -- ovo sam prvo ispustio.
        '
        ' Mislio sam da je AMB-10-ODL-20 (kapija u UpisiAmbDokument) zamena za
        ' RequireBrojSlobodanUNizu. Nije: ODL-20 sudi nad tblAmbalazaDokument, a
        ' stari niz zivi nad tblAmbalaza. To su DVA RAZLICITA NIZA, pa bi broj
        ' zauzet starim reversom bio slobodan za nov -- i obrnuto.
        '
        ' Dok oba oblika postoje (do 10e), oba niza moraju da vaze. Kad stari
        ' redovi nestanu, ova provera postaje mrtva i brise se sa njima.
        modBrojevi.RequireBrojSlobodanUNizu modBrojevi.KIND_REV, stanicaID, datum, _
                                            brojK, SRC
    Else
        brojK = modBrojevi.GenerateBrojAmbDokumenta(AMB_NALOG_STANICA, stanicaID, datum)
        If Len(brojK) = 0 Then
            Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                      "Broj reversa nije izracunat -- vidi Log."
        End If
    End If
    ReversBrojZaUpis = brojK
End Function

' JEZGRO UPISA NASEG REVERSA, U TRANSAKCIJI POZIVAOCA: zaglavlje pa JEDAN red
' knjige. `red` je vec proveren red mape smerova (ReversSmerZaUpis), a broj je
' konacan (ReversBrojZaUpis) -- jezgro ne odlucuje ni o jednom, samo pise.
Private Function UpisiReversUTx(ByVal tx As clsTransaction, ByVal red As Variant, _
                                ByVal datum As Date, ByVal brojK As String, _
                                ByVal stanicaID As String, ByVal tipAmb As String, _
                                ByVal kolicina As Double, ByVal smer As String, _
                                ByVal kooperantID As String, ByVal vozacID As String, _
                                ByVal napomena As String, ByVal SRC As String) As String
    Dim dokID As String, odTip As String, naTip As String, vrstaK As String
    odTip = CStr(red(1))
    naTip = CStr(red(2))
    vrstaK = CStr(red(3))

    dokID = UpisiAmbDokument(tx, AMB_DOK_REVERS, brojK, datum, _
                             AMB_NALOG_STANICA, stanicaID, napomena)

    PrenesiAmbalazu tx, datum, tipAmb, kolicina, _
                    odTip, ReversNalogID(odTip, stanicaID, kooperantID, vozacID, smer, SRC), _
                    naTip, ReversNalogID(naTip, stanicaID, kooperantID, vozacID, smer, SRC), _
                    vrstaK, DOK_TIP_AMBALAZA_DOKUMENT, dokID

    UpisiReversUTx = dokID
End Function

' KAPIJA ISPRAVKE NASEG REVERSA -- jedan izvor za ekran i za pisca. "" = moze.
'
' Ekran je pita PRE nego sto operater pocne da kuca zamenu (i pre nego sto
' uopste ponudi dugme), a IspravkaReversa_TX je dize kao gresku za pozivaoce
' koji nisu ekran. Dva spiska istih pravila bi se razisla prvom izmenom -- isti
' obrazac kao modOtkup.IspravkaPrepreka.
'
' Ispravka postoji SAMO za vrstu REVERS. Kupcev revers (REVERS_PARTNERA) nosi
' kupcev broj i svoj niz, pa njegova ispravka nije ista radnja; nabavka i otpis
' se ne ispravljaju ovim putem. Za sve njih ostaje storno.
Public Function IspravkaReversaRazlog(ByVal ambDokID As String) As String
    Dim redovi As Collection, vrsta As String, nPoslovnih As Long
    On Error GoTo EH

    ambDokID = Trim$(ambDokID)
    If Len(ambDokID) = 0 Then
        IspravkaReversaRazlog = "Revers za ispravku nije izabran (prazan AmbDokID)."
        Exit Function
    End If

    Set redovi = FindRows(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, ambDokID)
    If redovi.count <> 1 Then
        IspravkaReversaRazlog = "Ambalazni dokument '" & ambDokID & "' nije jednoznacan (" & _
                                CStr(redovi.count) & " redova) -- ispravka se ne nudi."
        Exit Function
    End If

    If UCase$(Trim$(NzToText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, ambDokID, _
                                         COL_STORNIRANO)))) = "DA" Then
        IspravkaReversaRazlog = "Revers '" & ambDokID & "' je vec storniran -- nema sta da se ispravi."
        Exit Function
    End If

    vrsta = Trim$(NzToText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, ambDokID, _
                                       COL_AMBD_VRSTA)))
    If StrComp(vrsta, AMB_DOK_REVERS, vbTextCompare) <> 0 Then
        IspravkaReversaRazlog = "Ispravka postoji samo za nas revers; dokument '" & ambDokID & _
                                "' je " & vrsta & ". Takav dokument se stornira i unosi ponovo."
        Exit Function
    End If

    ' TACNO JEDAN POSLOVNI RED (review #414, P2). Zamena se pise iz forme kao
    ' JEDAN red, a storno povlaci SVE originalne redove dokumenta. AMB-INV-10
    ' dozvoljava dokument sa vise redova (dva tipa ambalaze, oba smera prema
    ' istom partneru), a forma prikazuje samo prvi -- pa bi ispravka tiho
    ' zamenila dva kretanja jednim, i to ono koje operater nije video. Nula je
    ' zaglavlje bez knjige, dakle kvar, i takodje pada. Merenje ide POSLEDNJE:
    ' jedino ono cita celu knjigu.
    nPoslovnih = AmbDokBrojPoslovnihRedova(ambDokID)
    If nPoslovnih <> 1 Then
        IspravkaReversaRazlog = "Revers '" & ambDokID & "' nema tacno jedan poslovni red u knjizi (" & _
                                CStr(nPoslovnih) & ") -- zamena iz forme ne bi bila ista radnja. " & _
                                "Storniraj ga i unesi ponovo."
        Exit Function
    End If
    Exit Function
EH:
    ' FAIL-CLOSED: kapija koja ne moze da procita zaglavlje ne pusta ispravku.
    IspravkaReversaRazlog = "Revers se ne moze proveriti za ispravku: " & Err.description
End Function

' A9 za ispravku: zamena ne sme da nosi STARI poslovni broj. Poruka imenuje
' pravilo i stari broj -- ne zauzetost, koja je druga kapija i drugi razlog.
Private Sub RequireNovBrojIspravke(ByVal broj As String, ByVal stariBroj As String, _
                                   ByVal SRC As String)
    If StrComp(Trim$(broj), stariBroj, vbTextCompare) = 0 Then
        Err.Raise AMB_ERR_IDENTITET, SRC, _
                  "Ispravka reversa mora dobiti NOV broj (A9). Stari: " & stariBroj & "."
    End If
End Sub

' Podmetni pad ispravke reversa tik pre commit-a. ISKLJUCIVO iz test modula.
Public Sub IspravkaReversaPadniTest(ByVal onOff As Boolean)
    mIspravkaPadPreCommit = onOff
End Sub

' ISPRAVKA NASEG REVERSA (10d) -- JEDAN POTEZ, kao IspravkaOtkupa_TX (B-040).
'
' Ispravka reversa NE POMERA ROBU: gajbe su gde su bile, a stari dokument se
' povlaci i izdaje nov, sa NOVIM brojem i NOVIM AmbDokID-em (A9, prosireno na
' REV 14.09.2026). Kroz knjigu zato prolaze CETIRI reda -- kontra-stav starog i
' upis novog, svaki sa obe strane -- pa je saldo posle ispravke ISTI kad je
' promenjen samo broj ili datum, a razlikuje se TACNO za razliku kad je
' promenjena kolicina.
'
' JEDNA TRANSAKCIJA, NE DVOKORAK (odluka 11.10.2026). Stari tok
' (RunReversCorrection -> F7 -> CompleteReversIspravka) je stornirao odmah, a
' zamenu vezivao kasnije, pa je izmedju postojao prozor u kome starog vise nema a
' novog jos nema -- okvir ga je krpio PENDING kontekstom i MANUAL zadatkom. Ovde
' odustanak ne ostavlja trag, a pad bilo kog koraka vraca sve.
'
' STORNO PRVI, kao kod otkupa. AMB-INV-07 se meri posle kontra-stava (u
' StornirajAmbalazuDokumenta), dakle i u medjukoraku: revers cije su gajbe vec
' otisle dalje ne moze da se ispravi -- isto pravilo koje otkup ima danas.
'
' VEZA JE PO ID-u, na oba kraja (IspravkaOdID na novom, ZamenjenSaID na starom)
' i u tragu tblStornoVeze -- nikad po broju. Tip traga je AmbalazaDokument, ne
' FLOW_DOC_REVERS: stari tip je nosio ReversID u OldDocID, pa bi ga citaoci starog
' oblika (IspravkaReversOpis, obrisan u 10e-2) procitali kao tudji identitet.
'
' Vraca AmbDokID novog reversa. Greska se DIZE, kao kod UpisiReversAmbalaze_TX;
' pozivalac je pretvara u poruku operateru.
Public Function IspravkaReversa_TX(ByVal stariAmbDokID As String, ByVal datum As Date, _
                                   ByVal broj As String, ByVal stanicaID As String, _
                                   ByVal tipAmb As String, ByVal kolicina As Double, _
                                   ByVal smer As String, ByVal kooperantID As String, _
                                   ByVal vozacID As String, _
                                   Optional ByVal napomena As String = "") As String
    Const SRC As String = "modAmbalaza.IspravkaReversa_TX"

    Dim tx As clsTransaction
    Dim red As Variant, prepreka As String, cid As String
    Dim stariBroj As String, brojK As String, noviID As String
    Dim rStari As Long, rNovi As Long
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    stariAmbDokID = Trim$(stariAmbDokID)
    prepreka = IspravkaReversaRazlog(stariAmbDokID)
    If Len(prepreka) > 0 Then Err.Raise AMB_ERR_IDENTITET, SRC, prepreka

    red = ReversSmerZaUpis(kolicina, smer, SRC)

    ' A9: NOV POSLOVNI BROJ -- dvaput, jer broj nastaje na dva nacina.
    '
    ' ZADAT broj se poredi PRE kapija broja (ReversBrojZaUpis). Na istoj stanici
    ' i danu stari broj je zauzet (storno ga ne oslobadja), pa bi ga kapija
    ' zauzetosti odbila prva -- porukom o ZAUZETOSTI, ne o ispravci. Dokaz je to
    ' izmerio 11.10.2026: dok je ovo poredjenje stajalo posle kapija, nijedan put
    ' do njega nije vodio za taj slucaj.
    '
    ' IZRACUNAT broj se poredi POSLE: generator broji niz NOVE stanice, pa zamena
    ' na drugoj stanici istog dana ume da dobije isti TEKST broja.
    stariBroj = Trim$(NzToText(LookupValue(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, _
                                           stariAmbDokID, COL_AMBD_BROJ)))
    RequireNovBrojIspravke broj, stariBroj, SRC
    brojK = ReversBrojZaUpis(datum, broj, stanicaID, SRC)
    RequireNovBrojIspravke brojK, stariBroj, SRC

    Set tx = New clsTransaction
    tx.BeginTx
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA
    ' Trag ispravke je deo ovog poteza: pad posle traga ne sme da ostavi red koji
    ' tvrdi zamenu koja se nije desila. CreateCorrectionContext otvara SVOJU
    ' transakciju, pa je snapshot ovde ono sto ga vraca.
    tx.AddTableSnapshot TBL_STORNO_VEZE

    StornirajAmbDokument tx, stariAmbDokID
    noviID = UpisiReversUTx(tx, red, datum, brojK, stanicaID, tipAmb, kolicina, _
                            smer, kooperantID, vozacID, napomena, SRC)

    cid = modStornoContext.CreateCorrectionContext( _
        SV_MODE_ISPRAVKA, DOK_TIP_AMBALAZA_DOKUMENT, stariAmbDokID, stariBroj, _
        DOK_TIP_AMBALAZA_DOKUMENT, noviID, brojK, , , , _
        "Ispravka reversa: " & stariBroj & " -> " & brojK & ".")
    If Len(cid) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Trag ispravke (tblStornoVeze) nije upisan."
    End If

    ' Veza na OBA kraja. Redovi se traze ponovo: upis je pomerio tabelu.
    RequireTacnoJedan TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, noviID, "Nov revers", SRC
    RequireTacnoJedan TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, stariAmbDokID, "Stari revers", SRC
    rNovi = CLng(FindRows(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, noviID)(1))
    rStari = CLng(FindRows(TBL_AMBALAZA_DOKUMENT, COL_AMBD_ID, stariAmbDokID)(1))
    RequireUpdateCell TBL_AMBALAZA_DOKUMENT, rNovi, COL_TRACE_ISPRAVKA_OD_ID, stariAmbDokID, SRC
    RequireUpdateCell TBL_AMBALAZA_DOKUMENT, rStari, COL_TRACE_ZAMENJEN_SA_ID, noviID, SRC

    ' Povratna vrednost se CITA: zatvaranje traga koje tiho ne uspe ostavilo bi
    ' PENDING red nad zavrsenom ispravkom -- MANUAL zadatak koji niko ne treba.
    If Not modStornoContext.CompleteCorrectionContext(cid, noviID, brojK, _
            "Ispravka reversa zavrsena: nov dokument " & noviID & ".") Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Trag ispravke nije zatvoren -- vidi Log."
    End If

    If mIspravkaPadPreCommit And IsTestMode() Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "TEST: podmetnut pad ispravke pre commit-a."
    End If

    tx.CommitTx
    Set tx = Nothing

    IspravkaReversa_TX = noviID
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0
    LogError SRC, errDesc, errNum
    Err.Raise errNum, SRC, errDesc
End Function


' POSLOVNI RED AMBALAZNOG DOKUMENTA -- za citaoce koji prikazuju DOKUMENT.
'
' Zaglavlje nema ni strane dogadjaja ni kolicinu ni tip ambalaze: po
' AMB-10-ODL-3 one zive na REDU knjige, a kopija na zaglavlju bila bi DRUGA
' ISTINA koja moze da se razidje. Ekran koji lista dokumente ih zato uzima
' odavde, a ne iz zaglavlja.
'
' JEDAN PROLAZ, jer mreza crta stotine redova; po redu dokumenta bi bio sken
' po redu.
'
' Sta se NE racuna u poslovni red, i zasto:
'   kontra-stav (StornoOd)   nije dogadjaj nego njegov inverz
'   ULAZ_TUDJE_AMBALAZE      pokrice deficita -- posledica, ne posao dokumenta
'                            (AmbVrstaJeZahtev je i zove 'posledica')
' Ostaje POSLOVNI red. Danasnji pisci ambalaznog dokumenta pisu tacno jedan, pa
' se uzima PRVI -- ali AMB-INV-10 dozvoljava i vise (dva tipa ambalaze, oba
' smera prema istom partneru). Mapa je PRIKAZ i tu pretpostavku sme da nosi;
' kapija koja na njoj odlucuje ne sme (IspravkaReversaRazlog broji kroz
' AmbDokBrojPoslovnihRedova, review #414 P2), a ni papir: stampa bira red po
' (dokument, tip) iz AmbDokPoslovniRedovi (KI-013). Definicija reda je JEDNA
' (PoslovniRedDokumenta), i oblik niza je JEDAN (PoslovniRedNiz).
'
' Vrednost je PAKOVAN STRING, ne ugnjezden objekat: late-bound Dictionary nema
' Property Set, pa `Set d(kljuc) = objekat` puca.
'
'   AmbDokID -> vrstaKretanja | tipAmb | kolicina | odTip | odID | naTip | naID
Public Function AmbDokRedMapa() As Object
    Const SRC As String = "modAmbalaza.AmbDokRedMapa"

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set AmbDokRedMapa = d

    On Error GoTo EH

    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function

    Dim kol As Object, vrste As Object
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    Dim i As Long, dokID As String
    For i = 1 To UBound(data, 1)
        dokID = PoslovniRedDokumenta(data, i, kol)
        If Len(dokID) > 0 Then
            If Not d.Exists(dokID) Then d.Add dokID, PoslovniRedNiz(data, i, kol)
        End If
    Next i

    Exit Function
EH:
    ' Citalac liste ne sme da obori ekran: prazna mapa daje dokument bez detalja,
    ' a razlog ostaje u logu. Kapije koje sude ispravnost knjige su drugde.
    LogErr SRC
End Function

' POSLOVNI RED AMBALAZNOG DOKUMENTA -- JEDNA definicija, za mapu prikaza
' (AmbDokRedMapa) i za kapiju ispravke (AmbDokBrojPoslovnihRedova). Dve kopije
' bi se razisle, a razlika bi se videla tek kao ispravka koja povuce red koji
' forma nije prikazala.
'
' Vraca AmbDokID kome red (data, i) pripada kao POSLOVNI red, inace "".
' Nije poslovni red: red van knjige, red tudjeg tipa dokumenta, kontra-stav
' (StornoOd) i pokrice deficita (ULAZ_TUDJE_AMBALAZE -- posledica, ne posao).
Private Function PoslovniRedDokumenta(ByRef data As Variant, ByVal i As Long, _
                                      ByRef kol As Object) As String
    If Not RedDoticeKnjigu(data, i, kol) Then Exit Function
    If StrComp(AmbText(data(i, kol(COL_AMB_DOK_TIP))), DOK_TIP_AMBALAZA_DOKUMENT, _
               vbTextCompare) <> 0 Then Exit Function
    If Len(AmbText(data(i, kol(COL_AMB_STORNO_OD)))) > 0 Then Exit Function
    If StrComp(AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA))), AMB_VK_ULAZ_TUDJE, _
               vbTextCompare) = 0 Then Exit Function
    PoslovniRedDokumenta = AmbText(data(i, kol(COL_AMB_DOK_ID)))
End Function

' POSLOVNI RED KAO NIZ -- jedan oblik za mapu prikaza (AmbDokRedMapa) i za
' stampu po tipu (AmbDokPoslovniRedovi). Indeksi su ugovor citaoca:
'   vrstaKretanja | tipAmb | kolicina | odTip | odID | naTip | naID
' NIZ, ne spojen string: TipAmbalaze je operaterski podatak iz lookup tabele,
' pa i on sme da nosi '|' -- ista rupa kao kod poslovnog broja (P2, 08.10.2026).
Private Function PoslovniRedNiz(ByRef data As Variant, ByVal i As Long, _
                                ByRef kol As Object) As Variant
    PoslovniRedNiz = Array(AmbText(data(i, kol(COL_AMB_VRSTA_KRETANJA))), _
                           AmbText(data(i, kol(COL_AMB_TIP))), _
                           CDbl(nz(data(i, kol(COL_AMB_KOLICINA)), 0)), _
                           AmbText(data(i, kol(COL_AMB_OD_TIP))), _
                           AmbText(data(i, kol(COL_AMB_OD_ID))), _
                           AmbText(data(i, kol(COL_AMB_NA_TIP))), _
                           AmbText(data(i, kol(COL_AMB_NA_ID))))
End Function

' SVI POSLOVNI REDOVI JEDNOG DOKUMENTA -- STROGO, za papir (KI-013).
'
' AmbDokRedMapa drzi PRVI poslovni red, jer je prikaz. Papir na to ne sme da se
' osloni: AMB-INV-10 dozvoljava dokument sa dva tipa ambalaze i sa oba smera
' prema istom partneru, a papir je po (dokument, tip) (AMB-10-ODL-30). Red bira
' pozivalac, iz SVIH redova.
'
' Greska citanja knjige se DIZE, kao u AmbDokBrojPoslovnihRedova: papir za
' potpis ne sme da nastane iz knjige koja nije procitana do kraja. Element je
' niz oblika PoslovniRedNiz. Prazan ID vraca praznu kolekciju -- inace bi se
' poklopio sa svakim redom koji NIJE poslovni.
Public Function AmbDokPoslovniRedovi(ByVal ambDokID As String) As Collection
    Const SRC As String = "modAmbalaza.AmbDokPoslovniRedovi"
    Dim data As Variant, kol As Object, vrste As Object
    Dim i As Long, c As Collection

    Set c = New Collection
    Set AmbDokPoslovniRedovi = c
    ambDokID = Trim$(ambDokID)
    If Len(ambDokID) = 0 Then Exit Function

    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then Exit Function
    Set kol = KnjigaZaCitanje(data, SRC, vrste)

    For i = 1 To UBound(data, 1)
        If StrComp(PoslovniRedDokumenta(data, i, kol), ambDokID, vbTextCompare) = 0 Then
            c.Add PoslovniRedNiz(data, i, kol)
        End If
    Next i
End Function

' Broj POSLOVNIH redova jednog ambalaznog dokumenta -- STROGO.
'
' Za razliku od AmbDokRedMapa (prikaz, guta greske), ovo je ulaz u KAPIJU:
' greska u citanju knjige se DIZE, jer kapija ne sme da procita "jedan" iz
' knjige koju nije procitala do kraja. Integritet knjige ide kroz isti ulaz kao
' kod svakog citaoca (KnjigaZaCitanje).
Private Function AmbDokBrojPoslovnihRedova(ByVal ambDokID As String) As Long
    AmbDokBrojPoslovnihRedova = AmbDokPoslovniRedovi(ambDokID).count
End Function

' POVRAT PRAZNIH OD KUPCA BEZ PRIJEMNICE -- KUPCEV DOKUMENT (AMB-10-ODL-23).
'
' Presuda operatera 06.10.2026: kupac vraca prazne gajbe i bez prijemnice, i to
' redovno. Takav povrat je NJEGOV dokument -- nosi njegov broj, vrsta je
' REVERS_PARTNERA, par je Kupac -> Vozac (lanac kupac -> vozac -> stanica,
' AMB-10-ODL-9), kretanje POVRAT_PRAZNE.
'
' ZASTO NIJE PETI SMER U AmbReversSmerovi. Ta mapa je mapa NASEG reversa: njen
' vlasnik broja je stanica, pa bi peti red tiho uveo dokument sa TUDJIM brojem i
' drugom vrstom u mapu koja o njima ne zna nista. Dva pisca se ovde razlikuju po
' vlasniku niza, ne po stilu.
'
' SVE INVARIJANTE SU U JEZGRU, ne ovde: AmbDokKretanjeProblem sudi par i
' vlasnika broja (grana jePartnerov), AmbDokDozvoljavaKretanje pusta samo
' POVRAT_PRAZNE uz ovu vrstu, UpisiAmbDokument meri zauzetost broja u opsegu
' (Kupac, KupacID, dan) po ODL-20, a PrenesiAmbalazu drzi AMB-INV-04 i -07.
' Ovaj pisac zato nosi tacno jedno pravilo koje nigde drugde ne postoji: BROJ JE
' OBAVEZAN I NE PREDLAZE SE.
'
' STORNO JE VEC POKRIVEN: StornirajAmbDokument_TX radi nad svakim ambalaznim
' dokumentom, pa ova vrsta ne trazi svoju putanju.
'
' Telo deli sa pozajmicom (UpisiKupcevDokument_TX): isti papir, isti broj, isti
' par -- razlikuje se samo vrsta kretanja.
Public Function UpisiReversPartnera_TX(ByVal datum As Date, ByVal broj As String, _
                                       ByVal kupacID As String, _
                                       ByVal vozacID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal kolicina As Double, _
                                       Optional ByVal napomena As String = "", _
                                       Optional ByVal potvrdaDeficita As Double = -1) As String
    UpisiReversPartnera_TX = UpisiKupcevDokument_TX(AMB_VK_POVRAT_PRAZNE, datum, broj, _
                                                    kupacID, vozacID, tipAmb, kolicina, _
                                                    napomena, potvrdaDeficita, _
                                                    "modAmbalaza.UpisiReversPartnera_TX")
End Function

' POZAJMICA AMBALAZE OD KUPCA -- KUPCEV DOKUMENT (AMB-10-ODL-24, odluke operatera
' 19.10.2026).
'
' Kupci pre sezone predaju SVOJE prazne gajbe vozacu. Papir je kupcev, kao kod
' povrata (REVERS_PARTNERA, njegov broj, par kupac -> vozac), a red nosi vrstu
' POZAJMICA_PARTNERA. Jezgro uz nju upisuje ULAZ_TUDJE iste kolicine
' (AMB-INV-13), pa kupcev saldo ostaje isti, a obaveza raste za kolicinu.
'
' POTVRDE MANJKA NEMA, i to je odluka (ODL-24 (f)): pozajmica ne gleda kupcev
' saldo, pa ni manjka nema. Zato ni parametra za nju nema.
Public Function UpisiPozajmicuKupca_TX(ByVal datum As Date, ByVal broj As String, _
                                       ByVal kupacID As String, _
                                       ByVal vozacID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal kolicina As Double, _
                                       Optional ByVal napomena As String = "") As String
    UpisiPozajmicuKupca_TX = UpisiKupcevDokument_TX(AMB_VK_POZAJMICA, datum, broj, _
                                                    kupacID, vozacID, tipAmb, kolicina, _
                                                    napomena, -1, _
                                                    "modAmbalaza.UpisiPozajmicuKupca_TX")
End Function

' JEDNO TELO ZA OBA KUPCEVA DOKUMENTA -- povrat (ODL-23) i pozajmicu (ODL-24).
' Pravila su ista jer je papir isti; vrstu kretanja bira javni ulaz, a sve ostalo
' sudi jezgro (AmbDokDozvoljavaKretanje, AmbDokKretanjeProblem, PrenesiAmbalazu).
Private Function UpisiKupcevDokument_TX(ByVal vrstaKretanja As String, _
                                        ByVal datum As Date, ByVal broj As String, _
                                        ByVal kupacID As String, _
                                        ByVal vozacID As String, _
                                        ByVal tipAmb As String, _
                                        ByVal kolicina As Double, _
                                        ByVal napomena As String, _
                                        ByVal potvrdaDeficita As Double, _
                                        ByVal SRC As String) As String
    Dim tx As clsTransaction
    Dim dokID As String, brojK As String
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    If kolicina <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Kupcev dokument trazi pozitivnu kolicinu."
    End If

    ' BROJ JE OBAVEZAN I NE PREDLAZE SE -- to je jedino pravilo koje je ovde, i
    ' jedina razlika prema UpisiReversAmbalaze_TX, koji na prazan broj zove
    ' generator. Predlog iz NASEG niza bio bi izmisljen broj TUDJE serije: papir
    ' koji operater drzi u ruci nosi kupcev broj, pa bi nas predlog bio drugi broj
    ' za isti dokument -- i ni jedan ne bi bio onaj po kome se dokument trazi.
    brojK = Trim$(broj)
    If Len(brojK) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Broj je obavezan i ne predlaze se: dokument je kupcev."
    End If

    ' Oba naloga su obavezna, i razlog je razlicit za svaki: bez kupca nema
    ' vlasnika broja, bez vozaca nema odredista lanca.
    If Len(Trim$(kupacID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Kupcev dokument trazi nalog Kupac -- on je i vlasnik broja."
    End If
    If Len(Trim$(vozacID)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Kupcev dokument trazi nalog Vozac -- lanac ide kupac -> vozac."
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    ' Obe tabele: zaglavlje i knjiga nastaju u ISTOM rollback-u (AMB-INV-08).
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    dokID = UpisiAmbDokument(tx, AMB_DOK_REVERS_PARTNERA, brojK, datum, _
                             AMB_NALOG_KUPAC, Trim$(kupacID), napomena)

    PrenesiAmbalazu tx, datum, tipAmb, kolicina, _
                    AMB_NALOG_KUPAC, Trim$(kupacID), _
                    AMB_NALOG_VOZAC, Trim$(vozacID), _
                    vrstaKretanja, DOK_TIP_AMBALAZA_DOKUMENT, dokID, _
                    potvrdaDeficita

    tx.CommitTx
    Set tx = Nothing

    UpisiKupcevDokument_TX = dokID
    Exit Function

EH:
    errNum = Err.Number
    errDesc = Err.description
    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0

    ' POTVRDA DEFICITA NIJE KVAR NEGO PITANJE POZIVAOCU (6.5), pa izlazi PRE
    ' loga -- isti obrazac kao CreateOtkup_TX. Da ide kroz LogError, log bi
    ' prestao da bude signal: kupac koji vrati vise nego sto knjiga kaze je
    ' REDOVAN slucaj, ne kvar. Greska se i dalje DIZE, sa istim brojem, pa
    ' pozivalac prepoznaje slucaj po BROJU a ne po tekstu.
    If errNum <> AMB_ERR_POTVRDA_DEFICITA Then LogError SRC, errDesc, errNum
    Err.Raise errNum, SRC, errDesc
End Function

' Red zatvorene mape za trazeni smer. Nepoznat smer je FAIL-CLOSED, i poruka
' nabraja sta je dozvoljeno -- isto kao zateceni Case Else u SaveOMUlaz_TX, samo
' sto spisak sada dolazi IZ MAPE, pa ne moze da se razidje sa njom.
Private Function ReversSmerRed(ByVal smer As String, ByVal SRC As String) As Variant
    Dim mapa As Variant, i As Long, spisak As String
    mapa = modAmbalazaUgovor.AmbReversSmerovi()
    For i = LBound(mapa) To UBound(mapa)
        If StrComp(Trim$(smer), CStr(mapa(i)(0)), vbTextCompare) = 0 Then
            ReversSmerRed = mapa(i)
            Exit Function
        End If
        If Len(spisak) > 0 Then spisak = spisak & ", "
        spisak = spisak & CStr(mapa(i)(0))
    Next i
    Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
              "Nepoznat smer reversa '" & Trim$(smer) & "'. Dozvoljeni: " & spisak & "."
End Function

' Koji ID popunjava stranu datog tipa. Prazan ID je FAIL-CLOSED: stari pisac je
' imao po dve takve provere u svakom od cetiri Case bloka -- osam kopija jednog
' pravila. Ovde je jedno telo, pa ne moze da se razidje po granama.
Private Function ReversNalogID(ByVal nalogTip As String, ByVal stanicaID As String, _
                               ByVal kooperantID As String, ByVal vozacID As String, _
                               ByVal smer As String, ByVal SRC As String) As String
    Select Case Trim$(nalogTip)
        Case AMB_NALOG_STANICA: ReversNalogID = Trim$(stanicaID)
        Case AMB_NALOG_KOOPERANT: ReversNalogID = Trim$(kooperantID)
        Case AMB_NALOG_VOZAC: ReversNalogID = Trim$(vozacID)
    End Select

    If Len(ReversNalogID) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "Revers (" & Trim$(smer) & ") trazi nalog " & Trim$(nalogTip) & _
                  ", a nije zadat."
    End If
End Function

' ============================================================
' NABAVKA -- jedini put kojim NASE gajbe ulaze u opticaj
' ============================================================
'
' Postoji jer AMB-INV-07 i AMB-10-ODL-8 ne dozvoljavaju stanici da izda gajbe
' koje nema: njen manjak nije "tudja ambalaza" nego SKRIVEN MANJAK, pa je put
' NABAVKA, sa svojim dokumentom. Na svezoj instalaciji svaka stanica pocinje od
' nule, pa bi bez ovog ulaza prvo izdavanje praznih bilo odbijeno. Odluka
' operatera 03.10.2026; ulaz sa ekrana (F7, sedmi segment) od AMB-10-ODL-34.
' Cene nema (ODL-34 d): ambalazni dokument ne nosi novac (ODL-5).
'
' VLASNIK BROJA JE STANICA, izvedeno a ne izmisljeno: AMB-10-ODL-3 trazi jedan
' protivpartner po dokumentu, a 6.9 izricito kaze da bi "jedan NABAVKA dokument
' nad dve stanice" tu invarijantu odmah prekrsio. Nalog je samo stanica (ODL-34
' b): Firma nema nijedan tok koji iz nje izdaje.
'
' NEMA PROTOKOLA POTVRDE: AmbDeficitZaPrenos vraca 0 za sve sto nije REALAN, a
' SpoljniSvet je iz te klase iskljucen (6.3 -- granica nema fizicko stanje), pa
' prenos IZ nje nikad nije u manjku.
'
' Broj se moze proslediti ili izracunati. Oba puta idu kroz istu kapiju
' zaglavlja -- a ona od 03.10.2026 sudi i ZAUZETOST u nizu (BrojOwnerTip,
' BrojOwnerID, dan), ne samo oblik. Zato prosledjen broj nije povlasten: ako je
' zauzet, upis pada.
Public Function NabaviAmbalazu_TX(ByVal datum As Date, _
                                  ByVal stanicaID As String, _
                                  ByVal tipAmb As String, _
                                  ByVal kolicina As Double, _
                                  Optional ByVal broj As String = "", _
                                  Optional ByVal napomena As String = "") As String
    NabaviAmbalazu_TX = UpisiNasJednoredni_TX(AMB_DOK_NABAVKA, AMB_VK_NABAVKA, datum, stanicaID, _
                                              AMB_NALOG_SPOLJNI, "", _
                                              AMB_NALOG_STANICA, stanicaID, _
                                              tipAmb, kolicina, broj, napomena, False, _
                                              "Nabavka", "modAmbalaza.NabaviAmbalazu_TX")
End Function

' ============================================================
' OTPIS -- jedini put kojim gajbe IZLAZE iz opticaja (KI-026)
' ============================================================
'
' Polomljena ili izgubljena gajba ide nalog -> SpoljniSvet, sa BILO KOG realnog
' naloga: stanice, vozaca, kooperanta ili kupca (AMB-10-ODL-34 e; matrica REALAN
' -> GRANICA). Bez ovoga bi ostala na saldu zauvek, a "ukupno u opticaju" bi
' raslo od stvarnog stanja sa svakim lomom. Zaobilaznica preko storna originala
' je POGRESNA: menja istoriju dokumenta koji je bio tacan.
'
' Dokument je NAS: broj iz niza stanice (ODL-34 c), kao i nabavka. Naknade nema
' (ODL-34 f): otpis nema finansijsku posledicu.
'
' RAZLOG JE OBAVEZAN I U PISCU (ODL-34 k), ne samo na ekranu: zapisnik bez
' razloga nije dokument, a pisac ima i pozivaoce mimo ekrana.
'
' Otpis preko salda se ODBIJA u jezgru (AMB-INV-14), bez pitanja: zato ovde nema
' parametra potvrde.
Public Function OtpisiAmbalazu_TX(ByVal datum As Date, _
                                  ByVal stanicaID As String, _
                                  ByVal nalogTip As String, _
                                  ByVal nalogID As String, _
                                  ByVal tipAmb As String, _
                                  ByVal kolicina As Double, _
                                  ByVal razlog As String, _
                                  Optional ByVal broj As String = "") As String
    OtpisiAmbalazu_TX = UpisiNasJednoredni_TX(AMB_DOK_OTPIS, AMB_VK_OTPIS, datum, stanicaID, _
                                              nalogTip, nalogID, _
                                              AMB_NALOG_SPOLJNI, "", _
                                              tipAmb, kolicina, broj, razlog, True, _
                                              "Otpis", "modAmbalaza.OtpisiAmbalazu_TX")
End Function

' JEDNO TELO ZA NASA DVA JEDNOREDNA DOKUMENTA -- nabavku i otpis (ODL-34).
' Oba su NASA (broj iz niza stanice) i nose tacno jedan red knjige jednog naloga
' prema granici (ODL-3). Pravila prenosa sudi jezgro (PrenesiAmbalazu:
' AMB-INV-07, -14), pa ovde zive samo kolicina, razlog i broj -- isti oblik kao
' kupcev par u UpisiKupcevDokument_TX.
Private Function UpisiNasJednoredni_TX(ByVal dokVrsta As String, _
                                       ByVal vrstaKretanja As String, _
                                       ByVal datum As Date, _
                                       ByVal stanicaID As String, _
                                       ByVal odTip As String, ByVal odID As String, _
                                       ByVal naTip As String, ByVal naID As String, _
                                       ByVal tipAmb As String, _
                                       ByVal kolicina As Double, _
                                       ByVal broj As String, _
                                       ByVal napomena As String, _
                                       ByVal napomenaObavezna As Boolean, _
                                       ByVal naziv As String, _
                                       ByVal SRC As String) As String
    Dim tx As clsTransaction
    Dim dokID As String, brojK As String
    Dim errNum As Long, errDesc As String

    On Error GoTo EH

    If kolicina <= 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  naziv & " trazi pozitivnu kolicinu."
    End If

    If napomenaObavezna And Len(Trim$(napomena)) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  naziv & " trazi razlog (lom, gubitak) -- AMB-10-ODL-34 (k)."
    End If

    brojK = Trim$(broj)
    If Len(brojK) = 0 Then
        brojK = modBrojevi.GenerateBrojAmbDokumenta(AMB_NALOG_STANICA, stanicaID, datum)
        If Len(brojK) = 0 Then
            Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                      "Broj dokumenta (" & naziv & ") nije izracunat -- vidi Log."
        End If
    End If

    Set tx = New clsTransaction
    tx.BeginTx
    ' Obe tabele: zaglavlje i knjiga nastaju u ISTOM rollback-u (AMB-INV-08).
    tx.AddTableSnapshot TBL_AMBALAZA_DOKUMENT
    tx.AddTableSnapshot TBL_AMBALAZA

    ' UpisiAmbDokument sam vezuje dokument za ovu transakciju.
    dokID = UpisiAmbDokument(tx, dokVrsta, brojK, datum, _
                             AMB_NALOG_STANICA, stanicaID, napomena)

    PrenesiAmbalazu tx, datum, tipAmb, kolicina, _
                    odTip, odID, naTip, naID, _
                    vrstaKretanja, DOK_TIP_AMBALAZA_DOKUMENT, dokID

    tx.CommitTx
    Set tx = Nothing

    UpisiNasJednoredni_TX = dokID
    Exit Function

EH:
    ' Err se cuva PRE rollback-a, i PONOVO se dize SA ISTIM BROJEM: pozivalac
    ' (UI) razlikuje kapije po kodu greske, pa bi zamena broja unistila tu
    ' informaciju -- i protokol potvrde deficita na drugim putanjama.
    errNum = Err.Number
    errDesc = Err.description
    LogErr SRC

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    Set tx = Nothing
    On Error GoTo 0

    Err.Raise errNum, SRC, errDesc
End Function

' AMB-INV-08: nijedan upis u knjigu ne nastaje van vlasnistva transakcije
' izvornog dokumenta.
'
' Ovo je bila STATICKA kapija u planu, i napisana je -- pa pala na svom drugom
' nivou dokaza. Pravilo "neki predak u pozivnom lancu poseduje tx sa snapshotom"
' je nad 4120 procedura i 17 vlasnika uvek istinito ako se ide dovoljno visoko:
' skinuta su oba AddTableSnapshot iz modOtkup, a kapija je ostala zelena. Alat
' koji se moze prevariti dubinom ne sme da cuva invarijantu.
'
' Zato pisac TRAZI transakciju i sam pita. Razlika nije stilska:
'   - poziv bez tx je COMPILE ERROR, ne nalaz koji se moze ignorisati;
'   - nema grafa, nema dubine, nema laznog zelenog ni laznog nalaza;
'   - sabotaza je prava: skini snapshot, pisac padne po imenu.
Private Sub RequireAmbTxVlasnistvo(ByVal tx As clsTransaction, _
                                   ByVal tabela As String, _
                                   ByVal sourceName As String)
    If tx Is Nothing Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: upis u " & tabela & " bez transakcije. Pisac " & _
                  "trazi clsTransaction izvornog dokumenta."
    End If
    If Not tx.ImaSnapshot(tabela) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: transakcija ne snapshotuje " & tabela & _
                  " -- rollback izvornog dokumenta ne bi vratio ambalazu."
    End If
End Sub

' Tacka 3 iste invarijante: izvorni dokument mora biti vezan BAS za ovu
' transakciju. Snapshot tabele dokazuje da se knjiga MOZE vratiti; ovo dokazuje
' da se vraca ZAJEDNO sa dokumentom koji ju je izazvao.
Private Sub RequireAmbTxIzvorniDokument(ByVal tx As clsTransaction, _
                                        ByVal dokTip As String, _
                                        ByVal dokID As String, _
                                        ByVal sourceName As String)
    If Not tx.OwnsSourceDocument(dokTip, dokID) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: izvorni dokument " & Trim$(dokTip) & " '" & _
                  Trim$(dokID) & "' nije vezan za ovu transakciju " & _
                  "(BindSourceDocument) -- knjiga i dokument ne dele vlasnika, " & _
                  "pa rollback dokumenta ne bi vratio ambalazu."
    End If

    ' Vezivanje dokazuje IDENTITET transakcije, ne i da je tabela tog dokumenta
    ' u njenom snapshotu. Bez ovoga prolazi bind nad Otkup-om uz snapshot samo
    ' knjige -- rollback tada nije zajednicki. Nepoznat tip je fail-closed.
    Dim izvornaTbl As String
    izvornaTbl = modAmbalazaUgovor.AmbIzvornaTabela(dokTip)
    If Len(izvornaTbl) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: tip izvornog dokumenta '" & Trim$(dokTip) & _
                  "' nema poznatu izvornu tabelu -- ne moze se dokazati da " & _
                  "knjiga i dokument dele rollback."
    End If
    If Not tx.ImaSnapshot(izvornaTbl) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, sourceName, _
                  "AMB-INV-08: transakcija ne snapshotuje " & izvornaTbl & _
                  " -- izvorni dokument je vezan, ali se ne bi vratio " & _
                  "rollback-om zajedno sa knjigom."
    End If
End Sub

' ============================================================
' PRENESI AMBALAZU -- javni ulaz u knjigu
' ============================================================
'
' Vraca AmbID glavnog reda. Jedan poziv moze da upise do TRI reda, i svaki je
' posledica imenovane invarijante:
'
'   pokrice deficita   SpoljniSvet -> izvor   ULAZ_TUDJE_AMBALAZE   (AMB-INV-07)
'   trazeni prenos     od -> na               trazena vrsta
'   ostatak podele     od -> na               IZDATA_PRAZNA         (AMB-INV-09)
'
' potvrdaDeficita: -1 znaci "nije data". Nula NIJE sentinela -- deficit nula ne
' trazi potvrdu, pa bi 0 bila dvosmislena vrednost.
Public Function PrenesiAmbalazu(ByVal tx As clsTransaction, _
                                ByVal datum As Date, ByVal tipAmb As String, _
                                ByVal kolicina As Double, _
                                ByVal odTip As String, ByVal odID As String, _
                                ByVal naTip As String, ByVal naID As String, _
                                ByVal vrsta As String, _
                                ByVal dokTip As String, ByVal dokID As String, _
                                Optional ByVal potvrdaDeficita As Double = -1) As String
    Const SRC As String = "modAmbalaza.PrenesiAmbalazu"

    RequireKnjigaSchema SRC
    modAmbalazaUgovor.RequireAmbPrenos odTip, odID, naTip, naID, kolicina, tipAmb, vrsta, SRC

    ' ULAZ_TUDJE_AMBALAZE nije zahtev nego POSLEDICA -- generise je ovaj pisac kao
    ' pokrice deficita i kao ulaz pozajmice (AMB-INV-13). Da je i zahtev, pokrice
    ' i eksplicitan zahtev nad istim dokumentom delili bi par i vrstu, pa bi
    ' jedan progutao drugi.
    If Not modAmbalazaUgovor.AmbVrstaJeZahtev(vrsta) Then
        Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                  "'" & Trim$(vrsta) & "' nije zahtev nego posledica -- pisac je generise sam."
    End If

    ' IDENTITET DOKUMENTA SE NE PROVERAVA OVDE. Ista provera stoji u
    ' UpisiRedKnjige, kroz koji prolazi SVAKI red -- i pokrice deficita i ostatak
    ' podele. Kopija na ulazu bila bi placebo: nijedna sabotaza je ne bi mogla
    ' oboriti, jer bi jezgro odgovorilo isto. Invarijanta zivi u jezgru.
    Dim vrstaK As String
    vrstaK = modAmbalazaUgovor.AmbVrstaKanon(vrsta)

    ' VEZA DOKUMENT <-> KRETANJE. Vazi samo za tblAmbalazaDokument: za robne
    ' dokumente (otkup, otpremnica, prijemnica) tipovi nisu zatvoren skup, pa
    ' ugovor o njima namerno ne tvrdi nista.
    Dim dokVrsta As String, dokOwnerTip As String, dokOwnerID As String
    If StrComp(Trim$(dokTip), DOK_TIP_AMBALAZA_DOKUMENT, vbTextCompare) = 0 Then
        AmbDokZaglavlje dokID, dokVrsta, dokOwnerTip, dokOwnerID
        If Not modAmbalazaUgovor.AmbDokDozvoljavaKretanje(dokVrsta, vrstaK) Then
            Err.Raise AMB_ERR_KNJIGA_ULAZ, SRC, _
                      "Dokument vrste " & dokVrsta & " ne nosi kretanje " & vrstaK & "."
        End If
    End If

    ' AMB-10-ODL-22: robni dokument koji je I SAM partnerov objavljuje vlasnika
    ' svog broja. Prijemnica je takav: eksterna je, njen broj je kupcev, i
    ' povrat praznih se knjizi POD TIM brojem (operater, 05.10.2026). Bez ovoga
    ' bi obrnuta kapija odbila prijemnicu -- merila bi VRSTU dokumenta kao
    ' zamenu za vlasnika broja, a to je zamena koju je ODL-22 pobio.
    If Len(dokOwnerTip) = 0 Then
        AmbRobniZaglavlje dokTip, dokID, dokOwnerTip, dokOwnerID
    End If

    ' AMB-10-ODL-9/-10: vrsta, vlasnik broja i par naloga se gledaju ZAJEDNO.
    ' Obrnuta kapija vazi i kad dokument NIJE ambalazni (dokVrsta ostaje prazna):
    ' povrat praznih od kupca mora da nosi kupcev broj, pa ga dokument koji
    ' vlasnika broja ne objavi ne sme nositi.
    Dim parProblem As String
    parProblem = modAmbalazaUgovor.AmbDokKretanjeProblem( _
        dokVrsta, dokOwnerTip, dokOwnerID, odTip, odID, naTip, vrstaK)
    If Len(parProblem) > 0 Then
        Err.Raise AMB_ERR_IDENTITET, SRC, parProblem
    End If

    ' AMB-INV-04, idempotencija: isti zahtev nad istim dokumentom i parem naloga.
    Dim vecUpisano As Double, vecID As String
    vecUpisano = ZbirZahteva(dokTip, dokID, datum, tipAmb, odTip, odID, naTip, naID, _
                             vrstaK, vecID, SRC)
    If Len(vecID) > 0 Then
        If vecUpisano = kolicina Then
            PrenesiAmbalazu = vecID
            Exit Function
        End If

        Err.Raise AMB_ERR_IDENTITET, SRC, _
                  "AMB-INV-04: isti dogadjaj je vec knjizen sa " & CStr(vecUpisano) & _
                  ", a trazi se " & CStr(kolicina) & " (" & Trim$(dokTip) & " '" & _
                  Trim$(dokID) & "', " & vrstaK & ")."
    End If

    ' AMB-INV-07: nijedan realan nalog ne sme posle commit-a biti ispod nule.
    Dim deficit As Double, pokriceZaUpis As Double
    If modAmbalazaUgovor.AmbVrstaUnosiUOpticaj(vrstaK) Then
        ' AMB-INV-13 (ODL-24): pozajmica unosi u opticaj TACNO svoju kolicinu,
        ' bez obzira na partnerov saldo. Deficita zato nema, pa ni pitanja: saldo
        ' partnera ostaje kakav je bio. Ulaz se upisuje POSLE trazenog reda, istim
        ' putem kao pokrice -- AMB-INV-10 trazi zakljucan par pre njega.
        pokriceZaUpis = kolicina
    Else
        deficit = AmbDeficitZaPrenos(odTip, odID, tipAmb, kolicina)
    End If

    If deficit > 0 Then
        ' AMB-INV-14 (AMB-10-ODL-34 g): OTPIS SE NE POKRIVA. Pitanje o manjku i
        ' ULAZ_TUDJE bi uneli gajbe samo da bi se otpisale, a firma bi ih partneru
        ' dugovala. Odbija se PRE pokrica i PRE potvrde, pa ni potvrda jednaka
        ' manjku ne prolazi. Pravilo je u ugovoru (AmbVrstaSmePokrice).
        If Not modAmbalazaUgovor.AmbVrstaSmePokrice(vrstaK) Then
            Err.Raise AMB_ERR_DEFICIT_NEPOKRIV, SRC, _
                      "AMB-INV-14: " & vrstaK & " ne sme vise nego sto nalog drzi -- " & _
                      Trim$(odTip) & " '" & Trim$(odID) & "' drzi " & CStr(kolicina - deficit) & _
                      ", trazi se " & CStr(kolicina) & " '" & Trim$(tipAmb) & "'."
        End If

        Dim pokrice As String
        pokrice = modAmbalazaUgovor.AmbPokriceProblem(odTip, odID)
        If Len(pokrice) > 0 Then
            Err.Raise AMB_ERR_DEFICIT_NEPOKRIV, SRC, _
                      pokrice & " Manjak: " & CStr(deficit) & " '" & Trim$(tipAmb) & "'."
        End If

        If potvrdaDeficita < 0 Then
            Err.Raise AMB_ERR_POTVRDA_DEFICITA, SRC, _
                      "Nalog " & Trim$(odTip) & " '" & Trim$(odID) & "' nema " & CStr(kolicina) & _
                      " '" & Trim$(tipAmb) & "'. Manjak " & CStr(deficit) & " ulazi u opticaj kao " & _
                      "tudja ambalaza -- potvrdi tacno taj broj."
        End If

        ' Potvrda se meri prema SVEZE izracunatom deficitu: izmedju pitanja i
        ' odgovora stanje se moglo promeniti drugim unosom, pa stara potvrda ne vazi.
        If potvrdaDeficita <> deficit Then
            Err.Raise AMB_ERR_POTVRDA_DEFICITA, SRC, _
                      "Potvrda " & CStr(potvrdaDeficita) & " se ne slaze sa manjkom " & _
                      CStr(deficit) & " -- stanje se promenilo, potvrdi ponovo."
        End If

        ' POKRICE SE UPISUJE POSLE TRAZENOG PRENOSA, i to je sada deo ugovora:
        ' AMB-INV-10 poredi odredisten pokrica sa ZAKLJUCANIM parom dokumenta, pa par
        ' mora postojati pre njega. Redosled unutar transakcije je slobodan --
        ' AMB-INV-07 govori o stanju POSLE commit-a, ne o putu do njega.
        pokriceZaUpis = deficit
    End If

    ' AMB-INV-09: ne moze se vratiti vise tudje ambalaze nego sto je uzeto. Visak
    ' nije vracanje nego NOVO zaduzenje partnera -- podelu radi pisac, jer samo on
    ' cita obavezu pouzdano u trenutku upisa.
    Dim glavna As Double, ostatak As Double
    glavna = kolicina
    ostatak = 0

    If StrComp(vrstaK, AMB_VK_VRACANJE_TUDJE, vbTextCompare) = 0 Then
        Dim obaveza As Double
        obaveza = AmbObavezaPartneru(naTip, naID, tipAmb)

        ' NEGATIVNA OBAVEZA SE NE SPUSTA NA NULU, NEGO PADA. Ona znaci da je
        ' AMB-INV-09 prekrsena PRE ovog poziva, a tisina bi je pretvorila u
        ' normalno stanje. Do njega 10d MOZE da dodje: storno ULAZA tudje
        ' ambalaze cija je obaveza vec zatvorena vracanjem daje -N. Odgovor na to
        ' je da se takav storno odbije, i to je zahtev koji 10d nasledjuje -- ne
        ' da ga ovaj pisac prekrije.
        If obaveza < 0 Then
            Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, _
                      "Obaveza prema " & Trim$(naTip) & " '" & Trim$(naID) & "' je " & _
                      CStr(obaveza) & " -- AMB-INV-09 je prekrsena pre ovog poziva."
        End If

        If kolicina > obaveza Then
            glavna = obaveza
            ostatak = kolicina - obaveza
        End If
    End If

    Dim glavniID As String
    If glavna > 0 Then
        glavniID = UpisiRedKnjige(tx, datum, tipAmb, glavna, odTip, odID, naTip, naID, _
                                  dokTip, dokID, vrstaK, "", SRC)
    End If

    If ostatak > 0 Then
        Dim ostatakID As String
        ostatakID = UpisiRedKnjige(tx, datum, tipAmb, ostatak, odTip, odID, naTip, naID, _
                                   dokTip, dokID, AMB_VK_IZDATA_PRAZNA, "", SRC)
        If Len(glavniID) = 0 Then glavniID = ostatakID
    End If

    If Len(glavniID) = 0 Then
        Err.Raise AMB_ERR_KNJIGA_KVAR, SRC, "Prenos nije upisao ni jedan red."
    End If

    If pokriceZaUpis > 0 Then
        UpisiRedKnjige tx, datum, tipAmb, pokriceZaUpis, _
                       AMB_NALOG_SPOLJNI, "", odTip, odID, _
                       dokTip, dokID, AMB_VK_ULAZ_TUDJE, "", SRC
    End If

    PrenesiAmbalazu = glavniID
End Function
