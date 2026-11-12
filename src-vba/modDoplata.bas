Attribute VB_Name = "modDoplata"
Option Explicit

' ============================================================
' modDoplata - DOPLATA KOOPERANTU (knjizno odobrenje, KO)
'
' Domen i odluke: docs/DOMEN/DOPLATA.md (DOP-ODL-*, DOP-*).
'
' Globalni unos: stopa po KLJUCU CENOVNIKA (VrstaVoca + SortaVoca + Klasa) za
' period otkupa, na jedan od dva nacina (DOP-ODL-15): bruto doplata po kg, ili
' konacna bruto cena po kg. Obracun pravi JEDAN DOKUMENT DOPLATE PO OTKUPNOM
' BLOKU (stavke po klasi), svi pod istim ObracunID-em, u JEDNOJ transakciji;
' operater pre knjizenja moze da iskljuci blokove (DOP-ODL-16). Cena stavke je
' BRUTO doplata po kg -- sa PDV nadoknadom, isti model kao cena otkupa
' (DOP-ODL-14). Iznos je ono sto se isplacuje.
'
' PRIKAZ I PISAC DELE JEDAN PLAN (obrazac BimAutoPlan): DoplataPlan racuna
' stavke bez ijednog upisa; ekran ga prikazuje, KnjiziDoplatu_TX ga izvrsava.
' Druga kopija racuna bi se razisla prvom izmenom.
'
' Sta doplata NE dira (DOP-01): robu -- kg doplate ne ulazi ni u jedan robni
' zbir --, cenu na stavkama otkupa, otkupni list, hladnjacki lanac. Knjizen KO
' je SNIMAK (DOP-ODL-07): kasniji storno ili ispravka otkupa ga ne menja.
'
' Storno zaglavlja je u modStorno (StornoDoplata_TX), kao za ostale dokumente.
' Stavka nema svoj storno (BEZ_STORNA): aktivna je dok je aktivno zaglavlje.
' ============================================================

' Greske pisca (opseg vbObjectError + 6300..6399; potvrda je poslednji broj
' opsega, da je nijedna kapija ne bi slucajno ponela).
'
' POTVRDA PONAVLJANJA NIJE KVAR NEGO PITANJE POZIVAOCU (DOP-ODL-06, obrazac
' AMB_ERR_POTVRDA_DEFICITA): pisac vraca "" bez loga i monitora, uz tekst koji
' imenuje raniju doplatu; pozivalac pita operatera i ponavlja poziv sa potvrdom.
Public Const DOP_ERR_POTVRDA_PONAVLJANJA As Long = vbObjectError + 6399
Private Const DOP_ERR_BASE As Long = vbObjectError + 6300

' Stopa je Scripting.Dictionary sa TACNO ovim kljucevima (zatvoren spisak, kao
' stavka otkupa): tipfeler u imenu kljuca ne sme da prodje kao "nije uneto".
Private Const DOP_STOPA_KLJUCEVI As String = "|VRSTAVOCA|SORTAVOCA|KLASA|CENA|"

' Minuta u danu: granica perioda se poredi na minut, isto koliko otkupni list
' stampa (hh:nn).
Private Const MIN_U_DANU As Long = 1440

' Kolone plana (DoplataPlan). Nove kolone idu NA KRAJ: citaoci indeksiraju
' poziciono.
Private Const PL_KOOP As Long = 1
Private Const PL_STANICA As Long = 2
Private Const PL_VRSTA As Long = 3
Private Const PL_SORTA As Long = 4
Private Const PL_KLASA As Long = 5
Private Const PL_KOLICINA As Long = 6
Private Const PL_CENA As Long = 7
Private Const PL_IZNOS As Long = 8
Private Const PL_BROJ_OTKUPA As Long = 9
Private Const PL_KLJUC As Long = 10
' Revizija po bloku (DOPLATA.md par. 2a):
Private Const PL_OTKUP As Long = 11
Private Const PL_BROJ_BLOKA As Long = 12
Private Const PL_CENA_OTKUPA As Long = 13
Private Const PL_KONACNA As Long = 14
Private Const PL_VEC_DOPLACENO As Long = 15
Private Const PL_STATUS As Long = 16
Private Const PL_DATUM_BLOKA As Long = 17
Private Const PL_KOLONA As Long = 17

' Status reda plana: knjizi se, ili je "bez doplate" (nacin 2: cena otkupa i
' vec knjizena doplata su vec na konacnoj ceni ili preko nje -- DOP-ODL-15).
Public Const DOP_PLAN_KNJIZI As Long = 1
Public Const DOP_PLAN_BEZ_DOPLATE As Long = 0

' Ishod kapije KO (KOKapija). Prazno = KO postoji TACNO jednom, aktivan je i
' pripada kooperantu. Svaki pisac koji vezuje novac za KO (izvod, gotovina,
' primena avansa) prevodi ishod u SVOJU gresku ili poruku -- pravilo je jedno.
Public Const KO_KAPIJA_NEMA As String = "NEMA"
Public Const KO_KAPIJA_DUPLI As String = "DUPLI"
Public Const KO_KAPIJA_STORNIRAN As String = "STORNIRAN"
Public Const KO_KAPIJA_TUDJ As String = "TUDJ"

' ============================================================
' PLAN -- bez ijednog upisa
' ============================================================

' Stavke obracuna -- JEDAN RED PO BLOKU I KLASI (DOPLATA.md par. 2a). 2D niz
' (1..n, 1..17):
'   1 KooperantID | 2 StanicaID (stanica bloka) | 3 VrstaVoca | 4 SortaVoca
'   5 Klasa | 6 Kolicina (kg, 2 dec.) | 7 Cena = BRUTO doplata po kg
'   8 Iznos (din, 2 dec.) | 9 broj otkupa u stavci (uvek 1) | 10 kljuc poretka
'   11 OtkupID | 12 broj bloka | 13 cena otkupa (bruto, sa stavke) | 14 konacna
'   cena (nacin 2; inace Empty) | 15 vec knjizena aktivna doplata po kg na tom
'   bloku i klasi | 16 status (DOP_PLAN_KNJIZI / DOP_PLAN_BEZ_DOPLATE) |
'   17 datum bloka
' poredano po kooperantu, stanici, datumu i broju bloka, klasi. Empty kad nema
' nijedne.
'
' nacin (DOP-ODL-15, jedan za ceo obracun): DOP_NACIN_PO_KG -- vrednost stope je
' bruto doplata po kg; DOP_NACIN_KONACNA -- vrednost stope je konacna bruto cena
' po kg, a doplata po kg = konacna - cena otkupa - vec knjizena doplata po kg.
' Red sa doplatom <= 0 ostaje u planu sa statusom "bez doplate": ekran ga
' pokazuje, pisac ga ne knjizi.
'
' outBezOtkupa: stope (vrsta|sorta|klasa) za koje u periodu nema nijednog
' aktivnog otkupa. Nije greska plana -- ekran ih pokazuje uz pregled -- ali pisac
' ih odbija: tipfeler u sorti bi inace tiho dao doplatu nula za ceo kljuc.
'
' Obuhvat (DOP-ODL-03): svaki AKTIVAN otkup, na svakom otkupnom mestu,
' ukljucujuci hladnjacke blokove. Kilogrami i cena otkupa su sa STAVKI
' (StavkeOtkupaRedovi, strog citac kanona) -- isti izvor koji citaju kartica i
' saldo.
'
' kooperantID (DOP-ODL-31): prazno = obracun za sve kooperante. Zadat = obracun
' SAMO za tog kooperanta: tudji blokovi ne ulaze u plan ni u jednom koraku, pa
' stopa koja ne pogadja nijedan NJEGOV blok ide u outBezOtkupa iako pogadja
' tudje, a blok drugog kooperanta na granicnom danu bez poznatog sata ne obara
' ovaj obracun.
Public Function DoplataPlan(ByVal periodOd As Date, ByVal periodDo As Date, _
                            ByVal stope As Collection, _
                            Optional ByRef outBezOtkupa As String, _
                            Optional ByVal nacin As Long = DOP_NACIN_PO_KG, _
                            Optional ByVal kooperantID As String = "") As Variant
    Const SRC As String = "modDoplata.DoplataPlan"

    DoplataPlan = Empty
    outBezOtkupa = ""

    Dim obuhvat As String
    obuhvat = Trim$(kooperantID)

    RequirePeriod periodOd, periodDo, SRC
    RequireNacin nacin, SRC

    Dim cene As Object, vrsteSorte As Object, prikaz As Object
    UcitajStope stope, cene, vrsteSorte, prikaz, SRC

    ' --- 1. Aktivna zaglavlja u periodu, za vrstu i sortu iz stopa -----------
    Dim d As Variant
    d = GetTableData(TBL_OTKUP)
    If Not IsArray(d) Then
        outBezOtkupa = SpisakKljuceva(prikaz, Nothing)
        Exit Function
    End If
    d = ExcludeStornirano(d, TBL_OTKUP)
    If Not IsArray(d) Then
        outBezOtkupa = SpisakKljuceva(prikaz, Nothing)
        Exit Function
    End If

    Dim cId As Long, cDat As Long, cKoop As Long, cSta As Long
    Dim cVr As Long, cSo As Long, cBr As Long, cSrc As Long, cCre As Long
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    cSta = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cVr = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)
    cSo = RequireColumnIndex(TBL_OTKUP, COL_OTK_SORTA, SRC)
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cSrc = RequireColumnIndex(TBL_OTKUP, COL_OTK_SOURCE_CREATED_AT, SRC)
    cCre = RequireColumnIndex(TBL_OTKUP, COL_AUDIT_CREATED_AT, SRC)

    Dim uPeriodu As Object
    Set uPeriodu = CreateObject("Scripting.Dictionary")
    uPeriodu.CompareMode = vbTextCompare

    Dim i As Long, oid As String, vs As String
    For i = 1 To UBound(d, 1)
        vs = KljucVrstaSorta(d(i, cVr), d(i, cSo))
        If vrsteSorte.Exists(vs) And UObuhvatu(d(i, cKoop), obuhvat) Then
            oid = Trim$(NzToText(d(i, cId)))
            If OtkupUPeriodu(d(i, cDat), d(i, cSrc), d(i, cCre), periodOd, periodDo, _
                             oid, NzToText(d(i, cBr)), SRC) Then
                ' Dupli OtkupID obara StavkeOtkupaRedovi pre agregacije; ovde
                ' bi tiho pobedio poslednji red.
                uPeriodu(oid) = Array(Trim$(NzToText(d(i, cKoop))), _
                                      Trim$(NzToText(d(i, cSta))), vs, _
                                      Trim$(NzToText(d(i, cBr))), CDate(d(i, cDat)), i)
            End If
        End If
    Next i

    ' --- 2. Stavke tih otkupa, po bloku i klasi -------------------------------
    Dim agg As Object, pokriveno As Object
    Set agg = CreateObject("Scripting.Dictionary")
    Set pokriveno = CreateObject("Scripting.Dictionary")
    pokriveno.CompareMode = vbTextCompare

    If uPeriodu.count > 0 Then
        Dim s As Variant
        s = modOtkup.StavkeOtkupaRedovi()
        If IsArray(s) Then
            Dim h As Variant, kStope As String, kAgg As String, v As Variant
            For i = 1 To UBound(s, 1)
                oid = CStr(s(i, 1))
                If uPeriodu.Exists(oid) Then
                    h = uPeriodu(oid)
                    kStope = CStr(h(2)) & "|" & UCase$(Trim$(CStr(s(i, 3))))
                    If cene.Exists(kStope) Then
                        pokriveno(kStope) = True
                        kAgg = KljucPoretka(CStr(h(0)), CStr(h(1)), CDate(h(4)), CLng(h(5)), CStr(s(i, 3)))
                        If agg.Exists(kAgg) Then
                            v = agg(kAgg)
                            ' Dve stavke iste klase u istom bloku: kanon otkupa to
                            ' NE dozvoljava -- pisac (CreateOtkup_TX, greska 1865) i
                            ' citalac (StavkeOtkupaRedovi, 1924, pada pre ovoga).
                            ' Grana je odbrana u dubinu za oba nacina (review P1,
                            ' DOP-P1-REV-02): sa razlicitom cenom nacin 2 ne bi
                            ' znao od cega racuna, pa pada po imenu umesto da bira.
                            If Abs(CDbl(v(4)) - CDbl(s(i, 5))) > 0.000001 Then
                                Err.Raise DOP_ERR_BASE + 32, SRC, _
                                          "Blok " & CStr(h(3)) & " (" & oid & ") ima dve stavke klase " & _
                                          CStr(s(i, 3)) & " sa razlicitom cenom."
                            End If
                        Else
                            v = Array(CStr(h(0)), CStr(h(1)), kStope, 0#, CDbl(s(i, 5)), oid, _
                                      CStr(h(3)), CDate(h(4)))
                        End If
                        v(3) = CDbl(v(3)) + CDbl(s(i, 4))
                        agg(kAgg) = v
                    End If
                End If
            Next i
        End If
    End If

    outBezOtkupa = SpisakKljuceva(prikaz, pokriveno)
    If agg.count = 0 Then Exit Function

    ' --- 3. Niz plana, poredan ------------------------------------------------
    Dim vec As Object
    Set vec = VecDoplacenoPoKg()

    Dim res() As Variant, n As Long, k As Variant, deo As Variant, kg As Double
    Dim unos As Double, cenaOtk As Double, vecKg As Double, dpk As Double, kVec As String
    ReDim res(1 To agg.count, 1 To PL_KOLONA)
    For Each k In agg.keys
        v = agg(k)
        deo = prikaz(CStr(v(2)))
        n = n + 1
        kg = modNovac.ZaokruziNovac(CDbl(v(3)))   ' kg na dve decimale, kao stavka otkupa
        unos = CDbl(cene(CStr(v(2))))
        cenaOtk = CDbl(v(4))
        kVec = UCase$(CStr(v(5))) & "|" & UCase$(Trim$(CStr(deo(2))))
        vecKg = 0
        If vec.Exists(kVec) Then vecKg = CDbl(vec(kVec))

        If nacin = DOP_NACIN_KONACNA Then
            dpk = modNovac.ZaokruziNovac(unos - cenaOtk - vecKg)
            res(n, PL_KONACNA) = unos
        Else
            dpk = unos
            res(n, PL_KONACNA) = Empty
        End If

        res(n, PL_KOOP) = v(0)
        res(n, PL_STANICA) = v(1)
        res(n, PL_VRSTA) = deo(0)
        res(n, PL_SORTA) = deo(1)
        res(n, PL_KLASA) = deo(2)
        res(n, PL_KOLICINA) = kg
        res(n, PL_BROJ_OTKUPA) = 1
        res(n, PL_KLJUC) = CStr(k)
        res(n, PL_OTKUP) = v(5)
        res(n, PL_BROJ_BLOKA) = v(6)
        res(n, PL_CENA_OTKUPA) = cenaOtk
        res(n, PL_VEC_DOPLACENO) = vecKg
        res(n, PL_DATUM_BLOKA) = v(7)
        If dpk > 0 Then
            res(n, PL_CENA) = dpk
            res(n, PL_IZNOS) = modNovac.ZaokruziNovac(kg * dpk)
            res(n, PL_STATUS) = DOP_PLAN_KNJIZI
        Else
            res(n, PL_CENA) = 0#
            res(n, PL_IZNOS) = 0#
            res(n, PL_STATUS) = DOP_PLAN_BEZ_DOPLATE
        End If
    Next k

    DoplataPlan = SortArrayPoKljucu(res, PL_KLJUC)
End Function

' PONAVLJANJE (DOP-ODL-06, revidirano po bloku): blokovi iz obracuna koji VEC
' imaju aktivnu doplatu za istu klasu. "" kad ih nema; inace tekst za operatera
' koji imenuje ranije doplate. Vazi u oba nacina unosa -- i u nacinu 2, gde plan
' ranije doplaceno vec oduzima, operater treba da zna da se blok doplacuje opet.
' izabrani: kao u KnjiziDoplatu_TX (Nothing = svi blokovi plana sa doplatom).
' kooperantID: obuhvat obracuna, kao u DoplataPlan (DOP-ODL-31).
Public Function DoplataPreklapanja(ByVal periodOd As Date, ByVal periodDo As Date, _
                                   ByVal stope As Collection, _
                                   Optional ByVal nacin As Long = DOP_NACIN_PO_KG, _
                                   Optional ByVal izabrani As Object = Nothing, _
                                   Optional ByVal kooperantID As String = "") As String
    Const SRC As String = "modDoplata.DoplataPreklapanja"
    Dim plan As Variant, bez As String
    plan = DoplataPlan(periodOd, periodDo, stope, bez, nacin, kooperantID)
    If Not IsArray(plan) Then Exit Function
    DoplataPreklapanja = PonavljanjaUPlanu(PlanZaKnjizenje(plan, izabrani, SRC, False))
End Function

' Isti test nad VEC IZRACUNATIM planom -- pisac ga zove nad planom koji ce
' knjiziti, pa pitanje i upis ne mogu da se raziidju.
Private Function PonavljanjaUPlanu(ByVal plan As Variant) As String
    If Not IsArray(plan) Then Exit Function

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim postojece As Object, i As Long, k As String
    Set postojece = CreateObject("Scripting.Dictionary")
    postojece.CompareMode = vbTextCompare
    For i = 1 To UBound(r, 1)
        k = CStr(r(i, 16)) & "|" & UCase$(Trim$(CStr(r(i, 9))))
        If Not postojece.Exists(k) Then postojece.Add k, CStr(r(i, 2))
    Next i

    Dim brojevi As Object
    Set brojevi = CreateObject("Scripting.Dictionary")
    brojevi.CompareMode = vbTextCompare
    For i = 1 To UBound(plan, 1)
        k = CStr(plan(i, PL_OTKUP)) & "|" & UCase$(Trim$(CStr(plan(i, PL_KLASA))))
        If postojece.Exists(k) Then brojevi(CStr(postojece(k))) = True
    Next i
    If brojevi.count = 0 Then Exit Function

    Dim t As String, kk As Variant, n As Long
    For Each kk In brojevi.keys
        n = n + 1
        If n <= 5 Then
            If Len(t) > 0 Then t = t & ", "
            t = t & CStr(kk)
        End If
    Next kk
    If brojevi.count > 5 Then t = t & " (+" & CStr(brojevi.count - 5) & ")"

    ' Tekst je iz kataloga poruka (dijakritika); ide i u gresku pisca i u
    ' pitanje ekrana (modScrDoplata.KnjiziKlik), pa je isti na oba mesta.
    PonavljanjaUPlanu = Poruka("DOP_ASK_PONAVLJANJE") & " " & t & ". " & _
                        Poruka("DOP_ASK_PONAVLJANJE_POTVRDI")
End Function

' Redovi plana koji se KNJIZE: status "knjizi" i, kad je izbor zadat, samo
' izabrani blokovi (DOP-ODL-16). izabrani je Scripting.Dictionary sa OtkupID-em
' kao kljucem; Nothing = svi blokovi plana sa doplatom. Pisac knjizi TACNO ono
' sto je operater video: izabran blok koji vise nije u planu sa doplatom
' (storniran, van perioda, van kljuca, bez doplate) obara knjizenje po imenu
' (strogo = True), a blok koji se pojavio posle pregleda se ne knjizi tiho.
' Empty kad ne ostane nijedan red.
Private Function PlanZaKnjizenje(ByVal plan As Variant, ByVal izabrani As Object, _
                                 ByVal sourceName As String, _
                                 Optional ByVal strogo As Boolean = True) As Variant
    PlanZaKnjizenje = Empty
    If Not IsArray(plan) Then Exit Function

    Dim izbor As Object, k As Variant
    If Not izabrani Is Nothing Then
        Set izbor = CreateObject("Scripting.Dictionary")
        izbor.CompareMode = vbTextCompare
        For Each k In izabrani.keys
            If Len(Trim$(CStr(k))) > 0 Then izbor(Trim$(CStr(k))) = True
        Next k
    End If

    ' Jedna odluka po redu, pa se brojanje i prepisivanje ne mogu raziici.
    Dim sDoplatom As Object, i As Long, n As Long, oid As String, uzmi() As Boolean
    Set sDoplatom = CreateObject("Scripting.Dictionary")
    sDoplatom.CompareMode = vbTextCompare
    ReDim uzmi(1 To UBound(plan, 1))
    For i = 1 To UBound(plan, 1)
        oid = CStr(plan(i, PL_OTKUP))
        If CLng(plan(i, PL_STATUS)) = DOP_PLAN_KNJIZI Then
            sDoplatom(oid) = True
            uzmi(i) = True
            If Not izbor Is Nothing Then uzmi(i) = izbor.Exists(oid)
        End If
        If uzmi(i) Then n = n + 1
    Next i

    If strogo And Not izbor Is Nothing Then
        For Each k In izbor.keys
            If Not sDoplatom.Exists(CStr(k)) Then
                Err.Raise DOP_ERR_BASE + 33, sourceName, _
                          "Izabran blok vise ne ulazi u obracun sa doplatom (storniran, van " & _
                          "perioda, van kljuca ili bez doplate): " & CStr(k) & ". Osvezi pregled."
            End If
        Next k
    End If
    If n = 0 Then Exit Function

    Dim res() As Variant, j As Long, m As Long
    ReDim res(1 To n, 1 To PL_KOLONA)
    For i = 1 To UBound(plan, 1)
        If uzmi(i) Then
            m = m + 1
            For j = 1 To PL_KOLONA
                res(m, j) = plan(i, j)
            Next j
        End If
    Next i
    PlanZaKnjizenje = res
End Function

' Vec knjizena AKTIVNA doplata po kg: "OtkupID|KLASA" -> zbir bruto doplate po
' kg svih aktivnih stavki tog bloka i klase. Nacin 2 je oduzima od konacne cene.
Private Function VecDoplacenoPoKg() As Object
    Dim d As Object, r As Variant, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set VecDoplacenoPoKg = d
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function
    For i = 1 To UBound(r, 1)
        k = UCase$(CStr(r(i, 16))) & "|" & UCase$(Trim$(CStr(r(i, 9))))
        d(k) = CDbl(d(k)) + CDbl(r(i, 11))
    Next i
End Function

' Problem sa stopama kao TEKST za operatera, "" kad su ispravne. Isto pravilo
' koje pisac dize kao gresku (UcitajStope) -- jedna implementacija, dva
' pozivaoca: ekran Doplata ga zove pri dodavanju stope, da operater dobije
' razlog uz polje umesto tek pri knjizenju.
Public Function DoplataStopeProblem(ByVal stope As Collection) As String
    Dim cene As Object, vrsteSorte As Object, prikaz As Object
    On Error GoTo EH
    UcitajStope stope, cene, vrsteSorte, prikaz, "modDoplata.DoplataStopeProblem"
    Exit Function
EH:
    DoplataStopeProblem = Err.description
End Function

' ============================================================
' KNJIZENJE -- jedini ulaz upisa
' ============================================================

' Knjizi obracun: JEDAN DOKUMENT DOPLATE PO BLOKU iz plana (DOP-ODL-16), svi pod
' jednim ObracunID-em, u jednoj transakciji. Vraca ObracunID ili "" (outGreska
' kaze zasto).
'
' datum         datum knjizenja = datum zaduzenja na kartici (DOP-ODL-04)
' periodOd/Do   period otkupa, sa satom (DOP-ODL-03a)
' stope         Collection stopa (v. UcitajStope); vrednost stope znaci ono sto
'               kaze nacin
' potvrdaPonavljanja  odgovor operatera na DoplataPreklapanja (DOP-ODL-06)
' nacin         DOP_NACIN_PO_KG ili DOP_NACIN_KONACNA (DOP-ODL-15)
' izabrani      Scripting.Dictionary OtkupID -> True: blokovi koje operater
'               knjizi; Nothing = svi blokovi plana sa doplatom (v. PlanZaKnjizenje)
' kooperantID   obuhvat obracuna (DOP-ODL-31): prazno = svi kooperanti; zadat =
'               obracun samo za tog kooperanta. Plan se racuna u TOM obuhvatu, pa
'               izabran blok drugog kooperanta obara knjizenje po imenu (isti
'               put kao blok koji vise nije u obracunu), a obuhvat se upisuje na
'               svaki dokument (COL_DOP_OBUHVAT, DOP-12).
'
' Kapije pre ijednog upisa (DOP-ODL-10): period je zatvoren i vec prosao, a
' datum knjizenja nije pre kraja perioda ni posle danas -- zaduzenje ne sme da
' stoji na kartici pre robe koju placa. Svaka stopa mora da pogodi bar jedan
' otkup. Nema praznog obracuna.
Public Function KnjiziDoplatu_TX(ByVal datum As Date, _
                                 ByVal periodOd As Date, _
                                 ByVal periodDo As Date, _
                                 ByVal stope As Collection, _
                                 Optional ByVal napomena As String = "", _
                                 Optional ByVal potvrdaPonavljanja As Boolean = False, _
                                 Optional ByRef outGreska As String, _
                                 Optional ByRef outErrNum As Long, _
                                 Optional ByVal nacin As Long = DOP_NACIN_PO_KG, _
                                 Optional ByVal izabrani As Object = Nothing, _
                                 Optional ByVal kooperantID As String = "") As String
    Const SRC As String = "modDoplata.KnjiziDoplatu_TX"

    Dim tx As clsTransaction
    Set tx = New clsTransaction

    outGreska = ""
    outErrNum = 0

    On Error GoTo EH

    ' Sema pre upisa: AppendRow pise POZICIONO. Ide PRE BeginTx, kao kod otkupa.
    modSchema.SchemaReadyOrFail SRC, TBL_DOPLATA & "|" & TBL_DOPLATA_STAVKE

    RequireDatumKnjizenja datum, periodDo, SRC

    ' Obuhvat je identitet, ne tekst: kooperant mora da postoji TACNO jednom.
    ' Nepoznat ID bi inace dao prazan plan i poruku "nema otkupa" umesto razloga.
    Dim obuhvat As String
    obuhvat = Trim$(kooperantID)
    If Len(obuhvat) > 0 Then
        RequireTacnoJedan TBL_KOOPERANTI, COL_KOOP_ID, obuhvat, "KooperantID (obuhvat obracuna)", SRC
    End If

    Dim plan As Variant, bezOtkupa As String
    plan = DoplataPlan(periodOd, periodDo, stope, bezOtkupa, nacin, obuhvat)

    If Len(bezOtkupa) > 0 Then
        Err.Raise DOP_ERR_BASE + 30, SRC, _
                  "Stope bez ijednog aktivnog otkupa u periodu: " & bezOtkupa & _
                  ". Ukloni ih ili ispravi kljuc (kljuc cenovnika: vrsta, sorta, klasa)."
    End If
    If Not IsArray(plan) Then
        Err.Raise DOP_ERR_BASE + 31, SRC, _
                  "U periodu nema nijednog aktivnog otkupa za zadate stope."
    End If
    plan = PlanZaKnjizenje(plan, izabrani, SRC)
    If Not IsArray(plan) Then
        Err.Raise DOP_ERR_BASE + 31, SRC, _
                  "Nema nijednog bloka za knjizenje: svi su iskljuceni ili su bez doplate."
    End If

    Dim preklop As String
    preklop = PonavljanjaUPlanu(plan)
    If Len(preklop) > 0 And Not potvrdaPonavljanja Then
        Err.Raise DOP_ERR_POTVRDA_PONAVLJANJA, SRC, preklop
    End If

    tx.BeginTx
    tx.AddTableSnapshot TBL_DOPLATA
    tx.AddTableSnapshot TBL_DOPLATA_STAVKE

    KnjiziDoplatu_TX = KnjiziPlan(datum, periodOd, periodDo, plan, napomena, nacin, obuhvat, SRC)

    tx.CommitTx
    Set tx = Nothing

    On Error Resume Next
    Monitor_Event _
        eventType:="DOPLATA_KNJIZENA", _
        severity:="INFO", _
        message:="Doplata knjizena. ObracunID=" & KnjiziDoplatu_TX, _
        userId:="Operator", _
        moduleName:="modDoplata", _
        procedureName:=SRC, _
        entityType:="Doplata", _
        entityID:=KnjiziDoplatu_TX, _
        correlationId:=KnjiziDoplatu_TX
    On Error GoTo 0
    Exit Function

EH:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.description
    errSrc = Err.SOURCE

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx

    outErrNum = errNum
    outGreska = errDesc
    KnjiziDoplatu_TX = ""

    ' Pitanje nije kvar: bez loga i monitora, da log ostane signal.
    If errNum <> DOP_ERR_POTVRDA_PONAVLJANJA Then
        LogError SRC, errDesc, errNum
        Monitor_Error _
            moduleName:="modDoplata", _
            procedureName:=SRC, _
            entityType:="Doplata", _
            entityID:="", _
            correlationId:="", _
            errorNumber:=errNum, _
            errorDescription:=errDesc, _
            errorSource:=errSrc
    End If
    On Error GoTo 0
End Function

' Core -- NE zovi spolja. Jedini ulaz je KnjiziDoplatu_TX, koji drzi snapshot
' obe tabele. Plan je vec filtriran (PlanZaKnjizenje) i poredan, pa su stavke
' istog bloka uzastopne: jedan prolaz pravi jedan dokument po bloku.
' Broj = broj bloka + "-0" + redni broj doplate na tom bloku (DOP-ODL-30); redni
' broj se racuna preko SVIH ranijih dokumenata bloka, i storniranih.
' obuhvat: "" ili KooperantID za koga je obracun radjen (DOP-ODL-31); ide na
' svako zaglavlje, a dokument drugog kooperanta pod tim obuhvatom je greska.
Private Function KnjiziPlan(ByVal datum As Date, _
                            ByVal periodOd As Date, _
                            ByVal periodDo As Date, _
                            ByVal plan As Variant, _
                            ByVal napomena As String, _
                            ByVal nacin As Long, _
                            ByVal obuhvat As String, _
                            ByVal sourceName As String) As String
    Dim obracunID As String
    obracunID = NewEntityID("DOB-")
    If Len(obracunID) = 0 Then
        Err.Raise DOP_ERR_BASE + 40, sourceName, "NewEntityID nije vratio ObracunID."
    End If

    Dim poBloku As Object, brojevi As Object
    UcitajBrojeveDoplata poBloku, brojevi, sourceName

    ' Specifikacija (DOP-ODL-21/30): jedan broj po kooperantu u ovom obracunu,
    ' x/ddmmyy-00y -- maticna stanica kooperanta, datum knjizenja, redni broj te
    ' stanice tog dana; storno ga ne oslobadja. specKoop: KooperantID -> broj.
    Dim specMax As Object, specBrojevi As Object, specKoop As Object, spec As String
    Dim staKoop As String, kljucSpec As String, rSpec As Long
    UcitajSpecifikacije specMax, specBrojevi, sourceName
    Set specKoop = CreateObject("Scripting.Dictionary")
    specKoop.CompareMode = vbTextCompare

    Dim proveren As Object
    Set proveren = CreateObject("Scripting.Dictionary")
    proveren.CompareMode = vbTextCompare

    Dim i As Long, tekuci As String, doplataID As String, rb As Long, koop As String
    Dim rowData() As Variant, stavkaID As String, broj As String, brojBloka As String

    For i = 1 To UBound(plan, 1)
        If StrComp(CStr(plan(i, PL_OTKUP)), tekuci, vbTextCompare) <> 0 Then
            tekuci = CStr(plan(i, PL_OTKUP))
            koop = CStr(plan(i, PL_KOOP))
            ' DOP-12, kapija PISCA: plan je vec suzen na obuhvat (DoplataPlan),
            ' ali upis ne veruje pozivaocu -- dokument tudjeg kooperanta pod
            ' ovim obuhvatom bi signalu osnovice sakrio njegove blokove.
            If Not UObuhvatu(koop, obuhvat) Then
                Err.Raise DOP_ERR_BASE + 36, sourceName, _
                          "Blok " & tekuci & " pripada kooperantu " & koop & _
                          ", a obracun je samo za kooperanta " & obuhvat & "."
            End If
            If Not proveren.Exists(koop) Then
                RequireTacnoJedan TBL_KOOPERANTI, COL_KOOP_ID, koop, "KooperantID", sourceName
                proveren.Add koop, True
            End If
            If Not specKoop.Exists(koop) Then
                staKoop = Trim$(NzToText(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, koop, COL_KOOP_STANICA)))
                If modBrojevi.ExtractNumericFromEntityID(staKoop) = 0 Then
                    Err.Raise DOP_ERR_BASE + 58, sourceName, _
                              "Kooperant " & koop & " nema maticnu stanicu sa brojem, pa " & _
                              "specifikacija nema broj (x/ddmmyy-00y)."
                End If
                kljucSpec = CStr(modBrojevi.ExtractNumericFromEntityID(staKoop)) & "/" & _
                            Format$(datum, "ddmmyy")
                rSpec = CLng(specMax(kljucSpec)) + 1
                spec = BrojSpecifikacije(staKoop, datum, rSpec)
                If specBrojevi.Exists(spec) Then
                    Err.Raise DOP_ERR_BASE + 48, sourceName, _
                              "Broj specifikacije " & spec & " je vec zauzet."
                End If
                specBrojevi.Add spec, True
                specMax(kljucSpec) = rSpec
                specKoop.Add koop, spec
            End If

            brojBloka = Trim$(CStr(plan(i, PL_BROJ_BLOKA)))
            If Len(brojBloka) = 0 Then
                Err.Raise DOP_ERR_BASE + 46, sourceName, _
                          "Blok " & tekuci & " nema broj dokumenta, pa doplata nema broj."
            End If
            broj = BrojDoplate(brojBloka, CLng(poBloku(tekuci)) + 1)
            If brojevi.Exists(broj) Then
                Err.Raise DOP_ERR_BASE + 47, sourceName, _
                          "Broj doplate " & broj & " je vec zauzet drugim dokumentom (blok sa " & _
                          "istim brojem na drugoj stanici?). Iskljuci blok " & brojBloka & " iz obracuna."
            End If
            brojevi.Add broj, True
            poBloku(tekuci) = CLng(poBloku(tekuci)) + 1

            doplataID = NewEntityID("DOP-")
            If Len(doplataID) = 0 Then
                Err.Raise DOP_ERR_BASE + 41, sourceName, "NewEntityID nije vratio DoplataID."
            End If

            rowData = PrazanRed(TBL_DOPLATA, sourceName)
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_ID, doplataID, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_BROJ, broj, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_DATUM, datum, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_KOOPERANT, koop, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_OBRACUN, obracunID, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_PERIOD_OD, periodOd, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_PERIOD_DO, periodDo, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_NAPOMENA, Trim$(napomena), sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_STORNIRANO, "", sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_OTKUP, tekuci, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_STANICA, plan(i, PL_STANICA), sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_NACIN, nacin, sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_SPEC, specKoop(koop), sourceName
            SetRowValueByColumn rowData, TBL_DOPLATA, COL_DOP_OBUHVAT, obuhvat, sourceName

            If AppendRow(TBL_DOPLATA, rowData) <= 0 Then
                Err.Raise DOP_ERR_BASE + 42, sourceName, _
                          "AppendRow nije upisao zaglavlje doplate za blok " & brojBloka & "."
            End If
            rb = 0
        End If

        rb = rb + 1
        stavkaID = NewEntityID("DPS-")
        If Len(stavkaID) = 0 Then
            Err.Raise DOP_ERR_BASE + 43, sourceName, "NewEntityID nije vratio DoplataStavkaID."
        End If

        rowData = PrazanRed(TBL_DOPLATA_STAVKE, sourceName)
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_ID, stavkaID, sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_DOPLATA_ID, doplataID, sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_RB, rb, sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_STANICA, plan(i, PL_STANICA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_VRSTA, plan(i, PL_VRSTA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_SORTA, plan(i, PL_SORTA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_KLASA, plan(i, PL_KLASA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_KOLICINA, plan(i, PL_KOLICINA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_CENA, plan(i, PL_CENA), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_IZNOS, plan(i, PL_IZNOS), sourceName
        SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_CENA_OTKUPA, plan(i, PL_CENA_OTKUPA), sourceName
        If nacin = DOP_NACIN_KONACNA Then
            SetRowValueByColumn rowData, TBL_DOPLATA_STAVKE, COL_DOPS_KONACNA, plan(i, PL_KONACNA), sourceName
        End If

        If AppendRow(TBL_DOPLATA_STAVKE, rowData) <= 0 Then
            Err.Raise DOP_ERR_BASE + 44, sourceName, _
                      "AppendRow nije upisao stavku doplate za blok " & brojBloka & "."
        End If
    Next i

    KnjiziPlan = obracunID
End Function

' ============================================================
' CITACI
' ============================================================

' Stavke AKTIVNIH dokumenata doplate, sa zaglavljem. 2D niz (1..n, 1..21):
'   1 DoplataID | 2 BrojDokumenta | 3 Datum | 4 KooperantID | 5 ObracunID
'   6 StanicaID | 7 VrstaVoca | 8 SortaVoca | 9 Klasa | 10 Kolicina
'   11 Cena (bruto doplata po kg) | 12 Iznos | 13 DoplataStavkaID
'   14 PeriodOd | 15 PeriodDo | 16 OtkupID (blok) | 17 CenaOtkupa
'   18 KonacnaCena (nacin 2; inace prazno) | 19 NacinUnosa
'   20 BrojSpecifikacije (prazno za dokument knjizen pre P3)
'   21 ObuhvatKooperantID (prazno = obracun za sve kooperante; DOP-ODL-31)
' redosledom tabele stavki. Empty kad nema nijedne aktivne stavke.
'
' STROG, kao StavkeOtkupaRedovi: dupli DoplataID, stavka bez zaglavlja i
' aktivno zaglavlje bez stavki padaju PO IMENU -- tiha nula bi sakrila dug firme
' prema kooperantu. Stavka storniranog zaglavlja nije aktivna (DOP-03).
' DOP-12: obuhvat je osobina OBRACUNA -- svi dokumenti obracuna nose isti, a
' zadat obuhvat je kooperant dokumenta. Odstupanje pada po imenu: signal
' osnovice bi inace gledao pogresnog kooperanta.
Public Function DoplataStavkeRedovi() As Variant
    Const SRC As String = "modDoplata.DoplataStavkeRedovi"

    DoplataStavkeRedovi = Empty

    ' Sveska PRE NADOGRADNJE: obe tabele jos ne postoje (pravi ih EnsureAllTables
    ' na startu), pa KO nema -- prazno, isti obrazac kao modUtovar.UtovarenoPoPreradi.
    ' Bez ovoga bi kartica i Saldo OM pali SVIM kooperantima do prvog restarta.
    ' Samo jedna od dve je drift i pada po imenu nize (RequireColumnIndex).
    If GetTable(TBL_DOPLATA) Is Nothing And GetTable(TBL_DOPLATA_STAVKE) Is Nothing Then
        Exit Function
    End If

    Dim hd As Variant
    hd = GetTableData(TBL_DOPLATA)

    Dim sve As Object, aktivna As Object
    Set sve = CreateObject("Scripting.Dictionary")
    sve.CompareMode = vbTextCompare
    Set aktivna = CreateObject("Scripting.Dictionary")
    aktivna.CompareMode = vbTextCompare

    Dim cId As Long, cBr As Long, cDat As Long, cKoop As Long, cObr As Long
    Dim cOd As Long, cDo As Long, cOtk As Long, cNac As Long, cSpec As Long
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cBr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_BROJ, SRC)
    cDat = RequireColumnIndex(TBL_DOPLATA, COL_DOP_DATUM, SRC)
    cKoop = RequireColumnIndex(TBL_DOPLATA, COL_DOP_KOOPERANT, SRC)
    cObr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_OBRACUN, SRC)
    cOd = RequireColumnIndex(TBL_DOPLATA, COL_DOP_PERIOD_OD, SRC)
    cDo = RequireColumnIndex(TBL_DOPLATA, COL_DOP_PERIOD_DO, SRC)
    cOtk = RequireColumnIndex(TBL_DOPLATA, COL_DOP_OTKUP, SRC)
    cNac = RequireColumnIndex(TBL_DOPLATA, COL_DOP_NACIN, SRC)
    cSpec = RequireColumnIndex(TBL_DOPLATA, COL_DOP_SPEC, SRC)
    Dim cObuh As Long
    cObuh = RequireColumnIndex(TBL_DOPLATA, COL_DOP_OBUHVAT, SRC)

    ' Specifikacija je par (ObracunID, KooperantID); broj je njena oznaka i mora
    ' biti jedan po paru i ne sme da pripada dvama parovima -- i preko storniranih
    ' (storno ne oslobadja broj). Prazan broj je dokument knjizen pre P3.
    Dim parSpec As Object, specPar As Object, par As String, spec As String
    Set parSpec = CreateObject("Scripting.Dictionary")
    parSpec.CompareMode = vbTextCompare
    Set specPar = CreateObject("Scripting.Dictionary")
    specPar.CompareMode = vbTextCompare

    ' obrObuh: ObracunID -> obuhvat (DOP-12), i preko storniranih dokumenata.
    Dim obrObuh As Object, obuh As String, obrID As String
    Set obrObuh = CreateObject("Scripting.Dictionary")
    obrObuh.CompareMode = vbTextCompare

    Dim i As Long, id As String
    If IsArray(hd) Then
        For i = 1 To UBound(hd, 1)
            id = Trim$(NzToText(hd(i, cId)))
            If Len(id) = 0 Then
                Err.Raise DOP_ERR_BASE + 50, SRC, _
                          "Zaglavlje KO bez DoplataID-a (red " & CStr(i) & ")."
            End If
            If sve.Exists(id) Then
                Err.Raise DOP_ERR_BASE + 51, SRC, "DoplataID nije jedinstven: " & id & "."
            End If
            sve.Add id, True

            par = Trim$(NzToText(hd(i, cObr))) & "|" & Trim$(NzToText(hd(i, cKoop)))
            spec = Trim$(NzToText(hd(i, cSpec)))
            If parSpec.Exists(par) Then
                If StrComp(CStr(parSpec(par)), spec, vbTextCompare) <> 0 Then
                    Err.Raise DOP_ERR_BASE + 56, SRC, _
                              "Dokumenti iste specifikacije nose razlicit broj (" & _
                              CStr(parSpec(par)) & " / " & spec & "), dokument " & id & "."
                End If
            Else
                parSpec.Add par, spec
            End If
            If Len(spec) > 0 Then
                If specPar.Exists(spec) Then
                    If StrComp(CStr(specPar(spec)), par, vbTextCompare) <> 0 Then
                        Err.Raise DOP_ERR_BASE + 57, SRC, _
                                  "Broj specifikacije " & spec & " pripada dvema specifikacijama."
                    End If
                Else
                    specPar.Add spec, par
                End If
            End If

            obuh = Trim$(NzToText(hd(i, cObuh)))
            obrID = Trim$(NzToText(hd(i, cObr)))
            If Len(obuh) > 0 Then
                If StrComp(obuh, Trim$(NzToText(hd(i, cKoop))), vbTextCompare) <> 0 Then
                    Err.Raise DOP_ERR_BASE + 59, SRC, _
                              "Dokument doplate " & id & " je kooperanta " & _
                              Trim$(NzToText(hd(i, cKoop))) & ", a obracun je samo za " & _
                              "kooperanta " & obuh & "."
                End If
            End If
            If obrObuh.Exists(obrID) Then
                If StrComp(CStr(obrObuh(obrID)), obuh, vbTextCompare) <> 0 Then
                    Err.Raise DOP_ERR_BASE + 60, SRC, _
                              "Dokumenti istog obracuna nose razlicit obuhvat ('" & _
                              CStr(obrObuh(obrID)) & "' / '" & obuh & "'), dokument " & id & "."
                End If
            Else
                obrObuh.Add obrID, obuh
            End If
        Next i

        Dim akt As Variant
        akt = ExcludeStornirano(hd, TBL_DOPLATA)
        If IsArray(akt) Then
            For i = 1 To UBound(akt, 1)
                aktivna.Add Trim$(NzToText(akt(i, cId))), _
                    Array(NzToText(akt(i, cBr)), akt(i, cDat), Trim$(NzToText(akt(i, cKoop))), _
                          Trim$(NzToText(akt(i, cObr))), akt(i, cOd), akt(i, cDo), _
                          Trim$(NzToText(akt(i, cOtk))), akt(i, cNac), _
                          Trim$(NzToText(akt(i, cSpec))), Trim$(NzToText(akt(i, cObuh))))
            Next i
        End If
    End If

    Dim sd As Variant
    sd = GetTableData(TBL_DOPLATA_STAVKE)

    Dim sId As Long, sDop As Long, sSta As Long, sVr As Long, sSo As Long
    Dim sKl As Long, sKol As Long, sCena As Long, sIzn As Long, sCOtk As Long, sKon As Long
    sId = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_ID, SRC)
    sDop = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_DOPLATA_ID, SRC)
    sSta = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_STANICA, SRC)
    sVr = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_VRSTA, SRC)
    sSo = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_SORTA, SRC)
    sKl = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_KLASA, SRC)
    sKol = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_KOLICINA, SRC)
    sCena = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_CENA, SRC)
    sIzn = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_IZNOS, SRC)
    sCOtk = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_CENA_OTKUPA, SRC)
    sKon = RequireColumnIndex(TBL_DOPLATA_STAVKE, COL_DOPS_KONACNA, SRC)

    Dim imaStavku As Object
    Set imaStavku = CreateObject("Scripting.Dictionary")
    imaStavku.CompareMode = vbTextCompare

    Dim n As Long
    If IsArray(sd) Then
        For i = 1 To UBound(sd, 1)
            id = Trim$(NzToText(sd(i, sDop)))
            If Not sve.Exists(id) Then
                Err.Raise DOP_ERR_BASE + 52, SRC, _
                          "Stavka KO pokazuje na zaglavlje koje ne postoji: DoplataID=" & id & _
                          ", stavka " & NzToText(sd(i, sId)) & "."
            End If
            If aktivna.Exists(id) Then
                If Not IsNumeric(sd(i, sKol)) Or Not IsNumeric(sd(i, sCena)) _
                   Or Not IsNumeric(sd(i, sIzn)) Then
                    Err.Raise DOP_ERR_BASE + 53, SRC, _
                              "Stavka KO nema brojcanu kolicinu, cenu ili iznos: " & _
                              NzToText(sd(i, sId)) & "."
                End If
                ' Iznos je ono sto se isplacuje: mora biti TACNO kg x cena u
                ' cent-domenu, kako ga pisac racuna. Rucno prepravljen iznos bi
                ' tiho promenio dug firme prema kooperantu (review P1, REV-03).
                If modNovac.ZaokruziNovac(CDbl(sd(i, sKol)) * CDbl(sd(i, sCena))) <> _
                   modNovac.ZaokruziNovac(CDbl(sd(i, sIzn))) Then
                    Err.Raise DOP_ERR_BASE + 55, SRC, _
                              "Stavka doplate nije dosledna (kolicina x cena <> iznos): " & _
                              NzToText(sd(i, sId)) & "."
                End If
                imaStavku(id) = True
                n = n + 1
            End If
        Next i
    End If

    Dim k As Variant, h As Variant
    For Each k In aktivna.keys
        If Not imaStavku.Exists(CStr(k)) Then
            h = aktivna(k)
            Err.Raise DOP_ERR_BASE + 54, SRC, _
                      "Aktivan KO nema nijednu stavku: " & CStr(h(0)) & _
                      " (" & CStr(k) & ")."
        End If
    Next k

    If n = 0 Then Exit Function

    Dim res() As Variant
    ReDim res(1 To n, 1 To 21)
    n = 0
    For i = 1 To UBound(sd, 1)
        id = Trim$(NzToText(sd(i, sDop)))
        If aktivna.Exists(id) Then
            h = aktivna(id)
            n = n + 1
            res(n, 1) = id
            res(n, 2) = h(0)
            res(n, 3) = h(1)
            res(n, 4) = h(2)
            res(n, 5) = h(3)
            res(n, 6) = Trim$(NzToText(sd(i, sSta)))
            res(n, 7) = Trim$(NzToText(sd(i, sVr)))
            res(n, 8) = Trim$(NzToText(sd(i, sSo)))
            res(n, 9) = Trim$(NzToText(sd(i, sKl)))
            res(n, 10) = CDbl(sd(i, sKol))
            res(n, 11) = CDbl(sd(i, sCena))
            res(n, 12) = CDbl(sd(i, sIzn))
            res(n, 13) = Trim$(NzToText(sd(i, sId)))
            res(n, 14) = h(4)
            res(n, 15) = h(5)
            res(n, 16) = h(6)
            res(n, 17) = 0#
            If IsNumeric(sd(i, sCOtk)) Then res(n, 17) = CDbl(sd(i, sCOtk))
            res(n, 18) = sd(i, sKon)
            res(n, 19) = h(7)
            res(n, 20) = h(8)
            res(n, 21) = h(9)
        End If
    Next i

    DoplataStavkeRedovi = res
End Function

' Iznos AKTIVNIH KO po kooperantu: KooperantID -> din (zbir stavki).
'
' stanicaID  samo stavke te stanice -- stavka pripada stanici na kojoj je roba
'            predata (DOPLATA.md par. 3), pa Saldo OM svake stanice vidi svoj
'            deo, a zbir po stanicama je ceo KO. Prazno = sve stanice.
' datumOd/Do Datum KO = datum knjizenja = datum zaduzenja (DOP-ODL-04), po
'            danu, obe granice ukljucene. Izostavljeno = bez granice.
'
' Jedan citac za Saldo OM i PWA izvoz salda -- isti filter storna i ista
' pripadnost stanici, bez druge kopije.
Public Function DoplataIznosPoKooperantu(Optional ByVal stanicaID As String = "", _
                                         Optional ByVal datumOd As Variant, _
                                         Optional ByVal datumDo As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set DoplataIznosPoKooperantu = dict

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim imaOd As Boolean, imaDo As Boolean, danOd As Long, danDo As Long
    If Not IsMissing(datumOd) Then
        If IsDate(datumOd) Then imaOd = True: danOd = Int(CDbl(CDate(datumOd)))
    End If
    If Not IsMissing(datumDo) Then
        If IsDate(datumDo) Then imaDo = True: danDo = Int(CDbl(CDate(datumDo)))
    End If

    Dim i As Long, dan As Long, koop As String, ok As Boolean
    For i = 1 To UBound(r, 1)
        ok = True
        If Len(Trim$(stanicaID)) > 0 Then
            ok = (StrComp(CStr(r(i, 6)), Trim$(stanicaID), vbTextCompare) = 0)
        End If
        If ok And (imaOd Or imaDo) Then
            dan = Int(CDbl(CDate(r(i, 3))))
            If imaOd And dan < danOd Then ok = False
            If imaDo And dan > danDo Then ok = False
        End If
        If ok Then
            koop = CStr(r(i, 4))
            dict(koop) = CDbl(dict(koop)) + CDbl(r(i, 12))
        End If
    Next i
End Function

' AKTIVNI dokumenti doplate jednog kooperanta, jedan red po dokumentu -- za
' listu cilja rucnog mapiranja. (Kartica od P4 cita DoplataObracuniKooperanta --
' red po obracunu, ne po dokumentu.)
' 2D niz (1..n, 1..6):
'   1 DoplataID | 2 BrojDokumenta | 3 Datum | 4 Iznos (zbir stavki)
'   5 delovi opisa: Collection od Array(vrsta, sorta, klasa, kg, cena)
'   6 StanicaID bloka
' Stavke istog kljuca i iste stope sa razlicitih stanica se sabiraju: kartica
' je kooperantova, ne stanicina. Redosled: redosled prve stavke u tabeli; poredak
' na kartici ionako daje kljuc hronologije (KART-HRON-01). Empty kad nema KO.
Public Function DoplataKOKooperanta(ByVal kooperantID As String) As Variant
    DoplataKOKooperanta = Empty

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim ko As Object, red As Variant, delovi As Collection, kDeo As Object
    Dim i As Long, id As String, kljuc As String, j As Long, deo As Variant
    Set ko = CreateObject("Scripting.Dictionary")
    ko.CompareMode = vbTextCompare
    Set kDeo = CreateObject("Scripting.Dictionary")
    kDeo.CompareMode = vbTextCompare

    For i = 1 To UBound(r, 1)
        If StrComp(CStr(r(i, 4)), Trim$(kooperantID), vbTextCompare) = 0 Then
            id = CStr(r(i, 1))
            If Not ko.Exists(id) Then
                Set delovi = New Collection
                ko.Add id, Array(r(i, 2), r(i, 3), 0#, delovi, CStr(r(i, 6)))
            End If
            red = ko(id)
            red(2) = CDbl(red(2)) + CDbl(r(i, 12))
            Set delovi = red(3)
            kljuc = id & "|" & UCase$(CStr(r(i, 7))) & "|" & UCase$(CStr(r(i, 8))) & "|" & _
                    UCase$(CStr(r(i, 9))) & "|" & CStr(r(i, 11))
            If kDeo.Exists(kljuc) Then
                j = CLng(kDeo(kljuc))
                deo = delovi(j)
                deo(3) = CDbl(deo(3)) + CDbl(r(i, 10))
                delovi.Remove j
                If j > delovi.count Then
                    delovi.Add deo
                Else
                    delovi.Add deo, Before:=j
                End If
            Else
                delovi.Add Array(CStr(r(i, 7)), CStr(r(i, 8)), CStr(r(i, 9)), CDbl(r(i, 10)), _
                                 CDbl(r(i, 11)))
                kDeo.Add kljuc, delovi.count
            End If
            ko(id) = red
        End If
    Next i

    If ko.count = 0 Then Exit Function

    Dim res() As Variant, n As Long, k As Variant
    ReDim res(1 To ko.count, 1 To 6)
    For Each k In ko.keys
        red = ko(k)
        n = n + 1
        res(n, 1) = CStr(k)
        res(n, 2) = red(0)
        res(n, 3) = red(1)
        res(n, 4) = modNovac.ZaokruziNovac(CDbl(red(2)))
        Set res(n, 5) = red(3)
        res(n, 6) = red(4)
    Next k
    DoplataKOKooperanta = res
End Function

' KARTICA ZBIRNO PO OBRACUNU (DOP-ODL-04/18, P4): jedan red po obracunu doplate
' za kooperanta -- to je njegova specifikacija (par obracun x kooperant), ma
' koliko blokova i stanica nosila. 2D niz (1..n, 1..5):
'   1 ObracunID | 2 Broj: broj specifikacije; obracun knjizen pre P3 (bez broja)
'     nosi broj svog prvog dokumenta | 3 Datum (knjizenja) | 4 Iznos (zbir
'     aktivnih dokumenata)
'   5 delovi opisa PO KULTURI I KLASI: Collection od Array(vrsta, sorta, klasa,
'     kg, din/kg). Kg i iznos se sabiraju preko blokova i stanica; din/kg je
'     iznos / kg -- u nacinu "konacna cena" stopa se razlikuje po bloku, pa je to
'     prosek, isto kao rekapitulacija specifikacije.
' Redosled: redosled prve stavke obracuna u tabeli (poredak na kartici ionako
' daje kljuc hronologije, KART-HRON-01). Storniran dokument ne ulazi (strog
' citac), pa obracun ciji su svi dokumenti stornirani nema red. Empty kad nema.
Public Function DoplataObracuniKooperanta(ByVal kooperantID As String) As Variant
    DoplataObracuniKooperanta = Empty

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim ob As Object, red As Variant, pd As Object, a As Variant
    Dim i As Long, id As String, kljuc As String
    Set ob = CreateObject("Scripting.Dictionary")
    ob.CompareMode = vbTextCompare

    For i = 1 To UBound(r, 1)
        If StrComp(CStr(r(i, 4)), Trim$(kooperantID), vbTextCompare) = 0 Then
            id = CStr(r(i, 5))
            If Not ob.Exists(id) Then
                Set pd = CreateObject("Scripting.Dictionary")
                pd.CompareMode = vbTextCompare
                ob.Add id, Array(r(i, 2), r(i, 3), 0#, pd, Trim$(CStr(r(i, 20))))
            End If
            red = ob(id)
            red(2) = CDbl(red(2)) + CDbl(r(i, 12))
            If Len(CStr(red(4))) = 0 Then red(4) = Trim$(CStr(r(i, 20)))
            Set pd = red(3)
            kljuc = UCase$(CStr(r(i, 7))) & "|" & UCase$(CStr(r(i, 8))) & "|" & UCase$(CStr(r(i, 9)))
            If pd.Exists(kljuc) Then
                a = pd(kljuc)
            Else
                a = Array(CStr(r(i, 7)), CStr(r(i, 8)), CStr(r(i, 9)), 0#, 0#)
            End If
            a(3) = CDbl(a(3)) + CDbl(r(i, 10))
            a(4) = CDbl(a(4)) + CDbl(r(i, 12))
            pd(kljuc) = a
            ob(id) = red
        End If
    Next i

    If ob.count = 0 Then Exit Function

    Dim res() As Variant, n As Long, k As Variant, kd As Variant, delovi As Collection
    ReDim res(1 To ob.count, 1 To 5)
    For Each k In ob.keys
        red = ob(k)
        n = n + 1
        res(n, 1) = CStr(k)
        If Len(CStr(red(4))) > 0 Then res(n, 2) = CStr(red(4)) Else res(n, 2) = red(0)
        res(n, 3) = red(1)
        res(n, 4) = modNovac.ZaokruziNovac(CDbl(red(2)))
        Set delovi = New Collection
        Set pd = red(3)
        For Each kd In pd.keys
            a = pd(kd)
            If CDbl(a(3)) <> 0 Then
                delovi.Add Array(a(0), a(1), a(2), CDbl(a(3)), modNovac.ZaokruziNovac(CDbl(a(4)) / CDbl(a(3))))
            Else
                delovi.Add Array(a(0), a(1), a(2), 0#, 0#)
            End If
        Next kd
        Set res(n, 5) = delovi
    Next k
    DoplataObracuniKooperanta = res
End Function

' ============================================================
' ISPLATA KO (F4) -- otvorena obaveza po dokumentu
' ============================================================
'
' KO je obaveza firme prema kooperantu kao i otkupni blok: isplacuje se kroz
' naloge za banku i zatvara uvozom izvoda (DOP-ODL-05). Isplaceno je zbir
' AKTIVNOG novca vezanog DoplataID-em (modNovac.BuildIsplataDictByDoplata) --
' isto pravilo kao za otkup (BuildIsplataDictByOtkup), drugi kljuc.

' OTVORENA KO. Blizanac modNovac.GetOpenOtkupi; cita ga
' modBankaExportPregled.BuildBlokIsplataList. 2D niz (1..n, 1..8):
'   1 DoplataID | 2 BrojDokumenta | 3 Datum | 4 KooperantID
'   5 StanicaID -- stanica BLOKA (DOP-ODL-19): dokument je po bloku, pa i
'     isplata pripada mestu kupovine, kao kod otkupa
'   6 Iznos | 7 Isplaceno | 8 Otvoreno
'   9 BrojSpecifikacije ("" za dokument knjizen pre P3) -- nalozi po njoj
'     biraju red specifikacije umesto dokumenta (DOP-ODL-27)
' U listi je KO ciji je ostatak, u cent-domenu, veci od nule. kooperantID suzava
' na jednog kooperanta. Empty kad nema nijednog.
Public Function OtvoreniKORedovi(Optional ByVal kooperantID As String = "") As Variant
    OtvoreniKORedovi = Empty

    Dim glava As Object, iznosi As Object, isplate As Object
    Set iznosi = IznosPoKO(glava)
    If iznosi.count = 0 Then Exit Function
    Set isplate = modNovac.BuildIsplataDictByDoplata()

    Dim k As Variant, h As Variant, n As Long, otv As Double, ispl As Double
    Dim res() As Variant
    ReDim res(1 To iznosi.count, 1 To 9)
    For Each k In iznosi.keys
        h = glava(k)
        If Len(Trim$(kooperantID)) > 0 Then
            If StrComp(CStr(h(2)), Trim$(kooperantID), vbTextCompare) <> 0 Then GoTo Sledeci
        End If
        ispl = 0
        If isplate.Exists(CStr(k)) Then ispl = CDbl(isplate(CStr(k)))
        otv = modNovac.ZaokruziNovac(CDbl(iznosi(k)) - ispl)
        If otv <= 0 Then GoTo Sledeci
        n = n + 1
        res(n, 1) = CStr(k)
        res(n, 2) = h(0)
        res(n, 3) = h(1)
        res(n, 4) = CStr(h(2))
        res(n, 5) = CStr(h(3))
        res(n, 6) = modNovac.ZaokruziNovac(CDbl(iznosi(k)))
        res(n, 7) = modNovac.ZaokruziNovac(ispl)
        res(n, 8) = otv
        res(n, 9) = CStr(h(4))
Sledeci:
    Next k
    If n = 0 Then Exit Function

    Dim outA() As Variant, i As Long, j As Long
    ReDim outA(1 To n, 1 To 9)
    For i = 1 To n
        For j = 1 To 9
            outA(i, j) = res(i, j)
        Next j
    Next i
    OtvoreniKORedovi = outA
End Function

' Bruto iznos AKTIVNOG dokumenta doplate (zbir stavki); 0 kad nije aktivan.
Public Function IznosKO(ByVal doplataID As String) As Double
    Dim glava As Object, iznosi As Object, id As String
    id = Trim$(doplataID)
    If Len(id) = 0 Then Exit Function
    Set iznosi = IznosPoKO(glava)
    If iznosi.Exists(id) Then IznosKO = modNovac.ZaokruziNovac(CDbl(iznosi(id)))
End Function

' Otvoreno na JEDNOM KO, u cent-domenu: iznos minus isplaceno. 0 za storniran ili
' nepostojeci KO -- kapiju vlasnistva i zivota drzi pisac pre ovoga
' (modBankaMapiranje.PotvrdiKOZaKooperanta).
Public Function OtvorenoNaKO(ByVal doplataID As String) As Double
    Dim glava As Object, iznosi As Object, isplate As Object, id As String
    id = Trim$(doplataID)
    If Len(id) = 0 Then Exit Function
    Set iznosi = IznosPoKO(glava)
    If Not iznosi.Exists(id) Then Exit Function
    Set isplate = modNovac.BuildIsplataDictByDoplata()
    Dim ispl As Double
    If isplate.Exists(id) Then ispl = CDbl(isplate(id))
    OtvorenoNaKO = modNovac.ZaokruziNovac(CDbl(iznosi(id)) - ispl)
End Function

' POZIV NA BROJ -> AKTIVAN KO. Array(DoplataID, KooperantID) kad TACNO jedan
' aktivan KO nosi taj broj; Empty inace. Blizanac
' modBankaMapiranje.TryResolveKooperantByOtkupPoziv: isto poredjenje
' (NormalizePozivKey, slova i cifre), isti odgovor na storno (storniran KO se ne
' razresava -- uplata tada ide putem kooperanta, kao za storniran otkup).
'
' Broj doplate je broj bloka + "-0" + redni broj (DOP-ODL-30): od cifara kao i
' broj otkupa, pa vrstu ne kaze poziv nego POGODAK -- uvoz izvoda trazi tacno
' jedan pogodak preko blokova, doplata i specifikacija
' (modBankaMapiranje.ResolveStrongKooperantByKeys / PozivViseVrsta).
Public Function KOPoPozivu(ByVal pozivNaBroj As String) As Variant
    Const SRC As String = "modDoplata.KOPoPozivu"
    KOPoPozivu = Empty

    Dim target As String
    target = modBankaMapiranje.NormalizePozivKey(pozivNaBroj)
    If Len(target) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function
    d = ExcludeStornirano(d, TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long, cBr As Long, cKoop As Long
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cBr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_BROJ, SRC)
    cKoop = RequireColumnIndex(TBL_DOPLATA, COL_DOP_KOOPERANT, SRC)

    Dim i As Long, hits As Long, hitID As String, hitKoop As String
    For i = 1 To UBound(d, 1)
        If modBankaMapiranje.NormalizePozivKey(NzToText(d(i, cBr))) = target Then
            hits = hits + 1
            hitID = Trim$(NzToText(d(i, cId)))
            hitKoop = Trim$(NzToText(d(i, cKoop)))
        End If
    Next i

    If hits = 1 And Len(hitID) > 0 And Len(hitKoop) > 0 Then KOPoPozivu = Array(hitID, hitKoop)
End Function

' KAPIJA PISCA NOVCA NAD KO: vlasnistvo i zivot dokumenta u trenutku upisa
' (obrazac IsplataBlokProblem / PotvrdiOtkupZaKooperanta). "" kad KO postoji
' tacno jednom, nije storniran i pripada kooperantu; inace KO_KAPIJA_*.
' Ostatak (koliko KO jos duguje) NIJE deo kapije -- svaki pisac ga cita sam
' (OtvorenoNaKO), jer ga razlicito tumaci: izvod visak salje u avans, gotovina
' ga odbija.
Public Function KOKapija(ByVal doplataID As String, ByVal kooperantID As String) As String
    Const SRC As String = "modDoplata.KOKapija"
    Dim id As String
    id = Trim$(doplataID)
    KOKapija = KO_KAPIJA_NEMA
    If Len(id) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long, cKoop As Long, cSt As Long
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cKoop = RequireColumnIndex(TBL_DOPLATA, COL_DOP_KOOPERANT, SRC)
    cSt = RequireColumnIndex(TBL_DOPLATA, COL_STORNIRANO, SRC)

    Dim i As Long, n As Long, r As Long
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cId))), id, vbTextCompare) = 0 Then
            n = n + 1
            r = i
        End If
    Next i
    If n = 0 Then Exit Function
    If n > 1 Then
        KOKapija = KO_KAPIJA_DUPLI
        Exit Function
    End If
    If UCase$(Trim$(NzToText(d(r, cSt)))) = "DA" Then
        KOKapija = KO_KAPIJA_STORNIRAN
        Exit Function
    End If
    If StrComp(Trim$(NzToText(d(r, cKoop))), Trim$(kooperantID), vbTextCompare) <> 0 Then
        KOKapija = KO_KAPIJA_TUDJ
        Exit Function
    End If
    KOKapija = ""
End Function

' Da li je ID knjizno odobrenje (bilo kog statusa). Lista naloga nosi KO pod
' istim kljucem bloka kao otkup (clsBlokIsplata.otkupID), pa radnja koja vazi
' samo za otkup (primena avansa) ovim pita pre nego sto ponudi potvrdu.
Public Function JeKO(ByVal id As String) As Boolean
    If Len(Trim$(id)) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function
    JeKO = (Len(NzToText(LookupValue(TBL_DOPLATA, COL_DOP_ID, Trim$(id), COL_DOP_ID))) > 0)
End Function

' Stanica BLOKA dokumenta doplate (zaglavlje, bilo kog statusa) -- vlasnik
' isplate (DOP-ODL-19): izvod i gotovina knjize na nju, kao za otkupni blok.
' "" kad dokumenta nema.
Public Function StanicaKO(ByVal doplataID As String) As String
    If Len(Trim$(doplataID)) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function
    StanicaKO = Trim$(NzToText(LookupValue(TBL_DOPLATA, COL_DOP_ID, Trim$(doplataID), COL_DOP_STANICA)))
End Function

' ============================================================
' SPECIFIKACIJA KAO CILJ PLACANJA (DOP-ODL-27..29)
' ============================================================
' Specifikacija nema svoj red ni tabelu: to je BROJ na zaglavljima dokumenata
' istog para (obracun, kooperant), 1:1 sa parom (DoplataStavkeRedovi +56/+57).
' Novac se NIKAD ne vezuje za specifikaciju -- raspodela ga vezuje za svaki
' dokument posebno (DoplataID), pa ostatak, storno (DOP-05), kartica i saldo OM
' ostaju po dokumentu, kao i do sada. Ovde su samo citaci; pisci su izvod
' (modBankaMapiranje) i primena avansa (modNovac.ApplyAvansToSpec_TX).

' OTVORENE SPECIFIKACIJE -- blizanac OtvoreniKORedovi za naloge za banku.
' 2D niz (1..n, 1..8):
'   1 BrojSpecifikacije | 2 Datum (knjizenja obracuna) | 3 KooperantID
'   4 StanicaID: MATICNA stanica kooperanta -- dokumenti iste specifikacije mogu
'     biti sa vise stanica, a red naloga je jedan
'   5 Iznos (zbir aktivnih dokumenata) | 6 Isplaceno | 7 Otvoreno
'   8 broj otvorenih dokumenata
' Otvoreno je zbir OTVORENOG po dokumentu (najmanje 0 po dokumentu), ne iznos
' minus isplaceno: preplata jednog dokumenta ne sme da umanji dug drugog, jer ga
' raspodela ionako ne bi prenela. U listi je specifikacija ciji je ostatak, u
' cent-domenu, veci od nule; kooperantID suzava. Empty kad nema nijedne.
Public Function OtvoreneSpecifikacijeRedovi(Optional ByVal kooperantID As String = "") As Variant
    OtvoreneSpecifikacijeRedovi = Empty

    Dim glava As Object, iznosi As Object, isplate As Object
    Set iznosi = IznosPoKO(glava)
    If iznosi.count = 0 Then Exit Function
    Set isplate = modNovac.BuildIsplataDictByDoplata()

    Dim sp As Object, k As Variant, h As Variant, s As String, a As Variant
    Dim ispl As Double, otv As Double
    Set sp = CreateObject("Scripting.Dictionary")
    sp.CompareMode = vbTextCompare
    For Each k In iznosi.keys
        h = glava(k)
        s = CStr(h(4))
        If Len(s) = 0 Then GoTo Sledeci
        If Len(Trim$(kooperantID)) > 0 Then
            If StrComp(CStr(h(2)), Trim$(kooperantID), vbTextCompare) <> 0 Then GoTo Sledeci
        End If
        ispl = 0
        If isplate.Exists(CStr(k)) Then ispl = CDbl(isplate(CStr(k)))
        otv = modNovac.ZaokruziNovac(CDbl(iznosi(k)) - ispl)
        If otv < 0 Then otv = 0
        If Not sp.Exists(s) Then sp.Add s, Array(h(1), CStr(h(2)), 0#, 0#, 0#, 0&)
        a = sp(s)
        a(2) = CDbl(a(2)) + CDbl(iznosi(k))
        a(3) = CDbl(a(3)) + ispl
        a(4) = CDbl(a(4)) + otv
        If otv > 0 Then a(5) = CLng(a(5)) + 1
        sp(s) = a
Sledeci:
    Next k
    If sp.count = 0 Then Exit Function

    Dim res() As Variant, n As Long
    ReDim res(1 To sp.count, 1 To 8)
    For Each k In sp.keys
        a = sp(k)
        If modNovac.ZaokruziNovac(CDbl(a(4))) > 0 Then
            n = n + 1
            res(n, 1) = CStr(k)
            res(n, 2) = a(0)
            res(n, 3) = CStr(a(1))
            res(n, 4) = Trim$(NzToText(LookupValue(TBL_KOOPERANTI, COL_KOOP_ID, CStr(a(1)), COL_KOOP_STANICA)))
            res(n, 5) = modNovac.ZaokruziNovac(CDbl(a(2)))
            res(n, 6) = modNovac.ZaokruziNovac(CDbl(a(3)))
            res(n, 7) = modNovac.ZaokruziNovac(CDbl(a(4)))
            res(n, 8) = CLng(a(5))
        End If
    Next k
    If n = 0 Then Exit Function

    Dim outA() As Variant, i As Long, j As Long
    ReDim outA(1 To n, 1 To 8)
    For i = 1 To n
        For j = 1 To 8
            outA(i, j) = res(i, j)
        Next j
    Next i
    OtvoreneSpecifikacijeRedovi = outA
End Function

' DOKUMENTI SPECIFIKACIJE ZA RASPODELU (DOP-ODL-28): OTVORENI aktivni dokumenti
' specifikacije, OD NAJSTARIJEG BLOKA -- datum bloka, pa broj bloka (cifre se
' porede kao brojevi: 9/... ide pre 10/...). 2D niz (1..n, 1..3):
'   1 DoplataID | 2 StanicaID bloka | 3 Otvoreno (cent-domen)
' stanicaID (opciono) suzava na dokumente ciji je blok na toj stanici (gotovina
' po specifikaciji, DOP-ODL-28). Empty kad otvorenog nema. Read-only.
Public Function SpecDokumentiZaRaspodelu(ByVal brojSpec As String, _
                                         Optional ByVal stanicaID As String = "") As Variant
    SpecDokumentiZaRaspodelu = Empty
    brojSpec = Trim$(brojSpec)
    If Len(brojSpec) = 0 Then Exit Function

    Dim glava As Object, iznosi As Object, isplate As Object
    Set iznosi = IznosPoKO(glava)
    If iznosi.count = 0 Then Exit Function
    Set isplate = modNovac.BuildIsplataDictByDoplata()

    Dim k As Variant, h As Variant, ispl As Double, otv As Double, n As Long
    Dim ids() As String, sta() As String, iz() As Double, kljuc() As String
    ReDim ids(1 To iznosi.count)
    ReDim sta(1 To iznosi.count)
    ReDim iz(1 To iznosi.count)
    ReDim kljuc(1 To iznosi.count)
    For Each k In iznosi.keys
        h = glava(k)
        If StrComp(CStr(h(4)), brojSpec, vbTextCompare) <> 0 Then GoTo Sledeci
        If Len(Trim$(stanicaID)) > 0 Then
            If StrComp(CStr(h(3)), Trim$(stanicaID), vbTextCompare) <> 0 Then GoTo Sledeci
        End If
        ispl = 0
        If isplate.Exists(CStr(k)) Then ispl = CDbl(isplate(CStr(k)))
        otv = modNovac.ZaokruziNovac(CDbl(iznosi(k)) - ispl)
        If otv <= 0 Then GoTo Sledeci
        n = n + 1
        ids(n) = CStr(k)
        sta(n) = CStr(h(3))
        iz(n) = otv
        kljuc(n) = RedosledBloka(CStr(h(5)))
Sledeci:
    Next k
    If n = 0 Then Exit Function

    ' Umetanje: dokumenata jedne specifikacije je malo (blokovi jednog
    ' kooperanta u jednom obracunu).
    Dim i As Long, j As Long, tId As String, tSt As String, tIz As Double, tK As String
    For i = 2 To n
        tId = ids(i): tSt = sta(i): tIz = iz(i): tK = kljuc(i)
        j = i - 1
        Do While j >= 1
            If StrComp(kljuc(j), tK, vbBinaryCompare) <= 0 Then Exit Do
            ids(j + 1) = ids(j): sta(j + 1) = sta(j): iz(j + 1) = iz(j): kljuc(j + 1) = kljuc(j)
            j = j - 1
        Loop
        ids(j + 1) = tId: sta(j + 1) = tSt: iz(j + 1) = tIz: kljuc(j + 1) = tK
    Next i

    Dim outA() As Variant
    ReDim outA(1 To n, 1 To 3)
    For i = 1 To n
        outA(i, 1) = ids(i)
        outA(i, 2) = sta(i)
        outA(i, 3) = iz(i)
    Next i
    SpecDokumentiZaRaspodelu = outA
End Function

' Otvoreno na specifikaciji: zbir otvorenog po njenim aktivnim dokumentima (sve
' stanice), u cent-domenu; 0 kad nema otvorenog ili specifikacije.
Public Function OtvorenoNaSpec(ByVal brojSpec As String) As Double
    Dim dok As Variant, i As Long, s As Double
    dok = SpecDokumentiZaRaspodelu(brojSpec)
    If Not IsArray(dok) Then Exit Function
    For i = 1 To UBound(dok, 1)
        s = s + CDbl(dok(i, 3))
    Next i
    OtvorenoNaSpec = modNovac.ZaokruziNovac(s)
End Function

' Kljuc redosleda bloka: datum (yyyymmdd) pa broj bloka u kome je svaki niz
' cifara dopunjen nulama do 10 mesta, pa se poredi kao tekst. Blok kog nema
' (ne bi smelo -- doplata je po aktivnom bloku) ide na kraj.
Private Function RedosledBloka(ByVal otkupID As String) As String
    Dim dt As Variant, br As String, i As Long, ch As String, cif As String, out As String
    dt = LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_DATUM)
    If Not IsDate(dt) Then
        RedosledBloka = "99999999|"
        Exit Function
    End If
    br = UCase$(Trim$(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, otkupID, COL_OTK_BR_DOK))))
    For i = 1 To Len(br)
        ch = Mid$(br, i, 1)
        If ch Like "#" Then
            cif = cif & ch
        Else
            If Len(cif) > 0 Then out = out & Right$(String$(10, "0") & cif, 10): cif = ""
            out = out & ch
        End If
    Next i
    If Len(cif) > 0 Then out = out & Right$(String$(10, "0") & cif, 10)
    RedosledBloka = Format$(CDate(dt), "yyyymmdd") & "|" & out
End Function

' POZIV NA BROJ -> SPECIFIKACIJA. Array(BrojSpecifikacije, KooperantID) kad TACNO
' jedna specifikacija sa bar jednim AKTIVNIM dokumentom nosi taj broj; Empty
' inace. Poredjenje je NormalizePozivKey (slova i cifre), kao KOPoPozivu; dva
' pogotka nisu "prvi pobedjuje" nego rucno mapiranje. Broj je od cifara
' (x/ddmmyy-00y, DOP-ODL-30), pa vrstu ne kaze poziv nego pogodak (v. KOPoPozivu).
Public Function SpecPoPozivu(ByVal pozivNaBroj As String) As Variant
    Const SRC As String = "modDoplata.SpecPoPozivu"
    SpecPoPozivu = Empty

    Dim target As String
    target = modBankaMapiranje.NormalizePozivKey(pozivNaBroj)
    If Len(target) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function
    d = ExcludeStornirano(d, TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cSpec As Long, cKoop As Long
    cSpec = RequireColumnIndex(TBL_DOPLATA, COL_DOP_SPEC, SRC)
    cKoop = RequireColumnIndex(TBL_DOPLATA, COL_DOP_KOOPERANT, SRC)

    Dim pog As Object, i As Long, s As String, koop As String
    Set pog = CreateObject("Scripting.Dictionary")
    pog.CompareMode = vbTextCompare
    For i = 1 To UBound(d, 1)
        s = Trim$(NzToText(d(i, cSpec)))
        If Len(s) > 0 Then
            If modBankaMapiranje.NormalizePozivKey(s) = target Then
                koop = Trim$(NzToText(d(i, cKoop)))
                If pog.Exists(s) Then
                    ' Isti broj kod dva kooperanta je kvar (citac +57); ovde se
                    ' samo ne razresava.
                    If StrComp(CStr(pog(s)), koop, vbTextCompare) <> 0 Then Exit Function
                Else
                    pog.Add s, koop
                End If
            End If
        End If
    Next i

    If pog.count <> 1 Then Exit Function
    s = CStr(pog.keys()(0))
    koop = CStr(pog(s))
    If Len(koop) = 0 Then Exit Function
    SpecPoPozivu = Array(s, koop)
End Function

' Vlasnik specifikacije: kooperant njenih AKTIVNIH dokumenata (tacan broj, bez
' normalizacije). "" kad aktivnog dokumenta nema ili ga nose dva kooperanta --
' pisac tada odbija po imenu (kapija vlasnistva, obrazac KOKapija).
Public Function SpecKooperant(ByVal brojSpec As String) As String
    Const SRC As String = "modDoplata.SpecKooperant"
    brojSpec = Trim$(brojSpec)
    If Len(brojSpec) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function
    d = ExcludeStornirano(d, TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cSpec As Long, cKoop As Long, i As Long, koop As String, k As String
    cSpec = RequireColumnIndex(TBL_DOPLATA, COL_DOP_SPEC, SRC)
    cKoop = RequireColumnIndex(TBL_DOPLATA, COL_DOP_KOOPERANT, SRC)
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cSpec))), brojSpec, vbTextCompare) = 0 Then
            k = Trim$(NzToText(d(i, cKoop)))
            If Len(koop) = 0 Then
                koop = k
            ElseIf StrComp(koop, k, vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
    Next i
    SpecKooperant = koop
End Function

' Da li je kljuc reda broj specifikacije (bilo kog statusa). Lista naloga nosi
' specifikaciju pod istim kljucem bloka kao otkup i KO (clsBlokIsplata.otkupID),
' pa radnja nad redom (primena avansa) ovim bira motor. OtkupID ("OTK-...") i
' DoplataID ("DOP-...") nemaju kosu crtu, a broj specifikacije ima, pa kljucevi
' reda ni ne citaju tabelu.
Public Function JeSpecifikacija(ByVal id As String) As Boolean
    id = Trim$(id)
    If Len(id) = 0 Then Exit Function
    If InStr(1, id, "/", vbBinaryCompare) = 0 Then Exit Function
    If GetTable(TBL_DOPLATA) Is Nothing Then Exit Function
    JeSpecifikacija = (Len(NzToText(LookupValue(TBL_DOPLATA, COL_DOP_SPEC, id, COL_DOP_SPEC))) > 0)
End Function

' Iznos svakog AKTIVNOG KO: DoplataID -> din (zbir stavki, DoplataStavkeRedovi --
' strog citac, pa storniran KO ne ulazi). outGlava: DoplataID ->
' Array(BrojDokumenta, Datum, KooperantID, StanicaID bloka, BrojSpecifikacije,
' OtkupID bloka).
Private Function IznosPoKO(ByRef outGlava As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set outGlava = CreateObject("Scripting.Dictionary")
    outGlava.CompareMode = vbTextCompare
    Set IznosPoKO = d

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    Dim i As Long, id As String
    For i = 1 To UBound(r, 1)
        id = CStr(r(i, 1))
        If Not d.Exists(id) Then
            d.Add id, 0#
            outGlava.Add id, Array(r(i, 2), r(i, 3), CStr(r(i, 4)), CStr(r(i, 6)), _
                                   Trim$(CStr(r(i, 20))), Trim$(CStr(r(i, 16))))
        End If
        d(id) = CDbl(d(id)) + CDbl(r(i, 12))
    Next i
End Function

' OSNOVICA KNJIZENE DOPLATE SE PROMENILA (DOP-ODL-07). Knjizen dokument je
' SNIMAK: kasniji storno, ispravka ili retro unos otkupa ga ne menja -- ovde se
' to VIDI. Signal, ne popravka: ispravka je storno doplate pa nov obracun.
'
' Po BLOKU (DOPLATA.md par. 2a), dve vrste nalaza:
'   - stavka knjizenog DOKUMENTA (blok, klasa) cija se kolicina razlikuje od
'     sadasnje na AKTIVNOM otkupu -- storniran blok ima 0 kg. Poredi se svaki
'     dokument posebno sa svojim snimkom: ponovljena doplata (D, D2) je ista
'     roba, pa se kg preko dokumenata NE sabira (review P1, DOP-P1-REV-01);
'   - RETRO UNOS: aktivan otkup u periodu obracuna, za kljuc obracuna, NASTAO
'     POSLE knjizenja (CreatedAt), a bez AKTIVNE doplate u BILO KOM obracunu.
'     Blok placen kasnijim obracunom (npr. nov obracun samo za njega) nije
'     nalaz -- nema sta da se ispravi; storno te doplate nalaz vraca (odluka
'     05.11.2026, review P1 DOP-P1-REV-04; test 337).
' Blok koji je postojao pri knjizenju a operater ga je iskljucio (DOP-ODL-16)
' NIJE nalaz: plan se zato ne racuna ponovo nad celim periodom.
' Obracun radjen SAMO ZA JEDNOG kooperanta (DOP-ODL-31) trazi retro unos samo
' medju NJEGOVIM blokovima -- isto pravilo obuhvata kao plan (UObuhvatu). Tudji
' blok tom obracunu nikad nije pripadao, pa ni na granicnom danu bez poznatog
' sata ne daje "osnovica se ne moze proveriti".
'
' 2D niz (1..n, 1..6): 1 broj doplate ("" za retro unos) | 2 KooperantID |
' 3 kljuc ("vrsta sorta klasa @ stanica (blok broj)") | 4 kg knjizeno |
' 5 kg sada | 6 ObracunID. Obracun ciji se period danas ne moze proveriti (npr.
' granicni dan bez poznatog sata) daje JEDAN red sa razlogom u koloni 3.
' Empty kad odstupanja nema. Read-only.
Public Function DoplataOdstupanjaOsnovice() As Variant
    Const SRC As String = "modDoplata.DoplataOdstupanjaOsnovice"
    DoplataOdstupanjaOsnovice = Empty

    Dim r As Variant
    r = DoplataStavkeRedovi()
    If Not IsArray(r) Then Exit Function

    ' Trenutak knjizenja svakog dokumenta (CreatedAt zaglavlja).
    Dim hd As Variant, kad As Object, i As Long, cId As Long, cCre As Long
    Set kad = CreateObject("Scripting.Dictionary")
    kad.CompareMode = vbTextCompare
    hd = GetTableData(TBL_DOPLATA)
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cCre = RequireColumnIndex(TBL_DOPLATA, COL_AUDIT_CREATED_AT, SRC)
    If IsArray(hd) Then
        For i = 1 To UBound(hd, 1)
            kad(Trim$(NzToText(hd(i, cId)))) = hd(i, cCre)
        Next i
    End If

    ' knjizeno: "DoplataID|KLASA" -> Array(kg, broj, koop, opis, ObracunID,
    '           "OtkupID|KLASA") -- po DOKUMENTU, ne po bloku (v. gore)
    ' blokovi:  "OtkupID|KLASA" koji imaju bar jedan aktivan dokument
    ' obr:      ObracunID -> Array(periodOd, periodDo, kljucevi vs|KLASA, knjizeno u,
    '                              obuhvat)
    Dim knj As Object, blokovi As Object, obr As Object, o As Variant, kl As Object, k As String
    Set knj = CreateObject("Scripting.Dictionary")
    knj.CompareMode = vbTextCompare
    Set blokovi = CreateObject("Scripting.Dictionary")
    blokovi.CompareMode = vbTextCompare
    Set obr = CreateObject("Scripting.Dictionary")
    obr.CompareMode = vbTextCompare
    Dim cr As Variant, kBlok As String
    For i = 1 To UBound(r, 1)
        kBlok = CStr(r(i, 16)) & "|" & UCase$(Trim$(CStr(r(i, 9))))
        k = CStr(r(i, 1)) & "|" & UCase$(Trim$(CStr(r(i, 9))))
        blokovi(kBlok) = True
        If knj.Exists(k) Then
            o = knj(k)
            o(0) = CDbl(o(0)) + CDbl(r(i, 10))
            knj(k) = o
        Else
            knj.Add k, Array(CDbl(r(i, 10)), CStr(r(i, 2)), CStr(r(i, 4)), _
                             OpisKljuca(r(i, 7), r(i, 8), r(i, 9), r(i, 6)), CStr(r(i, 5)), kBlok)
        End If
        If Not obr.Exists(CStr(r(i, 5))) Then
            Set kl = CreateObject("Scripting.Dictionary")
            kl.CompareMode = vbTextCompare
            cr = Empty
            If kad.Exists(CStr(r(i, 1))) Then cr = kad(CStr(r(i, 1)))
            obr.Add CStr(r(i, 5)), Array(r(i, 14), r(i, 15), kl, cr, CStr(r(i, 21)))
        End If
        o = obr(CStr(r(i, 5)))
        Set kl = o(2)
        kl(KljucVrstaSorta(r(i, 7), r(i, 8)) & "|" & UCase$(Trim$(CStr(r(i, 9))))) = True
    Next i

    ' Trenutno stanje AKTIVNIH otkupa: zaglavlja i kg po "OtkupID|KLASA".
    ' kls: OtkupID -> "KLASA;KLASA;" -- klase bloka, da retro provera ne prolazi
    ' kroz sve stavke za svaki otkup.
    Dim d As Variant, zag As Object, sada As Object, kls As Object
    Set zag = CreateObject("Scripting.Dictionary")
    zag.CompareMode = vbTextCompare
    Set sada = CreateObject("Scripting.Dictionary")
    sada.CompareMode = vbTextCompare
    Set kls = CreateObject("Scripting.Dictionary")
    kls.CompareMode = vbTextCompare
    d = GetTableData(TBL_OTKUP)
    If IsArray(d) Then d = ExcludeStornirano(d, TBL_OTKUP)
    If IsArray(d) Then
        Dim zId As Long, zDat As Long, zKoop As Long, zSta As Long, zVr As Long
        Dim zSo As Long, zBr As Long, zSrc As Long, zCre As Long
        zId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
        zDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
        zKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
        zSta = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
        zVr = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)
        zSo = RequireColumnIndex(TBL_OTKUP, COL_OTK_SORTA, SRC)
        zBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
        zSrc = RequireColumnIndex(TBL_OTKUP, COL_OTK_SOURCE_CREATED_AT, SRC)
        zCre = RequireColumnIndex(TBL_OTKUP, COL_AUDIT_CREATED_AT, SRC)
        For i = 1 To UBound(d, 1)
            zag(Trim$(NzToText(d(i, zId)))) = Array(d(i, zDat), d(i, zSrc), d(i, zCre), _
                Trim$(NzToText(d(i, zKoop))), Trim$(NzToText(d(i, zSta))), _
                d(i, zVr), d(i, zSo), Trim$(NzToText(d(i, zBr))))
        Next i
        Dim s As Variant
        s = modOtkup.StavkeOtkupaRedovi()
        If IsArray(s) Then
            For i = 1 To UBound(s, 1)
                If zag.Exists(CStr(s(i, 1))) Then
                    k = CStr(s(i, 1)) & "|" & UCase$(Trim$(CStr(s(i, 3))))
                    If Not sada.Exists(k) Then _
                        kls(CStr(s(i, 1))) = kls(CStr(s(i, 1))) & UCase$(Trim$(CStr(s(i, 3)))) & ";"
                    sada(k) = CDbl(sada(k)) + CDbl(s(i, 4))
                End If
            Next i
        End If
    End If

    Dim out As New Collection, kk As Variant, v As Variant, kgSada As Double
    For Each kk In knj.keys
        v = knj(kk)
        kgSada = 0
        If sada.Exists(CStr(v(5))) Then kgSada = CDbl(sada(CStr(v(5))))
        If Abs(kgSada - CDbl(v(0))) >= 0.005 Then
            out.Add Array(v(1), v(2), v(3), CDbl(v(0)), kgSada, v(4))
        End If
    Next kk

    ' Retro unos, po obracunu. Redosled provera je namerno: vreme nastanka, pa
    ' KLJUC, pa tek onda period -- period na granicnom danu bez poznatog sata
    ' dize gresku, a otkup tudjeg kljuca ne sme da obori proveru ovog obracuna.
    Dim ok As Variant, z As Variant, greska As String, uPer As Boolean
    Dim oid As Variant, delovi As Variant, p As Long, klasa As String, kRed As String
    Dim nove As Collection, nv As Variant
    For Each ok In obr.keys
        o = obr(ok)
        Set kl = o(2)
        If Not IsDate(o(3)) Then GoTo SledeciObracun
        greska = ""
        For Each oid In zag.keys
            z = zag(oid)
            If Not IsDate(z(2)) Then GoTo SledeciOtkup
            ' Nastao pre knjizenja = postojao je, pa ga je operater svesno
            ' iskljucio (ili nije bio u kljucu). Isti sekund se racuna kao posle.
            If CDate(z(2)) < CDate(o(3)) Then GoTo SledeciOtkup
            ' Van obuhvata obracuna (DOP-ODL-31): nije mu pripadao, nije nalaz.
            If Not UObuhvatu(z(3), CStr(o(4))) Then GoTo SledeciOtkup
            If Not kls.Exists(CStr(oid)) Then GoTo SledeciOtkup
            Set nove = New Collection
            delovi = Split(CStr(kls(CStr(oid))), ";")
            For p = LBound(delovi) To UBound(delovi)
                klasa = CStr(delovi(p))
                kRed = CStr(oid) & "|" & klasa
                If Len(klasa) > 0 Then
                    If kl.Exists(KljucVrstaSorta(z(5), z(6)) & "|" & klasa) And Not blokovi.Exists(kRed) Then
                        nove.Add Array(klasa, kRed)
                    End If
                End If
            Next p
            If nove.count = 0 Then GoTo SledeciOtkup
            On Error Resume Next
            uPer = OtkupUPeriodu(z(0), z(1), z(2), CDate(o(0)), CDate(o(1)), CStr(oid), CStr(z(7)), SRC)
            If Err.Number <> 0 Then greska = Err.description
            Err.Clear
            On Error GoTo 0
            If Len(greska) > 0 Then Exit For
            If uPer Then
                For Each nv In nove
                    out.Add Array("", z(3), OpisKljuca(z(5), z(6), nv(0), z(4)) & " (blok " & z(7) & ")", _
                                  0#, CDbl(sada(CStr(nv(1)))), CStr(ok))
                Next nv
            End If
SledeciOtkup:
        Next oid
        If Len(greska) > 0 Then
            out.Add Array("", "", "osnovica se ne moze proveriti: " & greska, 0#, 0#, CStr(ok))
        End If
SledeciObracun:
    Next ok

    If out.count = 0 Then Exit Function
    Dim res() As Variant, red As Variant, j As Long
    ReDim res(1 To out.count, 1 To 6)
    For i = 1 To out.count
        red = out(i)
        For j = 1 To 6
            res(i, j) = red(j - 1)
        Next j
    Next i
    DoplataOdstupanjaOsnovice = res
End Function

' "vrsta sorta klasa @ stanica" -- citljiv kljuc reda za nalaz integriteta.
Private Function OpisKljuca(ByVal vrsta As Variant, ByVal sorta As Variant, _
                            ByVal klasa As Variant, ByVal stanica As Variant) As String
    OpisKljuca = Trim$(Trim$(NzToText(vrsta) & " " & NzToText(sorta)) & " " & NzToText(klasa)) & _
                 " @ " & NzToText(stanica)
End Function

' DoplataID-evi AKTIVNIH KO jednog obracuna -- za storno celog obracuna
' (modStorno.StornoObracunDoplate_TX). Prazna kolekcija kad ih nema.
Public Function AktivneDoplateObracuna(ByVal obracunID As String) As Collection
    Const SRC As String = "modDoplata.AktivneDoplateObracuna"

    Dim c As New Collection
    Set AktivneDoplateObracuna = c

    obracunID = Trim$(obracunID)
    If Len(obracunID) = 0 Then Exit Function

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function
    d = ExcludeStornirano(d, TBL_DOPLATA)
    If Not IsArray(d) Then Exit Function

    Dim cId As Long, cObr As Long, i As Long
    cId = RequireColumnIndex(TBL_DOPLATA, COL_DOP_ID, SRC)
    cObr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_OBRACUN, SRC)
    For i = 1 To UBound(d, 1)
        If StrComp(Trim$(NzToText(d(i, cObr))), obracunID, vbTextCompare) = 0 Then
            c.Add Trim$(NzToText(d(i, cId)))
        End If
    Next i
End Function

' ============================================================
' BROJ DOPLATE I SPECIFIKACIJE -- bez slova (DOP-ODL-30)
' ============================================================

' Broj n-te doplate na bloku: broj bloka + "-0" + n. "5/250626-13", 1 ->
' "5/250626-13-01"; "5/280626", 2 -> "5/280626-02".
Public Function BrojDoplate(ByVal brojBloka As String, ByVal redni As Long) As String
    BrojDoplate = Trim$(brojBloka) & DOP_BROJ_ODVAJAC & CStr(redni)
End Function

' Broj specifikacije: x/ddmmyy-00y -- broj maticne stanice kooperanta, datum
' knjizenja, redni broj specifikacije te stanice tog dana: "STA-5", 05.11.2026,
' 1 -> "5/051126-001". Broj stanice daje isto pravilo kao broj bloka
' (modBrojevi.ExtractNumericFromEntityID, FormatBroj).
Public Function BrojSpecifikacije(ByVal stanicaID As String, ByVal datum As Date, _
                                  ByVal redni As Long) As String
    BrojSpecifikacije = CStr(modBrojevi.ExtractNumericFromEntityID(stanicaID)) & "/" & _
                        Format$(datum, "ddmmyy") & DOP_SPEC_ODVAJAC & CStr(redni)
End Function

' Kljuc niza ("x/ddmmyy") i redni broj iz broja specifikacije x/ddmmyy-00y;
' False kad broj nije tog oblika.
Private Function ParsirajSpecifikaciju(ByVal spec As String, ByRef outKljuc As String, _
                                       ByRef outRedni As Long) As Boolean
    Dim p As Long, kljuc As String, sRb As String, ps As Long, sX As String, sD As String
    spec = Trim$(spec)
    p = InStrRev(spec, DOP_SPEC_ODVAJAC)
    If p <= 1 Then Exit Function
    kljuc = Left$(spec, p - 1)
    sRb = Mid$(spec, p + Len(DOP_SPEC_ODVAJAC))
    If Len(sRb) = 0 Or Len(sRb) > 9 Or sRb Like "*[!0-9]*" Then Exit Function
    If Left$(sRb, 1) = "0" Then Exit Function
    ps = InStr(1, kljuc, "/")
    If ps <= 1 Then Exit Function
    sX = Left$(kljuc, ps - 1)
    sD = Mid$(kljuc, ps + 1)
    If sX Like "*[!0-9]*" Then Exit Function
    If Len(sD) <> 6 Or sD Like "*[!0-9]*" Then Exit Function
    outKljuc = kljuc
    outRedni = CLng(sRb)
    ParsirajSpecifikaciju = True
End Function

' Stari oblik S<redni>/<godina>-<sifra> (pre DOP-ODL-30, samo u test sveskama).
Private Function JeStariOblikSpecifikacije(ByVal spec As String) As Boolean
    Dim p1 As Long, p2 As Long, sRb As String, sGod As String, pl As Long
    spec = Trim$(spec)
    pl = Len(DOP_SPEC_STARI_PREFIKS)
    If UCase$(Left$(spec, pl)) <> UCase$(DOP_SPEC_STARI_PREFIKS) Then Exit Function
    p1 = InStr(pl + 1, spec, "/")
    If p1 = 0 Then Exit Function
    p2 = InStr(p1 + 1, spec, "-")
    If p2 = 0 Then Exit Function
    sRb = Mid$(spec, pl + 1, p1 - pl - 1)
    sGod = Mid$(spec, p1 + 1, p2 - p1 - 1)
    If Len(sRb) = 0 Or sRb Like "*[!0-9]*" Then Exit Function
    If Len(sGod) <> 4 Or sGod Like "*[!0-9]*" Then Exit Function
    JeStariOblikSpecifikacije = True
End Function

' Iz SVIH dokumenata doplate (i storniranih -- storno ne oslobadja broj):
' outMax "x/ddmmyy" (stanica i dan) -> najveci redni broj specifikacije;
' outBrojevi skup zauzetih brojeva. Stari oblik S<redni>/<godina>-<sifra> je
' zauzet broj, ali ne ulazi u niz. Broj koji nije ni u jednom obliku pada po
' imenu -- pisac ga nije mogao napraviti.
Private Sub UcitajSpecifikacije(ByRef outMax As Object, ByRef outBrojevi As Object, _
                                ByVal sourceName As String)
    Set outMax = CreateObject("Scripting.Dictionary")
    outMax.CompareMode = vbTextCompare
    Set outBrojevi = CreateObject("Scripting.Dictionary")
    outBrojevi.CompareMode = vbTextCompare

    Dim dSpec As Variant
    dSpec = GetTableData(TBL_DOPLATA)
    If Not IsArray(dSpec) Then Exit Sub

    Dim cSpec As Long, i As Long, spec As String, kljuc As String, rb As Long
    cSpec = RequireColumnIndex(TBL_DOPLATA, COL_DOP_SPEC, sourceName)
    For i = 1 To UBound(dSpec, 1)
        spec = Trim$(NzToText(dSpec(i, cSpec)))
        If Len(spec) > 0 Then
            If ParsirajSpecifikaciju(spec, kljuc, rb) Then
                If rb > CLng(outMax(kljuc)) Then outMax(kljuc) = rb
            ElseIf Not JeStariOblikSpecifikacije(spec) Then
                Err.Raise DOP_ERR_BASE + 49, sourceName, _
                          "Broj specifikacije nije u obliku x/ddmmyy-00y: " & spec & "."
            End If
            outBrojevi(spec) = True
        End If
    Next i
End Sub

' Iz SVIH dokumenata doplate (i storniranih -- storno ne oslobadja broj, A9):
' outPoBloku OtkupID -> broj dosadasnjih dokumenata tog bloka; outBrojevi skup
' zauzetih brojeva.
Private Sub UcitajBrojeveDoplata(ByRef outPoBloku As Object, ByRef outBrojevi As Object, _
                                 ByVal sourceName As String)
    Set outPoBloku = CreateObject("Scripting.Dictionary")
    outPoBloku.CompareMode = vbTextCompare
    Set outBrojevi = CreateObject("Scripting.Dictionary")
    outBrojevi.CompareMode = vbTextCompare

    Dim d As Variant
    d = GetTableData(TBL_DOPLATA)
    If Not IsArray(d) Then Exit Sub

    Dim cOtk As Long, cBr As Long, i As Long, oid As String, br As String
    cOtk = RequireColumnIndex(TBL_DOPLATA, COL_DOP_OTKUP, sourceName)
    cBr = RequireColumnIndex(TBL_DOPLATA, COL_DOP_BROJ, sourceName)
    For i = 1 To UBound(d, 1)
        oid = Trim$(NzToText(d(i, cOtk)))
        If Len(oid) > 0 Then outPoBloku(oid) = CLng(outPoBloku(oid)) + 1
        br = Trim$(NzToText(d(i, cBr)))
        If Len(br) > 0 Then outBrojevi(br) = True
    Next i
End Sub

' ============================================================
' PERIOD SA SATOM (DOP-ODL-03a)
' ============================================================

' Da li otkup ulazi u period [Od, Do], obe granice ukljucene, na minut.
'
' DAN je Datum dokumenta. SAT je vreme nastanka po pravilu otkupnog lista
' (modOtkup.OtkupVremeNastanka: SourceCreatedAt, pa CreatedAt), i trazi se SAMO
' kad granica sece dan: Od posle 00:00 na prvom danu, Do pre 23:59 na poslednjem.
' Unutrasnji dani i cele granicne dane odlucuje datum.
'
' STROZE OD STAMPE, namerno: kad SourceCreatedAt postoji a nije vreme (PWA salje
' ISO tekst), sat se NE uzima iz CreatedAt -- za PWA red to je trenutak
' SINHRONIZACIJE, ne otkupa. Tada je sat nepoznat, a otkup na granicnom danu
' obara obracun po imenu: pogresan dinar na granici je gori od zastoja.
Private Function OtkupUPeriodu(ByVal datum As Variant, ByVal srcCreated As Variant, _
                               ByVal created As Variant, ByVal periodOd As Date, _
                               ByVal periodDo As Date, ByVal oid As String, _
                               ByVal broj As String, ByVal sourceName As String) As Boolean
    If Not IsDate(datum) Then
        Err.Raise DOP_ERR_BASE + 20, sourceName, _
                  "Otkup " & broj & " (" & oid & ") nema ispravan Datum."
    End If

    Dim dan As Long, danOd As Long, danDo As Long
    dan = Int(CDbl(CDate(datum)))
    danOd = Int(CDbl(periodOd))
    danDo = Int(CDbl(periodDo))

    If dan < danOd Or dan > danDo Then Exit Function

    Dim trebaSat As Boolean
    If dan = danOd And (MinutOd(periodOd) Mod MIN_U_DANU) > 0 Then trebaSat = True
    If dan = danDo And (MinutOd(periodDo) Mod MIN_U_DANU) < MIN_U_DANU - 1 Then trebaSat = True

    If Not trebaSat Then
        OtkupUPeriodu = True
        Exit Function
    End If

    Dim v As Variant
    If Len(Trim$(NzToText(srcCreated))) > 0 And Not IsDate(srcCreated) Then
        v = Empty
    Else
        v = modOtkup.OtkupVremeNastanka(srcCreated, created)
    End If

    If Not IsDate(v) Then
        Err.Raise DOP_ERR_BASE + 21, sourceName, _
                  "Otkup " & broj & " (" & oid & ") je na granicnom danu perioda, a vreme " & _
                  "nastanka mu nije poznato (SourceCreatedAt: '" & NzToText(srcCreated) & _
                  "'). Pomeri granicu na pocetak ili kraj dana."
    End If

    Dim m As Double
    m = CDbl(dan) * MIN_U_DANU + Hour(CDate(v)) * 60 + Minute(CDate(v))
    OtkupUPeriodu = (m >= MinutOd(periodOd) And m <= MinutOd(periodDo))
End Function

' Trenutak na minut: dan * 1440 + sat * 60 + minut. Sekunde ne ulaze.
Private Function MinutOd(ByVal t As Date) As Double
    MinutOd = Int(CDbl(t)) * CDbl(MIN_U_DANU) + Hour(t) * 60 + Minute(t)
End Function

' ============================================================
' KAPIJE ULAZA
' ============================================================

Private Sub RequirePeriod(ByVal periodOd As Date, ByVal periodDo As Date, _
                          ByVal sourceName As String)
    If MinutOd(periodOd) > MinutOd(periodDo) Then
        Err.Raise DOP_ERR_BASE + 10, sourceName, _
                  "Pocetak perioda je posle kraja: " & Format$(periodOd, "dd.mm.yyyy hh:nn") & _
                  " > " & Format$(periodDo, "dd.mm.yyyy hh:nn") & "."
    End If
    If MinutOd(periodDo) > MinutOd(Now) Then
        Err.Raise DOP_ERR_BASE + 11, sourceName, _
                  "Kraj perioda je u buducnosti (" & Format$(periodDo, "dd.mm.yyyy hh:nn") & _
                  "). Doplata se obracunava za otkupe koji su vec nastali."
    End If
End Sub

Private Sub RequireDatumKnjizenja(ByVal datum As Date, ByVal periodDo As Date, _
                                  ByVal sourceName As String)
    If Int(CDbl(datum)) < Int(CDbl(periodDo)) Then
        Err.Raise DOP_ERR_BASE + 12, sourceName, _
                  "Datum knjizenja (" & Format$(datum, "dd.mm.yyyy") & ") je pre kraja " & _
                  "perioda (" & Format$(periodDo, "dd.mm.yyyy") & "). Zaduzenje ne sme da " & _
                  "stoji na kartici pre robe koju placa."
    End If
    If Int(CDbl(datum)) > Int(CDbl(Date)) Then
        Err.Raise DOP_ERR_BASE + 13, sourceName, _
                  "Datum knjizenja je u buducnosti: " & Format$(datum, "dd.mm.yyyy") & "."
    End If
End Sub

' Stope -> tri recnika:
'   cene        "VRSTA|SORTA|KLASA" -> din/kg
'   vrsteSorte  "VRSTA|SORTA"       -> True   (prefilter zaglavlja)
'   prikaz      "VRSTA|SORTA|KLASA" -> Array(vrsta, sorta, klasa) kako je uneto
' Kljuc je kljuc cenovnika (DOP-ODL-02): trim, bez razlike u velicini slova,
' prazna sorta = kultura bez sorte (ne "sve sorte"). Klasa I ili II, cena > 0,
' bez dupliranog kljuca.
Private Sub UcitajStope(ByVal stope As Collection, ByRef cene As Object, _
                        ByRef vrsteSorte As Object, ByRef prikaz As Object, _
                        ByVal sourceName As String)
    If stope Is Nothing Then
        Err.Raise DOP_ERR_BASE + 1, sourceName, "Nije zadata nijedna stopa doplate."
    End If
    If stope.count = 0 Then
        Err.Raise DOP_ERR_BASE + 1, sourceName, "Nije zadata nijedna stopa doplate."
    End If

    Set cene = CreateObject("Scripting.Dictionary")
    cene.CompareMode = vbTextCompare
    Set vrsteSorte = CreateObject("Scripting.Dictionary")
    vrsteSorte.CompareMode = vbTextCompare
    Set prikaz = CreateObject("Scripting.Dictionary")
    prikaz.CompareMode = vbTextCompare

    Dim i As Long, s As Object, k As Variant
    Dim vrsta As String, sorta As String, klasa As String, cena As Variant, kljuc As String

    For i = 1 To stope.count
        If Not IsObject(stope(i)) Then
            Err.Raise DOP_ERR_BASE + 2, sourceName, "Stopa " & CStr(i) & " nije Dictionary."
        End If
        Set s = stope(i)

        For Each k In s.keys
            If InStr(1, DOP_STOPA_KLJUCEVI, "|" & UCase$(CStr(k)) & "|") = 0 Then
                Err.Raise DOP_ERR_BASE + 3, sourceName, _
                          "Stopa " & CStr(i) & " ima nepoznat kljuc: " & CStr(k) & "."
            End If
        Next k
        If Not (s.Exists("VrstaVoca") And s.Exists("SortaVoca") And s.Exists("Klasa") _
                And s.Exists("Cena")) Then
            Err.Raise DOP_ERR_BASE + 4, sourceName, _
                      "Stopa " & CStr(i) & " mora imati VrstaVoca, SortaVoca, Klasa i Cena."
        End If

        vrsta = Trim$(NzToText(s("VrstaVoca")))
        sorta = Trim$(NzToText(s("SortaVoca")))
        klasa = Trim$(NzToText(s("Klasa")))
        cena = s("Cena")

        If Len(vrsta) = 0 Then
            Err.Raise DOP_ERR_BASE + 5, sourceName, "Stopa " & CStr(i) & ": vrsta je obavezna."
        End If
        modOtkup.RequireValidOtkupClass klasa, sourceName
        If Not IsNumeric(cena) Then
            Err.Raise DOP_ERR_BASE + 6, sourceName, _
                      "Stopa " & CStr(i) & ": iznos doplate nije broj."
        End If
        If CDbl(cena) <= 0 Then
            Err.Raise DOP_ERR_BASE + 7, sourceName, _
                      "Stopa " & CStr(i) & ": iznos doplate mora biti veci od nule (din/kg)."
        End If

        kljuc = KljucVrstaSorta(vrsta, sorta) & "|" & UCase$(klasa)
        If cene.Exists(kljuc) Then
            Err.Raise DOP_ERR_BASE + 8, sourceName, _
                      "Ista stopa je zadata dvaput: " & vrsta & " " & sorta & " " & klasa & "."
        End If
        cene.Add kljuc, CDbl(cena)
        prikaz.Add kljuc, Array(vrsta, sorta, klasa)
        vrsteSorte(KljucVrstaSorta(vrsta, sorta)) = True
    Next i
End Sub

' ============================================================
' POMOCNICI
' ============================================================

Private Function KljucVrstaSorta(ByVal vrsta As Variant, ByVal sorta As Variant) As String
    KljucVrstaSorta = UCase$(Trim$(NzToText(vrsta))) & "|" & UCase$(Trim$(NzToText(sorta)))
End Function

' Kljuc poretka plana: kooperant, stanica, vrsta|sorta|klasa, pa klasa I pre II.
' Chr$(1) razdvaja delove (binarno poredjenje, SortArrayPoKljucu).
' Poredak plana: kooperant, stanica, datum bloka, redosled upisa bloka (red u
' tblOtkup), klasa (I pre II) -- stavke istog bloka su uzastopne, pa pisac pravi
' dokument u jednom prolazu. Redosled upisa, a NE broj bloka kao tekst: "-10"
' bi tekstualno dosao pre "-9".
Private Function KljucPoretka(ByVal koop As String, ByVal stanica As String, _
                              ByVal datumBloka As Date, ByVal redBloka As Long, _
                              ByVal klasa As String) As String
    Dim rang As String
    If UCase$(Trim$(klasa)) = UCase$(KLASA_I) Then rang = "1" Else rang = "2"
    KljucPoretka = UCase$(koop) & Chr$(1) & UCase$(stanica) & Chr$(1) & _
                   Format$(datumBloka, "yyyymmdd") & Chr$(1) & Format$(redBloka, "0000000") & _
                   Chr$(1) & rang
End Function

' Da li kooperant ulazi u obuhvat obracuna (DOP-ODL-31): prazan obuhvat = svi.
' JEDNO pravilo za plan, pisca i signal osnovice -- inace bi obracun knjizio
' jedne blokove, a signal trazio retro unos medju drugima.
Private Function UObuhvatu(ByVal kooperantID As Variant, ByVal obuhvat As String) As Boolean
    If Len(obuhvat) = 0 Then
        UObuhvatu = True
    Else
        UObuhvatu = (StrComp(Trim$(NzToText(kooperantID)), obuhvat, vbTextCompare) = 0)
    End If
End Function

Private Sub RequireNacin(ByVal nacin As Long, ByVal sourceName As String)
    If nacin <> DOP_NACIN_PO_KG And nacin <> DOP_NACIN_KONACNA Then
        Err.Raise DOP_ERR_BASE + 34, sourceName, _
                  "Nepoznat nacin unosa doplate: " & CStr(nacin) & _
                  " (1 = doplata po kg, 2 = konacna cena po kg)."
    End If
End Sub

' Spisak kljuceva stopa koje NISU pokrivene (pokriveno Nothing = nijedna).
Private Function SpisakKljuceva(ByVal prikaz As Object, ByVal pokriveno As Object) As String
    Dim k As Variant, t As String, deo As Variant
    For Each k In prikaz.keys
        Dim pogodjen As Boolean
        pogodjen = False
        If Not pokriveno Is Nothing Then pogodjen = pokriveno.Exists(CStr(k))
        If Not pogodjen Then
            deo = prikaz(k)
            If Len(t) > 0 Then t = t & "; "
            t = t & Trim$(CStr(deo(0)) & " " & CStr(deo(1))) & " " & CStr(deo(2))
        End If
    Next k
    SpisakKljuceva = t
End Function

Private Function PrazanRed(ByVal tblName As String, ByVal sourceName As String) As Variant()
    Dim n As Long
    n = TabelaBrojKolona(tblName)
    If n <= 0 Then
        Err.Raise DOP_ERR_BASE + 45, sourceName, _
                  "Ne mogu da odredim broj kolona za " & tblName & "."
    End If
    Dim r() As Variant
    ReDim r(0 To n - 1)
    PrazanRed = r
End Function
