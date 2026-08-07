Attribute VB_Name = "modIzvestajTests"
Option Explicit

' ============================================================
' modIzvestajTests
'
' Dva nivoa:
'   RunIzvestajTests      - ASSERT suite (RF-06 / AUD-023): fiksira svaku
'                           ispravljenu brojku; regresija ponovo obara test.
'                           Radi nad cistim racunskim seam-ovima (bez tabela),
'                           pa je deterministican na svakoj instalaciji.
'   SmokeTest_modIzvestaj  - shape smoke nad ZIVIM podacima (ne tvrdi brojke).
'
' Output: Immediate Window / Ctrl+G
' ============================================================

Private m_izvFail As Long
Private m_izvPass As Long

' Deterministicki podaci za end-to-end testove. Datum je namerno van svakog
' realnog opsega, pa seed redovi ne mogu da se pomesaju sa produkcijskim.
Private Const IZVT_BROJ As String = "IZVT-1/150199"
Private Const IZVT_STANICA As String = "IZVT-OM"
Private Const IZVT_STANICA2 As String = "IZVT-OM2"   ' izolovana za Klasa I+II test
Private Const IZVT_VOZAC_A As String = "IZVT-VZ-A"
Private Const IZVT_VOZAC_B As String = "IZVT-VZ-B"
Private Const IZVT_KUPAC_A As String = "IZVT-KP-A"
Private Const IZVT_KUPAC_B As String = "IZVT-KP-B"

Private Function IZVT_DATUM() As Date
    IZVT_DATUM = DateSerial(1999, 1, 15)
End Function

' ============================================================
' RF-06 ASSERT SUITE
' Svaki blok odgovara jednoj stavki iz AUD-023 / FM-0028.
'
' TVRD GATE: suite PODIZE gresku kad ijedan assert padne ili kad interno pukne
' (ista konvencija kao RunSheetsJsonParserTests / RunMasterSyncSmokeSuite posle
' RF-14 -- inace automatizovani pozivalac vidi "makro zavrsen" i prijavi PASS
' iako su assertioni pali).
' ============================================================
Public Sub RunIzvestajTests()
    Dim tx As clsTransaction
    On Error GoTo EH

    m_izvFail = 0
    m_izvPass = 0

    Debug.Print String(70, "=")
    Debug.Print "RunIzvestajTests START (RF-06 / AUD-023)"
    Debug.Print String(70, "-")

    ' --- Cisti racunski seam-ovi (bez tabela) ---
    T_IzplateStanicaAtribucija
    T_KarticaPocetnoStanje
    T_KarticaAmbPocetnoStanje
    T_NemaPrijemaOznaka
    T_PrijemVlasnikRazresenje
    T_UplataSrazmernoPoVrsti
    T_DispatchNepodrzanTip
    T_EntitetKod
    T_TabMatrica
    T_ReversKljucPoTipu

    ' --- End-to-end nad tabelama (u transakciji, UVEK rollback) ---
    ' Seam testovi ne mogu da uhvate gresku u samom table-join-u, a upravo je
    ' join po `BrojZbirne` bio nosilac pogresnih brojki.
    If GetTable(TBL_ZBIRNA) Is Nothing Or GetTable(TBL_PRIJEMNICA) Is Nothing _
       Or GetTable(TBL_OTPREMNICA) Is Nothing Or GetTable(TBL_AMBALAZA) Is Nothing Then
        Debug.Print "  SKIP | end-to-end: nema tabela zbirna/prijemnica/otpremnica/ambalaza"
    Else
        Set tx = New clsTransaction
        tx.BeginTx
        tx.AddTableSnapshot TBL_ZBIRNA
        tx.AddTableSnapshot TBL_PRIJEMNICA
        tx.AddTableSnapshot TBL_OTPREMNICA
        tx.AddTableSnapshot TBL_OTKUP
        tx.AddTableSnapshot TBL_AMBALAZA

        T_E2E_ManjakDvaVlasnikaIstiBroj
        T_E2E_RobaOMDvaVlasnikaIstiBroj
        T_E2E_KlasaIiIINeMesajuPrijem
        T_E2E_ProsecnaCenaZbirniKupac
        T_E2E_AmbPregledRazdvajaTipDokumenta

        tx.RollbackTx
        Set tx = Nothing
    End If

    Debug.Print String(70, "-")
    If m_izvFail = 0 Then
        Debug.Print "RunIzvestajTests OK  | " & m_izvPass & " provera proslo"
        Debug.Print String(70, "=")
        Exit Sub
    End If

    Debug.Print "RunIzvestajTests FAIL | " & m_izvFail & " od " & _
                (m_izvFail + m_izvPass) & " provera palo"
    Debug.Print String(70, "=")
    On Error GoTo 0
    Err.Raise vbObjectError + 7610, "modIzvestajTests.RunIzvestajTests", _
              "RunIzvestajTests: " & m_izvFail & " od " & (m_izvFail + m_izvPass) & _
              " provera palo (detalji u Immediate Window)."
    Exit Sub

EH:
    ' errNum/errDesc/errSrc -- ista imena kao u ostatku projekta (modStorno
    ' LogAndReraise, modAgrohemija, modBankaImport). NE "eNum": VBA je
    ' case-insensitive, pa se `eNum` poklapa sa rezervisanom reci `Enum`.
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number: errDesc = Err.description: errSrc = Err.SOURCE

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    Debug.Print String(70, "!")
    Debug.Print "RunIzvestajTests ERROR | " & errNum & " | " & errSrc & " | " & errDesc
    Debug.Print String(70, "!")
    Err.Raise errNum, errSrc, errDesc
End Sub

' --- (a) FM-0028 #3/#9: isplata pripada stanici po OMID-u REDA ---
Private Sub T_IzplateStanicaAtribucija()
    Const S As String = "Isplata->stanica: "

    ' Red NOSI OMID -> odlucuje OMID reda, ne maticna stanica kooperanta.
    IzvChk NovacRedPripadaStanici("OM-1", "OM-2", "OM-1") = True, _
           S & "OMID reda = trazena stanica -> pripada"
    IzvChk NovacRedPripadaStanici("OM-1", "OM-2", "OM-2") = False, _
           S & "OMID reda != trazena stanica -> NE pripada (bez preliva na maticnu)"

    ' Red BEZ OMID-a (nepoznata istorijska stanica) -> maticna stanica kooperanta.
    IzvChk NovacRedPripadaStanici("", "OM-2", "OM-2") = True, _
           S & "bez OMID-a -> maticna stanica kooperanta"
    IzvChk NovacRedPripadaStanici("", "OM-2", "OM-1") = False, _
           S & "bez OMID-a i druga maticna -> NE pripada"

    ' Razmaci u ID-u ne smeju da razbiju poredjenje.
    IzvChk NovacRedPripadaStanici(" OM-1 ", "", "OM-1") = True, _
           S & "trim na OMID-u reda"
End Sub

' --- (b) FM-0028 #1: kartica krece od pocetnog stanja, ne od nule ---
Private Sub T_KarticaPocetnoStanje()
    Const S As String = "Kartica pocetno stanje: "

    ' 2 period-reda: zaduzenje 1000, razduzenje 400; amb delta +5 pa -2.
    Dim arr(1 To 2, 1 To 8) As Variant
    arr(1, 1) = DateSerial(2026, 5, 10): arr(1, 2) = "OTK-1": arr(1, 3) = ""
    arr(1, 4) = "Otkup": arr(1, 5) = 1000#: arr(1, 6) = 0#: arr(1, 7) = 5#: arr(1, 8) = "OTK|1"
    arr(2, 1) = DateSerial(2026, 5, 12): arr(2, 2) = "NOV-1": arr(2, 3) = ""
    arr(2, 4) = "Kes": arr(2, 5) = 0#: arr(2, 6) = 400#: arr(2, 7) = -2#: arr(2, 8) = "NOV"

    ' --- sa pocetnim stanjem (dug 300, ambalaza 7 gajbi) ---
    Dim r As Variant
    r = KarticaRezultatSaPocetnim(arr, 300#, 7#)

    IzvChk IsArray(r), S & "vraca niz"
    IzvChkEq UBound(r, 1), 4, S & "1 pocetni + 2 prometna + UKUPNO = 4 reda"
    IzvChkEqText CStr(r(1, 4)), IZV_POCETNO_STANJE, S & "prvi red je pocetno stanje"
    IzvChkEqD CDbl(r(1, 7)), 300#, S & "pocetni saldo u koloni Saldo"
    IzvChkEqD CDbl(r(1, 8)), 7#, S & "pocetni saldo ambalaze"
    IzvChk IsEmpty(r(1, 5)) And IsEmpty(r(1, 6)), S & "pocetni red ne ulazi u promet"

    ' Running saldo krece OD pocetnog: 300 + 1000 = 1300, pa - 400 = 900.
    IzvChkEqD CDbl(r(2, 7)), 1300#, S & "running saldo posle 1. reda = 300 + 1000"
    IzvChkEqD CDbl(r(3, 7)), 900#, S & "running saldo posle 2. reda = 1300 - 400"
    IzvChkEqD CDbl(r(2, 8)), 12#, S & "running amb posle 1. reda = 7 + 5"
    IzvChkEqD CDbl(r(3, 8)), 10#, S & "running amb posle 2. reda = 12 - 2"

    ' UKUPNO: kolone 5/6 = PROMET perioda, kolona 7 = ZAVRSNI saldo.
    IzvChkEqText CStr(r(4, 4)), "UKUPNO", S & "poslednji red je UKUPNO"
    IzvChkEqD CDbl(r(4, 5)), 1000#, S & "UKUPNO zaduzenje = promet perioda"
    IzvChkEqD CDbl(r(4, 6)), 400#, S & "UKUPNO razduzenje = promet perioda"
    IzvChkEqD CDbl(r(4, 7)), 900#, S & "UKUPNO saldo = pocetno + promet (NE promet sam)"
    IzvChkEqD CDbl(r(4, 8)), 10#, S & "UKUPNO saldo ambalaze = zavrsno stanje"

    ' --- bez pocetnog stanja: stari oblik (nema dodatnog reda) ---
    r = KarticaRezultatSaPocetnim(arr, 0#, 0#)
    IzvChkEq UBound(r, 1), 3, S & "bez pocetnog stanja nema dodatnog reda"
    IzvChkEqD CDbl(r(3, 7)), 600#, S & "bez pocetnog: UKUPNO saldo = 1000 - 400"

    ' --- samo pocetno stanje, bez prometa: kartica NIJE prazna ---
    r = KarticaRezultatSaPocetnim(Empty, 250#, 0#)
    IzvChk IsArray(r), S & "samo pocetno stanje i dalje daje karticu"
    IzvChkEqD CDbl(r(1, 7)), 250#, S & "pocetni saldo bez prometa"
    IzvChkEqD CDbl(r(2, 7)), 250#, S & "UKUPNO = pocetni saldo"

    ' --- ni prometa ni pocetnog stanja -> Empty (kao i pre) ---
    IzvChk IsEmpty(KarticaRezultatSaPocetnim(Empty, 0#, 0#)), _
           S & "bez ijednog podatka ostaje Empty"
End Sub

' --- (b) isto za "Pregled ambalaze" ---
Private Sub T_KarticaAmbPocetnoStanje()
    Const S As String = "Kartica ambalaze pocetno stanje: "

    Dim arr(1 To 2, 1 To 5) As Variant
    arr(1, 1) = DateSerial(2026, 5, 10): arr(1, 2) = "D-1": arr(1, 3) = "Gajbica"
    arr(1, 4) = 10#: arr(1, 5) = 0#
    arr(2, 1) = DateSerial(2026, 5, 11): arr(2, 2) = "D-2": arr(2, 3) = "Gajbica"
    arr(2, 4) = 0#: arr(2, 5) = 4#

    Dim r As Variant
    r = KarticaAmbRezultatSaPocetnim(arr, 6#)

    IzvChkEq UBound(r, 1), 4, S & "1 pocetni + 2 prometna + UKUPNO"
    IzvChkEqText CStr(r(1, 3)), IZV_POCETNO_STANJE, S & "prvi red je pocetno stanje"
    IzvChkEqD CDbl(r(1, 6)), 6#, S & "pocetno stanje gajbi"
    IzvChkEqD CDbl(r(2, 6)), 16#, S & "running saldo = 6 + 10"
    IzvChkEqD CDbl(r(3, 6)), 12#, S & "running saldo = 16 - 4"
    IzvChkEqD CDbl(r(4, 4)), 10#, S & "UKUPNO ulaz = promet perioda"
    IzvChkEqD CDbl(r(4, 5)), 4#, S & "UKUPNO izlaz = promet perioda"
    IzvChkEqD CDbl(r(4, 6)), 12#, S & "UKUPNO saldo = zavrsno stanje (NE 10-4)"

    r = KarticaAmbRezultatSaPocetnim(arr, 0#)
    IzvChkEq UBound(r, 1), 3, S & "bez pocetnog stanja nema dodatnog reda"
    IzvChk IsEmpty(KarticaAmbRezultatSaPocetnim(Empty, 0#)), S & "prazno ostaje Empty"
End Sub

' --- (c) FM-0028 #5: nedostajuca prijemnica = oznaka, ne 0% / 100% ---
Private Sub T_NemaPrijemaOznaka()
    Const S As String = "Nema prijema: "

    ' Sa prijemom: normalan racun manjka.
    Dim m As Variant
    m = ManjakStavka(1000#, 900#, True)
    IzvChkEqD CDbl(m(0)), 900#, S & "prijem kg"
    IzvChkEqD CDbl(m(1)), 100#, S & "manjak kg = 1000 - 900"
    IzvChkEqD CDbl(m(2)), 10#, S & "manjak % = 10"
    IzvChkEqText CStr(m(3)), "", S & "nema oznake kad prijem postoji"

    ' Bez prijema: SVE brojke prazne + oznaka. Pre RF-06 je isti slucaj bio
    ' 0 kg / 0,00% u RobaOM i 100% u Manjak izvestaju.
    m = ManjakStavka(1000#, 0#, False)
    IzvChk IsEmpty(m(0)), S & "prijem kg je prazan (ne 0)"
    IzvChk IsEmpty(m(1)), S & "manjak kg je prazan (ne 1000)"
    IzvChk IsEmpty(m(2)), S & "manjak % je prazan (ne 0 i ne 100)"
    IzvChkEqText CStr(m(3)), IZV_NEMA_PRIJEMA, S & "oznaka 'nema prijema'"

    ' Nulta osnovica ne sme da deli nulom.
    m = ManjakStavka(0#, 0#, True)
    IzvChkEqD CDbl(m(2)), 0#, S & "osnovica 0 -> procenat 0, bez greske"
End Sub

' --- (c2) AUD-052 posledica: BrojZbirne NIJE identitet ---
Private Sub T_PrijemVlasnikRazresenje()
    Const S As String = "Prijem/vlasnik: "

    ' Broj ima JEDNOG vlasnika -> join po broju je siguran (stara, brza putanja).
    Dim r As Variant
    r = PrijemZaZbirnu(1, True, 0, 2, 900#)
    IzvChk CBool(r(0)) = True, S & "jedan vlasnik -> prijem vazi"
    IzvChkEqD CDbl(r(1)), 900#, S & "jedan vlasnik -> kg po broju"
    IzvChkEqText CStr(r(2)), "", S & "jedan vlasnik -> bez oznake"

    ' Nema nijedne zbirne / nema prijemnice -> nema prijema (ne 100% manjka).
    r = PrijemZaZbirnu(1, True, 0, 0, 0#)
    IzvChk CBool(r(0)) = False, S & "bez prijemnice -> nema prijema"
    IzvChkEqText CStr(r(2)), IZV_NEMA_PRIJEMA, S & "oznaka 'nema prijema'"

    ' Broj dele DVA vlasnika, ali je vlasnik reda razresen i sve prijemnice
    ' nose kompletnog vlasnika -> racuna se owner-scoped prijem.
    r = PrijemZaZbirnu(2, True, 0, 1, 1500#)
    IzvChk CBool(r(0)) = True, S & "dva vlasnika + razresen -> prijem vazi"
    IzvChkEqD CDbl(r(1)), 1500#, S & "koristi se owner-scoped kg, ne zbir broja"

    ' Fail-closed 1: vlasnik reda se ne moze razresiti.
    r = PrijemZaZbirnu(2, False, 0, 1, 1500#)
    IzvChk CBool(r(0)) = False, S & "dva vlasnika + nerazresen -> bez brojke"
    IzvChkEqText CStr(r(2)), IZV_VLASNIK_NEJASAN, S & "oznaka 'nejasan vlasnik'"

    ' Fail-closed 2: postoji prijemnica koja se ne moze pripisati nijednoj zbirnoj.
    r = PrijemZaZbirnu(2, True, 1, 1, 1500#)
    IzvChk CBool(r(0)) = False, S & "nepripisiva prijemnica -> bez brojke"
    IzvChkEqText CStr(r(2)), IZV_VLASNIK_NEJASAN, S & "nepripisiva -> 'nejasan vlasnik'"

    ' Nepripisiva prijemnica NE blokira kad broj ima jednog vlasnika
    ' (tada je pripadnost dokazana samim brojem -- starije instalacije).
    r = PrijemZaZbirnu(1, True, 3, 3, 900#)
    IzvChk CBool(r(0)) = True, S & "jedan vlasnik: prazan vlasnik na prijemnici ne smeta"

    ' Vlasnik kljuc: ista definicija kao modStorno.RequireJedanVlasnikPoBroju.
    IzvChkEqText ZbirnaVlasnikKljuc("1/010126", "VZ-1", "KP-1"), "1/010126|VZ-1|KP-1", _
                 S & "vlasnik kljuc = broj|vozac|kupac"
    IzvChkEqText ZbirnaVlasnikKljuc(" 1/010126 ", " VZ-1", "KP-1 "), "1/010126|VZ-1|KP-1", _
                 S & "vlasnik kljuc trimuje delove"
End Sub

' --- (d) FM-0028 #6: uplata po fakturi se deli SRAZMERNO stavkama ---
Private Sub T_UplataSrazmernoPoVrsti()
    Const S As String = "Uplata po vrsti: "

    Dim udeli As Object
    Set udeli = CreateObject("Scripting.Dictionary")
    udeli.Add "Malina", 30000#
    udeli.Add "Kupina", 70000#

    Dim p As Object
    Set p = RaspodeliPoUdelima(50000#, udeli)

    IzvChkEq p.count, 2, S & "obe vrste dobijaju deo (ne samo prva stavka)"
    IzvChkEqD CDbl(p("Malina")), 15000#, S & "Malina = 30% od 50.000"
    IzvChkEqD CDbl(p("Kupina")), 35000#, S & "Kupina = 70% od 50.000"
    IzvChkEqD CDbl(p("Malina")) + CDbl(p("Kupina")), 50000#, S & "zbir podele = ceo iznos"

    ' Jedna vrsta -> cela uplata na nju (bez promene ponasanja).
    Dim jedna As Object
    Set jedna = CreateObject("Scripting.Dictionary")
    jedna.Add "Malina", 12345#
    Set p = RaspodeliPoUdelima(999#, jedna)
    IzvChkEqD CDbl(p("Malina")), 999#, S & "jedna vrsta nosi ceo iznos"

    ' Bez upotrebljivih tezina -> "(Nepoznato)", bez deljenja nulom.
    Dim prazno As Object
    Set prazno = CreateObject("Scripting.Dictionary")
    prazno.Add "Malina", 0#
    Set p = RaspodeliPoUdelima(500#, prazno)
    IzvChkEqD CDbl(p("(Nepoznato)")), 500#, S & "nulte tezine -> (Nepoznato)"

    Set p = RaspodeliPoUdelima(500#, Nothing)
    IzvChkEqD CDbl(p("(Nepoznato)")), 500#, S & "nepoznata faktura -> (Nepoznato)"

    ' Zaokruzivanje: 100 / 3 jednake stavke. Delovi moraju biti VEC zaokruzeni na
    ' 2 decimale (33,33 / 33,33 / 33,34), inace forma prikaze 3 x 33,33 = 99,99
    ' uz UKUPNO 100,00 -- vidljiva razlika od jednog centa.
    Dim tri As Object
    Set tri = CreateObject("Scripting.Dictionary")
    tri.Add "A", 1#: tri.Add "B", 1#: tri.Add "C", 1#
    ' NAPOMENA: ovde se poredi na NIVOU CENTA (IzvChkEqC), ne sa tolerancijom
    ' 0,01 -- nezaokruzeno 33,3333 bi proslo tolerantnu proveru i propustilo
    ' bas onu gresku koju test treba da fiksira.
    ' Metod je largest-remainder: visak para ide na NAJVECI ostatak, pa kod tri
    ' jednaka udela dodatnu paru dobija PRVI kljuc (svi ostaci su isti).
    Set p = RaspodeliPoUdelima(100#, tri)
    IzvChkEqC CDbl(p("A")), 33.34, S & "prvi deo nosi visak pare -> 33,34"
    IzvChkEqC CDbl(p("B")), 33.33, S & "drugi deo 33,33"
    IzvChkEqC CDbl(p("C")), 33.33, S & "treci deo 33,33"
    IzvChkEqC CDbl(p("A")) + CDbl(p("B")) + CDbl(p("C")), 100#, _
              S & "zbir ZAOKRUZENIH delova = 100,00 (ne 99,99)"
    IzvChk IsRoundedTo2(CDbl(p("A"))) And IsRoundedTo2(CDbl(p("B"))) And IsRoundedTo2(CDbl(p("C"))), _
           S & "nijedan deo nema vise od 2 decimale"

    ' Sitan iznos na vise vrsta: raniji oblik (poslednji nosi ostatak) je prva
    ' cetiri dela zaokruzivao navise preko cilja, pa je POSLEDNJA vrsta dobijala
    ' -0,01 -- negativan cent u izvestaju salda kupca.
    Set p = RaspodeliPoUdelima(0.03, IzvUdeliJednaki(5))
    Dim minDeo As Double: minDeo = IzvMinVrednost(p)
    IzvChk minDeo >= 0, S & "0,03 na 5 vrsta: nijedan deo nije negativan"
    IzvChkEqC IzvZbirVrednosti(p), 0.03, S & "0,03 na 5 vrsta: zbir = 0,03"

    ' Invarijanta (zbir == iznos I nijedan deo < 0) na vise oblika ulaza.
    ChkRaspodelaInvarijanta 0.01, IzvUdeliJednaki(2), S & "0,01 na 2"
    ChkRaspodelaInvarijanta 0.03, IzvUdeliJednaki(5), S & "0,03 na 5"
    ChkRaspodelaInvarijanta 0.07, IzvUdeliJednaki(3), S & "0,07 na 3"
    ChkRaspodelaInvarijanta 100#, IzvUdeliJednaki(7), S & "100 na 7"
    ChkRaspodelaInvarijanta 1234.56, IzvUdeliJednaki(3), S & "1234,56 na 3"
    ChkRaspodelaInvarijanta 999999.99, IzvUdeliJednaki(11), S & "999.999,99 na 11"

    ' Isto na iznosu koji se ne deli lepo na dve vrste.
    Dim dve As Object
    Set dve = CreateObject("Scripting.Dictionary")
    dve.Add "A", 1#: dve.Add "B", 2#
    Set p = RaspodeliPoUdelima(10#, dve)
    IzvChkEqC CDbl(p("A")), 3.33, S & "1/3 od 10 = 3,33"
    IzvChkEqC CDbl(p("B")), 6.67, S & "ostatak = 6,67"
    IzvChkEqC CDbl(p("A")) + CDbl(p("B")), 10#, S & "zbir = 10,00"

    ' Zaokruzivanje navise na .xx5 se NE testira na knife-edge vrednosti
    ' (2,345 u Double-u nije tacno 2,345, pa bi test bio nedeterministican).
    IzvChkEqC ZaokruziNovac(2.344), 2.34, S & "2,344 -> 2,34"
    IzvChkEqC ZaokruziNovac(2.346), 2.35, S & "2,346 -> 2,35"
    IzvChkEqC ZaokruziNovac(-2.346), -2.35, S & "negativan iznos je simetrican"
    IzvChkEqC ZaokruziNovac(0#), 0#, S & "nula ostaje nula"
End Sub

' Da li vrednost stane u 2 decimale (bez repa tipa 33,3333).
Private Function IsRoundedTo2(ByVal v As Double) As Boolean
    IsRoundedTo2 = (Abs(v * 100 - Int(v * 100 + 0.5)) < 0.000001)
End Function

' Dve invarijante raspodele: zbir delova == zaokruzen iznos, i nijedan deo < 0.
' Clamp na nulu bi zadovoljio drugu a razbio prvu -- zato se proveravaju zajedno.
Private Sub ChkRaspodelaInvarijanta(ByVal iznos As Double, ByVal udeli As Object, _
                                    ByVal labela As String)
    Dim p As Object
    Set p = RaspodeliPoUdelima(iznos, udeli)

    IzvChk Abs(IzvZbirVrednosti(p) - ZaokruziNovac(iznos)) < 0.005, _
           labela & ": zbir delova = iznos"
    IzvChk IzvMinVrednost(p) >= 0, labela & ": nijedan deo nije negativan"
End Sub

Private Function IzvZbirVrednosti(ByVal d As Object) As Double
    Dim k As Variant
    For Each k In d.keys
        IzvZbirVrednosti = IzvZbirVrednosti + CDbl(d(k))
    Next k
End Function

Private Function IzvMinVrednost(ByVal d As Object) As Double
    Dim prvi As Boolean: prvi = True
    Dim k As Variant
    For Each k In d.keys
        If prvi Then
            IzvMinVrednost = CDbl(d(k))
            prvi = False
        ElseIf CDbl(d(k)) < IzvMinVrednost Then
            IzvMinVrednost = CDbl(d(k))
        End If
    Next k
End Function

' N jednakih udela ("V1".."Vn") -- za provere zaokruzivanja na sitnim iznosima.
Private Function IzvUdeliJednaki(ByVal n As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To n
        d.Add "V" & i, 1#
    Next i
    Set IzvUdeliJednaki = d
End Function

' ============================================================
' END-TO-END (nad tabelama, u transakciji pozivaoca -- uvek rollback)
' Seam testovi ne mogu da uhvate gresku u samom join-u zbirna<->prijemnica,
' a upravo je join po `BrojZbirne` (koji NIJE identitet -- RF-05/AUD-052)
' pravio pogresnu prijemnu kolicinu, manjak i procenat.
' ============================================================

' Dve AKTIVNE zbirne istog broja, dva kupca, razlicite prijemne kolicine.
Private Sub T_E2E_ManjakDvaVlasnikaIstiBroj()
    Const S As String = "E2E Manjak (dva vlasnika, isti broj): "
    On Error GoTo EH

    SeedDveZbirneIstogBroja

    Dim d As Date: d = IZVT_DATUM
    Dim r As Variant
    r = ReportManjak("", "", d, d)

    IzvChk IsArray(r), S & "izvestaj vraca redove"
    If Not IsArray(r) Then Exit Sub

    ' Redovi se razlikuju po zbirnoj kilazi (broj im je isti -- to je i poenta).
    Dim rowA As Long: rowA = IzvFindRowByNum(r, 2, 1000#)
    Dim rowB As Long: rowB = IzvFindRowByNum(r, 2, 2000#)

    IzvChk rowA > 0, S & "zbirna A (1000 kg) je zaseban red"
    IzvChk rowB > 0, S & "zbirna B (2000 kg) je zaseban red"
    If rowA = 0 Or rowB = 0 Then Exit Sub

    ' Pre RF-06 bi OBA reda videla 900 + 1500 = 2400 primljenih kg.
    IzvChkEqD CDbl(r(rowA, 3)), 900#, S & "A prima SVOJIH 900 kg (ne 2400)"
    IzvChkEqD CDbl(r(rowB, 3)), 1500#, S & "B prima SVOJIH 1500 kg (ne 2400)"
    IzvChkEqD CDbl(r(rowA, 4)), 100#, S & "A manjak = 1000 - 900"
    IzvChkEqD CDbl(r(rowB, 4)), 500#, S & "B manjak = 2000 - 1500"
    IzvChkEqD CDbl(r(rowA, 5)), 10#, S & "A manjak % = 10"
    IzvChkEqD CDbl(r(rowB, 5)), 25#, S & "B manjak % = 25"

    ' Filter po kupcu mora da izoluje samo svoju zbirnu.
    r = ReportManjak("Kupac", IZVT_KUPAC_A, d, d)
    IzvChk IsArray(r), S & "filter po kupcu vraca redove"
    If IsArray(r) Then
        IzvChkEq UBound(r, 1), 2, S & "kupac A: 1 red + UKUPNO"
        IzvChkEqD CDbl(r(1, 2)), 1000#, S & "kupac A vidi samo svoju zbirnu"
        IzvChkEqD CDbl(r(1, 3)), 900#, S & "kupac A vidi samo svoj prijem"
    End If
    Exit Sub
EH:
    IzvChk False, S & "neocekivana greska: " & Err.description
End Sub

' Ista postavka iz ugla "Otkupljena roba (OM)" -- razresenje ide preko vozaca
' otpremnice (otpremnica nema KupacID).
Private Sub T_E2E_RobaOMDvaVlasnikaIstiBroj()
    Const S As String = "E2E RobaOM (dva vlasnika, isti broj): "
    On Error GoTo EH

    ' Zbirne/prijemnice je vec zasejao prethodni test (ista transakcija).
    IzvSeed TBL_OTPREMNICA, _
        Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_DATUM, COL_OTP_STANICA, COL_OTP_VOZAC, _
              COL_OTP_BROJ_ZBIRNE, COL_OTP_VRSTA, COL_OTP_KLASA, COL_OTP_KOLICINA), _
        Array("IZVT-OTP-A", "IZVT-OTP-A", IZVT_DATUM, IZVT_STANICA, IZVT_VOZAC_A, _
              IZVT_BROJ, "Malina", "I", 1000#)

    IzvSeed TBL_OTPREMNICA, _
        Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_DATUM, COL_OTP_STANICA, COL_OTP_VOZAC, _
              COL_OTP_BROJ_ZBIRNE, COL_OTP_VRSTA, COL_OTP_KLASA, COL_OTP_KOLICINA), _
        Array("IZVT-OTP-B", "IZVT-OTP-B", IZVT_DATUM, IZVT_STANICA, IZVT_VOZAC_B, _
              IZVT_BROJ, "Malina", "I", 2000#)

    Dim d As Date: d = IZVT_DATUM
    Dim r As Variant
    r = ReportOtkupRoba("OM", IZVT_STANICA, d, d)

    IzvChk IsArray(r), S & "izvestaj vraca redove"
    If Not IsArray(r) Then Exit Sub

    Dim rowA As Long: rowA = IzvFindRowByText(r, 2, "IZVT-OTP-A")
    Dim rowB As Long: rowB = IzvFindRowByText(r, 2, "IZVT-OTP-B")
    IzvChk rowA > 0 And rowB > 0, S & "obe otpremnice su u izvestaju"
    If rowA = 0 Or rowB = 0 Then Exit Sub

    ' Otpremnica kg = zbirna kg te iste zbirne, pa je rezultat isti i u malina
    ' i u standardnom modu (srazmera = 1). Pre RF-06: osnovica 3000 i prijem
    ' 2400 (obe zbirne skupa) -> A bi dobila 800 kg i manjak 200.
    IzvChkEqD CDbl(r(rowA, 9)), 900#, S & "A prijemnica kg = 900 (ne 800)"
    IzvChkEqD CDbl(r(rowB, 9)), 1500#, S & "B prijemnica kg = 1500"
    IzvChkEqD CDbl(r(rowA, 10)), 100#, S & "A manjak = 100 (ne 200)"
    IzvChkEqD CDbl(r(rowB, 10)), 500#, S & "B manjak = 500"
    Exit Sub
EH:
    IzvChk False, S & "neocekivana greska: " & Err.description
End Sub

' Klasa I i II ISTOG dokumenta: isti broj, isti vozac, isti kupac, ali zasebna
' otpremnica/zbirna/prijemnica po klasi (tako ih pravi auto-lanac hladnjace).
' Bez Klase u kljucu prijem obe klase se sabere pa dodeli SVAKOJ klasi -- u
' malina modu UKUPNO prijem postaje dvostruk.
Private Sub T_E2E_KlasaIiIINeMesajuPrijem()
    Const S As String = "E2E Klasa I+II (isti dokument): "
    On Error GoTo EH

    Dim brDok As String: brDok = "IZVT-2/150199"
    Dim voz As String: voz = "IZVT-VZ-K"
    Dim kup As String: kup = "IZVT-KP-K"

    ' Klasa I: zbirna/otpremnica 1000 kg, prijem 900. Klasa II: 200 kg, prijem 150.
    IzvSeed TBL_ZBIRNA, _
        Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_DATUM, COL_ZBR_VOZAC, COL_ZBR_KUPAC, _
              COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB, COL_ZBR_KLASA), _
        Array("IZVT-ZBR-K1", brDok, IZVT_DATUM, voz, kup, 1000#, 100, KLASA_I)
    IzvSeed TBL_ZBIRNA, _
        Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_DATUM, COL_ZBR_VOZAC, COL_ZBR_KUPAC, _
              COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB, COL_ZBR_KLASA), _
        Array("IZVT-ZBR-K2", brDok, IZVT_DATUM, voz, kup, 200#, 20, KLASA_II)

    ' Klasa I i II dele BROJ prijemnice (kao u produkciji), ali su zasebni redovi.
    IzvSeed TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_BROJ_ZBIRNE, _
              COL_PRJ_VOZAC, COL_PRJ_KUPAC, COL_PRJ_KOLICINA, COL_PRJ_KLASA), _
        Array("IZVT-PRJ-K1", "IZVT-PRJ-K", IZVT_DATUM, brDok, voz, kup, 900#, KLASA_I)
    IzvSeed TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_BROJ_ZBIRNE, _
              COL_PRJ_VOZAC, COL_PRJ_KUPAC, COL_PRJ_KOLICINA, COL_PRJ_KLASA), _
        Array("IZVT-PRJ-K2", "IZVT-PRJ-K", IZVT_DATUM, brDok, voz, kup, 150#, KLASA_II)

    IzvSeed TBL_OTPREMNICA, _
        Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_DATUM, COL_OTP_STANICA, COL_OTP_VOZAC, _
              COL_OTP_BROJ_ZBIRNE, COL_OTP_VRSTA, COL_OTP_KLASA, COL_OTP_KOLICINA), _
        Array("IZVT-OTP-K1", "IZVT-OTP-K1", IZVT_DATUM, IZVT_STANICA2, voz, _
              brDok, "Malina", KLASA_I, 1000#)
    IzvSeed TBL_OTPREMNICA, _
        Array(COL_OTP_ID, COL_OTP_BROJ, COL_OTP_DATUM, COL_OTP_STANICA, COL_OTP_VOZAC, _
              COL_OTP_BROJ_ZBIRNE, COL_OTP_VRSTA, COL_OTP_KLASA, COL_OTP_KOLICINA), _
        Array("IZVT-OTP-K2", "IZVT-OTP-K2", IZVT_DATUM, IZVT_STANICA2, voz, _
              brDok, "Malina", KLASA_II, 200#)

    Dim d As Date: d = IZVT_DATUM

    ' --- Otkupljena roba (OM): po klasi, bez mesanja ---
    Dim r As Variant
    r = ReportOtkupRoba("OM", IZVT_STANICA2, d, d)
    IzvChk IsArray(r), S & "RobaOM vraca redove"
    If Not IsArray(r) Then Exit Sub

    Dim r1 As Long: r1 = IzvFindRowByText(r, 2, "IZVT-OTP-K1")
    Dim r2 As Long: r2 = IzvFindRowByText(r, 2, "IZVT-OTP-K2")
    IzvChk r1 > 0 And r2 > 0, S & "obe klase su zasebni redovi"
    If r1 = 0 Or r2 = 0 Then Exit Sub

    ' Pre fix-a (malina mod): obe klase bi dobile 1050 kg prijema.
    IzvChkEqD CDbl(r(r1, 9)), 900#, S & "Klasa I prima 900 (ne 1050)"
    IzvChkEqD CDbl(r(r2, 9)), 150#, S & "Klasa II prima 150 (ne 1050)"
    IzvChkEqD CDbl(r(r1, 10)), 100#, S & "Klasa I manjak = 1000 - 900"
    IzvChkEqD CDbl(r(r2, 10)), 50#, S & "Klasa II manjak = 200 - 150"

    ' UKUPNO prijem mora biti 1050, ne 2100.
    Dim uk As Long: uk = UBound(r, 1)
    IzvChkEqD CDbl(r(uk, 9)), 1050#, S & "UKUPNO prijem = 1050 (ne 2x1050)"
    IzvChkEqD CDbl(r(uk, 10)), 150#, S & "UKUPNO manjak = 100 + 50"

    ' --- Manjak: ceo dokument u JEDNOM redu, prijem sabran po klasama ---
    Dim m As Variant
    m = ReportManjak("Kupac", kup, d, d)
    IzvChk IsArray(m), S & "Manjak vraca redove"
    If Not IsArray(m) Then Exit Sub

    IzvChkEq UBound(m, 1), 2, S & "dokument je JEDAN red + UKUPNO"
    IzvChkEqD CDbl(m(1, 2)), 1200#, S & "zbirna kg = 1000 + 200 (ceo dokument)"
    IzvChkEqD CDbl(m(1, 3)), 1050#, S & "prijem = 900 + 150 (ne 2x1050)"
    IzvChkEqD CDbl(m(1, 4)), 150#, S & "manjak = 1200 - 1050"
    Exit Sub
EH:
    IzvChk False, S & "neocekivana greska: " & Err.description
End Sub

' Dve aktivne zbirne sa ISTIM BrojZbirne, razliciti vozac+kupac, svaka sa
' svojom prijemnicom. Poziva se unutar transakcije pozivaoca.
Private Sub SeedDveZbirneIstogBroja()
    IzvSeed TBL_ZBIRNA, _
        Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_DATUM, COL_ZBR_VOZAC, COL_ZBR_KUPAC, _
              COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB), _
        Array("IZVT-ZBR-A", IZVT_BROJ, IZVT_DATUM, IZVT_VOZAC_A, IZVT_KUPAC_A, 1000#, 100)

    IzvSeed TBL_ZBIRNA, _
        Array(COL_ZBR_ID, COL_ZBR_BROJ, COL_ZBR_DATUM, COL_ZBR_VOZAC, COL_ZBR_KUPAC, _
              COL_ZBR_KOLICINA, COL_ZBR_KOL_AMB), _
        Array("IZVT-ZBR-B", IZVT_BROJ, IZVT_DATUM, IZVT_VOZAC_B, IZVT_KUPAC_B, 2000#, 200)

    IzvSeed TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_BROJ_ZBIRNE, _
              COL_PRJ_VOZAC, COL_PRJ_KUPAC, COL_PRJ_KOLICINA), _
        Array("IZVT-PRJ-A", "IZVT-PRJ-A", IZVT_DATUM, IZVT_BROJ, _
              IZVT_VOZAC_A, IZVT_KUPAC_A, 900#)

    IzvSeed TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_BROJ_ZBIRNE, _
              COL_PRJ_VOZAC, COL_PRJ_KUPAC, COL_PRJ_KOLICINA), _
        Array("IZVT-PRJ-B", "IZVT-PRJ-B", IZVT_DATUM, IZVT_BROJ, _
              IZVT_VOZAC_B, IZVT_KUPAC_B, 1500#)
End Sub

' Append reda PO IMENU kolone (preskace kolone kojih nema u semi).
Private Sub IzvSeed(ByVal tblName As String, ByVal cols As Variant, ByVal vals As Variant)
    Dim lo As ListObject
    Set lo = GetTable(tblName)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 7611, "modIzvestajTests.IzvSeed", "Nema tabele: " & tblName
    End If

    Dim nr As ListRow
    Set nr = lo.ListRows.Add

    Dim i As Long, ci As Long
    For i = LBound(cols) To UBound(cols)
        ci = GetColumnIndex(tblName, CStr(cols(i)))
        If ci > 0 Then nr.Range.cells(1, ci).value = vals(i)
    Next i
End Sub

' Indeks reda ciji je `col` brojcano jednak `value` (0 = nije nadjen).
Private Function IzvFindRowByNum(ByVal arr As Variant, ByVal col As Long, _
                                 ByVal value As Double) As Long
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, col)) And Not IsEmpty(arr(i, col)) Then
            If Abs(CDbl(arr(i, col)) - value) < 0.01 Then
                IzvFindRowByNum = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IzvFindRowByText(ByVal arr As Variant, ByVal col As Long, _
                                  ByVal value As String) As Long
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Trim$(NzToText(arr(i, col))) = value Then
            IzvFindRowByText = i
            Exit Function
        End If
    Next i
End Function

' --- (e) FM-0028 #12/#13/#14: nepodrzana kombinacija ne daje tudji izvestaj ---
Private Sub T_DispatchNepodrzanTip()
    Const S As String = "Dispatch nepodrzan tip: "

    Dim od As Date: od = DateSerial(Year(Date), 1, 1)
    Dim doD As Date: doD = Date

    ' Zbirni mod nudi tabove Prosecna cena / Manjak i za Kooperante i Vozace.
    ' Pre RF-06 su te kombinacije vracale GLOBALNI izvestaj pod tim naslovom.
    IzvChk IsEmpty(ReportProsecnaCena("Kooperant", "", od, doD)), _
           S & "ProsecnaCena za Kooperante -> Empty (ne globalni otkup)"
    IzvChk IsEmpty(ReportProsecnaCena("Vozac", "", od, doD)), _
           S & "ProsecnaCena za Vozace -> Empty"
    IzvChk IsEmpty(ReportManjak("Kooperant", "", od, doD)), _
           S & "Manjak za Kooperante -> Empty (ne globalni manjak)"
    IzvChk IsEmpty(ReportAmbalaza("Kooperant", "", od, doD, False)), _
           S & "Ambalaza za Kooperante -> Empty (ne globalni ledger)"
    IzvChk IsEmpty(ReportOtkupRoba("Kooperant", "", od, doD)), _
           S & "OtkupRoba za Kooperante -> Empty"

    ' Podrzane kombinacije SMEJU da prodju (Empty je ovde legitiman ishod ako
    ' nema podataka -- proverava se samo da ne pukne).
    Dim dummy As Variant
    dummy = ReportProsecnaCena("OM", "", od, doD)
    dummy = ReportManjak("", "", od, doD)
    IzvChk True, S & "podrzane kombinacije se i dalje izvrsavaju bez greske"
End Sub

' --- RF-07 (FM-0029 #3): UI labela entiteta -> kod za Report* dispatch ---
Private Sub T_EntitetKod()
    Const S As String = "Entitet kod: "

    IzvChkEqText IzvestajEntitetKod("Otkupna mesta"), "OM", S & "Otkupna mesta -> OM"
    IzvChkEqText IzvestajEntitetKod("Kupci"), "Kupac", S & "Kupci -> Kupac"
    IzvChkEqText IzvestajEntitetKod("Vozaci"), "Vozac", S & "Vozaci -> Vozac"
    IzvChkEqText IzvestajEntitetKod("Kooperanti"), "Kooperant", S & "Kooperanti -> Kooperant"
    IzvChkEqText IzvestajEntitetKod("nepoznato"), "OM", S & "nepoznata labela -> OM (fallback kao ranije)"
End Sub

' --- RF-07 (AUD-024 / FM-0029 #3): matrica dostupnih tabova ---
' Invarijanta: tab se nudi SAMO ako odgovarajuci Report* ima granu za taj tip.
' Regresija (npr. vracanje tabova 5/6/7 svim tipovima u zbirnom rezimu) obara
' ove assert-e.
Private Sub T_TabMatrica()
    Const S As String = "Tab matrica: "

    ' Zbirni rezim -- nevalidne kombinacije koje su pre RF-07 bile ponudjene.
    IzvChk IzvestajTabDostupan("Kooperant", True, IZV_TAB_ZBIRNI) = False, _
           S & "zbirni Kooperant: Zbirni NE (ReportZbirni nema granu)"
    IzvChk IzvestajTabDostupan("Kooperant", True, IZV_TAB_PROSECNA_CENA) = False, _
           S & "zbirni Kooperant: Prosecna cena NE"
    IzvChk IzvestajTabDostupan("Kooperant", True, IZV_TAB_MANJAK) = False, _
           S & "zbirni Kooperant: Manjak NE"
    IzvChk IzvestajTabDostupan("Vozac", True, IZV_TAB_PROSECNA_CENA) = False, _
           S & "zbirni Vozac: Prosecna cena NE (ReportProsecnaCena nema vozacku granu)"
    ' Kupac + zbirni: grana POSTOJI ali ide kroz GetPrijemniceByKupac koji
    ' bezuslovno filtrira KupacID = "" -> tab bi bio ponudjen a trajno prazan.
    ' Dokazuje T_E2E_ProsecnaCenaZbirniKupac nad stvarnim prijemnicama.
    IzvChk IzvestajTabDostupan("Kupac", True, IZV_TAB_PROSECNA_CENA) = False, _
           S & "zbirni Kupac: Prosecna cena NE (KupacID='' filter -> uvek prazno)"

    ' Zbirni rezim -- validne kombinacije moraju da ostanu.
    IzvChk IzvestajTabDostupan("Vozac", True, IZV_TAB_ZBIRNI), _
           S & "zbirni Vozac: Zbirni DA"
    IzvChk IzvestajTabDostupan("Vozac", True, IZV_TAB_MANJAK), _
           S & "zbirni Vozac: Manjak DA"
    IzvChk IzvestajTabDostupan("OM", True, IZV_TAB_ZBIRNI), S & "zbirni OM: Zbirni DA"
    ' OM zadrzava tab: ReportProsecnaCena grana `Case "OM", ""` eksplicitno hvata
    ' prazan entitetID kao "svi" (bez filtera po stanici).
    IzvChk IzvestajTabDostupan("OM", True, IZV_TAB_PROSECNA_CENA), S & "zbirni OM: Prosecna cena DA"
    IzvChk IzvestajTabDostupan("OM", True, IZV_TAB_MANJAK), S & "zbirni OM: Manjak DA"
    IzvChk IzvestajTabDostupan("Kupac", True, IZV_TAB_ZBIRNI), S & "zbirni Kupac: Zbirni DA"

    ' Zbirni rezim NE sme da nudi pojedinacne tabove.
    IzvChk IzvestajTabDostupan("OM", True, IZV_TAB_SALDO_OM) = False, _
           S & "zbirni OM: Saldo OM NE"
    IzvChk IzvestajTabDostupan("OM", True, IZV_TAB_ISPLATA) = False, _
           S & "zbirni OM: Isplata NE"
    IzvChk IzvestajTabDostupan("Kooperant", True, IZV_TAB_KARTICA) = False, _
           S & "zbirni Kooperant: Kartica NE"

    ' Pojedinacni rezim -- zatecena (validna) matrica se ne sme suziti.
    IzvChk IzvestajTabDostupan("OM", False, IZV_TAB_SALDO_OM), S & "OM: Saldo OM DA"
    IzvChk IzvestajTabDostupan("OM", False, IZV_TAB_OTKUP_ROBA), S & "OM: Otkupljena roba DA"
    IzvChk IzvestajTabDostupan("OM", False, IZV_TAB_AMBALAZA), S & "OM: Ambalaza DA"
    IzvChk IzvestajTabDostupan("OM", False, IZV_TAB_ISPLATA), S & "OM: Isplata DA"
    IzvChk IzvestajTabDostupan("OM", False, IZV_TAB_PROSECNA_CENA), S & "OM: Prosecna cena DA"
    IzvChk IzvestajTabDostupan("OM", False, IZV_TAB_SALDO_KUPCI) = False, S & "OM: Saldo Kupci NE"
    IzvChk IzvestajTabDostupan("OM", False, IZV_TAB_MANJAK) = False, S & "OM: Manjak NE"

    IzvChk IzvestajTabDostupan("Kupac", False, IZV_TAB_SALDO_KUPCI), S & "Kupac: Saldo Kupci DA"
    IzvChk IzvestajTabDostupan("Kupac", False, IZV_TAB_MANJAK), S & "Kupac: Manjak DA"
    IzvChk IzvestajTabDostupan("Kupac", False, IZV_TAB_SALDO_OM) = False, S & "Kupac: Saldo OM NE"
    IzvChk IzvestajTabDostupan("Kupac", False, IZV_TAB_ISPLATA) = False, S & "Kupac: Isplata NE"

    IzvChk IzvestajTabDostupan("Vozac", False, IZV_TAB_AMBALAZA), S & "Vozac: Ambalaza DA"
    IzvChk IzvestajTabDostupan("Vozac", False, IZV_TAB_MANJAK), S & "Vozac: Manjak DA"
    IzvChk IzvestajTabDostupan("Vozac", False, IZV_TAB_PROSECNA_CENA) = False, _
           S & "Vozac: Prosecna cena NE"
    IzvChk IzvestajTabDostupan("Vozac", False, IZV_TAB_OTKUP_ROBA) = False, _
           S & "Vozac: Otkupljena roba NE"

    IzvChk IzvestajTabDostupan("Kooperant", False, IZV_TAB_KARTICA), S & "Kooperant: Kartica DA"
    IzvChk IzvestajTabDostupan("Kooperant", False, IZV_TAB_AMBALAZA) = False, _
           S & "Kooperant: Ambalaza (staticki tab) NE"

    ' Runtime tabovi (dinamicki indeks preko statickih) nisu deo matrice.
    IzvChk IzvestajTabDostupan("Kooperant", False, 9) = False, S & "runtime tab nije u matrici"
    IzvChk IzvestajTabDostupan("OM", True, 10) = False, S & "runtime tab nije u matrici (zbirni)"
End Sub

' --- RF-07 (AUD-012 / FM-0029 #4): kljuc reversa nosi i TIP AMBALAZE ---
Private Sub T_ReversKljucPoTipu()
    Const S As String = "Revers kljuc: "

    ' Isti dokument, DVA tipa gajbica -> samo izabrani tip pripada reversu.
    IzvChk ReversRedPripada("OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica", _
                            "OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica"), _
           S & "isti dokument + isti tip -> pripada"
    IzvChk ReversRedPripada("OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Plasticna", _
                            "OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica") = False, _
           S & "isti dokument + DRUGI tip -> NE pripada (bez mesanja tipova)"

    ' Kljuc i dalje trazi isti dokument i isti tip dokumenta.
    IzvChk ReversRedPripada("OTK-2", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica", _
                            "OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica") = False, _
           S & "drugi DokumentID -> NE pripada"
    IzvChk ReversRedPripada("OTK-1", DOK_TIP_OM_ULAZ_KOOP, "Letvarica", _
                            "OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica") = False, _
           S & "drugi DokumentTip -> NE pripada"

    ' Slobodan unos sifarnika: razmaci i velicina slova ne smeju da razbiju match.
    IzvChk ReversRedPripada("OTK-1", DOK_TIP_OM_IZLAZ_KOOP, " letvarica ", _
                            "OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica"), _
           S & "tip: trim + case-insensitive"

    ' Kanonizacija mora da bude ISTA kao u pregledu (`AmbTipKljuc` na obe
    ' putanje) -- inace pregled napravi dva reda a svaki revers sabere oba.
    IzvChkEqText AmbTipKljuc(" letvarica "), AmbTipKljuc("Letvarica"), _
           S & "AmbTipKljuc: trim + velicina slova daju isti kljuc"
    IzvChk AmbTipKljuc("Letvarica") <> AmbTipKljuc("Plasticna"), _
           S & "AmbTipKljuc: razliciti tipovi ostaju razliciti"

    ' Prazan tip je legitimna grupa (red pregleda bez tipa), ne wildcard.
    IzvChk ReversRedPripada("OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "", _
                            "OTK-1", DOK_TIP_OM_IZLAZ_KOOP, ""), _
           S & "prazan tip sa obe strane -> pripada"
    IzvChk ReversRedPripada("OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "Letvarica", _
                            "OTK-1", DOK_TIP_OM_IZLAZ_KOOP, "") = False, _
           S & "prazan trazeni tip NIJE wildcard"
End Sub

' --- RF-07 review nalaz: matrica se ne sme raziti od core-a ---
' Zbirni rezim salje entitetID = "". `ReportProsecnaCena` grana za Kupca ide
' kroz `GetPrijemniceByKupac`, koji BEZUSLOVNO dodaje filter `KupacID = ""` --
' upit trazi prijemnice BEZ kupca, pa vraca prazno i kad prijemnice postoje.
' Matricni test proverava samo True/False; OVAJ test dokazuje da je False
' ISPRAVAN -- nad stvarnim prijemnicama dva razlicita kupca.
'
' TVRD GATE U OBA SMERA: ako neko implementira globalni prosek svih kupaca,
' prvi assert pada i tera da se matrica ponovo odluci -- umesto da tab ostane
' skriven iako bi radio.
Private Sub T_E2E_ProsecnaCenaZbirniKupac()
    Const S As String = "E2E ProsecnaCena (zbirni Kupac): "
    On Error GoTo EH

    Dim d As Date: d = IZVT_DATUM

    ' Dve prijemnice, dva RAZLICITA kupca, obe u sentinel prozoru (1999) -- ni
    ' jedan produkcijski red ne moze da upadne u ovaj datumski opseg.
    IzvSeed TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_BROJ_ZBIRNE, _
              COL_PRJ_KUPAC, COL_PRJ_KOLICINA, COL_PRJ_CENA), _
        Array("IZVT-PC-A", "IZVT-PC-A", d, IZVT_BROJ, IZVT_KUPAC_A, 1000#, 100#)

    IzvSeed TBL_PRIJEMNICA, _
        Array(COL_PRJ_ID, COL_PRJ_BROJ, COL_PRJ_DATUM, COL_PRJ_BROJ_ZBIRNE, _
              COL_PRJ_KUPAC, COL_PRJ_KOLICINA, COL_PRJ_CENA), _
        Array("IZVT-PC-B", "IZVT-PC-B", d, IZVT_BROJ, IZVT_KUPAC_B, 2000#, 200#)

    ' Pojedinacno (entitetID popunjen) izvestaj RADI -- seed je validan i grana
    ' nije mrtva; prazan rezultat ispod je posledica bas praznog entitetID-a.
    IzvChk IsArray(ReportProsecnaCena("Kupac", IZVT_KUPAC_A, d, d)), _
           S & "pojedinacni Kupac A vraca redove (seed validan)"
    IzvChk IsArray(ReportProsecnaCena("Kupac", IZVT_KUPAC_B, d, d)), _
           S & "pojedinacni Kupac B vraca redove"

    ' Zbirni (entitetID = "") NE vraca nista -- zato tab nije u matrici.
    IzvChk IsEmpty(ReportProsecnaCena("Kupac", "", d, d)), _
           S & "zbirni Kupac -> Empty i kad postoje prijemnice dva kupca"
    IzvChk IzvestajTabDostupan("Kupac", True, IZV_TAB_PROSECNA_CENA) = False, _
           S & "matrica se slaze sa core-om: tab se NE nudi"
    Exit Sub

EH:
    m_izvFail = m_izvFail + 1
    Debug.Print "  FAIL | " & S & "ERROR " & Err.Number & ": " & Err.description
End Sub

' --- RF-07 review nalaz: pregled ambalaze mora da grupise PUNIM kljucem ---
' `modOtkup.SaveOtkup` na NORMALNOJ putanji upisuje isti otkupID i isti tip
' ambalaze pod DVA tipa dokumenta: primljene pune gajbe kao DOK_TIP_OTKUP,
' izdate prazne kao DOK_TIP_OM_IZLAZ_KOOP. Dok je grupni kljuc bio samo
' DokumentID + TipAmbalaze, oba su padala u JEDAN red koji nosi tip PRVOG
' zapisa -- skriveni ref-kljuc `AMB|Otkup|<id>`, pa je "Stampaj dokument" uvek
' rutirao na otkupni list, a revers OM-Izlaz-Koop nije imao svoj red i bio je
' NEDOSTUPAN za stampu. Test tvrdi DVA reda i DVA razlicita ref-kljuca.
Private Sub T_E2E_AmbPregledRazdvajaTipDokumenta()
    Const S As String = "E2E Amb pregled (dva tipa dokumenta, isti otkupID): "
    On Error GoTo EH

    Dim d As Date: d = IZVT_DATUM
    Const DOK As String = "IZVT-OTK-AMB"
    Const TIPA As String = "IZVT-Letvarica"

    ' Samo OM/Stanica noge -- ReportAmbalaza("OM", ...) filtrira EntitetTip="Stanica".
    ' Pune gajbe stizu na OM (Ulaz) pod tipom dokumenta "Otkup".
    IzvSeed TBL_AMBALAZA, _
        Array(COL_AMB_ID, COL_AMB_DATUM, COL_AMB_TIP, COL_AMB_KOLICINA, COL_AMB_SMER, _
              COL_AMB_ENTITET, COL_AMB_ENTITET_TIP, COL_AMB_DOK_ID, COL_AMB_DOK_TIP), _
        Array("IZVT-AMB-1", d, TIPA, 20, "Ulaz", _
              IZVT_STANICA, "Stanica", DOK, DOK_TIP_OTKUP)

    ' Prazne gajbe OM izdaje kooperantu (Izlaz) -- ISTI DokumentID, ISTI tip
    ' ambalaze, ali tip dokumenta "OM-Izlaz-Koop".
    IzvSeed TBL_AMBALAZA, _
        Array(COL_AMB_ID, COL_AMB_DATUM, COL_AMB_TIP, COL_AMB_KOLICINA, COL_AMB_SMER, _
              COL_AMB_ENTITET, COL_AMB_ENTITET_TIP, COL_AMB_DOK_ID, COL_AMB_DOK_TIP), _
        Array("IZVT-AMB-2", d, TIPA, 8, "Izlaz", _
              IZVT_STANICA, "Stanica", DOK, DOK_TIP_OM_IZLAZ_KOOP)

    Dim r As Variant
    r = ReportAmbalaza("OM", IZVT_STANICA, d, d, False)
    IzvChk IsArray(r), S & "izvestaj vraca redove"
    If Not IsArray(r) Then Exit Sub

    ' Poslednji red je UKUPNO -> dva dokumenta = 3 reda.
    IzvChkEq UBound(r, 1), 3, S & "dva dokumenta + UKUPNO = 3 reda"
    If UBound(r, 1) < 3 Then Exit Sub

    ' Skriveni ref-kljuc (kol. 7) mora da razlikuje tip dokumenta.
    Dim k1 As String: k1 = CStr(r(1, 7))
    Dim k2 As String: k2 = CStr(r(2, 7))
    IzvChkEqText k1, "AMB|" & DOK_TIP_OTKUP & "|" & DOK, S & "1. red -> ref-kljuc Otkup"
    IzvChkEqText k2, "AMB|" & DOK_TIP_OM_IZLAZ_KOOP & "|" & DOK, _
           S & "2. red -> ref-kljuc OM-Izlaz-Koop (revers dostupan za stampu)"
    IzvChk k1 <> k2, S & "ref-kljucevi su razliciti"

    ' Kolicine ostaju na svom dokumentu (nema spajanja Ulaz+Izlaz tudjih tipova).
    IzvChkEqD NzNum(r(1, 5)), 20#, S & "Otkup red: Ulaz 20"
    IzvChk Len(Trim$(CStr(r(1, 6)))) = 0, S & "Otkup red: Izlaz prazan"
    IzvChk Len(Trim$(CStr(r(2, 5)))) = 0, S & "OM-Izlaz-Koop red: Ulaz prazan"
    IzvChkEqD NzNum(r(2, 6)), 8#, S & "OM-Izlaz-Koop red: Izlaz 8"

    ' UKUPNO se ne menja grupisanjem -- i dalje 20 / 8.
    IzvChkEqD NzNum(r(3, 5)), 20#, S & "UKUPNO Ulaz 20"
    IzvChkEqD NzNum(r(3, 6)), 8#, S & "UKUPNO Izlaz 8"
    Exit Sub

EH:
    m_izvFail = m_izvFail + 1
    Debug.Print "  FAIL | " & S & "ERROR " & Err.Number & ": " & Err.description
End Sub

' Prazna celija ("" kad je smer bez kolicine) -> 0, za brojcano poredjenje.
Private Function NzNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then NzNum = CDbl(v)
End Function

' ============================================================
' ASSERT HELPERI
' ============================================================
Private Sub IzvChk(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        m_izvPass = m_izvPass + 1
    Else
        m_izvFail = m_izvFail + 1
        Debug.Print "  FAIL | " & testName
    End If
End Sub

Private Sub IzvChkEq(ByVal actual As Long, ByVal expected As Long, ByVal testName As String)
    If actual = expected Then
        m_izvPass = m_izvPass + 1
    Else
        m_izvFail = m_izvFail + 1
        Debug.Print "  FAIL | " & testName & " | ocekivano " & expected & ", dobijeno " & actual
    End If
End Sub

' Poredjenje na nivou CENTA (za novcane iznose gde tolerancija 0,01 propusta
' bas gresku koju test fiksira -- npr. nezaokruzeno 33,3333 vs 33,33).
Private Sub IzvChkEqC(ByVal actual As Double, ByVal expected As Double, ByVal testName As String)
    If Abs(actual - expected) < 0.000001 Then
        m_izvPass = m_izvPass + 1
    Else
        m_izvFail = m_izvFail + 1
        Debug.Print "  FAIL | " & testName & " | ocekivano " & expected & ", dobijeno " & actual
    End If
End Sub

Private Sub IzvChkEqD(ByVal actual As Double, ByVal expected As Double, ByVal testName As String)
    If Abs(actual - expected) < 0.01 Then
        m_izvPass = m_izvPass + 1
    Else
        m_izvFail = m_izvFail + 1
        Debug.Print "  FAIL | " & testName & " | ocekivano " & expected & ", dobijeno " & actual
    End If
End Sub

Private Sub IzvChkEqText(ByVal actual As String, ByVal expected As String, ByVal testName As String)
    If actual = expected Then
        m_izvPass = m_izvPass + 1
    Else
        m_izvFail = m_izvFail + 1
        Debug.Print "  FAIL | " & testName & " | ocekivano '" & expected & "', dobijeno '" & actual & "'"
    End If
End Sub

Public Sub SmokeTest_modIzvestaj()
    On Error GoTo EH
    
    Dim datumOd As Date
    Dim datumDo As Date
    
    datumOd = DateSerial(Year(Date), 1, 1)
    datumDo = Date
    
    Debug.Print String(70, "=")
    Debug.Print "SmokeTest_modIzvestaj START | Period: " & _
                Format$(datumOd, "yyyy-mm-dd") & " - " & Format$(datumDo, "yyyy-mm-dd")
    Debug.Print String(70, "-")
    
    Dim stanicaID As String
    Dim kupacID As String
    Dim vozacID As String
    Dim kooperantID As String
    
    stanicaID = Smoke_FirstValue(TBL_STANICE, "StanicaID")
    kupacID = Smoke_FirstValue(TBL_KUPCI, COL_KUP_ID)
    vozacID = Smoke_FirstValue(TBL_VOZACI, "VozacID")
    kooperantID = Smoke_FirstValue(TBL_KOOPERANTI, COL_KOOP_ID)
    
    Debug.Print "Sample IDs:"
    Debug.Print "  StanicaID:   " & Smoke_TextOrSkip(stanicaID)
    Debug.Print "  KupacID:     " & Smoke_TextOrSkip(kupacID)
    Debug.Print "  VozacID:     " & Smoke_TextOrSkip(vozacID)
    Debug.Print "  KooperantID: " & Smoke_TextOrSkip(kooperantID)
    Debug.Print String(70, "-")
    
    ' ========================================================
    ' SALDO
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportSaldoOM", ReportSaldoOM(stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportSaldoOM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportSaldoKupci", ReportSaldoKupci(kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportSaldoKupci", "Nema KupacID u " & TBL_KUPCI
    End If
    
    ' ========================================================
    ' KARTICA KOOPERANTA
    ' ========================================================
    If kooperantID <> "" Then
        Smoke_RunReport "ReportKarticaKooperanta", ReportKarticaKooperanta(kooperantID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportKarticaKooperanta", "Nema KooperantID u " & TBL_KOOPERANTI
    End If
    
    ' ========================================================
    ' ISPLATA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportIsplata OM", ReportIsplata("OM", stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportIsplata OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportIsplata Kupac", ReportIsplata("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportIsplata Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    ' ========================================================
    ' OTKUPLJENA ROBA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportOtkupRoba OM", ReportOtkupRoba("OM", stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportOtkupRoba OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportOtkupRoba Kupac", ReportOtkupRoba("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportOtkupRoba Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    If vozacID <> "" Then
        Smoke_RunReport "ReportOtkupRoba Vozac", ReportOtkupRoba("Vozac", vozacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportOtkupRoba Vozac", "Nema VozacID u " & TBL_VOZACI
    End If
    
    ' ========================================================
    ' AMBALAZA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportAmbalaza OM pojedinacni", ReportAmbalaza("OM", stanicaID, datumOd, datumDo, False)
        Smoke_RunReport "ReportAmbalaza OM zbirni", ReportAmbalaza("OM", stanicaID, datumOd, datumDo, True)
    Else
        Smoke_Skip "ReportAmbalaza OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportAmbalaza Kupac pojedinacni", ReportAmbalaza("Kupac", kupacID, datumOd, datumDo, False)
        Smoke_RunReport "ReportAmbalaza Kupac zbirni", ReportAmbalaza("Kupac", kupacID, datumOd, datumDo, True)
    Else
        Smoke_Skip "ReportAmbalaza Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    If vozacID <> "" Then
        Smoke_RunReport "ReportAmbalaza Vozac pojedinacni", ReportAmbalaza("Vozac", vozacID, datumOd, datumDo, False)
        Smoke_RunReport "ReportAmbalaza Vozac zbirni", ReportAmbalaza("Vozac", vozacID, datumOd, datumDo, True)
    Else
        Smoke_Skip "ReportAmbalaza Vozac", "Nema VozacID u " & TBL_VOZACI
    End If
    
    ' ========================================================
    ' PROSECNA CENA
    ' ========================================================
    If stanicaID <> "" Then
        Smoke_RunReport "ReportProsecnaCena OM", ReportProsecnaCena("OM", stanicaID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportProsecnaCena OM", "Nema StanicaID u " & TBL_STANICE
    End If
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportProsecnaCena Kupac", ReportProsecnaCena("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportProsecnaCena Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    Smoke_RunReport "ReportProsecnaCena zbirni/all", ReportProsecnaCena("OM", "", datumOd, datumDo)
    
    ' ========================================================
    ' MANJAK
    ' ========================================================
    Smoke_RunReport "ReportManjak zbirni/all", ReportManjak("", "", datumOd, datumDo)
    
    If kupacID <> "" Then
        Smoke_RunReport "ReportManjak Kupac", ReportManjak("Kupac", kupacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportManjak Kupac", "Nema KupacID u " & TBL_KUPCI
    End If
    
    If vozacID <> "" Then
        Smoke_RunReport "ReportManjak Vozac", ReportManjak("Vozac", vozacID, datumOd, datumDo)
    Else
        Smoke_Skip "ReportManjak Vozac", "Nema VozacID u " & TBL_VOZACI
    End If
    
    ' ========================================================
    ' ZBIRNI
    ' ========================================================
    Smoke_RunReport "ReportZbirni OM", ReportZbirni("OM", datumOd, datumDo)
    Smoke_RunReport "ReportZbirni Kupac", ReportZbirni("Kupac", datumOd, datumDo)
    Smoke_RunReport "ReportZbirni Vozac", ReportZbirni("Vozac", datumOd, datumDo)
    
    Debug.Print String(70, "-")
    Debug.Print "SmokeTest_modIzvestaj OK"
    Debug.Print String(70, "=")
    Exit Sub

EH:
    Debug.Print String(70, "!")
    Debug.Print "SmokeTest_modIzvestaj ERROR"
    Debug.Print "Err.Number:      " & Err.Number
    Debug.Print "Err.Source:      " & Err.SOURCE
    Debug.Print "Err.Description: " & Err.description
    Debug.Print String(70, "!")
End Sub

Private Sub Smoke_RunReport(ByVal reportName As String, ByVal data As Variant)
    On Error GoTo EH
    
    Debug.Print Smoke_Pad(reportName, 38) & " | " & Smoke_ArrayShape(data)
    Exit Sub
    
EH:
    Debug.Print Smoke_Pad(reportName, 38) & " | ERROR | " & _
                Err.Number & " | " & Err.SOURCE & " | " & Err.description
End Sub

Private Sub Smoke_Skip(ByVal reportName As String, ByVal reason As String)
    Debug.Print Smoke_Pad(reportName, 38) & " | SKIP  | " & reason
End Sub

Private Function Smoke_ArrayShape(ByVal data As Variant) As String
    On Error GoTo EH
    
    If IsEmpty(data) Then
        Smoke_ArrayShape = "EMPTY"
    ElseIf Not IsArray(data) Then
        Smoke_ArrayShape = "NOT ARRAY"
    Else
        Smoke_ArrayShape = "OK    | " & _
                           CStr(UBound(data, 1)) & " rows x " & _
                           CStr(UBound(data, 2)) & " cols"
    End If
    
    Exit Function
    
EH:
    Smoke_ArrayShape = "INVALID ARRAY | " & Err.description
End Function

Private Function Smoke_FirstValue(ByVal tableName As String, ByVal columnName As String) As String
    On Error GoTo EH
    
    Dim data As Variant
    data = GetTableData(tableName)
    
    If IsEmpty(data) Or Not IsArray(data) Then
        Smoke_FirstValue = ""
        Exit Function
    End If
    
    Dim colIdx As Long
    colIdx = GetColumnIndex(tableName, columnName)
    
    If colIdx <= 0 Then
        Smoke_FirstValue = ""
        Exit Function
    End If
    
    Smoke_FirstValue = Trim$(CStr(data(1, colIdx)))
    Exit Function
    
EH:
    Smoke_FirstValue = ""
End Function

Private Function Smoke_TextOrSkip(ByVal value As String) As String
    If Trim$(value) = "" Then
        Smoke_TextOrSkip = "(nije pronadjen)"
    Else
        Smoke_TextOrSkip = value
    End If
End Function

Private Function Smoke_Pad(ByVal value As String, ByVal width As Long) As String
    If Len(value) >= width Then
        Smoke_Pad = Left$(value, width)
    Else
        Smoke_Pad = value & Space$(width - Len(value))
    End If
End Function

