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

    ' --- End-to-end nad tabelama (u transakciji, UVEK rollback) ---
    ' Seam testovi ne mogu da uhvate gresku u samom table-join-u, a upravo je
    ' join po `BrojZbirne` bio nosilac pogresnih brojki.
    If GetTable(TBL_ZBIRNA) Is Nothing Or GetTable(TBL_PRIJEMNICA) Is Nothing _
       Or GetTable(TBL_OTPREMNICA) Is Nothing Then
        Debug.Print "  SKIP | end-to-end: nema tabela zbirna/prijemnica/otpremnica"
    Else
        Set tx = New clsTransaction
        tx.BeginTx
        tx.AddTableSnapshot TBL_ZBIRNA
        tx.AddTableSnapshot TBL_PRIJEMNICA
        tx.AddTableSnapshot TBL_OTPREMNICA
        tx.AddTableSnapshot TBL_OTKUP

        T_E2E_ManjakDvaVlasnikaIstiBroj
        T_E2E_RobaOMDvaVlasnikaIstiBroj

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
    Dim eNum As Long, eDesc As String, eSrc As String
    eNum = Err.Number: eDesc = Err.description: eSrc = Err.SOURCE

    On Error Resume Next
    If Not tx Is Nothing Then tx.RollbackTx
    On Error GoTo 0

    Debug.Print String(70, "!")
    Debug.Print "RunIzvestajTests ERROR | " & eNum & " | " & eSrc & " | " & eDesc
    Debug.Print String(70, "!")
    Err.Raise eNum, eSrc, eDesc
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
    Set p = RaspodeliPoUdelima(100#, tri)
    IzvChkEqC CDbl(p("A")), 33.33, S & "prvi deo zaokruzen na 33,33"
    IzvChkEqC CDbl(p("B")), 33.33, S & "drugi deo zaokruzen na 33,33"
    IzvChkEqC CDbl(p("C")), 33.34, S & "poslednji deo nosi zaokruzeni ostatak 33,34"
    IzvChkEqC CDbl(p("A")) + CDbl(p("B")) + CDbl(p("C")), 100#, _
              S & "zbir ZAOKRUZENIH delova = 100,00 (ne 99,99)"
    IzvChk IsRoundedTo2(CDbl(p("A"))) And IsRoundedTo2(CDbl(p("B"))) And IsRoundedTo2(CDbl(p("C"))), _
           S & "nijedan deo nema vise od 2 decimale"

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

