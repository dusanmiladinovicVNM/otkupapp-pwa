Attribute VB_Name = "modIzvestaj"
'Attribute VB_Name = "modIzvestaj"
Option Explicit

' ============================================================
' modIzvestaj v3.0 - Report Business Logic
' Alle Funktionen geben 2D-Arrays zurueck
' Form ist nur noch fuer UI-Darstellung zustaendig
' ============================================================

' ============================================================
' RF-06 (AUD-023) - deljeni racunski seam-ovi izvestaja.
' Cist ulaz -> cist izlaz, bez citanja tabela: pokrivaju ih assert-i u
' modIzvestajTests.RunIzvestajTests (regresija ponovo obara test).
' Labele su ASCII (kao "UKUPNO" / "OM AVANS (nerasporedjen)") jer se po njima
' traze redovi u testovima -- ne idu kroz modPoruke katalog.
' ============================================================

' Oznaka reda bez prijemnice. Pre RF-06: RobaOM je prikazivao 0 kg / 0,00%
' manjka, a Manjak isti slucaj kao 100% manjka -- dva izvestaja, dva odgovora
' za isti podatak (FM-0028 #5).
Public Const IZV_NEMA_PRIJEMA As String = "nema prijema"

' Oznaka reda kod kog se prijem NE moze pouzdano pripisati zbirnoj: isti
' `BrojZbirne` nose dve aktivne zbirne razlicitih vlasnika (RF-05/AUD-052 je
' dokazao da poslovni broj nije identitet), a red/prijemnica nema podatak koji
' bi jednoznacno razresio vlasnika. Fail-closed: bolje vidljivo "ne znam" nego
' tudja kilaza upisana kao manjak.
Public Const IZV_VLASNIK_NEJASAN As String = "nejasan vlasnik"

' Labela reda pocetnog stanja u karticama (FM-0028 #1).
Public Const IZV_POCETNO_STANJE As String = "POCETNO STANJE"

' RF-07 (AUD-024 / FM-0029 #3) - indeksi STATICKIH stranica mpReports u
' frmIzvestaj (redosled iz .frx); koristi ih matrica IzvestajTabDostupan.
' Runtime tabovi ("Pregled ambalaze", "Otkupni listovi") dobijaju dinamicki
' indeks >= broja statickih stranica i NE prolaze kroz matricu.
' NAPOMENA: modul-level Const MORA u deklaracionu sekciju (pre prve
' procedure) -- VBA ne kompajlira Const izmedju procedura.
Public Const IZV_TAB_SALDO_OM As Long = 0
Public Const IZV_TAB_SALDO_KUPCI As Long = 1
Public Const IZV_TAB_OTKUP_ROBA As Long = 2
Public Const IZV_TAB_AMBALAZA As Long = 3
Public Const IZV_TAB_ISPLATA As Long = 4
Public Const IZV_TAB_ZBIRNI As Long = 5
Public Const IZV_TAB_PROSECNA_CENA As Long = 6
Public Const IZV_TAB_MANJAK As Long = 7
Public Const IZV_TAB_KARTICA As Long = 8

' ============================================================
' SLEDLJIVOST (v6-ui-187) - lanac dokumenata kao read-model.
' Ekran modScrSledljivost je PRIKAZ nad ReportSledljivostLanac /
' ReportSledljivostProblemi (dno modula); nijedno pravilo razresenja se
' tamo ne izmislja:
'  - otkup -> otpremnica ide iskljucivo po OtpremnicaID;
'  - otpremnica/prijemnica -> zbirna ide kroz ISTO pravilo vlasnika kao
'    ReportOtkupRobaOM i ReportManjak (BuildManjakDict + PrijemZaZbirnu:
'    #V>1 bez razresenja po vozacu = fail-closed IZV_VLASNIK_NEJASAN,
'    bez prijemnice = IZV_NEMA_PRIJEMA);
'  - prijemnica -> faktura ide po denorm FakturaID koloni, istoj koju
'    cita ekran Fakturisanja (tblFakturaStavke je normativ, ne cita se).
' Oznake su ASCII konstante (kao IZV_NEMA_PRIJEMA gore) jer se po njima
' traze redovi u testovima -- ne idu kroz modPoruke katalog.
' ============================================================
Public Const SLED_OZN_NEPOVEZAN As String = "nepovezan"
Public Const SLED_OZN_OTP_STORNIRANA As String = "otpremnica stornirana"
Public Const SLED_OZN_VEZA As String = "veza neusaglasena"
Public Const SLED_OZN_BEZ_ZBIRNE As String = "bez zbirne"
Public Const SLED_OZN_ZBIRNA_NEMA As String = "zbirna ne postoji"
' Krug 9: "Fakturisano=Ne" je LEGITIMAN tok (roba u sopstvenu hladnjacu
' -> paleta -> prerada; faktura nije obavezna karika sledljivosti).
' Oznaka postoji SAMO za podatkovnu kontradikciju: prijemnica TVRDI da
' je fakturisana, a veza ne pokazuje na postojecu aktivnu fakturu.
Public Const SLED_OZN_FAK_NEISPRAVNA As String = "faktura neusaglasena"
Public Const SLED_OZN_KG As String = "kg razlika"

' Stanje reda LANCA (kolona 30, GP grana) -- NAJDALJA dostignuta karika
' toka robe. Pise se SAMO kad je red bez oznake (potpun lanac) i kad je
' broj zbirne jednoznacan (nVl = 1); dvosmislen broj ne sme da pripise
' tudje palete (isti fail-closed kao mete, krug 8 R3). Redosled provere:
' GP faktura > prerada > sveza faktura > paleta > otvoren tok.
Public Const SLED_ST_PRODATO_GP As String = "prodato GP"
' Krug 5 (parcijalna prodaja): deo proizvedene robe je validno
' utovaren i fakturisan, deo je jos na stanju.
Public Const SLED_ST_DELIMICNO As String = "delimicno prodato"
' Model B (revizija #10 B1): utovar je samostalan dogadjaj -- roba je
' FIZICKI otisla i pre fakture, pa lanac to mora reci (ne "preradjeno"
' dok kamion vozi robu kupcu).
Public Const SLED_ST_UTOVARENO As String = "utovareno / ceka fakturu"
Public Const SLED_ST_DELIM_UTOVAR As String = "delimicno utovareno"
Public Const SLED_ST_PRERADJENO As String = "preradjeno"
Public Const SLED_ST_PRODATO_SVEZE As String = "prodato svezo"
Public Const SLED_ST_HLADNJACA As String = "u hladnjaci"
Public Const SLED_ST_OTVOREN As String = "otvoren tok"
Public Const SLED_OZN_BEZ_PARCELE As String = "bez parcele"

' Klase problema u ReportSledljivostProblemi (kolona 1). ASCII kodovi;
' prikazni tekst daje ekran kroz modPoruke.
Public Const SLEDP_BEZ_OTPREMNICE As String = "OTKUP-BEZ-OTPREMNICE"
Public Const SLEDP_VEZA As String = "VEZA-NEUSAGLASENA"
Public Const SLEDP_BEZ_ZBIRNE As String = "OTPREMNICA-BEZ-ZBIRNE"
Public Const SLEDP_BROJ_DVOSMISLEN As String = "BROJ-ZBIRNE-DVOSMISLEN"
Public Const SLEDP_BEZ_PRIJEMA As String = "ZBIRNA-BEZ-PRIJEMA"
Public Const SLEDP_FAK_NEISPRAVNA As String = "FAKTURA-VEZA-NEISPRAVNA"
Public Const SLEDP_KG_RAZLIKA As String = "KG-RAZLIKA"

' Tipovi meta sledljivosti (ReportSledljivostMete, kolona 1) -- ASCII
' kodovi za rutiranje stampe; prikazno ime daje ekran kroz modPoruke.
' NEJASNA (krug 8 R3): broj zbirne dele RAZLICITI aktivni vlasnici
' (vozac+kupac) -- sledljivost po prostom broju bi sabrala tudje tokove,
' pa se takva meta nudi samo kao oznaka, BEZ stampe (fail-closed).
Public Const SLEDM_ZBIRNA As String = "ZBIRNA"
Public Const SLEDM_PALETA As String = "PALETA"
Public Const SLEDM_PRERADA As String = "PRERADA"
Public Const SLEDM_NEJASNA As String = "ZBIRNA-NEJASNA"

' Vrsta karike "zbirna" za rutu stampe u NEPOTPUNI listi ekrana. Zbirna
' nema svoju stampu, pa vrsta postoji da radnja ume da ODBIJE s razlogom
' (legacy Case Else obrazac) -- ne da bi se stampalo.
Public Const SLED_DOK_ZBIRNA As String = "Zbirna"

' Vrsta karike "prerada" (GP grana): NEPOTPUNI red kontradiktorne
' prerade nosi je kao DokTip, a radnja stampe rutira na preradni list
' (ExportPreradaPDF -- ista ruta kao meta PRERADA).
Public Const SLED_DOK_PRERADA As String = "Prerada"
' Vrsta karike "utovar" (krug 5): utovarna lista jos nema stampu --
' radnja ODBIJA s razlogom (isti obrazac kao zbirna).
Public Const SLED_DOK_UTOVAR As String = "Utovar"

' Prag poredjenja kg niz lanac -- ISTA vrednost kao (privatni)
' modDokumentInvariant.EPS_KG: dva mesta, jedan prag (par. 12.4 "prag je
' isti kao kod slaganja"). Ne menjati jedno bez drugog.
Public Const SLED_EPS_KG As Double = 0.01

' Pripada li tblNovac red stanici. Primarno po OMID-u SAMOG REDA (istorijska
' pripadnost -- isti kljuc koji ReportIsplata("OM") vec koristi), pa se isplate
' vise ne prelivaju izmedju stanica (FM-0028 #3). Red bez OMID-a (npr. stariji
' upis ili virman bez stanice) nema istorijsku stanicu, pa pada na maticnu
' stanicu kooperanta -- inace bi takav novac nestao iz SVIH stanica (FM-0028 #9
' ostaje pokriven samo za redove koji nose OMID; sire je stvar migracije).
Public Function NovacRedPripadaStanici(ByVal rowOMID As String, _
                                       ByVal koopMaticnaStanica As String, _
                                       ByVal stanicaID As String) As Boolean
    If Len(Trim$(rowOMID)) > 0 Then
        NovacRedPripadaStanici = (Trim$(rowOMID) = Trim$(stanicaID))
    Else
        NovacRedPripadaStanici = (Trim$(koopMaticnaStanica) = Trim$(stanicaID))
    End If
End Function

' Odluka o pouzdanosti prijema za jednu zbirnu -- deljena izmedju ReportManjak
' i ReportOtkupRobaOM. Cist racun, bez tabela (testira RunIzvestajTests).
'
'   brojVlasnika     koliko RAZLICITIH aktivnih vlasnika (vozac+kupac) nosi taj
'                    BrojZbirne; 1 = broj je pouzdan identitet
'   vlasnikRazresen  da li je pozivalac uspeo da odredi TACNOG vlasnika reda
'   cntNejasan       prijemnice tog broja bez kompletnog vlasnika
'   cntPrijem/kgPrijem  prijem u opsegu koji je pozivalac razresio
'
' Returns: Array(imaPrijem As Boolean, prijemKg As Double, oznaka As String)
Public Function PrijemZaZbirnu(ByVal brojVlasnika As Long, _
                               ByVal vlasnikRazresen As Boolean, _
                               ByVal cntNejasan As Long, _
                               ByVal cntPrijem As Long, _
                               ByVal kgPrijem As Double) As Variant
    If brojVlasnika > 1 Then
        ' Broj dele dve zbirne. Ako vlasnik reda nije razresen, ili postoji
        ' prijemnica koja se ne moze pripisati nijednoj -- ne racunamo manjak.
        If (Not vlasnikRazresen) Or cntNejasan > 0 Then
            PrijemZaZbirnu = Array(False, 0#, IZV_VLASNIK_NEJASAN)
            Exit Function
        End If
    End If

    If cntPrijem <= 0 Then
        PrijemZaZbirnu = Array(False, 0#, IZV_NEMA_PRIJEMA)
        Exit Function
    End If

    PrijemZaZbirnu = Array(True, kgPrijem, "")
End Function

' Jedan racun manjka za obe putanje: ReportOtkupRobaOM (po otpremnici) i
' ReportManjak (po zbirnoj). Bez prijema brojke ostaju PRAZNE (ne 0, ne 100%)
' i nose oznaku (IZV_NEMA_PRIJEMA ili IZV_VLASNIK_NEJASAN); pozivalac ih tada
' ne sme uracunati u UKUPNO.
' Returns: Array(prijemKg, manjakKg, manjakPct, oznaka)
Public Function ManjakStavka(ByVal osnovicaKg As Double, _
                             ByVal prijemKg As Double, _
                             ByVal imaPrijem As Boolean, _
                             Optional ByVal oznakaBezPrijema As String = IZV_NEMA_PRIJEMA) As Variant
    If Not imaPrijem Then
        ManjakStavka = Array(Empty, Empty, Empty, oznakaBezPrijema)
        Exit Function
    End If

    Dim manjak As Double
    manjak = osnovicaKg - prijemKg

    Dim pct As Double
    pct = 0
    If osnovicaKg > 0 Then pct = manjak / osnovicaKg * 100

    ManjakStavka = Array(prijemKg, manjak, pct, "")
End Function

' Kartica kooperanta: sortirani period-redovi + pocetno stanje -> rezultat.
' arr = (1..N, 1..8): 1 Datum, 2 BrojDok, 3 Parcela, 4 Opis, 5 Zaduzenje,
' 6 Razduzenje, 7 AmbDelta, 8 RefKljuc.
' Pre RF-06 je running saldo krenuo od NULE, pa je kolona "Saldo" zapravo
' prikazivala neto promenu perioda (FM-0028 #1). Sada se, kad postoji promet
' pre datumOd, ubacuje red IZV_POCETNO_STANJE i saldo krece od njega.
' UKUPNO zadrzava PROMET PERIODA u kolonama 5/6, a kolona 7 je ZAVRSNI saldo
' (pocetno + promet) -- to je red koji operater cita kao dug kooperanta.
' Returns: 2D Array (1..N[+1], 1..9)
Public Function KarticaRezultatSaPocetnim(ByVal arr As Variant, _
                                          ByVal pocetniSaldo As Double, _
                                          ByVal pocetniSaldoAmb As Double) As Variant
    Dim redova As Long
    redova = 0
    If IsArray(arr) Then
        If Not IsEmpty(arr) Then redova = UBound(arr, 1)
    End If

    Dim imaPocetno As Boolean
    imaPocetno = (pocetniSaldo <> 0 Or pocetniSaldoAmb <> 0)

    If redova = 0 And Not imaPocetno Then
        KarticaRezultatSaPocetnim = Empty
        Exit Function
    End If

    Dim offset As Long
    offset = 0
    If imaPocetno Then offset = 1

    Dim result() As Variant
    ReDim result(1 To redova + offset + 1, 1 To 9)

    Dim runSaldo As Double, runSaldoAmb As Double
    runSaldo = pocetniSaldo
    runSaldoAmb = pocetniSaldoAmb

    If imaPocetno Then
        result(1, 1) = ""                  ' bez datuma: nije promet, nego stanje
        result(1, 2) = ""
        result(1, 3) = ""
        result(1, 4) = IZV_POCETNO_STANJE
        result(1, 5) = Empty               ' ne ulazi u promet perioda
        result(1, 6) = Empty
        result(1, 7) = pocetniSaldo
        result(1, 8) = pocetniSaldoAmb
        result(1, 9) = ""
    End If

    Dim totZad As Double, totRaz As Double
    Dim i As Long
    For i = 1 To redova
        result(i + offset, 1) = arr(i, 1)
        result(i + offset, 2) = arr(i, 2)
        result(i + offset, 3) = arr(i, 3)
        result(i + offset, 4) = arr(i, 4)
        result(i + offset, 5) = arr(i, 5)
        result(i + offset, 6) = arr(i, 6)

        runSaldo = runSaldo + arr(i, 5) - arr(i, 6)
        result(i + offset, 7) = runSaldo

        runSaldoAmb = runSaldoAmb + arr(i, 7)
        result(i + offset, 8) = runSaldoAmb

        result(i + offset, 9) = arr(i, 8)

        totZad = totZad + arr(i, 5)
        totRaz = totRaz + arr(i, 6)
    Next i

    Dim ukRow As Long
    ukRow = redova + offset + 1
    result(ukRow, 4) = "UKUPNO"
    result(ukRow, 5) = totZad
    result(ukRow, 6) = totRaz
    result(ukRow, 7) = runSaldo        ' zavrsni saldo = pocetno + promet perioda
    result(ukRow, 8) = runSaldoAmb
    result(ukRow, 9) = ""

    KarticaRezultatSaPocetnim = result
End Function

' Pregled ambalaze kooperanta: isti princip kao KarticaRezultatSaPocetnim.
' arr = (1..N, 1..5): 1 Datum, 2 BrojDok, 3 Opis, 4 Ulaz, 5 Izlaz.
' Returns: 2D Array (1..N[+1], 1..6); kol. 6 = running saldo od pocetnog stanja.
Public Function KarticaAmbRezultatSaPocetnim(ByVal arr As Variant, _
                                             ByVal pocetniSaldo As Double) As Variant
    Dim redova As Long
    redova = 0
    If IsArray(arr) Then
        If Not IsEmpty(arr) Then redova = UBound(arr, 1)
    End If

    Dim imaPocetno As Boolean
    imaPocetno = (pocetniSaldo <> 0)

    If redova = 0 And Not imaPocetno Then
        KarticaAmbRezultatSaPocetnim = Empty
        Exit Function
    End If

    Dim offset As Long
    offset = 0
    If imaPocetno Then offset = 1

    Dim result() As Variant
    ReDim result(1 To redova + offset + 1, 1 To 6)

    If imaPocetno Then
        result(1, 1) = ""
        result(1, 2) = ""
        result(1, 3) = IZV_POCETNO_STANJE
        result(1, 4) = Empty
        result(1, 5) = Empty
        result(1, 6) = pocetniSaldo
    End If

    Dim runSaldo As Double, totU As Double, totI As Double
    runSaldo = pocetniSaldo

    Dim i As Long
    For i = 1 To redova
        result(i + offset, 1) = arr(i, 1)
        result(i + offset, 2) = arr(i, 2)
        result(i + offset, 3) = arr(i, 3)
        result(i + offset, 4) = arr(i, 4)
        result(i + offset, 5) = arr(i, 5)
        runSaldo = runSaldo + arr(i, 4) - arr(i, 5)
        result(i + offset, 6) = runSaldo
        totU = totU + arr(i, 4)
        totI = totI + arr(i, 5)
    Next i

    Dim ukRow As Long
    ukRow = redova + offset + 1
    result(ukRow, 3) = "UKUPNO"
    result(ukRow, 4) = totU
    result(ukRow, 5) = totI
    result(ukRow, 6) = runSaldo        ' zavrsno stanje, ne neto promena perioda

    KarticaAmbRezultatSaPocetnim = result
End Function

' ============================================================
' RF-07 (AUD-024 / AUD-012) - deljeni UI seam-ovi izvestaja.
' Cist ulaz -> cist izlaz, bez citanja tabela; pokrivaju ih assert-i u
' modIzvestajTests.RunIzvestajTests. Zive OVDE (a ne u frmIzvestaj) jer se
' privatne procedure forme ne mogu testirati, a bas su te odluke nosile
' pogresne izvestaje (nevalidne zbirne kombinacije, mesanje tipova ambalaze).
' ============================================================

' UI labela entiteta (caption toggle dugmeta) -> interni kod koji Report*
' funkcije dispecuju. Jedno mesto istine: pre RF-07 je isti Select Case
' postojao samo u btnUnos_Click, dok je UpdateReportMode radio nad labelama.
Public Function IzvestajEntitetKod(ByVal uiLabel As String) As String
    Select Case uiLabel
        Case "Otkupna mesta": IzvestajEntitetKod = "OM"
        Case "Kupci":         IzvestajEntitetKod = "Kupac"
        Case "Vozaci":        IzvestajEntitetKod = "Vozac"
        Case "Kooperanti":    IzvestajEntitetKod = "Kooperant"
        Case Else:            IzvestajEntitetKod = "OM"
    End Select
End Function

' Sme li tab `pageIdx` da bude ponudjen za dati entitet + rezim (FM-0029 #3).
' Matrica prati STVARNI dispatch Report* funkcija -- tab se nudi samo ako
' odgovarajuci izvestaj ima granu za taj tip:
'   ReportZbirni       OM / Kupac / Vozac
'   ReportProsecnaCena OM (uklj. zbirno "") / Kupac SAMO pojedinacno
'   ReportManjak       OM / Kupac / Vozac
'   ReportAmbalaza     OM / Kupac / Vozac
'   ReportOtkupRoba    OM / Kupac / Vozac
' Pre RF-07 su zbirni tabovi 5/6/7 bili vidljivi SVIM tipovima, pa su npr.
' Kooperanti u zbirnom rezimu dobijali prazne liste pod punim naslovom.
'
' Kriterijum NIJE "postoji Case grana za taj tip" nego "grana vraca podatke za
' TAJ entitetID" -- zbirni rezim salje entitetID = "", pa grana koja taj prazan
' ID ubacuje u filter ne moze nista da vrati (vidi Kupac ispod).
Public Function IzvestajTabDostupan(ByVal entitetTip As String, _
                                    ByVal zbirni As Boolean, _
                                    ByVal pageIdx As Long) As Boolean
    If zbirni Then
        Select Case entitetTip
            Case "OM"
                ' ReportProsecnaCena grana `Case "OM", ""` eksplicitno hvata
                ' entitetID = "" kao "svi" (bez filtera po stanici) -> radi.
                ' SALDO_OM i ISPLATA zbirno (krug 9, odluka operatera --
                ' "fali sadrzaj za zbirne"): red = stanica, isti racun kao
                ' pojedinacni oblik (ReportSaldoOMZbirni/ReportIsplataZbirniOM).
                ' AMBALAZA zbirno je legacy grana ReportAmbalazeZbirni
                ' (agregat po tipu gajbe ZA izabranog entiteta) -- do sada
                ' implementirana a neponudjena u UI (par. 23.7).
                Select Case pageIdx
                    Case IZV_TAB_ZBIRNI, IZV_TAB_PROSECNA_CENA, IZV_TAB_MANJAK, _
                         IZV_TAB_SALDO_OM, IZV_TAB_AMBALAZA, IZV_TAB_ISPLATA, _
                         IZV_TAB_OTKUP_ROBA
                        IzvestajTabDostupan = True
                End Select
            Case "Kupac"
                ' Krug 11 ("fale salda po kupcima, u robi roba po kupcu"):
                ' SALDO_KUPCI i OTKUP_ROBA zbirno = red po kupcu, UKUPNO red
                ' pojedinacnog izvestaja (ReportSaldoKupciZbirni /
                ' ReportRobaKupciZbirni) -- isti obrazac kao stanice.
                Select Case pageIdx
                    Case IZV_TAB_ZBIRNI, IZV_TAB_MANJAK, IZV_TAB_AMBALAZA, _
                         IZV_TAB_SALDO_KUPCI, IZV_TAB_OTKUP_ROBA
                        IzvestajTabDostupan = True
                End Select
            Case "Vozac"
                ' Prosecna cena zbirno ne postoji ni za kupca ni za vozaca:
                ' vozacka grana u ReportProsecnaCena ne postoji, a kupceva ide
                ' kroz GetPrijemniceByKupac koji bezuslovno filtrira po
                ' KupacID (zbirno "" = trajno prazno). Globalni prosek je nov
                ' izvestaj (poslovna odluka), ne UI podesavanje. Gate:
                ' T_E2E_ProsecnaCenaZbirniKupac. AMBALAZA zbirno = legacy
                ' agregat po tipu za izabranog (krug 9).
                ' + OTKUP_ROBA (krug 12): otpremljeno PO VOZACU.
                Select Case pageIdx
                    Case IZV_TAB_ZBIRNI, IZV_TAB_MANJAK, IZV_TAB_AMBALAZA, _
                         IZV_TAB_OTKUP_ROBA
                        IzvestajTabDostupan = True
                End Select
            Case Else
                ' Kooperant: zbirni izvestaj ne postoji ni u jednom Report*
                ' (odluka i kruga 12: rang JE njihov zbirni pogled).
                IzvestajTabDostupan = False
        End Select
        Exit Function
    End If

    Select Case entitetTip
        Case "OM"
            Select Case pageIdx
                Case IZV_TAB_SALDO_OM, IZV_TAB_OTKUP_ROBA, IZV_TAB_AMBALAZA, _
                     IZV_TAB_ISPLATA, IZV_TAB_PROSECNA_CENA
                    IzvestajTabDostupan = True
            End Select
        Case "Kupac"
            Select Case pageIdx
                Case IZV_TAB_SALDO_KUPCI, IZV_TAB_OTKUP_ROBA, IZV_TAB_AMBALAZA, _
                     IZV_TAB_PROSECNA_CENA, IZV_TAB_MANJAK
                    IzvestajTabDostupan = True
            End Select
        Case "Vozac"
            Select Case pageIdx
                Case IZV_TAB_AMBALAZA, IZV_TAB_MANJAK
                    IzvestajTabDostupan = True
            End Select
        Case "Kooperant"
            IzvestajTabDostupan = (pageIdx = IZV_TAB_KARTICA)
        Case Else
            IzvestajTabDostupan = False
    End Select
End Function

' Kanonski oblik tipa ambalaze -- JEDAN izvor istine za grupisanje reda u
' pregledu (`ReportAmbalazePojedinacni`) i za match u reversu
' (`ReversRedPripada`). Tip dolazi iz slobodnog unosa sifarnika, pa se razlikuju
' po razmacima i velicini slova. Dok su dve putanje normalizovale RAZLICITO
' (pregled: sirov string; revers: trim + vbTextCompare), "Letvarica" i
' "letvarica" su davali DVA reda pregleda, a svaki revers je sabirao OBA --
' tiho vracanje bas onog mesanja koje RF-07 zatvara.
Public Function AmbTipKljuc(ByVal tipAmb As String) As String
    AmbTipKljuc = UCase$(Trim$(tipAmb))
End Function

' Pripada li red tblAmbalaza reversu koji se stampa (AUD-012 / FM-0029 #4).
' Kljuc reversa je DokumentID + DokumentTip + TIP AMBALAZE: jedan dokument sme
' da nosi vise tipova gajbica, a pre RF-07 se tip uzimao sa PRVOG reda dok su
' se kolicine sabirale preko SVIH tipova -> revers na 40 "letvarica" za promet
' 25 letvarica + 15 plasticnih. Tip se poredi preko `AmbTipKljuc` -- istog
' kljuca po kom pregled grupise redove, pa je poklapanje reda i reversa 1:1.
' Za samostalan revers identitet dokumenta je ReversID (ReversStampaNoge); ovo je
' samo poklapanje broja, smera i tipa ambalaze.
Public Function ReversRedPripada(ByVal rowDokID As String, ByVal rowDokTip As String, _
                                 ByVal rowTipAmb As String, _
                                 ByVal dokID As String, ByVal dokTip As String, _
                                 ByVal tipAmb As String) As Boolean
    If Trim$(rowDokID) <> Trim$(dokID) Then Exit Function
    If Trim$(rowDokTip) <> Trim$(dokTip) Then Exit Function
    ReversRedPripada = (AmbTipKljuc(rowTipAmb) = AmbTipKljuc(tipAmb))
End Function

' NOGE REVERSA ZA STAMPU -- redovi (indeksi u GetTableData) jednog papira: jedan
' dokument, jedan tip ambalaze (AUD-012), bez storniranih.
'   Samostalan revers -> identitet je ReversID izabranog reda pregleda
'                        (REV-IDENT-01). Red bez njega se odbija -- bez pogadjanja
'                        po broju, stanici ili danu (A2 red REV: isti broj legalno
'                        nose reversi druge stanice ili drugog dana).
'   Ambalaza uz otkup  -> DokumentID je OtkupID, vec jedinstven; ReversID-a nema.
' Vraca "" i redove u `redovi`; inace razlog. Red istog ReversID-a sa drugim brojem
' ili smerom je integritetska greska (B10) -- stampa se tada odbija, red se ne
' preskace tiho.
Public Function ReversStampaNoge(ByVal dokID As String, ByVal dokTip As String, _
                                 ByVal tipAmb As String, ByVal reversID As String, _
                                 ByRef redovi As Collection) As String
    Const SRC As String = "modIzvestaj.ReversStampaNoge"
    Set redovi = New Collection

    Dim d As Variant: d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then
        ReversStampaNoge = "Revers " & dokID & " nije pronadjen."
        Exit Function
    End If
    Dim cTip As Long, cDok As Long, cDokTip As Long, cStorno As Long, cRid As Long
    cTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, SRC)
    cDok = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, SRC)
    cDokTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, SRC)
    cRid = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_REVERS_ID, SRC)
    cStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)

    Dim uzOtkup As Boolean
    uzOtkup = (Len(NzToText(LookupValue(TBL_OTKUP, COL_OTK_ID, Trim$(dokID), COL_OTK_ID))) > 0)
    reversID = Trim$(reversID)
    If Not uzOtkup And Len(reversID) = 0 Then
        ReversStampaNoge = "Revers " & dokID & " nema ReversID u izabranom redu -> dokument nije " & _
                           "jednoznacan (integritet B10). Stampa odbijena."
        Exit Function
    End If
    If Not uzOtkup Then
        ' Granica dokumenta: papir ne sme da pokupi red tudjeg dokumenta sa istim
        ' ReversID-om, pa ni da se sklopi preko njega (modStorno.ReversIDGranica).
        Dim granica As String: granica = modStorno.ReversIDGranica(reversID)
        If Len(granica) > 0 Then
            ReversStampaNoge = granica & " Stampa odbijena."
            Exit Function
        End If
    End If

    Dim i As Long
    For i = 1 To UBound(d, 1)
        If Not IzvAmbRedStorniran(d, i, cStorno) Then
            If uzOtkup Then
                If ReversRedPripada(NzToText(d(i, cDok)), NzToText(d(i, cDokTip)), _
                                    NzToText(d(i, cTip)), dokID, dokTip, tipAmb) Then redovi.Add i
            ElseIf StrComp(Trim$(NzToText(d(i, cRid))), reversID, vbTextCompare) = 0 Then
                If AmbTipKljuc(NzToText(d(i, cTip))) = AmbTipKljuc(tipAmb) Then
                    If Not ReversRedPripada(NzToText(d(i, cDok)), NzToText(d(i, cDokTip)), _
                                            NzToText(d(i, cTip)), dokID, dokTip, tipAmb) Then
                        Set redovi = New Collection
                        ReversStampaNoge = "Revers " & reversID & " nose redovi razlicitog broja ili " & _
                                           "smera (integritet B10). Stampa odbijena."
                        Exit Function
                    End If
                    redovi.Add i
                End If
            End If
        End If
    Next i
    If redovi.count = 0 Then
        ReversStampaNoge = "Revers nije moguce rekonstruisati (nema aktivnih nogu za izabrani red)."
    End If
End Function

Public Function ReportSaldoOM(ByVal stanicaID As String, _
                              ByVal datumOd As Date, _
                              ByVal datumDo As Date) As Variant
                              
    Const SRC As String = "modIzvestaj.ReportSaldoOM"
    On Error GoTo EH
    ' Returns: 2D Array (Name, Kolicina, Vrednost, Novac, Saldo, Ambalaza)
    ' Letzte Zeile = UKUPNO
    
    Dim otkupData As Variant
    otkupData = GetOtkupByStation(stanicaID, datumOd, datumDo)
    
    ' --- Otkup pro Kooperant aggregieren ---
    ' Nema early-exit ako nema otkupa:
    ' report mora i dalje da prikaze novac / OM avans ako postoje u periodu.
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    
    If Not IsEmpty(otkupData) Then
        If IsArray(otkupData) Then
            otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
            
            If IsArray(otkupData) Then
                Dim colKoop As Long, colOtkID As Long
                colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, "modIzvestaj.ReportSaldoOM")
                colOtkID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "modIzvestaj.ReportSaldoOM")

                ' Kolicina i vrednost dokumenta su na STAVKAMA: CreateOtkup_TX
                ' linijska polja zaglavlja ostavlja prazna (REFAKTOR S14.7, kvar 3).
                ' Kolona Ambalaza je saldo iz tblAmbalaza (nize), ne zbir zaglavlja.
                Dim stavkeZbir As Object
                Set stavkeZbir = modOtkup.ZbirStavkiPoOtkupu()
                
                For i = 1 To UBound(otkupData, 1)
                    Dim key As String
                    key = CStr(otkupData(i, colKoop))
                    
                    If key <> "" Then
                        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#)
                        
                        Dim vals As Variant
                        vals = dict(key)
                        
                        Dim zSal As Variant
                        zSal = modOtkup.ZbirStavkiZaOtkup(stavkeZbir, _
                                   CStr(otkupData(i, colOtkID)), "modIzvestaj.ReportSaldoOM")
                        vals(0) = vals(0) + CDbl(zSal(0))
                        vals(1) = vals(1) + CDbl(zSal(1))
                        
                        dict(key) = vals
                    End If
                Next i
            End If
        End If
    End If
    
    ' Mape kooperanata (ime/prezime, stanica) -- jednom, umesto LookupValue u petljama nize.
    Dim koopNameDict As Object
    Set koopNameDict = BuildLookupDict(TBL_KOOPERANTI, "KooperantID", "Ime", "Prezime")
    Dim koopStanicaDict As Object
    Set koopStanicaDict = BuildLookupDict(TBL_KOOPERANTI, "KooperantID", "StanicaID")

    ' --- Novac pro Kooperant aus tblNovac ---
    Dim novacDict As Object
    Set novacDict = CreateObject("Scripting.Dictionary")
    
    Dim novacData As Variant
    novacData = GetTableData(TBL_NOVAC)
    
    If IsArray(novacData) Then
        novacData = ExcludeStornirano(novacData, TBL_NOVAC)
    End If
    
    Dim colNovKoop As Long, colNovIsplata As Long, colNovDatum As Long
    Dim colNovTip As Long, colNovOMID As Long
    Dim n As Long

    If IsArray(novacData) And Not IsEmpty(novacData) Then
        colNovKoop = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, "modIzvestaj.ReportSaldoOM")
        colNovIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, "modIzvestaj.ReportSaldoOM")
        colNovDatum = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, "modIzvestaj.ReportSaldoOM")
        colNovTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, "modIzvestaj.ReportSaldoOM")
        colNovOMID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OM_ID, "modIzvestaj.ReportSaldoOM")

        For n = 1 To UBound(novacData, 1)
            Dim koopID As String
            koopID = CStr(novacData(n, colNovKoop))
            If koopID <> "" Then
                If IsDate(novacData(n, colNovDatum)) Then
                    If CDate(novacData(n, colNovDatum)) >= datumOd And _
                       CDate(novacData(n, colNovDatum)) <= datumDo Then
                        ' Isplata pripada stanici po OMID-u REDA (istorijski), a tek
                        ' za redove bez OMID-a po maticnoj stanici kooperanta.
                        ' Pre RF-06 se gledala samo maticna stanica, pa je isplata
                        ' izvrsena na jednom OM-u ulazila u izvestaj drugog OM-a.
                        Dim koopStation As String
                        If koopStanicaDict.Exists(koopID) Then koopStation = koopStanicaDict(koopID) Else koopStation = ""

                        If NovacRedPripadaStanici(CStr(novacData(n, colNovOMID)), koopStation, stanicaID) Then
                            If Not dict.Exists(koopID) Then dict.Add koopID, Array(0#, 0#, 0#)

                            If Not novacDict.Exists(koopID) Then novacDict.Add koopID, 0#
                            If IsNumeric(novacData(n, colNovIsplata)) Then
                                novacDict(koopID) = novacDict(koopID) + CDbl(novacData(n, colNovIsplata))
                            End If
                        End If
                    End If
                End If
            End If
        Next n
    End If

        ' --- OM Avans berechnen (VOR dem ReDim) ---
    Dim omAvans As Double
    omAvans = 0

    If IsArray(novacData) And Not IsEmpty(novacData) Then
        For n = 1 To UBound(novacData, 1)
            If CStr(novacData(n, colNovOMID)) = stanicaID Then
                If IsDate(novacData(n, colNovDatum)) Then
                    If CDate(novacData(n, colNovDatum)) >= datumOd And _
                       CDate(novacData(n, colNovDatum)) <= datumDo Then
                        If IsNumeric(novacData(n, colNovIsplata)) Then
                            ' Avans Firma->Otkupac: oba kanala (kes + virman iz izvoda).
                            If IsFirmaOtkupacAvansTip(CStr(novacData(n, colNovTip))) Then
                                omAvans = omAvans + CDbl(novacData(n, colNovIsplata))
                            ElseIf CStr(novacData(n, colNovTip)) = NOV_KES_OTKUPAC_KOOP Then
                                omAvans = omAvans - CDbl(novacData(n, colNovIsplata))
                            End If
                        End If
                    End If
                End If
            End If
        Next n
    End If
    
    Dim hasOMAvans As Boolean
    hasOMAvans = (omAvans <> 0)
    
    ' --- Agrohemija pro Kooperant (Dict) ---
    Dim magData As Variant
    magData = GetTableData(TBL_MAGACIN)
    
    If IsArray(magData) Then
        magData = ExcludeStornirano(magData, TBL_MAGACIN)
    End If
    
    Dim colMagKoop As Long, colMagTip As Long, colMagVrednost As Long, colMagDat As Long
    If IsArray(magData) And Not IsEmpty(magData) Then
        colMagKoop = RequireColumnIndex(TBL_MAGACIN, COL_MAG_KOOP, "modIzvestaj.ReportSaldoOM")
        colMagTip = RequireColumnIndex(TBL_MAGACIN, COL_MAG_TIP, "modIzvestaj.ReportSaldoOM")
        colMagVrednost = RequireColumnIndex(TBL_MAGACIN, COL_MAG_VREDNOST, "modIzvestaj.ReportSaldoOM")
        colMagDat = RequireColumnIndex(TBL_MAGACIN, COL_MAG_DATUM, "modIzvestaj.ReportSaldoOM")
    End If
    
    Dim agroKoopDict As Object
    Set agroKoopDict = CreateObject("Scripting.Dictionary")
    Dim agroBezStanica As Double  ' nerasporedjena Agrohemija (kein Kooperant)
    agroBezStanica = 0
    
    If IsArray(magData) And Not IsEmpty(magData) Then
        Dim m As Long
            For m = 1 To UBound(magData, 1)
                If CStr(magData(m, colMagTip)) = MAG_IZLAZ Then
                    If IsDate(magData(m, colMagDat)) Then
                        If CDate(magData(m, colMagDat)) >= datumOd And _
                           CDate(magData(m, colMagDat)) <= datumDo Then
                            If IsNumeric(magData(m, colMagVrednost)) Then
                                Dim magKoopID As String
                                magKoopID = CStr(magData(m, colMagKoop))
                                
                                If magKoopID <> "" And dict.Exists(magKoopID) Then
                                    If Not agroKoopDict.Exists(magKoopID) Then agroKoopDict.Add magKoopID, 0#
                                    agroKoopDict(magKoopID) = agroKoopDict(magKoopID) + CDbl(magData(m, colMagVrednost))
                                ElseIf magKoopID = "" Then
                                    agroBezStanica = agroBezStanica + CDbl(magData(m, colMagVrednost))
                                End If
                            End If
                        End If
                    End If
                End If
            Next m
    End If
    
    ' --- Aktivni saldo ambalaze po kooperantu (neto iz ledgera: Ulaz - Izlaz,
    '     EntitetTip="Kooperant"); prikazuje se umesto zbira predatih gajbica. ---
    Dim koopAmbDict As Object: Set koopAmbDict = CreateObject("Scripting.Dictionary")
    Dim ambData As Variant: ambData = GetTableData(TBL_AMBALAZA)
    If IsArray(ambData) Then
        ambData = ExcludeStornirano(ambData, TBL_AMBALAZA)
        If IsArray(ambData) And Not IsEmpty(ambData) Then
            Dim caEnt As Long, caEntTip As Long, caKol As Long, caSmer As Long
            caEnt = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, SRC)
            caEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, SRC)
            caKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, SRC)
            caSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, SRC)
            Dim ai As Long
            For ai = 1 To UBound(ambData, 1)
                If Trim$(CStr(ambData(ai, caEntTip))) = "Kooperant" Then
                    Dim akoop As String: akoop = Trim$(CStr(ambData(ai, caEnt)))
                    If akoop <> "" And IsNumeric(ambData(ai, caKol)) Then
                        If Not koopAmbDict.Exists(akoop) Then koopAmbDict.Add akoop, 0&
                        Select Case Trim$(CStr(ambData(ai, caSmer)))
                            Case "Ulaz":  koopAmbDict(akoop) = koopAmbDict(akoop) + CLng(ambData(ai, caKol))
                            Case "Izlaz": koopAmbDict(akoop) = koopAmbDict(akoop) - CLng(ambData(ai, caKol))
                        End Select
                    End If
                End If
            Next ai
        End If
    End If

    ' --- Ergebnis-Array: 7 Spalten ---
    ' Kooperant | Kolicina | Vrednost | Isplaceno | AgroZaduzenje | Saldo | Ambalaza
    
    Dim rowCount As Long
    rowCount = dict.count + 1  ' +UKUPNO
    If hasOMAvans Then rowCount = rowCount + 1
    If agroBezStanica > 0 Then rowCount = rowCount + 1
    
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 7)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totKol As Double, totVr As Double, totNov As Double
    Dim totAgro As Double, totAmb As Long
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        
        Dim novacSum As Double
        novacSum = 0
        If novacDict.Exists(keys(i)) Then novacSum = novacDict(keys(i))
        
        Dim agroSum As Double
        agroSum = 0
        If agroKoopDict.Exists(keys(i)) Then agroSum = agroKoopDict(keys(i))

        Dim koopNaziv As String
        If koopNameDict.Exists(CStr(keys(i))) Then koopNaziv = koopNameDict(CStr(keys(i))) Else koopNaziv = ""

        result(i + 1, 1) = koopNaziv
        result(i + 1, 2) = vals(0)                          ' Kolicina
        result(i + 1, 3) = vals(1)                          ' Vrednost
        result(i + 1, 4) = novacSum                         ' Isplaceno
        result(i + 1, 5) = agroSum                          ' AgroZaduzenje
        result(i + 1, 6) = vals(1) - novacSum - agroSum     ' Saldo
        Dim ambSaldo As Long: ambSaldo = 0
        If koopAmbDict.Exists(keys(i)) Then ambSaldo = CLng(koopAmbDict(keys(i)))
        result(i + 1, 7) = ambSaldo                         ' Ambalaza (aktivni saldo, neto)
        
        totKol = totKol + vals(0)
        totVr = totVr + vals(1)
        totNov = totNov + novacSum
        totAgro = totAgro + agroSum
        totAmb = totAmb + ambSaldo
    Next i
    
    ' OM Avans (nerasporedjen)
    If hasOMAvans Then
        Dim omAvansRow As Long
        omAvansRow = dict.count + 1
        result(omAvansRow, 1) = "OM AVANS (nerasporedjen)"
        result(omAvansRow, 4) = omAvans
        totNov = totNov + omAvans
    End If
    
    ' Agrohemija (nerasporedjena -- ohne Kooperant).
    ' tblMagacin nema kolonu stanice, pa se ovaj iznos NE moze pripisati ni jednom
    ' OM-u: isti broj se pojavljuje u izvestaju SVAKE stanice. Zato ostaje kao
    ' informativan red, ali od RF-06 NE ulazi u UKUPNO (inace bi zbir po stanicama
    ' visestruko brojao isti trosak -- FM-0028 #10).
    If agroBezStanica > 0 Then
        Dim agroRow As Long
        agroRow = rowCount - 1
        result(agroRow, 1) = "AGROHEMIJA (nerasporedjena, van UKUPNO)"
        result(agroRow, 5) = agroBezStanica
    End If
    
    ' UKUPNO
    result(rowCount, 1) = "UKUPNO"
    result(rowCount, 2) = totKol
    result(rowCount, 3) = totVr
    result(rowCount, 4) = totNov
    result(rowCount, 5) = totAgro
    result(rowCount, 6) = totVr - totNov - totAgro
    result(rowCount, 7) = totAmb
    
    ReportSaldoOM = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Public Function ReportKarticaKooperanta(ByVal kooperantID As String, _
                                        ByVal datumOd As Date, _
                                        ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportKarticaKooperanta"
    On Error GoTo EH
    ' Returns: 2D Array
    ' (1)=Datum, (2)=BrojDok, (3)=BrojParcele, (4)=Opis,
    ' (5)=Zaduzenje, (6)=Razduzenje, (7)=Saldo,
    ' (8)=SaldoAmbalaze (running; gajbe = Izdata - Primljena;
    '     ukljucuje i samostalna kretanja ambalaze van otkupa),
    ' (9)=RefKljuc reda ("OTK|<OtkupID>" / "NOV" / "MAG" / "AMB") za Detalje otkupa
    '
    ' RF-06: promet PRE datumOd se vise ne odbacuje nego se sabira u pocetno
    ' stanje, pa kartica krece od stanja duga a ne od nule (FM-0028 #1).

    Dim moves As New Collection

    Dim pocetniSaldo As Double, pocetniSaldoAmb As Double
    pocetniSaldo = 0
    pocetniSaldoAmb = 0

    Dim i As Long

    ' 1. Otkup = Zaduzenje
    Dim otkData As Variant
    otkData = GetTableData(TBL_OTKUP)
    If IsArray(otkData) Then
        otkData = ExcludeStornirano(otkData, TBL_OTKUP)
        If IsArray(otkData) Then
            Dim colOtkDat As Long, colOtkKoop As Long
            Dim colOtkVrsta As Long
            Dim colOtkBrDok As Long, colParcela As Long
            Dim colOtkID As Long

            colOtkID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "modIzvestaj.ReportKarticaKooperanta")
            colOtkDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, "modIzvestaj.ReportKarticaKooperanta")
            colOtkKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, "modIzvestaj.ReportKarticaKooperanta")
            colOtkVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "modIzvestaj.ReportKarticaKooperanta")
            colOtkBrDok = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, "modIzvestaj.ReportKarticaKooperanta")
            colParcela = RequireColumnIndex(TBL_OTKUP, COL_OTK_PARCELA, "modIzvestaj.ReportKarticaKooperanta")

            ' Kolicina, vrednost, klase i PRIMLJENE gajbe dokumenta su na STAVKAMA
            ' (CreateOtkup_TX ih na zaglavlju ostavlja prazne -- REFAKTOR S14.7,
            ' kvar 3). Red kartice ostaje DOKUMENT; klase se nabrajaju u opisu.
            Dim stavkeZbir As Object
            Set stavkeZbir = modOtkup.ZbirStavkiPoOtkupu()

            ' Izdate gajbe su header cinjenica (S3). KolAmbIzdata je noviji stup ->
            ' GetColumnIndex (0 = stara sema, tretiraj kao 0).
            Dim colOtkAmbIzd As Long
            colOtkAmbIzd = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA)

            For i = 1 To UBound(otkData, 1)
                If CStr(otkData(i, colOtkKoop)) = kooperantID Then
                    If IsDate(otkData(i, colOtkDat)) Then
                        Dim otkDatum As Date
                        otkDatum = CDate(otkData(i, colOtkDat))
                        
                        If otkDatum <= datumDo Then
                            Dim vr As Double, otkKol As Double, otkKlase As String
                            Dim ambPrimljena As Double, ambIzdata As Double
                            vr = 0: otkKol = 0: otkKlase = ""
                            ambPrimljena = 0: ambIzdata = 0

                            Dim zKar As Variant, oidKar As String
                            oidKar = Trim$(CStr(otkData(i, colOtkID)))
                            zKar = modOtkup.ZbirStavkiZaOtkup(stavkeZbir, oidKar, _
                                       "modIzvestaj.ReportKarticaKooperanta")
                            otkKol = CDbl(zKar(0))
                            vr = CDbl(zKar(1))
                            ambPrimljena = CDbl(zKar(2))
                            otkKlase = CStr(zKar(3))

                            Dim opis As String
                            opis = "Otkup " & CStr(otkData(i, colOtkVrsta)) & " " & _
                                   otkKlase & " " & _
                                   FmtKolicina(otkKol) & "kg"

                            ' Saldo ambalaze (gajbe): Izdata (OM->koop) - Primljena (koop->OM).
                            ' Isti smer kao kanonski entitetski saldo (modAmbalaza.GetAmbalazeStanje).
                            If colOtkAmbIzd > 0 Then
                                If IsNumeric(otkData(i, colOtkAmbIzd)) Then ambIzdata = CDbl(otkData(i, colOtkAmbIzd))
                            End If

                            If otkDatum < datumOd Then
                                pocetniSaldo = pocetniSaldo + vr
                                pocetniSaldoAmb = pocetniSaldoAmb + (ambIzdata - ambPrimljena)
                            Else
                                moves.Add Array( _
                                    otkDatum, _
                                    CStr(otkData(i, colOtkBrDok)), _
                                    CStr(otkData(i, colParcela)), _
                                    opis, _
                                    vr, _
                                    0#, _
                                    ambIzdata - ambPrimljena, _
                                    "OTK|" & CStr(otkData(i, colOtkID)))
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End If
    
    ' 2. Novac = Razduzenje
    Dim novData As Variant
    novData = GetTableData(TBL_NOVAC)
    If IsArray(novData) Then
        novData = ExcludeStornirano(novData, TBL_NOVAC)
        If IsArray(novData) Then
            Dim colNovDat As Long, colNovKoop As Long
            Dim colNovIsplata As Long, colNovTip As Long, colNovBrDok As Long
            
            colNovDat = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, "modIzvestaj.ReportKarticaKooperanta")
            colNovKoop = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, "modIzvestaj.ReportKarticaKooperanta")
            colNovIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, "modIzvestaj.ReportKarticaKooperanta")
            colNovTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, "modIzvestaj.ReportKarticaKooperanta")
            colNovBrDok = RequireColumnIndex(TBL_NOVAC, COL_NOV_BROJ_DOK, "modIzvestaj.ReportKarticaKooperanta")
            
            Dim n As Long
            For n = 1 To UBound(novData, 1)
                If CStr(novData(n, colNovKoop)) = kooperantID Then
                    If IsDate(novData(n, colNovDat)) Then
                        Dim novDatum As Date
                        novDatum = CDate(novData(n, colNovDat))
                        
                        If novDatum <= datumDo Then
                            Dim iznos As Double
                            iznos = 0
                            If IsNumeric(novData(n, colNovIsplata)) Then
                                iznos = CDbl(novData(n, colNovIsplata))
                            End If

                            If iznos > 0 Then
                                If novDatum < datumOd Then
                                    pocetniSaldo = pocetniSaldo - iznos
                                Else
                                    Dim tipNovca As String
                                    Dim novOpis As String

                                    tipNovca = CStr(novData(n, colNovTip))
                                    Select Case tipNovca
                                        Case NOV_KES_OTKUPAC_KOOP: novOpis = "Ke" & ChrW(353) & " Otkupac"
                                        Case NOV_VIRMAN_FIRMA_KOOP: novOpis = "Virman Firma"
                                        Case NOV_VIRMAN_AVANS_KOOP: novOpis = "Virman Avans"
                                        Case Else: novOpis = tipNovca
                                    End Select

                                    moves.Add Array( _
                                        novDatum, _
                                        CStr(novData(n, colNovBrDok)), _
                                        "", _
                                        novOpis, _
                                        0#, _
                                        iznos, _
                                        0#, _
                                        "NOV")
                                End If
                            End If
                        End If
                    End If
                End If
            Next n
        End If
    End If
    
    ' 3. Agrohemija = Razduzenje
    Dim magData As Variant
    magData = GetTableData(TBL_MAGACIN)
    If IsArray(magData) Then
        magData = ExcludeStornirano(magData, TBL_MAGACIN)
        If IsArray(magData) Then
            Dim colMagDat As Long, colMagKoop As Long, colMagTip As Long
            Dim colMagVrednost As Long, colMagArtikal As Long, colMagBrDok As Long
            
            colMagDat = RequireColumnIndex(TBL_MAGACIN, COL_MAG_DATUM, "modIzvestaj.ReportKarticaKooperanta")
            colMagKoop = RequireColumnIndex(TBL_MAGACIN, COL_MAG_KOOP, "modIzvestaj.ReportKarticaKooperanta")
            colMagTip = RequireColumnIndex(TBL_MAGACIN, COL_MAG_TIP, "modIzvestaj.ReportKarticaKooperanta")
            colMagVrednost = RequireColumnIndex(TBL_MAGACIN, COL_MAG_VREDNOST, "modIzvestaj.ReportKarticaKooperanta")
            colMagArtikal = RequireColumnIndex(TBL_MAGACIN, COL_MAG_ARTIKAL, "modIzvestaj.ReportKarticaKooperanta")
            colMagBrDok = RequireColumnIndex(TBL_MAGACIN, COL_MAG_BR_DOK, "modIzvestaj.ReportKarticaKooperanta")

            ' Mapa artikala (ID -> naziv) -- jednom, umesto LookupValue u petlji nize.
            Dim artikalDict As Object
            Set artikalDict = BuildLookupDict(TBL_ARTIKLI, COL_ART_ID, COL_ART_NAZIV)

            Dim m As Long
            For m = 1 To UBound(magData, 1)
                If CStr(magData(m, colMagKoop)) = kooperantID Then
                    If CStr(magData(m, colMagTip)) = MAG_IZLAZ Then
                        If IsDate(magData(m, colMagDat)) Then
                            Dim magDatum As Date
                            magDatum = CDate(magData(m, colMagDat))
                            
                            If magDatum <= datumDo Then
                                Dim magVr As Double
                                magVr = 0
                                If IsNumeric(magData(m, colMagVrednost)) Then
                                    magVr = CDbl(magData(m, colMagVrednost))
                                End If

                                If magVr > 0 Then
                                    If magDatum < datumOd Then
                                        pocetniSaldo = pocetniSaldo - magVr
                                    Else
                                        Dim artNaziv As String
                                        Dim artKey As String: artKey = CStr(magData(m, colMagArtikal))
                                        If artikalDict.Exists(artKey) Then artNaziv = artikalDict(artKey) Else artNaziv = ""

                                        moves.Add Array( _
                                            magDatum, _
                                            CStr(magData(m, colMagBrDok)), _
                                            "", _
                                            "Agrohemija " & artNaziv, _
                                            0#, _
                                            magVr, _
                                            0#, _
                                            "MAG")
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next m
        End If
    End If

    ' 4. Ambalaza (samostalna kretanja, van otkupa) -> menja samo "Saldo amb."
    '    Otkup-vezane amb stavke (primljene pune /DokTip=Otkup/ I izdate prazne
    '    /DokTip=OM-Izlaz-Koop/) imaju DokID = otkupID i VEC su uracunate kroz
    '    otkup redove (ambIzdata - ambPrimljena). Zato uzimamo SAMO one ciji DokID
    '    NIJE otkupID (prava samostalna kretanja, npr. izdate prazne gajbe bez
    '    otkupa) -> postaju vidljivi redovi i UKUPNO/saldo postaju tacni.
    Dim ambData As Variant
    ambData = GetTableData(TBL_AMBALAZA)
    If IsArray(ambData) Then
        ambData = ExcludeStornirano(ambData, TBL_AMBALAZA)
        If IsArray(ambData) Then
            Dim cAmbDat As Long, cAmbEnt As Long, cAmbEntTip As Long, cAmbTip As Long
            Dim cAmbKol As Long, cAmbSmer As Long, cAmbDokID As Long, cAmbDokTip As Long
            cAmbDat = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, SRC)
            cAmbEnt = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, SRC)
            cAmbEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, SRC)
            cAmbTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, SRC)
            cAmbKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, SRC)
            cAmbSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, SRC)
            cAmbDokID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, SRC)
            cAmbDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)   ' opciono (friendly opis)

            ' Kljucevi = svi otkupID-evi (za iskljucivanje otkup-vezanih amb stavki).
            Dim otkIdDict As Object
            Set otkIdDict = BuildOtkupBrojDokDict()

            Dim a As Long
            For a = 1 To UBound(ambData, 1)
                If NzToText(ambData(a, cAmbEntTip)) = "Kooperant" And _
                   NzToText(ambData(a, cAmbEnt)) = kooperantID Then
                    Dim aDokID As String
                    aDokID = NzToText(ambData(a, cAmbDokID))
                    If Not otkIdDict.Exists(aDokID) Then          ' samostalno (ne otkup)
                        If IsDate(ambData(a, cAmbDat)) Then
                            Dim aDat As Date
                            aDat = CDate(ambData(a, cAmbDat))
                            If aDat <= datumDo Then
                                Dim aKol As Double
                                aKol = 0
                                If IsNumeric(ambData(a, cAmbKol)) Then aKol = CDbl(ambData(a, cAmbKol))

                                Dim aDelta As Double
                                If NzToText(ambData(a, cAmbSmer)) = "Ulaz" Then
                                    aDelta = aKol            ' OM -> koop (drzi vise)
                                Else
                                    aDelta = -aKol           ' koop -> OM (vratio)
                                End If

                                If aDat < datumOd Then
                                    pocetniSaldoAmb = pocetniSaldoAmb + aDelta
                                    GoTo NextAmbRed
                                End If

                                Dim aTip As String
                                aTip = NzToText(ambData(a, cAmbTip))
                                Dim aLbl As String
                                aLbl = ""
                                If cAmbDokTip > 0 Then aLbl = KarticaAmbDocLabel(NzToText(ambData(a, cAmbDokTip)))
                                Dim aOpis As String
                                aOpis = "Ambala" & ChrW(382) & "a"
                                If aLbl <> "" Then aOpis = aOpis & ": " & aLbl
                                aOpis = aOpis & " (" & aTip & " x " & CStr(CLng(aKol)) & ")"

                                moves.Add Array( _
                                    aDat, _
                                    aDokID, _
                                    "", _
                                    aOpis, _
                                    0#, _
                                    0#, _
                                    aDelta, _
                                    "AMB")
                            End If
                        End If
                    End If
                End If
NextAmbRed:
            Next a
        End If
    End If

    If moves.count = 0 And pocetniSaldo = 0 And pocetniSaldoAmb = 0 Then
        ReportKarticaKooperanta = Empty
        Exit Function
    End If

    ' Prebaci u niz za sortiranje:
    ' 1 Datum, 2 BrojDok, 3 BrojParcele, 4 Opis, 5 Zaduzenje, 6 Razduzenje,
    ' 7 AmbDelta (Izdata - Primljena), 8 RefKljuc
    Dim arr As Variant
    arr = Empty

    If moves.count > 0 Then
        Dim tmpArr() As Variant
        ReDim tmpArr(1 To moves.count, 1 To 8)

        For i = 1 To moves.count
            Dim mv As Variant
            mv = moves(i)
            tmpArr(i, 1) = mv(0)
            tmpArr(i, 2) = mv(1)
            tmpArr(i, 3) = mv(2)
            tmpArr(i, 4) = mv(3)
            tmpArr(i, 5) = mv(4)
            tmpArr(i, 6) = mv(5)
            tmpArr(i, 7) = mv(6)
            tmpArr(i, 8) = mv(7)
        Next i

        ' Sort po datumu, sekundarno po broju dokumenta
        arr = SortArray(tmpArr, 1, True, 2)
    End If

    ' Rezultat: red pocetnog stanja (ako ga ima) + running saldo novca (7) i
    ' ambalaze (8); kol. 9 = ref-kljuc reda za "Detalji otkupa".
    ' GenerateKarticaReport i PrintKarticaPDF citaju kol. 1-8; kol. 9 je skrivena.
    ReportKarticaKooperanta = KarticaRezultatSaPocetnim(arr, pocetniSaldo, pocetniSaldoAmb)
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' REKAPITULACIJA ROBE ZA KARTICU KOOPERANTA
' Zbir kilaze otkupljene robe grupisano po (Vrsta, Sorta, Klasa) za
' kooperanta u periodu. Koristi se kao poseban blok "REKAPITULACIJA ROBE (kg)"
' ispod finansijske kartice u PDF-u (modPrint.FillKarticaSablon). Isti obuhvat
' kao ReportKarticaKooperanta: storno iskljucen, isti datumski opseg, isti
' KooperantID (samo otkup redovi -- oni jedini nose robu).
' Returns: 2D Array (1..N+1, 1..4): 1=Vrsta 2=Sorta 3=Klasa 4=Kg; poslednji
' red = UKUPNO (kol.1="UKUPNO", kol.4 = zbir kg). Empty ako nema robe.
' ============================================================
Public Function ReportKarticaRobaRekap(ByVal kooperantID As String, _
                                       ByVal datumOd As Date, _
                                       ByVal datumDo As Date) As Variant

    Const SRC As String = "modIzvestaj.ReportKarticaRobaRekap"
    On Error GoTo EH

    Dim otkData As Variant
    otkData = GetTableData(TBL_OTKUP)
    If Not IsArray(otkData) Then Exit Function
    otkData = ExcludeStornirano(otkData, TBL_OTKUP)
    If Not IsArray(otkData) Then Exit Function

    Dim cDat As Long, cKoop As Long, cId As Long
    Dim cVrsta As Long, cSorta As Long
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)
    cSorta = GetColumnIndex(TBL_OTKUP, COL_OTK_SORTA)   ' opciono (schema drift -> 0)

    ' Dokumenti kooperanta u periodu: OtkupID -> (vrsta, sorta) sa zaglavlja.
    Dim docs As Object
    Set docs = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(otkData, 1)
        If CStr(otkData(i, cKoop)) = kooperantID Then
            If IsDate(otkData(i, cDat)) Then
                Dim d As Date: d = CDate(otkData(i, cDat))
                If d >= datumOd And d <= datumDo Then
                    Dim sorta As String
                    sorta = ""
                    If cSorta > 0 Then sorta = Trim$(CStr(otkData(i, cSorta)))
                    docs(Trim$(CStr(otkData(i, cId)))) = _
                        Array(Trim$(CStr(otkData(i, cVrsta))), sorta)
                End If
            End If
        End If
    Next i

    ' Klasa i kilaza su na STAVKAMA (REFAKTOR S14.7, kvar 3): red rekapitulacije
    ' je (vrsta, sorta, klasa STAVKE), pa dvoklasni dokument daje dva reda.
    Dim agg As Object
    Set agg = CreateObject("Scripting.Dictionary")

    Dim st As Variant
    st = modOtkup.StavkeOtkupaRedovi()
    If IsArray(st) Then
        For i = 1 To UBound(st, 1)
            If docs.Exists(CStr(st(i, 1))) Then
                Dim vs As Variant
                vs = docs(CStr(st(i, 1)))
                Dim klasa As String, kg As Double
                klasa = CStr(st(i, 3))
                kg = CDbl(st(i, 4))

                Dim key As String: key = CStr(vs(0)) & "|" & CStr(vs(1)) & "|" & klasa
                Dim rec As Variant
                If agg.Exists(key) Then
                    rec = agg(key)
                    rec(3) = CDbl(rec(3)) + kg
                Else
                    rec = Array(CStr(vs(0)), CStr(vs(1)), klasa, kg)
                End If
                agg(key) = rec
            End If
        Next i
    End If

    If agg.count = 0 Then Exit Function

    ' Sortiraj kljuceve (vrsta|sorta|klasa) rastuce -> stabilan, predvidiv prikaz.
    Dim keys() As String
    ReDim keys(0 To agg.count - 1)
    Dim kk As Variant, n As Long
    n = 0
    For Each kk In agg.keys
        keys(n) = CStr(kk): n = n + 1
    Next kk
    Dim a As Long, b As Long, tmp As String
    For a = 0 To UBound(keys) - 1
        For b = a + 1 To UBound(keys)
            If keys(b) < keys(a) Then
                tmp = keys(a): keys(a) = keys(b): keys(b) = tmp
            End If
        Next b
    Next a

    Dim result() As Variant
    ReDim result(1 To agg.count + 1, 1 To 4)
    Dim totKg As Double
    For a = 0 To UBound(keys)
        Dim rr As Variant: rr = agg(keys(a))
        result(a + 1, 1) = CStr(rr(0))
        result(a + 1, 2) = CStr(rr(1))
        result(a + 1, 3) = CStr(rr(2))
        result(a + 1, 4) = CDbl(rr(3))
        totKg = totKg + CDbl(rr(3))
    Next a

    Dim uk As Long: uk = agg.count + 1
    result(uk, 1) = "UKUPNO"
    result(uk, 2) = ""
    result(uk, 3) = ""
    result(uk, 4) = totKg

    ReportKarticaRobaRekap = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Public Function ReportKarticaAmbalaze(ByVal kooperantID As String, _
                                      ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportKarticaAmbalaze"
    On Error GoTo EH
    ' Tok ambalaze kooperanta iz tblAmbalaza (EntitetTip=Kooperant).
    ' Returns: 2D Array (1)=Datum (2)=BrojDok (3)=Opis (4)=Ulaz (5)=Izlaz (6)=Saldo
    ' Smer kanonski (kao GetAmbalazeStanje): Ulaz (+ OM izdao prazne),
    ' Izlaz (- koop predao pune). Saldo (running) = SumaUlaz - SumaIzlaz =
    ' koliko gajbica kooperant drzi/duguje.
    ' RF-06: kretanja PRE datumOd ulaze u red IZV_POCETNO_STANJE, pa saldo vise
    ' ne krece od nule (FM-0028 #1, ista greska kao na novcanoj kartici).

    Dim ambData As Variant
    ambData = GetTableData(TBL_AMBALAZA)
    If Not IsArray(ambData) Then
        ReportKarticaAmbalaze = Empty
        Exit Function
    End If
    ambData = ExcludeStornirano(ambData, TBL_AMBALAZA)
    If Not IsArray(ambData) Then
        ReportKarticaAmbalaze = Empty
        Exit Function
    End If

    Dim colDat As Long, colEnt As Long, colEntTip As Long, colTip As Long
    Dim colKol As Long, colSmer As Long, colDokID As Long
    colDat = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, SRC)
    colEnt = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, SRC)
    colEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, SRC)
    colTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, SRC)
    colKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, SRC)
    colSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, SRC)
    colDokID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, SRC)
    Dim colDokTip As Long
    colDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)   ' opciono (friendly opis)

    ' DokumentID (otkupID) -> BrojDok, jednim prolazom (bez per-row LookupValue).
    Dim brDokDict As Object
    Set brDokDict = BuildOtkupBrojDokDict()

    Dim moves As New Collection
    Dim pocetniSaldo As Double
    pocetniSaldo = 0

    Dim i As Long
    For i = 1 To UBound(ambData, 1)
        If NzToText(ambData(i, colEntTip)) = "Kooperant" And _
           NzToText(ambData(i, colEnt)) = Trim$(kooperantID) Then
            If IsDate(ambData(i, colDat)) Then
                Dim d As Date
                d = CDate(ambData(i, colDat))
                If d <= datumDo Then
                    Dim kol As Double
                    kol = 0
                    If IsNumeric(ambData(i, colKol)) Then kol = CDbl(ambData(i, colKol))

                    Dim ulaz As Double, izlaz As Double
                    ulaz = 0
                    izlaz = 0
                    ' Ledger Smer: "Ulaz" = kooperant dobija (+), inace izlaz (-).
                    If NzToText(ambData(i, colSmer)) = "Ulaz" Then
                        ulaz = kol
                    Else
                        izlaz = kol
                    End If

                    If d < datumOd Then
                        pocetniSaldo = pocetniSaldo + ulaz - izlaz
                        GoTo NextAmbKartRed
                    End If

                    Dim dokID As String
                    dokID = NzToText(ambData(i, colDokID))
                    Dim brojDok As String
                    If brDokDict.Exists(dokID) Then
                        brojDok = CStr(brDokDict(dokID))
                    Else
                        brojDok = dokID
                    End If

                    Dim opis As String
                    opis = NzToText(ambData(i, colTip))   ' TipAmbalaze
                    If colDokTip > 0 Then
                        Dim lbl As String
                        lbl = KarticaAmbDocLabel(NzToText(ambData(i, colDokTip)))
                        If lbl <> "" Then opis = Trim$(opis & " (" & lbl & ")")
                    End If

                    moves.Add Array(d, brojDok, opis, ulaz, izlaz)
                End If
            End If
        End If
NextAmbKartRed:
    Next i

    If moves.count = 0 And pocetniSaldo = 0 Then
        ReportKarticaAmbalaze = Empty
        Exit Function
    End If

    Dim arr As Variant
    arr = Empty

    If moves.count > 0 Then
        Dim tmpArr() As Variant
        ReDim tmpArr(1 To moves.count, 1 To 5)
        For i = 1 To moves.count
            Dim mv As Variant
            mv = moves(i)
            tmpArr(i, 1) = mv(0)
            tmpArr(i, 2) = mv(1)
            tmpArr(i, 3) = mv(2)
            tmpArr(i, 4) = mv(3)
            tmpArr(i, 5) = mv(4)
        Next i
        arr = SortArray(tmpArr, 1, True, 2)
    End If

    ReportKarticaAmbalaze = KarticaAmbRezultatSaPocetnim(arr, pocetniSaldo)
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Mapa: DokumentID (otkupID) -> BrojDok, jednim prolazom kroz tblOtkup.
Private Function BuildOtkupBrojDokDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error GoTo EH

    Dim otk As Variant
    otk = GetTableData(TBL_OTKUP)
    If Not IsArray(otk) Then
        Set BuildOtkupBrojDokDict = dict
        Exit Function
    End If

    Dim colID As Long, colBr As Long
    colID = GetColumnIndex(TBL_OTKUP, COL_OTK_ID)
    colBr = GetColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK)
    If colID = 0 Or colBr = 0 Then
        Set BuildOtkupBrojDokDict = dict
        Exit Function
    End If

    Dim i As Long
    For i = 1 To UBound(otk, 1)
        Dim k As String
        k = NzToText(otk(i, colID))
        If k <> "" Then
            If Not dict.Exists(k) Then dict.Add k, NzToText(otk(i, colBr))
        End If
    Next i

    Set BuildOtkupBrojDokDict = dict
    Exit Function
EH:
    Set BuildOtkupBrojDokDict = dict
End Function

' Friendly oznaka tipa dokumenta za "Pregled ambalaze".
Private Function KarticaAmbDocLabel(ByVal dokTip As String) As String
    Select Case Trim$(dokTip)
        Case DOK_TIP_OTKUP:         KarticaAmbDocLabel = "otkup"
        Case DOK_TIP_OM_IZLAZ_KOOP: KarticaAmbDocLabel = "izdate prazne"
        Case Else:                  KarticaAmbDocLabel = Trim$(dokTip)
    End Select
End Function

' ============================================================
' OTKUPNI LISTOVI (Otkupna mesta) -- sve otkup linije jedne stanice.
' Grain = po OtkupID (linija/klasa), kao kartica; Klasa I/II dele BrDok ali su
' zasebni redovi. Kol. 8 = ref-kljuc "OTK|<OtkupID>" za panel "Detalji otkupa"
' (detalj panel u modScrIzvestaji) i za stampu celog lista po BrDok-u.
' Returns: (1)Datum (2)BrDok (3)Kooperant (4)Vrsta (5)Klasa (6)Kolicina (7)Vrednost (8)RefKljuc
' ============================================================
Public Function ReportOtkupListe(ByVal stanicaID As String, _
                                 ByVal datumOd As Date, _
                                 ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportOtkupListe"
    On Error GoTo EH

    Dim d As Variant
    d = GetTableData(TBL_OTKUP)
    If Not IsArray(d) Then
        ReportOtkupListe = Empty
        Exit Function
    End If
    ' Bez zasebnog ExcludeStornirano -> storno se preskace u glavnoj petlji nize
    ' (izbegnuta jos jedna kopija cele tblOtkup).
    Dim cId As Long, cDat As Long, cBr As Long, cSt As Long, cKoop As Long
    Dim cVr As Long, cStorno As Long
    cStorno = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    cVr = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)

    ' Red liste je DOKUMENT (REFAKTOR S14.7, kvar 3): kolicina, vrednost i klase
    ' su na stavkama, jer ih CreateOtkup_TX na zaglavlju ostavlja prazne. Klase
    ' dokumenta se nabrajaju ("I, II"); linije po klasi nosi detalj reda
    ' (modScrIzvestaji.IzDetaljOtkupLista).
    Dim stavkeZbir As Object
    Set stavkeZbir = modOtkup.ZbirStavkiPoOtkupu()

    ' KooperantID -> "Ime Prezime (ID)" jednim prolazom (bez per-row LookupValue).
    Dim koopDict As Object
    Set koopDict = BuildKooperantNameDict()

    Dim moves As New Collection
    Dim i As Long
    For i = 1 To UBound(d, 1)
        Dim okStorno As Boolean: okStorno = True
        If cStorno > 0 Then okStorno = (NzToText(d(i, cStorno)) <> "Da")   ' VBA Or ne short-circuituje
        If okStorno And NzToText(d(i, cSt)) = Trim$(stanicaID) Then
            If IsDate(d(i, cDat)) Then
                Dim dt As Date
                dt = CDate(d(i, cDat))
                If dt >= datumOd And dt <= datumDo Then
                    Dim koopID As String
                    koopID = NzToText(d(i, cKoop))
                    Dim koopNm As String
                    If koopDict.Exists(koopID) Then
                        koopNm = CStr(koopDict(koopID))
                    Else
                        koopNm = koopID
                    End If

                    Dim kol As Double, vrednost As Double, klase As String
                    kol = 0: vrednost = 0: klase = ""
                    Dim zLst As Variant
                    zLst = modOtkup.ZbirStavkiZaOtkup(stavkeZbir, NzToText(d(i, cId)), _
                               "modIzvestaj.ReportOtkupListe")
                    kol = CDbl(zLst(0))
                    vrednost = CDbl(zLst(1))
                    klase = CStr(zLst(3))

                    moves.Add Array( _
                        dt, _
                        NzToText(d(i, cBr)), _
                        koopNm, _
                        NzToText(d(i, cVr)), _
                        klase, _
                        kol, _
                        vrednost, _
                        "OTK|" & NzToText(d(i, cId)))
                End If
            End If
        End If
    Next i

    If moves.count = 0 Then
        ReportOtkupListe = Empty
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To moves.count, 1 To 8)
    For i = 1 To moves.count
        Dim mv As Variant
        mv = moves(i)
        Dim j As Long
        For j = 0 To 7
            arr(i, j + 1) = mv(j)
        Next j
    Next i
    arr = SortArray(arr, 1, True, 2)   ' po datumu, pa BrDok

    ReportOtkupListe = arr
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Mapa: KooperantID -> "Ime Prezime (ID)", jednim prolazom kroz tblKooperanti.
Private Function BuildKooperantNameDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error GoTo EH

    Dim k As Variant
    k = GetTableData(TBL_KOOPERANTI)
    If Not IsArray(k) Then
        Set BuildKooperantNameDict = dict
        Exit Function
    End If

    Dim cId As Long, cIme As Long, cPr As Long
    cId = GetColumnIndex(TBL_KOOPERANTI, COL_KOOP_ID)
    cIme = GetColumnIndex(TBL_KOOPERANTI, "Ime")
    cPr = GetColumnIndex(TBL_KOOPERANTI, "Prezime")
    If cId = 0 Then
        Set BuildKooperantNameDict = dict
        Exit Function
    End If

    Dim i As Long
    For i = 1 To UBound(k, 1)
        Dim id As String
        id = NzToText(k(i, cId))
        If id <> "" Then
            Dim nm As String
            nm = ""
            If cIme > 0 Then nm = NzToText(k(i, cIme))
            If cPr > 0 Then nm = Trim$(nm & " " & NzToText(k(i, cPr)))
            If nm = "" Then nm = id Else nm = nm & " (" & id & ")"
            If Not dict.Exists(id) Then dict.Add id, nm
        End If
    Next i

    Set BuildKooperantNameDict = dict
    Exit Function
EH:
    Set BuildKooperantNameDict = dict
End Function

Public Sub PrintKarticaPDF(ByVal kooperantID As String, _
                           ByVal datumOd As Date, ByVal datumDo As Date)
                           
    Const SRC As String = "modIzvestaj.PrintKarticaPDF"
    On Error GoTo EH

    Dim data As Variant
    data = ReportKarticaKooperanta(kooperantID, datumOd, datumDo)
    If IsEmpty(data) Then
        Err.Raise vbObjectError + 7502, SRC, _
                  "Nema podataka za ovog kooperanta."
    End If

    Dim ime As String, prezime As String, bpg As String
    ime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime"))
    prezime = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"))
    bpg = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, COL_KOOP_BPG))

    Dim koopNaziv As String
    koopNaziv = ime & " " & prezime & " (" & kooperantID & ")"
    Dim period As String
    period = Format$(datumOd, "DD.MM.YYYY") & " - " & Format$(datumDo, "DD.MM.YYYY")

    ' Rekapitulacija robe (kg) po vrsti/sorti/klasi -- poseban blok ispod kartice.
    Dim rekap As Variant
    rekap = ReportKarticaRobaRekap(kooperantID, datumOd, datumDo)

    Dim ws As Worksheet
    Set ws = FillKarticaSablon(koopNaziv, bpg, period, data, rekap)
    If ws Is Nothing Then Exit Sub

    Dim pdfPath As String
    pdfPath = EnsureDocFolder(PDF_DIR_KARTICE) & "\Kartica_" & Replace(kooperantID, "-", "") & "_" & _
              Format$(datumOd, "YYYYMMDD") & "-" & Format$(datumDo, "YYYYMMDD") & ".pdf"

    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_KARTICA_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            DocExportPdf ws, pdfPath, True
        ' OFF -> bez izlaza
    End Select
    Exit Sub

EH:
    Application.ScreenUpdating = True
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Sub

' ============================================================
' KARTICA AMBALAZE (PDF) -- pandan PrintKarticaPDF za tab "Pregled ambalaze".
' Bez "KarticaSablon" templejta (on je finansijski): layout se gradi u kodu na
' posvecenom skrivenom sheetu (_KartAmbPrint), pa export u PDF (otvara po zavrsetku).
' Podaci iz ReportKarticaAmbalaze (6 kol: Datum, BrojDok, Opis, Ulaz, Izlaz, Saldo;
' poslednji red = UKUPNO). Gajbe = ceo broj.
' ============================================================
Public Sub PrintKarticaAmbalazePDF(ByVal kooperantID As String, _
                                   ByVal datumOd As Date, ByVal datumDo As Date)

    Const SRC As String = "modIzvestaj.PrintKarticaAmbalazePDF"
    Const NUM_COLS As Long = 6   ' Datum, BrojDok, Opis, Ulaz, Izlaz, Saldo
    On Error GoTo EH

    Dim data As Variant
    data = ReportKarticaAmbalaze(kooperantID, datumOd, datumDo)
    If IsEmpty(data) Then
        Err.Raise vbObjectError + 7502, SRC, _
                  "Nema podataka o ambalazi za ovog kooperanta."
    End If

    Dim ime As String, prezime As String
    ime = NzToText(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Ime"))
    prezime = NzToText(LookupValue(TBL_KOOPERANTI, "KooperantID", kooperantID, "Prezime"))
    Dim koopNaziv As String
    koopNaziv = Trim$(ime & " " & prezime) & " (" & kooperantID & ")"
    Dim period As String
    period = Format$(datumOd, "DD.MM.YYYY") & " - " & Format$(datumDo, "DD.MM.YYYY")

    Dim ws As Worksheet
    Set ws = FillKarticaAmbalazeSablon(koopNaziv, period, data)
    If ws Is Nothing Then Exit Sub

    Dim pdfPath As String
    pdfPath = EnsureDocFolder(PDF_DIR_KARTICE) & "\KarticaAmbalaze_" & Replace(kooperantID, "-", "") & "_" & _
              Format$(datumOd, "YYYYMMDD") & "-" & Format$(datumDo, "YYYYMMDD") & ".pdf"

    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_KARTICA_AMB_PRINT_MODE), "PDF")
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case "PDF"
            DocExportPdf ws, pdfPath, True
        ' OFF -> bez izlaza
    End Select
    Exit Sub

EH:
    Application.ScreenUpdating = True
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Sub

' ============================================================
' ZBIRNI OBLICI PO STANICAMA (krug 9 -- "fali sadrzaj za zbirne")
' Red = stanica, kolone = UKUPNO red pojedinacnog izvestaja te stanice --
' isti racun, nijedno pravilo se ne prepisuje. Stanica ciji su svi brojevi
' nula se preskace (sum bez prometa je red-shum), ali stanica sa saldom bez
' prometa perioda OSTAJE. Poslednji red = UKUPNO preko svih stanica.
' ============================================================
' (1)=StanicaID (2)=Naziv (3)=Kg (4)=Vrednost (5)=Isplaceno (6)=Agro
' (7)=Saldo (8)=Amb -- kolone 3..8 su UKUPNO red (2..7) pojedinacnog
' ReportSaldoOM te stanice (agro PRIPISAN stanicama ucestvuje u saldu;
' "nerasporedjena" agro linija je i tamo van UKUPNO pa je nema ni ovde).
Public Function ReportSaldoOMZbirni(ByVal datumOd As Date, _
                                    ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportSaldoOMZbirni"
    On Error GoTo EH

    ' Univerzum stanica IZ PODATAKA (krug 16) -- sifarnik samo imenuje.
    Dim st As Variant
    st = IzvStaniceIzPodataka()
    If Not IsArray(st) Then Exit Function

    Dim outA() As Variant, n As Long, i As Long, j As Long
    Dim r As Variant, uk As Long, imaSta As Boolean
    Dim tot(3 To 8) As Double
    ReDim outA(1 To UBound(st, 1) + 1, 1 To 8)
    For i = 1 To UBound(st, 1)
        Dim stID As String
        stID = Trim$(CStr(st(i, 1)))
        If Len(stID) > 0 Then
            r = ReportSaldoOM(stID, datumOd, datumDo)
            uk = IzvUkupnoRed(r, 1)
            If uk > 0 Then
                imaSta = False
                For j = 2 To 7
                    If IzvNum(r(uk, j)) <> 0 Then imaSta = True
                Next j
                If imaSta Then
                    n = n + 1
                    outA(n, 1) = stID
                    outA(n, 2) = CStr(st(i, 2))
                    For j = 3 To 8
                        outA(n, j) = IzvNum(r(uk, j - 1))
                        tot(j) = tot(j) + IzvNum(outA(n, j))
                    Next j
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function

    outA(n + 1, 2) = "UKUPNO"
    For j = 3 To 8
        outA(n + 1, j) = tot(j)
    Next j
    ReportSaldoOMZbirni = IzvIseciRedove(outA, n + 1, 8)
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' (1)=StanicaID (2)=Naziv (3)=Kes (4)=VirmanFirma (5)=VirmanAvans (6)=Ukupno
Public Function ReportIsplataZbirniOM(ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportIsplataZbirniOM"
    On Error GoTo EH

    ' Univerzum stanica IZ PODATAKA (krug 16) -- sifarnik samo imenuje.
    Dim st As Variant
    st = IzvStaniceIzPodataka()
    If Not IsArray(st) Then Exit Function

    Dim outA() As Variant, n As Long, i As Long, j As Long
    Dim r As Variant, uk As Long
    Dim tot(3 To 6) As Double
    ReDim outA(1 To UBound(st, 1) + 1, 1 To 6)
    For i = 1 To UBound(st, 1)
        Dim stID As String
        stID = Trim$(CStr(st(i, 1)))
        If Len(stID) > 0 Then
            r = ReportIsplata("OM", stID, datumOd, datumDo)
            uk = IzvUkupnoRed(r, 1)
            If uk > 0 Then
                If IzvNum(r(uk, 5)) <> 0 Then
                    n = n + 1
                    outA(n, 1) = stID
                    outA(n, 2) = CStr(st(i, 2))
                    For j = 3 To 6
                        outA(n, j) = IzvNum(r(uk, j - 1))
                        tot(j) = tot(j) + IzvNum(outA(n, j))
                    Next j
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function

    outA(n + 1, 2) = "UKUPNO"
    For j = 3 To 6
        outA(n + 1, j) = tot(j)
    Next j
    ReportIsplataZbirniOM = IzvIseciRedove(outA, n + 1, 6)
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Zbirno po KUPCIMA (krug 11 -- "fale salda po kupcima"): isti obrazac
' kao stanice. (1)=KupacID (2)=Naziv (3)=Kg (4)=Vrednost (5)=Uplaceno
' (6)=Saldo (7)=Amb -- iz UKUPNO reda ReportSaldoKupci (kolona 3, cena,
' je prosek pa se u zbir ne prenosi).
Public Function ReportSaldoKupciZbirni(ByVal datumOd As Date, _
                                       ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportSaldoKupciZbirni"
    On Error GoTo EH

    ' Kupci iz PODATAKA (distinct po prijemnicama), ne iz sifarnika: kupac
    ' sa prometom a bez reda u tblKupci mora da se vidi; naziv iz sifarnika
    ' uz fallback na ID (IzvKupciIzPodataka).
    Dim ku As Variant
    ku = IzvKupciIzPodataka()
    If Not IsArray(ku) Then Exit Function

    Dim outA() As Variant, n As Long, i As Long, j As Long
    Dim r As Variant, uk As Long, imaSta As Boolean
    Dim srcKol As Variant, tot(3 To 7) As Double
    srcKol = Array(0, 0, 0, 2, 4, 5, 6, 7)   ' out kolona j <- pojedinacna srcKol(j)
    ReDim outA(1 To UBound(ku, 1) + 1, 1 To 7)
    For i = 1 To UBound(ku, 1)
        Dim kuID As String
        kuID = Trim$(CStr(ku(i, 1)))
        If Len(kuID) > 0 Then
            r = ReportSaldoKupci(kuID, datumOd, datumDo)
            uk = IzvUkupnoRed(r, 1)
            If uk > 0 Then
                imaSta = False
                For j = 3 To 7
                    If IzvNum(r(uk, srcKol(j))) <> 0 Then imaSta = True
                Next j
                If imaSta Then
                    n = n + 1
                    outA(n, 1) = kuID
                    outA(n, 2) = CStr(ku(i, 2))
                    For j = 3 To 7
                        outA(n, j) = IzvNum(r(uk, srcKol(j)))
                        tot(j) = tot(j) + IzvNum(outA(n, j))
                    Next j
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function

    outA(n + 1, 2) = "UKUPNO"
    For j = 3 To 7
        outA(n + 1, j) = tot(j)
    Next j
    ReportSaldoKupciZbirni = IzvIseciRedove(outA, n + 1, 7)
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Roba po kupcu (krug 11): (1)=KupacID (2)=Naziv (3)=Kg (4)=Vrednost --
' UKUPNO red kupcevog agregata ReportOtkupRoba("Kupac") preko svih vrsta.
Public Function ReportRobaKupciZbirni(ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportRobaKupciZbirni"
    On Error GoTo EH

    Dim ku As Variant
    ku = IzvKupciIzPodataka()
    If Not IsArray(ku) Then Exit Function

    Dim outA() As Variant, n As Long, i As Long
    Dim r As Variant, uk As Long
    Dim totKg As Double, totVr As Double
    ReDim outA(1 To UBound(ku, 1) + 1, 1 To 4)
    For i = 1 To UBound(ku, 1)
        Dim kuID2 As String
        kuID2 = Trim$(CStr(ku(i, 1)))
        If Len(kuID2) > 0 Then
            r = ReportOtkupRoba("Kupac", kuID2, datumOd, datumDo)
            uk = IzvUkupnoRed(r, 2)
            If uk > 0 Then
                If IzvNum(r(uk, 3)) <> 0 Or IzvNum(r(uk, 4)) <> 0 Then
                    n = n + 1
                    outA(n, 1) = kuID2
                    outA(n, 2) = CStr(ku(i, 2))
                    outA(n, 3) = IzvNum(r(uk, 3))
                    outA(n, 4) = IzvNum(r(uk, 4))
                    totKg = totKg + IzvNum(outA(n, 3))
                    totVr = totVr + IzvNum(outA(n, 4))
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function

    outA(n + 1, 2) = "UKUPNO"
    outA(n + 1, 3) = totKg
    outA(n + 1, 4) = totVr
    ReportRobaKupciZbirni = IzvIseciRedove(outA, n + 1, 4)
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Roba po OM zbirno (krug 12): kg i vrednost su TACNO kolone 3 i 4
' zbirnog salda po stanicama -- projekcija istog izvora, ne drugi racun
' ("tu realno idu podaci o robi koji su vec u saldu").
' (1)=StanicaID (2)=Naziv (3)=Kg (4)=Vrednost; poslednji red = UKUPNO.
Public Function ReportRobaOMZbirni(ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportRobaOMZbirni"
    On Error GoTo EH
    Dim s As Variant, outA() As Variant, i As Long, j As Long
    s = ReportSaldoOMZbirni(datumOd, datumDo)
    If Not IsArray(s) Then Exit Function
    ReDim outA(1 To UBound(s, 1), 1 To 4)
    For i = 1 To UBound(s, 1)
        For j = 1 To 4
            outA(i, j) = s(i, j)
        Next j
    Next i
    ReportRobaOMZbirni = outA
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Roba po VOZACU zbirno (krug 12): OTPREMLJENO -- kg i vrednost IZDATIH,
' nestorniranih otpremnica u opsegu, po vozacu; naziv iz tblVozaci sa
' fallback-om na ID. (1)=VozacID (2)=Naziv (3)=Kg (4)=Vrednost; UKUPNO.
Public Function ReportRobaVozaciZbirni(ByVal datumOd As Date, _
                                       ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportRobaVozaciZbirni"
    On Error GoTo EH

    ' OTPREMLJENO = IZDATO (review #362, P1). Nacrt je najava: nema izvora, gajbe
    ' nisu knjizene, i sme da ostane neizdat -- u otpremljenu robu ne ulazi.
    ' Kilaza dolazi sa STAVKI (zaglavlje je od S3a nema). Vrednost je vrednost
    ' IZVORNIH otkupa (Kolicina x Cena njihovih stavki), NE Kolicina x
    ' PredlogCena: predlog je polje za prefill, ne finansijska cinjenica.
    Dim d As Variant, i As Long
    Dim cVoz As Long, cId As Long, cDat As Long, cStorno As Long, cIzd As Long
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(d) Then Exit Function
    cVoz = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    cId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    cDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    cIzd = RequireColumnIndex(TBL_OTPREMNICA, COL_TRACE_IZDATO_STATUS, SRC)
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)

    Dim zbir As Object, vredIzv As Object
    Set zbir = modDokumenta.ZbirStavkiPoOtpremnici()
    Set vredIzv = modDokumenta.VrednostIzvoraPoOtpremnici()

    Dim kg As Object, vr As Object, k As String, dv As Date, z As Variant
    Set kg = CreateObject("Scripting.Dictionary")
    Set vr = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(d, 1)
        If (cStorno = 0 Or CStr(d(i, cStorno)) <> "Da") And _
           modDokumenta.IzdatoStatusJeIzdato(d(i, cIzd)) Then
            If IsDate(d(i, cDat)) Then
                dv = CDate(d(i, cDat))
                If dv >= datumOd And dv <= datumDo Then
                    k = Trim$(CStr(d(i, cVoz)))
                    If Len(k) > 0 Then
                        z = modDokumenta.ZbirStavkiZaOtpremnicu(zbir, _
                                Trim$(NzToText(d(i, cId))), SRC)
                        kg(k) = IzvNum(kg(k)) + CDbl(z(0))
                        vr(k) = IzvNum(vr(k)) + modDokumenta.VrednostIzvoraZaOtpremnicu( _
                                    vredIzv, Trim$(NzToText(d(i, cId))), SRC)
                    End If
                End If
            End If
        End If
    Next i
    If kg.count = 0 Then Exit Function

    Dim outA() As Variant, kk As Variant, n As Long, nm As String
    Dim totKg As Double, totVr As Double
    ReDim outA(1 To kg.count + 1, 1 To 4)
    For Each kk In kg.keys
        n = n + 1
        outA(n, 1) = CStr(kk)
        nm = ""
        On Error Resume Next
        nm = Trim$(Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", CStr(kk), "Ime"))) & _
                   " " & Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", CStr(kk), "Prezime"))))
        On Error GoTo EH
        outA(n, 2) = IIf(Len(nm) > 0, nm, CStr(kk))
        outA(n, 3) = IzvNum(kg(kk))
        outA(n, 4) = IzvNum(vr(kk))
        totKg = totKg + IzvNum(outA(n, 3))
        totVr = totVr + IzvNum(outA(n, 4))
    Next kk
    outA(n + 1, 2) = "UKUPNO"
    outA(n + 1, 3) = totKg
    outA(n + 1, 4) = totVr
    ReportRobaVozaciZbirni = outA
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Zbirna ambalaza preko SVIH entiteta tipa (krug 14: "sumarno stanje po
' tipu za svakog vozaca... dropdown je besmislen"): red = entitet x tip
' gajbe. Entiteti dolaze IZ PODATAKA (distinct po nestorniranom ledgeru,
' uz isti DOK_TIP_OTKUP izuzetak za vozace kao ReportAmbalaza); po
' entitetu se zove POSTOJECI legacy zbirni racun (ReportAmbalaza sa
' zbirni=True) -- smerovi/isVozac pravila se ne prepisuju.
' (1)=EntID (2)=EntNaziv (3)=Tip (4)=Ulaz (5)=Izlaz; UKUPNO u koloni 2.
Public Function ReportAmbalazaZbirnoSvi(ByVal entitetTip As String, _
                                        ByVal datumOd As Date, _
                                        ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportAmbalazaZbirnoSvi"
    On Error GoTo EH

    Dim d As Variant, i As Long
    Dim cEnt As Long, cEntTip As Long, cVoz As Long, cDokTip As Long, cStorno As Long
    d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Function
    cEnt = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, SRC)
    cEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, SRC)
    cVoz = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC, SRC)
    cDokTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, SRC)
    cStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)

    Dim ents As Object, k As String
    Set ents = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(d, 1)
        If cStorno = 0 Or CStr(d(i, cStorno)) <> "Da" Then
            k = ""
            Select Case entitetTip
                Case "OM"
                    If CStr(d(i, cEntTip)) = "Stanica" Then k = Trim$(CStr(d(i, cEnt)))
                Case "Kupac"
                    If CStr(d(i, cEntTip)) = "Kupac" Then k = Trim$(CStr(d(i, cEnt)))
                Case "Vozac"
                    If CStr(d(i, cDokTip)) <> DOK_TIP_OTKUP Then k = Trim$(CStr(d(i, cVoz)))
            End Select
            If Len(k) > 0 Then ents(k) = True
        End If
    Next i
    If ents.count = 0 Then Exit Function

    Dim linije As Collection, kk As Variant, r As Variant
    Dim nm As String, totU As Double, totI As Double
    Set linije = New Collection
    For Each kk In ents.keys
        r = ReportAmbalaza(entitetTip, CStr(kk), datumOd, datumDo, True)
        If IsArray(r) Then
            nm = IzvEntNaziv(entitetTip, CStr(kk))
            For i = 1 To UBound(r, 1)
                If CStr(r(i, 1)) <> "UKUPNO" Then
                    linije.Add Array(CStr(kk), nm, CStr(r(i, 1)), _
                                     IzvNum(r(i, 5)), IzvNum(r(i, 6)))
                    totU = totU + IzvNum(r(i, 5))
                    totI = totI + IzvNum(r(i, 6))
                End If
            Next i
        End If
    Next kk
    If linije.count = 0 Then Exit Function

    Dim outA() As Variant, n As Long, red As Variant
    ReDim outA(1 To linije.count + 1, 1 To 5)
    For n = 1 To linije.count
        red = linije(n)
        For i = 0 To 4
            outA(n, i + 1) = red(i)
        Next i
    Next n
    outA(linije.count + 1, 2) = "UKUPNO"
    outA(linije.count + 1, 4) = totU
    outA(linije.count + 1, 5) = totI
    ReportAmbalazaZbirnoSvi = outA
    Exit Function
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Naziv entiteta za zbirne redove: sifarnik sa fallback-om na ID.
Private Function IzvEntNaziv(ByVal entitetTip As String, ByVal iD As String) As String
    Dim nm As String
    On Error Resume Next
    Select Case entitetTip
        Case "OM"
            nm = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", iD, "Naziv")))
        Case "Kupac"
            nm = Trim$(CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, iD, COL_KUP_NAZIV)))
        Case "Vozac"
            nm = Trim$(Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", iD, "Ime"))) & _
                       " " & Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", iD, "Prezime"))))
    End Select
    On Error GoTo 0
    IzvEntNaziv = IIf(Len(nm) > 0, nm, iD)
End Function

' Distinct STANICE iz podataka (recenzija #245, krug 16): union StanicaID
' iz tblOtkup + OMID iz tblNovac + Stanica-entiteta iz tblAmbalaza
' (nestornirano). Sifarnik daje samo ime (fallback ID) -- stanica sa
' prometom a bez reda u tblStanice NE SME tiho da ispadne iz "Svi OM"
' zbirova (silent omission je gori od ruznog ID-a).
' 2D (1..n, 1..2): 1=StanicaID, 2=naziv.
Private Function IzvStaniceIzPodataka() As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    IzvStaniceUnion dict, TBL_OTKUP, COL_OTK_STANICA, "", ""
    IzvStaniceUnion dict, TBL_NOVAC, COL_NOV_OM_ID, "", ""
    IzvStaniceUnion dict, TBL_AMBALAZA, COL_AMB_ENTITET, COL_AMB_ENTITET_TIP, "Stanica"
    If dict.count = 0 Then Exit Function

    Dim outA() As Variant, kk As Variant, n As Long, nm As String
    ReDim outA(1 To dict.count, 1 To 2)
    For Each kk In dict.keys
        n = n + 1
        outA(n, 1) = CStr(kk)
        nm = ""
        On Error Resume Next
        nm = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", CStr(kk), "Naziv")))
        On Error GoTo 0
        outA(n, 2) = IIf(Len(nm) > 0, nm, CStr(kk))
    Next kk
    IzvStaniceIzPodataka = outA
End Function

' Dodaj distinct vrednosti kolone (nestornirano; uz opcioni filter druge
' kolone) u dict -- pomocna za IzvStaniceIzPodataka.
Private Sub IzvStaniceUnion(ByVal dict As Object, ByVal tblName As String, _
                            ByVal kolona As String, ByVal filtKol As String, _
                            ByVal filtVal As String)
    Dim d As Variant, i As Long, c As Long, cF As Long, cStorno As Long, k As String
    ' Fail-visible (recenzija #245): obavezna ID/filter kolona koja fali =
    ' greska, ne tihi nepotpun univerzum finansijskog zbira.
    d = GetTableData(tblName)
    If Not IsArray(d) Then Exit Sub
    c = RequireColumnIndex(tblName, kolona, "modIzvestaj.IzvStaniceUnion")
    cStorno = GetColumnIndex(tblName, COL_STORNIRANO)
    cF = 0
    If Len(filtKol) > 0 Then cF = RequireColumnIndex(tblName, filtKol, "modIzvestaj.IzvStaniceUnion")
    ' VBA Or NEMA kratki spoj: "cF = 0 Or d(i, cF)" evaluira i d(i, 0) i
    ' puca -- zato ugnjezdeni uslovi (greska je do kruga 17 bila gutana
    ' starim On Error Resume Next, a Resume-Next je slucajno ulazio u telo).
    For i = 1 To UBound(d, 1)
        If cStorno > 0 Then
            If CStr(d(i, cStorno)) = "Da" Then GoTo Sledeci
        End If
        If cF > 0 Then
            If CStr(d(i, cF)) <> filtVal Then GoTo Sledeci
        End If
        k = Trim$(CStr(d(i, c)))
        If Len(k) > 0 Then dict(k) = True
Sledeci:
    Next i
End Sub

' Distinct kupci IZ PODATAKA (nestornirane prijemnice), 2D (1..n, 1..2):
' 1=KupacID, 2=naziv iz tblKupci sa fallback-om na ID. Sifarnik nije izvor
' spiska -- kupac sa prometom bez reda u tblKupci mora da se vidi (fixture
' to namerno drzi tako).
Private Function IzvKupciIzPodataka() As Variant
    Dim d As Variant, i As Long, cKup As Long, cStorno As Long
    Dim dict As Object, k As String, nm As String
    d = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(d) Then Exit Function
    cKup = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, "modIzvestaj.IzvKupciIzPodataka")
    cStorno = GetColumnIndex(TBL_PRIJEMNICA, COL_STORNIRANO)
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(d, 1)
        If cStorno = 0 Or CStr(d(i, cStorno)) <> "Da" Then
            k = Trim$(CStr(d(i, cKup)))
            If Len(k) > 0 Then dict(k) = True
        End If
    Next i
    If dict.count = 0 Then Exit Function

    Dim outA() As Variant, kk As Variant, n As Long
    ReDim outA(1 To dict.count, 1 To 2)
    For Each kk In dict.keys
        n = n + 1
        outA(n, 1) = CStr(kk)
        nm = ""
        On Error Resume Next
        nm = Trim$(CStr(LookupValue(TBL_KUPCI, COL_KUP_ID, CStr(kk), COL_KUP_NAZIV)))
        On Error GoTo 0
        outA(n, 2) = IIf(Len(nm) > 0, nm, CStr(kk))
    Next kk
    IzvKupciIzPodataka = outA
End Function

' Bezbedan broj (Empty/tekst -> 0) -- lokalni pandan NumVal-a iz
' modOtkupBlok (tamo je Private, odavde nevidljiv).
Private Function IzvNum(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then IzvNum = CDbl(v)
End Function

' Indeks reda "UKUPNO" u koloni k (0 = nema ga).
Private Function IzvUkupnoRed(ByVal r As Variant, ByVal k As Long) As Long
    Dim i As Long
    If IsEmpty(r) Or Not IsArray(r) Then Exit Function
    For i = UBound(r, 1) To 1 Step -1
        If CStr(r(i, k)) = "UKUPNO" Then
            IzvUkupnoRed = i
            Exit Function
        End If
    Next i
End Function

' Prvih n redova 2D niza (petlja po stanicama alocira za sve, popuni manje).
Private Function IzvIseciRedove(ByRef a As Variant, ByVal n As Long, _
                                ByVal nCols As Long) As Variant
    Dim outA() As Variant, i As Long, j As Long
    ReDim outA(1 To n, 1 To nCols)
    For i = 1 To n
        For j = 1 To nCols
            outA(i, j) = a(i, j)
        Next j
    Next i
    IzvIseciRedove = outA
End Function

' ============================================================
' KUPCI
' ============================================================
' Otkupljena roba za kupca kao LISTA PRIJEMNICA (smoke krug 4) -- ne agregat
' po vrsti: operater trazi dokumenta, agregat vec daje tab Zbirni. Izvor je
' GetPrijemniceByKupac (isti read-model kao korpa fakturisanja), ovde samo
' normalizovan u fiksne kolone nezavisne od rasporeda u tabeli (schema drift):
' (1)=Datum (2)=BrojPrijemnice (3)=BrojZbirne (4)=Vrsta (5)=Klasa
' (6)=Kg (7)=Cena (8)=Vrednost=kg*cena (9)=PrijemnicaID (10)=Sorta.
' Poslednji red = UKUPNO (kolona 2), kao ostali Report*.
Public Function ReportPrijemniceKupca(ByVal kupacID As String, _
                                      ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportPrijemniceKupca"
    On Error GoTo EH

    Dim data As Variant
    data = GetPrijemniceByKupac(kupacID, datumOd, datumDo, False)
    If IsEmpty(data) Or Not IsArray(data) Then Exit Function

    Dim cDat As Long, cBr As Long, cZb As Long, cVr As Long, cKl As Long
    Dim cKol As Long, cCe As Long, cId As Long
    cDat = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, SRC)
    cBr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)
    cZb = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
    cVr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA, SRC)
    cKl = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA, SRC)
    cKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, SRC)
    cCe = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, SRC)
    cId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
    Dim cSor As Long
    cSor = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_SORTA, SRC)

    Dim n As Long, i As Long, kg As Double, cena As Double
    Dim totKg As Double, totVr As Double
    n = UBound(data, 1)
    Dim result() As Variant
    ReDim result(1 To n + 1, 1 To 10)
    For i = 1 To n
        kg = 0: cena = 0
        If IsNumeric(data(i, cKol)) Then kg = CDbl(data(i, cKol))
        If IsNumeric(data(i, cCe)) Then cena = CDbl(data(i, cCe))
        result(i, 1) = data(i, cDat)
        result(i, 2) = Trim$(CStr(data(i, cBr)))
        result(i, 3) = Trim$(CStr(data(i, cZb)))
        result(i, 4) = Trim$(CStr(data(i, cVr)))
        result(i, 5) = Trim$(CStr(data(i, cKl)))
        result(i, 6) = kg
        result(i, 7) = cena
        result(i, 8) = kg * cena
        result(i, 9) = Trim$(CStr(data(i, cId)))
        result(i, 10) = Trim$(CStr(data(i, cSor)))
        totKg = totKg + kg
        totVr = totVr + kg * cena
    Next i
    result(n + 1, 2) = "UKUPNO"
    result(n + 1, 6) = totKg
    result(n + 1, 8) = totVr

    ReportPrijemniceKupca = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Public Function ReportSaldoKupci(ByVal kupacID As String, _
                                 ByVal datumOd As Date, _
                                 ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportSaldoKupci"
    On Error GoTo EH
    
    ' Returns: 2D Array (Vrsta, Kolicina, Cena, Vrednost, Novac, Saldo, Ambalaza)
    ' Letzte Zeile = UKUPNO
    '
    ' Napomena:
    ' Ne izlazimo ako nema prijemnica, jer kupac moze imati uplatu/avans
    ' bez robe u periodu. Takav novac mora biti vidljiv u saldu.
    
    Dim prijData As Variant
    prijData = GetPrijemniceByKupac(kupacID, datumOd, datumDo)
    
    ' --- Prijemnice pro VrstaVoca aggregieren ---
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim vals As Variant
    
    If Not IsEmpty(prijData) Then
        If IsArray(prijData) Then
            prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
            
            If IsArray(prijData) Then
                Dim colVrsta As Long, colKol As Long, colCena As Long, colAmb As Long
                colVrsta = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA, "modIzvestaj.ReportSaldoKupci")
                colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "modIzvestaj.ReportSaldoKupci")
                colCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, "modIzvestaj.ReportSaldoKupci")
                colAmb = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB, "modIzvestaj.ReportSaldoKupci")
                
                For i = 1 To UBound(prijData, 1)
                    Dim key As String
                    key = CStr(prijData(i, colVrsta))
                    If key = "" Then key = "(Nepoznato)"
                    
                    If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#, 0#)
                    
                    vals = dict(key)
                    
                    If IsNumeric(prijData(i, colKol)) Then vals(0) = vals(0) + CDbl(prijData(i, colKol))
                    If IsNumeric(prijData(i, colCena)) Then vals(1) = CDbl(prijData(i, colCena))
                    If IsNumeric(prijData(i, colKol)) And IsNumeric(prijData(i, colCena)) Then
                        vals(2) = vals(2) + CDbl(prijData(i, colKol)) * CDbl(prijData(i, colCena))
                    End If
                    If IsNumeric(prijData(i, colAmb)) Then vals(3) = vals(3) + CLng(prijData(i, colAmb))
                    
                    dict(key) = vals
                Next i
            End If
        End If
    End If
    
    ' --- Novac pro Vrsta ---
    Dim novacDict As Object
    Set novacDict = GetUplataByVrsta(kupacID, datumOd, datumDo)
    
    If dict.count = 0 And novacDict.count = 0 Then
        ReportSaldoKupci = Empty
        Exit Function
    End If
    
    ' --- Gesamt-Novac (fuer UKUPNO Saldo) ---
    Dim novacTotal As Double
    Dim novacData As Variant
    novacData = GetTableData(TBL_NOVAC)
    
    If IsArray(novacData) Then
        novacData = ExcludeStornirano(novacData, TBL_NOVAC)
    End If
    
    If IsArray(novacData) Then
        Dim colNovPartnerID As Long, colNovUplata As Long, colNovDatum As Long
        colNovPartnerID = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID, "modIzvestaj.ReportSaldoKupci")
        colNovUplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_UPLATA, "modIzvestaj.ReportSaldoKupci")
        colNovDatum = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, "modIzvestaj.ReportSaldoKupci")
        
        Dim n As Long
        For n = 1 To UBound(novacData, 1)
            If CStr(novacData(n, colNovPartnerID)) = kupacID Then
                If IsDate(novacData(n, colNovDatum)) Then
                    If CDate(novacData(n, colNovDatum)) >= datumOd And _
                       CDate(novacData(n, colNovDatum)) <= datumDo Then
                        If IsNumeric(novacData(n, colNovUplata)) Then
                            novacTotal = novacTotal + CDbl(novacData(n, colNovUplata))
                        End If
                    End If
                End If
            End If
        Next n
    End If
    
    ' --- Novac-only vrste: novac postoji, ali nema prijemnice za tu vrstu ---
    Dim novacOnlyCount As Long
    Dim novacKeys As Variant
    
    If novacDict.count > 0 Then
        novacKeys = novacDict.keys
        
        For i = 0 To novacDict.count - 1
            Dim novacKey As String
            novacKey = CStr(novacKeys(i))
            
            If Not dict.Exists(novacKey) Then
                novacOnlyCount = novacOnlyCount + 1
            End If
        Next i
    End If
    
    ' --- Ergebnis-Array ---
    Dim rowCount As Long
    rowCount = dict.count + novacOnlyCount + 1  ' +1 UKUPNO
    
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 7)
    
    Dim keys As Variant
    Dim totKol As Double, totVr As Double, totNov As Double, totAmb As Long
    Dim idx As Long
    
    If dict.count > 0 Then
        keys = dict.keys
        
        For i = 0 To dict.count - 1
            idx = idx + 1
            vals = dict(keys(i))
            
            Dim novacVrsta As Double
            novacVrsta = 0
            If novacDict.Exists(keys(i)) Then novacVrsta = CDbl(novacDict(keys(i)))
            
            result(idx, 1) = keys(i)              ' Vrsta
            result(idx, 2) = vals(0)              ' Kolicina
            result(idx, 3) = vals(1)              ' Cena (letzte)
            result(idx, 4) = vals(2)              ' Vrednost
            result(idx, 5) = novacVrsta           ' Novac pro Vrsta
            result(idx, 6) = vals(2) - novacVrsta ' Saldo pro Vrsta
            result(idx, 7) = vals(3)              ' Ambalaza
            
            totKol = totKol + vals(0)
            totVr = totVr + vals(2)
            totNov = totNov + novacVrsta
            totAmb = totAmb + vals(3)
        Next i
    End If
    
    ' Novac-only redovi
    If novacDict.count > 0 Then
        novacKeys = novacDict.keys
        
        For i = 0 To novacDict.count - 1
            novacKey = CStr(novacKeys(i))
            
            If Not dict.Exists(novacKey) Then
                idx = idx + 1
                
                result(idx, 1) = novacKey
                result(idx, 2) = ""
                result(idx, 3) = ""
                result(idx, 4) = ""
                result(idx, 5) = CDbl(novacDict(novacKey))
                result(idx, 6) = 0 - CDbl(novacDict(novacKey))
                result(idx, 7) = ""
                
                totNov = totNov + CDbl(novacDict(novacKey))
            End If
        Next i
    End If
    
    ' UKUPNO
    result(rowCount, 1) = "UKUPNO"
    result(rowCount, 2) = totKol
    result(rowCount, 3) = ""       ' Keine Durchschnittscena
    result(rowCount, 4) = totVr
    result(rowCount, 5) = novacTotal
    result(rowCount, 6) = totVr - novacTotal
    result(rowCount, 7) = totAmb
    
    ReportSaldoKupci = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function


Public Function ReportIsplata(ByVal entitetTip As String, _
                              ByVal entitetID As String, _
                              ByVal datumOd As Date, _
                              ByVal datumDo As Date) As Variant

    Const SRC As String = "modIzvestaj.ReportIsplata"
    On Error GoTo EH
    ' Returns: 2D Array pro Kooperant
    ' Spalten: Kooperant | KesOtkupac | VirmanFirma | VirmanAvans | Ukupno
    ' + Summary: OM Avans primljeno | OM Avans podeljeno | Kod Otkupca
    
    Dim data As Variant
    data = GetTableData(TBL_NOVAC)
    If IsEmpty(data) Then
        ReportIsplata = Empty
        Exit Function
    End If
    data = ExcludeStornirano(data, TBL_NOVAC)
    If IsEmpty(data) Or Not IsArray(data) Then
        ReportIsplata = Empty
        Exit Function
    End If
    
    Dim colDatum As Long, colOMID As Long, colTip As Long
    Dim colIsplata As Long, colKoopID As Long, colPartnerID As Long
    
    colDatum = RequireColumnIndex(TBL_NOVAC, COL_NOV_DATUM, "modIzvestaj.ReportIsplata")
    colOMID = RequireColumnIndex(TBL_NOVAC, COL_NOV_OM_ID, "modIzvestaj.ReportIsplata")
    colTip = RequireColumnIndex(TBL_NOVAC, COL_NOV_TIP, "modIzvestaj.ReportIsplata")
    colIsplata = RequireColumnIndex(TBL_NOVAC, COL_NOV_ISPLATA, "modIzvestaj.ReportIsplata")
    colKoopID = RequireColumnIndex(TBL_NOVAC, COL_NOV_KOOP_ID, "modIzvestaj.ReportIsplata")
    colPartnerID = RequireColumnIndex(TBL_NOVAC, COL_NOV_PARTNER_ID, "modIzvestaj.ReportIsplata")
    
    ' Dicts: KooperantID ? Array(KesOtkupac, VirmanFirma, VirmanAvans)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim totalOMAvans As Double
    Dim totalKesOtkupac As Double
    
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Not IsDate(data(i, colDatum)) Then GoTo NextRow
        If CDate(data(i, colDatum)) < datumOd Or CDate(data(i, colDatum)) > datumDo Then GoTo NextRow
        
        Dim match As Boolean: match = False
        Select Case entitetTip
            Case "OM":    match = (CStr(data(i, colOMID)) = entitetID)
            Case "Kupac": match = (CStr(data(i, colPartnerID)) = entitetID)
        End Select
        If Not match Then GoTo NextRow
        
        Dim tipNovca As String
        tipNovca = CStr(data(i, colTip))
        Dim iznos As Double: iznos = 0
        If IsNumeric(data(i, colIsplata)) Then iznos = CDbl(data(i, colIsplata))
        If iznos <= 0 Then GoTo NextRow
        
        Dim koopID As String
        koopID = CStr(data(i, colKoopID))
        
        ' OM Avans (Firma ? Otkupac) -- kein Kooperant; oba kanala (kes + virman).
        If IsFirmaOtkupacAvansTip(tipNovca) Then
            totalOMAvans = totalOMAvans + iznos
            GoTo NextRow
        End If
        
        ' Kooperant-bezogene Isplate
        If koopID = "" Then GoTo NextRow
        
        If Not dict.Exists(koopID) Then dict.Add koopID, Array(0#, 0#, 0#)
        Dim vals As Variant
        vals = dict(koopID)
        
        Select Case tipNovca
            Case NOV_KES_OTKUPAC_KOOP
                vals(0) = vals(0) + iznos
                totalKesOtkupac = totalKesOtkupac + iznos
            Case NOV_VIRMAN_FIRMA_KOOP
                vals(1) = vals(1) + iznos
            Case NOV_VIRMAN_AVANS_KOOP
                vals(2) = vals(2) + iznos
        End Select
        
        dict(koopID) = vals
NextRow:
    Next i
    
    If dict.count = 0 And totalOMAvans = 0 Then
        ReportIsplata = Empty
        Exit Function
    End If
    
    ' Ergebnis: Kooperanten + UKUPNO + 3 Summary-Zeilen
    Dim rowCount As Long
    rowCount = dict.count + 4  ' UKUPNO + 3 Kontrolle
    
    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To 5)
    
    Dim keys As Variant
    If dict.count > 0 Then keys = dict.keys
    Dim koopNameDict As Object
    Set koopNameDict = BuildLookupDict(TBL_KOOPERANTI, "KooperantID", "Ime", "Prezime")
    Dim totKes As Double, totVirman As Double, totAvans As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        
        Dim koopNaziv As String
        If koopNameDict.Exists(CStr(keys(i))) Then koopNaziv = koopNameDict(CStr(keys(i))) Else koopNaziv = ""

        result(i + 1, 1) = koopNaziv
        result(i + 1, 2) = vals(0)                          ' KesOtkupac
        result(i + 1, 3) = vals(1)                          ' VirmanFirma
        result(i + 1, 4) = vals(2)                          ' VirmanAvans
        result(i + 1, 5) = vals(0) + vals(1) + vals(2)      ' Ukupno
        
        totKes = totKes + vals(0)
        totVirman = totVirman + vals(1)
        totAvans = totAvans + vals(2)
    Next i
    
    ' UKUPNO
    Dim ukRow As Long
    ukRow = dict.count + 1
    result(ukRow, 1) = "UKUPNO"
    result(ukRow, 2) = totKes
    result(ukRow, 3) = totVirman
    result(ukRow, 4) = totAvans
    result(ukRow, 5) = totKes + totVirman + totAvans
    
    ' Kontrolle
    result(ukRow + 1, 1) = "OM Avans (primljeno)"
    result(ukRow + 1, 5) = totalOMAvans
    
    result(ukRow + 2, 1) = "OM Avans (podeljeno)"
    result(ukRow + 2, 5) = totalKesOtkupac
    
    result(ukRow + 3, 1) = "Kod Otkupca"
    result(ukRow + 3, 5) = totalOMAvans - totalKesOtkupac
    
    ReportIsplata = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Public Function ReportOtkupRoba(ByVal entitetTip As String, _
                                ByVal entitetID As String, _
                                ByVal datumOd As Date, _
                                ByVal datumDo As Date) As Variant
                                
    Const SRC As String = "modIzvestaj.ReportOtkupRoba"
    On Error GoTo EH
    ' Returns: 2D Array (Col1, Col2, Kolicina, Vrednost)
    '   OM:    Datum, BrojOtp+Vrsta, Kg, RSD
    '   Kupac: Nr, Vrsta, Kg, RSD
    '   Vozac: Nr, Vrsta, Kg, RSD
    ' Letzte Zeile = UKUPNO
    
    ' Eksplicitan dispatch (RF-06): nepodrzan tip daje Empty, nikad "neki drugi"
    ' izvestaj pod pogresnim naslovom.
    Select Case entitetTip
        Case "OM":    ReportOtkupRoba = ReportOtkupRobaOM(entitetID, datumOd, datumDo)
        Case "Kupac": ReportOtkupRoba = ReportOtkupRobaKupac(entitetID, datumOd, datumDo)
        Case "Vozac": ReportOtkupRoba = ReportOtkupRobaVozac(entitetID, datumOd, datumDo)
        Case Else:    ReportOtkupRoba = Empty
    End Select
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportOtkupRobaOM(ByVal stanicaID As String, _
                                   ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportOtkupRobaOM"
    On Error GoTo EH
    
    Dim otpData As Variant
    otpData = GetOtpremniceByStation(stanicaID, datumOd, datumDo)
    If IsEmpty(otpData) Then
        ReportOtkupRobaOM = Empty
        Exit Function
    End If
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    If IsEmpty(otpData) Or Not IsArray(otpData) Then
        ReportOtkupRobaOM = Empty
        Exit Function
    End If
    
    Dim colVrsta As Long, colBrOtp As Long
    Dim colDatum As Long, colVozac As Long
    Dim colOtpID As Long, colBrZbirne As Long, colIzd As Long
    colIzd = RequireColumnIndex(TBL_OTPREMNICA, COL_TRACE_IZDATO_STATUS, SRC)
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, SRC)
    colBrOtp = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, SRC)
    colDatum = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    colVozac = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    colOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    colBrZbirne = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)

    ' Kolone 7 (kg blokova) i 8 (razlika) ostaju PRAZNE od S1b-3: racunale su se
    ' preko veze Otkup.OtpremnicaID, koju S3 zamenjuje sa tblOtpremnicaIzvori.
    ' Oblik rezultata se ne menja (ekran i stampa citaju kolone po polozaju);
    ' S3 ih puni iz novog modela.

    ' --- Manjak pro Zbirna ---
    Dim manjakDict As Object
    Set manjakDict = BuildManjakDict()
    
    ' --- JEDAN RED = JEDNA KLASA (S3b) --------------------------------------
    ' Klasa i kilaza vise nisu na zaglavlju. Red izvestaja OSTAJE po klasi, i to
    ' nije kozmetika: manjak se razresava kroz kljuc stavke zbirne, koji nosi
    ' klasu (v. komentar nize) -- spajanje klasa u jedan red bi prijem obe klase
    ' sabralo i pripisalo jednoj. Zato se zaglavlja RAZVIJAJU u parove
    ' (zaglavlje, stavka) pre petlje; sve ostalo racuna kao i pre.
    Dim stavkeDok As Object
    Set stavkeDok = modDokumenta.StavkeOtpremnicePoDokumentu()

    Dim mapRed() As Long, mapKlasa() As String, mapKg() As Double
    Dim rowCount As Long, h As Long, s As Long
    Dim stavke As Collection, stavka As Variant

    ' Samo IZDATE otpremnice su otpremljena roba (review #362, P1): nacrt nema
    ' izvore ni knjizene gajbe, pa u "otkupljeno i otpremljeno" ne ulazi.
    For h = 1 To UBound(otpData, 1)
        If modDokumenta.IzdatoStatusJeIzdato(otpData(h, colIzd)) Then
            Set stavke = modDokumenta.StavkeZaOtpremnicu(stavkeDok, _
                             Trim$(NzToText(otpData(h, colOtpID))), SRC)
            rowCount = rowCount + stavke.count
        End If
    Next h

    If rowCount = 0 Then
        ReportOtkupRobaOM = Empty
        Exit Function
    End If

    ReDim mapRed(1 To rowCount)
    ReDim mapKlasa(1 To rowCount)
    ReDim mapKg(1 To rowCount)

    Dim nPar As Long
    For h = 1 To UBound(otpData, 1)
        If modDokumenta.IzdatoStatusJeIzdato(otpData(h, colIzd)) Then
            Set stavke = modDokumenta.StavkeZaOtpremnicu(stavkeDok, _
                             Trim$(NzToText(otpData(h, colOtpID))), SRC)
            For s = 1 To stavke.count
                stavka = stavke(s)
                nPar = nPar + 1
                mapRed(nPar) = h
                mapKlasa(nPar) = KlasaOrDefault(stavka(3))
                mapKg(nPar) = CDbl(stavka(4))
            Next s
        End If
    Next h

    Dim result() As Variant
    ReDim result(1 To rowCount + 1, 1 To 12)   ' +Prijemnica kg (9), +skriveni OTP|<id> (12)
    
    Dim totOtp As Double, totManjak As Double
    Dim totPrijemnica As Double
    Dim totOtpSaPrijemom As Double   ' osnovica za UKUPNO manjak % (samo redovi sa prijemom)
    Dim malinaMode As Boolean: malinaMode = IsMalinaMode()
    ' Mapa vozaca (ID -> "Ime Prezime") -- jednom, umesto LookupValue u petlji nize.
    Dim vozacDict As Object
    Set vozacDict = BuildLookupDict(TBL_VOZACI, "VozacID", "Ime", "Prezime")
    Dim i As Long

    For i = 1 To rowCount
        ' hr = red ZAGLAVLJA kome ovaj red izvestaja pripada; i = par
        ' (zaglavlje, stavka). Zaglavlje se cita preko hr, stavka preko i.
        Dim hr As Long
        hr = mapRed(i)

        Dim kgOtp As Double
        kgOtp = mapKg(i)

        Dim thisOtpID As String
        thisOtpID = CStr(otpData(hr, colOtpID))

        ' Manjak proportional berechnen
        Dim thisBrZbirne As String
        thisBrZbirne = Trim$(CStr(otpData(hr, colBrZbirne)))

        ' Vozac i klasa otpremnice -- treba za razresenje STAVKE zbirne (dole) i
        ' za prikaz, pa se citaju pre oba.
        Dim vozID As String
        vozID = Trim$(CStr(otpData(hr, colVozac)))
        Dim klasaOtp As String
        klasaOtp = mapKlasa(i)

        ' Prijem po otpremnici -- vezan za STAVKU (vlasnik + Klasa), ne za broj.
        ' Otpremnica nosi BrojZbirne, VozacID i Klasu, ali ne i KupacID, pa se
        ' vlasnik razresava ovako:
        '   #V = 1  -> broj ima jednog vlasnika: agregat po (broj, klasa) je
        '              dokazano siguran (i hvata starije prijemnice bez vlasnika),
        '   #V > 1  -> broj dele dve zbirne: pokusaj razresenja po vozacu (#O);
        '              ako ne uspe ili postoji nepripisiva prijemnica -> fail-closed
        '              oznaka IZV_VLASNIK_NEJASAN, bez izmisljene brojke.
        ' Klasa MORA biti u kljucu: Klasa I i II istog dokumenta dele broj, vozaca
        ' i kupca, ali imaju zasebnu otpremnicu/zbirnu/prijemnicu. Bez nje bi se
        ' prijem obe klase sabrao i taj zbir dodelio SVAKOJ klasi (u malina modu
        ' bukvalno duplo -- UKUPNO prijem 2x stvarni).
        ' Bez prijema red NEMA brojku manjka nego oznaku -- pre RF-06 se isti
        ' slucaj prikazivao kao 0 kg / 0,00% (FM-0028 #5).
        Dim prijemnicaKg As Double: prijemnicaKg = 0
        Dim imaPrijem As Boolean: imaPrijem = False
        Dim oznakaBez As String: oznakaBez = IZV_NEMA_PRIJEMA
        Dim zbirnaTotal As Double: zbirnaTotal = 0
        Dim prijTotal As Double: prijTotal = 0

        Dim nVlasnika As Long: nVlasnika = 0
        If manjakDict.Exists("#V|" & thisBrZbirne) Then nVlasnika = CLng(manjakDict("#V|" & thisBrZbirne))

        Dim cntNejasan As Long: cntNejasan = 0
        If manjakDict.Exists("#N|" & thisBrZbirne) Then cntNejasan = CLng(manjakDict("#N|" & thisBrZbirne))

        Dim stavkaKey As String: stavkaKey = ""
        Dim razresen As Boolean: razresen = False
        Dim cntPrijem As Long: cntPrijem = 0

        If nVlasnika = 1 Then
            razresen = True
            stavkaKey = ZbirnaStavkaKljuc(CStr(manjakDict("#1|" & thisBrZbirne)), klasaOtp)
            ' Prijem po (broj, klasa): dokazano jedan vlasnik.
            Dim bkKey As String
            bkKey = thisBrZbirne & "|" & klasaOtp
            If manjakDict.Exists("#C|" & bkKey) Then cntPrijem = CLng(manjakDict("#C|" & bkKey))
            If manjakDict.Exists("#K|" & bkKey) Then prijTotal = CDbl(manjakDict("#K|" & bkKey))
        ElseIf nVlasnika > 1 Then
            Dim vozKey As String
            vozKey = "#O|" & thisBrZbirne & "|" & Trim$(vozID)
            If manjakDict.Exists(vozKey) Then
                razresen = True
                stavkaKey = ZbirnaStavkaKljuc(CStr(manjakDict(vozKey)), klasaOtp)
            End If
            If razresen Then
                If manjakDict.Exists(stavkaKey) Then
                    Dim ownVals As Variant
                    ownVals = manjakDict(stavkaKey)
                    prijTotal = CDbl(ownVals(1))
                    cntPrijem = CLng(ownVals(2))
                End If
            End If
        End If

        If Len(stavkaKey) > 0 Then
            If manjakDict.Exists(stavkaKey) Then
                Dim zbVals As Variant
                zbVals = manjakDict(stavkaKey)
                zbirnaTotal = CDbl(zbVals(0))   ' osnovica srazmere = kg TE klase
            End If
        End If

        Dim pz As Variant
        pz = PrijemZaZbirnu(nVlasnika, razresen, cntNejasan, cntPrijem, prijTotal)
        imaPrijem = CBool(pz(0))
        oznakaBez = CStr(pz(2))

        If imaPrijem Then
            prijTotal = CDbl(pz(1))
            If malinaMode Then
                ' Malina: 1 otpremnica = 1 zbirna = 1 prijemnica PO KLASI -> direktno.
                prijemnicaKg = prijTotal
            ElseIf zbirnaTotal > 0 Then
                ' Srazmerno udelu otpremnice u zbirnoj -- UNUTAR iste klase.
                prijemnicaKg = prijTotal * (kgOtp / zbirnaTotal)
            Else
                ' Nema upotrebljive osnovice za srazmeru -> ne izmisljaj manjak.
                imaPrijem = False
                oznakaBez = IZV_NEMA_PRIJEMA
            End If
        End If

        Dim mStavka As Variant
        mStavka = ManjakStavka(kgOtp, prijemnicaKg, imaPrijem, oznakaBez)

        ' Vozac Name
        Dim vozNaziv As String
        If vozID <> "" Then
            If vozacDict.Exists(vozID) Then vozNaziv = vozacDict(vozID) Else vozNaziv = ""
        Else
            vozNaziv = ""
        End If
        
        result(i, 1) = CDate(otpData(hr, colDatum))
        result(i, 2) = CStr(otpData(hr, colBrOtp))
        result(i, 3) = CStr(otpData(hr, colVrsta))
        result(i, 4) = klasaOtp
        result(i, 5) = vozNaziv
        result(i, 6) = kgOtp
        result(i, 7) = ""               ' kg blokova -- S3
        result(i, 8) = ""               ' razlika    -- S3
        result(i, 9) = mStavka(0)      ' Prijemnica kg (prazno kad nema prijema)
        result(i, 10) = mStavka(1)     ' Manjak kg     (prazno kad nema prijema)
        If imaPrijem Then
            result(i, 11) = mStavka(2) ' Manjak %
        Else
            result(i, 11) = mStavka(3) ' oznaka "nema prijema"
        End If
        result(i, 12) = "OTP|" & thisOtpID

        totOtp = totOtp + kgOtp

        ' Manjak-total ide SAMO preko redova sa prijemom; inace bi otpremnice bez
        ' prijemnice pomerale i zbir i procenat manjka.
        If imaPrijem Then
            totManjak = totManjak + CDbl(mStavka(1))
            totPrijemnica = totPrijemnica + prijemnicaKg
            totOtpSaPrijemom = totOtpSaPrijemom + kgOtp
        End If
    Next i

    ' UKUPNO
    result(rowCount + 1, 2) = "UKUPNO"
    result(rowCount + 1, 6) = totOtp
    result(rowCount + 1, 7) = ""
    result(rowCount + 1, 8) = ""
    result(rowCount + 1, 9) = totPrijemnica
    result(rowCount + 1, 10) = totManjak
    If totOtpSaPrijemom > 0 Then result(rowCount + 1, 11) = totManjak / totOtpSaPrijemom * 100

    ReportOtkupRobaOM = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportOtkupRobaKupac(ByVal kupacID As String, _
                                      ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportOtkupRobaKupac"
    On Error GoTo EH
    ' Aggregiert pro VrstaVoca
    Dim prijData As Variant
    prijData = GetPrijemniceByKupac(kupacID, datumOd, datumDo)
    If IsEmpty(prijData) Or Not IsArray(prijData) Then
        ReportOtkupRobaKupac = Empty
        Exit Function
    End If
    prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If IsEmpty(prijData) Or Not IsArray(prijData) Then
        ReportOtkupRobaKupac = Empty
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colVrsta As Long, colKol As Long, colCena As Long
    colVrsta = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VRSTA, "modIzvestaj.ReportOtkupRobaKupac")
    colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "modIzvestaj.ReportOtkupRobaKupac")
    colCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, "modIzvestaj.ReportOtkupRobaKupac")
    
    Dim i As Long
    For i = 1 To UBound(prijData, 1)
        Dim key As String
        key = CStr(prijData(i, colVrsta))
        If key = "" Then key = "(Nepoznato)"
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(prijData(i, colKol)) Then vals(0) = vals(0) + CDbl(prijData(i, colKol))
        If IsNumeric(prijData(i, colKol)) And IsNumeric(prijData(i, colCena)) Then
            vals(1) = vals(1) + CDbl(prijData(i, colKol)) * CDbl(prijData(i, colCena))
        End If
        dict(key) = vals
    Next i
    
    ReportOtkupRobaKupac = DictToResultArray(dict)
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportOtkupRobaVozac(ByVal vozacID As String, _
                                      ByVal datumOd As Date, _
                                      ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportOtkupRobaVozac"
    On Error GoTo EH
    
    Dim otpData As Variant
    otpData = GetVozacDokumenta(vozacID, datumOd, datumDo)
    
    If IsEmpty(otpData) Or Not IsArray(otpData) Then
        ReportOtkupRobaVozac = Empty
        Exit Function
    End If
    
    otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    
    If IsEmpty(otpData) Or Not IsArray(otpData) Then
        ReportOtkupRobaVozac = Empty
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Samo IZDATE otpremnice su otpremljena roba (review #362, P1). Vrsta je na
    ' zaglavlju, kilaza na STAVKAMA, a vrednost je vrednost IZVORNIH otkupa --
    ' ne Kolicina x PredlogCena.
    Dim colVrsta As Long, colOtpID As Long, colIzd As Long
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, SRC)
    colOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
    colIzd = RequireColumnIndex(TBL_OTPREMNICA, COL_TRACE_IZDATO_STATUS, SRC)

    Dim zbir As Object, vredIzv As Object
    Set zbir = modDokumenta.ZbirStavkiPoOtpremnici()
    Set vredIzv = modDokumenta.VrednostIzvoraPoOtpremnici()

    Dim i As Long
    For i = 1 To UBound(otpData, 1)
        If Not modDokumenta.IzdatoStatusJeIzdato(otpData(i, colIzd)) Then GoTo SledecaOtp
        Dim key As String
        key = CStr(otpData(i, colVrsta))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        Dim z As Variant
        z = modDokumenta.ZbirStavkiZaOtpremnicu(zbir, _
                Trim$(NzToText(otpData(i, colOtpID))), SRC)
        vals(0) = vals(0) + CDbl(z(0))
        vals(1) = vals(1) + modDokumenta.VrednostIzvoraZaOtpremnicu( _
                                vredIzv, Trim$(NzToText(otpData(i, colOtpID))), SRC)
        dict(key) = vals
SledecaOtp:
    Next i

    ReportOtkupRobaVozac = DictToResultArray(dict)
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' AMBALAZA REPORT
' ============================================================

Public Function ReportAmbalaza(ByVal entitetTip As String, _
                               ByVal entitetID As String, _
                               ByVal datumOd As Date, _
                               ByVal datumDo As Date, _
                               ByVal zbirni As Boolean) As Variant
    
    Const SRC As String = "modIzvestaj.ReportAmbalaza"
    On Error GoTo EH
    ' Zbirni Returns: 2D Array (Tip, "", "", "", Ulaz, Izlaz)
    ' Einzeln Returns: 2D Array (Datum, Mesto, Tip, DokID, Ulaz, Izlaz)
    ' Letzte Zeile = UKUPNO
    
    Dim data As Variant
    data = GetTableData(TBL_AMBALAZA)
    If IsEmpty(data) Then
        ReportAmbalaza = Empty
        Exit Function
    End If
    ' --- Filter aufbauen ---
    ' Storno se filtrira UNUTAR FilterArray (umesto zasebnog ExcludeStornirano koji
    ' je pravio JOS jednu kopiju cele tblAmbalaza) -> jedan prolaz umesto dva.
    Dim filters As New Collection
    Dim fp As clsFilterParam

    Dim colStornoAmb As Long
    colStornoAmb = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    If colStornoAmb > 0 Then
        Set fp = New clsFilterParam
        fp.Init colStornoAmb, "<>", "Da"
        filters.Add fp
    End If

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, "modIzvestaj.ReportAmbalaza"), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    ' Eksplicitan dispatch (RF-06): nepoznat tip je pre padao kroz SVE grane bez
    ' entitet filtera -> globalni ambalazni izvestaj pod naslovom entiteta
    ' (FM-0028 #12).
    Select Case entitetTip
    Case "OM"
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, "modIzvestaj.ReportAmbalaza"), "=", entitetID
        filters.Add fp

        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, "modIzvestaj.ReportAmbalaza"), "=", "Stanica"
        filters.Add fp

    Case "Kupac"
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, "modIzvestaj.ReportAmbalaza"), "=", entitetID
        filters.Add fp

        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, "modIzvestaj.ReportAmbalaza"), "=", "Kupac"
        filters.Add fp

    Case "Vozac"
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC, "modIzvestaj.ReportAmbalaza"), "=", entitetID
        filters.Add fp

        ' Otkup (Kooperant-nabavka) NIJE vozaceva transportna noga. Iste gajbice
        ' se vec broje na otpremnici, pa bi otkup duplo teretio vozacev saldo
        ' (narocito uz auto-hladnjacu, koja mirror-vozaca vezuje za svaki otkup).
        ' Vozacev saldo = otpremnica (utovar) - prijemnica (predaja).
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, "modIzvestaj.ReportAmbalaza"), "<>", DOK_TIP_OTKUP
        filters.Add fp

    Case Else
        ReportAmbalaza = Empty
        Exit Function
    End Select

    Dim filtered As Variant
    filtered = FilterArray(data, filters)
    If IsEmpty(filtered) Or Not IsArray(filtered) Then
        ReportAmbalaza = Empty
        Exit Function
    End If
    
    Dim colTip As Long, colKol As Long, colSmer As Long
    Dim colDokID As Long, colDokTip As Long, colDatum As Long
    Dim colEntitet As Long, colEntTip As Long, colRID As Long
    colTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, "modIzvestaj.ReportAmbalaza")
    colKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, "modIzvestaj.ReportAmbalaza")
    colSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, "modIzvestaj.ReportAmbalaza")
    colDokID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, "modIzvestaj.ReportAmbalaza")
    colDokTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, "modIzvestaj.ReportAmbalaza")
    colDatum = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, "modIzvestaj.ReportAmbalaza")
    colEntitet = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, "modIzvestaj.ReportAmbalaza")
    colEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, "modIzvestaj.ReportAmbalaza")
    colRID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_REVERS_ID, "modIzvestaj.ReportAmbalaza")
    
    ' Vozac = inverzni protivpartner entiteta (Stanica / Kupac); kompletna ruta
    ' otpremnica -> prijemnica daje saldo 0. Otkup nema vozaca -> izuzet (filter).
    ' Entitetski izvestaji (OM / Kupac) koriste sirovi Smer (isVozac = False).
    Dim isVozac As Boolean
    isVozac = (entitetTip = "Vozac")

    If zbirni Then
        ReportAmbalaza = ReportAmbalazeZbirni(filtered, colTip, colKol, colSmer, colEntTip, isVozac)
    Else
        ReportAmbalaza = ReportAmbalazePojedinacni(filtered, colDatum, colEntitet, colEntTip, _
                                                    colTip, colDokID, colDokTip, colKol, colSmer, isVozac, colRID)
    End If
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportAmbalazeZbirni(ByVal filtered As Variant, _
                                      ByVal colTip As Long, ByVal colKol As Long, _
                                      ByVal colSmer As Long, _
                                      ByVal colEntTip As Long, _
                                      ByVal isVozac As Boolean) As Variant
    
    Const SRC As String = "modIzvestaj.ReportAmbalazeZbirni"
    On Error GoTo EH
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim key As String
        key = CStr(filtered(i, colTip))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        Dim kol As Long: kol = 0
        If IsNumeric(filtered(i, colKol)) Then kol = CLng(filtered(i, colKol))
        Dim effSmer As String
        effSmer = CStr(filtered(i, colSmer))
        If isVozac Then effSmer = VozacAmbEffectiveSmer(effSmer, CStr(filtered(i, colEntTip)))
        If effSmer = "Ulaz" Then
            vals(0) = vals(0) + kol
        Else
            vals(1) = vals(1) + kol
        End If
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportAmbalazeZbirni = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 6)
    
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = ""
        result(i + 1, 3) = ""
        result(i + 1, 4) = ""
        result(i + 1, 5) = vals(0)
        result(i + 1, 6) = vals(1)
    Next i
    
    ReportAmbalazeZbirni = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportAmbalazePojedinacni(ByVal filtered As Variant, _
                                            ByVal colDatum As Long, ByVal colEntitet As Long, _
                                            ByVal colEntTip As Long, ByVal colTip As Long, _
                                            ByVal colDokID As Long, ByVal colDokTip As Long, _
                                            ByVal colKol As Long, _
                                            ByVal colSmer As Long, ByVal isVozac As Boolean, _
                                            ByVal colRID As Long) As Variant
    
    Const SRC As String = "modIzvestaj.ReportAmbalazePojedinacni"
    On Error GoTo EH
    
    Dim rowCount As Long
    rowCount = UBound(filtered, 1)

    ' Grupisanje po JEDNOM dokumentu (DokumentTip + DokumentID + TipAmbalaze):
    ' ako isti dokument ima i Ulaz i Izlaz red, prikazi oba u istom redu. Ako ima
    ' samo jedan smer -> red ostaje kao i do sada. Redovi RAZLICITOG DokumentTip-a
    ' su razliciti dokumenti (uz-otkup: `Otkup` + `OM-Izlaz-Koop` dele otkupID) i
    ' NE smeju u isti red -- vidi gkey nize.
    ' Scripting.Dictionary cuva redosled umetanja (kao redosled filtriranih redova).
    Dim grp As Object
    Set grp = CreateObject("Scripting.Dictionary")

    ' Memo za ResolveEntitetName: (tip|id) -> naziv za trajanje ovog izvestaja, da se
    ' LookupValue ne ponavlja po svakom redu (O(jedinstvenih) umesto O(redova)).
    Dim nameMemo As Object
    Set nameMemo = CreateObject("Scripting.Dictionary")

    Dim totalUlaz As Long, totalIzlaz As Long
    Dim i As Long
    For i = 1 To rowCount
        Dim kol As Long: kol = 0
        If IsNumeric(filtered(i, colKol)) Then kol = CLng(filtered(i, colKol))

        Dim entID As String: entID = CStr(filtered(i, colEntitet))
        Dim entTipVal As String: entTipVal = CStr(filtered(i, colEntTip))
        Dim dokIDv As String: dokIDv = CStr(filtered(i, colDokID))
        Dim tipv As String: tipv = CStr(filtered(i, colTip))
        Dim dokTipv As String: dokTipv = CStr(filtered(i, colDokTip))

        Dim effSmer As String
        effSmer = CStr(filtered(i, colSmer))
        If isVozac Then effSmer = VozacAmbEffectiveSmer(effSmer, entTipVal)

        ' PUN identitet dokumenta, isti koji `ReversRedPripada` koristi za match:
        ' DokumentTip + DokumentID + TipAmbalaze. `DokumentTip` je nuzan jer
        ' `modOtkup.SaveOtkup` na NORMALNOJ putanji upisuje isti `otkupID` i isti
        ' tip ambalaze pod DVA tipa dokumenta -- primljene pune gajbe kao
        ' `DOK_TIP_OTKUP`, izdate prazne kao `DOK_TIP_OM_IZLAZ_KOOP`. Bez njega su
        ' se spajali u JEDAN red koji nosi tip PRVOG zapisa ("Otkup"), pa je
        ' skriveni ref-kljuc bio `AMB|Otkup|<id>` i "Stampaj dokument" je uvek
        ' rutirao na `ReprintOtkupniListByOtkupID` -- revers `OM-Izlaz-Koop` nije
        ' imao svoj red i bio je NEDOSTUPAN za stampu iz pregleda.
        Dim gkey As String
        gkey = Trim$(dokTipv) & "|" & Trim$(dokIDv) & "|" & AmbTipKljuc(tipv)
        ' REVERS: identitet je ReversID (REV-IDENT-01) -- isti broj legalno nose
        ' reversi druge stanice ili drugog dana (A2 red REV), pa bez njega pregled
        ' po vozacu spaja dva reversa u jedan red sa datumom prvog i zbirom
        ' kolicina. ReversID ide i u red (ref-kljuc), da stampa bira noge po
        ' njemu, a ne po broju. Ambalaza uz otkup ReversID nema (DokumentID =
        ' OtkupID je vec jedinstven): za nju ostaje (stanica, dan) -- u pregledu
        ' po OM i po vozacu red daje samo nogu Stanica, pa EntitetID JESTE stanica.
        ' Samostalan revers bez ReversID-a (nalaz B10) grupise se isto, a stampa
        ' ga odbija.
        Dim ridv As String: ridv = Trim$(NzToText(filtered(i, colRID)))
        If Len(ridv) > 0 Then
            gkey = gkey & "|" & ridv
        ElseIf modStorno.ReversTipJe(dokTipv) Then
            Dim danKljuc As String: danKljuc = ""
            If IsDate(filtered(i, colDatum)) Then danKljuc = CStr(Int(CDbl(CDate(filtered(i, colDatum)))))
            gkey = gkey & "|" & UCase$(Trim$(entID)) & "|" & danKljuc
        End If
        Dim rec As Variant
        If grp.Exists(gkey) Then
            rec = grp(gkey)
        Else
            ' Datum, Mesto, Tip, Dokument, Ulaz, Izlaz, DokTip, ReversID
            Dim entMemoKey As String: entMemoKey = entTipVal & "|" & entID
            If Not nameMemo.Exists(entMemoKey) Then nameMemo.Add entMemoKey, ResolveEntitetName(entID, entTipVal)
            rec = Array(filtered(i, colDatum), CStr(nameMemo(entMemoKey)), _
                        tipv, dokIDv, 0&, 0&, dokTipv, ridv)
        End If
        If effSmer = "Ulaz" Then
            rec(4) = CLng(rec(4)) + kol
            totalUlaz = totalUlaz + kol
        Else
            rec(5) = CLng(rec(5)) + kol
            totalIzlaz = totalIzlaz + kol
        End If
        grp(gkey) = rec
    Next i

    Dim nGrp As Long: nGrp = grp.Count
    Dim result() As Variant
    ReDim result(1 To nGrp + 1, 1 To 7)  ' +1 UKUPNO, kol.7 = skriveni ref-kljuc

    ' Poslovni brojevi dokumenata JEDNIM prolazom po tabeli (mape), umesto
    ' LookupValue po redu: na svesci sa 1.596 amb redova je razresenje broja
    ' radilo 1.596 punih skenova tabela i tab je delovao zamrznuto (smoke
    ' 28.08, krug 3) -- isti potez kao BuildOtkupBrojDokDict u karticama.
    Dim mapaOtp As Object, mapaPrj As Object, mapaOtk As Object
    Set mapaOtp = BuildLookupDict(TBL_OTPREMNICA, COL_OTP_ID, COL_OTP_BROJ)
    Set mapaPrj = BuildLookupDict(TBL_PRIJEMNICA, COL_PRJ_ID, COL_PRJ_BROJ)
    Set mapaOtk = BuildLookupDict(TBL_OTKUP, COL_OTK_ID, COL_OTK_BR_DOK)

    Dim keys As Variant: keys = grp.keys
    Dim r As Long
    For r = 0 To nGrp - 1
        Dim rr As Variant: rr = grp(keys(r))
        If IsDate(rr(0)) Then
            result(r + 1, 1) = CDate(rr(0))
        Else
            result(r + 1, 1) = rr(0)
        End If
        result(r + 1, 2) = rr(1)
        result(r + 1, 3) = rr(2)
        result(r + 1, 4) = ResolveDokBrojMape(CStr(rr(6)), CStr(rr(3)), _
                                              mapaOtp, mapaPrj, mapaOtk)
        result(r + 1, 5) = IIf(CLng(rr(4)) <> 0, CLng(rr(4)), "")
        result(r + 1, 6) = IIf(CLng(rr(5)) <> 0, CLng(rr(5)), "")
        ' "AMB|<DokTip>|<DokID>[|<ReversID>]" -- ReversID samo kad ga red nosi.
        result(r + 1, 7) = "AMB|" & CStr(rr(6)) & "|" & CStr(rr(3))
        If Len(CStr(rr(7))) > 0 Then result(r + 1, 7) = CStr(result(r + 1, 7)) & "|" & CStr(rr(7))
    Next r

    ' UKUPNO
    result(nGrp + 1, 1) = "UKUPNO"
    result(nGrp + 1, 2) = ""
    result(nGrp + 1, 3) = ""
    result(nGrp + 1, 4) = "Saldo: " & Format$(totalUlaz - totalIzlaz, "#,##0")
    result(nGrp + 1, 5) = totalUlaz
    result(nGrp + 1, 6) = totalIzlaz
    result(nGrp + 1, 7) = ""

    ReportAmbalazePojedinacni = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Poslovni broj dokumenta iz internog DokumentID-a (za prikaz u Ambalaza
' pregledu), nad UNAPRED izgradjenim mapama ID -> broj. Isto pravilo kao
' nekadasnji ResolveDokBroj (LookupValue po redu), samo O(1) po redu:
' BuildLookupDict je "prvi pojav pobedjuje", identicno LookupValue-u.
' Vraca DokumentID ako broj nije razresiv.
Private Function ResolveDokBrojMape(ByVal dokTip As String, ByVal dokID As String, _
                                    ByVal mapaOtp As Object, ByVal mapaPrj As Object, _
                                    ByVal mapaOtk As Object) As String
    On Error Resume Next
    Dim sOut As String: sOut = dokID
    Select Case dokTip
        Case DOK_TIP_OTPREMNICA
            If mapaOtp.Exists(dokID) Then sOut = CStr(mapaOtp(dokID))
        Case DOK_TIP_PRIJEMNICA
            If mapaPrj.Exists(dokID) Then sOut = CStr(mapaPrj(dokID))
        Case DOK_TIP_OTKUP, DOK_TIP_OM_IZLAZ_KOOP, DOK_TIP_OM_ULAZ_KOOP
            ' uz-otkup: DokumentID = otkupID -> BrojDokumenta; standalone
            ' revers: DokumentID = brojDok
            If mapaOtk.Exists(dokID) Then
                If Len(Trim$(CStr(mapaOtk(dokID)))) > 0 Then sOut = CStr(mapaOtk(dokID))
            End If
    End Select
    If Len(Trim$(sOut)) = 0 Then sOut = dokID
    ResolveDokBrojMape = sOut
End Function

' ============================================================
' PROSECNA CENA i MANJAK
' ============================================================

Public Function ReportProsecnaCena(ByVal entitetTip As String, _
                                   ByVal entitetID As String, _
                                   ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportProsecnaCena"
    On Error GoTo EH
    
    ' Returns: 2D Array (Vrsta, Kolicina, Vrednost, ProsecnaCena)
    '
    ' RF-06: dispatch je eksplicitan. Pre toga je SVAKI tip koji nije "Kupac"
    ' padao u otkup granu, pa je zbirni mod za Kooperante/Vozace (tab 6 je i njima
    ' vidljiv) prikazivao GLOBALNU prosecnu cenu otkupa pod pogresnim naslovom
    ' (FM-0028 #13). Sada takva kombinacija daje Empty = cista prazna lista;
    ' vidljiva poruka i suzavanje tab-matrice idu u RF-07 (frmIzvestaj).

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long

    Select Case entitetTip
    Case "Kupac"
        Dim prijData As Variant
        prijData = GetPrijemniceByKupac(entitetID, datumOd, datumDo)
        If IsEmpty(prijData) Then
            ReportProsecnaCena = Empty
            Exit Function
        End If
        prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
        If IsEmpty(prijData) Or Not IsArray(prijData) Then
            ReportProsecnaCena = Empty
            Exit Function
        End If
        Dim vrstaCache As Object
        Set vrstaCache = BuildZbirnaVrstaCache()
        
        Dim colBrZbr As Long, colPrijKol As Long, colPrijCena As Long
        colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "modIzvestaj.ReportProsecnaCena")
        colPrijKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "modIzvestaj.ReportProsecnaCena")
        colPrijCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, "modIzvestaj.ReportProsecnaCena")
        
        For i = 1 To UBound(prijData, 1)
            Dim vrsta As String
            vrsta = GetVrstaFromCache(vrstaCache, CStr(prijData(i, colBrZbr)))
            If vrsta = "" Then vrsta = "(Nepoznato)"
            
            If Not dict.Exists(vrsta) Then dict.Add vrsta, Array(0#, 0#)
            Dim vals As Variant
            vals = dict(vrsta)
            If IsNumeric(prijData(i, colPrijKol)) Then vals(0) = vals(0) + CDbl(prijData(i, colPrijKol))
            If IsNumeric(prijData(i, colPrijKol)) And IsNumeric(prijData(i, colPrijCena)) Then
                vals(1) = vals(1) + CDbl(prijData(i, colPrijKol)) * CDbl(prijData(i, colPrijCena))
            End If
            dict(vrsta) = vals
        Next i

    Case "OM", ""
        ' OM einzeln (entitetID) oder Zbirni/alle (entitetID = "")
        Dim otkData As Variant
        If entitetID <> "" Then
            otkData = GetOtkupByStation(entitetID, datumOd, datumDo)
        Else
            otkData = GetTableData(TBL_OTKUP)
            If Not IsEmpty(otkData) Then
                Dim filters As New Collection
                Dim fp As clsFilterParam
                Set fp = New clsFilterParam
                fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, "modIzvestaj.ReportProsecnaCena"), "BETWEEN", datumOd, datumDo
                filters.Add fp
                otkData = FilterArray(otkData, filters)
            End If
        End If
        If IsEmpty(otkData) Then
            ReportProsecnaCena = Empty
            Exit Function
        End If
        otkData = ExcludeStornirano(otkData, TBL_OTKUP)
        If IsEmpty(otkData) Or Not IsArray(otkData) Then
            ReportProsecnaCena = Empty
            Exit Function
        End If
        
        Dim colVrsta As Long, colOtkID As Long
        colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "modIzvestaj.ReportProsecnaCena")
        colOtkID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "modIzvestaj.ReportProsecnaCena")

        ' Kolicina i vrednost su na STAVKAMA (REFAKTOR S14.7, kvar 3).
        Dim stavkeZbir As Object
        Set stavkeZbir = modOtkup.ZbirStavkiPoOtkupu()

        For i = 1 To UBound(otkData, 1)
            Dim key As String
            key = CStr(otkData(i, colVrsta))
            If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
            vals = dict(key)
            Dim zPc As Variant
            zPc = modOtkup.ZbirStavkiZaOtkup(stavkeZbir, CStr(otkData(i, colOtkID)), _
                      "modIzvestaj.ReportProsecnaCena")
            vals(0) = vals(0) + CDbl(zPc(0))
            vals(1) = vals(1) + CDbl(zPc(1))
            dict(key) = vals
        Next i

    Case Else
        ' "Vozac" / "Kooperant" i sve nepoznato: prosecna cena za taj entitet
        ' nije definisana -> prazan izvestaj umesto tudjih brojki.
        ReportProsecnaCena = Empty
        Exit Function
    End Select

    If dict.count = 0 Then
        ReportProsecnaCena = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 4)
    
    Dim keys As Variant
    keys = dict.keys
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        result(i + 1, 1) = keys(i)
        result(i + 1, 2) = vals(0)
        result(i + 1, 3) = vals(1)
        If vals(0) > 0 Then
            result(i + 1, 4) = vals(1) / vals(0)
        Else
            result(i + 1, 4) = 0
        End If
    Next i
    
    ReportProsecnaCena = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Public Function ReportManjak(ByVal entitetTip As String, _
                             ByVal entitetID As String, _
                             ByVal datumOd As Date, _
                             ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportManjak"
    On Error GoTo EH
    
    ' Returns: 2D Array (BrojZbirne, ZbirnaKg, PrijKg, ManjakKg, ManjakPct, ProsekGajbe)
    ' Letzte Zeile = UKUPNO
    '
    ' RF-06:
    '  - dispatch je eksplicitan; nepodrzan tip (npr. zbirni Kooperanti, kojima
    '    je tab Manjak vidljiv) vise ne dobija GLOBALNI izvestaj (FM-0028 #4/#14);
    '  - zbirna bez prijemnice nosi oznaku IZV_NEMA_PRIJEMA umesto 100% manjka
    '    (isti podatak koji RobaOM prikazuje kao "nema prijema" -- FM-0028 #5);
    '  - UKUPNO se racuna SAMO nad zbirnama koje imaju prijem, pa nepreuzete
    '    posiljke ne naduvavaju zbir i procenat manjka.

    Select Case entitetTip
        Case "", "OM", "Kupac", "Vozac"
            ' podrzano (tblZbirna nema kolonu stanice -> "OM"/"" = bez entitet filtera)
        Case Else
            ReportManjak = Empty
            Exit Function
    End Select

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then
        ReportManjak = Empty
        Exit Function
    End If

    Dim filters As New Collection
    Dim fp As clsFilterParam

    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, "modIzvestaj.ReportManjak"), "BETWEEN", datumOd, datumDo
    filters.Add fp

    If entitetTip = "Kupac" And entitetID <> "" Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, "modIzvestaj.ReportManjak"), "=", entitetID
        filters.Add fp
    ElseIf entitetTip = "Vozac" And entitetID <> "" Then
        Set fp = New clsFilterParam
        fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, "modIzvestaj.ReportManjak"), "=", entitetID
        filters.Add fp
    End If

    Dim filtered As Variant
    filtered = FilterArray(zbrData, filters)
    If IsEmpty(filtered) Then
        ReportManjak = Empty
        Exit Function
    End If
    filtered = ExcludeStornirano(filtered, TBL_ZBIRNA)
    If IsEmpty(filtered) Or Not IsArray(filtered) Then
        ReportManjak = Empty
        Exit Function
    End If
    
    ' Prijem + indeks vlasnika: deljeni owner-scoped agregat (modHelpers).
    ' Raniji oblik je ovde rucno sabirao prijemnice po SAMOM BrojZbirne -- isti
    ' propust koji je RF-05/AUD-052 vec dokazao na storno putanji.
    Dim manjakDict As Object
    Set manjakDict = BuildManjakDict()

    ' Zbirna-Daten
    Dim colBroj As Long, colZbrKol As Long, colZbrAmb As Long
    Dim colZbrVoz As Long, colZbrKup As Long, colZbrKla As Long
    colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modIzvestaj.ReportManjak")
    colZbrKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, "modIzvestaj.ReportManjak")
    colZbrAmb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB, "modIzvestaj.ReportManjak")
    colZbrVoz = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, "modIzvestaj.ReportManjak")
    colZbrKup = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, "modIzvestaj.ReportManjak")
    colZbrKla = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KLASA, "modIzvestaj.ReportManjak")

    ' Zbirne aggregieren po VLASNIKU (broj|vozac|kupac), ne po broju: dve aktivne
    ' zbirne mogu deliti isti poslovni broj. Klasa I i II istog dokumenta ostaju
    ' u ISTOM redu (izvestaj namerno prikazuje ceo dokument jednim redom), ali se
    ' prijem svake klase cita ZASEBNO pa sabira -- prijemnice su po klasi.
    Dim zbrDict As Object
    Set zbrDict = CreateObject("Scripting.Dictionary")

    ' vlasnikKljuc -> Dictionary(klasa -> True): koje klase red obuhvata.
    Dim klasePoVlasniku As Object
    Set klasePoVlasniku = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim brZbr As String
        brZbr = Trim$(CStr(filtered(i, colBroj)))
        Dim vlKey As String
        vlKey = ZbirnaVlasnikKljuc(brZbr, _
                                   Trim$(NzToText(filtered(i, colZbrVoz))), _
                                   Trim$(NzToText(filtered(i, colZbrKup))))

        ' vals: 0 = zbirna kg, 1 = ambalaza, 2 = BrojZbirne (za prikaz)
        If Not zbrDict.Exists(vlKey) Then
            zbrDict.Add vlKey, Array(0#, 0#, brZbr)
        End If
        Dim zv As Variant
        zv = zbrDict(vlKey)
        If IsNumeric(filtered(i, colZbrKol)) Then zv(0) = zv(0) + CDbl(filtered(i, colZbrKol))
        If IsNumeric(filtered(i, colZbrAmb)) Then zv(1) = zv(1) + CLng(filtered(i, colZbrAmb))
        zbrDict(vlKey) = zv

        If Not klasePoVlasniku.Exists(vlKey) Then
            klasePoVlasniku.Add vlKey, CreateObject("Scripting.Dictionary")
        End If
        Dim klSet As Object
        Set klSet = klasePoVlasniku(vlKey)
        Dim thisKlasa As String
        thisKlasa = KlasaOrDefault(filtered(i, colZbrKla))
        If Not klSet.Exists(thisKlasa) Then klSet.Add thisKlasa, True
    Next i

    ' Ergebnis
    Dim rowCount As Long
    rowCount = zbrDict.count
    If rowCount = 0 Then
        ReportManjak = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To rowCount + 1, 1 To 6)  ' +1 UKUPNO

    Dim keys As Variant
    keys = zbrDict.keys
    Dim totalZbrKg As Double, totalPrijKg As Double

    For i = 0 To zbrDict.count - 1
        zv = zbrDict(keys(i))
        Dim zbrKg As Double: zbrKg = zv(0)
        Dim zbrAmb As Long: zbrAmb = CLng(zv(1))
        Dim rowBroj As String: rowBroj = CStr(zv(2))

        ' Vlasnik reda je ovde POZNAT (zbirna nosi i vozaca i kupca), pa se
        ' prijem cita owner-scoped. Kad broj ima jednog vlasnika koristi se
        ' agregat po (broj, klasa) -- hvata i starije prijemnice bez vlasnika.
        ' Red pokriva ceo dokument, pa se prijem sabira PO KLASAMA koje red
        ' obuhvata (prijemnice postoje po klasi, ne po dokumentu).
        Dim nVlasnika As Long: nVlasnika = 0
        If manjakDict.Exists("#V|" & rowBroj) Then nVlasnika = CLng(manjakDict("#V|" & rowBroj))

        Dim cntNejasan As Long: cntNejasan = 0
        If manjakDict.Exists("#N|" & rowBroj) Then cntNejasan = CLng(manjakDict("#N|" & rowBroj))

        Dim cntPrijem As Long: cntPrijem = 0
        Dim prijKg As Double: prijKg = 0

        Dim rowKlase As Object
        Set rowKlase = klasePoVlasniku(CStr(keys(i)))
        Dim kl As Variant
        For Each kl In rowKlase.keys
            If nVlasnika <= 1 Then
                Dim bkKey As String
                bkKey = rowBroj & "|" & CStr(kl)
                If manjakDict.Exists("#C|" & bkKey) Then cntPrijem = cntPrijem + CLng(manjakDict("#C|" & bkKey))
                If manjakDict.Exists("#K|" & bkKey) Then prijKg = prijKg + CDbl(manjakDict("#K|" & bkKey))
            Else
                Dim skKey As String
                skKey = ZbirnaStavkaKljuc(CStr(keys(i)), CStr(kl))
                If manjakDict.Exists(skKey) Then
                    Dim ownVals As Variant
                    ownVals = manjakDict(skKey)
                    prijKg = prijKg + CDbl(ownVals(1))
                    cntPrijem = cntPrijem + CLng(ownVals(2))
                End If
            End If
        Next kl

        Dim pz As Variant
        pz = PrijemZaZbirnu(nVlasnika, True, cntNejasan, cntPrijem, prijKg)
        Dim imaPrijem As Boolean: imaPrijem = CBool(pz(0))
        prijKg = CDbl(pz(1))

        Dim mStavka As Variant
        mStavka = ManjakStavka(zbrKg, prijKg, imaPrijem, CStr(pz(2)))

        Dim prosek As Double: prosek = 0
        If zbrAmb > 0 Then prosek = zbrKg / zbrAmb

        result(i + 1, 1) = rowBroj
        result(i + 1, 2) = zbrKg
        result(i + 1, 3) = mStavka(0)
        result(i + 1, 4) = mStavka(1)
        If imaPrijem Then
            result(i + 1, 5) = mStavka(2)
        Else
            result(i + 1, 5) = mStavka(3)      ' "nema prijema" / "nejasan vlasnik"
        End If
        result(i + 1, 6) = prosek

        If imaPrijem Then
            totalZbrKg = totalZbrKg + zbrKg
            totalPrijKg = totalPrijKg + prijKg
        End If
    Next i

    ' UKUPNO -- samo zbirne sa prijemom (v. napomenu na vrhu funkcije).
    result(rowCount + 1, 1) = "UKUPNO"
    result(rowCount + 1, 2) = totalZbrKg
    result(rowCount + 1, 3) = totalPrijKg
    result(rowCount + 1, 4) = totalZbrKg - totalPrijKg
    If totalZbrKg > 0 Then
        result(rowCount + 1, 5) = (totalZbrKg - totalPrijKg) / totalZbrKg * 100
    Else
        result(rowCount + 1, 5) = 0
    End If
    result(rowCount + 1, 6) = ""

    ReportManjak = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Public Function ReportZbirni(ByVal entitetTip As String, _
                             ByVal datumOd As Date, _
                             ByVal datumDo As Date) As Variant
                             
    Const SRC As String = "modIzvestaj.ReportZbirni"
    On Error GoTo EH
    
    ' Returns: 2D Array (Entitet, Info, Col3, Col4, Col5)
    '   OM:    StanicaNaziv, Vrsta, Kolicina, Vrednost, ProsekCena
    '   Kupac: KupacNaziv, Vrsta, Kolicina, Vrednost, ProsekCena
    '   Vozac: VozacIme, AmbIzlaz, AmbVracena, ManjakKg, ManjakPct
    ' Letzte Zeile = UKUPNO
    
    Select Case entitetTip
        Case "OM":    ReportZbirni = ReportZbirniOM(datumOd, datumDo)
        Case "Kupac": ReportZbirni = ReportZbirniKupac(datumOd, datumDo)
        Case "Vozac": ReportZbirni = ReportZbirniVozac(datumOd, datumDo)
        Case Else:    ReportZbirni = Empty
    End Select
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportZbirniOM(ByVal datumOd As Date, _
                                ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportZbirniOM"
    On Error GoTo EH
    
    Dim data As Variant
    data = GetTableData(TBL_OTKUP)
    If IsEmpty(data) Then
        ReportZbirniOM = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, "modIzvestaj.ReportZbirniOM"), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    Dim filtered As Variant
    filtered = FilterArray(data, filters)
    If IsEmpty(filtered) Then
        ReportZbirniOM = Empty
        Exit Function
    End If
    filtered = ExcludeStornirano(filtered, TBL_OTKUP)
    If IsEmpty(filtered) Or Not IsArray(filtered) Then
        ReportZbirniOM = Empty
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colStation As Long, colVrsta As Long, colOtkID As Long
    colStation = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, "modIzvestaj.ReportZbirniOM")
    colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "modIzvestaj.ReportZbirniOM")
    colOtkID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "modIzvestaj.ReportZbirniOM")

    ' Kolicina i vrednost su na STAVKAMA (REFAKTOR S14.7, kvar 3).
    Dim stavkeZbir As Object
    Set stavkeZbir = modOtkup.ZbirStavkiPoOtkupu()

    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim key As String
        key = CStr(filtered(i, colStation)) & "|" & CStr(filtered(i, colVrsta))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        Dim zZb As Variant
        zZb = modOtkup.ZbirStavkiZaOtkup(stavkeZbir, CStr(filtered(i, colOtkID)), _
                  "modIzvestaj.ReportZbirniOM")
        vals(0) = vals(0) + CDbl(zZb(0))
        vals(1) = vals(1) + CDbl(zZb(1))
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportZbirniOM = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 5)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totalKg As Double, totalRSD As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        Dim parts As Variant
        parts = Split(keys(i), "|")
        
        result(i + 1, 1) = CStr(LookupValue(TBL_STANICE, "StanicaID", parts(0), "Naziv"))
        result(i + 1, 2) = parts(1)
        result(i + 1, 3) = vals(0)
        result(i + 1, 4) = vals(1)
        If vals(0) > 0 Then result(i + 1, 5) = vals(1) / vals(0) Else result(i + 1, 5) = 0
        
        totalKg = totalKg + vals(0)
        totalRSD = totalRSD + vals(1)
    Next i
    
    result(dict.count + 1, 1) = ""
    result(dict.count + 1, 2) = "UKUPNO"
    result(dict.count + 1, 3) = totalKg
    result(dict.count + 1, 4) = totalRSD
    result(dict.count + 1, 5) = ""
    
    ReportZbirniOM = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportZbirniKupac(ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportZbirniKupac"
    On Error GoTo EH
    
    Dim data As Variant
    data = GetTableData(TBL_PRIJEMNICA)
    If IsEmpty(data) Then
        ReportZbirniKupac = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, "modIzvestaj.ReportZbirniKupac"), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    Dim filtered As Variant
    filtered = FilterArray(data, filters)
    If IsEmpty(filtered) Then
        ReportZbirniKupac = Empty
        Exit Function
    End If
    filtered = ExcludeStornirano(filtered, TBL_PRIJEMNICA)
    If IsEmpty(filtered) Or Not IsArray(filtered) Then
        ReportZbirniKupac = Empty
        Exit Function
    End If
    
    ' Cache fuer Vrsta-Lookup
    Dim vrstaCache As Object
    Set vrstaCache = BuildZbirnaVrstaCache()
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colKupac As Long, colKol As Long, colCena As Long, colBrZbr As Long
    colKupac = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, "modIzvestaj.ReportZbirniKupac")
    colKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "modIzvestaj.ReportZbirniKupac")
    colCena = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_CENA, "modIzvestaj.ReportZbirniKupac")
    colBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "modIzvestaj.ReportZbirniKupac")
    
    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim vrsta As String
        vrsta = GetVrstaFromCache(vrstaCache, CStr(filtered(i, colBrZbr)))
        If vrsta = "" Then vrsta = "(Nepoznato)"
        
        Dim key As String
        key = CStr(filtered(i, colKupac)) & "|" & vrsta
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(filtered(i, colKol)) Then vals(0) = vals(0) + CDbl(filtered(i, colKol))
        If IsNumeric(filtered(i, colKol)) And IsNumeric(filtered(i, colCena)) Then
            vals(1) = vals(1) + CDbl(filtered(i, colKol)) * CDbl(filtered(i, colCena))
        End If
        dict(key) = vals
    Next i
    
    If dict.count = 0 Then
        ReportZbirniKupac = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 5)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totalKg As Double, totalRSD As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        Dim parts As Variant
        parts = Split(keys(i), "|")
        
        result(i + 1, 1) = CStr(LookupValue(TBL_KUPCI, "KupacID", parts(0), "Naziv"))
        result(i + 1, 2) = parts(1)
        result(i + 1, 3) = vals(0)
        result(i + 1, 4) = vals(1)
        If vals(0) > 0 Then result(i + 1, 5) = vals(1) / vals(0) Else result(i + 1, 5) = 0
        
        totalKg = totalKg + vals(0)
        totalRSD = totalRSD + vals(1)
    Next i
    
    result(dict.count + 1, 1) = ""
    result(dict.count + 1, 2) = "UKUPNO"
    result(dict.count + 1, 3) = totalKg
    result(dict.count + 1, 4) = totalRSD
    result(dict.count + 1, 5) = ""
    
    ReportZbirniKupac = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

Private Function ReportZbirniVozac(ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Variant
    
    Const SRC As String = "modIzvestaj.ReportZbirniVozac"
    On Error GoTo EH
    
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then
        ReportZbirniVozac = Empty
        Exit Function
    End If
    
    Dim filters As New Collection
    Dim fp As clsFilterParam
    Set fp = New clsFilterParam
    fp.Init RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, "modIzvestaj.ReportZbirniVozac"), "BETWEEN", datumOd, datumDo
    filters.Add fp
    
    Dim zbrFiltered As Variant
    zbrFiltered = FilterArray(zbrData, filters)
    If IsEmpty(zbrFiltered) Then
        ReportZbirniVozac = Empty
        Exit Function
    End If
    zbrFiltered = ExcludeStornirano(zbrFiltered, TBL_ZBIRNA)
    If IsEmpty(zbrFiltered) Or Not IsArray(zbrFiltered) Then
        ReportZbirniVozac = Empty
        Exit Function
    End If
    
    ' Prijemnica-Daten EINMAL laden (Performance-Fix)
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    
    If IsArray(prijData) Then
        prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    End If
    
    Dim colPBrZbr As Long, colPAmbVr As Long, colPKol As Long
    If IsArray(prijData) And Not IsEmpty(prijData) Then
        colPBrZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, "modIzvestaj.ReportZbirniVozac")
        colPAmbVr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOL_AMB_VRACENA, "modIzvestaj.ReportZbirniVozac")
        colPKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, "modIzvestaj.ReportZbirniVozac")
    End If
    
    ' Prijemnice aggregieren po BrojZbirne:
    '   vals(0) = AmbVracena
    '   vals(1) = PrijKg
    Dim prijAgg As Object
    Set prijAgg = CreateObject("Scripting.Dictionary")
    
    Dim j As Long
    
    If IsArray(prijData) And Not IsEmpty(prijData) Then
        For j = 1 To UBound(prijData, 1)
            Dim pBrZbrAgg As String
            pBrZbrAgg = CStr(prijData(j, colPBrZbr))
            
            If pBrZbrAgg <> "" Then
                If Not prijAgg.Exists(pBrZbrAgg) Then
                    prijAgg.Add pBrZbrAgg, Array(0#, 0#)
                End If
                
                Dim prijAggVals As Variant
                prijAggVals = prijAgg(pBrZbrAgg)
                
                If IsNumeric(prijData(j, colPAmbVr)) Then
                    prijAggVals(0) = prijAggVals(0) + CLng(prijData(j, colPAmbVr))
                End If
                
                If IsNumeric(prijData(j, colPKol)) Then
                    prijAggVals(1) = prijAggVals(1) + CDbl(prijData(j, colPKol))
                End If
                
                prijAgg(pBrZbrAgg) = prijAggVals
            End If
        Next j
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim colVozac As Long, colBroj As Long, colKol As Long, colAmb As Long
    colVozac = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, "modIzvestaj.ReportZbirniVozac")
    colBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, "modIzvestaj.ReportZbirniVozac")
    colKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, "modIzvestaj.ReportZbirniVozac")
    colAmb = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOL_AMB, "modIzvestaj.ReportZbirniVozac")
    
    Dim i As Long
    For i = 1 To UBound(zbrFiltered, 1)
        Dim vozacID As String
        vozacID = CStr(zbrFiltered(i, colVozac))
        If Not dict.Exists(vozacID) Then dict.Add vozacID, Array(0#, 0#, 0#, 0#)
        ' (0)=AmbIzlaz, (1)=AmbVracena, (2)=ZbirnaKg, (3)=PrijKg
        
        Dim vals As Variant
        vals = dict(vozacID)
        
        If IsNumeric(zbrFiltered(i, colAmb)) Then vals(0) = vals(0) + CLng(zbrFiltered(i, colAmb))
        If IsNumeric(zbrFiltered(i, colKol)) Then vals(2) = vals(2) + CDbl(zbrFiltered(i, colKol))
        
        ' Prijemnica-Daten fuer diese Zbirna aus vorgeladenem Array
        Dim brZbr As String
        brZbr = CStr(zbrFiltered(i, colBroj))
        
        If prijAgg.Exists(brZbr) Then
            Dim prijLookupVals As Variant
            prijLookupVals = prijAgg(brZbr)
            
            vals(1) = vals(1) + CLng(prijLookupVals(0))   ' AmbVracena
            vals(3) = vals(3) + CDbl(prijLookupVals(1))   ' PrijKg
        End If
        
        dict(vozacID) = vals
    Next i
    
    If dict.count = 0 Then
        ReportZbirniVozac = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 5)
    
    Dim keys As Variant
    keys = dict.keys
    Dim vozacDict As Object
    Set vozacDict = BuildLookupDict(TBL_VOZACI, "VozacID", "Ime", "Prezime")
    Dim tAmbIzl As Long, tAmbVr As Long, tZbrKg As Double, tPrijKg As Double
    
    For i = 0 To dict.count - 1
        vals = dict(keys(i))
        
        Dim manjakKg As Double
        manjakKg = vals(2) - vals(3)
        Dim manjakPct As Double
        If vals(2) > 0 Then manjakPct = manjakKg / vals(2) * 100 Else manjakPct = 0
        
        Dim vozNaz As String
        If vozacDict.Exists(CStr(keys(i))) Then vozNaz = vozacDict(CStr(keys(i))) Else vozNaz = ""
        result(i + 1, 1) = vozNaz
        result(i + 1, 2) = vals(0)       ' AmbIzlaz
        result(i + 1, 3) = vals(1)       ' AmbVracena
        result(i + 1, 4) = manjakKg      ' ManjakKg
        result(i + 1, 5) = manjakPct     ' ManjakPct
        
        tAmbIzl = tAmbIzl + vals(0)
        tAmbVr = tAmbVr + vals(1)
        tZbrKg = tZbrKg + vals(2)
        tPrijKg = tPrijKg + vals(3)
    Next i
    
    ' UKUPNO
    result(dict.count + 1, 1) = "UKUPNO"
    result(dict.count + 1, 2) = tAmbIzl
    result(dict.count + 1, 3) = tAmbVr
    result(dict.count + 1, 4) = tZbrKg - tPrijKg
    If tZbrKg > 0 Then
        result(dict.count + 1, 5) = (tZbrKg - tPrijKg) / tZbrKg * 100
    Else
        result(dict.count + 1, 5) = 0
    End If
    
    ReportZbirniVozac = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' AUSGABE (unveraendert)
' ============================================================

Public Sub OutputToSheet(ByVal data As Variant, ByVal targetRange As Range, _
                         Optional ByVal headers As Variant)
    If IsEmpty(data) Then
        targetRange.value = "Nema podataka"
        Exit Sub
    End If
    
    Dim startRow As Long
    startRow = 0
    
    If Not IsMissing(headers) Then
        Dim h As Long
        For h = LBound(headers) To UBound(headers)
            targetRange.Offset(0, h - LBound(headers)).value = headers(h)
            targetRange.Offset(0, h - LBound(headers)).Font.Bold = True
        Next h
        startRow = 1
    End If
    
    Dim r As Long, c As Long
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            targetRange.Offset(startRow + r - 1, c - 1).value = data(r, c)
        Next c
    Next r
End Sub


' ============================================================
' SHARED HELPER - Dict(Key zu Array(Kg, RSD)) zu 2D Result
' ============================================================

Private Function DictToResultArray(ByVal dict As Object) As Variant
    
    Const SRC As String = "modIzvestaj.DictToResultArray"
    On Error GoTo EH
    
    ' Konvertiert Dictionary(String ? Array(Double, Double))
    ' zu 2D Array (Nr, Key, Kg, RSD) + UKUPNO
    
    If dict.count = 0 Then
        DictToResultArray = Empty
        Exit Function
    End If
    
    Dim result() As Variant
    ReDim result(1 To dict.count + 1, 1 To 4)
    
    Dim keys As Variant
    keys = dict.keys
    Dim totalKg As Double, totalRSD As Double
    
    Dim i As Long
    For i = 0 To dict.count - 1
        Dim vals As Variant
        vals = dict(keys(i))
        result(i + 1, 1) = CStr(i + 1)
        result(i + 1, 2) = keys(i)
        result(i + 1, 3) = vals(0)
        result(i + 1, 4) = vals(1)
        totalKg = totalKg + vals(0)
        totalRSD = totalRSD + vals(1)
    Next i
    
    result(dict.count + 1, 1) = ""
    result(dict.count + 1, 2) = "UKUPNO"
    result(dict.count + 1, 3) = totalKg
    result(dict.count + 1, 4) = totalRSD
    
    DictToResultArray = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' SHARED HELPER - Entitet-Name aufloesen
' ============================================================

Private Function ResolveEntitetName(ByVal entitetID As String, _
                                    ByVal entitetTip As String) As String
    Select Case entitetTip
        Case "Stanica"
            ResolveEntitetName = CStr(LookupValue(TBL_STANICE, "StanicaID", entitetID, "Naziv"))
        Case "Kupac"
            ResolveEntitetName = CStr(LookupValue(TBL_KUPCI, "KupacID", entitetID, "Naziv"))
        Case "Kooperant"
            ResolveEntitetName = CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Ime")) & " " & _
                                 CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", entitetID, "Prezime"))
        Case Else
            ResolveEntitetName = entitetID
    End Select
End Function

' ============================================================
' REVERS AMBALAZE IZ PREGLEDA (v6-ui-186) -- racun izdvojen iz
' frmIzvestaj.StampajReversAmbDok za ekran Izvestaji (novi UI). Forma
' zadrzava svoju kopiju i NE menja se (katalog par. 5 / Faza B: dve kopije
' zive namerno dok legacy ne ode). Pravila su ISTA, AUD-012 / FM-0029:
'   - argumenti reversa se rekonstruisu iz nogu ledgera (Kooperant +
'     Stanica) koje dele ReversID;
'   - STORNIRANI redovi se preskacu INLINE (bez kopije cele tblAmbalaza);
'   - tip ambalaze je DEO KLJUCA (ReversRedPripada) -- dokument sa dve vrste
'     gajbica daje dva reversa, ne jedan sa pogresnim zbirom;
'   - vise od dve noge po tipu se PRIJAVLJUJE operateru, ne sabira tiho;
'   - IDENTITET je ReversID izabranog reda pregleda (REV-IDENT-01): broj je
'     labela, a isti broj legalno nose reversi druge stanice ili drugog dana
'     (A2 red REV). Noge bira ReversStampaNoge; red bez ReversID-a se ODBIJA,
'     ne nudi na "Da" -- papir za potpis ne sme da spoji dva reversa. Ambalaza
'     uz otkup ReversID nema: njen identitet je OtkupID (DokumentID).
' ============================================================
Public Sub StampajReversAmbalaze(ByVal dokID As String, ByVal dokTip As String, _
                                 ByVal tipSel As String, _
                                 Optional ByVal reversID As String = "")
    Const SRC As String = "modIzvestaj.StampajReversAmbalaze"
    On Error GoTo EH

    Dim d As Variant: d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Sub
    Dim cDat As Long, cTip As Long, cKol As Long, cEnt As Long
    Dim cEntTip As Long, cDok As Long, cDokTip As Long, cVoz As Long
    Dim cStorno As Long, cRid As Long
    cStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    cDat = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    cTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    cKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    cEnt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    cEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    cDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cVoz = GetColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC)
    cRid = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_REVERS_ID, SRC)
    If cDok = 0 Or cDokTip = 0 Then Exit Sub

    Dim isFirma As Boolean
    isFirma = (dokTip = DOK_TIP_OM_IZLAZ_FIRMA Or dokTip = DOK_TIP_OM_ULAZ_FIRMA)

    Dim datum As Date, haveDatum As Boolean
    Dim tipAmb As String, omID As String, koopID As String
    Dim kolAmb As Long
    Dim revVozacID As String
    Dim i As Long, v As Variant

    ' Tip ambalaze = tip IZABRANOG reda pregleda. Fallback (poziv bez tipa):
    ' tip prvog aktivnog reda dokumenta (tog ReversID-a, kad je dat) -- i tada
    ' se sabira SAMO taj tip.
    tipAmb = Trim$(tipSel)
    If Len(tipAmb) = 0 Then
        For i = 1 To UBound(d, 1)
            If Trim$(CStr(d(i, cDok))) = Trim$(dokID) And _
               Trim$(CStr(d(i, cDokTip))) = Trim$(dokTip) And _
               Not IzvAmbRedStorniran(d, i, cStorno) Then
                If Len(Trim$(reversID)) = 0 Or _
                   StrComp(Trim$(NzToText(d(i, cRid))), Trim$(reversID), vbTextCompare) = 0 Then
                    tipAmb = Trim$(CStr(d(i, cTip)))
                    Exit For
                End If
            End If
        Next i
    End If

    Dim redovi As Collection, revRaz As String
    revRaz = ReversStampaNoge(dokID, dokTip, tipAmb, reversID, redovi)
    If Len(revRaz) > 0 Then Err.Raise vbObjectError + 7504, SRC, revRaz

    Dim nogeKoop As Long, nogeOM As Long
    For Each v In redovi
        i = CLng(v)
        If Not haveDatum And IsDate(d(i, cDat)) Then
            datum = CDate(d(i, cDat)): haveDatum = True
        End If
        Dim et As String: et = CStr(d(i, cEntTip))
        If et = "Stanica" Then
            omID = CStr(d(i, cEnt))
            nogeOM = nogeOM + 1
            If isFirma Then
                If IsNumeric(d(i, cKol)) Then kolAmb = kolAmb + CLng(d(i, cKol))
                If cVoz > 0 And Len(revVozacID) = 0 Then revVozacID = CStr(d(i, cVoz))
            End If
        ElseIf et = "Kooperant" Then
            koopID = CStr(d(i, cEnt))
            nogeKoop = nogeKoop + 1
            If IsNumeric(d(i, cKol)) Then kolAmb = kolAmb + CLng(d(i, cKol))
        End If
    Next v

    ' Ocekivana je po JEDNA noga sa svake strane (FM-0029 #16). Vise = duplikat
    ' ili vise generacija istog dokumenta -> zbir je verovatno naduvan.
    If nogeOM > 1 Or nogeKoop > 1 Then
        If MsgBox(Poruka("RPT_MSG_REVERS_VISE_NOGU") & vbCrLf & vbCrLf & _
                  "OM: " & nogeOM & " | kooperant: " & nogeKoop & vbCrLf & _
                  Poruka("RPT_MSG_NASTAVITI_STAMPU"), _
                  vbExclamation + vbYesNo, APP_NAME) <> vbYes Then Exit Sub
    End If

    If isFirma Then
        If Len(Trim$(omID)) = 0 Then
            Err.Raise vbObjectError + 7503, SRC, _
                      "Revers (firma) nije moguce rekonstruisati (nedostaje OM noga)."
        End If
    ElseIf Len(Trim$(koopID)) = 0 Or Len(Trim$(omID)) = 0 Then
        Err.Raise vbObjectError + 7503, SRC, _
                  "Revers nije moguce rekonstruisati (nedostaje OM ili kooperant noga)."
    End If
    If Not haveDatum Then datum = Date

    Dim omNaziv As String, koopNaziv As String, vrsta As String
    omNaziv = CStr(LookupValue(TBL_STANICE, "StanicaID", omID, "Naziv"))
    ' Uz-otkup revers: DokumentID = otkupID -> vrsta iz otkupa; standalone -> prazno.
    vrsta = CStr(LookupValue(TBL_OTKUP, COL_OTK_ID, dokID, COL_OTK_VRSTA))

    If isFirma Then
        Dim prijemF As Boolean: prijemF = (dokTip = DOK_TIP_OM_ULAZ_FIRMA)
        Dim revVozacNaziv As String
        revVozacNaziv = Trim$(CStr(LookupValue(TBL_VOZACI, "VozacID", revVozacID, "Ime")) & " " & _
                              CStr(LookupValue(TBL_VOZACI, "VozacID", revVozacID, "Prezime")))
        OutputIzdavanjeAmbalaze datum, dokID, omNaziv, omID, _
                                revVozacNaziv, "", _
                                tipAmb, kolAmb, vrsta, prijemF, "FIRMA"
        Exit Sub
    End If

    koopNaziv = Trim$(CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Ime")) & " " & _
                      CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", koopID, "Prezime")))
    Dim prijem As Boolean: prijem = (dokTip = DOK_TIP_OM_ULAZ_KOOP)
    OutputIzdavanjeAmbalaze datum, dokID, omNaziv, omID, koopNaziv, koopID, _
                            tipAmb, kolAmb, vrsta, prijem
    Exit Sub
EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Sub

' Je li red tblAmbalaza storniran -- IDENTICNO pravilo kao ExcludeStornirano
' (CStr poredjenje sa "Da"), primenjeno po redu. Kolone nema = nema storna.
Private Function IzvAmbRedStorniran(ByRef d As Variant, ByVal r As Long, _
                                    ByVal cStorno As Long) As Boolean
    If cStorno <= 0 Then Exit Function
    IzvAmbRedStorniran = (CStr(d(r, cStorno)) = "Da")
End Function

Private Sub IzvRethrow(ByVal sourceName As String, _
                       ByVal errNum As Long, _
                       ByVal errDesc As String, _
                       ByVal errSrc As String)
    On Error Resume Next
    LogErr sourceName
    On Error GoTo 0

    Err.Raise errNum, sourceName, _
              "Source=" & errSrc & " | " & errDesc
End Sub

' ============================================================
' SLEDLJIVOST (v6-ui-187) - dva read-modela lanca dokumenata.
' Konstante (SLED_OZN_* / SLEDP_*) su u deklaracionoj sekciji na vrhu.
' ============================================================

' ============================================================
' LANAC-DOKUMENT (krug 6 S14): ozbiljan A4 dokument po ugledu na
' SledljivostSablon -- zaglavlje firme, naslov, info blok korena
' (otkup/kooperant/stanica/datum + vozac/kupac/period), tabela karika
' sa nosiocima, red kompletnosti, potpis/pecat. Sopstveni list
' (_SlLanacPrint) sa EKSPLICITNIM sirinama kolona -- zajednicki
' PrintIzvestajHouse je za SIROKE liste (MERGE naslov se sece na sirinu
' uske tabele) i ne dira se.
'
' paket = modScrSledljivost.SlLanacZaPdf:
'   (0) dataS 1..5 x 1..5 (karika, broj, nosilac, kg, oznaka)
'   (1) broj redova  (2) kontekst-linija (ne stampa se ovde)
'   (3) info(0..7): broj, kooperant, stanica, datum, vozac, kupac,
'       period, oznaka ("" = potpun)
' mode: PDF/PRINT/PREVIEW -- OFF je vec odbio pozivalac.
' ============================================================
Public Sub StampajSledljivostLanacDoc(ByRef paket As Variant, ByVal mode As String)
    Const SRC As String = "modIzvestaj.StampajSledljivostLanacDoc"
    On Error GoTo EH

    Dim dataS As Variant, nR As Long, info As Variant
    dataS = paket(0)
    nR = CLng(paket(1))
    info = paket(3)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("_SlLanacPrint")
    On Error GoTo EH
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.name = "_SlLanacPrint"
    End If
    ws.Visible = xlSheetVisible

    Dim oldScr As Boolean: oldScr = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ws.cells.Clear
    ws.cells.Font.name = "Calibri"
    ws.cells.Font.Size = 10

    Const NC As Long = 5
    ws.columns(1).ColumnWidth = 14
    ws.columns(2).ColumnWidth = 18
    ws.columns(3).ColumnWidth = 30
    ws.columns(4).ColumnWidth = 11
    ws.columns(5).ColumnWidth = 20

    Dim r As Long
    r = DocSellerHeader(ws, 1, NC, NC)
    r = DocTitleBlock(ws, r, NC, Poruka("OTKUI_SLPDF_SUB"), _
                      Poruka("OTKUI_SL_LANAC_NASLOV"))

    ' Info blok korena -- levo dokument, desno prevoz/period (kao sablon).
    r = r + 1
    SlpInfo ws, r, 1, Poruka("OTKUI_SLPDF_OTKUP"), info(0)
    SlpInfo ws, r, 4, Poruka("OTKUI_SLPDF_VOZAC"), info(4)
    SlpInfo ws, r + 1, 1, Poruka("OTKUI_SLPDF_KOOP"), info(1)
    SlpInfo ws, r + 1, 4, Poruka("OTKUI_SLPDF_KUPAC"), info(5)
    SlpInfo ws, r + 2, 1, Poruka("OTKUI_SLPDF_STANICA"), info(2)
    SlpInfo ws, r + 2, 4, Poruka("OTKUI_SLPDF_PERIOD"), info(6)
    SlpInfo ws, r + 3, 1, Poruka("OTKUI_SLPDF_DATUM"), info(3)
    r = r + 5

    ' Tabela karika.
    Dim hdr As Long, i As Long, c As Long
    hdr = r
    ws.cells(hdr, 1).value = Poruka("OTKUI_HDS_KARIKA")
    ws.cells(hdr, 2).value = Poruka("OTKUI_HDI_BRDOK")
    ws.cells(hdr, 3).value = Poruka("OTKUI_HDS_NOSILAC")
    ws.cells(hdr, 4).value = Poruka("OTKUI_HD_KG")
    ws.cells(hdr, 5).value = Poruka("OTKUI_HDS_OZNAKA")
    With ws.Range(ws.cells(hdr, 1), ws.cells(hdr, NC))
        .Font.Bold = True
        .Interior.Color = DocColHeaderFill()
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ws.Range(ws.cells(hdr + 1, 1), ws.cells(hdr + nR, NC)).NumberFormat = "@"
    For i = 1 To nR
        For c = 1 To NC
            ws.cells(hdr + i, c).value = dataS(i, c)
        Next c
    Next i
    With ws.Range(ws.cells(hdr + 1, 1), ws.cells(hdr + nR, NC))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ws.Range(ws.cells(hdr + 1, 4), ws.cells(hdr + nR, 4)) _
      .HorizontalAlignment = xlRight
    r = hdr + nR + 2

    ' Kompletnost -- bold; prazna oznaka je i ovde dobra vest, receno.
    ws.cells(r, 1).value = IIf(Len(info(7)) = 0, _
        Poruka("OTKUI_SLPDF_POTPUN"), _
        Poruka("OTKUI_SLPDF_STAO") & " " & info(7))
    ws.cells(r, 1).Font.Bold = True
    r = r + 2

    ' Podnozje kao sablon: datum stampe + potpis levo, pecat desno.
    ws.cells(r, 1).value = Poruka("OTKUI_SLPDF_DATSTAMPE") & " " & _
                           Format$(Date, "dd.MM.yyyy")
    r = r + 1
    ws.cells(r, 1).value = Poruka("OTKUI_SLPDF_POTPIS") & " ____________"
    ws.cells(r, 4).value = Poruka("OTKUI_SLPDF_PECAT") & " ____________"

    On Error Resume Next
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = ws.Range(ws.cells(1, 1), ws.cells(r, NC)).Address
    End With
    On Error GoTo EH
    Application.ScreenUpdating = oldScr

    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case Else
            DocExportPdf ws, ThisWorkbook.path & "\Sledljivost_Lanac_" & _
                Replace(info(0), "/", "-") & ".pdf", True
    End Select
    Exit Sub

EH:
    Application.ScreenUpdating = True
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Sub

' Par labela+vrednost u info bloku lanac-dokumenta.
Private Sub SlpInfo(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, _
                    ByVal lbl As String, ByVal v As String)
    ws.cells(r, c).value = lbl
    ws.cells(r, c).Font.Color = DocColGray()
    ws.cells(r, c + 1).value = v
    ws.cells(r, c + 1).Font.Bold = True
End Sub

