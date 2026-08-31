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
Public Const SLED_OZN_NEFAKTURISANO As String = "nefakturisano"
Public Const SLED_OZN_KG As String = "kg razlika"
Public Const SLED_OZN_BEZ_PARCELE As String = "bez parcele"

' Klase problema u ReportSledljivostProblemi (kolona 1). ASCII kodovi;
' prikazni tekst daje ekran kroz modPoruke.
Public Const SLEDP_BEZ_OTPREMNICE As String = "OTKUP-BEZ-OTPREMNICE"
Public Const SLEDP_VEZA As String = "VEZA-NEUSAGLASENA"
Public Const SLEDP_BEZ_ZBIRNE As String = "OTPREMNICA-BEZ-ZBIRNE"
Public Const SLEDP_BROJ_DVOSMISLEN As String = "BROJ-ZBIRNE-DVOSMISLEN"
Public Const SLEDP_BEZ_PRIJEMA As String = "ZBIRNA-BEZ-PRIJEMA"
Public Const SLEDP_NEFAKTURISANA As String = "PRIJEMNICA-BEZ-FAKTURE"
Public Const SLEDP_KG_RAZLIKA As String = "KG-RAZLIKA"

' Tipovi meta sledljivosti (ReportSledljivostMete, kolona 1) -- ASCII
' kodovi za rutiranje stampe; prikazno ime daje ekran kroz modPoruke.
Public Const SLEDM_ZBIRNA As String = "ZBIRNA"
Public Const SLEDM_PALETA As String = "PALETA"
Public Const SLEDM_PRERADA As String = "PRERADA"

' Vrsta karike "zbirna" za rutu stampe u NEPOTPUNI listi ekrana. Zbirna
' nema svoju stampu, pa vrsta postoji da radnja ume da ODBIJE s razlogom
' (legacy Case Else obrazac) -- ne da bi se stampalo.
Public Const SLED_DOK_ZBIRNA As String = "Zbirna"

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
Public Function ReversRedPripada(ByVal rowDokID As String, ByVal rowDokTip As String, _
                                 ByVal rowTipAmb As String, _
                                 ByVal dokID As String, ByVal dokTip As String, _
                                 ByVal tipAmb As String) As Boolean
    If Trim$(rowDokID) <> Trim$(dokID) Then Exit Function
    If Trim$(rowDokTip) <> Trim$(dokTip) Then Exit Function
    ReversRedPripada = (AmbTipKljuc(rowTipAmb) = AmbTipKljuc(tipAmb))
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
                Dim colKoop As Long, colKol As Long, colCena As Long, colAmb As Long
                colKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, "modIzvestaj.ReportSaldoOM")
                colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "modIzvestaj.ReportSaldoOM")
                colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, "modIzvestaj.ReportSaldoOM")
                colAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB, "modIzvestaj.ReportSaldoOM")
                
                For i = 1 To UBound(otkupData, 1)
                    Dim key As String
                    key = CStr(otkupData(i, colKoop))
                    
                    If key <> "" Then
                        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#, 0#)
                        
                        Dim vals As Variant
                        vals = dict(key)
                        
                        If IsNumeric(otkupData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkupData(i, colKol))
                        If IsNumeric(otkupData(i, colKol)) And IsNumeric(otkupData(i, colCena)) Then
                            vals(1) = vals(1) + CDbl(otkupData(i, colKol)) * CDbl(otkupData(i, colCena))
                        End If
                        If IsNumeric(otkupData(i, colAmb)) Then vals(2) = vals(2) + CLng(otkupData(i, colAmb))
                        
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
            Dim colOtkKol As Long, colOtkCena As Long
            Dim colOtkVrsta As Long, colOtkKlasa As Long
            Dim colOtkBrDok As Long, colParcela As Long
            Dim colOtkID As Long

            colOtkID = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "modIzvestaj.ReportKarticaKooperanta")
            colOtkDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, "modIzvestaj.ReportKarticaKooperanta")
            colOtkKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, "modIzvestaj.ReportKarticaKooperanta")
            colOtkKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "modIzvestaj.ReportKarticaKooperanta")
            colOtkCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, "modIzvestaj.ReportKarticaKooperanta")
            colOtkVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "modIzvestaj.ReportKarticaKooperanta")
            colOtkKlasa = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, "modIzvestaj.ReportKarticaKooperanta")
            colOtkBrDok = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, "modIzvestaj.ReportKarticaKooperanta")
            colParcela = RequireColumnIndex(TBL_OTKUP, COL_OTK_PARCELA, "modIzvestaj.ReportKarticaKooperanta")

            ' Ambalaza (gajbe) za running saldo: Primljena (koop->OM) i Izdata (OM->koop).
            ' KolAmbIzdata je noviji stup -> GetColumnIndex (0 = stara sema, tretiraj kao 0).
            Dim colOtkAmb As Long, colOtkAmbIzd As Long
            colOtkAmb = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB, "modIzvestaj.ReportKarticaKooperanta")
            colOtkAmbIzd = GetColumnIndex(TBL_OTKUP, COL_OTK_KOL_AMB_IZDATA)

            For i = 1 To UBound(otkData, 1)
                If CStr(otkData(i, colOtkKoop)) = kooperantID Then
                    If IsDate(otkData(i, colOtkDat)) Then
                        Dim otkDatum As Date
                        otkDatum = CDate(otkData(i, colOtkDat))
                        
                        If otkDatum <= datumDo Then
                            Dim vr As Double
                            Dim otkKol As Double

                            vr = 0
                            otkKol = 0

                            If IsNumeric(otkData(i, colOtkKol)) Then
                                otkKol = CDbl(otkData(i, colOtkKol))
                            End If

                            If IsNumeric(otkData(i, colOtkCena)) Then
                                vr = otkKol * CDbl(otkData(i, colOtkCena))
                            End If

                            Dim opis As String
                            opis = "Otkup " & CStr(otkData(i, colOtkVrsta)) & " " & _
                                   CStr(otkData(i, colOtkKlasa)) & " " & _
                                   FmtKolicina(otkKol) & "kg"

                            ' Saldo ambalaze (gajbe): Izdata (OM->koop) - Primljena (koop->OM).
                            ' Isti smer kao kanonski entitetski saldo (modAmbalaza.GetAmbalazeStanje).
                            Dim ambPrimljena As Double, ambIzdata As Double
                            ambPrimljena = 0
                            ambIzdata = 0
                            If IsNumeric(otkData(i, colOtkAmb)) Then ambPrimljena = CDbl(otkData(i, colOtkAmb))
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

    Dim cDat As Long, cKoop As Long, cKol As Long
    Dim cVrsta As Long, cKlasa As Long, cSorta As Long
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    cKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, SRC)
    cVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)
    cKlasa = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, SRC)
    cSorta = GetColumnIndex(TBL_OTKUP, COL_OTK_SORTA)   ' opciono (schema drift -> 0)

    Dim agg As Object
    Set agg = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To UBound(otkData, 1)
        If CStr(otkData(i, cKoop)) = kooperantID Then
            If IsDate(otkData(i, cDat)) Then
                Dim d As Date: d = CDate(otkData(i, cDat))
                If d >= datumOd And d <= datumDo Then
                    Dim vrsta As String, sorta As String, klasa As String
                    vrsta = Trim$(CStr(otkData(i, cVrsta)))
                    klasa = Trim$(CStr(otkData(i, cKlasa)))
                    sorta = ""
                    If cSorta > 0 Then sorta = Trim$(CStr(otkData(i, cSorta)))

                    Dim kg As Double: kg = 0
                    If IsNumeric(otkData(i, cKol)) Then kg = CDbl(otkData(i, cKol))

                    Dim key As String: key = vrsta & "|" & sorta & "|" & klasa
                    Dim rec As Variant
                    If agg.Exists(key) Then
                        rec = agg(key)
                        rec(3) = CDbl(rec(3)) + kg
                    Else
                        rec = Array(vrsta, sorta, klasa, kg)
                    End If
                    agg(key) = rec
                End If
            End If
        End If
    Next i

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
' (modKarticaDetalji, KART_REFKEY_COL=7) i za stampu celog lista po BrDok-u.
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
    Dim cVr As Long, cKl As Long, cKol As Long, cCe As Long, cStorno As Long
    cStorno = GetColumnIndex(TBL_OTKUP, COL_STORNIRANO)
    cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    cVr = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)
    cKl = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, SRC)
    cKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, SRC)
    cCe = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, SRC)

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

                    Dim kol As Double, cena As Double
                    kol = 0: cena = 0
                    If IsNumeric(d(i, cKol)) Then kol = CDbl(d(i, cKol))
                    If IsNumeric(d(i, cCe)) Then cena = CDbl(d(i, cCe))

                    moves.Add Array( _
                        dt, _
                        NzToText(d(i, cBr)), _
                        koopNm, _
                        NzToText(d(i, cVr)), _
                        NzToText(d(i, cKl)), _
                        kol, _
                        kol * cena, _
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

' Roba po VOZACU zbirno (krug 12): otpremljeno -- Sum kg i Sum kg*cena
' nestorniranih otpremnica u opsegu, po vozacu; naziv iz tblVozaci sa
' fallback-om na ID. (1)=VozacID (2)=Naziv (3)=Kg (4)=Vrednost; UKUPNO.
Public Function ReportRobaVozaciZbirni(ByVal datumOd As Date, _
                                       ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportRobaVozaciZbirni"
    On Error GoTo EH

    Dim d As Variant, i As Long
    Dim cVoz As Long, cKol As Long, cCen As Long, cDat As Long, cStorno As Long
    d = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(d) Then Exit Function
    cVoz = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
    cKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
    cCen = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA, SRC)
    cDat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)
    cStorno = GetColumnIndex(TBL_OTPREMNICA, COL_STORNIRANO)

    Dim kg As Object, vr As Object, k As String, dv As Date
    Set kg = CreateObject("Scripting.Dictionary")
    Set vr = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(d, 1)
        If cStorno = 0 Or CStr(d(i, cStorno)) <> "Da" Then
            If IsDate(d(i, cDat)) Then
                dv = CDate(d(i, cDat))
                If dv >= datumOd And dv <= datumDo Then
                    k = Trim$(CStr(d(i, cVoz)))
                    If Len(k) > 0 Then
                        kg(k) = IzvNum(kg(k)) + IzvNum(d(i, cKol))
                        vr(k) = IzvNum(vr(k)) + IzvNum(d(i, cKol)) * IzvNum(d(i, cCen))
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
    
    Dim colVrsta As Long, colKol As Long, colBrOtp As Long
    Dim colDatum As Long, colKlasa As Long, colVozac As Long
    Dim colOtpID As Long, colBrZbirne As Long
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, "modIzvestaj.ReportOtkupRobaOM")
    colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, "modIzvestaj.ReportOtkupRobaOM")
    colBrOtp = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, "modIzvestaj.ReportOtkupRobaOM")
    colDatum = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, "modIzvestaj.ReportOtkupRobaOM")
    colKlasa = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA, "modIzvestaj.ReportOtkupRobaOM")
    colVozac = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, "modIzvestaj.ReportOtkupRobaOM")
    colOtpID = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, "modIzvestaj.ReportOtkupRobaOM")
    colBrZbirne = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, "modIzvestaj.ReportOtkupRobaOM")
    
    ' --- Otkup-Summen pro OtpremnicaID ---
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    Dim otkupDict As Object
    Set otkupDict = CreateObject("Scripting.Dictionary")
    
    If IsArray(otkupData) Then
        otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
        If IsArray(otkupData) And Not IsEmpty(otkupData) Then
            Dim colOtkOtpID As Long, colOtkKol As Long
            colOtkOtpID = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, "modIzvestaj.ReportOtkupRobaOM")
            colOtkKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "modIzvestaj.ReportOtkupRobaOM")
            
            Dim j As Long
            For j = 1 To UBound(otkupData, 1)
                Dim otpKey As String
                otpKey = CStr(otkupData(j, colOtkOtpID))
                If otpKey <> "" Then
                    If Not otkupDict.Exists(otpKey) Then otkupDict.Add otpKey, 0#
                    If IsNumeric(otkupData(j, colOtkKol)) Then
                        otkupDict(otpKey) = otkupDict(otpKey) + CDbl(otkupData(j, colOtkKol))
                    End If
                End If
            Next j
        End If
    End If
    
    ' --- Manjak pro Zbirna ---
    Dim manjakDict As Object
    Set manjakDict = BuildManjakDict()
    
    ' --- Ergebnis ---
    Dim rowCount As Long
    rowCount = UBound(otpData, 1)
    
    Dim result() As Variant
    ReDim result(1 To rowCount + 1, 1 To 12)   ' +Prijemnica kg (9), +skriveni OTP|<id> (12)
    
    Dim totOtp As Double, totBlokovi As Double
    Dim totRazlika As Double, totManjak As Double
    Dim totPrijemnica As Double
    Dim totOtpSaPrijemom As Double   ' osnovica za UKUPNO manjak % (samo redovi sa prijemom)
    Dim malinaMode As Boolean: malinaMode = IsMalinaMode()
    ' Mapa vozaca (ID -> "Ime Prezime") -- jednom, umesto LookupValue u petlji nize.
    Dim vozacDict As Object
    Set vozacDict = BuildLookupDict(TBL_VOZACI, "VozacID", "Ime", "Prezime")
    Dim i As Long

    For i = 1 To rowCount
        Dim kgOtp As Double: kgOtp = 0
        If IsNumeric(otpData(i, colKol)) Then kgOtp = CDbl(otpData(i, colKol))
        
        Dim thisOtpID As String
        thisOtpID = CStr(otpData(i, colOtpID))
        
        Dim kgBlokovi As Double: kgBlokovi = 0
        If otkupDict.Exists(thisOtpID) Then kgBlokovi = otkupDict(thisOtpID)
        
        Dim razlika As Double
        razlika = kgBlokovi - kgOtp

        ' Manjak proportional berechnen
        Dim thisBrZbirne As String
        thisBrZbirne = Trim$(CStr(otpData(i, colBrZbirne)))

        ' Vozac i klasa otpremnice -- treba za razresenje STAVKE zbirne (dole) i
        ' za prikaz, pa se citaju pre oba.
        Dim vozID As String
        vozID = Trim$(CStr(otpData(i, colVozac)))
        Dim klasaOtp As String
        klasaOtp = KlasaOrDefault(otpData(i, colKlasa))

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
        
        result(i, 1) = CDate(otpData(i, colDatum))
        result(i, 2) = CStr(otpData(i, colBrOtp))
        result(i, 3) = CStr(otpData(i, colVrsta))
        result(i, 4) = CStr(otpData(i, colKlasa))
        result(i, 5) = vozNaziv
        result(i, 6) = kgOtp
        result(i, 7) = kgBlokovi
        result(i, 8) = razlika
        result(i, 9) = mStavka(0)      ' Prijemnica kg (prazno kad nema prijema)
        result(i, 10) = mStavka(1)     ' Manjak kg     (prazno kad nema prijema)
        If imaPrijem Then
            result(i, 11) = mStavka(2) ' Manjak %
        Else
            result(i, 11) = mStavka(3) ' oznaka "nema prijema"
        End If
        result(i, 12) = "OTP|" & thisOtpID

        totOtp = totOtp + kgOtp
        totBlokovi = totBlokovi + kgBlokovi
        totRazlika = totRazlika + razlika

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
    result(rowCount + 1, 7) = totBlokovi
    result(rowCount + 1, 8) = totRazlika
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
    
    Dim colVrsta As Long, colKol As Long, colCena As Long
    colVrsta = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VRSTA, "modIzvestaj.ReportOtkupRobaVozac")
    colKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, "modIzvestaj.ReportOtkupRobaVozac")
    colCena = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_CENA, "modIzvestaj.ReportOtkupRobaVozac")
    
    Dim i As Long
    For i = 1 To UBound(otpData, 1)
        Dim key As String
        key = CStr(otpData(i, colVrsta))
        If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
        Dim vals As Variant
        vals = dict(key)
        If IsNumeric(otpData(i, colKol)) Then vals(0) = vals(0) + CDbl(otpData(i, colKol))
        If IsNumeric(otpData(i, colKol)) And IsNumeric(otpData(i, colCena)) Then
            vals(1) = vals(1) + CDbl(otpData(i, colKol)) * CDbl(otpData(i, colCena))
        End If
        dict(key) = vals
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
    Dim colEntitet As Long, colEntTip As Long
    colTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_TIP, "modIzvestaj.ReportAmbalaza")
    colKol = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA, "modIzvestaj.ReportAmbalaza")
    colSmer = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_SMER, "modIzvestaj.ReportAmbalaza")
    colDokID = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID, "modIzvestaj.ReportAmbalaza")
    colDokTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP, "modIzvestaj.ReportAmbalaza")
    colDatum = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM, "modIzvestaj.ReportAmbalaza")
    colEntitet = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET, "modIzvestaj.ReportAmbalaza")
    colEntTip = RequireColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP, "modIzvestaj.ReportAmbalaza")
    
    ' Vozac = inverzni protivpartner entiteta (Stanica / Kupac); kompletna ruta
    ' otpremnica -> prijemnica daje saldo 0. Otkup nema vozaca -> izuzet (filter).
    ' Entitetski izvestaji (OM / Kupac) koriste sirovi Smer (isVozac = False).
    Dim isVozac As Boolean
    isVozac = (entitetTip = "Vozac")

    If zbirni Then
        ReportAmbalaza = ReportAmbalazeZbirni(filtered, colTip, colKol, colSmer, colEntTip, isVozac)
    Else
        ReportAmbalaza = ReportAmbalazePojedinacni(filtered, colDatum, colEntitet, colEntTip, _
                                                    colTip, colDokID, colDokTip, colKol, colSmer, isVozac)
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
                                            ByVal colSmer As Long, ByVal isVozac As Boolean) As Variant
    
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
        Dim rec As Variant
        If grp.Exists(gkey) Then
            rec = grp(gkey)
        Else
            ' Datum, Mesto, Tip, Dokument, Ulaz, Izlaz
            Dim entMemoKey As String: entMemoKey = entTipVal & "|" & entID
            If Not nameMemo.Exists(entMemoKey) Then nameMemo.Add entMemoKey, ResolveEntitetName(entID, entTipVal)
            rec = Array(filtered(i, colDatum), CStr(nameMemo(entMemoKey)), _
                        tipv, dokIDv, 0&, 0&, dokTipv)
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
        result(r + 1, 7) = "AMB|" & CStr(rr(6)) & "|" & CStr(rr(3))
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
        
        Dim colVrsta As Long, colKol As Long, colCena As Long
        colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "modIzvestaj.ReportProsecnaCena")
        colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "modIzvestaj.ReportProsecnaCena")
        colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, "modIzvestaj.ReportProsecnaCena")
        
        For i = 1 To UBound(otkData, 1)
            Dim key As String
            key = CStr(otkData(i, colVrsta))
            If Not dict.Exists(key) Then dict.Add key, Array(0#, 0#)
            vals = dict(key)
            If IsNumeric(otkData(i, colKol)) Then vals(0) = vals(0) + CDbl(otkData(i, colKol))
            If IsNumeric(otkData(i, colKol)) And IsNumeric(otkData(i, colCena)) Then
                vals(1) = vals(1) + CDbl(otkData(i, colKol)) * CDbl(otkData(i, colCena))
            End If
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
    
    Dim colStation As Long, colVrsta As Long, colKol As Long, colCena As Long
    colStation = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, "modIzvestaj.ReportZbirniOM")
    colVrsta = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, "modIzvestaj.ReportZbirniOM")
    colKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, "modIzvestaj.ReportZbirniOM")
    colCena = RequireColumnIndex(TBL_OTKUP, COL_OTK_CENA, "modIzvestaj.ReportZbirniOM")
    
    Dim i As Long
    For i = 1 To UBound(filtered, 1)
        Dim key As String
        key = CStr(filtered(i, colStation)) & "|" & CStr(filtered(i, colVrsta))
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
'   - argumenti reversa se rekonstruisu iz dve noge ledgera (Kooperant +
'     Stanica) koje dele DokumentID;
'   - STORNIRANI redovi se preskacu INLINE (bez kopije cele tblAmbalaza);
'   - tip ambalaze je DEO KLJUCA (ReversRedPripada) -- dokument sa dve vrste
'     gajbica daje dva reversa, ne jedan sa pogresnim zbirom;
'   - vise od dve noge po tipu se PRIJAVLJUJE operateru, ne sabira tiho.
' ============================================================
Public Sub StampajReversAmbalaze(ByVal dokID As String, ByVal dokTip As String, _
                                 ByVal tipSel As String)
    Const SRC As String = "modIzvestaj.StampajReversAmbalaze"
    On Error GoTo EH

    Dim d As Variant: d = GetTableData(TBL_AMBALAZA)
    If Not IsArray(d) Then Exit Sub
    Dim cDat As Long, cTip As Long, cKol As Long, cEnt As Long
    Dim cEntTip As Long, cDok As Long, cDokTip As Long, cVoz As Long
    Dim cStorno As Long
    cStorno = GetColumnIndex(TBL_AMBALAZA, COL_STORNIRANO)
    cDat = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DATUM)
    cTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_TIP)
    cKol = GetColumnIndex(TBL_AMBALAZA, COL_AMB_KOLICINA)
    cEnt = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET)
    cEntTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_ENTITET_TIP)
    cDok = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_ID)
    cDokTip = GetColumnIndex(TBL_AMBALAZA, COL_AMB_DOK_TIP)
    cVoz = GetColumnIndex(TBL_AMBALAZA, COL_AMB_VOZAC)
    If cDok = 0 Or cDokTip = 0 Then Exit Sub

    Dim isFirma As Boolean
    isFirma = (dokTip = DOK_TIP_OM_IZLAZ_FIRMA Or dokTip = DOK_TIP_OM_ULAZ_FIRMA)

    Dim datum As Date, haveDatum As Boolean
    Dim tipAmb As String, omID As String, koopID As String
    Dim kolAmb As Long
    Dim revVozacID As String
    Dim i As Long

    ' Tip ambalaze = tip IZABRANOG reda pregleda. Fallback (poziv bez tipa):
    ' tip prvog reda dokumenta -- i tada se sabira SAMO taj tip.
    tipAmb = Trim$(tipSel)
    If Len(tipAmb) = 0 Then
        For i = 1 To UBound(d, 1)
            If Trim$(CStr(d(i, cDok))) = Trim$(dokID) And _
               Trim$(CStr(d(i, cDokTip))) = Trim$(dokTip) And _
               Not IzvAmbRedStorniran(d, i, cStorno) Then
                tipAmb = Trim$(CStr(d(i, cTip)))
                Exit For
            End If
        Next i
    End If

    Dim nogeKoop As Long, nogeOM As Long
    For i = 1 To UBound(d, 1)
        If ReversRedPripada(CStr(d(i, cDok)), CStr(d(i, cDokTip)), CStr(d(i, cTip)), _
                            dokID, dokTip, tipAmb) _
           And Not IzvAmbRedStorniran(d, i, cStorno) Then
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
        End If
    Next i

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

Private Function SledTxt(ByVal v As Variant) As String
    SledTxt = Trim$(NzToText(v))
End Function

Private Function SledDbl(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then SledDbl = CDbl(v)
End Function

' Razresenje zbirne za jednu otpremnicu-stavku -- ISTI koraci kao u
' ReportOtkupRobaOM (#V / #1 / #O kljucevi BuildManjakDict-a), izdvojeni da
' ih lanac i lista problema ne prepisuju. Fail-closed: #V>1 bez jednoznacnog
' razresenja po vozacu ostavlja razresen=False.
Private Sub SledResolveZbirna(ByVal manjakDict As Object, ByVal brZbr As String, _
                              ByVal vozID As String, ByVal klasa As String, _
                              ByRef nVlasnika As Long, ByRef razresen As Boolean, _
                              ByRef stavkaKey As String, ByRef cntNejasan As Long, _
                              ByRef cntPrijem As Long, ByRef prijKg As Double, _
                              ByRef zbirnaKg As Double)
    Dim ownVals As Variant
    nVlasnika = 0
    razresen = False
    stavkaKey = ""
    cntNejasan = 0
    cntPrijem = 0
    prijKg = 0
    zbirnaKg = 0

    If manjakDict.Exists("#V|" & brZbr) Then nVlasnika = CLng(manjakDict("#V|" & brZbr))
    If manjakDict.Exists("#N|" & brZbr) Then cntNejasan = CLng(manjakDict("#N|" & brZbr))

    If nVlasnika = 1 Then
        razresen = True
        stavkaKey = ZbirnaStavkaKljuc(CStr(manjakDict("#1|" & brZbr)), klasa)
        ' Prijem po (broj, klasa) -- dokazano jedan vlasnik; hvata i starije
        ' prijemnice bez popunjenog vlasnika.
        If manjakDict.Exists("#C|" & brZbr & "|" & klasa) Then _
            cntPrijem = CLng(manjakDict("#C|" & brZbr & "|" & klasa))
        If manjakDict.Exists("#K|" & brZbr & "|" & klasa) Then _
            prijKg = CDbl(manjakDict("#K|" & brZbr & "|" & klasa))
    ElseIf nVlasnika > 1 Then
        If manjakDict.Exists("#O|" & brZbr & "|" & vozID) Then
            razresen = True
            stavkaKey = ZbirnaStavkaKljuc(CStr(manjakDict("#O|" & brZbr & "|" & vozID)), klasa)
            If manjakDict.Exists(stavkaKey) Then
                ownVals = manjakDict(stavkaKey)
                prijKg = CDbl(ownVals(1))
                cntPrijem = CLng(ownVals(2))
            End If
        End If
    End If

    If Len(stavkaKey) > 0 Then
        If manjakDict.Exists(stavkaKey) Then
            ownVals = manjakDict(stavkaKey)
            zbirnaKg = CDbl(ownVals(0))
        End If
    End If
End Sub

' Mapa parcela: ParcelaID -> Array(KatBroj, Kultura, PovrsinaHa, GGAPStatus).
' Aktivnost se NE filtrira: sledljivost je istorijska, i neaktivna parcela je
' istina o poreklu robe (isto kao legacy TraceByZbirna).
Private Function SledParceleMapa() As Object
    Dim d As Object, src As Variant, i As Long
    Dim cId As Long, cKat As Long, cKul As Long, cPov As Long, cGgap As Long
    Dim pid As String
    Set d = CreateObject("Scripting.Dictionary")
    Set SledParceleMapa = d
    On Error Resume Next
    src = GetTableData(TBL_PARCELE)
    On Error GoTo 0
    If Not IsArray(src) Then Exit Function
    cId = GetColumnIndex(TBL_PARCELE, COL_PAR_ID)
    cKat = GetColumnIndex(TBL_PARCELE, COL_PAR_KAT_BROJ)
    cKul = GetColumnIndex(TBL_PARCELE, COL_PAR_KULTURA)
    cPov = GetColumnIndex(TBL_PARCELE, COL_PAR_POVRSINA)
    cGgap = GetColumnIndex(TBL_PARCELE, COL_PAR_GGAP)
    If cId = 0 Then Exit Function
    For i = 1 To UBound(src, 1)
        pid = SledTxt(src(i, cId))
        If Len(pid) > 0 And Not d.Exists(pid) Then
            d.Add pid, Array( _
                IIf(cKat > 0, SledTxt(src(i, cKat)), ""), _
                IIf(cKul > 0, SledTxt(src(i, cKul)), ""), _
                IIf(cPov > 0, SledDbl(src(i, cPov)), 0#), _
                IIf(cGgap > 0, SledTxt(src(i, cGgap)), ""))
        End If
    Next i
End Function

' Mapa AKTIVNIH otpremnica: OtpremnicaID -> Array(Broj, BrojZbirne, VozacID,
' Klasa, Kolicina, Datum). Storniranih NEMA u mapi -- otkup vezan za takvu
' otpremnicu nosi oznaku SLED_OZN_OTP_STORNIRANA (fail-closed, ne premoscuje).
Private Function SledOtpMapa() As Object
    Dim d As Object, src As Variant, i As Long
    Dim cId As Long, cBr As Long, cZbr As Long, cVoz As Long
    Dim cKl As Long, cKol As Long, cDat As Long
    Dim oid As String
    Set d = CreateObject("Scripting.Dictionary")
    Set SledOtpMapa = d
    src = GetTableData(TBL_OTPREMNICA)
    If Not IsArray(src) Then Exit Function
    src = ExcludeStornirano(src, TBL_OTPREMNICA)
    If Not IsArray(src) Then Exit Function
    cId = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_ID)
    cBr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ)
    cZbr = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE)
    cVoz = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC)
    cKl = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KLASA)
    cKol = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA)
    cDat = GetColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM)
    If cId = 0 Then Exit Function
    For i = 1 To UBound(src, 1)
        oid = SledTxt(src(i, cId))
        If Len(oid) > 0 And Not d.Exists(oid) Then
            d.Add oid, Array( _
                SledTxt(src(i, cBr)), SledTxt(src(i, cZbr)), _
                SledTxt(src(i, cVoz)), KlasaOrDefault(src(i, cKl)), _
                SledDbl(src(i, cKol)), src(i, cDat))
        End If
    Next i
End Function

' Zbir kg NESTORNIRANIH blokova po OtpremnicaID (svi datumi -- kg karike se
' poredi nad CELIM dokumentom, ne nad periodom prikaza).
Private Function SledBlokSumMapa(ByRef otkupData As Variant, ByVal cOtkOtp As Long, _
                                 ByVal cOtkKol As Long) As Object
    Dim d As Object, i As Long, oid As String
    Set d = CreateObject("Scripting.Dictionary")
    Set SledBlokSumMapa = d
    If Not IsArray(otkupData) Then Exit Function
    For i = 1 To UBound(otkupData, 1)
        oid = SledTxt(otkupData(i, cOtkOtp))
        If Len(oid) > 0 Then
            If Not d.Exists(oid) Then d.Add oid, 0#
            d(oid) = CDbl(d(oid)) + SledDbl(otkupData(i, cOtkKol))
        End If
    Next i
End Function

' Prijemnice po scope-u razresenja: "B|broj|klasa" (agregat bez vlasnika --
' vazi samo uz #V=1) i "S|stavkaKljuc" (pun vlasnik). Vrednost je Collection
' zapisa "PrijemnicaID|Broj|Kg|Fakturisano|FakturaID|KupacID".
Private Function SledPrijMapa() As Object
    Dim d As Object, src As Variant, i As Long
    Dim cId As Long, cBr As Long, cZbr As Long, cVoz As Long, cKup As Long
    Dim cKl As Long, cKol As Long, cFakt As Long, cFid As Long
    Dim brZbr As String, rec As String, kk As String
    Set d = CreateObject("Scripting.Dictionary")
    Set SledPrijMapa = d
    src = GetTableData(TBL_PRIJEMNICA)
    If Not IsArray(src) Then Exit Function
    src = ExcludeStornirano(src, TBL_PRIJEMNICA)
    If Not IsArray(src) Then Exit Function
    cId = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID)
    cBr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ)
    cZbr = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE)
    cVoz = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_VOZAC)
    cKup = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC)
    cKl = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KLASA)
    cKol = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA)
    cFakt = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO)
    cFid = GetColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID)
    For i = 1 To UBound(src, 1)
        brZbr = SledTxt(src(i, cZbr))
        If Len(brZbr) = 0 Then GoTo SledeciP
        rec = SledTxt(src(i, cId)) & "|" & SledTxt(src(i, cBr)) & "|" & _
              CStr(SledDbl(src(i, cKol))) & "|" & SledTxt(src(i, cFakt)) & "|" & _
              SledTxt(src(i, cFid)) & "|" & SledTxt(src(i, cKup))
        kk = "B|" & brZbr & "|" & KlasaOrDefault(src(i, cKl))
        If Not d.Exists(kk) Then d.Add kk, New Collection
        d(kk).Add rec
        kk = "S|" & ZbirnaStavkaKljuc( _
                 ZbirnaVlasnikKljuc(brZbr, SledTxt(src(i, cVoz)), SledTxt(src(i, cKup))), _
                 KlasaOrDefault(src(i, cKl)))
        If Not d.Exists(kk) Then d.Add kk, New Collection
        d(kk).Add rec
SledeciP:
    Next i
End Function

' ============================================================
' LANAC. Zrno = JEDAN nestorniran otkupni list u [datumOd, datumDo];
' kolone su razresene karike NAPRED (i podaci za projekciju PO PARCELI).
'
' Returns: 2D Array (1..N, 1..26) ili Empty:
'   1  Datum otkupa             14 Oznaka ("" = potpun lanac)
'   2  BrojDokumenta            15 OtkupID
'   3  KooperantID              16 OtpremnicaID
'   4  Kooperant (naziv)        17 ParcelaID
'   5  VrstaVoca                18 KatBroj
'   6  Klasa                    19 Kultura (parcele)
'   7  Kolicina (kg otkupa)     20 PovrsinaHa
'   8  BrojOtpremnice           21 GGAPStatus
'   9  BrojZbirne (otpremnicin) 22 BPGBroj (kooperanta)
'   10 Prijem (broj/"N prij.")  23 Stanica (naziv)
'   11 Prijem kg (ili Empty)    24 Vozac (naziv, otpremnicin)
'   12 Faktura (broj/"N fakt.") 25 Otpremnica kg (ili Empty)
'   13 Kupac (naziv vlasnika)   26 Zbirna kg stavke (ili Empty)
'
' Oznaka (14) je PRVA prekinuta/visesmislena karika (SLED_OZN_* /
' IZV_VLASNIK_NEJASAN / IZV_NEMA_PRIJEMA); potpun lanac kome kg curi na
' nekoj karici nosi SLED_OZN_KG (prag SLED_EPS_KG). NISTA se ne
' premoscuje: zbirna se cita ISKLJUCIVO iz otpremnice (otkupov denorm
' BrojZbirne sluzi samo za proveru saglasnosti -> SLED_OZN_VEZA), prijem
' se pripisuje istim fail-closed pravilom kao ReportOtkupRobaOM. Red
' UKUPNO se NE vraca: podnozje racuna ekran, stampa dodaje svoj.
' ============================================================
Public Function ReportSledljivostLanac(ByVal datumOd As Date, _
                                       ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportSledljivostLanac"
    On Error GoTo EH

    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If Not IsArray(otkupData) Then Exit Function
    otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)
    If Not IsArray(otkupData) Then Exit Function

    Dim cOtkId As Long, cOtkDat As Long, cOtkKoop As Long, cOtkSt As Long
    Dim cOtkVr As Long, cOtkKl As Long, cOtkKol As Long, cOtkBr As Long
    Dim cOtkOtp As Long, cOtkZbr As Long, cOtkPar As Long
    cOtkId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
    cOtkDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
    cOtkKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
    cOtkSt = RequireColumnIndex(TBL_OTKUP, COL_OTK_STANICA, SRC)
    cOtkVr = RequireColumnIndex(TBL_OTKUP, COL_OTK_VRSTA, SRC)
    cOtkKl = RequireColumnIndex(TBL_OTKUP, COL_OTK_KLASA, SRC)
    cOtkKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, SRC)
    cOtkBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
    cOtkOtp = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, SRC)
    cOtkZbr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE, SRC)
    cOtkPar = RequireColumnIndex(TBL_OTKUP, COL_OTK_PARCELA, SRC)

    ' Mape PRE petlje -- nijedan LookupValue po redu (par. 23.11/S5).
    Dim otpMapa As Object: Set otpMapa = SledOtpMapa()
    Dim blokSum As Object: Set blokSum = SledBlokSumMapa(otkupData, cOtkOtp, cOtkKol)
    Dim manjakDict As Object: Set manjakDict = BuildManjakDict()
    Dim prijMapa As Object: Set prijMapa = SledPrijMapa()
    Dim parcele As Object: Set parcele = SledParceleMapa()
    Dim koopMapa As Object: Set koopMapa = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    Dim bpgMapa As Object: Set bpgMapa = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, COL_KOOP_BPG)
    Dim kupMapa As Object: Set kupMapa = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
    Dim vozMapa As Object: Set vozMapa = BuildLookupDict(TBL_VOZACI, "VozacID", "Ime", "Prezime")
    Dim staMapa As Object: Set staMapa = BuildLookupDict(TBL_STANICE, "StanicaID", "Naziv")
    Dim fakMapa As Object: Set fakMapa = BuildLookupDict(TBL_FAKTURE, COL_FAK_ID, COL_FAK_BROJ)

    ' Zbir kg AKTIVNIH otpremnica po (broj, klasa) -- za karicni kg test
    ' otpremnice -> zbirna (validan samo uz #V = 1, inace se ne racuna).
    Dim otpSumBK As Object: Set otpSumBK = CreateObject("Scripting.Dictionary")
    Dim ok As Variant, oInfo As Variant
    For Each ok In otpMapa.keys
        oInfo = otpMapa(ok)
        If Len(CStr(oInfo(1))) > 0 Then
            Dim bk As String
            bk = CStr(oInfo(1)) & "|" & CStr(oInfo(3))
            If Not otpSumBK.Exists(bk) Then otpSumBK.Add bk, 0#
            otpSumBK(bk) = CDbl(otpSumBK(bk)) + CDbl(oInfo(4))
        End If
    Next ok

    Dim odN As Double, doN As Double
    odN = Int(CDbl(datumOd))
    doN = Int(CDbl(datumDo))

    ' Prvi prolaz: koliko redova ulazi u period.
    Dim i As Long, n As Long, dSer As Double
    For i = 1 To UBound(otkupData, 1)
        If IsDate(otkupData(i, cOtkDat)) Then
            dSer = Int(CDbl(CDate(otkupData(i, cOtkDat))))
            If dSer < odN Or dSer > doN Then GoTo PreskociBroj
        End If
        n = n + 1
PreskociBroj:
    Next i
    If n = 0 Then Exit Function

    Dim result() As Variant
    ReDim result(1 To n, 1 To 26)

    Dim r As Long
    Dim otpID As String, brZbr As String, blokZbr As String
    Dim vozID As String, klasa As String, otpKg As Double
    Dim nVl As Long, razresen As Boolean, stavkaKey As String
    Dim cntNej As Long, cntPr As Long, prijKg As Double, zbirnaKg As Double
    Dim oznaka As String, kupacID As String, koopID As String, pid As String
    Dim pz As Variant, pInfo As Variant
    Dim prijC As Collection, fakture As Object
    Dim prikazPrij As String, prikazFak As String
    Dim kg1 As Boolean, kg2 As Boolean

    For i = 1 To UBound(otkupData, 1)
        If IsDate(otkupData(i, cOtkDat)) Then
            dSer = Int(CDbl(CDate(otkupData(i, cOtkDat))))
            If dSer < odN Or dSer > doN Then GoTo Sledeci
        End If
        r = r + 1

        koopID = SledTxt(otkupData(i, cOtkKoop))
        result(r, 1) = otkupData(i, cOtkDat)
        result(r, 2) = SledTxt(otkupData(i, cOtkBr))
        result(r, 3) = koopID
        If koopMapa.Exists(koopID) Then result(r, 4) = Trim$(CStr(koopMapa(koopID))) Else result(r, 4) = koopID
        result(r, 5) = SledTxt(otkupData(i, cOtkVr))
        result(r, 6) = KlasaOrDefault(otkupData(i, cOtkKl))
        result(r, 7) = SledDbl(otkupData(i, cOtkKol))
        result(r, 15) = SledTxt(otkupData(i, cOtkId))

        pid = SledTxt(otkupData(i, cOtkPar))
        result(r, 17) = pid
        If parcele.Exists(pid) Then
            pInfo = parcele(pid)
            result(r, 18) = CStr(pInfo(0))
            result(r, 19) = CStr(pInfo(1))
            result(r, 20) = CDbl(pInfo(2))
            result(r, 21) = CStr(pInfo(3))
        Else
            result(r, 18) = ""
            result(r, 19) = ""
            result(r, 20) = Empty
            result(r, 21) = ""
        End If
        If bpgMapa.Exists(koopID) Then result(r, 22) = Trim$(CStr(bpgMapa(koopID))) Else result(r, 22) = ""
        Dim stId As String
        stId = SledTxt(otkupData(i, cOtkSt))
        If staMapa.Exists(stId) Then result(r, 23) = Trim$(CStr(staMapa(stId))) Else result(r, 23) = stId

        ' --- karika 2: otpremnica (iskljucivo po OtpremnicaID) ---
        oznaka = ""
        kg1 = False
        kg2 = False
        otpID = SledTxt(otkupData(i, cOtkOtp))
        blokZbr = SledTxt(otkupData(i, cOtkZbr))
        result(r, 16) = otpID
        result(r, 8) = ""
        result(r, 9) = ""
        result(r, 10) = ""
        result(r, 11) = Empty
        result(r, 12) = ""
        result(r, 13) = ""
        result(r, 24) = ""
        result(r, 25) = Empty
        result(r, 26) = Empty

        If Len(otpID) = 0 Then
            oznaka = SLED_OZN_NEPOVEZAN
        ElseIf Not otpMapa.Exists(otpID) Then
            ' Veza pokazuje na storniran ili nepostojeci dokument -- lanac
            ' STAJE ovde, nista se ne premoscuje.
            oznaka = SLED_OZN_OTP_STORNIRANA
        Else
            oInfo = otpMapa(otpID)
            result(r, 8) = CStr(oInfo(0))
            brZbr = CStr(oInfo(1))
            vozID = CStr(oInfo(2))
            klasa = CStr(oInfo(3))
            otpKg = CDbl(oInfo(4))
            result(r, 9) = brZbr
            result(r, 25) = otpKg
            If vozMapa.Exists(vozID) Then result(r, 24) = Trim$(CStr(vozMapa(vozID)))

            ' kg karika 1: blokovi <-> otpremnica (nad celim dokumentom).
            If blokSum.Exists(otpID) Then
                If Abs(CDbl(blokSum(otpID)) - otpKg) > SLED_EPS_KG Then kg1 = True
            End If

            ' Saglasnost denorma: blok koji tvrdi zbirnu koju otpremnica nema
            ' (ili drugu) je drift veze koju ReassignOtkupToOtpremnica_TX
            ' odrzava -- prijavljuje se, ne premoscuje.
            If Len(blokZbr) > 0 And UCase$(blokZbr) <> UCase$(brZbr) Then
                oznaka = SLED_OZN_VEZA
            ElseIf Len(brZbr) = 0 Then
                oznaka = SLED_OZN_BEZ_ZBIRNE
            End If

            ' --- karika 3: zbirna (vlasnicko razresenje) ---
            If Len(oznaka) = 0 Then
                SledResolveZbirna manjakDict, brZbr, vozID, klasa, _
                                  nVl, razresen, stavkaKey, cntNej, cntPr, prijKg, zbirnaKg
                If nVl = 0 Then
                    ' Broj postoji na otpremnici, a nijedna AKTIVNA zbirna ga
                    ' ne nosi. (Bez fixture vozila -- ne tvrdi se testom.)
                    oznaka = SLED_OZN_ZBIRNA_NEMA
                Else
                    result(r, 26) = zbirnaKg
                    ' kg karika 2: otpremnice <-> zbirna, samo uz #V = 1.
                    If nVl = 1 And otpSumBK.Exists(brZbr & "|" & klasa) Then
                        If Abs(CDbl(otpSumBK(brZbr & "|" & klasa)) - zbirnaKg) > SLED_EPS_KG Then kg2 = True
                    End If

                    ' --- karika 4: prijemnice (fail-closed pravilo) ---
                    pz = PrijemZaZbirnu(nVl, razresen, cntNej, cntPr, prijKg)
                    If Not CBool(pz(0)) Then
                        oznaka = CStr(pz(2))    ' IZV_VLASNIK_NEJASAN / IZV_NEMA_PRIJEMA
                    Else
                        result(r, 11) = CDbl(pz(1))
                        ' Razlika zbirna <-> prijem se NE proverava:
                        ' to je TRANSPORTNO KALO (smoke nalaz S1) --
                        ' poslovna velicina koju mere Manjak izvestaji,
                        ' ne kvar lanca. Vidljiva je u detalju (kg po
                        ' karici), ne obelezava se kao oznaka.

                        Set prijC = Nothing
                        If nVl = 1 Then
                            If prijMapa.Exists("B|" & brZbr & "|" & klasa) Then _
                                Set prijC = prijMapa("B|" & brZbr & "|" & klasa)
                        Else
                            If prijMapa.Exists("S|" & stavkaKey) Then _
                                Set prijC = prijMapa("S|" & stavkaKey)
                        End If

                        prikazPrij = ""
                        prikazFak = ""
                        Set fakture = CreateObject("Scripting.Dictionary")
                        If Not prijC Is Nothing Then
                            Dim recV As Variant, p() As String
                            For Each recV In prijC
                                p = Split(CStr(recV), "|")
                                If prijC.count = 1 Then prikazPrij = p(1)
                                If p(3) = "Da" And Len(p(4)) > 0 Then
                                    If Not fakture.Exists(p(4)) Then fakture.Add p(4), True
                                End If
                            Next recV
                            If prijC.count > 1 Then prikazPrij = CStr(prijC.count) & " prij."
                        End If
                        result(r, 10) = prikazPrij

                        ' --- karika 5: faktura (denorm FakturaID) ---
                        If fakture.count = 0 Then
                            oznaka = SLED_OZN_NEFAKTURISANO
                        ElseIf fakture.count = 1 Then
                            Dim fk As Variant
                            For Each fk In fakture.keys
                                If fakMapa.Exists(CStr(fk)) And Len(Trim$(CStr(fakMapa(CStr(fk))))) > 0 Then
                                    prikazFak = Trim$(CStr(fakMapa(CStr(fk))))
                                Else
                                    prikazFak = CStr(fk)
                                End If
                            Next fk
                        Else
                            prikazFak = CStr(fakture.count) & " fakt."
                        End If
                        result(r, 12) = prikazFak
                    End If

                    ' Kupac iz RAZRESENOG vlasnika (broj|vozac|kupac).
                    If razresen And Len(stavkaKey) > 0 Then
                        Dim vkDel() As String
                        vkDel = Split(stavkaKey, "|")
                        If UBound(vkDel) >= 2 Then
                            kupacID = vkDel(2)
                            If kupMapa.Exists(kupacID) Then
                                result(r, 13) = Trim$(CStr(kupMapa(kupacID)))
                            Else
                                result(r, 13) = kupacID
                            End If
                        End If
                    End If
                End If
            End If
        End If

        ' Oznaka je PRVA anomalija PO POZICIJI u lancu (merilo #2 zadatka:
        ' kg koji curi je vidljiva razlika, nikad precutana): prekid pre/na
        ' otpremnici -> kg blok<->otp -> prekid na zbirni -> kg otp<->zbirna
        ' -> nema prijema -> nefakturisano. Kg se poredi SAMO na podatkovnim
        ' karikama (roba se nije mrdala, brojevi moraju biti isti); zbirna
        ' <-> prijem je TRANSPORTNO KALO (smoke S1) i ne ulazi -- njega mere
        ' Manjak izvestaji. Prekid DUBLJE u lancu ne sme da sakrije kg
        ' curenje na RANIJOJ karici; svako curenje ponaosob nosi i lista
        ' problema, pa se nista ne gubi.
        Select Case oznaka
            Case SLED_OZN_NEPOVEZAN, SLED_OZN_OTP_STORNIRANA, SLED_OZN_VEZA
                ' karika pre kg1 -- ostaje
            Case SLED_OZN_BEZ_ZBIRNE, SLED_OZN_ZBIRNA_NEMA, IZV_VLASNIK_NEJASAN
                If kg1 Then oznaka = SLED_OZN_KG
            Case IZV_NEMA_PRIJEMA
                If kg1 Or kg2 Then oznaka = SLED_OZN_KG
            Case Else
                ' "" ili nefakturisano -- kg bilo koje karike je ranije.
                If kg1 Or kg2 Then oznaka = SLED_OZN_KG
        End Select
        result(r, 14) = oznaka
Sledeci:
    Next i

    ReportSledljivostLanac = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' PROBLEMI. Zrno = JEDNA prekinuta/visesmislena karika u [datumOd,
' datumDo] -- radni spisak fail-closed nalaza, dedupliran po karici.
'
' Returns: 2D Array (1..N, 1..9) ili Empty:
'   1 Klasa (SLEDP_*)   5 Kg karike (ili Empty)
'   2 Datum karike      6 Detalj (ASCII, sa brojkama)
'   3 Broj dokumenta    7 DokTip (DOK_TIP_* / SLED_DOK_ZBIRNA)
'   4 Nosilac (naziv)   8 DokID
'   9 Lanac-brojevi za PRETRAGU (krug 4 S8): brojevi karika reda koji
'     NISU vec u kolonama 3/6 (npr. broj zbirne nefakturisane
'     prijemnice) -- ekran obecava "pretraga nalazi svaki broj u lancu"
'     i na listi NEPOTPUNI. Ne prikazuje se.
'
' Klase: OTKUP-BEZ-OTPREMNICE (i veza na storniranu -- detalj kaze),
' VEZA-NEUSAGLASENA, OTPREMNICA-BEZ-ZBIRNE (i broj bez aktivne zbirne),
' BROJ-ZBIRNE-DVOSMISLEN (jednom po broju), ZBIRNA-BEZ-PRIJEMA (po
' stavki; uz #V>1 sa nepripisivim prijemnicama se NE tvrdi -- fail-closed
' i ovde), PRIJEMNICA-BEZ-FAKTURE (i "Fakturisano=Da bez FakturaID"),
' KG-RAZLIKA (po karici, prag SLED_EPS_KG; uz #V>1 se ne racuna).
' ============================================================
Public Function ReportSledljivostProblemi(ByVal datumOd As Date, _
                                          ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportSledljivostProblemi"
    On Error GoTo EH

    Dim rows As Collection
    Set rows = New Collection

    Dim odN As Double, doN As Double
    odN = Int(CDbl(datumOd))
    doN = Int(CDbl(datumDo))

    Dim koopMapa As Object: Set koopMapa = BuildLookupDict(TBL_KOOPERANTI, COL_KOOP_ID, "Ime", "Prezime")
    Dim kupMapa As Object: Set kupMapa = BuildLookupDict(TBL_KUPCI, COL_KUP_ID, COL_KUP_NAZIV)
    Dim vozMapa As Object: Set vozMapa = BuildLookupDict(TBL_VOZACI, "VozacID", "Ime", "Prezime")
    Dim otpMapa As Object: Set otpMapa = SledOtpMapa()
    Dim manjakDict As Object: Set manjakDict = BuildManjakDict()

    ' --- otkupi: bez otpremnice / mrtva veza / neusaglasen denorm ---
    Dim otkupData As Variant
    otkupData = GetTableData(TBL_OTKUP)
    If IsArray(otkupData) Then otkupData = ExcludeStornirano(otkupData, TBL_OTKUP)

    Dim cOtkId As Long, cOtkDat As Long, cOtkKoop As Long, cOtkKol As Long
    Dim cOtkBr As Long, cOtkOtp As Long, cOtkZbr As Long
    Dim i As Long, dSer As Double
    Dim koopID As String, otpID As String, blokZbr As String, naziv As String
    Dim oInfo As Variant

    Dim blokSum As Object
    Set blokSum = CreateObject("Scripting.Dictionary")

    If IsArray(otkupData) Then
        cOtkId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, SRC)
        cOtkDat = RequireColumnIndex(TBL_OTKUP, COL_OTK_DATUM, SRC)
        cOtkKoop = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOOPERANT, SRC)
        cOtkKol = RequireColumnIndex(TBL_OTKUP, COL_OTK_KOLICINA, SRC)
        cOtkBr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BR_DOK, SRC)
        cOtkOtp = RequireColumnIndex(TBL_OTKUP, COL_OTK_OTPREMNICA_ID, SRC)
        cOtkZbr = RequireColumnIndex(TBL_OTKUP, COL_OTK_BROJ_ZBIRNE, SRC)
        Set blokSum = SledBlokSumMapa(otkupData, cOtkOtp, cOtkKol)

        For i = 1 To UBound(otkupData, 1)
            If IsDate(otkupData(i, cOtkDat)) Then
                dSer = Int(CDbl(CDate(otkupData(i, cOtkDat))))
                If dSer < odN Or dSer > doN Then GoTo SledeciOtk
            End If
            koopID = SledTxt(otkupData(i, cOtkKoop))
            If koopMapa.Exists(koopID) Then naziv = Trim$(CStr(koopMapa(koopID))) Else naziv = koopID
            otpID = SledTxt(otkupData(i, cOtkOtp))
            blokZbr = SledTxt(otkupData(i, cOtkZbr))

            If Len(otpID) = 0 Then
                rows.Add Array(SLEDP_BEZ_OTPREMNICE, otkupData(i, cOtkDat), _
                               SledTxt(otkupData(i, cOtkBr)), naziv, _
                               SledDbl(otkupData(i, cOtkKol)), "", _
                               DOK_TIP_OTKUP, SledTxt(otkupData(i, cOtkId)), blokZbr)
            ElseIf Not otpMapa.Exists(otpID) Then
                rows.Add Array(SLEDP_BEZ_OTPREMNICE, otkupData(i, cOtkDat), _
                               SledTxt(otkupData(i, cOtkBr)), naziv, _
                               SledDbl(otkupData(i, cOtkKol)), _
                               "otpremnica stornirana ili ne postoji (" & otpID & ")", _
                               DOK_TIP_OTKUP, SledTxt(otkupData(i, cOtkId)), blokZbr)
            Else
                oInfo = otpMapa(otpID)
                If Len(blokZbr) > 0 And UCase$(blokZbr) <> UCase$(CStr(oInfo(1))) Then
                    rows.Add Array(SLEDP_VEZA, otkupData(i, cOtkDat), _
                                   SledTxt(otkupData(i, cOtkBr)), naziv, _
                                   SledDbl(otkupData(i, cOtkKol)), _
                                   "blok nosi zbirnu " & blokZbr & ", otpremnica " & _
                                   IIf(Len(CStr(oInfo(1))) > 0, CStr(oInfo(1)), "(prazno)"), _
                                   DOK_TIP_OTKUP, SledTxt(otkupData(i, cOtkId)), "")
                End If
            End If
SledeciOtk:
        Next i
    End If

    ' --- otpremnice: bez zbirne / broj bez aktivne zbirne / kg blokova ---
    Dim otpData As Variant
    otpData = GetTableData(TBL_OTPREMNICA)
    If IsArray(otpData) Then otpData = ExcludeStornirano(otpData, TBL_OTPREMNICA)
    If IsArray(otpData) Then
        Dim cOId As Long, cOBr As Long, cOZbr As Long, cOVoz As Long
        Dim cOKol As Long, cODat As Long
        cOId = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_ID, SRC)
        cOBr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ, SRC)
        cOZbr = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_BROJ_ZBIRNE, SRC)
        cOVoz = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_VOZAC, SRC)
        cOKol = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_KOLICINA, SRC)
        cODat = RequireColumnIndex(TBL_OTPREMNICA, COL_OTP_DATUM, SRC)

        Dim vozID As String, brZbr As String, oid As String
        Dim otpKg As Double, sumB As Double
        For i = 1 To UBound(otpData, 1)
            If IsDate(otpData(i, cODat)) Then
                dSer = Int(CDbl(CDate(otpData(i, cODat))))
                If dSer < odN Or dSer > doN Then GoTo SledeciOtp
            End If
            oid = SledTxt(otpData(i, cOId))
            vozID = SledTxt(otpData(i, cOVoz))
            If vozMapa.Exists(vozID) Then naziv = Trim$(CStr(vozMapa(vozID))) Else naziv = vozID
            brZbr = SledTxt(otpData(i, cOZbr))
            otpKg = SledDbl(otpData(i, cOKol))

            If Len(brZbr) = 0 Then
                rows.Add Array(SLEDP_BEZ_ZBIRNE, otpData(i, cODat), _
                               SledTxt(otpData(i, cOBr)), naziv, otpKg, _
                               "BrojZbirne je prazan", DOK_TIP_OTPREMNICA, oid, "")
            ElseIf Not manjakDict.Exists("#V|" & brZbr) Then
                rows.Add Array(SLEDP_BEZ_ZBIRNE, otpData(i, cODat), _
                               SledTxt(otpData(i, cOBr)), naziv, otpKg, _
                               "zbirna " & brZbr & " ne postoji medju aktivnima", _
                               DOK_TIP_OTPREMNICA, oid, "")
            End If

            If blokSum.Exists(oid) Then
                sumB = CDbl(blokSum(oid))
                If Abs(sumB - otpKg) > SLED_EPS_KG Then
                    rows.Add Array(SLEDP_KG_RAZLIKA, otpData(i, cODat), _
                                   SledTxt(otpData(i, cOBr)), naziv, otpKg, _
                                   "blokovi " & Format$(sumB, "#,##0.##") & _
                                   " kg / otpremnica " & Format$(otpKg, "#,##0.##") & " kg", _
                                   DOK_TIP_OTPREMNICA, oid, brZbr)
                End If
            End If
SledeciOtp:
        Next i
    End If

    ' --- zbirne: dvosmislen broj / stavka bez prijema / kg karike ---
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
    If IsArray(zbrData) Then
        Dim cZId As Long, cZBr As Long, cZVoz As Long, cZKup As Long
        Dim cZKl As Long, cZKol As Long, cZDat As Long
        cZId = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_ID, SRC)
        cZBr = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
        cZVoz = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, SRC)
        cZKup = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, SRC)
        cZKl = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KLASA, SRC)
        cZKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, SRC)
        cZDat = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, SRC)

        ' Zbir kg AKTIVNIH otpremnica po (broj, klasa) -- za kg karike.
        Dim otpSumBK As Object: Set otpSumBK = CreateObject("Scripting.Dictionary")
        Dim ok As Variant
        For Each ok In otpMapa.keys
            oInfo = otpMapa(ok)
            If Len(CStr(oInfo(1))) > 0 Then
                Dim bk As String
                bk = CStr(oInfo(1)) & "|" & CStr(oInfo(3))
                If Not otpSumBK.Exists(bk) Then otpSumBK.Add bk, 0#
                otpSumBK(bk) = CDbl(otpSumBK(bk)) + CDbl(oInfo(4))
            End If
        Next ok

        Dim dvosmisleni As Object
        Set dvosmisleni = CreateObject("Scripting.Dictionary")

        Dim zBr As String, zKup As String, zKla As String, zKg As Double
        Dim nVl As Long, cntNej As Long, cntPr As Long
        Dim prijKg As Double, zbirnaKg As Double
        Dim stavkaKey As String, razresen As Boolean
        For i = 1 To UBound(zbrData, 1)
            If IsDate(zbrData(i, cZDat)) Then
                dSer = Int(CDbl(CDate(zbrData(i, cZDat))))
                If dSer < odN Or dSer > doN Then GoTo SledeciZbr
            End If
            zBr = SledTxt(zbrData(i, cZBr))
            zKup = SledTxt(zbrData(i, cZKup))
            zKla = KlasaOrDefault(zbrData(i, cZKl))
            zKg = SledDbl(zbrData(i, cZKol))
            If kupMapa.Exists(zKup) Then naziv = Trim$(CStr(kupMapa(zKup))) Else naziv = zKup

            nVl = 0
            If manjakDict.Exists("#V|" & zBr) Then nVl = CLng(manjakDict("#V|" & zBr))

            If nVl > 1 And Not dvosmisleni.Exists(zBr) Then
                dvosmisleni.Add zBr, True
                rows.Add Array(SLEDP_BROJ_DVOSMISLEN, zbrData(i, cZDat), zBr, naziv, _
                               Empty, CStr(nVl) & " aktivnih vlasnika (vozac+kupac) deli broj", _
                               SLED_DOK_ZBIRNA, SledTxt(zbrData(i, cZId)), "")
            End If

            ' Prijem za OVU stavku (vlasnik reda je poznat -- red zbirne
            ' nosi svog vozaca i kupca).
            SledResolveZbirna manjakDict, zBr, SledTxt(zbrData(i, cZVoz)), zKla, _
                              nVl, razresen, stavkaKey, cntNej, cntPr, prijKg, zbirnaKg
            If nVl > 1 Then
                ' Stavka reda je poznata i bez #O razresenja.
                stavkaKey = ZbirnaStavkaKljuc( _
                    ZbirnaVlasnikKljuc(zBr, SledTxt(zbrData(i, cZVoz)), zKup), zKla)
                cntPr = 0
                prijKg = 0
                If manjakDict.Exists(stavkaKey) Then
                    Dim sVals As Variant
                    sVals = manjakDict(stavkaKey)
                    prijKg = CDbl(sVals(1))
                    cntPr = CLng(sVals(2))
                End If
            End If

            If cntPr = 0 Then
                ' Uz #V>1 sa nepripisivim prijemnicama se ne tvrdi "bez
                ' prijema" -- prijem mozda postoji a ne sme se pripisati.
                If Not (nVl > 1 And cntNej > 0) Then
                    rows.Add Array(SLEDP_BEZ_PRIJEMA, zbrData(i, cZDat), zBr, naziv, _
                                   zKg, "nijedna prijemnica za broj " & zBr & _
                                   " (klasa " & zKla & ")", _
                                   SLED_DOK_ZBIRNA, SledTxt(zbrData(i, cZId)), "")
                End If
            End If
            ' Razlika zbirna <-> prijem se NE prijavljuje: transportno
            ' kalo (smoke S1) -- poslovna velicina, meri je Manjak.

            If nVl = 1 And otpSumBK.Exists(zBr & "|" & zKla) Then
                Dim sumO As Double
                sumO = CDbl(otpSumBK(zBr & "|" & zKla))
                If Abs(sumO - zKg) > SLED_EPS_KG Then
                    rows.Add Array(SLEDP_KG_RAZLIKA, zbrData(i, cZDat), zBr, naziv, _
                                   zKg, "otpremnice " & Format$(sumO, "#,##0.##") & _
                                   " kg / zbirna " & Format$(zKg, "#,##0.##") & " kg", _
                                   SLED_DOK_ZBIRNA, SledTxt(zbrData(i, cZId)), "")
                End If
            End If
SledeciZbr:
        Next i
    End If

    ' --- prijemnice: bez fakture (i "Da" bez FakturaID) ---
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If IsArray(prijData) Then prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
    If IsArray(prijData) Then
        Dim cPId As Long, cPBr As Long, cPKup As Long, cPKol As Long
        Dim cPDat As Long, cPFakt As Long, cPFid As Long
        cPId = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_ID, SRC)
        cPBr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ, SRC)
        cPKup = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KUPAC, SRC)
        cPKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, SRC)
        cPDat = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_DATUM, SRC)
        cPFakt = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURISANO, SRC)
        cPFid = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_FAKTURA_ID, SRC)
        ' Broj zbirne ide u kolonu 9 (pretraga): "koje prijemnice moje
        ' zbirne nisu fakturisane" je pitanje smera NAZAD i na NEPOTPUNIMA.
        Dim cPZbr As Long
        cPZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)

        Dim pKup As String
        For i = 1 To UBound(prijData, 1)
            If IsDate(prijData(i, cPDat)) Then
                dSer = Int(CDbl(CDate(prijData(i, cPDat))))
                If dSer < odN Or dSer > doN Then GoTo SledeciPrj
            End If
            pKup = SledTxt(prijData(i, cPKup))
            If kupMapa.Exists(pKup) Then naziv = Trim$(CStr(kupMapa(pKup))) Else naziv = pKup

            If SledTxt(prijData(i, cPFakt)) <> "Da" Then
                rows.Add Array(SLEDP_NEFAKTURISANA, prijData(i, cPDat), _
                               SledTxt(prijData(i, cPBr)), naziv, _
                               SledDbl(prijData(i, cPKol)), "", _
                               DOK_TIP_PRIJEMNICA, SledTxt(prijData(i, cPId)), _
                               SledTxt(prijData(i, cPZbr)))
            ElseIf Len(SledTxt(prijData(i, cPFid))) = 0 Then
                ' Poznato nepotpuno stanje (PRJ-FAK-2 klasa): oznacena kao
                ' fakturisana, a broj fakture ne postoji -- karika je slepa.
                rows.Add Array(SLEDP_NEFAKTURISANA, prijData(i, cPDat), _
                               SledTxt(prijData(i, cPBr)), naziv, _
                               SledDbl(prijData(i, cPKol)), _
                               "Fakturisano=Da bez FakturaID", _
                               DOK_TIP_PRIJEMNICA, SledTxt(prijData(i, cPId)), _
                               SledTxt(prijData(i, cPZbr)))
            End If
SledeciPrj:
        Next i
    End If

    If rows.count = 0 Then Exit Function

    Dim result() As Variant, rr As Variant, r As Long, c As Long
    ReDim result(1 To rows.count, 1 To 9)
    r = 0
    For Each rr In rows
        r = r + 1
        For c = 1 To 9
            result(r, c) = rr(c - 1)
        Next c
    Next rr

    ReportSledljivostProblemi = result
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' SLEDLJIVOST PDF PO ZBIRNOJ -- postojeci sablon (WS_SLEDLJIVOST_SABLON),
' izdvojen iz frmSledljivost.PrintTracePDF za ekran Sledljivost (smoke
' krug 2: "sledljivost ima vec definisanu formu za pdf"). Forma zadrzava
' svoju kopiju i ne menja se (par. 5 / Faza B) -- isti obrazac kao
' StampajReversAmbalaze. Ponasanje je verno legacy-ju, ukljucujuci i to
' da se zaglavlje cita sa PRVOG reda tblZbirna po broju (dvosmislen broj
' zbirne na sablonu razresava operater -- pregled, ne knjizenje).
'
' Returns (ASCII kod za ekran; ovde nema MsgBox-a):
'   ""     stampa je izasla po rezimu (PDF/PRINT/PREVIEW)
'   "OFF"  SLEDLJIVOST_PRINT_MODE = OFF -- ekran to PRIJAVLJUJE
'   "NEMA" nema podataka za taj broj zbirne
' ============================================================
Public Function StampajSledljivostZbirne(ByVal brojZbirne As String) As String
    Const SRC As String = "modIzvestaj.StampajSledljivostZbirne"
    On Error GoTo EH

    Dim mode As String
    mode = DocResolveMode(GetConfigValue(CFG_SLEDLJIVOST_PRINT_MODE), "PDF")
    If mode = "OFF" Then
        StampajSledljivostZbirne = "OFF"
        Exit Function
    End If

    Dim traceData As Variant
    traceData = TraceByZbirna(brojZbirne)
    If IsEmpty(traceData) Then
        StampajSledljivostZbirne = "NEMA"
        Exit Function
    End If

    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsEmpty(zbrData) Then
        StampajSledljivostZbirne = "NEMA"
        Exit Function
    End If

    Dim colZbrBroj As Long, colZbrDatum As Long, colZbrVozac As Long
    Dim colZbrKupac As Long, colZbrVrsta As Long
    colZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
    colZbrDatum = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, SRC)
    colZbrVozac = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VOZAC, SRC)
    colZbrKupac = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KUPAC, SRC)
    colZbrVrsta = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_VRSTA, SRC)

    Dim zbrRow As Long, z As Long
    For z = 1 To UBound(zbrData, 1)
        If CStr(zbrData(z, colZbrBroj)) = brojZbirne Then zbrRow = z: Exit For
    Next z
    If zbrRow = 0 Then
        StampajSledljivostZbirne = "NEMA"
        Exit Function
    End If

    Dim vozacID As String, vozacNaziv As String
    vozacID = CStr(zbrData(zbrRow, colZbrVozac))
    vozacNaziv = Trim$(NzToText(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Ime")) & _
                 " " & NzToText(LookupValue(TBL_VOZACI, "VozacID", vozacID, "Prezime")))
    Dim kupacNaziv As String
    kupacNaziv = NzToText(LookupValue(TBL_KUPCI, COL_KUP_ID, _
                          CStr(zbrData(zbrRow, colZbrKupac)), COL_KUP_NAZIV))
    Dim datumOtpreme As String
    If IsDate(zbrData(zbrRow, colZbrDatum)) Then
        datumOtpreme = Format$(CDate(zbrData(zbrRow, colZbrDatum)), "DD.MM.YYYY")
    End If
    Dim vrsta As String
    vrsta = CStr(zbrData(zbrRow, colZbrVrsta))

    Dim prijKg As Double
    Dim prijData As Variant
    prijData = GetTableData(TBL_PRIJEMNICA)
    If Not IsEmpty(prijData) Then
        prijData = ExcludeStornirano(prijData, TBL_PRIJEMNICA)
        If IsArray(prijData) Then
            Dim colPrjZbr As Long, colPrjKol As Long, p As Long
            colPrjZbr = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_BROJ_ZBIRNE, SRC)
            colPrjKol = RequireColumnIndex(TBL_PRIJEMNICA, COL_PRJ_KOLICINA, SRC)
            For p = 1 To UBound(prijData, 1)
                If CStr(prijData(p, colPrjZbr)) = brojZbirne Then
                    If IsNumeric(prijData(p, colPrjKol)) Then _
                        prijKg = prijKg + CDbl(prijData(p, colPrjKol))
                End If
            Next p
        End If
    End If

    Dim ws As Worksheet
    Set ws = FillSledljivostSablon(brojZbirne, datumOtpreme, vozacNaziv, kupacNaziv, _
                                   vrsta, traceData, prijKg)
    If ws Is Nothing Then
        StampajSledljivostZbirne = "NEMA"
        Exit Function
    End If

    Dim pdfPath As String
    pdfPath = ThisWorkbook.path & "\Sledljivost_" & Replace(brojZbirne, "/", "-") & ".pdf"
    Select Case mode
        Case "PRINT", "PREVIEW"
            DocPrintWs ws, mode
        Case Else
            DocExportPdf ws, pdfPath, True
    End Select

    StampajSledljivostZbirne = ""
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' METE SLEDLJIVOSTI za jednu zbirnu (smoke krug 3): kojim dokumentom se
' sledljivost te robe STVARNO dokazuje.
'   - ZBIRNA: prevoz -- sablon (roba prodata dalje kao sveza).
'   - PALETA: nestornirana, NEpreradjena paleta cija stavka nosi taj
'     BrojZbirne (roba u magacinu sveze robe) -> paletni list.
'   - PRERADA: nestornirana prerada cija stavka (join po PaletaID, kao
'     modIntegritet D2) pokazuje na nadjenu paletu (roba preradjena /
'     u magacinu preradjene robe) -> preradni list.
' NISTA se ne premoscuje: veza ide iskljucivo kroz BrojZbirne na
' paletnoj stavci i PaletaID na preradnoj stavci; preradjena paleta bez
' preradne stavke ne izmislja metu (fail-closed, D2 je prijavljuje).
'
' Returns: 2D Array (1..N, 1..4) ili Empty (prazan broj):
'   1 Tip (SLEDM_*)  2 ID (broj zbirne / PaletaID / PreradaID)
'   3 Broj (prikaz: "31/2026")  4 Opis (status/tip + kg)
' Red 1 je UVEK zbirna -- to je danasnje ponasanje dugmeta.
' ============================================================
Public Function ReportSledljivostMete(ByVal brojZbirne As String) As Variant
    Const SRC As String = "modIzvestaj.ReportSledljivostMete"
    On Error GoTo EH

    If Len(Trim$(brojZbirne)) = 0 Then Exit Function

    Dim rows As Collection: Set rows = New Collection
    rows.Add Array(SLEDM_ZBIRNA, Trim$(brojZbirne), Trim$(brojZbirne), "")

    ' 1) Palete cija stavka nosi ovaj broj zbirne.
    Dim palIDs As Object: Set palIDs = CreateObject("Scripting.Dictionary")
    palIDs.CompareMode = vbTextCompare
    Dim stData As Variant, i As Long
    stData = GetTableData(TBL_PALETA_STAVKA)
    If IsArray(stData) Then stData = ExcludeStornirano(stData, TBL_PALETA_STAVKA)
    If IsArray(stData) Then
        Dim cStPal As Long, cStZbr As Long
        cStPal = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_PALETA_ID, SRC)
        cStZbr = RequireColumnIndex(TBL_PALETA_STAVKA, COL_PALS_BROJ_ZBIRNE, SRC)
        For i = 1 To UBound(stData, 1)
            If Trim$(SledTxt(stData(i, cStZbr))) = Trim$(brojZbirne) Then
                If Len(Trim$(SledTxt(stData(i, cStPal)))) > 0 Then _
                    palIDs(Trim$(SledTxt(stData(i, cStPal)))) = True
            End If
        Next i
    End If

    Dim nadjene As Object: Set nadjene = CreateObject("Scripting.Dictionary")
    nadjene.CompareMode = vbTextCompare
    If palIDs.count > 0 Then
        Dim palData As Variant
        palData = GetTableData(TBL_PALETA)
        If IsArray(palData) Then palData = ExcludeStornirano(palData, TBL_PALETA)
        If IsArray(palData) Then
            Dim cPalId As Long, cPalBroj As Long, cPalGod As Long
            Dim cPalStat As Long, cPalPre As Long, cPalNeto As Long
            cPalId = RequireColumnIndex(TBL_PALETA, COL_PAL_ID, SRC)
            cPalBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, SRC)
            cPalGod = RequireColumnIndex(TBL_PALETA, COL_PAL_GODINA, SRC)
            cPalStat = RequireColumnIndex(TBL_PALETA, COL_PAL_STATUS, SRC)
            cPalPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, SRC)
            cPalNeto = RequireColumnIndex(TBL_PALETA, COL_PAL_NETO, SRC)
            For i = 1 To UBound(palData, 1)
                If palIDs.Exists(Trim$(SledTxt(palData(i, cPalId)))) Then
                    nadjene(Trim$(SledTxt(palData(i, cPalId)))) = True
                    ' Preradjena paleta NIJE "u magacinu sveze robe" -- njena
                    ' sledljivost je preradni list (prolaz 2).
                    If UCase$(Trim$(SledTxt(palData(i, cPalPre)))) <> "DA" Then
                        rows.Add Array(SLEDM_PALETA, _
                            Trim$(SledTxt(palData(i, cPalId))), _
                            SledTxt(palData(i, cPalBroj)) & "/" & _
                                SledTxt(palData(i, cPalGod)), _
                            SledTxt(palData(i, cPalStat)) & " " & ChrW(183) & _
                                " " & FmtKolicina(SledDbl(palData(i, cPalNeto))) & _
                                " kg")
                    End If
                End If
            Next i
        End If
    End If

    ' 2) Prerade nad nadjenim paletama (join po PaletaID).
    If nadjene.count > 0 Then
        Dim preIDs As Object: Set preIDs = CreateObject("Scripting.Dictionary")
        preIDs.CompareMode = vbTextCompare
        Dim prsData As Variant
        prsData = GetTableData(TBL_PRERADA_STAVKA)
        If IsArray(prsData) Then prsData = ExcludeStornirano(prsData, TBL_PRERADA_STAVKA)
        If IsArray(prsData) Then
            Dim cPrsPre As Long, cPrsPal As Long
            cPrsPre = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PRERADA_ID, SRC)
            cPrsPal = RequireColumnIndex(TBL_PRERADA_STAVKA, COL_PRES_PALETA_ID, SRC)
            For i = 1 To UBound(prsData, 1)
                If nadjene.Exists(Trim$(SledTxt(prsData(i, cPrsPal)))) Then
                    If Len(Trim$(SledTxt(prsData(i, cPrsPre)))) > 0 Then _
                        preIDs(Trim$(SledTxt(prsData(i, cPrsPre)))) = True
                End If
            Next i
        End If
        If preIDs.count > 0 Then
            Dim preData As Variant
            preData = GetTableData(TBL_PRERADA)
            If IsArray(preData) Then preData = ExcludeStornirano(preData, TBL_PRERADA)
            If IsArray(preData) Then
                Dim cPreId As Long, cPreBroj As Long, cPreGod As Long
                Dim cPreTip As Long, cPreNeto As Long
                cPreId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, SRC)
                cPreBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, SRC)
                cPreGod = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, SRC)
                cPreTip = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, SRC)
                cPreNeto = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
                For i = 1 To UBound(preData, 1)
                    If preIDs.Exists(Trim$(SledTxt(preData(i, cPreId)))) Then
                        rows.Add Array(SLEDM_PRERADA, _
                            Trim$(SledTxt(preData(i, cPreId))), _
                            SledTxt(preData(i, cPreBroj)) & "/" & _
                                SledTxt(preData(i, cPreGod)), _
                            SledTxt(preData(i, cPreTip)) & " " & ChrW(183) & _
                                " " & FmtKolicina(SledDbl(preData(i, cPreNeto))) & _
                                " kg")
                    End If
                Next i
            End If
        End If
    End If

    Dim res() As Variant, r As Variant, n As Long
    ReDim res(1 To rows.count, 1 To 4)
    For Each r In rows
        n = n + 1
        res(n, 1) = r(0): res(n, 2) = r(1): res(n, 3) = r(2): res(n, 4) = r(3)
    Next r
    ReportSledljivostMete = res
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' ============================================================
' SVI DOKUMENTI SLEDLJIVOSTI u periodu -- ponuda za polje izbora na
' ekranu (smoke krug 3b: "mora da postoji jasno polje za izbor sa
' dropdown i filter poljem"). Operater bira ZBIRNU (roba prodata dalje
' kao sveza -- sablon), SVEZU PALETU (magacin sveze robe -- paletni
' list) ili PRERADU (roba preradjena -- preradni list) i dobija njegov
' izvestaj.
'   - Zbirne: DISTINCT BROJ (stampa sablona je po broju -- legacy
'     cmbZbirna semantika; vlasnicko razresenje radi sablon-ruta).
'   - Palete: nestornirane, NEpreradjene (preradjena nije "sveza roba").
'   - Prerade: nestornirane.
' Dokument bez validnog datuma se UKLJUCUJE (vidljiv je bolji od tiho
' sakrivenog); sa datumom mora biti u [datumOd, datumDo].
'
' Returns: 2D Array (1..N, 1..5) ili Empty:
'   1 Tip (SLEDM_*)  2 ID (broj zbirne / PaletaID / PreradaID)
'   3 Broj (prikaz)  4 Datum (serijski Double ili Empty)  5 Opis
' ============================================================
Public Function ReportSledljivostDokumenti(ByVal datumOd As Date, _
                                           ByVal datumDo As Date) As Variant
    Const SRC As String = "modIzvestaj.ReportSledljivostDokumenti"
    On Error GoTo EH

    Dim rows As Collection: Set rows = New Collection
    Dim i As Long

    ' 1) Zbirne -- distinct broj, zbir kg, najraniji datum broja.
    Dim zbrData As Variant
    zbrData = GetTableData(TBL_ZBIRNA)
    If IsArray(zbrData) Then zbrData = ExcludeStornirano(zbrData, TBL_ZBIRNA)
    If IsArray(zbrData) Then
        Dim cZbrBroj As Long, cZbrDat As Long, cZbrKol As Long
        cZbrBroj = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_BROJ, SRC)
        cZbrDat = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_DATUM, SRC)
        cZbrKol = RequireColumnIndex(TBL_ZBIRNA, COL_ZBR_KOLICINA, SRC)
        Dim brojevi As Object: Set brojevi = CreateObject("Scripting.Dictionary")
        brojevi.CompareMode = vbTextCompare
        Dim bz As String, acc As Variant
        For i = 1 To UBound(zbrData, 1)
            bz = Trim$(SledTxt(zbrData(i, cZbrBroj)))
            If Len(bz) > 0 Then
                If SledDatumUPeriodu(zbrData(i, cZbrDat), datumOd, datumDo) Then
                    If brojevi.Exists(bz) Then
                        acc = brojevi(bz)
                        acc(0) = acc(0) + SledDbl(zbrData(i, cZbrKol))
                        If IsDate(zbrData(i, cZbrDat)) Then
                            If IsEmpty(acc(1)) Or CDbl(CDate(zbrData(i, cZbrDat))) < acc(1) Then _
                                acc(1) = CDbl(CDate(zbrData(i, cZbrDat)))
                        End If
                        brojevi(bz) = acc
                    Else
                        brojevi(bz) = Array(SledDbl(zbrData(i, cZbrKol)), _
                            IIf(IsDate(zbrData(i, cZbrDat)), _
                                CDbl(CDate(zbrData(i, cZbrDat))), Empty))
                    End If
                End If
            End If
        Next i
        Dim kb As Variant
        For Each kb In brojevi.keys
            acc = brojevi(kb)
            rows.Add Array(SLEDM_ZBIRNA, CStr(kb), CStr(kb), acc(1), _
                           FmtKolicina(CDbl(acc(0))) & " kg")
        Next kb
    End If

    ' 2) Sveze palete (nestornirane, NEpreradjene).
    Dim palData As Variant
    palData = GetTableData(TBL_PALETA)
    If IsArray(palData) Then palData = ExcludeStornirano(palData, TBL_PALETA)
    If IsArray(palData) Then
        Dim cPalId As Long, cPalBroj As Long, cPalGod As Long
        Dim cPalStat As Long, cPalPre As Long, cPalNeto As Long, cPalDat As Long
        cPalId = RequireColumnIndex(TBL_PALETA, COL_PAL_ID, SRC)
        cPalBroj = RequireColumnIndex(TBL_PALETA, COL_PAL_BROJ, SRC)
        cPalGod = RequireColumnIndex(TBL_PALETA, COL_PAL_GODINA, SRC)
        cPalStat = RequireColumnIndex(TBL_PALETA, COL_PAL_STATUS, SRC)
        cPalPre = RequireColumnIndex(TBL_PALETA, COL_PAL_PRERADJENO, SRC)
        cPalNeto = RequireColumnIndex(TBL_PALETA, COL_PAL_NETO, SRC)
        cPalDat = RequireColumnIndex(TBL_PALETA, COL_PAL_DATUM, SRC)
        For i = 1 To UBound(palData, 1)
            If UCase$(Trim$(SledTxt(palData(i, cPalPre)))) <> "DA" Then
                If SledDatumUPeriodu(palData(i, cPalDat), datumOd, datumDo) Then
                    rows.Add Array(SLEDM_PALETA, Trim$(SledTxt(palData(i, cPalId))), _
                        SledTxt(palData(i, cPalBroj)) & "/" & SledTxt(palData(i, cPalGod)), _
                        IIf(IsDate(palData(i, cPalDat)), _
                            CDbl(CDate(palData(i, cPalDat))), Empty), _
                        SledTxt(palData(i, cPalStat)) & " " & ChrW(183) & " " & _
                            FmtKolicina(SledDbl(palData(i, cPalNeto))) & " kg")
                End If
            End If
        Next i
    End If

    ' 3) Prerade (nestornirane).
    Dim preData As Variant
    preData = GetTableData(TBL_PRERADA)
    If IsArray(preData) Then preData = ExcludeStornirano(preData, TBL_PRERADA)
    If IsArray(preData) Then
        Dim cPreId As Long, cPreBroj As Long, cPreGod As Long
        Dim cPreTip As Long, cPreNeto As Long, cPreDat As Long
        cPreId = RequireColumnIndex(TBL_PRERADA, COL_PRE_ID, SRC)
        cPreBroj = RequireColumnIndex(TBL_PRERADA, COL_PRE_BROJ, SRC)
        cPreGod = RequireColumnIndex(TBL_PRERADA, COL_PRE_GODINA, SRC)
        cPreTip = RequireColumnIndex(TBL_PRERADA, COL_PRE_TIP_GP, SRC)
        cPreNeto = RequireColumnIndex(TBL_PRERADA, COL_PRE_NETO_IZLAZ, SRC)
        cPreDat = RequireColumnIndex(TBL_PRERADA, COL_PRE_DATUM, SRC)
        For i = 1 To UBound(preData, 1)
            If SledDatumUPeriodu(preData(i, cPreDat), datumOd, datumDo) Then
                rows.Add Array(SLEDM_PRERADA, Trim$(SledTxt(preData(i, cPreId))), _
                    SledTxt(preData(i, cPreBroj)) & "/" & SledTxt(preData(i, cPreGod)), _
                    IIf(IsDate(preData(i, cPreDat)), _
                        CDbl(CDate(preData(i, cPreDat))), Empty), _
                    SledTxt(preData(i, cPreTip)) & " " & ChrW(183) & " " & _
                        FmtKolicina(SledDbl(preData(i, cPreNeto))) & " kg")
            End If
        Next i
    End If

    If rows.count = 0 Then Exit Function
    Dim res() As Variant, r As Variant, n As Long
    ReDim res(1 To rows.count, 1 To 5)
    For Each r In rows
        n = n + 1
        res(n, 1) = r(0): res(n, 2) = r(1): res(n, 3) = r(2)
        res(n, 4) = r(3): res(n, 5) = r(4)
    Next r
    ReportSledljivostDokumenti = res
    Exit Function

EH:
    IzvRethrow SRC, Err.Number, Err.description, Err.SOURCE
End Function

' Datum dokumenta u [od, do]? Nevalidan datum PROLAZI -- dokument bez
' datuma je anomalija koja mora ostati VIDLJIVA u ponudi, ne tiho skrivena.
Private Function SledDatumUPeriodu(ByVal v As Variant, ByVal datumOd As Date, _
                                   ByVal datumDo As Date) As Boolean
    If Not IsDate(v) Then
        SledDatumUPeriodu = True
    Else
        SledDatumUPeriodu = (CDate(v) >= datumOd And CDate(v) <= datumDo)
    End If
End Function

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
    SlpInfo ws, r, 1, Poruka("OTKUI_SLPDF_OTKUP"), NzS(info(0))
    SlpInfo ws, r, 4, Poruka("OTKUI_SLPDF_VOZAC"), NzS(info(4))
    SlpInfo ws, r + 1, 1, Poruka("OTKUI_SLPDF_KOOP"), NzS(info(1))
    SlpInfo ws, r + 1, 4, Poruka("OTKUI_SLPDF_KUPAC"), NzS(info(5))
    SlpInfo ws, r + 2, 1, Poruka("OTKUI_SLPDF_STANICA"), NzS(info(2))
    SlpInfo ws, r + 2, 4, Poruka("OTKUI_SLPDF_PERIOD"), NzS(info(6))
    SlpInfo ws, r + 3, 1, Poruka("OTKUI_SLPDF_DATUM"), NzS(info(3))
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
    ws.cells(r, 1).value = IIf(Len(NzS(info(7))) = 0, _
        Poruka("OTKUI_SLPDF_POTPUN"), _
        Poruka("OTKUI_SLPDF_STAO") & " " & NzS(info(7)))
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
                Replace(NzS(info(0)), "/", "-") & ".pdf", True
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

