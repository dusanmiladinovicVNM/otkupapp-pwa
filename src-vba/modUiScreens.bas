Attribute VB_Name = "modUiScreens"
'=====================================================================
' modUiScreens - REGISTAR EKRANA i ugovor prema njima (faza S3a).
'
' Jedno mesto koje zna koji ekrani postoje, kako se zovu, u koju grupu
' sidebara idu i koja im je oblast za dozvole. Sidebar se gradi odavde,
' ne iz nabrojanih nizova u BuildNav.
'
' KLJUCNO - LJUSKA NE SME DA ZNA NIJEDAN EKRAN PO IMENU. Svi pozivi ka
' ekranskom modulu idu kroz Application.Run, dakle KASNO VEZANO. Da je
' vezivanje rano, klijent kome jedan ekranski modul nedostaje ne bi se
' kompajlirao i pao bi ceo StartApp - ista klasa greske kao zamka #19 u
' CLAUDE.md. Ovako ekran koji nedostaje samo bude prikazan kao neaktivan.
'
' UGOVOR EKRANA (sve je opciono osim Scr_Meta):
'   Scr_Meta()                 -> opis ekrana; sluzi i kao provera da modul
'                                 postoji i da je ekran spreman
'   Scr_Build(z)               -> izgradi kontrole u svoju zonu (jednom)
'   Scr_Layout(z, w, h)        -> rasporedi; vraca zauzetu visinu
'   Scr_Rows(filter, q)        -> Array(kolone, redovi, n, zbirKg, zbirVal)
'                                 za DELJENU mrezu ljuske
'   Scr_Liste()                -> prekidac lista ekrana; niz redova
'                                 "KLJUC|natpis|naslov mreze|sirina"
'   Scr_Lista()                -> kljuc aktivne liste
'   Scr_Radnje()               -> radnje nad redom za AKTIVNU listu; redovi
'                                 "kljuc:natpis:sirina:stil:trebaRed" spojeni "|"
'   Scr_Event(tag, ev)         -> obradi klik; True ako je obradio
'   Scr_Dozvoljen()            -> dodatna brana ekrana (npr. administracija);
'                                 ekran koji je nema je dozvoljen
'   Scr_Sort()                 -> podrazumevani sort aktivne liste, "kol:asc"
'   Scr_Save()                 -> upisi; "" ako je proslo, inace greska
'   Scr_ResetCache()           -> zaboravi izvedene mape (posle upisa)
'
' Od S4b SVI ekrani daju svoje redove kroz Scr_Rows - i ekran dokumenata,
' koji je do tada jedini punio mrezu po svom. Mreza je time neutralna.
'
' Od v6-ui-213 SVAKI red registra ima svoj modul: MARZA je bio poslednji bez
' njega i zamenio ga je ANALIZA. Sposobnost da se ekran bez modula nacrta
' prigusen OSTAJE (v. ScrPostoji) -- ona je tu za ekran koji tek dolazi, ne za
' zatecen red. Prazno polje modula danas znaci PANEL, i test 165 to tvrdi u oba
' smera.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const UISCR_BUILD As String = "v6-ui-207"

' Redosled polja u redu registra
Public Const SCR_KLJUC   As Long = 0
Public Const SCR_MODUL   As Long = 1
Public Const SCR_NASLOV  As Long = 2
Public Const SCR_IKONICA As Long = 3
Public Const SCR_GRUPA   As Long = 4
Public Const SCR_OBLAST  As Long = 5
' SEKCIJA ljuske -- koji SKUP stavki sidebara nosi ovaj ekran. Uvedena zato sto
' sidebar NEMA SKROL: na MIN_H (620) ostaje 492 pt za stavke, zauzeto je 405, pa
' bi maticne stavke (5 stavki + 2 grupe = 199 pt) ispale ispod profila -- tiho.
' Dva skupa se nikad ne crtaju zajedno; zlatno dugme u zaglavlju ih menja.
' V. docs/UI_MIGRACIJA_KATALOG.md, 26.1 i 26.2.
Public Const SCR_SEKCIJA As Long = 6

' Nosi li ekran SOPSTVENU branu (Scr_Dozvoljen). "1" = nosi, sve ostalo = ne.
'
' Do v6-ui-204 se to zakljucivalo iz BROJA GRESKE: 1004 pri pozivu je citano kao
' "ekran nema tu funkciju". Ali 1004 je obicna Excel greska koju i POSTOJECA
' brana moze da digne iznutra -- i tada bi ekran bio propusten. Sposobnost je
' zato podatak u registru, a ne pogodjena iz nacina na koji je poziv pukao.
' Test 179 tvrdi da se polje i stvarnost slazu, u oba smera.
Public Const SCR_BRANA   As Long = 7

' Redosled polja u redu GRUPE sidebara
Public Const SCRG_KLJUC   As Long = 0
Public Const SCRG_NASLOV  As Long = 1
Public Const SCRG_SEKCIJA As Long = 2

Public Const SEK_RAD     As String = "RAD"
Public Const SEK_MATICNI As String = "MATICNI"

' Poslednja greska iz ekrana: clan ugovora koji se NIJE izvrsio (v. ScrRunIshod).
' Gusenje BEZ TRAGA znaci da dugme "ne radi" i niko ne zna zasto - tacno to se
' desilo sa "Po datumu". Ljuska ovo cita i prikazuje.
Public ScrLastErr As String

' Greska koju je clan ugovora OSTAVIO u Err, a ipak VRATIO rezultat (KI-011 za
' Scr_Save, KI-014 za ostale). To nije neuspeh clana nego kvar koda -- neki
' pomocni korak je gresku prigusio i nije je obrisao. Ne prikazuje se operateru
' (clan JESTE odgovorio), ali se ne guta: ide u log kao WARN, a test ga tvrdi.
' Nosi POSLEDNJI zaostatak; prazni ga samo ScrSave, na pocetku upisa.
Public ScrZaostalaGreska As String

' Ishod Application.Run nad clanom ugovora (ScrRunIshod). Redosled je deo
' ugovora: sve do SCR_RUN_PRAZNO znaci "clan se izvrsio".
Public Const SCR_RUN_VRATIO As Long = 0  ' vratio vrednost (zaostatak zabelezen)
Public Const SCR_RUN_PRAZNO As Long = 1  ' vratio Empty, Err = 0 (npr. ekran bez liste)
Public Const SCR_RUN_NEMA   As Long = 2  ' nije se izvrsio: 1004, clana nema
Public Const SCR_RUN_PAO    As Long = 3  ' nije se izvrsio iz drugog razloga (npr. 449)

' Kes odgovora "da li modul postoji" - Application.Run na nepostojeci modul
' baca gresku, a to je skupo raditi pri svakom crtanju sidebara.
Private mHas As Object

' TEST SEAM-ovi. Oba smeju SAMO da ZABRANE -- nema vrednosti kojom se pristup
' ili sposobnost dodaje.
'
' DEJSTVO IM JE VEZANO ZA TEST REZIM, ne samo postavljanje: potrosaci pitaju
' IsTestMode() pri SVAKOM citanju, pa seam koji ostane postavljen (test koji
' je pukao pre ciscenja) postaje inertan u trenutku kad RunAllTests vrati
' SetTestMode prevMode. Bez toga bi zaboravljen seam ostavio aplikaciju
' degradiranu do restarta -- a ciscenje kroz EH svakog testa je obecanje koje
' se ne moze proveriti.
'
' mSekcijaZabranjenaTest: sekcija u kojoj nijedan ekran nije dostupan. To se u
' harnessu ne moze izazvati drugacije (AUTH je iskljucen, pa KorisnikImaPravo
' vraca True za svakoga), a bas na tom stanju stoji vracanje sekcije u ljusci.
Private mSekcijaZabranjenaTest As String

' mGradnjaPadaTest: Scr_Build "pada" za svaki ekran. Modeluje ekran koji se ne
' izgradi -- jedini put kojim ActivateScreen moze da vrati False za ekran koji
' je i dostupan i postoji, pa je i jedini nacin da se SEKCIJSKI prelazak obori
' kroz pravi prekidac (v. T_Sekcija_OdbijenPrelazakNePomeraSekciju).
Private mGradnjaPadaTest As Boolean

' mCuriTest (KI-014): clan ugovora cije telo na svom izlazu pozove
' ScrCuriAkoTest sa ovim imenom VRACA rezultat i ostavlja gresku u Err -- oblik
' koji je sonda izmerila (proba pod Resume Next, izlaz van handlera). Greska nije
' dignuta, pa kroz Application.Run stize do citaoca. Ne dodaje ni pristup ni
' sposobnost; prazno gasi seam.
Private mCuriTest As String

' mPadTest (KI-016): clan ugovora cije telo na svom vrhu pozove ScrPadAkoTest sa
' ovim imenom DIGNE gresku -- podmetnut pad. Clan koji ga ne uhvati pusta gresku
' pozivaocu, a kroz Application.Run je to runtime dijalog i reset projekta. Zato
' test clan BEZ brane zove DIREKTNO, nikad kroz Run. Prazno gasi seam.
Private mPadTest As String

' KANAL PADA (KI-016, UI-PAD-01). Clan ugovora koji uhvati svoju gresku upisuje
' je ovde (ScrPadZabelezi), a citalac je PREUZIMA posle Run-a, PRE nego sto
' pogleda vracenu vrednost. Vrednost ne moze da nosi pad (sonda 15.10.2026):
' posle handlera clana je Err = 0, Boolean bez rezultata je False (dakle
' "vratio"), a clan koji je vec dodelio rezultat pa pao vraca taj rezultat.
Private mPadClana As String

'------------------------------------------------------------ REGISTAR
' kljuc | modul | naslov (kljuc kataloga) | MDL2 kod | grupa | oblast | sekcija
'       | brana ("1" = ekran nosi svoj Scr_Dozvoljen; v. SCR_BRANA)
Public Function ScrRows() As Variant
    Dim c As Collection: Set c = New Collection
    c.Add "DOKUMENTI|modScrDokumenti|OTKUI_NAV_UNOS|" & IC_OTKUP & _
          "|OPERACIJE|" & OBL_DOKUMENTA & "|" & SEK_RAD
    c.Add "PALETE|modScrPalete|OTKUI_NAV_PALETE|" & IC_PALETE & _
          "|OPERACIJE|" & OBL_PALETE & "|" & SEK_RAD
    ' Storno je do v6-ui-141 bio rezim F8 unosnog ekrana. Zaseban ekran zato sto
    ' NIJE unos: forma i "Sacuvaj" mu ne pripadaju (Scr_Save je za njega padao u
    ' Case Else), a pregled posledica pre odluke trazi svoju zonu -- cetiri moda,
    ' lanac, palete i blokovi ne staju u MsgBox. Oblast je OBL_DOKUMENTA: ko sme
    ' da unese dokument, sme i da ga stornira.
    c.Add "STORNO|modScrStorno|OTKUI_NAV_STORNO|" & IC_STORNO & _
          "|OPERACIJE|" & OBL_DOKUMENTA & "|" & SEK_RAD
    ' Oporavak stoji uz Dokumenta i po oblasti prava: sve sto radi je
    ' prevezivanje i vracanje DOKUMENATA, pa ko sme da ih unosi sme i da ih
    ' popravi. Zaseban ekran, a ne jos jedan rezim, jer ovo nisu dokumenti nego
    ' POSAO koji ceka - i ne bira se po tipu nego po problemu.
    c.Add "OPORAVAK|modScrOporavak|OTKUI_NAV_OPORAVAK|" & IC_OPORAVAK & _
          "|OPERACIJE|" & OBL_DOKUMENTA & "|" & SEK_RAD
    c.Add "AGRO|modScrAgro|OTKUI_NAV_AGRO|" & IC_AGRO & _
          "|OPERACIJE|" & OBL_AGROHEMIJA & "|" & SEK_RAD
    c.Add "FAKTURE|modScrFakture|OTKUI_NAV_FAKT|" & IC_FAKT & _
          "|FINANSIJE|" & OBL_FAKTURISANJE & "|" & SEK_RAD
    c.Add "BANKA_UVOZ|modScrBankaUvoz|OTKUI_NAV_BANKA_UVOZ|" & IC_UVOZ & _
          "|FINANSIJE|" & OBL_BANKA & "|" & SEK_RAD
    c.Add "BANKA_NALOZI|modScrBankaNalozi|OTKUI_NAV_BANKA_NALOZI|" & IC_NALOZI & _
          "|FINANSIJE|" & OBL_BANKA & "|" & SEK_RAD
    ' ANALIZA je do v6-ui-213 bila MARZA, red koji je pokazivao na modul koji
    ' nikad nije napisan. Oblast prava ostaje OBL_MARZA -- to je NAZIV KOLONE u
    ' tblKorisnici, isti u svakoj instalaciji; pravo je isto, promenio se ekran
    ' koji ga trosi, pa preimenovanje kolone ne bi kupilo nista.
    c.Add "ANALIZA|modScrAnaliza|OTKUI_NAV_ANALIZA|" & IC_ANALIZA & _
          "|FINANSIJE|" & OBL_MARZA & "|" & SEK_RAD
    c.Add "IZVESTAJI|modScrIzvestaji|OTKUI_NAV_IZVESTAJI|" & IC_IZVEST & _
          "|ANALITIKA|" & OBL_IZVESTAJI & "|" & SEK_RAD
    ' SLEDLJIVOST obrisana u S1b-3: lanac je krenuo od veze Otkup.OtpremnicaID; vraca S9.

    ' ---- SEKCIJA MATICNI ------------------------------------------------
    ' Ono sto danas stoji iza zlatnog dugmeta: frmMaticniPodaci (popup meni,
    ' 16 sekcija u 4 grupe) -> frmStammdaten. Granica ekrana je LEGACY GRUPA,
    ' jer MAX_SEG (11) ne prima svih 13 sekcija sa podacima u jedan ekran, a
    ' grupisanje je u modMaticniLookups uvedeno svesno i operater ga zna.
    ' Plan i obrazlozenje: docs/UI_MIGRACIJA_KATALOG.md, 26.3.
    '
    ' Od M4 svih pet ima svoj modul; prigusene u sekciji MATICNI vise nema.
    ' Od v6-ui-213 je nema nigde -- svaki red registra ima modul. Sposobnost
    ' prigusivanja ostaje za ekran koji tek dolazi (v. vrh modula).
    c.Add "MAT_PARTNERI|modScrMatPartneri|OTKUI_NAV_MAT_PARTNERI|" & IC_MAT_PARTNERI & _
          "|SIFARNICI|" & OBL_MATICNI & "|" & SEK_MATICNI
    c.Add "MAT_ROBA|modScrMatRoba|OTKUI_NAV_MAT_ROBA|" & IC_MAT_ROBA & _
          "|SIFARNICI|" & OBL_MATICNI & "|" & SEK_MATICNI
    c.Add "MAT_PAKOVANJE|modScrMatPakovanje|OTKUI_NAV_MAT_PAKOVANJE|" & IC_MAT_PAKOVANJE & _
          "|SIFARNICI|" & OBL_MATICNI & "|" & SEK_MATICNI
    ' Korisnici, Podesavanja i Admin nose JOS JEDNU branu preko oblasti --
    ' administraciju. Ona ne moze u SCR_OBLAST (to je naziv kolone prava u
    ' tblKorisnici), pa se odgovara kroz neobavezan Scr_Dozvoljen za ekran,
    ' odnosno kroz modUiPanel za panel (v. ScrDozvoljen nize).
    c.Add "MAT_KORISNICI|modScrMatKorisnici|OTKUI_NAV_MAT_KORISNICI|" & IC_MAT_KORISNICI & _
          "|SISTEM|" & OBL_MATICNI & "|" & SEK_MATICNI & "|1"
    ' PANEL, ne ekran: modul je prazan jer ovi redovi nemaju Scr_* ugovor.
    ' Ljuska ih prepoznaje po tome sto ih modUiPanel poznaje, pa klik u sidebaru
    ' otvara panel umesto ekrana.
    '
    ' Do v6-ui-200 su stajali iza ekrana MAT_SISTEM i dugmeta "Otvori alatku" --
    ' jedan klik i jedan spisak vise, bez ijednog dobitka: spisak od dve stavke
    ' je ponavljao ono sto sidebar vec ume. Otud ekran alatki vise ne postoji.
    c.Add "MAT_PODESAVANJA||OTKUI_MS_PODESAVANJA|" & IC_MAT_SISTEM & _
          "|SISTEM|" & OBL_MATICNI & "|" & SEK_MATICNI
    c.Add "MAT_ADMIN||OTKUI_MS_ADMIN|" & IC_MAT_ADMIN & _
          "|SISTEM|" & OBL_MATICNI & "|" & SEK_MATICNI

    Dim a() As Variant, i As Long
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = c(i)
    Next i
    ScrRows = a
End Function

' Grupe sidebara, redom. Naslov grupe je kljuc kataloga, trece polje je SEKCIJA.
' Redosled ovde je i redosled crtanja: prvo sve grupe radne sekcije, pa maticne.
Public Function ScrGroups() As Variant
    ScrGroups = Array("OPERACIJE|OTKUI_NAVG_OPERACIJE|" & SEK_RAD, _
                      "FINANSIJE|OTKUI_NAVG_FINANSIJE|" & SEK_RAD, _
                      "ANALITIKA|OTKUI_NAVG_ANALITIKA|" & SEK_RAD, _
                      "SIFARNICI|OTKUI_NAVG_SIFARNICI|" & SEK_MATICNI, _
                      "SISTEM|OTKUI_NAVG_SISTEM|" & SEK_MATICNI)
End Function

' Sekcija reda registra ili reda grupe. Prazno polje znaci RAD -- red koji je
' ostao bez sedmog polja (star export, rucna izmena) ne sme da nestane iz
' sidebara nego se ponasa kao pre uvodjenja sekcija.
Public Function SekcijaIli(ByVal s As String) As String
    SekcijaIli = SEK_RAD
    If Len(Trim$(s)) > 0 Then SekcijaIli = Trim$(s)
End Function

Public Function ScrSekcija(ByVal row As String) As String
    ScrSekcija = SekcijaIli(ScrField(row, SCR_SEKCIJA))
End Function

Public Function ScrGrupaSekcija(ByVal grpRow As String) As String
    ScrGrupaSekcija = SekcijaIli(ScrField(grpRow, SCRG_SEKCIJA))
End Function

' Prvi ekran sekcije na koji korisnik SME i koji postoji. Prazno = nijedan
' (npr. maticni ekrani pre nego sto im modul stigne). Ljuska ovim ne saznaje
' nijedan kljuc unapred -- pita registar i dobija onaj koji je dostupan.
Public Function ScrPrviUSekciji(ByVal sekcija As String) As String
    Dim r As Variant, kljuc As String
    For Each r In ScrRows()
        If ScrSekcija(CStr(r)) = sekcija Then
            kljuc = ScrField(CStr(r), SCR_KLJUC)
            If ScrAktivan(kljuc) Then
                ScrPrviUSekciji = kljuc
                Exit Function
            End If
        End If
    Next r
End Function

Public Function ScrField(ByVal row As String, ByVal idx As Long) As String
    Dim p() As String
    p = Split(row, "|")
    If idx > UBound(p) Then Exit Function
    ScrField = p(idx)
End Function

Public Function ScrRowByKey(ByVal kljuc As String) As String
    Dim r As Variant
    For Each r In ScrRows()
        If ScrField(CStr(r), SCR_KLJUC) = kljuc Then
            ScrRowByKey = CStr(r)
            Exit Function
        End If
    Next r
End Function

'-------------------------------------------------------------- STANJE
' Postoji li modul ekrana i odgovara li na ugovor. Provera je namerno
' pokusaj poziva, ne pretraga po projektu: VBComponents trazi programski
' pristup VBA projektu, koji na korisnickoj masini cesto nije ukljucen.
Public Function ScrPostoji(ByVal kljuc As String) As Boolean
    Dim modul As String, dummy As Variant
    If mHas Is Nothing Then Set mHas = CreateObject("Scripting.Dictionary")
    If mHas.Exists(kljuc) Then
        ScrPostoji = mHas(kljuc)
        Exit Function
    End If
    ' Panel nema modul ekrana, ali POSTOJI -- inace bi ga sidebar crtao
    ' prigusenim kao ekran koji jos nije napisan.
    If modUiPanel.PanelPostoji(kljuc) Then
        ScrPostoji = True
        mHas(kljuc) = True
        Exit Function
    End If
    modul = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(modul) = 0 Then Exit Function
    On Error Resume Next
    dummy = Application.Run(modul & ".Scr_Meta")
    ScrPostoji = (ScrRunIshod(dummy, modul & ".Scr_Meta") <= SCR_RUN_PRAZNO)
    Err.Clear
    On Error GoTo 0
    mHas(kljuc) = ScrPostoji
End Function

' Ekran je dostupan ako postoji I ako korisnik ima pravo na njegovu oblast I
' ako sam ekran ne kaze da nije.
'
' TRECI uslov je uveden zbog maticnih sekcija Korisnici / Podesavanja / Admin:
' one traze ADMINISTRACIJU, a to nije oblast iz tblKorisnici pa ne moze da stane
' u SCR_OBLAST. Legacy je istu branu drzao u modMaticniLookups.MaticniMenu_OnClick;
' ovde je pita EKRAN, kroz neobavezan Scr_Dozvoljen -- ljuska i dalje ne zna
' nijedan ekran po imenu. Ekran koji ga ne implementira se ponasa kao pre.
'
' Brana je i dalje samo UI brana; tvrde su u modAdmin/modPodesavanja ulaznim
' tackama i tamo ostaju.
' Test seam: svi ekrani date sekcije su zabranjeni. Prazno gasi seam.
Public Sub ScrSekcijuZabraniTest(ByVal sekcija As String)
    If Not IsTestMode() Then Exit Sub
    mSekcijaZabranjenaTest = sekcija
End Sub

' Test seam: gradnja zone ekrana "pada". False gasi seam.
Public Sub ScrGradnjuOboriTest(ByVal obori As Boolean)
    If Not IsTestMode() Then Exit Sub
    mGradnjaPadaTest = obori
End Sub

' Test seam (KI-014): koji clan ugovora "curi" (v. mCuriTest). Prazno gasi seam.
Public Sub ScrCuriTestSet(ByVal clan As String)
    If Not IsTestMode() Then Exit Sub
    mCuriTest = clan
End Sub

' Kuka seam-a. Zove je telo clana ugovora NA SVOM IZLAZU, posle poslednjeg
' On Error/Err.Clear. Van test rezima i za drugi clan ne radi nista i ne dira
' Err. Vraca True kad je gresku ostavila.
Public Function ScrCuriAkoTest(ByVal clan As String) As Boolean
    Dim ws As Object
    If Len(mCuriTest) = 0 Then Exit Function
    If mCuriTest <> clan Then Exit Function
    If Not IsTestMode() Then Exit Function
    ScrCuriAkoTest = True
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("__KI014_CURI__")
End Function

' Kuka seam-a BEZ curenja: True kad je seam postavljen na ovaj clan, a clan tada
' vraca Empty uz Err = 0 -- oblik "ekran nema listu" (236). Danas ga nijedan
' ekran ne proizvodi u normalnom toku (sonda 14.10.2026: svih 14 vraca niz, i
' ANALIZA), pa bez ovoga grana SCR_RUN_PRAZNO u ScrGridData nema slucaj.
Public Function ScrPraznoAkoTest(ByVal clan As String) As Boolean
    If Len(mCuriTest) = 0 Then Exit Function
    If mCuriTest <> clan Then Exit Function
    ScrPraznoAkoTest = IsTestMode()
End Function

' Test seam (KI-016): koji clan ugovora "pada" (v. mPadTest). Prazno gasi seam.
Public Sub ScrPadTestSet(ByVal clan As String)
    If Not IsTestMode() Then Exit Sub
    mPadTest = clan
End Sub

' Kuka seam-a. Zove je telo clana ugovora na SVOM VRHU, odmah posle brane, pa
' podmetnut pad gadja branu koja pokriva celo telo. Van test rezima i za drugi
' clan ne radi nista i ne dira Err.
Public Sub ScrPadAkoTest(ByVal clan As String)
    If Len(mPadTest) = 0 Then Exit Sub
    If mPadTest <> clan Then Exit Sub
    If Not IsTestMode() Then Exit Sub
    Err.Raise 9, "modUiScreens.ScrPadAkoTest[" & clan & "]", _
              "podmetnut pad clana ugovora (KI-016)"
End Sub

Public Function ScrDozvoljen(ByVal kljuc As String) As Boolean
    Dim obl As String
    On Error Resume Next
    If Len(mSekcijaZabranjenaTest) > 0 And IsTestMode() Then
        If ScrSekcija(ScrRowByKey(kljuc)) = mSekcijaZabranjenaTest Then Exit Function
    End If
    obl = ScrField(ScrRowByKey(kljuc), SCR_OBLAST)
    If Len(obl) = 0 Then
        ScrDozvoljen = True
    Else
        ScrDozvoljen = modAuth.KorisnikImaPravo(obl)
    End If
    If Not ScrDozvoljen Then Exit Function
    ' Panel nema Scr_Dozvoljen jer nema modul ekrana -- njegovu branu drzi
    ' registar panela. Bez ovoga bi stavka u sidebaru stajala puna, a otvaranje
    ' bi odbilo: prigusenje mora da kaze istinu PRE klika.
    If modUiPanel.PanelPostoji(kljuc) Then
        ScrDozvoljen = modUiPanel.PanelDozvoljen(kljuc)
        Exit Function
    End If
    ScrDozvoljen = ScrSopstvenaBrana(kljuc)
End Function

' Odgovor samog ekrana na pitanje "smem li da te otvorim".
'
' KO IMA BRANU KAZE REGISTAR (SCR_BRANA), ne broj greske. Ekran koji je ne
' deklarise se ne pita uopste -- prolazi. Ekran koji je deklarise mora da
' odgovori: brana koja NIJE ODGOVORILA (Run nije vratio vrednost) je ZABRANJENO.
'
' Do v6-ui-204 se sposobnost pogadjala iz greske 1004 ("Cannot run the macro").
' Ali 1004 je obicna Excel greska koju i postojeca brana moze da digne iznutra,
' pa je pukla brana bila nerazluciva od nepostojece -- i propustala. Sada
' registar kaze da li brana postoji, a ako postoji, neodgovor je uvek "ne".
' Greska koja ostane u Err uz VRACEN odgovor je zaostatak, ne "ne" (KI-014).
'
' Odgovor se NIKAD ne kesira: prava se menjaju zamenom operatera, pa bi kesiran
' odgovor ostavio otvoren ekran kome novi operater nema pristup.
Private Function ScrSopstvenaBrana(ByVal kljuc As String) As Boolean
    Dim m As String, red As String, v As Variant
    ScrSopstvenaBrana = True
    red = ScrRowByKey(kljuc)
    If ScrField(red, SCR_BRANA) <> "1" Then Exit Function
    m = ScrField(red, SCR_MODUL)
    If Len(m) = 0 Then Exit Function

    On Error Resume Next
    Err.Clear
    v = Application.Run(m & ".Scr_Dozvoljen")
    If ScrRunIshod(v, m & ".Scr_Dozvoljen") = SCR_RUN_VRATIO Then
        ScrSopstvenaBrana = CBool(v)
    Else
        ScrSopstvenaBrana = False
        ScrLastErr = m & ".Scr_Dozvoljen -> " & Err.Number & " " & Err.description
    End If
    Err.Clear
End Function

' Ekran napusta scenu. Neobavezno -- ekran koji to ne implementira nema sta da
' sprema. Zove se PRE nego sto se predje na drugi: ekran tada zatvara editore i
' brise izbor, jer je sve to njegovo stanje.
'
' Greska poziva znaci "nema takvu proceduru", ne "pad" -- isto kao ScrPostoji.
' Ima li ekran nesacuvanih izmena. Neobavezno -- ekran koji to ne implementira
' nema sta da izgubi, pa clan koji se nije izvrsio znaci False, ne pad. Isti
' obrazac i isto ime kao modUiPanel.PanelImaNesacuvano; ljuska pita oba pre nego
' sto zatvori.
'
' Jedini FAIL-OPEN u ugovoru: dok se ovde citao Err.Number, zaostala greska uz
' odgovor True je davala False -- ljuska bi zatvorila ekran bez pitanja i
' odbacila unos (KI-014, test 238).
Public Function ScrImaNesacuvano(ByVal kljuc As String) As Boolean
    Dim m As String, v As Variant
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    On Error Resume Next
    Err.Clear
    v = Application.Run(m & ".Scr_ImaNesacuvano")
    If ScrRunIshod(v, m & ".Scr_ImaNesacuvano") = SCR_RUN_VRATIO Then _
        ScrImaNesacuvano = CBool(v)
    Err.Clear
End Function

' Ekran je USAO na scenu. Neobavezan clan ugovora, kao Scr_Dozvoljen -- ekran
' koji ga ne implementira ponasa se kao pre.
'
' Postoji zbog Banke: u legacy meniju je klik na Banka UVOZIO izvode pa otvarao
' mapiranje. Ljuska nema to dugme -- ulazak u ekran JESTE taj klik. Bez ove
' kuke ekran bi pokazivao samo ono sto je neko ranije uvezao.
'
' Zove se POSLE aktivacije (mScreen je vec postavljen), da radnja i njena poruka
' padnu na ekran koji je vec na sceni, a njen pad ne obori navigaciju.
Public Sub ScrAktiviraj(ByVal kljuc As String)
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Sub
    Application.Run m & ".Scr_Aktiviraj"
    Err.Clear
End Sub

Public Sub ScrDeaktiviraj(ByVal kljuc As String)
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Sub
    Application.Run m & ".Scr_Deaktiviraj"
    Err.Clear
End Sub

Public Function ScrAktivan(ByVal kljuc As String) As Boolean
    ScrAktivan = ScrPostoji(kljuc)
    If ScrAktivan Then ScrAktivan = ScrDozvoljen(kljuc)
End Function

' Kes se prazni pri gradnji ekrana - posle self-update-a moduli koji nisu
' postojali mogu da postoje.
Public Sub ScrResetCache()
    Set mHas = Nothing
End Sub

'--------------------------------------------------------------- POZIV
' Omotaci ugovora. Ishod svakog poziva tumaci ScrRunIshod (dole), ne Err.Number.
'
' Omotac NE MOZE da uhvati pad ekrana. Izmereno sondom (14. i 15.10.2026): clan
' koji gresku DIGNE je kroz Application.Run ne vraca pozivaocu -- ni pod On Error
' Resume Next, ni pod On Error GoTo, ni iz dogadjaja prikazane forme -- nego
' otvara runtime dijalog ("Run-time error ...", Continue UGASEN, End, Debug). "End"
' resetuje projekat: stanje modula se brise, forma nestaje, a Excel ostaje
' sakriven. Zato pad hvata CLAN, u svom handleru, i javlja ga kroz kanal
' (ScrPadZabelezi); citaoci sa porukom za pad -- ScrGridData, ScrEvent, ScrBuild
' -- kanal citaju pre vracene vrednosti (KI-016, UI-PAD-01).
Public Function ScrMeta(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrMeta = CStr(Application.Run(m & ".Scr_Meta"))
End Function

' Gradnja zone. Scr_Build je Sub, pa Run ne vraca nista: uspeh je Err = 0. Pad
' koji clan uhvati ostavlja Err = 0 kao i uspeh (sonda 15.10.2026), pa se kanal
' cita PRVI -- bez njega bi ljuska presla na ekran cija zona nije izgradjena.
Public Function ScrBuild(ByVal kljuc As String, ByVal z As Object) As Boolean
    Dim m As String
    On Error Resume Next
    If mGradnjaPadaTest And IsTestMode() Then Exit Function
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    ScrPadPreuzmi
    Application.Run m & ".Scr_Build", z
    If Len(ScrPadPreuzmi()) > 0 Then
        Err.Clear
        Exit Function
    End If
    ScrBuild = (Err.Number = 0)
End Function

Public Function ScrLayout(ByVal kljuc As String, ByVal z As Object, _
                          ByVal w As Single, ByVal h As Single) As Single
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrLayout = CSng(Application.Run(m & ".Scr_Layout", z, w, h))
End Function

Public Function ScrGrid(ByVal kljuc As String) As Variant
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrGrid = Application.Run(m & ".Scr_Grid")
End Function

' Radnja ekrana. Pad radnje je samo clan koji se NIJE izvrsio iz razloga koji
' nije "nema ga" (SCR_RUN_PAO). Ekran bez Scr_Event (1004) radnju nije ni
' imao -- to nije pad: ANALIZA ga nema, a gradnja zone sledeceg ekrana joj salje
' "chg:" dok je jos aktivna (izmereno 14.10.2026, test 237).
'
' Pad radnje koji clan uhvati (KI-016) javlja kanal, i to PRE vrednosti: Scr_Event
' je Boolean, pa posle uhvacenog pada Run vraca False -- "vratio", ne Empty.
Public Function ScrEvent(ByVal kljuc As String, ByVal tag As String, _
                         ByVal ev As String) As Boolean
    Dim m As String, v As Variant, ishodRadnje As Long, pad As String
    ScrLastErr = ""
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    ScrPadPreuzmi
    v = Application.Run(m & ".Scr_Event", tag, ev)
    pad = ScrPadPreuzmi()
    If Len(pad) > 0 Then
        ScrLastErr = pad
        Err.Clear
        Exit Function
    End If
    ishodRadnje = ScrRunIshod(v, m & ".Scr_Event " & tag)
    If ishodRadnje = SCR_RUN_VRATIO Then ScrEvent = CBool(v)
    If ishodRadnje = SCR_RUN_PAO Then ScrLastErr = m & ".Scr_Event " & tag & " -> " & Err.Number & " " & Err.description
    Err.Clear
End Function

' Redovi za deljenu mrezu. Vraca Array(kolone, redovi, n, zbirKg, zbirVal)
' ili Empty ako ekran nema listu.
' NE zove se ScrRows - to ime vec nosi spisak redova REGISTRA, a dve funkcije
' istog imena u istom modulu su "Ambiguous name".
'
' NEUSPEH SE BELEZI, ne samo gusi. Gusenje BEZ TRAGA je vec jednom skupo
' naplaceno ("Po datumu", v. ScrLastErr gore): na cipu "Svi" je lista koja se
' nije procitala davala Empty, LoadGridFromScreen na ne-niz radi Exit Sub -- pa
' je mreza OSTAJALA na prethodnoj listi, bez ijedne poruke i bez traga u logu.
'
' Tri oblika (KI-014, test 236), jer Scr_Rows LEGITIMNO vraca Empty:
'   - lista uz gresku u Err      -> lista vazi, greska je zaostatak;
'   - Empty uz Err = 0           -> ekran nema listu, nije neuspeh;
'   - Empty uz gresku            -> neuspeh. Ovde se nepokrenut clan (1004) i
'     Empty uz zaostatak NE MOGU razlikovati (sonda 14.10.2026), pa vazi
'     strozi ishod: mreza javlja da se lista nije procitala.
' Cetvrti oblik je PAD koji je clan uhvatio (KI-016): Empty uz Err = 0, isto kao
' "nema liste", a clan koji je vec dodelio listu pa pao vraca i nju. Zato kanal
' ide pre svih: pad je neuspeh liste, sa imenom clana i korakom u ScrLastErr.
Public Function ScrGridData(ByVal kljuc As String, ByVal filter As String, _
                            ByVal q As String) As Variant
    Dim m As String, v As Variant, pad As String
    ScrLastErr = ""
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    ScrPadPreuzmi
    v = Application.Run(m & ".Scr_Rows", filter, q)
    pad = ScrPadPreuzmi()
    If Len(pad) > 0 Then
        ScrLastErr = pad
        Err.Clear
        Exit Function
    End If
    Select Case ScrRunIshod(v, m & ".Scr_Rows")
        Case SCR_RUN_VRATIO: ScrGridData = v
        Case SCR_RUN_PRAZNO
        Case Else: ScrLastErr = m & ".Scr_Rows -> " & Err.Number & " " & Err.description
    End Select
    Err.Clear
End Function

' Cipovi AKTIVNE liste, oblika 'kljuc:KATALOG:sirina|...'. Opciono -- ekran
' koji ih nema vraca prazno i mreza se ne suzava nicim osim pretragom.
' Kljuc se vraca ljusci samo kao mFilter i putuje nazad u Scr_Rows; ljuska ga
' ne tumaci.
Public Function ScrCipovi(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    ScrCipovi = CStr(Application.Run(m & ".Scr_Cipovi"))
    Err.Clear
End Function

' Prekidac lista ekrana. Prazno = ekran ima samo jednu listu, pa prekidaca
' nema. Ljuska ne zna nijedan kljuc unapred - ni "OTPREMNICE" ni "PRERADE".
Public Function ScrListe(ByVal kljuc As String) As Variant
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrListe = Application.Run(m & ".Scr_Liste")
End Function

Public Function ScrLista(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrLista = CStr(Application.Run(m & ".Scr_Lista"))
End Function

' Podrazumevani sort AKTIVNE liste ekrana, oblika "kolona:asc" ili "kolona:desc".
' Prazno = ekran nema misljenje, pa ostaje pravilo ljuske (SortZaListu).
'
' Uvedeno zbog sifarnika: ljuskino pravilo je "druga kolona opadajuce", sto je
' datum dokumenta -- nad Stanicama to znaci nazive unazad. Ekran koji zna sta mu
' je glavna kolona to kaze sam, umesto da se ljusci dopisuje jos trinaest imena
' lista pored zatecena dva.
Public Function ScrSort(ByVal kljuc As String) As String
    Dim m As String, v As Variant
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    v = Application.Run(m & ".Scr_Sort")
    If ScrRunIshod(v, m & ".Scr_Sort") = SCR_RUN_VRATIO Then ScrSort = CStr(v)
    Err.Clear
End Function

' Radnje nad izabranim redom za trenutno aktivnu listu ekrana.
' Koliko stavki na ovom ekranu CEKA operatera. Opciono: ekran koji nema sta da
' broji je ne implementira i dobija nulu.
'
' Ljuska ovim ne saznaje NISTA o ekranu -- ne zna sta se broji ni zasto, samo
' dobija broj koji ce nacrtati uz stavku menija. Bez ovoga bi morala da zove
' GetNedovrseno po imenu, a to je tacno ono sto ceo ugovor izbegava.
Public Function ScrBrojac(ByVal kljuc As String) As Long
    Dim m As String, v As Variant
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    v = Application.Run(m & ".Scr_Brojac")
    If ScrRunIshod(v, m & ".Scr_Brojac") = SCR_RUN_VRATIO Then ScrBrojac = CLng(v)
    Err.Clear
End Function

' Broji li podnozje KOMADE umesto dinara na ovom ekranu.
'
' Ljuska je do sada pitala `ModeBrojiKomade(ActiveMode)` -- a ActiveMode je rezim
' UNOSA DOKUMENATA i na ugovornom ekranu ostaje onakav kakav ga je Dokumenta
' ostavila. Ko je bio na F7 (Reversi) pa presao na Uvoz izvoda, u podnozju je
' video "Ukupno 8.950 kom" umesto "Vrednost 8.950,00 RSD": novac izbrojan kao
' komadi, bez decimala. Ista klasa kao traka `zOtp` koja je ostajala upaljena.
'
' Ekran koji ovo ne implementira dobija False -- dinari. Ljuska ovim ne saznaje
' NISTA o ekranu, isto kao kod ScrBrojac.
Public Function ScrBrojiKomade(ByVal kljuc As String) As Boolean
    Dim m As String, v As Variant
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) = 0 Then Exit Function
    Err.Clear
    v = Application.Run(m & ".Scr_BrojiKomade")
    If ScrRunIshod(v, m & ".Scr_BrojiKomade") = SCR_RUN_VRATIO Then ScrBrojiKomade = CBool(v)
    Err.Clear
End Function

Public Function ScrRadnje(ByVal kljuc As String) As String
    Dim m As String
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    If Len(m) > 0 Then ScrRadnje = CStr(Application.Run(m & ".Scr_Radnje"))
End Function

' Upis dokumenta. Ljuska predaje RECNIK vrednosti pod logickim imenima; ekran
' zna sta su i sta se sa njima radi. Vraca "" kad je proslo, inace poruku za
' operatera (jedan razmak = operater je odustao, ne prikazuje se nista).
' Ekran u isti recnik upisuje "fokus", "rezultat" i "poruke".
'
' NEUSPEH JE SAMO NEVRACEN REZULTAT (KI-011, UI-ISHOD-01). Izmereno sondom
' 13.10.2026:
'   - makro koji DIGNE gresku kroz Application.Run je NE vraca pozivaocu, nego
'     otvara runtime dijalog -- pa "greska u Err posle Run-a" nikad nije bila
'     pad Scr_Save;
'   - Run koji makro ne moze da pokrene (1004) ne vraca nista -- to JESTE
'     neuspeh, upis se nije ni desio;
'   - makro koji VRATI rezultat a ostavi gresku u Err vraca i jedno i drugo.
' Treci slucaj je do KI-011 citan kao neuspeh: revers je bio snimljen, a ljuska je
' javila "Radnja nije uspela", nije ocistila formu ni osvezila listu, pa je
' operater isti revers snimio dvaput. Sada vazi ono sto je ekran VRATIO, a
' zaostala greska ide u ScrZaostalaGreska i u log -- vidljiva, ne progutana.
Public Function ScrSave(ByVal kljuc As String, ByVal polja As Object) As String
    Dim m As String, v As Variant, ishodUpisa As Long
    ScrLastErr = ""
    ScrZaostalaGreska = ""
    On Error Resume Next
    m = ScrField(ScrRowByKey(kljuc), SCR_MODUL)
    ' Ekran bez modula u registru: nijedan pisac nije pozvan, pa to JE neuspeh
    ' (review #419, P2). Prazan rezultat uz prazan ScrLastErr CommitDokument cita
    ' kao upis -- ocistio bi formu i javio "sacuvan" bez dokumenta.
    If Len(m) = 0 Then
        ScrLastErr = "modUiScreens.ScrSave: ekran '" & kljuc & "' nema modul u registru"
        Err.Clear
        Exit Function
    End If
    Err.Clear
    v = Application.Run(m & ".Scr_Save", polja)
    ishodUpisa = ScrRunIshod(v, m & ".Scr_Save")
    If ishodUpisa = SCR_RUN_VRATIO Then
        ScrSave = CStr(v)
    Else
        ScrLastErr = m & ".Scr_Save -> " & Err.Number & " " & Err.description
    End If
    Err.Clear
End Function

' Pad koji je clan zabelezio od poslednjeg preuzimanja; vraca ga i BRISE.
' Citalac ga zove DVAPUT: pre Run-a (odbacuje pad koji nije iz ovog poziva --
' clan pozvan direktno, mimo citaoca, ostavlja ga u kanalu) i posle Run-a. Tako
' i ugnjezden citalac (radnja koja osvezi mrezu) preuzme SVOJ pad, pa spoljni
' ne dobije tudji. Bez On Error naredbe, iz istog razloga kao ScrRunIshod.
Public Function ScrPadPreuzmi() As String
    ScrPadPreuzmi = mPadClana
    mPadClana = ""
End Function

' Zove ga CLAN UGOVORA iz svog handlera (UI-PAD-01): upisuje pad u kanal i ERROR
' u log. Jedna implementacija za sve clanove -- clan nosi samo branu i ime.
' Err se cita PRVI, pre bilo kog poziva. Izvor greske ide u zapis kad nosi ime
' (funkcija liste ga dize sa korakom, npr. "modScrFakture.RedoviFakture[fakture]"),
' jer je to ono sto ReloadGrid prikazuje posle "na koraku:". Err se ovde ne brise:
' izlaz iz handlera clana ga brise sam (sonda 15.10.2026).
'
' Oblik je "clan [izvor] -> broj opis": izvor ide PRE opisa. Opis nema granicu
' duzine (greska seme nabraja celo zaglavlje), a traka poruke odseca KRAJ -- pa
' je korak, dok je stajao iza opisa, ispadao sa ekrana (rucna provera #423).
Public Sub ScrPadZabelezi(ByVal clan As String)
    Dim errNum As Long, errSrc As String, errDesc As String
    errNum = Err.Number: errSrc = Err.Source: errDesc = Err.description
    mPadClana = clan
    If InStr(1, errSrc, ".", vbBinaryCompare) > 0 Then _
        mPadClana = mPadClana & " [" & errSrc & "]"
    mPadClana = mPadClana & " -> " & errNum & " " & errDesc
    LogError clan, "pad clana ugovora (KI-016): " & errDesc, errNum, LOG_ERROR, errSrc
End Sub

' ISHOD Application.Run NAD CLANOM UGOVORA -- UI-ISHOD-01 za ceo ugovor ekrana
' (KI-011 za Scr_Save, KI-014 za ostale). Jedno pravilo za sve citaoce: do
' KI-014 je svaki citao Err.Number po svom, i svaki ga je tumacio drugacije.
'
' Zove se ODMAH posle Run-a, sa onim sto je Run vratio, pre bilo cega sto bi
' moglo da obrise Err. Izmereno sondom (13. i 14.10.2026):
'   - clan koji VRATI vrednost a ostavi gresku u Err vraca i jedno i drugo;
'     Boolean/Long/String/niz su tada UVEK ne-Empty -> SCR_RUN_VRATIO, a greska
'     je ZAOSTATAK pomocnog koraka: ScrZaostalaGreska + WARN, Err se brise;
'   - clan koji se ne izvrsi ne vraca nista: 1004 (nema ga) -> SCR_RUN_NEMA,
'     druga greska (449, pogresan potpis) -> SCR_RUN_PAO. Err OSTAJE pozivaocu
'     da ga upise u ScrLastErr;
'   - Empty uz Err = 0 je clan koji je vratio Empty -> SCR_RUN_PRAZNO.
' Variant clan koji vrati Empty UZ zaostatak, i Sub uz zaostatak, nerazlucivi su
' od nepokrenutog clana (isti oblik). Zato ScrBuild (Sub) i dalje cita Err.
' Clan koji gresku DIGNE ovde ne stize (v. POZIV gore); pad koji clan uhvati
' nosi kanal (ScrPadPreuzmi), a citalac ga gleda PRE ovog pomocnika.
'
' Bez On Error naredbe -- ona bi obrisala Err koji pozivalac jos treba da
' procita. v je ByRef: Scr_Rows vraca celu mrezu, ByVal bi je kopirao.
Public Function ScrRunIshod(ByRef v As Variant, ByVal poziv As String) As Long
    If IsEmpty(v) Then
        If Err.Number = 0 Then
            ScrRunIshod = SCR_RUN_PRAZNO
        ElseIf Err.Number = 1004 Then
            ScrRunIshod = SCR_RUN_NEMA
        Else
            ScrRunIshod = SCR_RUN_PAO
        End If
        Exit Function
    End If
    ScrRunIshod = SCR_RUN_VRATIO
    If Err.Number = 0 Then Exit Function
    ScrZaostalaGreska = poziv & ": " & Err.Number & " " & Err.description
    Err.Clear
    LogWarn "modUiScreens.ScrRunIshod", "zaostala greska posle " & poziv & _
            " (KI-014): " & ScrZaostalaGreska
End Function
