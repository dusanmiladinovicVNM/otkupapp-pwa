Attribute VB_Name = "modMaticniIzvor"
'=====================================================================
' modMaticniIzvor - JEDAN opis 13 maticnih sekcija (korak M1).
'
' Sve sto tri maticna ekrana (Partneri, Proizvodi i cene, Ambalaza i pakovanje)
' znaju o podacima stoji ovde: koja tabela, koji je PK, koje kolone se vide i
' odakle im vrednost. Ekrani su tanki -- oni samo kazu KOJE sekcije nose.
'
' ODAKLE DOLAZI: frmStammdaten. Trinaest Setup* procedura je davalo naslov i
' polja, a LoadList (351 linija) je punio listbox -- sest sekcija po imenu
' kolone i sa izvedenim vrednostima, ostalih sedam kroz zajednicki Case Else
' (pozicijski: kolona 1 je ID, pa redom). Ovde je to jedan opis, pa se dodavanje
' sekcije svede na jedan red umesto na dve grane u dve procedure.
'
' STA JE PRENETO 1:1: koje kolone operater vidi i u kom redosledu, i sve
' izvedene vrednosti (puno ime, naziv stanice umesto ID-ja, spojena adresa,
' kooperant sa ID-jem uz ime, geo status/izvor). Cenovnik zadrzava svoj poredak
' (ExcludeStornirano -> SortArray po datumu opadajuce, tie-break CenaID).
'
' STA JE NAMERNO DRUGACIJE:
'   - kolona statusa se TRAZI u semi (Aktivan / Aktivna), ne pogadja. Sema je
'     izvor istine, ne kod -- isti probe koji legacy radi u AktivanColName.
'     Sekcija bez te kolone nema ni cipove; nema tihog "prikazujem sve".
'   - identitet reda je PK u koloni 1, uvek. Legacy je red birao po poziciji u
'     listboxu (m_RowMap); posle sortiranja i pretrage pozicija ne znaci nista,
'     a istoimeni zapisi su u sifarnicima obicna pojava.
'   - status ostaje TEKST, ne pilula. Vrsta celije "pill" ima zatvoren recnik
'     od tri stanja dokumenta (Otkazana / Poslato / Sacuvana), pa bi aktivnom
'     kooperantu pisalo "Sacuvana". Pilula za Aktivan/Neaktivan bi bila nova
'     vrsta celije -- to nije M1.
'
' OVAJ MODUL NE UPISUJE NISTA. Upis dolazi u M2 (modMaticniUnos), i tada se
' izdvaja iz frmStammdaten pa ga zovu OBE strane -- v. UI_MIGRACIJA_KATALOG 24.5.
'
' Fajl mora ostati 100% ASCII.
'=====================================================================
Option Explicit

Public Const MATIZ_BUILD As String = "v6-ui-188"

' Cipovi su svuda isti: jedina poslovna osa koju sifarnik ima je soft-delete.
Public Const MAT_CIP_SVI As String = "sve"
Public Const MAT_CIP_AKT As String = "aktivni"
Public Const MAT_CIP_NEAKT As String = "neaktivni"

' Brojke za zonu ekrana; racunaju se u istom prolazu kroz podatke kao i redovi,
' pa se prikaz i broj ne mogu razici.
Private mUkupno As Long
Private mAktivnih As Long
Private mNeaktivnih As Long

' Kes odgovora "kako se u OVOJ tabeli zove kolona statusa". Odgovor trazi prolaz
' kroz zaglavlje tabele, a pita se pri svakom crtanju mreze i svakog cipa.
Private mStatusKol As Object

'------------------------------------------------------------- SEKCIJE
' Sekcije po ekranu. Redosled = redosled u prekidacu lista.
' Red: "KLJUC|natpis|naslov mreze|sirina" -- ugovor Scr_Liste.
Public Function MatSekcijeEkrana(ByVal ekran As String) As Variant
    Select Case ekran
        Case "MAT_PARTNERI"
            MatSekcijeEkrana = Array( _
                "KOOPERANTI|OTKUI_SEGM_KOOP|OTKUI_GTM_KOOP|92", _
                "STANICE|OTKUI_SEGM_STAN|OTKUI_GTM_STAN|76", _
                "KUPCI|OTKUI_SEGM_KUPCI|OTKUI_GTM_KUPCI|64", _
                "VOZACI|OTKUI_SEGM_VOZ|OTKUI_GTM_VOZ|68", _
                "PARCELE|OTKUI_SEGM_PARC|OTKUI_GTM_PARC|76")
        Case "MAT_ROBA"
            MatSekcijeEkrana = Array( _
                "ARTIKLI|OTKUI_SEGM_ART|OTKUI_GTM_ART|68", _
                "KULTURE|OTKUI_SEGM_KULT|OTKUI_GTM_KULT|72", _
                "CENOVNIK|OTKUI_SEGM_CEN|OTKUI_GTM_CEN|80", _
                "VRSTAGP|OTKUI_SEGM_VGP|OTKUI_GTM_VGP|120")
        Case "MAT_PAKOVANJE"
            MatSekcijeEkrana = Array( _
                "AMBALAZA|OTKUI_SEGM_AMB|OTKUI_GTM_AMB|84", _
                "PALETE|OTKUI_SEGM_PAL|OTKUI_GTM_PAL|68", _
                "KUTIJE|OTKUI_SEGM_KUT|OTKUI_GTM_KUT|68", _
                "KESE|OTKUI_SEGM_KES|OTKUI_GTM_KES|60")
    End Select
End Function

' Prva sekcija ekrana - podrazumevana lista. Cita se iz istog spiska, da se ne
' moze razici sa prekidacem.
Public Function MatPrvaSekcija(ByVal ekran As String) As String
    Dim a As Variant
    a = MatSekcijeEkrana(ekran)
    If Not IsArray(a) Then Exit Function
    MatPrvaSekcija = Split(CStr(a(LBound(a))), "|")(0)
End Function

' Prevod legacy Tag-a (frmStammdaten.Tag) u kljuc sekcije. Jedno mesto na kom
' se dva imenovanja sretnu -- forma i ekran posle toga govore istim jezikom.
' Korisnici NISU u spisku: oni idu u M4, a do tada forma zadrzava svoju granu.
Public Function MatKljucIzLegacyTag(ByVal tg As String) As String
    Select Case tg
        Case "Kooperanti":  MatKljucIzLegacyTag = "KOOPERANTI"
        Case "Stanice":     MatKljucIzLegacyTag = "STANICE"
        Case "Kupci":       MatKljucIzLegacyTag = "KUPCI"
        Case "Vozaci":      MatKljucIzLegacyTag = "VOZACI"
        Case "Parcele":     MatKljucIzLegacyTag = "PARCELE"
        Case "Artikli":     MatKljucIzLegacyTag = "ARTIKLI"
        Case "Kulture":     MatKljucIzLegacyTag = "KULTURE"
        Case "Cenovnik":    MatKljucIzLegacyTag = "CENOVNIK"
        Case "VrstaGP":     MatKljucIzLegacyTag = "VRSTAGP"
        Case "TipAmbalaze": MatKljucIzLegacyTag = "AMBALAZA"
        Case "TipPalete":   MatKljucIzLegacyTag = "PALETE"
        Case "Kutije":      MatKljucIzLegacyTag = "KUTIJE"
        Case "Kese":        MatKljucIzLegacyTag = "KESE"
    End Select
End Function

Public Function MatTabela(ByVal kljuc As String) As String
    Select Case kljuc
        Case "KOOPERANTI": MatTabela = TBL_KOOPERANTI
        Case "STANICE":    MatTabela = TBL_STANICE
        Case "KUPCI":      MatTabela = TBL_KUPCI
        Case "VOZACI":     MatTabela = TBL_VOZACI
        Case "PARCELE":    MatTabela = TBL_PARCELE
        Case "ARTIKLI":    MatTabela = TBL_ARTIKLI
        Case "KULTURE":    MatTabela = TBL_KULTURE
        Case "CENOVNIK":   MatTabela = TBL_CENOVNIK
        Case "VRSTAGP":    MatTabela = TBL_VRSTA_GP
        Case "AMBALAZA":   MatTabela = TBL_TIP_AMBALAZE
        Case "PALETE":     MatTabela = TBL_TIP_PALETE
        Case "KUTIJE":     MatTabela = TBL_KUTIJE
        Case "KESE":       MatTabela = TBL_KESE
    End Select
End Function

' Kolona koja NOSI IDENTITET reda. Za cetiri sifarnika pakovanja i za VrstaGP
' to je sam tip -- te tabele nemaju surogat kljuc, i tako ih legacy i pise.
Public Function MatPK(ByVal kljuc As String) As String
    Select Case kljuc
        Case "KOOPERANTI": MatPK = "KooperantID"
        Case "STANICE":    MatPK = "StanicaID"
        Case "KUPCI":      MatPK = "KupacID"
        Case "VOZACI":     MatPK = "VozacID"
        Case "PARCELE":    MatPK = COL_PAR_ID
        Case "ARTIKLI":    MatPK = "ArtikalID"
        Case "KULTURE":    MatPK = "KulturaID"
        Case "CENOVNIK":   MatPK = COL_CEN_ID
        Case "VRSTAGP":    MatPK = COL_VGP_TIP
        Case "AMBALAZA":   MatPK = COL_TAMB_TIP
        Case "PALETE":     MatPK = COL_TPAL_TIP
        Case "KUTIJE":     MatPK = COL_KUT_TIP
        Case "KESE":       MatPK = COL_KES_TIP
    End Select
End Function

' Naziv kolone statusa u OVOJ tabeli, ili "" ako je nema.
'
' TRAZI SE U SEMI, ne pogadja se: sema je izvor istine, a ne kod (CLAUDE.md 3).
' Parcele nose "Aktivna", ostali "Aktivan"; Cenovnik, Ambalaza i Palete je po
' zatecenoj semi nemaju uopste. Isti probe koji legacy radi u AktivanColName --
' i isti razlog: sekcija bez te kolone ne sme da dobije cip koji tiho ne
' filtrira nista.
Public Function MatStatusKolona(ByVal kljuc As String) As String
    Dim tbl As String
    If mStatusKol Is Nothing Then Set mStatusKol = CreateObject("Scripting.Dictionary")
    If mStatusKol.Exists(kljuc) Then
        MatStatusKolona = CStr(mStatusKol(kljuc))
        Exit Function
    End If
    tbl = MatTabela(kljuc)
    If Len(tbl) = 0 Then Exit Function
    On Error Resume Next
    If GetColumnIndex(tbl, "Aktivan") > 0 Then
        MatStatusKolona = "Aktivan"
    ElseIf GetColumnIndex(tbl, "Aktivna") > 0 Then
        MatStatusKolona = "Aktivna"
    End If
    Err.Clear
    mStatusKol(kljuc) = MatStatusKolona
End Function

' Cipovi sekcije. Sekcija bez kolone statusa nema sta da filtrira, pa ne
' prijavljuje nijedan cip -- mreza se tada suzava samo pretragom.
Public Function MatCipovi(ByVal kljuc As String) As String
    If Len(MatStatusKolona(kljuc)) = 0 Then Exit Function
    MatCipovi = MAT_CIP_SVI & ":OTKUI_CHIP_SVE:40|" & _
                MAT_CIP_AKT & ":OTKUI_CIPM_AKTIVNI:70|" & _
                MAT_CIP_NEAKT & ":OTKUI_CIPM_NEAKTIVNI:84"
End Function

'-------------------------------------------------------------- KOLONE
' Opis kolona je ISTI ugovor koji mreza vec koristi:
'   "kljuc naslova | izvor | vrsta | sirina | prioritet"
' Izvor je naziv kolone u tabeli ili izvedena vrednost ("@ime"). Prioritet 1 se
' crta uvek, 3 se sklanja na uskom ekranu, 4 nikad (identitet).
'
' NIJEDNA kolona nije vrste "kg" ni novcane ("rsd"/"mult"/"sum0"/"rest") -- v.
' UI_MIGRACIJA_KATALOG 24.4: podnozje pita opis kolona, pa sifarnik tako sam
' po sebi nema lazne zbirove. Tezine su "num" jer su broj, ali nisu zbirna
' velicina.
Public Function MatKolone(ByVal kljuc As String) As Variant
    Select Case kljuc
        Case "KOOPERANTI"
            MatKolone = Array( _
                "OTKUI_HDM_ID|KooperantID|txt|84|1", _
                "OTKUI_HDM_IME_PREZ|@ime_prezime|part|0|1", _
                "OTKUI_HDM_TELEFON|Telefon|txt|84|2", _
                "OTKUI_HDM_STANICA|@stanica|txt|96|2", _
                "OTKUI_HDM_BPG|BPGBroj|txt|72|3", _
                "OTKUI_HDM_RACUN|TekuciRacun|txt|140|3", _
                "OTKUI_HDM_ADRESA|@adresa_mesto|txt|150|3", _
                "OTKUI_HDM_JMBG|JMBG|txt|110|3", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "STANICE"
            MatKolone = Array( _
                "OTKUI_HDM_ID|StanicaID|txt|76|1", _
                "OTKUI_HDM_NAZIV|Naziv|part|0|1", _
                "OTKUI_HDM_MESTO|Mesto|txt|110|2", _
                "OTKUI_HDM_TELEFON|Telefon|txt|92|2", _
                "OTKUI_HDM_KONTAKT|@ime_prezime|txt|130|3", _
                "OTKUI_HDM_HLADNJACA|JeHladnjaca|txt|76|2", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "KUPCI"
            MatKolone = Array( _
                "OTKUI_HDM_ID|KupacID|txt|76|1", _
                "OTKUI_HDM_NAZIV|Naziv|part|0|1", _
                "OTKUI_HDM_ADRESA|@kupac_adresa|txt|170|2", _
                "OTKUI_HDM_DRZAVA|@drzava|txt|84|3", _
                "OTKUI_HDM_PIB|PIB|txt|76|2", _
                "OTKUI_HDM_MB|MaticniBroj|txt|84|3", _
                "OTKUI_HDM_EMAIL|Email|txt|150|3", _
                "OTKUI_HDM_RACUN|TekuciRacun|txt|140|3", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "VOZACI"
            MatKolone = Array( _
                "OTKUI_HDM_ID|VozacID|txt|76|1", _
                "OTKUI_HDM_IME|Ime|txt|130|1", _
                "OTKUI_HDM_PREZIME|Prezime|part|0|1", _
                "OTKUI_HDM_TELEFON|Telefon|txt|100|2", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "PARCELE"
            MatKolone = Array( _
                "OTKUI_HDM_ID|" & COL_PAR_ID & "|txt|84|1", _
                "OTKUI_HDM_KOOPERANT|@koop_naziv|part|0|1", _
                "OTKUI_HDM_KATBROJ|" & COL_PAR_KAT_BROJ & "|txt|96|1", _
                "OTKUI_HDM_KATOPSTINA|" & COL_PAR_KAT_OPSTINA & "|txt|120|2", _
                "OTKUI_HDM_KULTURA|" & COL_PAR_KULTURA & "|txt|96|2", _
                "OTKUI_HDM_POVRSINA|" & COL_PAR_POVRSINA & "|txt|72|1", _
                "OTKUI_HDM_GGAP|" & COL_PAR_GGAP & "|txt|84|3", _
                "OTKUI_HDM_GEO|@geo|txt|120|3", _
                "OTKUI_HDM_RIZIK|" & COL_PAR_RIZIK & "|txt|84|3", _
                "OTKUI_HDM_NAPOMENA|" & COL_PAR_NAPOMENA & "|txt|140|3", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "ARTIKLI"
            MatKolone = Array( _
                "OTKUI_HDM_ID|ArtikalID|txt|84|1", _
                "OTKUI_HDM_NAZIV|Naziv|part|0|1", _
                "OTKUI_HDM_TIP|Tip|txt|110|2", _
                "OTKUI_HDM_JM|JedinicaMere|txt|56|2", _
                "OTKUI_HDM_CENA_JED|CenaPoJedinici|num|84|1", _
                "OTKUI_HDM_DOZA|DozaPoHa|txt|76|3", _
                "OTKUI_HDM_KULTURA|Kultura|txt|100|3", _
                "OTKUI_HDM_PAKOVANJE|Pakovanje|txt|96|3")
        Case "KULTURE"
            MatKolone = Array( _
                "OTKUI_HDM_ID|KulturaID|txt|84|1", _
                "OTKUI_HDM_VRSTA|VrstaVoca|txt|110|1", _
                "OTKUI_HDM_SORTA|SortaVoca|part|0|1", _
                "OTKUI_HDM_GAJBICA_PAL|" & COL_KUL_GAJBICA_PALETA & "|num|96|2", _
                "OTKUI_HDM_TIP_AMB|" & COL_KUL_TIP_AMBALAZE & "|txt|110|2", _
                "OTKUI_HDM_PRAG_UPOZ|" & COL_KUL_PRAG_PROSEK_UPOZ & "|txt|92|3", _
                "OTKUI_HDM_PRAG_BLOK|" & COL_KUL_PRAG_PROSEK_BLOK & "|txt|92|3", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "CENOVNIK"
            MatKolone = Array( _
                "OTKUI_HDM_ID|" & COL_CEN_ID & "|txt|84|1", _
                "OTKUI_HDM_DATUM|" & COL_CEN_DATUM & "|date|72|1", _
                "OTKUI_HDM_VRSTA|" & COL_CEN_VRSTA & "|txt|110|1", _
                "OTKUI_HDM_SORTA|" & COL_CEN_SORTA & "|part|0|1", _
                "OTKUI_HDM_KLASA|" & COL_CEN_KLASA & "|txt|72|1", _
                "OTKUI_HDM_CENA|" & COL_CEN_CENA & "|num|84|1")
        Case "VRSTAGP"
            MatKolone = Array( _
                "OTKUI_HDM_TIP_GP|" & COL_VGP_TIP & "|part|0|1", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "AMBALAZA"
            MatKolone = Array( _
                "OTKUI_HDM_TIP_AMB|" & COL_TAMB_TIP & "|part|0|1", _
                "OTKUI_HDM_TEZINA_GAJ|" & COL_TAMB_TEZINA & "|num|130|1")
        Case "PALETE"
            MatKolone = Array( _
                "OTKUI_HDM_TIP_PAL|" & COL_TPAL_TIP & "|part|0|1", _
                "OTKUI_HDM_TEZINA|" & COL_TPAL_TEZINA & "|num|110|1")
        Case "KUTIJE"
            MatKolone = Array( _
                "OTKUI_HDM_TIP_KUT|" & COL_KUT_TIP & "|part|0|1", _
                "OTKUI_HDM_TEZINA|" & COL_KUT_TEZINA & "|num|110|1", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
        Case "KESE"
            MatKolone = Array( _
                "OTKUI_HDM_TIP_KES|" & COL_KES_TIP & "|part|0|1", _
                "OTKUI_HDM_TEZINA|" & COL_KES_TEZINA & "|num|110|1", _
                "OTKUI_HDM_STATUS|@status|txt|76|1")
    End Select
End Function

''-------------------------------------------------------------- POLJA
' Opis POLJA UNOSA po sekciji. Jedan spisak za dve stvari: sta pisac prima
' (modMaticniUnos) i sta editor crta (M2b). Dva spiska bi se razisla.
'
' Red: "kljuc | natpis(katalog) | vrsta | obavezno | kolona | combo"
'   vrsta    txt | num | cmb | date
'   obavezno 1 = prazno polje odbija upis, 0 = sme prazno
'   kolona   naziv kolone u tabeli; "@alias:A,B" = prva koja POSTOJI u semi
'            (drift po instalaciji -- isti probe koji legacy radi u
'            UpdateFirstExistingCol); "" = polje se ne upisuje direktno
'   combo    izvor stavki za padajucu listu (v. MatComboIzvor)
'
' Redosled je redosled u legacy formi -- operater ne uci nov raspored.
Public Function MatPolja(ByVal kljuc As String) As Variant
    Select Case kljuc
        Case "KOOPERANTI"
            MatPolja = Array( _
                "ime|OTKUI_HDM_IME|txt|1|Ime|", _
                "prezime|OTKUI_HDM_PREZIME|txt|1|Prezime|", _
                "mesto|OTKUI_HDM_MESTO|txt|0|Mesto|", _
                "telefon|OTKUI_HDM_TELEFON|txt|0|Telefon|", _
                "stanica|OTKUI_HDM_STANICA|cmb|1|StanicaID|@stanice", _
                "bpg|OTKUI_HDM_BPG|txt|0|BPGBroj|", _
                "racun|OTKUI_HDM_RACUN|txt|0|TekuciRacun|", _
                "pin|OTKUI_MP_PIN|txt|0|Pin|", _
                "adresa|OTKUI_HDM_ADRESA|txt|0|Adresa|", _
                "jmbg|OTKUI_HDM_JMBG|txt|0|JMBG|")
        Case "STANICE"
            MatPolja = Array( _
                "naziv|OTKUI_HDM_NAZIV|txt|1|Naziv|", _
                "mesto|OTKUI_HDM_MESTO|txt|1|Mesto|", _
                "telefon|OTKUI_HDM_TELEFON|txt|0|@alias:Kontakt,Telefon|", _
                "kime|OTKUI_MP_KONTAKT_IME|txt|0|@alias:Ime,KontaktIme|", _
                "kprezime|OTKUI_MP_KONTAKT_PREZ|txt|0|@alias:Prezime,KontaktPrezime|", _
                "pin|OTKUI_MP_PIN|txt|0|@alias:PIN,Pin|", _
                "hladnjaca|OTKUI_HDM_HLADNJACA|cmb|0|" & COL_STA_JE_HLADNJACA & "|@dane")
        Case "KUPCI"
            MatPolja = Array( _
                "naziv|OTKUI_HDM_NAZIV|txt|1|Naziv|", _
                "ulica|OTKUI_MP_ULICA|txt|0|Ulica|", _
                "mesto|OTKUI_HDM_MESTO|txt|0|Mesto|", _
                "posta|OTKUI_MP_POSTA|txt|0|PostanskiBroj|", _
                "drzava|OTKUI_HDM_DRZAVA|txt|0|@alias:Dr" & ChrW(382) & "ava,Drzava|", _
                "pib|OTKUI_HDM_PIB|txt|0|PIB|", _
                "mb|OTKUI_HDM_MB|txt|0|MaticniBroj|", _
                "email|OTKUI_HDM_EMAIL|txt|0|Email|", _
                "hladnjaca|OTKUI_HDM_HLADNJACA|txt|0|Hladnjaca|", _
                "racun|OTKUI_HDM_RACUN|txt|0|TekuciRacun|")
        Case "VOZACI"
            MatPolja = Array( _
                "ime|OTKUI_HDM_IME|txt|1|Ime|", _
                "prezime|OTKUI_HDM_PREZIME|txt|1|Prezime|", _
                "telefon|OTKUI_HDM_TELEFON|txt|0|Telefon|", _
                "pin|OTKUI_MP_PIN|txt|0|PIN|")
        Case "PARCELE"
            MatPolja = Array( _
                "kooperant|OTKUI_HDM_KOOPERANT|cmb|1|" & COL_PAR_KOOP & "|@kooperanti", _
                "katbroj|OTKUI_HDM_KATBROJ|txt|1|" & COL_PAR_KAT_BROJ & "|", _
                "katopstina|OTKUI_HDM_KATOPSTINA|txt|1|" & COL_PAR_KAT_OPSTINA & "|", _
                "kultura|OTKUI_HDM_KULTURA|cmb|1|" & COL_PAR_KULTURA & "|@kulture", _
                "povrsina|OTKUI_HDM_POVRSINA|num|1|" & COL_PAR_POVRSINA & "|", _
                "ggap|OTKUI_HDM_GGAP|cmb|1|" & COL_PAR_GGAP & "|@ggap", _
                "napomena|OTKUI_HDM_NAPOMENA|txt|0|" & COL_PAR_NAPOMENA & "|")
        Case "ARTIKLI"
            MatPolja = Array( _
                "naziv|OTKUI_HDM_NAZIV|txt|1|Naziv|", _
                "tip|OTKUI_HDM_TIP|cmb|1|Tip|@tipartikla", _
                "jm|OTKUI_HDM_JM|cmb|1|JedinicaMere|@jm", _
                "cena|OTKUI_HDM_CENA_JED|num|1|CenaPoJedinici|", _
                "doza|OTKUI_HDM_DOZA|num|1|DozaPoHa|", _
                "kultura|OTKUI_HDM_KULTURA|cmb|0|Kultura|@kulture", _
                "pakovanje|OTKUI_HDM_PAKOVANJE|num|1|Pakovanje|")
        Case "KULTURE"
            MatPolja = Array( _
                "vrsta|OTKUI_HDM_VRSTA|txt|1|VrstaVoca|", _
                "sorta|OTKUI_HDM_SORTA|txt|1|SortaVoca|", _
                "gajbica|OTKUI_HDM_GAJBICA_PAL|num|0|" & COL_KUL_GAJBICA_PALETA & "|", _
                "tipamb|OTKUI_HDM_TIP_AMB|cmb|0|" & COL_KUL_TIP_AMBALAZE & "|@tipambalaze", _
                "pragupoz|OTKUI_HDM_PRAG_UPOZ|num|0|" & COL_KUL_PRAG_PROSEK_UPOZ & "|", _
                "pragblok|OTKUI_HDM_PRAG_BLOK|num|0|" & COL_KUL_PRAG_PROSEK_BLOK & "|")
        Case "CENOVNIK"
            MatPolja = Array( _
                "vrsta|OTKUI_HDM_VRSTA|cmb|1|" & COL_CEN_VRSTA & "|@vrste", _
                "sorta|OTKUI_HDM_SORTA|cmb|0|" & COL_CEN_SORTA & "|@sorte", _
                "klasa|OTKUI_HDM_KLASA|cmb|1|" & COL_CEN_KLASA & "|@klase", _
                "datum|OTKUI_HDM_DATUM|date|0|" & COL_CEN_DATUM & "|", _
                "cena|OTKUI_HDM_CENA|num|1|" & COL_CEN_CENA & "|")
        Case "VRSTAGP"
            MatPolja = Array("tip|OTKUI_HDM_TIP_GP|txt|1|" & COL_VGP_TIP & "|")
        Case "AMBALAZA"
            MatPolja = Array( _
                "tip|OTKUI_HDM_TIP_AMB|txt|1|" & COL_TAMB_TIP & "|", _
                "tezina|OTKUI_HDM_TEZINA_GAJ|num|1|" & COL_TAMB_TEZINA & "|")
        Case "PALETE"
            MatPolja = Array( _
                "tip|OTKUI_HDM_TIP_PAL|txt|1|" & COL_TPAL_TIP & "|", _
                "tezina|OTKUI_HDM_TEZINA|num|1|" & COL_TPAL_TEZINA & "|")
        Case "KUTIJE"
            MatPolja = Array( _
                "tip|OTKUI_HDM_TIP_KUT|txt|1|" & COL_KUT_TIP & "|", _
                "tezina|OTKUI_HDM_TEZINA|num|1|" & COL_KUT_TEZINA & "|")
        Case "KESE"
            MatPolja = Array( _
                "tip|OTKUI_HDM_TIP_KES|txt|1|" & COL_KES_TIP & "|", _
                "tezina|OTKUI_HDM_TEZINA|num|1|" & COL_KES_TEZINA & "|")
    End Select
End Function

' Polje opisa polja: 0=kljuc 1=natpis 2=vrsta 3=obavezno 4=kolona 5=combo
Public Function PoljeF(ByVal spec As String, ByVal idx As Long) As String
    Dim p() As String
    p = Split(spec, "|")
    If idx > UBound(p) Then Exit Function
    PoljeF = p(idx)
End Function

' Opis JEDNOG polja po kljucu, ili "" ako ga sekcija nema.
Public Function MatPolje(ByVal kljuc As String, ByVal poljeKljuc As String) As String
    Dim a As Variant, r As Variant
    a = MatPolja(kljuc)
    If Not IsArray(a) Then Exit Function
    For Each r In a
        If PoljeF(CStr(r), 0) = poljeKljuc Then
            MatPolje = CStr(r)
            Exit Function
        End If
    Next r
End Function

' Stvarna kolona za polje: razresava "@alias:A,B" nad semom tabele. Vraca "" kad
' nijedna ne postoji -- pisac tada to polje PRESKACE umesto da obori ceo upis.
Public Function MatKolonaPolja(ByVal kljuc As String, ByVal spec As String) As String
    Dim kol As String, tbl As String, imena As Variant, ime As Variant
    kol = PoljeF(spec, 4)
    If Len(kol) = 0 Then Exit Function
    If Left$(kol, 7) <> "@alias:" Then
        MatKolonaPolja = kol
        Exit Function
    End If
    tbl = MatTabela(kljuc)
    If Len(tbl) = 0 Then Exit Function
    imena = Split(Mid$(kol, 8), ",")
    For Each ime In imena
        If GetColumnIndex(tbl, Trim$(CStr(ime))) > 0 Then
            MatKolonaPolja = Trim$(CStr(ime))
            Exit Function
        End If
    Next ime
End Function

' Prefiks novog ID-ja. Prazno znaci da sekcija NEMA surogat kljuc -- PK je sama
' unesena vrednost (tipovi ambalaze, paleta, kutija, kesa i gotovog proizvoda).
Public Function MatPrefiksID(ByVal kljuc As String) As String
    Select Case kljuc
        Case "KOOPERANTI": MatPrefiksID = "KOOP-"
        Case "STANICE":    MatPrefiksID = "ST-"
        Case "KUPCI":      MatPrefiksID = "KUP-"
        Case "VOZACI":     MatPrefiksID = "VOZ-"
        Case "PARCELE":    MatPrefiksID = "PAR-"
        Case "ARTIKLI":    MatPrefiksID = "ART-"
        Case "KULTURE":    MatPrefiksID = "KUL-"
    End Select
End Function

' Vrednost koja se pri UNOSU upisuje u kolonu statusa.
'
' Parcele dobijaju "Da", sve ostalo STATUS_AKTIVAN ("Aktivan"). To NIJE greska
' ovde nego zateceno stanje: legacy btnDodaj upisuje bas "Da" u tblParcele, dok
' soft-delete u istu kolonu upisuje "Aktivan"/"Neaktivan". Citac zato aktivnim
' smatra sve sto NIJE "Neaktivan" -- oba oblika prolaze. Poravnanje bi promenilo
' ono sto sinhronizacija vec vidi, pa se ne radi usput.
Public Function MatStatusNaUnosu(ByVal kljuc As String) As String
    If kljuc = "PARCELE" Then
        MatStatusNaUnosu = "Da"
    Else
        MatStatusNaUnosu = STATUS_AKTIVAN
    End If
End Function

'---------------------------------------------------------------- SORT
' Podrazumevani sort sekcije, oblika "kolona:asc" (ugovor Scr_Sort).
'
' Pravilo je JEDNO: rastuce po GLAVNOJ koloni, a glavna je ona koju opis oznaci
' kao rastegljivu ("part") -- ista ona koju mreza siri preko slobodnog prostora.
' Racuna se iz opisa kolona, pa dodavanje sekcije ne trazi jos jedan spisak koji
' bi se razisao sa prvim.
'
' CENOVNIK je jedini izuzetak i to je poslovno pravilo, ne ukus: cenovnik je
' append-only i vazeca cena je POSLEDNJA, pa se otvara po datumu opadajuce --
' isto sto legacy forma vec radi svojim sortom.
Public Function MatSort(ByVal kljuc As String) As String
    Dim cols As Variant, i As Long, dat As Long
    If kljuc = "CENOVNIK" Then
        dat = IndeksKolone(kljuc, COL_CEN_DATUM)
        If dat > 0 Then MatSort = CStr(dat) & ":desc"
        Exit Function
    End If
    cols = MatKolone(kljuc)
    If Not IsArray(cols) Then Exit Function
    For i = LBound(cols) To UBound(cols)
        If ColF(CStr(cols(i)), 2) = "part" Then
            MatSort = CStr(i - LBound(cols) + 1) & ":asc"
            Exit Function
        End If
    Next i
    MatSort = "1:asc"
End Function

' Redni broj kolone U MREZI (1-bazirano) po nazivu izvora. Vraca 0 ako je nema.
Private Function IndeksKolone(ByVal kljuc As String, ByVal izvor As String) As Long
    Dim cols As Variant, i As Long
    cols = MatKolone(kljuc)
    If Not IsArray(cols) Then Exit Function
    For i = LBound(cols) To UBound(cols)
        If ColF(CStr(cols(i)), 1) = izvor Then
            IndeksKolone = i - LBound(cols) + 1
            Exit Function
        End If
    Next i
End Function

'-------------------------------------------------------------- REDOVI
' Ugovor je isti kao kod svakog ekrana:
'   Array(kolone, redovi, n, zbirKg, zbirVal, brojaci cipova)
Public Function MatRedovi(ByVal kljuc As String, ByVal filter As String, _
                          ByVal q As String) As Variant
    Dim cols As Variant, data As Variant, tbl As String
    Dim nCol As Long, i As Long, c As Long, n As Long
    Dim outA() As Variant, hay As String, pkIdx As Long, statIdx As Long
    Dim v As Variant, statTxt As String, aktivan As Boolean
    Dim idx() As Long, izv As String

    On Error GoTo EH
    mUkupno = 0: mAktivnih = 0: mNeaktivnih = 0

    cols = MatKolone(kljuc)
    tbl = MatTabela(kljuc)
    If Not IsArray(cols) Or Len(tbl) = 0 Then
        MatRedovi = Array(Array(), Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    nCol = UBound(cols) - LBound(cols) + 1

    data = GetTableData(tbl)
    ' Cenovnik je append-only i cita se ISTIM redosledom kao u legacy formi:
    ' stornirano napolje, pa datum opadajuce sa tie-break-om po CenaID (kasniji
    ' unos gore). Bez toga bi vazeca cena zavisila od redosleda upisa u tabelu.
    If kljuc = "CENOVNIK" And Not IsEmpty(data) Then data = CenovnikPoredak(data)
    If IsEmpty(data) Then
        MatRedovi = Array(cols, Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If

    pkIdx = GetColumnIndex(tbl, MatPK(kljuc))
    statIdx = 0
    If Len(MatStatusKolona(kljuc)) > 0 Then _
        statIdx = GetColumnIndex(tbl, MatStatusKolona(kljuc))

    ' Indeksi izvornih kolona se traze JEDNOM, ne po redu. Kolona koje nema
    ' ostaje prazna (indeks 0) -- schema drift ne sme da obori celu listu, isto
    ' pravilo koje legacy vec primenjuje na Kupce.
    ReDim idx(0 To nCol - 1)
    For c = 0 To nCol - 1
        izv = ColF(CStr(cols(LBound(cols) + c)), 1)
        If Len(izv) > 0 And Left$(izv, 1) <> "@" Then idx(c) = GetColumnIndex(tbl, izv)
    Next c

    ReDim outA(1 To UBound(data, 1), 1 To nCol)
    For i = 1 To UBound(data, 1)
        If pkIdx = 0 Then GoTo Sledeci
        If Trim$(NzToText(data(i, pkIdx))) = "" Then GoTo Sledeci

        mUkupno = mUkupno + 1
        aktivan = True
        If statIdx > 0 Then
            statTxt = Trim$(NzToText(data(i, statIdx)))
            aktivan = (StrComp(statTxt, STATUS_NEAKTIVAN, vbTextCompare) <> 0)
            If aktivan Then mAktivnih = mAktivnih + 1 Else mNeaktivnih = mNeaktivnih + 1
        End If

        ' Cip radi SAMO tamo gde kolona statusa postoji. Gde je nema, "aktivni"
        ' se ne moze ni izabrati (MatCipovi ne prijavljuje nijedan cip).
        If statIdx > 0 Then
            If filter = MAT_CIP_AKT And Not aktivan Then GoTo Sledeci
            If filter = MAT_CIP_NEAKT And aktivan Then GoTo Sledeci
        End If

        n = n + 1
        hay = ""
        For c = 0 To nCol - 1
            v = VrednostCelije(kljuc, tbl, data, i, CStr(cols(LBound(cols) + c)), idx(c))
            outA(n, c + 1) = v
            hay = hay & "|" & NzToText(v)
        Next c
        ' Pretraga ide preko SVEGA sto operater vidi, ukljucujuci izvedene
        ' vrednosti -- inace se kooperant ne bi nasao po imenu stanice.
        If Len(q) > 0 Then
            If InStr(1, hay, q, vbTextCompare) = 0 Then
                n = n - 1
                GoTo Sledeci
            End If
        End If
Sledeci:
    Next i

    If n = 0 Then
        MatRedovi = Array(cols, Empty, 0, 0#, 0#, Array(0, 0, 0))
        Exit Function
    End If
    MatRedovi = Array(cols, outA, n, 0#, 0#, Array(0, 0, 0))
    Exit Function
EH:
    Err.Raise Err.Number, "modMaticniIzvor.MatRedovi[" & kljuc & "]", Err.description
End Function

' Poredak cenovnika, izdvojen zato sto je poslovno pravilo a ne prikaz:
' stornirano napolje, pa datum opadajuce (tie-break CenaID).
Private Function CenovnikPoredak(ByRef data As Variant) As Variant
    Dim cen As Variant, sortCol As Long, idCol As Long
    cen = ExcludeStornirano(data, TBL_CENOVNIK)
    If IsEmpty(cen) Then Exit Function
    idCol = GetColumnIndex(TBL_CENOVNIK, COL_CEN_ID)
    sortCol = GetColumnIndex(TBL_CENOVNIK, COL_CEN_DATUM)
    If sortCol = 0 Then sortCol = idCol
    If sortCol = 0 Then
        CenovnikPoredak = cen
        Exit Function
    End If
    CenovnikPoredak = SortArray(cen, sortCol, False, idCol)
End Function

' Vrednost jedne celije: ili gola kolona, ili izvedena vrednost.
'
' IZVEDENE VREDNOSTI SU PRENETE IZ LoadList, ne izmisljene. Svaka od njih je
' tamo vec postojala; ovde su na jednom mestu umesto u sest grana.
Private Function VrednostCelije(ByVal kljuc As String, ByVal tbl As String, _
                                ByRef data As Variant, ByVal r As Long, _
                                ByVal spec As String, ByVal kolIdx As Long) As Variant
    Dim izv As String
    izv = ColF(spec, 1)
    If Len(izv) = 0 Then Exit Function
    If Left$(izv, 1) <> "@" Then
        If kolIdx > 0 Then VrednostCelije = data(r, kolIdx)
        Exit Function
    End If

    Select Case izv
        Case "@status":       VrednostCelije = StatusTekst(kljuc, tbl, data, r)
        Case "@ime_prezime":  VrednostCelije = SpojiPolja(tbl, data, r, "Ime", "Prezime", " ")
        Case "@stanica":      VrednostCelije = NazivStanice(tbl, data, r)
        Case "@adresa_mesto": VrednostCelije = SpojiPolja(tbl, data, r, "Adresa", "Mesto", ", ")
        Case "@kupac_adresa": VrednostCelije = AdresaKupca(tbl, data, r)
        Case "@drzava":       VrednostCelije = DrzavaKupca(tbl, data, r)
        Case "@koop_naziv":   VrednostCelije = NazivKooperanta(tbl, data, r)
        Case "@geo":          VrednostCelije = GeoInfo(tbl, data, r)
    End Select
End Function

' Prazna kolona statusa znaci AKTIVAN -- isto sto legacy podrazumeva kad polje
' nije popunjeno. Prikazuje se pun tekst, ne "DA"/"NE".
Private Function StatusTekst(ByVal kljuc As String, ByVal tbl As String, _
                             ByRef data As Variant, ByVal r As Long) As String
    Dim c As Long, s As String
    If Len(MatStatusKolona(kljuc)) = 0 Then Exit Function
    c = GetColumnIndex(tbl, MatStatusKolona(kljuc))
    If c = 0 Then Exit Function
    s = Trim$(NzToText(data(r, c)))
    If StrComp(s, STATUS_NEAKTIVAN, vbTextCompare) = 0 Then
        StatusTekst = Poruka("OTKUI_MAT_NEAKTIVAN")
    Else
        StatusTekst = Poruka("OTKUI_MAT_AKTIVAN")
    End If
End Function

Private Function Kol(ByVal tbl As String, ByRef data As Variant, ByVal r As Long, _
                     ByVal ime As String) As String
    Dim c As Long
    c = GetColumnIndex(tbl, ime)
    If c > 0 Then Kol = Trim$(NzToText(data(r, c)))
End Function

' Spaja dva polja razdvojnikom; prazno polje ne ostavlja visece zareze.
Private Function SpojiPolja(ByVal tbl As String, ByRef data As Variant, ByVal r As Long, _
                            ByVal a As String, ByVal b As String, ByVal sep As String) As String
    Dim x As String, y As String
    x = Kol(tbl, data, r, a)
    y = Kol(tbl, data, r, b)
    If Len(x) = 0 Then
        SpojiPolja = y
    ElseIf Len(y) = 0 Then
        SpojiPolja = x
    Else
        SpojiPolja = x & sep & y
    End If
End Function

' Operater bira stanicu po NAZIVU, pa je i vidi po nazivu. StanicaID ostaje u
' podacima (kolona identiteta je PK reda, ne ovo).
Private Function NazivStanice(ByVal tbl As String, ByRef data As Variant, _
                              ByVal r As Long) As String
    Dim id As String
    id = Kol(tbl, data, r, "StanicaID")
    If Len(id) = 0 Then Exit Function
    NazivStanice = Trim$(CStr(LookupValue(TBL_STANICE, "StanicaID", id, "Naziv")))
    If Len(NazivStanice) = 0 Then NazivStanice = id
End Function

' Kooperant uz parcelu: "Ime Prezime (ID)" -- tacno kako legacy puni listu i
' combo, pa se pretraga po ID-ju i po imenu ponasa isto kao pre.
Private Function NazivKooperanta(ByVal tbl As String, ByRef data As Variant, _
                                 ByVal r As Long) As String
    Dim id As String, ime As String, prez As String
    id = Kol(tbl, data, r, COL_PAR_KOOP)
    If Len(id) = 0 Then Exit Function
    ime = Trim$(CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", id, "Ime")))
    prez = Trim$(CStr(LookupValue(TBL_KOOPERANTI, "KooperantID", id, "Prezime")))
    NazivKooperanta = Trim$(ime & " " & prez) & " (" & id & ")"
End Function

' Adresa kupca: ulica, postanski broj mesto -- isti sastav kao u LoadList.
Private Function AdresaKupca(ByVal tbl As String, ByRef data As Variant, _
                             ByVal r As Long) As String
    Dim s As String, posta As String, mesto As String
    s = Kol(tbl, data, r, "Ulica")
    posta = Kol(tbl, data, r, "PostanskiBroj")
    mesto = Kol(tbl, data, r, "Mesto")
    If Len(posta) > 0 Then
        If Len(s) > 0 Then s = s & ", "
        s = s & posta
    End If
    If Len(mesto) > 0 Then
        If Len(s) > 0 Then s = s & " "
        s = s & mesto
    End If
    AdresaKupca = s
End Function

' Drzava se na nekim instalacijama zove sa dijakritikom, na nekim bez -- ista
' tolerancija koju LoadList vec ima za Kupce.
Private Function DrzavaKupca(ByVal tbl As String, ByRef data As Variant, _
                             ByVal r As Long) As String
    DrzavaKupca = Kol(tbl, data, r, "Dr" & ChrW(382) & "ava")
    If Len(DrzavaKupca) = 0 Then DrzavaKupca = Kol(tbl, data, r, "Drzava")
End Function

Private Function GeoInfo(ByVal tbl As String, ByRef data As Variant, _
                         ByVal r As Long) As String
    Dim st As String, src As String
    st = Kol(tbl, data, r, COL_PAR_GEO_STATUS)
    src = Kol(tbl, data, r, COL_PAR_GEO_SOURCE)
    GeoInfo = st
    If Len(src) > 0 Then
        If Len(GeoInfo) > 0 Then GeoInfo = GeoInfo & " / "
        GeoInfo = GeoInfo & src
    End If
End Function

'-------------------------------------------------------------- BROJKE
' Brojke iz POSLEDNJEG citanja. Racunaju se u istom prolazu kao redovi, pa broj
' u zoni i lista u mrezi ne mogu da se razidju. "Ukupno" broji SVE zapise
' sekcije, bez obzira na cip -- inace bi brojka menjala znacenje sa cipom.
Public Function MatUkupno() As Long
    MatUkupno = mUkupno
End Function

Public Function MatAktivnih() As Long
    MatAktivnih = mAktivnih
End Function

Public Function MatNeaktivnih() As Long
    MatNeaktivnih = mNeaktivnih
End Function

Public Sub MatResetCache()
    Set mStatusKol = Nothing
End Sub
