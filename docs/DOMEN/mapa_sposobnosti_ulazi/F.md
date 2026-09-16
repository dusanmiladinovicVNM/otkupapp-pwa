# Mapa sposobnosti -- oblast F, deo F1: MATICNI PODACI I PRIJAVA

Moduli: `modScrMatPartneri` + `modScrMatRoba` + `modScrMatPakovanje` +
`modScrMatKorisnici` (tanke ljuske), zajednicki motor `modMaticniEkran`, izvor
`modMaticniIzvor`, pisac `modMaticniUnos`, pisac naloga i prava
`modMaticniKorisnici`, geo panel `modMaticniGeo` (+ `modGeoParcele`), prijava
`modAuth` + faze ljuske `modUiFaze`, zivotni ciklus sveske
`ThisWorkbook.doccls` + `modMain`.

Podoblasti: **F1a** motor maticnih ekrana - **F1b** sekcije (sifarnici) -
**F1c** geo panel parcele - **F1d** nalozi i prava - **F1e** prijava i zamena
operatera - **F1f** zivotni ciklus sveske (otvaranje, snimanje, zatvaranje).

Ovo je PRVI deo oblasti F. Admin panel, Podesavanja, Setup, self-update i
izdanja, integritet, health i ostali makroi su **deo F2**, nize u istom fajlu
(F-053 nadalje).

Ne ponavlja se ono sto je vec mapirano: upis nove cene u cenovnik je **D-061**,
citanje vazece cene je **D-059**, oporavak zaglavljenih SEF faktura pri startu
je **D-024**, pokretanje i gasenje zakazanog sync-a su **E-008** i **E-009**,
ciscenje zaostalih stanica-lockova pri otvaranju sveske je **E-029**.

KO: **operater** = svakodnevni rad na ekranu; **administrator** = nalozi, prava
i sve iza `modAuth.MozeAdministraciju`; **odrzavanje** = Alt+F8 i oporavak.

## F1a -- motor maticnih ekrana (vazi za sve sekcije)

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-001 | operater | Doci do sifarnika kroz tri ekrana sekcije MATICNI, a do naloga kroz cetvrti koji trazi jos i administraciju | sidebar, sekcija `SEK_MATICNI` -- redovi registra `MAT_PARTNERI` (`modUiScreens.bas:161`), `MAT_ROBA` (`:163`), `MAT_PAKOVANJE` (`:165`), `MAT_KORISNICI` (`:171`) | otvoren ekran sa prekidacem lista; zabranjen ekran je prigusena stavka | Sva cetiri stoje iza prava `OBL_MATICNI` (kolona `SCR_OBLAST` u registru); `MAT_KORISNICI` nosi JOS jednu branu -- `Scr_Dozvoljen` trazi `modAuth.MozeAdministraciju` (`modScrMatKorisnici.bas:41-43`), jer su prava jedina lista na kojoj se pristup moze dodeliti samom sebi | cita `tblKorisnici` (kolone oblasti) | `modUiScreens.ScrRows:113`, `modScrMatKorisnici.Scr_Dozvoljen:41` | ne | `modTest.T_Matic_SekcijaTraziPravo:1183` -- "bez prava na Maticne podatke nijedan njihov ekran ni panel ne sme", "admin sme sve iz sekcije Maticni", "prekidac sekcije se NE crta bez prava na Maticne podatke" |
| F-002 | operater | Videti jednu listu sifarnika sa brojem zapisa, brojem aktivnih i neaktivnih, i suziti je pretragom | prekidac lista (`modMaticniIzvor.MatSekcijeEkrana:94`), citanje reda kroz `modMaticniEkran.Redovi:479` | mreza + zona: naslov liste, "N zapisa", AKTIVNIH / NEAKTIVNIH (`modMaticniEkran.bas:424-428`) | Brojke se racunaju U ISTOM prolazu kroz podatke kao i redovi, pa se prikaz i brojka ne mogu razici (`modMaticniEkran.bas:485-487`); sekcija bez kolone statusa umesto brojke pokazuje crticu (`modMaticniEkran.bas:446-447`) | cita tabelu sekcije preko `modMaticniIzvor.MatTabela:155` | `modMaticniIzvor.MatRedovi:783`, `MatUkupno:1087`, `MatAktivnih:1091`, `MatNeaktivnih:1095` | ne | `modTest.T_MatUnos_OpisPoljaISema:15368` -- "svako polje pise u kolonu koja POSTOJI u semi" |
| F-003 | operater | Suziti listu na aktivne ili neaktivne zapise jednim cipom | cipovi `sve` / `aktivni` / `neaktivni` (`modMaticniIzvor.MatCipovi:247`) | filtrirana mreza; brojke u zoni ostaju za CELU sekciju | Sekcija koja u semi nema kolonu statusa NE prijavljuje nijedan cip -- dugme koje tiho ne filtrira nista je gore od dugmeta kog nema (`modMaticniIzvor.MatCipovi:247-252`); naziv kolone se TRAZI u semi (`Aktivan` ili `Aktivna`), ne pogadja se (`modMaticniIzvor.MatStatusKolona:207`) | cita kolonu statusa sekcije | `modMaticniIzvor.MatStatusKolona:207`, `MatCipovi:247` | ne | `modTest.T_MatEkran_RadnjeIRezim:15656` -- "svaka radnja ima pet polja opisa" |
| F-004 | operater | Preci na drugu listu ili drugi ekran bez tihog gubitka otkucanog unosa | prekidac lista `ls<KLJUC>` (`modMaticniEkran.bas:540`), promena ekrana iz sidebara (`modOtkupUI.bas:5525`) | `MsgBox` `Poruka("MATU_ASK_ODBACI_UNOS")` = "Otvoren unos nije sacuvan. Odbaciti ga i preci na drugu listu?" | Polja pripadaju sekciji koja se napusta -- ostavljena bi trazila upis u DRUGU tabelu, pa se pita umesto da unos tiho nestane (`modMaticniEkran.bas:543-548`); ekran javlja ima li nesacuvano kroz `Scr_ImaNesacuvano` -> `ImaNesacuvano:907` | -- | `modMaticniEkran.ObradiDogadjaj:540`, `ImaNesacuvano:907`, `modUiScreens.ScrImaNesacuvano:368` | ne | NEPROVERENO -- nije nadjen test koji tvrdi bas pitanje pri promeni liste (`MATU_ASK_ODBACI_UNOS`) |
| F-005 | operater | Uneti nov zapis u sifarnik kroz editor u zoni ekrana | dugme `scrMatNovi` (`modMaticniEkran.bas:593`), pa `scrMatSacuvaj` (`:595`) -> `modMaticniUnos.MatDodaj:91` | nov red u tabeli + toast `Poruka("MATU_OK_DODATO")` = "Dodato:" sa dodeljenim ID-em (`modMaticniEkran.bas:868-869`) | ID dodeljuje `GetNextID` sa prefiksom sekcije (`modMaticniUnos.bas:131`, prefiksi `modMaticniIzvor.MatPrefiksID:557`); obavezno polje bez vrednosti se odbija imenovano (`MATU_ERR_OBAVEZNO`, `modMaticniUnos.bas:395`); combo polje MORA da bira iz svoje liste -- slobodan tekst se odbija sa `MATU_ERR_VAN_LISTE` = "izaberite vrednost iz ponudjene liste" (`modMaticniUnos.bas:404-409`); zapis se otvara kao aktivan (`modMaticniIzvor.MatStatusNaUnosu:577`); upis ide u transakciji, uz `SchemaReadyOrFail` pre prve celije (`modMaticniUnos.bas:148`) | pise tabelu sekcije (`MatTabela:155`), kolona po kolona po imenu (`UpisiPolja:462`) | `modMaticniEkran.Sacuvaj:846` -> `modMaticniUnos.MatDodaj:91`, provera `Proveri:380` | ne | `modTest.T_MatUnos_ProveraOdbija:15479` -- "kooperant bez imena se odbija", "odbijeno polje se imenuje (fokus)"; `modTest.T_Maticni_KapijeUpisaIZivotniCiklus:16247` -- "izmisljena stanica se odbija pre upisa", "kvar pri citanju spiska ODBIJA vrednost, ne propusta je" |
| F-006 | operater | Izmeniti postojeci zapis -- dvoklikom na red ili dugmetom "Izmeni" | dvoklik `dbl:<red>` (`modMaticniEkran.bas:564`), radnja `izmeni` (`:715`) -> `modMaticniUnos.MatIzmeni:187` | polja editora popunjena zatecenim vrednostima; po snimanju toast `Poruka("MATU_OK_IZMENJENO")` = "Izmenjeno." | Red se bira po IDENTITETU (`MatRedPoID:317`), nikad po poziciji u mrezi -- mreza se sortira i pretrazuje, a istoimeni zapisi su u sifarnicima obicna pojava (`modMaticniUnos.bas:313-316`); identitet reda se cita iz kolone koju sekcija prijavljuje (`MatKolonaID:237`); editor otvoren na DRUGOM ekranu ne sme da pise odavde -- odbija se sa `MATU_ERR_TUDJI_EDITOR` (`modMaticniEkran.bas:851-856`) | cita i pise tabelu sekcije | `modMaticniEkran.OtvoriIzmenu:742`, `modMaticniUnos.MatIzmeni:187`, `MatVrednostiReda:341` | ne | `modTest.T_Maticni_KapijeUpisaIZivotniCiklus:16247` (isti test nosi i tvrdnju o prirodnom PK) |
| F-007 | operater | Deaktivirati ili ponovo aktivirati zapis, uz potvrdu | radnja `status` (`modMaticniEkran.bas:716`) -> `PromeniStatus:820` -> `modMaticniUnos.MatPromeniStatus:255` | `MsgBox` `Poruka("OTKUI_MAT_STATUS_ASK")` = "Promeniti status ovog zapisa? Vidi se u svim izvestajima i listama.", pa toast `Poruka("MATU_OK_STATUS")` = "Status promenjen u:" | Soft-delete, ne brisanje: status se OBRCE u istoj koloni (`modMaticniUnos.bas:288-293`); sekcija bez kolone statusa se ODBIJA porukom `MATU_ERR_NEMA_STATUSA` umesto da dugme tiho ne uradi nista (`modMaticniUnos.bas:275-278`) | pise kolonu statusa sekcije (`Aktivan` / `Aktivna`) | `modMaticniEkran.PromeniStatus:820`, `modMaticniUnos.MatPromeniStatus:255` | ne | `modTest.T_MatEkran_RadnjeIRezim:15656` -- radnja "Deaktiviraj" postoji samo gde sekcija stvarno ima kolonu statusa |
| F-008 | operater | Dobiti spisak sorti koji prati izabranu vrstu voca, i pri otvaranju i pri promeni polja | promena teksta `chg:<kontrola>` (`modMaticniEkran.bas:583`) -> `OsveziZavisne:1021`; zavisnost iz opisa polja (`modMaticniIzvor.MatComboZavisi:597`) | prepunjen combo sorti | Bez izabrane vrste spisak sorti je PRAZAN -- prazan spisak je tacan odgovor, spisak svih sorti nije (`modMaticniIzvor.bas:646-650`); isti spisak zavisnosti sluzi punjenju pri otvaranju, ponovnom punjenju i punjenju posle ucitavanja zapisa (`modMaticniIzvor.bas:592-597`) | cita `tblKulture` (`VrstaVoca`, `SortaVoca`) | `modMaticniIzvor.MatComboStavke:617`, `modMaticniEkran.OsveziZavisne:1021` | ne | `modTest.T_MatEkran_KaskadaZavisnogCombo:15996` -- "sorta cenovnika zavisi od vrste", "bez izabrane vrste spisak sorti je PRAZAN, ne spisak svih sorti" |
| F-009 | administrator | Biti odbijen na upisu i kad se do pisca dodje mimo ekrana (Alt+F8, nov pozivalac) | `modMaticniUnos.MatBranaUpisa:71`, zvana iz `MatDodaj:91`, `MatIzmeni:187`, `MatPromeniStatus:255`, i iz `modMaticniKorisnici.Brana:241` | poruka `MATU_ERR_BEZ_PRAVA` = "Nemate pravo izmene maticnih podataka", odnosno `AUTH_MSG_SAMO_ADMIN_SEKCIJA` = "Ovoj sekciji pristupa samo administrator" | Tvrda kapija stoji U PISCU, ne samo u ekranu -- brana ekrana pada zajedno sa svojim ekranom (`modMaticniUnos.bas:60-64`); nalozi i prava traze JOS administraciju (`modMaticniUnos.bas:83-86`) | cita `tblKorisnici.MaticniPodaci`, `tblKorisnici.Uloga` | `modMaticniUnos.MatBranaUpisa:71`, `modAuth.KorisnikImaPravo:282`, `MozeAdministraciju:314` | ne | `modTest.T_Maticni_KapijeUpisaIZivotniCiklus:16247` -- pisac naloga nosi kapiju sam (`modTest.bas:16329-16340`) |

## F1b -- sekcije (koji sifarnici postoje i sta je u svakom posebno)

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-010 | operater | Voditi kooperante: ime, prezime, mesto, telefon, otkupno mesto, BPG broj, tekuci racun, PIN, adresa, JMBG | ekran `MAT_PARTNERI`, lista `KOOPERANTI` (`modMaticniIzvor.bas:98`) | red u `tblKooperanti` sa ID-em `KOOP-...` | Otkupno mesto je OBAVEZNO i bira se iz spiska stanica (`modMaticniIzvor.bas:411`); spisak je oblika "Naziv (ID)" (`PrikazLista:680`) | pise `tblKooperanti` (`Ime`, `Prezime`, `Mesto`, `Telefon`, `StanicaID`, `BPGBroj`, `TekuciRacun`, `Pin`, `Adresa`, `JMBG`) | `modMaticniIzvor.MatPolja:404-413`, upis kroz F-005/F-006 | ne | `modTest.T_MatUnos_ProveraOdbija:15479` -- "kooperant sa imenom, prezimenom i stanicom prolazi" |
| F-011 | operater | Voditi otkupna mesta (stanice), ukljucujuci oznaku da li je stanica hladnjaca | ekran `MAT_PARTNERI`, lista `STANICE` (`modMaticniIzvor.bas:99`) | red u `tblStanice` sa ID-em `ST-...` | Telefon i kontakt podaci se citaju kroz alias jer se u zatecenoj semi kolona zove razlicito (`@alias:Kontakt,Telefon`, `modMaticniIzvor.bas:418`); u malina rezimu nova stanica ODMAH dobija par-vozaca sa istim ID-em, van transakcije i bez prava da obori unos stanice (`modMaticniUnos.bas:168-177`) | pise `tblStanice` (`Naziv`, `Mesto`, `Kontakt`, `Ime`, `Prezime`, `PIN`, `JeHladnjaca`); posredno `tblVozaci` (`VozacID` = `StanicaID`) | `modMaticniIzvor.MatPolja:415-421`, `modMalina.EnsureVozacMirrorForStanica:65` | ne | `modBusinessFlowProTests.bas:1716-1737` -- mirror je idempotentan i odbija se za nepostojecu stanicu |
| F-012 | operater | Voditi kupce sa punom adresom i poreskim podacima (PIB, maticni broj, tekuci racun, e-posta, hladnjaca) | ekran `MAT_PARTNERI`, lista `KUPCI` (`modMaticniIzvor.bas:100`) | red u `tblKupci` sa ID-em `KUP-...` | Drzava se cita kroz alias (`@alias:Drzava`, `modMaticniIzvor.bas:431`) jer zatecena sema nosi oba pisanja; adresa i drzava se u mrezi prikazuju kao izvedene kolone (`AdresaKupca:1046`, `DrzavaKupca:1065`) | pise `tblKupci` (`Naziv`, `Ulica`, `Mesto`, `PostanskiBroj`, `Drzava`, `PIB`, `MaticniBroj`, `Email`, `Hladnjaca`, `TekuciRacun`) | `modMaticniIzvor.MatPolja:425-434` | ne | NEPROVERENO -- nije nadjen test bas za sekciju `KUPCI` (pokrivena je samo opstim prolazom kroz sve sekcije u `modTest.T_MatUnos_OpisPoljaISema:15368`) |
| F-013 | operater | Voditi vozace (ime, prezime, telefon, PIN) | ekran `MAT_PARTNERI`, lista `VOZACI` (`modMaticniIzvor.bas:101`) | red u `tblVozaci` sa ID-em `VOZ-...` | Sifarnik je nezavisan od dodele vozaca dokumentu -- ovde se vozac samo zavodi (`modMaticniIzvor.MatPolja:436-441`) | pise `tblVozaci` (`Ime`, `Prezime`, `Telefon`, `PIN`) | `modMaticniIzvor.MatPolja:436-441` | ne -- sifarnik ne cita ni jedno polje zaglavlja otkupa; veza `Otkup.VozacID` je oblast A/E | NEPROVERENO -- nije nadjen test bas za sekciju `VOZACI` |
| F-014 | operater | Voditi parcele kooperanata: katastarski broj i opstina, kultura, povrsina, GGAP status, napomena | ekran `MAT_PARTNERI`, lista `PARCELE` (`modMaticniIzvor.bas:102`) | red u `tblParcele` sa ID-em `PAR-...` | Povrsina mora biti STROGO pozitivna -- nula hektara nije podatak nego prazan unos (`modMaticniUnos.TraziPozitivan:447-450`); kooperant, kultura i GGAP su obavezni i biraju se iz spiska (`modMaticniIzvor.bas:443-450`); parcela se pri unosu upisuje sa statusom "Da", ne "Aktivan" -- tako je zateceno i tako sinhronizacija vec vidi (`modMaticniIzvor.MatStatusNaUnosu:577-584`) | pise `tblParcele` (`KooperantID`, `KatBroj`, `KatOpstina`, `Kultura`, `PovrsinaHa`, `GGAPStatus`, `Napomena`, `Aktivna`) | `modMaticniIzvor.MatPolja:443-450` | ne | `modTest.T_MatUnos_OpisPoljaISema:15368` -- "PK nije polje za rucni unos tamo gde ga dodeljuje GetNextID" |
| F-015 | operater | Voditi artikle agrohemije: tip, jedinica mere, cena po jedinici, doza po hektaru, kultura, pakovanje | ekran `MAT_ROBA`, lista `ARTIKLI` (`modMaticniIzvor.bas:105`) | red u `tblArtikli` sa ID-em `ART-...` | Tip je zatvoren spisak (`Pesticid` / `Djubrivo` / `SadniMaterijal`, `modMaticniIzvor.bas:638`), jedinica mere takodje (`kg` / `l` / `kom`, `:639`); cena artikla je JEDNA tekuca cena po artiklu i NE ide kroz cenovnik otkupa (v. **D-061**) | pise `tblArtikli` (`Naziv`, `Tip`, `JedinicaMere`, `CenaPoJedinici`, `DozaPoHa`, `Kultura`, `Pakovanje`) | `modMaticniIzvor.MatPolja:453-460` | ne | `modTest.T_MatUnos_OpisPoljaISema:15368` -- "vrsta polja je iz zatvorenog skupa", "svaki combo ima poznat izvor stavki" |
| F-016 | operater | Voditi kulture (vrsta + sorta) sa gajbicama po paleti, tipom ambalaze i pragovima proseka | ekran `MAT_ROBA`, lista `KULTURE` (`modMaticniIzvor.bas:106`) | red u `tblKulture` sa ID-em `KUL-...` | Jedino pravilo koje gleda DVA polja odjednom: prag blokade ne sme biti ispod praga upozorenja, inace `MATU_ERR_PRAG_BLOK` (`modMaticniUnos.bas:430-440`); ista tabela je izvor spiska vrsta i sorti za cenovnik i za unos otkupa (`modMaticniIzvor.bas:634`, `:652-668`) | pise `tblKulture` (`VrstaVoca`, `SortaVoca`, `GajbicaPoPaleti`, `TipAmbalaze`, `PragProsekUpoz`, `PragProsekBlok`) | `modMaticniIzvor.MatPolja:461-468`, provera `modMaticniUnos.Proveri:430` | ne | `modTest.T_MatEkran_KaskadaZavisnogCombo:15996` -- "fixture ima bar jednu vrstu" |
| F-017 | operater | Videti istoriju cena otkupa kao listu i dodati novu vazecu cenu bez ponovnog kucanja proizvoda | ekran `MAT_ROBA`, lista `CENOVNIK` (`modMaticniIzvor.bas:107`); dvoklik na stari red otvara NOV unos sa istim proizvodom (`modMaticniEkran.OtvoriIzmenu:766-785`) | nov red u `tblCenovnik`; toast `Poruka("MATU_ERR_CENOVNIK_APPEND")` = "Cenovnik se ne menja -- nova cena se DODAJE kao nov vazeci red." | Cenovnik je APPEND-ONLY: sekcija NEMA radnju "Izmeni" (`modMaticniEkran.Radnje:505`), a `MatIzmeni` je izricito odbija (`modMaticniUnos.bas:202-207`); posle snimanja editor OSTAJE otvoren sa istim proizvodom i praznom cenom, jer se obicno unosi vise cena zaredom (`modMaticniEkran.bas:874-877`); prazan datum znaci DANAS (`modMaticniUnos.DodajCenu:549`). Sam upis je **D-061**, citanje vazece cene je **D-059** | cita/pise `tblCenovnik` (`Vrsta`, `Sorta`, `Klasa`, `Datum`, `Cena`) | `modMaticniIzvor.MatPolja:469-475`, `modMaticniEkran.OtvoriIzmenu:742`, `modMaticniUnos.DodajCenu:536` | ne | `modTest.T_MatEkran_RadnjeIRezim:15656` -- cenovnik je append-only pa "Izmeni" nema |
| F-018 | operater | Voditi cetiri sifarnika pakovanja (ambalaza, palete, kutije, kese), svaki kao par tip + tezina | ekran `MAT_PAKOVANJE`, liste `AMBALAZA`, `PALETE`, `KUTIJE`, `KESE` (`modMaticniIzvor.bas:110-113`) | red u `tblTipAmbalaze` / `tblTipPalete` / `tblKutije` / `tblKese` | Naziv JESTE kljuc (nema surogata, `modMaticniIzvor.MatPK:189-192`), pa se duplikat odbija PRE upisa sa `MATU_ERR_VEC_POSTOJI` -- inace bi dva tipa istog imena nizvodno znacila dve razlicite tezine za istu gajbicu (`modMaticniUnos.bas:134-141`); postojeci naziv se NE preimenuje: `MATU_ERR_PK_ZAKLJUCAN` = "Naziv je kljuc ovog sifarnika i ne menja se -- dodajte nov zapis, pa deaktivirajte stari" (`modMaticniUnos.bas:592`); kolona statusa se TRAZI u semi pri svakom crtanju, ne pogadja se (`modMaticniIzvor.MatStatusKolona:207`); po kanonu (`schema/schema.json`) sve cetiri tabele imaju `Aktivan`, pa sve cetiri imaju i cipove i "Deaktiviraj" -- v. nalaz 3 | pise `tblTipAmbalaze` (`TipAmbalaze`, `TezinaGajbiceKg`, `Aktivan`), `tblTipPalete` (`TipPalete`, `TezinaKg`, `Aktivan`), `tblKutije` (`TipKutije`, `TezinaKg`, `Aktivan`), `tblKese` (`TipKese`, `TezinaKg`, `Aktivan`) | `modMaticniIzvor.MatPolja:477-492`, `modMaticniUnos.PkNepromenjen:567` | ne | `modTest.T_Maticni_KapijeUpisaIZivotniCiklus:16247` -- prirodan PK se pri izmeni ne preimenuje (`modTest.bas:16249-16251`) |
| F-019 | operater | Voditi spisak vrsta gotovog proizvoda | ekran `MAT_ROBA`, lista `VRSTAGP` (`modMaticniIzvor.bas:108`) | red u `tblVrstaGotovihProizvoda` | Kao i pakovanje, tip JESTE kljuc (`modMaticniIzvor.MatPK:187`) -- duplikat se odbija, naziv se ne preimenuje; sekcija ima tacno JEDNO polje za unos (`modMaticniIzvor.MatPolja:476`), a kanonska tabela ima tri kolone -- `RokMeseci` se sa ekrana ne moze uneti (v. nalaz 3) | pise `tblVrstaGotovihProizvoda` (`TipGotovogProizvoda`) | `modMaticniIzvor.MatPolja:476` | ne | `modTest.T_MatEkran_BazenPoljaIVisina:15579` -- test prolazi kroz svih 13 sekcija |

## F1c -- geo panel parcele

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-020 | operater | Otvoriti GeoSrbija portal sa vec pripremljenom pretragom izabrane parcele | radnja `geo` (`modMaticniEkran.bas:717`) -> dugme `scrGeoPortal` (`:598`) -> `modMaticniGeo.GeoOtvoriPortal:60` | otvoren `https://a3.geosrbija.rs/` + pretraga "katbroj opstina" u klipbordu | Parcela bez katastarskog broja ili opstine se ODBIJA (`MATG_ERR_NEMA_KATASTRA`) -- portal bi se otvorio na praznu pretragu, a operater bi mislio da je nesto uradjeno (`modMaticniGeo.bas:68-72`); prefiks "KO " se skida jer ga portal ne prepoznaje (`modMaticniGeo.bas:74`) | cita `tblParcele` (`KatBroj`, `KatOpstina`) | `modMaticniGeo.GeoOtvoriPortal:60` | ne | NEPROVERENO -- `ThisWorkbook.FollowHyperlink` i klipbord traze Excel; test pokriva samo sastav adrese i parsiranje teksta |
| F-021 | operater | Otvoriti Google mape tacno na tacki izabrane parcele | dugme `scrGeoMape` (`modMaticniEkran.bas:599`) -> `modMaticniGeo.GeoOtvoriMape:82` | otvorena mapa na `lat,lng` | Parcela bez upotrebljive tacke se odbija sa `MATG_ERR_NEMA_TACKE`; prazno i neispravno se tretiraju isto, jer se mapa ne moze otvoriti ni na jedno od toga (`modMaticniGeo.GeoTacka:158-165`); adresa mora imati decimalnu TACKU -- na masini sa zarezom kao separatorom URL bi bio neispravan | cita `tblParcele` (`Lat`, `Lng`) | `modMaticniGeo.GeoOtvoriMape:82`, `GeoUrlMape:41` | ne | `modTest.T_MatGeo_TekstIAdrese:15724` -- adresa Google Maps mora imati decimalnu tacku (`modTest.bas:15728-15729`) |
| F-022 | operater | Otvoriti poligon izabrane parcele na portalu | dugme `scrGeoPoligon` (`modMaticniEkran.bas:600`) -> `modMaticniGeo.GeoOtvoriPoligon:95` | otvoren portal na poligonu parcele | Bez parcele nema sta da se otvori -- `MATG_ERR_NEMA_PARCELE` (`modMaticniGeo.bas:98-101`) | cita `tblParcele` (poligon) | `modMaticniGeo.GeoOtvoriPoligon:95`, `GeoUrlPoligon:47` | ne | NEPROVERENO -- `FollowHyperlink` trazi Excel i pregledac |
| F-023 | operater | Nalepiti koordinate prekopirane sa portala i dobiti popunjena polja N i E | dugme `scrGeoNalepi` (`modMaticniEkran.bas:597`) -> `GeoNalepi:632` -> `modMaticniGeo.GeoIzTeksta:207` | popunjena polja `scrGeoN` / `scrGeoE` + toast `MATG_OK_NALEPLJENO`; neuspeh: `MATG_ERR_NALEPLJENO` | Uzimaju se PRVA DVA broja veca od 1000 -- UTM34 nad Srbijom je sedmocifren, pa se time odbacuju broj parcele, godina i sve ostalo sto se zateklo u redu (`modMaticniGeo.bas:203-206`, prag `:231`); oznake `N=`, `E:` i zagrade se skidaju (`OcistiToken:246`) | -- (samo polja panela) | `modMaticniEkran.GeoNalepi:632`, `modMaticniGeo.GeoIzTeksta:207` | ne | `modTest.T_MatGeo_TekstIAdrese:15724` -- "koordinate sa oznakama N=/E= se prepoznaju", "red sa portala daje koordinate", "mali brojevi se preskacu -- broj parcele nije koordinata" |
| F-024 | operater | Sacuvati koordinate parcele; aplikacija sama racuna lat/lng iz UTM-a | dugme `scrGeoSacuvaj` (`modMaticniEkran.bas:601`) -> `GeoSacuvajKlik:645` -> `modMaticniGeo.GeoSacuvaj:112` | upisana tacka + toast `MATG_OK_SACUVANO`; greska imenuje polje (fokus na N ili E) | Koordinate moraju biti POZITIVNE -- UTM34 nad Srbijom nema negativnih, a nula znaci "nije uneto" (`modMaticniGeo.bas:129-133`); upis ide u transakciji i puni N, E, Lat, Lng, status, izvor, meteo i dva datuma odjednom (`modGeoParcele.bas:72-79`) | pise `tblParcele` (`N_Coord`, `E_Coord`, `Lat`, `Lng`, `GeoStatus`, `GeoSource`, `MeteoEnabled`, `DatumGeoUnosa`, `DatumAzuriranja`) | `modMaticniGeo.GeoSacuvaj:112` -> `modGeoParcele.SaveParcelGeoPointByID:13` -> `SaveParcelGeoPointByID_TX:40` | ne | NEPROVERENO -- upis trazi Excel; testom je pokriven samo racun `ConvertUTM34ToLatLng` kroz `modTest.T_MatGeo_TekstIAdrese:15724` |
| F-025 | operater | Obrisati geo podatak parcele, uz potvrdu | dugme `scrGeoObrisi` (`modMaticniEkran.bas:602`) -> `GeoObrisiKlik:661` -> `modMaticniGeo.GeoObrisi:142` | `MsgBox` `Poruka("MATG_ASK_OBRISI")` sa ID-em parcele, pa toast `MATG_OK_OBRISANO` i ispraznjena polja | Trazi se potvrda jer se gube i tacka i poligon, a poligon se rucno crta (`modMaticniEkran.bas:657-659`) | pise `tblParcele` (geo kolone) | `modMaticniEkran.GeoObrisiKlik:661` -> `modGeoParcele.ClearParcelGeoByID:108` | ne | NEPROVERENO -- brisanje trazi Excel |
| F-026 | operater | Videti u traci panela da li izabrana parcela uopste ima tacku, odakle je i ima li poligon | traka GEO panela, crta se pri svakom rasporedu (`modMaticniEkran.RasporediGeo:216` -> `modMaticniGeo.GeoOpis:178`) | jedan red: ID parcele, koordinate na pet decimala, status / izvor, oznaka poligona | Ono sto je legacy pokazivao kroz `lblGeoStatus` posle svake radnje, ovde stoji STALNO (`modMaticniGeo.bas:174-176`); bez izabrane parcele traka to i kaze (`MATG_OPIS_NEMA_IZBORA`, `modMaticniGeo.bas:181`) | cita `tblParcele` (`GeoStatus`, `GeoSource`, `PolygonGeoJSON`, `Lat`, `Lng`) | `modMaticniGeo.GeoOpis:178` | ne | NEPROVERENO -- crta se iz rasporeda forme, sto trazi Excel |

## F1d -- nalozi i prava (ekran MAT_KORISNICI)

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-027 | administrator | Videti sve naloge sa korisnickim imenom, imenom, ulogom, aktivnoscu i sazetkom prava | ekran `MAT_KORISNICI`, lista `KORISNICI` (`modMaticniIzvor.bas:120`) -> `modMaticniKorisnici.KorRedovi:96` | mreza: ID, Username, Ime, Uloga, Aktivan, Prava (`KorKolone:86`) | Aktivan je SVE sto nije "NE" -- isti kriterijum koji `modAuth` primenjuje pri prijavi, pa se zapis sa zatecenim "Neaktivan" prikazuje onako kako se stvarno ponasa (`modMaticniKorisnici.bas:124-128`); brojke se pune PRE cipa, da ne menjaju znacenje sa filtrom (`:131-134`) | cita `tblKorisnici` (`KorisnikID`, `Username`, `ImePrezime`, `Uloga`, `Aktivan`, kolone oblasti) | `modMaticniKorisnici.KorRedovi:96`, `PravaOpis:627` | ne | `modTest.T_MatKor_RecnikDaNeIPrava:15794` -- "zatecena vrednost se ne prepisuje kao dozvola" |
| F-028 | administrator | Otvoriti nov nalog: korisnicko ime, ime, PIN, uloga, aktivnost, otkupno mesto | dugme `scrMatNovi` na listi `KORISNICI` -> `modMaticniUnos.MatDodaj:91` -> `modMaticniKorisnici.KorDodaj:245` | nov red u `tblKorisnici` sa ID-em `KOR-...` i upisanim `CreatedAt` | Nalozi imaju SVOG pisca jer opsti upis ne zna ni jedno od ovoga: PIN se hesira, uloga i aktivnost se pisu recnikom "DA"/"NE" koji cita `modAuth`, a prava su kolone istog reda (`modMaticniUnos.bas:108-114`); PIN je OBAVEZAN pri unosu -- `MATK_ERR_PIN` = "PIN je obavezan za novog korisnika" (`modMaticniKorisnici.bas:255-260`); korisnicko ime je obavezno i jedinstveno; nov korisnik je podrazumevano aktivan | pise `tblKorisnici` (`KorisnikID`, `Username`, `ImePrezime`, `PIN`, `Uloga`, `Aktivan`, `StanicaID`, `CreatedAt`, kolone oblasti) | `modMaticniKorisnici.KorDodaj:245`, `UpisiPolja:487`, `UpisiPrava:522` | ne | `modTest.T_MatKor_RecnikDaNeIPrava:15794` -- "nov korisnik je podrazumevano aktivan", "korisnicko ime je obavezno i jedinstveno" (`modTest.bas:15797-15800`) |
| F-029 | administrator | Izmeniti nalog, a da PIN ostane isti ako se ne kuca | dvoklik / "Izmeni" na listi `KORISNICI` -> `modMaticniUnos.MatIzmeni:187` -> `modMaticniKorisnici.KorIzmeni:295` | izmenjen red + toast "Izmenjeno." | Prazan PIN pri IZMENI znaci "ostaje isti", a pri UNOSU ne prolazi (`modMaticniKorisnici.bas:254-256`); sopstveni ID se cita IZ REDA, inace bi provera jedinstvenosti korisnickog imena pala na sopstvenom zapisu pri svakoj izmeni (`modMaticniKorisnici.bas:291-293`) | pise `tblKorisnici` | `modMaticniKorisnici.KorIzmeni:295`, `Proveri:465` | ne | `modTest.T_MatKor_RecnikDaNeIPrava:15794` -- "prazan PIN je 'ostaje isti' pri izmeni, ali NE prolazi pri unosu" (`modTest.bas:15795-15796`) |
| F-030 | administrator | Ugasiti ili vratiti nalog tako da to stvarno zaustavi prijavu | radnja `status` na listi `KORISNICI` -> `modMaticniUnos.MatPromeniStatus:255` -> `modMaticniKorisnici.KorPromeniStatus:336` | upisano "DA" ili "NE" u `Aktivan` + toast "Status promenjen u:" | Kolona `Aktivan` u `tblKorisnici` NIJE obicna kolona statusa: `modAuth` neaktivnim smatra samo "NE", pa bi opsti upis napisao "Neaktivan" i korisnik bi se i dalje prijavljivao -- zato ide kroz svog pisca (`modMaticniUnos.bas:265-271`) | pise `tblKorisnici.Aktivan` | `modMaticniKorisnici.KorPromeniStatus:336`, `DaNe:561` | ne | `modTest.T_MatKor_RecnikDaNeIPrava:15794` -- "pisac korisnika NE upisuje ono sto modAuth ne prepoznaje", "recnik kolone Aktivan je DA" / "je NE" |
| F-031 | administrator | Videti prava izabranog korisnika po oblastima, i odakle svako pravo dolazi | prekidac na listu `PRAVA` (`modMaticniIzvor.bas:121`), red se cita za korisnika izabranog u listi `KORISNICI` -> `modMaticniKorisnici.KorPravaRedovi:176` | mreza: Oblast, Ima/Nema, Odakle (admin ili pojedinacno), + skrivena kolona kljuca (`KorPravaKolone:166`) | Lista je po JEDNOJ oblasti iz `modAuth.OblastiList:319` (dvanaest oblasti); adminu se sve prikazuje kao "DA" jer to nije zapis nego pravilo (`modMaticniKorisnici.bas:191-194`); bez izabranog korisnika lista je prazna -- prava bez korisnika nisu podatak (`:181-185`); napomena u zoni kaze CIJA su prava na ekranu (`modMaticniEkran.Napomena:452`) | cita `tblKorisnici` (kolone oblasti, `Uloga`) | `modMaticniKorisnici.KorPravaRedovi:176`, `KorOblastNaziv:224`, `VrednostOblasti:594` | ne | `modTest.T_MatKor_RecnikDaNeIPrava:15794` -- "lista prava bira red po SKRIVENOJ koloni, a ne po vidljivom nazivu oblasti" (`modTest.bas:15801-15803`) |
| F-032 | administrator | Ukljuciti ili iskljuciti jedno pravo izabranog korisnika jednim potezom | dvoklik ili radnja `pravo` na listi `PRAVA` (`modMaticniEkran.bas:718`, `:567`) -> `PromeniPravo:799` -> `modMaticniKorisnici.KorPromeniPravo:362` | upisano "DA"/"NE" u kolonu oblasti + toast "<Oblast>: ima pravo / nema pravo" | Prava se NE uredjuju u editoru -- red je oblast, a ne zapis, pa sekcija nema ni "Izmeni" ni "Deaktiviraj" (`modMaticniEkran.Radnje:501-505`); adminu se promena ODBIJA sa `MATK_ERR_ADMIN_SVE` jer bi bila prividna -- sledeci upis bi je vratio na DA (`modMaticniKorisnici.bas:385-388`); bez izabranog korisnika se to i kaze umesto da dugme tiho ne radi (`modMaticniEkran.bas:801-803`); nepoznata oblast se odbija (`MATK_ERR_NEMA_OBLASTI`) | pise `tblKorisnici.<Oblast>` | `modMaticniKorisnici.KorPromeniPravo:362` | ne | `modTest.T_Maticni_KapijeUpisaIZivotniCiklus:16247` -- pisac prava nosi kapiju sam (`modTest.bas:16329-16340`) |

## F1e -- prijava i zamena operatera

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-033 | operater | Prijaviti se korisnickim imenom i PIN-om pri pokretanju aplikacije | `modMain.StartApp:96-102` -> `modAuth.Login:74` -> `modUiFaze.FazaPrijava:187`; polja `fzlUser` / `fzlPin`, dugme `fzlOK` (`modUiFaze.FazaKlik:368`) | kartica prijave kao FAZA iste ljuske; po uspehu ekran, po neuspehu `Poruka("AUTH_MSG_PRIJAVA_NEUSPESNA")` = "Prijava neuspesna. Aplikacija se zatvara." | Tri promasaja gase aplikaciju: brojac `MAX_ATT` (`modUiFaze.Prijavi:395-399`), pa `StartApp` zakazuje `QuitAfterFailedLogin` na sledeci tik (`modMain.bas:98-100`, `modAuth.QuitAfterFailedLogin:345`); posle svakog promasaja poruka nosi i redni broj pokusaja (`modUiFaze.bas:401-402`); dok traje prijava ljuska se NE gradi, jer bi dobila prava prazne sesije i zapamtila ih | cita `tblKorisnici` (`Username`, `PIN`, `Aktivan`, `Uloga`, `ImePrezime`) | `modAuth.Login:74`, `PrikaziPrijavu:114`, `modUiFaze.FazaPrijava:187`, `Prijavi:380`, `modAuth.ValidateLogin:207` | ne | `modTest.T_Faza_PrijavaNeGradiLjusku:6405` -- prijava je faza ljuske i dok traje ljuska nije izgradjena |
| F-034 | operater | Raditi bez prijave dok administrator ne ukljuci autentikaciju | kljuc `AUTH_ENABLED` u konfiguraciji (`modAuth.AuthEnabled:52`) | aplikacija se otvara pravo na ekran, bez kartice prijave | Prijava je opt-in: prihvata se samo `YES` / `DA` / `TRUE` / `1` (`modAuth.bas:57-59`); dok je iskljucena `KorisnikImaPravo` vraca True za svaku oblast (`modAuth.bas:285-288`), a `MozeAdministraciju` je anti-lockout -- svi su admini, da bi se nalozi mogli pripremiti PRE ukljucenja (`modAuth.bas:310-315`) | cita konfiguraciju `AUTH_ENABLED` | `modAuth.AuthEnabled:52`, `KorisnikImaPravo:282`, `MozeAdministraciju:314` | ne | `modTest.T_Ljuska_SuzenaPravaStartIAlatke:1117` -- "posle testa alatke opet prolaze" (mereno kroz `AuthTestUkljuci`) |
| F-035 | operater | Zameniti operatera bez gasenja aplikacije, tako da se sa imenom promene i prava | dugme `btnOperater` u zaglavlju ljuske (`modOtkupUI.bas:4213`) -> `DoSwitchOperater:4576` -> `modAuth.Login:74` | toast `Poruka("OTKUI_MSG_OPERATER")` sa novim imenom; radna povrsina precrtana po novim pravima | "Otkazi" znaci "ne menjam operatera": prethodna sesija se PAMTI pa tek onda gasi, i vraca se ako prijava ne prodje (`modAuth.bas:88-100`); tada poruka `OTKUI_MSG_OPERATER_ISTI` = "Operater nije promenjen, ostaje"; ako sesija NIJE vracena, prikaz ne sme da tvrdi da je stari operater tu -- prava se primenjuju na praznu sesiju i povrsina se isprazni (`modOtkupUI.bas:4592-4599`); zamena OVDE ne gasi aplikaciju posle tri promasaja, za razliku od starta | cita `tblKorisnici` | `modOtkupUI.DoSwitchOperater:4576`, `modAuth.Login:74`, `VratiSesiju:125`, `modOtkupUI.PrimeniNovaPrava:4610` | ne | `modTest.T_Auth_OtkazanaPrijavaNeLazePrikaz:15074` -- "otkazana prijava vraca PRETHODNOG operatera, ne prazno", "prazna radna povrsina BRISE redove, ne samo sto ih sakriva" |
| F-036 | operater | Videti samo one ekrane i alatke na koje prijavljeni nalog ima pravo | `modUiScreens.ScrDozvoljen` po koloni `SCR_OBLAST`; alatke `btnExcel` i `btnSync` kroz `modOtkupUI.AlatkaOblast:4471` i `AlatkaSme:4479` | prigusene stavke sidebara; zabranjena alatka javlja `Poruka("OTKUI_SCR_ZABRANJEN")` | Admin ima sve, bez obzira na kolone (`modAuth.KorisnikImaPravo:293-296`); pravo se priznaje za "DA"/"YES"/"TRUE"/"1"/"X" (`modAuth.bas:303`); provera je FAIL-CLOSED -- greska znaci NE SME (`modOtkupUI.AlatkaSme:4485-4488`, `modAuth.bas:305-307`); "Otvori Excel" i "Sinhronizuj" nisu ekrani ali JESU oblasti prava, pa se pitaju pre radnje (`modOtkupUI.bas:4461-4466`) | cita `tblKorisnici` (kolone oblasti iz `modAuth.OblastiList:319`) | `modAuth.KorisnikImaPravo:282`, `modOtkupUI.AlatkaSme:4479` | ne | `modTest.T_Ljuska_AlatkeTrazePravo:1063`; `modTest.T_Ljuska_StartEkranDozvoljen:1092`; `modTest.T_Ljuska_SuzenaPravaStartIAlatke:1117` -- "Excel alatka NE sme bez prava na oblast", "start vodi na prvi dozvoljen ekran" |
| F-037 | administrator | Biti siguran da ugasen nalog stvarno ne moze da udje, i da se to razlikuje od "nema prava" | `modAuth.ValidateLogin:207`, provera kolone `Aktivan` (`modAuth.bas:226-231`) | prijava odbijena; monitoring dobija `AUTH_LOGIN_FAIL` sa razlogom "(deaktiviran)" | Prazno znaci AKTIVAN, blokira samo "NE" -- drift-safe (`modAuth.bas:225-226`); nepoznat korisnik i pogresan PIN daju svoje razloge ("(nepoznat)", "(pogresan PIN)"), pa se u telemetriji razlikuju (`modAuth.bas:218-238`) | cita `tblKorisnici.Aktivan` | `modAuth.ValidateLogin:207`, `AuditAuth:486` | ne | `modTest.T_Ljuska_SuzenaPravaStartIAlatke:1117` -- "deaktiviran nalog se ne prijavljuje" (`modTest.bas:1151`) |
| F-038 | administrator | Imati PIN-ove cuvane hesovano, bez rucne migracije zatecenih | `modAuth.PinHashEnabled:358`, `PreparePin:405`, `VerifyPin:419`, migracija pri prijavi (`MigratePinToHash:464`) | zapis oblika `sha256$<salt>$<hex>` u koloni PIN | Podrazumevano UKLJUCENO (opt-out); prijava zatecenim plaintext PIN-om i dalje prolazi, a zapis se TIHO prevede u hes pri prvoj uspesnoj prijavi (`modAuth.bas:240-245`); ako SHA nije dostupan sve pada nazad na plaintext, bez lockout-a (`modAuth.bas:352-356`) | cita/pise `tblKorisnici.PIN` | `modAuth.VerifyPin:419`, `MigratePinToHash:464`, `HashPin:400`, `Sha256Hex:371` | ne | NEPROVERENO -- provera je makro `modAuth.TestPinHash:442` (Alt+F8), nije nadjena tvrdnja u automatskoj suiti |
| F-039 | odrzavanje | Imati svaku prijavu, neuspeh i otkaz zabelezen u telemetriji | `modAuth.AuditAuth:486`, zvana iz `ValidateLogin:207` i `VratiSesiju:125` | dogadjaji `AUTH_LOGIN`, `AUTH_LOGIN_FAIL`, `AUTH_LOGIN_OTKAZ` sa korisnikom i razlogom | Audit je fail-soft (`On Error Resume Next`, `modAuth.bas:487`) -- telemetrija ne sme da obori prijavu; otkazana zamena operatera se belezi izricito, kao "(prijava nije promenjena)" (`modAuth.bas:135`) | salje monitoring dogadjaje (`entityType:="Auth"`) | `modAuth.AuditAuth:486` | ne | NEPROVERENO -- slanje trazi mrezu; nije nadjen test koji tvrdi bas ove dogadjaje |

## F1f -- zivotni ciklus sveske (otvaranje, snimanje, zatvaranje)

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-040 | operater | Otvoriti fajl i odmah dobiti aplikaciju, bez ijednog pogleda na sirove listove | dvoklik na svesku -> `ThisWorkbook.Workbook_Open:10` -> `modMain.StartApp:16` | skriven Excel, splash sa statusom ("Licenca", "Verzija", "Ekran"), pa ljuska sa ekranom | `Application.Visible = False` je PRVA naredba, pre svake mrezne provere (`ThisWorkbook.doccls:20-21`); red je: sakrij -> splash -> kapije -> prijava -> ekran (`modMain.bas:45-50`); sveska se otkriva SAMO na tri mesta i svako od njih to zeli -- odbijena kapija, first-run setup i dugme "Otvori Excel"; splash se dopunjava do najmanjeg trajanja da znak ne bljesne kad su sve kapije prosle trenutno (`modMain.bas:183-186`) | -- | `ThisWorkbook.Workbook_Open:10`, `modMain.StartApp:16`, `modUiFaze.FazaBoot:132`, `FazaStatus:169`, `FazaBootSacekaj:152` | ne | NEPROVERENO -- ceo lanac trazi Excel (`Application.Visible`, `OnTime`) |
| F-041 | odrzavanje | Zaustaviti pokretanje na masini bez vazece licence ili sa isteklim probnim periodom, a ponuditi unos kljuca umesto tihog gasenja | `modMain.StartApp:60` -> `modLicense.AccessGateOrQuit:92` | dijalog za kljuc ili blokada sa zakazanim zatvaranjem; `modLicense.AccessWasDenied:626` prekida start odmah | Opt-in: ni licenca ni trial ukljuceni -> propusta (`modLicense.bas:120-122`); masina sa kljucem ide na licencu, bez kljuca a u trialu na trial; trial istekao a licenca ukljucena -> kljuc se trazi INLINE, jer bi se masina inace zatvorila pre nego sto se stigne do makroa (`modLicense.bas:105-111`); bug u orkestraciji NE sme da zakljuca korisnika -- fail-OPEN uz log (`modLicense.bas:126-129`) | cita konfiguraciju (`CFG_LIC_KEY`) | `modLicense.AccessGateOrQuit:92`, `AccessWasDenied:626`, `ThisWorkbook.doccls:27` | ne | NEPROVERENO -- nema testa koji tvrdi ishod kapije; `modLicenseTests.bas:46-105` pokriva samo razlaganje kljuca, poklapanje delova i otisak uredjaja |
| F-042 | odrzavanje | Upozoriti ili zaustaviti zastarelog klijenta u floti, po odluci servera | `modMain.StartApp:88` -> `modUpdateGate.UpdateGateOrQuit:35` | `MsgBox` sa porukom servera; uz `enforce` blokada i zakazano zatvaranje, inace samo preporuka | Fail-open na svakom koraku: nema endpointa ili tajne -> ne diramo, offline -> propusti, server bez policy-ja -> propusti (`modUpdateGate.bas:42-48`); podrazumevano je WARN, blokira se samo na izricito `YES`/`TRUE`/`1`/`ON` (`modUpdateGate.bas:60-69`); kapija stoji POSLE ponude self-update-a, da `enforce` ne bi blokirao bas onog klijenta kome je azuriranje najpotrebnije (`modMain.bas:72-76`) | cita konfiguraciju (`MONITORING_ENDPOINT`, `MONITORING_SECRET`), cita odgovor GAS-a | `modUpdateGate.UpdateGateOrQuit:35`, `VersionHttpCheck:87` | ne | NEPROVERENO -- trazi mrezu; nije nadjen test koji tvrdi ishod kapije |
| F-043 | odrzavanje | Dobiti ponudu za podesavanje racunara pri prvom pokretanju, i ne dobijati je vise kad prodje | `modMain.StartApp:114-120`, kljuc `APP_SETUP_COMPLETED` u `tblLocalConfig` | `MsgBox` `Poruka("SETUP_MSG_FIRSTRUN_PONUDA")` = "Ovaj racunar jos nije podesen. Zelite li da pokrenete podesavanje (SetupNewPC) sada?" | Jednokratno -- cim `SetupNewPC` prodje i upise "DA", kapija se vise ne javlja; fail-soft, ne obara start (`modMain.bas:108-112`); ovo je JEDINO mesto u startu koje stvarno trazi vidljiv Excel, jer setup bira foldere kroz `FileDialog` -- zato se sveska otkriva samo za njega, a splash se sklanja da ne pokriva dijaloge (`modMain.bas:116-119`). Sam `SetupNewPC` je oblast F2 | cita `tblLocalConfig.APP_SETUP_COMPLETED` | `modMain.StartApp:114`, `modSetup.SetupNewPC:41` | ne | NEPROVERENO -- trazi Excel i `FileDialog` |
| F-044 | odrzavanje | Imati svesku cija se sema sama poravna sa kanonom pri svakom pokretanju, a odstupanje koje se ne moze izleciti prijavljeno | `modMain.StartApp:136-158` -> `modSchema.SchemaCheckOnStart`, pa `EnsureAllTables` | log `WARN` "SEMA: ..." + monitoring `SCHEMA_DRIFT` | Provera je jeftina (otisak zaglavlja), pun `VerifySchema` se placa tek kad se otisci razlikuju (`modMain.bas:126-129`); neslaganje se PRVO pokusava izleciti pa tek onda prijavljuje (`modMain.bas:139-146`); REDOSLED se ne leci -- premestanje bi pomerilo podatke, pa poruka posle drugog prolaza znaci "mora rucno"; fail-SOFT je namerno: pogresna sema ne sme da zakljuca aplikaciju, tvrda kapija stoji pred samim upisom (`modSchema.SchemaReadyOrFail`) | cita zaglavlja svih tabela, poredi sa `schema/schema.json` | `modMain.StartApp:136`, `modSchema.SchemaCheckOnStart`, `EnsureAllTables` | ne | NEPROVERENO -- trazi Excel; posredno je pokriveno kapijom `python tools/gen_schema_module.py --check` |
| F-045 | odrzavanje | Imati ugovor o formatu celije primenjen pri svakom pokretanju, a neuspelu primenu prijavljenu umesto progutanu | `modMain.StartApp:171-183` -> `modSchema.PrimeniFormateKanona` | log `WARN` "FORMAT: ..." + monitoring `SCHEMA_DRIFT` | Kolona bez ugovora TIHO menja vrednost pri upisu ("3/2026" -> datum, vodeca nula otpadne), a steta se ne vidi ni u jednoj kasnijoj proveri; primena je fail-soft, ali NIJE tiha -- bez ovog reda bi neuspela primena prosla bez ijednog traga (`modMain.bas:166-169`) | postavlja formate kolona po kanonu | `modMain.StartApp:171`, `modSchema.PrimeniFormateKanona` | ne | `modTest.T_Sema_FormatCelijeCuvaVrednost:6588`; `modTest.T_Sema_ZurnalCuvaVrednostKrozJournalCell:6661` |
| F-046 | odrzavanje | Dobiti kopiju sveske pri svakom pokretanju i ocisceno skladiste starih kopija, zurnala i logova | `modMain.StartApp:194-208` -> `BackupFileOnStart`, `PurgeOldBackups`, `PurgeOldJournals`, `PurgeOldLogs`, `LogAppStart` | backup fajl; toast `Poruka("APP_MSG_BACKUP_NIJE_USPEO")` kad ne prodje | Backup je sigurnosna mreza, NE preduslov rada: do 12.09.2026. je pun disk obarao celo pokretanje i operater je ostajao bez alata (`modMain.bas:188-193`); sada je fail-soft ali ne i tih -- ljuska je vec podignuta, pa se izostanak zastite kaze toast-om | pise backup fajl, brise stare fajlove; pise log start | `modMain.StartApp:194`, `BackupFileOnStart`, `PurgeOldBackups`, `PurgeOldJournals`, `PurgeOldLogs` | ne | NEPROVERENO -- trazi disk i Excel |
| F-047 | odrzavanje | Biti upozoren pri pokretanju da je u zurnalu ostalo nesto sto moze znaciti gubitak podataka | `modMain.StartApp:215-236` -> `CheckJournalForRecovery` | `MsgBox` "UPOZORENJE - Moguc gubitak podataka!" sa sadrzajem nalaza + monitoring `JOURNAL_RECOVERY_WARN` | Poruka imenuje folder zurnala i trazi reimport ako je potrebno (`modMain.bas:234-236`); ne blokira start | cita folder zurnala | `modMain.StartApp:215`, `CheckJournalForRecovery` | ne | NEPROVERENO -- trazi disk; nije nadjen test koji tvrdi ovaj nalaz |
| F-048 | operater | Snimiti radnu svesku dugmetom u zaglavlju ljuske | dugme `btnSnimi` (`modOtkupUI.bas:4211`) -> `DoSaveWorkbook:4537` | toast `Poruka("OTKUI_MSG_WB_SNIMLJENA")` = "Radna sveska je snimljena."; pad javlja `OTKUI_MSG_WB_PALO` | Dok stoji marker prekinutog VBA uvoza snimanje se ODBIJA glasno -- toast kaze i zasto i sta da se uradi (`modOtkupUI.bas:4538-4544`); `SaveApp` (javan Save bez pozivaoca) je obrisan 12.09.2026. bas zato sto je bio zaobilaznica te kapije (`modMain.bas:410-416`) | pise fajl sveske | `modOtkupUI.DoSaveWorkbook:4537`, `modImportState.ImportNijeDovrsen` | ne | `modTest.T_Save_PrekinutImportZatvaraSvaVrata:7203` |
| F-049 | odrzavanje | Imati snimanje odbijeno na SVIM vratima dok je projekat mozda nepotpun, ukljucujuci Ctrl+S i File > Save | `ThisWorkbook.Workbook_BeforeSave:93` | `Cancel = True`, log `WARN` "Save ODBIJEN: prekinut VBA import...", i `MsgBox` `Poruka("APP_MSG_IMPORT_PREKINUT_NE_SNIMAM")` van test rezima | Tri VBA puta (AutoSave, dugme "Snimi", gasenje) NISU sva vrata -- Ctrl+S, File > Save i Save As idu mimo svakog od njih; 12.09.2026. je tako polomljen `modLogo` prezive- o zatvaranje i ponovno otvaranje fajla (`ThisWorkbook.doccls:76-81`); self-update je izuzet PO KONSTRUKCIJI: on gasi `Application.EnableEvents` pre ijedne izmene koda, pa ovaj event tada ne opali, a to i `tools/vba_selfupdate_gates.py` proverava kao kapiju (`ThisWorkbook.doccls:83-89`) | -- | `ThisWorkbook.Workbook_BeforeSave:93`, `modImportState.ImportNijeDovrsen` | ne | `modTest.T_Save_PrekinutImportZatvaraSvaVrata:7203` -- "sa markerom je i direktan Save (Ctrl+S put) odbijen", "bez markera direktan Save i dalje prolazi (kapija nije 'uvek otkazi')" |
| F-050 | operater | Zatvoriti aplikaciju tako da se unos snimi, a zauzeto otkupno mesto oslobodi | X na formi -> `modOtkupUI.OtkupUI_Sakrij:5391` -> `modMain.ZatvoriAplikaciju:395`; i `ThisWorkbook.Workbook_BeforeClose:108` kao poslednja brana | zatvorena sveska; pri zauzetom otkupnom mestu statusna traka "Sinhronizujem unos pre zatvaranja..." | Normalno zatvaranje snima (`SaveChanges:=True`, `modMain.bas:407`); pri prekinutom uvozu se zatvara BEZ snimanja i operater dobija razlog -- nesnimljen red se moze uneti ponovo, polomljen projekat prezivi zatvaranje fajla (`modMain.bas:400-405`); oslobadjanje locka je fail-soft, a bulk push se ne radi bezuslovno jer bi pri odbacenim promenama sledeca sesija poslala iste redove PONOVO (`ThisWorkbook.doccls:112-119`). Ciscenje zaostalih lockova je **E-029**, gasenje zakazanog sync-a **E-009** | pise fajl sveske; posredno Google `SyncControl` | `modMain.ZatvoriAplikaciju:395`, `ShutdownApp:339`, `ThisWorkbook.Workbook_BeforeClose:108` | ne | `modTest.T_Save_PrekinutImportZatvaraSvaVrata:7203` -- "izlazak pri prekinutom importu NE push-uje u cloud (SaveChanges:=False)", "normalan izlazak i dalje radi bulk push" |
| F-051 | operater | Doci do sirove radne sveske i vratiti se u aplikaciju jednim klikom, bez ponovnog gradjenja ekrana | dugme `btnExcel` (`modOtkupUI.bas:4212`) -> `DoShowExcel:4557` -> `modUiFaze.FazaMini:255`; povratak dugmetom `fzmBack` (`modUiFaze.FazaKlik:370`) | vidljiv Excel + plutajuca kartica "Nazad u aplikaciju"; povratak vraca pun ekran | Iza dugmeta je sirova sveska sa svim tabelama, pa se PRE radnje pita za pravo `OtvoriExcel` (`modOtkupUI.bas:4562-4566`); kartica je FAZA iste forme, ne zasebna forma -- ljuska ostaje izgradjena pa je povratak trenutan (`modOtkupUI.bas:4567-4570`); povratak sam vraca `Application.Visible = False` (`modUiFaze.NazadUAplikaciju:296`) | -- | `modOtkupUI.DoShowExcel:4557`, `modUiFaze.FazaMini:255`, `FazaApp:270` | ne | `modTest.T_Ljuska_AlatkeTrazePravo:1063`; `modTest.T_Ljuska_SuzenaPravaStartIAlatke:1117` -- "Excel alatka NE sme bez prava na oblast" |
| F-052 | odrzavanje | Otkriti ili sakriti Excel prozor makroom, kad ljuska nije upotrebljiva | Alt+F8 -> `modMain.OpenExcel:382` / `modMain.CloseExcel:386` | vidljiv, odnosno skriven Excel | Oba su `Public Sub` bez argumenata i bez ijednog pozivaoca u kodu (`popis_citalaca`: `status=ZIV_MAKRO`, `prod_pozivaoci=0`), dakle iskljucivo Alt+F8 put; NE pitaju za pravo `OtvoriExcel`, za razliku od dugmeta **F-051** | -- | `modMain.OpenExcel:382`, `modMain.CloseExcel:386` | ne | NEPROVERENO -- trazi Excel; nije nadjen test koji ih zove |

## Pokrivenost ulaznih tacaka

### Registar ekrana, sekcija MATICNI (`modUiScreens.bas:161-183`)

| Ulazna tacka | ID |
|---|---|
| `MAT_PARTNERI` (`:161`) | F-001, F-010, F-011, F-012, F-013, F-014 |
| `MAT_ROBA` (`:163`) | F-001, F-015, F-016, F-017, F-019 |
| `MAT_PAKOVANJE` (`:165`) | F-001, F-018 |
| `MAT_KORISNICI` (`:171`) | F-001, F-027, F-028, F-029, F-030, F-031, F-032 |
| `MAT_PODESAVANJA` (`:180`), `MAT_ADMIN` (`:182`) | oblast F2 -- paneli, ne ekrani; nisu mapirani u ovoj sesiji |

### Motor maticnih ekrana (`modMaticniEkran`)

| Ulazna tacka | ID |
|---|---|
| prekidac lista `ls<KLJUC>` (`:540`) | F-004 |
| izbor reda `row:<n>` (`:555`) | tehnicko (pamti izbor, ne cita listu ponovo) |
| dvoklik `dbl:<n>` (`:564`) | F-006, F-032 |
| radnja `act:izmeni` (`:715`) | F-006 |
| radnja `act:status` (`:716`) | F-007 |
| radnja `act:geo` (`:717`) | F-020 (otvara panel) |
| radnja `act:pravo` (`:718`) | F-032 |
| promena polja `chg:<kontrola>` (`:583`) | F-008 |
| `scrMatNovi` (`:593`) | F-005 |
| `scrMatOdustani` (`:594`) | tehnicko (zatvara editor) |
| `scrMatSacuvaj` (`:595`) | F-005, F-006, F-028, F-029 |
| `scrGeoZatvori` (`:596`) | tehnicko |
| `scrGeoNalepi` (`:597`) | F-023 |
| `scrGeoPortal` (`:598`) | F-020 |
| `scrGeoMape` (`:599`) | F-021 |
| `scrGeoPoligon` (`:600`) | F-022 |
| `scrGeoSacuvaj` (`:601`) | F-024 |
| `scrGeoObrisi` (`:602`) | F-025 |
| traka statusa GEO panela (`RasporediGeo:216`) | F-026 |

### Ljuska -- zaglavlje i zivotni ciklus

| Ulazna tacka | ID |
|---|---|
| `btnOperater` (`modOtkupUI.bas:4213`) | F-035 |
| `btnSnimi` (`modOtkupUI.bas:4211`) | F-048 |
| `btnExcel` (`modOtkupUI.bas:4212`) | F-051 |
| `btnSync` (`modOtkupUI.bas:4210`) | oblast E (**E-008** i dalje); brana alatke je F-036 |
| `fzlOK` / `fzlCancel` (`modUiFaze.FazaKlik:368-369`) | F-033, F-035 |
| `fzmBack` (`modUiFaze.FazaKlik:370`) | F-051 |
| `ThisWorkbook.Workbook_Open:10` | F-040 (+ F-041, F-042, F-043, F-044, F-045, F-046, F-047) |
| `ThisWorkbook.Workbook_BeforeSave:93` | F-049 |
| `ThisWorkbook.Workbook_BeforeClose:108` | F-050 |

### Makroi (Alt+F8) iz ovog dela oblasti

| Makro | ID |
|---|---|
| `modMain.StartApp:16` | F-040 (javan je i rucno se pokrece; isti lanac) |
| `modMain.OpenExcel:382`, `modMain.CloseExcel:386` | F-052 |
| `modAuth.TestPinHash:442` | tehnicko -- provera pred ukljucenje hesa (`modAuth.bas:352-356`), ne daje ishod operateru |
| `modMain.InitApp:300`, `modMain.ShutdownApp:339` | tehnicko -- inicijalizacija i gasenje, zovu ih `StartApp` i `Workbook_BeforeClose` |

### Mrtvo / tehnicko (NIJE sposobnost)

- **`modMaticniLookups` je ceo mrtav osim jedne procedure.** `MaticniSekcije:33` ima
  jedinog pozivaoca u testu (`popis_citalaca`: `status=SAMO_TEST`,
  `prod_pozivaoci=0`), a `MaticniSekcijeGrupisano:66` zove samo ona. Modul je
  gradio meni forme `frmMaticniPodaci`, koje u `src-vba/` vise nema -- jedina
  forma je `frmOtkupUI`. Ziva je samo `MaticniMenu_Release:93`
  (`status=ZIV_UI`), i to kao tehnicki korak: otpusta `WithEvents` reference pre
  VBA importa i self-update-a (`modVbaTools.PrepareRuntimeForImport:1594`,
  `modSelfUpdate.PrepareRuntimeForSelfUpdate:424`).
- **`modGeoParcele`: tri mrtve procedure.** `SaveParcelGeoPoint:4`,
  `ClearParcelGeo:99` (obe rade po REDNOM BROJU reda, sto je legacy nacin
  biranja) i `SyncSelectedParcelaToGoogle:302` -- sve tri `status=MRTAV`,
  `prod_pozivaoci=0`, `test_pozivaoci=0`. Ziv put ide po ID-u
  (`SaveParcelGeoPointByID:13`, `ClearParcelGeoByID:108`).
- **Test seam-ovi** (ne daju ishod korisniku, a samo suzavaju sta prolazi):
  `modScrMatPartneri.Scr_MpTestSet:84`, `modScrMatRoba.Scr_MrTestSet:81`,
  `modScrMatPakovanje.Scr_MkTestSet:81`, `modScrMatKorisnici.Scr_MkorTestSet:99`
  i `Scr_MkorBranaZatvoriTest:104`, `modMaticniUnos.MatBranaZatvoriTest:54`,
  `modMaticniIzvor.MatComboPadTest:610`, `modAuth.AuthTestUkljuci:46`,
  `modUiFaze.FazaGradiTest:861` / `FazaRezimTest:875` / `FazaPokusajiTest:879` /
  `FazaCekaTest:883` / `FazaIshodTest:887`, i `Mat*Test` funkcije
  `modMaticniEkran.bas:1088-1126`.
- **`modAuth.AuthRegresijaOtkaz:162`** je javan, ali nije sposobnost: vozi pravi
  `Login` i vraca nalaz za test `modTest.T_Auth_OtkazanaPrijavaNeLazePrikaz:15074`.
  Namerno NIJE setter sesije -- postavljanje sesije ne postoji ni za test
  (`modAuth.bas:145-150`).
- **`modUiScreens` ugovorne procedure** maticnih ekrana (`Scr_Build`,
  `Scr_Layout`, `Scr_Deaktiviraj`, `Scr_ResetCache`) su tehnicke: prosledjuju
  posao motoru i ljusci.

## Sposobnosti koje postoje samo u PWA / GAS

Za deo F1 -- nijedna nadjena. Maticni ekrani, prijava i zivotni ciklus sveske su
u celini VBA. Ono sto PWA radi sa istim tabelama (citanje sifarnika kroz
`Stammdaten` sync) je vec mapirano u oblasti E.

## NEPROVERENO (sa razlogom)

| ID | Sta nije provereno | Razlog |
|---|---|---|
| F-004 | pitanje `MATU_ASK_ODBACI_UNOS` pri promeni liste i ekrana | nema testa koji tvrdi bas to pitanje |
| F-012 | sekcija `KUPCI` (polja, alias `Drzava`, izvedena adresa) | nema testa bas za tu sekciju -- pokrivena je samo opstim prolazom kroz sve sekcije |
| F-013 | sekcija `VOZACI` | nema testa bas za tu sekciju |
| F-018 | koje od cetiri tabele pakovanja ZAISTA imaju kolonu `Aktivan` u zatecenoj svesci | `MatStatusKolona:207` ne cita kanon nego radi probe nad otvorenom sveskom (`GetColumnIndex`), sto trazi Excel; po kanonu je imaju sve cetiri |
| F-019 | da li je `RokMeseci` negde drugde upisiv | nije pretrazeno van maticnih ekrana u ovoj sesiji |
| F-020 | otvaranje GeoSrbija portala i punjenje klipborda | `ThisWorkbook.FollowHyperlink` i `CopyToClipboard` traze Excel |
| F-022 | otvaranje poligona | `FollowHyperlink` trazi Excel i pregledac |
| F-024 | upis geo tacke (N, E, Lat, Lng, status, izvor, meteo, datumi) | upis trazi Excel; testom je pokriven samo racun `ConvertUTM34ToLatLng` |
| F-025 | brisanje geo podatka | trazi Excel |
| F-026 | traka statusa geo panela | crta se iz rasporeda forme, trazi Excel |
| F-033 | ponasanje kartice prijave (fokus, brojac na ekranu, gasenje posle tri promasaja) | `FazaPrijava` se vrti na `DoEvents` dok neko ne klikne dugme; `T_Faza_PrijavaNeGradiLjusku:6405` meri ugovor faze, ne prikaz |
| F-038 | hesovanje i migracija PIN-a | provera je makro `modAuth.TestPinHash:442`, nije nadjena tvrdnja u automatskoj suiti |
| F-039 | audit dogadjaji prijave | slanje trazi mrezu; nema testa koji tvrdi bas te dogadjaje |
| F-040 | ceo start lanac (skrivanje, splash, redosled kapija) | trazi Excel (`Application.Visible`, `Application.OnTime`) |
| F-041 | ishod licencne i trial kapije | nema testa koji to tvrdi; `modLicenseTests.bas:46-105` pokriva samo razlaganje kljuca, poklapanje delova i otisak uredjaja |
| F-042 | ishod min-version kapije | trazi mrezu; nema testa koji tvrdi ishod |
| F-043 | first-run setup ponuda | trazi Excel i `FileDialog` |
| F-044 | provera i samolecenje seme na startu | trazi Excel (posredno pokriveno kapijom `gen_schema_module.py --check`) |
| F-046 | backup i ciscenje starih fajlova | trazi disk i Excel |
| F-047 | upozorenje iz zurnala | trazi disk; nema testa koji tvrdi nalaz |
| F-052 | makroi `OpenExcel` / `CloseExcel` | trazi Excel; nema testa koji ih zove |

## Nalaz uz mapu (za korak posle)

**1. Deo F1 je JEDINA oblast do sada bez ijedne zavisnosti od starog modela.**
Svih 52 sposobnosti nose presudu "ne". Maticni ekrani citaju i pisu iskljucivo
sifarnike (`tblKooperanti`, `tblStanice`, `tblKupci`, `tblVozaci`, `tblParcele`,
`tblArtikli`, `tblKulture`, `tblCenovnik`, `tblVrstaGP`, cetiri tabele
pakovanja, `tblKorisnici`); nijedan combo izvor ne cita dokument
(`modMaticniIzvor.MatComboStavke:632-643` -- svi izvori su sifarnici ili
zatvoreni spiskovi). Prijava cita samo `tblKorisnici`. Za redosled slajseva to
znaci da se F1 moze ostaviti netaknut kroz ceo refaktor dokumenata.

**2. Dva javna makroa zaobilaze branu koju dugme postuje.**
`modMain.OpenExcel:382` i `modMain.CloseExcel:386` rade tacno ono sto i dugme
`btnExcel`, ali bez provere prava `OtvoriExcel` koju `DoShowExcel:4562` radi pre
radnje. Isti obrazac je vec jednom naveden kao razlog brisanja `SaveApp`
(`modMain.bas:410-416`: "mrtav javan Save nema zasto da postoji"), i isti razlog
zbog kog kapija upisa stoji u piscu a ne samo u ekranu
(`modMaticniUnos.bas:60-64`). Nije regresija refaktora -- postojece stanje koje
treba presuditi.

**3. Komentar o kolonama statusa se razisao sa kanonom, i to menja koje
sekcije imaju "Deaktiviraj".**
`modMaticniIzvor.bas:203-206` tvrdi da "Cenovnik, Ambalaza i Palete" po
zatecenoj semi nemaju kolonu statusa. Po kanonu `schema/schema.json` je nemaju
SAMO `tblCenovnik`; `tblTipAmbalaze`, `tblTipPalete`, `tblKutije` i `tblKese`
sve nose `Aktivan`. Kod to ne cita iz kanona nego probe-om nad zatecenom
sveskom (`MatStatusKolona:207` -> `GetColumnIndex`), pa se stvarno ponasanje
razlikuje od sveske do sveske -- a komentar koji vodi citaoca je zastareo.
Uz to `tblVrstaGotovihProizvoda` ima kolonu `RokMeseci` koju ekran NE nudi
(`MatPolja:476` ima samo `tip`), pa se ta vrednost sa ekrana ne moze ni uneti ni
videti. Nov model treba da presudi oboje iz kanona, ne iz probe-a.

**4. Parcele pri unosu dobijaju status "Da", a pri deaktivaciji
"Aktivan"/"Neaktivan".** `MatStatusNaUnosu:577` to i kaze: dva oblika u istoj
koloni, a citac aktivnim smatra sve sto nije "Neaktivan". Poravnanje bi promenilo
ono sto sinhronizacija vec vidi, pa se svesno ne radi -- ali nova tabela parcela
ne treba da nasledi dva pisanja istog stanja.

**5. Rucno unesena geo tacka se belezi kao da je dosla iz automatike.**
`SaveParcelGeoPointByID_TX` upisuje `GeoSource = "selenium"` bez obzira odakle
je tacka dosla (`modGeoParcele.bas:78`), a jedini ziv pozivalac je rucni upis sa
GEO panela (`modMaticniGeo.GeoSacuvaj:136`). Traka statusa (F-026) taj izvor
prikazuje operateru, pa on cita "selenium" i za ono sto je sam otkucao.

**6. `PRAVA` su jedina lista koja radi bez zapisa.** Red je oblast iz
`modAuth.OblastiList:319`, ne zapis u tabeli; identitet nosi skrivena kolona
(`MatKolonaID:237` -> `KOR_COL_OBLAST`). Dodavanje oblasti prava je zato jedan
red u `OblastiList` plus jedan red u `modPoruke` (`KorOblastNaziv:224`) -- i
jedna NOVA KOLONA u `tblKorisnici`, jer se pravo cita po imenu kolone
(`modAuth.bas:302`). To je jedino mesto u F1 gde sema zavisi od spiska u kodu.

---

# Mapa sposobnosti -- oblast F, deo F2: ADMIN, PODESAVANJA, SETUP, AZURIRANJE, INTEGRITET, HEALTH, PREOSTALI MAKROI

Moduli: `modAdmin` (admin panel), `modPodesavanja` (config editor), `modSetup`
(setup nove masine, seme, prekidaci prijave), `modSelfUpdate` + `modUpdateGate`
(azuriranje), `modRelease` + `modE2EReleaseGate` (izdanja), `modLicense`
(aktivacija), `modIntegritet` (22 provere podataka), `modProductionHealthCheck`
(19 provera zdravlja), `modAutoHladnjaca` (backfill), `modPregledListova`
(sirova ljuska sveske), `modMigracija`, `modGoogleAuth`, `modDrive`.

Podoblasti: **F2a** admin panel - **F2b** podesavanja - **F2c** setup i rezimi
masine - **F2d** azuriranje, izdanja i licenca - **F2e** integritet i zdravlje -
**F2f** preostali makroi.

Numeracija se nastavlja na F1 (F-001..F-052).

Ne ponavlja se ono sto je vec mapirano: prikaz nalaza integriteta na ekranu
oporavka je **B-048** (sadrzaj 22 provere je ovde), zivi auto-lanac hladnjace je
**A-014** (`AutoChainHladnjaca`), zaustavljanje zakazanog sync-a pri self-update-u
je **E-009**, oslobadjanje stanica-locka pri self-update-u je **E-030**, licencna
i min-version kapija na startu su **F-041** i **F-042**, prva ponuda za setup na
startu je **F-043**.

KO: kao u F1 -- **operater** / **administrator** / **odrzavanje**, uz
**razvoj-release** za build masinu (VBA import/export, objava i rollback izdanja,
release kapija).

## F2a -- admin panel (`modAdmin`)

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-053 | administrator | Doci do admin panela iz sidebara, i biti odbijen ako nije administrator | sidebar, sekcija `SEK_MATICNI`, red `MAT_ADMIN` (`modUiScreens.bas:182`); registar panela `modUiPanel.PanelRedovi:71` | panel u radnoj povrsini ljuske, naslov `Poruka("OTKUI_MS_ADMIN")` | Red nema modul ekrana -- ljuska ga prepoznaje kao PANEL jer ga `modUiPanel` zna. Brana je dvostruka: sidebar prigusi stavku, a `BuildAdminPanel` sam odbija ne-administratora sa `Poruka("AUTH_MSG_SAMO_ADMIN_SEKCIJA")` (AUD-033) | cita `tblKorisnici` (prava, kroz `modAuth.MozeAdministraciju`) | `modAdmin.BuildAdminPanel:51`, `modUiPanel.PanelRedovi:57` | ne | `modTest.T_UiPanel_StavkaSidebara:14549` -- "registar panela nosi dva kljuca", "stavka sidebara otvara panel, ne ekran"; `T_UiPanel_UgovorIUstupanje:16063` -- "graditelj svakog panela postoji pod imenom iz registra" |
| F-054 | administrator | Naci komandu kroz pet grupa umesto kroz jedan naslagan spisak | prekidac grupa u panelu; akcija `grp:<naziv>` (`modAdmin.AdminPanel_OnClick:226`) | crta se SAMO aktivna grupa, komande u dve kolone | Dvanaest komandi u pet grupa: "Azuriranje", "Setup i provere", "Google / Drive", "VBA (dev)", "Podaci (oprezno)". Dodavanje komande = jedan red u `AdminGroups` | -- | `modAdmin.AdminGroups:168`, `RelayoutAdmin:370` | ne | NEPROVERENO |
| F-055 | administrator | Rucno proveriti ima li novije verzije i pokrenuti azuriranje, uz jasan odgovor i kad je nema | dugme "Proveri azuriranje" (akcija `checkupdate`, `modAdmin.bas:235`) | `MsgBox` sa daljinskom i tekucom verzijom, ili "Nema novih azuriranja" (+ napomena kad je kanal nedostupan) | Provera je read-only (`ReleaseManifestVersion` ne dira `modSelfUpdate`), pa je self-update-safe; samo "Da" pokrece azuriranje, i to preko `Application.OnTime` (prazan stack), nikad direktnim pozivom | cita `version.json` u Drive folderu `REL_FOLDER_ID` | `modAdmin.AdminCheckUpdate:264`, `modUpdateGate.ReleaseManifestVersion:179` | ne | NEPROVERENO |
| F-056 | odrzavanje | Jednim potezom napraviti i dopuniti sve tabele, kolone i kataloge koje aplikacija trazi | dugme "Ensure (setup + seme)" (akcija `ensure`, `modAdmin.bas:236`) | poruke pojedinacnih koraka + zavrsni rezime "Ensure zavrsen" | Agregat: `SetupNewPC` pa sve pojedinacne seme; sve su idempotentne, pa ponovljeno pokretanje ne kvari nista | pise `tblLocalConfig`, `tblPoruke`, `tblCenovnik`, `tblPaleta`, `tblKorisnici` + audit kolone; tabele iz `schema/schema.json` kroz `modSchema.EnsureAllTables` | `modAdmin.AdminEnsureEverything:289` -> `modSetup.SetupNewPC:41`, `EnsurePoruke:1120`, `EnsureCenovnikSchema:999`, `EnsurePaletniListSchema:919`, `EnsureDoradeSchema:1619`, `EnsureKorisniciSchema:1729`, `EnsureAuditColumns:1064` | ne -- spisak tabela i kolona dolazi iz kanona (`modSchema.EnsureAllTables:235`), pa se menja sa kanonom, ne ovde | NEPROVERENO |
| F-057 | odrzavanje | Autorizovati ovu svesku kod Google-a (OAuth) bez rucnog kopiranja tokena u tabelu | dugme "Google autorizacija" (akcija `googleauth`, `modAdmin.bas:240`) | otvoren pregledac sa Google prijavom, pa unos koda -> upisani tokeni | Odbija se odmah ako `GOOGLE_CLIENT_ID` ili `GOOGLE_CLIENT_SECRET` nisu upisani | cita/pise `tblSEFConfig` (`GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, tokeni) | `modGoogleAuth.RunGoogleAuthSetup:35` | ne | NEPROVERENO |
| F-058 | razvoj-release | Objaviti tekuci kod i `version.json` celoj floti, uz sifru i bez mogucnosti da se objavljena verzija tiho prepise | dugme "Objavi release na Drive" (akcija `publish`, `modAdmin.bas:241`); Alt+F8 `PublishReleaseToDrive` | fajlovi u `AgriX_Release\releases\<APP_VERSION>\` + `manifest.json` + `current.json` | Sifra (`RELEASE_PUBLISH_SIFRA`) je ujedno i potvrda; versioned folder se pravi PRE ijednog uploada (pad = prekid bez izmene kanala); verzija koja vec ima `manifest.json` je immutable -- re-objava je tvrdo odbijena osim uz `ALLOW_REPUBLISH` | cita `src-vba\`; pise Drive `REL_FOLDER_ID` | `modAdmin.AdminPublishToDrive:309`, `modRelease.PublishReleaseToDrive:26` | ne | NEPROVERENO |
| F-059 | razvoj-release | Usaglasiti VBA projekat sa `src-vba` folderom, ili ga izvesti nazad u fajlove | dugmad "VBA Import" / "VBA Export" (akcije `vbaimport` / `vbaexport`, `modAdmin.bas:242-243`); Alt+F8 `ImportAllVBA` / `ExportAllVBA`; dugme "Uvezi VBA" na listu `Pregled listova` (`modPregledListova.UveziVBA:123`) | prepisani moduli u projektu, odnosno `.bas`/`.cls`/`.frm` fajlovi na disku | Import prepisuje kod, pa nosi potvrdu; sam `ImportAllVBA` jos i pokazuje PLAN izmena, pravi backup pre prve izmene i odbija izvor koji nije pravi `src-vba` (mora imati `modVbaTools.bas`, tacno jednu `.frm` i njen `.frx`) | -- | `modAdmin.AdminVbaImport:328`, `modVbaTools.ImportAllVBA:235`, `ExportAllVBA:205` | ne | NEPROVERENO |
| F-060 | razvoj-release | Otvoriti VBA editor i kad "Trust access to the VBA project object model" nije dozvoljen | dugme "Otvori VBA editor" (akcija `vbaopen`, `modAdmin.bas:244`); dugme na listu `Pregled listova` | otvoren VBE prozor | Prvo cist put (`Application.VBE.MainWindow.Visible`), pa fallback na `SendKeys "%{F11}"` -- korisnik ne vidi gresku nego editor | -- | `modPregledListova.OtvoriVBA:93` | ne | NEPROVERENO |
| F-061 | administrator | Jednokratno uvesti podatke iz starog OtkupApp fajla, bez tihog prepisa postojecih | dugme "Migracija iz starog fajla" (akcija `migracija`, `modAdmin.bas:245`); Alt+F8 `MigrirajPodatkeIzStarog`; dugme "Migracije" na listu `Pregled listova` (`modPregledListova.PokreniMigraciju:105`) | popunjene tabele + izvestaj o preskocenim kolonama | Bira se stari fajl; ako `tblOtkup` vec ima redove, trazi se izricita potvrda jer migracija PREPISUJE; prazna "kriticna" kolona posle prenosa se prijavljuje, audit kolone se tiho preskacu | pise skoro sve tabele sveske; poredi po IMENU kolone | `modMigracija.MigrirajPodatkeIzStarog:28`, `JeKriticnaKolona:615` | da -- spisak "kriticnih" veznih kolona imenuje `brojzbirne`, `brojotpremnice`, `brojdokumenta` i `klasa` kao identitet/vezu (`modMigracija.bas:625-627`) | NEPROVERENO |
| F-062 | administrator | Isprazniti 22 tabele od podataka, a zadrzati zaglavlja i same tabele | dugme "Ocisti tabele od podataka" (akcija `ocisti`, `modAdmin.bas:246`); Alt+F8 `OcistiTabele`; dugme na listu `Pregled listova` | obrisani `DataBodyRange` redovi + poruka koliko je tabela ocisceno | Nema undo, pa se ne pita Da/Ne nego se mora UKUCATI "OBRISI" -- da jedan pogresan klik ne moze da obrise svesku | brise `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke`, `tblAmbalaza`, `tblPaleta`, `tblPaletaStavka`, `tblPrerada`, `tblPreradaStavka`, `tblArtikli`, `tblMagacin`, `tblNovac`, `tblKupci`, `tblStanice`, `tblKulture`, `tblTipAmbalaze`, `tblTipPalete`, `tblKooperanti`, `tblVozaci`, `tblCenovnik` | `modPregledListova.OcistiTabele:131` | ne -- spisak je po tabelama, ne po kolonama | NEPROVERENO |

## F2b -- podesavanja (`modPodesavanja`, config editor)

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-063 | administrator | Doci do podesavanja iz sidebara, i biti odbijen ako nije administrator | sidebar, red `MAT_PODESAVANJA` (`modUiScreens.bas:180`); registar panela `modUiPanel.PanelRedovi:69` | panel sa poljima, naslov `Poruka("CFG_LBL_PODESAVANJA_TBLSEFCONFIG")` | Ista dvostruka brana kao Admin: sidebar prigusi, a `BuildConfigEditor` sam odbija sa `Poruka("AUTH_MSG_SAMO_ADMIN_SEKCIJA")` (AUD-033) | cita `tblKorisnici` (prava) | `modPodesavanja.BuildConfigEditor:251` | ne | `modTest.T_UiPanel_ZivotniCiklusIPrava:14855` -- "svaki modul panela zatvara SVOJ panel po imenu modula", "panel je uzeo radnu povrsinu" |
| F-064 | administrator | Urediti sve operativne kljuceve kroz imenovana polja umesto kroz sirovu tabelu, po grupama | prekidac grupa (akcija `grp`, `modPodesavanja.ConfigEditor_ToggleGroup:461`) | jedna grupa polja u dve kolone; `memo` polja preko celog reda | Trinaest grupa: "Prodavac (firma)", "Otkup / dokumenta", "Stampa", "Malina rezim", "Management / Klijent", "SEF", "Monitoring", "Sinhronizacija", "Google", "Banka / lokalno", "Banka / nalozi", "Napredno / Test". Tip polja odredjuje kontrolu (`bool` = DA/NE, `list:` = combo, `int`, `secret`, `memo`, `text`); dodavanje kljuca = jedan red u `ConfigEditorFields` | cita `tblSEFConfig` i `tblLocalConfig`; `list:` izvori citaju `tblKulture` i `tblKupci` | `modPodesavanja.ConfigEditorFields:61`, `CfgAdd:225`, `RelayoutGroups:499` | ne | NEPROVERENO |
| F-065 | administrator | Sacuvati podesavanja i tacno videti koje polje nije proslo, umesto delimicno snimljenog config-a | dugme "Sacuvaj" (akcija `save`, `modPodesavanja.ConfigEditor_OnClick:446`) | `MsgBox` "Sacuvano: N polja", uz `Poruka("CFG_MSG_PRESKOCENO_GRESKA")` + spisak kad neko polje padne | Svako polje se upisuje pod SVOJIM rukovaocem greske, pa jedan pad ne prekida petlju; `int` polje koje nije broj se odbija po imenu. Polje ide u `tblLocalConfig` ako mu je `store="local"` (putanje po masini), inace u `tblSEFConfig` | pise `tblSEFConfig` (`SetConfigValue`) i `tblLocalConfig` (`SetLocalConfigValue`) | `modPodesavanja.SaveConfigEditor:689` | ne | NEPROVERENO |
| F-066 | administrator | Ne izgubiti otkucano tiho -- ni pri izlasku iz panela ni pri navigaciji ljuskom | dugme "Nazad" (akcija `back`) i navigacija ljuske (`modUiPanel.PanelSmemoDaZatvorimo:261`) | pitanje `Poruka("UIPAN_ASK_ODBACI")` pre zatvaranja | Pitanje postavlja PANEL, ne ljuska, pa vazi i kad ljuske nema; posle uspesnog snimanja polje se vise ne broji kao izmenjeno | poredi vrednosti polja sa ucitanim | `modPodesavanja.Podesavanja_ImaNesacuvano:663`, `modUiPanel.PanelImaNesacuvano:245` | ne | NEPROVERENO |
| F-067 | administrator | Podesiti `pdftotext.exe` (Poppler) za uvoz izvoda, tako da putanja prati radnu svesku kad god moze | inline dugme uz polje `PDFTOTEXT_EXE_PATH` (`modPodesavanja.ConfigEditor_PickPoppler:475`); Alt+F8 `SetupPopplerInteractive` | poruka sa nadjenom putanjom; polje u editoru se odmah osvezi | Ako Poppler VEC stoji pored `.xlsm`-a, upisuje se PRAZNA vrednost = automatski rezim (putanja se racuna relativno, pa premestanje paketa ne kvari uvoz); inace picker + apsolutna putanja, uz napomenu sta to znaci. Trazi se i u `\Library\bin`, `\bin`, `\poppler\Library\bin`, `\poppler\bin` | pise `tblLocalConfig` (`PDFTOTEXT_EXE_PATH`) | `modSetup.SetupPopplerInteractive:301`, `FindPdfToTextExe:349` | ne | NEPROVERENO |
| F-068 | administrator | Izabrati folder klikom umesto kucanjem putanje, za svako `BANKA_*_PATH` polje | inline "..." dugme uz path polja (`modPodesavanja.ConfigEditor_PickFolderInto:489`) | izabrana putanja upisana u polje | Picker samo PUNI polje -- upis u `tblLocalConfig` radi "Sacuvaj", isti model kao ostala polja | pise `tblLocalConfig` tek kroz F-065 | `modPodesavanja.ConfigEditor_PickFolderInto:489`, `modSetup.PickFolder:2193` | ne | NEPROVERENO |
| F-069 | administrator | Otkriti ili sakriti sirovu tabelu `tblSEFConfig`, kao izlaz u nuzdi kad editor nije upotrebljiv | dugme "Prikazi/Sakrij" u panelu (akcija `toggle`, `modPodesavanja.ConfigEditor_OnClick:447`); Alt+F8 `ShowConfigSheet` / `HideConfigSheet` | list postaje vidljiv ili `xlSheetVeryHidden`, uz poruku sta to znaci | `ShowConfigSheet` nosi svoju branu (`AUTH_MSG_SAMO_ADMIN_SEKCIJA`, AUD-033) da se editor ne zaobidje; `SetupNewPC` sam sakriva list cim je setup zelen | menja vidljivost lista `tblSEFConfig` | `modPodesavanja.ToggleConfigSheet:775`, `ShowConfigSheet:801`, `HideConfigSheet:793` | ne | NEPROVERENO |

## F2c -- setup i rezimi masine (`modSetup`)

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-070 | odrzavanje | Podesiti nov racunar jednim potezom i dobiti spisak svega sto jos fali | Alt+F8 `SetupNewPC`; ponuda na prvom pokretanju (**F-043**); dugme "Ensure" u Admin panelu (**F-056**) | `MsgBox` `Poruka("SETUP_MSG_SETUP_USPESNO_ZAVRSEN")` ili `Poruka("SETUP_MSG_SETUP_ZAVRSEN_ALI")` + spisak; upis u `tblLocalConfig` i setup log | Zeleno se upisuje SAMO ako nijedna provera nije nista vratila: `APP_SETUP_COMPLETED = DA` + masina, Windows korisnik i vreme. Zivi link ka serveru je ADVISORY -- prijavljuje se, ali offline Google NE obara zeleno. Cim je zeleno, `tblSEFConfig` se sakriva (anti-tamper) | pise `tblLocalConfig` (`APP_SETUP_COMPLETED`, `APP_SETUP_COMPLETED_AT`, `APP_SETUP_MACHINE_NAME`, `APP_SETUP_WINDOWS_USER`, `APP_LAST_HEALTHCHECK_AT`); pravi tabele iz `schema/schema.json` | `modSetup.SetupNewPC:41`, `modSchema.EnsureAllTables:235` | ne | NEPROVERENO |
| F-071 | odrzavanje | Proveriti zdravlje podesenog racunara bez ponovnog setup-a, i znati da li aplikacija sme da radi | dugme "Health check (setup)" u Admin panelu (akcija `healthsetup`, `modAdmin.bas:237`); Alt+F8 `RunSetupHealthCheck` | `MsgBox` `Poruka("SETUP_MSG_HEALTH_CHECK_PROSAO")` ili `Poruka("SETUP_MSG_HEALTH_CHECK_NASAO")` + spisak | Sedam provera: okruzenje, osnovni folderi, `pdftotext.exe`, Google OAuth config, SEF config, obavezne tabele, obavezne kolone; plus advisory server link. Kratku verziju (`IsSetupHealthy`) koristi aplikacija: zeleno je samo ako `APP_SETUP_COMPLETED = DA` I folderi postoje I tabele postoje | cita `tblLocalConfig`, `tblSEFConfig`; pise `APP_LAST_HEALTHCHECK_AT` | `modSetup.RunSetupHealthCheck:130`, `IsSetupHealthy:171` | delimicno -- `CheckRequiredColumnsForSetup` proverava zatecen spisak kolona; spisak nije citan u ovoj sesiji (NEPROVERENO), ali isti obrazac u `modProductionHealthCheck` je zavistan (v. **F-089**, AUD-055) | NEPROVERENO |
| F-072 | odrzavanje | Prebaciti aplikaciju u potpuno lokalni rad, ili je vratiti na cloud/PWA, jednim makroom | Alt+F8 `EnableDesktopOnlyMode` / `EnableCloudSyncMode` | `MsgBox` sa uputstvom da se `SetupNewPC` pokrene ponovo; upis u setup log | Isti prekidac (`CLOUD_SYNC_ENABLED`) koji vec gasi runtime sync (`modBrojevi`, `modStanicaLock`); dok je iskljucen, `SetupNewPC` ne trazi Google kredencijale | pise `tblSEFConfig` (`CLOUD_SYNC_ENABLED`) | `modSetup.EnableDesktopOnlyMode:206`, `EnableCloudSyncMode:226` | ne | NEPROVERENO |
| F-073 | odrzavanje | Podesiti sve bankarske foldere kroz pickere, i dobiti ih napravljene | Alt+F8 `SetupBankFoldersInteractive` | `MsgBox` `Poruka("SETUP_MSG_BANKARSKI_FOLDERI_PODESENI")`; napravljeni folderi | Drive izvor se NE pravi (pravi ga Drive sync) nego se samo pamti, i Cancel ga preskace; Inbox je obavezan (Cancel = odustajanje), a Processed i Error dobijaju podrazumevane podfoldere Inbox-a ako se preskoce | pise `tblLocalConfig` (`BANKA_DRIVE_SOURCE_PATH`, `BANKA_INBOX_PATH`, `BANKA_PROCESSED_PATH`, `BANKA_ERROR_PATH`) | `modSetup.SetupBankFoldersInteractive:243` | ne | NEPROVERENO |
| F-074 | odrzavanje | Proveriti da li desktop stvarno vidi server (Google / GAS / banka Drive folder) | Alt+F8 `TestServerLink` | `MsgBox` "Server link OK" ili spisak stavki za proveru | Ista provera koja u `SetupNewPC` stoji kao advisory; ovde je sama sebi ishod, pa se moze pokrenuti kad se sumnja na mrezu a ne na setup | cita `tblSEFConfig` (Google/monitoring kljucevi), `tblLocalConfig` (banka putanje) | `modSetup.TestServerLink:855` | ne | NEPROVERENO |
| F-075 | administrator | Napraviti prvog administratora na svezoj svesci, pre nego sto se prijava uopste ukljuci | Alt+F8 `KreirajPrvogAdmina` | upisan red u `tblKorisnici` | Anti-lockout: dok nema nijednog naloga `MozeAdministraciju` propusta, pa se prvi admin uopste moze napraviti. Pre upisa se tvrdi ugovor o formatu celije (`SchemaReadyOrFail`) -- kolona koja nije "@" bi TIHO pokvarila vrednost i to se ne bi videlo ni u jednoj kasnijoj proveri | pise `tblKorisnici` | `modSetup.KreirajPrvogAdmina:1750`, `modSchema.SchemaReadyOrFail` | ne | NEPROVERENO |
| F-076 | administrator | Ukljuciti ili iskljuciti prijavu za celu svesku, ali ne tako da se svi zakljucaju napolje | Alt+F8 `EnableAuth` / `DisableAuth` | `MsgBox` `Poruka("AUTH_MSG_PRIJAVA_UKLJUCENA")` / `Poruka("AUTH_MSG_NEMA_ADMINA")` | `EnableAuth` odbija posao ako ne postoji nijedan AKTIVAN admin -- inace bi sledece pokretanje trazilo prijavu koju niko ne moze da prodje | cita `tblKorisnici`; pise `tblSEFConfig` (`AUTH_ENABLED`) | `modSetup.EnableAuth:1823`, `DisableAuth:1847` | ne | NEPROVERENO |
| F-077 | administrator | Ukljuciti ili iskljuciti hesovanje PIN-ova, uz dokaz da SHA uopste radi na ovoj masini | Alt+F8 `EnablePinHash` / `DisablePinHash` | `MsgBox` `Poruka("AUTH_MSG_PINHASH_ISKLJUCEN")` ili `Poruka("AUTH_ERR_SHA_NE_RADI")` | Ukljucenje se odbija ako `Sha256Hex("abc")` ne daje poznati otisak -- da se PIN-ovi ne hesuju necim sto ne radi i time zakljucaju nalozi (isto sto meri makro `modAuth.TestPinHash:442`) | pise `tblSEFConfig` (`PIN_HASH_ENABLED`); posredno `tblKorisnici` (**F-038**) | `modSetup.EnablePinHash:1866`, `DisablePinHash:1892` | ne | NEPROVERENO |
| F-078 | odrzavanje | Imati sve PDF foldere pored sveske napravljene i pre prvog generisanog dokumenta | Alt+F8 `EnsureAllDocFolders`; poziva ga i setup | napravljeni folderi na disku | Devet kategorija: otkupni listovi, prijemnice, otpremnice, reversi, kartice, paletni listovi, prerada, specifikacije, izvestaji | -- | `modSetup.EnsureAllDocFolders:2180`, `EnsureDocFolder:2151` | ne | NEPROVERENO |

## F2d -- azuriranje, izdanja i licenca

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-079 | operater | Dobiti ponudu za novu verziju pri otvaranju fajla i azurirati se bez ijednog rucnog koraka | otvaranje sveske -> `modMain.bas:74` -> `CheckForUpdateOnOpen`; "Da" pokrece `RunSelfUpdate` kroz `Application.OnTime` (`modMain.bas:75`); rucno kroz Admin panel (**F-055**) | `MsgBox` sa novom i tekucom verzijom + `Poruka("SU_AZURIRATI_SADA")`; po zavrsetku restart sveske na novoj verziji | Opt-in (bez `REL_FOLDER_ID` nema nicega) i tih kad je offline ili vec azurno. ATOMARNOST: mora stici SVAKI fajl iz listinga, inace se NISTA ne snima; nema li izmena koda, delta-skip ne dira nista; forma ili sheet koji se ne mogu azurirati = nista se ne snima i sveska se zatvara bez snimanja, pa na disku ostaje stara ISPRAVNA verzija | cita Drive `REL_FOLDER_ID`; pise `src-vba` module u projekat i backup fajl | `modSelfUpdate.CheckForUpdateOnOpen:58`, `RunSelfUpdate:110`, `RunSelfUpdateCore:155`, `RunSelfUpdatePhase2:262` | ne | NEPROVERENO |
| F-080 | odrzavanje | Ne ostati na polusnimljenoj verziji kad se azuriranje prekine (Excel zatvoren, `OnTime` otkazan) | automatski pri otvaranju, iz `CheckForUpdateOnOpen` (`modSelfUpdate.bas:66`); `AbortSelfUpdateClose` kao rucni fallback | `MsgBox` "Prethodni self-update nije zavrsen do kraja. Radna verzija je ocuvana" | Posto se pre punog uspeha NIKAD ne snima, fajl na disku je stara ispravna verzija -- zato se samo cisti `pending` marker i temp folder, a faza 2 se NE pokrece nad starim projektom (mesala bi staru i novu verziju). `AbortSelfUpdateClose` pre zatvaranja postavlja `Saved = True` da ni `ShutdownApp` ni `Close` ne upisu polu-nov projekat preko dobre kopije | `SaveSetting`/`GetSetting` sekcija po svesci (`P2Section`) | `modSelfUpdate.RecoverPendingSelfUpdate:93`, `AbortSelfUpdateClose:564` | ne | NEPROVERENO |
| F-081 | razvoj-release | Isprobati self-update iz lokalnog git klona, bez objave na Drive | Alt+F8 `RunSelfUpdateDev` | isti tok kao pravi self-update, ali iz izabranog `src-vba` foldera; backup u `Backup\AgriX_pre-update_*.xlsm` | DEV GUARD: folder se odbija ako nije git klon (ime `src-vba` + `.git` u roditelju), pa se makro ne moze slucajno pokrenuti nad proizvoljnim folderom na klijentskoj masini | cita izabrani `src-vba` folder | `modSelfUpdate.RunSelfUpdateDev:1608`, `IsDevCloneFolder` | ne | NEPROVERENO |
| F-082 | razvoj-release | Vratiti celu flotu na stariju vec objavljenu verziju | Alt+F8 `RollbackReleaseTo` | prepisan `current.json` + poruka sa `manifest_sha256` | Verzija mora postojati u `releases\` I imati `manifest.json`, inace odbijanje; `manifest_sha256` se racuna iz STVARNO skinutih bajtova, ne prepisuje se; potvrda izricito kaze da ce klijenti koji jos nisu azurirani povuci STARIJU verziju | cita/pise Drive `REL_FOLDER_ID` | `modRelease.RollbackReleaseTo:250` | ne | NEPROVERENO |
| F-083 | razvoj-release | Videti koje su verzije objavljene i na koju pokazuje kanal, pre nego sto se bira cilj rollback-a | Alt+F8 `ListReleases` | `MsgBox` sa sortiranim spiskom verzija i tekucom iz `current.json` | Citanje bez ijedne izmene -- namenjeno kao korak PRE `RollbackReleaseTo` | cita Drive `REL_FOLDER_ID` | `modRelease.ListReleases:304` | ne | NEPROVERENO |
| F-084 | razvoj-release | Proci jednu release kapiju umesto da se svaka suite pokrece rucno | Alt+F8 `RunE2EReleaseGate_v610` | `Debug.Print` dnevnik sa PASS/WARN/FAIL brojacima | Kapija ORKESTRIRA, ne zamenjuje suite: vozi sedam VBA suita (JSON parser, Google sync, MasterSync, Novac, Faktura, Business flow, Production health check) i DOKUMENTUJE dve koje se moraju pokrenuti iz Apps Script editora (`runGasRouteHealthCheck`, `runGasSmokeSuite`) | -- | `modE2EReleaseGate.RunE2EReleaseGate_v610:23`, `E2E_RunVbaSuite:73` | delimicno -- sam ne cita kolone, ali vozi `RunProductionHealthCheck` koji jeste zavistan (**F-089**); modul to i najavljuje kao poznat WARN (`modE2EReleaseGate.bas:55-56`) | NEPROVERENO |
| F-085 | odrzavanje | Aktivirati licencu za ovaj racunar kad je start odbijen, umesto samo da se vidi blokada | Alt+F8 `ActivateLicensePrompt`; poruke kapije same upucuju na njega (`modLicense.bas:166`, `:323`) | `InputBox` za kljuc, pa sveza online provera | Unos kljuca forsira svezu proveru: brisu se `LICENSE_NEXT_CHECK` i `LICENSE_BOUND_PARTS`, pa kesirana odbijenica ne moze da prezivi aktivaciju. Kapija koja odbija start je **F-041** | pise `tblSEFConfig` (`LICENSE_KEY`, `LICENSE_NEXT_CHECK`, `LICENSE_BOUND_PARTS`) | `modLicense.ActivateLicensePrompt:341` | ne | NEPROVERENO |
| F-086 | odrzavanje | Procitati otisak ovog racunara da bi se izdao kljuc za bas njega | Alt+F8 `LicenseShowDevice` | `MsgBox` sa `MachineGuid`, `SMBIOS UUID`, `Volume SN` i imenom racunara | Cista dijagnostika, bez ijedne izmene -- podatak koji dobavljac trazi da bi vezao kljuc | -- | `modLicense.LicenseShowDevice:397` | ne | `modLicenseTests.TestLicense_DeviceFingerprint:94` |

## F2e -- integritet i zdravlje

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-087 | odrzavanje | Dobiti sve neuskladjenosti u podacima kao jedan pregledan list, sa ukupnim brojem | dugme "Integritet provere (tabele)" u Admin panelu (akcija `integritet`, `modAdmin.bas:239`); ekran oporavka, lista `INTEGRITET` (**B-048**) | sheet `INTEGRITET` sa blokom po proveri + `MsgBox` rezime "UKUPNO: N neuskladjenih zapisa" | Dvadeset dve provere, uvek sve (v. tabelu ispod); nalaz je OPIS neslaganja, ne stavka koja se prevezuje -- popravka ide svojim tokom. `GetIntegritetRows` vraca isti prolaz kao ravan spisak za ekran, bez diranja sheet-a | cita `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblPaleta`, `tblPaletaStavka`, `tblPrerada`, `tblPreradaStavka`, `tblAmbalaza` | `modIntegritet.RunIntegritetProvere:36`, `GetIntegritetRows:70`, `RunAllChecks:93` | **da** -- 15 od 22 provere; puna presuda po proveri u tabeli ispod | `modTest.T_Integritet_VidiDvosmislenBrojIPraznuGeneraciju:17573` -- "B9 vidi aktivnu zbirnu bez GeneracijaID", "B8 vidi broj sa dva aktivna dokumenta" |
| F-088 | odrzavanje | Znati koliko nalaza integriteta postoji bez ponovnog prolaza kroz podatke | `modIntegritet.IntegritetUkupno:89` (naslov liste `INTEGRITET`, `modScrOporavak.Scr_NaslovDopuna:114`) | broj u naslovu liste, uz `Poruka("OTKUI_OPO_INT_NALAZA")` | Broj je iz POSLEDNJEG prolaza -- ne pokrece provere ponovo, pa naslov ne kosta drugi prolaz kroz svesku | -- | `modIntegritet.IntegritetUkupno:89` | da -- broji nalaze provera iz **F-087** | NEPROVERENO |
| F-089 | odrzavanje | Dobiti presudu o zdravlju PRODUKCIJSKE sveske (sema, duplikati, novac, fakture, isplate, sync) | dugme "Production health check" u Admin panelu (akcija `healthprod`, `modAdmin.bas:238`); Alt+F8 `RunProductionHealthCheck`; release kapija (**F-084**) | dnevnik sa PASS/WARN/FAIL po proveri + zavrsni rezime | Devetnaest provera (v. tabelu ispod), grupisane u semu, kljuceve, novac, fakture, isplate i Google sync. Health check SAMO CITA -- `EnsureAllTables` se ovde namerno NE zove, jer provera ne sme da menja ono sto meri | cita `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke`, `tblNovac`, `tblAmbalaza`, `tblBankaImport`, `tblParcele`, `tblKooperanti`, `tblStanice`, `tblKulture`, `tblSEFConfig` | `modProductionHealthCheck.RunProductionHealthCheck:28` | **da** -- `Check_CoreTablesAndColumns` trazi linijska polja na zaglavljima i vezne kolone starog modela; poznat kvar **AUD-055** (`docs/KNOWN_ISSUES.md`) | NEPROVERENO |

### Provere integriteta -- presuda po proveri (**F-087**)

| Provera | Sta tvrdi | Kolone koje cita | Zavisi od starog modela? |
|---|---|---|---|
| `Chk_A1_OtpremnicaVsZbirna:124` | zbir kg otpremnica po `BrojZbirne` = ukupno na zbirnoj | `tblOtpremnica.BrojZbirne`, `tblZbirna.BrojZbirne` | **da** -- poslovni broj kao veza; `ValidateZbirna` agregira Klasu I+II sa zaglavlja |
| `Chk_A2_ManjakAnomalije:160` | manjak prijemnica prema zbirnoj nije van praga | `tblPrijemnica.BrojZbirne`, `tblPrijemnica.Kolicina`, `tblZbirna.UkupnoKolicina` | **da** -- poslovni broj kao veza + linijska `Kolicina` na zaglavlju prijemnice |
| `Chk_A3_StavkeVsPrijemnica:822` | zbir neta paletnih stavki = kolicina na prijemnici | `tblPaletaStavka.PrijemnicaID`, `tblPaletaStavka.Neto`, `tblPrijemnica.Kolicina` | **da** -- veza je ispravna (`PrijemnicaID`), ali poredi se sa linijskom `Kolicina` na ZAGLAVLJU prijemnice |
| `Chk_A4_PaletaHeaderVsStavke:853` | zaglavlje palete = zbir njenih stavki (neto i gajbice) | `tblPaleta.Neto/BrGajbica/PaletaID`, `tblPaletaStavka.*` | ne -- paleta je vec zaglavlje + stavke |
| `Chk_A5_PreradaKg:998` | ulazni i izlazni kg prerade se slazu sa stavkama | `tblPrerada.NetoUlaz/NetoIzlaz`, `tblPreradaStavka.Neto` | ne -- prerada je vec zaglavlje + stavke |
| `Chk_B1_Verwaiste:207` | ziva otpremnica/prijemnica ciji je zbirni dokument potpuno storniran | `BrojZbirne`, `BrojOtpremnice`, `Kolicina` na zaglavljima | **da** -- poslovni broj kao veza + linijska `Kolicina` na zaglavlju |
| `Chk_B2_UnlinkedOtkupi:229` | aktivan otkup bez otpremnice | `tblOtkup.OtpremnicaID`, `VozacID`, `Kolicina` | **da** -- `Otkup.OtpremnicaID` i `Otkup.VozacID`, plus linijska `Kolicina` |
| `Chk_B3_IzgubljeniBlokovi:247` | aktivan otkup ciji `OtpremnicaID` pokazuje na storniranu ili nepostojecu otpremnicu | `tblOtkup.OtpremnicaID`, `Kolicina` | **da** -- `Otkup.OtpremnicaID` je sama premisa provere |
| `Chk_B4_DanglingBrojZbirne:265` | ziv dokument sa `BrojZbirne` koji ne postoji ni u jednom redu `tblZbirna` | `tblOtpremnica.BrojZbirne/Kolicina`, `tblPrijemnica.BrojZbirne/Kolicina` | **da** -- poslovni broj kao veza |
| `Chk_B5_PrijemnicaBezZbirne:287` | ziva prijemnica bez `BrojZbirne` | `tblPrijemnica.BrojZbirne`, `Kolicina` | **da** -- poslovni broj kao veza |
| `Chk_B5b_OtpremnicaBezZbirne:303` | ziva otpremnica bez `BrojZbirne` | `tblOtpremnica.BrojZbirne`, `Kolicina` | **da** -- poslovni broj kao veza |
| `Chk_B6_ZbirnaCaseMismatch:593` | isti `BrojZbirne` pisan razlicitim slovima kroz cetiri tabele | `tblOtkup.BrojZbirne`, `tblOtpremnica.BrojZbirne`, `tblPrijemnica.BrojZbirne`, `tblPaletaStavka.BrojZbirne` | **da** -- cetiri kopije poslovnog broja kao veze, `Otkup.BrojZbirne` medju njima |
| `Chk_B7_ZbirnaNulaKg:320` | aktivna zbirna sa nula kg | `tblZbirna.UkupnoKolicina` | **da** -- linijska ukupna kolicina na zaglavlju zbirne (po klasi) |
| `Chk_B8_DvosmislenBrojZbirne:351` | isti `BrojZbirne` na vise aktivnih zbirnih sa razlicitim kupcem ili vozacem | `tblZbirna.BrojZbirne`, `KupacID`, `VozacID` | **da** -- dvosmislenost POSTOJI zato sto je zbirna red-po-klasi; u novom modelu je broj kljuc zaglavlja |
| `Chk_B9_ZbirnaBezGeneracije:412` | aktivna zbirna bez `GeneracijaID` | `tblZbirna.GeneracijaID` | **da** -- `GeneracijaID` |
| `Chk_B10_ReversBezID:459` | red ambalaze koji je revers a nema `ReversID` | `tblAmbalaza.ReversID/DokumentID/DokumentTip/EntitetTip`; `tblOtkup.OtkupID` (samo za izuzimanje) | ne -- cita samo `OtkupID` |
| `Chk_C1_C4_StavkaPrijemnica:625` | paletna stavka vezana na nepostojecu prijemnicu / neslaganje sa zbirnom | `tblPaletaStavka.PrijemnicaID/BrojPrijemnice`, `tblZbirna.UkupnoKolicina/UkupnoAmbalaze` | **da** -- poslovni broj uz ID kao veza + linijska ukupna polja na zaglavlju zbirne |
| `Chk_C2_StavkaBezZbirne:690` | paletna stavka bez `BrojZbirne` | `tblPaletaStavka.BrojZbirne` | **da** -- poslovni broj kao veza |
| `Chk_C3_PaletaBezStavke:733` | paleta bez ijedne stavke, a sa netom | `tblPaleta.Neto/Broj/Godina`, `tblPaletaStavka.Neto` | ne |
| `Chk_C5_DupliBrojPalete:773` | dve palete sa istim brojem u istoj godini | `tblPaleta.Broj`, `Godina` | ne |
| `Chk_D1_PreradjenoBezPrerade:903` | paleta oznacena kao preradjena bez ijedne stavke prerade | `tblPaleta.Preradjeno`, `tblPreradaStavka.PaletaID` | ne |
| `Chk_D2_PreradaNesvezaPaleta:947` | stavka prerade ciji `BrojPalete` i `PaletaID` ne pokazuju na isto | `tblPreradaStavka.BrojPalete`, `PaletaID` | delimicno -- poslovni broj stoji uz ID kao druga veza; provera POSTOJI zato sto postoje obe |

### Provere zdravlja -- presuda po proveri (**F-089**)

| Provera | Sta tvrdi | Zavisi od starog modela? |
|---|---|---|
| `Check_SchemaRegistry:68` | cela sveska odgovara registru `modSchema` | ne -- poredi sa kanonom, pa se menja sa kanonom |
| `Check_CoreTablesAndColumns:109` | sest kljucnih tabela ima svaku nabrojanu kolonu | **da, tvrdo** -- rucno kucan spisak trazi `Otkup.Kolicina/Cena/Klasa/VozacID/BrojZbirne/OtpremnicaID` (`:122-124`), `Otpremnica.Kolicina/Cena/KolAmbalaze/Klasa/BrojZbirne` (`:127-128`), `Zbirna.UkupnoKolicina/UkupnoAmbalaze/Klasa` (`:131-132`), `Prijemnica.Kolicina/Cena/KolAmbalaze/Klasa` (`:135-137`). Vec danas javlja FAIL na ispravnoj svesci (**AUD-055**) |
| `Check_DuplicateDocumentKeys:162` | nijedan dokumentni ID nije dupliran | ne -- samo ID kolone |
| `Check_NovacRowsAreFinanciallyValid:191` | red novca je finansijski ispravan | ne |
| `Check_FakturaStavkeReferences:268` | stavke fakture pokazuju na postojecu fakturu i prijemnicu | ne |
| `Check_PrijemnicaFakturaFlags:325` | `Fakturisano` i `FakturaID` na prijemnici se slazu | ne -- to nisu linijska polja |
| `Check_FakturaPaymentConsistency:389` | status i datum placanja fakture se slazu | ne |
| `Check_OtkupPaymentConsistency:505` | otvorena obaveza po otkupu se slaze sa isplatama | ne -- vrednost vec dolazi sa STAVKI (`HealthVrednostPoOtkupu`), a dokument bez stavki ispada iz liste |
| `Check_KooperantOtkupReconciliation:600` | zbir po kooperantu i stanici se slaze sa ukupnim | ne -- isto, vrednost sa stavki; komentar `:620-622` izricito kaze zasto `Kolicina x Cena` sa zaglavlja ne valja |
| `Check_FakturaIznosReconciliation:736` | iznos fakture = zbir njenih stavki | ne -- `tblFakturaStavke` je vec stavke |
| `Check_OtkupOtpremnicaCrossZbirnaLinks:849` | `BrojZbirne` otkupa se slaze sa otpremnicom na koju pokazuje | **da** -- `Otkup.BrojZbirne` + `Otkup.OtpremnicaID` + `Otpremnica.BrojZbirne` (`:865-867`, `:891`) |
| `Check_SEFOutboundConsistency:923` | SEF stanja faktura su konzistentna | ne |
| `Check_DocumentSoftDeleteReferences:1016` | ziv dokument ne pokazuje na storniran roditelj | **da** -- `Otkup.OtpremnicaID` (`:1022-1023`) |
| `Check_GoogleSyncHealth:1050` | sync deo kao celina | ne -- agregat pet provera ispod |
| `Check_GoogleSyncConfig:1066` | Google kljucevi upisani | ne |
| `Check_GoogleSyncAuth:1090` | token se dobija | ne |
| `Check_GoogleSyncFolderReachable:1115` | PWA folder dostupan | ne |
| `Check_GoogleSyncMasterSchema:1155` | tabele koje sync cita imaju svoje kolone | **da** -- trazi `Otkup.Kolicina/Cena/Klasa/OtpremnicaID` (`:1170-1174`) |
| `Check_GoogleSyncFeatureFlags:1190` | prekidaci sync-a su smisleno postavljeni | ne |

## F2f -- preostali makroi

| ID | KO | Sposobnost | Ulaz | Izlaz | Pravilo/ishod koji korisnik ocekuje | Podaci | Legacy implementacija | Zavisi od starog modela? | Testovi |
|---|---|---|---|---|---|---|---|---|---|
| F-090 | odrzavanje | Naknadno napraviti prijemnice za hladnjacke otpremnice koje ih nemaju | Alt+F8 `BackfillPrijemniceHladnjaca` | napravljene prijemnice + `MsgBox` sa brojem uspesnih i palih | Nadoknada za pauziran auto-lanac (**A-014**). Odbija posao ako `MALINA_DEFAULT_KUPAC` nije podesen; svaka prijemnica je atomicna, pa pad jedne ne rusi ostale; ako jedna klasa dokumenta vec ima prijemnicu, druga dobija ISTI broj (jedna prijemnica = jedan broj) | cita `tblOtpremnica`, pise `tblPrijemnica`; cita `tblSEFConfig` (`MALINA_DEFAULT_KUPAC`) | `modAutoHladnjaca.BackfillPrijemniceHladnjaca:439`, `BackfillPrijemniceHladnjacaCore:454` | **da** -- kljuc obrade je `BrojZbirne\|Klasa`, a izvor su po-klasna polja na ZAGLAVLJU otpremnice: `Kolicina`, `Cena`, `KolAmbalaze`, `Klasa`, `Bruto` (`modAutoHladnjaca.bas:488-493`) | NEPROVERENO |
| F-091 | odrzavanje | Popuniti `GeneracijaID` na dete-zbirnama koje su ostale bez njega | Alt+F8 `BackfillDeteZbirnaGeneracija` | `MsgBox` sa brojem popunjenih i preskocenih (nema generacije / vise kandidata / integritet) | Preskace se sve gde izbor nije jednoznacan -- backfill ne pogadja generaciju, nego je popunjava samo kad je jedna moguca | pise `tblZbirna.GeneracijaID` (i srodne tabele iz `EnsureSledljivostSchema`) | `modSetup.BackfillDeteZbirnaGeneracija:1454`, `BackfillDeteZbirnaGeneracija_Core:1459` | **da** -- `GeneracijaID` | NEPROVERENO |
| F-092 | odrzavanje | Popuniti `Otkup.BrojOtpremnice` iz otpremnica na koje otkupi pokazuju | Alt+F8 `BackfillOtkupBrojOtpremnice` | popunjena kolona + `MsgBox` sa brojem redova | Kolona se prvo obezbedi, pa se tvrdi ugovor o formatu (`SchemaReadyOrFail`) PRE upisa -- kolona koja nije "@" bi tiho pokvarila broj ("3/2026" -> datum) i to se ne bi videlo ni u jednoj kasnijoj proveri | cita `tblOtpremnica.OtpremnicaID/BrojOtpremnice`; pise `tblOtkup.BrojOtpremnice` | `modSetup.BackfillOtkupBrojOtpremnice:1569` | **da** -- `Otkup.BrojOtpremnice` i `Otkup.OtpremnicaID` | NEPROVERENO |
| F-093 | odrzavanje | Naknadno uskladiti kolicine na paletama za prijemnicu, kad je pri unosu preskoceno | Alt+F8 `PaletaAdjust_Prompt` | `MsgBox` "Kolicine na paletama korigovane" ili razlog zasto nije | Isti tok koji nudi i unos dokumenta: ako visak gajbica ne staje na paletu, operater bira PRELIJ (na sledecu otvorenu/novu paletu) ili PREKO kapaciteta na istoj; OTKAZI ostavlja palete netaknute i poruka izricito upucuje na ovaj makro za kasnije | cita/pise `tblPaleta`, `tblPaletaStavka` po broju prijemnice | `modPaletniListUI.PaletaAdjust_Prompt:253`, `PaletaAdjustPrompt:214` | **da** -- radi po POSLOVNOM broju prijemnice (`AdjustPaletaGajbiceZaPrijemnicu_TX(brojPrij, ...)`, `modPaletniListUI.bas:218`), ne po ID-u | NEPROVERENO |
| F-094 | odrzavanje | Doci do svakog lista sveske kroz jedan pregled sa linkovima i dugmadima za glavne komande | Alt+F8 `NapraviPregledListova` | list `Pregled listova`: red po listu sa klikabilnim linkom + dugmad iznad tabele, zamrznuto zaglavlje | Regenerise se u celini (i stara dugmad se brisu, jer ih `Cells.Clear` ne dira); dugmad zovu postojece ulazne tacke -- "Pokreni program" (**F-051**), "Otvori VBA" (**F-060**), "Migracije" (**F-061**), "Uvezi VBA" (**F-059**), "Ocisti tabele" (**F-062**) | cita imena svih radnih listova | `modPregledListova.NapraviPregledListova:24`, `DodajDugmad`, `EnsurePregledDugmad:249` | ne | NEPROVERENO |
| F-095 | razvoj-release | Proveriti da li OVAJ fajl uopste ima Google token i vidi release folder | Alt+F8 `DriveSelfTest` | `MsgBox` sa `REL_FOLDER_ID`, duzinom tokena i ishodom listinga | Prazan token se prijavljuje kao JASAN uzrok ("Google nije autentifikovan na OVOM fajlu") sa uputstvom sta uraditi, umesto tihog neuspeha objave ili self-update-a | cita Drive `REL_FOLDER_ID` | `modDrive.DriveSelfTest:277` | ne | NEPROVERENO |

## Pokrivenost ulaznih tacaka (F2)

### Admin panel -- svih dvanaest komandi (`modAdmin.AdminPanel_OnClick:215`)

| Akcija | ID |
|---|---|
| `checkupdate` (`:235`) | F-055 |
| `ensure` (`:236`) | F-056 |
| `healthsetup` (`:237`) | F-071 |
| `healthprod` (`:238`) | F-089 |
| `integritet` (`:239`) | F-087 |
| `googleauth` (`:240`) | F-057 |
| `publish` (`:241`) | F-058 |
| `vbaimport` (`:242`), `vbaexport` (`:243`) | F-059 |
| `vbaopen` (`:244`) | F-060 |
| `migracija` (`:245`) | F-061 |
| `ocisti` (`:246`) | F-062 |
| `back` (`:247`) -> `CloseAdminPanel:353` | tehnicko -- vracanje na dashboard; ciscenje radi `Admin_Release:344` |
| `grp:<naziv>` (`:226`) | F-054 |

### Podesavanja -- ruter (`modPodesavanja.ConfigEditor_OnClick:443`)

| Akcija | ID |
|---|---|
| `save` (`:446`) | F-065 |
| `toggle` (`:447`) | F-069 |
| `back` (`:448`) -> `CloseConfigEditor:749` | F-066 |
| `grp` (`:449`) | F-064 |
| `poppler` / `browsepoppler` (`:450`) | F-067 |
| `browsefolder` (`:451`) | F-068 |
| `ConfigEditorFields:61` | F-064 (registar polja; `LookupCSV:217` i `MigrateBankaRacuniLegacy:235` su tehnicki) |
| `ApplyDefaultProizvod:890` | tehnicko -- primena kljuceva `DEFAULT_VRSTA_VOCA` / `DEFAULT_SORTA_VOCA` na combo-e ekrana unosa; sposobnost je u oblasti A |
| `Podesavanja_Release:903` | tehnicko -- otpusta `WithEvents` omotace pre self-update importa |

### Makroi (Alt+F8) razvrstani u ovoj sesiji

Svi javni `Sub` bez argumenata koji se ne pominju ni u `A.md`--`F.md` (206 imena,
od ukupno 287 u `src-vba/*.bas`).

**Sposobnosti (red u tabeli iznad ili u vec mapiranoj oblasti):**

| Makro | ID |
|---|---|
| `modSetup.SetupNewPC:41` | F-070 (i F-043, F-056) |
| `modSetup.RunSetupHealthCheck:130` | F-071 |
| `modSetup.EnableDesktopOnlyMode:206`, `EnableCloudSyncMode:226` | F-072 |
| `modSetup.SetupBankFoldersInteractive:243` | F-073 |
| `modSetup.SetupPopplerInteractive:301`, `modPodesavanja.ConfigEditor_PickPoppler:475` | F-067 |
| `modSetup.TestServerLink:855` | F-074 |
| `modSetup.KreirajPrvogAdmina:1750` | F-075 |
| `modSetup.EnableAuth:1823`, `DisableAuth:1847` | F-076 |
| `modSetup.EnablePinHash:1866`, `DisablePinHash:1892` | F-077 |
| `modSetup.EnsureAllDocFolders:2180` | F-078 |
| `modSetup.BackfillDeteZbirnaGeneracija:1454` | F-091 |
| `modSetup.BackfillOtkupBrojOtpremnice:1569` | F-092 |
| `modPodesavanja.ToggleConfigSheet:775`, `ShowConfigSheet:801`, `HideConfigSheet:793` | F-069 |
| `modSelfUpdate.RunSelfUpdate:110`, `RunSelfUpdatePhase2:262` | F-079 |
| `modSelfUpdate.RecoverPendingSelfUpdate:93`, `AbortSelfUpdateClose:564` | F-080 |
| `modSelfUpdate.RunSelfUpdateDev:1608` | F-081 |
| `modRelease.PublishReleaseToDrive:26` | F-058 |
| `modRelease.RollbackReleaseTo:250` | F-082 |
| `modRelease.ListReleases:304` | F-083 |
| `modE2EReleaseGate.RunE2EReleaseGate_v610:23` | F-084 |
| `modLicense.ActivateLicensePrompt:341` | F-085 |
| `modLicense.LicenseShowDevice:397` | F-086 |
| `modProductionHealthCheck.RunProductionHealthCheck:28` | F-089 |
| `modAutoHladnjaca.BackfillPrijemniceHladnjaca:439` | F-090 |
| `modPaletniListUI.PaletaAdjust_Prompt:253` | F-093 |
| `modPregledListova.NapraviPregledListova:24` | F-094 |
| `modPregledListova.PokreniProgram:82` | F-051 (ista ljuska, drugi ulaz) |
| `modPregledListova.OtvoriVBA:93` | F-060 |
| `modPregledListova.PokreniMigraciju:105` | F-061 |
| `modPregledListova.UveziVBA:123` | F-059 |
| `modPregledListova.OcistiTabele:131` | F-062 |
| `modVbaTools.ImportAllVBA:235`, `ExportAllVBA:205` | F-059 |
| `modMigracija.MigrirajPodatkeIzStarog:28` | F-061 |
| `modGoogleAuth.RunGoogleAuthSetup:35` | F-057 |
| `modDrive.DriveSelfTest:277` | F-095 |
| `modAuth.Logout:325` | F-035 (zamena operatera) |
| `modBankaImport.ImportBankaInbox:193` | D-026 (isti lanac, `ImportBankaInbox_WithDrivePull`) |
| `modPaletniListUI.ExportPaletniListPDF_Prompt:14` | C-040 / C-051 |
| `modPaletniListUI.ExportPreradaPDF_Prompt:39` | C-040 |
| `modPaletniListUI.SavePrerada_Prompt:65` | C-055 |
| `modPaletniListUI.PrintNepotpunePalete_Prompt:107` | C-054 |
| `modPaletniListUI.ClosePaleta_Prompt:123` | C-052 |
| `modPaletniListUI.StornoPaleta_Prompt:149` | C-053 |
| `modPaletniListUI.StornoPrerada_Prompt:179` | C-058 |

### Mrtvo / tehnicko (NIJE sposobnost)

- **Test suite-ovi i pojedinacni testovi** -- razvojni alat, ne daju ishod
  operateru, administratoru ni odrzavanju: `modTest.RunAllTests:271`,
  `modGoldenTests.RunGoldenSuite:76`, `modBusinessFlowProTests.RunBusinessFlowProSuite:98`
  (+ `RunBusinessFlowProSeedOnly:332`, `RunBusinessFlowProTraceabilityOnly:350`,
  `RunBusinessFlowProAuditOnly:369`, `SoftStornoBusinessFlowTestRows:1404`,
  `HardDeleteBusinessFlowTestRows:13725`), `modFakturaTests.RunFakturaSmokeSuite:16`,
  `modNovacTests.RunNovacSmokeSuite:16`, `modIzvestajTests.RunIzvestajTests:46` i
  `SmokeTest_modIzvestaj:1126`, `modPaleteTests` (`modTestPalete.RunPaleteTestSuite:61`),
  `modTestStorno.RunStornoTestSuite:42`, `modTestStornoCentar.Test_StornoCentar_All:27`
  (+ 24 `Test_*_Auto`), `modGoogleSyncSmokeTests.RunGoogleSyncSmokeSuite:31` /
  `RunMasterSyncSmokeSuite:388` / `RunSheetsJsonParserTests:1206`,
  `modSEFTests.RunSEFTestSuite:195` / `RunSEFStateTransitionSuite:246` /
  `RunHttpUtilsSmokeSuite:2294` / `RunSEFDocumentIdShapeSuite:2391`,
  `modSEFClient.RunSEFClientParserSmokeSuite:968`,
  `modLicenseTests.TestLicense_All:22` (+ `TestLicense_SplitParts:46`,
  `TestLicense_PartsMatch:66`, `TestLicense_NonEmptyParts:83`,
  `TestLicense_DeviceFingerprint:94`), `modMonitoringTests.TestMonitoring_All:4`
  (+ `TestMonitoring_Config:22`, `TestMonitoring_HTTP:33`, `TestMonitoring_ErrorEvent:49`,
  `TestMonitoring_SEFUnknown:72`, `TestMonitoring_BackupSuccess:99`,
  `TestMonitoring_BackupFail:115`), i sva `Test_*` imena
  (`modSelfUpdate.Test_DeltaSkip:1680`, `modDrive.Test_Sha256File:170`,
  `modSEFMapper.Test_BuildSEFInvoiceDto:1287` / `Test_SerializeSEFRequest:1326` /
  `Test_PayloadHash:1345` / `Test_SerializeUBLInvoice:1361`,
  `modSEFClient.Test_SubmitUBLInvoice:935`, `modBankaAlta.Test_AltaParse:607`,
  `modBankaProCredit.Test_ProCreditParse:483`,
  `modBankaImport.Test_BankParse:500` / `Test_SaldoIntegrityOnSamplePDF:1634`,
  `modBankaMapiranje.Test_GetBankaImportOpen:3659` i pet `Test_Map*_TX` /
  `Test_SkipBankaImportRow_TX:3753`, `modStornoImpact.Test_BuildStornoImpact:472`,
  `modStornoRecovery.Test_GetNedovrseno:432` / `Test_UndoStorno:445`,
  `modDokumentInvariant.Test_FindSingleActiveRow:643`).
- **Test seam-ovi i reset-i za testove** (suzavaju sta prolazi, ne daju ishod):
  `Scr_*TestReset` u `modScrAgro:1732`, `modScrBankaNalozi:1741`,
  `modScrBankaUvoz:1980`, `modScrFakture:2327`, `modScrIzvestaji:2893`,
  `modScrSledljivost:1955`; `modScrStorno.Scr_ErrTestPrljav:193`;
  `modOtkupUI.GridSortAktivacijaTest:3097`, `OtkupUI_PrimeniNovaPravaTest:4255`,
  `OtkupUI_IsprazniPovrsinuTest:4273`, `GridOtkaciFormuTest:6107`;
  `modJournaling.ResetAutoSaveStateForTests:877`;
  `modDokumentInvariant.ResetIssuedZbirnaAudit:460`.
- **Dijagnostika za razvoj** (ispis u `Debug`/list, bez ishoda korisniku):
  `modOtkupUI.DiagOtkupUI:9206`, `Diag_Hladnjaca:7396`, `DumpMdl2Sheet:9155`,
  `DumpMdl2Used:9285`, `DetectDisplayFont:7816`; `modScrBankaNalozi.Diag_BnRedovi:1634`,
  `modScrBankaUvoz.Diag_BuRedovi:1819`, `modScrIzvestaji.Diag_IzRedovi:2791`,
  `modScrSledljivost.Diag_SlRedovi:1869`; `modBankaImport.Diag_DumpPdfTextAroundStanje:1563`,
  `modBankaImportParserPdfToText.Diag_DumpFullPdfText:1197`;
  `modSetup.DebugKoloneTabele:1944`.
- **Idempotentni graditelji seme i sablona** -- zovu ih setup, `Ensure` agregat
  (**F-056**) i sami tokovi stampe pre prve upotrebe; nemaju svoj ishod:
  `modSetup.EnsurePoruke:1120`, `EnsureCenovnikSchema:999`, `EnsurePaletniListSchema:919`,
  `EnsureDoradeSchema:1619`, `EnsureKorisniciSchema:1729`, `EnsureAuditColumns:1064`,
  `EnsureStornoVezeSchema:1040`, `EnsureStornoVezeSchemaCore:1021`,
  `EnsureStornoZurnalSchemaCore:1032`, `EnsureUtovarSchemaCore:1154`,
  `EnsureRuntimeSchema:1173`, `EnsureSledljivostSchema:1283`;
  `modPrint.EnsureOtpremnicaSablon:356`, `EnsureOtkupSablon:1029`,
  `EnsureGrupniOtkupSablon:1282`, `EnsurePrijemnicaSablon:1565`,
  `EnsureIzdavanjeAmbalazeSablon:2048`, `EnsureFakturaSablon:2077`,
  `EnsureUtovarSablon:2322`, `EnsureKarticaSablon:2584`, `EnsureSledljivostSablon:2823`,
  `EnsureKarticaAmbalazeSablon:3019`, `EnsureSpecifikacijaSablon:3180`,
  `EnsureIsplataSpecSablon:3332`; `modPaletniList.EnsurePaletaSablon:1061`,
  `EnsurePreradaSablon:3226`; `modPregledListova.EnsurePregledDugmad:249`.
- **Otpustanje referenci pre VBA importa / self-update-a** (`CallOptional` u
  `modSelfUpdate.bas:425`, `modVbaTools.bas:1584`): `modAdmin.Admin_Release:344`,
  `modPodesavanja.Podesavanja_Release:903`, `modOtkupUI.OtkupUI_Release:7908`,
  `modOtkupBlok.OtkupBlok_Release:2218`, `modLogo.LogoOtpusti:74`,
  `modUiFaze.FazaOtpusti:113`.
- **Kes i osvezavanje** (tehnicki, zovu ih ekrani i pisci):
  `modUiScreens.ScrResetCache:413`, `modUiData.ResetCache:39`,
  `modMaticniIzvor.MatResetCache:1099`, `modMaticniKorisnici.KorResetCache:620`,
  `modPoruke.InvalidateCache:20`, `modBrojevi.ClearSpreadsheetIDCache:668`,
  `modDataAccess.BeginTableCache:59` / `EndTableCache:76`, `modUiKit.ResetNumFields:34`,
  `modStornoWarm.InvalidateStornoWarm:42` / `ScheduleStornoWarm:49` /
  `StopStornoWarm:59` / `StornoWarmTick:65`, i `Scr_ResetCache` u svakom `modScr*`.
- **Ljuska i njen raspored** (zove ih `frmOtkupUI` i `modUiFaze`, nisu zasebne
  komande): `modOtkupUI.ShowOtkupUI:5332`, `OtkupUI_EnsureShellBuilt:5262`,
  `OtkupUI_FormClosed:7888`, `OsveziRasporedEkrana:4287`, `RenderGrid:2707`,
  `RefreshFromData:5869`, `EnsureGridLoaded:5821`, `OsveziNavBrojace:5835`,
  `RefreshPartnerLista:6824`, `PanelVracenNaEkran:1324`, `ResetAllBtnVisuals:7971`,
  `FocusKontekst:8030`, `FocusPretraga:8035`, `HideToast:7870`,
  `StopOtkupUITimers:7884`, `OtkupUI_StatsZone:9436`, `MakeResizable:5179`,
  `FixVbeCaption:5184`, `ClearForm:8725`; `modUiFaze.FazaSakrij:284`;
  `modMaticniEkran.ZatvoriEditor:885` i `Deaktiviraj:912` (ugovor motora, F1);
  `modOtkupBlok.OtkupBlok_ClearActiveOtp:769` i `OtkupBlok_RefreshKoopTotal:1858`.
- **Tajmeri i zivotni ciklus** (`Application.OnTime` ciljevi ili koraci gasenja):
  `modJournaling.AutoSaveTick:807`, `StopAutoSaveTimer:822`;
  `modLogError.LogAppShutdown:127`; `modStornoZurnal.AbortStornoOp:76`;
  `modLicense.DenyAccessAndScheduleClose:633`, `ForceCloseDeniedWorkbook:642`
  (ishod licencne kapije je **F-041**); `modVbaTools.ImportAllVBA_Phase2:486`,
  `ImportAllVBA_MergeStep:654` (koraci **F-059**).

## Sposobnosti koje postoje samo u PWA / GAS

Za deo F2 -- nijedna. Admin, podesavanja, setup, azuriranje, integritet i health
su u celini VBA; PWA nema svoj ekran nad `tblSEFConfig`/`tblLocalConfig`, a GAS
strana release kapije (`runGasRouteHealthCheck`, `runGasSmokeSuite`) se ne poziva
iz VBA nego se rucno pokrece iz Apps Script editora (**F-084**).

## NEPROVERENO (sa razlogom)

| ID | Sta nije provereno | Razlog |
|---|---|---|
| F-054 | raspored grupa i prelamanje segmenata | crta se iz rasporeda forme, trazi Excel |
| F-055 | ishod rucne provere azuriranja | trazi Drive i mrezu |
| F-056 | ishod "Ensure" agregata nad zatecenom sveskom | trazi Excel; nema testa koji tvrdi zavrsni rezime |
| F-057 | OAuth tok | trazi pregledac i Google nalog |
| F-058 | objava, versioned layout i immutability odbijanje | trazi Drive i autorizaciju |
| F-059 | VBA import/export | trazi Excel sa "Trust access to the VBA project object model" |
| F-060 | otvaranje VBE i `SendKeys` fallback | trazi Excel |
| F-061 | migracija iz starog fajla | trazi Excel i stari `.xlsm`; nema testa koji tvrdi ishod |
| F-062 | brisanje 22 tabele | trazi Excel; nema testa koji tvrdi da se zaglavlja cuvaju |
| F-064 | koje grupe i polja se stvarno iscrtaju | `list:` izvori citaju svesku (`LookupCSV:217`), pa spisak zavisi od podataka -- trazi Excel |
| F-065 | delimicno snimanje i spisak preskocenih polja | trazi Excel; nema testa koji tvrdi bas taj rezime |
| F-066 | pitanje pre odbacivanja nesacuvanog | nema testa koji tvrdi bas to pitanje za panel podesavanja |
| F-067 | pronalazenje `pdftotext.exe` i automatski rezim | trazi disk i `FileDialog` |
| F-068 | folder picker | trazi `FileDialog` |
| F-069 | skrivanje/otkrivanje lista | trazi Excel |
| F-070 | ceo setup lanac i zeleni gate | trazi Excel, disk i Google |
| F-071 | spisak obaveznih kolona koji `CheckRequiredColumnsForSetup` trazi | procedura nije citana u ovoj sesiji; presuda "delimicno" je iz obrasca `modProductionHealthCheck.Check_CoreTablesAndColumns:109`, ne iz njenog koda |
| F-072 | ishod prekidaca rezima | trazi Excel |
| F-073 | pickeri i pravljenje foldera | trazi `FileDialog` i disk |
| F-074 | zivi server link | trazi Google/mrezu |
| F-075 | upis prvog admina | trazi Excel; nema testa koji tvrdi ishod |
| F-076 | odbijanje `EnableAuth` bez aktivnog admina | nema testa koji tvrdi bas to odbijanje |
| F-077 | ukljucenje hesa i SHA dokaz | nema tvrdnje u automatskoj suiti (isto kao **F-038**) |
| F-078 | pravljenje PDF foldera | trazi disk |
| F-079 | ceo self-update lanac (download, delta-skip, faza 1/2, save) | trazi Drive, Excel i pristup VBA projektu; `Test_DeltaSkip:1680` je razvojni makro, ne suite |
| F-080 | oporavak prekinutog azuriranja | trazi registar (`SaveSetting`) i Excel |
| F-081 | DEV self-update iz git klona | trazi Excel i dev masinu |
| F-082 | rollback `current.json` | trazi Drive |
| F-083 | listing objavljenih verzija | trazi Drive |
| F-084 | ishod release kapije | vozi sedam suita koje traze Excel, plus dva rucna GAS koraka |
| F-085 | aktivacija licence | trazi mrezu i licencni endpoint |
| F-087 | ishod 21 od 22 provere | `T_Integritet_...:17573` tvrdi samo B8 i B9; ostale traze Excel i zatecene podatke |
| F-088 | broj u naslovu liste | crta se u ljusci, trazi Excel |
| F-089 | ishod 19 provera zdravlja | trazi Excel; poznato je samo da `Check_CoreTablesAndColumns` danas pada (**AUD-055**), i to nije pokrenuto |
| F-090 | backfill prijemnica hladnjace | trazi Excel; test seam (`BackfillPrijemniceHladnjacaCore` sa `silent`) postoji, ali nije nadjena tvrdnja koja ga vozi u ovoj sesiji |
| F-091 | backfill `GeneracijaID` | trazi Excel |
| F-092 | backfill `BrojOtpremnice` | trazi Excel |
| F-093 | korekcija kolicina na paletama | trazi Excel |
| F-094 | pregled listova i njegova dugmad | trazi Excel |
| F-095 | Drive self-test | trazi Google/mrezu |

## Nalaz uz mapu (za korak posle) -- F2

**1. Provera zdravlja je danas najtvrdji citalac starog modela u celoj
aplikaciji, i to ne posredno nego po spisku.**
`Check_CoreTablesAndColumns:109` NE cita kanon nego nosi RUCNO KUCAN spisak
kolona, i u njemu izricito trazi `Otkup.Kolicina/Cena/Klasa/VozacID/BrojZbirne/
OtpremnicaID`, `Otpremnica.Kolicina/Cena/KolAmbalaze/Klasa`, `Zbirna.
UkupnoKolicina/UkupnoAmbalaze/Klasa` i `Prijemnica.Kolicina/Cena/KolAmbalaze/
Klasa` (`:122-137`). Cim slajs premesti neku od njih na stavke, provera javi FAIL
na ISPRAVNOJ svesci -- sto se vec desilo sa `Isplaceno`/`DatumIsplate` (**AUD-055**,
`docs/KNOWN_ISSUES.md`). Isti obrazac nosi `Check_GoogleSyncMasterSchema:1155`
(`:1170-1174`).

**2. Unutar istog modula vec postoje DVA nacina da se ista stvar izracuna, i
noviji je ispravan.** `Check_OtkupPaymentConsistency:505` i
`Check_KooperantOtkupReconciliation:600` vrednost otkupa uzimaju sa STAVKI
(`HealthVrednostPoOtkupu`), i komentar `:620-622` izricito kaze zasto: "dve klase
legitimno nose dve cene, pa `Kolicina x Cena` sa zaglavlja posle cutover-a daje
nulu za svaki nov otkup -- a zbir po kooperantu bi se i dalje 'slagao', na nuli".
Dakle jedan deo modula je vec presao, a `Check_CoreTablesAndColumns` iz tacke 1
jos tvrdi suprotno. To nije dva posla nego jedan: spisak kolona treba da dodje iz
kanona, kao sto vec radi `Check_SchemaRegistry:68`.

**3. Petnaest od 22 provere integriteta postoji SAMO zato sto postoji stari
model.** Sve koje presudjuju "da" (A1, A2, A3, B1, B2, B3, B4, B5, B5b, B6, B7,
B8, B9, C1/C4, C2) mere posledice tri stvari: poslovni broj kao veza
(`BrojZbirne` u cetiri tabele), linijska polja na zaglavlju (`Kolicina`, `Cena`,
`Klasa`, `KolAmbalaze`) i `Otkup.OtpremnicaID`. Sedam koje presudjuju "ne" (A4,
A5, B10, C3, C5, D1, D2) mere zaglavlje protiv stavki -- i to su tacno palete i
prerada, jedine dve celine koje su vec po novom modelu. Posle refaktora
najmanje ovih petnaest provera nema sta da meri: `Chk_B6_ZbirnaCaseMismatch:593`
(ista rec pisana razlicito u cetiri tabele) i `Chk_B8_DvosmislenBrojZbirne:351`
(isti broj na vise aktivnih zbirnih sa razlicitim kupcem) su po konstrukciji
nemoguci kad je broj kljuc jednog zaglavlja.

**4. `BackfillPrijemniceHladnjaca:439` je najzavisniji makro u F2 i u novom
modelu nema sta da radi u ovom obliku.** Kljuc obrade mu je `BrojZbirne|Klasa`
(dakle red PO KLASI), a izvor su po-klasna polja na zaglavlju otpremnice
(`modAutoHladnjaca.bas:488-493`). Isto vazi za pravilo "druga klasa dobija ISTI
broj prijemnice", koje u novom modelu prestaje da bude pravilo i postaje samo
posledica toga sto je prijemnica jedno zaglavlje. Sposobnost (**F-090**) ostaje:
naknadno napraviti prijemnice za hladnjacke otpremnice koje ih nemaju.

**5. Dva javna backfill makroa pisu bas kolone koje slajs brise.**
`BackfillOtkupBrojOtpremnice:1569` popunjava `Otkup.BrojOtpremnice`, a
`BackfillDeteZbirnaGeneracija:1454` `Zbirna.GeneracijaID`. Oba su
`status=ZIV_MAKRO` sa nula pozivalaca u produkciji (`popis_citalaca`), pa su
ulazne tacke samo za odrzavanje. Za redosled slajseva to znaci da se oni gase
ZAJEDNO sa kolonom, bez zasebnog koraka -- ali i da niko nece primetiti ako
ostanu, jer ih ne zove nijedan tok.

**6. `PaletaAdjust_Prompt:253` vezuje palete na PRIJEMNICU po poslovnom broju.**
`AdjustPaletaGajbiceZaPrijemnicu_TX(brojPrij, ...)` (`modPaletniListUI.bas:218`)
prima broj, ne ID -- iako paletne stavke nose i `PrijemnicaID`
(`Chk_A3_StavkeVsPrijemnica:822` ga koristi). To je jedno od mesta gde su obe
veze zive istovremeno, pa `Chk_D2_PreradaNesvezaPaleta:947` i postoji: da bi
proverio da dve veze pokazuju na isto.

**7. Admin panel i Podesavanja nose TVRDU branu, ali samo oni.** Oba modula
odbijaju ne-administratora i na izgradnji i na akciji (AUD-033,
`modAdmin.bas:57-61` i `:219-222`, `modPodesavanja.bas:257-261`), a
`ShowConfigSheet:801` isto. Ostali makroi ove oblasti nemaju branu: `SetupNewPC`,
`RunSelfUpdate`, `PublishReleaseToDrive`, `RollbackReleaseTo`, `OcistiTabele` i
`MigrirajPodatkeIzStarog` se iz Alt+F8 pokrecu bez provere prava (`OcistiTabele` i
`MigrirajPodatkeIzStarog` imaju samo potvrdu, `PublishReleaseToDrive` sifru kad se
zove iz panela -- ali ne i kad se zove direktno). Isti obrazac je u F1 vec naveden
za `modMain.OpenExcel`/`CloseExcel`. Nije regresija refaktora -- postojece stanje
koje treba presuditi.
