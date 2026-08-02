# AGROSOFT — feature teardown i poređenje sa AgriX-om

**Status:** Prvi feature-level teardown konkurenta u repou
**Datum analize:** 2026-08-02
**Izvor:** `AgroSoft-Korisnicko-Uputsvo.pdf` (161 strana), korisničko uputstvo proizvođača
**Analiza AgriX strane:** čitanje koda u `src-vba/`, `src/`, `gas/` na grani `claude/uputstvo-agrix-poredenje-dlmi5h`

---

## 1. Izvor i njegova ograničenja

`PRODUCT DOCUMENTATION EVIDENCE`: kompletno korisničko uputstvo za programski paket
**AGROSOFT**, proizvođač **„DATA SOFT" Vrbas** (sa naslovne strane: Palih boraca 6,
MB 65849577, PIB 112067720, šifra delatnosti 6201, podrska@datasoft.rs).

`LIMITATION` — starost sadržaja:

- uputstvo opisuje podešavanje za **Windows 7, Vista i XP**;
- primeri kroz ceo dokument koriste sezone **„rod 2011"** i **„rod 2012"**;
- sam PDF je renderovan kasnije (metapodaci: `Skia/PDF m115 Google Docs Renderer`,
  izvorni fajl `AgroSoft Korisnicko Uputsvo.doc`).

`INFERENCE`: sadržaj je verovatno iz perioda **2012–2013**, dok je PDF export noviji.
Aktuelna verzija AGROSOFT-a može imati funkcije kojih u ovom uputstvu nema. Sve
„nema" tvrdnje o AGROSOFT-u u ovom dokumentu znače **„nije opisano u ovom uputstvu"**,
a ne „proizvod to danas nema".

`LIMITATION` — asimetrija dokaza: AgriX strana je proverena u izvornom kodu, AGROSOFT
strana samo u dokumentaciji. Nije viđena instalacija, baza, cenovnik ni ugovor.

---

## 2. Šta je AGROSOFT

Iz uputstva (str. 4–8):

- ciljni segment: **zemljoradničke zadruge, agrokombinati i preduzeća koja skladište
  žitarice i industrijsko bilje**;
- arhitektura: **klijent-server, jedna MySQL baza**, Windows desktop klijent
  (poruke o greškama „MySQL server has gone away", „Could not connect to MySQL server");
- isporuka: integrisan paket ili pojedinačni moduli; postoji **TEST baza** za vežbanje;
- deklarisana ekspertiza proizvođača: automatizacija **protočnih vaga, kolskih
  nagaznih vaga i mernih instrumenata**.

Deklarisani funkcionalni obim (str. 7–8), šest domena:

| # | Domen | Ključne stavke iz uputstva |
|---|---|---|
| 1 | Zadrugarsko poslovanje | sezone, pariteti, uslovi svođenja na JUS, ugovori, zaduživanje kooperanta, analiza, obračun predate robe, praćenje isplate, obračunske liste, proizvodnja po partnerima i po parcelama |
| 2 | Silosno poslovanje | arhitektura silosa, lager liste, praćenje i simulacija eleviranja, kvalitet eleviranja, dispozicije, troškovi sušenja i lagera, ulazne/izlazne transakcije |
| 3 | Finansijsko poslovanje | šifarnik proizvoda, cenovnik, cene i poreske stope, ulazno-izlazna dokumenta, ponderisanje analize, promena vlasništva, kalkulacija troškova lagera |
| 4 | Tehnički sistemi | akvizicija podataka sa nagaznih vaga i mernih instrumenata |
| 5 | Laboratorija | unos podataka za laboratoriju, pregled nalaza, **uslužne laboratorijske usluge za treća lica**, izveštaji |
| 6 | Zaštita i kontrola | korisnici, privilegije, praćenje rada korisnika, oporavak od grešaka, arhiviranje |

`INFERENCE`: ovo **nije** isti proizvod kao AgriX — to je sistem za **žito/uljarice u
silosu**, gde je vrednost u vagi, laboratorijskom obračunu kvaliteta i skladišnoj
usluzi. AgriX je sistem za **voće u hladnjači**, gde je vrednost u otkupnom bloku,
kooperantima na terenu, ambalaži i paletama.

---

## 3. Matrica poređenja po oblastima

Legenda: ✅ ima · ⚠️ delimično / drugačije rešeno · ❌ nema · „—" nije opisano u uputstvu

### 3.1. Merenje i prijem robe

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz na AgriX strani |
|---|:--:|:--:|---|
| Kolska vaga preko serijskog porta (COM, BaudRate 9600/19200, izbor šeme mernog instrumenta) — str. 13–15, 160 | ✅ | ❌ | nema serijske komunikacije u `src-vba/` (nula pogodaka na `MSComm`/`SerialPort`/`COM1`) |
| Dva merenja (dolazak i povratak vozila), automatsko očitavanje težine u trenutku klika — str. 30–32 | ✅ | ❌ | težine se unose ručno; `OTKUP_BRUTO_UNOS` flag, `COL_OTK_BRUTO` (`modConfig`) |
| Ručni unos vage kao fallback za analognu vagu — str. 45–47 | ✅ | ✅ | ručni unos je jedini režim u AgriX-u |
| **Uslužno merenje** (merenje za treće lice bez predaje robe) — str. 43–44 | ✅ | ❌ | nema |
| Vagarska potvrda (štampa) — str. 35, 46 | ✅ | ❌ | nema takvog obrasca u `modPrint` |
| Zaključavanje težine: izmenu težine može samo proizvođač uz pismeni zahtev — str. 42 | ✅ | ⚠️ | AgriX ima storno lanac (`modStorno*`) i žurnal, ali težina je editabilna operateru |

### 3.2. Kvalitet i laboratorija

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz na AgriX strani |
|---|:--:|:--:|---|
| Unos laboratorijske analize po elementima (vlaga, primese, lom, defekt, hektolitar, klijavost, zelena zrna, rastur, analiza) — str. 22, 33–34 | ✅ | ❌ | `modKvalitet.bas` je stub od 6 linija („TODO: Implementierung"); `TBL_KVALITET` deklarisan u `modConfig`, nigde nije korišćen |
| Definisanje elementa obračuna: količina sa koje se skida (P/J/I/E), način skidanja (kg/%/din), „se odbija sa" (K/T/R), PDV po elementu, min/max, **formula** (npr. `(-1)*(1-((100-CEIL(UNOS))/86))`) — str. 106–110, 121–122 | ✅ | ❌ | nema parametarskog obračuna |
| Intervali elemenata — boniteti i penali sa predznakom +1/-1/0, cena na intervalu, generisanje N intervala sa offsetom — str. 111–113, 123 | ✅ | ❌ | nema |
| Svođenje na JUS/SRPS, „novi element kao nova osnovica za JUS" — str. 7, 111 | ✅ | ❌ | nema; AgriX radi sa **klasom** na dokumentu, ne sa svođenjem |
| Setovi elemenata obračuna kao šabloni + prepis na novu robu („Podesi otkup") — str. 20, 119–125 | ✅ | ⚠️ | cenovnik po vrsti/sorti/klasi (`modCenovnik`, append-only) pokriva cenu, ne kvalitet |
| **Reobračun** — masovno osvežavanje svih prijemnica/otpremnica posle izmene elemenata — str. 113–114 | ✅ | ❌ | nema |
| Obračunske liste sa filtriranjem po elementu (npr. vlaga 12,5–14%) i export u Excel/OpenOffice — str. 85–88 | ✅ | ⚠️ | `modIzvestaj` ima bogat set izveštaja, ali bez elemenata kvaliteta; AgriX je ionako u Excelu |

`INFERENCE`: ovo je **najveća funkcionalna razlika**. Kod žitarica cena se izvodi iz
laboratorijskog nalaza kroz parametarski obračun; AgriX taj sloj nema uopšte.

### 3.3. Skladištenje, silos i lager

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz na AgriX strani |
|---|:--:|:--:|---|
| Magacini i objekti sa odgovornim licem — str. 25–26 | ✅ | ⚠️ | `tblStanice` (otkupna mesta) + `modStanicaLock`; nema „odgovornog lica" po magacinu |
| **Ćelije silosa** sa stanjem po ćeliji — str. 26–27 | ✅ | ❌ | nema; AgriX prostorno vodi **palete** (`tblPaleta`, `modPaletniList`, `frmPalete`) |
| Rekapitulacija silosa (ulaz/izlaz/obračunato/stanje, po vlasnicima) — str. 81–83 | ✅ | ⚠️ | `modIzvestaj.ReportSaldoOM`, `ReportOtkupRoba` pokrivaju saldo, ne silos |
| Lager liste sa **rasturom** (procenat, rastur na ulazu/izlazu/oboje, zaliha iz prethodnog meseca, preračun) — str. 114–117 | ✅ | ❌ | nula pogodaka na `Rastur`; `Kalo` postoji samo u `modPaletniList`/`modIntegritet` |
| **Obračun troškova skladištenja i sušenja** sa storniranjem i proverom obračuna — str. 103–106 | ✅ | ❌ | nema; nula pogodaka na `Susenje` |
| Preseci stanja — 6 kombinacija (jedan/svi partner × jedna roba/grupa/sve) — str. 75–76 | ✅ | ⚠️ | izveštaji postoje, ali bez ove matrice parametara |
| Kartica partnera (lager) + analitička kartica robe, materijalna i finansijska — str. 78–80 | ✅ | ⚠️ | `ReportKarticaKooperanta`, `ReportKarticaRobaRekap`, `ReportKarticaAmbalaze` |
| Početna stanja po magacinu, robi, partneru, ceni i datumu — str. 63–64 | ✅ | ⚠️ | početno stanje postoji samo za **dug kooperanta** (`ART_POCETNI_DUG`, `BookPocetniDug`) |
| Hladnjača, palete, paletni list, prerada, gotovi proizvodi | — | ✅ | `modPaletniList`, `modPaletniListUI`, `frmPalete`, `tblPrerada`, `modAutoHladnjaca` |

### 3.4. Skladišna usluga i vlasništvo nad robom

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz |
|---|:--:|:--:|---|
| **Potvrda o skladištenju** (dokaz da je roba primljena na čuvanje) — str. 48–49 | ✅ | ❌ | nema |
| **Ugovor o skladištenju** iz šablona, sa štampom — str. 49–52 | ✅ | ❌ | nema |
| **Prenos vlasništva** — roba se vodi na predavaoca do trenutka prodaje — str. 64–70 | ✅ | ❌ | nema; u AgriX-u otkup odmah prenosi vlasništvo |
| **Kompenzacija** sa preračunom (roba za robu, cene i poreske stope) — str. 66–67 | ✅ | ❌ | nula pogodaka na `Kompenzac` |
| Pregled prenosa vlasništva po tipovima (otkup od poljoprivrednika / prodaja / treća lica / prenos bez otkupa) — str. 68 | ✅ | ❌ | nema |

`INFERENCE`: ceo poslovni model „roba na čuvanju u tuđem vlasništvu" AgriX ne modeluje.
To je standard u silosnom poslovanju i praktično ga nema u voćarskom otkupu, ali je
blokator za svaki silos/zadrugu kao kupca.

### 3.5. Dokumenti otkupa

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz |
|---|:--:|:--:|---|
| Prijemnica sa stavkama, magacinom, vlasnikom robe, brojem ugovora, PDV-om — str. 52–58 | ✅ | ✅ | `modDokumenta`, `tblPrijemnica` |
| Otpremnica sa stavkama i dodatnim podacima — str. 59–62 | ✅ | ✅ | `modDokumenta`, `tblOtpremnica`, + `tblZbirna` (nema pandana kod AGROSOFT-a) |
| „Merenje ne stavlja robu na stanje — to čini prijemnica"; jedna odvaga = jedan dokument — str. 34, 38 | ✅ | ⚠️ | AgriX ima `modDokumentInvariant` i `modIntegritet` kao ekvivalentnu zaštitu |
| **Priznanica / otkupni list**, 4 tipa: Tel-Kel (JUS količine), otkup umanjen za usluge sa i bez ulaza, otkup po tarifiranim cenama po vlazi — str. 57, 70–74 | ✅ | ⚠️ | AgriX ima otkupni list (`modPrint`, PWA `otkup/otkupni-list.js`), ali **jedan** model obračuna |
| Automatsko kreiranje priznanice za partnera u periodu — str. 57–58 | ✅ | ⚠️ | grupni otkup / `GRUPNI_OTKUP_PRINT_MODE` |
| Prikaz usluga po dokumentu (analiza, ulaz, sušenje) — str. 58 | ✅ | ❌ | nema usluga kao stavki na dokumentu |
| Veza dokumenta sa fakturom („dokumenta prebačena / neprebačena u fakture") — str. 77 | ✅ | ✅ | `modFaktura`, `tblFakture`/`tblFakturaStavke` |
| **SEF e-faktura** | — | ✅ | `modSEFClient`, `modSEFMapper`, `modSEFValidator`, `modSEFStatusSync`, `modSEFPersistance` |
| **Ambalaža i reversi** kao ledger | — | ✅ | `modAmbalaza`, `ReportKarticaAmbalaze` |

### 3.6. Ugovaranje proizvodnje i zaduženje kooperanata

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz |
|---|:--:|:--:|---|
| Ugovor o organizovanju proizvodnje: ugovorena površina (ha), BPG, opština, min. cena, datum obračuna, aneks — str. 128–134 | ✅ | ❌ | nema modula ugovora (nula pogodaka na `Ugovor` van bankarskih parsera) |
| **Pariteti**: naturalni odnos 1:X, cena duženja u din i u evrima, način računanja (naturalno / finansijski din / finansijski din po otpremnicama / finansijski eur) — str. 117–119, 135–136 | ✅ | ❌ | nula pogodaka na `Paritet` |
| Zaduženje kooperanta po ugovoru (seme, đubrivo, zaštita → obaveza u robi ili novcu) — str. 7, 143–149 | ✅ | ⚠️ | AgriX zadužuje kroz agrohemijski magacin: `SaveMagacin` (`MAG_ULAZ`/`MAG_IZLAZ`), `GetAgrohemijaDug` sabira **dinarsku** vrednost izlaza — nema naturalnog pariteta ni veze sa ugovorom/površinom |
| Table/parcele u ugovoru (naziv, lokacija, ha, klasa, broj parcele, vlasnik) — str. 137–138 | ✅ | ✅✅ | AgriX ide dalje: `tblParcele`, `modGeoParcele`, GeoJSON poligoni, `parcel-draw.html`, PWA `kooperant/parcele.js`, GGAP polja |
| Isplate uz ugovor: redovna/avansna, procenat i iznos **PDV nadoknade** — str. 138–139 | ✅ | ✅ | `CFG_PDV_NADOKNADA_STOPA` (default 8), avansi `ApplyAvansToOtkup_TX`, `CFG_KES_ISPLATE` |
| Obračun po ugovoru: cena + kurs evra → sumar prijemi / isplate / zaduženja — str. 140 | ✅ | ⚠️ | `ReportKarticaKooperanta` daje saldo, ali bez ugovora i bez kursa |
| Rekapitulacija ugovaranja, kompenzacija ugovora („pokriva ugovore" / „pokriven ugovorima") — str. 141–143 | ✅ | ❌ | nema |
| Pregledi zaduženja: po ugovoru, po vrsti ugovora, po robi, po kooperantu, ukupno — str. 143–152 | ✅ | ⚠️ | dug po kooperantu postoji (`GetAgrohemijaDug`, kartica), ostali preseci ne |
| Ulaz/izlaz u proizvodnju kao dokument sa stavkama — str. 151–153 | ✅ | ❌ | nema |
| **Sezona** kao entitet (godina × sorta × partner, ulazna/izlazna) kao osnova ugovaranja i obračuna — str. 23–25 | ✅ | ❌ | nema sezone kao entiteta |

### 3.7. Šifarnici i cene

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz |
|---|:--:|:--:|---|
| Partneri: pravna lica, fizička lica, računi, „partneri sa nepotpunim podacima" — str. 16–18 | ✅ | ⚠️ | `tblKooperanti`, `tblKupci`, `frmMaticniPodaci` + `modMaticniLookups`; nema posebnog pregleda nepotpunih partnera |
| Jedan partner se automatski pojavljuje u svim listama (dobavljač, kupac, vlasnik, prevoznik) — str. 40 | ✅ | ⚠️ | AgriX razdvaja kooperante, kupce i vozače u zasebne tabele |
| Robe/usluge sa grupama, cenama, PDV-om i poreskim stopama — str. 18–21 | ✅ | ✅ | `tblArtikli`, `tblKulture`, `modCenovnik` |
| Otkupne cene za poljoprivredne proizvode od fizičkih lica kao poseban set — str. 18 | ✅ | ✅ | `tblCenovnik` (append-only, `GetVazecaCena`/`AddCena`) |
| Parcele kao prost šifarnik (samo naziv) — str. 27 | ✅ | ✅✅ | AgriX ima geometriju, površinu, kulturu, GGAP i meteo polja |

### 3.8. Izveštaji

| AGROSOFT (str. 91–103) | AgriX ekvivalent |
|---|---|
| Dnevni promet | `ReportOtkupListe`, dnevni pregledi |
| Spisak prometa po sorti / po kulturi / ukupno po sortama / ukupno po kulturama | `ReportOtkupRoba`, `ReportProsecnaCena` |
| Potvrda o preuzimanju — sorte / kulture (broj_kp, registracija, bruto, tara, neto) | otkupni list, prijemnica |
| Spisak prodaje (otkupa) po sortama / kulturama | `ReportSaldoKupci`, `ReportOtkupRoba` |
| Kumulativ trgovine (ulaz/izlaz po lageru, sorti, partneru) | `ReportZbirni`, `ReportSaldoOM` |
| — | **Marža** po kupcu / OM / ukupno (`modMarza`) — nema pandana u uputstvu |
| — | Kartica ambalaže, specifikacija isplata, sledljivost (`modSledljivost.TraceByZbirna`) |

### 3.9. Administracija i sistem

| Funkcija (AGROSOFT, strana) | AGROSOFT | AgriX | Dokaz |
|---|:--:|:--:|---|
| Korisnici, lozinke, privilegije po aplikaciji/modulu, role — str. 154–159 | ✅ | ✅ | `modAuth`, `tblKorisnici`, 12 oblasti prava, PIN hashing (`docs/UPUTSTVO_KORISNICI.md`) |
| Praćenje rada korisnika, oporavak od grešaka, arhiviranje — str. 8 | ✅ | ✅ | `modJournaling`, `modStornoRecovery`, `modStornoZurnal`, Drive backup (`modDrive`) |
| Podaci o preduzeću, logo, memorandumi — str. 11–12 | ✅ | ✅ | `modPodesavanja`, `tblSEFConfig`, `modDocStyle` |
| **TEST baza** za vežbanje odvojena od produkcije — str. 9 | ✅ | ❌ | nema odvojenog test-režima u aplikaciji |
| Klijent-server, jedna MySQL baza, više radnih stanica — str. 4 | ✅ | ⚠️ | Excel workbook + Google Sheets sinhronizacija (`modMasterSync`, `modStammdatenSync`, `gas/`); `modStanicaLock` rešava konkurentni pristup |
| **Self-update klijenta** | — | ✅ | `modSelfUpdate`, `modRelease`, `docs/SELF_UPDATE.md` |
| **Licenciranje i trial** | — | ✅ | `modLicense`, `modTrial` |
| **Monitoring / health check** | — | ✅ | `modMonitoring`, `modProductionHealthCheck`, `gas/Monitoring.gs` |

### 3.10. Teren i mobilnost

Uputstvo AGROSOFT-a **ne pominje nijednu mobilnu, web ni offline komponentu** — sve je
Windows desktop uz bazu. AgriX ovde nema šta da poredi, samo da nabroji:

- PWA sa četiri role: `kooperant`, `vozac`, `otkup`, `management` (`src/js/features/`);
- offline rad i sinhronizacija (`src/js/services/db.js`, `features/*/sync.js`);
- QR skener, PDF u pregledaču, elektronski potpis (`services/qr.js`, `services/pdf.js`, `ui/signatures.js`);
- dispečer i transport (`management/dispecer.js`, `vozac/transport.js`, `vozac/zbirna.js`);
- knjiga polja, agromere, karenca (`kooperant/knjiga-polja.js`, `kooperant/agromere.js`);
- skeniranje fiskalnog računa u lager inputa (`kooperant/fiskalni.js`, `docs/production-runbook-fiskalni-lager.md`);
- banka: uvoz izvoda i auto-mapiranje, CSV nalozi za prenos (`modBankaImport` + 4 parsera, `modBankaExportPregled`).

---

## 4. Sažetak: gde je AGROSOFT jači

1. **Vaga.** Direktna akvizicija sa kolske vage je jezgro proizvoda, ne dodatak.
   AgriX to nema. Veza sa odlukama: **66** i **234** (integracije sa vagama su u
   scope-u proizvodnje 2027, ali posle radnih naloga, normi, ambalaže i prinosa),
   **67** (standardizuju se samo unapred odobrene vage i senzori).
2. **Parametarski obračun kvaliteta.** Elementi, formule, intervali, boniteti/penali,
   svođenje na JUS, reobračun. AgriX ima samo klasu i cenu po klasi.
3. **Skladišna usluga i tuđa roba.** Potvrda i ugovor o skladištenju, prenos
   vlasništva, kompenzacija, troškovi skladištenja i sušenja, rastur.
4. **Ugovaranje sa paritetima.** Naturalni i finansijski (din/eur) paritet, zaduženje
   po ugovoru, aneksi, kompenzacija ugovora, pregledi zaduženja u četiri preseka.
5. **Baza.** Prava klijent-server RDBMS instalacija sa TEST bazom.

`DECISION` referenca: odluka **242** — žitarice, silosi i mlinovi su „mogući red posle
2027" i **nisu sadašnji prioritet**. Tačke 2–4 su prvenstveno silosne/zadružne funkcije,
pa ovaj teardown **ne otvara** zahtev za razvojem; on utvrđuje granicu segmenta.

## 5. Sažetak: gde je AgriX jači

Funkcije kojih u uputstvu nema uopšte:

| Oblast | AgriX |
|---|---|
| Teren i mobilnost | PWA za kooperanta, vozača, otkup i management, offline, QR, potpisi |
| Elektronsko fakturisanje | pun SEF stack sa state machine-om i statusnom sinhronizacijom |
| Banka | uvoz izvoda 4 banke, auto-mapiranje po jakim ključevima, CSV nalozi za prenos |
| Ambalaža | reversi i ledger ambalaže, kartica ambalaže |
| Parcele | GIS poligoni, GeoJSON, GGAP, knjiga polja, agromere i karenca |
| Sledljivost | lanac otkup → otpremnica → zbirna (`TraceByZbirna`) |
| Storno | storno centar sa žurnalom, impact analizom i recovery-jem |
| Paletizacija | palete, paletni list, prerada, gotovi proizvodi |
| Isporuka softvera | self-update, licenciranje po uređaju, monitoring, health check |
| Marža | `modMarza` — u uputstvu nema izveštaja o marži |

`LIMITATION`: većina ovih stavki nije poređena sa **današnjim** AGROSOFT-om, već sa
uputstvom iz ~2012. SEF (2022+) i PWA po definiciji nisu mogli biti u njemu.

## 6. Segmentna implikacija

`INFERENCE`: AGROSOFT i AgriX se u ovom trenutku **ne takmiče direktno**. AGROSOFT
pokriva žito/uljarice/silos/zadruga; AgriX pokriva voće/hladnjača/kooperant/izvoz.
Preklapanje postoji samo u zajedničkom jezgru: partneri, prijem, dokument, kartica,
isplata, prava korisnika.

Praktične posledice:

1. Ako se AgriX ikad pomeri prema žitaricama (odluka **242**), AGROSOFT postaje
   incumbent, a ulaznica su vaga + obračun kvaliteta + skladišna usluga — ne otkupni
   dokument.
2. Za voćarskog kupca koji dolazi sa AGROSOFT-a (ili sličnog silosnog sistema),
   diskvalifikacioni rizik nije funkcija nego **navika**: očekivaće automatsko
   očitavanje vage i „vagarsku potvrdu".
3. Zadruga koja radi i žito i voće je mešoviti slučaj — tu AgriX ne može da zameni
   ceo sistem, samo voćarski deo. To treba reći otvoreno u kvalifikaciji.

## 7. Šta ostaje da se proveri

1. Aktuelna verzija AGROSOFT-a — da li danas ima web/mobilni klijent, SEF i e-fakturu.
2. Reference i broj instalacija (DATA SOFT Vrbas nije u `competitor_references.csv`
   ni u `infosys_agro_references.csv` — trenutno nula referentnih redova).
3. Cenovnik i model održavanja.
4. Da li se AGROSOFT ikad pojavio u AgriX prodajnom procesu — ako jeste, ide u
   `competitive_events.csv` kao founder-confirmed događaj.
5. Da li ijedan postojeći ili ciljni AgriX kupac drži žito pored voća — to bi
   pretvorilo ovaj teardown iz segmentne granice u konkurentski dodir.

`DECISION` (predlog, bez ID-ja): AGROSOFT se vodi kao **adjacent competitor**, ne kao
replacement target, dok se tačke 1–5 ne popune. Ako se usvoji, treba mu dodeliti ID u
`docs/Master Plan/09_QA_DECISION_LOG.md` i upisati ga u `05_COMPETITION.md`.
