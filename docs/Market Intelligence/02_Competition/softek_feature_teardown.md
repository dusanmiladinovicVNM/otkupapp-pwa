# SOFTEK — feature teardown i poređenje sa AgriX-om

**Status:** Drugi feature-level teardown konkurenta u repou
**Datum analize:** 2026-08-02
**Izvori:**
- `SOFTEK_uputstvp_otkup_poljoproizvoda.pdf` (34 strane, 2017) — korisničko uputstvo za modul „Otkup poljoprivrednih proizvoda"
- `Softek-otkup.pdf` (16 strana, „Verzija 2.1", 2014) — uputstvo na koje sam program linkuje iz menija Pomoć
- **~20 screenshot-ova žive aplikacije `Ver.20.2.4`, poslovna godina 2021** (vidi §8)

**Analiza AgriX strane:** čitanje koda u `src-vba/`, `src/`

---

## 1. Izvor i njegova ograničenja

`PRODUCT DOCUMENTATION EVIDENCE`: korisničko uputstvo za **SOFTEK** modul otkupa
(www.softek.rs). Metapodaci PDF-a: Microsoft Word 2010, **kreirano 2017-06-06**.

`LIMITATION`:

- uputstvo pokriva **samo modul otkupa** (34 strane, pretežno screenshotovi). SOFTEK
  po `competitor_references.csv` ima 13 referenci za „softver za otkup poljoprivrednih
  proizvoda", ali ovo uputstvo ne opisuje ceo proizvod;
- iz sadržaja se vidi da modul stoji **na knjigovodstvenoj platformi** (KEP knjiga,
  kontni nalog za knjiženje, poslovne knjige) — ta platforma nije dokumentovana ovde;
- dokument je iz 2017; aktuelna verzija je šira. Sve „nema" tvrdnje u §§2–7 znače
  **„nije opisano u uputstvu"**, ne „proizvod to nema".

> **Ova ograničenja su delimično prevaziđena.** §8 opisuje živu aplikaciju `Ver.20.2.4`
> na osnovu screenshot-ova i sadrži dve izričite ispravke zaključaka iz §§4 i 6.
> Kada se §§2–7 i §8 razlikuju, **važi §8**.

---

## 2. Zašto je ovaj dokaz važniji od AGROSOFT-a

`INFERENCE`: **SOFTEK je prvi dokumentovani direktni konkurent AgriX-a.** Za razliku od
AGROSOFT-a (žito, silos, laboratorija, zadruge), SOFTEK-ovo uputstvo radi tačno posao
koji AgriX radi:

- radni primer kroz celo uputstvo je **`MALINA VILAMET I KLASA`** i ambalaža
  **`GAJBICA MALINE`** (str. 11);
- obračun neto količine je **bruto − (broj gajbica × težina gajbice)** (str. 11);
- ambalaža se vodi kroz **revers** i ima `ZADUŽENO / RAZDUŽENO / STANJE` na otkupnom
  listu (str. 12, 16);
- uvod objašnjava **PDV nadoknadu od 8%** za poljoprivrednika van sistema PDV-a i
  prag od 8.000.000 RSD (str. 4);
- geografija referenci (Užice, Arilje, Brus, Kosjerić, Prijepolje, Priboj) je
  identična AgriX klasteru.

Drugim rečima: isti kupac, isti proizvod, ista regulativa, isti kraj Srbije.

---

## 3. Funkcionalni obim modula (iz sadržaja)

| Grupa | Stavke |
|---|---|
| Šifarnici | poljoprivredni proizvodi, otkupna mesta, ambalaža, **uslovi plaćanja (rok za uplatu)**, poslovni partneri |
| Otkup | otkup na centralnom otkupnom mestu, štampa otkupnog lista, **štampa nalepnica**, **nalog za knjiženje** |
| Ostala dokumenta | **revers za preuzetu ambalažu** na centralnom mestu |
| Izveštaji | ukupne otkupljene količine (ceo period / period / po poljoprivrednicima), preuzeto robe zbirno/pojedinačno/po datumu, vraćeno-preuzeto ambalaže |
| Kartica | kartica robe (uz **grafički prikaz** kretanja), kartica ambalaže |
| Lager lista | stanje i vrednost robe u skladištu, robe i ambalaže |
| Statistika | otkupljene količine po proizvođačima, otkupljeno u periodu, ambalaža preuzeto/vraćeno, **pivot tabele ambalaže** |
| Finansije | **kartica poljoprivrednog proizvođača** (dugovanje) → dugme **Virman** → nalog za prenos + štampa |
| Poslovne knjige | **KEP knjiga**, automatsko proknjižavanje dokumenata |
| Prečice | toolbar za tri šifarnika |

---

## 4. Matrica poređenja

Legenda: ✅ ima · ⚠️ delimično / drugačije · ❌ nema · „—" nije opisano u uputstvu

### 4.1. Otkup i obračun

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz na AgriX strani |
|---|:--:|:--:|---|
| Otkupni list u dva TAB-a: zaglavlje (datum, otkupno mesto, poljoprivrednik, uslovi plaćanja) → stavke (str. 10–11) | ✅ | ✅ | `frmOtkup` + `modOtkup`, blokovi (`modOtkupBlok`, `clsBlokUI`) |
| **Neto = bruto − (kom ambalaže × jed. težina)**, program računa automatski (str. 11) | ✅ | ✅✅ | `frmOtkup` linije 938–974: isti obračun, uz zaštitu „tara ≥ bruto" (`DOK_MSG_TEZINA_AMBALAZE`) i **odvojeno za Klasu I i Klasu II**; težina iz `tblTipAmbalaze.TezinaGajbiceKg` |
| Fakturna cena = nabavna cena sa PDV-om iz šifarnika (str. 11) | ✅ | ⚠️ | AgriX ima dva modela cene: `tblArtikli.CenaPoJedinici` i append-only `tblCenovnik` (`GetVazecaCena`) — bogatije, sa istorijom |
| Ambalaža na otkupnom listu: `ZADUŽENO / RAZDUŽENO / STANJE` (str. 12) | ✅ | ✅ | `modAmbalaza` ledger, `ReportKarticaAmbalaze` |
| **Šifarnik uslova plaćanja** kao kodirana lista koja se bira na otkupnom listu (str. 8) | ✅ | ⚠️ | AgriX ima `OTKUP_ROK_ISPLATE` i `OTKUP_KLAUZULA` kao **jednu globalnu vrednost** (`modPrint`), ne šifarnik po dogovoru |
| Odvojen prostor šifara: šifre proizvoda i ambalaže se ne smeju poklapati (str. 5) | ✅ | ✅ | odvojene tabele `tblArtikli`, `tblKulture`, `tblTipAmbalaze` |
| Polje „Sistem PDV-a" na partneru (šifra 3 = registrovani poljoprivrednik), JMBG + matični broj gazdinstva odvojeno (str. 9) | ✅ | ⚠️ | AgriX ima BPG i PIB/JMBG polja; status u sistemu PDV-a kao kodirano polje treba proveriti |
| PDV nadoknada 8%, ukalkulisana u cenu (str. 4, 15) | ✅ | ✅ | `CFG_PDV_NADOKNADA_STOPA`, default 8 (`modConfig`) |
| Dve klase na jednom otkupu | — | ✅ | `chkDveKlase` u `frmOtkup` |
| Otkup sa terena preko mobilnog uređaja | — | ✅ | PWA `src/js/features/otkup/`, offline sync |

### 4.2. Dokumenti i štampa

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz |
|---|:--:|:--:|---|
| Štampa otkupnog lista (str. 12) | ✅ | ✅ | `modPrint`, `OTKUP_PRINT_MODE`, PWA `otkup/otkupni-list.js` |
| 🟢 **Štampa nalepnica** — broj nalepnica = `neto / 12 + 1`; na nalepnici: firma, proizvod, datum otkupa, naziv iz šifarnika, poljoprivrednik + trocifrena šifra (str. 13–14) | ✅ | ❌ | AgriX nema štampu nalepnica; „etiketa" u `modPaletniList` je logički identitet palete (RELABEL), ne odštampana nalepnica |
| 🟡 **„Štampaj kao priznanicu"** — isti dokument, drugi ispis (str. 13) | ✅ | ⚠️ | AgriX ima print-mode po dokumentu, ne alternativni pravni oblik istog dokumenta |
| 🟢 **Nalog za knjiženje** sa kontima — 1311 magacin robe, 287 PDV plaćen poljoprivredniku, 435 dobavljači u zemlji (str. 15) | ✅ | ❌ | AgriX nema kontni nalog ni izlaz ka knjigovodstvu |
| **Revers za preuzetu ambalažu**, pravi se **pre početka otkupa** za svakog proizvođača (str. 16–18) | ✅ | ✅ | `modAmbalaza`; AgriX vodi i palete (`tblPaleta`, `modPaletniList`) |
| Otpremnice, zbirne, fakture, SEF | — | ✅ | `modDokumenta`, `modFaktura`, `modSEF*` |

### 4.3. Pregledi, kartice i lager

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz |
|---|:--:|:--:|---|
| Ukupne otkupljene količine: ceo period / period / po poljoprivrednicima (str. 19–21) | ✅ | ✅ | `ReportOtkupRoba`, `ReportOtkupListe`, `ReportProsecnaCena` |
| Preuzeto robe zbirno / pojedinačno / po datumu; vraćeno-preuzeto ambalaže (str. 22) | ✅ | ✅ | `ReportAmbalaza`, `ReportKarticaAmbalaze` |
| Kartica robe — sve promene po artiklu (str. 23) | ✅ | ✅ | `ReportKarticaRobaRekap` |
| 🟡 **Grafički prikaz kretanja proizvoda** uz karticu (str. 24) | ✅ | ❌ | AgriX nema grafikone ni u VBA ni u PWA izveštajima |
| Kartica ambalaže (str. 25) | ✅ | ✅ | `ReportKarticaAmbalaze`, `PrintKarticaAmbalazePDF` |
| Lager lista robe/ambalaže — stanje **i vrednost** u skladištu (str. 26) | ✅ | ⚠️ | `TBL_LAGER` je deklarisan a nekorišćen; stanje se vodi kroz palete/prijemnice, ne kroz jednu lager listu sa vrednošću |
| 🟡 **Pivot tabele** preuzete/vraćene ambalaže (str. 30) | ✅ | ⚠️ | AgriX živi u Excelu, pa je pivot dostupan ručno — ali nije ponuđen kao gotov izveštaj |
| Marža, saldo kupaca, sledljivost | — | ✅ | `modMarza`, `ReportSaldoKupci`, `modSledljivost` |

### 4.4. Finansije i knjige

| Funkcija (SOFTEK, strana) | SOFTEK | AgriX | Dokaz |
|---|:--:|:--:|---|
| Kartica poljoprivrednog proizvođača — obaveze prema dobavljaču (str. 31) | ✅ | ✅ | `ReportKarticaKooperanta`, `PrintKarticaPDF` |
| 🟢 **Virman iz kartice** — jedan klik → nalog za prenos → štampa (str. 31) | ✅ | ⚠️ | AgriX ima jače: `GenerisiNalogeCSV` (CSV za e-banking, poziv na broj = broj bloka) + `PrintIsplataSpecifikacija`; ali **nema štampu pojedinačnog virmana iz kartice** |
| 🟢 **KEP knjiga** — automatsko vođenje, štampa (str. 33) | ✅ | ❌ | AgriX nema KEP ni poslovne knjige |
| Automatsko proknjižavanje dokumenata u poslovne knjige (str. 33) | ✅ | ❌ | AgriX se zaustavlja na fakturi i SEF-u |
| Uvoz bankarskih izvoda i auto-mapiranje uplata | — | ✅ | `modBankaImport` + 4 parsera, `modBankaMapiranje` |
| Avansi, kes isplate, storno lanac | — | ✅ | `ApplyAvansToOtkup_TX`, `CFG_KES_ISPLATE`, `modStorno*` |

---

## 5. Sažetak: šta SOFTEK ima a AgriX nema

1. 🟢 **KEP knjiga i nalog za knjiženje sa kontima** (str. 15, 33). Ovo je najveća
   razlika i ujedno objašnjenje SOFTEK-ovog ugla: otkup je modul na knjigovodstvenoj
   platformi. Kupac koji hoće „i otkup i knjige na jednom mestu" tamo dobija oboje.
2. 🟢 **Štampa nalepnica** sa formulom `neto / 12 + 1` (str. 13–14). Sitno, vidljivo,
   svakodnevno.
3. 🟢 **Virman/nalog za prenos direktno iz kartice proizvođača** (str. 31).
4. 🟡 **Grafički prikaz** kretanja artikla (str. 24) i **pivot tabele** ambalaže (str. 30).
5. 🟡 **Šifarnik uslova plaćanja** umesto jedne globalne klauzule (str. 8).

`INFERENCE`: nijedna od ovih stavki nije arhitektonska. Sve su u dometu AgriX-a; KEP i
kontni nalog su jedini koji traže poslovnu odluku (da li AgriX ulazi u knjigovodstvo
ili se namerno zaustavlja na SEF-u).

## 6. Sažetak: šta AgriX ima a u uputstvu ga nema

Cela desna strana lanca i ceo teren:

| Oblast | AgriX |
|---|---|
| Prodaja | kupci, otpremnice, zbirne, fakture, **SEF** sa state machine-om |
| Teren | PWA za kooperanta, vozača, otkup i management; offline, QR, potpisi |
| Logistika | vozači, dispečer, transport, palete, paletni list, hladnjača, prerada |
| Finansije | uvoz izvoda 4 banke, auto-mapiranje, CSV nalozi, avansi, marža |
| Agronomija | parcele sa GIS poligonima, GGAP, knjiga polja, agromere, karenca, agrohemijski magacin i dug |
| Kontrola | storno centar sa žurnalom i recovery, sledljivost lanca, monitoring, prava po 12 oblasti |
| Isporuka | self-update, licenciranje po uređaju, health check |

`LIMITATION`: uputstvo pokriva samo modul otkupa — SOFTEK kao firma verovatno ima i
fakturisanje i SEF u drugim modulima. Ne tvrditi suprotno bez dokaza.

## 7. Prodajna implikacija

`INFERENCE`:

1. **SOFTEK je referentna tačka za cenu i za „šta je dovoljno".** Kupac koji ga koristi
   ima pokriven otkup, ambalažu, karticu i knjige — dakle nije u bolu zbog osnovnog
   dokumenta. AgriX se ne prodaje protiv otkupnog lista, nego protiv svega što SOFTEK
   modul ne dodiruje: terena, hladnjače, transporta, banke i SEF-a.
2. **Dve male stvari mogu da odluče demo** — nalepnice i „gde su mi knjige". Vredi imati
   spreman odgovor na oba (nalepnice: da/ne/kada; knjige: AgriX izvozi ka knjigovodstvu,
   ne vodi ih).
3. **Preklapanje geografije je potpuno** (Zapadna i Centralna Srbija), pa je verovatnoća
   susreta u istom poslu visoka — viša nego sa AGROSOFT-om.
4. **Prava linija razdvajanja nije funkcija nego to ko sme da radi u programu** (dokaz
   u §8.2). Kod SOFTEK-a otkup se unosi kroz kontni UI — šifra dokumenta, konto
   dobavljača, konto magacina, mesto troška. Kod AgriX-a operater ne mora da zna
   nijedan konto, a kooperant i vozač uopšte ne ulaze u desktop aplikaciju.
   Ovo je jači i pošteniji argument od nabrajanja modula.

## 8. Verzija 20.2.4 uživo — šta screenshot-ovi pokazuju preko uputstava

`PRODUCT DOCUMENTATION EVIDENCE` (screenshot): oko 20 snimaka žive aplikacije,
naslovna traka `OTKUP POLJOPRIVREDNIH PROIZVODA - Zemljoradnicka Zadruga
Ver.20.2.4 - poslovna godina 2021`. Demo firma je podešena na
`Zemljoradnicka Zadruga, Svetog Ahilija bb, 31000 Arilje`.

`LIMITATION`: jedna mašina, demo baza sa dva artikla i jednim otkupnim mestom,
**nemački Windows locale** (statusna traka: `DEU`, datum `2.8.2026.`). Sve što
sledi treba potvrditi na srpskoj produkcionoj instalaciji pre upotrebe u prodaji.

### 8.1. Tehnička arhitektura

| Nalaz | Dokaz |
|---|---|
| **Backend je Microsoft Access (Jet)**, ne server baza | `Datoteka → Compact baze` (Jet „Compact and Repair"); runtime greška `-2147217913 (80040e07)` = Jet OLEDB `Data type mismatch in criteria expression` |
| Klijent je **VB6-era desktop** | format dijaloga `Run-time error '...'` |
| Baza je **po poslovnoj godini** | `Datoteka → Izbor poslovne godine`; godina u naslovnoj traci |
| Održavanje baze je ručno | `Datoteka`: `Izbor poslovne godine · Arhiviranje baze · Compact baze · Kraj rada` — to je ceo meni |
| **Aplikacija nije Unicode** | mojibake kroz ceo meni na nemačkom locale-u: `Pomoæne knjige`, `skraæeni unos`, `proizvoðaèi`, `Vraæeno`, `pojedinaèno` (CP1250 prikazan kao CP1252) |
| **Nema prijave ni prava pristupa** | nijedan od ~20 snimaka ne pokazuje login; nigde nema menija za korisnike/uloge. `LIMITATION`: moguć zaseban administratorski alat koji nije otvaran |

`INFERENCE`: na osi baze podataka poredak je **AGROSOFT (MySQL, klijent-server) >
SOFTEK (Access/Jet) ≈ AgriX (Excel workbook + Google Sheets)**. AgriX ovde nema
prednost i ne treba je tvrditi.

### 8.2. Kontna arhitektura — otkup je vrsta dokumenta u glavnoj knjizi

Ovo je najvažniji strukturni nalaz i ne vidi se ni u jednom uputstvu.

- **Šifarnik dokumenata** je pun kontni katalog: `0` Osnovna dokumenta · `02` Početno
  stanje · `03` Zatvaranje klasa prihoda i rashoda · `04` Predzaključna knjiženja ·
  `05` Dinarski izvod (`051`–`054`, **četiri banke**) · `06` Devizni izvodi (`061`, `062`) ·
  `07` Blagajna (`071`) · `08` Kompenzacije i cesije (`081`, `082`)
- **Otkup na otkupnom mestu = šifra dokumenta `381`**, revers ambalaže = `385`
- dobavljač je vezan za sintetički konto **`4358 – Dobavljaci u zemlji poljoprivrednici`**,
  a „Šif. poljop." u gridu je **`4358100`** — šifra poljoprivrednika **jeste analitički konto**
- kolona **Magacin = `1311`** (isti konto iz naloga za knjiženje u uputstvu, str. 15)
- **otkupno mesto = „mesto troška"** (polja u šifarniku doslovno nose taj naziv)
- svaki dokument se zatvara dugmetom **`Knjiženje`**; dok nije proknjižen, stoji u
  `Pomoćne knjige → Dokumenta koja nisu knjižena`

`INFERENCE`: operater na otkupnom mestu radi u knjigovodstvenom UI-ju — da unese
otkup maline kreće se kroz šifru dokumenta, konto dobavljača, konto magacina i mesto
troška. To je ekran za knjigovođu, ne za sezonskog radnika. **To je najoštrija linija
razdvajanja prema AgriX-u** — jasnija od bilo koje pojedinačne funkcije.
`LIMITATION`: moguće je da su u realnoj instalaciji konta predpodešena pa operater ne
bira ništa; proveriti pre upotrebe kao argument.

### 8.3. Funkcije kojih nema ni u jednom uputstvu

| Funkcija | Gde |
|---|---|
| **Otkup – skraćeni unos** | meni Otkup (ekran nije viđen) |
| **Otkup na ostalim otkupnim mestima** — multi-site | meni Otkup |
| **Otpremnica** | meni Ostala dokumenta — **ali puca pri otvaranju**, vidi §8.5 |
| **Revers za ambalažu na ostalim mestima** | meni Ostala dokumenta |
| **Vrste dokumenata** kao šifarnik | meni Šifarnici |
| Izveštaji **Po otkupnim mestima ▸**, `Lager robe`, `Lager ambalaže` | meni Izveštaji |
| **Štampa IOS-a** (izvod otvorenih stavki), `Štampa svih kartica`, `Saldo veći od` + Dugovni/Potražni, ABC sortiranje | Promet dobavljača — finansijski |
| **Otvorene/zatvorene stavke** (`Prikaz: Sve / Otvr. / Zatv.`), duguje/potražuje/saldo, sortiranje po nalogu, `Virman`, `Grafik` | Stanje finansijske kartice |
| **Specifikacija – promet poljoprivrednih proizvođača** | meni Finansije |
| **BAR kod** kao kriterijum pretrage artikala | Kartica robe |
| Tabovi **`Zalihe`** i **`Oznake`** na artiklu; kolone `Fab.oznaka`, `Proizvodjac` | šifarnik „Materijal" |
| **`Broj palete`** direktno na otkupnom listu (kolona `Paleta`) | Otkup 381 |
| Kolone **`PDV poljop.`** i **`PDV nep.`** — razdvojena nadoknada poljoprivredniku od običnog PDV-a | Otkup 381 |
| Podešavanje firme: **5 tekućih računa**, PDV period (mesečni/kvartalni), vlasništvo (privredno društvo/preduzetnik/budžet), veličina pravnog lica (mikro/malo/nedobitno), `Sistem PDV: U sistemu / Van sistema`, **„Osoba za kontakt na IOS-u"** | Podešavanje radnog okruženja → Firma |
| Per-radna-stanica podešavanja: tabovi **Monitor / Printer / Boja** | isto |
| Uniforman CRUD kostur na svakoj formi: grid + tabovi + `Upiši/Izmeni/Briši/Isprazni polja/Opcije >>` + panel **Pretraži/Sortiraj** (`Pretraži po` · `Deo izraza` · `Sortiraj po`) + `Od–Do datum` + `Izdvoj po datumu` | sve forme |

`INFERENCE` — dve ispravke ranijih zaključaka u ovom dokumentu:

1. **§4.2 i §6:** otpremnica **postoji** u proizvodu (meni Ostala dokumenta), pa tvrdnja
   „nema je" važi samo za uputstva. Ostaje da nije upotrebljiva u ovom snimku.
2. **§4.4:** ocena „AgriX ima jače" za isplate treba razdvojiti. Za **masovnu isplatu**
   AgriX i dalje vodi (`GenerisiNalogeCSV`, poziv na broj = broj bloka, auto-mapiranje
   izvoda kroz `modBankaImport`). Za **knjigovodstveno usaglašavanje** — otvorene stavke,
   IOS, bruto bilans po partnerima — SOFTEK je jasno ispred i AgriX tu nema ekvivalent.

### 8.4. KEP knjiga

Prozor: `OBJEKAT-PRODAJNO MESTO: 1311 Magacin robe po nabavnim vrednostima ZA PERIOD
OD 1.1.2021. DO 31.12.2021.` Kolone: `Datum · Dok. · Broj · Opis promene · Zaduženje
sa PDV-om · Razduženje sa PDV-om`. Polja za unos postoje i za iznose **bez** PDV-a i za
**Uplatu**. Dugme `Štampa KEP`.

Dve napomene:

- knjiga se može **ručno unositi i menjati** (`Upiši / Izmeni / Briši`), iako uputstvo
  kaže da se vodi automatski proknjižavanjem;
- prikazani sadržaj ne odgovara periodu iz naslova — vidi §8.5.

### 8.5. Uočeni defekti

`LIMITATION`: sve viđeno na jednoj mašini sa demo bazom i stranim locale-om. Ovo su
zapažanja za proveru, ne dokazana ponašanja u produkciji.

| # | Defekt | Dokaz |
|---:|---|---|
| 1 | **Otpremnica se ne otvara** — neuhvaćena `Run-time error '-2147217913 (80040e07)' Data type mismatch in criteria expression` stiže do korisnika | snimak greške; korisnik potvrdio da je za otpremnicu |
| 2 | **Zbir meša jedinice mere** — `MALINA I KLASA 104,00` (kg) + `ZELENA SALATA 120,00` (kom.) = ukupno `224,00`, na dva izveštaja | Ukupne otkupljene količine; Otkupljeni proizvodi od poljoprivrednika |
| 3 | **KEP: period u naslovu ≠ sadržaj** — naslov kaže 1.1.2021–31.12.2021, a prikazana su i dva reda iz 2020 (`330,00` i `29.400,00`) i ulaze u zbir `51.330,00` | KEP knjiga |
| 4 | **Prozor kartice robe se zove `Form1`** — zaostalo podrazumevano ime forme | kartica robe |
| 5 | **Mojibake na stranom locale-u** — aplikacija nije Unicode | glavni meni |
| 6 | **In-app pomoć kasni ~11 godina** — program linkuje uputstvo „Verzija 2.1" iz 2014, dok je aplikacija na 20.2.4 | `Softek-otkup.pdf`, metapodaci `D:20140528` |

`INFERENCE`: prva četiri su vidljiva svakom korisniku i ne traže tehničko znanje da se
prepoznaju. Za razliku od arhitektonskih argumenata, ovo su stvari koje kupac sam vidi.
Ne koristiti ih napadački — koristiti ih kao pitanja u kvalifikaciji („da li vam se
dešava da…"), u skladu sa pravilom iz `README.md` §9 da se konkurentske reference ne
kontaktiraju agresivno.

`INFERENCE` (kontrast, ne tvrdnja o kvalitetu): AgriX ima zaseban odbrambeni sloj koji
u ovim snimcima nema pandana — `modLogError`, `modIntegritet`, `modDokumentInvariant`,
`modSchemaGuard`, `modJournaling`, `modStornoRecovery`, `RunProductionHealthCheck`.
To ne dokazuje da je AgriX stabilniji u produkciji; dokazuje samo da postoji sloj koji
greške hvata pre korisnika.

## 9. Šta ostaje da se proveri

1. Puni obim SOFTEK proizvoda — da li postoje moduli za prodaju, fakturisanje i SEF.
   Kontni katalog (`05`–`08`) potvrđuje glavnu knjigu, blagajnu i kompenzacije, ali
   fakturisanje i SEF i dalje nisu viđeni.
2. Cena i model održavanja.
3. Da li postoji mobilna ili web komponenta (ni jedno uputstvo ni jedan snimak je ne pominju).
4. Status 13 referenci iz `competitor_references.csv` — koliko ih danas radi.
5. Da li se SOFTEK pojavio u nekom AgriX poslu → ako jeste, red u `competitive_events.csv`.
6. **Ekran `Otkup – skraćeni unos`** — jedina neviđena otkupna forma; verovatno je
   odgovor na zamerku iz §8.2 (knjigovodstveni UI u sezonskom špicu).
7. Da li se runtime greška na otpremnici (§8.5 #1) reprodukuje na **srpskom locale-u**.
   Ako se ne reprodukuje, uzrok je locale i tvrdnja se svodi na „ne radi van srpskog
   Windows-a" — što i dalje stoji, ali je drugačija tvrdnja.
8. Postoje li korisnici i prava pristupa u zasebnom administratorskom alatu (§8.1).
9. Gde se BAR kod unosi i da li se štampa — to bi promenilo raniju procenu da nalepnice
   iz uputstva (str. 13–14) nemaju barkod.

`DECISION` (predlog, bez ID-ja): SOFTEK se vodi kao **direktan konkurent** — za razliku
od AGROSOFT-a koji je adjacent. Ako se usvoji, treba mu dodeliti ID u
`docs/Master Plan/09_QA_DECISION_LOG.md` i upisati ga u `05_COMPETITION.md`.
