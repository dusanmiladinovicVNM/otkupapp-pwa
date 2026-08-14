# PLAN STAVKI ZA REŠAVANJE

> **Šta je ovo:** jedno trajno mesto za stavke koje treba uraditi — bez obzira na to
> iz kog chata/sesije su došle. Rad se seli iz chata u chat, kontekst se gubi; ovaj
> fajl je ono što ostaje. Vlasnik ga popunjava ručno, a **svaka sesija koja radi po
> planu upisuje taj plan ovde pre nego što se završi.**
>
> **Ovo NIJE:** registar bugova (`KNOWN_ISSUES.md`), arhitektonski roadmap
> (`ROADMAP.md`), zatvoren AUD/RF program sanacije (`PLAN_SANACIJE.md` +
> `REFAKTOR_PLAYBOOK.md`), niti storno-specifičan backlog (`STORNO_BACKLOG.md`).
> Ako stavka pripada nekom od njih — živi tamo, a ovde stoji samo red sa linkom.

**Poslednja izmena:** 2026-08-14 · **Grana:** `claude/plan-stavki-resavanje-jw2kgc`

---

## 1) Pravila vođenja

Kratka, jer se dugačka ne poštuju.

1. **Stavka dobija ID i ne gubi ga.** `PS-NNN`, redom. ID se ne reciklira ni kad se
   stavka zatvori — zatvorene idu u §5, ne brišu se.
2. **Svaka stavka ima „Definiciju gotovog".** Bez toga se ne zna kad je gotova, pa
   ostaje otvorena zauvek. Definicija mora biti provera koja može da padne
   (test, `vba_check` izlaz, merenje) — ne „pregledano".
3. **Kontekst se piše kao provereno / neprovereno.** Ono što je izmereno (grep,
   test, brojevi) stoji sa komandom kojom je dobijeno. Pretpostavka se označava sa
   `[pretpostavka]`. Nikad se pretpostavka ne upisuje kao nalaz — po CLAUDE.md §0.
4. **Sesija koja radi po planu upisuje svoj plan u §4 (Dnevnik) pre kraja.**
   Naslov = grana. Sadržaj = šta je planirano, šta urađeno, šta ostalo i zašto.
   To je ono što preživljava gašenje chata.
5. **Stavke se ne rešavaju u ovom fajlu.** Ovde stoji plan i status; kod, testovi i
   release beleške idu na svoja mesta.
6. **Redosled ne diktira ovaj fajl.** Prioritet je kolona, a šta se radi sledeće
   bira vlasnik.

**Legenda statusa:** 🔴 otvoreno · 🟡 u radu / delimično · ✅ zatvoreno ·
⏸️ odloženo (sa razlogom) · ❔ traži odluku vlasnika

---

## 2) Registar

| ID | Stavka | Status | Prio | Oblast | Izvor |
|---|---|---|---|---|---|
| [PS-001](#ps-001--ujednačavanje-ensure-familije) | Ujednačavanje `Ensure*` familije | 🔴 | P2 | VBA / setup | vlasnik, 2026-08-14 |
| [PS-002](#ps-002--ujednačavanje-testova) | Ujednačavanje testova | 🔴 | P2 | testovi | vlasnik, 2026-08-14 |
| [PS-003](#ps-003--logo-i-logotip--poliranje) | Logo i logotip — poliranje | 🔴 | P3 | brend / UI | vlasnik, 2026-08-14 |
| [PS-004](#ps-004--tekstovi-koje-korisnik-vidi-težište-storno) | Tekstovi koje korisnik vidi (težište: storno) | 🔴 | P2 | UX / poruke | vlasnik, 2026-08-14 |
| [PS-005](#ps-005--moderna-struktura-koda-slojevi-api-servisi) | Moderna struktura koda (slojevi, API, servisi) | 🔴 | P1 | arhitektura | vlasnik, 2026-08-14 |

---

## 3) Otvorene stavke

### PS-001 — Ujednačavanje `Ensure*` familije

**Status:** 🔴 otvoreno · **Prioritet:** P2 · **Oblast:** VBA / setup / šeme

**Zadatak (kako je postavljen):** provera da li su svi `Ensure` ujedinjeni i
unificirani u meri u kojoj je to moguće i korisno.

**Kontekst — provereno** (`grep -rhoE "^(Public |Private )?(Sub|Function) +Ensure[A-Za-z0-9_]*" src-vba/`):

- **62 različita `Ensure*` imena**, 63 definicije (`EnsureRuntimeControls` postoji
  dvaput), raspoređena u **17 fajlova** (`modSetup`, `modOtkupUI`, `modAgrohemija`,
  `modBankaImport`, `modStammdatenSync`, `modPaletniList`, `modPrint`,
  `modSEFPersistance`, `modStornoContext`, `modMalina`, `modMouseWheel`,
  `modWindow`, `frmStammdaten`, `frmDokumenta`, `frmIzvestaj`, `frmBankaImport`,
  `frmBankaExportPregled`).
- Četiri jasne familije: **šema/kolone/tabele — 18**, **UI paneli i kontrole — 20**,
  **šabloni dokumenata (`*Sablon`) — 13**, **folderi — 5**; ostatak (6) je
  pojedinačan (`EnsurePoruke`, `EnsureHook`, `EnsureArtikalPocetniDug`,
  `EnsureSEFDocumentIdTextFormat`, `EnsurePrijemnicaNotAlreadyPaletized`,
  `EnsureVozacMirrorForStanica`).
- **`*Core` parovi NISU duplikat** — to je namerni obrazac: javni `Ensure*` sa
  `MsgBox` potvrdom za ručno pokretanje (`Alt+F8`) + `Ensure*Core` radnik koji zovu
  `modSetup`/`modMigracija` bez dijaloga (`modSetup.bas:1007,1018,1026,1050,1067`).
  Ovo treba **dokumentovati kao konvenciju**, ne spajati.

**Konkretni kandidati za spajanje** (telo procedura još nije čitano — prvi korak je
čitanje, ne izmena):

- `modSetup.EnsureFolder` (`:1718`, Private) vs `modBankaImport.EnsureFolderExists`
  (`:1153`, Private) — dve privatne implementacije istog posla u dva modula.
- `modStornoContext.EnsureTable` (`:344`) vs `modSetup.EnsureDataTable` (`:1580`) —
  proveriti da li je semantika ista pre bilo kakvog spajanja.
- 13 × `Ensure*Sablon` — kandidat za jedan tabelom vođen helper `[pretpostavka]`,
  potvrditi tek pošto se uporede tela; šabloni se mogu razlikovati više nego što
  imena sugerišu.
- 3 × `Ensure*TabsBestEffort` (`Stammdaten`, `MgmtReport`, `Kartice`).
- Kolone: `EnsureColumnOnTable`, `EnsureAktivanColumn`, `EnsurePreradaCol`,
  `EnsurePreradaCols` — familija oko iste operacije.

**Šta uraditi:**

1. Inventar u tabelu: ime → fajl → linija → familija → potpis → ko zove.
2. Za svaku familiju odlučiti: **spojiti / ostaviti + dokumentovati konvenciju /
   preimenovati radi doslednosti**. Odluka se zapisuje uz stavku, i za „ostaviti".
3. Spajati **samo** gde je telo dokazano isto ili razlika nebitna. `reuse > new`, ali
   i `minimal change over idealized redesign` — 62 poziva nije poziv na veliki refaktor.
4. Konvenciju (`Ensure*` javni + `*Core` radnik, gde ide šta) upisati u
   `.claude/rules/podaci-i-config.md`.

**Definicija gotovog:**

- Inventar postoji i pokriva svih 62.
- Svaka familija ima zapisanu odluku (uključujući „ostaje kako jeste, evo zašto").
- Za svako stvarno spajanje: `python tools/vba_check.py` zelen **i** dokaz u oba
  smera po CLAUDE.md §5 (namerno pokvariti → pokazati crveno po imenu → vratiti →
  zeleno). Bez oba smera nije gotovo.
- Ponašanje šema/setup-a ne sme se promeniti: `RunAllTests` + relevantne suite na
  Windows mašini. **U Linux/web sesiji se ovo ne može verifikovati** i tako se prijavljuje.

**Rizik:** `Ensure*` diraju šemu tabela, a šema je izvor istine i drifta po
instalaciji (CLAUDE.md §4). Spajanje koje „pojednostavi" upis po poziciji umesto po
imenu kolone je regresija, ne čišćenje. Vidi i `modStornoContext.bas:46` — redosled
tamo MORA da prati `EnsureStornoVezeSchemaCore`.

---

### PS-002 — Ujednačavanje testova

**Status:** 🔴 otvoreno · **Prioritet:** P2 · **Oblast:** test suite

**Zadatak:** ista provera kao PS-001, ali za testove.

**Kontekst — provereno:**

- **15 test modula**, u dva nekompatibilna imenovanja: sufiks `mod*Tests` (9:
  `modAgrohemijaTests`, `modBusinessFlowProTests`, `modFakturaTests`,
  `modGoogleSyncSmokeTests`, `modIzvestajTests`, `modLicenseTests`,
  `modMonitoringTests`, `modNovacTests`, `modSEFTests`) i prefiks `modTest*` (6:
  `modTest`, `modTestBanka`, `modTestMode`, `modTestPalete`, `modTestStorno`,
  `modTestStornoCentar`). `modTestMode` je verovatno infrastruktura a ne suite —
  proveriti pre nego što uđe u bilo kakvu klasifikaciju.
- **~20 različitih assert helpera** preko modula za isti posao: `AssertTrue` (4×),
  `AssertEquals` (3×), `ChkEq` (3×), `ChkEqD` (3×), `Chk` (3×), `AssertEq` (2×), plus
  po-modulu varijante (`AssertNovacTrue`, `AssertFakturaTextEquals`,
  `AssertNovacDoubleEquals`, `AssertDoubleNear`, `AssertContains`…). `modTest.bas`
  već ima javne `AssertEq` / `AssertSnapshot` / `DumpKontrole` (`:406,432,458`).
- **Katalog `SUITES` u `tools/run_vba.py:58` ima 18 ulaza**, a u izvoru je **27
  `Run*` ulaznih tačaka** u test modulima. Deo razlike je legitiman (pod-runneri
  `RunBusinessFlowPro*Only`, SEF pod-suite pod `RunSEFTestSuite`, privatni `RunOne`).
- **Potvrđena rupa:** `RunHttpUtilsSmokeSuite` je `Public Sub`
  (`modSEFTests.bas:2294`) i **ne postoji ni u `SUITES` ni bilo gde u `.claude/`** —
  znači ne pokreće ga nijedan set, ni pun ni brzi. Postoji, a nikad se ne izvrši.
- **3 blind suite** (`gate: False` — rezultat samo u Immediate, runner ih ne vidi kao
  crvene): `RunNovacSmokeSuite`, `RunProductionHealthCheck`, `TestMonitoring_All`.
  Prevođenje blind → gate je opisano u `.claude/rules/testovi.md` §3.

**Šta uraditi:**

1. Inventar: modul → suite → `gate`/`blind` → u podrazumevanom setu? → u katalogu?
2. Zatvoriti rupe: svaka javna suite je ili u `SUITES`, ili obrisana, ili
   eksplicitno označena kao ručna sa razlogom. Počevši od `RunHttpUtilsSmokeSuite`.
3. Blind → gate za 3 suite gore, po postupku iz `testovi.md` §3.
4. Jedan skup assert helpera (`modTest` je prirodan domaćin — već je javni), ostali
   se svode na njega. Ovo je najveći deo posla i ide postepeno, modul po modul.
5. Odabrati jedno imenovanje modula i zapisati ga; **preimenovanje modula je skupo**
   (menja `Attribute VB_Name`, import listu, katalog) — moguće je da je ispravna
   odluka „ostaje kako jeste, novi moduli idu po X". To je validan ishod.

**Definicija gotovog:**

- Nema javne suite van kataloga bez zapisanog razloga.
- Nula `gate: False` u podrazumevanom setu, ili zapisan razlog za svaki.
- Za svaku izmenu helpera: dokaz u oba smera — `tools/sabotaza.py` ili ručno
  kvarenje, pa pokazano da baš ta provera pukne **po imenu**, pa povratak i zeleno.
  Suite zelena nad ispravnim kodom bez pokazane crvene ne dokazuje ništa (CLAUDE.md §5).
- `python tools/run_vba.py` na Windows mašini — pun set zelen.

**Napomena:** ovo je stavka koja se ne može verifikovati u web sesiji (traži
Windows + Excel + `pywin32`). Planiranje i inventar mogu; izmene idu na Windows.

---

### PS-003 — Logo i logotip — poliranje

**Status:** 🔴 otvoreno · **Prioritet:** P3 · **Oblast:** brend / UI / štampa

**Zadatak:** ažuriranje odnosno poliranje logoa i logotipa.

**Kontekst — provereno:**

- **PWA:** `img/AgriX-Otkup-Logo-Final.png` (koristi se, `index.html:58` + keširan u
  `sw.js:86`), `img/AgriX-Logo-Final_Novi.png` i `img/AgriX-Gazdinstvo-Logo-Final.png`
  (u repou; nijedan grep ih ne nalazi u `index.html`/`sw.js`/`src/js` — proveriti da
  li su mrtvi ili se koriste dinamički).
- **Ikone:** `icons/icon-192x192.png`, `icon-256x256.png`, `icon-512x512.png`.
  `manifest.json` upisuje samo 192 i 512; `256` nije referenciran. Apple touch ikona
  pokazuje na 192 (`index.html:22`).
- **Excel/štampa:** logo na dokumentima ide kroz `modDocStyle.DocLogoPath` (`:46`) →
  config `SELLER_LOGO_PATH`, pa fallback `<workbook>\logo.png` / `logo.jpg`;
  crtanje u `DocDrawLogo` (`:61`), poziv iz zaglavlja (`:114`). Podešavanje je
  izloženo korisniku u `modPodesavanja.bas:68`.
- Znači: **tri nezavisna izvora logoa** (PWA `img/`, PWA ikone, Excel `SELLER_LOGO_PATH`)
  koji se danas ne usklađuju automatski.

**Otvorena pitanja za vlasnika:** ❔

- Da li poliranje znači **novi dizajn** ili **tehničko sređivanje postojećeg**
  (rezolucije, maskable ikona, prozirnost, konzistentne margine)? Obim se bitno razlikuje.
- Ostaju li tri varijante loga (Otkup / Gazdinstvo / opšti AgriX) ili se svodi na jednu?

**Šta uraditi (tehnički deo, nezavisno od dizajnerske odluke):**

1. Popisati sva mesta prikaza: PWA header, splash/instalacija, `manifest.json`,
   `sw.js` keš lista, `install/`, štampani dokumenti, forme.
2. Odlučiti sudbinu nereferenciranih fajlova (`AgriX-Logo-Final_Novi.png`,
   `AgriX-Gazdinstvo-Logo-Final.png`, `icon-256x256.png`) — koristi se ili se briše.
3. `manifest.json`: dodati `maskable` varijantu ako je nema (Android adaptive ikone).
4. Pri svakoj zameni fajla: **bump keša u `sw.js`**, inače stari logo ostaje kod
   korisnika koji već imaju instaliranu PWA.

**Definicija gotovog:** ista slika na svim površinama (PWA, instalirana ikona,
štampani dokument); nijedan referenciran fajl ne nedostaje; posle deploy-a hard
refresh pokazuje novi logo (ručna provera — ovo se ne testira automatski, pa ide
kao checklista, u skladu sa CLAUDE.md §6).

---

### PS-004 — Tekstovi koje korisnik vidi (težište: storno)

**Status:** 🔴 otvoreno · **Prioritet:** P2 · **Oblast:** UX / poruke

**Zadatak:** prolazak kroz tekstove koje korisnik vidi, pogotovo storno deo.

**Kontekst — provereno:**

- **623 poziva `Poruka("KLJUC")`** u `src-vba/` — infrastruktura za lokalizovan
  tekst sa dijakritikom (`modPoruke.UpsertPoruke`) postoji i široko se koristi.
- **Storno moduli i dalje imaju direktan `MsgBox`:** `modStornoFlow` 5×,
  `modStorno` 2×, `modStornoRecovery` 2×, `modStornoContext` 1×, `modStornoImpact` 1×
  (= 11 ukupno; `modStornoWarm` i `modStornoZurnal` su čisti). Svaki takav tekst je
  ASCII bez dijakritike i van sistema poruka.
- Storno je najosetljiviji deo za formulaciju: korisnik u tom trenutku poništava
  dokument koji je već negde odštampan/poslat, pa poruka mora jasno reći **šta je
  poništeno, šta nije, i šta sledi**.
- Povezano: `ROADMAP.md` §2.6 već traži izmeštanje `MsgBox`-a iz poslovnih modula u
  UI sloj. **Ista izmena rešava i jedno i drugo** — ne raditi dvaput.

**Šta uraditi:**

1. Popisati sve korisnički vidljive tekstove u storno putanji: 11 `MsgBox`-eva +
   ključevi `Poruka()` koje storno koristi + tekstovi u `frmDokumenta` overlay-u.
2. Za svaki: da li je tačan, razumljiv operateru (ne programeru), i da li kaže šta
   je sledeći korak. Poruka „Greska u X: <Err.Description>" nije poruka korisniku.
3. Preseliti tekstove u `modPoruke` + `Poruka("KLJUC")` — time dobijaju dijakritiku
   i jedno mesto izmene. **VBA izvor ostaje 100% ASCII** (CLAUDE.md §4); dijakritika
   isključivo kroz `UpsertPoruke`/`ChrW`.
4. Gde je `MsgBox` u poslovnom modulu — dići grešku/vratiti rezultat, a prikaz
   ostaviti UI sloju (zajedno sa `ROADMAP.md` §2.6).
5. Posle storna proći isti prolaz za ostatak aplikacije (otkup, faktura, banka, sync).

**Definicija gotovog:**

- Nijedan `MsgBox` u `modStorno*` poslovnim modulima (osim eksplicitno opravdanih,
  sa upisanim razlogom).
- Svi storno tekstovi idu kroz `Poruka()`; `python tools/vba_check.py` zelen
  (ASCII disciplina se time i proverava).
- Test u `modTestStorno`/`modTestStornoCentar` koji tvrdi da poslovna funkcija
  **vraća grešku** umesto da prikaže dijalog — jer izmena menja ponašanje, a
  izmena ponašanja nosi test, ne checklistu (CLAUDE.md §6).
- Sam tekst poruka je stvar procene i ide na checklistu — to je legitiman izuzetak.

**Napomena:** `STORNO_BACKLOG.md` već ima P2/P3 UX stavke za storno centar (inline
potvrde umesto modala, baner „ISPRAVKA u toku"). Ovde se radi **tekst**; te stavke
ostaju tamo i ne prepisuju se.

---

### PS-005 — Moderna struktura koda (slojevi, API, servisi)

**Status:** 🔴 otvoreno · **Prioritet:** P1 · **Oblast:** arhitektura

**Zadatak:** rad na tome da kod dobije modernu strukturu (layeri, API-ji, servisi itd).

**Kontekst — provereno:**

- **PWA (`src/`) je već slojevit** i to je dobra osnova: `services/` (`api`, `auth`,
  `db`, `pdf`, `pwa`, `qr`), `features/` po ulogama (`kooperant`, `management`,
  `otkup`, `vozac`), `ui/`, `utils/`, `state.js`, `config.js`.
- **Ali `index.html` ima 2.932 linije** — monolit pored uredne `src/` strukture.
  Prvo pitanje je koliko je logike ostalo u njemu naspram `src/js`.
- **VBA (`src-vba/`) je ~103.000 linija**, sa modulima koji su prerasli granicu:
  `modOtkupUI` 6.194, `modMasterSync` 4.062, `modDokumenta` 3.833, `modIzvestaj` 3.443,
  `modPaletniList` 3.106, `modPrint` 3.017, `modBankaMapiranje` 2.927.
  (`frmDokumenta.frm` 6.062 je već zaveden u `STORNO_BACKLOG.md` P3.)
- **Slojevitost delom postoji** i u VBA: `modDataAccess` kao sloj podataka,
  `modConfig` za `TBL_*`/`COL_*`, `modPoruke` za tekst, `modDocStyle` za prezentaciju.
  Nije greenfield — postoji obrazac koji treba dosledno primeniti, ne izmisliti.

**Zašto ovo traži odluku pre rada:** ❔ ovako postavljena, stavka je veća od svih
ostalih zajedno i može da se razlije. Traži se **odluka o obimu** pre prvog reda koda:

- (a) **Samo PWA** — istanjiti `index.html`, ostatak u `src/js` po postojećoj šemi.
  Najniži rizik, najbrži vidljiv efekat, ne dira poslovnu logiku desktopa.
- (b) **Samo VBA** — dosledan sloj `UI → servis → podaci`, počev od najvećih modula.
  Najveća vrednost, ali svaka izmena traži Windows + Excel za verifikaciju.
- (c) **Oba, fazno** — realno, ali samo ako se drži pravila „jedan paket po sesiji"
  iz `PLAN_SANACIJE.md` §2.

**Šta uraditi (nezavisno od izbora):**

1. Zapisati ciljni model slojeva **za postojeći kod**, ne idealizovan: šta sme da
   zove šta, gde ide validacija, gde `MsgBox`, gde pristup tabelama.
2. Razbijati **samo modul koji se ionako dira**, uz merljiv razlog. Refaktor bez
   povoda je najskuplji način da se uvede regresija.
3. Svaki paket = jedan release-kandidat, sa testom ponašanja pre i posle.

**Definicija gotovog:** ovo je programska stavka i **nema jedno „gotovo"** — deli se
na pakete, svaki sa svojom definicijom. Prvo gotovo je **zapisan ciljni model +
odabran obim (a/b/c)**.

**Preduslov:** pročitati `REFAKTOR_PLAYBOOK.md` i `PLAN_SANACIJE.md` §2 pre početka —
tamo je već definisan način vođenja ovakvih paketa (minimal-delta, schema-first,
re-baza pre svakog paketa). Ne praviti paralelni proces.

---

## 4) Dnevnik po sesijama / chatovima

> Ovde svaka sesija ostavlja svoj plan, da ne nestane sa chatom. Format:
> grana → datum → šta je planirano → šta urađeno → šta ostalo i zašto.
> **Piše se pre kraja sesije, ne posle.**

### `claude/plan-stavki-resavanje-jw2kgc` — 2026-08-14

**Planirano:** napraviti trajni fajl za plan stavki i uneti prvih pet stavki vlasnika.

**Urađeno:**

- Provereno da ekvivalent ne postoji: `backlog/` su sirovi dumpovi iz starih chatova
  (nisu održavani), `ROADMAP.md` je arhitektonski, `PLAN_SANACIJE.md` je zatvoren
  AUD/RF program, `STORNO_BACKLOG.md` je usko storno, `KNOWN_ISSUES.md` je registar
  bugova. Novi fajl je opravdan; granice prema tim fajlovima upisane u zaglavlje.
- Uneto PS-001…PS-005, svaka sa provereno izmerenim kontekstom (inventar `Ensure*`,
  inventar testova i `SUITES` kataloga, mapa logoa, brojanje `MsgBox`-eva u storno
  modulima, veličine modula) i definicijom gotovog.
- `PLAN_STAVKI.md` upisan u `CLAUDE.md` §1 (reference-first) da bi ga svaka sledeća
  sesija videla na startu — to je ono što ga čini trajnim.

**Nalazi uz put** (nusproizvod inventara, nisu deo zadatka):

- `RunHttpUtilsSmokeSuite` (`modSEFTests.bas:2294`) je javna suite koja nije ni u
  `tools/run_vba.py` `SUITES` ni bilo gde u `.claude/` — ne pokreće je nijedan set.
- `icon-256x256.png`, `AgriX-Logo-Final_Novi.png`, `AgriX-Gazdinstvo-Logo-Final.png`
  nemaju referencu u `index.html`/`sw.js`/`src/js`.
- `*Core` parovi u `modSetup` su namerni obrazac (javni sa `MsgBox` + radnik bez
  dijaloga), a nigde nisu dokumentovani kao konvencija.

**Ostalo:** ništa iz ove sesije. Same stavke PS-001…PS-005 su nedirnute — ovo je
sesija koja pravi plan, ne izvršava ga. PS-003 i PS-005 čekaju odluku vlasnika (❔).

**Verifikacija:** izmena je isključivo dokumentaciona (`docs/`, `CLAUDE.md`) — nema
promene u `src-vba/` ni `src/`, pa nema šta da se testira ponašanjem.
`python3 tools/vba_check.py` je ipak pokrenut i zelen (izlaz u odgovoru sesije).

---

## 5) Zatvoreno (arhiva)

> Zatvorene stavke se premeštaju ovde sa datumom, granom i kratkim dokazom.
> Ne brišu se — plan je i istorija, ne samo TODO.

_(prazno)_

---

## 6) Šta gde živi (da se ne duplira)

| Tražiš | Idi u |
|---|---|
| Aktivan bug / prihvaćen rizik | `docs/KNOWN_ISSUES.md` |
| Arhitektonski roadmap, post-launch hardening | `docs/ROADMAP.md` |
| AUD nalazi → RF paketi, program sanacije | `docs/PLAN_SANACIJE.md`, `docs/REFAKTOR_PLAYBOOK.md`, `docs/AUDIT_FM_TRIJAZA.md` |
| Storno backlog (P1–P3, ADR veze) | `docs/STORNO_BACKLOG.md`, `docs/adr/` |
| Šta dokumenti jesu, invarijante, ko piše koju tabelu | `docs/DOMEN/README.md`, `docs/DOMEN/WHO_WRITES.md` |
| Stanje UI migracije na `frmOtkupUI` | `docs/UI_MIGRACIJA_KATALOG.md` |
| Katalog test suite-ova, `gate` vs blind | `.claude/rules/testovi.md` |
| Pravila po oblastima koda | `.claude/rules/` (tabela u `CLAUDE.md` §3) |
