# AgriX Test Platform Foundation (Faza 1) — kapije, identitet run-a, lifecycle

> **Šta ovaj dokument rešava.** Do sada je pitanje bilo „ima li AgriX dovoljno
> testova?" — ima ih, ~1050 provera. Pitanje je postalo drugo:
>
> **„Da li proces garantuje da se ti testovi zaista izvrše nad tačno onim kodom
> koji se merge-uje i isporučuje?"**
>
> `vba-v2.40.0` je odgovor na to pitanje dao u praksi: tag je nastao, a release
> notes su pošteno zapisale da behavior runner nije ni pokrenut. Pošteno — i
> prekasno, jer je tag već postojao. Uzrok nije bila nepažnja nego **redosled**:
> `release.ps1` je verziju, commit, tag i push radio pre nego što operateru
> uopšte kaže da uradi Import i Compile.
>
> Ovaj dokument opisuje sloj koji to zatvara. Detalji o pojedinačnim suite-ovima
> ostaju u `.claude/rules/testovi.md`; ovde je *proces*.
>
> **Ovo je Faza 1, ne završena unifikacija.** Skelet platforme (kapije, manifest,
> identitet run-a, lifecycle API) stoji; migracija suite-ova na zajednički
> lifecycle je tek počela — jedan od jedanaest. Šta konkretno nije urađeno stoji
> u §11 i nije popis želja nego popis duga.

---

## 1. Tri kapije

| Kapija | Kada | Komanda | Šta mora da prođe |
|---|---|---|---|
| **DEV** | namerno, posle izmene koja dira ponašanje | `python tools\run_vba.py` | STATIC + brzi determinističan Excel set (11 suite-ova, ~1050 provera) |
| **PR** | pre merge-a u `main` | `python tools\run_vba.py --gate pr` | sve iz DEV + `RunNovacSmokeSuite`, COUNTS, CLEANUP izveštaj |
| **RELEASE** | pre taga | `powershell -File tools\release.ps1 <verzija>` | sve iz PR + **dokazan** compile + fixture provenance + external ugovori (PASS/NOT_RUN/WAIVED) |

Google/SEF live test ne blokira svaku izmenu boje dugmeta. Ali release **tačno
zna** da li je external ugovor `PASS`, `NOT_RUN` ili `WAIVED` — i to piše u
poruci taga.

Katalog koje suite ulaze u koju kapiju je **`tests/suite_manifest.json`**, ne
kod. Nova suite ulazi u kapiju time što je upisana tamo.

**Nijedna od ove tri kapije nije automatika.** Stop hook je uklonjen u #191 sa
merenim razlogom (Excel na svakom zaustavljanju, i na turnovima bez ijedne
izmene). Automatski i bez Excela vrti se samo CI (`.github/workflows/static.yml`):
JSON, `vba_check`, `who_writes`, manifest i self-testovi alata.

RELEASE kapija vrti **dva odvojena Excel run-a**: `--gate pr` (BEHAVIOR, bez
mreže) i `--category external` (EXTERNAL). Bez tog razdvajanja nedostupan SEF
sandbox obara BEHAVIOR i GREEN, pa `--waive external` ne bi rešavao ništa.

---

## 2. Verdikt nije jedan

`REZULTAT: ZELENO` je bio jedan broj za sedam različitih pitanja, pa je
`COMPILE NEJASNO` u njemu nestajao. Sada svaka istina ima svoj red:

```
STATIC    PASS | FAIL                 vba_check + who_writes + manifest
COMPILE   PASS | UNKNOWN | FAIL       VBE Debug > Compile
SCHEMA    PASS | FAIL | NOT_RUN       EnsureRuntimeSchema pre suite-ova
BEHAVIOR  PASS | FAIL | NOT_RUN       suite iz izabrane kapije
COUNTS    PASS | FAIL | NOT_RUN       broj provera vs min_asserts iz manifesta
CLEANUP   PASS | FAIL | NOT_RUN       tragovi test podataka posle run-a
EXTERNAL  PASS | FAIL | NOT_RUN       Google / MasterSync / SEF
```

Dva pravila koja nisu očigledna:

- **`BEHAVIOR PASS` uz `COMPILE UNKNOWN` je dozvoljen ishod za DEV i PR** — da bi
  se suite uopšte pokrenula, VBA je morao da kompajlira nju i sve što ona
  referencira. Ali mora tako i da **piše**. Release kapija `UNKNOWN` ne prihvata:
  isporučuje se ceo projekat, uključujući module koje nijedna suite ne dodiruje.
- **Sama `BLIND` suite ne daje `BEHAVIOR PASS`.** „Prošlo bez greške" nije isto
  što i „sve provere prošle". Isto važi za ceo run: ako nijedan skup ne da
  `PASS`, izlazni kod je 2. `NOT_RUN` nigde nije prolaz.
- **Suitin sopstveni izveštaj je jači dokaz od `Run()`.** Ako je `Err.Raise` u
  suite-u uklonjen ili zaobiđen, VBA završi uredno i runner bi rekao `OK` — iznad
  reda `SUITE|...|188|1|...` koji je suite sama upisala. Prijavljen `FAIL > 0`
  zato prepravlja status u `FAIL`, a `NOT_RUN` u `NOT_RUN`.

---

## 3. Identitet run-a — šta je tačno testirano

`zeleno` bez identiteta je apstraktna tvrdnja. Svaki run zapisuje u
`tests/last_run.json` (i u ispis):

```
IZVOR   src-vba <kanonski sha256>
        git <sha> <grana> <describe>       + oznaka ako je radno stablo prljavo
        fixture <ime> <sha256>
        <platforma>  python <v>  locale <v>
KAPIJA  dev|pr|release   [shuffle seed=N] [repeat=N]
```

Tek uz to „zeleno" znači: **ovaj tačan VBA izvor, nad ovim tačnim fixture-om,
prošao je ovaj set suite-ova.**

Hash mora da opisuje **kod koji je stvarno uvezen**, ne samo repo:

- `--no-import` (testiraj zatečen kod u svesci) **nikad** ne piše marker.
  Bez toga je bila moguća sekvenca: repo = nov kod, sveska = star kod, run zelen
  → marker tvrdi da je nov izvor testiran. `last_run.json` nosi
  `source_imported: true/false`.
- `modVbaTools` **se uvozi**. Self-skip postoji u `ImportAllVBA` jer VBA
  procedura ne može da ukloni modul iz kog se izvršava; Python nema taj problem,
  a `source_hash` taj modul uračunava.
- Nedostajući `.doccls` nad **našim** fixture-om je tvrd pad: kod tih listova nije
  ni uvezen, a hash ga uračunava. Nad automatski napravljenom praznom sveskom
  nije pad, ali se takav run ne zove pun compile izvora (`SOURCE` red u ispisu).

### Kanonski hash — zašto normalizacija prelomaka

`src-vba/` se na Windows-u checkout-uje kao CRLF, a na Linuxu kao LF. Sirov hash
bajtova bi nad **istim commitom** dao dva različita broja, pa bi „zeleno na build
mašini" svuda drugde izgledalo kao nepoznat izvor. Zato se `.bas/.cls/.frm/.doccls`
normalizuju na LF pre hesiranja; `.frx` je binaran i ide sirov.

`modBuildInfo.bas` je **izuzet**: `stamp-build` ga prepisuje pri svakom build-u i
ta izmena se ne commit-uje. Da je u hešu, stamp bi obarao zelen marker baš u
trenutku release-a — a modul ne nosi nikakvo ponašanje, samo tri konstante.

### Ustajao fixture — druga polovina istog pitanja

Last-green marker odgovara na „da li je OVAJ izvor testiran". Postoji i obrnuto
pitanje: **da li je ovo onaj fixture nad kojim su testovi pisani.** Fixture je
gitignored, pa ga `git checkout` NE menja — posle prelaska na drugu granu na disku
ostaje sveska prethodne, testovi padaju *na podacima*, a pad izgleda kao regresija
koda (jednom je pojeo pola sata trijaže).

`make_fixture` zato pored sveske piše `otkup_test.sig` sa hash-om posejanih
podataka, a `run_vba` ga poredi **pre podizanja Excela**. Neslaganje ili
nedostajući potpis zaustavljaju run uz komandu za regeneraciju; jedini izlaz je
svestan `--ignore-fixture-sig`.

**Crveno posle prelaska grane — prvo regeneriši fixture, pa tek onda traži krivca
u kodu.**

### Last-green marker

Posle zelenog punog run-a upisuje se `.git/agrix-vba-last-green`:

```json
{ "source_hash": "...", "gate": "pr", "suites": [...], "timestamp_utc": "...", ... }
```

U `.git` namerno: nikad ne može da se commit-uje i ne preživljava svež klon (na
novoj mašini kapija mora da se dokaže iznova).

```bash
python3 tools/vba_gate.py --status              # izvor vs poslednji zeleni
python3 tools/vba_gate.py --require-green --gate pr
python3 tools/vba_gate.py --hash
```

**Marker pamti i koja je kapija bila zelena.** Zelen `dev` run zadovoljava Stop
hook, ali **ne** zadovoljava release — `dev` ne vrti ni `RunNovacSmokeSuite` ni
external ugovore.

### Zašto hash, a ne git istorija

Raniji uslov (`git diff HEAD~1 -- src-vba/`, u Stop hook-u koji je u međuvremenu
uklonjen) bio je netačan u oba smera:

- commit koji dira samo `docs/` „resetuje" potrebu za testom, jer poslednji
  commit više ne dira `src-vba` — a izvor je u međuvremenu promenjen;
- rebase / amend / cherry-pick pomeraju `HEAD~1` pod nogama, pa se ista izmena
  jednom traži a drugi put ne.

Sada se poredi hash. Isti hash = nema šta da se testira. Različit = testira se,
bez obzira na to kako git istorija u tom trenutku izgleda. To pitanje se sada
postavlja **namerno** (`vba_gate.py --status`) i u release kapiji, a ne na svakom
zaustavljanju sesije.

---

## 4. Manifest — jedini izvor istine o suite-ovima

`tests/suite_manifest.json` nosi za svaku suite: `id`, `module`, `category`,
`gates`, `raises`, `dialogs`, `min_asserts`, `reports_counts`, `timeout_s`,
`uses_workbook`, `mutates`, `external`, `result_file`, `note`.

`python3 tools/vba_gate.py --manifest-check` proverava **oba smera**:

1. **manifest → kod:** upisana suite mora da postoji kao `Public Sub` u navedenom
   modulu. Bez ovoga preimenovanje suite-a tiho izbaci ~200 provera iz kapije.
2. **kod → manifest:** ulazna tačka (`Run*Suite`, `Run*Tests`, `Test*_All`) koju
   **niko ne poziva** i koja nije upisana je nalaz. To je suite koju niko nikad
   neće pokrenuti — napisana pa zaboravljena.

Prvo pokretanje te provere našlo je **10 takvih** (SEF sub-suite-ovi +
`RunHttpUtilsSmokeSuite`). Nisu izbrisane niti tiho dodate u kapiju — upisane su
u `unlisted` sa razlogom po komadu. Prazan razlog je greška u proveri: nepokretana
suite bez zapisanog razloga je tiho preskakanje.

### COUNTS — protiv tihog pada pokrivenosti

Suite prijavljuje broj provera kroz `modTestRunner.TR_Report`, koji piše jedan red
u `suite_results.txt` pored sveske. Runner ga poredi sa `min_asserts`:

```
ASSERTS RunBankaImportTestSuite: 189 provera (min 189)
ASSERTS RunPaleteTestSuite: 120 provera, a manifest trazi >= 97
```

Ako Banka danas ima 189 provera, sutrašnjih 120 je **crveno** dok neko ne spusti
`min_asserts` svesno, u commit-u koji se vidi.

---

## 5. Lifecycle i „nije pokrenuto ≠ prošlo"

`modTestRunner` daje jedan lifecycle:

```
TR_BeginSuite -> TR_BeginTest -> (assert-i) -> TR_EndTest -> TR_EndSuite
                                            \-> TR_NotRun  (eksplicitan NOT_RUN)
```

`TR_EndSuite` upisuje izveštaj i **podiže grešku** ako je išta palo — to je ono
što suite čini kapijom.

`TR_NotRun` postoji zbog tihog `Exit Sub`. Takvih putanja je u ovom projektu bilo
pet; runner ih je sve video kao `OK`:

| Suite | Uslov koji je tiho izlazio | Status |
|---|---|---|
| `RunPaleteTestSuite` | paletiranje isključeno | zatvoreno ranije |
| `RunPaleteTestSuite` | zatečen `TST-` ostatak | zatvoreno ranije |
| `RunPaleteTestSuite` | operater odustao na potvrdi | zatvoreno ranije |
| `RunAgrohemijaSmokeSuite` | dev-guard odbijen | zatvoreno ranije |
| `RunBankaImportTestSuite` | `tblBankaImport`/`tblOtkup` ne postoje | **zatvoreno sada** |
| `RunBankaImportTestSuite` | operater odustao na potvrdi | **zatvoreno sada** |

`NOT_RUN` je **lepljiv** u izveštaju: suite ume da upiše dva reda u istom prolazu
(`TR_NotRun`, pa EH grana pozove `ReportResults` koji upiše 0/0 kao da je sve
normalno). Da poslednji red pobeđuje, „nije pokrenut" bi se izgubilo — tačno ono
stanje zbog kog provera postoji.

### `clsTestContext` — VBA nema `finally`, ali ima `Class_Terminate`

Obrazac `wasQuiet` / `quietSet` / `RestoreJournalQuiet` u dva mesta (normalan
izlaz i EH) živeo je u svakom test modulu zasebno. Svaka kopija je jedna prilika
da se stanje ne vrati — a nevraćeno stanje ne obara suite koja ga je ostavila,
nego **sledeću**.

```vba
Dim ctx As clsTestContext
Set ctx = New clsTestContext     ' snapshot odmah (Class_Initialize)
ctx.Quiet                        ' gasi journal trag i UI cekanja
' ... nema rucnog cleanup-a: Class_Terminate vraca stanje na SVAKOM izlazu
```

`ctx.Drift()` posle restore-a vraća prazan string ako je stanje stvarno vraćeno.
Snima se: `Calculation`, `EnableEvents`, `ScreenUpdating`, `DisplayAlerts`,
`StatusBar`, `Cursor`, `modTestMode` i `modJournaling` test-flagovi.

---

## 6. Assertion API

Jedan modul, `modTestAssert`, umesto pet lokalnih kopija:

```
AssertTrue / AssertFalse
AssertEqual / AssertNotEqual        (poredi kao tekst -- v. komentar u modulu)
AssertNear (tolerancija, default 0.001)
AssertEmpty / AssertNotEmpty / AssertContains
AssertRaised / AssertRaisedNumber / AssertNoError
AssertRowCount / AssertTableRowExists / AssertTableRowMissing
TableRowCount / AssertUnchanged
```

Nijedan ne prekida izvršavanje — jedan pao assert ne sme da sakrije preostalih
dvadeset u istom testu. Suite pada na kraju, kroz `TR_EndSuite`.

**„Ovo mora da bude odbijeno"** je najčešći oblik pravila u ovom projektu, pa je
`AssertRaised` najvažniji. VBA nema delegate, pa se obrazac ne može sakriti do
kraja — ali je svuda isti:

```vba
On Error Resume Next
Err.Clear
SaveNesto "nevalidno"                  ' ovo MORA da pukne
AssertRaised "prazan kooperant se odbija"
On Error GoTo 0
```

Između poziva koji treba da pukne i `AssertRaised` **ne sme** da stoji nijedna
druga naredba: ubačena naredba koja uspe ostavlja `Err.Number = 0` i provera bi
lagala u smeru „nije podignuto".

---

## 7. Release Gate

Redosled je promenjen i to je suština:

```
1  main + pull + cisto radno stablo
2  bump APP_VERSION U RADNO STABLO (bez commita)   <- gate vrti BAS to
3  release gate: static / fixture / behavior / green / compile / external
4  commit + push                                    <- tek posle zelene kapije
5  anotiran tag sa verdiktom kapije + push
6  stamp build otisak
7  preostali Excel koraci (isporuka, ne provera)
```

**Zašto bump ide pre gate-a:** gate mora da vrti tačno ono što će biti tagovano.
Da se `APP_VERSION` menja posle zelenog run-a, hash `src-vba` bi se promenio i tag
bi pokrio izvor koji niko nije testirao — ista rupa, samo manja. Ako kapija padne,
bump se vraća i radno stablo ostaje čisto.

Verdikt završava **u poruci anotiranog taga**. `git show vba-vX.Y.Z` kasnije tačno
kaže šta je bilo zeleno, šta izuzeto i nad kojim hashom.

### Waiver

```
powershell -File tools\release.ps1 2.41.0 -Waive external -WaiveReason "SEF sandbox nedostupan 14.08."
bash tools/release.sh 2.41.0 --waive external --reason "..."
```

Waiver bez razloga se **odbija**. Izuzeta kapija zadržava svoj originalni status u
zapisu (`WAIVED (razlog) -- bilo je: FAIL ...`), pa se u tagu vidi i šta je tačno
zaobiđeno. „Testovi nisu pokrenuti" više ne može da bude fusnota u release notes
koju niko ne traži.

`NOT_RUN` blokira isto kao `FAIL`. „Nije pokrenuto" nije „prošlo".

---

## 8. Nezavisnost od redosleda i stanja

```
python tools\run_vba.py --gate pr --shuffle --seed 12345
python tools\run_vba.py --gate pr --repeat 3
```

- `--shuffle --seed N` — nasumičan redosled, ponovljiv istim seed-om. Hvata test
  koji prolazi samo zato što ga je prethodni pripremio. Seed ide u ispis i u
  `last_run.json`, pa se pad ponavlja tačno.
- `--repeat N` — isti set N puta **u istom Excel procesu**. Hvata klasu kvarova
  koju svež proces nikad ne vidi: stale cache, `Static` Boolean, license latch,
  neotpušteni eventi, globalni `Dictionary`.

**Per-suite timeout** (`timeout_s` u manifestu): ako BFP počne da traje 40 s umesto
5 s, runner kaže `SUITE TIMEOUT RunBusinessFlowProSuite` i ubije run — test je i
detektor regresije u performansama, ne samo correctness gate.

---

## 9. Mašinski čitljiv izveštaj

**Pored sveske** (u temp folderu run-a) suite pišu tri različite stvari, i ne
preklapaju se:

| Fajl | Ko piše | Šta nosi | Ko čita |
|---|---|---|---|
| `last_run.txt` | `modTest` | ime palog testa | `run_vba` (`result_file`), `dokaz.py` |
| `last_run_banka.txt` | `modTestBanka` | ime palog testa | isto |
| `last_run_<suite>.txt` | `modTestRunner.TR_EndSuite` | ime palog testa | isto |
| `suite_results.txt` | `modTestRunner.TR_Report` | **broj** provera | `run_vba` (COUNTS) |

Detalj pada ne preživi COM granicu — `xl.Run` vrati golo „Exception occurred" —
pa svaka suite koja hoće da se vidi **koja** provera je pala mora da ga napiše u
fajl. Broj provera je zaseban zapis jer ga čita druga kapija.

**U repou** (`tests/`, svi gitignored):

| Fajl | Šta je | Ide u git |
|---|---|---|
| `tests/last_run.txt` | ljudski ispis + VERDIKT blok | ne |
| `tests/last_run.json` | pun zapis: provenance, suite-ovi, brojevi, verdikti | ne |
| `tests/last_run.xml` | JUnit XML — GitHub/IDE prikazuju koji test je pao | ne |
| `tests/last_release.json` | zapis release kapije (i waiver-i) | ne |
| `.git/agrix-vba-last-green` | marker poslednjeg zelenog run-a | ne (u `.git`) |

---

## 10. Definition of Done za test infrastrukturu

Stavka je gotova kad važi **sve**:

1. Nijedna suite u DEV/PR kapiji nije blind (`raises: true`).
2. Nema tihog `Exit Sub` — svaki rani izlaz ide kroz `TR_NotRun`.
3. Suite prijavljuje broj provera (`reports_counts: true`) i ima `min_asserts`.
4. Suite može samostalno da se pokrene (`--suite X`).
5. Suite prolazi i u promenjenom redosledu (`--shuffle`).
6. Cleanup je **dokazan**, ne pretpostavljen (`ctx.Drift()` prazan, CLEANUP PASS).
7. Compile status je eksplicitan — `UNKNOWN` se ispisuje, ne prećutkuje.
8. Release nije moguć sa stale behavior izveštajem (GREEN kapija).
9. Broj očekivanih provera ne može tiho da padne (COUNTS kapija).
10. Izmena ponašanja nosi test u `modTest`/odgovarajućem suite-u, ne checklistu.

---

## 11. Šta NIJE urađeno — otvoren dug

Zapisano da ne bi izgledalo kao pokriveno:

| Stavka | Stanje |
|---|---|
| **Migracija suite-ova na `modTestAssert`** | urađen samo kanarinac (`modLicenseTests`) — **1 od 11**. Ostale su na `TR_Report`: broj provera je vidljiv runneru, ali harness, brojači i `Err.Raise` su i dalje njihovi. Zato ovo nije „unified framework" nego temelj. |
| **`dialogs: false` svuda** | 9 od 17 suite-ova i dalje otvara `MsgBox` i oslanja se na watchdog. Cilj je `IUserPrompt`/`TestMode` seam; watchdog je sigurnosna mreža, ne model izvršavanja. Neočekivan dijalog i dalje NIJE FAIL. |
| **CLEANUP** | **blokirajuć u `pr` i `release`**, prijava u `dev`. `--no-enforce-cleanup` je izlaz dok detektor ne bude dokazan u oba smera nad pravim Excelom — a to još nije. |
| **Test Data Builder (`modTestData`)** | nije urađen. Svaki modul i dalje sam sklapa `TST-*` podatke. |
| **Dependency seams (`Now`/ID/FS/HTTP/MsgBox)** | nije urađeno. |
| **Contract testovi legacy vs `frmOtkupUI`** | nije urađeno — kategorija `contract` postoji u manifestu, prazna je. |
| **Windows self-hosted CI runner** | nije urađen. Kapije se i dalje pokreću ručno na build mašini. |
| **Mutation testing kao alat** | nije urađen; sabotaže su i dalje ručne. |
| **Coverage matrica po RF/AUD invarijantama** | nije urađena. |
| **`RunSEFTestSuite` kategorija** | zaveden kao `external`, ali sopstveni header kaže „offline hard gate" i zove samo `Test_*` seam provere. Možda spada u `pr`. **Neprovereno** — traži jedan Windows run. |
| **Release dokazuje izvor, ne artefakt** | kapija pokriva `src-vba` + fixture. Finalni `.xlsm` (ImportAllVBA → Compile → AssertBlankBuild → hash) i dalje nastaje POSLE taga. Sledeći nivo: napravi artefakt, hešuj ga, pa tek onda tag. |
| **`dialogs` i dalje `true` za 9 suite-ova** | watchdog klika dijaloge. Neočekivan dijalog i dalje NIJE `FAIL`. |
| **10 `unlisted` suite-ova** | zapisani sa razlogom, nijedan nije priključen kapiji. |

### Šta je pregled (NO-MERGE nad 7ec4052) našao i gde je zatvoreno

| Nalaz | Gde je zatvoreno |
|---|---|
| P0 `--no-import` označava netestiran izvor kao GREEN | `run_vba.py` — `--no-import` isključen iz uslova za marker; `source_imported` u zapisu |
| P0 prijavljen `FAIL=1` prolazi kao zeleno | `check_counts` + `apply_reported_failures` |
| P1 `--waive external` ne odblokira release | BEHAVIOR = `--gate pr`, EXTERNAL = zaseban run |
| P1 hash tvrdi više nego što je uvezeno | `modVbaTools` se uvozi; nedostajući `.doccls` je pad |
| P1 `AssertTableRowMissing` fail-open | `CountMatching` vraća `okFlag`; infrastrukturna greška ≠ nula pogodaka |
| P1 Novac nije izolovan, a ušao je u kapiju | jedna transakcija nad 4 tabele + `clsTestContext` |
| P1 `NOT_RUN` / prazan izbor daju uspeh | `rc_from` traži pozitivan dokaz; prazan izbor je greška |
| P2 `Drift()` se nigde ne proverava | `TestLicense_All` ga asertuje; neuspeo restore izlazi kroz `Drift` |
| P2 fixture nije „100% sintetički" | tačna formulacija u `make_fixture.py`; `AutomationSecurity=ForceDisable` |

### I jedno upozorenje za prvi Windows run

Sve VBA izmene u ovom sloju su **statički** proverene (`vba_check`: ASCII,
deklaracije, rezervisane reči, duplikati, nepostojeći simbol, arnost) i dokazane
u oba smera nad samim checker-om. **Nijedna nije izvršena nad Excelom** — u
Linux/web sesiji to nije moguće.

Konkretno, prvi `python tools\run_vba.py --gate pr` na Windows mašini može da
pokaže:

- suite čiji `TR_Report` poziv nije dosegnut na zelenoj putanji → `COUNTS FAIL` sa
  imenom te suite. Popravka je pomeriti poziv ili spustiti `reports_counts` na
  `false` za nju;
- `RunNovacSmokeSuite` konvertovan iz blind u gate **i** izolovan transakcijom —
  ako je do sada imao skrivene padove ili je zavisio od podataka koje je sam
  ostavljao, sada će se videti. To je i bila poenta;
- `RunIzvestajTests` `min_asserts: 0` → upisati izmerenu vrednost posle run-a;
- `RunAllTests` `min_asserts: 17` je broj `RunOne` poziva u `modTest`, ne broj
  provera — proveriti da li se poklapa sa izmerenim;
- `CLEANUP` je sada blokirajuć u `pr`/`release` a nikad nije pokrenut nad Excelom;
  ako da lažan nalaz, izlaz je `--no-enforce-cleanup` dok se ne popravi.

Framework-sabotaže su u istom katalogu kao poslovne (`tools/sabotaza.py`) i idu
kroz istu mašineriju: `--proveri-sidra` (statički, ide i kroz `vba_check`) i
`tools/dokaz.py` (pušta suite i traži da padne **baš ta tvrdnja**). Zato
`TR_EndSuite` piše `last_run_<suite>.txt` u formatu koji `dokaz.py` već čita, a
katalog framework-sabotažu imenuje **suite-om** umesto `T_` testom.

Izuzetak je `counts-pad`: on ne obara nijednu *tvrdnju* nego *kapiju* (suite i
dalje prolazi, samo prijavi manje provera). `dokaz.py` po konstrukciji to ne vidi,
pa je zapisan u `POZNATI_NALAZI_DOKAZ` sa razlogom; crveno se dobija sa
`run_vba.py --gate pr`, red `ASSERTS TestLicense_All: ...`.

**Redosled dokaza koji se traži pre merge-a** (`tools/sabotaza.py --lista` nosi
prve tri):

1. `python tools\run_vba.py --gate pr` → `exit 0`
2. `python tools\sabotaza.py license-assert` → `FAIL TestLicense_All`, pa `--vrati`
3. `python tools\sabotaza.py license-cleanup` → `FAIL` na `ctx.Drift()`, pa `--vrati`
4. `python tools\sabotaza.py counts-pad` → `COUNTS FAIL` (suite prolazi, broj pao)
5. ručno ostavi `TST-` red → `CLEANUP FAIL` u `--gate pr`
6. `--no-import` posle izmene izvora → marker se NE upisuje (`vba_gate --status`)
7. izmena izvora posle zelenog → `vba_gate --require-green` `exit 2`
8. bez mreže + `--waive external --reason "..."` → release prolazi ostale kapije
