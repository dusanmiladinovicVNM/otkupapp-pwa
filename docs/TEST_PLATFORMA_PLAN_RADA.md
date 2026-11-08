# Test platforma — detaljan plan rada

> Cilj: dovesti verifikaciju ponašanja na nivo na kome je statička osa već —
> da „zeleno" nastaje iz **dokaza**, a ne iz odsustva greške. Pet stavki,
> poređanih po zavisnostima, svaka sa izmerenim zatečenim stanjem, dokazom
> u oba smera i cenom.
>
> Mereno nad `main` = `e20640d` (08.10.2026). Svaki broj u ovom dokumentu je
> izmeren, ne procenjen; gde nije — tako i piše.

---

## 0. Zašto ovaj plan, i šta je već rešeno

`main` je kroz #402 dobio **identitet onoga što je dokazano**: popis suita
(`vba_gate --popis`), marker zelenog po suite-u (`--require-green`) i potvrdu
compile-a vezanu za otisak (`--mark-compile`). To je zatvorilo pitanje „da li je
**baš ovaj** izvor prošao testove".

Ostalo je pitanje koje marker ne može da postavi sam sebi: **da li je ono što je
upisano kao `OK` zaista nešto izmerilo.**

### Dve grane koje čekaju merge

| Grana | Zatvara | Stanje |
|---|---|---|
| `claude/vba-gate-compile-po-izvoru` | `marker["compile"]` je bio jedan objekat → `--mark-compile` na drugoj grani gazio potvrdu prve (#405/#406). Sada rečnik po izvoru, uz migraciju starog oblika. Dokaz 9/9 dvosmerno | **54 behind**, nije merge-ovana |
| `claude/release-kapija-pred-tagom` | Tag nije mogao da nastane bez dokaza. `bump → KAPIJA → commit → anotiran tag → push`. Dokaz `tools/dokaz_release.sh` 18 tvrdnji + 5/5 sabotaža | **54 behind**, sedi na prvoj |

Na `main`-u je i dalje: `release.sh` kapija **0 pojava**, `vba_gate.potvrde_compile`
**0 pojava**.

### Dva nalaza iz merenja koja menjaju redosled

**N1 — `SKIP` se broji u VBA i baca se na granici fajla.**
`modBusinessFlowProTests`, `modSEFTests`, `modFakturaTests` i
`modAgrohemijaTests` svaki ima `LogSkip` koji radi:

```vb
Private Sub LogSkip(ByVal testName As String, ByVal reason As String)
    m_Total = m_Total + 1
    m_Skipped = m_Skipped + 1
```

Zaglavlje rezultat-fajla je `TESTS=n FAIL=m`. **`SKIP=` nema nigde u `src-vba`
— 0 pojava.** Dakle:

> BFP run u kome **sve** provere preskoče piše `TESTS=310 FAIL=0`, `run_vba`
> javi „310 ukupno, 0 palo", marker upiše `status: OK`, `--require-green` to
> prizna, a `min_asserts: 300` bi prošao.

Ne fali samo prag — fali razlika između **izmereno** i **preskočeno**. Broj koji
danas postoji je već nepouzdan. Zato je ta stavka podignuta na prvo mesto posle
merge-a.

**N2 — `--popis` ima slepu tačku.**
`src-vba/modE2EReleaseGate.bas` (160 linija, v6.10, engleski header) ima javnu
ulaznu tačku `RunE2EReleaseGate_v610` koja:

- **nije** u `SUITES`,
- **nije** u `SUITE_VAN_KAPIJA`,
- `IME_SUITE` regex je **ne matchuje** (traži završetak na `Suite`/`Tests`, a
  ime se završava na `_v610`).

`vba_gate --popis` javlja „nema nalaza". Šesta suite van kapije, nevidljiva
popisu — i orkestrira `RunNovacSmokeSuite` + `RunProductionHealthCheck`, dve
`gate: False` suite.

**N3 — `dokaz.py` ne dosegne većinu modula, pa plan dokaza nije izvodljiv bez
pripreme.** Tabela „Dug sa imenom" u `STANJE_REFAKTORA.md` to već imenuje:

> kapija kataloga sabotaža priznaje samo `modTest`, `modTestBanka` i
> `modBusinessFlowProTests`, pa tvrdnja iz `modIzvestajTests` **ne može da se
> obori**.

Ovaj plan na više mesta traži „jednu novu sabotažu po gejtovanoj suite". Za
devet od dvanaest modula to **danas nije izvodljivo**. Dohvat `dokaz.py`-ja je
zato **prethodna stavka**, a ne detalj — i po main-ovom zapisu ide kao
**zaseban process PR** sa svojim dvosmernim dokazom. Dodata kao stavka **0**.

---

## 0. Dohvat `dokaz.py`-ja (prethodna stavka)

**Zašto pre svega ostalog:** bez nje stavke 2b, 3 i 4 ne mogu da zadovolje
`CLAUDE.md` §5 — tvrdnja koja se ne može oboriti nije dokazana, a plan bi se
sveo na „suite je zelena".

**Zatečeno:** kapija kataloga priznaje tri modula (`modTest`, `modTestBanka`,
`modBusinessFlowProTests`) od dvanaest koji nose tvrdnje.

**Posao:** proširiti kapiju na preostalih devet, uz dva ograničenja koja je main
već platio:

1. **Ključ tvrdnje mora biti ceo statičan literal.** `dokaz.py` za BFP beleži
   *naziv* tvrdnje (runner ne ispisuje ime `Sub`-a), pa je tvrdnja ključ.
   Merenje je dalo **69** tvrdnji koje su prolazile kapiju kao **podniz** a
   `dokaz.py` ih je 25 minuta kasnije javljao kao „NE OBARA SVOJ TEST". Kapija
   je pojačana; **8 ih i dalje stoji imenovano u `POZNATI_NALAZI` sa receptom.**
2. **Svaki nov modul u dohvatu nosi svoju zamku.** `modIzvestajTests` i
   `modTestStornoCentar` broje inline (`m_izvPass`, `mPass = mPass + 1`), bez
   imena `Sub`-a u ispisu — pa im format ključa treba **izmeriti**, ne
   pretpostaviti.

**Dokaz:** po modulu — jedna tvrdnja oborena, `dokaz.py <prefiks>` je javi
**po imenu**, vrati, zeleno. Plus kapija kataloga mora da padne na skraćenom
(ne-celom) literalu, u oba smera.

**Cena:** ~2 dana, **zaseban process PR**. Bez Excela za kapiju; `dokaz.py`
prolaz traži Windows.

---

## 1. Merge dve gurnute grane

**Cilj:** ono što je dokazano prestane da čeka.

### Koraci

1. `git fetch origin main`
2. `powershell -File tools/check_merge.ps1` nad prvom granom
3. rebase prve na `origin/main` **lokalno**, pa **pokazati rezultat** (log, diff
   vs `main`, pun set jeftinih kapija)
4. `git push --force-with-lease` **tek po eksplicitnom odobrenju** (`CLAUDE.md` §6)
5. isto za drugu, rebase-ovanu na prvu
6. PR-ovi u redosledu: `vba-gate-compile-po-izvoru → main`, pa
   `release-kapija-pred-tagom → main`

### Rizik rebase-a

Prva grana dira `tools/vba_gate.py` (+241 linija), druga
`tools/release.sh` / `release.ps1` / `.github/workflows/static.yml`.
Kroz 54 commita na `main`-u **nijedan od tih fajlova nije u diff-u** — proveriti
pred rebase, ne pretpostaviti.

### Dokaz

`vba_gate --self-test` (9 pravila dvosmerno) · `bash tools/dokaz_release.sh`
(18 tvrdnji) · **18 poziva jeftinih kapija** iz `static.yml` (`vba_check`,
`vba_gate --popis`, `vba_parity_check`, `vba_selfupdate_gates`,
`vba_hard_census`, `vba_import_marker_gate`, `who_writes --check`
+ `--check-ownership`, `gen_schema_module --check`, `run_vba --self-test`,
i njihovi `--self-test` koraci). **Bez Excela**, ~15 s ukupno.

---

## 2. Tihi izlazi i šest suita van kapije

### 2a. `modTestBanka` — dva reda

| Mesto | Šta se dešava |
|---|---|
| `src-vba/modTestBanka.bas:66-68` | nema `tblBankaImport`/`tblOtkup` → `MsgBox` + `Exit Sub` |
| `src-vba/modTestBanka.bas:74` | „Ne" na `vbYesNo` → `Exit Sub` |

Oba izlaze **pre** `mPass = 0: mFail = 0`, dakle pre ijedne provere, i pre
`WriteResultFileBanka`.

Danas to **slučajno** radi ispravno: fajla nema → `_read_test_results` vrati 2 →
run pada. Ali se oslanja na to da u `tests/fixtures/` ne stoji zatečen
`last_run_banka.txt` od pre. `run_vba._copy_txt` kopira `*.txt` u temp folder,
pa zatečen fajl **može** da pređe. **To nije provereno** i ne tvrdi se ni u
jednu stranu — ali kapija ne sme da zavisi od toga.

**Rešenje:** eksplicitno stanje umesto zaključivanja po odsustvu fajla —
rezultat-fajl sa `NOTRUN=1` i razlogom, pa `Err.Raise`. Čeka deljeni pisač iz
3b; dok njega nema, dva reda inline.

### 2b. Šest suita van kapije

`SUITE_VAN_KAPIJA` **već nosi posao po suite-u** — ovo je izvršenje, ne
istraživanje.

| Suite | Fajl | Posao (po registru) |
|---|---|---|
| `RunHttpUtilsSmokeSuite` | `modSEFTests.bas` | `Err.Raise` na kraju tela + `SUITES` `gate: True`. Zaglavlje tvrdi `PASS=18` → taj broj ide u `min_asserts` |
| `RunSEFDocumentIdShapeSuite` | `modSEFTests.bas` | isto; zaglavlje tvrdi `PASS=14` |
| `RunSEFOfflineSuite` | `modSEFTests.bas` | telo ima `LogFatal` a ne `Err.Raise`. **Deklarisana sa opcionim argumentom** (`Optional ByVal fakturaID As String = ""`) — proveriti da je `xl.Run("<ime>")` zove bez problema |
| `RunSEFStateTransitionSuite` | `modSEFTests.bas` | **prvo odluka:** registar kaže da deo istih `Test_*` provera već vrti `RunSEFTestSuite` (koja JE u `SUITES` i diže grešku). Spojiti tvrdnje ili gejtovati zasebno |
| `RunSEFClientParserSmokeSuite` | **`modSEFClient.bas`** | **test živi u produkcionom modulu** i pad broji u lokalnu promenljivu. Premestiti u `modSEFTests` + `Err.Raise`, ili obrisati ako je pokriveno drugde |
| `RunE2EReleaseGate_v610` | `modE2EReleaseGate.bas` | **novo, v. N2.** Odluka: u registar sa razlogom, ili obrisati — orkestrira tri suite koje `run_vba` ionako pušta, a dve su blind |

### Slepa tačka popisa (N2) — dve opcije, obe sa cenom

| Opcija | Cena |
|---|---|
| proširiti `IME_SUITE` regex | rizik lažnih nalaza nad celim `src-vba` — ista klasa greške zbog koje je širenje `ARNOST`-a odbijeno sa **406 lažnih nalaza** |
| novo pravilo: javna procedura u `mod*Test*` / `mod*Gate*` modulu koja nije ni u `SUITES` ni u registru je nalaz | uže, vezano za **fajl** a ne za ime |

**Preporuka: drugo.** Prvo pogađa ceo izvor.

### Pre-flight

`RunSEFStateTransitionSuite` i `RunSEFClientParserSmokeSuite` menjaju **premisu
postojećeg testa** (spajanje tvrdnji / brisanje), a `RunE2EReleaseGate_v610` je
kandidat za brisanje. Po `.claude/rules` izuzeci prestaju da važe čim rez traži
promenu premise testa → **skill `pre-flight` ide pred kod** za te tri. Ostale
tri (`Err.Raise` + unos u `SUITES`) su mehaničke i ne traže ga.

### Dokaz

Po `CLAUDE.md` §5: **u rezu samo nove sabotaže**,
`python tools/dokaz.py <prefiks>`, **jednom**, kad su `vba_check` i ciljana
suite zeleni i kad je diff pročitan. Po gejtovanoj suite jedna nova sabotaža:
obori jednu njenu tvrdnju → `run_vba --suite <ime>` pada **po imenu** → vrati →
zeleno. Plus: `vba_gate --popis` mora da padne na zastarelom unosu kad suite
izađe iz registra a nije u `SUITES`.

---

## 3. `SKIP` i counts kapija

### 3a. `SKIP` prestaje da bude nevidljiv

**Najveća vrednost po liniji koda u celom planu.** V. N1.

#### Koraci

1. Zaglavlje postaje `TESTS=n FAIL=m SKIP=s` u sva četiri pisača:
   `modTest.bas:8688`, `modTestBanka.bas:2333`,
   `modBusinessFlowProTests.bas:23420`, `modTestStorno.bas:1442`.
2. `run_vba._read_test_results` parsira `SKIP=`. **Unazad kompatibilno:** fajl
   bez `SKIP=` daje `skip = -1` („nepoznato"), što se **ne** tumači kao 0.
3. Ispis: `TESTS <suite>: n ukupno, f palo, s preskoceno`.
4. Kapija: `skip > 0` **ne obara** run sam po sebi, ali
   - `n - s` (stvarno izmereno) je ono što `min_asserts` meri,
   - `s == n` (sve preskočeno) je **pad**, ne zeleno.
5. `vba_gate`: marker pamti `preskoceno` uz `ukupno`/`palo`; `--require-green`
   ne priznaje suite kojoj je sve preskočeno. **Menja `kapija` deo ugovora**, pa
   obara zapisane GREEN-ove **sam**, bez bumpa `MARKER_VERZIJA` — isti obrazac
   kao u `claude/vba-gate-compile-po-izvoru`.

#### Dokaz — dvosmerno, bez Excela

`run_vba --self-test` dobija slučajeve: fajl sa `SKIP=310 TESTS=310` mora da da
**pad po imenu**; fajl bez `SKIP=` mora da da „nepoznato" a ne 0; `min_asserts`
mora da meri `n-s`. `vba_gate --self-test` dobija slučaj „suite sa svim
preskočenim nije dokaz". Ugasi svako pravilo pojedinačno → crveno po imenu.

**Cena:** ~1 dan. Dokaz bez Excela; jedna Excel sesija za potvrdu da zaglavlje
izlazi ispravno.

### 3b. Rezultat-fajl za sve gejtovane suite + pragovi

**Zatečeno: 4 od 22 suita ima `result_file`** (`RunAllTests`,
`RunBankaImportTestSuite`, `RunStornoTestSuite`, `RunBusinessFlowProSuite`). Za
ostalih 18 verdikt je „`Run()` nije pukao" — **nikakav broj ne postoji**, pa je
`min_asserts` za njih neizvodljiv.

I: pisač je **već četvorostruko dupliran** — `WriteResultFile`,
`WriteResultFileBanka`, `WriteResultFileBFP`, plus `WriteTextFile` /
`WriteTextFileBanka`. Peta kopija nije opcija (`extend > duplicate`).

#### Koraci

1. **Jedan deljeni pisač** — `modTestIzvestaj.bas`:
   `TI_Upisi(suiteId, pass, fail, skip, report)` piše `last_run_<suiteId>.txt` u
   **postojećem formatu**. Ovo je seme koje stavka 4 proširuje, ne nov sloj.
2. Četiri postojeća pisača pozivaju njega; svoje telo brišu. Format se ne menja
   osim `SKIP=` iz 3a — `run_vba` ih čita istim putem.
3. Po **jedan** poziv `TI_Upisi` u završnu proceduru svake od 18 preostalih
   suita, uz postojeće brojače. **Aserti se ne diraju.**
4. `result_file` u `SUITES` za svaku.
5. `min_asserts` po suite-u — prag iz **izmerenog**:

| Suite | Izmereno | Napomena |
|---|---|---|
| `RunAllTests` | **209** `RunOne` poziva | statički proverljivo → self-test poredi prag sa brojem poziva |
| `RunBusinessFlowProSuite` | ~2419 kandidata | prag iz **prvog pravog run-a**, ne iz grep-a |
| `RunSEFTestSuite` | 319 | isto |
| `RunStornoTestSuite` · `RunFakturaSmokeSuite` | 27 · 27 | |
| `RunAgrohemijaSmokeSuite` · `RunBankaImportTestSuite` | 26 · 25 | |
| `TestLicense_All` | 23 | |
| `RunPaleteTestSuite` · `RunNovacSmokeSuite` | 11 · 11 | |
| `Test_StornoCentar_All`, `RunIzvestajTests`, `RunGoldenSuite`, `RunSheetsJsonParserTests` | **nemerljivo grep-om** | broje inline (`mPass = mPass + 1`, `m_izvPass`) → prag iz prvog run-a |

6. Za `RunAllTests` prag ide **sa self-testom** koji ga poredi sa brojem
   `RunOne` poziva. Bez toga zastari tiho — **to se već desilo**: prag je bio 17
   dok je kod imao 163.

#### Dokaz

`run_vba --self-test`: prag ispod izmerenog mora da obori run **po imenu**; prag
za `RunAllTests` mora da pukne kad se broj `RunOne` poziva promeni. Dvosmerno.

**Cena:** ~2 dana. **Traži pun Excel prolaz** — bez njega su pragovi iz grep-a,
dakle pretpostavka, ne merenje.

### `.claude/rules/testovi.md`

Posle 2b i 3b pravilo je zastarelo na tri mesta: „katalog `SUITES` je jedini
izvor istine" (postoji i `SUITE_VAN_KAPIJA`), „nova suite mora biti `gate`", i
nema ni `vba_gate`, ni markera, ni `SKIP`-a. **Zaseban process PR, jedan po
jedan** (`CLAUDE.md` §6) — nikad uz ovaj rez.

---

## 4. Objedinjavanje test koda

Najveći deo zadatka, i jedini koji traži pravu migraciju.

**Zatečeno: `0/4` deljenih modula; 12 modula nosi svoju infrastrukturu**, i
aserti su **imenski različiti** — nema jedinstvene semantike:

| Modul | Aserti | Brojači |
|---|---|---|
| `modTest` | `AssertEq`, `AssertSnapshot` | `m_Total`/`m_Failed` |
| `modBusinessFlowProTests` | `AssertTrue`, `AssertFalse`, `AssertEquals`, `AssertDoubleNear` | + `m_Skipped` |
| `modSEFTests` | `AssertTrue`, `AssertEquals`, `AssertContains`, `AssertTransitionAllowed/Blocked` | + `m_Skipped` |
| `modFakturaTests` | `AssertFakturaTrue` / `DoubleEquals` / `TextEquals` | — |
| `modNovacTests` | `AssertNovacTrue` / `DoubleEquals` / `TextEquals` | — |
| `modAgrohemijaTests` | `AssertAgroTrue` / `DoubleEquals` | — |
| `modGoogleSyncSmokeTests` | `AssertTrue`, `AssertEquals`, `AssertConfigKeyPresent` | — |
| `modLicenseTests` | `AssertEq`, `AssertTrue` | `mPass`/`mFail` |
| `modTestBanka`, `modTestPalete`, `modTestStorno` | **nema asserta** | `mPass`/`mFail`/`mFails` |
| `modTestStornoCentar` | nema | `mPass`/`mFail`/`mFailImena` |

### Zamka koja odlučuje redosled

`AssertTrue` / `AssertEquals` postoje kao **`Private`** u tri modula. Jedan
`Public AssertTrue` u deljenom modulu **ne** daje `DUPLIKAT` (privatni član se
ne sudara globalno) — ali u modulu koji ima svoj privatni, deljeni **nikad neće
biti pozvan**: lokalni ga zaklanja, a VBA je case-insensitive.

> Migracija zato mora da **briše privatni assert u istom rezu**. „Dodaj deljeni
> pa postepeno migriraj" daje dve semantike pod istim imenom, bez ijednog
> nalaza.

### Redosled — najmanji rizik prvi, svaki svoj PR

| # | Modul | Zašto taj red |
|---|---|---|
| 1 | `modLicenseTests` | 23 tvrdnje, 2 asserta, bez tabela i transakcije — čista kanarinka |
| 2 | `modTestStornoCentar` | nema asserta, samo brojače → migracija je **uvođenje** asserta, ne zamena |
| 3 | `modTestPalete` | 11 tvrdnji, isti obrazac |
| 4 | `modTestBanka` | 25 tvrdnji; **zatvara i 2a** istim rezom |
| 5 | `modTestStorno` | 27 |
| 6 | `modFakturaTests`, `modNovacTests`, `modAgrohemijaTests` | prefiksirani aserti → **nema zaklanjanja**, mehanički |
| 7 | `modGoogleSyncSmokeTests`, `modSEFTests` | imaju `AssertTrue`/`AssertEquals` → **zaklanjanje**, briše se u istom rezu |
| 8 | `modBusinessFlowProTests` | ~2419 kandidata, 23k linija — poslednji, verovatno u više rezova |
| 9 | `modTest` | `AssertEq` + `AssertSnapshot`, 209 `RunOne` — jezgro, poslednje |

### Šta se gradi

| Modul | Uloga |
|---|---|
| `modTestAssert.bas` | jedinstven API, `AssertRaises`, **fail-closed** tabelarne tvrdnje (infrastrukturna greška se razlikuje od poslovne nule) |
| `modTestRunner.bas` | **nad `modTestIzvestaj` iz 3b** — proširenje, ne nov sloj |
| `clsTestContext.cls` | snapshot/restore stanja Excela (`Calculation`, `EnableEvents`, `ScreenUpdating`, `DisplayAlerts`, `TestMode`, journal), restore u `Class_Terminate` — jedini „finally" koji VBA ima |
| `clsTestResult.cls` | jedan red rezultata, jedan format |

Nacrt sa mrtve grane `44e9e4b` je polazna tačka, ali: **nijedna linija nikad
nije izvršena u Excelu**, i nastao je **pre** `DUPLIKAT_LOKALNI`, `ZAKLONJENO` i
`SEMA_REGISTAR` — mora ponovo kroz `vba_check` pred bilo kakvu tvrdnju.

### Dokaz po rezu

Jedna **nova** sabotaža nad migriranim assertom: obori tvrdnju →
`run_vba --suite <ime>` padne **po imenu** → vrati → zeleno. Pun
`python tools/dokaz.py` ide **pred release**, uz FULL prolaz (`CLAUDE.md` §5).
`clsTestContext` nosi svoju tvrdnju: `AssertEmpty ctx.Drift()` — stanje Excela
vraćeno na ulazno.

**Dva ograničenja koja migracija mora da poštuje, oba već naplaćena na main-u:**

1. **Zavisi od stavke 0.** Za devet od dvanaest modula `dokaz.py` danas ne može
   da obori tvrdnju. Bez stavke 0 ovaj dokaz nije izvodljiv.
2. **Labela tvrdnje je KLJUČ sabotaže i mora ostati ceo statičan literal.**
   Migracija menja upravo te labele. Main je to već platio: **69** tvrdnji je
   prolazilo kapiju kao podniz, a `dokaz.py` ih je javljao kao „NE OBARA SVOJ
   TEST"; 57 je širen mehanički, 4 ispravljena u testu, **8 i dalje stoji u
   `POZNATI_NALAZI` sa receptom**. Svaki rez migracije mora da pusti kapiju
   kataloga **pre** `dokaz.py`-ja, inače se greška vidi 25 minuta kasnije.

**Zaklanjanje ima i svoju nepostojeću kapiju.** Main vodi dug „nema kapije
*modul ne sme da koristi tuđ `Private` simbol*": VBA kompajlira na zahtev, pa
`Sub or Function not defined` pukne tek kad neki test prvi put pozove tu
proceduru — **posle 600 s i ubijenog Excela**, uz poruku bez fajla i linije.
Jednokratni merač je napisan i dvosmerno dokazan (1 nalaz nad pokvarenim
izvorom, 0 posle). Migracija koja briše privatni assert treba da ga pusti pred
svaki rez; trajna kapija u `vba_check` je zaseban process PR.

**Cena:** 9 rezova, svaki sa svojom Excel sesijom. Realno **2–3 nedelje**
kalendarski; `modBusinessFlowProTests` sam je verovatno trećina.

---

## 5. Dijalozi, artefakt, razdeljen verdikt

### 5a. `dialogs: false` — dva različita posla, ne jedan

**Zatečeno, test moduli:** 12 × `dialogs: True`. U gejtovanim test modulima
**42** `MsgBox`-a, ali samo **4 su `vbYesNo`** (`modTestBanka`, `modTestPalete`,
`modTestStorno`, `modAgrohemijaTests`) — ostalo su izveštajni na kraju.
`modBusinessFlowProTests` sam ima 11.

**Zatečeno, pisac-putanja — i to je teži deo.** Main to već vodi kao dug:

> Protokol potvrde deficita pita operatera na **tri** mesta (`modOtkupUnos` od
> 03.10.2026, `modNovacUnos` i `modDokUnos` od 07.10.2026). Test koji uđe u tu
> granu ne pada nego **visi do timeout-a** i ostavlja Excel u `[break]` — ista
> cena kao compile greška (**585 s + ubijen Excel**). Danas to drže samo
> komentari uz tri grane.

Dakle `MsgBox` nije samo u test modulima. Dva posla:

| | Posao | Cena |
|---|---|---|
| **5a-1** | `If Not IsTestMode()` oko `MsgBox`-a u **test modulima**. Seam je zatečen (`modOtkupUI` oko tri `SetFocus`-a, `modGoldenTests`), a `modBusinessFlowProTests` već koristi `IsTestMode` na 10 mesta | ~1 dan, rizik nizak |
| **5a-2** | `MsgBox` u **pisac-putanji** (`modOtkupUnos`, `modNovacUnos`, `modDokUnos`). Po main-ovom zapisu „kapija bi morala da zna koji su pozivi iz suite-a dostupni, pa traži **svoj rez i svoj dvosmerni dokaz**" | zaseban rez, ne deo ovog |

`dialogs: False` u `SUITES` se sme upisati **tek posle 5a-1 i 5a-2** — dok
pisac-putanja visi, `dialogs: False` bi tvrdio nešto netačno. Tek tada watchdog
koji **uhvati** dijalog postaje **pad**, ne zabeležen događaj, i to je prava
vrednost: dijalog u headless run-u tada znači da je neki put izbegao seam.

**Dokaz (5a-1):** ukloni jedan `If Not IsTestMode()` → `run_vba` javi `DIALOG` i
padne. Dvosmerno.

### 5b. Artefakt vezan za tag

**Zatečeno je bolje nego što je ranije u lancu tvrđeno:**
`modRelease.PublishReleaseToDrive` **već** računa SHA-256 po fajlu i piše ih u
`version.json` / `manifest.json`, a `manifest_sha256` u `current.json`. Otisci
artefakta postoje.

**Šta fali:** veza između taga i tih otisaka. Tag nastaje u koraku 2d procedure,
`.xlsm` se gradi u 4–8 — artefakt **fizički ne postoji** kad tag nastaje, pa ga
tag ne može nositi.

**Rešenje koje ne laže o redosledu** — potvrda **posle**, vezana za izvor, isti
obrazac kao `--mark-compile`:

1. `vba_gate --mark-artifact <putanja>` → upis u marker **po izvoru**:
   `{izvor: {sha256, ime, kada}}`. Reuse mehanike `potvrde_compile`, ne nova.
2. `vba_gate --require-artifact` → `rc=2` ako artefakt nije potvrđen nad **ovim**
   izvorom.
3. `release.sh` dobija korak posle Excel dela: `--mark-artifact`, pa
   `git notes add` na tag sa otiskom. `git show <tag>` time pokazuje i izvor
   (pred tagom) i artefakt (posle).
4. `manifest_sha256` iz `current.json` se poredi sa upisanim — razlika znači da
   objavljeno nije ono što je potvrđeno.

**Pošteno:** ovo **ne** čini artefakt proverenim *pred* tagom — to je nemoguće
jer ne postoji. Čini ga **proverljivim posle**, i vezanim za izvor koji je tag
dokazao. Alternativa (tag posle artefakta) značila bi da `.xlsm` nastaje iz
necommit-ovanog stanja, što je gore.

**Cena:** ~1 dan, reuse grane 1. Traži Excel sesiju za `PublishReleaseToDrive`.

### 5c. Razdeljen verdikt + `--waive` u `run_vba`

**Zatečeno:** jedan `REZULTAT: ZELENO/PALO`; `waive` **0 pojava** u `run_vba`.
Dijagnoza **jeste** razdeljena (`IMPORT` / `COMPILE` / `SCHEMA` / `SUITE` /
`TESTS` / `DIALOG` / `FATAL`), verdikt nije — pa `--waive` nema na šta da se
primeni.

**Rešenje:** `rc` po osama (`STATIC` / `COMPILE` / `SCHEMA` / `BEHAVIOR` /
`COUNTS` / `CLEANUP`), ispis sa po jednom linijom verdikta, `--waive <osa>
--reason` traži razlog i upisuje ga u izveštaj **i u marker**, da
`--require-green` vidi da je osa waived a ne dokazana.

**Realno: najmanje vredna stavka u planu.** `--waive` u release kapiji pokriva
stvaran scenario; waiver u runneru je ugodnost. **Preporuka: poslednje, ili
nikad** — i odluka tek kad 3a/3b pokažu kako se ose ponašaju u praksi.

---

## 6. Zbirno i redosled

| # | Posao | Cena | Excel | Pre-flight | Vrednost |
|---|---|---|---|---|---|
| 1 | merge dve grane | pola dana | ne | ne | **visoka** — gotovo |
| 0 | dohvat `dokaz.py`-ja na 9 modula | ~2 dana, **process PR** | prolaz da | ne | **blokira 2b, 3, 4** |
| 2a | banka tihi izlazi | 2 reda | 1 sesija | ne | srednja |
| 2b | šest suita + slepa tačka popisa | ~1 dan + 3 odluke | 1 sesija | **da** (3 suite) | visoka |
| 3a | `SKIP` vidljiv | ~1 dan | 1 sesija | ne | **najviša po liniji** |
| 3b | rezultat-fajl za 18 + pragovi | ~2 dana | **pun prolaz** | ne | visoka |
| 4 | objedinjavanje, 9 rezova | 2–3 nedelje | 9 sesija | ne | visoka, dugoročna |
| 5a-1 | `dialogs: false` u test modulima | ~1 dan | 1 sesija | ne | srednja |
| 5a-2 | `MsgBox` u pisac-putanji | **zaseban rez** | da | ne | srednja, main-ov dug |
| 5b | artefakt vezan za tag | ~1 dan | 1 sesija | ne | srednja |
| 5c | waive u runneru | ~1 dan | ne | ne | **niska — preskočiti** |

### Tri kritične zavisnosti

1. **Stavka 0 blokira dokaze u 2b, 3 i 4.** Za devet od dvanaest modula
   `dokaz.py` ne može da obori tvrdnju, pa bi „dokazano" značilo samo „suite je
   zelena". Ide pre njih, i kao **zaseban process PR**.
2. **3b gradi deljeni pisač koji 4 proširuje.** Ako 4 krene pre 3b, pisač se
   pravi dva puta — a duplirani pisač je tačno ono što 3b uklanja.
3. **Ništa od ovoga nije izvršeno u Excelu.** 3b bez punog prolaza daje pragove
   iz grep-a, dakle pretpostavku umesto merenja.

### Predloženi redosled

```
1  ->  3a  ->  0  ->  2a + 2b  ->  3b  ->  4  ->  5a-1  ->  5b
                                                    (5a-2: svoj rez)
                                                    (5c:   preskočiti)
```

`3a` ide odmah posle merge-a: najjeftiniji pravi dobitak u planu, **ne zavisi
ni od čega**, i njegov dokaz leži u `run_vba --self-test` / `vba_gate
--self-test` — dakle **ne traži dohvat `dokaz.py`-ja**. Stavka 0 dolazi tek
posle njega, jer je prva koja blokira ostalo.

---

## 7. Šta ovaj plan ne tvrdi

- Da su pragovi iz tabele u 3b tačni. Šest je izmereno grep-om nad izvorom,
  četiri su **nemerljiva** statički, a BFP-ov i SEF-ov su kandidati a ne brojevi.
  Pragovi se postavljaju iz **prvog punog prolaza**.
- Da `modTestBanka` danas daje lažno zeleno. Oslanja se na odsustvo zatečenog
  `last_run_banka.txt`; da li ga `_copy_txt` može preneti — **nije provereno**.
- Da nacrt sa `44e9e4b` radi. Nijedna njegova linija nikad nije izvršena u
  Excelu.
- Da je spisak suita van kapije konačan. `--popis` ima slepu tačku (N2), pa dok
  se ona ne zatvori, „šest" je donja granica a ne broj.
- Da je plan dokaza izvodljiv bez stavke 0. Za devet od dvanaest modula
  `dokaz.py` danas ne dohvata tvrdnje (N3), pa bez te stavke „dokazano" u 2b, 3
  i 4 znači samo „suite je zelena".
- Da su cene u satima merene. One su procena po obimu zatečenog koda; jedino
  što je izmereno jesu brojevi tvrdnji, modula, `MsgBox`-a i suita.
