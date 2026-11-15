# Plan optimizacije verifikacije — provere, testovi, dokazi

> **Status:** predlog za odluku vlasnika. **Datum:** 15.11.2026.
> **Podloga:** osam nezavisnih analiza nad `origin/main @1fa4937` (08.11.2026),
> sačuvanih u `docs/engineering/analize/2026-11-verifikacija/`.
>
> Nijedna stavka odavde nije sprovedena, osim koraka 0 i 1 (izbor testova i vreme
> po testu) na grani `claude/dazzling-sagan-3o531f`, commit `9d6e780`, čija VBA
> strana još **nije izvršena u Excelu**.

**Oznake uz brojeve.** `[M]` izmereno u ovoj analizi (Linux sesija, bez Excela).
`[D]` broj iz dokumentacije ili dnevnika repoa, sa datumom. `[I]` izvedeno
modelom iz `[M]` i `[D]`. `[?]` nepoznato dok se ne izmeri. Nijedno trajanje
Excel prolaza nije `[M]`: Excel u ovoj sesiji ne postoji.

**Oznake uz stavke.** Trud: `S` do jednog radnog dana, `M` dva do četiri dana,
`L` jedna do dve nedelje (procena; uključuje dokaz u oba smera, ne uključuje
čekanje na Windows prolaz). Gde se dokazuje: `L` u celosti na Linux-u i u CI-ju,
`W` traži Windows + Excel. U zagradi iza naslova stoji oznaka iste stavke u
analizi (na primer `F1`, `RV-04`, `D4`), gde su dokaz i brojevi u celini.

---

## 1) Rezime

**Ocena.** Strogoća sistema je iznad proseka i ostaje: dvosmerni dokaz po imenu
tvrdnje, self-test svakog checkera kroz produkcionu putanju, marker zelenog vezan
za otisak izvora. Neprofesionalan je **model izvršavanja**: sve se plaća iznova,
na svakom mestu, i ništa se ne meri.

**Problem u četiri broja.**

| Šta | Koliko |
|---|---|
| Hook posle svake VBA izmene | 9–14 s `[M]`; kod malog fajla je 100% tog vremena posao koji ne zavisi od izmenjenog fajla |
| Pun dvosmerni dokaz kataloga | 1265 prolaza `[M]`; 44–82 h pojedinačno, 11–20 h u grupama po 6 `[I]`; poslednji put pušten 28.08.2026 nad 261 sabotažom `[D]` |
| Rast | za 30 dana sabotaže +76%, testovi +67%, produkcioni VBA +9% `[M]`; trajanje `RunAllTests` 0,10 → 0,64 s po testu `[D]` |
| Merenje | u 1560 commit-a oko 14 upotrebljivih brojeva o trajanju, svi u prozi `[M]` |

**Strategija u pet poteza.**

1. **Meri pre nego što menjaš.** Trajanje po fazi poziva, po suiti, po testu i po
   prolazu dokaza ide u mašinski zapis.
2. **FAST sloj:** algoritamski rezovi sa bajt-identičnim izlazom. Prototip postoji
   i izmeren je.
3. **Dokaz:** sa „cela suite po mutaciji" na „ciljani test po mutaciji" (cena
   postaje linearna), zatim manja režija poziva, zatim paralelne instance.
4. **Ritam:** pun dokaz se vezuje za kalendar (noćno), ne za release, a ishod
   Excel prolaza beleži alat, ne rečenica u dnevniku.
5. **Struktura:** tri fajla-žarišta (`sabotaza.py`, `modTest.bas`,
   `modBusinessFlowProTests.bas`) se dele pre nego što udare u VBA granice.

**Očekivani ishod.**

| Mera | Danas | Posle talasa 1 (bez Excela) | Posle talasa 3 | Posle talasa 5 |
|---|---|---|---|---|
| Hook po Edit-u | 9–14 s `[M]` | 3–5 s `[M, prototip]` | isto | do 2 s tipično `[M, prototip keša]` |
| Pun `vba_check` | 30–34 s `[M]` | 13–14 s `[M, prototip]` | isto | isto |
| CI job `static` | medijana 76 s `[M]` | 30–35 s `[I]` | isto | isto |
| Dokaz jednog reza (4–7 sabotaža) | oko 34 min `[D]` | isto | 3–5 min `[I]` | isto |
| Pun pojedinačni dokaz | 44–82 h `[I]` | isto | 10–14 h `[I]` | do 5 h na jednoj mašini, do 2 h sa 3–4 radnika `[I]` |
| Starost najstarijeg dokaza u katalogu | oko 11 nedelja `[D]` | isto | isto | do 7 dana |
| Excel prolazi sa zabeleženim trajanjem | približno 0% | isto | 100% | 100% |

Sve vrednosti u kolonama „posle" koje se tiču Excela su model. Kapija odluke G0
(§8) postoji upravo zato da se posle prvog mernog prolaza redosled potvrdi ili
promeni.

---

## 2) Polazno stanje

### 2.1 FAST sloj i CI

| Mera | Vrednost | Izvor |
|---|---|---|
| Hook nad malim ili srednjim `.bas` fajlom | 9,2–10,3 s; na manje opterećenoj mašini 6,2–6,3 s | `[M]` |
| Hook nad `modTest.bas` / BFP | 12,0–13,7 s | `[M]` |
| Deo hook-a nezavisan od izmenjenog fajla | 9,05 od 9,05 s (mali fajl); 9,4 od 12,6 s (`modTest`) | `[M]` |
| Pun `vba_check.py` | 30–34 s | `[M]` |
| `vba_check.py --self-test` | 17,4–17,8 s; 152 slučaja zajedno ispod 0,2 s, ostatak su dva puna prolaza globalnih kapija | `[M]` |
| Lokalni pun set kapija | 51,4 s redom; 22,8 s paralelno | `[M]` |
| CI job | medijana 76 s (25 poslednjih run-ova); 12 s u avgustu | `[M]` |
| Rast hook-a na istoj mašini | 0,32 s (16.08) → 6,19 s (08.11); oko 4,5 ms po sabotaži | `[M]` |
| Dnevno čekanje na hook | medijana 143 VBA hunk-a dnevno × oko 11 s = oko 27 min; p90 dan oko 90 min | `[I]` |

### 2.2 Excel prolazi

| Mera | Vrednost | Izvor |
|---|---|---|
| Režija jednog `run_vba` poziva | oko 18 s; od toga 4,5–7,9 s su fiksni `sleep`-ovi u kodu | `[D]`; sleep `[M]` iz koda |
| `RunAllTests` | 215,8 s pri 336 testova (05.11); 73,5 s pri 204 (09.10) | `[D]` |
| `RunBusinessFlowProSuite` | 210 s (07.10); oko 350–390 s početkom novembra | `[D]`; `[I]` |
| FULL (12 suita) | ne staje u podrazumevanih 600 s; zapisi „oko 37 min" i „20–60 min" | `[D]` |
| Dokaz reza | 34 min za 4 BFP sabotaže (04.11; procena pre puštanja 20–25) | `[D]` |
| Jedna runda prijema reza | 7–20 podizanja Excela, 21–88 min | `[D]` |
| Pun katalog | 1265 prolaza pojedinačno, 215 uz `--grupe 6` | `[M]` kroz `--plan` |
| Trajanje punog kataloga | 44–82 h pojedinačno; 11–20 h grupno | `[I]` |
| Poslednji pun katalog | 28.08.2026, 260/261 | `[D]` |

### 2.3 Rast i održavanje

| Mera | Vrednost | Izvor |
|---|---|---|
| Sabotaže | 718 → 1265 za 30 dana; +18,7 dnevno | `[M]` |
| Testovi u `RunAllTests` | 209 → 350 | `[M]` |
| Test moduli / produkcioni VBA | +40% / +9% (linije); odnos test/prod 0,17 → 0,51 | `[M]` |
| `modTest.bas` | 31.061 linija, +452 dnevno; granica 65.534 linije za oko 76 dana | `[M]`; `[I]` |
| BFP modul | 33.077 linija, +323 dnevno; granica za oko 100 dana | `[M]`; `[I]` |
| `TestName` (registar) | 22.178 bajtova; prag `PROCEDURA_VELIKA` 40.000 za oko 60 dana | `[M]`; `[I]` |
| Presidrenje kataloga | 64 u 29 dana (21 od 50 merge koraka) | `[M]` |
| Commit-i koji diraju samo `docs/` | 52,7%; `STANJE_REFAKTORA.md` 844 KB | `[M]` |

Da li granica od 65.534 linije važi u instaliranoj verziji Excela nije provereno
`[?]`; stavka L-3 to proverava jednom, u Excelu.

---

## 3) Dijagnoza: osam uzroka

1. **Kvadratna cena dokaza.** Jedna mutacija pušta celu suitu (350 ili 391 test)
   da bi posmatrala jednu tvrdnju. Broj prolaza raste sa katalogom, a trajanje
   prolaza sa brojem testova. Projekcija za šest meseci uz današnji rast: 129–256 h
   u grupama `[I]`.
2. **Fiksna režija po pozivu.** Svaki `run_vba` diže nov Excel, uvozi 192
   komponente (10,4 MB), radi compile probu sa fiksnim čekanjem, šemu, pa gašenje.
   Mutacija pritom dira jedan modul od prosečno 170 KB, to jest 1,6% izvora `[M]`.
3. **Globalan posao na lokalan događaj.** Hook posle izmene jednog fajla ponovo
   parsira ceo katalog i ceo `src-vba`. Tri uzroka su čisto algoritamska:
   `_dodela_tudjoj_proceduri` obrađuje 214,5 MB teksta umesto 5,7 MB jedinstvenog;
   `collect_arities` tokenizuje svaku liniju dvaput; `check_scalar_call` čisti
   istu liniju jednom po skalarnom imenu `[M]`.
4. **Suite plaćaju sporedne efekte koje nijedna tvrdnja ne meri.** `RunAllTests`
   i BFP ne pale tihi režim, pa svaki `AppendRow` piše CSV journal, a svaki
   `CommitTx` log i `OnTime`; BFP upisuje 7 ćelija po tvrdnji; 51 test gradi celu
   ljusku forme; tabele kroz prolaz samo rastu `[M]`. Koliko to košta: `[?]`.
5. **Ništa se ne meri mašinski.** `last_run.json` nema ni trajanje ni SHA i
   prepisuje se; `dokaz.py` ne meri ništa. Brojevi u pravilima su ustajali 2–3
   puta. Posledica je ona iz postmortema §10: pravilo koje košta više nego što se
   plati tiho se svede na podskup.
6. **Izvršavanje vezano za jednu mašinu i radno stablo.** Mutacija se radi na
   mestu, pa je `src-vba` zaključan dok dokaz traje. Fiksne putanje izveštaja,
   markera i jednog `%TEMP%` foldera, uz `SendKeys` koji traži fokus, sprečavaju
   paralelne instance i noćni rad.
7. **Tri fajla-žarišta.** `sabotaza.py` (735 KB, diran u 35 od 40 PR-ova),
   `modTest.bas` (28 od 40) i BFP modul (25 od 40) `[M]`. Tu se sudara i grana sa
   izborom testova. Tri ručna registra testova imaju 1.106 linija.
8. **Praksa se odvojila od pisanih pravila.** „Jeftine i skupe kapije", `--grupe`
   i `--require-green` vode svakodnevni rad, a ne postoje u `CLAUDE.md` ni u
   `.claude/rules`. `CLAUDE.md` §5 opisuje cenu dokaza koja ne odgovara kodu.

---

## 4) Nalazi o tačnosti

Ovo nisu ubrzanja. Ulaze u plan jer je svaki rupa u onome što sistem već tvrdi, a
većina se zatvara usput.

| # | Nalaz | Izvor | Stavka |
|---|---|---|---|
| 1 | Hook se ne pokreće kad se menja `tools/sabotaza.py`, iako komentar u `vba_check.py` i dokumentacija tvrde da se pokreće | `[M]`: rc=0 za 0,06 s | F-6 |
| 2 | `sabotaza.py --self-test` (43 slučaja) ne vrti ni CI ni `vba_check` | `[M]` | L-4 |
| 3 | Hook ne vidi lom preko granice fajla: izmena potpisa `LogErr` daje rc=0, a pun prolaz 492 nalaza `ARNOST` u 78 fajlova | `[M]` | F-9 |
| 4 | Ne-ASCII bajt u fajlu koji čita katalog ruši `vba_check` traceback-om, bez imena fajla i reda | `[M]` | L-5 |
| 5 | Kad `run_vba` ubije Excel na 600 s, `dokaz.py` to prijavi kao „NE OBARA NISTA", ne kao prekid | `[M]` iz koda | L-1 |
| 6 | Zastavica `dialogs` u `SUITES` se nigde ne čita; „neočekivan dijalog = pad" iz dokumentacije nije implementirano | `[M]` iz koda | S-3 |
| 7 | KI-036: prolaz koji pušta dokaz upiše zeleni marker nad sabotiranim izvorom | `[D]` | L-2 |
| 8 | `release.sh` ne zove nijednu kapiju i taguje pre ručnog Compile-a; bump verzije menja otisak izvora, pa poništava marker | `[M]` iz koda; drugi deo `[I]`, nije izvršeno | N-3 |
| 9 | Osam podrazumevanih gate-suita nema rezultat-fajl, pa ime pale tvrdnje ne preživi COM granicu | `[M]` | K-3 |
| 10 | Kvar u `_strip_comment` ostavlja `vba_check --self-test` zelenim | `[M]` | F-1 |
| 11 | `popis_citalaca.py --check` postoji samo lokalno i crven je na `main`-u (33 preko praga 31) | `[M]` | F-5 |
| 12 | Sedam priznatih statičkih nalaza u katalogu nije u spisku koji `dokaz.py` čita, pa pun katalog verovatno ne može da vrati 1265/1265 | `[I]`, nije pušteno | D-3 |
| 13 | PWA: dva fajla koja `index.html` učitava nisu u `ASSETS` u `sw.js`; keširan fajl je menjan bez promene `CACHE_NAME` | `[M]` | F-11 |

---

## 5) Načela

### 5.1 Tvrdnje koje ostaju netaknute

Katalog tvrdnji sa incidentima koji su ih doneli je u analizi
`06-pravila-i-incidenti.md` (T1–T26). Plan ih ne menja. Najvažnije:

- baza je zelena pre prve mutacije; pala je baš ciljana tvrdnja, po imenu;
- izvor se vraća i kad prolaz pukne, i to se proverava otiskom;
- nema rezultat-fajla = pad; BLIND nije dokaz; „nejasno" se prijavljuje kao nejasno;
- zeleno važi samo za baš ovaj izvor pod ovim test sistemom (marker);
- self-test checkera ide kroz produkcionu putanju, u oba smera, i vrti se u CI-ju;
- grupni dokaz tvrdi manje i tako se zove;
- bez Excela se izmena ponašanja prijavljuje kao neverifikovana.

### 5.2 Pravila ovog plana

- **N1. Nov režim nosi svoje ime.** Kao što danas postoji „DOKAZANO (grupno)",
  svaki režim koji meri drugačije dobija oznaku u verdiktu. Nijedna optimizacija
  ne menja značenje postojećeg verdikta.
- **N2. Skraćeni put deli kod sa punim i dokazuje se diferencijalno.** Isti
  verdikt po imenu nad uzorkom, jednom nad celim skupom, i crveno po imenu kad se
  skraćeni put namerno pokvari.
- **N3. Preskočeno se broji.** Sve što neki režim ne izmeri ulazi u imenilac kao
  „NIJE IZMERENO". Nikad se ne prećutkuje.
- **N4. Merenje pre i posle, na istoj mašini.** Gde je moguće, kapija meri
  **brojač rada** (obrađeni bajtovi, broj poziva), a ne štopericu.
- **N5. Keš je fail-closed.** Ključ je sadržaj plus otisak alata; bez keša se
  računa sve; CI uvek radi bez keša.
- **N6. Izmena ugovora se grupiše.** `run_vba.py`, `vba_gate.py` i
  `make_fixture.py` ulaze u otisak ugovora, pa svaka njihova izmena obara sve
  zelene markere. Izmene se spajaju u što manje PR-ova, da se FULL plati jednom.
- **N7. `.claude/` ide zasebnim process PR-om, jedan po jedan.**
- **N8. Mali rezovi.** Svaka stavka je zaseban, povratan rez sa sopstvenim
  kriterijumom prihvatanja.

---

## 6) Ciljna slika: slojevi

| Sloj | Okidač | Šta se pušta | Budžet | Šta tvrdi |
|---|---|---|---|---|
| **Edit** | `PostToolUse` hook | provere izmenjenog fajla i globalne kapije (iz keša činjenica) | do 4 s, zatim do 2 s | mehanička ispravnost; katalog je živ |
| **Commit / push** | jedna komanda | pun `vba_check`, `who_writes` oba režima, `gen_schema --check`, popis suita | do 25 s | isti skup koji pušta CI |
| **Rez** | posle „CODE GO" | TARGETED suite i dokaz nad novim i dirnutim sabotažama koje bira alat | do 10 min | nove i dirnute tvrdnje su dokazane |
| **PR** | CI, i jedan završni prijem na glavi koja se spaja | `static`; `vba_gate --require-green --require-compile` | CI do 40 s | spojeni izvor je meren |
| **Noćno** | zakazan zadatak, zaseban klon | FULL nad `main`-om, rotacioni isečak dokaza, revizije | do 3 h | regresija se vidi za najviše 24 h; nijedan dokaz stariji od N dana |
| **Nedeljno** | vikend | grupni prolaz cele suite nad celim katalogom; obrnut redosled testova | jedna noć | kontekst redosleda; kill-matrica |
| **Release** | ručno | bump → FULL, pun dokaz, ručni Compile nad tim izvorom → `vba_gate` → tag | jedna noć | isporučuje se baš ono što je mereno |

Dva sloja su nova (noćno i nedeljno). Ostali postoje u praksi, ali nisu zapisani
ili nemaju budžet.

---

## 7) Radne stavke

### Faza 0 — Temelj: merenje i zatvaranje litica

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **M-1** | Prenos izbora testova na tekući `main` (`D4`, `RV-11`) | otključava fazu 3 | VBA strana nikad izvršena | M | L + W | Linux: `vba_check`, self-test `run_vba` / `vba_gate` / `dokaz`, `--proveri-sidra`, 23 dvosmerne provere. Windows: ručni Compile; četiri sabotaže `izbor-*` crvene po imenu |
| **M-2** | Trajanje po fazi i istorija prolaza u `run_vba` (`RV-01`, `IM-1`) | svaka dalja procena postaje merenje | menja ugovor (N6) | S | L | self-test: lažan izveštaj nosi sve faze, `REZIJA` jednaka zbiru; uklonjen dopis → crveno po imenu |
| **M-3** | Knjiga dokaza: red po prolazu sa svim palim tvrdnjama, trajanjem i režimom; `--plan` daje procenu iz istorije (`D1`, `IM-2`) | cena reza poznata pre puštanja; ukida ručni prepis | nizak | S | L | self-test sa lažnim runner-om: red za svaki ishod, nijedan dvaput |
| **M-4** | Fajl napretka po testu: red pre i posle svakog testa (`T-02`) | ubijen prolaz imenuje test koji je radio | format koji čita `dokaz.py` mora ostati isti | S | W | prekid usred suite → izlaz imenuje taj test |
| **M-5** | Dijagnostički otisak tabela pre i posle svakog testa (`T-01`, `P5`) | tačan spisak testova koji ostavljaju stanje (danas statička granica 96–121) | oko 50 čitanja po testu | S | W | test koji namerno ostavi red je imenovan; kad se red vrati, nije |
| **M-6** | `tools/metrike.py`: rast i održavanje iz git-a, tabela u CI rezimeu (`IM-5`) | rast vidljiv nedeljno | nizak | S | L | brojevi nad `HEAD`-om se slažu sa `sabotaza.py --lista` i `vba_gate --popis` |
| **L-1** | Prekid po suiti; `dokaz.py` prosleđuje izričit `--timeout`; prekid se zove TIMEOUT (`RV-04`, `IM-3`, `D2d`) | uklanja liticu: BFP oko 370 s + režija prema granici od 600 s | prekratak budžet daje lažan TIMEOUT (fail-closed) | S–M | L + W | self-test: tajmer se naoružava po suiti; izlaz sa FATAL daje TIMEOUT, ne „NE OBARA NISTA" |
| **L-2** | KI-036: mutirani prolaz ne piše zeleni marker (`P7`, `IM-9`) | jedan ponovljen prolaz manje posle svakog dokaza | menja ugovor (N6) | S | L + W | self-test: mutirani poziv nosi prekidač, bazni ne; Windows: FULL → dokaz → `--require-green` ostaje zelen |
| **L-3** | `MODUL_VELIK` u `vba_check` (upozorenje 45.000, greška 60.000 linija) i jedna proba granice u Excelu (`T-05`) | granica se vidi mesecima ranije | lažna uzbuna ako je prag nizak | S | L + W | self-test u oba smera po uzoru na `PROCEDURA_VELIKA` |
| **L-4** | `sabotaza.py --self-test` u CI; self-test piše u privremeni folder, ne u `src-vba` (`F8a`) | dokaz za najskuplju proveru hook-a postaje automatski | nizak | S | L | namerni kvar `_dodela_tudjoj_proceduri` → CI korak crven |
| **L-5** | Ne-ASCII u fajlu kataloga daje nalaz `ASCII` sa fajlom i redom (`F6`) | tačna poruka za najčešću mehaničku grešku | vrlo nizak | S | L | nov slučaj kroz ceo CLI; pre izmene pada traceback-om |
| **M-7** | Prvi merni prolaz na Windows-u (Prilog A) | popunjava sve `[?]` iz §12 | — | 1–1,5 h | W | tabela iz §2.2 nosi `[M]` |

Napomena uz M-1: grana nosi jedan commit, a `main` je odmakao (158 commit-a na
dan pisanja). Python strana se primenjuje bez sudara. VBA strana se ne rešava
ručno: nov test dobija sledeći slobodan broj, omotavanje BFP poziva se ponovo
generiše skriptom nad tekućim telom suite, a četiri sabotaže idu na kraj
kataloga. Skripte su u `analize/2026-11-verifikacija/prenos-izbora/`.

### Faza 1 — FAST sloj i CI

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **F-1** | Čuvari pre ubrzanja: self-test slučajevi za tokenizere; brojač obrađenih bajtova u proveri sidara (`F8b`, `F8c`) | ubrzanje ne može ući bez provere koja ume da pocrveni | nizak | S | L | kvar u `_strip_comment` (danas zelen) pada po imenu |
| **F-2** | Egzaktni rezovi: rani izlaz i inkrementalno parsiranje u proveri sidara; jedan prolaz tokenizera; čišćenje tela jednom po proceduri (`F1`–`F4`) | hook 9–10 → 3–4 s; pun prolaz 30–34 → 13–14 s; self-test 17 → 6 s `[M, prototip]` | pet tokenizera je osnova gotovo svih provera | S | L | izlaz bajt-identičan nad čistim repoom u tri režima, nad dve „prljave" kopije (261 i 434 nalaza), nad 301.079 redova po funkciji i nad 1265 + 6000 zamena; tri namerna kvara padaju po imenu |
| **F-3** | Self-test ne ponavlja globalne kapije u dva „ceo CLI" slučaja; nov slučaj tvrdi da `main` zove sve tri kapije alata (`F5`, `CI-1`) | 17 s → 2–3 s | isključenje ne sme da procuri u produkcioni poziv | S | L | otkačen poziv kapije u `main` → nov slučaj pada |
| **F-4** | CI: `timeout-minutes`, `concurrency` za PR, Node 22; četiri paralelna job-a i zbirni `static`; `who_writes` u jednom prolazu (`CI-2`, `CI-3`, `CI-11`) | zid 76 s → 30–35 s `[I]`; pad jedne kapije ne sakriva ostale | ime obaveznog check-a mora ostati isto | S | L | namerno crven korak → ostali job-ovi daju verdikt, zbirni je crven |
| **F-5** | Kapije koje postoje samo lokalno ulaze u CI (`CI-4`) | zatvara klasu „kapija postoji, niko je ne pušta" | dve nemaju self-test i moraju ga dobiti pre toga | M | L | po jedan namerno crven commit na test grani za svaki nov korak |
| **F-6** | Hook pokriva `tools/sabotaza.py`, `dokaz.py`, `vba_gate.py`, `run_vba.py` (`F7`, `P2`) — **process PR** | greška u katalogu se vidi u istom koraku | nizak | S | L | kroz pravi hook: pokvareno sidro u kopiji kataloga → rc=2 (danas rc=0) |
| **F-7** | Manifest kapija i lokalni pokretač; `--check` poredi manifest, `static.yml` i tabele u pravilima (`P3`, `CI-5`) | jedna komanda umesto 4–22; drift obara CI | nov checker | M | L | korak samo u `static.yml`, korak samo u manifestu i izmenjena tabela padaju po imenu |
| **F-8** | Keš činjenica po fajlu, ključ sadržaj + otisak alata (`F9`) | hook 0,2 s van kataloga, 1,1–1,7 s u katalogu, 3,1–3,8 s za `modTest`/BFP `[M, prototip]` | ustajao keš daje lažno zeleno; zato N5 | M | L | 16 scenarija izmene sa identičnim izlazom; pokvaren ključ → scenario „vraćena izmena" pada |
| **F-9** | Hook sa semantikom punog prolaza: lom pozivalaca vidljiv odmah (`F10`) | zatvara nalaz 3 iz §4 | tabela zavisnosti mora biti potpuna; blokira međukorake refaktora | L | L | ekvivalencija sa hladnim punim prolazom posle niza nasumičnih izmena |
| **F-10** | Budžet trajanja kao kapija: brojač rada odmah, štoperica u CI prvo kao upozorenje (`P12`, `IM-6`) | cena ne može tiho da se vrati | lažna uzbuna na sporom runner-u | S | L | veštačko kašnjenje → prekoračenje po imenu |
| **F-11** | JS harness jasan bez `node_modules`; statička kapija nad PWA ljuskom (`CI-6`, `CI-7`) | prva kapija nad delom koji se isporučuje samim merge-om | deo je izmena ponašanja PWA, traži odluku | S | L | fiksture u memoriji u oba smera |

**Odluka unutar faze.** Analiza pravila je predložila i sužavanje hook-a na
sabotaže izmenjenog fajla (`P1` u `06-…`). Ne preporučuje se: odlaže proveru
celog kataloga do commit-a, a F-2 i F-8 daju isto ubrzanje bez tog slabljenja.

### Faza 2 — Suite: skinuti trošak koji nijedna tvrdnja ne meri

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **S-1** | Tihi režim journal-a i loga u `RunAllTests`, BFP i ostalim suitama; gašenje `OnTime` tajmera na kraju suite; jedan imenovan test tvrdi journal i zakazivanje, sa svojom sabotažom (`T-03`) | hipoteza 5–25% trajanja `[?]`; uklanja kandidata za „RPC failed" posle 10 min | nizak; 0 sabotaža je usidreno u te module | S | W | isti skup imena i isti `TESTS`/`FAIL`; nov test pada na svoju sabotažu |
| **S-2** | BFP log u memoriji umesto 7 ćelija po tvrdnji; VBE prozor se sakriva posle compile probe (`T-04`) | 1–5 s po BFP prolazu `[?]` | bafer se gubi ako Excel padne, zato M-4 pre | S | W | list sa logom identičan po redovima; `last_run_bfp.txt` identičan |
| **S-3** | „Neočekivan dijalog = pad" se sprovodi; `MsgBox` se preskače pod harnesom (`RV-10`) | pouzdanost; oko 11 dijaloga manje po FULL-u | dira 7 test modula | S–M | L + W | dijalog u suiti sa `dialogs: False` → rc=2 po imenu; FULL bez reda `DIALOG` |
| **S-4** | Revizija nezavisnosti: svaki test sam (`--izolacija`), obrnut redosled, 3–5 permutacija sa zapisanim semenom; zavisnosti postaju deklaracije (`T-09`, `P5`) | izmeren spisak zavisnih testova; preduslov za D-2, S-6, K-4 | može otkriti više zavisnosti nego što se očekuje | M | W | isti skup PASS/FAIL po imenu u svim permutacijama; namerno zavisan par je prijavljen |
| **S-5** | Trijaža najsporijih testova po izmerenom vremenu; vruća putanja pristupa podacima samo ako merenje pokaže, svaki rez uz `pre-flight` (`T-13`) | nepoznat dok M-7 ne da brojeve | visok: jezgro kroz koje ide svaki upis | M po rezu | W | vreme po testu pre i posle; nov test poništavanja keša crven svojom sabotažom |
| **S-6** | Podrazumevane suite u dva paralelna Excela (`T-10`) | FULL preko 600 s → oko 265–330 s `[I]` | pravilo unije u markeru je mesto za lažno zeleno | M | L + W | isti skup imena i verdikata; tri negativna slučaja markera u self-testu |
| **S-7** | Provera okruženja pre podizanja Excela: slobodan disk (`IM-7`) | sprečava „tri prolaza za nula nalaza" | nizak | S | L | lažiran disk ispod praga → odbijeno po imenu |

### Faza 3 — Dokaz: linearna cena

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **D-1** | Sitna režija: pamćenje `_suita_testa`; `--vrati` samo nad primenjenim članovima; bazne suite u jednom pozivu (`D2`, `P8c`) | 27 s → ispod 1 s po pozivu; 1,0 s → 0,05 s po prolazu | nizak | S | L | isti plan (215 prolaza, ista podela); self-test oba alata zelen |
| **D-2** | Izbor testa kao podrazumevani put dokaza u rezu, posle diferencijalnog prijema (`D4`, `P2`) | prolaz 100–400 s → 20–26 s; rez 34 min → 3–5 min; rast postaje linearan `[I]` | testovi koji sami ne prolaze padaju na celu suitu (udeo `[?]`) | M | W | isti rez u oba režima, verdikt identičan po stavci, nad najmanje 50 sabotaža iz obe velike suite; sabotaža koja u izboru ne obori svoju tvrdnju meri se celom suitom pre verdikta „NIJE DOKAZANO" |
| **D-3** | Kontrolne sabotaže u svakom pozivu: placebo mutacija koja menja samo komentar ne sme biti priznata (`D12`) | dvosmerni dokaz samog dokaza u Excelu; razrešava nalaz 12 iz §4 | jedan prolaz više | S | L + W | lažan izlaz u kome placebo „obara" tvrdnju → NIJE DOKAZANO |
| **D-4** | Baza iz markera zelenog kad marker važi za isti otisak; inače se baza pušta (`D3`, `RV-12`) | 2–9 min po pozivu `[I]` | marker ne hvata nedeterminizam između dva puštanja | S | L | važeći marker → nula baznih poziva; drugi otisak → baza se pušta |
| **D-5** | Obim reza bira alat: `dokaz.py --od <ref>` (izmenjeni unosi, sabotaže čiji je fajl, procedura ili test diran) (`D8`, `P4`) | nema propuštenih i odloženih „dirnutih" stavki (2 od 11 odloženih je istrunulo) | pravilo o pomoćnom kodu može izabrati previše | M | L | sintetički diff: svaka vrsta pogotka bira svoju stavku; nad diff-om stavke 288 bira svih 11 |
| **D-6** | Mutacija van radnog stabla: `run_vba --src <kopija>`, `dokaz.py` nikad ne piše u `src-vba` (`D5`, `RV-05`) | radno stablo nije zaključano; nestaje klasa REVERT-FAIL; otvara paralelne instance | `sabotaza.py` primena traži ciljni direktorijum | M | L + W | otisak `src-vba` isti pre, tokom i posle; ista sabotaža starim i novim putem daje isti verdikt |
| **D-7** | Nastavak prekinutog dokaza po konzervativnom ključu (ceo izvor, unos, sveska, alati, režim) (`D7`) | prekid posle 20 od 40 h ne košta 20 h | nizak | S | L | promena bilo kog sastojka ključa → stavka se ponovo meri |
| **D-8** | Statička predprovera novih i izmenjenih sabotaža: mutirani tekst kroz provere `vba_check` po fajlu (`P7c`) | sabotaža koja obara compile se hvata bez Excela | nizak | S | L | sabotaža sa duplom deklaracijom → statički nalaz |

**Šta izbor slabi i kako se to zatvara.** Izbor dokazuje da je test osetljiv na
mutaciju kad se pusti sam, uz bazu nad istim izborom. Ne dokazuje da je osetljiv
u kanonskom redosledu cele suite: raniji test može ostaviti stanje koje mutaciju
zaklanja. Zato verdikt nosi oznaku „DOKAZANO (izbor)" (N1), a razlika se zatvara
sa tri mere: izmerenom nezavisnošću testova (S-4), nedeljnim grupnim prolazom
cele suite (N-1) i revizijom uzorkom u punoj suiti (N-5). Posle podele po oblasti
(K-4) izbor može da ide na nivou oblasti, pa redosled unutar oblasti ostaje.

### Faza 4 — Režija poziva i paralelizam

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **R-1** | Čekanje na događaj umesto fiksnih `sleep`-ova; compile proba samo tamo gde nešto tvrdi (`RV-02`, `RV-03`) | 1,5–4 s i 2,5–5,9 s po pozivu `[I]` | slomljen izvor se u suite režimu vidi kasnije; zato L-1 pre | S–M | L + W | `--compile-only` daje isti verdikt nad čistim i nad slomljenim modulom; 20 poziva bez zaostalog procesa |
| **R-2** | Proba: 50 puta zamena 6 modula u uvezenoj, kompajliranoj svesci (`RV-06`, preduslov) | odgovor na pitanje da li je R-3 izvodljiv | — | S | W | bez pada VBE-a i bez modula sa sufiksom `1` |
| **R-3** | Jedan uvoz po pozivu dokaza i zamena samo izmenjenih modula u svežoj kopiji sveske (`RV-06`, `D6a`) | režija 18 → 6–8 s `[I]` | ponašanje `Remove` + `Import` nad kompajliranim projektom `[?]` | M | W | isti skup sabotaža hladno i delta putem daje isti verdikt; prazna delta → ista baza; pokvaren read-back pada po imenu modula |
| **R-4** | Paralelni Excel radnici, svaki nad svojom kopijom izvora; podela po izmerenom vremenu (`RV-09`, `P4`) | približno N puta kraće, realno 2–4× `[I]` | fokus i memorija; mašina zauzeta | M | W | isti skup serijski i paralelno, tri ponavljanja bez razlike; dva radnika sa istim izlazom → alat odbija |
| **R-5** | Topla sesija sa „sendvičem": mutacija, test, vraćanje, isti test mora biti zelen; recikliranje na N mutacija (`RV-08`, `P3`) | preostala režija 6–8 → 2–4 s `[I]` | visok: `OnTime`, zastavice aplikacije, break mode | L | W | ekvivalencija nad celim katalogom; namerno „curava" pseudo-sabotaža uhvaćena po imenu |

R-5 se radi samo ako posle R-3 start i gašenje Excela i dalje nose više od oko
30% prolaza (kapija G2).

### Faza 5 — Ritam i evidencija

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **N-1** | Noćni zadatak u zasebnom klonu (Task Scheduler, prijavljena sesija): FULL nad `main`-om i rotacioni dokaz sa knjigom; nedeljno grupni prolaz cele suite (`CI-9`, `P6`, `D9`) | regresija vidljiva za 24 h; katalog svež | mašina mora biti upaljena i prijavljena | M–L | W | namerno pokvaren test → jutarnji ishod crven po imenu suite; ubijen proces → klon čist; tri noći bez intervencije |
| **N-2** | Objava dokaza po otisku izvora (commit status) i mašinski zapis po PR-u umesto prepisa u dnevnik (`CI-8`, `IM-8`) | oko 150 ručnih docs commit-a mesečno manje | status postavlja operaterov token; lokalna putanja ne sme u javni repo | M | L + W | pao prolaz → nema objave; izmenjen izvor → „drugi izvor" |
| **N-3** | Release kapija: bump → FULL, dokaz, Compile nad tim izvorom → `vba_gate` → tek tada tag (`P10a`, `CI-10`) | tag se ne može postaviti nad nedokazanim izvorom | pogrešan redosled blokira hitno izdanje; treba svestan, zabeležen izlaz | S–M | L | suvo pokretanje u klonu bez markera staje pre taga |
| **N-4** | Politika obima dokaza pred release (odluka O-1) | pravilo postaje izvršivo | zavisi od izabrane varijante | odluka | — | — |
| **N-5** | Revizija uzorkom, isključivo kao dopuna: n sabotaža dokazanih u izboru meri se pojedinačno u punoj suiti; granica piše u verdiktu (`P6`) | 0 neslaganja u 100 → najviše 2,95% uz 95% | statistička tvrdnja se lako pročita kao univerzalna | S | L + W | isto seme → isti uzorak; pokvaren zapis u uzorku je nađen |
| **N-6** | Registar dokaza preko commit-a (ključ po proceduri i zavisnostima): **samo istraživanje, u senci** (`P1`, `D10`, `P11`) | po PR-u bi se ponovo merilo oko 2–12% kataloga `[M]` | visok: propuštena zavisnost tiho potvrđuje zastarelo | L | L + W | kill-matrica iz nedeljnog prolaza: procedura svake sabotaže mora biti u zatvaranju njenog testa |

### Faza 6 — Struktura

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **K-1** | Katalog po modulu, bez escape-a; alat `--presidri` (`D11`, `P10a`) | sudari ostaju u jednom modulu; presidrenje dobija mehanički deo | jednokratna velika izmena fajla koji dira skoro svaki PR | M | L | rečnik iz starog fajla jednak rečniku iz novih, uključujući redosled; `--lista` identičan |
| **K-2** | Generisan registar testova, identitet je ime (`T-06`) | nov test je jedna procedura umesto četiri izmene; nestaje rok za `TestName` | nov checker; prva generacija menja oko 1.100 linija | M | L + W | istih 350 imena istim redom; `--check` pada na sva četiri današnja razlaza |
| **K-3** | Jedan test-kit i jedan format rezultata za sve suite (`T-07`) | osam gate-suita dobija ime pale tvrdnje; tvrdnje mogu da žive u svom modulu | dira 10+ test modula; radi se suita po suita | M | W | isti broj provera po suiti; namerno oborena tvrdnja stiže po imenu |
| **K-4** | Podela `modTest` i BFP po oblasti, skriptom; suita po oblasti (`T-08`) | TARGETED oko 6× kraći `[I]`; uklanja rok od 65.534 linije | visok po obimu; sudar sa svakom otvorenom granom | L | W | unija rezultata = ista imena i isti zbir tvrdnji; pun dokaz jednom posle podele |

**Rokovi koje nameće rast** `[I]`: K-2 pre kraja decembra 2026 (`TestName`), K-4
pre sredine januara 2027 (`modTest.bas`). L-3 iz faze 0 ih pretvara u alarm.

### Faza 7 — Pravila i dokumenti

| ID | Stavka | Put |
|---|---|---|
| **P-1** | Hook pokriva alatne fajlove (isto što F-6) | process PR |
| **P-2** | `CLAUDE.md` §5 i `.claude/rules/testovi.md` usklađeni sa kodom: jedna suita po prolazu, bez zakucanih trajanja, `/` umesto `\`, slojevi iz manifesta, `paths:` dopunjen (`P11`, `IM-10`) | process PR, posle F-7 |
| **P-3** | Zastareli brojevi i tvrdnje van `.claude/`: docstring `dokaz.py`, `EXCEL_TEST_HARNESS.md`, komentar hook-a u `vba_check.py`, postmortem §7/§8 | običan docs PR |
| **P-4** | Zapisati slojevitu politiku iz §6, uključujući ishod odluka O-1, O-4 i O-8 | process PR, posle odluka |
| **P-5** | `STANJE_REFAKTORA.md`: arhiva starih stavki, bez ponavljanja rezultata kapija (`P13`) | običan docs PR, odluka O-9 |

---

## 8) Redosled, zavisnosti i kapije odluke

### 8.1 Talasi

| Talas | Sadržaj | Traži Excel | Ishod |
|---|---|---|---|
| **1** | F-1, F-2, F-3, L-3, L-4, L-5, D-1, M-6, F-4 | ne | hook 3–5 s; CI 30–35 s; sve proverljivo u CI-ju |
| **1b** | jedan PR nad ugovorom: M-2, M-3, L-1, L-2 | FULL jednom | merenje i zatvorene litice |
| **2** | M-1 (prenos izbora), pa M-7 (merni prolaz) | da | brojevi umesto modela → **G0** |
| **3** | D-2, D-3, D-4, S-1, S-2, S-3, M-4, M-5, S-4, F-6 | da | dokaz reza 3–5 min → **G1** |
| **4** | D-5, D-6, D-7, D-8, R-1, R-2, R-3, K-2, F-7 | da | pun pojedinačni dokaz reda 10 h → **G2** |
| **5** | N-1, N-2, N-3, N-5, R-4, K-1, K-3, F-8, F-10, S-6 | da | noćni ritam; hook do 2 s |
| **6** | K-4, F-9, N-6 (istraživanje), R-5 (uslovno), P-2, P-4 | da | struktura; pravila usklađena |

Talas 1 može da počne odmah i ne zavisi ni od čega. Talasi 1 i 1b mogu da teku
uporedo sa talasom 2.

### 8.2 Zavisnosti

- F-1 → F-2 → F-3; F-2 → F-6 (da cena izmene kataloga bude oko 1 s); F-2 → F-8 → F-9.
- M-2 → sve stavke čiji se dobitak meri po fazi (R-1, R-3, S-1).
- L-1 → R-1 (skraćena compile proba traži brz prekid na dijalog) i → D-2.
- M-1 → M-7 → D-2 → D-5, D-7; S-4 zatvara slabljenje iz D-2.
- D-6 → R-3 → R-4; R-3 → R-5.
- M-3 → D-7, N-1, N-5, N-6. L-2 → N-1.
- K-2 → K-4; S-4 → K-4.
- F-7 → P-2.

### 8.3 Kapije odluke

- **G0, posle M-7.** Tabela iz §2.2 je izmerena. Ako je režija znatno veća ili
  manja od 18 s, ili trajanje pojedinačnog testa znatno veće od 4 s, menja se
  odnos faza 3 i 4.
- **G1, posle D-2.** Koliko se verdikt u izboru razlikuje od verdikta cele suite
  i koliko testova ne prolazi samo. Ako je razlika nula nad uzorkom, izbor postaje
  podrazumevan u rezu. Ako nije, prvo S-4.
- **G2, posle R-3.** Udeo starta i gašenja Excela u prolazu odlučuje o R-5.
- **G3, pre K-4.** Miran trenutak (malo otvorenih grana nad test modulima) i
  rezultat S-4.

### 8.4 Prvih deset radnih dana

1. F-1, F-2, F-3 u jednom PR-u nad `tools/` (prototip i diferencijalni testovi
   su u `analize/…/prototip-fast/`).
2. L-3, L-4, L-5 kao mali PR; D-1 kao mali PR.
3. M-1: prenos izbora na tekući `main`, kapije na Linux-u, pa predaja na Windows.
4. PR nad ugovorom: M-2, M-3, L-1, L-2.
5. Windows prolaz (1–1,5 h): ručni Compile, FULL (vraća marker), protokol iz
   Priloga A.
6. G0: brojevi u §2.2, potvrda ili izmena redosleda.
7. Do desetog dana: D-2 prijem, D-3, D-4, S-1, S-2, F-4, i process PR za F-6.

---

## 9) Odluke koje traže vlasnika

| # | Odluka | Preporuka |
|---|---|---|
| **O-1** | Obim dokaza pred release. **V1:** grupno nad RC-om, uz uslov da je svaka sabotaža dokazana u poslednjih N dana. **V2:** pojedinačno u izboru nad zamrznutim commit-om, uz nedeljni grupni prolaz cele suite i reviziju uzorkom. | V2, kad G1 prođe. Do tada V1, jer je današnje „pojedinačno, cela suita" 44–82 h `[I]` i nije izvršeno od 28.08. |
| **O-2** | Može li Windows mašina noću da ostane upaljena i prijavljena, sa zasebnim klonom | Da; bez toga nema N-1. Uslove licence za zakazano pokretanje Excela treba proveriti `[?]`. |
| **O-3** | Da li nalaz van izmenjenog fajla u hook-u (F-9) blokira sledeći Edit | Upozorenje, ne blokada: višekoračni refaktor (prvo potpis, pa pozivaoci) bi inače stajao na međukoraku. |
| **O-4** | „Crveni prolaz nad starim kodom" obavezan samo za bug (`P5`) | Da. Za rez koji uvodi ponašanje crveni smer daje sabotaža. Ovo sužava praksu (ne pisano pravilo) i štedi 2–3 podizanja po rundi. |
| **O-5** | Miran trenutak za K-1, K-2, K-4 | K-2 pre kraja decembra; K-1 odmah posle M-1. |
| **O-6** | Gde žive knjiga dokaza i istorija prolaza | Lokalno (gitignored), a objava po otisku izvora kroz N-2. Self-hosted runner se ne preporučuje dok je repo javan. |
| **O-7** | Da li test moduli (35,8% VBA koda) namerno idu klijentu kroz self-update | Van obima plana; zabeleženo da se odluči posebno. |
| **O-8** | Jedan završni prijem po PR-u, na glavi koja se spaja (`P8`) | Da; rane kapije ostaju dozvoljene za rizične rezove. |
| **O-9** | Arhiviranje `STANJE_REFAKTORA.md` | Da, posle N-2, kad zapis merenja pravi alat. |
| **O-10** | Dva fajla van `ASSETS` i `CACHE_NAME` u PWA (nalaz 13) | Izmena ponašanja PWA, ne kapija; odlučuje se posebno. |

---

## 10) Rizici

| Rizik | Verovatnoća | Ublažavanje |
|---|---|---|
| VBA strana izbora ne radi kako je zamišljeno | srednja | M-1 prijem pre svega ostalog u fazi 3; Python strana je odvojiva |
| Mnogo testova zavisi od redosleda, pa izbor donosi manje | srednja | `--izolacija` u M-7 daje broj; S-4; izbor na nivou oblasti posle K-4 |
| Ubrzanje FAST sloja tiho promeni nalaze | niska | F-1 pre F-2; diferencijalni test ulazi u self-test; stara funkcija ostaje jedan rez kao referenca |
| Nov checker ili keš sam daje lažno zeleno | srednja | N2 i N5; svaki nov alat ima self-test u oba smera kroz `main()` |
| Izmene ugovora troše FULL prolaze | visoka | N6: dva PR-a nad ugovorom (talas 1b i R-1), ne pet |
| Excel nestabilan u dugoj sesiji („RPC failed" posle oko 10 min) | srednja | S-1 gasi tajmere; R-5 reciklira; prekid po suiti (L-1) |
| Rast nastavlja istom stopom | visoka | F-10 i M-6 čine ga vidljivim; faza 3 čini cenu linearnom |
| Refaktor žarišta se sudara sa otvorenim granama | visoka | radi skripta, u mirnom trenutku, jednim commit-om bez druge izmene |
| Noćni prolaz zavisi od jedne mašine; Office bez nadzora nije podržan scenario | srednja | zaseban klon, tvrdi prekid, ubijanje sopstvenog PID-a; knjiga je fail-closed |
| Statistička granica iz N-5 se pročita kao „sve je dokazano" | srednja | granica stoji u samom verdiktu; uzorak nikad ne zamenjuje prvi dokaz nove sabotaže |

---

## 11) Šta se ne radi

| Predlog | Zašto ne |
|---|---|
| Uzorkovanje umesto punog dokaza | 1264 od 1265 sabotaža dokazuje različitu tvrdnju; uzorak sme samo kao revizija (N-5) |
| Mutant schemata (sve mutacije u jednom build-u iza prekidača) | mehanički obmotljivo je samo 54% kataloga; meri se kod koji se ne isporučuje |
| Predviđanje ishoda mutacije umesto merenja | suprotno pravilu da se zeleno ne tvrdi bez merenja |
| Sužavanje hook-a na sabotaže izmenjenog fajla | slabi „ceo katalog posle svake izmene"; F-2 i F-8 daju isto bez slabljenja |
| Path-filter u CI-ju | mala ušteda (76 s, javan repo), a slabi pokrivenost po commit-u |
| Self-hosted runner na razvojnoj mašini | repo je javan; PR iz forka može da cilja takav runner |
| LibreOffice, twinBASIC ili emulator kao izvor verdikta | nema Excel objektnog modela ni formi |
| Jedna dugoveka Excel sesija bez recikliranja | zaostali `OnTime` i zastavice aplikacije prelaze u sledeći nalog |
| Dalje ulaganje u headless Compile | Excel nema API za to; ostaje ručna kapija |
| Deljena forma između testova (`T-15`) | slabi izolaciju UI stanja; tek ako merenje pokaže najmanje 10% suite |
| Veći `--grupe` kao trajno rešenje | više partnera po grupi, više zaklanjanja; samo kao privremena mera do D-2 |
| FULL u hook-u | već jednom je učinio sesiju neupotrebljivom (postmortem §7) |
| `merge=union` za katalog | može tiho spojiti dva unosa istog imena |

---

## 12) Nepoznanice i kako se zatvaraju

| Nepoznanica | Zatvara je |
|---|---|
| Raspodela oko 18 s režije po fazama; koji compile put se danas izvršava (2,5 ili 5,9 s čekanja) | Prilog A.1; zatim M-2 |
| Današnje trajanje `RunAllTests`, BFP i FULL | Prilog A.1 |
| Trajanje po testu; koliko traje BFP seed koji se u izboru pušta uvek | Prilog A.2 |
| Koliko testova ne prolazi samo | Prilog A.2 (`--izolacija`) |
| Koliko testova ostavlja stanje | M-5 |
| Razlika verdikta: izbor prema celoj suiti | D-2 |
| Da li `Application.Run` prima spisak testova duži od 255 znakova (grupa od 6 ide do 239, od 12 do 455) | Prilog A.2 |
| Udeo journal-a, loga i `OnTime` u trajanju suite | S-1 uz M-7 |
| Ponašanje `Remove` + `Import` nad kompajliranim projektom preko spoljnog COM-a | R-2 |
| Broj jezgara i RAM Windows mašine; ponašanje pod zaključanim ekranom | pitanje operateru; proba u N-1 |
| Da li granica od 65.534 linije važi u instaliranom Excelu | L-3 |
| Da li je `static` obavezan check i postoji li zaštita grane | podešavanja repoa |
| Zašto bazni prolaz dokaza uvek uključuje `RunAllTests` | pitanje vlasniku pre D-4; dok se ne zna, ne uklanja se |
| Stvarni broj Edit poziva po sesiji | jedan red u hook-u koji beleži vreme i trajanje (uz F-6) |

---

## Prilog A — Prvi merni prolaz na Windows-u

Sve komande su za Git Bash, iz korena repoa.

### A.1 Bez ijedne izmene koda, nad `main`-om

```bash
cd ~/Documents/GitHub/otkupapp-pwa
```

```bash
time python tools/run_vba.py --compile-only
```

```bash
time python tools/run_vba.py --suite TestLicense_All
```

```bash
time python tools/run_vba.py --suite RunAllTests
```

```bash
time python tools/run_vba.py --suite RunBusinessFlowProSuite --timeout 1500
```

Šta se čita: trajanje prvog poziva je režija bez šeme i suita; razlika prema
drugom je šema i jedna mala suita; treći i četvrti daju današnje trajanje dve
velike suite. Iz `tests/last_run.txt` posle prvog poziva ključevi `com_*` i
`keys_*` pokazuju koji se compile put izvršava.

### A.2 Posle M-1, nad granom sa izborom

```bash
python tools/run_vba.py --suite RunAllTests --suite RunBusinessFlowProSuite --timeout 1500
```

Blok `VREME` daje trajanje po testu i deset najsporijih u obe suite.

```bash
python tools/dokaz.py amb-pisac --izolacija
```

Nabraja testove koji sami ne prolaze.

```bash
python tools/dokaz.py izbor- bfp-izbor- --izbor
```

Četiri sabotaže samog izbora moraju biti crvene po imenu.

```bash
python tools/dokaz.py amb-pisac
```

```bash
python tools/dokaz.py amb-pisac --izbor
```

Isti skup u oba režima: verdikt mora biti isti po stavci, a odnos trajanja je
prvi izmeren dobitak faze 3.

---

## Prilog B — Podloga

Folder `docs/engineering/analize/2026-11-verifikacija/`:

| Fajl | Oblast |
|---|---|
| `01-fast-sloj.md` | `vba_check`, hook, prototipovi ubrzanja |
| `02-run-vba.md` | anatomija poziva, režija, paralelne instance |
| `03-dokaz-i-katalog.md` | model troška dokaza, semantika ubrzanja, održavanje kataloga |
| `04-suite-i-testovi.md` | suite, zavisnost od redosleda, sporedni efekti, rast modula |
| `05-ci-i-ne-vba.md` | CI koraci i trajanja, JS, migracija, noćni prolaz |
| `06-pravila-i-incidenti.md` | katalog tvrdnji T1–T26, nesaglasnosti dokument–kod, slojevita politika |
| `07-istorija-i-metrike.md` | krive rasta, udeo verifikacije u radu, otpad |
| `08-praksa-u-industriji.md` | šta je primenljivo iz mutacionog testiranja, a šta nije |
| `prototip-fast/` | zakrpe i diferencijalni testovi za F-2 i F-8 (rađeno nad kopijom alata) |
| `prenos-izbora/` | skripte kojima je nastao commit `9d6e780` i 23 dvosmerne provere |
| `skripte/` | merne skripte iz kojih su brojevi u §2 |

Reference `fajl:red` u analizama važe za `main @1fa4937`. Putanje `<scratch>` su
radni folder sesije u kojoj je analiza rađena i više ne postoje; skripte koje su
tamo bile, a vrede za dalji rad, prenete su u tri foldera iznad.
