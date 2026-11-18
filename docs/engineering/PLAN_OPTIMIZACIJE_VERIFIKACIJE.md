# Plan optimizacije verifikacije — provere, testovi, dokazi

> **Status:** predlog za odluku vlasnika. **Datum:** 15.11.2026, verzija 2.
> **Podloga:** osam nezavisnih analiza nad `origin/main @1fa4937` (08.11.2026),
> sačuvanih u `docs/engineering/analize/2026-11-verifikacija/`.
>
> Verzija 2 nosi ispravke posle protivničke provere (tri nezavisna pregleda:
> brojevi, činjenice u kodu, semantika i redosled; 46 nalaza, od toga 4 teška).
> Šta je provera promenila stoji u Prilogu C.
>
> **Stanje 17.11.2026:** prvi talasi su spojeni u `main` kroz #482 (merge
> `638bf299`): F-1–F-3, M-1, M-2 i L-1 (delom), S-1, prvi rez S-5, D-6 i R-4,
> uz usputnu ispravku AMB-INV-04. Šta je izmereno i šta sledi stoji u
> **Prilogu D**. Tekst ispod je plan kakav je bio pre sprovođenja; brojevi u
> njemu su polazni.

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
| Hook posle svake VBA izmene | 6–9 s na neopterećenoj mašini, 9–14 s pod opterećenjem `[M]`; kod malog fajla je 100% tog vremena posao koji ne zavisi od izmenjenog fajla |
| Pun dvosmerni dokaz kataloga | 1265 prolaza `[M]`; 44–100 h pojedinačno, 11–25 h u grupama po 6 `[I]`; poslednji put pušten 28.08.2026 nad 261 sabotažom `[D]` |
| Rast | za 30 dana sabotaže +76%, testovi +67%, produkcioni VBA +9% `[M]`; `RunAllTests` 0,36 → 0,64 s po testu za 27 dana `[D]` |
| Merenje | u 1560 commit-a oko 14 upotrebljivih brojeva o trajanju, svi u prozi `[M]` |

**Strategija u pet poteza.**

1. **Meri pre nego što menjaš.** Trajanje po fazi poziva, po suiti, po testu i po
   prolazu dokaza ide u mašinski zapis.
2. **FAST sloj:** algoritamski rezovi sa bajt-identičnim izlazom. Prototip postoji
   i izmeren je dva puta, nezavisno.
3. **Dokaz:** sa „cela suite po mutaciji" na „ciljani test po mutaciji" (cena
   postaje linearna), zatim manja režija poziva, zatim paralelne instance. Ovaj
   potez sužava jednu tvrdnju i zato nosi svoje ime u verdiktu (§5.3).
4. **Ritam:** pun dokaz se vezuje za kalendar (noćno), ne za release, a ishod
   Excel prolaza beleži alat, ne rečenica u dnevniku.
5. **Struktura:** tri fajla-žarišta (`sabotaza.py`, `modTest.bas`,
   `modBusinessFlowProTests.bas`) se dele pre nego što udare u VBA granice.

**Očekivani ishod.** Vrednosti u zagradi su merene pod opterećenjem paralelnih
sesija; odnos pre i posle je u oba merenja isti (2,3–2,9 puta). Na Windows-u
ništa od ovoga nije mereno.

| Mera | Danas | Posle talasa 1 (bez Excela) | Posle talasa 3 | Posle talasa 5 |
|---|---|---|---|---|
| Hook po Edit-u | 6–9 s (9–14 s) `[M]` | 2,3–3,5 s (3–5 s) `[M, prototip]` | isto | 0,1–2,0 s (0,2–3,8 s) `[M, prototip keša]` |
| Pun `vba_check` | 21,5–22,5 s (30–34 s) `[M]` | oko 10 s (13–14 s) `[M, prototip]` | isto | isto |
| CI job `static` | medijana 76 s `[M]` | 30–35 s `[I]` | isto | isto |
| Dokaz reza (primer: 4 BFP sabotaže) | 34 min `[D]` | isto | 3–4 min `[I]`; rez od 13–20 sabotaža 10–17 min `[I]` | isto |
| Pun dokaz kataloga | 44–100 h `[I]`, pojedinačno, cela suita | isto | oko 4–5 h `[I]` u izboru, verdikt „DOKAZANO (izbor)" | isto na jednoj mašini; oko 1–2 h sa 3–4 radnika `[I]` |
| Starost najstarijeg dokaza u katalogu | oko 11 nedelja `[D]` | isto | isto | do 7 dana `[I]`, uz rotaciju u izboru |
| Excel prolazi sa zabeleženim trajanjem | približno 0% | isto | 100% | 100% |

Red „Pun dokaz kataloga" je računat za oko 1290 sabotaža i bez testova koji
sami ne prolaze `[?]`. Posle kapije G0 (16.11.2026) računa se sa **izmerenom**
režijom od 6,5 s po pozivu i izmerenim prosekom po testu (0,66 s u
`RunAllTests`-u, 1,7 s u BFP-u, plus seed BFP-a do 6,7 s): oko 1.800 poziva,
3,8–5,2 h. Raste linearno sa katalogom: uz oko 18 sabotaža dnevno oko 6–7 h za
mesec dana i 9–12 h za tri meseca `[I]`.
Sve vrednosti „posle" koje se tiču Excela su model. Kapija odluke G0 (§8.3)
postoji da se posle prvog mernog prolaza redosled potvrdi ili promeni.

**Obim.** Talasi 1–5 nose oko 45–105 radnih dana jednog izvršioca `[I]`, ne
računajući čekanje na Windows prolaze. Talas 1 se u celosti dokazuje u CI-ju i
može da počne odmah.

---

## 2) Polazno stanje

### 2.1 FAST sloj i CI

Dve kolone merenja: analiza `01` je merila dok su druge analize radile na istoj
mašini, a provera plana je isto ponovila na neopterećenoj.

| Mera | Neopterećeno | Pod opterećenjem | Izvor |
|---|---|---|---|
| Hook nad malim ili srednjim `.bas` fajlom | 6,1–6,4 s | 9,2–10,3 s | `[M]` |
| Hook nad `modTest.bas` / BFP | 8,4–8,6 s | 12,0–13,7 s | `[M]` |
| Pun `vba_check.py` | 21,5–22,5 s | 30–34 s | `[M]` |
| `vba_check.py --self-test` | 12,3–12,5 s | 17,4–17,8 s | `[M]` |
| Deo hook-a nezavisan od izmenjenog fajla | — | 9,05 od 9,05 s (mali fajl); 9,4 od 12,6 s (`modTest`) | `[M]` |

| Mera | Vrednost | Izvor |
|---|---|---|
| Lokalni pun set kapija | 51,4 s redom; 22,8 s paralelno | `[M]` |
| CI job | medijana 76 s (25 poslednjih run-ova); 12 s u avgustu | `[M]` |
| Rast hook-a na istoj mašini | 0,32 s (16.08) → 6,19 s (08.11); oko 4,5 ms po sabotaži | `[M]` |
| Dnevno čekanje na hook | medijana 143 VBA hunk-a dnevno × 6,9–11,2 s = 16–27 min; p90 dan 55–90 min | `[I]` |

U self-testu 152 slučaja zajedno traju ispod 0,2 s; ostatak su dva puna prolaza
globalnih kapija nad pravim repoom.

### 2.2 Excel prolazi

| Mera | Vrednost | Izvor |
|---|---|---|
| Režija jednog `run_vba` poziva | **6,5 s bez compile probe** (16.11.2026, grana `claude/run-vba-ugovor`): priprema 0,3 · start Excela 0,4 · otvaranje 0,3 · uvoz 2,2 · šema 0,3 · zatvaranje i gašenje 3,0 (od toga 2,0 s fiksni `sleep`). Compile proba je dodavala 5,9 s čekanja (ide i COM i SendKeys put). Dokumentovanih „oko 18 s" je bila procena | `[M]` 16.11.2026; ranije `[D]` |
| `RunAllTests` | **230,9 s pri 354 testa** (16.11.2026, Windows); ranije 215,8 s pri 336 (05.11), 73,5 s pri 204 (09.10) | `[M]` 16.11; ranije `[D]` |
| `RunBusinessFlowProSuite` | **703,7 s pri 414 testova i 3487 provera**, pušten posle `RunAllTests` u istom pozivu (16.11.2026); 687,9 s u punom prolazu na grani #476; ranije 210 s (07.10) | `[M]` 16.11; ostalo `[D]` |
| FULL (12 suita) | ne staje u podrazumevanih 600 s; zapisi „oko 37 min" i „20–60 min" | `[D]` |
| Dokaz reza | 34 min za 4 BFP sabotaže uz `--grupe 6` (04.11; procena pre puštanja 20–25) | `[D]` |
| Jedna runda prijema reza | 7–20 podizanja Excela, 21–88 min | `[D]` |
| Pun katalog | 1265 prolaza pojedinačno, 215 uz `--grupe 6` | `[M]` kroz `--plan` |
| Trajanje punog kataloga, pojedinačno | 44 h uz dokumentovana trajanja suita (80 i 150 s); oko 100 h uz poslednje zapise (215,8 s i 350–390 s) | `[I]` |
| Trajanje punog kataloga, grupe po 6 | 11 h uz dokumentovana trajanja; 17 h uz poslednje zapise; oko 25 h sa 8% ponavljanja | `[I]` |
| Poslednji pun katalog | 28.08.2026, 260/261 | `[D]` |

**Prvo merenje po testu (16.11.2026, Windows, grana izbora `ec8e3bf`)** `[M]`.
Zbir koraka je 99,8% trajanja suite (702,4 od 703,7 s; 230,8 od 230,9 s), pa
izvan testova skoro nema skrivenog vremena.

- **Raspodela je koncentrisana.** Deset najsporijih nosi 22% `RunAllTests`
  (50,7 s) i 23% BFP-a (163,8 s).
- **Jedan BFP test traje 75,2 s** (11% suite): `Test_OTP_AutoKvarGrupeNeObaraOstale`.
  Zove `AutoCreateOtpremniceFromPWA_TX("")`, a prazan argument znači da batch
  obrađuje **svaki** nevezan otkup u svesci, ne samo dva koja test pravi. Cena
  mu raste sa svakim testom koji ostavi otkup, a u punom prolazu i sa
  ostacima `RunAllTests`-a. To je kandidat za uzrok sporijeg BFP-a posle
  `RunAllTests`-a (v. KI-096); nije izmereno.
- **Stvarna snimanja sveske** su među najsporijim testovima `RunAllTests`-a:
  tri testa snimanja po 5,6 s i `T_TxRollback_NepotpunZatvaraUpisISnimanje`
  11,4 s. To su testovi samog snimanja, pa ih tihi režim (S-1) ne skraćuje.

**S-1 izmeren (16.11.2026, Windows, grana `claude/s1-tihi-rezim` @`52e69af3`)**
`[M]`. Tihi režim CSV dnevnika i zakazivanja AutoSave-a u `RunAllTests`-u i BFP-u,
jedno merenje po stanju:

| Suita | Pre | Posle | Razlika |
|---|---|---|---|
| `RunAllTests` | 230,9 s, 354 testa (234,9 s na osnovnoj grani `claude/run-vba-ugovor`) | 222,2 s, 355 | −8,7 do −12,7 s (−4 do −5%) |
| BFP | 703,7 s, 414 testova, 3487 provera | 651,2 s, 415, 3488 | −52,5 s (−7,5%) |
| Obe | 934,6 s | 873,4 s | −61,2 s (−6,5%) |

Razlika kod `RunAllTests`-a je blizu šuma jednog merenja; kod BFP-a nije.
`Test_OTP_AutoKvarGrupeNeObaraOstale` 75,2 → 69,5 s. Režija 5,9 s (zatvaranje
2,5 s) je u šumu i ne pripisuje se S-1. Dvosmerni dokaz `s1-*`: 3/3, „DOKAZANO
(izbor)". Isti dobitak važi po prolazu dokaza koji vrti obe suite.

Analize se oko trajanja punog kataloga razilaze (40–100 h pojedinačno, 7–20 h
grupno), zavisno od toga koja trajanja suita uzimaju. Ništa od toga nije mereno;
plan nosi ceo raspon.

### 2.3 Rast i održavanje

| Mera | Vrednost | Izvor |
|---|---|---|
| Sabotaže | 718 → 1265 za 30 dana; oko +18 dnevno | `[M]` |
| Testovi u `RunAllTests` | 209 → 350 | `[M]` |
| Test moduli / produkcioni VBA | +40% / +9% (linije); odnos test/prod 0,17 → 0,51 | `[M]` |
| `modTest.bas` | 31.061 linija, +452 dnevno; granica 65.534 linije za oko 76 dana | `[M]`; `[I]` |
| BFP modul | 33.077 linija, +323 dnevno; granica za oko 100 dana | `[M]`; `[I]` |
| `TestName` (registar) | 22.178 bajtova; prag `PROCEDURA_VELIKA` 40.000 za oko 60 dana | `[M]`; `[I]` |
| Presidrenje kataloga | najmanje 64 u 29 dana (21 od 50 merge koraka) | `[M]` |
| Commit-i koji diraju samo `docs/` | 52,7%; `STANJE_REFAKTORA.md` 844 KB | `[M]` |

Rokovi su osetljivi na stopu rasta. U sedam dana posle analize `modTest.bas` je
rastao oko 57 linija dnevno, ne 452 (`origin/main @ed506dc`: 31.459 linija, 353
testa) `[M]`. Stavka M-6 zato prati rast nedeljno. Da li granica od 65.534 linije
važi u instaliranoj verziji Excela nije provereno `[?]`; proverava je L-3.

---

## 3) Dijagnoza: osam uzroka

1. **Kvadratna cena dokaza.** Jedna mutacija pušta celu suitu (350 ili 391 test)
   da bi posmatrala jednu tvrdnju. Broj prolaza raste sa katalogom, a trajanje
   prolaza sa brojem testova. Projekcija za šest meseci u grupama po 6: 129–256 h
   ako trajanje suite raste srazmerno katalogu, 40–75 h ako ne raste `[I]`.
2. **Fiksna režija po pozivu.** Svaki `run_vba` diže nov Excel, uvozi 193
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
većina se zatvara usput. Nalazi 1, 2, 5, 6, 8, 9 i 11 su u proveri plana ponovo
potvrđeni direktno u kodu.

| # | Nalaz | Izvor | Stavka |
|---|---|---|---|
| 1 | Hook se ne pokreće kad se menja `tools/sabotaza.py`, iako komentar u `vba_check.py` i dokumentacija tvrde da se pokreće | `[M]`: rc=0 za 0,06 s | F-6 |
| 2 | `sabotaza.py --self-test` (43 slučaja) ne vrti ni CI ni `vba_check` | `[M]` | L-4 |
| 3 | Hook ne vidi lom preko granice fajla: izmena potpisa `LogErr` daje rc=0, a pun prolaz 492 nalaza `ARNOST` u 78 fajlova | `[M]` | F-7, F-9 |
| 4 | Ne-ASCII bajt u fajlu koji čita katalog ruši `vba_check` traceback-om, bez imena fajla i reda | `[M]` | L-5 |
| 5 | Kad `run_vba` ubije Excel na 600 s, `dokaz.py` to u mutiranom prolazu prijavi kao „NE OBARA NISTA", a u baznom kao „nema označenog reda TESTS"; prekid nigde nije imenovan | `[M]` iz koda | L-1 |
| 6 | Zastavica `dialogs` u `SUITES` se nigde ne čita; „neočekivan dijalog = pad" iz dokumentacije nije implementirano | `[M]` iz koda | S-3 |
| 7 | KI-036: mutirani prolaz koji ostane zelen (sabotaža u toj suiti ne obori ništa) upiše zeleni marker nad sabotiranim izvorom | uslov `[M]` iz koda; incident `[D]`, mehanizam u `KNOWN_ISSUES` nije potvrđen | L-2 |
| 8 | `release.sh` ne zove nijednu kapiju i taguje pre ručnog Compile-a; bump verzije menja otisak izvora, pa poništava marker | `[M]` iz koda; poništavanje `[I]`, nije izvršeno; `[?]` da li se skripta danas koristi | N-3 |
| 9 | Osam podrazumevanih gate-suita nema rezultat-fajl, pa ime pale tvrdnje ne preživi COM granicu | `[M]` | K-3 |
| 10 | Kvar u `_strip_comment` ostavlja `vba_check --self-test` zelenim | `[M]` | F-1 |
| 11 | `popis_citalaca.py --check` postoji samo lokalno i crven je na `main`-u (33 preko praga 31) | `[M]` | F-5 |
| 12 | Sedam sabotaža sa priznatim statičkim nalazom nije u spisku koji `dokaz.py` čita, pa pun katalog verovatno ne može da vrati 1265/1265 | `[I]`, nije pušteno | M-7 (Prilog A.2), zatim upis u spisak ili ispravka tvrdnji |
| 13 | PWA: dva fajla koja `index.html` učitava nisu u `ASSETS` u `sw.js`; keširan fajl je menjan bez promene `CACHE_NAME` | `[M]` | F-11 |

---

## 5) Načela

### 5.1 Šta ostaje netaknuto

Katalog tvrdnji sa incidentima koji su ih doneli je u analizi
`06-pravila-i-incidenti.md` (T1–T26). Plan ne menja značenje nijednog postojećeg
verdikta. Bez izuzetka ostaju:

- pala je baš ciljana tvrdnja, po imenu;
- izvor se vraća i kad prolaz pukne, i to se proverava otiskom;
- nema rezultat-fajla = pad; BLIND nije dokaz; „nejasno" se prijavljuje kao nejasno;
- zeleno važi samo za baš ovaj izvor pod ovim test sistemom (marker);
- self-test checkera ide kroz produkcionu putanju, u oba smera, i vrti se u CI-ju;
- grupni dokaz tvrdi manje i tako se zove;
- bez Excela se izmena ponašanja prijavljuje kao neverifikovana.

### 5.2 Pravila ovog plana

- **N1. Nov režim nosi svoje ime.** Kao što danas postoji „DOKAZANO (grupno)",
  svaki režim koji meri drugačije dobija oznaku u verdiktu.
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
  zelene markere. Izmene se spajaju u jednu grupu po talasu, da se FULL plati
  jednom po talasu (§10).
- **N7. `.claude/` ide zasebnim process PR-om, jedan po jedan.**
- **N8. Mali rezovi.** Svaka stavka je zaseban, povratan rez sa sopstvenim
  kriterijumom prihvatanja.

### 5.3 Stavke koje sužavaju ono što se meri

Šest mesta u planu menja **šta** se meri ili **kada**. Nijedno nije čisto
ubrzanje i svako je odluka vlasnika.

| Stavka | Šta se sužava | Kako se zove i gde je granica | Odluka |
|---|---|---|---|
| D-2, izbor testa | tvrdnja se meri kad test ide sam, ne u kanonskom redosledu suite | verdikt „DOKAZANO (izbor)"; granice u odeljku ispod tabele D | O-1 |
| D-4, baza iz markera | „baza je zelena u ovom pozivu" postaje „bila je zelena ranije, nad istim otiscima" | opcija, ne podrazumevano; pred release je baza uvek sveža | posle G1 |
| O-1, obim pred release | pisano pravilo „pred release pojedinačno, cela suita" | do kraja talasa 5 grupno; zatim izbor uz reviziju | O-1 |
| S-6, dva Excela | BFP, Agrohemija i Golden više ne rade nad sveskom sa tragovima prethodnih suita | FULL pred release ostaje serijski, u jednom Excelu | O-11 |
| O-4, crveni prolaz nad starim kodom | praksa (ne pisano pravilo): „crveno nad stvarnim prethodnim kodom" postaje „crveno po sabotaži" | za bug ostaje obavezan | O-4 |
| N-5, revizija uzorkom | statistička tvrdnja, ne potpuna | granica piše u verdiktu; nikad ne zamenjuje prvi dokaz | — |

---

## 6) Ciljna slika: slojevi

| Sloj | Okidač | Šta se pušta | Budžet | Šta tvrdi |
|---|---|---|---|---|
| **Edit** | `PostToolUse` hook | provere izmenjenog fajla i globalne kapije (iz keša činjenica) | do 4 s, zatim do 2 s, na neopterećenoj mašini | mehanička ispravnost; katalog je živ |
| **Commit / push** | jedna komanda | pun `vba_check`, `who_writes` oba režima, `gen_schema --check`, popis suita | do 25 s | isti skup koji pušta CI; lom preko granice fajla |
| **Rez** | posle „CODE GO" | TARGETED suite i dokaz nad novim i dirnutim sabotažama koje bira alat | do 10 min | nove i dirnute tvrdnje su dokazane |
| **PR** | CI, i jedan završni prijem na glavi koja se spaja | `static`; `vba_gate --require-green --require-compile` | CI do 40 s | spojeni izvor je meren |
| **Noćno** | zakazan zadatak, zaseban klon | FULL nad `main`-om, rotacioni isečak dokaza, revizije | do 3 h | regresija se vidi za najviše 24 h; nijedan dokaz stariji od N dana |
| **Nedeljno** | vikend | grupni prolaz cele suite nad celim katalogom | danas 17–25 h `[I]`; raste sa katalogom | kontekst redosleda; kill-matrica |
| **Release** | ručno | bump → FULL, dokaz, ručni Compile nad tim izvorom → `vba_gate` → tag | do talasa 5: 17–25 h `[I]` plus FULL; jedna noć tek posle | isporučuje se baš ono što je mereno |

Dva sloja su nova (noćno i nedeljno). Ostali postoje u praksi, ali nisu zapisani
ili nemaju budžet.

Nedeljni sloj je jedini koji ništa u planu ne skraćuje osim paralelnih radnika
(R-4). Kad pređe oko 36 h, deli se po oblasti (posle K-4) ili rotira polovinu
kataloga nedeljno. To drugo je slabljenje i tako se zapisuje.

---

## 7) Radne stavke

### Faza 0 — Temelj: merenje i zatvaranje litica

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **M-1** | Prenos izbora testova na tekući `main`; verdikt dobija oznaku „(izbor)"; stavka koja u izboru ne obori svoju tvrdnju meri se celom suitom (`D4`, `RV-11`) | otključava fazu 3 | VBA strana nikad izvršena; menja ugovor (N6) | M | L + W | Linux: `vba_check`, self-test `run_vba` / `vba_gate` / `dokaz`, `--proveri-sidra`, 23 dvosmerne provere; self-test pada po imenu ako izbor vrati golo „DOKAZANO" i ako se ukloni povratak na celu suitu. Windows: ručni Compile; sabotaže `izbor-podniz`, `izbor-ne-trosi`, `bfp-izbor-podniz`, `bfp-izbor-ne-trosi` crvene po imenu |
| **M-2** | Trajanje po fazi i istorija prolaza u `run_vba` (`RV-01`, `IM-1`) | svaka dalja procena postaje merenje | menja ugovor (N6) | S | L | self-test: lažan izveštaj nosi sve faze, `REZIJA` jednaka zbiru; uklonjen dopis → crveno po imenu |
| **M-3** | Knjiga dokaza: red po prolazu sa svim palim tvrdnjama, trajanjem i režimom; `--plan` daje procenu iz istorije (`D1`, `IM-2`) | cena reza poznata pre puštanja; ukida ručni prepis | nizak | S | L | self-test sa lažnim runner-om: red za svaki ishod, nijedan dvaput |
| **M-4** | Fajl napretka po testu: red pre i posle svakog testa (`T-02`) | ubijen prolaz imenuje test koji je radio | format koji čita `dokaz.py` mora ostati isti; menja ugovor (N6) | S | W | prekid usred suite → izlaz imenuje taj test |
| **M-5** | Dijagnostički otisak tabela pre i posle svakog testa (`T-01`, `P5`) | tačan spisak testova koji ostavljaju stanje (danas statička granica 96–121 u `modTest`, 235 koraka u BFP) | oko 50 čitanja po testu | S | W | test koji namerno ostavi red je imenovan; kad se red vrati, nije |
| **M-6** | `tools/metrike.py`: rast i održavanje iz git-a, tabela u CI rezimeu (`IM-5`) | rast vidljiv nedeljno; rokovi iz §2.3 dobijaju stvarnu stopu | nizak | S | L | brojevi nad `HEAD`-om se slažu sa `sabotaza.py --lista` i `vba_gate --popis` |
| **L-1** | Prekid po suiti; `dokaz.py` prosleđuje izričit `--timeout`, čita red FATAL i izlazni kod; brz prekid na compile ili runtime dijalog; bazne suite u jednom pozivu (`RV-04`, `IM-3`, `D2c`, `D2d`) | uklanja liticu: BFP oko 370 s + režija prema granici od 600 s | prekratak budžet daje lažan TIMEOUT (fail-closed); menja ugovor (N6) | M | L + W | self-test: tajmer se naoružava po suiti; FATAL daje TIMEOUT i u mutiranom i u baznom prolazu. Windows: slomljen modul → prekid ispod 15 s sa tekstom dijaloga |
| **L-2** | KI-036: mutirani prolaz ne piše zeleni marker (`P7`, `IM-9`) | ponovljen prolaz suite (oko 4 min za BFP `[D]`) manje posle dokaza u kome je bar jedan mutirani prolaz ostao zelen | menja ugovor (N6) | S | L + W | self-test: mutirani poziv nosi prekidač, bazni ne. Windows: FULL → dokaz sa placebo sabotažom iz D-3 → `--require-green` ostaje zelen; isti niz bez prekidača daje „dokazan nad drugim izvorom" |
| **L-3** | `MODUL_VELIK` u `vba_check`: upozorenje na 45.000 linija u punom prolazu i CI rezimeu, greška na 60.000; jedna proba granice u Excelu (`T-05`) | upozorenje oko šest nedelja pre granice pri rastu iz §2.3 | `vba_check` danas nema neblokirajući nivo, uvodi se; lažna uzbuna ako je prag nizak | S | L + W | self-test u oba smera po uzoru na `PROCEDURA_VELIKA`; proba u Excelu ide uz M-7 |
| **L-4** | `sabotaza.py --self-test` u CI; self-test piše u privremeni folder, ne u `src-vba` (`F8a`) | dokaz za najskuplju proveru hook-a postaje automatski | nizak | S | L | namerni kvar `_dodela_tudjoj_proceduri` → CI korak crven |
| **L-5** | Ne-ASCII u fajlu kataloga daje nalaz `ASCII` sa fajlom i redom (`F6`) | tačna poruka za najčešću mehaničku grešku | vrlo nizak | S | L | nov slučaj kroz ceo CLI; pre izmene pada traceback-om |
| **M-7** | Prvi merni prolaz na Windows-u (Prilog A) | popunjava `[?]` iz §12 koji se tiču trajanja | — | 2–2,5 h mašinskog vremena `[I]` | W | tabela iz §2.2 nosi `[M]` |

Napomena uz M-1. Grana nosi jedan commit, a `main` je odmakao (158 commit-a na
dan pisanja). `dokaz.py`, `run_vba.py` i `vba_gate.py` se primenjuju bez sudara.
`sabotaza.py` ima jedan konfliktni blok, pa četiri sabotaže idu na kraj kataloga.
VBA strana se ne rešava ručno: nov test dobija sledeći slobodan broj, a
omotavanje BFP poziva se ponovo generiše skriptom nad tekućim telom suite.
Skripte su u `analize/2026-11-verifikacija/prenos-izbora/`. Ako VBA strana izbora
ne prođe prijem, vreme po testu se prenosi zasebno, bez izbora (`IM-4`), da G0 ne
čeka.

### Faza 1 — FAST sloj i CI

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **F-1** | Čuvari pre ubrzanja: self-test slučajevi za tokenizere; brojač obrađenih bajtova u proveri sidara (`F8b`, `F8c`) | ubrzanje ne može ući bez provere koja ume da pocrveni | nizak | S | L | kvar u `_strip_comment` (danas zelen) pada po imenu |
| **F-2** | Egzaktni rezovi: rani izlaz i inkrementalno parsiranje u proveri sidara; jedan prolaz tokenizera; čišćenje tela jednom po proceduri (`F1`–`F4`) | hook 6,1–6,4 → 2,3–2,4 s; `modTest` 8,4 → 3,5 s; pun prolaz 21,5–22,5 → oko 10 s; self-test 12,5 → 4,3 s `[M, prototip, neopterećeno]` | pet tokenizera je osnova gotovo svih provera | S | L | izlaz bajt-identičan nad čistim repoom u tri režima, nad dve „prljave" kopije (261 i 434 nalaza), nad 301.079 redova po funkciji i nad 1265 + 6000 zamena; tri namerna kvara padaju po imenu |
| **F-3** | Self-test ne ponavlja globalne kapije u dva „ceo CLI" slučaja; nov slučaj tvrdi da `main` zove sve tri kapije alata (`F5`, `CI-1`) | self-test posle F-2 → oko 2–3 s `[I]` | isključenje ne sme da procuri u produkcioni poziv | S | L | otkačen poziv kapije u `main` → nov slučaj pada |
| **F-4** | CI: `timeout-minutes`, `concurrency` za PR, Node 22; četiri paralelna job-a i zbirni `static`; `who_writes` u jednom prolazu (`CI-2`, `CI-3`, `CI-11`) | zid 76 s → 50–55 s samom podelom; 30–35 s zajedno sa F-3 `[I]`; pad jedne kapije ne sakriva ostale | ime obaveznog check-a mora ostati isto | S | L | namerno crven korak → ostali job-ovi daju verdikt, zbirni je crven |
| **F-5** | Kapije koje postoje samo lokalno ulaze u CI (`CI-4`) | zatvara klasu „kapija postoji, niko je ne pušta" | dve nemaju self-test i moraju ga dobiti pre toga; za `popis_citalaca` prvo odluka o pragu | M | L | po jedan namerno crven commit na test grani za svaki nov korak |
| **F-6** | Hook pokriva `tools/sabotaza.py`, `dokaz.py`, `vba_gate.py`, `run_vba.py` (`F7`, `P2`) — **process PR** | greška u katalogu se vidi u istom koraku | nizak | S | L | kroz pravi hook: pokvareno sidro u kopiji kataloga → rc=2 (danas rc=0) |
| **F-7** | Manifest kapija i lokalni pokretač; `--check` poredi manifest, `static.yml` i tabele u pravilima (`P3`, `CI-5`) | jedna komanda umesto 4–22; drift obara CI; sloj commit/push zatvara nalaz 3 | nov checker | M | L | korak samo u `static.yml`, korak samo u manifestu i izmenjena tabela padaju po imenu |
| **F-8** | Keš činjenica po fajlu, ključ sadržaj + otisak alata (`F9`) | hook 0,13 s van kataloga, 0,9–1,1 s u katalogu, oko 2,0 s za `modTest`/BFP `[M, prototip, neopterećeno]` | ustajao keš daje lažno zeleno; zato N5 | M | L | 16 scenarija izmene sa identičnim izlazom; pokvaren ključ → scenario „vraćena izmena" pada |
| **F-9** | Hook izveštava i o lomu pozivalaca u drugim fajlovima (`F10`) | lom vidljiv odmah u izlazu hook-a; uz O-3 = upozorenje tvrdnja hook-a ostaje ista kao danas | tabela zavisnosti mora biti potpuna | L | L | ekvivalencija sa hladnim punim prolazom posle niza nasumičnih izmena; kroz pravi hook: izmena potpisa → tekst upozorenja je u onome što agent dobije |
| **F-10** | Budžet trajanja kao kapija: brojač rada odmah, štoperica u CI prvo kao upozorenje (`P12`, `IM-6`) | cena ne može tiho da se vrati | lažna uzbuna na sporom runner-u | S | L | veštačko kašnjenje → prekoračenje po imenu |
| **F-11** | JS harness jasan bez `node_modules`; statička kapija nad PWA ljuskom (`CI-6`, `CI-7`) | prva kapija nad delom koji se isporučuje samim merge-om | deo je izmena ponašanja PWA (O-10) | S | L | fiksture u memoriji u oba smera |

**Odluka unutar faze.** Analiza pravila je predložila i sužavanje hook-a na
sabotaže izmenjenog fajla (`P1` u `06-…`). Ne preporučuje se: odlaže proveru
celog kataloga do commit-a, a F-2 i F-8 daju isto ubrzanje bez tog slabljenja.

### Faza 2 — Suite: skinuti trošak koji nijedna tvrdnja ne meri

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **S-1** | Tihi režim journal-a i loga u `RunAllTests`, BFP i ostalim suitama; gašenje `OnTime` tajmera na kraju suite; jedan imenovan test tvrdi journal i zakazivanje, sa svojom sabotažom (`T-03`). **Delom urađeno (16.11.2026):** journal i zakazivanje AutoSave-a u `RunAllTests` i BFP, test 355 i BFP provera, sabotaže `s1-*`; tihi režim loga (`modLogError`) i ostale suite nisu | hipoteza 5–25% trajanja `[?]`; **izmereno 6,5% obe suite, BFP 7,5%** `[M]` (§2.2); uklanja kandidata za „RPC failed" posle 10 min | nizak za journal (0 sabotaža usidreno u ta tri modula `[M]`); srednji za log: tihi režim loga danas ne postoji (nov prekidač u `modLogError`), a testovi koji tvrde AutoSave ili log moraju ga lokalno gasiti | S–M | W | isti skup imena i isti `TESTS`/`FAIL`; nov test pada na svoju sabotažu |
| **S-2** | BFP log u memoriji umesto 7 ćelija po tvrdnji; VBE prozor se sakriva posle compile probe (`T-04`) | 1–5 s po BFP prolazu `[?]` | bafer se gubi ako Excel padne, zato M-4 pre; menja ugovor (N6) | S | W | list sa logom identičan po redovima; `last_run_bfp.txt` identičan |
| **S-3** | „Neočekivan dijalog = pad" se sprovodi; `MsgBox` se preskače pod harnesom (`RV-10`) | pouzdanost; oko 11 dijaloga manje po FULL-u | dira 7 test modula; menja ugovor (N6) | S–M | L + W | dijalog u suiti sa `dialogs: False` → rc=2 po imenu; FULL bez reda `DIALOG` |
| **S-4a** | Svaki ciljani test sam: `dokaz.py --izolacija` nad celim katalogom, noću (`T-09`, `P5`) | izmeren spisak testova koji sami ne prolaze; uslov da izbor postane podrazumevan | 545 podizanja Excela, oko 3,3 h `[I]` | S | W | spisak imena; namerno zavisan test je u njemu |
| **S-4b** | Obrnut redosled i 3–5 permutacija sa zapisanim semenom; zavisnosti postaju deklaracije u registru (`T-09`) | zavisnost od redosleda postaje podatak; preduslov za K-4 | traži redosled kao podatak (K-2); može otkriti više zavisnosti nego što se očekuje | M | W | isti skup PASS/FAIL po imenu u svim permutacijama; namerno zavisan par je prijavljen |
| **S-5** | Trijaža najsporijih testova po izmerenom vremenu; vruća putanja pristupa podacima samo ako merenje pokaže, svaki rez uz `pre-flight` (`T-13`) — **uslovno, posle G0** **Prvi rez (17.11.2026):** `Test_OTP_AutoKvarGrupeNeObaraOstale` 69,5 s → ispod 6,4 s (test više ne obrađuje ~150 tuđih grupa iz sesije), dokaz 4/4; v. §2.2 | nepoznat dok M-7 ne da brojeve | visok: jezgro kroz koje ide svaki upis | M po rezu | W | vreme po testu pre i posle; nov test poništavanja keša crven svojom sabotažom |
| **S-6** | Podrazumevane suite u dva paralelna Excela (`T-10`) | FULL preko 600 s → oko 265–330 s `[I]` | sužava (§5.3, O-11); pravilo unije u markeru je mesto za lažno zeleno; menja ugovor (N6) | M | L + W | svaka podrazumevana suita zelena sama nad čistom kopijom; isti skup imena i verdikata; tri negativna slučaja markera u self-testu |
| **S-7** | Provera okruženja pre podizanja Excela: slobodan disk (`IM-7`) | sprečava „tri prolaza za nula nalaza" | nizak | S | L | lažiran disk ispod praga → odbijeno po imenu |
| **S-8** | Hermetički `RunOne`: snimak pre testa i vraćanje izmenjenih tabela (`T-12`) — **uslovno, posle M-5 i S-4a** | svaki test kreće od istog stanja, pa izolovan test i test u suiti mere isto | srednji do visok: nov kritičan mehanizam; u BFP-u se ne primenjuje bez posebne odluke | M | W | test koji ostavi red → sledeći vidi čistu tabelu; isti verdikti u pravom i obrnutom redosledu |

### Faza 3 — Dokaz: linearna cena

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **D-1** | Sitna režija: pamćenje `_suita_testa`; `--vrati` samo nad primenjenim članovima (`D2a`, `D2b`, `P8c`) | 27 s → ispod 1 s po pozivu; 1,0 s → 0,05 s po prolazu | nizak | S | L | isti plan (215 prolaza, ista podela); self-test oba alata zelen |
| **D-2** | Izbor testa kao podrazumevani put dokaza u rezu, posle diferencijalnog prijema (`D4`, `P2`) | prolaz 100–400 s → 20–26 s; rast postaje linearan `[I]` | testovi koji sami ne prolaze i stavke vraćene na celu suitu koštaju po 100–400 s (udeo `[?]`) | M | L + W | isti rez u oba režima, verdikt identičan po stavci, nad najmanje 50 sabotaža iz obe velike suite (0 razlika u 50 dopušta do 5,8% uz 95%); spisak iz S-4a postoji. **Alat (G1):** `python tools/dokaz.py --prijem --radnika 4` — jedan poziv nad jednim snimkom, izbor pa cela suite po stavci; imenilac su samo slučajne stavke koje je izbor presudio sam (43 ciljane i povratci na celu suitu ne ulaze), najmanje 50, od toga najmanje 15 u RAT i u BFP |
| **D-3** | Kontrolne sabotaže u svakom pozivu: placebo mutacija koja menja samo komentar ne sme biti priznata (`D12`) | dvosmerni dokaz samog dokaza u Excelu | jedan prolaz više | S | L + W | lažan izlaz u kome placebo „obara" tvrdnju → NIJE DOKAZANO |
| **D-4** | Opcija `--baza-iz-markera`, samo za suite koje se mere celom suitom (`D3`, `RV-12`) — **uslovno, posle G1** | 4–11 min po pozivu bez izbora `[I]`; uz izbor približno ništa | sužava (§5.3); marker ne hvata nedeterminizam između dva puštanja | S | L | važeći marker → nula baznih poziva cele suite; drugi otisak → baza se pušta; uz `--izbor` se baza izbora uvek pušta |
| **D-5** | Obim reza bira alat: `dokaz.py --od <ref>` (izmenjeni unosi, sabotaže čiji je fajl, procedura ili test diran) (`D8`, `P4`) **Urađeno lokalno (18.11.2026), bez Excela:** pet razloga (`KATALOG` uz `POZNATI_NALAZI_DOKAZ`, `PROCEDURA`, `TEST`, `FAJL`, `FIXTURE`); prihvat 11/11 nad stavkom 288 **nije ispunjen** — fale pozivaoci, otvorena odluka D-5b (Prilog D.3, tačka 4) | nema propuštenih i odloženih „dirnutih" stavki (2 od 11 odloženih je istrunulo) | pravilo o pomoćnom kodu može izabrati previše | M | L | sintetički diff: svaka vrsta pogotka bira svoju stavku; nad diff-om stavke 288 bira svih 11 |
| **D-6** | Mutacija van radnog stabla: `run_vba --src <kopija>`, `dokaz.py` nikad ne piše u `src-vba` (`D5`, `RV-05`) **Urađeno (17.11.2026):** snimak + kopija + zamrznut katalog, obnova iz snimka, straža; Windows: isti verdikti starim i novim putem (4/4), `src-vba` čist posle dokaza | radno stablo nije zaključano; nestaje klasa REVERT-FAIL; otvara paralelne instance | `sabotaza.py` primena traži ciljni direktorijum; menja ugovor (N6) | M | L + W | otisak `src-vba` isti pre, tokom i posle; ista sabotaža starim i novim putem daje isti verdikt |
| **D-7** | Nastavak prekinutog dokaza po konzervativnom ključu (ceo izvor, unos, sveska, alati, režim) (`D7`) | prekid posle 20 od 40 h ne košta 20 h | nizak | S | L | promena bilo kog sastojka ključa → stavka se ponovo meri |
| **D-8** | Statička predprovera novih i izmenjenih sabotaža: mutirani tekst kroz provere `vba_check` po fajlu (`P7c`) | sabotaža koja obara compile se hvata bez Excela | nizak | S | L | sabotaža sa duplom deklaracijom → statički nalaz |

**Šta izbor sužava, i dokle to mere pokrivaju.** Izbor dokazuje da je test
osetljiv na mutaciju kad se pusti sam, uz bazu nad istim izborom. Ne dokazuje da
je osetljiv u kanonskom redosledu cele suite: raniji test može ostaviti stanje
koje mutaciju zaklanja. FULL ide kanonskim redosledom, pa bi takvu regresiju
propustio, a knjiga bi govorila „DOKAZANO (izbor)". Mere u planu ovu razliku
**sužavaju, ne zatvaraju**:

- **S-4a** nalazi testove koji sami ne prolaze. To je obrnut smer (test traži
  tuđe stanje) i rešava ga povratak na celu suitu iz M-1. Zaklanjanje ne vidi,
  jer meri nemutiran izvor.
- **Nedeljni prolaz cele suite** (N-1) vidi zaklonjenu mutaciju kao člana koji u
  celoj suiti ne obori svoju tvrdnju; takav se meri sam. Ne vidi je kad tu
  tvrdnju obori partner iz iste grupe. To je poznata granica grupnog režima.
- **Revizija uzorkom** (N-5): 0 neslaganja u 100 dopušta do 2,95%, to jest do 37
  od 1265 sabotaža, i košta 6,5–11 h po reviziji `[I]`.
- **Uklanjanje uzroka** je jedino što razliku zatvara: test koji ne ostavlja
  stanje meri isto sam i u suiti. M-5 daje spisak testova koji stanje ostavljaju;
  S-8 ga uklanja, ako spisak pokaže da je to potrebno.

Posle podele po oblasti (K-4) izbor može da ide na nivou oblasti, pa redosled
unutar oblasti ostaje.

### Faza 4 — Režija poziva i paralelizam

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **R-1** | Čekanje na događaj umesto fiksnih `sleep`-ova; compile proba samo tamo gde nešto tvrdi (`RV-02`, `RV-03`) | compile deo urađen na grani `claude/run-vba-ugovor` (5,9 s); preostalo 2,0 s fiksnog čekanja pri gašenju `[M]` | slomljen izvor se u suite režimu vidi kasnije, zato L-1 pre; menja ugovor (N6) | S–M | L + W | `--compile-only` daje isti verdikt nad čistim i nad slomljenim modulom; 20 poziva bez zaostalog procesa |
| **R-2** | Proba: 50 puta zamena 6 modula u uvezenoj, kompajliranoj svesci (`RV-06`, preduslov) | odgovor na pitanje da li je R-3 izvodljiv | — | S | W | bez pada VBE-a i bez modula sa sufiksom `1` |
| **R-3** | Jedan uvoz po pozivu dokaza i zamena samo izmenjenih modula u svežoj kopiji sveske (`RV-06`, `D6a`) — **odloženo posle G0** | najviše oko 2 s po pozivu: ceo uvoz je 2,2 s `[M]` | ponašanje `Remove` + `Import` nad kompajliranim projektom `[?]`; menja ugovor (N6) | M | W | isti skup sabotaža hladno i delta putem daje isti verdikt; prazna delta → ista baza; pokvaren read-back pada po imenu modula |
| **R-4** | Paralelni Excel radnici, svaki nad svojom kopijom izvora; podela po izmerenom vremenu (`RV-09`, `P4`) **Povučeno napred posle G0:** zavisi samo od D-6 (R-3 je odložen); 6 fizičkih jezgara → podrazumevano 5 radnika | približno N puta kraće, realno 2–4× `[I]` | fokus i memorija; mašina zauzeta; menja ugovor (N6) | M | W | isti skup serijski i paralelno, tri ponavljanja bez razlike; dva radnika sa istim izlazom → alat odbija |
| **R-5** | Topla sesija sa „sendvičem": mutacija, test, vraćanje, isti test mora biti zelen; recikliranje na N mutacija (`RV-08`, `P3`) — **odloženo posle G0** | najviše oko 1 s po pozivu: start Excela 0,4 s, otvaranje 0,3 s `[M]` | visok: `OnTime`, zastavice aplikacije, break mode | L | W | ekvivalencija nad celim katalogom; namerno „curava" pseudo-sabotaža uhvaćena po imenu |

### Faza 5 — Ritam i evidencija

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **N-1** | Noćni zadatak u zasebnom klonu (Task Scheduler, prijavljena sesija): FULL nad `main`-om i rotacioni dokaz sa knjigom; nedeljno grupni prolaz cele suite (`CI-9`, `P6`, `D9`) | regresija vidljiva za 24 h; katalog svež | mašina mora biti upaljena i prijavljena | M–L | W | namerno pokvaren test → jutarnji ishod crven po imenu suite; ubijen proces → klon čist; tri noći bez intervencije |
| **N-2** | Objava dokaza po otisku izvora (commit status) i mašinski zapis po PR-u umesto prepisa u dnevnik (`CI-8`, `IM-8`) | većina od oko 150 ručnih docs commit-a mesečno | status postavlja operaterov token; lokalna putanja ne sme u javni repo | M | L + W | pao prolaz → nema objave; izmenjen izvor → „drugi izvor" |
| **N-3** | Release kapija: bump → FULL, dokaz, Compile nad tim izvorom → `vba_gate` → tek tada tag (`P10a`, `CI-10`) | tag se ne može postaviti nad nedokazanim izvorom | pogrešan redosled blokira hitno izdanje; kapija mora ići na put kojim se izdanje stvarno pravi `[?]` | S–M | L | suvo pokretanje u klonu bez markera staje pre taga |
| **N-4** | Politika obima dokaza pred release (odluka O-1) | pravilo postaje izvršivo | zavisi od izabrane varijante | odluka | — | — |
| **N-5** | Revizija uzorkom, isključivo kao dopuna: n sabotaža dokazanih u izboru meri se pojedinačno u punoj suiti; granica piše u verdiktu (`P6`) | 0 neslaganja u 100 → najviše 2,95% uz 95% | statistička tvrdnja se lako pročita kao univerzalna; 6,5–11 h po reviziji `[I]` | S | L + W | isto seme → isti uzorak; pokvaren zapis u uzorku je nađen |
| **N-6** | Registar dokaza preko commit-a (ključ po proceduri i zavisnostima): **samo istraživanje, u senci** (`P1`, `D10`, `P11`) | po PR-u bi se ponovo merilo u medijani 1,8% kataloga (prosek 3,2%, najviše 12,7%; 40 PR-ova) `[M]` | visok: propuštena zavisnost tiho potvrđuje zastarelo | L | L + W | kill-matrica iz nedeljnog prolaza: procedura svake sabotaže mora biti u zatvaranju njenog testa |

### Faza 6 — Struktura

| ID | Stavka | Dobitak | Rizik | Trud | Gde | Kriterijum prihvatanja |
|---|---|---|---|---|---|---|
| **K-1** | Katalog po modulu, bez escape-a; alat `--presidri` (`D11`, `P10a`) | sudari ostaju u jednom modulu; presidrenje dobija mehanički deo | jednokratna velika izmena fajla koji dira skoro svaki PR | M | L | rečnik iz starog fajla jednak rečniku iz novih, uključujući redosled; `--lista` identičan |
| **K-2** | Generisan registar testova, identitet je ime (`T-06`) | nov test je jedna procedura umesto četiri izmene; nestaje rok za `TestName` | nov checker; prva generacija menja oko 1.100 linija | M | L + W | istih 350 imena istim redom; `--check` pada na sva četiri današnja razlaza |
| **K-3** | Jedan test-kit i jedan format rezultata. **Prvi deo:** katalog priznaje više test modula, zajednički modul sa brojačima i tvrdnjama. **Drugi deo:** ostale suite ga usvajaju i dobijaju rezultat-fajl (`T-07`) | prvi deo je preduslov za K-4; drugi daje ime pale tvrdnje iz osam gate-suita | dira 10+ test modula; radi se suita po suita | M | W | isti broj provera po suiti; namerno oborena tvrdnja stiže po imenu |
| **K-4** | Podela `modTest` i BFP po oblasti, skriptom; suita po oblasti (`T-08`) | TARGETED oko 6× kraći `[I]`; uklanja rok od 65.534 linije | visok po obimu; sudar sa svakom otvorenom granom; menja ugovor (N6) | L | W | unija rezultata = ista imena i isti zbir tvrdnji; pun dokaz jednom posle podele |

**Rokovi koje nameće rast** `[I]`. Uz stopu iz §2.3 (+452 linije dnevno):
`TestName` prelazi 40.000 bajtova oko 07.01.2027; `modTest.bas` prelazi 45.000
linija oko 09.12.2026, 60.000 oko 11.01.2027, a 65.534 oko 23.01.2027. Pri toj
stopi K-2 mora početi do oko 10.12.2026, a K-4 ne stiže rok ako čeka talase 1–3.
Zato su K-2, prvi deo K-3 i K-4 **zaseban kolosek**, nezavisan od faze 3. Pri
stopi iz poslednjih sedam dana (+57 dnevno) rok je više od godinu dana daleko;
M-6 pokazuje koja stopa važi. Ako K-4 ne stigne, privremena mera je da novi
testovi idu u nov test modul iste suite, što traži samo prvi deo K-3.

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
| **1** | F-1, F-2, F-3, F-4, L-3 (statički deo), L-4, L-5, D-1, M-3, M-6, S-7, P-3 | ne | hook 2,3–3,5 s; CI 30–35 s; sve proverljivo u CI-ju |
| **1b** | jedna grupa nad ugovorom: M-2, L-1, L-2 | FULL jednom, zajedno sa talasom 2 | merenje i zatvorene litice |
| **2** | M-1 (prenos izbora), pa M-7 (merni prolaz, sa probom granice iz L-3) | da | brojevi umesto modela → **G0** |
| **3** | D-2 (prijem), D-3, S-4a, S-1, S-2, S-3, M-4, M-5, F-5, F-6; kolosek K: K-1, K-2, prvi deo K-3 | da | dokaz reza 3–4 min → **G1** |
| **4** | D-5, D-6, D-7, D-8, R-1 (gašenje), F-7, F-11; kolosek K: S-4b, K-4. R-2 i R-3 odloženi posle G0 | da | pun dokaz u izboru oko 4–5 h `[I]` → **G2** |
| **5** | N-1, N-2, N-3, N-5, R-4, S-6, F-8, F-10, drugi deo K-3, P-5 | da | noćni ritam; hook do 2 s |
| **6** | F-9, N-6 (istraživanje), P-2, P-4; uslovno: R-5, S-5, S-8, D-4 | da | pravila usklađena |

Talas 1 može da počne odmah i ne zavisi ni od čega. Talasi 1 i 1b mogu da teku
uporedo sa talasom 2. Kolosek K ima svoj rok (§7, faza 6) i ne čeka fazu 3.
N-4 nije u talasima jer je odluka (O-1), a P-1 je isto što i F-6.

### 8.2 Zavisnosti

- F-1 → F-2 → F-3; F-2 → F-6 (da cena izmene kataloga bude oko 1 s); F-2 → F-8 → F-9.
- M-2 → sve stavke čiji se dobitak meri po fazi (R-1, R-3, S-1).
- L-1 → R-1 (skraćena compile proba traži brz prekid na dijalog), → D-2, i →
  spajanje baznih suita u jedan poziv (215,8 + 350–390 + 18 = 584–624 s prema
  jednom tajmeru od 600 s).
- M-1 → M-7 → D-2 → D-5, D-7. S-4a → D-2 kao podrazumevan put.
- D-6 → R-4 (posle G0; ranije D-6 → R-3 → R-4, ali R-3 je odložen, a radnicima treba samo privatna kopija izvora); R-1 → R-3 (režija 6–8 s važi samo za obe); R-3 → R-5.
- M-3 → D-7, N-1, N-5, N-6. L-2 i D-3 → N-1.
- M-1 → K-1 (katalog se deli kad grana sa izborom uđe u `main`).
- K-2 → S-4b → K-4; prvi deo K-3 → K-4.
- M-5 i S-4a → S-8.
- N-1 i N-5 → varijanta V2 u O-1.
- F-7 → P-2.

### 8.3 Kapije odluke

- **G0, posle M-7.** Tabela iz §2.2 je izmerena. Ako je režija znatno veća ili
  manja od 18 s, ili trajanje pojedinačnog testa znatno veće od 4 s, menja se
  odnos faza 3 i 4. Ovde se odlučuje i o S-5.
  **Ishod (16.11.2026):** režija je 6,5 s, a prosečan test 0,66 s (`RunAllTests`)
  i 1,7 s (BFP). Skoro sve vreme nose same suite, pa faza 4 gubi vrednost: uvoz
  je 2,2 s, start Excela 0,4 s. R-2, R-3 i R-5 se odlažu (najviše oko 2 s, odnosno
  1 s po pozivu, uz srednji do visok rizik). Ostaju R-1 (preostalih 2,0 s
  fiksnog čekanja pri gašenju) i R-4 (paralelni radnici), a sledeća poluga su
  D-2 (izbor u rezu), S-1 (tihi režim) i najskuplji testovi iz merenja po testu
  (S-5): jedan BFP test od 75 s nosi 11% suite.
  **S-1 posle G0 (16.11.2026):** −6,5% obe suite (§2.2), na donjoj granici
  hipoteze. `Test_OTP_AutoKvarGrupeNeObaraOstale` (69,5 s, 10,7% BFP-a) je sam
  skuplji od cele uštede S-1, pa je on prvi kandidat za S-5.
  **S-5, prvi rez (17.11.2026):** taj test je ispao iz top-10 (< 6,4 s); BFP
  651,2 → 617,6 s, ali je isti prolaz bio sporiji i za testove koje rez ne
  dira (devet ostalih iz top-10 +6,4 %, od toga sedam PRE ovog testa;
  RunAllTests +1,8 %), pa suite-nivo uštede nije čisto izmeren.
  **D-6 urađen (17.11.2026)**, a R-4 povučen napred: zavisi samo od D-6.
- **G1, posle D-2 i S-4a.** Izbor postaje podrazumevan u rezu kad je razlika
  prema celoj suiti nula nad najmanje 50 sabotaža **i** kad postoji spisak testova
  koji sami ne prolaze. Za release (O-1, V2) traži se nula nad najmanje 100.
- **G2, posle R-3.** Udeo starta i gašenja Excela u prolazu odlučuje o R-5.
- **G3, pre K-4.** Miran trenutak (malo otvorenih grana nad test modulima) i
  rezultat S-4b.

### 8.4 Prve tri nedelje

1. F-1, F-2, F-3 u jednom PR-u nad `tools/` (prototip i diferencijalni testovi
   su u `analize/…/prototip-fast/`).
2. L-3 (statički deo), L-4, L-5 kao mali PR; D-1 i M-3 kao mali PR; F-4 kao PR
   nad `.github/`.
3. M-1: prenos izbora na tekući `main`, sa oznakom „(izbor)" i povratkom na celu
   suitu; kapije na Linux-u.
4. Grupa nad ugovorom: M-2, L-1, L-2. Spaja se sa M-1 pre Windows prolaza, da se
   FULL plati jednom.
5. Windows prolaz (oko 2–2,5 h `[I]`): ručni Compile, FULL (vraća marker),
   protokol iz Priloga A. Noću: `--izolacija` nad celim katalogom (S-4a).
6. G0: brojevi u §2.2, potvrda ili izmena redosleda.
7. Treća nedelja: prijem D-2, D-3, S-1, S-2, process PR za F-6; početak K-2.

---

## 9) Odluke koje traže vlasnika

| # | Odluka | Preporuka |
|---|---|---|
| **O-1** | Obim dokaza pred release. Pisano pravilo („pojedinačno, cela suita") je 44–100 h `[I]` i nije izvršeno od 28.08. **V1:** grupno nad RC-om. **V2:** pojedinačno u izboru nad zamrznutim commit-om, uz nedeljni grupni prolaz cele suite i reviziju uzorkom. | Do kraja talasa 5: V1 (215 prolaza, 17–25 h `[I]`), bez uslova „dokazano u poslednjih N dana", jer knjiga i rotacija još ne postoje. To sužava pisano pravilo i tako se zapisuje. V2 tek kad važe sva tri: G1 prošao nad najmanje 100 sabotaža iz obe velike suite; N-1 završio bar jedan pun nedeljni ciklus; N-5 dao prvi rezultat. |
| **O-2** | Može li Windows mašina noću da ostane upaljena i prijavljena, sa zasebnim klonom | Da; bez toga nema N-1. Uslove licence za zakazano pokretanje Excela treba proveriti `[?]`. |
| **O-3** | Da li nalaz van izmenjenog fajla u hook-u (F-9) blokira sledeći Edit | Upozorenje, ne blokada: višekoračni refaktor (prvo potpis, pa pozivaoci) bi inače stajao na međukoraku. Upozorenje mora stvarno stići agentu; lom i dalje zatvaraju pun prolaz pre commit-a (F-7) i CI. |
| **O-4** | „Crveni prolaz nad starim kodom" obavezan samo za bug (`P5`) | Da. Za rez koji uvodi ponašanje crveni smer daje sabotaža. Štedi 2–3 podizanja po rundi; sužava praksu (§5.3). |
| **O-5** | Miran trenutak za K-1, K-2, K-4 | K-1 čim M-1 uđe u `main`; K-2 do oko 10.12.2026 ako M-6 potvrdi stopu rasta iz §2.3. |
| **O-6** | Gde žive knjiga dokaza i istorija prolaza | Lokalno (gitignored), a objava po otisku izvora kroz N-2. Self-hosted runner se ne preporučuje dok je repo javan. |
| **O-7** | Da li test moduli (35,8% VBA koda) namerno idu klijentu kroz self-update | Van obima plana; zabeleženo da se odluči posebno. |
| **O-8** | Jedan završni prijem po PR-u, na glavi koja se spaja (`P8`) | Da; rane kapije ostaju dozvoljene za rizične rezove. Posle izmene se ponavlja ceo skup. |
| **O-9** | Arhiviranje `STANJE_REFAKTORA.md` | Da, posle N-2, kad zapis merenja pravi alat. |
| **O-10** | Dva fajla van `ASSETS` i `CACHE_NAME` u PWA (nalaz 13) | Izmena ponašanja PWA, ne kapija; odlučuje se posebno. |
| **O-11** | Podrazumevane suite u dva Excela (S-6): gubi se međusuitni kontekst | Da za dnevni rad; FULL pred release ostaje serijski. |

---

## 10) Rizici

| Rizik | Verovatnoća | Ublažavanje |
|---|---|---|
| VBA strana izbora ne radi kako je zamišljeno | srednja | M-1 prijem pre svega ostalog u fazi 3; Python strana je odvojiva; vreme po testu ima rezervni put (`IM-4`) |
| Mnogo testova zavisi od redosleda, pa izbor donosi manje | srednja | S-4a daje broj (M-7 daje samo uzorak); S-8; izbor na nivou oblasti posle K-4 |
| Izbor proglasi dokazanim ono što bi cela suita zaklonila | srednja | oznaka „(izbor)"; nedeljni prolaz cele suite; revizija uzorkom sa granicom; M-5 i S-8 |
| Ubrzanje FAST sloja tiho promeni nalaze | niska | F-1 pre F-2; diferencijalni test ulazi u self-test; stara funkcija ostaje jedan rez kao referenca |
| Nov checker ili keš sam daje lažno zeleno | srednja | N2 i N5; svaki nov alat ima self-test u oba smera kroz `main()` |
| Izmene ugovora troše FULL prolaze | visoka | N6: jedna grupa po talasu. Ima ih najmanje pet: talas 1b zajedno sa M-1; talas 3 (M-4, S-2, S-3); talas 4 (D-6, R-1, R-3, K-4); talas 5 (R-4, S-6); R-5 ako dođe na red |
| Excel nestabilan u dugoj sesiji („RPC failed" posle oko 10 min) | srednja | S-1 gasi tajmere; R-5 reciklira; prekid po suiti (L-1) |
| Rast nastavlja istom stopom | srednja | F-10 i M-6 čine ga vidljivim; faza 3 čini cenu linearnom; kolosek K ima svoj rok |
| Refaktor žarišta se sudara sa otvorenim granama | visoka | radi skripta, u mirnom trenutku, jednim commit-om bez druge izmene |
| Nedeljni prolaz cele suite prerasta vikend | visoka | prag od oko 36 h; R-4; podela po oblasti posle K-4 (§6) |
| Noćni prolaz zavisi od jedne mašine; Office bez nadzora nije podržan scenario | srednja | zaseban klon, tvrdi prekid, ubijanje sopstvenog PID-a; knjiga je fail-closed |
| Statistička granica iz N-5 se pročita kao „sve je dokazano" | srednja | granica stoji u samom verdiktu; uzorak nikad ne zamenjuje prvi dokaz nove sabotaže |

---

## 11) Šta se ne radi

| Predlog | Zašto ne |
|---|---|
| Uzorkovanje umesto punog dokaza | 1264 od 1265 sabotaža dokazuje različitu tvrdnju; uzorak sme samo kao revizija (N-5) |
| Mutant schemata (sve mutacije u jednom build-u iza prekidača) | mehanički obmotljivo je samo 54% kataloga; meri se kod koji se ne isporučuje |
| Predviđanje ishoda mutacije umesto merenja | suprotno pravilu da se zeleno ne tvrdi bez merenja |
| Sužavanje hook-a na sabotaže izmenjenog fajla (`06/P1`) | slabi „ceo katalog posle svake izmene"; F-2 i F-8 daju isto bez slabljenja |
| Marker zelenog po suiti (`06/P9`) | statički graf ne vidi `Application.Run` po imenu; propuštena zavisnost je lažno zeleno |
| Trajni keš uvezene sveske između sesija (`RV-07`) | ustajao keš znači da izvršen kod nije `src-vba`; R-3 daje veći deo dobiti bez toga |
| Drugi talas rezova bez keša (`01/F11`) | suvišan uz F-8 |
| Path-filter u CI-ju | mala ušteda (76 s, javan repo), a slabi pokrivenost po commit-u |
| Self-hosted runner na razvojnoj mašini | repo je javan; PR iz forka može da cilja takav runner |
| LibreOffice, twinBASIC ili emulator kao izvor verdikta | nema Excel objektnog modela ni formi |
| Jedna dugoveka Excel sesija bez recikliranja | zaostali `OnTime` i zastavice aplikacije prelaze u sledeći nalog |
| Dalje ulaganje u headless Compile | Excel nema API za to; ostaje ručna kapija |
| Deljena forma između testova (`T-15`) | slabi izolaciju UI stanja; tek ako merenje pokaže najmanje 10% suite |
| Veći `--grupe` kao trajno rešenje | više partnera po grupi, više zaklanjanja; samo kao privremena mera do D-2 |
| Paralelni delovi unutar jedne suite (`T-11`) i konsolidacija `modTest` sa BFP (`T-14`) | tek posle K-4; van obima ovog plana |
| FULL u hook-u | već jednom je učinio sesiju neupotrebljivom (postmortem §7) |
| `merge=union` za katalog | može tiho spojiti dva unosa istog imena |

---

## 12) Nepoznanice i kako se zatvaraju

| Nepoznanica | Zatvara je |
|---|---|
| Raspodela oko 18 s režije po fazama | Prilog A.1; zatim M-2. Compile put je izmeren: oba, 5,9 s (16.11.2026) |
| Trajanje FULL-a i BFP-a pustenog samog | Prilog A.1; `RunAllTests` i BFP posle njega izmereni 16.11.2026 (§2.2) |
| Koliko traje BFP seed koji se u izboru pušta uvek; da li je `Test_OTP_AutoKvarGrupeNeObaraOstale` uzrok sporijeg BFP-a posle `RunAllTests`-a | BFP pušten sam, pa poređenje vremena po koraku sa merenjem od 16.11.2026 |
| Koliko testova ne prolazi samo | S-4a: `--izolacija` nad celim katalogom, noću (Prilog A.3) |
| Koliko testova ostavlja stanje | M-5 |
| Razlika verdikta: izbor prema celoj suiti | Prilog A.2 (uzorak od 6); D-2 (najmanje 50) |
| Da li sedam sabotaža sa priznatim statičkim nalazom obara svoju tvrdnju u dokazu | Prilog A.2 |
| Da li `Application.Run` prima spisak testova duži od 255 znakova (grupa od 6 ide do 239, od 12 do 455) | D-2, uz `--grupe` |
| Udeo journal-a, loga i `OnTime` u trajanju suite | S-1 uz M-7 |
| Ponašanje `Remove` + `Import` nad kompajliranim projektom preko spoljnog COM-a | R-2 |
| Broj jezgara i RAM Windows mašine; ponašanje pod zaključanim ekranom | pitanje operateru; proba u N-1 |
| Da li granica od 65.534 linije važi u instaliranom Excelu | L-3, uz M-7 |
| Koja stopa rasta test modula važi (+452 ili +57 linija dnevno) | M-6 |
| Kojim putem se izdanje danas stvarno pravi (`release.sh` ili ručno) | pitanje vlasniku, pre N-3 |
| Da li je `static` obavezan check i postoji li zaštita grane | podešavanja repoa, pre F-4 |
| Zašto bazni prolaz dokaza uvek uključuje `RunAllTests` | pitanje vlasniku pre D-4; dok se ne zna, ne uklanja se |
| Stvarni broj Edit poziva po sesiji | jedan red u hook-u koji beleži vreme i trajanje (uz F-6) |

---

## Prilog A — Prvi merni prolaz na Windows-u

Sve komande su za Git Bash, iz korena repoa. Ukupno oko 2–2,5 h mašinskog
vremena `[I]`, zajedno sa ručnim Compile-om i FULL prolazom.

### A.1 Bez ijedne izmene koda, nad `main`-om (oko 11 min `[I]`)

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

### A.2 Posle M-1, nad granom sa izborom (oko 1,2 h `[I]`)

```bash
python tools/run_vba.py --suite RunAllTests --suite RunBusinessFlowProSuite --timeout 1500
```

Blok `VREME` daje zbir i deset najsporijih po suiti. Trajanje svakog testa je u
`tests/last_run.json`, koji sledeći poziv prepisuje, pa ga odmah treba sačuvati:

```bash
cp tests/last_run.json ~/vreme_m7.json
```

```bash
python tools/dokaz.py izbor- bfp-izbor- --izbor
```

Četiri sabotaže samog izbora moraju biti crvene po imenu.

```bash
python tools/dokaz.py izvestaji-cip zbirna-nacrt
```

```bash
python tools/dokaz.py izvestaji-cip zbirna-nacrt --izbor
```

Isti skup u oba režima: tri sabotaže nad tri testa iz `RunAllTests` i tri nad tri
testa iz BFP-a. Verdikt mora biti isti po stavci. Prvi poziv traje oko 43 min, a
drugi oko 5 min `[I]`; njihov odnos je prvi izmeren dobitak faze 3. Šest stavki
je uzorak, ne prijem: D-2 traži najmanje 50.

```bash
python tools/dokaz.py clanstvo-bulk- ispravka-bloka-bez-vadjenja otp-f8-storno-po-broju otp-prefill-bez-stavki malina-zbirna-nasledjuje-broj kanonska-zbirna-ne-sme-da-se-razveze --izbor
```

Sedam sabotaža sa priznatim statičkim nalazom (nalaz 12 iz §4). Ishod kaže da li
pun katalog danas uopšte može da vrati 1265 od 1265.

### A.3 Noću (oko 3,3 h `[I]`)

```bash
python tools/dokaz.py --izolacija
```

Nabraja ciljane testove koji sami ne prolaze: 545 testova iz dve velike suite,
po jedno podizanje Excela za svaki.

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
| `provera-plana/` | tri izveštaja protivničke provere verzije 1 ovog plana |
| `prototip-fast/` | zakrpe i diferencijalni testovi za F-2 i F-8 (rađeno nad kopijom alata) |
| `prenos-izbora/` | skripte kojima je nastao commit `9d6e780` i 23 dvosmerne provere |
| `skripte/` | merne skripte iz kojih su brojevi u §2 |

Reference `fajl:red` u analizama važe za `main @1fa4937`. Putanje `<scratch>` su
radni folder sesije u kojoj je analiza rađena i više ne postoje; skripte koje su
tamo bile, a vrede za dalji rad, prenete su u foldere iznad. Spoljni brojevi u
`08-praksa-u-industriji.md` (PIT, Stryker i slični) potiču iz sažetaka pretrage,
ne sa otvorenih stranica, i plan se na njih ne oslanja.

---

## Prilog C — Šta je promenila protivnička provera

Tri nezavisna pregleda su pokušala da obore verziju 1. Proverila su trinaest
činjenica o kodu na kojima plan stoji (jedanaest stoji u celosti, dve delimično)
i ponovila merenje prototipa (isti odnos, identičan izlaz). Našla su 46 stvari;
najvažnije ispravke:

| Šta je verzija 1 tvrdila | Šta stoji |
|---|---|
| „Plan ne menja nijednu tvrdnju T1–T26" | Šest mesta sužava ono što se meri ili kada; sada su u §5.3, svako kao odluka |
| Tri mere „zatvaraju" slabljenje izbora | Sužavaju ga, ne zatvaraju; zatvara ga samo uklanjanje ostavljenog stanja (M-5, S-8) |
| Verdikt izbora već nosi oznaku „(izbor)" | Grana ispisuje golo „DOKAZANO"; oznaku uvodi M-1 |
| Stavka koja u izboru ne obori tvrdnju meri se celom suitom | Grana to radi samo kad baza izbora nije zelena; ostalo je nov zahtev u M-1 |
| O-1: V2 čim prođe G1 | V2 zavisi od N-1 i N-5 iz talasa 5; do tada V1, bez uslova „N dana" |
| Pun katalog 44–82 h, grupno 11–20 h | 44–100 h i 11–25 h, uz trajanja suita koja plan sam navodi |
| Nedeljni prolaz cele suite staje u jednu noć | Danas 17–25 h i raste; prag i podela po oblasti u §6 |
| Hook 9–14 s | To je merenje pod opterećenjem; neopterećeno 6–9 s. Odnos pre i posle je isti |
| Prvi Windows prolaz 1–1,5 h; `amb-pisac` kao uzorak | `amb-pisac` je 19 sabotaža, sve u BFP-u, nad 4 testa (preko 2 h); zamenjen skupom od 6 iz obe suite |
| K-4 u talasu 6, uz rok sredinom januara | Zaseban kolosek K od talasa 3; rok zavisi od stope rasta koju prati M-6 |
| „Dva PR-a nad ugovorom" | Najmanje pet grupa, po jedna po talasu |
| „Prvih deset radnih dana" | Sadržaj je nosio 12–28 dana truda; sada „prve tri nedelje" |
| KI-036 nastaje „posle svakog dokaza" | Samo kad mutirani prolaz ostane zelen |
| D-4 štedi 2–9 min po pozivu | Uz izbor približno ništa; sada uslovna opcija |
| Šest stavki bez talasa; sedam predloga iz analiza bez reči | Raspoređene u §8.1; predlozi uneti u §7 ili §11 |

---

## Prilog D — Stanje posle #482 (17.11.2026)

Spojeno u `main` kroz #482 (merge `638bf299`, CI `static` zelen). Pre spajanja,
nad integracionom granom: FULL zelen (12 suita), ručni Compile bez poruke,
dokaz novih sabotaža 14/14 „DOKAZANO (izbor)". Oznake kao gore: `[M]` znači
izmereno, ovde na Windows-u operatera, osim gde piše Linux.

### D.1 Šta je sprovedeno

| Stavka | Šta | Izmereno |
|---|---|---|
| F-1–F-3 (#478) | Egzaktni rezovi u `vba_check`/`sabotaza.py`, izlaz bajt-identičan; čuvari i self-test bez ponavljanja kapija | hook 6 → 1,9 s; pun prolaz 22 → 9,3 s; self-test 12 → 2 s `[M, Linux]` |
| M-1 (#479) | `run_vba --test`, vreme po testu; `dokaz --izbor` sa oznakom „(izbor)" i povratkom na celu suite; `--izolacija` | mutacija pušta ciljani test, ne celu suite |
| M-2, L-1, delom (#480) | Faze poziva (`FAZE`), prekid po fazi, koban dijalog; compile proba samo uz `--compile-only` (KI-096) | režija poziva 12,4 → 5,4–6,5 s `[M]` |
| S-1 (#481) | `RunAllTests` i BFP u tihom režimu dnevnika i zakazivanja; test 355, BFP provera, sabotaže `s1-*` | obe suite −6,5% `[M]` |
| S-5, prvi rez | AutoKvar test meri samo svoje grupe | 69,5 → < 6,4 s; dokaz 4/4 `[M]` |
| D-6 | Dokaz mutira kopiju snimka (`--src`); `src-vba` se nikad ne piše; straža radnog stabla | isti verdikti starim i novim putem `[M]` |
| R-4 | `dokaz --radnika N`: zajednički bazen Excel radnika (poslovi po redu), `run_vba --izlaz`, brave, handle procesa; baza cele suite na zahtev (samo uz stavku koja ide na celu suite) i rana mutacija; jedna ocena faze 3 za N=1 i N>1; provera kopije pre svakog prolaza; #184 | v. D.2 |
| usput | AMB-INV-04: provera identiteta bila nedostižna iza AMB-INV-08; sabotaža `amb-pisac-red-bez-dokumenta` sada obara svoju tvrdnju | crveno po imenu pre, zeleno posle `[M]` |

R-4 je prošao pet krugova protivničkog review-a. Poslednja tri su svaki put
našla novu klasu greške, pa su ispravke išle strukturno, ne zakrpama:
- jedna ocena faze 3 za oba puta;
- ožičenje izdvojeno u funkcije koje self-test meri ponašanjem;
- poslovi po redu i provera kopije pre prolaza.

Self-test `dokaz.py` traje 0,65 s, a dokaz u oba smera nad njim obuhvata 26
mutanata. Bez Excela, 1384 stavke daju isto za N=1 i N=4. Dokumentovane razlike
između N=1 i N>1 postoje samo uz STOP, kad je verdikt ionako NIJE DOKAZANO
(`docs/EXCEL_TEST_HARNESS.md`, „Radnici").

### D.2 Brojevi posle

| Mera | Polazno (§2) | Sada |
|---|---|---|
| `RunAllTests` u FULL-u | 230,9 s | 222,9 s `[M]` |
| BFP u FULL-u, posle `RunAllTests` | 703,7 s | 581,1 s `[M]` (S-1 + S-5, jedno merenje) |
| Režija `run_vba` poziva | 12,4 s (sa compile probom) | 5,4–6,5 s `[M]` |
| Dokaz `amb-pisac` (19 sabotaža, jedna u fazi 3) | 21:22 (N=1) | 11:15 (N=5) `[M]`; posle AMB-INV-04 bez faze 3 |
| Dokaz `dop-spec` (20 sabotaža, bez faze 3) | po starom pravilu ~3,5 min po stavci | 3:09 (N=1), 1:32 (N=5) `[M]` |
| Dokaz 14 novih sabotaža lanca | — | 1:13 (N=4) `[M]` |

Skaliranje radnika na mašini operatera je oko 2× za N=5, ne 5×. Jedan Excel
koristi oko 1,4 logička jezgra, pa 5 Excel-a traži oko 7 od 6 jezgara.
`dop-spec` je meren dok je baza još kretala spekulativno i dok je druga sesija
puštala pun `vba_check`. Merenje sa N=4 i N=5 posle baze na zahtev je otvoreno
(D.3).

### D.3 Šta sledi

1. ~~**Merenje broja radnika**~~ — urađeno: `dop-spec` 3:09 (N=1), 1:15 (N=4),
   1:12 (N=5) `[M]`. Preporuka je **4**, jer peti Excel skoro ništa ne donosi.
2. ~~**S-4a**~~ — urađeno. `--izolacija` nad celim katalogom (574 ciljana
   testa, 4 radnika, 22:48 `[M]`) našla je 13 testova koji sami ne prolaze.
   Svaki sada sam priprema svoje stanje ili meri relativno. Posle ispravke
   svih 13 prolazi sam, `RunAllTests` 355/355, BFP 3491/3491, a dokaz nad
   njihovih 28 sabotaža daje 26 OK `[M]`. Preostale dve su zatečene i ne
   obaraju ništa nad fixture-om (`POZNATI_NALAZI_DOKAZ`, zaseban rez).
3. **G1:** prijem izbora kao podrazumevanog puta u rezu. Uslov je nula razlika
   prema celoj suiti nad bar 50 sabotaža. Komanda:
   `python tools/dokaz.py --prijem --radnika 4` (uzorak bez Excela:
   `--prijem --plan`). Uzorak čine 43 ciljane stavke (S-4a, mehanizam izbora,
   `POZNATI_NALAZI_DOKAZ`) i 60 slučajnih (RAT 36, BFP 24), determinističan,
   sa otiskom. **Imenilac** su samo slučajne stavke koje je izbor presudio sam.
   Povratak na celu suitu (`SKUPA`) ne ulazi, jer je tamo verdikt po
   konstrukciji jednak celoj suiti. Traži se najmanje 50 takvih, od toga
   najmanje 15 u RAT i u BFP. Za V2 (O-1) poziv je `--prijem 115`, uz najmanje
   100 presuđenih; alat i tada traži 50, pa se M čita iz reda `G1 UZORAK`. Pravila i klase: `docs/EXCEL_TEST_HARNESS.md`, „Prijem izbora (G1)".

   **Ishod (18.11.2026, Windows operatera, `--radnika 4`, nad `main`
   `aa6c03f1`, otisak uzorka `3ff4e52a91839e23`) — G1 PROŠAO, rc 0 `[M]`:**
   - presuđeno izborom **59** slučajnih (RAT 35, BFP 24), **0 razlika**; van
     ciljanih orijentaciono najviše **5,0%** uz 95% (pretpostavka
     reprezentativnog uzorka; uzorak je stratifikovan, pa to nije garancija za
     ceo katalog);
   - OPASNA 0, STROZA 0, NEJASNO 0; ISTO 98 (59 slučajnih + 39 ciljanih, među
     njima svih 28 sabotaža S-4a, bez ijedne S-4a regresije);
   - **SKUPA slučajni 0/60:** nijedna slučajna stavka nije morala na celu
     suitu, pa je udeo povrataka iz D-2 (`[?]`) u uzorku nula; 0/60 nije
     dokaz da ih nema. SKUPA ciljani 4/43: tačno četiri `POZNATI_NALAZI_DOKAZ`
     (negativne kontrole), u oba režima isti verdikt;
   - **NEUPOREDIVO 1:** `panel-kljuc-se-razisao-sa-sidebarom`. Zamena je
     nosila Python komentar (`# SABOTAZA`), pa mutacija od nastanka (03.09.)
     ne kompajlira — EXCEL UBIJEN, Compile error, na obe strane. Ispravljeno u
     katalogu; jedini takav slučaj u 1406. Ispravljena mutacija obara svoju
     tvrdnju: `--izbor` → OK, DOKAZANO (izbor) `[M]`.
   - Trajanje nije zabeleženo.

   Sledi zaseban process PR: izbor (`--izbor --radnika 4`) kao podrazumevan
   put dokaza u rezu, u `CLAUDE.md` §5 i `.claude/rules/testovi.md`.

   Nezavisni review (18.11.2026): #487 CONDITIONAL GO, #488 GO posle #487.
   Uslov G1-REV-01 (ciljani test koji sam ne prolazi obara prijem) je zatvoren
   u #487 kao pravilo presude; Windows prijem se zbog toga ne ponavlja, jer je
   izmereno `baza izbora crvena 0`. **Otvoreno pre V2 (O-1):** alat danas
   programski traži najmanje 50 presuđenih; za V2 prag od 100 mora postati
   izvršna kapija (rc), ne samo broj koji se čita iz reda.
4. **D-5** (`--od <ref>`) — urađeno lokalno, bez Excela. Komanda u rezu:
   `python tools/dokaz.py --od origin/main --izbor --radnika 4` (šta rez dira,
   bez Excela: `--od origin/main --plan`). Pravila, razlozi i odbijanja:
   `docs/EXCEL_TEST_HARNESS.md`, „Obim reza iz diff-a (D-5)". Izmereno
   (Linux, 18.11.2026):
   - self-test `dokaz.py` sa `_self_test_od` 0,79–0,87 s (šest merenja; pre
     D-5 0,59–0,62 s). Novo je ponašanje `main --od --plan` uz lažan git sloj
     i git sloj nad privremenim repoom (oko 0,1 s, oko 25 poziva git-a). Na
     Windows-u, gde je poziv git-a sporiji, vreme nije mereno;
   - dvosmerni dokaz (`d5_dvosmerno3.py`): 42 od 43 D-5 mutanta padaju po
     imenu (`D5: ...`). Zelen ostaje D14 („red između procedura je
     deklaracija"), koji je posle review-a ekvivalentan: oba slučaja idu istom
     granom („kod van procedura" je `FAJL`). Raniji mutanti nad istim fajlom:
     G1 30/30, R-4/D-6 26/26;
   - bez `--od` je izlaz bajt-isti: `--grupe 8 --plan` pre i posle
     (187 785 bajtova), takođe `--plan` i `relink --grupe 6 --plan`;
   - skup dirnutih procedura iz `difflib`-a je isti kao iz `git diff -U0` u
     417 od 417 fajl-diffova (poslednjih 120 commit-a nad `src-vba`, bez
     pravila o komentaru; sa njim 404, a 13 razlika su samo komentar u
     deklaracijama); bez pomeranja umetka 378 od 417;
   - `60a7957d^` → `60a7957d` (PRJ-ZBR-02): izabrano 30 od 1001; svih 8
     sabotaža sa sidrom u `ReassignPrijemnicaToZbirna_TX` su tu, među njima
     `cilj-zbirna-bez-provere-postojanja` (`PROCEDURA`). Prefiks porodice tog
     reza (`s6k8-`) hvata samo 2 od tih 8, i nju ne;
   - diff koji dira samo dokumente (`3c216acf`, `5655661e`): izbor prazan. U
     radnom stablu je to rc 2 („nema šta da meri"), izmereno sa
     `--od origin/main --plan` posle ove izmene.

   **Review 18.11.2026** (tri ugla: propušten izbor, git na Windows-u,
   integracija; skeptik po nalazu): 11 nalaza, nijedan oboren, svi ispravljeni,
   svaki sa mutantom koji pada po imenu:
   - izmena `POZNATI_NALAZI_DOKAZ` nije birala svoju sabotažu, a menja joj
     verdikt. `1559def1` (S-4a, upis dva priznanja) je davao 0, sada 2
     (`cilj-zbirna-*`, `KATALOG`);
   - rez koji dira samo fixture nije birao ništa. Sada je `FIXTURE` peti
     razlog i bira ceo katalog: `9ca82b61` 1158 od 1158, `7da9dad6` 613 od 613
     (oba samo `tools/make_fixture.py`). Od 1.10.2026 je 315 commit-a nad
     `src-vba` i 27 nad fixture-om ili kanonom (bez merge-a);
   - kod deklaracija test-modula (`FX_*`) sada bira sve testove tog modula;
     sidro sa više od jednog pogotka sada bira `FAJL`; kod između procedura
     (`#If VBA7` grana `modOtkupUI.FormHwnd`) sada bira `FAJL`;
   - git: samo plumbing (porcelain `git diff` je prepisivao `.git/index` i
     mogao da obori `git add` paralelne sesije u istom klonu); git koji nije
     dostupan i folder van repoa se više ne javljaju kao loša referenca; bez
     zajedničkog pretka poruka to kaže (u plitkom klonu uz savet
     `git fetch --unshallow`);
   - `main --od` se meri ponašanjem, ne iglom u izvoru: poređenje sa radnim
     stablom, učitan katalog kao nova strana, prosleđen `POZNATI_NALAZI_DOKAZ`
     i fixture, prazan izbor rc 2;
   - drugi prolaz `difflib`-a ima prag (blok 1000×1000): prepis celog BFP-a
     (razmak na kraju svakog reda) bez praga nije završio ni za 5 minuta, sa
     pragom 0,08–0,39 s. Posle praga je skup procedura i dalje isti kao iz
     `git diff -U0` u 417 od 417.

   Istorijski brojevi gore su ponovo izmereni posle ispravki i nisu se
   promenili (30, 12, 35, 0).

   **Windows (Git Bash, 18.11.2026):** self-test zelen (pravi git, privremeni
   repo); `--od origin/main --plan` nad zastarelim lokalnim `origin/main`
   (baza `aa6c03f1`) je izabrao 1 stavku (`panel-kljuc-se-razisao-sa-sidebarom`,
   `KATALOG`), isto što daje `od_istorija` nad `aa6c03f1..0b811e51` na Linux-u.
   Pouka: pre `--od` ide `git fetch origin`, inače se meri i već spojen tuđi rez
   (bezbedan smer, ali nepotreban posao). `--od` uz `--izbor` i Excel nije
   pušten.

   **Prihvat „nad diff-om stavke 288 bira svih 11" NIJE ispunjen.** Stavke 288,
   289 i 322 ne imenuju šest `s6k4` sabotaža. Uzete su one čije sidro živi u
   pozivaocu `IsPaletaPreradjena`, direktnom ili preko `PreradjenaPaletaRazlog`,
   i tačno ih je šest. Ostalih pet su tri `relink-*`, `cotenant-po-broju` i
   `writer-isti-broj-odbija`.

   | Diff | Izabrano | Od 11 | Fale |
   |---|---|---|---|
   | `5c3d72a8^` → `5c3d72a8` (review REV-04/05) | 12 od 1134 | 5 (`PROCEDURA`, `ReassignPaleteToPrijemnica_TX`) | svih 6 `s6k4` |
   | `d19c2dae` → `a1d98e76` (ceo rez 1b, prvi roditelj merge-a) | 35 od 1134 | 8 | 3 `s6k4-okvir-*` u `modStornoFlow.RunPrijemnicaCorrection` |

   Razlog koji fali je **pozivalac**, koji D-5 namerno ne bira.
   `SkiniPaleteZaStornoPrijemnice`, gde žive tri od njih, direktno zove
   `IsPaletaPreradjena` (dubina 1). `RunPrijemnicaCorrection`, gde žive ostale
   tri, zove je preko `PreradjenaPaletaRazlog` (dubina 2). Nad celim rezom su
   prve tri izabrane samo kao `FAJL` (nov `Private mPaletaPraznoTest` u
   deklaracijama), dakle slučajno.

   Proba proširenja, van alata: graf poziva po imenu procedure, bez stringova i
   komentara, pa je približan (`Application.Run` mu je nevidljiv).

   | Diff | Bez pozivalaca | Dubina 1 | Dubina 2 |
   |---|---|---|---|
   | `5c3d72a8` | 12, 5/11 | 21, 8/11 | 39, 11/11 |
   | rez 1b | 35 (proba 37), 8/11 | 53, 8/11 | 80, 11/11 |
   | `60a7957d` | 30 | 34 | 39 |

   Odluka za vlasnika: `POZIVALAC` do dubine 2 kao šesti razlog (D-5b) ili
   ostaje ručno dodavanje. Uz oko 5 s po stavci izborom (CLAUDE.md §5) 80
   stavki je oko 7 min serijski, ali to nije mereno.
5. **D-3** (placebo u svakom pozivu), **R-1** (2 s fiksnog čekanja pri
   gašenju), **F-8** (keš hook-a).
6. **Kolosek K:** podela kataloga i test modula, po roku iz §7 (faza 6) i stopi
   rasta koju meri M-6.
7. **Van plana, otvoreno:**
   - `On Error Resume Next` u aktivnom obrađivaču ne guta grešku (sonda
     17.11.2026). Pogođeno je 383 procedure, od toga 125 produkcionih.
   - Scenario B u `Test_Amb_ZaglavljeDokumentaPisac` tvrdi samo `errNum <> 0`,
     ista slabost kao scenario A pre AMB-INV-04.
