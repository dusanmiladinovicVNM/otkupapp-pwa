# Istorija i metrike -- koliko verifikacija stvarno kosta u praksi i kako raste (git istorija origin/main 03.04-08.11.2026, docs, poruke commit-a, GitHub Actions istorija, merenje FAST kapija nad istorijskim snimcima)

## SAZETAK
Istorija je uspesno produbljena (git fetch --deepen=2000: 4582 commit-a, 446 first-parent merge-ova, od 03.04.2026), pa su krive izvedene nad celom erom harness-a (katalog sabotaza postoji od 14.08.2026). Verifikacija raste 4-8 puta brze od produkcije: za 30 dana (09.10-08.11) sabotaze +76% (718->1265), RunOne +67% (209->350), test moduli +40% (56.6k->79.2k linija), a produkcioni VBA samo +9% (141.4k->154.3k); odnos test/prod je sa 0.17 (16.08) dosao na 0.51, a na svaku novu produkcionu liniju u poslednjih 30 dana dolazi 2.14 linije test koda + kataloga. U diff-u poslednjih 59 PR-ova verifikacija nosi 41% churn-a, docs 15%, produkcioni VBA 26% (bez migracionog alata: 50% / 18% / 32%); 52.7% svih commit-a dira SAMO docs (263 commit-a u STANJE_REFAKTORA.md, koji je sa 186 KB narastao na 844 KB), a medijana commit-a po PR-u je skocila sa 2 na 7 i lead time sa 2.3 h na 5.7 h. Trajanje Excel suita NE raste linearno sa brojem testova: RunAllTests 12.0 s/123 testa (27.08) -> 58.2 s/199 (25.09) -> 73.5 s/204 (09.10) -> 215.8 s/336 (05.11), tj. 0.10 -> 0.29 -> 0.36 -> 0.64 s/test; BFP ~150 s (02.10) -> 210 s (07.10) -> ~350-390 s (04.11, izvedeno). Brojevi u pravilima su ustajali 2-3x (CLAUDE.md: 3,5 min po sabotazi, 'sat i vise'; harness doc: BFP ~150 s): po poslednjim zapisima pun katalog sa --grupe 6 (215 prolaza) kosta oko 17 h Excela, a dokaz 4 sabotaze je 04.11 trajao 34 min. FAST sloj sam izmerio nad 8 istorijskih snimaka na ovoj masini: hook je sa 0.32 s (16.08) dosao na 6.19 s (08.11), priblizno linearno sa katalogom (~4.5 ms po sabotazi); CI medijana je sa 12 s na 73 s (1233 run-a, 1.9% failure). Otpad je realan i ponavlja se: 29% stavki u STANJE pominje bar jedan kvar samog alata (sidro/fixture/visenje/lazno crveno/placebo), 9 commit-a belezi visenje od 585 s, 139 postojecih unosa kataloga je presidreno ili prepravljeno u 30 dana, 42 razlicita crvena ishoda suita su zapisana. Najveca rupa: trajanja se uopste ne beleze masinski -- last_run.json nema ni vreme ni SHA i prepisuje se, dokaz.py ne meri nista; u 1560 commit-a postoji oko 14 upotrebljivih brojeva o trajanju, svi u prozi. Predlozi su zato pre svega merna infrastruktura (istorija prolaza, vreme po testu, git KPI u CI, budzeti) plus tri jeftine ustede koje iz istorije direktno slede (timeout iz istorije, KI-036, preflight okruzenja); nijedan ne slabi tvrdnje sistema.

## CINJENICE
- [M] Istorija je produbljena: origin/main ima 4582 commit-a (2116 first-parent, 446 first-parent merge-ova), 03.04-08.11.2026; repo je i dalje shallow (1 granica), ali cela era harness-a je pokrivena. Oba stabla su ostala cista (HEAD B=9d6e780b, A=1fa4937a).
    DOKAZ: cd /home/user/otkupapp-pwa && git fetch --deepen=2000 origin main; git log --oneline origin/main | wc -l -> 4582; git log --first-parent --merges --oneline origin/main | wc -l -> 446; git status --short -> prazno u oba stabla
- [M] Kriva rasta (poslednji first-parent commit u danu; dan: sabotaza / RunOne / BFP poziva / linije test modula / linije produkcionog VBA / test:prod): 16.08: 61/34/45/19.492/115.741/0.168; 30.08: 322/149/45/28.305/134.047/0.211; 13.09: 507/212/167/43.961/137.934/0.319; 27.09: 617/199/268/51.866/136.504/0.380; 11.10: 735/221/298/57.111/141.927/0.402; 24.10: 961/271/331/64.096/141.925/0.452; 01.11: 1091/308/344/70.565/145.173/0.486; 08.11: 1265/350/391/79.151/154.340/0.513. Pre 14.08 katalog ne postoji; pre 09.08 test/prod je bio stabilno 0.10-0.14.
    DOKAZ: python3 <scratch>/istorija-metrike/rast.py (git log --first-parent + git ls-tree -r + git cat-file --batch; broj sabotaza = ast broj kljuceva SABOTAZE; RunOne = regex ^\s*RunOne \d+; BFP = goli Test_ pozivi unutar RunBusinessFlowProSuite; test modul = *Test(s).bas / modTest*.bas osim modTestMode) -> rast.csv u istom direktorijumu (168 dana); kontrola: 1265 / 350 / 391 se slazu sa zadatim brojevima
- [M] Stope rasta 09.10->08.11 (30 dana): sabotaze +76.2% (+127.6/nedelji, udvostrucenje za ~37 dana), RunOne +67.5% (~40 dana), BFP poziva +31.2%, modTest.bas +75.6% (17.690->31.061), BFP modul +39.6%, test moduli ukupno +40.0% (+5.272 linije/nedelji), tools/*.py +20.6% (25.493->30.741), sabotaza.py +57.8% (8.745->13.801), produkcioni VBA +9.1% (+3.018 linija/nedelji), broj fajlova u src-vba +1% (192->194), vba_check.py +0.2%. U poslednjih 21 dan: +22.6 sabotaza/dan, +5.1 RunOne/dan, +4.3 BFP poziva/dan, +1002 test linije/dan naspram +654 produkcione.
    DOKAZ: python3 skripta nad <scratch>/istorija-metrike/rast.csv (redovi 2026-10-09 i 2026-11-08; kolone sabotaza, runone, bfp_poziva, modTest_lin, bfp_lin, test_mod_lin, prod_vba_lin, tools_py_lin, sabotaza_py_lin)
- [M] Marginalni odnos: 09.10->08.11 je dodato +12.935 produkcionih linija, +22.595 linija test modula i +5.056 linija kataloga -> 2.14 linije verifikacije (neto) po novoj produkcionoj liniji (16.08->08.11: 1.88). Gustina: 1.71 sabotaza po testu (RunOne+BFP poziv), 10.9 linija sabotaza.py po sabotazi, 89 linija modTest po testu, 85 linija BFP modula po pozivu.
    DOKAZ: isti proracun nad rast.csv: (79151-56556)/(154340-141405)=1.75; (13801-8745)/12935=0.39; 1265/(350+391)=1.71; 13801/1265=10.9
- [M] Diff po merge PR-u (git diff --numstat m^1 m), 59 PR-ova 09.10-08.11, 125.438 linija churn-a: test moduli 34.7%, katalog sabotaza 5.5%, ostali verif. alati+tests/ 0.9% => verifikacija 41.1%; docs+.claude 15.2%; produkcioni VBA 26.3%; migracioni alat tools/migracija 17.3% (dva PR-a, #455 i #472). Bez migracionog alata: verifikacija 49.7%, docs 18.4%, produkcija 31.9%. Odnos verifikacija : produkcioni VBA po periodima: 0.57 (09.08-08.09) -> 1.03 (09.09-08.10) -> 1.56 (09.10-08.11); sa docs: 0.84 -> 1.46 -> 2.14.
    DOKAZ: python3 <scratch>/istorija-metrike/pr.py 2026-08-09 -> pr.csv (280 PR-ova) + agregacija po datumu merge-a
- [M] Svih 280 PR-ova od 09.08: 1548 commit-a (prosek 5.5, medijana 3, max 90 = #408); churn 426.784 linije: verifikacija 36.7%, docs+pravila 15.6%, produkcija 47.6%. 76 PR-ova (27.1%) nema NIJEDNU produkcionu liniju: 47 samo docs/pravila, 29 samo verifikacija. Po periodima PR/30 dana: 113 -> 108 -> 59; commit/PR medijana 2 -> 2 -> 7; lead time (prvi commit -> merge) medijana 2.8 h -> 2.3 h -> 5.7 h, p90 26 h -> 15 h -> 78 h.
    DOKAZ: pr.csv iz pr.py (isti direktorijum) + python3 skripta: git log --no-merges --format=%aI m^1..m^2 za lead time
- [M] Commit-i po FAJLOVIMA koje diraju (bez merge-a, od 09.08: 1560): produkcija+verifikacija 39.7%, samo docs 30.9%, samo verifikacija 16.0%, samo produkcija 13.3%. U prozoru 09.10-08.11 (505): samo docs 52.7% (266), prod+verif 27.5%, samo verif 13.1%, samo prod 6.7%. Udeo 'samo docs' po periodima: 17.3% -> 24.4% -> 52.7%. 33% svih commit-a dira tools/sabotaza.py (520), 46% dira test module (713).
    DOKAZ: python3 <scratch>/istorija-metrike/commiti.py 2026-08-09 (git log --no-merges --numstat origin/main) -> commiti.csv
- [M] Commit-i po PORUCI (strogo: prefiks/scope test|tools|dokaz|harness|golden|gate|ci|fixture...): od 09.08 verifikacija 17.5% (273; od toga 221 dira samo test/alat), docs o verifikaciji 6.6%, ostali docs 24.4%, funkcionalni koji pominju test/dokaz 4.2%, funkcionalno 47.4%. U prozoru 09.10-08.11: 16.0% / 13.9% / 38.8% / 4.0% / 27.3%. Od 553 fix-commit-a, 119 (21.5%) popravlja SAMO test/alat/katalog. U 30 dana: 86 subject-a pominje review, 53 compile, 68 je zapis merenja (kapije / N/0 / izmeren), 29 zapis merge-a.
    DOKAZ: python3 klasifikator nad commiti.csv (regex prefiksa i kljucnih reci: dokaz|sabotaz|sidro|harness|vba_check|run_vba|fixture|self-test|test|golden|who_writes|vba_gate|dvosmern|RunAllTests|BFP); git log --no-merges --since=2026-10-09 --format=%s origin/main | grep -ciE 'review|...'. Heuristika po poruci je priblizna; klasifikacija po fajlovima je objektivna.
- [M] docs/STANJE_REFAKTORA.md ('Kratko i azurno -- cita se umesto celog plana') je glavni ponor rucnog rada: 8.954 B (16.09) -> 186.350 B (09.10) -> 844.024 B / 11.059 linija (08.11); u 30 dana ga dira 263 commit-a (52% svih), +26.790 / -18.131 linija. docs/RELEASE_NOTES.md je zamrznut od 09.09 (562 KB), pa otpad posle toga zivi samo u STANJE i porukama commit-a.
    DOKAZ: git cat-file -s <c>:docs/STANJE_REFAKTORA.md po datumima; git log --no-merges --since=2026-10-09 --until=2026-11-09 --numstat --format= origin/main -- docs/STANJE_REFAKTORA.md | awk (263 / 26790 / 18131); git log -1 --format=%cd origin/main -- docs/RELEASE_NOTES.md -> 2026-09-09; citat: main-wt/docs/STANJE_REFAKTORA.md:3-5
- [M] Odrzavanje kataloga sabotaza (poredjenje SABOTAZE recnika pre/posle svakog merge-a, ast bez izvrsavanja): u 30 dana +636 novih, -26 obrisanih, 139 POSTOJECIH unosa izmenjeno (70 sidro, 56 zamena, 67 tekst tvrdnje); 25 od 51 PR-a koji dira katalog menja ili brise zatecene unose. Od 09.08: +1395, -130, 350 izmenjenih (143 sidra); 67 od 160 PR-ova. Dakle na svaka 4 nova unosa odrzava se 1 stari.
    DOKAZ: python3 <scratch>/istorija-metrike/sidra.py 2026-08-09 -> sidra.csv
- [N] Vremenska serija RunAllTests iz zapisa (trajanje / broj testova / s po testu): 27.08: 12.0 s / 123 / 0.098; 25.09: 58.2 s / 199 / 0.292; 09.10: 73.5 s / 204 / 0.360; 04.11: 210 s / 314 (crven) / 0.669; 05.11: 215.8 s / 336 / 0.642. Trajanje NE raste linearno: testovi x2.7, vreme x18; marginalno 09.10->05.11 je 1.08 s po DODATOM testu naspram proseka 0.36 s. Skok 58.2->73.5 s pri +5 testova poklapa se sa prelaskom na kanonski fixture (09.10).
    DOKAZ: citati: docs/engineering/postmortems/2026-08-verifikacija.md:1214-1215 ('Vreme RunAllTests: 12.0s', 123 testa u :1203); poruka commit-a 9892b5cd ('RunAllTests 199/0 (SUITE OK, 58.2s)'); docs/STANJE_REFAKTORA.md:2115-2116 ('SUITE OK RunAllTests (73.5s)', '204 ukupno, 0 palo'); STANJE:10525 ('crveni RunAllTests 210 s'); docs/DOMEN/DOPLATA.md:306 ('RunAllTests 336/0 (215,8 s)'). Izvuceno: python3 .../istorija-metrike/trajanja.py 2026-08-01 -> trajanja_raw.csv. Nije izmereno ovde (nema Excela); uslovi masine izmedju tacaka nisu poznati.
- [N] Vremenska serija BFP: 02.10 ~150 s (276 poziva, 0.54 s/poziv); 07.10 210 s (2417 tvrdnji, ~295 poziva, 0.71 s/poziv); 20.10 ~4 min (308 poziva); 04.11 izmedju ~350 i ~390 s (368 poziva, 2953-3017 tvrdnji, ~1.0 s/poziv) -- izvedeno iz 'ceo poziv RunAllTests 210 s + BFP prelazi 600 s' i iz dokaza 4 grupe za 34 min. Marginalno 07.10->04.11: ~2-2.5 s po dodatom BFP pozivu naspram proseka 0.7 s. Isti obrazac kao RunAllTests: novi testovi su ~3x skuplji od starog proseka.
    DOKAZ: poruka commit-a f9b03a38 ('BFP ~150s, rezija jednog poziva ~18s'); STANJE stavka 76 ('RunBusinessFlowProSuite 2417/0 (210 s)'); docs/KNOWN_ISSUES.md:62 KI-036 ('~4 min za BFP'); STANJE:10525 i DOPLATA.md:309 ('ceo poziv sa BFP-om prelazi 600 s', '--timeout 1500'); STANJE:10280 ('dokaz 11:00-11:34 (oko 34 min; procena pre pustanja je bila 20-25)'). Broj poziva iz rast.csv.
- [N] Cena dvosmernog dokaza kroz vreme: 25.08 pun katalog 220 sabotaza ~2,5 h (~41 s po sabotazi); 23.09 ~3,5 min po sabotazi (RunAllTests+BFP), 4 pustanja u jednom rezu ~85 min; 02.10 cetiri kruga review-a #400 = 2h+ dokaza; 10.10 pun ponovni prolaz ~37 min; 04.11 dokaz 4 sabotaze = 34 min (8.5 min po sabotazi sa bazom), dokaz 13 sabotaza u 7 prolaza, ceo ciklus kapija jednog reza 14:50-16:18 = 88 min.
    DOKAZ: citati: commit 5c33d0ab / docs/UI_MIGRACIJA_KATALOG.md ('nad 220 sabotaza traje oko dva i po sata'); commit 1982fdcb; commit f9b03a38; STANJE:2484; STANJE:10280; STANJE:10514-10541
- [N] Model troska u pravilima je ustajao: CLAUDE.md:113-114 tvrdi 'oko 3,5 minuta po stavci, pa pun katalog traje sat i vise' (merenje od 23.09), EXCEL_TEST_HARNESS.md:732 'BFP ~150 s, rezija ~18 s' (02.10), poruka commit-a grane B (04.11) 'RunAllTests ~80 s, BFP ~150 s' -- a poslednji zapis je RunAllTests 215.8 s i BFP >= ~350 s. Procena sa tim brojevima: pun katalog sa --grupe 6 = 134 RunAllTests + 78 BFP + 3 banka prolaza = 134*(216+18) + 78*(370+18) + ~120 s = ~61.700 s = ~17 h; bez grupisanja ~100 h. To je racun, ne merenje.
    DOKAZ: main-wt/CLAUDE.md:113-114; main-wt/docs/EXCEL_TEST_HARNESS.md:732; git show 9d6e780 (poruka); podela 134/78/3 iz <scratch>/ci-analiza/plan6.out (grep '^grupa' | awk po suiti)
- [M] Serija broja testova iz zapisanih rezultata: RunAllTests zeleno 103 (21.08) -> 217 (16.09) -> 204 (09.10) -> 314 (04.11) -> 349 (06.11); BFP tvrdnji 400 (09.09) -> 1985 (25.09) -> 2417 (07.10) -> 3017 (04.11). Zapisano je 16 razlicitih CRVENIH ishoda RunAllTests i 26 crvenih BFP (npr. 2741/54, 2670/34, 274/175 nije brojan), svaki = bar jedan ponovljen prolaz. Zapisa 'dokaz N/M' u porukama commit-a: 61 (09.08-08.09, zbir M=1712), 38 (zbir 231), 62 (09.10-08.11, zbir 878).
    DOKAZ: trajanja_raw.csv (regex RunAllTests N/M, BFP N/M, dokaz N/M nad porukama i DODATIM linijama u docs/, datum = datum commit-a) + python3 dedup po (N, palo); ukupno 714 / 625 / 521 pominjanja
- [M] Trajanja se prakticno ne beleze: u 1560 commit-a i svim docs izmenama od 01.08 ima 5 pogodaka oblika 'Suite (N.Ns)' i oko 14 upotrebljivih brojeva o trajanju suita/dokaza (svi u prozi), naspram 1860 pominjanja rezultata N/M. tests/last_run.json (gitignored, prepisuje se svakim pozivom) nosi: workbook, import[], compile{ok,error,detail}, schema, suites[{name,gate,status,seconds,error}], suite_results{total,failed,detail}, dialogs, fatal, green -- BEZ vremena pocetka/kraja, git SHA, trajanja faza (start Excela, import, compile, schema), vremena po testu i oznake pozivaoca. tools/dokaz.py ne meri vreme uopste (nema import time).
    DOKAZ: trajanja_raw.csv: Counter(vrsta); main-wt/tools/run_vba.py:864 (report = {...}), :947-972 (jedini t0 je oko xl.Run suite), :1064-1068 (_write_report prepisuje tests/last_run.json); main-wt/.gitignore ('tests/last_run.json'); grep -n 'time\.' main-wt/tools/dokaz.py -> nema
- [M] FAST sloj izmeren nad istorijskim snimcima na OVOJ masini (git archive u scratch, min od 2 merenja, load 0.4-1.0) -- dan: sabotaza -> vba_check ceo / --hook modConfig.bas / sabotaza --proveri-sidra: 16.08: 61 -> 1.37 / 0.32 / n.p.; 30.08: 322 -> 10.84 / 1.94 / 0.48; 13.09: 507 -> 14.37 / 2.83 / 0.91; 27.09: 617 -> 14.53 / 3.15 / 1.23; 11.10: 735 -> 17.27 / 4.66 / 1.37; 24.10: 961 -> 18.40 / 5.01 / 1.82; 01.11: 1091 -> 19.80 / 5.34 / 2.17; 08.11: 1265 -> 21.58 / 6.19 / 2.64 s. Hook je x19 za 12 nedelja i raste priblizno linearno sa katalogom (~4.5 ms po sabotazi: (6.19-1.94)/(1265-322)); proveri-sidra blago nadlinearno (~N^1.25).
    DOKAZ: <scratch>/istorija-metrike/fast_istorija.sh 2026-08-16 ... 2026-11-08 (izlaz u transkriptu; snimci obrisani posle merenja). Apsolutne vrednosti su nize od ranije izmerenih 31.0 / 9-11.9 / 3.5 s (drugo opterecenje masine), trend je isti.
- [M] CI (static.yml) istorija: 1233 run-a od 14.08 (958 pull_request, 275 push), 24 failure (1.9%), 0 cancelled; medijana trajanja uspesnog run-a po nedelji: 12, 20, 26, 33, 37, 43, 43, 51, 54, 62, 62, 62, 73 s (x6 za 12 nedelja, ~+5 s nedeljno). U prozoru 09.10-08.11: 327 run-ova (10.5/dan), 6 failure, zbir 388 min; ukupno od pocetka 927 min.
    DOKAZ: GitHub Actions API (mcp github actions_list list_workflow_runs static.yml, 13 strana x 100) -> <scratch>/istorija-metrike/ci_runs.csv; trajanje = updated_at - run_started_at (ukljucuje cekanje runner-a)
- [M] Otpad u STANJE_REFAKTORA.md (323 numerisane stavke, 16.09-06.11) -- stavki koje pominju / ukupno pominjanja: presidravanje ili zastarelo sidro 36 / 49; ustajao ili regenerisan fixture 10 / 11; visenje, break mode, 585/600 s, RPC, ubijen Excel 17 / 36; COMPILE NEJASNO 33 / 35; lazno crveno, pad u premisi testa, ispravka testa 10 / 12; ponovljen prolaz ili pokusaj 55 / 67; NE OBARA SVOJ TEST / PALA DRUGA TVRDNJA / placebo 50 / 61. Bar jedan kvar samog alata (bez NEJASNO i 'ponovljen') pominje 94 stavke = 29%. RELEASE_NOTES.md (493 odeljka, 23.06-09.09): sidro 2, fixture 0, visenje 10/12, NEJASNO 35/37, lazno crveno 3, ponovljen 19/20, ne obara 4.
    DOKAZ: python3 <scratch>/istorija-metrike/otpad.py (regexi po kategoriji u samoj skripti; jedinica = numerisana stavka odnosno ##/### odeljak). Pominjanje nije isto sto i ponovljen prolaz -- gornja granica broja incidenata.
- [M] Otpad u porukama commit-a od 09.08: 9 commit-a eksplicitno belezi visenje od 585 s (19 ukupno pominje 585 s / 600 s / RPC / break mode); 54 pominje NEJASNO; 23 subject-a su samo zapis rucnog Compile-a; 53 pominje presidravanje ili zastarelo/mrtvo sidro; 109 pominje NE OBARA / PALA DRUGA TVRDNJA / placebo; 9 regeneraciju ili ustajalost fixture-a; 9 gresku u premisi testa.
    DOKAZ: git log --no-merges --since=2026-08-09 -E -i --grep='585 ?s|visio 585' origin/main (9); isti oblik sa --grep za ostale kategorije (komande u transkriptu)
- [N] Najskuplji primer 1 -- dokaz koji zastari pre merge-a: 'Cetiri kruga review-a na #400 su tako pojela 2h+ u `dokaz` prolazima koji su svi zastareli pre merge-a.'
    DOKAZ: git log -1 f9b03a38 (02.10.2026, perf(dokaz): --grupe)
- [N] Najskuplji primer 2 -- pun prolaz od 2,5 h u kome se kvar alata vidi tek na kraju: 'timeout ili prekid usred prolaza od 2.5h ostavljao namerno pokvaren radni izvor'; 'nad 220 sabotaza traje oko dva i po sata'; mrtvo sidro se 'hvata za sekundu umesto za dva i po sata'.
    DOKAZ: git log -1 9e70e6e1 (25.08.2026); main-wt/docs/RELEASE_NOTES.md:5433; main-wt/docs/engineering/postmortems/2026-08-verifikacija.md:394
- [N] Najskuplji primer 3 -- ceo katalog iz opreza: 'U tom rezu je pusten cetiri puta -- oko 85 minuta, od toga dva puta zbog moje greske.' Uzrok: 'Pravilo je bilo precutno ... bez reci o OBIMU -- pa se svaki rez iz opreza svodio na pun katalog.'
    DOKAZ: git log -1 1982fdcb (23.09.2026, docs(proces): obim dokaza)
- [N] Najskuplji primer 4 -- visenje do tvrdog prekida, svaki put ~10 min + ubijen Excel + poruka bez fajla i linije: 'To je ona ista zamka koja je dvaput naplatila 585 s i ubijen Excel'; 'posle 585 s i ubijenog Excela, uz poruku bez fajla i linije'; 04.11: 'prvi pokusaj je pao na ogranicenju od 600 s za CEO poziv run_vba (crveni RunAllTests 210 s + BFP); BFP je ponovljen sam, sa --timeout 900'; 29.10: 'RunGoldenSuite: Excel pao posle 6 s (RPC failed), zatim 600 s bez odgovora i proces ubijen'.
    DOKAZ: main-wt/docs/STANJE_REFAKTORA.md stavke 99, 71, 301 (:10525), 232; 9 commit-a sa '585 s' (a0053a5a, d029542a, 5a0ae65d, 63d6828c, 7cdc7941, acca3c1f, 1f5cc683, 4aee46b8, 61761134)
- [N] Najskuplji primer 5 -- crveno koje nije kvar koda: (a) okruzenje: 'Prvi RunAllTests posle ispravke je dao 274 / 3 sa "The disk is full" ... drugi 274 / 175, jer je ciscenje diska obrisalo otkup_test.xlsm i .sig ... pa treci prolaz: 274 / 0' (tri prolaza za nula nalaza); (b) ustajao fixture: 'To je vec pojelo pola sata trijaze -- cetiri crvena testa'; (c) staticka kapija propusti, a 'dokaz.py ih 25 minuta Excela kasnije javi kao NE OBARA SVOJ TEST'; (d) preneti brojevi: 'kroz osam narednih krugova je pisalo ... RunAllTests 199/0, BFP 1985/0. Ako projekat ne kompajlira, ti brojevi nisu mogli biti izmereni'.
    DOKAZ: main-wt/docs/STANJE_REFAKTORA.md:7496 (stavka 205); main-wt/docs/EXCEL_TEST_HARNESS.md:613; git log -1 66de2b3d (04.10); git log -1 9892b5cd (25.09)
- [N] dokaz.py zove run_vba sa podrazumevanim tvrdim prekidom od 600 s po POZIVU (ne po suiti), a BFP je po poslednjim zapisima ~350-390 s + ~18 s rezije i raste ~5.7 s/dan (210 s 07.10 -> ~370 s 04.11). Ako se trend nastavi, oko sredine decembra 2026 svaki BFP prolaz dokaza udara u prekid. FULL poziv (RunAllTests + BFP u istom pozivu) vec udara: dva zapisa, 04.11 i 05.11.
    DOKAZ: main-wt/tools/dokaz.py:155 i :349 (_pusti(sys.executable, 'tools/run_vba.py', '--suite', suite) bez --timeout); main-wt/tools/run_vba.py:705 (default=600.0), :886 (jedan threading.Timer za ceo poziv); STANJE:10525; DOPLATA.md:309. Datum litice je ekstrapolacija.
- [M] KI-036 je i dalje OPEN: dokaz.py kroz run_vba upisuje zeleni marker nad sabotiranim izvorom, pa posle svakog dokaza --require-green trazi ponovljen prolaz suite ('Cena je ponovljen prolaz suite-a posle dokaza (~4 min za BFP)'; danas ~6+ min). U kodu nema gasenja markera kad poziva dokaz.
    DOKAZ: main-wt/docs/KNOWN_ISSUES.md:62; grep -n 'marker\|green' main-wt/tools/dokaz.py -> nema; main-wt/tools/run_vba.py:1051-1058 (zabelezi_prolaz se zove uvek)
- [M] Donja granica broja pokretanja PostToolUse hook-a: u 30 dana 201 commit dira VBA fajlove, 765 parova (commit, fajl), 4.798 diff hunk-ova u .bas/.cls/.frm/.doccls (prethodni periodi: 6.374 i 7.574). Hook se pali samo za VBA ekstenzije.
    DOKAZ: git log --no-merges --since=2026-10-09 --until=2026-11-09 -p -U0 --format= origin/main -- 'src-vba/*.bas' 'src-vba/*.cls' 'src-vba/*.frm' 'src-vba/*.doccls' | grep -c '^@@' -> 4798; main-wt/.claude/hooks/vba-check.sh (case *.bas|*.cls|*.frm|*.doccls)

## TROSKOVI
[
 {
  "sta": "PostToolUse hook (vba_check --hook) po Edit/Write nad VBA fajlom",
  "koliko": "6.2 s na ovoj masini pri niskom opterecenju (ranije mereno 9.0-11.9 s); 16.08 je bio 0.32 s; raste ~0.45 s na svakih 100 novih sabotaza",
  "koliko_cesto": "najmanje 765 puta u 30 dana (parovi commit-fajl), verovatnije red velicine 4.800 (diff hunk-ovi) -> 8-16 h cekanja mesecno pri 6-12 s; procena ucestalosti, ne merenje",
  "izmereno": true
 },
 {
  "sta": "vba_check ceo repo (FAST, i u CI)",
  "koliko": "21.6 s ovde pri niskom opterecenju (ranije 31.0 s); 16.08: 1.4 s; 30.08: 10.8 s",
  "koliko_cesto": "svaki CI run + rucno pred kapije",
  "izmereno": true
 },
 {
  "sta": "CI static.yml (ceo workflow)",
  "koliko": "medijana 73 s u nedelji od 02.11 (12 s u avgustu); failure 1.9%",
  "koliko_cesto": "327 run-ova u 30 dana (10.5 dnevno) = 388 min masinskog vremena; 1233 run-a / 927 min od 14.08",
  "izmereno": true
 },
 {
  "sta": "RunAllTests (Excel)",
  "koliko": "215.8 s za 336 testova (05.11); 73.5 s za 204 (09.10); 58.2 s za 199 (25.09); 12.0 s za 123 (27.08) -- 0.64 s/test i raste",
  "koliko_cesto": "bazni prolaz svakog dokaza + svaka grupa RunAllTests sabotaza (134 od 215 prolaza punog kataloga) + svaki FULL; 59 commit-a u 30 dana nosi zapis rezultata",
  "izmereno": false
 },
 {
  "sta": "RunBusinessFlowProSuite (Excel)",
  "koliko": "~350-390 s (04.11, izvedeno); 210 s (07.10); ~150 s (02.10) -- ~1.0 s po pozivu testa",
  "koliko_cesto": "isto kao gore; 78 od 215 prolaza punog kataloga; 65 commit-a u 30 dana nosi zapis rezultata; 26 razlicitih crvenih ishoda zapisano",
  "izmereno": false
 },
 {
  "sta": "Dvosmerni dokaz u rezu",
  "koliko": "34 min za 4 sabotaze (04.11; procena pre pustanja 20-25 min); 13 sabotaza u 7 prolaza; ceo ciklus kapija jednog reza 88 min (14:50-16:18)",
  "koliko_cesto": "62 zapisa 'dokaz N/M' u porukama commit-a u 30 dana, zbir 878 sabotaza; uz efektivno grupisanje 1-2.2x to je reda 400-450 prolaza suite mesecno",
  "izmereno": false
 },
 {
  "sta": "Pun katalog dokaza pred release (--grupe 6, 215 prolaza)",
  "koliko": "~17 h Excela po racunu sa poslednjim zapisanim trajanjima (CLAUDE.md tvrdi 'sat i vise'); bez grupisanja ~100 h",
  "koliko_cesto": "pred svaki release; trosak raste kao proizvod (broj sabotaza x trajanje suite), oba +60-190% mesecno",
  "izmereno": false
 },
 {
  "sta": "Ukupno Excel vreme mesecno (model, ne merenje)",
  "koliko": "reda 40-60 h u prozoru 09.10-08.11: dokaz ~37 h (878 sabotaza / ~2 po prolazu x ~240 s + 62 bazna para), FULL/ciljani ~9 h (65 x ~520 s), crveni ponovljeni ~5 h (42 x ~7 min)",
  "koliko_cesto": "mesecno; sa trenutnim rastom isti obim posla sledeceg meseca kosta ~1.7x vise",
  "izmereno": false
 },
 {
  "sta": "Visenje do tvrdog prekida (compile greska na zahtev, MsgBox u test prolazu, RPC pad, prekoracenje 600 s)",
  "koliko": "585-600 s po dogadjaju + ubijen Excel + dijagnoza bez fajla i linije",
  "koliko_cesto": "9 commit-a sa '585 s' + 2 zapisa 600 s od 16.08; 17 stavki STANJE pominje visenje",
  "izmereno": false
 },
 {
  "sta": "Rucni Compile (automatski verdikt NEJASNO)",
  "koliko": "svaki rez ceka operatera; trajanje cekanja nigde nije zapisano",
  "koliko_cesto": "53 subject-a u 30 dana pominje compile, 23 su cist zapis 'rucni Compile PASS'; NEJASNO u 33 stavke STANJE i 54 poruke commit-a",
  "izmereno": true
 },
 {
  "sta": "Odrzavanje kataloga sabotaza (presidravanje i prepravka zatecenih unosa)",
  "koliko": "139 izmenjenih + 26 obrisanih unosa na 636 novih u 30 dana (0.26 po novom); 10.9 linija po unosu",
  "koliko_cesto": "25 od 51 PR-a koji dira katalog; 136 commit-a u 30 dana dira sabotaza.py",
  "izmereno": true
 },
 {
  "sta": "Rucno prepisivanje merenja i stanja u prozu (STANJE_REFAKTORA.md)",
  "koliko": "263 commit-a, +26.790 / -18.131 linija u 30 dana; fajl 844 KB (rast ~22 KB dnevno)",
  "koliko_cesto": "52% svih commit-a; ~150 commit-a mesecno su cisti zapisi merenja (68), merge-a (29) i compile-a (53)",
  "izmereno": true
 },
 {
  "sta": "Linije verifikacije po liniji produkcije",
  "koliko": "2.14 : 1 marginalno u 30 dana (1.75 test + 0.39 katalog); stanje 0.51 test/prod, 0.71 sa tools/*.py",
  "koliko_cesto": "trajno; odnos churn-a verifikacija : produkcioni VBA 0.57 -> 1.03 -> 1.56 po mesecima",
  "izmereno": true
 },
 {
  "sta": "Ponovljen BFP posle svakog dokaza zbog markera (KI-036)",
  "koliko": "~4 min (20.10), danas ~6+ min po dokazu",
  "koliko_cesto": "do 62 puta u 30 dana ako se --require-green pusta posle svakog dokaza (gornja granica ~6-7 h)",
  "izmereno": false
 }
]

## PREDLOZI
### IM-1 Istorija prolaza: run_vba dopisuje jedan red po pozivu [trud S] [slabi: NE]
tools/run_vba.py vec pravi report i meri sekunde po suiti. Dodati: ts_start/ts_end, git SHA, otisak izvora (vec se racuna u kontekst_dokaza), argumente (suite, no_import, compile_only, timeout), trajanje faza (start Excela, otvaranje, import, compile, schema, svaka suite), rc i klasu ishoda. _write_report pored last_run.json DOPISUJE red u tests/history/run_vba.jsonl (gitignored, po masini). Pozivalac se prenosi kroz promenljivu okruzenja (npr. DOKAZ_ID, DOKAZ_GRUPA). Samo Python, bez izmene VBA.
  DOBITAK: Nema direktne ustede sekundi; pretvara svaku buducu optimizaciju u merljivu. Danas postoji ~14 brojeva o trajanju u prozi za 3 meseca i 1560 commit-a; posle ovoga 100% prolaza ima trajanje po fazi. Trosak: +<50 ms po pozivu.
  RIZIK: Nizak. Greska u upisu istorije ne sme da obori run (isti try/except obrazac kao kod markera zelenog). Fajl raste ~0.5 KB po prolazu.
  PREDUSLOVI: Nema. Ide u svoj PR nad tools/ (ne dira .claude/).
  DOKAZATI: Prosiriti run_vba --self-test (radi na Linuxu): laznim report-om tvrditi da red postoji i nosi sva polja; dvosmerno: ukloniti dopis -> self-test pada po imenu -> vratiti. Na Windows-u jedan stvaran prolaz i pregled reda.

### IM-2 dokaz.py meri sebe i daje procenu pre pustanja [trud S] [slabi: NE]
dokaz.py danas ne meri nista. Dodati po prolazu: trajanje, suite, velicinu grupe, ishod (DOKAZANO, grupno, ponavlja se sam, TIMEOUT, NE OBARA); na kraju zbir (prolaza, sekundi, sekundi po sabotazi, efektivan faktor grupisanja, koliko je islo na ponavljanje). --plan ispisuje PROCENU iz istorije IM-1 (bazni prolazi + broj prolaza x medijana trajanja te suite + rezija). Sve se dopisuje u tests/history/dokaz.jsonl.
  DOBITAK: Operater bira obim dokaza sa stvarnom cenom (04.11: procena 20-25 min, stvarno 34). Daje jedini nedostajuci broj za plan: koliko prolaza i sekundi kosta rez. Otkriva gde grupisanje ne radi (u rezovima 1.0-2.2x naspram 5.9x u planu celog kataloga).
  RIZIK: Nizak; izmena alata koji je i sam checker, pa trazi svoj dvosmerni dokaz.
  PREDUSLOVI: IM-1 (za procenu); samo merenje moze i bez njega.
  DOKAZATI: dokaz.py --self-test vec ima lazni prolaz (_lazni_prolaz): tvrditi da zbir i red istorije postoje i da je procena = zbir medijana; sabotirati racun -> crveno po imenu.

### IM-3 Tvrdi prekid iz istorije i po suiti, umesto fiksnih 600 s po pozivu [trud S] [slabi: NE -- prekid nije tvrdnja o kodu; visenje se i dalje prijavljuje kao TIMEOUT, ne kao zeleno.]
run_vba ima jedan Timer od 600 s za CEO poziv, a dokaz.py ga zove bez --timeout. BFP je ~350-390 s i raste ~5.7 s dnevno; FULL poziv vec puca (04.11, 05.11). Predlog: prekid se racuna po suiti kao max(600, k x medijana poslednjih zelenih trajanja te suite iz istorije, k oko 2.5), sat se resetuje izmedju suita, a dokaz.py prosledjuje isti racun. Bez istorije vazi dosadasnjih 600 s.
  DOBITAK: Sprecava liticu: kad BFP + rezija predje 600 s, svaki BFP prolaz dokaza postaje TIMEOUT (78 od 215 prolaza punog kataloga, vecina prolaza u rezovima S6-S8), svaki po 10 min izgubljenog vremena + ponavljanje. Vec sada stedi 10 min po FULL pozivu koji bi inace pukao. Pravo visenje se i dalje hvata, ali relativno prema ocekivanom trajanju.
  RIZIK: Stvarno visenje se otkriva kasnije nego danas ako je k prevelik (ograniciti gornjom granicom, npr. 1500 s). Izmena dira ponasanje alata pod prekidom.
  PREDUSLOVI: IM-1 za istoriju; prelazno resenje bez nje je eksplicitan --timeout u dokaz.py.
  DOKAZATI: Self-test bez Excela: lazna istorija -> ocekivan prekid; lazni prolaz duzi od prekida -> TIMEOUT po imenu; sabotaza racuna -> crveno. Na Windows-u: FULL poziv koji danas puca na 600 s prolazi bez rucnog --timeout 1500.

### IM-4 Vreme po testu kao zaseban, mali rez (odvojeno od izbora testova sa grane B) [trud M] [slabi: NE -- samo meri; ne menja nijedan test ni redosled.]
Serija pokazuje da s/test raste (0.29 -> 0.36 -> 0.64 u RunAllTests; 0.54 -> ~1.0 s po BFP pozivu), a uzrok je nepoznat jer vreme po testu ne postoji na main-u. Grana B ga ima (last_run_vreme.txt), ali spojeno sa izborom testova, 149 commit-a iza i u sudaru sa modTest/BFP. Predlog: na main-u ponovo uraditi SAMO merenje: u modTest.RunOne Timer oko InvokeTest i dopis reda u last_run_vreme.txt (jedno mesto); u BFP-u isto na postojecoj granici testa. run_vba cita fajl u istoriju (IM-1) i ispisuje 10 najsporijih. Posle prvog prolaza odgovoriti: da li je trosak u malom broju teskih testova ili raste sa rednim brojem testa (nagomilano stanje ili velicina fixture-a).
  DOBITAK: Nepoznat dok se ne izmeri -- to je poenta. Ako je uzrok nagomilano stanje i s/test se vrati na nivo od 09.10 (0.36 s; BFP 0.71 s po pozivu): RunAllTests 216 -> ~125 s (-42%), BFP ~370 -> ~260 s (-30%), sto se mnozi kroz svaki prolaz dokaza (oko -35% Excel vremena, reda 15-20 h mesecno po modelu). Ako je uzrok nekoliko teskih testova, dobija se spisak kandidata.
  RIZIK: Srednji: VBA izmena u test modulima, mora da se izvrsi u Excelu; Timer ima rezoluciju ~15 ms i prelazi ponoc (koristiti razliku sa korekcijom). Upis fajla po testu dodaje malo I/O (pisati jednom na kraju iz niza).
  PREDUSLOVI: Windows + Excel za verifikaciju; ASCII izvor; bez novih WithEvents. Korisno uz IM-1.
  DOKAZATI: Test u modTest: broj redova u fajlu vremena = m_Total i zbir vremena <= trajanje suite; sabotaza (preskoci dopis) obara taj test po imenu. Zatim dva prolaza iste suite: vreme po testu stabilno u +-10%.

### IM-5 tools/metrike.py: KPI iz git-a, bez Excela, u CI [trud S] [slabi: NE]
Objediniti ono sto je ovde uradjeno skriptama rast.py, pr.py, sidra.py, commiti.py u jedan alat koji cita samo git: (a) velicina i rast: sabotaze (ukupno i po suiti), RunOne, BFP pozivi, linije test modula / produkcije / tools / kataloga, odnosi test/prod i sabotaza po testu; (b) odrzavanje kataloga po PR-u: dodato / izmenjeno zatecenih / obrisano; (c) tok: commit-a po PR-u, udeo commit-a koji diraju samo docs, lead time, udeo diff-a po kategoriji. Izlaz: tabela u GITHUB_STEP_SUMMARY na svakom push-u na main + jedan red nedeljno u docs/engineering/metrike/rast.csv (dopisuje se samo na main-u, pa nema sudara grana).
  DOBITAK: Rast postaje vidljiv svake nedelje umesto jednom u retrospektivi: +76% sabotaza i +9% produkcije za 30 dana niko nije video kao broj. Cena: ~18 s za celu istoriju (izmereno: 4.6 + 8.6 + 4.2 s), inkrementalno ispod 2 s.
  RIZIK: Nizak. Regexi za brojanje (RunOne, Test_ pozivi) mogu da zastare ako se promeni oblik suite -- zato kontrola protiv alata (broj sabotaza == sabotaza.py --lista).
  PREDUSLOVI: Odluka gde zivi CSV (predlog: docs/engineering/metrike/). Nije .claude/, pa ne trazi process PR.
  DOKAZATI: --self-test nad malim sintetickim repoom u temp-u (poznati brojevi); kontrola nad HEAD-om: 1265 / 350 / 391 moraju da se sloze sa sabotaza.py --lista i vba_gate --popis; sabotirati regex -> self-test crven.

### IM-6 Budzet za FAST sloj: meri se u CI i alarmira na rast [trud S] [slabi: NE]
Hook je za 12 nedelja porastao 19x (0.32 -> 6.19 s na ovoj masini) i niko to nije merio; model je hook ~ 0.5 s + 4.5 ms x broj sabotaza. U CI obmotati tri koraka (vba_check ceo, vba_check --hook nad jednim fajlom, sabotaza --proveri-sidra) merenjem i upisati u step summary; uporediti sa budzetom (npr. hook <= 3 s, ceo <= 15 s na CI runner-u) -- prvo kao upozorenje, kasnije kao kapija. Trajanje celog workflow-a vec postoji u GitHub API-ju (1233 run-a) i metrike.py ga moze povuci.
  DOBITAK: Sam budzet ne skracuje nista, ali vezuje optimizaciju FAST sloja (koju radi druga oblast) za broj koji ne sme da se vrati. Bez toga hook po trendu (+18 sabotaza dnevno) dolazi na ~13-14 s na ovoj masini do februara 2027. Pri ~4.800 izmena mesecno svaka sekunda hook-a je ~1.3 h cekanja.
  RIZIK: CI runner ima promenljivu brzinu -> lazni alarmi; zato medijana od 3 merenja i prag sa rezervom, u prvoj fazi samo upozorenje.
  PREDUSLOVI: IM-5 (za prikaz trenda) pozeljan, nije nuzan.
  DOKAZATI: U grani namerno usporiti proveru (sleep) -> korak prijavi prekoracenje po imenu; ukloniti -> zeleno. Uporediti CI brojeve sa lokalnim merenjem iz fast_istorija.sh.

### IM-7 Klasa ishoda i preflight okruzenja: otpad kao broj, ne kao anegdota [trud S] [slabi: NE]
run_vba i dokaz vec znaju zasto prolaz nije zelen (fatal/hard_stop, compile NEJASNO, potpis fixture-a, sidro nije nadjeno, NE OBARA SVOJ TEST, crvena baza). Upisati to kao polje klasa_ishoda u istoriju (IM-1/IM-2). Dodati preflight pre podizanja Excela: slobodan disk (STANJE: lazni padovi krecu ispod ~150 MB; incident sa tri prolaza na 61 MB) -- odbiti start ispod praga sa jasnom porukom, kao sto vec radi provera potpisa fixture-a. KPI: udeo Excel-sekundi potrosenih na prolaze koji nisu zeleni iz razloga koji nije kvar proizvoda.
  DOBITAK: Sprecava klasu incidenta 'tri RunAllTests prolaza za nula nalaza' (~11 min Excela + trijaza po dogadjaju) i daje mesecni broj otpada umesto 94 stavke proze. Usmerava sledecu optimizaciju na najskuplju klasu (po istoriji: visenje ~10 min po dogadjaju, 9+ puta).
  RIZIK: Nizak; prag diska mora biti podesiv (--ignore kao kod fixture potpisa).
  PREDUSLOVI: IM-1.
  DOKAZATI: Self-test: laziran shutil.disk_usage ispod praga -> run odbijen porukom po imenu, iznad -> prolazi; sabotaza provere -> crveno. Klasa ishoda: lazni report za svaku klasu -> ocekivana oznaka.

### IM-8 Masinski zapis merenja po PR-u umesto rucnog prepisivanja u STANJE [trud M] [slabi: NE -- pojacava: broj je vezan za otisak izvora alatom, ne recenicom.]
U 30 dana ~150 commit-a su cisti zapisi 'kapije: RunAllTests 314/0, BFP 3014/0, dokaz 8/8', 'rucni Compile PASS', 'merge #N = sha'; STANJE je narastao na 844 KB. Predlog: vba_gate.py dobija --izvestaj koji iz last_green.json + istorije (IM-1/2) pise jedan fajl po PR-u, docs/engineering/merenja/<grana ili PR>.json (suite N/0, otisak izvora, dokaz N/N, trajanja, compile marker, SHA). Fajl po PR-u -> nema sudara grana. Stavka u STANJE tada nosi jednu recenicu i referencu, a zatvorene stavke se sele u arhivu da ulazni dokument ostane kratak.
  DOBITAK: Uklanja vecinu od ~150 rucnih docs commit-a mesecno (30% svih commit-a u prozoru) i zaustavlja rast STANJE (~22 KB dnevno). Zatvara klasu greske 'preneti brojevi koji zvuce kao merenje' (25.09: osam krugova sa brojevima koji nisu mogli biti izmereni), jer broj nosi otisak izvora iz alata.
  RIZIK: Promena procesa: pravila u .claude/rules i CLAUDE.md koja traze zapis u STANJE moraju se uskladiti -- to je zaseban process PR. Ljudski kontekst (odluke, razlozi) i dalje ide u prozu.
  PREDUSLOVI: IM-1 i IM-2; odluka operatera o formatu zapisa i o arhiviranju STANJE; process PR za pravila.
  DOKAZATI: Self-test vba_gate: iz laznog markera i istorije nastaje fajl sa ocekivanim poljima; izvestaj se ODBIJA kad marker nije nad tekucim izvorom (dvosmerno, po imenu). Posle mesec dana: broj docs-only commit-a i velicina STANJE iz metrike.py.

### IM-9 KI-036: prolaz koji pusta dokaz ne upisuje zeleni marker [trud S] [slabi: NE -- marker postaje stroziji (nikad nad sabotiranim izvorom).]
Poznat i otvoren nalaz: run_vba pozvan iz dokaz.py upisuje marker zelenog nad sabotiranim izvorom (sabotaza koja ne obara tu suite), pa --require-green posle dokaza trazi ponovljen prolaz. Resenje koje KNOWN_ISSUES vec imenuje: run_vba ne upisuje marker kad ga zove dokaz (promenljiva okruzenja iz IM-1 ili zastavica), uz dvosmerni dokaz.
  DOBITAK: Do ~6 min (jedan BFP prolaz) po dokazu; gornja granica u 30 dana 62 dokaza x ~6.5 min = ~6-7 h, plus nestaje zbunjujuca poruka posle uspesnog FULL-a.
  RIZIK: Dira ugovor kapije (vba_gate) -- mora svoj dvosmerni dokaz i zaseban rez nad tools/.
  PREDUSLOVI: Nema; lakse uz IM-1 (ista oznaka pozivaoca).
  DOKAZATI: vba_gate/run_vba self-test: prolaz sa oznakom dokaza -> marker netaknut; bez oznake -> marker upisan; sabotaza uslova -> crveno po imenu. Na Windows-u: FULL, pa dokaz, pa --require-green prolazi bez ponovljenog BFP-a.

### IM-10 Pravila citiraju izmeren trosak, ne broj od pre dva meseca [trud S] [slabi: NE -- ali otkriva da je 'pun katalog pred release' po trenutnoj ceni visednevni posao; kako se to resava je odluka van ove oblasti.]
CLAUDE.md paragraf 5 ('oko 3,5 minuta po stavci, pun katalog sat i vise') i EXCEL_TEST_HARNESS.md ('BFP ~150 s') nose brojeve od 23.09 i 02.10; po poslednjim zapisima suite su 2-3x sporije, a pun katalog je ~17 h, ne 'sat i vise'. Zameniti literale referencom na generisani fajl metrika (IM-5 + IM-1) i dodati datum merenja uz svaki broj. Pravilo 'pun dokaz pred release' ostaje, ali sa stvarnom cenom pred ocima.
  DOBITAK: Odluke o obimu dokaza se donose nad tacnim redom velicine (17 h umesto 1 h). Nema direktne ustede vremena.
  RIZIK: Nizak. Izmena CLAUDE.md i .claude/ ide iskljucivo zasebnim process PR-om.
  PREDUSLOVI: IM-1/IM-5 da postoji izvor broja; do tada rucno azurirati na vrednosti sa datumom (05.11: RunAllTests 215.8 s).
  DOKAZATI: Provera u vba_check ili metrike.py da dokument ne nosi literal trajanja bez datuma ili reference (dvosmerno: vratiti stari literal -> nalaz).


## NEPOZNANICE
[
 "Zasto s/test raste (RunAllTests 0.29 -> 0.64, BFP 0.54 -> ~1.0 s po pozivu): nagomilano stanje u svesci, veci fixture posle 09.10, ili novi testovi S6-S8/doplate koji grade ekrane. Bez vremena po testu (IM-4) ne moze se razdvojiti; na Linuxu se ne moze izmeriti.",
 "Tacke serije trajanja (12.0 / 58.2 / 73.5 / 215.8 s) poticu iz proze i nije poznato da li su merene na istoj masini i pod istim uslovima (STANJE belezi prolaze sa 61-923 MB slobodnog diska). Serija ima 4-5 tacaka.",
 "Tacno trajanje BFP-a danas: 350-390 s je izvedeno iz dva posredna zapisa (prekoracenje 600 s; 34 min za 4 grupe), ne iz zapisanog '(N s)'.",
 "Stvaran broj Excel prolaza po rezu i ukupno Excel vreme mesecno: nigde se ne beleze. Procena 40-60 h je model (zapisi 'dokaz N/M' x pretpostavljeno grupisanje x prosecno trajanje), sa greskom reda +-50%.",
 "Koliko puta se hook stvarno pali: 4.798 hunk-ova i 765 parova (commit, fajl) su posredne mere; broj Edit poziva po sesiji nije dostupan iz git-a.",
 "Koliko operater ceka na rucni Compile i koliko traje rucna provera: zapisano je samo da su se desili (53 subject-a), ne i trajanje.",
 "Brojevi otpada iz STANJE/RELEASE_NOTES su pominjanja po stavci, ne prebrojani ponovljeni prolazi; ista epizoda se pominje u vise stavki, a deo ponavljanja nije ni zapisan.",
 "Klasifikacija commit-a po poruci je heuristika (prefiks i kljucne reci); klasifikacija po fajlovima je objektivna, ali commit 'produkcija+verifikacija' ne kaze koliki je deo rada bio verifikacija.",
 "Trajanje CI run-a je updated_at - run_started_at i ukljucuje cekanje runner-a; trajanje po koraku nije povuceno.",
 "Projekcije (hook ~13-14 s do februara, BFP preko 600 s sredinom decembra, pun katalog x2.7 mesecno) su linearne ekstrapolacije poslednjih 3-4 nedelje; tempo rada se menja (PR/30 dana: 113 -> 108 -> 59).",
 "VBA strana merenja vremena sa grane B nikad nije izvrsena u Excelu, pa nije poznato da li format last_run_vreme.txt i Timer rade kako je zamisljeno.",
 "Istorija pre 03.04.2026 nije povucena (repo je i dalje shallow); za eru harness-a (od 14.08) to nije bitno."
]
