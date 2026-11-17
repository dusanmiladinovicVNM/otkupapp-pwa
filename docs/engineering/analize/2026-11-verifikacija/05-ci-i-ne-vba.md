# CI i ne-VBA provere: .github/workflows/static.yml, tools/js_sintaksa.js, tests/js/, tools/migracija/ (+ provera/), gas/, src/ (PWA), i ručni rituali oko release-a/merge-a/evidencije (stablo A = origin/main @1fa4937)

## SAZETAK
CI je jedan workflow i jedan job (static.yml, ubuntu-latest): 28 koraka strogo redom, bez matrice, keša, path-filtera, concurrency i timeout-a; medijana poslednjih 25 run-ova je 76 s, a raste sa katalogom (27-40 s 12.09. -> 63 s 12.10. -> 76 s u novembru). Tri koraka nose 65-70% job-a: vba_check 22-30 s, vba_check --self-test 13-18 s, migracija self-test 13-20 s; sve ostalo je ispod 3 s po koraku, osim setup-node koji je u 2 od 6 uzoraka trajao 45-64 s. Najveće udvajanje je izmereno: vba_check --self-test je 99% (12,21 od 12,28 s) dva puna prolaza main() nad PRAVIM repoom (katalog od 1265 sabotaža dvaput, popis suita dvaput, arnosti dvaput), pa se popis suita u jednom CI run-u izvrši 4 puta. Samo CI vreme nije usko grlo (oko 17 min dnevno na javnom repou) -- usko grlo su pokrivenost i evidencija: Excel se nikad ne pušta, a rezultat Excel prolaza se ručno prepisuje u docs i commit poruke jer je marker zelenog gitignored. Lokalno postoje kapije kojih u CI-ju nema (popis_citalaca --check je zato CRVEN na main-u: 33 > prag 31; sabotaza.py --self-test; check-banka-eh; check-sef-asserts; make_dev_workbook --self-test; logo_to_vba --proveri; sintaksa .ps1 i sw.js). node tests/js/pokreni.js rc=1 nije crven test nego MODULE_NOT_FOUND (node_modules ne postoji); posle npm ci daje 47/0 i 23/0 za oko 1 s, a CI ga pušta zeleno nad 1fa4937. tools/migracija testovi (600 kom, unittest discover, bez spoljnih zavisnosti) idu automatski samo u CI-ju, 10,8 s lokalno, a po modulu paralelno 3,2 s; nemaju mehanički dvosmerni dokaz. PWA strana nema nijednu kapiju nad sw.js: dva fajla koja index.html učitava nisu u ASSETS, a precache-ovan fajl je menjan bez promene CACHE_NAME. Za Excel: repo je JAVAN, pa je self-hosted runner na razvojnoj mašini bezbednosno loš izbor; preporuka je noćni lokalni zadatak (Task Scheduler u interaktivnoj sesiji, zaseban klon) plus OBJAVA markera -- vba_gate.py već ima otisak izvora nezavisan od platforme, --require-green i --require-compile, a fale objava, trajan izveštaj dokaz.py i omotač. release.sh ne zove nijednu kapiju, a njegov .ps1 dvojnik je već izgubio dva koraka.

## CINJENICE
- [M] CI je jedan workflow (`static`) sa jednim job-om na ubuntu-latest; okidači su `pull_request` i `push` na main. Job ima 3 setup akcije (checkout@v4, setup-python@v5, setup-node@v4) i 28 run koraka koji idu strogo redom. Nema matrice, path-filtera, keša, `concurrency` ni `timeout-minutes`; nema `pip install` (svi Python alati su čist stdlib).
    DOKAZ: <scratch>/main-wt/.github/workflows/static.yml:12-22 (okidači, job), 24-243 (koraci); u fajlu ne postoje ključevi paths/strategy/cache/concurrency/timeout-minutes. GitHub MCP actions_list list_workflows: samo `static` i `pages-build-deployment`.
- [M] Poslednjih 25 run-ova `static` (06.-15.11.2026): 24 success, 1 failure. Trajanje created->updated: medijana 76 s, min 45 s, max 163 s (dva run-a sa 314/351 s imaju naduvan updated_at; njihovi job-ovi su trajali 64 s i 70 s). Ukupno 1232 run-a od 15.08.2026, tj. oko 13,4 dnevno, oko 17 min CI vremena dnevno.
    DOKAZ: GitHub MCP actions_list list_workflow_runs(resource_id=static.yml, perPage=25): run_number 1208-1232, total_count 1232; workflow created_at 2026-08-15. Job-ovi run-ova 43099160039 i 43099036442: started 20:59:19 -> completed 21:00:23, odnosno 20:57:34 -> 20:58:44.
- [M] Trajanje po koraku (6 job-ova, GitHub API): vba_check 22/23/23/25/30/30 s; vba_check --self-test 13/13/13/15/17/18 s; migracija self-test 13/14/16/19/20 s; Setup-AgriX -SelfTest (pwsh) 1-11 s; who_writes --check + --check-ownership 4-6 s zajedno; setup-node 3-5 s tipično, ali 45 s i 64 s u 2 od 6 job-ova; checkout 2-3 s; svaki drugi korak <= 1 s (JS koraci zajedno < 2 s). Job ukupno: 64 (pao rano), 70, 73, 86, 132, 160 s. Tri najskuplja koraka su 65-70% job-a bez repa setup-node.
    DOKAZ: GitHub MCP actions_list list_workflow_jobs za run-ove 43702317002 (#1232), 43099160039 (#1229), 43099036442 (#1228), 42474792419 (#1227, main @1fa4937), 42470304120 (#1225), 42464641211 (#1223): started_at/completed_at po koraku.
- [M] CI trajanje raste sa repoom: 12.09.2026. run-ovi #533/#534 27 s i 40 s; 11.-12.10. #933/#934 64 s i 63 s; novembar medijana 76 s (migracija self-test je dodat 28.10).
    DOKAZ: GitHub MCP list_workflow_runs(static.yml, perPage=2, page=350 i page=150): created_at/updated_at; git log -- .github pokazuje a1fd428 2026-10-28 feat(migracija).
- [M] Lokalno (ova mašina, 4 jezgra): 23 komande kapija iz CI-ja (bez pwsh koraka) redom traju 51,4 s, a paralelno (xargs -P 4) 22,8 s -- granica je sam vba_check. Pojedinačno: who_writes --check 1,62 s, --check-ownership 1,51 s, vba_gate --popis 0,79 s, sve ostale male Python kapije 0,02-0,15 s, js_sintaksa 0,30 s, js harness 0,25 s, js self-test 0,84 s, npm ci 0,67 s, migracija self-test 10,8 s.
    DOKAZ: Komande iz <scratch>/ci-analiza/koraci.txt puštene `while read` petljom i `xargs -P 4`; izlaz: `redom: 51.43 s`, `paralelno: 22.76 s`; pojedinačna merenja funkcijom t() (date +%s.%N).
- [M] `vba_check.py --self-test` je u 99% ponavljanje provere pravog repoa: self_test() dva puta zove main([tmp fajl]), a svaki main vrti check_katalog_sabotaza (svih 1265 sabotaža), check_dokaz_grupe, check_popis_suita i collect_definitions/members/arities nad celim src-vba. Izmereno: self_test 12,28 s, od toga main (n=2) 12,21 s; unutra check_katalog_sabotaza 5,04 s (n=2), collect_arities 4,09 s, check_popis_suita 1,19 s (n=2). Svih 150 slučajeva sa fiksturama zajedno traju < 0,1 s. Self-test usput ispisuje `provereno 1265 sabotaza`.
    DOKAZ: <scratch>/main-wt/tools/vba_check.py:3867-3909 (dva `rc = main([put])`), 4115-4117 (kapije alata pre `if not files`), 4128-4130 (collect_* nad vba_files([])). Merenje: python3 <scratch>/ci-analiza/mjeri_selftest.py -> `self_test 12.28s`, `12.21s n=2 main`. cProfile potvrđuje main n=3, check_katalog_sabotaza n=2.
- [M] Popis test suita se u jednom CI run-u izvrši 4 puta: jednom unutar koraka vba_check, dva puta unutar vba_check --self-test i jednom kao zaseban korak `vba_gate.py --popis`. Katalog sabotaža (proveri_sidra) se izvrši 3 puta (vba_check + 2x self-test).
    DOKAZ: <scratch>/main-wt/tools/vba_check.py:4115-4117; <scratch>/main-wt/.github/workflows/static.yml:79-86 i 187-188; merenje iz prethodne stavke (n=2 u self-testu).
- [M] Profil PUNOG vba_check (ne hook-a) je drugačiji od hook profila: check_file 69%, od toga check_scalar_call (ZAKLONJENO) 40% jer zove _clean(line) za svaki red tela puta svako skalarno ime -- 1,93 miliona poziva nad 233.491 linija izvora; katalog sabotaža 11,5%, collect_arities 4,8%, popis 2,2%. Memoizacija _clean samo u memoriji (bez izmene fajla) skraćuje pun prolaz sa 23,29 s na 17,73 s (-24%) uz isti verdikt `cisto (193 fajlova)`.
    DOKAZ: cProfile: python3 -m cProfile -o vbacheck_full.prof tools/vba_check.py (ukupno 60,7 s pod profilerom; check_scalar_call 24,45 s, _clean n=1927489). <scratch>/main-wt/tools/vba_check.py:945-975. Eksperiment: <scratch>/ci-analiza/mjeri_clean.py -> `kes=False main([]) 23.29s`, `kes=True main([]) 17.73s`.
- [M] Koraci su fail-fast: kad jedan padne, svi posle njega se preskaču, pa crven run ne kaže ništa o ostatku. U run-u #1229 pao je korak 23 (popis citalaca) i preskočeno je 10 koraka, među njima migracija self-test i svi JS koraci.
    DOKAZ: GitHub MCP list_workflow_jobs run 43099160039: korak 23 `popis citalaca -- pragovi starog modela` conclusion=failure, koraci 24-33 conclusion=skipped.
- [M] Runner upozorava da su korišćene akcije zastarele: `Node.js 20 is deprecated ... forced to run on Node.js 24: actions/checkout@v4, actions/setup-node@v4, actions/setup-python@v5`. Workflow traži `node-version: '20'`. Harness lokalno prolazi i na Node v22.23.4.
    DOKAZ: GitHub MCP get_job_logs job 128699269974 (run #1225), poslednja linija loga; <scratch>/main-wt/.github/workflows/static.yml:24-28, 211-213; lokalno `node --version` = v22.23.4 uz 47/0 i 23/0.
- [M] tools/migracija testovi SE puštaju automatski, ali samo u CI-ju: korak `migracija self-test` zove unittest.defaultTestLoader.discover(pattern test_*.py) koji ulazi i u paket provera/. Ukupno 600 testova: test_migracija 110, provera.test_otkup 95, test_novac 89, test_prijemnica 86, test_amb_pal_fak 80, test_otpremnica_zbirna 71, test_storno 58, test_zajednicko 11. Bez pytest-a i bez spoljnih zavisnosti. Lokalno 10,6-10,9 s (ranije mereno 15,1 s), u CI-ju 13-20 s. Četiri performansna testa nose 3,24 s (30%). Po modulu paralelno (P=4): 3,17 s.
    DOKAZ: <scratch>/main-wt/.github/workflows/static.yml:197-198; <scratch>/main-wt/tools/migracija/migracija.py:391-394. Izlaz: `Ran 600 tests in 10.587s OK`; brojanje i vreme po modulu/testu skriptom nad discover-om (provera.test_amb_pal_fak 3,07 s; CistaSveska.test_performanse 1,29 s); `xargs -P 4 python3 -m unittest -q <modul>` -> 3,17 s.
- [M] migracija kapija čita i fajlove van svog foldera (schema/schema.json kao kanon; tools/make_dev_workbook.py), a nije pomenuta ni u CLAUDE.md §5 ni u .claude/rules/testovi.md, niti je u hook-u. Za Python testove ne postoji mehanički dvosmerni dokaz (katalog mutacija); dokaz je rađen ručno po rezu.
    DOKAZ: <scratch>/main-wt/tools/migracija/dokumenti.py:22-23 (KANON = schema/schema.json); <scratch>/main-wt/tools/migracija/migracija.py:342-343; grep `migracija` u CLAUDE.md i .claude/ = prazno; `grep -rln -i sabotaz tools/migracija` = samo test_migracija.py; commit poruka 7c3e2e1: `Oba dokazana sabotazom`.
- [M] `node tests/js/pokreni.js` rc=1 na stablu A NIJE crven test: proces pada pri učitavanju sa MODULE_NOT_FOUND `fake-indexeddb/auto` jer node_modules ne postoji (gitignored, instalira ga tek CI). Posle `npm ci` (0,67 s) u scratch folderu i NODE_PATH-a: `js harness: 47/0 (proslo/palo) od 47` za 0,25 s i `sabotaza: 23/0 (dokazano/palo) od 23` za 0,84 s. CI ga pušta zeleno nad 1fa4937 jer pre harness-a ide `npm ci`.
    DOKAZ: Izlaz: `Error: Cannot find module 'fake-indexeddb/auto'` sa require stack-om <scratch>/main-wt/tests/js/suites/db-claim.js:21; <scratch>/main-wt/.gitignore:31; static.yml:233-243; run 42474792419 koraci 30-32 success; log job-a 128699269974: `js harness: 47/0`, `sabotaza: 23/0`.
- [M] Jedna nedostajuća zavisnost obara svih 5 JS suita, iako je treba samo db-claim: pokreni.js učitava sve suite na vrhu modula, pa greška stiže kao sirov stack trace pre ijedne tvrdnje.
    DOKAZ: <scratch>/main-wt/tests/js/pokreni.js:15-21; `grep -n fake-indexeddb tests/js/suites/*.js` pogađa samo db-claim.js.
- [M] JS harness meri 4 produkciona fajla: src/js/services/db.js, src/js/features/vozac/zbirna.js, src/js/features/otkup/otkup-stavke.js i gas/Code.gs -- od 52 fajla u src/js (20.336 linija) i 4 .gs fajla (10.772 linije). 47 tvrdnji, 23 sabotaže; 24 tvrdnje nemaju sabotažu (ispisuje se, ne obara).
    DOKAZ: `grep -n "const FAJL" tests/js/suites/*.js`; `find src/js -name '*.js' | wc -l` = 52; `wc -l gas/*.gs gas/bank-pdf-downloader/*.gs` = 10772; izlaz self-testa: `bez sabotaze (24): ...`.
- [M] JS harness čita i VBA izvor (ugovor kolona PWA/VBA), pa path-filter koji bi JS korake vezao samo za src/, gas/ i tests/js oslabio bi tu tvrdnju.
    DOKAZ: <scratch>/main-wt/tests/js/suites/gas-otk-stavke.js:109-110 (procitaj('src-vba/modMasterSync.bas'), procitaj('src-vba/modConfig.bas')).
- [M] js_sintaksa pokriva src/js, gas, tests/js i tools (plitko) = 65 fajlova. Ne pokriva sw.js u korenu ni inline skriptu u parcel-draw.html.
    DOKAZ: <scratch>/main-wt/tools/js_sintaksa.js:24-29; izlaz `js_sintaksa: 65/0 (proslo/palo) od 65 fajlova`; `grep -n "<script" parcel-draw.html` -> linija 122 inline.
- [M] JS kapije nisu pomenute ni u CLAUDE.md, ni u .claude/rules, ni u docs/EXCEL_TEST_HARNESS.md. static.yml tvrdi da je CI jedini izvršilac jer node ne postoji na Windows mašini; u web (Linux) sesiji node postoji i sve četiri JS komande rade posle `npm ci` za oko 1,4 s.
    DOKAZ: `grep -rn "tests/js\|js_sintaksa\|npm ci\|pokreni.js" .claude CLAUDE.md docs/EXCEL_TEST_HARNESS.md` = prazno; <scratch>/main-wt/.github/workflows/static.yml:207-210; lokalna merenja iznad.
- [M] PWA nema nijednu kapiju nad service worker-om, a stanje je već odstupilo: (a) index.html učitava src/js/features/otkup/otkup-stavke.js i src/js/features/vozac/status.js, a nijedan nije u sw.js ASSETS; (b) sw.js je cache-first za sve osim dokumenata i briše keš samo kad se promeni CACHE_NAME ('AgriX-v25'), a precache-ovan src/js/features/management/kooperanti.js je menjan 03.11.2026 (a4f1760) bez izmene sw.js; index.html nema ?v= u putanjama skripti. Uticaj u pregledaču NIJE proveren -- ovo je statički nalaz.
    DOKAZ: Petlja nad `find src/js -name '*.js'` naspram sw.js i index.html: `van kesa: src/js/features/otkup/otkup-stavke.js`, `van kesa: src/js/features/vozac/status.js` (oba prisutna u index.html). <scratch>/main-wt/sw.js:2, 72 i fetch handler (cache first); `git show --stat a4f1760 -- src sw.js`; `git log -- sw.js` bez commit-a u vidljivoj istoriji; `grep -c 'src="[^"]*?v=' index.html` = 0.
- [M] Kapije koje postoje lokalno a NISU u CI-ju (sve rade bez Excela): tools/popis_citalaca.py --check (8,8 s; rc=1 na main-u: `prj_brojzbirne 33 > prag 31 PAO`), tools/sabotaza.py --self-test (0,69 s, 43 slučaja), tools/check-banka-eh.py (0,85 s), tools/check-sef-asserts.py (0,02 s), tools/make_dev_workbook.py --self-test (0,04 s), tools/logo_to_vba.py --proveri (0,90 s). Sem prve sve su zelene. PR #474 koji dodaje popis_citalaca u CI je otvoren i sam crven dok #469 ne podigne prag.
    DOKAZ: Lokalno pokretanje svake komande sa merenjem (rc i poslednje linije izlaza); `grep -n "sabotaza\|popis_citalaca\|check-banka\|check-sef" .github/workflows/static.yml` = prazno; run 43099160039 (PR #474) korak `popis citalaca` failure, commit poruka: `Pre toga je korak na main-u crven (exit 1)`.
- [M] proveri_sidra (kapija nad katalogom sabotaža) ide u CI kroz vba_check, ali njen sopstveni self-test (`sabotaza.py --self-test`) ne ide nigde automatski -- suprotno pravilu repoa da checker bez pokazanog crvenog ne dokazuje ništa. dokaz.py --self-test jeste pokriven (zove ga vba_check kroz check_dokaz_grupe).
    DOKAZ: <scratch>/main-wt/tools/vba_check.py:4022-4033 (zove samo modul.proveri_sidra), 4046-4057 (dokaz._self_test); <scratch>/main-wt/tools/sabotaza.py:13568 (_self_test) bez pozivaoca u static.yml.
- [M] CI parsira samo .sh skripte (bash -n nad .claude/hooks/*.sh i tools/*.sh). Četiri PowerShell skripte se ne proveravaju: tools/check_merge.ps1, tools/release.ps1, tools/stamp-build.ps1 (install/Setup-AgriX.ps1 se pokreće kroz -SelfTest, pa je implicitno parsiran).
    DOKAZ: <scratch>/main-wt/.github/workflows/static.yml:59-77, 104-106; `find . -name '*.ps1'` -> 4 fajla.
- [M] Spisak kapija je ručno prepisan na tri mesta i već se razišao: testovi.md kaže da ŠEST alata ima --self-test i da gen_schema_module NEMA self-test, a static.yml ga pušta; testovi.md ne pominje vba_import_marker_gate, vba_gate --popis, migraciju, JS ni Setup-AgriX.
    DOKAZ: <scratch>/main-wt/.claude/rules/testovi.md:63-75 naspram <scratch>/main-wt/.github/workflows/static.yml:165-171, 135-139, 187-198, 215-243.
- [M] Ne postoji jedna lokalna komanda koja pušta isto što i CI (nema Makefile, pyproject, pre-commit). Hook pušta samo `vba_check.py --hook <fajl>` za VBA fajlove i JSON validaciju settings fajlova; tools/test_grana.sh pušta samo vba_check, gen_schema --check i schema_diff. Ostalih ~25 koraka lokalno je ručno.
    DOKAZ: <scratch>/main-wt/.claude/hooks/vba-check.sh:36-66; <scratch>/main-wt/tools/test_grana.sh (sekcija `kapije`); `ls requirements*.txt pyproject.toml Makefile .pre-commit-config.yaml` -> ne postoje.
- [M] Repo je JAVAN, a CI po sopstvenoj izjavi nikad ne pušta Excel; `--require-green` namerno nije u CI-ju jer je marker zelenog mašinsko stanje.
    DOKAZ: mcp claude-code-remote list_repos(query=otkupapp-pwa): `visibility: public`; <scratch>/main-wt/.github/workflows/static.yml:8-10, 184-186.
- [M] Infrastruktura za vezivanje Excel dokaza za izvor već postoji u tools/vba_gate.py: marker tests/last_green.json (gitignored) pamti po suite-u status, brojeve, otisak izvora, otisak test-ugovora (runner + fixture generator + golden + kapija + verzija 4), identitet sveske, git glavu i platformu; upis je fail-closed nad snimkom uzetim pre run-a; CLI ima --status, --hash, --require-green [--suite] [--require-compile] [--sveska], --mark-compile. Otisak izvora je LF-normalizovan pa je isti na Linux-u i Windows-u: `--hash` ovde radi za 0,10 s. Nedostaje samo objava van mašine.
    DOKAZ: <scratch>/main-wt/tools/vba_gate.py:1-106 (docstring), 123, 130, 155, 644-714 (zabelezi_prolaz), 735-798 (zahtevaj_zeleno), 1367-1385 (CLI). Lokalno: `python3 tools/vba_gate.py --hash` -> 656a1563... za 0,096 s; `--status` -> `marker: nema ga`.
- [M] run_vba.py traži interaktivnu desktop sesiju, ne samo Windows + Excel: watchdog nabraja prozore (win32gui.EnumWindows/IsWindowVisible), compile pokušaj pali VBE prozor i šalje SendKeys posle AppActivate, a preduslov je `Trust access to the VBA project object model`. Rezultat svakog run-a već ide u mašinski čitljiv tests/last_run.json i u marker.
    DOKAZ: <scratch>/main-wt/tools/run_vba.py:6-8, 193, 201, 530, 572-635, 1038-1069.
- [M] dokaz.py nema trajan izveštaj (samo stdout) ni nastavak posle prekida; kvari src-vba na disku (vraća u finally i poredi potpis) i traži zelenu bazu pre prve mutacije. Plan `--grupe 6` nad katalogom od 1265 daje 215 prolaza: 134 RunAllTests + 78 RunBusinessFlowProSuite + 3 RunBankaImportTestSuite.
    DOKAZ: <scratch>/main-wt/tools/dokaz.py:1-55, 433-443 (nema argumenta za izlazni fajl; grep json/open -> samo čitanje izvora na liniji 87). `python3 tools/dokaz.py --grupe 6 --plan` -> `sabotaza 1265, prolaza 215`, brojanje grupa po suite-u u izlazu.
- [N] Procena trajanja dokaza iz dokumentovanih brojeva (RunAllTests 68-91 s, BFP ~150 s, režija ~18 s po pozivu): grupni dokaz --grupe 6 oko 7-8 h (134 x 86-109 s + 78 x 168 s); pojedinačni dokaz celog kataloga 40-74 h (gornja granica po CLAUDE.md: 3,5 min x 1265).
    DOKAZ: Računica nad planom iz prethodne stavke i brojevima iz docs/EXCEL_TEST_HARNESS.md:732 i CLAUDE.md §5; ništa od toga nije izvršeno u Excelu.
- [M] Excel dokaz se danas prenosi ručno: operater pusti run_vba/dokaz, rezultat ide u chat, pa u docs commit. U poslednjih ~30 dana 48 commit poruka nosi NEVERIFIKOVANO, 111 linija u telima commit-a nosi prepisane brojeve prolaza, a docs/STANJE_REFAKTORA.md (11.059 linija) je diran u 252 od 493 commit-a. Sam dokument kaže zašto: marker je gitignored, pa ga review koji čita repo ne vidi.
    DOKAZ: `git log --since=2026-10-09 --format=%B | grep -c -i NEVERIFIKOVANO` = 48; grep brojeva prolaza = 111; `git log --since=2026-10-09 --no-merges --format=%h -- docs/STANJE_REFAKTORA.md | wc -l` = 252 od 493; <scratch>/main-wt/docs/STANJE_REFAKTORA.md:1437-1440; head commit run-a #1226: `docs: KI-094 posle zelenog BFP-a, RunAllTests i dokaza`.
- [M] tools/release.sh ne zove nijednu proveru: radi checkout main, pull, bump APP_VERSION, commit i push direktno na main, tag, push taga, stamp -- bez vba_check, bez provere CI statusa i bez vba_gate --require-green/--require-compile. Posle toga ostaje 5-6 ručnih Excel koraka i ručno vraćanje modBuildInfo.bas.
    DOKAZ: <scratch>/main-wt/tools/release.sh (ceo fajl, 5 koraka + ispis ručnih koraka); `grep -n "require-green\|vba_check" tools/release.sh` = prazno.
- [M] release.ps1 je dvojnik release.sh koji se već razišao: .sh u ručnim koracima navodi AssertBlankBuild i PublishReleaseToDrive, .ps1 ne. Isto udvajanje postoji za stamp-build.sh/.ps1, a CLAUDE.md propisuje da je ljuska korisnika isključivo Git Bash.
    DOKAZ: Poređenje ispisa ručnih koraka u <scratch>/main-wt/tools/release.sh i <scratch>/main-wt/tools/release.ps1; CLAUDE.md, pasus o Git Bash-u.
- [M] Otisak izvora izuzima samo modBuildInfo.bas; release bump menja APP_VERSION u modConfig.bas, pa marker zelenog po konstrukciji prestaje da važi za release commit.
    DOKAZ: <scratch>/main-wt/tools/vba_gate.py:155, 424-438; release.sh korak 2/5 (sed nad modConfig.bas).
- [M] Sastav commit-a (493 bez merge-a od 09.10.2026): 51% dira samo docs/; 99% ne dira PWA sloj; 95% ne dira tools/migracija; 62 PR merge-a u ~30 dana. U uzorku od 25 run-ova najmanje 6 ima docs-only head commit, a svaki pušta pun CI.
    DOKAZ: `git log --no-merges --since=2026-10-09 --name-only` klasifikovan skriptom: `samo docs: 253 51%`, `ne dira pwa: 487 99%`, `ne dira tools/migracija: 467 95%`; `git log --merges | wc -l` = 62; head commit poruke run-ova #1230, #1226, #1221, #1220, #1219, #1208.
- [M] Generisani artefakti u gitu prave ponovljen ručni posao pri spajanju: WHO_WRITES.md je menjan u 21 commit-u, modSchema.bas u 14, a test_grana.sh ima posebnu granu baš za sukob u WHO_WRITES.md. who_writes --check i --check-ownership su dva procesa koji oba rade isti scan(); zajedno zadate zastavice izvrše samo ownership.
    DOKAZ: `git log --since=2026-10-09 --no-merges --format=%h -- docs/DOMEN/WHO_WRITES.md | wc -l` = 21; isto za src-vba/modSchema.bas = 14; <scratch>/main-wt/tools/test_grana.sh:129-132; <scratch>/main-wt/tools/who_writes.py:457-476.
- [M] Grana sa izborom testova (stablo B, commit 9d6e780) ne dira ni CI ni JS ni migraciju: menja docs/EXCEL_TEST_HARNESS.md, dva VBA test modula i tools/dokaz.py, run_vba.py, sabotaza.py, vba_gate.py.
    DOKAZ: `git -C /home/user/otkupapp-pwa show --stat --format= HEAD` -> 7 fajlova, bez .github, tests/js, tools/migracija.
- [N] GitHub-hosted Windows runner-i nemaju instaliran Office; GitHub preporučuje da se self-hosted runner ne koristi na javnom repou (PR iz forka može da izvrši kod na runner mašini); Microsoft ne podržava automatizaciju Office-a iz neinteraktivnog naloga/servisa.
    DOKAZ: Opšte znanje iz dokumentacije (GitHub runner-images, GitHub Docs: Security hardening / self-hosted runners, Microsoft KB 257757) -- nije provereno u ovoj sesiji.

## TROSKOVI
[
 {
  "sta": "Jedan CI run (`static`), ceo job",
  "koliko": "medijana 76 s (45-163 s); oko 17 min dnevno ukupno",
  "izmereno": true,
  "koliko_cesto": "oko 13,4 run-a dnevno (1232 od 15.08.2026); svaki push na PR i svaki merge u main"
 },
 {
  "sta": "Korak vba_check (ceo repo)",
  "koliko": "22-30 s u CI-ju; 23-31 s lokalno; od toga 40% check_scalar_call, 11,5% katalog sabotaža",
  "izmereno": true,
  "koliko_cesto": "svaki CI run"
 },
 {
  "sta": "Korak vba_check --self-test",
  "koliko": "13-18 s u CI-ju; 12,3-16,9 s lokalno; 99% je dvostruko ponavljanje provere pravog repoa",
  "izmereno": true,
  "koliko_cesto": "svaki CI run"
 },
 {
  "sta": "Korak migracija self-test (600 unittest testova)",
  "koliko": "13-20 s u CI-ju; 10,6-15,1 s lokalno; 3,2 s kad se moduli puste paralelno",
  "izmereno": true,
  "koliko_cesto": "svaki CI run, iako 95% commit-a ne dira tools/migracija"
 },
 {
  "sta": "setup-node@v4 sa node-version 20",
  "koliko": "3-5 s tipično; 45 s i 64 s u 2 od 6 uzorkovanih job-ova",
  "izmereno": true,
  "koliko_cesto": "svaki CI run, iako 99% commit-a ne dira PWA sloj"
 },
 {
  "sta": "Setup-AgriX.ps1 -SelfTest (pwsh)",
  "koliko": "1-11 s u CI-ju; lokalno nije mereno (pwsh ne postoji u ovoj sesiji)",
  "izmereno": true,
  "koliko_cesto": "svaki CI run"
 },
 {
  "sta": "who_writes --check + --check-ownership (dva procesa, isti scan)",
  "koliko": "4-6 s u CI-ju; 3,1 s lokalno",
  "izmereno": true,
  "koliko_cesto": "svaki CI run"
 },
 {
  "sta": "JS kapije (js_sintaksa x2, npm ci, harness, harness --self-test)",
  "koliko": "< 2 s u CI-ju; 1,4 s + 0,7 s npm ci lokalno",
  "izmereno": true,
  "koliko_cesto": "svaki CI run; lokalno na Windows mašini nikad (nema node)"
 },
 {
  "sta": "Fail-fast: crven korak sakriva verdikt svih narednih",
  "koliko": "10 koraka preskočeno u run-u #1229; cena je dodatni push + novi run (~76 s) po skrivenom nalazu",
  "izmereno": true,
  "koliko_cesto": "pri svakom crvenom run-u"
 },
 {
  "sta": "Pun CI za docs-only izmenu",
  "koliko": "~76 s po run-u bez ijedne relevantne kapije (osim who_writes kad se dira docs/DOMEN)",
  "izmereno": true,
  "koliko_cesto": "51% commit-a je docs-only; najmanje 6 od poslednjih 25 run-ova"
 },
 {
  "sta": "Lokalno puštanje svih statičkih kapija ručno (nema jedne komande)",
  "koliko": "51,4 s redom, 23 komande koje treba znati napamet ili prepisati iz static.yml",
  "izmereno": true,
  "koliko_cesto": "pred svaki commit koji dira više oblasti; u praksi se preskače pa nalaz stiže tek iz CI-ja"
 },
 {
  "sta": "popis_citalaca --check (lokalna kapija van CI-ja)",
  "koliko": "8,8 s lokalno, 10 s u CI-ju; trošak izostanka: prag probijen kroz tri PR-a (#471-#473) bez crvenog koraka",
  "izmereno": true,
  "koliko_cesto": "jednom se već desilo; kapija je crvena na main-u od 08.11.2026"
 },
 {
  "sta": "Ručno prenošenje Excel dokaza u docs i commit poruke",
  "koliko": "48 commit-a sa NEVERIFIKOVANO i 111 prepisanih linija rezultata u ~30 dana; minuti po prenosu nisu mereni",
  "izmereno": true,
  "koliko_cesto": "posle skoro svakog reza koji menja ponašanje (62 PR-a mesečno)"
 },
 {
  "sta": "FULL prolaz u Excelu (run_vba bez argumenata)",
  "koliko": "20-60 min po dokumentaciji",
  "izmereno": false,
  "koliko_cesto": "pred release i za rizične izmene; nikad automatski"
 },
 {
  "sta": "Pun dvosmerni dokaz nad katalogom od 1265 sabotaža",
  "koliko": "procena 40-74 h pojedinačno; 7-8 h grupno (--grupe 6, 215 prolaza)",
  "izmereno": false,
  "koliko_cesto": "po pravilu pred release; u praksi podskup po rezu"
 },
 {
  "sta": "Release ritual",
  "koliko": "1 skripta (git deo) + 5-6 ručnih Excel koraka + ručno vraćanje modBuildInfo.bas; trajanje nije mereno",
  "izmereno": false,
  "koliko_cesto": "po release-u"
 },
 {
  "sta": "Održavanje dvojnika release.sh/.ps1 i stamp-build.sh/.ps1",
  "koliko": "4 fajla umesto 2; dva koraka već izgubljena u .ps1 kopiji",
  "izmereno": true,
  "koliko_cesto": "pri svakoj izmeni release procedure"
 },
 {
  "sta": "Ručni bump CACHE_NAME u sw.js pri izmeni PWA fajlova",
  "koliko": "jedna linija; trošak izostanka je zastareo JS kod već instaliranih klijenata (nije provereno u pregledaču)",
  "izmereno": false,
  "koliko_cesto": "pri svakoj izmeni fajla iz ASSETS; propušteno bar jednom (a4f1760)"
 }
]

## PREDLOZI
### CI-1 vba_check --self-test: ne ponavljati kapije alata i skeniranje celog repoa u dva CLI slučaja [trud S] [slabi: NE. Oba CLI slučaja i dalje idu kroz main() nad pravim fajlom i traže exit 2 po imenu nalaza; katalog nad pravim repoom i dalje proverava korak vba_check. Dodaje se tvrdnja koja danas ne postoji (CLI zove kapije alata).]
self_test() dva puta zove main([tmp fajl]) (tools/vba_check.py:3877, 3901), a svaki main pušta check_katalog_sabotaza, check_dokaz_grupe, check_popis_suita (4115-4117) i collect_definitions/members/arities nad celim src-vba (4128-4130). Predlog: (a) main dobija interni keyword parametar (bez CLI prekidača) kojim self-test isključuje tri kapije alata u ta dva poziva; (b) collect_* nad vba_files([]) se memoizuje u procesu, pa drugi poziv ne skenira ponovo; (c) dodaje se JEDAN nov slučaj koji stub-ovima dokazuje da main zaista zove sve tri kapije alata i da njihov rc=2 izlazi kao rc=2 -- to se danas nigde ne tvrdi, samo se usput izvrši.
  DOBITAK: Self-test lokalno 12,3 s -> oko 1,5-2 s (izmereno: 12,21 od 12,28 s je u dva main poziva); CI korak 13-18 s -> oko 2-3 s, tj. 15-20% job-a. Katalog se u CI run-u proverava 1x umesto 3x.
  RIZIK: Nizak. Jedini put ka slabljenju je da se isključenje kapija omakne u produkcioni poziv -- zato samo keyword parametar bez CLI prekidača, plus nov slučaj (c) koji baš to hvata.
  PREDUSLOVI: Nema. Izmena checker-a, pa je dvosmerni dokaz obavezan.
  DOKAZATI: Vreme pre/posle istom skriptom (scratchpad/ci-analiza/mjeri_selftest.py). Sabotaža 1: otkači `findings += check_file(...)` u main -> slučaj `ceo CLI nad pravim .bas fajlom` mora pasti. Sabotaža 2: otkači `rc_kat = check_katalog_sabotaza(...)` -> novi slučaj mora pasti po imenu. U CI-ju: trajanje koraka kroz API u 5 uzastopnih run-ova.

### CI-2 Higijena workflow-a: Node 22, novije verzije akcija, timeout-minutes, concurrency za PR [trud S] [slabi: NE]
node-version '20' -> '22' (harness lokalno prolazi na v22.23.4: 47/0 i 23/0; fake-indexeddb 6.0.0 traži >= 18); checkout/setup-node/setup-python na verzije koje runner ne prijavljuje kao zastarele; `timeout-minutes: 10` na job; `concurrency: group: static-${{ github.ref }}` sa cancel-in-progress samo za pull_request (push na main mora uvek dobiti verdikt).
  DOBITAK: Uklanja rep od 45-64 s u setup-node (2 od 6 uzoraka; pretpostavka je da se Node 20 više ne nalazi u tool cache-u pa se preuzima). Timeout ograničava viseći korak na 10 min umesto podrazumevanih 6 h (pokreni.js sam navodi da suite sa nezatvorenim pozadinskim poslom visi). Concurrency otkazuje zastarele run-ove (11 run-ova na grani PR #471 za dva dana).
  RIZIK: Nizak. Razlika u ponašanju Node 20/22 je pokrivena lokalnim prolazom; cancel-in-progress ne sme da važi za main.
  PREDUSLOVI: Nema; .github nije pod pravilom zasebnog process PR-a, ali je zaseban PR čistiji.
  DOKAZATI: 20 uzastopnih job-ova sa setup-node <= 5 s (API); na test grani korak koji namerno spava -> prekid posle timeout-a; dva brza push-a na isti PR -> prvi run otkazan.

### CI-3 Podela na paralelne job-ove uz jedan zbirni check; pun verdikt i kad nešto padne [trud S] [slabi: NE]
Četiri job-a: `vba` (vba_check, self-test i 14 malih kapija redom), `migracija` (moduli testova pušteni paralelno), `pwa` (node koraci), `skripte` (JSON, bash -n, pwsh). Zbirni job `static` sa `needs` i `if: always()` pada ako ijedan padne, pa ime obaveznog check-a ostaje isto. Alternativa bez podele: jedan korak koji lokalnim pokretačem (CI-5) pušta kapije paralelno -- brže, ali se gubi pregled po koraku u Actions UI.
  DOBITAK: Zid 73-86 s -> oko 50-55 s odmah (granica je `vba` job), oko 30-35 s uz CI-1 i CI-11. Migracija 10,8 s -> 3,2 s (izmereno, P=4). Pad jedne kapije više ne sakriva ostale (u run-u #1229 je preskočeno 10 koraka). Spor setup-node više nije na kritičnom putu VBA verdikta.
  RIZIK: Četiri pokretanja runner-a umesto jednog (po ~5 s, besplatno na javnom repou). Ako postoji branch protection, zbirni job mora zadržati ime `static`.
  PREDUSLOVI: Proveriti da li je `static` obavezan check u podešavanjima grane.
  DOKAZATI: Namerno crven popis suita na test PR-u -> `migracija` i `pwa` i dalje daju svoj verdikt, zbirni `static` je crven. Medijana trajanja iz API-ja pre/posle na 20 run-ova.

### CI-4 Uvesti u CI kapije koje danas postoje samo lokalno [trud M] [slabi: NE -- jača: proveri_sidra dobija svoj self-test u CI-ju, a pragovi čitalaca prestaju da zavise od pamćenja.]
Dodati: `sabotaza.py --self-test` (0,69 s -- self-test kapije proveri_sidra koja već jeste u CI-ju), `check-banka-eh.py` (0,85 s), `check-sef-asserts.py` (0,02 s), `make_dev_workbook.py --self-test` (0,04 s), `logo_to_vba.py --proveri` (0,90 s), `popis_citalaca.py --check` (8,8 s; PR #474 već postoji, ide posle #469), parsiranje tools/*.ps1 kroz pwsh parser, i sw.js u spisak js_sintaksa.
  DOBITAK: Zatvara klasu `kapija postoji, ali je niko ne pušta`: popis_citalaca je probijen kroz tri PR-a i sada je crven na main-u. Cena: oko +12,5 s redom (od toga 8,8 s popis_citalaca), bez uticaja na zid ako ide u paralelnom job-u.
  RIZIK: check-banka-eh i check-sef-asserts nemaju --self-test: po pravilu repoa treba im dvosmerni dokaz pre nego što im se veruje kao kapiji (to je M deo posla). popis_citalaca mora sačekati odluku o pragu (#469).
  PREDUSLOVI: Merge #469 pre #474; self-test za dva check-* alata.
  DOKAZATI: Za svaki dodat korak jedan namerno crven commit na test grani (npr. ubačen `Err.Raise Err.Number` posle LogErr u banka modul; prag čitalaca spušten za 1) -> korak pada po imenu; vrati -> zeleno.

### CI-5 Jedan manifest kapija + lokalni pokretač (paralelno, izbor po putanjama sa dokazanim ulazima) [trud M] [slabi: NE dok CI pušta sve. Path-filter u samom CI-ju se NE predlaže: ušteda je mala (javni repo, 76 s), a slabio bi pokrivenost po commit-u za kapije sa unakrsnim ulazima.]
Spisak kapija (ime, komanda, da li traži node/pwsh, ulazni glob-ovi) živi na jednom mestu; `python tools/kapije.py` pušta sve paralelno, a `--izmenjeno <baza>` samo one čiji ulazi seku `git diff --name-only`. static.yml se generiše iz manifesta ili ima `--check` da je u koraku; testovi.md i CLAUDE.md §5 upućuju na manifest umesto da prepisuju spisak. Ulazi se ne veruju na reč: self-test pušta svaku Python kapiju pod sys.addaudithook('open') i tvrdi da je svaki pročitan fajl iz repoa pokriven deklarisanim glob-om (to bi uhvatilo da JS harness čita src-vba/modMasterSync.bas i modConfig.bas, a migracija schema/schema.json).
  DOBITAK: Lokalno pun set 51,4 s -> 22,8 s (izmereno). Za docs-only izmenu (51% commit-a) lokalno 0-3 s. Jedan spisak umesto tri koja su se već razišla (testovi.md tvrdi šest self-testova i da gen_schema nema self-test). Agent u web sesiji dobija isti verdikt kao CI pre push-a.
  RIZIK: Izbor po putanjama je jedino mesto u ovom planu gde se tvrdnja može oslabiti: nepotpun spisak ulaza znači da se kapija ne pusti kad treba. Zato CI (pull_request i push na main) i dalje pušta SVE; izbor važi samo lokalno, a audit-hook self-test čuva spisak ulaza.
  PREDUSLOVI: CI-3 poželjan. Izmena .claude/rules/testovi.md ide zasebnim process PR-om.
  DOKAZATI: Sabotaža manifesta: izbaci src-vba/modConfig.bas iz ulaza JS harness-a -> self-test manifesta crven po imenu kapije. Poređenje: skup kapija koje pusti `kapije.py` bez argumenata == skup koraka u static.yml. Vreme pre/posle.

### CI-6 JS kapije upotrebljive lokalno: jasna greška bez node_modules, pravilo u .claude/rules, sintaksni hook [trud S] [slabi: NE]
(a) pokreni.js hvata MODULE_NOT_FOUND za fake-indexeddb i ispisuje jednu liniju `node_modules nema -- pusti npm ci`, i dalje sa nenultim izlazom (crveno, ne preskočeno); (b) kratko pravilo sa paths src/js/**, gas/**, tests/js/** i tri komande (npm ci, npm test, npm run self-test) -- zaseban process PR; (c) PostToolUse hook za *.js/*.gs pušta `node tools/js_sintaksa.js` kad node postoji (0,3 s), a tiho izlazi kad ne postoji, istim obrascem kojim vba-check.sh traži python.
  DOBITAK: U web sesiji JS verdikt za ~1,4 s lokalno umesto push + ~76 s CI. Stack trace se više ne čita kao crven test (upravo se to desilo u ovom zadatku). Agent koji menja gas/ ili src/ saznaje da harness postoji.
  RIZIK: Na Windows mašini bez node hook tiho ne radi i izgleda kao čist -- isto ograničenje kao danas, mora da piše u pravilu da je CI tamo jedini verdikt.
  PREDUSLOVI: Process PR za .claude/ (pravilo i hook), odvojeno od izmene pokreni.js.
  DOKAZATI: Bez node_modules: izlaz je jedna imenovana linija i rc != 0. Sa node_modules: 47/0. Hook: namerno neuravnotežena vitica u .js fajlu -> exit 2 sa imenom fajla; isti hook sa `PATH` bez node -> exit 0 bez izlaza.

### CI-7 Statička kapija nad PWA ljuskom: index.html, sw.js ASSETS i CACHE_NAME [trud S] [slabi: NE -- nova tvrdnja.]
Nova mala kapija (Node, sa --self-test): (1) svaki lokalni src/href iz index.html postoji na disku; (2) svaki je u sw.js ASSETS; (3) sw.js ulazi u sintaksnu proveru; (4) u PR-u: ako je promenjen sadržaj bilo kog fajla iz ASSETS, a CACHE_NAME nije -- pad (poređenje sa merge-base). Varijanta bez rituala: CACHE_NAME izveden iz heša sadržaja ASSETS, generisan i čuvan `--check`-om kao modSchema.bas.
  DOBITAK: Hvata dva zatečena odstupanja (otkup-stavke.js i status.js van ASSETS) i klasu `izmena PWA bez promene keša` (kooperanti.js u a4f1760). Trajanje ispod 0,5 s. Prva kapija nad jedinim delom sistema koji se isporučuje samim merge-om u main (GitHub Pages).
  RIZIK: Stavka (4) traži istoriju u CI checkout-u (fetch-depth ili merge-base) i daje crveno i za izmenu komentara -- prihvatljivo jer je bump jedna linija. Stvarni uticaj zastarelog keša u pregledaču nije proveren.
  PREDUSLOVI: Odluka da li se dva fajla dodaju u ASSETS i da li se CACHE_NAME diže odmah (to je izmena ponašanja PWA, ne samo kapija).
  DOKAZATI: Self-test sa fiksturama u memoriji: skripta van ASSETS -> crveno; sve u ASSETS -> zeleno; izmenjen fajl bez bump-a -> crveno; sa bump-om -> zeleno. Nad pravim repoom pre ispravke mora pasti na dva imenovana fajla.

### CI-8 Objava Excel dokaza: marker zelenog izlazi iz gitignore-a kao commit status + zapis po otisku izvora [trud M] [slabi: NE, pod uslovom da lokalni --require-green ostane nepromenjen i da se udaljeni ishod ne predstavlja kao isti dokaz.]
`vba_gate.py --objavi`: posle zelenog run_vba (i posle --mark-compile) šalje commit status `excel/full` (i `excel/compile`) na HEAD sa opisom `N/N suita; izvor <otisak>` i upisuje ceo marker bez lokalne putanje sveske u namensku granu (npr. dokazi/<otisak>.json). U CI-ju informativni job `excel-dokaz` računa `vba_gate.py --hash` (radi na Linux-u, 0,1 s) i traži zapis za taj otisak; ishod je jasno imenovan kao dokaz sa druge mašine, ne kao lokalni --require-green. Pošto se traži po otisku src-vba a ne po SHA commit-a, docs-only commit i rebase bez izmene izvora nasleđuju dokaz sami.
  DOBITAK: Zamenjuje ručno prepisivanje rezultata u docs i commit poruke (48 NEVERIFIKOVANO commit-a i 111 prepisanih linija mesečno; STANJE_REFAKTORA diran u 51% commit-a). Review koji čita repo vidi da li je BAŠ taj izvor prošao, bez rečenice u PR-u koja zastari posle rebase-a. Preduslov za svaku automatizaciju Excel prolaza.
  RIZIK: Status postavlja operaterov token -- isti nivo poverenja kao današnji docs commit, ali ga je lakše falsifikovati slučajno (pogrešan klon); zato objava ide samo iz markera koji je run_vba upisao fail-closed. Lokalna putanja sveske nosi korisničko ime i ne sme u javni repo. Udaljena provera ne poredi identitet sveske, pa mora nositi svoje, slabije ime.
  PREDUSLOVI: Fine-grained token (statuses: write, contents: write nad jednom granom); odluka gde zapis živi (grana, issue ili release asset).
  DOKAZATI: vba_gate --self-test dobija slučajeve: pao run -> nema objave; --no-import -> nema objave; izmenjen izvor posle objave -> CI job kaže `drugi izvor` i imenuje suite; zapis sa starom verzijom markera -> odbijen. Ručno: docs-only commit posle zelenog -> job i dalje nalazi dokaz.

### CI-9 Noćni Excel prolaz: lokalni zakazani zadatak (preporuka), self-hosted runner samo uz uslove [trud M] [slabi: NE. Grupni dokaz i dalje nosi ime `DOKAZANO (grupno)`; noćni prolaz ne zamenjuje ručni Compile ni pojedinačni dokaz novih sabotaža u rezu, nego dodaje pokrivenost koje danas nema.]
Varijanta B (preporuka): Task Scheduler zadatak `Run only when user is logged on` (interaktivna sesija, ne servis) u ZASEBNOM klonu (dokaz.py kvari src-vba na disku): fetch + checkout origin/main -> vba_check -> regeneracija fixture-a ako potpis ne odgovara -> `run_vba.py` (FULL, piše marker) -> `dokaz.py --grupe 6` ili rotirajući isečak kataloga -> objava kroz CI-8. Mašina nema nijedan ulazni kanal; izlazi samo `gh` poziv. Varijanta A (self-hosted runner): daje dugme Run workflow, status i artifact bez dodatnog koda, ali traži isto (interaktivna sesija -- runner kao proces u prijavljenoj sesiji, ne kao servis; Excel licenca na toj mašini; Trust access to VBOM) i izlaže razvojnu mašinu: repo je JAVAN, a PR iz forka može da cilja `runs-on: self-hosted`. A dolazi u obzir samo ako repo postane privatan ili postoji namenska VM bez klijentskih podataka (tada dodatna Windows i Office licenca i pitanje unattended licenciranja). Hosted Windows runner nije opcija jer nema Office.
  DOBITAK: Regresija na main-u se vidi za najviše 24 h umesto pri sledećem ručnom FULL-u. Grupni dokaz celog kataloga (procena 7-8 h) može da ide svake noći ili vikendom umesto da troši operaterov dan; sidra koja više ne obaraju svoj test izlaze sama. Postojeće već podržava: marker sa dva otiska, fail-closed upis, last_run.json, potpis fixture-a, zelena baza i vraćanje izvora u dokaz.py, `--plan` bez Excela.
  RIZIK: Mašina mora biti upaljena i prijavljena; ponašanje pod zaključanim ekranom nije provereno (SendKeys/AppActivate za compile verovatno ne rade, suite bi trebalo da rade). Viseći Excel ili dijalog: omotač mora da ubije svoj Excel PID i da ima ukupni timeout. Noćni zadatak ne sme da deli klon ni fixture sa dnevnim radom. Pojedinačni dokaz celog kataloga (40-74 h) ne staje u noć -- treba rotacija ili izbor testova (druga oblast).
  PREDUSLOVI: CI-8; trajan izveštaj dokaz.py (JSON po sabotaži + nastavak posle prekida); donor sveska za fixture na toj mašini; odluka o rotaciji kataloga.
  DOKAZATI: (1) Test grana sa namerno pokvarenim testom -> jutarnji status `excel/full` = failure sa imenom suite. (2) Zelena noć -> `vba_gate --require-green` RC=0 u noćnom klonu i CI job nalazi zapis za isti otisak. (3) Ubijen proces usred dokaza -> `git status` u noćnom klonu čist, sledeća noć nastavlja. (4) Tri uzastopne noći bez ručne intervencije.

### CI-10 release.sh: kapija pre bump-a i ukidanje .ps1 dvojnika [trud S] [slabi: NE -- uvodi kapiju tamo gde je nema.]
Pre koraka 2 (bump) skripta proverava: vba_check čist; poslednji `static` run nad HEAD-om main-a je success (gh run list --json conclusion,headSha); `vba_gate.py --require-green --require-compile` RC=0. Izlaz mimo kapije samo izričitim prekidačem koji ispiše šta nije dokazano. tools/release.ps1 i tools/stamp-build.ps1 se brišu (ljuska je po CLAUDE.md samo Git Bash, a .ps1 kopija je već izgubila korake AssertBlankBuild i PublishReleaseToDrive) ili dobijaju CI proveru pariteta. Vraćanje modBuildInfo.bas postaje jedna komanda umesto podsetnika.
  DOBITAK: Tag se više ne može postaviti nad izvorom za koji ne postoji dokaz -- danas može, skripta ne zove nijednu proveru. Dve skripte manje za održavanje. Compile potvrda (--mark-compile) dobija mesto gde se stvarno traži.
  RIZIK: Bump APP_VERSION menja modConfig.bas, pa marker posle bump-a po konstrukciji ne važi: zato provera ide PRE bump-a, a razlika od jedne konstante se prihvata i ispisuje (ili se FULL ponovi nad release commit-om noću). Bez markera na build mašini release staje -- to je namera, ali traži naviku.
  PREDUSLOVI: gh CLI na build mašini (test_grana.sh ga već koristi); CI-8 nije neophodan ali ga dopunjuje.
  DOKAZATI: Suvo pokretanje u klonu bez markera -> staje i imenuje suite koja nedostaje; sa markerom nad drugim izvorom -> `DRUGI izvor`; sa važećim markerom i potvrđenim compile-om -> prolazi do bump-a. bash -n već pokriva sintaksu u CI-ju.

### CI-11 Skratiti sam vba_check u CI-ju: predčišćenje tela u ZAKLONJENO proveri i jedan prolaz who_writes [trud S] [slabi: NE]
check_scalar_call zove _clean(line) za svaki red tela puta svako skalarno ime (tools/vba_check.py:964-968): 1,93 M poziva nad 233.491 linija. Telo se očisti jednom po proceduri. who_writes: --check i --check-ownership u jednom procesu (oba koriste isti scan(); danas zajedno zadate zastavice tiho izvrše samo ownership). Zaseban korak `vba_gate.py --popis` u static.yml može da ode jer isti popis već vrti vba_check. Napomena: unutrašnjost vba_check-a je tuđa oblast; ovde stoji zbog CI budžeta.
  DOBITAK: Pun vba_check -24% (23,29 -> 17,73 s izmereno memoizacijom u memoriji, isti verdikt); CI korak 22-30 s -> oko 17-23 s. who_writes -1,5 s. Popis -1 s.
  RIZIK: Nizak; izlaz mora ostati identičan. Uklanjanje zasebnog popis koraka smanjuje vidljivost u Actions UI (nalaz bi se video u koraku vba_check).
  PREDUSLOVI: Nema.
  DOKAZATI: Diff izlaza vba_check nad celim repoom pre/posle; ZAKLONJENO_CASES u --self-test; ubačen pravi ZAKLONJENO slučaj u kopiju modula -> isti nalaz na istoj liniji; vreme pre/posle. Za who_writes: namerno zastareo WHO_WRITES.md i nov pisac van WRITE_OWNERSHIP.json u istom pozivu -> oba nalaza.


## NEPOZNANICE
[
 "Da li je `static` obavezan status check i da li postoji branch protection na main-u -- nije čitljivo alatima koje sam koristio; utiče na podelu job-ova (ime check-a) i na to da li crven CI uopšte blokira merge.",
 "Uzrok setup-node repova od 45-64 s nije potvrđen iz loga tog koraka (preuzet je samo kraj loga); pretpostavka je da Node 20 više nije u tool cache-u runner-a.",
 "Stvarno trajanje FULL prolaza i dokaza na razvojnoj mašini: svi brojevi (68-91 s, ~150 s, ~18 s režije, 20-60 min, 3,5 min po sabotaži) su iz dokumentacije; procene 7-8 h i 40-74 h su računica, ne merenje.",
 "Da li run_vba i dokaz rade pod zaključanim ekranom ili posle RDP diskonekta (EnumWindows watchdog, SendKeys/AppActivate) -- ne može se proveriti bez Windows mašine.",
 "Da li razvojna mašina može da ostane upaljena i prijavljena noću, i da li uslovi licence za Excel dozvoljavaju zakazano (unattended) pokretanje -- pravno/organizaciono pitanje.",
 "Stvarni uticaj nalaza o sw.js u pregledaču (dva fajla van ASSETS, izmena bez promene CACHE_NAME) nije proveren; ne znam koliko klijenata ima instaliran keš AgriX-v25 ni da li je CACHE_NAME dizan pre početka vidljive (plitke) istorije.",
 "Kako se GAS kod isporučuje (u repou nema clasp konfiguracije): ne postoji provera da je ono što radi u Apps Script-u isto što i gas/Code.gs u gitu.",
 "Da li se GitHub Pages objavljuje iz korena main grane pretpostavljeno je iz CNAME fajla i workflow-a pages-build-deployment; podešavanje nije pročitano.",
 "Setup-AgriX.ps1 -SelfTest nije meren lokalno (pwsh ne postoji u ovoj sesiji); raspon 1-11 s je iz CI-ja.",
 "check-banka-eh.py i check-sef-asserts.py nemaju --self-test, pa ne znam da li danas išta mere -- zeleni su, ali nikad nisu ovde pokazani crveni.",
 "Za tools/migracija ne postoji katalog mutacija; koliko od 600 testova bi palo na ciljanoj mutaciji nije mereno (alat tipa mutmut nije instaliran niti pušten).",
 "Da li je main i dalje na 1fa4937: poslednji push run na main-u u API-ju je #1227 od 08.11.2026, a PR-ovi #469, #474 i #475 su imali run-ove 11.-15.11. i deluju otvoreno; stanje posle njih nisam analizirao.",
 "Koliko operatera košta jedan ručni prenos dokaza (fetch grane, ImportAllVBA, Compile, prolaz, prepis rezultata) u minutima -- izmeren je samo broj takvih commit-a, ne vreme."
]
