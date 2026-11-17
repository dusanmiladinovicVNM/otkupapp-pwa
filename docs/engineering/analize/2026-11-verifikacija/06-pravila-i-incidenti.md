# Pravila verifikacije i njihova cena (CLAUDE.md §2/§5/§6, .claude/rules/testovi.md, git-i-release.md, skill pre-flight, EXCEL_TEST_HARNESS, postmortem 2026-08, STANJE_REFAKTORA, RELEASE_PROCEDURE) -- stablo A (origin/main @1fa4937)

## SAZETAK
Sve putanje su u stablu A = <scratch>/main-wt (pisano kao A/...); nijedan fajl u stablima A i B nije menjan, pomocne skripte su u <scratch>/pravila-analiza/. Excel nije dostupan, pa su sva trajanja Excel prolaza izvedena iz zapisa u STANJE_REFAKTORA i oznacena kao neizmerena.

1) Najskuplje pravilo je i najmanje izvrsivo. CLAUDE.md §5 tvrdi da dokaz.py po sabotazi vrti RunAllTests i BFP, oko 3,5 min po stavci, i da pun katalog traje "sat i vise". Kod pusta jednu suitu po prolazu. `dokaz.py --plan` daje 1265 prolaza pojedinacno, odnosno 215 uz --grupe 6. Po vremenima iz STANJE to je reda 60-95 h pojedinacno i 11-16 h grupno.

2) Pun katalog je poslednji put pusten 28.08.2026 nad 261 sabotazom; od tada je narastao na 1265. Okidac "pred release" ne okida: poslednje datirano izdanje je 03.09, a v2.93.0 stoji "u pripremi". To je isti mehanizam koji postmortem §10 opisuje kao uzrok deset mrtvih sidara: skupo pravilo se tiho svede na podskup.

3) Praksa se odvojila od pisanih pravila. "Jeftine kapije", "skupe kapije posle reviewer GO", "izvrsni acceptance", "crveni prolaz nad starim kodom", `--grupe` i `--require-green` ne postoje u CLAUDE.md ni u .claude/rules. Jedna runda acceptance-a prijavljuje 7-20 podizanja Excela (13 pregledanih stavki, 261-323).

4) Nadjeno je vise nesaglasnosti dokument-kod, popisanih u cinjenicama sa prefiksom NESAGLASNOST. Tri diraju tvrdnje sistema:
- hook ne pokrece nijednu proveru kad se menja tools/sabotaza.py (izmereno: rc=0 za 0,056 s), iako dokumentacija i komentar u vba_check.py tvrde suprotno;
- `sabotaza.py --self-test` (43 slucaja) ne vrti ni CI ni vba_check;
- release.sh taguje i gura tag pre Compile-a i bez ijedne kapije, a bump APP_VERSION u modConfig.bas menja otisak izvora, pa ponistava marker dobijen pre release-a.

5) Pravila se ucitavaju skupo, a malo ih je politika. Svaka VBA izmena ucita najmanje 270 linija; 40 od 193 fajla ucita najmanje 815 (testovi.md ima 544 linije, od kojih je "sta se pusta kada" oko 65). Tabela nivoa postoji u dva fajla i vec se razisla (testovi.md koristi backslash koji CLAUDE.md zabranjuje). Masinski izvor postoji samo za suite (SUITES u run_vba.py) i za CI (static.yml); lokalne komande koja pusta isti skup kao CI nema.

6) Predlog je slojevita politika (po Edit-u, pre commit-a, po rezu, po PR-u, nocno, pred release) sa jednim manifestom kapija. Kljucni potez je da se periodicni pun dokaz veze za kalendar umesto za release. Dva predloga slabe po jednu tvrdnju i tako su oznacena (P5, P10); P7 i P8 nose uslovno, usko suzenje; ostali ne slabe nista ili jacaju.

## CINJENICE
- [M] KATALOG T1 -- 'Baza mora biti zelena pre prve mutacije' (inace se meri 'posle mutacije postoji crveno', ne 'mutacija je izazvala crveno'). Incident: postmortem §10, 'Cetiri rupe' br. 2. Sprovodi: dokaz.py `_baza_zelena`, rc=2 bez ijedne mutacije.
    DOKAZ: A/tools/dokaz.py:17-21, 152-167, 465-473; A/docs/engineering/postmortems/2026-08-verifikacija.md:216-222
- [M] KATALOG T2 -- 'Pala je bas NJENA tvrdnja, ne samo njen test' (AssertEq puca na prvom padu, zamka 6). Incidenti: postmortem §10 br. 6 (7 od 13 neuskladjenih), §12 (119 zastarelih tvrdnji), §15 (5 sabotaza obaralo tudju tvrdnju). Sprovodi: dokaz.py `_oceni` ('NE OBARA NISTA' / 'NE OBARA SVOJ TEST' / 'PALA DRUGA TVRDNJA'), staticki sabotaza.py `_nalazi`.
    DOKAZ: A/tools/dokaz.py:29-35, 276-305; postmortem:261-277, 391-422, 966-1058
- [M] KATALOG T3 -- 'Izvor se vraca i kad run pukne; potpis celog src-vba pre i posle je identican, inace STOP'. Incidenti: postmortem §10 br. 3 (Ctrl+C ili timeout ostavlja pokvaren izvor) i §15 (ubijen proces ostavio sabotiran modOtkupUI.bas; dijagnoza 'sidro ZASTARELO' bila pogresna). Sprovodi: dokaz.py `finally` + `_otisak` + REVERT-FAIL; sabotaza.py nalaz 'izvor ZATECEN SABOTIRAN' kroz vba_check.
    DOKAZ: A/tools/dokaz.py:23-27, 83-90, 355-365, 530; A/tools/sabotaza.py:12491-12498; postmortem:224-228, 1069-1097
- [M] KATALOG T4 -- 'Zelena provera koja nikad nije pokazana crvena ne dokazuje da ista meri' (dvosmerni dokaz obavezan za nov ili izmenjen test ili checker i za kriticnu invarijantu; nije obavezan za obicnu funkcionalnu izmenu). Incident: postmortem §4 (PR #181, cetiri puta zeleno-ali-nedokazano). Sprovodi: katalog u sabotaza.py + dokaz.py za VBA testove; `--self-test` koraci u CI za checkere.
    DOKAZ: A/CLAUDE.md:108-111; A/.claude/rules/testovi.md:436-447; postmortem:62-68; A/.github/workflows/static.yml:85-243
- [M] KATALOG T5 -- 'Self-test checkera ide KROZ produkcionu putanju (check_file / main / ista funkcija koju zove CLI), u oba smera, i vrti se u CI'. Incidenti: postmortem §12 (self-test sklapao svoje 'telo' i zaobisao `_tela_testova`), §15 (fixture biran rukom, guard zelen nad stanjem koje --vrati odbija). Sprovodi: 12 self-test koraka u static.yml. RUPA: `sabotaza.py --self-test` nije medju njima.
    DOKAZ: A/.claude/rules/testovi.md:173-181; postmortem:485-502, 1099-1144; `grep -n -i 'sabotaz\|dokaz' A/.github/workflows/static.yml` -> nijedan poziv sabotaza.py
- [M] KATALOG T6 -- 'Nema fajla = pad': verdikt dolazi iz last_run*.txt pored sveske, po slotu suite, ne iz toga da li Run() pukne. Incidenti: postmortem §10 br. 4-5 (banka se nije mogla izmeriti; dve suite prepisivale rezultat jedna drugoj), §14 (sabotaza koja se ne kompajlira izgledala kao 'NE OBARA NISTA'). Sprovodi: run_vba.py `_read_test_results`.
    DOKAZ: A/tools/run_vba.py:116-135; A/.claude/rules/testovi.md:194-195; postmortem:230-259
- [M] KATALOG T7 -- 'Suite koja se nije pokrenula nije OK' (tihi Exit Sub podize 'suite NIJE pokrenut:'). Incident: postmortem §5 (cetiri putanje laznog zelenog). Sprovodi: VBA suite (modTestPalete 3 mesta, modAgrohemijaTests 1). Poznata otvorena rupa po dokumentu: RunBankaImportTestSuite rani Exit Sub.
    DOKAZ: postmortem:70-86; A/docs/EXCEL_TEST_HARNESS.md:166-170; `grep -c 'suite NIJE pokrenut' A/src-vba/*.bas` -> modTestPalete.bas:3, modAgrohemijaTests.bas:1
- [M] KATALOG T8 -- 'BLIND nije dokaz' (gate:False znaci prosla bez greske, ne da su sve provere prosle; nova suite mora biti gate; marker je ne priznaje). Sprovodi: SUITES u run_vba.py (18 suita, 12 default+gate, 3 blind), vba_gate `potrebne_suite` i pravila markera. Dopuna T17: nijedna suite van kapije bez zapisa (`vba_gate --popis`, registar SUITE_VAN_KAPIJA, pet suita).
    DOKAZ: A/tools/run_vba.py:60-92, 970, 1117-1119; A/tools/vba_gate.py:80-82, 178, 633-641; A/docs/EXCEL_TEST_HARNESS.md:185-206
- [M] KATALOG T9 -- 'Sve sto nije eksplicitno OK je pad; NEJASNO se prijavljuje kao nejasno'. Incident: tri-cetiri lazna zelena headless compile-a (stanje menija, lenja kompilacija probe modula, Execute bez vidljivog VBE). Sprovodi: run_vba za --compile-only (rc=2). Kad idu suite, NEJASNO ne obara run -- v. nesaglasnost sa EXCEL_TEST_HARNESS.
    DOKAZ: A/docs/EXCEL_TEST_HARNESS.md:99-145; A/tools/run_vba.py:906-918; A/CLAUDE.md:118-119
- [M] KATALOG T10 -- 'Zeleno vazi samo za BAS OVAJ izvor pod OVIM test sistemom' (marker po suiti sa dva otiska: izvor i ugovor = izvor + runner + fixture + kapija + golden; kontekst se snima pre run-a; upis je fail-closed; ne upisuje pao run, --no-import ni BLIND). Incident: review #400 P3; rebase ili amend posle zelenog. Sprovodi: vba_gate.py `snimi_kontekst` / `zabelezi_prolaz` / `--require-green`. Otvoreno: KI-036.
    DOKAZ: A/tools/vba_gate.py:41-60, 155, 563, 644-660; A/docs/EXCEL_TEST_HARNESS.md:284-360; A/docs/KNOWN_ISSUES.md:62
- [M] KATALOG T11 -- 'Fixture odgovara grani: bez vazeceg potpisa run ne prolazi' (jedini izlaz je svestan --ignore-fixture-sig). Incident: pola sata trijaze, cetiri crvena testa nad ispravnim kodom posle prelaska grane. Sprovodi: make_fixture .sig + run_vba `proveri_fixture_potpis` pre podizanja Excela.
    DOKAZ: A/.claude/rules/testovi.md:196-203; A/docs/EXCEL_TEST_HARNESS.md:609-645; A/tools/run_vba.py:724-772, 832
- [M] KATALOG T12/T13 -- 'Sema se priprema iz svezeg koda pre suita (EnsureRuntimeSchema posle importa); pala priprema obara run' i 'fixture ne nosi VBA kod donora'. Incidenti: postmortem §1 (147/310 padova citano kao nalaz o proizvodu), §2 ('Cannot run the macro' zbog zaostalog duplikata medju 131 modulom donora). Sprovodi: run_vba.py i make_fixture.py `strip_vba`.
    DOKAZ: postmortem:11-41; A/tools/run_vba.py:915-938; A/tools/make_fixture.py:2820, 3034
- [M] KATALOG T14 -- 'Jedan uzrok = jedan pad': ciscenje ide iz EH grane, Err se cita pre ciscenja, a seam sme samo da pokvari, zabrani ili posmatra i gasi ga ResetSeamova. Incident: postmortem §3 (jedna sabotaza obarala dva testa). Sprovodi: modTest `CleanupPosleTesta` / `ResetSeamova`; alata koji to proverava nema, cuvaju ga pravilo i testovi.
    DOKAZ: postmortem:43-60; A/src-vba/modTest.bas:692-698, 717, 786; A/.claude/rules/testovi.md:222-230
- [M] KATALOG T15 -- 'Katalog sabotaza je ziv bez Excela': sidro tacno jednom, test postoji, tvrdnja je poruka assert-a BAS tog testa, zamke 4/5/7/8; priznat nalaz nosi tacan prefiks; mrtav upis je nalaz. Incidenti: postmortem §10 (10 od 222 mrtvih sidara), §12 (119 zastarelih tvrdnji), §14 (zamena koja se ne kompajlira). Sprovodi: `sabotaza.py --proveri-sidra` kroz vba_check (hook i CI); 9 upisa u POZNATI_NALAZI, 2 u POZNATI_NALAZI_DOKAZ.
    DOKAZ: A/tools/sabotaza.py:13213, 13357-13390; A/tools/vba_check.py:4022-4034, 4115; postmortem:172-203, 594-600; `proveri_sidra(True)` -> rc 0 za 2,86 s nad 1265 unosa
- [M] KATALOG T16 -- 'Tri registra testova i stvarna tela su saglasni' (RunOne / TestName / InvokeTest; rupa u numeraciji, dupli indeks, neregistrovano telo). Incident: razlaz pri rebase-u; KI-084 (pogresno upisan kao da kapije nema). Sprovodi: vba_check REGISTAR (`check_test_registry`).
    DOKAZ: A/docs/EXCEL_TEST_HARNESS.md:379-406; A/tools/vba_check.py:2000-2038; A/docs/KNOWN_ISSUES.md:110
- [M] KATALOG T18 -- 'Grupni dokaz tvrdi manje i tako se prijavljuje' (DOKAZANO (grupno)); clan koji u grupi ne obori svoju tvrdnju meri se sam; cetiri pravila grupisanja (ista suita, razlicit test, kljuc i procedura). Sprovodi: dokaz.py `naprav_grupe` / `_prolaz` + `_self_test` kroz vba_check. Pravilo o obimu: u rezu grupno, pred release pojedinacno.
    DOKAZ: A/tools/dokaz.py:37-42, 206-273, 575-580; A/docs/EXCEL_TEST_HARNESS.md:749-758; A/tools/vba_check.py:4046-4057
- [M] KATALOG T19/T20/T21 -- (T19) bez Excela se izmena ponasanja prijavljuje kao NEVERIFIKOVANA, nikad zelena; (T20) Compile je rucna kapija vezana za otisak izvora (--mark-compile); (T21) bug: prvo reprodukcija crveno po imenu, pa zakrpa, a nereprodukovano se ne krpi. Incidenti: postmortem §13 (m_BlokoviOk: vba_check cist, run_vba visi), testovi.md (12 ZAKLONJENO poziva od v6-ui-119 do 141; CLAN_FORME uz zelen FULL), postmortem §11 i §16. T19 i T21 ne sprovodi nijedan alat; T20 sprovodi vba_gate.
    DOKAZ: A/CLAUDE.md:35-43, 96-98, 103-104; A/.claude/rules/testovi.md:123-136, 155-160; postmortem:667-685, 308-328, 1194-1205; `grep -c NEVERIFIKOVANO A/docs/STANJE_REFAKTORA.md` -> 19
- [M] KATALOG T22/T23/T24/T25 -- (T22) FULL nije u hook-u, cena provere je proporcionalna riziku (postmortem §7); (T23) lazna uzbuna u hook-u je gora od propusta, pa su provere namerno uske (406 laznih nalaza pri sirenju ARNOST-a; §13: 54 lazne uzbune); (T24) izvestaj ne sme da pukne niti da sam sebi protivreci (§10: ChrW(183) na cp1252); (T25) infrastruktura se smoke-testira u stvarnom okruzenju (§8 zarez u settings.json, §9 python3 Store alias). Sprovode: settings.json (samo PostToolUse), run_vba errors=replace i slotovi, CI koraci 'settings JSON je validan' i 'hook skripte se parsiraju'.
    DOKAZ: postmortem:95-137, 160-170, 687-708; A/.claude/rules/testovi.md:105-109; A/.github/workflows/static.yml:36-73; A/.claude/hooks/vba-check.sh:12-19
- [M] KATALOG T26 -- 'Pun prolaz dokaza je PERIODICAN, ne jednokratan': staticka provera ne zna sta sabotaza stvarno obara (8 od 251, oko 3%), a ispravna izmena (nova centralna kapija) zastareva tudje dokaze. Incident: postmortem §12 i §17. Sprovodi: nista automatski; pravilo glasi 'pun katalog pred release'.
    DOKAZ: postmortem:621-663, 1232-1294; A/CLAUDE.md:112-117
- [M] NESAGLASNOST 1 -- CLAUDE.md §5 kaze: '`tools/dokaz.py` po svakoj sabotazi vrti `RunAllTests` i BFP -- oko 3,5 minuta po stavci, pa pun katalog traje sat i vise'. Kod pusta JEDNU suitu po prolazu, onu u cijem modulu test zivi. Obe suite idu samo u baznom prolazu, jednom po pozivu. Racunica je i sama sa sobom u neskladu: 1265 x 3,5 min = 73,8 h.
    DOKAZ: A/CLAUDE.md:112-114; A/tools/dokaz.py:92-104, 349, 466; `python3 tools/dokaz.py --plan` -> RunAllTests 801, BFP 456, Banka 8 prolaza
- [M] NESAGLASNOST 2 -- `dokaz.py --grupe 6 --plan` daje 215 prolaza: RunAllTests 134, BFP 78, Banka 3. Pojedinacno je 1265. Pravilo 'pred release pojedinacno' zato znaci 1265 prolaza suite.
    DOKAZ: komanda `python3 tools/dokaz.py --grupe 6 --plan` -> 'sabotaza 1265, prolaza 215 (najvise 6 po grupi) -- 5.9x manje'; A/docs/EXCEL_TEST_HARNESS.md:752; A/docs/STANJE_REFAKTORA.md:11025-11027
- [M] NESAGLASNOST 3 -- Docstring dokaz.py tvrdi da 'CLAUDE.md paragraf 5 trazi da se posle izmene pusti ceo dvosmerni dokaz'. CLAUDE.md §5 danas trazi suprotno: 'u rezu samo NOVE sabotaze, pun katalog pred release'. Komanda reza u CLAUDE.md je bez --grupe (pojedinacno), a EXCEL_TEST_HARNESS i praksa kazu 'u rezu grupno'. `--grupe`, `--require-green` i `--proveri-sidra` imaju 0 pojava u CLAUDE.md i testovi.md.
    DOKAZ: A/tools/dokaz.py:4-5; A/CLAUDE.md:112-117; A/docs/EXCEL_TEST_HARNESS.md:752; grep brojanje: --grupe CLAUDE 0 / testovi 0 / HARNESS 5 / STANJE 76; --require-green 0 / 0 / 2 / 57
- [M] NESAGLASNOST 4 -- Hook ne pokriva alatne fajlove, a dokumentacija i komentar u kodu tvrde da pokriva ('Hook se zove i sa putanjom koja nije VBA fajl -- na primer bas tools/sabotaza.py, gde se greska u katalogu i pravi'). Skripta hook-a izlazi sa 0 za sve sto nije .bas/.cls/.frm/.doccls, pre poziva vba_check.
    DOKAZ: A/.claude/hooks/vba-check.sh:52-55; A/tools/vba_check.py:4109-4114; A/docs/EXCEL_TEST_HARNESS.md:765-769; izmereno: payload sa tools/sabotaza.py -> rc=0 za 0,056 s (isto za tools/dokaz.py); nad src-vba/modParse.bas -> rc=0 za 6,3 s
- [M] NESAGLASNOST 5 -- `sabotaza.py --self-test` (dokaz da pravila --proveri-sidra grizu; 43 slucaja, 0,54 s) ne pokrece ni CI ni vba_check. vba_check zove samo `proveri_sidra`, a od alata oko dokaza samo `dokaz._self_test`. Postmortem §12 obrazlaze smestanje pravila u `_nalazi` upravo time sto joj '--self-test podmece izmisljene unose'.
    DOKAZ: A/tools/vba_check.py:4034, 4057; A/.github/workflows/static.yml (nijedan poziv sabotaza.py); komanda `python3 tools/sabotaza.py --self-test` -> 'self-test: cisto (43 slucajeva)', 0,54 s; postmortem:594-600
- [M] NESAGLASNOST 6 -- testovi.md §1 je zastareo u odnosu na CI: 'Sest od tih alata ima i --self-test' i '`gen_schema_module.py` nema self-test, i to je poznata rupa'. static.yml ima 12 self-test koraka, ukljucujuci gen_schema_module --self-test (od 13.09.2026). CI spisak u testovi.md navodi 'tri kapije nad self-update motorom', a CI ima i cetvrtu (vba_import_marker_gate), plus vba_gate --popis, migraciju i JS harness.
    DOKAZ: A/.claude/rules/testovi.md:63-75; A/.github/workflows/static.yml:85-243 (posebno 135-139, 165-171); A/tools/gen_schema_module.py:875, 1005
- [M] NESAGLASNOST 7 -- testovi.md daje komande sa backslash-om (`python tools\vba_check.py`, `python tools\run_vba.py --suite <ime>`, `python tools\run_vba.py`, `python tools\vba_check.py --self-test`), a CLAUDE.md to izricito zabranjuje ('nikad `\` u putanji... Backslash u Git Bash-u nestaje bez greske'). Ista tabela nivoa postoji u oba fajla i vec se razisla. git-i-release.md koristi `python3 tools/vba_check.py`, a postmortem §9 belezi da je python3 na Windows-u Store alias (rc=49).
    DOKAZ: A/.claude/rules/testovi.md:55-57, 173; A/CLAUDE.md:11-16, 90-94; A/.claude/rules/git-i-release.md:28-29; postmortem:119-123
- [M] NESAGLASNOST 8 -- Brzina jeftine polovine opisana je kao 'u istoj sekundi' (hook), 'u milisekundama' (EXCEL_TEST_HARNESS), '0,14 s, tiho kad je cisto' i 'za sekundu' (postmortem), 'Traje sekundu' (vba_check.py). Izmereno ovde: hook 6,3 s po VBA fajlu (ranije 9,0-11,9 s), --proveri-sidra 2,86-3,5 s, pun vba_check 31 s. Katalog je rastao 222 -> 1265, a brojke u dokumentima nisu.
    DOKAZ: A/.claude/hooks/vba-check.sh:5-7; A/docs/EXCEL_TEST_HARNESS.md:91; postmortem:174-175, 214; A/tools/vba_check.py:4020-4021; merenja u ovoj sesiji i zadata merenja
- [M] NESAGLASNOST 9 -- EXCEL_TEST_HARNESS kaze da je kod COMPILE NEJASNO 'izlazni kod i tada 2' i da je 'sve sto nije eksplicitno OK pad (izlazni kod 2)'. run_vba.py to radi samo uz --compile-only; kad idu suite, 'NEJASNO vise ne obara run' (testovi.md §3 to kaze tacno). Compile proba se pri tom radi na SVAKOM pozivu run_vba, a STANJE ishod opisuje kao 'automatski NEJASNO, kao i uvek'.
    DOKAZ: A/docs/EXCEL_TEST_HARNESS.md:140-145; A/tools/run_vba.py:899-918; A/.claude/rules/testovi.md:204-205; A/docs/STANJE_REFAKTORA.md:9714 ('automatski NEJASNO, kao i uvek'); grep NEJASNO u STANJE = 35
- [M] NESAGLASNOST 10 -- 'Pun dokaz ide pred release, uz FULL prolaz' (CLAUDE.md §5) ne postoji u izvrsnoj proceduri release-a. RELEASE_PROCEDURE.md i tools/release.sh imaju 0 pojava dokaz / run_vba / vba_check / vba_gate / require-green. release.sh pravi i gura tag PRE rucnog Compile-a.
    DOKAZ: `grep -n -i 'dokaz\|run_vba\|vba_check\|vba_gate\|require-green' A/docs/RELEASE_PROCEDURE.md A/tools/release.sh A/tools/release.ps1` -> prazno; A/tools/release.sh:44-50 (tag + push), 58-60 (Compile rucno posle)
- [N] NESAGLASNOST 11 -- Sam release ponistava dokaz dobijen pre njega. release.sh menja APP_VERSION u src-vba/modConfig.bas, a otisak `izvor` izuzima samo modBuildInfo.bas. Posle bump-a `--require-green --require-compile` vidi drugi izvor za tagovani commit. Zakljucak je izveden citanjem koda; nije izvrsen.
    DOKAZ: A/tools/release.sh:30-39; A/tools/vba_gate.py:155 (`IZUZET_IZ_OTISKA = ("modBuildInfo.bas",)`); A/src-vba/modConfig.bas:13
- [M] NESAGLASNOST 12 -- Postmortem §7 i §8 navode kao vazece pravilo da 'hook pusta brzi set + zig u .git/vba-test-stamp' i 'JSON gate na dva ulaza -- PostToolUse i Stop... Detalji u testovi.md'. U settings.json postoji samo PostToolUse; vba-test.sh i Stop hook ne postoje u repou, a testovi.md nema odeljak o JSON gate-u.
    DOKAZ: postmortem:102-104, 116-117, 129; A/.claude/settings.json (hooks: samo PostToolUse); `ls A/.claude/hooks` -> samo vba-check.sh
- [M] NESAGLASNOST 13 -- CLAUDE.md §5 kaze da FAST 'ide i automatski kroz PostToolUse hook' sa komandom `python tools/vba_check.py`. Hook zove `vba_check.py --hook <fajl>`, sto je podskup: DUPLIKAT ('Ambiguous name') se racuna samo bez putanja. Hook zato ne zamenjuje pun vba_check pre commit-a.
    DOKAZ: A/CLAUDE.md:92; A/.claude/hooks/vba-check.sh:59; A/tools/vba_check.py:4144-4156 ('Duplikate trazimo samo kad se gleda ceo src-vba')
- [M] NESAGLASNOST 14 -- testovi.md se ne ucitava kad se menjaju tools/dokaz.py, tools/vba_gate.py i tools/vba_import_marker_gate.py (nisu u `paths:`), iako su to alati koji odlucuju sta je dokaz. git-i-release.md ne aktivira nijedan od 193 VBA fajla, pa se pravilo 'na kraju SVAKE izmene koda' odatle ne ucitava u obicnoj sesiji (sustinu ponavlja CLAUDE.md §6).
    DOKAZ: A/.claude/rules/testovi.md:2-25; A/.claude/rules/git-i-release.md:2-8; skripta nad `paths:` -> git-i-release.md aktivira 0 VBA fajlova
- [M] NESAGLASNOST 15 -- STANJE_REFAKTORA se predstavlja kao 'Ulaz za svaku novu sesiju. Kratko i azurno', a ima 11.059 linija. Zaglavlje kaze 'Azurirano: 04.11.2026 (stavka 322)', poslednja stavka je 323 (06.11), a main je na 08.11 (PR #471-#473; ishodi su u KNOWN_ISSUES KI-091 i KI-094, ne u STANJE). APP_VERSION u izvoru je 2.28.4, dok RELEASE_NOTES vodi v2.92.0 i v2.93.0 'u pripremi'.
    DOKAZ: A/docs/STANJE_REFAKTORA.md:3-7, 10954; `wc -l` -> 11059; A/docs/KNOWN_ISSUES.md:117, 120; A/src-vba/modConfig.bas:13; A/docs/RELEASE_NOTES.md:7098, 7151
- [M] PRAKSA 1 -- Pojmovi koji vode svakodnevni rad ne postoje u pravilima. 'jeftine kapije' (59 pojava u STANJE), 'skupe kapije' (14; 'Skupe kapije cekaju reviewer GO'), 'izvrsni acceptance' (11) i 'crveni prolaz / crveno nad starim kodom' (35) imaju 0 pojava u CLAUDE.md, .claude/rules, pre-flight i EXCEL_TEST_HARNESS. Slojevita politika se dakle vec primenjuje, ali nije zapisana.
    DOKAZ: `grep -c` nad A/docs/STANJE_REFAKTORA.md; `grep -n -i 'skupe kapije\|jeftine kapije\|izvrsni acceptance\|crveni prolaz'` nad pravilima -> prazno; STANJE:922, 990
- [M] PRAKSA 2 -- Uzorak poslednjih 20 stavki (304-323): 5 prijavljuje izvrsene Excel kapije (309, 316, 320, 322, 323), 5 su merge zapisi koji ponavljaju iste brojke kao 'Osnova za main' (311, 318, 321), a 3 su 'kod lokalno -- NEVERIFIKOVANO'. Prolazi po stavci, izbrojani iz teksta: 309 = 7 podizanja Excela (Palete, Storno, bazni RunAllTests + BFP, 2 prolaza dokaza, compile-only) za 21 min; 316 = 7-12; 320 = 7-13; 322 = najmanje 24 (11 odlozenih sabotaza + FULL 12 suita + ponavljanja); 323 = FULL 12 suita najmanje dvaput + dokaz 9/9. FULL je u tih 20 pomenut u 4, a izvrsen u 2.
    DOKAZ: A/docs/STANJE_REFAKTORA.md:10728-10753, 10842-10985; skripta <scratch>/pravila-analiza/stavke.py
- [M] PRAKSA 3 -- Siri uzorak od 13 stavki sa izvrsenim kapijama (261, 267, 271, 279, 285, 288, 292, 301, 309, 316, 320, 322, 323): jedna runda acceptance-a = 7-20 podizanja Excela. Sastav: crveni prolaz nad starim kodom 2-3 (uz rucno vracanje 4-6 produkcionih modula), zeleno nad novim 4-8 suita, dokaz = bazni RunAllTests + BFP + 2-7 prolaza, i `run_vba --compile-only --keep` za svesku rucnog Compile-a. Runda se ponavlja posle review kruga (#454: stavke 267 i 271; #459: 285 i 288).
    DOKAZ: A/docs/STANJE_REFAKTORA.md:9097-9136 (17 podizanja), 9277-9324, 9409-9427, 9681-9727, 9947-10000, 10099-10151, 10263-10298, 10514-10552
- [M] PRAKSA 4 -- FULL se pustao daleko cesce nego sto pravilo trazi ('pred release i za rizicne izmene u jezgru'). Od 322 stavke, 59 pominje zelen pun prolaz (12/12): 13 u dekadi 01-10.10, 29 u 11-20.10, 14 u 21-30.10, a samo 3 od 01.11. Za to vreme nije bilo nijednog datiranog izdanja. Pominjanje nije isto sto i izvrsenje, jer merge stavke ponavljaju brojke.
    DOKAZ: skripta <scratch>/pravila-analiza/full.py -> 'stavki: 322 sa FULL/12-12 zelenim: 59 sa dokaz N/N: 78'; A/docs/RELEASE_NOTES.md:7098, 7151
- [M] PRAKSA 5 -- Odlaganje 'dirnutih' sabotaza je vec proizvelo trulez. U stavci 288 operater je, pozivajuci se na CLAUDE.md §5, pustio samo nove sabotaze, a 11 regresionih (pet presidrenih iz 1b i sest s6k4) odlozio 'uz pun katalog pred release'. U zavrsnom krugu (stavka 322) prvi prolaz je bio 9/11: jedna se nije kompajlirala, a jedna je obarala tudji test.
    DOKAZ: A/docs/STANJE_REFAKTORA.md:10109-10121, 10941-10943
- [M] PRAKSA 6 -- Pun katalog nije pusten od 28.08.2026 (260/261). Posle toga se 'pun katalog sabotaza pred release' samo ponavlja kao odlozena obaveza (stavke 288, 294, 296, 321, 322, 323), a katalog je narastao 261 -> 1265 (4,8x). Poslednje datirano izdanje je vba-v2.92.0 (03.09.2026); u avgustu je bilo vise izdanja dnevno (78 unosa ukupno), pa je 'pred release' tada znacilo skoro svaki dan.
    DOKAZ: postmortem:1234-1240; `grep -rn 'crvenih: [0-9]\{3,4\} / sabotaza'` -> samo 251 i 261; A/docs/STANJE_REFAKTORA.md:10115, 10336, 10375, 10934, 10951, 10981; A/docs/RELEASE_NOTES.md (78 zaglavlja '## vba-v')
- [M] PRAKSA 7 -- Pravilo o markeru cini sitnu izmenu skupom, i to se uzima u obzir pri odlukama. 'Cena jednog premestenog reda je dakle pun prolaz + dokaz + drugi operaterov compile'; 'Cena zatvaranja je pun ponovni prolaz (~37 min) i jos jedan rucni compile'; KI-058 (slovna greska) je odlozen 'da #443 ne trazi nov FULL i Compile'; KI-036: posle dokaz.py marker BFP nosi drugi izvor, cena je ponovljen prolaz (~4 min).
    DOKAZ: A/docs/STANJE_REFAKTORA.md:1452-1456, 2484-2485; A/docs/KNOWN_ISSUES.md:84, 62
- [M] OBIM 1 -- Ucitavanje pravila po sesiji. CLAUDE.md (137 linija, 8,6 KB) ide uvek. vba-izvor.md (132 linije) ide za svaki od 193 VBA fajla. testovi.md (544 linije, 31,9 KB, 4534 reci) aktivira 40 od 193 fajla i 23 obrasca, medju njima modTest*, mod*Tests i svi modScr*. Posto izmena ponasanja mora da nosi test, testovi.md se ucitava u prakticno svakoj takvoj sesiji. Minimum je 270 linija / 16,1 KB; sa testovi.md najmanje 815 linija / 48 KB; najvise 979 linija / 55,8 KB (modOtkupUI.bas).
    DOKAZ: `wc -l -w -c` nad A/CLAUDE.md i A/.claude/rules/*.md; skripta nad `paths:` frontmatter-om (ispis: testovi.md 40 fajlova, vba-izvor.md 193, max 979 linija)
- [M] OBIM 2 -- U testovi.md je politika 'sta se pusta kada' oko 65 od 544 linije (12%): §1 = 25, §3 = 23, §6 = 12, §7 = 5. Ostalo: §2 opis vba_check pravila sa istorijom incidenata = 105, §4 katalog seam-ova KI-011/014/015/016 = 208, §5 = 19, §8 i §9 opis statickih kapija = 89. Postmortem je 14.08 nastao da bi istoriju izvukao iz ovog fajla ('arheologija koja ne mora da se ucitava u svakoj sesiji'), ali se istorija vratila.
    DOKAZ: A/.claude/rules/testovi.md:51-75, 77-181, 183-205, 207-414, 436-453, 455-544; postmortem:3-5
- [M] OBIM 3 -- Jednog kratkog, masinski proverljivog izvora 'nivo -> komanda -> kada' nema. Tabela nivoa je rucno prepisana u CLAUDE.md §5 i testovi.md §1 (i razisla se). Masinski izvor postoji za suite (SUITES u run_vba.py, proverava ga vba_gate --popis) i za CI (static.yml, oko 22 koraka kapija). Lokalni spisak 'pre push-a' u STANJE navodi 4 komande. Ista pravila se ponavljaju: Compile kao rucna kapija na 6 mesta (CLAUDE 1, testovi 3, git-i-release 1, HARNESS 7, RELEASE_PROC 3, STANJE 23 pominjanja); pasus 'Grupni dokaz' i pasus 'Marker zelenog' stoje gotovo doslovno i u EXCEL_TEST_HARNESS i u STANJE 'Alati i kapije'.
    DOKAZ: A/CLAUDE.md:90-94; A/.claude/rules/testovi.md:53-57; A/docs/STANJE_REFAKTORA.md:11020-11054; A/docs/EXCEL_TEST_HARNESS.md:284-360, 729-763; grep tabela ponavljanja po fajlovima
- [M] OBIM 4 -- Pravilo 'izmene u .claude/ samo kroz zaseban process PR, jedan po jedan' proizvodi zaostajanje samih pravila. STANJE vodi otvorene stavke 'process PR za .claude/rules' i '.claude/rules/testovi.md ne zna za JS kapiju -- samo process PR'. testovi.md §1 je netacan o self-testovima od 13.09 (oko dva meseca).
    DOKAZ: A/CLAUDE.md:132-134; A/docs/STANJE_REFAKTORA.md:10979-10980, 11014; A/.github/workflows/static.yml:165-171
- [M] POUKA 1 (postmortem §10) -- Pravilo koje kosta vise nego sto ljudi plate ne krsi se glasno, nego se tiho svodi na podskup: 'Nad 222 sabotaze to traje oko dva i po sata, pa se u praksi pustao podskup', i deset sidara je istrunulo. Lek je bio trostruk: jeftina polovina staticki i automatski (hook), skupa polovina kao alat, kriterijum 'crvenih = sabotaza'. Plan zato ne sme ostaviti nijedno pravilo cija je cena van onoga sto ce se zaista platiti; danasnje 'pun katalog pojedinacno pred release' je bas takvo.
    DOKAZ: postmortem:138-151, 172-188
- [M] POUKA 2 (postmortem §10, §14) -- 'Nije primenjeno' ili 'nije se pokrenulo' izgleda kao zeleno ako nije u imeniocu: 'nisam nasla sidro' na stderr usred polusatnog izlaza; sabotaza koja se ne kompajlira prijavljena kao 'NE OBARA NISTA'; puknut ispis (cp1252) pretvorio crven test u 'mrtvu proveru'. Svaka optimizacija koja nesto preskace (izbor testova, kes, rotacija) mora preskoceno da broji i prijavi ('NIJE IZMERENA'), nikad da ga cuti.
    DOKAZ: postmortem:149-151, 160-170, 821-836; A/tools/dokaz.py:520-526
- [M] POUKA 3 (postmortem §10 'Cetiri rupe' br. 1; §9) -- Provera je bila napisana, ali nije bila prikljucena na putanju kojom se stvarno zove (`not args.paths`, a hook zove sa putanjom); hook je mesecima bio mrtav zbog python3 aliasa. Isti oblik danas postoji za tools/sabotaza.py (NESAGLASNOST 4). Svaki novi sloj mora da ima dokaz da ga okidac zaista zove, tj. pad nad namerno pokvarenim ulazom kroz pravi hook.
    DOKAZ: postmortem:210-214, 119-137; izmereno: hook nad tools/sabotaza.py -> rc=0 za 0,056 s
- [M] POUKA 4 (postmortem §12) -- Provera koja je STALNO crvena je isto sto i provera koja nikad nije crvena: 119 zastarelih tvrdnji ucinilo je da 'alat nauci da laje', pa pravu tudju tvrdnju vise niko ne cita. Odatle: priznat nalaz nosi tacan prefiks, ne kategoriju; mrtav upis je nalaz; lazna uzbuna u hook-u je skuplja od propusta. Brze, uze provere ne smeju da uvedu lazne uzbune, a spisak izuzetaka mora biti registar sa razlogom.
    DOKAZ: postmortem:404-412, 190-203, 587-589, 687-691
- [M] POUKA 5 (postmortem §12, §15) -- Dokaz same popravke je dva puta bio prazan: self-test je sklapao sopstveni ulaz i zaobisao funkciju u kojoj je rupa; guard 'izvor sabotiran' je bio zelen nad fixture-om koji --vrati odbija. 'Dvosmerni dokaz pokazuje da grana bira, ne i da bira po pravom pravilu.' Lek je bila jedna deljena funkcija (`_telo_podaci`, `_pogodaka`) umesto dve kopije pravila. Svaki skraceni put (izbor, kes, manifest) treba da deli kod sa punim putem, a ne da ga oponasa.
    DOKAZ: postmortem:485-502, 1099-1144
- [M] POUKA 6 (postmortem §12, §17) -- Staticka provera ne moze da zameni izvrsenje: 8 od 251 (oko 3%) nalaza video je samo pun prolaz, a nova centralna kapija u trecoj proceduri je zastarela dva dokaza kojima se ni sidro ni test nisu menjali. Kes dokaza po otisku procedure sidra i tela testa zato NIJE ispravan sam po sebi; ponovni dokaz mora biti i vremenski periodican.
    DOKAZ: postmortem:630-650, 1246-1261, 1287-1294
- [M] POUKA 7 (postmortem §7; EXCEL_TEST_HARNESS 'Compile') -- Skupa provera na pogresnom okidacu (FULL na svakom Stop-u, 11 suita, oko 1050 provera) ucinila je sesiju neupotrebljivom i sklonjena je. Ishod koji nije eksplicitno OK prolazio je kao zeleno tri-cetiri puta (NEJASNO, lenja kompilacija, tihi Exit Sub). Masovni pad najcesce nije nalaz o proizvodu, nego fixture, sema ili vektor (§1, §2, §6).
    DOKAZ: postmortem:95-104, 70-93, 11-41; A/docs/EXCEL_TEST_HARNESS.md:99-145

## TROSKOVI
[
 {
  "sta": "PostToolUse hook `vba_check.py --hook <fajl>` posle svakog Edit/Write/MultiEdit nad .bas/.cls/.frm/.doccls",
  "koliko": "6,3 s izmereno u ovoj sesiji (modParse.bas); 9,0-11,9 s u zadatim merenjima. Za fajl koji nije VBA: 0,056 s, bez ijedne provere.",
  "izmereno": true,
  "koliko_cesto": "Po svakom Edit-u VBA fajla. Broj Edit-a po rezu nije izmeren; uz pretpostavku 30-100 to je oko 3-20 min cekanja po rezu."
 },
 {
  "sta": "Pun `vba_check.py` + who_writes --check + --check-ownership + gen_schema --check (praksa 'jeftine kapije')",
  "koliko": "31,0 s + 2,4 s + 2,2 s + manje od 1 s, ukupno oko 36 s (zadata merenja)",
  "izmereno": true,
  "koliko_cesto": "Pre svakog commit-a ili push-a VBA izmene (testovi.md:79; STANJE:11020); 59 pominjanja 'jeftine kapije' u STANJE."
 },
 {
  "sta": "CI static.yml: oko 22 koraka kapija + 12 self-testova + JS (setup-python, setup-node, npm ci)",
  "koliko": "Zbir izmerenih Python delova je najmanje 70 s (vba_check 31 + self-test 16,9 + migracija 15,1 + who_writes 4,6 + ostalo manje od 1 s po koraku). Ukupno trajanje posla sa setup-om i npm-om nije izmereno.",
  "izmereno": false,
  "koliko_cesto": "Na svaki push u PR i na push u main."
 },
 {
  "sta": "Jedno podizanje Excela (`run_vba --suite X`): import oko 194 fajla + compile proba (uvek NEJASNO) + EnsureRuntimeSchema + suite",
  "koliko": "Dokumentacija: rezija oko 18 s, RunAllTests 68-91 s, BFP oko 150 s. Te brojke su iz vremena manjih suita. Noviji zapisi: 'crveni RunAllTests 210 s' (04.11), BFP oko 4 min (KI-036), prolaz dokaza nad 4 sabotaze 34 min, runda od 7 podizanja 21 min, runda od oko 18 podizanja 88 min. Efektivno oko 3-6 min po podizanju.",
  "izmereno": false,
  "koliko_cesto": "7-20 puta po rundi acceptance-a."
 },
 {
  "sta": "Runda 'skupih kapija' / 'izvrsni acceptance' po rezu: crveno nad starim kodom + zeleno + dokaz + compile-only",
  "koliko": "7-20 podizanja Excela; 21-88 min po zapisima sa vremenom (stavke 309, 292, 301). Uz to rucni Compile operatera i rucna provera.",
  "izmereno": false,
  "koliko_cesto": "Po rezu, i ponovo posle review kruga koji menja kod (9 stavki 'kapije/acceptance' i 24 review stavke u 70 stavki za 01-06.11)."
 },
 {
  "sta": "Dvosmerni dokaz u rezu (`dokaz.py <prefiks> --grupe 6`)",
  "koliko": "Bazni prolaz je uvek RunAllTests + suite sabotaza (2 podizanja i kad su sve sabotaze u BFP) + 2-7 grupnih ili pojedinacnih prolaza. Medijana 7 sabotaza po dokazu (77 pominjanja 'dokaz N/N', zbir 788). Jedini zapis sa vremenom: oko 34 min za 4 sabotaze.",
  "izmereno": false,
  "koliko_cesto": "Po rezu koji dodaje ili menja test; 78 od 322 stavke pominje dokaz N/N."
 },
 {
  "sta": "Pun katalog dokaza (pravilo: pred release, pojedinacno)",
  "koliko": "1265 prolaza pojedinacno; 215 uz --grupe 6 (oba broja izmerena kroz --plan). Trajanje je procena iz zapisa: oko 60-95 h pojedinacno, oko 11-16 h grupno, plus ponavljanja clanova koji u grupi ne obore svoju tvrdnju. CLAUDE.md navodi 'sat i vise'.",
  "izmereno": false,
  "koliko_cesto": "Po pravilu pred svaki release; stvarno nijednom od 28.08.2026 (tada 261 sabotaza)."
 },
 {
  "sta": "FULL prolaz (12 default gate suita)",
  "koliko": "'~37 min' (STANJE:2484, 10.10) i 'prolaz traje 20-60 min' (EXCEL_TEST_HARNESS:326); nije izmereno ovde.",
  "izmereno": false,
  "koliko_cesto": "Pravilo: pred release i za rizicne izmene u jezgru. Praksa u oktobru: prakticno po rezu (56 stavki pominje zelen 12/12 u 01-30.10); od 01.11 tri pominjanja."
 },
 {
  "sta": "Rucni Compile (Alt+F11 -> Debug -> Compile) + priprema sveske `run_vba --compile-only --keep` + uvoz modVbaTools",
  "koliko": "Jedno dodatno podizanje Excela i rucni korak operatera; trajanje nije zabelezeno. U STANJE: 56 pominjanja Compile PASS, 35 NEJASNO, 129 'rucna provera'.",
  "izmereno": false,
  "koliko_cesto": "Po rundi acceptance-a i posle svake izmene src-vba nakon prethodnog Compile-a (marker compile-a je vezan za otisak izvora)."
 },
 {
  "sta": "Ponistavanje markera: svaka izmena u src-vba (i samo test modula) obara zeleno svih suita i compile; uz to KI-036 (dokaz upise marker nad sabotiranim izvorom)",
  "koliko": "'pun prolaz + dokaz + drugi compile' za jedan premesten red (STANJE:1452-1456); oko 37 min + compile (STANJE:2484); oko 4 min ponovljenog BFP posle svakog dokaza (KI-036).",
  "izmereno": false,
  "koliko_cesto": "Po svakoj naknadnoj ispravci u PR-u; odluke o odlaganju sitnih ispravki (KI-058, 'badge nesacuvano') donose se zbog te cene."
 },
 {
  "sta": "Ucitavanje pravila u kontekst sesije",
  "koliko": "Najmanje 270 linija / 16,1 KB za svaku VBA izmenu; najmanje 815 linija / 48 KB kad se dira bilo koji od 40 fajlova koji aktiviraju testovi.md (prakticno svaki rez sa testom); najvise 979 linija / 55,8 KB. Od 544 linije testovi.md, politika je oko 65.",
  "izmereno": true,
  "koliko_cesto": "U svakoj sesiji; testovi.md u svakoj sesiji koja dira test module, modScr*, modOtkupUI i sl."
 },
 {
  "sta": "Vodjenje dnevnika verifikacije (STANJE_REFAKTORA.md) i ponavljanje istih rezultata",
  "koliko": "11.059 linija, 322 stavke; 28 stavki samo 04.11.2026; merge stavke ponavljaju 'Osnova za main' (iste brojke kao stavka kapija); pasusi o alatima udvojeni sa EXCEL_TEST_HARNESS.",
  "izmereno": true,
  "koliko_cesto": "Na kraju svakog koraka ('u istom commit-u'); fajl je deklarisan kao ulaz za svaku sesiju."
 },
 {
  "sta": "Zaostajanje pravila zbog obaveznog zasebnog process PR-a za .claude/",
  "koliko": "testovi.md §1 netacan o self-testovima oko 2 meseca; komande sa backslash-om; najmanje 2 otvorene stavke 'process PR' u STANJE; pojmovi prakse (jeftine/skupe kapije, --grupe, --require-green) nisu u pravilima.",
  "izmereno": true,
  "koliko_cesto": "Trajno; raste sa svakim alatom dodatim van .claude/."
 },
 {
  "sta": "Crveni prolaz nad starim kodom (rucno vracanje 4-6 produkcionih modula na bazu, pa vracanje nazad)",
  "koliko": "2-3 podizanja Excela + rucna manipulacija radnim stablom; cesto daje i sum ('crveni zbog API-ja'), a ponekad nije moguc (stavka 309: stari kod se ne kompajlira sa novim testovima).",
  "izmereno": false,
  "koliko_cesto": "U 6 od 13 pregledanih rundi kapija (261, 267, 271, 279, 285, 301); pravilo ga trazi samo za bug (CLAUDE.md §2)."
 }
]

## PREDLOZI
### P1 Sloj 'po Edit-u': hook proverava samo ono sto taj Edit moze da pokvari [trud M] [slabi: DA, usko: tvrdnja 'ceo katalog se proverava posle SVAKE VBA izmene' (postmortem §10 t.2, T15) postaje 'deo kataloga koji ta izmena moze da dirne posle svake izmene; ceo katalog pre commit-a i u CI'. Kasnjenje je najvise do sledeceg commit-a, i samo za efekte preko fajlova.]
U `tools/vba_check.py --hook <fajl>` ostaju pravila nad tim fajlom (ASCII, KRAJ_REDA, DEKLARACIJA, REZERVISANO, DUPLIKAT_LOKALNI, DUPLI_LOKAL, ZAKLONJENO, NEDEKLARISAN, REGISTAR za test module, NEDEFINISAN/ARNOST uz kes globalnih definicija po otisku src-vba). Od kataloga se proveravaju samo sabotaze ciji je `fajl` jednak izmenjenom fajlu; kad je izmenjen test modul (modTest, modTestBanka, BFP), jos imena testova i tvrdnje sabotaza tog modula. Ceo katalog, popis suita, amb_bind i dokaz self-test prelaze u sloj 'pre commit-a' i ostaju u CI. Zasto je sloj dovoljan: Edit fajla F moze da pomeri samo sidra u F i, ako je F test modul, tvrdnje i imena u F. Pravila preko vise unosa (zamka 5 i 8) menjaju se samo izmenom kataloga, a nju pokriva P2.
  DOBITAK: Hook sa 6,3-11,9 s na cilj od najvise 2-3 s po Edit-u. Procena iz zadatog profila: katalog 53%, arnosti 21%, popis 10%, amb_bind 7%. 116 od 193 VBA fajla nema nijednu sabotazu; medijana je 7 po fajlu, maksimum 128.
  RIZIK: Pretpostavka o lokalnosti moze biti pogresna za neko buduce pravilo kataloga. Kes definicija moze da zastari, pa mora biti vezan za otisak i fail-closed (bez kesa = pun obracun).
  PREDUSLOVI: Izmena je u tools/ (obican PR, ali serijski: test harness je deljen fajl, pre-flight §8). Ako se menja i skripta hook-a, to je zaseban process PR. P3 (pun set pre commit-a) mora postojati pre nego sto se hook suzi.
  DOKAZATI: (1) Ekvivalencija bez Excela: za svaki od 77 sabotiranih fajlova na kopiji src-vba pokvari jedno sidro; suzeni hook nad tim fajlom mora dati isti nalaz kao pun vba_check. (2) Self-test u oba smera kroz main(): sidro u drugom fajlu se u hook rezimu ne prijavljuje, a u punom pada po imenu; sabotaza koda suzavanja (filter vrati prazno) obara slucaj. (3) Merenje trajanja pre i posle na istoj masini.

### P2 Hook pokriva i alatne fajlove; `sabotaza.py --self-test` ulazi u CI [trud S] [slabi: NE -- jaca T15 (katalog ziv) i T5 (self-test checkera u CI).]
U `.claude/hooks/vba-check.sh` granu `case` prosiriti na tools/sabotaza.py, tools/dokaz.py, tools/vba_gate.py i tools/run_vba.py. Za njih se zove `vba_check.py --hook <fajl>`, koji vec ume da vrati verdikt kapija nad alatima kad nema VBA fajlova (`if not files: return rc_alata`). U static.yml dodati korak `python3 tools/sabotaza.py --self-test`.
  DOBITAK: Zatvara izmerenu rupu: Edit kataloga danas ne pokrece nijednu proveru (rc=0 za 0,056 s), pa greska u katalogu stize tek u CI. Cena je oko 3-6 s po Edit-u ta cetiri fajla i 0,54 s u CI.
  RIZIK: Nizak. Edit velikog tools/sabotaza.py bice sporiji za trajanje provere kataloga.
  PREDUSLOVI: Izmena hook-a je u .claude/, dakle zaseban process PR; CI korak je obican PR.
  DOKAZATI: Kroz pravi hook, sa payload-om kao u ovoj analizi: nad kopijom tools/sabotaza.py sa namerno pokvarenim sidrom izlaz mora biti rc=2 (danas je rc=0); vracen fajl daje rc=0. Za CI: privremeno pokvari jedno pravilo u `_nalazi` na grani, pa korak mora pasti po imenu slucaja.

### P3 Jedan manifest kapija ('sloj -> komanda -> kada') sa pokretacem i proverom dokumenata [trud M] [slabi: NE.]
Dodati `tools/kapije.py` sa deklarativnim spiskom: id, komanda, sloj (edit, commit, rez, pr, nocno, release), da li trazi Excel, self-test, budzet u sekundama. `--sloj commit` pusta danasnje 'jeftine kapije' jednom komandom (pun vba_check, who_writes oba rezima, gen_schema --check, vba_gate --popis; self-test alata samo kad je taj alat u diff-u). `--check` u CI tvrdi da svaki korak kapije u static.yml postoji u manifestu i obrnuto, i da tabela nivoa u CLAUDE.md §5 i testovi.md §1 odgovara manifestu. Zasto je sloj 'pre commit-a' dovoljan: staticke kapije su deterministicke nad stablom, pa je njihovo ponavljanje po Edit-u cist trosak; CI ih i dalje vrti nezavisno od masine.
  DOBITAK: Jedna lokalna komanda (oko 36-40 s) umesto pamcenja 4-22 komande. Uklanja klasu nesaglasnosti 6, 7 i 13 (tabela u dva fajla, spisak self-testova) trajno, jer drift obara CI.
  RIZIK: Nov checker je nova stvar koja moze da lazno prolazi; zato ima svoj self-test. Parser YAML koraka mora biti uzak, da ne daje lazne uzbune.
  PREDUSLOVI: Obican PR za tools/ i CI; izmena tabela u CLAUDE.md i testovi.md ide kroz process PR (P11).
  DOKAZATI: Self-test u oba smera kroz main(): korak koji postoji samo u static.yml, korak samo u manifestu i izmenjena komanda u tabeli dokumenta -- svaki pada po imenu; cist repo je zelen. Sabotaza poredjenja (uvek True) obara sva tri slucaja.

### P4 Sloj 'po rezu': zapisati 'skupe kapije posle CODE GO' i dokaz nad NOVIM i DIRNUTIM sabotazama koje racuna alat [trud M] [slabi: NE -- danasnje pravilo trazi samo 'NOVE ili dirane' po izboru coveka; ovo ga cini mehanickim i sirim.]
Pravilo u CLAUDE.md §5 / testovi.md: posle reviewer CODE GO pusta se TARGETED suita (ili vise njih) i `dokaz.py --grupe 6` nad novim i dirnutim sabotazama. 'Dirnute' ne bira covek, nego alat (`dokaz.py --dirnuto <baza> --plan`): (a) unosi kataloga izmenjeni u odnosu na merge-base, (b) sabotaze cije je sidro u proceduri koju diff menja, (c) sabotaze ciji je test menjan. Bazni prolaz dokaza se priznaje kao zeleni prolaz tih suita (vec praksa, stavka 288). Zasto je sloj dovoljan: sve sto rez moze neposredno da pokvari u katalogu jeste (a)-(c); posredne efekte (postmortem §17) hvata nocna rotacija iz P6.
  DOBITAK: Bez odlozene trulezi: 2 od 11 odlozenih 'dirnutih' sabotaza je istrunulo (stavke 288 i 322). Nema dvostrukog RunAllTests/BFP (zasebno 'zeleno' + bazni prolaz), sto je 2 podizanja manje po rundi, oko 6-10 min.
  RIZIK: Racunanje (b) trazi mapu procedura; dokaz.py je vec ima (`_procedura`). Preuzak skup ako sabotaza zavisi od trece procedure -- to pokriva P6.
  PREDUSLOVI: Izmena tools/dokaz.py sa self-testom kroz vba_check; tekst pravila kroz process PR.
  DOKAZATI: Bez Excela: self-test `--dirnuto` nad sintetickim diff-om (izmenjen unos, izmenjena procedura sidra, izmenjen test: svaki ulazi u skup; nepovezana izmena ne ulazi). Retroaktivno: nad diff-om reza iz stavke 288 alat mora da izabere svih 11 odlozenih sabotaza.

### P5 'Crveni prolaz nad starim kodom' obavezan samo za bug; kad se pusta, ide iz git worktree-a [trud S] [slabi: DA, za rezove koji nisu bug: tvrdnja 'novi test je crven nad STVARNIM prethodnim kodom' menja se u 'novi test je crven nad sabotazom koja vraca prethodno ponasanje'. Tu tvrdnju danas ne trazi nijedno pisano pravilo (samo praksa i pre-flight §1 'crveno -> zeleno'); T21 za bug ostaje netaknuta.]
Zapisati: reprodukcija nad stvarnim starim kodom je obavezna po CLAUDE.md §2 samo za bug ili KI. Za rez koji uvodi ponasanje, crveni smer daje sabotaza (to je vec priznato u stavci 309: 'Staro ponasanje reprodukuju sabotaze'). Kad se crveni prolaz ipak radi, produkcioni moduli sa baze se uzimaju u zasebnom worktree-ju ili privremenoj kopiji izvora, a ne rucnim vracanjem 4-6 modula u radnom stablu.
  DOBITAK: 2-3 podizanja Excela manje po rundi za rezove koji nisu bug (oko 6-15 min). Nestaje rizik da vraceni moduli ostanu u radnom stablu.
  RIZIK: Sabotaza nije uvek identicna starom kodu, pa 'crveno po sabotazi' i 'crveno na stvarnoj bazi' nisu ista tvrdnja.
  PREDUSLOVI: Odluka operatera (praksa bez pisanog pravila); tekst kroz process PR. Varijanta sa worktree-om trazi da run_vba primi putanju izvora -- nije provereno da li to danas ume.
  DOKAZATI: Na jednom sledecem rezu uporediti: skup tvrdnji crvenih nad stvarnom bazom i skup crvenih kroz sabotaze. Ako su isti, pravilo je bez gubitka; razlika se zapisuje kao nalaz.

### P6 Sloj 'nocno/periodicno': FULL nad main + rotacioni grupni dokaz sa knjigom [trud L] [slabi: NE -- uvodi sprovodjenje za T26, koju danas nista ne sprovodi. Grupni rezim se i dalje prijavljuje kao grupni (T18).]
Zakazan prolaz na Windows masini sa Excelom, nad svezim klonom ili worktree-om `main`-a (ne nad radnim stablom operatera). Radi: (1) FULL 12 suita i marker za glavu main-a; (2) `dokaz.py --grupe 6` sa vremenskim budzetom (npr. 2-3 h) nad sabotazama po prioritetu: nikad dokazane nad tekucim main-om, pa najstariji dokaz. Ishod po sabotazi ide u knjigu (datum, otisak izvora, rezim grupno ili pojedinacno, ishod). Knjiga je fail-closed: stavka bez dokaza mladjeg od N dana vodi se kao 'NIJE DOKAZANO', ne kao zelena. Zasto je sloj dovoljan: postmortem §17 kaze da je pun prolaz periodican jer ispravna izmena u trecoj proceduri zastareva tudji dokaz; to se ne moze vezati ni za rez ni za release, samo za kalendar.
  DOBITAK: Zatvara rupu od oko 10 nedelja bez punog kataloga (261 -> 1265). Ceo katalog obidje se za oko 5-8 noci (215 grupnih prolaza, oko 11-16 h po proceni). FULL po rezu postaje nepotreban, jer se regresija van ciljanih suita vidi u roku od 24 h.
  RIZIK: Masina operatera: sabotaza `pdf-izvoz-ignorise-kljuc` otvara PDF-ove ('pusta se kad operater ne radi'). Prekinut dokaz ostavlja sabotiran izvor; zato zaseban klon i `--vrati` + provera otiska na pocetku. Knjiga je nova tvrdnja i mora imati dokaz u oba smera.
  PREDUSLOVI: P7 (dokaz ne sme da upisuje marker nad mutiranim izvorom, KI-036). Zakazivanje na Windows-u (Task Scheduler) je van repoa. Odluka gde zivi knjiga (gitignored, kao last_green.json). Ako grana B (`--izbor`) udje u main, rotacija se drasticno skracuje, ali VBA strana te grane nikad nije izvrsena u Excelu.
  DOKAZATI: Bez Excela: self-test knjige -- stavka starija od N dana, stavka sa drugim otiskom i stavka koja fali padaju po imenu; plan rotacije pokriva svaku sabotazu tacno jednom po ciklusu (ista tvrdnja kao `_plan`). Sa Excelom: prvi pun ciklus nad main-om; ocekuju se nalazi (stope: oko 3% u §12, 2 od 261 u §17, 2 od 11 u stavci 322), a njihov broj je merilo vrednosti.

### P7 KI-036: mutirani prolazi dokaza ne upisuju zeleni marker [trud S] [slabi: NE -- jaca T10: marker vise ne moze nastati nad sabotiranim izvorom.]
`dokaz.py` zove `run_vba.py` za mutirane prolaze sa prekidacem koji gasi upis markera (npr. `--bez-markera`); bazni prolaz sme da ga upise. Danas run_vba nema takav argument (argumenti: --workbook, --compile-only, --no-import, --suite, --all, --keep, --timeout-dialog, --timeout, --self-test, --ignore-fixture-sig).
  DOBITAK: Jedan ponovljen prolaz suite manje posle svakog dokaza (oko 4 min za BFP po KI-036). Nestaje poruka 'dokazan nad DRUGIM izvorom' posle uspesnog FULL-a. Preduslov za P6.
  RIZIK: Nizak. Ne sme da ugasi upis i za bazni prolaz; zato ide tvrdnja u self-testu.
  PREDUSLOVI: Menja run_vba.py i vba_gate.py, koji su delovi 'ugovora' markera, pa obara postojece zelene markere jednom. Dvosmerni dokaz je obavezan (menja se checker).
  DOKAZATI: Bez Excela: self-test orkestracije u dokaz.py vec lazira `_pusti`; dodati tvrdnju da mutirani poziv nosi prekidac, a bazni ne. U run_vba/vba_gate self-testu: sa prekidacem `zabelezi_prolaz` nije pozvan; sabotaza (ignorisan prekidac) pada po imenu. Sa Excelom: FULL -> dokaz -> `--require-green` ostaje zelen bez ponavljanja.

### P8 Sloj 'po PR-u': CI + jedan zavrsni acceptance na glavi koja se spaja (Compile jednom po PR-u) [trud S] [slabi: NE za spojeni kod: T10 i T20 se mere na glavi koja ulazi u main. Medjuglave se ne dokazuju, ali se ni ne spajaju.]
Zapisati kao pravilo ono sto je vec praksa ('Skupe kapije cekaju reviewer GO'). Review krugovi se rade nad kodom, jeftinim kapijama i CI-jem. Excel kapije i rucni Compile idu jednom, nad zavrsnom glavom, a merge se vezuje za `vba_gate --require-green --require-compile` nad tim izvorom (uz --suite za suite koje rez trazi). Naknadna izmena posle acceptance-a ponavlja samo suite koje ta izmena dira, pod uslovom da P9 to masinski dokaze. Zasto je sloj dovoljan: spaja se samo zavrsna glava, a marker vec dokazuje da je bas taj izvor meren.
  DOBITAK: Jedna runda umesto 2-3 po PR-u (primer #454: kapije u stavkama 267 i 271; #459: 285 i 288). Usteda je 7-20 podizanja Excela i jedan rucni Compile po izbegnutoj rundi.
  RIZIK: Ako review tek posle acceptance-a nadje P1, runda se ponavlja kao i danas. Operater moze hteti rane kapije za rizicne rezove; to ostaje dozvoljeno.
  PREDUSLOVI: Tekst pravila kroz process PR; alat vec postoji (vba_gate --require-green).
  DOKAZATI: Na sledeca tri PR-a zabeleziti broj rundi i podizanja Excela pre i posle. Kontrola: nijedan merge bez `--require-green --require-compile` rc=0 nad spojenom glavom; izlaz se prilaze uz PR.

### P9 Marker zelenog po suiti: izmena samo tudjeg test modula ne obara zeleno suite koja ga ne poziva [trud L] [slabi: DA, uslovno: T10 'bas OVAJ izvor' postaje 'bas ovaj izvor u delu koji suite moze da izvrsi'. Ako graf propusti zavisnost, marker bi bio lazno zelen; otuda fail-closed pravila i Compile nad celim izvorom.]
Danas je `izvor` otisak celog src-vba (bez modBuildInfo), pa izmena samo modBusinessFlowProTests.bas obara i marker Palete/Storno suita. Operater to vec rucno obrazlaze ('ddf0230c menja samo modBusinessFlowProTests, pa Palete i Storno sa d69937ff vaze', stavka 279). Predlog: otisak suite = produkcioni moduli + test moduli do kojih ta suite staticki dopire (graf poziva iz definicija koje vba_check vec skuplja). Ako se graf ne moze pouzdano utvrditi, vazi danasnje ponasanje (fail-closed).
  DOBITAK: Kod ispravki koje diraju samo test ne ponavljaju se nepovezane suite: 2-6 podizanja Excela po rundi. Rucno obrazlozenje iz stavke 279 postaje masinsko.
  RIZIK: Srednji: staticki graf ne vidi `Application.Run` po imenu ni kasno vezivanje; modul koji se ne kompajlira obara ceo projekat bez obzira na graf. Zato se Compile i dalje vezuje za ceo izvor.
  PREDUSLOVI: Menja `kapija` deo ugovora (obara postojece markere jednom); dvosmerni dokaz obavezan; serijski sa P7. Predlog je nizeg prioriteta.
  DOKAZATI: Bez Excela: self-test sa sintetickim modulima (suite A zove X; izmena Y ne obara A; izmena X obara; `Application.Run "ime"` u dometu suite vraca na pun otisak). Sabotaza: graf uvek prazan -> slucaj 'izmena X obara A' pada. Sa Excelom: namerno izmeni produkcioni modul u dometu i potvrdi da marker pada.

### P10 Sloj 'pred release': izvrsna release kapija i ostvariv obim punog dokaza [trud M] [slabi: V1: DA -- T18 za release: 'svaka sabotaza POJEDINACNO pred release' postaje 'grupno nad RC + rotaciono u poslednjih N dana'; grupni prolaz ne razlikuje sopstvenu mutaciju od partnerove (osim za clanove koji se ponove sami). V2: NE, ako se baza tvrdi nad istim izborom (T1) i test koji ne prolazi sam meri celom suitom.]
(a) U RELEASE_PROCEDURE.md i tools/release.sh uvesti redosled: bump APP_VERSION (commit) -> FULL + dokaz + rucni Compile nad TIM izvorom -> `vba_gate --require-green --require-compile` rc=0 -> tek tada tag i push. Danas release.sh taguje i gura tag pre Compile-a i bez ijedne kapije, a bump menja otisak izvora. (b) Obim punog dokaza pred release, dve varijante. V1: grupno nad RC (215 prolaza, oko 11-16 h, jedna noc ili vikend) + uslov da je svaka sabotaza u knjizi iz P6 dokazana u poslednjih N dana. V2: pojedinacno, ali tek kad izbor testova (grana B, `--izbor`) udje u main i bude izmeren. Zasto je sloj dovoljan: FULL i Compile nad tagovanim izvorom dokazuju bas ono sto se isporucuje; periodicna rotacija drzi katalog svezim izmedju izdanja.
  DOBITAK: Pravilo postaje izvrsivo: danasnje 'pojedinacno pred release' je reda 60-95 h i nije izvrseno nijednom nad 1265. Zatvara nesaglasnosti 10 i 11 (proceduru bez kapija i bump koji ponistava marker).
  RIZIK: V1 daje slabiju tvrdnju od pisane (v. slabi_tvrdnju). Izmena release.sh dira distribuciju; pogresan redosled moze da blokira hitno izdanje, pa treba svestan, zabelezen izlaz kao --ignore-fixture-sig.
  PREDUSLOVI: P6 i P7; odluka operatera izmedju V1 i V2; izmena git-i-release.md i CLAUDE.md kroz process PR; release.sh i RELEASE_PROCEDURE.md kroz obican PR. NESAGLASNOST 11 mora prvo da se potvrdi izvrsenjem.
  DOKAZATI: Bez Excela: self-test release kapije u izolovanom privremenom repou (bez push-a) -- marker za drugi izvor -> release.sh staje pre taga; marker za isti izvor -> prolazi. Otisak: izmena APP_VERSION menja `otisak_izvora` (potvrda nesaglasnosti 11) i posle promene redosleda marker vazi za tagovani commit.

### P11 Jedan process PR: uskladiti pravila sa kodom i praksom i svesti testovi.md na politiku [trud M] [slabi: NE formalno (nijedan alat ne sprovodi tekst pravila). Zastita seam-ova (T14) je danas samo dokumentaciona; premestanje teksta je ne menja, ali smanjuje verovatnocu da bude procitan.]
(1) CLAUDE.md §5 'Obim dokaza': jedna suita po prolazu, bez zakucanih brojki (uputiti na `dokaz.py --plan`), komanda reza sa --grupe, slojevi iz manifesta (P3). (2) testovi.md: `/` umesto `\` (linije 55-57, 173); §1 spisak self-testova zameniti pokazivacem na manifest; u `paths:` dodati tools/dokaz.py, tools/vba_gate.py, tools/vba_import_marker_gate.py i .claude/hooks/*. (3) Katalog seam-ova (§4, 208 linija) i istorije incidenata preseliti u docs/ (npr. docs/TEST_SEAMOVI.md i postmortem). U pravilima ostaje kratak spisak zabrana sa pokazivacem: ne vracaj MsgBox u putanju pisca, kuka na vrhu ili na izlazu clana, nov seam ide u ResetSeamova. (4) dokaz.py:4-5 docstring, EXCEL_TEST_HARNESS:91 i 140-145, komentar hook-a i postmortem §7/§8 (napomena da Stop hook vise ne postoji) -- ovi nisu u .claude/, pa mogu u obican docs PR.
  DOBITAK: testovi.md sa 544 na oko 120-150 linija: oko -400 linija / -24 KB u svakoj sesiji koja dira test ili ekran. Uklanja nesaglasnosti 1, 3, 6, 7, 8, 9, 12, 13 i 14.
  RIZIK: Agent koji menja modScr* ili modOtkupUI vise ne dobija ceo opis seam-ova automatski i moze ih 'pocistiti'. Ublazenje: kratak spisak zabrana ostaje u pravilu, a opis stoji uz same seam-ove u kodu i u docs/.
  PREDUSLOVI: Zaseban process PR, jedan po jedan (CLAUDE.md §6), posle P3 da tabela ima izvor. Bez paralelne sesije nad .claude/.
  DOKAZATI: `wc -l` pre i posle; skripta nad `paths:` (ista kao u ovoj analizi) pokazuje nove minimume i maksimume ucitanih linija. `tools/kapije.py --check` (P3) je zelen nad novim tabelama, a crven kad se tabela rucno promeni. grep: 0 backslash komandi u .claude/rules.

### P12 Budzet trajanja kao kapija: hook i vba_check ne smeju tiho da uspore [trud S] [slabi: NE.]
U CI izmeriti trajanje `vba_check.py --hook` nad jednim reprezentativnim fajlom i punog `vba_check.py` (medijana tri pokretanja) i pasti preko budzeta iz manifesta (npr. hook 3 s posle P1, vba_check 45 s). Razlog: dokumenti i dalje tvrde '0,14 s' i 'sekundu', a merenje daje 6-12 s; katalog raste oko 60% za 24 dana i nista to ne signalizira.
  DOBITAK: Sprecava povratak danasnjeg stanja posle P1. Brojke u dokumentima prestaju da budu prepisivane, jer ih nosi manifest.
  RIZIK: Razlika u brzini CI izvrsilaca moze dati laznu uzbunu; zato sirok prag i medijana. Lazna uzbuna u kapiji je skupa (T23), pa prag treba kalibrisati merenjem na CI-ju, ne na lokalnoj masini.
  PREDUSLOVI: P1 i P3; nekoliko CI pokretanja za kalibraciju praga.
  DOKAZATI: Dvosmerno: na grani dodati vestacko kasnjenje u `check_katalog_sabotaza` -> korak pada sa porukom o budzetu; bez kasnjenja prolazi. Pet uzastopnih CI pokretanja bez lazne uzbune pre nego sto korak postane blokirajuci.

### P13 STANJE_REFAKTORA: arhivirati istoriju, ne ponavljati rezultate kapija [trud S] [slabi: NE -- zapis ishoda kapija ostaje; uklanja se samo ponavljanje.]
Fajl koji je deklarisan kao 'ulaz za svaku novu sesiju, kratko i azurno' svesti na: 'Gde smo', tekuci slajs (poslednjih oko 20 stavki), 'Dug sa imenom' i pokazivac na manifest kapija. Starije stavke idu u arhivu (docs/arhiva/). Merge stavke ne ponavljaju 'Osnova za main', nego pokazuju na stavku kapija. Pasuse 'Grupni dokaz' i 'Marker zelenog' ostaviti samo u EXCEL_TEST_HARNESS.
  DOBITAK: Sa 11.059 linija na oko 600-800. Manje pisanja po koraku (oko 1/3 poslednjih 70 stavki su merge i review zapisi koji ponavljaju brojke). Dnevnik prestaje da zaostaje za main-om (danas: poslednja stavka 06.11, main 08.11).
  RIZIK: Gubitak grep-traga po brojevima stavki ako arhiva ne zadrzi numeraciju; zadrzati brojeve i jednu datoteku arhive.
  PREDUSLOVI: Odluka operatera (fajl je njegov radni dnevnik); nije u .claude/, pa ide obican docs PR.
  DOKAZATI: `wc -l` pre i posle; grep po uzorku od 20 brojeva stavki i KI oznaka nalazi svaku u glavnom fajlu ili u arhivi; nijedan pokazivac 'stavka N' iz KNOWN_ISSUES ne ostaje bez cilja.


## NEPOZNANICE
[
 "Stvarno trajanje jednog prolaza suite danas (RunAllTests 350 testova; BFP 3208 tvrdnji) i rezije run_vba poziva. Procene za pun katalog (60-95 h pojedinacno, 11-16 h grupno) izvedene su iz cetiri zapisa u STANJE (stavke 292, 301, 309 i KI-036) i starih brojki iz dokumentacije; nista od toga nije izmereno ovde.",
 "Koliko je od 1265 sabotaza danas zastarelo u smislu postmortem §17 (tvrdnja koja vise nije ona koja pada). Poslednji pun prolaz je bio nad 261 unosom. Poznate stope su oko 3% (§12), 2 od 261 za jedan dan izmena (§17) i 2 od 11 odlozenih (stavka 322); ekstrapolacija na 1265 nije pouzdana.",
 "Da li je od 28.08.2026 ipak pusten pun katalog a da to nije zapisano. Pretrazeni su STANJE_REFAKTORA, postmortem, RELEASE_NOTES i KNOWN_ISSUES; REFAKTOR_DOKUMENT_HEADER_STAVKE.md i UI_MIGRACIJA_KATALOG.md nisu citani u celini.",
 "Zasto bazni prolaz dokaz.py uvek ukljucuje RunAllTests (`| {SUITE_ALL}`, linija 466) i kad su sve izabrane sabotaze u BFP. Komentara nema; moguce je da sluzi kao compile signal. Dok se to ne zna, uklanjanje se ne predlaze.",
 "Koliko traje compile proba koja se radi na SVAKOM pozivu run_vba i po zapisima uvek zavrsi kao NEJASNO, i da li se moze preskociti kad idu suite. To je oblast analize run_vba i ovde nije mereno.",
 "NESAGLASNOST 11 (bump APP_VERSION ponistava marker pred release) izvedena je citanjem release.sh i vba_gate.py i nije izvrsena. Nije jasno ni da li se tools/release.sh uopste koristi u tekucem ciklusu: APP_VERSION je 2.28.4, a RELEASE_NOTES vodi v2.92.0 i v2.93.0.",
 "Broj Edit/Write poziva nad VBA fajlovima po rezu (mnozilac cene hook-a) nije izmeren; trazi transkripte sesija ili brojac u hook-u.",
 "Ukupno trajanje CI posla `static` po push-u (sa setup-python, setup-node i npm ci) nije izmereno; dostupno je iz GitHub Actions istorije.",
 "Da li staticki graf poziva (P9) moze dovoljno pouzdano da odredi domet suite u ovom kodu. Ugovor ekrana ide kroz `Application.Run` po imenu, sto graf ne vidi; nije ispitano koliko suita to pogadja.",
 "Da li operater moze da ustupi Windows masinu za nocni prolaz (P6) i da li zakazani Excel COM prolaz radi bez prijavljene interaktivne sesije; VBE i SendKeys deo compile probe trazi vidljiv prozor.",
 "Grana B (`run_vba --test`, `dokaz --izbor/--izolacija`): VBA strana nikad nije izvrsena u Excelu i sudara se sa main-om. Varijanta V2 u P10 i skracenje rotacije u P6 zavise od nje, a ovde nije ocenjivana.",
 "Uzorak prakse pokriva period 01-06.11.2026 (stavke 254-323) i oktobarske zbirne brojke; rezovi posle stavke 323 (PR #471-#473) nisu u STANJE, pa njihova cena kapija nije uracunata."
]
