# tools/run_vba.py -- headless Excel runner (režija jednog poziva i arhitektura izvršavanja)

## SAZETAK
Jedan poziv run_vba je strogo linearan i sve plaća iznova: nov EXCEL.EXE, Open, uvoz 192 komponente (43 .doccls merge + 149 Import), compile-detekcija, EnsureRuntimeSchema, suite, gašenje. Od dokumentovanih ~18 s režije, 4,5-7,9 s su fiksni sleep-ovi u kodu (25-44%); to je jedini deo režije koji se zna bez Excela, a Python strana (heševi, potpis fixture-a, kopije) je izmerena na manje od 0,2 s. Compile korak u režimu sa suitama ne utiče na rc osim eksplicitnog FAIL-a i nikad se ne upisuje u marker, a verdikt je po dokumentaciji često NEJASNO, pa najskuplji fiksni deo najčešće ne proizvodi nikakvu tvrdnju. Dijalozi su jeftini (oko 0,2 s po dijalogu, 11 po FULL prolazu, 1 po BFP), ali su MsgBox-ovi bezuslovni, zastavica `dialogs` se nigde ne čita, i "neočekivan dijalog = pad" iz dokumentacije nije implementiran. Tvrdi prekid je JEDAN tajmer od 600 s za ceo poziv; to je već jednom oborilo prolaz i uz rast testova od ~60% za 24 dana postaje redovan kvar. Katalog sabotaža je 99,4% nad .bas fajlovima (0 nad .frm i tvrdim klasama), prosečan mutirani fajl je 1,6% izvora, a svaki prolaz ponovo uvozi 100%. Preporučeni redosled: instrumentacija faza; uklanjanje fiksnih čekanja i compile koraka iz mutacionih prolaza uz timeout po suiti; rebase grane B; `--src` i jednokratni uvoz po dokaz-pozivu sa delta zamenom modula; paralelne instance; trajna sesija tek ako merenje pokaže da start Excela i dalje dominira. Trajna sesija nosi najveći rizik (OnTime živi u Excelu, Application zastavice, BFP ostavlja podatke) i traži kanarinca i automatski pad na hladan put. Grana B: Python strana se primenjuje bez sudara, sudari su samo u modTest.bas, modBusinessFlowProTests.bas i sabotaza.py, a VBA strana traži prenumeraciju testa 311 i mehaničko omotavanje 391 BFP poziva. Sve procene o Excelu su neizmerene i tako su označene.

## CINJENICE
- [M] Oznake putanja u ovom izveštaju: [A] = <scratch>/main-wt (main @1fa4937), [B] = /home/user/otkupapp-pwa (grana izbora, 9d6e780). Privremeni fajl merenja: <scratch>/runvba-analiza/mjeri_kontekst.py.
    DOKAZ: git -C [B] rev-parse --short origin/main -> 1fa4937; git log -1 -> 9d6e780
- [M] Faze jednog poziva, redom: potpis fixture-a (830-835) -> mkdtemp + kopija sveske + kopija golden-a (837-844) -> snimi_kontekst (854-863) -> CoInitialize + DispatchEx + pid + watchdog + killer tajmer (873-886) -> Visible/DisplayAlerts/AutomationSecurity/EnableEvents=False (888-891) -> Workbooks.Open (893) -> import_src_vba + report_orphan_components (895-897) -> compile_project (903) -> EnsureRuntimeSchema (938) -> petlja suita (944-972) -> Close svih sveski (999-1000) -> sleep 1.0 + watchdog.stop (1012-1013) -> Quit, CoUninitialize, sleep 1.0, TerminateProcess (1015-1027) -> golden nazad u repo (1031) -> rmtree (1036) -> zabelezi_prolaz (1055-1056) -> _write_report (1060).
    DOKAZ: [A]/tools/run_vba.py:795-1061
- [M] Fiksni sleep-ovi: 2.5 s posle svakog compile pokušaja (617); 0.5 s i 0.4 s u SendKeys pokušaju (632, 634); 1.0 s pre gašenja watchdog-a (1012); 1.0 s pre TerminateProcess (1021). Zbir je 4.5 s kad se radi samo COM pokušaj, a 7.9 s kad COM Execute ne uradi ništa pa se ide i na SendKeys (uslov na 575).
    DOKAZ: [A]/tools/run_vba.py:617, 632, 634, 1012, 1021, 575-577
- [M] Uvoz nije jedan poziv nego dve petlje: (1) 43 .doccls: VBComponents(base) -> DeleteLines -> AddFromString(telo bez header-a), 42 od 43 su prazni stubovi od 9 redova; (2) 149 fajlova (138 .bas bez modVbaTools + 10 .cls + 1 .frm): pokušaj Remove (izuzetak se guta) pa VBComponents.Import. Nema preskakanja neizmenjenih modula.
    DOKAZ: [A]/tools/run_vba.py:353-406; ls [A]/src-vba po ekstenziji: 139 bas, 10 cls, 43 doccls, 1 frm, 1 frx; wc -l *.doccls: ThisWorkbook 156, ostali 9
- [M] Nad podrazumevanim fixture-om Remove grana se praktično nikad ne izvrši: make_fixture.strip_vba uklanja sve std/class/form komponente, pa VBComponents(base) baca izuzetak 149 puta po pozivu. Remove+Import nad postojećim modulima se vozi samo uz --workbook (npr. DEV sveska).
    DOKAZ: [A]/tools/make_fixture.py:2820-2856; [A]/tools/run_vba.py:394-397; [A]/docs/EXCEL_TEST_HARNESS.md:484
- [M] Obim koji se uvozi u svakom pozivu: 232.957 linija / 10,44 MB u .bas/.cls/.frm; modTest.bas (31.061 linija, 1,58 MB) i modBusinessFlowProTests.bas (33.077 linija, 1,53 MB) su zajedno 30% bajtova.
    DOKAZ: wc -lc [A]/src-vba/*.bas *.cls *.frm
- [M] Python strana režije je zanemarljiva: snimi_kontekst 0,043-0,045 s (zove se dvaput: pre run-a i u zabelezi_prolaz), import make_fixture 0,027 s + signature 0,004 s, run_vba --self-test 0,05 s, _git_glava 0,003 s.
    DOKAZ: python3 mjeri_kontekst.py [A] -> 'snimi_kontekst 0.045 s / 0.043 s / 0.043 s; otisak_izvora 0.043 s; _git_glava 0.003 s; import make_fixture 0.027 s; signature 0.004 s'; time python3 tools/run_vba.py --self-test -> real 0m0.050s
- [M] U režimu sa suitama compile verdikt menja rc samo kad je eksplicitno False; NEJASNO ne obara run. Compile se u marker ne upisuje ni kad je OK. Sam korak detekcije je izmereno remetio Excel (posle njega nekvalifikovan Application.Run prestaje da radi).
    DOKAZ: [A]/tools/run_vba.py:905-918, 927-937, 1046-1048
- [M] DialogWatchdog obilazi prozore klase #32770 svog PID-a na 0,4 s (podrazumevano, --timeout-dialog). Dijalog zato u proseku čeka poll/2 = 0,2 s, najviše 0,4 s. Dugme bira po natpisu (end/ok/u redu/da/yes/cancel...), nikad Debug/Help.
    DOKAZ: [A]/tools/run_vba.py:168, 187, 197, 247-261, 704
- [M] Suite sa dialogs:True u podrazumevanom setu ima 7 i zajedno otvaraju 11 dijaloga: SheetsJson 1 (rezime), Banka 2 (potvrda + rezime), Faktura 1, Storno 2, Palete 2, BFP 1 (EndRun), Agrohemija 2. RunAllTests nema nijedan izvršni MsgBox.
    DOKAZ: [A]/src-vba/modGoogleSyncSmokeTests.bas:1267-1275; modTestBanka.bas:71, 2367-2371; modFakturaTests.bas:986-992; modTestStorno.bas:54, 1133; modTestPalete.bas:93, 975-979; modBusinessFlowProTests.bas:3113-3119; modAgrohemijaTests.bas:48, 555-558; grep -c izvršnih MsgBox u modTest.bas -> 0
- [M] Ne postoji test-režim koji preskače te MsgBox-ove: svi su bezuslovni. modTestMode.IsTestMode postoji, ali ga EndRun i ostali završeci ne pitaju. BFP BeginRun je svoj disclaimer namerno učinio nemodalnim zbog automatizacije, a EndRun i dalje diže MsgBox.
    DOKAZ: [A]/src-vba/modTestMode.bas:21-29; modBusinessFlowProTests.bas:3077-3087 (komentar) i 3113-3119 (MsgBox)
- [M] Zastavica `dialogs` u SUITES se nigde ne čita, i nijedan dijalog viđen tokom suite ne utiče na rc (watchdog.seen se koristi samo u compile koraku i za ispis DIALOG redova). Docstring i dokumentacija tvrde suprotno ("2 = ... neočekivan dijalog").
    DOKAZ: grep -n 'dialogs\|watchdog' [A]/tools/run_vba.py -> upotrebe samo na 575-618, 865, 1014, 1106; grep u vba_gate.py/dokaz.py/vba_check.py -> ništa; tvrdnja: run_vba.py:25 i docs/EXCEL_TEST_HARNESS.md:28-29
- [M] Tvrdi prekid je jedan threading.Timer za CEO poziv (uvoz + compile + sve suite), podrazumevano 600 s. To je već oborilo prolaz: "prvi pokušaj je pao na ograničenju od 600 s za CEO poziv run_vba (crveni RunAllTests 210 s + BFP); BFP je ponovljen sam, sa --timeout 900".
    DOKAZ: [A]/tools/run_vba.py:705-706, 884-886; [A]/docs/STANJE_REFAKTORA.md:10525-10526
- [M] run_vba.py je deo test-ugovora: svaka izmena fajla menja otisak `runner`, pa svi upisi u tests/last_green.json zastarevaju i traže nov prolaz. self_test uz to traži TAČAN tekst dva poziva u main() (snimi_kontekst i zabelezi_prolaz), pa refaktor main() lomi self-test.
    DOKAZ: [A]/tools/vba_gate.py:141-145, 513-546; [A]/tools/run_vba.py:446-459
- [M] Katalog sabotaža po vrsti fajla: 1258 nad .bas, 6 nad clsTransaction.cls (meka klasa), 1 nad ThisWorkbook.doccls, 0 nad .frm i 0 nad pet tvrdih klasa. Prosečan mutirani fajl je 170 kB (medijana 111 kB), tj. oko 1,6% izvora; po suiti: RunAllTests 801, BFP 456, banka 8.
    DOKAZ: python3 (uvoz [A]/tools/sabotaza.py kao modula, samo čitanje SABOTAZE) -> "po ekstenziji: {'.bas': 1258, '.cls': 6, '.doccls': 1}; nad .frm: {}; sabotaze nad tvrdim klasama: {}; prosecna velicina fajla po sabotazi: 170 kB; po suiti: {RunAllTests: 801, RunBusinessFlowProSuite: 456, RunBankaImportTestSuite: 8}"
- [M] dokaz.py pušta jedan run_vba potproces po prolazu i u njemu JEDNU suitu (onu u kojoj test živi); bazu pušta zasebnim pozivom po suiti i uvek dodaje RunAllTests. CLAUDE.md tvrdi da se po svakoj sabotaži vrte RunAllTests i BFP -- to se ne slaže sa kodom.
    DOKAZ: [A]/tools/dokaz.py:344-349, 466-468, 92-94; [A]/CLAUDE.md paragraf 5 ("po svakoj sabotaži vrti RunAllTests i BFP -- oko 3,5 minuta po stavci")
- [M] Nije bezbedno za paralelan rad u istom stablu: tests/last_run.json i last_run.txt na fiksnoj putanji; tests/last_green.json se čita-menja-piše bez brave; golden se na kraju svakog poziva kopira nazad u repo; SendKeys/AppActivate cilja naslov VBE prozora koji je isti za sve instance (isto ime temp sveske); SRC_VBA je konstanta; jedan test koristi fiksni folder %TEMP%\agrix_backup_test sa Kill/RmDir.
    DOKAZ: [A]/tools/run_vba.py:1065-1067, 1123, 844, 1031, 625-635, 530-533, 96; [A]/tools/vba_gate.py:687-709, 616-620; [A]/src-vba/modTest.bas:17763-17779
- [M] Bezbedno za paralelan rad: watchdog filtrira po PID-u, killer i završni TerminateProcess ciljaju svoj PID, temp folder je mkdtemp, Excel je zaseban proces (DispatchEx). U src-vba nema upotrebe clipboard-a, AddressOf, SetTimer ni Windows hook-ova.
    DOKAZ: [A]/tools/run_vba.py:197, 470-479, 837, 875-876, 1022-1027; grep clipboard/AddressOf/SetTimer/SetWindowsHookEx u [A]/src-vba -> bez pogodaka
- [M] Stanje koje preživljava zatvaranje sveske u istom Excel procesu: Application.OnTime registracije (žive u Excelu, ne u VBA) -- 25 izvršnih mesta u 10 modula, među njima modOtkupUI HideToast, modJournaling AutoSaveTick, modStanicaLock, modStornoWarm; i Application zastavice -- modTest.bas ih postavlja na 13 mesta, a suite ume da vrati DisplayAlerts na True.
    DOKAZ: [A]/docs/SELF_UPDATE.md:540-547 (zamka 29); grep -c '^[^']*Application\.OnTime' [A]/src-vba -> modAdmin 1, modGoogleSyncOrchestrator 3, modJournaling 2, modLicense 1, modMain 2, modOtkupUI 4, modSelfUpdate 2, modStanicaLock 2, modStornoWarm 2, modVbaTools 5; grep Application.* setera u modTest.bas -> 13; [A]/tools/run_vba.py:981-985
- [M] Podaci u svesci ne preživljavaju bezbedno drugi nalog: BFP namerno ostavlja test redove (TST-PRO / *-9000x), a Palete suita odbija da krene kad zatekne TST- ostatke. FULL prolaz već danas vrti 12 suita u istoj svesci bez ponovnog otvaranja, pa suite u FULL i u `--suite X` polaze iz različitog stanja.
    DOKAZ: [A]/src-vba/modBusinessFlowProTests.bas:3077-3087; modTestPalete.bas:83-91; [A]/tools/run_vba.py:944-972 (jedna sveska za sve suite)
- [M] Zamke zamene modula koje kod i dokumentacija već beleže: VBComponents.Remove je odložen dok makro radi; forme se nikad ne smeju Remove+Import; tvrdi moduli idu isključivo kroz Import; više DeleteLines+AddFromString nad KOMPAJLIRANIM projektom u jednom makrou ruši VBE (mereno 14.09.2026); izmena bilo kog modula briše module-level stanje svih modula (mereno 07.09.2026); Err=0 posle AddFromString nije dokaz da je telo primenjeno; Import ume da napravi modX1.
    DOKAZ: [A]/src-vba/modVbaTools.bas:14-18, 44-48, 60-67, 72-74; [A]/docs/SELF_UPDATE.md:160-181 (zamke 1-4), 406-419 (25), 521-538 (29); [A]/docs/EXCEL_TEST_HARNESS.md:446-454
- [M] Za proveru da je modul stvarno zamenjen postoji gotova Python referenca poređenja koda pod paritetnom kapijom sa VBA kopijama: ref_same_code, ref_canon, ref_is_hard.
    DOKAZ: [A]/tools/vba_parity_check.py:148, 207, 215
- [M] Grana B: tools/run_vba.py, tools/vba_gate.py, tools/dokaz.py, tools/vba_check.py i docs/EXCEL_TEST_HARNESS.md nisu menjani na main-u od baze d19c2da, pa se Python strana grane primenjuje bez sudara. Probno spajanje u zasebnom klonu daje sudare samo u 3 fajla: modBusinessFlowProTests.bas (2 bloka), modTest.bas (3 bloka), sabotaza.py (1 blok).
    DOKAZ: git -C [B] diff --quiet d19c2da origin/main -- <fajl> -> NEPROMENJEN za svih pet; git merge-tree --write-tree --name-only main-novi origin/claude/dazzling-sagan-3o531f (u klonu u scratchpad-u, posle obrisanom) -> CONFLICT samo u ta tri fajla; grep -c '^<<<<<<<' -> 2, 3, 1
- [M] Grana B dodaje test pod indeksom 311, a na main-u je 311 već zauzet (T_S8R1b_PrevezivanjeTraziObaID); main ima 350 RunOne poziva. BFP telo na main-u ima 391 poziv testa, grana B ih je omotala 361 (360 iz baze + svoj nov test), dakle 31 poziv sa main-a nije omotan u `If BfpIzabran(...) Then`.
    DOKAZ: [A]/src-vba/modTest.bas:1083, 1441; grep -c '^\s*RunOne [0-9]' -> 350; brojanje poziva u telu RunBusinessFlowProSuite: d19c2da 360, main 391, grana B 361; git -C [B] diff HEAD~1 HEAD -- src-vba/modTest.bas (Case 311)
- [M] Grana B u run_vba.py dodaje: --test (ulazna tačka iz ključa `izbor`, spisak kao jedan string argument kroz xl.Run), problem_izbora pre provere platforme, čitanje <rezultat>_vreme.txt i blok VREME (10 najsporijih), oznaku entry["izbor"] po kojoj vba_gate ne piše marker, i statičke provere wire-up-a u self-testu. Ne meri faze poziva (uvoz, compile, start Excela).
    DOKAZ: git -C [B] diff HEAD~1 HEAD -- tools/run_vba.py (252 linije); tools/vba_gate.py (izbori -> 'marker nije upisan')
- [M] Pomoćni Python koraci po prolazu dokaza su zanemarljivi: otisak celog src-vba 0,012 s, kopija celog src-vba (195 fajlova, 10,4 MB) 0,012 s, učitavanje sabotaza.py --lista 0,07 s (Linux).
    DOKAZ: python3 merenje nad [A]/src-vba -> 'dokaz._otisak ekvivalent: 0.012 s; kopija celog src-vba: 0.012 s'; time python3 tools/sabotaza.py --lista -> real 0m0.071s
- [N] Procena raspodele ~18 s režije (iz koda, bez merenja): start Excela 1,5-3 s; Open 0,5-1,5 s; uvoz 3-7 s; compile korak 3-7 s (od toga 2,5 ili 5,9 s fiksno); EnsureRuntimeSchema 0,5-3 s; zatvaranje i gašenje 2,3-3 s (od toga 2,0 s fiksno); Python ostatak ispod 1 s.
    DOKAZ: procena iz [A]/tools/run_vba.py (linije faza gore) i dokumentovane ~18 s u [A]/docs/EXCEL_TEST_HARNESS.md:732; nijedna vrednost nije merena u Excelu

## TROSKOVI
[
 {
  "sta": "Režija jednog run_vba poziva (sve osim same suite)",
  "koliko": "~18 s (dokumentacija, docs/EXCEL_TEST_HARNESS.md:732); nije poznato da li uključuje demand-compile koji se plaća unutar vremena suite",
  "izmereno": false,
  "koliko_cesto": "svaki poziv: 215 puta po grupnom dokazu (64,5 min), 1265 puta po pojedinačnom (6,3 h), plus 2-3 bazna poziva"
 },
 {
  "sta": "Fiksni sleep-ovi u compile koraku",
  "koliko": "2,5 s (samo COM pokušaj) ili 5,9 s (COM + SendKeys: 2,5 + 0,5 + 0,4 + 2,5)",
  "izmereno": true,
  "koliko_cesto": "svaki poziv, i kad se puštaju suite kod kojih verdikt ne utiče na rc"
 },
 {
  "sta": "Fiksni sleep-ovi na gašenju",
  "koliko": "2,0 s (1,0 pre watchdog.stop + 1,0 pre TerminateProcess)",
  "izmereno": true,
  "koliko_cesto": "svaki poziv"
 },
 {
  "sta": "Uvoz src-vba: 43 merge-a + 149 Import-a, 10,4 MB",
  "koliko": "procena 3-7 s",
  "izmereno": false,
  "koliko_cesto": "svaki poziv, iako se u mutacionom prolazu razlikuje 1-6 fajlova (prosečno 1,6% bajtova)"
 },
 {
  "sta": "Start novog Excel procesa + Open temp sveske",
  "koliko": "procena 2-4,5 s",
  "izmereno": false,
  "koliko_cesto": "svaki poziv"
 },
 {
  "sta": "EnsureRuntimeSchema (uključuje demand-compile modSetup/modSchema i zavisnosti)",
  "koliko": "procena 0,5-3 s",
  "izmereno": false,
  "koliko_cesto": "svaki poziv sa suitama"
 },
 {
  "sta": "Demand-compile sakriven u trajanju suite (modTest i BFP povlače skoro ceo projekat)",
  "koliko": "nepoznato; moguće reda veličine samog uvoza ili više",
  "izmereno": false,
  "koliko_cesto": "svaki poziv koji pušta RunAllTests ili BFP, i uz --test"
 },
 {
  "sta": "Čekanje watchdog-a po dijalogu",
  "koliko": "prosečno 0,2 s, najviše 0,4 s po dijalogu (poll 0,4 s); 11 dijaloga po FULL prolazu = ~2,2 s; BFP prolaz 1 dijalog; banka 2",
  "izmereno": false,
  "koliko_cesto": "FULL: jednom; dokaz: po BFP/banka prolazu (~0,2-0,4 s)"
 },
 {
  "sta": "Visenje do tvrdog prekida kad mutacija obori compile ili Excel uđe u break",
  "koliko": "600 s po slučaju, pa u grupi još i ponavljanje svakog člana pojedinačno",
  "izmereno": true,
  "koliko_cesto": "retko (statički čuvari u sabotaza.py), učestalost nije izmerena"
 },
 {
  "sta": "Lažan FATAL kad ceo poziv pređe 600 s",
  "koliko": "ceo prolaz propada i ponavlja se sa --timeout 900",
  "izmereno": true,
  "koliko_cesto": "dogodilo se bar jednom (STANJE_REFAKTORA.md:10525); učestalost raste sa brojem testova"
 },
 {
  "sta": "Invalidacija markera zelenog pri svakoj izmeni tools/run_vba.py (deo ugovora)",
  "koliko": "jedan FULL prolaz da se marker povrati",
  "izmereno": true,
  "koliko_cesto": "po svakom PR-u koji dira run_vba.py, vba_gate.py ili make_fixture.py"
 },
 {
  "sta": "Bazni pozivi dokaza: svaka potrebna suita + uvek RunAllTests, svaka u zasebnom pozivu",
  "koliko": "~100 s (RunAllTests) + ~170 s (BFP) po dokaz-pozivu, po dokumentovanim trajanjima",
  "izmereno": false,
  "koliko_cesto": "jednom po dokaz-pozivu; u rezu sa 1-3 prolaza to je 40-60% trajanja"
 },
 {
  "sta": "Snimak konteksta i marker (heš 10,4 MB izvora + sveska, dvaput)",
  "koliko": "0,045 s po snimku (Linux)",
  "izmereno": true,
  "koliko_cesto": "svaki poziv; u mutacionim prolazima marker se ionako ne piše"
 },
 {
  "sta": "Održavanje: tačan tekst poziva u main() zakucan u self_test, i 391 BFP poziv koji (posle grane B) mora da bude omotan",
  "koliko": "svaki refaktor main() traži izmenu igala i nov dvosmerni dokaz; svaki nov BFP test se piše u omotanom obliku",
  "izmereno": true,
  "koliko_cesto": "po refaktoru runnera; po novom BFP testu"
 }
]

## PREDLOZI
### RV-01 Vreme po fazi u izveštaju (instrumentacija) [trud S] [slabi: NE]
Kontekst-menadžer oko svake faze u main(): potpis fixture-a, kopija, snimak, start Excela, Open, uvoz .doccls, uvoz modula, orphan, compile COM, compile SendKeys, probe, schema, svaka suita, Close, Quit, marker. Upis u report["faze"] i redovi `FAZA <ime> <s>` plus `REZIJA <zbir bez suita>` u last_run.txt/json; u uvozu 10 najsporijih fajlova; zapis verzije i build-a Excela. Uz granu B dodati `VAN TESTOVA = trajanje suite - zbir VREME` (skriveni demand-compile, seed, čekanje dijaloga). Bez ijedne izmene koda već danas se može: `time python tools/run_vba.py --compile-only`, `--suite TestLicense_All`, i čitanje ključeva com_*/keys_* iz postojećeg tests/last_run.txt (pokazuje da li se plaća 2,5 ili 5,9 s).
  DOBITAK: 0 s direktno. Jednim prolazom na Windows-u (3 poziva: --compile-only, --suite RunAllTests, --suite RunBusinessFlowProSuite, oko 6 min) sve procene iz ovog izveštaja postaju merenja, pa se ostali predlozi rangiraju po stvarnim brojevima.
  RIZIK: Nizak. Menja run_vba.py, dakle otisak `runner`, pa markeri u last_green.json zastarevaju (jedan FULL).
  PREDUSLOVI: Windows + Excel za samo merenje. Spojiti u isti PR sa RV-02/03/04 da se invalidacija markera plati jednom.
  DOKAZATI: Self-test bez Excela: lažan report sa fazama -> _write_report ispisuje FAZA redove i REZIJA jednak zbiru; dvosmerno: izbaci jednu fazu iz zbira -> self-test pada po imenu. Na Windows-u: zbir FAZA + zbir suita se slaže sa ukupnim trajanjem poziva na +-0,5 s.

### RV-02 Fiksna čekanja zameniti čekanjem na događaj [trud S] [slabi: NE]
(a) _attempt_compile: umesto sleep(2.5) petlja do 2,5 s sa korakom 0,1 s koja izlazi čim ctl.Enabled postane False ili watchdog zabeleži dijalog; bez pozitivnog signala čeka pun budžet kao danas. (b) Linija 1012: umesto sleep(1.0) jedan sinhroni obilazak prozora pa stop + join. (c) Linija 1021: WaitForSingleObject nad procesom do 3 s, TerminateProcess samo ako je i dalje živ. (d) Jedno listanje imena komponenti umesto 149 COM izuzetaka u uvozu i ponovnog obilaska u report_orphan_components.
  DOBITAK: Procena 1,5-4 s po pozivu: do ~2,4 s kad COM compile uspe, ~0,9 s i ~0,5-0,9 s na gašenju. Na 215 prolaza 5-14 min; na 1265 prolaza 0,5-1,4 h.
  RIZIK: Nizak do srednji. Prerani izlaz bi mogao da promaši kasni compile dijalog, zato se izlazi samo na pozitivan signal. Excel koji ne izađe za 3 s i dalje se ubija.
  PREDUSLOVI: RV-01, da se dobitak vidi po fazi.
  DOKAZATI: Self-test sa lažnim ctl i watchdog objektima u tri smera (Enabled rano False; dijalog rano; ništa -> pun budžet), svaki pada po imenu kad se uslov ukloni. Windows: --compile-only nad čistim izvorom i nad namerno slomljenim modulom daje isti verdikt kao pre izmene; 20 uzastopnih poziva bez zaostalog EXCEL.EXE.

### RV-03 Compile korak samo tamo gde nešto tvrdi [trud S] [slabi: NE za rc i marker. Menja se dijagnostika: u suite režimu se gubi rani eksplicitni COMPILE FAIL iz drugog pokušaja; nadoknađuje ga RV-04, a FULL i ručni Compile pred release ostaju isti.]
Kad se puštaju suite (nije --compile-only): ne raditi SendKeys pokušaj (run_vba.py:575-577) ni probe. Za mutacione prolaze dokaz.py prosleđuje zastavicu koja preskače ceo compile_project. --compile-only i FULL (podrazumevani set) ostaju netaknuti. Osnov u kodu: NEJASNO ne utiče na rc (914-918), compile se ne piše u marker (1046-1048), a korak je izmereno remetio Excel (929-934).
  DOBITAK: 3,4 s po pozivu na mašini gde COM Execute ne uradi ništa, plus ~0,3-0,5 s probe; sa potpunim preskakanjem 2,5-5,9 s plus otvaranje VBE prozora. Na 215 prolaza oko 18 min, na 1265 oko 1,75 h (uz ~5 s po prolazu). Sporedno: nema krađe fokusa ni SendKeys-a dok operater radi na istoj mašini; preduslov za paralelne instance.
  RIZIK: Compile-slomljen izvor se u suite režimu vidi kasnije i mutnije (dijalog tokom xl.Run, 'Exception occurred' ili visenje) umesto ranog COMPILE FAIL. Zato ide zajedno sa RV-04.
  PREDUSLOVI: RV-04. Iz postojećeg tests/last_run.txt pročitati koji se put danas stvarno izvršava.
  DOKAZATI: Windows: (1) čist izvor, `--suite RunIzvestajTests` pre i posle: isti rc i isti SUITE/TESTS redovi, faza compile kraća; (2) namerno slomljen modul: rc=2 pre i posle, a posle se u izlazu vidi tekst Compile error dijaloga; (3) --compile-only daje iste detail ključeve kao pre.

### RV-04 Timeout po suiti i brz prekid na compile/runtime dijalog [trud M] [slabi: NE]
(a) Tajmer se ponovo naoružava pred svaku suitu, sa zasebnim budžetom za uvoz i compile; dokaz.py prosleđuje budžet izveden iz izmerenog trajanja baze (npr. 3x + 60 s) umesto 600 s. (b) Kad watchdog tokom xl.Run suite vidi dijalog čiji tekst odgovara istim ključnim rečima kao na liniji 583-584, a Run se ne vrati za nekoliko sekundi, proces se ubija i prijavljuje se FATAL sa imenom suite i tekstom dijaloga.
  DOBITAK: Uklanja lažni FATAL kad ceo poziv pređe 600 s (već se desilo). Visenje mutacije koja obori compile pada sa 600 s na ~10 s po slučaju, a izbegava se i do 6 ponavljanja po 600 s kad visi grupa. Učestalost nije izmerena.
  RIZIK: Prekratak budžet na sporoj mašini daje lažan TIMEOUT; on je fail-closed (nalaz, nikad dokaz).
  PREDUSLOVI: Trajanja baze iz RV-01 ili iz bloka VREME grane B.
  DOKAZATI: Self-test sa lažnim satom i lažnom petljom suita: tajmer se naoružava po suiti; dvosmerno: vrati jedan globalni tajmer -> self-test pada. Windows: suita sa veštačkim čekanjem preko budžeta -> FATAL imenuje tu suitu; slomljen modul -> prekid ispod 15 s sa tekstom dijaloga.

### RV-05 `--src` i `--izlaz`: mutacija nad kopijom izvora [trud M] [slabi: NE]
SRC_VBA (run_vba.py:96) i odredište izveštaja (1065) postaju parametri; za tuđ --src marker se ne piše (isto pravilo kao --no-import). dokaz.py mutira KOPIJU src-vba (kopija celog stabla je izmerena na 0,012 s na Linuxu) i pušta run_vba nad njom. Radno stablo se nikad ne prlja: nestaje klasa REVERT-FAIL, ne treba --vrati, hook i razvoj rade dok dokaz traje, a paralelni delovi postaju mogući u jednom stablu.
  DOBITAK: 0 s po prolazu. Uklanja klasu kvara (namerno pokvaren radni izvor posle prekida) i otključava RV-06 i RV-09.
  RIZIK: sabotaza.py primena i vraćanje rade nad ROOT/src-vba pa im treba ciljni direktorijum (oblast drugog alata). Provera 'promenilo se TOKOM run-a' u vba_gate više ne vidi mutaciju, što je ispravno jer mutacija nije u stablu.
  PREDUSLOVI: Dogovor sa vlasnikom dokaz.py i sabotaza.py.
  DOKAZATI: dokaz.py --self-test proširen: orkestracija ne dira ROOT/src-vba (otisak isti pre, tokom i posle prolaza); dvosmerno: vrati mutaciju u stablo -> self-test pada. Windows: isti skup od 20 sabotaža starim i novim putem daje identične verdikte po imenu.

### RV-06 Jednokratni uvoz po dokaz-pozivu + delta zamena modula (varijante a i c, uži oblik) [trud M] [slabi: NE za marker (delta put ga ne piše) ni za verdikt dokaza (baza i mutacija dele istu uvezenu svesku, razlika je samo delta). Verdikt treba da nosi oznaku puta, a pun dokaz pred release ostaje hladan dok ekvivalencija ne bude izmerena nad celim katalogom.]
Na početku dokaza run_vba jednom uveze čist src-vba u kopiju fixture-a i snimi je (SaveCopyAs) uz manifest: sha256 po fajlu, heš fixture-a, build Excela, heš snimljene sveske. Svaki prolaz: proveri heš uvezene sveske, kopiraj je u nov temp folder, otvori, izračunaj deltu između manifesta i mutirane kopije izvora, zameni samo te module istim putevima koje uvoz već koristi (Remove+Import za .bas/.cls, DeleteLines+AddFromString za .doccls), proveri ime i broj komponenti i pročitaj kod nazad kroz vba_parity_check.ref_same_code, pa EnsureRuntimeSchema i suita. Sveska sa mutacijom se uvek baca, nikad se ne 'vraća'. Delta koja sadrži .frm/.frx, tvrdu klasu, nov ili obrisan fajl se odbija i ide hladnim putem (danas 0 takvih sabotaža). Keš živi koliko jedan dokaz-poziv, pa nema problema ustajalosti između sesija.
  DOBITAK: Uvoz 192 komponente -> 1-6 modula po prolazu: procena -3 do -7 s, plus deo skrivenog demand-compile-a ako snimljena sveska zadrži kompajlirano stanje (nepoznato). Sa RV-02/03 režija prolaza ~18 -> ~6-8 s: na 215 prolaza ~65 -> ~25 min, na 1265 prolaza 6,3 h -> ~2,5 h. Uz izbor testova (grana B), gde je prolaz približno režija + nekoliko sekundi, to je oko 2-3x na ceo dokaz.
  RIZIK: Srednji. Nije izmereno: da li je Remove preko spoljnog COM-a trenutan nad kompajliranim projektom (u makrou je odložen); da li više zamena zaredom nad kompajliranom sveskom ponavlja pad VBE-a zabeležen za merge u jednom makrou; da li Import ume da napravi modX1.
  PREDUSLOVI: RV-05 i RV-01. Spike na Windows-u: 50 puta zamena 6 modula nad uvezenom, kompajliranom sveskom bez pada i bez duplikata.
  DOKAZATI: (1) Ekvivalencija: isti skup sabotaža hladno i delta putem -> isti verdikt po imenu. (2) Negativna kontrola: prazna delta -> zeleno sa istim TESTS=N kao hladna baza. (3) Dvosmerno nad alatom: ugasi primenu delte -> svi članovi 'NE OBARA NIŠTA'; izmeni jedan bajt u uvezenoj svesci -> odbijena po hešu; pokvari read-back -> pad po imenu modula. (4) Član koji delta putem nije dokazan meri se hladno, isto pravilo kao grupa -> pojedinačno.

### RV-07 Trajni keš uvezene sveske po otisku za TARGETED petlju (varijanta a, puni oblik) [trud M] [slabi: DA, umereno: tvrdnju markera 'izvršen je baš ovaj src-vba, sveže uvezen' menja u 'izvršen je sadržaj identičan po hešu, uvezen ranije'. Ublažava se oznakom u markeru i zahtevom za svež uvoz u FULL-u pred release. Nizak prioritet: RV-06 daje veći deo dobiti bez ovog slabljenja.]
tests/.cache/uvoz_<ključ>.xlsm; ključ = sha256 nad manifestom po fajlu (uključujući .frx i modBuildInfo), hešom fixture-a, hešom run_vba.py i build-om/bitnošću Excela. Pogodak preskače uvoz; promašaj sa postojećim manifestom radi delta uvoz pa snima nov keš. Marker beleži da li je uvoz bio svež ili iz keša, a `--require-green` pred release traži svež.
  DOBITAK: Procena -3 do -7 s po TARGETED pozivu kad izvor nije menjan (ponovljene suite), neto -2 do -4 s kad su menjana 1-3 modula jer se plaća snimanje keša.
  RIZIK: Srednji do visok: ustajao keš znači da izvršen kod nije src-vba, tj. lažno zeleno. Keš postaje nov ulaz u ugovor; troši disk.
  PREDUSLOVI: RV-06 u pogonu i izmeren; RV-01.
  DOKAZATI: Self-test ključa: svaki ulaz menja ključ (dvosmerno, po ulazu). Windows: izmeni modul -> promašaj ili delta -> test nad izmenjenim ponašanjem pada; poziv bez keša daje iste rezultate.

### RV-08 Trajna Excel sesija za više naloga (varijanta b) [trud L] [slabi: DA ako bi sesija pisala marker zelenog (izolacija po procesu je deo onoga što 'zeleno' znači), zato sesija marker NE piše. Za verdikt dokaza slabljenje je ograničeno kanarincem i padom na hladan put, ali tvrdnja postaje 'izmereno u deljenom procesu' i tako treba da se zove.]
Klasa sesije u run_vba.py; CLI ostaje sesija sa jednim nalogom, da hladan put bude isti kod. dokaz.py drži jednu sesiju za K prolaza. Između naloga: zatvori sve sveske bez snimanja; proveri invarijante (nema otvorenih sveski, nema #32770 prozora, nijedan dijalog viđen između naloga); ponovo postavi EnableEvents=False, DisplayAlerts=False, AutomationSecurity; nov temp folder; kopija uvezene sveske iz RV-06; delta; schema; suita. Proces se reciklira na svakih N naloga, na svaki timeout i na svaki dijalog van naloga. Kanarinac: na svakih M naloga i na kraju ista baza u istoj sesiji mora biti zelena sa istim TESTS=N; ako nije, svi nalozi od poslednjeg dobrog kanarinca se poništavaju i mere hladno.
  DOBITAK: Preostala režija posle RV-06 (start Excela i gašenje): procena -3 do -5 s po prolazu, režija ~6-8 -> ~2-4 s. Ima smisla tek uz izbor testova, kad je režija većina prolaza.
  RIZIK: Visok. Preživljava zatvaranje sveske: Application.OnTime (25 mesta u 10 modula) -- zaostao tik ponovo otvara zatvorenu svesku ili diže dijalog u tuđem nalogu; Application zastavice (EnableEvents=True pre Open pokreće Workbook_Open i modalni start); break mode ubija sesiju; curenje memorije. Refaktor main() lomi zakucane igle self-testa (446-456) i traži nov dvosmerni dokaz.
  PREDUSLOVI: RV-04, RV-05, RV-06, grana B; merenje koje pokazuje da start i gašenje Excela i dalje nose više od ~30% prolaza.
  DOKAZATI: Ekvivalencija nad celim katalogom (sesija naspram hladnog, verdikt po imenu). Test zagađenja: nalog koji namerno ostavi OnTime, EnableEvents=True ili otvoren dijalog -> sledeći nalog mora biti odbijen i sesija reciklirana; dvosmerno: ugasi proveru -> self-test pada. Namerno pokvaren kanarinac -> poništavanje i hladno ponavljanje.

### RV-09 Više Excel instanci paralelno (varijanta d) [trud M] [slabi: NE, pod uslovom da FULL i upis markera ostanu serijski; svaki prolaz je i dalje zaseban proces nad svojim izvorom.]
N delova dokaza, svaki nad svojom kopijom izvora (RV-05) ili u svom git worktree-u (bez izmene koda, uz kopiju fixture-a i .sig). Potrebno: bez SendKeys i VBE fokusa (RV-03); izveštaj po delu (`--izlaz`); marker se u delu ne piše, ili se piše atomski pod bravom; golden se iz dela ne vraća u repo (nov golden je nalaz); test sa fiksnim %TEMP%\agrix_backup_test (modTest.bas:17763-17779) prebaciti pod folder sveske; spajanje rezultata uz tvrdnju da je svaka sabotaža izmerena tačno jednom.
  DOBITAK: Približno min(N, broj jezgara) puta na ukupno trajanje, uključujući i trajanje suite -- jedina poluga koja skalira deo koji nije režija (80-150 s po prolazu). Za N=4 grupni dokaz od oko 7 h -> oko 2 h (procena, uz dovoljno jezgara i memorije).
  RIZIK: Srednji: testovi osetljivi na vreme pod opterećenjem (dva Application.Wait od 1 s, modTest.bas:10415 i 10465); memorija po instanci nepoznata; operater radi na istoj mašini; deljeni resursi koje nisam našao.
  PREDUSLOVI: RV-03 i RV-05 (ili worktree varijanta); broj jezgara i RAM Windows mašine.
  DOKAZATI: Isti skup serijski i paralelno (N=2, pa 4) -> identični verdikti po imenu, tri ponavljanja bez razlike. Dvosmerno: dva dela namerno dele isti `--izlaz` -> alat mora da odbije umesto da pregazi.

### RV-10 Dijalozi: sprovesti 'neočekivan dijalog = pad' i ukloniti MsgBox iz headless prolaza [trud S] [slabi: NE -- jača: dokumentovano pravilo postaje izvršeno.]
(a) U run_vba: za suitu sa dialogs:False svaki dijalog viđen tokom njenog Run-a je FAIL; za dialogs:True dozvoljen je samo očekivan broj i obrazac (potvrda + rezime). (b) U VBA: potvrdni i završni MsgBox u 7 suita preskočiti kad nema operatera (Excel je u run_vba nevidljiv, linija 888), pa sve suite postaju dialogs:False.
  DOBITAK: Vreme je zanemarljivo: najviše ~2 s po FULL prolazu i ~0,2 s po BFP prolazu (oko 40 s na 215 prolaza). Stvarna vrednost je pouzdanost: zatvara mogućnost da watchdog klikne End na runtime-error dijalogu u suiti bez rezultat-fajla i ona ostane OK, i uklanja zavisnost od jezika dugmadi.
  RIZIK: (b) dira 7 test modula; dev-guard 'Pokrenuti?' štiti klijenta od slučajnog pokretanja, pa se preskače samo kad je aplikacija nevidljiva. Važi ASCII pravilo.
  PREDUSLOVI: Za (a) ništa; za (b) Windows prolaz.
  DOKAZATI: (a) Self-test sa lažnim izveštajem: dijalog u suiti sa dialogs:False -> rc=2 po imenu suite; dvosmerno. Windows: suita sa ubačenim MsgBox-om -> FAIL. (b) FULL prolaz bez ijednog reda DIALOG.

### RV-11 Rebase grane B: šta prepraviti u runneru i oko njega [trud M] [slabi: NE: prolaz sa izborom ne piše marker i baza se tvrdi nad istim izborom. Ostaje ono što grana sama kaže: izbor nije dokaz cele suite.]
1) Python strana se uzima cela sa grane (run_vba.py, vba_gate.py, dokaz.py, docs/EXCEL_TEST_HARNESS.md) -- na main-u nisu menjani od baze. 2) modTest.bas: nov test prenumerisati sa 311 na 351 u RunOne/TestName/InvokeTest; ostatak (RunOne sa TestIzabran i merenjem, PreuzmiIzbor, ProveriIzbor, WriteVremeFile) preneti u tri sudarena bloka. 3) BFP: ne rešavati sudar rukom -- uzeti telo sa main-a (391 poziv) i skriptom omotati svaki poziv u `If BfpIzabran("Test_X") Then Test_X`; provera golih poziva u run_vba --self-test to čuva. 4) sabotaza.py: četiri nove stavke na kraj kataloga, pa --proveri-sidra. 5) U run_vba.py pri rebase-u dodati FAZA i VAN TESTOVA (RV-01) i timeout po suiti (RV-04); ispraviti docstring (red `--workbook` je ostao ispod pasusa o izboru). 6) Prvi Windows prolaz: --compile-only, pa dokaz.py --izolacija nad rezom, pa --izbor; posebno proveriti spisak testova duži od 255 znakova kroz Application.Run.
  DOBITAK: Otključava izbor testova: prolaz suite od 80-150 s postaje nekoliko sekundi + režija. Tada režija nosi oko 80-90% prolaza, što je razlog za RV-02, RV-03 i RV-06.
  RIZIK: VBA strana grane nikad nije izvršena u Excelu; testovi koji sami ne prolaze (zavisnost od redosleda); skriveni demand-compile ograničava dobitak izbora jer modTest povlači skoro ceo projekat.
  PREDUSLOVI: Windows + Excel; FULL prolaz posle spajanja jer se menja ugovor (runner i kapija).
  DOKAZATI: Linux: vba_check, run_vba --self-test, vba_gate --self-test, dokaz --self-test i sabotaza --proveri-sidra zeleni. Windows: ručni Compile, FULL, --izolacija; dokaz nad četiri sabotaže izbor-* i bfp-izbor-* crven po imenu.

### RV-12 Baza dokaza iz markera zelenog, ili sve bazne suite u jednom pozivu (granična oblast sa dokaz.py) [trud S] [slabi: DELIMIČNO: tvrdnju 'baza je zelena PRE prve mutacije, u ovom pozivu' zamenjuje 'baza je dokazano zelena nad identičnim otiscima, ranije'. Treba da ostane opcija; sveža baza ostaje podrazumevana pred release. Varijanta sa jednim pozivom ne slabi ništa.]
dokaz.py pre mutacija pušta bazu za svaku potrebnu suitu i uvek RunAllTests, svaku u zasebnom run_vba pozivu (dokaz.py:466-468). Ako vba_gate već potvrđuje da je baš ovaj izvor, ugovor i sveska zelen za tu suitu, baza se preuzima iz markera; inače se sve bazne suite puste u JEDNOM pozivu (`--suite A --suite B`).
  DOBITAK: Oko 100-270 s po dokaz-pozivu po dokumentovanim trajanjima; u rezu sa 1-3 prolaza to je 40-60% ukupnog trajanja. CLAUDE.md ionako traži da ciljana suita bude zelena neposredno pre dokaza.
  RIZIK: Marker ne hvata nedeterminizam između dva puštanja (flaky ili vremenski zavisan test), a sveža baza hvata. Više suita u jednom pozivu zavisi od RV-04 zbog globalnog timeouta.
  PREDUSLOVI: RV-04 za varijantu sa jednim pozivom.
  DOKAZATI: dokaz --self-test: važeći marker -> nula baznih poziva; marker sa drugim otiskom ili palom suitom -> bazni poziv se izvrši (dvosmerno).


## NEPOZNANICE
[
 "Stvarno trajanje svake faze na Windows mašini: sve vrednosti za Excel su procene iz koda. Meri se jednim prolazom posle RV-01 ili, bez izmene koda, sa `time python tools/run_vba.py --compile-only` i `--suite TestLicense_All`.",
 "Koji se compile put danas stvarno izvršava na korisnikovoj mašini (COM Execute uspe, ili se ide i na SendKeys), tj. da li se plaća 2,5 ili 5,9 s. Vidi se iz ključeva com_*/keys_* u postojećem tests/last_run.txt.",
 "Kako je izmereno dokumentovanih ~18 s režije i da li uključuje demand-compile koji se plaća unutar trajanja suite. Ako ne uključuje, stvarna režija po prolazu je veća.",
 "Da li je VBComponents.Remove pozvan spolja preko COM-a (bez makroa na steku) trenutan nad kompajliranim projektom, i da li više zamena modula zaredom ponavlja pad VBE-a zabeležen 14.09.2026 za merge u jednom makrou.",
 "Da li sveska snimljena posle uvoza zadržava kompajlirano stanje pri ponovnom otvaranju u istom build-u Excela, i koliko traje Open sveske sa ~10 MB VBA izvora.",
 "Da li probe (dodavanje i uklanjanje modula posle compile-a) poništava kompajlirano stanje projekta pa se compile plaća ponovo pri prvom Run-u.",
 "Da li testovi u test-režimu ostavljaju zakazan Application.OnTime (npr. HideToast u modOtkupUI.bas:8308). Ključno za trajnu sesiju; nisam proverio garde.",
 "Šta xl.Run vrati kad watchdog klikne End na runtime-error dijalogu: izuzetak ili normalan povratak. Od toga zavisi da li suita bez rezultat-fajla može da ostane lažno OK.",
 "Da li Application.Run prima string argument duži od 255 znakova (spisak testova u --test iz grane B). Grupa od 6 testova sa imenima od ~40 znakova je na toj granici.",
 "Trajanje FULL prolaza danas i koliko je blizu granice od 600 s; veličina fixture-a; broj jezgara i RAM Windows mašine (za paralelne instance).",
 "Koliko često mutacija obori compile ili izazove visenje do timeouta, i koliko se članova grupe zbog toga ponavlja pojedinačno.",
 "Da li testovi dele još neki resurs van sveske osim %TEMP%\\agrix_backup_test (npr. podrazumevani korenski folder Documents\\OtkupApp iz modSetup.bas:2379).",
 "VBA strana grane B (ulazne tačke izbora, BfpIzabran, upis vremena) nikad nije izvršena u Excelu; koliko ciljanih testova ne prolazi samo, pokazaće tek `dokaz.py --izolacija`."
]
