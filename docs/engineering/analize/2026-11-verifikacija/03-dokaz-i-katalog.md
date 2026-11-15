# Dvosmerni dokaz: tools/dokaz.py i katalog tools/sabotaza.py (anatomija, model troška, trend, semantika, održavanje, slojevitost)

## SAZETAK
Oznake: A = <scratch>/main-wt (main @1fa4937), B = /home/user/otkupapp-pwa (grana izbora), S = <scratch>/dokaz-analiza (moje skripte; nijedno stablo nije dirano). (1) Katalog po tvrdnji nije redundantan (1264 različita para test+tvrdnja na 1265 sabotaža), ali pokriva samo oko 18% literala tvrdnji (1246 od 6803); redundantno je IZVRŠAVANJE: pojedinačan prolaz pusti celu suite od 350 ili 391 testa da bi posmatrao jednu tvrdnju. (2) Model: pun pojedinačni dokaz je 44 h (dokumentovana trajanja) do 82 h (trajanja izvedena iz dnevnika za početak novembra), --grupe 6 je 11–20 h, i oba rastu kvadratno (broj prolaza x trajanje suite): za 6 meseci 517–1025 h odnosno 129–256 h. (3) BFP je po dnevniku oko 350–380 s, ne dokumentovanih 150 s; dokaz.py ne prosleđuje --timeout, pa je run_vba limit od 600 s litica na 1–2 meseca rasta, a taj ishod bi se prijavio kao NE OBARA NISTA. (4) Nema traga da je pun katalog ikad pušten nad sadašnjom veličinom; rezultati dokaza se nigde mašinski ne beleže, a CLAUDE.md opisuje cenu po stavci drugačije nego što kod radi. (5) Najveća poluga koja ne dira release tvrdnju je izbor testa (grana B): rast postaje linearan, pun pojedinačni dokaz oko 10,5–13,5 h, tipičan rez sa 34 min na 3–5 min; posle toga 79% vremena je režija poziva, pa slede mutacija van radnog stabla, paralelne instance i jeftinija režija. (6) U izbor+grupe 6 keš baze praktično ne radi (208 baza na 212 prolaza), a pravila grupisanja nad celim katalogom ne koštaju ništa (215 prema 212) — jedina kočnica je cap. (7) Održavanje: +18,7 sabotaža dnevno, 64 presidrenja u 29 dana (42% merge koraka), fajl od 735 KB sa 1474 escape-ovana navodnika. (8) Tri jeftine režijske rupe su izmerene: 27 s po pozivu dokaz.py nad celim katalogom, 1,0 s po --vrati, 3,4 s od 4,0 s u --proveri-sidra (hook). (9) Uzorkovanje, rotacija i keš po zatvaranju poziva slabe release tvrdnju ako ZAMENE pun dokaz; mutant schemata mehanički pokriva samo 54% kataloga i menja ono što se meri, pa se ne preporučuje.

## CINJENICE
- [M] Katalog: 1265 sabotaža; RunAllTests 801, BFP 456, banka 8. 77 fajlova (75 .bas sa 1258 sabotaža, 1 .cls sa 6, 1 .doccls sa 1; nijedan .frm i nijedan test modul). Top 10 fajlova nosi 49% (modDokumenta 128, modAmbalaza 97, modOtkupUI 88, modScrDokumenti 58, modIzvestaj 55).
    DOKAZ: python3 S/anat.py A (modul učitan preko importlib, sys.dont_write_bytecode) -> 'po suite: {RunAllTests: 801, RunBusinessFlowProSuite: 456, RunBankaImportTestSuite: 8}', 'top 10 fajlova = 621 (49%)'
- [M] Po proceduri (dokaz._procedura): 695 različitih procedura, nijedno sidro nenađeno, 5 sabotaža u deklaracijama. 435 procedura ima 1 sabotažu, 129 ima 2, 41 ima 5 ili više, 4 imaju 10 ili više; max 12 (modNovacUnos.bas::ReversValidiraj), zatim 11 (modDokUnos.bas::PrijemnicaValidiraj).
    DOKAZ: S/anat.py -> 'histogram: [(1, 435), (2, 129), (3, 57), (4, 33), (5, 17), (6, 8), (7, 5), (8, 5), (9, 2), (10, 2), (11, 1), (12, 1)]'
- [M] Po testu: 551 različit test (318 + 227 + 6). Tačno 1 sabotažu ima 240 testova (116 / 119 / 5). 27 testova sa 6 ili više sabotaža nosi 223 sabotaže (17,6%). Max 14 (T_FakturaGP_WriterKapijeIStorno), u BFP 12 (Test_OTK_CitaociCitajuStavke). Prosek 2,5 sabotaže po testu u RunAllTests, 2,0 u BFP.
    DOKAZ: S/anat.py -> 'RunAllTests testova 318 sa 1 sab: 116 ... >=10: 6 max: 14'; 'RunBusinessFlowProSuite testova 227 sa 1 sab: 119 ... max: 12'
- [M] Po ključu (test, tvrdnja): 1264 različita ključa na 1265 sabotaža. Jedini par koji deli ključ je u T_Oporavak_CiljneListe i već je priznat nalaz. Katalog dakle nije redundantan po tvrdnji — svaka sabotaža je svedok druge tvrdnje.
    DOKAZ: S/anat.py -> 'razlicitih (test,tvrdnja): 1264 ; kljuceva sa >1 sabotazom: 1'; A/tools/sabotaza.py:12628-12629
- [M] Sidra: 1224 različita para (fajl, sidro); 40 identičnih sidara deli 81 sabotaža; 45 istih početnih redova (95 sabotaža); 117 sabotaža ima opseg redova koji se preklapa sa nekom drugom. Sidro je u medijani 1 red (913 jednorednih), najviše 17.
    DOKAZ: S/anat.py -> 'sidara sa >1 sabotazom: 40 ; sabotaza u njima: 81', 'sabotaza cije se sidro preklapa: 117', 'sidro redova: median 1, max 17; 1-redna: 913'
- [M] POZNATI_NALAZI ima 9 upisa (7x 'tvrdnja nije CEO STATICAN literal' u BFP, 1x deljena tvrdnja, 1x mrtva ljuska-rez-bez-potvrde); POZNATI_NALAZI_DOKAZ ima 2 (ljuska-rez-bez-potvrde: NE OBARA NISTA; f8-identitet-po-broju: PALA DRUGA TVRDNJA). Ukupno 10 različitih sabotaža (0,8%). dokaz.py čita samo drugi spisak.
    DOKAZ: A/tools/sabotaza.py:12586-12667; A/tools/dokaz.py:453
- [M] Pokrivenost tvrdnji sabotažom (grubo: tekst iz kataloga upoređen sa literalom poruke tvrdnje u telu testa): modTest 789 od 3376 (23,4%), BFP 449 od 3188 (14,1%), banka 8 od 239 (3,3%); ukupno 1246 od 6803 = 18,3%, to jest oko 5557 tvrdnji nema nijednu sabotažu. Testova sa tvrdnjama a bez ijedne sabotaže: modTest 29 (166 tvrdnji), BFP 166 (956), banka 19 (173).
    DOKAZ: python3 S/pokr.py A (koristi sab._poruke_tvrdnji i _poruka_delovi) -> 'razlicitih (test,poruka): 3376 ; pokrivenih sabotazom: 789 (23.4%)' itd. Dinamičke poruke (123 u modTest, 334 u BFP, sve u banci) su procenjene grubo.
- [M] Plan grupa nad celim katalogom: cap 1 -> 1265 prolaza, 2 -> 634, 3 -> 425, 4 -> 320, 6 -> 215 (134 + 78 + 3), 8 -> 163, 12 -> 112, 24 -> 63, bez cap-a -> 29 (donja granica je najveći broj sabotaža po testu: 14 + 12 + 3).
    DOKAZ: S/anat.py (dok.naprav_grupe) -> 'cap 6: prolaza 215 {RunAllTests: 134, RunBusinessFlowProSuite: 78, RunBankaImportTestSuite: 3}', 'cap 1000: prolaza 29'
- [M] Četiri pravila grupisanja nad celim katalogom ne koštaju gotovo ništa: cap 6 sa svim pravilima 215 prolaza, samo suite+cap 212. Strože pravilo 'različit FAJL' umesto 'različita procedura' daje 247 (+15%); 192 od 215 današnjih grupa ima bar dva člana u istom fajlu. Komentar uz GRUPA_PODRAZUMEVANO ('pravila vezu grupu na ~5 clanova, pa veci broj skoro nista ne doda') i brojevi 118/119 u docstring-u su zastareli.
    DOKAZ: inline skripta nad S/stavke_main-wt.json -> 'cap 6 sva pravila: 215 | bez razlicit test: 213 | bez razlicita procedura: 215 | samo suite+cap: 212'; 'razlicit FAJL: 247'; A/tools/dokaz.py:67-71 i :230-232
- [M] Pojedinačan prolaz pušta JEDNU suite (onu u kojoj živi test sabotaže), ne obe; obe se puštaju samo kao baza (suite stavki unija RunAllTests, i kad nijedna izabrana sabotaža nije u RunAllTests). CLAUDE.md tvrdi da dokaz 'po svakoj sabotaži vrti RunAllTests i BFP — oko 3,5 minuta po stavci', što ne odgovara kodu.
    DOKAZ: A/tools/dokaz.py:344-349 (suite = primenjeni[0]['suite']; jedan run_vba poziv), :466 (potrebne = ... | {SUITE_ALL}); A/CLAUDE.md:113
- [M] Tvrdnje koje dokaz daje: T1 baza zelena pre prve mutacije (jednom, na početku celog prolaza); T2 posle svakog prolaza --vrati i otisak celog src-vba jednak početnom, inače STOP (otisak ne pokriva svesku ni golden fajlove); T3 pao je baš njen test i u poruci tog testa stoji tekst tvrdnje (podniz za RunAllTests i banku, ceo literal kao ključ za BFP); T4 u grupi se priznaje samo član koji je oborio svoju tvrdnju, ostali se ponavljaju sami, verdikt nosi '(grupno)' jer se ne razlučuje čija je mutacija oborila tvrdnju; imenilac: plan pokriva svaku sabotažu tačno jednom i crvenih mora biti >= broj stavki minus priznati nalazi. T5 (samo grana B): baza se tvrdi nad ISTIM spiskom testova.
    DOKAZ: A/tools/dokaz.py:16-41 (docstring), :82-89 (_otisak samo src-vba), :211-273 (naprav_grupe), :276-305 (_oceni), :355-365, :413-417, :574-575; B/tools/dokaz.py docstring tačka 5
- [M] Svaki run_vba poziv podiže nov Excel, uvozi svih 194 fajla iz src-vba, radi compile probu i EnsureRuntimeSchema. Fiksni sleep-ovi po pozivu: 2,5 s po pokušaju compile-a (do dva pokušaja, drugi još 0,9 s) plus 1,0 s i 1,0 s pri gašenju = 4,5 do 7,9 s od dokumentovanih ~18 s režije.
    DOKAZ: A/tools/run_vba.py:353-407 (import_src_vba), :617, :632, :634, :1012, :1021, :875 (DispatchEx)
- [M] Python strana jednog prolaza (Linux, kopija alata i src-vba u S/kopija): sabotaza.py --lista 0,08 s; sabotaza.py --vrati 1,00 s jer prolazi kroz svih 1265 unosa i za svaki ponovo čita fajl; _otisak 0,013 s.
    DOKAZ: komande u S/kopija: 'python3 tools/sabotaza.py --vrati' -> 1.03 / 1.00 / 1.00 s; A/tools/sabotaza.py:12535-12547
- [M] dokaz.py nad celim katalogom troši 27–29 s pre prvog Excela (isto i za --plan): _suita_testa se zove 1265 puta i svaki put ponovo čita i regex-uje test module.
    DOKAZ: 'python3 tools/dokaz.py --grupe 6 --plan' -> 28.88 s; cProfile: _suita_testa 1265 poziva, 26.8 od 27.4 s; A/tools/sabotaza.py:12681-12696, A/tools/dokaz.py:97-108
- [M] sabotaza.py --proveri-sidra (ide u hook posle svakog Edit-a) traje 4,0 s; 3,37 s od toga je _dodela_tudjoj_proceduri, koja za svaki unos ponovo parsira ceo mutirani fajl. Sa pred-proverom (zamena uopšte dodeljuje imenu neke procedure ili uvodi proceduru) pun račun treba samo 178 od 1265 unosa i traje 0,45 s, uz isti skup nalaza (0) nad zatečenim katalogom.
    DOKAZ: 'python3 tools/sabotaza.py --proveri-sidra' -> 4.01 s; cProfile: _dodela_tudjoj_proceduri 8.8 od 9.9 s (sa profilerom), 8,7M regex poklapanja; inline merenje: 'sa predproverom: pun racun za 178 unosa, nalaza 0, 0.45 s' / 'bez predprovere: nalaza 0, 3.37 s'; A/tools/sabotaza.py:13150-13209
- [N] Dokumentovana trajanja se ne slažu: režija ~18 s i BFP ~150 s (harness dokument); BFP 2417/0 za 210 s (07.10.); 'crveni RunAllTests 210 s + BFP' je probio limit od 600 s (početak novembra), dakle BFP > oko 370 s; dokaz.py --grupe 6 nad 4 BFP sabotaže (4 grupe + 2 baze = 6 poziva) trajao je 34 min, to jest oko 340 s po pozivu. BFP je dakle početkom novembra oko 350–380 s, 2,5 puta više od broja u dokumentaciji i u docstring-u dokaz.py.
    DOKAZ: A/docs/EXCEL_TEST_HARNESS.md:731-732; A/docs/STANJE_REFAKTORA.md:995, :10525-10526, :10258 ('dokaz.py --plan s8r1c (četiri grupe, BFP)') i :10280 ('dokaz 11:00–11:34 (oko 34 min; procena pre puštanja je bila 20–25)'); A/tools/dokaz.py:214
- [M] dokaz.py ne prosleđuje --timeout, pa za CEO run_vba poziv važi podrazumevanih 600 s. Kad run_vba ubije Excel, u izlazu nema FAIL redova i ocena glasi 'NE OBARA NISTA' (u bazi: STOP), a ne TIMEOUT. Uz BFP od ~370 s + 18 s režije prag se dostiže na oko 1,5x rasta.
    DOKAZ: A/tools/dokaz.py:92 (subprocess timeout 1200), :155 i :349 (bez --timeout), :287-288; A/tools/run_vba.py:705-706, :1006-1010
- [M] Rezultat dokaza se nigde mašinski ne beleži: dokaz.py samo ispisuje, trajanje suite koje run_vba štampa ('SUITE ... (Ns)') se odbacuje, tests/last_run.* je gitignored i prepisuje se, a vba_gate marker ne zna za dokaz. U STANJE_REFAKTORA nema zapisa da je pun katalog ikad pušten nad sadašnjom veličinom; 'pun katalog pred release' se više puta odlaže.
    DOKAZ: A/tools/run_vba.py:1091; A/.gitignore:15-18; grep 'dokaz' A/tools/vba_gate.py -> samo :101; A/docs/STANJE_REFAKTORA.md:6973, :10115, :10336, :10951, :10981
- [M] Rast i održavanje kataloga (git, first-parent, 50 koraka nad tools/sabotaza.py, 10.10.–08.11.2026): 724 -> 1265, to jest +18,7 dnevno; dodato 572, obrisano 25, izmenjeno 79 postojećih (sidro ili fajl 64, zamena 53, tvrdnja 13, test 2); 21 od 50 koraka (42%) nosi presidrenje. To je donja granica, jer first-parent krije popravke unutar PR-a.
    DOKAZ: python3 S/churn.py A S -> 'ukupno: dodato 572, obrisano 25, izmenjenih postojecih 79 (sidro/fajl 64, zamena 53, test 2, tvrdnja 13)', 'koraka ... sa presidrenjem: 21'
- [M] Truljenje u dokumentima: kad je dokaz prvi put pušten šire, 10 od 222 sidara je bilo mrtvo, 1 sabotaža nije obarala ništa, 1 test preimenovan; kasnije je 119 od 251 unosa nosilo zastareo tekst tvrdnje. STANJE_REFAKTORA ima 43 reda sa 'presidren*' (eksplicitni brojevi uz njih u zbiru 54) i 8 puta 'NIJE DOKAZANO' kao verdikt; truljenje se danas hvata odmah (hook), ali popravka i ponovni dokaz ostaju ručni.
    DOKAZ: A/docs/engineering/postmortems/2026-08-verifikacija.md:138-158; A/tools/sabotaza.py:12707-12709; grep -c nad A/docs/STANJE_REFAKTORA.md: 'presidren' 43, 'NIJE DOKAZANO' 8
- [M] Format: katalog je 12.368 redova (9,8 po unosu, od toga 2.673 reda komentara), fajl ima 735 KB; 1.474 escape-ovana navodnika u 532 reda, a 371 red koristi jednostruke navodnike da izbegne escape (mešano pisanje). Nova dodavanja padaju i u sredinu fajla, a uzastopni PR-ovi u isti region (redovi 9112–9283), pa je fajl magnet za konflikte — grana B se sa main-om sudara upravo tu.
    DOKAZ: inline merenje nad A/tools/sabotaza.py -> 'katalog redova: 12368 komentara: 2673 ... po unosu 9.8'; 'escape-ovanih navodnika: 1474 ; redova sa njima: 532'; git diff --unified=0 za poslednjih 12 first-parent commit-a (hunk pozicije 7626, 12425, 7447, 11500, 9232, 9283, 9249, 9207, 9149, 9112)
- [M] Grana B, --izbor: baza i mutacija su DVA run_vba poziva nad istim spiskom testova, a baza se pamti po (suite, spisak). Nad katalogom sa main-a: bez grupa treba 545 baza (318 + 227) + 1257 prolaza + banka (nema izbor: 8 prolaza cele suite + 1 baza) = 1811 poziva. Sa grupama 6: 208 baza (132 + 76) na 212 prolaza — keš pogodi samo 4 puta — ukupno 424 poziva plus ponavljanja (svako traži i svoju novu bazu).
    DOKAZ: python3 S/baze.py A -> 'cap 1: ... baza izbora A=318 B=227 (ukupno 545) | run_vba poziva = 1811'; 'cap 6: ... baza izbora A=132 B=76 (ukupno 208) | run_vba poziva = 424'; B/tools/dokaz.py (_prolaz: kljuc = (suite, tuple(testovi)))
- [M] Poravnanje grupa po bloku testova (blok od 6 testova, sloj = po jedna sabotaža svakog testa, baza po bloku) dalo bi 91 bazu + 250 prolaza = 341 poziv umesto 420 (−19%) uz ista pravila o ključu i proceduri.
    DOKAZ: python3 S/blok.py A -> 'cap 6 blok-plan: baza 91 + prolaza 250 = 341 poziva'
- [M] Spajanje grane B je mehanički malo: probni merge u scratch klonu daje 6 konfliktnih blokova (modTest 3, modBusinessFlowProTests 2, sabotaza.py 1); dokaz.py i run_vba.py se spajaju bez konflikta. Na main-u je 391 goli poziv Test_* u RunBusinessFlowProSuite (grana obmotava 361), pa 31 nov poziv traži BfpIzabran; test broj 311 sa grane se sudara sa numeracijom main-a (350 RunOne).
    DOKAZ: git merge-tree --write-tree glavna HEAD u S/klon -> CONFLICT u tri fajla; brojanje '^<<<<<<<' po fajlu: 3 / 2 / 1 / 0 / 0; inline brojanje golih poziva: d19c2da 360, main 391, grana 0 golih i 361 kroz BfpIzabran
- [M] Spisak testova koji --izbor šalje kroz Application.Run: uz cap 6 najduži je 239 znakova, uz cap 12 je 455 (103 od 109 grupa prelazi 255).
    DOKAZ: inline skripta nad dok.naprav_grupe -> 'cap 6 max duzina spiska testova: 239'; 'cap 12 max: 455 ; grupa > 255 znakova: 103 od 109'
- [M] BFP testovi statički deluju nezavisno jedan od drugog: 354 od 398 Test_* koristi NewScenarioCode, samo 2 diraju modul-stanje (m_RunID), 96 koristi transakciju; seed (SeedBusinessFlowProMasterData) se u izboru pušta uvek. Ponašanje u izolaciji nije izvršeno.
    DOKAZ: inline regex analiza A/src-vba/modBusinessFlowProTests.bas -> 'NewScenarioCode 354', 'm_ modul-stanje 2 / 398', 'TxBegin 96 / 398'; B/src-vba/modBusinessFlowProTests.bas telo RunBusinessFlowProSuite
- [M] Mutant schemata: samo 682 od 1265 unosa (54%) je mehanički obmotljivo u 'If prekidac Then zamena Else sidro End If'. Kod ostalih sidro otvara blok bez zatvaranja (305), jeste Else/Case (100), seče nastavak reda (139), ili je deklaracija/potpis (30).
    DOKAZ: python3 S/schemata.py A -> 'mehanicki obmotljivo u If-prekidac (grubo): 682 / 1265 (54%)'
- [M] Po klasi modula: pisci domen-tabela (row_owner iz WRITE_OWNERSHIP.json) 515 sabotaža (41%), UI/ekrani 371 (29%), ostala logika i čitaoci 294 (23%), štampa/izveštaji 85 (7%).
    DOKAZ: inline skripta: A/docs/DOMEN/WRITE_OWNERSHIP.json (26 row_owner modula) x S/stavke_main-wt.json
- [M] vba_gate već ume da odgovori da li je BAŠ OVAJ izvor zelen za datu suite (otisak izvora, sveske i ugovora), ali dokaz.py to ne koristi nego bazu uvek pušta iznova, iako radni tok reza pušta ciljanu suite neposredno pre dokaza.
    DOKAZ: A/tools/vba_gate.py:13-16 (--require-green --suite), :735 (zahtevaj_zeleno); A/tools/dokaz.py:465-473; A/docs/STANJE_REFAKTORA.md:992 ('ciljana BFP suite PRVA, pa dokaz.py')
- [M] Dokaz mutira RADNO stablo (sabotaza.py piše u src-vba, run_vba uvozi iz istog ROOT-a), pa je src-vba zaključan dok dokaz traje: svaka paralelna izmena daje REVERT-FAIL i STOP. Pravila kataloga o praznoj zameni, zameni kao podnizu sidra i deljenoj zameni (zamke 3, 7, 8, 11) postoje samo zato što se vraćanje radi obrnutom zamenom teksta u radnom stablu.
    DOKAZ: A/tools/dokaz.py:325, :358-365; A/tools/sabotaza.py:24-61 (zamke), :13271-13334

## TROSKOVI
[
 {
  "sta": "Pun pojedinačni dokaz, cela suite po sabotaži. Formula: C1 = n_A(r + T_A + p) + n_B(r + T_B + p) + n_K(r + T_K + p) + baze; n = 801 / 456 / 8, r = 18 s (dokumentovano), p = 1,1 s (izmereno, Linux).",
  "koliko": "43,6 h uz dokumentovana trajanja (T_A = 80 s, T_B = 150 s); 81,7 h uz trajanja izvedena iz dnevnika za početak novembra (T_A = 120 s, T_B = 380 s). 1268 run_vba poziva.",
  "izmereno": false,
  "koliko_cesto": "po pravilu pred svaki release; nema traga da je ikad pušten nad sadašnjim katalogom"
 },
 {
  "sta": "Dokaz --grupe 6, cela suite. Formula: C2 = G_A(r + T_A) + G_B(r + T_B) + G_K(r + T_K) + rho x C1; G = 134 / 78 / 3 (izmereno planom), rho = 8% ponavljanja (procena iz pet zapisa u STANJE_REFAKTORA, npr. 3 od 36).",
  "koliko": "10,8 h (dokumentovana trajanja) do 20,4 h (novembarska); od toga ponavljanja 3,4 do 6,5 h.",
  "izmereno": false,
  "koliko_cesto": "nije pušten nad celim katalogom; isti režim se koristi po rezu"
 },
 {
  "sta": "Dokaz --izbor (grana B), pojedinačno. Formula: C3 = (n_A + n_B)(r + tau + p) + (D_A + D_B)(r + tau) + banka; D = 318 / 227 baza, tau = trajanje izbora jednog testa (nepoznato; BFP uključuje seed).",
  "koliko": "10,5 h (tau = 2 s), 11,5 h (4 s), 13,5 h (8 s); 1811 run_vba poziva; režija je oko 79% vremena. Testovi koji sami ne prolaze dodaju po jedan prolaz cele suite (udeo nepoznat).",
  "izmereno": false,
  "koliko_cesto": "VBA strana nikad izvršena"
 },
 {
  "sta": "Dokaz --izbor --grupe 6 (grana B). Formula: C4 = (G + B6)(r + tau6) + 2 x rho x n x (r + tau); G = 212, B6 = 208 baza.",
  "koliko": "3,9 do 6,4 h; 424 poziva + oko 202 za ponavljanja.",
  "izmereno": false
 },
 {
  "sta": "Projekcija uz +18,7 sabotaža dnevno (1825 za 1 mesec, 2944 za 3, 4624 za 6) i pretpostavku da trajanje suite raste srazmerno katalogu.",
  "koliko": "Pojedinačno cela suite: 86–166 h / 215–421 h / 517–1025 h (kvadratno). Grupe 6: 22–41 h / 54–105 h / 129–256 h (kvadratno). Izbor pojedinačno (tau = 4 s): 16,5 / 26,7 / 41,8 h (linearno). Izbor + grupe 6: 6,9 / 11,1 / 17,4 h. Izbor pojedinačno uz režiju 5 s: 7,1 / 11,4 / 18,0 h.",
  "izmereno": false,
  "koliko_cesto": "1 / 3 / 6 meseci; ako trajanje suite ne raste, prva dva reda rastu linearno (x1,44 / x2,33 / x3,66)"
 },
 {
  "sta": "Tipičan dokaz reza danas (4 BFP sabotaže, --grupe 6: 2 baze + 4 prolaza)",
  "koliko": "34 min (zapis u dnevniku); baze same oko 9–10 min od toga. U izboru bi isti rez bio 8 poziva po oko 22–26 s, to jest 3–4 min (model).",
  "izmereno": false,
  "koliko_cesto": "po rezu; u STANJE_REFAKTORA najmanje 23 DOKAZANO i 8 NIJE DOKAZANO verdikta u oko mesec dana"
 },
 {
  "sta": "Bazni prolaz pri svakom pozivu dokaz.py (RunAllTests uvek + suite stavki), iako je ista suite često puštena neposredno pre",
  "koliko": "1–3 run_vba poziva; uz novembarska trajanja oko 2–9 min po pozivu",
  "izmereno": false,
  "koliko_cesto": "svaki poziv dokaz.py"
 },
 {
  "sta": "Fiksni sleep-ovi u run_vba (compile proba 2,5 s po pokušaju + gašenje 2 x 1,0 s)",
  "koliko": "4,5–7,9 s po pozivu: 1,6–2,8 h u punom pojedinačnom dokazu, 2,3–4,0 h u izboru pojedinačno (1811 poziva)",
  "izmereno": false,
  "koliko_cesto": "svaki run_vba poziv (redovi su iz koda; udeo u stvarnih 18 s nije meren)"
 },
 {
  "sta": "_suita_testa bez pamćenja u dokaz.py",
  "koliko": "27–29 s pre prvog Excela, i za --plan",
  "izmereno": true,
  "koliko_cesto": "svaki poziv dokaz.py nad celim katalogom (uz filter srazmerno manje)"
 },
 {
  "sta": "sabotaza.py --vrati prolazi ceo katalog i čita fajl za svaki unos",
  "koliko": "1,0 s po prolazu; oko 21 min u punom pojedinačnom dokazu (Linux; Windows verovatno sporije)",
  "izmereno": true,
  "koliko_cesto": "svaki prolaz"
 },
 {
  "sta": "--proveri-sidra u PostToolUse hook-u",
  "koliko": "4,0 s, od toga 3,37 s u _dodela_tudjoj_proceduri; raste sa katalogom i veličinom fajlova",
  "izmereno": true,
  "koliko_cesto": "posle svakog Edit/Write nad VBA fajlom i nad tools/sabotaza.py"
 },
 {
  "sta": "Održavanje kataloga: nove stavke i presidrenje",
  "koliko": "572 nove, 25 obrisanih, 79 izmenjenih (64 presidrenja) u 29 dana; oko 2,2 presidrenja dnevno, jedno na svakih 9 novih; svako presidrenje traži i ponovni dokaz u Excelu",
  "izmereno": true,
  "koliko_cesto": "21 od 50 merge koraka nad katalogom"
 },
 {
  "sta": "Zaključano radno stablo tokom dokaza (src-vba se mutira na mestu)",
  "koliko": "koliko traje dokaz: 34 min za mali rez, 11–82 h za pun katalog",
  "izmereno": false,
  "koliko_cesto": "svaki poziv dokaz.py"
 },
 {
  "sta": "Litica: run_vba limit 600 s za ceo poziv, dokaz.py ga ne podešava",
  "koliko": "BFP oko 370 s + 18 s danas; prag na oko 1,5x rasta, to jest 0,7–1,9 meseci uz rast trajanja 25–80% mesečno",
  "izmereno": false,
  "koliko_cesto": "jednom, ali tada pada svaki BFP prolaz i baza"
 }
]

## PREDLOZI
### D1 Knjiga dokaza i telemetrija (JSONL) [trud S] [slabi: NE]
dokaz.py za svaki prolaz dopiše red: imena članova, suite, režim (cela suite / izbor, pojedinačno / grupno), verdikt po članu, SVA pala imena i tvrdnje (kill-skup, danas se od njega čuva samo broj 'uz jos N testa'), sekunde prolaza i sekunde suite iz reda 'SUITE ... (Ns)', otisak izvora, heš unosa kataloga, heš sveske, datum. run_vba uz to ispiše trajanje faza (start Excela, uvoz, compile proba, schema, suite, gašenje).
  DOBITAK: Model troška postaje meren umesto izveden iz dnevnika (T_A, T_B, r, rho, tau su danas nepoznati). Osnova za D7, D8, D9, D10. Ukida ručno prepisivanje verdikta u STANJE_REFAKTORA.
  RIZIK: Nizak. Knjiga ne sme postati izvor verdikta bez provere otiska (v. D7).
  PREDUSLOVI: nema
  DOKAZATI: dokaz.py --self-test dobija slučajeve sa lažnim runner-om: red se upiše za svaki ishod (priznat, ponavlja se sam, APPLY-FAIL, TIMEOUT, REVERT-FAIL) i ne upiše se dvaput; pokazati crveno izbacivanjem upisa.

### D2 Jeftine režijske popravke u dokaz.py i sabotaza.py [trud S] [slabi: NE]
(a) pamćenje rezultata _suita_testa; (b) --vrati samo nad primenjenim članovima ili sa čitanjem fajla jednom po fajlu; (c) sve bazne suite u jednom run_vba pozivu (run_vba već prima više --suite); (d) dokaz.py prosleđuje izričit --timeout, a FATAL iz run_vba imenuje kao prekid, ne kao 'NE OBARA NISTA'; (e) pred-provera u _dodela_tudjoj_proceduri pre punog parsiranja; (f) ispraviti zastarele brojeve u dokaz.py:67-71 i :230-232 i u harness dokumentu, i rečenicu u CLAUDE.md:113.
  DOBITAK: (a) 27 s -> ispod 1 s po pozivu; (b) 1,0 s -> oko 0,05 s po prolazu (21 min na punom dokazu); (c) 18–36 s po pozivu; (d) uklanja liticu od 600 s i lažnu etiketu; (e) 3,37 s -> 0,45 s po hook-u, to jest oko 3 s manje posle svakog Edit-a.
  RIZIK: Nizak. (e) mora dati isti skup nalaza kao danas; (d) veći timeout produžava čekanje na sabotažu koja obori compile.
  PREDUSLOVI: (f) izmena CLAUDE.md i .claude/rules ide zasebnim process PR-om
  DOKAZATI: Postojeći --self-test oba alata ostaje zelen; za (e) uporediti spisak nalaza _nalazi pre i posle nad celim katalogom i nad _DODELA_CASES, pa sabotirati pred-proveru i pokazati da slučaj pada po imenu; za (d) nov slučaj u self-testu sa izlazom koji nosi FATAL.

### D3 Baza iz markera zelenog i samo za suite koje se mere [trud S] [slabi: NE ako je marker vezan za isti otisak (nema markera -> baza se pušta, fail-closed). Gubi se jedino usputna provera RunAllTests kad se meri samo BFP — to nije deo tvrdnje T1.]
Pre bazne suite dokaz.py pita vba_gate (zahtevaj_zeleno za tu suite); ako marker važi za isti otisak izvora, sveske i ugovora, baza se ne pušta ponovo. RunAllTests se kao baza pušta samo kad neka izabrana sabotaža živi u njoj.
  DOBITAK: 1–2 prolaza cele suite manje po pozivu dokaza: oko 2–9 min od tipičnih ~34 min; ukupno nekoliko sati mesečno uz tridesetak poziva.
  RIZIK: Marker i dokaz moraju računati isti otisak (danas su dva različita: dokaz._otisak i vba_gate.otisak_izvora). Marker se ne upisuje za izbor niti za --no-import, pa važi samo za celu suite.
  PREDUSLOVI: nema
  DOKAZATI: Self-test: važeći marker -> nula baznih poziva; marker nad drugim izvorom ili drugom sveskom -> baza se pušta; oba smera po imenu.

### D4 Spojiti izbor testa (grana B) na main i izvršiti ga prvi put u Excelu [trud M] [slabi: DELIMIČNO, T3: tvrdnja se meri kad se test pusti sam, a ne u punom redosledu suite — mutacija koju bi stanje prethodnih testova zaklonilo ovde izgleda dokazano; gubi se i podatak o kolateralnim padovima. Istovremeno jača atribuciju: susedni testovi više ne mogu da ostave stanje koje obori ciljanu tvrdnju. T1 ostaje (baza nad istim spiskom). Razlika se pokriva u D9 jednim grupnim prolazom cele suite pred release.]
Rebase commit-a 9d6e780: 6 konfliktnih blokova, obmotati 31 nov BFP poziv u BfpIzabran, prenumerisati test 311. Prvi Windows prolaz nad jednim rezom u oba režima (cela suite i izbor) i uporediti verdikte stavku po stavku. Posle toga je za rez podrazumevan POJEDINAČAN dokaz u izboru (bez oznake grupno). Opciono: grupe poravnati po bloku testova da se baza deli (341 umesto 420 poziva).
  DOBITAK: Prolaz sa r + T (oko 100–400 s) pada na r + tau (oko 20–26 s). Tipičan rez 34 min -> 3–5 min. Pun pojedinačni dokaz 44–82 h -> 10,5–13,5 h, a rast postaje linearan (za 6 meseci 42 h umesto 517–1025 h).
  RIZIK: VBA strana nikad nije izvršena. Testovi koji sami ne prolaze padaju na celu suite i plaćaju i promašenu bazu (udeo nepoznat; BFP statički deluje nezavisno). Spisak od 6 testova ide do 239 znakova, a moguć limit Application.Run nije proveren. Seed u BFP izboru može da dominira.
  PREDUSLOVI: D2(d); Windows + Excel
  DOKAZATI: Diferencijalno: isti rez (npr. 20 amb-pisac sabotaža) u oba režima, verdikti moraju biti identični po stavci; negativne kontrole iz D12 moraju ostati crvene; self-test grane (23 provere) već postoji i mora ostati zelen posle rebase-a.

### D5 Mutacija van radnog stabla [trud M] [slabi: NE. T2 postaje jača: izvor nije vraćen nego nikad nije diran, a otisak to i dalje tvrdi.]
run_vba dobija opciju da jedan modul uveze sa druge putanje (npr. --zameni modX.bas=putanja). dokaz.py mutirani tekst pravi u temp folderu iz čistog fajla (jedna zamena, tvrdi se tačno jedan pogodak) i nikad ne piše u src-vba. Otisak pre/posle ostaje kao provera da niko nije dirao izvor. Ručni sabotaza.py <ime> / --vrati ostaje kakav jeste.
  DOBITAK: Radno stablo više nije zaključano tokom dokaza (danas 34 min do desetina sati). Nestaju klase REVERT-FAIL, zatečen sabotiran izvor i pokvaren izvor posle ubijenog procesa. Otvara više paralelnih Excel instanci nad istim checkout-om (zidno vreme deljeno brojem instanci). Za dokaz prestaju da budu potrebna pravila o praznoj zameni, podnizu i deljenoj zameni (zamke 7, 8, 11), pa sabotaža koja briše red više ne traži veštačku oznaku.
  RIZIK: Uvoz sa temp putanje za .cls (6 sabotaža) i .doccls (1) ide drugim putem nego .bas. Paralelne instance: compile proba koristi SetFocus / AppActivate / SendKeys, a tests/last_run.* i marker su deljeni — traži merenje na Windows-u.
  PREDUSLOVI: nema (nezavisno od D4)
  DOKAZATI: run_vba --self-test: zamenjeni modul se uvozi sa date putanje, ostali iz src-vba (oba smera). U Excelu: ista sabotaža starim i novim putem daje isti verdikt; otisak src-vba identičan i posle nasilnog prekida prolaza.

### D6 Smanjiti režiju poziva (posle merenja faza) [trud L] [slabi: NE za (a) i (b). Za (c) NE samo ako baza posle vraćanja modula mora biti zelena (sendvič) — tada je T1 čak jača; bez toga DA, jer stanje jedne mutacije može da procuri u sledeću.]
Redom, tek kad D1 pokaže gde ide 18 s: (a) unapred sagrađena sveska po otisku izvora — uvoz 194 modula jednom, zatim --no-import plus zamena jednog modula po prolazu, sveža kopija sveske za svaki prolaz; (b) skratiti fiksne sleep-ove gde merenje pokaže da ne nose ništa; (c) sesija-sendvič: baza -> zamena modula -> isti izbor -> vraćanje modula -> baza, više mutacija po Excel sesiji uz recikliranje na N mutacija.
  DOBITAK: U izboru je režija oko 79% vremena. Uz režiju 5 s pun pojedinačni dokaz je oko 4,9 h (umesto 11,5 h); sesije po 20 mutacija oko 3,4 h (model). Ubrzava i svaki TARGETED poziv run_vba.
  RIZIK: (c) nosi stanje sveske između mutacija (oko 300 od 398 BFP testova ne koristi transakciju) i poznat neistražen pad Excela posle oko 10 min rada (STANJE_REFAKTORA:8347-8354); zamena modula resetuje stanje VBA projekta. (a) je znatno niži rizik od (c).
  PREDUSLOVI: D1 (faze), D4, D5
  DOKAZATI: Diferencijalno nad uzorkom od stotinak sabotaža: verdikt po stavci jednak onom iz zasebnih poziva; kanarinac iz D12 u svakoj sesiji; self-test orkestracije sa lažnim runner-om (redosled baza-mutacija-baza, recikliranje).

### D7 Nastavak prekinutog dokaza po potpuno konzervativnom ključu K0 [trud S] [slabi: NE (isti ulazi, isti režim; pretpostavlja determinizam koji se i danas pretpostavlja)]
Sabotaža se preskače samo ako u knjizi (D1) postoji POJEDINAČNO DOKAZANO za isti ključ K0 = heš( otisak celog src-vba, heš unosa (fajl, sidro, zamena, test, tvrdnja), heš sveske + .sig + tests/golden, heš alata dokaz.py / run_vba.py / vba_gate.py i ne-kataloškog dela sabotaza.py, režim i spisak testova ). Negativni i grupni ishodi se ne keširaju.
  DOBITAK: Prekid posle 20 od 40 sati ne košta 20 sati. Release dokaz može da ide u više noći nad zamrznutim commit-om i da se nastavi posle nalaza koji dira samo katalog (menja se heš unosa samo te stavke). Između commit-a koji menjaju src-vba pogodaka nema — to je cena konzervativnosti.
  RIZIK: Nizak. Ključ ne pokriva okruženje: verziju Excela, locale (DMY/MDY), datum i vreme, nedeterminizam testa — to ne pokriva ni današnji dokaz.
  PREDUSLOVI: D1
  DOKAZATI: Self-test: promena bilo kog sastojka ključa (jedan bajt u src-vba, unos, sveska, alat, režim) -> stavka se ponovo meri; nepromenjeno -> preskače se; svaka osa posebno i u oba smera, po uzoru na test osa u dokaz.py --self-test.

### D8 Automatski obim reza: dokaz.py --od <git-ref> [trud M] [slabi: NE u odnosu na današnju praksu reza (ista definicija, automatizovana). Ne zamenjuje pun katalog.]
Alat sam bira stavke: nove i izmenjene u katalogu, sabotaže čiji je FAJL diran, sabotaže čije je telo TESTA dirano, i sve nad test modulom čiji je pomoćni kod diran. Danas se taj spisak sastavlja ručno (skripta u scratchpad-u, nabrajanje prefiksa u komandi).
  DOBITAK: Manje ručnog rada po rezu i nema propuštenih presidrenih ili posredno pogođenih stavki (dnevnik beleži slučajeve gde su dodavane naknadno).
  RIZIK: Pravilo o pomoćnom kodu test modula može da izabere stotine stavki; treba ga ograničiti na procedure koje test zaista poziva. Preimenovanje fajla.
  PREDUSLOVI: nema; bolje uz D1
  DOKAZATI: Self-test nad izmišljenim katalogom i lažnim diff-om: svaka od četiri vrste pogotka bira svoju stavku, a stavka bez pogotka ne ulazi; oba smera.

### D9 Slojevit raspored puštanja [trud M] [slabi: NE za release tvrdnju — release ostaje pun pojedinačni dokaz nad release izvorom. Sama rotacija tvrdi manje (dokazano u poslednjih K dana, nad tadašnjim izvorom), pa sme da bude samo dodatak, nikad zamena. Nedeljni prolaz je grupni i tako se i označava.]
Rez: obim iz D8, pojedinačno u izboru. Noćno: sve novo ili dirano od poslednjeg noćnog prolaza + sloj kritičnih (pisci domen-tabela po WRITE_OWNERSHIP.json, danas 515 stavki = 41%) + rotacioni isečak ostatka (750 / K, stabilno po hešu imena). Nedeljno: jedan grupni prolaz CELE suite nad celim katalogom (kontekst redosleda i kill-matrica). Release: pun pojedinačni dokaz nad zamrznutim commit-om, uz nastavak po K0.
  DOBITAK: Release dokaz prestaje da bude jednokratan udar od više dana: trulež i mrtve sabotaže izlaze u roku od dana, a release prolaz je u jeftinom režimu (oko 5–12 h na jednoj mašini, manje uz D5). Kritične invarijante se mere svake noći.
  RIZIK: Traži mašinu sa Excelom koja radi noću i stabilan runner. Klasifikacija po modulu je gruba (kritična tvrdnja može da živi u UI modulu).
  PREDUSLOVI: D1, D4, D7, D8
  DOKAZATI: Plan rotacije je čista funkcija: self-test tvrdi da svaka stavka uđe tačno jednom u K noći i da kritični sloj uđe svake noći. Vrednost rasporeda se meri iz knjige: koliko dana prođe od nastanka nalaza do njegovog otkrivanja.

### D10 Keš preko commit-a po zatvaranju poziva (K1) — istraživanje, ne za release [trud L] [slabi: DA dok nije validiran: tvrdnja postaje 'dokazano nad izvorom koji je po K1 ekvivalentan', ne 'nad ovim izvorom'. Zato ne ulazi u release tvrdnju.]
K1 = heš unosa + tekst mutirane procedure posle zamene + tela svih procedura u zatvaranju po imenu od testa (identifikatori, i oni u string literalima zbog Application.Run / CallByName / OnTime) i harness-a + deklaracione sekcije svih modula iz zatvaranja + ceo tekst svake klase ili forme koja se pominje (događaji) + sveska, golden, alati. Važi samo za izolovan test (izbor). Ne pokriva: ime procedure sastavljeno iz podataka, redosled u suite, uslovnu kompilaciju, okruženje.
  DOBITAK: Ako se pokaže zdravim: sabotaža se ne meri ponovo dok se ne promeni kod koji njen test zaista doseže — noćni i release prolazi padaju na ono što je dirano. Sporedni proizvod: izbor testova za TARGETED po izmenjenoj proceduri.
  RIZIK: Visok: jedno propušteno dinamičko pozivanje i keš tiho potvrđuje ono što više ne važi. Verdikti su skoro uvek DOKAZANO, pa poređenje sa ponovnim merenjem slabo otkriva grešku ključa.
  PREDUSLOVI: D1 (kill-matrica iz nedeljnog prolaza cele suite), D4
  DOKAZATI: Kill-matrica daje stotine izmerenih zavisnosti: za svaki par 'sabotaža X obara test T' procedura sabotaže X MORA biti u zatvaranju testa T; jedan promašaj znači da ključ nije zdrav. Uz to senka-režim: K1 se računa, ali se ništa ne preskače dok se ne skupi dovoljno poređenja.

### D11 Katalog po modulu, bez escape-a, i alat za presidrenje [trud M] [slabi: NE: isti rečnik, ista pravila. Marker u izvoru bi slabio — sidro koje preživi izmenu koda pretvara glasno 'sidro zastarelo' u tiho 'mutacija više ne znači isto', a 1265 markera bi ušlo u produkcioni kod.]
Rečnik SABOTAZE se sklapa iz fajlova po modulu (npr. tools/sabotaze/modDokumenta.py sa sirovim trostrukim stringovima, ili YAML blok-skalari — PyYAML 6.0.3 postoji u ovom okruženju, na Windows mašini neprovereno). Potrošači (dokaz.py, vba_check) i dalje vide isti rečnik. Uz to sabotaza.py --presidri <ime>: nađe najbliže mesto starog sidra u izmenjenom fajlu, predloži novo sidro i zamenu, a stavku označi za ponovni dokaz. Marker-komentar u VBA izvoru kao sidro se NE predlaže.
  DOBITAK: VBA tekst se lepi doslovno (danas 1474 escape-a i mešani navodnici), konflikti ostaju u jednom modulu umesto u fajlu od 735 KB, pregled i izmena čitaju desetine umesto hiljada redova. Presidrenje (oko 2,2 dnevno) dobija mehanički deo.
  RIZIK: Jednokratna velika izmena fajla koji dira skoro svaki PR (sudar sa otvorenim granama, uključujući granu B). Sirovi stringovi moraju čuvati uvlačenje i završni novi red. Nov loader je nov checker.
  PREDUSLOVI: miran trenutak bez otvorenih PR-ova nad katalogom; posle D4
  DOKAZATI: Migracija je dokaziva jednakošću: rečnik iz starog fajla == rečnik iz novih, uključujući redosled. Loader dobija self-test (duplo ime u dva fajla, fajl bez unosa, pokvaren blok) u oba smera; --proveri-sidra daje isti izlaz pre i posle.

### D12 Kontrolne sabotaže u svakom prolazu dokaza [trud S] [slabi: NE, jača]
Svaki poziv dokaz.py (i svaki nov režim: izbor, zamena modula, sesija) pusti i kontrolu koja NE SME biti priznata: placebo mutaciju koja menja samo komentar (očekivano NE OBARA NISTA). Dve zatečene priznate stavke (ljuska-rez-bez-potvrde -> NE OBARA NISTA; f8-identitet-po-broju -> PALA DRUGA TVRDNJA) služe kao iste kontrole pri uvođenju novog režima.
  DOBITAK: Dvosmerni dokaz samog dokaza u Excelu. Danas se crveni smer orkestracije meri samo lažnim runner-om; pipeline koji bi uvek govorio DOKAZANO ne bi bio uhvaćen.
  RIZIK: Jedan prolaz više po pozivu (u izboru oko 20–25 s). Placebo mora imati svoj test i tvrdnju.
  PREDUSLOVI: nema
  DOKAZATI: Kontrola koja bude priznata obara verdikt po imenu; self-test sa lažnim izlazom u kome je placebo 'oborio' tvrdnju mora dati NIJE DOKAZANO.

### D13 Privremeno, do izbora: veći cap uz strože pravilo o fajlu [trud S] [slabi: DA, blago: ista vrsta tvrdnje (grupno), ali sa do 11 partnera umesto 5 raste verovatnoća da je partner oborio tuđu tvrdnju. Pravilo 'različit fajl' to delimično vraća. Verdikt ostaje DOKAZANO (grupno).]
Za grupne prolaze cele suite (dok D4 ne uđe, i kasnije za nedeljni prolaz iz D9): --grupe 12, opciono sa pravilom 'različit fajl' umesto 'različita procedura'.
  DOBITAK: 215 -> 112 prolaza (cap 12), ili 191 uz 'različit fajl'; uz novembarska trajanja oko 14 h -> oko 7 h pre ponavljanja. Bez izmene koda za sam cap.
  RIZIK: Veća grupa znači više zaklanjanja i više ponavljanja (stopa po cap-u nije merena) i duže čekanje kad jedan član obori compile.
  PREDUSLOVI: D2(d) zbog timeout-a
  DOKAZATI: Iz knjige (D1): uporediti stopu ponavljanja i broj stavki koje grupno prođu a pojedinačno padnu, za cap 6 i cap 12 nad istim isečkom.

### D14 NE PREPORUČUJE SE: uzorkovanje umesto punog dokaza i mutant schemata [trud L] [slabi: DA. Uzorkovanje menja 'crvenih = sabotaža' u verovatnosnu tvrdnju (uzorak od k stavki otkriva bar jednu mrtvu samo sa verovatnoćom 1 − (1 − x)^k). Schemata menja ono što se meri: tvrdnja više nije o tekstualnoj mutaciji produkcionog izvora nego o grani generisanog build-a, čija ekvivalencija sa izvorom nije dokazana.]
Uzorkovanje: slučajan podskup pred release. Schemata: sve mutacije u jednom build-u iza prekidača, bez uvoza po mutaciji.
  DOBITAK: Uzorkovanje: trošak srazmeran uzorku. Schemata: nula režije po mutaciji, ali dobitak preko zamene jednog modula (D6) je samo sekunda-dve po mutaciji.
  RIZIK: Schemata: samo 54% kataloga je mehanički obmotljivo, pa bi dva mehanizma živela uporedo; generator je nov checker koji i sam traži dvosmerni dokaz; 117 sabotaža ima sidra koja se preklapaju.
  PREDUSLOVI: nema
  DOKAZATI: Ne sprovodi se; brojevi su iz S/schemata.py i iz modela.


## NEPOZNANICE
[
 "Stvarna trajanja danas: T_A (RunAllTests zelen), T_B (BFP), T_K (banka) i režija r sa raspadom po fazama. Sve je iz dokumentacije ili izvedeno iz dva zapisa u dnevniku; tests/last_run.* nije u repou, a Excel ovde ne postoji.",
 "Trajanje jednog testa i jednog izbora (tau), posebno koliko traje BFP seed koji se u izboru pušta uvek. Grana B to beleži (last_run_vreme.txt), ali nikad nije izvršena.",
 "Koliko ciljanih testova ne prolazi samo (zavisnost od redosleda). Od toga zavisi da li izbor uopšte pomaže za BFP: svaki takav test plaća promašenu bazu i prolaz cele suite. Statika (NewScenarioCode u 354 od 398 testova) je samo indicija.",
 "Stopa ponavljanja u grupi (rho) nad celim katalogom i kako raste sa cap-om; 8% je procena iz pet zapisa o malim rezovima.",
 "Da li Application.Run prima spisak testova duži od 255 znakova (cap 6 ide do 239, cap 12 do 455).",
 "Da li više Excel instanci može paralelno na istoj mašini: compile proba koristi SetFocus, AppActivate i SendKeys; watchdog je vezan za PID, ali fokus nije.",
 "Cena zamene jednog modula u otvorenoj svesci i otvaranja unapred sagrađene sveske; bez toga se dobitak D6 ne može potvrditi.",
 "Da li pun katalog danas uopšte može da vrati DOKAZANO: sedam priznatih statičkih nalaza 'tvrdnja nije CEO STATICAN literal' po komentaru samog alata daje NE OBARA SVOJ TEST, a nisu u POZNATI_NALAZI_DOKAZ (koji dokaz.py jedini čita). Izvedeno iz koda, nije pušteno.",
 "Koliko se verdikt razlikuje između izolovanog testa i pune suite (mutacija zaklonjena ili lažno potvrđena stanjem susednih testova). Meri se tek diferencijalnim prolazom iz D4.",
 "Nestabilnost i nedeterminizam: koliko prolaza pada iz razloga van mutacije (pad Excela posle oko 10 min rada je zabeležen kao neistražen), i da li ijedna tvrdnja zavisi od datuma ili locale-a.",
 "Oblik rasta: linearna projekcija (+18,7 dnevno) je iz 29 dana plitke istorije; da li trajanje suite raste srazmerno katalogu je pretpostavka (BFP je u mesec dana porastao više nego broj tvrdnji).",
 "Udeo uvoza 194 modula, compile probe i sleep-ova u 18 s režije — iz koda se vidi samo da je 4,5–7,9 s fiksnog čekanja.",
 "PyYAML na Windows mašini operatera i u hook okruženju (ovde postoji 6.0.3); bitno samo ako se za D11 izabere YAML umesto Python fajlova po modulu."
]
