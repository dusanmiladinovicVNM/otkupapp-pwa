# FAST sloj -- tools/vba_check.py (4192 linije) i PostToolUse hook (.claude/hooks/vba-check.sh, .claude/settings.json); stablo A = origin/main @1fa4937

## SAZETAK
Hook danas traje 9,2-10,3 s po Edit-u malog/srednjeg .bas fajla i 12-13,7 s za modTest/BFP, a kod malog fajla je 100% tog vremena potrošeno na stvari koje NE zavise od izmenjenog fajla (katalog sabotaža 3,7 s, collect_arities 3,0 s, popis suita 1,0 s, AMB_BIND 0,5 s, dva skupljača imena 0,8 s). Tri uzroka su čisto algoritamska: (1) sabotaza._dodela_tudjoj_proceduri za svaku od 1265 sabotaža ponovo parsira CEO mutirani fajl (214,5 MB teksta umesto 5,7 MB jedinstvenog; 856 od 1265 poziva ne može ni teoretski dati nalaz), (2) collect_arities zove _split_statements dvaput po svakoj liniji celog src-vba, znak po znak, (3) check_scalar_call čisti istu liniju tela jednom po svakom skalarnom imenu (8,3x više poziva nego linija) -- to je 32% punog prolaza. Napravio sam prototip na KOPIJI alata (original netaknut): hook 9,2-10,3 s -> 2,9-3,8 s, modTest hook 12-13 s -> 4,3-4,9 s, pun prolaz 30-34 s -> 13,4-13,9 s, --self-test 17,4 s -> 5,8 s, uz IDENTIČAN izlaz nad čistim repoom, nad dve 'prljave' kopije src-vba (261 i 434 nalaza u 13-15 kodova), nad 301.079 redova po funkciji i nad 1265 + 6000 fuzz zamena kataloga. Drugi prototip (keš činjenica po fajlu, ključ = sha1 sadržaja + otisak alata) spušta hook na 0,2 s (fajl van kataloga) / 1,1-1,7 s (fajl iz kataloga) / 3,1-3,8 s (modTest/BFP), takođe sa identičnim izlazom. Procena dnevne cene: medijana 143 VBA hunk-a dnevno x ponderisanih 11,2 s = oko 27 min čekanja dnevno (p90 dan oko 90 min; oko 14,5 h za 27 radnih dana); posle prvog talasa oko 9 min dnevno. Usput su izmerene tri rupe u onome što FAST sloj tvrdi: hook ne vidi lomove preko granice fajla (promena potpisa LogErr: hook rc=0, pun prolaz 492 ARNOST u 78 fajlova), hook se uopšte ne pokreće kad se menja tools/sabotaza.py iako komentar u vba_check.py tvrdi suprotno, i ne-ASCII bajt u fajlu koji čita katalog ruši alat traceback-om bez imena fajla i reda umesto nalaza ASCII. Trošak raste sa katalogom: isti hook je 10.10. (724 sabotaže) trajao oko 7,0 s, a 08.11. (1265) 9,2-10,3 s. Preporučen redosled: prvo jeftini egzaktni rezovi F1-F4 uz preduslov F8 (sabotaza --self-test u CI), zatim F6/F7 koji jačaju tvrdnju, a keš (F9/F10) samo ako se želi da hook hvata i lomove preko fajlova.

## CINJENICE
- [M] Hook nad MALIM .bas fajlom (modMeteo.bas, 154 B): 9,05 s u main(), raspodela po fazama (wall-clock, bez profajlera): check_katalog_sabotaza 3,68 s (40,6%), collect_arities 3,01 s (33,2%), check_popis_suita 1,00 s (11,0%), check_amb_bind_vlasnik 0,53 s (5,9%), collect_module_members 0,41 s, collect_definitions 0,41 s, check_dokaz_grupe 0,01 s; sve provere koje zavise od samog fajla zajedno 0,00 s. Ceo proces: 9,2-10,4 s kroz 8 ponavljanja; kroz pravi hook skript 9,75 s.
    DOKAZ: python3 <scratch>/fast-sloj/stage_timer.py <A>/tools --hook src-vba/modMeteo.bas -> 'rc=0 main=9.05s'; 'time python3 tools/vba_check.py --hook src-vba/modMeteo.bas' -> 10.425/10.855/9.212/9.436/10.274/9.800/9.738 s; 'printf {payload} | bash .claude/hooks/vba-check.sh' -> wall=9.754s exit=0. Alat za merenje: <scratch>/fast-sloj/stage_timer.py
- [M] cProfile udvostručuje vreme (19,4 s umesto 9-10 s) i iskrivljuje udele: pod profajlerom katalog izgleda kao 52% a collect_arities kao 22%, dok je stvarni wall-clock udeo 41% i 33%. Tabela iz zadatka (53%/21%) je zato cProfile artefakt; collect_arities je skoro jednako skup kao katalog.
    DOKAZ: python3 -m cProfile -o hook_small.prof tools/vba_check.py --hook src-vba/modMeteo.bas -> '25000008 function calls in 19.403 seconds': check_katalog_sabotaza 10.147 cum, collect_arities 4.259, check_popis_suita 2.207, check_amb_bind_vlasnik 1.386; re.Pattern.match 11.245.994 poziva. Poređenje sa stage_timer.py izlazom iznad.
- [M] Hook nad modTest.bas (1546 KB): 12,57 s u main(): katalog 3,58 s (28,5%), collect_arities 3,20 s (25,5%), check_file 3,14 s (24,9% -- od toga check_scalar_call 1,69 s, check_nedeklarisan 0,40, check_undefined 0,28, check_test_registry 0,20, check_dupli_lokal 0,17), popis 0,82 s, AMB_BIND 0,64 s, collect_module_members 0,49, collect_definitions 0,33, check_clan_forme 0,14, collect_public 0,13 (računa se, a u hook režimu se rezultat ne koristi). Ceo proces 12,0-14,9 s; BFP modul 13,7 s.
    DOKAZ: stage_timer.py <A>/tools --hook src-vba/modTest.bas -> 'rc=0 main=12.57s'; time ... --hook src-vba/modTest.bas -> 14.859/13.600/12.012/13.020 s; modBusinessFlowProTests.bas -> 13.673 s. collect_public bez upotrebe: vba_check.py:4140-4146 (publics se čita samo pod 'if not args.paths').
- [M] PUN prolaz (193 VBA fajla, 233.491 linija, 10,2 MB): 32,04 s u main(): check_file x193 = 19,74 s (61,6%), od toga check_scalar_call 10,29 s (32,1% celog prolaza), check_nedeklarisan 2,06, check_undefined 1,81, check_dupli_lokal 1,32, check_test_registry 1,12, check_array_copy 0,65, check_local_dupes 0,61, check_dead_log 0,50, check_reserved 0,44; katalog 3,82 s (11,9%); collect_arities 2,95 s (9,2%); popis 0,95; collect_public 0,80; check_clan_forme 0,75; check_storno_registar 0,59; AMB_BIND 0,57; check_rollback_red 0,53; collect_definitions 0,49; collect_module_members 0,43; check_rollback_tvrdnja 0,26.
    DOKAZ: stage_timer.py <scratch>/fast-sloj/base/tools (kopija alata, izlaz identičan stablu A: 'IDENTICAL_TO_TREE_A') -> 'rc=0 main=32.04s' + tabela; time python3 tools/vba_check.py -> 33.405/30.184/33.711 s
- [M] ZAŠTO katalog košta: sabotaza._dodela_tudjoj_proceduri se zove 1265 puta (jednom po sabotaži) i svaki put radi tekst.replace + _procedure_po_redu(mutirani) -- split celog fajla na redove, strip i dva regex-a po redu. To je 2,89 s od 3,42 s provere sidara (85%). Obrađuje se 214,5 MB teksta, a jedinstvenih je 5,72 MB (77 fajlova) -- 37,5x ponavljanja. Najteži: modDokumenta.bas 128 sabotaža x 434 KB = 56,8 MB, modOtkupUI.bas 88 x 447 KB = 40,3 MB, modAmbalaza.bas 97 x 202 KB = 20,0 MB. Samo 409 od 1265 zamena uopšte sadrži liniju-dodelu (_ZAMENA_DODELA); preostalih 856 poziva po konstrukciji vraća None posle punog parsiranja.
    DOKAZ: tools/sabotaza.py:13150-13209 (funkcija), :13183-13185 (replace + _procedure_po_redu + split), :13040-13051 (_procedure_po_redu), :13286 (poziv iz _nalazi). Merenje: python3 <scratch>/fast-sloj/katalog_timer.py <A>/tools -> 'proveri_sidra: 3.42s; 2.89s x1265 _dodela_tudjoj_proceduri; 2.20s x1265 _procedure_po_redu; 0.28s _tela_testova; suma (sabotaza x velicina fajla) = 214.5 MB; suma velicina jedinstvenih = 5.72 MB; sabotaza cija ZAMENA ima bar jednu liniju-dodelu: 409 od 1265'
- [M] Odluka da se ne kešira doneta je za katalog 5x manji i komentari o ceni više ne važe: 'Kes nad originalnim tekstom je time otpao. 251 unos je premalo da bi ustedu vredelo platiti netacnoscu', a na dva mesta piše da provera 'traje sekundu'; hook skript tvrdi da se greške 'vide u istoj sekundi'. Izmereno je 3,4-3,9 s za proveru sidara i 9-14 s za hook.
    DOKAZ: tools/sabotaza.py:13181-13182, tools/sabotaza.py:12559, tools/vba_check.py:4021, .claude/hooks/vba-check.sh:7; merenje: 'time python3 tools/sabotaza.py --proveri-sidra' -> 3.532 s / 3.903 s
- [M] ZAŠTO collect_arities košta 3,0 s: za SVAKU logičku liniju celog src-vba zove _split_statements DVAPUT (jednom za uslov, jednom za vrednost), a to je petlja znak po znak u Pythonu; 401.088 poziva po jednom hook-u. Rezultat treba samo za linije koje su deklaracija procedure (6.171 ime); ostalih ~95% rada se baca.
    DOKAZ: tools/vba_check.py:514 ('glava = _split_statements(line)[0][0] if _split_statements(line) else line'), :464-490 (petlja po znaku); cProfile hook_small: '401088 ... vba_check.py:464(_split_statements)'; diff_p2.py: 'collect_arities: imena 6171'
- [M] ZAŠTO check_scalar_call košta 10,3 s u punom prolazu: za svaku proceduru, za svako skalarno ime (parametri + Dim), za svaku liniju tela ponovo računa _clean(line) = _strip_strings + _strip_comment, obe petlje znak po znak. 1.927.489 poziva _clean i 2.125.820 poziva _strip_strings za 233.491 fizičku liniju (8,3x).
    DOKAZ: tools/vba_check.py:964-967 (petlja ime x linija sa hit.search(_clean(line))), :909-922; cProfile full.prof: '1927489 ... _clean', '2125820 17.232 ... 28.534 _strip_strings', '193 ... 34.320 check_scalar_call' od 85.493 s
- [M] Popis suita i AMB_BIND rade isti posao dvaput: obe provere zovu logicke_izjave() nad (skoro) svim fajlovima -- 372 poziva po jednom hook-u (193 iz vba_gate.skeniraj + 177 ne-test fajlova iz check_amb_bind_vlasnik), a _izjava_deo ide znak po znak kroz svaki red (388.186 poziva). Reč 'BindSourceDocument' postoji u samo 7 od 193 fajla.
    DOKAZ: tools/vba_gate.py:290-335 (skeniraj), tools/vba_check.py:1339-1379 (check_amb_bind_vlasnik), :129-152 (_izjava_deo); cProfile hook_small: '372 ... logicke_izjave', '388186 ... _izjava_deo'; 'grep -il bindsourcedocument src-vba/* | wc -l' -> 7
- [M] OD ČEGA ZAVISI ISHOD svake provere u hook režimu nad fajlom F: (a) samo F: ASCII, KRAJ_REDA, DEKLARACIJA, REZERVISANO, DUPLIKAT_LOKALNI, DUPLI_LOKAL, ZAKLONJENO, MRTAV_LOG, ODSECEN, KOPIJA_NIZA, REGISTAR, STORNO_PROGUTAN, NEDEKLARISAN, PROCEDURA_VELIKA, ROLLBACK_TVRDNJA_RED; (b) F + jedan registar: PORUKA i ROLLBACK_TVRDNJA (modPoruke.bas), STORNO_REGISTAR (modSchemaGuard.bas), CLAN_FORME (sve .frm, ima ih 1); (c) F + SVI fajlovi: NEDEFINISAN/ARNOST preko collect_definitions/collect_module_members/collect_arities -- i to samo kad je F .bas (check_undefined za ostale odmah vraća []); (d) GLOBALNO, nezavisno od F: katalog sabotaža (tools/sabotaza.py + 77 sabotiranih + 3 test fajla = 80 od 193 fajla), popis suita (svih 193 + SUITES iz run_vba.py + SUITE_VAN_KAPIJA iz vba_gate.py), AMB_BIND (177 ne-test fajlova), SEMA_REGISTAR (modConfig.bas + modSchema.bas), pravila grupisanja (samo tools/dokaz.py, 0,01 s).
    DOKAZ: tools/vba_check.py:4115-4117 i :4157-4163 (šta main zove i sa kojim argumentima), :2390-2409 (check_file), :530-531 (check_undefined samo .bas), :1529-1547, :1453-1494, :1600-1630, :1663-1693, :1107-1157; tools/sabotaza.py:95 (_TEST_FAJLOVI), :13222-13230; broj fajlova: <scratch>/fast-sloj/katalog_fajlovi.txt (80)
- [M] --self-test traje 17,4-17,8 s, a 152 slučaja zajedno koštaju ispod 0,2 s. Ostatak su DVA poziva main([put]) ('ceo CLI' slučajevi) od kojih svaki plaća sve globalne kapije nad pravim repoom: katalog 2 x 3,65 s = 7,31 s (41%), collect_arities 5,93 s (33%, 3 poziva), popis 1,81 s, AMB_BIND 1,08 s (12 poziva), collect_module_members 0,88, collect_definitions 0,69; check_file x82 = 0,08 s. Usput dvaput ispisuje ceo izveštaj kataloga.
    DOKAZ: tools/vba_check.py:3877 i :3901 (main([put])); stage_timer.py <scratch>/fast-sloj/base/tools --self-test -> 'rc=0 main=17.84s; 7.31s x2 check_katalog_sabotaza; 5.93s x3 collect_arities; 0.08s x82 check_file'; time ... --self-test -> 17.378 s
- [M] PROTOTIP 1 (F1+F2+F3+F4, samo egzaktne brze putanje, bez keša): hook mali fajl 9,2-10,3 s -> 2,9-3,8 s; hook modTest.bas 12,0-13,0 s -> 4,3-4,9 s; hook BFP 13,7 s -> 5,1 s; hook modDokumenta.bas 9,4-10,3 s -> 3,4-3,7 s; pun prolaz 30,2-33,7 s -> 13,4-13,9 s; --self-test 17,4 s -> 5,8 s; sabotaza.py --proveri-sidra 3,5 s -> 0,93 s (sama funkcija nad katalogom 3,10 s -> 0,24 s).
    DOKAZ: Kopije: <scratch>/fast-sloj/base/tools (original) i <scratch>/fast-sloj/proto/tools (izmenjeno skriptama patch_p1.py i patch_p2.py). Izlaz: 'base hook modMeteo wall=9.212s / proto 2.938s; modTest 12.012 / 4.336; modDokumenta 9.396 / 3.363; pun prolaz 30.184 / 13.744'; 'base --self-test: wall=17.378s, proto: 5.795s'; 'base proveri-sidra wall=3.532s, proto 0.932s'; tri ponavljanja: base 10.274/9.800/9.738, proto 2.965/3.784/3.106
- [M] Prototip 1 daje IDENTIČAN izlaz originalu: (1) čist repo -- stdout, stderr i rc u sva tri režima; (2) dve 'prljave' kopije src-vba sa 373 i 390 nasumičnih povreda: 261 nalaz u 13 kodova i 434 nalaza u 15 kodova, bajt-identično, plus hook nad 8 prljavih fajlova; (3) po funkciji: 301.079 redova (147.812 različitih redova iz src-vba + 153.267 fuzz) za _izjava_deo, _strip_comment, _split_statements, _strip_strings, _clean -- 0 razlika; collect_arities rečnik jednak (6171 ime); logicke_izjave nad 193 fajla 0 razlika; (4) katalog: 1265 pravih unosa + 6000 fuzz zamena (2199 sa nalazom, 11 vrsta zamena, tri oblika sidra) -- 0 razlika; (5) vba_check --self-test (152), sabotaza --self-test (43) i vba_gate --self-test identični.
    DOKAZ: Izlazi: 'IDENTICNO: hook.modMeteo.bas.out / hook.modTest.bas.out / hook.modDokumenta.bas.out / full.out / full.err'; 'PRLJAV_STDERR_IDENTICAN ... vba_check: 261 nalaza (ARNOST=15, ASCII=14, CLAN_FORME=48, ...)'; 'PRLJAV_777_IDENTICAN: vba_check: 434 nalaza (ARNOST=158, ...)'; diff_p2.py: 'poredjenja=301079 razlika=0' x5; diff_p1.py: '(1) katalog: 1265 unosa, razlika=0 ... (2) fuzz: 6000 zamena, razlika=0, ne-None=2199 (37%)'; 'SELFTEST_IDENTICNO', 'SABOTAZA_SELFTEST_IDENTICNO', 'GATE_SELFTEST_IDENTICNO'. Skripte: <scratch>/fast-sloj/diff_p1.py, diff_p2.py, napravi_prljav.py
- [M] Dokaz u crvenom smeru za sam prototip: tri namerna kvara nove _dodela_tudjoj_proceduri obaraju postojeći sabotaza.py --self-test (5, 6 i 2 od 43 slučaja 'NE MERI'). Kod tokenizera slika je mešovita: kvar u _strip_strings i _split_statements obara vba_check --self-test (7 i 3 od 152), kvar u _izjava_deo hvata tek vba_gate --self-test ('suite u obliku Prelom je nevidljiva'), a kvar u _strip_comment (apostrof unutar stringa pre komentara) ostavlja vba_check --self-test ZELEN (rc=0) -- vidi se samo kao 8 lažnih ARNOST nad pravim repoom i kao 12.807 razlika u diferencijalnom testu.
    DOKAZ: <scratch>/fast-sloj/kvar/tools: 'self-test: 5 od 43 slucajeva NE MERI rc=1', '6 od 43', '2 od 43'; 'kvar strings: self-test rc=2 (7 od 152)', 'kvar split: rc=2 (3 od 152)', 'kvar izjava: self-test rc=0' + 'vba_gate --self-test: 1 nalaza POPIS: suite u obliku Prelom je nevidljiva'; 'kvar A (samo _strip_comment) self-test rc=0' + 'kvar A pun prolaz nad cistim repoom rc=2, 8 nalaza (ARNOST=8)'; diff_p2.py kvar -> '_strip_comment poredjenja=301079 razlika=12807'
- [M] PROTOTIP 2 (keš činjenica po fajlu: ključ = (ime fajla, sha1 sadržaja) + sha1 pet alata; keširaju se definicije, članovi modula, sirova arnost, deklaracije za popis, pozivi BindSourceDocument i presuda kataloga): hook bez ijedne izmene 0,23 s; izmenjen fajl van kataloga 0,18-0,24 s; izmenjen fajl iz kataloga 1,1-1,7 s (katalog se ponovo vrti, 0,8-1,06 s); modTest/BFP 3,1-3,8 s (katalog 1,06 + check_file 1,38 + činjenice 0,41); hladan keš 3,5 s. Čitanje i sha1 svih 194 fajla (10,5 MB) = 0,012 s; keš fajl 1,3 MB.
    DOKAZ: <scratch>/fast-sloj/proto2/tools (patch_p4.py). Izlazi: 'hladan kes wall=3.502s; topao kes, nista izmenjeno wall=0.226s'; '(a2) modCenovnik.bas VAN kataloga: base 8.789s -> proto2 0.180s'; 'modDokumenta.bas: base=10.323s P1+P2=3.712s +kes=2.255s/1.615s'; 'modTest.bas: 13.020 / 4.435 / 3.850 / 3.532'; 'modBusinessFlowProTests.bas: 13.673 / 5.098 / 3.467 / 3.102'; stage_timer2.py: 'main=3.10s: _katalog_kroz_kes 1.06, check_file 1.38, _cinjenice_fajla 0.41'; 'citanje + sha1 svih 194 fajlova (10.5 MB): 0.012s'
- [M] Prototip 2 daje identičan izlaz originalu u 16 scenarija izmene (čista izmena; NEDEFINISAN+ARNOST+REZERVISANO+PORUKA; +ne-ASCII; AMB_BIND+nova suite; dirnuto sidro sabotaže -> KATALOG; vraćanje izmene; promena potpisa u DRUGOM fajlu -> 7 ARNOST kod pozivaoca; brisanje i vraćanje fajla) i nad prljavim korpusom u punom prolazu sa hladnim i toplim kešom (438 redova stderr) i u hook-u nad 10 prljavih fajlova.
    DOKAZ: Izlazi: '(c) ... izlaz=IDENTICNO rc=2 kodovi=[1 AMB_BIND_VLASNIK 1 NEPOKRETANA]', '(d) modStorno.bas: sidro sabotaze dirnuto ... IDENTICNO rc=2 [1 KATALOG:]', '(e) ... vracen ... IDENTICNO rc=0', '(f) pozivalac modAdmin.bas posle izmene potpisa LogErr ... IDENTICNO rc=2 [7 ARNOST]', '(g) fajl obrisan ... IDENTICNO'; 'proto2 prljav(777) pun prolaz, hladan kes: 14.554 IDENTICNO; topao kes: 10.466 IDENTICNO (438 redova stderr)'; 'hook nad 10 prljavih fajlova (kes): ok=1'
- [M] RUPA 1 -- hook ne vidi lomove preko granice fajla. Hook proverava samo izmenjeni fajl, a DUPLIKAT (isto Public ime u dva modula) se u hook režimu uopšte ne proverava. Izmena potpisa LogErr u modLogError.bas (dodat obavezan parametar): hook nad tim fajlom vraća rc=0, a pun prolaz prijavljuje 492 ARNOST nalaza u 78 fajlova. Preimenovanje iste procedure: hook rc=0, pun prolaz 552 NEDEFINISAN u 82 fajla. Te greške se danas vide tek u punom prolazu (30+ s) ili u CI.
    DOKAZ: tools/vba_check.py:4132-4136 (check_file samo nad 'files') i :4144-4146 ('Duplikate trazimo samo kad se gleda ceo src-vba'). Merenje nad kopijom src-vba originalnim alatom: 'HOOK nad izmenjenim modLogError.bas (original alat): rc=0; PUN PROLAZ: rc=2; vba_check: 492 nalaza (ARNOST=492); 78 fajlova'; 'HOOK posle preimenovanja LogErr: rc=0; PUN PROLAZ: rc=2; 552 nalaza (NEDEFINISAN=552); 82 fajla'
- [M] RUPA 2 -- hook se NE pokreće kad se menja katalog. vba_check.py tvrdi da je putanja 'hook nad tools/sabotaza.py' zatvorena ('Ovde se to vise ne moze zaobici'), ali hook skript izlazi sa 0 za svaki fajl koji nije .bas/.cls/.frm/.doccls pre nego što pozove vba_check. Izmereno: hook nad tools/sabotaza.py i tools/dokaz.py izlazi sa 0 za 63-69 ms, bez provere. Greška u novom unosu kataloga se vidi tek pri sledećoj VBA izmeni ili u CI.
    DOKAZ: tools/vba_check.py:4104-4114 (komentar); .claude/hooks/vba-check.sh:54-57 (case ... *) exit 0); 'printf {"tool_input":{"file_path":"<A>/tools/sabotaza.py"}} | bash .claude/hooks/vba-check.sh' -> 'exit=0 wall=0.063s'; isto za tools/dokaz.py -> 'exit=0 wall=0.069s'
- [M] RUPA 3 -- ne-ASCII bajt u fajlu koji čita katalog (80 fajlova, uključujući modTest.bas i sve jezgro) ruši vba_check traceback-om PRE nego što provera ASCII stigne da prijavi nalaz: sabotaza._procitaj čita sa errors='strict', a check_katalog_sabotaza hvata izuzetak samo oko učitavanja modula. Izlaz je rc=1 i 'UnicodeDecodeError ... position 176168' bez imena fajla i bez reda. Hook i dalje blokira (izlaz nije 0), ali poruka ne kaže gde je greška -- a ASCII je provera broj 1 zbog koje hook postoji. Za fajl van kataloga nalaz ASCII stiže uredno.
    DOKAZ: tools/sabotaza.py:12478-12483 (_procitaj, errors='strict'), tools/vba_check.py:4028-4033 (try samo oko exec_module). Reprodukcija nad kopijom: 'Traceback ... sabotaza.py line 12676 _imena_testova ... _procitaj ... UnicodeDecodeError: ascii codec cant decode byte 0xc5 in position 176168; rc=1'; u hook režimu nad modConfig.bas isto (rc=1, traceback); kontrola: '(b3) modCenovnik.bas + ne-ASCII (fajl van kataloga) ... rc=2 kodovi=[... 1 ASCII ...]'
- [M] sabotaza.py --self-test (43 slučaja, 0,65 s) -- dvosmerni dokaz za proveri_sidra, najskuplju stavku hook-a -- ne vrti se ni u CI ni kroz vba_check; pominje se samo u dokumentima kao ručni korak. dokaz.py --self-test se vrti (kroz check_dokaz_grupe), vba_gate --self-test je u CI.
    DOKAZ: 'grep -c sabotaza .github/workflows/static.yml' -> 0; static.yml:79-86 (vba_check i vba_check --self-test), :188-191 (vba_gate); tools/vba_check.py:4022-4033 zove samo modul.proveri_sidra; 'time python3 tools/sabotaza.py --self-test' -> 0.653 s, 'self-test: cisto (43 slucajeva)'
- [M] Cena hook-a raste sa katalogom i izvorom. Isti hook nad malim fajlom, alat i src-vba iz tog dana, ova mašina: 10.10.2026. (724 sabotaže, 8,5 MB) 6,8-7,2 s; 20.10. (886) 7,0-8,3 s; 30.10. (1070, 9,3 MB) 8,5-8,9 s; 08.11. (1265, 10,2 MB) 9,1-10,3 s. Provera sidara: 2,2 s -> 3,5-3,9 s. Pun prolaz: 27,0 s -> 30-34 s. Šum merenja na ovoj mašini je oko +-10% (load average 0,8-0,9, verovatno drugi agenti).
    DOKAZ: git archive <commit> src-vba tools/... u <scratch>/fast-sloj/stari/<datum> (commit-i 09d9c46, 8e970d9, 5a9bd86), pa time: '2026-10-10: sabotaza=724 src-vba=8472KB hook=7.196s/6.774s proveri-sidra=2.206s pun=27.017s'; '2026-10-20: 886 ... 7.037/8.318 ... 2.149 ... 29.846'; '2026-10-30: 1070 ... 8.916/8.495 ... 4.035 ... 31.025'; '2026-11-08: 1265 ... 9.091 ... 3.903 ... 33.711'
- [M] Obim VBA izmena (plitka istorija, 09.10-08.11.2026): 192 od 493 ne-merge commit-a dira VBA; 4.848 hunk-ova (4.657 bez jednog masovnog commit-a od 191 fajla); 27 radnih dana sa VBA izmenom; hunk-ova po danu: medijana 143, prosek 180, p90 480, max 605; po PR-u (neto diff 53 PR-a): medijana 46, prosek 74, p90 161, max 363. Raspodela hunk-ova: modTest+BFP 30% (883 + 529), ostali fajlovi iz kataloga 64%, van kataloga 6%, .frm/.cls/.doccls ukupno 7 hunk-ova.
    DOKAZ: python3 <scratch>/fast-sloj/git_stat.py (git log --no-merges --since=2026-10-09 -p --unified=0 -- src-vba) -> 'ne-merge commit-a od 09.10: 493; od toga dira VBA: 192; hunks ... zbir=4848 medijana=11.0; radnih dana sa VBA izmenom: 27; hunk-ova po danu: medijana=143 prosek=180 p90=480 max=605; PR merge-ova: 62; sa VBA izmenom: 53; hunk-ova po PR-u: medijana=46 prosek=74 p90=161'; 'modTest+BFP: 1412 (30%); ostali fajlovi iz kataloga: 2977 (64%); van kataloga: 268 (6%)'
- [M] Funkcije koje prototip dira (sabotaza._dodela_tudjoj_proceduri i _procedure_po_redu; vba_check._split_statements, _strip_comment, _strip_strings, _izjava_deo, collect_arities, check_scalar_call, check_amb_bind_vlasnik, main) su bajt-identične u stablu A i na grani B (9d6e780), a commit grane B u tools/sabotaza.py samo dodaje 34 linije na drugom mestu. Rez nad FAST slojem se zato ne sudara sa granom 'izbor testova'.
    DOKAZ: diff <(sed -n '/^def X(/,/^def /p' A/tools/...) <(... B/tools/...) za svih 10 funkcija -> 'A==B'; 'git -C /home/user/otkupapp-pwa diff --stat d19c2da 9d6e780 -- tools/vba_check.py tools/sabotaza.py .claude' -> 'tools/sabotaza.py | 34 +'
- [M] Oba stabla su ostala netaknuta; sav rad je u scratch direktorijumu nad kopijama alata i kopijom src-vba (sabotaza --self-test privremeno piše modOtkupUI.bas.selftest.tmp u src-vba, pa je namerno puštan samo nad kopijom).
    DOKAZ: 'git status --short | wc -l' u stablu A -> 0 (posle svih merenja), u /home/user/otkupapp-pwa -> 0; tools/sabotaza.py:13691-13701 (kopija + os.remove); 'diff -rq <A>/src-vba <scratch>/fast-sloj/srcvba-kopija' -> KOPIJA_IDENTICNA

## TROSKOVI
[
 {
  "sta": "PostToolUse hook po jednom Edit/Write nad VBA fajlom (stanje danas)",
  "koliko": "9,2-10,3 s za mali/srednji .bas; 10,1-10,9 s za veće module iz kataloga; 12,0-13,7 s za modTest.bas i BFP. Ponderisano po raspodeli hunk-ova (30% test moduli / 64% katalog / 6% ostalo): oko 11,2 s",
  "koliko_cesto": "posle SVAKOG Edit/Write/MultiEdit nad .bas/.cls/.frm/.doccls",
  "izmereno": true
 },
 {
  "sta": "Deo hook-a koji ne zavisi od izmenjenog fajla (katalog + popis + AMB_BIND + tri skupljača imena)",
  "koliko": "9,05 s od 9,05 s kod malog fajla (100%); 9,4 s od 12,6 s kod modTest.bas (75%)",
  "koliko_cesto": "svaki hook",
  "izmereno": true
 },
 {
  "sta": "Provera sidara kataloga (check_katalog_sabotaza -> proveri_sidra)",
  "koliko": "3,4-3,9 s, od toga 2,89 s je ponovno parsiranje celog fajla po sabotaži (214,5 MB umesto 5,7 MB)",
  "koliko_cesto": "svaki hook, svaki pun prolaz, dvaput u --self-test; raste linearno sa (broj sabotaža x veličina fajla)",
  "izmereno": true
 },
 {
  "sta": "collect_arities (dupli _split_statements nad svakom linijom)",
  "koliko": "3,0-3,2 s",
  "koliko_cesto": "svaki hook i pun prolaz; tri puta u --self-test",
  "izmereno": true
 },
 {
  "sta": "check_scalar_call u punom prolazu",
  "koliko": "10,3 s od 32 s (32%); 1,7 s u hook-u nad modTest.bas",
  "koliko_cesto": "svaki pun prolaz (CI na svakom PR-u, ručno pred commit)",
  "izmereno": true
 },
 {
  "sta": "Pun prolaz vba_check",
  "koliko": "30,2-33,7 s",
  "koliko_cesto": "CI po PR-u/push-u; ručno kao FAST nivo",
  "izmereno": true
 },
 {
  "sta": "vba_check --self-test",
  "koliko": "17,4-17,8 s, od čega oko 17,6 s nisu slučajevi nego dva puna prolaza globalnih kapija",
  "koliko_cesto": "CI po PR-u; ručno pri izmeni checkera",
  "izmereno": true
 },
 {
  "sta": "Dnevno čekanje na hook (procena: 1 Edit po hunk-u u commit diff-u)",
  "koliko": "medijana 143 x 11,2 s = oko 27 min dnevno; prosečan dan (180) oko 34 min; p90 dan (480) oko 90 min; ukupno 4.657 x 11,2 s = oko 14,5 h za 27 radnih dana",
  "koliko_cesto": "svaki radni dan sa VBA izmenom; množi se brojem paralelnih sesija",
  "izmereno": false
 },
 {
  "sta": "Isto čekanje posle prvog talasa (F1-F4, izmeren prototip: ponderisano oko 3,9 s po Edit-u)",
  "koliko": "oko 9 min dnevno (medijana), oko 31 min na p90 dan, oko 5,0 h za 27 dana",
  "koliko_cesto": "isto",
  "izmereno": false
 },
 {
  "sta": "Isto čekanje sa kešom po fajlu (F9, izmeren prototip: ponderisano oko 2,0 s po Edit-u)",
  "koliko": "oko 5 min dnevno (medijana), oko 16 min na p90 dan, oko 2,6 h za 27 dana",
  "koliko_cesto": "isto",
  "izmereno": false
 },
 {
  "sta": "Rast cene hook-a bez ikakve izmene",
  "koliko": "oko 7,0 s (10.10.) -> 9,1-10,3 s (08.11.): +30-45% za 29 dana; provera sidara 2,2 s -> 3,5-3,9 s",
  "koliko_cesto": "trajno, uz svaki novi unos u katalog i svaki novi red u jezgru",
  "izmereno": true
 },
 {
  "sta": "Režija samog hook skripta (bash + 2 pomoćna Python procesa pre vba_check)",
  "koliko": "0,06-0,08 s na ovom Linux-u za putanju koja nije VBA; na Windows/Git Bash nije mereno",
  "koliko_cesto": "svaki Edit/Write bilo kog fajla",
  "izmereno": true
 },
 {
  "sta": "Trošak održavanja: tri odvojena prolaza kroz src-vba za imena (collect_definitions, collect_module_members, collect_arities) + dva identična parsiranja u logičke izjave (popis, AMB_BIND) + 15 provera po fajlu, svaka sa svojom petljom znak po znak",
  "koliko": "oko 4,7 s od 9 s hook-a; pet nezavisnih 'tokenizera' stringa/komentara (_strip_comment, _strip_strings, _izjava_deo, _bez_teksta, _split_statements) koji moraju ostati usaglašeni",
  "koliko_cesto": "pri svakoj izmeni checkera",
  "izmereno": true
 }
]

## PREDLOZI
### F1 Katalog: _dodela_tudjoj_proceduri bez ponovnog parsiranja celog fajla po sabotaži [trud S] [slabi: NE -- rezultat je identičan po konstrukciji i po merenju; pravilo 'zamka 10' ostaje isto.]
Dve egzaktne izmene u tools/sabotaza.py:13150-13209. (a) Rani izlaz: ako nijedna linija zamene ne odgovara _ZAMENA_DODELA, funkcija svakako vraća None -- 856 od 1265 poziva otpada bez parsiranja. (b) Za preostalih 409: 'događaji' zdravog fajla (redovi koji otvaraju/zatvaraju proceduru) računaju se jednom po fajlu; za mutirani tekst se ponovo klasifikuju samo redovi koje zamena dodiruje, ostali se pomeraju za razliku u broju redova, a isti automat kao u _procedure_po_redu ide nad događajima. Simboli se i dalje čitaju iz MUTIRANOG koda (razlog zbog koga je raniji keš odbačen ostaje ispoštovan). Prototip: <scratch>/fast-sloj/patch_p1.py.
  DOBITAK: Izmereno: funkcija nad celim katalogom 3,10 s -> 0,24 s; sabotaza.py --proveri-sidra 3,53 s -> 0,93 s; hook -2,7 do -2,9 s (katalog 3,7 s -> 0,8 s); isto u punom prolazu i dvaput u --self-test. Cena više ne raste sa (sabotaža x veličina fajla), pa prestaje i rast od oko +1,5 s mesečno.
  RIZIK: Nizak. Jedina suptilnost je granica mutiranog regiona kad sidro ne završava prelomom reda; pokriveno fuzz-om u tri oblika sidra (sa prelomom, bez, pola reda).
  PREDUSLOVI: F8a (sabotaza.py --self-test u CI), da dokaz u crvenom smeru za ovu funkciju bude automatski, a ne ručni. Izmena je u tools/, nije u .claude/.
  DOKAZATI: Već pokazano na prototipu: (1) diff_p1.py -- 1265 pravih unosa + 6000 fuzz zamena (2199 sa nalazom, 11 vrsta), 0 razlika prema originalu; (2) sabotaza.py --self-test identičan (43 slučaja) i --proveri-sidra identičan izlaz; (3) crveni smer: tri namerna kvara nove funkcije obaraju self-test po imenu ('dodela/Sub nema povratnu vrednost', 'dodela/zamena UVODI deklaraciju', 'dodela/zamena UKLANJA deklaraciju'). Pri uvođenju staru funkciju zadržati jedan rez kao referencu i u self-testu tvrditi jednakost nad _DODELA_CASES + celim katalogom.

### F2 collect_arities i tokenizeri: egzaktne brze putanje umesto petlje znak po znak [trud S] [slabi: NE]
U tools/vba_check.py: (a) collect_arities (:514) -- jedan poziv _split_statements umesto dva, i predfilter: linija bez reči sub/function/property ne može pogoditi PROC_DEF/DECLARE_DEF; (b) _split_statements -- linija bez dvotačke je jedna naredba; (c) _strip_comment -- bez apostrofa nema komentara, a apostrof pre svakog navodnika je početak komentara; (d) _strip_strings -- split po navodniku umesto petlje; (e) _izjava_deo -- bez apostrofa je stanje stringa parnost broja navodnika. Petlja ostaje kao spora putanja za preostale slučajeve. Prototip: .../fast-sloj/patch_p2.py.
  DOBITAK: Izmereno: collect_arities 3,0 s -> 0,57 s; popis suita 1,0 s -> 0,46 s; AMB_BIND 0,53 s -> 0,32 s; check_undefined u punom prolazu 1,81 s -> 1,03 s. Zajedno sa F1: hook 9,2-10,3 s -> 2,9-3,8 s.
  RIZIK: Nizak do srednji: pet funkcija su osnova gotovo svih provera, pa bi greška u brzoj putanji tiho promenila nalaze. Zato dokaz mora biti diferencijalni, ne samo self-test (v. F8b: kvar u _strip_comment danas ostavlja self-test zelenim).
  PREDUSLOVI: F8b (self-test slučajevi za tokenizere). Poželjno uvesti zajedno sa F1 i F3 u jednom PR-u nad tools/.
  DOKAZATI: Pokazano: diff_p2.py -- 301.079 redova po funkciji (svi različiti fizički i logički redovi src-vba + 153.267 fuzz redova od navodnika, apostrofa, dvotački, zagrada, ' _'), 0 razlika; collect_arities rečnik jednak (6171 ime); izlaz identičan nad čistim repoom i nad dve prljave kopije (261 i 434 nalaza); crveni smer: namerni kvar brze putanje daje 12.807 razlika u diferencijalnom testu. Predlog: diferencijalni test (stara petlja kao referenca u test kodu) ući u vba_check --self-test nad fiksnim skupom neprijateljskih redova.

### F3 check_scalar_call: telo procedure se čisti jednom, ne jednom po skalarnom imenu [trud S] [slabi: NE]
U tools/vba_check.py:964-967 _clean(line) se računa za svaki par (ime, linija). Izračunati listu očišćenih linija jednom po proceduri (lenjo, tek kad procedura ima bar jedno skalarno ime) i nad njom vrteti regex za svako ime. Redosled nalaza ostaje isti (ime, pa linija).
  DOBITAK: Izmereno (zajedno sa brzim putanjama iz F2): check_scalar_call u punom prolazu 10,29 s -> 2,40 s; pun prolaz 30,2-33,7 s -> 13,4-13,9 s; hook nad modTest.bas 12,0-13,0 s -> 4,3-4,9 s; CI korak vba_check oko -18 s.
  RIZIK: Vrlo nizak -- ista funkcija nad istim ulazom, samo zapamćena.
  PREDUSLOVI: Nema.
  DOKAZATI: ZAKLONJENO slučajevi u self-testu (identičan izlaz, 152 slučaja); prljav korpus: ZAKLONJENO=10 nalaza identično u oba semena; dvosmerno: izbaciti liniju iz očišćene liste -> ZAKLONJENO slučaj mora pasti po imenu.

### F4 Sitni egzaktni rezovi globalnih provera u hook režimu [trud S] [slabi: NE]
(a) check_amb_bind_vlasnik: fajl koji ne sadrži reč 'bindsourcedocument' (bez obzira na velika/mala slova) ne može dati nalaz -- preskočiti parsiranje; reč ima 7 od 193 fajla. (b) collect_definitions/collect_module_members/collect_arities računati lenjo, samo kad je bar jedan proveravani fajl .bas (check_undefined za .frm/.cls/.doccls odmah vraća []). (c) collect_public ne zvati u hook režimu (rezultat se koristi samo bez putanja). (d) sabotaza._tela_testova: pattern.search(tekst, pos) umesto isecanja tekst[m.start():] po proceduri.
  DOBITAK: (a) izmereno 0,32 s -> 0,07 s po hook-u; (b) procena -1,3 s za izmene .frm/.cls/.doccls (retke: 7 hunk-ova za mesec); (c) izmereno 0,13 s kod modTest.bas; (d) procena 0,32 s -> oko 0,15 s. Ukupno još oko 0,4-0,5 s po tipičnom hook-u: 2,9 s -> oko 2,5 s.
  RIZIK: Nizak. (a) počiva na tome da se identifikator ne može prelomiti preko reda -- važi i u VBA i u logicke_izjave (spaja delove razmakom).
  PREDUSLOVI: Nema.
  DOKAZATI: (a) pokazano: self-test identičan (AMB_BIND_CASES, 10 slučajeva), scenario sa novim pozivom van liste daje isti AMB_BIND_VLASNIK nalaz, prljav korpus identičan. Za (b)-(d): identičan izlaz nad čistim i prljavim korpusom + po jedan self-test slučaj kroz ceo CLI za .cls fajl.

### F5 --self-test: globalne kapije jednom po procesu, ne po svakom 'ceo CLI' slučaju [trud S] [slabi: NE]
Dva slučaja 'ceo CLI' (vba_check.py:3877, :3901) zovu main([put]) i time dvaput plaćaju katalog, popis, AMB_BIND i tri skupljača nad pravim repoom, a služe samo da dokažu da CLI zove proveru i vraća exit 2. Zapamtiti presudu globalnih kapija i skupljene tabele u okviru jednog procesa (ulaz se između dva poziva ne menja), i ne ispisivati izveštaj kataloga dvaput.
  DOBITAK: Izmereno polazno: 17,4-17,8 s, od toga oko 17,6 s su ta dva poziva. Sa F1-F3 već izmereno 5,8 s; sa pamćenjem u procesu procena oko 3 s. CI korak 'vba_check self-test' sa 17 s na 3-6 s.
  RIZIK: Nizak; paziti da se pamćenje ne primeni između slučajeva koji namerno podmeću različit src_dir (oni već idu preko parametara funkcija, ne preko main).
  PREDUSLOVI: F1-F3 (inače dobitak ostaje, ali manji deo).
  DOKAZATI: Izlaz self-testa identičan (152 slučaja); dvosmerno: izbaciti check_storno_registar iz main -> slučaj 'ceo CLI: STORNO_REGISTAR nad pravim registrom' mora i dalje pasti po imenu.

### F6 Ne-ASCII u fajlu kataloga mora dati nalaz ASCII sa putanjom i redom, ne traceback [trud S] [slabi: NE -- jača: nalaz koji danas ne stiže počinje da stiže.]
check_katalog_sabotaza hvata izuzetak samo oko učitavanja modula; _procitaj (errors='strict') baca UnicodeDecodeError iz proveri_sidra pre nego što check_ascii stigne na red. Rešenje bez promene semantike kataloga: u vba_check-u uhvatiti UnicodeDecodeError oko proveri_sidra, prijaviti 'KATALOG: nije proveren -- ne-ASCII u izvoru' kao tvrd nalaz (rc 2) i NASTAVITI na provere po fajlu, da nalaz ASCII izađe sa fajlom i redom.
  DOBITAK: Nije brzina nego tačnost poruke: danas hook za najčešću mehaničku grešku (dijakritik u .bas) nad 80 fajlova vraća 20 redova traceback-a bez imena fajla; posle izmene jedan red 'fajl:red ASCII'. Štedi bar jedan krug traženja po incidentu.
  RIZIK: Vrlo nizak; rc ostaje različit od 0 u oba slučaja.
  PREDUSLOVI: Nema.
  DOKAZATI: Novi self-test slučaj kroz ceo CLI: kopija jednog fajla iz kataloga sa bajtom 0xC5 -> očekuje se exit 2 i 'ASCII' u izlazu; pre izmene taj slučaj pada (traceback, rc=1) -- reprodukovano u ovoj analizi.

### F7 Hook da pokrije i izmene samog kataloga i alata (zaseban process PR nad .claude/) [trud S] [slabi: NE -- jača: provera se pokreće tamo gde greška nastaje.]
U .claude/hooks/vba-check.sh dodati granu: za tools/sabotaza.py, tools/dokaz.py, tools/vba_gate.py i tools/run_vba.py pozvati 'vba_check.py --hook <putanja>' -- vba_check tada nema VBA fajlova i vrti samo tri kapije nad alatima (taj put već postoji: vba_check.py:4115-4121). Time ponašanje sustiže ono što komentar u vba_check.py:4104-4114 već tvrdi. Usput: čitanje payload-a prebaciti u sam vba_check (--hook-stdin), da se po Edit-u ne dižu tri Python procesa.
  DOBITAK: Zatvara izmerenu rupu (hook nad tools/sabotaza.py: exit 0 za 63 ms, bez provere). Pogrešno sidro u novom unosu kataloga vidi se u istom koraku, a ne pri sledećoj VBA izmeni ili u CI. Cena po izmeni kataloga posle F1/F2: oko 1,3 s (danas bi bila oko 4,7 s). Ušteda dva procesa po Edit-u: nije mereno na Windows-u.
  RIZIK: Nizak; hook nad sabotaza.py postaje blokirajući, pa unos sa namerno 'poznatim' nalazom mora u POZNATI_NALAZI u istom Edit-u.
  PREDUSLOVI: F1 (da cena bude oko 1 s). Izmena je u .claude/ -> zaseban process PR, jedan po jedan, ne uz feature izmenu.
  DOKAZATI: Test hook skripta sa lažnim payload-om (kao u ovoj analizi): putanja tools/sabotaza.py uz pokvareno sidro u kopiji -> exit 2 i 'KATALOG:' u stderr; zdrav katalog -> exit 0; putanja README -> exit 0 bez pokretanja provere.

### F8 Dvosmerni dokaz za najskuplje delove FAST sloja da bude automatski [trud S] [slabi: NE -- jača.]
(a) Dodati 'python3 tools/sabotaza.py --self-test' u .github/workflows/static.yml (0,65 s; danas ga nema ni u CI ni u vba_check). Napomena: self-test privremeno piše modOtkupUI.bas.selftest.tmp u src-vba -- prebaciti na tempfile da ne dira radno stablo. (b) Dodati self-test slučajeve za tokenizere: apostrof unutar string literala praćen pravim komentarom (za _strip_comment preko check_undefined/ARNOST) i prelomljena izjava (za _izjava_deo) -- izmereno je da kvar u _strip_comment ostavlja 152 slučaja zelenim. (c) Deterministički čuvar složenosti umesto merenja sekundi: brojač obrađenih bajtova u proveri sidara mora biti <= k x zbir veličina jedinstvenih fajlova (danas 37,5x) -- pada ako se neko vrati na 'parsiraj ceo fajl po sabotaži'.
  DOBITAK: Nije brzina. Sprečava da se ubrzanja (F1, F2) uvedu bez provere koja ume da pocrveni, i da se cena ponovo tiho popne kao sa 'traje sekundu' na 3,5 s.
  RIZIK: Nizak. (c) mora biti brojač rada, ne štoperica, da ne bude nestabilan u CI.
  PREDUSLOVI: Nema; ide PRE ili zajedno sa F1/F2.
  DOKAZATI: (a) u CI zeleno, a sa namernim kvarom _dodela_tudjoj_proceduri crveno (pokazano lokalno: 5/43, 6/43, 2/43). (b) novi slučajevi padaju nad pokvarenim _strip_comment (danas rc=0) i prolaze nad zdravim. (c) vratiti staru funkciju -> čuvar pada.

### F9 Keš činjenica po fajlu (ključ: ime + sha1 sadržaja + otisak alata) [trud M] [slabi: NE, pod uslovom da CI ostane bez keša i da je ključ sadržaj + otisak alata. Bez tih uslova: DA -- 'zeleno' bi moglo da znači 'zeleno pre poslednje izmene'.]
Globalne provere čitaju iz svakog fajla samo 'činjenice': definisana imena, članove modula, sirovu arnost, deklaracije procedura (popis), pozive BindSourceDocument, i presudu kataloga za dati skup otisaka. Keširati ih po (ime, sha1) i računati ponovo samo za izmenjene fajlove; spajanje ide istim redosledom kao danas. Otisak pet alata (vba_check, sabotaza, vba_gate, run_vba, dokaz) je deo ključa. Keš van praćenog stabla (temp ili ignorisan fajl uz tests/last_*), JSON a ne pickle, atomski upis, prekidač --bez-kesa; CI uvek bez keša. U repou već postoji srodna infrastruktura (vba_gate.otisak_izvora, sha256) -- proširiti nju, ne praviti drugu. Prototip: .../fast-sloj/patch_p4.py (pickle, samo za merenje).
  DOBITAK: Izmereno prototipom: hook 0,18-0,24 s kad izmenjen fajl nije u katalogu (6% izmena), 1,1-1,7 s za fajl iz kataloga (64%), 3,1-3,8 s za modTest/BFP (30%); ponderisano oko 2,0 s prema 3,9 s posle F1-F4. Pošteno rečeno: nad F1-F4 keš donosi još oko 1-2 s po Edit-u, jer posle njega ostaju katalog (0,8-1,06 s) i provere samog fajla. Glavna vrednost keša je što omogućava F10.
  RIZIK: Srednji: ustajali keš je klasa greške koja daje lažno zeleno. Ublažavanje: ključ je sadržaj (ne mtime), otisak alata u ključu, CI bez keša, i self-test koji dokazuje nevažeći keš (izmena fajla -> nalaz se pojavi; vraćanje -> nestane). Dve sesije nad istim stablom: atomski upis (os.replace). Dodatni kod za održavanje: oko 120-150 linija.
  PREDUSLOVI: F1-F4 (bez njih bi keš maskirao algoritamske greške umesto da ih ukloni); odluka gde keš živi (.gitignore red ako je u repou).
  DOKAZATI: Pokazano na prototipu: 16 scenarija izmene sa identičnim izlazom (uključujući dirnuto sidro -> KATALOG, promenu potpisa u drugom fajlu -> 7 ARNOST, brisanje fajla, vraćanje izmene) i prljav korpus sa hladnim i toplim kešom (438 redova stderr identično). U proizvodu: self-test 'keš ne menja izlaz' (isti korpus hladno/toplo) + dvosmerno: pokvariti ključ (izostaviti sha1) -> scenario 'vraćena izmena' mora pasti.

### F10 Hook sa semantikom punog prolaza: lomovi preko granice fajla vidljivi odmah [trud L] [slabi: NE -- jača: hook počinje da tvrdi ono što danas tvrdi samo pun prolaz. (Ako se nalazi van fajla spuste na upozorenje, tvrdnja hook-a ostaje ista kao danas, ne slabija.)]
Nad kešom iz F9 keširati i NALAZE po fajlu, sa ključem (sha1 fajla + sha1 globalnih tabela od kojih zavise: definicije, arnost, članovi modula, ključevi modPoruke, članovi formi, registri). Hook tada izveštava o celom repou: ponovo se proveravaju izmenjeni fajl i samo oni fajlovi čiji se ulaz promenio (npr. promena potpisa -> NEDEFINISAN/ARNOST nad .bas fajlovima), a DUPLIKAT se računa iz keširanih javnih imena. Nalaze van izmenjenog fajla prijaviti odvojeno ('tvoja izmena je slomila N poziva u M fajlova').
  DOBITAK: Zatvara izmerenu rupu: promena potpisa LogErr -> hook danas rc=0, pun prolaz 492 ARNOST u 78 fajlova; preimenovanje -> 552 NEDEFINISAN u 82 fajla; DUPLIKAT se u hook-u danas ne gleda. Cena: tipičan Edit ostaje kao u F9; Edit koji menja potpis ili briše proceduru plaća ponovni check_undefined nad svim .bas (izmereno 1,03 s posle F2) -- procena 2-3 s. Alternativa bez keša je pun prolaz u hook-u: 13,4-13,9 s posle F1-F3.
  RIZIK: Srednji/viši: tačna lista zavisnosti po proveri mora biti potpuna (propuštena zavisnost = ustajao nalaz). Hook postaje stroži, pa će tokom višekoračnog refaktora (prvo potpis, pa pozivaoci) blokirati međukorake -- treba odlučiti da li su nalazi van izmenjenog fajla blokirajući ili upozorenje.
  PREDUSLOVI: F9; odluka vlasnika o blokiranju međukoraka; tabela zavisnosti iz ove analize kao polazna specifikacija.
  DOKAZATI: Ekvivalencija sa punim prolazom bez keša nad prljavim korpusom posle niza od N nasumičnih izmena (keš inkrementalno vs. hladan pun prolaz -- izlaz mora biti identičan posle svake izmene); dvosmerno: izostaviti jednu zavisnost iz ključa (npr. arnost) -> scenario 'promena potpisa' mora pasti po imenu.

### F11 Drugi talas egzaktnih rezova bez keša: jedan prolaz za simbole i provere velikih modula [trud M] [slabi: NE]
(a) collect_definitions + collect_module_members + collect_arities spojiti u JEDAN prolaz po fajlu (danas tri čitanja i 2-3 regex-a po liniji u svakom); (b) popis suita: preskočiti fajl koji ne sadrži nijedno ime iz SUITES/SUITE_VAN_KAPIJA niti oblik Run*Suite/Run*Tests/Test*_All (deklaracije drugih imena ne utiču na nalaz); (c) check_nedeklarisan._bez_teksta (petlja znak po znak nad celim fajlom, 2,06 s u punom prolazu) i _registar_blokovi (re.match sa nekompajliranim izrazima po liniji) prevesti na iste brze putanje kao u F2; (d) logicke_izjave zapamtiti po tekstu u okviru procesa (popis, AMB_BIND, rollback_red i collect_public parsiraju isto). Napomena iz merenja: neusidren case-insensitive predfilter po liniji je SPORIJI od usidrenog match-a (collect_module_members 0,44 s -> 0,64 s) -- taj oblik ne koristiti.
  DOBITAK: Procena (nije prototipirano kao celina): hook 2,5-2,9 s -> 1,5-2,0 s za tipičan fajl bez ikakvog keša; modTest hook 4,3 s -> oko 3 s; pun prolaz 13,4 s -> oko 9-10 s. Izmereno pojedinačno: memo logicke_izjave spušta AMB_BIND 0,32 s -> 0,20 s.
  RIZIK: Nizak do srednji: (a) i (b) menjaju strukturu koda, ne samo brze putanje; (d) drži oko 40-50 MB u memoriji tokom punog prolaza (procena).
  PREDUSLOVI: F2 i F8b. Ako se ide na F9, (a) i (b) postaju suvišni -- birati jedno od dva.
  DOKAZATI: Isti diferencijalni okvir kao za F2 (diff po funkciji + prljav korpus + self-test + vba_gate --self-test); za (b) dodatno self-test slučaj 'suite po konvenciji u fajlu koji se inače preskače'.


## NEPOZNANICE
[
 "Vremena na korisnikovoj Windows mašini (Git Bash + Windows Python, Defender, sporije dizanje procesa) NISU merena -- svi brojevi su sa ovog Linux-a (4 jezgra, Python 3.13.18). Odnosi pre/posle bi trebalo da važe jer je posao čist CPU u Pythonu, ali apsolutne sekunde i režija hook skripta (tri Python procesa po Edit-u) mogu biti veće.",
 "Stvarni broj Edit/Write poziva po sesiji nije poznat -- ne postoji log hook-a. Dnevna cena je procena '1 Edit po hunk-u u commit diff-u': potcenjuje (ispravke istog mesta, odbačene izmene, Edit-i između commit-a) i precenjuje (Write celog fajla ili jedan Edit koji pokriva više hunk-ova). Jedan red u hook skriptu koji beleži vreme i trajanje dao bi pravu brojku za nedelju dana.",
 "Šum merenja je oko +-10% (load average 0,8-0,9 tokom rada, verovatno drugi agenti na istoj mašini): hook originala 9,1-10,4 s, prototipa 2,9-3,8 s. Brojevi su dati kao rasponi; pojedinačne razlike ispod 0,3 s ne treba tumačiti.",
 "Prljavi korpus (dva semena) pokriva 15 kodova nalaza; kodovi NEDEKLARISAN, PROCEDURA_VELIKA, ODSECEN, STORNO_PROGUTAN, ROLLBACK_TVRDNJA, ROLLBACK_TVRDNJA_RED, SEMA_REGISTAR i KRAJ_REDA u njemu nisu okinuti -- za njih je ekvivalencija prototipa pokazana samo kroz self-test (152 slučaja, identičan izlaz) i čist repo.",
 "Prototip keša (F9) koristi pickle i kešira samo činjenice za globalne provere, ne i nalaze po fajlu; F10 (nalazi po fajlu + tabela zavisnosti) NIJE prototipiran, pa su njegovi brojevi procena.",
 "Nije provereno da li je __pycache__ upisiv na korisnikovoj mašini i u web sesiji; bez njega se sabotaza.py (717 KB) kompajlira pri svakom hook-u -- ovde izmereno samo 0,04 s, pa je uticaj mali.",
 "Nije ispitano koliko Claude Code čeka na hook kad paralelno radi više sesija nad istim stablom, niti da li postoji timeout hook-a koji bi 13,7 s nad BFP modulom mogao da preseče na sporijoj mašini.",
 "Predlog F10 otvara pitanje koje traži odluku vlasnika, ne merenje: da li nalaz van izmenjenog fajla (lom pozivalaca tokom višekoračnog refaktora) treba da BLOKIRA sledeći Edit ili samo da se prijavi.",
 "Grana B (izbor testova) nije izvršavana; provereno je samo da su funkcije koje prototip dira identične u A i B. Da li njen commit u tools/sabotaza.py menja cenu provere sidara nije mereno."
]
