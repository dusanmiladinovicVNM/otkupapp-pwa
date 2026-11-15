# BROJEVI I OZNAKE IZVORA

Nema P0. Ima 9 nalaza P1 i 9 nalaza P2. Nijedan ne menja smer plana, ali nekoliko brojeva u rezimeu i ciljnoj slici ne stoji.

Oznake u dokazima: PLAN = /home/user/otkupapp-pwa/docs/engineering/PLAN_OPTIMIZACIJE_VERIFIKACIJE.md; "0X:red" = fajl u /home/user/otkupapp-pwa/docs/engineering/analize/2026-11-verifikacija/; A = stablo main-wt @1fa4937.

Najteže stvari:
1. Trajanje punog kataloga (44–82 h, 11–20 h) je potcenjeno prema ulazima koje sam plan navodi u §2.2. Sa RunAllTests 215,8 s isti model daje oko 100 h pojedinačno i 17–26 h grupno. Analize se razilaze (40–100 h, 7–20 h), a plan to prećutkuje.
2. Budžet "jedna noć" za nedeljni grupni prolaz celog kataloga nema pokriće, a taj prolaz je jedna od tri mere kojima plan zatvara slabljenje izbora.
3. Red "Pun pojedinačni dokaz" u tabeli ishoda računa nad zamrznutih 1265 sabotaža, prećutkuje promenu režima u "izbor" i ne slaže se sa §8.1.
4. Polazne vrednosti FAST sloja (9–14 s, 30–34 s, 17 s, 27 min dnevno) su merenje pod opterećenjem. Dve druge analize i moje ponovljeno merenje na neopterećenoj mašini daju oko 1,45 puta manje. Odnos pre/posle iz prototipa je potvrđen (oko 2,3–2,9 puta, izlaz identičan).
5. Prvi Windows prolaz "1–1,5 h" je nemoguć: jedna komanda iz Priloga A.2 sama traje 2,1–2,3 h po brojevima plana.
6. Rokovi K-2 i K-4 se ne slažu sa talasima u koje su stavljeni.

Radio sam nad kopijama u <scratch>/provera-plana/; stabla A i B nisu menjana. Prototip keša (patch_p4.py:221) piše pickle u sistemski /tmp; taj fajl sam obrisao. Excel nije korišćen, pa nijedno Excel trajanje nisam mogao da proverim merenjem.

## NALAZI
### 1. [P1] §1 r.36 i r.61; §2.2 r.98; §9 O-1 r.416
TVRDNJA: Pun katalog: 44–82 h pojedinačno, 11–20 h u grupama po 6 [I]
PROBLEM: Gornje granice su računate sa RunAllTests = 120 s (03:72), a sam plan u §2.2 navodi poslednji zapis 215,8 s (05.11) i BFP 350–390 s. Sa ulazima iz §2.2 isti model daje oko 100 h pojedinačno i 17 h grupno (25 h sa ponavljanjima). Uz to šest analiza daje različite vrednosti za istu veličinu, a plan uzima jednu i razilaženje ne pominje.
DOKAZ: 03:72 '43,6 h (T_A=80, T_B=150); 81,7 h (T_A=120, T_B=380)'; 03:78 '10,8 h do 20,4 h; od toga ponavljanja 3,4 do 6,5 h'; PLAN r.92–93. Računica formulom iz 03 (801·(18+T_A+1,1)+456·(18+T_B+1,1)+banka; grupe 134·(18+T_A)+78·(18+T_B)+3·48 + 8% ponavljanja): za ulaze iz 03 reprodukuje 43,6 / 81,6 / 10,8 / 20,3 h; za T_A=215,8 i T_B=350..390 daje 99,1–104,2 h pojedinačno, 16,7–17,6 h grupno bez ponavljanja, 24,6–25,9 h sa ponavljanjima. Ostale analize: 07:33 '~17 h; bez grupisanja ~100 h'; 06:6 '60-95 h pojedinacno i 11-16 h grupno'; 05:63 '40-74 h', '7-8 h'; 08:72 i 08:78 '~43 h', '~7,3 h'; 04:98 '13-15 h'.
ISPRAVKA: §2.2 r.98: 'Trajanje punog kataloga | pojedinačno od 44 h (zastarela dokumentovana trajanja 80/150 s) do oko 100 h (poslednji zapisi: RunAllTests 215,8 s, BFP 350–390 s); grupno od 7–11 h (zastarela trajanja) do 17 h bez ponavljanja, oko 25 h sa 8% ponavljanja | [I]; analize daju 40–100 h odnosno 7–20 h zavisno od toga koja trajanja uzimaju, ništa nije mereno'. U §1 r.36, r.61 i O-1 zameniti '44–82 h' sa '44–100 h' i '11–20 h' sa '11–25 h'.

### 2. [P1] §6 r.224 (sloj Nedeljno) i r.225 (Release); §1 r.62; §7 r.306–307; O-1
TVRDNJA: Nedeljno: grupni prolaz cele suite nad celim katalogom, budžet 'jedna noć'. Starost najstarijeg dokaza posle talasa 5: 'do 7 dana'.
PROBLEM: Budžet 'jedna noć' nema pokriće ni u podlozi ni u brojevima samog plana. Plan isti prolaz ceni na 11–20 h danas, a sa ulazima iz §2.2 to je 17–26 h. Po projekciji iz 03 raste kvadratno: 22–41 h za mesec dana, 54–105 h za tri meseca; plan u §3.1 sam navodi 129–256 h za šest meseci. Jedini izvor za 'noć' je 05 sa zastarelim trajanjima. Na ovaj prolaz se oslanja zatvaranje slabljenja izbora i varijanta V2 u O-1. 'Do 7 dana' nema oznaku i važi samo za rotaciju u izboru.
DOKAZ: PLAN r.224, r.98, r.123–124, r.460 (veći --grupe odbijen kao trajno rešenje). 03:213 (D9: nedeljni prolaz bez budžeta); 03:95 'Grupe 6: 22–41 h / 54–105 h / 129–256 h'; 06:253 'Ceo katalog obidje se za oko 5-8 noci (215 grupnih prolaza, oko 11-16 h)'; 06:280 'oko 11-16 h, jedna noc ili vikend'; 05:253 'procena 7-8 h ... svake noci ili vikendom' (trajanja iz 05:63: BFP ~150 s). Računica za 'do 7 dana': noćni budžet 3 h minus FULL 20–60 min [D] ostavlja 2–2,7 h; grupno 17–26 h je 6–13 noći; izbor posle R-3 (5,1–6,5 h) je 2–3 noći.
ISPRAVKA: §6 Nedeljno, kolona Budžet: 'danas 17–26 h [I] (vikend); jedna noć tek posle R-4 ili K-4; raste sa katalogom i brojem testova, pa stavka N-1 nosi prag: iznad oko 36 h prelazi se na grupni prolaz po oblasti (K-4)'. §6 Release, Budžet: 'jedna noć tek posle talasa 4–5; pod V1 danas 17–26 h plus FULL'. §1 r.62: 'do 7 dana [I], uz rotaciju u izboru posle R-3 (2–3 noći); rotacija --grupe 6 nad celom suitom traje 6–13 noći'.

### 3. [P1] §1 tabela 'Očekivani ishod' r.61; §8.1 r.368
TVRDNJA: Pun pojedinačni dokaz: 44–82 h → posle talasa 3: 10–14 h [I] → posle talasa 5: do 5 h na jednoj mašini, do 2 h sa 3–4 radnika [I]. §8.1 talas 4: 'pun pojedinačni dokaz reda 10 h'.
PROBLEM: Četiri stvari. (1) Vrednosti su računate za današnjih 1265 sabotaža, a plan sam tvrdi +18,7 dnevno; podloga za isti režim daje 16,5 h za mesec i 26,7 h za tri meseca. (2) Posle talasa 3 to više nije isti verdikt nego 'DOKAZANO (izbor)', što red ne kaže, protivno N1. (3) 'Do 5 h' je broj iz 03 za režiju 5 s, a R-3 po planu daje 6–8 s; sa 6–8 s izlazi 5,1–6,5 h, a sa 2–4 puta ubrzanja 1,3–3,3 h. (4) §8.1 za talas 4 kaže 'reda 10 h', što je isto kao talas 3, iako R-1 i R-3 ulaze u talas 4.
DOKAZ: 03:84 '10,5 h (tau = 2 s), 11,5 h (4 s), 13,5 h (8 s); 1811 run_vba poziva ... Testovi koji sami ne prolaze dodaju po jedan prolaz cele suite (udeo nepoznat)'; 03:95 'Izbor pojedinačno (tau = 4 s): 16,5 / 26,7 / 41,8 h', 'uz režiju 5 s: 7,1 / 11,4 / 18,0 h'; 03:193 'Uz režiju 5 s pun pojedinačni dokaz je oko 4,9 h'; 02:188 'režija prolaza ~18 -> ~6-8 s'; 08:156 'realno 2-4x'. Računica 1257·(r+4+p)+545·(r+4)+432 s: r=5 daje 4,6–5,0 h; r=6 daje 5,1–5,5 h; r=8 daje 6,1–6,5 h.
ISPRAVKA: Red preimenovati u 'Pun pojedinačni dokaz (od talasa 3: u izboru, verdikt DOKAZANO (izbor))'. Posle talasa 3: '10,5–13,5 h [I] pri 1265 sabotaža, bez testova koji sami ne prolaze [?]; uz +18,7 dnevno 16,5 h za mesec, 26,7 h za tri meseca'. Posle talasa 5: '5–6,5 h na jednoj mašini (režija 6–8 s), 1,5–3 h sa 3–4 radnika [I], pri 1265; uz rast 7–11 h na jednoj mašini za 1–3 meseca'. §8.1 talas 4: 'pun pojedinačni dokaz u izboru reda 5–6,5 h pri 1265'.

### 4. [P1] §1 r.35, r.57, r.58; §2.1 r.77–81 i r.85; §6 r.219
TVRDNJA: Hook 9–14 s [M]; pun vba_check 30–34 s [M]; --self-test 17,4–17,8 s [M]; dnevno čekanje oko 27 min, p90 oko 90 min; posle: 3–5 s, 13–14 s, do 2 s [M, prototip].
PROBLEM: Plan je za 'danas' uzeo najsporije merenje (analiza 01, pod opterećenjem). Dve druge analize i moje ponovljeno merenje na neopterećenoj mašini daju oko 1,45 puta manje. Plan to pominje samo u jednom redu (mali hook), a ne za pun prolaz, self-test, dnevno čekanje ni rezime. Odnos pre/posle iz prototipa jeste potvrđen, ali apsolutne vrednosti zavise od opterećenja i nijedna nije merena na Windows-u. Budžeti u §6 ('do 4 s, zatim do 2 s') nisu ispunjeni po brojevima koje plan sam navodi za modTest/BFP (30% izmena): 4,3–4,9 s posle F-2 i 3,1–3,8 s sa kešom.
DOKAZ: Moje merenje nad kopijom alata iz A, load 0,02–0,21: hook modMeteo.bas 6,11 / 6,10 / 6,42 s; modTest.bas 8,38 s; BFP 8,55 s; pun vba_check 21,53 i 22,53 s; --self-test 12,51 s. Sa patch_p1 + patch_p2 + patch_p2b: 2,40 / 2,37 / 2,29 s; 3,49 s; 3,24 s; 9,89 s; 4,26 s; cmp punog izlaza: identičan. Sa patch_p4: 0,13 s (van kataloga), 0,90–1,12 s (modAmbalaza, modDokumenta), 1,95–1,97 s (modTest, BFP). Analize: 06:74 'hook 6,3 s po VBA fajlu (ranije 9,0-11,9 s)'; 07:39 '21.58 / 6.19'; 07:40 'Apsolutne vrednosti su nize od ranije izmerenih 31.0 / 9-11.9 / 3.5 s (drugo opterecenje masine)'; 05:197 'Self-test lokalno 12,3 s'; 01:155 i 01:197 (4,3–4,9 s; 3,1–3,8 s; 30% izmena). Računica: 143·(0,7·6,2+0,3·8,5) s = 16,4 min; p90 480 hunk-ova = 55 min.
ISPRAVKA: §1 r.35: 'Hook posle svake VBA izmene | 6–9 s na neopterećenoj mašini, 9–14 s pod opterećenjem [M]'. r.57: '6–9 s (9–14 s) → 2,3–3,5 s (3–5 s) → 0,1–2,0 s (0,2–3,8 s) [M, prototip]'. r.58: '21,5–22,5 s (30–34 s) → oko 10 s (13–14 s)'. r.81: '12,3–12,5 s (17,4–17,8 s pod opterećenjem)'. r.85: '143 × 6,9–11,2 s = 16–27 min; p90 dan 55–90 min'. Ispod tabele dodati: 'Vrednosti u zagradi su merene pod opterećenjem paralelnih sesija. Odnos pre/posle je isti u oba merenja (oko 2,3–2,9 puta); na Windows-u nije mereno.' U §6 uz budžet Edit sloja dopisati 'na neopterećenoj mašini'.

### 5. [P1] §1 'Problem u četiri broja', r.37
TVRDNJA: Rast: za 30 dana sabotaže +76%, testovi +67%, produkcioni VBA +9% [M]; trajanje RunAllTests 0,10 → 0,64 s po testu [D]
PROBLEM: Vrednost 0,10 s po testu je od 27.08 (123 testa), a ne iz prozora od 30 dana u čijem redu stoji. U tom prozoru (09.10 → 05.11) promena je 0,36 → 0,64 s. Red tako prikazuje 6,4 puta umesto 1,8 puta. Podloga uz to kaže da tačke potiču iz proze i da uslovi mašine nisu poznati.
DOKAZ: 07:27 '27.08: 12.0 s / 123 / 0.098; ... 09.10: 73.5 s / 204 / 0.360; ... 05.11: 215.8 s / 336 / 0.642'; 04:9 '0,36 s/test -> 0,64 s/test'; 07:227 'nije poznato da li su merene na istoj masini i pod istim uslovima'. Od 27.08 do 05.11 je 70 dana.
ISPRAVKA: 'trajanje RunAllTests 0,36 → 0,64 s po testu za 27 dana (09.10 → 05.11); od 27.08 0,10 → 0,64 s [D, četiri tačke iz proze, uslovi mašine nepoznati]'.

### 6. [P1] §7 M-7 (Trud '1–1,5 h') r.249; §8.4 korak 5 r.405; Prilog A.2 r.539
TVRDNJA: Prvi merni prolaz na Windows-u: 1–1,5 h, uključujući ručni Compile, FULL i protokol iz Priloga A.
PROBLEM: Komanda 'python tools/dokaz.py amb-pisac' iz A.2 bira 19 sabotaža, sve u BFP-u, i bez --grupe i --izbor pušta 19 prolaza cele BFP suite plus dve baze. Po trajanjima iz §2.2 samo ona traje 2,1–2,3 h. Sa FULL-om (20–60 min [D]) i ostatkom protokola to je 3–4 h.
DOKAZ: Komanda nad kopijom alata iz B: 'python3 tools/dokaz.py amb-pisac --plan' → 'sabotaza 19, prolaza 19 (najvise 1 po grupi)'; lista: 8 × Test_Amb_PisacKnjige, 6 × Test_Amb_DeficitIObaveza, 3 × Test_Amb_ZaglavljeDokumentaPisac, 2 × Test_Amb_JedanProtivpartnerPoDokumentu. Računica: 19·(18+350..390) + baze (234 + 368..408 s) = 127–140 min. Ostatak: A.1 oko 12 min, zajednički poziv dve suite oko 10 min, amb-pisac --izbor (23 poziva po oko 25 s) oko 10 min, FULL 20–60 min (PLAN r.94).
ISPRAVKA: M-7 Trud i §8.4 korak 5: 'oko 3–4 h mašinskog vremena, od toga poređenje bez --izbor oko 2,2 h'. Ili u A.2 zameniti 'amb-pisac' užim prefiksom sa 3–5 sabotaža, pa prolaz staje u oko 1,5–2 h zajedno sa FULL-om.

### 7. [P1] §7 R-3 r.316 (i R-1 r.314)
TVRDNJA: R-3: režija 18 → 6–8 s [I]
PROBLEM: Podloga pripisuje 18 → 6–8 s zbiru RV-02, RV-03 i RV-06, to jest stavkama R-1 i R-3 zajedno. Sam R-3 je procenjen na −3 do −7 s. Plan isti dobitak broji dvaput: jednom kao R-1 (1,5–4 s i 2,5–5,9 s), pa ceo kumulativ kao R-3. To iskrivljuje poređenje R-1 i R-3 i ulaz za kapiju G2.
DOKAZ: 02:188 'Uvoz 192 komponente -> 1-6 modula po prolazu: procena -3 do -7 s ... Sa RV-02/03 režija prolaza ~18 -> ~6-8 s'; 02:160 i 02:167 (RV-02 i RV-03, što plan vodi kao R-1).
ISPRAVKA: R-3, Dobitak: '−3 do −7 s po prolazu [I]; zajedno sa R-1 režija oko 18 → 6–8 s'. U §8.2 dodati zavisnost 'R-1 → R-3 (broj 6–8 s važi samo za obe)'.

### 8. [P1] §4 nalaz 12 r.172; §7 D-3 r.295; §12
TVRDNJA: D-3 (kontrolne placebo sabotaže) 'razrešava nalaz 12 iz §4'.
PROBLEM: Nalaz 12 je da sedam sabotaža sa priznatim nalazom 'tvrdnja nije CEO STATICAN literal' nije u spisku koji dokaz.py jedini čita. Placebo kontrola iz D12 se toga ne tiče: ona tvrdi da prolaz koji ništa ne obara ne bude priznat. U podlozi je ovo otvoreno pitanje, a ne predlog koji D12 zatvara, i nema ga ni u §12.
DOKAZ: 03:17 'POZNATI_NALAZI ima 9 upisa (7x tvrdnja nije CEO STATICAN literal u BFP ...); POZNATI_NALAZI_DOKAZ ima 2 ... dokaz.py čita samo drugi spisak'; 03:264 'Da li pun katalog danas uopšte može da vrati DOKAZANO ... Izvedeno iz koda, nije pušteno'; 03:234 (tekst D12: placebo i dve zatečene kontrole). Kod u A: tools/dokaz.py:453 'poznati_spisak = getattr(sab, "POZNATI_NALAZI_DOKAZ", {})'; tools/sabotaza.py:12586 i 12651.
ISPRAVKA: U D-3 obrisati 'razrešava nalaz 12 iz §4'. U §4 red 12, kolona Stavka: 'M-7, zatim upis u POZNATI_NALAZI_DOKAZ ili ispravka tvrdnji'. U §12 dodati red: 'Da li sedam sabotaža sa priznatim statičkim nalazom obara svoju tvrdnju u dokazu | Prilog A.2: dokaz.py nad tih sedam imena'.

### 9. [P1] §7 'Rokovi koje nameće rast' r.343–344; §8.1 r.368–370; O-5 r.420
TVRDNJA: K-2 pre kraja decembra 2026, K-4 pre sredine januara 2027; K-2 je u talasu 4, K-4 u talasu 6, K-1 u talasu 5.
PROBLEM: Datumi su tačno izračunati, ali se ne slažu sa talasima. Zbir truda iz kolona 'Trud' do kraja talasa 4 je 30–66 radnih dana, a do K-4 51–116. K-2 stiže na rok samo po donjoj granici truda, K-4 ne stiže ni po donjoj. Uz to L-3 (talas 1) pravi vba_check crvenim na 60.000 linija, što je za modTest oko 11.01.2027, pre 'sredine januara'. O-5 kaže 'K-1 odmah posle M-1', a §8.1 ga stavlja u talas 5.
DOKAZ: Računica roka od 08.11.2026: TestName (40.000−22.178)/63/4,7 = 60,2 dana → 07.01.2027; modTest (65.534−31.061)/452 = 76,3 dana → 23.01.2027; prag 60.000: 64 dana → 11.01.2027; prag 45.000: 31 dan → 09.12.2026. Trud po legendi (PLAN r.16–17), jedan izvršilac, redom po talasima, bez čekanja na Windows: donja granica (S=0,5 d, M=2 d, L=5 d) do kraja talasa 4 je 30 radnih dana (oko 25.12.2026), do K-4 51 (oko 25.01.2027); sredina (M=3, L=7,5) 40 (oko 08.01.2027) i 74 (oko 25.02.2027); gornja (S=1, M=4, L=10) 66 (oko 15.02.2027) i 116 (kraj aprila 2027). Napomena: u 7 dana posle analize rast je bio mnogo manji od modela (origin/main ed506dc7: modTest 31.459 linija, to jest +57 dnevno; BFP bez promene; RunOne 350 → 353), pa su rokovi osetljivi na stopu.
ISPRAVKA: U §8.1 K-2 premestiti u talas 2–3 (nema preduslova), a K-4 odmah posle S-4 i K-2, ispred talasa 5. U r.343: 'K-2 mora početi do oko 10.12.2026 (TestName prelazi 40.000 bajtova oko 07.01.2027); K-4 mora biti gotov pre oko 11.01.2027, kada modTest prelazi prag greške 60.000 iz L-3 (granica 65.534 oko 23.01.2027); rokovi važe uz +452 linije dnevno, a poslednjih 7 dana rast je bio +57'. Uskladiti O-5 i §8.1 za K-1.

### 10. [P2] §7 R-1 r.314
TVRDNJA: R-1: 1,5–4 s i 2,5–5,9 s po pozivu [I]
PROBLEM: Dve uštede se čitaju kao zbir (do 9,9 s), a preklapaju se na istim sleep-ovima compile probe. Svi fiksni sleep-ovi zajedno su 4,5–7,9 s, što plan sam navodi u §2.2.
DOKAZ: A/tools/run_vba.py:617 sleep(2.5), :632 sleep(0.5), :634 sleep(0.4), :1012 i :1021 sleep(1.0); 02:160 (RV-02: do ~2,4 s + ~0,9 s na compile probi, ~0,5–0,9 s na gašenju); 02:167 (RV-03: preskakanjem compile koraka 2,5–5,9 s). Zajedno najviše 5,9 + 0,9 = 6,8 s.
ISPRAVKA: '3–7 s po mutacionom pozivu [I]: 2,5–5,9 s preskakanjem compile probe i još 0,5–0,9 s na gašenju; ušteda čekanja na compile probi se ne sabira sa preskakanjem'.

### 11. [P2] §7 F-3 r.263 i F-4 r.264
TVRDNJA: F-3: 17 s → 2–3 s. F-4: zid 76 s → 30–35 s [I].
PROBLEM: Oba dobitka su kumulativna, a pripisana jednoj stavci. F-2 već spušta self-test na oko 6 s, pa F-3 donosi 6 → 2–3 s i to je procena bez oznake. Sama podela na job-ove u F-4 daje 50–55 s; 30–35 s važi tek uz F-3 i jedan prolaz who_writes.
DOKAZ: 01:169 'Sa F1-F3 već izmereno 5,8 s; sa pamćenjem u procesu procena oko 3 s'; 05:197 'CI korak 13-18 s -> oko 2-3 s'; 05:211 'Zid 73-86 s -> oko 50-55 s odmah, oko 30-35 s uz CI-1 i CI-11'.
ISPRAVKA: F-3: 'self-test 5,8 s (posle F-2) → oko 2–3 s [I]'. F-4: 'zid 76 s → 50–55 s samom podelom; 30–35 s zajedno sa F-3 i jednim prolazom who_writes [I]'.

### 12. [P2] §7 N-6 r.332
TVRDNJA: po PR-u bi se ponovo merilo oko 2–12% kataloga [M]
PROBLEM: Raspon spaja medijanu i maksimum i čita se kao tipičan opseg.
DOKAZ: 08:19 'po PR-u zastari medijana 1,8% kataloga (prosek 3,2%, max 12,7%)', 40 PR-ova; zbirno 3,1%.
ISPRAVKA: 'po PR-u bi se ponovo merilo u medijani 1,8% kataloga (prosek 3,2%, najviše 12,7%; 40 PR-ova) [M]'.

### 13. [P2] §2.3 r.105 i r.111
TVRDNJA: Sabotaže 718 → 1265 za 30 dana; +18,7 dnevno. Presidrenje kataloga 64 u 29 dana (21 od 50 merge koraka).
PROBLEM: Spojena su dva prozora iz dve analize: 718 → 1265 za 30 dana daje +18,2, a +18,7 pripada paru 724 → 1265 za 29 dana. Broj 64 je u podlozi donja granica; druga analiza za isti period meri više.
DOKAZ: (1265−718)/30 = 18,23; (1265−724)/29 = 18,66. 03:43 '724 -> 1265, to jest +18,7 dnevno ... To je donja granica'; 07:25 'u 30 dana ... 139 POSTOJECIH unosa izmenjeno (70 sidro ...); 25 od 51 PR-a'.
ISPRAVKA: r.105: '718 → 1265 za 30 dana; +18,2 dnevno'. r.111: 'najmanje 64 u 29 dana po first-parent koracima (21 od 50); po merge-u 70 presidrenih od 139 izmenjenih unosa u 30 dana'.

### 14. [P2] §3 tačka 1, r.123–124
TVRDNJA: Projekcija za šest meseci uz današnji rast: 129–256 h u grupama [I]
PROBLEM: Izostavljena je pretpostavka pod kojom broj važi: da trajanje suite raste srazmerno katalogu. Bez nje podloga daje linearan rast.
DOKAZ: 03:94 'i pretpostavku da trajanje suite raste srazmerno katalogu'; 03:97 'ako trajanje suite ne raste, prva dva reda rastu linearno (x1,44 / x2,33 / x3,66)'; 10,8–20,4 h × 3,66 = 40–75 h.
ISPRAVKA: 'Projekcija za šest meseci uz +18,7 sabotaža dnevno: 129–256 h u grupama ako trajanje suite raste srazmerno katalogu, 40–75 h ako ne raste [I]'.

### 15. [P2] §1 tabela ishoda r.60
TVRDNJA: Dokaz jednog reza (4–7 sabotaža): oko 34 min [D] → 3–5 min [I]
PROBLEM: Raspona '4–7 sabotaža' nema u podlozi. Zapis od 34 min je za 4 BFP sabotaže uz --grupe 6, a model 3–4 min je za taj isti rez. Druge analize kao primer reza navode 13 i oko 20 sabotaža, za koje model daje više.
DOKAZ: 03:100–101 'Tipičan dokaz reza danas (4 BFP sabotaže, --grupe 6: 2 baze + 4 prolaza) ... U izboru bi isti rez bio 8 poziva po oko 22–26 s, to jest 3–4 min (model)'; 07:98 '13 sabotaza u 7 prolaza'; 08:184 'kod tipicnog reza od ~20 sabotaza'. Računica za 13–20 sabotaža: do 26–40 poziva × 22–26 s = 10–17 min.
ISPRAVKA: 'Dokaz jednog reza (primer: 4 BFP sabotaže, --grupe 6) | 34 min [D] | isto | 3–4 min [I]; rez od 13–20 sabotaža oko 10–17 min [I]'.

### 16. [P2] §3 tačka 2, r.125
TVRDNJA: uvozi 192 komponente (10,4 MB) [M]
PROBLEM: Broj se ne slaže sa stablom ni između analiza: u A/src-vba je 193 komponente, a dve analize navode 194 fajla (sa .frx).
DOKAZ: ls A/src-vba po nastavku: 139 .bas, 10 .cls, 43 .doccls, 1 .frm, 1 .frx; A/tools/run_vba.py:33 ('.bas/.cls/.frm se uklone pa uvezu'). 02:4 '192 komponente (43 .doccls merge + 149 Import)'; 03:29 'svih 194 fajla'; 08:9 '194 fajla'.
ISPRAVKA: 'uvozi 193 komponente (43 .doccls spajanjem, 150 uvozom: 139 .bas, 10 .cls, 1 .frm; 10,4 MB)'.

### 17. [P2] §7 napomena uz M-1, r.252
TVRDNJA: Python strana se primenjuje bez sudara.
PROBLEM: Jedan od četiri Python fajla grane ima konfliktni blok prema tekućem main-u. Plan ga posredno rešava ('četiri sabotaže idu na kraj kataloga'), ali ga vodi pod VBA stranom.
DOKAZ: git merge-file -p nad izvezenim verzijama (baza d19c2dae, origin/main ed506dc7, 9d6e780) u scratch-u: tools/sabotaza.py 1 konfliktni blok; tools/dokaz.py, run_vba.py, vba_gate.py 0; modTest.bas 3; modBusinessFlowProTests.bas 2. Isto u 03:53.
ISPRAVKA: 'Od Python strane, dokaz.py, run_vba.py i vba_gate.py se primenjuju bez sudara; tools/sabotaza.py ima jedan konfliktni blok, pa četiri sabotaže idu na kraj kataloga.'

### 18. [P2] §7 N-2 r.328
TVRDNJA: oko 150 ručnih docs commit-a mesečno manje
PROBLEM: Podloga tvrdi manje: uklanja se većina od oko 150, ne svih 150.
DOKAZ: 07:204 'Uklanja vecinu od ~150 rucnih docs commit-a mesecno'; 07:135 '~150 commit-a mesecno su cisti zapisi merenja (68), merge-a (29) i compile-a (53)'.
ISPRAVKA: 'većina od oko 150 ručnih docs commit-a mesečno'.

## POTVRDJENO
- Katalog: 1265 sabotaža, 215 prolaza uz --grupe 6 (pustio dokaz.py --plan --grupe 6 nad kopijom A)
- 1264 od 1265 različitih ključeva (03:13) i 54% mehanički obmotljivo (03:59)
- modTest.bas 31.061 i BFP 33.077 linija (wc -l u A); sabotaza.py 735.170 B; STANJE_REFAKTORA.md 844.024 B
- Računica rasta: +76,2% i +67,5%; modTest 76,3 dana, BFP 100,5 dana, TestName 60,2 dana; prag MAX_PROC_BAJTOVA = 40000 (vba_check.py)
- Dnevno čekanje: 143 × 11,2 s = 26,7 min, p90 oko 90 min; oznaka [I] odgovara podlozi (01:102, izmereno=false)
- Rast hook-a 0,32 s → 6,19 s i oko 4,5 ms po sabotaži (07:39)
- N-5: 0 neslaganja u 100 daje najviše 2,951% uz 95% (1 − 0,05^(1/100))
- Fiksni sleep-ovi 4,5–7,9 s (run_vba.py:617, 632, 634, 1012, 1021); režija oko 18 s je [D]
- Brojevi §2.2 sa oznakom [D] postoje u podlozi: 215,8 s / 336; 73,5 s / 204; BFP 210 s i 350–390 s; 34 min za 4 sabotaže (procena 20–25); 7–20 podizanja i 21–88 min; 28.08.2026 260/261; '~37 min' i '20–60 min'
- CI: medijana 76 s na 25 run-ova (05:9), 12 s u avgustu (07:41); lokalno 51,4 s redom i 22,8 s paralelno (05:15)
- Prototip F-2: izlaz punog prolaza bajt-identičan nad čistim repoom, ubrzanje oko 2,3–2,9 puta; brojevi iz plana odgovaraju 01 (2,9–3,8 s; 4,3–4,9 s; 13,4–13,9 s; 5,8 s)
- Prototip keša F-8: 0,2 / 1,1–1,7 / 3,1–3,8 s odgovara 01:197; ponovljeno merenje daje 0,13 / 0,9–1,1 / 1,95 s
- Nalaz 1: hook pokriva samo *.bas|*.cls|*.frm|*.doccls (.claude/hooks/vba-check.sh:54–55); rc=0 za 0,056 s (06:67)
- Nalaz 2: static.yml ne zove sabotaza.py --self-test; 43 slučaja (05:45)
- Nalaz 3: 492 ARNOST u 78 fajlova (01:204)
- Nalaz 5: dokaz.py ne prosleđuje --timeout (03:39; dokaz.py:92)
- Nalaz 6: ključ 'dialogs' iz SUITES se nigde ne čita (run_vba.py)
- Nalaz 11: popis_citalaca.py --check daje rc=1, 'prj_brojzbirne 33 > prag 31 PAO'
- Nalaz 12, osnova: dokaz.py čita samo POZNATI_NALAZI_DOKAZ (dokaz.py:453)
- 12 podrazumevanih suita u SUITES; osam gate-suita bez rezultat-fajla (04:15)
- Žarišta: sabotaza.py 35 od 40, modTest 28, BFP 25 PR-ova (08:23); registri 1.106 linija (04:61)
- --grupe, --require-green i 'jeftine/skupe kapije' se ne pominju u CLAUDE.md ni u .claude/rules u A (grep: 0)
- main je odmakao 158 commit-a od baze grane (149 na 1fa4937); 1560 je broj ne-merge commit-a od 09.08 (07:19)
- 23 dvosmerne provere: SLUCAJEVI u prenos-izbora/dvosmerno.py ima 23 elementa
- Dobici koji se slažu sa podlogom: M-5 (96–121), S-1 (5–25%), S-2 (1–5 s), S-6 (265–330 s), D-1 (27 s → ispod 1 s; 1,0 → 0,05 s), D-2 (100–400 s → 20–26 s), D-4 (2–9 min), D-5 (2 od 11), R-5 (6–8 → 2–4 s), R-4 (2–4 puta), K-4 (oko 6 puta), F-7 (4–22 komande), prag G2 oko 30%
- Starost dokaza oko 11 nedelja: 28.08 → 15.11 je 79 dana
