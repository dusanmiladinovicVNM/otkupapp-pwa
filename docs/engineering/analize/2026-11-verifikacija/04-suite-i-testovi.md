# VBA test suite (modTest, modBusinessFlowProTests i 13 manjih test modula), katalog SUITES u tools/run_vba.py i fixture (tools/make_fixture.py) -- stablo A (origin/main @1fa4937)

## SAZETAK
Sve putanje su relativne na stablo A (<scratch>/main-wt/) osim gde pise drugacije; skripte kojima su brojevi dobijeni su u <scratch>/suite-analiza/ (an_modtest.py, an_class.py, an_bfp.py, an_suites.py, an_size.py, an_dup.py, an_overlap.py). Nista sto trazi Excel nije izmereno; trajanja su citati iz dokumentacije. (1) Na main-u NE postoji merenje vremena po testu, a dokumentovano trajanje RunAllTests je sa 73,5 s (204 testa) poraslo na 215,8 s (336 testova): prosek 0,64 s/test, a svaki novi test kosta ~1,08 s; pun podrazumevani poziv sa BFP-om vise ne staje u podrazumevani timeout od 600 s. (2) Dva diva nose dva razlicita modela izolacije u istoj svesci: od 350 modTest testova 133 su cisto citanje, 39 samo gradi formu, 43 pise uz snimak+rollback, a do 121 (najmanje 96) pise kroz produkcione pisce i NE vraca stanje; u BFP-u 235 od 392 koraka ostavlja redove. Komentari o redosledu u RunAllTests su zastareli (pozivaju se na brojeve testova koji danas znace nesto drugo), a zavisnosti nigde nisu masinski zapisane. (3) Najverovatniji potrosaci vremena koje suite same izazivaju, a nijedna tvrdnja ih ne meri: CSV journal po svakom AppendRow i log+OnTime po svakom CommitTx (RunAllTests i BFP ne pale tihi rezim, Banka/Agro/SEF pale), 7 upisa u celiju po svakoj od ~3000 BFP tvrdnji, gradnja cele ljuske forme u 51 testu, 5 stvarnih ThisWorkbook.Save u RunAllTests, i produkciona vruca putanja bez kesa (GetTable linearno skenira sve listove pri svakom pozivu). (4) Rast je merljivo na sudaru sa granicama: modTest raste ~452 linije/dan i BFP ~323 linije/dan, pa dokumentovanu granicu od 65.534 linije po modulu stizu za ~76 odnosno ~100 dana, a TestName za ~60 dana prelazi prag PROCEDURA_VELIKA (40.000 bajtova) u vba_check. (5) Tri rucna registra (1.106 linija) i pravilo da katalog sabotaza priznaje samo tri test fajla guraju sve nove testove u modTest/BFP. (6) Deljenje na shard-ove po suite-u je mehanicki blizu (svaki run_vba vec ima svoju temp kopiju i svoj Excel), ali ga blokiraju tri stvari u alatu (fiksan izvestaj tests/last_run.*, compile proba koja trazi fokus VBE prozora, marker zelenog) i jedna u testu (fiksan %TEMP% folder); deljenje UNUTAR RunAllTests/BFP trazi generisan registar i reviziju nezavisnosti. Predlog redosleda: prvo merenje, pa bezrizicna gasenja sporednih efekata, pa generisan registar i podela po oblasti, pa shard-ovi; nijedan od glavnih predloga ne slabi imenovanu tvrdnju, a gde postoji sporedni gubitak pokrivenosti to je izricito navedeno.

## CINJENICE
- [M] RunAllTests (modTest.bas): 350 testova; tri registra su usaglasena (RunOne=350, TestName=350, InvokeTest=350, bez duplikata); 3386 AssertEq/AssertSnapshot u telima testova (~9,7 po testu); 637 procedura od kojih 287 pomocnih; 31.061 linija, 25% su komentari. gate=True, dialogs=False (0 MsgBox van komentara), verdikt iz last_run.txt.
    DOKAZ: python3 suite-analiza/an_modtest.py -> 'InvokeTest: 350 TestName: 350 RunOne: 350 unique 350 ... asserts own total: 3386'; an_size.py -> 'modTest.bas lines 31062 comment lines 7673 =25% procs 637'; tools/run_vba.py:63-64
- [N] Dokumentovana trajanja RunAllTests: 73,5 s pri 204 testa; 210 s (crven prolaz, ~314 testova); 215,8 s pri 336 testova (05.11.2026). Izvedeno: 0,36 s/test -> 0,64 s/test; granicni trosak 132 dodata testa = 142 s, tj. ~1,08 s po novom testu.
    DOKAZ: docs/STANJE_REFAKTORA.md:2115-2116; docs/STANJE_REFAKTORA.md:10525; docs/DOMEN/DOPLATA.md:306
- [N] RunBusinessFlowProSuite: 391 test + seed (392 direktna poziva), 398 Test_ procedura, 764 procedure, 3189 statickih mesta tvrdnji; poslednji dokumentovan broj izvrsenih tvrdnji 3014/0. gate=True kroz Err.Raise u EndRun, zavrsava MsgBox-om (dialogs=True), verdikt iz last_run_bfp.txt. Dokumentovano trajanje: ~150 s (stariji dokument) i 210 s pri 2417 tvrdnji (07.10.2026); novije merenje nije zapisano.
    DOKAZ: python3 suite-analiza/an_bfp.py -> 'procs 764 tests in suite 392 ... asserts (static, own body) sum: 3189'; modBusinessFlowProTests.bas:3092-3125 (EndRun: MsgBox 3114/3117, Err.Raise); docs/DOMEN/DOPLATA.md:308; docs/EXCEL_TEST_HARNESS.md:731-732; docs/STANJE_REFAKTORA.md:995
- [N] Pun podrazumevani poziv run_vba (12 default suita u JEDNOM Excelu) vise ne staje u podrazumevani --timeout 600 s; BFP, agrohemija i golden se pustaju posebnim pozivom sa --timeout 1500 (odnosno 900).
    DOKAZ: docs/DOMEN/DOPLATA.md:308-309 ('ceo poziv sa BFP-om prelazi 600 s'); docs/STANJE_REFAKTORA.md:10525-10526; tools/run_vba.py:705
- [M] Katalog SUITES: 18 suita, 12 default, 15 gate, 12 sa dijalozima, samo 4 pisu rezultat-fajl (RunAllTests, banka, storno, BFP). 8 default gate-suita nema rezultat-fajl (Izvestaj, JSON, Faktura, StornoCentar, License, Palete, Agrohemija, Golden) -- kod njih ime pale tvrdnje ne prezivi COM granicu. Redosled izvrsavanja je redosled recnika: RunAllTests prva, BFP deseta, pa Agrohemija i Golden.
    DOKAZ: python3 (ast.literal_eval nad SUITES iz tools/run_vba.py:60-93) -> '18 suites; default: 12 gate: 15 dialogs: 12 result_file: 4'; tools/run_vba.py:784-792
- [N] Manje suite (broj testova / provera, upis i ciscenje, trajanje ako je zapisano): Banka 25 T / 241 provera, JEDNA spoljna TX sa 9 snimaka + rollback, journal utisan, 2 MsgBox; Storno 23 T / 111, jedna TX sa 17 snimaka + rollback, 2 MsgBox, 2,1 s; Palete 14 T / 115, jedna TX sa 8 snimaka + rollback, do 3 MsgBox, 2,9-3,3 s; StornoCentar 15 / 56, TX+rollback PO TESTU, bez dijaloga; Golden 11 scenarija, TX+rollback po scenariju, 9,2-9,5 s samostalno; Izvestaj 15 T, E2E deo u TX sa 8 snimaka + rollback; Faktura 12 / 39, delimicno TX, MsgBox na kraju; Agrohemija 19 / 25, jedna TX + rollback, journal utisan, 2 MsgBox; JSON 72 provere, bez upisa; License 4 testa, bez upisa.
    DOKAZ: modTestBanka.bas:91-105,129; modTestStorno.bas:62-79,116; modTestPalete.bas:105-113,132; modTestStornoCentar.bas:9-10; modIzvestajTests.bas:78-94; modAgrohemijaTests.bas:61-94; python3 suite-analiza/an_suites.py; brojevi provera i sekunde: docs/DOMEN/DOPLATA.md:307-308, docs/STANJE_REFAKTORA.md:10273, 10741, 8351-8352, docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md:11186
- [M] Van default seta: GoogleSync (5 testova, mreza), MasterSync (11, mreza, u kodu opisana kao zatecena crvena 9/26), SEF (12 testova, 187 statickih tvrdnji), Novac (8 testova, gate=False, pise bez transakcije -- ostavlja TST-NOV redove), ProductionHealthCheck (14 provera, gate=False), Monitoring (6, HTTP, gate=False). Za Banku, Fakturu, Izvestaje, StornoCentar, Agrohemiju, JSON i License trajanje nije zapisano nigde u docs/.
    DOKAZ: python3 suite-analiza/an_suites.py; modBusinessFlowProTests.bas:288; grep -rnE '\([0-9]+[,.]?[0-9]* ?s\)' docs .claude -> samo 8 pogodaka, svi navedeni u ovom izvestaju
- [M] Klasifikacija 350 modTest testova (staticki, kroz zatvaranje poziva unutar modula): 133 cisto citanje/logika; 39 gradi formu bez upisa; 14 menja samo config ili celiju; 43 pise uz AddTableSnapshot + RollbackTx na nivou testa; 121 pise bez rollbacka na nivou testa (9 od njih i gradi formu). Od tih 121, 96 zove produkcione _TX pisce / AppendRow / DeleteRow (pouzdano), a 25 pogadja samo obrazac Storno*/UpdateCell (moguci lazni pogoci). Ukupno 51 test gradi formu.
    DOKAZ: python3 suite-analiza/an_class.py -> '133 A_citanje, 39 B_forma_bez_upisa, 13+1 C0, 41+2 C_upis_sa_snimkom_i_rollback, 112+9 D_upis_bez_test_rollbacka, forma ukupno: 51'; 'D total 121 strong 96'
- [M] Mutirajuci testovi NISU na kraju niza: prvi je na 10. poziciji, 69 od prvih 100 izvrsenih testova pise. Primer obrasca: T_AmbSaldo_ObeStraneJednogReda seje kroz NabaviAmbalazu_TX i UpisiReversAmbalaze_TX, nista ne vraca i tvrdi DELTU; blok doplate (33 testa) pravi svoje maticne podatke i otkupe kroz CreateOtkup_TX i ostavlja ih.
    DOKAZ: an_class.py -> 'first mutating pos 10 count before pos 100: 69'; modTest.bas:7486-7498, 7505-7525; modTest.bas:27809-27816, 27836-27858
- [M] Komentari koji vezuju redosled u RunAllTests: 377-379 (275/276 pre 35 jer 35 ostavlja PRJ-TEST-I1), 382-383 (277/278 mere deltu), 422-426 (339 prvi u bloku doplate -- zbog SABOTAZE, ne zbog testa), 563-566, 573-574, 592-593, 595, 597-599 (citanja pre mutirajucih), 656-657, 663-664 (mutirajuci na kraju), 670. Brojevi u tim komentarima su ZASTARELI: '124-126 DIRAJU fixture i moraju ostati POSLEDNJA' -- danas su 124-126 T_Izv_* cista citanja; '127-131 (Platni nalozi)' stoji iznad RunOne 116-121; '162 MUTIRA' stoji iznad RunOne 141; 'test 22 stornira ZBI-DUPL-2' -- danas je 22 T_Prefill_PoIdentitetuNePoBroju.
    DOKAZ: modTest.bas:377-379, 422-426, 563-566, 573-574, 592-599, 656-657, 663-664, 263; python3 (mapa InvokeTest) -> '124 T_Izv_IdentitetURedu_NeCrtaSe', '116 T_BankaNalozi_UgovorEkrana', '141 T_FakturaGP_WriterKapijeIStorno', '22 T_Prefill_PoIdentitetuNePoBroju'
- [M] Kod sam dokumentuje curenje stanja izmedju testova: raniji test plati FAK-TEST-1; raniji test stornira ZBI-DUPL-2; test vraca podatke jer bi inace menjao ishod testa 25; 'testovi dele svesku' pa se meri delta; prethodni testovi ostavljaju gajbe; raniji test zatvori FX_FAKTURA; T_WriterGuard ostavlja virman sa danasnjim datumom; raniji testovi pune mapu snimaka; 'ovaj test ide 190. po redu -- posle 189 testova koji fixture MENJAJU'. RunOne cisti (CleanupPosleTesta) samo kad test PADNE.
    DOKAZ: modTest.bas:82, 263, 4434-4435, 5525-5528, 7276-7278, 8095-8096, 15616-15617, 20529-20532, 23153-23155, 23368-23371, 26979-26981; modTest.bas:680-703 (CleanupPosleTesta samo u EH, linija 698)
- [M] BFP: 46 koraka su cisto citanje, 111 pise uz snimak+rollback (63 u samom testu, 48 kroz pomocnu rutinu), 235 pise bez rollbacka -- redovi ostaju ('Tests create TST-* rows and do not physically delete data', 'Test redovi (TST-PRO / *-9000x) ostaju'). Ostaje i list BUSINESS_FLOW_PRO_TEST_LOG. 46 testova menja config (SetConfigValue) i mora ga vratiti.
    DOKAZ: python3 suite-analiza/an_bfp.py -> '46 A_citanje, 48 C2, 63 C_, 235 D_upis_bez_rollbacka ... tests with SetConfigValue: 46'; modBusinessFlowProTests.bas:19, 3084, 3081-3085
- [M] BFP testovi NE citaju redove drugog testa po ID-u: modul ima 23 konstante (samo seed sifarnici ST-90001, KOOP-90001...) i modulsko stanje su samo brojaci; 351 od 392 koraka pravi jedinstven kljuc kroz NewScenarioCode (RunID + m_Total+1), 346 koristi NextTestDate (2090. godina + redni broj). Svi zavise od SeedBusinessFlowProMasterData (idempotentan). Zavisnost postoji samo posredno, kroz zbirove nad deljenim sifarnicima (saldo ambalaze, opticaj) -- sto autori zaobilaze merenjem delte ili sopstvenim tipom.
    DOKAZ: grep '^Private Const' modBusinessFlowProTests.bas (23 reda); modBusinessFlowProTests.bas:34-57, 32722-32729, 1252-1276; komentari 80-81, 5315-5316, 25762-25763, 27233-27235; an_bfp.py -> 'NewScenarioCode users: 351 NextTestDate users: 346'
- [M] BFP vise nije 'suite za praznu svesku' kako pise u zaglavlju: 6 testova (S7R2/S7R3) cita redove koje pravi tools/make_fixture.py (FAK-SLED-GP4, UT-SLED-P, FST-SLED-P), dakle zavisi od modTest fixture-a. U podrazumevanom prolazu BFP radi nad sveskom u kojoj je RunAllTests vec ostavio redove iz >=96 testova, a Agrohemija i Golden rade nad sveskom sa svim BFP redovima; samostalan --suite poziv krece od ciste kopije. Stanje na pocetku suite zavisi od toga kako je pozvana.
    DOKAZ: modBusinessFlowProTests.bas:8, 8476-8487, 9152-9180; grep -c FAK-SLED-GP4 tools/make_fixture.py -> 3, src-vba/modTest.bas -> 0; tools/run_vba.py:837-839 (jedna temp kopija po pozivu), 944-949 (suite redom u istoj svesci)
- [M] Suite same NE upravljaju stanjem aplikacije: RunAllTests ne gasi ScreenUpdating/Calculation/EnableEvents (jedino clsTransaction to radi dok TX traje). run_vba postavlja Visible=False i EnableEvents=False. U izvoru nema nijedne .Formula dodele, pa prelazak Calculation na Manual verovatno ne nosi dobitak.
    DOKAZ: modTest.bas:281-675 (nema Application.*); grep -c ScreenUpdating/Application.Calculation po test modulima -> 0 svuda osim modTest (2/3, u pojedinacnim testovima); clsTransaction.cls:61-69, 245-251; tools/run_vba.py:888-891; grep '\.Formula' src-vba/*.bas -> 0 pogodaka van komentara
- [M] RunAllTests i BFP NE pale tihi rezim journaling-a (pale ga samo Banka, Agrohemija i SEF). Posledica po SVAKOM AppendRow: WriteJournalRow = 2x Dir() + Open For Append + Print + Close CSV fajla pored sveske. Po SVAKOM CommitTx: MarkDirtyAndSchedule (otkaz + zakazivanje Application.OnTime + LogInfo -> jos jedan Open/Append/Close log fajla + Debug.Print) i InvalidateStornoWarm (jos jedan par OnTime poziva, koji tihi rezim NE gasi). Nijedna sabotaza u katalogu nije usidrena u modJournaling.bas, modLogError.bas ni modStornoWarm.bas.
    DOKAZ: grep -n 'SetTestModeQuiet True' src-vba/*.bas -> samo modAgrohemijaTests.bas:61, modSEFTests.bas:953/1164, modTestBanka.bas:92; modDataAccess.bas:294-295; modJournaling.bas:66-123, 801-827, 849-857; modLogError.bas:32-92; clsTransaction.cls:116, 121; modStornoWarm.bas:42-55; grep -c '"modJournaling.bas"' tools/sabotaza.py -> 0, isto za modLogError/modStornoWarm
- [M] OnTime tajmeri koje commit zakaze (AutoSaveTick posle 60 s, odnosno odmah kad je od poslednjeg snimanja proslo 600 s; StornoWarmTick posle 60 s) ostaju zakazani posle suite; run_vba ih ne otkazuje, a AutoSaveTick nad prljavom sveskom radi stvarni ThisWorkbook.Save. Dokumentacija belezi neistrazen pad Excela ('RPC failed') u punom prolazu posle ~10 min rada, dok je ista suite sama zelena.
    DOKAZ: modJournaling.bas:40-41, 816-825, 830-835; modStornoWarm.bas:26, 49-55, 63-72; grep 'StopAutoSaveTimer|StopStornoWarm|OnTime' tools/run_vba.py -> 0 pogodaka; docs/STANJE_REFAKTORA.md:8346-8354
- [M] BFP po SVAKOJ tvrdnji (LogPass/LogFail/LogSkip) radi Debug.Print i AppendTestLog: trazenje poslednjeg reda (End(xlUp)) + 7 pojedinacnih upisa u celije lista. Pri ~3000 tvrdnji to je ~21.000 upisa u celiju i ~3000 Debug.Print po prolazu. run_vba pri compile probi ucini VBE prozor vidljivim i vise ga ne sakriva.
    DOKAZ: modBusinessFlowProTests.bas:32764-32770, 32829-32846; grep -n 'MainWindow.Visible' tools/run_vba.py -> samo linija 530 (= True)
- [M] Svaki od 51 testa sa formom gradi CELU ljusku: NewOtkupUIForm = New frmOtkupUI + Controls.count -> UserForm_Initialize -> OtkupUI_Init -> BuildOtkupScreen; ReleaseOtkupUIForm je Unload + OtkupUI_Release. Produkcioni komentar nosi merenje: ekran dokumenata 863 kontrole, od toga 623 deljeni hrom i mreza, prosecan ekran 240 kontrola ~180 ms.
    DOKAZ: modTest.bas:17180-17197, 18017-18020; frmOtkupUI.frm:43-52; modOtkupUI.bas:5608-5637, 5818-5821
- [M] RunAllTests izvrsava do 5 STVARNIH ThisWorkbook.Save nad temp sveskom sa sveze uvezenim projektom od 194 komponente (testovi AutoSave/Save/TxRollback) i 2 fiksna cekanja Application.Wait od 1 s (PDF testovi).
    DOKAZ: modTest.bas:17648, 17846, 17907, 18328, 18403 (Save koji prolazi); modTest.bas:10415, 10465
- [M] Produkciona vruca putanja koju svaki test placa: GetTable linearno prolazi sve listove i sve ListObject-e pri SVAKOM pozivu (50 tabela u kanonu); kes tabela/kolona radi samo unutar BeginTableCache prozora (sam kod kaze '~80 GetTable skenova'); GetConfigValue = LookupValue = puno citanje config tabele + 3 skena (311 mesta poziva); RequireColumnIndex ima 1617 mesta poziva; AppendRow pise celiju po celiju, a StampRowAudit 4 puta petljom trazi audit kolonu po imenu (i pri svakom UpdateCell).
    DOKAZ: modDataAccess.bas:118-131, 133-153, 197-234, 286-295, 382-423, 631-643; modConfig.bas:1273-1285; schema/schema.json (50 tabela, 684 kolone); grep -c RequireColumnIndex -> 1617, GetConfigValue( -> 311
- [M] AddTableSnapshot kopira CELU tabelu (DataBodyRange.Value2), a RestoreTable vraca red po red (ListRows.Add / ListRows(n).Delete u petlji) pa upise ceo niz. U BFP-u ima 294 mesta AddTableSnapshot (najvise TBL_AMBALAZA: 45), a tabele kroz prolaz samo rastu jer 235 koraka ostavlja redove -- cena snimka i citanja po testu raste sa rednim brojem testa.
    DOKAZ: clsTransaction.cls:73-90, 262-302; python3 (brojanje AddTableSnapshot) -> 'modBusinessFlowProTests.bas 294 [(TBL_AMBALAZA, 45), ...]', 'modTest.bas 56', 'modGoldenTests.bas 21'
- [M] Rezultat-fajl se pise TEK NA KRAJU suite (RunAllTests: WriteResultFile kao poslednja naredba; BFP: u EndRun). Ako Excel visi ili ga timeout ubije, nema nijednog reda -- ne zna se koji test je radio, a ceo prolaz se ponavlja.
    DOKAZ: modTest.bas:674, 18712-18716; modBusinessFlowProTests.bas:3107-3110, 32783-32791; tools/run_vba.py:132-135 ('suite nije upisala ... nije stigla do kraja')
- [M] Rast izmeren iz git istorije (10.10 -> 08.11.2026, 29 dana): modTest 17.948 -> 31.061 linija (+452/dan), 213 -> 350 testova (+4,7/dan); BFP 23.698 -> 33.077 linija (+323/dan), 305 -> 398 Test_ procedura; sabotaza.py 8.804 -> 13.801 linija. Fajlovi: modTest.bas 1,58 MB, BFP 1,53 MB.
    DOKAZ: git show <commit>:src-vba/modTest.bas | wc -l za commit-e 09d9c46 (10.10), 2ea80dd, 6ef0b77, 3aa6e6f, 7bae8cd, 1fa4937 (08.11); ls -la src-vba/modTest.bas src-vba/modBusinessFlowProTests.bas
- [M] Najvece procedure u test modulima su sami registri: TestName 22.178 bajtova (357 linija), InvokeTest 17.401, RunBusinessFlowProSuite 16.886; najveci test T_FakturaGP_WriterKapijeIStorno 23.574 bajta (443 linije). Prag vba_check PROCEDURA_VELIKA je 40.000 bajtova izvora. TestName raste ~63 bajta po testu, pa prag prelazi na ~631 testu -- za ~60 dana pri +4,7 testa/dan.
    DOKAZ: python3 (bajtovi procedure sa razmacima) -> 'TestName lines 357 raw bytes 22178', 'InvokeTest ... 17401', 'RunBusinessFlowProSuite ... 16886', 'T_FakturaGP_WriterKapijeIStorno ... 23574'; tools/vba_check.py:2362
- [N] VBA granice su vec jednom ujele ovaj projekat: 'Procedure too large' na modPoruke.UpsertPoruke (1657 poziva, ~200 KB izvora) uhvacen tek rucnim Compile-om posle spajanja dve grane; i 'Out of memory' pri uvozu u jednom makrou nad kompajliranim projektom (14.09.2026, kad je modTest imao 17,7k linija -- danas ima 31k).
    DOKAZ: docs/UI_MIGRACIJA_KATALOG.md:7103-7127; docs/EXCEL_TEST_HARNESS.md:446-451; src-vba/modVbaTools.bas:65
- [M] Test kod je 79.151 linija od 233.684 (34%) i 3,74 MB od 10,46 MB VBA izvora (35,8%). run_vba ga uvozi CELOG pri svakom pozivu (Remove + Import svake komponente), a objava release-a salje sve .bas/.cls/.frm iz src-vba bez filtera -- test moduli idu i klijentu kroz self-update.
    DOKAZ: wc -l src-vba/mod*Test*.bas; python3 (os.path.getsize) -> 'test bytes 3741146 all code bytes 10456658 share 35.8%'; tools/run_vba.py:382-406; src-vba/modRelease.bas:78-99
- [M] Tri rucna registra u modTest zauzimaju 1.106 linija (RunAllTests 395 + TestName 357 + InvokeTest 354) i svaki nov test trazi tri unosa na tri mesta; vba_check REGISTAR postoji bas zato sto su se 'vec razisli pri rebase-u'. Identitet testa u katalogu sabotaza je IME (npr. T_BrojKapija_IstoZaSvakiCase), ne broj. BFP nema registar: test je jedan red u RunBusinessFlowProSuite (435 linija).
    DOKAZ: tools/vba_check.py:1888-1905; tools/sabotaza.py:6705-6712 (cetvrti clan zapisa je ime testa); an_size.py
- [M] Katalog sabotaza priznaje samo tri test fajla (modTest, modTestBanka, BFP) i mapu modul->suite za ta tri. Zato tvrdnje o drugim oblastima zavrsavaju u modTest umesto u svom modulu -- sam kod to kaze. Uz to, komentar u sabotaza.py 'modTest ne pise u tabele' vise nije tacan (>=96 testova pise).
    DOKAZ: tools/sabotaza.py:86-88, 95, 100-104; src-vba/modTest.bas:7572-7575
- [M] Preklapanje modTest <-> BFP: modTest referencira 37 produkcionih _TX pisaca, BFP 41, zajednicka su 22 (CreateOtkup_TX, StornoOtkup_TX, UpisiReversAmbalaze_TX, NabaviAmbalazu_TX...). Ambalaza je najjace udvojena: UpisiReversAmbalaze_TX 58 mesta u modTest / 24 u BFP; NabaviAmbalazu_TX 45 / 32. Oba modula imaju sopstvene fabrike scenarija (modTest: AmbSejRevers, DopOtkup, KartOtkup, S6K6Zbirna...; BFP: OtkHeader/OtkStavka/Pr3Header...). Od 92 produkciona _TX pisca 66 referencira bar jedan test modul.
    DOKAZ: python3 suite-analiza/an_overlap.py -> '_TX writers: modTest 37 BFP 41 both 22 ... prod _TX total 92 referenced by any test module 66'; grep -c po imenu pisca
- [M] Kopirane pomocne rutine medju test modulima: 19 imena definisano u vise modula (Chk/ChkEq/ChkEqD/Fail/ReportResults u Banka/Storno/Palete; AssertTrue/AssertEquals/LogPass/LogFail u BFP/SEF/GoogleSync/License; SeedKupac/SeedKooperant/SeedStanica u BFP/Banka/Agro), nijedno telo nije identicno, ukupno ~222 suvisne linije. Sustinski problem nije obim nego to sto postoji najmanje 6 razlicitih mini-okvira sa tri razlicita kanala verdikta (rezultat-fajl, Err.Raise, MsgBox+watchdog).
    DOKAZ: python3 suite-analiza/an_dup.py -> 'procedure names defined in >1 test module: 19 ... identical bodies: 0 / 19 ; redundant body lines: 222'
- [M] Prepreke paralelnim Excel instancama: (a) run_vba pise izvestaj na FIKSNU putanju tests/last_run.json i tests/last_run.txt; (b) compile proba trazi vidljiv i aktivan VBE prozor i pada na SendKeys uz AppActivate -- zavisi od fokusa; (c) SRC_VBA je fiksan na src-vba radnog stabla (dokaz mutira izvor u mestu); (d) test u modTest koristi fiksan folder %TEMP%\agrix_backup_test. Ono sto vec radi: svaki poziv ima svoju temp kopiju sveske i svoj Excel proces (DispatchEx), a watchdog je vezan za PID.
    DOKAZ: tools/run_vba.py:1064-1067, 1123; 504-505, 530-533, 572, 631-635; 96; 837-839, 875-878; src-vba/modTest.bas:17763-17779
- [M] Grana B (stablo /home/user/otkupapp-pwa, commit 9d6e780) vec sadrzi merni deo koji je odvojiv od logike izbora: Timer oko testa u RunOne i fajl last_run_vreme.txt, za BFP omotac BfpIzabran oko 364 poziva, i dokaz.py --izolacija koji nabraja testove koji sami ne prolaze. Njen VBA deo nikad nije izvrsen u Excelu.
    DOKAZ: /home/user/otkupapp-pwa/src-vba/modTest.bas:289-292, 691, 784-802; /home/user/otkupapp-pwa/src-vba/modBusinessFlowProTests.bas:63-71, 135-140; git show --stat 9d6e780 ('VBA strana NEVERIFIKOVANA')
- [M] Nezavisnost testova nije nigde propisana kao ugovor: ni .claude/rules/testovi.md ni docs/EXCEL_TEST_HARNESS.md nemaju pravilo o redosledu ili izolaciji testova.
    DOKAZ: grep -n -i 'paralel|shard|izolacij|redosled|nezavis' .claude/rules/testovi.md docs/EXCEL_TEST_HARNESS.md -> nijedan pogodak koji govori o medjusobnoj zavisnosti testova

## TROSKOVI
[
 {
  "sta": "Jedan prolaz RunAllTests",
  "koliko": "215,8 s pri 336 testova (0,64 s/test); pri 350 testova procena ~225 s",
  "koliko_cesto": "svaki TARGETED nad modTest-om, svaki FULL, i svaki prolaz dokaza cija je meta test iz modTest-a (801 sabotaza)",
  "izmereno": false
 },
 {
  "sta": "Jedan prolaz BFP",
  "koliko": "210 s pri 2417 tvrdnji (07.10.2026); danas 3014+ tvrdnji, novije merenje ne postoji; ceo podrazumevani poziv > 600 s",
  "koliko_cesto": "svaki FULL i svaki prolaz dokaza nad BFP-om (456 sabotaza)",
  "izmereno": false
 },
 {
  "sta": "Granicni trosak novog modTest testa",
  "koliko": "~1,08 s po dodatom testu (izvedeno: (215,8-73,5)/(336-204)); pri +4,7 testa/dan to je ~+5 s trajanja suite dnevno, a kroz dokaz se mnozi brojem prolaza",
  "koliko_cesto": "trajno, raste",
  "izmereno": false
 },
 {
  "sta": "Pun katalog dokaza u grupama po 6 (izvedena procena iz dokumentovanih trajanja)",
  "koliko": "215 prolaza: ~134 nad RunAllTests x (18+216 s) = ~8,7 h, ~76 nad BFP x (18+210..300 s) = 4,8-6,7 h; ukupno reda velicine 13-15 h",
  "koliko_cesto": "pred release",
  "izmereno": false
 },
 {
  "sta": "Gradnja cele ljuske forme u 51 testu",
  "koliko": "procena 25-33 s po prolazu RunAllTests (623 kontrole hroma + ekran; izvedeno iz produkcionog merenja '240 kontrola ~180 ms'), tj. 11-15% suite",
  "koliko_cesto": "svaki prolaz RunAllTests",
  "izmereno": false
 },
 {
  "sta": "Stvarni ThisWorkbook.Save u testovima (5 puta) + 2x Application.Wait 1 s",
  "koliko": "procena 10-25 s po prolazu (snimanje .xlsm sa sveze uvezenih 194 komponente; nije mereno)",
  "koliko_cesto": "svaki prolaz RunAllTests",
  "izmereno": false
 },
 {
  "sta": "Journal CSV + log + OnTime po upisu i commit-u u RunAllTests i BFP",
  "koliko": "po AppendRow: 2 Dir + 1 otvaranje/zatvaranje fajla; po CommitTx: 1 otvaranje/zatvaranje log fajla + 4 OnTime poziva. Broj upisa po prolazu nepoznat (BFP ima 726 mesta poziva _TX pisaca); cena po fajl-operaciji zavisi od antivirusa (1-20 ms) -- raspon sekunde do vise desetina sekundi",
  "koliko_cesto": "svaki prolaz RunAllTests i BFP",
  "izmereno": false
 },
 {
  "sta": "BFP log po tvrdnji",
  "koliko": "~21.000 upisa u celiju + ~3.000 End(xlUp) + ~3.000 Debug.Print po prolazu; procena 1-5 s",
  "koliko_cesto": "svaki prolaz BFP",
  "izmereno": false
 },
 {
  "sta": "Dijalozi u 7 default suita (potvrda na pocetku i/ili MsgBox na kraju)",
  "koliko": "~11 dijaloga x do 0,4 s cekanja watchdog-a = do ~4,4 s; vazniji je rizik visenja ako watchdog ne prepozna prozor",
  "koliko_cesto": "svaki FULL",
  "izmereno": false
 },
 {
  "sta": "Uvoz test koda pri svakom run_vba pozivu",
  "koliko": "79.151 linija / 3,74 MB (34-36% izvora) se uvozi i kad se pusta jedna mala suite; udeo u reziji od ~18 s nije meren",
  "koliko_cesto": "svaki run_vba poziv i svaki prolaz dokaza",
  "izmereno": true
 },
 {
  "sta": "Ponovljeni prolazi zbog timeout-a i padova bez imena testa",
  "koliko": "ceo poziv (do 600 s) se baca; zabelezeno najmanje 2 puta (prekoracenje 600 s) + neistrazen 'RPC failed' u punom prolazu",
  "koliko_cesto": "povremeno, pri FULL prolazu",
  "izmereno": false
 },
 {
  "sta": "Rucno odrzavanje tri registra",
  "koliko": "3 unosa na 3 mesta po testu u fajlu od 31k linija; 1.106 linija registra; sukobi pri rebase-u (grana B se sudara bas u modTest/BFP/sabotaza)",
  "koliko_cesto": "~4,7 puta dnevno",
  "izmereno": true
 },
 {
  "sta": "Zastareli komentari o redosledu i zavisnostima",
  "koliko": "najmanje 5 komentara u RunAllTests + 1 kod konstanti imenuje brojeve testova koji danas znace drugi test; zavisnost se mora ponovo otkrivati citanjem ili padom",
  "koliko_cesto": "pri svakom pomeranju ili dodavanju mutirajuceg testa",
  "izmereno": true
 },
 {
  "sta": "Fixture kao zaobilaznica za deljeno stanje",
  "koliko": "tools/make_fixture.py ima 3.171 liniju; namenski redovi 'koje nijedan drugi test ne dira' se dodaju po scenariju; regeneracija trazi Windows + Excel + donor svesku",
  "koliko_cesto": "pri svakom novom scenariju koji trazi netaknute redove",
  "izmereno": true
 }
]

## PREDLOZI
### T-01 Merenje pre svega: vreme po testu i otisak tabela pre/posle testa [trud S] [slabi: NE]
U modTest.RunOne (modTest.bas:680-703) staviti Timer oko InvokeTest i pisati red 'sekunde TAB status TAB ime' u last_run_vreme.txt -- samo merni deo commit-a 9d6e780 iz grane B (njen modTest.bas:289-292, 691, 784-802), bez logike izbora koja se sudara sa main-om. Za BFP prvo bez izmene koda: run_vba --suite RunBusinessFlowProSuite --keep, pa vremena iz kolone Timestamp lista BUSINESS_FLOW_PRO_TEST_LOG (rezolucija 1 s); kasnije omotac po testu. Uz to dijagnosticki (ne ulazi u verdikt) otisak: broj redova svih 50 tabela pre i posle svakog testa -> spisak 'OSTAVLJA STANJE: test, tabela, +n'.
  DOBITAK: Nema direktnog ubrzanja. Pretvara sve procene iz ove analize u merenja: 10 najsporijih testova, udeo forme/Save/journal-a, i TACAN broj testova koji ostavljaju stanje (danas samo staticka granica 96-121). Bez ovoga se ostali predlozi ne mogu ni rangirati ni dokazati.
  RIZIK: Nizak. Timer se u ponoc vraca na nulu (grana B to vec obradjuje). Otisak kosta ~50 citanja ListRows.Count po testu (procena < 2 s po prolazu).
  PREDUSLOVI: Jedan Windows+Excel prolaz. Ne zavisi od spajanja grane B.
  DOKAZATI: Dvosmerno za sam merac: test koji namerno ceka 1 s mora dobiti >= 1 s; zbir vremena po testu mora biti u +-5% od sekundi koje run_vba ispise za suite; test koji namerno ostavi red mora biti imenovan u otisku, a kad se red vrati -- ne sme. Pokvaren merac (npr. Timer citan posle upisa) mora oboriti te provere po imenu.

### T-02 Inkrementalni rezultat-fajl: pad ili timeout imenuje test koji je radio [trud S] [slabi: NE (strozije: nedovrsen prolaz i dalje pada, ali sada sa imenom).]
RunOne i BFP upisuju po jedan red PRE testa ('RUN ime') i POSLE ('OK/FAIL ime') dopisivanjem u fajl, a prvi red 'TESTS=n FAIL=m' se pise na kraju kao i danas (ili u zaseban fajl napretka da format last_run.txt ostane isti za dokaz.py). run_vba kad ne nadje zavrsni red ispise poslednji 'RUN' red.
  DOBITAK: Ubijen ili obesen prolaz (do 600 s bacenog rada) daje ime testa umesto 'suite nije stigla do kraja'; delimicni rezultati su upotrebljivi. Direktno skracuje dijagnostiku neistrazenog 'RPC failed' i visenja na MsgBox-u u pisac-putanji.
  RIZIK: Nizak; 350 (odnosno ~400) dopisivanja u fajl po prolazu (< 1 s). Paziti da se format koji cita dokaz.py ne promeni.
  PREDUSLOVI: Nema.
  DOKAZATI: Test-seam koji usred suite izazove prekid (ili --timeout 5): izlaz run_vba mora da imenuje bas taj test; bez izmene ne imenuje nista. run_vba --self-test dobija slucaj 'fajl napretka bez zavrsnog reda'.

### T-03 Ugasiti sporedne efekte koje nijedna tvrdnja ne meri: journal CSV, log i OnTime tajmeri u RunAllTests, BFP i ostalim suitama [trud S] [slabi: NE za imenovane tvrdnje (0 sabotaza usidreno u modJournaling/modLogError/modStornoWarm). Gubi se USPUTNO izvrsavanje WriteJournalRow i MarkDirtyAndSchedule pri svakom upisu u testovima; nadoknadjuje se jednim imenovanim testom iz opisa, koji je jaci od danasnjeg stanja (danas te putanje niko ne tvrdi).]
Na ulazu u RunAllTests (posle SetTestMode True, modTest.bas:284) i u BFP BeginRun/EndRun primeniti isti obrazac koji Banka vec ima (modTestBanka.bas:91-92, 165-168): zapamti IsTestModeQuiet, SetTestModeQuiet True, vrati na kraju. Isto za Golden, Storno, Palete, StornoCentar, Faktura, Izvestaj. Na kraju svake suite (ili u run_vba pre Close) pozvati modJournaling.StopAutoSaveTimer i modStornoWarm.StopStornoWarm, jer InvalidateStornoWarm nije pod tihim rezimom (clsTransaction.cls:121). Dodati JEDAN imenovan test koji sa ukljucenim journal-om tvrdi da AppendRow upise red u CSV i da commit zakaze AutoSave -- da ta pokrivenost bude izricita umesto usputna.
  DOBITAK: Uklanja 2 Dir + otvaranje/zatvaranje fajla po svakom AppendRow i otvaranje/zatvaranje log fajla + 4 OnTime poziva po svakom CommitTx u dve najvece suite. Velicina nije merena: hipoteza 5-25% trajanja RunAllTests/BFP (zavisi od antivirusa); meri se kroz T-01. Sporedni moguci dobitak: nestaje kandidat za neistrazen 'RPC failed' posle ~10 min (tajmer koji opali izmedju dve suite i snima svesku).
  RIZIK: Nizak. T_AutoSave_* testovi zovu AutoSaveAfterCommit direktno, pa ih tihi rezim ne dotice (modTest.bas:17623-17664, 17897-17921). Treba proveriti da nijedan test ne tvrdi posledicu MarkDirtyAndSchedule (grep nije nasao nijedan).
  PREDUSLOVI: T-01 (da se dobitak izmeri). Izmena je samo u test modulima; ako se tajmeri gase u run_vba, to je izmena alata.
  DOKAZATI: Pre/posle: isti skup imena testova i isti TESTS/FAIL u last_run.txt i last_run_bfp.txt; vreme po testu iz T-01 pokazuje razliku; u temp folderu posle prolaza nema Journal CSV fajlova (danas ih ima). Novi journal test dobija svoju sabotazu (WriteJournalRow bez upisa) i mora pasti po imenu. Pun dokaz pred release ostaje isti.

### T-04 BFP: log u memoriji umesto 7 celija po tvrdnji; sakriti VBE prozor posle compile probe [trud S] [slabi: NE (list sa logom nije kanal verdikta; verdikt je last_run_bfp.txt + Err.Raise).]
AppendTestLog (modBusinessFlowProTests.bas:32829-32846) skuplja redove u niz i upisuje ih jednim Range.Value u EndRun i u LogFatal; tools/read_test_log.py i dalje cita isti list uz --keep. U run_vba posle compile probe vratiti vbe.MainWindow.Visible = False (danas ostaje True od linije 530), da Debug.Print ne crta u Immediate prozor.
  DOBITAK: ~21.000 upisa u celiju i ~3.000 End(xlUp) manje po BFP prolazu; procena 1-5 s po prolazu (nije mereno), puta broj BFP prolaza u dokazu (~76 u grupama po 6).
  RIZIK: Nizak. Ako Excel padne usred suite, bafer se gubi -- zato T-02 (fajl napretka) ide pre ili zajedno. Sakrivanje VBE prozora ne sme da pokvari compile probu (radi se POSLE nje).
  PREDUSLOVI: T-02 pozeljno.
  DOKAZATI: Sadrzaj lista BUSINESS_FLOW_PRO_TEST_LOG posle --keep prolaza identican po redovima (osim Timestamp) starom; last_run_bfp.txt identican; vreme iz T-01.

### T-05 Rani alarm za VBA granice u vba_check + jedna proba granice u Excelu [trud S] [slabi: NE]
Dodati u vba_check nalaz MODUL_VELIK (upozorenje na 45.000 linija, greska na 60.000) po uzoru na postojeci PROCEDURA_VELIKA (tools/vba_check.py:2346-2380), sa self-test slucajevima u oba smera. Jednom, rucno: sinteticki modul od ~65.600 linija kroz run_vba --compile-only da se potvrdi da granica od 65.534 linije vazi u ovoj verziji Excela.
  DOBITAK: Granica se vidi u FAST kapiji mesecima ranije, umesto kao pad rucnog Compile-a posle spajanja grana (tako je vec jednom bilo sa 'Procedure too large'). Pri sadasnjem rastu modTest stize na 65.534 linije za ~76 dana, BFP za ~100, a TestName prelazi 40.000 bajtova za ~60 dana.
  RIZIK: Nizak; lazna uzbuna ako se prag postavi prenisko (zato dva nivoa).
  PREDUSLOVI: Nema. Ako se pravilo upisuje u .claude/rules/testovi.md, to ide zasebnim process PR-om.
  DOKAZATI: vba_check --self-test: modul preko praga MORA dati nalaz, modul ispod NE SME; podignut prag obara self-test po imenu (isti obrazac kao PROC_SIZE_CASES).

### T-06 Generisan registar testova umesto tri rucna spiska (identitet = ime) [trud M] [slabi: NE. Kapija REGISTAR u vba_check se zamenjuje ili zadrzava kao druga linija; nova kapija (--check) mora biti pokazana crvena za sva cetiri danasnja razlaza (nema u InvokeTest / nema u TestName / nema u RunOne / pogresno ime).]
Nov alat tools/gen_test_registar.py po uzoru na gen_schema_module.py: kanon su same procedure 'Private Sub T_*' (redosled = redosled u fajlu, uz masinski citljive anotacije za izuzetke, npr. '@posle T_X' ili '@grupa lanac'); artefakt su RunAllTests lista, TestName i InvokeTest izmedju GEN markera u modTest.bas (direktni pozivi ostaju -- compile signal se ne gubi). --check u CI obara razlaz. Numericki ID postaje interni redni broj; ime ostaje identitet (katalog sabotaza vec radi po imenu). Dispatch se generise u komadima od po ~100 Case grana da nijedna procedura ne raste ka pragu. Za BFP isti generator proizvodi listu poziva u RunBusinessFlowProSuite i hvata 'Test_ postoji a niko ga ne zove'.
  DOBITAK: Nov test = jedna procedura umesto 4 izmene; 1.106 linija registra prestaje da se pise rukom; nestaje klasa rebase sukoba u registrima (bas tu se sudara grana B); nestaje rok od ~60 dana za PROCEDURA_VELIKA na TestName; zavisnosti redosleda postaju deklaracija koju alat proverava, umesto komentara sa zastarelim brojevima.
  RIZIK: Srednji. Generator i --check su novi checker -> obavezan dvosmerni dokaz. Prva generacija menja ~1.100 linija u modTest-u (sukob sa otvorenim granama) -- treba je uraditi u mirnom trenutku i jednim commit-om. Redosled izvrsavanja mora ostati IDENTICAN danasnjem u prvom koraku.
  PREDUSLOVI: Dogovor o trenutku (minimum otvorenih grana nad modTest.bas). Windows+Excel prolaz i rucni Compile posle generacije.
  DOKAZATI: (1) Posle prve generacije: last_run.txt ima istih 350 imena ISTIM redom i isti TESTS/FAIL; (2) gen_test_registar.py --self-test sa slucajevima u oba smera; (3) rucno obrisan Case ili dodat T_ bez regeneracije -> --check pada po imenu testa; (4) sabotaza.py --proveri-sidra ostaje zelen.

### T-07 Jedan test-kit i jedan format rezultata za SVE suite; bez MsgBox-a pod harnesom [trud M] [slabi: NE; jaca (imenovan pad iz svih suita). Paznja: suite koja je do sada bila 'gate bez fajla' ne sme tokom prelaza postati BLIND.]
Izdvojiti zajednicki modul (radno ime modTestKit) sa: brojacima, AssertEq/Chk varijantama, upisom rezultata 'TESTS=n FAIL=m' + redovi po tvrdnji, i jednim nacinom podizanja verdikta. Postojece suite ga usvajaju postepeno (Banka/Storno/Palete dele Chk/ChkEq/ChkEqD/Fail/ReportResults -- tri kopije; BFP/SEF/GoogleSync/License dele AssertTrue/AssertEquals/LogPass/LogFail). Potvrdni i zavrsni MsgBox se preskace kad je IsTestMode/harness aktivan (rucno pokretanje iz Alt+F8 zadrzava dijalog). Svaka gate-suita dobija result_file u SUITES. Posle toga sabotaza._TEST_FAJLOVI i _SUITA_PO_MODULU (tools/sabotaza.py:95-104) mogu da prime i ostale test module.
  DOBITAK: 8 default gate-suita koje danas prijavljuju samo 'Exception occurred' dobijaju ime pale tvrdnje; tvrdnje mogu da zive u svom modulu umesto u modTest-u (uklanja uzrok rasta zapisan u modTest.bas:7572-7575); ~11 dijaloga manje po FULL prolazu (do ~4 s i jedan izvor visenja manje); ~220 kopiranih linija manje.
  RIZIK: Srednji: dira 10+ test modula; javna imena u zajednickom modulu ne smeju da se sudare (AssertEq je vec Public u modTest-u, a Private u modLicenseTests). Radi se suita po suita.
  PREDUSLOVI: T-02 (format fajla). Prosirenje _TEST_FAJLOVI je izmena alata sa sopstvenim dokazom.
  DOKAZATI: Po suiti: isti broj provera pre/posle (Banka 241, Storno 111, Palete 115, Faktura 39...); namerno oborena jedna tvrdnja mora stici u izvestaj run_vba PO IMENU (danas kod 8 suita ne stize); run_vba bez watchdog klikova u dnevniku dijaloga.

### T-08 Podela modTest i BFP po oblasti (mehanicki, skriptom) + suita po oblasti [trud L] [slabi: NE, ako se dokaze da je skup (ime testa, verdikt) i broj tvrdnji isti pre i posle. Jedna stvar se menja i mora se reci: ciljani prolaz po oblasti vise ne izvrsava ostale oblasti, pa tvrdnja 'cela RunAllTests zelena' ostaje samo za FULL.]
Posle T-06/T-07: testovi se sele u module po oblasti. Prirodne granice vec postoje u imenima -- modTest: Doplata (33 testa, 1.896 linija), KI (31), S6 (18), Izv (15), Amb* (~30), Banka* (13), Fak (9), Sema (9), ljuska/UI (~25)...; BFP: OTP (79), OTK (62), ZBR (48), S6K (33), PR3 (24), Amb (24), LNC (18), S7R (15), PWA (12), BKTX (12). Zajednicke konstante FX_* i pomocne rutine idu u deljeni modul (Public), harness u modTestKit. Selidbu radi SKRIPTA po mapi 'procedura -> modul' (ponovljiva nad bilo kojom granom, da se sukobi ne resavaju rukom). Svaki modul dobija generisanu ulaznu tacku i red u SUITES; generator pise manifest 'test -> modul -> suita' koji citaju sabotaza (--proveri-sidra) i dokaz (koju suitu pustiti za metu). RunAllTests ostaje kao zbirna ulazna tacka.
  DOBITAK: TARGETED postaje stvarno ciljan: oblast od ~30 testova je ~20 s + ~18 s rezije umesto 216 + 18 s (oko 6x krace). Prolaz dokaza po sabotazi ~40-60 s umesto ~234 s (modTest) odnosno ~230-320 s (BFP) -- 4-6x po prolazu, pre grupisanja (izvedeno iz dokumentovanih trajanja, nije mereno). Uklanja rok od 65.534 linije i rizik 'Out of memory' pri uvozu modula od 1,5 MB. VBA kompajlira po modulu na zahtev, pa ciljana suita ne povlaci 31k linija.
  RIZIK: Visok po obimu: dira dva najveca fajla i tri alata (vba_check REGISTAR, sabotaza, dokaz). Javna imena posle podele mogu da se sudare ('Ambiguous name') -- vidi se tek u Excel Compile-u. Grupno pravilo dokaza 'ista suita' menja znacenje kad suita postane oblast. Sukob sa svakom otvorenom granom nad ova dva fajla.
  PREDUSLOVI: T-06, T-07, T-09 (da se zna koje grupe moraju ostati zajedno). Rucni Compile. Pun dokaz (ne grupni) bar jednom posle podele.
  DOKAZATI: (1) Unija last_run fajlova po oblasti = istih 350 (odnosno 391) imena i isti zbir tvrdnji (3386 AssertEq staticki; BFP isti broj PASS redova); (2) sabotaza.py --proveri-sidra zelen nad manifestom; (3) u rezu dokaz nad uzorkom sabotaza iz SVAKE preseljene oblasti, pun katalog pred release; (4) vba_check MODUL_VELIK i PROCEDURA_VELIKA bez nalaza.

### T-09 Revizija nezavisnosti testova: obrnut i nasumican redosled, pa deklaracija zavisnosti [trud M] [slabi: NE (otkriva skrivene pretpostavke; nijedna tvrdnja se ne uklanja).]
Pustiti RunAllTests i BFP (a) obrnutim redom, (b) u 3-5 nasumicnih permutacija sa zapisanim semenom, (c) svaki test sam (dokaz.py --izolacija iz grane B radi upravo ovo). Svaki test kome se verdikt promeni je zavisan. Zavisne ili popraviti (sopstveni redovi, delta umesto apsolutnog broja, snimak+rollback) ili zapisati kao deklaraciju u generisanom registru ('@posle', '@grupa'). Zastarele komentare o redosledu (modTest.bas:563-599, 656-657, 263) zameniti deklaracijama. Isto na nivou suita: svaka default suita mora biti zelena SAMA nad cistom kopijom fixture-a.
  DOBITAK: Tacan, izmeren spisak zavisnih testova (danas: ~12 mesta u komentarima i staticka gornja granica 121). Preduslov za izbor testova, shard-ove i pouzdano grupisanje sabotaza. Uklanja klasu laznih padova 'test pao bez svoje krivice' pri pomeranju testova.
  RIZIK: Srednji po trajanju: N punih prolaza u Excelu (svaki 4-10 min). Nasumican redosled moze otkriti vise zavisnosti nego sto se ocekuje, narocito u BFP-u kroz deljene saldo-zbirove.
  PREDUSLOVI: T-06 (redosled kao podatak, da permutacija ne trazi rucno preslaganje 350 linija) ili spajanje mernog/izbornog dela grane B posle njene Excel verifikacije.
  DOKAZATI: Kriterijum: isti skup PASS/FAIL po imenu u svim permutacijama. Dvosmerno: namerno uveden zavisan par (test B cita red koji ostavi test A) MORA biti prijavljen u obrnutom redosledu.

### T-10 Shard nivo 1: default suite u 2 paralelna Excela (bez izmene VBA testova) [trud M] [slabi: DELIMICNO, i to treba svesno prihvatiti: danas BFP, Agrohemija i Golden rade nad sveskom u kojoj su ostali tragovi RunAllTests (odnosno BFP-a); u shard-u svaka grupa krece od ciste kopije. Gubi se ta SLUCAJNA medjusuitna interakcija, koju nijedna imenovana tvrdnja ne meri; dobija se ponovljivost (isto pocetno stanje kao kod --suite poziva, koji dokaz vec koristi).]
Podeliti 12 default suita na dva run_vba procesa, npr. A = RunAllTests + Izvestaj + JSON + Banka + Faktura + StornoCentar + License + Storno + Palete; B = BFP + Agrohemija + Golden. Potrebno u alatu: (1) izvestaj po procesu umesto fiksnog tests/last_run.json/.txt (run_vba.py:1064-1123); (2) compile proba samo u jednom procesu ili jednom pre oba (trazi fokus VBE prozora i SendKeys, run_vba.py:530-635); (3) marker zelenog priznaje FULL tek kad unija shard-ova pokrije sve default suite nad ISTIM potpisom izvora i fixture-a. U testovima: fiksan folder %TEMP%\agrix_backup_test (modTest.bas:17763-17779) vezati za putanju sveske.
  DOBITAK: Pun podrazumevani prolaz sa > 600 s (ne staje u timeout) na ~max(18+216+~30, 18+BFP+~12) = oko 265-330 s, tj. 1,8-2,2x, i svaki shard staje u svoj timeout. Kraca Excel sesija smanjuje izlozenost neistrazenom padu 'posle ~10 min rada'.
  RIZIK: Srednji. Dve Excel instance na istoj radnoj povrsini: dijalozi i fokus (watchdog je po PID-u, to radi; SendKeys ne). Dvostruka rezija uvoza (2 x ~18 s CPU). Marker zelenog je mesto gde se lako napravi lazno zeleno -- pravilo unije mora imati dokaz.
  PREDUSLOVI: T-09 na nivou suita (svaka default suita zelena sama nad cistom kopijom). Izmene u run_vba/vba_gate sa svojim self-testom.
  DOKAZATI: Isti skup imena i verdikata kao jednoprocesni prolaz nad istim commit-om; marker: (a) jedan shard crven -> nema zelenog, (b) shard-ovi sa razlicitim potpisom izvora -> nema zelenog, (c) nedostaje suita -> nema zelenog; sve tri u vba_gate --self-test, crvene po imenu.

### T-11 Shard nivo 2: paralelni delovi UNUTAR RunAllTests/BFP i paralelni dokaz [trud L] [slabi: NE po sebi; isti sporedni gubitak kao T-10 (nema slucajnih interakcija izmedju oblasti u istoj svesci).]
Kad postoje suite po oblasti (T-08) i spisak zavisnosti (T-09): 3-4 run_vba procesa nad disjunktnim oblastima. Za dokaz: svaki proces radi u svom git worktree-u (mutacija izvora je u mestu; SRC_VBA je danas fiksan, run_vba.py:96), fixture se deli samo za citanje. Opcija za manju reziju: uvoz jednom, pa kopije pripremljene sveske po shard-u -- ali to trazi novu, dokazanu putanju u vba_gate jer se --no-import prolaz danas namerno ne priznaje kao dokaz.
  DOBITAK: FULL: ~4 shard-a x (18 s + ~130 s) u paraleli = oko 150 s zida umesto > 600 s (~4x). Dokaz: N worktree-a daje ~N puta kraci zid povrh dobitka iz T-08. Sve su izvedene procene.
  RIZIK: Visok: zavisi od tri prethodna koraka; vise Excel instanci trosi memoriju i CPU na razvojnoj masini; marker i grupna pravila dokaza postaju slozeniji.
  PREDUSLOVI: T-06, T-08, T-09, T-10.
  DOKAZATI: Unija rezultata = jednoprocesni FULL po imenima; za dokaz: isti verdikt DOKAZANO/NE po sabotazi kao pojedinacni prolaz, nad uzorkom od bar 30 sabotaza iz razlicitih oblasti; potpis izvora pre/posle identican u svakom worktree-u.

### T-12 Hermeticki RunOne (opciono): snimak pre testa i vracanje samo izmenjenih tabela [trud M] [slabi: DELIMICNO: gubi se slucajna integraciona pokrivenost 'kasniji test radi nad podacima ranijih' (nije imenovana tvrdnja). U BFP-u, gde je ostavljanje redova deo dizajna, NE primenjivati bez posebne odluke.]
Ako T-09 pokaze da je zavisnih testova mnogo: RunOne pre testa uzme snimak (Value2) svih tabela, a posle testa vrati samo one koje se razlikuju (poredjenje nizova u memoriji), uz ResetSeamova i reset keseva. Testovi koji menjaju SEMU (T_Sema_*, kolone/formati) i T_TxRollback_* ostaju izuzeti deklaracijom, jer snimak ne podnosi promenu broja kolona (clsTransaction.cls:269-272).
  DOBITAK: Svaki test krece od istog stanja -> redosled, izbor i shard postaju bezbedni bez pojedinacnih popravki; fixture prestaje da raste zbog 'redova koje niko drugi ne dira'. Tabele ostaju male, pa nestaje i rast cene po rednom broju testa.
  RIZIK: Srednji/visok: cena ~50 citanja tabela po testu (procena 5-30 ms na malom fixture-u, tj. 2-10 s po prolazu; nije mereno); testovi koji danas nesvesno zavise od prethodnog stanja pocece da padaju (to je i cilj, ali je posao). Sam mehanizam je nova kriticna tacka.
  PREDUSLOVI: T-01 (otisak), T-09 (broj zavisnih). Ne raditi pre nego sto brojevi pokazu da je potrebno.
  DOKAZATI: Dvosmerno: test koji ostavi red -> sledeci test vidi cistu tabelu; sabotaza mehanizma (vracanje preskoceno) mora oboriti namenski test po imenu. Skup verdikata isti u pravom i obrnutom redosledu.

### T-13 Vruca putanja pristupa podacima (PRODUKCIONI kod) -- tek ako merenje pokaze [trud M] [slabi: NE, ali menja produkciono ponasanje pod testom; svaki rez mora doneti svoj test i sabotazu (npr. 'kes prezivi promenu seme').]
Kandidati, svaki zaseban rez sa pre-flight-om: (a) kes ListObject-a po imenu u GetTable (modDataAccess.bas:118-131) sa ponistavanjem pri promeni seme; (b) AppendRow upisuje red jednim nizom umesto celiju po celiju (289-291); (c) StampRowAudit pamti indekse 4 audit kolone po tabeli umesto 4 petlje po imenu pri svakom upisu (382-423); (d) GetConfigValue cita config kroz kes sa ponistavanjem u SetConfigValue. Ovo NIJE izmena testova nego jezgra, i ubrzava i operatera.
  DOBITAK: Nepoznat dok T-01 ne da brojeve; hipoteza je da je ovo najveci pojedinacni udeo u 0,64 s/test jer svaki pisac desetinama puta skenira listove i cita cele tabele. Dobitak se prenosi na SVE suite i na svaki prolaz dokaza.
  RIZIK: Visok: jezgro kroz koje prolazi svaki upis i citanje. Kes koji zastari daje pogresne podatke bez greske (vec postoji postmortem o zapamcenoj nuli u kesu kolona, modDataAccess.bas:214-229). Trazi FULL + pun dokaz + rucni Compile po rezu.
  PREDUSLOVI: T-01 sa brojacima poziva (GetTable, GetTableData, AppendRow) u test rezimu; skill pre-flight; ne raditi u istom rezu sa izmenama testova.
  DOKAZATI: Vreme po testu pre/posle nad istim commit-om testova; novi testovi za ponistavanje kesa pokazani crveni svojom sabotazom; FULL i pun katalog dokaza zeleni.

### T-14 Mapa pokrivenosti po piscu i konsolidacija modTest <-> BFP [trud M] [slabi: NE ako se nijedna tvrdnja ne brise bez para sabotaza; DA ako bi se brisalo po slicnosti imena -- to se ne predlaze.]
Generisati (alatom, iz izvora) tabelu 'produkcioni _TX pisac -> koji testovi ga zovu, u kom modulu, koje sabotaze ga ciljaju'. Danas 22 pisca vuku oba modula, a ambalaza ima dve nezavisne fabrike scenarija (modTest: AmbSejRevers i srodne; BFP: OtkHeader/Pr3*). Pri podeli po oblasti (T-08) testovi iste oblasti iz oba modula idu u ISTI modul i dele jednu fabriku scenarija. Brisanje udvojene tvrdnje dozvoljeno samo kad obe imaju sabotazu koja obara istu granu.
  DOBITAK: Manje koda za odrzavanje (dve fabrike scenarija -> jedna po oblasti) i osnova za odluku gde ide NOV test; vidljivo je i 26 _TX pisaca koje ne referencira nijedan test modul. Vremenski dobitak mali i nemeren.
  RIZIK: Srednji: 'isti pisac' ne znaci 'ista tvrdnja' -- modTest cesto meri ekran/citaoca, BFP pisca. Brisanje bez dokaza bi oslabilo sistem.
  PREDUSLOVI: T-08 (ili bar manifest iz T-06).
  DOKAZATI: Za svaku uklonjenu tvrdnju: njena sabotaza i dalje obara preostalu tvrdnju PO IMENU. Ukupan broj crvenih u punom katalogu ostaje isti.

### T-15 Deljena forma po suiti umesto gradnje ljuske u svakom od 51 testa -- NE preporucuje se bez merenja [trud M] [slabi: DA: slabi tvrdnju 'svaki test polazi od sveze izgradjene forme' (izolacija UI stanja izmedju testova). Prihvatljivo samo za podskup testova koji formu samo citaju, i tek posle merenja iz T-01.]
Graditi frmOtkupUI jednom i izmedju testova vracati stanje (OtkupUI_Release + reset ekrana) umesto Unload/New u svakom testu. Navodi se jer je verovatno drugi po velicini potrosac u RunAllTests, ali je cena u izolaciji visoka.
  DOBITAK: Procena 20-30 s po prolazu RunAllTests (10-14%), ako se potvrdi izvedena cena od ~0,5-0,65 s po gradnji.
  RIZIK: Visok: testovi 236-243 i slicni mere bas gradnju, pad usred gradnje i ciscenje; zaostalo stanje forme je vec proizvodilo lazne padove (razlog postojanja CleanupPosleTesta, modTest.bas:776-791).
  PREDUSLOVI: T-01; odluka vlasnika.
  DOKAZATI: Isti verdikti u pravom i obrnutom redosledu; sabotaze UI oblasti i dalje crvene po imenu; izmeren dobitak >= 10% suite, inace odustati.


## NEPOZNANICE
[
 "Stvarna raspodela vremena po testu: na main-u ne postoji nikakvo merenje ispod nivoa suite; svi udeli (forma, Save, journal, GetTable, snimci) u ovom izvestaju su zakljuceni iz koda i oznaceni kao procene.",
 "Sadasnje trajanje BFP-a: poslednji zapis je 210 s pri 2417 tvrdnji (07.10.2026); danas ima 3014+ tvrdnji i 391 test, a zna se samo da ceo poziv prelazi 600 s.",
 "Da li Application.OnTime tajmeri (AutoSaveTick, StornoWarmTick) zaista opale izmedju dve suite pod run_vba automatizacijom i da li su uzrok neistrazenog 'RPC failed' posle ~10 min -- hipoteza izvedena iz koda (AUTOSAVE_MAX_AGE_SECONDS = 600, tajmeri se ne otkazuju), nije reprodukovana.",
 "Cena jedne fajl-operacije journal-a i loga na test masini (zavisi od antivirusa) i broj AppendRow/CommitTx po prolazu -- bez toga se dobitak T-03 ne moze kvantifikovati.",
 "Koliko testova STVARNO zavisi od redosleda: statika daje gornju granicu (121 modTest test pise bez rollbacka; 235 BFP koraka ostavlja redove) i ~12 mesta gde kod to sam priznaje; pravi broj daje tek T-09. Za BFP posebno: koliko tvrdnji nad deljenim sifarnicima meri apsolutan broj umesto delte.",
 "Da li granica od 65.534 linije po modulu vazi u instaliranoj verziji Excela (dokumentovana je za VB6/VBA; ovde nije proverena) i da li je Excel 32-bitni ili 64-bitni (bitno za 'Out of memory' pri kompajliranju projekta od 233.684 linije i ~16.000 razlicitih identifikatora).",
 "Koliko traje gradnja ljuske forme BEZ .Show u nevidljivom Excelu -- produkciono merenje (240 kontrola ~180 ms) je verovatno uzeto nad prikazanom formom.",
 "Trajanja Banke, Fakture, Izvestaja, StornoCentra, Agrohemije, JSON-a i License suite nisu zapisana nigde u docs/; pretpostavka da zajedno traju ~30 s nije potvrdjena.",
 "Da li je VBA deo grane B (RunAllTestsIzbor, BfpIzabran, vreme po testu) ispravan -- nikad nije izvrsen u Excelu, a grana je 149 commit-a iza main-a i sudara se u modTest, BFP i sabotaza.py.",
 "Da li test moduli NAMERNO idu u release i klijentu kroz self-update (3,74 MB, 35,8% koda); ako ne, njihovo izostavljanje je zasebna odluka van ove oblasti.",
 "Koliko otvorenih grana trenutno dira modTest.bas i modBusinessFlowProTests.bas -- od toga zavisi kada je bezbedno uraditi T-06 i T-08.",
 "Zasto su komentari o redosledu ostali sa starim brojevima (plitka istorija ne pokazuje kada su testovi prenumerisani); moguce je da postoje i druge zastarele reference na brojeve testova u docs/."
]
