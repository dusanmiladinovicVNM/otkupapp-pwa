# Praksa i istrazivanje: kako se trosak mutacionog dokaza, izbora testova, Excel/COM harness-a, piramide verifikacije i odrzavanja registara resava drugde -- primenjeno na OtkupApp (katalog od 1265 rucno pisanih sabotaza, suite 80-150 s, rezija ~18 s, jedan Windows racunar)

## SAZETAK
Web: WebSearch je radio, WebFetch NIJE (pitest.org i stryker-mutator.io se ne razresavaju -- treba ih dozvoliti u mreznim podesavanjima sesije), pa su svi spoljni izvori citani kao sazeci pretrage, ne kao otvorene stranice, i nose izmereno=false. Lokalno sam merio (samo citanje, oba stabla ostala cista): katalog je po prirodi 'fault-injection' katalog, ne klasicno mutaciono testiranje -- 1265 sabotaza nosi 1264 razlicita kljuca (test, tvrdnja), pa tehnike koje smanjuju BROJ mutanata (selektivna mutacija, uzorkovanje, subsumpcija, predictive mutation testing) ovde ili ne vaze ili direktno ukidaju tvrdnje; primenljive su one koje smanjuju CENU PO MUTANTU i broj ponovljenih merenja. Najveca izmerena poluga je inkrementalni registar dokaza (analog PIT history / Stryker --incremental): sa kljucem na nivou procedure po PR-u zastari medijana 1,8% kataloga (prosek 3,2%, max 12,7%; 40 poslednjih PR-ova), a sa kljucem na nivou fajla 77% -- jer se modTest/BFP diraju u 28/25 od 40 PR-ova. Druga poluga je cena prolaza: svaki prolaz ponovo uvozi svih 194 fajla (10,5 MB) i placa 4,5-7,9 s fiksnih sleep-ova, a sabotaza dira jedan modul od prosecno 170 KB (1,6% izvora); 1258/1265 sabotaza je nad .bas modulima, pa je 'topla' Excel sesija sa zamenom jednog modula izvodljiva istim COM pozivima koji se vec koriste. Izbor testa (grana B) plus topla sesija po proceni (ne merenju) spustaju pun POJEDINACAN dokaz sa ~43 h na ~1-2,5 h -- ispod danasnjeg grupnog (~7,3 h), sto dozvoljava da se slabiji verdikt 'DOKAZANO (grupno)' penzionise. Izbor i registar slabe tvrdnju na tacno dva imenovana nacina (izolovan test vs kanonski redosled; direktni ulazi vs ceo izvor); oba se zatvaraju izmerenom nezavisnoscu testova (PolDet analog) i revizijom uzorkom (0 neslaganja u 100 -> najvise 2,95% uz 95%). VBA se van Excela ne moze verno izvrsiti, Microsoft ne podrzava nenadgledanu automatizaciju Office-a, pa se projektuje za recikliranje instanci, ne za jednu dugovecnu. Merge queue na GitHub-u nije dostupan repou u vlasnistvu korisnickog naloga; piramida se zato gradi oko registra dokaza koji CI (Linux) samo proverava. Tri fajla-zarista (sabotaza.py 35/40 PR-ova, modTest 28/40, BFP 25/40) su i uzrok sudara nespojene grane B; lek iz prakse su fragmenti kataloga po modulu i generisan registar testova.

## CINJENICE
- [M] Katalog nije klasican skup mutanata: 1265 sabotaza, 551 razlicit test (318 RunAllTests + 227 BFP + 6 banka), 1264 razlicita kljuca (suita, test, tvrdnja) -- samo JEDAN par sabotaza deli istu tvrdnju. Sabotaza po testu: medijana 2, prosek 2,30, max 14 (T_FakturaGP_WriterKapijeIStorno).
    DOKAZ: python3 <scratch>/praksa/stat.py <main-wt> -> 'ukupno 1265 / razlicitih testova 551 / razlicitih (test,tvrdnja) 1264 / visak (redundantno po kljucu): 1 / min 1 med 2.0 max 14 mean 2.30'
- [M] 1258 od 1265 sabotaza gadja .bas module (75 fajlova), 6 jednu .cls, 1 jednu .doccls; nijedna ne gadja test module ni formu. Prosecan sabotiran modul (ponderisano brojem sabotaza) ima 170 KB = 1,6% izvora (src-vba = 10,46 MB, 194 fajla). 388 od 1265 ima viselinijsko sidro ili zamenu.
    DOKAZ: python3 skripta nad <scratch>/main-wt/tools/sabotaza.py: 'sabotaza po ekstenziji: {.bas: 1258, .cls: 6, .doccls: 1}; sabotaze nad test modulima: {}; prosecna velicina 170 KB; udeo 1.6%; viselinijskih 388 od 1265'
- [M] Svaki poziv run_vba.py pravi NOVU Excel instancu, uvozi SVE komponente (Remove + Import za svaki .bas/.cls/.frm, DeleteLines + AddFromString za .doccls), pokusava compile, zove EnsureRuntimeSchema, pa suite. Nema parametra za drugi izvorni folder (SRC_VBA je konstanta).
    DOKAZ: <scratch>/main-wt/tools/run_vba.py:96 (SRC_VBA), :353-406 (import_src_vba), :875 (DispatchEx), :896, :903, :938
- [M] run_vba.py nosi fiksne sleep-ove po pozivu: 2,5 s po pokusaju compile-a (1 ili 2 pokusaja), jos 0,5+0,4 s na SendKeys putu, i 1,0+1,0 s na gasenju -- dakle najmanje 4,5 s, a 7,9 s kad ide i SendKeys pokusaj (sto je slucaj kad je verdikt NEJASNO). To je 25-44% dokumentovane rezije od ~18 s.
    DOKAZ: <scratch>/main-wt/tools/run_vba.py:617, :632, :634, :1012, :1021 (citanje koda; zidno vreme NIJE mereno jer nema Excela)
- [M] dokaz.py po prolazu pokrece po jedan podproces sabotaza.py za svakog clana, jedan run_vba.py --suite (CELA suita) i jedan sabotaza.py --vrati; mutacija se upisuje u RADNO stablo src-vba i posle se poredi otisak celog src-vba. Timeout podprocesa je 1200 s, a tvrdi prekid run_vba 600 s.
    DOKAZ: <scratch>/main-wt/tools/dokaz.py:92, :325, :349, :358-365; run_vba.py:705
- [M] dokaz.py pre prve mutacije uvek pusta RunAllTests + svaku suitu koju izbor trazi (1-3 pune suite) kao bazu, i NE ostavlja nikakav trajan zapis ishoda (samo print). Marker u vba_gate.py je vezan za otisak CELOG src-vba i pamti rezultat po suiti, ne po sabotazi.
    DOKAZ: <scratch>/main-wt/tools/dokaz.py:466-473; grep 'json.dump|open(.*"w"' tools/dokaz.py -> bez pogodaka; tools/vba_gate.py:43-48, :424
- [M] INKREMENTALNI KLJUC, mereno nad 40 poslednjih PR-ova (first-parent, 19.10-08.11.2026): ako se sabotaza ponovo dokazuje samo kad se promeni njena n-torka, procedura u kojoj je sidro ili procedura njenog testa, po PR-u zastari medijana 1,8% kataloga (prosek 3,2%, max 12,7%; u komadima: medijana 20,5, prosek 32,5, max 138). Zbir preko 40 PR-ova: 1298 ponovnih dokaza naspram 42112 kad bi se svaki put dokazivao ceo katalog (3,1%).
    DOKAZ: python3 <scratch>/praksa/inkrement.py <main-wt> 40 -> 'K_proc + nove + izmenjene, % kataloga po PR-u: medijana 1.8, prosek 3.2, max 12.7; zbir 1298 ; zbir velicina kataloga 42112' (cita samo git objekte)
- [M] Isti racun sa grubljim kljucevima: produkcija na nivou FAJLA + test na nivou procedure -> medijana 12,3% (prosek 15,0%, max 50,7%); oba na nivou fajla -> medijana 77,1% (prosek 68,8%, max 99,8%). Kljuc na nivou fajla je dakle skoro bezvredan jer se test moduli menjaju u vecini PR-ova.
    DOKAZ: python3 <scratch>/praksa/inkrement2.py <main-wt> 40 -> 'K_mid ... medijana 12.3, prosek 15.0, max 50.7'; 'K_fajl ... medijana 77.1, prosek 68.8, max 99.8'
- [M] Tri fajla su zarista sudara: tools/sabotaza.py je diran u 35 od 40 poslednjih PR-ova, src-vba/modTest.bas u 28, src-vba/modBusinessFlowProTests.bas u 25 (docs/STANJE_REFAKTORA.md u 33). U tih 40 PR-ova modTest je dobio +20.536/-10.214 redova, BFP +19.216/-9.870, sabotaza.py +4.273/-162. Katalog je u istom prozoru porastao 811 -> 1265 (+454, ~11 po PR-u).
    DOKAZ: petlja 'git diff --quiet C^1 C -- <fajl>' nad 'git log --first-parent -n 40' u main-wt; 'git diff --numstat <41. commit> HEAD'; kolona 'kat' u izlazu inkrement.py
- [M] dokaz.py --grupe 6 --plan nad celim katalogom traje 20,8 s cistog Pythona; 20,3 s od toga je sabotaza._suita_testa pozvana 1265 puta, koja svaki put ponovo cita (8,3 s) i regex-om skenira (11,9 s) test module od po ~1,5 MB. Plan: 134 grupe RunAllTests + 78 BFP + 3 banka = 215.
    DOKAZ: time python3 tools/dokaz.py --grupe 6 --plan -> real 0m20.829s; cProfile: '_suita_testa 1265 poziva 20.325 s; findall 11.878 s; _procitaj 8.259 s'; <scratch>/praksa/plan6.txt
- [M] Katalog pokriva manji deo tvrdnji: modTest ima 3386 AssertEq naredbi (na pocetku reda) naspram 801 sabotaze (~24%), BFP ~3200 Assert* poziva naspram 456 (~14%). Ako politika ostane 'nova tvrdnja dobija sabotazu', katalog ima prostora da poraste 4-7 puta.
    DOKAZ: grep -cE '^\s*Assert' src-vba/modTest.bas -> 3386; BFP: AssertTrue 1740 + AssertEquals 1395 + AssertFalse 45 + AssertDoubleNear 22
- [N] SendKeys rezerva za compile aktivira VBE prozor PO NASLOVU (AppActivate(caption)) pa salje Alt+D, C. Dve paralelne Excel instance sa istom sveskom imale bi isti naslov -- tastatura moze da ode u pogresnu instancu. (Sudar je zakljucak iz koda, nije izveden.)
    DOKAZ: <scratch>/main-wt/tools/run_vba.py:625-634
- [M] Dokumentacija tvrdi da hook 'radi u milisekundama', a izmereno je 9-12 s po izmeni (podatak iz zadatka) -- dokument je zastareo u odnosu na alat.
    DOKAZ: <scratch>/main-wt/docs/EXCEL_TEST_HARNESS.md:91
- [M] Grana B vec sadrzi izbor testa (run_vba --test, dokaz --izbor sa bazom nad ISTIM izborom, --izolacija, vreme po testu) i sama kaze da VBA strana nikad nije izvrsena; sudara se sa main-om bas u tri fajla-zarista.
    DOKAZ: git -C /home/user/otkupapp-pwa diff HEAD~1 HEAD -- docs/EXCEL_TEST_HARNESS.md (odeljak 'Izbor testova i vreme po testu', pasus 'Sta je izmereno, a sta nije'); /home/user/otkupapp-pwa/tools/dokaz.py:532-538
- [N] PIT inkrementalna analiza: ishod mutanta se preuzima iz istorije ako se nisu promenili ni klasa pod testom ni test koji ga je ubio; PIT sam navodi da pociva na pretpostavci da promene u ZAVISNOSTIMA retko menjaju status, i da je ta pretpostavka 'trenutno nedokazana'.
    DOKAZ: https://pitest.org/quickstart/incremental_analysis -- WebSearch sazetak; stranica nije otvorena (WebFetch: ENOTFOUND pitest.org)
- [N] Stryker --incremental cuva rezultate u JSON-u i ponovo pusta samo izmenjene mutante; ne vidi izmene u fajlovima koji nisu ni mutirani ni test fajlovi; --force ponistava preuzimanje.
    DOKAZ: https://stryker-mutator.io/docs/stryker-js/incremental/ -- WebSearch sazetak; stranica nije otvorena
- [N] Google mutaciono testira samo izmenjen kod u code review-u, najvise jedan mutant po liniji, uz agresivno potiskivanje neproduktivnih mutanata; nikad ceo kod pri svakoj izmeni.
    DOKAZ: https://arxiv.org/abs/2102.11378 (Petrovic, Ivankovic, Fraser, Just, 'Practical Mutation Testing at Scale', TSE 2021); https://research.google/pubs/state-of-mutation-testing-at-google/ -- WebSearch sazeci
- [N] Mutant schemata (Untch/Offutt/Harrold 1993): svi mutanti u jednom programu, kompajlira se jednom, mutant se bira u run-time-u; prijavljena ubrzanja ~5x (zavise od odnosa kompajliranje/izvrsavanje). Stryker koristi sandbox kopiju projekta (.stryker-tmp) -- ne mutira radno stablo.
    DOKAZ: https://www.arxiv.org/pdf/2210.17215 (Clang studija, 4.87-5.16x); https://github.com/stryker-mutator/stryker-js -- WebSearch sazeci; 'mutation switching' kod Stryker-a NIJE potvrdjen pretragom
- [N] Uzorkovanje mutanata: konstantan uzorak (~1000) daje statisticki slican MUTACIONI SKOR kao pun skup (Gopinath i dr., ISSRE 2015). To je procena udela, ne provera svakog unosa.
    DOKAZ: https://www.cefns.nau.edu/~adg326/issre15.pdf -- WebSearch sazetak; tacna granica greske nije proverena u PDF-u
- [N] Predictive mutation testing obecava do 151x ubrzanja, ali kasniji rad navodi da postojeci pristupi pogresno klasifikuju ~33% ishoda.
    DOKAZ: https://ar5iv.labs.arxiv.org/html/2309.02389 (Contextual PMT); Zhang i dr. ISSTA 2016 -- WebSearch sazeci
- [N] Redundantni (subsumed) mutanti naduvavaju skor; TCE odbacuje ~7% mutanata kao ekvivalentne i ~21% kao duplikate (C programi). Tehnika trazi izlaz kompajlera, kog za VBA nema.
    DOKAZ: https://orbilu.uni.lu/bitstream/10993/20289/1/ICSE15.pdf (Papadakis i dr., ICSE 2015); Ammann/Delamaro/Offutt ICST 2014 -- WebSearch sazeci
- [N] PIT: timeoutFactor je podrazumevano 1,25 (uz konstantu; 4000 ms vidjeno samo u primeru trece strane); mutant koji istekne broji se kao UBIJEN. Infection ima opciju da istek broji kao PREZIVELOG. OtkupApp vec radi strozu varijantu (TIMEOUT = nije dokazano).
    DOKAZ: https://pitest.org/quickstart/maven ; https://infection.github.io/2026/01/14/whats-new-in-0.32.3 -- WebSearch sazeci; dokaz.py:377-378 u main-wt
- [N] Regression test selection: Ekstazi belezi DINAMICKE zavisnosti testa i poredi checksum-e; Microsoft TIA bira pogodjene + prethodno pale + nove testove, PADA NA 'pusti sve' kad izmenu ne ume da protumaci, i ima podesivo periodicno 'pusti sve'.
    DOKAZ: https://learn.microsoft.com/en-us/vsts/build-release/test/test-impact-analysis ; Gligoric i dr. ISSTA 2015 (Ekstazi) -- WebSearch sazeci
- [N] Zavisnost od redosleda: zavisni testovi nadjeni u svakoj proucenoj suiti (Zhang i dr. ISSTA 2014); iDFlakies skup: 422 flaky testa, 50,5% zavisno od redosleda; podela victim/polluter/brittle; PolDet poredi snimak deljenog stanja pre i posle svakog testa; 'dependent-test-aware' izbor daje 65% manje laznih padova uz ~1% sporiji rad (Lam i dr. ISSTA 2020).
    DOKAZ: https://cs.gmu.edu/~winglam/publications/2019/LamETAL19iDFlakies.pdf ; https://mir.cs.illinois.edu/gyori/pubs/issta15pollution.pdf ; https://cs.gmu.edu/~winglam/publications/2020/LamETAL20ISSTA.pdf -- WebSearch sazeci
- [N] 'Can't enter break mode at this time' je dokumentovana posledica PROGRAMSKE izmene projekta kroz extensibility model (izvrsavanje se moze nastaviti ili prekinuti, ne i pauzirati). Programska izmena koda (AddFromString, DeleteLines, InsertLines...) resetuje sve promenljive projekta.
    DOKAZ: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/can-t-enter-break-mode-at-this-time ; MS KB 198637 (Access) i https://learn.microsoft.com/en-us/answers/questions/1183083 -- WebSearch sazeci; za Excel konkretno NIJE provereno
- [N] Microsoft ne preporucuje i ne podrzava nenadgledanu automatizaciju Office-a (STA, nije reentrantan, dijalozi 'zamrznu' proces); servis u sesiji 0 nema korisnicki interfejs. DispatchEx pravi novu Excel instancu (proces) po pozivu.
    DOKAZ: https://support.microsoft.com/help/257757 ; https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa ; pywin32 mailing lista -- WebSearch sazeci
- [N] Nema vernog nacina da se VBA izvrsi van Office hosta: Rubberduck parsira ali ne izvrsava (testovi idu iz VBE-a, bez headless CLI-ja); LibreOffice 'Option VBASupport 1' pokriva deo jezika, a Excel objektni model nema; twinBASIC je u beti i ne podrzava VBA UserForm-e. vba-blocks/vbapm grade svesku iz izvora, ali i dalje traze Excel.
    DOKAZ: https://rubberduckvba.blog/2017/10/19/how-to-unit-test-vba-code/ ; https://ask.libreoffice.org/t/how-does-one-get-complete-compatibility-with-vba/55876 ; https://github.com/GCuser99/VBA-UserForm-to-twinBASIC ; https://github.com/vba-blocks/vba-blocks -- WebSearch sazeci
- [N] Budzeti iz prakse: 'commit build' do ~10 minuta uz sekundarne (etapne) build-ove (Fowler); pre-commit do ~10 s i pre-push do ~2 min su konvencije praktikanata, ne standard; Google velicine testova 60/300/900 s (sekundarni izvor).
    DOKAZ: https://martinfowler.com/bliki/ContinuousIntegrationCertification.html ; https://www.thoughtworks.com/insights/blog/pre-commit-don-t-git-hooked ; https://testing.googleblog.com/2010/12/test-sizes.html -- WebSearch sazeci
- [N] GitHub merge queue postoji samo za javne repoe u vlasnistvu ORGANIZACIJE i za privatne na Enterprise Cloud-u; trazi merge_group okidac u workflow-u. Repo je, po remote URL-u, pod korisnickim nalogom -- pa najverovatnije nije dostupan (vrsta naloga nije proverena).
    DOKAZ: https://docs.github.com/en/pull-requests/how-tos/merge-and-close-pull-requests/merging-a-pull-request-with-a-merge-queue -- WebSearch sazetak; 'git remote -v' u main-wt
- [N] Za registre koji se samo dopunjuju praksa nudi dva leka: merge=union u .gitattributes (jeftino, ali ga hosting ne postuje uvek i moze tiho spojiti sustinski sukob) i 'jedan fragment-fajl po izmeni' sa generisanjem celine (towncrier obrazac). Za registar testova: generisan runner po konvenciji imena (Unity generate_test_runner.rb; Rubberduck '@TestMethod).
    DOKAZ: https://about.gitlab.com/blog/gitlab-reduced-merge-conflicts-by-90-percent-with-changelog-placeholders/ ; https://www.Github.com/orgs/community/discussions/9288 ; Unity UnityHelperScriptsGuide.md -- WebSearch sazeci

## TROSKOVI
[
 {
  "sta": "Pun pojedinacan dokaz celog kataloga, danasnjim putem (cela suita po sabotazi)",
  "koliko": "~43 h: 801 x (18+80) s + 456 x (18+150) s + 8 x ~48 s. Procena iz dokumentovanih brojeva, nije mereno.",
  "izmereno": false,
  "koliko_cesto": "pred release (po CLAUDE.md); u praksi se zato ne pusta"
 },
 {
  "sta": "Pun grupni dokaz (--grupe 6), bez ponavljanja clanova",
  "koliko": "~7,3 h: 134 x 98 s + 78 x 168 s + 3 x ~48 s (215 prolaza izmereno planom; trajanja iz dokumentacije)",
  "izmereno": false,
  "koliko_cesto": "pred release"
 },
 {
  "sta": "Rezija jednog run_vba poziva: nova instanca + uvoz 194 fajla / 10,5 MB + compile pokusaj + gasenje",
  "koliko": "~18 s po dokumentaciji; od toga 4,5-7,9 s su fiksni sleep-ovi (citano iz koda), a uvozi se 100% izvora iako sabotaza dira ~1,6%",
  "izmereno": false,
  "koliko_cesto": "svaki prolaz dokaza i svaki TARGETED/FULL poziv"
 },
 {
  "sta": "Baza pre prve mutacije u dokaz.py (RunAllTests uvek + suite iz izbora)",
  "koliko": "1-3 pune suite: ~1,6 min (samo RunAllTests) do ~4,4 min (RunAllTests + BFP). Procena.",
  "izmereno": false,
  "koliko_cesto": "svaki poziv dokaz.py, i za rez od 2 sabotaze"
 },
 {
  "sta": "Koliko sabotaza STVARNO trazi ponovni dokaz po PR-u (nove + izmenjene n-torke + promenjena procedura sidra ili testa)",
  "koliko": "medijana 20,5, prosek 32,5, max 138 komada (1,8% / 3,2% / 12,7% kataloga)",
  "izmereno": true,
  "koliko_cesto": "po PR-u (40 poslednjih)"
 },
 {
  "sta": "Pokretanje dokaz.py nad celim katalogom pre ijednog Excel poziva (_suita_testa bez memoizacije)",
  "koliko": "20,8 s",
  "izmereno": true,
  "koliko_cesto": "svaki pun poziv i svaki --plan; kod reza srazmerno broju izabranih (~17 ms po sabotazi)"
 },
 {
  "sta": "Sabotaza koja obori compile ili otvori dijalog",
  "koliko": "do 600 s po dogadjaju (tvrdi prekid run_vba), u grupi jos i ponavljanje svih clanova pojedinacno",
  "izmereno": false,
  "koliko_cesto": "nepoznato -- u zdravom katalogu 0, ali svaka nova sabotaza je kandidat"
 },
 {
  "sta": "Odrzavanje tri fajla-zarista (sudari, rebase, citanje diff-a)",
  "koliko": "sabotaza.py diran u 35/40 PR-ova, modTest 28/40, BFP 25/40; ~30.000 obrisanih/prepisanih redova u dva test modula za 40 PR-ova; grana B (149 commit-a iza) se sudara bas tu",
  "izmereno": true,
  "koliko_cesto": "skoro svaki PR"
 },
 {
  "sta": "Rast kataloga",
  "koliko": "+454 sabotaze u 40 PR-ova (~11 po PR-u); pun dokaz raste linearno s tim",
  "izmereno": true,
  "koliko_cesto": "trajno"
 },
 {
  "sta": "Revizija uzorkom u punom 'hladnom' rezimu (cela suita, pojedinacno)",
  "koliko": "n=59 -> ~2,0 h; n=100 -> ~3,4 h (prosek ~123 s po sabotazi; procena). U toplom rezimu sa izborom testa: n=100 -> ~10 min (procena).",
  "izmereno": false,
  "koliko_cesto": "po release-u, ako se usvoje P1/P2"
 }
]

## PREDLOZI
### P1 Inkrementalni registar dokaza po sabotazi (analog PIT history / Stryker --incremental) [trud M] [slabi: DA. Danas: 'svaka sabotaza je nad OVIM izvorom oborila svoju tvrdnju'. Posle: 'svaka je oborila svoju tvrdnju nad izvorom u kome su njeni direktni ulazi identicni danasnjim'. Ostatak rizika se iskazuje brojem kroz P6 (0 nalaza u 100 -> najvise 2,95% zastarelih uz 95%).]
dokaz.py upisuje za svaku sabotazu zapis {verdikt, kljuc, datum, commit, rezim}. Kljuc = hes n-torke + hes procedure u kojoj je sidro + hes procedure testa + otisak 'ugovora' harness-a koji vba_gate vec racuna (run_vba.py, make_fixture, golden) + otisak motora sabotaza.py/dokaz.py. Ponovo se meri samo ono ciji se kljuc promenio. Za release se koristi stroziji kljuc (ceo sabotirani FAJL + procedura testa). Verdikt nosi oznaku rezima, isto kao danasnje '(grupno)': 'DOKAZANO (inkrementalno, revizija n=...)'. Registar je prirodan nastavak markera u vba_gate, samo na zrnu sabotaze umesto suite.
  DOBITAK: Izmereno nad 40 PR-ova: po PR-u se ponovo meri medijana 1,8% kataloga (20 komada) umesto 100%; ukupno 1298 naspram 42112 merenja (-97%). Sa strozim kljucem 12,3% medijana (i dalje -85%). Release dokaz prestaje da raste linearno sa katalogom.
  RIZIK: Isti koji PIT sam priznaje: promena u ZAVISNOSTI (pozvana procedura, konstanta u modConfig, fixture, zajednicki seed) moze da promeni ishod a da kljuc ostane isti. Ublazenja: strozi kljuc za release, revizija uzorkom (P6), pun dokaz kad se promeni ugovor harness-a, starenje zapisa (svaki se ponovo izmeri najkasnije na X dana kroz nocni prolaz).
  PREDUSLOVI: Nista sa Windows strane za sam registar (cist Python + self-test). Za smisleno starenje: nocni prolaz (P9).
  DOKAZATI: (1) Self-test kljuca u oba smera: izmena u proceduri sidra ili testa -> zapis zastareva; izmena komentara u trecoj proceduri -> ne zastareva. (2) Ponoviti inkrement.py nad istorijom kao regresioni test procene. (3) Jedan pun dokaz uporedjen sa registrom: broj neslaganja JE izmerena stopa zastarelosti i ulazi u dokument.

### P2 Izbor testa po sabotazi kao podrazumevani put dokaza (preneti granu B na main) [trud M] [slabi: DA, usko. Izbor dokazuje 'test je SAM osetljiv na mutaciju', a ne 'RunAllTests u kanonskom redosledu je osetljiv': raniji test moze da ostavi stanje koje mutaciju MASKIRA u punoj suiti. Zatvara se sa P5 (izmerena nezavisnost) i P6 (revizija uzorkom u punoj suiti).]
Svaka sabotaza vec imenuje svoj test, pa je 'coverage-based test selection' ovde besplatan: pusta se samo taj test, sa bazom nad ISTIM izborom (grana B to vec radi: run_vba --test, dokaz --izbor, --izolacija, tražen-a-nepusten test = FAIL). Preporuka je da se logika prenese na svez main u malim rezovima (prvo Python strana sa self-testom, pa VBA ulazne tacke), a ne da se rebase-uje 149 commit-a preko tri fajla-zarista.
  DOBITAK: Prolaz 98-168 s -> ~20-23 s (prosecan test je 80/350 = 0,23 s, odnosno 150/391 = 0,38 s; procena). Pun grupni dokaz ~7,3 h -> ~1,4-2,9 h; pojedinacan ~43 h -> ~7 h. Posle ovoga rezija od 18 s postaje ~90% cene prolaza -- zato P3.
  RIZIK: VBA strana grane nikad nije izvrsena. BFP 'seed se pusta uvek' moze da pojede veci deo dobitka u toj suiti (nepoznato). Test zavisan od redosleda: grana B ga vec vraca na celu suitu, ali ne vidi obrnut slucaj (v. slabljenje).
  PREDUSLOVI: Windows prolaz 'dokaz.py --izolacija' nad rezom pa nad svim; vreme po testu (isti commit).
  DOKAZATI: Na uzorku od ~50 sabotaza uporediti verdikt --izbor sa verdiktom cele suite (moraju biti isti); sabotaze izbor-* / bfp-izbor-* iz grane B moraju pokazati crveno; pogresno ime testa mora dati FAIL, ne zelen prolaz sa nula testova.

### P3 Topla Excel sesija: jedan uvoz, pa zamena samo sabotiranog modula ('mutation worker') [trud L] [slabi: NE -- uz uslov povratka na zeleno po mutaciji. To je cak jace od danasnjeg: danas se baza tvrdi jednom na pocetku, a ovde se za SVAKU sabotazu meri crveno sa mutacijom i zeleno posle vracanja, u istoj sesiji.]
Jedna instanca i jedan pun uvoz po seriji. Po sabotazi: mutacija se primeni NA TEKST modula u memoriji ili temp folderu (nikad u radno stablo), zameni se samo ta komponenta (Remove + Import za .bas/.cls, DeleteLines + AddFromString za .doccls -- isti pozivi koje run_vba vec koristi 194 puta po pozivu), pusti se ciljani test, vrati se originalni modul, i ISTI test se pusti jos jednom i mora biti zelen. Posle svake zamene se tvrdi ime i broj komponenti (poznata zamka: odlozen Remove pa 'Modul1'). Zdravlje sesije se proverava izmedju mutacija (probe vraca 42, watchdog nije video dijalog, fajl rezultata postoji); na svaku anomaliju i na svakih K mutacija instanca se ubija i dize nova nad 'ispecenom' sveskom (uvoz uradjen jednom i snimljen u temp).
  DOBITAK: Nestaje rezija ~18 s po prolazu (ukljucujuci 4,5-7,9 s sleep-ova) i uvoz pada sa 100% na ~1,6% izvora po mutaciji. Uz P2: pun POJEDINACAN dokaz ~1265 x (2 + 1..5 s) = ~1,1-2,5 h (procena) -- manje od danasnjeg grupnog, pa se '--grupe' i slabiji verdikt 'DOKAZANO (grupno)' mogu povuci iz release puta. Sporedno: nestaje klasa REVERT-FAIL, razvoj moze da tece dok dokaz radi, i otvara se P4.
  RIZIK: Curenje stanja izmedju mutacija (podaci u svesci ako mutiran pisac pokvari rollback; stanje modula se po MS dokumentaciji resetuje pri programskoj izmeni koda, sto ovde pomaze, ali za Excel nije provereno). [break] ili dijalog zatruje sesiju -> zato recikliranje. Dugoveka COM sesija je scenario koji Microsoft ne podrzava. Nepoznato kako VBE podnosi stotine zamena modula u jednoj sesiji.
  PREDUSLOVI: P2 (ulazna tacka za jedan test). Run_vba mora da primi izvor modula kao parametar (danas je SRC_VBA konstanta).
  DOKAZATI: (1) A/B: ista ~50 sabotaza hladnim i toplim putem -- verdikti identicni. (2) Namerno 'curava' pseudo-sabotaza (ostavi red u tabeli) -> povratak-na-zeleno je mora uhvatiti po imenu. (3) Orkestracija bez Excela, laznim COM objektom, kao sto dokaz.py --self-test vec radi za _prolaz. (4) Ubiti Excel usred serije -> preostale stavke se izmere u novoj instanci, nijedna ne ispadne iz imenioca.

### P4 Paralelni Excel radnici, podela po izmerenom vremenu [trud M] [slabi: NE -- isto merenje po sabotazi, samo u vise procesa.]
N procesa, svaki sa svojom DispatchEx instancom, svojom temp kopijom sveske i svojim redom sabotaza; podela reda po vremenu testa iz prethodnog prolaza (praksa 'split by timings'). U radnicima se compile korak i SendKeys rezerva NE izvrsavaju (AppActivate po naslovu prozora nije bezbedan uz vise instanci). Radnik koji padne vraca svoj red u zajednicki.
  DOBITAK: Priblizno N puta krace do broja fizickih jezgara; realno 2-4x na razvojnoj masini (procena, nije mereno). Posle P2+P3 pun dokaz bi stao u ~20-60 min.
  RIZIK: Fokus prozora i clipboard su deljeni; anegdotski izvestaji o 'RPC server unavailable' uz vise Excel sesija; memorija po instanci nepoznata; masina je za to vreme tesko upotrebljiva za rad.
  PREDUSLOVI: P3 (mutacija van radnog stabla) -- bez toga paralelni radnici pisu u isti src-vba. Vreme po testu (P2).
  DOKAZATI: Isti skup sabotaza serijski i paralelno -> identican skup verdikata; ubijen jedan radnik -> ostali nedirnuti, njegov red ponovo izmeren; dve instance istovremeno ne smeju da daju nijedan dijalog koji watchdog jednog pripise drugom (watchdog je vec vezan za PID).

### P5 Nezavisnost testova kao IZMERENA osobina (analog PolDet + iDFlakies) [trud M] [slabi: NE -- dodaje tvrdnju koja danas ne postoji ('testovi ne ostavljaju stanje') i time zatvara slabljenje iz P2.]
Dijagnosticki rezim suite: pre i posle svakog testa uzima se otisak deljenog stanja (broj redova i kontrolni zbir svake tabele, Application zastavice, stanje seam-ova, broj otvorenih svesaka, fajlovi u temp folderu). Test koji promeni otisak je 'polluter' i imenuje se. Dozvoljeni trajni efekti (log listovi) idu u registar sa razlogom, po obrascu koji repo vec koristi (MRTAV_UNOS). Uz to nocno: --izolacija (lista testova koji sami ne prolaze = 'brittle') i jedan prolaz u OBRNUTOM redosledu.
  DOBITAK: Cini P2 ispravnim po konstrukciji: ako nijedan test ne ostavlja stanje, izolovan test i test u suiti mere isto. Pravilo 'prvo proveri izolaciju, pa proizvod' iz dokumentacije postaje provera. Trosak po proceni: nekoliko sekundi po suiti u dijagnostickom rezimu + jedna dodatna suita nocno.
  RIZIK: Cena otiska nad velikim tabelama u VBA (fixture je sintetican, ali nije mereno); lazni nalazi dok se registar dozvoljenih efekata ne ustali; obrnut redosled ce verovatno prvi put otkriti stvarne zavisnosti koje treba leciti (state-setter obrazac) -- to je posao, ne trosak alata.
  PREDUSLOVI: Windows; ulazna tacka koja prima redosled (deli mehanizam sa izborom iz P2).
  DOKAZATI: Sabotaza nad samim mehanizmom: test koji namerno ostavi red u tabeli mora biti imenovan kao polluter; uklonjen -> zeleno. Test koji zavisi od prethodnog mora pasti u --izolacija po imenu.

### P6 Revizija uzorkom umesto punog ponavljanja -- sa iskazanom granicom [trud S] [slabi: Ovo JE cena P1 i P2, iskazana brojem: umesto 'sve je ponovo izmereno' -- 'sve je izmereno bar jednom nad svojim ulazima, a udeo zastarelih je uz 95% ispod p'. Literatura o uzorkovanju mutanata (Gopinath 2015) opravdava procenu UDELA, ne proveru svakog unosa -- zato uzorak sluzi samo za reviziju, nikad kao zamena za prvi dokaz nove sabotaze.]
Dve revizije po release-u, obe nad slucajnim uzorkom stratifikovanim po suiti, sa semenom iz commit hesa (ponovljivo): (a) n zapisa iz registra koji NISU ponovo mereni -> ponovo se izmere (jeftino, toplim putem); (b) n sabotaza dokazanih izolovanim testom -> izmere se u punoj suiti, hladno, pojedinacno. Rezultat se pise uz verdikt: '0 neslaganja u n'. Svako neslaganje je nalaz i povecava n za sledeci put.
  DOBITAK: Zamenjuje 7-43 h punog ponavljanja sa ~10 min (a, toplo) + ~2-3,4 h (b, n=59-100, hladno; nocno). Tacna binomna granica: 0 u 59 -> najvise 4,95%; 0 u 100 -> 2,95%; 0 u 149 -> 1,99%; 0 u 299 -> 1,00% (95%).
  RIZIK: Statisticka tvrdnja se lako procita kao univerzalna -- zato granica mora da stoji u samom verdiktu. Uzorak bez stratifikacije moze da promasi malu suitu (banka: 8 sabotaza -> uvek sve).
  PREDUSLOVI: P1 (ima sta da se revidira); P3 za jeftinu polovinu.
  DOKAZATI: Self-test: namerno pokvaren zapis u registru mora biti nadjen kad upadne u uzorak; isto seme -> isti uzorak; uzorak pokriva svaku suitu.

### P7 Politika timeout-a i 'mrtvorodjenih' sabotaza [trud S] [slabi: NE.]
(a) Timeout prolaza se racuna iz baze, kao kod PIT-a: max(pod, k x izmereno vreme izabranih testova + rezija), umesto fiksnih 600/1200 s. (b) TIMEOUT ostaje 'nije dokazano' (stroze od PIT-a, kao Infection opcija) -- to se ne menja. (c) Za NOVE i IZMENJENE sabotaze: mutirani tekst modula se u memoriji provuce kroz jeftine provere vba_check-a po fajlu (duplo ime, deklaracija posle procedure, arnost...), pa se sabotaza koja bi oborila compile uhvati bez Excela.
  DOBITAK: Zaglavljen prolaz kosta ~1 min umesto 10 min po dogadjaju; u grupi se izbegava i ponavljanje svih clanova. Ucestalost nepoznata, pa je ukupan dobitak nepoznat.
  RIZIK: Pretesan timeout na sporoj masini daje lazan TIMEOUT -- bezbedan smer (prijavi se kao nedokazano), cena je ponovljen prolaz.
  PREDUSLOVI: Vreme po testu (P2).
  DOKAZATI: Sabotaza koja namerno otvori dijalog -> TIMEOUT u izracunatom roku i po imenu; sabotaza sa duplom deklaracijom -> staticki nalaz pre Excela; obe vracene -> zeleno.

### P8 Tri jeftine ustede u hladnom putu (pre P3) [trud S] [slabi: NE -- (a) koristi istu tvrdnju pod istim otiskom; (b) ne dira tvrdnje dokaza; (c) cista optimizacija.]
(a) Baza iz markera: ako vba_gate marker vec kaze da je BAS OVAJ otisak izvora i ugovora zelen za trazenu suitu, dokaz.py ne pusta bazu ponovo (dokaz.py:466). (b) Prolazi dokaza preskacu compile pokusaj (verdikt dokaza se cita iz FAIL redova, a compile je ionako rucna kapija; 'NEJASNO' ne obara run) -- stedi 4,5-7,9 s fiksnih sleep-ova po prolazu. (c) Memoizovati modul -> skup imena testova u sabotaza._suita_testa.
  DOBITAK: (a) 1,6-4,4 min po pozivu dokaz.py (procena) -- kod tipicnog reza od ~20 sabotaza to je 30-50% zidnog vremena. (b) 215 prolaza x 4,5-7,9 s = 16-28 min po punom grupnom dokazu; 1,6-2,8 h po pojedinacnom. (c) 20,8 s -> <0,1 s na svakom punom pozivu i --plan (izmereno gde odlazi).
  RIZIK: (a) Marker iz prolaza nad drugom sveskom -- vba_gate vec vezuje identitet sveske, pa se samo koristi. (b) Gubi se rani pad na compile dijalog kad detekcija radi; nadoknadjuje P7(c) i kraci timeout.
  PREDUSLOVI: Nema.
  DOKAZATI: (a) Izmena jednog bajta u src-vba -> baza se ponovo pusta; self-test u oba smera. (b) Isti rez sa i bez compile koraka -> isti verdikti. (c) cProfile pre/posle; isti plan (215 prolaza, ista podela).

### P9 Piramida sa eksplicitnim budzetima, oslonjena na registar dokaza [trud M] [slabi: Delimicno, isto kao P1: release tvrdnja postaje 'registar potpun i svez + revizija', ne 'sve ponovo izmereno pred release'. Dobija se nesto sto danas ne postoji: dokaz se odrzava KONTINUIRANO, a ne kao ritual koji se zbog trajanja preskace.]
Predlog nivoa i budzeta: (1) hook po izmeni: cilj <=2 s, samo provere tog fajla (danas 9-12 s); (2) lokalno pred commit/push: ceo vba_check + kapije <=60 s, a --self-test nekog alata samo kad je BAS taj alat ili njegov ulaz diran (TIA pravilo: gde se ne zna -- pusti sve); (3) PR na CI (Linux): staticke kapije + PROVERA registra ('svaka sabotaza ciji se kljuc promenio u ovom PR-u ima svez zapis') -- CI ne pokrece Excel, samo cita; (4) razvojna masina, TARGETED: izabrani testovi + inkrementalni dokaz, <=10 min; (5) nocno (Task Scheduler u prijavljenoj interaktivnoj sesiji, ne servis): sve suite, obrnut redosled, dijagnostika zagadjenja, revizija uzorkom i osvezavanje najstarijih zapisa; (6) release: sve suite + rucni Compile + nijedan zapis stariji od X + revizija n>=100. Merge queue ne uvoditi: GitHub ga ne nudi repou pod korisnickim nalogom; postojece pravilo 'rebase lokalno -> pokazi -> push' uz proveru registra na rebase-ovanoj glavi daje isti efekat za jednog-dva autora.
  DOBITAK: Release vise ne trazi prolaz od 7-43 h -- trosak se amortizuje nocno (npr. toplim putem ~300 zapisa za ~20 min: ceo katalog svake ~4 noci; procena). Svaki nivo ima brojku protiv koje se meri regresija trajanja.
  RIZIK: Nocni prolaz zavisi od ukljucene i prijavljene masine; Office automatizacija bez nadzora nije podrzan scenario (watchdog i tvrdi prekid vec postoje). Provera registra na CI-ju trazi da se registar commit-uje ili prilozi -- nov artefakt koji moze da zastari.
  PREDUSLOVI: P1; za punu vrednost P3 i P5. Izmena hook-a je izmena u .claude/ -> zaseban process PR.
  DOKAZATI: Svaki budzet je merljiv ('time' po nivou, upisan uz PR). Provera registra na CI-ju mora pasti na PR-u koji izmeni proceduru sidra a ne donese svez zapis (dvosmerno). Nocni prolaz ostavlja izvestaj sa trajanjem i brojem osvezenih zapisa.

### P10 Fragmenti kataloga i generisan registar testova (lek za fajlove-zarista) [trud L] [slabi: NE -- ako se ocuvaju broj, imena i redosled, sto se proverava alatom.]
(a) SABOTAZE iz tools/sabotaza.py (735 KB, diran u 35/40 PR-ova) podeliti u fragmente po sabotiranom modulu ili po oblasti (tools/sabotaze/<modul>.py), koje motor ucitava glob-om; motor i provere ostaju gde jesu. Dva PR-a tada dopisuju u razlicite fajlove. merge=union NE preporucujem: GitHub ga pri spajanju kroz web ne postuje pouzdano i moze tiho spojiti dve n-torke istog imena. (b) modTest.bas (1,58 MB) i BFP (1,53 MB) podeliti po oblasti, a tabelu TestName/InvokeTest generisati alatom iz konvencije imena (obrazac Unity generate_test_runner / Rubberduck '@TestMethod), sa --check u CI-ju -- isti obrazac koji repo vec ima za modSchema (gen_schema_module.py). Redosled testova ostaje EKSPLICITAN (fajl redosleda) dok P5 ne pokaze da nije nosiv.
  DOBITAK: Manje sudara (grana B se sudara sa main-om bas u ova tri fajla); manji diff za citanje; manji moduli = brza zamena jednog modula u toploj sesiji kad se menja test; inkrementalni kljuc i hook provere mogu da rade po fragmentu (rani prekid po hesu sadrzaja).
  RIZIK: (b) je veliki mehanicki potez koji se sudara sa SVAKOM otvorenom granom -- radi se alatom, u mirnom trenutku, u jednom PR-u bez ikakve druge izmene. Private pomocnici postaju vidljivi izmedju modula ili se moraju deliti; redosled se mora ocuvati bajt-u-bajt.
  PREDUSLOVI: Za (b): dogovoren prozor bez otvorenih grana nad test modulima. (a) moze odmah i nezavisno.
  DOKAZATI: (a) sabotaza.py --lista pre i posle daje identican izlaz (ista imena, isti redosled, 1265); --proveri-sidra zeleno; dvosmerno: obrisan fragment -> provera pada. (b) Popis testova po imenu pre/posle identican (350 / 391 poziva), isti redosled, FULL zeleno na Windows-u; generator --check pada kad se doda test van registra.

### P11 Dinamicka mapa 'test -> procedure u koje je usao' (analog Ekstazi / PIT coverage) [trud L] [slabi: NE -- smanjuje slabljenje koje unosi P1.]
Samo u test svesci, pri uvozu (nikad u src-vba), na ulaz svake procedure doda se jedan poziv brojaca; po testu se zapise skup procedura u koje je usao. Tri upotrebe: (1) TARGETED izbor za PR: diff -> izmenjene procedure -> testovi koji u njih ulaze; izmene koje mapa ne ume da protumaci (deklaracije, sema, forme) -> pusti sve, uz uvek ukljucene prethodno pale i nove testove (pravilo Microsoft TIA); (2) staticki nalaz nad katalogom: deklarisan test uopste ne ulazi u sabotiranu proceduru -> unos ne moze biti dokazan; (3) najvaznije -- inkrementalni kljuc iz P1 dobija STVARNE zavisnosti (hesovi svih procedura u koje je test usao), cime se uklanja PIT-ova nedokazana pretpostavka za tela procedura.
  DOBITAK: Kljuc registra postaje bezbedan na nivou procedure; koliko bi to povecalo ponovna merenja (izmedju izmerenih 1,8% i 12,3% po PR-u) nije poznato. TARGETED nivo prestaje da bude rucni izbor suite.
  RIZIK: Spekulativno -- nigde nije isprobano na ovom kodu. Instrumentacija menja brojeve redova, moze da uspori vrele petlje i dira osetljivu semantiku Err-a (KI-014); kod koji se zove kroz Application.Run po imenu se belezi (dinamicki), ali modul-level stanje i podaci u svesci ne -- zato fixture i sekcije deklaracija moraju u kljuc.
  PREDUSLOVI: P2 i P3 (bez tople sesije je prolaz za prikupljanje mape preskup da bi se cesto osvezavao).
  DOKAZATI: Suite sa i bez instrumentacije -> identicni verdikti i isti tekstovi palih tvrdnji; sabotaza u proceduri P mora biti u mapi svog testa (provera nad celim katalogom je ujedno test mape); izmena procedure koju test poziva posredno -> zapis zastareva.

### P12 Odluke 'ne' -- da se ne otvaraju ponovo [trud S] [slabi: NE.]
(1) Predictive mutation testing: NE -- menja merenje predvidjanjem (kasniji rad: ~33% pogresnih klasifikacija), sto je suprotno pravilu 'zeleno se ne tvrdi bez merenja'. (2) Selektivna mutacija i izbacivanje 'subsumed' unosa: NE -- 1264 od 1265 sabotaza dokazuje RAZLICITU tvrdnju, pa svako izbacivanje ukida jedan dokaz; nema operatora koji bi se birali. (3) Uzorkovanje kao zamena za prvi dokaz: NE -- samo kao revizija (P6). (4) Mutant schemata u produkcionom izvoru (If Mutant = n Then ...): NE -- testirao bi se kod koji se ne isporucuje, a 388 viselinijskih sabotaza se ne da mehanicki umotati; varijanta 'samo u test svesci' posle P3 donosi jos ~1-2 s po mutaciji, premalo za rizik. (5) LibreOffice / twinBASIC / emulatori kao izvor verdikta: NE -- nema Excel objektnog modela ni UserForm-i; mogu samo kao neobavezna rana provera cistih funkcija, nikad kao 'zeleno'. (6) Jedna dugoveka Excel sesija bez recikliranja: NE. (7) Dalje ulaganje u headless Compile: NE -- Excel nema API za to (komanda 578 trazi vidljiv i aktivan VBE), ostaje rucna kapija. (8) Grupni dokaz kao release put: povuci cim P2+P3 prorade -- literatura ga zna kao 'higher-order' mutante sa maskiranjem, a repo vec priznaje da tvrdi manje.
  DOBITAK: Usteda vremena na predlozima koji izgledaju privlacno a slabe sistem ili ne donose merljivu korist.
  RIZIK: (4) i (7) vredi ponovo proceniti ako se okolnosti promene (npr. P3 se pokaze nestabilnim).
  PREDUSLOVI: Nema.
  DOKAZATI: Nije primenljivo -- ovo su obrazlozene odluke; brojke na kojima pocivaju su u cinjenicama (1264/1265, 388/1265).


## NEPOZNANICE
[
 "Raspodela dokumentovanih ~18 s rezije (start Excela / uvoz 194 fajla / compile pokusaj / EnsureRuntimeSchema / gasenje) nije merena -- nema Excela. Iz koda se zna samo da su 4,5-7,9 s fiksni sleep-ovi.",
 "Vreme po testu ne postoji: grana B ga uvodi, ali VBA strana nikad nije izvrsena. Sve procene za P2/P3 (1-5 s po ciljanom testu) pocivaju na proseku suite/broj testova; BFP seed koji se pusta uvek moze znatno da ih pokvari.",
 "Koliko testova zavisi od redosleda (ne prolaze sami, ili ih prethodni maskira) -- 'dokaz.py --izolacija' nije nikad pusten; obrnut redosled takodje.",
 "Da li Excel VBE stabilno podnosi stotine uzastopnih zamena jednog modula u istoj sesiji (rast memorije, [break], 'Modul1' pri odlozenom Remove-u) i da li programska zamena modula u Excelu resetuje stanje celog projekta kao sto MS KB opisuje za Access.",
 "Stvarna stopa zastarelosti inkrementalnog kljuca: merio sam samo koliko se kljuc MENJA (1,8% / 12,3% po PR-u), ne koliko cesto se verdikt promeni a kljuc ne. To se dobija jedino jednim punim dokazom uporedjenim sa registrom.",
 "Ucestalost TIMEOUT-a i compile-obarajucih sabotaza u praksi (odredjuje vrednost P7).",
 "Kapacitet razvojne masine za paralelne Excel instance (jezgra, RAM, ponasanje antivirusa nad temp .xlsm, licencni dijalozi).",
 "Da li je GitHub nalog repoa korisnicki ili organizacioni (zakljuceno iz remote URL-a, nije provereno) -- od toga zavisi tvrdnja da merge queue nije dostupan.",
 "Spoljni brojevi (PIT 1,25 / 4000 ms; iDFlakies 50,5%; TCE 7% / 21%; Lam 65% / 1%; Shi 79,4%; schemata ~5x; Gopinath ~1000 mutanata) su iz sazetaka pretrage -- nijedna stranica nije otvorena. Pre nego sto udju u formalan dokument treba ih potvrditi na izvoru; pitest.org i stryker-mutator.io moraju biti dozvoljeni u mreznim podesavanjima sesije.",
 "Da li bi twinBASIC mogao da kompajlira .bas module ovog projekta kao dodatnu compile proveru van Excela -- nije isprobano; tretirati kao spekulaciju.",
 "Koliko sabotaza zavisi od stanja koje postavlja RANIJI test u suiti (pa izolovan test mutaciju ne bi oborio) -- grana B takav slucaj vraca na celu suitu, ali njihov broj odredjuje stvarni dobitak P2."
]
