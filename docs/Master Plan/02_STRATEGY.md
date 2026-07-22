# 02 — Strategija i identitet AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-22

---

## 1. Polazna tačka

AgriX je specijalizovan poslovni sistem za organizaciju otkupa poljoprivrednih proizvoda. Trenutni proizvod povezuje desktop administraciju, management PWA, monitoring, self-update i niz poslovnih modula. Za narednu sezonu planirani su usklađeni PWA Otkupac, kiosk tableti i termalna štampa otkupnih listova na mestu otkupa.

Potvrđene polazne činjenice i radne pretpostavke:

- `FACT`: postoje tri aktivna klijenta;
- `TARGET`: u startu je realan cilj približno pet novih klijenata;
- `FACT`: prosečna firma ima približno deset otkupnih stanica;
- `FACT`: prosečno postoji jedan desktop korisnik po firmi, uz management PWA;
- `FACT`: prosečno postoji oko 100 kooperanata po firmi;
- `MEASURED`: support je približno jedan poziv nedeljno u dosadašnjem periodu;
- `FACT`: self-update i granularni monitoring postoje;
- `FACT`: onboarding je do sada rađen remote;
- `ASSUMPTION`: onboarding se može standardizovati sa jednog na približno pola dana;
- `FACT`: sve klijentske varijacije ostaju u zajedničkom kodu i rešavaju se konfiguracijom;
- `HYPOTHESIS`: u Srbiji postoji približno 500–1.000 relevantnih firmi; procena mora biti potvrđena tržišnim istraživanjem;
- `TARGET`: AgriX treba da osvoji najmanje 200 firmi u naredne 3–4 godine, ako readiness i tržišna validacija to podrže.

AgriX više nije prototip, ali još nije dokazano skaliran proizvod. Sledeće faze moraju dokazati da proizvod, onboarding, support, hardver i organizacija mogu da rastu bez pada pouzdanosti.

---

## 2. Strateška definicija

AgriX nije generički ERP i nije zbir nepovezanih modula. AgriX je vertikalna operativna platforma za firme koje organizuju otkup poljoprivrednih proizvoda, njihove terenske mreže i kooperante.

Platforma ima tri osnovna sloja:

1. **AgriX Otkup / Desktop** — centralna administracija, dokumentacija, finansijski i operativni tokovi;
2. **AgriX Field / PWA Otkupac** — rad na otkupnim stanicama, kiosk terminali i štampa na licu mesta;
3. **AgriX Gazdinstvo** — digitalna veza kooperanta sa otkupljivačem i samostalna evidencija gazdinstva.

Hladnjače i druge firme sa razgranatom mrežom otkupnih stanica primarni su kupci i glavni izvor prihoda u kratkom roku. Gazdinstvo trenutno nije finansijski oslonac, ali može postati glavni proizvod ili glavni izvor prihoda ako tržište potvrdi konverziju i dugoročnu vrednost.

---

## 3. Vizija

Do 2030. AgriX treba da bude vodeća regionalna platforma za digitalizaciju otkupa poljoprivrednih proizvoda, sa snažnom bazom u Srbiji i prenosivim operativnim modelom za tržišta regiona.

`TARGET`: u naredne 3–4 godine izgraditi bazu od najmanje 200 firmi, uz rast koji je dozvoljen readiness-om, a ne proizvoljnim godišnjim plafonom.

Vizija podrazumeva da AgriX:

- pouzdano vodi kritični tok otkupa tokom sezone;
- povezuje centralnu administraciju sa terenskim stanicama;
- digitalizuje odnos firme i kooperanta;
- smanjuje ručni rad, prepisivanje i kašnjenje informacija;
- daje managementu trenutni uvid i kontrolu;
- ostaje standardizovan proizvod sa jednim kodom;
- može da opslužuje stotine firmi kroz procese, automatizaciju i delegiranje;
- stvara dovoljno recurring prihoda za stalni razvoj, podršku i regionalno širenje;
- postane prirodni tehnološki dobavljač ciljnom segmentu, a ne samo dobavljač jedne aplikacije.

Cilj od 200 firmi je strateški cilj, ne finansijska prognoza. Mora se razložiti na godišnje akvizicione, operativne i kadrovske kapacitete.

---

## 4. Misija

AgriX omogućava hladnjačama i drugim organizovanim otkupljivačima da vode otkup, dokumentaciju, ambalažu, finansijske tokove, terenske stanice i kontrolu poslovanja kao jedan povezan sistem, bez troška i složenosti velikog ERP projekta.

Za kooperante AgriX treba da obezbedi jasan digitalni pregled saradnje sa otkupljivačem i jednostavan alat za sopstvenu evidenciju gazdinstva.

---

## 5. Šta AgriX jeste, a šta nije

AgriX jeste:

- vertikalna platforma za otkup;
- zajednički proizvod za veliki broj firmi;
- sistem koji se prilagođava konfiguracijom;
- softverski i operativni paket za centralu i terenska otkupna mesta;
- potencijalni dobavljač šireg IT sistema za ciljne klijente.

AgriX namerno nije:

- univerzalni knjigovodstveni program;
- potpuna zamena za BizniSoft, PANTHEON ili drugi računovodstveni ERP;
- custom software studio koji pravi poseban proizvod za svakog klijenta;
- jeftin program samo za štampanje otkupnih listova;
- klasičan hardverski distributer bez sopstvene tehnološke vrednosti;
- proizvod koji obećava funkcionalnosti koje nisu production-ready;
- projekat koji menja tehnologiju samo zato što je nova tehnologija atraktivnija.

AgriX može prodavati hardver i širu IT opremu, ali samo kada time povećava pouzdanost, standardizaciju, prihod i kontrolu ukupnog rešenja.

---

## 6. Ciljno tržište i idealni kupac

### Primarni fokus

Trenutni geografski fokus je Srbija.

Primarni ciljni kupci su:

- hladnjače sa sopstvenom mrežom otkupnih stanica;
- firme koje se bave organizovanim otkupom i imaju razgranatu mrežu stanica i kooperanata;
- firme kojima generički ERP ne rešava dovoljno dobro ulazni tok robe;
- firme koje žele centralnu kontrolu terenskog otkupa.

Tipičan idealni kupac ima:

- više otkupnih stanica, često oko deset;
- jednog centralnog administrativnog korisnika;
- management koji želi pregled i kontrolu kroz PWA;
- postojeći knjigovodstveni sistem;
- odgovornu osobu za implementaciju i podatke;
- spremnost da standardizuje proces umesto da zahteva poseban fork;
- dovoljno veliki operativni problem da godišnja licenca ima jasnu vrednost.

Promet od 1–2 miliona EUR jeste čest profil sadašnjih ciljnih klijenata, ali nije jedini kriterijum. Broj stanica, složenost toka, obim dokumenata, potreba za kontrolom i spremnost na standardizaciju važniji su od samog prometa.

### Anti-ICP

AgriX ne treba aktivno da prihvata klijente koji:

- zahtevaju sopstvenu verziju koda ili poseban release;
- očekuju neograničen custom development uključen u licencu;
- nemaju odgovornu osobu za podatke, obuku i komunikaciju;
- traže rollout neposredno pred sezonu bez vremena za test;
- odbijaju standardne procese backupa, update-a i monitoringa;
- kupuju isključivo po najnižoj ceni;
- očekuju da AgriX preuzme fizičke kvarove i zloupotrebu hardvera bez ugovorne granice.

---

## 7. Strateški principi razvoja

### 7.1 Jedan kod, bez forkova

Sve funkcionalnosti ostaju u zajedničkom kodu. Razlike među firmama rešavaju se kroz podešavanja, module, dozvole, workflow konfiguraciju i feature flags.

Trajni klijentski fork nije dozvoljen.

### 7.2 Pouzdanost i brzina rasta nisu suprotnosti

AgriX ne bira između sigurnosti i rasta. Cilj je da automatizacijom, monitoringom, self-update-om, standardizovanim onboardingom i delegiranjem poveća brzinu rasta bez pada pouzdanosti.

### 7.3 Sezonski cap određuje readiness

Ne postoji unapred fiksiran hard cap od 10, 15 ili 20 firmi. Maksimalan broj novih firmi za sezonu određuje se isključivo kroz readiness score pre prodajnog i implementacionog ciklusa.

Readiness mora najmanje da obuhvati:

- stabilnost kritičnog codebase-a;
- production readiness PWA Otkupac toka;
- pouzdanost sync-a i termalne štampe;
- automatizaciju i trajanje onboardinga;
- kapacitet customer supporta i eskalacije;
- monitoring, recovery i release procese;
- dostupnost i logistiku hardvera;
- finansijsku rezervu i obrtni kapital;
- broj osoba koje mogu sprovesti standardan onboarding bez osnivača.

Cap određuje najslabija kritična komponenta, ne prosečna ocena. Detaljna metodologija biće definisana u posebnom operations/readiness poglavlju.

### 7.4 Kontrolisan staged rollout

Velike promene prolaze kroz interni test, pilot firmu, ograničenu grupu i tek zatim pun rollout. Self-update omogućava distribuciju, ali ne uklanja potrebu za staged rolloutom.

### 7.5 Bez rewrite-a bez merljivog razloga

Promena tehnološke platforme razmatra se kada postojeća platforma stvara merljiv limit u pouzdanosti, brzini razvoja, zapošljavanju, integracijama ili ukupnom trošku održavanja.

### 7.6 Operativna jednostavnost je funkcionalnost

Remote onboarding, monitoring, self-update, backup, kiosk konfiguracija, manuali i runbook-ovi imaju isti strateški značaj kao korisničke funkcije.

### 7.7 Nema prikrivenog custom developmenta

Zahtev jednog klijenta ulazi u proizvod samo kada predstavlja opšti problem segmenta i može se rešiti kroz zajednički model. Poseban razvoj se posebno ugovara ili odbija.

### 7.8 Ne obećavati budući proizvod kao postojeći

PWA Otkupac, termalna štampa, kiosk terminali i budući moduli prodaju se kao production tek kada zadovolje release kriterijume.

---

## 8. Strategija rasta 2026–2030

### Faza 1 — Dokaz readiness modela

**Period:** naredna sezona.

Polazni komercijalni cilj je približno pet novih firmi, ali stvarni broj može biti 10, 15 ili 20 ako readiness score pokaže da onboarding, podrška, proizvod i hardver mogu bezbedno da iznesu taj obim.

Ciljevi:

- standardizovati remote onboarding kroz manuale i checklistu;
- omogućiti da standardan onboarding vodi customer support / implementation osoba;
- potvrditi PWA Otkupac, kiosk i termalnu štampu;
- meriti vreme po onboardingu i support case-u;
- potvrditi da broj novih firmi može rasti bez forkova i bez rasta incidenta po firmi;
- napraviti formalni readiness score pre početka aktivne prodaje.

### Faza 2 — Ubrzana nacionalna penetracija

**Okvir:** od približno 10 do 50 firmi.

Ciljevi:

- customer support / implementation osoba preuzima standardna pitanja i onboarding;
- osnivač ostaje eskalacija za bugove i poslovnu logiku;
- osnivač se postepeno prebacuje na marketing, prodaju i partnerstva;
- developer se dodaje kada razvoj postane usko grlo ili kada je potreban da oslobodi osnivača za tržište;
- razviti sistem preporuka, case studies i direktnog pristupa ciljnom segmentu;
- potvrditi pricing i unit economics na različitim segmentima.

### Faza 3 — Liderstvo u Srbiji

**Okvir:** približno 50–200 firmi.

Ciljevi:

- izgraditi najprepoznatljiviji specijalizovani brend za otkup u Srbiji;
- organizovati support, implementaciju i razvoj tako da dnevni rad ne zavisi od osnivača;
- imati standardizovan hardverski katalog i logistiku;
- razviti partnerstva sa knjigovođama, dobavljačima opreme i relevantnim organizacijama;
- pretvoriti Gazdinstvo iz hipoteze u dokazani kanal ili dokazano odbaciti njegovu ekonomiku;
- pripremiti proizvod i organizaciju za regionalnu ekspanziju.

### Faza 4 — Regionalna platforma

Regionalna platforma je cilj, ne samo jedna od opcija.

Početna tržišta za procenu su:

1. Srbija kao baza;
2. BiH;
3. Crna Gora i Severna Makedonija;
4. Hrvatska i druga tržišta nakon pravne, jezičke, poreske i prodajne procene.

Regionalno širenje mora imati lokalni prodajni kanal, regulatorno mapiranje, support model i jasnu odgovornost za implementaciju.

---

## 9. Strategija prihoda

Redosled važnosti u kratkom roku:

1. godišnje licence firmi;
2. implementacija i obuka kao odvojene usluge;
3. dodatni moduli i multi-company licence;
4. marža na terminalima i drugoj IT opremi;
5. Gazdinstvo Basic i Pro;
6. buduće integracije, premium support i SLA.

### Gazdinstvo

Gazdinstvo trenutno ne finansira osnovni biznis. Sa približno 100 kooperanata po firmi i radnom hipotezom konverzije od 5%, početni prihod je trivijalan.

To ne ograničava njegov dugoročni potencijal. Gazdinstvo može postati glavni proizvod ili glavni izvor prihoda ako se potvrde:

- dovoljno velika baza aktivnih kooperanata;
- održiva konverzija na Basic i Pro;
- niska cena podrške i akvizicije;
- jaka retencija i svakodnevna korisnost;
- dodatni proizvodi koji povećavaju ARPU.

### Hardver i širi IT sistem

Hardver nije glavni profitni centar, ali treba da bude profitabilan sporedni centar.

AgriX treba da razmotri pozicioniranje kao dobavljač kompletnog IT sistema za ciljne klijente, uključujući:

- kiosk tablete;
- termalne štampače;
- mrežnu i rezervnu opremu;
- desktop računare i osnovnu konfiguraciju kada je komercijalno opravdano;
- standardizaciju uređaja, remote management i zamenske jedinice;
- koordinaciju sa vagama, štampačima i drugim perifernim sistemima.

Ovaj pravac je prihvatljiv samo kada:

- svaka kategorija ima pozitivnu stvarnu maržu nakon rada i garancije;
- ne odvlači organizaciju od razvoja softvera;
- povećava pouzdanost i vezanost klijenta;
- postoji standardizovan katalog, nabavka i support granica.

---

## 10. Strategija organizacije

### Osnivač

U narednoj fazi osnivač zadržava:

- product ownership;
- arhitekturu i ključni razvoj;
- finalnu eskalaciju supporta;
- prodaju važnim klijentima;
- marketing strategiju;
- ključna partnerstva.

Cilj nije da osnivač trajno ostane operativno usko grlo, već da se standardni rad postepeno delegira.

### Prvo zaposlenje

Prva operativna osoba je customer support / implementation.

Njena uloga uključuje:

- rešavanje standardnih i baznih korisničkih pitanja;
- remote onboarding prema manualima i checklistama;
- pomoć oko tableta, štampača i konfiguracije;
- praćenje monitoringa;
- evidenciju i trijažu problema;
- rešavanje poznatih problema prema runbook-u;
- delegiranje i eskalaciju ostatka osnivaču ili developeru;
- vođenje evidencije o vremenu, uzroku i rešenju support slučajeva.

Osnivač pruža podršku van smene te osobe tokom sezone i rešava složene eskalacije.

### Dodatni developer

Developer se dodaje kada:

- razvoj postane dokazano usko grlo;
- roadmap ne može biti isporučen uz postojeći kapacitet;
- osnivač treba značajno da se prebaci na marketing i prodaju;
- trošak propuštenog rasta postane veći od punog troška developera.

---

## 11. Partner i kapital

AgriX ne treba partnera samo zbog kapitala.

Partner ima smisla kada donosi najmanje jednu teško zamenljivu sposobnost:

- direktan pristup velikom broju kvalitetnih kupaca;
- dokazanu distribuciju u agraru;
- operativno vođenje prodaje i implementacije;
- relevantno iskustvo skaliranja B2B softvera;
- regionalnu mrežu koju AgriX ne može brzo sam da izgradi;
- kapital vezan za precizan, validiran plan ubrzanja.

Pre prodaje udela moraju biti poznati upotreba kapitala, očekivani dodatni ARR, rok, odgovornost partnera, upravljačka prava, dilution i scenario neuspeha.

Partner ili investicija mogu postati racionalni ranije nego što je prvobitno planirano ako readiness pokaže da je potražnja veća od kapaciteta i da kapital direktno uklanja dokazano usko grlo.

---

## 12. Strateška hitnost i tržišni prozor

`HYPOTHESIS`: tržište specijalizovanog softvera za otkup u Srbiji ima ograničen prozor u kojem AgriX može izgraditi dominantnu poziciju pre nego što postojeći ERP dobavljači ili novi vertikalni konkurent razviju sličan proizvod.

Zbog toga strategija ne sme biti pasivna. Pouzdanost ostaje uslov, ali cilj nije beskonačno dokazivanje na malom broju firmi. Cilj je što brže povećavati readiness i zatim pretvarati readiness u tržišni udeo.

`TARGET`: najmanje 200 firmi u periodu od 3–4 godine.

Ovaj cilj zahteva poseban plan:

- broj novih firmi po sezoni;
- broj osoba za onboarding i support;
- potrebni razvojni kapacitet;
- marketing i prodajni kanali;
- hardverski obrtni kapital;
- regionalni roadmap;
- minimalni ARR i cash reserve po fazi.

---

## 13. Strateški rizici

| Rizik | Verovatnoća | Uticaj | Primarna zaštita |
|---|---|---|---|
| Osnivač ostaje jedina osoba koja razume ceo sistem | Visoka | Visok | dokumentacija, support osoba, developer, ownership |
| Previše razvoja pred sezonu | Visoka | Visok | scope freeze i release gate |
| Field štampa ili sync nisu dovoljno pouzdani | Srednja | Kritičan | pilot, staged rollout, rezervni proces |
| Readiness score preceni kapacitet | Srednja | Kritičan | weakest-link model i konzervativna rezerva |
| Cena je niža od punog troška usluge | Srednja | Visok | unit economics po firmi |
| Hardver veže previše kapitala | Srednja | Visok | predujam, standardni modeli, ograničena zaliha |
| Širi IT portfolio odvlači fokus | Srednja | Srednji/visok | profitabilnost po kategoriji i jasne granice |
| Klijenti traže prikriven custom razvoj | Visoka | Srednji | ugovorne granice i product pravila |
| Gazdinstvo nema dovoljnu konverziju | Visoka | Srednji | tretirati prihod kao nulu u konzervativnom planu |
| Rast bude prespor i konkurent zauzme tržište | Srednja/visoka | Kritičan | ambiciozan GTM i rast readiness-a |
| Agresivan rast pogorša kvalitet | Srednja | Kritičan | readiness-based cap i staged onboarding |
| Partnerstvo prerano smanji kontrolu | Srednja | Visok | jasni kriterijumi pre prodaje udela |

---

## 14. Ključni strateški KPI-jevi

- broj aktivnih firmi;
- broj novih firmi po sezoni;
- readiness score pre sezone i po kritičnoj komponenti;
- onboarding sati po firmi;
- procenat onboardinga koji support osoba vodi bez osnivača;
- prosečno support vreme po firmi mesečno;
- procenat problema rešenih bez osnivača;
- kritični incidenti i recovery vreme;
- uspešnost self-update rollouta;
- broj aktivnih Field terminala;
- stopa neuspešne ili ponovljene štampe;
- zahtevi rešeni konfiguracijom naspram novog koda;
- ARR po firmi i ukupni ARR;
- ostvarena cena naspram cenovnika;
- stvarna hardverska marža;
- renewal i churn;
- Gazdinstvo aktivacija, konverzija, ARPU i support cost;
- tržišni udeo u procenjenom adresabilnom segmentu.

---

## 15. Odobrene strateške odluke

### STR-001 — Readiness-based rast

Ne postoji unapred fiksiran sezonski hard cap. Maksimalan broj novih firmi određuje readiness score organizacije i proizvoda u trenutku prodaje i implementacije.

### STR-002 — Primarni tržišni fokus

Trenutni fokus je Srbija. Ciljni kupci su hladnjače i druge firme koje imaju razgranatu mrežu otkupnih stanica i kooperanata.

### STR-003 — Jedan proizvod, jedan kod

Klijentske razlike rešavaju se zajedničkim kodom i konfiguracijom. Trajni klijentski fork nije dozvoljen.

### STR-004 — Dinamična uloga Gazdinstva

Licence firmi trenutno finansiraju osnovni biznis. Gazdinstvo se kratkoročno ne tretira kao ključni prihod, ali nema strateško ograničenje koje bi sprečilo da postane glavni proizvod ili izvor prihoda ako podaci to potvrde.

### STR-005 — Hardver kao sporedni profitni centar

Hardver nije glavni profitni centar, ali mora imati pozitivnu stvarnu maržu. AgriX može postati dobavljač šireg IT sistema ciljnih klijenata kada to povećava pouzdanost, prihod i stratešku poziciju.

### STR-006 — Bez partnera samo zbog novca

Partner ili investitor razmatra se kada rešava dokazano usko grlo i donosi merljivu sposobnost pored kapitala.

### STR-007 — Prvo operativno zaposlenje

Prva operativna uloga je customer support / implementation. Ta osoba rešava bazne slučajeve, sprovodi onboarding prema manualima i delegira složene probleme.

### STR-008 — Regionalna platforma

Dugoročni cilj AgriX-a je regionalna vertikalna platforma, ne trajno ograničen lokalni specijalista.

### STR-009 — Strateški cilj tržišnog udela

AgriX cilja najmanje 200 firmi u naredne 3–4 godine. Cilj se tretira kao ambicija koju treba operacionalizovati, ne kao garantovana prognoza.

---

## 16. Otvorene teme za naredna poglavlja

1. Potvrditi procenu da u Srbiji postoji 500–1.000 relevantnih firmi.
2. Definisati matematički readiness score i pragove po sezonskom kapacitetu.
3. Razložiti cilj od 200 firmi na godišnji prodajni i kadrovski plan.
4. Odrediti trenutak zapošljavanja prve support / implementation osobe.
5. Izračunati finansiranje hardverske nabavke i minimalnu zalihu.
6. Definisati koje IT kategorije AgriX prodaje, a koje ne.
7. Odrediti ARR i operativne pragove za partnera ili investitora.

---

## 17. Naredni koraci

1. upisati STR-001 do STR-009 u `DECISION_LOG.md`;
2. razviti `03_CUSTOMERS_AND_JOBS.md`;
3. razviti `04_MARKET.md` i potvrditi adresabilno tržište;
4. razviti `07_PRODUCT_PORTFOLIO.md`;
5. napraviti readiness model pre finalnog finansijskog plana;
6. zatim finalizovati pricing, unit economics i plan rasta do 200 firmi.
