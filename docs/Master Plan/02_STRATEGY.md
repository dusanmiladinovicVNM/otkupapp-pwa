# 02 — Strategija i identitet AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-22

---

## 1. Polazna tačka

AgriX je vertikalni poslovni operativni sistem za firme koje organizuju otkup poljoprivrednih proizvoda. Proizvod ne pokriva samo evidenciju otkupa, već povezuje terenski rad, prijem robe, logistiku, transport, dokumentaciju, finansije, prodaju, regulatorne obaveze, upravljanje i saradnju sa kooperantima.

Potvrđene činjenice i radne pretpostavke:

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
- `FACT`: AgriX ima management Dispečer modul i operativni tok za vozače;
- `FACT`: AgriX podržava SEF tok;
- `FACT`: AgriX automatski preuzima i obrađuje bankarske izvode, rasknjižava prilive i odlive na kooperante i kupce i priprema naloge za plaćanje;
- `HYPOTHESIS`: u Srbiji postoji približno 500–1.000 relevantnih firmi; procena mora biti potvrđena tržišnim istraživanjem;
- `TARGET`: AgriX treba da osvoji najmanje 200 firmi u naredne 3–4 godine, ako readiness i tržišna validacija to podrže.

AgriX više nije prototip, ali još nije dokazano skaliran proizvod. Sledeće faze moraju dokazati da ceo sistem, onboarding, support, hardver i organizacija mogu da rastu bez pada pouzdanosti.

---

## 2. Strateška definicija

AgriX nije samo program za otkup i nije skup nepovezanih modula. AgriX je **end-to-end vertikalni poslovni operativni sistem** za hladnjače i druge firme sa razgranatom mrežom otkupnih stanica, vozača, kupaca i kooperanata.

Strateški cilj proizvoda je da pokrije ceo posao klijenta u domenu u kojem AgriX ima smisla: sve glavne tokove i najvažnije sporedne tokove, od prvog kontakta sa kooperantom do konačne naplate kupca, isplate kooperanta i management kontrole.

### 2.1 Poslovni domeni

AgriX se definiše kroz poslovne domene, a ne samo kroz aplikacije ili ekrane.

1. **Kooperanti i priprema sezone**
   - matični podaci;
   - ugovorni i proizvodni kontekst;
   - stanice, otkupljivači, cenovnici i pravila;
   - Gazdinstvo i digitalna veza sa kooperantom.

2. **Terenski otkup i prijem robe**
   - PWA Otkupac;
   - unos na otkupnoj stanici;
   - kiosk tablet i termalna štampa na licu mesta;
   - otkupni listovi, prijemnice, reversi i prateća dokumentacija;
   - ambalaža, kvalitet, klase i povezani podaci.

3. **Logistika, dispečer i vozači**
   - real-time pregled otkupljene i neraspoređene robe;
   - planiranje preuzimanja po stanicama;
   - raspodela kamiona i vozača;
   - kapaciteti, rute i statusi transporta;
   - praćenje izvršenja i zatvaranje transportnog toka;
   - organizacija isporuke kupcima kada je primenljivo.

4. **Lager, ambalaža, palete, prerada i sledljivost**
   - stanje robe i ambalaže;
   - paletni i proizvodni tokovi;
   - prerada, povezivanje izvora i izlaza;
   - sledljivost od kooperanta i stanice do kupca;
   - operativni dokumenti i kontrole.

5. **Kupci, prodaja i regulatorna dokumentacija**
   - kupci, fakture i potraživanja;
   - izlazna dokumentacija;
   - SEF slanje, statusi i kontrola elektronskih faktura;
   - veza prodajnog toka sa robom, dokumentima i finansijama.

6. **Finansije i trezor**
   - obaveze prema kooperantima;
   - potraživanja od kupaca;
   - avansi, uplate, isplate i salda;
   - automatsko preuzimanje bankarskih izvoda;
   - automatsko ili kontrolisano rasknjižavanje stavki na kooperante i kupce;
   - priprema naloga za plaćanje;
   - kontrola preplata, duplikata i neusaglašenih transakcija.

7. **Management, kontrola i monitoring**
   - real-time pregled poslovanja;
   - Dispečer i operativna kontrola mreže;
   - KPI, izveštaji i analitika;
   - granularni tehnički monitoring;
   - audit trag, storno, korekcije i odgovornost;
   - kontrola korisnika, rola i pristupa.

8. **Ekosistem kooperanata — Gazdinstvo**
   - Partner nalog povezan sa hladnjačom;
   - Basic i Pro samostalna evidencija;
   - dugoročna digitalna veza između firme i kooperanta;
   - mogući budući izvor dominantnog prihoda ako tržište to potvrdi.

### 2.2 Kanali pristupa

Desktop, PWA i hardver nisu odvojeni proizvodi bez veze, već kanali kroz koje različite uloge pristupaju istom poslovnom sistemu:

- **Desktop** — centralna administracija i složeni back-office tokovi;
- **Management PWA** — pregled, kontrola, Dispečer i odlučivanje;
- **PWA Otkupac** — rad na otkupnoj stanici;
- **PWA Vozač** — izvršenje i status transportnog zadatka;
- **Gazdinstvo PWA** — kooperant;
- **Kiosk terminali i termalni štampači** — standardizovan terenski rad;
- **Integracije** — SEF, banke i budući eksterni sistemi.

Ova razlika je strateški važna: AgriX se prodaje kao jedan povezan sistem, a ne kao kolekcija aplikacija.

---

## 3. Vizija

Do 2030. AgriX treba da bude vodeći regionalni poslovni operativni sistem za organizovani otkup poljoprivrednih proizvoda, sa snažnom bazom u Srbiji i prenosivim modelom za tržišta regiona.

`TARGET`: u naredne 3–4 godine izgraditi bazu od najmanje 200 firmi, uz rast koji je dozvoljen readiness-om, a ne proizvoljnim godišnjim plafonom.

Vizija podrazumeva da AgriX:

- pokriva glavni poslovni ciklus firme od kooperanta i otkupa do transporta, kupca, fakture, naplate i isplate;
- povezuje centralu, stanice, management, dispečere, vozače i kooperante;
- smanjuje ručni rad, prepisivanje i kašnjenje informacija;
- daje managementu trenutni uvid i realnu operativnu kontrolu;
- zatvara regulatorne i finansijske tokove kroz SEF i bankarske integracije;
- ostaje standardizovan proizvod sa jednim kodom;
- može da opslužuje stotine firmi kroz procese, automatizaciju i delegiranje;
- stvara dovoljno recurring prihoda za stalni razvoj, podršku i regionalno širenje;
- postane prirodni tehnološki dobavljač ciljnom segmentu, a ne dobavljač jedne aplikacije.

Cilj od 200 firmi je strateški cilj, ne finansijska prognoza. Mora se razložiti na godišnje akvizicione, operativne, kadrovske i finansijske kapacitete.

---

## 4. Misija

AgriX omogućava hladnjačama i drugim organizovanim otkupljivačima da vode celokupan operativni posao kao jedan povezan sistem: kooperante, otkup, terenske stanice, dokumentaciju, ambalažu, transport, vozače, dispečersko planiranje, lager, kupce, fakture, SEF, banku, isplate, naplate i management kontrolu.

Za kooperante AgriX obezbeđuje jasan digitalni pregled saradnje sa otkupljivačem i jednostavan alat za sopstvenu evidenciju gazdinstva.

---

## 5. Šta AgriX jeste, a šta nije

AgriX jeste:

- vertikalni poslovni operativni sistem;
- end-to-end platforma za glavni i povezane sporedne tokove otkupa;
- zajednički proizvod za veliki broj firmi;
- sistem koji se prilagođava konfiguracijom;
- softverski i operativni paket za centralu, teren, transport i management;
- integracioni sloj prema SEF-u, bankama i drugim relevantnim sistemima;
- potencijalni dobavljač šireg IT sistema za ciljne klijente.

AgriX namerno nije:

- generički ERP za sve industrije;
- univerzalni knjigovodstveni program;
- custom software studio koji pravi poseban proizvod za svakog klijenta;
- jeftin program samo za štampanje otkupnih listova;
- klasičan hardverski distributer bez sopstvene tehnološke vrednosti;
- proizvod koji obećava funkcionalnosti koje nisu production-ready;
- projekat koji menja tehnologiju samo zato što je nova tehnologija atraktivnija.

AgriX ne mora da zameni svaki računovodstveni program, ali treba da bude primarni operativni sistem klijenta i da zatvori sve tokove koji su specifični za otkupni biznis.

---

## 6. Ciljno tržište i idealni kupac

Trenutni geografski fokus je Srbija.

Primarni ciljni kupci su:

- hladnjače sa sopstvenom mrežom otkupnih stanica;
- firme koje se bave organizovanim otkupom i imaju razgranatu mrežu stanica i kooperanata;
- firme kojima generički ERP ne rešava dovoljno dobro ulazni tok robe, logistiku i finansijsko zatvaranje otkupa;
- firme koje žele centralnu kontrolu terenskog rada, vozača, transporta, kupaca, naplate i isplate.

Tipičan idealni kupac ima:

- više otkupnih stanica, često oko deset;
- jednog centralnog administrativnog korisnika;
- management koji želi pregled i kontrolu kroz PWA;
- dispečersku ili transportnu potrebu;
- postojeći knjigovodstveni sistem;
- odgovornu osobu za implementaciju i podatke;
- spremnost da standardizuje proces umesto da zahteva poseban fork;
- dovoljno veliki operativni problem da godišnja licenca ima jasnu vrednost.

Promet od 1–2 miliona EUR jeste čest profil sadašnjih ciljnih klijenata, ali nije jedini kriterijum. Broj stanica, složenost logistike, broj kooperanata, obim dokumenata, bankarskih transakcija i potreba za centralnom kontrolom važniji su od samog prometa.

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

Sve funkcionalnosti ostaju u zajedničkom kodu. Razlike među firmama rešavaju se kroz podešavanja, module, dozvole, workflow konfiguraciju i feature flags. Trajni klijentski fork nije dozvoljen.

### 7.2 Pokriti posao, ne samo funkcije

Nova funkcionalnost se vrednuje prema tome da li zatvara važan poslovni tok, uklanja ručni prelaz između sistema ili povećava kontrolu nad procesom. Lista funkcija bez zatvorenog procesa nije dovoljna.

### 7.3 Glavni i sporedni tokovi

AgriX mora pouzdano pokriti glavne tokove, ali i sporedne tokove bez kojih glavni proces u praksi ostaje nedovršen. Primer: otkup nije završen ako transport, prijem, dokumentacija, isplata i management kontrola ostanu van sistema.

### 7.4 Jedinstven podatak kroz ceo tok

Podatak nastao na stanici treba da se koristi dalje u logistici, prijemu, finansijama, izveštajima, SEF-u i banci bez ponovnog ručnog unosa. Duplo unošenje je signal da proces nije potpuno zatvoren.

### 7.5 Pouzdanost i brzina rasta nisu suprotnosti

AgriX automatizacijom, monitoringom, self-update-om, standardizovanim onboardingom i delegiranjem povećava brzinu rasta bez pada pouzdanosti.

### 7.6 Sezonski cap određuje readiness

Ne postoji unapred fiksiran hard cap od 10, 15 ili 20 firmi. Maksimalan broj novih firmi određuje formalni readiness score.

Readiness mora najmanje da obuhvati:

- stabilnost kritičnog codebase-a;
- readiness svih obaveznih poslovnih tokova;
- PWA Otkupac, Vozač i Dispečer;
- sync, termalnu štampu, SEF i bankarske integracije;
- automatizaciju i trajanje onboardinga;
- kapacitet customer supporta i eskalacije;
- monitoring, recovery i release procese;
- dostupnost i logistiku hardvera;
- finansijsku rezervu i obrtni kapital;
- broj osoba koje mogu sprovesti standardan onboarding bez osnivača.

Cap određuje najslabija kritična komponenta, ne prosečna ocena.

### 7.7 Kontrolisan staged rollout

Velike promene prolaze kroz interni test, pilot firmu, ograničenu grupu i tek zatim pun rollout.

### 7.8 Bez rewrite-a bez merljivog razloga

Promena tehnološke platforme razmatra se kada postojeća platforma stvara merljiv limit u pouzdanosti, brzini razvoja, zapošljavanju, integracijama ili ukupnom trošku održavanja.

### 7.9 Operativna jednostavnost je funkcionalnost

Remote onboarding, monitoring, self-update, backup, kiosk konfiguracija, manuali i runbook-ovi imaju isti strateški značaj kao korisničke funkcije.

### 7.10 Nema prikrivenog custom developmenta

Zahtev jednog klijenta ulazi u proizvod samo kada predstavlja opšti problem segmenta i može se rešiti kroz zajednički model.

### 7.11 Ne obećavati budući proizvod kao postojeći

Modul se prodaje kao production tek kada zadovolji release kriterijume.

---

## 8. Strategija rasta 2026–2030

### Faza 1 — Dokaz readiness modela

Polazni komercijalni cilj je približno pet novih firmi, ali stvarni broj može biti 10, 15 ili 20 ako readiness score pokaže da sistem i organizacija mogu bezbedno da iznesu taj obim.

Ciljevi:

- standardizovati remote onboarding kroz manuale i checklistu;
- omogućiti da onboarding vodi customer support / implementation osoba;
- potvrditi PWA Otkupac, Vozač, Dispečer, kiosk i termalnu štampu;
- potvrditi SEF i bankarske tokove u realnoj upotrebi;
- meriti vreme po onboardingu i support case-u;
- potvrditi da broj novih firmi može rasti bez forkova i bez rasta incidenta po firmi;
- napraviti formalni readiness score pre početka aktivne prodaje.

### Faza 2 — Ubrzana nacionalna penetracija

**Okvir:** približno 10–50 firmi.

- support / implementation osoba preuzima standardna pitanja i onboarding;
- osnivač ostaje eskalacija za bugove i poslovnu logiku;
- osnivač se postepeno prebacuje na marketing, prodaju i partnerstva;
- developer se dodaje kada razvoj postane usko grlo;
- razvijaju se case studies, preporuke i direktna prodaja;
- potvrđuju se pricing i unit economics po segmentima.

### Faza 3 — Liderstvo u Srbiji

**Okvir:** približno 50–200 firmi.

- izgraditi najprepoznatljiviji specijalizovani brend za ceo otkupni biznis;
- organizovati support, implementaciju i razvoj tako da dnevni rad ne zavisi od osnivača;
- standardizovati hardver i širi IT katalog;
- razviti partnerstva sa knjigovođama, bankama, dobavljačima opreme i relevantnim organizacijama;
- potvrditi ili odbaciti ekonomiku Gazdinstva;
- pripremiti sistem za regionalnu ekspanziju.

### Faza 4 — Regionalna platforma

Regionalna platforma je cilj. Početna tržišta za procenu su Srbija, BiH, Crna Gora, Severna Makedonija i zatim Hrvatska i druga tržišta nakon pravne, jezičke, poreske i prodajne procene.

---

## 9. Strategija prihoda

Redosled važnosti u kratkom roku:

1. godišnje licence firmi za osnovni operativni sistem;
2. napredni moduli i paketi: Field, Logistika/Dispečer, Finansije/Banke, SEF i drugi;
3. implementacija i obuka kao odvojene usluge;
4. multi-company licence i premium support;
5. marža na terminalima i drugoj IT opremi;
6. Gazdinstvo Basic i Pro;
7. buduće integracije i SLA.

Pricing ne treba nužno da rascepa sistem na deset malih doplata. Packaging mora sačuvati jasnu vrednost celog sistema, uz mogućnost skupljih paketa za firme koje koriste pun operativni obim.

### Gazdinstvo

Gazdinstvo trenutno ne finansira osnovni biznis. Može postati glavni proizvod ili glavni izvor prihoda ako se potvrde aktivacija, konverzija, retencija, niska cena podrške i dodatni proizvodi koji povećavaju ARPU.

### Hardver i širi IT sistem

Hardver nije glavni profitni centar, ali treba da bude profitabilan sporedni centar. AgriX može postati dobavljač kompletnog IT sistema: kiosk tableti, termalni štampači, mrežna i rezervna oprema, računari, remote management i integracija perifernih sistema.

---

## 10. Strategija organizacije

### Osnivač

Osnivač zadržava product ownership, arhitekturu, ključni razvoj, finalnu eskalaciju, prodaju važnim klijentima, marketing strategiju i partnerstva.

### Prvo zaposlenje

Prva operativna osoba je customer support / implementation. Ona:

- rešava standardna i bazna pitanja;
- sprovodi remote onboarding prema manualima i checklistama;
- pomaže oko tableta, štampača i konfiguracije;
- prati monitoring;
- trijažira problem po poslovnom domenu;
- rešava poznate slučajeve prema runbook-u;
- eskalira složene slučajeve osnivaču ili developeru;
- vodi evidenciju vremena, uzroka i rešenja.

Osnivač pruža podršku van smene te osobe tokom sezone i rešava složene eskalacije.

### Dodatni developer

Developer se dodaje kada razvoj postane dokazano usko grlo, roadmap kasni, osnivač treba da se prebaci na marketing ili je trošak propuštenog rasta veći od punog troška developera.

---

## 11. Partner i kapital

AgriX ne treba partnera samo zbog kapitala. Partner ima smisla kada donosi distribuciju, direktan pristup kupcima, operativno vođenje prodaje i implementacije, iskustvo skaliranja B2B softvera, regionalnu mrežu ili kapital vezan za validiran plan ubrzanja.

Partner ili investicija mogu postati racionalni kada potražnja premaši kapacitet i kada kapital direktno uklanja dokazano usko grlo.

---

## 12. Strateška hitnost i tržišni prozor

`HYPOTHESIS`: tržište specijalizovanog softvera za otkup u Srbiji ima ograničen prozor u kojem AgriX može izgraditi dominantnu poziciju pre nego što postojeći ERP dobavljači ili novi vertikalni konkurent razviju sličan sistem.

AgriX ima jaču odbranu kada pokriva ceo posao nego kada prodaje samo otkupne listove. Što je više ključnih tokova zatvoreno u jednom sistemu, veća je korisnička vrednost, veći switching cost i teže je konkurentu da kopira ponudu.

`TARGET`: najmanje 200 firmi u periodu od 3–4 godine.

---

## 13. Strateški rizici

| Rizik | Verovatnoća | Uticaj | Primarna zaštita |
|---|---|---|---|
| Osnivač ostaje jedina osoba koja razume ceo sistem | Visoka | Visok | dokumentacija, support osoba, developer, ownership |
| Širina proizvoda postane prevelika za mali tim | Visoka | Visok | domeni, prioriteti, release gate i product ownership |
| Glavni tok radi, ali sporedni tokovi ostanu ručni | Srednja | Visok | end-to-end process mapping |
| Dispečer, vozači ili transport nisu dovoljno stabilni | Srednja | Kritičan | pilot i staged rollout |
| SEF ili banka naprave finansijski pogrešan rezultat | Srednja | Kritičan | validacija, audit trag, fail-closed i reconciliation |
| Readiness score preceni kapacitet | Srednja | Kritičan | weakest-link model i rezerva |
| Cena je niža od punog troška celog sistema | Srednja | Visok | unit economics i value-based packaging |
| Hardver veže previše kapitala | Srednja | Visok | predujam, standardni modeli, ograničena zaliha |
| Gazdinstvo nema dovoljnu konverziju | Visoka | Srednji | računati prihod kao nulu u konzervativnom planu |
| Rast bude prespor i konkurent zauzme tržište | Srednja/visoka | Kritičan | ambiciozan GTM i rast readiness-a |
| Agresivan rast pogorša kvalitet | Srednja | Kritičan | readiness-based cap i staged onboarding |

---

## 14. Ključni strateški KPI-jevi

- broj aktivnih firmi i novih firmi po sezoni;
- readiness score ukupno i po poslovnom domenu;
- procenat ključnih procesa potpuno zatvorenih u AgriX-u;
- broj ručnih prelaza i duplih unosa između procesa;
- onboarding sati po firmi;
- procenat onboardinga bez osnivača;
- support vreme po firmi i domenu;
- procenat problema rešenih bez osnivača;
- kritični incidenti i recovery vreme;
- uspešnost self-update rollouta;
- aktivni Field i Driver terminali;
- stopa neuspešne štampe;
- tačnost i vreme Dispečer planiranja;
- SEF uspešnost i broj neusaglašenih statusa;
- procenat automatski rasknjiženih bankarskih stavki;
- broj ručnih korekcija bankarskog mapiranja;
- broj i vrednost automatski pripremljenih naloga za plaćanje;
- ARR po firmi, modulu i ukupno;
- stvarna hardverska marža;
- renewal, churn i korišćenje modula;
- Gazdinstvo aktivacija, konverzija, ARPU i support cost.

---

## 15. Odobrene strateške odluke

### STR-001 — Readiness-based rast
Ne postoji unapred fiksiran sezonski hard cap. Maksimalan broj novih firmi određuje readiness score organizacije i celog proizvoda.

### STR-002 — Primarni tržišni fokus
Trenutni fokus je Srbija. Ciljni kupci su hladnjače i druge firme sa razgranatom mrežom stanica, vozača, kupaca i kooperanata.

### STR-003 — Jedan proizvod, jedan kod
Klijentske razlike rešavaju se zajedničkim kodom i konfiguracijom. Trajni klijentski fork nije dozvoljen.

### STR-004 — Dinamična uloga Gazdinstva
Licence firmi trenutno finansiraju osnovni biznis. Gazdinstvo može postati glavni proizvod ili prihod ako podaci to potvrde.

### STR-005 — Hardver kao sporedni profitni centar
Hardver mora imati pozitivnu stvarnu maržu. AgriX može postati dobavljač šireg IT sistema ciljnih klijenata.

### STR-006 — Bez partnera samo zbog novca
Partner ili investitor razmatra se kada rešava dokazano usko grlo i donosi merljivu sposobnost pored kapitala.

### STR-007 — Prvo operativno zaposlenje
Prva operativna uloga je customer support / implementation. Ta osoba rešava bazne slučajeve, sprovodi onboarding i delegira složene probleme.

### STR-008 — Regionalna platforma
Dugoročni cilj AgriX-a je regionalna vertikalna platforma.

### STR-009 — Strateški cilj tržišnog udela
AgriX cilja najmanje 200 firmi u naredne 3–4 godine.

### STR-010 — AgriX pokriva ceo poslovni sistem
AgriX se razvija i pozicionira kao end-to-end poslovni operativni sistem koji pokriva sve glavne i ključne sporedne tokove ciljnog klijenta. Desktop, PWA, Dispečer, Vozači, Gazdinstvo, SEF, banka i hardver predstavljaju povezane delove jednog sistema, ne zasebne nepovezane proizvode.

---

## 16. Otvorene teme za naredna poglavlja

1. Potvrditi procenu da u Srbiji postoji 500–1.000 relevantnih firmi.
2. Napraviti mapu svih glavnih i sporednih poslovnih tokova.
3. Za svaki tok označiti `Production`, `Pilot`, `Planned`, `Gap` ili `Out of scope`.
4. Definisati readiness score po poslovnim domenima.
5. Razložiti cilj od 200 firmi na godišnji prodajni i kadrovski plan.
6. Odrediti packaging celog sistema i premium modula.
7. Definisati koje IT kategorije AgriX prodaje, a koje ne.
8. Odrediti ARR i operativne pragove za partnera ili investitora.

---

## 17. Naredni koraci

1. upisati STR-010 u `DECISION_LOG.md`;
2. razviti `03_CUSTOMERS_AND_JOBS.md` po svim ulogama: vlasnik, administracija, otkupljivač, dispečer, vozač, kupac i kooperant;
3. razviti `07_PRODUCT_PORTFOLIO.md` kao mapu poslovnih domena i tokova;
4. razviti `04_MARKET.md` i potvrditi adresabilno tržište;
5. napraviti readiness model po domenima;
6. zatim finalizovati pricing, unit economics i plan rasta do 200 firmi.
