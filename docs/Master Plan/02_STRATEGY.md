# 02 — Strategija i identitet AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-22

---

## 1. Polazna tačka

AgriX je vertikalni poslovni operativni sistem za organizovani otkup poljoprivrednih proizvoda i povezano upravljanje gazdinstvom. Proizvod ne pokriva samo evidenciju otkupa, već povezuje kooperante, parcele, repromaterijal, terenski rad, prijem robe, logistiku, transport, lager, prodaju, dokumentaciju, finansije, regulatorne obaveze, management kontrolu i farm-management funkcije.

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
- `FACT`: management rola ima Dispečer modul za real-time planiranje i praćenje otkupa i transporta;
- `FACT`: postoji posebna operativna rola za vozače;
- `FACT`: AgriX podržava SEF tok;
- `FACT`: AgriX automatski preuzima i obrađuje bankarske izvode, rasknjižava stavke na kooperante i kupce i priprema naloge za plaćanje;
- `FACT`: management modul omogućava QR identifikaciju kooperanta, izbor parcela, preporuku količine agrohemije, skeniranje barkoda artikala, formiranje korpe i izdavanje otpremnice;
- `FACT`: Kooperant rola ima karticu prema hladnjači sa zaduženjima, razduženjima i saldom;
- `FACT`: Kooperant rola ima GIS prikaz parcela, parcelnu meteo prognozu, rizike, upozorenja i termine pogodne za tretmane;
- `FACT`: Kooperant rola ima pametno doziranje, evidenciju tretmana i opreme;
- `FACT`: Kooperant rola ima unos i kategorizaciju troškova, raspodelu po parcelama i sezonski bilans ukupno i po parceli;
- `HYPOTHESIS`: u Srbiji postoji približno 500–1.000 relevantnih firmi; procena mora biti potvrđena tržišnim istraživanjem;
- `TARGET`: AgriX treba da osvoji najmanje 200 firmi u naredne 3–4 godine, ako readiness i tržišna validacija to podrže.

AgriX više nije prototip, ali još nije dokazano skaliran proizvod. Sledeće faze moraju dokazati da ceo sistem, onboarding, support, hardver i organizacija mogu da rastu bez pada pouzdanosti.

---

## 2. Strateška definicija

AgriX nije samo program za otkup i nije skup nepovezanih modula. AgriX je **end-to-end vertikalni poslovni operativni sistem** sa dva povezana jezgra:

1. **AgriX Enterprise** — kompletan operativni sistem za hladnjače i druge firme koje organizuju otkup;
2. **AgriX Gazdinstvo** — farm-management i relationship sistem za kooperante.

Ta dva jezgra dele podatke i poslovni kontekst, ali mogu imati različitu ekonomiku, pakete i dugoročne puteve rasta.

Strateški cilj je da AgriX pokrije sve glavne i ključne sporedne tokove ciljnog klijenta: od parcele i izdavanja repromaterijala, preko otkupa i transporta, do otpreme, fakture, SEF-a, banke, isplate kooperanta, naplate kupca i konačne management kontrole.

### 2.1 AgriX Enterprise — poslovni domeni

#### 1. Kooperanti, parcele i priprema sezone

- matični podaci kooperanata;
- parcele, kulture, površine i proizvodni kontekst;
- otkupne stanice, otkupljivači, vozači, kupci, artikli i cenovnici;
- pravila, dozvole, role i konfiguracija;
- digitalna veza firme sa kooperantom.

#### 2. Repromaterijal i podrška proizvodnji

- agrohemija i drugi artikli;
- QR identifikacija kooperanta;
- izbor parcele ili parcela;
- preporučena količina prema površini, dozi i pakovanju;
- skeniranje EAN, Code i QR barkodova kamerom;
- formiranje korpe, količina, cena i ukupne vrednosti;
- izdavanje otpremnice i potpisi;
- evidencija izdate robe po kooperantu, parceli i artiklu;
- povezivanje sa lagerom, karticom i zaduženjem kada poslovni model to zahteva.

#### 3. Terenski otkup i prijem robe

- PWA Otkupac;
- unos na otkupnoj stanici;
- kiosk tablet i termalna štampa na licu mesta;
- otkupni listovi, otpremnice, zbirne, prijemnice i prateća dokumentacija;
- kvalitet, klase, bruto/neto, ambalaža i povezani podaci;
- offline-first rad i naknadna sinhronizacija.

#### 4. Logistika, Dispečer i vozači

- real-time pregled otkupljene i neraspoređene robe;
- planiranje preuzimanja po stanicama;
- raspodela kamiona i vozača;
- kapaciteti, rute i statusi transporta;
- PWA tok za vozača;
- praćenje izvršenja i zatvaranje transporta;
- organizacija isporuke kupcima kada je primenljivo.

#### 5. Lager, ambalaža, palete, prerada i sledljivost

- ulaz, izlaz i stanje robe i repromaterijala;
- ambalaža, reversi i zaduženja;
- paletni i proizvodni tokovi;
- prerada i povezivanje izvora i izlaza;
- sledljivost od kooperanta i stanice do kupca;
- operativni dokumenti, kontrole i storno tokovi.

#### 6. Kupci, prodaja i otprema

- kupci i komercijalni uslovi;
- izdavanje robe iz lagera;
- korpe, otpremnice i isporuke;
- fakture i potraživanja;
- veza prodaje sa robom, sledljivošću i finansijama.

#### 7. SEF i regulatorni tok

- priprema elektronske fakture;
- validacija i mapiranje;
- slanje na SEF;
- praćenje workflow i remote statusa;
- audit trag, recovery i kontrola neusaglašenosti.

#### 8. Finansije i trezor

- obaveze prema kooperantima;
- zaduženja za repromaterijal i druge stavke;
- potraživanja od kupaca;
- avansi, uplate, isplate i salda;
- automatsko preuzimanje bankarskih izvoda;
- automatsko ili kontrolisano rasknjižavanje na kooperante i kupce;
- priprema naloga za plaćanje;
- kontrola preplata, duplikata i neusaglašenih transakcija.

#### 9. Management, kontrola i monitoring

- real-time pregled poslovanja;
- Dispečer i operativna kontrola mreže;
- KPI, izveštaji i analitika;
- management operacije putem PWA;
- granularni tehnički monitoring;
- audit trag, storno, korekcije i odgovornost;
- korisnici, role i pristup.

### 2.2 AgriX Gazdinstvo — farm-management domeni

Gazdinstvo nije samo portal za pregled odnosa sa hladnjačom. To je zaseban operativni proizvod za planiranje, izvršenje i ekonomsku kontrolu proizvodnje.

#### 1. Kartica prema hladnjači

- dokumenti i stavke kartice;
- zaduženja;
- razduženja;
- tekući saldo;
- pregled finansijskog odnosa sa hladnjačom.

#### 2. Parcele i GIS

- spisak svih parcela;
- kultura, površina, katastarski i GGAP podaci;
- poligoni i lokacije na satelitskoj mapi;
- pregled i fokus pojedinačne parcele;
- osnova za parcelno vezivanje tretmana, troškova i rezultata.

#### 3. Realna parcelna prognoza i upozorenja

- prognoza vezana za konkretnu parcelu;
- temperatura, vlaga, vetar, padavine i drugi relevantni parametri;
- rizici i aktivna upozorenja;
- povoljni i nepovoljni termini za prskanje;
- upozorenja po kulturi i parceli;
- management odluke zasnovane na lokaciji, ne na opštoj prognozi grada.

#### 4. Digitalni agronom i pametno doziranje

- izbor parcele, kulture, mere i artikla;
- preporučena doza prema površini i karakteristikama artikla;
- stanje sopstvenog lagera agrohemije;
- karenca i relevantna upozorenja;
- oprema: traktor, prskalica i druga sredstva;
- vreme rada, lokacija i meteo snapshot;
- evidencija izvršenih tretmana;
- offline-first rad i sinhronizacija.

#### 5. Knjiga polja i proizvodnja

- istorija radova i tretmana;
- proizvodnja i otkup povezani sa parcelom;
- utrošena agrohemija;
- radni sati i korišćena oprema;
- pregled po parceli, kulturi i sezoni.

#### 6. Troškovi

- ručni unos troškova;
- kategorije: gorivo, popravke, osiguranje, sertifikacija, analize, navodnjavanje, ambalaža, radna snaga, zakup, transport i ostalo;
- opšti troškovi ili vezivanje za konkretnu parcelu;
- automatski obračun pojedinih troškova, uključujući radnu snagu kada je konfigurisan;
- pregled ukupno i po kategoriji.

#### 7. Sezonski bilans

- vrednost proizvodnje;
- trošak agrohemije;
- ostali troškovi;
- radni sati;
- rezultat sezone;
- bilans celog gazdinstva;
- bilans pojedinačne parcele;
- osnova za poređenje kultura, parcela i sezona.

#### 8. Dodatni farm-management tokovi

- evidencija opreme;
- fiskalni računi i privatne evidencije;
- preporuke i dnevni pregled;
- sync status i rad bez stabilne veze;
- buduće funkcije koje povećavaju ekonomsku i agronomsku vrednost po gazdinstvu.

### 2.3 Kanali pristupa

Aplikacije i uređaji su kanali pristupa jedinstvenom sistemu:

- **Excel/VBA Desktop** — centralni master, dokumentacija, finansije, SEF i složeni back-office tokovi;
- **Management PWA** — pregled, Dispečer, kontrola, repromaterijal i operativno odlučivanje;
- **PWA Otkupac** — rad na otkupnoj stanici;
- **PWA Vozač** — transportni zadaci i statusi;
- **Gazdinstvo PWA** — farm-management i odnos kooperanta sa hladnjačom;
- **kamera mobilnog uređaja** — QR i barkod skeniranje;
- **kiosk tableti i termalni štampači** — standardizovan terenski rad;
- **integracije** — SEF, banke i budući eksterni sistemi.

AgriX se prodaje kao jedan povezan poslovni sistem, a ne kao kolekcija aplikacija.

---

## 3. Vizija

Do 2030. AgriX treba da bude vodeći regionalni poslovni operativni sistem za organizovani otkup i povezano upravljanje gazdinstvima, sa snažnom bazom u Srbiji i prenosivim modelom za tržišta regiona.

`TARGET`: u naredne 3–4 godine izgraditi bazu od najmanje 200 firmi, uz rast koji je dozvoljen readiness-om, a ne proizvoljnim godišnjim plafonom.

Vizija podrazumeva da AgriX:

- pokriva ciklus saradnje sa kooperantom od parcele i repromaterijala do otkupa i isplate;
- pokriva poslovni ciklus firme od otkupa do transporta, otpreme, fakture i naplate;
- povezuje centralu, stanice, management, dispečere, vozače, kupce i kooperante;
- daje kooperantu ozbiljan farm-management alat, a ne samo portal;
- zatvara regulatorne i finansijske tokove kroz SEF i bankarske integracije;
- smanjuje ručni rad, dupli unos i kašnjenje informacija;
- koristi kameru, barkodove, QR, GIS, meteo i offline rad kao prirodne operativne alate;
- ostaje standardizovan proizvod sa jednim kodom;
- može da opslužuje stotine firmi i veliki broj gazdinstava kroz automatizaciju i delegiranje;
- stvara dovoljno recurring prihoda za stalni razvoj, podršku i regionalno širenje;
- postane prirodni tehnološki dobavljač ciljnom segmentu.

Cilj od 200 firmi je strateški cilj, ne finansijska prognoza. Mora se razložiti na godišnje akvizicione, operativne, kadrovske i finansijske kapacitete.

---

## 4. Misija

AgriX omogućava hladnjačama i drugim organizovanim otkupljivačima da vode celokupan operativni posao kao jedan povezan sistem: kooperante, parcele, repromaterijal, otkup, stanice, dokumentaciju, ambalažu, transport, vozače, dispečersko planiranje, lager, kupce, otpremu, fakture, SEF, banku, isplate, naplate i management kontrolu.

Istovremeno, AgriX Gazdinstvo omogućava kooperantu da vodi parcele, tretmane, pametno doziranje, opremu, troškove, proizvodnju, prognozu, upozorenja, sezonski bilans i karticu prema hladnjači u jednom sistemu.

---

## 5. Šta AgriX jeste, a šta nije

AgriX jeste:

- vertikalni poslovni operativni sistem;
- end-to-end platforma za glavne i ključne sporedne tokove otkupnog biznisa;
- farm-management platforma za kooperante;
- B2B2C ekosistem koji povezuje firmu i gazdinstvo;
- zajednički proizvod za veliki broj firmi bez forkova;
- sistem koji se prilagođava konfiguracijom;
- integracioni sloj prema SEF-u, bankama i drugim relevantnim sistemima;
- potencijalni dobavljač šireg IT sistema za ciljne klijente.

AgriX namerno nije:

- generički ERP za sve industrije;
- univerzalni knjigovodstveni program;
- custom software studio sa posebnom verzijom za svakog klijenta;
- jeftin program samo za štampanje otkupnih listova;
- jednostavan portal za kooperante bez stvarne operativne vrednosti;
- klasičan hardverski distributer bez sopstvene tehnološke vrednosti;
- proizvod koji obećava funkcionalnosti koje nisu production-ready;
- projekat koji menja tehnologiju bez merljivog razloga.

AgriX ne mora da zameni svaki računovodstveni program, ali treba da bude primarni operativni sistem klijenta i da zatvori tokove specifične za otkupni biznis.

---

## 6. Ciljno tržište i idealni kupac

Trenutni geografski fokus je Srbija.

Primarni B2B kupci su:

- hladnjače sa sopstvenom mrežom otkupnih stanica;
- firme koje se bave organizovanim otkupom i imaju razgranatu mrežu stanica i kooperanata;
- firme koje kooperantima izdaju agrohemiju, ambalažu ili drugi repromaterijal;
- firme sa sopstvenom logistikom, vozačima ili dispečerskom potrebom;
- firme kojima generički ERP ne rešava ulaz robe, sledljivost, logistiku i finansijsko zatvaranje otkupa;
- firme koje žele centralnu kontrolu terena, lagera, kupaca, naplate i isplate.

Tipičan idealni kupac ima više stanica, veliki broj kooperanata, značajan obim dokumenata i finansijskih transakcija, odgovornu osobu za implementaciju i spremnost da standardizuje proces bez forka.

Promet od 1–2 miliona EUR jeste čest profil sadašnjih klijenata, ali broj stanica, logistička složenost, broj kooperanata, repromaterijal, obim dokumenata i potreba za kontrolom važniji su od samog prometa.

Primarni B2C/B2B2C korisnici su kooperanti koji imaju više parcela, koriste agrohemiju, žele ekonomsku kontrolu proizvodnje ili žele transparentan odnos sa hladnjačom.

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

AgriX mora pokriti i sporedne tokove bez kojih glavni proces ostaje nedovršen. Otkup nije zatvoren bez transporta, prijema, otpreme, fakture, SEF-a, banke i isplate. Odnos sa kooperantom nije zatvoren bez parcela, repromaterijala, troškova, tretmana, kartice i bilansa.

### 7.4 Jedinstven podatak kroz ceo tok

Podatak nastao na parceli, pri izdavanju repromaterijala ili na stanici treba da se koristi dalje u lageru, logistici, finansijama, izveštajima, SEF-u i banci bez ponovnog ručnog unosa. Dupli unos je signal da proces nije potpuno zatvoren.

### 7.5 Mobilni uređaj je operativni terminal

Kamera, QR, barkod, digitalni potpis, GIS, geolokacija i PWA nisu pomoćni UX detalji. Oni omogućavaju da se transakcija evidentira na mestu nastanka.

### 7.6 Podatak po parceli je strateška imovina

Parcelno povezivanje prognoze, tretmana, troškova, proizvodnje i rezultata stvara vrednost koju generički evidencioni sistemi teško kopiraju.

### 7.7 Pouzdanost i brzina rasta nisu suprotnosti

AgriX automatizacijom, monitoringom, self-update-om, standardizovanim onboardingom i delegiranjem povećava brzinu rasta bez pada pouzdanosti.

### 7.8 Sezonski cap određuje readiness

Ne postoji unapred fiksiran hard cap. Maksimalan broj novih firmi određuje formalni readiness score.

Readiness mora obuhvatiti:

- stabilnost kritičnog codebase-a;
- readiness svih obaveznih poslovnih tokova;
- PWA Otkupac, Vozač, Dispečer, repromaterijal i Gazdinstvo;
- barkod/QR, GIS, meteo i offline sync;
- termalnu štampu, SEF i bankarske integracije;
- automatizaciju i trajanje onboardinga;
- kapacitet supporta i eskalacije;
- monitoring, recovery i release procese;
- logistiku hardvera;
- finansijsku rezervu i obrtni kapital;
- broj osoba koje mogu sprovesti standardan onboarding bez osnivača.

Cap određuje najslabija kritična komponenta, ne prosečna ocena.

### 7.9 Kontrolisan staged rollout

Velike promene prolaze kroz interni test, pilot firmu, ograničenu grupu i tek zatim pun rollout.

### 7.10 Bez rewrite-a bez merljivog razloga

Promena platforme razmatra se kada postojeća tehnologija stvara merljiv limit u pouzdanosti, brzini razvoja, zapošljavanju, integracijama ili trošku održavanja.

### 7.11 Operativna jednostavnost je funkcionalnost

Remote onboarding, monitoring, self-update, backup, kiosk konfiguracija, manuali i runbook-ovi imaju isti strateški značaj kao korisničke funkcije.

### 7.12 Nema prikrivenog custom developmenta

Zahtev jednog klijenta ulazi u proizvod samo kada predstavlja opšti problem segmenta i može se rešiti kroz zajednički model.

### 7.13 Ne obećavati budući proizvod kao postojeći

Modul se prodaje kao production tek kada zadovolji release kriterijume.

---

## 8. Strategija rasta 2026–2030

### Faza 1 — Dokaz readiness modela

Polazni komercijalni cilj je približno pet novih firmi, ali stvarni broj može biti 10, 15 ili 20 ako readiness pokaže da sistem i organizacija mogu bezbedno da iznesu obim.

Ciljevi:

- standardizovati remote onboarding kroz manuale i checkliste;
- omogućiti da onboarding vodi customer support / implementation osoba;
- potvrditi PWA Otkupac, Vozač, Dispečer, repromaterijal, kiosk i termalnu štampu;
- potvrditi SEF i bankarske tokove u realnoj upotrebi;
- definisati status i production readiness Gazdinstvo funkcija;
- meriti vreme po onboardingu i support case-u;
- potvrditi rast bez forkova i bez rasta incidenta po firmi;
- napraviti formalni readiness score pre aktivne prodaje.

### Faza 2 — Ubrzana nacionalna penetracija

**Okvir:** približno 10–50 firmi.

- support / implementation osoba preuzima standardna pitanja i onboarding;
- osnivač ostaje eskalacija za bugove i poslovnu logiku;
- osnivač se prebacuje na marketing, prodaju i partnerstva;
- developer se dodaje kada razvoj postane usko grlo;
- razvijaju se case studies, preporuke i direktna prodaja;
- potvrđuju se pricing i unit economics Enterprise i Gazdinstvo proizvoda.

### Faza 3 — Liderstvo u Srbiji

**Okvir:** približno 50–200 firmi.

- izgraditi najprepoznatljiviji specijalizovani brend za ceo otkupni biznis;
- organizovati support, implementaciju i razvoj tako da dnevni rad ne zavisi od osnivača;
- standardizovati hardver i širi IT katalog;
- razviti partnerstva sa knjigovođama, bankama, agronomima, dobavljačima opreme i relevantnim organizacijama;
- dokazati ili odbaciti ekonomiku Gazdinstva;
- pripremiti sistem za regionalnu ekspanziju.

### Faza 4 — Regionalna platforma

Regionalna platforma je cilj. Početna tržišta za procenu su Srbija, BiH, Crna Gora, Severna Makedonija i zatim Hrvatska i druga tržišta nakon pravne, jezičke, poreske i prodajne procene.

---

## 9. Strategija prihoda

Kratkoročni izvori prihoda:

1. godišnje licence firmi za AgriX Enterprise;
2. paketi i moduli: Field, Logistika/Dispečer, Repromaterijal, Finansije/Banke, SEF i drugi;
3. implementacija i obuka;
4. multi-company licence i premium support;
5. marža na terminalima i drugoj IT opremi;
6. Gazdinstvo Basic i Pro;
7. buduće integracije i SLA.

Pricing ne treba da rascepa sistem na deset sitnih doplata. Packaging mora sačuvati vrednost celog sistema, uz skuplje pakete za firme koje koriste veći operativni obim.

### 9.1 Gazdinstvo ekonomika

Gazdinstvo trenutno ne finansira osnovni biznis, ali funkcionalna širina opravdava da se tretira kao ozbiljan zaseban proizvod.

Može postati glavni proizvod ili prihod ako se potvrde:

- aktivacija kooperanata;
- učestalo korišćenje tokom sezone i van sezone;
- retencija;
- willingness-to-pay za Basic i Pro;
- niska cena podrške;
- dodatni prihod od premium agronomskih, finansijskih ili tržišnih funkcija;
- mrežni efekat kroz hladnjače i njihove kooperante.

### 9.2 Hardver i širi IT sistem

Hardver nije glavni profitni centar, ali treba da bude profitabilan sporedni centar. AgriX može postati dobavljač kiosk tableta, termalnih štampača, uređaja sa pouzdanom kamerom, mrežne i rezervne opreme, računara, remote managementa i integracije perifernih sistema.

---

## 10. Strategija organizacije

### Osnivač

Osnivač zadržava product ownership, arhitekturu, ključni razvoj, finalnu eskalaciju, prodaju važnim klijentima, marketing strategiju i partnerstva.

### Prvo zaposlenje

Prva operativna osoba je customer support / implementation. Ona:

- rešava standardna i bazna pitanja;
- sprovodi onboarding prema manualima i checklistama;
- pomaže oko tableta, kamera, skeniranja, štampača i konfiguracije;
- prati monitoring;
- trijažira problem po poslovnom domenu i roli;
- rešava poznate slučajeve prema runbook-u;
- eskalira složene slučajeve;
- vodi evidenciju vremena, uzroka i rešenja.

Osnivač pruža podršku van smene te osobe tokom sezone i rešava složene eskalacije.

### Dodatni developer

Developer se dodaje kada razvoj postane dokazano usko grlo, roadmap kasni, osnivač treba da se prebaci na marketing ili je trošak propuštenog rasta veći od punog troška developera.

---

## 11. Partner i kapital

AgriX ne treba partnera samo zbog kapitala. Partner ima smisla kada donosi distribuciju, pristup kupcima, operativno vođenje prodaje i implementacije, iskustvo skaliranja B2B/B2B2C softvera, regionalnu mrežu ili kapital vezan za validiran plan ubrzanja.

Partner ili investicija mogu postati racionalni kada potražnja premaši kapacitet i kada kapital direktno uklanja dokazano usko grlo.

---

## 12. Strateška hitnost i tržišni prozor

`HYPOTHESIS`: tržište ima ograničen prozor u kojem AgriX može izgraditi dominantnu poziciju pre nego što postojeći ERP dobavljači ili novi vertikalni konkurent razviju sličan sistem.

AgriX ima jaču odbranu kada pokriva ceo posao firme i daje kooperantu pun farm-management proizvod. Što je više ključnih tokova zatvoreno u jednom sistemu, veća je korisnička vrednost, veći switching cost i teže je konkurentu da kopira ponudu.

`TARGET`: najmanje 200 firmi u periodu od 3–4 godine.

---

## 13. Strateški rizici

| Rizik | Verovatnoća | Uticaj | Primarna zaštita |
|---|---|---|---|
| Osnivač ostaje jedina osoba koja razume ceo sistem | Visoka | Visok | dokumentacija, support osoba, developer, ownership |
| Širina proizvoda postane prevelika za mali tim | Visoka | Visok | domeni, statusi, prioriteti i release gate |
| Funkcije postoje, ali nisu povezane u zatvoren tok | Srednja | Visok | end-to-end process mapping |
| Dispečer, vozači ili transport nisu dovoljno stabilni | Srednja | Kritičan | pilot i staged rollout |
| SEF ili banka proizvedu finansijski pogrešan rezultat | Srednja | Kritičan | validacija, audit, fail-closed i reconciliation |
| GIS/meteo preporuka bude protumačena kao stručna garancija | Srednja | Visok | jasna ograničenja, izvori i upozorenja |
| Gazdinstvo ima mnogo funkcija, ali nisku aktivaciju | Visoka | Visok | product analytics, onboarding i test monetizacije |
| Readiness score preceni kapacitet | Srednja | Kritičan | weakest-link model i rezerva |
| Cena je niža od punog troška celog sistema | Srednja | Visok | unit economics i value-based packaging |
| Hardver veže previše kapitala | Srednja | Visok | predujam, standardni modeli i ograničena zaliha |
| Rast bude prespor i konkurent zauzme tržište | Srednja/visoka | Kritičan | ambiciozan GTM i rast readiness-a |
| Agresivan rast pogorša kvalitet | Srednja | Kritičan | readiness-based cap i staged onboarding |

---

## 14. Ključni strateški KPI-jevi

### Enterprise

- broj aktivnih firmi i novih firmi po sezoni;
- readiness score ukupno i po poslovnom domenu;
- procenat ključnih procesa potpuno zatvorenih u AgriX-u;
- broj ručnih prelaza i duplih unosa;
- onboarding sati po firmi;
- procenat onboardinga bez osnivača;
- support vreme po firmi i domenu;
- kritični incidenti i recovery vreme;
- aktivni Field i Driver terminali;
- Dispečer korišćenje i uspešnost planova;
- stopa neuspešne štampe;
- SEF uspešnost i neusaglašeni statusi;
- procenat automatski rasknjiženih bankarskih stavki;
- broj ručnih korekcija mapiranja;
- broj i vrednost pripremljenih naloga za plaćanje;
- ARR po firmi, modulu i ukupno;
- hardverska marža;
- renewal i churn.

### Gazdinstvo

- broj Partner, Basic i Pro naloga;
- aktivacija po hladnjači;
- mesečno i nedeljno aktivni korisnici;
- broj aktivnih parcela;
- broj otvorenih GIS/meteo pregleda;
- broj evidentiranih tretmana;
- korišćenje pametnog doziranja;
- broj unetih troškova i procenat vezan za parcelu;
- broj korisnika koji pregledaju sezonski bilans;
- broj pregleda kartice prema hladnjači;
- konverzija Partner → Basic/Pro;
- ARPU, retencija i support cost.

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
AgriX se razvija i pozicionira kao end-to-end poslovni operativni sistem koji pokriva sve glavne i ključne sporedne tokove ciljne firme. Desktop, PWA, Dispečer, Vozači, repromaterijal, lager, prodaja, SEF, banka, monitoring i hardver predstavljaju povezane delove jednog sistema.

### STR-011 — Gazdinstvo je pun farm-management proizvod
AgriX Gazdinstvo se ne tretira samo kao portal ili dodatak hladnjači. To je zaseban farm-management proizvod koji pokriva karticu prema hladnjači, parcele i GIS, prognozu i upozorenja, pametno doziranje, tretmane, troškove, proizvodnju i sezonski bilans ukupno i po parceli.

---

## 16. Otvorene teme za naredna poglavlja

1. Potvrditi procenu da u Srbiji postoji 500–1.000 relevantnih firmi.
2. Napraviti mapu svih glavnih i sporednih poslovnih tokova.
3. Za svaki tok označiti `Production`, `Pilot`, `Planned`, `Gap` ili `Out of scope`.
4. Definisati readiness score po poslovnim domenima.
5. Razložiti cilj od 200 firmi na godišnji prodajni i kadrovski plan.
6. Definisati Gazdinstvo Partner, Basic i Pro granice prema stvarnim funkcijama.
7. Definisati koje agronomske preporuke su informativne, a koje zahtevaju stručnu validaciju.
8. Odrediti packaging Enterprise sistema i premium modula.
9. Definisati koje IT kategorije AgriX prodaje, a koje ne.
10. Odrediti ARR i operativne pragove za partnera ili investitora.

---

## 17. Naredni koraci

1. upisati STR-010 i STR-011 u `DECISION_LOG.md`;
2. razviti `03_CUSTOMERS_AND_JOBS.md` po svim ulogama: vlasnik, administracija, otkupljivač, dispečer, vozač, magacioner, kupac i kooperant;
3. razviti `07_PRODUCT_PORTFOLIO.md` kao mapu poslovnih domena i tokova;
4. razviti `04_MARKET.md` i potvrditi adresabilno tržište;
5. napraviti readiness model po domenima;
6. zatim finalizovati pricing, unit economics i plan rasta do 200 firmi.
