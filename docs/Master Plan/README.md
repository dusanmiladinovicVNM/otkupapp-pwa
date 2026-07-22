# AgriX Master Plan

**Status:** inicijalni kostur  
**Horizont:** 2026–2030  
**Namena:** interni strateški, finansijski i operativni sistem za vođenje AgriX-a  
**Način rada:** living documentation u Git-u; svako poglavlje se razvija kroz zaseban commit ili PR.

## Osnovna pravila

1. Odvojiti potvrđene činjenice, radne pretpostavke i tržišne hipoteze.
2. Ne koristiti hardver kao prihod bez paralelnog prikaza nabavne vrednosti, garancije, zalihe i marže.
3. Razdvojiti recurring prihod, implementaciju i hardver.
4. Svaka projekcija mora imati konzervativni, bazni i optimistični scenario.
5. Svako poglavlje mora završiti odlukama, KPI-jima, rizicima i narednim akcijama.
6. Master Plan nije marketing dokument; neprijatni zaključci imaju prednost nad motivacionim narativom.

## Potvrđene polazne činjenice

- Tri postojeća klijenta koriste desktop sistem; management koristi PWA za pregled i kontrolu.
- Plan za narednu sezonu je približno pet novih firmi, odnosno oko osam ukupno; preko deset nije cilj.
- Prosečna firma ima oko deset otkupnih stanica, jednog desktop korisnika i management PWA korisnike.
- Prosečna firma ima oko 100 kooperanata.
- Početna pretpostavka konverzije kooperanata na plaćeno Gazdinstvo je oko 5%.
- Gazdinstvo: planirane cene 19 EUR Basic i 39 EUR Pro; prvih 50 Partner naloga uključeno uz paket hladnjače.
- Onboarding je remote; trenutno realno jedan dan, sa ciljem od pola dana nakon standardizacije.
- Trenutno približno jedan support poziv nedeljno: oko 15 minuta za trivijalan slučaj ili više sati kada je bug.
- Funkcionalni self-update i granularni monitoring već postoje.
- Sve klijentske razlike ostaju u zajedničkom kodu, a aktiviraju se kroz konfiguraciju i podešavanja.
- Razvoj vodi osnivač; mogući naredni kapacitet je još jedan developer.
- Planirana je jedna osoba za customer support, uz podršku osnivača van njene smene tokom sezone.
- Za sledeću sezonu planirani su PWA Otkupac, kiosk tableti i termalna štampa na mestu otkupa.

## Struktura Master Plana

### 00 — Upravljanje dokumentom

`00_GOVERNANCE.md`

- svrha i publika dokumenta;
- vlasništvo nad poglavljima;
- statusi: Draft, Review, Approved, Superseded;
- ritam mesečnog i kvartalnog ažuriranja;
- pravila za pretpostavke, izvore i verzionisanje;
- decision log i change log;
- definicija osnovnog, konzervativnog i optimističnog scenarija.

### 01 — Executive Summary

`01_EXECUTIVE_SUMMARY.md`

- AgriX u jednoj stranici;
- problem, rešenje i ciljna grupa;
- trenutna faza proizvoda;
- poslovni model;
- ciljevi za 12, 24 i 48 meseci;
- ključni finansijski pokazatelji;
- najveći rizici i odluke koje trenutno nisu reverzibilne.

### 02 — Strategija i identitet kompanije

`02_STRATEGY.md`

- vizija, misija i strateški cilj do 2030;
- šta AgriX jeste, a šta namerno nije;
- vertikalna platforma naspram generičkog ERP-a;
- strateški principi razvoja;
- zajednički kod plus konfiguracija kao temelj skaliranja;
- bootstrapped rast naspram partnera/investitora;
- kriterijumi pod kojima bi partner imao smisla;
- geografski redosled širenja.

### 03 — Problem, kupci i segmentacija

`03_CUSTOMERS_AND_JOBS.md`

- vlasnici i management hladnjača;
- administrativni desktop korisnik;
- terenski otkupljivač;
- kooperant/gazdinstvo;
- jobs-to-be-done po ulozi;
- tipični procesi, bolne tačke i postojeća rešenja;
- segmentacija firmi po broju stanica, obimu i složenosti;
- ideal customer profile;
- anti-ICP: klijenti koje ne treba prihvatiti.

### 04 — Tržište

`04_MARKET.md`

- tržište otkupa u Srbiji;
- broj i vrste potencijalnih klijenata;
- TAM, SAM i realni SOM;
- sezonalnost i regionalna koncentracija;
- tržišna spremnost za desktop, PWA i kiosk terminale;
- tržište Gazdinstva;
- kasnije: BiH, Crna Gora, Severna Makedonija i Hrvatska;
- izvori podataka i nivo pouzdanosti svake procene.

### 05 — Konkurencija i alternative

`05_COMPETITION.md`

- BizniSoft, PANTHEON, Minimax i drugi ERP sistemi;
- specijalizovani programi za otkup;
- interna Excel rešenja, papir i ručni procesi;
- sopstveni razvoj kod hladnjače;
- funkcionalno i cenovno poređenje;
- switching costs;
- prednosti koje konkurencija može lako kopirati i prednosti koje teško kopira;
- realni razlozi zbog kojih kupac može odbiti AgriX.

### 06 — Pozicioniranje i ponuda vrednosti

`06_POSITIONING.md`

- centralna tržišna poruka;
- pozicioniranje po segmentima;
- merljiva vrednost za management, administraciju, otkupljivača i kooperanta;
- razlika između trenutnog i budućeg proizvoda;
- dozvoljene i zabranjene marketinške tvrdnje;
- dokazni materijal: reference, rezultati, studije slučaja i demonstracije.

### 07 — Portfolio proizvoda

`07_PRODUCT_PORTFOLIO.md`

- AgriX Desktop;
- Management PWA;
- PWA Otkupac;
- AgriX Gazdinstvo Partner, Basic i Pro;
- kiosk tablet i termalni štampač;
- banke, SEF, ambalaža, agrohemija, palete, izveštaji i ostali moduli;
- granice svakog proizvoda;
- zavisnosti između proizvoda;
- status: production, pilot, planned, deprecated.

### 08 — Product roadmap

`08_PRODUCT_ROADMAP.md`

- roadmap za narednih 12, 24 i 48 meseci;
- obavezno za sezonu, važno posle sezone i opcionalno;
- PWA Otkupac usklađivanje sa desktop tokom;
- pouzdana termalna štampa i kiosk upravljanje;
- Gazdinstvo onboarding i monetizacija;
- smanjenje tehničkog i operativnog rizika;
- kriterijumi za odlaganje funkcionalnosti;
- capacity budget osnivača i eventualnog developera.

### 09 — Tehnološka strategija

`09_TECHNOLOGY_STRATEGY.md`

- sadašnja arhitektura i razlozi njenog izbora;
- desktop, GAS/Sheets/PWA i integracije;
- self-update, monitoring, backup i recovery;
- konfiguracija klijenata i feature flags;
- sigurnost, privatnost i pristup podacima;
- SLA i RTO/RPO ciljevi;
- tehnički pragovi za migraciju sa trenutne platforme;
- pravilo bez rewrite-a bez merljivog razloga;
- plan za dodatnog developera, code ownership i review.

### 10 — Pricing i packaging

`10_PRICING_AND_PACKAGING.md`

- paketi za hladnjače prema realnosti tržišta Srbije;
- broj stanica kao glavna mera operativnog obima;
- desktop i management PWA;
- PWA Otkupac i terminali;
- prvih 50 Gazdinstvo Partner naloga;
- Basic 19 EUR i Pro 39 EUR;
- implementacija, obuka i posebni zahtevi;
- popusti, multi-company i founding-customer politika;
- godišnje korekcije cena;
- willingness-to-pay eksperimenti;
- pravila da se cena ne zasniva samo na prometu klijenta.

### 11 — Unit economics

`11_UNIT_ECONOMICS.md`

- ARR po hladnjači;
- prihod po stanici;
- prihod i konverzija Gazdinstva;
- CAC po kanalu;
- onboarding cost;
- support cost po firmi;
- gross margin po vrsti prihoda;
- LTV i payback period;
- hardware margin nakon rada, garancije i zamena;
- contribution margin po firmi i paketu;
- minimalna održiva cena.

### 12 — Finansijski model

`12_FINANCIAL_MODEL.md`

Za 5, 10, 20 i 50 hladnjača, uz poseban plan za narednu sezonu sa najviše oko 8–10 firmi:

- konzervativni, bazni i optimistični scenario;
- P&L;
- cash-flow;
- recurring i non-recurring prihodi;
- hardver: prodajna vrednost, nabavna vrednost, marža, zaliha i obrtni kapital;
- plate i pun trošak zaposlenih;
- hosting, alati, računovodstvo, osiguranje, pravni troškovi i banke;
- marketing i prodaja;
- putovanja, obuka i implementacija;
- rezervna oprema, reklamacije i garancija;
- porezi i PDV tretman kao posebne pretpostavke za proveru sa računovođom;
- osnivačka plata i dobit;
- break-even i runway;
- minimalna gotovinska rezerva pred sezonu;
- sensitivity analiza za cenu, prodaju, churn, support i konverziju Gazdinstva.

### 13 — Hardver i terenski terminali

`13_HARDWARE_OPERATIONS.md`

- standardni modeli tableta i printera;
- kiosk konfiguracija i daljinsko upravljanje;
- printer bridge i pouzdana štampa;
- nabavka, ulazna kontrola i konfiguracija;
- evidencija serijskih brojeva i dodela klijentu;
- minimalna zaliha i rezervni uređaji;
- garancija, kvar, lom, krađa i odgovornost;
- kupovina naspram najma;
- povrat i rashodovanje;
- realna marža nakon ukupnog rada i rizika;
- plan za 50, 100, 200 i 500 terminala.

### 14 — Go-to-market strategija

`14_GO_TO_MARKET.md`

- cilj od približno pet novih firmi u prvoj narednoj sezoni;
- redosled segmenata i kultura;
- SEO, Google Ads, direktni outreach i preporuke;
- sajmovi, udruženja, knjigovođe i partneri;
- demo i pilot model;
- case studies postojećih klijenata;
- budžet, očekivani leadovi i merila uspeha;
- kanali koje ne treba finansirati bez dokaza;
- prelazak osnivača sa razvoja ka marketingu nakon dodavanja developera.

### 15 — Sales playbook

`15_SALES_PLAYBOOK.md`

- kvalifikacija lead-a;
- discovery pitanja;
- priprema i vođenje demo prezentacije;
- ponuda, pregovori i ugovor;
- dokaz vrednosti bez nerealnih ROI tvrdnji;
- obrada prigovora;
- odluka pilot ili puna implementacija;
- prodajni pipeline i CRM minimum;
- prodajni ciklus i win/loss analiza;
- referral i upsell.

### 16 — Onboarding i implementacija

`16_ONBOARDING.md`

- remote onboarding kao podrazumevani model;
- pre-onboarding upitnik;
- priprema podataka i šifarnika;
- instalacija, konfiguracija i health check;
- obuka desktop, management i otkupac uloga;
- isporuka i test terminala;
- acceptance checklist;
- cilj skraćenja sa jednog na pola dana;
- prvih 7 i 30 dana;
- merila uspešnog onboardinga;
- automatizacije i dokumentacija potrebne da proces preuzme druga osoba.

### 17 — Customer support i customer success

`17_CUSTOMER_SUCCESS.md`

- radno vreme i sezonsko pokriće;
- jedna support osoba plus osnivač van smene;
- klasifikacija zahteva i prioriteti;
- maksimalna vremena odgovora;
- monitoring-first support;
- baza znanja i standardni odgovori;
- eskalacija bugova osnivaču/developeru;
- incident communication;
- obnova licenci, zadovoljstvo i preporuke;
- support capacity model za 5, 10, 20 i 50 firmi.

### 18 — Operacije, release i incidenti

`18_OPERATIONS.md`

- release ciklus i zabrana rizičnih promena usred sezone;
- pilot i staged rollout;
- self-update procedura;
- monitoring i alerting;
- backup, restore i disaster recovery;
- incident severity nivoi;
- postmortem bez okrivljavanja;
- sezonski readiness review;
- business continuity ako je osnivač nedostupan;
- service inventory i vlasništvo.

### 19 — Organizacija i zapošljavanje

`19_ORGANIZATION.md`

- trenutne uloge osnivača;
- customer support kao prva planirana operativna uloga;
- kriterijumi za dodatnog developera;
- kada je potrebna implementacija/teren osoba;
- kada prodaja ili marketing postaju zasebna uloga;
- opis posla, plata i pun trošak svake pozicije;
- sezonski i stalni angažman;
- delegation matrix;
- bus factor i transfer znanja;
- organizacioni modeli za 5, 10, 20 i 50 hladnjača.

### 20 — Pravni, poreski i bezbednosni okvir

`20_LEGAL_SECURITY_COMPLIANCE.md`

- ugovori, licenciranje i uslovi korišćenja;
- obrada podataka i privatnost;
- odgovornost za poslovne dokumente;
- SLA i ograničenje odgovornosti;
- hardverske garancije i fizička oštećenja;
- fiskalni, SEF i drugi regulatorni rizici;
- poreski i PDV model potvrđen sa stručnim licem;
- upravljanje korisnicima, tajnama i pristupom;
- incident sa podacima i obaveštavanje.

### 21 — Rizici i kontrole

`21_RISK_REGISTER.md`

- tržišni, prodajni, finansijski i likvidnosni rizici;
- tehnički i integracioni rizici;
- rizici sezone i koncentracije klijenata;
- zavisnost od osnivača;
- podrška, reputacija i reference;
- dobavljači hardvera i cloud platforme;
- pravni i regulatorni rizici;
- verovatnoća, uticaj, vlasnik i mitigacija;
- early-warning indikatori;
- kvartalni pregled top 10 rizika.

### 22 — KPI i CEO dashboard

`22_KPI_DASHBOARD.md`

- broj aktivnih hladnjača i stanica;
- ARR, renewal i churn;
- prihod i gross margin po izvoru;
- pipeline, conversion i CAC;
- onboarding vreme;
- support zahtevi i sati po firmi;
- critical incidents;
- uspešnost update-a i health status;
- PWA aktivnost;
- Gazdinstvo activation i paid conversion;
- founder time allocation;
- cash balance i runway;
- mesečni i sezonski dashboard.

### 23 — Scenario planovi

`23_SCENARIOS.md`

Za 5, 10, 20 i 50 hladnjača:

- prihod i marža;
- broj stanica, terminala i kooperanata;
- broj plaćenih Gazdinstvo korisnika;
- potreban tim;
- support i onboarding kapacitet;
- marketing i prodajni budžet;
- zaliha hardvera i obrtni kapital;
- ključni operativni rizici;
- tačke na kojima se rast namerno zaustavlja;
- uslovi za prelazak u sledeću fazu.

### 24 — Kapital, partnerstvo i vlasništvo

`24_CAPITAL_AND_PARTNERSHIP.md`

- bootstrapping kao osnovni scenario;
- koliko kapitala je stvarno potrebno po fazi;
- upotreba kapitala od 50.000 ili 100.000 EUR;
- zašto novac bez distribucije možda ne rešava usko grlo;
- kriterijumi za strateškog partnera;
- valuacija i dilution scenariji;
- kontrolna prava i zaštitne odredbe;
- debt, grant, leasing i revenue-financing alternative;
- go/no-go okvir za agresivno širenje.

### 25 — Decision log

`25_DECISION_LOG.md`

Za svaku veliku odluku:

- datum i status;
- kontekst;
- razmatrane alternative;
- odluka i razlog;
- očekivani efekat;
- rizik i trigger za ponovno razmatranje.

Početne teme: ostanak na VBA runtime-u, PWA, Gazdinstvo, zajednički kod plus konfiguracija, kiosk tableti, termalna štampa, pricing, maksimalno deset firmi naredne sezone i odluka o partneru.

### 26 — Izvori i data room

`26_SOURCES_AND_DATA_ROOM.md`

- tržišni izvori;
- cenovnici konkurencije;
- prodajni podaci;
- support i monitoring metrike;
- onboarding evidencija;
- finansijske pretpostavke;
- ugovori i šabloni;
- hardverske ponude;
- reference ka tehničkoj dokumentaciji u repozitorijumu;
- evidencija zastarelosti izvora.

## Redosled razvoja

Master Plan se ne piše numeričkim redom. Predloženi redosled je:

1. `00_GOVERNANCE.md` — pravila i model pretpostavki.
2. `02_STRATEGY.md` — strateške granice i cilj.
3. `03_CUSTOMERS_AND_JOBS.md` — kome se prodaje i koji problem se rešava.
4. `07_PRODUCT_PORTFOLIO.md` — precizan obim postojećeg i planiranog proizvoda.
5. `10_PRICING_AND_PACKAGING.md` — paketiranje i cene kao hipoteze za testiranje.
6. `11_UNIT_ECONOMICS.md` — stvarna ekonomika jednog klijenta.
7. `12_FINANCIAL_MODEL.md` — P&L i cash-flow scenariji.
8. `13_HARDWARE_OPERATIONS.md` — terminali, zaliha, garancija i obrtni kapital.
9. `16_ONBOARDING.md` i `17_CUSTOMER_SUCCESS.md` — kapacitet isporuke.
10. `23_SCENARIOS.md` — objedinjeni modeli za 5, 10, 20 i 50 firmi.
11. `14_GO_TO_MARKET.md` i `15_SALES_PLAYBOOK.md` — rast nakon potvrđene isporučivosti.
12. Ostala poglavlja, zatim `01_EXECUTIVE_SUMMARY.md` poslednje.

## Standard svakog poglavlja

Svaki dokument treba da sadrži:

1. cilj poglavlja;
2. potvrđene činjenice;
3. pretpostavke i nivo pouzdanosti;
4. analizu;
5. konzervativni, bazni i optimistični scenario gde je relevantno;
6. odluke koje proizlaze iz analize;
7. KPI-je i pragove;
8. rizike i mitigacije;
9. otvorena pitanja;
10. akcije sa vlasnikom i rokom;
11. izvore i datum poslednje provere;
12. change log.

## Statusna tabla

| Dokument | Prioritet | Status |
|---|---:|---|
| 00 Governance | P0 | nije započeto |
| 02 Strategy | P0 | nije započeto |
| 03 Customers and Jobs | P0 | nije započeto |
| 07 Product Portfolio | P0 | nije započeto |
| 10 Pricing and Packaging | P0 | nije započeto |
| 11 Unit Economics | P0 | nije započeto |
| 12 Financial Model | P0 | nije započeto |
| 13 Hardware Operations | P0 | nije započeto |
| 16 Onboarding | P1 | nije započeto |
| 17 Customer Success | P1 | nije započeto |
| 23 Scenarios | P1 | nije započeto |
| Ostala poglavlja | P2 | nije započeto |

## Prvi sledeći korak

Prvo se izrađuje `00_GOVERNANCE.md`, a odmah zatim `02_STRATEGY.md`. Tek kada se potvrde strateške granice, prelazi se na pricing i finansijski model. Time se sprečava da finansijske tabele budu precizna matematika zasnovana na pogrešnoj strategiji.
