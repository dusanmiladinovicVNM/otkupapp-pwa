# 00 — Upravljanje AgriX Master Planom

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-22  
**Sledeći pregled:** pre zaključavanja `02_STRATEGY.md`

---

## 1. Svrha

AgriX Master Plan je interni sistem za donošenje odluka o proizvodu, tržištu, cenama, finansijama, organizaciji i rastu. Nije investitorski pitch, marketing brošura niti skup optimističnih projekcija.

Njegov zadatak je da:

- pretvori rasute ideje i podatke u proverljive odluke;
- jasno odvoji činjenice, pretpostavke i hipoteze;
- pokaže finansijske i operativne posledice svake važne odluke;
- spreči rast koji prevazilazi kapacitet proizvoda i organizacije;
- omogući da novi saradnik razume zašto je odluka doneta;
- čuva istoriju promena bez prepravljanja prošlosti.

Kada je neprijatan zaključak bolje potkrepljen od poželjnog zaključka, neprijatan zaključak ima prednost.

---

## 2. Publika

### Primarna publika

- osnivač i vlasnik proizvoda;
- budući developer;
- customer support / implementation osoba;
- budući rukovodioci prodaje, operacija ili finansija.

### Sekundarna publika

Odabrani delovi mogu se koristiti za računovođu, pravnika, banku, potencijalnog partnera ili investitora. Interna verzija se ne prilagođava da bi izgledala privlačnije eksternoj publici.

---

## 3. Izvor istine

Master Plan je strateški izvor istine za poslovne odluke, ali nije izvor istine za tehničke detalje koji već imaju vlasničku dokumentaciju u repozitorijumu.

Prioritet izvora:

1. produkcijski podaci, ugovori, računi i stvarne metrike;
2. aktuelni kod, konfiguracija i operativni runbook-ovi;
3. potvrđene izjave klijenata i zapisnici razgovora;
4. zvanični javni izvori;
5. interne procene zasnovane na navedenoj metodologiji;
6. hipoteze koje tek treba testirati.

Kada su izvori u konfliktu, konflikt se navodi. Ne bira se automatski broj koji više odgovara željenoj priči.

---

## 4. Klasifikacija tvrdnji

Svaka važna tvrdnja, broj ili zaključak mora imati jednu od oznaka.

| Oznaka | Značenje | Dozvoljena upotreba |
|---|---|---|
| **FACT** | Direktno potvrđena činjenica | Može biti osnova odluke |
| **MEASURED** | Izmereno na ograničenom uzorku | Koristiti uz veličinu uzorka i period |
| **ASSUMPTION** | Radna pretpostavka | Obavezna sensitivity analiza |
| **HYPOTHESIS** | Tržišno ili produktno verovanje koje nije potvrđeno | Mora imati test i rok |
| **TARGET** | Željeni rezultat | Ne prikazivati kao prognozu |
| **DECISION** | Odobrena odluka | Mora imati datum i razlog |
| **UNKNOWN** | Nedostaje pouzdan podatak | Ne popunjavati izmišljenom preciznošću |

Primeri:

- `FACT`: postoje tri aktivna klijenta.
- `MEASURED`: približno jedan support poziv nedeljno tokom posmatranog perioda.
- `ASSUMPTION`: onboarding će nakon standardizacije trajati pola dana.
- `HYPOTHESIS`: 5% kooperanata će kupiti Gazdinstvo Basic ili Pro.
- `TARGET`: narednu sezonu završiti sa najviše 8–10 hladnjača.

---

## 5. Nivoi pouzdanosti

Za projekcije i tržišne procene koristi se ocena:

- **High** — više nezavisnih izvora ili dovoljno produkcijskih podataka;
- **Medium** — jedan kvalitetan izvor ili ograničen, ali relevantan uzorak;
- **Low** — rana procena sa malo podataka;
- **Speculative** — scenario za istraživanje, ne za budžetiranje.

Finansijski bazni scenario ne sme zavisiti od `Speculative` pretpostavke. Ako zavisi od `Low` pretpostavke, mora postojati konzervativna alternativa.

---

## 6. Statusi poglavlja

| Status | Značenje |
|---|---|
| **Planned** | Postoji samo u sadržaju |
| **Draft** | Početna verzija, nepotpuna ili neproverena |
| **Review** | Dovoljno kompletno za kritičku proveru |
| **Approved** | Osnivač ga prihvata kao trenutno važeću osnovu |
| **Needs Update** | Ranije odobreno, ali je važna pretpostavka zastarela |
| **Superseded** | Zamenjeno novijom odlukom ili dokumentom |
| **Archived** | Istorijski relevantno, više nije operativno |

Samo `Approved` poglavlje može biti obavezna osnova godišnjeg plana ili budžeta.

---

## 7. Definition of Done za poglavlje

Poglavlje može preći u `Approved` tek kada:

1. ima jasan cilj i poslovno pitanje koje rešava;
2. navodi potvrđene činjenice i izvore;
3. navodi sve materijalne pretpostavke;
4. razlikuje trenutno stanje od budućeg proizvoda;
5. sadrži hard-truth zaključke, ne samo prednosti;
6. kvantifikuje posledice gde je moguće;
7. sadrži rizike i kontraargumente;
8. definiše odluke, KPI-je i naredne akcije;
9. navodi šta još nije poznato;
10. dobije eksplicitno odobrenje osnivača.

Lep stil, dužina ili broj tabela nisu kriterijum kvaliteta.

---

## 8. Finansijska disciplina

Finansijski model mora odvojeno prikazivati:

### Recurring prihod

- godišnje licence hladnjača;
- Gazdinstvo Basic i Pro;
- održavanje ili recurring moduli;
- druge ponovljive pretplate.

### Non-recurring prihod

- onboarding;
- migracija;
- obuka;
- custom integracija;
- druga profesionalna usluga.

### Hardver

Za hardver se obavezno prikazuju:

- fakturisana prodajna vrednost;
- nabavna vrednost;
- transport i ulazni troškovi;
- konfiguracija i testiranje;
- rezervni uređaji;
- reklamacije, garancije i zamene;
- rizik zastarevanja zalihe;
- PDV i obrtni kapital;
- stvarna bruto i contribution marža.

Ukupna prodajna vrednost hardvera ne sme se predstavljati kao doprinos dobiti.

### Plate

Za svakog zaposlenog prikazuje se pun trošak poslodavca, ne samo neto plata. Plata osnivača se prikazuje čak i kada se privremeno ne isplaćuje, da bi model pokazao da li je poslovanje održivo bez besplatnog rada vlasnika.

### Porezi

Poreski i PDV tretman predstavljaju radnu pretpostavku dok ih ne potvrdi računovođa. Master Plan ne zamenjuje poresko ili pravno mišljenje.

---

## 9. Scenariji

Svaka važna finansijska projekcija ima najmanje tri scenarija.

### Konzervativni

- niža prodaja;
- sporija aktivacija Gazdinstva;
- veći support i onboarding trošak;
- više reklamacija hardvera;
- bez prihoda koji još nije potvrđen.

### Bazni

- najverovatniji ishod prema trenutno dostupnim podacima;
- bez nerealne konverzije ili zapošljavanja unapred;
- uključuje tržišnu platu osnivača u mature-state analizi.

### Optimistični

- bolja prodaja i konverzija, ali uz realan operativni kapacitet;
- nije fantasy scenario;
- mora objasniti šta konkretno mora biti tačno da bi se ostvario.

Pored toga, model mora imati operativne pragove za 5, 10, 20 i 50 hladnjača. Scenario sa 50 hladnjača nije prognoza, već test potrebne organizacije i kapitala.

---

## 10. Vremenski horizonti

Svako strateško poglavlje razlikuje:

- **sada:** aktuelna produkcija i tri postojeća klijenta;
- **naredna sezona:** cilj približno osam, uz maksimum oko deset hladnjača;
- **12–24 meseca:** validacija skaliranja, pricinga i PWA Otkupac modela;
- **24–48 meseci:** mogući rast prema 20 i više hladnjača;
- **2030 horizont:** pravac, ne obećanje.

Ne sme se mešati funkcionalnost koja postoji danas sa onom koja je tek planirana za narednu sezonu.

---

## 11. Upravljanje pretpostavkama

Svaka materijalna pretpostavka treba da sadrži:

| Polje | Sadržaj |
|---|---|
| ID | npr. `ASM-PRC-001` |
| Tvrdnja | Šta pretpostavljamo |
| Vlasnik | Ko proverava |
| Pouzdanost | High/Medium/Low/Speculative |
| Uticaj | Na koje odluke utiče |
| Test | Kako se proverava |
| Rok | Kada mora biti proverena |
| Rezultat | Confirmed/Rejected/Revised/Open |

Primer početnih pretpostavki:

- `ASM-GAZ-001`: oko 5% kooperanata prelazi na plaćeni Gazdinstvo nalog.
- `ASM-ONB-001`: standardizovan onboarding može pasti na pola dana.
- `ASM-PRC-001`: paket za oko deset stanica može se prodavati u planiranom cenovnom rasponu Srbije.
- `ASM-SUP-001`: trenutna stopa supporta ostaje približno stabilna sa većim brojem firmi.

Pretpostavka koja nije testirana u dogovorenom roku prelazi u `UNKNOWN`, ne ostaje neograničeno u baznom scenariju.

---

## 12. Decision log

Materijalne odluke se zapisuju u `DECISION_LOG.md`.

Obavezna polja:

- ID odluke;
- datum;
- status;
- kontekst;
- razmatrane opcije;
- odluka;
- razlog;
- očekivane posledice;
- rizici;
- uslovi za ponovno otvaranje odluke;
- povezani dokumenti ili commit-i.

Odluka se ne briše kada se promeni. Dobija status `Superseded` i vezu ka novoj odluci.

---

## 13. Change log

Bitne promene Master Plana vode se u `CHANGELOG.md`.

Promena se smatra bitnom kada menja:

- ciljnu grupu;
- paket ili cenu;
- plan prihoda i troškova;
- prag zapošljavanja;
- proizvodni roadmap;
- strategiju finansiranja;
- toleranciju rizika;
- geografski fokus.

Sitne jezičke i formaterske izmene ne zahtevaju poseban zapis.

---

## 14. Ritam ažuriranja

### Mesečno

- broj aktivnih firmi;
- pipeline i zaključeni ugovori;
- support incidenti i utrošeno vreme;
- onboarding vreme;
- troškovi alata i infrastrukture;
- korišćenje PWA i Gazdinstva;
- otvorene pretpostavke kojima ističe rok.

### Kvartalno

- pricing i prodajna konverzija;
- P&L i cash-flow odstupanja;
- roadmap i kapacitet;
- hiring pragovi;
- ključni rizici;
- odluke koje treba potvrditi ili promeniti.

### Pred sezonu

Obavezan pregled:

- release readiness;
- podrška i dežurstva;
- hardverska zaliha;
- incident plan;
- cash rezerva;
- maksimalan broj novih klijenata;
- zabrana rizičnih promena tokom kritičnog perioda.

### Posle sezone

- post-mortem bez traženja krivca;
- stvarni support i incident load;
- profitabilnost po klijentu i paketu;
- churn i zadovoljstvo;
- validacija pretpostavki;
- odluka o narednom kapacitetu.

---

## 15. Vlasništvo i prava odlučivanja

Dok je AgriX founder-led kompanija:

- osnivač je finalni vlasnik strategije, proizvoda, pricinga i kapitala;
- developer daje tehničku procenu i može blokirati release zbog dokumentovanog kritičnog rizika;
- customer support/implementation osoba poseduje operativne metrike i povratne informacije klijenata;
- računovođa potvrđuje poreske i računovodstvene pretpostavke;
- pravnik potvrđuje ugovore, privatnost, odgovornost i garancijske uslove.

Konsenzus nije obavezan. Neslaganje mora biti zapisano kada je materijalno.

---

## 16. Pravila protiv fluff-a

U Master Planu nisu dozvoljene tvrdnje poput:

- „tržište je ogromno“ bez izračunatog dostupnog segmenta;
- „AgriX štedi mnogo vremena“ bez definisanog procesa i merenja;
- „proizvod se lako skalira“ bez support, onboarding i release podataka;
- „Gazdinstvo ima veliki potencijal“ kao zamena za model konverzije;
- „partner ubrzava rast“ bez plana korišćenja kapitala;
- „promet klijenta opravdava cenu“ bez willingness-to-pay dokaza;
- „hardver donosi X prihoda“ bez nabavne vrednosti i marže;
- „osnivač može sve sam“ bez capacity modela i cene njegovog vremena.

Svaka velika prednost mora imati odgovarajući rizik ili ograničenje.

---

## 17. Pravila za korišćenje AI analiza

AI može pomagati u istraživanju, strukturiranju, modeliranju i kritici, ali:

- AI procena nije izvor sama po sebi;
- broj bez izvora ili modela dobija oznaku `ASSUMPTION`;
- AI ne odobrava poglavlje;
- finansijski, pravni i poreski zaključci zahtevaju stručnu proveru;
- motivacioni ton ne povećava pouzdanost zaključka;
- konflikt između AI procene i produkcijskih podataka rešava se u korist podataka.

---

## 18. Reversibility klasifikacija odluka

Odluke se dele na:

- **Type 1 — teško reverzibilne:** prodaja udela, veliki kredit, zapošljavanje velikog tima, dugoročni hardverski ugovor, rewrite platforme;
- **Type 2 — reverzibilne:** test cene, pilot paket, marketinški kanal, privremeni feature flag;
- **Type 3 — operativne:** redovne male odluke bez strateškog uticaja.

Type 1 odluka zahteva pisani decision record, downside scenario i jasno definisane uslove pod kojima se donosi.

---

## 19. Prag materijalnosti

Promena mora ući u Master Plan kada može da utiče na najmanje jedno od sledećeg:

- više od 5% godišnjeg prihoda ili troška;
- više od 40 sati rada osnivača godišnje;
- produkcijski rizik za više klijenata;
- cenu, ugovor ili obećani SLA;
- zapošljavanje ili prodaju udela;
- obradu podataka i pravnu odgovornost;
- strateški položaj proizvoda.

---

## 20. Početni registar poznatih činjenica

| ID | Tvrdnja | Tip | Pouzdanost |
|---|---|---|---|
| F-001 | Tri firme trenutno koriste desktop sistem | FACT | High |
| F-002 | Management koristi PWA za pregled i kontrolu | FACT | High |
| F-003 | Prosečno oko deset otkupnih stanica po firmi | MEASURED | Medium |
| F-004 | Prosečno oko 100 kooperanata po firmi | MEASURED | Medium |
| F-005 | Jedan desktop korisnik po firmi | MEASURED | Medium |
| F-006 | Onboarding je do sada izveden remote | FACT | High |
| F-007 | Trenutni onboarding traje približno jedan dan | MEASURED | Medium |
| F-008 | Self-update i granularni monitoring postoje | FACT | High |
| F-009 | Klijentske razlike rešavaju se zajedničkim kodom i konfiguracijom | FACT | High |
| F-010 | Sledeća sezona cilja oko osam, uz maksimum oko deset firmi | TARGET | High |
| F-011 | Gazdinstvo Basic/Pro planirani su na 19/39 EUR | DECISION-DRAFT | Medium |
| F-012 | Prvih 50 Partner naloga ide uz paket hladnjače | DECISION-DRAFT | Medium |

Ovaj registar nije zamena za detaljna poglavlja. On samo sprečava da se početne činjenice izgube ili promene bez traga.

---

## 21. Otvorena governance pitanja

Pre odobrenja ovog poglavlja treba odlučiti:

1. Da li se Master Plan vodi na srpskom kao jedinom službenom jeziku?
2. Da li finansijski model koristi EUR kao osnovnu plansku valutu, uz RSD poreske i gotovinske tokove?
3. Da li se svako poglavlje razvija u zasebnom PR-u ili više povezanih poglavlja može deliti PR?
4. Koji podaci o klijentima smeju biti navedeni, a koji moraju biti anonimizovani?
5. Da li se formira zaseban privatni repozitorijum ako finansijski i ugovorni podaci postanu previše osetljivi za postojeći repo?

---

## 22. Odluke koje ovaj dokument odmah uvodi

- Master Plan se vodi kao living documentation u Git-u.
- Činjenice, pretpostavke, hipoteze i ciljevi moraju biti eksplicitno razdvojeni.
- Hardware revenue se nikada ne posmatra bez hardware cost-a i stvarne marže.
- Executive Summary se piše poslednji.
- Scenario sa 50 hladnjača služi za capacity i organization test, ne kao automatska prognoza.
- Nijedna Type 1 odluka ne donosi se na osnovu jednog optimističnog scenarija.
- Konačno odobrenje svakog poglavlja daje osnivač.

---

## 23. Naredne akcije

1. Pregledati i odobriti ili izmeniti pet otvorenih governance pitanja.
2. Kreirati `DECISION_LOG.md` i `CHANGELOG.md` kada se potvrdi format.
3. Zatim izraditi `02_STRATEGY.md`.
4. `01_EXECUTIVE_SUMMARY.md` ostaviti za kraj prve pune verzije Master Plana.
