# AgriX CRM Pipeline Playbook

**Status:** DRAFT v1 — VALIDATION  
**Datum:** 30.07.2026.  
**Revizija:** 02.08.2026. — usklađen kanonski stage model i dodat sales motion  
**Vlasnik:** osnivač AgriX-a  
**Svrha:** Uvesti jedinstven, proverljiv i operativan standard za evidentiranje naloga, kontakata, prilika, aktivnosti, forecast-a, no-deal ishoda i predaje dobijenih poslova implementaciji.

---

## 1. Osnovni princip

CRM nije arhiva kontakata niti zbir beleški. CRM je operativni sistem komercijalnog procesa i mora u svakom trenutku da pokaže:

- ko je kupac i koje firme/objekti pripadaju istom poslovnom sistemu;
- ko učestvuje u odluci i kakvu ulogu ima;
- koji potvrđeni problem se rešava;
- u kojoj fazi se prilika zaista nalazi;
- koji sales motion se koristi;
- šta je sledeći korak, ko ga vodi i kada se dešava;
- koji rizik može da zaustavi posao;
- koliko je forecast zasnovan na dokazima;
- zašto je posao dobijen, izgubljen ili odložen;
- šta implementacija mora da zna pre početka rada.

**DECISION:** Faza se ne određuje po tome koliko je razgovora održano, već po ispunjenim exit kriterijumima.

**DECISION:** `04_SALES_PROCESS.md` je kanonski izvor za nazive i značenje opportunity faza. CRM ne uvodi paralelni stage model.

---

## 2. Objekti i relacije

### 2.1. Account

Predstavlja pravno ili poslovno lice sa kojim AgriX ima ili može imati odnos.

Obavezna polja:

- naziv;
- PIB/matični broj kada je poznat i dozvoljen za evidenciju;
- mesto i region;
- tip poslovanja;
- dominantne kulture/proizvodi;
- broj stanica/lokacija kada je poznat;
- sezonski prozor;
- izvor naloga;
- ICP tier;
- owner naloga;
- lifecycle status;
- poslednja relevantna aktivnost;
- sledeći planirani kontakt.

Povezani account-i se ne dupliraju. Matična firma, povezana firma, hladnjača, otkupna stanica i druga instanca povezuju se relacijom.

### 2.2. Contact

Predstavlja konkretnu osobu.

Obavezna polja:

- ime i prezime;
- funkcija;
- account;
- kontakt kanal;
- uloga u buying committee-ju;
- nivo uticaja;
- stav prema promeni;
- preferirani način komunikacije;
- dozvola/status komunikacije gde je relevantno;
- poslednji kontakt;
- sledeći korak.

Buying role vrednosti:

- economic buyer;
- decision maker;
- champion;
- operational owner;
- end user;
- finance evaluator;
- technical evaluator;
- legal/procurement;
- blocker;
- influencer;
- unknown.

### 2.3. Lead

Lead je nepotvrđena osoba ili firma sa mogućim interesom. Lead nije opportunity.

Minimalni statusi:

- New;
- Researching;
- Contact Attempted;
- Connected;
- Qualified to Account;
- Nurture;
- Disqualified.

Lead se konvertuje kada postoji stvarni account, relevantan kontakt i razlog za dalji komercijalni rad.

### 2.4. Opportunity

Opportunity postoji tek kada su potvrđeni:

- stvarna firma i relevantan kontakt;
- poslovni problem ili cilj;
- razlog za promenu ili rok;
- realan sledeći korak.

Obavezna osnovna polja:

- naziv opportunity-ja;
- account;
- owner;
- stage;
- `sales_motion`;
- source;
- primary use case;
- problem statement;
- consequence;
- expected scope;
- procenjena vrednost;
- expected close date;
- season/implementation deadline;
- economic buyer;
- champion status;
- decision process;
- competition/status quo;
- top risk;
- next step;
- next step date;
- last meaningful activity;
- forecast category;
- confidence evidence;
- loss/no-decision reason kada se zatvori.

### 2.5. Activity

Standardni tipovi:

- email;
- call;
- meeting;
- discovery;
- demo;
- technical validation;
- scope review;
- proposal review;
- negotiation;
- site visit;
- implementation handoff;
- customer success review;
- internal task.

Aktivnost mora imati ishod, ne samo oznaku da se dogodila.

### 2.6. Quote / Proposal

Ponuda mora biti vezana za jednu opportunity i verzionisana.

Minimalna polja:

- verzija;
- datum;
- scope;
- vrednost;
- važenje;
- izvor cenovnika;
- status;
- razlog izmene;
- ko je odobrio odstupanje scope-a ili uslova.

### 2.7. Implementation Handoff

Kreira se tek za S8 Closed Won i sadrži komercijalno-tehnički kontekst potreban implementaciji.

---

## 3. Lifecycle status account-a

| Status | Značenje |
|---|---|
| Target | odgovara ICP-u, kontakt još nije potvrđen |
| Prospect | postoji relevantan kontakt ili signal |
| Active Opportunity | postoji najmanje jedna otvorena opportunity |
| Customer | postoji aktivan ugovorni odnos |
| Expansion | aktivna prilika za dodatni scope |
| Nurture | trenutno nema aktivnog procesa, ali postoji budući razlog |
| Dormant | nema potvrđenog razloga ni roka |
| Disqualified | ne odgovara ICP-u ili nije prihvatljiv nalog |

Account lifecycle ne zamenjuje opportunity stage.

---

## 4. Kanonski opportunity stage model

CRM koristi potpuno isti model kao `04_SALES_PROCESS.md`:

| Faza | Naziv | Suština |
|---|---|---|
| S0 | Target Account | nalog odgovara ICP-u, ali nema potvrđenog kontakta ili potrebe |
| S1 | Connected | uspostavljen je relevantan dvosmerni kontakt |
| S2 | Qualified Problem | potvrđeni su problem, posledica, owner i vremenski kontekst |
| S3 | Discovery | mapirani su proces, stakeholderi, success criteria i rizici |
| S4 | Solution Evaluation | AgriX se proverava prema konkretnim procesima i kriterijumima |
| S5 | Risk Alignment | usaglašavaju se scope, implementacija, migracija, obuka, podrška i tehnički uslovi |
| S6 | Proposal Review | konačni scope i komercijalni model predstavljeni su relevantnim ljudima |
| S7 | Decision / Approval | kupac završava odobrenje, komercijalni review, ugovor i commitment |
| S8 | Closed Won | postoji formalna odluka i prihvaćen handoff |
| S9 | Closed Lost | izabran je konkurent, status quo ili drugi eksplicitan negativan ishod |
| SN | Nurture | postoji fit/potencijal, ali nema aktivnog projekta ili roka |

**DECISION:** Aktivni forecast obuhvata samo S2–S7. S0, S1 i SN nisu aktivan prihod u forecast-u. S8 i S9 su zatvoreni ishodi.

### 4.1. Entry/exit minimum

#### S0 — Target Account

**Entry:** potvrđen nalog i potencijalni use case.  
**Exit:** postoji relevantan kontakt i plan prvog razgovora.

Obavezno:

- source;
- ICP tier;
- hypothesized use case;
- next step.

#### S1 — Connected

**Entry:** ostvaren dvosmerni kontakt.  
**Exit:** potvrđen smislen razlog za kvalifikaciju/discovery ili je nalog prebačen u SN/S9.

Obavezno:

- contact role;
- osnovni trigger ili signal;
- trenutni pristup;
- sledeći sastanak ili jasan ishod.

#### S2 — Qualified Problem

**Entry:** potvrđeni problem/cilj, consequence, približan rok i relevantan sagovornik.  
**Exit:** zakazan discovery sa pravim učesnicima.

Minimalni PACT:

- Problem;
- Authority path;
- Consequence;
- Timing.

#### S3 — Discovery

**Entry:** kvalifikovan problem i dogovoren discovery.  
**Exit:** dokumentovani proces, posledice, success criteria, rizici, buying committee i dogovoren demo/tehnički korak.

Opportunity se ne pomera dalje ako postoji samo lista funkcija koje kupac želi.

#### S4 — Solution Evaluation

**Entry:** demo, workflow review ili tehnička validacija povezani su sa potvrđenim problemima.  
**Exit:** potvrđen fit, evidentirani gap-ovi, definisan preliminarni scope i dogovoren risk/scope korak.

#### S5 — Risk Alignment

**Entry:** funkcionalni fit je potvrđen i postoji realna namera za procenu uvođenja.  
**Exit:** scope, implementation pretpostavke, odgovornosti i kritični rizici dovoljno su stabilni za ponudu; proposal review je zakazan.

#### S6 — Proposal Review

**Entry:** ponuda je predstavljena uživo ili na strukturisanom pozivu; samo slanje dokumenta nije dovoljno.  
**Exit:** evidentirane su reakcije, otvorena pitanja, approval path, decision date i sledeći korak.

#### S7 — Decision / Approval

**Entry:** proposal review je završen i kupac aktivno rešava preostala odobrenja, uslove ili ugovor.  
**Exit:** formalno prihvatanje vodi u S8; eksplicitan negativan ishod vodi u S9; legitimno odlaganje sa budućim triggerom vodi u SN.

#### S8 — Closed Won

Obavezno:

- konačni scope;
- ugovorena vrednost;
- datum odluke;
- razlog dobitka;
- ključni dokaz;
- konkurencija/status quo;
- implementation owner;
- prihvaćen handoff;
- kickoff ili prvi onboarding korak.

#### S9 — Closed Lost

Obavezno:

- outcome subtype;
- primary reason;
- root cause;
- poslednja relevantna faza;
- competitor/status quo;
- lesson learned;
- reactivation datum samo kada je stvarno legitiman.

Outcome subtype:

- Lost to competitor;
- No Decision / Status Quo;
- Disqualified;
- AgriX Withdrawn;
- Project Cancelled;
- Timing Missed;
- Other — explanation required.

#### SN — Nurture

Koristi se kada postoji fit ili potencijal, ali nema aktivnog projekta, prioriteta ili roka.

Obavezno:

- nurture reason;
- future trigger;
- reactivation datum/prozor;
- relevantna persona;
- owner;
- sledeći smislen sadržaj ili kontakt.

Nurture bez triggera i datuma je Dormant, ne aktivni proces.

---

## 5. Migracija prethodnih CRM oznaka

Prethodna verzija CRM dokumenta koristila je neke iste oznake sa drugačijim značenjem. Sledeća mapa čuva sve prethodne koncepte, ali ih više ne tretira kao zasebne stage-ove.

| Prethodna CRM oznaka/koncept | Novi kanonski zapis |
|---|---|
| S0 Identified | S0 Target Account |
| S2 Qualified | S2 Qualified Problem |
| S3 Discovery Complete | S3 Discovery + milestone `DISCOVERY_COMPLETE` |
| S4 Solution Validated | S4 Solution Evaluation + milestone `SOLUTION_VALIDATED` |
| S5 Scope Confirmed | S5 Risk Alignment + `commercial_substatus = SCOPE_CONFIRMED` |
| S6 Proposal Presented | S6 Proposal Review + `proposal_status = PRESENTED` |
| S7 Commercial Review | S7 Decision / Approval + `commercial_substatus = COMMERCIAL_REVIEW` |
| S8 Commit / Contracting | S7 Decision / Approval + `forecast_category = Commit` + `commercial_substatus = CONTRACTING` |
| S9 Closed Won | S8 Closed Won |
| SN Closed Lost / No Decision | S9 Closed Lost sa outcome subtype-om |
| Deferred | SN Nurture sa potvrđenim triggerom i datumom |

### 5.1. Commercial substatus

Da se ne izgubi detalj kasne faze, S5–S7 mogu koristiti opciono polje `commercial_substatus`:

- SCOPE_DRAFT;
- SCOPE_CONFIRMED;
- IMPLEMENTATION_OUTLINE_CONFIRMED;
- PROPOSAL_PREPARED;
- PROPOSAL_PRESENTED;
- COMMERCIAL_REVIEW;
- LEGAL_REVIEW;
- PROCUREMENT;
- CONTRACTING;
- SIGNATURE_PENDING.

Substatus ne menja stage i ne koristi se za ulepšavanje forecast-a.

---

## 6. Sales motion

Svaka opportunity ima polje `sales_motion`:

- `FAST_TRACK`;
- `STANDARD`;
- `COMPLEX`.

### FAST_TRACK

Koristi se samo kada su ispunjeni kriterijumi iz `04A_FAST_TRACK_SALES_MOTION.md`.

Fast Track:

- koristi iste S0–S9/SN faze;
- može završiti više faza istog dana;
- ne preskače entry/exit dokaze;
- automatski prelazi u Standard/Complex kada se pojavi složenost;
- ne dobija posebne forecast kategorije niti poseban cenovnik.

### STANDARD

Podrazumevani konsultativni proces iz `04_SALES_PROCESS.md` za tipične B2B prilike.

### COMPLEX

Koristi se za više firmi/instanci, složene integracije, migracije, custom razvoj, formalni procurement, poseban SLA ili visok implementation/vendor rizik.

Promena motion-a evidentira:

- prethodni motion;
- novi motion;
- datum;
- razlog;
- uticaj na scope, close date i implementation deadline.

---

## 7. Progressive CRM fields

Sva polja ne moraju biti poznata pri kreiranju opportunity-ja. Obaveznost raste sa fazom.

### Do S2

- account;
- relevantan kontakt;
- owner;
- stage;
- sales motion;
- source;
- trigger;
- problem statement;
- consequence;
- problem owner;
- okvirni timing;
- okvirni scope;
- next step, owner i date.

### Do izlaska iz S3

Dodati:

- current-state process summary;
- 3–5 success criteria;
- buying committee;
- economic buyer poznat/nepoznat;
- champion status;
- current solution/ERP;
- implementation deadline;
- decision process — poznat/nepoznat;
- top risks.

### Do izlaska iz S4

Dodati:

- fit conclusion;
- gap classification;
- proof koji je kupac video;
- broj firmi, instanci, stanica, korisnika i uređaja;
- paket/moduli — preliminarno;
- technical/operational evaluator;
- sledeći risk/scope korak.

### Do izlaska iz S5

Dodati:

- finalni preliminarni scope;
- implementation outline;
- migraciju;
- integracije;
- obuku;
- odgovornosti obe strane;
- out-of-scope;
- capacity status;
- procenjenu vrednost prema cenovniku;
- proposal review datum.

### Do izlaska iz S6

Dodati:

- proposal version;
- finalni scope i cena;
- izvor cenovnika;
- reakcije kupca;
- otvorena komercijalna/pravna pitanja;
- approval path;
- decision date;
- competition/status quo;
- forecast category i evidence.

### Pre S8

Dodati:

- formalno prihvatanje;
- ugovorenu vrednost;
- potpisnika;
- project owner-e;
- handoff paket;
- kickoff/prvi onboarding korak;
- reason won i ključni dokaz.

Polje bez informacije označava se `UNKNOWN`, ne popunjava pretpostavkom.

---

## 8. Stage discipline

Opportunity ne sme napredovati zbog:

- lepog razgovora;
- verbalnog interesa bez sledećeg koraka;
- održanog demo-a bez potvrđenog fit-a;
- poslate ponude bez review sastanka;
- neodređenog „javićemo se“;
- procene prodavca da će posao verovatno biti dobijen;
- Fast Track oznake;
- želje da pipeline izgleda naprednije.

Kada se otkrije da prethodni kriterijum nije ispunjen, opportunity se vraća u odgovarajuću fazu.

---

## 9. Next-step standard

Svaka otvorena opportunity mora imati:

- konkretan sledeći korak;
- datum;
- vlasnika;
- drugu stranu koja je prihvatila korak;
- očekivani rezultat;
- stage evidence koji korak treba da završi.

Loše:

- pratiti;
- javiti se;
- poslati više informacija;
- čekamo odgovor.

Dobro:

- 4. avgusta vlasnik i operativa potvrđuju broj stanica i obim migracije;
- AgriX do 6. avgusta dostavlja reviziju scope-a, kupac je razmatra 8. avgusta;
- tehnička validacija izvoza podataka zakazana za 12. avgust sa internim IT kontaktom.

**DECISION:** Opportunity bez dogovorenog sledećeg koraka nije aktivan forecast signal.

---

## 10. Meaningful activity

Meaningful activity menja razumevanje, odluku ili napredak. Primeri:

- potvrđen problem;
- identifikovan decision maker;
- dobijen podatak za scope;
- završen demo sa fit zaključkom;
- rešen prigovor;
- potvrđen decision process;
- dogovoren rok;
- kupac dostavio podatke;
- ponuda pregledana;
- ugovorni uslov usaglašen.

Automatski email, pokušaj poziva i poruka bez odgovora ne resetuju stage aging kao meaningful activity.

---

## 11. Aging i stagnacija

Početni pragovi za validaciju:

| Stage | Warning | Critical |
|---|---:|---:|
| S0–S1 | 10 dana | 20 dana |
| S2 | 14 dana | 30 dana |
| S3–S4 | 21 dan | 45 dana |
| S5 | 21 dan | 45 dana |
| S6 | 14 dana bez potvrđenog decision koraka | 30 dana |
| S7 | prema decision date-u | propušten decision date bez novog dokaza |

S8 i S9 su zatvoreni i nemaju stage aging. SN se prati prema reactivation datumu, ne prema aktivnom pipeline aging-u.

Fast Track može imati kraće realno trajanje, ali koristi iste warning/critical principe dok se ne validira dovoljan uzorak.

Aging se tumači u odnosu na sezonski kontekst. Duga prilika sa potvrđenim future triggerom prelazi u SN, ne ostaje veštački otvorena.

Stalled opportunity zahteva jedno od:

- novi mutual next step;
- povratak u raniju fazu;
- promena sales motion-a;
- SN Nurture;
- S9 Closed Lost;
- disqualification.

---

## 12. Forecast kategorije

### Pipeline

- tipično S2–S5;
- postoji kvalifikovana prilika;
- nema dovoljno dokaza da će odluka biti završena u posmatranom periodu.

### Best Case

- tipično S6–S7;
- fit, scope i proces odluke su dovoljno poznati;
- realna je mogućnost odluke u periodu;
- ostaje najmanje jedan važan rizik.

### Commit

Dozvoljeno samo u S7 kada postoje:

- potvrđen economic buyer;
- potvrđen finalni scope i cena;
- rešeni materijalni prigovori;
- poznat proces odobrenja/potpisa;
- konkretan datum odluke;
- kupčev eksplicitni commitment;
- nema nepoznatog kritičnog veto faktora;
- realan implementation slot.

`commercial_substatus = CONTRACTING` sam po sebi nije dovoljan za Commit.

### Closed

- S8 Closed Won;
- S9 Closed Lost sa outcome subtype-om.

SN nije Closed niti aktivni forecast.

**PROHIBITED:** Forecast se ne zasniva na osećaju, simpatiji kupca, brzini Fast Track-a ili količini komunikacije.

---

## 13. Opportunity confidence score

Interni score 0–10:

- problem potvrđen: 1;
- posledica potvrđena: 1;
- timing/trigger potvrđen: 1;
- champion potvrđen ponašanjem: 1;
- economic buyer uključen: 1;
- decision process poznat: 1;
- scope potvrđen: 1;
- fit validiran: 1;
- komercijalni uslovi potvrđeni: 1;
- mutual next step sa datumom: 1.

Score ne zamenjuje stage. Koristi se za proveru kvaliteta i otkrivanje slabih prilika.

Fast Track ne dobija bonus bodove zbog kraćeg trajanja.

---

## 14. Champion test

Kontakt je champion samo ako:

- interno objašnjava problem i vrednost;
- daje informacije koje nisu javne;
- povezuje AgriX sa drugim učesnicima;
- pomaže da se razume proces odluke;
- preuzima dogovorene aktivnosti;
- upozorava na rizike i blokere;
- ima lični ili poslovni razlog da promena uspe.

Ljubazan i zainteresovan kontakt nije automatski champion.

Champion status:

- Unknown;
- Potential;
- Tested;
- Confirmed;
- Lost.

---

## 15. Close date pravila

Expected close date mora biti izveden iz kupčevog procesa, ne iz želje prodavca.

Izvor može biti:

- datum upravnog sastanka;
- budžetski rok;
- sezonski rok;
- datum isteka postojećeg ugovora;
- planirani početak implementacije;
- potvrđen procurement/contracting kalendar.

Kada datum prođe:

1. utvrditi šta se promenilo;
2. evidentirati razlog pomeranja;
3. postaviti novi datum samo uz dokaz;
4. promeniti forecast kategoriju;
5. promeniti motion kada je složenost porasla;
6. vratiti stage kada buyer task nije završen;
7. zatvoriti ili prebaciti u SN kada nema realnog osnova.

---

## 16. Pipeline hygiene

Nedeljna revizija proverava:

- opportunity bez next step-a;
- next step u prošlosti;
- close date u prošlosti;
- critical aging;
- stage bez obaveznih polja;
- duplirane account-e/kontakte;
- opportunity bez economic buyer-a posle S4;
- S6 bez proposal review-a;
- S7 Commit bez dokaza;
- neaktivne prilike koje treba zatvoriti;
- vrednost koja nije usklađena sa scope-om i cenovnikom;
- izgubljene poslove bez razloga i beleške;
- Fast Track koji više ne zadovoljava eligibility;
- sales motion bez obrazloženja.

Mesečna revizija dodatno proverava:

- stage conversion;
- velocity;
- source quality;
- forecast accuracy;
- conversion i pre-sales sate po sales motion-u;
- najčešće loss/no-decision razloge;
- promene close date-a;
- stale nurture naloge;
- kvalitet CRM unosa po owner-u.

---

## 17. No-deal i closure standard

Obavezni primary reason:

- No budget;
- No priority;
- No decision/status quo;
- Lost to competitor;
- Build internally;
- Existing ERP/solution retained;
- Missing critical capability;
- Timing/season missed;
- Authority/process unavailable;
- Commercial terms;
- Implementation capacity;
- Trust/vendor risk;
- Poor fit;
- AgriX withdrew;
- Unresponsive after confirmed process;
- Other — explanation required.

Dodatno beležiti:

- outcome subtype;
- stvarni root cause;
- konkurenta/status quo;
- odlučujući kriterijum;
- poslednji dokaz ili prigovor;
- stage u kom je ishod postao verovatan;
- da li je reactivation legitimna;
- future trigger i datum, ako postoji;
- šta treba promeniti u proizvodu, poruci ili procesu.

„Cena“ se ne bira kao razlog ako je stvarni uzrok nedokazana vrednost, pogrešan scope ili odsustvo prioriteta.

S9 se ne koristi za legitimno odlaganje sa potvrđenim budućim triggerom; takva prilika ide u SN.

---

## 18. Nurture i Deferred

Nurture zapis mora imati:

- razlog;
- konkretan future trigger;
- datum/prozor;
- relevantan sadržaj ili kontakt plan;
- owner-a.

Prihvatljivi triggeri:

- predsezonsko planiranje;
- promena ERP-a;
- otvaranje nove stanice;
- rast obima;
- odlazak ključnog administratora;
- regulatorna promena;
- završetak aktuelne sezone;
- budžetski ciklus;
- istek ugovora sa postojećim dobavljačem.

Bez triggera i datuma nalog je Dormant, ne aktivni Nurture.

Deferred je `SN Nurture` sa potvrđenim razlogom, budućim događajem i datumom — nije poseban opportunity stage.

---

## 19. Expansion, renewal i customer opportunity

Novi modul, dodatna firma, stanica ili veći scope vode se kao posebna expansion opportunity.

Ne mešati:

- podršku postojećem ugovoru;
- implementacioni change request;
- obnovu pretplate;
- novu komercijalnu ekspanziju.

Expansion zahteva novi problem/cilj, scope, vrednost, decision process, sales motion i next step.

Renewal se vodi prema potvrđenom ugovornom i customer-success procesu, a kada uključuje materijalnu promenu scope-a ili uslova dobija posebnu opportunity.

---

## 20. Implementation handoff

S8 Closed Won se ne smatra operativno završenim dok handoff nije prihvaćen.

Obavezni sadržaj:

- ugovoreni pravni subjekti;
- lokacije i stanice;
- paketi, moduli i instance;
- korisničke uloge;
- trenutni proces;
- potvrđeni problemi;
- success criteria;
- obim migracije;
- integracije;
- poznati gap-ovi;
- eksplicitno out-of-scope;
- custom razvoj i odobreni rokovi;
- hardver;
- obuke;
- implementacioni rok i sezonski deadline;
- kupčev project owner;
- AgriX implementation owner;
- rizici i zavisnosti;
- komercijalne obaveze relevantne za isporuku;
- obećanja data tokom prodaje;
- sales motion i razlog njegovog izbora;
- plan prvog kickoff-a.

Prodaja ne sme predati implicitna ili usmena obećanja kao ugovoreni scope.

Handoff sastanak završava se potvrdom:

- šta je prihvaćeno;
- šta zahteva razjašnjenje;
- ko je vlasnik svake otvorene stavke;
- datum kickoff-a.

Implementation owner ima pravo da odbije handoff ako je Fast Track korišćen za prikrivanje složenosti ili ako postoji neodobreno obećanje.

---

## 21. CRM activity note format

Svaka važna beleška koristi strukturu:

1. **Context** — ko je učestvovao i zašto;
2. **Confirmed** — koje činjenice su potvrđene;
3. **Changed** — šta se promenilo u prilici;
4. **Risks** — novi ili uklonjeni rizici;
5. **Decision** — šta je dogovoreno;
6. **Next step** — aktivnost, vlasnik i datum.

Ne unositi transkript bez zaključka.

Kod promene stage-a ili sales motion-a beleška mora navesti dokaz i razlog.

---

## 22. Minimalni dashboard pogledi

CRM mora omogućiti najmanje:

- pipeline po stage-u i vrednosti;
- pipeline po sales motion-u;
- pipeline po sezoni/kulturi;
- opportunities bez next step-a;
- overdue next steps;
- aging i stale opportunities;
- forecast po kategoriji;
- close-date slippage;
- conversion po stage-u;
- conversion i pre-sales sati po sales motion-u;
- source conversion;
- win/loss/no-decision razloge;
- aktivnosti i meaningful activity;
- opportunities bez economic buyer-a/champion-a;
- Fast Track escalation rate;
- implementation handoff status;
- expansion pipeline.

Detaljna KPI definicija pripada dokumentu `14_KPI_DASHBOARD_PLAYBOOK.md`.

---

## 23. Data governance

- Kontakt i poslovni podaci čuvaju se samo u legitimnom poslovnom kontekstu.
- Ne upisuju se nepotrebni lični podaci, privatne procene ličnosti ili uvredljive kvalifikacije.
- Beleže se ponašanja relevantna za proces odluke, ne psihološke etikete.
- Poverljive ponude, ugovori i dokumenti čuvaju se u odgovarajućem kontrolisanom prostoru, a CRM sadrži referencu i sažetak.
- Brisanje, pristup i izvoz moraju pratiti pravne i interne obaveze.
- Sales motion ne sme biti korišćen kao procena kvaliteta osobe ili firme, već isključivo složenosti komercijalnog i implementacionog procesa.

---

## 24. Zabranjeni obrasci

- kreiranje lažnih opportunity-ja radi većeg pipeline-a;
- držanje izgubljenih poslova otvorenim;
- pomeranje close date-a bez razloga;
- Commit bez kupčevog commitment-a;
- menjanje stage-a da bi forecast izgledao bolje;
- activity logging bez ishoda;
- dupliranje firmi i kontakata;
- čuvanje ključnih informacija samo u privatnim porukama;
- brisanje negativnih signala;
- upisivanje neodobrenih obećanja kao činjenica;
- označavanje kontakta kao blocker samo zato što postavlja legitimna pitanja;
- korišćenje Fast Track-a da bi se preskočio discovery, risk alignment ili proposal review;
- otvaranje novog opportunity-ja radi skrivanja povratka u raniju fazu;
- zadržavanje starih konfliktnih S8/S9 značenja u dashboardu, automatizaciji ili izveštaju.

---

## 25. CRM quality score

Za otvorenu opportunity, 0–16:

- account i kontakti kompletni: 1;
- buying roles evidentirane: 1;
- problem potvrđen: 1;
- consequence potvrđena: 1;
- timing potvrđen: 1;
- process map/dokumentovan kontekst: 1;
- champion status testiran: 1;
- economic buyer poznat: 1;
- decision process poznat: 1;
- scope dokumentovan: 1;
- stage exit kriterijumi ispunjeni: 1;
- realan close date: 1;
- top risk evidentiran: 1;
- mutual next step: 1;
- meaningful activity ažurna: 1;
- forecast kategorija potkrepljena: 1.

Tumačenje:

- 14–16: visoka CRM pouzdanost;
- 10–13: upotrebljivo uz otvorene rizike;
- 6–9: slab dokazni osnov;
- 0–5: opportunity verovatno nije pravilno kvalifikovana.

Dodatni guardrail bez boda:

- `sales_motion` je opravdan i ažuran;
- nema konflikta između kanonskog stage-a i commercial substatus-a.

---

## 26. Operativna rutina

### Posle svakog kontakta — isti dan

- uneti ishod;
- ažurirati polja koja su se promenila;
- zabeležiti rizik;
- postaviti next step;
- promeniti stage samo ako su kriterijumi ispunjeni;
- proveriti da li sales motion i dalje odgovara složenosti.

### Nedeljno

- pipeline hygiene;
- overdue i aging;
- forecast review;
- next-step review;
- Fast Track eligibility review;
- closure neaktivnih prilika.

### Mesečno

- conversion i velocity;
- forecast accuracy;
- win/loss/no-decision;
- source kvalitet;
- performanse po sales motion-u;
- CRM quality sampling;
- korekcija polja, definicija i playbook-a.

---

## 27. Validacioni plan

Prvih 30 aktivnih opportunity-ja koristiće se za proveru:

- da li kanonske stage definicije odgovaraju stvarnom procesu;
- gde opportunity najčešće stagnira;
- koja obavezna polja ne daju odluku ili su suvišna;
- koliko je next-step disciplina realno održiva;
- koji aging pragovi odgovaraju sezonskoj prodaji;
- koliko forecast kategorije predviđaju ishod;
- koji loss reasons se preklapaju;
- da li implementation handoff sprečava gubitak konteksta;
- koliko je Fast Track prilika pogrešno klasifikovano;
- koliki su pre-sales sati i conversion po motion-u;
- da li commercial substatus daje koristan detalj bez stvaranja paralelnog stage modela.

Na 10 zatvorenih prilika radi se prva revizija. Na 30 zatvorenih prilika usvajaju se v2 stage verovatnoće, aging pragovi i forecast standard.

Posebno se proverava da nijedan dashboard, formula ili automatizacija više ne koristi prethodno značenje `S8 = Commit/Contracting` ili `S9 = Closed Won`.

---

## 28. Veze sa drugim dokumentima

- `03_BUYING_PROCESS.md` — buying committee i proces odluke;
- `04_SALES_PROCESS.md` — kanonske faze i komercijalna disciplina;
- `04A_FAST_TRACK_SALES_MOTION.md` — eligibility, kompresovani tok i escalation Fast Track-a;
- `05_DISCOVERY_PLAYBOOK.md` — problem, posledice i success criteria;
- `08_DEMO_PLAYBOOK.md` — solution validation;
- `09_OBJECTION_HANDLING.md` — objection record;
- `10_NEGOTIATION_PLAYBOOK.md` — concession i approval evidencija;
- `11_CASE_STUDIES_PLAYBOOK.md` — reference i merljivi rezultati;
- `12_ROI_CALCULATOR_PLAYBOOK.md` — business case i pretpostavke;
- `14_KPI_DASHBOARD_PLAYBOOK.md` — metrike i dashboard;
- `15_ANNUAL_SALES_CALENDAR.md` — sezonski cadence i workload;
- `16_WEBSITE_SALES_ALIGNMENT_REVIEW.md` — javne poruke, claim i pricing usklađenost.

---

## 29. Definition of Done

CRM Pipeline segment prelazi iz DRAFT u DONE kada:

- CRM, Sales Process, dashboardi i automatizacije koriste isti S0–S9/SN model;
- stara konfliktna značenja S8/S9 više ne postoje u aktivnim izvorima;
- stage model bude testiran na najmanje 30 zatvorenih prilika;
- sva obavezna polja budu praktično proverena;
- aging pragovi budu prilagođeni stvarnom ciklusu;
- forecast kategorije imaju izmerenu tačnost;
- loss/no-decision taxonomy pokriva najmanje 90% ishoda bez `Other`;
- Fast Track, Standard i Complex imaju merljive kriterijume i porediv rezultat;
- handoff bude korišćen na najmanje pet implementacija;
- owner može iz CRM-a da rekonstruiše svaku aktivnu priliku bez privatnih beleški prodavca;
- postoji dokumentovana revizija v2.
