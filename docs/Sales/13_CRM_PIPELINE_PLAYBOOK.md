# AgriX CRM Pipeline Playbook

**Status:** DRAFT v1 — VALIDATION  
**Datum:** 30.07.2026.  
**Vlasnik:** osnivač AgriX-a  
**Svrha:** Uvesti jedinstven, proverljiv i operativan standard za evidentiranje naloga, kontakata, prilika, aktivnosti, forecast-a, no-deal ishoda i predaje dobijenih poslova implementaciji.

---

## 1. Osnovni princip

CRM nije arhiva kontakata niti zbir beleški. CRM je operativni sistem komercijalnog procesa i mora u svakom trenutku da pokaže:

- ko je kupac i koje firme/objekti pripadaju istom poslovnom sistemu;
- ko učestvuje u odluci i kakvu ulogu ima;
- koji potvrđeni problem se rešava;
- u kojoj fazi se prilika zaista nalazi;
- šta je sledeći korak, ko ga vodi i kada se dešava;
- koji rizik može da zaustavi posao;
- koliko je forecast zasnovan na dokazima;
- zašto je posao dobijen, izgubljen ili odložen;
- šta implementacija mora da zna pre početka rada.

**DECISION:** Faza se ne određuje po tome koliko je razgovora održano, već po ispunjenim exit kriterijumima.

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

- stvarna firma i kontakt;
- poslovni problem ili cilj;
- razlog za promenu ili rok;
- realan sledeći korak.

Obavezna polja:

- naziv opportunity-ja;
- account;
- owner;
- stage;
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

Kreira se tek za Closed Won i sadrži komercijalno-tehnički kontekst potreban implementaciji.

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

## 4. Opportunity stages

### S0 — Identified

**Entry:** potvrđen nalog i potencijalni use case.  
**Exit:** postoji relevantan kontakt i plan prvog razgovora.

Obavezno:

- source;
- ICP tier;
- hypothesized use case;
- next step.

### S1 — Connected

**Entry:** ostvaren dvosmerni kontakt.  
**Exit:** potvrđeno da postoji smislen razlog za discovery ili je prilika disqualified/nurture.

Obavezno:

- contact role;
- osnovni trigger;
- trenutni pristup;
- sledeći sastanak ili jasan ishod.

### S2 — Qualified

**Entry:** potvrđeni problem/cilj, kontekst, približan rok i relevantan sagovornik.  
**Exit:** zakazan discovery sa pravim učesnicima.

Minimalni PACT:

- Problem;
- Authority;
- Consequence;
- Timing.

### S3 — Discovery Complete

**Entry:** discovery održan.  
**Exit:** dokumentovani proces, posledice, success criteria, rizici, buying committee i dogovoren demo/tehnički korak.

Opportunity se ne pomera dalje ako postoji samo lista funkcija koje kupac želi.

### S4 — Solution Validated

**Entry:** demo, workflow review ili tehnička validacija povezani su sa potvrđenim problemima.  
**Exit:** potvrđen fit, evidentirani gap-ovi, definisan preliminarni scope i dogovoren komercijalni korak.

### S5 — Scope Confirmed

**Entry:** strane razumeju šta ulazi i ne ulazi u rešenje.  
**Exit:** potvrđena struktura ponude, implementacione pretpostavke i učesnici u odluci.

### S6 — Proposal Presented

**Entry:** ponuda je predstavljena uživo ili na strukturisanom pozivu; samo slanje dokumenta nije dovoljno.  
**Exit:** evidentirane reakcije, otvorena pitanja, decision process i sledeći datum.

### S7 — Commercial Review

**Entry:** kupac aktivno ocenjuje scope, cenu, uslove ili ugovor.  
**Exit:** svi materijalni prigovori i uslovi su rešeni ili je posao vraćen u raniju fazu/no-deal.

### S8 — Commit / Contracting

**Entry:** ekonomski kupac je potvrdio nameru pod jasno navedenim uslovima.  
**Exit:** potpis/obavezujuća potvrda i definisan početak implementacije.

### S9 — Closed Won

Obavezno:

- konačni scope;
- ugovorena vrednost;
- datum odluke;
- razlog dobitka;
- ključni dokaz;
- konkurencija/status quo;
- implementation owner;
- handoff datum.

### SN — Closed Lost / No Decision

Razlikovati:

- Closed Lost — izabran konkurent ili drugo rešenje;
- No Decision — status quo ostao;
- Disqualified — fit nije postojao;
- Withdrawn — AgriX je odustao;
- Deferred — potvrđeno odlaganje sa realnim budućim triggerom.

---

## 5. Stage discipline

Opportunity ne sme napredovati zbog:

- lepog razgovora;
- verbalnog interesa bez sledećeg koraka;
- održanog demo-a bez potvrđenog fit-a;
- poslate ponude bez review sastanka;
- neodređenog „javićemo se“;
- procene prodavca da će posao verovatno biti dobijen.

Kada se otkrije da prethodni kriterijum nije ispunjen, opportunity se vraća u odgovarajuću fazu.

---

## 6. Next-step standard

Svaka otvorena opportunity mora imati:

- konkretan sledeći korak;
- datum;
- vlasnika;
- drugu stranu koja je prihvatila korak;
- očekivani rezultat.

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

## 7. Meaningful activity

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

## 8. Aging i stagnacija

Početni pragovi za validaciju:

| Stage | Warning | Critical |
|---|---:|---:|
| S0–S1 | 10 dana | 20 dana |
| S2 | 14 dana | 30 dana |
| S3–S4 | 21 dan | 45 dana |
| S5–S6 | 21 dan | 45 dana |
| S7 | 30 dana | 60 dana |
| S8 | 21 dan | 45 dana |

Aging se tumači u odnosu na sezonski kontekst. Duga prilika sa potvrđenim future triggerom prelazi u Deferred/Nurture, ne ostaje veštački otvorena.

Stalled opportunity zahteva jedno od:

- novi mutual next step;
- povratak u raniju fazu;
- Deferred/Nurture;
- Closed Lost/No Decision;
- disqualification.

---

## 9. Forecast kategorije

### Pipeline

Postoji kvalifikovana prilika, ali nema dovoljno dokaza za period odluke.

### Best Case

Postoji validiran fit, uključen decision process i realna mogućnost odluke u periodu, ali ostaje važan rizik.

### Commit

Dozvoljeno samo kada postoje:

- potvrđen economic buyer;
- potvrđen scope i cena;
- rešeni materijalni prigovori;
- poznat proces odobrenja;
- konkretan datum odluke;
- kupčev eksplicitni commitment;
- nema nepoznatog kritičnog veto faktora.

### Closed

Won, Lost, No Decision, Withdrawn ili Disqualified.

**PROHIBITED:** Forecast se ne zasniva na osećaju, simpatiji kupca ili količini komunikacije.

---

## 10. Opportunity confidence score

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

---

## 11. Champion test

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

## 12. Close date pravila

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
5. zatvoriti ili deferovati kada nema realnog osnova.

---

## 13. Pipeline hygiene

Nedeljna revizija proverava:

- opportunity bez next step-a;
- next step u prošlosti;
- close date u prošlosti;
- critical aging;
- stage bez obaveznih polja;
- duplirane account-e/kontakte;
- opportunity bez economic buyer-a posle S4;
- proposal stage bez proposal review-a;
- Commit bez dokaza;
- neaktivne prilike koje treba zatvoriti;
- vrednost koja nije usklađena sa scope-om i cenovnikom;
- izgubljene poslove bez razloga i beleške.

Mesečna revizija dodatno proverava:

- stage conversion;
- velocity;
- source quality;
- forecast accuracy;
- najčešće loss/no-decision razloge;
- promene close date-a;
- stale nurture naloge;
- kvalitet CRM unosa po owner-u.

---

## 14. No-deal i closure standard

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

- stvarni root cause;
- konkurenta/status quo;
- odlučujući kriterijum;
- poslednji dokaz ili prigovor;
- da li je reactivation legitimna;
- future trigger i datum, ako postoji;
- šta treba promeniti u proizvodu, poruci ili procesu.

„Cena“ se ne bira kao razlog ako je stvarni uzrok nedokazana vrednost, pogrešan scope ili odsustvo prioriteta.

---

## 15. Nurture i Deferred

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

---

## 16. Expansion, renewal i customer opportunity

Novi modul, dodatna firma, stanica ili veći scope vode se kao posebna expansion opportunity.

Ne mešati:

- podršku postojećem ugovoru;
- implementacioni change request;
- obnovu pretplate;
- novu komercijalnu ekspanziju.

Expansion zahteva novi problem/cilj, scope, vrednost, decision process i next step.

---

## 17. Implementation handoff

Closed Won se ne smatra operativno završenim dok handoff nije prihvaćen.

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
- plan prvog kickoff-a.

Prodaja ne sme predati implicitna ili usmena obećanja kao ugovoreni scope.

Handoff sastanak završava se potvrdom:

- šta je prihvaćeno;
- šta zahteva razjašnjenje;
- ko je vlasnik svake otvorene stavke;
- datum kickoff-a.

---

## 18. CRM activity note format

Svaka važna beleška koristi strukturu:

1. **Context** — ko je učestvovao i zašto;
2. **Confirmed** — koje činjenice su potvrđene;
3. **Changed** — šta se promenilo u prilici;
4. **Risks** — novi ili uklonjeni rizici;
5. **Decision** — šta je dogovoreno;
6. **Next step** — aktivnost, vlasnik i datum.

Ne unositi transkript bez zaključka.

---

## 19. Minimalni dashboard pogledi

CRM mora omogućiti najmanje:

- pipeline po stage-u i vrednosti;
- pipeline po sezoni/kulturi;
- opportunities bez next step-a;
- overdue next steps;
- aging i stale opportunities;
- forecast po kategoriji;
- close-date slippage;
- conversion po stage-u;
- source conversion;
- win/loss/no-decision razloge;
- aktivnosti i meaningful activity;
- opportunities bez economic buyer-a/champion-a;
- implementation handoff status;
- expansion pipeline.

Detaljna KPI definicija pripada dokumentu `14_KPI_DASHBOARD_PLAYBOOK.md`.

---

## 20. Data governance

- Kontakt i poslovni podaci čuvaju se samo u legitimnom poslovnom kontekstu.
- Ne upisuju se nepotrebni lični podaci, privatne procene ličnosti ili uvredljive kvalifikacije.
- Beleže se ponašanja relevantna za proces odluke, ne psihološke etikete.
- Poverljive ponude, ugovori i dokumenti čuvaju se u odgovarajućem kontrolisanom prostoru, a CRM sadrži referencu i sažetak.
- Brisanje, pristup i izvoz moraju pratiti pravne i interne obaveze.

---

## 21. Zabranjeni obrasci

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
- označavanje kontakta kao blocker samo zato što postavlja legitimna pitanja.

---

## 22. CRM quality score

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

---

## 23. Operativna rutina

### Posle svakog kontakta — isti dan

- uneti ishod;
- ažurirati polja koja su se promenila;
- zabeležiti rizik;
- postaviti next step;
- promeniti stage samo ako su kriterijumi ispunjeni.

### Nedeljno

- pipeline hygiene;
- overdue i aging;
- forecast review;
- next-step review;
- closure neaktivnih prilika.

### Mesečno

- conversion i velocity;
- forecast accuracy;
- win/loss/no-decision;
- source kvalitet;
- CRM quality sampling;
- korekcija polja, definicija i playbook-a.

---

## 24. Validacioni plan

Prvih 30 aktivnih opportunity-ja koristiće se za proveru:

- da li stage definicije odgovaraju stvarnom procesu;
- gde opportunity najčešće stagnira;
- koja obavezna polja ne daju odluku ili su suvišna;
- koliko je next-step disciplina realno održiva;
- koji aging pragovi odgovaraju sezonskoj prodaji;
- koliko forecast kategorije predviđaju ishod;
- koji loss reasons se preklapaju;
- da li implementation handoff sprečava gubitak konteksta.

Na 10 zatvorenih prilika radi se prva revizija. Na 30 zatvorenih prilika usvajaju se v2 stage verovatnoće, aging pragovi i forecast standard.

---

## 25. Veze sa drugim dokumentima

- `03_BUYING_PROCESS.md` — buying committee i proces odluke;
- `04_SALES_PROCESS.md` — faze i komercijalna disciplina;
- `05_DISCOVERY_PLAYBOOK.md` — problem, posledice i success criteria;
- `08_DEMO_PLAYBOOK.md` — solution validation;
- `09_OBJECTION_HANDLING.md` — objection record;
- `10_NEGOTIATION_PLAYBOOK.md` — concession i approval evidencija;
- `11_CASE_STUDIES_PLAYBOOK.md` — reference i merljivi rezultati;
- `12_ROI_CALCULATOR_PLAYBOOK.md` — business case i pretpostavke;
- budući `14_KPI_DASHBOARD_PLAYBOOK.md` — metrike i dashboard;
- budući `15_ANNUAL_SALES_CALENDAR.md` — sezonski cadence i workload.

---

## 26. Definition of Done

CRM Pipeline segment prelazi iz DRAFT u DONE kada:

- stage model bude testiran na najmanje 30 zatvorenih prilika;
- sva obavezna polja budu praktično proverena;
- aging pragovi budu prilagođeni stvarnom ciklusu;
- forecast kategorije imaju izmerenu tačnost;
- loss/no-decision taxonomy pokriva najmanje 90% ishoda bez `Other`;
- handoff bude korišćen na najmanje pet implementacija;
- owner može iz CRM-a da rekonstruiše svaku aktivnu priliku bez privatnih beleški prodavca;
- postoji dokumentovana revizija v2.
