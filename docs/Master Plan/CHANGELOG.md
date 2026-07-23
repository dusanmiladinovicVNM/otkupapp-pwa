# AgriX Master Plan — Change Log

## 2026-07-23

### Added

- `04_MARKET.md` sa APR zasnovanim tržišnim universe-om, segmentacijom, koncentracijom prihoda i ograničenjima šifara delatnosti;
- `05_COMPETITION.md` i konkurentski evidence skup za SOFTEK, KRUNET, Yuteam i Infosys;
- `05A_COMPETITOR_EVIDENCE.md` i `05B_INFOSYS_REPLACEMENT_GTM.md` sa Infosys replacement analizom;
- reproduktivni Infosys APR matching, wide enrichment i sales-readiness pipeline;
- account-research skup, win-interview obrazac i migration-discovery checklist;
- `06_POSITIONING.md` sa odlukom o tržišnoj kategoriji i dozvoljenim/zabranjenim tvrdnjama;
- `07_PRODUCT_PORTFOLIO.md` sa proizvodnim stubovima, komercijalnim statusima, modulima, uslugama, hardverom i readiness gate-ovima;
- `07A_PRODUCT_STATUS_MATRIX.csv` kao strukturisani izvor za roadmap i pricing;
- `08_PRODUCT_ROADMAP.md` kao gate-based roadmap od core safety-ja do Field Operations, Gazdinstvo i GGAP validacije;
- `08A_ROADMAP_MILESTONES.csv` kao operativna matrica faza, ciljnih prozora, zavisnosti i exit gate-ova.

### Approved / Proposed decisions

- AgriX se pozicionira kao vertikalni poslovni operativni sistem za organizovani otkup poljoprivrednih proizvoda;
- Enterprise je primarno komercijalno jezgro;
- Management PWA je deo Enterprise proizvoda, ne zaseban BI proizvod;
- PWA Otkupac, kiosk i termalna štampa ostaju `Pilot only` do sezonske validacije;
- Gazdinstvo Partner/Basic/Pro ostaje kontrolisana rana ponuda dok se ne potvrde activation, retention, willingness-to-pay i support cost;
- GGAP ostaje discovery/pilot proizvod i ne prodaje se kao završena produkciona ili sertifikaciona garancija;
- postojanje funkcije u kodu nije dovoljno za status `Standard offer`;
- hardver, migracija, onboarding i posebne integracije imaju odvojenu ekonomiku;
- trajni klijentski forkovi ostaju zabranjeni;
- otvoreni P0 data-safety, statusni i authorization rizici imaju prednost nad novim funkcijama;
- Field Operations prelazi u standardnu prodaju tek posle kontrolisanog realnog sezonskog pilota;
- uvodi se sezonski feature freeze najmanje 30 dana pre kritične sezone pilot-klijenta;
- Gazdinstvo se prioritetno validira kroz activation/retention/WTP, a ne kroz širenje premium scope-a;
- pun GGAP razvoj ne počinje bez stručnog domain owner-a, standarda/verzije, pilot-klijenta, data mapiranja i ekonomske hipoteze;
- `HOLD`, `REDUCE SCOPE` i `STOP` su legitimne roadmap odluke kada inicijativa ne prolazi dokazni ili ekonomski gate.

### Evidence and qualification

- Infosys je potvrđen kao prioritetni replacement konkurent kroz dve postojeće migracije ka AgriX-u;
- 114 agro/prehrambenih Infosys referenci čini početni universe, sa 49 visokopotencijalnih redova;
- wide APR enrichment je identifikovao 30 jedinstvenih pravnih lica, ali je identity match odvojen od stvarnog AgriX process fit-a;
- masovni outbound ka celoj referentnoj bazi je odbijen; prioritet je mali, spoljno validiran account-research talas;
- prihod je pomoćni signal, dok su broj stanica, kooperanata, dokumenata, logistika i procesna složenost važniji ICP kriterijumi;
- poslovni roadmap je povezan sa aktivnim tehničkim auditom i ne može proglasiti proizvod spremnijim od runtime/release evidence-a.

### Next

- završiti Fazu 0: rebase i verifikacija aktivnih P0 nalaza, target-workbook health i runtime release evidence;
- zaključati standardni Enterprise Core onboarding i migracioni scope;
- izabrati jednog Field Operations pilot-klijenta, 1–3 početne stanice i podržani tablet/printer paket;
- razviti `10_PRICING_AND_PACKAGING.md` na osnovu portfolija, roadmap gate-ova i stvarnog support/onboarding troška;
- sprovesti dva Infosys win interview-a kada termini budu dostupni;
- rezultate intervjua pretvoriti u battlecard, migration package i dokazne prodajne poruke.

## 2026-07-22

### Added

- početni sadržaj i mapa svih planiranih poglavlja u `README.md`;
- governance pravila i klasifikacija tvrdnji u `00_GOVERNANCE.md`;
- formalni `DECISION_LOG.md`;
- prva puna verzija `02_STRATEGY.md`;
- `02A_GGAP_STRATEGY.md` kao posebna strategija trećeg proizvodnog stuba;
- `03_CUSTOMERS_AND_JOBS.md` sa ulogama, jobs-to-be-done, buying committee modelom, segmentacijom i ICP scoringom.

### Approved

- Master Plan se vodi na srpskom;
- osnovna planska valuta je EUR, uz RSD za lokalne tokove;
- velika poglavlja razvijaju se odvojeno;
- klijenti se anonimizuju;
- osetljivi podaci se izdvajaju iz javnog tehničkog repoa;
- sezonski cap određuje readiness score, ne unapred fiksiran broj firmi;
- tržišni fokus je Srbija, uz hladnjače i druge firme sa razgranatom mrežom stanica i kooperanata;
- klijentski forkovi nisu dozvoljeni;
- Gazdinstvo trenutno nije osnovni prihod, ali može postati glavni proizvod ako podaci to potvrde;
- hardver je sporedni profitni centar i potencijalni ulaz u širi IT portfolio;
- partner se ne uzima samo zbog kapitala;
- prva operativna osoba je customer support / implementation;
- dugoročni cilj je regionalna platforma;
- strateški cilj je najmanje 200 firmi u naredne 3–4 godine;
- AgriX je end-to-end poslovni sistem;
- Gazdinstvo je pun farm-management proizvod;
- GGAP je treći puni proizvodni stub.

### Changed

- strategija rasta promenjena je sa fiksnog limita od 8–10 firmi na readiness-based model;
- vizija je podignuta sa lokalnog profitabilnog specijaliste na regionalnu vertikalnu platformu;
- ciljna grupa je proširena sa hladnjača na sve organizovane otkupljivače sa mrežom stanica i kooperanata;
- hardver je redefinisan iz enablementa u profitabilni sporedni centar uz mogući širi IT sistem;
- tržišni cilj od 200 firmi uveden je kao ambicija, ne prognoza;
- proizvodna arhitektura je definisana kroz tri povezana stuba: Enterprise, Gazdinstvo i GGAP;
- kupac se više ne modeluje samo kao vlasnik, operater i kooperant, već kao višeuloga buying committee i operativni lanac.

### Review

- predložene odluke CUS-001 do CUS-005 čekaju potvrdu nakon pregleda `03_CUSTOMERS_AND_JOBS.md`;
- hipoteze o najboljem segmentu, activation funnel-u Gazdinstva i willingness-to-pay za GGAP zahtevaju intervjue i merenje.

### Next

- pregledati i zaključati `03_CUSTOMERS_AND_JOBS.md`;
- razviti `04_MARKET.md` i potvrditi procenu 500–1.000 relevantnih firmi;
- sprovesti GGAP discovery: standard, verzija, liste, uloge, dokazi i audit tok;
- definisati formalni readiness score;
- zatim razviti portfolio, pricing, unit economics i finansijski plan do 200 firmi.
