# AgriX Master Plan — Change Log

## 2026-07-27

### Added

- odluke **401–408** u `09_QA_DECISION_LOG.md` (odeljak 25);
- napomena o nedostajućim odlukama **322–400** (odeljak 26) — na njih se pozivaju dokumenti u `docs/Product/`, `docs/Sales/` i `docs/Finance/`, ali tekst nije u repou;
- `STR-013` (Savetnik kao treći stub) i `STR-014` (fiksan ciljni broj klijenata) u `DECISION_LOG.md`;
- odeljak `02_STRATEGY.md` §2.3 „AgriX Savetnik“ i pododeljak „Moduli uz Enterprise“;
- odeljak `07_PRODUCT_PORTFOLIO.md` §9 „AgriX Savetnik“ i §9A „GGAP — modul Enterprise-a“;
- red za Savetnik u `07A_PRODUCT_STATUS_MATRIX.csv`;
- poslovni dokumenti van Master Plana: `docs/Product/AgriX_Definicija_proizvoda.pdf`, `docs/Legal/AgriX_Mapa_tokova_podataka.pdf`, `docs/Sales/AgriX_Materijal_za_prvi_kontakt.pdf`, `docs/Sales/AgriX_Sablon_ponude.xlsx`, `docs/Finance/AgriX_Finansijski_model.xlsx`, uz indeks u README-ju svakog direktorijuma.

### Changed

- **Savetnik je treći stub** (odluka 401, potvrđuje 269): `02_STRATEGY.md` §2 i `07_PRODUCT_PORTFOLIO.md` §3;
- **GGAP je modul Enterprise-a, ne stub** (odluka 402): `02_STRATEGY.md` §2, `02A_GGAP_STRATEGY.md` §1, §7 i §11, `07_PRODUCT_PORTFOLIO.md` §3 i §9A, `07A_PRODUCT_STATUS_MATRIX.csv`;
- **readiness cap zamenjen fiksnim ciljem** (odluka 403): `02_STRATEGY.md` §9, §10 Faza 1, §15 i §17; readiness prelazi u kontrolnu listu pred onboarding;
- **cilj rasta usklađen sa odlukom 375**: `04_MARKET.md` §9.1 — ubrzani raspon 12–15 zamenjen izabranim scenarijem C, 17–18 aktivnih firmi do sezone 2027 (14–15 novih uz postojeće 3);
- **Gazdinstvo iz `Pilot only` u `Standard offer`** (odluka 404): `07_PRODUCT_PORTFOLIO.md` §8 i §11, `07A_PRODUCT_STATUS_MATRIX.csv`;
- **GGAP ostaje van komercijalne ponude do validacije** (odluka 405): `07_PRODUCT_PORTFOLIO.md` §9A i §11;
- **jedinstvena cena po stanici** (odluka 406): `07B_ENTERPRISE_OPERATING_MODES.md` odluka 9 zatvorena — razliku pokriva cena Mobile paketa; posledica upisana u `07_PRODUCT_PORTFOLIO.md` §13.

### Superseded

- `STR-001` — readiness-based rast → odluka 403 / STR-014;
- `STR-012` — GGAP kao treći proizvodni stub → odluke 401 i 402 / STR-013;
- `07B` odluka 9 — pricing koji razlikuje desktop-only od PWA-led cene po stanici → odluka 406.

### Open

- odluke 322–400 nisu unete; do tada tvrdnje izvedene iz njih nisu proverljive u repou;
- Savetnik nema product strategy, packaging ni cenu — nema komercijalni status;
- hardverska marža ostaje planska do izbora dobavljača (odluka 407), a cena hardverske podrške (357) je i dalje predlog;
- troškovi u finansijskom modelu nisu popunjeni, pa neto rezultat i cash-flow još nemaju smisla;
- LEG1 nije razrešen — bez njega nema Priloga 3 ugovora ni politika privatnosti za Gazdinstvo i Savetnik.

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
- `08_PRODUCT_ROADMAP.md` kao gate-based roadmap;
- `08A_ROADMAP_MILESTONES.csv` kao operativna matrica milestone-a.

### Critical correction — PWA-led operating model

- PWA Otkupac i PWA Vozač nisu sporedni Field Operations dodatak desktop proizvodu;
- AgriX Enterprise Core je end-to-end sistem `teren → sync → centralna baza → prijem/faktura/izveštaj`;
- otkupljivači i vozači sami stvaraju osnovne poslovne događaje i dokumente na mestu nastanka;
- centralni operater se primarno bavi kontrolom, prijemom, fakturama, finansijama i izveštajima, a ne ponovnim unosom terenskih podataka;
- PWA, GAS, Sheets/MasterSync i desktop backoffice predstavljaju jedan proizvodni tok;
- kiosk, tablet i termalna štampa dobijaju zasebne readiness statuse i ne smeju da obore status funkcionalne PWA aplikacije;
- roadmap je promenjen tako da PWA-led productization i core correctness imaju jednak strateški prioritet;
- glavni product KPI postaje procenat poslovnih događaja koji od terena do centrale prolaze bez ponovnog unosa.

### Approved / Proposed decisions

- AgriX se pozicionira kao terenski i centralni operativni sistem za organizovani otkup;
- Enterprise je primarno komercijalno jezgro;
- PWA Otkupac i PWA Vozač su centralne komponente Enterprise Core-a;
- Management PWA je deo Enterprise proizvoda, ne zaseban BI proizvod;
- centralni desktop je canonical backoffice posle sinhronizacije, ali nije zamišljen kao mesto rutinskog prepisivanja terenskih događaja;
- PWA status se određuje prema konkretnom aktivnom scope-u i release evidence-u: `Standard offer` ili `Controlled rollout`;
- kiosk standardizacija i termalna štampa ostaju odvojeni `Controlled rollout` tokovi;
- Gazdinstvo Partner/Basic/Pro ostaje kontrolisana rana ponuda dok se ne potvrde activation, retention, willingness-to-pay i support cost;
- GGAP ostaje discovery/pilot proizvod i ne prodaje se kao završena produkciona ili sertifikaciona garancija;
- postojanje funkcije u kodu nije dovoljno za status `Standard offer`, ali nepostojanje standardnog hardware paketa nije dokaz da sama PWA nije spremna;
- hardver, migracija, onboarding i posebne integracije imaju odvojenu ekonomiku;
- trajni klijentski forkovi ostaju zabranjeni;
- potvrđeni P0 data-safety, statusni i authorization rizici imaju prednost nad novim nepovezanim funkcijama;
- uvodi se sezonski feature freeze najmanje 30 dana pre kritične sezone;
- Gazdinstvo se validira kroz activation/retention/WTP, a ne kroz širenje premium scope-a;
- pun GGAP razvoj ne počinje bez stručnog domain owner-a, standarda/verzije, pilot-klijenta, data mapiranja i ekonomske hipoteze.

### Evidence and qualification

- Infosys je potvrđen kao prioritetni replacement konkurent kroz dve postojeće migracije ka AgriX-u;
- 114 agro/prehrambenih Infosys referenci čini početni universe, sa 49 visokopotencijalnih redova;
- wide APR enrichment je identifikovao 30 jedinstvenih pravnih lica, ali je identity match odvojen od stvarnog AgriX process fit-a;
- masovni outbound ka celoj referentnoj bazi je odbijen; prioritet je mali, spoljno validiran account-research talas;
- prihod je pomoćni signal, dok su broj stanica, terenskih korisnika, dokumenata, logistika i procesna složenost važniji ICP kriterijumi;
- poslovni roadmap je povezan sa aktivnim tehničkim auditom i ne može proglasiti proizvod spremnijim od runtime/release evidence-a.

### Next

- izmeriti postojeći PWA-led tok: procenat terenskih unosa, sync uspeh, ručne centralne korekcije i vreme operatera;
- završiti P0 closeout i sačuvati end-to-end release evidence;
- zaključati standardni field-to-office onboarding i migracioni scope;
- odvojeno standardizovati tablet/kiosk i termalni print paket;
- razviti `10_PRICING_AND_PACKAGING.md` tako da vrednuje broj stanica, terenskih korisnika i obim dokumenata, a ne samo desktop licencu;
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
