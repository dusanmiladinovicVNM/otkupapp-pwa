# AgriX Master Plan — Change Log

## 2026-07-27 (5) — cenovnik verzija 3, odluke 416 i 423

### Added

- **odluka 423** u `09_QA_DECISION_LOG.md` §26.1 — Savetnik se u cenovniku prikazuje samo uz vidljivu oznaku „u pripremi“, bez vizuelne oznake preporuke. Analogno odluci 417 za GGAP;
- strana 9 cenovnika dobila blok **„Prestanak i obnova“** — read-only 30 dana i godišnja pretplata bez obzira na dužinu sezone. Oba podatka provereni: 30 dana u odluci 117 i članu 7(4) ugovora, sezonska pretplata u odluci 28.

### Changed

- **odluka 416 preformulisana:** bazne cene paketa objavljuju se kao **tačan iznos uz kvalifikator obima**, a „od X €“ samo tamo gde je iznos stvarno donja granica — trenutno samo GGAP. Time 416 **zamenjuje i odluku 297**, koja je ranije bila potvrđena; 297 je označena `Superseded`, a u `09B` je dobila `→ 416`;
- **strana 4:** Dispatch je vraćena obračunska jedinica („po pravnom licu“) uz zasebnu neutralnu oznaku dostupnosti („samo uz Mobile“); Hladnjača ima punu formulaciju „po proizvodnom pogonu“ umesto skraćene;
- **strana 6:** Savetnik nosi oznaku „u pripremi“, zlatni okvir uklonjen, dodata napomena da se ne ugovara; blok „Savetnik u praksi“ dopunjen izuzetkom da se Pro ne plaća dvaput kada su gazdinstva pokrivena partnerskim paketom;
- **strana 7:** naslovna rečenica više ne protivreči redovima „uključeno“;
- **strana 8:** treći primer prikazuje 2.450 € kao ukupan iznos, a GGAP stoji ispod tamnog bloka sa oznakom „na upit“ i „od 1.000 €“ — veliki broj sada sadrži samo fiksno kotirano;
- **tipografija:** linijske cifre na svim cenama, all-in cena izjednačena sa baznom po veličini, kvačice na strani 9 crtaju se CSS-om umesto znakom U+2713 — `DejaVuSans-Bold` više nije u PDF-u;
- `07_PRODUCT_PORTFOLIO.md` i `docs/Sales/README.md` dopunjeni statusom „u pripremi“;
- u `AgriX_Sablon_ponude.xlsx` Savetnik redovi označeni `[u pripremi]` sa napomenom da se ne uvrštavaju u ponudu.

### Napomena

- Cene nisu dirane. Iznos od 2.450 € u trećem primeru je **prikaz**, ne promena cene — Prilog 1 ugovora i finansijski model ostaju nepromenjeni.

## 2026-07-27 (4) — cenovnik dobio izvor, PDF regenerisan

### Added

- `docs/Sales/AgriX_Cenovnik_2027.html` — izvor istine za cenovnik;
- `tools/cenovnik.sh` — `build` renderuje PDF preko headless Chromium-a, `check` programski poredi cene na sva četiri mesta.

### Changed

- **`AgriX_Cenovnik_2027.pdf` regenerisan** sa svih jedanaest izmena iz cenovne revizije: all-in 1.200 i 2.200 €, izričit sastav all-in paketa uz naglašeno da Desktop all-in ne sadrži Dispatch, Hladnjača po proizvodnom pogonu i dodatni pogon 200 €, dve satnice, vreme puta po implementacionoj satnici, Enterprise tarifa Savetnika, GGAP sa oznakom „na upit“, preračunati primeri (500 / 1.450 / 3.450 €) i datum 27.07.2026.

### Vizuelni identitet

- Cenovnik je usklađen sa **AgriX vizuelnim identitetom**: brand tokeni iz `src/styles/base.css` (forest `#1E2D14`, accent `#5EA135`, gold `#C8A84B`, cream skala, radius skala), tipografija Cormorant Garamond + DM Sans iz `vendor/fonts/`, logo `img/AgriX-Logo-Final_Novi.png` na naslovnoj i wordmark u podnožju;
- forest naslovna strana, kartice sa brand radius-om, zelene oznake za uključeno i zlatne za „na upit“;
- `check` više ne zavisi od izgleda — cene se čitaju iz `data-cena`/`data-eur` atributa i porede sa prikazanim tekstom.

## 2026-07-27 (3) — cenovna revizija, odluke 409–422

### Added

- **odluke 409–422** u `09_QA_DECISION_LOG.md` §26.1: dve satnice, obračunska jedinica modula, dodatna instanca, formiranje i prikaz cena, politika popusta, tarife Savetnika;
- odeljak „Satnice“ u `07_PRODUCT_PORTFOLIO.md` §10;
- Prilog 1 ugovora: sastav all-in paketa, obračunska jedinica modula, dve satnice, tarife Savetnika, klauzula o popustima;
- u `AgriX_Sablon_ponude.xlsx` list `Cenovnik`: dodatni pogon, Gazdinstvo Basic kanalska, četiri reda tarifa Savetnika;
- u `AgriX_Finansijski_model.xlsx`: sekcija G i dva prihodna reda na listu `Prihod`.

### Changed

- **cene all-in paketa:** Desktop all-in 1.100 → **1.200 €**, Mobile all-in 2.100 → **2.200 €**. Struktura je sada aditivna: Mobile dodatak +1.000 €, all-in doplata +700 €, oba fiksna na oba nivoa;
- odluka 349 prepravljena na mestu, bez novog broja;
- satnice preimenovane u **razvojnu (50 €/h)** i **implementacionu (30 €/h)**; biraju se po prirodi posla, ne po mestu izvođenja;
- **Hladnjača/Proizvodnja se plaća po proizvodnom pogonu**, ne po pravnom licu; dodatni pogon 200 €;
- GGAP se u cenovniku sme prikazati samo uz oznaku „na upit, uz potvrdu obima — nije deo standardne ponude“;
- Savetnik dobija drugu, Enterprise tarifu 100 € / 10 €;
- `09B_ODLUKE_PO_OBLASTIMA.md` regenerisan na **verziju 3** — obuhvata 1–321, 323–378 i 401–422, uklonjena oznaka „zastarelo“, dopunjena odluka 370 koje ranije nije bilo u indeksu.

### Superseded / Deleted

- `Deleted` odlukom 418: **C3**, **111**, **112** — pregovaračkih i individualnih popusta nema;
- `Superseded`: 110 → 409 · 126 i 127 → 414 · 133 → 413 · 156 i 157 → 416 · 198 → 420 · 207 → 419;
- `Closed`: 130 → 421 · `Rewritten`: IP4 → 422.

### Resolved

- cena po gazdinstvu kod Savetnika (341) — potvrđena, 150 € / 15 €;
- neusaglašenost indeksa 09B sa logom.

### Open

- ugovor je nacrt bez pravnog pregleda (376); Prilog 3 nije dovršen (LEG1);
- spisak podobrađivača i lokacija obrade (373);
- dve pretpostavke u finansijskom modelu namerno ostavljene na 0 (broj Savetnik licenci i dodatnih pogona po klijentu), pa ARR ostaje nepromenjen dok se ne popune;
- Desktop all-in štedi klijentu samo 100 € — slaba bundle poruka, komercijalno pitanje.

## 2026-07-27 (2) — odluke 323–378, cenovnik, ugovor i razrešenja

> Ovaj unos ispravlja i dopunjuje unos `2026-07-27 (1)` niže. Gde se razlikuju, važi ovaj.

### Added

- **odluke 323–378** u `09_QA_DECISION_LOG.md` §25, grupisane po oblastima (cene, ugovor, podrška, onboarding, prodaja, tržišni cilj, Gazdinstvo, Savetnik, podaci i bezbednost, razvoj i organizacija), uz tabelu otvorenih stavki §25.11;
- `09B_ODLUKE_PO_OBLASTIMA.md` — tematski indeks svih odluka, i `09B_ODLUKE_PO_OBLASTIMA_2026-07-26.pdf` kao renderovani snimak;
- `docs/Sales/AgriX_Cenovnik_2027.pdf` — zvanični cenovnik od sezone 2027;
- `docs/Legal/AgriX_Ugovor_o_licenciranju.docx` — nacrt ugovora o licenciranju sa Prilozima 1–3.

### Changed

- `09_QA_DECISION_LOG.md`: odluke 401–408 pomerene u §26, napomena o numeraciji u §27 i ispravljena — brojevi **322 i 379–400 se ne koriste**, opseg 323–378 je popunjen;
- **odluka 375 ispravljena**: nije „17–18 ukupno“ nego **12–15 novih / 15–18 ukupno**; ispravljeno u `02_STRATEGY.md` §9 i §10, `04_MARKET.md` §9.1 i §22 i `DECISION_LOG.md` STR-014;
- **Savetnik ima objavljenu cenu** (150 € do 10 gazdinstava + 15 € po gazdinstvu, odluke 341 i 347): `02_STRATEGY.md` §2.3, `07_PRODUCT_PORTFOLIO.md` §9, §11 i §13, `07A_PRODUCT_STATUS_MATRIX.csv`. Raniji zapis „packaging i cena nisu zaključani“ bio je netačan;
- README-ji u `docs/Sales/`, `docs/Legal/` i `docs/Finance/` dopunjeni novim dokumentima i pravilom da cenovnik, šablon ponude, Prilog 1 ugovora i finansijski model moraju imati iste iznose.

### Resolved — odluke osnivača 27.07.2026.

- **cilj rasta:** merodavna je planska vrednost iz finansijskog modela — **14 novih / 17 ukupno** do sezone 2027. Odluka 375 preformulisana sa raspona na jednu vrednost; usklađeni `02_STRATEGY.md` §9 i §10, `04_MARKET.md` §9.1 i §22, `DECISION_LOG.md` STR-014 i `docs/Finance/README.md`. Kolona od 18 na listu `Kapacitet` označena je kao stress-test, ne cilj;
- **prozor za rok od jednog sata (359):** tokom sezone svakog dana 08.00–20.00 uključujući vikend, van sezone radnim danima 08.00–16.00. Upisano u Prilog 2 ugovora i u odluku 359;
- **hardverska podrška (357):** potvrđena kako stoji — 40 € po stanici godišnje, minimum 200 € po pravnom licu. Oznaka `PROPOSAL` uklonjena;
- **sati obuke (362):** pet sati ukupno, i za onboarding i za uvođenje modula; odluka 362 usklađena sa 354 i 365, bez odvojene kvote po modulu.

### Tooling

- `.claude/skills/poslovni-dokumenti/SKILL.md` — pravila za rad sa `docs/`: gde je istina, numeracija odluka, sinhronizacija cena na četiri mesta, klase dokumenata, gde ide koji upload, provere pre commit-a i spisak naučenih grešaka;
- `docs/Legal/AgriX_Ugovor_o_licenciranju.md` — tekst ugovora kao izvor istine, sa mapom svih članova na ID-jeve odluka;
- `tools/ugovor.sh` — `build` generiše `.docx` iz `.md`, `check` prijavljuje razlike između njih. Zavisnost: `pandoc`.
- **mesto nadležnog suda:** Niš (član 15), jer je sedište AgriX-a Merošina.

### Open

- ugovor je nacrt bez pravnog pregleda (376); Prilog 3 nije dovršen i ne potpisuje se u ovom obliku (LEG1);
- nepopunjeno u ugovoru: spisak podobrađivača i lokacija obrade (373);
- cena po gazdinstvu kod Savetnika 15 € (341) čeka potvrdu;
- `09B_ODLUKE_PO_OBLASTIMA.md` još ne obuhvata odluke 401–408.

## 2026-07-27 (1) — odluke 401–408 i poslovni dokumenti

### Added

- odluke **401–408** u `09_QA_DECISION_LOG.md` (odeljak 25; kasnije premešten u odeljak 26);
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
