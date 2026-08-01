# Sales

Prodajni playbook, discovery, demo, kvalifikacija, ponude, prigovori, pipeline pravila i win/loss proces.

Poverljive ponude, kontakt podaci i ugovorni detalji ne čuvaju se u javnom repozitorijumu.

## Commercial Operating System

| Dokument | Status | Sadržaj |
|---|---|---|
| `00_COMMERCIAL_OPERATING_SYSTEM_ROADMAP.md` | DONE v1 · 28.07.2026. | Redosled izrade, standard dokaza, Definition of Done, Customer Intelligence Loop i status svih oblasti. |
| `01_MARKET_POSITIONING.md` | DRAFT v1 — VALIDATION · 28.07.2026. | Tržišna kategorija, centralni problem narrative, poziciona teza, value pillars, competitive frame, ICP/anti-positioning, message house, proof hierarchy, website implications i plan validacije. |
| `02_PSYCHOLOGICAL_PROFILES.md` | DRAFT v1 — VALIDATION · 28.07.2026. | Evidence-based buying committee: vlasnik, operativa, administrator, finansije, teren, IT i skriveni influencer; motivi, rizici, dokazni prag, buying signals i validacione hipoteze. |
| `03_BUYING_PROCESS.md` | DRAFT v1 — VALIDATION · 28.07.2026. | Buying triggers, committee i champion test, faze B0–B7, decision criteria, skriveni veto, sezonski timing, Mutual Action Plan, stage advancement, CRM polja i no-deal pravila. |
| `04_SALES_PROCESS.md` | DRAFT v1 — VALIDATION · 28.07.2026. | Faze S0–S9/SN, PACT kvalifikacija, entry/exit kriterijumi, SLA, stage aging, next-step disciplina, forecast, no-deal, pipeline hygiene i implementation handoff. |
| `05_DISCOVERY_PLAYBOOK.md` | DRAFT v1 — VALIDATION · 28.07.2026. | Pre-call intelligence, C-P-I-O-R-D tok razgovora, process mapping, pitanja po personama, consequence chain, success criteria, risk/decision discovery, scoring, recap i CRM zapis. |
| `06_EMAIL_SEQUENCES.md` | DRAFT v1 — VALIDATION · 29.07.2026. | Šestomesečni cadence i gotovi tekstovi za cold outbound, inbound, post-call, discovery, demo, proposal, nurture, reaktivaciju i referral; grananje, SLA, CRM i A/B test pravila. |
| `07_CALL_PLAYBOOKS.md` | DRAFT v1 — VALIDATION · 29.07.2026. | O-R-E-D-A struktura, cold i inbound pozivi, kvalifikacija, discovery, demo confirmation, post-demo, scope, proposal review, stalled deal, reaktivacija, reakcije, CRM i coaching standard. |
| `08_DEMO_PLAYBOOK.md` | DRAFT v1 — VALIDATION · 29.07.2026. | R-E-L-A-Y struktura, demo brief, executive/operational/technical tokovi, storyline po problemu i personi, gap management, fit review, exit kriterijumi, CRM i quality score. |
| `09_OBJECTION_HANDLING.md` | DRAFT v1 — VALIDATION · 29.07.2026. | A-C-T-I-O-N dijagnostika, cena/vrednost, status quo, ERP, implementacija, vendor rizik, gap-ovi, konkurencija, no-deal, CRM i quality score. |
| `10_NEGOTIATION_PLAYBOOK.md` | DRAFT v1 — VALIDATION · 29.07.2026. | P-A-C-T-S okvir, cenovna disciplina, give/get, scope trade-offs, plaćanje, rokovi, SLA, pilot, custom razvoj, approval matrix, concession log, walk-away i CRM. |
| `11_CASE_STUDIES_PLAYBOOK.md` | DRAFT v1 — VALIDATION · 29.07.2026. | Izbor kandidata, L0–L3 dozvole, Evidence Pack, baseline, metrike, intervju, proof card/kratka/puna/anonimna forma, approval workflow, CRM i quality score. |
| `12_ROI_CALCULATOR_PLAYBOOK.md` | DRAFT v1 — VALIDATION · 30.07.2026. | Konzervativni/base/upside scenariji, TCO, direktne koristi, faktor realizacije, ramp-up, payback, break-even, sensitivity, assumption register, CRM i quality score. |
| `13_CRM_PIPELINE_PLAYBOOK.md` | DRAFT v1 — VALIDATION · 30.07.2026. | Account/contact/lead/opportunity model, S0–S9/SN stages, next-step i aging disciplina, forecast, confidence/champion test, hygiene, no-deal, nurture, handoff i CRM quality score. |
| `14_KPI_DASHBOARD_PLAYBOOK.md` | DRAFT v1 — VALIDATION · 30.07.2026. | Activity, funnel, conversion, coverage, velocity, forecast accuracy, source quality, CRM hygiene, win/loss, handoff, customer health, alerts i KPI governance. |
| `15_ANNUAL_SALES_CALENDAR.md` | DRAFT v1 — VALIDATION · 31.07.2026. | Godišnji ritam, sezonski account plan, campaign waves, kanali, Google Ads/SEO, partneri, capacity gate, nurture, customer calendar i review cadence. |

## Postojeći prodajni dokumenti

| Dokument | Verzija / datum | Sadržaj |
|---|---|---|
| `AgriX_Cenovnik_2027.html` | 27.07.2026. | **Izvor istine za cenovnik.** Cene se menjaju isključivo ovde, u `data-eur` atributima. Koristi AgriX brand tokene iz `src/styles/base.css`, fontove iz `vendor/fonts/` i logo iz `img/` — relativnim putanjama, pa fajl mora ostati u `docs/Sales/`. |
| `AgriX_Cenovnik_2027.pdf` | važi od sezone 2027 · 27.07.2026. | Generisani cenovnik za klijenta, 9 strana: paketi Desktop/Mobile i all-in varijante sa izričitim sastavom, moduli sa obračunskom jedinicom, stanice i dodatna instanca, Gazdinstvo, **dve tarife Savetnika** (standalone i Enterprise), **dve satnice** (razvojna 50 €/h i implementaciona 30 €/h), primeri obračuna, šta je uključeno u pretplatu. |
| `AgriX_Materijal_za_prvi_kontakt.pdf` | v1 · 26.07.2026. | Prodajni prozori po kulturama, tri tira i tri poruke, skripta telefonskog razgovora, email šabloni, prigovori i odgovori, šta se nikada ne obećava, evidencija posle poziva, model talasa. |
| `AgriX_Sablon_ponude.xlsx` | v1 · 26.07.2026. | Radni šablon ponude sa listom `Cenovnik` kao jedinim mestom za cene. Ponuda povlači vrednosti iz cenovnika; cene se ne kucaju u ponudu. |

Napomene:

- **Cenovnik se ne menja u PDF-u** — menja se `AgriX_Cenovnik_2027.html` pa se PDF regeneriše:

  ```bash
  tools/cenovnik.sh build    # .html -> AgriX_Cenovnik_2027.pdf
  tools/cenovnik.sh check    # poredi cene u .html sa ostala tri mesta
  ```

  Zavisnost je Chromium/Chrome (headless print-to-pdf); `CHROME_BIN` može da nadjača automatsko pronalaženje. `check` čita cene iz `data-cena`/`data-eur` atributa, pa ne puca kad se menja dizajn, i proverava da se atribut poklapa sa prikazanim tekstom.

- **Cene moraju biti identične na četiri mesta:** `AgriX_Cenovnik_2027.html`, list `Cenovnik` u šablonu ponude, Prilog 1 ugovora (`docs/Legal/AgriX_Ugovor_o_licenciranju.md`) i finansijski model. `tools/cenovnik.sh check` to proverava programski;
- cene se menjaju samo kada se promeni odluka o ceni (izvor: odluke 339, 341, 349–358, 409–422);
- šablon ponude je prazan obrazac — popunjene ponude sa podacima klijenta se ne commit-uju;
- hardverska podrška (odluka 357) i cena po gazdinstvu kod Savetnika (odluka 341) potvrđene su 27.07.2026.;
- Dispatch se nudi samo uz Mobile paket (odluka 293). **GGAP se sme prikazati samo uz vidljivu oznaku „na upit, uz potvrdu obima — nije deo standardne ponude“** (odluka 417); ostaje van redovne komercijalne ponude do validacije (odluka 405);
- **Savetnik nosi oznaku „u pripremi“ i ne kotira se kao redovna stavka** (odluka 423). Ima dve objavljene tarife — standalone 150 €/15 € i Enterprise 100 €/10 € (odluke 419, 420) — ali se ne ugovara dok proizvod ne bude stabilan (odluka 217). U cenovniku ne sme nositi zlatni okvir ni drugu vizuelnu oznaku preporuke, i ne uvrštava se automatski u ponudu;
- **nema pregovaračkih ni individualnih popusta** (odluka 418). Jedina cenovna razlika unutar istog obima je −50 % na drugu i svaku narednu instancu (odluka 413); objavljene razlike iz cenovnika nisu popusti;
- marža na hardver (~100 €/stanici, odluke 356 i 407) je **interni podatak i ne prikazuje se klijentu**.