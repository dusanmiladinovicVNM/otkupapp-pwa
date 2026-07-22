# 00 — Upravljanje AgriX Master Planom

**Status:** Approved  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-22  
**Sledeći pregled:** pre odobravanja `02_STRATEGY.md`

---

## 1. Svrha

AgriX Master Plan je interni sistem za donošenje odluka o proizvodu, tržištu, cenama, finansijama, organizaciji i rastu. Nije investitorski pitch, marketing brošura niti skup optimističnih projekcija.

Njegov zadatak je da:

- pretvori rasute ideje i podatke u proverljive odluke;
- jasno odvoji činjenice, pretpostavke i hipoteze;
- pokaže finansijske i operativne posledice svake važne odluke;
- spreči rast koji prevazilazi kapacitet proizvoda i organizacije;
- omogući novom saradniku da razume zašto je odluka doneta;
- čuva istoriju promena bez prepravljanja prošlosti.

Kada je neprijatan zaključak bolje potkrepljen od poželjnog zaključka, neprijatan zaključak ima prednost.

---

## 2. Publika

Primarna publika su osnivač, budući developer, customer support / implementation osoba i budući rukovodioci funkcija. Odabrani delovi mogu se dati računovođi, pravniku, banci, partneru ili investitoru, ali interna verzija se ne ulepšava za eksternu publiku.

---

## 3. Izvor istine

Prioritet izvora:

1. produkcijski podaci, ugovori, računi i stvarne metrike;
2. aktuelni kod, konfiguracija i operativni runbook-ovi;
3. potvrđene izjave klijenata;
4. zvanični javni izvori;
5. interne procene sa navedenom metodologijom;
6. hipoteze koje tek treba testirati.

Kada su izvori u konfliktu, konflikt se navodi. Ne bira se broj koji više odgovara željenoj priči.

---

## 4. Klasifikacija tvrdnji

| Oznaka | Značenje | Dozvoljena upotreba |
|---|---|---|
| **FACT** | Direktno potvrđena činjenica | Može biti osnova odluke |
| **MEASURED** | Izmereno na ograničenom uzorku | Navesti uzorak i period |
| **ASSUMPTION** | Radna pretpostavka | Obavezna sensitivity analiza |
| **HYPOTHESIS** | Verovanje koje nije potvrđeno | Mora imati test i rok |
| **TARGET** | Željeni rezultat | Ne prikazivati kao prognozu |
| **DECISION** | Formalno odobrena odluka | Mora imati datum i razlog |
| **UNKNOWN** | Nedostaje pouzdan podatak | Ne popunjavati izmišljenom preciznošću |

Primeri:

- `FACT`: postoje tri aktivna klijenta.
- `MEASURED`: približno jedan support poziv nedeljno tokom posmatranog perioda.
- `ASSUMPTION`: onboarding će nakon standardizacije trajati pola dana.
- `HYPOTHESIS`: 5% kooperanata će kupiti Gazdinstvo Basic ili Pro.
- `TARGET`: narednu sezonu završiti sa najviše 8–10 hladnjača.

---

## 5. Statusi i odobravanje

Statusi su `Draft`, `Review`, `Approved`, `Superseded` i `Archived`.

Poglavlje prelazi u `Review` tek kada sadrži činjenice, pretpostavke, kontra-tezu, finansijske ili kapacitivne posledice, rizike, KPI-jeve i preporučene odluke. U `Approved` prelazi tek nakon eksplicitne potvrde osnivača i upisa ključnih odluka u `DECISION_LOG.md`.

---

## 6. Finansijska disciplina

Finansijski model mora odvojiti recurring prihod, implementaciju, obuku i hardver. Kod hardvera se paralelno prikazuju prodajna i nabavna cena, konfiguracija, transport, garancija, zamene, zaliha i obrtni kapital.

Plate se prikazuju kao pun trošak poslodavca. Rad osnivača ima ekonomsku cenu čak i kada se ne isplaćuje. Svaka projekcija ima konzervativni, bazni i optimistični scenario.

---

## 7. Ritam pregleda

- mesečno: prodaja, support, korišćenje i nove pretpostavke;
- kvartalno: pricing, troškovi, roadmap i kapacitet;
- pred sezonu: release readiness, support raspored, hardver i maksimalni broj novih klijenata;
- posle sezone: support sati, incidenti, onboarding trošak, churn i ekonomika po klijentu.

---

## 8. Odobrene governance odluke

| ID | Odluka | Status |
|---|---|---|
| GOV-001 | Master Plan se vodi na srpskom jeziku | Approved |
| GOV-002 | Osnovna planska valuta je EUR; RSD se koristi za lokalne tokove | Approved |
| GOV-003 | Svako veliko poglavlje razvija se kroz zaseban PR ili jasno izdvojen commit | Approved |
| GOV-004 | Klijenti se anonimizuju u strateškim dokumentima | Approved |
| GOV-005 | Osetljivi finansijski i ugovorni podaci izdvajaju se iz javnog tehničkog repoa | Approved |

Odluke su potvrđene 2026-07-22.

---

## 9. Pravila protiv fluff-a

Master Plan ne sme da tretira cilj kao prognozu, TAM kao očekivanu prodaju, promet klijenta kao automatsko opravdanje cene, hardver kao čistu zaradu, Gazdinstvo kao značajan kratkoročni prihod bez konverzije, zapošljavanje bez dokazanog uskog grla ili investiciju bez precizne upotrebe kapitala.

---

## 10. Sledeći koraci

1. održavati `DECISION_LOG.md`;
2. održavati `CHANGELOG.md`;
3. razviti i odobriti `02_STRATEGY.md`;
4. potom razvijati kupce, portfolio, pricing i finansijski model;
5. `01_EXECUTIVE_SUMMARY.md` napisati poslednji.
