# AgriX Master Plan — Decision Log

**Status:** Active  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-22

Ovaj dokument čuva formalno odobrene strateške i governance odluke. Odluke se ne brišu; kada prestanu da važe, označavaju se kao `Superseded` i povezuju sa novom odlukom.

---

## Governance odluke

### GOV-001 — Jezik Master Plana
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Master Plan se vodi na srpskom jeziku.

### GOV-002 — Planska valuta
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Osnovna planska valuta je EUR; RSD se koristi za lokalne poreske, platne i gotovinske tokove.

### GOV-003 — Način razvoja poglavlja
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Svako veliko poglavlje razvija se kroz zaseban PR ili jasno izdvojen commit.

### GOV-004 — Anonimizacija klijenata
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Klijenti se u strateškim dokumentima anonimizuju.

### GOV-005 — Osetljivi podaci
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Osetljivi finansijski, ugovorni i identifikacioni podaci izdvajaju se iz javnog tehničkog repozitorijuma.

---

## Strateške odluke

### STR-001 — Readiness-based rast
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Maksimalan broj novih firmi po sezoni određuje formalni readiness score, ne unapred definisan broj.

### STR-002 — Primarni tržišni fokus
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Trenutni geografski fokus je Srbija. Primarni kupci su hladnjače i druge firme sa razgranatom mrežom otkupnih stanica i kooperanata.

### STR-003 — Jedan proizvod, jedan kod
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Sve klijentske razlike rešavaju se zajedničkim kodom i konfiguracijom. Trajni klijentski fork nije dozvoljen.

### STR-004 — Dinamična uloga Gazdinstva
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Licence firmi trenutno finansiraju osnovni biznis. Gazdinstvo se kratkoročno ne tretira kao ključni prihod, ali može postati glavni proizvod ili izvor prihoda ako podaci potvrde ekonomiku.

### STR-005 — Hardver kao sporedni profitni centar
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Hardver nije glavni profitni centar, ali mora biti profitabilan sporedni centar. AgriX može postati dobavljač šireg IT sistema ciljnih klijenata.

### STR-006 — Bez partnera samo zbog novca
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Partner ili investitor razmatra se samo kada rešava dokazano usko grlo i donosi merljivu sposobnost pored kapitala.

### STR-007 — Prvo operativno zaposlenje
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Prva operativna uloga je customer support / implementation.

### STR-008 — Regionalna platforma
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Dugoročni cilj AgriX-a je regionalna vertikalna platforma.

### STR-009 — Strateški cilj tržišnog udela
- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** AgriX cilja najmanje 200 firmi u naredne 3–4 godine.

### STR-010 — End-to-end poslovni operativni sistem
- **Datum:** 2026-07-22
- **Status:** Approved
- **Kontekst:** AgriX već obuhvata terenski otkup, management, Dispečer, vozače, repromaterijal, otpremu, SEF, bankarske izvode, rasknjižavanje i pripremu naloga za plaćanje. Tretiranje tih funkcija kao sporednih modula potcenjuje stvarni obim proizvoda.
- **Odluka:** AgriX se razvija i pozicionira kao end-to-end vertikalni poslovni operativni sistem koji pokriva sve glavne i ključne sporedne tokove ciljnog klijenta.
- **Obuhvat:** Kooperanti, parcele, repromaterijal, otkup, stanice, dokumentacija, ambalaža, logistika, Dispečer, vozači, lager, sledljivost, kupci, otprema, fakture, SEF, banka, naplate, isplate, nalozi za plaćanje, management i monitoring.
- **Posledice:** Product roadmap i pricing polaze od poslovnih tokova, ne od liste ekrana ili tehničkih aplikacija.

### STR-011 — Gazdinstvo kao pun farm-management proizvod
- **Datum:** 2026-07-22
- **Status:** Approved
- **Kontekst:** Kooperant rola već sadrži karticu prema hladnjači, GIS i parcele, parcelnu prognozu i upozorenja, digitalnog agronoma, pametno doziranje, tretmane, opremu, troškove, proizvodnju i sezonski bilans.
- **Odluka:** AgriX Gazdinstvo se tretira kao zaseban pun farm-management proizvod, a ne kao portal ili mali dodatak Enterprise sistemu.
- **Obuhvat:** Kartica sa zaduženjima, razduženjima i saldom; GIS i parcele; meteo i rizici po parceli; pametno doziranje; tretmani; oprema; knjiga polja; troškovi po kategorijama i parcelama; proizvodnja; bilans ukupno i po parceli; offline-first rad.
- **Posledice:** Gazdinstvo dobija sopstveni product strategy, activation funnel, packaging, KPI-jeve, support model i unit economics. Njegov prihod se kratkoročno ne precenjuje, ali se njegov dugoročni potencijal ne ograničava B2B licencom hladnjače.
- **Ponovno otvaranje:** Nakon merenja aktivacije, retencije, willingness-to-pay i stvarnog support cost-a.
