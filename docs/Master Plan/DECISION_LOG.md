# AgriX Master Plan — Decision Log

**Status:** Active  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-27

Ovaj dokument čuva formalno odobrene strateške i governance odluke. Odluke se ne brišu; kada prestanu da važe, označavaju se kao `Superseded` i povezuju sa novom odlukom.

Numerisane odluke iz Q&A sesija vode se u `09_QA_DECISION_LOG.md`; ovde se upisuju samo one koje menjaju STR/GOV odluke.

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
- **Status:** `Superseded` 2026-07-27 — zamenjena odlukom 403 (vidi STR-014)
- **Odluka (više ne važi):** Maksimalan broj novih firmi po sezoni određuje formalni readiness score, ne unapred definisan broj.
- **Razlog povlačenja:** Rast se planira prema fiksnom ciljnom broju klijenata. Readiness ostaje preduslov kvaliteta isporuke i može zaustaviti pojedinačan onboarding, ali više ne određuje ciljni broj.

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
- **Kontekst:** AgriX već obuhvata terenski otkup, management, Dispečer, vozače, repromaterijal, otpremu, SEF, bankarske izvode, rasknjižavanje i pripremu naloga za plaćanje.
- **Odluka:** AgriX se razvija i pozicionira kao end-to-end vertikalni poslovni operativni sistem koji pokriva sve glavne i ključne sporedne tokove ciljnog klijenta.
- **Obuhvat:** Kooperanti, parcele, repromaterijal, otkup, stanice, dokumentacija, ambalaža, logistika, Dispečer, vozači, lager, sledljivost, kupci, otprema, fakture, SEF, banka, naplate, isplate, nalozi za plaćanje, management i monitoring.
- **Posledice:** Product roadmap i pricing polaze od poslovnih tokova, ne od liste ekrana ili tehničkih aplikacija.

### STR-011 — Gazdinstvo kao pun farm-management proizvod
- **Datum:** 2026-07-22
- **Status:** Approved
- **Kontekst:** Kooperant rola već sadrži karticu prema hladnjači, GIS i parcele, parcelnu prognozu i upozorenja, digitalnog agronoma, pametno doziranje, tretmane, opremu, troškove, proizvodnju i sezonski bilans.
- **Odluka:** AgriX Gazdinstvo se tretira kao zaseban pun farm-management proizvod, a ne kao portal ili mali dodatak Enterprise sistemu.
- **Posledice:** Gazdinstvo dobija sopstveni product strategy, activation funnel, packaging, KPI-jeve, support model i unit economics.

### STR-012 — GGAP kao treći proizvodni stub
- **Datum:** 2026-07-22
- **Status:** `Superseded` 2026-07-27 — zamenjena odlukama 401 i 402 (vidi STR-013)
- **Kontekst:** Enterprise i Gazdinstvo već stvaraju veliki deo operativnih, parcelnih, agronomskih, robnih i dokumentacionih podataka potrebnih za GGAP evidencije.
- **Odluka (više ne važi):** AgriX GGAP je treći puni proizvod, pored Enterprise i Gazdinstvo. Pokriva GGAP liste, evidencije, dokaze, zadatke, neusaglašenosti, korektivne mere i kompletan dokumentacioni tok.
- **Razlog povlačenja:** GGAP nema sopstveni ICP — koriste ga isključivo hladnjače koje su već Enterprise klijenti. Zato je modul u okviru Enterprise-a, a treći stub je AgriX Savetnik.
- **Šta ostaje:** Funkcionalni obuhvat i osnovni princip (podatak se unosi jednom na mestu nastanka, evidencija i dokaz se izvode automatski) ostaju na snazi kao obuhvat modula. Softver ne garantuje sertifikaciju i ne zamenjuje sertifikaciono telo, auditora ili stručnog konsultanta.
- **Šta otpada:** Zaseban stub-level product strategy, packaging i unit economics. GGAP se vodi u Enterprise ekonomici.

### STR-013 — AgriX Savetnik kao treći proizvodni stub
- **Datum:** 2026-07-27
- **Status:** Approved
- **Izvor:** odluke 401 i 402; potvrđuje odluku 269
- **Kontekst:** Krovni brend AgriX već je definisan kroz tri proizvoda — Enterprise, Gazdinstvo i Savetnik. Poglavlja 02 i 07 su na mestu trećeg stuba i dalje vodila GGAP.
- **Odluka:** Treći ravnopravan proizvodni stub je AgriX Savetnik — upravljački sloj nad većim brojem gazdinstava, za agronome i savetodavne službe. GGAP je modul u okviru Enterprise-a.
- **Posledice:** Savetnik dobija sopstveni product strategy, packaging, KPI-jeve i unit economics; GGAP ih gubi kao zaseban proizvod. Savetnik ne može javno na tržište bez samostalne registracije i politike privatnosti, a uloga za tok T13 ostaje otvorena do razrešenja LEG1.
- **Ponovno otvaranje:** Ako se pokaže da GGAP ima kupce van postojeće Enterprise baze.

### STR-014 — Fiksan ciljni broj klijenata
- **Datum:** 2026-07-27
- **Status:** Approved
- **Izvor:** odluka 403; ciljni broj iz odluke 375 (scenario C)
- **Kontekst:** Readiness-based cap (STR-001) nije davao broj oko kojeg se planiraju prodaja, kapacitet i finansije.
- **Odluka:** Rast se planira prema fiksnom ciljnom broju klijenata. Aktuelan cilj je 17–18 aktivnih Enterprise firmi do sezone 2027 (14–15 novih uz postojeće 3).
- **Posledice:** Readiness prelazi iz cap mehanizma u kontrolnu listu pred svaki onboarding. Crven P0 domen zaustavlja pojedinačan onboarding, a odgovor je otklanjanje uzroka ili dodavanje kapaciteta, ne spuštanje cilja. Isti broj mora ostati usklađen u `02_STRATEGY.md` §9, `04_MARKET.md` §9.1 i `docs/Finance/AgriX_Finansijski_model.xlsx`.
- **Rizik i trigger:** Ako se posle sezone 2027 pokaže da fiksni cilj obara kvalitet isporuke ili stopu obnove, odluka se preispituje.
