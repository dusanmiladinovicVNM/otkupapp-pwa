# AgriX Master Plan — Decision Log

**Status:** Active  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-22

Ovaj dokument čuva formalno odobrene strateške i governance odluke. Odluke se ne brišu; kada prestanu da važe, označavaju se kao `Superseded` i povezuju sa novom odlukom.

## Format odluke

Svaka odluka treba da sadrži:

- ID;
- datum;
- status;
- kontekst;
- odluku;
- razlog;
- posledice;
- kriterijum za ponovno otvaranje.

---

## GOV-001 — Jezik Master Plana

- **Datum:** 2026-07-22
- **Status:** Approved
- **Kontekst:** Dokument treba da bude svakodnevno upotrebljiv osnivaču i budućem domaćem timu.
- **Odluka:** Master Plan se vodi na srpskom jeziku.
- **Razlog:** Smanjuje trenje pri radu i omogućava precizno zapisivanje domaćih tržišnih i operativnih specifičnosti.
- **Posledice:** Eksterni engleski sažetak pravi se samo kada postoji konkretna potreba.
- **Ponovno otvaranje:** Kada većina rukovodećeg tima više ne koristi srpski kao radni jezik.

## GOV-002 — Planska valuta

- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Osnovna planska valuta je EUR; RSD se koristi za lokalne poreske, platne i gotovinske tokove.
- **Razlog:** Cene softvera, hardvera i buduća regionalna poređenja lakše su uporedivi u EUR, dok se stvarni domaći rashodi i obaveze često realizuju u RSD.
- **Posledice:** Svaki finansijski model mora navesti kursnu pretpostavku kada kombinuje EUR i RSD.
- **Ponovno otvaranje:** Ako se promeni dominantna ugovorna ili računovodstvena valuta poslovanja.

## GOV-003 — Način razvoja poglavlja

- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Svako veliko poglavlje razvija se kroz zaseban PR ili jasno izdvojen commit.
- **Razlog:** Omogućava fokusiranu reviziju, istoriju promena i vraćanje odluka bez mešanja tema.
- **Posledice:** Velika kombinovana ažuriranja bez jasne granice izbegavaju se.
- **Ponovno otvaranje:** Ako obim dokumentacije postane toliko mali da odvojeni PR-ovi stvaraju više troška nego koristi.

## GOV-004 — Anonimizacija klijenata

- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Klijenti se u strateškim dokumentima anonimizuju.
- **Razlog:** Smanjuje reputacioni i poverljivi rizik i omogućava iskrenu analizu problema.
- **Posledice:** Koriste se oznake poput `Klijent A`, `Klijent B`, kultura i segment, bez direktnih naziva firmi.
- **Ponovno otvaranje:** Samo uz eksplicitnu dozvolu klijenta za javnu studiju slučaja.

## GOV-005 — Osetljivi podaci

- **Datum:** 2026-07-22
- **Status:** Approved
- **Odluka:** Osetljivi finansijski, ugovorni i identifikacioni podaci izdvajaju se iz javnog tehničkog repozitorijuma.
- **Razlog:** Master Plan treba da bude iskren i numerički precizan bez izlaganja poverljivih informacija.
- **Posledice:** Javni repo sadrži metodologiju i anonimizovane vrednosti; detaljni modeli idu u privatni repo ili privatni dodatak.
- **Ponovno otvaranje:** Kada se promeni status repozitorijuma ili se uvede formalna kontrola pristupa.

---

## STR-001 — Kontrolisan rast pre agresivnog skaliranja

- **Datum:** 2026-07-22
- **Status:** Draft for strategy approval
- **Kontekst:** Postoje tri aktivna klijenta; cilj za narednu sezonu je približno pet novih, ukupno oko osam, uz svesnu granicu od najviše deset.
- **Predložena odluka:** Naredna sezona se tretira kao sezona dokazivanja operativne skale, ne kao sezona maksimalne prodaje.
- **Razlog:** PWA Otkupac, kiosk terminali i termalna štampa uvode novi operativni rizik koji mora biti potvrđen na ograničenom broju firmi.
- **Posledice:** Prodaja preko deset firmi odlaže se čak i ako postoji interesovanje, osim ako kapacitet i readiness budu ponovo potvrđeni.
- **Ponovno otvaranje:** Nakon završene pune sezone sa izmerenim onboardingom, supportom, incidentima i korišćenjem terminala.
