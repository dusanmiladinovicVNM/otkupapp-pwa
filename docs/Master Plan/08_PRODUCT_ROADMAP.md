# 08 — Product roadmap AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 23.07.2026–31.12.2027  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `06_POSITIONING.md`, `07_PRODUCT_PORTFOLIO.md`, `07A_PRODUCT_STATUS_MATRIX.csv`, `09_TECHNOLOGY_STRATEGY.md`, `10_PRICING_AND_PACKAGING.md`, `14_GO_TO_MARKET.md`, `16_ONBOARDING_AND_IMPLEMENTATION.md`, `../ROADMAP.md`, `../KNOWN_ISSUES.md`, `../RELEASE_GATES.md`

---

## 1. Centralna roadmap odluka

AgriX roadmap ima dva jednako važna glavna toka:

1. **pouzdanost i istinitost podataka/release-a**;
2. **PWA-led teren–centrala operativni tok**.

Oni nisu konkurenti za pažnju. P0/P1 hardening postoji da bi zaštitio glavni model proizvoda:

> otkupljivač i vozač unose osnovne poslovne događaje na terenu, sistem automatski puni centralu, a operater radi kontrolu, fakture i izveštaje.

Desktop-only productization nije cilj AgriX-a.

---

## 2. Redosled prioriteta

1. zatvoriti potvrđene P0 rizike koji mogu ugroziti terenski unos, sync, dokumente ili centralnu bazu;
2. učvrstiti PWA Otkupac i PWA Vozač kao standardne operativne kanale;
3. standardizovati onboarding kompletne celine teren–centrala;
4. odvojeno standardizovati kiosk, tablet i termalnu štampu;
5. meriti realnu sezonu i povećavati broj stanica;
6. productizovati opcione Enterprise module;
7. validirati Gazdinstvo;
8. pokrenuti GGAP discovery kada core kapacitet nije ugrožen.

`DECISION`: nijedna nova funkcija nema prioritet nad P0 problemom, ali PWA terenski tok ima prioritet nad desktop-only širenjem i nad novim sporednim modulima.

---

## 3. Roadmap principi

### 3.1 Teren je mesto nastanka podatka

Svaka inicijativa se procenjuje prema tome da li:

- omogućava unos na mestu nastanka;
- uklanja ponovno prekucavanje;
- ubrzava dolazak podatka u centralu;
- smanjuje rad centralnog operatera na osnovnom unosu;
- povećava kontrolu i sledljivost.

### 3.2 PWA i desktop su jedan proizvod

PWA i desktop nisu dva nezavisna roadmap-a. Uspeh se meri end-to-end:

- unos na stanici;
- lokalno čuvanje/offline;
- sync;
- import u canonical centralu;
- dokumentna veza;
- prijem/faktura/izveštaj;
- bez ručnog popravljanja istog događaja.

### 3.3 Hardver se ne meša sa PWA zrelošću

Kiosk, tablet i printer mogu imati niži readiness status od PWA aplikacije. Neuspeh standardizacije jednog printera ne znači da PWA Otkupac nije vredan ili operativan.

### 3.4 Gate pre datuma

Datumi su ciljni prozori. Faza se završava tek kada prođu acceptance kriterijumi.

### 3.5 Jedan codebase

Nema trajnih klijentskih forkova. Novi proces mora biti konfiguracija, reusable funkcija ili jasno ograničen adapter.

### 3.6 Sezonski freeze

Najmanje 30 dana pre kritične sezone uvodi se freeze za pogođeni tok. Dozvoljene su samo P0/P1 ispravke, konfiguracija, dokumentacija i promene koje smanjuju rizik.

---

## 4. Polazni status

| Celina | Status |
|---|---|
| PWA Otkupac | implementiran i funkcionalan u aktivnom scope-u; `Standard offer` ili `Controlled rollout` prema release evidence-u |
| PWA Vozač | implementiran za aktivni transportni scope; deo end-to-end Enterprise toka |
| Sync/MasterSync | centralna platformna zavisnost; zahteva najviši correctness prioritet |
| Desktop backoffice | production centralni kontrolni, dokumentni i finansijski sloj |
| Management PWA | production kanal u okviru Enterprise-a |
| Kiosk/tablet standardizacija | `Controlled rollout` |
| Termalna štampa | zaseban readiness tok po podržanom hardveru |
| Gazdinstvo | kontrolisani pilot |
| GGAP | discovery / `Not for sale` kao završen proizvod |

---

## 5. Track A — Core safety i release truth

**Ciljni prozor:** odmah i kontinuirano; prvi closeout do 31.08.2026.

### Obavezni ishodi

- potvrditi i zatvoriti aktivne P0 data-safety nalaze;
- zaštititi JSON parsing i PWA→centrala import;
- zatvoriti potvrđene SEF statusne defekte;
- izvršiti VBA, GAS, PWA, monitoring i Production Health gate-ove;
- potvrditi autorizaciju `saveParcelPolygon`;
- rešiti `ART_POCETNI_DUG` phantom projekciju;
- čuvati release evidence: verzija, commit, workbook, environment, testovi i poznati rizici;
- ukloniti false-green release rezultate.

### Exit gate

- nema nepotvrđenog otvorenog P0;
- target workbook ima nula health failure-a;
- obavezni end-to-end tok stanica→centrala prolazi;
- release evidence je sačuvan i proverljiv;
- known issues i containment su ažurni.

### Stop pravilo

Dok potvrđeni P0 postoji, nema novih GGAP/Gazdinstvo premium funkcija niti novih nepovezanih integracija. PWA razvoj je dozvoljen kada direktno zatvara correctness ili rollout gate.

---

## 6. Track B — PWA-led Enterprise Core productization

**Ciljni prozor:** 23.07.2026–31.10.2026.

### 6.1 End-to-end canonical tok

Mora biti eksplicitno dokumentovan i testiran:

1. otkupljivač unosi otkup;
2. PWA dodeljuje stabilan client identity;
3. zapis ostaje bezbedan offline;
4. GAS prihvata idempotentni sync;
5. Google/transport sloj čuva operativni zapis;
6. MasterSync prenosi zapis u centralni desktop;
7. canonical dokumentni lanac nastavlja bez ponovnog unosa;
8. centralni operater rešava samo izuzetke, prijem, fakture i finansije.

Za vozača isto važi za zbirni transportni tok.

### 6.2 Standardizovani scope

- podržane varijante otkupa;
- podržane klase, bruto/neto i ambalaža;
- pravila brojeva dokumenata;
- correction/storno ugovor;
- sync status i recovery;
- odgovornost PWA, GAS, Sheets i desktop sloja;
- jasni unsupported slučajevi.

### 6.3 Metrike

- procenat terenskih otkupa koji ulaze bez centralnog re-entry-ja;
- sync success rate;
- duplicate rate;
- data-loss rate;
- vreme unos→centrala;
- broj ručnih centralnih korekcija na 100 otkupa;
- vreme centralnog operatera na 100 otkupa;
- support minuti po stanici.

### 6.4 Exit gate

- kompletan tok prolazi bez ručnog popravljanja baze;
- postoje regression testovi za retry, stale syncing, duplicate i partial failure;
- operater vidi izuzetke i ne mora da rekonstruiše normalan tok;
- najmanje jedan realan klijent koristi dokumentovani scope;
- onboarding i support procedura postoje.

---

## 7. Track C — Repeatable onboarding teren–centrala

**Ciljni prozor:** 01.08.2026–30.11.2026.

Standardni onboarding mora obuhvatiti:

- discovery procesa;
- firme, stanice, korisnike i uloge;
- PWA uređaje i pristupe;
- master podatke i cenovnike;
- sync/GAS/Sheets konfiguraciju;
- desktop build i centralnu bazu;
- obuku otkupljivača;
- obuku vozača;
- obuku centralnog operatera;
- management pristup;
- go-live, fallback i rollback;
- monitoring i support handoff.

### Exit gate

- jedan onboarding izveden po checklisti;
- nema founder-only skrivenih koraka;
- poznato vreme po stanici i po firmi;
- greške i odstupanja su izmereni;
- podrška može da dijagnostikuje tipične PWA/sync probleme.

---

## 8. Track D — Kiosk, tablet i termalna štampa

**Ciljni prozor:** Q3 2026–Q1 2027.

Ovaj track je odvojen od osnovnog statusa PWA aplikacije.

### Kiosk/tablet

- odobren minimalni profil uređaja;
- kiosk zaključavanje;
- remote support;
- asset evidencija;
- rezervni uređaj;
- recovery posle restarta i gubitka veze.

### Termalna štampa

- odobreni printer modeli;
- stabilan print path;
- potvrda uspeha/retry;
- fallback kada printer nije dostupan;
- garancija i zamena;
- poznat support cost.

### Exit gate

- setup se ponavlja bez posebnog koda;
- print/recovery test prolazi na odobrenom modelu;
- hardver ima odvojenu cenu i maržu;
- neuspeh printera ne ugrožava očuvanje poslovnog zapisa.

---

## 9. Track E — Sezonsko skaliranje

**Ciljni prozor:** prema kulturi i sezoni klijenata, 2026–2027.

Cilj nije prvi dokaz da PWA radi, već dokaz koliko stanica i obima može pouzdano da podrži.

### Rollout model

- postojeći klijent ili kvalifikovani novi klijent;
- postepeno širenje stanica;
- definisan dnevni obim;
- dnevni monitoring u kritičnom periodu;
- success, containment i abort kriterijumi;
- feature freeze pre vrhunca sezone.

### Radni pragovi

- nula izgubljenih potvrđenih zapisa;
- nula nekontrolisanih canonical duplikata;
- najmanje 99% sync uspeha bez developerske intervencije;
- svi P0 incidenti zatvoreni root-cause analizom;
- ručne centralne korekcije ispod dogovorenog praga;
- support cost održiv za planirani broj stanica.

Print KPI se vodi samo za klijente i stanice koje koriste AgriX termalni print paket.

---

## 10. Track F — Opcione Enterprise ekstenzije

Finance, Logistics, Agrohemija i Warehouse dobijaju kapacitet kada:

- aktivni klijent ima blocking problem;
- više klijenata traži isti tok;
- kvalifikovani kupac finansira reusable rad;
- promena smanjuje support ili data-risk;
- ekstenzija direktno jača teren–centrala model.

Nova banka, ERP, vaga ili printer integracija zahteva testne podatke, finansiran razvoj, acceptance, maintenance i fallback.

---

## 11. Track G — Gazdinstvo validation

Gazdinstvo se razvija ograničeno dok PWA-led Enterprise nema stabilan rollout.

Meriti:

- pozvani→registrovani;
- registrovani→prvi unos;
- 30/90/180-day active rate;
- retenciju između sezona;
- Partner→Basic/Pro konverziju;
- WTP;
- support cost;
- uticaj na kvalitet Enterprise podataka.

Ne širiti AI/Digitalni agronom i premium scope bez dokaza i stručne odgovornosti.

---

## 12. Track H — GGAP discovery

Preduslovi:

- stručni domain owner;
- standard i verzija;
- imenovani pilot-klijent;
- mapiranje polja na Enterprise/Gazdinstvo izvore;
- pravne i marketinške granice;
- pricing i support hipoteza.

Dozvoljeni prvi scope: readiness dashboard, nedostajući dokazi, rokovi, provenance i kontrolisani export.

---

## 13. Capacity budget

### Dok postoje potvrđeni P0 rizici

| Oblast | Udeo |
|---|---:|
| Core correctness, release i sync data-safety | 35% |
| PWA-led teren–centrala productization | 35% |
| Onboarding i support enablement | 20% |
| Kiosk/print standardizacija | 7% |
| Gazdinstvo/GGAP discovery | 3% |

### Posle P0 closeout-a

| Oblast | Udeo |
|---|---:|
| PWA-led Enterprise i sezonsko skaliranje | 45% |
| Enterprise correctness i opcione ekstenzije | 20% |
| Onboarding, migracija i support alati | 20% |
| Kiosk/print productization | 8% |
| Gazdinstvo | 5% |
| GGAP discovery | 2% |

Ovo su zaštitne smernice, ne timesheet kvote.

---

## 14. Uloga dodatnog developera

Prioritetni ownership:

1. PWA/GAS regression i fixture okruženje;
2. release tooling i evidence;
3. sync/import correctness;
4. jasno ograničene P1 popravke;
5. onboarding alati i adapteri;
6. kiosk/print stabilizacija.

Osnivač zadržava product prioritization, kritični document/finance model i release approval dok review ownership nije dokazano prenet.

---

## 15. Roadmap anti-prioriteti

Ne prioritetizovati:

- desktop-only funkcije koje povećavaju centralni unos umesto da ga smanjuju;
- nove premium Gazdinstvo funkcije bez usage dokaza;
- puni GGAP bez domain owner-a;
- generički ERP/TMS/WMS scope;
- nepovezane integracije bez klijenta;
- novi hardver bez jasnog support i warranty modela;
- kozmetičke promene koje ne poboljšavaju field-to-office tok.

---

## 16. Odluke

1. PWA-led teren–centrala tok je srž AgriX roadmap-a.
2. Core safety i PWA productization imaju jednak strateški prioritet.
3. PWA Otkupac i Vozač ne čekaju 2027. da postanu relevantni; već se productizuju i skaliraju prema aktivnom dokazu.
4. Kiosk i termalna štampa imaju zasebne roadmap gate-ove.
5. Primarni KPI nije broj novih ekrana, već procenat poslovnih događaja koji od terena do centrale prolaze bez ponovnog unosa.
6. Centralni operater treba da troši vreme na kontrolu, fakture i izveštaje, ne na prepisivanje.
