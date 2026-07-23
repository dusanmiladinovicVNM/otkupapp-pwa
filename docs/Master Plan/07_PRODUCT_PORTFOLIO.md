# 07 — Portfolio proizvoda AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `02_STRATEGY.md`, `03_CUSTOMERS_AND_JOBS.md`, `06_POSITIONING.md`, `08_PRODUCT_ROADMAP.md`, `10_PRICING_AND_PACKAGING.md`, `14_GO_TO_MARKET.md`, `15_SALES_PLAYBOOK.md`, `16_ONBOARDING_AND_IMPLEMENTATION.md`

---

## 1. Svrha

Portfolio definiše šta kupac stvarno kupuje, koje komponente moraju ostati zajedno i šta se ne sme pogrešno izdvojiti kao sporedni dodatak.

Ključna odluka:

> **AgriX Enterprise Core je end-to-end sistem teren–centrala, a ne desktop proizvod sa opcionom PWA aplikacijom.**

Otkupci i vozači stvaraju osnovne poslovne podatke i dokumente na terenu. Sinhronizacioni sloj ih prenosi u centralnu bazu. Centralni operater završava kontrolu, prijem, fakture, finansije i izveštavanje.

---

## 2. Statusi proizvoda

Za svaku celinu odvojeno se vode:

### Implementacioni status

- `Implemented`
- `Partial`
- `Planned`
- `Deprecated`

### Dokazni status

- `Production-proven`
- `Limited production evidence`
- `Pilot evidence`
- `Unvalidated`

### Komercijalni status

- `Standard offer`
- `Optional extension`
- `Controlled rollout`
- `Pilot only`
- `Not for sale`

Postojanje koda nije jedini kriterijum, ali ni nedovoljno standardizovan hardver ne sme da obori status dobro funkcionalne PWA aplikacije.

---

## 3. Arhitektura portfolija

AgriX ima tri proizvodna stuba:

1. **AgriX Enterprise** — povezani terenski i centralni operativni sistem firme;
2. **AgriX Gazdinstvo** — farm-management sistem kooperanta/proizvođača;
3. **AgriX GGAP** — dokumentacioni i compliance workflow.

Zajednički tehnički sloj je **AgriX Platform Services**.

---

## 4. AgriX Platform Services

Platform Services je ugrađen u svaki plaćeni proizvod i obuhvata:

- zajednički codebase i konfiguraciju;
- tenant, firmu, stanicu, korisnika i uloge;
- autentifikaciju i autorizaciju;
- offline queue, retry, idempotency i sync;
- MasterSync i centralizaciju podataka;
- transakcione kontrole, storno i audit;
- monitoring, logging, self-update, backup i recovery;
- release gates i Production Health Check;
- support dijagnostiku.

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` |
| Dokaz | `Production-proven`, uz ograničen broj klijenata |
| Komercijalno | uključeno u plaćene proizvode |

---

## 5. AgriX Enterprise Core

### 5.1 Poslovna definicija

Enterprise Core obuhvata ceo minimalni poslovni lanac:

> **otkupac na stanici → osnovni dokument → vozač i transport → sinhronizacija → centralna baza → prijem → faktura → izveštaj i management kontrola**

Core nije završen ako centralni operater mora rutinski da ponovo unosi otkupe koje su terenski korisnici već evidentirali.

### 5.2 Terenski operativni kanali

#### PWA Otkupac

- identifikacija kooperanta;
- stanica, kultura, klasa, cena i ambalaža;
- bruto/neto i podržane varijante otkupa;
- kreiranje osnovnog otkupnog dokumenta;
- offline rad i lokalni queue;
- kontrolisana sinhronizacija u centralu;
- status pending/syncing/synced/error;
- korekcija/storno prema podržanom ugovoru.

#### PWA Vozač

- preuzimanje robe sa stanica;
- zbirni dokument i broj zbirne;
- količine, poslednja stanica i status transporta;
- sinhronizacija prema centrali;
- podrška dispečerskom i prijemnom toku.

PWA Otkupac i PWA Vozač su **glavni izvori terenskih poslovnih događaja**, a ne marketinški dodatak desktop proizvodu.

### 5.3 Centralni backoffice

Centralni desktop sloj obuhvata:

- master podatke i pripremu sezone;
- kontrolu i import sinhronizovanih terenskih događaja;
- prijemnicu i povezivanje dokumentnog lanca;
- fakture i stavke;
- SEF;
- banku, avanse, salda, rasknjižavanje i naloge za plaćanje;
- lager, ambalažu, repromaterijal i relevantne kartice;
- izveštaje, monitoring, audit i storno.

Centralni operater treba da radi **kontrolu i završnu obradu**, ne rutinski osnovni unos sa terena.

### 5.4 Management PWA

Management PWA je deo Enterprise Core-a i daje:

- pregled stanica, količina i dokumenata;
- stanje transporta i prijema;
- finansijske i kartične preglede;
- upozorenja i odstupanja;
- kontrolisane operacije prema ulozi.

### 5.5 Status Enterprise Core-a

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented` za aktivni end-to-end scope |
| Dokaz | `Production-proven / limited production evidence` po konkretnom toku i klijentu |
| Komercijalno | `Standard offer` ili `Controlled rollout` prema release evidence-u |

`DECISION`: PWA deo ne izdvaja se automatski u „pilot only“ ako već radi u realnom procesu. Status se određuje po konkretnom toku i dokazu, ne po najnezrelijoj hardverskoj komponenti.

---

## 6. Delivery komponente odvojene od PWA proizvoda

### 6.1 Kiosk režim

- zaključavanje uređaja;
- standardna konfiguracija;
- remote pristup i podrška;
- recovery posle restarta ili gubitka veze.

### 6.2 Tablet paket

- odobreni modeli;
- zaštita, napajanje i stalak;
- asset evidencija;
- zamena i garancija.

### 6.3 Termalna štampa

- podržani printeri;
- printer bridge ili drugi stabilan način štampe;
- retry i potvrda uspeha;
- rezervni uređaj i fallback.

Ove komponente imaju odvojene statuse:

| Celina | Komercijalni status |
|---|---|
| PWA Otkupac | `Standard offer` / `Controlled rollout` prema aktivnom release scope-u |
| PWA Vozač | `Standard offer` / `Controlled rollout` prema aktivnom release scope-u |
| Kiosk standardizacija | `Controlled rollout` |
| Tablet hardverski paket | `Optional hardware` |
| Termalna štampa | `Controlled rollout` dok modeli i recovery nisu standardizovani |

---

## 7. Enterprise opcione ekstenzije

### 7.1 Finance & Regulatory Integration

- SEF;
- bankarski izvod i BankaImport;
- partner mapping;
- salda, avansi i status naplate;
- rasknjižavanje;
- nalozi za plaćanje;
- ERP import/export adapteri.

Status: `Optional extension`.

### 7.2 Logistics & Fleet

- Dispečer;
- raspodela vozila i vozača;
- rute, statusi i kapaciteti;
- neraspoređena roba;
- pregled realizacije.

Status: deo osnovnog toka gde je potreban; napredne funkcije su `Optional extension`.

### 7.3 Inputs, Agrohemija & Cooperant Balance

- izdavanje i razduženje repromaterijala;
- kartice i saldo;
- korpa i otpremnica;
- artikli, doziranje i pomoćne kontrole;
- veza sa Gazdinstvom.

Status: `Optional extension` ili viši paket.

### 7.4 Advanced Warehouse, Pallets & Traceability

- palete i skladišne jedinice;
- prerada, lager i izlaz;
- sledljivost kooperant/parcela–prijem–kupac;
- napredni kontrolni izveštaji.

Status: discovery-scoped `Optional extension`; ne predstavljati kao generički WMS/MES.

---

## 8. AgriX Gazdinstvo

Gazdinstvo obuhvata:

- karticu prema hladnjači;
- parcele, kulture i GIS;
- tretmane i karencu;
- opremu i rad;
- troškove i sezonski bilans;
- lager agrohemije;
- proizvodnju i dokumente;
- prognozu i upozorenja;
- offline/sync.

Radni paketi: Partner, Basic i Pro.

| Dimenzija | Status |
|---|---|
| Implementacija | `Implemented/Partial` |
| Dokaz | `Pilot evidence` |
| Komercijalno | `Pilot only` / kontrolisana rana ponuda |

Pre skaliranja: activation, 30/90/180-day retention, WTP i support cost.

---

## 9. AgriX GGAP

GGAP koristi Enterprise i Gazdinstvo podatke za:

- evidencije i liste;
- dokumente i dokaze;
- rokove i odgovorne osobe;
- neusaglašenosti i korektivne mere;
- audit readiness i export.

| Dimenzija | Status |
|---|---|
| Implementacija | `Planned/Discovery` |
| Dokaz | `Unvalidated` |
| Komercijalno | `Not for sale`, osim kontrolisanog discovery-ja/pilota |

AgriX ne garantuje sertifikaciju i ne zamenjuje auditora ili konsultanta.

---

## 10. Usluge

### Standardne

- discovery i procesno mapiranje;
- konfiguracija firme, stanica, korisnika i uloga;
- migracija master podataka;
- setup PWA, GAS, Sheets i desktop sloja;
- obuka otkupljivača, vozača, operatera i managementa;
- go-live, monitoring i readiness review;
- support handoff.

### Posebne

- Infosys i druge složene migracije;
- novi ERP/banka adapter;
- posebni reusable izveštaji;
- nova vaga/printer integracija;
- hardverska instalacija i remote management.

Trajni klijentski fork nije dozvoljen.

---

## 11. Trenutno dozvoljena ponuda

### Standardno ili kontrolisano prodavati

**AgriX Enterprise — teren do centrale**, uključujući:

- PWA Otkupac;
- PWA Vozač gde klijent koristi transportni tok;
- sync i centralno punjenje baze;
- centralni desktop/backoffice;
- dokumentni lanac;
- Management PWA;
- monitoring, update, backup i podršku;
- onboarding po ulozi.

Tačan status `Standard offer` ili `Controlled rollout` određuje se prema verziji, klijentskom procesu i sačuvanom release evidence-u.

### Opciono

- Finance & Regulatory;
- napredni Logistics & Fleet;
- Agrohemija i kooperantska zaduženja;
- Advanced Warehouse/Traceability;
- složene migracije i adapteri;
- hardver.

### Pilot only

- novi ili nepotvrđeni printer/hardware modeli;
- nove integracije bez produkcionog dokaza;
- Gazdinstvo u skaliranom komercijalnom modelu;
- GGAP prototip/pilot.

---

## 12. Readiness gate

Celina prelazi u `Standard offer` kada ima:

- jasan end-to-end tok;
- poznate podržane varijante;
- compile/smoke/regression dokaz;
- realni produkcioni dokaz relevantnog scope-a;
- monitoring i recovery;
- dokumentovan onboarding;
- support boundary i trošak;
- standardnu konfiguraciju bez forka;
- cenu i ugovornu granicu.

Za PWA terenski tok posebno meriti:

- sync success;
- duplicate i data-loss rate;
- broj ručnih centralnih korekcija;
- vreme od unosa do centrale;
- broj dokumenata bez ponovnog unosa;
- support minute po stanici.

---

## 13. Portfolio odluke

1. Enterprise Core je teren–centrala sistem.
2. PWA Otkupac i PWA Vozač su centralne komponente, ne sporedni add-on.
3. Desktop je centralni backoffice i canonical sloj nakon sinhronizacije.
4. Operater se fokusira na kontrolu, prijem, fakture, finansije i izveštaje.
5. Kiosk, tablet i termalna štampa imaju odvojene readiness statuse.
6. Pricing mora vrednovati broj stanica, terenskih korisnika i obim dokumenata.
7. Svaki demo mora prikazati teren → sync → centrala → faktura/izveštaj.
8. Gazdinstvo i GGAP ostaju zasebni proizvodi sa sopstvenim dokaznim pragovima.
