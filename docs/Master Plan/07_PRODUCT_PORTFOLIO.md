# 07 — Portfolio proizvoda AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-27  
**Povezani dokumenti:** `02_STRATEGY.md`, `03_CUSTOMERS_AND_JOBS.md`, `06_POSITIONING.md`, `07A_PRODUCT_STATUS_MATRIX.csv`, `07B_ENTERPRISE_OPERATING_MODES.md`, `08_PRODUCT_ROADMAP.md`, `09_QA_DECISION_LOG.md`, `10_PRICING_AND_PACKAGING.md`, `14_GO_TO_MARKET.md`, `15_SALES_PLAYBOOK.md`, `16_ONBOARDING_AND_IMPLEMENTATION.md`, `docs/Product/AgriX_Definicija_proizvoda.pdf`

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

AgriX ima tri proizvodna stuba (odluke 269 i 401):

1. **AgriX Enterprise** — povezani terenski i centralni operativni sistem firme;
2. **AgriX Gazdinstvo** — farm-management sistem kooperanta/proizvođača;
3. **AgriX Savetnik** — upravljački sloj nad većim brojem gazdinstava, za agronome i savetodavne službe.

**GGAP nije stub.** GGAP je modul u okviru Enterprise-a; kupuju ga isključivo hladnjače koje su već Enterprise klijenti, a aktivacija otključava dodatne GGAP funkcije u Gazdinstvu njihovih kooperanata (odluka 402, PRT4). Ranija odluka STR-012 time je ukinuta.

Moduli uz Enterprise: Hladnjača/Proizvodnja, SEF, Banka, Dispatch (samo uz Mobile) i GGAP. Obračunska jedinica nije ista za sve — SEF, Banka i Dispatch plaćaju se jednom po pravnom licu i važe kroz sve instance, dok se **Hladnjača/Proizvodnja plaća po proizvodnom pogonu** (odluke 412, 421).

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
| Komercijalno | `Standard offer` — odluka 404 |

`DECISION` (odluka 404): Gazdinstvo je launch ready i prelazi iz `Pilot only` u `Standard offer`.

Cene (odluka 339): maloprodajna **19 € Basic / 39 € Pro**; kanalska, za naloge posredovane preko hladnjače ili savetnika, **10 € Basic / 20 € Pro**. Prvih 50 Basic naloga partner dobija bez naknade (odluka 161). Proizvođač ima jedan Pro nalog — ko ga prvi aktivira, taj ga plaća (odluka 343). Pro se plaća direktno ili preko hladnjače (PRT2).

Gazdinstvo mora biti kompletan i vredan proizvod i bez ijedne povezane AgriX hladnjače (odluka 321); Enterprise povezivanje prvenstveno donosi korist hladnjači.

Ono što se odlukom 404 **ne** menja: activation, 30/90/180-day retention, WTP i support cost i dalje se mere. Oni sada služe za korekciju paketa, cene i support modela, a ne kao kapija pred prodaju. Gazdinstvo takođe ne može javno na tržište bez sopstvene politike privatnosti — zavisnost od LEG1 (`docs/Legal/AgriX_Mapa_tokova_podataka.pdf`).

---

## 9. AgriX Savetnik

Treći stub (odluke 269, 401, PRT3). Upravljački sloj nad većim brojem gazdinstava, za agronome i savetodavne službe:

- pregled portfelja gazdinstava;
- radni nalozi i preporuke;
- praćenje izvršenja i kontrola rada;
- agronomska istorija po gazdinstvu, parceli i kulturi.

| Dimenzija | Status |
|---|---|
| Implementacija | `Planned` — osnovna verzija do 2027 (odluka 203) |
| Dokaz | `Unvalidated` |
| Komercijalno | **`U pripremi`** — cena objavljena u `Cenovnik 2027` (odluke 341, 347), ali se ne ugovara dok proizvod ne bude stabilan (odluke 217, 423) |

`DECISION` (odluka 423): kad god se Savetnik prikaže u cenovniku, mora nositi **vidljivu oznaku „u pripremi“** i napomenu da se ne ugovara dok ne bude stabilan. Zlatni okvir ili druga vizuelna oznaka preporuke ne sme se koristiti na Savetniku dok traje taj status. Analogno odluci 417 za GGAP.

Komercijalni model je dvostruk: savetnik plaća alat, a gazdinstva u portfelju zadržavaju sopstvenu Pro pretplatu po kanalskoj ceni od 20 € (odluke 340, 339). Postoje **dve objavljene tarife** (odluka 419):

| Tarifa | Osnovica (do 10 aktivnih gazdinstava) | Svako preko 10 |
|---|---:|---:|
| Standalone — bez drugog ugovornog odnosa sa AgriX-om | 150 € | 15 € |
| Enterprise — uz aktivan Enterprise ugovor pravnog lica | 100 € | 10 € |

Enterprise tarifa traje dok traje Enterprise ugovor; prestankom se prelazi na standalone pri prvoj narednoj obnovi, bez retroaktivnog obračuna. Model naplate je osnovica plus fiksni iznos po aktivnom gazdinstvu preko deset (odluka 420).

Savetnik ne dobija proviziju za gazdinstva u portfelju; podsticaj je sam alat, koji bez Pro naloga ne funkcioniše (odluka 345). Provizija po odluci 221 ostaje samo za preporuke van portfelja.

Aktivno gazdinstvo je ono kojem je savetnik u toku godine poslao makar jedan nalog ili preporuku (odluka 342). Proba obuhvata i Pro za do 10 gazdinstava (odluka 346). Kada su gazdinstva u portfelju već pokrivena partnerskim paketom hladnjače, savetnik plaća samo alat po Enterprise tarifi i Pro pretplate se ne plaćaju ponovo (odluke 348, 419).

Zavisnosti pre izlaska na tržište: samostalna registracija i politika privatnosti, i pravna ocena toka T13 (LEG1).

`UNKNOWN`: packaging i cena su zaključani, ali product strategy trećeg stuba još nije napisana — funkcionalni obim ovog odeljka je okvir, ne specifikacija.

---

## 9A. GGAP — modul Enterprise-a

`DECISION` (odluka 402): GGAP nije stub nego **modul u okviru Enterprise-a**. Kupac je uvek postojeća hladnjača sa Enterprise ugovorom; aktivacija modula otključava dodatne GGAP funkcije u Gazdinstvu njenih kooperanata (PRT4).

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
| Komercijalno | `Not for sale` — van komercijalne ponude do validacije (odluka 405); samo kontrolisan discovery/pilot |

Posledice statusa modula: GGAP nema sopstveni ICP, packaging ni unit economics — vodi se unutar Enterprise ekonomike. Cena „od 1.000 € godišnje po pravnom licu“ (odluka 352) ostaje referentna za pilot uz potvrdu obima, ne za redovnu ponudu.

`DECISION` (odluka 417): kad god se GGAP prikaže u cenovniku, uz njega **mora** stajati vidljiva oznaka **„na upit, uz potvrdu obima — nije deo standardne ponude“**. Bez te oznake stavka se ne sme prikazati, jer bi je prodaja mogla kotirati kao redovnu.

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

### Satnice

Postoje tačno **dve** standardne satnice, i biraju se prema **prirodi posla, ne prema mestu izvođenja** (odluka 409):

| Satnica | Iznos | Obuhvata |
|---|---:|---|
| Razvojna | 50 €/h | razvoj po zahtevu, složena migracija, novi adapteri, posebni izveštaji, masovne korekcije podataka |
| Implementaciona | 30 €/h | obuka preko uključenih pet sati, konfiguracija, čišćenje podataka, IT setup, procesni konsalting, rad na lokaciji |

Izlazak na teren je 50 € po izlasku, uvećano za gorivo, vreme puta i vreme rada; **vreme puta se uvek obračunava po implementacionoj satnici**, a vreme rada po prirodi posla (odluka 410). Usluge iz C7 raspoređene su po satnicama odlukom 411.

Satnice su fiksne i nepregovaračke. Pregovaračkih i individualnih popusta nema (odluka 418) — jedina cenovna razlika unutar istog obima je −50 % na drugu i svaku narednu instancu (odluka 413).

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

**AgriX Gazdinstvo** — Basic i Pro, `Standard offer` od odluke 404; prodaje se preko hladnjače, savetnika ili direktno proizvođaču.

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
- GGAP modul — van komercijalne ponude do validacije (odluka 405).

### U pripremi — cena objavljena, isporuka tek predstoji

- AgriX Savetnik — cena je u `Cenovnik 2027` (odluke 341, 347), prikazuje se uz obaveznu oznaku „u pripremi“ (odluka 423) i ne ugovara se dok proizvod ne bude stabilan (odluka 217).

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
6. Pricing vrednuje broj stanica, terenskih korisnika i obim dokumenata, ali je **cena po stanici jedinstvena** bez obzira na režim rada; razliku pokriva cena Mobile paketa (odluka 406).
7. Svaki demo mora prikazati teren → sync → centrala → faktura/izveštaj.
8. Tri stuba su Enterprise, Gazdinstvo i Savetnik (odluka 401); GGAP je modul Enterprise-a (odluka 402).
9. Gazdinstvo je `Standard offer` (odluka 404); GGAP ostaje van komercijalne ponude do validacije (odluka 405).
10. Savetnik ima objavljenu cenu (odluke 341, 347), ali se ne ugovara dok proizvod ne bude stabilan (odluka 217).
