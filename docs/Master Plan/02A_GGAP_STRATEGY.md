# 02A — AgriX GGAP strategija

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 2026–2030  
**Poslednje ažuriranje:** 2026-07-27  
**Nivo:** strategija **modula**, ne stuba (odluka 402)

---

## 1. Strateška uloga

`DECISION` (odluka 402): GGAP je **modul u okviru AgriX Enterprise-a**, ne treći proizvodni stub. Kupci su isključivo hladnjače koje su već Enterprise klijenti; aktivacija modula otključava dodatne GGAP funkcije u Gazdinstvu njihovih kooperanata (PRT4). Treći stub je AgriX Savetnik (odluka 401).

Ceo ovaj dokument zato važi kao strategija modula: funkcionalni obuhvat, izvori podataka, granice odgovornosti i KPI-jevi ostaju na snazi, a delovi koji podrazumevaju zaseban proizvod — sopstveni ICP, packaging i unit economics — otpadaju i vode se u Enterprise ekonomici.

Pozicija modula u platformi:

1. `AgriX Enterprise` — operativni sistem firme, nosilac GGAP modula;
2. `AgriX Gazdinstvo` — farm-management sistem kooperanta, u kome se GGAP funkcije otključavaju;
3. `AgriX Savetnik` — upravljački sloj nad većim brojem gazdinstava.

GGAP je dokumentaciona kruna sistema zato što koristi podatke koje Enterprise i Gazdinstvo već stvaraju tokom realnog rada.

Osnovni princip:

> Podatak se unosi jednom na mestu nastanka, a GGAP evidencija i dokaz se iz njega automatski izvode.

GGAP ne treba da bude paralelni sistem u kojem se ručno prepisuju tretmani, parcele, artikli, karenca, proizvodnja, otkup i sledljivost.

---

## 2. Izvori podataka

AgriX GGAP treba da koristi:

- kooperante, parcele, kulture i površine;
- GIS poligone i lokacije;
- tretmane, mere, artikle i doze;
- opremu, vreme rada, geolokaciju i meteo snapshot;
- karencu i upozorenja;
- izdati repromaterijal i otpremnice;
- proizvodnju i otkup po parceli;
- ambalažu, prijem, preradu, lager i sledljivost;
- otpremu, kupce i povezane dokumente;
- potpise, korisnike, datume i audit događaje;
- račune, analize, sertifikate, ugovore i druge priloge.

Svaki automatski izvedeni podatak mora zadržati vezu sa izvornim poslovnim događajem.

---

## 3. Planirani funkcionalni domeni

### 3.1 Registar zahteva i listi

- katalog obaveznih listi, procedura, evidencija i dokaza;
- primenljivost po firmi, grupi, kooperantu, parceli, kulturi i sezoni;
- statusi: nije započeto, u toku, kompletno, neprimenljivo i blokirano;
- rok, prioritet i odgovorna osoba;
- verzija standarda i datum važenja.

### 3.2 Automatsko popunjavanje

- preuzimanje poznatih podataka iz Enterprise i Gazdinstvo sistema;
- jasno označavanje automatski generisanih i ručno dopunjenih polja;
- validacija kompletnosti i logičke usklađenosti;
- ponovno generisanje kada se izvorni podatak promeni, uz očuvanje istorije.

### 3.3 Dokumentacioni workflow

- generisanje i popunjavanje listi;
- obavezna polja i validaciona pravila;
- prilozi: fotografije, skenovi, računi, analize, sertifikati, ugovori i izjave;
- digitalni potpisi i odobravanja;
- verzije dokumenta i istorija promena;
- zaključavanje odobrenih dokumenata;
- obnova dokumenata po isteku ili promeni relevantnog podatka.

### 3.4 Zadaci i upozorenja

- zadaci po osobi, firmi, kooperantu i parceli;
- rokovi i podsetnici;
- nedostajući dokumenti;
- dokumenti pred istekom;
- neusaglašeni podaci između listi i operativnog sistema;
- upozorenja vezana za tretmane, karencu i dokazni tok;
- eskalacija managementu.

### 3.5 Neusaglašenosti i korektivne mere

- evidentiranje neusaglašenosti;
- nivo rizika i vlasnik problema;
- korektivna mera;
- rok i dokaz izvršenja;
- verifikacija i zatvaranje;
- audit trag;
- pregled otvorenih nalaza po gazdinstvu, parceli i organizaciji.

### 3.6 Interna kontrola i audit readiness

- kontrolna tabla spremnosti;
- procenat kompletiranosti po oblasti;
- pregled dokaza koji nedostaju;
- interna kontrolna lista;
- audit paket za firmu, grupu, kooperanta, parcelu ili period;
- kontrolisan read-only pristup ovlašćenom eksternom licu;
- export kompletnog dokumentacionog paketa.

### 3.7 Sledljivost dokaza

Za relevantnu stavku treba prikazati:

- izvorni poslovni događaj;
- ko je uneo podatak;
- vreme i uređaj;
- kooperanta, parcelu, artikal ili dokument;
- istoriju korekcija;
- ko je odobrio dokument;
- dokaz ili prilog koji potvrđuje tvrdnju.

---

## 4. Nivoi proizvoda

AgriX GGAP treba da podrži tri nivoa korišćenja:

### Gazdinstvo

- sopstvene liste i dokumentacija proizvođača;
- zadaci, rokovi i upozorenja;
- pregled spremnosti sopstvenog gazdinstva.

### Enterprise

- pregled i upravljanje dokumentacijom svih kooperanata;
- centralni zadaci i kontrole;
- standardizacija procesa i dokaza.

### Grupa / GGAP management

- upravljanje velikim brojem gazdinstava;
- grupne kontrole i neusaglašenosti;
- centralni audit readiness;
- koordinacija agronoma, quality osoba i odgovornih lica.

---

## 5. Produktni princip

AgriX GGAP nije statična arhiva PDF obrazaca.

Vrednost proizvoda mora dolaziti iz:

- automatskog popunjavanja iz realnih podataka;
- sprečavanja duplog unosa;
- upozorenja pre nego što propust postane problem;
- centralne kontrole velikog broja kooperanata;
- povezivanja dokaza sa izvornim događajem;
- ubrzane pripreme dokumentacije;
- manjeg rizika od nedostajućih ili zastarelih dokaza.

---

## 6. Granice odgovornosti

AgriX GGAP:

- vodi workflow;
- organizuje podatke i dokaze;
- proverava kompletnost i definisana pravila;
- upozorava na propuste;
- priprema dokumentaciju i audit paket.

AgriX GGAP ne sme da se predstavlja kao:

- sertifikaciono telo;
- nezavisni auditor;
- garancija dobijanja sertifikata;
- zamena za stručnog GGAP konsultanta;
- garancija agronomske ispravnosti svake odluke korisnika.

Standard, verzija, lokalni zahtevi i pravila moraju biti stručno validirani i verzionisani.

---

## 7. Ekonomika

Od odluke 402 GGAP nema sopstvenu ekonomiku — cena i marža se vode kao stavka Enterprise ponude. Cenovna referenca za pilot je „od 1.000 €“ uz potvrdu obima (odluka 352).

Modeli koji se i dalje mogu testirati u okviru Enterprise cenovnika:

- godišnja naknada za modul po firmi;
- osnovna cena modula plus broj aktivnih kooperanata ili sertifikovanih gazdinstava;
- onboarding i migracija postojeće dokumentacije;
- napredni workflow, audit priprema i integracije kao posebno procenjen rad.

Modeli koji otpadaju: zaseban paket za pojedinačno Gazdinstvo i licenca za grupu proizvođača bez Enterprise ugovora — GGAP se ne prodaje van postojeće Enterprise baze.

GGAP se ne sme ceniti kao generator obrazaca. Cena treba da odražava:

- uštedu administrativnog rada;
- manji rizik propusta;
- centralnu kontrolu;
- automatsko izvođenje dokaza;
- vreme pripreme za internu i eksternu proveru.

Trošak održavanja sadržaja, verzija standarda, validacija i stručnih provera mora biti deo unit economics-a.

---

## 8. Ključni rizici

| Rizik | Uticaj | Zaštita |
|---|---|---|
| Pravilo ne prati važeću verziju standarda | Kritičan | verzionisanje, domain owner, stručna revizija |
| Klijent očekuje garantovanu sertifikaciju | Kritičan | jasne ugovorne i UX granice |
| Automatski dokaz koristi pogrešan izvor | Kritičan | provenance, validacija i approval workflow |
| Širina scope-a preoptereti mali tim | Visok | fazni razvoj i ograničen pilot |
| Održavanje sadržaja je skuplje od prihoda | Visok | pilot unit economics |
| Korisnik vodi paralelne ručne evidencije | Visok | end-to-end onboarding i eliminacija duplog unosa |

---

## 9. KPI-jevi

- broj aktivnih firmi, grupa i gazdinstava;
- procenat automatski popunjenih polja i listi;
- broj eliminisanih ručnih unosa;
- procenat kompletiranosti dokumentacije;
- broj nedostajućih i isteklih dokaza;
- vreme pripreme dokumentacije po gazdinstvu;
- broj otvorenih i zatvorenih neusaglašenosti;
- vreme zatvaranja korektivnih mera;
- broj audit paketa i exporta;
- broj grešaka pronađenih pre kontrole;
- ARR, gross margin, renewal i support cost.

---

## 10. Pre implementacije

Pre razvoja produkcionog proizvoda potrebno je uraditi poseban discovery:

1. odabrati standard, verziju i tip sertifikacije;
2. prikupiti kompletan set listi, procedura i dokaza;
3. mapirati svako polje na postojeći AgriX podatak ili ručni unos;
4. definisati uloge, odobravanja i potpise;
5. definisati neusaglašenosti i korektivne mere;
6. odrediti obavezne priloge i rokove čuvanja;
7. definisati audit paket i export;
8. pronaći stručnog domain owner-a;
9. odabrati pilot firmu i ograničeni početni scope;
10. napraviti pricing i unit economics pre punog razvoja.

---

## 11. Odluka

### STR-012 — GGAP kao treći proizvodni stub — `Superseded` 2026-07-27

Predlog je bio: AgriX GGAP je treći puni proizvod, pored Enterprise i Gazdinstvo.

**Zamenjen odlukama 401 i 402** (vidi STR-013 u `DECISION_LOG.md`): treći stub je AgriX Savetnik, a GGAP je modul Enterprise-a.

### Važeći status modula

- **Obuhvat:** pokriti GGAP liste i kompletan dokumentacioni tok koristeći podatke iz operativnog i farm-management sistema bez ponovnog ručnog unosa — ostaje nepromenjen.
- **Kupac:** isključivo postojeći Enterprise klijent; GGAP se ne prodaje samostalno.
- **Komercijalno:** van komercijalne ponude do validacije (odluka 405); samo kontrolisan pilot uz potvrdu obima. Referentna cena „od 1.000 €“ (odluka 352) važi za pilot.
- **Ekonomika:** nema zasebnog packaging-a ni unit economics-a; vodi se u Enterprise ekonomici.
- **Granica odgovornosti:** nepromenjena — AgriX ne garantuje sertifikaciju i ne zamenjuje sertifikaciono telo, auditora ni stručnog konsultanta.
