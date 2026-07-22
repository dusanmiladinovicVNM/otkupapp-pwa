# 02 — Competition

**Status:** Initial evidence set  
**Poslednje ažuriranje:** 2026-07-23  
**Izvor:** reference liste koje je osnivač AgriX-a dostavio iz javno prikazanih materijala konkurenata. Stavke u ovom commit-u nisu nezavisno proverene sa samim korisnicima.

Konkurenti, alternative, interne Excel varijante, ERP sistemi, specijalizovani proizvodi, cene, funkcije, screenshotovi i win/loss dokazi pripadaju ovom direktorijumu.

Podaci o konkurenciji moraju imati datum, izvor i jasno odvojene činjenice od procena.

## 1. Sažetak reference evidence-a

| Konkurent | Ukupno navedenih projekata/referenci | Direktno relevantni otkupni sistemi | Susedni projekti |
|---|---:|---:|---:|
| SOFTEK | 13 | 13 | 0 |
| KRUNET | 15 | 11 | 3 relevantna susedna + 1 ostali |
| **Ukupno** | **28** | **24** | **4** |

`FACT — user-provided evidence`: dostavljene liste sadrže najmanje 24 reference direktno povezane sa softverom za otkup poljoprivrednih proizvoda ili otkup, preradu i skladištenje voća.

`INFERENCE`: specijalizovani softver za otkup u Srbiji nije hipotetička kategorija. Najmanje dva dobavljača prikazuju višegodišnje reference u ciljnom ili neposredno susednom segmentu.

`LIMITATION`: broj referenci nije isto što i broj trenutno aktivnih instalacija, aktivnih licenci, zadovoljnih korisnika ili tržišni udeo. Moguće su stare, ugašene, preimenovane ili migrirane firme.

## 2. SOFTEK

Dostavljena lista korisnika softvera za otkup poljoprivrednih proizvoda:

1. Agropromet — Užice;
2. Valerije — Vranje;
3. AgroFrigo Lukić — Arilje;
4. AS Group — Brus;
5. TR Džuver — Kosjerić;
6. Fortis doo — Bojnik;
7. Tomović — Koštunići;
8. Rosa — Aleksandrovac;
9. Eko-Nikolić — Kosjerić;
10. Eurohorizont doo — Prijepolje;
11. Paun — Teočin;
12. Županjka — Rekovac;
13. STUR Gorava — Priboj.

### Signali

- 13 direktno relevantnih referenci;
- široka geografska raspodela;
- vidljiv klaster Zapadne i Centralne Srbije;
- najmanje dve reference u Kosjeriću;
- reference uključuju privredna društva, trgovinske radnje i manje lokalne subjekte.

`INFERENCE`: SOFTEK verovatno pokriva širok raspon veličina kupaca, uključujući manji i srednji segment, ali bez podataka o funkcijama, cenama i aktuelnom statusu instalacija to nije potvrđeno.

## 3. KRUNET

Dostavljena lista projekata Agencije za poslovni konsalting KRUNET:

### Direktno relevantni sistemi

#### Hladis

1. Strela d.o.o.;
2. Vule komerc d.o.o.;
3. Maks d.o.o.;
4. Voće produkt d.o.o.

#### K.I.A. fruit

5. Sweet home;
6. MENEX;
7. Janeks d.o.o.;
8. Žileks trgovina;
9. MDDP Janić;
10. Eko Food — Vlasotince;
11. Jeka Fruit.

Za svih 11 navedena je namena: softver za otkup, preradu i skladištenje voća.

### Susedni projekti

- PET-ING — softver za magacinsko zaduženje;
- AMIGOS Kafa — praćenje proizvodnje, sledljivosti i uporedna analiza cena;
- Dunipak — praćenje proizvodnje kroz sve faze u industriji kartonske ambalaže;
- Marić Centar d.o.o. — softver za marketinški nastup na tržištu, koji nije direktno relevantan za AgriX konkurentsku analizu.

### Signali

- KRUNET prikazuje najmanje dve imenovane produktne porodice za voćarski sektor: `Hladis` i `K.I.A. fruit`;
- fokus nije samo na otkupnom listu, već eksplicitno uključuje preradu i skladištenje;
- susedni projekti ukazuju na iskustvo u proizvodnji, magacinu i sledljivosti;
- model može biti kombinacija proizvoda i namenski prilagođenih implementacija.

`INFERENCE`: KRUNET je funkcionalno bliži AgriX Enterprise viziji od dobavljača koji nudi samo dokumente za otkup. Njegov stvarni današnji obim proizvoda i održavanja mora se nezavisno proveriti.

## 4. Tržišne implikacije

1. Postojanje 24 direktne reference potvrđuje da kupci u Srbiji već plaćaju ili su istorijski plaćali specijalizovani softver za otkup.
2. Konkurencija ima reference u manjim gradovima i proizvodnim klasterima; tržište nije koncentrisano samo u Beogradu ili velikim centrima.
3. Funkcije `otkup + prerada + skladištenje` predstavljaju tržišno prepoznatu kategoriju, ne samo internu AgriX viziju.
4. Reference konkurenata mogu biti početni seed za win/loss i replacement-market istraživanje, ali ne smeju se kontaktirati agresivno niti se pretpostaviti da su nezadovoljne.
5. AgriX mora dokazati razliku kroz povezanost terena, managementa, logistike, finansija, Gazdinstva i GGAP-a, a ne samo kroz postojanje otkupnih dokumenata.

## 5. Sledeća validacija

Za svaku referencu treba dopuniti:

- tačan pravni naziv i matični broj;
- status firme;
- prihod i broj zaposlenih;
- kulture i proizvode;
- broj lokacija/stanica;
- da li se konkurentski sistem još koristi;
- koje module koristi;
- približnu godinu implementacije;
- javno dostupne dokaze: sajt, objava, screenshot, oglas za posao ili dokument;
- potencijalni switching trigger bez pretpostavljanja nezadovoljstva.

Strukturisani redovi nalaze se u `competitor_references.csv`.
