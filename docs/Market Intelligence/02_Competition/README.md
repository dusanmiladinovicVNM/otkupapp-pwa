# 02 — Competition

**Status:** Initial evidence set  
**Poslednje ažuriranje:** 2026-08-02  
**Izvori:** reference liste koje je osnivač AgriX-a dostavio iz javno prikazanih materijala konkurenata, founder-confirmed AgriX win/loss događaji i zvanična korisnička dokumentacija konkurenta (AGROSOFT). Javne reference nisu nezavisno proverene sa navedenim korisnicima.

Konkurenti, alternative, interne Excel varijante, ERP sistemi, specijalizovani proizvodi, cene, funkcije, screenshotovi i win/loss dokazi pripadaju ovom direktorijumu.

Podaci o konkurenciji moraju imati datum, izvor i jasno odvojene činjenice od procena.

## 1. Vrste dokaza

- `PUBLIC REFERENCE EVIDENCE` — firma ili projekat naveden u dostavljenoj konkurentskoj referentnoj listi;
- `PRODUCT DOCUMENTATION EVIDENCE` — funkcija potvrđena u zvaničnoj dokumentaciji samog konkurenta (uputstvo, priručnik, specifikacija);
- `CLASSIFIED REFERENCE EVIDENCE` — stavka izdvojena iz šire liste na osnovu naziva, lokacije i sektorskog signala;
- `FOUNDER-CONFIRMED BUSINESS EVENT` — stvarni AgriX prodajni ili migracioni događaj koji je potvrdio osnivač;
- `INFERENCE` — zaključak izveden iz više dokaza;
- `LIMITATION` — razlog zbog kojeg se dokaz ne sme tumačiti šire nego što dozvoljava.

`LIMITATION`: broj referenci nije isto što i broj trenutno aktivnih instalacija, licenci, zadovoljnih korisnika ili tržišni udeo. Moguće su stare, ugašene, preimenovane ili migrirane firme.

## 2. Evidence dataset-i

| Fajl | Sadržaj |
|---|---|
| `competitor_references.csv` | 40 javno navedenih projekata/reference SOFTEK-a, KRUNET-a i Yuteam-a |
| `infosys_agro_references.csv` | 113 agro/prehrambenih stavki izdvojenih iz šire Infosys referentne liste |
| `competitive_events.csv` | dva founder-confirmed prelaska sa Infosys-a na AgriX |
| `infosys_agro_references_summary.md` | agregati, kategorije, geografija i prioriteti za Infosys replacement pool |
| `AgroSoft-Korisnicko-Uputsvo.pdf` | korisničko uputstvo za AGROSOFT („DATA SOFT" Vrbas), 161 strana — jedina konkurentska **produkt-dokumentacija** u repou; sadržaj datiran ~2012–2013 (Windows 7/XP, sezone „rod 2011/2012") |
| `agrosoft_feature_teardown.md` | feature-level poređenje AGROSOFT ↔ AgriX po deset oblasti, sa dokazima iz `src-vba/`, `src/` i `gas/` |
| `SOFTEK_uputstvp_otkup_poljoproizvoda.pdf` | korisničko uputstvo za SOFTEK modul „Otkup poljoprivrednih proizvoda", 34 strane; PDF kreiran 2017-06-06 |
| `Softek-otkup.pdf` | **uputstvo na koje sam program linkuje** (Pomoć), „Verzija 2.1", 16 strana; PDF kreiran 2014-05-28. Tanje od verzije iz 2017 i ne pominje funkcije viđene u aplikaciji v20.2.4 — in-app pomoć kasni za proizvodom |
| `softek_feature_teardown.md` | feature-level poređenje SOFTEK ↔ AgriX; prvi dokumentovani **direktan** konkurent (malina, gajbice, PDV nadoknada 8%). §8 opisuje **živu aplikaciju `Ver.20.2.4`** iz ~20 screenshot-ova: kontna arhitektura, IOS i otvorene stavke, KEP, Access/Jet backend, šest uočenih defekata |

## 3. Sažetak javnih i klasifikovanih referenci

### Direktne i susedne liste SOFTEK, KRUNET i Yuteam

| Konkurent | Ukupno navedenih projekata/referenci | Direktni otkup/voće | Poljoprivredna kooperacija i susedni projekti | Ostalo |
|---|---:|---:|---:|---:|
| SOFTEK | 13 | 13 | 0 | 0 |
| KRUNET | 15 | 11 | 3 | 1 |
| YUTEAM | 12 | 0 potvrđenih u užem voćarskom scope-u | 12 | 0 |
| **Ukupno** | **40** | **24** | **15** | **1** |

### Infosys agro/prehrambeni reference universe

Iz šire Infosys liste izdvojeno je:

| Pokazatelj | Broj |
|---|---:|
| Agro/prehrambene stavke | 113 |
| Jasno iz naziva | 89 |
| Potrebna dodatna provera | 24 |
| Visok AgriX potencijal | 48 |
| Srednji potencijal | 49 |
| Nizak potencijal | 16 |
| Navedene lokacije | 47 |

Visokopotencijalnih 48 čine:

- 25 voće, povrće i hladnjače;
- 13 poljoprivreda i kooperative;
- 9 mlekara;
- 1 duvanski račun.

`INFERENCE`: Infosys ima znatno širu agro/prehrambenu instalacionu ili istorijsku referentnu bazu od ranije dokumentovane konkurencije. Ipak, dataset ne potvrđuje da svih 113 firmi danas koristi isti Infosys proizvod ili agro modul.

## 4. INFOSYS — prioritet 1

`FOUNDER-CONFIRMED BUSINESS EVENT`: dva postojeća AgriX klijenta prethodno su koristila Infosys sistem.

Dodatni kontekst:

- oba računa imaju približno 150 miliona RSD godišnjih prihoda;
- prelazak predstavlja potvrđene AgriX pobede protiv incumbent sistema;
- osnivač Infosys procenjuje kao jednog od najvećih konkurenata u segmentu;
- šira referentna lista sadrži najmanje 113 agro/prehrambenih stavki, od kojih je 48 označeno kao visok AgriX potencijal.

`INFERENCE`: Infosys mora imati najviši prioritet u product teardown-u, win/loss istraživanju, migracionom playbook-u i replacement-market prodaji.

Za oba migrirana računa treba dokumentovati:

- šta ih je navelo da traže zamenu;
- koje procese prethodni sistem nije dovoljno rešavao;
- koji je switching cost postojao;
- zašto je AgriX izabran;
- šta im je bilo najteže u migraciji;
- u čemu je Infosys bio bolji;
- šta bi ih moglo vratiti prethodnom dobavljaču;
- koje AgriX prednosti su stvarno korišćene.

## 5. SOFTEK

Dostavljena lista sadrži 13 korisnika softvera za otkup poljoprivrednih proizvoda, među njima firme iz Užica, Vranja, Arilja, Brusa, Kosjerića, Bojnika, Aleksandrovca, Prijepolja, Rekovca i Priboja.

Signali:

- 13 direktno relevantnih referenci;
- široka geografska raspodela;
- vidljiv klaster Zapadne i Centralne Srbije;
- lista uključuje privredna društva i trgovinske radnje.

`PRODUCT DOCUMENTATION EVIDENCE`: uputstvo za modul „Otkup poljoprivrednih proizvoda" (34 strane, PDF iz 2017) je u ovom folderu. Radni primer kroz ceo dokument je `MALINA VILAMET I KLASA` sa ambalažom `GAJBICA MALINE`, neto se računa kao bruto minus gajbice, ambalaža ide kroz revers sa `ZADUŽENO/RAZDUŽENO/STANJE`, a uvod objašnjava PDV nadoknadu od 8%.

`PRODUCT DOCUMENTATION EVIDENCE` (screenshot): oko 20 snimaka žive aplikacije `Ver.20.2.4` (poslovna godina 2021, demo firma podešena na Arilje). Oni pokazuju da otkup nije samostalan proizvod nego **vrsta dokumenta `381` u glavnoj knjizi** — dobavljač je analitika konta `4358`, magacin je konto `1311`, otkupno mesto je „mesto troška", a svaki dokument se zatvara `Knjiženjem`. Backend je Microsoft Access (Jet), ne server baza.

`INFERENCE`: SOFTEK je **prvi dokumentovani direktan konkurent** — isti proizvod, isti kupac, ista regulativa i isti geografski klaster kao AgriX. Njegov ugao je knjigovodstvo (KEP, IOS, otvorene stavke, bruto bilans); AgriX-ov je teren i lanac posle otkupa. Prava linija razdvajanja je **ko sme da radi u programu**, ne spisak funkcija. Cene i aktuelni status instalacija i dalje nisu potvrđeni.

Puna analiza: `softek_feature_teardown.md`.

## 6. KRUNET

KRUNET prikazuje 11 direktno relevantnih projekata kroz porodice `Hladis` i `K.I.A. fruit`, sa namenom otkup, prerada i skladištenje voća.

Susedni projekti uključuju magacinsko zaduženje, proizvodnju i sledljivost.

`INFERENCE`: KRUNET je funkcionalno bliži AgriX Enterprise kategoriji od dobavljača koji nudi samo otkupne dokumente. Stvarni današnji obim proizvoda i održavanja mora se proveriti.

## 7. YUTEAM

Dostavljena lista `RUGE I KOOPERACIJE` sadrži 12 referenci u agraru, zemljoradničkim zadrugama, kooperaciji i mlinarstvu, sa jakim klasterom Vojvodine.

`INFERENCE`: Yuteam je posebno relevantan za buduće AgriX širenje prema ratarstvu, žitaricama, zadrugama i organizovanoj kooperaciji.

## 8. AGROSOFT — „DATA SOFT" Vrbas

`PRODUCT DOCUMENTATION EVIDENCE`: kompletno korisničko uputstvo (161 strana) je u ovom folderu. To je jedini konkurent za koga postoji funkcionalni dokaz, a ne samo referentna lista.

Ciljni segment po uputstvu: **zemljoradničke zadruge, agrokombinati i skladištenje žitarica i industrijskog bilja**. Jezgro proizvoda su kolska vaga preko serijskog porta, laboratorijski obračun kvaliteta sa formulama i intervalima, silosne ćelije, skladišna usluga (potvrda/ugovor o skladištenju, prenos vlasništva, kompenzacija) i ugovaranje proizvodnje sa naturalnim i finansijskim paritetima.

`INFERENCE`: AGROSOFT i AgriX se **ne takmiče direktno** — različit segment (žito/silos vs voće/hladnjača). Preklapanje je samo u zajedničkom jezgru (partneri, prijem, dokument, kartica, isplata, prava korisnika).

`LIMITATION`: sadržaj uputstva je datiran ~2012–2013; aktuelna verzija proizvoda nije proverena. DATA SOFT nema nijedan red u `competitor_references.csv` ni u `infosys_agro_references.csv` — broj instalacija je nepoznat.

Puna analiza: `agrosoft_feature_teardown.md`.

## 9. Tržišne implikacije

1. Tržište nije greenfield.
2. Infosys ima najjači trenutni replacement dokaz i najširi evidentirani agro/prehrambeni reference universe.
3. Voćarski konkurentski klaster vidljiv je kod Infosys-a, SOFTEK-a i KRUNET-a.
4. Yuteam potvrđuje poseban Vojvodina–ratarstvo–zadruge segment.
5. Infosys lista dodatno potvrđuje da APR scan `1039/4631` propušta mlekare, duvan, kooperative i druge relevantne procese.
6. Prethodni sistem mora biti obavezno CRM polje.
7. Konkurentske reference mogu biti seed za replacement-market istraživanje, ali ne smeju se kontaktirati agresivno niti se pretpostaviti da su nezadovoljne.
8. AgriX mora dokazati razliku kroz povezanost terena, managementa, logistike, finansija, Gazdinstva i GGAP-a, a ne samo kroz otkupni dokument.

## 10. Sledeća validacija

Za svaku javnu ili klasifikovanu referencu treba dopuniti:

- tačan pravni naziv i matični broj;
- aktivan status;
- prihod i broj zaposlenih;
- šifru delatnosti;
- kulture i proizvode;
- broj lokacija/stanica;
- da li se konkurentski sistem još koristi;
- koje module koristi;
- približnu godinu implementacije;
- javno dostupne dokaze;
- potencijalni switching trigger;
- ownera sledeće dozvoljene akcije.

Za Infosys je prvi prioritet:

1. dva strukturisana win interview-a;
2. pravna i APR validacija 48 visokopotencijalnih referenci;
3. dokumentovan migracioni format i playbook;
4. evidence-based battlecard.
