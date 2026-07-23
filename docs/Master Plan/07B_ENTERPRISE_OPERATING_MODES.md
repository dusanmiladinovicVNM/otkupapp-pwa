# 07B — AgriX Enterprise režimi rada

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `06_POSITIONING.md`, `07_PRODUCT_PORTFOLIO.md`, `08_PRODUCT_ROADMAP.md`, `16_ONBOARDING_AND_IMPLEMENTATION.md`

---

## 1. Svrha

Ovaj dokument razdvaja dve činjenice koje moraju istovremeno ostati tačne:

1. AgriX VBA Desktop omogućava da centralni operater unese kompletan poslovni tok na osnovu papirnih dokumenata.
2. Glavna diferencijacija AgriX-a nastaje kada se koristi PWA, jer se dokument i poslovni podatak stvaraju jednom na terenu, a zatim automatski sinhronizuju ka centralnoj bazi.

Desktop fallback nije arhitektonska greška. On je važan continuity, migration i recovery mehanizam. Međutim, desktop-only način rada ne ostvaruje punu vrednost distribuiranog AgriX modela.

---

## 2. Režim A — PWA-led teren–centrala

Ovo je ciljni i preporučeni operativni režim za firme sa otkupnim stanicama i vozačima.

### 2.1 Tok otkupa

1. Otkupac/vagač identifikuje kooperanta i unosi otkup na mestu nastanka.
2. PWA formira podatke osnovnog otkupnog dokumenta.
3. Dokument se štampa ili prosleđuje kooperantu/otkupljivaču prema podržanom toku.
4. Paralelno sa lokalnim čuvanjem i štampom, zapis ulazi u offline/sync queue.
5. GAS/Google Sheets transportni sloj prenosi zapis ka desktop master toku.
6. VBA MasterSync importuje zapis u centralne tabele, formira ili povezuje pripadajući dokumentni lanac i vraća sync status.
7. Centralni operater kontroliše izuzetke i nastavlja prijem, fakturisanje, SEF, banku, salda i izveštaje.

### 2.2 Tok vozača

1. Vozač preuzima robu sa jedne ili više stanica.
2. PWA Vozač kreira ili dopunjava zbirnu i transportne podatke.
3. Podatak se lokalno čuva i sinhronizuje prema centrali.
4. Centralni sistem povezuje transport sa otkupom, otpremnicama, prijemom i sledljivošću.

### 2.3 Glavna vrednost

> **Ista radnja na terenu istovremeno daje dokument učesniku procesa i podatak centrali.**

Zbog toga:

- nema ponovnog prekucavanja istog otkupa;
- centrala dobija podatke pre nego što stigne papir;
- operater se bavi kontrolom i završnom obradom;
- smanjuje se kašnjenje između stanice i centrale;
- smanjuje se mogućnost razlike između papirnog dokumenta i centralne evidencije;
- broj stanica može rasti bez proporcionalnog povećanja centralnog ručnog unosa.

---

## 3. Režim B — Desktop manual/fallback

VBA Desktop mora omogućiti kompletan unos kada:

- klijent privremeno ili trajno ne koristi PWA;
- dokument je nastao na papiru;
- stanica ili uređaj nisu bili dostupni;
- potrebno je uneti istorijsku dokumentaciju;
- radi se migracija sa prethodnog sistema;
- postoji incident i aktivirana je fallback procedura;
- klijent je ugovorio desktop-only konfiguraciju.

Centralni operater tada može ručno da unese:

- otkup;
- otpremnicu;
- zbirnu;
- prijemnicu;
- ambalažu i povezana zaduženja;
- fakturu i finansijski nastavak;
- korekciju i storno prema desktop pravilima.

Desktop-only tok ostaje potpuno podržan poslovni tok, ali podrazumeva veći operativni rad centrale i ne eliminiše prenos podataka sa papira.

---

## 4. Source-of-truth pravilo

PWA nije zamena za centralni poslovni model i VBA dokumentnu logiku.

- PWA je primarni owner terenskog unosa u PWA-led režimu.
- Google Sheets/GAS predstavljaju operativni transportni i shared-state sloj.
- VBA/Excel ostaje canonical centralni master za dokumenta, finansije i master podatke nakon kontrolisanog importa.
- Desktop ručni unos ostaje canonical kada je dokument unet kroz fallback režim.

I PWA-led i desktop fallback moraju završiti u istim centralnim poslovnim pravilima, tabelama, kontrolama, izveštajima i audit toku.

---

## 5. Offline i sync ugovor

Štampa i sync nisu ista operacija i ne smeju lažno potvrđivati jedna drugu.

- Terenski zapis se prvo pouzdano čuva lokalno.
- Štampa koristi sačuvani poslovni zapis i podržani print payload.
- Sync može biti neposredan ili odložen zbog mreže ili `MASTER_SYNC_LOCK` stanja.
- Neuspešan trenutni sync ne sme da izbriše lokalni zapis niti da spreči dokumentovan fallback.
- Korisnik mora da vidi status `pending`, `syncing`, `synced` ili `error`.
- Retry mora koristiti stabilan identifikator radi idempotency zaštite.
- Centrala mora razlikovati zapis koji čeka import, zapis koji je importovan i zapis koji zahteva intervenciju.

Prodajna poruka zato nije da štampa zavisi od interneta, već da se podatak čuva na mestu nastanka i sinhronizuje čim transportni uslovi dozvole.

---

## 6. Posledice za portfolio

### Enterprise Core

Enterprise Core mora podržavati oba režima:

- PWA-led teren–centrala kao preporučeni režim sa najvećom vrednošću;
- desktop manual/fallback kao continuity i alternativni režim.

PWA Otkupac i PWA Vozač nisu obavezni da bi VBA Desktop tehnički mogao da vodi posao. Oni su, međutim, ključni da bi kupac dobio glavnu AgriX prednost — automatsko punjenje centrale iz događaja nastalih na terenu.

### Hardver

Kiosk tablet, konkretan printer i printer bridge imaju odvojeni readiness status. Njihova nezrelost ne sme da se koristi kao dokaz da sama PWA aplikacija ili PWA-led poslovni model nisu funkcionalni.

---

## 7. Posledice za demo i prodaju

Demo treba prvo da pokaže PWA-led tok:

1. otkupac unosi otkup;
2. nastaje dokument za kooperanta/otkupljivača;
3. zapis odlazi u sync queue;
4. centrala preuzima podatak;
5. operater nastavlja prijem, fakturu i izveštaj bez ponovnog unosa.

Zatim treba pokazati desktop fallback:

- isti poslovni događaj može se potpuno uneti u centrali kada PWA nije korišćena;
- fallback nije glavni benefit, već zaštita kontinuiteta i fleksibilnosti klijenta.

Dozvoljena formulacija:

> **AgriX podržava kompletan desktop unos, ali najveću uštedu ostvaruje kada otkupci i vozači sami kreiraju dokumente na terenu: dokument se odmah izdaje učesniku, a isti podatak automatski puni centralnu bazu.**

---

## 8. KPI-jevi po režimu

### PWA-led

- procenat otkupa unetih na terenu;
- procenat dokumenata odštampanih iz istog PWA zapisa;
- vreme od terenskog unosa do centralnog importa;
- sync success rate;
- duplicate i data-loss rate;
- broj ručnih korekcija u centrali;
- broj eliminisanih ponovnih unosa;
- vreme centralnog operatera po 100 otkupa.

### Desktop fallback

- broj i procenat ručno unetih dokumenata;
- razlog korišćenja fallback-a;
- vreme ručnog unosa po dokumentu;
- stopa grešaka u prepisivanju;
- broj dana ili stanica koje su radile bez PWA;
- vreme povratka u PWA-led režim posle incidenta.

---

## 9. Odluke

1. AgriX podržava i PWA-led i desktop manual/fallback režim.
2. PWA-led režim je preporučeni model i glavna diferencijacija proizvoda.
3. Desktop fallback je važan za kontinuitet, migraciju, incidente i desktop-only klijente.
4. Kompletna desktop funkcionalnost ne umanjuje strateški značaj PWA aplikacije.
5. PWA nije zamena za VBA business layer; ona je terenski ulaz u isti centralni dokumentni i finansijski sistem.
6. Štampa i sync polaze iz istog sačuvanog poslovnog događaja, ali imaju odvojene success/failure statuse.
7. Operater ne mora da prepisuje dokument koji je uspešno nastao i sinhronizovan kroz PWA.
8. Kada PWA nije korišćena, operater može kompletno da unese papirni dokument kroz VBA Desktop.
9. Pricing treba da razlikuje desktop-only korišćenje od PWA-led vrednosti po stanici i terenskom toku.
10. Roadmap treba da meri smanjenje fallback korišćenja zbog tehničkih razloga, ali nikada ne sme ukloniti desktop fallback sposobnost.
