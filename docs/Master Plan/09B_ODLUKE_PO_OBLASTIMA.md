# 09B — AgriX, sve odluke grupisane po oblastima

**Datum:** 27.07.2026.  **Verzija:** 3
**Izvor:** Q&A Decision Log 1–321, blokovi A/M/S/ML/I/D/P/IP/BC/L/Q/ON/C/MKT/PRT/LEG, plus odluke 323–378 iz sesije 26.07, 401–408 i 409–422 (cenovna revizija 27.07).
**Namena:** indeks, ne zamena za 09a. Tekst je sažet na jednu liniju po odluci; merodavna formulacija ostaje u Decision Logu.
**Numeracija:** 323–378 nastavlja niz posle 321; 401–422 posle toga. Brojevi 322 i 379–400 se ne koriste.

> **Mesto u repou.** Merodavan Decision Log je `09_QA_DECISION_LOG.md`; ovaj fajl je tematski indeks nad njim. Odluke 323–378 unete su u `09_QA_DECISION_LOG.md` §25 na osnovu ovog indeksa i pratećih dokumenata (`docs/Sales/AgriX_Cenovnik_2027.pdf`, `docs/Legal/AgriX_Ugovor_o_licenciranju.docx`, `docs/Product/AgriX_Definicija_proizvoda.pdf`, `docs/Finance/AgriX_Finansijski_model.xlsx`). Odluke 401–408 su u §26, a 409–422 u §26.1; ovaj indeks ih obuhvata.
>
> `09B_ODLUKE_PO_OBLASTIMA_2026-07-26.pdf` je renderovani snimak verzije 2 na dan 26.07.2026. i zadržava se kao istorijski snimak. Kad se indeks menja, menja se `.md`; PDF je snimak i ne mora se osvežavati uz svaku izmenu.

Oznake: `⚠` stavka učestvuje u nerazrešenom konfliktu · `→ X` izmenjena odlukom X · `(nova)` doneta 26.07.2026. · `(27.07)` doneta 27.07.2026.

---

## 1. Definicija proizvoda i paketi

| ID | Odluka |
|---|---|
| 1 | Jedinstven sistem teren–centrala; do fakture PWA i Desktop ravnopravni, od fakture Desktop-only |
| 2 | Desktop-only je legitimno kompletan proizvod, jeftiniji, operativno manje efikasan |
| 3 | Desktop Core: otkup, dokumenta, prijemnice, fakture, ambalaža, repromaterijal, skladište, izveštaji |
| 4 | SEF, Banka i Dispatch su posebno plaćeni moduli |
| 5 | Management PWA uključen u svaki paket bez veštačkih ograničenja |
| 6 | Nazivi paketa: AgriX Desktop i AgriX Mobile |
| 7 | Mobile standardno uključuje PWA Otkupac i PWA Vozač |
| 8 | Mobile nikada nije samostalan; uvek uključuje kompletan Desktop |
| 9 | Osnovni PWA Vozač: preuzimanje, zbirni dokument, količine, status, sinhronizacija |
| 10 | Dispatch: vozila, vozači, rute, kapaciteti, neraspoređena roba, dispečerski pregled |
| 11 | SEF i Banka se u početku prodaju odvojeno; bundle moguć kasnije |
| 12 | Core dozvoljava ručni unos novčanih transakcija |
| 13 | Banka automatizuje izvode, uplate, rasknjižavanje, avanse, platne naloge |
| 14 | Kartice i salda ostaju u Core-u; Banka ih automatizuje |
| 15 | SEF: slanje, statusi, storniranje, preuzimanje ulaznih faktura, povezivanje |
| 16 | Sledljivost i WMS ostaju u Core-u |
| 17 | Repromaterijal i agrohemija zasad ostaju u Core-u |
| 18 | Kreiranje faktura je Core; SEF integracija je modul |
| 19 | Svi postojeći standardni izveštaji su Core |
| 59 | Jedan proizvod — nema small/mid/large izdanja ⚠ |
| 60 | Klijent bira režim rada kroz podešavanja istog proizvoda |
| 128 | Desktop baza, Mobile dodatak, Hladnjača/Proizvodnja nezavisan dodatak, ostali odvojeni |
| 129 | Mobile pokriva teren i transport; proizvodnja ostaje Desktop |
| 136 | Nema limita ni licence po korisniku na Desktopu |
| 137 | Desktop je single-active-user; Management PWA podržava više pregleda |
| 138 | Budući multi-user ulazi u postojeću pretplatu |
| 139 | Multi-user pre 2027 samo ako ga zahteva konkretan ugovor |
| 147 | Core kriterijum: funkcija potrebna za osnovno obećanje proizvoda |
| 148 | Modul kriterijum: merljiva dodatna vrednost koja se može posebno platiti |
| 270 | Enterprise je proizvod; Desktop i Mobile su komercijalni paketi |
| 286 | Desktop se predstavlja kao potpuno validan i kompletan paket |
| 287 | Desktop i Mobile kao dve ravnopravne kolone |
| 288 | Mobile: sve iz Desktopa + Otkupac + Vozač + real-time sync + Otkup uživo |
| 289 | Otkup uživo je atraktivna, ali ne centralna prodajna funkcija |
| 290 | Management Mobile se ističe u obe kolone |
| 291 | Obim Management Mobile-a se razlikuje; Otkup uživo i Dispatch samo uz Mobile |
| 292 | Javno: Mobile Otkupac, Mobile Vozač, Mobile MGMT; termin PWA se ne koristi |
| 293 | Dispatch dostupan isključivo uz Mobile |
| 294 | Sekcija dodatnih modula: Hladnjača/Proizvodnja, SEF, Banka, Dispatch |
| 295 | Redosled modula prema stranici; na glavnoj prednost Hladnjača/Proizvodnja |
| 296 | Moduli zasad nemaju posebne stranice |

---

## 2. Cenovni model

| ID | Odluka |
|---|---|
| 25 | Godišnja pretplata uključuje korišćenje, ažuriranja, bug-fix, podršku |
| 26 | Nema mesečnog modela |
| 28 | Sezonski klijenti plaćaju punu godišnju pretplatu |
| 29 | Cena po pravnom licu + naknada po aktivnoj stanici |
| 30 | Svako pravno lice plaća punu osnovnu cenu → 418|
| 31 | Nema naplate po korisniku ili uređaju |
| 32 | Hardver je odvojena jednokratna stavka |
| 33 | Hardverska podrška je posebna godišnja naknada |
| 34 | Svaki izlazak na teren se naplaćuje |
| 110 | Jedna standardna satnica uz mogući individualni popust → 409|
| 113 | Isti paket i broj stanica = ista cena → 418|
| 118 | Desktop → Mobile usred godine: srazmerna razlika |
| 119 | Mobile → Desktop samo pri obnovi, bez refundacije |
| 120 | Nova stanica usred godine: srazmerna naknada → 324 |
| 121 | Osnovni paket uključuje do pet stanica |
| 122 | Svaka stanica preko pet ima istu fiksnu cenu |
| 123 | Cena stanice ista u oba paketa |
| 124 | SEF, Banka, Dispatch: fiksna godišnja cena po pravnom licu |
| 125 | Cena modula ne zavisi od paketa |
| 126 | Cena Mobile-a je utvrđen odnos prema Desktopu → 414|
| 127 | Desktop + Mobile najmanje dvostruko od Desktop Otkup → 414|
| 130 | Proizvodni dodatak dobija fiksnu naknadu nakon standardizacije → 421|
| 131 | Proizvodni dodatak pokriva jedan pogon |
| 132 | Dodatni pogon: dodatna instanca i dodatni proizvodni dodatak |
| 133 | Cena dodatne instance individualno → 413|
| 135 | Moduli se plaćaju jednom po pravnom licu, važe kroz sve instance |
| 149 | Kod izdvajanja modula postojeći plaćaju tek od obnove |
| 150 | Nove cene odmah novima, prelazni period postojećima |
| 151 | Zaštita cene: jedna naredna godina po staroj ceni |
| 152 | Osnovna pretplata se ne menja tokom plaćenog perioda |
| 156 | Objavljuju se rasponi za Desktop i Mobile → 416|
| 157 | Objavljuje se raspon za Hladnjača/Proizvodnja → 416|
| 158 | Objavljuje se tačan iznos po dodatnoj stanici |
| 159 | Objavljuju se tačne cene Gazdinstvo Basic i Pro |
| 297 | U kolonama „od … godišnje", detalji niže |
| 298 | Početna cena: jedno pravno lice + do pet stanica |
| C4 | Promene cena pri obnovi su diskrecione |
| **323** | **(nova)** Aktivna stanica = stanica sa makar jednim otkupnim blokom |
| **324** | **(nova)** Stanice se prijavljuju unapred; usklađenje na obnovi, bez povraćaja |
| **329** | **(nova)** Cene u EUR, plaćanje u RSD po srednjem kursu na dan uplate |
| **334** | **(nova)** Postojeći klijenti zadržavaju cene za ovu sezonu |
| **337** | **(nova)** Smanjenje broja stanica ne menja cenu ako je stanica bila aktivna |
| **349** | **(nova)** Cene paketa: Desktop 500, Desktop all-in 1.200, Mobile 1.500, Mobile all-in 2.200 € → 414, 415 |
| **350** | **(nova)** Moduli 200 € svaki; Hladnjača/Proizvodnja 400 € → 412|
| **351** | **(nova)** Aktivna stanica preko pet — 50 € godišnje |
| **352** | **(nova)** GGAP od 1.000 € po pravnom licu |
| **353** | **(nova)** Satnica 50 € za razvoj po zahtevu i složenu migraciju |
| **354** | **(nova)** Obuka: 5 sati uključeno u onboarding, preko toga 30 € po satu |
| **355** | **(nova)** Izlazak na teren 50 € + gorivo + vreme puta i vreme rada |
| **356** | **(nova)** Hardver se prodaje sa oko 100 € marže po stanici |
| **357** | **(nova)** Hardverska podrška 40 € po stanici godišnje, minimum 200 € — potvrđeno 27.07. |
| **358** | **(nova)** Druga instanca −50 % na sve što ta instanca dodatno koristi; moduli po #135 se ne dupliraju → 413|
| **367** | **(nova)** Jedno gratis pravilo: prva godina od produkcijskog puštanja *(zamenjuje #144 i #149)* |
| **368** | **(nova)** Softverski modul i konfiguracija besplatni; fizički rad na lokaciji i puštanje opreme u rad se naplaćuju *(razrešava #68 vs #245)* |
| **406** | **(27.07)** Cena stanice ista bez obzira na režim rada; razliku pokriva cena Mobile paketa |
| **407** | **(27.07)** Hardver ostaje na planiranoj marži do izbora dobavljača |
| **408** | **(27.07)** Cene su određene; unit economics je kasnije fino podešavanje, ne preduslov |
| **409** | **(27.07)** Dve satnice: razvojna 50 €/h i implementaciona 30 €/h, prema prirodi posla *(menja 110)* |
| **410** | **(27.07)** Izlazak na teren 50 € + gorivo; vreme puta uvek po 30 €/h, vreme rada po prirodi posla |
| **411** | **(27.07)** C7 raspoređen: čišćenje/IT/konsalting po 30 €/h, masovne korekcije i izveštaji po 50 €/h |
| **412** | **(27.07)** SEF, Banka i Dispatch po pravnom licu; Hladnjača/Proizvodnja po proizvodnom pogonu |
| **413** | **(27.07)** Dodatna instanca −50 % na listu cena stavki koje koristi; dodatni pogon 200 € *(zamenjuje 133)* |
| **414** | **(27.07)** Mobile = Desktop + fiksni dodatak 1.000 €, isti na oba nivoa *(zamenjuje 126 i 127)* |
| **415** | **(27.07)** All-in = bazna + fiksna doplata 700 €; Desktop all-in bez Dispatch-a, Mobile all-in sa njim |
| **416** | **(27.07)** Bazne cene paketa „od X €"; moduli, stanica, Gazdinstvo i Savetnik tačni iznosi *(zamenjuje 156 i 157)* |
| **417** | **(27.07)** GGAP u cenovniku samo uz oznaku „na upit, uz potvrdu obima — nije deo standardne ponude" |
| **418** | **(27.07)** Nema pregovaračkih ni individualnih popusta *(briše C3, 111, 112)* |
| **421** | **(27.07)** Hladnjača/Proizvodnja 400 € po proizvodnom pogonu; uslov iz 130 ispunjen |
| **422** | **(27.07)** Prva godina od produkcijskog puštanja je jedino gratis pravilo *(prepisuje IP4)* |

---

## 3. Ugovor, obnova, raskid

| ID | Odluka |
|---|---|
| 27 | Standardni ugovor traje 12 meseci |
| 44 | Pri prestanku: kompletan izvoz + opciona plaćena migracija |
| 45 | Tranzicioni rok 30 dana |
| 114 | Pretplata teče od datuma aktivacije → 326 |
| 115 | Obnova nije automatska |
| 116 | Obaveštenje 30 dana pre isteka |
| 117 | Bez obnove: read-only 30 dana → 361 |
| 134 | Sve instance istog lica imaju isti ugovor i datum obnove |
| 153 | Novi klijenti plaćaju pre aktivacije |
| 154 | Obnova: faktura sa rokom 30 dana |
| 155 | Kašnjenje pri obnovi se rešava individualno |
| C5 | Prevremeni raskid samo kod bitne povrede od strane AgriX-a |
| C6 | Kašnjenje krivicom klijenta pauzira implementaciju, rok teče → 338 |
| LEG4 | Odgovornost ograničena na plaćanja u prethodnih 12 meseci |
| **325** | **(nova)** Nema plaćanja usred sezone; ugovori se sklapaju pre sezone |
| **326** | **(nova)** Jedinstveni predsezonski datum obnove; prvi period srazmeran |
| **338** | **(nova)** Jednokratni prenos neiskorišćenog perioda kod kašnjenja klijenta |
| **361** | **(nova)** Read-only režim je razvojni prioritet sa rokom pre 1. juna 2027; do tada se tretira kao postojeći |
| **376** | **(nova)** Ugovor piše osnivač; pravnik radi pregled gotovog nacrta |

---

## 4. Podrška i SLA

| ID | Odluka |
|---|---|
| 23 | Bug-fix, bezbednosne ispravke i ažuriranja u pretplati |
| 38 | Standardna podrška: kanali, dijagnostika, pomoć, bug-fix, sezonski prioritet |
| 39 | Kritični incidenti pokriveni i van radnog vremena ⚠ |
| 50 | Potvrda prijema i početak dijagnostike u roku od 1 sata ⚠ |
| 51 | Definicija kritičnog incidenta |
| 52 | Nekritični: odgovor u jednom radnom danu |
| 53 | Radno vreme podrške 08:00–16:00 radnim danima |
| 54 | Vikend podrška u sezoni → 332 |
| 55 | Vikend radno vreme 08:00–16:00 |
| 56 | Sezona po klijentu → 331 |
| 57 | Vikend podrška uključena u pretplatu → 332 |
| C7 | Čišćenje podataka, masovne korekcije, posebni izveštaji, teren, IT setup i procesni konsalting se naplaćuju |
| M12 | Jedinstvena matrica podrške; nema SLA nivoa po klijentu |
| M13 | Nema javne garancije dostupnosti |
| M14 | Interni cilj dostupnosti se meri |
| M15 | Dostupnost se meri poslovnim tokovima |
| **327** | **(nova)** Klijent proverava svoja dokumenta; AgriX ispravlja bug u najkraćem roku |
| **331** | **(nova)** Sezona se definiše jedinstveno na nivou AgriX-a |
| **332** | **(nova)** Vikend podrška u sezoni samo za kritične incidente |
| **359** | **(nova)** Rok od 1 sata važi unutar definisanog proširenog prozora; van njega best effort *(menja #50)* |
| **378** | **(nova)** Sezona traje od 1. juna do 30. novembra *(precizira 331)* |

---

## 5. Onboarding i implementacija

| ID | Odluka |
|---|---|
| 35 | Onboarding se može proceniti, ali prvim klijentima uglavnom nije naplaćivan ⚠ |
| 36 | Osnovni uvoz šifarnika u početku može biti besplatan |
| 37 | Složena migracija se naplaćuje posebno |
| 245 | Uvođenje modula postojećem klijentu uvek besplatno ⚠ |
| 302 | Tipičan raspon uvođenja + individualni plan |
| 303 | Uslovi: odgovorna osoba, vreme pre sezone, sređeni podaci |
| 304 | Hijerarhija: odgovorna osoba → vreme → podaci |
| 305 | Pred sezonu standardni paket bez customa, onboarding ~pola dana |
| 306 | Nema skraćenog proizvoda; nestandardni zahtevi se odlažu |
| C2 | Instalacija uvek uključena; onboarding preko osnovnog nivoa se naplaćuje ⚠ |
| ON1 | Klijent odgovara za poslovnu tačnost početnih podataka |
| ON2 | Početna obuka uključena, kasnije se naplaćuju ⚠ |
| ON3 | Zajednički produkcijski sign-off pre starta ⚠ |
| ON4 | Pojačana podrška posle go-live samo ako je posebno ugovorena |
| **362** | **(nova)** Modul je besplatan uz uključenih X sati obuke; preko toga naplata *(razrešava #245 vs ON2)* |
| **363** | **(nova)** Skraćeni sign-off sa izričitim prihvatanjem rizika za predsezonski start *(razrešava ON3 vs #305)* |
| **365** | **(nova)** Fiksni uključeni obim: instalacija, povezivanje svega i 5 sati obuke; preko toga naplata *(razrešava C2 vs #35)* |

---

## 6. Prodaja, demo, kvalifikacija

| ID | Odluka |
|---|---|
| 71 | Ciljna reakcija posle demoa: „ovo pokriva celu firmu" |
| 100 | Svi segmenti u scoring; veliki sistemi zasad nisu primarni |
| 101 | Rang lead-a: brzina zatvaranja, prihod, fit, referentna vrednost |
| 102 | Nema automatskog odbijanja segmenta |
| 252 | Demo počinje problemom klijenta, završava celim tokom |
| 253 | Pre demoa kratak razgovor o procesima |
| 254 | Prilagođen demo samo uz pristup odlučiocu |
| 255 | Nema probnog perioda ni pilota; samo dummy demo ⚠ → C1 |
| 256 | Kvalifikovan lead može dobiti ograničen samostalni demo pristup |
| 257 | Jedan standardni demo scenario za sve |
| 258 | Demo prikazuje ceo ekosistem; ponuda odvaja kupljeno od opcionog |
| 259 | Prototipovi jasno označeni kao nedostupni za ugovaranje |
| 273 | Primarni CTA „Zakažite demonstraciju" |
| 274 | Kvalifikaciona forma sa pet polja |
| 275 | Kvalifikovan odmah zakazuje; nejasan ide na razgovor |
| C1 | Trial režim sa punom funkcionalnošću → 371 |
| **371** | **(nova)** Trial postoji i dolazi posle vođenog demoa i kvalifikacije *(razrešava C1 vs #255)* |

---

## 7. Marketing, brend i sajt

| ID | Odluka |
|---|---|
| 99 | Kanali: direktno i demo → SEO i oglasi → partneri |
| 194 | Gazdinstvo zasad sekcija glavnog sajta |
| 195 | Zaseban sajt tek kada sadržaj naruši jasnoću |
| 224 | Savetnik: posebna stranica unutar glavnog sajta |
| 268 | Javno lice je brend, ne osnivač |
| 269 | Krovni brend: Enterprise, Gazdinstvo, Savetnik |
| 271 | Suština kompletan operativni sistem; ulaz jednostavniji |
| 272 | SEO „Softver za otkup i hladnjače"; glavna stranica šira formulacija |
| 276 | Glavna stranica + tri funkcionalne stranice |
| 277 | Sve tri paralelno; blaga prednost stranici Otkup |
| 278 | Otkup: QR kao dokaz brzine + jednokratan unos kao obećanje |
| 279 | Hladnjača: sledljivost kao obećanje, kontrola prerade kao vrednost |
| 280 | Skladište: kontrola paleta + istorija porekla |
| 281 | Prvo tok kroz firmu, zatim paketi i moduli |
| 282 | Dijagram objašnjava, snimci dokazuju |
| 283 | Dijagram pa postojeći mobilni video-demo |
| 284 | Mobilni video odmah prati Desktop prikaz |
| 285 | Prvo tok teren → centrala, zatim širi Desktop pregled |
| 299 | Odnos prema ERP-u u FAQ-u i na demou |
| 300 | Prvo FAQ: radi li AgriX bez Mobile paketa |
| 301 | Drugo FAQ: koliko traje uvođenje |
| MKT1 | Hladnjače i otkupljivači prioritetni segment |
| MKT2 | Primarno tržište Srbija |
| MKT3 | Samo referral partneri; nema resellera |
| MKT4 | Gate za marketing budžet → 374 |
| **374** | **(nova)** Tri nivoa potrošnje, svaki sa dvostrukim uslovom (kanal konvertuje + kapacitet postoji), plus pravilo povratka na niži nivo |
| MKT5 | Prva nova uloga: onboarding i podrška |

---

## 8. Reference i studije slučaja

| ID | Odluka |
|---|---|
| 260 | Javna referenca bez izričite zabrane → 318 |
| 261 | Vredni su i javno ime i dokazivi rezultati |
| 262 | Mere se vreme, greške, kontrola robe, sledljivost, upravljački pregled |
| 263 | AgriX priprema nacrt, klijent potvrđuje pre objave |
| 264 | Objava čim postoje merljivi rezultati |
| 265 | Pisana studija + kratka video izjava |
| 266 | Precizne brojke gde nisu osetljive, procenti gde jesu |
| 267 | Nema finansijskih podsticaja za referencu |
| 318 | Ime, logo, studija, foto i video samo uz izričitu saglasnost |

---

## 9. Tržište, pozicioniranje i granice

| ID | Odluka |
|---|---|
| 58 | Od jedne stanice do velikih sistema |
| 61 | Voće i povrće glavni fokus |
| 62 | Hladnjače su ključna ciljna grupa |
| 64 | Kompletan operativni sistem za otkup, preradu, skladište, sledljivost |
| 72 | AgriX nije računovodstveni ERP |
| 73 | AgriX operacije, ERP glavna knjiga i zakonsko računovodstvo |
| 74 | Trajne granice: ne opšti ERP, ne generički ERP, ostati u agroindustriji |
| 75 | Ekosistem: Enterprise, Gazdinstvo, GGAP |
| 76 | North Star: platforma celog lanca |
| 92 | Prednost: cela firma + unos na izvoru + brz razvoj |
| 93 | Pretnja: finansiran konkurent sa sales timom |
| 94 | Bottleneck na 30–50 klijenata je prodaja |
| 95 | Prvi prodavac tek oko 30–50 klijenata |
| 97 | Vrednost: Enterprise → + Gazdinstvo → podaci |
| 98 | Rizici: prespor rast, generički ERP, previše customa |
| 238 | Drugi prerađivači samo kada se prirodno uklapaju |
| 239 | Zajedničko jezgro, posebni paketi po segmentu ⚠ |
| 240 | Svaki vertikalni paket ima svoj scope i cenu ⚠ |
| 241 | Razvoj vertikale tek uz kupca ili potvrđenu tražnju |
| 242 | Mogući red posle 2027: žitarice, duvan, sušare, vinarije |
| 247 | Reference do 2027 primarno hladnjače za voće i povrće |
| 248 | Aktivna prodaja samo u Srbiji do 2027 |
| 249 | Cilj 10–20 aktivnih pravnih lica do 2027 → 375 |
| 250 | Proizvodni modul koristi preko 80% klijenata |
| 251 | Hladnjača/Proizvodnja standardni deo svake ponude |
| **333** | **(nova)** Nema ekskluzivnosti; konkurenti mogu oba biti klijenti |
| **375** | **(nova)** Scenario rasta C: 12–15 novih Enterprise klijenata do sezone 2027, ukupno 15–18 |
| **377** | **(nova)** Nema referral provizija van slučaja Savetnika |
| **403** | **(27.07)** Važi fiksan ciljni broj klijenata, ne readiness cap *(povlači STR-001)* |

---

## 10. AgriX Gazdinstvo

| ID | Odluka |
|---|---|
| 77 | Četiri kanala finansiranja Basic i Pro |
| 78 | Vrednost i bez Enterprise veze |
| 79 | Dva growth engine-a |
| 80 | Preko Enterprise-a podrazumeva se Basic |
| 81 | Proizvođač je primarni korisnik; nije white-label |
| 160 | Basic se može kupiti bez Enterprise-a |
| 161 | Prvih 50 Basic korisnika partner dobija bez naknade |
| **339** | **(nova)** Jedinstvena kanalska cena za sve partnerski posredovane naloge — 10 € Basic, 20 € Pro; maloprodajna ostaje 19 / 39 € |
| **343** | **(nova)** Jedan Pro po proizvođaču — ko prvi aktivira taj plaća, druga strana ne plaća ponovo |
| 162 | Prioritet 2027: ključne Pro funkcije bez usporavanja Enterprise-a |
| 163 | Oba kanala paralelno, prioritet Enterprise |
| 164 | Cilj 2027 je rast plaćenih Pro korisnika |
| 165 | Direktno i preko hladnjače ravnopravno |
| 166 | Pro preko hladnjače traje do isteka njenog ugovora |
| 167 | Pro usred godine srazmerno |
| 168 | Prekid saradnje deaktivira finansirani Pro |
| 169 | Posle gašenja povratak na Basic, podaci vidljivi |
| 170 | 30 dana besplatnog Basic-a bez kartice |
| 171 | Posle probe bez uplate read-only |
| 172 | Samo godišnja pretplata |
| 173 | Aktivacija na poverenje; blokada posle 7 dana |
| 174 | Blokada traje do evidentirane uplate |
| 175 | Pro proba samo uz plaćen Basic |
| 176 | Jedna Pro proba na poverenje |
| 177 | Po isteku povratak na Basic, Pro podaci zaključani |
| 178 | Upgrade uz srazmernu doplatu |
| 179 | Obnova Pro-a puna godišnja cena |
| 180 | Downgrade pri obnovi; podaci ostaju vidljivi |
| 181 | Prestanak Enterprise ugovora: 30 dana za samostalnu obnovu |
| 182 | Tih 30 dana read-only |
| 183 | Istorija saradnje trajno dostupna proizvođaču |
| 184 | Samostalan izvoz kompletne istorije |
| 185 | Brisanje naloga briše samostalne podatke; dokumenti ostaju hladnjači |
| 186 | Hladnjača priprema, proizvođač aktivira i prihvata uslove |
| 187 | Hladnjača vidi samo svoj poslovni odnos |
| 188 | Dodatni podaci samo uz posebno odobrenje ⚠ |
| 189 | Pravila povlačenja saglasnosti → 369 |
| **369** | **(nova)** Povlačenje saglasnosti deluje samo ubuduće; već izdati dokumenti ostaju nepromenjeni |
| 190 | Širenje i validacija paralelno |
| 191 | Poruka zavisi od kanala |
| 192 | Brend AgriX Gazdinstvo, nije white-label |
| 193 | Basic bez veštačkih limita |
| 321 | Kompletan proizvod i bez ijedne AgriX hladnjače |
| PRT2 | Pro direktno ili preko hladnjače |
| **404** | **(27.07)** Gazdinstvo je launch ready — prelazi iz Pilot only u Standard offer |

---

## 11. AgriX Savetnik

| ID | Odluka |
|---|---|
| 196 | Ciljni korisnici: komercijalni proizvođači i savetnici |
| 197 | Jedan interfejs za više gazdinstava |
| 198 | Naplata po broju aktivnih gazdinstava → 420|
| 199 | Poseban proizvod AgriX Savetnik |
| 200 | Cena pokriva gazdinstva; ona ne plaćaju Pro → **zamenjena odlukom 340** |
| 201 | Proizvođač zadržava svoj nalog |
| 202 | Kupci: samostalni agronomi i savetodavne firme |
| 203 | Osnovna verzija do 2027 |
| 204 | Prva verzija: jedan savetnik, više gazdinstava |
| 205 | Gazdinstvo može biti povezano i sa Savetnikom i sa hladnjačama |
| 206 | Interne agronomske službe ravnopravna grupa |
| 207 | Ista tarifa za oba tipa → 419|
| 208 | Samo godišnja pretplata |
| 209 | Besplatnih 30 dana → 346 |
| 210 | Limit probe 10 gazdinstava |
| 211 | Savetnik dobija planerske i kontrolne funkcije |
| 212 | Obavezujući nalog ili neobavezna preporuka |
| 213 | Automatski uvid u status i odstupanja |
| 214 | Proizvođač evidentira odstupanje, savetnik dobija upozorenje |
| 215 | Enterprise klijent sa agronomima posebno kupuje Savetnik |
| 216 | GGAP minimum u GGAP modulu; iznad toga Savetnik |
| 217 | Javno se nudi čim bude stabilan |
| 218 | Direktno obraćanje savetnicima |
| 219 | Savetnik može prodavati Gazdinstvo uz proviziju |
| 220 | Provizija za prvu prodaju i svaku obnovu |
| 221 | Fiksni iznos, ne procenat |
| 222 | I Gazdinstvo i Enterprise, različiti iznosi |
| 223 | Enterprise provizija samo jednokratno |
| 225 | Cena se ne objavljuje → **zamenjena odlukom 347** |
| 226 | Samostalna registracija i trenutni start probe |
| 227 | Posle probe read-only |
| 228 | Prekid: proizvođač zadržava sve, savetnik gubi pristup |
| 229 | Zasad alat, dugoročno platforma |
| 230 | Marketplace ne pre kraja 2027 |
| 231 | Redosled post-2027 ostaje svesno otvoren |
| PRT3 | Savetnik kao management sloj nad više gazdinstava |
| **340** | **(nova)** Gazdinstva u portfelju nisu uključena u cenu Savetnika; drže sopstvenu Pro pretplatu po kanalskoj ceni *(zamenjuje #200)* |
| **341** | **(nova)** Cena Savetnika: osnovica 150 € uključuje do 10 gazdinstava, svako preko toga 15 € — potvrđeno 27.07. → 419 |
| **342** | **(nova)** Aktivno gazdinstvo = ono kojem je savetnik u toku godine poslao makar jedan nalog ili preporuku |
| **344** | **(nova)** Savetnik može platiti Pro u ime proizvođača i ugraditi to u svoju naknadu — posrednička uloga po MKT3 |
| **345** | **(nova)** Nema cashbacka za gazdinstva u portfelju; podsticaj je sam alat. #221 ostaje samo za preporuke van portfelja |
| **346** | **(nova)** Proba obuhvata i Pro za do 10 gazdinstava |
| **347** | **(nova)** Cena Savetnika se objavljuje — osnovica i cena po gazdinstvu *(zamenjuje #225)* |
| **348** | **(nova)** Interne agronomske službe plaćaju samo alat kada su kooperanti već pokriveni partnerskim paketom |
| **401** | **(27.07)** Savetnik je treći ravnopravan stub uz Enterprise i Gazdinstvo *(potvrđuje 269)* |
| **419** | **(27.07)** Dve tarife: standalone 150 €/15 €, Enterprise 100 €/10 € *(zamenjuje 207)* |
| **420** | **(27.07)** Osnovica do 10 gazdinstava + fiksni iznos po gazdinstvu preko 10 *(zamenjuje 198)* |

---

## 12. Hladnjača / Proizvodnja i oprema

| ID | Odluka |
|---|---|
| 63 | Proizvodni domen: sirovina, klasiranje, otpad, partije, ambalaža, palete, skladište |
| 65 | Cilj 2027: planiranje, norme, kapaciteti, radnici, učinak, integracije |
| 66 | Minimalni scope proizvodnje 2027 |
| 67 | Integracije samo za odobrene vage, PLC, senzore, mašine |
| 68 | Svaka instalacija i puštanje u rad se naplaćuje ⚠ |
| 69 | AgriX zadržava framework i kod |
| 70 | Široko primenjivu integraciju finansira AgriX, specifičnu prvi klijent |
| 232 | Prodaje se čim osnovni tok bude stabilan |
| 233 | Palete sveže i prerađene robe već u produkciji |
| 234 | Red razvoja: nalozi i norme → integracije → kapaciteti |
| 235 | Klijent kupuje samo ono što postoji na dan prodaje |
| 236 | Precizno navođenje postojećih funkcija |
| 237 | Primarna grupa: hladnjače sa prijemom, klasiranjem, zamrzavanjem |
| 243 | Nema naknade dok se modul standardizuje |
| 244 | Naplata posle provere kod jednog klijenta |
| 246 | Prednost funkcijama koje donose ugovor |

---

## 13. GGAP

| ID | Odluka |
|---|---|
| §8 | GGAP je Enterprise dodatak koji kupuje hladnjača, nije deo Pro-a |
| §8 | Korisnik u GGAP-u dobija potrebne funkcije bez Pro naknade |
| §8 | Jedna fiksna godišnja cena po pravnom licu |
| §8 | Konsalting i dokumentacija se naplaćuju odvojeno |
| §8 | Prvo softver, zatim mreža stručnjaka, dugoročno interni tim |
| §8 | Konsultant direktno ili podugovoren kroz AgriX |
| §8 | Softver nikada ne garantuje sertifikat |
| §8 | Prodaja tek posle validacije i jednog uspešnog projekta |
| §8 | Do 2027 samo konceptualna priprema |
| PRT4 | Aktivacija GGAP-a otključava funkcije u Gazdinstvu |
| **402** | **(27.07)** GGAP je modul Enterprise-a, ne stub; koriste ga samo Enterprise klijenti *(ukida STR-012)* |
| **405** | **(27.07)** GGAP ostaje van komercijalne ponude do validacije |

---

## 14. Vlasništvo podataka i Multi-Enterprise

| ID | Odluka |
|---|---|
| 40 | Desktop podaci lokalni uz periodične kopije |
| 41 | PWA/GAS/Sheets obezbeđuje AgriX |
| 42 | Silo po pravnom licu |
| 43 | Klijent vlasnik, AgriX obrađivač ⚠ → LEG1 |
| 82 | Multi-Enterprise dugoročno |
| 83 | Trenutno jedna Enterprise veza |
| 84 | Dugoročno globalni identitet proizvođača |
| 85 | Globalni identitet firme već postoji |
| 86 | Dugoročno globalni identitet parcele |
| 87 | Dugoročno globalni katalog proizvoda |
| 88 | Desktop broj kanonski, PWA broj privremen → 372 |
| 89 | Duplikati se sprečavaju, brojevi trajni |
| 319 | Podaci pripadaju onome ko ih je stvorio |
| 320 | Gazdinstvo master ličnih podataka; firma odobrava promenu |
| A1 | Proizvođač zadržava istoriju i posle odlaska hladnjače |
| A3 | Matični podaci se uređuju kroz Desktop |
| A4 | Workflow odobravanja izmena ličnih podataka |
| **372** | **(nova)** Numeracija i prelazak privremenog u konačan broj rešeni su u kodu; pravilo treba zapisati u dokumentaciju |

---

## 15. Poslovna pravila i autorizacija

| ID | Odluka |
|---|---|
| A5 | Plan i najava nisu osnova za otkupni list |
| A6 | Najave se koriste za planiranje kapaciteta |
| A7 | Prati se pouzdanost najava |
| A8 | Preporuke upozoravaju; regulatorna pravila blokiraju ⚠ |
| A9 | Svako pravilo definiše uloge za odobrenje izuzetka |
| A10 | Apsolutna i uslovna pravila |
| A11 | Invarijante u proizvodu; klijent konfiguriše parametre |
| A12 | Dokument se tumači prema verziji u trenutku nastanka |

**Principi:** Business Ownership · Immutable Business History · Product-Driven Business Rules · Planning ≠ Execution ≠ Legal · Versioned Behavior

---

## 16. Bezbednost i pristup

| ID | Odluka |
|---|---|
| S1 | Stalni pristup podrške produkciji, po ulozi i uz audit ⚠ |
| S2 | Nema produkcionih podataka u razvoju |
| S3 | Imenovani nalozi; nema deljenih |
| S4 | MFA za privilegovane uloge |
| S5 | Stroga izolacija pravnih lica |
| **373** | **(nova)** Jedini podobrađivač je Google; lokaciju podataka treba verifikovati u Workspace konzoli i zapisati |

---

## 17. Monitoring, incidenti, operabilnost

| ID | Odluka |
|---|---|
| M1 | Monitoring je zasad interni alat |
| M2 | Samo bezbedan i idempotentan auto-recovery |
| M3 | Centralna matrica incidenata ⚠ |
| M4 | SLA sat kreće od detekcije ⚠ |
| M5 | Agregirana telemetrija dozvoljena ⚠ → LEG3 |
| M6 | Ograničeno čuvanje detaljnih događaja |
| M7 | Rollout stop i rollback samo ručno |
| M8 | Obaveštavanje prema ozbiljnosti |
| M9 | Obavezan postmortem ⚠ |
| M10 | AuditCritical odvojen i neizmenjiv |
| M11 | Zasad na zahtev; dugoročno read-only pregled klijentu |
| M16 | Spoljni servisi ne ulaze u metriku |
| M17 | Kontrolisana degradacija po toku |
| M18 | Lokalni status odvojen od integracionog |
| M19 | Operativni koraci dalje; pravni i finansijski čekaju |

---

## 18. Životni ciklus podataka

| ID | Odluka |
|---|---|
| 46 | Dnevne kopije 30 dana, mesečne 12 meseci |
| 47 | Backup posle Journal promene, pri otvaranju, dnevno off-site |
| 48 | Cloud RPO dnevni |
| 49 | RTO 24 sata ⚠ |
| A2 | Istorijski dokumenti nepromenljivi; ispravka kroz storno |
| D1 | Neograničeno arhiviranje samo uz saglasnost |
| D2 | Standardni izvoz uključen; migracija posebno |
| D3 | Brisanje obuhvata backupe → LEG2 |
| D4 | Anonimizovani agregati ostaju ⚠ → LEG3 |
| D5 | Osnovni izvoz samostalno, arhivski preko podrške |
| LEG2 | Brisanje iz aktivnih; backupi netaknuti do isteka retention-a ⚠ |
| **370** | **(nova)** Restore rezervne kopije ponovo primenjuje izvršena brisanja, uz zabelešku |

---

## 19. Integracije

| ID | Odluka |
|---|---|
| I1 | Zatvoren model; nema opšteg API-ja |
| I2 | Jedan standardni konektor po sistemu |
| I3 | Autoritativni izvor po domenu |
| I4 | Konflikti se prijavljuju, ne prepisuju |
| I5 | Osnovni uvoz/izvoz uključen; dvosmerna integracija se plaća |

---

## 20. ML i AgriX Intelligence

| ID | Odluka |
|---|---|
| ML1 | Zajednički modeli na anonimizovanim podacima → 364 |
| ML2 | ML ne određuje cenu, ne odbija, ne blokira |
| ML3 | Čuvaju se verzija, ulazi, pouzdanost, objašnjenje |
| ML4 | Jedan zajednički model po nameni |
| ML5 | Plaćeni modul AgriX Intelligence |
| PRT5 | Zrelost prema kvalitetu podataka |
| LEG3 | Saglasnost i za nepovratno anonimizovane podatke → **revidirana odlukom 364** |
| **364** | **(nova)** Nepovratno anonimizovani podaci ne traže saglasnost; upotreba se transparentno navodi u ugovoru |

---

## 21. Platformska arhitektura i evolucija

| ID | Odluka |
|---|---|
| 90 | Sheets baza se ne briše → 366 |
| P1 | Migracija samo po objektivnim pragovima ⚠ (pragovi nedefinisani) |
| P2 | Poslovni model stabilan kroz tehnološku zamenu |
| P3 | Danas kanonska specifikacija u oba kanala; dugoročno backend ⚠ |
| P4 | Jedan autoritativni izvršilac po domenu |
| P5 | Postepena migracija klijenata |
| **366** | **(nova)** Sheets ostaje PWA backend dok se ne dostignu pragovi iz P1 *(preformuliše #90)* |

---

## 22. Razvoj, backlog i prioriteti

| ID | Odluka |
|---|---|
| 20 | Klijentski izveštaji se naplaćuju odvojeno |
| 21 | Klijentska funkcionalnost se naplaćuje; opšte vredna ulazi u proizvod |
| 22 | Nema trajnih forkova |
| 24 | Novi procesi i izmene se naplaćuju ⚠ → 328 |
| 91 | Proizvodni sistem ima najviši prioritet do 2027 |
| 103 | Veliki custom mora opravdati odlaganje roadmapa |
| 104 | Eksplicitno navesti šta se odlaže |
| 105 | Prioritet: incidenti → ugovor → roadmap → ostalo |
| 106 | Promena roadmapa uz pisanu nameru |
| 107 | Custom je time-and-materials |
| 108 | Procena sati i budžetski limit |
| 109 | Fakturisanje mesečno ili po završetku |
| 140 | Celine se puštaju kada postanu stabilne |
| 141 | Niskorizične šire, kritične prvo kod jednog |
| 142 | Uslovi za širi rollout |
| 143 | Pilot je onaj koji je tražio funkciju |
| 144 | Pilot funkcija besplatna do kraja ugovorne godine ⚠ |
| 145 | Posle pilota Core ili modul |
| 146 | Pilot uslovi individualno |
| 307 | Custom tokom sezone samo za ugovor ili ozbiljan problem |
| 308 | Tokom sezone direktno samo bug-fix, bezbednost, kritično |
| 309 | Postsezonski razgovor sa svakim klijentom |
| 310 | Postsezonski razgovor je i osnova za obnovu |
| 311 | Klijent dobija pisani rezime |
| 312 | Jedinstveni backlog sa četiri klase |
| 313 | Periodični pregled prioriteta ⚠ |
| 314 | Kod nesaglasnosti odlučuje osnivač |
| 315 | Roadmap se ne objavljuje |
| 316 | Nema glasanja klijenata |
| 317 | Reinvestiranje: podrška i onboarding → prodaja i marketing |
| **328** | **(nova)** Zakonske izmene AgriX radi o svom trošku |

---

## 23. Kvalitet i release

| ID | Odluka |
|---|---|
| L1 | Stari tok se uklanja posle prelaznog perioda |
| Q1 | Release sa nekritičnim nedostatkom uz workaround |
| Q2 | Svi release gate-ovi obavezni za redovan release |
| Q3 | Rollout prema riziku |
| Q4 | Hotfix prolazi iste gate-ove → 360 |
| **360** | **(nova)** Emergency gate: smanjen obavezan skup provera tokom incidenta, puna validacija i dokumentacija u roku od 24h po stabilizaciji |

---

## 24. Intelektualna svojina

| ID | Odluka |
|---|---|
| IP1 | AgriX zadržava vlasništvo i nad finansiranim funkcijama |
| IP2 | Ugovorom se definiše šta se sme generalizovati |
| IP3 | Nema izvornog koda ni escrow-a |
| IP4 | Finansijer dobija prioritet i prvu godinu besplatno → 422|
| IP5 | Ekskluzivnost samo uz poseban ugovor i višu cenu |

---

## 25. Kontinuitet i organizacija

| ID | Odluka |
|---|---|
| 96 | Arhitektura i strategija kod osnivača; ostalo se delegira |
| BC1 | Prihvaćena zavisnost od osnivača ⚠ |
| BC2 | Nema vanrednog continuity paketa ⚠ |
| BC3 | Druga tehnička osoba na 15–20 firmi → 336 |
| **335** | **(nova)** Standardni odgovor na pitanje kontinuiteta |
| **336** | **(nova)** BC3 postaje ugovorna obaveza |

---

## 26. Pravni okvir

| ID | Odluka |
|---|---|
| LEG1 | Uloge rukovalac/obrađivač otvorene do pravnika ⚠ |
| LEG4 | Odgovornost ograničena na 12 meseci plaćanja |
| LEG5 | Obaveštavanje o incidentu bez nepotrebnog odlaganja |
| **330** | **(nova)** Hosting van Googlea nije u ponudi do 2028 |

---

## 27. Pravilo tumačenja (§10)

- Kasnija odluka ima prednost nad ranijom.
- Implementirane funkcije ne predstavljati kao buduće.
- Roadmap nije prodajno obećanje.
- Nema trajnih klijentskih forkova.

> **Predlog dopune:** odluke koje su ušle u potpisan ugovor menjaju se samo aneksom, bez obzira na datum kasnije odluke.

---

## Otvorene stavke

Stanje posle cenovne revizije 27.07.2026.

| # | Stavka | Vezano za | Status |
|---|---|---|---|
| 1 | Pravni pregled ugovora nije obavljen | 376 | otvoreno |
| 2 | Prilog 3 — obrada podataka o ličnosti | LEG1 | otvoreno |
| 3 | Mesto nadležnog suda u ugovoru | član 15 | **zatvoreno 27.07.** — Niš, sedište je Merošina |
| 4 | Lokacija podataka u Google Workspace konzoli | 373 | otvoreno |
| 5 | Redosled post-2027 inicijativa | 231 | otvoreno svesno |
| 6 | Vertikalni paketi protiv odluke 59 | K-06 | odloženo |
| 7 | Desktop all-in štedi klijentu samo 100 € — slaba bundle poruka | 415 | komercijalno, nije konflikt |

Uz stavku 3: brief cenovne revizije vodio ju je kao otvorenu, ali je `09_QA_DECISION_LOG.md` §25.11 zatvara istog dana. Merodavan je log.

Uz stavku 7: Desktop nosi tri modula u vrednosti 800 €, a all-in doplata je 700 €. Snižavanje doplate ispod 600 € vratilo bi mrtvu tačku na dva modula. Prodajno rešenje je pozicioniranje kao „Hladnjača + dva modula gratis", što je tačno.

**Zatvoreno ovom revizijom:** cena po gazdinstvu kod Savetnika (341), hardverska podrška (357), vremenski prozor za rok od 1 sata (359), broj uključenih sati obuke (362), i nesaglasje ovog indeksa sa logom.

---

## Razrešeni konflikti

Svih četrnaest konflikata iz audita zatvoreno je u sesiji 26.07.2026. Cenovna revizija 27.07.2026. zatvorila je i pet preostalih stavki iz cenovnog i komercijalnog dela — odluke 409, 412, 413, 418 i 422.

| ID | Konflikt | Razrešenje |
|---|---|---|
| K-01 | #39/#50 vs BC1/BC2 — odziv 1h uz jednog čoveka | 359 |
| K-02 | #88 vs A2 — kada dokument postaje važeći | rešeno u kodu; 372 nalaže zapis pravila |
| K-03 | #117 vs `LicenseBlock` — read-only ne postoji | 361 |
| K-04 | #68 vs #245 — instalacija vs besplatan modul | 368 |
| K-05 | #144 / #149 / IP4 — tri gratis perioda | 367 |
| K-06 | #59 vs #239/#240 — jedan proizvod vs vertikale | odloženo do prve vertikale |
| K-07 | #90 vs P1/P3 — Sheets kao invarijanta | 366 |
| K-08 | Q2/Q4 — svi gate-ovi i za hotfix | 360 |
| K-09 | C1 vs #255 — trial vs bez probnog perioda | 371 |
| K-10 | LEG3 vs ML1/M5/D4 — saglasnost za anonimizovane podatke | 364 |
| K-11 | C2 vs #245/#35 — granica naplate onboardinga | 365 |
| K-12 | #245 vs ON2 — modul besplatan, obuka naplativa | 362 |
| K-13 | ON3 vs #305 — sign-off kod predsezonskog starta | 363 |
| K-14 | LEG2 — restore vraća obrisane podatke | 370 |

---

## Zamenjene odluke

| Stara | Zamenjena odlukom |
|---|---|
| #50 | 359 — rok od 1h važi unutar definisanog prozora |
| #56 | 331 — sezona se definiše jedinstveno |
| #54, #57 | 332 — vikend podrška samo za kritične incidente |
| #59 | ostaje, uz napomenu iz K-06 |
| #90 | 366 — Sheets do pragova iz P1 |
| #114 | 326 — jedinstveni predsezonski datum obnove |
| #144, #149 | 367 — jedno gratis pravilo |
| #200 | 340 — gazdinstva nisu uključena u cenu Savetnika |
| #225 | 347 — cena Savetnika se objavljuje |
| #249 | 375 — scenario C |
| #255 | 371 — trial postoji, posle demoa |
| #260 | 318 — izričita saglasnost za referencu |
| D3 | LEG2 + 370 |
| C1 | 371 |
| LEG3 | 364 |
| BC3 | 336 — postaje ugovorna obaveza |
| #110 | 409 — dve satnice umesto jedne |
| #126, #127 | 414 — Mobile = Desktop + fiksni dodatak 1.000 € |
| #133 | 413 — dodatna instanca −50 %, bez individualnog dogovora |
| #156, #157 | 416 — politika prikaza cena |
| #198 | 420 — osnovica + iznos po gazdinstvu preko 10 |
| #207 | 419 — dve tarife Savetnika |
| C3, #111, #112 | 418 — obrisane; nema pregovaračkih popusta |
| IP4 | 422 — ista prva besplatna godina, ne dodatna |
| #130 | 421 — cena Hladnjače više nije uslovljena |
