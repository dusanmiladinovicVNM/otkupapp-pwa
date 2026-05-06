# Production runbook: Fiskalni scanner, save i mapiranje artikala

Status: **operativni runbook za incidente “skenirao sam fiskalni račun, ali nije ušao u lager”, “račun je duplikat”, “artikal je pogrešno mapiran”, “račun je sačuvan, ali mapiranje nije naučeno”.**

Aplikacija: **OtkupApp / AgriX PWA**
Domen: **Kooperant PWA fiskalni račun → GAS parse → FISKALNI sheet / lager → mapiranje artikala**
Glavni kod: `src/js/features/kooperant/fiskalni.js`, `gas/Code.gs`

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Skenirao sam račun, ali nije ušao u lager.”
* “Piše da je račun već skeniran.”
* “QR neće da se očita.”
* “Slikaj opcija ne nalazi QR.”
* “Stavka nema artikal.”
* “Artikal je pogrešno mapiran.”
* “Napravio sam privatni artikal, ali ga ne vidim u šifarniku.”
* “Sačuvao sam račun, ali sledeći put opet moram ručno da mapiram.”
* “Jedna stavka je ušla, druga nije.”
* “Račun je ušao pod pogrešnog kooperanta.”
* “Management vidi račun, ali kooperant tvrdi da nije sačuvan.”

Prvo pravilo:

> Ne skeniraj isti račun ponovo i ne pravi ručni lager unos dok ne proveriš `VerificationUrl`, `InvoiceNumber`, `KooperantID` i `ClientRecordID` stavki.

Minimalni podaci koje operator mora da prikupi:

```text
KooperantID:
Korisnik/role:
InvoiceNumber:
Kompanija:
DatumRacuna:
VerificationUrl:
Ukupan iznos:
Broj stavki na računu:
Koje stavke su čekirane:
Za svaku stavku: NazivStavke, ArtikalID, ArtikalNaziv, Kolicina, JedCena, Ukupno
Da li je parse išao preko QR scan-a ili slike:
Da li je prikazana duplicate poruka:
Da li je saveFiskalni uspeo:
Da li je saveFiskalniMapiranje uspelo ili je bilo fire-and-forget:
```

---

## 2. Source of truth: gde se gleda

### 2.1. Prvo mesto: PWA fiskalni ekran

Na PWA strani proveri:

* da li je QR scan radio preko kamere;
* da li je fallback “Slikaj QR kod računa” korišćen;
* da li je parse vratio `duplicate`;
* da li je parse vratio `success`;
* da li su stavke prikazane;
* da li svaka čekirana stavka ima `ArtikalID`;
* da li je stavka mapirana automatski, ručno ili kao privatni artikal;
* da li je korisnik kliknuo “Prenesi u lager” / save.

### 2.2. Drugo mesto: GAS FISKALNI sheet

GAS definiše `FISKALNI_COLUMNS`:

```text
ClientRecordID
CreatedAtClient
SyncStatus
KooperantID
InvoiceNumber
Kompanija
DatumRacuna
VerificationUrl
NazivStavke
ArtikalID
ArtikalNaziv
Kolicina
JedCena
Ukupno
PDVStopa
Mapirano
ReceivedAt
```

Za incident proveri redove po:

* `VerificationUrl`;
* `InvoiceNumber`;
* `KooperantID`;
* `ClientRecordID`;
* `NazivStavke`.

### 2.3. Treće mesto: fiskalno mapiranje

GAS definiše `FISKALNI_MAP_COLUMNS`:

```text
FiskalniNaziv
ArtikalID
ArtikalNaziv
KooperantID
CreatedAt
```

Ovo je mapa koja pomaže da sledeći put fiskalni naziv bude automatski vezan na artikal.

Važno:

> `saveFiskalni` i `saveFiskalniMapiranje` su odvojeni. `saveFiskalni` može uspeti, a mapiranje da ne bude sačuvano, jer PWA mapiranje šalje “fire and forget”.

### 2.4. Četvrto mesto: PWA/GAS ErrorLog

Ako parse/save pukne, proveri `ErrorLog`:

```text
Timestamp
Source
Action
Message
Details
EntityID
Severity
```

Traži action:

```text
parseFiskalni
parseFiskalniImage
saveFiskalni
saveFiskalniMapiranje
```

---

## 3. Koji ID pratiš

Primarni incident ID-jevi:

```text
VerificationUrl      = najjači dedupe ključ računa
InvoiceNumber        = poslovni broj računa
KooperantID          = vlasnik lagera
ClientRecordID       = ID pojedinačne sačuvane stavke
ArtikalID            = master ili privatni artikal
FiskalniNaziv        = naziv sa računa, ulaz za mapiranje
```

Za svaku stavku prati:

```text
ClientRecordID:
NazivStavke:
ArtikalID:
ArtikalNaziv:
Kolicina:
JedCena:
Ukupno:
PDVStopa:
Mapirano:
```

Incident ticket minimum:

```text
KooperantID:
InvoiceNumber:
Kompanija:
DatumRacuna:
VerificationUrl:
Parse method: QR / image
Duplicate: Da/Ne
Save success: Da/Ne
Broj izabranih stavki:
ClientRecordID-jevi:
Problematicni ArtikalID:
Da li je mapiranje sačuvano:
Odluka:
```

---

## 4. Normalan tok

### 4.1. QR scan

1. Korisnik klikne scan.
2. Ako browser podržava `BarcodeDetector`, PWA otvara kameru.
3. Ako ne podržava, prikazuje fallback “Slikaj QR kod računa”.
4. Kada se QR očita, PWA poziva:

```text
action = parseFiskalni
kooperantID = CONFIG.ENTITY_ID
verificationUrl = url iz QR-a
```

5. GAS proverava token/role/entity.
6. GAS parse-uje fiskalni račun i vraća metadata + stavke.
7. Ako je račun već unet, vraća `duplicate`.

### 4.2. Image fallback

1. Korisnik slika QR kod.
2. PWA smanjuje sliku na max 1024 px i šalje base64.
3. PWA poziva:

```text
action = parseFiskalniImage
kooperantID = CONFIG.ENTITY_ID
imageBase64 = base64
```

4. GAS pokušava da pronađe QR u slici.
5. Ako uspe, vraća isti rezultat kao QR parse.

### 4.3. Mapiranje stavki

Za svaku fiskalnu stavku PWA može imati:

| Stanje           | Značenje                                   |
| ---------------- | ------------------------------------------ |
| `exact`          | automatski exact match na artikal          |
| `mapped`         | match iz prethodno naučenog mapiranja      |
| `fuzzy`          | približan match                            |
| `manual`         | korisnik ručno izabrao artikal             |
| `new-private`    | korisnik kreirao privatni `PRIV-*` artikal |
| nema `ArtikalID` | ne sme se čuvati ako je čekirano           |

### 4.4. Save u lager

1. Korisnik čekira stavke koje želi da prenese.
2. PWA validira da svaka čekirana stavka ima `ArtikalID`.
3. PWA za svaku izabranu stavku pravi `ClientRecordID`.
4. PWA poziva:

```text
action = saveFiskalni
kooperantID = CONFIG.ENTITY_ID
invoiceNumber
company
date
verificationUrl
stavke = selected[]
```

5. GAS čuva stavke u FISKALNI/lager sheet.
6. Ako je save uspešan, PWA šalje `saveFiskalniMapiranje` za nova ručna mapiranja, ali “fire and forget”.
7. PWA prikazuje: “N stavki preneseno u lager”.

---

## 5. Statusi i značenje

### 5.1. `duplicate = true` na parse-u

Značenje:

* GAS smatra da je račun već skeniran;
* najverovatnije postoji isti `VerificationUrl` ili drugi duplicate kriterijum.

Akcija:

* ne skenirati ponovo;
* proveriti FISKALNI sheet po `VerificationUrl`;
* proveriti da li su sve stavke računa već sačuvane;
* ako su sačuvane pod pogrešnim kooperantom, eskalirati.

### 5.2. `success = false` na parse-u

Značenje:

* QR nije fiskalni;
* fiskalni portal nije dostupan;
* image parse nije našao QR;
* GAS parser nije podržao format.

Akcija:

* pokušati direktan QR scan ako je image fallback pao;
* proveriti internet;
* proveriti ErrorLog;
* ako je format problem, tehnički owner.

### 5.3. Stavka bez `ArtikalID`

Značenje:

* stavka nije mapirana ni automatski ni ručno;
* PWA ne sme da je sačuva ako je čekirana.

Akcija:

* izabrati postojeći artikal;
* ili napraviti privatni artikal;
* ili odčekirati stavku ako ne treba u lager.

### 5.4. `PRIV-*` artikal

Značenje:

* korisnik je kreirao privatni artikal samo za ovaj fiskalni/lager zapis;
* ne upisuje se u globalni `Artikli` sheet;
* može biti vidljiv u fiskalnom/lager kontekstu, ali ne kao master šifarnik.

Akcija:

* ako artikal treba da postane globalni master artikal, Management mora da ga kreira kroz `createArtikal` ili master data proceduru.

### 5.5. `saveFiskalniMapiranje` nije uspelo

Značenje:

* račun/stavke mogu biti sačuvane;
* ali sledeći put isti fiskalni naziv neće biti automatski mapiran.

Akcija:

* ne duplirati račun;
* proveriti FISKALNI sheet da je lager unos uspeo;
* zatim ručno dodati mapiranje ili ponoviti mapping endpoint kao Management.

---

## 6. Standardni incident flow

### Korak 1: Identifikuj račun

Zapiši:

```text
KooperantID:
InvoiceNumber:
Kompanija:
DatumRacuna:
VerificationUrl:
Ukupan iznos:
```

Ako korisnik nema `VerificationUrl`, neka otvori fiskalni ekran ili pošalje screenshot QR/računa.

### Korak 2: Proveri da li račun postoji u FISKALNI sheet-u

Traži po:

```text
VerificationUrl
InvoiceNumber
KooperantID
Kompanija + DatumRacuna + Ukupno
```

Zapiši broj redova/stavki.

### Korak 3: Proveri stavke

Za svaku stavku proveri:

```text
ClientRecordID
NazivStavke
ArtikalID
ArtikalNaziv
Kolicina
JedCena
Ukupno
Mapirano
ReceivedAt
```

### Korak 4: Klasifikuj problem

| Signal                                 | Kategorija                   | Sledeći korak                              |
| -------------------------------------- | ---------------------------- | ------------------------------------------ |
| nema QR / kamera ne radi               | device/browser               | image fallback ili drugi uređaj            |
| image fallback ne nalazi QR            | image quality/parser         | bolja slika ili tehnički owner             |
| `duplicate = true`                     | već sačuvano                 | proveri FISKALNI po `VerificationUrl`      |
| parse uspeo, save nije kliknut         | korisnički tok               | korisnik mora izabrati stavke i sačuvati   |
| čekirana stavka bez `ArtikalID`        | mapiranje                    | mapirati ili odčekirati                    |
| saveFiskalni failed                    | backend/save problem         | ErrorLog + ne skenirati ponovo bez provere |
| save uspeo, mapiranje nije naučeno     | mapping fire-and-forget fail | dodati mapiranje kao Management            |
| pogrešan ArtikalID                     | pogrešno mapiranje           | poslovno/tehnički owner, korekcija lagera  |
| `PRIV-*` ne vidi se u master šifarniku | očekivano                    | kreirati globalni artikal ako treba        |

### Korak 5: Izaberi dozvoljenu akciju

| Stanje                                | Dozvoljena akcija                            |
| ------------------------------------- | -------------------------------------------- |
| Račun ne postoji nigde, parse ok      | ponoviti save                                |
| Račun postoji u FISKALNI sheet-u      | ne skenirati ponovo                          |
| Račun postoji, ali fali mapiranje     | popraviti mapiranje, ne lager unos           |
| Pogrešan artikal na stavci            | korekcija lagera + mapiranja uz owner-a      |
| Duplicate poruka, a FISKALNI nema red | proveriti ErrorLog i duplicate kriterijum    |
| Save nepoznatog ishoda                | prvo proveriti FISKALNI po `VerificationUrl` |

---

## 7. Retry pravila

### 7.1. Kada sme retry

Retry je dozvoljen ako:

* parse nije uspeo i nema FISKALNI redova za `VerificationUrl`;
* save nije uspeo i provereno je da nema redova za račun;
* korisnik nije kliknuo save;
* mapiranje nije uspelo, ali save jeste — retry samo mapiranje, ne račun;
* image fallback nije našao QR, pa se proba direktan QR scan.

### 7.2. Kada ne sme retry

Ne retry-ovati save računa ako:

* FISKALNI sheet već ima redove za `VerificationUrl`;
* parse vraća `duplicate = true`;
* nije jasno da li je prethodni `saveFiskalni` uspeo;
* korisnik je već video “N stavki preneseno u lager”;
* pogrešan je samo mapping, ne lager unos.

### 7.3. Kako se sprečava dupli račun

Zaštite:

1. Parse vraća `duplicate` ako je račun već skeniran.
2. `VerificationUrl` je najjači dedupe ključ.
3. Svaka sačuvana stavka ima `ClientRecordID`.
4. `saveFiskalni` ide pod GAS `withLock`.
5. Role/entity auth sprečava da kooperant upisuje za drugog kooperanta.
6. PWA ne čuva čekiranu stavku bez `ArtikalID`.

Operativno pravilo:

> Ako postoji isti `VerificationUrl` u FISKALNI sheet-u, ne čuvaj račun ponovo. Rešavaj mapiranje ili korekciju stavke, ne novi save.

---

## 8. Recovery scenariji

### 8.1. QR kamera ne radi

Postupak:

1. Proveri da li browser podržava kameru i `BarcodeDetector`.
2. Proveri dozvole za kameru.
3. Ako uređaj ne podržava scan, koristi “Slikaj QR kod računa”.
4. Ako ni slika ne radi, probati drugi uređaj ili ručno poslati screenshot tehničkom owner-u.

### 8.2. Image fallback ne pronalazi QR

Postupak:

1. Napraviti oštriju sliku QR-a.
2. Slikati samo QR zonu, bez refleksije i zamućenja.
3. Proveriti da li je QR fiskalni URL.
4. Ako više slika ne radi, tehnički owner proverava `parseFiskalniImage` i ErrorLog.

### 8.3. Parse kaže duplicate

Postupak:

1. Ne ponavljati save.
2. Pretražiti FISKALNI sheet po `VerificationUrl`.
3. Ako postoje redovi, proveriti da li broj stavki odgovara računu.
4. Ako fali stavka, proveriti da li je korisnik prethodno odčekirao stavku.
5. Ako su redovi pod pogrešnim `KooperantID`, eskalirati tehničkom + poslovnom owner-u.
6. Ako FISKALNI sheet nema red, proveriti duplicate kriterijum u GAS-u i ErrorLog.

### 8.4. Parse uspeo, ali stavka nema artikal

Postupak:

1. Ne klikati save dok je čekirana stavka bez `ArtikalID`.
2. Izabrati postojeći artikal iz dropdown-a.
3. Ako artikal ne postoji, korisnik može napraviti `PRIV-*` privatni artikal.
4. Ako artikal treba da bude master, Management kreira globalni artikal.
5. Tek posle toga save.

### 8.5. Save uspeo, ali mapiranje nije naučeno

Simptom:

* stavke su u lageru/FISKALNI sheet-u;
* sledeći put isti fiskalni naziv ne dobija auto-map.

Postupak:

1. Ne skenirati račun ponovo.
2. Proveriti `FISKALNI_MAP` sheet.
3. Ako nema mapiranja, Management dodaje mapiranje.
4. Proveriti da li `saveFiskalniMapiranje` zahteva Management token.
5. Ako Kooperant šalje mapping fire-and-forget bez Management prava, ovo je očekivani auth/design problem i treba ga popraviti.

### 8.6. Pogrešan artikal je sačuvan

Postupak:

1. Identifikuj račun po `VerificationUrl`.
2. Identifikuj stavku po `ClientRecordID` i `NazivStavke`.
3. Proveri pogrešni `ArtikalID` i ispravni `ArtikalID`.
4. Poslovni owner odlučuje korekciju lagera.
5. Tehnički owner koriguje FISKALNI/lager red ili radi storno/korekcioni red, zavisno od modela.
6. Popraviti `FISKALNI_MAP` da se greška ne ponavlja.

### 8.7. `PRIV-*` artikal nije u master artiklima

Postupak:

1. Objasniti da je `PRIV-*` lokalni/privatni artikal za taj fiskalni unos.
2. Ako treba da bude globalni artikal, Management kreira novi artikal u master šifarniku.
3. Nakon kreiranja, dodati fiskalno mapiranje `FiskalniNaziv → novi ArtikalID`.
4. Postojeći `PRIV-*` red ostaje audit trag ili se koriguje po odluci.

### 8.8. Račun je sačuvan pod pogrešnim kooperantom

Postupak:

1. Proveri token/entity iz ErrorLog-a ako postoji.
2. Proveri `KooperantID` u FISKALNI sheet-u.
3. Proveri da li je korisnik bio prijavljen kao pogrešan kooperant.
4. Ne premeštati red ručno bez poslovne odluke.
5. Tehnički owner i poslovni owner odlučuju korekciju: promena `KooperantID`, storno/korekcioni red ili ponovno knjiženje.

### 8.9. Save nepoznatog ishoda

Simptom:

* korisnik kliknuo save;
* mreža pukla ili PWA nije prikazala potvrdu;
* ne zna se da li je GAS upisao.

Postupak:

1. Ne skenirati ponovo odmah.
2. Proveriti FISKALNI sheet po `VerificationUrl`.
3. Ako postoje redovi, save je uspeo.
4. Ako nema redova, proveriti ErrorLog.
5. Ako nema server traga, retry save je dozvoljen.
6. Ako je neodređeno, tehnički owner proverava GAS log/Drive history.

---

## 9. Admin/GAS/DevTools provere

### 9.1. PWA lokalni state

U browser console proveriti:

```js
fiskalniMeta
fiskalniStavke
```

Zabeležiti:

```text
invoiceNumber
company
date
totalAmount
verificationUrl
stavke[].artikalID
stavke[].matchConfidence
```

### 9.2. Endpoint akcije

Relevantne GAS akcije:

```text
parseFiskalni
parseFiskalniImage
saveFiskalni
saveFiskalniMapiranje
createArtikal
```

### 9.3. Ručna provera FISKALNI sheet-a

Filter:

```text
VerificationUrl = <url>
KooperantID = <kooperant>
InvoiceNumber = <broj>
```

### 9.4. Ručna provera mapiranja

Filter `FISKALNI_MAP`:

```text
FiskalniNaziv = <naziv sa računa>
KooperantID = <kooperant>
ArtikalID = <artikal>
```

---

## 10. Ko donosi odluku

### Operator sme sam

* ponoviti QR scan ako parse nije uspeo i račun nije sačuvan;
* koristiti image fallback;
* izabrati postojeći artikal za nemapiranu stavku;
* odčekirati stavku koja ne treba u lager;
* proveriti duplicate poruku i prijaviti `VerificationUrl`.

### Kooperant sme sam

* izabrati artikal iz ponuđenog šifarnika;
* napraviti `PRIV-*` artikal za sopstvenu evidenciju;
* odlučiti koje stavke sa računa ulaze u lager, ako je to poslovno dozvoljeno.

### Management / poslovni owner odlučuje

* da li `PRIV-*` postaje globalni master artikal;
* ispravku pogrešno mapiranog artikla;
* korekciju lagera;
* šta raditi ako je račun pod pogrešnim kooperantom;
* da li se odčekirana stavka naknadno unosi.

### Tehnički owner odlučuje

* ručnu izmenu FISKALNI sheet-a;
* ručnu izmenu FISKALNI_MAP sheet-a;
* recovery posle save nepoznatog ishoda;
* popravku `parseFiskalni` / `parseFiskalniImage`;
* popravku auth/design problema oko `saveFiskalniMapiranje`;
* korekciju duplicate kriterijuma.

### Niko ne sme bez odobrenja

* brisati FISKALNI redove;
* menjati `VerificationUrl`;
* menjati `KooperantID` bez odluke;
* skenirati isti račun ponovo ako već postoji;
* menjati `ArtikalID` bez lager korekcije;
* tretirati `PRIV-*` kao master artikal.

---

## 11. Checklist za zatvaranje incidenta

```text
[ ] Identifikovan KooperantID
[ ] Identifikovan InvoiceNumber
[ ] Identifikovan VerificationUrl
[ ] Proveren FISKALNI sheet
[ ] Proveren broj sačuvanih stavki
[ ] Provereni ClientRecordID-jevi
[ ] Provereni ArtikalID-jevi
[ ] Ako je duplicate, potvrđeno gde je račun već sačuvan
[ ] Ako je pogrešno mapiranje, FISKALNI_MAP korigovan
[ ] Ako je PRIV artikal, odlučeno da li ostaje privatni ili postaje master
[ ] Ako je save nepoznat, potvrđeno da li postoji server red
[ ] Ako je ručna korekcija, postoji ticket/backup
[ ] Korisnik obavešten
```

---

## 12. Primeri odluke

### Primer A: Parse kaže “Ovaj račun je već skeniran”

Zaključak: račun verovatno već postoji.
Akcija: tražiti FISKALNI redove po `VerificationUrl`. Ne skenirati ponovo.

### Primer B: Stavka nema artikal

Zaključak: save je blokiran dok je čekirana stavka bez `ArtikalID`.
Akcija: mapirati na postojeći artikal, napraviti `PRIV-*`, ili odčekirati.

### Primer C: Račun je sačuvan, ali mapiranje nije naučeno

Zaključak: `saveFiskalni` uspeo, `saveFiskalniMapiranje` nije.
Akcija: ne dirati lager unos; dodati mapiranje kao Management.

### Primer D: Pogrešan artikal je ušao u lager

Zaključak: ledger/lager korekcija, ne samo mapiranje.
Akcija: poslovni owner odlučuje korekciju; tehnički owner ispravlja FISKALNI/lager i mapiranje.

### Primer E: `PRIV-*` nije u šifarniku artikala

Zaključak: očekivano ponašanje.
Akcija: ako treba globalni artikal, Management kreira master artikal i mapiranje.

### Primer F: Save pukao posle klika

Zaključak: nepoznat ishod.
Akcija: proveriti FISKALNI po `VerificationUrl`; retry samo ako nema server reda.

---

## 13. Poznate production rupe koje treba zatvoriti

1. `saveFiskalniMapiranje` trenutno zahteva Management, dok ga Kooperant PWA šalje fire-and-forget; odlučiti da li je to namerno ili bug.
2. Dodati eksplicitnu potvrdu/grešku za `saveFiskalniMapiranje`, umesto fire-and-forget bez vidljivosti.
3. Dodati `FiskalniEventLog`: parse, duplicate, save, mapping-save, correction, operator, razlog.
4. Dodati admin ekran “Find fiscal by VerificationUrl”.
5. Dodati korekcionu proceduru za pogrešno mapiran artikal koja automatski ažurira lager posledice.
6. Dodati jasnu UI oznaku da je `PRIV-*` privatan i nije master šifarnik.
7. Dodati server-side idempotency za `saveFiskalni` po `VerificationUrl + KooperantID + NazivStavke`.
8. Dodati report fiskalnih stavki bez master artikla ili sa `PRIV-*` starijih od X dana.
9. Dodati ErrorLog dashboard za `parseFiskalniImage` neuspehe.
10. Dodati test za duplicate račun i partial selected items.
11. Dodati mogućnost naknadnog dodavanja odčekirane stavke bez dupliranja celog računa.
12. Dodati jasnu politiku: ko sme da pravi globalne artikle iz fiskalnih stavki.

Do tada važi konzervativno pravilo:

> `VerificationUrl` je granica idempotency-ja fiskalnog računa. Ako taj URL već postoji, ne čuvaj račun ponovo; rešavaj mapiranje, stavke ili korekciju lagera.
