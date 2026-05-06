# Production runbook: Dokumentni chain, Zbirna–Prijemnica–Faktura i orphan recovery

Status: **operativni runbook za incidente gde se ne slažu kg/ambalaža, dokument je storniran, faktura nema stavke, prijemnica je već fakturisana ili je link otišao na pogrešnu zbirnu.**

Aplikacija: **OtkupApp / AgriX**
Domen: **Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF**
Glavni kod: `src-vba/modDokumenta.bas`, `src-vba/modFaktura.bas`, `src-vba/modBusinessFlowProTests.bas`, `src-vba/modNovac.bas`

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Ne mogu da napravim zbirnu.”
* “Ne slažu se kg sa otpremnicama.”
* “Ambalaža nije dobra.”
* “Prijemnica je uneta, ali ne mogu da je fakturišem.”
* “Faktura nema stavke.”
* “Prijemnica je već fakturisana, a ne vidim fakturu.”
* “Faktura je napravljena za pogrešnu prijemnicu.”
* “Otkup je povezan na pogrešnu otpremnicu.”
* “Jedna zbirna je pokupila otkupe iz druge zbirne.”
* “Stornirao sam dokument i sada lanac visi.”
* “Manjak nije tačan.”
* “Klasa II nije ušla / ambalaža je duplirana.”

Prvo pravilo:

> Ne popravljaj jedan red izolovano. Dokumentni lanac se proverava kroz ceo trace: `OtkupID → OtpremnicaID → BrojZbirne/ZbirnaID → PrijemnicaID → FakturaID → SEFDocumentId`.

Minimalni podaci koje operator mora da prikupi:

```text
BrojZbirne:
OtkupID / BrojDokumenta:
OtpremnicaID / BrojOtpremnice:
ZbirnaID:
PrijemnicaID / BrojPrijemnice:
FakturaID / BrojFakture:
KupacID:
VozacID:
StanicaID:
Klasa: I / II / I+II
Kolicina:
Cena:
TipAmbalaze:
KolAmbalaze:
KolAmbVracena:
Stornirano statusi:
Fakturisano status prijemnice:
SEFWorkflowState ako faktura već postoji:
```

---

## 2. Source of truth: gde se gleda

### 2.1. Prvo mesto: `BrojZbirne`

Za dokumentni chain incident, `BrojZbirne` je najvažniji poslovni ključ.

Njime povezuješ:

* otkupe koji treba da uđu u transport;
* otpremnice koje stanica predaje vozaču;
* zbirne koje vozač predaje kupcu/hladnjači;
* prijemnice koje kupac potvrđuje;
* kasnije fakture.

Pravilo:

> Ako auto-link ili ručna korekcija ne proverava `BrojZbirne`, postoji rizik cross-zbirna pogrešnog linka.

### 2.2. `tblOtkup`

Početni terenski/otkupni red.

Proveriti:

| Kolona                       | Značenje                                             |
| ---------------------------- | ---------------------------------------------------- |
| `OtkupID`                    | interni otkup ID                                     |
| `Datum`                      | datum otkupa                                         |
| `KooperantID`                | dobavljač/kooperant                                  |
| `StanicaID`                  | otkupno mesto                                        |
| `VrstaVoca`, `SortaVoca`     | roba                                                 |
| `Kolicina`, `Cena`           | otkupna količina/cena                                |
| `TipAmbalaze`, `KolAmbalaze` | ambalaža, obično samo Klasa I nosi količinu ambalaže |
| `VozacID`                    | vozač zadužen za otpremu                             |
| `BrojDokumenta`              | otkupni dokument/blok                                |
| `Klasa`                      | `I` ili `II`                                         |
| `BrojZbirne`                 | poslovni broj zbirne za transportni lanac            |
| `OtpremnicaID`               | link na otpremnicu                                   |
| `Stornirano`                 | ako postoji, isključuje red iz aktivnog toka         |

### 2.3. `tblOtpremnica`

Dokument: stanica predaje robu vozaču.

Proveriti:

| Kolona                       | Značenje                 |
| ---------------------------- | ------------------------ |
| `OtpremnicaID`               | interni ID               |
| `Datum`                      | datum otpreme            |
| `StanicaID`                  | stanica                  |
| `VozacID`                    | vozač                    |
| `BrojOtpremnice`             | poslovni broj otpremnice |
| `BrojZbirne`                 | zbirna kojoj pripada     |
| `VrstaVoca`, `SortaVoca`     | roba                     |
| `Kolicina`, `Cena`           | količina/cena            |
| `TipAmbalaze`, `KolAmbalaze` | ambalaža, obično Klasa I |
| `Klasa`                      | `I` / `II`               |
| `Stornirano`                 | ako postoji              |

### 2.4. `tblZbirna`

Dokument: vozač/kupac agregira isporuku.

Proveriti:

| Kolona                          | Značenje             |
| ------------------------------- | -------------------- |
| `ZbirnaID`                      | interni ID           |
| `Datum`                         | datum                |
| `VozacID`                       | vozač                |
| `BrojZbirne`                    | poslovni broj zbirne |
| `KupacID`                       | kupac/hladnjača      |
| `VrstaVoca`, `SortaVoca`        | roba                 |
| `UkupnoKolicina`                | ukupno kg za klasu   |
| `TipAmbalaze`, `UkupnoAmbalaze` | ambalaža             |
| `Klasa`                         | `I` / `II`           |
| `Stornirano`                    | ako postoji          |

### 2.5. `tblPrijemnica`

Dokument: kupac potvrđuje prijem i realnu količinu/cenu.

Proveriti:

| Kolona                       | Značenje                                |
| ---------------------------- | --------------------------------------- |
| `PrijemnicaID`               | interni ID                              |
| `Datum`                      | datum prijema                           |
| `KupacID`                    | kupac                                   |
| `VozacID`                    | vozač                                   |
| `BrojPrijemnice`             | poslovni broj prijemnice                |
| `BrojZbirne`                 | zbirna kojoj prijemnica pripada         |
| `VrstaVoca`, `SortaVoca`     | roba                                    |
| `Kolicina`, `Cena`           | canonical količina/cena za fakturisanje |
| `TipAmbalaze`, `KolAmbalaze` | ambalaža zadužena/primljena             |
| `KolAmbVracena`              | vraćena ambalaža                        |
| `Klasa`                      | `I` / `II`                              |
| `Fakturisano`                | `Da` ako je ušla u fakturu              |
| `FakturaID`                  | faktura kojoj pripada                   |
| `Stornirano`                 | ako postoji                             |

### 2.6. `tblFakture` i `tblFakturaStavke`

Faktura se pravi iz prijemnica, ne direktno iz otkupa ili zbirne.

`tblFakture` proveriti:

```text
FakturaID
BrojFakture
Datum
KupacID
Iznos
Status
DatumPlacanja
SEFWorkflowState
SEFDocumentId
Stornirano
```

`tblFakturaStavke` proveriti:

```text
StavkaID
FakturaID
PrijemnicaID
Kolicina
Cena
Klasa
BrojPrijemnice
Stornirano
```

Ključni invariant:

> Aktivna `PrijemnicaID` sme biti na najviše jednoj aktivnoj fakturi. Ako `Fakturisano = Da` ili `FakturaID` nije prazno, ne sme se ponovo fakturisati.

### 2.7. `tblAmbalaza`

Za incidente ambalaže proveriti pokrete:

```text
Datum
TipAmbalaze
Kolicina
Smer / Ulaz-Izlaz
EntitetID
EntitetTip
ReferenceID
DokumentTip
Stornirano
```

Ambalaža je često zadužena na Klasa I redu; Klasa II obično nosi `KolAmbalaze = 0` da se ambalaža ne duplira.

---

## 3. Koji ID pratiš

Uvek prati ceo set:

```text
BrojZbirne          = glavni poslovni chain key
OtkupID             = početni otkupni red
BrojDokumenta       = otkupni dokument/blok
OtpremnicaID        = transport od stanice ka vozaču
BrojOtpremnice      = poslovni broj otpremnice
ZbirnaID            = interni red zbirne
PrijemnicaID        = canonical izvor za fakturisanje
BrojPrijemnice      = poslovni broj prijemnice
FakturaID           = faktura
BrojFakture         = poslovni broj fakture
SEFDocumentId       = ako je faktura već poslata SEF-u
```

Incident ticket minimum:

```text
BrojZbirne:
OtkupID-i:
OtpremnicaID-i:
ZbirnaID-i:
PrijemnicaID-i:
FakturaID:
SEF state/document:
Razlika kg:
Razlika ambalaže:
Stornirano redovi:
Fakturisano redovi:
Odluka operator/poslovni/tehnički:
```

---

## 4. Normalan dokumentni tok

### 4.1. Otkup

`SaveOtkupMulti_TX` može napraviti dva reda:

* Klasa I;
* Klasa II.

Pravila:

* Klasa I nosi ambalažu;
* Klasa II obično nosi `KolAmbalaze = 0`;
* oba reda mogu imati isti `BrojDokumenta` i isti `BrojZbirne`;
* transakcija treba da appenduje oba reda ili nijedan.

### 4.2. Otpremnica

`SaveOtpremnicaMulti_TX` / `SaveOtpremnica_TX` pravi otpremnicu po klasi.

Pravila:

* `BrojOtpremnice` povezuje fizički dokument;
* `BrojZbirne` povezuje transportni chain;
* ambalaža se trackuje kroz `TrackAmbalaza` kada je `KolAmbalaze > 0`;
* Klasa II ne sme duplirati ambalažu.

### 4.3. Zbirna

`SaveZbirnaMulti_TX` / `SaveZbirna_TX` pravi zbirnu po klasi.

Pre unosa treba validirati:

```vb
ValidateZbirnaPreUnosa(brojZbirne, inputKgKlI, inputKgKlII, inputAmb)
```

Posle unosa proveriti:

```vb
ValidateZbirna(brojZbirne)
```

Pravila:

* zbirna kg treba da se složi sa otpremnicama po `BrojZbirne`;
* razlika može biti poslovna realnost samo ako je tako odobreno;
* ambalaža mora biti pod kontrolom, jer dupliranje odmah remeti saldo.

### 4.4. Prijemnica

`SavePrijemnicaMulti_TX` / `SavePrijemnica_TX` pravi prijemnicu po klasi.

Pravila:

* prijemnica je canonical izvor za fakturu;
* `Kolicina × Cena` iz prijemnice ulazi u fakturu;
* prijemnica se markira kao `Fakturisano = Da` tek kroz `CreateFaktura`;
* `KolAmbVracena` je važan za ambalažu i manjak.

### 4.5. Faktura

`CreateFaktura_TX(kupacID, stavke)`:

* pre-validira sve prijemnice pre bilo kog upisa;
* blokira duplu prijemnicu u izboru;
* blokira prijemnicu koja je već fakturisana;
* blokira storniranu prijemnicu;
* računa iznos iz canonical `tblPrijemnica` vrednosti;
* pravi `tblFakture` header;
* pravi `tblFakturaStavke`;
* markira prijemnice `Fakturisano = Da` i popunjava `FakturaID`;
* automatski primenjuje kupčev avans kroz `ApplyAvansToFaktura`.

Pravilo:

> Kada faktura postoji, prijemnica više nije slobodna. Korekcije posle fakture su finansijsko/SEF pitanje, ne običan edit dokumentnog chain-a.

---

## 5. Statusi i invarianti

### 5.1. Aktivni vs stornirani redovi

Većina read helper-a koristi `ExcludeStornirano`. To znači:

* stornirani red postoji kao audit trag;
* ne učestvuje u aktivnim zbirnim/prijemnicama/fakturama;
* ne sme se fizički brisati bez posebne odluke.

### 5.2. Prijemnica availability za fakturu

Prijemnica je dostupna za fakturisanje samo ako:

```text
Stornirano nije Da
Fakturisano nije Da
FakturaID je prazno
Kolicina je numerička i > 0
Cena je numerička i >= 0
```

### 5.3. Faktura stavke

Faktura mora imati bar jednu stavku. Ako `tblFakture` ima header, ali `tblFakturaStavke` nema redove, to je nekonzistentno stanje i ne sme se slati na SEF.

### 5.4. Cross-zbirna link

Ne sme se desiti da se `OtkupID` iz jedne `BrojZbirne` poveže sa `OtpremnicaID` iz druge `BrojZbirne`.

Ako se desi, to je P0 data integrity incident.

---

## 6. Standardni incident flow

### Korak 1: Identifikuj `BrojZbirne`

Ako korisnik nema `BrojZbirne`, nađi ga po:

* vozaču;
* datumu;
* kupcu;
* otkupnom dokumentu;
* prijemnici;
* fakturi.

Zapiši:

```text
BrojZbirne:
Datum:
VozacID:
KupacID:
StanicaID-i:
Vrsta/Sorta:
```

### Korak 2: Izvuci sve redove iz chain-a

Za isti `BrojZbirne` izvuci:

```text
tblOtkup redove
tblOtpremnica redove
tblZbirna redove
tblPrijemnica redove
tblFakturaStavke preko PrijemnicaID
tblFakture preko FakturaID
tblAmbalaza pokrete preko reference ID-jeva
```

### Korak 3: Proveri klase

Za svaku fazu proveri:

```text
Klasa I postoji?
Klasa II postoji ako očekivano?
Da li Klasa II ima KolAmbalaze = 0?
Da li ukupno kg odgovara očekivanju?
Da li cena dolazi iz ispravnog canonical izvora?
```

### Korak 4: Proveri stornirano/fakturisano

Zapiši:

```text
Otkup stornirano:
Otpremnica stornirano:
Zbirna stornirano:
Prijemnica stornirano:
Prijemnica Fakturisano:
Prijemnica FakturaID:
Faktura stornirano:
Faktura SEFWorkflowState:
Faktura SEFDocumentId:
```

### Korak 5: Klasifikuj problem

| Signal                                          | Kategorija             | Sledeći korak                     |
| ----------------------------------------------- | ---------------------- | --------------------------------- |
| Zbirna kg ne odgovara otpremnicama              | količinski mismatch    | validacija + poslovna odluka      |
| Prijemnica kg manji od zbirne kg                | manjak                 | izračunati i poslovno potvrditi   |
| Ambalaža duplirana u Klasa II                   | ambalaža bug/korekcija | tehnički + poslovni owner         |
| Prijemnica `Fakturisano = Da`, ali nema fakture | nekonzistentno         | recovery kroz faktura stavke/log  |
| Faktura header postoji bez stavki               | nekonzistentno         | ne slati SEF, tehnički recovery   |
| Faktura poslata SEF-u                           | pravno/SEF ograničenje | SEF runbook, ne editovati lokalno |
| Otkup linkovan na pogrešnu otpremnicu           | traceability incident  | stop auto-link, audit             |
| Cross-zbirna link                               | P0 data integrity      | zaustaviti chain, tehnički audit  |
| Stornirano srednji dokument                     | orphan chain           | procena posledica downstream      |

---

## 7. Dozvoljene akcije po stanju

| Stanje                                             | Dozvoljena akcija                                  |
| -------------------------------------------------- | -------------------------------------------------- |
| Još nema fakture                                   | korekcija dokumentnog chain-a uz poslovnu odluku   |
| Prijemnica nije fakturisana                        | može se korigovati/stornirati po proceduri         |
| Prijemnica fakturisana, faktura nije poslata SEF-u | korekcija samo uz finansijski/tehnički owner       |
| Faktura poslata SEF-u i ima `SEFDocumentId`        | ne editovati; SEF cancel/storno flow               |
| Cross-zbirna link                                  | stop operacije, audit, ne nastavljati fakturisanje |
| Faktura header bez stavki                          | zabranjen SEF send; tehnički recovery              |
| `BrojZbirne` mismatch                              | ne auto-linkovati; ručna analiza                   |

---

## 8. Recovery scenariji

### 8.1. Ne slažu se kg između otpremnica i zbirne

Postupak:

1. Izvuci sve `tblOtpremnica` redove za `BrojZbirne`.
2. Izvuci sve `tblZbirna` redove za isti `BrojZbirne`.
3. Pozovi ili ručno izračunaj `ValidateZbirna(brojZbirne)`.
4. Razdvoj Klasa I i Klasa II.
5. Ako je razlika greška unosa, poslovni owner odlučuje koji dokument se koriguje/stornira.
6. Ako je razlika realan manjak, dokumentovati kao manjak, ne “popravljati” količine da se slažu.
7. Ne praviti fakturu dok nije jasno šta je canonical prijemna količina.

### 8.2. Prijemnica pokazuje manjak u odnosu na zbirnu

Postupak:

1. Uporedi `tblZbirna.UkupnoKolicina` i `tblPrijemnica.Kolicina` po `BrojZbirne` i klasi.
2. Izračunaj razliku.
3. Proveri da li postoji poslovno odobren manjak.
4. Ako je manjak realan, faktura se pravi iz prijemnice, ne iz zbirne.
5. Ako je greška unosa, stornirati/korigovati prijemnicu pre fakturisanja.

### 8.3. Ambalaža se ne slaže

Postupak:

1. Proveri `KolAmbalaze` u `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`.
2. Proveri `KolAmbVracena` u prijemnici.
3. Proveri `tblAmbalaza` pokrete po referencama.
4. Proveri da li je Klasa II slučajno dobila ambalažu.
5. Ako je ambalaža duplirana, ne popravljati samo saldo; pronaći dokument koji je napravio dupli pokret.
6. Poslovni owner odlučuje korekciju zaduženja/vraćanja ambalaže.

### 8.4. Prijemnica je `Fakturisano = Da`, ali korisnik ne vidi fakturu

Postupak:

1. Uzmi `PrijemnicaID`.
2. Proveri `tblPrijemnica.FakturaID`.
3. Ako postoji `FakturaID`, proveri `tblFakture`.
4. Proveri `tblFakturaStavke` po `PrijemnicaID`.
5. Ako faktura postoji, problem je UI/search/report.
6. Ako `FakturaID` pokazuje na nepostojeću fakturu, to je nekonzistentno stanje; tehnički owner radi recovery iz backup/journal/log.
7. Ne čistiti `Fakturisano` ručno bez provere da faktura/stavka stvarno ne postoji.

### 8.5. Faktura header postoji, ali nema stavke

Ovo je P0 za fakturisanje/SEF.

Postupak:

1. Ne štampati i ne slati na SEF.
2. Proveri `tblFakturaStavke` po `FakturaID`.
3. Proveri `tblPrijemnica` koje su možda markirane na taj `FakturaID`.
4. Proveri `Journal/` i backup.
5. Ako je `CreateFaktura_TX` rollback trebalo da vrati sve, proveriti zašto je header ostao.
6. Tehnički owner odlučuje da li se faktura stornira/uklanja iz aktivnog toka ili rekonstruiše iz prijemnica.
7. Ako je faktura već dobila `SEFDocumentId`, ide SEF/legal flow.

### 8.6. Pokušaj duple fakture za istu prijemnicu

Simptom:

* `CreateFaktura_TX` failuje;
* poruka da je prijemnica već fakturisana ili stornirana.

Postupak:

1. Proveri `PrijemnicaID`.
2. Proveri `Fakturisano` i `FakturaID`.
3. Ako faktura postoji, ne praviti novu.
4. Ako faktura ne postoji, proveriti `tblFakturaStavke` i backup.
5. Tehnički owner odlučuje recovery.

### 8.7. Otkup je linkovan na pogrešnu otpremnicu

Postupak:

1. Uzmi `OtkupID` i njegov `BrojZbirne`.
2. Uzmi `OtpremnicaID` iz `tblOtkup`.
3. Proveri `tblOtpremnica.BrojZbirne`.
4. Ako se `BrojZbirne` ne poklapa, ovo je pogrešan link.
5. Zaustaviti automatski auto-link proces dok se ne uradi audit.
6. Tehnički owner ispravlja link uz backup.
7. Pokrenuti audit za sve redove gde `tblOtkup.BrojZbirne <> tblOtpremnica.BrojZbirne`.

### 8.8. Cross-zbirna link incident

Simptom:

* otkup iz jedne zbirne je vezan za otpremnicu druge zbirne;
* regression test / audit ukazuje da auto-link može pogrešiti ako `BrojZbirne` nije deo ključa.

Postupak:

1. Zaustaviti dalje fakturisanje pogođenih zbirnih.
2. Izvući sve `OtkupID` i `OtpremnicaID` za oba `BrojZbirne` broja.
3. Napraviti tabelu:

```text
OtkupID | Otkup.BrojZbirne | OtpremnicaID | Otpremnica.BrojZbirne | Klasa | Kg | Stanica | Vozac
```

4. Sve mismatch redove označiti kao data integrity incident.
5. Tehnički owner popravlja linkove.
6. Poslovni owner potvrđuje koji dokumentni chain je stvarno važeći.
7. Tek posle toga nastaviti prijemnice/fakture.

### 8.9. Storniran je srednji dokument

Primer: stornirana otpremnica, ali zbirna/prijemnica/faktura postoje.

Postupak:

1. Identifikuj downstream dokumente.
2. Ako nema fakture, poslovni owner odlučuje da li se pravi nova otpremnica/zbirna/prijemnica ili se postojeći chain stornira.
3. Ako postoji faktura, proveri da li je poslata SEF-u.
4. Ako postoji `SEFDocumentId`, ne dirati lokalno bez SEF/legal flow-a.
5. Ako faktura nije poslata SEF-u, tehnički + finansijski owner odlučuju korekciju/storno lokalne fakture.

### 8.10. Faktura već poslata SEF-u, a chain je pogrešan

Postupak:

1. Ne editovati prijemnice/faktura stavke direktno.
2. Preći na SEF production runbook.
3. Proveriti `SEFWorkflowState` i `SEFDocumentId`.
4. Pravni/računovodstveni owner odlučuje cancel/storno.
5. Tek nakon pravno ispravnog SEF ishoda rešavati lokalni chain.

---

## 9. Kako sprečavaš duple i pogrešne dokumente

Sistem ima zaštite:

1. TX wrapper-i za multi-row dokumente: sve ili ništa.
2. `ValidateZbirnaPreUnosa` i `ValidateZbirna` proveravaju kg/ambalažu oko zbirne.
3. `CreateFaktura` pre-validira sve prijemnice pre append-a fakture.
4. Faktura se računa iz `tblPrijemnica`, ne iz user-prosleđenih stavki.
5. Dupla prijemnica u izboru fakture se blokira.
6. Već fakturisana ili stornirana prijemnica se blokira.
7. `ExcludeStornirano` čuva stornirane redove van aktivnih helper-a.
8. Regression suite testira happy path, invalid saves, duplicate faktura block i traceability.

Operativno pravilo:

> Ako link ne možeš dokazati preko `BrojZbirne` i ID lanca, ne pravi sledeći dokument.

---

## 10. Admin/VBA komande

Koristiti samo ako UI nije dovoljan ili ako tehnički owner radi incident.

```vb
' Validacija zbirne pre unosa
Debug.Print Join(ValidateZbirnaPreUnosa("ZBR-...", 1000, 200, 100), " | ")

' Validacija postojeće zbirne
Debug.Print Join(ValidateZbirna("ZBR-..."), " | ")

' Kreiranje otpremnice
Debug.Print SaveOtpremnica_TX(Date, "ST-...", "VOZ-...", "OTP-...", "ZBR-...", _
                              "Jabuka", "Ajdared", 1000, 120, "Gajba", 100, "I")

' Kreiranje zbirne
Debug.Print SaveZbirna_TX(Date, "VOZ-...", "ZBR-...", "KUP-...", _
                          "Hladnjaca", "Pogon", "Jabuka", "Ajdared", _
                          1000, "Gajba", 100, "I")

' Kreiranje prijemnice
Debug.Print SavePrijemnica_TX(Date, "KUP-...", "VOZ-...", "PRJ-...", "ZBR-...", _
                              "Jabuka", "Ajdared", 990, 120, "Gajba", 100, 95, "I")

' Kreiranje fakture iz prijemnica
Dim stavke As New Collection
stavke.Add Array("PRJ-00001")
Debug.Print CreateFaktura_TX("KUP-...", stavke)

' Recalculate status fakture posle finansijske korekcije
Call UpdateFakturaStatus("FAK-...")

' Regression / audit suite
Call RunBusinessFlowProSuite
Call RunBusinessFlowProTraceabilityOnly
Call RunBusinessFlowProAuditOnly
```

Napomena: direktno ručno menjanje linkova i `Stornirano` kolona nije standardna operator akcija. To je tehnički recovery uz backup/ticket.

---

## 11. Ko donosi odluku

### Operator sme sam

* pokrenuti validaciju zbirne;
* pronaći dokumente po `BrojZbirne`;
* proveriti `Fakturisano`, `Stornirano`, `FakturaID`;
* odbiti fakturisanje ako prijemnica nije dostupna;
* eskalirati mismatch sa kompletnim ID lancem.

### Tehnički owner odlučuje

* ručnu korekciju linkova `OtkupID → OtpremnicaID`;
* recovery fakture header bez stavki;
* čišćenje `Fakturisano/FakturaID` nekonzistentnog stanja;
* popravku auto-link logike;
* audit cross-zbirna linkova;
* intervencije kroz backup/journal.

### Poslovni / logistički owner odlučuje

* da li je razlika kg realan manjak ili greška;
* koja količina je canonical za prijem/fakturu;
* da li se dokument stornira ili pravi novi;
* šta raditi sa ambalažom i vraćanjem;
* koji chain je ispravan ako postoje dva slična.

### Finansijski / pravni owner odlučuje

* korekcije posle fakture;
* storno fakture;
* izmene kada je faktura poslata na SEF;
* da li se kupcu šalje novi dokument ili storno.

### Niko ne sme bez odobrenja

* fizički brisati dokumentne redove;
* menjati `BrojZbirne` na jednom redu bez celog chain audit-a;
* čistiti `Fakturisano` ručno da bi se napravila nova faktura;
* menjati faktura stavke ako faktura ima `SEFDocumentId`;
* nastaviti fakturisanje posle cross-zbirna mismatch-a.

---

## 12. Checklist za zatvaranje incidenta

```text
[ ] Identifikovan BrojZbirne
[ ] Izvučeni svi tblOtkup redovi
[ ] Izvučeni svi tblOtpremnica redovi
[ ] Izvučeni svi tblZbirna redovi
[ ] Izvučeni svi tblPrijemnica redovi
[ ] Proverene tblFakturaStavke
[ ] Proverena tblFakture
[ ] Proveren SEF status ako faktura postoji
[ ] Proverene klase I/II
[ ] Proverena kg razlika
[ ] Proverena ambalaža i KolAmbVracena
[ ] Provereni Stornirano flagovi
[ ] Proveren Fakturisano/FakturaID status
[ ] Ako je link korekcija, postoji tehnički owner odobrenje
[ ] Ako je manjak/ambalaža sporna, postoji poslovna odluka
[ ] Ako je faktura/SEF sporno, postoji finansijsko-pravna odluka
[ ] Posle korekcije ponovo validirana zbirna/chain
[ ] Korisnik obavešten
```

---

## 13. Primeri odluke

### Primer A: Zbirna ima 1200 kg, prijemnica 1180 kg

Zaključak: postoji manjak 20 kg ili greška prijema.
Akcija: poslovni owner odlučuje. Ako je realan manjak, faktura ide iz 1180 kg prijemnice. Ako je greška, korigovati/stornirati prijemnicu pre fakture.

### Primer B: Prijemnica je `Fakturisano = Da`, ali faktura nije vidljiva

Zaključak: mora se naći `FakturaID` ili `tblFakturaStavke`.
Akcija: ne čistiti status. Proveriti `FakturaID`, stavke, backup/journal.

### Primer C: Faktura header postoji bez stavki

Zaključak: nekonzistentna faktura, SEF stop.
Akcija: ne slati SEF. Tehnički owner rekonstruiše ili stornira lokalno uz ticket.

### Primer D: Otkup je linkovan na otpremnicu druge zbirne

Zaključak: cross-zbirna data integrity incident.
Akcija: zaustaviti dalji chain, audit svih linkova, popravka uz tehničkog owner-a.

### Primer E: Klasa II ima ambalažu

Zaključak: verovatna duplirana ambalaža.
Akcija: proveriti `tblAmbalaza` pokrete i dokument koji je napravio duplu količinu; poslovni owner odlučuje korekciju.

### Primer F: Prijemnica već fakturisana, korisnik hoće novu fakturu

Zaključak: zabranjeno bez storno/korekcije.
Akcija: naći postojeću fakturu. Ako je pogrešna, finansijski/pravni owner odlučuje storno/cancel/novi dokument.

---

## 14. Poznate production rupe koje treba zatvoriti

1. Dodati `tblDocumentChainEventLog`: svaki link/unlink/storno/korekcija sa operatorom i razlogom.
2. Dodati admin ekran `TraceByZbirna` koji prikazuje ceo chain u jednoj tabeli.
3. Dodati hard audit: `tblOtkup.BrojZbirne` mora biti jednak `tblOtpremnica.BrojZbirne` za svaki aktivni link.
4. Popraviti/zaključati auto-link da `BrojZbirne` bude deo preferiranog ključa.
5. Dodati “orphan dashboard”: prijemnice bez fakture, fakture bez stavki, otkupi bez otpremnice, stornirani srednji dokumenti sa downstream dokumentima.
6. Dodati eksplicitnu Undo/Correct proceduru za link `OtkupID → OtpremnicaID`.
7. Dodati eksplicitnu Storno chain proceduru koja prikazuje downstream posledice pre izvršenja.
8. Dodati blokadu SEF slanja ako faktura nema stavke ili ako bilo koja faktura stavka pokazuje na storniranu prijemnicu.
9. Dodati formalnu politiku za manjak: ko odobrava, kako se beleži, kako utiče na fakturu.
10. Dodati formalnu politiku za ambalažu: ko odobrava korekcije salda.
11. Dodati automatski dnevni audit cross-zbirna linkova.
12. Dodati test koji mora proći pre release-a: duplicate faktura block, cross-zbirna no-link, faktura iz prijemnica, Klasa II bez ambalaže.

Do tada važi konzervativno pravilo:

> `BrojZbirne` je granica dokumentnog chain-a. Ako dokument ili link prelazi tu granicu bez eksplicitne poslovne odluke, zaustavi fakturisanje i uradi audit.
