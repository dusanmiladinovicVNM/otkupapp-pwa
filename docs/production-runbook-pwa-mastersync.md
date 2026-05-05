# Production runbook: PWA MasterSync OTK/VOZ import i writeback

Status: **operativni runbook za incidente “PWA unos postoji, ali nije u Excel masteru” i “Vozač je napravio zbirnu, ali nema/pogrešan je BrojZbirne.”**

Aplikacija: **OtkupApp / AgriX**
Domen: **PWA → GAS/Google Sheets → Excel/VBA master sync**
Glavni kod: `src-vba/modMasterSync.bas`, `gas/Code.gs`, PWA IndexedDB sync engine
Glavne tabele/sheetovi: `OTK-*`, `VOZ-*`, `tblOtkup`, `tblAmbalaza`, `tblZbirna`, `tblOtpremnica`

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Unos sa telefona se vidi u PWA, ali ga nema u Excelu.”
* “Otkup je sinhronizovan, ali master import ga nije povukao.”
* “U Google Sheet-u piše `SyncError`.”
* “U Google Sheet-u piše `Synced`, ali Excel ga ne uvozi.”
* “Vozač je napravio zbirnu, ali nema `BrojZbirne`.”
* “U VOZ sheet-u u koloni T stoji interni `ZBR-*` umesto poslovnog broja zbirne.”
* “Red je dupliran / označen kao `Duplicate`.”
* “Import kaže uspešno, ali writeback nije upisan u Google.”

Prvo pravilo:

> Ne popravljaj ručno status u Google Sheet-u pre nego što identifikuješ `ClientRecordID`, sheet, lokalni master red i writeback stanje.

Minimalni podaci koje operator mora da prikupi:

```text
Uloga korisnika: Otkupac / Vozac
Korisnik / entityID: OtkupacID ili VozacID
Vreme unosa u PWA:
Tip unosa: OTK otkup ili VOZ zbirna
ClientRecordID:
ServerRecordID:
Google sheet ime: OTK-* ili VOZ-*
Google row number:
SyncStatus:
Ako je VOZ: BrojZbirne:
Ako je OTK: KooperantID, Datum, Klasa, Kolicina:
Screenshot iz PWA ako postoji:
```

---

## 2. Source of truth: gde se gleda

### 2.1. Prvo mesto: Google Sheet red

Za PWA → Excel sync incident, prvo se gleda odgovarajući Google Sheet:

* `OTK-<OtkupacID>` za otkupe;
* `VOZ-<VozacID>` za vozačke zbirne.

Uvek pronađi red po `ClientRecordID`. Ako korisnik ne zna `ClientRecordID`, traži po:

* `Datum`;
* `KooperantID` ili `KupacID`;
* `Kolicina` / `KolicinaKlI` / `KolicinaKlII`;
* `CreatedAtClient`;
* `ReceivedAt`;
* korisnik/stanica/vozač.

### 2.2. Drugo mesto: Excel master tabele

Za OTK:

* `tblOtkup`
* `tblAmbalaza`
* eventualno kasnije `tblOtpremnica`

Za VOZ:

* `tblZbirna`
* `tblOtkup`, ako se nakon import-a linkuje `BrojZbirne` na otkup redove
* `tblOtpremnica`, ako se linkuje zbirna na transportne dokumente

### 2.3. Treće mesto: lokalni logovi

Proveriti:

* `Log/` dnevni log;
* `Journal/` CSV journal, ako postoji sumnja na crash;
* eventualni Google API log/error iz `LogError`, `LogWarn`, `LogInfo` poziva;
* `ErrorLog` u Google Drive-u samo ako je problem došao iz PWA/GAS sync-a pre master import-a.

### 2.4. Četvrto mesto: PWA lokalni uređaj

Ako red nije stigao do Google Sheet-a, problem nije MasterSync nego PWA local sync.

Tada se incident prebacuje na PWA offline sync runbook i proveravaju se:

* IndexedDB store: `otkupi` ili `zbirne`;
* `syncStatus`;
* `lastServerStatus`;
* `lastSyncError`;
* `clientRecordID`;
* `serverRecordID`.

---

## 3. Koji ID pratiš

### 3.1. OTK incident

Za OTK red obavezno prati:

```text
ClientRecordID      = primarni PWA lokalni ID i dedupe key
ServerRecordID      = GAS tehnički server ID / kasnije Excel writeback vrednost
OtkupacID           = vlasnik OTK sheet-a
Google sheet        = OTK-<OtkupacID>
Google row number   = fizički red u Sheet1
SyncStatus          = Synced / Synced>Master / Duplicate / SyncError...
OtkupID             = Excel master ID, OTK-xxxxx
Datum
KooperantID
Stanica/OtkupacID
Klasa
Kolicina
```

### 3.2. VOZ incident

Za VOZ red obavezno prati:

```text
ClientRecordID      = primarni PWA lokalni ID i dedupe key
ServerRecordID      = tehnički sync/master ID; u writeback-u treba da bude ZbirnaID
VozacID             = vlasnik VOZ sheet-a
Google sheet        = VOZ-<VozacID>
Google row number   = fizički red u Sheet1
SyncStatus          = Synced / Synced>Master / Duplicate / SyncError...
ZbirnaID            = Excel master ID, ZBR-xxxxx
BrojZbirne          = poslovni broj zbirne, npr. 4/040526 ili 4/040526-2
OtkupRecordIDs      = lista PWA otkup redova koje zbirna pokriva
Datum
KupacID
KolicinaKlI
KolicinaKlII
```

Najvažnije pravilo za VOZ:

> `ServerRecordID` nije isto što i `BrojZbirne`. `ServerRecordID` je tehnički ID. `BrojZbirne` je poslovni dokumentni broj koji generiše VBA Master import.

---

## 4. Statusi u Google Sheet-u

### 4.1. `SyncStatus = Synced`

Značenje:

* PWA je uspešno poslala red u GAS;
* GAS je upisao/confirmovao red u Google Sheet;
* Excel MasterSync još treba da ga uveze.

Akcija:

* Dozvoljeno je pokrenuti odgovarajući MasterSync import.
* Ne menjati ručno status pre import-a.

### 4.2. `SyncStatus = Synced>Master`

Značenje:

* Excel MasterSync je obradio red;
* red treba da postoji u master tabeli;
* Google writeback je upisao da je red u masteru.

Akcija:

* Ako korisnik tvrdi da ga nema u Excelu, proveri po `ClientRecordID` / `ServerRecordID` / poslovnim poljima.
* Ako nema master reda, proveri rollback/crash/journal i log.
* Ne vraćati automatski na `Synced` bez analize, jer može napraviti dupli import.

### 4.3. `SyncStatus = Duplicate`

Značenje:

* MasterSync smatra da red već postoji ili je poslovno duplikat.

Akcija:

* Pronađi postojeći master red.
* Uporedi `ClientRecordID`, datum, kooperanta/kupca, količinu, klasu i vozača.
* Ako je stvarno isti unos, incident se zatvara kao duplikat.
* Ako nije isti unos, eskalirati tehničkom owner-u; ne menjati status u `Synced` bez dokumentovane odluke.

### 4.4. `SyncStatus = SyncError` ili `SyncError:<reason>`

Značenje:

* MasterSync je pokušao ili pripremio obradu, ali red nije validan za master.

Tipični uzroci:

* schema drift Google Sheet header-a;
* nedostaje `ClientRecordID`;
* nedostaje obavezno poslovno polje;
* količina/cena nisu validne;
* kooperant/kupac/vozač ne postoji u masteru;
* writeback nije uspeo;
* fatal Google API error.

Akcija:

* Ne brisati red.
* Ne duplirati ručno unos.
* Zabeležiti `SyncError` razlog i `ClientRecordID`.
* Ispravka zavisi od tipa greške: master data, schema, validacija, Google auth ili writeback.

### 4.5. Prazan / nepoznat `SyncStatus`

Značenje:

* Red nije u normalnom master import contract-u.

Akcija:

* Ako je red ručno dodat u Google Sheet, tretirati kao sumnjiv.
* Ne uvoziti ručno bez tehničkog pregleda.
* Ako je PWA red, proveriti PWA sync engine i GAS response.

---

## 5. OTK flow: kako treba da radi

Normalan tok:

1. Otkupac unese otkup u PWA.
2. PWA snimi lokalno u IndexedDB store `otkupi` sa `clientRecordID`.
3. PWA sync pošalje red GAS-u kroz action `sync`.
4. GAS validira token/ulogu/entity ownership i upisuje u `OTK-<OtkupacID>`.
5. GAS vraća batch response sa `serverRecordID` i statusom.
6. PWA lokalni red postaje `synced`.
7. Excel operator pokrene `ImportOtkupFromPWA_TX`.
8. VBA traži `OTK-*` sheetove u Google PWA folderu.
9. Za svaki sheet čita `Sheet1`.
10. Proverava header contract.
11. Uvozi samo redove sa `SyncStatus = Synced`.
12. Kreira red u `tblOtkup`, i po potrebi ambalažu u `tblAmbalaza`.
13. Piše writeback u Google Sheet:

    * `SyncStatus = Synced>Master`;
    * `ServerRecordID = OtkupID` ili relevantni master ID.
14. Transaction se commit-uje.

---

## 6. VOZ flow: kako treba da radi

Normalan tok:

1. Vozač u PWA napravi zbirnu.
2. PWA snimi lokalno u IndexedDB store `zbirne` sa `clientRecordID`.
3. PWA sync pošalje red GAS-u kroz action `syncZbirna`.
4. GAS validira token/ulogu/entity ownership i upisuje u `VOZ-<VozacID>`.
5. Google Sheet red dobija `SyncStatus = Synced`.
6. Excel operator pokrene VOZ/Zbirna master import.
7. VBA čita `VOZ-*` sheetove.
8. Uvozi samo `SyncStatus = Synced`.
9. Kreira `tblZbirna` red i generiše poslovni `BrojZbirne` kroz VBA Master import.
10. Linkuje `BrojZbirne` nazad na povezane `tblOtkup` / `tblOtpremnica` redove kada je moguće.
11. Piše writeback u Google Sheet:

    * kolona B / `ServerRecordID` = `ZbirnaID`;
    * kolona F / `SyncStatus` = `Synced>Master`;
    * kolona T / `BrojZbirne` = poslovni broj zbirne.

Hard rule:

> Ako u VOZ sheet-u kolona T / `BrojZbirne` sadrži `ZBR-*`, writeback je pogrešan. Kolona T mora sadržati poslovni broj zbirne, ne tehnički master ID.

---

## 7. Standardni incident flow

### Korak 1: Odredi domen

Pitanje:

```text
Da li problem ima PWA red u Google Sheet-u?
```

* Ako **nema** red u Google Sheet-u: ovo je PWA/GAS sync incident, ne MasterSync.
* Ako **ima** red u Google Sheet-u: nastavi ovaj runbook.

### Korak 2: Pronađi Google red

U `OTK-*` ili `VOZ-*` sheet-u pronađi red i zapiši:

```text
Sheet name:
Row number:
ClientRecordID:
ServerRecordID:
SyncStatus:
CreatedAtClient:
UpdatedAtClient:
UpdatedAtServer:
ReceivedAt:
```

Za OTK dodatno:

```text
OtkupacID:
Datum:
KooperantID:
KooperantName:
VrstaVoca:
SortaVoca:
Klasa:
Kolicina:
Cena:
TipAmbalaze:
KolAmbalaze:
VozacID:
ParcelaID:
```

Za VOZ dodatno:

```text
VozacID:
Datum:
KupacID:
KupacName:
KolicinaKlI:
KolicinaKlII:
TipAmbalaze:
KolAmbalaze:
OtkupRecordIDs:
BrojZbirne:
```

### Korak 3: Klasifikuj status

| Google `SyncStatus` | Značenje                            | Operator sme                                 |
| ------------------- | ----------------------------------- | -------------------------------------------- |
| `Synced`            | spremno za master import            | pokrenuti MasterSync                         |
| `Synced>Master`     | već uvezeno                         | proveriti master red, ne reimport bez odluke |
| `Duplicate`         | MasterSync ga smatra duplikatom     | naći master duplikat i dokumentovati         |
| `SyncError`         | import/writeback/validacija problem | analizirati razlog, ne ručno menjati         |
| prazno/nepoznato    | van contract-a                      | tehnički pregled                             |

### Korak 4: Proveri master tabelu

Za OTK:

* traži u `tblOtkup` po `ClientRecordID` ako postoji kolona/sync lineage;
* ako nema lineage, traži po poslovnom ključu:

  * datum;
  * kooperant;
  * stanica/otkupac;
  * vrsta/sorta;
  * klasa;
  * količina;
  * cena;
  * vozač;
  * parcela.

Za VOZ:

* traži u `tblZbirna` po `ClientRecordID` ako postoji;
* traži po `ZbirnaID` iz `ServerRecordID`;
* traži po `BrojZbirne`;
* traži po vozaču, datumu, kupcu i količinama.

### Korak 5: Proveri log

Ako import nije uspeo ili je status čudan, proveri dnevni log za:

```text
ImportOtkupFromPWA
ImportOtkupFromPWA_Core
FindOTKSheets
ImportOneOTKSheet
ValidateOTKSheetHeader
ImportZbirneFromPWA / VOZ import funkcije
WriteBackOTKSyncStatus
WriteBackVOZSyncStatus
Google auth / ReadSheetData / WriteSheetData errors
```

### Korak 6: Izaberi akciju

| Situacija                                         | Akcija                                                                                              |
| ------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| Google red `Synced`, nema master reda             | pokrenuti MasterSync                                                                                |
| Google red `Synced>Master`, master red postoji    | zatvoriti kao već uvezeno                                                                           |
| Google red `Synced>Master`, master red ne postoji | proveriti rollback/crash/journal; ne vraćati status ručno bez tehničke odluke                       |
| Google red `Duplicate`                            | pronaći postojeći master red i dokumentovati                                                        |
| Google red `SyncError` zbog master data           | ispraviti master data, zatim kontrolisano reprocess po odluci tehničkog owner-a                     |
| Google red `SyncError` zbog schema drift          | popraviti header contract; ne dodavati kolone ad hoc                                                |
| Google red `SyncError` zbog writeback fail        | proveriti da li je master append commitovan; ako jeste, ručno writeback samo uz ticket              |
| VOZ ima `BrojZbirne = ZBR-*`                      | tretirati kao writeback bug; ne koristiti taj broj kao poslovni dokumentni broj                     |
| VOZ nema `BrojZbirne`, ali ima `Synced>Master`    | proveriti `tblZbirna`; ako postoji poslovni broj, dopisati writeback uz odobrenje tehničkog owner-a |

---

## 8. Retry i reprocess pravila

### 8.1. Kada smeš ponovo da pokreneš MasterSync

Smeš ponovo pokrenuti MasterSync kada:

* Google red ima `SyncStatus = Synced`; ili
* prethodni pokušaj nije commitovao master promenu; ili
* problem je bio Google auth / Drive listing / ReadSheetData pre lokalnih append-ova; ili
* tehnički owner je potvrdio da nema master reda i da je reprocess bezbedan.

### 8.2. Kada ne smeš ponovo da uvoziš

Ne reimportovati ako:

* `SyncStatus = Synced>Master` i postoji master red;
* `SyncStatus = Duplicate` i postojeći master red je potvrđen;
* nije jasno da li je Excel commitovao, a Google writeback nije uspeo;
* VOZ red već ima `ZbirnaID`, ali `BrojZbirne` writeback nedostaje;
* postoji mogućnost da će se napraviti drugi `OTK-*` ili `ZBR-*` master ID za isti `ClientRecordID`.

### 8.3. Kada sme ručni writeback

Ručni writeback u Google Sheet sme samo tehnički owner, i to kada je dokazano:

```text
[ ] master red postoji u Excelu
[ ] master ID je poznat
[ ] ClientRecordID odgovara Google redu
[ ] nema drugog master reda za isti ClientRecordID
[ ] zna se tačan Google row number
[ ] urađen backup / postoji ticket
```

Za OTK ručni writeback minimalno:

```text
ServerRecordID = OtkupID ili odgovarajući master ID
SyncStatus = Synced>Master
```

Za VOZ ručni writeback minimalno:

```text
ServerRecordID = ZbirnaID
SyncStatus = Synced>Master
BrojZbirne = poslovni broj zbirne, ne ZBR-*
```

---

## 9. Kako sprečavaš dupli master dokument

Glavni anti-duplicate mehanizmi:

1. `ClientRecordID` je primarni PWA dedupe key.
2. GAS sync upisuje/updates by `ClientRecordID`, umesto da svaki retry appenduje novi red.
3. Excel MasterSync uvozi samo `SyncStatus = Synced`.
4. Posle uspešnog master import-a writeback postavlja `Synced>Master`.
5. `Duplicate` status se koristi kada MasterSync otkrije da red ne treba drugi put uvesti.
6. `ImportOtkupFromPWA_TX` ima transaction rollback za fatal sync greške.
7. Schema drift se tretira kao fatal greška, ne kao “probaj najbolje”.
8. VOZ `BrojZbirne` je odvojen od `ServerRecordID`, da se ne pomešaju tehnički i poslovni identiteti.

Operativno pravilo:

> Ako Google red više nije `Synced`, ne vraćaj ga na `Synced` bez dokaza da master red ne postoji.

---

## 10. Recovery scenariji

### 10.1. Red je u PWA, ali ga nema u Google Sheet-u

Ovo nije MasterSync problem.

Postupak:

1. Na uređaju proveri lokalni IndexedDB red.
2. Zapiši `clientRecordID`, `syncStatus`, `lastServerStatus`, `lastSyncError`.
3. Ako je `pending`, pokreni PWA sync.
4. Ako je `syncing`, proveri stale recovery.
5. Ako je `auth-error`, korisnik se mora ponovo prijaviti.
6. Ako je `empty-response` ili `missing-result`, tretirati kao PWA/GAS sync incident.

### 10.2. Red je u Google Sheet-u kao `Synced`, ali Excel ga nema

Postupak:

1. Proveri da je sheet u pravom folderu `GOOGLE_PWA_FOLDER_ID`.
2. Proveri Google auth.
3. Proveri da li ga `FindOTKSheets` / VOZ finder vidi.
4. Proveri header contract.
5. Pokreni MasterSync.
6. Posle import-a proveri master tabelu i writeback.

### 10.3. Red je `Synced>Master`, ali Excel ga nema

Ovo je nekonzistentno stanje.

Postupak:

1. Ne vraćaj odmah status na `Synced`.
2. Proveri master po `ServerRecordID`.
3. Proveri master po poslovnom ključu.
4. Proveri `Journal/` i `Backup/`.
5. Proveri da li je došlo do rollback-a posle Google writeback-a.
6. Ako master red stvarno ne postoji, tehnički owner odlučuje da li se red vraća na `Synced` za reprocess.

### 10.4. Red je `Duplicate`

Postupak:

1. Nađi master red koji je duplikat.
2. Uporedi `ClientRecordID`, datum, entity, količinu, klasu i cenu.
3. Ako je isti unos, zatvori incident kao već obrađen.
4. Ako nije isti, ne menjati status ručno; eskalirati.

### 10.5. Red je `SyncError` zbog schema drift-a

Postupak:

1. Ne dodavati kolone proizvoljno na kraj sheet-a.
2. Uporediti header sa canonical OTK/VOZ header-om.
3. Popraviti sheet schema po canonical redosledu.
4. Proveriti da li je problem nastao iz stare deployment verzije GAS-a ili ručne izmene sheet-a.
5. Tek posle schema fix-a odlučiti da li se red vraća u `Synced` za import.

### 10.6. VOZ `BrojZbirne` je prazan posle master import-a

Postupak:

1. Proveri `SyncStatus`.
2. Ako je `Synced`, master import još nije završen.
3. Ako je `Synced>Master`, proveri `ServerRecordID` i pronađi `tblZbirna` red.
4. Ako `tblZbirna` ima poslovni `BrojZbirne`, a Google kolona T je prazna, to je writeback problem.
5. Tehnički owner sme ručno upisati `BrojZbirne` u kolonu T samo uz ticket i dokaz.

### 10.7. VOZ `BrojZbirne` je `ZBR-*`

Postupak:

1. Tretirati kao bug/nekonzistentan writeback.
2. `ZBR-*` je `ZbirnaID`, nije poslovni broj zbirne.
3. Naći `tblZbirna` red po tom `ZbirnaID`.
4. Uzeti stvarni `BrojZbirne` iz master tabele.
5. Ispraviti Google kolonu T samo uz ticket.
6. Proveriti kod `WriteBackVOZSyncStatus` pre sledećeg produkcionog import-a.

### 10.8. Google writeback nije uspeo posle master append-a

Ovo je najopasniji scenario za dupli import.

Postupak:

1. Ne pokretati import ponovo naslepo.
2. Proveriti da li master red postoji.
3. Ako postoji, ručno ili kontrolisano upisati writeback status.
4. Ako ne postoji, proveriti rollback/journal.
5. Ako je neodređeno, zaustaviti operaciju i eskalirati tehničkom owner-u.

---

## 11. Canonical OTK sheet contract

OTK Google Sheet mora imati kolone ovim redosledom:

```text
ClientRecordID
ServerRecordID
CreatedAtClient
UpdatedAtClient
UpdatedAtServer
SyncStatus
DeviceID
OtkupacID
Datum
KooperantID
KooperantName
VrstaVoca
SortaVoca
Klasa
Kolicina
Cena
TipAmbalaze
KolAmbalaze
ParcelaID
VozacID
Napomena
ReceivedAt
```

Ako ovaj header nije tačan, MasterSync treba da fail-fast prijavi schema drift.

---

## 12. Canonical VOZ sheet contract

VOZ Google Sheet mora imati kolone ovim redosledom:

```text
ClientRecordID
ServerRecordID
CreatedAtClient
UpdatedAtClient
UpdatedAtServer
SyncStatus
VozacID
Datum
KupacID
KupacName
VrstaVoca
SortaVoca
KolicinaKlI
KolicinaKlII
TipAmbalaze
KolAmbalaze
Klasa
OtkupRecordIDs
ReceivedAt
BrojZbirne
```

Posebno:

* `TipAmbalaze` mora ostati plain-text, jer vrednosti tipa `12/1` ne smeju postati datum.
* `BrojZbirne` mora ostati plain-text, jer vrednosti tipa `4/040526` ne smeju postati datum.
* `BrojZbirne` nije `ZbirnaID`.

---

## 13. Operator komande / akcije

Tipične VBA akcije:

```vb
' OTK import iz PWA u master
Call ImportOtkupFromPWA_TX

' OTK import bez TX wrapper-a samo za tehničku analizu, ne za operatora
Call ImportOtkupFromPWA

' Kreiranje OTK sheetova za stanice
Call CreateOTKSheetsForAllStanice

' Auto kreiranje otpremnica posle PWA otkupa sa VozacID
Debug.Print AutoCreateOtpremniceFromPWA()
```

Za VOZ import koristiti postojeći operator/UI entrypoint za `ImportZbirneFromPWA` ako je izložen u aplikaciji. Ako nije, tehnički owner treba da pozove odgovarajući VBA import samo nakon provere sheet-a i header-a.

Ne koristiti ručno editovanje master tabela osim kao poslednju meru.

---

## 14. Ko donosi odluku

### Operator sme sam

* pokrenuti MasterSync kada je red `Synced`;
* prijaviti `SyncError` sa podacima;
* proveriti master red;
* proveriti Google status;
* obavestiti korisnika da red čeka sync/import.

### Tehnički owner odlučuje

* vraćanje `SyncStatus` iz `Synced>Master`, `Duplicate` ili `SyncError` nazad na `Synced`;
* ručni Google writeback;
* schema drift repair;
* popravku `WriteBackVOZSyncStatus`;
* recovery posle writeback-fail / rollback neodređenog stanja;
* ručnu izmenu `ServerRecordID` ili `BrojZbirne`.

### Poslovni owner odlučuje

* da li je duplikat stvarno poslovni duplikat ili drugi realni unos;
* koji od dva slična otkupa je važeći;
* da li se pogrešna zbirna stornira ili koriguje;
* da li se povezana otpremnica/prijemnica/faktura menja.

---

## 15. Checklist za zatvaranje incidenta

```text
[ ] Identifikovan domen: OTK ili VOZ
[ ] Identifikovan Google sheet name
[ ] Identifikovan Google row number
[ ] Identifikovan ClientRecordID
[ ] Identifikovan ServerRecordID
[ ] Proveren SyncStatus
[ ] Proveren master red u Excelu
[ ] Ako je OTK, proveren tblOtkup / tblAmbalaza efekat
[ ] Ako je VOZ, proveren tblZbirna / BrojZbirne
[ ] Proveren writeback u Google Sheet
[ ] Ako je SyncError, dokumentovan razlog
[ ] Ako je Duplicate, dokumentovan master red koji je duplikat
[ ] Ako je ručni writeback, postoji ticket i odobrenje tehničkog owner-a
[ ] Ako je poslovno sporno, postoji odluka poslovnog owner-a
[ ] Korisnik obavešten
```

---

## 16. Primeri odluke

### Primer A: OTK red je `Synced`, nema ga u Excelu

Zaključak: čeka master import.
Akcija: pokreni `ImportOtkupFromPWA_TX`, zatim proveri `tblOtkup` i Google writeback.

### Primer B: OTK red je `Synced>Master`, korisnik kaže da ga nema

Zaključak: moguć pogrešan search ili nekonzistentno stanje.
Akcija: traži master po `ServerRecordID` i poslovnom ključu. Ako ne postoji, proveri journal/backup/log. Ne vraćaj status na `Synced` bez tehničke odluke.

### Primer C: VOZ red je `Synced>Master`, ali `BrojZbirne` je prazan

Zaključak: master import je verovatno prošao, ali writeback kolone T nije.
Akcija: pronađi `tblZbirna` po `ServerRecordID`, uzmi stvarni poslovni `BrojZbirne`, tehnički owner radi kontrolisani writeback.

### Primer D: VOZ `BrojZbirne = ZBR-00012`

Zaključak: u kolonu poslovnog broja upisan je tehnički `ZbirnaID`.
Akcija: ne koristiti taj broj u dokumentnom toku. Naći stvarni `BrojZbirne` u `tblZbirna`, ispraviti Google red i popraviti writeback kod pre produkcionog import-a.

### Primer E: Red je `SyncError` zbog schema drift-a

Zaključak: Google sheet header nije canonical.
Akcija: popraviti header, proveriti GAS/VBA deployment verzije, tek zatim odlučiti reprocess.

### Primer F: Google writeback fail posle lokalnog append-a

Zaključak: visok rizik duplog import-a.
Akcija: ne pokretati ponovo. Prvo dokazati da li master red postoji. Ako postoji, uraditi kontrolisani writeback. Ako ne postoji, recovery kroz log/journal.

---

## 17. Poznate production rupe koje treba zatvoriti

1. Napraviti poseban operator ekran za PWA import incidente: filter po `ClientRecordID`, `SyncStatus`, sheet name, last error.
2. Eksplicitno logovati svaki Google writeback rezultat: sheet, row, old status, new status, master ID.
3. Uvesti `tblPWAMasterSyncEventLog` ili sličan audit log za import/writeback događaje.
4. Dodati admin proceduru “Verify PWA row imported” koja proverava Google row + master row + writeback konzistentnost.
5. Zatvoriti/validirati `WriteBackVOZSyncStatus` tako da kolona B dobija `ZbirnaID`, a kolona T dobija poslovni `BrojZbirne`.
6. Obezbediti da OTK/VOZ sheetovi imaju schema guard pre svake produkcione sezone.
7. Dodati dokumentovan recovery za slučaj: master append uspeo, Google writeback pao.
8. Dodati dnevni report redova sa `SyncStatus = Synced` starijih od npr. 2 sata.
9. Dodati dnevni report redova sa `SyncStatus = SyncError` i `Duplicate`.
10. U PWA prikazati jasnije razliku između “synced to server” i “imported into Excel master”.

Do tada važi konzervativno pravilo:

> `Synced` znači “stiglo do Google transportnog sloja”. Ne znači “ušlo u Excel master”. `Synced>Master` znači “Excel tvrdi da je obradio”. Ako postoji sumnja između ta dva sveta, `ClientRecordID` je glavni trag, a ne datum/količina sama za sebe.
