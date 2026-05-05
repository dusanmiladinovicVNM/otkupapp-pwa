# Production runbook: PWA offline sync i stuck pending/syncing

Status: **operativni runbook za incidente “na telefonu postoji unos, ali nije stigao”, “stoji pending/syncing”, “dupliralo se posle loše konekcije”.**

Aplikacija: **OtkupApp / AgriX PWA**
Domen: **PWA IndexedDB → GAS sync endpoint → Google Sheets → Excel MasterSync**
Glavni kod: `src/js/utils/sync-engine.js`, `src/js/services/db.js`, `src/js/features/otkup/sync.js`, `src/js/features/kooperant/sync.js`, `src/js/features/vozac/zbirna.js`, `gas/Code.gs`

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Uneo sam otkup na telefonu, ali ga nema u sistemu.”
* “Na telefonu piše ČEKA / pending.”
* “Na telefonu piše SYNC..., ali se ništa ne menja.”
* “Bio sam offline, sad sam online, ali unos nije otišao.”
* “Zbirna je kreirana, ali nije stigla.”
* “Tretman/trošak je unet kod kooperanta, ali ga nema kod managementa.”
* “Duplirao mi se unos posle ponovnog klika.”
* “Sesija je istekla, a imam pending podatke.”
* “Aplikacija se resetovala i nestali su lokalni podaci.”

Prvo pravilo:

> Ne briši IndexedDB i ne odjavljuj korisnika pre nego što izvučeš `clientRecordID`, `storeName`, `syncStatus`, `lastServerStatus` i `lastSyncError`.

Minimalni podaci koje operator mora da prikupi:

```text
Korisnik / uloga: Otkupac / Vozac / Kooperant
EntityID: OtkupacID / VozacID / KooperantID
Uređaj / browser:
Online ili offline:
Vreme unosa:
Store: otkupi / zbirne / tretmani / troskovi
clientRecordID:
syncStatus:
serverRecordID:
syncAttempts:
syncAttemptAt:
lastServerStatus:
lastSyncError:
Da li postoji red u Google Sheet-u:
Da li postoji red u Excel masteru:
```

---

## 2. Source of truth: gde se gleda

### 2.1. Prvo mesto: PWA lokalni IndexedDB

Ako korisnik kaže “na telefonu postoji”, prvo se gleda lokalni uređaj.

IndexedDB database ima store-ove:

| Store                | Uloga / domen               | Backend action        |
| -------------------- | --------------------------- | --------------------- |
| `CONFIG.STORE_NAME`  | Otkupac otkupi              | `sync`                |
| `zbirne`             | Vozačke zbirne              | `syncZbirna`          |
| `tretmani`           | Kooperant tretmani / radovi | `syncTretman`         |
| `troskovi`           | Kooperant troškovi          | `syncTrosak`          |
| `CONFIG.STAMM_STORE` | lokalni šifarnici/cache     | nije write sync store |

Svaki sync store ima `clientRecordID` kao key i index `syncStatus`.

### 2.2. Drugo mesto: PWA UI badge/lista

Za Otkupac PWA:

* `ONLINE` znači nema pending/syncing redova;
* `OFFLINE (n)` znači korisnik nema konekciju i ima `n` pending/syncing redova;
* `ČEKA: n` znači postoji `n` redova koji nisu potvrđeni;
* `SYNC...` znači sync je u toku ili postoji `syncing` red.

Queue lista prikazuje pending/syncing redove i `lastSyncError` ako postoji.

Za Vozača:

* lokalne zbirne prikazuju `pending`, `sync...`, `serverRecordID` ili `BrojZbirne`;
* `BrojZbirne` može doći iz backend rezultata ili kasnije iz master writeback-a.

### 2.3. Treće mesto: GAS / Google Sheet

Ako je PWA lokalni red `synced`, sledeći source of truth je Google Sheet:

* Otkupac: `OTK-*` sheet;
* Vozač: `VOZ-*` sheet;
* Kooperant: odgovarajući GAS sheet za tretmane/troškove, zavisno od implementacije.

Traži po `ClientRecordID`.

### 2.4. Četvrto mesto: Excel MasterSync

Ako red postoji u Google-u, ali nije u Excelu, to više nije PWA offline sync incident nego MasterSync incident.

Prebaci na runbook:

```text
Production runbook: PWA MasterSync OTK/VOZ import i writeback
```

### 2.5. PWA/GAS ErrorLog

PWA greške se šalju kroz `reportClientError` na GAS action `logClientError`.

GAS čuva `ErrorLog` sa:

```text
Timestamp
Source
Action
Message
Details
EntityID
Severity
```

Koristi se kada lokalni uređaj više nije dostupan ili kada treba naći greške iz produkcije.

---

## 3. Koji ID pratiš

### 3.1. Primarni ID: `clientRecordID`

`clientRecordID` je glavni incident ID za PWA sync.

On povezuje:

* lokalni IndexedDB red;
* GAS batch result;
* Google Sheet red;
* kasniji Excel MasterSync red;
* dedupe/render merge logiku.

Nikada ne koristi samo datum/količinu kao dokaz. Uvek traži `clientRecordID`.

### 3.2. Sekundarni ID: `serverRecordID`

`serverRecordID` je ID koji vrati server/GAS ili kasniji master writeback.

Važno:

* za OTK može biti server/master ID;
* za VOZ može biti `ZbirnaID`;
* `serverRecordID` nije isto što i poslovni `BrojZbirne`.

### 3.3. Status polja

Obavezno prati:

```text
syncStatus
syncAttempts
syncAttemptAt
lastServerStatus
lastSyncError
syncedAt
updatedAtServer
```

### 3.4. Store-specific poslovni ID-jevi

Za Otkup:

```text
clientRecordID
serverRecordID
otkupacID / entityID
datum
kooperantID
kooperantName
vrstaVoca
sortaVoca
klasa
kolicina
cena
vozacID
parcelaID
```

Za Zbirnu:

```text
clientRecordID
serverRecordID
brojZbirne
vozacID
datum
kupacID
kupacName
kolicinaKlI
kolicinaKlII
otkupRecordIDs
```

Za Tretman:

```text
clientRecordID
kooperantID
parcelaID
datum
mera
artikalID
kolicinaUpotrebljena
vremePocetka
vremeZavrsetka
```

Za Trošak:

```text
clientRecordID
kooperantID
parcelaID
datum
kategorija
iznos
dokumentBroj
```

---

## 4. Sync status značenje

### 4.1. `pending`

Značenje:

* red postoji lokalno;
* još nije poslat ili prethodni pokušaj nije dobio potvrdu;
* red je retryable.

Operator sme:

* proveriti konekciju;
* pokrenuti ručni sync;
* proveriti `lastSyncError`;
* proveriti da li postoji isti `clientRecordID` u Google Sheet-u.

Operator ne sme:

* ručno brisati red;
* ručno praviti isti unos opet;
* resetovati IndexedDB.

### 4.2. `syncing`

Značenje:

* red je označen kao u slanju;
* sync engine je postavio `syncAttemptAt`;
* ako sync traje predugo ili je browser pukao, red može ostati zaglavljen.

Kod ima stale recovery: `syncing` stariji od približno 2 minuta vraća se u `pending` tokom sledećeg sync/bootstrap ciklusa.

Operator sme:

* sačekati normalan završetak ako je sync upravo pokrenut;
* refresh/reopen aplikacije da se pokrene bootstrap recovery;
* ručno pokrenuti sync posle recovery-ja.

Operator ne sme:

* praviti isti unos ponovo;
* brisati `syncing` red;
* ručno menjati status ako ne zna `clientRecordID` i da li je server već primio red.

### 4.3. `synced`

Značenje:

* PWA je dobila uspešan server rezultat;
* red je potvrđen kao `synced`, `duplicate`, `existing`, `inserted` ili `updated` prema server response logici;
* lokalno se čuva `syncedAt`, `serverRecordID`, `updatedAtServer` ako ih server vrati.

Važno:

> `synced` znači “stiglo do GAS/Google transportnog sloja”. Ne znači automatski “ušlo u Excel master”.

Ako korisnik kaže da “ga nema u Excelu”, proveri Google Sheet i MasterSync runbook.

### 4.4. `deleted`

Neki lokalni modeli koriste `deleted` flag za render/filter.

Ako postoji `deleted = true`, ne tretirati red kao aktivan poslovni unos bez dodatne provere.

---

## 5. `lastServerStatus` značenje

Tipični statusi koje sync engine upisuje:

| `lastServerStatus`                                           | Značenje                                              | Sledeći korak                                                      |
| ------------------------------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------ |
| `request-failed`                                             | ceo request je pao                                    | proveri error, retry kada je mreža OK                              |
| `empty-response`                                             | server nije vratio validan JSON/body                  | proveri GAS/API dostupnost                                         |
| `auth-error`                                                 | token/session problem                                 | korisnik se mora ponovo prijaviti, ali ne brisati lokalne podatke  |
| `missing-result`                                             | server response nema rezultat za taj `clientRecordID` | visok rizik neodređenog stanja; proveri Google po `clientRecordID` |
| `feature-disabled`                                           | backend kaže da funkcija nije aktivna                 | ne brojati kao neuspešni pokušaj, proveriti deployment/config      |
| `exception`                                                  | JS exception tokom sync-a                             | proveri ErrorLog / console                                         |
| `stale-syncing-recovered`                                    | lokalni recovery vratio syncing u pending             | sada sme kontrolisani retry                                        |
| `legacy-success`                                             | stari server success bez results array-a              | proveriti backend contract                                         |
| `synced` / `duplicate` / `existing` / `inserted` / `updated` | server smatra uspešno                                 | proveriti Google/MasterSync ako ga nema dalje                      |

---

## 6. Normalan sync tok

### 6.1. Otkupac otkup

1. Korisnik unese otkup.
2. PWA snimi red u IndexedDB store `CONFIG.STORE_NAME` sa `syncStatus = pending`.
3. Ako je online, `syncQueueSafe('post-save')` ili interval/manual pokreće sync.
4. `requestOtkupSync` sprečava paralelni sync i ako tokom sync-a stigne novi zahtev, radi drugi serijski pass.
5. `syncStore` vraća stale `syncing` redove u `pending` ako su prestari.
6. Svi pending redovi prelaze u `syncing` i dobijaju `syncAttemptAt`.
7. PWA šalje action `sync` i `otkupacID` GAS-u.
8. GAS vraća `results[]` po `clientRecordID`.
9. Za svaki uspešan rezultat lokalni red postaje `synced`.
10. Za neuspešne rezultate red se vraća u `pending` uz `lastSyncError`.

### 6.2. Vozač zbirna

1. Vozač kreira zbirnu.
2. PWA pravi lokalni red u store `zbirne` sa `syncStatus = pending`.
3. `confirmZbirna` koristi `withSubmitLock('zbirna:confirm')` da spreči double-click duplikat.
4. Ako je online, pokreće `syncQueueSafe('post-save')`.
5. `syncZbirne` šalje store `zbirne` na GAS action `syncZbirna`.
6. Ako server vrati `brojZbirne`, lokalni red ga upisuje kroz `onResultRecord`.
7. Red postaje `synced` ili se vraća u `pending` uz grešku.

### 6.3. Kooperant tretmani/troškovi

1. Kooperant unese tretman ili trošak.
2. PWA snimi red u `tretmani` ili `troskovi` sa `syncStatus = pending`.
3. `requestKooperantSync` serijski pokreće `syncTretmani` i `syncTroskovi`.
4. Ako je neki sync već u toku, novi zahtev se pamti i radi drugi pass posle prvog.
5. Tretmani idu na action `syncTretman`.
6. Troškovi idu na action `syncTrosak`.

---

## 7. Standardni incident flow

### Korak 1: Utvrdi da li je problem lokalni, GAS ili MasterSync

Postavi tri pitanja:

```text
1. Da li red postoji u lokalnom IndexedDB-u?
2. Da li red postoji u Google Sheet-u / GAS backend-u?
3. Da li red postoji u Excel masteru?
```

Tumačenje:

| Lokalno | Google/GAS | Excel | Problem                                                            |
| ------- | ---------- | ----- | ------------------------------------------------------------------ |
| Da      | Ne         | Ne    | PWA offline/GAS sync                                               |
| Da      | Da         | Ne    | MasterSync import/writeback                                        |
| Da      | Da         | Da    | verovatno UI/search/report problem                                 |
| Ne      | Da         | Možda | lokalni cache očišćen posle uspešnog sync-a ili server-only stanje |
| Ne      | Ne         | Ne    | unos nije sačuvan ili je izgubljen lokalni DB                      |

### Korak 2: Izvuci lokalni red

Na uređaju korisnika, u DevTools → Application → IndexedDB, pronađi odgovarajući store:

```text
CONFIG.STORE_NAME  -> otkupi
zbirne             -> vozačke zbirne
tretmani           -> kooperant tretmani
troskovi           -> kooperant troškovi
```

Zapiši:

```text
storeName:
clientRecordID:
syncStatus:
serverRecordID:
syncAttempts:
syncAttemptAt:
lastServerStatus:
lastSyncError:
syncedAt:
updatedAtServer:
createdAtClient:
updatedAtClient:
entityID:
```

### Korak 3: Proveri sync badge i mrežu

Zapiši:

```text
navigator.onLine:
Badge: ONLINE / OFFLINE / ČEKA / SYNC
Broj pending redova:
Broj syncing redova:
```

Ako je offline, ne radi ništa destruktivno. Korisnik mora prvo imati stabilnu konekciju.

### Korak 4: Klasifikuj status

| Status                         | Akcija                                           |
| ------------------------------ | ------------------------------------------------ |
| `pending`                      | ručni sync kada je online                        |
| `syncing` < 2 min              | sačekati ili proveriti da li sync stvarno traje  |
| `syncing` > 2 min              | pokrenuti bootstrap/stale recovery, zatim retry  |
| `synced`                       | proveriti Google po `clientRecordID`             |
| `pending` + `auth-error`       | ponovna prijava, pa sync                         |
| `pending` + `missing-result`   | proveriti Google pre retry-ja                    |
| `pending` + `empty-response`   | proveriti GAS/API, retry kad je stabilno         |
| `pending` + validaciona greška | ispraviti podatke ili prijaviti poslovni problem |

### Korak 5: Proveri server trag

Ako lokalni red nije `synced`, ali postoji šansa da je request otišao, obavezno proveri Google/GAS po `clientRecordID` pre ručnog dupliranja.

Posebno za:

```text
lastServerStatus = missing-result
lastServerStatus = empty-response
lastServerStatus = exception posle apiPost
syncing koji je pukao tokom request-a
```

### Korak 6: Izvrši dozvoljenu akciju

| Problem                 | Dozvoljena akcija                                         |
| ----------------------- | --------------------------------------------------------- |
| offline                 | sačekati konekciju, zatim manual sync                     |
| pending bez greške      | manual sync                                               |
| stale syncing           | recover stale syncing, zatim manual sync                  |
| auth-error              | ponovna prijava, zatim sync; ne brisati DB                |
| feature-disabled        | proveriti deployment/config; ne forsirati retry loop      |
| server validation error | ispraviti podatke ili eskalirati owner-u                  |
| missing-result          | prvo proveriti Google; zatim kontrolisani retry           |
| duplicate na serveru    | tretirati kao uspeh ako isti `clientRecordID` već postoji |

---

## 8. Retry pravila

### 8.1. Dozvoljen retry

Retry je dozvoljen ako:

* red je `pending`; ili
* red je `syncing`, ali ga je stale recovery vratio u `pending`; ili
* request-level failure nije stigao do servera; ili
* `auth-error` je rešen novom prijavom; ili
* `feature-disabled` je rešen deployment/config promenom; ili
* provereno je da server nema taj `clientRecordID`.

### 8.2. Nije dozvoljen retry bez provere

Ne retry-ovati naslepo ako:

* `syncStatus = synced`;
* `lastServerStatus = missing-result`;
* postoji server/Google red za isti `clientRecordID`;
* korisnik je ručno napravio drugi unos za isti poslovni događaj;
* nije jasno da li je prethodni request stigao do GAS-a.

### 8.3. Kako sync engine sprečava duplikate

Sistem ima zaštite:

1. `clientRecordID` je IndexedDB key.
2. `syncStore` šalje pending batch i očekuje rezultat po `clientRecordID`.
3. `applyServerResults` ažurira tačno onaj lokalni red koji server pominje.
4. Ako server ne pomene neki `clientRecordID`, red se vraća u `pending` sa `missing-result`.
5. `requestOtkupSync` i `requestKooperantSync` sprečavaju paralelne sync pozive i rade drugi pass ako je stigao novi zahtev.
6. `withSubmitLock` sprečava double-submit kod kritičnih UI akcija, npr. `confirmZbirna`.
7. `dedupeRecordsForRender` smanjuje rizik da korisnik vidi isti server/local red dva puta.

Operativno pravilo:

> Retry istog `clientRecordID` je prihvatljiv. Ručno pravljenje novog poslovnog unosa zato što je stari “pending” nije prihvatljivo dok se ne proveri server.

---

## 9. Recovery scenariji

### 9.1. Red je `pending`, korisnik je bio offline

Postupak:

1. Proveri da li je uređaj online.
2. Proveri da li postoji validna sesija/token.
3. Pokreni manual sync.
4. Proveri da li je red prešao u `synced`.
5. Ako nije, zapiši `lastSyncError` i `lastServerStatus`.

### 9.2. Red je `syncing` i stoji zaglavljen

Postupak:

1. Zapiši `clientRecordID` i `syncAttemptAt`.
2. Ako je mlađe od 2 minuta, proveri da li request još traje.
3. Ako je starije od 2 minuta, zatvori/otvori PWA ili pokreni novi sync da se izvrši `recoverStaleSyncingRecords`.
4. Proveri da li je red vraćen u `pending`.
5. Pre retry-ja proveri Google po `clientRecordID` ako postoji sumnja da je request stigao.
6. Pokreni manual sync.

### 9.3. `lastServerStatus = missing-result`

Ovo je visokorizično neodređeno stanje.

Značenje:

* PWA je poslala batch;
* server je vratio `results[]`;
* ali taj `clientRecordID` nije pomenut u rezultatima.

Postupak:

1. Ne praviti novi unos.
2. Proveriti Google/GAS po `clientRecordID`.
3. Ako red postoji u Google-u, lokalno stanje se sme tretirati kao server-side success uz tehnički recovery.
4. Ako red ne postoji, retry je dozvoljen.
5. Ako nije moguće proveriti, eskalirati tehničkom owner-u.

### 9.4. `lastServerStatus = empty-response`

Značenje:

* request je otišao, ali PWA nije dobila validan response.

Postupak:

1. Proveriti GAS dostupnost.
2. Proveriti browser console/network.
3. Proveriti ErrorLog za action.
4. Proveriti Google po `clientRecordID`.
5. Ako nema server reda, retry.
6. Ako postoji server red, ne praviti novi unos.

### 9.5. `lastServerStatus = auth-error`

Značenje:

* token je istekao ili nije validan.

Postupak:

1. Ne brisati lokalne pending podatke.
2. Korisnik se ponovo prijavljuje.
3. Proveriti da je `CONFIG.ENTITY_ID` isti kao pre.
4. Pokrenuti sync.
5. Ako je entity promenjen, ne syncovati dok se ne potvrdi da podaci pripadaju tom korisniku.

### 9.6. `lastServerStatus = feature-disabled`

Značenje:

* backend action je eksplicitno isključen ili nije aktivan.

Postupak:

1. Ne pritiskati sync u loop-u.
2. Proveriti deployment verziju GAS-a.
3. Proveriti da li je feature namerno isključen.
4. Nakon fix-a, redovi ostaju `pending` i mogu se syncovati.

### 9.7. `syncStatus = synced`, ali korisnik kaže “nema u sistemu”

Postupak:

1. Proveriti Google Sheet po `clientRecordID`.
2. Ako Google red postoji, incident prebaciti na MasterSync runbook.
3. Ako Google red ne postoji, proveriti da li je server vratio `duplicate/existing` i gde je postojeći red.
4. Ako nema server traga, ovo je nekonzistentno lokalno stanje; tehnički owner odlučuje recovery.

### 9.8. Dupli unos posle loše konekcije

Postupak:

1. Proveri da li duplikati imaju isti `clientRecordID` ili različite.
2. Ako je isti `clientRecordID`, verovatno je render/server merge problem, ne poslovni duplikat.
3. Ako su različiti `clientRecordID`, korisnik je verovatno napravio dva poslovna unosa ili double-submit lock nije pokrio tok.
4. Proveri vreme unosa i poslovne podatke.
5. Poslovni owner odlučuje koji unos ostaje.
6. Tehnički owner proverava da li treba pojačati `withSubmitLock` / dedupe.

### 9.9. IndexedDB open/migration problem

Kod pokušava recovery reset ako je DB open greška recoverable: version/upgrade/object store/index/blocked/timeout.

Postupak:

1. Ako postoje pending podaci, ne pokretati ručni `resetIndexedDb` bez export-a.
2. Proveriti da li drugi tab drži DB i blokira upgrade.
3. Zatvoriti druge tabove.
4. Ponovo otvoriti aplikaciju.
5. Ako DB mora da se resetuje, korisnik mora znati da lokalni pending podaci mogu biti izgubljeni.
6. Pre reset-a pokušati izvući podatke iz IndexedDB preko DevTools.

### 9.10. Korisnik se odjavio dok ima pending podatke

Postupak:

1. Proveri da li logout briše lokalne store-ove u toj verziji aplikacije.
2. Ako podaci postoje, korisnik mora da se prijavi istim entityID.
3. Ne syncovati pending redove pod drugim korisnikom.
4. Ako je entity mismatch, tehnički owner odlučuje export/import ili ručni unos.

---

## 10. Role-specific procedure

### 10.1. Otkupac: otkup nije stigao

Proveri store:

```text
storeName = CONFIG.STORE_NAME
backend action = sync
entityIdField = otkupacID
runtime flags = otkupacRequestInFlight, otkupacInFlight
```

Postupak:

1. Nađi lokalni red po `clientRecordID`.
2. Proveri `syncStatus`.
3. Ako je `pending`, manual sync.
4. Ako je `synced`, traži u `OTK-*` Google Sheet-u.
5. Ako je u Google-u, prebaci na MasterSync runbook.
6. Ako je duplikat, proveri postojeći red po `clientRecordID`.

### 10.2. Vozač: zbirna nije stigla / nema `BrojZbirne`

Proveri store:

```text
storeName = zbirne
backend action = syncZbirna
entityIdField = vozacID
runtime flag = zbirnaInFlight
```

Postupak:

1. Nađi lokalnu zbirnu po `clientRecordID`.
2. Proveri `syncStatus`, `serverRecordID`, `brojZbirne`.
3. Ako je `pending`, sync.
4. Ako je `synced`, traži `VOZ-*` red po `ClientRecordID`.
5. Ako Google red postoji, ali nema Excel `BrojZbirne`, prebaci na MasterSync runbook.
6. Ako lokalno ima `serverRecordID`, ali nema `brojZbirne`, proveri da li GAS ili MasterSync treba da vrati poslovni broj.

### 10.3. Kooperant: tretman nije stigao

Proveri store:

```text
storeName = tretmani
backend action = syncTretman
entityIdField = kooperantID
runtime flag = tretmaniInFlight
```

Postupak:

1. Nađi lokalni tretman po `clientRecordID`.
2. Proveri parcelu, datum, meru i artikal.
3. Ako je `pending`, sync.
4. Ako je validaciona greška, proveri šifarnike: parcela, artikal, kooperant.
5. Ako je `synced`, proveri GAS sheet/report po `clientRecordID`.

### 10.4. Kooperant: trošak nije stigao

Proveri store:

```text
storeName = troskovi
backend action = syncTrosak
entityIdField = kooperantID
runtime flag = troskoviInFlight
```

Postupak:

1. Nađi lokalni trošak po `clientRecordID`.
2. Proveri kategoriju, iznos, dokument broj i parcelu.
3. Ako je `pending`, sync.
4. Ako backend vrati `FEATURE_DISABLED`, proveri deployment.
5. Ako je `synced`, proveri GAS sheet/report po `clientRecordID`.

---

## 11. Admin/DevTools komande

Koristiti samo kada operator ima pristup uređaju i razume rizik.

### 11.1. Provera pending/syncing redova

U browser console:

```js
await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'pending')
await dbGetByIndex(db, CONFIG.STORE_NAME, 'syncStatus', 'syncing')
await dbGetByIndex(db, 'zbirne', 'syncStatus', 'pending')
await dbGetByIndex(db, 'zbirne', 'syncStatus', 'syncing')
await dbGetByIndex(db, 'tretmani', 'syncStatus', 'pending')
await dbGetByIndex(db, 'tretmani', 'syncStatus', 'syncing')
await dbGetByIndex(db, 'troskovi', 'syncStatus', 'pending')
await dbGetByIndex(db, 'troskovi', 'syncStatus', 'syncing')
```

### 11.2. Provera jednog reda

```js
await dbGet(db, CONFIG.STORE_NAME, '<clientRecordID>')
await dbGet(db, 'zbirne', '<clientRecordID>')
await dbGet(db, 'tretmani', '<clientRecordID>')
await dbGet(db, 'troskovi', '<clientRecordID>')
```

### 11.3. Pokretanje stale recovery-ja

```js
await recoverStaleSyncingRecords(CONFIG.STORE_NAME)
await recoverStaleSyncingRecords('zbirne')
await recoverStaleSyncingRecords('tretmani')
await recoverStaleSyncingRecords('troskovi')
```

### 11.4. Pokretanje role sync-a

```js
await requestOtkupSync('manual-debug')
await requestKooperantSync('manual-debug')
await syncZbirne()
```

### 11.5. Export pre rizične intervencije

Pre bilo kakvog DB reset-a:

```js
JSON.stringify(await dbGetAll(db, CONFIG.STORE_NAME), null, 2)
JSON.stringify(await dbGetAll(db, 'zbirne'), null, 2)
JSON.stringify(await dbGetAll(db, 'tretmani'), null, 2)
JSON.stringify(await dbGetAll(db, 'troskovi'), null, 2)
```

Sačuvati output u incident ticket.

### 11.6. Ručna promena statusa

Ručna promena `syncStatus` je dozvoljena samo tehničkom owner-u.

Primer za stale row koji je proveren da ne postoji na serveru:

```js
const r = await dbGet(db, 'zbirne', '<clientRecordID>')
r.syncStatus = 'pending'
r.lastServerStatus = 'manual-recovery'
r.lastSyncError = 'Manual recovery after server check: not found'
await dbPut(db, 'zbirne', r)
```

Ne koristiti ovo bez server provere.

---

## 12. Ko donosi odluku

### Operator sme sam

* proveriti online/offline status;
* pokrenuti manual sync;
* zabeležiti lokalni red;
* zatražiti od korisnika da ne unosi isti događaj ponovo;
* proveriti da li je `syncing` star i pokrenuti app reload;
* eskalirati sa `clientRecordID` i statusima.

### Tehnički owner odlučuje

* ručno vraćanje `syncing` u `pending`;
* bilo kakav `dbPut` recovery;
* IndexedDB export/import;
* IndexedDB reset;
* tretman `missing-result` i `empty-response` slučajeva;
* zaključak da je server-side duplicate bez lokalnog success-a;
* popravku sync engine-a, endpoint-a ili deployment-a.

### Poslovni owner odlučuje

* šta raditi ako postoje dva različita `clientRecordID` za isti poslovni događaj;
* koji dupli otkup/zbirna/tretman/trošak ostaje;
* da li se pogrešan unos stornira, ignoriše ili koriguje;
* da li se ručno unosi podatak u master ako je lokalni uređaj izgubljen.

### Niko ne sme bez odobrenja

* brisati IndexedDB;
* odjaviti korisnika ako postoje pending podaci i ne zna se logout efekat;
* menjati `clientRecordID`;
* praviti drugi poslovni unos zato što prvi stoji pending;
* syncovati pending podatke pod drugim `entityID`;
* tretirati `synced` kao “ušlo u Excel”.

---

## 13. Checklist za zatvaranje incidenta

```text
[ ] Identifikovana uloga korisnika
[ ] Identifikovan EntityID
[ ] Identifikovan storeName
[ ] Identifikovan clientRecordID
[ ] Proveren syncStatus
[ ] Proveren serverRecordID
[ ] Proveren syncAttempts / syncAttemptAt
[ ] Proveren lastServerStatus
[ ] Proveren lastSyncError
[ ] Proveren da li uređaj ima konekciju
[ ] Proveren da li je sesija validna
[ ] Ako je synced, proveren Google/GAS po clientRecordID
[ ] Ako je Google red postoji, prebačeno na MasterSync ako nema Excel reda
[ ] Ako je retry urađen, potvrđeno da je isti clientRecordID
[ ] Ako postoji duplikat, poslovni owner odlučio koji ostaje
[ ] Ako je ručna DB intervencija, sačuvan export i ticket
[ ] Korisnik obavešten
```

---

## 14. Primeri odluke

### Primer A: Otkup je `pending`, korisnik je bio offline

Zaključak: normalan offline mode.
Akcija: povezati internet, pokrenuti manual sync, proveriti da li prelazi u `synced`.

### Primer B: Zbirna je `syncing` od pre 30 minuta

Zaključak: stale syncing.
Akcija: reload PWA ili pokrenuti recovery, red se vraća u `pending`, proveriti server, zatim retry.

### Primer C: Tretman ima `auth-error`

Zaključak: sesija istekla.
Akcija: korisnik se ponovo prijavljuje istim `KooperantID`, zatim sync. Ne brisati lokalne podatke.

### Primer D: Trošak ima `feature-disabled`

Zaključak: backend action nije aktivan/deployment problem.
Akcija: ne retry loop; proveriti GAS deployment/config. Posle fix-a syncovati pending redove.

### Primer E: Otkup je `synced`, ali nije u Excelu

Zaključak: PWA sync je verovatno uspeo; problem je MasterSync.
Akcija: proveriti `OTK-*` Google Sheet po `clientRecordID`, zatim MasterSync runbook.

### Primer F: `missing-result`

Zaključak: neodređeno stanje.
Akcija: proveriti Google po `clientRecordID`. Ako postoji, ne retry. Ako ne postoji, retry istog `clientRecordID`.

### Primer G: Duplirana zbirna

Zaključak: prvo utvrditi da li je isti ili različit `clientRecordID`.
Akcija: isti ID = render/merge problem; različiti ID = poslovni duplikat, owner odlučuje.

---

## 15. Poznate production rupe koje treba zatvoriti

1. Dodati user-facing ekran “Pending sync details” koji prikazuje `clientRecordID`, `lastServerStatus`, `lastSyncError`.
2. Dodati export dugme za pending lokalne redove pre logout/reset-a.
3. Dodati admin endpoint “find by clientRecordID” preko svih GAS sheet-ova.
4. Dodati server-side idempotency/dedupe garanciju jasno dokumentovanu za svaki action.
5. Dodati durable request log u GAS: requestId, action, entityID, clientRecordID list, result status.
6. Dodati alert ako `syncing` redovi ostanu stariji od 2 minuta.
7. Dodati alert ako `missing-result` nastane više od 0 puta u produkciji.
8. Dodati jasnu UI razliku: “Sinhronizovano na server” vs “Uvezeno u Excel master”.
9. Dodati recovery UI za tehničkog owner-a umesto ručnog DevTools rada.
10. Dodati smoke test za svaki store: `CONFIG.STORE_NAME`, `zbirne`, `tretmani`, `troskovi`.
11. Dodati zaštitu da logout upozori ako postoje pending/syncing redovi.
12. Dodati structured `entityID` mismatch detector pre sync-a.

Do tada važi konzervativno pravilo:

> Lokalni `pending` podatak je poslovno važan podatak. Dok ne znaš da li postoji u GAS/Google po istom `clientRecordID`, ne briši ga, ne pravi novi unos i ne resetuj IndexedDB.
