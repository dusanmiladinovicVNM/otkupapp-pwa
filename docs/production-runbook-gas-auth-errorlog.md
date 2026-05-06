# Production runbook: GAS auth, token fallback i ErrorLog

Status: **operativni runbook za incidente “401/403”, “korisnik ne može da se uloguje”, “sync radi pod pogrešnim entityID”, “ErrorLog se puni”, “backend vraća INTERNAL_ERROR”.**

Aplikacija: **OtkupApp / AgriX PWA + GAS backend**
Domen: **PWA auth/session → GAS token validation → role/entity authorization → endpoint execution → ErrorLog observability**
Glavni kod: `gas/Code.gs`, PWA API/auth klijent, `reportClientError`, global error/rejection handlers

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* “Ne mogu da se ulogujem.”
* “Izbacuje me iz aplikacije.”
* “Dobijam Neautorizovan pristup.”
* “Dobijam Nemate pristup.”
* “Otkupac ne može da syncuje otkup.”
* “Kooperant vidi ili šalje podatke za pogrešnog kooperanta.”
* “Vozač ne vidi svoje zbirne.”
* “Management akcija ne radi.”
* “ErrorLog se puni.”
* “ErrorLog se ne puni iako korisnik prijavljuje greške.”
* “GAS vraća Interna greška.”
* “Token fallback pravi gomilu starih tokena.”

Prvo pravilo:

> 401 nije isto što i 403. 401 znači token/session problem. 403 znači token postoji, ali role/entity nema pravo na traženu akciju.

Minimalni podaci koje operator mora da prikupi:

```text
Korisnik:
Role: Otkupac / Vozac / Kooperant / Management
EntityID iz sesije:
Action koja pada:
HTTP/JSON code: 401 / 403 / INTERNAL_ERROR / BAD_JSON / FEATURE_DISABLED...
Request time:
Device/browser:
Da li login uspeva:
Da li drugi korisnici iste role imaju isti problem:
Da li ErrorLog ima red:
ErrorLog Timestamp / Action / Message / EntityID:
ClientRecordID ako je sync incident:
```

---

## 2. Source of truth: gde se gleda

### 2.1. PWA session state

Na PWA strani proveriti:

```text
token
role
entityID
username/display name ako postoji
CONFIG.ENTITY_ID
CONFIG.ROLE / current role config
last API action
last API response
```

Ako je sync incident, proveri i lokalni IndexedDB red:

```text
clientRecordID
syncStatus
lastServerStatus
lastSyncError
```

### 2.2. GAS `doPost`

GAS `doPost` tok je:

1. kreira `requestId`;
2. parsira JSON body;
3. pusti public read akcije;
4. `login` ide pre token provere;
5. `saveParcelPolygon` je trenutno public-write izuzetak ako je namerno;
6. `logClientError` ide pre hard auth check-a, best-effort;
7. za ostale akcije radi `validateToken(data.token)`;
8. ako token nije validan: `401`;
9. uzima `tokenData` kroz `getTokenData`;
10. `handleAuthorizedRead` pokriva read endpoint-e;
11. write/sync endpoint-i proveravaju role i entity ownership;
12. ako nema prava: `403`;
13. endpoint izvršava akciju pod `withLock` gde treba;
14. neuhvaćene greške idu u `logUnhandledGasError` i vraćaju `INTERNAL_ERROR` sa `requestId`.

### 2.3. ErrorLog sheet

`logError` pravi ili koristi Drive spreadsheet `ErrorLog` u `MASTER_FOLDER_ID` folderu.

Kolone:

```text
Timestamp
Source
Action
Message
Details
EntityID
Severity
```

`logError` je best-effort: ako logging padne, ne sme da sruši glavni endpoint.

### 2.4. Token storage

Arhitektura navodi persistent token fallback u GAS sloju:

* primarni sloj: `CacheService`;
* fallback sloj: `PropertiesService`;
* purge/cleanup mora biti aktivan kroz `purgeExpiredTokens` / trigger setup.

Runbook pravilo:

> Ako token validacija puca sporadično, proveri i cache i fallback ponašanje. Ako tokeni nikad ne ističu/čiste se, proveri purge trigger.

### 2.5. GAS deployment

Proveriti:

```text
Web App URL u PWA config-u
Deployment version
Execute as: Me
Access: Anyone
Da li PWA gađa staru GAS deployment URL
Da li repo Code.gs odgovara deployed Code.gs
```

---

## 3. Koji ID pratiš

Za auth incident prati:

```text
username
role
entityID
token prefix / token hash, ne pun token u ticket-u
action
requestId
HTTP/JSON code
ErrorLog timestamp
```

Za sync/auth incident dodatno:

```text
clientRecordID
storeName
syncStatus
lastServerStatus
lastSyncError
```

Za role/entity incident:

```text
tokenData.role
tokenData.entityID
payload entityID: otkupacID / vozacID / kooperantID
expected owner entityID
```

Incident ticket minimum:

```text
Action:
Role:
Session EntityID:
Payload EntityID:
Expected EntityID:
Response code:
Error message:
RequestId:
ErrorLog row:
ClientRecordID ako postoji:
Decision:
```

---

## 4. Normalan login/session tok

1. PWA šalje:

```text
action = login
username
pin
```

2. GAS poziva `authenticateUser`.
3. Ako su kredencijali validni, GAS vraća token i podatke korisnika.
4. PWA čuva token i role/entity session state.
5. Svi naredni endpoint-i šalju token.
6. GAS radi `validateToken` i `getTokenData`.
7. Endpoint proverava role/entity ownership.
8. Ako je sve u redu, akcija se izvršava.

---

## 5. 401 vs 403 vs INTERNAL_ERROR

### 5.1. 401 / `Neautorizovan pristup`

Značenje:

* token ne postoji;
* token je istekao;
* token nije pronađen ni u cache ni fallback store-u;
* token payload ne može da se pročita;
* PWA šalje pogrešan/star token;
* korisnik je možda odjavljen ili je deployment promenjen.

Akcija:

1. Ne brisati pending lokalne podatke.
2. Korisnik se ponovo prijavljuje.
3. Proveriti da je role/entityID isti kao pre.
4. Pokrenuti sync ponovo.
5. Ako se 401 ponavlja odmah posle login-a, proveriti token storage/fallback/deployment.

### 5.2. 403 / `Nemate pristup`

Značenje:

* token je validan;
* role nije dozvoljena za action;
* ili `tokenData.entityID` ne odgovara payload entityID-u.

Primeri:

```text
Otkupac pokušava sync za drugi OtkupacID
Kooperant šalje tretman za drugi KooperantID
Vozač update-uje tuđi VozacID
Kooperant pokušava Management-only akciju
saveFiskalniMapiranje traži Management
```

Akcija:

1. Ne retry u loop-u.
2. Proveriti session role/entityID.
3. Proveriti payload entityID.
4. Ako je korisnik prijavljen pogrešno, logout/login ispravnim nalogom.
5. Ako je endpoint pogrešno klasifikovan kao Management-only, tehnički owner odlučuje promenu auth matrice.

### 5.3. `INTERNAL_ERROR`

Značenje:

* neuhvaćena GAS greška;
* response treba da sadrži `requestId`;
* detalji su u ErrorLog-u kroz `logUnhandledGasError`.

Akcija:

1. Zapiši `requestId`.
2. Proveri ErrorLog po vremenu/action-u/requestId u Details.
3. Ne zaključivati da je user kriv.
4. Ako je sync akcija, proveriti da li je partial write već nastao.
5. Tehnički owner popravlja backend ili podatke.

### 5.4. `BAD_JSON`

Značenje:

* request body nije validan JSON.

Akcija:

* proveriti PWA API client;
* proveriti da li je mreža/proxy modifikovao request;
* proveriti console i ErrorLog.

### 5.5. `FEATURE_DISABLED`

Značenje:

* endpoint je namerno isključen ili placeholder.

Akcija:

* ne retry u loop-u;
* proveriti da li je feature u scope-u launch-a;
* ako mora da radi, deployment/config fix.

---

## 6. Endpoint auth matrix: šta proveravaš

### 6.1. Public/pre-auth endpoint-i

| Action                                                                        | Napomena                                                           |
| ----------------------------------------------------------------------------- | ------------------------------------------------------------------ |
| `login`                                                                       | nema token, služi za token issue                                   |
| `logClientError`                                                              | pre-auth best-effort; može pokušati token lookup ako token postoji |
| `getParcelGeo`, `getParcelMeteo`, `getParcelMeteoLatest`, `getAllMeteoLatest` | public read                                                        |
| `saveParcelPolygon`                                                           | trenutno public write exception / open production decision         |

### 6.2. Role + entity protected sync endpoint-i

| Action               | Role                  | Entity check                                  |
| -------------------- | --------------------- | --------------------------------------------- |
| `sync`               | Otkupac, Management   | Otkupac mora imati isti `otkupacID`           |
| `syncZbirna`         | Vozac, Management     | Vozač mora imati isti `vozacID`               |
| `syncTretman`        | Kooperant, Management | Kooperant mora imati isti `kooperantID`       |
| `syncTrosak`         | Kooperant, Management | Kooperant mora imati isti `kooperantID`       |
| `syncAgromere`       | Kooperant, Management | Kooperant mora imati isti `kooperantID`       |
| `syncOprema`         | Kooperant, Management | Kooperant mora imati isti `kooperantID`       |
| `parseFiskalni`      | Kooperant, Management | Kooperant ne sme za drugog `kooperantID`      |
| `parseFiskalniImage` | Kooperant, Management | Kooperant ne sme za drugog `kooperantID`      |
| `saveFiskalni`       | Kooperant, Management | Kooperant se forsira/proverava na svoj entity |
| `uploadPdf`          | Otkupac, Management   | Otkupac samo za svoj `otkupacID`              |
| `updateKamionStatus` | Vozac, Management     | Vozaču se namešta sopstveni `vozacID`         |

### 6.3. Management-only endpoint-i

Primeri:

```text
getMgmtAll
getMgmtKartica
getMgmtFakture
getMgmtFakturaStavke
getMgmtSaldoOM
getMgmtSaldoKupci
getMgmtOtkupPoOM
getMgmtPredatoPoKupcu
saveWarRoomDemand
removeWarRoomDemand
updateDemandPrimljeno
saveDispecer
updateDispecer
removeDispecer
saveIzdavanje
saveFiskalniMapiranje
createArtikal
```

Ako non-Management korisnik dobije 403 na ove akcije, to je očekivano.

---

## 7. Standardni incident flow

### Korak 1: Zapiši response

Iz PWA/network/console/ticket-a zapiši:

```text
Action:
Response success:
Response code:
Response error:
RequestId:
Role:
EntityID:
Payload entity field:
```

### Korak 2: Klasifikuj grešku

| Code             | Kategorija          | Sledeći korak                     |
| ---------------- | ------------------- | --------------------------------- |
| 401              | token/session       | login, token fallback, deployment |
| 403              | role/entity authz   | proveri role i payload entityID   |
| BAD_JSON         | client/request body | proveri PWA api client / console  |
| INTERNAL_ERROR   | backend exception   | ErrorLog + requestId              |
| FEATURE_DISABLED | namerno isključeno  | scope/deployment odluka           |
| network/timeout  | transport           | retry samo ako idempotency jasna  |

### Korak 3: Proveri ErrorLog

Filter:

```text
Timestamp oko incidenta
Action = <action>
EntityID = <entityID>
Message sadrži code/error
Details sadrži requestId
```

Ako nema ErrorLog reda:

* 401/403 se često vraćaju kontrolisano i ne moraju uvek biti error log;
* `logError` može biti best-effort failure;
* client možda nije pozvao `reportClientError`;
* action je možda public read ili login fail bez unhandled exception.

### Korak 4: Proveri token/session

Za 401:

```text
Da li token postoji u PWA?
Da li je korisnik tek ulogovan?
Da li se 401 javlja svim korisnicima ili jednom?
Da li je deployment URL promenjen?
Da li token fallback/purge radi?
Da li korisnik ima pending podatke pre logout-a?
```

### Korak 5: Proveri role/entity

Za 403:

```text
Role iz tokena:
EntityID iz tokena:
Payload otkupacID/vozacID/kooperantID:
Da li action traži Management:
Da li korisnik pokušava tuđi entitet:
```

### Korak 6: Izvrši dozvoljenu akciju

| Problem                                | Dozvoljena akcija                                |
| -------------------------------------- | ------------------------------------------------ |
| istekao token, nema pending podataka   | ponovni login                                    |
| istekao token, ima pending podataka    | login istim entityID, zatim sync                 |
| role mismatch                          | login ispravnim nalogom ili promena prava        |
| entity mismatch                        | ne syncovati; proveriti čiji su podaci           |
| Management-only endpoint iz Kooperanta | očekivan 403 ili promena auth matrice ako je bug |
| INTERNAL_ERROR                         | tehnički owner + ErrorLog/requestId              |
| ErrorLog ne radi                       | proveriti Drive folder/permissions/logError      |

---

## 8. Recovery scenariji

### 8.1. Korisnik dobija 401 pri sync-u

Postupak:

1. Ne brisati IndexedDB.
2. Zapisati store i pending `clientRecordID`.
3. Korisnik se ponovo prijavljuje.
4. Proveriti da novi session ima isti `role` i `entityID`.
5. Pokrenuti sync.
6. Ako se 401 ponavlja odmah, tehnički owner proverava token storage/fallback i deployment URL.

### 8.2. Korisnik dobija 403 pri sync-u

Postupak:

1. Proveriti `tokenData.entityID`.
2. Proveriti payload entityID:

   * `otkupacID`;
   * `vozacID`;
   * `kooperantID`.
3. Ako se ne poklapaju, ne syncovati.
4. Ako je PWA config pogrešno postavio entityID, tehnički owner popravlja session/config.
5. Ako korisnik stvarno treba pravo na drugi entity, Management/owner mora promeniti model prava.

### 8.3. Korisnik ne može da se uloguje

Postupak:

1. Proveriti username/pin.
2. Proveriti da li `login` action stiže do GAS-a.
3. Ako login vraća kontrolisani fail, proveriti Users/šifarnik u GAS data source-u.
4. Ako login vraća `INTERNAL_ERROR`, proveriti ErrorLog.
5. Ako svi korisnici ne mogu da se uloguju, proveriti deployment URL, Apps Script runtime i Drive permissions.

### 8.4. PWA šalje podatke pod pogrešnim entityID

Postupak:

1. Stopirati sync dok se ne razjasni.
2. Izvući pending lokalne redove iz IndexedDB.
3. Proveriti session role/entityID.
4. Proveriti da li je korisnik menjao nalog na istom uređaju.
5. Ne slati redove pod drugim nalogom.
6. Tehnički owner radi export/import ili ručno usmeravanje podataka.

### 8.5. ErrorLog se puni istom greškom

Postupak:

1. Grupisati po `Action`, `Message`, `EntityID`.
2. Utvrditi da li je client bug, auth loop, network loop ili backend exception.
3. Ako je 401/auth loop, rešiti session/token, ne backend logiku.
4. Ako je `INTERNAL_ERROR`, koristiti requestId/details.
5. Ako je timeout warning, proveriti Apps Script quota/latency.
6. Ako ErrorLog eksplodira, privremeno ograničiti client retry ili noisy reporter.

### 8.6. ErrorLog se ne puni

Postupak:

1. Proveriti da li greška zaista poziva `reportClientError` ili je kontrolisan 401/403.
2. Proveriti da `logClientError` action radi pre auth check-a.
3. Proveriti da li `MASTER_FOLDER_ID` postoji i da GAS ima Drive permissions.
4. Proveriti da li `ErrorLog` spreadsheet postoji u folderu.
5. Ručno testirati `logError` ili `reportClientError` smoke.
6. Ako logging failuje, `logError` neće srušiti endpoint; mora se gledati Apps Script execution log.

### 8.7. Token fallback ne čisti stare tokene

Postupak:

1. Proveriti da li postoji purge trigger.
2. Pokrenuti `purgeExpiredTokens` ručno ako postoji u deployment-u.
3. Proveriti PropertiesService key count/veličinu.
4. Ako se stari tokeni gomilaju, tehnički owner uvodi/aktivira cleanup.
5. Ako se validni tokeni prebrzo brišu, proveriti TTL i clock assumptions.

### 8.8. Deployment URL je pogrešan

Postupak:

1. Proveriti PWA config Web App URL.
2. Proveriti Apps Script deployment verziju.
3. Ako repo Code.gs i deployed Code.gs nisu isti, odlučiti koji je source of truth.
4. Deploy nove verzije ili vratiti config na ispravnu URL.
5. Posle promene, korisnici moraju reload/cache refresh ako PWA drži staru config verziju.

### 8.9. `saveFiskalniMapiranje` vraća 403 za Kooperanta

Postupak:

1. Prepoznati da je endpoint Management-only.
2. Ako je PWA Kooperant šalje fire-and-forget mapping, 403 može ostati nevidljiv korisniku.
3. Proveriti da li je račun/stavka sačuvana kroz `saveFiskalni`.
4. Ako jeste, ne duplirati račun.
5. Management dodaje mapiranje ili tehnički owner menja auth politiku.

### 8.10. `saveParcelPolygon` public write decision

Postupak:

1. Ako postoji sumnja na neovlašćenu izmenu polygon-a, tretirati kao security/data incident.
2. Proveriti da li action ima token u trenutnom deployment-u.
3. Ako je još public write, proceniti da li je to eksplicitno prihvaćen rizik.
4. Pre produkcije preporuka je prebaciti ga ispod auth check-a.
5. Recovery polygon podataka ide kroz GIS runbook.

---

## 9. Admin/GAS provere

### 9.1. Smoke test ErrorLog-a

Pokrenuti client-side test ili ručni GAS call za `logClientError`.

Očekivanje:

```text
ErrorLog spreadsheet postoji u MASTER_FOLDER_ID
novi red ima Source=PWA
Action=<test action>
Message=<test message>
EntityID=<entity>
```

### 9.2. Ping test

`doGet?action=ping` treba da vrati:

```text
success=true
timestamp=<ISO>
```

Ako ping ne radi, problem je deployment/availability, ne auth.

### 9.3. Token/session test

Za tehničkog owner-a:

```text
login -> token
authorized read -> success
same role wrong entity -> 403
expired/invalid token -> 401
Management-only action as non-management -> 403
```

### 9.4. ErrorLog lookup

Filteri:

```text
Timestamp >= incident time - 10 min
Action = failing action
EntityID = user entity
Details contains requestId
Severity = error/warning
```

---

## 10. Kako sprečavaš pogrešan pristup i gubitak pending podataka

Zaštite u sistemu:

1. `login` je jedini normalan token issue path.
2. `validateToken` blokira endpoint-e bez validne sesije.
3. `getTokenData` je source za role/entity odluke.
4. `requireRole` blokira pogrešnu rolu.
5. `requireEntity` blokira tuđi entityID.
6. Management ima šira prava, ali mora biti eksplicitno `role=Management`.
7. `logClientError` radi pre auth check-a da bi i auth/session greške bile vidljive.
8. `ErrorLog` je best-effort i ne sme srušiti glavni flow.
9. Persistent token fallback smanjuje rizik da CacheService evikcija izbaci aktivnog korisnika.
10. Purge/cleanup sprečava gomilanje tokena u fallback store-u.

Operativno pravilo:

> Ako ima pending lokalnih podataka, prvo sačuvaj `clientRecordID` i proveri entityID. Tek onda logout/login ili sync recovery.

---

## 11. Ko donosi odluku

### Operator sme sam

* razlikovati 401 i 403;
* tražiti od korisnika ponovni login kod 401;
* prikupiti screenshot/network response;
* zabeležiti role/entityID/action;
* proveriti da korisnik ne unosi isti podatak ponovo;
* eskalirati sa ErrorLog redom.

### Tehnički owner odlučuje

* token fallback/purge intervencije;
* deployment rollback/upgrade;
* promenu endpoint auth matrix-a;
* ručni recovery pending podataka posle entity mismatch-a;
* ErrorLog debugging;
* zaključavanje `saveParcelPolygon`;
* smanjenje noisy client logging-a.

### Poslovni/security owner odlučuje

* ko sme Management prava;
* da li korisnik sme raditi za više entity-ja;
* šta raditi ako su podaci poslati pod pogrešnim entitetom;
* da li je public GIS write prihvatljiv rizik;
* da li se sumnjiv access tretira kao security incident.

### Niko ne sme bez odobrenja

* menjati role/entityID u tokenu ručno;
* syncovati pending podatke pod drugim nalogom;
* dati Management rolu da “reši” 403 bez poslovne odluke;
* brisati ErrorLog tokom incidenta;
* ignorisati ponavljajući 403 kao “mrežni problem”;
* kopirati pune tokene u ticket ili chat.

---

## 12. Checklist za zatvaranje incidenta

```text
[ ] Identifikovan action
[ ] Identifikovan response code
[ ] Razlikovan 401 vs 403 vs INTERNAL_ERROR
[ ] Identifikovan role
[ ] Identifikovan session entityID
[ ] Identifikovan payload entityID
[ ] Ako je sync, identifikovan clientRecordID
[ ] Proveren ErrorLog po action/time/entity/requestId
[ ] Ako je 401, korisnik se loginovao istim entityID pre sync-a
[ ] Ako je 403, potvrđeno da li je očekivana authz blokada ili bug
[ ] Ako je INTERNAL_ERROR, tehnički owner ima requestId/details
[ ] Ako su pending podaci postojali, nisu obrisani
[ ] Ako je deployment problem, potvrđena ispravna Web App URL/verzija
[ ] Korisnik obavešten
```

---

## 13. Primeri odluke

### Primer A: Otkupac dobija 401 pri sync-u

Zaključak: session/token problem.
Akcija: ne brisati pending otkupe; login istim OtkupacID; sync ponovo.

### Primer B: Kooperant dobija 403 na `syncTretman`

Zaključak: token je validan, ali payload `kooperantID` se ne poklapa ili role nije Kooperant/Management.
Akcija: proveriti session i payload; ne syncovati pod drugim kooperantom.

### Primer C: Management-only `saveFiskalniMapiranje` pada iz Kooperant PWA

Zaključak: očekivan 403 prema trenutnoj auth politici, ali UX/design problem ako korisnik očekuje da se mapiranje nauči.
Akcija: ne duplirati račun; Management doda mapping ili tehnički owner promeni auth flow.

### Primer D: Svi korisnici dobijaju 401 posle deploy-a

Zaključak: deployment/token/config problem.
Akcija: proveriti Web App URL, deployed Code.gs, token validation/fallback. Ne resetovati PWA data.

### Primer E: ErrorLog nema red za korisničku grešku

Zaključak: možda je kontrolisani 401/403, ili `reportClientError` nije pozvan, ili ErrorLog failuje best-effort.
Akcija: testirati `logClientError`, proveriti Drive permissions i execution log.

### Primer F: `INTERNAL_ERROR` sa requestId

Zaključak: backend exception.
Akcija: naći ErrorLog red po requestId, proveriti da li je partial write nastao, tek onda retry.

---

## 14. Poznate production rupe koje treba zatvoriti

1. Dodati admin endpoint “whoami” koji vraća `role`, `entityID`, token expiry i deployment version.
2. Dodati token hash / session ID u ErrorLog, nikad pun token.
3. Dodati structured audit za login/logout/token refresh.
4. Dodati dashboard za 401/403 po action-u i entityID-u.
5. Dodati smoke test za celu auth matrix-u pre svakog deploy-a.
6. Dodati automatski purge trigger verification za token fallback.
7. Dodati alert ako PropertiesService token count raste preko praga.
8. Dodati jasnu PWA poruku: “Sesija istekla — podaci su sačuvani lokalno, prijavite se ponovo istim nalogom.”
9. Dodati blocking guard za entity mismatch pre sync-a na client strani.
10. Zaključati `saveParcelPolygon` tokenom ili dokumentovati prihvaćen rizik.
11. Razjasniti auth policy za `saveFiskalniMapiranje` iz Kooperant PWA toka.
12. Dodati deployment version u svaki GAS response.
13. Dodati correlation/requestId u client-side API error prikaz.
14. Dodati ErrorLog retention/archival politiku osim 30-day purge.

Do tada važi konzervativno pravilo:

> 401 rešavaš sesijom, 403 rešavaš pravima i entity ownership-om, INTERNAL_ERROR rešavaš preko requestId/ErrorLog-a. Pending lokalne podatke ne brišeš dok ne znaš pod kojim entityID-em treba da se pošalju.
