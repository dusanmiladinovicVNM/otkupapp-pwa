# Production runbook: v6.14 Monitoring, Alerts i Observability

Predložen path: `docs/production-runbook-monitoring-observability.md`

Status: **operativni runbook za incidente “monitoring ne radi”, “Health je CRITICAL/WARN”, “nema event-a”, “Alerts se pune”, “SEF je unknown/manual review”, “backup nije svež”, “MasterData Sync ne daje signal”, “monitoring možda curi tajne podatke”.**

Aplikacija: **OtkupApp / AgriX**
Domen: **VBA/PWA/GAS event source → GAS Monitoring.gs ingest → OtkupApp_Monitoring_PROD workbook → Health/Events/Errors/SEFStatus/Backups/Alerts/AuditCritical**
Glavni slojevi: `modMonitoring`, `Monitoring.gs`, `OtkupApp_Monitoring_PROD`, `tblSEFConfig`, GAS Script Properties, watchdog triggers

---

## 1. Kada korisnik/operator kaže problem

Tipični incidenti:

* “Monitoring ne radi.”
* “Health je CRITICAL.”
* “GAS API je prazan/nepoznat.”
* “Google Sheets DB je WARN/CRITICAL.”
* “VBA Client ne šalje event-e.”
* “SEF alert stoji otvoren.”
* “Backup je star ili se ne vidi u monitoringu.”
* “MasterData Sync health je prazan.”
* “Alerts se pune istim incidentom.”
* “AuditCritical ima SEF event, ali ne znamo šta da radimo.”
* “Monitoring test vraća HTTP 200, ali nema reda u sheet-u.”
* “Monitoring test pada zbog secret-a.”
* “Nema daily summary-ja.”
* “Watchdog trigger nije radio.”
* “Plašim se da monitoring loguje SEF API key / token / UBL XML.”

Prvo pravilo:

> Monitoring je dijagnostički sloj, nije business source of truth. Ako monitoring ne radi, ne smeš zaključiti da business operacija nije uspela. Proveri canonical tabele/logove za konkretan domen.

Drugo pravilo:

> Ako business operacija radi, a monitoring ne radi, ne zaustavljaš posao automatski. Ako monitoring pokazuje CRITICAL za SEF, backup, auth ili master sync, zaustavljaš samo pogođeni domen dok ne potvrdiš stanje u canonical izvorima.

Minimalni podaci koje operator mora da prikupi:

```text
Incident time:
Component: PWA Sync / GAS API / Google Sheets DB / VBA Client / SEF API / Backup / Auth / MasterData Sync / Offline Queue
Health status: OK / WARN / CRITICAL / UNKNOWN
EventType:
Severity: INFO / WARN / ERROR / CRITICAL
CorrelationId:
EntityType:
EntityId:
Workbook/environment: DEV / PROD
DeviceId:
AppVersion:
GAS deployment URL:
Monitoring workbook ID:
Alert ID / status:
AuditCritical row if exists:
Canonical business source checked: Da/Ne
```

---

## 2. Source of truth: gde se gleda

### 2.1. Monitoring workbook

Canonical monitoring workbook:

```text
OtkupApp_Monitoring_PROD
```

Tabovi:

| Tab             | Svrha                                  |
| --------------- | -------------------------------------- |
| `Health`        | trenutni status komponenti             |
| `Events`        | append-only centralni event stream     |
| `Errors`        | strukturisane runtime/business greške  |
| `SyncStatus`    | PWA sync/offline queue samples         |
| `SEFStatus`     | SEF status per faktura/correlation     |
| `UserSessions`  | PWA/VBA session/heartbeat record-i     |
| `Backups`       | istorija backup success/fail           |
| `Alerts`        | otvoreni/rešeni operational alerts     |
| `AuditCritical` | kritični audit/business-impact event-i |

### 2.2. Health tab

Health je prvi ekran, ali nije dovoljan za root-cause.

Canonical komponente:

```text
PWA Sync
GAS API
Google Sheets DB
VBA Client
SEF API
Backup
Auth
MasterData Sync
Offline Queue
```

Za svaku komponentu proveri:

```text
Component
Status
LastSeen / LastUpdated
LastEventType
LastSeverity
Message
CorrelationId
NextAction
```

Tumačenje:

| Status         | Značenje                                       | Akcija                                          |
| -------------- | ---------------------------------------------- | ----------------------------------------------- |
| `OK`           | poslednji signal je zdrav                      | proveri samo ako korisnik prijavljuje problem   |
| `WARN`         | potencijalno degradirano stanje                | proveri Events/Alerts i domen                   |
| `CRITICAL`     | operativno rizično stanje                      | proveri Alerts/AuditCritical i canonical source |
| prazno/unknown | nema dovoljno signala ili setup nije kompletan | proveri setup/watchdog                          |

### 2.3. Events tab

Events je append-only stream. Traži po:

```text
Timestamp
source
severity
eventType
correlationId
entityType
entityId
module
functionName
message
payload
```

Events ti govori **šta je sistem prijavio**, ne garantuje da je business state konačan.

### 2.4. Errors tab

Errors sadrži `ERROR` i `CRITICAL` event-e.

Filteri:

```text
severity = ERROR / CRITICAL
eventType
correlationId
entityId
module/functionName
```

Ako event postoji u Errors, proveri i Events za kontekst pre/posle.

### 2.5. Alerts tab

Alerts je radna lista za operativne intervencije.

Proveri:

```text
AlertId
Component
AlertCode
Severity
CorrelationId
Status: Open / Resolved
CreatedAt
LastSeenAt
NextAction
Message
```

Pravilo:

> Alert se ne zatvara zato što je nestao iz UI-ja. Zatvara se kada je canonical stanje provereno i kada Health/Events više ne pokazuju aktivan problem.

### 2.6. AuditCritical tab

AuditCritical je za događaje koji imaju poslovno-pravni ili finansijski rizik.

Tipični kandidati:

```text
SEF unknown/manual review
SEF send exception after local sending
payment/faktura critical event
critical bank mapping failure
backup critical failure
manual override
security/auth critical event
```

Ako događaj postoji u AuditCritical, mora postojati odgovorna osoba i odluka.

### 2.7. Canonical business izvori

Monitoring nikad nije jedini dokaz. Uvek proveri domain source:

| Domen            | Canonical source                                                              |
| ---------------- | ----------------------------------------------------------------------------- |
| SEF              | `tblFakture`, `tblSEFSubmission`, `tblSEFEventLog`, SEF portal/API refresh    |
| Novac            | `tblNovac`, `tblFakture`, `tblOtkup`                                          |
| Dokumentni chain | `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakturaStavke` |
| PWA sync         | IndexedDB, role Google Sheet, GAS response, MasterSync status                 |
| Banka            | `tblBankaImport`, PDF folderi, `tblNovac`, `tblPartnerMap`                    |
| Backup           | `Backup\`, Backups tab, Drive/file system backup evidence                     |
| MasterData Sync  | Stammdaten workbook, MgmtReports/Kartice exports, MasterSync logs             |
| Auth             | GAS token/session logs, Users/Stammdaten, ErrorLog                            |

---

## 3. Koji ID pratiš

Primarni monitoring ID:

```text
correlationId
```

Obavezni sekundarni ID-jevi po domenu:

```text
SEF: FakturaID, SEFSubmissionID, SEFDocumentId
Faktura: FakturaID, BrojFakture
Novac: NovacID, FakturaID, OtkupID
Otkup: OtkupID, clientRecordID, serverRecordID
Zbirna: clientRecordID, serverRecordID, BrojZbirne
Banka: BankaImportID, NovacID, PartnerID
Backup: BackupFileId / file path / timestamp
MasterData: sync run timestamp, sheet ID, counts
Auth: userId, role, entityID, token hash/session marker, not full token
```

Incident ticket minimum:

```text
Monitoring Component:
Health Status:
EventType:
Severity:
CorrelationId:
EntityType:
EntityId:
Source:
Module/Function:
AlertId:
AuditCritical row: Da/Ne
Canonical source checked:
Business decision:
Technical action:
Resolved by:
Resolved at:
```

---

## 4. Normalan monitoring tok

### 4.1. VBA event emission

`modMonitoring` šalje event u GAS monitoring endpoint.

VBA config dolazi iz `tblSEFConfig`:

```text
MONITORING_ENDPOINT
MONITORING_SECRET
MONITORING_ENV
```

`APP_VERSION` dolazi iz `modConfig.APP_VERSION`. `deviceId` se generiše iz lokalnog workstation/user identiteta.

Normalan payload sadrži:

```text
action
monitoringSecret
environment
source
severity
eventType
userId
role
deviceId
appVersion
module
functionName
entityType
entityId
correlationId
message
payload
```

### 4.2. GAS ingest

VBA koristi public ingest:

```json
{ "action": "monitorPublic" }
```

`monitorPublic` proverava `monitoringSecret` protiv GAS Script Property:

```text
MONITORING_INGEST_SECRET
```

PWA/internal monitoring može koristiti authenticated:

```json
{ "action": "monitor" }
```

### 4.3. Routing u monitoring workbook

`Monitoring.gs` normalizuje event i route-uje:

| Event                 | Ide u tab                    |
| --------------------- | ---------------------------- |
| svaki validan event   | `Events`                     |
| `ERROR` / `CRITICAL`  | `Errors`                     |
| sync/heartbeat        | `SyncStatus`, `UserSessions` |
| `SEF_*`               | `SEFStatus`                  |
| `BACKUP_*`            | `Backups`                    |
| alert-worthy          | `Alerts`                     |
| critical/audit-impact | `AuditCritical`              |
| svaki event           | update `Health` komponente   |

### 4.4. Watchdog

`runMonitoringWatchdog()` treba da radi na schedule-u.

Canonical trigger:

```text
svakih 15 minuta
```

Watchdog proverava:

```text
GAS API watchdog execution
Google Sheets DB / required monitoring tabs accessibility
Auth secret presence
backup freshness
MasterData/Stammdaten sync freshness/failure
SEF unknown/stuck state age
PWA sync/offline queue failures/conflicts/stale pending
recent GAS error spikes
```

### 4.5. Daily summary

`sendDailyMonitoringSummary` treba da radi jednom dnevno.

Koristi se za:

```text
otvoreni alert-i
critical/error count
backup freshness
SEF manual-review/unknown
sync/offline queue health
masterdata sync state
```

---

## 5. Severity i nextAction

### 5.1. Severity

| Severity   | Značenje                                                                | Akcija                                                       |
| ---------- | ----------------------------------------------------------------------- | ------------------------------------------------------------ |
| `INFO`     | normalan signal                                                         | bez akcije osim ako nedostaje očekivani kasniji event        |
| `WARN`     | degradirano stanje ili potrebna pažnja                                  | proveriti domen i pratiti ponavljanje                        |
| `ERROR`    | operacija nije uspela ili sistem ima grešku                             | otvoriti incident ako utiče na rad                           |
| `CRITICAL` | moguć poslovno-pravni, finansijski, data-integrity ili sigurnosni rizik | odmah proveriti canonical source i zaustaviti pogođeni domen |

### 5.2. nextAction

Canonical `nextAction` vrednosti:

```text
WAIT
RETRY
MANUAL_REVIEW
CHECK_SEF_PORTAL
```

Tumačenje:

| nextAction         | Značenje                                                             |
| ------------------ | -------------------------------------------------------------------- |
| `WAIT`             | ne retry odmah; sačekati scheduled refresh/watchdog ili remote state |
| `RETRY`            | retry je dozvoljen posle provere idempotency-ja                      |
| `MANUAL_REVIEW`    | čovek mora pregledati canonical state pre dalje akcije               |
| `CHECK_SEF_PORTAL` | proveriti SEF portal/API pre retry/storno/cancel                     |

---

## 6. Standardni incident flow

### Korak 1: Počni iz Health-a

Zapiši:

```text
Component:
Status:
LastEventType:
LastSeverity:
CorrelationId:
Message:
NextAction:
```

### Korak 2: Pronađi event chain

U `Events` filter:

```text
correlationId = <CorrelationId>
```

Ako nema correlation ID:

```text
eventType + entityId + timestamp window
```

### Korak 3: Proveri Alerts

Ako postoji alert:

```text
AlertId
Status
AlertCode
CreatedAt
LastSeenAt
NextAction
```

Ako je isti alert već otvoren, ne pravi novi incident; update-uj postojeći.

### Korak 4: Proveri AuditCritical

Ako event postoji u `AuditCritical`, ne radi retry dok nije jasno:

```text
šta je canonical stanje
ko je business owner
ko je technical owner
da li je potrebna pravna/finansijska odluka
```

### Korak 5: Proveri canonical source

Pre odluke proveri domain source:

```text
SEF tables / portal
Novac/Faktura/Otkup tables
PWA IndexedDB / Google sheet
Backup folder / Drive history
MasterSync/Stammdaten workbook
GAS execution logs
```

### Korak 6: Klasifikuj

| Signal                               | Kategorija                                | Akcija                                                    |
| ------------------------------------ | ----------------------------------------- | --------------------------------------------------------- |
| Health CRITICAL, canonical source OK | monitoring false-positive ili stale alert | proveri watchdog/dedup/resolution                         |
| Health OK, korisnik ima problem      | monitoring coverage gap                   | proveri domain runbook i dodaj monitoring event ako treba |
| Event ne stiže                       | ingest/config problem                     | proveri endpoint/secret/triggers                          |
| Events stižu, Health se ne menja     | routing/update problem                    | tehnički owner za Monitoring.gs                           |
| Alerts se dupliraju                  | dedupe problem                            | proveri component + alertCode + correlationId             |
| AuditCritical bez owner-a            | process gap                               | dodeli owner-a pre zatvaranja                             |

---

## 7. Recovery scenariji

### 7.1. Monitoring uopšte ne radi

Simptom:

```text
nema novih Events
Health stale
VBA test ne upisuje red
```

Postupak:

1. Proveri `MONITORING_ENDPOINT` u `tblSEFConfig`.
2. Proveri `MONITORING_SECRET` u `tblSEFConfig`.
3. Proveri `MONITORING_ENV`.
4. Proveri GAS Script Properties:

```text
MONITORING_SPREADSHEET_ID
MONITORING_ALERT_EMAIL
MONITORING_INGEST_SECRET
```

5. Proveri da li GAS Web App URL odgovara deployment-u.
6. Pokreni monitoring HTTP test.
7. Ako HTTP nije 200, problem je endpoint/deployment/network.
8. Ako HTTP 200, ali nema reda, problem je workbook ID / permissions / Monitoring.gs routing.

### 7.2. Monitoring secret mismatch

Simptom:

```text
monitorPublic odbija event
HTTP možda 200 sa success=false ili forbidden style response
nema Events reda
```

Postupak:

1. Ne kopirati secret u ticket.
2. Proveriti da `tblSEFConfig.MONITORING_SECRET` odgovara GAS `MONITORING_INGEST_SECRET`.
3. Rotirati secret ako je kompromitovan.
4. Posle promene pokrenuti `TestMonitoring_Config` i `TestMonitoring_HTTP`.
5. Proveriti da debug output ne prikazuje pun secret.

### 7.3. Health je prazan/unknown za GAS API, Auth, Backup ili MasterData Sync

Postupak:

1. Proveri da li postoji `runMonitoringWatchdog` trigger.
2. Pokreni watchdog ručno.
3. Proveri da li monitoring workbook ima sve required tabs.
4. Proveri da Script Properties postoje.
5. Za Backup proveri poslednji `BACKUP_SUCCESS`.
6. Za MasterData proveri `MASTERDATA_SYNC_*` i `STAMMDATEN_SYNC_*` event-e.
7. Ako ručni watchdog popuni Health, problem je trigger setup.

### 7.4. Health je CRITICAL za SEF API

Postupak:

1. Otvori `SEFStatus` za isti `correlationId` / `FakturaID`.
2. Otvori `Alerts` i `AuditCritical`.
3. Proveri canonical SEF tabele:

```text
tblFakture.SEFWorkflowState
tblFakture.SEFStatus
tblFakture.SEFDocumentId
tblSEFSubmission
tblSEFEventLog
```

4. Ako `SEFDocumentId` postoji, prvo ide status refresh / SEF portal check.
5. Ako nema `SEFDocumentId`, proveri da li je lokalno stanje `SEF_SENDING`, `TECH_FAILED`, `UNKNOWN` ili manual-review.
6. Ne retry slanje fakture dok nije jasno da remote dokument ne postoji.
7. Pređi na SEF production runbook.

### 7.5. `SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING`

Ovo je visokorizičan scenario.

Postupak:

1. Ne ponavljati submit automatski.
2. Proveriti da li postoji `SEFSubmissionID`.
3. Proveriti da li postoji `SEFDocumentId`.
4. Proveriti `tblSEFEventLog` i SEF portal.
5. Ako postoji remote dokument, lokalno stanje treba završiti kroz refresh/recovery, ne novi submit.
6. Ako ne postoji remote dokument, tehnički owner i finansijsko-pravni owner odlučuju retry.
7. Alert ostaje otvoren dok nije dokumentovano konačno stanje.

### 7.6. Backup health je WARN/CRITICAL

Postupak:

1. Otvori `Backups` tab.
2. Nađi poslednji `BACKUP_SUCCESS`.
3. Proveri:

```text
BackupType
BackupFileId / path
BackupLocation
Status
ErrorMessage
Timestamp
```

4. Proveri stvarni `Backup\` folder ili Drive backup lokaciju.
5. Ako monitoring nema event, ali backup postoji, problem je monitoring emission/routing.
6. Ako backup ne postoji ili je star, pokrenuti backup procedure/runbook.
7. Ne zatvarati alert dok ne postoji nova uspešna backup potvrda.

### 7.7. MasterData Sync je WARN/CRITICAL

Postupak:

1. Proveri `MASTERDATA_SYNC_SUCCESS/FAIL` event-e.
2. Proveri `STAMMDATEN_SYNC_SUCCESS/FAIL` event-e.
3. Za Stammdaten proveri da li je export svih 13 tabova uspeo:

```text
Kooperanti
Kulture
Parcele
Config
Users
Fakture
FakturaStavke
SaldoOMDetail
Stanice
Kupci
Vozaci
Artikli
MagacinKoop
```

4. Ako je partial export, tretirati kao fail.
5. Proveri Google OAuth/config/PWA folder ID.
6. Pređi na MasterSync/Stammdaten runbook.

### 7.8. PWA Sync / Offline Queue je WARN/CRITICAL

Postupak:

1. Otvori `SyncStatus`.
2. Filter po role/entity/correlationId.
3. Proveri broj pending/syncing/failed/offline queue record-a.
4. Za konkretan incident proveri lokalni IndexedDB na uređaju:

```text
clientRecordID
syncStatus
lastServerStatus
lastSyncError
syncAttempts
```

5. Ako je `syncing` stale, pređi na PWA offline sync runbook.
6. Ako su duplikati, pređi na submit-lock/duplicate runbook.
7. Ako je auth 401/403, pređi na GAS auth runbook.

### 7.9. Alerts se pune istim problemom

Postupak:

1. Grupisati po:

```text
component
alertCode
correlationId
entityId
```

2. Ako je isti correlation/component/code, treba da bude jedan unresolved alert.
3. Ako se pravi više alert-a za isti incident, tehnički owner proverava dedupe key.
4. Ako su različiti entity/correlation, možda je stvarno više incidenata.
5. Ne gasiti alert generation globalno osim ako monitoring ugrožava rad.

### 7.10. Nema alert-a, ali postoji CRITICAL event

Postupak:

1. Proveri `Events` i `Errors`.
2. Proveri routing u `Monitoring.gs`.
3. Proveri da li `alert-worthy` logika pokriva taj `eventType`.
4. Ako je business-impacting, ručno otvoriti incident i dodati eventType u alert mapping.
5. Ako je namerno bez alert-a, dokumentovati zašto.

### 7.11. AuditCritical ima event, ali Alerts nema otvoren alert

Postupak:

1. Tretirati kao process gap.
2. Proveri da li je event audit-only ili mora imati alert.
3. Ako je SEF, bank, payment, backup critical, mora postojati owner i nextAction.
4. Ručno otvoriti incident ako alert nije kreiran.
5. Tehnički owner popravlja mapping.

### 7.12. Monitoring potencijalno loguje secret/token/XML/PDF

Postupak:

1. Ne deliti screenshot javno dok se ne proveri sadržaj.
2. Proveri `Events.payload`, `Errors.details`, debug output.
3. Zabranjeni podaci:

```text
monitoring secret
Google access/refresh token
SEF API key
full UBL XML
full PDF/base64
full raw SEF response body
password/PIN/Authorization header
```

4. Ako je procurelo, tretirati kao security incident.
5. Rotirati pogođene secrets/tokens.
6. Očistiti ili ograničiti pristup monitoring workbook-u po security odluci.
7. Popraviti sanitization/truncation.

### 7.13. Daily summary ne stiže

Postupak:

1. Proveri `MONITORING_ALERT_EMAIL` Script Property.
2. Proveri daily trigger za `sendDailyMonitoringSummary`.
3. Pokreni funkciju ručno u GAS.
4. Proveri Apps Script executions/mail quota.
5. Proveri spam/inbox pravila.
6. Ako summary failuje, Health/Alerts i dalje ostaju source za incident rad.

---

## 8. Setup checklist

Monitoring nije production-ready dok ovo nije potvrđeno:

```text
[ ] OtkupApp_Monitoring_PROD postoji
[ ] Health tab postoji
[ ] Events tab postoji
[ ] Errors tab postoji
[ ] SyncStatus tab postoji
[ ] SEFStatus tab postoji
[ ] UserSessions tab postoji
[ ] Backups tab postoji
[ ] Alerts tab postoji
[ ] AuditCritical tab postoji
[ ] GAS Script Property MONITORING_SPREADSHEET_ID setovan
[ ] GAS Script Property MONITORING_ALERT_EMAIL setovan
[ ] GAS Script Property MONITORING_INGEST_SECRET setovan
[ ] tblSEFConfig.MONITORING_ENDPOINT setovan
[ ] tblSEFConfig.MONITORING_SECRET setovan
[ ] tblSEFConfig.MONITORING_ENV setovan na DEV/PROD
[ ] GAS Web App redeploy urađen
[ ] runMonitoringWatchdog trigger instaliran
[ ] sendDailyMonitoringSummary trigger instaliran
[ ] TestMonitoring_All prolazi
[ ] Ručni end-to-end event vidljiv u Events
[ ] Health red za VBA Client ažuriran
[ ] Debug output ne prikazuje secret/body
```

---

## 9. Test/smoke procedure

### 9.1. VBA monitoring smoke

Pokrenuti:

```text
TestMonitoring_All
TestMonitoring_Config
TestMonitoring_HTTP
TestMonitoring_ErrorEvent
TestMonitoring_SEFUnknown
TestMonitoring_BackupSuccess
TestMonitoring_BackupFail
```

Uspešan HTTP test očekuje:

```text
HTTP Status = 200
success = true
eventId postoji
timestamp postoji
severity postoji
component = VBA Client
```

### 9.2. GAS watchdog smoke

Pokrenuti ručno:

```text
runMonitoringWatchdog()
```

Očekivanje:

```text
Health popunjen za GAS API
Health popunjen za Google Sheets DB
Health popunjen za Auth
Health popunjen za Backup
Health popunjen za MasterData Sync
```

### 9.3. Routing smoke

Poslati test event-e:

```text
INFO VBA_TEST_EVENT -> Events + Health
ERROR TEST_ERROR -> Events + Errors + Alerts ako alert-worthy
SEF_STATUS_PENDING -> Events + SEFStatus + Health
BACKUP_SUCCESS -> Events + Backups + Health
CRITICAL TEST_AUDIT -> Events + Errors + Alerts + AuditCritical
```

### 9.4. Security smoke

U test payload ubaciti lažni sensitive key i proveriti da monitoring ne loguje pun sadržaj:

```text
monitoringSecret
access_token
refresh_token
sefApiKey
Authorization
base64/pdf body
```

Očekivanje:

```text
redacted/truncated
nema punog secret-a
debug output nema full JSON body
```

---

## 10. Event coverage mapa

### 10.1. App lifecycle

Event-i:

```text
VBA_APP_OPEN
VBA_STARTAPP_START
VBA_STARTAPP_SUCCESS
JOURNAL_RECOVERY_WARN
Monitor_Error za startup failure
```

Ako `JOURNAL_RECOVERY_WARN` postoji, preći na Startup/Journal recovery runbook.

### 10.2. SEF

Startup recovery:

```text
SEF_STARTUP_RECOVERY_START
SEF_RECOVERY_INVOICE_FOUND
SEF_RECOVERY_INVOICE_SUCCESS
SEF_RECOVERY_INVOICE_FAIL
SEF_STARTUP_RECOVERY_SUCCESS
SEF_STARTUP_RECOVERY_FAIL
```

Submit:

```text
SEF_SEND_START
SEF_SEND_ACCEPTED
SEF_SEND_SUCCESS
SEF_SEND_REJECTED
SEF_SEND_FAIL
SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
```

Refresh:

```text
SEF_STATUS_ACCEPTED
SEF_STATUS_REJECTED
SEF_STATUS_PENDING
SEF_STATUS_TERMINAL
SEF_STATUS_UPDATE
SEF_STATUS_REFRESH_FAIL
SEF_STATUS_REFRESH_EXCEPTION
SEF_REFRESH_PENDING_START
SEF_PENDING_REFRESH_INVOICE_FAIL
SEF_REFRESH_PENDING_SUMMARY
SEF_REFRESH_PENDING_FAIL
```

High-risk event-i:

```text
SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
SEF_RECOVERY_INVOICE_FAIL
SEF_STATUS_REFRESH_EXCEPTION
unknown/manual-review states
```

### 10.3. Finance / Novac / Faktura

Event-i:

```text
FAKTURA_CREATE_SUCCESS
FAKTURA_CREATE_FAIL
NOVAC_SAVE_SUCCESS
NOVAC_SAVE_FAIL
AVANS_APPLY_TO_FAKTURA_FAIL
Monitor_Error
```

Ako monitoring prijavi fail, canonical source je `tblFakture`, `tblFakturaStavke`, `tblPrijemnica`, `tblNovac`.

### 10.4. Otkup

Event-i:

```text
OTKUP_SAVE_SUCCESS
OTKUP_SAVE_FAIL
OTKUP_MULTI_SAVE_SUCCESS
OTKUP_MULTI_SAVE_FAIL
Monitor_Error
```

Ako otkup failuje, proveri i `tblAmbalaza`, jer otkup može imati packaging side-effect.

### 10.5. Dokumentni chain

Event-i su fail-only:

```text
DOKUMENT_SAVE_FAIL
Monitor_Error
```

Pokriveni tokovi:

```text
SaveOtpremnica_TX
SaveOtpremnicaMulti_TX
SaveZbirna_TX
SaveZbirnaMulti_TX
SavePrijemnica_TX
SavePrijemnicaMulti_TX
```

Normalni success dokument-chain event-i nisu namerno noisy. Ako nema success event-a, to nije dokaz da save nije uspeo.

### 10.6. Banka

Event-i:

```text
BANKA_MAP_SUCCESS
BANKA_MAP_FAIL
BANKA_IMPORT_SKIP
BANKA_AUTOMAP_ALL_START
BANKA_AUTOMAP_ALL_SUMMARY
BANKA_AUTOMAP_ALL_FAIL
```

Canonical wrappers:

```text
AutoMapBankaImportRow_TX
MapBankaImportAsKupac_TX
MapBankaImportAsKooperant_TX
MapBankaImportAsOM_TX
MapBankaImportAsKooperantBlock_TX
MapBankaImportAsKooperantBlockManual_TX
SkipBankaImportRow_TX
AutoMapAllBankaImport_TX
```

### 10.7. MasterData / Stammdaten

Event-i:

```text
MASTERDATA_SYNC_SUCCESS
MASTERDATA_SYNC_FAIL
STAMMDATEN_SYNC_SUCCESS
STAMMDATEN_SYNC_FAIL
```

Pravilo:

> MasterData Sync health se zatvara realnim sync/export signalom, ne placeholder health redom.

### 10.8. Backup

Event-i:

```text
BACKUP_SUCCESS
BACKUP_FAIL
```

Backups tab treba da sadrži:

```text
BackupType
SourceSpreadsheetId
BackupFileId
BackupLocation
RowsCount
Checksum
Status
DurationMs
ErrorMessage
```

---

## 11. Dozvoljene i zabranjene akcije

### Operator sme sam

* pogledati Health/Alerts/Events;
* prikupiti correlationId i entityId;
* pokrenuti definisane monitoring testove ako su dostupni;
* proveriti da li postoji otvoren alert;
* proveriti canonical source za domen ako ima prava;
* eskalirati sa kompletnim monitoring paketom.

### Operator ne sme sam

* zatvoriti CRITICAL alert bez canonical provere;
* retry-ovati SEF submit samo zato što monitoring kaže fail;
* brisati Events/Errors/AuditCritical redove;
* menjati Script Properties;
* kopirati monitoring secret/token u ticket/chat;
* ignorisati AuditCritical.

### Tehnički owner odlučuje

* promenu `Monitoring.gs` routing-a;
* trigger setup/reinstall;
* monitoring workbook schema repair;
* secret rotation;
* redaction/sanitization fix;
* dedupe alert logiku;
* deployment URL/config fix.

### Business/domain owner odlučuje

* da li SEF manual review prelazi u retry/cancel/storno;
* da li bank mapping failure utiče na finansije;
* da li document/faktura/novac incident zahteva korekciju;
* da li backup failure blokira rad;
* da li AuditCritical event ima pravno/računovodstveni uticaj.

### Security owner odlučuje

* šta raditi ako monitoring procure secrets/tokens;
* da li je `monitorPublic` secret kompromitovan;
* da li se monitoring workbook access ograničava;
* da li incident ide kao security incident.

---

## 12. Checklist za zatvaranje incidenta

```text
[ ] Identifikovan component
[ ] Identifikovan eventType
[ ] Identifikovan severity
[ ] Identifikovan correlationId
[ ] Proveren Events chain
[ ] Proveren Errors ako severity ERROR/CRITICAL
[ ] Proveren Alerts status
[ ] Proveren AuditCritical ako postoji
[ ] Proveren canonical source za pogođeni domen
[ ] Ako je SEF, proveren SEFDocumentId / SEF portal / tblSEFSubmission / tblSEFEventLog
[ ] Ako je Backup, proveren stvarni backup file/location
[ ] Ako je MasterData, proveren sync/export outcome
[ ] Ako je PWA Sync, proveren IndexedDB/GAS/Google trag
[ ] Ako je auth/config, proveren endpoint/secret bez deljenja secret-a
[ ] Ako je security leak, pokrenut security flow
[ ] Dodeljen owner
[ ] Alert resolved samo posle potvrde
[ ] Korisnik/operator obavešten
```

---

## 13. Primeri odluke

### Primer A: Monitoring ne šalje event-e iz VBA

Zaključak: verovatno config, endpoint, secret ili deployment.
Akcija: proveriti `tblSEFConfig` keys, GAS Script Properties, Web App URL i `TestMonitoring_HTTP`. Business operacije se ne smatraju neuspelim samo zato što monitoring ne radi.

### Primer B: SEF API Health je CRITICAL sa `SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING`

Zaključak: moguć remote/local split-brain.
Akcija: ne retry submit. Proveriti `SEFSubmissionID`, `SEFDocumentId`, `tblSEFEventLog` i SEF portal; SEF/legal owner odlučuje.

### Primer C: Backup CRITICAL, ali backup fajl postoji

Zaključak: monitoring routing/emission problem ili stale Health.
Akcija: proveriti Backups tab, stvarni folder, watchdog. Alert se zatvara tek kad je Health ažuriran ili incident dokumentovan.

### Primer D: MasterData Sync health prazan

Zaključak: watchdog ili `MASTERDATA_SYNC_*` / `STAMMDATEN_SYNC_*` event-i ne stižu.
Akcija: pokrenuti watchdog i Stammdaten/MasterSync smoke; proveriti Script Properties i workbook routing.

### Primer E: Alerts se dupliraju

Zaključak: alert dedupe key ne radi ili svaki event ima novi correlationId.
Akcija: grupisati po component/alertCode/correlationId; tehnički owner popravlja dedupe ili correlation policy.

### Primer F: Monitoring payload sadrži full UBL XML

Zaključak: security/privacy incident.
Akcija: ograničiti pristup workbook-u, rotirati relevantne secrets ako su procureli, popraviti redaction i ukloniti/ograničiti osetljiv sadržaj po security odluci.

---

## 14. Poznate production rupe koje treba zatvoriti

1. Dodati dashboard “Open Alerts by severity and age”.
2. Dodati owner/assignee kolonu u Alerts ako ne postoji.
3. Dodati alert SLA: koliko dugo WARN/CRITICAL sme da stoji otvoren.
4. Dodati explicit resolved workflow: ko, kada, zašto, canonical source checked.
5. Dodati monitoring schema validator za sve tabove.
6. Dodati automatic trigger verifier za watchdog i daily summary.
7. Dodati rate limit za noisy eventType.
8. Dodati canonical correlationId generator po business operaciji.
9. Dodati PWA role-level heartbeat / sync health standardizaciju ako nije kompletna.
10. Dodati monitoring access policy: ko sme da vidi Events/Errors/AuditCritical.
11. Dodati redaction tests za svaki zabranjeni sensitive field.
12. Dodati cross-link iz Alert-a ka domain runbook-u.
13. Dodati stale alert auto-reminder, ne auto-resolve.
14. Dodati evidence pack export za incident: Health row + Events chain + Alerts + AuditCritical + canonical source references.

Do tada važi konzervativno pravilo:

> Monitoring govori gde da gledaš, ali ne menja business istinu. Health/Alerts pokreću istragu; konačna odluka dolazi iz canonical tabela, SEF portala/API-ja, backup fajlova, IndexedDB/Google sync traga i odgovornog owner-a.
