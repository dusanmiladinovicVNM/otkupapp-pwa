# AgriX / OtkupApp Release Gates

**Purpose:** Operational compile, smoke, regression, launch and production readiness checks extracted from the Architecture Reference and Changelog.

---

## 0. Purpose
This document owns detailed verification steps. `ARCHITECTURE_REFERENCE.md` keeps only the high-level requirement that the relevant gates must pass.

---

## 1. Current Mandatory Gates

These are the current top-level gates for a v6.23 production handoff. Detailed domain gates below are selected based on the changed surfaces.

### 1.1 Always Required

- VBA compile for all touched desktop modules and forms.
- `RunProductionHealthCheck` clean for the specific production workbook before workbook launch.
- Review of unresolved `NEEDS REVIEW` items before handoff.
- Confirmation that smoke/test data did not leave active broken references.
- Confirmation that accepted warnings/risks are recorded in `ROADMAP.md` or AR `Current Known Risks`.

### 1.2 Required When Desktop Business Logic Changes

- `RunBusinessFlowProSuite` or successor.
- Relevant domain smoke: Faktura, Novac, Dokumenta, Storno, BankaImport, BankaMapiranje, Agrohemija, SEF.
- Transaction rollback negative tests for changed write paths.
- AutoSave/journal/backup behavior where transaction code changed.

### 1.3 Required When GAS / Google Sheets Changes

- `runGasRouteHealthCheck()` or successor.
- `runGasSmokeSuite()` or successor.
- Sheet header/schema drift checks for affected sheets.
- Role/entity authorization checks for changed write endpoints.
- `SyncControl / MASTER_SYNC_LOCK` checks where MasterSync/sync paths changed.

### 1.4 Required When PWA Changes

- Role smoke for affected roles: Otkupac, Kooperant, Vozač, Management.
- Offline queue / stale `syncing` recovery smoke when sync code changes.
- Render dedupe smoke when merge/render code changes.
- Submit lock smoke when save flows change.
- Service-worker `CACHE_NAME` bump when critical cached assets change.

### 1.5 Required When Monitoring / Security Changes

- Monitoring config and HTTP ingest smoke.
- Monitoring workbook tab/header check.
- Watchdog and daily summary trigger check where applicable.
- SEF HTTPS-only negative config check when SEF/HTTP/config changes.
- Secret/redaction check where monitoring, SEF, Google auth or logging changes.

### 1.6 Required When External Side Effects Change

- Google writeback failure behavior check.
- BankaImport file move timing check.
- PDF extraction tool path/exit-code check.
- Confirmation that external side effects are not assumed rollback-able unless explicitly designed so.


---

## 2. VBA Gates

### 2.1 Compile Gate

Acceptance:

- all touched VBA modules compile without reserved-name, missing-procedure or missing-constant errors;
- `clsTransaction`, `modDataAccess`, `modSchemaGuard`, `modParse`, `modComboBinding`, `modJournal`, `modLogError`, `modMonitoring` and domain modules referenced by this release are available;
- `RequireColumnIndex`, `RequireColumns` and `RequireUpdateCell` compile and are visible to hardened modules;
- no form event handler references deleted modules or stale form names;
- `Attribute VB_Name` headers are preserved in exported `.bas`/`.cls`/`.frm` modules.


### 2.1a App Lifecycle Gate

Acceptance:

- `Workbook_Open` delegates to `modMain.StartApp` and contains no direct business import/sync side effects;
- startup error handling restores `Application.Visible = True` before user-facing failure output;
- `StartApp` validates canonical tables and runs startup housekeeping without blocking the shell unnecessarily;
- `ImportBankaInbox_TX()` is not run automatically from boot;
- `frmOtkupAPP` close and `Workbook_BeforeClose` both route through `ShutdownApp`;
- normal shutdown writes `LogAppShutdown` and restores Excel visibility.

### 2.1b Setup / LocalConfig Gate

Acceptance:

- `LocalConfig` sheet and `tblLocalConfig` exist with `Kljuc | Vrednost | Opis`;
- `GetLocalConfigValue` and `SetLocalConfigValue` are public and callable from parser/setup modules;
- `PDFTOTEXT_EXE_PATH` is initialized or reported as a health warning;
- `tblConfig` is not used for workstation-local tool paths;
- `tblSEFConfig` remains the owner for SEF/monitoring runtime keys;
- clean-machine setup creates expected folders: `Backups`, `Logs`, `Journal`, `Export`, `Temp`, `Secrets`, `Bank_Izvodi`;
- Excel trusted location and workbook unblock/certificate steps are documented or automated by setup.

### 2.1c AutoSave / Journal / Backup Gate

Acceptance:

- successful `clsTransaction.CommitTx` triggers `AutoSaveAfterCommit(sourceName)` best-effort;
- AutoSave handles read-only and unsaved/no-path workbooks without raising into committed business callers;
- AutoSave restores `Application.DisplayAlerts` on every exit path;
- `AppendRow()` writes a journal row through `WriteJournalRow()`;
- startup backup creates a dated workbook copy under `Backup/`;
- journal/backup/log purge routines are best-effort and do not block normal workbook use;
- `CheckJournalForRecovery()` surfaces advisory warnings instead of attempting unsafe automatic replay.

### 2.2 BusinessFlowPro Gate


Acceptance:

- End-to-end `Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF` smoke runs against non-stornirano active rows.
- Dual-class `Otpremnica`, `Zbirna` and `Prijemnica` flows save atomically through their multi-row transaction wrappers.
- `tblPrijemnica.PrimjenicaID` / `PrijemnicaID` row identity is unique per physical row; `BrojPrijemnice` groups class rows. *(NEEDS REVIEW: confirm spelling in any local gate script output; schema uses `PrijemnicaID`.)*
- Faktura creation from both class rows succeeds and writes one faktura line per selected prijemnica tuple.
- Duplicate faktura prevention works for already-fakturisana prijemnica rows.
- Invalid input rolls back the full business operation.
- Stornirano rows are excluded from active document reads and reporting caches.
- Strict `BrojZbirne` auto-link/cascade does not create cross-zbirna links.
- Cross-zbirna audit remains green after import/linking.
- Ambalaža side effects are created/reversed only through the owning transaction/storno flow.
### 2.3 Faktura Gate

Acceptance:

- `CreateFaktura_TX` snapshots fakture, stavke, prijemnice and novac before creation/allocation work.
- Total amount equals the sum of selected `tblPrijemnica.Kolicina × tblPrijemnica.Cena`.
- One `tblFakturaStavke` row is created per selected prijemnica tuple and stores `PrijemnicaID`, `Kolicina`, `Cena`, `Klasa` and `BrojPrijemnice`.
- Selected prijemnice are marked `Fakturisano = "Da"` and linked with `FakturaID`.
- Duplicate already-fakturisana prijemnica selection is rejected.
- Buyer avans auto-application does not break transaction rollback.
- Storno/relink scenario: replacement prijemnica with same `BrojPrijemnice + Klasa` relinks orphaned stavka exactly once.
- `CreateFaktura_TX` emits `FAKTURA_CREATE_SUCCESS` / `FAKTURA_CREATE_FAIL` / `Monitor_Error` as applicable.

### 2.4 Novac Gate
### 2.5 Document Flow Non-Storno Gate

Acceptance:

- `SaveOtpremnica` fails if `GetNextID` does not return an `OtpremnicaID`.
- `SaveOtpremnicaMulti_TX`, `SaveZbirnaMulti_TX` and `SavePrijemnicaMulti_TX` rollback all rows from a dual-class save on forced failure.
- `SavePrijemnica` uses the row returned by `AppendRow`; no required post-append lookup is needed for normal write completion.
- `GetVozacDokumenta` and `BuildZbirnaVrstaCache` exclude stornirano otpremnica rows.
- `RelinkFakturaStavke(newPrijemnicaID, brojPrijemnice, klasaFilter)` relinks only the targeted class.
- `AutoLinkOtkupOtpremnica_TX` requires exactly one matching `OtkupID` before writing `OtpremnicaID`.
- Ambiguous auto-link candidates remain unresolved for manual review.
- `TrackAmbalaza` rejects negative quantity and invalid `Smer`.
- `TrackAmbalaza` treats quantity `0` as legal no-op.
- `GetAmbalazeStanje` treats `Ulaz` as positive and `Izlaz` as negative; unknown `Smer` fails.
- `GetVozacAmbSaldo` works with only `datumOd`, only `datumDo`, both, or neither.

### 2.6 Storno Gate

Acceptance:

- Missing target ID fails and rolls back.
- Duplicate target ID fails and rolls back for ID-based storno operations.
- Already-stornirano target is rejected by `CanStorno()`.
- Required schema reads use `RequireColumnIndex`.
- Required writes use `RequireUpdateCell`.
- `_TX` wrappers rollback hard failures.
- `_TX` wrappers emit success/failure monitoring best-effort.
- `StornoOtkup` marks otkup stornirano, stornira related ambalaža and removes otkup links from `tblNovac`.
- `StornoOtpremnica` marks otpremnica stornirano and stornira related ambalaža.
- `StornoZbirna` marks all active rows with the same `BrojZbirne` stornirano and does not reclaim sequence numbers.
- `StornoPrijemnica` marks prijemnica stornirano, resets faktura linkage when applicable, marks orphaned faktura/stavka state when needed and stornira related ambalaža.
- `StornoFaktura` marks faktura/stavke stornirano, releases linked prijemnice, removes faktura links from `tblNovac` and updates faktura status.
- `StornoNovac` marks novac stornirano and recomputes faktura status when relevant.

### 2.7 BankaImport Gate
### 2.8 BankaMapiranje Gate
### 2.9 SEF Gate
### 2.10 ProductionHealthCheck

Acceptance:

- `RunProductionHealthCheck` completes without failures before declaring a workbook production-launch-ready;
- smoke/regression fixtures are rolled back, stornirano-marked or removed from active production data;
- legacy demo/test references do not leave active broken links;
- BankaImport, SEF, document-chain, finance and master-data health checks do not report active blocking issues;
- warnings that remain accepted risks are explicitly documented in `ROADMAP.md` or `Current Known Risks`.


---

## 3. GAS Gates

### 3.1 Route Health Check

Required checks:

- run `runGasRouteHealthCheck()` where available;
- verify every active action has a real handler;
- verify intentionally disabled actions such as `saveOtkupniListPdf` return `FEATURE_DISABLED`;
- verify disabled actions are not counted as active healthy handlers;
- verify public/read bridge actions are explicitly listed and reviewed.

### 3.2 Auth / Authorization Check

Required checks:

- `login` succeeds with valid test credentials and fails with invalid PIN;
- unauthenticated protected actions return structured unauthorized failure;
- wrong-role actions return structured forbidden failure;
- entity mismatch fails for Otkupac, Kooperant and Vozac write/read scopes;
- Management-only actions reject non-Management callers;
- `updateKamionStatus` allows Vozac only for own status and Management for any driver;
- `saveFiskalniMapiranje` and `createArtikal` are Management-only;
- `saveParcelPolygon` deployed auth state is verified and reconciled with AR `NEEDS REVIEW` marker.

### 3.3 Token / Session Maintenance Check

Required checks:

- tokens are cached under `TOKEN_<token>`;
- valid fallback token properties can restore cache state;
- expired/malformed token properties are rejected/deleted;
- failed login throttle blocks repeated failures;
- `setupTokenPurgeTrigger()` exists or maintenance trigger is installed;
- `LoginLog` appends success/failure attempts without exposing PINs.

### 3.4 Sync Endpoint Check

Required checks:

- `sync` writes OTK records idempotently by `ClientRecordID`;
- `syncZbirna` writes VOZ records idempotently by `ClientRecordID`;
- `syncTretman` writes treatment records idempotently;
- `syncTrosak` writes/updates `TROSKOVI-<KooperantID>` and returns `buildBatchSyncResponse(results)`;
- `syncOprema` keeps the same batch/idempotency standard;
- HTTP 200 with empty body is not accepted as success by PWA smoke;
- mixed success/failure returns `PARTIAL_FAILURE`;
- all failed batch returns `BATCH_FAILED`;
- terminal states such as `Synced>Master`, `Duplicate` and `SyncError:*` are not reset by retry.

### 3.5 Schema Drift Check

Required checks:

- empty sheets may receive canonical headers;
- existing sheets with missing required headers fail with `SCHEMA_DRIFT`;
- existing sheets with incompatible headers fail clearly;
- runtime code does not silently append guessed columns to active sync sheets;
- schema repair remains manual/operator-controlled.

### 3.6 Master-Sync Soft-Lock Check

Required checks:

- `getMasterSyncState` reads `Stammdaten / SyncControl`;
- active `MASTER_SYNC_LOCK = YES` blocks write dispatch;
- public/read actions and login remain available where intended;
- blocked writes return retryable `MASTER_SYNC_ACTIVE` or equivalent;
- PWA keeps local pending records retryable;
- stale lock behavior is visible and recoverable.

### 3.7 Management / Dispatch / Agro-Izdavanje Check

Required checks:

- `saveWarRoomDemand`, `removeWarRoomDemand`, `updateDemandPrimljeno` are Management-only and locked;
- `saveDispecer`, `updateDispecer`, `removeDispecer` are Management-only and locked;
- `getDispecer` returns today demand plus active plans;
- `updateKamionStatus` upserts one row per `VozacID`;
- `saveIzdavanje` persists an `IZD-*` row and serializes `stavke` safely;
- ambiguous retry of `saveIzdavanje` remains documented as client-lock mitigation / roadmap server idempotency.

### 3.8 GIS / Meteo / Fiskalni Endpoint Check

Required checks:

- public geo/meteo reads return bounded read payloads only;
- `getParcelMeteo` uses cached `MeteoLatest` before live fallback;
- `scheduledMeteoFetch` refreshes `MeteoLatest` and appends `MeteoHistory`;
- fiscal parse endpoints require Kooperant/Management auth;
- `saveFiskalni` enforces kooperant entity scope;
- `saveFiskalni` rejects duplicate `VerificationUrl`;
- `saveFiskalniMapiranje` and `createArtikal` reject Kooperant callers.

### 3.9 Monitoring Ingest Check

Required checks:

- `logClientError` works as a pre-auth exception;
- valid token enriches entity context where present;
- payloads are truncated/redacted before `ErrorLog` append;
- logger failure does not break normal app flow;
- `monitorPublic` validates monitoring secret;
- authenticated `monitor` path follows the monitoring contract in section 15 of AR.

### 3.10 ErrorLog Check

Required checks:

- `ErrorLog` exists or is lazily created;
- columns match `Timestamp | Source | Action | Message | Details | EntityID | Severity`;
- timeout-like messages classify as warning where implemented;
- retention cleanup removes old rows;
- logger/purge failures remain non-blocking.

### 3.11 Disabled Endpoint Check

Required checks:

- disabled endpoints return explicit disabled responses;
- missing handlers do not crash route dispatch;
- route health does not report intentionally disabled endpoints as active implemented handlers;
- `syncTrosak` is confirmed active and no longer treated as disabled.

---

## 4. PWA Gates

### 4.1 App Shell / Cache Gate

Acceptance:

- `index.html` loads `src/js/services/db.js` only once.
- `sw.js` `CACHE_NAME` is bumped after critical runtime JS changes.
- Critical runtime files are present in the service-worker asset list.
- On a test field device, reload/force update confirms the new service worker and changed JS are active.
- Leaflet marker image assets are cached when map/offline marker surfaces are launch-relevant.

### 4.2 Otkupac Smoke

Acceptance:

- Create new otkup while online and verify local save.
- Create new otkup while offline and verify record remains pending.
- Restore online and trigger `syncQueueSafe('manual')` / post-save sync.
- Verify pending row uploads and moves to synced on valid backend confirmation.
- Verify missing backend per-record result returns the row to `pending` with diagnostics.
- Verify sync badge states: `OFFLINE`, `SYNC...`, `ČEKA: n`, `ONLINE`.
- Double-tap save creates one local record due to `withSubmitLock('otkup:save', ...)`.

### 4.3 Kooperant Smoke

Acceptance:

- Treatment save/sync works through `syncTretmani()`.
- Expense save/sync works through `syncTroskovi()` and GAS `syncTrosak`.
- `syncKooperantNow()` returns one normalized aggregate result for both stores.
- Both child stores returning `no-pending` still produces canonical role result.
- Offline treatment/expense remains pending and syncs after reconnect.
- Treatment/expense history render passes through dedupe before display.
- Stale `syncing` rows in `tretmani` and `troskovi` recover to `pending` at bootstrap.
- Double-tap `agroSaveTretman()` is protected by `withSubmitLock('agro:tretman:save', ...)`.

### 4.4 Vozač Smoke

Acceptance:

- `getVozacOtkupi` loads assigned otkupi.
- `confirmZbirna()` creates local `zbirne` record with `clientRecordID`, technical `serverRecordID` field and separate `brojZbirne`.
- `brojZbirne` is visible immediately before server/master sync.
- Two zbirne for same driver/day produce `x/ddmmyy` then `x/ddmmyy-2`.
- Soft-delete/storno does not reclaim the sequence number.
- `syncZbirne()` syncs pending zbirna through `syncZbirna`.
- Backend statuses `duplicate`, `existing`, `inserted`, `updated`, `synced` all confirm local success.
- Double-tap `confirmZbirna()` produces one local record due to `withSubmitLock('zbirna:confirm', ...)`.

### 4.5 Management Smoke

Acceptance:

- Management session loads dashboard/overview data.
- Management role sync returns canonical `no-sync-for-role` where no local sync store applies.
- Management navigation does not trigger low-level sync internals.
- Management app shell does not access Kooperant-only `window.agroState` paths.

### 4.6 Offline / Sync Recovery Smoke

Acceptance:

- Force a record in each relevant store to `syncStatus = 'syncing'` with stale timestamp.
- Reload app and verify bootstrap calls `recoverStaleSyncingForCurrentRole('bootstrap')`.
- Otkupac store recovers stale rows to `pending`.
- Kooperant `tretmani` / `troskovi` recover stale rows to `pending`.
- Vozač `zbirne` recovers stale rows to `pending`.
- Recovered rows set `lastServerStatus = 'stale-syncing-recovered'`.
- Fresh in-flight rows are not blindly recovered.

### 4.7 Submit Lock Smoke

Acceptance:

- `withSubmitLock` exists globally.
- `saveOtkup()` wraps `saveOtkupUnlocked()`.
- `confirmZbirna()` wraps `confirmZbirnaUnlocked()`.
- `agroSaveTretman()` wraps `agroSaveTretmanUnlocked()`.
- Active lock returns early and optionally shows already-saving feedback.
- Matching action buttons are disabled during the save and restored in `finally`.
- `confirmZbirnaUnlocked()` calls `syncQueueSafe('post-save')` and does not call low-level `syncZbirne()` directly.

### 4.8 Client Error Reporting Smoke

Acceptance:

- `reportClientError(error, context)` exists in `src/js/utils/async.js`.
- `safeAsync(...)` catch paths report best-effort.
- sync-engine exception paths report best-effort.
- bootstrap/startup catch paths report best-effort.
- global `window.error` and `window.unhandledrejection` are wired.
- GAS `logClientError` writes a row to `ErrorLog`.
- Logging failure does not break the app runtime.
- Payload is bounded: no secrets, no tokens, no large sensitive bodies.

### 4.9 Master-Sync Guard / Soft-Lock Smoke

Acceptance:

- VBA full-cycle sync writes `MASTER_SYNC_LOCK = YES` in `Stammdaten / SyncControl` during sync.
- GAS `getMasterSyncState` exposes the lock to PWA.
- GAS write actions are blocked while lock is active.
- PWA shows operator-visible lock/overlay/waiting state.
- PWA local capture can continue for ordinary otkup workflows.
- `MASTER_SYNC_ACTIVE` does not become a permanent sync error.
- Pending local rows retry after lock release.
- Stale lock timeout/manual recovery behavior is visible and controlled.

### 4.10 Render Dedupe Smoke

Acceptance:

- `dedupeRecordsForRender(records, aliasesFn?)` exists in `src/js/utils/merge.js`.
- A local pending row with same `clientRecordID` as server synced row wins.
- Synced/non-priority duplicates choose the newest timestamp.
- Rows without identity aliases are preserved.
- Otkup queue, Otkup pregled, Vozač zbirna pregled, Kooperant tretman history, Otkup otprema overview and otprema assign state call dedupe before render.

### 4.11 Business Date Smoke

Acceptance:

- Business date-only fields use `getTodayIsoDate()`, `getRelativeIsoDate(...)`, `toIsoDateOnly(...)`, `fmtDate(...)` or `localIsoDateFromDate(...)`.
- Feature code does not use `toISOString().slice(0, 10)` or `toISOString().split('T')[0]` for business dates.
- Late-evening Serbia local time does not drift the displayed business date by one day.
- UTC timestamp diagnostics continue using real ISO timestamps.

---

## 5. BankaImport / BankaMapiranje Gates

### 5.1 PDF Extract Success

Acceptance:

- `PDFTOTEXT_EXE_PATH` exists in `tblLocalConfig` or resolves through the standard setup fallback.
- Valid PDF statement extracts text through local `pdftotext`.
- Extraction uses a unique temp txt path, not `%TEMP%\pdf_extract.txt`.
- Temp txt path is cleaned up after extraction.
- Parsed text contains required statement header and transaction data.

### 5.2 PDF Extract Failure

Acceptance:

- Missing/wrong `PDFTOTEXT_EXE_PATH` fails hard.
- Forced non-zero `pdftotext` exit code fails hard.
- Missing output temp txt after zero exit fails hard.
- Failed extraction returns `extract error` / error outcome and must not enter processed success flow.
- Failed/problematic PDF does not create partial staged transaction rows.

### 5.3 Statement Integrity Failure

Acceptance:

- Missing `BrojIzvoda`, `DatumIzvoda`, `BrojRacuna` or saldo block fails before staging.
- `PocetnoStanje + sum(Uplata) - sum(Isplata) <> ZavrsnoStanje ±0.01` fails before staging.
- Parsed `Uplata` total mismatch against `UkupanPotrazuje` fails before staging.
- Parsed `Isplata` total mismatch against `UkupanDuguje` fails before staging.
- Parsed uplata/isplata count mismatch against `BrojNalogaOdobrenje` / `BrojNalogaZaduzenje` fails before staging.
- On failure, no transaction row from that statement is staged.

### 5.4 Append Failure Rollback

Acceptance:

- Missing `tblBankaImport` fails immediately.
- Missing required `tblBankaImport` column fails immediately.
- Invalid input data array fails immediately.
- Empty `GetNextID` result fails immediately.
- Forced `AppendRow <= 0` rolls back the BankaImport transaction.
- `IsDuplicateBankaImport` uses required column guards.

### 5.5 Deferred File Move

Acceptance:

- `ImportOnePdfIntoBankaImport` stages data and records pending successful moves; it does not move successful files immediately.
- `ImportBankaInbox_TX` commits before moving successful PDFs to `Processed`.
- Batch failure before commit moves no successful PDF to `Processed`.
- Failed/problematic PDF remains in Inbox or error path for operator review according to the import outcome.
- Post-commit file-move failure is reported clearly for manual folder recovery.

### 5.6 Duplicate-Key Mapping Failure

Acceptance:

- Duplicate `BankaImportID`, `NovacID`, `OtkupID`, or `FakturaID` fails mapping.
- Missing critical link row fails mapping.
- `UpdateBankaImportStatus` uses exact-row guard and checked write.
- `LinkNovacToOtkupStrict` checks exact `NovacID`, exact `OtkupID` and checked update.
- Manual and auto-map paths share the same exact-row integrity standard.

### 5.7 Stornirano and Skip/Processed Guards

Acceptance:

- `ValidateBankaImportNotProcessed` rejects `Obradjeno = Da`.
- `ValidateBankaImportNotProcessed` rejects `Obradjeno = Skip`.
- `ValidateBankaImportNotProcessed` rejects `Stornirano = DA`.
- `GetBankaImportOpen()` excludes stornirano, processed and skipped rows.

### 5.8 `GetBankaImportRowByID` Contract Smoke

Acceptance:

- Function returns the legacy 1x10 semantic shape, not raw table row.
- `bim(1, 1)` through `bim(1, 10)` map to:

```text
BrojDokumenta
DatumTransakcije
Partner
PartnerKonto
Uplata
Isplata
Opis
SvrhaPlacanja
BankaReferenz
PozivNaBroj
```

### 5.9 Kooperant Block Allocation Smoke

Acceptance:

- Empty `novID` after `SaveNovac` is a hard error.
- `LinkNovacToOtkupStrict` is called exactly once per consumed otkup candidate.
- Candidate count increments exactly once.
- `preostaloZaRaspodelu` is reduced exactly once per consumed candidate.
- Excess amount routes to kooperant avans according to current business rules.

---

## 6. SEF Gates

### 6.1 HTTPS Config Gate
### 6.2 Submit Gate
### 6.3 Status Refresh Gate
### 6.4 Recovery Gate
### 6.5 Parser Smoke

---

## 7. Monitoring Gates

Monitoring gates verify the observability pipeline without making monitoring a business transaction participant. All monitoring sends remain best-effort in production.

### 7.1 VBA Monitoring Configuration Gate

Acceptance:

- `tblSEFConfig` contains `MONITORING_ENDPOINT`, `MONITORING_SECRET` and `MONITORING_ENV`.
- `APP_VERSION` is resolved from `modConfig.APP_VERSION`.
- `deviceId` is generated from local workstation/user identity.
- Endpoint/secret/env are not hardcoded into `modMonitoring`.

### 7.2 GAS Monitoring Configuration Gate

Acceptance:

- Script Properties contain `MONITORING_SPREADSHEET_ID`, `MONITORING_ALERT_EMAIL` and `MONITORING_INGEST_SECRET`.
- GAS constants store property names, not the actual values.
- `monitorPublic` validates `monitoringSecret`.
- Authenticated `monitor` remains under the existing token model.

### 7.3 VBA Monitoring HTTP Gate

Acceptance:

- `TestMonitoring_HTTP` or equivalent diagnostic returns `HTTP Status = 200`.
- Response has `success = true`.
- Response includes `eventId`, timestamp, severity and component.
- Debug output prints only redacted diagnostics, not the full JSON body.

### 7.4 Workbook Tabs Gate

Acceptance:

`OtkupApp_Monitoring_PROD` exists with these tabs and canonical headers:

```text
Health
Events
Errors
SyncStatus
SEFStatus
UserSessions
Backups
Alerts
AuditCritical
```

### 7.5 Event Routing Gate

Acceptance:

- Every valid event appends to `Events`.
- `ERROR` and `CRITICAL` events route to `Errors`.
- Sync/heartbeat events route to `SyncStatus` and/or `UserSessions`.
- `SEF_*` events route to `SEFStatus`.
- `BACKUP_*` events route to `Backups`.
- Alert-worthy events route to `Alerts`.
- Audit-impacting events route to `AuditCritical`.
- Every valid event updates relevant `Health` component.

### 7.6 Health Component Gate

Acceptance:

`Health` contains maintained rows for:

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

Watchdog checks keep state-driven components current even when user events are not emitted frequently.

### 7.7 SEF Monitoring Gate

Acceptance:

- SEF startup recovery emits `SEF_STARTUP_RECOVERY_*` and `SEF_RECOVERY_INVOICE_*` events.
- SEF submit emits `SEF_SEND_*` events.
- SEF status refresh emits `SEF_STATUS_*` and `SEF_REFRESH_PENDING_*` events.
- `SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING`, `SEF_RECOVERY_INVOICE_FAIL`, `SEF_STATUS_REFRESH_EXCEPTION` and unknown/manual-review states create alert/audit-critical candidates.
- SEF monitoring does not replace `tblSEFSubmission`, `tblSEFEventLog`, `AppendSEFEvent_Row`, `SaveSEFSubmissionResult_Row` or the SEF state machine.

### 7.8 Business Transaction Monitoring Gate

Acceptance:

- `CreateFaktura_TX` emits faktura success/fail monitoring.
- `SaveNovac_TX` emits novac success/fail monitoring.
- `ApplyAvansToFaktura_TX` emits failure monitoring.
- `SaveOtkup_TX` / `SaveOtkupMulti_TX` emit otkup success/fail monitoring.
- Document-chain TX wrappers emit fail-only `DOKUMENT_SAVE_FAIL`.
- Bank mapping TX wrappers emit bank mapping/skip/batch events.
- MasterSync and Stammdaten export emit success/fail monitoring.
- Helper/read/update-level functions do not emit row-level noise.

### 7.9 Bank Mapping Monitoring Gate

Acceptance:

- Bank mapping events are emitted at TX-wrapper level only.
- Base mapping helpers are not double-instrumented.
- Batch automap emits start, summary and failure events.
- Payloads avoid sensitive bank account details beyond operational identifiers needed for diagnosis.

### 7.10 Backup Monitoring Gate

Acceptance:

- Backup success emits `BACKUP_SUCCESS`.
- Backup failure emits `BACKUP_FAIL`.
- `Backups` captures backup type, source, backup file, location, rows, checksum, status, duration and error.
- Backup monitoring failure does not change the backup procedure result.

### 7.11 Watchdog Gate

Acceptance:

- `runMonitoringWatchdog` exists and can run.
- The monitoring workbook is initialized/accessible.
- Required tabs are accessible.
- Auth/secret configuration is checked.
- Backup freshness is checked.
- MasterData/Stammdaten freshness and failure state are checked.
- SEF stuck/unknown/manual-review state age is checked.
- PWA sync/offline queue failures and stale pending rows are checked.
- Recent GAS error spikes are checked.
- Alerts are deduplicated while unresolved by component, alert code and correlation ID.

### 7.12 Scheduled Jobs Gate

Acceptance:

- `runMonitoringWatchdog` trigger is installed every 15 minutes.
- `sendDailyMonitoringSummary` trigger is installed once per day.

### 7.13 Alert / AuditCritical Gate

Acceptance:

- `CRITICAL` and alert-worthy `ERROR` events create alerts.
- `needsManualReview = true` creates operator-visible manual-review state.
- `nextAction` uses one of `WAIT`, `RETRY`, `MANUAL_REVIEW`, `CHECK_SEF_PORTAL`.
- Audit-critical events include SEF critical states, payment/faktura critical events, manual overrides and critical bank/backup failures.

### 7.14 Monitoring Test Suite Gate

Acceptance:

Run or confirm the active VBA monitoring test surface:

```text
TestMonitoring_All
TestMonitoring_Config
TestMonitoring_HTTP
TestMonitoring_ErrorEvent
TestMonitoring_SEFUnknown
TestMonitoring_BackupSuccess
TestMonitoring_BackupFail
```

### 7.15 Deployment Setup Gate

Acceptance:

- `OtkupApp_Monitoring_PROD` exists with canonical tabs and headers.
- GAS Script Properties are configured.
- `tblSEFConfig` monitoring keys are configured.
- GAS Web App is redeployed after monitoring code changes.
- Watchdog and daily-summary triggers are installed.
- At least one end-to-end business-flow smoke confirms monitoring does not block business success.

---

## 8. Agrohemija / Digitalni Agronom Gates

### 8.1 Desktop `SaveMagacin` Validation Gate

Acceptance:

- Missing article fails before append.
- Invalid movement type fails before append.
- Non-positive quantity fails before append.
- `MAG_IZLAZ` without required kooperant fails before append.
- `MAG_IZLAZ` with insufficient current stock fails before append.
- Required column reads use fail-fast schema guards.

### 8.2 Desktop `SaveMagacin_TX` Transaction and Monitoring Gate

Acceptance:

- `SaveMagacin_TX` snapshots `tblMagacin` before business write.
- Forced save failure rolls back the snapshot.
- Successful commit emits `MAGACIN_SAVE_SUCCESS` best-effort.
- Failure emits `MAGACIN_SAVE_FAIL` and `Monitor_Error` where monitoring is configured.
- A monitoring exception after commit does not turn the committed business save into a false failure.
- The original `SaveMagacin` failure reason is preserved in logs/operator diagnostics.

### 8.3 Desktop `frmAgrohemija` Basket Gate

Acceptance:

- Issue and receipt baskets commit through one explicit `clsTransaction` over `tblMagacin`.
- A forced failure on one basket line rolls back all earlier basket writes.
- Issue basket stock check aggregates quantities by `ArtikalID` before commit.
- Multiple basket rows for the same article cannot individually pass and collectively exceed stock.
- Multiple parcel IDs serialize with semicolon (`;`).
- Return navigation routes back to `frmOtkupAPP` without embedding business writes.

### 8.4 PWA Management Agrohemija Issuing Gate

Acceptance:

- Recommendation quantity is real unit-of-measure quantity.
- Packaging example: raw 2.4 kg with 1 kg package persists/displays 3 kg, not 3 packages.
- Package count is displayed only as explanatory metadata where shown.
- Selected parcel IDs serialize with semicolon (`;`).
- `izdZavrsi()` opens the printable/signable otpremnica modal and does not write directly.
- Final save is protected by `withSubmitLock`.
- Double-click final save produces one user submit in client smoke.
- `izdReset()` clears cart, selected kooperant, selected article/quantity, parcel list, note, recommendation state, modal data and barcode debounce state.
- Render functions tolerate missing DOM targets during partial screen lifecycle transitions.

### 8.5 PWA Kooperant Digitalni Agronom Gate

Acceptance:

- `agroCalcPreporuka()` writes real quantity semantics in every branch.
- Packaged article recommendation uses `ceil(rawQty / pakovanje) * pakovanje`.
- Lager warning compares real final quantity against `stammdaten.magacinkoop` stock.
- `agroPrimeniPreporuku()` writes real quantity into `agroKolicina` and `agroState.kolicina`.
- `agroValidateKolicinaNaLageru()` blocks quantity greater than local stock before local treatment save.
- `agroSaveTretman()` is protected by `withSubmitLock('agro:tretman:save', ...)`.
- Offline treatment remains pending and syncs after reconnect.
- Treatment history reload shows one deduplicated treatment.
- `agroResetState()` clears both `active` and `selected` measure-button classes.

### 8.6 GAS Treatment Sync Gate

Acceptance:

- `syncTretman` accepts authenticated POST for allowed `Kooperant` and `Management` callers.
- Kooperant caller is scoped to its own `kooperantID`.
- Payload requires `records[]`.
- Handler executes under `withLock(...)`.
- Each row is processed through `processTretmanRecord(record, kooperantID)`.
- Response uses `buildBatchSyncResponse(results)` and includes `results[]`.
- Storage target is `TRETMAN-<KooperantID>`.
- Idempotency by `ClientRecordID` returns existing/terminal success instead of duplicate append.
- `getTretmaniForKooperant(kooperantID)` reads the same per-kooperant sheet.

### 8.7 GAS Management `saveIzdavanje` Boundary Gate

Acceptance:

- `saveIzdavanje` remains Management-only.
- Existing `Izdavanje` sheet contract is unchanged.
- Client submit-lock is present because current server retry idempotency by stable client issuance ID is not guaranteed.
- Server-side `saveIzdavanje` idempotency remains roadmap unless implemented in a future release.

### 8.8 Quantity Semantics Cross-Layer Gate

Acceptance:

- Desktop Agrohemija, PWA management issuing and PWA kooperant agromere all persist real unit-of-measure quantities.
- Package count is never persisted as `Kolicina`, `KolicinaUpotrebljena`, `DozaPreporucena`, `DozaPrimenjena` or issue quantity unless the article unit itself is a package.
- Treatment save does not decrement server-side stock unless a future backend stock-consumption contract is explicitly implemented.

---


## 9. Security and Compliance Gates

Security gates verify that production deployment has the expected auth, transport, redaction and local-state boundaries. These gates are required whenever GAS routing, PWA shell/session code, SEF config, monitoring, geo endpoints or workstation setup changes.

### 9.1 SEF HTTPS Config Gate

Acceptance:

- Temporarily set `SEF_BASE_URL` to an `http://` URL in a non-production/test config.
- Run SEF config validation / SEF smoke path.
- Confirm the request fails locally before any network call.
- Restore valid `https://` config.

### 9.2 GAS Auth / Authorization Matrix Gate

Acceptance:

- `ping` / allowed public health action behaves as expected.
- Missing token on protected actions returns 401/auth error.
- Valid token with wrong role returns 403/forbidden.
- Valid token with mismatched entity ID returns forbidden/error for scoped writes.
- Management-only writes reject Kooperant/Otkupac/Vozac roles.
- Write actions are below the normal token-validation gate unless explicitly documented as public exceptions.

### 9.3 Sync Entity Ownership Gate

Acceptance:

- Otkupac cannot sync another `otkupacID`.
- Kooperant cannot sync another `kooperantID` for treatments/troškovi.
- Vozac cannot sync or update another `vozacID`; GAS uses `tokenData.entityID` for Vozac-owned status updates.
- Management override works only on endpoints documented to support Management override.
- `syncTrosak` returns a normal batch result and never an empty HTTP 200 body.

### 9.4 GAS Write Lock Gate

Acceptance:

- Sync writes, dispatch/demand writes, PDF/Drive writes, fiskalni mapping writes, kamion status writes, parcel polygon writes if enabled, and master-artikal writes use `withLock(...)`.
- Concurrent write smoke does not produce duplicate/partial rows for the same idempotency key.
- Disabled endpoints return explicit disabled responses rather than silently doing nothing.

### 9.5 Shared Fiscal / Master-Data Authorization Gate

Acceptance:

- `saveFiskalniMapiranje` is Management-only.
- `createArtikal` is Management-only.
- Fiscal parse actions require Kooperant or Management and are not public utilities.
- Kooperant-scoped fiscal payloads match authenticated entity when a kooperant scope is supplied.

### 9.6 Monitoring Secret and Redaction Gate

Acceptance:

- Wrong `MONITORING_SECRET` / `MONITORING_INGEST_SECRET` is rejected.
- Valid secret produces HTTP 200 and `success: true` response in test mode.
- Debug output does not print full JSON body.
- Monitoring payloads do not expose tokens, secrets, SEF API key, full UBL XML, full PDF/base64/file content or authorization headers.
- Client error reporting truncates stack/details and redacts token/PIN/password/base64-like content.

### 9.7 PWA CSP / Asset Gate

Acceptance:

- CSP remains compatible with `script-src 'self'` target.
- Runtime/vendor assets needed for offline app-shell behavior are self-hosted.
- Service-worker cache includes changed critical assets.
- `CACHE_NAME` is bumped after critical runtime asset changes.
- `index.html` does not double-load critical JS services.
- Leaflet marker assets are cached when map/offline marker surfaces are launch-relevant.

### 9.8 Local State Audit Gate

Acceptance:

- Shared operational state is not stored in `localStorage`.
- `localStorage` usage is limited to device-local preference/helper state.
- IndexedDB stores own offline queues/caches.
- Pending/error records remain visible/retryable and are not hidden by clean server copies.

### 9.9 Local Workstation Security Gate

Acceptance:

- `PDFTOTEXT_EXE_PATH` is read from `tblLocalConfig` or validated setup fallback.
- `tblConfig` is not used as workstation-local config storage.
- `tblSEFConfig` owns SEF/monitoring runtime config.
- Missing `pdftotext.exe` produces setup/health warning or hard parser error as appropriate.
- PDF extraction uses unique temp file paths and cleans them before/after extraction.

### 9.10 Public Geo / Meteo / Parcel Polygon Review Gate

Acceptance:

- Confirm deployed authorization state for `saveParcelPolygon`.
- If deployed as Management-only: verify missing token, wrong role and non-management calls fail; valid Management call succeeds under `withLock(...)`.
- If any public/pre-auth polygon write remains deployed: record explicit accepted risk, deployment URL protection and remediation plan.
- Confirm public geo/meteo actions are read-only if retained.
- Confirm no new public write endpoints were introduced.


## 10. GIS / Parcele / Meteo Gates

### 10.1 VBA Compile / Dependency Gate

Acceptance:

- Compile all VBA modules.
- Confirm `SyncParceleToGoogle_Core` exists and is public where `modGeoParcele` can call it.
- Confirm `COL_PAR_*` constants compile.
- Confirm `clsTransaction` is available.
- Confirm `RequireUpdateCell` is available.
- Confirm existing Google helper functions are available: `ReadSheetData`, `WriteSheetData`, `CreateSpreadsheet`, `GetSpreadsheetID`, `AddSheetTab`.
- Confirm no reserved-name compile issues in geo modules/forms.

### 10.2 Form Designer Gate

Acceptance:

- `lblGeoStatus` exists.
- `btnGeoOpen`, `btnPasteCoords`, `btnGeoClear`, `btnGeoSave`, `btnOpenMap`, `btnOpenPolygonEditor` exist.
- `txtNCoord` and `txtECoord` exist.
- Labels exist or `SetGeoControlsVisible` safely tolerates missing optional labels.
- Geo controls are hidden outside `Parcele` mode.

### 10.3 Geo UI Smoke

Acceptance:

1. Open `frmStammdaten` in non-`Parcele` mode.
2. Verify geo controls are hidden.
3. Open `Parcele` mode.
4. Verify geo controls are hidden before row selection.
5. Select parcel.
6. Verify geo controls become visible.
7. Change selection.
8. Verify clear-confirm state resets.
9. Verify normal geo flow uses inline `lblGeoStatus`, not blocking `MsgBox` prompts.

### 10.4 Geo Save Smoke

Acceptance:

1. Paste coordinates.
2. Save point.
3. Verify `N_Coord`, `E_Coord`, `Lat`, `Lng`, `GeoStatus`, `GeoSource`, `MeteoEnabled`, `DatumGeoUnosa` and `DatumAzuriranja` update.
4. Verify selected parcel remains selected after `LoadList`.
5. Force a controlled write error and verify transaction rollback.
6. Verify no partial geo state remains after rollback.

### 10.5 Geo Clear Smoke

Acceptance:

1. Click clear once.
2. Verify inline status/caption says second click is required.
3. Click clear again.
4. Verify point fields clear.
5. Verify `GeoStatus`, `GeoSource`, `MeteoEnabled` and update date fields reflect cleared state.
6. Verify `COL_PAR_POLYGON` is not cleared by point clear.
7. Verify selected parcel remains selected after `LoadList`.
8. Force a controlled write error and verify rollback.

### 10.6 Selected Parcel Sync Gate

Acceptance:

- Select a valid parcel.
- Run `SyncSelectedParcelaToGoogle(parcelaID)`.
- Verify it calls/reuses `SyncParceleToGoogle_Core(False)` rather than duplicating `Parcele` export mapping.
- Verify selected `ParcelaID` exists in Google `Stammdaten / Parcele` after sync.
- Verify missing local row fails closed.
- Verify missing Google sheet ID, unreadable `Parcele` tab or missing `ParcelaID` header fails closed.

### 10.7 Polygon Editor Gate

Acceptance:

1. Select parcel.
2. Save point locally.
3. Click open polygon editor.
4. Verify `SyncSelectedParcelaToGoogle` succeeds.
5. Verify editor opens for the same encoded `ParcelaID`.
6. Verify editor sees current point/row data.
7. Force selected sync failure and verify editor does not open.

### 10.8 Polygon Overwrite Safety Gate

Acceptance:

1. Create or edit polygon in editor/PWA/HTML.
2. Confirm polygon exists in Google.
3. Run full-cycle PWA/Google sync.
4. Confirm `ImportParcelGeoFromGoogleToMaster` pulls polygon into master before outbound Stammdaten export.
5. Confirm later Stammdaten export does not wipe the Google polygon.

### 10.9 PWA Parcel Map Gate

Acceptance:

- Kooperant parcel screen initializes one Leaflet map instance per shell lifetime.
- `getParcelGeo(parcelaId)` is used for parcel geometry reads.
- `PolygonGeoJSON` renders polygon parcels.
- Valid `Lat` / `Lng` renders point-only markers when polygon is absent.
- Parcel popup shows key parcel fields and explicit navigation to parcel detail.
- GIS interactions use delegated `data-action` hooks, not inline handlers.
- `focusParcel(...)` synchronizes list click, map focus, popup open, highlight and detail navigation.
- `openParcelaDetail(...)` renders `osnovno`, `meteo`, `radovi`, `troskovi` and `proizvodnja` tabs.

### 10.10 Meteo Scheduled Fetch Gate

Acceptance:

- `setupMeteoTriggers()` creates daily triggers for 00:00, 06:00, 12:00 and 18:00 in `Europe/Belgrade`.
- `scheduledMeteoFetch()` groups parcels by rounded 0.01 lat/lng buckets.
- Batch Open-Meteo fetch works for grouped coordinates.
- Individual retry/fallback works when batch fetch fails.
- `MeteoHistory` receives append-only history rows.
- `MeteoLatest` receives current overwrite-state rows.
- Risk/spray/forecast JSON fields are serialized for frontend reads.

### 10.11 Meteo Cached-First Read Gate

Acceptance:

- `getParcelMeteo(parcelaId)` uses `MeteoLatest` when `LastFetch` is younger than 12 hours.
- Stale/missing cache falls back to live forecast retrieval.
- PWA prewarms `window.meteoCache` from exported `stammdaten.meteoLatest` before per-parcel API fallback.
- Card-level meteo respects a 6-hour `METEO_CACHE_TTL`.
- Parcel detail meteo tab and home-dashboard alerts consume the same current-state data.

### 10.12 Meteo Risk / Spray Window Gate

Acceptance:

- Culture thresholds exist for `Visnja`, `Jabuka`, `Sljiva`, `Kruska`, `Breskva`, `Malina` and `_default`.
- Threshold payload includes `frostWarn`, `frostDanger`, `heatWarn`, `heatDanger`, `sprayWindMax`, `sprayRainHours`, `optimalTempMin`, `optimalTempMax`.
- `assessRisk(...)` derives frost, heat, rain and disease risk plus aggregate level.
- `calculateSprayWindow(...)` evaluates a 72-hour horizon and keeps valid contiguous spray-safe windows.
- Digitalni Agronom treatment validation and spray timing use the same meteo/risk outputs.

### 10.13 KPI Robustness Gate

Acceptance:

1. Put a blank/text date in a relevant test row.
2. Put a non-numeric value in a kg field in a test row.
3. Open `frmOtkupAPP`.
4. Refresh sidebar KPI.
5. Verify no `13 | Type mismatch` log from `SumOtkupKgForDate` or `CountDocsForDate`.
6. Verify KPI shows safe zero, valid aggregate or safe placeholder.
7. Verify UI remains usable.

---


## 11. Reports and Derived Views Gates

### 11.1 Management Reports Gate

Acceptance:

- `getMgmtAll` returns a complete management bundle for dashboard, pregled, dispatch, partner, saldo and agro views.
- `MgmtReports` and `SaldoOMDetail` refresh from desktop/exported source facts and are not used as write-back sources.
- Management dashboard renders with fresh export data and with stale/cache state clearly represented when applicable.
- Dispatcher board remains planning-only and does not directly assign `VozacID` to OTK rows.

### 11.2 Financial Reports Gate

Acceptance:

- `SaldoOM`, `SaldoOMDetail`, `SaldoKupci`, `Kartica Kooperanta` and open BankaImport queue render from canonical source tables.
- Stornirano rows are excluded from active finance views unless the report is explicitly audit/history.
- Kartica parsing ignores `UKUPNO` rows in production parsing.
- Required source columns fail fast rather than producing misleading empty reports.
- BankaImport queue reflects current workflow state without deleting or overwriting staged bank facts.

### 11.3 Agrohemija / Warehouse Reports Gate

Acceptance:

- `GetMagacinStanje()` returns current `(Ulaz, Izlaz, Stanje)` after excluding stornirano rows.
- `ReportIzdavanjePoKooperantu()` works with `datumOd` only, `datumDo` only, both dates or neither.
- `ReportStanjePoDobavljacu()` is available as the correct public spelling.
- `ReportStanjePoDoabvljacu()` remains available only as compatibility wrapper if callers still exist.
- `GetAgrohemijaDug()` and `GetAgroAbzug()` remain separated by ownership: warehouse debt helper vs finance deduction helper.

### 11.4 PWA Role View Gate

Acceptance:

- Otkupac today overview renders merged local/server records through the canonical dedupe helper.
- Kooperant treatment/expense/knjiga polja views keep pending/syncing/error local records visible.
- Vozač zbirna overview displays `BrojZbirne` as the business number and does not confuse it with `ServerRecordID`.
- Management views render dashboard/dispatch/partner/agro sections without hidden dependency on deprecated globals.

### 11.5 Dashboard KPI Robustness Gate

Acceptance:

1. Put a blank, text or Excel-error value in a date field used by a KPI helper.
2. Put a non-numeric value in a kg/amount field used by a KPI helper.
3. Open/refresh the relevant dashboard or operator shell.
4. Verify no runtime `13 | Type mismatch` is logged from KPI helpers such as `SumOtkupKgForDate`, `CountDocsForDate` or `RefreshSidebarKpi`.
5. Verify the UI remains usable and displays a safe zero, valid aggregate or placeholder.

### 11.6 Monitoring Read-Model Gate

Acceptance:

- `Health` reflects current component status from events and watchdog checks.
- `Events` remains append-only for valid monitoring events.
- `Errors` receives `ERROR` / `CRITICAL` rows without leaking sensitive payloads.
- `SEFStatus`, `Backups`, `Alerts`, `AuditCritical` and `SyncStatus` are updated by their owning monitoring routes.
- Monitoring read models do not replace business transaction tables or SEF persistence tables.

---


## 12. Data Architecture Gates

### 12.1 Canonical Entity Inventory Gate

Acceptance:

- Every production workbook contains the canonical master, transaction, document, finance, BankaImport, Agrohemija and config tables required by the current AR.
- `tblLocalConfig`, `tblConfig` and `tblSEFConfig` are present and used for their distinct configuration scopes.
- `tblBankaImport` contains `PocetnoStanje`, `ZavrsnoStanje`, `UkupanDuguje` and `UkupanPotrazuje` before new bank imports are accepted.
- Business tables that rely on soft delete/status have the required `Stornirano`, `Aktivan`, `Aktivna` or equivalent visibility fields.

### 12.2 Schema Guard Gate

Acceptance:

- Required-column reads in production save/import/map/report paths use `RequireColumnIndex` or `RequireColumns`.
- Critical writes use `RequireUpdateCell` or an approved exact-row update helper.
- Missing required tables/columns fail loudly; they do not produce empty reports, partial imports or silent no-op writes.
- Optional columns are explicitly documented as optional and have safe missing-column behavior.
- Append flows treat empty `GetNextID` and `AppendRow <= 0` as hard failures.

### 12.3 Source-of-Truth / Derived-View Gate

Acceptance:

- `MgmtReports`, `SaldoOMDetail`, `Kartice`, `MeteoLatest`, monitoring tabs and PWA caches are verified as read models unless explicitly documented otherwise.
- Report/read-model code does not silently correct source tables during rendering.
- Desktop exports refresh Google/PWA read models from canonical source tables.
- PWA local pending/error records remain visible and are not hidden by stale server reads before reconciliation.

### 12.4 Google Sheets Transport Gate

Acceptance:

- `Stammdaten` tabs reflect active desktop master data and do not become independent editable source tables.
- OTK/VOZ/TRETMAN/TROSKOVI/FISKALNI role sheets preserve `ClientRecordID`, `ServerRecordID`, `SyncStatus` and relevant sync timestamps.
- VOZ column B / `ServerRecordID` and column T / `BrojZbirne` remain separate.
- `SyncControl` / `MASTER_SYNC_LOCK` blocks unsafe writes during full-cycle sync where required.
- Google writeback side effects are treated as external and not falsely covered by local Excel rollback.

### 12.5 Data Cleanliness / Migration Gate

Acceptance:

- `RunProductionHealthCheck` passes for the target workbook before production launch.
- Legacy/demo/test rows are removed or marked safely inactive/stornirano where they would pollute health checks.
- Historical `tblBankaImport` rows with blank saldo fields are understood as historical/non-backfilled rows, not automatic corruption.
- Any future table/column migration has an explicit migration note, backfill plan or accepted no-migration statement.

---


## 13. Google Sheets Data Layer Gates

### 13.1 Stammdaten Workbook Gate

Acceptance:

- `SyncStammdatenToGoogle()` finds or creates the `Stammdaten` spreadsheet.
- `GOOGLE_STAMMDATEN_SHEET_ID` is persisted and reused.
- Required tabs exist: `Kooperanti`, `Kulture`, `Parcele`, `Config`, `Users`, `Fakture`, `FakturaStavke`, `SaldoOMDetail`, `Stanice`, `Kupci`, `Vozaci`, `Artikli`, `MagacinKoop`.
- Stammdaten remains an exported projection, not an uncontrolled manual source table.

### 13.2 OTK Sheet Contract Gate

Acceptance:

- OTK sheets use the canonical GAS-first `COLUMNS` order.
- `ClientRecordID`, `ServerRecordID`, `SyncStatus`, timestamps and business fields are preserved.
- Desktop import treats `SyncStatus = "Synced"` as pending-for-master.
- Writeback targets remain `Sheet1!F` for `SyncStatus` and `Sheet1!B` for `ServerRecordID`.
- Invalid/skipped rows receive controlled statuses such as `Duplicate` or `SyncError[:reason]`.

### 13.3 VOZ Sheet Contract Gate

Acceptance:

- VOZ sheets use the canonical GAS-first `ZBIRNA_COLUMNS` order.
- Column B / `ServerRecordID` and column T / `BrojZbirne` remain separate.
- `ServerRecordID` is never reused as business `BrojZbirne`.
- `WriteBackVOZSyncStatus` writes B/F/T correctly.
- `TipAmbalaze` and `BrojZbirne` are plain-text formatted to prevent date coercion.

### 13.4 Per-Kooperant Sheet Gate

Acceptance:

- `TRETMAN-<KooperantID>`, `TROSKOVI-<KooperantID>` and `FISKALNI-<KooperantID>` are entity-scoped.
- Kooperant access cannot read/write another kooperant's private sheets.
- `ClientRecordID` idempotency is preserved in active sync processors.
- Empty HTTP 200 responses are rejected by smoke tests for sync endpoints.

### 13.5 Exported Read-Model Gate

Acceptance:

- `Kartice` export uses `KooperantID | Datum | BrojDok | BrojParcele | Opis | Zaduzenje | Razduzenje | Saldo`.
- `MgmtReports` contains `SaldoOM`, `SaldoKupci`, `OtkupPoOM` and `PredatoPoKupcu`.
- Exported read models are not used as source tables for business corrections.
- Freshness/last-export diagnostics are available where operationally required.

### 13.6 Sheet Registry / Header Drift Gate

Acceptance:

- `rebuildSheetRegistry()` is run/verified after backend deployment when registry behavior changes.
- Required header mismatches produce schema-drift errors, not shifted-column writes.
- OTK/VOZ import/export constants match live headers.
- Schema drift is visible through controlled error/log/writeback paths.

### 13.7 SyncControl / Master Lock Gate

Acceptance:

- `Stammdaten / SyncControl` contains `MASTER_SYNC_LOCK`, `MASTER_SYNC_UPDATED_AT`, `MASTER_SYNC_MESSAGE` and `MASTER_SYNC_OWNER`.
- VBA sets and releases `MASTER_SYNC_LOCK` during full-cycle sync.
- GAS blocks unsafe writes while the lock is active.
- PWA treats active lock as temporary retry/soft-lock, not a permanent error.
- Stale-lock behavior is smoke-tested.

### 13.8 Parcel Geo Pull Gate

Acceptance:

- `ImportParcelGeoFromGoogleToMaster` runs before outbound Stammdaten export.
- If geo pull fails, outbound Stammdaten export aborts.
- Newer Google-side polygon/geo data is not overwritten by stale local `tblParcele` values.

### 13.9 Google External Side-Effect Gate

Acceptance:

- Documentation and code comments do not claim Google writebacks are covered by local Excel rollback.
- Writeback failures are operator-visible and recoverable.
- Critical writeback paths are ordered so irreversible Google updates do not create misleading success states.

---

## 17. Pass 17 GO Hardening Closeout Gates

### 17.1 Google Sheets Staging / Verify / Replace Gate

Acceptance:

- `WriteSheetData` does not clear the target tab before a verified replacement is ready.
- A staging tab is created, values are written to staging, staging is verified, target is replaced and final target is verified.
- Target replacement uses phased rename: `target -> backup`, `staging -> target`, then backup deletion.
- If backup deletion fails after replacement, the new target remains live and the leftover backup is reported for manual cleanup.

### 17.2 Google Sheets Quota / Cache Gate

Acceptance:

- Sheet ID lookup uses cache per spreadsheet and forced refresh only where required.
- `AddSheetTab` is no-op for an already-existing tab.
- Google HTTP retry covers `429`, `500`, `502`, `503` and `504`.
- Write-request throttling is active.
- Quota-window `429` uses the longer wait handling expected by the implementation.

### 17.3 Kartice Named-Tab Gate

Acceptance:

- `ExportKarticeToGoogle_Core` uses `KARTICE_TAB_NAME = "Kartice"`.
- Kartice export ensures named tab `Kartice` before write.
- Kartice export does not use `Sheet1` fallback behavior.
- Successful export writes through the staging/replace model.

### 17.4 Google/PWA Unlock Outcome Gate

Acceptance:

- Full sync cannot be reported green if final PWA unlock fails.
- Unlock failure produces partial/degraded operator result.
- Monitoring event is emitted for failed unlock.
- PWA/GAS stale-lock TTL recovery remains expected recovery path, but does not justify a green operator result.

### 17.5 MasterSync Exact-Row Link Gate

Acceptance:

- `AutoCreateOtpremniceFromPWA` validates grouped `OtkupID` values before creating/linking `Otpremnica`.
- `ImportVOZRow_RowTX` owns row-level import/link atomicity.
- `LinkZbirnaToOtkupAndOtpremnica`, `LinkOtkupToOtpremnicaStrict` and `LinkOtpremnicaToBrojZbirneStrict` reject missing and duplicate targets.
- Critical updates use `RequireUpdateCell`.
- `FindRows(...).Count > 0` is not used for critical document-chain link decisions.

### 17.6 SaveNovac Append Failure Gate

Acceptance:

- Empty `GetNextID` raises an error.
- `AppendRow <= 0` raises an error.
- Successful append returns a non-empty `NovacID`.
- BankaMapiranje, faktura payments, avans allocation/split and otkup payment flows fail rather than proceeding with empty `NovacID`.

### 17.7 ProductionHealthCheck Duplicate-Key Gate

Acceptance:

- Duplicate-key preflight checks run after core schema checks.
- Checks cover `OtkupID`, `OtpremnicaID`, `ZbirnaID`, `PrijemnicaID`, `FakturaID`, `NovacID`, `BankaImportID` and `ParcelaID`.
- Duplicate keys produce blocking health failures.
- Health check remains read-only and does not auto-repair duplicates.

### 17.8 GO Validation Evidence

Accepted validation evidence for this closeout:

```text
Geo=True
Otkup=True
Otpremnice=True
Zbirne=True
Stammdaten=True
Kartice=True
MgmtReports=True
PWA lock=NO
```


## 14. Final Validation Gate

Acceptance:

- `ARCHITECTURE_REFERENCE.md` contains current-state architecture only, not large release-delta bodies.
- `ARCHITECTURE_CHANGELOG.md` owns historical version deltas.
- `RELEASE_GATES.md` owns detailed smoke/regression/checklist steps.
- `ROADMAP.md` owns future hardening and accepted risks.
- `SECTION_MIGRATION_MAP.md` records where major old sections were moved.
- `archive/` contains historical full snapshots and legacy changelog material.
- AR section numbering is stable and domain-based.
- AR metadata is v6.22 and does not claim v6.20 as the active current version.
- No active rule is replaced with only “see previous version” or “see changelog”.
- All remaining uncertainties are explicit `NEEDS REVIEW` items.
- Final package includes `FINAL_VALIDATION_REPORT.md`.

---

## 15. Operator Pre-Launch Checklist

- Confirm local workstation setup.
- Confirm secrets/config values are present but not logged.
- Confirm monitoring is configured for the target environment.
- Confirm backups and recovery procedures are operational.


## 16. Pass 15 Omission-Fix Gates

### 16.1 Finance Architecture Gate

- Run or confirm `RunNovacSmokeSuite` when `modNovac`, avans allocation, partner mapping or BankaMapiranje changes are included.
- Confirm invalid both-direction money rows are rejected.
- Confirm valid `Uplata` and valid `Isplata` rows append correctly.
- Confirm stornirano novac rows are excluded from live aggregates.
- Confirm partner-map conflict raises a fail-fast error and identical mapping is idempotent.
- Confirm partial buyer avans split and partial otkup avans split preserve original/split row semantics.
- Confirm `ResetNovacOtkupLink_TX` clears links and recomputes otkup paid state in one transaction.
- Confirm required finance helpers use `RequireColumnIndex` and critical writes use `RequireUpdateCell`.

### 16.2 Ownership Matrix Gate

- Confirm every domain has exactly one canonical source of truth and any online/local projections are marked as projection/cache.
- Confirm `ServerRecordID` is never documented as a business number.
- Confirm `BrojZbirne` remains PWA-first with VBA fallback.
- Confirm LocalConfig, Config and SEFConfig boundaries are not mixed.

### 16.3 Endpoint Matrix Reconciliation Gate

- Compare deployed `Code.gs` actions against `ARCHITECTURE_REFERENCE.md` section 9.7.
- Confirm every write action has token/role/entity/lock behavior matching the matrix.
- Confirm disabled routes return `FEATURE_DISABLED`.
- Confirm `syncTrosak` is active and returns a batch response.
- Confirm `saveParcelPolygon` authorization state and remove `NEEDS REVIEW` only after code verification.

### 16.4 Deprecated / Transitional Gate

- Confirm no new code uses deprecated patterns listed in AR section 21.
- Confirm compatibility wrappers remain tested where still exposed.
- Confirm dev/test modules are not reachable from normal operator UI.

### 16.5 Glossary Gate

- Confirm business terms and identifiers in section 22 match terminology used in code, forms, sheets and operator training.


---

## 17. v6.22 Residual GO Hardening Gates

### 17.1 Faktura Duplicate-ID Guards

Required:

- create or fixture a duplicate `FakturaID` condition in a safe test workbook;
- `PrintFaktura` must fail before selecting or printing a first-found row;
- `UpdateFakturaStatus` must fail before recomputing a first-found row;
- failure must preserve original error context and not mutate unrelated faktura/novac state.

### 17.2 ParcelaID-Based Geo Save/Clear

Required:

- `SaveParcelGeoPointByID[_TX]` updates exactly one `tblParcele` row by `ParcelaID`;
- `ClearParcelGeoByID[_TX]` clears exactly one `tblParcele` row by `ParcelaID`;
- missing `ParcelaID` fails hard;
- duplicate `ParcelaID` fails hard;
- rollback removes partial point/lat/lng/status writes on forced error;
- polygon fields remain unchanged by point clear unless a separate polygon-clear contract is introduced.

### 17.3 Row-Index Compatibility Wrapper Gate

Required where row-index public wrappers are retained:

- wrapper resolves the target `ParcelaID` before mutation;
- wrapper delegates to the ByID API;
- wrapper is not the preferred architecture in new code;
- sorted/filtered/reloaded table scenarios do not update the wrong parcel.

### 17.4 Storno Eligibility Helper Gate

Required:

- `RequireStornoAllowed` / `CanStorno` / `LookupActiveID` pattern rejects missing rows;
- duplicate target IDs fail hard;
- already-stornirano rows fail before mutation;
- critical updates use `RequireUpdateCell`;
- business-layer storno logic does not depend on `MsgBox`.


---

## 18. v6.23 PWA Otkup Read-Model Gates

### 18.1 Master Projection Availability

Required:

- `MgmtReports/OtkupiAll` exists in the reports folder/read path;
- GAS/PWA reads `OtkupiAll` from the reports source, not from an incidental Drive fallback;
- `OtkupiAll` represents `tblOtkup` master data and is read-only for PWA display.

### 18.2 Operational Queue Availability

Required:

- `OTK-ST-*` / `OTK-*` sheets remain active operational inbox/live queue sources;
- PWA-origin rows in operational sheets are not hidden merely because `OtkupiAll` exists;
- role/station filters still apply where required.

### 18.3 Management Browser Smoke

Required / reported browser-tested:

- Management sees VBA/master-created otkupi;
- Management sees PWA-created/synced otkupi;
- Management Partneri / Kooperanti / Kupci load expected data;
- `getMgmtAll` returns populated report/kartica/read-model surfaces where data exists.

### 18.4 Otkupac Browser Smoke

Required / reported browser-tested:

- Otkupac sees VBA/master-created otkupi in its scope;
- Otkupac sees PWA/operational otkupi from `OTK-ST-*` / `OTK-*`;
- merged overview displays both worlds without dropping operational rows.

### 18.5 Otkup Merge Dedup Smoke

Required:

- a synced PWA otkup present in both `OTK-ST-*` and `OtkupiAll` renders once;
- dedup first checks `ServerRecordID` / `OtkupID`;
- `ClientRecordID` is fallback, not first priority, because it can differ between operational and master projection rows;
- natural-key fallback does not collapse legitimate distinct rows.


## 17. v6.24 PWA Design/Runtime and Numbering Gates

### 17.1 Target Git Repository Confirmation

- Confirm the target AgriX/OtkupApp frontend repository is available.
- Confirm the connected repo is not accidentally the unrelated `handoverApp` repository.
- Confirm origin/main or commit `ce90970` contains the listed work.

### 17.2 Design-System Gate

- `base.css` contains current brand tokens and compatibility aliases.
- `fonts.css` declares self-hosted Cormorant Garamond and DM Sans, including Latin-ext unicode ranges.
- `components_v2.css` contains shared app header/body/card/field/button/pill/list/record primitives.
- Existing role views reuse the shared primitives instead of duplicating local CSS.

### 17.3 Otkup Form Gate

- Otkup form renders as a 5-step flow.
- Class picker exposes only `I` and `II`.
- Package picker exposes canonical package options.
- Save bar shows live total.
- Driver selection is absent from Otkup form and present in Otprema.
- Picker handlers use event delegation, not inline scripts.
- Serbian comma decimal input is parsed correctly.

### 17.4 Otkupni List / Otprema / Pregled Gate

- Otkupni list modal renders with `.ol-*` classes and preserves required `data-action` hooks.
- Otprema summary/detail/success views render with current `.otp-*` patterns.
- Pregled/Danas pills, date range, stats and `.danas-*` cards render correctly.
- Problem badges map to existing semantic classes.

### 17.5 Runtime / Cache / Lazy Loading Gate

- `sw.js` cache version is bumped for redesigned app shell.
- Required font assets are included or reachable.
- `lazy.js` exists and is loaded before feature lazy-load callers.
- jsPDF, Leaflet, Chart.js and Firebase compat are lazy-loaded by feature paths.
- Heavy vendor files are not unnecessarily precached in the initial service-worker install.

### 17.6 VBA / Google Sync Diagnostic Gate

- Document-numbering helpers for `BrojDokumenta`, `BrojOtpremnice`, `BrojZbirne` are present in the real VBA code.
- Lock-based per-station numbering behavior is confirmed.
- `BuildOTKOperationalHeaders_` column count is aligned with operational row shape.
- `blockWriteIfMasterSyncActive` callers pass the required data/context.
- `modGoogleSheets` `sheetId = 0` sentinel issue is either patched or tracked as open.
