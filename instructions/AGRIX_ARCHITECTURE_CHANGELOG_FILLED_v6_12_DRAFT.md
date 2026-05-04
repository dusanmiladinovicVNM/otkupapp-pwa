# AgriX / OtkupApp Architecture Changelog

**Document Purpose:** Delta notes between canonical architecture snapshots  
**Companion to:** `AGRIX_ARCHITECTURE_REFERENCE_FILLED.md`  
**Owner:** Architecture documentation compiled from supplied reference set  

---

## 0. Changelog Contract

Ovaj dokument služi samo za **razlike između verzija**.

Dozvoljeno je da sadrži:
- added
- changed
- fixed
- deprecated
- removed
- known regressions
- roadmap changes

Nije dozvoljeno da bude zamena za canonical architecture reference.

Ako je neka stvar važeća arhitektura, ona mora biti upisana i u canonical reference dokument.

### 0.1 Writing Rule
Svaka stavka u changelog-u mora jasno da kaže:
- šta se promenilo
- zašto
- koji sloj je pogođen
- da li zahteva update reference dokumenta
- da li zahteva migraciju

### 0.2 Status Tags
Koristi sledeće tagove:
- **ADDED**
- **CHANGED**
- **FIXED**
- **DEPRECATED**
- **REMOVED**
- **KNOWN ISSUE**
- **ROADMAP**

---

## 1. Version Index

| Version | Date | Summary | Reference updated | Notes |
|---|---|---|---|---|
| v2.2 | 2026-03-08 | Sledljivost v2.0, novac renaming, Agrohemija, report refactor | Yes | post-session stabilization |
| v2.3 | 2026-03-13 | BankaImport + mapping matured, Kartica Kooperanta, orphan checks | Yes | finance/reporting expansion |
| v2.4 | 2026-03-17 | SEF v1.0 introduced | Yes | major e-faktura milestone |
| v2.5 | 2026-03-22 | Full NOVAC/SEF/BANKA/PWA snapshot and UI/module documentation | Yes | baseline full reference generation |
| v2.6 | 2026-03-28 | PWA OTK/VOZ flow, AutoCreateOtpremnice, stronger error/backup patterns | Yes | PWA transport integration |
| v2.7 | 2026-03-28 | Parcel GIS + Meteo pipeline | Yes | spatial/agro enrichment |
| v2.8 | 2026-03-29 | Dispečer implemented, localStorage ban enforced for shared state | Yes | planning architecture shift |
| v3.0 | 2026-04-10 | Modular PWA, Knjiga Polja, Fiskalni Scanner, Meteo batch | Yes | major platform expansion |
| v3.1 | 2026-04-19 | AgriX branding, formal system overview, role/state/sync/known issues consolidation | Yes | architecture narrative refocus |
| v5.6 | 2026-04-22 | MasterSync Zbirna import activated, OTK/VOZ GAS-first sheet alignment, Zbirna business numbering and VOZ sheet contract cleanup | Yes | sync/document architecture correction |
| v5.7 | 2026-04-23 | Management shell consolidation, dashboard stabilization, legacy mount/runtime cleanup | Yes | frontend management cleanup |
| v5.8 | 2026-04-23 | Session/auth canonicalization, role-wide sync, kooperant store cleanup, SW/CSP hardening, runtime-state normalization | Yes | frontend launch-hardening |
| v5.9 | 2026-04-23 | PWA self-hosted vendor assets, CSP `script-src 'self'`, normalized API client result shape, GAS POST-first bridge | Yes | frontend asset and contract hardening |
| v6.0 | 2026-04-24 | Frontend hardening, CSP cleanup, IndexedDB recovery layer, navigation shell cleanup | Yes | PWA pre-launch readiness |
| v6.1 | 2026-04-25 | OtkupApp v2.2.1 pre-launch desktop hardening: app lifecycle wired through StartApp, dead UI/theming code purged, frmBankaImport preview deduplicated, ReportSaldoOM avans edge bug closed, SEF mapper EH discipline restored | Yes | desktop pre-launch readiness |
| v6.2 | 2026-04-25 | GAS observability, persistent token fallback, ErrorLog client bridge, SheetRegistry lookup clarification | Yes | backend/runtime observability and session resilience |
| v6.3 | 2026-04-25 | GAS endpoint authorization matrix, role/entity ownership checks, write-lock coverage, parcel/fiskalni/kamion-status security hardening and launch smoke-test validation | Yes | endpoint-authz launch hardening |
| v6.4 | 2026-04-26 | VBA desktop pre-launch hardening: app lifecycle, Dokumenta atomic saves, Fakturisanje duplicate prevention/print selection, Otkup atomic multi-wrapper, Sledljivost checked linking, PWA-first/
VBA-fallback traceability rule, shared parse/combo/schema guards and business/UI separation | Yes | desktop launch hardening across app lifecycle, document-flow, invoicing, otkup and traceability |
| v6.12 | 2026-05-04 | PWA launch-smoke hardening: business-date cleanup, canonical sync-result convergence, client ErrorLog reporting, active syncTrosak endpoint, Kooperant expense sync, and VOZ/BrojZbirne post-VBA ownership clarification | Yes | no production data migration; PWA smoke passed for Otkupac/Kooperant/Vozac/Management; VBA VOZ writeback fix remains required for BrojZbirne column T |
| v6.11 | 2026-04-30 | Pre-launch persistence/security/data-health hardening: AR-002 central AutoSave after TX commit, canonical UTF-8 HTTP utilities, SEF HTTPS-only enforcement, SEF parser/http utility smoke coverage, and legacy test-data health cleanup | Yes | no schema migration; compile + HTTP/SEF/Google/MasterSync/Faktura/BusinessFlow/E2E smoke rerun required; ProductionHealthCheck must be clean before final launch |
| v6.10 | 2026-04-29 | Google VBA + GAS hardening: GoogleAuth/GoogleSheets ownership/fail-fast cleanup, MasterSync schema/idempotency/writeback guards, GAS auth/schema/sync processor hardening, inactive endpoint disablement and Google/GAS smoke suites | Yes | no data migration; compile + RunGoogleSyncSmokeSuite PASS 29/29 + runGasRouteHealthCheck PASS 29 handlers + runGasSmokeSuite PASS 10/10 |
| v6.9 | 2026-04-29 | Business-core hardening after v6.8: modFaktura canonical prijemnica values and status/print guards, modDokumenta EH/input/stornirano hardening, modOtkup validation/read-helper hardening, expanded Faktura and BusinessFlowPro regression suites | Yes | no schema migration; compile + RunFakturaSmokeSuite + RunBusinessFlowProSuite passed |
| v6.8 | 2026-04-29 | frmSEF operator-shell cleanup and modNovac finance hardening: guarded SEF form activation/send cleanup, destructive action confirmations, SaveNovac validation, avans split fail-fast, partner-map conflict guard, otkup payment recompute, reset-link TX and Novac smoke suite | Yes | no schema/data migration; compile + RunNovacSmokeSuite passed |
| v6.7 | 2026-04-28 | SEF P0/P1 follow-up hardening: external status persistence on submit, missing `SEFDocumentId` fail-fast, refresh helper convergence, EH preservation, transition matrix suite, stricter cancel/storno tests, mapper total consistency and lightweight parser hardening | Yes | no schema/data migration; compile + SEF offline/state-transition/refresh smoke required |
| v6.6 | 2026-04-27 | Desktop tested baseline: strict BrojZbirne traceability auto-link, professional business-flow regression suite, SEF live submit/refresh evidence, SEF DeliveryDate/InvoiceDate validation, destructive cancel/storno test scaffolding | Yes | no data migration; compile + business-flow + SEF smoke passed; cancel/storno final outcome still P1 |
| v6.5 | 2026-04-26 | SEF P0/P1 desktop hardening, Stammdaten update-guard convergence and frmOtkupAPP shell save/navigation cleanup | Yes | no data migration; compile OK; SEF smoke test required |
---

## 2. Changelog Entries

Slede svi popunjeni changelog unosi iz dostavljenog reference seta, normalizovani u zajednički format.

 ## 3. Version Entries






## v6.12 — 2026-05-04

### Summary
- v6.12 closes the PWA launch-smoke hardening pass after the v6.11 desktop persistence/security/data-health baseline.
- PWA business-date handling was normalized around local-calendar helpers in `utils/format.js`, eliminating UTC date slicing for business dates.
- PWA role sync now returns a canonical result object for Otkupac, Kooperant and Vozac.
- The shared sync engine now has request-level rollback and stale-`syncing` recovery behavior suitable for launch use.
- Kooperant sync now covers both `tretmani` and `troskovi`.
- GAS `syncTrosak` is active and returns the same batch response contract as other sync endpoints.
- PWA client error reporting to GAS `ErrorLog` is active through `reportClientError` and global error/rejection handlers.
- Full PWA smoke passed across Otkupac, Kooperant, Vozac and Management roles.
- `BrojZbirne` ownership was clarified: PWA/GAS own technical `ServerRecordID`, while VBA Master import owns business `BrojZbirne` generation.

### ADDED
- [Layer: PWA/Observability] `reportClientError(error, context)` client error reporting.
  - What changed: PWA now reports selected client exceptions to GAS `logClientError`, which writes to the existing `ErrorLog` sheet.
  - Why: field/browser failures must be visible without relying on user console screenshots.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Formatting] Shared `formatNumber`, `formatKg` and `formatMoney` helpers.
  - What changed: display formatting needed by Otkup Pregled is now owned by `utils/format.js`.
  - Why: feature modules must not depend on undefined formatter globals.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Kooperant Sync] Active `syncTrosak` endpoint and `processTrosakRecord` implementation.
  - What changed: troškovi are no longer blocked by `FEATURE_DISABLED`; GAS now validates records, writes/updates `TROSKOVI-<KooperantID>` and returns `buildBatchSyncResponse(results)`.
  - Why: Kooperant expense sync is required for launch behavior in Knjiga Polja.
  - Reference update required: Yes
  - Migration required: No production data migration; stale test sheets may be deleted/recreated during smoke.

### CHANGED
- [Layer: PWA/Date Handling] Business date generation moved to local-calendar helpers.
  - Previous behavior: some feature code used `toISOString().slice(0, 10)` or `toISOString().split('T')[0]`, causing UTC day drift in Serbia around late-evening local times.
  - New behavior: business date-only fields use `getTodayIsoDate()`, `getRelativeIsoDate(...)`, `toIsoDateOnly(...)` or `localIsoDateFromDate(...)`.
  - Why: PWA, GAS and Sheets must agree on local agricultural business dates.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Sync] Role sync result shape is canonical across roles.
  - Previous behavior: app-level sync and role entrypoints could return arrays, role-only success objects or different per-role shapes.
  - New behavior: `syncQueueSafe()` / role sync returns `{ ok, role, synced, failed, results, reason, code, partial }` for Otkupac, Kooperant and Vozac.
  - Why: UI feedback, diagnostics and retry behavior need one contract.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Sync Engine] Request-level rollback now covers the attempted batch.
  - Previous behavior: rollback logic depended on the local record still being `syncing`, which could leave records non-retryable after request-level failures.
  - New behavior: every record in the attempted pending batch is returned to `pending` on request-level failure unless it has a confirmed server result.
  - Why: launch sync must be recoverable after empty/failed responses.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Kooperant] `syncKooperantNow()` includes both treatment and expense stores.
  - Previous behavior: Kooperant manual/app-level sync was treatment-centric.
  - New behavior: Kooperant sync aggregates `syncTretmani()` and `syncTroskovi()` and returns a role-level result.
  - Why: Knjiga Polja expense data is part of the launch Kooperant workflow.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/GAS/PWA Zbirna Boundary] `BrojZbirne` ownership clarified.
  - Previous behavior: PWA/GAS smoke initially treated empty `BrojZbirne` after `syncZbirna` as a potential GAS bug.
  - New behavior: `ServerRecordID` remains technical PWA/GAS sync ID; `BrojZbirne` is generated by VBA Master import through `GenerateBrojZbirne(vozacID, datum)`.
  - Why: VBA document flow links Otkup/Otpremnica/Zbirna/Prijemnica through business `BrojZbirne`, not technical `ServerRecordID`.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: PWA/Otkup Pregled] Missing kilogram/money formatter crash.
  - Symptom: Otkupac `Danas`/Pregled crashed with `ReferenceError: formatKg is not defined`.
  - Resolution: shared formatter helpers were added to `utils/format.js`.
  - Reference update required: Yes

- [Layer: PWA/Date Display] One-day drift in PWA reads of GAS/Sheets date values.
  - Symptom: rows written with correct `Datum` in Sheets could display as the previous day in PWA views.
  - Resolution: `toIsoDateOnly` and `fmtDate` now normalize parseable timestamps to local date-only values.
  - Reference update required: Yes

- [Layer: GAS/Troškovi Sync] `syncTrosak` returned HTTP 200 with `data = null`.
  - Symptom: GAS inserted/updated a trošak row, but PWA marked the local row `pending` with `lastServerStatus = empty-response`.
  - Resolution: the `syncTrosak` branch now returns `buildBatchSyncResponse(results)` from inside the `withLock` callback.
  - Reference update required: Yes

- [Layer: PWA/Kooperant Runtime] Server-confirmed troškovi can reconcile from prior pending/empty-response state.
  - Symptom: a trošak already visible on the server could remain local `pending` after an empty response during sync.
  - Resolution: Kooperant expense read/merge and subsequent sync confirmation now converge to local `synced` state.
  - Reference update required: Yes

### DEPRECATED
- [Layer: PWA/Date Handling] UTC slicing for business date fields.
  - Replacement: local-calendar date helpers in `utils/format.js`.
  - Removal target: immediate in v6.12 PWA codebase.
  - Reference update required: Yes

- [Layer: GAS/Sync] Disabled `syncTrosak` launch placeholder.
  - Replacement: active `syncTrosak` endpoint and `processTrosakRecord`.
  - Removal target: immediate in v6.12 backend.
  - Reference update required: Yes

### KNOWN ISSUE
- [KI-v6.12-01] VBA VOZ writeback must use the correct `BrojZbirne` payload index.
  - Affected layer: VBA / MasterSync / VOZ sheet writeback
  - Impact: `ImportOneVOZSheet` appends `Array(i, SYNC_STATUS_MASTER, newZbirnaID, brojZbirne)`, so `WriteBackVOZSyncStatus` must write column B from `update(2)` and column T from `update(3)`. Writing `update(2)` to both columns stores the internal/master ID as the business document number.
  - Workaround: update `WriteBackVOZSyncStatus` before treating post-VBA `BrojZbirne` writeback as launch-ready.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.12-01] Complete VBA VOZ `BrojZbirne` writeback correction.
  - Target: next VBA patch before full document-flow launch.
  - Dependency: keep `GenerateBrojZbirne(vozacID, datum)` as owner of business numbering.

- [RM-v6.12-02] Sync deployed GAS `Code.gs` back into repository source.
  - Target: release consistency cleanup.
  - Dependency: deployed GAS currently contains the newer `syncTrosak` implementation and should not remain ahead of repo source.

### VERIFICATION
- Otkupac smoke: new otkup, sync, `getOtkupi`, Danas/Pregled, Otprema, date and ambalaža display passed.
- Kooperant smoke: treatment sync passed after test sheet schema correction; troškovi sync passed after `syncTrosak` response contract fix.
- Vozac smoke: `getVozacOtkupi`, zbirna creation, `syncZbirna`, `getVozacZbirne`, date and ambalaža display passed.
- Management smoke: Management session sanity, `getMgmtAll`, navigation and no unexpected `ErrorLog` rows passed.
- Error reporting smoke: manual `reportClientError` wrote a row to Drive `ErrorLog`.
- Final PWA smoke: no new unexpected `ErrorLog` rows after role smoke.

### Migration Notes
- No production-data migration is required for v6.12.
- Test sheets created with stale headers may be deleted/recreated during smoke because they are not production data.
- Do not add repair/backfill logic for prior smoke rows; test data is disposable.

### Documentation Actions
- [x] Canonical reference updated for PWA launch-smoke contracts.
- [x] Changelog updated for v6.12.
- [x] Sync result contract reviewed.
- [x] GAS `syncTrosak` contract reviewed.
- [x] PWA/VBA `BrojZbirne` ownership boundary documented.


## v6.11 — 2026-04-30

### Summary
- v6.11 closes the pre-launch persistence/security/data-health follow-up after the v6.10 Google/GAS hardening pass.
- AR-002 is implemented centrally: every successful `clsTransaction.CommitTx` now triggers best-effort workbook persistence through `AutoSaveAfterCommit`, with debounce, read-only/path guards, `DisplayAlerts` suppression and non-propagating error handling.
- AutoSave logging records a transaction source derived from snapshotted tables, e.g. `clsTransaction[tblOtkup,tblAmbalaza,tblNovac]`, improving post-crash/post-mortem diagnosis.
- HTTP request construction is consolidated in `modHttpUtils`, replacing duplicated ANSI `UrlEncode` / JSON escape helpers across Google Auth, Google Sheets/Drive, MasterSync and SEF.
- `UrlEncode` now performs RFC 3986 percent-encoding over UTF-8 bytes, preventing Serbian diacritics from being encoded as ANSI codepoints.
- SEF client and validator config now enforce HTTPS-only `SEF_BASE_URL`; plain `http://` is rejected locally before any API key or payload leaves Excel.
- SEF parser risk remains acknowledged but is bounded by baseline parser smoke coverage until the planned VBA-JSON migration.
- Test hygiene and data-health were tightened: the Faktura already-fakturisana-prijemnica fixture no longer leaves `FAK-EXISTING` broken references, and legacy test/demo reference failures are treated as cleanup items before final production launch.

### ADDED
- [Layer: VBA/Persistence] AR-002 central AutoSave after transaction commit.
  - What changed: `clsTransaction.CommitTx` now calls `AutoSaveAfterCommit(sourceName)` after successful commit.
  - Why: closes the crash-loss window between in-memory transaction commit and disk persistence.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Persistence] Transaction-source AutoSave logging.
  - What changed: `BuildAutoSaveSourceName()` captures the snapshotted table list before transaction cleanup and passes it to the AutoSave log.
  - Why: post-mortem logs need to show which business areas were persisted or debounce-skipped.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/HTTP Utilities] `modHttpUtils` canonical helper module.
  - What changed: shared `UrlEncode(s)` and `JsonEscape(s)` are now owned by one canonical module.
  - Why: Google, SEF and future HTTP clients must not carry duplicated request-building helpers with divergent bug fixes.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Tests] HTTP utility and SEF parser smoke coverage.
  - What changed: `RunHttpUtilsSmokeSuite` validates UTF-8 URL encoding and JSON escape behavior; `RunSEFClientParserSmokeSuite` provides baseline coverage for current manual SEF parser behavior.
  - Why: the UTF-8 encoder and JSON parsing surface are cross-client risk points and must be regression-tested before launch.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/Transaction] AutoSave is now transaction-level, not wrapper-by-wrapper.
  - Previous behavior: saving depended on manual workbook save or ad-hoc wrapper-level hooks.
  - New behavior: successful `CommitTx` triggers centralized best-effort AutoSave with debounce.
  - Why: every production commit must have the same persistence contract, including future transaction wrappers.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Google/SEF HTTP] URL encoding consolidated and upgraded to UTF-8.
  - Previous behavior: duplicated helpers used `Asc(ch)` and encoded ANSI/codepage values.
  - New behavior: outbound query/body form parameters use RFC 3986 percent-encoding over UTF-8 bytes.
  - Why: Serbian diacritics and non-ASCII values must not produce malformed Google/SEF requests.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Security] SEF base URL validation is HTTPS-only.
  - Previous behavior: `http://` and `https://` were both accepted as syntactically valid.
  - New behavior: `SEF_BASE_URL` must start with `https://`; plain HTTP raises `ERR_SEF_CONFIG` locally.
  - Why: SEF API key and invoice data must never be sent over unencrypted transport.
  - Reference update required: Yes
  - Migration required: Existing local test configs using `http://` must be corrected.

- [Layer: VBA/Faktura Tests] Already-fakturisana prijemnica fixture no longer creates broken fake faktura references.
  - Previous behavior: the fixture could create active `PRJ-TST-ALRPRJ-*` rows pointing to non-existent `FAK-EXISTING`.
  - New behavior: the test still verifies blocking via `Fakturisano=Da` but does not leave invalid active `FakturaID` references.
  - Why: regression tests must not pollute `ProductionHealthCheck` with permanent broken references.
  - Reference update required: Yes
  - Migration required: Existing demo/test fixture rows should be cleaned or stornirano-marked.

### FIXED
- [Layer: VBA/Persistence] Crash-loss window after successful TX commit.
  - Symptom: a successful in-memory commit could be lost if Excel crashed before manual save.
  - Resolution: centralized post-commit AutoSave with debounce and safe failure handling.
  - Reference update required: Yes

- [Layer: VBA/HTTP] ANSI URL encoding for Serbian characters.
  - Symptom: non-ASCII characters could be encoded as codepage bytes instead of UTF-8 and trigger HTTP 400 behavior.
  - Resolution: central `UrlEncode` uses UTF-8 byte expansion and RFC 3986 unreserved-character rules.
  - Reference update required: Yes

- [Layer: VBA/SEF] Plain HTTP allowed for SEF endpoint config.
  - Symptom: `SEF_BASE_URL` could be configured with `http://`.
  - Resolution: `modSEFClient` and `modSEFValidator` enforce `https://`.
  - Reference update required: Yes

- [Layer: VBA/Test Hygiene] Faktura smoke fixture polluted production-health output.
  - Symptom: repeated test runs added `PRJ-TST-ALRPRJ-* -> FAK-EXISTING` failures.
  - Resolution: test fixture was adjusted and legacy test/demo rows are handled as cleanup data.
  - Reference update required: Yes

### DEPRECATED
- [Layer: VBA/HTTP] Module-local `UrlEncodeGoogle`, private `UrlEncode`, `JsonEscapeGoogle` and private SEF `JsonEscape` helpers.
  - Replacement: `modHttpUtils.UrlEncode` and `modHttpUtils.JsonEscape`.
  - Removal target: immediate in v6.11 codebase.
  - Reference update required: Yes

- [Layer: VBA/Persistence] Relying on operator/manual save after committed transactions.
  - Replacement: central post-commit AutoSave plus backup/journal recovery model.
  - Removal target: immediate in v6.11 codebase.
  - Reference update required: Yes

### KNOWN ISSUE
- [KI-v6.11-01] SEF still uses manual JSON extraction helpers.
  - Affected layer: VBA / SEF client parser
  - Impact: current parser baseline covers simple string/number/bool payloads but does not fully decode all JSON edge cases such as escaped quotes, nested same-name keys, arrays or `null` objects.
  - Workaround: current SEF smoke/parser suite documents baseline behavior; migrate to VBA-JSON in the next controlled P1 pass.
  - Reference update required: Yes

- [KI-v6.11-02] ProductionHealthCheck final readiness depends on workbook data cleanliness.
  - Affected layer: VBA / data health
  - Impact: code/test gates can pass while legacy/demo data still produces reference-integrity failures.
  - Workaround: clean or stornirano-mark demo/test rows before marking a production workbook as final launch-ready.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.11-01] Replace manual SEF JSON extraction with VBA-JSON wrappers.
  - Target: P1 after v6.11 launch hardening.
  - Dependency: add parser regression suite before replacing `ExtractJson*` usage in `ParseSubmitResponse`, `ParseStatusResponse` and `BuildHttpErrorMessage`.

- [RM-v6.11-02] Add optional batch AutoSave suspend/resume for bulk import loops if performance/log noise becomes a problem.
  - Target: P1 if import loops produce excessive save attempts.
  - Dependency: identify real bulk operations that call multiple TX wrappers in a tight loop.

### VERIFICATION
- Compile VBA project after `clsTransaction`, `modJournaling`, `modHttpUtils`, `modSEFClient`, `modSEFValidator`, Google and MasterSync helper changes.
- Run `RunHttpUtilsSmokeSuite` and confirm UTF-8/JSON utility checks pass.
- Run `RunSEFClientParserSmokeSuite` and confirm parser baseline pass.
- Run SEF HTTPS config negative test with temporary `http://` base URL and confirm local fail-fast.
- Run `RunGoogleSyncSmokeSuite` and `RunMasterSyncSmokeSuite` after switching Google/MasterSync to canonical `UrlEncode` / `JsonEscape`.
- Run `RunFakturaSmokeSuite` and confirm `Passed=18 Failed=0` and no new `PRJ-TST-ALRPRJ-* -> FAK-EXISTING` health failure is introduced.
- Run `RunBusinessFlowProSuite` and confirm `Passed=111 Failed=0` with AutoSave logs showing saved/debounce behavior.
- Run `RunProductionHealthCheck`; final production launch requires `Fail=0` after demo/test cleanup.
- Run `RunE2EReleaseGate_v610` or successor release gate; expected status is no runtime failures, with GAS manual gates still documented separately unless integrated.

### RELEASE DECISION
- v6.11 code hardening is launch-ready after compile and smoke/E2E rerun.
- A specific workbook is production-launch-ready only when `RunProductionHealthCheck` is clean after legacy/demo test-data cleanup.

---

## v6.10 — 2026-04-29

### Summary
- v6.10 promotes the Google integration hardening pass across VBA GoogleAuth/GoogleSheets, desktop MasterSync import/writeback, and the GAS `Code.gs` backend.
- VBA Google config ownership was corrected: `tblSEFConfig` remains the central `ConfigKey/ConfigValue` store, while `SetConfigValue` ownership moves to `modConfig` instead of `modGoogleAuth`.
- Google Sheets wrappers now fail fast on clear/write/read/create/find/move failures and include stronger exact-name lookup, non-silent Drive move handling and smoke-test coverage.
- Desktop PWA import paths now guard OTK header schema drift, empty `ClientRecordID`, status writeback failure and Drive-list failure before treating a sync run as successful.
- GAS now fails fast on schema drift, normalizes duplicate lookup, preserves terminal sync statuses on PWA retry, validates critical business fields, hardens token/session behavior and returns partial/batch failure semantics.
- Inactive GAS endpoints `syncTrosak` and `saveOtkupniListPdf` are explicitly disabled instead of leaving dead-route runtime crashes.
- Verification evidence was added through `RunGoogleSyncSmokeSuite`, `runGasRouteHealthCheck` and `runGasSmokeSuite`.

### ADDED
- [Layer: VBA/Google Tests] `RunGoogleSyncSmokeSuite`.
  - What changed: added a Google integration smoke suite covering `tblSEFConfig` schema/config, OAuth token retrieval, spreadsheet create/write/read/find/add-tab and Drive cleanup.
  - Why: GoogleAuth/GoogleSheets transport must be testable independently from business import paths.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Smoke Tests] `runGasRouteHealthCheck`.
  - What changed: added a route/handler presence healthcheck for active `doPost`/`doGet` handler dependencies.
  - Why: `syncTrosak` and `saveOtkupniListPdf` exposed the risk of routing to missing functions.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Smoke Tests] `runGasSmokeSuite`.
  - What changed: added lightweight GAS helper coverage for route health, batch response semantics, authz helpers, login config validation, schema guard behavior and normalized lookup.
  - Why: GAS backend hardening needs repeatable smoke evidence without inserting production OTK/ZBR/TRETMAN/OPREMA business rows.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Sync] Batch response builder for sync actions.
  - What changed: sync batches now report `success=false` when one or more records fail, with `partial`, `OK`, `PARTIAL_FAILURE` and `BATCH_FAILED` semantics.
  - Why: a batch containing failed records must not look like a global success to the PWA.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/Config] Google token/config write ownership moved to `modConfig`.
  - Previous behavior: `SetConfigValue` lived in `modGoogleAuth` while writing into the central `tblSEFConfig` table.
  - New behavior: `GetConfigValue` and `SetConfigValue` live together in `modConfig`; `modGoogleAuth` remains an OAuth client and no longer owns general config writes.
  - Why: `tblSEFConfig` is the central app config store, not a Google-only table.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/GoogleAuth] OAuth token handling hardened.
  - Previous behavior: auth error logging could include raw response bodies and `expires_in` parsing could silently become zero.
  - New behavior: auth error logs redact configured Google secrets/tokens and `expires_in` falls back safely when missing/invalid.
  - Why: auth modules must not leak credentials and must avoid avoidable token-refresh loops.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/GoogleSheets] Sheet write/read/create/find/move flows hardened.
  - Previous behavior: `WriteSheetData` ignored `ClearSheet` failure, Drive move failures were silent, and spreadsheet lookup returned the first parsed ID.
  - New behavior: writes fail if clear fails, Drive move failures are logged, exact-name lookup is required, inputs are validated and HTTP failures log bounded response bodies.
  - Why: partial Google Sheet writes and silent Drive failures create stale rows, orphaned files and wrong-target sync risk.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/MasterSync] OTK import schema/idempotency/writeback hardening.
  - Previous behavior: import relied on positional headers, empty `ClientRecordID` could be treated as non-duplicate, and Google status writeback failure was not visible to the caller.
  - New behavior: OTK header must match canonical GAS `COLUMNS`, empty `ClientRecordID` is blocked, writeback returns success/failure and fatal sync errors can prevent a false transaction commit.
  - Why: schema drift and retry ambiguity are primary risks in the Excel/Google sync bridge.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Schema] `ensureSheetColumns` changed from silent mutation to fail-fast guard.
  - Previous behavior: missing headers were appended to the end of an existing sheet.
  - New behavior: empty sheets receive canonical headers; existing sheet header mismatch/extra named columns raise `SCHEMA_DRIFT`.
  - Why: appending missing columns hides schema drift and breaks VBA header-position contracts.
  - Reference update required: Yes
  - Migration required: Existing malformed sync sheets require manual header repair.

- [Layer: GAS/Sync] Terminal sync statuses are preserved on PWA retry.
  - Previous behavior: existing records could be reset to `Synced`, rewriting `Synced>Master`, `Duplicate` or `SyncError:*` states.
  - New behavior: terminal/master/error statuses are not reset by idempotent retry, and terminal rows are not enriched with new client business fields.
  - Why: PWA retry must not reverse Excel master processing or erase forensic error states.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Auth] Token/session and login config handling hardened.
  - Previous behavior: cache token existence could be treated as valid without payload validation, token generation used `Math.random`, and user role/entity drift could still issue tokens.
  - New behavior: cache and property tokens are payload/expiry validated, malformed/expired tokens are purged, UUID-chain token generation replaces `Math.random`, and `Users` header/role/entity config is validated fail-closed.
  - Why: auth state must fail closed and avoid predictable or stale token behavior.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Observability] Client and login logs sanitized.
  - Previous behavior: `logClientError` and login attempt logging could persist raw token/PIN/password-like values or oversized payloads.
  - New behavior: log payloads are truncated and sensitive fields/base64 blobs are redacted before persistence.
  - Why: operational logging must not become credential leakage or uncontrolled Sheet growth.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA/GoogleSheets] Stale rows could remain after failed clear-before-write.
  - Symptom: a failed `ClearSheet` followed by write to `A1` could leave old rows below newly written data.
  - Resolution: `WriteSheetData` returns failure when `ClearSheet` fails.
  - Reference update required: Yes

- [Layer: VBA/GoogleSheets] Spreadsheet lookup and Drive move were under-observed.
  - Symptom: created sheets could remain outside the PWA folder without clear warning; lookup could trust the first `id` in a response.
  - Resolution: `MoveFileToFolder` returns Boolean/logs HTTP failures and `GetSpreadsheetID` requires exact-name match.
  - Reference update required: Yes

- [Layer: GAS/Routing] Missing handler runtime crashes removed for known inactive features.
  - Symptom: `syncTrosak` and `saveOtkupniListPdf` could route to missing implementation and fail at runtime.
  - Resolution: both endpoints now return explicit `FEATURE_DISABLED` responses until implemented.
  - Reference update required: Yes

- [Layer: GAS/Sync] `processZbirnaRecord` used an undefined `clientRecordID` variable in the insert branch.
  - Symptom: new Zbirna insert could fail or produce invalid ID behavior.
  - Resolution: canonical trimmed `clientRecordID` is now created and reused consistently.
  - Reference update required: Yes

### DEPRECATED
- [Layer: GAS/Sync] Silent schema repair by appending missing columns.
  - Replacement: fail-fast `SCHEMA_DRIFT` and manual/operator-controlled sheet repair.
  - Removal target: immediate in active backend.
  - Reference update required: Yes

- [Layer: GAS/Routing] Treating inactive routes as present but unimplemented.
  - Replacement: explicit `FEATURE_DISABLED` response until implementation exists.
  - Removal target: immediate in active backend.
  - Reference update required: Yes

### KNOWN ISSUE
- [KI-v6.10-01] `syncTrosak` is intentionally disabled.
  - Affected layer: GAS / Kooperant trošak sync
  - Impact: online trošak sync is not active through GAS until `processTrosakRecord` is implemented and smoke-tested.
  - Workaround: local PWA storage/export or later implementation.
  - Reference update required: Yes

- [KI-v6.10-02] `saveOtkupniListPdf` is intentionally disabled.
  - Affected layer: GAS / PDF generation hook
  - Impact: GAS-side otkupni-list PDF generation/save endpoint returns `FEATURE_DISABLED` until the real implementation exists.
  - Workaround: use active PDF upload/export paths that are implemented.
  - Reference update required: Yes

- [KI-v6.10-03] `saveParcelPolygon` remains an intentional public/pre-auth exception.
  - Affected layer: GAS / GIS
  - Impact: the endpoint remains outside the normal token gate by explicit product decision in this hardening pass.
  - Workaround: restrict deployment URL exposure operationally and revisit auth gating if abuse risk increases.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.10-01] Implement `processTrosakRecord` before re-enabling `syncTrosak`.
  - Why it matters: the route must not be re-enabled until its schema, idempotency, validation and smoke coverage match the other sync processors.
  - Affected modules: GAS `Code.gs`, PWA kooperant troškovi flow.
  - Target state: `syncTrosak` active with tested `TROSKOVI_COLUMNS` contract.

- [RM-v6.10-02] Re-enable `saveOtkupniListPdf` only with a real, tested implementation.
  - Why it matters: disabled is safer than runtime crash or stubbed PDF output.
  - Affected modules: GAS `Code.gs`, PWA/desktop PDF flows.
  - Target state: explicit PDF generation/save implementation with route healthcheck and smoke evidence.

- [RM-v6.10-03] Add full integration fixture tests for real OTK/ZBR/TRETMAN/OPREMA insert paths.
  - Why it matters: current GAS smoke suite avoids production business rows; fixture tests should use isolated test sheets/folders.
  - Affected modules: GAS, Google Sheets, VBA MasterSync.
  - Target state: repeatable fixture suite with cleanup.

### TESTED
- [Layer: VBA/Google] `RunGoogleSyncSmokeSuite` passed.
  - Evidence: user reported `TOTAL=29 PASS=29 FAIL=0`.
  - Covered cases: config schema/keys, auth token retrieval, spreadsheet create/write/read/find/add-tab and cleanup.
  - Reference update required: Yes

- [Layer: GAS/Routing] `runGasRouteHealthCheck` passed.
  - Evidence: user reported `GAS route healthcheck PASS: 29 handlers present`.
  - Covered cases: active route handler presence after disabling missing-handler endpoints.
  - Reference update required: Yes

- [Layer: GAS/Smoke] `runGasSmokeSuite` passed.
  - Evidence: user reported `TOTAL=10 PASS=10 FAIL=0`.
  - Covered cases: route health, batch response semantics, authz/login helper behavior, schema guard, normalized lookup and disabled endpoint contract.
  - Reference update required: Yes

### Migration Notes
- No Excel/VBA table schema migration is required.
- No Google business sheet schema migration is required for healthy canonical sync sheets.
- Existing malformed OTK/VOZ/TRETMAN/OPREMA headers must be repaired manually because GAS now fails fast on schema drift instead of appending missing columns.
- `syncTrosak` and `saveOtkupniListPdf` callers must tolerate `FEATURE_DISABLED` until those features are implemented.
- Existing Google OAuth tokens may continue to work, but deployments should run the Google VBA smoke suite after applying GoogleAuth/GoogleSheets changes.

### Verification / Smoke Tests
- Compile VBA project.
- Run `RunGoogleSyncSmokeSuite` and confirm 0 failed.
- Run `runGasRouteHealthCheck()` in Apps Script and confirm all active handlers are present.
- Run `runGasSmokeSuite()` in Apps Script and confirm 0 failed.
- Run existing business suites if the release package also touches finance/document/otkup modules: `RunNovacSmokeSuite`, `RunFakturaSmokeSuite`, `RunBusinessFlowProSuite`.

### Documentation Actions
- [x] Reference updated to v6.10 final
- [x] Changelog v6.10 created
- [x] GoogleAuth/GoogleSheets ownership and fail-fast hardening documented
- [x] MasterSync schema/idempotency/writeback hardening documented
- [x] GAS schema/auth/sync processor hardening documented
- [x] Disabled GAS endpoints documented
- [x] Smoke evidence documented

## v6.9 — 2026-04-29

### Summary
- v6.9 promotes the post-v6.8 hardening pass focused on `modFaktura`, `modDokumenta` and `modOtkup`.
- `CreateFaktura` now treats `PrijemnicaID` as the only trusted caller-provided faktura line input; quantity, price, class and receipt number are read from canonical `tblPrijemnica` rows.
- `UpdateFakturaStatus` now has two-way recompute behavior and preserves existing `DatumPlacanja` instead of overwriting it on every refresh.
- `PrintFaktura` blocks active printing of stornirana faktura rows.
- `modDokumenta` now has stronger EH preservation, base-writer input validation and default stornirano exclusion in core document read helpers.
- `modOtkup` now preserves original errors, validates class and excludes stornirano rows from station/kooperant read helpers.
- Regression coverage was expanded through `RunFakturaSmokeSuite` and the existing `modBusinessFlowProTests.RunBusinessFlowProSuite`.

### ADDED
- [Layer: VBA/Faktura Tests] `RunFakturaSmokeSuite`.
  - What changed: added dev/test-only regression coverage for canonical prijemnica values, duplicate/stornirano/already-fakturisano blocks, status recompute/date preservation and print blocking for stornirana faktura.
  - Why: invoice creation and payment status are finance-critical and must be protected against UI/caller payload mistakes.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Business Flow Tests] Expanded `RunBusinessFlowProSuite` coverage.
  - What changed: added tests for document input validation, document read-helper stornirano exclusion, dual-class wrapper correctness, otkup invalid price/class rejection and otkup read-helper stornirano exclusion.
  - Why: the full document chain is the best end-to-end regression guard for `Otkup -> Otpremnica -> Zbirna -> Prijemnica -> Faktura`.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/Faktura] `CreateFaktura` now derives line values from `tblPrijemnica`.
  - Previous behavior: caller/form-supplied stavka arrays could carry quantity, price, class and receipt number into the faktura creation path.
  - New behavior: caller-supplied stavke are trusted only for `PrijemnicaID`; all financial/line metadata comes from canonical `tblPrijemnica`.
  - Why: business modules must not trust UI payload for invoice financial values when canonical receipt data already exists.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Faktura] `UpdateFakturaStatus` now recomputes in both directions.
  - Previous behavior: paid status/date logic could overwrite `DatumPlacanja` on repeated recompute and was not aligned with the strengthened otkup status recompute pattern.
  - New behavior: sufficient active uplata sets paid status and fills `DatumPlacanja` only if empty; insufficient uplata reopens the faktura and clears `DatumPlacanja`; stornirana rows are skipped.
  - Why: payment removal/storno and repeated refresh must not corrupt payment date or leave stale paid state.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Faktura] `PrintFaktura` blocks stornirana faktura.
  - Previous behavior: direct calls could print a stornirana invoice as if active.
  - New behavior: active `PrintFaktura` raises a business error for stornirana rows; any archival reprint must be a separate marked workflow.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Dokumenta] `_TX` EH preservation and base-writer error propagation aligned with SEF/Novac/Faktura patterns.
  - Previous behavior: some wrapper/base save flows could reduce original validation/relink errors to generic empty-result failures.
  - New behavior: original `Err.Number`, `Err.Source` and `Err.Description` are captured before logging/rollback; `SaveZbirna` and `SavePrijemnica` re-raise original errors.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Dokumenta] Base document input validation strengthened.
  - Previous behavior: Otpremnica/Zbirna/Prijemnica base writers validated only minimal required fields.
  - New behavior: required IDs/numbers, valid class, positive quantity, non-negative/positive price semantics and ambalaža constraints are validated fail-fast before append.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Otkup] Otkup input/read hardening completed.
  - Previous behavior: `SaveOtkup` masked original errors and read helpers returned raw table rows unless caller filtered stornirano.
  - New behavior: `SaveOtkup` propagates original errors, validates class, and `GetOtkupByStation` / `GetOtkupByKooperant` exclude stornirano rows internally.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA/Dokumenta] Corrected returned-ambalaža column naming for `tblPrijemnica`.
  - Symptom: tests failed when code/test assertions used non-canonical `KolAmbalazeVracena`.
  - Resolution: canonical column name is `kolAmbVracena`; writers/tests/constants must use that exact name.
  - Reference update required: Yes

- [Layer: VBA/Faktura] UI/caller tampering risk for invoice line values removed.
  - Symptom: a caller could theoretically pass wrong quantity/price/class/receipt number inside stavka arrays.
  - Resolution: `CreateFaktura` derives those values from `tblPrijemnica` by `PrijemnicaID`.
  - Reference update required: Yes

### TESTED
- [Layer: VBA/Faktura] `RunFakturaSmokeSuite` passed.
  - Evidence: user reported 18 pass / 0 fail.
  - Covered cases: canonical prijemnica values, duplicate selected prijemnica, stornirana prijemnica, already-fakturisana prijemnica, payment-date preservation, reopen on missing payment, stornirana skip and print block.
  - Reference update required: Yes

- [Layer: VBA/Business Flow] Expanded `RunBusinessFlowProSuite` passed.
  - Evidence: user reported all tests passing after correcting the `kolAmbVracena` column name.
  - Covered cases: full document chain, validation rejection, stornirano read-helper exclusion, dual-class document wrappers, otkup invalid price/class and autolink/cross-zbirna audit.
  - Reference update required: Yes

### KNOWN ISSUE
- No new active known issue introduced by v6.9.

### ROADMAP
- [RM-v6.9-01] Keep dev/test modules out of operator UI before production packaging.
  - Why it matters: `modBusinessFlowProTests`, `modFakturaTests`, `modNovacTests` and `modSEFTests` intentionally mutate test data and must remain engineering-only.
  - Affected modules: VBA test modules and desktop navigation.
  - Target state: tests available for engineering regression, inaccessible from normal operator UI.

### Migration Notes
- No Excel table schema migration is required.
- No GAS/Google Sheet schema migration is required.
- Existing data does not require migration.
- If any VBA constant or helper still points to `KolAmbalazeVracena`, update it to `kolAmbVracena`.
- Compile VBA project after applying changes.

### Verification / Smoke Tests
- Compile VBA project.
- Run `RunFakturaSmokeSuite` and confirm 0 failed.
- Run expanded `RunBusinessFlowProSuite` and confirm 0 failed.
- Run `RunNovacSmokeSuite` if `modNovac` was touched in the same release package.
- Confirm SEF smoke tests remain available for SEF-specific changes.

### Documentation Actions
- [x] Reference updated to v6.9 final
- [x] Changelog v6.9 created
- [x] `modFaktura` canonical prijemnica-value rule documented
- [x] `UpdateFakturaStatus` two-way recompute documented
- [x] `PrintFaktura` storno guard documented
- [x] `modDokumenta` EH/input/stornirano hardening documented
- [x] `tblPrijemnica.kolAmbVracena` canonical name documented
- [x] `modOtkup` validation/read-helper hardening documented
- [x] Regression evidence documented

## v6.8 — 2026-04-29

### Summary
- v6.8 promotes the post-SEF cleanup pass focused on `frmSEF` and `modNovac`.
- `frmSEF` is hardened as an operator shell: guarded activation, explicit two-column faktura combo setup, safe send-button cleanup, and extra confirmation for destructive cancel/storno operations.
- `modNovac` is hardened across finance-critical read/write paths: money-row validation, fail-fast avans split, partner-map conflict detection, stornirano-safe reads, guarded required columns, otkup payment recompute and reset-link transaction coverage.
- The update confirms the active transaction model as snapshot/rollback, not pending-write. Direct `AppendRow` writes remain valid inside `_TX` flows when affected tables are snapshotted before mutation.
- `modNovacTests.RunNovacSmokeSuite` now provides regression evidence for validation, stornirano exclusion, partner-map conflict, partial buyer/otkup avans split and otkup reset-link status recompute.

### ADDED
- [Layer: VBA/Finance Tests] `modNovacTests.RunNovacSmokeSuite`.
  - What changed: added dev/test-only finance smoke coverage for `modNovac`.
  - Why: avans split, stornirano exclusion and partner-map conflict are finance-critical edge cases that need regression evidence.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Finance] `GetIsplataForOtkup` alias.
  - What changed: added a clearer alias over the historical `GetUplataForOtkup` helper, which actually aggregates linked `Isplata` by `OtkupID`.
  - Why: preserve caller compatibility while allowing new code to use the correct financial wording.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Finance] `ResetNovacOtkupLink_TX`.
  - What changed: reset of money-to-otkup links is now available as a transaction wrapper that snapshots both `tblNovac` and `tblOtkup`.
  - Why: unlink and otkup paid-status recompute must commit or rollback together.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/SEF UI] `frmSEF` activation and send command handling hardened.
  - Previous behavior: activation had limited guarding and `btnPosalji_Click` could leave the send button disabled after early validation/user-cancel exits.
  - New behavior: activation uses guarded EH; `btnPosalji_Click` uses a cleanup path that restores button state and recomputes allowed actions on all exits.
  - Why: operator-facing forms must not degrade the desktop lifecycle or leave controls disabled after non-error exits.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF UI] `frmSEF` faktura combo setup made explicit.
  - Previous behavior: code wrote to a second combo column while depending on designer-time column setup.
  - New behavior: `LoadFaktureIntoCombo()` explicitly configures `ColumnCount`, `ColumnWidths` and `BoundColumn`.
  - Why: form behavior should not depend on hidden designer state.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF UI] Cancel/storno UI flow now requires explicit destructive confirmation.
  - Previous behavior: comment entry was required but final destructive confirmation was not enforced in the reviewed form code.
  - New behavior: cancel and storno require comment plus explicit confirmation before calling SEF mutation functions.
  - Why: cancel/storno change real external SEF state and must be deliberate.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Finance] `SaveNovac` changed from append-only writer to validated writer.
  - Previous behavior: invalid combinations such as both `Uplata` and `Isplata`, negative amounts or zero/zero rows could be accepted if caller validation missed them.
  - New behavior: `SaveNovac` validates money direction, `Tip`, partner/entity context and amount sign before append.
  - Why: finance core writer must reject impossible rows even if a form or caller fails to validate.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Finance] Avans allocation changed to fail-fast split/update logic.
  - Previous behavior: `ApplyAvansToFaktura` and `ApplyAvansToOtkup` could update original avans rows or create split rows without checking every write result.
  - New behavior: required updates use `RequireUpdateCell`, missing avans rows raise explicit errors, and split row creation must return a non-empty `NOV-*` ID.
  - Why: partial avans consumption must not silently create inconsistent money allocation state.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Finance] `UpdateOtkupStatus` changed to two-way recompute.
  - Previous behavior: the helper could mark an otkup as paid but did not reliably clear paid status when linked payment was removed or became insufficient.
  - New behavior: linked `Isplata` greater than or equal to otkup value sets paid status; insufficient linked payment clears `Isplaceno` and `DatumIsplate`.
  - Why: storno/relink/reset flows must be able to recompute payment state accurately.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA/Finance] Stornirano money rows no longer affect key finance aggregates.
  - Symptom: some `tblNovac` read helpers could count stornirano rows in invoice payment, open-faktura, buyer payment-by-fruit and avans allocation logic.
  - Resolution: key read helpers now exclude stornirano rows before aggregation/allocation.
  - Reference update required: Yes

- [Layer: VBA/Finance] Partner-map conflicts no longer return silent success.
  - Symptom: saving a bank partner name already mapped to a different partner could return `True`.
  - Resolution: identical existing mapping remains idempotent success; conflicting mapping raises a fail-fast error.
  - Reference update required: Yes

- [Layer: VBA/Finance] Filtered-array update index bug fixed.
  - Symptom: using an `ExcludeStornirano` filtered array together with physical row indexes from `FindRows` / `UpdateCell` could produce `Subscript out of range` or wrong-row update risk.
  - Resolution: update flows use full table arrays / physical row indexes and skip `Stornirano="Da"` rows manually; read-only aggregates may still use filtered arrays.
  - Reference update required: Yes

- [Layer: VBA/Finance] Remaining required finance column lookups hardened.
  - Symptom: helper functions such as `GetBankaByPartner`, `LookupPartnerMap`, `GetVrstaFromFaktura`, `BuildVrstaFakturaCache`, `GetUplataForOtkup`, `GetOMAvansSaldo` and `GetAgroAbzug` could use column index `0` if schema drift occurred.
  - Resolution: required columns now use `RequireColumnIndex`; optional columns require explicit zero-index handling.
  - Reference update required: Yes

### TESTED
- [Layer: VBA/Finance] `RunNovacSmokeSuite` passed.
  - Evidence: user reported all tests passing after fixes.
  - Covered cases: invalid both-direction money row rejection, valid uplata, valid isplata, stornirano novac exclusion, partner-map conflict block, partial buyer avans split, partial otkup avans split, and reset-link status recompute.
  - Reference update required: Yes

- [Layer: VBA/SEF UI / Finance] Compile status confirmed.
  - Evidence: user reported the patches compiled through the `frmSEF` and `modNovac` passes.
  - Reference update required: Yes

### KNOWN ISSUE
- [KI-v6.8-01] `GetUplataForOtkup` remains as a historical compatibility name.
  - Affected layer: VBA / finance
  - Impact: the helper actually sums linked `Isplata` values by `OtkupID`; the clearer `GetIsplataForOtkup` alias should be preferred in new code.
  - Workaround: keep the alias and migrate call sites gradually.
  - Reference update required: Yes

- [KI-v6.8-02] Transaction journal is still intent/recovery style, not pure commit-only audit.
  - Affected layer: VBA / data access / transaction observability
  - Impact: snapshot rollback restores tables, but journal rows emitted during failed transactions may remain as intent/recovery traces.
  - Workaround: interpret journal as recovery/diagnostic evidence, not a strict committed-state ledger.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.8-01] Review `UpdateFakturaStatus` for two-way recompute parity with `UpdateOtkupStatus`.
  - Why it matters: buyer invoice status should have the same recompute reliability as otkup paid status.
  - Affected modules: `modFakturisanje`, `modNovac`.
  - Target state: faktura status recompute handles payment removal/storno as well as payment closure.

- [RM-v6.8-02] Production packaging cleanup for dev-only smoke suites.
  - Why it matters: `modNovacTests`, `modSEFTests` and business-flow tests are valuable regression tools but must not be operator-facing.
  - Affected modules: VBA test modules and desktop navigation.
  - Target state: dev/test modules remain available to engineering and inaccessible from normal operator UI.

### Migration Notes
- No Excel table schema migration is required.
- No GAS/Google Sheet schema migration is required.
- Existing data does not require migration.
- VBA compile requires updated `frmSEF`, `modNovac` and `modNovacTests` if the smoke suite is included.

### Verification / Smoke Tests
- Compile VBA project.
- Run `RunNovacSmokeSuite`.
- Confirm `RunNovacSmokeSuite` passes validation, stornirano exclusion, partner-map conflict, partial buyer avans split, partial otkup avans split and reset-link recompute.
- Open `frmSEF`, verify faktura combo loads two columns, send button re-enables after cancel/validation exits, and cancel/storno prompts require explicit confirmation.

### Documentation Actions
- [x] Reference updated to v6.8 final
- [x] Changelog v6.8 created
- [x] `frmSEF` operator-shell hardening documented
- [x] `modNovac` validation and avans split hardening documented
- [x] Partner-map conflict guard documented
- [x] Filtered-array update-index rule documented
- [x] `RunNovacSmokeSuite` evidence documented


## v6.7 — 2026-04-28

### Summary
- v6.7 closes the SEF P0/P1 follow-up items identified after the v6.6 live baseline and Git review.
- The update keeps the existing outbound status model: `SEFWorkflowState` is internal/local process control and `SEFStatus` is the exact latest external SEF/API status. No `WF_SEF_DRAFT` is introduced.
- Submit result persistence now writes `response.apiStatus` into `SEFStatus` instead of internal workflow constants, and successful submit without `SEFDocumentId` is normalized to a technical failure.
- SEF refresh, tests, mapper, parser and error-handling paths were hardened without changing business status semantics.
- Cancel/storno destructive test semantics are stricter: service Boolean result is asserted, already-`STORNO` invoices are expected SKIP, and final external outcome is separated from API-call smoke evidence.

### ADDED
- [Layer: VBA/SEF Tests] `RunSEFStateTransitionSuite`.
  - What changed: added an offline transition-matrix suite around `ValidateAllowedTransition`.
  - Why: the SEF state machine now has explicit regression coverage for allowed and blocked local transitions, including terminal `WF_SEF_STORNO`.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Mapper] `ValidateSEFTotalMatch`.
  - What changed: mapper now fail-fast checks that header totals match line totals for net, VAT and gross within tolerance.
  - Why: UBL payloads should not be submitted when `tblFakture` totals and `tblFakturaStavke` line sums disagree.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Client] Lightweight parser helpers.
  - What changed: added numeric-or-string document ID extraction and tolerant simple boolean parsing.
  - Why: SEF responses can vary between numeric and string ID representations and may format boolean fields with whitespace.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/SEF Service] Submit result persistence now stores external API status.
  - Previous behavior: `SendInvoiceToSEF_TX` could write internal workflow constants such as `WF_SEF_SENT`, `WF_SEF_ACCEPTED` or `WF_SEF_REJECTED` into `SEFStatus`.
  - New behavior: submit result persistence writes `response.apiStatus` into `SEFStatus`; `SEFWorkflowState` remains the internal workflow field.
  - Why: preserve the v6.6/v6.7 canonical split between internal workflow and external SEF status.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Service] Missing `SEFDocumentId` after successful submit is now technical failure evidence.
  - Previous behavior: a successful response with no `SEFDocumentId` could still progress toward a successful outbound workflow state or be treated too softly by tests.
  - New behavior: `response.Success=True` without `SEFDocumentId` is normalized to `FAILED` / `MISSING_SEF_DOCUMENT_ID` and persisted as a technical failure path.
  - Why: refresh/cancel/storno require stable SEF document identity; successful outbound state without document ID is invalid.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Status Sync] Pending/non-final refresh routing changed to helper-based convergence.
  - Previous behavior: some pending status branches could still perform direct workflow/update logic.
  - New behavior: `SENT`, `NEW`, `DRAFT` and non-final fallback statuses route through `ApplySEFStateOrRefreshOnly`.
  - Why: avoid same-state errors and prevent accidental backwards transitions from final local states.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Tests] Live send test semantics changed.
  - Previous behavior: missing `SEFDocumentId` after a successful outbound workflow could be logged as SKIP/PASS.
  - New behavior: `WF_SEF_SENT` or `WF_SEF_ACCEPTED` without `SEFDocumentId` is FAIL; returned submission ID must exist and match the faktura's last submission.
  - Why: successful submit evidence is incomplete without document identity.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Tests] Cancel/storno destructive test assertions changed.
  - Previous behavior: cancel/storno smoke could pass mainly from event/status population.
  - New behavior: tests assert Boolean return from `CancelInvoiceOnSEF_TX` / `StornoInvoiceOnSEF_TX`, classify already-`STORNO` as SKIP, and only mark final outcome verified when external status is cancel-like or `STORNO`.
  - Why: separate API/event smoke from final business outcome evidence and avoid false positives.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA/SEF Service] Original error preservation fixed in remaining SEF mutation/recovery handlers.
  - Symptom: validator/API error details could be masked if `LogErr` or rollback ran before capturing `Err`.
  - Resolution: `CancelInvoiceOnSEF_TX`, `StornoInvoiceOnSEF_TX`, `RecoverStuckSEFSendingInvoice` and `PrepareRejectedInvoiceForResubmit` now capture `Err.Number`, `Err.Description` and `Err.Source` before logging/rollback.
  - Reference update required: Yes

- [Layer: VBA/SEF Validator] `WF_SEF_STORNO` terminal transition handling fixed.
  - Symptom: `WF_SEF_STORNO` existed and was reachable but had no explicit old-state case in `ValidateAllowedTransition`.
  - Resolution: added terminal blocked transition behavior for `WF_SEF_STORNO`.
  - Reference update required: Yes

- [Layer: VBA/SEF Tests] Already-STORNO storno test classification fixed.
  - Symptom: attempting storno on an already externally `STORNO` invoice could appear as test failure even though validator behavior was correct.
  - Resolution: destructive storno test pre-checks `SEFStatus=STORNO` and reports SKIP.
  - Reference update required: Yes

- [Layer: VBA/SEF Client] Simple response parser robustness improved.
  - Symptom: document ID parsing could fail if ID appeared as a string instead of a JSON number; accepted detection could fail with whitespace around boolean values.
  - Resolution: parser now uses numeric-or-string ID extraction and `ExtractJsonBoolean`.
  - Reference update required: Yes

### TESTED
- [Layer: VBA/SEF] Compile status confirmed after v6.7 patches.
  - Evidence: user reported all patches compiled/worked during the hardening pass.
  - Reference update required: Yes

- [Layer: VBA/SEF Tests] State-transition suite passed.
  - Evidence: `RunSEFStateTransitionSuite` was added and reported as working.
  - Reference update required: Yes

- [Layer: VBA/SEF Tests] Live/test semantics hardened.
  - Evidence: missing-document-ID live-send test patch compiled and passed; cancel/storno assertion patch was applied.
  - Reference update required: Yes

### KNOWN ISSUE
- [KI-v6.7-01] SEF cancel final business outcome still needs controlled final-status evidence.
  - Affected layer: VBA / SEF cancel
  - Impact: v6.7 distinguishes API success from final cancel-like external status, but a complete matrix of allowed cancel outcomes still needs live confirmation.
  - Workaround: treat API smoke and final cancel-like status as separate evidence.
  - Reference update required: Yes

- [KI-v6.7-02] Full JSON parser is still not introduced.
  - Affected layer: VBA / SEF client
  - Impact: numeric/string ID and simple boolean handling are improved, but nested/escaped JSON remains a limitation of the lightweight parser.
  - Workaround: keep parser usage limited to known simple SEF response fields.
  - Reference update required: Yes

- [KI-v6.7-03] Dev-only `Test_*` routines remain in some production SEF modules.
  - Affected layer: VBA / packaging
  - Impact: no runtime business bug, but production module cleanliness and accidental invocation risk remain.
  - Workaround: move to `modSEFTests` or clearly mark as dev-only before final packaging.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.7-01] Complete SEF accepted/final lifecycle evidence.
  - Why it matters: live submit and `SENT` refresh are proven; controlled `ACCEPTED` evidence still closes the buyer-side final lifecycle.
  - Affected modules: `modSEFStatusSync`, `modSEFTests`, `frmSEF`.
  - Target state: external `ACCEPTED` refresh converges idempotently with event/submission evidence.

- [RM-v6.7-02] Complete cancel/storno final outcome matrix.
  - Why it matters: destructive corrective actions need allowed/disallowed and final-status evidence, not just API-call smoke.
  - Affected modules: `modSEFTests`, `modSEFService`, `modSEFStatusSync`.
  - Target state: cancel/storno allowed, blocked and already-final scenarios have stable PASS/SKIP/FAIL semantics.

- [RM-v6.7-03] Move dev-only SEF `Test_*` procedures out of production modules.
  - Why it matters: package hygiene before production operator rollout.
  - Affected modules: `modSEFService`, `modSEFClient`, `modSEFStatusSync`, `modSEFMapper`, `modSEFTests`.
  - Target state: one formal test module holds dev-only test routines.

### Migration Notes
- No Excel table schema migration is required.
- No GAS/Google Sheet schema migration is required.
- Existing SEF rows do not need data migration; future submit/result writes will preserve external `SEFStatus` semantics more precisely.
- Compile requires the updated SEF service, tests, status sync, validator, mapper and client modules.

### Verification / Smoke Tests
- Compile VBA project.
- Run `RunSEFOfflineSuite` on a valid current-date SEF candidate.
- Run `RunSEFStateTransitionSuite`.
- Run `RunSEFRefreshIdempotencySuite` on a previously submitted faktura with `SEFDocumentId`.
- For a new dummy current-date faktura, run `RunSEFLiveSendSuite` if live SEF testing is intentionally enabled.
- Run destructive cancel/storno suites only with explicit config/user confirmation and only on intended test invoices.
- Verify that successful outbound workflow states never exist without `SEFDocumentId`.

### Documentation Actions
- [x] Reference updated to v6.7 final
- [x] Changelog v6.7 created
- [x] SEF external status persistence on submit documented
- [x] Missing `SEFDocumentId` failure rule documented
- [x] Refresh helper convergence documented
- [x] EH preservation cleanup documented
- [x] State-transition suite documented
- [x] Mapper total consistency documented
- [x] Parser hardening documented
- [x] Cancel/storno stricter test semantics documented


## v6.6 — 2026-04-27

### Summary
- v6.6 promotes the post-v6.5 tested baseline into the canonical reference/changelog set.
- The update is not limited to SEF tests: it also documents the professional business-flow regression suite, the P1 traceability auto-link bug found by that suite, and the strict `BrojZbirne` auto-link fix.
- SEF is upgraded from “hardened but smoke-test pending” to a partially live-tested baseline: valid submit + refresh passed, invalid receiver rejection persisted correctly, and local UBL date validation blocks known-bad payloads before HTTP.
- The outbound status model is clarified: `SEFWorkflowState` is local process control, while `SEFStatus` is the exact latest external SEF API status. They are not required to match.
- Cancel/storno destructive tests were introduced, but final cancel/storno business-outcome certification remains a P1 known issue.

### ADDED
- [Layer: VBA/Tests] `modBusinessFlowProTests` professional regression suite.
  - What changed: added a dev/test-only suite that seeds an empty workbook and runs Otkup → Otpremnica → Zbirna → Prijemnica → Faktura with validation, duplicate, traceability and audit checks.
  - Why: the system previously had no reliable regression suite for an empty workbook.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Tests] SEF live/negative/destructive test coverage.
  - What changed: `modSEFTests` now covers offline DTO/UBL checks, live submit, SEF rejection persistence, local date-validation blocking, repeated refresh idempotency and cancel/storno destructive scaffolding.
  - Why: SEF must be proven through real API and state-machine scenarios, not only compile/offline checks.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Mapper] `DeliveryDate` added to `clsSEFInvoiceSnapshot`.
  - What changed: DTO now carries both `InvoiceDate` and `DeliveryDate`.
  - Why: serializer and validator must operate on the same business dates.
  - Reference update required: Yes
  - Migration required: No schema migration; class property/code compile update required.

- [Layer: VBA/SEF Tests] `CreateSEFLiveDummyFaktura` / equivalent live fixture helper.
  - What changed: controlled dummy invoices can be created with real current dates for SEF live tests.
  - Why: business-flow tests intentionally use far-future dates for isolation, which are invalid for SEF UBL live submit.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Dev Ops] dev-only full table clear/reset utility pattern.
  - What changed: a controlled utility can remove all ListObject data from a test workbook while preserving table structure.
  - Why: repeated smoke/regression tests need a clean workbook when no production data exists.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/Sledljivost] Auto-link canonical key changed to include `BrojZbirne`.
  - Previous behavior: `AutoLinkOtkupOtpremnica` could match only on `StanicaID + Datum + VozacID + Klasa`.
  - New behavior: strict match uses `StanicaID + Datum + VozacID + Klasa + BrojZbirne`; legacy fallback without `BrojZbirne` is allowed only when `BrojZbirne` is missing and the match is unique.
  - Why: prevent cross-zbirna wrong links when multiple otkup groups share the same station/date/driver/class.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Mapper] UBL date ownership changed to DTO-owned dates.
  - Previous behavior: serializer could recalculate or derive delivery date independently.
  - New behavior: `BuildSEFInvoiceDto` sets `InvoiceDate` and `DeliveryDate`, and `SerializeUBLInvoice` renders `dto.InvoiceDate` and `dto.DeliveryDate`.
  - Why: serializer and validator must not diverge.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Mapper] Delivery date derivation changed to invoice-line based lookup.
  - Previous behavior: delivery date logic was incomplete/inconsistent and could depend on the wrong call shape.
  - New behavior: `GetInvoiceDeliveryDate(fakturaID)` resolves `tblFakturaStavke.FakturaID -> PrijemnicaID -> tblPrijemnica.Datum` and uses the latest linked prijemnica date.
  - Why: SEF delivery date must represent the faktura’s actual linked receiving documents.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Status] Outbound status model clarified.
  - Previous behavior: `SEF_SENT` could be misunderstood as external status exactly equal to `SENT`.
  - New behavior: `SEFWorkflowState` is local process control and `SEFStatus` is the exact external SEF status. `SEFWorkflowState=SEF_SENT` can coexist with `SEFStatus=DRAFT`, `SENT`, `STORNO` etc.
  - Why: logs showed valid cases where submit succeeded locally while refresh returned `DRAFT`, then later `SENT`.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Tests] Cancel/storno tests changed to destructive, gated tests.
  - Previous behavior: cancel/storno were in the smoke matrix but not covered by explicit destructive test scaffolding.
  - New behavior: live cancel/storno suites require config flag and confirmation.
  - Why: these calls mutate real SEF state and must not run accidentally.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA/Sledljivost] Cross-`BrojZbirne` auto-link bug fixed.
  - Symptom: an otkup row from `BrojZbirne=A` could receive an `OtpremnicaID` from `BrojZbirne=B` if station/date/driver/class matched.
  - Resolution: strict key includes `BrojZbirne`; regression test now verifies that mismatched zbirna remains unlinked while the matching zbirna links correctly.
  - Reference update required: Yes

- [Layer: VBA/SEF Mapper] Known SEF UBL date rejection is prevented locally.
  - Symptom: SEF rejected payloads with `UBLDeliveryDateMustNotBeLatterThanIssueDate`.
  - Resolution: `ValidateSEFDtoForUBL` blocks `DeliveryDate > InvoiceDate` before HTTP submit.
  - Reference update required: Yes

- [Layer: VBA/SEF Mapper] Serializer date divergence fixed.
  - Symptom: `SerializeUBLInvoice` could use a local delivery-date calculation instead of the DTO date.
  - Resolution: serializer must render `dto.DeliveryDate`.
  - Reference update required: Yes

- [Layer: VBA/SEF Tests] Live submit baseline fixed from false fixture selection.
  - Symptom: trying to live-submit invoices created by far-future business-flow tests caused local date validation SKIP.
  - Resolution: live SEF tests use current-date dummy invoices and validate linked prijemnica dates through debug helper evidence.
  - Reference update required: Yes

### TESTED
- [Layer: VBA/Business Flow] Professional business-flow suite passed.
  - Evidence: empty-workbook seed and full document chain tests passed, including autolink positive/negative regression.
  - Reference update required: Yes

- [Layer: VBA/SEF] Live submit passed.
  - Evidence: `FAK-00003` submitted successfully with HTTP 200, `SubmissionStatus=SENT`, `SEFDocumentId=5317568`, local workflow `LOCAL_FINALIZED -> SEF_SENT`.
  - Reference update required: Yes

- [Layer: VBA/SEF] Refresh idempotency passed.
  - Evidence: repeated refresh of `FAK-00003` passed twice; workflow remained `SEF_SENT` and external status reached `SENT`.
  - Reference update required: Yes

- [Layer: VBA/SEF] Negative SEF validation paths passed.
  - Evidence: invalid receiver PIB produced persisted `ReceiverCompanyNotFound`; future delivery date was blocked locally by `DeliveryDate > InvoiceDate`.
  - Reference update required: Yes

- [Layer: VBA/SEF] Cancel/storno scaffolding executed.
  - Evidence: cancel API smoke wrote event log and refreshed status; storno validator correctly blocked already-`STORNO` invoice.
  - Reference update required: Yes

### KNOWN ISSUE
- [KI-v6.6-01] Cancel final business outcome is not fully certified.
  - Affected layer: VBA / SEF cancel
  - Impact: cancel API/event smoke can pass while the external status after refresh may still require interpretation or follow-up.
  - Workaround: treat cancel test as API smoke until final cancel-like external status is verified.
  - Reference update required: Yes

- [KI-v6.6-02] Storno already-`STORNO` scenario needs expected-SKIP classification.
  - Affected layer: VBA / SEF tests
  - Impact: validator behavior is correct, but the test should not report a failure when the invoice is already externally `STORNO`.
  - Workaround: pre-check `SEFStatus=STORNO` and log SKIP.
  - Reference update required: Yes

- [KI-v6.6-03] `StornoInvoiceOnSEF_TX` EH can still mask the original validator message.
  - Affected layer: VBA / SEF service
  - Impact: original messages can be replaced by generic/invalid-procedure errors.
  - Workaround: capture original `Err.Number`, `Err.Description`, `Err.Source` before logging/rollback.
  - Reference update required: Yes

- [KI-v6.6-04] Final `ACCEPTED` status refresh still lacks live evidence.
  - Affected layer: VBA / SEF status sync
  - Impact: `SENT` live baseline is proven, but accepted-side final lifecycle still needs a controlled test.
  - Workaround: keep refreshing a receiver-accepted test invoice or perform an accepted-flow test when available.
  - Reference update required: Yes

- [KI-v6.6-05] Dev/test modules must not become operator surface.
  - Affected layer: VBA / tests / packaging
  - Impact: reset and destructive tests can mutate data.
  - Workaround: keep dev-only modules hidden from UI and remove/lock before production packaging if required.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.6-01] Complete cancel/storno final outcome certification.
  - Why it matters: SEF corrective actions are destructive and must be proven beyond API smoke.
  - Affected modules: `modSEFTests`, `modSEFService`, `modSEFStatusSync`.
  - Target state: allowed/disallowed cancel and storno scenarios have PASS/SKIP/FAIL semantics and preserved original errors.

- [RM-v6.6-02] Capture live `ACCEPTED` refresh evidence.
  - Why it matters: closes the remaining final-status path after successful submit.
  - Affected modules: `modSEFStatusSync`, `frmSEF`, `modSEFTests`.
  - Target state: external `ACCEPTED` refresh converges local workflow to `SEF_ACCEPTED` idempotently.

- [RM-v6.6-03] Keep test modules as formal dev/test layer.
  - Why it matters: the suites found real bugs and should remain usable, but must not be operator-accessible.
  - Affected modules: `modBusinessFlowProTests`, `modSEFTests`, reset helper.
  - Target state: repeatable smoke suite exists outside normal operator navigation.

### Migration Notes
- No Excel table schema migration is required.
- No GAS/Google Sheet schema migration is required.
- VBA compile requires `clsSEFInvoiceSnapshot.DeliveryDate`, the updated SEF mapper date guard and strict BrojZbirne auto-link implementation.
- Existing far-future test fixture invoices are expected to be locally blocked by SEF date validation and should not be used as positive live SEF fixtures.

### Verification / Smoke Tests
- Compile VBA project.
- Run `RunBusinessFlowProSuite` on a clean test workbook.
- Confirm cross-BrojZbirne autolink negative regression passes.
- Create a current-date SEF dummy faktura and run `RunSEFLiveSendSuite`.
- Run `RunSEFRefreshIdempotencySuite` on the submitted faktura.
- Verify invalid receiver rejection persists `ErrorCode`/`ErrorMessage`.
- Verify `DeliveryDate > InvoiceDate` is locally blocked before HTTP.
- Run cancel/storno destructive suites only with explicit config/user confirmation.
- Treat already-STORNO storno as expected SKIP after test cleanup.

### Documentation Actions
- [x] Reference updated to v6.6 final
- [x] Changelog v6.6 created
- [x] Strict BrojZbirne autolink rule documented
- [x] Business-flow professional test suite documented
- [x] SEF live submit/refresh evidence documented
- [x] SEF DeliveryDate/InvoiceDate validation documented
- [x] SEF dual-status model clarified
- [x] Cancel/storno test scaffolding and known issues documented


## v6.5 — 2026-04-26

### Summary
- v6.5 promotes the post-v6.4 cleanup work into a full canonical version.
- `frmStammdaten` update writes were aligned with the global v6.4/v6.5 `modSchemaGuard.RequireUpdateCell` standard.
- `frmOtkupAPP` main shell behavior was hardened so startup setup, save button behavior and matični-podaci navigation are safer and operator-facing.
- The SEF subsystem received P0/P1 hardening across status sync, mapper, persistence, validator, client and service modules.
- SEF status refresh is now idempotent, tax calculation is consistent between DTO and UBL, persistence writes use `RequireUpdateCell`, validator config checks are active, HTTP client handling is centralized, and stuck `SEF_SENDING` recovery is explicit.

### ADDED
- [Layer: VBA/SEF] Idempotent SEF status refresh helper behavior.
  - What changed: refresh can update status fields without forcing a workflow transition when the local state already matches the target state.
  - Why: repeated refresh of `ACCEPTED`, `REJECTED`, `SENT`, `NEW` or `DRAFT` must be safe.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF] SEF persistence schema guard helpers.
  - What changed: `RequireFaktureSEFSchema`, `RequireSEFSubmissionSchema`, `RequireSEFEventLogSchema` and `GetFakturaSEFFieldText` were introduced/standardized in `modSEFPersistance`.
  - Why: SEF reads/writes must fail fast on missing columns instead of returning misleading blanks.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF] Centralized SEF HTTP client helpers.
  - What changed: `GetSEFClientConfig`, `CreateSEFHttpRequest`, `ApplySEFHeaders`, `ApplyRateLimitResponse`, `GetJsonNumericIdLiteral` and debug helpers were added/standardized.
  - Why: submit/status/cancel/storno calls need one config/header/rate-limit pattern.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF] Stuck `SEF_SENDING` recovery flow.
  - What changed: `RecoverStuckSEFSendingInvoice` and guarded `RecoverAllStuckSEFSendingInvoices` handle invoices stuck in sending state.
  - Why: operator recovery must not require manual table edits.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/Stammdaten] `frmStammdaten.btnIzmeni_Click` changed from local `MustUpdateCell` update guard usage to canonical `RequireUpdateCell`.
  - Previous behavior: update writes could use a private form-local helper.
  - New behavior: critical update writes use `modSchemaGuard.RequireUpdateCell`.
  - Why: avoid duplicated fail-fast update semantics.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Main Shell] `frmOtkupAPP` initialization and navigation were hardened.
  - Previous behavior: shell setup and save/navigation behavior could drift; `btnSnapshot_Click` did not necessarily perform an actual save.
  - New behavior: initialization has EH, `SetupShell` is restored in the setup chain, snapshot calls `SaveApp`, and matični-podaci navigation has explicit EH.
  - Why: the main shell must reliably initialize, save and navigate without business writes in shell code.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Mapper] VAT/tax calculation changed to one source of truth.
  - Previous behavior: mapper DTO values used hardcoded 10% VAT while UBL serialization used config.
  - New behavior: DTO and UBL use `GetDefaultTaxPercent`.
  - Why: prevent monetary/tax mismatch in outbound SEF payloads.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Mapper] `SEF_PAYMENT_DUE_DAYS` parsing changed to guarded config parsing.
  - Previous behavior: raw `CLng` could fail unclearly.
  - New behavior: empty config defaults to 15 days, invalid values raise `ERR_SEF_CONFIG`.
  - Why: invalid config should fail before send.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Persistence] SEF write helpers changed to `RequireUpdateCell`.
  - Previous behavior: local `ok = UpdateCell(...)` plus `RaiseUpdateError` pattern.
  - New behavior: canonical fail-fast update helper is used for faktura, submission and event state writes.
  - Why: align SEF with v6.5 schema/update guard standard.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Validator] `ValidateSEFConfig` changed from placeholder to active validation.
  - Previous behavior: SEF config could fail later in the HTTP layer.
  - New behavior: `SEF_BASE_URL`, `SEF_API_KEY` and URL scheme are validated before send.
  - Why: fail early and clearly.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Client] Debug logging changed from unconditional Immediate Window output to config-controlled output.
  - Previous behavior: submit/status debug output could always print response data.
  - New behavior: debug output is gated by `SEF_DEBUG_LOG = DA`.
  - Why: production runs should not dump unnecessary response data.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Client] SEF document ID handling for cancel/storno changed.
  - Previous behavior: `CLng(sefDocumentId)` could overflow for large IDs.
  - New behavior: numeric string validation is used without `Long` conversion.
  - Why: avoid overflow and preserve the external ID literal.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Service] SEF transaction EH changed to preserve original errors before rollback.
  - Previous behavior: rollback cleanup could mask original error state.
  - New behavior: error number/description are captured before rollback.
  - Why: diagnostics and logs should show the original failure.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA/SEF Status] Repeated final status refresh no longer attempts illegal same-state transitions.
  - Symptom: refreshing already accepted/rejected invoices could hit state-machine errors.
  - Resolution: same-state refresh updates refresh fields only.
  - Reference update required: Yes

- [Layer: VBA/SEF Client] `GetInvoiceStatus` parser correctness fixed.
  - Symptom: status endpoint could be parsed like submit response.
  - Resolution: `GetInvoiceStatus` must call `ParseStatusResponse`, not `ParseSubmitResponse`.
  - Reference update required: Yes

- [Layer: VBA/SEF Mapper] VAT mismatch risk fixed.
  - Symptom: DTO totals and UBL tax percent could diverge.
  - Resolution: one tax source through `GetDefaultTaxPercent`.
  - Reference update required: Yes

- [Layer: VBA/SEF Persistence] Local `RaiseUpdateError` pattern removed after migration.
  - Symptom: local update error pattern duplicated global guard behavior.
  - Resolution: callers migrated to `RequireUpdateCell`, helper removed.
  - Reference update required: Yes

- [Layer: VBA/SEF Client] HTTP 429 handling no longer relies on parser fallback.
  - Symptom: rate limit handling was incomplete and duplicated.
  - Resolution: `ApplyRateLimitResponse` maps 429 to `RATE_LIMITED` and includes `Retry-After` when available.
  - Reference update required: Yes

### DEPRECATED
- [Layer: VBA/Update Guards] Form-local `MustUpdateCell` helpers as canonical update API.
  - Replacement: `modSchemaGuard.RequireUpdateCell`.
  - Removal target: delete when no callers remain; wrapper allowed only temporarily.
  - Reference update required: Yes

- [Layer: VBA/SEF Persistence] Local `RaiseUpdateError` style update wrappers.
  - Replacement: `RequireUpdateCell`.
  - Removal target: removed from the hardened SEF persistence path.
  - Reference update required: Yes

- [Layer: VBA/SEF Client] `CLng(sefDocumentId)` in JSON request bodies.
  - Replacement: `GetJsonNumericIdLiteral`.
  - Removal target: cancel/storno paths.
  - Reference update required: Yes

### REMOVED
- [Layer: VBA/SEF Persistence] `RaiseUpdateError` removed after all relevant callers migrated to `RequireUpdateCell`.
  - Reason: duplicate fail-fast update pattern.
  - Reference update required: Yes
  - Migration required: No

### KNOWN ISSUE
- [KI-v6.5-01] `modSEFClient` still uses lightweight JSON extraction.
  - Affected layer: VBA / SEF client
  - Impact: simple fields are supported; nested/escaped JSON can still be fragile.
  - Workaround: keep parsing limited to known simple SEF fields.
  - Reference update required: Yes

- [KI-v6.5-02] `ComputePayloadHash` is a lightweight fingerprint, not a cryptographic hash.
  - Affected layer: VBA / SEF mapper
  - Impact: acceptable for current payload identity, not audit-grade proof.
  - Workaround: document clearly; consider SHA-256 later.
  - Reference update required: Yes

- [KI-v6.5-03] SEF debug response truncation marker not fully standardized.
  - Affected layer: VBA / SEF client
  - Impact: low; debug is config-gated, but explicit `[truncated]` helper can improve readability.
  - Workaround: keep `SEF_DEBUG_LOG` disabled in production unless troubleshooting.
  - Reference update required: Yes

- [KI-v6.5-04] SEF test procedures still exist in production modules.
  - Affected layer: VBA / SEF
  - Impact: low if not exposed in UI.
  - Workaround: move to `modSEFTests` or mark dev-only.
  - Reference update required: Yes

- [KI-v6.5-05] `modSEFPersistance` spelling debt remains.
  - Affected layer: VBA / SEF
  - Impact: none while compile references are consistent.
  - Workaround: preserve current name until controlled rename.
  - Reference update required: Yes

- [KI-v6.5-06] Two low-risk `frmOtkupAPP` cleanup items remain.
  - Affected layer: VBA main shell
  - Impact: low; `DisplayAlerts` cleanup and `ResetHover` guard should still be added.
  - Workaround: manual patch later.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.5-01] Replace lightweight JSON extraction with a proper JSON parser/wrapper for SEF responses.
  - Why it matters: improves robustness for escaped/nested error responses.
  - Affected modules: `modSEFClient`.
  - Target state: response parsing does not rely on fragile string extraction.

- [RM-v6.5-02] Optional SHA-256 payload hash for SEF submissions.
  - Why it matters: stronger audit-grade payload identity if required.
  - Affected modules: `modSEFMapper`, `modSEFPersistance`.
  - Target state: cryptographic hash available without breaking existing history.

- [RM-v6.5-03] Move SEF `Test_*` procedures to `modSEFTests`.
  - Why it matters: production modules should not carry ad-hoc dev routines.
  - Affected modules: SEF modules.
  - Target state: tests isolated from production modules.

- [RM-v6.5-04] Adaptive SEF retry scheduler.
  - Why it matters: `RATE_LIMITED` and `Retry-After` are now surfaced, but retry scheduling remains manual.
  - Affected modules: `modSEFClient`, `modSEFService`, `modSEFStatusSync`.
  - Target state: controlled retry/backoff behavior.

- [RM-v6.5-05] Finish low-risk `frmOtkupAPP` cleanup.
  - Why it matters: defensive UI shell hygiene.
  - Affected form: `frmOtkupAPP`.
  - Target state: `DisplayAlerts` cleanup and `ResetHover` guard applied.

### Migration Notes
- No Excel table schema migration is required.
- No Google Sheet schema migration is required.
- VBA project must compile with `modParse`, `modSchemaGuard`, `modComboBinding` and the SEF modules present.
- Confirm `GetInvoiceStatus` calls `ParseStatusResponse`.
- Confirm `modSEFPersistance.RaiseUpdateError` has no remaining callers before deletion.

### Verification / Smoke Tests
- Compile VBA project.
- Build and serialize a valid SEF invoice.
- Try invalid faktura cases: no stavke, missing PIB, zero/invalid amount.
- Send invoice successfully and verify `tblSEFSubmission` and `tblSEFEventLog` rows.
- Refresh `SENT -> ACCEPTED` and `SENT -> REJECTED`.
- Refresh already `ACCEPTED` and already `REJECTED` invoices again.
- Test `SEF_SYNC_ERROR` recovery.
- Test `SEF_TECH_FAILED` retry.
- Test `RecoverStuckSEFSendingInvoice` with and without `SEFDocumentId`.
- Test cancel/storno allowed status, disallowed status and empty comment.

### Documentation Actions
- [x] Reference updated to v6.5 final
- [x] Changelog v6.5 created
- [x] Stammdaten update guard convergence documented
- [x] frmOtkupAPP shell hardening documented
- [x] SEF status sync hardening documented
- [x] SEF mapper hardening documented
- [x] SEF persistence hardening documented
- [x] SEF validator hardening documented
- [x] SEF client hardening documented
- [x] SEF service/recovery hardening documented


## v6.4 — 2026-04-26

### Summary
- VBA desktop pre-launch hardening was extended from `Dokumenta` through `Fakturisanje`, `Otkup`, form lifecycle utilities and `Sledljivost` repair/traceability flows.
- Desktop startup/shutdown now has a clearer lifecycle contract through `Workbook_Open`, `modMain`, startup backup/log verification, controlled shutdown and safe form activation.
- `frmDokumenta` / `modDokumenta` were hardened with atomic multi-class saves, centralized parsing, schema/update guards, hidden-ID combos, no-`MsgBox` business modules and guarded relink/orphan handling.
- `frmFakturisanje` / `modFaktura` were hardened against duplicate invoicing and partial writes, with explicit invoice selection for printing and `PrintFaktura` aligned to the current `tblFakturaStavke` schema.
- `frmOtkup` / `modOtkup` were hardened so Klasa I, optional Klasa II, ambalaža, cash payout and avans allocation run as one atomic business operation through `SaveOtkupMulti_TX`.
- `frmOtkupniBlokovi` / `modSledljivost` were hardened without changing business logic: existing matching keys, return shapes, trace output and PWA/desktop flow were preserved while adding column guards, checked updates and a transaction wrapper for batch auto-link.
- PWA-first traceability was documented: ideally otkup blocks and zbirna/driver context come from PWA; VBA remains the fallback and repair layer for non-PWA users and exceptional cases.
- Shared helper architecture was formalized around `modParse`, `modComboBinding` and `modSchemaGuard`.

### ADDED
- [Layer: VBA/Documents] Atomic multi-wrapper document saves: `SaveOtpremnicaMulti_TX`, `SaveZbirnaMulti_TX`, `SavePrijemnicaMulti_TX`.
  - Impact: dual-class document saves now commit or rollback as one operator action.
  - Why: prevent Klasa I from persisting while Klasa II or related side effects fail.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Otkup] `SaveOtkupMulti_TX` high-level transaction wrapper.
  - Impact: Klasa I, optional Klasa II, ambalaža, cash payout and avans allocation now share one rollback boundary.
  - Why: desktop otkup is one business operation and must not be split across independent `_TX` calls.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Fakturisanje] Explicit `cmbFaktura` invoice selection for printing.
  - Impact: operators print a selected FakturaID instead of the last faktura found for a kupac.
  - Why: printing should be deterministic and operator-controlled.
  - Reference update required: Yes
  - Migration required: Form control addition required in VBA form designer.

- [Layer: VBA/Sledljivost] `AutoLinkOtkupOtpremnica_TX` transaction wrapper.
  - Impact: batch auto-link of `tblOtkup.OtpremnicaID` can rollback if a required link update fails.
  - Why: traceability repair writes must be fail-fast and recoverable.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Parsing] Shared `modParse` helpers.
  - Impact: numeric/date parsing is centralized through `TryParseDouble`, `TryParseLong`, `TryParseDateValue` and numeric normalization.
  - Why: avoid locale-sensitive `Val`, raw `CDbl`, raw `CLng` and duplicated private parser logic.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/UI Binding] Shared `modComboBinding` helpers.
  - Impact: ComboBoxes can display human labels while storing stable hidden IDs, with `GetComboID`, `GetComboDisplay` and `SetComboByID`.
  - Why: display text is not a stable primary key.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Schema/Data Guards] `RequireUpdateCell` in `modSchemaGuard`.
  - Impact: critical updates now raise when `UpdateCell` fails.
  - Why: avoid silent partial relink/status/linking corruption.
  - Reference update required: Yes
  - Migration required: No

- [Layer: Architecture/PWA+VBA] PWA-first / VBA-fallback traceability rule.
  - Impact: PWA is recognized as preferred field source for otkup blocks and driver/zbirna context, while VBA remains the complete fallback/repair system.
  - Why: launch architecture must support both modern PWA operation and non-PWA/manual desktop operators.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA/App Lifecycle] Desktop boot changed to a controlled lifecycle path.
  - Previous behavior: startup responsibilities could drift into `Workbook_Open` or form activation.
  - New behavior: `Workbook_Open` delegates to `modMain.StartApp`; startup backup/log checks and safe failure behavior are part of the lifecycle contract.
  - Why: startup must be recoverable and must not leave Excel invisible on failure.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/App Lifecycle] Shutdown changed to a controlled `ShutdownApp` path.
  - Previous behavior: forms could hide/unload and leave uncertain Excel visibility or lifecycle trace.
  - New behavior: normal exits route through `ShutdownApp`, restore `Application.Visible = True`, unload the shell and write shutdown log marker.
  - Why: operator sessions need predictable close behavior and traceability.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Documents] Dual-class document saves changed from multi-call form orchestration to one business wrapper.
  - Previous behavior: Klasa I and Klasa II could be saved through separate operations.
  - New behavior: Otpremnica/Zbirna/Prijemnica multi-class flows commit or rollback together.
  - Why: a dual-class entry is one operator action.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Fakturisanje] Faktura creation changed to fail-fast, guarded write semantics.
  - Previous behavior: `CreateFaktura_TX` could commit even when the base save returned `""`, and stavka/prijemnica updates were not fully checked.
  - New behavior: empty result raises, `AppendRow` is checked, `RequireUpdateCell` is used, and selected prijemnice are prevalidated.
  - Why: invoice creation must not double-fakturisati or partially mark prijemnice.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Fakturisanje] `PrintFaktura` changed from legacy `KulturaID`-based output to current prijemnica-based stavka output.
  - Previous behavior: print logic referenced a legacy `KulturaID` column in `tblFakturaStavke`.
  - New behavior: print logic reads current faktura stavke columns: broj prijemnice, klasa, količina, cena and vrednost.
  - Why: the invoice data model is prijemnica-based.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Otkup] Desktop otkup save flow changed from separate transaction calls to `SaveOtkupMulti_TX`.
  - Previous behavior: `frmOtkup` called `SaveOtkup_TX`, then `SaveNovac_TX`, then `ApplyAvansToOtkup_TX`.
  - New behavior: `frmOtkup` validates/parses UI input and calls one high-level wrapper.
  - Why: prevent partial persistence across otkup, ambalaža, cash and avans side effects.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Sledljivost] Auto-link and manual-link changed to checked update semantics without changing matching logic.
  - Previous behavior: `UpdateCell` result was not enforced.
  - New behavior: link writes use `RequireUpdateCell`, and batch auto-link can run inside `AutoLinkOtkupOtpremnica_TX`.
  - Why: traceability repairs must not silently fail.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Forms] Form lifecycle cleanup changed to a pragmatic standard.
  - Previous behavior: some activation/close/navigation paths lacked error handling or had debug/test buttons in operator surfaces.
  - New behavior: forms get minimal EH, QueryClose discipline and shared navigation where safe, while proven titlebar/chrome patterns are preserved instead of over-refactored.
  - Why: launch safety should not break working UI chrome behavior.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA/Documents] Dual-class partial-save risk fixed.
  - Symptom: Klasa I could remain saved while Klasa II failed.
  - Root cause: separate saves for one operator action.
  - Resolution: atomic multi-wrapper saves for document flows.
  - Reference update required: Yes

- [Layer: VBA/Documents] Faktura/prijemnica relink partial-update risk fixed.
  - Symptom: invoice-line relink could partially update faktura stavke or prijemnica status.
  - Root cause: unchecked `UpdateCell` calls.
  - Resolution: `RequireUpdateCell` in guarded relink flow.
  - Reference update required: Yes

- [Layer: VBA/Fakturisanje] Duplicate invoice risk reduced/fixed in the create path.
  - Symptom: a selected prijemnica could be included despite being stornirana or already fakturisana.
  - Root cause: UI and module did not both enforce invoiceability.
  - Resolution: form-level block and module-level prevalidation.
  - Reference update required: Yes

- [Layer: VBA/Fakturisanje] Blind last-faktura printing fixed.
  - Symptom: `btnStampaj_Click` printed the last faktura for a kupac rather than an explicitly selected invoice.
  - Root cause: no faktura selection control.
  - Resolution: added `cmbFaktura` with hidden FakturaID and refreshed/selected new faktura after creation.
  - Reference update required: Yes

- [Layer: VBA/Fakturisanje] Legacy `KulturaID` print dependency fixed.
  - Symptom: `PrintFaktura` referenced a no-longer-active `KulturaID` column.
  - Root cause: print logic lagged the prijemnica-based faktura model.
  - Resolution: print output now reads current `tblFakturaStavke` schema.
  - Reference update required: Yes

- [Layer: VBA/Otkup] Desktop otkup transaction split fixed.
  - Symptom: otkup, cash and avans side effects could persist inconsistently.
  - Root cause: form orchestrated several independent `_TX` wrappers.
  - Resolution: `SaveOtkupMulti_TX` wraps the whole business operation.
  - Reference update required: Yes

- [Layer: VBA/Otkup] Locale-sensitive parsing risk reduced.
  - Symptom: raw `Val`, `CDbl`, `CLng`, `CDate` could misread user input.
  - Root cause: form-local parsing style.
  - Resolution: `frmOtkup` moved save-path parsing to `modParse` helpers.
  - Reference update required: Yes

- [Layer: VBA/Sledljivost] Unchecked link writes fixed.
  - Symptom: auto-link/manual-link could appear successful without confirming `OtpremnicaID` update.
  - Root cause: unchecked `UpdateCell`.
  - Resolution: `RequireUpdateCell` and transaction wrapper for batch auto-link.
  - Reference update required: Yes

- [Layer: VBA/App Lifecycle] Form activation and popup lifecycle rough edges reduced.
  - Symptom: selected forms could throw during activation or over-refactor of titlebar removal could alter layout.
  - Root cause: UI lifecycle patterns were not standardized pragmatically.
  - Resolution: minimal EH and rollback to proven chrome-removal patterns where needed.
  - Reference update required: Yes

### DEPRECATED
- [Layer: VBA/Forms] Resolving canonical entity IDs from ComboBox display text.
  - Replacement: hidden-ID ComboBox binding through `modComboBinding`.
  - Removal target: critical flows handled in v6.4; remaining low-risk forms migrate progressively.
  - Reference update required: Yes

- [Layer: VBA/Business Modules] `MsgBox` inside business/data modules.
  - Replacement: modules log/raise/return; forms display operator-facing messages.
  - Removal target: immediate for hardened flows, progressive for remaining modules.
  - Reference update required: Yes

- [Layer: VBA/Write Logic] Unchecked `UpdateCell` in critical relink/status/linking logic.
  - Replacement: `RequireUpdateCell`.
  - Removal target: immediate for document, faktura, otkup and sledljivost critical paths.
  - Reference update required: Yes

- [Layer: VBA/Otkup] Calling `SaveOtkup_TX`, `SaveNovac_TX` and `ApplyAvansToOtkup_TX` separately for one desktop otkup operator action.
  - Replacement: `SaveOtkupMulti_TX`.
  - Removal target: immediate for `frmOtkup` save path.
  - Reference update required: Yes

- [Layer: VBA/Sledljivost] Treating traceability as only a PDF/report surface.
  - Replacement: traceability is also the repair/audit surface for missing `Otkup → Otpremnica` links.
  - Removal target: documentation-level deprecation; existing UI remains with hardened behavior.
  - Reference update required: Yes

### REMOVED
- [Layer: VBA/Fakturisanje] Legacy `KulturaID` dependency from invoice print path.
  - Reason: invoice stavke are prijemnica-based.
  - Impact: `PrintFaktura` aligns with current `tblFakturaStavke` schema.
  - Reference update required: Yes

- [Layer: VBA/Otkup] Known partial-save behavior for dual-class/cash/avans desktop otkup path.
  - Reason: replaced by `SaveOtkupMulti_TX`.
  - Impact: one operator save action has one rollback boundary.
  - Reference update required: Yes

- [Layer: VBA/App Lifecycle] Debug/test button exposure in production menu surfaces should not remain canonical.
  - Reason: production menu must expose operator functions only.
  - Impact: debug hooks should be renamed, hidden, or removed for launch.
  - Reference update required: Yes

- No Excel table schema or Google Sheet schema was removed in v6.4.

### KNOWN ISSUES
- [KI-v6.4-01] Remaining business modules outside the hardened flows may still contain older `MsgBox`, raw parsing or unchecked update patterns.
  - Affected layer: VBA modules
  - Impact: lower-priority flows may not yet match the new launch standard.
  - Workaround: migrate module-by-module using the v6.4 patterns.
  - Should remain in canonical reference: Yes

- [KI-v6.4-02] PWA-first traceability can be improved with suggested desktop repair hints, but this is not part of the hardening pass.
  - Affected layer: PWA/VBA traceability
  - Impact: operators still use existing repair UI rather than richer PWA-suggested matching.
  - Workaround: keep current canonical chain and explicit repair workflow.
  - Should remain in canonical reference: Yes

- [KI-v6.4-03] `GetSaldoByStation` remains a guarded gross helper, not a full cross-domain accounting saldo engine.
  - Affected layer: VBA reporting
  - Impact: Banka/Novac/Isporuka deductions require separate reporting design.
  - Workaround: keep TODO as roadmap; do not expand `modOtkup` save module into accounting reports.
  - Should remain in canonical reference: Yes

- [KI-v6.4-04] Some form chrome/titlebar removal patterns are sensitive to timing and should not be over-refactored.
  - Affected layer: VBA forms
  - Impact: visual artifacts can appear if working patterns are changed aggressively.
  - Workaround: preserve proven working pattern and add only minimal EH around it.
  - Should remain in canonical reference: Yes

### ROADMAP
- [RM-v6.4-01] Extend `modSchemaGuard` / `RequireUpdateCell` adoption to remaining critical modules.
  - Why it matters: finance, SEF, banka, agro and reports benefit from fail-fast data access semantics.
  - Affected modules: `modNovac`, `modFaktura`, `modBankaMapiranje`, `modAgrohemija`, reports.
  - Target state: no unchecked business-critical `UpdateCell` calls.

- [RM-v6.4-02] Finish no-`MsgBox` cleanup across remaining business modules.
  - Why it matters: UX feedback should remain in forms/controllers, not business/data modules.
  - Affected modules: all `mod*.bas` with business logic.
  - Target state: modules log/return/raise; forms display.

- [RM-v6.4-03] PWA-assisted traceability suggestions.
  - Why it matters: PWA will usually know driver/zbirna context and can reduce desktop repair work.
  - Affected modules: PWA otkup flow, MasterSync, `frmOtkupniBlokovi`, `modSledljivost`.
  - Target state: desktop can show PWA-suggested context without bypassing the canonical `OtpremnicaID` bridge.

- [RM-v6.4-04] Dedicated saldo/reporting module for station/kooperant balance.
  - Why it matters: cross-domain saldo should not be bolted onto `modOtkup` core save logic.
  - Affected modules: reports, novac, banka, otkup, dokumenta.
  - Target state: clear accounting/reporting aggregation with Banka/Novac/Isporuka rules.

- [RM-v6.4-05] Complete form lifecycle cleanup pass.
  - Why it matters: remaining forms should have safe Activate/QueryClose/navigation patterns, but proven UI chrome behavior should be preserved.
  - Affected modules/forms: remaining `frm*.frm` files.
  - Target state: minimal EH, no debug/test buttons in production, safe navigation and no visual regressions.

### Migration Notes
- No Excel/VBA table schema migration is required.
- No Google Sheet schema migration is required.
- VBA codebase must include the shared helper modules before compile:
  - `modParse`
  - `modComboBinding`
  - `modSchemaGuard`
- `RequireUpdateCell` is part of `modSchemaGuard`, not local to `modDokumenta` or other domain modules.
- `frmFakturisanje` requires the added `cmbFaktura` control for explicit invoice printing.
- Forms that use hidden-ID ComboBoxes must initialize them before calling `GetComboID(...)`.
- `SaveOtkupMulti_TX` assumes base `SaveNovac(...)` and base `ApplyAvansToOtkup(...)` are callable inside an existing transaction.
- Sledljivost hardening intentionally preserves existing matching logic and output shapes; any PWA-specific matching changes require explicit design approval.

### Verification / Smoke Tests
- Compile VBA project.
- Dokumenta: save one-class and two-class Otpremnica/Zbirna/Prijemnica; force/verify rollback behavior where possible.
- Fakturisanje: create faktura from valid prijemnice; verify selected prijemnice are marked; try duplicate fakturisanje and verify block; print selected faktura.
- Otkup: save one-class and two-class otkup, with/without cash, with/without ambalaža, with/without parcela; verify rollback behavior and no separate form-level finance transaction calls.
- Sledljivost: run auto-link, manual link, load trace by zbirna and export PDF; verify `Application.ScreenUpdating` returns to True after print path.
- App lifecycle: open workbook, verify startup backup/log path, main shell, close/shutdown path and Excel visibility.

### Documentation Actions
- [x] Reference updated to v6.4 final
- [x] Changelog v6.4 created
- [x] Dokumenta hardening documented
- [x] Fakturisanje hardening documented
- [x] Otkup hardening documented
- [x] Sledljivost hardening documented
- [x] PWA-first / VBA-fallback traceability rule documented
- [x] Shared parse/combo/schema guard architecture documented
- [x] Form lifecycle cleanup documented
- [x] Known issues and roadmap reviewed

## v6.3 — 2026-04-25

### Summary
- GAS `doPost` endpoint authorization was hardened with explicit role/entity ownership checks across write and quota-sensitive actions.
- `saveParcelPolygon` was moved behind token validation and made Management-only.
- Sync endpoints now enforce role scope and entity ownership before writing to OTK/VOZ/AGRO/TRETMAN/OPREMA/TROSKOVI sheets.
- Sheet/Drive write actions are now required to run through `withLock(...)`.
- `saveFiskalniMapiranje` and `createArtikal` are Management-only shared/master-data actions.
- `updateKamionStatus` now allows Vozac only for own status and Management for any driver; Vozac caller scope is derived from `tokenData.entityID`.
- `parseFiskalniImage` and `parseFiskalni` are no longer treated as public utilities; they require authenticated Kooperant or Management scope.
- Backend launch smoke testing confirmed expected 401/403 behavior, role ownership checks and core role sync flows.

### ADDED
- [Layer: GAS/AuthZ] Local authorization helper contract in `Code.gs`: `isManagement`, `requireRole`, `requireEntity` and `forbiddenResponse`.
  - Impact: endpoint role checks are now expressed through shared helpers instead of ad-hoc repeated checks.
  - Why: reduce authorization drift across the action router.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Security] Explicit endpoint authorization matrix for GAS `doPost` actions.
  - Impact: each action now has documented public/auth/role/ownership/lock expectations.
  - Why: launch readiness requires every write and quota-sensitive action to have a known authorization boundary.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Concurrency] `withLock(...)` coverage requirement for actions that write to Google Sheets or Google Drive.
  - Impact: sync writes, dispatch writes, PDF/Drive writes, fiscal saves/mapping, parcel polygon writes and master-artikal creation are protected against common concurrent write races.
  - Why: GAS/Sheets write operations can be called concurrently from field devices and management UI.
  - Reference update required: Yes
  - Migration required: No

- [Layer: QA/Launch] Backend smoke-test checklist for GAS launch gate.
  - Impact: verified `ping`, login, 401 without token, 403 for wrong role/scope, Management flow, Otkupac sync, Vozac syncZbirna, Kooperant syncAgromere, `saveFiskalniMapiranje` denial for Kooperant and `logClientError`.
  - Why: validate that endpoint-authz changes did not break core launch flows.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: GAS/Router] `saveParcelPolygon` changed from pre-auth write surface to authenticated Management-only write action.
  - Previous behavior: `saveParcelPolygon` could execute before the normal token validation gate.
  - New behavior: action executes only after `validateToken(...)` / `getTokenData(...)`, requires Management role and runs under `withLock(...)`.
  - Why: parcel polygon writes are GIS/master-data mutations and must not be public.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Sync] Sync endpoints changed from token-only protection to role/entity-scoped write authorization.
  - Previous behavior: after token validation, sync actions did not consistently verify that the caller owned the requested `otkupacID`, `kooperantID` or `vozacID`.
  - New behavior: `sync` requires Otkupac/Management and Otkupac callers must match `otkupacID`; `syncAgromere`, `syncTretman`, `syncOprema` and `syncTrosak` require Kooperant/Management and Kooperant callers must match `kooperantID`; `syncZbirna` requires Vozac/Management and Vozac callers must match `vozacID`.
  - Why: prevent authenticated users from writing into another role/entity scope.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Dispatch] `updateKamionStatus` changed from broadly token-protected write to role-scoped dispatch status update.
  - Previous behavior: an authenticated caller could invoke the action without explicit Vozac/Management role enforcement.
  - New behavior: Management may update any driver; Vozac may update only own status, with backend forcing `data.vozacID = tokenData.entityID`.
  - Why: client-supplied driver identity must not be trusted for Vozac-owned status writes.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Fiskalni] Fiscal parser actions changed from authenticated-but-ungated utility calls to role-gated actions.
  - Previous behavior: `parseFiskalniImage` and `parseFiskalni` executed after auth but without explicit Kooperant/Management role checks.
  - New behavior: parser actions require Kooperant or Management; Kooperant-scoped payloads must match the authenticated entity when a `kooperantID` is supplied.
  - Why: parsing consumes backend/API quota and processes private kooperant fiscal data.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Fiskalni] `saveFiskalniMapiranje` changed to Management-only.
  - Previous behavior: action executed for any authenticated token.
  - New behavior: action requires Management and runs under `withLock(...)`.
  - Why: fiscal name mapping is shared matching metadata and must not be writable by ordinary kooperant sessions.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Artikli] `createArtikal` authorization was formalized as Management-only and locked.
  - Previous behavior: Management check existed, but the launch contract did not explicitly document it as master-data write authority.
  - New behavior: Management-only master artikal creation is part of the endpoint authorization matrix and runs under `withLock(...)`.
  - Why: master `Artikli` remains operator/management controlled; private kooperant fiscal items must not enter master catalog.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/PDF/Drive] PDF generation/upload actions now have explicit Otkupac/Management role and ownership expectations.
  - Previous behavior: `saveOtkupniListPdf` and `uploadPdf` were token-protected but not explicitly role/entity scoped in the documented action matrix.
  - New behavior: Otkupac may act only on own `otkupacID`; Management may act as override; write path runs through `withLock(...)`.
  - Why: PDF and Drive writes are business-document side effects.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: GAS/Security] Public write exposure for `saveParcelPolygon` was closed.
  - Symptom: parcel polygon save was still reachable before the normal auth gate.
  - Root cause: legacy placement in `doPost` before `validateToken(...)`.
  - Resolution: move action below auth, require Management role and wrap write in `withLock(...)`.
  - Reference update required: Yes

- [Layer: GAS/Security] Authenticated cross-entity write risk was reduced for sync actions.
  - Symptom: a valid token could be paired with another entity ID in payload unless each endpoint checked ownership.
  - Root cause: endpoint router validated token before action handling but did not consistently bind action scope to `tokenData.entityID`.
  - Resolution: role/entity ownership checks were added for all sync write actions.
  - Reference update required: Yes

- [Layer: GAS/Security] Kooperant access to shared fiscal mapping was blocked.
  - Symptom: `saveFiskalniMapiranje` could be called by non-Management authenticated sessions.
  - Root cause: missing role gate on a shared mapping write.
  - Resolution: Management-only gate and lock wrapper were added.
  - Reference update required: Yes

- [Layer: GAS/Dispatch] Vozac status updates no longer trust client-supplied `vozacID`.
  - Symptom: a Vozac caller could attempt to submit another driver's ID in the payload.
  - Root cause: endpoint did not force identity from token data.
  - Resolution: Vozac branch overwrites `data.vozacID` with `tokenData.entityID`.
  - Reference update required: Yes

### DEPRECATED
- [Layer: GAS/Router] Any business write action before the normal token validation gate.
  - Replacement: `login` and `logClientError` remain the only intentional pre-auth POST exceptions; public geo/meteo reads remain an acknowledged read-only auth gap.
  - Removal target: immediate in v6.3 endpoint-authz contract.
  - Reference update required: Yes

- [Layer: GAS/AuthZ] Ad-hoc per-endpoint role checks without shared helper vocabulary.
  - Replacement: local helper contract through `isManagement`, `requireRole`, `requireEntity` and `forbiddenResponse`.
  - Removal target: ongoing cleanup; v6.3 establishes the active pattern.
  - Reference update required: Yes

### REMOVED
- [Layer: GAS/Security] Public write behavior for `saveParcelPolygon`.
  - Reason: parcel polygon save mutates GIS/master data and cannot remain public.
  - Impact: clients must call it with a valid Management token.
  - Reference update required: Yes

- No business schema, Google Sheet column layout or Excel/VBA table was removed in v6.3.

### KNOWN ISSUES
- [KI-v6.3-01] Token age validation should also be enforced for token payloads found directly in `CacheService`.
  - Affected layer: GAS auth/session
  - Impact: cache-hit tokens should be checked against the same absolute expiry discipline as fallback tokens.
  - Workaround: keep current bounded cache TTL and daily purge; harden `validateToken(...)` in the next auth pass.
  - Should remain in canonical reference: Yes

- [KI-v6.3-02] `logError(...)` still needs sensitive-field redaction before writing details.
  - Affected layer: GAS/PWA observability/security
  - Impact: tokens, PINs, base64 image/PDF payloads or signatures could be logged if passed through details.
  - Workaround: keep logged details minimal from callers; add `sanitizeLogDetails(...)` in next observability pass.
  - Should remain in canonical reference: Yes

- [KI-v6.3-03] `logClientError` remains intentionally pre-auth and still needs throttle/dedupe review.
  - Affected layer: GAS/PWA observability/security
  - Impact: repeated client loops could create noisy ErrorLog rows or consume quotas.
  - Workaround: keep payload limited and truncated; add cache-based dedupe/throttle after launch.
  - Should remain in canonical reference: Yes

- [KI-v6.3-04] Public geo/meteo read bridge remains an acknowledged read-only auth gap.
  - Affected layer: GAS/PWA/GIS/Meteo security
  - Impact: parcel geo/meteo read actions remain callable before token validation.
  - Workaround: treat as accepted current behavior until frontend request model and role-scoped map reads are fully gated.
  - Should remain in canonical reference: Yes

### ROADMAP
- [RM-v6.3-01] Harden `validateToken(...)` so cache-hit payloads are age-validated against the absolute expiry rule.
  - Why it matters: persistent and cached token paths should follow one expiry model.
  - Affected modules: GAS auth helpers
  - Target state: cache hit, fallback restore and token-data access all share the same validation function.

- [RM-v6.3-02] Add `sanitizeLogDetails(...)` before `logError(...)` persistence.
  - Why it matters: operational logs must not store bearer tokens, PINs, large base64 payloads or signatures.
  - Affected modules: GAS observability, PWA error reporting
  - Target state: redacted, bounded and safe ErrorLog details.

- [RM-v6.3-03] Add dedupe/throttle protection for `logClientError`.
  - Why it matters: a client-side loop should not flood ErrorLog or consume quota.
  - Affected modules: GAS `doPost`, PWA error bridge
  - Target state: repeated equivalent client errors are rate-limited while preserving critical visibility.

- [RM-v6.3-04] Consider optional `logout` / `revokeToken` endpoint.
  - Why it matters: users and support flows may need explicit session invalidation before natural expiry.
  - Affected modules: GAS auth, PWA session UI
  - Target state: token deleted from both CacheService and PropertiesService.

- [RM-v6.3-05] Consider Management-only `getErrorLog` endpoint for operational review.
  - Why it matters: support should be able to inspect recent backend/client errors without manually opening Google Sheets.
  - Affected modules: GAS observability, Management PWA
  - Target state: recent bounded ErrorLog read surface for Management only.

### Migration Notes
- No Excel/VBA table migration is required.
- No Google business sheet schema migration is required.
- Existing clients must deploy the updated GAS `Code.gs` so endpoint authorization changes are active.
- Any code calling `saveParcelPolygon` must now use a valid Management token.
- Any code calling `saveFiskalniMapiranje` must now use a valid Management token.
- Otkupac, Kooperant and Vozac write payloads must supply entity IDs that match the authenticated session, except where Management override is intentionally used.
- Deployments should still verify `setupTokenPurgeTrigger()` and `rebuildSheetRegistry()` after backend updates.
- Smoke-test rows created during launch validation should be removed from operational sheets or marked as expected test data.

### Documentation Actions
- [x] Canonical reference updated to v6.3
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint authorization matrix added
- [x] Write authority reviewed
- [x] Known issues reviewed
- [x] Roadmap status reviewed
- [x] Deprecated elements reviewed
- [x] Migration notes reviewed



## v6.2 — 2026-04-25

### Summary
- GAS backend observability was formalized through `logError(...)`, the `ErrorLog` workbook and the `logClientError` PWA bridge.
- Session/token handling was hardened from cache-only storage to cache + `PropertiesService` fallback with a 48h hard-expiry rule.
- Backend maintenance now includes token-property purge and ErrorLog retention cleanup through the same scheduled path.
- GAS workbook topology was updated to treat `ErrorLog` as an active shared operational workbook and `SheetRegistry` as the registry-assisted lookup path.
- Remote error logging moved from roadmap-only status to partially implemented / verification-required status.

### ADDED
- [Layer: GAS/Observability] Central `logError(source, action, message, details, entityID)` helper.
  - Impact: GAS runtime errors can now be persisted in a shared operational log instead of only being returned to clients or hidden in Apps Script runtime output.
  - Why: field/PWA and backend failures need traceability across devices and sessions.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Sheets] `ErrorLog` workbook in `MASTER_FOLDER_ID`.
  - Impact: operational errors are appended to a lazily created workbook with columns `Timestamp | Source | Action | Message | Details | EntityID | Severity`.
  - Why: keep error observability inside the existing per-client Google Drive / Sheets deployment model.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/GAS] `logClientError` action in `doPost`.
  - Impact: PWA clients can send runtime/client errors to GAS even before the normal authenticated action router.
  - Why: client-side field failures must remain observable even when the current token is missing, expired or only usable for best-effort entity attribution.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Auth] Token fallback storage through `PropertiesService.getScriptProperties()`.
  - Impact: session tokens are no longer lost immediately on script-cache eviction; valid fallback tokens can be restored into cache.
  - Why: `CacheService` is a fast but volatile storage layer and can cause avoidable session failures.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Maintenance] `purgeExpiredTokens()` and `setupTokenPurgeTrigger()`.
  - Impact: expired or malformed `TOKEN_*` script properties can be cleaned daily, and the same maintenance path also triggers ErrorLog retention cleanup.
  - Why: persistent token fallback requires bounded retention and a cleanup contract.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Maintenance] `purgeOldErrorLogs()` retention helper.
  - Impact: `ErrorLog` rows older than 30 days are removed best-effort.
  - Why: remote logging must not grow without retention control in client-owned Google Sheets.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Sheets] `SheetRegistry` lookup contract documented as an active lookup accelerator.
  - Impact: selected aggregate reads can use name-to-spreadsheet-id registry entries instead of relying only on repeated Drive folder scans.
  - Why: reduce lookup cost and make role-scoped workbook discovery more deterministic while preserving folder-scan fallback.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: GAS/Auth] Session token contract changed from cache-only to cache + persistent fallback.
  - Previous behavior: successful login stored `TOKEN_<token>` only in `CacheService.getScriptCache()` for 24h.
  - New behavior: successful login stores the token payload in cache and mirrors it into `PropertiesService`; cache misses may be recovered from script properties until the 48h hard-expiry window.
  - Why: avoid fragile session loss caused by cache volatility while still keeping a bounded expiry model.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Auth] Token validation now includes age validation for fallback tokens.
  - Previous behavior: token validity was equivalent to presence in cache.
  - New behavior: fallback token payloads older than 48h are rejected and deleted; valid fallback tokens are restored into cache.
  - Why: persistent fallback must not become unbounded long-lived authentication state.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Router] `doPost` action surface changed to include pre-auth `logClientError`.
  - Previous behavior: pre-auth actions were limited to public reads, `login`, and the existing `saveParcelPolygon` auth-gap path.
  - New behavior: `logClientError` is accepted before the normal token gate, with optional token validation only for entity attribution.
  - Why: logging must be available in failure states, including session expiry and client runtime degradation.
  - Reference update required: Yes
  - Migration required: No

- [Layer: Architecture/Roadmap] Remote error logging status changed from open roadmap item to partially implemented / verify.
  - Previous behavior: remote logging was tracked as a missing capability.
  - New behavior: GAS-side logging and endpoint support exist; remaining work is deployment verification and full frontend role wiring.
  - Why: refactored `Code.gs` implements the backend side of the capability.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Sheets] Shared workbook topology changed to include `ErrorLog`.
  - Previous behavior: active shared workbooks were documented around `Stammdaten`, `Kartice`, `MgmtReports` and `LoginLog`.
  - New behavior: `ErrorLog` is part of the active shared operational workbook family under `MASTER_FOLDER_ID`.
  - Why: logging is now a first-class operational backend surface.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: GAS/Observability] Backend catch paths now have a reusable non-blocking logging target.
  - Symptom: GAS failures could be returned as error payloads but were not consistently persisted for later diagnosis.
  - Root cause: no canonical GAS-side error log existed in the active architecture.
  - Resolution: `logError(...)` writes best-effort rows to `ErrorLog`, and failures inside the logger are swallowed.
  - Reference update required: Yes

- [Layer: GAS/Auth] Token cache volatility is mitigated.
  - Symptom: sessions could fail if script-cache entries disappeared before the user expected the session to expire.
  - Root cause: `CacheService` was the only token store.
  - Resolution: token payloads are mirrored into `PropertiesService` and restored into cache when still inside the 48h hard-expiry window.
  - Reference update required: Yes

- [Layer: GAS/Maintenance] Persistent token fallback now has a cleanup path.
  - Symptom: adding persistent token storage without cleanup would create unbounded stale auth state.
  - Root cause: persistent fallback requires lifecycle management beyond cache TTL.
  - Resolution: `purgeExpiredTokens()` removes expired/malformed `TOKEN_*` properties and can be scheduled daily.
  - Reference update required: Yes

### DEPRECATED
- [Layer: GAS/Auth] Treating `CacheService` as the only session-token store.
  - Replacement: cache-first token storage with `PropertiesService` fallback and 48h hard-expiry.
  - Removal target: immediate in v6.2 backend contract.
  - Reference update required: Yes

- [Layer: Roadmap/Observability] Treating remote error logging as purely unimplemented.
  - Replacement: mark backend logging as implemented and keep verification/frontend rollout as the remaining work.
  - Removal target: immediate documentation cleanup.
  - Reference update required: Yes

### REMOVED
- No business domain behavior was removed in v6.2.
- No PWA store, table or Google Sheet business schema migration was removed or replaced.

### KNOWN ISSUES
- [KI-v6.2-01] `logClientError` is intentionally pre-auth and must stay minimal.
  - Affected layer: GAS/PWA observability/security
  - Impact: endpoint can receive unauthenticated payloads and must not become a general write surface.
  - Workaround: keep payload limited to error metadata, truncate message/details, avoid accepting arbitrary operational commands.
  - Should remain in canonical reference: Yes

- [KI-v6.2-02] Remote logging backend exists, but full frontend role-wide wiring still requires verification.
  - Affected layer: PWA/GAS observability
  - Impact: not every frontend failure path is guaranteed to call `logClientError` until role-level smoke tests confirm it.
  - Workaround: verify Otkupac, Kooperant, Vozac and Management runtime catch paths during launch testing.
  - Should remain in canonical reference: Yes

- [KI-v6.2-03] `PropertiesService` token fallback improves resilience but increases cleanup/security responsibility.
  - Affected layer: GAS auth/session
  - Impact: stale token properties must be purged reliably to avoid long-lived auth debris.
  - Workaround: run `setupTokenPurgeTrigger()` during deployment and verify the daily `purgeExpiredTokens` trigger exists.
  - Should remain in canonical reference: Yes

### ROADMAP
- [RM-v6.2-01] Verify PWA role-wide `logClientError` adoption.
  - Why it matters: backend logging is useful only if critical frontend catch paths actually report failures.
  - Affected modules: PWA app shell, role modules, API/sync wrappers, GAS `doPost`
  - Target state: all role-critical runtime errors can be observed through `ErrorLog`.

- [RM-v6.2-02] Add deployment checklist item for token/error-log maintenance trigger.
  - Why it matters: `PropertiesService` fallback depends on bounded cleanup.
  - Affected modules: GAS deployment procedure
  - Target state: `setupTokenPurgeTrigger()` is run and verified for each client deployment.

- [RM-v6.2-03] Review whether `logClientError` needs throttling or severity normalization after field testing.
  - Why it matters: remote logs should remain operationally useful and not become noisy or abusable.
  - Affected modules: GAS observability, PWA error reporting
  - Target state: clear volume control and severity taxonomy if real usage requires it.

### Migration Notes
- No Excel/VBA table migration is required.
- No existing Google business sheet schema migration is required.
- New deployments should allow the backend to lazily create `ErrorLog` in `MASTER_FOLDER_ID`.
- Existing deployments should run or manually verify `setupTokenPurgeTrigger()` so expired `TOKEN_*` properties and old ErrorLog rows are cleaned.
- `SheetRegistry` is additive and fallback-safe; missing registry data should not block folder-scan lookup behavior.

### Documentation Actions
- [x] Canonical reference updated to v6.2
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Roadmap status reviewed
- [x] Deprecated elements reviewed
- [x] Migration notes reviewed


## v6.1 — OtkupApp v2.2.1 desktop pre-launch hardening

### App lifecycle wiring
- `Workbook_Open` was reduced to a thin EH-wrapped delegator to `modMain.StartApp`; previously it ran a mini-lifecycle that bypassed `StartApp` entirely
- `StartApp` is now actually invoked at boot, which means `ValidateAllTables`, `BackupFileOnStart`, `PurgeOldBackups`, `PurgeOldJournals`, `PurgeOldLogs`, `LogAppStart`, `RecoverAllStuckSEFSendingInvoices` and `CheckJournalForRecovery` finally run on every cold start as the reference always intended
- `ImportBankaInbox_TX` was removed from the `Workbook_Open` boot path; bank inbox import is now invoked only through `frmBankaImport` UI affordances, eliminating the silent auto-import side effect at boot
- `Workbook_Open` EH path now restores `Application.Visible = True` before showing the user-facing error message, so a startup failure can no longer leave the operator locked out of an invisible Excel
- `Workbook_BeforeClose` now routes through `ShutdownApp`, so the daily log has a paired open/close lifecycle marker regardless of whether the operator closes the form or the workbook
- `frmSplash` invocation moved inside `StartApp`, keeping the splash → `frmOtkupAPP` handoff intact while making `StartApp` the single canonical boot contract
- **Reference impact:** new section §6.1.29a App Startup Orchestration added; §6.1.28 and §6.1.30 updated to reflect the new shutdown symmetry and dynamic version label

### Versioning
- `APP_VERSION` constant in `modConfig.bas` advanced from `"2.1"` to `"2.2.1"`
- `frmSplash.lblVersion.Caption` is now derived as `"v" & APP_VERSION` instead of the hardcoded `"v2.1.0"` literal
- the splash form is now a single source of truth consumer of `APP_VERSION`; future bumps require touching only `modConfig`
- **Reference impact:** §6.1.30 Desktop Splash Form Capabilities updated

### Theming and UI cleanup
- removed `modReportModernUI.bas` (224 lines of dead code) — it was never invoked from any flow but declared `Public Const UI_BG`/`UI_PANEL`/etc. with values that conflicted with `modModernUI.Private UI_BG`, polluting the global constant scope across the project
- removed `modUI.bas` (empty placeholder, 1 line)
- removed `frmModernUIIzvestaji.frm` + `.frx` (391 lines of dead form code) — `frmIzvestaj` is the canonical reporting shell per the reference and was the only one ever opened from `frmOtkupAPP`
- canonical theming surface is now exactly `modTheme` (palette, colors, button styling) plus `modModernUI` (form-level helpers); no third paralel theming module exists anymore
- **Reference impact:** none — these modules were never documented as canonical in any reference version

### Bank reconciliation form deduplication
- the 8 `_Preview` helper functions in `frmBankaImport.frm` were removed (~290 lines): `GetBankaImportRowByID_Preview`, `TryResolveKupacBIM_Preview`, `TryResolveKooperantBIM_Preview`, `TryResolveOMBIM_Preview`, `TryResolveFakturaForKupac_Preview`, `GetOtkupCandidatesForKooperantBlock_Preview`, `NormalizeLoosePreview`, `NzBankaPreview`, plus the small `GetKooperantNazivPreview` helper
- the corresponding 8 functions in `modBankaMapiranje` were promoted from `Private` to `Public` so the form can call them directly: `GetBankaImportRowByID`, `TryResolveKupacBIM`, `TryResolveKooperantBIM`, `TryResolveOMBIM`, `TryResolveFakturaForKupac`, `GetOtkupCandidatesForKooperantBlock`, `NormalizeLooseBIM`, `NzBIM`
- new `Public Function GetKooperantNaziv(kooperantID)` was added to `modBankaMapiranje` to replace the form-local `GetKooperantNazivPreview`
- `GetBankaImportRowByID` was extended from a 9-column to a 10-column return shape (added `COL_BIM_POZIV_NA_BROJ` as column 10) so the form preview can show the call-reference line; existing module callers use only columns 1-9, so the change is backward-compatible
- `frmBankaImport.frm` shrank from 794 lines to about 385 lines (≈52% smaller)
- preview-parity with backend is now a structural property of the codebase rather than two independently maintained copies
- **Reference impact:** §6.1.27 frmBankaImport updated to list the actual public helpers used by preview

### ReportSaldoOM avans edge bug
- removed the premature `Exit Function` after `GetOtkupByStation` in `modIzvestaj.ReportSaldoOM` so an empty otkup result no longer aborts the report before Novac aggregation runs
- removed the inner `If dict.count = 0 Then ... Exit Function` guard which had the same effect 20 lines later
- the Novac aggregation loop was already correctly structured to dynamically add kooperants who appear only via avans, but the early-exits prevented it from ever running in the affected case
- as of this version, a station whose only money activity in the period is kooperant avans (no Otkup) now correctly surfaces those kooperants with negative saldo in the report
- **Reference impact:** known-issue **KI-024** (`ReportSaldoOM avans edge bug`) status changed from **Open** to **Closed**

### SEF mapper EH discipline
- error message for missing `SELLER_NAME` and `SELLER_PIB` config values in `modSEFMapper.BuildSEFInvoiceDto` was corrected from `"missing in tblConfig"` to `"missing in " & TBL_SEF_CONFIG & "."`; previously the message pointed operators to the wrong table (a separate `tblConfig` exists for non-SEF settings, while SEF config actually lives in `tblSEFConfig`)
- four functions in `modSEFMapper` (`BuildSEFInvoiceDto`, `SerializeSEFRequest`, `SerializeUBLInvoice`, `GetInvoiceDeliveryDate`) were missing an `Exit Function` between their success-path return and the `EH:` label, causing successful calls to fall through into the error handler with `Err.Number = 0`
- the missing `Exit Function` statements were added; the EH labels now also include `LogErr "ProcName"` as the first line, restoring the project-wide convention already followed by every other module
- **Reference impact:** none — this is a regression fix that restores the documented EH/`LogErr` invariant from §6.1.11

### Form navigation cleanup
- `frmMarza`, `frmOtkup`, and `frmStammdaten` previously had their "Povratak" return action hardcoded to `frmMain.Show`, but `frmOtkupAPP` is the canonical home shell per reference §6.1.28
- all three return paths were changed to `frmOtkupAPP.Show`, so the operator now consistently lands on the actual home shell instead of a legacy `frmMain` instance
- `frmMain` is implicitly deprecated by this change but not deleted; deletion can follow once a separate audit confirms no other call sites
- **Reference impact:** none — this aligns code with reference §6.1.28 which already names `frmOtkupAPP` as the canonical home

### Form EH coverage parity
- `frmOtkupAPP.UserForm_Activate` EH block now calls `LogErr "frmOtkupAPP.UserForm_Activate"` as its first line, matching the project-wide pattern already used by every `*_TX` wrapper in `modBankaMapiranje`, `modSEFService` and `modSEFStatusSync`
- the previously unmonitored EH path was the only remaining EH block in a form-level handler that was not feeding the daily log file
- **Reference impact:** none — implicit invariant from §6.1.11 now uniformly applied

### Removed and migrated artifacts
- **REMOVED** `src-vba/modReportModernUI.bas` (224 lines, dead code, polluted global UI_* constant scope)
- **REMOVED** `src-vba/modUI.bas` (1 line, empty placeholder)
- **REMOVED** `src-vba/frmModernUIIzvestaji.frm` + `.frx` (391 lines, dead form, never invoked)
- **REMOVED** `src-vba/clsSEFValidationResult.cls` (header-only class, 0 members, 0 references)
- **CHANGED** `ThisWorkbook.Workbook_Open`, `ThisWorkbook.Workbook_BeforeClose`
- **CHANGED** `modMain.StartApp` (frmSplash.Show migrated inside)
- **CHANGED** `modConfig.APP_VERSION = "2.2.1"`
- **CHANGED** `frmSplash.UserForm_Initialize` (lblVersion now derived)
- **CHANGED** `modSEFMapper.bas` (EH discipline + correct error message table reference)
- **CHANGED** `modBankaMapiranje.bas` (8 Private → Public, GetBankaImportRowByID extended to 10 cols, new GetKooperantNaziv helper)
- **CHANGED** `frmBankaImport.frm` (8 _Preview helpers removed, 31 call-site renames)
- **CHANGED** `modIzvestaj.ReportSaldoOM` (early-exit guards removed)
- **CHANGED** `frmOtkupAPP.frm` (LogErr added to EH block)
- **CHANGED** `frmMarza.frm`, `frmOtkup.frm`, `frmStammdaten.frm` (frmMain.Show → frmOtkupAPP.Show)

### Net code change
- approximately **−750 lines deleted, +50 lines added or modified** across 14 files in `src-vba/`
- four files removed entirely; no new files added
- public API contracts of `modBankaMapiranje`, `modIzvestaj`, `modSEFMapper`, `clsTransaction`, `modDataAccess`, `modNovac` remain unchanged at the call-site level

### Closed known issues
- **KI-024** ReportSaldoOM avans edge bug — Closed
- **KI-028** (new) `Workbook_Open` ran `ImportBankaInbox_TX` automatically and could leave Excel invisible on failure — Closed
- **KI-029** (new) `frmSplash` hardcoded version label `v2.1.0` while `APP_VERSION` advanced — Closed
- **KI-030** (new) `modReportModernUI` global `UI_BG` collision with `modModernUI` private — Closed
- **KI-031** (new) `frmBankaImport` carried 8 `_Preview` duplicates of `modBankaMapiranje` private helpers — Closed

### Open and deferred
- **P2-2** seven parallel `Nz*` implementations across modules and forms — deferred to post-launch refactor; no migration risk in current state
- **P2-3** Phase 3 stub modules (`modMeteo`, `modML`, `modKvalitet`) — kept as documented placeholders per AR-007/AR-009 roadmap items
- **P2-5** `MapBankaImportAsKooperant` still does post-save `FindRows` + `UpdateCell` for `OtkupID` instead of using the optional 14th `SaveNovac` argument — functionally correct, performance-only refactor for post-launch
- **KI-025**, **KI-026**, **KI-027** remain Open per their reference status

### Migration notes
- no schema changes (`tblOtkup`, `tblNovac`, `tblBankaImport`, `tblFakture`, `tblSEFSubmission`, `tblSEFEventLog`, `tblConfig`, `tblSEFConfig` all unchanged)
- no Google Sheets contract changes
- no GAS endpoint changes
- no PWA changes
- existing workbooks open and operate with the new build without any manual migration step
- operators will see the splash version label as `v2.2.1` instead of `v2.1.0` on next boot, which is the only externally visible change

## v6.0 — frontend hardening / CSP cleanup / IndexedDB recovery

### Frontend security and runtime stability
- completed repo-wide cleanup of remaining runtime inline event handlers across feature modules
- static `index.html` and runtime-rendered feature HTML are now aligned with `script-src 'self'`
- action handling has been standardized through delegated `data-action` / `data-route` patterns and local module-level listeners where appropriate
- management, kooperant, otkupac and shared modal/action flows were updated to avoid inline `onclick` / `onchange` / `oninput` dependencies
- frontend runtime contract now explicitly treats delegated event handling as canonical behavior for shell and dynamic UI surfaces

### Navigation and shell cleanup
- legacy bottom-nav helper behavior was reduced further and aligned with the canonical `role-nav.js` ownership model
- obsolete `onclick`-dependent tab button lookup pattern was retired in favor of `data-route="tab"` / `data-tab` based lookup
- management routing remains owned by `showMgmtRoot(...)` and the V2 management shell only
- `tabs.js` remains non-management router and no longer acts as implicit management owner

### IndexedDB hardening
- `db.js` was upgraded from a minimal open/create-store helper to a recovery-first IndexedDB layer
- added blocked/open timeout handling
- added `versionchange` connection lifecycle handling
- added controlled reset/reopen recovery path
- added explicit `resetIndexedDb()` helper
- added safer schema provisioning helpers instead of ad-hoc store bootstrap only
- added store/index existence guards in IndexedDB access helpers
- local database layer no longer depends on manual browser-storage clearing as the primary recovery path

### Deployment and readiness impact
- CSP / inline-handler cleanup is effectively closed for current shell and reviewed feature surfaces
- runtime-generated UI is now compatible with the stricter script CSP posture already adopted in the app shell
- IndexedDB layer is materially more resilient to upgrade/open/store drift issues
- sync unification and smoke-test automation remain outside this version and are tracked as subsequent work

## Changelog — v5.9

### PWA shell / CSP / assets
- `index.html` contract ažuriran na self-hosted vendor model:
  - `./vendor/html5-qrcode.min.js`
  - `./vendor/jspdf.umd.min.js`
  - `./vendor/leaflet.css`
  - `./vendor/leaflet.js`
  - `./vendor/chart.umd.min.js`
- CSP ažuriran:
  - `script-src 'self'`
  - uklonjena potreba za script-side `'unsafe-inline'`
  - dokumentovano da `style-src` i dalje koristi `'unsafe-inline'` zbog postojećih inline style atributa
- uklonjene stare reference da shell i dalje zavisi od CDN runtime učitavanja
- service worker contract ažuriran na aktivni cache generation `AgriX-v10`
- dokumentovano da SW sada kešira self-hosted vendor assete
- offline napomena ažurirana:
  - Leaflet asset gap je zatvoren
  - live tile slojevi su i dalje network-dependent

### Frontend navigation / shell ownership
- `tabs.js` više nije dokumentovan kao management intercept layer
- canonical nav ownership sada glasi:
  - non-management: `showTab(...)`
  - management: `showMgmtRoot(...)` + `role-nav.js`
- uklonjen otvoreni arhitektonski dug koji je tvrdio da management routing i dalje zavisi od `tabs.js` intercept-a
- management shell dokumentovan kao čistiji V2-only root model

### Frontend API client
- dodat novi canonical `api.js` contract section
- dokumentovano:
  - POST-first request model
  - `apiBuildUrl()` više ne lepi token u query string
  - `apiRequestSafe(...)` koristi timeout (`AbortController`)
  - proverava `resp.ok`
  - normalizuje parse/network/timeout/auth failure
  - novi normalized result shape:
    `{ ok, status, data, error, code, isTimeout, isNetworkError, isAuthError }`
- `apiFetchSafe(...)` i `apiPostSafe(...)` uneseni kao canonical safe helpers
- uklonjen open issue da token još ide kroz URL i da API client nema launch-grade hardening

### Auth / session
- reference je usklađen sa canonical `entityID` session modelom
- `otkupacID` je dokumentovan samo kao compatibility alias za starije Otkupac tokove
- login shell je usklađen sa novim auth/runtime contract-om

### GAS router / backend contract
- `doPost` dokumentovan kao primarni frontend contract
- dodat public POST-read bridge za:
  - `getParcelGeo`
  - `getParcelMeteo`
  - `getParcelMeteoLatest`
  - `getAllMeteoLatest`
- dokumentovano da public geo/meteo exposure i dalje ostaje auth-gap, ali sada kroz POST-first bridge
- u current action surface dodato:
  - `syncTrosak`
  - `saveFiskalniMapiranje`
  - `createArtikal`
- `doGet` ostavljen kao compatibility/read surface + health check
- failure contract dopunjen public-read bridge napomenom

### Kooperant / parcele
- `parcele.js` contract dopunjen:
  - popup/list/detail akcije su sada CSP-safe
  - koriste delegated `data-action` model umesto inline handlera
- time je GIS/detail flow usklađen sa strožim script CSP-om

### Known Issues cleanup
- uklonjeni zastareli/open issue unosi za:
  - management `tabs.js` intercept
  - missing `agroState` guard
  - token u query string-u
  - API client hardening
  - inline handlers / script-side CSP slabost
  - CDN runtime dependency
- `KI-007` preformulisan:
  - više nije “SW asset incompleteness”
  - sada je “full offline GIS still limited by live tile dependency”
- dodati novi aktivni issue unosi:
  - `KI-034` IndexedDB migration / recovery plan is still minimal
  - `KI-035` sync result shape still not fully unified across role modules
  - `KI-036` `style-src` still needs `'unsafe-inline'`

### Roadmap cleanup
- uklonjene zatvorene roadmap stavke za:
  - query-token cleanup
  - API hardening
  - inline-handler/script-CSP hardening
  - self-host vendor libs
- roadmap dodatno usklađen sa preostalim realnim radom:
  - `AR-021` IndexedDB migration / recovery strategy
  - `AR-022` unified sync result and retry contract across roles
  - `AR-023` cross-role smoke test matrix for online/offline transitions
  - `AR-024` style-CSP hardening / inline-style reduction

### Deprecated / transitional updates
- dodat deprecated zapis za runtime CDN library loading
- dodat deprecated zapis za management `tabs.js` root intercept
- replacement jasno naveden:
  - local `vendor/` assets
  - `role-nav.js` + `showMgmtRoot(...)`

## v5.8 — 2026-04-23

### Summary
- Frontend session/auth contract was normalized around `entityID` as the canonical cross-role identity key.
- Background sync was expanded from Otkupac-only behavior to role-aware orchestration across Otkupac, Kooperant and Vozac.
- Kooperant offline model was simplified to one canonical treatment store (`tretmani`) and legacy dual-store behavior was removed.
- Service worker/runtime hardening and runtime-state cleanup reduced launch risk in the PWA shell.

### ADDED
- [Layer: PWA/Core] Canonical runtime-state ownership through the shared state/runtime path instead of separate private app-shell runtime assumptions.
  - Impact: bootstrap, refresh and sync guards now have a cleaner single-state direction.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Sync] Role-aware background sync orchestration in app bootstrap.
  - Impact: Kooperant and Vozac now participate in the same scheduled sync discipline already used by Otkupac.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: PWA/Auth] Session identity contract changed from legacy generic `otkupacID` reliance to canonical `entityID`.
  - Previous behavior: frontend session/bootstrap still treated `otkupacID` as a generic persisted entity key even outside Otkupac-only flows.
  - New behavior: `entityID` is the canonical persisted identity; `otkupacID` remains only as a compatibility alias where older Otkupac-specific payloads still expect it.
  - Why: cleaner cross-role auth/session semantics.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Config] Runtime identity surface changed so `CONFIG.ENTITY_ID` is canonical and `CONFIG.OTKUPAC_ID` is now derived compatibility state.
  - Previous behavior: config still exposed older Otkupac-centered identity semantics.
  - New behavior: config is aligned to one role-neutral entity identity surface.
  - Why: reduce naming debt and session drift.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Sync] Background sync changed from Otkupac-only interval behavior to role-aware orchestration.
  - Previous behavior: `syncQueueSafe()` / scheduled sync were effectively Otkupac-centric.
  - New behavior: app shell delegates background sync by role to existing role-specific sync functions.
  - Why: operational consistency across active offline-first roles.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Kooperant] Offline persistence contract changed from dual agro store model to canonical `tretmani` store ownership.
  - Previous behavior: kooperant runtime still carried both `agromere`/`AGRO_STORE` legacy store semantics and `tretmani`.
  - New behavior: `tretmani` is the one canonical kooperant treatment store and sync path.
  - Why: remove dual-store drift before launch testing.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: PWA/Offline] Service worker asset contract was hardened.
  - Previous behavior: obsolete management asset references, incomplete third-party asset coverage and fail-all install risk remained.
  - New behavior: service worker cache list matches active shell dependencies better, obsolete assets are removed and install is more resilient.
  - Why: stronger offline/deploy stability.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Security] Entry-shell CSP was expanded to include the active Chart.js CDN origin.
  - Previous behavior: Chart.js origin and CSP policy could diverge.
  - New behavior: CSP now explicitly allows the concrete CDN origin used by the current shell.
  - Why: avoid runtime blocking of active chart dependency.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/State] App bootstrap/runtime logic changed from private `appRuntime` ownership to canonical shared runtime-state usage.
  - Previous behavior: `app.js` kept a second private runtime object while other modules also used `window.appRuntime`.
  - New behavior: runtime flags are normalized around the shared runtime path and compatibility aliasing instead of parallel ownership.
  - Why: remove split runtime-state drift.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: PWA/Auth] Cross-role session validation no longer depends on legacy generic `otkupacID` semantics.
  - Symptom: bootstrap/session validity still reflected older Otkupac-centered identity assumptions.
  - Root cause: auth/config/app bootstrap were not yet aligned to canonical `entityID`.
  - Resolution: session, config and bootstrap were normalized around `entityID`.
  - Reference update required: Yes
- [Layer: PWA/Kooperant] Kooperant offline/sync model no longer maintains parallel treatment-store semantics.
  - Symptom: offline agronomy persistence carried unnecessary dual-store drift.
  - Root cause: legacy `agromere` compatibility path remained active beside `tretmani`.
  - Resolution: `agromere` store/runtime path was removed and `tretmani` remained as the single source of offline truth.
  - Reference update required: Yes
- [Layer: PWA/ServiceWorker] Service worker registration/install path was stabilized after asset-list cleanup and syntax correction.
  - Symptom: service worker registration could fail due to stale asset references and fragile asset-list handling.
  - Root cause: active shell/runtime assets and worker asset list drifted.
  - Resolution: worker asset list was cleaned up, obsolete management shell reference removed, active third-party assets included and install strategy hardened.
  - Reference update required: Yes
- [Layer: PWA/State] Runtime synchronization bug caused by split runtime ownership was reduced.
  - Symptom: different modules could effectively reason about different runtime state surfaces.
  - Root cause: private app-shell runtime state and global runtime usage coexisted.
  - Resolution: runtime state was consolidated around one shared path.
  - Reference update required: Yes

### DEPRECATED
- [Layer: PWA/Auth] Using `otkupacID` as a generic cross-role session/entity key.
  - Replacement: canonical `entityID`.
  - Removal target: ongoing cleanup of remaining compatibility references.
  - Reference update required: Yes
- [Layer: PWA/Kooperant] Legacy `agromere`/`AGRO_STORE` offline treatment-store semantics.
  - Replacement: `tretmani` as the single kooperant treatment store.
  - Removal target: completed in active runtime.
  - Reference update required: Yes

### REMOVED
- [Layer: PWA/Kooperant] Legacy `agromere` IndexedDB store contract from active kooperant offline model.
  - Reason: no legacy data was being preserved for this launch path; testing starts from the canonical store model.
  - Impact: kooperant treatment persistence and sync now revolve around `tretmani` only.
  - Reference update required: Yes
- [Layer: PWA/Config] `CONFIG.AGRO_STORE` from the active canonical runtime contract.
  - Reason: dual-store kooperant persistence is no longer part of the active architecture.
  - Impact: cleaner offline model and less feature drift.
  - Reference update required: Yes

### KNOWN ISSUES
- [KI-v3.3-01] `tabs.js` / `role-nav.js` / `showMgmtRoot` navigation ownership is still partially transitional.
  - Affected layer: PWA navigation/runtime
  - Impact: management navigation still spans multiple routing surfaces
  - Workaround: keep management root state and role-nav state aligned
  - Should remain in canonical reference: Yes
- [KI-v3.3-02] API client still lacks launch-grade timeout / `resp.ok` / normalized error-shape handling.
  - Affected layer: PWA API/runtime
  - Impact: transport failures still degrade into weak `null`-style semantics
  - Workaround: rely on wrapper usage and console diagnostics until client hardening pass
  - Should remain in canonical reference: Yes
- [KI-v3.3-03] Token transport for GET-style API reads is still not fully hardened.
  - Affected layer: PWA auth/API
  - Impact: query-string token transport remains a security/operational debt
  - Workaround: treat as priority hardening item before production launch
  - Should remain in canonical reference: Yes
- [KI-v3.3-04] Entry shell still depends on inline handlers and therefore retains `'unsafe-inline'` CSP posture.
  - Affected layer: PWA UI/security
  - Impact: CSP is improved but not yet strict production-grade
  - Workaround: keep current shell stable until event-binding hardening pass
  - Should remain in canonical reference: Yes

### ROADMAP
- [RM-v3.3-01] Finish navigation ownership cleanup.
  - Why it matters: management routing still spans `tabs.js`, `role-nav.js` and management shell state
  - Affected modules: `src/js/ui/tabs.js`, `src/js/ui/role-nav.js`, management shell
  - Target state: one explicit owner for management root navigation
- [RM-v3.3-02] Harden API client contract for production.
  - Why it matters: launch-ready transport should not depend on `null`-style failure handling
  - Affected modules: `src/js/services/api.js`, auth/runtime shell
  - Target state: timeout + `resp.ok` + normalized result shape
- [RM-v3.3-03] Remove inline handlers and tighten CSP posture.
  - Why it matters: current shell still requires `'unsafe-inline'`
  - Affected modules: `index.html`, UI binding layer
  - Target state: stricter production-ready CSP
- [RM-v3.3-04] Self-host critical frontend third-party libraries.
  - Why it matters: runtime/offline stability should not depend on CDN availability
  - Affected modules: `index.html`, `sw.js`, asset pipeline
  - Target state: self-hosted QR/PDF/map/chart dependencies

### Migration Notes
- If any browser still has an older IndexedDB schema, database version upgrade must recreate stores under the new canonical kooperant offline model.
- No backend data migration is required for the `entityID` session/auth cleanup; this is a frontend contract normalization.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

## v5.7 — 2026-04-23

### Summary
- Management frontend runtime was consolidated onto one canonical shell.
- Legacy management compatibility wrapper and runtime DOM transplant were removed.
- Management dashboard chart/data pipeline was stabilized.
- Local dead code and shell-level observability were improved.

### ADDED
- [Layer: PWA/Management] Shared V2-shell helpers for canonical Management rendering and state access.
  - Impact: overview/dashboard/sub-tab flows now rely on one local helper set instead of repeated shell-local patterns.
  - Reference update required: Yes
  - Migration required: No
  - [Layer: PWA/Management] Explicit Management shell error logging in catch paths.
    - Impact: boot/render failures are now visible in console instead of being silently swallowed.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: PWA/Management] Management shell ownership changed from parallel legacy+V2 loading to V2-only ownership.
  - Previous behavior: `mgmt-shell-v2.js` and legacy `mgmt-shell.js` were both loaded, with bootstrap still depending on compatibility glue.
  - New behavior: Management runtime is owned only by `mgmt-shell-v2.js`; required legacy boot helpers were migrated into the V2 shell.
  - Why: reduce routing drift, remove parallel shell behavior and lower regression surface.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Management] Management DOM ownership changed from runtime transplant to canonical in-place V2 mount ownership.
  - Previous behavior: `mgmtMountLegacyBlocks()` physically moved old Management DOM blocks into V2 containers during init.
  - New behavior: Management content now lives directly inside canonical V2 mount zones in `index.html`.
  - Why: eliminate hybrid shell behavior and reduce DOM/event-binding fragility.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Management] Overview and Dashboard data reads were normalized around shared field/date access.
  - Previous behavior: overview and dashboard used different field/date parsing paths, creating drift in KPI/chart behavior.
  - New behavior: shared helper-based reads now normalize otkup date/quantity access and period copy.
  - Why: consistency and chart correctness.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: PWA/Management] Dashboard chart no longer collapsed valid data into zero-valued series.
  - Symptom: chart rendered points as `0` even when Management data existed.
  - Root cause: field/date parsing drift between Management dashboard and other reads.
  - Resolution: canonicalized Management otkup date/quantity access and aligned dashboard/overview parsing.
  - Reference update required: Yes
- [Layer: PWA/Bootstrap] Duplicate `db.js` include removed from the active entry shell.
  - Symptom: duplicated service load in frontend boot chain.
  - Root cause: `index.html` still carried an extra `db.js` include from transitional bootstrap state.
  - Resolution: kept one canonical DB service include.
  - Reference update required: Yes
- [Layer: PWA/Management] Silent Management-shell failures were made observable.
  - Symptom: shell boot/render failures could be swallowed without console signal.
  - Root cause: empty `catch (e) {}` blocks in Management runtime paths.
  - Resolution: explicit error logging added to Management shell catch paths.
  - Reference update required: Yes

### DEPRECATED
- [Layer: PWA/Management] Legacy Management compatibility wrapper model.
  - Replacement: canonical `mgmt-shell-v2.js` ownership with direct V2 mount zones.
  - Removal target: completed in active runtime.
  - Reference update required: Yes
- [Layer: PWA/Management] Legacy `tab-mgmt` / `mgmtSubBar` wrapper.
  - Replacement: explicit `tab-mgmt-*` V2 root surfaces.
  - Removal target: completed in active runtime.
  - Reference update required: Yes

### REMOVED
- [Layer: PWA/Management] Legacy `mgmt-shell.js` from active runtime.
  - Reason: Management bootstrap no longer depends on parallel legacy shell behavior.
  - Impact: Management now runs through one canonical shell only.
  - Reference update required: Yes
- [Layer: PWA/Management] `mgmtMountLegacyBlocks()` runtime DOM transplant.
  - Reason: Management blocks are now placed directly in canonical V2 mount containers.
  - Impact: shell init no longer reparents legacy DOM on startup.
  - Reference update required: Yes
- [Layer: PWA/Management] Dead shell-local bottom-nav helpers:
  - `showMgmtBottomRoot`
  - `updateMgmtBottomNavActive`
  - `updateMgmtBottomNavVisibility`
  - Reason: no active repo/runtime references remained after Management V2 shell consolidation.
  - Impact: smaller shell surface and less dead code.
  - Reference update required: Yes

### KNOWN ISSUES
- [KI-v3.2-01] `tabs.js` management root intercept remains transitional.
  - Affected layer: PWA management navigation
  - Impact: management routing still depends on coordination between `tabs.js`, `role-nav.js` and `mgmt-shell-v2.js`
  - Workaround: keep V2 root state and role-nav state aligned
  - Should remain in canonical reference: Yes
- [KI-v3.2-02] Session key naming debt (`otkupacID` vs `entityID`) remains open.
  - Affected layer: PWA auth/bootstrap
  - Impact: role-neutral session semantics are still partially obscured by legacy naming
  - Workaround: keep `entityID` authoritative in future cleanup and preserve compatibility meanwhile
  - Should remain in canonical reference: Yes

### ROADMAP
- [RM-v3.2-01] Finish management navigation simplification.
  - Why it matters: root routing still spans `tabs.js`, `role-nav.js` and `mgmt-shell-v2.js`
  - Affected modules: PWA management navigation
  - Target state: one explicit management root-routing contract
- [RM-v3.2-02] Finish PWA session-key normalization.
  - Why it matters: `entityID` should become the one canonical cross-role session/entity key
  - Affected modules: auth/bootstrap/config/runtime state
  - Target state: remove generic reliance on legacy `otkupacID` naming

### Migration Notes
- No data migration required.
- Frontend runtime migration completed: Management entry shell now expects only `mgmt-shell-v2.js` and direct V2 mount ownership.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed


## v5.6 — 2026-04-22

### Summary
- Desktop MasterSync architecture was extended with active Zbirna import from VOZ Google Sheets.
- OTK and VOZ Google Sheet contracts were aligned to GAS-first column order.
- `tblZbirna` canonical active schema was corrected to match the real Excel table, including `ClientRecordID` and `SyncSource`.
- `BrojZbirne` was separated from `ServerRecordID` and moved to explicit business-number generation in VBA.
- VOZ sheet contract was extended with dedicated `BrojZbirne` column and plain-text formatting safeguards for slash-based values such as `TipAmbalaze` and `BrojZbirne`.
- VBA Google OAuth scope was widened because GAS-created VOZ files were not visible under limited Drive access.

### ADDED
- [Layer: VBA/Desktop Sync] Active `ImportZbirneFromPWA` / `ImportZbirneFromPWA_TX` flow for VOZ sheet ingestion into desktop master.
  - Impact: Zbirna is now part of the canonical desktop master-sync flow rather than remaining an open roadmap item.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Sheets] Dedicated `BrojZbirne` column added to canonical `ZBIRNA_COLUMNS`.
  - Impact: VOZ Google Sheets now store sync/server ID and business document number separately.
  - Reference update required: Yes
  - Migration required: Yes

- [Layer: GAS/Sheets] `ensurePlainTextColumn(...)` introduced for VOZ sheet columns whose business values may contain `/`.
  - Impact: prevents Google Sheets from coercing `TipAmbalaze` values like `12/1` and `BrojZbirne` values into dates.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Vozac] `brojZbirne` added to the local/client zbirna model and render path in `zbirna.js`.
  - Impact: driver UI can display business `BrojZbirne` instead of overloading `serverRecordID`.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: Sheets/VBA/GAS] OTK Google Sheet contract was aligned to GAS-first `COLUMNS` order.
  - Previous behavior: desktop OTK import and sheet bootstrap followed legacy VBA-oriented header order.
  - New behavior: VBA `GS_*`, sheet bootstrap and writeback now follow GAS `COLUMNS`.
  - Why: remove schema drift between GAS-created sheets and desktop import expectations.
  - Reference update required: Yes
  - Migration required: No

- [Layer: Sheets/VBA/GAS] VOZ/Zbirna Google Sheet contract was aligned to GAS-first `ZBIRNA_COLUMNS` order.
  - Previous behavior: desktop VOZ import expected an older header/position layout inconsistent with GAS-created sheets.
  - New behavior: VBA `VS_*` and VOZ writeback now follow GAS `ZBIRNA_COLUMNS`.
  - Why: remove incorrect row parsing and writeback mismatches in VOZ sheets.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Documents] `BrojZbirne` generation moved away from implicit `ServerRecordID` reuse.
  - Previous behavior: early Zbirna import versions overloaded `ServerRecordID` with business-number semantics.
  - New behavior: `BrojZbirne` is generated in desktop VBA from business rule based on `VozacID` + `Datum` + daily sequence.
  - Why: separate technical sync identity from business-visible document numbering.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Schema Docs] Active `tblZbirna` schema was corrected to the real 16-column desktop table.
  - Previous behavior: reference/documentation still reflected a 14-column schema.
  - New behavior: active schema includes `ClientRecordID` and `SyncSource`, matching the real Excel table and current import row mapping.
  - Why: remove documentation drift and ensure dedupe/source semantics are documented.
  - Reference update required: Yes
  - Migration required: No

- [Layer: OAuth/Drive] VBA Google scope widened from limited Drive-file visibility to broader Drive access.
  - Previous behavior: VBA could miss GAS-created `VOZ-*` spreadsheets despite correct folder and naming.
  - New behavior: Drive listing can see GAS-created VOZ files after broader scope + re-authentication.
  - Why: `drive.file` was too narrow for the mixed VBA/GAS file-creation model.
  - Reference update required: Yes
  - Migration required: Yes

### FIXED
- [Layer: VBA/GAS Sheets] VOZ files were not discoverable by desktop sync when created by GAS.
  - Symptom: `FindVOZSheets()` returned 0 files although VOZ sheets existed in the configured folder.
  - Root cause: limited VBA OAuth scope could not reliably enumerate GAS-created files.
  - Resolution: widen Drive scope and re-authorize.
  - Reference update required: Yes

- [Layer: VBA/GAS] OTK and VOZ sheet imports were vulnerable to wrong column reads on GAS-created sheets.
  - Symptom: desktop import could read/write wrong Google Sheet columns because VBA constants assumed legacy order.
  - Root cause: mismatch between legacy VBA positions and GAS-created header order.
  - Resolution: align `GS_*` and `VS_*` to GAS `COLUMNS` / `ZBIRNA_COLUMNS`, and align writeback columns.
  - Reference update required: Yes

- [Layer: VBA/Documents] Early Zbirna import row mapping did not match real `tblZbirna` field order.
  - Symptom: imported Zbirna rows could shift values into wrong columns.
  - Root cause: `rowData` order did not match active `tblZbirna` schema / `SaveZbirna` expectations.
  - Resolution: row mapping corrected to `ZbirnaID, Datum, VozacID, BrojZbirne, ... , ClientRecordID, SyncSource`.
  - Reference update required: Yes

- [Layer: Sheets/Data formatting] `TipAmbalaze` values like `12/1` were coerced into date format in Google Sheets.
  - Symptom: VOZ sheet displayed/held date-like values instead of packaging code.
  - Root cause: slash-based values entered into unformatted Google Sheet columns.
  - Resolution: plain-text formatting rule added through `ensurePlainTextColumn(...)`.
  - Reference update required: Yes

### DEPRECATED
- [Layer: VBA/GAS Sync] Treating `ServerRecordID` as a carrier for business `BrojZbirne`.
  - Replacement: dedicated `BrojZbirne` business field in VOZ sheet + local PWA model + desktop-generated business numbering.
  - Removal target: immediate in active architecture.
  - Reference update required: Yes

- [Layer: Sheets/VBA] Legacy VBA-first OTK/VOZ Google Sheet header order.
  - Replacement: GAS-first `COLUMNS` / `ZBIRNA_COLUMNS`.
  - Removal target: immediate for active sync architecture.
  - Reference update required: Yes

### REMOVED
- [Layer: Roadmap] `MasterSync zbirna import` removed from active roadmap.
  - Reason: desktop Zbirna import is now implemented and part of the active system.
  - Impact: roadmap and known-issue framing must no longer treat Zbirna import as missing.
  - Reference update required: Yes

### KNOWN ISSUES
- [KI-v5.6-01] `_TX` rollback still covers only local Excel table snapshots, not Google Sheets writeback.
  - Affected layer: VBA desktop sync / Google Sheets integration
  - Impact: a desktop rollback can leave Google-side SyncStatus already updated
  - Workaround: operational awareness; future refactor should move external writeback outside transactional illusion
  - Should remain in canonical reference: Yes

### ROADMAP
- [RM-v5.6-01] Consider moving Google Sheets writeback out of desktop `_TX` illusion for OTK/Zbirna flows.
  - Why it matters: current rollback model is local-only while Google writeback is external and irreversible
  - Affected modules: VBA MasterSync, Google Sheets sync
  - Target state: clearer split between local TX and remote writeback

### Migration Notes
- Existing VOZ sheets require the additional `BrojZbirne` column in the canonical GAS header.
- Existing VOZ sheet columns containing slash-based business values should be formatted as plain text.
- VBA OAuth re-authentication is required after Drive scope widening.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

## v3.1 — 2026-04-19

### Summary
- AgriX branding and product framing introduced.
- Architecture reference switched from feature-heavy delta style to explicit system overview with roles, state, boot, sync and known issues.
- v3.1 surfaced several active client/runtime defects as named known issues.

### ADDED
- [Layer: Docs] Full system overview connecting PWA, GAS, Sheets and VBA desktop.
  - Impact: makes the stack easier to onboard and review.
  - Reference update required: Yes
  - Migration required: No
- [Layer: Docs/PWA] Explicit role architecture, state management, boot sequence and service worker documentation.
  - Impact: formalizes client architecture instead of only listing features.
  - Reference update required: Yes
  - Migration required: No
- [Layer: UI/Brand] AgriX role-based branding with role-specific logos and header branding.
  - Impact: product identity unified across roles.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: Docs] Reference narrative moved from “feature dump + roadmap” toward “system architecture”.
  - Previous behavior: references emphasized feature sections and cumulative roadmap.
  - New behavior: system overview, PWA file structure, role architecture, state/sync and known issues are first-class.
  - Why: better onboarding and architecture review.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA] Known runtime issues made explicit in docs.
  - Previous behavior: defects were implicit or buried in roadmap.
  - New behavior: issues like `role-nav.js`, `tabs.js`, duplicate includes and absolute paths are named.
  - Why: operational transparency.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- No new documented production fixes were explicitly declared as completed in the v3.1 file itself; v3.1 focuses on consolidation and surfacing open defects.

### DEPRECATED
- [Layer: Docs] Informal architecture storytelling without explicit state/sync ownership.
  - Replacement: canonical architecture sections for roles, state, sync and known issues.
  - Removal target: immediate editorial standard.
  - Reference update required: Yes

### REMOVED
- No explicit removal was documented for v3.1 beyond editorial repositioning.

### KNOWN ISSUES
- [KI-v3.1-01] `role-nav.js` uses `cfg.showMode` while defining `cfg.type`.
  - Affected layer: PWA management navigation
  - Impact: broken management bottom-nav routing in some flows
  - Workaround: patch role-nav and keep manual fallback routes
  - Should remain in canonical reference: Yes
- [KI-v3.1-02] `tabs.js` lacks `agroState` guard for non-Kooperant roles.
  - Affected layer: PWA tabs
  - Impact: crash risk on role mismatch
  - Workaround: add guard
  - Should remain in canonical reference: Yes
- [KI-v3.1-03] Leaflet CSS/JS missing from service-worker assets.
  - Affected layer: offline GIS
  - Impact: map unavailable offline
  - Workaround: online usage / asset completion
  - Should remain in canonical reference: Yes

### ROADMAP
- [RM-v3.1-01] Remote error logging.
  - Why it matters: field failures remain hard to observe remotely
  - Affected modules: PWA, GAS, ops
  - Target state: structured remote error log
- [RM-v3.1-02] BankaImport re-entry into active roadmap.
  - Why it matters: finance workflow still incomplete in modern platform framing
  - Affected modules: VBA finance/reporting
  - Target state: active canonical finance tooling
- [RM-v3.1-03] Supabase migration exploration.
  - Why it matters: GAS boilerplate and scale concerns
  - Affected modules: whole stack
  - Target state: future platform migration decision

### Migration Notes
- No data migration documented.
- Documentation migration implied: future references should follow canonical snapshot style.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v3.0 — 2026-04-10

### Summary
- Major platform release.
- PWA monolith was refactored into a modular client architecture.
- Knjiga Polja and Fiskalni Scanner were added.
- Meteo moved to batch-oriented architecture with retries and prefetch.

### ADDED
- [Layer: PWA] Modular file structure with ~30 JS files separated into utils, services, ui and role-specific features.
  - Impact: maintainability and clearer ownership improved.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA/Sheets/GAS] Knjiga Polja (bilans, proizvodnja, troškovi, lager).
  - Impact: farm-management layer became first-class.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: PWA/GAS/Sheets] Fiskalni Scanner with photo → QR → SUF parse → auto-match → save.
  - Impact: private kooperant purchase capture introduced.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: Sheets] `FISKALNI-KOOP`, `TROSKOVI-KOOP`, `FiskalniMapiranje`, `BrojParcele` in Kartice.
  - Impact: new data domains and parsing paths.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: Meteo] Batch fetch, retry logic, meteo prefetch in stammdaten.
  - Impact: lower request volume and faster parcel UI.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: PWA] Architecture changed from single-file/monolithic style to modular structure.
  - Previous behavior: large monolithic JS file.
  - New behavior: dedicated modules under `src/js`.
  - Why: maintainability and scalability.
  - Reference update required: Yes
  - Migration required: No
- [Layer: Stammdaten] `normalizeStammdaten()` extended with `meteoLatest`, `kartice` and unknown-field spread.
  - Previous behavior: smaller, more rigid bootstrap payload.
  - New behavior: richer bootstrap data and forward-compatible normalization.
  - Why: instant views and safer payload evolution.
  - Reference update required: Yes
  - Migration required: No
- [Layer: Service Worker] Asset caching moved to graceful per-asset catch instead of fail-all `addAll`.
  - Previous behavior: install failure could break cache build.
  - New behavior: partial cache resilience.
  - Why: robustness.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: Meteo/UI] Parceles without valid coordinates no longer create `0,0` markers.
  - Symptom: bad map markers.
  - Root cause: parcels with missing lat/lng rendered as world-origin point.
  - Resolution: suppress marker when coordinates are missing.
  - Reference update required: Yes
- [Layer: UI] Loader always hides via `finally`.
  - Symptom: spinner could remain stuck after errors.
  - Root cause: missing guaranteed cleanup.
  - Resolution: `finally` path standardized.
  - Reference update required: Yes

### DEPRECATED
- [Layer: PWA] Monolithic JS architecture.
  - Replacement: modular PWA.
  - Removal target: complete; only compatibility wrappers remain.
  - Reference update required: Yes

### REMOVED
- [Layer: PWA] Implicit reliance on per-parcela meteo calls for normal parcel viewing.
  - Reason: replaced by stammdaten prefetch and batch meteo.
  - Impact: parcel UI faster and API usage lower.
  - Reference update required: Yes

### KNOWN ISSUES
- [KI-v3.0-01] Endpoint registration/deploy verification remained P0.
- [KI-v3.0-02] DB version confirmation for `troskovi` remained open.
- [KI-v3.0-03] SW asset completeness and 13 user-flow E2E tests remained open.

### ROADMAP
- [RM-v3.0-01] Remote error logging
- [RM-v3.0-02] Token expiry handling
- [RM-v3.0-03] Offline banner
- [RM-v3.0-04] AutoCreateOtpremniceFromPWA integration into MasterSync
- [RM-v3.0-05] PHI control and LOT numbering
- [RM-v3.0-06] Notifications and photo documentation

### Migration Notes
- `Kartice` export now expects `BrojParcele`.
- Kooperant private fiscal/cost data introduces new dedicated sheets.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v2.8 — 2026-03-29

### Summary
- Dispečer replaced / rebuilt the earlier War Room idea.
- Shared-state localStorage usage was explicitly banned.
- Driver capacity (`KapacitetKG`) became part of active schema and dispatch logic.

### ADDED
- [Layer: Management/PWA] Dispečer 3-column planning flow with plans saved to `DispecerPlan`.
  - Impact: logistics planning became an explicit shared-state subsystem.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: Sheets] `DispecerPlan` and active `KamionStatus` server-side state.
  - Impact: dispatch moved from local heuristics to shared sheet state.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: Master data] `KapacitetKG` on `tblVozaci`.
  - Impact: dispatch can reason about truck capacity.
  - Reference update required: Yes
  - Migration required: Yes

### CHANGED
- [Layer: Architecture] Dispatch was redefined as planning-only.
  - Previous behavior: War Room concept still implied stronger operational writes.
  - New behavior: Dispečer may plan only; Otkupac remains sole authority for field-side `VozacID`.
  - Why: protect write authority and avoid state drift.
  - Reference update required: Yes
  - Migration required: No
- [Layer: PWA] Shared state moved out of localStorage.
  - Previous behavior: localStorage used for status/capacity-like shared state.
  - New behavior: server/sheet state for `KamionStatus` and capacities.
  - Why: phantom-entry bug and cross-session inconsistency.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: Dispatch] Phantom shared-state bug addressed by localStorage ban.
  - Symptom: stale/phantom logistics values across sessions/devices.
  - Root cause: localStorage used as shared source.
  - Resolution: server-backed status + explicit ban.
  - Reference update required: Yes

### DEPRECATED
- [Layer: Naming/UI] War Room naming (`wr*`) deprecated.
  - Replacement: Dispečer / `dp*`
  - Removal target: progressive code cleanup
  - Reference update required: Yes

### REMOVED
- [Layer: GAS] `assignVozacToOtkupi` removed from `Code.gs`.
  - Reason: dispatch should not directly write field assignment.
  - Impact: field-side authority boundary enforced.
  - Reference update required: Yes

### KNOWN ISSUES
- AutoSave after TX still open.
- MasterSync zbirna import still open.
- Domain/deploy and several Phase 2/3 dispatch improvements remained open.

### ROADMAP
- Dispečer Phase 2: drag-and-drop UX
- Dispečer Phase 3: planning suggestions / greedy algorithm
- AutoCreateOtpremniceFromPWA
- MasterSync zbirna import
- Notifications and PHI control

### Migration Notes
- `tblVozaci` schema now includes `KapacitetKG`.
- Dispatch screens and identifiers should migrate from `wr*` naming to `dp*`.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v2.7 — 2026-03-28

### Summary
- Parcel GIS and Meteo pipeline became active system domains.
- Operator tooling for geo maintenance and parcel map flows matured.

### ADDED
- [Layer: GIS] Parcel point/polygon workflow with geo utilities and polygon editor.
  - Impact: parcels gained explicit spatial representation.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: Meteo] Open-Meteo-based pipeline with risk and spray logic, scheduled fetch and current/history sheets.
  - Impact: agronomy and parcel insights became data-driven.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: VBA/PWA] `modClipboard`, `modGeoUtils`, frmStammdaten geo controls and parcel map UI.
  - Impact: operator and kooperant geo flows connected.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: Schema] `tblParcele` expanded to include geo status/source, coordinates, polygon JSON, meteo flags and risk status.
  - Previous behavior: parcel metadata without full geo/meteo support.
  - New behavior: parcel record carries geospatial and meteo state.
  - Why: map and agronomy features.
  - Reference update required: Yes
  - Migration required: Yes

### FIXED
- [Layer: Locale/backup/error pattern] v2.7 clarified locale and backup/error invariants and tightened `BackupFileOnStart`.
  - Symptom: broad error suppression and decimal conversions risked corruption.
  - Root cause: inconsistent helper usage.
  - Resolution: stronger invariant wording and code shifts.
  - Reference update required: Yes

### DEPRECATED
- [Layer: PWA/GAS] placing `fmtStanica()` or PWA JS behavior in `Code.gs`.
  - Replacement: keep display helpers in PWA JS only.
  - Removal target: immediate.
  - Reference update required: Yes

### REMOVED
- No major removals documented; v2.7 is primarily additive.

### KNOWN ISSUES
- saveParcelPolygon auth gap
- Google OAuth testing token limit
- OTK legacy `VozaciID`
- MeteoHistory long-term archival
- KapacitetKg still missing at this stage (resolved in v2.8)

### ROADMAP
- War Room / Operator dashboard
- AutoCreateOtpremniceFromPWA integration
- MasterSync zbirna import
- notifications, photo documentation, PHI, LOT, GlobalGAP audit

### Migration Notes
- Parcel export/import and map layers must understand new geo fields.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v2.6 — 2026-03-28

### Summary
- PWA document flow became operational for OTK and VOZ sheets.
- AutoCreateOtpremniceFromPWA and PWA sync discipline were introduced.
- Error/backup patterns were tightened.

### ADDED
- [Layer: PWA/Sheets] OTK and VOZ sheet operational flow.
  - Impact: field and driver roles became active contributors to document chain.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: VBA] `AutoCreateOtpremniceFromPWA`.
  - Impact: imported OTK rows could auto-generate transport docs.
  - Reference update required: Yes
  - Migration required: No
- [Layer: VBA/GAS/PWA] stronger sync and display helpers around station formatting and backup/journal handling.
  - Impact: PWA flow became more operationally safe.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: Error handling] Save pattern moved toward explicit `Err.Raise`.
  - Previous behavior: mixed validation/UI patterns.
  - New behavior: write functions expected to propagate errors.
  - Why: rollback reliability and cleaner boundaries.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA] `SaveOtpremnica` corrected to align with `Err.Raise` pattern.
  - Symptom: inconsistent error handling.
  - Root cause: direct UI-style validation feedback in business save path.
  - Resolution: align with module propagation pattern.
  - Reference update required: Yes

### DEPRECATED
- [Layer: VBA] broad `On Error Resume Next` in backup and business paths.
  - Replacement: structured EH paths.
  - Removal target: ongoing cleanup.
  - Reference update required: Yes

### REMOVED
- No major removals documented.

### KNOWN ISSUES
- sections 4–6 were still inherited conceptually from v2.5 in later docs
- zbirna import remained TODO
- autosave remained TODO

### ROADMAP
- AutoSave after TX
- War Room dashboard
- Domain deployment
- notifications, photo documentation, PHI and LOT

### Migration Notes
- PWA import/export paths require consistent OTK/VOZ sheet headers.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v2.5 — 2026-03-22

### Summary
- First broad full snapshot that documented NOVAC, SEF, BANKA, UI patterns, modules and PWA platform in one place.
- Became the baseline for finance/SEF/bank architecture.

### ADDED
- [Layer: Docs] Full sections for money flow, SEF architecture, bank import and PWA platform.
  - Impact: architecture became substantially more reviewable.
  - Reference update required: Yes
  - Migration required: No
- [Layer: Finance/SEF] formal SEF state machine, endpoints, payload contract and classes.
  - Impact: e-faktura domain was fully specified.
  - Reference update required: Yes
  - Migration required: No
- [Layer: UI/Modules] explicit form overview, combo cascade patterns and module catalog.
  - Impact: UI/runtime architecture documented, not just business flows.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: Tables] schema snapshot expanded and stabilized across finance, SEF, bank, parcel and agro domains.
  - Previous behavior: narrower table coverage in older refs.
  - New behavior: broad schema catalog.
  - Why: architecture maturity.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: Docs] documentation completeness improved, but still not yet self-contained by later standards.
  - Symptom: partial earlier versions required cross-reading.
  - Root cause: evolving reference style.
  - Resolution: v2.5 provided a much fuller snapshot.
  - Reference update required: Yes

### DEPRECATED
- Informal understanding of SEF/bank flow outside documentation.
  - Replacement: explicit documented flows.
  - Removal target: immediate.
  - Reference update required: Yes

### REMOVED
- No explicit removals documented.

### KNOWN ISSUES
- PWA still described in a more roadmap/MVP framing
- multiple finance/reporting and parser fragility issues remained open
- MsgBox debt in business modules remained

### ROADMAP
- LogError
- remote/mobile PWA phases
- audit trail
- notifications
- future IoT/sensor upsell

### Migration Notes
- None beyond process/documentation adoption.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v2.4 — 2026-03-17

### Summary
- SEF v1.0 entered the system.
- Existing finance, bank and document modules were carried forward and formalized further.

### ADDED
- [Layer: VBA/SEF] SEF v1.0 with validation, outbound, persistence, payload, HTTP module and frmSEF.
  - Impact: automatic/structured e-faktura support introduced.
  - Reference update required: Yes
  - Migration required: Yes

### CHANGED
- [Layer: Architecture] reference now included SEF as first-class domain alongside bank import and mapping.
  - Previous behavior: no full active SEF stack.
  - New behavior: documented 3-phase TX split, dual status, recovery.
  - Why: e-faktura integration.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: Finance/docs] earlier data and report flows remained stabilized as part of broader refactor set.
  - Symptom: fragmented logic across modules.
  - Root cause: pre-v2.4 growth.
  - Resolution: reference captured integrated state.

### DEPRECATED
- Manual/non-structured e-faktura handling.
  - Replacement: formal SEF module family.
  - Removal target: immediate.
  - Reference update required: Yes

### REMOVED
- No explicit removals documented.

### KNOWN ISSUES
- SEF 429 handling
- ACCEPTED→STORNO transition
- auto-recovery on workbook open
- retry counter
- broader payload validation

### ROADMAP
- bank reconciliation
- banka UI
- audit trail
- config unification
- Kontni plan

### Migration Notes
- `tblFakture`, `tblSEFSubmission` and `tblSEFEventLog` become essential active data structures.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v2.3 — 2026-03-13

### Summary
- BankaImport v1.0 and BankaMapiranje v2.0 became active.
- Kartica Kooperanta and orphan-document checks were added.
- Finance/reporting and partner mapping matured substantially.

### ADDED
- [Layer: VBA/Finance] BankaImport parser/staging/folder workflow.
  - Impact: statement import entered active system.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: VBA/Finance] `modBankaMapiranje` with auto/learning mapping, block resolution and skip/error statuses.
  - Impact: imported payments could feed novac flow.
  - Reference update required: Yes
  - Migration required: No
- [Layer: Reports] Kartica Kooperanta with print support.
  - Impact: customer-facing financial statement became available.
  - Reference update required: Yes
  - Migration required: No
- [Layer: QA/ops] `GetVerwaisteDokumente`.
  - Impact: document integrity warnings surfaced in main UI.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: Finance] Novac naming and semantics became more explicit and modernized.
  - Previous behavior: older less-clear constant names.
  - New behavior: clear `NOV_KES_*`, `NOV_VIRMAN_*`, `NOV_KUPCI_*`.
  - Why: readability and correctness.
  - Reference update required: Yes
  - Migration required: Yes

### FIXED
- [Layer: Mapping] partner/OM matching and faktura matching became more robust.
  - Symptom: more manual finance reconciliation pain.
  - Root cause: insufficient staged mapping paths.
  - Resolution: learning table + multiple match strategies.
  - Reference update required: Yes

### DEPRECATED
- Older less-clear novac constant vocabulary.
  - Replacement: renamed novac constants.
  - Removal target: immediate
  - Reference update required: Yes

### REMOVED
- No explicit removals documented.

### KNOWN ISSUES
- LogError still missing
- BankaImport UI still missing
- several finance semantics still open

### ROADMAP
- BankaImport UI
- audit trail
- Kontni plan
- meteo and SEF future integrations

### Migration Notes
- adopt renamed novac constants in code and docs.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## v2.2.1 — 2026-03-08

### Summary
- Early modern baseline.
- Sledljivost v2.0, Agrohemija, novac renaming groundwork and report refactors were established.

### ADDED
- [Layer: VBA] Sledljivost v2.0 with `ParcelaID` in `tblOtkup`, trace enhancements and parcela-aware document tracing.
  - Impact: traceability became parcel-aware.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: VBA] Agrohemija module with parcels, artikli, magacin and warehouse/cart flows.
  - Impact: agro stock and issue workflows entered system.
  - Reference update required: Yes
  - Migration required: Yes
- [Layer: Reports] richer saldo and otkupljena roba reporting.
  - Impact: operator visibility improved.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: Core] Error-handling cleanup and helper extraction.
  - Previous behavior: more ad hoc error suppression and duplicate helpers.
  - New behavior: cleaner helper module patterns and stricter `ExcludeStornirano` use.
  - Why: maintainability.
  - Reference update required: Yes
  - Migration required: No
- [Layer: Finance] Novac terminology started shifting to clearer semantics.
  - Previous behavior: older `NOV_OM_*` / `NOV_AVANS` names.
  - New behavior: renamed finance constants.
  - Why: clarity and better business meaning.
  - Reference update required: Yes
  - Migration required: Yes

### FIXED
- [Layer: Core/VBA] multiple storno/report/data-access inconsistencies were cleaned.
  - Symptom: fragile report and linkage behavior.
  - Root cause: pre-refactor duplication and missing helper patterns.
  - Resolution: helper extraction, cleanup, relink/reset fixes.
  - Reference update required: Yes

### DEPRECATED
- Older novac constants and broader `On Error Resume Next` habits.
  - Replacement: clearer novac constants and structured error handling.
  - Removal target: ongoing.
  - Reference update required: Yes

### REMOVED
- No explicit removals documented.

### KNOWN ISSUES
- LogError missing
- BankaImport not yet done
- audit trail absent
- meteo and SEF still future roadmap at this point

### ROADMAP
- LogError
- BankaImport
- audit trail
- Kontni plan
- Meteo and SEF future integrations

### Migration Notes
- adopt parcela linkage in otkup/tracing flows.
- adopt renamed novac constants.

### Documentation Actions
- [x] Canonical reference updated
- [x] Source-of-Truth Matrix reviewed
- [x] Endpoint list reviewed
- [x] Known issues reviewed
- [x] Deprecated elements reviewed

---

## 4. Editorial Policy

### 4.1 What belongs here
U changelog ide:
- šta je novo
- šta je promenjeno
- šta je popravljeno
- šta je deprecated ili removed
- koji problemi su aktivni ili rešeni
- koji roadmap elementi su otvoreni ili zatvoreni

### 4.2 What does not belong here
U changelog ne ide:
- kompletan popis svih tabela
- kompletan popis svih endpoint-a
- kompletan snapshot role modela
- kompletan source-of-truth opis
- kompletna aktivna arhitektura

To ide isključivo u canonical reference.

### 4.3 Update Rule
Ako changelog stavka menja trenutno važeću arhitekturu, ista promena mora biti preneta i u canonical reference dokument.

---

## 5. Cross-Check Matrix

| Change Type | Must update reference? | Must review known issues? | Must review roadmap? | Must review migration notes? |
|---|---|---|---|---|
| New table | Yes | Maybe | Maybe | Maybe |
| New endpoint | Yes | Maybe | Maybe | Maybe |
| Invariant change | Yes | Yes | Maybe | Maybe |
| UI-only cosmetic fix | No | No | No | No |
| Sync behavior change | Yes | Yes | Maybe | Yes |
| Deprecation | Yes | Maybe | Yes | Maybe |
| Removal | Yes | Maybe | Maybe | Yes |

---

## 6. Release Readiness Checklist

Pre zatvaranja verzije proveriti:

- [x] Sve arhitektonski relevantne promene su upisane u changelog.
- [x] Sve aktivne promene su prenete u canonical reference.
- [x] Nema otvorenih delta beleški koje nisu prepisane u snapshot.
- [x] Known issues su ažurirani.
- [x] Roadmap status je ažuriran.
- [x] Deprecated i removed stavke su jasno označene.
- [x] Ako postoji migracija, dokumentovana je.

---

## 7. Suggested Naming Convention

- `ARCHITECTURE_REFERENCE.md` → canonical snapshot
- `ARCHITECTURE_CHANGELOG.md` → delta history

Opcionalno po verziji:
- `ARCHITECTURE_REFERENCE_v3_2.md`
- `ARCHITECTURE_CHANGELOG_v3_2.md`

Ali preporuka je da canonical fajl zadrži stabilno ime, a verzija da stoji u zaglavlju dokumenta.
