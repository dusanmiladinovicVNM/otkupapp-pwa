# AgriX Documentation Refactor — Change Summary

## What this pass created

- Cleaned skeleton: `ARCHITECTURE_REFERENCE.md`
- Cleaned skeleton: `ARCHITECTURE_CHANGELOG.md`
- Extracted verification/runbook home: `RELEASE_GATES.md`
- Extracted future-work home: `ROADMAP.md`
- Detailed migration staging document: `SECTION_MIGRATION_MAP.md`
- Archive folder with extracted v6.18/v6.19 architecture snapshots and legacy changelog material.

## What was removed from active AR skeleton

Nothing was destructively deleted from source files. In the new AR skeleton, the following are intentionally not kept as active main-body structures:

- Large `v6.xx Delta` sections.
- `v6.xx Baseline` sections as release framing.
- Full historical snapshot appendices.
- Detailed smoke/evidence checklists.
- Long roadmap/future-hardening lists.

## What was moved to archive

- `Appendix B. Integrated Permanent v6.18 Architecture Section` → `archive/ARCHITECTURE_REFERENCE_v6_18.md`
- `Appendix C. Integrated Permanent v6.19 Architecture Section` → `archive/ARCHITECTURE_REFERENCE_v6_19.md`
- Regular legacy changelog entries v6.17 and older → `archive/CHANGELOG_legacy_v2_to_v6_17.md`
- Integrated permanent changelog v6.19/v6.18 → separate archive files.

## What was deduplicated conceptually

- Exact-row lookup rules are centralized under Global Architecture Invariants.
- Fail-fast schema/write rules are centralized under Global Architecture Invariants.
- Best-effort monitoring rule is centralized.
- Stornirano/soft-delete rule is centralized.
- Historical release framing is routed to changelog/archive, while current rules remain in AR domain sections.

## What still needs human review

- Canonical Banka parser module name: `modBankaImport_PdfText` vs `modBankaImportParserPdfToText`.
- Whether v2.2–v6.9 should stay compact in CL or fully move to archive.
- Whether endpoint reference remains inside AR or becomes generated API docs later.
- Whether current production gates should be fully outside AR or summarized in AR.
- Whether `saveParcelPolygon` token/security decision is closed.
- Whether every active v6.18/v6.19 rule is represented before deleting old appendices from active AR.

## Recommended next step

Use `SECTION_MIGRATION_MAP.md` as the working checklist. Move content into the AR skeleton one domain at a time, starting with:

1. Global Architecture Invariants.
2. BankaImport and BankaMapiranje.
3. Storno.
4. PWA Sync/Offline.
5. Monitoring.
6. Document Flow.
7. Agrohemija.
8. GIS/Parcele.


---

## Pass 1 — BankaImport / BankaMapiranje content migration

Status: completed as a domain-first migration pass.

What moved into `ARCHITECTURE_REFERENCE.md`:

- `tblBankaImport` and `tblPartnerMap` current table contracts.
- `modBankaImport` current active capabilities.
- `modBankaMapiranje` current active capabilities.
- Banka PDF parser executable/config/temp-file contract.
- Statement header/saldo parse contract.
- Four statement saldo integrity gates.
- Fail-fast staging rules for `SaveBankaImportRows` and `IsDuplicateBankaImport`.
- Explicit import outcome categories.
- Deferred file move sequence after DB commit.
- Exact-row mapping guard for `NovacID`, `BankaImportID`, `OtkupID`, `FakturaID`.
- `LinkNovacToOtkupStrict` pattern.
- `GetBankaImportRowByID` legacy 1x10 semantic contract.
- `ValidateBankaImportNotProcessed` stornirano guard.
- `MapBankaImportAsKooperantBlockCore` allocation invariant.
- `frmBankaImport` review-shell boundary.
- Banka monitoring boundary.

What moved/expanded into `RELEASE_GATES.md`:

- PDF extract success/failure gates.
- Statement integrity negative tests.
- Append failure rollback tests.
- Deferred file move tests.
- Duplicate-key mapping tests.
- Stornirano/processed/skip guard tests.
- `GetBankaImportRowByID` legacy-shape smoke.
- Kooperant block allocation smoke.

What remains for later passes:

- Storno domain section.
- PWA sync/offline domain section.
- Monitoring domain section.
- Document flow domain section.
- Agrohemija domain section.
- GIS/Parcele domain section.


---

## Pass 2 — Storno content migration

Status: completed as a domain-first migration pass.

What moved into `ARCHITECTURE_REFERENCE.md`:

- `modStorno` canonical scope.
- Exact-row / missing-record / duplicate-key storno guard rules.
- `CanStorno` and `LookupActiveID` responsibilities.
- Storno flow boundary: soft-delete plus explicit repair, not physical delete and not automatic chain-wide cascade.
- Entity-specific side effects for `StornoOtkup`, `StornoOtpremnica`, `StornoZbirna`, `StornoPrijemnica`, `StornoFaktura`, `StornoNovac`.
- Storno transaction and recovery contract.
- Tables touched by storno paths.

What moved/expanded into `RELEASE_GATES.md`:

- Missing target gate.
- Duplicate target gate.
- Already-stornirano rejection gate.
- Required schema/write helper gates.
- Rollback gate.
- Per-entity side-effect verification gates.

What remains for later passes:

- PWA Sync/Offline domain section.
- Monitoring domain section.
- Document Flow non-storno sections.
- Agrohemija domain section.
- GIS/Parcele domain section.


---

## Pass 3 — PWA Sync / Offline content migration

Status: completed as a domain-first migration pass.

What moved into `ARCHITECTURE_REFERENCE.md`:

- PWA app-shell/cache contract.
- Runtime state ownership around `AppState` and `window.appRuntime`.
- IndexedDB active stores and access helpers.
- Offline-first guarantees.
- Canonical sync result shape for Otkupac, Kooperant and Vozač.
- Shared sync-engine failure/retry/diagnostic rules.
- Single sync trigger entrypoint through `syncQueueSafe(reason)`.
- Parallel sync in-flight guard.
- VBA/GAS/PWA master-sync soft-lock behavior.
- Bootstrap stale-`syncing` recovery.
- Alias-aware render dedupe contract.
- Business date helper rules and forbidden UTC slicing patterns.
- Shared formatting helpers.
- Submit-lock helper and canonical lock keys.
- Client error reporting contract.
- Service-worker cache/versioning contract.
- localStorage ban/exceptions.
- Role-level Otkupac, Kooperant, Vozač, Management and Excel Operator sync/offline workflow summaries.
- PWA-first `BrojZbirne` ownership and format rules.
- GAS sync endpoint and `syncTrosak` current contracts.
- Google Sheets `SyncControl` master-sync lock contract.

What moved/expanded into `RELEASE_GATES.md`:

- PWA app-shell/cache gate.
- Otkupac, Kooperant, Vozač and Management role smoke gates.
- Offline/stale-syncing recovery gate.
- Submit-lock gate.
- Client error reporting gate.
- Master-sync guard / soft-lock gate.
- Render dedupe gate.
- Business-date gate.

What remains for later passes:

- Monitoring domain section.
- Document Flow non-storno sections.
- Agrohemija domain section.
- GIS/Parcele domain section.
- SEF/security pass.


---

## Pass 4 — Monitoring and Observability content migration

Status: completed as a domain-first migration pass.

What moved into `ARCHITECTURE_REFERENCE.md`:

- Monitoring pipeline and non-transaction-participant boundary.
- Layer ownership for `modMonitoring`, `Monitoring.gs` and `OtkupApp_Monitoring_PROD`.
- VBA monitoring configuration through `tblSEFConfig`.
- GAS monitoring configuration through Script Properties.
- Canonical monitoring workbook tabs.
- Health component model and component inference rules.
- Event payload contract and severity set.
- Routing rules for Events, Errors, SyncStatus, UserSessions, SEFStatus, Backups, Alerts, AuditCritical and Health.
- Redaction, forbidden payload content and timeout rules.
- VBA app lifecycle monitoring.
- SEF monitoring event coverage.
- Business transaction monitoring coverage.
- Bank mapping monitoring boundary.
- MasterData/Stammdaten health events.
- Backup monitoring contract.
- Alert, AuditCritical and manual-review contract.
- Watchdog and scheduled jobs.
- Monitoring test contract.
- Deployment/data-migration boundary.

What moved/expanded into `RELEASE_GATES.md`:

- VBA monitoring configuration gate.
- GAS monitoring configuration gate.
- HTTP/connectivity gate.
- Workbook tabs gate.
- Event routing gate.
- Health component gate.
- SEF monitoring gate.
- Business transaction monitoring gate.
- Bank mapping monitoring gate.
- Backup monitoring gate.
- Watchdog and scheduled jobs gates.
- Alert/AuditCritical gate.
- Monitoring test suite gate.
- Deployment setup gate.

What remains for later passes:

- Document Flow non-storno sections.
- Agrohemija domain section.
- GIS/Parcele domain section.
- SEF/security pass.
- Reports/derived views pass.


---

## Pass 5 — Document Flow Non-Storno

Populated `ARCHITECTURE_REFERENCE.md` section `6. Document Flow Architecture` for the non-storno chain:

- `Otkup`
- `Otpremnica`
- `Zbirna`
- `Prijemnica`
- `Faktura`
- `SEF Submission Flow`
- `Sledljivost`
- `Ambalaža Ledger`

Key current contracts moved into AR:

- canonical chain `Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF`;
- `tblOtkup.OtpremnicaID` as trace bridge;
- PWA-first `BrojZbirne` with VBA fallback;
- VOZ column B/T split for technical ID vs business number;
- `PrijemnicaID` row-unique model and `BrojPrijemnice + Klasa` relink model;
- `CreateFaktura[_TX]` prijemnica-based invoice model;
- `AutoLinkOtkupOtpremnica_TX` exact-row/checked-update rule;
- `TrackAmbalaza` fail-fast ledger rules and strict `Ulaz` / `Izlaz` semantics.

Expanded `RELEASE_GATES.md` with BusinessFlowPro, Faktura and Document Flow Non-Storno gates.

Expanded `ROADMAP.md` with `modFaktura` exact-row hardening and targeted relink regression test.

NEEDS REVIEW added for one gate-script wording issue around `PrijemnicaID` spelling in the release gate text.

---

## Pass 6 — Agrohemija / Digitalni Agronom

Populated `ARCHITECTURE_REFERENCE.md` section `13. Agrohemija / Digitalni Agronom`.

Key current contracts moved into AR:

- `modAgrohemija` remains canonical desktop business module for agrohemija warehouse operations.
- `SaveMagacin` owns `MAG_ULAZ` / `MAG_IZLAZ` single-row journal creation.
- `SaveMagacin` validates article, movement type, positive quantity and required kooperant for `MAG_IZLAZ`.
- `MAG_IZLAZ` checks current stock before append.
- `SaveMagacin_TX` snapshots `tblMagacin`, rolls back failures and emits best-effort monitoring.
- `GetMagacinStanje()` is canonical desktop stock read model after excluding stornirano rows.
- `frmAgrohemija` owns `m_KorpaIzlaz` and `m_KorpaUlaz` baskets and commits each basket atomically through one `clsTransaction`.
- Issue baskets aggregate stock by `ArtikalID` before commit.
- PWA management `agrohemija.js` persists real unit-of-measure quantity, not package count.
- PWA management issuing serializes multiple parcel IDs with semicolon (`;`).
- `izdZavrsi()` opens printable/signable modal; final save uses submit lock.
- PWA kooperant `agromere.js` uses real-quantity dosage semantics and validates local stock from `stammdaten.magacinkoop`.
- `agroSaveTretman()` uses canonical lock key `agro:tretman:save`.
- GAS `syncTretman` accepted current contract is documented with role scope, `withLock`, `processTretmanRecord`, `ClientRecordID` idempotency and `TRETMAN-<KooperantID>` storage.
- GAS `saveIzdavanje` remains Management-only and not yet server-idempotent by stable client issuance ID.
- Cross-layer quantity invariant documented.
- Current limitations documented: treatment sync does not decrement server-side `magacinkoop`; `saveIzdavanje` server idempotency remains future hardening.

Expanded `RELEASE_GATES.md` with Agrohemija / Digitalni Agronom gates covering:

- `SaveMagacin` validation;
- `SaveMagacin_TX` transaction/monitoring behavior;
- `frmAgrohemija` basket rollback and aggregated stock check;
- PWA management issuing quantity/package semantics;
- PWA kooperant treatment quantity/local-stock validation;
- GAS `syncTretman` boundary;
- GAS `saveIzdavanje` current limitation;
- cross-layer quantity semantics.

Expanded `ROADMAP.md` with clearer future hardening notes for `saveIzdavanje` idempotency, treatment stock decrement and Agrohemija-related UI/business separation.


## Pass 7 — GIS / Parcele / Meteo content migration

Status: completed as a domain-first migration pass.

What moved into `ARCHITECTURE_REFERENCE.md`:

- `tblParcele` local-master ownership and Google `Stammdaten / Parcele` online projection boundary.
- `frmStammdaten` UI-only geo responsibility.
- `modGeoParcele` service/domain ownership.
- Canonical parcel geo micro-flow: select parcel, open external geo source, paste coordinates, save point, clear point, open map, open polygon editor.
- Transaction-backed `SaveParcelGeoPoint_TX` and `ClearParcelGeo_TX` rules.
- UTM zone 34 conversion contract.
- `SyncSelectedParcelaToGoogle` fail-closed contract.
- Relationship between selected-parcel sync and full-cycle `ImportParcelGeoFromGoogleToMaster`.
- PWA parcel map and parcel detail contract.
- Meteo scheduled fetch, cached-first read, risk-threshold and spray-window contracts.
- GIS/meteo endpoint surface and public exposure boundary.
- `frmOtkupAPP` KPI robustness contract.
- Anti-duplication and error-handling rules.
- Current geo/meteo risks.

What moved into `RELEASE_GATES.md`:

- VBA compile/dependency gate.
- Form designer gate.
- Geo UI/save/clear smoke.
- Selected parcel sync gate.
- Polygon editor and polygon overwrite safety gates.
- PWA parcel map gate.
- Meteo scheduled fetch, cached-first read and risk/spray-window gates.
- KPI robustness gate.

What moved into `ROADMAP.md`:

- Dedicated geo editor/auth model for `saveParcelPolygon`.
- Possible true selected-parcel Google upsert.
- Polygon clear / geometry lifecycle decision.

Still needs human review:

- Whether the public/pre-auth `saveParcelPolygon` exception remains acceptable for launch or must be locked before production.
- Whether full `Parcele` export inside selected-parcel sync is operationally fast enough.
- Whether point clear should remain point-only forever or a separate polygon-clear feature is needed.


## Pass 8 — Security and Compliance

Populated `ARCHITECTURE_REFERENCE.md` section 16 with current-state security contracts:

- SEF HTTPS-only rule;
- secret/config storage boundaries;
- token/session scope;
- GAS endpoint authorization classes;
- role/entity ownership rules;
- monitoring redaction and privacy boundary;
- CSP/PWA asset rules;
- `localStorage` / IndexedDB local-state rules;
- local workstation config security;
- public geo/meteo and `saveParcelPolygon` review boundary;
- input validation/sanitization rules.

Expanded `RELEASE_GATES.md` with Security and Compliance gates covering SEF HTTPS, GAS authz matrix, sync ownership, write locks, fiscal/master-data authorization, monitoring redaction, CSP/assets, local-state audit, workstation config and public geo/polygon review.

Updated `ROADMAP.md` with endpoint authorization audit, secret/redaction regression suite and CSP/local-state audit items.

NEEDS REVIEW retained: deployed `saveParcelPolygon` authorization state is inconsistent across source documentation and must be verified before production handoff.


## Pass 9 — Reports and Derived Views

Populated `ARCHITECTURE_REFERENCE.md` section 17 with current-state report/read-model architecture:

- report/read-model ownership rule;
- source-of-truth distinction between canonical transaction tables and derived/exported views;
- report/view inventory covering finance, management, role dashboards, meteo and monitoring views;
- management reporting and dispatch planning-only boundary;
- finance report boundaries for `tblNovac`, `tblBankaImport`, faktura, otkup and magacin facts;
- kooperant, otkupac, vozač and management role-view rules;
- materialized/exported view rules;
- dirty-data KPI robustness rules;
- report safety rules for stornirano exclusion, technical/business ID separation, schema failures and freshness.

Expanded `RELEASE_GATES.md` with Reports and Derived Views gates for management reports, financial reports, agrohemija/warehouse reports, PWA role views, dashboard KPI robustness and monitoring read models.

Expanded `ROADMAP.md` with report/derived-view audit and report ownership cleanup tasks.

No active architecture rule was intentionally removed. Historical report-refactor notes remain owned by `ARCHITECTURE_CHANGELOG.md` / archive material.


## Pass 10 — VBA / Excel Desktop Architecture

Populated AR section `4. VBA / Excel Desktop Architecture` with current-state desktop rules:

- workbook role and single-operator desktop-master boundary;
- `modDataAccess` table/array access model;
- fail-fast schema/update helper model;
- `clsTransaction` snapshot/rollback semantics;
- `Workbook_Open` / `StartApp` / `ShutdownApp` lifecycle contract;
- local workstation setup through `modSetup`, `tblLocalConfig` and `Setup-OtkupApp.ps1`;
- `AutoSaveAfterCommit` durability rule;
- local journal/backup/recovery-warning layer;
- `RunProductionHealthCheck` workbook launch gate;
- shared helper module ownership;
- UI/business boundary for forms and business modules.

Expanded `RELEASE_GATES.md` with lifecycle, setup, AutoSave/journal/backup and production-health gates.



## Pass 11 — Data Architecture

Populated AR section `5. Data Architecture` as a current-state data contract instead of a placeholder.

Added:

- data ownership principles;
- canonical entity inventory;
- master table, transaction table, document table and finance table groups;
- BankaImport table and `tblPartnerMap` contracts;
- Agrohemija table/projection contracts;
- monitoring workbook data contract;
- Google Sheets transport table/family contract;
- identity and sync field rules;
- derived/exported read model rules;
- required schema guard rules;
- current data migration/backfill boundary.

Moved detailed checks into `RELEASE_GATES.md` under `12. Data Architecture Gates`.

Added roadmap entries for data inventory/schema-contract audit and source-of-truth/derived-view cleanup.


---

## Pass 12 — GAS API Architecture

Populated `ARCHITECTURE_REFERENCE.md` section `9. GAS API Architecture` as a current-state API contract.

Added / clarified:

- GAS role as online API/transport layer, not formal finance/document source of truth.
- `doPost` / `doGet` routing contract.
- Apps Script deployment and workbook topology.
- Auth/session token model with cache + PropertiesService fallback.
- Role/entity authorization rules.
- Public/pre-auth exceptions.
- Current action surface table.
- Response/failure envelope rules.
- Sync endpoint contract and canonical processors.
- `syncTrosak` active endpoint contract.
- Master-sync soft-lock readout and write blocking.
- Schema drift behavior.
- Management/dispatch/agro issuing endpoints.
- Meteo/GIS/fiskalni endpoint boundaries.
- `logClientError`, `ErrorLog`, monitoring ingest boundary.
- Locking/idempotency/concurrency rules.
- Disabled endpoint behavior.

Expanded `RELEASE_GATES.md` section `3. GAS Gates` with concrete checks for auth, token maintenance, sync, schema drift, master-sync soft-lock, management/dispatch, GIS/meteo/fiskalni, monitoring/ErrorLog and disabled endpoints.

Updated `ROADMAP.md` with GAS endpoint contract audit and GAS fixture sync suite items.

NEEDS REVIEW carried forward:

- `saveParcelPolygon` authorization state conflicts across source documentation and must be verified against deployed `Code.gs`.


## Pass 13 — Google Sheets Data Layer

Added a full current-state Google Sheets data-layer section covering Stammdaten, role-specific transport sheets, OTK/VOZ headers, SyncStatus/writeback semantics, per-kooperant sheets, Kartice, MgmtReports, monitoring workbook references, ErrorLog/LoginLog, SheetRegistry/header drift, SyncControl/master lock, parcel geo pull before export, Google external side-effect limits and manual-edit/test-data hygiene.

Expanded `RELEASE_GATES.md` with Google Sheets gates for Stammdaten, OTK, VOZ, per-kooperant sheets, exported read models, SheetRegistry/header drift, SyncControl/master lock, parcel geo pull and external writeback boundaries.

Expanded `ROADMAP.md` with Google Sheets schema/header audit and Google external writeback boundary cleanup items.


## Pass 14 — Current Production Gates / Final Validation

Status: completed.

What changed:

- Expanded AR `18. Current Production Gates` into a complete current-state launch-readiness summary.
- Added separate gate summaries for compile, business flow, Storno, BankaImport/Mapiranje, SEF, GAS, Google Sheets, PWA, Monitoring, Security, Data/Reports and ProductionHealthCheck.
- Expanded AR `19. Current Known Risks` into launch risks, accepted operational risks, `NEEDS REVIEW` items and technical debt.
- Expanded AR `20. Current Roadmap Summary` with schema/header registry, monitoring tuning and documentation QA items.
- Expanded `RELEASE_GATES.md` current mandatory gate matrix.
- Added `RELEASE_GATES.md` final validation gate.
- Restored archive files into final package.
- Added `FINAL_VALIDATION_REPORT.md`.

Remaining human review:

- Confirm deployed `saveParcelPolygon` authorization state.
- Confirm canonical Banka PDF parser module name.
- Confirm old v2.2–v6.9 changelog retention policy.
- Confirm whether endpoint tables should later become generated API docs.


## Pass 15 — Fix Omissions

This pass closes the previously acknowledged documentation gaps:

- filled `2. Source of Truth and Ownership Matrix`;
- filled `7. Finance Architecture` beyond BankaImport/BankaMapiranje;
- filled `21. Deprecated and Transitional Elements`;
- filled `22. Glossary`;
- replaced the placeholder changelog with a normalized version-history file;
- reconciled GAS endpoint authorization into a single matrix with `saveParcelPolygon` preserved as `NEEDS REVIEW`.

Remaining human review is intentionally narrow: deployed `saveParcelPolygon` auth state, canonical Banka PDF parser module naming, and whether future CL splitting should be more granular.


---

## Pass 16 — Known Issues Consolidation

Added `KNOWN_ISSUES.md` because known issues were previously spread across AR, ROADMAP, RELEASE_GATES and CL.

Changes:

- Added a dedicated active known-issues register.
- Updated AR companion docs and section 19 to point to `KNOWN_ISSUES.md`.
- Updated ROADMAP to clarify that roadmap tracks remediation, while `KNOWN_ISSUES.md` tracks active issues/accepted risks.
- Updated FINAL_VALIDATION_REPORT with a Pass 16 addendum.

Main active issues now centralized:

- `saveParcelPolygon` deployed authorization conflict.
- Banka PDF parser canonical module name conflict.
- Runtime gates not executed as part of documentation refactor.
- Production workbook still requires `RunProductionHealthCheck`.
- Endpoint matrix requires deployed `Code.gs` reconciliation.

## Pass 17 — GO Hardening Closeout Update

Integrated the user-supplied post-cutoff GO hardening diff for Google Sync, MasterSync, Novac and ProductionHealthCheck.

### Updated AR

Added current-state rules for:

- PWA unlock failure as degraded/partial result, not green success.
- Google Sheets staging/verify/replace full-tab writes.
- Google Sheets sheetId cache, quota retry/throttle and phased target replacement.
- Kartice named-tab export through `KARTICE_TAB_NAME = "Kartice"`.
- MasterSync exact-row guards for document-chain links.
- `ImportVOZRow_RowTX` row-level transaction ownership.
- `SaveNovac` hard-fail behavior for empty `GetNextID` and `AppendRow <= 0`.
- `ProductionHealthCheck` duplicate-key preflight.

### Updated CL

Added `v6.22 GO Hardening Closeout — Google Sync / MasterSync / Novac / HealthCheck` with added/changed/fixed/removed/validation/follow-up sections.

### Updated gates and issue registers

- Added Pass 17 GO hardening gates to `RELEASE_GATES.md`.
- Added non-blocking GoogleSheets helper cleanup items to `ROADMAP.md`.
- Added Pass 17 closeout notes to `KNOWN_ISSUES.md`.
- Added Pass 17 addendum to `FINAL_VALIDATION_REPORT.md`.


---

## v6.22 Change Summary

Added after the user-provided post-cut review:

- Architecture Reference now states `PrintFaktura` and `UpdateFakturaStatus` duplicate-`FakturaID` guards as current contract.
- Architecture Reference now makes `SaveParcelGeoPointByID[_TX]`, `ClearParcelGeoByID[_TX]` and `RequireSingleParcelaRow` the canonical parcel geo save/clear API.
- Row-index parcel geo functions are marked compatibility-only.
- Changelog includes a new v6.22 entry.
- Release gates include Faktura duplicate guard and Geo ByID gates.
- Roadmap marks Faktura print/status duplicate guard closed and tracks row-index geo wrapper retirement.


---

## v6.23 Change Summary

Added after the user-reported browser-tested PWA read-model correction:

- Architecture Reference now states `tblOtkup` is the canonical otkup master and `MgmtReports/OtkupiAll` is the PWA master read projection.
- Architecture Reference now states `OTK-ST-*` / `OTK-*` is operational inbound/live queue state, not the sole historical read source for otkup display.
- PWA Management and Otkupac otkup views now require merge of `OtkupiAll + OTK-ST-*` / `OTK-*` with duplicate removal.
- Otkup overview dedup priority is explicitly `ServerRecordID` / `OtkupID`, then `ClientRecordID`, then natural-key fallback.
- Changelog includes a new v6.23 entry.
- Release gates include Management/Otkupac browser smoke and otkup merge dedup gates.
- Known issues records prior PWA missing-master-otkup behavior as resolved by v6.23.

Status: integrated into package based on user-reported browser testing.
