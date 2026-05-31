# AgriX / OtkupApp Architecture Reference

**Version:** v6.24 canonical snapshot  
**Last Updated:** 2026-05-30  
**Status:** Canonical / Active Reference — v6.24 Vozač/Dispečer bugfix pass integrated
**Owner:** Architecture documentation compiled from supplied reference set  
**Audience:** Engineering, Product, Operations, Onboarding, Review  

> For historical deltas, see `ARCHITECTURE_CHANGELOG.md`.  
> For smoke suites and launch acceptance gates, see `RELEASE_GATES.md`.  
> For future hardening and cleanup items, see `ROADMAP.md`.  
> For active known issues and accepted risks, see `KNOWN_ISSUES.md`.

---

## 0. Document Contract

### 0.1 Purpose
Ovaj dokument je potpuni i samodovoljni snapshot trenutno važeće arhitekture sistema na datum verzije.

### 0.2 Canonical Snapshot Rule
Architecture Reference mora opisivati **šta je trenutno važeće**, ne hronologiju kako se do toga došlo. Svaka nova verzija mora restate-ovati važeće elemente arhitekture bez oslanjanja na starije reference fajlove.

### 0.3 Reading Rule
Ako nešto nije navedeno u ovom dokumentu, ne smatra se canonical arhitekturom dok ne bude eksplicitno uneseno.

### 0.4 Companion Documents
- `ARCHITECTURE_CHANGELOG.md` — istorija promena po verzijama.
- `RELEASE_GATES.md` — smoke, regression, launch i production gate checkliste.
- `ROADMAP.md` — aktivni i odloženi hardening/refactor rad.
- `KNOWN_ISSUES.md` — konsolidovani aktivni known issues, accepted risks i NEEDS REVIEW stavke.
- `archive/` — istorijski snapshotovi i legacy changelog materijal.

### 0.5 Editorial Rules
U glavnom reference dokumentu ne koristiti release-note formulacije kao primarni oblik dokumentacije: `unchanged from`, `same as previous`, `others omitted`, `see earlier version`, `rest unchanged`. Takve napomene pripadaju changelog-u ili arhivi.

---

## 1. Current System Overview

### 1.1 Product Scope
- **Naziv proizvoda:** AgriX / OtkupApp
- **Primarna namena:** digitalizacija srpskog lanca otkupa voća i paralelno farm-management sloja za kooperante.
- **Glavni korisnici:** Otkupac, Kooperant, Vozač, Management, Excel operator.
- **Operativni kontekst:** offline-first PWA na terenu + Google Sheets/GAS kao online transportni sloj + Excel/VBA kao desktop master i finansijsko-dokumentni backbone.

### 1.2 High-Level Architecture
Sistem ima 4 glavna sloja:

- **VBA/Excel desktop backend** — master data management, dokumentni tok, finansije, SEF, BankaImport, izveštaji.
- **Google Apps Script API** — action-router, auth, sync, meteo, dispatch, fiskalni i pomoćni endpoints.
- **Google Sheets data layer** — stammdaten, role-specific sheets, kartice, meteo, dispatch, fiskalni i management report tabovi.
- **AgriX PWA** — offline-first klijent sa 4 role-specific UI toka i IndexedDB lokalnim storage-om.

### 1.3 Deployment Model
The deployment model is per-client/per-firm and separates desktop, GAS, Google Sheets and PWA responsibilities.

### 1.4 Supported Roles
The supported roles are Otkupac, Kooperant, Vozac, Management and Excel operator.

### 1.5 Supported Platforms
Supported platforms are Excel Desktop, web/PWA, Android, iOS and browser offline mode within the documented constraints.

---

## 2. Source of Truth and Ownership Matrix

This section defines ownership of canonical facts, transport projections, local caches and derived read models. Historical changes in ownership belong in `ARCHITECTURE_CHANGELOG.md`; this section states the active contract.

### 2.1 Business Source of Truth

| Domain | Canonical source of truth | Online/shared projection | Local/offline copy | Primary writer | Notes |
|---|---|---|---|---|---|
| Master data (`Kooperanti`, `Kulture`, `Parcele`, `Users`, `Kupci`, `Vozaci`, `Artikli`, `Stanice`) | Excel/VBA master tables | `Stammdaten` workbook | PWA cached stammdaten | Excel operator / VBA export | PWA must not invent master data except where a Management endpoint explicitly creates controlled records such as `createArtikal`. |
| Otkup field records | `tblOtkup` after MasterSync import | `OTK-*` sheets | PWA `otkupi` store | Otkupac PWA first, then VBA import | Google/PWA records are operational transport until imported into the desktop master. |
| Zbirna transport records | `tblZbirna` after MasterSync import | `VOZ-*` sheets | PWA `zbirne` store | Vozac PWA first, then VBA import | `BrojZbirne` is a business number; `ServerRecordID` / `ZbirnaID` are technical IDs. |
| Document chain | Excel/VBA tables | exported reports only | PWA read models where supplied | Excel operator / VBA document modules | `Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF` remains desktop-owned after import. |
| Finance (`tblNovac`, avans, salda, faktura payment state) | Excel/VBA finance tables | `Kartice`, `MgmtReports`, saldo exports | PWA read models | Excel operator / VBA finance modules | BankaImport staging and BankaMapiranje feed this layer but do not replace it. |
| BankaImport staging | `tblBankaImport` | none unless exported | desktop form preview | Excel operator / VBA import | Imported bank rows are financial facts/staging records and must preserve parser/integrity metadata. |
| SEF | Excel/VBA SEF tables/event logs plus SEF API statuses | monitoring workbook summaries | desktop form state | Excel operator / SEF service modules | `SEFWorkflowState` and `SEFStatus` are intentionally separate. |
| Kooperant treatment evidence | Per-kooperant treatment sheets | `TRETMAN-<KooperantID>` | PWA `tretmani` store | Kooperant/Management PWA via GAS | Evidence/history/karenca sync; server-side stock decrement is not active behavior. |
| Kooperant expenses | Per-kooperant expense sheets | `TROSKOVI-<KooperantID>` | PWA `troskovi` store | Kooperant/Management PWA via GAS | Active `syncTrosak` endpoint owns the batch transport contract. |
| Fiscal private records | Per-kooperant fiscal sheets | `FISKALNI-<KooperantID>` | PWA fiscal UI/cache | Kooperant/Management via GAS | Shared fiscal mapping is Management-only. |
| Dispatch / demand / kamion status | Google Sheets operational state | dispatch sheets | PWA management/driver state | Management / Vozac where allowed | Dispatcher state is shared online operational state, not desktop canonical finance/document state. |
| Parcel geo and meteo | Excel `tblParcele` for master data; dedicated geo workbook for online geo/meteo projection | `GEO_SPREADSHEET_ID` tabs | PWA map/meteo cache | VBA/GAS according to endpoint | `saveParcelPolygon` authorization state remains `NEEDS REVIEW`. |
| Monitoring | `OtkupApp_Monitoring_PROD` | monitoring workbook tabs | none | VBA/PWA/GAS event sources | Monitoring is best-effort and not a business transaction participant. |
| Reports / dashboards | Derived from canonical sources | `MgmtReports`, `Kartice`, `SaldoOMDetail`, `MeteoLatest` | PWA caches | VBA/GAS export/read models | Derived views are never authoritative transaction sources. |

### 2.2 VBA / Excel Ownership

VBA/Excel owns the canonical desktop master and all finance/document/tax-critical business state after import.

Canonical VBA ownership includes:

- master data authoring and exported Stammdaten state;
- `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke` and `tblAmbalaza`;
- `tblNovac`, avans allocation, payment-status recompute, saldo reports and kartice export;
- `tblBankaImport`, `tblPartnerMap`, BankaImport parser/import and BankaMapiranje reconciliation;
- SEF local workflow state, SEF event log, submission persistence and desktop recovery;
- storno and rollback behavior for desktop-owned records;
- workbook setup, LocalConfig, backup, journal and ProductionHealthCheck.

VBA transaction wrappers can rollback only local Excel-table snapshots. External effects such as Google writeback, Drive writes, SEF HTTP calls, filesystem moves and monitoring sends must be sequenced or documented as non-transactional external side effects.

### 2.3 GAS Ownership

Google Apps Script owns online request routing, auth, role/entity authorization, lock-protected Google Sheets writes and selected external-service bridges.

Canonical GAS ownership includes:

- `doPost` / `doGet` action routing and structured JSON responses;
- login token creation, token validation and token purge;
- role/entity checks for Otkupac, Kooperant, Vozac and Management;
- idempotent sync writes to `OTK-*`, `VOZ-*`, `TRETMAN-*`, `TROSKOVI-*`, `OPREMA-*` and fiscal sheets;
- `getMasterSyncState` and master-sync write-blocking semantics;
- dispatch/demand/kamion status endpoints;
- fiscal parse/save/mapping endpoints;
- geo/meteo read bridges and scheduled meteo fetch;
- client error logging and monitoring ingest.

GAS does not own desktop finance/document truth unless the architecture explicitly says that a domain is online-operational rather than desktop-canonical.

### 2.4 Google Sheets Ownership

Google Sheets is the transport/projection layer between PWA, GAS and desktop VBA.

Canonical Google Sheets ownership includes:

- `Stammdaten` as exported master-data projection, not the primary authoring source;
- role-specific sheets as PWA-to-desktop transport queues;
- per-kooperant private sheets for treatment, expense, equipment and fiscal records;
- `Kartice`, `MgmtReports`, `SaldoOMDetail` and other derived read models;
- `SyncControl` / `MASTER_SYNC_LOCK` for master-sync coordination;
- `ErrorLog`, `LoginLog` and monitoring workbook tabs as operational logs.

Existing sync-sheet schema drift is a failure. GAS may create headers only for empty new sheets; it must not silently append missing columns to a populated production sheet.

### 2.5 PWA Ownership

The PWA owns role-specific field UX, local-first record creation and IndexedDB queue state.

Canonical PWA ownership includes:

- Otkupac local otkup queue and post-save sync trigger;
- Vozac local zbirna creation, PWA-first `BrojZbirne` generation and printable last-station workflow;
- Kooperant treatment, expense, equipment and fiscal UX;
- Management dashboards, dispatch/agro issuing/fiscal mapping UI surfaces;
- IndexedDB pending/syncing/synced/error lifecycle;
- local render dedupe, submit locks, stale-syncing recovery and client error reporting.

IndexedDB is a local-first queue/cache, not shared canonical truth. `localStorage` is allowed only for device-local preferences and lightweight helper state; shared operational state belongs in IndexedDB + GAS/Sheets or in canonical desktop tables.

DOM ID uniqueness rule: dynamically created modal elements must use IDs that are unique relative to any static elements in `index.html`. In particular, signature pad canvases (`initSignaturePad`, `clearSignature`, `getSignatureData`, `destroySignaturePad`) rely on `getElementById`; a modal canvas ID that collides with a static hidden canvas causes the pad to bind to the wrong element. Modal-specific canvas IDs must be distinct from any static canvas IDs (e.g. `sigKooperantOL` for the otkupni-list modal, not the shared `sigKooperant` used in the otpremnica view).

### 2.6 Technical IDs vs Business Document Numbers

| Identifier | Meaning | Owner | Canonical rule |
|---|---|---|---|
| `ClientRecordID` | local PWA idempotency key | PWA | Must be stable across retries and used by GAS to avoid duplicate inserts. |
| `ServerRecordID` | technical GAS/PWA server-side sync ID | GAS | Must not be used as a business document number. |
| Desktop IDs such as `OTK-*`, `ZBR-*`, `NOV-*`, `BIM-*` | desktop/master row identity | VBA | Allocated by desktop ID helpers or import logic, depending on flow. |
| `BrojZbirne` | business transport document number | PWA first, VBA fallback | Format `x/ddmmyy[-rb]`; preserved through VOZ column T and desktop import. |
| `BrojPrijemnice` | business prijemnica number | VBA | May group class rows; row identity remains `PrijemnicaID`. |
| `PrijemnicaID` | physical `tblPrijemnica` row ID | VBA | Unique per class row; faktura creation uses row identity. |
| `SEFDocumentId` | external SEF API document identity | SEF API / VBA persistence | Required for refresh/recovery once remote submission may exist. |

### 2.7 LocalConfig vs Remote Config

| Config surface | Owner | Intended content | Must not contain |
|---|---|---|---|
| `tblLocalConfig` | desktop setup / `modSetup` | workstation-local paths such as `PDFTOTEXT_EXE_PATH` | Google/PWA tenant config |
| `tblConfig` / exported `Config` | desktop master / PWA projection | PWA/Google business/runtime config | workstation-local executable paths |
| `tblSEFConfig` | desktop config / SEF/monitoring modules | SEF API config, monitoring endpoint/secret/env | PWA role data or local PDF parser paths |
| GAS Script Properties | Apps Script deployment | monitoring spreadsheet ID, alert email, ingest secret, token fallback | source-controlled secrets |
| PWA runtime config | deployed frontend | API endpoint, role/entity runtime constants | secrets that grant server-side authority |

---

## 3. Global Architecture Invariants

This section consolidates system-wide invariants that apply across modules unless a domain section explicitly narrows them.

### 3.1 Exact-Row Lookup Rule
Critical ID-based operations must require exactly one matching row unless explicitly documented as multi-row by business design. `Count = 0` is missing-record error. `Count > 1` is duplicate-key error.

### 3.2 Fail-Fast Schema Rule
Critical schema reads use `RequireColumnIndex`; missing required table/column is a hard error.

### 3.3 Checked Write Rule
Critical writes use `RequireUpdateCell` or equivalent checked write helper. Silent write failure is not acceptable.

### 3.4 Transaction Rollback Rule
`_TX` wrappers must rollback on hard failure. Side effects outside workbook state require explicit sequencing because they cannot rollback transactionally.

### 3.5 Best-Effort Monitoring Rule
Monitoring is operational visibility, not a business transaction participant. Monitoring failure must not convert a committed business operation into a failure.

### 3.6 Stornirano / Soft-Delete Rule
Soft-deleted/stornirano rows must be excluded from active business flows unless the flow explicitly documents historical/audit behavior.

### 3.7 No MsgBox-Controlled Business Logic Rule
Business-layer hardening should not depend on `MsgBox` as a control-flow mechanism. Forms own operator messaging; service/business modules raise errors or return structured results.

### 3.8 External Side-Effect Rule
File moves, SEF submissions, Google writes and other external side effects must be sequenced so local transaction success/failure remains truthful.

### 3.9 Idempotency Rule
Sync/import endpoints and staging flows should key idempotency by stable client/server identifiers where available.

### 3.10 Offline-First Rule
PWA role flows must tolerate offline entry, delayed sync, stale `syncing` recovery and normalized sync results.

---

## 4. VBA / Excel Desktop Architecture

### 4.1 Workbook Role

The Excel/VBA workbook is the canonical desktop master for backoffice operations. It owns master-data maintenance, document-chain creation, finance logic, SEF orchestration, BankaImport/BankaMapiranje, desktop reporting, production-health checks and operator-controlled synchronization with Google Sheets.

The workbook is intentionally a single-operator desktop backend, not a multi-user database server. Multi-user field capture belongs to the PWA/GAS/Google Sheets transport layer; final canonical desktop integration belongs to Excel/VBA import and transaction wrappers.

Canonical desktop responsibilities:

- maintain canonical business tables such as `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke`, `tblNovac`, `tblAmbalaza`, `tblMagacin`, `tblBankaImport` and master-data tables;
- run document, finance, SEF, BankaImport, Agrohemija and Stammdaten operations through VBA modules and forms;
- preserve rollback, journaling, backup and production-health behavior locally;
- emit best-effort monitoring without making monitoring part of the business transaction;
- export/sync approved data projections to Google Sheets and PWA surfaces.

### 4.2 Table Access Model

All VBA sheet/table access must go through approved helper layers. Business modules must not bypass these helpers with ad hoc range reads/writes.

Canonical surfaces:

| Surface | Responsibility |
|---|---|
| `modDataAccess` | canonical `ListObject`/array bridge, table resolver, append/update/search helpers, ID generation, duplicate checks and stornirano-aware helpers |
| `GetTable()` | workbook-wide resolver for named `ListObject` tables |
| `GetTableData()` | canonical table-to-2D-array read helper; returned arrays are VBA 1-based |
| `AppendRow()` | canonical row append helper and journaling hook |
| `UpdateCell()` | standard cell update helper where fail-fast guard is not required |
| `FindRows()` | row-location helper over full table data; update flows must use full-table row indexes, not filtered arrays |
| `LookupValue()` | simple lookup helper for non-critical reads |
| `CheckDuplicate()` | user-facing duplicate guard; skips rows marked `Stornirano="Da"` |
| `ExcludeStornirano()` | read-helper filter for active business rows; safe no-op on tables without `Stornirano` |

Critical schema reads and writes must use the fail-fast guard layer:

| Helper | Rule |
|---|---|
| `RequireColumnIndex` | required columns must exist before indexed read/write logic continues |
| `RequireColumns` | multi-column schema precondition check |
| `RequireUpdateCell` | critical updates must hard-fail when the target write cannot be confirmed |

Filtered arrays are allowed for read-only reporting and aggregation. They must not be used as row-index sources for `UpdateCell()` / `RequireUpdateCell()`. Critical update paths must operate on row indexes from full table data or exact-row helpers.

### 4.3 Transaction Model

The desktop transaction model is snapshot/rollback over Excel tables. `_TX` wrappers snapshot every table that may be mutated, delegate to base business logic, commit only after all required writes and side effects succeed, and rollback on hard failure.

Canonical transaction rules:

- every production write path has a `_TX` wrapper or is called inside an existing wrapper that has already snapshotted affected tables;
- nested business logic must not start an independent transaction when the caller already owns the transaction boundary;
- `AppendRow()` and `UpdateCell()` are rollback-safe only when the affected table has been snapshotted before the write;
- external side effects such as Google writeback, file moves, SEF network calls and monitoring sends are not automatically rollback-able;
- `_TX` wrappers must not silently commit when core import/save/writeback functions report errors;
- `_TX` wrappers propagate or return failure in a way the UI layer can surface without guessing;
- best-effort monitoring after commit must not convert a committed business operation into a false failure.

`clsTransaction.CommitTx` is also the central post-commit persistence hook. After successful in-memory commit cleanup, it calls `AutoSaveAfterCommit(sourceName)` best-effort. `sourceName` is built from the transaction snapshot table names before cleanup, e.g. `clsTransaction[tblOtkup,tblAmbalaza,tblNovac]`.

### 4.4 App Lifecycle

Desktop startup is centralized in `modMain.StartApp` and invoked from `ThisWorkbook.Workbook_Open`.

Canonical startup contract:

- `Workbook_Open` is only a safe entry wrapper;
- business work, imports and long-running side effects must not be added directly to `Workbook_Open`;
- `Workbook_Open` wraps `StartApp` in error handling and must restore `Application.Visible = True` before any user-facing error message;
- `StartApp` calls `InitApp` on first run;
- `InitApp` suspends `ScreenUpdating`, `Calculation` and `EnableEvents` while validating the workbook runtime;
- startup validation runs `ValidateAllTables` against the canonical table set and surfaces missing-table warnings;
- after initialization, `StartApp` may hide Excel, show `frmSplash`, and hand off to `frmOtkupAPP`;
- file imports such as `ImportBankaInbox_TX()` are explicit operator actions and must not run automatically at boot;
- SEF stuck-state recovery may run opportunistically but must not block boot;
- journal recovery warnings are advisory and must keep the operator in control.

Canonical shutdown contract:

- normal exits route through `ShutdownApp`;
- `ShutdownApp` restores `Application.Visible = True`;
- `ShutdownApp` unloads the shell and writes `LogAppShutdown`;
- form-control close on `frmOtkupAPP` and workbook-level `Workbook_BeforeClose` share the same shutdown contract;
- startup and shutdown logs must form a paired lifecycle trail when the app closes normally.

### 4.5 Setup and Local Workstation Config

`modSetup` owns local workstation setup state. Local workstation configuration is separate from Google/PWA remote configuration.

Canonical local setup storage:

```text
Sheet: LocalConfig
Table: tblLocalConfig
Columns: Kljuc | Vrednost | Opis
```

`GetLocalConfigValue` and `SetLocalConfigValue` are public helpers and may be used by desktop modules that need workstation-local settings, including Banka parser/import modules.

Configuration boundary:

| Store | Purpose |
|---|---|
| `tblLocalConfig` | workstation-local paths and machine setup values |
| `tblConfig` | Google/PWA/shared application configuration |
| `tblSEFConfig` | SEF and monitoring-related desktop runtime configuration |

`SetupNewPC` / health-check flow should initialize or validate local requirements, including `PDFTOTEXT_EXE_PATH` for bank-statement parsing.

`Setup-OtkupApp.ps1` is the clean-PC bootstrapper. Active responsibilities:

- create `C:\OtkupApp` or the configured install root;
- create core folders such as `Backups`, `Logs`, `Journal`, `Export`, `Temp`, `Secrets` and `Bank_Izvodi`;
- copy `OtkupApp.xlsm`;
- unblock the workbook;
- optionally install the VBA publisher certificate;
- add the Excel trusted location;
- create the desktop shortcut;
- optionally copy bundled tools such as Poppler under `Tools\poppler\Library\bin\pdftotext.exe`.

### 4.6 AutoSave After Commit

Every successful desktop business transaction is expected to become durable without relying on later manual operator save.

Canonical AutoSave rules:

- `clsTransaction.CommitTx` is the only central trigger point;
- `AutoSaveAfterCommit(sourceName)` runs after successful commit cleanup;
- AutoSave is best-effort and must not raise back to the business caller because the in-memory transaction has already committed;
- AutoSave suppresses `Application.DisplayAlerts` during `ThisWorkbook.Save` and restores it on every exit path;
- AutoSave guards read-only workbooks;
- AutoSave guards unsaved/no-path workbooks;
- AutoSave uses a debounce window to avoid rapid-fire saves during clustered transactions;
- AutoSave logging distinguishes actual saves from debounce skips.

A missing save event for every individual commit is not automatically a defect when the debounce rule intentionally coalesces clustered commits.

### 4.7 Journaling, Backup, and Recovery

The desktop runtime has three local resilience layers:

1. daily text logs through `modLogError` / lifecycle logging;
2. per-table append journals plus startup backups through `modJournal`;
3. transaction rollback through `_TX` wrappers.

`modJournal` canonical capabilities:

- `WriteJournalRow()` is called from `AppendRow()`;
- every successful append writes a semicolon-separated CSV record under `Journal/`;
- a new journal file starts with `JournalTime` plus live table headers from `GetTableHeaders(tblName)`;
- `BackupFileOnStart()` creates a timestamped workbook copy under `Backup/`;
- `CheckJournalForRecovery()` compares today's journal line count against current table row count and emits advisory crash-loss warnings;
- `PurgeOldJournals()` and `PurgeOldBackups()` prune old files according to the active retention policy;
- log/journal/backup maintenance is best-effort and must not block normal workbook use.

Recovery is advisory unless a domain module has an explicit recovery state machine such as SEF. Journal warnings tell the operator/support team where to inspect; they are not a full automatic replay engine.

### 4.8 ProductionHealthCheck

`RunProductionHealthCheck` is the final workbook data-cleanliness launch gate. A code build may pass compile/smoke checks, but a specific workbook is production-launch-ready only when production health has no failures.

Production-health expectations:

- active broken references must be cleaned, repaired or stornirano-marked;
- smoke/regression fixtures must not leave active broken production data;
- legacy demo/test rows must not pollute final launch output;
- final production workbook launch requires `RunProductionHealthCheck` failure count equal to zero.

### 4.8.1 Duplicate-Key Preflight in ProductionHealthCheck

`modProductionHealthCheck` includes read-only duplicate-key preflight checks before operator workflows are considered safe.

The check runs early after core schema checks and must detect duplicates for:

```text
OtkupID
OtpremnicaID
ZbirnaID
PrijemnicaID
FakturaID
NovacID
BankaImportID
ParcelaID
```

Duplicate keys in these identity columns are blocking health failures because they can cause exact-row guards, document-chain linking, finance allocation, BankaMapiranje and geo/master-data workflows to target an ambiguous row.

The preflight check is diagnostic/read-only. It must not auto-delete, merge or repair duplicate rows without explicit operator/admin action.


### 4.9 Shared Helper Modules

Current desktop helper ownership:

| Module | Canonical responsibility |
|---|---|
| `modDataAccess` | workbook table access, arrays, append/update/search, ID generation, duplicate checks, journal hook |
| `modSchemaGuard` | fail-fast schema/update helpers: `RequireColumnIndex`, `RequireColumns`, `RequireUpdateCell` |
| `modParse` | canonical parsing/normalization of user-entered numbers, integers and dates |
| `modComboBinding` | MSForms ComboBox binding with human display + hidden stable ID |
| `modArrayUtils` | in-memory filtering, sorting, grouping and summing for report-like use cases |
| `modHttpUtils` | canonical outbound `UrlEncode` and `JsonEscape` helpers |
| `modJournal` | append journals, startup backup and advisory crash-recovery checks |
| `modLogError` | daily local text logging and runtime diagnostics |
| `modMonitoring` | best-effort remote monitoring client |

Rules:

- forms must use `modParse` instead of raw `Val`, `CDbl`, `CLng` or duplicated private parsers for business input;
- entity ComboBoxes must store stable IDs in hidden columns and saves must read `GetComboID()`;
- in-memory report utilities replace worksheet copy/sort/group side effects;
- new HTTP clients must not introduce private ANSI/codepage URL encoders.

### 4.10 UI / Business Boundary

Desktop forms are operator surfaces. Business/data modules own validation, persistence and domain rules.

Canonical boundary:

- forms may show `MsgBox`, toasts/status labels and operator confirmations;
- business modules must raise errors, return result objects/IDs/Boolean failure or log details instead of using `MsgBox` as control flow;
- activation handlers use guarded error handling and must not crash the UI lifecycle;
- destructive actions require explicit operator confirmation at the form/operator-shell layer;
- `frmOtkupAPP` is the main operator shell and must not embed domain writes outside explicit button/action handlers;
- `frmSplash` performs no business reads/writes beyond UI presentation and handoff;
- Banka inbox import is triggered from the Banka navigation/review workflow, not from app startup;
- form-level close and workbook-level close must delegate to the centralized shutdown path.


### 4.11 v6.24 VBA Document Numbering Model

v6.24 documents the current document-numbering closeout from the frontend/VBA integration workstream.

Canonical VBA document numbering is no longer treated as manual/inconsistent UI state. The current model uses business document numbers aligned with the v6.15 `x/ddmmyy[-rb]` convention:

- `BrojDokumenta`;
- `BrojOtpremnice`;
- `BrojZbirne`.

The numbering model is lock-based per station/stanica where the number is generated or reserved during guarded sync/document flow. The purpose is to avoid duplicate business numbers during mixed VBA/PWA operation.

Current invariant:

```text
Business document numbers must be generated through canonical numbering helpers / guarded flows.
Manual or ad-hoc numbering in UI code is not canonical.
```

The v6.24 source summary reports that this was implemented across nine files. The final code-level file list must be confirmed from the actual AgriX repository before treating this as independent Git-verified implementation evidence.

## 5. Data Architecture

This section states the current data architecture and source-of-truth rules for canonical tables, transport sheets and derived views.

This section is the current-state data map for the system. It defines canonical stores, ownership, table groups, soft-delete expectations, transport projections and schema guard rules.

Architecture rule:

```text
Canonical source facts live in the owning system/table.
Transport sheets, exports, caches and dashboards are read models unless explicitly documented as writable source-of-truth surfaces.
```

### 5.1 Data Ownership Principles

| Principle | Current rule |
|---|---|
| Canonical owner wins | The domain owner listed for a table or sheet is the only layer allowed to create authoritative source facts for that domain. |
| Derived views do not write back | `MgmtReports`, `SaldoOMDetail`, `Kartice`, PWA caches and monitoring read tabs are derived/exported views, not correction surfaces. |
| Soft delete beats physical delete | Business rows are normally marked with `Stornirano`, `Aktivan`, `Aktivna` or equivalent status flags rather than removed. |
| Transport is explicit | Google Sheets role tabs carry operational sync state and are reconciled into desktop master tables through import/export contracts. |
| Schema failure is safer than silent default | Required columns must fail fast through `RequireColumnIndex` / `RequireColumns`; optional columns must be explicitly optional. |
| External writes are not rollback-able | Google writeback, file moves, monitoring sends and SEF HTTP calls must be documented as external side effects, not local transaction participants. |

### 5.2 Canonical Entities

| Entity / Table | Purpose | Primary key | Canonical owner | Soft delete / visibility | Notes |
|---|---|---|---|---|---|
| `tblKooperanti` | kooperant master data | `KooperantID` | Operator / desktop | `Aktivan` controls visibility | exported to PWA Stammdaten |
| `tblStanice` | station master data | `StanicaID` | Operator / desktop | `Aktivan` only | no active `Stornirano` model |
| `tblVozaci` | driver master data | `VozacID` | Operator / desktop | `Aktivan` only | includes capacity data such as `KapacitetKG` |
| `tblKupci` | buyer master data | `KupacID` | Operator / desktop | active-status based | exported to PWA / management views |
| `tblKulture` | fruit/sort catalog | `KulturaID` | Operator / desktop | n/a | lookup/catalog table |
| `tblArtikli` | agrohemija item catalog | `ArtikalID` | Operator / desktop | active flag | private fiscal receipts do not automatically become master articles |
| `tblParcele` | parcel master + geo/meteo flags | `ParcelaID` | Operator / desktop | `Aktivna` / active flag | enriched with point/polygon/meteo-facing fields |
| `tblOtkup` | field procurement records | `OtkupID` | Otkupac operationally, desktop after import | `Stornirano` | start of document chain |
| `tblOtpremnica` | shipment/transport rows | `OtpremnicaID` | Operator / desktop | `Stornirano` | may exist before `BrojZbirne` is known |
| `tblZbirna` | aggregated transport document | `ZbirnaID` | Vozac flow + desktop import | `Stornirano` | business number is `BrojZbirne` |
| `tblPrijemnica` | cold-storage receiving document | `PrijemnicaID` | Operator / desktop | `Stornirano` | row-unique ID; `BrojPrijemnice` groups class rows |
| `tblFakture` | invoice header | `FakturaID` | Operator / desktop | `Stornirano` | includes SEF workflow state |
| `tblFakturaStavke` | invoice lines | `StavkaID` | Operator / desktop | `Stornirano` / orphanable | relink workflows may repair document-chain references |
| `tblNovac` | money movement | `NovacID` | Operator / desktop | `Stornirano` | manual finance, avans and bank-mapping outputs |
| `tblAmbalaza` | packaging movement ledger | `AmbalazaID` | Operator / desktop | `Stornirano` | strict `Ulaz` / `Izlaz` semantics |
| `tblMagacin` | agrohemija warehouse movements | `MagacinID` | Operator / desktop | `Stornirano` | `MAG_ULAZ` and `MAG_IZLAZ` journal |
| `tblBankaImport` | bank statement staging | `BankaImportID` | Operator/system via import | `Stornirano` | staging only; source bank facts are not overwritten |
| `tblPartnerMap` | learned bank partner mapping | logical composite | Operator/system | n/a | helper table, not a business ledger |
| `tblSEFSubmission` | SEF submission journal | `SEFSubmissionID` | Operator/system | `Stornirano` / status | request/response persistence |
| `tblSEFEventLog` | SEF event timeline | `SEFEventID` | Operator/system | status/audit | append-style event stream |
| `tblConfig` | shared app / Google / PWA config | `Parameter` | Operator / desktop | n/a | must not store local workstation paths |
| `tblSEFConfig` | SEF and monitoring config | `ConfigKey` | Operator / desktop | n/a | includes SEF, seller and monitoring runtime config |
| `tblLocalConfig` | workstation-local setup config | `Kljuc` | Local workstation setup | n/a | owns paths such as `PDFTOTEXT_EXE_PATH` |

### 5.3 Master Tables

Master tables are maintained in the desktop workbook and exported to Google/PWA read models when needed.

Canonical master-data tables:

```text
tblKooperanti
tblStanice
tblVozaci
tblKupci
tblKulture
tblArtikli
tblParcele
tblConfig
tblSEFConfig
tblLocalConfig
```

Current rules:

- Operator/desktop is the canonical writer for master data.
- PWA must treat exported Stammdaten as read-only unless a specific endpoint is documented as writable.
- Visibility is controlled by `Aktivan`, `Aktivna` or the domain-specific active flag; do not infer soft delete where the table explicitly uses active status instead.
- `tblLocalConfig` is machine-local and must not be exported as shared PWA/Google configuration.
- `tblConfig` remains shared application / Google / PWA configuration.
- `tblSEFConfig` remains SEF, seller and selected monitoring runtime configuration.

### 5.4 Transaction Tables

Transaction tables record source business events or operational movements.

| Table | Domain | Owner | Current rule |
|---|---|---|---|
| `tblOtkup` | field procurement | PWA operational input, desktop canonical after import | append/import first; later document links update checked fields |
| `tblNovac` | finance | desktop | finance movement facts, avans and bank-map outputs |
| `tblAmbalaza` | packaging ledger | desktop | ledger semantics; no negative quantity; strict direction |
| `tblMagacin` | agrohemija warehouse | desktop | movement journal; stock is derived from active movements |
| `tblBankaImport` | bank staging | desktop/system import | staging source facts; mapping changes status/links, not raw bank facts |
| role-specific Google sheets | field/PWA transport | PWA/GAS | operational sync queues before desktop import/export reconciliation |

Transaction table writes must be protected by the transaction model defined in the VBA architecture section unless the write is an external transport write owned by GAS/PWA.

### 5.5 Document Tables

Document tables are the canonical desktop document-chain backbone.

| Table | Primary identity | Business identity | Notes |
|---|---|---|---|
| `tblOtpremnica` | `OtpremnicaID` | document number / shipment context | may be created before `BrojZbirne` exists |
| `tblZbirna` | `ZbirnaID` | `BrojZbirne` | `BrojZbirne` is PWA-first with VBA fallback |
| `tblPrijemnica` | `PrijemnicaID` | `BrojPrijemnice` | `PrijemnicaID` is row-unique; `BrojPrijemnice` may group class rows |
| `tblFakture` | `FakturaID` | invoice number | SEF workflow fields live on/around invoice state |
| `tblFakturaStavke` | `StavkaID` | faktura line + linked prijemnica/class | may become orphaned during storno/relink workflows |
| `tblSEFSubmission` | `SEFSubmissionID` | SEF submission correlation | audit/persistence table |
| `tblSEFEventLog` | `SEFEventID` | SEF event correlation | append-style timeline |

Current rules:

- `PrijemnicaID` is the row identity used by faktura logic.
- `BrojPrijemnice + Klasa` is the relink identity for orphaned faktura stavke when class rows are recreated.
- `BrojZbirne` is a business document number and must not be confused with `ServerRecordID` or `ZbirnaID`.
- Document-chain status changes must not be done by report/read-model code.

### 5.6 Finance Tables

Canonical finance tables and read models:

| Table / View | Type | Owner | Notes |
|---|---|---|---|
| `tblNovac` | source fact | desktop | money movement ledger |
| `tblFakture` | source fact | desktop | invoice state and SEF workflow relationship |
| `tblFakturaStavke` | source fact | desktop | invoice-line facts and prijemnica references |
| `tblBankaImport` | staging source | desktop/system import | bank statement facts before mapping |
| `tblPartnerMap` | helper/config | desktop/system | learned mapping; not financial source fact |
| `Kartice` export | derived read model | desktop export | PWA/management/kooperant read view |
| `SaldoOMDetail` | derived read model | desktop export/report | management/finance read view |
| `MgmtReports` | derived read model | desktop export/report | management dashboard data |

Current rules:

- Finance truth is in desktop source tables, not exported report tabs.
- Bank mapping produces `tblNovac` and explicit links/status transitions; it does not delete or rewrite imported bank facts.
- Report exports must exclude `Stornirano` rows unless the report is explicitly audit/history.

### 5.7 BankaImport Tables

#### 5.7.1 `tblBankaImport`

**Purpose:** bank-import staging table for parsed PDF statement rows before reconciliation into `tblNovac`, faktura/otkup links, or skip/error state.  
**Primary key:** `BankaImportID` (`BIM-*`).  
**Written by:** `modBankaImport` and the Banka PDF parser pipeline.  
**Read by:** `frmBankaImport`, `modBankaMapiranje`, finance reports and audit/review flows.  
**Soft delete:** `Stornirano = "Da"`.

Canonical columns:

```text
BankaImportID | BrojDokumenta | DatumIzvoda | BrojRacuna | DatumTransakcije | Partner | PartnerKonto | Opis | Uplata | Isplata | Valuta | PozivNaBroj | SvrhaPlacanja | BankaReferenz | IzvorFajl | ImportVreme | Obradjeno | Stornirano | PocetnoStanje | ZavrsnoStanje | UkupanDuguje | UkupanPotrazuje
```

Current invariants:

- Rows are append-only import staging rows; business reconciliation is represented through `Obradjeno`, related `tblNovac` rows and downstream links, not by overwriting the source bank import record.
- `IzvorFajl` and `ImportVreme` are mandatory traceability metadata for imported statement rows.
- `Obradjeno` lifecycle is:

```text
"" -> Da | Skip | Error
```

- Only rows with `Stornirano <> "Da"` and `Obradjeno` neither `Da` nor `Skip` are eligible for mapping.
- Header values `BrojIzvoda`, `DatumIzvoda` and `BrojRacuna` are copied onto every staged transaction row from the same parsed PDF statement.
- Statement-level saldo metadata `PocetnoStanje`, `ZavrsnoStanje`, `UkupanDuguje` and `UkupanPotrazuje` is copied onto every staged row from the same statement.
- Duplicate detection first uses `BankaReferenz` when present. If no bank reference exists, duplicate detection falls back to the active composite key:

```text
BrojDokumenta + DatumTransakcije + Uplata + Isplata + Partner
```

- A parsed statement is rejected entirely before staging if required header/saldo data is missing or any statement integrity gate fails.

#### 5.7.2 `tblPartnerMap`

**Purpose:** learned bank-side partner-name to internal entity mapping.  
**Primary key:** logical composite by normalized `BankaName`.  
**Foreign keys:** `PartnerID`, `OMID`.  
**Written by:** successful bank mapping flows.  
**Read by:** `modBankaMapiranje` auto-map and preview helpers.  
**Soft delete:** n/a.

Canonical columns:

```text
BankaName | PartnerID | EntitetTip | OMID
```

Current invariants:

- Lookup is exact-match on normalized bank name, using trimmed and case-insensitive comparison.
- Duplicate save attempts are treated as no-op success.
- Learned mappings accelerate auto-match but do not override exact-row integrity guards.

### 5.8 Agrohemija Tables

| Table / Projection | Type | Owner | Current rule |
|---|---|---|---|
| `tblArtikli` | master catalog | desktop | canonical item list and packaging metadata |
| `tblMagacin` | warehouse movement source | desktop | all stock is derived from active movement rows |
| `MagacinKoop` / `magacinkoop` export | PWA read model | desktop export / Stammdaten | Kooperant local validation uses this exported stock state |
| `Izdavanje` sheet | GAS/management issue persistence | Management endpoint | current `saveIzdavanje` boundary; server idempotency remains roadmap |
| `TRETMAN-<KooperantID>` | treatment evidence/history | GAS/PWA | idempotent by `ClientRecordID` through `syncTretman` |

Current rules:

- `tblMagacin` is a movement ledger, not a mutable stock-total table.
- `MAG_IZLAZ` requires positive quantity, required kooperant and sufficient stock before append.
- PWA management and kooperant flows must store real item-unit quantity, not package count.
- Treatment sync does not automatically decrement server-side `magacinkoop` stock in the current contract.

### 5.9 Monitoring Tables

The monitoring workbook is a separate operational observability store. It is not a source-of-truth replacement for business tables.

Canonical monitoring tabs:

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

Current rules:

- `Events` is append-only for valid monitoring events.
- `Errors` receives `ERROR` / `CRITICAL` rows and redacted error context.
- `Health` is a current component-status snapshot maintained by events and watchdog checks.
- `SEFStatus` summarizes SEF operational status but does not replace `tblSEFSubmission` or `tblSEFEventLog`.
- `Backups`, `Alerts` and `AuditCritical` are operational views for support/manual-review workflows.
- Monitoring failure must never block a business transaction or change source data.

### 5.10 Google Sheets Transport Tables

Google Sheets stores shared operational transport state and exported read models.

| Sheet / Family | Type | Writer | Reader | Current rule |
|---|---|---|---|---|
| `Stammdaten` tabs | exported read model | desktop export | PWA/GAS | Excel master wins; export refresh replaces stale read model |
| `OTK-*` | role transport | Otkupac/GAS | desktop import, Otkupac, Management | PWA-created otkup records before desktop canonical import |
| `VOZ-*` | role transport | Vozac/GAS | desktop import, Vozac | `ServerRecordID` and `BrojZbirne` are distinct |
| `TRETMAN-*` | role transport/history | Kooperant/GAS | Kooperant/PWA | idempotent by `ClientRecordID` |
| `TROSKOVI-*` | role transport/history | Kooperant/GAS | Kooperant/PWA | idempotent by `ClientRecordID` |
| `FISKALNI-*` | private receipts | Kooperant/GAS | Kooperant/GAS parse/read | private kooperant-scoped data |
| `DispecerPlan` | planning state | Management/GAS | Management UI | planning source, not OTK write authority |
| `KamionStatus` | driver status | Vozac/Management via GAS | Management/dispatch | scoped authorization required |
| `MgmtReports` | derived read model | desktop/export jobs | Management | dashboard/read-only |
| `Kartice` | derived read model | desktop/export jobs | Kooperant/Management | read-only finance view |
| `MeteoLatest` | current meteo read model | scheduled GAS | PWA/GAS | latest overwrite for current state |
| `MeteoHistory` | meteo history | scheduled GAS | reports/risk views | append/history |
| `SyncControl` | lock/control state | MasterSync/GAS | PWA/GAS/desktop | `MASTER_SYNC_LOCK` and full-cycle sync coordination |
| `ErrorLog` | runtime error log | GAS/PWA/client bridge | support/monitoring | operational diagnostics |

### 5.11 Identity and Sync Fields

Identity fields must not be conflated across layers.

| Field | Meaning |
|---|---|
| `ClientRecordID` | stable client-side idempotency key generated by PWA/client |
| `ServerRecordID` | technical GAS/transport/server identity |
| `OtkupID`, `ZbirnaID`, etc. | desktop/master technical business-row IDs |
| `BrojZbirne`, `BrojPrijemnice`, invoice numbers | human/business document numbers |
| `SyncStatus` | transport lifecycle status, not source business truth by itself |
| `ReceivedAt`, `updatedAtServer`, `syncedAt` | sync timestamps / diagnostics |

Current rules:

- `BrojZbirne` must never be generated by copying `ServerRecordID`.
- Google VOZ writeback writes column B / `ServerRecordID` separately from column T / `BrojZbirne`.
- Duplicate detection and idempotency must use the documented stable identity for the relevant layer.

### 5.12 Derived / Exported Read Models

Derived/cache layers include:

```text
SaldoOMDetail
MgmtReports
Kartice
MeteoLatest
PWA role caches
IndexedDB queues/stores
management overview bundles from getMgmtAll
warehouse/agro derived views such as GetMagacinStanje(), ReportIzdavanjePoKooperantu(), ReportStanjePoDobavljacu()
```

Current rules:

- Derived views may be rebuilt from source tables.
- Derived views must not become hidden correction surfaces.
- PWA local pending/error records may dominate stale server reads for display/sync safety, but desktop source tables remain canonical after import/export reconciliation.
- Report/read-model modules must not silently mutate source tables while rendering.

### 5.13 Required Schema Guard Rules

Data architecture is only safe if schema drift fails visibly.

Current rules:

- Required source columns must use `RequireColumnIndex` or `RequireColumns` before business logic relies on them.
- Critical writes must use `RequireUpdateCell` or a documented exact-row update helper.
- Optional columns must be explicitly marked optional in code/docs and must have safe behavior when missing.
- Missing required table is a hard error for production import/save flows.
- Empty `GetNextID` result is a hard error for append flows.
- `AppendRow <= 0` is a hard error for production source-table writes.
- Filtered array indexes must not be used as row indexes for source-table updates.

### 5.14 Data Migration and Backfill Boundary

Current documented data-migration posture:

- v6.21 requires no production data migration.
- Current BankaImport imports require the four `tblBankaImport` saldo columns before new bank statements are accepted.
- Historical `tblBankaImport` rows may have blank saldo fields unless backfilled from original statements.
- Test rows, demo rows and disposable smoke data are not production migration targets unless explicitly promoted.
- Production readiness depends on `RunProductionHealthCheck`, compile/smoke gates and workbook data cleanliness, not only code correctness.

---

## 6. Document Flow Architecture

The canonical desktop document chain is:

```text
Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF
```

The chain is **lineage-preserving**, not physically cascading. Each document/table keeps its own identity and `Stornirano` state. Recovery is done through explicit storno/relink/repair flows, not by deleting historical rows.

Document-flow modules follow the global invariants from section 3:

- critical table/column reads are fail-fast;
- critical row updates are checked;
- `_TX` wrappers rollback hard failures;
- exact-row behavior is required for row-unique IDs;
- multi-row business numbers are allowed only where documented, such as `BrojZbirne` and `BrojPrijemnice` grouping;
- monitoring is emitted at transaction/business boundaries, not at every helper call.

### 6.1 Otkup

`tblOtkup` is the canonical start of the procurement/document chain.

Current ownership:

- PWA Otkupac can create local offline-first otkup records and sync them through the OTK sheet.
- Desktop operator can enter otkup directly through `frmOtkup`.
- Desktop MasterSync imports synced PWA otkup rows into `tblOtkup`.
- `modOtkup` owns desktop save/validation behavior.

Canonical table identity:

```text
Primary key: OtkupID
Important links: KooperantID, StanicaID, KulturaID, VozacID, OtpremnicaID, ParcelaID, BrojZbirne
Soft delete: Stornirano = "Da"
```

Current rules:

- `tblOtkup` is the raw procurement source for later otpremnica/zbirna/prijemnica/faktura traceability.
- `OtpremnicaID` may be blank temporarily. It is populated by guarded auto-link/manual-link flows when a unique otpremnica match is available.
- `BrojZbirne` may be populated later from VOZ import cascade or direct document flow.
- Desktop `frmOtkup` supports direct operator entry and remains a valid fallback/repair path even though PWA-first is the preferred field model.
- `SaveOtkupMulti_TX` is the canonical desktop multi-class/high-level save wrapper for one operator action that may create Klasa I, Klasa II, ambalaža, novac and avans side effects.
- Forms must not separately call `SaveOtkup_TX`, `SaveNovac_TX` and `ApplyAvansToOtkup_TX` for one logical operator action when the multi-wrapper owns the transaction.
- PWA sync failures remain retryable; desktop storno resets explicit packaging and money links only through the documented storno side effects.

Monitoring boundary:

- `SaveOtkup_TX` emits `OTKUP_SAVE_SUCCESS` / `OTKUP_SAVE_FAIL` and `Monitor_Error` where applicable.
- `SaveOtkupMulti_TX` emits `OTKUP_MULTI_SAVE_SUCCESS` / `OTKUP_MULTI_SAVE_FAIL` and `Monitor_Error` where applicable.

### 6.2 Otpremnica

`tblOtpremnica` is the transport/shipping document layer between procurement rows and aggregated driver/buyer shipment state.

Current ownership:

- Desktop `frmDokumenta` and `modDokumenta` own direct otpremnica entry.
- PWA/Otkupac otprema and MasterSync/import flows may produce or update transport context before desktop document finalization.

Canonical table identity:

```text
Primary key: OtpremnicaID
Important links: StanicaID, VozacID, BrojZbirne
Soft delete: Stornirano = "Da"
```

Current rules:

- `BrojZbirne` on otpremnica may be empty at initial creation because otpremnica can exist before the matching vozač zbirna is created/imported.
- Empty `BrojZbirne` is valid only as a pending-link state; the VOZ import cascade or manual repair can populate it later.
- Klasa I and Klasa II rows are saved atomically through `SaveOtpremnicaMulti_TX` when a dual-class operator action is performed.
- Klasa II transport rows are stored separately, with class-level aggregation rolling up through the shared business context.
- `SaveOtpremnica` preserves error context and fails if `GetNextID` does not return an `OtpremnicaID`.
- Active reads/report caches, including `GetVozacDokumenta` and `BuildZbirnaVrstaCache`, exclude stornirano otpremnica rows.
- Otpremnica storno is explicit and does not auto-delete downstream documents.

Monitoring boundary:

- Document-chain saves from `modDokumenta`, including `SaveOtpremnica_TX` and `SaveOtpremnicaMulti_TX`, are monitored fail-only through `DOKUMENT_SAVE_FAIL` / `Monitor_Error` to avoid noisy success streams.

### 6.3 Zbirna

`tblZbirna` is the aggregate transport/buyer document, typically created from the Vozač PWA flow or desktop `frmDokumenta` fallback.

Current ownership:

- PWA Vozač owns primary `BrojZbirne` generation at `confirmZbirna` time.
- GAS `processZbirnaRecord` persists the provided `brojZbirne` into the VOZ sheet.
- Desktop MasterSync imports VOZ rows into `tblZbirna`.
- VBA `GenerateBrojZbirne` remains a fallback for legacy rows and desktop/manual entry.

Canonical table identity:

```text
Primary key: ZbirnaID
Business number: BrojZbirne
Important links: VozacID, KupacID, ClientRecordID, SyncSource
Soft delete: Stornirano = "Da"
```

Current `BrojZbirne` rule:

```text
Format: x/ddmmyy[-rb]
Example: 4/040526 or 4/040526-2
```

Generation semantics:

- PWA extracts the numeric part from `VozacID`, combines it with the local Serbian business date, and appends `-2`, `-3`, etc. for subsequent zbirne by the same driver on the same day.
- The daily sequence counts existing same-driver/same-day zbirne including soft-deleted/stornirano rows. Storno does not reclaim sequence numbers.
- Driver can print zbirna before server/master sync because `brojZbirne` exists locally before upload.
- GAS must not invent `BrojZbirne` from `ServerRecordID`.
- Desktop import reads VOZ column 20 / `BrojZbirne` first, validates the canonical format and falls back to `GenerateBrojZbirne(vozacID, datum)` only when the column is empty.
- `ClientRecordID` is the idempotency/dedupe key for imported VOZ rows.
- VOZ writeback keeps Google column B for technical/master `ServerRecordID` / `ZbirnaID` and column T for business `BrojZbirne`.

Current flow rules:

- `tblZbirna` may span multiple rows by class under the same `BrojZbirne`.
- Desktop validation aggregates kg by class and ambalaža by combined document where applicable.
- Import cascade links `BrojZbirne` to matching `tblOtkup` and `tblOtpremnica` rows through `OtkupRecordIDs` / linkage context.
- Cascade-link updates are traceability-critical and must use checked/fail-fast update semantics.
MasterSync document-chain linking must use exact-row guards for critical import/link identities:

```text
Count = 0  => missing document link error
Count = 1  => allowed
Count > 1  => duplicate document key error
```

The rule applies to `OtkupID`, `OtpremnicaID`, `ZbirnaID` and `ClientRecordID` where `ClientRecordID` is used as the PWA-origin row identity.

Affected MasterSync flows include:

- `AutoCreateOtpremniceFromPWA`;
- `ImportVOZRow_RowTX`;
- `LinkZbirnaToOtkupAndOtpremnica`;
- `LinkOtkupToOtpremnicaStrict`;
- `LinkOtpremnicaToBrojZbirneStrict`.

Critical updates from these flows must use `RequireUpdateCell` or equivalent checked-write semantics. `FindRows(...).Count > 0` is not acceptable for these document-chain links.

VOZ row processing is handled by `ImportVOZRow_RowTX`, whose row-level transaction covers:

```text
ImportRowToTblZbirna
RequireSingle ZbirnaID
GetBrojZbirne strict
LinkZbirnaToOtkupAndOtpremnica strict
Commit/Rollback
```

Google status writeback remains outside local `clsTransaction` rollback semantics and must remain explicit/operator-diagnosable.
- Storno zbirna marks the active rows for the same `BrojZbirne` as stornirano but does not auto-storno dependent otpremnice/prijemnice.

Monitoring boundary:

- Zbirna document saves through `modDokumenta` are monitored fail-only as part of `DOKUMENT_SAVE_FAIL` / `Monitor_Error`.

### 6.4 Prijemnica

`tblPrijemnica` is the receiving/intake layer and the canonical invoiceable source for faktura creation.

Current ownership:

- Desktop operator currently owns prijemnica creation.
- Future hladnjača/PWA intake is roadmap, not active architecture.
- `modDokumenta` owns prijemnica save/relink behavior.

Canonical table identity:

```text
Primary key: PrijemnicaID
Business number: BrojPrijemnice
Important links: KupacID, BrojZbirne, VozacID, FakturaID
Soft delete: Stornirano = "Da"
```

Identity model:

- `PrijemnicaID` is unique per physical row.
- `BrojPrijemnice` is the business document number and may group multiple class rows.
- `BrojPrijemnice + Klasa` is the canonical relink identity for class-aware orphan repair.
- A composite `PrijemnicaID + Klasa` redesign is not required because `PrijemnicaID` is row-unique.

Current rules:

- `SavePrijemnica` uses the row returned by `AppendRow` rather than a follow-up `FindRows` lookup.
- `BuildPrijemnicaRowData` owns writing `COL_PRJ_KOL_AMB_VRACENA`; redundant post-append updates are not canonical.
- `SavePrijemnicaMulti_TX` saves dual-class rows atomically and includes ambalaža effects/relink work in the rollback scope.
- `RelinkFakturaStavke(newPrijemnicaID, brojPrijemnice, Optional klasaFilter)` is class-aware and relinks only the class currently being recreated.
- Relink updates the replacement prijemnica to `Fakturisano = "Da"`, sets `FakturaID`, and recomputes faktura status when a `FakturaID` exists.
- Shortage/manjak preview and validation functions support operator review before save and analytics after save.

Current reconciliation helpers:

- `ValidateZbirna(brojZbirne)` compares summed otpremnice kg/ambalaža against saved zbirna rows.
- `ValidateZbirnaPreUnosa(brojZbirne, inputKgKlI, inputKgKlII, inputAmb)` validates pending desktop input by class while keeping ambalaža aggregated.
- `CalculateManjak(brojZbirne)` returns saved shortage totals.
- `CalculateManjakPreview(...)` includes unsaved pending prijemnica quantities for UI warning before save.
- `CalculateManjakByOtpremnica(brojZbirne)` allocates shortage proportionally by otpremnica share.
- `CalculateProsekGajbe(...)` helpers provide average kg-per-crate diagnostics.

### 6.5 Faktura

`tblFakture` and `tblFakturaStavke` are the invoice header/line layers and the local source for SEF submission.

Current ownership:

- `modFaktura` owns invoice creation, status and print behavior.
- `frmFakturisanje` is the operator shell for selecting invoiceable prijemnice.
- SEF modules own remote submission/status state after local faktura creation.

Canonical identity:

```text
Faktura header primary key: FakturaID
Faktura line primary key: StavkaID
Line source: PrijemnicaID
Soft delete: Stornirano = "Da"
```

Current rules:

- Faktura creation is prijemnica-based.
- `CreateFaktura[_TX]()` receives selected prijemnica tuples and computes total from canonical `tblPrijemnica.Kolicina × tblPrijemnica.Cena`.
- New faktura rows start as unpaid/local-finalized before SEF submission.
- One `tblFakturaStavke` row is appended per selected prijemnica tuple and stores `PrijemnicaID`, `Kolicina`, `Cena`, `Klasa` and `BrojPrijemnice`.
- Source prijemnice are marked `Fakturisano = "Da"` and linked with `FakturaID`.
- Buyer avans may be auto-applied immediately after create through the finance/novac layer.
- `PrijemnicaID` remains the canonical line reference because workbook identity confirms one unique row per class row.
- If a source prijemnica is later stornirana, faktura/stavke may be marked orphaned through `OsirocenoOd` until replacement relink or faktura storno is completed.
- Faktura storno frees prijemnice and removes novac faktura links through explicit storno side effects.

Monitoring boundary:

- `CreateFaktura_TX` emits `FAKTURA_CREATE_SUCCESS`, `FAKTURA_CREATE_FAIL` and `Monitor_Error` where applicable.

Exact-row guard boundary:

- `PrintFaktura(fakturaID)` must require exactly one matching active `FakturaID` row.
- `UpdateFakturaStatus(fakturaID)` must require exactly one matching `FakturaID` row before recomputing payment/status state.
- `Count = 0` is a missing-faktura error.
- `Count > 1` is a duplicate-key error.
- These guards close the duplicate-`FakturaID` wrong-row risk for faktura print/status flows without changing the `PrijemnicaID` row-unique line model.
- Any remaining `CreateFaktura` source-prijemnica guard review is a separate hardening check and must not weaken the current `PrijemnicaID` model.

### 6.6 SEF Submission Flow

SEF is the electronic invoice submission/status layer for locally created fakture.

Current ownership:

- `modSEFClient`, `modSEFService`, `modSEFStatusSync` and related mapper/persistence helpers own SEF integration.
- `tblSEFSubmission` stores submission journal state.
- `tblSEFEventLog` stores the SEF event timeline.

Current rules:

- SEF operates after local faktura creation; it must not replace local faktura state as the accounting source of truth.
- `SEF_BASE_URL` must be HTTPS-only and plaintext `http://` is rejected locally.
- SEF submission/result persistence must use existing SEF persistence helpers and event log functions.
- Monitoring does not replace `tblSEFSubmission`, `tblSEFEventLog`, `AppendSEFEvent_Row`, `SaveSEFSubmissionResult_Row` or the SEF state machine.
- Startup recovery and status refresh are explicit SEF flows, not implicit document-chain cascades.
- Unknown/manual-review states remain operator-visible and audit-critical candidates.

### 6.7 Sledljivost

`modSledljivost` owns traceability repair/reporting around the canonical chain:

```text
Zbirna → Otpremnica → Otkup → Kooperant/Parcela
```

Current ownership:

- `AutoLinkOtkupOtpremnica_TX` owns guarded automatic linking from active otkup rows to otpremnica rows.
- `frmOtkupniBlokovi` is both a repair/audit surface and a trace/export surface.
- `TraceByZbirna()` walks the chain for reverse trace output.

Current rules:

- The canonical bridge from raw/PWA/manual otkup records into the document chain is `tblOtkup.OtpremnicaID`.
- Records without `OtpremnicaID` are not part of canonical trace output until auto-link or manual-link resolves them.
- Auto-link matching preserves the existing matching model and includes `BrojZbirne`/driver/station/date/class context where applicable.
- `AutoLinkOtkupOtpremnica` uses the canonical `COL_OTK_BROJ_ZBIRNE` constant, not hardcoded column strings.
- Updating `tblOtkup.OtpremnicaID` after an auto-link requires exactly one matching `OtkupID` row.
- `FindRows` returning `Nothing`, zero rows or multiple rows is a data-integrity error for that update.
- Link writes use `RequireUpdateCell`.
- Ambiguous matches remain unresolved for manual review; the system must not silently pick the first candidate.
- Trace PDF export writes a file and does not alter business tables.

Monitoring boundary:

- `AutoLinkOtkupOtpremnica_TX` emits best-effort success after commit.
- Failure path emits `Monitor_Error` and `SLEDLJIVOST_AUTOLINK_FAIL` before rollback.

### 6.8 Ambalaža Ledger

`tblAmbalaza` is the packaging movement journal used by document side effects and packaging saldo reports.

Current ownership:

- `modAmbalaza` owns packaging ledger writes and read-model semantics.
- Existing document flows snapshot `tblAmbalaza` in their outer transactions; `modAmbalaza` must not introduce nested transaction wrappers that conflict with caller-owned transactions.

Canonical table identity:

```text
Primary key: AmbalazaID
Important links: EntitetID, EntitetTip, VozacID, DokumentID, DokumentTip
Soft delete: Stornirano = "Da"
```

Current write rules:

- `TrackAmbalaza` validates non-negative quantity.
- Quantity `0` is a legal no-op.
- `tipAmb`, `entitetID` and `entitetTip` are required when quantity is positive.
- `Smer` accepts only `Ulaz` or `Izlaz`.
- Unknown `Smer` is an error, not an implied opposite direction.
- `GetNextID` returning empty is a fail-fast error.
- `AppendRow <= 0` is a fail-fast error.
- Schema reads use fail-fast column guards.

Current read/saldo rules:

- `GetAmbalazeStanje` treats `Ulaz` as `+Kolicina` and `Izlaz` as `-Kolicina`.
- `GetVozacAmbSaldo` treats driver balance as all active movements with matching `VozacID`; there is no canonical `DokumentTip` filter.
- Open-ended date filters evaluate `datumOd` and `datumDo` independently.
- Only non-stornirano rows participate in active saldo helpers.

Transaction boundary:

- Existing document save/storno flows own the transaction snapshot for `tblAmbalaza` when packaging side effects are part of a larger business operation.


### 6.9 Storno

`modStorno` is the canonical desktop business module for per-entity soft-delete and repair side effects. It is aligned with the hardening standard used by finance/document modules such as `modNovac` and `modFaktura`.

#### 6.9.1 Storno Scope

Canonical public/hardened surface:

```text
StornoOtkup_TX / StornoOtkup
StornoOtpremnica_TX / StornoOtpremnica
StornoZbirna_TX / StornoZbirna
StornoPrijemnica_TX / StornoPrijemnica
StornoFaktura_TX / StornoFaktura
StornoNovac_TX / StornoNovac
RequireStornoAllowed
CanStorno
LookupActiveID
```

#### 6.9.2 Storno Guard Rules

Canonical rules:

- ID-based storno operations require exactly one matching row.
- `Count = 0` is a missing-record error.
- `Count > 1` is a duplicate-key error.
- Critical schema reads use `RequireColumnIndex`.
- Critical writes use `RequireUpdateCell`.
- `_TX` wrappers rollback on any hard failure.
- `_TX` wrappers emit monitoring success/failure events best-effort.
- Business-layer hardening must not depend on `MsgBox` as control flow.
- `RequireStornoAllowed` is the preferred helper-pattern name for hard-fail eligibility checks before mutation.
- `CanStorno()` validates that the target exists and is not already stornirano before destructive soft-delete side effects run.
- `LookupActiveID` must not silently choose a first matching row when the operation requires one exact target.
- Silent `UpdateCell` is not allowed in storno business paths; checked writes use `RequireUpdateCell`.

#### 6.9.3 Storno Flow Boundary

Storno is a **soft-delete plus explicit repair** mechanism. It is not a physical delete and it is not an automatic chain-wide cascade.

Generic flow:

```text
operator selects target
  -> CanStorno / active-row validation
  -> exact-row lookup where required
  -> mark target Stornirano = "Da"
  -> run entity-specific side effects
  -> recompute affected statuses where required
  -> commit transaction
  -> emit monitoring best-effort
```

Every storno path owns only the side effects explicitly documented for that entity. Dependent documents are not automatically stornirano-marked unless the entity-specific contract says so.

#### 6.9.4 Entity-Specific Side Effects

`StornoOtkup`:

- marks the otkup row as stornirano;
- stornira related ambalaža rows;
- removes otkup links from `tblNovac`.

`StornoOtpremnica`:

- marks the otpremnica row as stornirano;
- stornira related ambalaža rows.

`StornoZbirna`:

- supports multiple active rows for the same `BrojZbirne`;
- marks active rows with that `BrojZbirne` as stornirano;
- does not reclaim the `BrojZbirne` sequence number;
- does not automatically cascade-storno dependent otpremnice/prijemnice unless a future explicit rule introduces that behavior.

`StornoPrijemnica`:

- marks the prijemnica row as stornirano;
- resets faktura linkage when applicable;
- marks orphaned faktura/stavka state when needed;
- stornira related ambalaža.

`StornoFaktura`:

- marks the faktura header as stornirano;
- updates faktura status;
- stornira faktura stavke;
- releases linked prijemnice;
- removes faktura links from `tblNovac`.

`StornoNovac`:

- marks the novac row as stornirano;
- recomputes faktura status when a faktura link exists.

#### 6.9.5 Storno Transaction and Recovery Contract

- Each public mutating storno operation has a `_TX` wrapper.
- The `_TX` wrapper snapshots all tables affected by that storno path.
- Any hard failure triggers rollback.
- Monitoring after success/failure is best-effort and must not hide rollback/failure semantics.
- Storno preserves document lineage and auditability; it does not physically remove historical records.
- Replacement/recreate workflows must rely on orphan/relink eligibility rules documented in the affected domain sections.

#### 6.9.6 Storno Writes

Depending on the entity, storno may write:

```text
tblOtkup
tblOtpremnica
tblZbirna
tblPrijemnica
tblFakture
tblFakturaStavke
tblNovac
tblAmbalaza
```

No storno path may update a table outside its documented side-effect scope without updating this reference and the relevant release gates.

---

## 7. Finance Architecture

This section describes the current desktop finance contract outside the BankaImport-specific details that are expanded in section 8. Historical v6.8/v6.21 hardening notes belong in `ARCHITECTURE_CHANGELOG.md`.

### 7.1 Novac

`modNovac` is the canonical desktop finance module for money movement rows, payment allocation, avans handling and status recompute.

Current contract:

- `tblNovac` is the canonical money ledger for buyer uplata, kooperant isplata, OM/station movement, avans and bank-mapped payments.
- `SaveNovac()` validates money direction and basic partner/entity context before append.
- A valid `tblNovac` row has exactly one positive amount direction: `Uplata` or `Isplata`.
- Negative amounts are invalid.
- `Tip` is required.
- `SaveNovac_TX()` snapshots `tblNovac`, `tblFakture` and `tblOtkup` before delegating to append/update behavior.
- Direct `AppendRow` writes are rollback-safe only when the affected table has been snapshotted before mutation.
- Stornirano money rows are excluded from live finance aggregates, open-item resolution and allocation logic.
- Update flows that rely on physical row indexes must use the full table array and skip stornirano rows manually; filtered arrays are acceptable for read-only aggregates.

`SaveNovac()` must hard-fail on ID generation or append failure. It must never silently return an empty `NovacID` when the row was not appended.

Canonical append rules:

```text
GetNextID = ""      => Err.Raise
AppendRow <= 0      => Err.Raise
AppendRow success   => return NovacID
```

This protects every downstream flow that depends on a valid `NOV-*` identity, including BankaMapiranje, faktura payments, avans allocation/split and otkup payment links.

Canonical finance tips and exact enum lists should remain in code/constants. The architecture rule is that tip classification determines downstream status, partner, faktura/otkup and saldo behavior.

### 7.2 Open Items and Status Recompute

Open-item helpers define remaining receivables/payables by outstanding amount, not by status field alone.

Current contract:

- `GetOpenFakture()` returns buyer invoices with an unpaid remainder after excluding stornirano money/faktura rows.
- `GetOpenOtkupi()` returns otkup rows with an unpaid remainder after excluding stornirano rows.
- `UpdateOtkupStatus()` is a two-way recompute:
  - linked active `Isplata >= Kolicina × Cena` sets `Isplaceno` and fills `DatumIsplate` if empty;
  - insufficient linked payment clears `Isplaceno` and `DatumIsplate`.
- `UpdateFakturaStatus()` recomputes faktura payment state from active uplata aggregation and must not mutate stornirana faktura rows.
- `PrintFaktura()` must block active printing of stornirana faktura.
- Status helpers must tolerate payment removal, storno and reset-link flows.

### 7.3 Avans

Avans allocation is a finance-critical split/link flow.

Canonical flows:

- `ApplyAvansToFaktura[_TX]()` allocates buyer avans to faktura.
- `ApplyAvansToOtkup[_TX]()` allocates kooperant/otkup avans to otkup.
- Full consumption may link the existing avans row.
- Partial consumption reduces the original avans row and creates a linked split row.
- Required updates use `RequireUpdateCell`.
- Split creation must check that `SaveNovac()` returned a valid `NOV-*` ID.
- `_TX` wrappers snapshot affected money/status tables and preserve the original error before rollback/logging.
- Avans allocation must exclude stornirano money rows.

### 7.4 Saldo / Kartice

Saldo and kartice are derived finance read models, not canonical transaction sources.

Canonical derived views:

| View | Source | Owner | Rule |
|---|---|---|---|
| `Kartica Kooperanta` | `tblNovac`, `tblOtkup`, kartice export | VBA export + PWA reader | `UKUPNO` rows are ignored by production parsing. |
| `SaldoOM` | novac + otkup + agro flows | VBA/export | Derived station-level open balance; not a write source. |
| `SaldoOMDetail` | station/OM detail export | VBA/export | Shared management read model. |
| `SaldoKupci` | fakture + novac | VBA/export | Depends on correct faktura/payment mapping. |
| Invoice payment status views | `tblFakture`, `tblNovac` | VBA | Must use stornirano-safe uplata aggregation. |
| Bank reconciliation candidates | `tblBankaImport`, `tblPartnerMap`, `tblNovac` | VBA | Staging/open queue view only; not a separate ledger. |

### 7.5 BankaImport

`modBankaImport` is the canonical desktop orchestration layer for bank statement inbox import and staging.

Active capabilities:

- `ImportBankaInbox_TX()` ensures configured inbox/processed/error folders exist, snapshots `tblBankaImport`, and wraps the inbox import in rollback discipline.
- `ImportBankaInbox()` enumerates all `*.pdf` files from `APP_BANKA_INBOX` and processes them through the per-file routine.
- `ImportOnePdfIntoBankaImport()` extracts PDF text, prepares/parses the statement, stages valid non-duplicate rows and returns an explicit outcome. It must not immediately move successful PDFs to `Processed`; successful moves are deferred until after transaction commit.
- `ParseBankaIzvodForImport()` requires `BrojIzvoda`, `DatumIzvoda`, `BrojRacuna` and a parseable `STANJE` saldo block. Missing required header/saldo data is a hard parse failure.
- `SaveBankaImportRows()` allocates new `BIM-*` IDs, writes `Valuta = "RSD"`, stamps `ImportVreme = Now`, initializes `Obradjeno` / `Stornirano` as blank, and hard-fails on missing schema, invalid data, empty `GetNextID` or `AppendRow <= 0`.
- `IsDuplicateBankaImport()` participates in the staging integrity flow and must use fail-fast schema guards.
- `GetUniqueTargetPath()` prevents overwrite collisions in processed/error folders by suffixing duplicate filenames before move.

### 7.6 BankaMapiranje

`modBankaMapiranje` is the canonical desktop reconciliation layer that converts staged bank rows into canonical money flow, partner-map learning and optional faktura/otkup links.

Active capabilities:

- `GetBankaImportOpen()` returns only active bank-import rows whose `Obradjeno` is neither `Da` nor `Skip`; stornirano rows are excluded.
- `ValidateBankaImportNotProcessed()` is the canonical guard for mapping/skip flows; already processed, skipped or stornirano rows are rejected before financial write.
- `tblBankaImport.Obradjeno` values are `""`, `Da`, `Skip` and `Error`.
- Public mutating bank-map flows use `_TX` wrappers. They snapshot at least `tblBankaImport` and `tblNovac`; buyer/kooperant flows additionally protect `tblFakture`, `tblOtkup` and/or `tblPartnerMap` depending on downstream side effects.
- Manual mapping families are `MapBankaImportAsKupac[_TX]()`, `MapBankaImportAsKooperant[_TX]()`, `MapBankaImportAsOM[_TX]()` and kooperant-block variants.
- `AutoMapBankaImportRow[_TX]()` routes clean-direction rows by amount polarity; `AutoMapAllBankaImport_TX()` applies the same logic over the open staging queue.
- Auto-map checks learned mappings in `tblPartnerMap`, then falls back to normalized-name resolution against `tblKupci`, `tblKooperanti` and `tblStanice`.
- Incoming buyer rows create `NOV_KUPCI_UPLATA` when a unique faktura is resolved; otherwise they create `NOV_KUPCI_AVANS`.
- Outgoing kooperant rows resolve to OM-linked `tblNovac` rows; direct otkup-linked payouts use `NOV_VIRMAN_FIRMA_KOOP`, and unmatched remainder/advance-only flows use `NOV_VIRMAN_AVANS_KOOP`.
- Incoming OM/station funding creates `NOV_KES_FIRMA_OTKUPAC` with `Partner = Naziv stanice`, `PartnerID = OMID`, `EntitetTip = "OM"` and `OMID = StanicaID`.
- `TryResolveFakturaForKupac()` attempts a unique hit by normalized `PozivNaBroj`, then invoice number found inside `SvrhaPlacanja`, and finally exact amount match. Only one unambiguous hit is authoritative.
- `MapBankaImportAsKooperantBlock*()` uses `PozivNaBroj` or a manual block number to find up to two open otkup candidates, sort by larger open amount first, write one or more `tblNovac` rows, link consumed rows to `OtkupID`, update otkup paid status and push any remaining excess into kooperant avans.
- Reconciled `tblNovac` rows store generated napomena containing `BIM:<id>` plus selected bank reference, konto, opis, svrha and match reason metadata.
- `SkipBankaImportRow[_TX]()` is the first-class operator defer path; it postpones a row without deleting or storniranje the imported source record.

### 7.7 Partner Mapping

`tblPartnerMap` stores learned mapping between normalized bank partner names and internal entities.

Canonical rules:

- `LookupPartnerMap()` reads learned exact-match bank partner names.
- `savePartnerMap()` may persist mapping learned from bank reconciliation.
- Re-saving an identical mapping is idempotent success.
- A conflicting mapping for the same bank name with different `PartnerID`, `EntitetTip` or `OMID` is a fail-fast data-integrity error.
- Partner mapping must not silently override a previous learned mapping.
- Mapping helpers use `RequireColumnIndex` for required columns.

### 7.8 OM Station Saldo and Agro Deduction

Station/OM saldo is a derived finance calculation used by reports, bank mapping and operator review.

Canonical rules:

- `GetOMAvansSaldo()` defines OM avans as cash sent from firm to station minus cash paid from station to kooperant.
- OM saldo helpers exclude stornirano rows.
- `GetAgroAbzug()` remains a finance/deduction helper owned outside `modAgrohemija`; `modAgrohemija` must not duplicate its logic.
- Required finance/report columns use `RequireColumnIndex`; optional columns must have explicit optional handling.

### 7.9 Finance Monitoring

Finance monitoring is emitted at transaction boundaries, not for every helper/read.

Canonical monitored events include:

- `NOVAC_SAVE_SUCCESS` / `NOVAC_SAVE_FAIL` from `SaveNovac_TX`;
- `AVANS_APPLY_TO_FAKTURA_FAIL` and structured `Monitor_Error` on allocation failure;
- bank mapping events such as `BANKA_MAP_SUCCESS`, `BANKA_MAP_FAIL`, `BANKA_IMPORT_SKIP`, `BANKA_AUTOMAP_ALL_START`, `BANKA_AUTOMAP_ALL_SUMMARY` and `BANKA_AUTOMAP_ALL_FAIL`;
- faktura status/SEF monitoring described in document-flow and monitoring sections.

Monitoring is best-effort and must not turn a committed finance save into a false business failure.

### 7.10 Finance Acceptance Gate Summary

The detailed checklist lives in `RELEASE_GATES.md`. The current required coverage includes:

- invalid both-direction money rows are rejected;
- valid uplata and valid isplata append correctly;
- stornirano novac rows are excluded from live aggregates;
- partner-map conflict is blocked;
- partial buyer avans split works;
- partial otkup avans split works;
- reset-link recomputes otkup status;
- required finance column lookup uses fail-fast guards.

---

## 8. BankaImport and BankaMapiranje Current Contract

This section is the canonical current-state contract for Banka PDF import, staging and reconciliation. Historical introduction belongs in `ARCHITECTURE_CHANGELOG.md`.

### 8.1 Banka Import Pipeline

Canonical flow:

```text
APP_BANKA_INBOX PDFs
  -> pdftotext extraction
  -> statement header parse
  -> transaction block parse
  -> STANJE saldo parse
  -> statement integrity gates
  -> non-duplicate row staging
  -> transaction commit
  -> deferred successful file moves
  -> reconciliation through frmBankaImport / modBankaMapiranje
```

`ImportBankaInbox_TX()` is the transaction boundary for staging. It snapshots `tblBankaImport`, processes selected/inbox PDF files, commits only after successful staging, and performs successful file moves after commit.

### 8.2 PDF Parser Contract

`modBankaImportParserPdfToText.ExtractTextFromPdf` / the canonical Banka PDF parser layer shells out to local Poppler `pdftotext.exe` with UTF-8 text extraction semantics:

```text
pdftotext -raw -nopgbrk -enc UTF-8
```

Parser responsibilities:

- normalize page/line breaks and spaces;
- extract mandatory statement header fields `BrojIzvoda`, `DatumIzvoda`, `BrojRacuna`;
- collect transaction blocks from numeric sequence markers until hard-stop summary markers;
- parse execution date, partner, account, zaduženje/odobrenje, šifra, svrha, `PozivNaBroj` and `Referenca`;
- clean trailing summary noise such as `Ukupno za...`, dangling dates or dangling references;
- produce fixed 10-column in-memory transaction rows consumed by `ParseBankaIzvodForImport()`.

Required parser helper surface:

```text
ResolvePdfToTextExePath
BuildUniquePdfTextTempPath
GetBaseFileNameNoExt
QuoteArg
DeleteFileIfExists
```

`DeleteFileIfExists` must exist in `modBankaImportParserPdfToText.bas` or another shared module visible to the parser before compile.

### 8.3 Local `PDFTOTEXT_EXE_PATH` Contract

The parser executable path is workstation-local configuration.

Canonical rules:

- `PDFTOTEXT_EXE_PATH` is read from `tblLocalConfig` through public `GetLocalConfigValue` from `modSetup`.
- `tblLocalConfig` belongs to local workstation setup state.
- `tblConfig` remains Google/PWA config and must not become the local workstation config store.
- `tblSEFConfig` remains SEF/monitoring-related config.
- No user-specific hardcoded parser path is allowed.
- Setup fallback may derive from `APP_ROOT_PATH`, for example:

```text
C:\OtkupApp\Tools\poppler\Library\bin\pdftotext.exe
```

Recommended local config row:

```text
Kljuc: PDFTOTEXT_EXE_PATH
Vrednost: C:\OtkupApp\Tools\poppler\Library\bin\pdftotext.exe
Opis: Putanja do pdftotext.exe za PDF bankarske izvode
```

### 8.4 Unique Temp File and Extraction-Failure Rule

Canonical rules:

- No static `%TEMP%\pdf_extract.txt` output file is allowed.
- Every PDF extraction creates a unique temp txt path.
- The temp txt path is deleted defensively before extraction and after extraction.
- `WScript.Shell.Run` exit code is captured.
- Any non-zero `pdftotext` exit code is a hard error.
- Missing output temp txt after a zero exit code is a hard error.
- Failed extraction enters the import error path and must not be treated as processed.

### 8.5 Statement Header and Saldo Parse Contract

`ParseBankaIzvodForImport()` requires:

```text
BrojIzvoda
DatumIzvoda
BrojRacuna
PocetnoStanje
UkupanDuguje
UkupanPotrazuje
ZavrsnoStanje
BrojNalogaZaduzenje
BrojNalogaOdobrenje
```

`ExtractIzvodSaldoPdfText()` extracts the `STANJE` block anchored by `Prethodno stanje` and parses the statement-level totals into the Banka saldo data structure.

The parser must handle live token-count variants produced by `pdftotext` when zero values in `Duguje` or `Potrazuje` collapse. The accepted variants are 4-token, 5-token and 6-token forms, disambiguated by `BrojNaloga` integer anchors at the end of the line.

Missing mandatory header/saldo fields is a hard parse failure before staging.

### 8.6 Statement Saldo Integrity Gates

Before any transaction row is staged, the parser/import flow enforces four statement-level gates:

1. `PocetnoStanje + sum(parsed Uplata) - sum(parsed Isplata) = ZavrsnoStanje ±0.01`.
2. `sum(parsed Uplata) = UkupanPotrazuje ±0.01`.
3. `sum(parsed Isplata) = UkupanDuguje ±0.01`.
4. Parsed uplata/isplata counts match `BrojNalogaOdobrenje` / `BrojNalogaZaduzenje`.

Failure behavior:

```text
No transaction row from the statement is staged.
The source PDF goes to APP_BANKA_ERROR.
The tblBankaImport transaction wrapper rolls back.
```

This prevents partially parsed bank statements from entering mapping, partner-map learning, `tblNovac` creation or document linking.

### 8.7 Fail-Fast Staging Rule

`SaveBankaImportRows` is a financial staging function and must fail fast.

Canonical rules:

- all `tblBankaImport` columns used by staging are read through `RequireColumnIndex`;
- missing `tblBankaImport` is a hard error;
- missing required column is a hard error;
- invalid input data array is a hard error;
- `GetNextID` returning an empty ID is a hard error;
- `AppendRow <= 0` is a hard error;
- failures bubble to the surrounding transaction wrapper;
- `IsDuplicateBankaImport` also uses `RequireColumnIndex` because it participates in the staging integrity flow.

### 8.8 Import Outcome Categories

Banka import must distinguish operational outcomes explicitly:

```text
imported
duplicate-only
parse error
integrity error
append error
schema error
extract error
unknown error
```

`imported` and `duplicate-only` are the only reliable success outcomes.

`parse error`, `integrity error`, `append error`, `schema error`, `extract error` and `unknown error` are failure categories and must not lead to a processed-file success path.

### 8.9 Deferred File Move Rule

DB staging must commit before any successful PDF is moved to `Processed`.

`ImportOnePdfIntoBankaImport` must not move the PDF immediately. It must parse, validate and stage the data, return a status and add successful files to a pending move list.

`ImportBankaInbox_TX` sequence:

```text
1. create pendingMoves
2. begin transaction
3. stage all selected/inbox PDF files
4. commit transaction
5. execute pending successful file moves to Processed
```

Failure behavior before commit:

```text
tblBankaImport rollback runs.
No successful PDF is moved to Processed.
The failed/problematic PDF remains in Inbox for operator review or retry.
```

Failure behavior after commit:

```text
DB rollback is no longer possible.
File move failure must be reported clearly for manual folder recovery.
```

This is the canonical trade-off because file-system moves are not transactionally rollback-able.

### 8.10 Exact-Row Mapping Guard

`modBankaMapiranje` must not link by "first row found".

For critical IDs, exact-row behavior is required:

```text
NovacID       -> Count must be 1
BankaImportID -> Count must be 1
OtkupID       -> Count must be 1
FakturaID     -> Count must be 1
```

Canonical failure rules:

```text
Count = 0  -> missing link error
Count > 1  -> duplicate key error
```

The local helper accepted for the current implementation is:

```vba
Private Function RequireSingleRow(ByVal tblName As String, _
                                  ByVal idColumn As String, _
                                  ByVal idValue As String, _
                                  ByVal sourceName As String) As Long
```

Long-term consolidation into a shared data-access guard module is roadmap, not a blocker for the current contract.

`UpdateBankaImportStatus` must use:

```text
RequireSingleRow
RequireUpdateCell
```

Manual mapping and auto-mapping use the same integrity standard.

### 8.11 Strict Novac-to-Otkup Link Pattern

`LinkNovacToOtkupStrict` is the canonical pattern for linking a newly created `NovacID` to an `OtkupID`:

```text
RequireSingleRow(tblNovac, NovacID)
RequireSingleRow(tblOtkup, OtkupID)
RequireUpdateCell(tblNovac, COL_NOV_OTKUP_ID)
UpdateOtkupStatus
```

### 8.12 `GetBankaImportRowByID` Legacy 1x10 Shape

`GetBankaImportRowByID` must keep the old semantic 1x10 result shape. It must not return the raw table row.

Canonical public contract:

```vba
result(1, 1)  = BrojDokumenta
result(1, 2)  = DatumTransakcije
result(1, 3)  = Partner
result(1, 4)  = PartnerKonto
result(1, 5)  = Uplata
result(1, 6)  = Isplata
result(1, 7)  = Opis
result(1, 8)  = SvrhaPlacanja
result(1, 9)  = BankaReferenz
result(1, 10) = PozivNaBroj
```

Implementation may and should use `RequireSingleRow` and `RequireColumnIndex` internally, but callers that read `bim(1, 1)` through `bim(1, 10)` must receive the same business fields.

Changing this contract to return the raw `tblBankaImport` row is a P0/P1 regression.

### 8.13 Stornirano BankaImport Guard

`ValidateBankaImportNotProcessed` must check both processing state and soft-delete state.

Canonical checks:

```text
COL_BIM_OBRADJENO
COL_BIM_STORNIRANO
```

If `COL_BIM_STORNIRANO = "DA"`, the row is not eligible for mapping or skip flow.

### 8.14 Kooperant Block Allocation Invariant

After creating a `Novac` row for an otkup candidate, `MapBankaImportAsKooperantBlockCore` must enforce:

- empty `novID` is a hard error;
- `LinkNovacToOtkupStrict` is called exactly once;
- `MapBankaImportAsKooperantBlockCore` count is incremented exactly once;
- `preostaloZaRaspodelu` is reduced exactly once.

Canonical block:

```vba
If Len(Trim$(novID)) = 0 Then
    Err.Raise ERR_BMAP_BASE + 40, "MapBankaImportAsKooperantBlockCore", _
              "SaveNovac nije vratio NovacID za OtkupID=" & otkupID
End If

LinkNovacToOtkupStrict novID, otkupID, _
                        "MapBankaImportAsKooperantBlockCore"

MapBankaImportAsKooperantBlockCore = MapBankaImportAsKooperantBlockCore + 1
preostaloZaRaspodelu = preostaloZaRaspodelu - iznosZaRed
```

The invalid pattern of nested `If novID <> "" Then` and duplicate count/balance updates is not allowed.

### 8.15 Banka Review Form Boundary

`frmBankaImport` is a review/orchestration shell, not a business-rule owner.

Current form responsibilities:

- `LoadBankaRows()` reads `GetBankaImportOpen()` and shows only non-stornirano rows that are not processed or skipped.
- Row selection auto-suggests `Kupac` mapping for uplata-only rows and `Kooperant` mapping for isplata-only rows.
- `LoadManualTargets()` fills target combos from kupci, kooperanti or stanice based on `MapTip`.
- Kooperant mapping can load distinct open `BrojDok` block candidates from active `tblOtkup` rows.
- Preview UI calls shared `modBankaMapiranje` helpers directly, including `GetBankaImportRowByID`, `TryResolveKupacBIM`, `TryResolveKooperantBIM`, `TryResolveOMBIM`, `TryResolveFakturaForKupac`, `GetOtkupCandidatesForKooperantBlock`, `NormalizeLooseBIM`, `NzBIM` and `GetKooperantNaziv`.
- Auto-map controls call `AutoMapBankaImportRow_TX()` or `AutoMapAllBankaImport_TX()`.
- Manual commit controls call the corresponding `_TX` wrappers and then reload the queue.

### 8.16 Banka Monitoring Boundary

`modBankaMapiranje` is monitored at TX-wrapper level only. Base mapping helpers remain uninstrumented to avoid duplicate events.

Canonical bank events:

```text
BANKA_MAP_SUCCESS
BANKA_MAP_FAIL
BANKA_IMPORT_SKIP
BANKA_AUTOMAP_ALL_START
BANKA_AUTOMAP_ALL_SUMMARY
BANKA_AUTOMAP_ALL_FAIL
```

Covered bank TX wrappers:

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

Bank monitoring payloads may use `bankaImportID`, `partnerType`, `partnerId`, `resultId`, `linkedEntityId` and batch counts. They must not include sensitive bank account details beyond operational identifiers needed for diagnosis.

### 8.17 BankaImport Acceptance Gates Summary

Detailed executable gates live in `RELEASE_GATES.md`. Current AR only requires that parser, staging, deferred-file-move, exact-row mapping, stornirano guard and legacy-shape contract gates pass before launch.

---

## 9. GAS API Architecture

This section states the current GAS API architecture and endpoint contract.

This section describes the current Google Apps Script API as a production contract. Historical details about when routes were introduced or hardened belong in `ARCHITECTURE_CHANGELOG.md`.

### 9.1 Role of GAS

Google Apps Script is the online transport and API layer between the PWA, Google Sheets and selected external services.

Canonical responsibilities:

- route all frontend/API requests by explicit `action` values;
- authenticate users and issue short-lived session tokens;
- enforce role/entity authorization for reads and writes;
- write PWA sync records to role-specific Google Sheets;
- expose exported/read-model data to PWA roles;
- own dispatch, meteo, fiscal and selected management write endpoints;
- receive PWA client-error reports and monitoring events;
- protect write operations with `withLock(...)` where Sheets/Drive mutation occurs;
- return structured JSON responses instead of leaking raw Apps Script runtime errors.

GAS is not the formal finance/document source of truth. Formal finance, SEF, BankaImport, document-chain and desktop master state remain owned by VBA/Excel unless a section explicitly says otherwise.

### 9.2 Deployment and Router Contract

The active backend is a deployed Apps Script Web App.

Canonical runtime contract:

- `doPost(e)` is the primary frontend contract for authenticated actions and POST-first public/read bridges.
- `doGet(e)` remains a compatibility/read surface and health-check entrypoint.
- every request carries an explicit `action` string;
- the router delegates to action-specific handlers;
- all exits return through `jsonResponse(obj)` or an equivalent JSON envelope;
- top-level `doPost` / `doGet` catch blocks return structured failure payloads instead of throwing raw platform errors;
- missing/disabled features return explicit disabled responses rather than falling through to missing handlers.

The deployment model is per-client/per-firm, execute-as-owner, with the Web App URL stored in PWA/VBA configuration rather than hardcoded in feature code.

### 9.3 Workbook / Sheet Topology

GAS owns online transport workbooks and derived/shared state, not the Excel master workbook.

Canonical online storage families:

| Family | Purpose |
|---|---|
| `MASTER_FOLDER_ID` | Root Drive folder for the per-client GAS workbook family |
| `Stammdaten` | exported/shared master data and `SyncControl` |
| `OTK-*` | Otkupac otkup transport sheets |
| `VOZ-*` | Vozac zbirna transport sheets |
| `TRETMAN-<KooperantID>` | Kooperant treatment history/evidence |
| `TROSKOVI-<KooperantID>` | Kooperant farm expense records |
| `OPREMA-<KooperantID>` | Kooperant equipment records |
| `FISKALNI-<KooperantID>` | private kooperant fiscal receipt records |
| `GEO_SPREADSHEET_ID` | dedicated geo/meteo workbook with `Parcele`, `MeteoLatest`, `MeteoHistory` |
| `ErrorLog` | GAS/PWA runtime error log |
| Monitoring workbook | monitoring is described in section 15 |

Missing workbooks/sheets may be lazily created by controlled helpers such as `getOrCreateSheet(...)`, but existing sync sheet schema drift must not be silently repaired by appending missing columns.

### 9.4 Auth / Token Model

`login` is the explicit session entrypoint.

Canonical auth/session rules:

- `authenticateUser(username, pin)` reads the `Users` tab inside `Stammdaten`.
- successful login issues a random token and stores `{ entityID, role, created }` under `TOKEN_<token>`.
- token cache is backed by `CacheService.getScriptCache()` and mirrored into `PropertiesService.getScriptProperties()` as fallback.
- malformed, expired or invalid fallback token payloads must be rejected/deleted.
- cache misses may restore valid fallback tokens into cache.
- token purge removes expired/malformed `TOKEN_*` properties.
- `setupTokenPurgeTrigger()` provisions scheduled maintenance in `Europe/Belgrade`.
- failed login attempts are throttled per username.
- login attempts are written to `LoginLog`.
- PWA transport must not append `token=` into URL query strings; token/action data belongs in POST payloads or controlled request bodies.

### 9.5 Role and Entity Authorization

Authorization is part of the GAS contract, not a UI-only concern.

Canonical rules:

- management-only actions require explicit `Management` role checks;
- role-scoped writes require both role and entity ownership checks;
- Otkupac write payloads must match authenticated `OtkupacID`;
- Kooperant write payloads must match authenticated `KooperantID`;
- Vozac write payloads must match authenticated `VozacID`;
- Management override is allowed only where explicitly documented;
- failed authorization returns structured unauthorized/forbidden payloads with clear codes.

Local `Code.gs` helper patterns such as `isManagement`, `requireRole`, `requireEntity` and `forbiddenResponse` are the accepted centralization pattern for reducing endpoint authorization drift.

NEEDS REVIEW: The source documents conflict on the final deployed authorization state of `saveParcelPolygon`. Until verified against deployed code, AR keeps the conservative review marker already listed in sections 16 and 19.

### 9.6 Public / Pre-Auth Exceptions

The active architecture allows only explicit public/pre-auth exceptions.

Canonical exceptions:

| Action | Boundary |
|---|---|
| `login` | auth bootstrap only |
| `logClientError` | pre-auth best-effort observability; may enrich with token if valid |
| `getMasterSyncState` | public/read lock state for PWA soft-lock behavior |
| `getParcelGeo` | public/read bridge; acknowledged geo exposure |
| `getParcelMeteo` | public/read bridge; acknowledged meteo exposure |
| `getParcelMeteoLatest` | public/read bridge; acknowledged meteo exposure |
| `getAllMeteoLatest` | public/read bridge; acknowledged meteo exposure |
| `saveParcelPolygon` | NEEDS REVIEW: docs disagree whether this is still public/pre-auth or Management-only |

Public geo/meteo reads must remain read-only. No public write surface may be added without an explicit architecture decision and security review.

### 9.7 Current Endpoint Authorization Matrix

The table below is the maintained architecture view for active and known transitional actions. If deployed `Code.gs` differs, update this table and mark differences in `NEEDS REVIEW` before relying on the document for production handoff.

| Action | Public | Auth required | Roles | Entity ownership | Lock required | Purpose / notes |
|---|---:|---:|---|---|---:|---|
| `ping` | Yes | No | n/a | n/a | No | Lightweight health check. |
| `login` | Yes | No | n/a | username/PIN auth | Yes/controlled | Auth bootstrap; writes token state and `LoginLog`. |
| `logClientError` | Yes | No | optional token enrich | optional payload/entity fallback | Yes | Pre-auth best-effort client observability; payload must be redacted/truncated. |
| `monitorPublic` | Yes | Secret | VBA/monitoring source | monitoring secret | Yes | Public monitoring ingest guarded by `MONITORING_INGEST_SECRET`. |
| `monitor` | No | Yes | allowed authenticated roles | token context | Yes | Authenticated monitoring ingest. |
| `getMasterSyncState` | Yes | No | n/a | n/a | No | Public/read soft-lock state for PWA. |
| `getStammdaten` | No | Yes | all logged-in roles | role-specific filtering where applied | No | PWA bootstrap master-data projection. |
| `sync` | No | Yes | Otkupac, Management | Otkupac must match `OtkupacID`; Management override explicit | Yes | OTK batch sync; idempotent by `ClientRecordID`. |
| `getOtkupi` | No | Yes | Otkupac, Management | Otkupac scoped | No | Otkup readback. |
| `uploadPdf` | No | Yes | Otkupac, Management | Otkupac scoped unless Management | Yes | Drive upload for otkup-list PDF/file. |
| `saveOtkupniListPdf` | No | n/a | disabled | n/a | n/a | Must return `FEATURE_DISABLED` until real implementation exists. |
| `getVozacOtkupi` | No | Yes | Vozac, Management | Vozac must match `VozacID` | No | Driver assigned otkupi. |
| `syncZbirna` | No | Yes | Vozac, Management | Vozac must match `VozacID`; Management override explicit | Yes | VOZ zbirna sync; idempotent by `ClientRecordID`. |
| `getVozacZbirne` | No | Yes | Vozac, Management | Vozac scoped unless Management | No | Driver zbirna readback. |
| `syncTretman` | No | Yes | Kooperant, Management | Kooperant must match `KooperantID` | Yes | Treatment evidence/history sync. |
| `getTretmani` / `getTretmaniForKooperant` | No | Yes | Kooperant, Management | Kooperant scoped unless Management | No | Treatment readback. |
| `syncTrosak` | No | Yes | Kooperant, Management | Kooperant must match `KooperantID` | Yes | Active expense sync; must return batch response, never empty HTTP 200 body. |
| `getTroskovi` / `getTroskoviForKooperant` | No | Yes | Kooperant, Management | Kooperant scoped unless Management | No | Expense readback. |
| `syncOprema` | No | Yes | Kooperant, Management | Kooperant scoped unless Management | Yes | Equipment sync. |
| `getOprema` | No | Yes | Kooperant, Management | Kooperant scoped unless Management | No | Equipment readback. |
| `syncAgromere` | No | Yes | Kooperant, Management | Kooperant scoped unless Management | Yes | Compatibility/agromere sync surface where still present. |
| `getKartica` | No | Yes | Kooperant, Management | Kooperant scoped unless Management | No | Kooperant finance card read model. |
| `getKooperantProizvodnja` | No | Yes | Kooperant, Management | Kooperant scoped unless Management | No | Kooperant production read model. |
| `parseFiskalniImage` | No | Yes | Kooperant, Management | Kooperant must match `KooperantID` if supplied | No | Quota-sensitive parsing; no final fiscal row write. |
| `parseFiskalni` | No | Yes | Kooperant, Management | Kooperant scoped if supplied | No | Fiscal parse/verification surface. |
| `saveFiskalni` | No | Yes | Kooperant, Management | Kooperant must match `KooperantID` | Yes | Private fiscal receipt save. |
| `saveFiskalniMapiranje` | No | Yes | Management | n/a | Yes | Shared fiscal mapping write. |
| `createArtikal` | No | Yes | Management | n/a | Yes | Controlled master article creation. |
| `getMgmtAll` | No | Yes | Management | n/a | No | Management bootstrap/read bundle. |
| `getMgmtKartica` | No | Yes | Management | n/a | No | Management finance read. |
| `getMgmtOtkupiByStanica` | No | Yes | Management | n/a | No | Station otkup read. |
| `getMgmtSaldoOM` | No | Yes | Management | n/a | No | OM saldo report read. |
| `getMgmtSaldoKupci` | No | Yes | Management | n/a | No | Buyer saldo report read. |
| `getMgmtOtkupPoOM` | No | Yes | Management | n/a | No | Otkup report read. |
| `getMgmtPredatoPoKupcu` | No | Yes | Management | n/a | No | Delivered-by-buyer report read. |
| `getMgmtFakture` | No | Yes | Management | n/a | No | Faktura read model. |
| `getMgmtFakturaStavke` | No | Yes | Management | n/a | No | Faktura lines read model. |
| `getWarRoomDemand` | No | Yes | Management | n/a | No | Demand read. |
| `saveWarRoomDemand` | No | Yes | Management | n/a | Yes | Demand create. |
| `removeWarRoomDemand` | No | Yes | Management | n/a | Yes | Demand remove. |
| `updateDemandPrimljeno` | No | Yes | Management | n/a | Yes | Demand received update. |
| `getDispecer` | No | Yes | Management | n/a | No | Dispatch board read: today-only demand + active (non-`zavrseno`) plans. |
| `saveDispecer` | No | Yes | Management | n/a | Yes | Dispatch plan create; writes to `DispecerPlan` sheet only; must not assign `VozacID` to `OTK-*` rows. |
| `updateDispecer` | No | Yes | Management | n/a | Yes | Dispatch plan status update (`planned` → `u_toku` → `zavrseno`). |
| `removeDispecer` | No | Yes | Management | n/a | Yes | Dispatch plan remove. |
| `getVozacPlans` | No | Yes | Vozac | Vozac scoped to own `entityID` | No | Driver's active plans for today; reads `DispecerPlan` filtered by `VozacID` and date, excludes `zavrseno`. |
| `getKamionStatus` | No | Yes | Management, Vozac | Vozac reads own where applicable | No | Truck status read. |
| `updateKamionStatus` | No | Yes | Vozac, Management | Vozac forced to `tokenData.entityID`; Management any | Yes | Driver status upsert. |
| `saveIzdavanje` | No | Yes | Management | n/a | Yes | Agrohemija issuing write; retry idempotency remains roadmap. |
| `getParcelGeo` | Yes | No | public/read | n/a | No | Public/read geo bridge; acknowledged exposure. |
| `getParcelMeteo` | Yes | No | public/read | n/a | No | Cached-first meteo read. |
| `getParcelMeteoLatest` | Yes | No | public/read | n/a | No | Latest meteo by parcel. |
| `getAllMeteoLatest` | Yes | No | public/read | n/a | No | Meteo bundle read. |
| `scheduledMeteoFetch` | Trigger | n/a | scheduled | n/a | Yes | Writes `MeteoLatest` / `MeteoHistory`. |
| `saveParcelPolygon` | NEEDS REVIEW | NEEDS REVIEW | NEEDS REVIEW | NEEDS REVIEW | Yes if write-enabled | Source docs conflict: earlier current AR says public/pre-auth exception; v6.3 changelog says moved behind token and Management-only. Verify deployed `Code.gs`. |

Endpoint reconciliation rules:

- every Sheets/Drive write action must execute inside `withLock(...)`;
- read-only endpoints do not require locks unless they mutate logs/cache/state;
- disabled endpoints must return `FEATURE_DISABLED` and must not be listed as active route-health handlers;
- Management override is explicit, not assumed;
- Kooperant/Otkupac/Vozac writes must verify both role and entity ownership;
- public read exceptions are documented security exceptions and must not be extended casually;
- `saveParcelPolygon` remains unresolved until deployed backend code is checked.

### 9.8 Response and Failure Contract

Canonical response rules:

- every handler returns a structured JSON envelope;
- authentication failures return structured 401-like payloads;
- authorization failures return structured 403-like payloads;
- validation/schema failures return explicit error codes/messages;
- batch sync returns per-record `results` and aggregate counts;
- mixed batch sync failures return `PARTIAL_FAILURE`;
- all-failed batch sync returns `BATCH_FAILED`;
- idempotent duplicate/client replay may return success with `status = existing`;
- top-level catch blocks return `{ success: false, error: ... }` or equivalent controlled failures;
- PWA must treat HTTP 200 with empty/null JSON for batch sync as invalid `empty-response`.

### 9.9 Sync Endpoint Contract

All role sync endpoints follow the same general pattern:

1. authenticate token;
2. enforce role and entity ownership;
3. validate `records` as an array;
4. apply master-sync write blocking when active;
5. enter `withLock(...)` for Sheets/Drive mutation;
6. validate schema;
7. process each row with a stable client identity where available;
8. return normalized batch results.

Canonical sync processors:

| Processor | Target | Identity / idempotency |
|---|---|---|
| `processRecord(record, otkupacID)` | `OTK-<OtkupacID>` | trimmed `ClientRecordID` |
| `processZbirnaRecord(record, vozacID)` | `VOZ-<VozacID>` | trimmed `ClientRecordID` |
| `processTretmanRecord(record, kooperantID)` | `TRETMAN-<KooperantID>` | `ClientRecordID` |
| `processTrosakRecord(record, kooperantID)` | `TROSKOVI-<KooperantID>` | `ClientRecordID` |
| `processOpremaRecord(record, kooperantID)` | `OPREMA-<KooperantID>` | `ClientRecordID` |

Terminal/master/error states such as `Synced>Master`, `Duplicate` and `SyncError:*` must not be reset to ordinary `Synced` by idempotent PWA retry.

### 9.10 `syncTrosak` Endpoint

`syncTrosak` is an active GAS endpoint for Kooperant farm expenses.

Canonical contract:

- action: `syncTrosak`;
- allowed roles: `Kooperant`, `Management`;
- Kooperant callers must satisfy `tokenData.entityID === data.kooperantID`;
- `records` must be an array;
- processing is protected by `withLock(...)`;
- each row is handled by `processTrosakRecord(record, kooperantID)`;
- the endpoint must return `jsonResponse(withLock(function() { ... return buildBatchSyncResponse(results); }))` or equivalent behavior;
- HTTP 200 with empty JSON is invalid and treated by PWA as `empty-response`;
- idempotency is by `ClientRecordID`, updating an existing row instead of appending duplicates;
- `getTroskoviForKooperant(kooperantID)` reads `TROSKOVI-<KooperantID>` and returns normalized records scoped to that kooperant.

### 9.11 Master-Sync Readout and Write Blocking

GAS exposes master-sync lock state to PWA and blocks writes while the VBA full-cycle sync is active.

Canonical behavior:

- `getMasterSyncState` reads `Stammdaten / SyncControl` and is available as a public/read action so PWA can display lock state before attempting sync;
- `blockWriteIfMasterSyncActive(data.action)` or equivalent must run before write action dispatch;
- public reads and `login` may remain available during the lock;
- ordinary server writes return a clear soft-lock response such as `MASTER_SYNC_ACTIVE`;
- PWA treats this as pending/retry, not data loss;
- stale lock handling is allowed so a crashed VBA sync does not block field work indefinitely;
- operator-visible lock message/state remains in `SyncControl`.

Conceptual router order:

```js
const publicReadResponse = handlePublicRead(data);
if (publicReadResponse) return publicReadResponse;

if (data.action === 'login') {
  return jsonResponse(authenticateUser(data.username, data.pin));
}

const masterSyncWriteBlock = blockWriteIfMasterSyncActive(data.action);
if (masterSyncWriteBlock) return masterSyncWriteBlock;

// normal authenticated dispatch
```

Desktop full-sync orchestration must also treat PWA unlock failure as degraded/failed, not green success.

`modGoogleSyncOrchestrator` may rely on GAS/PWA lock TTL and stale-lock recovery for eventual recovery, but the operator-facing result must remain partial/degraded when the final unlock call fails.

Invariant:

```text
sync steps OK + unlock failed
=> full sync result False / partial
=> operator does not see success
=> monitoring event is emitted
=> PWA is expected to recover after lock TTL
```


### 9.12 Schema Drift and Sheet Guard Contract

`ensureSheetColumns(sheet, requiredColumns)` may create canonical headers only for empty sheets.

For existing sheets:

- missing required columns are `SCHEMA_DRIFT` failures;
- mismatched headers are `SCHEMA_DRIFT` failures;
- extra named columns must not be silently treated as harmless if they break canonical row mapping;
- sync processors must fail clearly instead of appending guessed columns;
- schema repair is manual/operator-controlled, not hidden in runtime write paths.

### 9.13 Management, Dispatch and Agro-Izdavanje Endpoints

Management endpoints are planning/operational surfaces and must not bypass finance/document source-of-truth rules.

Canonical contracts:

- `saveWarRoomDemand`, `removeWarRoomDemand` and `updateDemandPrimljeno` manage day-scoped demand rows in `WarRoomDemand`.
- `saveDispecer`, `updateDispecer` and `removeDispecer` manage `DispecerPlan` rows with explicit planned/status lifecycle and timestamp updates.
- `getDispecer` returns today-only demand plus active/non-`zavrseno` plans.
- `getVozacPlans` reads `DispecerPlan` scoped to the authenticated Vozac's `entityID` and today's date, excluding `zavrseno` rows; the action block must be present in `handleAuthorizedRead` and is restricted to the Vozac role.
- `DispecerPlan` schema: `PlanID`, `Datum`, `DemandID`, `VozacID`, `VozacName`, `StanicaID`, `StanicaName`, `KupacID`, `KupacName`, `PlannedKg`, `Status`, `CreatedAt`, `UpdatedAt`.
- Dispatcher write invariant: Management dispatch operations write exclusively to `DispecerPlan` and `KamionStatus`. `OTK-*` otkup records must never be mutated from the dispatcher flow; `VozacID` assignment in `OTK-*` rows is prohibited from the dispatcher path.
- `updateKamionStatus` upserts one row per `VozacID` in `KamionStatus`; Vozac may update only own status, Management may update any driver.
- `saveIzdavanje` persists one row per agro issuing document into `Izdavanje`, serializes `stavke` as JSON and returns an `IZD-*` identifier.
- `saveIzdavanje` server-side idempotency by stable client issuance ID remains roadmap, not current contract.

### 9.14 Meteo / GIS Endpoints

GAS owns the online geo/meteo read pipeline and polygon persistence boundary, with important security caveats documented in section 16.

Canonical meteo rules:

- `getParcelMeteo()` first uses `MeteoLatest` if `LastFetch` is younger than 12 hours;
- stale/missing cache falls back to live Open-Meteo forecast retrieval;
- `scheduledMeteoFetch()` refreshes `MeteoLatest` and appends `MeteoHistory`;
- scheduled refresh runs 4 times daily in `Europe/Belgrade`;
- crop thresholds and spray-window logic are part of the GAS meteo service contract;
- public geo/meteo reads are accepted read exposure until a gated model replaces them.

Canonical geo rules:

- `getParcelGeo` reads geo point/polygon state;
- `saveParcelPolygon` persists polygon and centroid data, but its auth state is `NEEDS REVIEW`;
- geo master/source-of-truth rules are documented in section 14.

### 9.15 Fiskalni Endpoints

Fiscal endpoints are private kooperant financial-intake surfaces and must not become public utilities.

Canonical rules:

- `parseFiskalniImage` decodes an image, extracts a fiscal verification URL and delegates to `parseFiskalni`;
- `parseFiskalni` fetches/parses SUF verification payload and extracts receipt/item data;
- duplicate receipts are rejected by `VerificationUrl` inside `FISKALNI-<KooperantID>`;
- fiscal item matching order is `FiskalniMapiranje` → exact Artikli name → contains match → keyword score fallback;
- `saveFiskalni` writes private kooperant fiscal rows;
- `saveFiskalniMapiranje` and `createArtikal` are Management-only shared/master-data actions;
- private fiscal rows do not automatically mutate master `Artikli` unless the explicit controlled endpoint is used.

### 9.16 Monitoring Ingest and ErrorLog

GAS has two observability paths:

1. `logClientError` / `logError(...)` for PWA and GAS runtime/client errors.
2. `Monitoring.gs` `monitorPublic` / `monitor` for the production monitoring workbook.

`logClientError` is a pre-auth exception so field devices can report failures even after token expiry. If a valid token is supplied, GAS resolves entity context from token data; otherwise it uses bounded payload fallback.

`logError(source, action, message, details, entityID)` writes to `ErrorLog` with columns:

```text
Timestamp | Source | Action | Message | Details | EntityID | Severity
```

Error logging is best-effort and must never break the main auth/sync/business response.

### 9.17 Locking, Idempotency and Concurrency

All GAS actions that mutate Google Sheets or Drive must use `withLock(...)` or an equivalent concurrency guard.

Canonical write surfaces requiring locks include:

- role sync writes;
- dispatch/demand writes;
- truck status updates;
- Drive/PDF uploads;
- fiscal saves/mapping;
- parcel polygon writes where enabled;
- master/shared `createArtikal` writes;
- management agro issuing writes.

Idempotency should use stable `ClientRecordID` or a documented business key. Blind append is not acceptable for retryable field-device workflows.

### 9.18 Disabled / Deprecated Endpoints

Inactive endpoints must return explicit disabled responses.

Canonical disabled/deprecated behavior:

- `saveOtkupniListPdf` remains `FEATURE_DISABLED` until a real, tested GAS PDF generation/save implementation exists;
- disabled endpoints must not be counted as active route-health handlers;
- historical disabled state for `syncTrosak` is superseded: `syncTrosak` is active in the current architecture;
- deprecated routes may remain only as compatibility aliases with clear ownership and tests.

### 9.19 GAS Acceptance Gate Summary

Detailed checks live in `RELEASE_GATES.md`. The current GAS gate set must cover:

- active route presence;
- auth/login validation;
- 401/403/ownership mismatch behavior;
- batch sync response semantics;
- schema drift failures;
- master-sync soft-lock behavior;
- ErrorLog/client-error bridge;
- monitoring ingest;
- disabled endpoint behavior;
- no production business rows inserted by route smoke unless a fixture cleanup path is explicit.

---

## 10. Google Sheets Data Layer

Google Sheets is the shared online transport, projection and operational-state layer between PWA/GAS and the desktop master. It is not a replacement for the Excel/VBA canonical business tables.

Current ownership rule:

```text
Excel/VBA owns canonical formal business state.
GAS owns API-safe reads/writes into Google Sheets.
Google Sheets owns transport/shared-state/read-model persistence.
PWA owns local offline state until sync reconciliation.
```

Google Sheets data must therefore be documented as one of these categories:

- **transport state** — pending/synced PWA records waiting for or reflecting master import;
- **exported projection** — Stammdaten, Kartice, MgmtReports and selected read models exported from desktop master;
- **operational log** — ErrorLog, LoginLog and monitoring workbook tabs;
- **coordination state** — `SyncControl` / `MASTER_SYNC_LOCK` and similar runtime coordination rows.

### 10.1 Spreadsheet Families

The active Google Sheets layer contains these workbook families:

| Workbook / sheet family | Purpose | Canonical owner | Notes |
|---|---|---|---|
| `Stammdaten` | PWA master-data projection and selected coordination state | VBA export + GAS read helpers | main shared master-data projection |
| `OTK-*` | Otkupac PWA transport sheets | GAS writes, VBA imports/writebacks | one sheet family per Otkupac/role deployment pattern |
| `VOZ-*` | Vozac / zbirna transport sheets | GAS writes, VBA imports/writebacks | carries technical ID and business `BrojZbirne` separately |
| `TRETMAN-<KooperantID>` | treatment/agromere history | GAS `syncTretman` | scoped per kooperant |
| `TROSKOVI-<KooperantID>` | kooperant expense records | GAS `syncTrosak` | scoped per kooperant |
| `FISKALNI-<KooperantID>` | private fiskalni records | GAS fiskalni flow | scoped per kooperant |
| `Kartice` | exported kooperant card/read model | VBA export | derived read model |
| `MgmtReports` | management report projections including `OtkupiAll` | VBA export | derived read model for PWA/Management/Otkupac overview |
| `ErrorLog` / `LoginLog` | GAS/PWA operational audit logs | GAS | not business source-of-truth |
| `OtkupApp_Monitoring_PROD` | production monitoring workbook | GAS `Monitoring.gs` | documented in section 15 |

### 10.2 `Stammdaten` Workbook

`Stammdaten` is the primary shared master-data projection consumed by PWA and GAS read endpoints.

`SyncStammdatenToGoogle()` must find or create a spreadsheet named `Stammdaten`, persist `GOOGLE_STAMMDATEN_SHEET_ID`, and provision the canonical export tabs:

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

`Stammdaten` is an exported projection of desktop master/business state. It must not become an uncontrolled independent edit source. The current exception is parcel geo/polygon data, which may be edited through the GIS/PWA/HTML map flow and must be pulled back into desktop master before outbound Stammdaten export.

### 10.3 `Stammdaten / Config` vs Local Workstation Config

Google `Stammdaten / Config` and desktop `tblConfig` belong to Google/PWA/shared configuration.

Local workstation configuration belongs in:

```text
Sheet: LocalConfig
Table: tblLocalConfig
Columns: Kljuc | Vrednost | Opis
```

Local-only values such as `PDFTOTEXT_EXE_PATH` must not be stored in Google/PWA config tabs. This prevents per-workstation executable paths from leaking into tenant/shared configuration.

### 10.4 Role-Specific Transport Sheets

Role-specific sheets are transport/shared operational state between PWA/GAS and the desktop master. They are not the final source of truth for formal finance/document state.

Canonical transport families:

```text
OTK / otkup queue sheets
VOZ / zbirna transport sheets
TRETMAN-<KooperantID>
TROSKOVI-<KooperantID>
FISKALNI-<KooperantID>
```

Common identity/lifecycle concepts:

| Field | Meaning |
|---|---|
| `ClientRecordID` | stable client-generated identity used for idempotency |
| `ServerRecordID` | technical server/master sync identifier |
| `SyncStatus` | lifecycle/status state for sync/import/writeback |
| `CreatedAtClient` | client-side creation timestamp |
| `UpdatedAtClient` | client-side update timestamp |
| `UpdatedAtServer` | backend/master update timestamp |
| `ReceivedAt` | GAS/server receive timestamp where applicable |
| `DeviceID` | client/device diagnostic identity where applicable |

PWA/GAS retry logic and VBA imports must preserve these fields. They are required for idempotency, duplicate detection, operator diagnosis and safe writeback.

### 10.5 OTK Sheet Contract

The canonical OTK Google Sheet header follows the GAS-first `COLUMNS` order:

```text
ClientRecordID | ServerRecordID | CreatedAtClient | UpdatedAtClient | UpdatedAtServer | SyncStatus | DeviceID | OtkupacID | Datum | KooperantID | KooperantName | VrstaVoca | SortaVoca | Klasa | Kolicina | Cena | TipAmbalaze | KolAmbalaze | ParcelaID | VozacID | Napomena | ReceivedAt
```

Desktop import treats:

```text
SyncStatus = "Synced"
```

as pending-for-master-import.

After successful master import, desktop writeback may set:

```text
Synced>Master
```

Skipped/invalid rows may receive controlled statuses such as:

```text
Duplicate
SyncError[:reason]
```

Canonical OTK writeback targets:

```text
Sheet1!F -> SyncStatus
Sheet1!B -> ServerRecordID
```

`OTK-*` / `OTK-ST-*` sheets are operational inbound/live queue sheets for PWA-origin otkup records. They are not the canonical historical read source for all otkup records. Formal otkup history is owned by `tblOtkup`; PWA display reads it through the `MgmtReports/OtkupiAll` projection and merges it with operational queue rows.

### 10.6 VOZ Sheet Contract

The canonical VOZ Google Sheet header follows the GAS-first `ZBIRNA_COLUMNS` order:

```text
ClientRecordID | ServerRecordID | CreatedAtClient | UpdatedAtClient | UpdatedAtServer | SyncStatus | VozacID | Datum | KupacID | KupacName | VrstaVoca | SortaVoca | KolicinaKlI | KolicinaKlII | TipAmbalaze | KolAmbalaze | Klasa | OtkupRecordIDs | ReceivedAt | BrojZbirne
```

The technical sync identity and business document number must remain separate:

```text
Column B / ServerRecordID -> technical/master ZbirnaID
Column T / BrojZbirne    -> business document number x/ddmmyy[-rb]
```

`ServerRecordID` must not be reused as `BrojZbirne`.

For VOZ writeback, `WriteBackVOZSyncStatus` writes:

```text
Sheet1!B -> ServerRecordID / ZbirnaID
Sheet1!F -> SyncStatus
Sheet1!T -> BrojZbirne
```

`BrojZbirne` is normally generated PWA-side at `confirmZbirna` time. VBA keeps `GenerateBrojZbirne` as fallback for legacy/empty rows and desktop manual entry.

### 10.7 Plain-Text Formatting Rule

GAS must apply plain-text formatting to Google Sheet columns whose business values may contain `/` and would otherwise be auto-coerced into dates.

At minimum this applies to:

```text
TipAmbalaze
BrojZbirne
```

The canonical helper surface is `ensurePlainTextColumn(...)` where the GAS sheet bootstrap/write path owns formatting.

### 10.8 Per-Kooperant Treatment, Expense and Fiskalni Sheets

Per-kooperant sheets are scoped transport/private operational sheets.

Canonical families:

```text
TRETMAN-<KooperantID>
TROSKOVI-<KooperantID>
FISKALNI-<KooperantID>
```

Rules:

- Kooperant writes must be scoped to the authenticated `KooperantID` unless Management override is explicitly allowed.
- `ClientRecordID` is the idempotency key for sync processors.
- Retried records with the same `ClientRecordID` must update/return the existing logical row rather than append duplicates where the endpoint contract defines idempotency.
- Empty HTTP 200 responses are invalid for sync; GAS must return a structured batch response.
- PWA local state remains the user-visible pending/error source until reconciliation.

### 10.9 Kartice Export Workbook

`ExportKarticeToGoogle()` maintains a dedicated `Kartice` spreadsheet.

Canonical row shape:

```text
KooperantID | Datum | BrojDok | BrojParcele | Opis | Zaduzenje | Razduzenje | Saldo
```

`Kartice` is a derived/exported read model. It is not the canonical finance ledger.

### 10.10 MgmtReports Export Workbook

`ExportMgmtReports()` maintains a dedicated management reporting spreadsheet with tabs:

```text
SaldoOM
SaldoKupci
OtkupPoOM
PredatoPoKupcu
OtkupiAll
```

`MgmtReports` is a derived read model exported from desktop/master data. It must not be treated as a source table for business corrections.

### 10.10.1 `OtkupiAll` Master Otkup Read Projection

`MgmtReports/OtkupiAll` is the canonical Google Sheets read-model projection for PWA otkup overview across Management and Otkupac roles. It exists because `OTK-ST-*` / `OTK-*` operational sheets are inbound/live queue sheets for PWA-origin records, not a complete historical read model of `tblOtkup`.

Canonical read model:

```text
tblOtkup -> VBA export -> MgmtReports/OtkupiAll -> GAS -> PWA Management + Otkupac
```

Operational queue model remains active:

```text
PWA/Otkupac input -> OTK-ST-* -> VBA import -> tblOtkup
```

PWA otkup display must therefore merge:

```text
MgmtReports/OtkupiAll
+ OTK-ST-* operational rows
- duplicates
```

This ensures PWA views include:

- otkupi entered directly in VBA/master;
- otkupi entered through PWA and already imported/synced to VBA;
- otkupi still present in `OTK-ST-*` operational sheets.

`OtkupiAll` is read-only for PWA/GAS consumers and must not be used as a business correction surface.

### 10.11 Monitoring Workbook

The monitoring workbook is:

```text
OtkupApp_Monitoring_PROD
```

Canonical tabs include:

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

Monitoring workbook ownership, routing and alert behavior are documented in section 15. Monitoring tabs are operational observability state, not business source-of-truth.

### 10.12 Operational Logs: `ErrorLog` and `LoginLog`

`ErrorLog` stores GAS/PWA remote runtime errors written through `logError(...)` and `logClientError`.

`LoginLog` stores GAS login/session audit entries.

These logs are operational audit/diagnostic state. They may support debugging and production support, but they must not drive formal business document or finance state.

### 10.13 Sheet Registry and Header Contracts

GAS and VBA code must not assume stale headers.

Rules:

- sheet bootstrap must create required headers in canonical order;
- `SheetRegistry` / registry-assisted lookup is the preferred path where active;
- required header mismatches are schema drift, not silent defaults;
- sync processors must validate required columns before append/update;
- VBA import/export code must use the canonical index constants for OTK/VOZ headers;
- `rebuildSheetRegistry()` should be verified after backend deployment when registry behavior changes.

Schema drift must produce controlled failure/writeback/error-log behavior rather than corrupting rows through shifted column positions.

### 10.14 `SyncControl` Master-Sync Lock

The master-sync lock is stored in:

```text
Workbook: Stammdaten
Tab: SyncControl
```

Canonical schema:

| Parameter | Vrednost |
|---|---|
| `MASTER_SYNC_LOCK` | `YES` / `NO` |
| `MASTER_SYNC_UPDATED_AT` | timestamp |
| `MASTER_SYNC_MESSAGE` | operator-visible message |
| `MASTER_SYNC_OWNER` | `VBA` |

During full-cycle sync:

```text
MASTER_SYNC_LOCK = YES
MASTER_SYNC_OWNER = VBA
```

After sync:

```text
MASTER_SYNC_LOCK = NO
```

VBA must release the lock in cleanup even on failure. GAS may treat stale locks older than the configured max age as stale/unlocked according to the master-sync guard contract.

### 10.15 Write Blocking During Master Sync

When `MASTER_SYNC_LOCK` is active, GAS blocks write actions that could conflict with the VBA full-cycle import/export.

PWA behavior:

- treat lock responses as temporary soft-lock/retry states;
- keep local records pending;
- do not convert soft-lock into permanent business error;
- show clear operator/user feedback where needed.

GAS behavior:

- return structured lock responses;
- do not partially write blocked records;
- preserve idempotency semantics for retried records after unlock.

### 10.16 Parcel Geo Pull Before Stammdaten Export

Parcel point/polygon data may be created or edited outside VBA through PWA/GAS/HTML map flow.

Before outbound Stammdaten export, VBA must run:

```text
ImportParcelGeoFromGoogleToMaster
```

This prevents stale local `tblParcele` geo values from overwriting newer Google-side polygon data during the export.

If geo pull fails, outbound Stammdaten export should be aborted. This is a safety gate, not a best-effort cosmetic step.

### 10.17 Google Sheets External Side-Effect Boundary

Google Sheets writeback is an external side effect. It is not transactionally rollback-able by `clsTransaction` snapshots in Excel.

Rules:

- local Excel rollback does not automatically rollback Google writeback;
- writeback status must be explicit and operator-diagnosable;
- code must not claim full transactional coverage for Google-side writes;
- where possible, irreversible writebacks should happen after the local commit decision or be clearly separated from local transaction semantics.

### 10.20 Manual Editing and Test Data Hygiene

Production Google Sheets should not be manually edited except through documented operator/support procedures.

Rules:

- smoke/test rows must be removed or clearly marked as test data;
- stale test sheets must not be treated as production migration targets;
- production health checks should not be polluted by demo/test references;
- schema or header corrections must be performed through documented setup/export/bootstrap procedures, not ad-hoc manual reshaping.

---

### 10.21 Google Sheets Staging / Verify / Replace Write Model

`modGoogleSheets.WriteSheetData` must not clear the target tab before writing replacement data.

The canonical full-tab write model is:

```text
create staging tab
write values to staging
verify staging
replace target tab through phased rename
verify final target
```

This prevents the target tab from being left empty if value write fails after a clear operation.

Target replacement is phased to avoid Google Sheets title collisions:

```text
target -> backup
staging -> target
delete backup
```

If backup deletion fails after replacement, the new target is already live. The leftover backup tab is an operational/manual cleanup item, not a data-loss condition.

Quota hardening is part of the Google Sheets architecture:

- sheetId cache per spreadsheet;
- cache updates after staging creation and phased replace;
- HTTP retry for `429`, `500`, `502`, `503` and `504` responses;
- longer wait handling for quota-window `429`;
- write-request throttling;
- `AddSheetTab` no-op behavior when a tab already exists.

These rules reduce Google Sheets read/write quota pressure without changing the business sync flow.

### 10.22 Named-Tab Export Rule for Kartice

Kartice export uses a named tab, not the default Google tab `Sheet1`.

`modStammdatenSync.ExportKarticeToGoogle_Core` follows the same architectural pattern as Stammdaten and MgmtReports:

```text
ensure named tab exists
write named tab through WriteSheetData
```

Canonical tab constant:

```vba
Private Const KARTICE_TAB_NAME As String = "Kartice"
```

This removes the need for `Sheet1` special-case fallback logic in `modGoogleSheets`.



### 10.7 v6.24 Frontend Read/Asset Folder Boundary

The v6.24 UI/runtime work reinforces a separation already present in the sync architecture:

- PWA app-shell assets live in the frontend app tree and are controlled by service-worker cache versioning;
- Google Sheets operational/read-model data remains owned by the Google Sheets data layer;
- PWA display code must not treat asset deploy concerns as data freshness guarantees.

When UI assets are redesigned or lazy-loaded, the deployment contract is still:

```text
code asset change -> cache version bump
data export/change -> Google read-model / operational-sheet validation
```

The two must not be conflated.

## 11. PWA Architecture

This section states the current PWA offline, sync and runtime architecture.

### 11.1 App Shell
The PWA is an offline-first field application. The app shell must remain usable under low-connectivity conditions and must load critical runtime assets consistently from the service-worker cache.

Runtime/app-shell rules:

- `index.html` loads `src/js/services/db.js` only once.
- `tabs.js` protects access to `window.agroState` for non-Kooperant roles.
- `role-nav.js` uses canonical `cfg.type` routing and must not depend on mismatched `cfg.showMode` behavior.
- `sw.js` bumps `CACHE_NAME` whenever critical JS/runtime assets change.
- critical offline assets, including self-hosted Leaflet marker images when map surfaces are launch-relevant, belong in the service-worker asset list.

Critical runtime assets include:

```text
src/js/utils/format.js
src/js/utils/async.js
src/js/utils/merge.js
src/js/utils/sync-engine.js
src/js/app.js
role feature files participating in save/sync/render flows
self-hosted vendor assets used by offline app-shell behavior
```

### 11.2 Role Navigation
PWA role routing is based on the authenticated/configured role and the canonical role config model. Role screens must not call low-level sync internals directly; they go through app-level request/safe wrappers.

### 11.3 Runtime State Ownership
`AppState` is the client in-memory authority for shared runtime UI state. Shared runtime flags are normalized around `window.appRuntime`.

Legacy globals may exist as compatibility aliases, but they must not become independent sources of truth:

```text
db
stammdaten
mgmtData
qrScanner
selectedMera
parcelExpertOpen
appRuntime
```

### 11.4 IndexedDB Stores
IndexedDB is the local persistence layer for offline-first flows. `openDB()` provisions active stores and removes removed legacy kooperant store schemas during upgrade where applicable.

Active stores:

| Store | Purpose | Key / important indices |
|---|---|---|
| `CONFIG.STORE_NAME` | Otkup queue | `clientRecordID`, `syncStatus`, `datum` |
| `CONFIG.STAMM_STORE` | cached Stammdaten blobs | `key` |
| `zbirne` | Vozač zbirna queue/history | `clientRecordID`, `syncStatus` |
| `tretmani` | Kooperant treatments/agromere | `clientRecordID`, `syncStatus`, `datum`, `parcelaID` |
| `troskovi` | Kooperant expenses | `clientRecordID`, `syncStatus`, `datum` |

Operational access is normalized through:

```text
dbPut
dbGet
dbGetAll
dbGetByIndex
dbDelete
```

### 11.5 Offline-First Contract
The app must support local capture without immediate server success for normal field workflows.

Offline guarantees:

- app shell remains available offline;
- cached Stammdaten is readable offline;
- otkup, zbirna, tretmani and troskovi records can be queued locally;
- local signatures and queue state survive session changes;
- sync resumes when online and when temporary locks are released;
- shared logistics/planning truth still requires online refresh.

### 11.6 Shared Sync Engine
The active PWA sync contract is canonical across Otkupac, Kooperant and Vozač roles.

All app-level role sync entrypoints return this normalized shape:

```js
{
  ok: true | false,
  role: 'Otkupac' | 'Kooperant' | 'Vozac',
  synced: number,
  failed: number,
  results: [],
  reason: '',
  code: '',
  partial: false
}
```

The shared sync engine owns these rules:

- offline, unavailable database and in-flight states return normalized results instead of ad-hoc values;
- request-level failures return every record from the current attempted batch to `pending`;
- stale `syncing` recovery is age-gated and does not blindly recover fresh in-flight rows;
- backend statuses `synced`, `duplicate`, `existing`, `inserted` and `updated` are successful confirmations;
- missing backend per-record results are non-terminal and revert the affected row to `pending` with diagnostics;
- diagnostics include `lastSyncError`, `lastServerStatus`, `syncAttempts`, `syncAttemptAt`, `updatedAtServer` and `syncedAt` where applicable.

### 11.7 Single Sync Trigger Entrypoint
All sync triggers must go through:

```js
syncQueueSafe(reason)
```

Supported reasons:

```text
manual
online
interval
post-save
```

`syncQueueSafe(reason)` calls:

```js
requestRoleSync(reason)
```

Role dispatch:

- Otkupac -> `requestOtkupSync(reason)`;
- Kooperant -> `requestKooperantSync(reason)`;
- Vozač -> role-level gate to `syncZbirne()` until a dedicated `requestVozacSync(reason)` wrapper is introduced;
- Management -> canonical `no-sync-for-role`.

`runRoleSync(reason)` remains only a compatibility alias:

```js
async function runRoleSync(reason) {
    return requestRoleSync(reason || 'manual');
}
```

Trigger ownership:

- online event -> `syncQueueSafe('online')`;
- background interval -> `syncQueueSafe('interval')`;
- post-save otkup -> `syncQueueSafe('post-save')`;
- post-save zbirna -> `syncQueueSafe('post-save')`;
- manual More sync -> `syncQueueSafe('manual')`.

Low-level sync functions remain implementation internals and should not be called directly by UI triggers:

```text
syncQueue()
syncTretmani()
syncTroskovi()
syncKooperantNow()
syncZbirne()
```

### 11.8 Parallel Sync Guard
`syncQueueSafe(reason)` owns role-level in-flight protection:

- if no sync is running, it starts one;
- if one is already running, it returns canonical `already-running` / `ALREADY_RUNNING` and may mark a follow-up pass;
- runtime flags are cleared in `finally`;
- store-level guards remain inside `syncStore(...)` through `inFlightKey`.

### 11.9 Master-Sync Guard / Soft Lock
PWA observes the VBA master-sync lock through GAS and must show operator-visible sync state while the lock is active.

Canonical policy:

```text
Keep guard.
Use it to block GAS/server writes during VBA sync.
Do not block ordinary local PWA capture by default.
Treat MASTER_SYNC_ACTIVE as pending/retry, not as sync error.
```

Implications:

- PWA local work continues in offline-first pending mode.
- `MASTER_SYNC_ACTIVE` must not become a permanent `lastSyncError` business failure.
- pending records retry after the lock is released.
- workflows requiring immediate server persistence may opt into stricter behavior, but ordinary otkup capture should use soft-lock semantics.

### 11.10 Bootstrap Stale-`syncing` Recovery
Records may become stuck as:

```js
syncStatus: 'syncing'
```

after crash, browser kill, refresh, deploy, network break or interrupted sync.

Canonical helpers in `src/js/utils/sync-engine.js`:

```js
recoverStaleSyncingRecords(storeName)
recoverStaleSyncingStores(storeNames)
```

`src/js/app.js` owns role-aware bootstrap recovery:

```js
recoverStaleSyncingForCurrentRole(reason)
```

After `openDB()` and before role render/sync badge calculation, bootstrap must call:

```js
await recoverStaleSyncingForCurrentRole('bootstrap');
```

Role store mapping:

- Otkupac -> `CONFIG.STORE_NAME`;
- Kooperant -> `tretmani`, `troskovi`;
- Vozač -> `zbirne`;
- Management -> no local sync recovery.

Recovered stale records become:

```js
syncStatus: 'pending'
lastServerStatus: 'stale-syncing-recovered'
```

`lastServerStatus` is the canonical diagnostic field for this case to avoid showing a false business error.

### 11.11 Render Dedupe
Before rendering key UI lists, merged local/server datasets must pass through:

```js
dedupeRecordsForRender(records, aliasesFn?)
```

`src/js/utils/merge.js` owns this helper.

Identity aliases:

```text
srv:<serverRecordID>
cli:<clientRecordID>
```

Canonical rules:

1. identity aliases use `serverRecordID` and `clientRecordID`;
2. local priority record beats server-synced version;
3. local priority means `syncStatus === 'pending'`, `syncStatus === 'syncing'` or non-empty `lastSyncError`;
4. if both candidates are synced/non-priority, newer timestamp wins;
5. timestamp priority uses `updatedAtServer`, `syncedAt`, `updatedAtClient`, `createdAtClient`, then `receivedAt` / `ReceivedAt`;
6. records without identity aliases are preserved and not collapsed.

Required render paths:

- Otkup queue;
- Otkup pregled;
- Vozač zbirna pregled;
- Kooperant istorija tretmana;
- Otkup otprema overview;
- Otkup otprema assign runtime state.

### 11.11.1 PWA Otkup Overview Master + Operational Merge

PWA otkup overview is not allowed to depend only on `OTK-ST-*` / `OTK-*` operational sheets. Those sheets are inbound/live queue state for PWA-origin records. They do not contain every otkup entered directly into the VBA/master workbook.

Canonical PWA otkup display model:

```text
PWA display rows = MgmtReports/OtkupiAll + OTK-ST-* operational rows - duplicates
```

Merge inputs:

| Source | Meaning | Inclusion rule |
|---|---|---|
| `MgmtReports/OtkupiAll` | master projection of `tblOtkup` | include VBA-created and master-imported otkupi visible to the role/scope |
| `OTK-ST-*` / `OTK-*` | operational inbox/live queue | include PWA-origin rows that may be pending, synced, imported or still operationally relevant |

Dedup priority for otkup overview differs from generic local/server render dedupe because the same synced PWA otkup can have different `ClientRecordID` values in the two sources:

```text
OTK-ST-*  : ClientRecordID = original PWA UUID, ServerRecordID = OTK-xxxxx
OtkupiAll : ClientRecordID = VBA-OTK-xxxxx,  ServerRecordID = OTK-xxxxx
```

Therefore the canonical otkup overview dedupe priority is:

1. `ServerRecordID` / `OtkupID`;
2. `ClientRecordID`;
3. natural-key fallback.

For synced records, `ServerRecordID` is the stable shared key and must be checked before `ClientRecordID`. Natural-key fallback is only a last-resort display safeguard and must not hide legitimate distinct otkup rows.

### 11.12 Business Date Helpers
PWA business dates are canonical `YYYY-MM-DD` local-calendar strings.

Date-only fields must use local-calendar helpers, not UTC slicing:

```text
Datum
treatment dates
expense dates
otkup dates
otprema dates
zbirna dates
agrohemija issue dates
```

Canonical helper surface in `src/js/utils/format.js`:

```js
getTodayIsoDate()
getRelativeIsoDate(...)
toIsoDateOnly(...)
fmtDate(...)
localIsoDateFromDate(...)
```

Forbidden date-only patterns in feature code:

```js
Date.prototype.toISOString().slice(0, 10)
toISOString().split('T')[0]
```

Real UTC timestamp fields may still use `new Date().toISOString()`:

```text
createdAtClient
updatedAtClient
updatedAtServer
syncAttemptAt
ReceivedAt
syncedAt
```

### 11.13 Format Helpers
`src/js/utils/format.js` owns shared display helpers:

```js
formatNumber(value, options)
formatKg(value)
formatMoney(value)
```

Feature renderers must not assume local/private formatter globals.

### 11.14 Submit Locks
Critical save flows must be protected against double-tap/repeated-click duplicate local record creation.

Canonical helper in `src/js/utils/async.js`:

```js
withSubmitLock(lockKey, fn, options)
```

The helper:

- stores lock state in `window.appRuntime.submitLocks`;
- returns early if the same lock is active;
- may show an “already saving” toast;
- disables matching `[data-action="..."]` elements while saving;
- restores button state and clears the lock in `finally`.

Critical public save functions are thin lock wrappers around unlocked implementations:

```js
saveOtkup()       -> saveOtkupUnlocked()
confirmZbirna()   -> confirmZbirnaUnlocked()
agroSaveTretman() -> agroSaveTretmanUnlocked()
```

Canonical locks:

```text
otkup:save
zbirna:confirm
agro:tretman:save
```

`confirmZbirnaUnlocked()` must call:

```js
syncQueueSafe('post-save')
```

and must not call low-level `syncZbirne()` directly.

### 11.15 Client Error Reporting
PWA client-side observability is active.

Canonical frontend helper:

```js
reportClientError(error, context)
```

Owned by `src/js/utils/async.js`, it reports best-effort to GAS `logClientError`. Logging failure must never break app runtime.

Active reporting surfaces:

- `safeAsync(...)` catch paths;
- sync-engine exception paths;
- bootstrap/app startup catch paths;
- global `window.error`;
- global `window.unhandledrejection`.

Payload is intentionally small and privacy-bounded:

```text
errorAction
message
stack/details truncated for transport
role
entityID
app version if configured
URL
user agent diagnostics
```

### 11.16 Service Worker / Cache Contract
The service-worker cache version is part of deployment. `CACHE_NAME` must be bumped whenever critical runtime assets change. Old field devices may keep old cached JS unless deployment explicitly forces reload/update.

### 11.17 localStorage Ban and Exceptions
`localStorage` must not hold shared/canonical operational state.

Allowed exceptions:

- auth token / session helper;
- device ID;
- otkupac signature asset;
- temporary helper fallback such as capacity until Stammdaten load.

Forbidden:

- kamion status;
- dispatch plans;
- shared logistics state;
- canonical business entities.

### 11.18 Sync Badge and Queue Diagnostics
The PWA header/sync badge is derived from connectivity and local queue state. Visible states include:

```text
OFFLINE
SYNC...
ČEKA: n
ONLINE
```

Queue views should render pending/syncing rows and include inline server-error diagnostics when present. Stats based only on local IndexedDB rows are local operational signals, not server truth.

---


### 11.19 v6.24 PWA Design System and Role UI Model

v6.24 documents the PWA visual-system and role-flow redesign work performed after the v6.23 read-model convergence.

The current PWA UI model is based on a reusable design-system layer rather than role-local one-off CSS. The canonical design-system concepts are:

- brand tokens in `base.css`;
- self-hosted font system in `fonts.css`;
- reusable component classes in `components_v2.css`;
- shared header/body/card/field/button/pill/list/record patterns;
- role feature files that compose these primitives instead of inventing local variants.

The canonical brand-token family includes:

```text
--forest
--accent
--gold
--cream
--text-primary
--text-secondary
--text-muted
--border
--border-strong
--shadow-sm
--shadow-md
--shadow-lg
--shadow-xl
--radius-sm
--radius-md
--radius-lg
--radius-xl
```

Legacy aliases remain for compatibility where existing code still expects them:

```text
--primary
--primary-light
--primary-dark
--bg
--card
--text
```

The current font architecture is:

- Cormorant Garamond for display headings / selected large numeric inputs;
- DM Sans for body text;
- self-hosted `woff2` assets;
- Latin and Latin-ext coverage, including Serbian characters such as `Č`, `ć`, `Š`, `š`, `Đ`, `đ`.

The reusable UI component layer includes:

- `.app-hd` header system;
- `.app-body` content container;
- `.step` wizard primitives;
- `.card` variants;
- `.scan-cta`;
- `.koop-chip`;
- `.field` / `.field__input` / `.field__select`;
- `.class-picker` / `.pkg-picker`;
- `.btn-v2` variants;
- `.sticky-bar`;
- `.pregled-hero`;
- `.pills` / `.pill`;
- `.list-head`;
- `.rec` record card family.

Role UI flows must reuse this layer before adding new CSS primitives.

### 11.20 v6.24 PWA Otkup Form, Otkupni List, Otprema and Pregled UI Contracts

The Otkupac PWA UI now uses a staged, mobile-first flow.

Otkup form current contract:

- form is a 5-step state machine;
- Step 1 selects Kooperant through scan CTA / chip / select flow;
- quantity and price use large field inputs;
- class picker is limited to `I` and `II`;
- package picker exposes canonical package options such as `12/1`, `6/1`, `2/1`;
- the save bar is sticky and shows live total;
- driver selection is not part of the otkup form and belongs to the Otprema flow;
- note/napomena field is not part of the current form contract;
- picker logic must use event delegation, not inline script, to preserve CSP compatibility.

Canonical Otkup form runtime helpers include:

```text
bindKlasaPicker
bindTipAmbalazePicker
evaluateOtkupFormState
applyOtkupFormState
bindOtkupFormStateListeners
bindOtkupFormUIEvents
```

Otkupni list modal current contract:

- uses `.ol-*` classes;
- has forest header;
- displays kg in accent/Cormorant presentation;
- uses a 2x2 information grid;
- supports expandable details;
- supports signature pad;
- uses a sticky action bar;
- preserves existing data-action hooks:
  - `otkupni-confirm`;
  - `otkupni-clear-signature`;
  - `otkupni-print`;
  - `otkupni-save-pdf`;
  - `close-otkupni-list-modal`;
  - `sigKooperant`.

Otprema current contract:

- uses the shared `.app-hd` / `.app-body` shell;
- supports summary, detail and success views;
- uses scan CTA and truck/hero card patterns;
- uses `btn-v2` button variants;
- shows pending summary by kooperants / blocks / kg;
- shows driver chip with real available driver fields only;
- uses sticky `Utovari` bar with live kg;
- selected block cards have visible selected state.

Pregled / Danas current contract:

- uses the shared app header with operational overview wording;
- uses 2x2 stats grid / summary-card primitives;
- uses filter pills for `Danas`, `Juče`, `Sve`, `Bez vozača`, `Problemi`;
- uses styled date range fields;
- renders dynamic `.danas-*` cards and detail modal;
- problem badges must map to existing semantic badge classes, not missing CSS variants.

### 11.21 v6.24 PWA Runtime Hygiene, Cache and Lazy Loading

v6.24 also documents runtime cleanup and performance rules discovered during the redesign.

Service-worker and cache rules:

- `CACHE_NAME` / equivalent cache version must be bumped when critical app-shell assets change;
- font CSS and self-hosted font assets must be part of the offline app-shell when they are required for stable role UI;
- heavy vendor scripts should not be forced into initial service-worker precache when they are loaded lazily at runtime;
- `lazy.js` is the canonical helper for idempotent Promise-based script loading.

Lazy-loading current contract:

- `jsPDF` is loaded lazily by document/PDF features;
- Leaflet is loaded lazily by parcel/map features;
- Chart.js is loaded lazily by management chart rendering;
- Firebase compat libraries are loaded lazily only by intercom features that need them;
- Kooperant and Vozač roles must not load Firebase compat libraries unless their active feature path requires it.

Runtime hygiene rules:

- `viewport-fit=cover` is required for iOS safe-area behavior;
- sticky save bars must account for bottom navigation and safe-area inset;
- date-range grid columns must use `minmax(0, 1fr)` to avoid overflow;
- form code must not use periodic polling for field state when event-driven state updates are available;
- `setFieldValue()`-style helpers should dispatch change events when code updates form values programmatically;
- Serbian decimal input must be parsed through a decimal helper that accepts comma input, rather than raw `parseFloat`.

Current helper additions / changes include:

```text
parseDecimalInput
setFieldValue
lazyLoadScript
```

The current PWA UI cleanup also removes hardcoded role/station text where runtime config can provide it. Role eyebrows and station labels must be derived from `CONFIG.ENTITY_NAME` / configured user context rather than hardcoded station names.

## 12. Role Workflows

### 12.1 Otkupac
Otkupac flow supports offline local otkup capture, queued sync and queue diagnostics.

Current contract:

- otkup records use client-generated identity for local queue/idempotency;
- local queue rows move through pending -> syncing -> synced/error-style diagnostics;
- post-save sync trigger goes through `syncQueueSafe('post-save')`;
- UI/manual sync goes through app-level wrappers, not low-level `syncQueue()`;
- sync badge and queue list show local state clearly;
- missing per-record backend confirmation reverts the row to `pending` with diagnostics rather than silently treating it as synced;
- otkup overview must show both VBA/master-created otkupi from `MgmtReports/OtkupiAll` and operational PWA queue rows from `OTK-ST-*` / `OTK-*`;
- merged otkup overview must deduplicate by `ServerRecordID` / `OtkupID` before `ClientRecordID`.

### 12.2 Kooperant
Kooperant sync covers both treatment/agromere records and farm expense records.

Current contract:

- `syncKooperantNow()` / role sync covers `syncTretmani()` and `syncTroskovi()`;
- top-level Kooperant result aggregates both child results;
- both stores returning `reason: 'no-pending'` is still a canonical normalized result;
- treatment history and expense history must dedupe local/server rows before render;
- Kooperant stale-`syncing` recovery covers `tretmani` and `troskovi`.

### 12.3 Vozač
Vozač flow supports assigned otkup overview, local zbirna creation, queued sync and driver-side business-number visibility.

Current contract:

- `loadVozacData()` loads assigned otkupi through `getVozacOtkupi` and merged zbirne through `getMergedZbirneForVozac()`;
- otkupi already referenced inside existing zbirne via `otkupRecordIDs` are removed from the active pool;
- `confirmZbirna()` creates one local `zbirne` record by summing class I/II kilograms, total ambalaža and concatenated `otkupRecordIDs` from still-free otkupi;
- new zbirne rows are stored in IndexedDB store `zbirne` with `entityType = "zbirna"`, `schemaVersion = 1`, sync metadata, `clientRecordID`, `serverRecordID` and separate business field `brojZbirne`;
- `syncZbirne()` follows the pending -> syncing -> synced lifecycle and targets backend action `syncZbirna`;
- backend statuses `duplicate`, `existing`, `inserted`, `updated` and `synced` are successful confirmations;
- driver UI renders `brojZbirne` before any technical server identifier.

### 12.4 `BrojZbirne` PWA-First Ownership
`ServerRecordID` is a technical PWA/GAS sync identifier. `BrojZbirne` is a business document number and must not be copied from `ServerRecordID`.

Primary generation is PWA-side at `confirmZbirna` time.

Format:

```text
x/ddmmyy[-rb]
```

Rules:

- extract numeric part from `VozacID`, e.g. `VOZ-00004` -> `4`;
- combine with local business date in `ddmmyy`, e.g. `4/040526`;
- append `-2`, `-3`, etc. for subsequent zbirne from the same driver on the same day;
- sequence counts all zbirne for that driver/date including soft-deleted rows;
- storno does not reclaim sequence numbers;
- VBA `GenerateBrojZbirne` remains fallback for legacy/pre-rollout rows and desktop manual entry.

### 12.5 Management
Management has no local sync recovery store by default and normally returns canonical `no-sync-for-role` from role sync. Management reads aggregated backend/Sheets state and owns planning/overview surfaces rather than local offline queue capture.

Management otkup overview must read both sides of the otkup architecture:

- canonical master projection: `MgmtReports/OtkupiAll`;
- operational queue/live rows: `OTK-ST-*` / `OTK-*`.

This ensures Management can see otkupi entered directly in VBA/master as well as PWA-created otkupi that are still present in operational queue sheets. Partneri, Kooperanti, Kupci, Kartice and report surfaces remain separate read paths within `getMgmtAll` / Management hydration, but otkup display must not regress to operational `OTK-ST-*` only.

### 12.6 Excel Operator
The Excel operator owns formal master sync through the VBA full-cycle orchestrator, not ad-hoc direct imports/exports. PWA field work during a full-cycle sync continues locally and retries after lock release.

---

## 13. Agrohemija / Digitalni Agronom

Agrohemija / Digitalni Agronom is the current architecture surface for warehouse agrohemija movements, PWA management issuing, kooperant treatment evidence, karenca/history support and related stock visibility.

The stack intentionally reuses the active surfaces below and does not introduce a separate magacin/agro backend service module:

- desktop `modAgrohemija`;
- desktop `frmAgrohemija`;
- PWA management `src/js/features/management/agrohemija.js`;
- PWA kooperant `src/js/features/kooperant/agromere.js`;
- GAS `syncTretman`, `getTretmaniForKooperant` and `saveIzdavanje` boundaries.

GAS treatment sync is accepted as the current launch contract for treatment evidence/history/karenca. Treatment sync does **not** automatically decrement server-side `magacinkoop` stock unless a future backend stock-consumption contract defines that behavior.

### 13.1 Desktop `modAgrohemija` Contract

`modAgrohemija` is the canonical desktop business module for agrohemija warehouse operations. It must not be replaced by a new magacin/agro service module without an explicit architecture revision.

Current rules:

- `SaveMagacin` owns single-row warehouse journal creation for `MAG_ULAZ` and `MAG_IZLAZ`.
- `SaveMagacin` validates:
  - required article;
  - valid movement type;
  - positive quantity;
  - required kooperant for `MAG_IZLAZ`.
- `MAG_IZLAZ` must check current article stock before writing a movement row.
- `SaveMagacin_TX` is the single-row transaction wrapper and snapshots `tblMagacin` before delegating to `SaveMagacin`.
- `SaveMagacin_TX` emits `MAGACIN_SAVE_SUCCESS` after successful commit.
- Failure paths emit `MAGACIN_SAVE_FAIL` and `Monitor_Error` where monitoring is configured.
- Monitoring after commit is best-effort and must never convert a committed business save into a false failure.
- The original failure reason from `SaveMagacin` is preserved for the transaction wrapper, logs and operator diagnostics.
- Required column reads in stock, report, debt and parcela paths use fail-fast schema guards such as `RequireColumnIndex`.
- `GetMagacinStanje()` is the canonical desktop read model for current article stock after excluding stornirano rows.
- `ReportIzdavanjePoKooperantu()` supports open-ended date filters; `datumOd` and `datumDo` are evaluated independently.
- `ReportStanjePoDobavljacu()` is the correct public spelling.
- The older `ReportStanjePoDoabvljacu()` spelling may remain as a compatibility wrapper.
- `GetAgroAbzug()` remains a finance/deduction helper owned outside `modAgrohemija`; `modAgrohemija` must not duplicate its logic.

### 13.2 Desktop `frmAgrohemija` Basket Contract

`frmAgrohemija` is the canonical desktop form for warehouse/agrohemija operations.

Current rules:

- The form owns two in-memory baskets:
  - issue basket: `m_KorpaIzlaz`;
  - receipt basket: `m_KorpaUlaz`.
- Basket commits use one explicit `clsTransaction` over `tblMagacin` so the entire basket commits or rolls back together.
- The form tracks whether the transaction has started and performs rollback in the error handler for both issue and receipt finish actions.
- Issue baskets run an aggregated pre-commit stock check by `ArtikalID`.
- Multiple basket lines for the same article are summed before comparing to current stock.
- Optional add-to-basket UX validation may block adding a line that would make the current basket exceed available stock.
- Multiple parcel IDs are serialized with semicolon (`;`) separators.
- The form may show user-facing `MsgBox` feedback because it is a UI layer.
- Business/data modules must not own UI popups as control flow.
- Return navigation and query-close behavior must route back to `frmOtkupAPP` without embedding business writes.

### 13.3 PWA Management Agrohemija Issuing

The management `agrohemija.js` module is the active PWA surface for barcode/parcel-aware agrohemija issuing and signature-backed otpremnica confirmation.

Current rules:

- Recommendation quantity always represents real quantity in the article unit of measure.
- If packaging exists, package rounding computes:

```text
pakCount = ceil(rawQty / pakovanje)
finalQty = pakCount * pakovanje
```

- `finalQty` is the persisted/displayed quantity.
- Package count may be displayed as explanatory `pakInfo`, but it is not the saved quantity.
- Multiple parcel IDs are serialized with semicolon (`;`) separators to match desktop `frmAgrohemija`.
- `izdZavrsi()` opens the printable/signable otpremnica modal; it does not write directly.
- The final save is protected by `withSubmitLock` to reduce double-submit duplicate risk.
- The modal payload carries a stable client-side issuance identity for display/PDF and future backend idempotency compatibility.
- GAS `saveIzdavanje` remains Management-only under the existing endpoint/sheet contract.
- Because GAS `saveIzdavanje` is not server-idempotent by the new client issuance identity in the current contract, client submit-lock is the launch mitigation for duplicate user submit.
- `izdReset()` clears cart, selected kooperant, selected article/quantity, parcel list, note, recommendation state, modal data and barcode debounce state.
- Render functions guard missing DOM targets so reset/render can safely run during partial screen lifecycle transitions.

### 13.4 PWA Kooperant Digitalni Agronom / Agromere

The kooperant `agromere.js` module is the active Digitalni Agronom / treatment evidence surface.

Current rules:

- `agroCalcPreporuka()` treats `finalQty` as real quantity in the selected article unit of measure in every branch.
- For packaged articles:

```text
finalQty = ceil(rawQty / pakovanje) * pakovanje
```

- Lager warnings compare the same real `finalQty` against locally loaded `art.stanje` from `stammdaten.magacinkoop`.
- `agroPrimeniPreporuku()` writes the real final quantity into `agroKolicina` and `agroState.kolicina`.
- `agroValidateKolicinaNaLageru()` validates the current input quantity against local stock before local treatment save and writes the accepted value back to `agroState.kolicina`.
- `agroSaveTretman()` is a thin submit-lock wrapper around `agroSaveTretmanUnlocked()`.
- The canonical lock key is:

```text
agro:tretman:save
```

- `agroSaveTretmanUnlocked()` persists a local IndexedDB `tretmani` record first, invalidates the treatment cache, then requests sync when online.
- `agroResetState()` clears both `active` and `selected` measure-button classes so stale UI selection cannot imply active business state.
- The treatment path is authoritative for evidence/history/karenca.
- The treatment path does not automatically decrement server-side stock.

### 13.5 GAS Treatment Sync Boundary

The accepted active backend contract for treatment sync is:

- action: `syncTretman`;
- authenticated POST action;
- allowed roles: `Kooperant`, `Management`;
- a Kooperant caller may sync only its own `kooperantID`;
- Management may override where explicitly authorized;
- payload requires `records[]`;
- processing executes under `withLock(...)`;
- each row is handled by `processTretmanRecord(record, kooperantID)`;
- batch response is returned through `buildBatchSyncResponse(results)` and includes `results[]` for PWA per-record reconciliation;
- storage target is `TRETMAN-<KooperantID>`;
- `processTretmanRecord` is idempotent by `ClientRecordID`;
- existing rows return a successful terminal/existing status instead of appending duplicates;
- `getTretmaniForKooperant(kooperantID)` reads the same per-kooperant treatment sheet and returns server records for history/karenca reload.

Accepted current limitations:

- GAS does not automatically decrement `magacinkoop` stock when a treatment syncs.
- Stricter backend validation for `Zastita` / `Prihrana` article and quantity remains recommended but is not part of the current GAS contract.

### 13.6 GAS Management Issuing Boundary

`saveIzdavanje` remains the active GAS endpoint for management agrohemija issuing.

Current rules:

- `saveIzdavanje` remains Management-only.
- It persists through the existing `Izdavanje` sheet contract.
- The current server contract does not guarantee retry idempotency by stable client issuance ID.
- Client-side submit lock is therefore required for the current launch mitigation.
- Optional server-side idempotency by stable client issuance ID is roadmap, not current architecture.

### 13.7 Quantity Semantics Invariant

Across desktop Agrohemija, management PWA Agrohemija and kooperant Digitalni Agronom, saved quantities are real quantities in the article unit of measure.

Package count is display/explanation metadata only. It must not be used as the persisted `Kolicina`, `KolicinaUpotrebljena`, `DozaPreporucena`, `DozaPrimenjena` or issuing quantity unless the article unit itself is literally a package.

### 13.8 Stock Ownership and Limits

Current stock truth is split by workflow:

- desktop `tblMagacin` / `GetMagacinStanje()` is canonical for desktop warehouse stock after excluding stornirano rows;
- PWA kooperant treatment uses locally loaded `stammdaten.magacinkoop` for client-side availability validation;
- synced treatment evidence does not itself decrement server-side `magacinkoop` stock;
- future treatment-stock decrement must define retry, idempotency, storno and reconciliation behavior before becoming canonical.

### 13.9 Monitoring Boundary

Agrohemija monitoring is TX-boundary focused:

- `SaveMagacin_TX` emits `MAGACIN_SAVE_SUCCESS` after successful commit;
- failures emit `MAGACIN_SAVE_FAIL` and `Monitor_Error` where monitoring is configured;
- monitoring is best-effort and non-transactional;
- monitoring failure must never invalidate a committed `tblMagacin` write.

### 13.10 Current Known Limitations

- Treatment save does not automatically decrement server-side `magacinkoop` stock.
- `saveIzdavanje` does not yet provide server-side retry idempotency by stable client issuance ID.
- Backend validation for `Zastita` / `Prihrana` article and positive `KolicinaUpotrebljena` is recommended but not part of the current GAS treatment contract.
- Client submit-lock reduces duplicate user submits but does not replace future server idempotency for ambiguous network responses.

---

## 14. GIS / Parcele / Meteo

The GIS / Parcele / Meteo domain covers parcel master data, local parcel point capture, polygon editing, public geo/meteo read surfaces, cached weather/risk calculations and the operator/PWA UX around parcel maps.

Current architecture rule:

```text
tblParcele remains the desktop/master source of parcel business data.
Google Stammdaten / Parcele is the online projection and editor/runtime transport layer.
modGeoParcele owns parcel geo business/service logic.
frmStammdaten owns UI only.
```

### 14.1 Parcel Geo Ownership

`tblParcele` remains the canonical local master table for parcel data in the Excel/VBA workbook.

It owns:

- `ParcelaID`;
- `KooperantID`;
- cadastral number / `KatBroj` / `KatOpstina` fields;
- culture and area fields;
- GGAP/status metadata;
- point geo fields;
- polygon field once imported back into master;
- geo/risk/status metadata;
- notes and other parcel master fields.

Google `Stammdaten / Parcele` remains the online projection used by PWA, GAS and HTML geo tooling. It is not the primary owner of parcel business master data, but it may temporarily hold point/polygon edits created outside VBA until the master full-cycle sync pulls them back.

The geo workbook split remains active: parcel geo/meteo persistence may live in the workbook referenced by `GEO_SPREADSHEET_ID`, with at least `Parcele`, `MeteoLatest` and `MeteoHistory` tabs.

### 14.2 `frmStammdaten` Geo UI

`frmStammdaten` owns the operator UI micro-flow for parcel geo work.

It owns:

- mode selection through `.Tag`;
- visible controls and field binding;
- ListBox display;
- operator status messages;
- button event handlers;
- row-map / selection preservation after `LoadList`;
- inline status feedback.

`frmStammdaten` must not duplicate Google export logic or own parcel geo business rules. It should call domain/service functions instead of directly mutating every geo field.

Canonical UI calls:

```vb
SaveParcelGeoPoint
ClearParcelGeo
SyncSelectedParcelaToGoogle
```

Required controls:

```text
btnGeoOpen
btnPasteCoords
btnGeoClear
btnGeoSave
btnOpenMap
btnOpenPolygonEditor
txtNCoord
txtECoord
lblNCoord
lblECoord
lblGeoStatus
```

Geo controls are visible only when:

```text
Me.Tag = "Parcele"
m_SelectedRow > 0
lstData.ListIndex >= 0
```

Normal geo flow uses inline status through `lblGeoStatus` / `SetGeoStatus`, not blocking `MsgBox` prompts.

### 14.3 Canonical Parcel Geo Flow

#### 14.3.1 Select parcel

A selected parcel is valid only when `frmStammdaten` is in `Parcele` mode and the selected ListBox row maps to a physical `tblParcele` row.

The row map is the bridge between the visible ListBox row and the physical `tblParcele` row.

#### 14.3.2 Open external GeoSrbija / GeoData

`btnGeoOpen_Click`:

1. checks selected parcel;
2. reads `KatBroj` and `KatOpstina` from the selected ListBox row;
3. copies search text to clipboard;
4. opens GeoSrbija / external geo search;
5. writes inline status through `SetGeoStatus`.

No normal-flow `MsgBox` is used.

#### 14.3.3 Paste coordinates

`btnPasteCoords_Click`:

1. requires `Parcele` mode;
2. requires selected parcel;
3. reads clipboard text;
4. extracts two coordinate values;
5. fills `txtNCoord` and `txtECoord`;
6. writes inline status.

#### 14.3.4 Save point

`btnGeoSave_Click`:

1. requires selected parcel;
2. validates N/E coordinate text through `TryParseDouble`;
3. captures selected `ParcelaID`;
4. captures selected `ParcelaID` and calls `SaveParcelGeoPointByID parcelaID, nVal, eVal`;
5. reloads the ListBox;
6. reselects the same `ParcelaID`;
7. clears coordinate entry fields;
8. writes inline status.

The public row-index `SaveParcelGeoPoint` wrapper remains a compatibility surface only. The canonical API is `SaveParcelGeoPointByID`, which resolves the target row by exact `ParcelaID` lookup immediately before update.

#### 14.3.5 Clear point

`btnGeoClear_Click`:

1. requires selected parcel;
2. uses two-click confirmation through `mGeoClearConfirmPending`;
3. captures selected `ParcelaID`;
4. captures selected `ParcelaID` and calls `ClearParcelGeoByID parcelaID`;
5. reloads the ListBox;
6. reselects the same `ParcelaID`;
7. clears coordinate fields;
8. writes inline status.

The public row-index `ClearParcelGeo` wrapper remains a compatibility surface only. The canonical API is `ClearParcelGeoByID`, which resolves the target row by exact `ParcelaID` lookup immediately before update.

Point clear does not clear polygon data unless a separate polygon-clear operation is explicitly designed.

#### 14.3.6 Open map

`btnOpenMap_Click`:

1. requires selected parcel;
2. reads `Lat` and `Lng` from `tblParcele`;
3. validates numeric values;
4. calls `OpenGoogleMaps lat, lng`;
5. writes inline status based on Boolean result.

#### 14.3.7 Open polygon editor

`btnOpenPolygonEditor_Click`:

1. requires selected parcel;
2. reads `ParcelaID` from `tblParcele`;
3. displays inline syncing status;
4. calls `SyncSelectedParcelaToGoogle(parcelaID)`;
5. opens editor only when sync returns `True`.

Canonical editor open call:

```vb
OpenParcelPolygonEditor parcelaID
```

Current editor URL pattern:

```text
https://dusanmiladinovicvnm.github.io/otkupapp-pwa/parcel-draw.html?parcelaId=<encoded ParcelaID>
```

### 14.4 `modGeoParcele` Contract

`modGeoParcele` is the canonical domain/service module for parcel geo operations.

It owns:

- saving parcel point coordinates;
- clearing parcel point coordinates;
- UTM zone 34 to `Lat`/`Lng` conversion;
- selected parcel sync service for polygon-editor launch;
- local/Google parcel existence checks required by selected-parcel sync;
- logging/rollback for domain failures.

Layering rule:

```text
UI form -> domain/service module -> table/update helpers -> persistence/sync layer
```

`modGeoParcele` must remain UI-free:

```text
No MsgBox
No form control access
No UI color/status logic
```

#### 14.4.1 Canonical ParcelaID-based entry points

Canonical geo save/clear is ID-based, not row-index based:

```vb
Public Sub SaveParcelGeoPointByID(ByVal parcelaID As String, _
                                  ByVal nCoord As Double, _
                                  ByVal eCoord As Double)

Public Function SaveParcelGeoPointByID_TX(ByVal parcelaID As String, _
                                          ByVal nCoord As Double, _
                                          ByVal eCoord As Double) As Boolean

Public Sub ClearParcelGeoByID(ByVal parcelaID As String)

Public Function ClearParcelGeoByID_TX(ByVal parcelaID As String) As Boolean

Private Function RequireSingleParcelaRow(ByVal parcelaID As String, _
                                         ByVal sourceName As String) As Long
```

`RequireSingleParcelaRow` must apply exact-row semantics:

```text
Count = 0  => missing parcela error
Count = 1  => allowed
Count > 1  => duplicate ParcelaID error
```

Reason: physical `rowIndex` is fragile when `tblParcele` is sorted, filtered or reloaded. The selected `ParcelaID` must be captured at UI level and resolved again immediately before the update.

#### 14.4.2 Public compatibility entry points

Existing public row-index names may remain stable for older callers:

```vb
Public Sub SaveParcelGeoPoint(ByVal rowIndex As Long, ByVal nCoord As Double, ByVal eCoord As Double)
Public Sub ClearParcelGeo(ByVal rowIndex As Long)
```

They are compatibility wrappers only and should resolve the corresponding `ParcelaID` before delegating to the ID-based transaction functions.

#### 14.4.3 Transaction functions

Canonical transaction functions are the `ByID_TX` functions. Row-index transaction functions, if still present, are transitional wrappers and must not be the preferred architecture.

Rules:

- validate input before transaction;
- use `clsTransaction`;
- call `tx.BeginTx`;
- snapshot `TBL_PARCELE`;
- use `RequireUpdateCell` for writes;
- commit only after all writes succeed;
- rollback on error;
- log error through `LogErr`;
- return `False` on failure.

#### 14.4.4 Save point write set

`SaveParcelGeoPoint_TX` writes:

```text
COL_PAR_N
COL_PAR_E
COL_PAR_LAT
COL_PAR_LNG
COL_PAR_GEO_STATUS = point
COL_PAR_GEO_SOURCE = selenium
COL_PAR_METEO = Da
COL_PAR_DATUM_GEO = Now
COL_PAR_DATUM_AZUR = Now
```

#### 14.4.5 Clear point write set

`ClearParcelGeo_TX` writes:

```text
COL_PAR_N = empty
COL_PAR_E = empty
COL_PAR_LAT = empty
COL_PAR_LNG = empty
COL_PAR_GEO_STATUS = none
COL_PAR_GEO_SOURCE = empty
COL_PAR_METEO = Ne
COL_PAR_DATUM_AZUR = Now
```

`COL_PAR_POLYGON` is intentionally not cleared by point clear.

#### 14.4.6 UTM conversion

Canonical conversion surface:

```vb
Public Sub ConvertUTM34ToLatLng(ByVal eCoord As Double, ByVal nCoord As Double, _
                                ByRef lat As Double, ByRef lng As Double)
```

Rules:

- input order is `E`, then `N`;
- UTM zone is 34;
- output is decimal `Lat` / `Lng`;
- save function rounds `Lat` / `Lng` to 6 decimals.

### 14.5 Selected Parcel Sync

Canonical entry point:

```vb
Public Function SyncSelectedParcelaToGoogle(ByVal parcelaID As String) As Boolean
```

Purpose:

```text
Make the currently selected local tblParcele row available to the online Google layer before opening the polygon editor.
```

Required behavior:

1. trim and validate `parcelaID`;
2. confirm the local row exists in `tblParcele`;
3. call the canonical Parcele export wrapper:

```vb
SyncParceleToGoogle_Core(False)
```

4. verify that `parcelaID` exists in Google `Stammdaten / Parcele`;
5. return `True` only after successful verification.

`frmStammdaten` must not call a broad full Stammdaten sync directly from the editor button. The selected-parcel service is the correct editor-launch boundary.

`modGeoParcele` depends on `modStammdatenSync` exposing:

```vb
Public Function SyncParceleToGoogle_Core(ByVal showMessages As Boolean) As Boolean
```

Expected behavior of that wrapper:

- ensure Google OAuth/config;
- ensure `Stammdaten` sheet;
- ensure `Parcele` tab;
- reuse existing canonical `ExportParcele(sheetID)` logic;
- return `True` only if Parcele export succeeds.

This avoids duplicating `Parcele` header/export mapping in `modGeoParcele`.

Google verification reads `GOOGLE_STAMMDATEN_SHEET_ID` / `Stammdaten / Parcele` and verifies by `COL_PAR_ID` / `ParcelaID` header.

Fail-closed cases:

- missing/empty `parcelaID`;
- local row missing;
- Google sheet ID missing;
- `Parcele` tab empty or unreadable;
- `ParcelaID` header missing;
- selected `ParcelaID` not found after export.

If sync/verification fails, polygon editor must not open.

### 14.6 Polygon Editor and Geo Pull Safety

Selected-parcel sync is not a replacement for the full-cycle geo pull rule.

Full-cycle sync must run:

```text
ImportParcelGeoFromGoogleToMaster before outbound Stammdaten export.
```

Why both flows exist:

| Flow | Direction | Purpose |
|---|---|---|
| `ImportParcelGeoFromGoogleToMaster` | Google -> VBA | Protect PWA/HTML-created polygon before master export |
| `SyncSelectedParcelaToGoogle` | VBA -> Google | Ensure editor sees the selected local parcel before opening |

Overwrite safety rule:

```text
A broad Parcele export can overwrite Google polygon data if local PolygonGeoJSON is empty and the full-cycle geo pull discipline is not respected.
```

Production rule:

- run real polygon overwrite smoke before launch;
- do not open editor against stale/empty Google data;
- do not duplicate Parcele export mapping;
- do not create a second Google Sheets client for this flow.

Existing Google helper surfaces remain the supported path:

```vb
ReadSheetData
WriteSheetData
CreateSpreadsheet
GetSpreadsheetID
AddSheetTab
```

### 14.7 PWA Parcel Map Contract

The Kooperant parcel GIS screen is owned by the PWA parcel module.

Current contract:

- `loadParcele()` owns the kooperant parcel list + map screen;
- one Leaflet map instance is initialized per shell lifetime;
- `_parceleLoaded` acts as the screen-level load guard;
- parcel geometry is read per parcel through `getParcelGeo(parcelaId)`;
- `PolygonGeoJSON` renders polygons;
- valid `Lat` / `Lng` without polygon renders point markers;
- list cards, popup detail buttons and expert-panel toggles use delegated `data-action` hooks, not inline handlers;
- `focusParcel(...)` synchronizes list click, map focus, popup open, highlight and detail navigation;
- `openParcelaDetail(...)` ensures `loadKnjigaPolja()` has run and renders parcel detail tabs for `osnovno`, `meteo`, `radovi`, `troskovi` and `proizvodnja`;
- detail data is assembled from existing kooperant datasets (`kpData`, cached meteo and stammdaten), not one-off parcel endpoints;
- search/filter is client-side by text and culture over the exported parcel set;
- `invalidateParceleCache()` resets the loaded flag so the next screen open reruns rendering and geo/meteo initialization.

Popup content includes `KatBroj`, `Kultura`, `PovrsinaHa`, `KatOpstina`, `GGAPStatus`, raw `ParcelaID` and an explicit navigation action into parcel detail.

### 14.8 Meteo Pipeline

Meteo is an active GAS/PWA pipeline, not a static display-only feature.

#### 14.8.1 Trigger and source data

Scheduled meteo jobs run four times daily in `Europe/Belgrade`:

```text
00:00
06:00
12:00
18:00
```

On-demand parcel reads are also supported.

Source data:

- parcels with `MeteoEnabled = "Da"`;
- parcel point / centroid / polygon geo data;
- Open-Meteo forecast payloads;
- culture-specific risk thresholds.

#### 14.8.2 Cached-first read contract

`getParcelMeteo(parcelaId)` first tries `getParcelMeteoLatest()` and uses cached `MeteoLatest` data when `LastFetch` is younger than 12 hours.

If cached data is missing or stale, the backend falls back to live Open-Meteo forecast retrieval for the parcel centroid.

#### 14.8.3 Scheduled batch fetch contract

`scheduledMeteoFetch()`:

1. groups parcels by rounded 0.01 lat/lng buckets;
2. performs batch Open-Meteo fetch;
3. falls back to individual fetch/retry when batch fails;
4. derives current, daily and hourly models;
5. assesses frost, heat, rain and disease risk;
6. calculates spray-safe windows;
7. appends `MeteoHistory`;
8. overwrites `MeteoLatest` with current-state snapshots;
9. serializes risk, spray-window and forecast JSON for frontend reads.

`MeteoHistory` is append-only history. `MeteoLatest` is the overwrite/current-state view.

#### 14.8.4 Risk threshold contract

Risk thresholds are culture-specific for at least:

```text
Visnja
Jabuka
Sljiva
Kruska
Breskva
Malina
```

`_default` fallback is used for other cultures.

Threshold payload fields include:

```text
frostWarn
frostDanger
heatWarn
heatDanger
sprayWindMax
sprayRainHours
optimalTempMin
optimalTempMax
```

`assessRisk(...)` derives frost, heat, rain and disease risk plus aggregate level, min/max temperature and 24h rain totals.

#### 14.8.5 Spray-window contract

`calculateSprayWindow(...)` searches the next 72 hours for contiguous spray-safe windows.

The current suitability rule requires:

- near-zero precipitation;
- low precipitation probability;
- wind below crop threshold;
- temperature roughly greater than 5°C and below 35°C;
- humidity below the high-risk cutoff;
- enough dry hours ahead according to `sprayRainHours`.

#### 14.8.6 PWA meteo rendering

PWA parcel cards prewarm `window.meteoCache` from exported `stammdaten.meteoLatest` before per-parcel API fallback.

Card-level meteo uses a 6-hour `METEO_CACHE_TTL` and renders compact current/risk/spray/3-day forecast state.

The expert panel may show soil moisture, soil temperature, ET0, UV and solar radiation when those fields exist.

Meteo output is consumed by:

- kooperant parcel cards;
- parcel detail meteo tab;
- home-dashboard alerts;
- Digitalni Agronom treatment validation;
- spray timing suggestions.

### 14.9 GIS and Meteo GAS Endpoints

Current endpoint surface:

```text
getParcelGeo
saveParcelPolygon
getParcelMeteo
getParcelMeteoLatest
getAllMeteoLatest
scheduledMeteoFetch
```

Public exposure / auth boundary:

- `getParcelGeo`, `getParcelMeteo`, `getParcelMeteoLatest` and `getAllMeteoLatest` are public/pre-auth read bridges in the current backend model.
- `saveParcelPolygon` is an intentional public/pre-auth write exception in the active backend and must be protected operationally until a dedicated geo-editor/auth model replaces it.
- This is an acknowledged security gap, not an accidental omission.

### 14.10 `frmOtkupAPP` KPI Robustness

The sidebar KPI surface must tolerate dirty production data.

Problem class:

```text
SumOtkupKgForDate | 13 | Type mismatch
CountDocsForDate  | 13 | Type mismatch
```

Required behavior:

- do not raw-compare unknown values to `Date`;
- do not raw `CDbl` dirty values;
- skip invalid rows;
- return safe `0` or valid aggregate;
- log only meaningful unexpected failures.

Recommended helpers:

```vb
SafeDateKey(ByVal v As Variant) As String
SafeKpiDouble(ByVal v As Variant) As Double
```

Date comparison should normalize both sides:

```vb
targetKey = Format$(targetDate, "yyyy-mm-dd")
rowKey = SafeDateKey(data(i, colDate))
```

`RefreshSidebarKpi` should catch unexpected errors and set safe UI captions:

```text
KPI value = —
KPI delta = empty
KPI subtext = empty
```

The KPI surface must not block operator work.

### 14.11 Anti-Duplication and Error-Handling Rules

Do not duplicate `Parcele` export mapping. `modGeoParcele` must not create a parallel full `Parcele` header/export implementation when `ExportParcele` already exists.

Do not create another Google Sheets client for this flow. Use existing Google helper functions.

Do not move UI logic into `modGeoParcele`.

Do not move business geo mutations into `frmStammdaten`.

Error-handling boundary:

- UI layer uses inline status for normal geo problems;
- UI layer logs exception paths through `LogErr`;
- UI layer resets mouse pointer and delete confirmation in error paths;
- domain layer returns Boolean for service calls;
- domain layer logs failures and rolls back transactions where applicable;
- selected sync fails closed.

### 14.12 Geo Risk Rules

Current accepted risks:

| Risk | Status | Mitigation |
|---|---|---|
| `SyncSelectedParcelaToGoogle` depends on `SyncParceleToGoogle_Core` | accepted until compile/smoke | compile and run selected-parcel sync smoke |
| Selected parcel sync may export the full `Parcele` tab internally | accepted | reuse canonical mapping; optimize only if field data volume makes it slow |
| Existing Google polygon can be overwritten by stale local export | governed by full-cycle geo pull rule | keep `ImportParcelGeoFromGoogleToMaster` before outbound Stammdaten export |
| Missing `lblGeoStatus` in designer | possible | designer smoke must verify control exists |
| Geo clear does not clear polygon | intentional | design separate polygon-clear before changing this behavior |
| Dirty KPI data can exist | accepted | KPI helpers skip invalid rows and return safe aggregate |
| Public geo/meteo read bridge and `saveParcelPolygon` public write exception | acknowledged security gap | protect deployment URL operationally and track dedicated auth model in roadmap |

---

## 15. Monitoring and Observability

The monitoring layer is the current production observability path for pilot/production diagnosis. It is intentionally **best-effort** and does not replace local VBA logs, journaling, transaction rollback, SEF persistence, `AppendSEFEvent_Row`, `ProductionHealthCheck`, or business-table state.

Monitoring is not a business transaction participant. Monitoring failure must not cause a business save, SEF state transition, Banka mapping, document save, sync, backup, startup operation, or operator workflow to fail.

### 15.1 Monitoring Pipeline

Active path:

```text
VBA / PWA / GAS event source -> GAS Web App ingest -> Monitoring.gs -> OtkupApp_Monitoring_PROD Google Sheet
```

The subsystem covers:

- runtime health;
- central event logging;
- structured runtime/business errors;
- PWA sync and offline-queue status;
- SEF operational state;
- backup status;
- alerting;
- critical audit traces.

### 15.2 Layer Ownership

| Layer | Owner | Contract |
|---|---|---|
| VBA monitoring client | `modMonitoring` | Emits best-effort structured events from Excel/VBA to the GAS monitoring endpoint. |
| GAS monitoring ingest | `Monitoring.gs` | Owns `monitorPublic`, authenticated `monitor`, normalization, sanitization, routing, health updates, alert creation, watchdog checks and daily summaries. |
| Monitoring workbook | `OtkupApp_Monitoring_PROD` | Canonical Google Sheets storage for health, events, errors, SEF state, backups, alerts and audit-critical records. |

Monitoring complements existing local logs and business persistence. It must not become the only place where business truth, SEF truth, rollback state, or document status is stored.

### 15.3 VBA Monitoring Client

`modMonitoring` is the VBA client surface for production monitoring. It reads configuration from `tblSEFConfig` with headers:

```text
ConfigKey
ConfigValue
```

Canonical keys:

| ConfigKey | Purpose |
|---|---|
| `MONITORING_ENDPOINT` | GAS Web App `/exec` URL. |
| `MONITORING_SECRET` | Shared ingest secret matching GAS `MONITORING_INGEST_SECRET`. |
| `MONITORING_ENV` | Environment label, for example `DEV` or `PROD`. |

The client sends `APP_VERSION` from `modConfig.APP_VERSION` as `appVersion`. `deviceId` is generated from local workstation/user identity. Endpoint, secret and environment must remain workbook/runtime configuration and must not be hardcoded into the module.

Production sends are synchronous but short-timeout and best-effort. Debug/test connectivity checks may use longer timeouts, but production business flows must not wait on monitoring for business-API-scale durations.

### 15.4 GAS Monitoring Ingest

`Monitoring.gs` owns both public and authenticated monitoring ingestion.

Script Properties are the canonical deployment-specific configuration surface:

| Script Property | Purpose |
|---|---|
| `MONITORING_SPREADSHEET_ID` | Explicit ID of `OtkupApp_Monitoring_PROD`. |
| `MONITORING_ALERT_EMAIL` | Recipient for `ERROR` / `CRITICAL` alert emails. |
| `MONITORING_INGEST_SECRET` | Shared secret for public VBA ingest. |

GAS constants must store property **names**, not secret/config values:

```js
const MONITORING_PROP_SPREADSHEET_ID = 'MONITORING_SPREADSHEET_ID';
const MONITORING_PROP_ALERT_EMAIL = 'MONITORING_ALERT_EMAIL';
const MONITORING_PROP_INGEST_SECRET = 'MONITORING_INGEST_SECRET';
```

The public VBA ingest action is:

```json
{ "action": "monitorPublic" }
```

`monitorPublic` validates `monitoringSecret` against `MONITORING_INGEST_SECRET`. PWA/internal monitoring may use authenticated `monitor` under the existing token model.

### 15.5 Monitoring Workbook

`OtkupApp_Monitoring_PROD` is the canonical workbook for centralized operational monitoring.

Canonical tabs:

| Tab | Purpose |
|---|---|
| `Health` | Current component status snapshot. |
| `Events` | Append-only central event stream. |
| `Errors` | Structured runtime/business errors. |
| `SyncStatus` | PWA sync and offline-queue status samples. |
| `SEFStatus` | Current SEF operational status per invoice/correlation. |
| `UserSessions` | PWA/VBA session and heartbeat records. |
| `Backups` | Backup success/failure history. |
| `Alerts` | Open/resolved operational alerts. |
| `AuditCritical` | Critical business/audit-impacting event trace. |

Monitoring workbook tabs and headers are setup/schema artifacts, not production business-data migrations.

### 15.6 Health Model

Canonical health components:

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

Every valid event should update the appropriate component row in `Health`. Component inference is based on explicit component, `source`, and `eventType` where applicable.

Canonical mappings include:

| Event/source pattern | Component |
|---|---|
| `SEF_*` | `SEF API` |
| `VBA_*` | `VBA Client` |
| `BACKUP_*` | `Backup` |
| `MASTERDATA_*` / `STAMMDATEN_*` | `MasterData Sync` |
| `AUTH_*` | `Auth` |
| sync/heartbeat/offline queue events | `PWA Sync` / `Offline Queue` |

The watchdog also maintains health rows for state that may not emit frequent user events, such as `GAS API`, `Google Sheets DB`, `Auth`, `Backup` and `MasterData Sync`.

### 15.7 Event Payload Contract

Canonical event fields:

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

Canonical severities:

```text
INFO
WARN
ERROR
CRITICAL
```

`correlationId` is mandatory for events that belong to a business operation, including Faktura, SEF request, BankaImport row, PWA sync batch, backup or startup sequence.

Every valid event is routed to `Events`. Additional routing depends on severity, type and payload.

### 15.8 Routing Rules

Canonical routing:

| Condition | Destination |
|---|---|
| Every valid event | `Events` |
| `ERROR` / `CRITICAL` | `Errors` |
| Sync/heartbeat/offline queue events | `SyncStatus` and/or `UserSessions` |
| `SEF_*` events | `SEFStatus` |
| `BACKUP_*` events | `Backups` |
| Critical or audit-impacting events | `AuditCritical` |
| Alert-worthy events | `Alerts` |
| Any valid event | Updates relevant `Health` component |

Routing must be normalized and sanitized before writing to the workbook.

### 15.9 Security, Redaction and Timeout Rules

Monitoring must not log or echo sensitive material. Forbidden in monitoring payloads/debug output:

```text
monitoring secret
Google access token
Google refresh token
SEF API key
full UBL XML
full PDF/base64/file content
full raw SEF response body
password/PIN/authorization header
```

VBA debug output must never print the full JSON body. Debug diagnostics may print body length, event-type presence, HTTP status and redacted response summary.

Timeout profiles:

| Timeout | Purpose |
|---|---|
| `HTTP_TIMEOUT_MS = 1200` | Production best-effort monitoring send. |
| `HTTP_DEBUG_TIMEOUT_MS = 10000` | Manual/debug connectivity test. |

A slow monitoring endpoint must not freeze operator work.

### 15.10 VBA App Lifecycle Monitoring

`ThisWorkbook.Workbook_Open` emits `VBA_APP_OPEN` best-effort before delegating to `StartApp`.

`modMain.StartApp` emits:

```text
VBA_STARTAPP_START
VBA_STARTAPP_SUCCESS
JOURNAL_RECOVERY_WARN
```

Startup errors emit `Monitor_Error` with:

```text
moduleName = ThisWorkbook / modMain
procedureName = Workbook_Open / StartApp
entityType = App
entityId = Startup
correlationId = VBA-STARTUP
```

`StartApp` remains an orchestration layer. It must not duplicate detailed SEF, backup, finance or document monitoring owned by domain modules.

### 15.11 SEF Monitoring Coverage

SEF monitoring is the most detailed operational surface because SEF failures can have tax, invoicing and manual-review impact.

Startup recovery events from `modSEFService.RecoverAllStuckSEFSendingInvoices`:

```text
SEF_STARTUP_RECOVERY_START
SEF_RECOVERY_INVOICE_FOUND
SEF_RECOVERY_INVOICE_SUCCESS
SEF_RECOVERY_INVOICE_FAIL
SEF_STARTUP_RECOVERY_SUCCESS
SEF_STARTUP_RECOVERY_FAIL
```

Submit events from `SendInvoiceToSEF_TX`:

```text
SEF_SEND_START
SEF_SEND_ACCEPTED
SEF_SEND_SUCCESS
SEF_SEND_REJECTED
SEF_SEND_FAIL
SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING
```

Status refresh events from `modSEFStatusSync`:

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

`SEF_SEND_EXCEPTION_AFTER_LOCAL_SENDING`, `SEF_RECOVERY_INVOICE_FAIL`, `SEF_STATUS_REFRESH_EXCEPTION` and unknown/manual-review states are alert and audit-critical candidates. SEF monitoring must never replace `tblSEFSubmission`, `tblSEFEventLog`, `AppendSEFEvent_Row`, `SaveSEFSubmissionResult_Row` or the SEF state machine.

### 15.12 Business Transaction Monitoring Coverage

Monitoring is attached at transaction boundaries and operationally important failures. It must not emit row-level noise for every helper, read or update operation.

Canonical coverage:

| Module | Procedure / surface | Event contract |
|---|---|---|
| `modFaktura` | `CreateFaktura_TX` | `FAKTURA_CREATE_SUCCESS`, `FAKTURA_CREATE_FAIL`, `Monitor_Error` |
| `modNovac` | `SaveNovac_TX` | `NOVAC_SAVE_SUCCESS`, `NOVAC_SAVE_FAIL`, `Monitor_Error` |
| `modNovac` | `ApplyAvansToFaktura_TX` | `AVANS_APPLY_TO_FAKTURA_FAIL`, `Monitor_Error` |
| `modOtkup` | `SaveOtkup_TX` | `OTKUP_SAVE_SUCCESS`, `OTKUP_SAVE_FAIL`, `Monitor_Error` |
| `modOtkup` | `SaveOtkupMulti_TX` | `OTKUP_MULTI_SAVE_SUCCESS`, `OTKUP_MULTI_SAVE_FAIL`, `Monitor_Error` |
| `modDokumenta` | document-chain save TX wrappers | fail-only `DOKUMENT_SAVE_FAIL`, `Monitor_Error` |
| `modBankaMapiranje` | bank import mapping TX wrappers | `BANKA_MAP_SUCCESS`, `BANKA_MAP_FAIL`, `BANKA_IMPORT_SKIP`, batch automap events |
| `modMasterSync` | `ImportOtkupFromPWA_Core`, `ImportOtkupFromPWA_TX` | `MASTERDATA_SYNC_SUCCESS`, `MASTERDATA_SYNC_FAIL`, `Monitor_Error` |
| `modStammdatenSync` | `SyncStammdatenToGoogle` | `STAMMDATEN_SYNC_SUCCESS`, `STAMMDATEN_SYNC_FAIL`, `Monitor_Error` |

Document-chain saves are fail-only to avoid noisy monitoring during normal operations. Bank mapping, otkup, novac and faktura use success+fail because they represent high-value operational transitions.

### 15.13 Bank Mapping Monitoring

`modBankaMapiranje` is monitored at TX-wrapper level only. Base mapping helpers remain uninstrumented to avoid duplicate events.

Canonical events:

```text
BANKA_MAP_SUCCESS
BANKA_MAP_FAIL
BANKA_IMPORT_SKIP
BANKA_AUTOMAP_ALL_START
BANKA_AUTOMAP_ALL_SUMMARY
BANKA_AUTOMAP_ALL_FAIL
```

Covered TX wrappers:

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

Bank monitoring payloads may include `bankaImportID`, `partnerType`, `partnerId`, `resultId`, `linkedEntityId` and batch counts. They must not include sensitive bank account details beyond operational identifiers needed for diagnosis.

### 15.14 MasterData and Stammdaten Monitoring

`modMasterSync` emits:

```text
MASTERDATA_SYNC_SUCCESS
MASTERDATA_SYNC_FAIL
```

for PWA `OTK-*` import into the desktop master. Success includes file/import/skip/error counts.

`modStammdatenSync.SyncStammdatenToGoogle` emits:

```text
STAMMDATEN_SYNC_SUCCESS
STAMMDATEN_SYNC_FAIL
```

for the 13-tab Stammdaten export:

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

A full `13/13` export is `STAMMDATEN_SYNC_SUCCESS`. Partial export, missing OAuth configuration, missing PWA folder ID, failure to create the Stammdaten sheet or runtime exception is `STAMMDATEN_SYNC_FAIL`.

### 15.15 Backup Monitoring

Backup monitoring uses:

```text
BACKUP_SUCCESS
BACKUP_FAIL
```

The `Backups` tab captures:

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

Backup monitoring is best-effort. Monitoring failure must not change the backup procedure result. The watchdog independently checks whether a successful backup exists and whether the latest successful backup is older than the accepted freshness window.

### 15.16 Alerts, Audit and Manual Review

Alerts are generated for operational conditions that require attention, including:

```text
CRITICAL severity
ERROR severity where alert-worthy
needsManualReview = true
SEF UNKNOWN / manual-review state
SEF stuck/unknown age thresholds
backup failure or missing backup freshness
bank automap batch failure
sync pending/conflict/failed queue thresholds
GAS error spike
```

`AuditCritical` captures critical business-impacting events, including SEF critical states, payment/faktura-related critical events, manual overrides and critical bank/backup failures.

`nextAction` is canonical for operator guidance:

```text
WAIT
RETRY
MANUAL_REVIEW
CHECK_SEF_PORTAL
```

`needsManualReview = true` means the operator/admin must inspect the relevant operational sheet before retrying or accepting the state.

### 15.17 Watchdog and Scheduled Jobs

`runMonitoringWatchdog()` is the canonical scheduled health worker. It ensures the monitoring workbook exists, runs active health checks and creates alerts for stale/critical conditions.

Active checks cover:

- GAS API watchdog execution;
- Google Sheets DB and required monitoring tab accessibility;
- Auth secret presence;
- backup freshness and last successful backup age;
- MasterData/Stammdaten sync freshness and failure state;
- SEF unknown/stuck state age;
- PWA sync/offline queue failures, conflicts and stale pending rows;
- recent GAS error spikes.

Canonical trigger setup:

- `runMonitoringWatchdog` every 15 minutes;
- `sendDailyMonitoringSummary` once per day.

Alerts are deduplicated while unresolved by component, alert code and correlation ID.

### 15.18 Monitoring Test Contract

Active VBA test surface:

```text
TestMonitoring_All
TestMonitoring_Config
TestMonitoring_HTTP
TestMonitoring_ErrorEvent
TestMonitoring_SEFUnknown
TestMonitoring_BackupSuccess
TestMonitoring_BackupFail
```

A successful connectivity test requires both:

```text
HTTP Status = 200
response success = true
```

Expected successful response shape:

```json
{
  "success": true,
  "eventId": "...",
  "timestamp": "...",
  "severity": "INFO",
  "component": "VBA Client"
}
```

Production monitoring remains best-effort. Debug mode may wait longer for GAS cold-start behavior, but production sends must stay short-timeout and non-blocking to business flows.

### 15.19 Deployment and Data-Migration Boundary

No production business-data migration is required for monitoring. Required deployment/configuration actions are operational setup:

- ensure `OtkupApp_Monitoring_PROD` exists with canonical tabs and headers;
- set GAS Script Properties for monitoring spreadsheet ID, alert email and ingest secret;
- set `tblSEFConfig` keys for monitoring endpoint, secret and environment;
- redeploy GAS Web App after monitoring code changes;
- install watchdog and daily summary triggers;
- run the monitoring test suite and at least one end-to-end business-flow smoke.

---

## 16. Security and Compliance

Security is enforced as a layered contract across VBA/Excel, GAS, Google Sheets, PWA and operator setup. The active security model is not a single perimeter; it combines endpoint authorization, role/entity scoping, checked writes, bounded logging, HTTPS-only external transport, service-worker/CSP discipline and explicit treatment of known public/read exceptions.

### 16.1 Security Ownership Model

Canonical ownership:

| Area | Owner | Contract |
|---|---|---|
| Session authentication | GAS + PWA | PIN/login creates session token; PWA stores runtime session state and presents role-specific UI. |
| Endpoint authorization | GAS | Every write/quota-sensitive action must enforce role/entity scope before mutation. |
| Desktop secrets/config | VBA workbook config tables + local setup | SEF/monitoring/workstation config is read from approved tables or local config, not hardcoded in modules. |
| Monitoring redaction | VBA + GAS monitoring layers | Secrets, tokens, full XML/PDF/base64 and sensitive raw payloads are not logged. |
| Offline/local state | PWA IndexedDB | IndexedDB is local queue/cache, not shared authority; shared operational state must flow through GAS/Sheets/VBA. |
| Production health | VBA + monitoring workbook | `ProductionHealthCheck` and monitoring health gates expose dirty/stale/broken states before launch. |

### 16.2 SEF HTTPS-Only Rule

`SEF_BASE_URL` must start with `https://`.

Canonical behavior:

- plaintext `http://` SEF endpoints are rejected locally before network calls;
- rejection applies to client/config validation paths;
- API keys and invoice payloads must never be sent over plaintext HTTP;
- negative SEF config smoke must temporarily set an `http://` base URL and confirm local failure.

This rule is current security architecture, not a release note.

### 16.3 Secret Handling

Secrets and deployment-specific sensitive values must be configuration, not source-code constants.

Canonical examples:

| Secret / config | Storage |
|---|---|
| SEF API credentials | `tblSEFConfig` / SEF config owner paths |
| Monitoring endpoint | `tblSEFConfig.MONITORING_ENDPOINT` |
| Monitoring shared secret in VBA | `tblSEFConfig.MONITORING_SECRET` |
| Monitoring environment | `tblSEFConfig.MONITORING_ENV` |
| Monitoring spreadsheet ID | GAS Script Property `MONITORING_SPREADSHEET_ID` |
| Monitoring alert email | GAS Script Property `MONITORING_ALERT_EMAIL` |
| Monitoring ingest secret | GAS Script Property `MONITORING_INGEST_SECRET` |
| Local `pdftotext.exe` path | `tblLocalConfig.PDFTOTEXT_EXE_PATH` |

Rules:

- secret values must not be committed as literal source constants;
- GAS constants for monitoring must store property names, not secret values;
- debug output must not print full request bodies containing secrets;
- operator setup can create `Secrets` folders, but architecture still requires explicit non-logging and non-hardcoding of secret material;
- any future export/debug command must redact keys named or shaped like `token`, `secret`, `apiKey`, `authorization`, `password`, `PIN`, XML/PDF/base64 payloads or raw SEF bodies.

### 16.4 Token Handling and Session Scope

The token model is role/entity scoped.

Canonical rules:

- authenticated requests carry token-derived role and entity identity;
- non-management users may mutate only their own entity scope;
- Management may act as explicit override only where the endpoint authorization matrix allows it;
- sync payload entity fields must be checked against `tokenData.entityID` for non-management callers;
- when a `Vozac` updates truck/driver state, GAS must use or overwrite with `tokenData.entityID` instead of trusting a client-supplied `vozacID`;
- expired/missing/invalid tokens return auth errors rather than falling through to writes;
- PWA session state is runtime/client state and must not be treated as authoritative for ownership without GAS-side token validation.

### 16.5 GAS Endpoint Authorization Matrix

Every GAS action belongs to one of these authorization classes:

| Class | Meaning | Examples / notes |
|---|---|---|
| Public pre-auth exception | Intentionally allowed before normal token validation | `login`; `logClientError` for field error reporting; selected public geo/meteo reads if still deployed that way. |
| Authenticated read | Valid token required; role/entity filtering applies | Role-specific reads, kooperant/vozac scoped data. |
| Authenticated write | Valid token plus role/entity ownership required | OTK/VOZ/AGRO/TRETMAN/OPREMA/TROSKOVI sync writes. |
| Management-only write | Management role required | shared fiscal mapping, master article creation, shared/master-data writes, parcel polygon write if locked. |
| Disabled endpoint | Must return explicit disabled response | inactive or legacy actions such as placeholders. |

Canonical helper vocabulary:

```text
isManagement
requireRole
requireEntity
forbiddenResponse
withLock
```

Rules:

- write actions must not sit before the normal auth gate unless explicitly documented as a public exception;
- every Google Sheets or Google Drive write path must execute inside `withLock(...)`;
- parsing-only endpoints do not require `withLock(...)` unless they persist data or write logs/mappings;
- role/entity checks must be enforced by GAS, not only by hiding UI controls in PWA;
- `syncTrosak` follows the same locked, role/entity-scoped, batch-response contract as other active sync endpoints and must not return empty HTTP 200 bodies.

### 16.6 Role-Based Authorization Rules

Current role rules:

| Surface | Authorization rule |
|---|---|
| Otkupac sync / OTK writes | Otkupac may write only own `otkupacID`; Management override only where explicitly supported. |
| Vozac sync / VOZ writes | Vozac may write only own driver scope; Management may override where supported. |
| Kooperant treatment/expense sync | Kooperant must match `tokenData.entityID`; Management may act where supported. |
| `syncTrosak` | Allowed roles: `Kooperant`, `Management`; Kooperant scope must match authenticated entity. |
| `parseFiskalniImage` / `parseFiskalni` | Not public; allowed only for `Kooperant` or `Management`; Kooperant-scoped payload must match authenticated entity where supplied. |
| `saveFiskalniMapiranje` | Management-only because it writes shared fiscal-name mapping. |
| `createArtikal` | Management-only master-data write. |
| PDF/Drive document writes | Role/entity scoped; Otkupac may act only on own records; Management override where supported; writes locked. |
| `updateKamionStatus` | Vozac uses token entity as authoritative driver; Management may update any driver. |

### 16.7 Monitoring Redaction and Privacy Boundary

Monitoring is operational telemetry, not a data dump.

Forbidden in monitoring payloads, event messages, debug output and response summaries:

```text
monitoring secret
Google access token
Google refresh token
SEF API key
full UBL XML
full PDF/base64/file content
full raw SEF response body
password
PIN
authorization header
```

Rules:

- VBA monitoring debug output may show body length, event-type presence, HTTP status and redacted response summary only;
- GAS monitoring sanitizes/truncates messages and payloads before persistence;
- client error reporting truncates stack/details and redacts token/PIN/password/base64-like content;
- monitoring failures must never cause business transaction failure;
- monitoring correlation IDs are safe operational identifiers, not substitutes for raw payloads.

### 16.8 CSP / PWA Asset Rules

The PWA security posture depends on predictable local assets and constrained script loading.

Canonical rules:

- CSP uses `script-src 'self'` where possible;
- vendor/runtime assets required for offline app-shell behavior are self-hosted;
- `sw.js` cache version must be bumped whenever critical runtime JS/assets change;
- critical JS and offline assets must be present in the service-worker asset list;
- `index.html` must not double-load critical services such as `src/js/services/db.js`;
- dynamic user-facing strings should use `escapeHtml()` / safe rendering helpers rather than raw HTML injection;
- role UI must not be considered a security boundary; backend authorization remains mandatory.

### 16.9 LocalStorage, IndexedDB and Local State Rules

Local browser state is not shared operational truth.

Rules:

- IndexedDB is the canonical offline queue/cache for local PWA records;
- `localStorage` is allowed only for device-local preference/helper state such as `deviceID`, lightweight UI flags, signatures or cached capacity hints;
- `localStorage` must not carry shared operational state that other users/devices rely on;
- pending/error local records must stay visible/retryable and must not be hidden by server copies;
- sensitive auth/session data must be minimized and scoped to the runtime session model.

### 16.10 Local Workstation Security

Local workstation configuration is explicitly separate from Google/PWA remote config.

Rules:

- `tblLocalConfig` owns local workstation settings such as `PDFTOTEXT_EXE_PATH`;
- `tblConfig` remains Google/PWA config and must not become workstation-local config storage;
- `tblSEFConfig` owns SEF and monitoring workbook/runtime configuration;
- `Setup-OtkupApp.ps1` may create install folders, logs, temp, backups, secrets and tools directories, but executable/tool paths must still be validated;
- missing required tools such as `pdftotext.exe` must be reported by health/setup checks;
- local PDF extraction uses unique temp output and deletes temp files before/after extraction to avoid stale-content leakage.

### 16.11 Public Geo / Meteo and Parcel Polygon Boundary

Current docs contain a security-state conflict that must be explicitly reviewed before production handoff.

Known positions in the supplied material:

- one documented security hardening path states that `saveParcelPolygon` was moved behind token validation, made Management-only and wrapped in `withLock(...)`;
- later active AR/pass material still describes `saveParcelPolygon` as an intentional public/pre-auth exception or accepted risk;
- public geo/meteo reads are also described as acknowledged auth gaps in some sections.

Canonical interim rule for this cleaned AR:

```text
NEEDS REVIEW: Confirm deployed `saveParcelPolygon` authorization state before final production handoff.
```

Until verified:

- treat parcel polygon writes as security-sensitive master-data mutation;
- prefer Management-only authenticated write behavior;
- protect deployment URLs operationally if any public/pre-auth write surface remains deployed;
- keep public geo/meteo reads read-only and explicitly documented if retained;
- do not add new public write endpoints.

### 16.12 Input Validation and Sanitization

All write paths require domain validation before persistence.

Rules:

- PWA forms validate required fields, numeric ranges and domain selections before local save;
- VBA business modules validate required inputs and raise/fail rather than relying on UI prompts as control flow;
- GAS validates payload shapes and `records[]` arrays for batch sync;
- numeric/date/business identifiers must be normalized before write;
- user-facing dynamic HTML must be escaped;
- backend logging must truncate long messages/details.

### 16.13 Security Gates Summary

Detailed security checks live in `RELEASE_GATES.md`. Current mandatory coverage includes:

- SEF `http://` rejection;
- GAS 401/403 authz matrix smoke;
- sync entity-mismatch denial;
- Management-only shared write checks;
- monitoring secret/redaction smoke;
- PWA CSP/app-shell asset smoke;
- `localStorage` shared-state audit;
- `saveParcelPolygon` deployed-state review.

---

## 17. Reports and Derived Views

Reports and derived views are read models. They help operators, management and field users understand current state, but they are not the canonical transaction source unless explicitly stated.

Canonical transaction facts remain in their owning tables and flows, such as `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke`, `tblNovac`, `tblBankaImport`, `tblMagacin`, `tblAmbalaza`, role-specific Google transport sheets and monitoring workbook event tables.

Derived views may be rebuilt, refreshed, exported or cached. A derived view must not be used to silently repair or overwrite canonical transaction state.

| Report / View | Purpose | Source data | Calculation owner | Refresh trigger | Derived or canonical | Caveats |
|---|---|---|---|---|---|---|
| `SaldoOM` | station-level open balance | `tblNovac`, `tblOtkup`, agro/OM flows | VBA/export | export or report run | derived | avans/edge cases must stay regression-covered |
| `SaldoOMDetail` | detailed station saldo projection for PWA/management | desktop finance/export model | VBA export + GAS/PWA read | export + PWA refresh | derived export | not the canonical payment ledger |
| `SaldoKupci` | buyer balance | `tblFakture`, `tblFakturaStavke`, `tblNovac` | VBA/export | export or report run | derived | depends on correct faktura/novac links |
| `OtkupPoOM` | procurement by otkupno mesto | `tblOtkup` | VBA/export | export or report run | derived | aggregation-only view |
| `PredatoPoKupcu` | delivered goods by buyer | `Zbirna` / `Prijemnica` / `Faktura` chain | VBA/export | export or report run | derived | depends on complete document chain |
| `Kartica Kooperanta` / `Kartice` | per-kooperant financial card | `tblOtkup`, `tblNovac`, kartice export | VBA/export + PWA reader | export run / PWA cache refresh | derived export | `UKUPNO` summary rows are ignored by production parsing |
| `BankaImport` open queue | bank reconciliation work queue | `tblBankaImport`, `tblPartnerMap`, linked `tblNovac`/document IDs | VBA | import/map/skip/storno actions | staged facts + derived queue state | raw bank facts are staged facts; UI status is workflow state |
| `MgmtReports` | management reporting bundle | desktop reports and exports | VBA export + GAS/PWA read | export run / PWA refresh | derived | not a transaction write source |
| Management KPI dashboard | management quick overview | `getMgmtAll`, `MgmtReports`, `SaldoOMDetail`, dispatch runtime | GAS/PWA Management | bootstrap, refresh, sync/export | derived | cache freshness and dispatch state matter |
| Dispatch board | planning board | demand, kamion status, `DispecerPlan` runtime | PWA Management + GAS | management write/refresh | operational derived/planning state | Management may plan via `DispecerPlan` only; assigning `VozacID` to `OTK-*` rows directly from dispatcher is prohibited; displayed unallocated supply is raw otkupi minus planned kg per station (display-only subtraction, no record mutation) |
| Kooperant home / knjiga polja | kooperant operational and agronomy summary | treatments, expenses, agrohemija stock, production data | PWA Kooperant + GAS reads | role bootstrap/refresh/sync | derived | local pending records must remain visible until synced |
| Otkupac today overview | field-day overview | local IndexedDB + server `getOtkupi` data | PWA Otkupac | save/sync/refresh | derived | render must pass through canonical dedupe helper |
| Vozač pregled zbirne | driver transport overview | local/server `zbirne`, `getVozacZbirne` | PWA Vozač | save/sync/refresh | derived | `BrojZbirne` is business number; `ServerRecordID` is technical sync ID |
| Vozač planovi | driver plan view | `DispecerPlan` via `getVozacPlans` | PWA Vozač + GAS | transport tab load | derived | plans scoped to authenticated Vozac's `entityID`; today only; `zavrseno` excluded |
| `MeteoLatest` | current parcel meteo state | scheduled fetch + parcel config | GAS | scheduled fetch + stale fallback | canonical current meteo read model | rate limits/offline map caveats remain |
| Monitoring `Health` | current component health | monitoring events and watchdog checks | GAS Monitoring | event ingest + watchdog | derived current status | not a replacement for business tables |
| Monitoring `Events` / `Errors` / `AuditCritical` | operational event history | VBA/PWA/GAS monitoring payloads | GAS Monitoring | event ingest | append-only monitoring facts | redacted; does not store full sensitive payloads |

### 17.1 Management Reports

Management consumes overview, dashboard, dispatch, partner, saldo and agro views primarily through `getMgmtAll`, `MgmtReports`, `SaldoOMDetail` and live dispatch runtime state hydrated into Management shell state.

Management report views are allowed to aggregate and cache, but they must not become a hidden transaction source. Writes that affect canonical business state still go through the owning modules/endpoints:

- dispatch and demand writes through Management/GAS dispatch surfaces;
- agrohemija issuing through the Management agrohemija flow and `saveIzdavanje` boundary;
- bank mapping through `modBankaMapiranje`;
- document/finance corrections through desktop transaction/storno flows.

Management KPI helpers must be defensive against dirty dates, blanks, non-numeric values and worksheet errors. KPI refresh must skip invalid rows or resolve them to safe zero/placeholder values instead of raising type mismatch errors in the operator shell.

### 17.2 Finance Reports

Finance read models include:

- `Kartica Kooperanta`;
- `SaldoOM`;
- `SaldoOMDetail`;
- `SaldoKupci`;
- BankaImport reconciliation candidates and open staging queue;
- faktura payment-status views;
- agrohemija debt projections where outbound warehouse value is linked to a kooperant.

Finance reports are derived from canonical ledgers and transaction tables. They must not replace:

- `tblNovac` as the money/payment ledger;
- `tblBankaImport` as bank-staged source facts;
- `tblFakture` / `tblFakturaStavke` as invoice facts;
- `tblOtkup` as procurement facts;
- `tblMagacin` as warehouse/agrohemija journal.

Report code that reads schema-critical fields must use fail-fast schema guards such as `RequireColumnIndex` where the source module contract requires them. Optional report fields must be explicit optional branches, not silent zero-index reads.

### 17.3 Kooperant Views

Kooperant views combine local/offline records and server/exported state:

- treatment history;
- expense history;
- parcel overview;
- agrohemija stock/usage views;
- knjiga polja summaries;
- financial/card views exposed through exported read models.

Pending, syncing, failed and locally edited records must remain visible according to the PWA sync/render rules. Kooperant views must use canonical dedupe before render when they merge local and server data.

Knjiga polja result semantics remain derived:

```text
result = proizvodnja - agrohemija - troškovi
```

The derived result must not be treated as a canonical payment or accounting transaction.

### 17.4 Operational Dashboards

Operational dashboards include:

- Management dashboard with KPIs, alerts and dispatch visibility;
- Otkupac today overview;
- Kooperant home dashboard;
- Vozač zbirna/transport overview;
- Excel operator/sidebar KPI widgets.

Dashboard rules:

- dashboards read from canonical tables, synced/exported read models, IndexedDB caches or GAS bundles;
- dashboards may cache, aggregate and format;
- dashboards must not silently mutate canonical source rows;
- stale cache and offline state must be visible through sync badges, queue diagnostics or clear UI status;
- dirty worksheet data must not crash desktop KPI refresh.

`frmOtkupAPP` KPI helpers must use safe date/number parsing patterns equivalent to `SafeDateKey` and `SafeKpiDouble` so invalid dates, blanks, Excel errors and non-numeric kg values do not produce repeated runtime errors.

### 17.5 Derived-Data Ownership

Derived-data ownership rules:

1. The module that owns the source facts owns the correctness of the base transaction data.
2. The export/report layer owns aggregation, filtering and presentation shape.
3. PWA role views own local render ordering, dedupe, offline visibility and display formatting.
4. GAS owns API response packaging and transport read models, not hidden accounting corrections.
5. Monitoring owns operational health/read models, not business transaction state.

Derived read models may be regenerated from source facts. If a derived view conflicts with canonical source tables, the source tables win and the derived view must be refreshed or repaired.

### 17.6 Materialized / Exported Views

Active materialized/exported views include:

- `SaldoOMDetail`;
- `MgmtReports`;
- `Kartice`;
- `MeteoLatest`;
- role-specific Google transport sheets;
- PWA IndexedDB caches;
- monitoring workbook current-state tabs such as `Health`, `SEFStatus`, `Backups` and `SyncStatus`.

Materialized/exported views should be considered replaceable projections unless their owning section explicitly defines them as canonical facts. `MeteoLatest` is the current canonical meteo read model, but its history remains in `MeteoHistory` and parcel identity remains in `tblParcele` / Google `Parcele` projection.

### 17.7 Report Safety Rules

Report/read-model code must follow these safety rules:

- exclude stornirano rows unless the report explicitly says it is an audit/history report;
- keep technical IDs and business document numbers distinct;
- avoid using derived report totals as write-back inputs;
- keep `UKUPNO`/summary rows out of production data parsing where specified;
- treat missing required source columns as schema errors, not empty reports;
- surface freshness/staleness where the view depends on export, sync or scheduled fetch;
- keep monitoring and dashboards non-blocking relative to business transactions.

---

## 18. Current Production Gates

Detailed operational procedures live in `RELEASE_GATES.md`. This section is the canonical AR summary of what must be true before a code release or a specific production workbook/deployment is accepted.

Production readiness is evaluated at two levels:

1. **Code/deployment readiness** — compile, smoke, role, API, sync, security and monitoring gates pass for the changed surfaces.
2. **Workbook/tenant readiness** — the specific production workbook and Google deployment are clean, configured, and pass `RunProductionHealthCheck` with no blocking failures.

A release may be code-ready while a specific workbook is not production-ready. Legacy/demo/test data must be removed, repaired or stornirano-marked before declaring a workbook ready.

### 18.1 Compile Gates

Required:

- touched VBA modules compile without missing procedure, missing constant, reserved-name or visibility errors;
- exported `.bas`, `.cls` and `.frm` files preserve their `Attribute VB_Name` headers;
- shared guard/helper modules required by hardened code are available;
- form event handlers do not reference deleted modules, renamed controls or stale form procedures;
- GAS changed files pass Apps Script syntax/deployment checks;
- PWA changed files load without app-shell/runtime errors after service-worker cache refresh.

### 18.2 Business Flow Gates

Required:

- end-to-end document-chain regression covers `Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF` where the release touches those surfaces;
- dual-class document flows rollback atomically on forced failure;
- `PrijemnicaID` remains row-unique while `BrojPrijemnice` groups class rows;
- active document reads exclude stornirano rows unless the read is explicitly historical/audit-oriented;
- `RunBusinessFlowProSuite` or its successor passes for desktop document-chain releases;
- `RunFakturaSmokeSuite` or its successor passes for faktura/prijemnica/SEF-adjacent releases.

### 18.3 Storno Gates

Required:

- ID-based storno rejects missing target rows;
- ID-based storno rejects duplicate target keys;
- already-stornirano targets are rejected by `CanStorno`;
- all required writes use checked update semantics;
- `_TX` wrappers rollback hard failures;
- expected side effects are verified for `StornoOtkup`, `StornoOtpremnica`, `StornoZbirna`, `StornoPrijemnica`, `StornoFaktura` and `StornoNovac`.

### 18.4 BankaImport / BankaMapiranje Gates

Required:

- valid PDF extraction succeeds with configured `PDFTOTEXT_EXE_PATH`;
- invalid/missing `pdftotext.exe`, non-zero exit code or missing temp output fails the import path;
- statement saldo integrity gates reject incomplete or inconsistent bank statements before staging;
- missing required `tblBankaImport` columns fail fast;
- `AppendRow <= 0` rolls back staging;
- successful PDFs move to `Processed` only after DB commit;
- duplicate or missing `BankaImportID`, `NovacID`, `OtkupID` and `FakturaID` cannot pass mapping;
- stornirani BankaImport rows cannot be mapped or skipped as active rows;
- `GetBankaImportRowByID` preserves the legacy 1x10 business-shape contract.

### 18.5 SEF Gates

Required:

- `SEF_BASE_URL` is HTTPS-only and rejects `http://` locally;
- SEF parser smoke covers the accepted current parser baseline;
- submit, accepted/rejected, status refresh, stuck recovery and manual-review paths remain smoke-covered when SEF code changes;
- SEF state persistence remains in `tblSEFSubmission` / `tblSEFEventLog` and is not replaced by monitoring;
- SEF monitoring events are emitted best-effort and never decide business state.

### 18.6 GAS Gates

Required:

- `runGasRouteHealthCheck()` or successor confirms all active actions have real handlers;
- `runGasSmokeSuite()` or successor confirms auth, batch response semantics, schema guards, normalized lookup and disabled endpoint behavior;
- role/entity authorization is checked for all write endpoints touched by the release;
- `withLock(...)` protects sync/write paths where required;
- schema drift is fail-fast and does not append missing critical columns silently;
- `getMasterSyncState` and write-blocking behavior remain valid when MasterSync surfaces change.

### 18.7 Google Sheets Gates

Required:

- `Stammdaten` workbook and required tabs exist;
- OTK/VOZ/per-kooperant sheet headers match canonical contracts;
- `SyncStatus`, `ClientRecordID` and `ServerRecordID` semantics remain intact;
- VOZ writeback preserves the B/T split: technical `ServerRecordID` / `ZbirnaID` in column B and business `BrojZbirne` in column T;
- `SyncControl / MASTER_SYNC_LOCK` is readable and writable;
- Google writebacks are treated as external side effects and do not replace local transaction rollback.
- full-tab Google writes use staging/verify/replace and do not clear the target before a verified replacement is ready;
- Google write helpers use sheetId caching, quota-aware retry/throttle and phased target replacement;
- Kartice export writes to named tab `Kartice`, not `Sheet1`;

### 18.8 PWA Role Smoke Gates

Required:

- Otkupac smoke covers save, local queue, sync, overview/render and post-save sync trigger;
- Kooperant smoke covers tretmani, troškovi, local stock validation where applicable and sync result handling;
- Vozač smoke covers zbirna creation, PWA-first `BrojZbirne`, local pending/synced state and printability before MasterSync;
- Management smoke covers dashboard navigation, dispatch/agro surfaces where changed and no-sync role behavior;
- offline recovery covers stale `syncing` records, request failures returning attempted records to `pending`, and render dedupe.
- full Google/PWA sync is not green if the final PWA unlock fails; unlock failure is degraded/partial because TTL recovery may recover later but the operator result must not show success;

### 18.9 Monitoring Gates

Required:

- `MONITORING_ENDPOINT`, `MONITORING_SECRET`, `MONITORING_ENV`, `MONITORING_SPREADSHEET_ID`, `MONITORING_ALERT_EMAIL` and `MONITORING_INGEST_SECRET` are configured where monitoring is in scope;
- `TestMonitoring_All` or successor verifies config, HTTP ingest, error event, SEF unknown, backup success and backup failure paths;
- monitoring workbook tabs exist with required headers;
- `runMonitoringWatchdog()` is installed on the expected schedule;
- `sendDailyMonitoringSummary` is installed where daily summaries are required;
- monitoring redaction prevents tokens, secrets, SEF keys, raw UBL/PDF/base64 and full raw SEF responses from being logged.

### 18.10 Security and Compliance Gates

Required:

- SEF HTTPS-only validation passes;
- monitoring redaction checks pass;
- PWA CSP/self-hosted asset rules are respected;
- local state audit confirms shared state is not stored in `localStorage`;
- GAS role/entity authorization matrix is verified for changed endpoints;
- `saveParcelPolygon` deployed authorization state is explicitly confirmed before production handoff.

### 18.11 Data / Report / Derived View Gates

Required:

- required schema guards fail fast on missing canonical columns;
- derived views such as `MgmtReports`, `SaldoOMDetail`, `Kartice`, `MeteoLatest`, PWA caches and monitoring tabs are not treated as write-back source-of-truth;
- report helpers tolerate dirty or partial production data where the architecture requires robust dashboards;
- demo/test rows are cleaned, repaired or stornirano-marked before production readiness is declared.

### 18.12 ProductionHealthCheck

Required:

- `RunProductionHealthCheck` completes with no blocking failures before a workbook is declared production-launch-ready;
- duplicate identity keys are checked for `OtkupID`, `OtpremnicaID`, `ZbirnaID`, `PrijemnicaID`, `FakturaID`, `NovacID`, `BankaImportID` and `ParcelaID` before operator workflows are accepted;
- any remaining warnings are classified as accepted operational risk or roadmap technical debt;
- unresolved `NEEDS REVIEW` items are not silently ignored;
- the final release package records which gates were run, which were deferred, and who accepted any remaining risks.


---

### 18.13 v6.22 Residual Hardening Gates

Required for the v6.22 documentation/code closeout surfaces:

- `PrintFaktura` duplicate-`FakturaID` guard smoke: duplicate faktura keys must fail before print selection/output.
- `UpdateFakturaStatus` duplicate-`FakturaID` guard smoke: duplicate faktura keys must fail before payment/status recompute.
- `SaveParcelGeoPointByID[_TX]` smoke: selected `ParcelaID` resolves exactly one physical `tblParcele` row and updates point/lat/lng/status fields transactionally.
- `ClearParcelGeoByID[_TX]` smoke: selected `ParcelaID` resolves exactly one row and clears point fields transactionally without clearing polygon fields.
- `RequireSingleParcelaRow` negative smoke: missing and duplicate `ParcelaID` values fail hard.
- Row-index geo wrappers, where retained, delegate to ID-based APIs and are not used as the preferred architecture.
- Storno eligibility checks use the `RequireStornoAllowed` / exact-row / checked-write pattern before mutation.

### 18.14 v6.23 PWA Otkup Read-Model Gates

Required for the v6.23 PWA otkup read-model convergence:

- `MgmtReports/OtkupiAll` export exists and is readable by GAS/PWA;
- Management sees otkupi entered directly in VBA/master;
- Management sees PWA-created/synced otkupi;
- Otkupac sees VBA/master-created otkupi within its role/station scope;
- Otkupac sees operational `OTK-ST-*` / `OTK-*` rows;
- merged display uses `ServerRecordID` / `OtkupID` before `ClientRecordID`;
- the same synced PWA otkup present in both `OTK-ST-*` and `OtkupiAll` renders once;
- operational rows do not disappear merely because `OtkupiAll` is present;
- browser smoke for Management and Otkupac read paths is recorded.


### 18.15 v6.24 PWA Design/Runtime and Numbering Gates

Required v6.24 documentation/implementation verification gates:

- confirm the target AgriX repo contains the current `base.css`, `fonts.css`, `components_v2.css`, role UI feature files, `sw.js`, `format.js` and `lazy.js`;
- confirm `CACHE_NAME` / service-worker cache version was bumped with the redesigned app shell;
- confirm self-hosted Cormorant/DM Sans font assets and Latin-ext coverage are deployed;
- confirm Otkup form uses event-delegated picker handlers and no inline script path;
- confirm Otkup save, Otkupni list modal, Otprema and Pregled flows work in browser after cache refresh;
- confirm `parseDecimalInput` handles Serbian comma decimal input in all money/quantity fields where used;
- confirm lazy-loaded `jsPDF`, Leaflet, Chart.js and Firebase paths load only when their feature needs them;
- confirm VBA document numbering helpers and lock-based per-station behavior are present in the real codebase;
- confirm `modGoogleSheets` `sheetId = 0` sentinel collision diagnosis was either patched or explicitly left as open issue.

Because the accessible GitHub connector repository does not match the AgriX/OtkupApp frontend, this gate remains source-summary verified and requires target-repo confirmation before production signoff.

## 19. Current Known Risks

This section summarizes current risks that are still relevant to the active architecture. The maintained issue register lives in `KNOWN_ISSUES.md`; historical fixed issues stay in `ARCHITECTURE_CHANGELOG.md` or archived snapshots.

### 19.1 Launch-Relevant Risks

- `saveParcelPolygon` authorization state is inconsistent across source documentation and must be verified against deployed `Code.gs` before production handoff.
- A specific production workbook is not launch-ready until `RunProductionHealthCheck` has no blocking failures, even if compile/smoke gates pass.
- Monitoring is best-effort and must not be interpreted as the canonical state machine for SEF, finance, BankaImport or document-chain operations.
- Google Sheets writebacks and file-system moves are external side effects and are not transactionally rollback-able with Excel table snapshots.

### 19.2 Accepted Operational Risks

- PWA-first `BrojZbirne` generation accepts the operational assumption of one active device per driver; multi-device same-driver offline collision remains a known edge case until optional GAS duplicate guard is implemented.
- File moves after BankaImport DB commit require manual recovery if the post-commit move fails.
- Treatment sync currently records evidence/history/karenca and does not automatically decrement server-side `magacinkoop` stock.
- Monitoring endpoint downtime must not block business transactions.

### 19.3 Needs Review

- Confirm row-index parcel geo wrappers are compatibility-only in exported VBA and all UI save/clear paths call the ParcelaID-based API.

- Canonical module name for Banka PDF parser: `modBankaImport_PdfText` vs `modBankaImportParserPdfToText`.
- `saveParcelPolygon` security decision: confirm whether token lock is closed or still accepted risk.
- Confirm v6.18/v6.19 active rules are represented before removing archived appendix copies from active AR.
- Confirm whether older changelog entries v2.2–v6.9 should remain compact in the main CL or live only in archive.
- Confirm whether endpoint tables remain in AR long-term or move to generated API docs.

### 19.4 Non-Launch Technical Debt

- Shared exact-row/data-access guards should be consolidated into a common module.
- SEF manual JSON parser remains a known limitation and should be migrated through a controlled VBA-JSON wrapper pass.
- Some residual business-layer `MsgBox` usage remains outside hardened blocks and should move toward UI-owned messaging.
- Test automation can be strengthened around GAS fixture sync, PWA dedupe/stale recovery, BankaImport negative paths and Agrohemija basket rollback.


---

## 20. Current Roadmap Summary

Full roadmap tracking lives in `ROADMAP.md`. AR keeps only the current architecture-facing summary.

### 20.1 Shared Data-Access Guard Consolidation

Move local exact-row helpers such as `RequireSingleRow` into a shared guard module when multiple finance/document modules converge on the same pattern.

### 20.2 SEF / JSON Parser Hardening

Replace manual SEF JSON extraction through a controlled parser-wrapper migration and regression suite.

### 20.3 GAS Idempotency Improvements

Add optional defensive guards for edge cases such as duplicate `BrojZbirne` for the same `VozacID` and server-side idempotency for management `saveIzdavanje`.

### 20.4 PWA Test Automation

Add minimal repeatable tests for sync result normalization, stale `syncing` recovery, render dedupe, submit locks and business-date helpers.

### 20.5 UI / Business Separation

Continue moving operator prompts and `MsgBox`-controlled branches out of business modules and into forms/operator shells.

### 20.6 Schema and Header Registry Hardening

Audit canonical Excel and Google Sheets headers, especially OTK/VOZ/per-kooperant transport sheets and finance/document tables.

### 20.7 Monitoring Noise Tuning

After pilot usage, tune event volume so monitoring keeps critical operational visibility without noisy helper-level logging.


---

## 21. Deprecated and Transitional Elements

This section lists patterns that still exist for compatibility or historical reasons but are not the preferred architecture. These entries must not be treated as permission to introduce more legacy behavior.

### 21.1 Deprecated Patterns

| Pattern | Status | Replacement / current rule |
|---|---|---|
| Version-delta blocks inside active AR | Deprecated documentation pattern | Current architecture belongs in domain sections; deltas belong in `ARCHITECTURE_CHANGELOG.md`. |
| Direct `FindRows(...).Count > 0` for critical IDs | Deprecated data-access pattern | Use exact-row guards: `Count = 1`, missing and duplicate are hard errors. |
| Silent schema fallback / column index `0` | Deprecated schema pattern | Required columns use `RequireColumnIndex` / `RequireColumns`; optional columns are explicit. |
| Business-layer control flow through `MsgBox` | Deprecated module pattern | Business modules raise errors/return results; forms decide operator messaging. |
| Broad `On Error Resume Next` in business/backup paths | Deprecated EH pattern | Structured EH with preserved original error context. |
| `Date.toISOString().slice(0, 10)` for PWA business dates | Deprecated PWA date pattern | Use local-calendar helpers such as `getTodayIsoDate()`, `toIsoDateOnly(...)`, `localIsoDateFromDate(...)`. |
| Direct low-level PWA sync calls from UI triggers | Deprecated sync pattern | Use `syncQueueSafe(reason)` / role request wrappers. |
| `localStorage` for shared operational state | Deprecated state pattern | Shared state belongs in IndexedDB + GAS/Sheets or canonical desktop tables. |
| Static `%TEMP%\pdf_extract.txt` for bank PDF extract | Removed/deprecated parser pattern | Use unique temp txt path per extraction and defensive cleanup. |
| User-specific hardcoded `pdftotext.exe` path | Removed/deprecated setup pattern | Use `tblLocalConfig.PDFTOTEXT_EXE_PATH` with setup fallback. |
| Row-index parcel geo save/clear API | Transitional compatibility surface | Use `SaveParcelGeoPointByID[_TX]`, `ClearParcelGeoByID[_TX]` and `RequireSingleParcelaRow`. |
| `ServerRecordID` as `BrojZbirne` | Deprecated identity pattern | `ServerRecordID` is technical; `BrojZbirne` is a business number. |
| Hardcoded SEF HTTP / `http://` endpoints | Forbidden config pattern | `SEF_BASE_URL` must be HTTPS. |
| Hardcoded 10% VAT in SEF mapper totals | Deprecated mapper pattern | Use configured tax source and fail-fast total consistency checks. |
| Helper duplication for URL/JSON escaping | Deprecated utility pattern | Use canonical `modHttpUtils.UrlEncode` and `modHttpUtils.JsonEscape`. |

### 21.2 Transitional Compatibility Surfaces

| Surface | Why it remains | Preferred future direction |
|---|---|---|
| `GenerateBrojZbirne(vozacID, datum)` | Fallback for legacy/pre-rollout VOZ rows and desktop manual entry | PWA-first generation at `confirmZbirna` remains primary. |
| `runRoleSync(reason)` | Compatibility alias | Delegate only to `requestRoleSync(reason || 'manual')`; do not reintroduce direct routing. |
| `ReportStanjePoDoabvljacu()` | Old misspelled report function may exist for callers | New code uses `ReportStanjePoDobavljacu()`. |
| `GetUplataForOtkup()` | Historical compatibility name | Prefer clearer `GetIsplataForOtkup()` in new code when the helper sums linked `Isplata`. |
| Local `RequireSingleRow` in `modBankaMapiranje` | Accepted v6.21 local helper | Consolidate into shared `modDataAccessGuards`. |
| Lightweight manual SEF JSON parser | Active but limited parser baseline | Controlled VBA-JSON wrapper migration with regression tests. |
| Public/read geo/meteo bridge | Supports current frontend request model | Decide long-term auth/read policy and document it. |
| `saveParcelPolygon` authorization state | Source docs conflict | `NEEDS REVIEW`: verify deployed `Code.gs` and update AR/security matrix. |
| Dev smoke modules (`modNovacTests`, `modSEFTests`, `modFakturaTests`, `modBusinessFlowProTests`) | Engineering regression value | Keep available to engineering, inaccessible from normal operator UI. |

### 21.3 Disabled / Removal Candidates

- temporary Kartice/`Sheet1` fallback helpers such as `GetFirstSheetId` and `ResolveTargetSheetIdForReplace` are removed/obsolete after Kartice moved to the named `Kartice` tab.
- `GoogleRetryDelayMs` and `ExtractSheetIdByTitle` appear to be non-blocking cleanup candidates after quota-aware retry and sheetId cache consolidation; remove only after compile/search confirms no callers.

| Item | Current behavior | Removal / replacement rule |
|---|---|---|
| `saveOtkupniListPdf` | Returns `FEATURE_DISABLED` until implemented | Do not list as active route-health handler. |
| Legacy disabled `syncTrosak` wording | Superseded | `syncTrosak` is active in current architecture. |
| Old v6.18/v6.19 full snapshots inside AR | Archived | Keep in `docs/archive/`, not active AR. |
| Legacy CL entries v2.2-v6.17 in active CL body | Compressed/archived | Active CL keeps compact summary and archive pointer. |
| Old `KolAmbalazeVracena` spelling | Non-canonical | Use `tblPrijemnica.kolAmbVracena`. |

---

## 22. Glossary

| Term | Meaning |
|---|---|
| AgriX / OtkupApp | Product/system for Serbian fruit procurement, field workflows, desktop document/finance backbone and farm-management support. |
| Otkup | Field procurement/purchase record from kooperant/parcel/culture/class/quantity/price context. |
| Otkupac | Field buyer/procurement role that creates otkup records and may assign drivers where allowed. |
| Kooperant | Producer/farmer role; owns parcel/treatment/expense/fiscal records in scoped PWA flows. |
| Vozac | Driver role; owns zbirna creation/transport overview for assigned otkup rows. |
| Management | Management PWA role for dashboards, dispatch, partner views, agro issuing and shared operational oversight. |
| Excel operator | Backoffice/admin user of the Excel/VBA desktop master. |
| Otpremnica | Dispatch/delivery document created before or around transport grouping; may initially have empty `BrojZbirne`. |
| Zbirna | Driver/transport aggregation document. Business number is `BrojZbirne`. |
| BrojZbirne | Business document number with format `x/ddmmyy[-rb]`, generated PWA-first with VBA fallback. |
| Prijemnica | Receipt/intake document at buyer/receiving side. |
| BrojPrijemnice | Business prijemnica number that may group multiple class rows. |
| PrijemnicaID | Unique physical row ID in `tblPrijemnica`; faktura creation uses this row identity. |
| Faktura | Invoice generated from prijemnica/document flow and optionally submitted to SEF. |
| SEF | Serbian e-invoice system/API integration. |
| SEFWorkflowState | Internal/local workflow state for SEF processing. |
| SEFStatus | Exact latest external SEF API status persisted locally. |
| Storno | Soft-cancel/reversal operation that marks rows as stornirano and performs business-specific repair/unlink side effects. |
| Stornirano | Soft-delete/cancel marker; active reads normally exclude `Stornirano = "Da"`. |
| Ambalaža | Packaging/containers tracked through `tblAmbalaza` movements. |
| `Ulaz` / `Izlaz` | Canonical ambalaža ledger directions; unknown directions are invalid. |
| Novac | Money ledger domain/table covering payments, payouts, avans and OM station movement. |
| Avans | Advance/prepayment amount later allocated to faktura or otkup. |
| BankaImport | Bank statement import/staging pipeline based on PDF text extraction and integrity checks. |
| BankaMapiranje | Reconciliation layer that maps staged bank rows to novac, faktura, otkup, OM or partner-map records. |
| PartnerMap | Learned bank partner-name mapping to internal buyer/kooperant/OM entities. |
| Stammdaten | Google Sheets exported master-data workbook used by PWA/GAS. |
| GAS | Google Apps Script API layer for auth, routing, sync, dispatch, meteo, fiscal and monitoring ingress. |
| PWA | Offline-first frontend for Otkupac, Kooperant, Vozac and Management. |
| IndexedDB | Browser local-first persistent queue/cache for PWA records. |
| `ClientRecordID` | Stable PWA-side idempotency key for retries. |
| `ServerRecordID` | Technical GAS/PWA server-side sync identifier; not a business number. |
| `SyncStatus` | Google/PWA sync lifecycle field such as pending/synced/master/error states. |
| `Synced>Master` | Transport row state indicating desktop master import/writeback has processed the record. |
| `SyncControl` | Stammdaten sheet/tab used for master-sync lock/readout state. |
| `MASTER_SYNC_LOCK` | Lock record used to block/soft-lock PWA writes during desktop master sync. |
| `RequireColumnIndex` | Fail-fast schema helper for required columns. |
| `RequireUpdateCell` | Checked write helper; critical updates must fail if the write does not land. |
| `_TX` wrapper | Transaction wrapper that snapshots affected Excel tables, commits or rolls back and emits monitoring where applicable. |
| `ProductionHealthCheck` | Desktop workbook launch gate for data/config/schema health. |
| `OtkupApp_Monitoring_PROD` | Monitoring Google Sheets workbook with health/events/errors/alerts/audit views. |
| `ErrorLog` | GAS/PWA runtime error log sheet. |
| `LocalConfig` / `tblLocalConfig` | Local workstation config table for machine-specific settings such as `PDFTOTEXT_EXE_PATH`. |
| `tblConfig` | PWA/Google config projection, not local workstation config. |
| `tblSEFConfig` | SEF and monitoring configuration table. |
| MeteoLatest | Current parcel meteo/risk read model populated by scheduled meteo fetch. |
| GeoSrbija / polygon editor | Parcel coordinate/polygon support surfaces. |
| Digitalni Agronom | Kooperant agronomy/treatment surface including treatments, dosage, karenca and evidence sync. |
| Agrohemija | Warehouse/chemical issuing and treatment-consumption domain. |

---

## 23. Revision Metadata

### 23.1 Current Version
v6.24 canonical snapshot.

### 23.2 Superseded Versions
Supersedes OtkupApp / AgriX reference versions v2.2.1–v6.23.

### 23.3 Companion Documents
- `ARCHITECTURE_CHANGELOG.md`
- `RELEASE_GATES.md`
- `ROADMAP.md`
- `KNOWN_ISSUES.md`

### 23.4 Archive References
- `archive/ARCHITECTURE_REFERENCE_v6_18.md`
- `archive/ARCHITECTURE_REFERENCE_v6_19.md`
- `archive/CHANGELOG_legacy_v2_to_v6_17.md`

### 23.5 v6.24 Closeout Scope
v6.24 documents the Vozač/Dispečer operational bugfix pass:

- `getVozacPlans` GAS action added to endpoint authorization matrix; Vozac-scoped, today-only, `zavrseno`-excluded.
- `DispecerPlan` schema contract made explicit in section 9.13.
- Dispatcher write invariant formalized: `OTK-*` records must not be mutated from the dispatcher path; `VozacID` assignment via dispatcher is prohibited.
- `dpGetSup()` display-only supply subtraction rule documented: unallocated supply = raw otkupi minus planned kg per station; no otkup records are mutated.
- PWA DOM ID uniqueness rule added to section 2.5: modal signature canvas IDs must not collide with static canvas IDs in `index.html`.
- Management dispatcher must load live operational otkupi (`includeLive=1`) on init to see today's quantities without requiring manual "Otkup uživo" navigation.
- Dispatch board report row updated to reflect `DispecerPlan` source and supply subtraction semantics.
- Vozač planovi report row added to section 17 report inventory.

### 23.6 v6.23 Closeout Scope
v6.23 documents the PWA otkup read-model convergence: `tblOtkup` remains canonical master, `MgmtReports/OtkupiAll` is the PWA master read projection, and `OTK-ST-*` / `OTK-*` remains the operational queue. Management and Otkupac views merge master projection plus operational queue rows with `ServerRecordID` / `OtkupID` before `ClientRecordID` dedupe. Browser smoke was reported as tested for the affected role views.

### 23.6 v6.22 Closeout Scope
v6.22 is a documentation/version cut that carries forward the v6.21 GO hardening closeout and adds residual current-contract items that were not explicit enough in the prior package:

- `modFaktura` duplicate-`FakturaID` guards for print/status paths;
- `modGeoParcele` ParcelaID-based save/clear APIs with rowIndex wrappers demoted to compatibility;
- explicit `RequireStornoAllowed` / checked-storno helper pattern naming;
- release gates for the above residual hardening.

No historical business-data migration is defined by this documentation cut.
