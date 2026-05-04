# AgriX / OtkupApp Architecture Reference

**Version:** v6.12 draft canonical snapshot  
**Last Updated:** 2026-05-04  
**Status:** Canonical / Active Reference Draft  
**Scope:** VBA/Excel desktop backend + desktop app lifecycle + Google Apps Script + Google Sheets + AgriX PWA + PWA-first otkup/traceability + document flow + finance + modNovac v6.8 hardening + frmSEF operator-shell hardening + modFaktura canonical-prijemnica hardening + modDokumenta EH/input/stornirano hardening + modOtkup validation/read-helper hardening + Google VBA/GAS v6.10 hardening + AR-002 transaction AutoSave + canonical HTTP utilities + SEF HTTPS/UTF-8 hardening + v6.11 data-health cleanup + Agrohemija/Stammdaten + GIS + meteo + dispatch + Knjiga Polja + Fiskalni + SEF P0/P1 hardening + tested SEF live baseline + SEF v6.7 state/parser/total-consistency hardening + professional regression suites + strict BrojZbirne trace bridge + shell/update-guard convergence + PWA launch-smoke hardening + canonical sync-result convergence + business-date cleanup + client error reporting + syncTrosak activation + VOZ/BrojZbirne post-VBA ownership clarification  
**Owner:** Architecture documentation compiled from supplied reference set  
**Supersedes:** OtkupApp / AgriX reference versions v2.2.1–v6.11 
**Audience:** Engineering, Product, Operations, Onboarding, Review  

---

## 0. Document Contract

### 0.1 Purpose
Ovaj dokument je **potpuni i samodovoljni snapshot** trenutno važeće arhitekture sistema na datum verzije.

### 0.2 Documentation Invariant
Every Architecture Reference version must be fully self-contained.  
It must restate all currently valid architecture elements, even if unchanged from previous versions.

U glavnom reference dokumentu **nisu dozvoljene** formulacije tipa:

- "unchanged from vX.Y"
- "same as previous"
- "others omitted"
- "see earlier version"
- "rest unchanged"

Takve napomene su dozvoljene samo u changelog-u, release notes ili appendix-u istorije verzija, ali **ne** u canonical reference dokumentu.

### 0.3 Reading Rule
Ako nešto nije navedeno u ovom dokumentu, ne smatra se canonical arhitekturom dok ne bude eksplicitno uneseno.

### 0.4 Versioning Rule
Svaka nova verzija reference dokumenta mora prepisati ceo trenutno važeći snapshot sistema, a ne samo razlike.

### 0.5 Scope Boundary
Ovde se dokumentuje:

- aktivna arhitektura
- aktivni sistemi i integracije
- važeći obrasci rada
- otvoreni arhitektonski problemi koji su i dalje relevantni
- važeći roadmap samo tamo gde utiče na arhitekturu

Ovaj dokument nije zamena za:

- detaljnu tehničku implementaciju po fajlovima
- task tracker
- release notes
- sprint plan

### 0.6 Editorial Rule
Reference mora biti čitljiv i upotrebljiv bez otvaranja bilo kog starijeg dokumenta.

---

## 1. System Overview

### 1.1 Product Scope
- **Naziv proizvoda:** AgriX / OtkupApp
- **Primarna namena:** digitalizacija srpskog lanca otkupa voća i paralelno farm-management sloja za kooperante
- **Glavni korisnici:** Otkupac, Kooperant, Vozač, Management, Excel operator
- **Operativni kontekst:** offline-first PWA na terenu + Google Sheets / GAS kao online transportni sloj + Excel/VBA kao desktop master i finansijsko-dokumentni backbone

### 1.2 High-Level Architecture
Sistem ima 4 glavna sloja:

- **VBA/Excel desktop backend** — master data management, dokumentni tok, finansije, SEF, BankaImport, izveštaji
- **Google Apps Script API** — action-router, auth, sync, meteo, dispatch, fiskalni i pomoćni endpoints
- **Google Sheets data layer** — stammdaten, role-specific sheets, kartice, meteo, dispatch, fiskalni i management report tabovi
- **AgriX PWA** — offline-first klijent sa 4 role-specific UI toka i IndexedDB lokalnim storage-om

## 1.6 v6.11 Pre-Launch Persistence, HTTP Security and Data-Health Delta

v6.11 is the active pre-launch hardening layer on top of the v6.10 Google/GAS sync baseline. It does not introduce a schema migration, but it changes the launch contract for persistence, outbound HTTP utility ownership and production-health readiness.

### 1.6.1 AR-002 Transaction AutoSave Contract

The active desktop transaction contract is now:

- `clsTransaction.CommitTx` is the central post-commit persistence hook.
- After successful commit cleanup, it calls `AutoSaveAfterCommit(sourceName)`.
- `sourceName` is built from the transaction snapshot table names before snapshots are cleaned up, for example `clsTransaction[tblOtkup,tblAmbalaza,tblNovac]`.
- AutoSave is best-effort and must not raise back to the business caller because the transaction has already committed in memory.
- AutoSave suppresses `Application.DisplayAlerts` during `ThisWorkbook.Save` and restores it on every exit path.
- AutoSave guards read-only workbooks and unsaved/no-path workbooks.
- AutoSave uses a debounce window to avoid rapid-fire saves during clustered transactions.
- AutoSave logging distinguishes actual saves from debounce skips.

This means a successful production transaction is no longer allowed to depend on a later manual operator save for persistence.

### 1.6.2 Canonical HTTP Utility Ownership

`modHttpUtils` is now the canonical owner for desktop outbound request helper functions:

- `UrlEncode(ByVal s As String) As String`
- `JsonEscape(ByVal s As String) As String`

The former duplicated helper surfaces are deprecated/removed:

- `modGoogleAuth.UrlEncodeGoogle`
- `modGoogleSheets.JsonEscapeGoogle`
- private `modSEFClient.UrlEncode`
- private `modSEFClient.JsonEscape`
- local Google/MasterSync URL encoding helpers where replaced by the canonical helper

`UrlEncode` follows RFC 3986 unreserved-character behavior and expands non-ASCII input into UTF-8 bytes before percent-encoding. This prevents Serbian characters from being encoded as ANSI/codepage bytes.

`JsonEscape` remains a basic outbound string-literal escaper and intentionally does not replace JSON parsing. JSON parser consolidation remains a separate VBA-JSON migration item.

### 1.6.3 SEF HTTPS-Only Rule

SEF endpoint configuration is now fail-closed:

- `SEF_BASE_URL` must start with `https://`.
- `http://` is rejected locally with `ERR_SEF_CONFIG`.
- This rule is enforced in both client/config validation paths so an API key and invoice payload are never sent over plaintext HTTP.

### 1.6.4 SEF Parser Baseline and Remaining JSON Risk

The manual SEF JSON parser remains active for v6.11, but its risk is now explicit:

- simple string, numeric/string ID and boolean cases are smoke-covered;
- escaped/nested/array/null-heavy payloads remain a known parser limitation;
- full replacement must be done through controlled VBA-JSON wrapper migration and parser regression tests.

### 1.6.5 Test Hygiene and ProductionHealthCheck Readiness

Regression/smoke tests must not leave active broken references in the workbook.

The Faktura already-fakturisana-prijemnica test no longer uses an active fake `FAK-EXISTING` reference that pollutes production health output. Legacy demo/test rows must be cleaned or stornirano-marked before declaring a workbook production-ready.

Canonical launch rule:

- Code release can be considered ready after compile + smoke/E2E pass.
- A specific workbook can be considered production-launch-ready only when `RunProductionHealthCheck` has no failures.

### 1.6.6 v6.11 Verification Evidence

Required v6.11 verification set:

- `RunHttpUtilsSmokeSuite` — UTF-8 `UrlEncode` and `JsonEscape` utility coverage.
- `RunSEFClientParserSmokeSuite` — baseline current-parser coverage.
- SEF negative config check — temporary `http://` `SEF_BASE_URL` must fail locally.
- `RunGoogleSyncSmokeSuite` — Google auth/sheets/drive transport still passes after central utility replacement.
- `RunMasterSyncSmokeSuite` — fixture import/writeback/idempotency still passes after central utility replacement and AutoSave hook.
- `RunFakturaSmokeSuite` — expected `18/18` pass and no new active `FAK-EXISTING` health pollution.
- `RunBusinessFlowProSuite` — expected `111/111` pass with AutoSave save/debounce evidence in logs.
- `RunProductionHealthCheck` — final launch gate for workbook data cleanliness.
- GAS checks remain the active v6.10 backend gates: `runGasRouteHealthCheck` and `runGasSmokeSuite`.



## 1.7 v6.12 PWA Launch-Smoke, Sync and Observability Delta

v6.12 is the active PWA launch-readiness hardening layer on top of the v6.11 desktop persistence/security baseline. It closes the concrete runtime defects found during role-by-role smoke testing and clarifies the cross-system boundary between PWA/GAS technical sync identifiers and VBA-owned business document numbering.

### 1.7.1 PWA Business-Date Contract

The active PWA date-only contract is now:

- business dates are represented as canonical `YYYY-MM-DD` strings;
- date-only fields such as `Datum`, treatment dates, expense dates, otkup dates, otprema dates, zbirna dates and agrohemija issue dates must be produced through local-calendar helpers, not by slicing UTC timestamps;
- `getTodayIsoDate()`, `getRelativeIsoDate(...)`, `toIsoDateOnly(...)`, `fmtDate(...)` and `localIsoDateFromDate(...)` are the canonical date helper surface in `src/js/utils/format.js`;
- `Date.prototype.toISOString().slice(0, 10)` and `toISOString().split('T')[0]` are not valid business-date patterns in PWA feature code;
- ISO timestamp fields such as `createdAtClient`, `updatedAtClient`, `updatedAtServer`, `syncAttemptAt`, `ReceivedAt` and `syncedAt` remain real UTC timestamps and may still use `new Date().toISOString()`.

Runtime smoke confirmed that `toIsoDateOnly('2026-05-01T22:00:00.000Z')` resolves to the local Serbian business date `2026-05-02`, preventing the prior one-day display drift in Pregled/Otprema-style views.

### 1.7.2 PWA Format Helper Contract

`src/js/utils/format.js` is also the canonical owner for simple display formatting helpers required by role views:

- `formatNumber(value, options)`;
- `formatKg(value)`;
- `formatMoney(value)`.

Feature renderers such as `otkup-pregled.js` must not assume local/private formatter globals. Shared display formatting belongs in the common format utility layer.

### 1.7.3 Shared Sync Engine and Canonical Role Result Contract

The active PWA sync contract is now canonical across Otkupac, Kooperant and Vozac roles.

All app-level role sync entrypoints must return this normalized top-level shape:

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

The shared sync engine owns these common runtime rules:

- offline, unavailable database and in-flight states return normalized results instead of ad-hoc values;
- request-level failures return every record from the current attempted batch to `pending`, not only records whose local object still says `syncing`;
- stale `syncing` recovery is age-gated and does not blindly recover fresh in-flight rows;
- successful backend statuses include `synced`, `duplicate`, `existing`, `inserted` and `updated`;
- missing backend per-record results are treated as non-terminal and revert the affected row to `pending` with diagnostics;
- record diagnostics include `lastSyncError`, `lastServerStatus`, `syncAttempts`, `syncAttemptAt`, `updatedAtServer` and `syncedAt` where applicable.

### 1.7.4 Kooperant Sync Scope

Kooperant sync is no longer treatment-only. The active `syncKooperantNow()` / role sync path covers:

- `syncTretmani()` for treatment/agromere records;
- `syncTroskovi()` for farm expense records.

The Kooperant top-level sync result aggregates both child sync results and remains canonical even when both stores return `reason: 'no-pending'`.

### 1.7.5 `syncTrosak` GAS Contract

`syncTrosak` is now an active GAS endpoint and is no longer a disabled launch placeholder.

The active backend contract is:

- action: `syncTrosak`;
- allowed roles: `Kooperant`, `Management`;
- Kooperant callers must match `tokenData.entityID === data.kooperantID`;
- `records` must be an array;
- processing is protected by `withLock(...)`;
- each row is handled by `processTrosakRecord(record, kooperantID)`;
- the endpoint must return `jsonResponse(withLock(function() { ... return buildBatchSyncResponse(results); }))`;
- an HTTP 200 response with empty JSON body is invalid and is treated by PWA as `empty-response`;
- `processTrosakRecord` must be idempotent by `ClientRecordID`, updating an existing row instead of appending duplicates.

`getTroskoviForKooperant(kooperantID)` reads the canonical `TROSKOVI-<KooperantID>` sheet and returns normalized records scoped to the requested kooperant.

### 1.7.6 Client Error Reporting Contract

PWA client-side observability is active.

The canonical frontend helper is `reportClientError(error, context)`, owned by `src/js/utils/async.js`. It reports to the GAS `logClientError` action and must be best-effort only. Logging failure must never break the app runtime.

The active reporting surfaces are:

- `safeAsync(...)` catch paths;
- sync-engine exception paths;
- bootstrap/app startup catch paths;
- global `window.error`;
- global `window.unhandledrejection`.

The payload is intentionally small and privacy-bounded:

- `errorAction`;
- message;
- stack/details truncated for transport;
- role;
- `entityID`;
- app version if configured;
- URL and user agent diagnostics.

GAS writes these into the existing `ErrorLog` sheet through `logError(...)`. The PWA smoke confirmed that a manual client error report creates an `ErrorLog` row and that the final role smoke did not add unexpected new rows.

### 1.7.7 PWA Smoke-Test Evidence

The v6.12 PWA launch smoke covered all role surfaces:

- **Otkupac:** new otkup save, sync, `getOtkupi` server read, Danas/Pregled render, Otprema render, date display and slash-based ambalaža display.
- **Kooperant:** treatment save/sync after sheet schema correction, expense save/sync after `syncTrosak` response contract fix, `getTroskovi` read and `kpFetchAll()` runtime visibility.
- **Vozac:** `getVozacOtkupi`, zbirna creation, `syncZbirna`, `getVozacZbirne`, local synced state, date and ambalaža display.
- **Management:** session sanity, `getMgmtAll`, management navigation and no unexpected `ErrorLog` rows.

Known smoke blockers found and closed during v6.12:

- missing `formatKg` / `formatMoney` helpers in Otkup Pregled rendering;
- stale/incorrect Google Sheet headers for treatment and expense test sheets;
- `syncTrosak` backend returning HTTP 200 with `data = null` because the batch response was not returned from the endpoint callback.

### 1.7.8 VOZ/Zbirna Numbering Boundary

The v6.12 cross-system rule for Zbirna identifiers is explicit:

- `ServerRecordID` is a technical PWA/GAS sync identifier;
- `BrojZbirne` is a business document number;
- PWA and GAS must not invent `BrojZbirne` by copying `ServerRecordID`;
- the business owner for `BrojZbirne` generation is VBA Master import.

The active VBA rule is implemented by `GenerateBrojZbirne(vozacID, datum)` in the VOZ/Zbirna import flow:

- extract numeric part from `VozacID`, e.g. `VOZ-00004` -> `4`;
- combine it with `Format$(datum, "ddmmyy")`, e.g. `4/040526`;
- append `-2`, `-3`, etc. for subsequent zbirne from the same driver on the same day.

`ImportOneVOZSheet` adds writeback updates as:

```vb
statusUpdates.Add Array(i, SYNC_STATUS_MASTER, newZbirnaID, brojZbirne)
```

Therefore `WriteBackVOZSyncStatus` must write:

- column B / `ServerRecordID` from `update(2)`;
- column T / `BrojZbirne` from `update(3)`.

Writing `update(2)` to both B and T is invalid because it stores the internal/master ID as the business document number.

### 1.7.9 v6.12 Launch Boundary

v6.12 does not require production-data migration. Test sheets and test records are not canonical production data and do not require repair/backfill documentation.

The only remaining cross-system P1 before full document-flow launch is the VBA VOZ writeback correction for `BrojZbirne` column T. PWA launch smoke itself is complete.

### 1.3 Major Components
| Component | Purpose | Owner | Notes |
|---|---|---|---|
| VBA/Excel Desktop | canonical desktop master, dokumenti, knjigovodstvena logika, SEF, banka | Operator / backoffice | single-operator model |
| Google Apps Script | REST-like action API, sync, auth, meteo batch, dispatch, fiskalni parse | Engineering | doGet/doPost routing |
| Google Sheets | operativni data transport i shared state između PWA i desktop sloja | Engineering / operator | per-client dedicated account |
| AgriX PWA | terenski i mobilni rad po rolama | Otkupac / Kooperant / Vozac / Management | offline-first, IndexedDB |
| External services | Open-Meteo, QR decode, Google Drive, SEF API | Engineering | deo aktivnog sistema |

### 1.4 Deployment Model
| Area | Model |
|---|---|
| Tenanting | per-client / per-firm deployment |
| Client isolation | dedicated Google account / sheets set per client |
| Operator model | one Excel file per firm, one active operator per file |
| Offline behavior | PWA offline-first sa IndexedDB + delayed sync |
| Environment separation | GAS deploy/version discipline; production tokenization expected |

### 1.5 User Roles
| Role | Primary responsibilities | Writes allowed | Sensitive actions |
|---|---|---|---|
| Otkupac | terenski unos otkupa, dodela vozača, štampa/izvoz otkupnog lista | OTK records, VozacID assignment on OTK records | vozac assignment, PDF upload |
| Kooperant | pregled parcela, tretmani, troškovi, knjiga polja, fiskalni unos | tretmani, troškovi, fiskalni private records | agronomical records, private expense/fiskalni data |
| Vozac | zbirna i transport overview | zbirna sync / transport data | aggregated transport document creation |
| Management | KPI pregled, dispatch planning, partneri, agro overview | dispatch plan, demand, kamion status | logistics planning, demand manipulation |
| Admin / Operator | Excel master, stammdaten, dokumenta, finansije, SEF, banka, exports | sve canonical desktop tabele | storno, faktura, SEF, BankaImport, export/sync |

### 1.6 Supported Platforms
| Platform | Supported | Notes |
|---|---|---|
| Excel Desktop | Yes | canonical desktop backoffice |
| Web / PWA | Yes | primary field UI |
| Android | Yes | camera/QR and field usage expected |
| iOS | Yes | photo fallback needed for fiscal/QR edge cases |
| Browser offline mode | Yes | app shell + IndexedDB; self-hosted Leaflet assets are cached, but live tile layers remain network-dependent unless previously cached by the browser |

---

## 2. Architecture Invariants

### 2.1 Always Rules
| Rule | Detail | Affected layers |
|---|---|---|
| ExcludeStornirano | Every query reading business data must call `ExcludeStornirano(data, tableName)` after `GetTableData()`; the helper is intentionally safe as a no-op on tables that do not have a `Stornirano` column | VBA |
| ID Format | IDs are `PREFIX-NNNNN`, zero-padded, with active prefixes including OTK, OTP, ZBR, PRJ, FAK, NOV, MAG, PAR, ART, KOOP, ST, VOZ, KUP, BIM, SFS, SFE, DPL, WRD, PRIV | VBA, GAS, PWA, Sheets |
| Array Basis | 2D VBA arrays are 1-based | VBA |
| Column Lookup | No hardcoded column index; always use `GetColumnIndex()` | VBA |
| Data access surface | All VBA sheet/table reads and writes must go through `modDataAccess` helpers such as `GetTableData()`, `AppendRow()`, `UpdateCell()`, `FindRows()`, `LookupValue()` and `CheckDuplicate()` | VBA |
| TX Wrapper | Every write operation has `_TX` wrapper with snapshot/rollback behavior | VBA |
| Desktop lifecycle entry rule | `Workbook_Open` is only the safe entry wrapper and delegates to `modMain.StartApp`; business work, imports and long-running side effects must not be added directly to `Workbook_Open`. | VBA/app lifecycle |
| Startup backup/log gate | Desktop boot must run startup backup/log housekeeping through `StartApp`/journal-log helpers and must never leave Excel invisible on startup failure. | VBA/app lifecycle |
| Controlled shutdown rule | Normal exits route through `ShutdownApp`, restore `Application.Visible = True`, unload the main shell and write `LogAppShutdown`; form close and workbook close must share the same exit contract. | VBA/app lifecycle |
| Main shell save/navigation rule | `frmOtkupAPP` is an operator shell, not a business module. Startup must call the full shell setup chain, `btnSnapshot_Click` must execute `SaveApp`, and popup/navigation helpers must use guarded EH without embedding domain writes. | VBA/app lifecycle, forms |
| Form activation safety rule | `UserForm_Activate` handlers in operator-facing forms use guarded EH, safe title-bar/chrome ordering and user-facing status/label feedback; they do not re-raise activation errors that would crash the UI lifecycle. | VBA/forms |
| SEF operator-shell safety rule | `frmSEF` is an operator-facing control surface only. Form activation must use guarded EH, two-column faktura combo setup must be explicit in code, `btnPosalji_Click` must always restore button enabled/caption state through a cleanup path, and cancel/storno buttons require explicit operator confirmation before destructive remote calls. | VBA/forms, SEF |
| Atomic document wrapper | Dual-class `Otpremnica`, `Zbirna` and `Prijemnica` saves must use one `Save*Multi_TX` wrapper so Klasa I and Klasa II commit or rollback together. Klasa II carries zero ambalaža side-effects. | VBA/document modules |
| Atomic money/packaging wrapper | `OM Ulaz` and `Izlaz Kupci` document-form paths must be saved through one wrapper transaction that covers packaging, money and status side-effects; wrappers call base `SaveNovac`, not nested `SaveNovac_TX`. | VBA/document/finance modules |
| Novac validation rule | `SaveNovac` is the canonical money-row writer and must validate `Tip`, entity/partner context and amount direction before append. A valid money row has exactly one positive direction (`Uplata` or `Isplata`), no negative amounts and no zero/zero amount combination. | VBA/finance |
| Error pattern | Business modules raise/propagate; forms catch and show UI feedback | VBA |
| Business UI boundary | Business/data modules must not show `MsgBox`; they log, raise or return failure. Only forms decide user-facing feedback. | VBA |
| Parse helper rule | User-entered numeric/date text is parsed only through shared `modParse` helpers (`TryParseDouble`, `TryParseLong`, `TryParseDateValue`, `NormalizeNumericText`). Forms must not parse business user input through `Val`, raw `CDbl`, raw `CLng` or duplicated private parser copies. | VBA/forms |
| Schema/update guard rule | Fail-fast data access guards live in `modSchemaGuard`: `RequireColumnIndex`, `RequireColumns` and `RequireUpdateCell`. Business modules must use these before indexed reads or critical updates. | VBA |
| Finance column guard rule | Financial helpers over `tblNovac`, `tblFakture`, `tblOtkup`, `tblPartnerMap` and faktura-stavka/prijemnica lookups must use `RequireColumnIndex` for required columns. Optional columns may use `GetColumnIndex` only with an explicit zero-index branch. | VBA/finance |
| Canonical update guard rule | `RequireUpdateCell` is the single canonical v6.6 fail-fast update helper. Form-local helpers such as `MustUpdateCell` and module-local helpers such as `RaiseUpdateError` are legacy compatibility debt and must be removed or kept only as thin wrappers. | VBA |
| Hidden-ID combo rule | ComboBox controls that select canonical entities must display human text in column 0 and store stable ID in hidden column 1 through `modComboBinding`; saves use `GetComboID`, not lookup-by-display-text. | VBA/forms |
| Soft delete | `Stornirano = "Da"` is soft-delete; no physical deletion of business records | VBA, Sheets |
| Filtered-array update-index rule | `ExcludeStornirano` may be used freely for read-only aggregation/report arrays, but filtered arrays must not supply row indexes to `UpdateCell` / `RequireUpdateCell`. Update flows must use full table row indexes from `GetTableData` / `FindRows` and skip `Stornirano="Da"` rows manually. | VBA/data access, finance |
| Storno isolation | Storno is performed per document/entity; there is no cross-document cascade through the chain, only explicit local side effects such as ambalaža storno, novac unlinking, prijemnica release, orphan marking or status recompute | VBA |
| Save return | Save functions return new ID or empty string on failure | VBA |
| Journal append | `AppendRow()` is the canonical row writer and must also emit `WriteJournalRow(...)` for transactional traceability | VBA |
| Display format | Desktop helper formatting supports both `"Name (ID)"` and `"ID - Name"` forms, while `ExtractIDFromDisplay()` must successfully recover the underlying ID from either style | VBA, PWA |
| Duplicate pre-check | Desktop document and money-entry forms must run `CheckDuplicate(...)` before `_TX` save when a user-entered document number exists; duplicate detection skips already stornirano rows | VBA/forms |
| Partner-map conflict rule | `savePartnerMap` treats an identical existing bank-partner mapping as idempotent success, but the same bank partner name mapped to a different PartnerID/EntitetTip/OMID is a fail-fast conflict, not a silent success. | VBA/finance, bank import |
| Faktura creation rule | `CreateFaktura()` trusts caller input only for `PrijemnicaID`. Quantity, price, class and receipt number are canonical values read from `tblPrijemnica`; invoice amount is the sum of canonical `Kolicina × Cena`, new invoices start `STATUS_NEPLACENO`, and buyer avans must be auto-applied immediately after creation. | VBA/fakturisanje |
| Faktura status/print guard rule | `UpdateFakturaStatus()` is a two-way recompute: sufficient active uplata sets `STATUS_PLACENO` and fills `DatumPlacanja` only if empty; insufficient uplata sets `STATUS_NEPLACENO` and clears `DatumPlacanja`. Stornirana faktura rows are not mutated. `PrintFaktura()` must block active-print of stornirana faktura. | VBA/fakturisanje, finance |
| Document read-helper active-row rule | Core document read helpers (`GetOtpremniceByZbirna`, `GetOtpremniceByStation`, `GetZbirnaByKupac`, `GetPrijemniceByKupac`, `GetOtkupByStation`, `GetOtkupByKooperant`) return active rows by default by applying `ExcludeStornirano` internally. Callers should not have to remember this filter for normal business reads. | VBA/document modules, otkup core |
| Document input validation rule | Base document writers for Otpremnica, Zbirna, Prijemnica and Otkup must fail fast on missing required IDs/numbers, invalid class, non-positive quantity, invalid price, negative money/ambalaža and missing ambalaža type when ambalaža quantity exists. | VBA/document modules, otkup core |
| tblPrijemnica returned-packaging column rule | The canonical workbook column for returned ambalaža on `tblPrijemnica` is `kolAmbVracena`. `KolAmbalazeVracena` is not canonical and must not be used in writers, tests or constants. | VBA/document modules, tests |
| Novac avans allocation rule | `ApplyAvansToFaktura[_TX]` and `ApplyAvansToOtkup[_TX]` are fail-fast allocation flows. Split avans updates use `RequireUpdateCell`, split-row creation checks that `SaveNovac` returned a `NOV-*` ID, stornirano money rows are excluded, and snapshot TX wrappers cover affected money/status tables. | VBA/finance |
| SEF dual-status rule | `SEFWorkflowState` is the internal/local process-control state, while `SEFStatus` stores the exact latest external SEF API status. They are related but intentionally do not have to be identical. `WF_SEF_SENT` means the local submit pipeline succeeded and `SEFDocumentId` exists, while the external status may still be `DRAFT`, `SENT`, `STORNO`, `ACCEPTED` or another SEF value returned by refresh. Submit result persistence must write `response.apiStatus` into `SEFStatus`, not internal workflow constants such as `WF_SEF_SENT`. | VBA/SEF |
| SEF submission identity rule | The outbound SEF `requestId` must equal the persisted `SEFSubmissionID`; payload hash and version number are stored alongside the submission and faktura state | VBA/SEF |
| SEF outbound split rule | SEF send orchestration must be split into PREP TX → SENDING TX → HTTP outside TX → RESULT TX so remote calls never happen inside an open desktop transaction | VBA/SEF |
| SEF idempotent refresh rule | `RefreshSEFStatus_TX` must be safe to run repeatedly. If the local workflow is already in the target state, the refresh updates status/error/document/sync fields only and must not attempt a same-state transition such as `SEF_ACCEPTED → SEF_ACCEPTED`. Pending and unknown non-final statuses such as `SENT`, `NEW`, `DRAFT` and parser fallback statuses route through `ApplySEFStateOrRefreshOnly`, not ad-hoc direct workflow writes. | VBA/SEF |
| SEF payload tax/total consistency rule | SEF DTO totals and UBL tax fields must use the same tax source through `GetDefaultTaxPercent`; hardcoded VAT math in mapper logic is not canonical. Header totals from `tblFakture` must match the sum of DTO line totals for net, VAT and gross within the active rounding tolerance before UBL submit; mismatch is a local validation failure, not a soft warning. | VBA/SEF |
| SEF HTTP client rule | `modSEFClient` centralizes base URL/API key/env lookup, WinHTTP timeout setup and headers. Debug output is controlled by `SEF_DEBUG_LOG`, HTTP 429 is mapped to `RATE_LIMITED`, and `GetInvoiceStatus` must parse with `ParseStatusResponse`. Lightweight parser hardening supports numeric-or-string SEF document IDs and tolerant boolean parsing for simple response fields such as `accepted`. | VBA/SEF |
| SEF persistence guard rule | `modSEFPersistance` read/write helpers must guard required SEF columns and use `RequireUpdateCell` for all critical faktura/submission/event state writes. | VBA/SEF |
| SEF stuck-sending recovery rule | `SEF_SENDING` rows must be recoverable: if `SEFDocumentId` exists, refresh external status; if it does not, transition to `SEF_TECH_FAILED` so normal retry paths can continue. Error handlers in recovery/send/cancel/storno flows must capture original `Err.Number`, `Err.Description` and `Err.Source` before rollback/logging so validator and API failure messages are preserved. | VBA/SEF |
| SEF invoice/delivery date rule | `clsSEFInvoiceSnapshot` owns both `InvoiceDate` and `DeliveryDate`. `InvoiceDate` is read from `tblFakture.Datum`; `DeliveryDate` is derived from prijemnice linked through `tblFakturaStavke.FakturaID -> PrijemnicaID -> tblPrijemnica.Datum`, using the latest linked prijemnica date when multiple lines exist. `SerializeUBLInvoice` must use `dto.InvoiceDate` and `dto.DeliveryDate` and must not recalculate business dates. | VBA/SEF |
| SEF UBL date validation rule | `ValidateSEFDtoForUBL` must fail locally before HTTP submit if `DeliveryDate > InvoiceDate`, because SEF rejects such UBL with delivery-date/issue-date validation errors. | VBA/SEF |
| SEF live smoke evidence rule | v6.7 baseline requires at least one valid live submit with HTTP 200, persisted `SEFDocumentId`, `tblSEFSubmission` row, `tblSEFEventLog` rows and idempotent repeated refresh. Negative evidence must include SEF-side business rejection persistence and local UBL validation failure before HTTP submit. A successful local workflow state such as `WF_SEF_SENT` or `WF_SEF_ACCEPTED` without `SEFDocumentId` is a failure, never a PASS/SKIP. | VBA/SEF/tests |
| SEF destructive-test safety rule | cancel/storno live tests are destructive and must be gated by explicit config and user confirmation. Test PASS for cancel/storno must assert the Boolean result returned by `CancelInvoiceOnSEF_TX` / `StornoInvoiceOnSEF_TX` and distinguish API/event smoke from final external business outcome. Already-`STORNO` storno attempts are expected SKIP/validator blocks, not production failures. | VBA/SEF/tests |
| SEF success-document rule | A successful SEF submit response must return a non-empty `SEFDocumentId`. If `response.Success=True` but `SEFDocumentId` is missing, the response is normalized to a technical failure (`FAILED` / `MISSING_SEF_DOCUMENT_ID`) and persisted as a failed submission instead of moving the faktura to a successful outbound state. | VBA/SEF |
| SEF state-transition test rule | `ValidateAllowedTransition` is covered by an offline transition-matrix test suite. `WF_SEF_STORNO` is a terminal local workflow state when entered; this does not imply that every external `SEFStatus=STORNO` from refresh must automatically rewrite local workflow to `WF_SEF_STORNO`. | VBA/SEF/tests |
| Two-class packaging rule | In dual-class otpremnica/zbirna/prijemnica saves, Klasa II is written as a separate row and carries `kolAmb = 0` and `kolAmbVracena = 0`; ambalaža remains tied to Klasa I aggregate | VBA/forms, document modules |
| Packaging journal rule | `TrackAmbalaza(datum, tipAmb, kolicina, smer, entitetID, entitetTip, VozacID, dokumentID, dokumentTip)` is the canonical writer for ambalaža movements and must preserve optional driver and document lineage on every business-side packaging effect | VBA/document modules |
| Packaging read rule | Every balance/query helper over `tblAmbalaza` must first apply `ExcludeStornirano(...)`; entity saldo is net `Ulaz - Izlaz`, while driver saldo is exposed as `Izlaz`, `Ulaz`, `Saldo = Izlaz - Ulaz` grouped by `TipAmbalaze` | VBA/report helpers |
| Desktop validation gate | Desktop zbirna save is blocked unless class-level kg validation and ambalaža validation are green through `UpdateValidacija()` | VBA/forms |
| Orphan-warning rule | After storno and on form startup, UI must surface orphaned otpremnice/prijemnice waiting for a new zbirna via `GetVerwaisteDokumente()` | VBA/forms |
| Orphan semantics | A document is orphaned only if it points to a `BrojZbirne` that is stornirana and no active zbirna with the same number exists as a replacement | VBA/document modules |
| Prijemnica relink rule | Saving a replacement prijemnica must attempt `RelinkFakturaStavke(newPrijemnicaID, brojPrijemnice)` so orphaned invoice lines reconnect and the new prijemnica inherits fakturisano status when applicable | VBA/document modules |
| Shortage allocation rule | Transport shortage analytics may be calculated per zbirna and proportionally allocated per otpremnica, including value impact via row price | VBA/report helpers |
| Locale | Serbian decimal conventions; no manual decimal replacement before `CDbl()` | VBA, PWA |
| Backup | Backup on app start + CSV journal on transactional writes | VBA |
| Active status semantics | `tblStanice` and `tblVozaci` use `Aktivan = "Aktivan"`; `tblKooperanti` active means not `"Ne"` | VBA, sync, PWA |
| localStorage ban | localStorage is forbidden for shared state; only auth/session/device helper values are allowed | PWA |
| Dispečer = planning only | Dispatch planning cannot write `VozacID` directly into OTK sheets; only Otkupac field flow may set it | PWA, GAS, Sheets |
| PWA modularity | No monolith; modular JS file structure with one responsibility per file | PWA |
| Artikli master rule | Master `Artikli` is operator-controlled; kooperant-private fiskalni artikli never enter master artikli | VBA, GAS, PWA, Sheets |
| Parcel selector rule | Kooperant-scoped parcela lookup returns canonical parcela tuples plus a preformatted display label built from katastarski broj, opština, kultura and površina | VBA/agro forms |
| Warehouse valuation rule | `SaveMagacin` must derive `CenaPoJedinici` from `tblArtikli` at write time and persist `Vrednost = Kolicina * CenaPoJedinici` on the journal row | VBA/agro stock |
| Warehouse read rule | Every stock/debt/report helper over `tblMagacin` must first apply `ExcludeStornirano(...)`; live stock is exposed as `Ulaz`, `Izlaz`, `Stanje = Ulaz - Izlaz` grouped by `ArtikalID` | VBA/report helpers |
| Dosage recommendation rule | Desktop agro recommendation is calculated from article dosage and parcel area as `DozaPoHa * PovrsinaHa`; missing or non-numeric dosage yields `0` recommendation | VBA/agro forms |
| Otkup write rule | `SaveOtkup_TX()` snapshots `tblOtkup` and `tblAmbalaza`; `SaveOtkup()` allocates `OTK-*`, resolves `KulturaID`, persists optional `ParcelaID` / `BrojZbirne`, and records packaging as `Izlaz` from the kooperant when ambalaža exists | VBA/otkup core |
| Traceability auto-link rule | Automatic `Otkup → Otpremnica` linkage first uses the strict key `StanicaID|Datum|VozacID|Klasa|BrojZbirne`. It may use the legacy fallback key `StanicaID|Datum|VozacID|Klasa` only when `BrojZbirne` is missing on the relevant side and the candidate is unique. Cross-`BrojZbirne` links are invalid and must remain unresolved for manual repair. | VBA/sledljivost |
| Otkup atomic save rule | Desktop `frmOtkup` must treat Klasa I, optional Klasa II, ambalaža, cash payout and avans allocation as one business operation through `SaveOtkupMulti_TX`; forms must not split this into separate `_TX` calls | VBA/otkup core |
| Fakturisanje duplicate-prevention rule | `CreateFaktura_TX` must fail if a selected prijemnica is stornirana, already `Fakturisano="Da"`, already has `FakturaID`, duplicated in the same selection, or if any faktura/stavka/prijemnica update fails | VBA/fakturisanje |
| PWA-first traceability rule | Preferred field source is PWA: PWA-created otkup records should carry station, driver, date, class, parcela when available and document/zbirna context when available; VBA remains the fallback and repair layer | PWA, Sheets, VBA/sledljivost |
| Canonical trace bridge | `tblOtkup.OtpremnicaID` is the desktop bridge from raw/PWA otkup records into the canonical chain `Zbirna → Otpremnica → Otkup → Kooperant/Parcela`; `BrojZbirne` is part of the preferred bridge key and prevents same-day same-driver cross-zbirna collisions; missing or ambiguous links remain repair work, not canonical trace output. | VBA/sledljivost |

| In-memory array rule | Filtering, sorting, grouping and summing over tabular business data must happen in memory through `modArrayUtils` helpers instead of sheet copy/paste or worksheet sort operations | VBA |
| Non-blocking observability | Error logs, remote client logs, journal writes, backups and retention purge helpers must never block the core business operation; failures in observability helpers are best-effort only | VBA, GAS, PWA/runtime |
| escapeHtml always | User-supplied strings must be sanitized before `innerHTML` rendering | PWA |

### 2.1.1 Google Apps Script Backend Invariants
| Rule | Detail | Affected layers |
|---|---|---|
| MASTER folder rule | The primary GAS workbook family lives under `MASTER_FOLDER_ID`; role- and domain-specific spreadsheets are created lazily inside that Drive folder through `getOrCreateSheet(...)` | GAS, Drive, Sheets |
| Geo workbook split | Parcel geo and meteo persistence are intentionally split out into a dedicated workbook `GEO_SPREADSHEET_ID` with at least `Parcele`, `MeteoLatest` and `MeteoHistory` tabs, instead of living under the main workbook family | GAS, Sheets, PWA |
| GAS idempotent sync rule | `sync`, `syncZbirna`, `syncTretman`, `syncOprema` and the current agromere sync flow key on trimmed `ClientRecordID`; an existing row must return success with `status = existing`, avoid duplicate inserts, and must not reset terminal/master/error statuses such as `Synced>Master`, `Duplicate` or `SyncError:*` | GAS, PWA, Sheets |
| Server ID allocation rule | New server-side records use timestamped random IDs via `generateServerRecordID(...)` or `generateEntityServerID(prefix, entityID)`; active backend prefixes explicitly visible in this snapshot include `OTK`, `ZBR`, `TRT`, `OPR`, `WRD`, `DPL` and `IZD` | GAS, Sheets |
| Sync status rule | GAS write-side sync flows normalize new successful rows to `SyncStatus = Synced` and stamp server timestamps; idempotent retry paths preserve terminal/master/error states and do not rewrite first-receive fields on already terminal rows | GAS, PWA, Sheets |
| Token persistence rule | Session tokens are stored first in `CacheService.getScriptCache()` under `TOKEN_<token>` with a 24h TTL and are mirrored into `PropertiesService.getScriptProperties()` as a persistent fallback; both cache and property payloads must pass created-time validation, malformed/expired payloads are rejected/deleted, and valid fallback tokens are restored into cache | GAS, PWA |
| Token purge rule | `purgeExpiredTokens()` removes expired or malformed `TOKEN_*` script properties and also calls `purgeOldErrorLogs()`; `setupTokenPurgeTrigger()` provisions a daily 03:00 `Europe/Belgrade` trigger for this maintenance path | GAS |
| Role gate rule | Management-only actions must explicitly verify `tokenData.role === 'Management'`; all role-scoped write actions must verify both role and entity ownership before executing. Otkupac write actions must match `tokenData.entityID === otkupacID`; Kooperant write actions must match `tokenData.entityID === kooperantID`; Vozac write actions must match `tokenData.entityID === vozacID`, except where Management is explicitly allowed as an override. | GAS, PWA |
| GAS endpoint authorization helper rule | GAS `doPost` uses small centralized authorization helpers (`isManagement`, `requireRole`, `requireEntity`, `forbiddenResponse`) instead of repeating ad-hoc role checks in every endpoint block. These helpers are local `Code.gs` helpers, not a separate auth layer. | GAS |
| GAS write endpoint auth rule | Business write endpoints execute after token validation unless explicitly documented as an exception. `login` and `logClientError` are intentional pre-auth exceptions, and `saveParcelPolygon` remains an intentional public/pre-auth exception in the v6.10 codebase by product decision. The public geo/meteo read bridge remains separately acknowledged under the public geo/meteo exposure rule. | GAS, Security |
| GAS write lock rule | Every active GAS action that writes to Google Sheets or Google Drive must execute inside `withLock(...)`. This includes sync writes, dispatch/demand writes, PDF upload writes, fiskalni save/mapping writes, kamion status writes and master-artikal creation. Disabled endpoints must return `FEATURE_DISABLED`; parsing-only endpoints do not require lock unless they persist data. | GAS, Sheets, Drive |
| GAS sync ownership rule | `sync`, `syncAgromere`, `syncZbirna`, `syncTretman`, `syncOprema` and `syncTrosak` are active role/entity-scoped sync actions. Non-management users may sync only their own entity scope. Management may act as explicit override where supported by the endpoint authorization matrix. `syncTrosak` must use the same lock-protected batch response contract as the other active sync endpoints and must never return an empty HTTP 200 body. | GAS, PWA, Sheets |
| Kamion status ownership rule | `updateKamionStatus` accepts `Vozac` and `Management`. When the caller is `Vozac`, GAS must ignore or overwrite any client-supplied `vozacID` and use `tokenData.entityID` as the authoritative driver ID. Management may update any driver status. | GAS, Dispatch |
| Fiskalni mapping authorization rule | `saveFiskalniMapiranje` is Management-only because it changes shared fiscal-name-to-artikal mapping. Kooperant fiscal receipt save/parse flows remain kooperant-scoped, but learned/global mapping writes are not available to Kooperant role. | GAS, Fiskalni, Security |
| Fiscal parsing auth rule | `parseFiskalniImage` and `parseFiskalni` are not public actions. They require a valid token and are allowed only for `Kooperant` or `Management`; if a kooperant scope is supplied, Kooperant callers must match `tokenData.entityID`. | GAS, PWA, Fiskalni |
| Parcel polygon public-write exception | `saveParcelPolygon` remains an intentional public/pre-auth endpoint in the active v6.10 codebase. This is an acknowledged security exception and must be protected operationally by deployment URL control until a future geo-editor/auth model is introduced. | GAS, GIS, Security |
| Public geo/meteo exposure rule | The active backend now supports a POST-first public-read bridge for `getParcelGeo`, `getParcelMeteo`, `getParcelMeteoLatest` and `getAllMeteoLatest` before token validation so the frontend can stay on one request model; public geo/meteo access remains an acknowledged auth-gap until those reads are fully gated | GAS, PWA, Security |
| GAS remote error logging rule | `logError(source, action, message, details, entityID)` is the canonical GAS-side logger; it lazily creates/uses workbook `ErrorLog` in `MASTER_FOLDER_ID` with columns `Timestamp \| Source \| Action \| Message \| Details \| EntityID \| Severity`, truncates long message/detail payloads, classifies timeout messages as warning, and must remain best-effort / non-blocking | GAS, PWA, Observability |
| PWA client error bridge rule | `doPost` accepts `logClientError` before the normal token gate so field devices can report client/runtime failures even after session expiry; if a valid token is present, GAS resolves `EntityID` from token data, otherwise it falls back to the payload `entityID`. Client log payloads are truncated and token/PIN/password/base64-like sensitive content is redacted before persistence. | GAS, PWA, Observability |
| GAS schema drift rule | `ensureSheetColumns(sheet, requiredColumns)` creates canonical headers only for empty sheets. Existing sync sheet header mismatch, missing required columns or extra named columns are `SCHEMA_DRIFT` failures and must not be silently repaired by appending columns. | GAS, Sheets, VBA sync |
| GAS batch response rule | Sync batch responses use `success=false` whenever at least one record fails. `code` is `OK`, `PARTIAL_FAILURE` or `BATCH_FAILED`, and per-record results remain available for retry/UI handling. | GAS, PWA |
| GAS disabled endpoint rule | Inactive routes such as `saveOtkupniListPdf` must return `FEATURE_DISABLED` until real implementations exist; active route healthcheck must not list handlers for intentionally disabled endpoints. `syncTrosak` is no longer inactive in v6.12. | GAS, PWA |
| GAS route smoke rule | `runGasRouteHealthCheck()` verifies active handler presence, and `runGasSmokeSuite()` verifies route health, batch response semantics, authz/login validation, schema drift guard and normalized lookup without inserting production business rows. | GAS, QA |
| Google VBA config ownership rule | `tblSEFConfig` is the central `ConfigKey/ConfigValue` app configuration table. `GetConfigValue` and `SetConfigValue` live in `modConfig`; Google modules consume config but do not own the central writer. | VBA, GoogleAuth, Config |
| Google Sheets fail-fast write rule | `WriteSheetData` must fail if `ClearSheet` fails; Drive move/create/find/read/write helpers must validate inputs, log bounded HTTP failures and avoid silent orphan/wrong-target behavior. | VBA, GoogleSheets |
| Meteo freshness rule | `getParcelMeteo(parcelaId)` first reuses `MeteoLatest` if `LastFetch` is younger than 12 hours and only then falls back to live Open-Meteo retrieval | GAS, PWA |
| Meteo trigger rule | Scheduled meteo is expected to run 4 times daily in `Europe/Belgrade` timezone, refreshing `MeteoLatest` as overwrite-state and `MeteoHistory` as append-only history | GAS, Sheets |
| Fiscal duplicate rule | Fiscal parse/save flows must reject already scanned receipts by `VerificationUrl` within the kooperant-specific `FISKALNI-<KooperantID>` workbook before saving parsed line items | GAS, PWA, Sheets |

### 2.2 Never Rules
| Anti-pattern | Why forbidden | Approved alternative |
|---|---|---|
| Hardcoded column indices | breaks schemas and maintenance | `GetColumnIndex()` |
| Raw `innerHTML` with user data | XSS risk | `escapeHtml()` + safe DOM helpers |
| Raw `fetch()` without wrapper | inconsistent auth/error handling | `apiFetch()` / `apiPost()` |
| `localStorage` for shared state | creates phantom stale state across sessions/devices | IndexedDB + server sync + AppState |
| Direct sheet/range access in business logic | bypasses access and rollback conventions | `modDataAccess` functions |
| Business logic inside `Workbook_Open` | boot failures can leave Excel hidden or block the operator before UI is available | `Workbook_Open` → `modMain.StartApp`; operator actions trigger imports explicitly |
| Auto-running bank import at workbook open | external/file import errors are not safe startup work and can block the app shell | Bank import is operator-triggered from Banka mapping/navigation flow |
| Re-raising `UserForm_Activate` errors after logging | activation failures become disruptive UI/runtime failures | log and show controlled label/MsgBox/status feedback from the form |
| `MsgBox` inside business modules | couples data layer to UI, weakens tests and makes transaction failure handling inconsistent | `Err.Raise` / return `""` or `False`; forms show messages |
| `Val` or raw `CDbl/CLng/CDate` on user input | breaks decimal/date semantics under Serbian/European input and can silently truncate values | `modParse.TryParseDouble`, `TryParseLong`, `TryParseDateValue` |
| Lookup by visible ComboBox text for IDs | duplicate names can write records to the wrong entity | hidden-ID ComboBox binding via `modComboBinding.GetComboID` |
| Unguarded critical `UpdateCell()` | partial relink/status updates can be committed without detection | `RequireUpdateCell(...)` inside an active TX scope |
| Parallel local update-guard helpers as canonical API | duplicates behavior and creates drift between forms/modules | `modSchemaGuard.RequireUpdateCell` as the single canonical helper |
| Parsing SEF status responses with submit parser | submit and status responses have different semantics and can corrupt workflow interpretation | `GetInvoiceStatus` must use `ParseStatusResponse` |
| Converting `SEFDocumentId` through `CLng` for cancel/storno | SEF IDs may exceed VBA `Long` range and overflow | validate numeric string with `GetJsonNumericIdLiteral` and write as JSON numeric literal |
| Recalculating SEF business dates inside serializer | serializer can diverge from DTO/validator and produce inconsistent UBL | `BuildSEFInvoiceDto` owns `InvoiceDate`/`DeliveryDate`; `SerializeUBLInvoice` only renders DTO fields |
| `CDbl(Replace(...))` | locale corruption | direct `CDbl()` |
| PWA logic in `Code.gs` | layer leakage and drift | keep UI helpers in PWA JS |
| Dispatcher writing `VozacID` | breaks planning/write-authority boundary | only Otkupac QR assignment flow may write |
| Private fiskalni artikli into master artikli | corrupts central agrohemija catalog | save into FISKALNI-KOOP or `PRIV-*` private scope |

### 2.3 Data Safety Rules
- **Backup policy:** `BackupFileOnStart` creates a dated workbook copy on app start, while `PurgeOldBackups()` removes backups older than 30 days.
- **Journal/audit policy:** every successful `AppendRow()` writes a CSV journal record immediately; `CheckJournalForRecovery()` compares today's journal counts against live table row counts to detect possible crash-loss situations.
- **Desktop log policy:** `modLogError` writes one daily text log under `Log/`, purges files older than 30 days, and must never block business execution.
- **GAS/PWA remote log policy:** `logError(...)` writes best-effort rows into the `ErrorLog` workbook in `MASTER_FOLDER_ID`; `purgeOldErrorLogs()` removes rows older than 30 days and is invoked from the token purge maintenance path.
- **GAS write concurrency policy:** all GAS action handlers that write to Google Sheets or Google Drive must wrap the write operation in `withLock(...)` to reduce duplicate/race-condition risk during concurrent sync or management operations.
- **Soft delete policy:** business data is soft-deleted via `Stornirano = "Da"`.
- **Recovery principle:** transaction rollback on VBA errors, sync retry on PWA/GAS errors, recovery paths for stuck SEF states, startup-time backup/journal/log maintenance in desktop runtime, and GAS token/error-log maintenance through scheduled purge helpers.
- **Fail-fast guard policy:** column lookup failures and failed critical row updates are explicit errors through `modSchemaGuard`, not silent fallbacks. This is mandatory for document relink, orphan repair and status refresh code paths.

### 2.4 Naming and ID Rules
| Entity | Prefix / format | Example | Notes |
|---|---|---|---|
| Otkup | `OTK-NNNNN` | `OTK-00042` | canonical otkup row |
| Otpremnica | `OTP-NNNNN` | `OTP-00011` | transport document |
| Zbirna | `ZBR-NNNNN` | `ZBR-00008` | aggregate transport / buyer doc |
| Prijemnica | `PRJ-NNNNN` | `PRJ-00005` | cold-storage receiving doc |
| Faktura | `FAK-NNNNN` | `FAK-00012` | accounting invoice |
| Novac | `NOV-NNNNN` | `NOV-00321` | money movement |
| Parcela | `PAR-NNNNN` | `PAR-00123` | parcel record |
| Artikal | `ART-NNNNN` or `PRIV-*` | `ART-00021`, `PRIV-171344` | private fiscals stay outside master artikli |
| Stanica | `ST-NNNNN` | `ST-00003` | otkup station |
| Vozac | `VOZ-NNNNN` | `VOZ-00007` | driver |
| Kupac | `KUP-NNNNN` | `KUP-00004` | buyer |
| Banka import | `BIM-NNNNN` | `BIM-00009` | import staging |
| SEF entities | `SFS-*` / `SFE-*` | `SFS-...` | SEF submission/event families |
| Dispatch plan | `DPL-NNNNN` | `DPL-00002` | dispatch planning |
| War room / dispatch demand | `WRD-NNNNN` | `WRD-00005` | legacy naming still present in some tabs |

### 2.5 Error Handling Rules
| Layer | Pattern | Propagation model | User feedback |
|---|---|---|---|
| VBA modules | `Err.Raise` / propagate | caller/TX wrapper handles rollback | no direct business MsgBox preferred |
| VBA forms | `On Error GoTo EH` | catches business errors | MsgBox / label feedback |
| GAS | action routing + structured response + `logError(...)` best-effort observability | client interprets failure; logger must not block primary response | returned error payload / remote `ErrorLog` row |
| PWA | `safeAsync()` + wrapper API calls | promise errors centralized | toast / inline validation / sync badge |

### 2.6 Security and Sanitization Rules
| Area | Rule | Enforcement |
|---|---|---|
| Authentication | PIN-based login with token per session | GAS auth + PWA session handling |
| Authorization | actions are role/entity scoped; GAS write actions must enforce role and ownership checks before any write or Drive operation executes | token + entity-based action routing + endpoint authorization matrix |
| HTML sanitization | all user-facing dynamic strings must use `escapeHtml()` | PWA render layer |
| Input validation | forms validate required fields and domain rules before save | PWA field validation + VBA validation |
| Sensitive data exposure | private fiskalni and auth/token data must stay scoped | role/entity isolation + storage rules |

### 2.7 Offline and Sync Rules
- **What works offline:** app shell, cached stammdaten subset, IndexedDB record creation, kooperant tretmani/troškovi, otkup queue, local render caches
- **What requires online:** server sync, meteo refresh, Drive uploads, QR image decode, dispatch shared state refresh
- **Queue/retry rule:** records save locally first, then sync with status transitions (`pending`, `synced`, `error`)
- **Merge/conflict rule:** pending/error local records win over server; synced clean server data wins otherwise
- **Idempotency rule:** sync and SEF flows rely on stable IDs/request IDs and duplicate-aware write logic

### 2.8 Source-of-Truth Rules
Za svaki domen mora biti jasno:

- **desktop master** je canonical za finansije, dokumenta, SEF, banka i master data authoring
- **Google Sheets shared tabs** su operativni shared state za PWA workflows
- **IndexedDB** je local-first cache/queue, ne canonical shared source
- **dispatch** ima shared server/sheet source; local fallback nije authoritative
- **private kooperant fiskalni/troškovi** ostaju u dedicated kooperant sheets, ne u master artikli
- **derived reports** nisu canonical izvor, već read modeli / agregati

### 2.9 Lessons Learned / Engineering Principles
Ove stavke su dovoljno puta potvrđene kroz implementaciju i debugging da ulaze u aktivna inženjerska pravila sistema:

- **LocalStorage is not shared operational truth:** `localStorage` je prihvatljiv samo za device-local preference, lightweight cache i helper state (`deviceID`, kapacitet kamiona, potpis, UI flags). Ne sme nositi shared operativni state koji više korisnika ili više uređaja treba da deli.
- **Sheet-backed caches beat redundant short-lived caches:** kada je `MeteoLatest` već aktivni materialized read model, dodatni kratki server-side cache sloj ima smisla samo ako stvarno smanjuje API load bez uvodjenja staleness confusion.
- **Modular refactor is valuable but regression-prone:** razbijanje monolita u `features/`, `services/`, `ui/` i `utils/` je ispravan target, ali svaka promena load-order-a, global compatibility layer-a ili mount ownership-a mora se posmatrati kao regression-risk.
- **Generic utilities reduce drift:** shared helper-i kao `mergeOfflineRecords`, DOM helper-i, sync lifecycle vokabular i signature-pad wrapper treba da ostanu generički; dupliciranje slične logike po feature modulima povećava divergence risk.
- **Signature pads must be explicitly cleaned up:** modal/rehydrate scenariji lako ostavljaju event-listener i memory leak tragove; zato `destroySignaturePad(...)` / `destroyAllSignaturePads()` ostaju deo aktivnog UI contract-a.
- **Smart dosage must round predictably:** preporuka `DozaPoHa × PovrsinaHa` sme da se zaokružuje samo po jasno definisanom packaging rule-u; korisnik mora videti i raw quantity i rounded package interpretation.
- **GPS and meteo features must respect battery/latency budgets:** geolocation watch, map/meteo prewarm i polling ne smeju postati hidden battery drain; zato je cleanup on tab-exit i cache-first meteo rendering obavezna praksa.
- **Token/session expiry must degrade visibly:** auth/session istek ne sme završiti u tihim 401 greškama; korisnik mora dobiti jasan UI signal i re-auth path.

---

## 3. Source-of-Truth Matrix

| Domain | Canonical Store | Writable By | Read By | Synced To | Conflict Rule | Notes |
|---|---|---|---|---|---|---|
| Master data | Excel tables (`tblKooperanti`, `tblStanice`, `tblKupci`, `tblVozaci`, `tblKulture`, `tblArtikli`, `tblParcele`) | Operator | VBA, sync export, PWA | Stammdaten sheets | Excel master wins; export overwrites shared read model | PWA reads exported copy |
| Otkup records | OTK-* sheets operationally, then `tblOtkup` canonically after import | Otkupac (sheet), Operator (master import/storno) | Otkupac, Management, VBA | `tblOtkup`, `tblOtpremnica` | pending local wins before sync; master canonical after import | append-oriented |
| Dispatch planning | `DispecerPlan` + `KamionStatus` sheets | Management / dispatch | Management, dispatch UI | Mgmt read models | server/sheet state wins; local fallback never canonical | planning only |
| Financial records | `tblNovac`, `tblFakture`, `tblFakturaStavke`, `Kartice` export | Operator; some bank import mappings | VBA, PWA kooperant/management | Kartice/MgmtReports exports | desktop master wins | Kartice is derived/exported |
| Parcel GIS | `tblParcele` + geo fields / polygon data | Operator; `saveParcelPolygon` remains a public/pre-auth GIS write exception in v6.10 | Kooperant, Management | Stammdaten export, map UIs | master parcel record wins | public polygon write is acknowledged security exception |
| Meteo | `MeteoLatest` current read model + `MeteoHistory` history | Scheduled GAS fetch | Kooperant, Management | PWA prefetch/cache | latest overwrite for current, append for history | Open-Meteo upstream |
| Fiskalni receipts | `FISKALNI-{KooperantID}` sheets | Kooperant-own-scope or Management through authorized GAS endpoint | Kooperant, GAS parse/save | none to master artikli | kooperant sheet is canonical for private fiscals | auto-match may reference master artikli IDs |
| Reports / KPIs | `MgmtReports`, `SaldoOMDetail`, PWA caches | operator exports / computed views | Management, Kooperant, VBA | PWA read models | recompute/export beats cache | derived only |

### 3.1 Ownership by Domain
- **Operator / desktop backend** owns master data, formal documents, finance, SEF and bank mapping.
- **Otkupac** owns creation of field otkup records and explicit field-side `VozacID` assignment.
- **Kooperant** owns treatment, expense and private fiscal intake records.
- **Vozac** owns zbirna creation and transport-side completion.
- **Management** owns dispatch plan and demand management, but not raw OTK write authority.

### 3.2 Write Authority
- **Create/edit/storno desktop documents:** Operator only
- **Create otkup in field:** Otkupac
- **Assign driver on field otkup:** Otkupac only
- **Create zbirna in PWA:** Vozac
- **Create/update dispatch plans:** Management
- **Create tretmani/troškovi/fiskalni private rows:** Kooperant
- **Create meteo rows:** scheduled GAS jobs
- **Create SEF submissions/events:** VBA desktop modules
- **GAS `sync` write:** Otkupac only for own `OtkupacID`; Management may override only through explicit endpoint authorization.
- **GAS `syncAgromere`, `syncTretman`, `syncOprema`, `syncTrosak`:** Kooperant only for own `KooperantID`; Management may override only through explicit endpoint authorization. `syncTrosak` is active in v6.12 and must return a normal batch response.
- **GAS `syncZbirna`:** Vozac only for own `VozacID`; Management may override only through explicit endpoint authorization.
- **GAS `updateKamionStatus`:** Vozac may update only own status; Management may update any driver.
- **GAS `saveParcelPolygon`:** intentional public/pre-auth exception in v6.10; protect operationally until a dedicated auth model is introduced.
- **GAS `saveFiskalniMapiranje`:** Management-only.
- **GAS `createArtikal`:** Management-only.
- **GAS `parseFiskalniImage` / `parseFiskalni`:** authenticated Kooperant or Management only; not public.

### 3.3 Derived Models
Derived / cache layers include:

- `SaldoOMDetail`
- `MgmtReports`
- `Kartice`
- `MeteoLatest`
- PWA role caches (`pregledCache`, `karticaCache`, `meteoCache`, `_parceleLoaded`, `_kpLoaded`)
- local IndexedDB queues/stores
- management overview bundles from `getMgmtAll`
- warehouse/agro derived views such as `GetMagacinStanje()`, `ReportIzdavanjePoKooperantu()`, `ReportStanjePoDoabvljacu()` and kooperant debt helpers

### 3.4 Conflict Resolution
- **PWA sync:** pending/error local records dominate over stale server reads
- **Desktop exports:** authoritative export refresh replaces stale PWA-side read model
- **Dispatch:** shared server/sheet state dominates local UI state
- **SEF:** state-machine transitions and explicit recovery procedures prevent blind overwrites
- **Bank import:** duplicate checks use bank reference or fallback composite keys, while reconciliation/mapping advances staged rows through explicit `Obradjeno` states rather than mutating imported source facts

---

## 4. Data Architecture

### 4.1 Canonical Entities
| Entity / Table | Purpose | Primary key | Canonical owner | Soft delete | Notes |
|---|---|---|---|---|---|
| tblKooperanti | kooperant master data | KooperantID | Operator | No explicit Stornirano in active use; `Aktivan` controls visibility | exported to PWA |
| tblStanice | station master data | StanicaID | Operator | no `Stornirano`; `Aktivan` only | special rule |
| tblVozaci | driver master data | VozacID | Operator | no `Stornirano`; `Aktivan` only | includes `KapacitetKG` |
| tblKupci | buyer master data | KupacID | Operator | active status based | exported |
| tblKulture | fruit/sort catalog | KulturaID | Operator | n/a | lookup |
| tblOtkup | field procurement records | OtkupID | Operator after import | Yes | central document chain start |
| tblOtpremnica | transport document rows | OtpremnicaID | Operator | Yes | may be auto-created from PWA |
| tblZbirna | aggregated transport document | ZbirnaID | Operator / Vozac flow import | Yes | buyer/cold storage aggregation |
| tblPrijemnica | receiving document | PrijemnicaID | Operator | Yes | cold storage intake |
| tblFakture | invoice header | FakturaID | Operator | Yes | includes SEF state |
| tblFakturaStavke | invoice lines | StavkaID | Operator | Yes / orphanable | linked to prijemnice and storno/relink workflows |
| tblNovac | money movement | NovacID | Operator | Yes | OM, kooperant, kupac, banka |
| tblAmbalaza | packaging movement | AmbalazaID | Operator | Yes | side-flow tied to documents |
| tblParcele | parcel master + geo + meteo flags | ParcelaID | Operator | Active/Aktivna flag | geo enriched |
| tblArtikli | agrohemija master catalog | ArtikalID | Operator | Active flag | private receipts excluded |
| tblMagacin | warehouse movements | MagacinID | Operator | Yes | agro issuance/stock |
| tblConfig | generic config | Parameter | Operator | n/a | operational config |
| tblSEFConfig | SEF and selected PWA config | ConfigKey | Operator | n/a | seller + integration config |
| tblBankaImport | bank import staging | BankaImportID | Operator | Yes | staging only |
| tblPartnerMap | learned bank mapping | composite logical key | Operator/system | n/a | helper table |
| tblSEFSubmission | SEF submission journal | SEFSubmissionID | Operator/system | Yes | request/response history |
| tblSEFEventLog | SEF event timeline | SEFEventID | Operator/system | Yes | audit/event stream |

### 4.2 Table Schemas

#### 4.2.1 tblKooperanti
**Purpose:** master list of suppliers / kooperants  
**Primary key:** `KooperantID`  
**Foreign keys:** `StanicaID`  
**Written by:** Operator  
**Read by:** VBA, PWA, exports  
**Soft delete:** operationally via `Aktivan`, not canonical `Stornirano` flow  
**Important invariants:** active filter is `Aktivan ≠ "Ne"`

```text
KooperantID | Ime | Prezime | Mesto | Telefon | StanicaID | Aktivan | BPGBroj | TekuciRacun | PIN | Adresa | JMBG
```

#### 4.2.2 tblStanice
**Purpose:** otkup station master data  
**Primary key:** `StanicaID`  
**Foreign keys:** none  
**Written by:** Operator  
**Read by:** VBA, PWA, dispatch, sync  
**Soft delete:** no `Stornirano`; use `Aktivan = "Aktivan"`  
**Important invariants:** never call `ExcludeStornirano()` on this table

```text
StanicaID | Naziv | Mesto | Kontakt | Aktivan | Ime | Prezime | PIN
```

#### 4.2.3 tblVozaci
**Purpose:** driver master data  
**Primary key:** `VozacID`  
**Foreign keys:** none  
**Written by:** Operator  
**Read by:** VBA, PWA, dispatch  
**Soft delete:** no `Stornirano`; use `Aktivan = "Aktivan"`  
**Important invariants:** includes truck capacity for dispatch planning

```text
VozacID | Ime | Prezime | Telefon | Aktivan | PIN | KapacitetKG
```

#### 4.2.4 tblKupci
**Purpose:** buyer master data  
**Primary key:** `KupacID`  
**Foreign keys:** none  
**Written by:** Operator  
**Read by:** VBA, PWA management, SEF  
**Soft delete:** active status / stornirano in dependent docs, not direct workflow focus  
**Important invariants:** used by faktura and zbirna chains

```text
KupacID | Naziv | Mesto | PIB | MaticniBroj | Ulica | PostanskiBroj | Drzava | Email | Hladnjaca | Aktivan | TekuciRacun
```

#### 4.2.5 tblKulture
**Purpose:** fruit type and sort lookup  
**Primary key:** `KulturaID`  
**Foreign keys:** none  
**Written by:** Operator  
**Read by:** VBA, PWA  
**Soft delete:** not emphasized  
**Important invariants:** drives vrsta/sorta UI filters

```text
KulturaID | VrstaVoca | SortaVoca
```

#### 4.2.6 tblOtkup
**Purpose:** canonical procurement records  
**Primary key:** `OtkupID`  
**Foreign keys:** `KooperantID`, `StanicaID`, `KulturaID`, `VozacID`, `OtpremnicaID`, `ParcelaID`  
**Written by:** Operator after import from OTK sheets; desktop forms  
**Read by:** VBA, reports, PWA management  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** start of the main document chain; carries payment, parcel and optional `BrojZbirne` linkage; `OtpremnicaID` may remain blank temporarily and later be auto-linked only on a unique `(StanicaID, Datum, VozacID, Klasa)` match.

```text
OtkupID | Datum | KooperantID | StanicaID | KulturaID | VrstaVoca | SortaVoca | Kolicina | Cena | TipAmbalaze | KolAmbalaze | VozacID | BrojDokumenta | Novac | PrimalacNovca | Klasa | Stornirano | BrojZbirne | Isplaceno | DatumIsplate | OtpremnicaID | ParcelaID
```

#### 4.2.7 tblOtpremnica
**Purpose:** transport document records  
**Primary key:** `OtpremnicaID`  
**Foreign keys:** `StanicaID`, `VozacID`, `BrojZbirne` logical link  
**Written by:** Operator / auto-create from PWA import  
**Read by:** VBA, reports, transport flows  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** PWA numbering format `{StanicaNum}/{DDMM}-{seq}`; Klasa II transport rows are stored separately with zero ambalaža while class-level aggregation still rolls up by shared `BrojZbirne`

```text
OtpremnicaID | Datum | StanicaID | VozacID | BrojOtpremnice | BrojZbirne | VrstaVoca | SortaVoca | Kolicina | Cena | TipAmbalaze | KolAmbalaze | Klasa | Stornirano
```

#### 4.2.8 tblZbirna
**Purpose:** aggregate transport / buyer document  
**Primary key:** `ZbirnaID`  
**Foreign keys:** `VozacID`, `KupacID`, `BrojZbirne` external linkage  
**Written by:** Operator / Vozac import path  
**Read by:** VBA, management, later document chain  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** active schema includes buyer/cold storage context plus sync lineage for imported rows; one `BrojZbirne` may span multiple rows by class, desktop validation aggregates kg by class while ambalaža is validated on the combined document, and PWA-imported rows preserve `ClientRecordID` and `SyncSource` for dedupe/traceability.

```text
ZbirnaID | Datum | VozacID | BrojZbirne | KupacID | Hladnjaca | Pogon | VrstaVoca | SortaVoca | UkupnoKolicina | TipAmbalaze | UkupnoAmbalaze | Klasa | Stornirano | ClientRecordID | SyncSource
```

#### 4.2.9 tblPrijemnica
**Purpose:** cold storage receiving document  
**Primary key:** `PrijemnicaID`  
**Foreign keys:** `KupacID`, `BrojZbirne`, `VozacID`, `FakturaID`  
**Written by:** Operator; PWA hladnjača flow is still roadmap  
**Read by:** VBA, fakturisanje, SEF line build  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** links collection shipment to invoiceable intake; storno clears `Fakturisano` / `FakturaID` on the prijemnica itself, may mark linked faktura and faktura-stavke as orphaned via `OsirocenoOd`, and saving a replacement prijemnica can relink orphaned invoice lines from a storno predecessor with the same `BrojPrijemnice`

```text
PrijemnicaID | Datum | KupacID | BrojPrijemnice | BrojZbirne | VrstaVoca | SortaVoca | Kolicina | Cena | TipAmbalaze | KolAmbalaze | KolAmbVracena | VozacID | Klasa | Fakturisano | FakturaID | Stornirano
```

#### 4.2.10 tblFakture
**Purpose:** invoice header, payment status and SEF process state  
**Primary key:** `FakturaID`  
**Foreign keys:** `KupacID`, `SEFSubmissionIDLast` logical link  
**Written by:** Operator / `modFaktura` / SEF modules  
**Read by:** VBA, PWA management exports, SEF modules  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** desktop invoice creation is prijemnica-based; `Iznos` is the sum of selected `Prijemnica.Kolicina × Cena`, new rows start as unpaid and locally finalized, buyer avans may be auto-applied immediately after create, and `OsirocenoOd` can temporarily mark a faktura whose source prijemnica was stornirana until repair or storno is completed.

```text
FakturaID | BrojFakture | Datum | KupacID | Iznos | Status | DatumPlacanja | Stornirano | OsirocenoOd | SEFWorkflowState | SEFStatus | SEFDocumentId | SEFVersionNo | SEFPayloadHash | SEFSubmissionIDLast | SEFLastErrorCode | SEFLastErrorMessage | SEFSentAt | SEFLastSyncAt | PoslatNaSEF
```

#### 4.2.11 tblFakturaStavke
**Purpose:** invoice line items  
**Primary key:** `StavkaID`  
**Foreign keys:** `FakturaID`, `PrijemnicaID`  
**Written by:** Operator / `modFaktura`  
**Read by:** VBA, exports, SEF payload build  
**Soft delete:** `Stornirano = "Da"` on line storno workflows  
**Important invariants:** each created line stores the selected prijemnica tuple (`PrijemnicaID`, `Kolicina`, `Cena`, `Klasa`, `BrojPrijemnice`); invoice lines may temporarily become orphaned after storno/reversal workflows, and `RelinkFakturaStavke()` repairs them by matching a replacement prijemnica on the same `BrojPrijemnice` and clearing `OsirocenoOd`.

```text
StavkaID | FakturaID | PrijemnicaID | Kolicina | Cena | Klasa | BrojPrijemnice | Stornirano | OsirocenoOd
```

#### 4.2.12 tblNovac
**Purpose:** money movements across OM, kooperant, buyer and bank flows  
**Primary key:** `NovacID`  
**Foreign keys:** `PartnerID`, `OMID`, `KooperantID`, `FakturaID`, `OtkupID`  
**Written by:** Operator, banka mapping flows, avans allocation helpers  
**Read by:** VBA, kartice, saldo exports, bank mapping and allocation helpers  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** `Tip` semantics drive saldo direction, allocation and status refresh logic. Rows may stay unlinked or become explicitly linked to `FakturaID` or `OtkupID`; avans application is canonical split-capable, meaning a larger advance can be reduced in place while a new linked novac row is created for the consumed portion. `OsirocenoOd` remains available for reversal/relink workflows.

```text
NovacID | BrojDokumenta | Datum | Partner | PartnerID | EntitetTip | OMID | KooperantID | FakturaID | VrstaVoca | Tip | Uplata | Isplata | Napomena | Stornirano | OtkupID | OsirocenoOd
```

#### 4.2.13 tblAmbalaza
**Purpose:** packaging movement journal  
**Primary key:** `AmbalazaID`  
**Foreign keys:** `EntitetID`, `VozacID`, `DokumentID`  
**Written by:** Operator / document side effects  
**Read by:** VBA, reports  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** stores packaging effects of document lifecycle; rows carry optional `VozacID`, `DokumentID` and `DokumentTip` lineage and are queried only through non-stornirano balance helpers. Canonical saldo semantics are entity net movement by packaging type and driver-level `Izlaz / Ulaz / Saldo` reporting.

```text
AmbalazaID | Datum | TipAmbalaze | Kolicina | Smer | EntitetID | EntitetTip | VozacID | DokumentID | DokumentTip | Stornirano
```

#### 4.2.14 tblParcele
**Purpose:** parcel master data with GIS and meteo flags  
**Primary key:** `ParcelaID`  
**Foreign keys:** `KooperantID`  
**Written by:** Operator / geo editor flow  
**Read by:** VBA, PWA kooperant, meteo scheduler, management  
**Soft delete:** active state via `Aktivna`  
**Important invariants:** parcels without valid lat/lng are not shown on map; kooperant-scoped desktop parcel lookup returns `(ParcelaID, KatBroj, KatOpstina, Kultura, PovrsinaHa, DisplayLabel)` for downstream agro selection helpers

```text
ParcelaID | KooperantID | KatBroj | KatOpstina | Kultura | PovrsinaHa | GGAPStatus | Aktivna | GeoStatus | GeoSource | N_Coord | E_Coord | Lat | Lng | PolygonGeoJSON | MeteoEnabled | RizikStatus | DatumGeoUnosa | DatumAzuriranja | Napomena
```

#### 4.2.15 tblArtikli
**Purpose:** operator-controlled agrohemija master catalog  
**Primary key:** `ArtikalID`  
**Foreign keys:** none  
**Written by:** Operator and Management-only `createArtikal` endpoint for master-approved entries  
**Read by:** VBA, PWA kooperant/agrohemija, fiskalni auto-match  
**Soft delete:** `Aktivan` field  
**Important invariants:** private kooperant items do not belong here

```text
ArtikalID | Naziv | Tip | JedinicaMere | CenaPoJedinici | DozaPoHa | Kultura | Pakovanje | BarKod | Karenca | Aktivan
```

#### 4.2.16 tblMagacin
**Purpose:** warehouse / agrohemija movement journal  
**Primary key:** `MagacinID`  
**Foreign keys:** `ArtikalID`, `KooperantID`, `ParcelaID`, `DobavljacID`, `EntitetID`  
**Written by:** Operator  
**Read by:** VBA, management reports, agro debt/saldo helpers, PWA lager read models  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** `Tip` is movement direction; `CenaPoJedinici` is copied from current `tblArtikli` price on write; `Vrednost` is persisted as row-level quantity × price at save time

```text
MagacinID | Datum | ArtikalID | Tip | Kolicina | KooperantID | ParcelaID | BrojDokumenta | CenaPoJedinici | Vrednost | Napomena | Stornirano | DobavljacID | EntitetID
```

#### 4.2.17 tblConfig
**Purpose:** generic app config  
**Primary key:** `Parameter`  
**Foreign keys:** none  
**Written by:** Operator  
**Read by:** VBA, selected exports to PWA  
**Soft delete:** n/a  
**Important invariants:** not all keys are exported to PWA

```text
Parameter | Vrednost
```

#### 4.2.18 tblSEFConfig
**Purpose:** SEF integration and selected app settings  
**Primary key:** `ConfigKey`  
**Foreign keys:** none  
**Written by:** Operator  
**Read by:** VBA SEF, some PWA export filters  
**Soft delete:** n/a  
**Important invariants:** contains seller/integration identity fields

```text
ConfigKey | ConfigValue
```

#### 4.2.19 tblBankaImport
**Purpose:** bank import staging  
**Primary key:** `BankaImportID`  
**Foreign keys:** logical mapping to partners/fakture  
**Written by:** `modBankaImport` + `modBankaImport_PdfText` bank PDF staging pipeline  
**Read by:** bank mapping UI/logic  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** staged rows are append-only `BIM-*` imports; duplicate detection uses bank reference when available and otherwise falls back to `(BrojDokumenta, DatumTransakcije, Uplata, Isplata, Partner)` matching; import rows persist `IzvorFajl` and `ImportVreme` for traceability while `Obradjeno` follows the active lifecycle `"" -> Da | Skip | Error`; only non-stornirano rows whose `Obradjeno` is neither `Da` nor `Skip` are eligible for further mapping; header values (`BrojIzvoda`, `DatumIzvoda`, `BrojRacuna`) are copied onto every staged transaction row after successful PDF-text parse.

```text
BankaImportID | BrojDokumenta | DatumIzvoda | BrojRacuna | DatumTransakcije | Partner | PartnerKonto | Opis | Uplata | Isplata | Valuta | PozivNaBroj | SvrhaPlacanja | BankaReferenz | IzvorFajl | ImportVreme | Obradjeno | Stornirano
```

#### 4.2.20 tblPartnerMap
**Purpose:** learned bank-name-to-entity mapping  
**Primary key:** logical composite by `BankaName`  
**Foreign keys:** `PartnerID`, `OMID`  
**Written by:** successful bank mapping flows  
**Read by:** `modBankaMapiranje`  
**Soft delete:** n/a  
**Important invariants:** accelerates future bank auto-matching; lookup is exact-match on normalized bank name (trimmed, case-insensitive) and duplicate save attempts are treated as no-op success.

```text
BankaName | PartnerID | EntitetTip | OMID
```

#### 4.2.21 tblSEFSubmission
**Purpose:** SEF outbound attempt journal  
**Primary key:** `SEFSubmissionID`  
**Foreign keys:** `FakturaID`  
**Written by:** SEF outbound / persistence modules  
**Read by:** SEF UI and recovery logic  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** request/response and status history lives here

```text
SEFSubmissionID | FakturaID | VersionNo | WorkflowStateAtSubmit | CreatedAt | SubmittedAt | SubmissionStatus | PayloadHash | RequestFormat | RequestBody | ResponseBody | HttpStatus | ApiStatus | CorrelationId | SEFDocumentId | ErrorCode | ErrorMessage | OperatorName | Stornirano | FinishedAt
```

#### 4.2.22 tblSEFEventLog
**Purpose:** SEF event timeline  
**Primary key:** `SEFEventID`  
**Foreign keys:** `FakturaID`, `SEFSubmissionID`  
**Written by:** SEF modules  
**Read by:** frmSEF / audit trail  
**Soft delete:** `Stornirano = "Da"`  
**Important invariants:** event types reflect state transitions and HTTP lifecycle

```text
SEFEventID | FakturaID | SEFSubmissionID | EventTime | EventType | Message | Details | OperatorName | Stornirano
```

#### 4.2.23 Kartice sheet
**Purpose:** exported financial card records per kooperant, also production parsing source for Knjiga Polja  
**Primary key:** row-level export record (no explicit stable PK in docs)  
**Foreign keys:** `KooperantID`, `BrojParcele` → `ParcelaID` reference  
**Written by:** `ExportKarticeToGoogle`  
**Read by:** Kooperant PWA, Knjiga Polja, management exports  
**Soft delete:** exported view, not canonical delete domain  
**Important invariants:** rows with `Opis.startsWith('Otkup')` and `Zaduzenje > 0` are production parse candidates; `UKUPNO` rows are skipped in PWA

```text
KooperantID | Datum | BrojDok | BrojParcele | Opis | Zaduzenje | Razduzenje | Saldo
```

#### 4.2.24 FISKALNI-KOOP sheet
**Purpose:** private fiscal receipt line items per kooperant  
**Primary key:** `ClientRecordID`  
**Foreign keys:** `KooperantID`, optional `ArtikalID` to master or `PRIV-*` local item  
**Written by:** Kooperant PWA + GAS parser/save flows  
**Read by:** Kooperant PWA, fiskalni matching logic  
**Soft delete:** not emphasized; sync/status based lifecycle  
**Important invariants:** never treated as source for master artikli

```text
ClientRecordID | CreatedAtClient | SyncStatus | KooperantID | InvoiceNumber | Company | Date | VerificationUrl | Naziv | ArtikalID | ArtikalNaziv | Kolicina | JedCena | Ukupno | PDVStopa | Matched | ReceivedAt
```

#### 4.2.25 TROSKOVI-KOOP sheet
**Purpose:** kooperant private expense records for Knjiga Polja  
**Primary key:** `ClientRecordID`  
**Foreign keys:** `KooperantID`, `ParcelaID`  
**Written by:** Kooperant PWA  
**Read by:** Kooperant PWA, Knjiga Polja  
**Soft delete:** not emphasized; sync/status based  
**Important invariants:** 11 category taxonomy, plus auto-generated labor costs

```text
ClientRecordID | CreatedAtClient | SyncStatus | KooperantID | ParcelaID | Datum | Kategorija | Opis | Iznos | DokumentBroj | Napomena | ReceivedAt
```

#### 4.2.26 FiskalniMapiranje tab
**Purpose:** learned fiscal name mapping for auto-match  
**Primary key:** logical composite by fiscal name + kooperant scope  
**Foreign keys:** `ArtikalID`, `KooperantID`  
**Written by:** Management-only `saveFiskalniMapiranje`  
**Read by:** fiscal parser/matcher  
**Soft delete:** none emphasized  
**Important invariants:** match order is mapped → exact → contains → keywords

```text
FiskalniNaziv | ArtikalID | ArtikalNaziv | KooperantID | CreatedAt
```

### 4.3 Spreadsheet Tabs / Sheets
| Sheet / Tab | Purpose | Producer | Consumer | Canonical or Derived | Update trigger |
|---|---|---|---|---|---|
| Stammdaten / 13 tabs | shared PWA read model for master data, finance snapshots and PWA bootstrap data | `modStammdatenSync` | all PWA roles | derived from desktop master | manual / sync export |
| OTK-* | field otkup rows per station / otkupac flow | Otkupac-own-scope or Management override through authorized GAS `sync` | MasterSync, management, otprema | operational shared source before master import | PWA save/sync |
| VOZ-* | zbirna rows per driver | Vozac-own-scope or Management override through authorized GAS `syncZbirna` | MasterSync, transport views | operational shared source before master import | PWA save/sync |
| Kartice | financial card export and production parse source | ExportKarticeToGoogle | Kooperant PWA, Knjiga Polja | derived export | export job |
| MeteoLatest | current meteo state per parcela | scheduled GAS | PWA, management | canonical current meteo read model | scheduled fetch overwrite |
| MeteoHistory | historical meteo series | scheduled GAS | analysis / future reporting | canonical history | scheduled fetch append |
| KamionStatus | truck status / current dispatch state | Management or Vozac-own-status through authorized GAS endpoint | management | canonical dispatch status | dispatch updates |
| DispecerPlan | dispatch plan rows | Management through authorized GAS endpoint | management/dispatch UI | canonical dispatch planning | save/update/remove plan |
| MgmtReports | exported management aggregates | operator export | management PWA | derived | export |
| FISKALNI-{KooperantID} | private fiscal receipts | Kooperant-own-scope or Management through authorized GAS endpoint | kooperant | canonical within private fiscal domain | save parse/sync |
| TROSKOVI-{KooperantID} | private costs | Kooperant-own-scope or Management through authorized GAS endpoint | kooperant | canonical within private cost domain | sync |
| Users | auth/role export tab | `ExportUsers` | GAS/PWA auth flows | derived auth view | stammdaten export |
| LoginLog | GAS login attempt audit trail | GAS auth | Ops / debugging | operational audit log | login attempt |
| ErrorLog | GAS/PWA remote runtime error log | GAS `logError(...)` / PWA `logClientError` | Ops / debugging | operational observability log | runtime/client error |


### 4.3.1 Google Spreadsheet Export and Sync Contracts

The current desktop snapshot adds these explicit Google-side sheet contracts:

- **Stammdaten workbook provisioning:** `SyncStammdatenToGoogle()` finds or creates a spreadsheet named `Stammdaten`, persists `GOOGLE_STAMMDATEN_SHEET_ID`, and provisions 13 tabs: `Kooperanti`, `Kulture`, `Parcele`, `Config`, `Users`, `Fakture`, `FakturaStavke`, `SaldoOMDetail`, `Stanice`, `Kupci`, `Vozaci`, `Artikli`, `MagacinKoop`.
- **Kartice export workbook:** `ExportKarticeToGoogle()` maintains a dedicated `Kartice` spreadsheet whose `Sheet1` rows carry `KooperantID | Datum | BrojDok | BrojParcele | Opis | Zaduzenje | Razduzenje | Saldo`.
- **MgmtReports export workbook:** `ExportMgmtReports()` maintains a dedicated spreadsheet with tabs `SaldoOM`, `SaldoKupci`, `OtkupPoOM` and `PredatoPoKupcu`.
- **OTK-* sheet contract:** desktop master-sync now follows the GAS-first `COLUMNS` order. Canonical OTK Google Sheet header is `ClientRecordID | ServerRecordID | CreatedAtClient | UpdatedAtClient | UpdatedAtServer | SyncStatus | DeviceID | OtkupacID | Datum | KooperantID | KooperantName | VrstaVoca | SortaVoca | Klasa | Kolicina | Cena | TipAmbalaze | KolAmbalaze | ParcelaID | VozacID | Napomena | ReceivedAt`.
- **OTK writeback semantics:** desktop import treats `SyncStatus = "Synced"` as pending-for-master, writes back `"Synced>Master"` on successful master import, and may also write `Duplicate` or `SyncError[:reason]` for skipped or invalid rows. Writeback targets are `Sheet1!F` for `SyncStatus` and `Sheet1!B` for `ServerRecordID`.
- **VOZ-* sheet contract:** desktop zbirna master-sync now follows the GAS-first `ZBIRNA_COLUMNS` order. Canonical VOZ Google Sheet header is `ClientRecordID | ServerRecordID | CreatedAtClient | UpdatedAtClient | UpdatedAtServer | SyncStatus | VozacID | Datum | KupacID | KupacName | VrstaVoca | SortaVoca | KolicinaKlI | KolicinaKlII | TipAmbalaze | KolAmbalaze | Klasa | OtkupRecordIDs | ReceivedAt | BrojZbirne`.
- **VOZ plain-text formatting rule:** GAS applies `ensurePlainTextColumn(...)` on VOZ sheet columns whose business values may contain `/` and would otherwise be auto-coerced by Google Sheets into dates. At minimum this includes `TipAmbalaze`, and where present also `BrojZbirne`.
- **VOZ writeback semantics:** successful desktop zbirna imports write back `Sheet1!F = Synced>Master` and `Sheet1!B = ZbirnaID`; duplicate and validation failures write `Duplicate` or `SyncError[:reason]` through the same status channel.
- **Plain-text packaging rule:** Google Sheet columns that can carry slash-delimited business values such as `TipAmbalaze` (for example `12/1`) must be kept in plain-text formatting to avoid Google Sheets date coercion.

### 4.3.2 Active GAS Workbook / Sheet Topology (Code.gs snapshot)

The current supplied `Code.gs` backend makes the following workbook and sheet topology explicit:

- **Master Drive folder:** all role-scoped operational spreadsheets are created under `MASTER_FOLDER_ID`.
- **Role-scoped spreadsheet naming:** `OTK-<OtkupacID>`, `VOZ-<VozacID>`, `AGRO-<KooperantID>`, `TRETMAN-<KooperantID>`, `OPREMA-<KooperantID>` and `FISKALNI-<KooperantID>` are lazily created Google spreadsheets with a single primary sheet and header bootstrap.
- **Shared operational workbooks:** `Stammdaten`, `Kartice`, `MgmtReports`, `LoginLog` and `ErrorLog` are looked up by workbook name inside the same master folder.
- **Sheet registry tab:** role-scoped and shared workbook lookup may use `SheetRegistry` inside `Stammdaten`; `getOrCreateSheet(...)` appends newly created workbook IDs when the registry tab exists, `getSheetRegistry()` returns the name-to-file-id map, and stale or missing registry entries fall back to classic folder lookup.
- **Geo split workbook:** parcel geometry and meteo persistence use separate constants `GEO_SPREADSHEET_ID` and `GEO_SHEET_PARCELE = 'Parcele'`; `MeteoLatest` and `MeteoHistory` are maintained in that geo workbook.
- **Shared tab expectations:** the current backend explicitly reads or writes tabs including `Users`, `Kooperanti`, `Kulture`, `Config`, `Parcele`, `Stanice`, `Kupci`, `Vozaci`, `Artikli`, optional `Oprema`, `MagacinKoop`, `KamionStatus`, `DispecerPlan`, `WarRoomDemand`, `Izdavanje`, `SaldoOM`, `SaldoKupci`, `OtkupPoOM`, `PredatoPoKupcu`, `SaldoOMDetail`, `Fakture` and `FakturaStavke`.
- **Fallback report rule:** management read helpers first look in workbook `MgmtReports`, then fall back to workbook `Stammdaten` for report tabs when the dedicated workbook or sheet is missing.

### 4.4 IndexedDB / Client Stores
| Store | Purpose | Key / indexes | Offline role | Synced | Retention |
|---|---|---|---|---|---|
| otkupi | field otkup local queue | `clientRecordID`; `syncStatus`, `datum` | Otkupac | Yes | keep active sync records; prune future enhancement |
| stammdaten | cached shared read model | `key` | all roles | refreshed from server | cached across sessions |
| tretmani | kooperant treatment queue | `clientRecordID`; `syncStatus`, `datum`, `parcelaID` | Kooperant | Yes | active local archive |
| troskovi | kooperant cost queue | `clientRecordID`; `syncStatus`, `datum` | Kooperant | Yes | active local archive |
| zbirne | driver zbirna queue | `clientRecordID`; `syncStatus` | Vozac | Yes | active local archive |

Additional active client-store contract from the supplied runtime snapshot:

- `openDB()` is now a recovery-first window-scoped bootstrap helper rather than a minimal raw open wrapper.
- IndexedDB open now includes timeout guarding, `onblocked` handling, `onversionchange` connection close behavior and an explicit recovery path through `resetIndexedDb()`.
- the active store bootstrap is centered on canonical stores `CONFIG.STORE_NAME`, `CONFIG.STAMM_STORE`, `tretmani`, `troskovi` and `zbirne`; legacy kooperant `agromere` store semantics are no longer part of the active runtime contract and are treated as legacy-cleanup migration state.
- store provisioning is normalized through schema-style helper logic during upgrade instead of only ad-hoc create/delete branches.
- low-level IndexedDB access is normalized through guarded promise wrappers `dbPut`, `dbGet`, `dbGetAll`, `dbGetByIndex` and `dbDelete`, with explicit missing-store / missing-index failure semantics.
- architecture consequence: frontend runtime no longer depends on manual browser storage clearing as the primary IndexedDB recovery mechanism for ordinary upgrade/store-shape failures.

### 4.5 DTOs / Payload Contracts
| DTO / Payload | Produced by | Consumed by | Versioned | Notes |
|---|---|---|---|---|
| LoginResponse | GAS login action | PWA auth/bootstrap | implicit | returns token, role, entity context |
| Otkup sync payload | Otkupac PWA | GAS `sync` | implicit | includes duplicate-aware VozacID update behavior |
| Zbirna payload | Vozac PWA | GAS `syncZbirna` | implicit | writes to VOZ sheet |
| SEF invoice snapshot | VBA `BuildSEFInvoiceDto` | `SerializeUBLInvoice` / SEF outbound | Yes via `VersionNo` | formal UBL payload source |
| Meteo record | scheduled GAS | MeteoLatest/MeteoHistory + PWA | implicit | includes risk and spray window evaluation |
| Fiskalni parse/save payload | Kooperant PWA / GAS parser | GAS save endpoints | implicit | QR URL or image base64 input |


### 4.5.1 GAS Action Authorization Matrix

| Action | Public | Auth required | Allowed roles | Ownership check | Lock required | Notes |
|---|---:|---:|---|---|---:|---|
| `login` | Yes | No | n/a | n/a | No | Auth bootstrap only |
| `logClientError` | Yes / pre-auth exception | No, but enriches from token when valid | n/a | Optional entity fallback | No | Best-effort observability; must not block field operation |
| `getParcelGeo` | Yes, current acknowledged gap | No | n/a | n/a | No | Public geo/meteo bridge |
| `getParcelMeteo` | Yes, current acknowledged gap | No | n/a | n/a | No | Public geo/meteo bridge |
| `getParcelMeteoLatest` | Yes, current acknowledged gap | No | n/a | n/a | No | Public geo/meteo bridge |
| `getAllMeteoLatest` | Yes, current acknowledged gap | No | n/a | n/a | No | Public geo/meteo bridge |
| `saveParcelPolygon` | Yes | Yes | Public exception | n/a | Not locked in normal auth path | Intentional v6.10 public/pre-auth GIS write exception |
| `sync` | No | Yes | Otkupac, Management | Otkupac must match `otkupacID` | Yes | OTK write |
| `syncAgromere` | No | Yes | Kooperant, Management | Kooperant must match `kooperantID` | Yes | AGRO write |
| `syncZbirna` | No | Yes | Vozac, Management | Vozac must match `vozacID` | Yes | VOZ write |
| `syncTretman` | No | Yes | Kooperant, Management | Kooperant must match `kooperantID` | Yes | Treatment write |
| `syncOprema` | No | Yes | Kooperant, Management | Kooperant must match `kooperantID` | Yes | Equipment write |
| `syncTrosak` | Yes | Kooperant / Management | kooperantID | Yes | Yes | Active v6.12 batch sync endpoint for `troskovi`; idempotent by `ClientRecordID` |
| `saveOtkupniListPdf` | No | Disabled | n/a | n/a | No | Returns `FEATURE_DISABLED` until real GAS PDF generation exists |
| `uploadPdf` | No | Yes | Otkupac, Management | Otkupac must match `otkupacID` when supplied | Yes | Drive upload |
| `saveWarRoomDemand` | No | Yes | Management | n/a | Yes | Shared demand write |
| `removeWarRoomDemand` | No | Yes | Management | n/a | Yes | Shared demand write |
| `updateDemandPrimljeno` | No | Yes | Management | n/a | Yes | Shared demand status write |
| `updateKamionStatus` | No | Yes | Vozac, Management | Vozac forced to `tokenData.entityID` | Yes | Driver status write |
| `saveDispecer` | No | Yes | Management | n/a | Yes | Dispatch planning write |
| `updateDispecer` | No | Yes | Management | n/a | Yes | Dispatch planning write |
| `removeDispecer` | No | Yes | Management | n/a | Yes | Dispatch planning write |
| `saveIzdavanje` | No | Yes | Management | n/a | Yes | Warehouse/issuance write |
| `parseFiskalniImage` | No | Yes | Kooperant, Management | Kooperant must match `kooperantID` if supplied | No | Parsing-only, quota-sensitive |
| `parseFiskalni` | No | Yes | Kooperant, Management | Kooperant must match `kooperantID` if supplied | No | Parsing-only |
| `saveFiskalni` | No | Yes | Kooperant, Management | Kooperant must match `kooperantID` | Yes | Private fiscal write |
| `saveFiskalniMapiranje` | No | Yes | Management | n/a | Yes | Shared mapping write |
| `createArtikal` | No | Yes | Management | n/a | Yes | Master artikal creation |

---

## 5. Business Flows

Za svaki flow obavezno navesti:
- trigger
- actors
- source data
- korake
- writes
- downstream efekte
- rollback / recovery ponašanje

### 5.1 End-to-End Document Flow
- **Trigger:** otkup on field or desktop entry
- **Actors:** Otkupac → Vozac / Operator → Hladnjača / Operator → Fakturisanje / SEF
- **Source data:** master data, otkup rows, buyer data, parcel linkage
- **Flow:** `Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF`
- **Writes:** `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke`, SEF journals
- **Downstream:** Novac, Ambalaza, Kartice, reports, status updates
- **Rollback / recovery:** storno is executed per document/entity with only local side effects (ambalaža storno, novac unlink, prijemnica release or orphan marking) and no automatic chain-wide cascade; SEF still has a dedicated recovery state machine

### 5.2 Otkup Flow
- **Trigger:** Otkupac opens new otkup in PWA or desktop form
- **Actors:** Otkupac, Kooperant
- **Source data:** kooperant, parcela, kultura, config pricing, vozac (optional)
- **Steps:** scan/select kooperant → choose parcela/vozac → choose vrsta/sorta/klasa → enter qty/price/ambalaza → local save → sync to OTK sheet → master import
- **Writes:** local `otkupi` store, OTK sheet, later `tblOtkup`
- **Authorization:** GAS `sync` requires a valid token. Otkupac may write only to own `OtkupacID`; Management override is allowed only by explicit endpoint authorization.
- **Downstream:** Otkupni list, otprema grouping, payment eligibility, kartica export
- **Rollback / recovery:** sync retry if offline; desktop storno resets packaging and money links


#### 5.2.1 Desktop Otkup Entry Form Flow
- **Trigger:** operator opens desktop `frmOtkup` and records procurement directly in the master workbook.
- **Actors:** Operator
- **Source data:** stanice, kooperanti, parcele, kulture, vozaci, ambalaža types, explicit quantity/price/cash inputs.
- **Steps:** choose station through hidden-ID `cmbOtkupnoMesto` → filter kooperanti by station → optionally choose parcela and auto-backfill fruit context from parcela kultura → parse date/quantity/price/cash through `modParse` → enter Klasa I quantity/price and optional Klasa II quantity/price → run duplicate check on `BrojDokumenta` → save the full business operation through `SaveOtkupMulti_TX`.
- **Writes:** `tblOtkup`, `tblAmbalaza`, optionally `tblNovac`, plus avans allocation updates executed inside the same high-level transaction boundary.
- **Downstream:** direct desktop procurement chain, packaging outflow, cash payout trace and later otpremnica linking/traceability.
- **Rollback / recovery:** `SaveOtkupMulti_TX` snapshots `tblOtkup`, `tblAmbalaza` and `tblNovac`; Klasa I, optional Klasa II, packaging, cash payout and avans allocation commit or rollback together. Forms do not call `SaveOtkup_TX`, `SaveNovac_TX` and `ApplyAvansToOtkup_TX` separately for one operator action.
- **Fallback rule:** this desktop flow remains supported for users who do not use PWA or for repair/exception cases, even though the preferred operating model is PWA-first.



### 5.3 Otpremnica / Transport Flow
- **Trigger:** Otkupac otprema tab assigns driver to selected otkupi, or operator directly enters otpremnica in desktop `frmDokumenta`
- **Actors:** Otkupac, Vozac, Operator
- **Source data:** unassigned OTK records, driver master data, station, fruit/sort, price and ambalaža
- **Steps:** open otprema or desktop document form → choose station/driver and fruit context → enter quantity/price/ambalaža → optional Klasa II entry → run duplicate pre-check → desktop save through `SaveOtpremnicaMulti_TX` so both classes are atomic → later sync/import where applicable
- **Writes:** local otkup records, OTK sheet updated rows, and/or direct `tblOtpremnica` rows
- **Downstream:** transport visibility, driver workload, zbirna candidate set, later zbirna kg/ambalaža validation
- **Rollback / recovery:** duplicate-aware update logic, form-level validation, one-transaction dual-class save, storno otpremnica cascades packaging effects

### 5.4 Zbirna Flow
- **Trigger:** Vozac creates zbirna in PWA, or operator directly enters zbirna in desktop `frmDokumenta`
- **Actors:** Vozac, Operator
- **Source data:** assigned otkupi or previously entered otpremnice, buyer/cold storage selection, class totals, ambalaža totals
- **Steps:** load assigned records or enter desktop summary → choose kupac/hladnjača/pogon → aggregate by classes and ambalaža → PWA persists one local `zbirne` record with `clientRecordID`, sync metadata and aggregated `otkupRecordIDs` → GAS writes/updates one row in `VOZ-<VozacID>` using the canonical VOZ sheet contract, including dedicated `BrojZbirne` column support and plain-text formatting for slash-based business values → desktop `ImportZbirneFromPWA()` imports rows whose Google-side `SyncStatus` equals `Synced`, validates payload, dedupes by `ClientRecordID`, generates `BrojZbirne`, writes `tblZbirna`, then cascade-links `BrojZbirne` to matching `tblOtkup` and `tblOtpremnica` rows through `OtkupRecordIDs`; manual desktop dual-class save uses `SaveZbirnaMulti_TX`
- **Writes:** local `zbirne` store, VOZ sheet, `tblZbirna`, and linked `BrojZbirne` updates in `tblOtkup` / `tblOtpremnica`
- **Authorization:** GAS `syncZbirna` requires a valid token. Vozac may write only to own `VozacID`; Management override is allowed only by explicit endpoint authorization. `updateKamionStatus` forces Vozac callers to `tokenData.entityID` regardless of client-supplied `vozacID`.
- **Downstream:** prijemnica creation, buyer chain, fakturisanje preparation, later shortage analytics by zbirna or by otpremnica share
- **Rollback / recovery:** sync retry on PWA; desktop duplicate pre-check on `ClientRecordID`; later storno zbirna does not auto-storno dependent otpremnice/prijemnice, but may create orphan-warning state for documents that still point to the stornirana zbirna

### 5.5 Prijemnica Flow
- **Trigger:** Hladnjača / operator enters intake against zbirna
- **Actors:** Operator today; future hladnjača role planned
- **Source data:** zbirna, buyer, measured quantities, packaging return, optional Klasa II pricing
- **Steps:** select/enter broj zbirne → compute live `CalculateManjakPreview()` against saved zbirna totals → show average crate weight → enter received qty, price and packaging return → optional Klasa II entry → run duplicate pre-check → save through `SavePrijemnicaMulti_TX`, which persists ambalaža effects and runs guarded `RelinkFakturaStavke()` inside the same rollback scope
- **Writes:** `tblPrijemnica`, `tblAmbalaza`, and relink updates on `tblFakturaStavke` / `tblPrijemnica` when orphaned invoice lines are repaired
- **Downstream:** fakturisanje lines, invoiceable stock, packaging reconciliation, shortage visibility
- **Rollback / recovery:** storno prijemnica triggers ambalaza reversal path and guarded faktura relink handling; live warning thresholds highlight shortage before save

#### 5.5.1 Transport / Intake Reconciliation Analytics
- **Zbirna saved-state validation:** `ValidateZbirna(brojZbirne)` compares summed otpremnice kg/ambalaža against already saved zbirna rows.
- **Pre-save desktop validation:** `ValidateZbirnaPreUnosa(brojZbirne, inputKgKlI, inputKgKlII, inputAmb)` validates Klasa I and Klasa II separately while keeping ambalaža aggregated.
- **Shortage totals:** `CalculateManjak(brojZbirne)` returns zbirna kg, prijemnica kg, `ManjakKg` and `ManjakPct` across all active rows.
- **Shortage preview:** `CalculateManjakPreview(...)` includes unsaved pending prijemnica quantities so desktop UI can warn before save.
- **Proportional shortage allocation:** `CalculateManjakByOtpremnica(brojZbirne)` distributes shortage by otpremnica share and derives `ManjakRSD` using row price.
- **Crate-weight helpers:** `CalculateProsekGajbe(brojOtp)` and `CalculateProsekGajbeByZbirna(brojZbirne)` provide average kg per crate for transport and intake review.

### 5.6 Faktura Flow
- **Trigger:** operator fakturisanje from prijemnice
- **Actors:** Operator
- **Source data:** non-stornirano prijemnice, kupac master data, selected `(PrijemnicaID, Kolicina, Cena, Klasa, BrojPrijemnice)` tuples, open buyer avans
- **Steps:** select invoiceable prijemnice → call `CreateFaktura[_TX]()` → generate new `FAK-*` plus human invoice number `N/GODINA` → compute total as the sum of canonical `tblPrijemnica.Kolicina × tblPrijemnica.Cena` for the selected `PrijemnicaID` values → append faktura header with unpaid/local-finalized defaults → append one stavka row per selected tuple → mark source prijemnice as `Fakturisano="Da"` and link `FakturaID` → auto-run buyer avans allocation → optionally print through the `FakturaSablon` worksheet template
- **Writes:** `tblFakture`, `tblFakturaStavke`, `tblPrijemnica`, and optionally `tblNovac` through automatic avans application
- **Downstream:** buyer saldo, SEF readiness, printable invoice output, bank reconciliation target
- **Rollback / recovery:** `CreateFaktura_TX` snapshots fakture/stavke/prijemnice/novac; storno faktura frees prijemnice and novac links, while storno prijemnica marks faktura and matching stavke orphaned via `OsirocenoOd` until a replacement prijemnica is relinked or the invoice is storned


#### 5.6.1 Desktop Fakturisanje UI Flow
- **Trigger:** operator opens `frmFakturisanje` for one selected kupac.
- **Actors:** Operator
- **Source data:** kupci, prijemnice for that kupac, optional already-fakturisane visibility, payment summary dictionary by faktura, existing fakture for print selection.
- **Steps:** choose kupac through hidden-ID combo → load kupac-specific prijemnice via `GetPrijemniceByKupac()` and `ExcludeStornirano()` → optionally include already-fakturisane rows as read-only/history context → multi-select only invoiceable rows → block stornirane or already-fakturisane rows before create → confirm the total amount → create faktura via `CreateFaktura_TX` → refresh `cmbFaktura` and select the newly created FakturaID for printing if applicable.
- **Writes:** `tblFakture`, `tblFakturaStavke`, `tblPrijemnica`, and optionally `tblNovac` through automatic avans application.
- **Downstream:** operator-facing invoice assembly, payment visibility per already-fakturisana row, explicit invoice print selection and fast handoff into SEF operations.
- **Rollback / recovery:** `CreateFaktura_TX` snapshots fakture/stavke/prijemnice/novac; create fails if selected prijemnica is stornirana, already fakturisana, already carries `FakturaID`, duplicated in the selection, or if any append/update fails. `PrintFaktura` uses guarded `tblFakturaStavke` schema, no legacy `KulturaID` dependency and blocks active-print of stornirana faktura rows.



### 5.7 SEF Flow
- **Trigger:** finalized invoice ready for SEF submit, later refresh, cancel or storno actions
- **Actors:** Operator / desktop SEF modules
- **Source data:** faktura header, faktura stavke, prijemnica-linked line context, buyer master data, seller/config values from `tblSEFConfig`, previous submission payload where technical retry is allowed
- **Steps:** `ValidateFakturaForSEF()` gates sendability → build fresh snapshot via `BuildSEFInvoiceDto()` or reuse last failed submission body when retrying a technical failure → validate DTO dates (`DeliveryDate <= InvoiceDate`) → serialize active outbound payload as UBL XML via `SerializeUBLInvoice()` using DTO-owned dates and compute `SEFPayloadHash` → if needed move faktura from `WF_LOCAL_FINALIZED` / `WF_SEF_TECH_FAILED` to `WF_SEF_READY` → create or reuse `tblSEFSubmission` row → move faktura to `WF_SEF_SENDING` in a dedicated TX → call `SubmitUBLInvoice(ublXml, requestId:=SEFSubmissionID)` outside TX → require successful responses to include `SEFDocumentId` or normalize them to `SEF_TECH_FAILED` / `MISSING_SEF_DOCUMENT_ID` → persist HTTP/API result in `tblSEFSubmission` and update faktura to `WF_SEF_ACCEPTED`, `WF_SEF_SENT`, `WF_SEF_REJECTED` or `WF_SEF_TECH_FAILED` while storing the exact external `response.apiStatus` in `SEFStatus` → later `RefreshSEFStatus_TX()` updates the exact external status (`SENT`, `NEW`, `DRAFT`, `ACCEPTED`, `REJECTED`, `STORNO`, `CANCELLED`, etc.) without blindly forcing a new local workflow transition.
- **Writes:** `tblFakture` SEF state/status fields, `tblSEFSubmission`, `tblSEFEventLog`


#### 5.7.1 Desktop SEF Management UI Flow
- **Trigger:** operator opens `frmSEF` to inspect or act on one faktura.
- **Actors:** Operator
- **Source data:** faktura SEF fields, buyer metadata, SEF event log rows, current local workflow state and external `SEFStatus`.
- **Steps:** guarded activation removes chrome and applies theme without crashing the UI lifecycle → explicit two-column faktura combo setup loads visible `FakturaID` / `BrojFakture` values → selected faktura loads workflow, external status, document ID, version, last-error labels and event log rows → button policy enables only actions legal for the current workflow/status → operator may send, refresh, prepare rejected invoice for resubmit, cancel, storno, recover one stuck `SEF_SENDING` faktura, refresh pending fakture, or recover all stuck sending fakture.
- **Writes:** `tblFakture`, `tblSEFSubmission`, `tblSEFEventLog` through called SEF orchestration/recovery modules.
- **UI safety:** `btnPosalji_Click` uses a cleanup path so the send button cannot remain disabled after validation/confirmation early exits or errors. Cancel/storno paths require operator comment plus explicit confirmation before calling destructive SEF endpoints.
- **Downstream:** operator-visible state machine control, audit visibility, resend/recovery decisions and safer handling of retry/recovery/cancel/storno scenarios.
- **Rollback / recovery:** all mutation actions delegate to `_TX` modules; the form itself only reloads visible state after action completion. `SEF_SENDING` is recoverable through refresh or explicit technical-failure repair, and `WF_SEF_REJECTED` can be prepared for corrected resubmission.

#### 5.7.2 SEF v6.7 Hardened Flow Invariants
- **Internal vs external status:** `SEFWorkflowState` remains the local state-machine field, while `SEFStatus` stores the exact latest external status returned by SEF API. Refresh operations may update `SEFStatus` without changing `SEFWorkflowState`.
- **Submit status persistence:** `SendInvoiceToSEF_TX` persists `response.apiStatus` into `SEFStatus`; internal workflow constants are not used as external status values on submit result persistence.
- **Successful submit guard:** a successful HTTP/API response without `SEFDocumentId` is normalized into a technical failure with `MISSING_SEF_DOCUMENT_ID` evidence and must not produce `WF_SEF_SENT` / `WF_SEF_ACCEPTED`.
- **Idempotent refresh:** repeated refresh of `ACCEPTED`, `REJECTED`, `SENT`, `NEW` or `DRAFT` must not fail just because the workflow is already in the target or compatible local state. Pending/non-final status handling routes through `ApplySEFStateOrRefreshOnly` to avoid accidental backwards transitions from final local states.
- **Final-state protection:** a local final state is not moved backwards to `SEF_SENT` merely because the external API returns a pending-like value.
- **Sync-error recovery:** `SEF_SYNC_ERROR` can recover through the allowed state path when a later status refresh returns a valid external final status.
- **Payload tax and total consistency:** DTO header totals, line totals and UBL tax percent use the same tax configuration source; hardcoded 10% VAT is not canonical. Header/line net, VAT and gross totals are fail-fast checked within the active rounding tolerance before UBL submission.
- **Config validation:** `ValidateSEFConfig` checks `SEF_BASE_URL`, `SEF_API_KEY` and URL scheme before send; mapper config such as `SEF_PAYMENT_DUE_DAYS` is parsed with explicit guard/default behavior.
- **HTTP boundary:** SEF HTTP calls remain outside desktop transaction scopes; only local state preparation and local result persistence are transactional.
- **Rate limit handling:** HTTP 429 is represented as `RATE_LIMITED`, optionally carrying `Retry-After`; automated backoff scheduling is roadmap, not active behavior.
- **Persistence/error guard:** faktura SEF state, submission result and event log writes use schema guards and `RequireUpdateCell`; local `RaiseUpdateError` is not canonical. Service/recovery error handlers capture original Err data before rollback/logging.
- **Recovery:** stuck `SEF_SENDING` rows are repaired through `RecoverStuckSEFSendingInvoice` or the guarded batch recovery helper.
- **Parser hardening:** simple JSON response extraction supports numeric-or-string document IDs and tolerant boolean parsing for fields such as `accepted`, while full nested/escaped JSON parsing remains outside the lightweight client scope.
- **Transition coverage:** `RunSEFStateTransitionSuite` validates allowed and blocked local workflow transitions, including terminal `WF_SEF_STORNO` behavior.

### 5.8 Novac / Banka Flow
- **Trigger:** manual finance entry, payout, buyer payment, OM station payout/input, avans allocation, otkup-link reset or bank import
- **Actors:** Operator
- **Source data:** fakture, open otkup blocks, bank statement PDFs plus normalized `pdftotext` output, OM pools, partner-map rows, ambalaža side-effects
- **Steps:** import/enter money movement → validate novac direction and entity context → classify by tip → map to kooperant/kupac/OM/faktura/otkup → validate remaining amount / available OM avans where relevant → optionally auto-allocate avans to open faktura/otkup → split oversized avans rows when only part of the amount is consumed → use fail-fast update guards for avans link/split/status writes → recompute otkup paid status in both directions when money links are added or reset.
- **Writes:** `tblNovac`, `tblBankaImport`, `tblPartnerMap`, `tblOtkup`, and optionally `tblAmbalaza` / faktura status fields.
- **Downstream:** salda, kartice, payment status on otkup/fakture, OM avans balance, customer return packaging, bank-name learned mapping and fruit-type aggregation for buyer payments.
- **Rollback / recovery:** `_TX` wrappers use snapshot rollback over affected tables; `AppendRow`/`UpdateCell` side effects are rollback-safe only when the caller has snapshotted the affected table before the write. Read-only helpers exclude stornirano rows; update helpers use full table indexes and skip stornirano rows manually.

#### 5.8.1 OM Ulaz Flow
- **Trigger:** operator enters station-side OM ulaz in `frmDokumenta`
- **Actors:** Operator
- **Source data:** selected station, optional driver, station-filtered kooperant list, open otkup blocks, OM avans saldo
- **Steps:** choose OM → optionally register incoming ambalaža with `TrackAmbalaza` → optionally choose kooperant and open otkup block → validate that entered novac does not exceed open remainder → if payout is from OM avans, validate station saldo → resolve `Tip` (`NOV_KES_OTKUPAC_KOOP`, `NOV_VIRMAN_FIRMA_KOOP`, `NOV_VIRMAN_AVANS_KOOP`, `NOV_KES_FIRMA_OTKUPAC`) → save through `SaveOMUlaz_TX` so ambalaža, novac and otkup-status side effects are atomic → refresh OM saldo label
- **Writes:** `tblNovac`, `tblAmbalaza`, otkup status fields
- **Downstream:** OM saldo, kooperant payout tracking, open-otkup closure logic
- **Rollback / recovery:** duplicate check on document number, no save if remainder or OM avans limits are violated, and one rollback restores ambalaža/novac/otkup-status effects if any linked write fails

#### 5.8.2 Izlaz Kupci Flow
- **Trigger:** operator enters buyer-side payment/advance and ambalaža return in `frmDokumenta`
- **Actors:** Operator
- **Source data:** kupac, open fakture for selected buyer, optional driver, returned ambalaža
- **Steps:** choose kupac → load open fakture into combo → optional faktura linking through hidden `FakturaID` → save through `SaveKupciIzlaz_TX` as `NOV_KUPCI_UPLATA` or `NOV_KUPCI_AVANS` → update faktura status if payment closes invoice → optionally register returned ambalaža through `TrackAmbalaza` inside the same TX
- **Writes:** `tblNovac`, `tblAmbalaza`, faktura status fields
- **Downstream:** buyer saldo, receivables closure, ambalaža stock reconciliation
- **Rollback / recovery:** duplicate check on document number, hidden-ID faktura binding, one rollback across buyer novac/ambalaža/faktura status, and refill of open-fakture list after save

#### 5.8.3 Storno and Reconciliation Flow
- **Trigger:** operator runs storno from `frmDokumenta`
- **Actors:** Operator
- **Source data:** active business documents looked up by number or ID
- **Steps:** choose type (`Otkup`, `Otpremnica`, `Zbirna`, `Prijemnica`, `Faktura`, `Novac`) → resolve active ID where needed → explicit confirmation dialog → call type-specific `_TX` storno function → refresh orphan-document warning
- **Writes:** soft-delete flags plus dependent reverse effects in linked tables
- **Downstream:** document chain rollback, relink requirements for dependent transport/intake docs, refreshed operator warnings
- **Rollback / recovery:** `CheckVerwaisteDokumente()` surfaces active otpremnice/prijemnice still pointing to a stornirana zbirna without an active replacement; later replacement prijemnice may also trigger faktura-stavka relink repair

#### 5.8.4 Avans Allocation Semantics
- **Trigger:** operator applies buyer or kooperant avans to a concrete faktura or otkup
- **Actors:** Operator
- **Source data:** open fakture, open otkupi, existing novac rows, linked payment aggregates
- **Steps:** compute outstanding remainder → iterate eligible unconsumed avans rows in existing order → fully link rows that fit into the remainder → if a row is larger than the remaining debt, reduce the original row and append a new linked row for the consumed portion → refresh paid status when remainder reaches zero
- **Writes:** `tblNovac`, `tblFakture`, `tblOtkup`
- **Downstream:** exact residual avans visibility, correct open-item reporting, deterministic payment linkage
- **Rollback / recovery:** both faktura-side and otkup-side allocation paths have `_TX` wrappers; `ResetNovacOtkupLink()` exists as a repair helper for otkup-side relinking

#### 5.8.5 Banka Inbox Import Flow
- **Trigger:** operator runs desktop bank-inbox import over statement PDFs
- **Actors:** Operator, `modBankaImport`, `modBankaImport_PdfText`
- **Source data:** PDFs in `APP_BANKA_INBOX`, UTF-8 text extracted via local `pdftotext`, parsed statement header fields and transaction lines
- **Steps:** ensure inbox/processed/error folders exist → list inbox PDFs → extract raw text per file with `pdftotext -raw -nopgbrk -enc UTF-8` → parse statement header (`BrojIzvoda`, `DatumIzvoda`, `BrojRacuna`) and transaction rows → stage non-duplicate rows into `tblBankaImport` with source-file metadata → move file to processed or error folder based on outcome
- **Writes:** `tblBankaImport`, filesystem moves from inbox to processed/error directories
- **Downstream:** staged reconciliation workload for bank mapping, partner-map learning, later novac creation and faktura/otkup linking
- **Rollback / recovery:** `ImportBankaInbox_TX()` wraps the staging table in a transaction; file moves are collision-safe via unique target naming, and parse/empty-text failures are routed to the error folder instead of partially polluting the import table

#### 5.8.6 Bank PDF Text Parse Contract
- **Trigger:** called from per-file bank import after successful PDF-to-text extraction
- **Actors:** `modBankaImport_PdfText`
- **Source data:** Komercijalna Banka statement text produced by local `pdftotext`
- **Steps:** normalize page/line breaks → extract statement header fields from text lines → collect transaction blocks starting from numeric sequence markers → within each block parse execution date, partner, account, zaduženje/odobrenje, šifra, svrha, poziv na broj and bank reference → clean trailing summary noise such as `Ukupno za...` or dangling dates/references
- **Writes:** in-memory 10-column transaction arrays consumed by `ParseBankaIzvodForImport()`
- **Downstream:** canonical staging row construction for `tblBankaImport`
- **Rollback / recovery:** missing mandatory header fields aborts staging for that file; parser attempts to recover references and `PozivNaBroj` from `Svrha` before final cleanup
- **Outbound status model:** `SEFWorkflowState = WF_SEF_SENT` is a local process milestone and can coexist with external `SEFStatus = DRAFT`, `SENT`, `STORNO` or other statuses returned by SEF until refresh converges the local workflow to a final state.
- **Delivery-date guard:** `DeliveryDate` is the latest linked prijemnica date for the faktura. UBL generation is blocked locally if `DeliveryDate > InvoiceDate`.
- **Serializer source of truth:** `SerializeUBLInvoice()` uses `dto.InvoiceDate` and `dto.DeliveryDate`; it must not call delivery-date lookup helpers again.
- **v6.6 live evidence:** baseline live submit has been proven with HTTP 200, `SubmissionStatus=SENT`, `SEFDocumentId` persistence and successful repeated status refresh. Negative tests include invalid receiver rejection persistence and local delivery-date validation blocking.


#### 5.8.7 Banka Review / Mapping UI Flow
- **Trigger:** operator opens desktop `frmBankaImport` to review staged `tblBankaImport` rows.
- **Actors:** Operator
- **Source data:** open bank-import rows, learned `tblPartnerMap`, kupci/kooperanti/OM master data, open fakture, open otkup blocks.
- **Steps:** load only non-stornirano rows whose `Obradjeno` is not `Da`/`Skip` → auto-set `MapTip` from payment direction (`Kupac` for uplata, `Kooperant` for isplata) → show detail panel (`Partner`, `PozivNaBroj`, `Opis`, `Svrha`, amounts) → render auto-preview using the same matching heuristics as `modBankaMapiranje` → either run single-row auto map, batch auto map, manual target mapping, block-specific kooperant mapping, skip, or refresh.
- **Writes:** `tblBankaImport`, `tblNovac`, `tblPartnerMap`, and optionally linked `tblOtkup` / `tblFakture` side effects through the called mapping modules.
- **Downstream:** operator-visible reconciliation queue, explainable auto-matching, faster exception handling for bank imports.
- **Rollback / recovery:** all save/skip actions use `_TX` wrappers from mapping modules; form itself is review/orchestration only and repaints the queue after every action.

### 5.9 Agrohemija Flow
- **Trigger:** management/operator issues agrohemija or kooperant records treatment consumption
- **Actors:** Operator, Kooperant, Management
- **Source data:** artikli, parcele, magacin state, dosage configuration
- **Steps:** choose kooperant → resolve parcela list via `GetParceleByKooperant()` → choose artikal → compute recommendation as `DozaPoHa × PovrsinaHa` via `CalculatePreporuka()` → validate/record warehouse movement through `SaveMagacin[_TX]()` or treatment save
- **Writes:** `tblMagacin`, tretmani store/sheet, read models
- **Downstream:** stock state, kooperant agro debt, supplier/issuance reporting, knjiga polja cost and lager calculations
- **Rollback / recovery:** transactional desktop save for warehouse; kooperant sync retry offline-first


#### 5.9.1 Desktop Agrohemija Issuance Flow
- **Trigger:** operator opens desktop `frmAgrohemija` and records agrohemija izlaz to a kooperant.
- **Actors:** Operator
- **Source data:** kooperanti, parcele, artikli, dosage config (`Doza`, `Pakovanje`, `JM`, `Cena`), current agro debt indicators.
- **Steps:** choose kooperant → load parcel display strings plus hidden `(ParcelaID, PovrsinaHa)` arrays → choose one or more parcels → choose artikal → compute recommendation from total selected hectares via `CalculatePreporuka()` → if packaging size exists, round up to whole packages and prefill issue quantity → validate quantity and package multiples → add line into an in-memory korpa item array → on finish, require `BrojDok`, start a transaction, and persist one `MAG_IZLAZ` row per korpa item through `SaveMagacin()` with semicolon-separated parcela references.
- **Writes:** `tblMagacin` (`MAG_IZLAZ`), while debt labels derive from `GetAgrohemijaDug() - GetAgroAbzug()`.
- **Downstream:** kooperant agro debt, parcel-linked issuance history, stock depletion, later knjiga polja and management reporting.
- **Rollback / recovery:** form batches all issue lines inside one explicit `clsTransaction`; any failed save rolls back the entire korpa.

#### 5.9.2 Desktop Agrohemija Receipt / Magacin Ulaz Flow
- **Trigger:** operator records warehouse receipt/purchase in `frmAgrohemija` ulaz mode.
- **Actors:** Operator
- **Source data:** artikli, supplier selector, entered price/quantity, generated or manual warehouse document number.
- **Steps:** choose artikal → prefill default price and show configured dose → enter quantity and price → add lines to a separate ulaz korpa array → on finish require `BrojDok`, optionally supplier, and persist one `MAG_ULAZ` row per korpa line through `SaveMagacin()` within a single transaction.
- **Writes:** `tblMagacin` (`MAG_ULAZ`) with explicit per-line value based on entered quantity × price.
- **Downstream:** stock replenishment, supplier value reporting, later issuance availability for kooperanti.
- **Rollback / recovery:** the whole ulaz basket is wrapped in one transaction; UI resets only after commit.

### 5.9.3 Kooperant Digitalni Agronom / Tretman Flow
- **Trigger:** kooperant opens `tab-agromere` or confirms a new treatment.
- **Actors:** Kooperant
- **Source data:** kooperant parcels, `stammdaten.magacinkoop`, merged equipment list, parcel geo/meteo data, treatment history cache, config-driven labor price and karenca metadata.
- **Steps:** reset agronom state → load current lager and deduplicated oprema list → populate parcels → start geo watch → optionally auto-detect/suggest nearby parcel → load meteo strip and active karenca → choose mera → for `Zastita/Prihrana` choose preparat from current lager and apply smart dosage → optionally block or override unfavorable meteo → start/stop timer → persist local treatment record with timer/geo/karenca/sync metadata → attempt online sync and refresh history.
- **Writes:** IndexedDB `tretmani`; GAS `syncTretman`; direct `syncOprema` only for newly added equipment when online.
- **Authorization:** Kooperant-side GAS writes including `syncAgromere`, `syncTretman`, `syncOprema` and `saveFiskalni` require a valid token and must match `tokenData.entityID` to the requested `kooperantID`, except where Management is explicitly allowed as override.
- **Downstream:** agronomy history, harvest blocking, knjiga polja bilans, seasonal potrošnja and parcel-level treatment context.
- **Rollback / recovery:** local-first persistence, merged treatment cache invalidation after save/sync, online retry path through sync engine and geo-watch cleanup on tab exit.

### 5.10 Parcela / GIS Flow
- **Trigger:** operator maintains parcela geo or kooperant opens map
- **Actors:** Operator, Kooperant
- **Source data:** `tblParcele`, geo editor, polygon/point data
- **Steps:** point or polygon capture → save geo → export to PWA → render map with markers/polygons
- **Writes:** `tblParcele` geo fields, polygon store via the intentional public/pre-auth `saveParcelPolygon` GAS endpoint


#### 5.10.1 Desktop Traceability / Otkupni Blokovi Flow
- **Trigger:** operator opens `frmOtkupniBlokovi` to repair missing otkup links or print lot trace for a zbirna.
- **Actors:** Operator
- **Source data:** unresolved `tblOtkup` rows without `OtpremnicaID`, candidate `tblOtpremnica` rows, selected `tblZbirna` number, reverse-trace payload from `TraceByZbirna()`.
- **Steps:** load unresolved queue → optionally run `AutoLinkOtkupOtpremnica()` → for one unresolved row inspect same-day/same-station otpremnica candidates and manually write `OtpremnicaID` if needed → independently select `BrojZbirne` to load reverse trace rows → render and export `SledljivostSablon` as PDF with totals and shortage summary.
- **Writes:** manual `OtpremnicaID` repairs on `tblOtkup` and exported trace PDF files on disk.
- **Downstream:** restored procurement-to-shipment integrity, printable lot trace package and easier manual cleanup of ambiguity left unresolved by auto-link rules.
- **Rollback / recovery:** auto-link only writes when key match is unique; manual linking is explicit operator action; trace printing is read-only.
- **Downstream:** meteo eligibility, geofencing, map and dispatch context
- **Rollback / recovery:** geo clear functions exist; polygon save remains a public/pre-auth exception in v6.10 and must be protected operationally; future auth gating remains a security roadmap item

### 5.11 Meteo Flow
- **Trigger:** scheduled jobs 4x daily (`00:00`, `06:00`, `12:00`, `18:00` Europe/Belgrade) plus on-demand parcel reads.
- **Actors:** GAS scheduler, kooperant parcel UI, digital agronom.
- **Source data:** parcels with `MeteoEnabled = "Da"`, parcel geo centroid/polygon data, Open-Meteo forecast payloads, culture-specific threshold map.
- **Threshold model:** active risk thresholds are culture-specific for at least `Visnja`, `Jabuka`, `Sljiva`, `Kruska`, `Breskva` and `Malina`, with `_default` fallback for all other cultures.
- **Steps:** group parcels by rounded lat/lng bucket → batch fetch Open-Meteo → fallback single fetch with retry when batch fails → calculate current/daily/hourly derived model → assess frost/heat/rain/disease risk → calculate contiguous 72-hour spray-safe windows → append history and overwrite latest snapshot.
- **Writes:** append-only `MeteoHistory` and current-state `MeteoLatest`, including serialized risk items, spray windows and compact daily forecast JSON.
- **Downstream:** kooperant parcel cards, parcel detail meteo tab, home-dashboard alerts, digital agronom meteo validation and spray timing suggestions.
- **Rollback / recovery:** retries with delay, cached-first read path for `getParcelMeteo()`, batch fetch to reduce API pressure and live fallback when cached snapshot is stale.

### 5.12 Fiskalni Lager / Fiskalni Scanner Flow
- **Trigger:** kooperant scans or photographs a fiscal receipt inside Knjiga Polja / Lager.
- **Actors:** Kooperant, PWA fiscal UI, GAS parser
- **Source data:** live QR scan or photo-captured QR, fiscal verification URL, `stammdaten.artikli`, learned `FiskalniMapiranje` rows
- **Steps:** try native `BarcodeDetector` live scan → fall back to photo capture with client-side resize → call authenticated `parseFiskalni` / `parseFiskalniImage` → render parsed lines with match confidence → optionally map to existing artikli or create private `PRIV-*` staged article IDs → save only checked rows → optionally request Management-approved learned mapping where applicable
- **Writes:** `FISKALNI-<KooperantID>` sheet rows through `saveFiskalni`; shared `FiskalniMapiranje` only through Management-only `saveFiskalniMapiranje`; no write into master `Artikli` for private new receipt items
- **Authorization:** parsing endpoints require authenticated Kooperant or Management. Kooperant save/parse payloads must match the caller's `KooperantID`; shared fiscal mapping writes are Management-only.
- **Downstream:** Knjiga Polja lager/proizvodnja awareness, private fiscal history and future auto-match improvement
- **Rollback / recovery:** duplicate receipt detection, checked-row validation that every saved line has an `artikalID`, cancel/reset of staged receipt and photo fallback when native scanning is unavailable

### 5.13 Knjiga Polja Flow
- **Trigger:** kooperant opens Knjiga Polja or records a new trošak.
- **Actors:** Kooperant
- **Source data:** `stammdaten.kartice`, merged tretmani cache/server, merged local+server `troskovi`, `stammdaten.magacinkoop`, config parameter `CenaRadaSat`
- **Steps:** derive proizvodnja from `Kartice` rows whose `Opis` starts with `Otkup` and skip `UKUPNO` → merge treatment and trošak datasets → derive synthetic labor cost from treatment duration → compute agrohemija cost from treatment consumption × `WAC/CenaPoJedinici` → calculate bilans → render pregled/proizvodnja/troškovi/lager sections plus KPI strip and seasonal consumption list
- **Writes:** primarily read-only aggregation plus local-first expense writes; `kpSaveTrosak()` writes pending rows into IndexedDB `troskovi`, and online `syncTrosak` is active in GAS for launch sync.
- **Authorization:** `syncTrosak` is Kooperant/Management scoped. Kooperant callers may sync only their own `KooperantID`; Management may override through explicit endpoint authorization.
- **Downstream:** parcel economics, profitability, consumption overview and management-style season summary for the kooperant
- **Rollback / recovery:** `mergeOfflineRecords` fallback, cache invalidation after saves, immediate bilans reload after trošak write and preservation of derived labor rows as non-authoritative `_auto` items

### 5.14 Dispatch / Planning Flow
- **Trigger:** management dispatch screen
- **Actors:** Management
- **Source data:** supply from stations, transport capacity/status, demand entries
- **Steps:** select supply card → choose truck → choose demand destination → save plan → update status during execution
- **Writes:** `DispecerPlan`, `KamionStatus`, demand sheet/tab
- **Downstream:** transport visibility, station/buyer balancing
- **Rollback / recovery:** update/remove plan endpoints; planning never overwrites raw OTK authority


### 5.15 Storno and Recovery Flow
- **Trigger:** operator initiates storno or business correction on an already saved desktop record
- **Actors:** Operator
- **Source data:** target entity ID or broj dokumenta plus linked ambalaža, novac, prijemnica and faktura rows
- **Steps:** validate target with `CanStorno()` so it exists and is not already stornirano → mark only the addressed entity/document as `Stornirano="Da"` → run entity-specific side effects: otkup resets linked money references and stornoes its packaging rows; otpremnica stornoes only its packaging rows; zbirna stornoes all active rows with the same `BrojZbirne` but leaves dependent otpremnice/prijemnice active; prijemnica clears invoice linkage and marks faktura / stavke orphaned; faktura stornoes header + lines, frees prijemnice and removes novac faktura links; novac storno recomputes invoice payment state if needed
- **Writes:** `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke`, `tblNovac`, `tblAmbalaza`
- **Downstream:** orphan warnings, relink eligibility for replacement prijemnice, refreshed faktura/otkup payment status, preserved document lineage without physical delete
- **Rollback / recovery:** every storno path has its own `_TX` wrapper; there is no automatic cross-document cascade beyond the explicit side effects coded for that entity


### 5.16 Otkup Traceability and Auto-Link Flow
- **Trigger:** operator runs traceability repair/reporting or the system evaluates unresolved procurement rows.
- **Actors:** Operator, `modSledljivost`, `frmOtkupniBlokovi`, `modOtkup`.
- **Preferred source model:** PWA-first. In the ideal operating model, most otkup blocks and zbirne context arrive from PWA with driver/zbirna-related context already captured, reducing manual desktop linking.
- **Fallback model:** VBA remains fully supported for operators who do not use PWA, for exception handling, and for repair of imported or manually-entered rows.
- **Source data:** active `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblParcele`, `tblKooperanti`.
- **Steps:** identify active otkupi without `OtpremnicaID` → build otpremnica index by `StanicaID|Datum|VozacID|Klasa` → auto-link only uniquely matching rows through the guarded auto-link transaction → leave ambiguous matches unresolved for manual review → allow operator manual linking from `frmOtkupniBlokovi` → use `TraceByZbirna()` to walk `Zbirna → Otpremnica → Otkup → Kooperant/Parcela` lineage when reverse-tracing shipments → render/export `SledljivostSablon` PDF.
- **Writes:** `tblOtkup.OtpremnicaID` only during auto-link or explicit manual link; trace PDF export writes a file and does not alter business tables.
- **Downstream:** lower orphan volume in document chain, stronger shipment provenance, richer audit/export context with kooperant BPG and parcela attributes.
- **Rollback / recovery:** `AutoLinkOtkupOtpremnica_TX` snapshots `tblOtkup`; every link write uses `RequireUpdateCell`. Manual linking is explicit operator action and must also require the target otkup row and checked update. Unresolved rows stay visible through `GetUnlinkedOtkupi()` for later manual correction.
- **Hardening boundary:** v6.4 hardening preserves existing matching rules, return shapes, PWA/desktop trace flow and PDF layout. Any future change to candidate filtering, additional PWA-driven matching, or trace output structure is a functional enhancement and requires explicit design approval.



### 5.17 Desktop Startup Resilience Flow
- **Trigger:** workbook/app startup
- **Actors:** Operator, splash shell, `frmOtkupAPP`, `modJournal`, `modLogError`
- **Source data:** workbook file, backup folder, journal folder, log folder, today's table state
- **Steps:** optional branded splash shell appears → create timestamped backup copy → purge old backups/logs/journals by retention window → inspect today's journal files against live table row counts → hand off to `frmOtkupAPP` which surfaces orphan/recovery warnings if present → continue startup without blocking business use even if maintenance helpers fail
- **Writes:** filesystem only (`Backup/`, `Journal/`, `Log/`)
- **Downstream:** local crash-recovery evidence, remote-support logs, bounded disk growth, visible orphan-warning surface on the home shell
- **Rollback / recovery:** helper failures are best-effort and must not stop the workbook; warnings remain advisory for operator follow-up


### 5.18 PWA-First / VBA-Fallback Traceability Model
- **Preferred path:** field users should use PWA wherever possible. PWA-created otkup records should carry enough context to support later traceability: station, driver, date, class, parcela when available, and document/zbirna context when available.
- **Desktop fallback:** the desktop VBA system must remain operational without PWA. `frmOtkup`, `frmDokumenta` and `frmOtkupniBlokovi` provide manual entry, document creation, repair linking and trace PDF generation.
- **Canonical chain:** trace output remains based on `Zbirna → Otpremnica → Otkup → Kooperant/Parcela`.
- **Bridge field:** `tblOtkup.OtpremnicaID` is the canonical link from raw/PWA/manual otkup records into the document chain. If missing, the record is not part of canonical trace output until auto-link or manual linking resolves it.
- **Repair role:** `frmOtkupniBlokovi` is not only a report form; it is the repair/audit surface for rows that PWA import or manual desktop entry did not fully connect.
- **Future enhancement boundary:** PWA-specific optimizations such as suggested zbirna, stricter candidate filtering, or direct PWA-provided document links are roadmap items, not silent hardening changes.


---

## 6. Application Architecture

### 6.1 VBA Module Overview
| Module | Responsibility | Inputs | Outputs | Tables touched | Critical invariants |
|---|---|---|---|---|---|
| modMain | desktop runtime coordinator for app startup, shell launch and shutdown | `Workbook_Open`, operator close/exit actions | `StartApp`, `ShutdownApp`, lifecycle logs, Excel visibility state | filesystem/logs, workbook runtime | `Workbook_Open` delegates only to `StartApp`; `ShutdownApp` restores Excel visibility and logs normal app shutdown |
| modConfig | constants and config access | config tables | lookup values | many | naming and enum semantics |
| modDataAccess | single approved sheet↔array bridge plus ID/duplicate/journal helpers | table names, column names, row arrays, search values | arrays, `ListObject` access, append/update ops, generated IDs, duplicate messages | all | no direct business logic may bypass this surface; `AppendRow()` journals writes and `CheckDuplicate()` skips stornirano rows |
| modDokumenta | document save logic, validation helpers, shortage analytics and orphan repair | forms, document data | IDs, linked docs, reconciliation aggregates, orphan scans | Otpremnica/Zbirna/Prijemnica/Ambalaza/FakturaStavke/Fakture/Novac/Otkup | Multi_TX wrappers for dual-class save paths, atomic OM/Kupci money-packaging wrappers, class-aware validation, guarded orphan detection and prijemnica relink repair |
| modAmbalaza | packaging movement journal and saldo helpers | document side effects, entity IDs, driver IDs, date filters | ambalaža rows, entity balances, driver saldo views | tblAmbalaza | canonical `TrackAmbalaza`, stornirano-safe reads, saldo semantics by packaging type |
| modNovac / payment logic | money movement, avans allocation, OM saldo and payment status rules | payment entries, bank partner names, open blocks, fakture, otkupi | money records, payment aggregates, updated faktura/otkup status | tblNovac, tblOtkup, tblFakture, tblPartnerMap | novac tip semantics, TX wrappers, split-avans allocation, stornirano-safe readers, no direct UI messages in business code |
| modFaktura | desktop invoice creation, numbering, printing and payment-status refresh | kupac, selected prijemnice tuples, faktura template sheet, buyer avans | `FAK-*` rows, invoice lines, linked prijemnice, printed output | tblFakture, tblFakturaStavke, tblPrijemnica, tblNovac | prijemnica-based amount calculation, `N/GODINA` numbering, auto-avans apply, local-finalized default |
| modStorno | per-entity soft-delete and repair side effects | entity IDs / broj zbirne, linked packaging/payment/invoice rows | stornirano flags, released links, orphan markers, recomputed statuses | tblOtkup, tblOtpremnica, tblZbirna, tblPrijemnica, tblFakture, tblFakturaStavke, tblNovac, tblAmbalaza | no chain-wide cascade; each storno path owns only its explicit local side effects |
| modMagacin / agro stock logic | parcela lookup, dosage recommendation, warehouse journal, stock/debt and supplier reports | kooperant/parcela/artikal context, article dosage/price, warehouse movement params | `MAG-*` rows, stock aggregates, issuance and supplier summaries | tblMagacin, tblArtikli, tblParcele, tblKooperanti | write-time price valuation, stornirano-safe warehouse reads, `Ulaz-Izlaz` stock semantics |
| modSEFValidation / modSEFPayload / modSEFOutbound / modSEFPersistence / modSEFHttp / modSEFStatus | e-faktura validation, payload, transport, persistence and refresh orchestration | fakture, stavke, prijemnice, config, previous submissions | UBL XML, payload hashes, submission/event journals, exact SEF status snapshots, state transitions | tblFakture, tblSEFSubmission, tblSEFEventLog, tblSEFConfig | exact external `SEFStatus` is separate from local `SEFWorkflowState`; outbound submit uses 3-phase TX split and requestId = submissionID |
| modBankaImport | bank inbox import and staging orchestration | inbox PDFs, parsed import arrays | staged `BIM-*` rows, processed/error file routing | tblBankaImport | folder triage, duplicate detection, source-file lineage |
| modBankaImport_PdfText | `pdftotext` extraction and Komercijalna Banka statement text parsing | statement PDFs, raw UTF-8 text lines | normalized 10-column transaction arrays and extracted header fields | none directly (feeds `tblBankaImport`) | parser assumes `pdftotext -raw -nopgbrk -enc UTF-8` layout and cleans summary noise |
| modBankaMapiranje | reconcile staged bank entries into canonical money flow through manual/auto mapping | staged rows, partner maps, fakture, otkupi | novac rows, updated processed states, learned partner mappings | tblBankaImport, tblNovac, tblPartnerMap, tblFakture, tblOtkup | `Obradjeno` lifecycle, not-processed gate, faktura/otkup resolution, block split semantics |
| modOtkup | procurement core save/read/saldo logic | otkup form data, kultura/artikal lookups, optional parcela/zbirna context | `OTK-*` rows, station/kooperant views, station saldo aggregates | tblOtkup, tblAmbalaza, tblKulture | otkup is the chain origin, packaging side-effect on save, saldo helper still marks bank/isporuka subtraction as open TODO |
| modSledljivost | auto-link and reverse-trace across procurement and shipment chain | active otkupi, otpremnice, zbirna identifiers, parcela/kooperant lookups | repaired `OtpremnicaID` links, unresolved lists, reverse-trace arrays | tblOtkup, tblOtpremnica, tblParcele, tblKooperanti | auto-link only on unique `(Stanica, Datum, Vozac, Klasa)` match; reverse trace enriches with BPG/parcela context |
| modLogError | non-blocking local runtime logging | source name, message, error context, lifecycle events | dated `.log` lines under `Log/` | filesystem | log writes must never break the app; daily file retention is 30 days |
| modJournal | append journaling, startup backup and crash-recovery checks | appended row arrays, workbook path, current table sizes | per-table daily CSV journals, dated workbook backups, advisory recovery warnings | filesystem, all tables indirectly | journaling hangs off `AppendRow()`, backup/journal maintenance is best-effort, retention is 30 days |
| modArrayUtils | in-memory filtering, sorting, grouping and aggregation over 2D arrays | table arrays, filter collections, group/sort specs | filtered/sorted/grouped arrays and numeric summaries | none directly (feeds all modules) | replaces sheet copy/paste/sort; works on 1-based in-memory arrays only |
| modParse | shared parsing/normalization for desktop user input | raw textbox strings | typed `Double`, `Long`, `Date` values or parse failure | none | one canonical parser for Serbian/European numeric/date input; forms do not duplicate parser functions |
| modSchemaGuard | fail-fast guards for schema and critical row updates | table/column names, row updates | required column indices, enforced `UpdateCell` success | all tables indirectly | `RequireColumnIndex`, `RequireColumns`, `RequireUpdateCell`; missing columns or failed updates are errors |
| modComboBinding | shared ComboBox display/hidden-ID binding | MSForms ComboBox controls, table/display/id columns | ComboBox rows with visible text and hidden stable IDs | master-data tables indirectly | entity selection uses hidden ID column and `GetComboID`, not display-text lookup |
| shared desktop helper module | display parsing, combo-fill helpers, generic `ExcludeStornirano()`, orphan warnings and aggregate dictionaries | display strings, combo targets, table arrays, optional filters | driver/kooperant display lists, filtered arrays, manjak dictionaries, warning text | tblVozaci, tblKooperanti, tblOtkup, tblOtpremnica, tblZbirna, tblPrijemnica | helpers must tolerate both display formats and `ExcludeStornirano()` must remain safe on tables without a `Stornirano` column |
| modStammdatenSync | export desktop master, finance and report read-models to Google Sheets | master/derived tables, Google config, report helpers | Stammdaten/Kartice/MgmtReports spreadsheets and tab payloads | many master + derived tables | provisioning + export semantics, 13-tab Stammdaten contract, filtered active-data exports |
| modMasterSync | import OTK-* Google Sheets into desktop master and auto-create grouped otpremnice | Google Sheets OTK tabs, OAuth helpers, master lookups | master otkup rows, sync-status writeback, optional auto-otpremnice | tblOtkup, tblOtpremnica, tblAmbalaza | imports only pending synced rows, duplicate-aware clientRecordID handling, grouped auto-otpremnica semantics |
| modGeoUtils / modClipboard | geo capture helpers | coords/clipboard | normalized geo data | tblParcele | point/polygon integrity |
| modGoogleAuth | OAuth2 token bootstrap, refresh and Google config persistence for VBA integrations | Google client credentials, auth code, token responses, config keys | access tokens, refresh tokens, expiry timestamps, configured-auth state | tblSEFConfig | browser-code flow, auto-refresh with expiry buffer, token writes centralized through `SetConfigValue()` |
| modGoogleSheets | REST wrapper for Google Sheets/Drive from VBA | spreadsheet IDs, tab names, 2D arrays, folder IDs | created spreadsheets/tabs, cleared tabs, read/write array payloads | Google APIs only (desktop config + logs indirectly) | writes arrays as strings, clears before overwrite, folder move + sheet lookup semantics |
| modTheme | shared dark desktop theme and control styling layer for VBA forms | forms, frames, controls, semantic button intents, status labels | consistent palette/fonts, recursive control styling, hover/active helpers, field enable/disable states | none directly | presentation-only helper; palette, typography and control semantics must stay centralized here |
| desktop parcel geo point helper | save/clear parcela point geometry and UTM34→WGS84 conversion | parcela row index, UTM easting/northing | updated parcel geo fields, meteo enable flags, timestamps | tblParcele | point save rounds lat/lng to 6 decimals, sets `GeoStatus=point`, `GeoSource=selenium`, and enables meteo |
| frmDokumenta | unified desktop document shell for Otpremnica, Zbirna, Prijemnica, OM Ulaz, Izlaz Kupci and Storno | parse-safe user input + hidden-ID master lookups + open fakture/otkupi | Multi_TX saves/storno calls, live validation labels, warning state | tblOtpremnica, tblZbirna, tblPrijemnica, tblNovac, tblAmbalaza, tblFakture | one-form six-frame design, duplicate pre-checks, class-II gating, hidden-ID combo binding, orphan-warning refresh |
| desktop stammdaten launcher form | desktop menu shell for master-data sections | operator menu choice | opens `frmStammdaten` with section tag and hides parent shells | master-data forms only | title-barless selector for `Kooperanti`, `Stanice`, `Kupci`, `Vozaci`, `Artikli`, `Parcele` |
| frmOtkupAPP | primary desktop home shell and navigation container | operator startup, navigation clicks, orphan warning text | fullscreen shell layout, child-form launches, warning surface, modeless maticni popup, Excel/exit actions | reads many tables indirectly through helper/report forms | responsive shell; alert card surfaces `CheckVerwaisteDokumente()` and close path delegates to app shutdown/save logic |
| frmStammdaten | universal desktop master-data CRUD shell driven by `frm.Tag` | requested entity section, table rows, master lookups, parcel geo hooks | add/edit/list/reset/navigation across kooperanti, stanice, kupci, vozaci, parcele and artikli | tblKooperanti, tblStanice, tblKupci, tblVozaci, tblParcele, tblArtikli | one form configures itself per entity; parcel mode surfaces geo/risk summary and geo-action hooks |
| desktop splash form | branded startup handoff shell | app/version labels and timer delay | auto-unload and show `frmOtkupAPP` after ~2s | none | presentation-only startup form; no business writes |
| frmFakturisanje | desktop invoice assembly shell over prijemnice | selected kupac, filtered prijemnice, print/SEF actions | `CreateFaktura_TX`, `PrintFaktura`, `frmSEF` launch | tblPrijemnica, tblFakture, tblFakturaStavke, tblNovac | kupac-driven list with optional already-fakturisane visibility and uplata summary per faktura |
| frmOtkup | desktop-only procurement entry shell | procurement input, parcela lookup, class-II toggle, optional cash payout | `SaveOtkup_TX`, `SaveNovac_TX`, `ApplyAvansToOtkup_TX` | tblOtkup, tblNovac, tblAmbalaza | left-side otkup-only form after document shell split; duplicate guard, parcela-kultura warning and optional dual-class save |
| frmSEF | desktop SEF operations shell for one faktura | selected faktura, workflow/status/event review, send/refresh/cancel/storno/recovery actions | SEF `_TX` orchestration modules and event loading helpers | tblFakture, tblSEFSubmission, tblSEFEventLog | button-state logic follows local workflow plus external SEF status |
| frmOtkupniBlokovi | desktop traceability and manual link shell | unlinked otkup queue, otpremnica candidates, zbirna trace selection | `AutoLinkOtkupOtpremnica`, manual `OtpremnicaID` update, `TraceByZbirna`, `PrintTracePDF` | tblOtkup, tblOtpremnica, tblZbirna, tblPrijemnica | combines unresolved link repair with printable lot trace view |

### 6.1.1 Desktop Form Shells

`frmDokumenta v2.1` is the canonical unified operator form for desktop-side document entry and reversal. Its active architecture is:

- **shell model:** one form, six frames, no MultiPage
- **frames:** `fraOtpremnica`, `fraZbirna`, `fraPrijemnica`, `fraOMUlaz`, `fraIzlazKupci`, `fraStorno`
- **startup behavior:** guarded activation with `On Error GoTo EH`, theme application, chrome/title-bar removal before emptying caption, lookup prefill and storno warning refresh
- **cascading lookups:** `VrstaVoca -> SortaVoca`, hidden-ID `Kupac -> Hladnjaca + open fakture`, hidden-ID `OtkupnoMesto -> kooperanti + OM saldo`; open fakture combo stores hidden `FakturaID`
- **reactive validation:** class-II toggles enable extra fields; parse-safe qty changes recompute totals, zbirna validation, manjak preview and average crate weight
- **save orchestration:** form validates/parses input and calls business wrappers (`SaveOtpremnicaMulti_TX`, `SaveZbirnaMulti_TX`, `SavePrijemnicaMulti_TX`, `SaveOMUlaz_TX`, `SaveKupciIzlaz_TX`); the form does not own transaction internals
- **navigation pattern:** form returns to `frmOtkupAPP`; hover/reset styling is UI-only and not business-authoritative

`frmAgrohemija` is the canonical desktop warehouse/agrohemija operation form. Its active architecture is:

- **dual-basket model:** one form manages separate in-memory `korpa izlaz` and `korpa ulaz` arrays before commit.
- **issuance shell:** kooperant selector drives parcela multi-select, recommendation label, debt label and issue basket assembly.
- **receipt shell:** dedicated article/price/quantity entry path drives supplier-linked warehouse receipts.
- **reactive calculation:** quantity/value labels recompute from article price; recommendation recalculates from selected hectares and packaging rules.
- **navigation pattern:** form can return to `frmOtkupAPP`; title bar is removed via the same desktop chrome pattern used by other forms.

`frmBankaImport` is the canonical desktop reconciliation review form for staged bank rows. Its active architecture is:

- **queue shell:** list-driven review of only still-open `tblBankaImport` rows.
- **detail panel:** selected row expands partner, reference, purpose, amount and mapping-direction context.
- **manual override controls:** operator can choose target type, partner and optional otkup block before committing a mapping.
- **preview-first behavior:** mapping preview uses the same heuristics as backend bank mapping helpers so UI explanation matches actual save logic.
- **batch affordances:** one-click auto-map current row or all open rows, plus explicit skip and refresh paths.


The desktop stammdaten launcher form is the canonical menu shell for opening master-data maintenance sections. Its active architecture is:

- **section menu model:** one lightweight launcher exposes buttons for `Kooperanti`, `Stanice`, `Kupci`, `Vozaci`, `Artikli` and `Parcele`.
- **child-open behavior:** each button hides the current shell plus `frmOtkupAPP`, opens `frmStammdaten`, and passes the requested section through `frm.Tag`.
- **auto-unload pattern:** the launcher unloads itself on deactivate unless it is intentionally opening a child form.
- **chrome pattern:** title bar is removed via the same `ThunderDFrame` desktop chrome helper used across other forms.

`frmOtkupAPP` is the canonical desktop home shell. Its active architecture is:

- **fullscreen responsive shell:** the form resizes itself to the Excel application window and recomputes header/sidebar/content-card layout at startup.
- **navigation container:** sidebar buttons open the active operator forms for otkup, dokumenta, agrohemija, izveštaji, fakturisanje, marža, sledljivost, Excel mini-shell and exit/save actions.
- **warning surface:** on activation the form calls `CheckVerwaisteDokumente()` and renders unresolved orphan warnings in a dedicated alert card instead of silently hiding them.
- **header affordances:** the shell provides a close icon, branding/logo area and a `Maticni podaci` entry point that opens `frmMaticniPodaci` modeless near the header button.
- **shutdown discipline:** form-control close delegates to `ShutdownApp`; explicit exit saves the workbook, restores Excel visibility and then quits the application; the workbook-level `Workbook_BeforeClose` handler also routes through `ShutdownApp`, so log finalization and shell teardown happen regardless of which exit path the operator takes.

`frmStammdaten` is the canonical universal desktop master-data maintenance form. Its active architecture is:

- **tag-driven setup:** `frm.Tag` selects entity-specific setup for `Kooperanti`, `Stanice`, `Kupci`, `Vozaci`, `Parcele` or `Artikli`, including table binding, visible fields and combo sources.
- **single-shell CRUD model:** the same form loads list projections, adds new rows through `AppendRow()`, edits existing rows through `UpdateCell()` and resets field state through `ClearFields()`.
- **entity-specific projections:** list rendering composes derived display values such as full kooperant names, station names, buyer address strings, and parcela geo/risk summaries rather than exposing raw row arrays directly.
- **parcel-specific maintenance:** parcela mode binds kooperant/kultura/GGAP combos, surfaces geo and rizik summary columns, and exposes geo-action hooks for parcel point maintenance flows.
- **navigation contract:** returning from the form unloads the maintenance shell and reopens `frmOtkupAPP`; the same chrome-removal/theme pattern is reused here as across other desktop forms.

The desktop splash form is the canonical startup branding handoff shell. Its active architecture is:

- **branding-only loader:** it shows `OtkupApp`, version text and `Powered by AgriX` labels with the shared dark theme and no business controls.
- **timed handoff:** `UserForm_Activate()` waits roughly 2 seconds via a `Timer` loop, then unloads itself and opens `frmOtkupAPP`.
- **startup chrome parity:** it removes the standard VBA title bar and uses the same shared label styling helpers as the rest of the desktop UI.

`frmFakturisanje` is the canonical desktop faktura assembly form. Its active architecture is:

- **kupac-first shell:** operator first selects a kupac/hlađnjača and then loads only relevant prijemnice for that buyer.
- **multi-select invoice assembly:** invoice lines are composed from selected prijemnice in a multi-select list and passed to `CreateFaktura_TX`.
- **visibility toggle:** the form can optionally show already-fakturisane prijemnice together with their faktura and uplata summary.
- **adjacent actions:** the same shell can print the latest buyer faktura and open `frmSEF` for outbound e-invoice operations.

`frmOtkup` is the canonical desktop procurement-entry form after the split that moved outbound documents into `frmDokumenta`. Its active architecture is:

- **left-side procurement shell:** the form now covers only otkup capture (`Kooperant -> Station`) while `Otpremnica`, `Zbirna` and `Prijemnica` remain in `frmDokumenta`.
- **cascading master lookups:** station selection filters kooperanti; kooperant selection loads parcel options; parcel choice can backfill fruit type/sort from parcela kultura.
- **dual-class entry:** optional `Dve klase` mode enables separate Klasa II quantity/price fields while live total kg is recomputed across class I + II.
- **optional cash sidecar:** when `Novac > 0`, procurement save also writes a kooperant cash payout row through `SaveNovac_TX` and then applies kooperant avans to saved otkup rows.

`frmSEF` is the canonical desktop shell for operational SEF management over existing fakture. Its active architecture is:

- **faktura-centric control panel:** one selected faktura drives workflow labels, external SEF status, document ID, version and last-error display.
- **event log surface:** the form loads SEF event rows into a dedicated list for operator audit/review.
- **action-state gating:** send, refresh, prepare-resubmit, cancel, storno and recovery buttons are enabled strictly from the current local workflow state plus exact external `SEFStatus`.
- **batch recovery affordances:** the shell exposes explicit refresh-pending and recover-all-stuck-sending actions on top of single-invoice commands.

`frmOtkupniBlokovi` is the canonical desktop shell for unresolved traceability repair and lot trace printing. Its active architecture is:

- **dual-purpose shell:** one form combines unresolved `Otkup -> Otpremnica` repair with reverse trace review by selected `BrojZbirne`.
- **unlinked queue view:** the first list surfaces unlinked otkup rows with resolved station, driver and kooperant names.
- **candidate otpremnica view:** clicking one unresolved row loads same-day same-station otpremnica candidates for manual link selection.
- **trace print surface:** selected zbirna renders lot trace rows and can export a formatted `Sledljivost_*.pdf` from the `SledljivostSablon` worksheet template.

### 6.1.2 `modDokumenta` Active Business Capabilities

The current canonical `modDokumenta` layer contains the following business responsibilities in addition to form-triggered save/storno entry points:

- **transactional save families:** `SaveOtpremnica_TX`, `SaveZbirna_TX` and `SavePrijemnica_TX` wrap single-class base saves with table snapshots and rollback discipline; `SaveOtpremnicaMulti_TX`, `SaveZbirnaMulti_TX` and `SavePrijemnicaMulti_TX` are the canonical dual-class wrappers and commit Klasa I/Klasa II atomically. Prijemnica TX scope explicitly protects `tblPrijemnica`, `tblAmbalaza`, `tblFakturaStavke` and `tblFakture` because intake save may repair downstream invoice linkage.
- **base save responsibilities:** `SaveOtpremnica`, `SaveZbirna` and `SavePrijemnica` allocate new IDs, append canonical rows and execute linked side-effects such as `TrackAmbalaza`; business save functions log/return failure or propagate errors but do not show UI messages.
- **read/query helpers:** `GetOtpremniceByZbirna`, `GetOtpremniceByStation`, `GetZbirnaByKupac`, `GetPrijemniceByKupac` and `GetVozacDokumenta` expose filtered document retrieval for validation, reporting and UI lookup flows.
- **validation helpers:** `ValidateZbirna` reconciles saved otpremnica vs zbirna totals, while `ValidateZbirnaPreUnosa` performs pre-save class-aware validation using current user input.
- **shortage helpers:** `CalculateManjak`, `CalculateManjakPreview` and `CalculateManjakByOtpremnica` provide total and proportional loss/shrink calculations across active zbirna/prijemnica chains.
- **crate analytics helpers:** `CalculateProsekGajbe` and `CalculateProsekGajbeByZbirna` compute average kg per ambalaža unit for transport/intake review.
- **orphan detection:** `GetVerwaisteDokumente` plus its otpremnica/prijemnica specializations detect documents whose `BrojZbirne` references only a stornirana zbirna and has not yet been replaced by a new active one.
- **invoice-link repair:** guarded `RelinkFakturaStavke` reattaches orphaned invoice lines to a replacement prijemnica with the same `BrojPrijemnice`, clears orphan markers and restores fakturisano linkage on the new intake row; each critical update uses `RequireUpdateCell` so any partial relink failure rolls back the surrounding transaction.
- **money/document side wrappers:** `SaveOMUlaz_TX` and `SaveKupciIzlaz_TX` group ambalaža, novac and status side-effects into one transaction and call base `SaveNovac` to avoid nested transactions.
- **schema hardening:** document query, validation, orphan and reporting helpers use `RequireColumnIndex`/`RequireColumns` before indexed array access.
- **v6.9 EH/input hardening:** `_TX` wrappers preserve original `Err.Number`, `Err.Source` and `Err.Description` before logging/rollback; `SaveZbirna` and `SavePrijemnica` propagate original business errors instead of masking them as empty IDs. Base writers validate required IDs/numbers, class, quantity, price and ambalaža inputs before append.
- **active-row read helpers:** `GetOtpremniceByZbirna`, `GetOtpremniceByStation`, `GetZbirnaByKupac` and `GetPrijemniceByKupac` apply `ExcludeStornirano` internally and therefore return active rows by default for normal business reads.
- **prijemnica returned ambalaža naming:** the canonical returned-packaging column is `tblPrijemnica.kolAmbVracena`; tests and writers must not use the non-canonical name `KolAmbalazeVracena`.
- **report/cache helpers:** `BuildZbirnaVrstaCache` and `GetVrstaFromCache` provide lightweight `BrojZbirne -> VrstaVoca` lookup caching for transport/report contexts.

### 6.1.2a Desktop Dokumenta Hardening Snapshot

The current canonical desktop document layer applies these implementation rules:

- **Form responsibility:** `frmDokumenta` performs UI validation, shared parsing, duplicate pre-checks, hidden-ID selection and user feedback. It does not own transaction orchestration internals.
- **Business responsibility:** `modDokumenta`, `modNovac` and related modules own writes, snapshots, rollback, status refresh and relink logic. They do not display `MsgBox`.
- **Dual-class documents:** `SaveOtpremnicaMulti_TX`, `SaveZbirnaMulti_TX` and `SavePrijemnicaMulti_TX` are the active dual-class save entry points; Klasa II rows carry zero ambalaža side-effects.
- **OM ulaz:** the active architecture requires one wrapper transaction covering `tblAmbalaza`, `tblNovac` and `tblOtkup` status effects.
- **Izlaz kupci:** the active architecture requires one wrapper transaction covering `tblAmbalaza`, `tblNovac` and `tblFakture` status effects.
- **Faktura combo binding:** open-fakture lists display invoice context but carry hidden `FakturaID`; payment save uses the hidden ID.
- **Prijemnica relink:** replacement prijemnica save calls guarded `RelinkFakturaStavke`; if any critical update fails, the surrounding TX rolls back.
- **Orphan warnings:** `GetVerwaisteDokumente` defines orphan semantics and feeds warning surfaces in the desktop shells.
- **Compile gate:** any change touching these flows must pass `Debug > Compile VBAProject` and regression tests for otpremnica, zbirna, prijemnica, OM ulaz, kupci izlaz, storno and relink. v6.9 also requires the expanded `RunBusinessFlowProSuite` document/otkup hardening tests to pass.

### 6.1.2b `modFaktura` Active Business Capabilities

The current canonical `modFaktura` layer contains the following invoice responsibilities:

- **transactional invoice create:** `CreateFaktura_TX()` snapshots `tblFakture`, `tblFakturaStavke`, `tblPrijemnica` and `tblNovac`, delegates to `CreateFaktura()` and rolls back header, lines, prijemnica linkage and buyer-avans effects on failure.
- **canonical prijemnica values:** `CreateFaktura()` trusts caller-provided stavke only for `PrijemnicaID`. It reads `Kolicina`, `Cena`, `Klasa` and `BrojPrijemnice` from `tblPrijemnica`, preventing UI/caller payload tampering from changing invoice amounts or line metadata.
- **pre-create guards:** duplicate selected prijemnica IDs, stornirana prijemnica rows, already-fakturisane prijemnice and prijemnice already carrying `FakturaID` are fail-fast blocks before any write.
- **status recompute:** `UpdateFakturaStatus()` recomputes status in both directions from active linked uplata. It preserves existing `DatumPlacanja` when already populated, fills it only on first closure and clears it when the faktura reopens. Stornirana faktura rows are skipped.
- **print guard:** `PrintFaktura()` blocks active printing of stornirana faktura rows. Any archival reprint of stornirana invoices must be a separate, clearly marked workflow.
- **test coverage:** `RunFakturaSmokeSuite` covers canonical prijemnica value usage, duplicate/stornirano/already-fakturisano blocks, status recompute/date preservation and print block for stornirana faktura.

### 6.1.3 `modNovac` Active Business Capabilities

The current canonical `modNovac` layer contains the following active financial capabilities:

- **validated novac writes:** `SaveNovac()` validates money direction and basic partner/entity context before appending. Valid rows have exactly one positive amount direction (`Uplata` or `Isplata`), no negative amounts and a non-empty `Tip`.
- **transactional novac writes:** `SaveNovac_TX()` snapshots `tblNovac`, `tblFakture` and `tblOtkup`, then delegates to `SaveNovac()` for append-oriented row creation. The active transaction model is snapshot/rollback, not pending-write; direct `AppendRow` writes are rollback-safe only when the affected table was snapshotted before the write.
- **partner-map learning:** `LookupPartnerMap()` and `savePartnerMap()` define learned exact-match mapping between bank partner names and internal entity identifiers. An identical existing mapping is idempotent success, while a conflicting PartnerID/EntitetTip/OMID for the same bank name is fail-fast.
- **open-item resolution:** `GetOpenFakture()` and `GetOpenOtkupi()` define open receivable/payable sets by outstanding remainder, not by status field alone. Stornirano money/faktura/otkup rows are excluded from live aggregates.
- **status refresh:** `UpdateOtkupStatus()` recomputes paid status in both directions: linked `Isplata >= Kolicina × Cena` sets `Isplaceno` and fills `DatumIsplate` if empty; insufficient linked payment clears both fields. Faktura closure relies on uplata aggregation and `UpdateFakturaStatus()`.
- **advance allocation:** `ApplyAvansToFaktura[_TX]()` and `ApplyAvansToOtkup[_TX]()` are canonical allocation flows. Full consumption links the existing avans row; partial consumption reduces the original row and creates a linked split row. Required updates use `RequireUpdateCell`, split creation checks the returned `NOV-*` ID, and wrapper EH preserves the original error before rollback/logging.
- **otkup link repair:** `ResetNovacOtkupLink_TX()` snapshots `tblNovac` and `tblOtkup`, clears active `OtkupID` money links through fail-fast updates, then calls `UpdateOtkupStatus()` so unlink and status recompute commit or rollback together.
- **OM station saldo:** `GetOMAvansSaldo()` defines OM avans as cash sent from firm to station minus cash paid from station to kooperant, using guarded required-column lookup and stornirano exclusion.
- **aggregation helpers:** `GetBankaByPartner()`, `GetUplataByVrsta()`, `BuildUplataDictByFaktura()`, `BuildIsplataDictByOtkup()`, `GetUplataForOtkup()` / `GetIsplataForOtkup()` and `GetAgroAbzug()` are active reporting helpers for finance and saldo views. Required columns are guarded with `RequireColumnIndex`.
- **test coverage:** `modNovacTests.RunNovacSmokeSuite` covers validation rejection, valid uplata/isplata, stornirano exclusion, partner-map conflict, partial buyer avans split, partial otkup avans split, and reset-link status recompute.

### 6.1.4 `modAmbalaza` Active Business Capabilities

The current canonical `modAmbalaza v2.1` layer defines the packaging-tracking and saldo read model used by document and payment flows:

- **canonical write helper:** `TrackAmbalaza` is the standard writer for `tblAmbalaza`; it allocates `AMB-*` IDs and appends rows with `Datum`, `TipAmbalaze`, `Kolicina`, `Smer`, `EntitetID`, `EntitetTip`, and optional `VozacID`, `DokumentID`, `DokumentTip` lineage.
- **entity balance helper:** `GetAmbalazeStanje(entitetID, entitetTip)` returns net packaging position per type for a business entity after excluding stornirano rows; `Ulaz` increments saldo and `Izlaz` decrements it.
- **driver saldo helper:** `GetVozacAmbSaldo(VozacID, datumOd, datumDo)` returns per-driver packaging exposure grouped by `TipAmbalaze` as `(Izlaz, Ulaz, Saldo)` with optional date filtering.
- **document integration:** document and finance flows call `TrackAmbalaza` as a side effect of otpremnica, prijemnica, OM ulaz and kupac-return operations, which means packaging is modeled as a journal and not as a directly overwritten stock field.
- **storno safety:** all read helpers first apply `ExcludeStornirano(TBL_AMBALAZA)` so storno rows do not distort live saldo reporting.

### 6.1.5 `modMagacin` Active Business Capabilities

The current canonical desktop agro/warehouse layer exposes the following active capabilities:

- **kooperant parcela lookup:** `GetParceleByKooperant()` returns kooperant-scoped parcela tuples together with a formatted display label for downstream operator selection flows.
- **dosage recommendation:** `CalculatePreporuka(artikalID, povrsinaHa)` computes recommendation strictly from `tblArtikli.DozaPoHa × PovrsinaHa`; non-numeric or missing dosage resolves to zero instead of throwing.
- **transactional warehouse save:** `SaveMagacin_TX()` snapshots `tblMagacin` and delegates to `SaveMagacin()` for append-oriented `MAG-*` row creation.
- **write-time valuation:** `SaveMagacin()` reads current article price from `tblArtikli`, persists `CenaPoJedinici`, and stores row `Vrednost = Kolicina × CenaPoJedinici` as part of the warehouse journal.
- **stock read model:** `GetMagacinStanje()` defines live warehouse state per article as `(Ulaz, Izlaz, Stanje)` after excluding stornirano rows and enriches results with article name, type and unit-of-measure lookups.
- **report helpers:** `ReportIzdavanjePoKooperantu()` aggregates outbound warehouse value per kooperant with optional date filtering and total row, while `ReportStanjePoDoabvljacu()` aggregates inbound value by supplier plus total.
- **debt helper:** `GetAgrohemijaDug()` defines kooperant agro debt as the sum of outbound (`MAG_IZLAZ`) warehouse row values linked to that kooperant after storno exclusion.

### 6.1.6 `modBankaImport` Active Business Capabilities

The current canonical desktop bank-import orchestration layer exposes the following active capabilities:

- **transactional inbox import wrapper:** `ImportBankaInbox_TX()` ensures configured inbox/processed/error folders exist, snapshots `tblBankaImport`, and wraps the full inbox import in rollback discipline.
- **batch inbox scan:** `ImportBankaInbox()` enumerates all `*.pdf` files from `APP_BANKA_INBOX` and processes them one by one through the per-file import routine.
- **per-file triage:** `ImportOnePdfIntoBankaImport()` delegates PDF text extraction and parse preparation, stages import rows into `tblBankaImport`, and then moves the source file either to `APP_BANKA_PROCESSED` on success or `APP_BANKA_ERROR` on empty/unparseable/error outcomes.
- **header extraction contract:** `ParseBankaIzvodForImport()` requires successfully extracted `BrojIzvoda`, `DatumIzvoda` and `BrojRacuna`; missing any of these is treated as a hard parse failure rather than a partially staged import.
- **staging row shape:** parsed transaction rows are normalized into the import schema with statement-level metadata copied onto each row and transaction-level values for date, partner, konto, amounts, payment purpose, reference and source filename.
- **append-time defaults:** `SaveBankaImportRows()` allocates new `BIM-*` IDs, persists `Valuta = "RSD"`, stamps `ImportVreme = Now`, and initializes `Obradjeno` / `Stornirano` as blank until later processing changes state.
- **idempotent duplicate check:** `IsDuplicateBankaImport()` first uses `BankaReferenz` when present; if no bank reference exists, it falls back to the composite duplicate key `(BrojDokumenta, DatumTransakcije, Uplata, Isplata, Partner)` within active staging rows.
- **safe file move helper:** `GetUniqueTargetPath()` prevents overwrite collisions in processed/error folders by suffixing duplicate filenames with `_NNN` before move.

### 6.1.7 `modBankaImport_PdfText` Active Business Capabilities

The current canonical desktop PDF-text parser layer exposes the following active capabilities:

- **local PDF-to-text dependency:** `ExtractTextFromPdf()` shells out to a local Poppler `pdftotext.exe` and requests `-raw -nopgbrk -enc UTF-8` output before the parser touches bank statement content.
- **statement header extraction:** `ExtractIzvodBrojPdfText()`, `ExtractIzvodDatumPdfText()` and `ExtractIzvodRacunPdfText()` define the mandatory header fields required upstream by `ParseBankaIzvodForImport()`.
- **transaction-block segmentation:** `CollectPdfTextTxnBlocks()` groups line ranges into candidate transactions using numeric sequence markers as starts and statement-summary markers (`Ukupno...`, fee totals, next statement headers) as hard stops.
- **strict block parsing:** `ParsePdfTextTxnBlock()` extracts execution date, partner, account, zaduženje, odobrenje, šifra, svrha, `PozivNaBroj` and `Referenca` from each normalized block and returns a fixed 10-column in-memory row.
- **amount/code parsing rule:** `ParsePdfOdobrenjeSifraLineStrict()` treats the odobrenje+šifra line as the canonical parse anchor and strips trailing summary text hanging on the same line before downstream cleanup.
- **cleanup and recovery helpers:** `ExtractReferenceFromSvrhaPdf()`, `ExtractPozivNaBrojPdf()` and `CleanSvrhaPdf()` recover embedded references from purpose text, remove dangling `[97]`-style tokens, and trim trailing dates or summary noise.
- **text-shape invariants:** `NormalizeSpacesPdf()`, `IsDateLinePdf()`, `IsAccountLinePdf()`, `IsAmountPdf()` and related helpers define the expected normalized line grammar for the Komercijalna Banka parser and are part of the active parse contract.
- **operator test hooks:** `PickPdf()`, `TestPdfTextParser()` and `TestPdfTextParser123()` remain active diagnostic helpers for validating parser behavior against real statement files during desktop maintenance.

### 6.1.8 `modBankaMapiranje` Active Business Capabilities

The current canonical desktop bank-reconciliation layer exposes the following active capabilities:

- **open staging view:** `GetBankaImportOpen()` returns active bank-import rows whose `Obradjeno` is neither `Da` nor `Skip`; stornirano rows are excluded before operator/manual or auto-map workflows begin.
- **processing gate:** `ValidateBankaImportNotProcessed()` is the canonical guard for all mapping actions; already processed, skipped or stornirano rows are rejected before any financial write is attempted.
- **status lifecycle:** `tblBankaImport.Obradjeno` uses the active values `""`, `Da`, `Skip` and `Error`; successful mapping marks `Da`, explicit operator deferral marks `Skip`, and unresolved/failed mapping paths mark `Error`.
- **transactional mapping wrappers:** all public mutating bank-map flows have `_TX` wrappers that snapshot at least `tblBankaImport` and `tblNovac`, while buyer/kooperant flows additionally protect `tblFakture`, `tblOtkup` and/or `tblPartnerMap` depending on downstream side effects.
- **manual mapping families:** `MapBankaImportAsKupac[_TX]()`, `MapBankaImportAsKooperant[_TX]()`, `MapBankaImportAsOM[_TX]()` and kooperant-block variants are the canonical operator-assisted reconciliation paths from staging rows into `tblNovac`.
- **auto-map orchestration:** `AutoMapBankaImportRow[_TX]()` routes clean-direction rows by amount polarity (`uplata` vs `isplata`), while `AutoMapAllBankaImport_TX()` replays that logic across the current open staging queue in one transaction scope.
- **entity resolution order:** auto-map first checks `tblPartnerMap` learned mappings, then falls back to normalized-name resolution against `tblKupci`, `tblKooperanti` and `tblStanice`; unresolved rows are left in staging and marked `Error`.
- **buyer payment semantics:** incoming bank rows mapped to buyers create `NOV_KUPCI_UPLATA` when a unique faktura is resolved and otherwise `NOV_KUPCI_AVANS`; successful faktura-linked mappings also refresh invoice payment status.
- **kooperant bank payout semantics:** outgoing bank rows for kooperants resolve to OM-linked `tblNovac` rows using the kooperant's `StanicaID`; direct otkup-linked payouts use `NOV_VIRMAN_FIRMA_KOOP`, while unmatched remainder or advance-only flows use `NOV_VIRMAN_AVANS_KOOP`.
- **OM funding semantics:** incoming rows resolved as station/OM funding create `NOV_KES_FIRMA_OTKUPAC` entries with `Partner = Naziv stanice`, `PartnerID = OMID`, `EntitetTip = "OM"` and `OMID = StanicaID`.
- **faktura matching heuristic:** `TryResolveFakturaForKupac()` attempts a unique hit by normalized `PozivNaBroj`, then by invoice number found inside `SvrhaPlacanja`, and finally by exact amount match; only a single unambiguous hit becomes authoritative.
- **kooperant block allocation:** `MapBankaImportAsKooperantBlock*()` uses `PozivNaBroj`/manual block number to find up to two open otkup candidates for the kooperant, sorts them by larger open amount first, writes one or more `tblNovac` rows, links consumed rows to `OtkupID`, updates otkup paid status, and pushes any remaining excess into kooperant avans.
- **mapping traceability:** every reconciled `tblNovac` row stores a generated napomena containing `BIM:<id>` plus selected bank reference, konto, opis, svrha and match reason metadata so later saldo and audit review can trace the source staging line.
- **learned mapping persistence:** successful manual/auto mappings can persist `savePartnerMap()` entries so future statements from the same bank-side partner name resolve without repeated operator selection.
- **operator defer path:** `SkipBankaImportRow[_TX]()` exists as a first-class workflow for intentionally postponing a staging row without deleting or storniranje the imported source record.



### 6.1.9 `modOtkup` Active Business Capabilities

The current canonical desktop procurement core exposes the following active capabilities:

- **transactional procurement save:** `SaveOtkup_TX()` snapshots `tblOtkup` and `tblAmbalaza`, delegates to `SaveOtkup()`, preserves original errors in EH, and rolls back the procurement write plus packaging side-effects on failure.
- **base otkup creation:** `SaveOtkup()` allocates a new `OTK-*` ID, validates required inputs and class, resolves `KulturaID` from `tblKulture` when possible, writes optional `ParcelaID` and `BrojZbirne`, propagates original business errors, and persists the row as the canonical start of the desktop-side procurement chain.
- **packaging side effect:** when `KolAmbalaze > 0`, procurement save writes `TrackAmbalaza(..., "Izlaz", kooperantID, "Kooperant", , newID, DOK_TIP_OTKUP)` so supplied ambalaža leaves kooperant-side balance as part of the same business action.
- **read/query helpers:** `GetOtkupByStation()` and `GetOtkupByKooperant()` apply `ExcludeStornirano` internally and provide active-row filtered array views over procurement rows with optional date-window filtering through `FilterArray()`.
- **station saldo helper:** `GetSaldoByStation()` aggregates kg, explicit novac and ambalaža by kooperant for one station and time window; the code explicitly marks bank/isporuka subtraction as unfinished TODO and therefore this helper is only a partial saldo model in the current desktop snapshot.
- **validation gate:** procurement save requires selected kooperant, station, fruit type, positive quantity, positive price, non-negative money/ambalaža, valid class and ambalaža type whenever ambalaža quantity exists before any row is written.

### 6.1.10 `modSledljivost` Active Business Capabilities

The current canonical desktop traceability layer exposes the following active capabilities:

- **auto-link repair:** `AutoLinkOtkupOtpremnica()` scans active otkupi whose `OtpremnicaID` is blank, builds an otpremnica dictionary by `StanicaID|Datum|VozacID|Klasa`, and writes the link only when there is exactly one candidate for that key.
- **manual unresolved queue:** `GetUnlinkedOtkupi()` is the canonical read model for procurement rows still lacking `OtpremnicaID`, returning compact fields needed for manual review.
- **reverse shipment trace:** `TraceByZbirna(brojZbirne)` walks active `Zbirna -> Otpremnica -> Otkup` relations and enriches each traced procurement row with kooperant name, BPG, parcela metadata, class and shipment context.
- **parcel enrichment contract:** reverse-trace includes parcela-derived katastarski broj, GGAP status, kultura and površina when `ParcelaID` exists on the otkup row.
- **ambiguity rule:** multiple matching otpremnice for the same auto-link key are intentionally left unresolved for manual operator handling rather than guessed automatically.

### 6.1.11 `modLogError` Active Business Capabilities

The current canonical desktop runtime logging layer exposes the following active capabilities:

- **daily log file contract:** log lines are written under `ThisWorkbook.Path\Log\` into files named `OtkupApp_YYYY-MM-DD.log`.
- **non-blocking write rule:** `LogError()` uses best-effort file append with `On Error Resume Next`; logging failure must never block the business flow that is being logged.
- **leveled logging:** `LOG_ERROR`, `LOG_WARN` and `LOG_INFO` are active log levels, with convenience wrappers `LogErr()`, `LogWarn()` and `LogInfo()`.
- **lifecycle hooks:** `LogAppStart()` and `LogAppShutdown()` provide app-session boundary markers and include workbook/user context in the startup log.
- **retention policy:** `PurgeOldLogs()` deletes log files older than 30 days based on the date embedded in the filename, not on arbitrary filesystem metadata.

### 6.1.12 `modJournal` Active Business Capabilities

The current canonical desktop journal/recovery layer exposes the following active capabilities:

- **append journaling hook:** `WriteJournalRow()` is called from `AppendRow()` and writes every successful append as a semicolon-separated CSV line into a per-table daily file under `Journal/`.
- **header contract:** a new journal file begins with `JournalTime` plus the live table headers returned by `GetTableHeaders(tblName)`, so row replay keeps schema context.
- **startup backup:** `BackupFileOnStart()` creates a timestamped workbook copy in `Backup/` once per start slot and logs successful backup creation through `LogInfo()`.
- **recovery warning:** `CheckJournalForRecovery()` compares today's journal line count (minus header) with the current Excel table row count and emits advisory warnings when journal rows exceed table rows, indicating potential post-crash data loss.
- **retention policy:** `PurgeOldJournals()` and `PurgeOldBackups()` prune files older than 30 days using dates encoded in filenames.
- **non-blocking maintenance:** journaling and purge helpers are explicitly best-effort and must not stop normal workbook use.

### 6.1.13 `modArrayUtils` Active Business Capabilities

The current canonical desktop in-memory array layer exposes the following active capabilities:

- **multi-filter selection:** `FilterArray()` applies a collection of `clsFilterParam` objects over 2D 1-based arrays and returns a compact filtered array or `Empty`.
- **filter operator grammar:** `MatchesFilter()` supports active operators `=`, `>=`, `<=`, `BETWEEN`, `<>` and `LIKE`, with numeric/date-aware comparisons where relevant.
- **in-memory sorting:** `SortArray()` plus `QuickSortIndex()` sort 2D arrays without touching worksheets and optionally support a secondary sort column.
- **aggregation helpers:** `SumColumn()` and `GroupBySum()` provide memory-side summarization for report-like use cases and are the approved replacement for ad hoc worksheet grouping.
- **sheet-bypass rule:** this module exists specifically to replace historical Copy/Paste/Sort behavior on worksheets with deterministic in-memory operations.

#### 6.1.14 Data Access and Desktop Helper Invariants

The desktop helper layer now includes three explicit shared helper modules beyond generic table access:

- **`modParse`:** the single canonical parse/normalization surface for user-entered numbers, integers and dates. It accepts Serbian/European decimal conventions and prevents local private parser drift across forms.
- **`modSchemaGuard`:** fail-fast data guard module containing `RequireColumnIndex`, `RequireColumns` and `RequireUpdateCell`. It converts missing schema columns and failed critical updates into explicit errors that are visible in logs and rollback paths.
- **`modComboBinding`:** shared MSForms ComboBox binding surface. Canonical entity combos display names to the operator but store stable IDs in hidden column 1; saves read `GetComboID()` rather than resolving IDs from human labels.


- `GetTable()` is the canonical workbook-wide resolver for named `ListObject` tables, and business modules are expected to work on arrays returned by `GetTableData()` rather than direct worksheet manipulation.
- `AppendRow()` is both the canonical row-append helper and the journaling hook, because every successful append also calls `WriteJournalRow(...)`.
- `CheckDuplicate()` is the standard desktop duplicate guard for user-facing document numbers and explicitly ignores rows already marked `Stornirano="Da"`.
- `ExcludeStornirano()` is generic and returns the source array unchanged when a table has no `Stornirano` column, which makes the helper safe on master tables like `tblStanice` and `tblVozaci`.
- Desktop display helpers must parse IDs from both `"ID - Name"` and `"Name (ID)"` strings, because both forms are active in the workbook UI.
- `CheckVerwaisteDokumente()` is the canonical warning builder for missing document links and active documents still pointing to a stornirana zbirna.
- `FilterArray()`, `SortArray()`, `GroupBySum()` and `SumColumn()` are the approved in-memory report utilities and replace worksheet-level copy/sort/group patterns.

#### 6.1.14a SEF Module Contracts and v6.6 Tested Hardening

The active SEF subsystem is a six-module desktop subsystem:

| Module | Canonical responsibility | v6.5 hardening requirement |
|---|---|---|
| `modSEFMapper` | Build `clsSEFInvoiceSnapshot`, serialize UBL XML and compute lightweight payload fingerprint | use `GetDefaultTaxPercent`, guarded config parsing, DTO-owned `InvoiceDate`/`DeliveryDate`, fail-fast `DeliveryDate <= InvoiceDate` validation, and serializer date output from DTO only |
| `modSEFValidator` | Gate send/cancel/storno/resubmit actions | preserve state machine rules; use schema/parser guards; validate SEF config before HTTP send |
| `modSEFPersistance` | Read/write faktura SEF fields, submissions and event log | use SEF schema helpers and `RequireUpdateCell`; no local `RaiseUpdateError` as canonical API |
| `modSEFClient` | Perform WinHTTP calls to SEF API | centralized config/headers/timeouts; debug behind `SEF_DEBUG_LOG`; 429 -> `RATE_LIMITED`; status endpoint uses `ParseStatusResponse` |
| `modSEFService` | Orchestrate send/cancel/storno/retry/recovery | keep HTTP outside local TX; preserve original errors before rollback; guarded stuck-send recovery |
| `modSEFStatusSync` | Refresh external status and batch refresh pending invoices | idempotent refresh; final-state protection; guarded per-invoice batch logging |

`modSEFPersistance` spelling remains intentionally preserved in v6.6 for compile/reference compatibility. A later controlled rename may standardize it to `Persistence`, but only if all references are migrated together.

`ComputePayloadHash` remains a lightweight fingerprint, not a cryptographic hash. It is sufficient for current payload identity tracking but must not be presented as audit-grade SHA-256 evidence.

`modSEFClient` JSON extraction remains lightweight and key-based. It is acceptable for simple SEF response fields, while nested/escaped JSON support remains an open cleanup item.


### 6.1.25a Dev/Test Module Baseline

The v6.6 desktop workbook may carry dev-only test modules during hardening. These modules are not operator features and should not be exposed in normal UI navigation.

| Module | Purpose | Canonical status |
|---|---|---|
| `modBusinessFlowProTests` | Seeds an empty workbook, runs Otkup → Otpremnica → Zbirna → Prijemnica → Faktura, validates duplicate/invalid-save behavior, checks trace output and audits cross-zbirna links | Dev/test-only regression suite |
| `modSEFTests` | Runs offline SEF mapper/validator tests, state-transition matrix tests, live submit/refresh smoke, repeated refresh idempotency, recovery smoke and destructive cancel/storno tests under explicit flags; successful outbound tests fail if `SEFDocumentId` is missing | Dev/test-only SEF integration suite |
| `modDevReset` / equivalent | Clears all ListObject data in a test workbook while preserving table structure and headers | Dev-only utility, never exposed to operators |

`modBusinessFlowProTests` is the canonical regression evidence for the strict BrojZbirne trace bridge. Its negative scenario must keep an otkup row with `BrojZbirne=A` unlinked when the only candidate otpremnica belongs to `BrojZbirne=B`, while the matching `BrojZbirne=B` otkup links correctly.

`modSEFTests` records results in `SEF_TEST_LOG` and treats successful SEF business rejections as plumbing PASS when the rejection is persisted with `ErrorCode`, `ErrorMessage`, submission row and event log. Local UBL validation failures such as `DeliveryDate > InvoiceDate` are expected SKIP/PASS evidence for validation hardening, not transport failures. The suite includes an offline state-transition matrix and treats missing `SEFDocumentId` after successful outbound workflow states as FAIL.

Destructive SEF cancel/storno tests must be explicitly gated by `SEF_TEST_ALLOW_LIVE = DA` and `SEF_TEST_ALLOW_CANCEL_STORNO = DA`, plus user confirmation. They must not be run on production invoices unless that is the intended test scenario. v6.7 test semantics require `CancelInvoiceOnSEF_TX` / `StornoInvoiceOnSEF_TX` to return `True` before a destructive smoke can pass, classify already-`STORNO` invoices as SKIP, and separate API-call success from final external outcome verification.

### 6.1.26 `frmAgrohemija` Active Desktop Form Capabilities

The current canonical desktop agrohemija form exposes the following active capabilities:

- **kooperant-driven parcel loading:** selecting a kooperant loads parcel display rows from `GetParceleByKooperant()` and caches hidden `ParcelaID` plus hectare arrays for later recommendation and save logic.
- **debt awareness:** kooperant change also computes the visible debt label as `GetAgrohemijaDug(koopID) - GetAgroAbzug(koopID)`.
- **recommendation engine:** `UpdatePreporuka()` sums all selected parcel hectares, calls `CalculatePreporuka(artikalID, totalHa)`, and when `Pakovanje` exists rounds the recommended amount up to whole packages before prefilling the issue quantity.
- **value preview:** both izlaz and ulaz paths recompute RSD value labels from current quantity × price inputs before save.
- **issuance basket model:** each issue line is staged inside `m_KorpaIzlaz()` with article, quantity, price, value, joined parcela IDs and unit of measure; save persists one `MAG_IZLAZ` row per basket line under one transaction.
- **receipt basket model:** inbound stock lines are staged separately inside `m_KorpaUlaz()` and persisted as `MAG_ULAZ` rows with optional supplier linkage under one transaction.
- **package-multiple validation:** issue save blocks quantities that are not whole multiples of configured packaging size.
- **document-number gate:** both ulaz and izlaz commit paths require explicit `BrojDok` before any warehouse movement is saved.
- **form-hardening status:** recent pre-launch cleanup aligns Agrohemija/Stammdaten forms with the shared parser/helper direction: local parser duplication is not canonical, missing helper functions are completed at form level where they are truly UI-specific, and add/edit/load handlers are expected to compile against shared helpers.

### 6.1.27 `frmBankaImport` Active Desktop Form Capabilities

The current canonical desktop bank-import review form exposes the following active capabilities:

- **open-queue loader:** `LoadBankaRows()` reads `GetBankaImportOpen()` and shows only non-stornirano rows that are not already processed or skipped.
- **direction-aware defaults:** selecting a row auto-suggests `Kupac` mapping for uplata-only rows and `Kooperant` mapping for isplata-only rows.
- **manual target loaders:** `LoadManualTargets()` dynamically fills target combos from kupci, kooperanti or stanice depending on selected `MapTip`.
- **otkup-block assist:** when mapping a kooperant, the form loads distinct open `BrojDok` block candidates from active `tblOtkup` rows into a dedicated combo for manual override.
- **preview parity with backend:** preview UI calls the same `modBankaMapiranje` helpers (`GetBankaImportRowByID`, `TryResolveKupacBIM`, `TryResolveKooperantBIM`, `TryResolveOMBIM`, `TryResolveFakturaForKupac`, `GetOtkupCandidatesForKooperantBlock`, `NormalizeLooseBIM`, `NzBIM`, `GetKooperantNaziv`) directly; preview-parity is now a structural property, not a duplicated copy of mapping logic.
- **auto-map controls:** the form can run `AutoMapBankaImportRow_TX()` for the selected row or `AutoMapAllBankaImport_TX()` for the entire open queue.
- **manual commit controls:** the form can explicitly save kupac, kooperant/block or OM mappings through the corresponding `_TX` wrappers, and can also mark rows as `Skip`.
- **review-only shell:** the form itself does not implement financial business rules; it orchestrates existing mapping modules and then reloads the queue.



### 6.1.28 `frmOtkupAPP` Active Desktop Form Capabilities

The current canonical desktop home shell exposes the following active capabilities:

- **fullscreen shell resize:** `ResizeMainForm()` binds the form to the Excel application viewport and `SetupShellResponsive()` lays out header, sidebar and content cards from live `InsideWidth/InsideHeight`.
- **alert-card warnings:** activation calls `CheckVerwaisteDokumente()` and projects the resulting warning text into `lblStatus` with alert coloring instead of burying orphaned-document issues.
- **sidebar route map:** the shell launches otkup, dokumenta, agrohemija, izveštaj, fakturisanje, marža, sledljivost, Excel mini-shell and snapshot/exit affordances from one navigation container.
- **Banka mapping action:** the Banka navigation button is the canonical operator-triggered point for `ImportBankaInbox_TX()` before opening the bank-mapping/review workflow; bank inbox import is intentionally not run from `Workbook_Open`.
- **modeless master-data popup:** `btnMaticni_Click()` loads `frmMaticniPodaci` modeless and positions it relative to the header button rather than opening it as a full modal workflow.
- **Excel visibility control:** `btnOpenExcel` intentionally restores Excel and opens `frmExcelMini` modeless, while `btnExit` saves the workbook and quits the application.
- **window-close discipline:** form-control close is trapped and delegated to `ShutdownApp`, keeping shell shutdown behavior centralized; the workbook-level `Workbook_BeforeClose` handler also routes through `ShutdownApp`, so the canonical shutdown path is invoked regardless of whether the operator closes the form or the workbook.

### 6.1.29 `frmStammdaten` Active Desktop Form Capabilities

The current canonical desktop master-data maintenance form exposes the following active capabilities:

- **section-by-tag setup:** on first activate the form switches between `Kooperanti`, `Stanice`, `Kupci`, `Vozaci`, `Parcele` and `Artikli` using `frm.Tag`, with entity-specific table names, headers, visible fields and combo data sources.
- **derived list projections:** `LoadList()` builds operator-friendly list rows such as `Ime + Prezime`, resolved station names, composed buyer addresses, and parcela `Geo/Rizik` summaries instead of showing raw table schema directly.
- **add workflow:** `btnDodaj_Click()` validates entity-required fields, generates the proper prefixed ID (`KOOP-`, `ST-`, `KUP-`, `VOZ-`, `PAR-`, `ART-`) and appends a new row with active/default flags.
- **edit workflow:** `btnIzmeni_Click()` updates only the relevant columns for the chosen entity through `UpdateCell()`, keeping business persistence inside `modDataAccess`.
- **parcel maintenance specialization:** parcela mode binds kooperant/kultura/GGAP combos, stores `Aktivna="Da"` by default on add, and surfaces parcel geo/risk columns plus geo-action hooks for further parcel geo entry.
- **navigation/reset helpers:** `ClearFields()` resets all text/combo state and return actions unload the form and reopen `frmOtkupAPP`.

### 6.1.29a App Startup Orchestration

The desktop app startup pipeline is centralized in `modMain.StartApp` and is invoked from `Workbook_Open`. Its active architecture is:

- **single entry point:** `Workbook_Open` does no business work itself; it wraps `StartApp` in an EH block and is responsible only for guaranteeing that `Application.Visible` is restored to `True` before any user-facing error message, so the operator can never end up locked out of an invisible Excel.
- **one-time initialization:** `StartApp` calls `InitApp` on the first run, which suspends `ScreenUpdating`, `Calculation` and `EnableEvents`, runs `ValidateAllTables` against the canonical table list (`tblKooperanti`, `tblStanice`, `tblVozaci`, `tblKupci`, `tblKulture`, `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`, `tblFakturaStavke`, `tblNovac`, `tblAmbalaza`, `tblConfig`) and surfaces a missing-tables warning if any are absent.
- **splash handoff:** after `InitApp` succeeds, `StartApp` sets `Application.Visible = False`, shows `frmSplash` (which auto-unloads after ~2s and opens `frmOtkupAPP`), and then proceeds with daily housekeeping in parallel with the splash timer.
- **daily housekeeping:** `BackupFileOnStart`, `PurgeOldBackups`, `PurgeOldJournals` and `PurgeOldLogs` run on every boot; the daily lifecycle marker is written via `LogAppStart`, which produces the first line in the daily log file under `Log/`.
- **no import-on-open policy:** file imports such as `ImportBankaInbox_TX()` are not startup work; they are triggered from explicit operator actions after the shell is available, so boot remains predictable and recoverable.
- **opportunistic SEF recovery:** `RecoverAllStuckSEFSendingInvoices` is called inside `On Error Resume Next` so a recovery failure does not block boot; this clears any fakture left in `WF_SEF_SENDING` from a previous crash and either refreshes their status from SEF or moves them to `WF_SEF_TECH_FAILED` for retry.
- **journal recovery surface:** `CheckJournalForRecovery` checks for unfinished master-sync sessions from a previous run and, if any are found, surfaces an explicit "moguć gubitak podataka" warning so the operator can manually inspect and reimport.
- **shutdown symmetry:** both form-control close on `frmOtkupAPP` and workbook-level `Workbook_BeforeClose` route through `ShutdownApp`, which writes the closing `LogAppShutdown` marker and unloads remaining shells; this guarantees the daily log always has a paired open/close lifecycle pair, regardless of how the operator exits.

`StartApp` is the canonical desktop boot contract; ad-hoc startup work (auto-imports, batch jobs, side-effect routines) must not be added to `Workbook_Open` directly — they belong either inside `StartApp`'s ordered pipeline or behind explicit operator action in `frmOtkupAPP` or its child forms.

### 6.1.30 Desktop Splash Form Capabilities

The current canonical desktop splash form exposes the following active capabilities:

- **branding labels:** startup text composes `OtkupApp`, the version label derived from the `APP_VERSION` constant in `modConfig` (currently `v2.2.1`) and `Powered by AgriX`; the version is no longer hardcoded in the form.
- **shared-theme usage:** activation applies `BG_MAIN()` and shared label styling helpers before showing the handoff screen.
- **fixed-timer handoff:** the form keeps itself visible for roughly two seconds using a `Timer`/`DoEvents` loop, then unloads and opens `frmOtkupAPP`.
- **startup-only responsibility:** the splash shell performs no data reads or writes beyond UI presentation and main-shell handoff.


### 6.1.31 Desktop Stammdaten Launcher Form Capabilities

The current canonical desktop stammdaten launcher form exposes the following active capabilities:

- **section dispatch:** dedicated buttons open master-data maintenance for `Kooperanti`, `Stanice`, `Kupci`, `Vozaci`, `Artikli` and `Parcele`.
- **tag-based child routing:** `OpenStammdatenForm(nazivSekcije)` passes the chosen section through `frmStammdaten.Tag`, which is the canonical routing contract used by the target maintenance form.
- **parent-shell hiding:** when opening a child form the launcher hides both itself and `frmOtkupAPP`, preventing parallel menu shells from remaining visible.
- **self-close invariant:** if the form deactivates without intentionally opening a child, it unloads itself to keep only one active menu shell.
- **desktop chrome parity:** the launcher uses the same title-bar removal and themed hover/active/exit button styling pattern as other operator forms.

### 6.1.32 `frmFakturisanje` Active Desktop Form Capabilities

The current canonical desktop fakturisanje form exposes the following active capabilities:

- **kupac-driven load:** `btnUnesi_Click()` resolves `KupacID`, loads prijemnice through `GetPrijemniceByKupac()`, filters stornirano rows, and projects them into a multi-select list.
- **fakturisano visibility toggle:** `chkPrikaziFakturisane` allows operators to include already-fakturisane prijemnice in the list instead of showing only invoiceable rows.
- **payment summary hint:** when a prijemnica is already linked to a faktura, the form shows the faktura number plus `uplaćeno / ukupno` using a prebuilt uplata dictionary.
- **selection-to-stavke mapping:** selected list rows are remapped back to original `m_PrijemniceData` indices and transformed into stavka arrays for `CreateFaktura_TX`; the business module now trusts only `PrijemnicaID` and derives `Kolicina`, `Cena`, `Klasa` and `BrojPrijemnice` from `tblPrijemnica`.
- **latest-faktura print helper:** `btnStampaj_Click()` finds the last non-stornirana faktura for the selected kupac and prints it through `PrintFaktura()`.
- **SEF handoff:** `btnSEF_Click()` opens `frmSEF` directly from the fakturisanje shell as the canonical outbound e-invoice handoff path.

### 6.1.33 `frmOtkup` Active Desktop Form Capabilities

The current canonical desktop otkup form exposes the following active capabilities:

- **master-data cascade:** `VrstaVoca -> SortaVoca`, `OtkupnoMesto -> Kooperanti`, and `Kooperant -> Parcele` are active UI cascades backed by workbook master tables.
- **parcela enrichment:** selecting a parcela exposes a display string containing katastarski broj, kultura, površina and hidden `ParcelaID`; parcela selection can also auto-fill fruit type/sort from parcela kultura.
- **two-class capture:** the `chkDveKlase` toggle enables separate Klasa II quantity/price fields while `lblUkupnoKG` recomputes total kg across both classes.
- **duplicate/document guard:** before save the form calls `CheckDuplicate(TBL_OTKUP, COL_OTK_BR_DOK, BrojDokumenta, COL_OTK_DATUM)` when a document number is present.
- **parcela-vs-fruit warning:** if selected parcela kultura conflicts with chosen fruit type, the form forces an explicit operator confirmation instead of silently saving.
- **dual-write business orchestration:** after saving procurement rows, the form can additionally write a kooperant cash payout through `SaveNovac_TX` and then run `ApplyAvansToOtkup_TX` for saved class I and optional class II rows.

### 6.1.34 `frmSEF` Active Desktop Form Capabilities

The current canonical desktop SEF form exposes the following active capabilities:

- **guarded activation:** `UserForm_Activate` uses guarded EH, safe title-bar/chrome ordering and user-facing error feedback so the SEF shell does not crash the desktop lifecycle on activation failures.
- **faktura selector:** `LoadFaktureIntoCombo()` explicitly configures the faktura combo as a two-column control and lists visible `FakturaID` plus `BrojFakture` so operators can choose one invoice context at a time.
- **status surface:** loading a faktura fills labels for local workflow state, exact external `SEFStatus`, `SEFDocumentId`, current version and last error message.
- **event-log review:** `LoadSEFEventsForSelectedFaktura()` loads `tblSEFEventLog` rows into a four-column list for operator audit visibility.
- **button-state policy:** `UpdateSEFButtonStates()` derives allowed actions from the local workflow plus external status, including retry caption changes for `WF_SEF_TECH_FAILED`.
- **safe send command:** `btnPosalji_Click` disables the send button only inside a cleanup-controlled execution path and always restores/recomputes UI state on validation exits, user cancellation or errors.
- **single-invoice command surface:** the form can send, refresh, prepare resubmit, cancel, storno and recover one stuck sending faktura via existing SEF `_TX` flows.
- **destructive action confirmation:** cancel and storno require operator-entered comment plus explicit confirmation before calling the remote SEF mutation functions.
- **batch recovery surface:** `btnRefreshPending` and `btnRecoverAllSending` expose global pending-refresh and stuck-sending recovery actions from the same UI shell.

### 6.1.35 `frmOtkupniBlokovi` Active Desktop Form Capabilities

The current canonical desktop traceability/repair form exposes the following active capabilities:

- **unlinked queue loader:** `LoadNepovezani()` reads `GetUnlinkedOtkupi()` and resolves station, driver and kooperant names for operator review.
- **candidate otpremnica list:** clicking one unresolved otkup loads candidate otpremnice for the same station and date, including number, zbirna reference, quantity and class.
- **manual link repair:** `btnPovezi_Click()` writes the chosen `OtpremnicaID` directly onto the selected otkup row, which is the canonical manual fallback after auto-link ambiguity.
- **auto-link launcher:** `btnAutoLink_Click()` runs `AutoLinkOtkupOtpremnica()` and then reloads queue and status labels.
- **zbirna trace viewer:** selecting a `BrojZbirne` loads reverse-trace rows from `TraceByZbirna()` into a dedicated list.
- **trace PDF export:** `PrintTracePDF()` fills `SledljivostSablon`, computes total otkup kg, prijemnica kg and manjak, then exports `Sledljivost_<BrojZbirne>.pdf` to the workbook path.


## 6.2 GAS Backend Overview
| Module / Area | Responsibility | Exposed actions | Writes | Notes |
|---|---|---|---|---|
| Code.gs router | action dispatch through doGet/doPost | all actions | various sheets | central gateway |
| auth | login, token validation/session, token purge | `login` | login/session logs and auth state | PIN-based; cache + script-properties fallback |
| observability | remote PWA/GAS error logging and retention | `logClientError` | `ErrorLog` workbook | best-effort, non-blocking |
| sync | shared record sync for roles | `sync`, `syncZbirna`, `syncTretman`, `syncOprema`, `syncTrosak` | role-specific sheets | duplicate-aware, idempotent and batch-status aware |
| dispatch | demand, plans, truck status | `getDispecer`, `saveDispecer`, `updateDispecer`, `removeDispecer`, `getKamionStatus`, `updateKamionStatus`, demand actions | DispecerPlan, KamionStatus, demand tab | planning-only boundary |
| meteo | scheduled batch fetch and parcel meteo reads | `getParcelMeteo`, `getParcelMeteoLatest`, `getAllMeteoLatest`, `scheduledMeteoFetch` | MeteoLatest, MeteoHistory | batch + retry |
| fiskalni | parse fiscal receipts and save mappings | `parseFiskalni`, `parseFiskalniImage`, `saveFiskalni`, `saveFiskalniMapiranje`, `createArtikal` | FISKALNI sheets, mapping tab, artikli | private vs master rule critical |
| files / PDF | pdf upload to Drive | `uploadPdf` | Drive | used by otkupni list |
| geo | polygon storage and retrieval plus public POST-read bridge | `saveParcelPolygon`, `getParcelGeo` | parcel geo data | public read/write auth gap still open |

### 6.2.1 Active `Code.gs` Deployment and Routing Contract

The supplied Google Apps Script backend is a single deployed Web App entrypoint whose canonical runtime contract is:

- **entrypoints:** `doPost(e)` is now the primary frontend contract for both authenticated actions and the public geo/meteo read bridge, while `doGet(e)` remains a compatibility/read surface plus health-check entrypoint.
- **deployment model:** manual Web App deployment, execute-as-owner, external URL stored in PWA config.
- **router style:** every request carries an `action` string and the router delegates to function-level handlers.
- **shared response model:** all handlers return `jsonResponse({...})` with structured `success`, `error`, `code`, `processed`, `results` or record payloads depending on action class.
- **sheet bootstrap style:** missing workbooks/sheets are lazily created with bold header row and frozen first row through `getOrCreateSheet(...)` or explicit `insertSheet(...)` branches.

### 6.2.2 Active `doPost` Action Surface (Code.gs snapshot)
| Action | Purpose | Auth gate in current code | Primary writes |
|---|---|---|---|
| `login` | username/PIN auth | none | token cache + `LoginLog` |
| `getParcelGeo` | public geo read bridge for POST-first frontend | none | none |
| `getParcelMeteo` | public parcel meteo read bridge for POST-first frontend | none | none |
| `getParcelMeteoLatest` | public latest meteo read bridge | none | none |
| `getAllMeteoLatest` | public latest meteo bundle read bridge | none | none |
| `saveParcelPolygon` | save parcela polygon and centroid | **currently before token check** | geo workbook `Parcele` |
| `logClientError` | remote client/runtime error logging from PWA into GAS `ErrorLog` | before normal token gate; token is optional and used only for entity attribution when valid | `ErrorLog` |
| `sync` | otkup row sync | token | `OTK-<OtkupacID>` |
| `syncAgromere` | agromere sync | token | `AGRO-<KooperantID>` |
| `syncZbirna` | driver zbirna sync | token | `VOZ-<VozacID>` |
| `saveOtkupniListPdf` | disabled otkupni-list PDF generation/save hook | disabled | `FEATURE_DISABLED` response |
| `uploadPdf` | binary/base64 PDF upload | token | Drive subfolder `OtkupniListovi` |
| `saveWarRoomDemand` | create war-room demand row | token + management | `WarRoomDemand` |
| `removeWarRoomDemand` | delete war-room demand row | token + management | `WarRoomDemand` |
| `updateDemandPrimljeno` | update received kg for demand | token + management | `WarRoomDemand` |
| `updateKamionStatus` | upsert truck status | token | `KamionStatus` |
| `saveDispecer` | create dispatch plan row | token + management | `DispecerPlan` |
| `updateDispecer` | update dispatch plan status | token + management | `DispecerPlan` |
| `removeDispecer` | delete dispatch plan | token + management | `DispecerPlan` |
| `saveIzdavanje` | save agro issuing document | token + management | `Izdavanje` |
| `syncTretman` | sync treatment rows | token | `TRETMAN-<KooperantID>` |
| `syncTrosak` | active kooperant expense sync route | Kooperant/Management role + entity ownership checks | batch response from `processTrosakRecord` |
| `syncOprema` | sync equipment rows | token | `OPREMA-<KooperantID>` |
| `parseFiskalniImage` | QR decode from uploaded image and fiscal parse | token | no canonical business row directly |
| `parseFiskalni` | fetch and parse SUF verification payload | token | no canonical business row directly |
| `saveFiskalni` | persist parsed fiscal lines | token | `FISKALNI-<KooperantID>` |
| `saveFiskalniMapiranje` | persist learned fiscal mapping | token | `FiskalniMapiranje` |
| `createArtikal` | controlled artikli create path | token | `Artikli` |

### 6.2.3 Active `doGet` Action Surface (Code.gs snapshot)
| Action | Purpose | Auth gate in current code | Primary source |
|---|---|---|---|
| `ping` | health check | none | runtime only |
| `getParcelGeo` | compatibility/read-only geo state surface | **none** | geo workbook |
| `getParcelMeteo` | compatibility/read-only parcel meteo fallback surface | **none** | `MeteoLatest` or Open-Meteo |
| `getParcelMeteoLatest` | compatibility/read-only latest meteo row by parcela | **none** | `MeteoLatest` |
| `getAllMeteoLatest` | compatibility/read-only latest meteo bundle | **none** | `MeteoLatest` |
| `getStammdaten` | bootstrap shared read model | token | `Stammdaten`, `Kartice`, optional meteo |
| `getOtkupi` | otkup rows for station/user scope | token + entity | `OTK-<OtkupacID>` |
| `getKartica` | kooperant financial card | token + entity | `Kartice` |
| `getAgromere` | kooperant agromere | token + entity | `AGRO-<KooperantID>` |
| `getMgmtKartica` | management kartica read | token + management | `Kartice` |
| `getMgmtOtkupiByStanica` | management otkupi by stanica | token + management | `OTK-<StanicaID>` |
| `getMgmtSaldoOM` | OM saldo report | token + management | report workbook |
| `getMgmtSaldoKupci` | buyer saldo report | token + management | report workbook |
| `getMgmtOtkupPoOM` | aggregated otkup report | token + management | report workbook |
| `getMgmtPredatoPoKupcu` | aggregated predato report | token + management | report workbook |
| `getMgmtAll` | combined management bootstrap | token + management | report + OTK workbooks |
| `getMgmtFakture` | fakture filtered by kupac | token + management | report workbook |
| `getMgmtFakturaStavke` | invoice lines by faktura | token + management | report workbook |
| `getVozacOtkupi` | otkup rows filtered by vozac | token + vozac | all `OTK-*` workbooks |
| `getVozacZbirne` | driver zbirne | token + vozac | `VOZ-<VozacID>` |
| `getWarRoomDemand` | today demand queue | token + management | `WarRoomDemand` |
| `getDispecer` | today demand + active plans | token + management | `WarRoomDemand`, `DispecerPlan` |
| `getKamionStatus` | current truck status list | token + management | `KamionStatus` |
| `getTretmani` | kooperant treatment rows | token + entity | `TRETMAN-<KooperantID>` |
| `getOprema` | kooperant equipment rows | token + entity | `OPREMA-<KooperantID>` |
| `getKooperantProizvodnja` | parsed production view from kartica | token + entity | `Kartice` |

### 6.2.4 Auth, Session and Access-Control Mechanics

The supplied backend makes the following auth/session mechanics explicit:

- **credential source:** `authenticateUser(username, pin)` reads the `Users` tab inside workbook `Stammdaten`.
- **session token format:** successful login issues a random 64-character token and stores `{ entityID, role, created }` under `TOKEN_<token>`.
- **token retention:** tokens are cached for 86400 seconds (24 hours) and mirrored to `PropertiesService`; cache misses may be recovered from script properties until the 48h hard-expiry window.
- **token cleanup:** `purgeExpiredTokens()` deletes expired or malformed `TOKEN_*` script properties and also triggers `purgeOldErrorLogs()`; `setupTokenPurgeTrigger()` creates a daily 03:00 maintenance trigger in `Europe/Belgrade`.
- **brute-force throttle:** failed attempts are cached per username for 900 seconds and hard-block after 5 failures within that window.
- **audit trail:** login attempts are appended to a lazily created `LoginLog` workbook with `Timestamp | Username | EntityID | Success | Message`.
- **remote error trail:** GAS runtime and PWA client errors are appended to a lazily created `ErrorLog` workbook with `Timestamp | Source | Action | Message | Details | EntityID | Severity`; logger failures are swallowed and logged only to Apps Script `Logger`.
- **frontend transport rule:** the active PWA contract now sends token and action data in POST JSON/text bodies; `apiBuildUrl()` no longer appends `token=` into URL query strings.
- **role boundary:** management-only mutating endpoints explicitly reject non-management users with 403-style payloads; kooperant and vozac reads compare `tokenData.entityID` to request scope.

### 6.2.4a GAS Observability and Runtime Maintenance

The active `Code.gs` includes the following backend observability and maintenance contract:

- **central GAS logger:** `logError(source, action, message, details, entityID)` is the shared writer for GAS-side exceptions and PWA-reported failures.
- **remote client bridge:** `logClientError` is intentionally accepted before the normal token gate so the app can report field/runtime failures even when a user session has expired or is broken.
- **entity attribution:** when `logClientError` includes a valid token, GAS resolves `EntityID` from token data; otherwise it falls back to the request payload `entityID`.
- **severity rule:** timeout-like messages are classified as `warning`; other logged failures default to `error`.
- **payload size guard:** `message` and `details` are truncated before append so large stack traces do not break the logging write path.
- **retention:** `purgeOldErrorLogs()` removes `ErrorLog` rows older than 30 days and is chained from `purgeExpiredTokens()`.
- **non-blocking rule:** logger and purge failures are swallowed or written only to Apps Script `Logger`; they must never block the main sync/auth/business response.

### 6.2.5 Current GAS Sync Semantics

The active backend uses the following hardened sync semantics:

- **station otkup sync:** `processRecord(record, otkupacID)` writes to workbook `OTK-<OtkupacID>` and uses trimmed `ClientRecordID` as the idempotency key.
- **zbirna sync:** `processZbirnaRecord(record, vozacID)` writes to workbook `VOZ-<VozacID>` and also behaves idempotently by trimmed `ClientRecordID`.
- **treatment/equipment sync:** `processTretmanRecord(...)` and `processOpremaRecord(...)` mirror the same idempotent pattern with dedicated server IDs and server timestamps.
- **trošak sync:** `syncTrosak` is active and returns a normal batch result; `processTrosakRecord` is the canonical GAS row processor for `troskovi`.
- **schema contract:** `ensureSheetColumns(sheet, requiredColumns)` creates canonical headers only for empty sheets; existing sheet header mismatch/extra named columns raise `SCHEMA_DRIFT` and block writes.
- **duplicate lookup:** `findByColumn(...)` normalizes compared values through string/trim semantics so whitespace or string/number differences do not create duplicate rows.
- **terminal status preservation:** existing rows with terminal/master/error states such as `Synced>Master`, `Duplicate` or `SyncError:*` are returned idempotently and are not reset to `Synced` by PWA retry.
- **retry mutation rule:** non-terminal existing rows may receive limited retry enrichment; terminal rows do not rewrite `UpdatedAtClient`, `ReceivedAt` or business/enrichment fields.
- **row creation contract:** fresh rows receive a server-side ID, canonical entity scope, server timestamping, business validation and append values in column order derived from the canonical header.
- **batch response contract:** sync batch responses are globally successful only when all records succeed; mixed results return `PARTIAL_FAILURE`, and all-failed batches return `BATCH_FAILED`.
- **registry-assisted lookup:** `getAllOtkupiSheets()` and selected driver/management aggregation reads can use `SheetRegistry` from `Stammdaten` to avoid repeated full folder scans; if registry data is missing or stale, code falls back to folder scanning / name lookup.



### 6.2.5a Google VBA Auth/Sheets Hardening

The active desktop Google integration follows the v6.10 hardening contract:

- **central config ownership:** Google OAuth keys, tokens, folder IDs and sheet IDs live in the central `tblSEFConfig` `ConfigKey/ConfigValue` table. `modConfig` owns both `GetConfigValue` and `SetConfigValue`.
- **OAuth boundary:** `modGoogleAuth` owns OAuth setup, token exchange, refresh and `GetAccessToken`; it does not own general application config writes.
- **safe auth logging:** Google auth error responses are bounded and redacted for configured client secret, access token and refresh token values.
- **expiry fallback:** missing or invalid Google `expires_in` values fall back to a safe 3600-second expiry.
- **clear-before-write safety:** `WriteSheetData` fails if `ClearSheet` fails, preventing stale rows from remaining below newly written data.
- **Drive move visibility:** spreadsheet creation can return a valid ID even if moving into the target folder fails, but the move failure is logged as a warning.
- **exact-name lookup:** `GetSpreadsheetID(title, folderID)` validates exact sheet name in the Drive result and does not trust the first raw `id` occurrence.
- **smoke evidence:** `RunGoogleSyncSmokeSuite` validates config, auth token retrieval, spreadsheet create/write/read/find/add-tab and cleanup behavior.

### 6.2.6 Active Meteo / Risk Pipeline

The supplied `Code.gs` backend makes the current meteo subsystem explicit:

- **geo source split:** parcela coordinates and polygons are read from the separate geo workbook referenced by `GEO_SPREADSHEET_ID`.
- **cached-first meteo read:** `getParcelMeteo()` first tries `getParcelMeteoLatest()` and uses cached data when `LastFetch` is younger than 12 hours.
- **live fallback:** stale or missing cached data falls back to live Open-Meteo forecast retrieval for the parcel centroid.
- **culture-specific thresholds:** risk logic uses `CROP_THRESHOLDS` keyed by culture with `_default` fallback.
- **implemented culture set:** the active threshold table explicitly covers `Visnja`, `Jabuka`, `Sljiva`, `Kruska`, `Breskva` and `Malina`, each with dedicated frost, heat and sprayability bounds.
- **threshold payload shape:** per culture the current threshold contract includes `frostWarn`, `frostDanger`, `heatWarn`, `heatDanger`, `sprayWindMax`, `sprayRainHours`, `optimalTempMin` and `optimalTempMax`.
- **risk outputs:** `assessRisk(...)` derives frost, heat, rain and disease risk, plus aggregate level, min/max temperature and 24h rain totals.
- **72h spray windows:** `calculateSprayWindow(...)` searches the next 72 hours for contiguous spray-safe windows and keeps only windows with at least ~2 valid hours and enough dry hours ahead according to `sprayRainHours`.
- **spray heuristics:** the current suitability rule requires near-zero precipitation, low precipitation probability, wind below crop threshold, temperature roughly `>5°C` and `<35°C`, and humidity below high-risk cutoff.
- **scheduled batch fetch:** `scheduledMeteoFetch()` groups parcels by rounded 0.01 lat/lng buckets, performs batch Open-Meteo fetch, falls back to individual retries, appends `MeteoHistory`, overwrites `MeteoLatest`, and serializes risk/spray/forecast JSON into the latest sheet.
- **trigger plan:** `setupMeteoTriggers()` creates 4 daily triggers for 00:00, 06:00, 12:00 and 18:00 in `Europe/Belgrade` timezone.
- **UI consequence:** this pipeline is not backend-only; its outputs are consumed inline on kooperant parcel cards, parcel-detail meteo panels, home alerts and digital agronom treatment validation.

### 6.2.7 Fiscal Parsing and Private Receipt Storage

The supplied backend confirms this fiscal flow:

- **image path:** `parseFiskalniImage()` decodes base64 image bytes, sends the image to `api.qrserver.com`, extracts the fiscal verification URL and then delegates to `parseFiskalni()`.
- **verification path:** `parseFiskalni()` fetches the supplied verification URL, parses SUF payload/journal text, extracts receipt metadata and item lines, and performs duplicate prevention by `VerificationUrl` inside the kooperant-specific fiscal workbook.
- **matching strategy:** fiscal item auto-match order is `FiskalniMapiranje` → exact artikli name → contains match → keyword score fallback.
- **private persistence:** `saveFiskalni()` writes parsed rows into workbook `FISKALNI-<KooperantID>`; those rows are canonical only inside the kooperant-private fiscal domain and do not automatically mutate master artikli.

### 6.2.8 Management, Dispatch and Agro-Izdavanje Contracts

The supplied backend makes these management-side contracts explicit:

- **war room demand:** `saveWarRoomDemand`, `removeWarRoomDemand` and `updateDemandPrimljeno` manage day-scoped demand rows in `WarRoomDemand`.
- **dispatch plan:** `saveDispecer`, `updateDispecer` and `removeDispecer` manage `DispecerPlan` rows with explicit `planned`/status lifecycle and timestamp updates.
- **kamion status:** `updateKamionStatus()` upserts one row per `VozacID` into `KamionStatus`.
- **combined dispatch read:** `getDispecer()` returns today-only demand plus non-`zavrseno` plans.
- **agro issuing:** `saveIzdavanje()` persists one row per issuing document into `Izdavanje`, serializing `stavke` as JSON and returning a generated `IZD-*` identifier.

### 6.3 PWA Architecture

#### 6.3.1 File Structure
PWA modular refactor is now the active canonical frontend shape and is organized under:

- `src/styles/` — base/layout/components/auth/feature styles/print
- `src/js/utils/` — storage, format, dom, sanitize, async, merge
- `src/js/services/` — db, api, auth, qr
- `src/js/ui/` — toast, signatures, tabs, role-nav
- `src/js/features/kooperant/` — pregled, kartica, koopinfo, parcele, agromere, knjiga-polja, fiskalni, sync, bottom-nav
- `src/js/features/otkup/` — otkup-form, otkup-pregled, otkupni-list, otpremnice, otkup-more, sync
- `src/js/features/vozac/` — zbirna, transport
- `src/js/features/management/` — kooperanti, stanice, kupci, agrohemija, dispecer, mgmt-shell-v2

- `app.js`, `sw.js`, `manifest.json`

The older monolithic/standalone PWA layout is no longer the target architecture. Any remaining legacy wrappers are compatibility bridges only and must not be treated as the primary file-structure model.

#### 6.3.1a Canonical `index.html` Entry Shell Contract
The supplied `index.html` is the active canonical PWA entry shell and makes the following frontend architecture explicit:

- **single-shell entrypoint:** one HTML document owns the app loader, header, root tab bars, role-specific bottom navigation, modal containers, toast surface and all top-level `tab-content` mount points.
- **PWA framing:** the shell declares `manifest.json`, `theme-color`, `apple-touch-icon` and mobile viewport settings with `user-scalable=no`, which means installability/mobile framing is defined at entry HTML level and not delegated to feature modules.
- **self-hosted runtime dependencies:** the shell now loads `html5-qrcode`, `jsPDF`, `Leaflet` and `Chart.js` from local `./vendor/` assets, so QR scanning, PDF output, parcel maps and management charts no longer depend on runtime CDN availability.
- **style segmentation:** CSS is split into `base.css`, `layout.css`, `components.css`, `auth.css`, role/feature styles (`features-otkup`, `features-kooperant`, `features-management`, `features-vozac`) and `print.css`.
- **header contract:** the top header always exposes branding, dynamic `headerInfo`, sync badge, QR-profile action and logout action across roles.
- **cross-role utility surfaces:** the shell keeps global `appLoader`, `qrProfileModal` and `toast` containers outside individual role screens so bootstrap, QR identity display and transient messaging are globally available.

#### 6.3.1b Content-Security-Policy and Static Asset Boundary
The active entry shell declares these concrete client-side boundary rules:

- **CSP default:** `default-src 'self'` is the base policy.
- **script policy:** the active shell now runs with `script-src 'self'`; inline event handlers were removed from both the entry shell and the runtime-rendered feature surfaces, so script-side `'unsafe-inline'` is no longer part of the active posture.
- **repo-wide runtime cleanup rule:** the active frontend contract now forbids runtime HTML strings with inline `onclick`, `onchange` or `oninput`; feature renderers must use delegated `data-action` / `data-route` patterns or explicit `addEventListener(...)` binding.
- **cleanup completion status:** this script-CSP transition is no longer limited to static `index.html`; the current canonical frontend assumes the same rule has been applied across kooperant, otkupac and management feature modules that render dynamic cards, tables, modals and detail views.
- **style policy:** `style-src` still retains `'unsafe-inline'` because the current shell and feature renderers still rely on inline `style="..."` attributes in multiple places.
- **image/network allowlist:** image access explicitly includes `data:`, `blob:`, QR decode service, ArcGIS/OpenStreetMap tiles and configured GAS/Open-Meteo endpoints; `connect-src` is narrowed to self plus GAS/Open-Meteo and no longer includes CDN origins.
- **frontend implication:** QR/image parsing, map tiles and GAS/Open-Meteo fetches are part of the sanctioned client runtime surface; any new provider must be reflected in the entry-shell CSP before it becomes deployable.
- **readiness note:** the shell now has self-hosted vendor assets plus strict script CSP, but style hardening, fuller manifest metadata (`id`, `scope`, categories, maskable purpose), update UX and install polish remain separate production-readiness work items rather than completed architecture.

#### 6.3.1c Canonical Frontend Core Runtime Files
The current modular PWA snapshot also makes these core frontend runtime modules explicit:

- **`src/js/state.js`:** owns a structured-clone based in-memory store with `get`, `set`, `patch`, `subscribe` and `reset`, plus a listener map for path-based subscriptions.
- **legacy-global compatibility layer:** the state layer still exposes `db`, `stammdaten`, `mgmtData`, `qrScanner`, `selectedMera` and `parcelExpertOpen` through `Object.defineProperty(window, ...)` so older feature code can continue reading/writing legacy globals while new code moves toward `AppState`.
- **IndexedDB helper surface:** the currently supplied frontend core snapshot also exposes `openDB`, `dbPut`, `dbGet`, `dbGetAll`, `dbGetByIndex` and `dbDelete` as window-scoped promise helpers around IndexedDB transactions.
- **`src/js/app.js`:** is the canonical bootstrap/orchestration owner for runtime init, stammdaten refresh, background sync, QR profile, service-worker registration and safe role boot.
- **`src/js/services/auth.js`:** owns login-shell rendering, logout cleanup, role visibility toggling and header-brand switching.
- **`src/js/ui/role-nav.js`:** owns bottom-nav visibility, active-state synchronization and role-specific root routing dispatch.

#### 6.3.2 App Bootstrap
`DOMContentLoaded → bootstrapApp()` performs:

1. session validation
2. IndexedDB open
3. stammdaten cache load
4. role visibility + branding
5. shell event binding
6. role bootstrap
7. bottom-nav engine init
8. sync badge update
9. online/offline listener bind
10. background sync start where applicable
11. background stammdaten refresh
12. guaranteed loader hide in `finally`

#### 6.3.2a Script Load Order and Bootstrap Invariants
The supplied entry shell makes the frontend boot pipeline explicit and canonical:

- **phase 1 — utilities:** storage/format/dom/sanitize/async/merge helpers must load before higher layers.
- **phase 2 — config/state:** `config.js` and `state.js` define runtime constants and mutable app state before services/features bind.
- **phase 3 — services:** `db.js`, `api.js`, `auth.js` and `qr.js` provide storage/API/auth/scan infrastructure consumed by feature modules.
- **phase 4 — UI helpers:** `toast.js`, `signatures.js`, `tabs.js`, `role-nav.js` attach shared rendering and navigation helpers.
- **phase 5+ — role features:** kooperant, otkup, vozac and management feature modules are loaded before final app bootstrap.
- **final bootstrap invariant:** `app.js` is the last script and is the canonical app-start owner.
- **current shell invariant:** current HTML keeps one canonical `db.js` include and one canonical management shell include (`mgmt-shell-v2.js`); legacy parallel shell loading is no longer part of the active runtime.

#### 6.3.2b `state.js` Runtime State Contract
The supplied runtime-state layer makes these active rules explicit:

- **initial in-memory envelope:** `AppState` starts from `initialState` holding `db`, `stammdaten`, `mgmtData`, `qrScanner`, `selectedMera`, `parcelExpertOpen`, an `init` block (`domReady`, `dbReady`, `stammdatenReady`, `appReady`, `bootError`) and a `sync` block (`inFlight`, `lastRunAt`).
- **path-addressable updates:** state writes are path-based (`'init.appReady'`, `'sync.inFlight'`, etc.), with `set()` creating missing nested objects and notifying listeners attached to the exact path.
- **patch semantics:** `patch(path, partial)` is shallow-merge based and intended for grouped runtime flags, not deep document merges.
- **listener model:** `subscribe(path, fn)` stores listeners in a `Map<path, Set<fn>>` and returns an unsubscribe function.
- **reset contract:** `reset()` restores all top-level keys from `initialState` via `structuredClone`, making session/runtime reset deterministic.
- **transitional compatibility:** legacy frontend code is still allowed to read/write `window.db`, `window.stammdaten`, `window.mgmtData`, `window.qrScanner`, `window.selectedMera`, `window.parcelExpertOpen` and `window.appRuntime`, but those are compatibility shims over canonical shared state.

#### 6.3.2c `app.js` Bootstrap and Runtime-Orchestrator Contract
The currently supplied `app.js` snapshot further clarifies the canonical bootstrap/runtime behavior:

- **runtime flags:** app-shell runtime ownership is normalized around the shared runtime path exposed through `window.appRuntime`, rather than a second independent private runtime object in `app.js`.
- **session gate:** `bootstrapApp()` short-circuits to `showLoginScreen()` when the app lacks `authToken` plus a canonical session entity identity; `entityID` is the active cross-role key, with legacy `otkupacID` tolerated only as compatibility fallback.
- **boot order:** after session validation the app opens IndexedDB, loads cached stammdaten, applies role visibility and header info, binds shell/connectivity events, runs role bootstrap, starts role-aware background sync, marks runtime ready and then refreshes stammdaten in background.
- **role bootstrap specifics:** otkupac boot initializes the new otkup form UI when available; kooperant boot guards on stammdaten and pre-populates parcel-aware agro flows; management boot populates station/kupac dropdowns, tries `prefetchMgmtData()` and prefers `mgmtShellInit()`.
- **default-date priming:** boot sets current date into pregled/otprema/management date filters when those inputs are empty.
- **cached-stammdaten contract:** local stammdaten are read from IndexedDB store `CONFIG.STAMM_STORE` under row key `all`; network refresh normalizes and rewrites the same cache row.
- **normalization rule:** `normalizeStammdaten(...)` guarantees array presence for `kooperanti`, `kulture`, `config`, `parcele`, `stanice`, `kupci`, `vozaci`, `artikli`, `magacinkoop`, `meteoLatest` and `kartice`, while also aliasing legacy `meteolatest` into `meteoLatest`.
- **shell event model:** app-shell interactivity is now centered on one delegated event layer (`document` click/change/input) plus targeted modal/feature listeners, replacing the old inline-handler ownership model.
- **update event contract:** successful stammdaten refresh dispatches `window.dispatchEvent(new CustomEvent('stammdaten:updated', ...))`; the update handler invalidates feature caches and repopulates role-specific dropdowns/views.
- **safe sync gate:** `syncQueueSafe()` runs only when online, skips Management, respects shared in-flight runtime guards, and delegates to role-specific sync functions for Otkupac, Kooperant and Vozac before always flipping badge/runtime state back in `finally`.
- **service worker runtime rule:** the shell registers `./sw.js`, calls `reg.update()` every 60 seconds and shows an informational toast when a newly installed worker reaches `activated`; the active cache generation in the current snapshot is `AgriX-v10`, and the worker cache list now includes the self-hosted vendor assets.

#### 6.3.2d Runtime Event-Delegation and Feature-Render Contract
The current frontend snapshot makes the following cross-feature UI interaction rules explicit:

- **delegation-first rule:** actions emitted from dynamic cards, rows, modals and section headers should be routed through `data-action`, `data-route`, `data-tab`, `data-index`, `data-record-key` and similar `data-*` contracts rather than inline JS attributes.
- **local-module exception:** feature modules that render complex runtime markup (for example fiscal tables, agrohemija baskets/modals, dispatch boards or invoice-detail expansions) may attach local delegated listeners to their own root containers instead of sending every action through the global shell dispatcher.
- **CSP-safe popup rule:** parcel popups, otkup detail modals, otprema detail cards, faktura drilldowns, signature modals and similar dynamic UI surfaces are expected to remain compatible with `script-src 'self'` and therefore must not rely on raw inline event handlers.
- **navigation helper cleanup:** legacy navigation helpers that depended on parsing `onclick` attributes are no longer part of the active runtime contract; tab/button lookup is expected to use current DOM `data-route` / `data-tab` markers instead.
- **architecture consequence:** frontend behavior ownership is now split cleanly between canonical shell delegation in `app.js`, role navigation in `role-nav.js`, and feature-local delegates where module-scoped rendering makes that the safer option.

#### 6.3.3 Role Routing
- Kooperant → `home`
- Otkupac → `otkup`
- Vozac → `zbirna`
- Management → `pregled`
- management shell supports two-level root/sub navigation through the canonical V2 shell only; legacy mount-compatibility runtime is no longer active

#### 6.3.3a Canonical Role Tab and Root-Surface Map
The supplied HTML shell defines the following active role-visible surfaces:

- **Otkupac root tabs:** `otkup`, `pregled`, `otpremnice`, `queue`.
- **Kooperant root tabs:** `home`, `parcele`, `agromere`, `knjigapolja`, `more`, with `kartica` and `koopinfo` reachable as secondary/detail routes.
- **Management root tabs:** `dashboard`, `pregled`, `dispecer`, `otkup`, `partneri`, `agro`.
- **Vozac root tabs:** `zbirna`, `transport`.
- **dual navigation contract:** desktop-style top tab bars remain in DOM, while role-specific bottom nav bars are the canonical mobile-first navigation shells for kooperant, otkupac, management and vozac roles.
- **mount-point rule:** every role/root route resolves into a dedicated `tab-content` container in the entry HTML; feature modules render into those predeclared DOM anchors instead of creating root screens dynamically.

#### 6.3.3b Canonical Feature Mount Surfaces by Role
The supplied entry shell further makes these UI mount contracts explicit:

- **Otkupac:** `tab-otkup` hosts a five-step otkup wizard; `tab-pregled` hosts quick filters, summary KPIs and detail modal; `tab-otpremnice` hosts root/assign/success otprema states; `tab-queue` hosts profile, signature and sync/queue diagnostics.
- **Kooperant:** `tab-home` hosts KPI dashboard plus quick actions; `tab-kartica` hosts financial card summary/list; `tab-parcele` hosts map/list/detail parcel workspace; `tab-agromere` hosts tretmani/evidencija/oprema/kalendar sections; `tab-knjigapolja` hosts pregled/proizvodnja/troškovi/lager sections; `tab-more` hosts secondary actions.
- **Management:** `tab-mgmt-dashboard`, `tab-mgmt-pregled`, `tab-mgmt-dispecer`, `tab-mgmt-otkup`, `tab-mgmt-partneri` and `tab-mgmt-agro` are now explicit root shells, each with its own nested sub-navigation/mount zone.
- **Vozac:** `tab-zbirna` separates main-view vs create-view behavior, while `tab-transport` is the compact transport overview surface.
- **global modal/popup rule:** detail modals (`pregled`, `otprema`, quick actions, QR profile) are declared once in entry HTML and reused by feature modules.


#### 6.3.3c `auth.js` Login Shell and Role-Visibility Contract
The supplied auth/runtime shell makes the current client auth flow explicit:

- **login-screen ownership:** `showLoginScreen()` hides the main header, top tab bar, all `tab-content` blocks and all `sub-tab-bar` blocks, removes visible role-bottom-nav classes and injects a dedicated `loginContainer` DOM shell if needed.
- **credential model:** login collects `username` plus 4-digit style PIN, then delegates authentication to `apiPost('login', { username, pin })`.
- **session persistence:** successful login stores `authToken`, `userRole`, canonical `entityID`, `entityName` and `username` in local storage, then hard-reloads the app.
- **compatibility alias rule:** when the active role is `Otkupac`, frontend login/bootstrap may still mirror `entityID` into legacy `otkupacID` storage for older Otkupac-specific payload contracts.
- **legacy naming debt cleanup:** `entityID` is now the canonical cross-role session key; `otkupacID` remains only as an Otkupac compatibility alias where older payload contracts still expect that field name.
- **logout cleanup:** `doLogout()` removes all session keys, clears visible bottom-nav states/body classes for active role shells and reloads the application.
- **role DOM gate:** `applyRoleVisibility()` toggles `.role-otkupac`, `.role-kooperant`, `.role-vozac` and `.role-management` element groups directly from `CONFIG.USER_ROLE`.
- **branding switch:** `applyHeaderBranding()` swaps the header logo between Gazdinstvo and Otkup branding depending on whether the active role is `Kooperant`.

#### 6.3.3d `role-nav.js` Bottom-Nav Engine Contract
The supplied role-navigation engine adds these active navigation rules:

- **per-role config map:** navigation is driven by `getRoleNavConfig()` which returns `navId`, `bodyClass`, `type`, `defaultTab` and a `tabMap` for each role.
- **dispatch mode split:** kooperant, otkupac and vozac bottom nav buttons route through `showTab(...)`, while management bottom nav routes through `showMgmtRoot(...)`.
- **visibility rule:** `updateRoleNavVisibility()` first hides all known role navs and removes all body spacing classes, then activates only the nav/body class for the current role.
- **active-state rule:** `updateRoleNavActive()` always clears active state across **all** role navs first and then re-applies the active button for the current role only.
- **DOM/state sync:** active nav resolution comes from the active `.tab-content`, except management where `window.mgmtShellState.activeRoot` is preferred when available.
- **layout contract:** role-specific body classes such as `has-koop-bottom-nav`, `has-otkup-bottom-nav`, `has-mgmt-bottom-nav` and `has-vozac-bottom-nav` are intentionally applied in both mobile and desktop layouts because spacing logic is shared.
- **window API:** `updateRoleNavVisibility`, `updateRoleNavActive`, `syncRoleNavActiveFromDom`, `showRoleNavTab` and `initRoleNavEngine` are exported globally for feature-shell interop.

#### 6.3.3e `tabs.js` Non-Management Router Contract
The currently supplied tab-router snapshot adds these active navigation rules:

- **base ownership:** `showTab(tabName, btn)` is now strictly the canonical non-management router for switching `.tab-content` and `.tab-btn` active states.
- **agromere cleanup rule:** leaving any tab other than `agromere` clears `agroState.geoWatchId` when present, so geolocation watch cleanup is tab-driven rather than component-destroy-driven.
- **feature-entry hooks:** tab routing eagerly triggers feature loads by tab name, including `loadOtkupPregled`, `loadOtpremaOverview`, `loadPregled`, `loadKartica`, `loadParcele`, `loadAgronom`, `loadKoopInfo`, `loadVozacData`, `loadVozacTransport` and `loadKnjigaPolja`; `queue` additionally dispatches to `loadOtkupacMore()` for otkupac role.
- **nav resync:** after each tab switch the router schedules `updateRoleNavActive()` through a zero-delay timeout so bottom-nav state re-syncs from the real DOM.
- **management ownership rule:** management root navigation is no longer intercepted in `tabs.js`; it stays owned by `showMgmtRoot(...)` and the dedicated role-nav / management-shell path.

#### 6.3.3f Legacy Feature Bottom-Nav Helper Contract
A second, feature-local bottom-nav helper layer is also present in the supplied PWA code and makes these constraints explicit:

- **limited scope:** the helper only models kooperant and otkupac bottom-nav concerns through `getActiveBottomNavConfig()` and does not own management routing.
- **button utility surface:** it provides `updateBottomNavButtons(...)`, `findLegacyTabBtn(...)`, `syncKooperantFromMore()` and `invalidatePregledCacheSafe()` as secondary helpers for older feature code.
- **explicit boundary:** the file itself states that it must not patch global `showTab` and must not take ownership over management navigation, which keeps the remaining kooperant/otkupac convenience helpers isolated from the canonical `role-nav.js` + `showMgmtRoot(...)` ownership model.
- **more-screen sync helper:** `syncKooperantFromMore()` delegates to `syncKooperantNow()`, uses toast feedback and invalidates pregled cache after successful manual sync.

#### 6.3.4 Shared Utilities
- `apiFetch` / `apiPost` plus normalized `apiFetchSafe` / `apiPostSafe`
- `escapeHtml`
- `safeAsync` and `reportClientError`
- storage wrappers
- DOM wrappers
- `mergeOfflineRecords`
- format helpers like `fmtDate`, `fmtStanica`, `normalizeIso`, `getTodayIsoDate`, `getRelativeIsoDate`, `toIsoDateOnly`, `localIsoDateFromDate`, `formatKg` and `formatMoney`

#### 6.3.4a `dom.js` Minimal DOM Utility Contract
The supplied DOM helper layer makes the current shared UI utility surface explicit:

- **selector helpers:** `qs(selector, root)` and `qsa(selector, root)` are the canonical query wrappers for one vs many DOM lookups.
- **element access rule:** `byId(id)` is the standard root-element accessor used throughout feature modules instead of repeated raw `document.getElementById(...)` calls.
- **display helpers:** `showEl(el, displayValue)` and `hideEl(el)` are the active visibility toggles and intentionally operate by inline `style.display` mutation.
- **content helpers:** `setText(el, text)` and `setHtml(el, html)` are the standard text vs trusted-HTML render entry points.
- **class helpers:** `addClass`, `removeClass` and `toggleClass` provide one canonical DOM-class API reused across tab, nav, toast and feature render code.

#### 6.3.4b `storage.js` Local Storage and Device Identity Contract
The supplied storage helper layer clarifies the active browser-persistence contract:

- **safe wrapper rule:** `getLs`, `setLs` and `removeLs` are thin exception-safe wrappers over `localStorage`, returning fallbacks/booleans rather than throwing.
- **session-storage boundary:** these wrappers are intended for lightweight local session/runtime values only and do not replace IndexedDB for operational entity queues.
- **device identity contract:** `getDeviceID()` lazily creates and persists one browser/device-scoped identifier shaped as `DEV-xxxxxxxx` using `crypto.randomUUID()`.
- **offline provenance rule:** local field records may therefore carry a stable device identifier even before first successful server sync.

#### 6.3.4c `merge.js` Generic Offline-Merge Contract
The supplied merge helper narrows the active offline-first reconciliation rules:

- **generic merge core:** `mergeOfflineRecords(local, server, normalizeLocal, primaryKey)` is the canonical client merge primitive across multiple list/detail modules.
- **identity rule:** `clientRecordID` is the default merge key unless an alternate `primaryKey` is explicitly supplied.
- **overlay order:** normalized server rows form the base snapshot and normalized local rows are then overlaid.
- **precedence rule:** local `pending` or `syncing` rows always replace matching server rows, regardless of timestamps.
- **freshness rule:** local already-synced rows replace server rows only when local `updatedAtClient` is newer than the server-side update timestamp chain.
- **orphan-local rule:** any local row without a server match is preserved in the merged result.

#### 6.3.4d `toast.js` User-Feedback Contract
The supplied toast helper defines a minimal but explicit global feedback model:

- **single-toast surface:** `showToast(msg, type)` writes into one global `#toast` element rather than creating stacked notifications.
- **state model:** toast visibility is expressed through CSS classes `toast show <type>`.
- **auto-dismiss rule:** the current implementation always clears the toast after 3 seconds and does not expose a per-message duration override.
- **type vocabulary:** the active runtime uses at least `info`, `success` and `error` semantic classes.

#### 6.3.4e `qr.js` Scan and QR-Profile Utility Contract
The supplied QR helper clarifies the current shared scanning/QR-render model:

- **scanner ownership:** `startQRScan()` owns the canonical kooperant QR-reader flow for the otkup form and renders into `#qr-reader`.
- **single-instance cleanup rule:** before creating a new `Html5Qrcode` instance, the helper stops/clears any previously retained `qrScanner` and nulls the global reference.
- **camera preference:** scanning always requests the environment-facing camera with a 250×250 QR box and `fps = 10`.
- **scan handoff:** successful decode stops the scanner, hides the reader container and delegates parsed content to `onQRScanned(decodedText)`.
- **QR generation rule:** `generateQRCode(canvasId, text)` currently uses the remote QRServer image endpoint to draw a 250×250 QR into a canvas, with a branded text fallback when image generation fails.

#### 6.3.4f `signatures.js` Shared Signature-Pad Contract
The supplied shared signature module makes the canvas-signature infrastructure explicit:

- **registry model:** active pads are tracked in a module-scoped `Map` keyed by canvas ID, which prevents duplicate binding and supports modal recreation.
- **high-DPI rule:** `setupCanvas(...)` resizes canvases using device-pixel ratio, resets transforms and re-applies drawing style so signatures remain sharp on mobile screens.
- **input model:** the pad supports both mouse and touch events with `preventDefault()` and round-cap stroke drawing.
- **rebind safety:** re-initializing the same canvas ID on a new DOM element automatically unbinds the old listeners and binds the new canvas instance.
- **public API:** `initSignaturePad`, `clearSignature`, `getSignatureData`, `destroySignaturePad` and `destroyAllSignaturePads` are exported globally for feature reuse.
- **empty-signature rule:** `getSignatureData(canvasId)` returns an empty string when a pad has no ink, which is how business flows distinguish unsigned vs signed documents.

#### 6.3.4g `api.js` Frontend API Client Contract
The currently supplied API client layer makes these runtime rules explicit:

- **single transport rule:** the active frontend request contract is POST-first; `apiNormalizePayload(...)` folds action/query-style parameters into one request body together with `CONFIG.TOKEN`.
- **URL rule:** `apiBuildUrl()` now returns only `CONFIG.API_URL`; token transport is no longer implemented through URL query strings.
- **safe-request core:** `apiRequestSafe(...)` is the canonical normalized fetch wrapper and uses `AbortController` with a default 20-second timeout.
- **HTTP/error contract:** the wrapper explicitly checks `resp.ok`, parses raw text defensively, handles bad JSON, network errors and timeouts, and returns one normalized result object through `apiBuildResult(...)`.
- **normalized result shape:** safe helpers resolve to `{ ok, status, data, error, code, isTimeout, isNetworkError, isAuthError }`.
- **auth-failure rule:** `apiHandleAuthFailure(...)` detects `401` payloads, shows a visible toast and triggers logout instead of leaving the app in a silent expired-session state.
- **compatibility surface:** `apiFetch(...)` / `apiPost(...)` remain raw/backwards-compatible helpers returning payload-or-null, while `apiFetchSafe(...)` / `apiPostSafe(...)` are the canonical normalized helpers for new code.

#### 6.3.5 Feature Modules
- **Kooperant:** dashboard, parcel GIS, agromere, knjiga polja, fiskalni, kartica
- **Otkupac:** form, pregled, otprema, otkupni list, queue/profile/sync
- **Vozac:** zbirna, transport
- **Management:** overview, dispatch, partneri, agrohemija

#### 6.3.5a Management Shell Split Contract
The supplied entry shell confirms that management UI currently runs on a split-root contract rather than one flat tab page:

- **overview root:** `tab-mgmt-pregled` is the lightweight overview/home shell.
- **dashboard root:** `tab-mgmt-dashboard` is the KPI/chart-centric executive shell.
- **operational roots:** `tab-mgmt-dispecer`, `tab-mgmt-otkup`, `tab-mgmt-partneri` and `tab-mgmt-agro` each own their own nested sub-nav and dedicated mount container.
- **canonical shell rule:** legacy `tab-mgmt` compatibility wrapper and legacy `mgmt-shell.js` are no longer part of the active runtime; `mgmt-shell-v2.js` plus explicit `tab-mgmt-*` roots are the canonical management shell contract.

#### 6.3.5b Kooperant `kartica.js` Contract
The supplied kooperant financial-card module makes the following behavior canonical:

- **data source:** `loadKartica()` fetches `action=getKartica&kooperantID=<ENTITY_ID>` and treats that endpoint as the active source for card lines.
- **cache rule:** successful loads are memoized in `karticaCache` until `invalidateKarticaCache()` is called.
- **export-cleanup rule:** rows where `Opis === 'UKUPNO'` are filtered out before render so aggregate export rows do not appear as business transactions.
- **render contract:** `renderKartica(...)` recomputes `Zaduzenje`, `Razduzenje` and terminal `Saldo` from the filtered rows and renders them into the dedicated summary cards plus detailed line list.

#### 6.3.5c Kooperant `koopinfo.js` Contract
The supplied kooperant info module is a read-only configuration surface with these rules:

- **source model:** all values come from `stammdaten.config`, without a dedicated fetch.
- **parameter contract:** active fields include `OtkupAktivan`, `RadnoVremeOd`, `RadnoVremeDo`, `SezonaOd`, `SezonaDo` and every config row whose `Parameter` starts with `Cena`.
- **display semantics:** otkup status is rendered as a green/red active flag, while price rows are displayed as human-readable `Cena*` parameters in RSD/kg.
- **write boundary:** the screen is read-only and does not author config.

#### 6.3.5d Kooperant `agromere.js` / Digitalni Agronom Contract
The supplied digital-agronom module significantly tightens the kooperant agronomy contract:

- **state model:** `agroState` holds parcela, mera, selected article, dosage, equipment, note, timer, geo start/end, meteo snapshot, karenca, local lager/oprema lists and active geolocation watch ID.
- **cache model:** `_tretmaniCache` is a 30-second merged local/server cache for treatment history and karenca checks.
- **boot orchestration:** `loadAgronom()` resets state, loads lager from `stammdaten.magacinkoop`, merges server/local/preset equipment, populates kooperant parcels, starts geolocation, loads history, restores the step-1 UI and optionally launches background `syncTretmani()`.
- **equipment merge rule:** server equipment, locally entered items and `OPREMA_PREDLOZI` presets are merged into one deduplicated `opremaList`; however, newly created equipment still syncs directly through `syncOprema` only when online.
- **parcel detection rule:** geolocation uses polygon containment or centroid distance; parcels may be auto-selected when inside/nearby, with automatic selection at about 50m and suggestion mode within about 200m.
- **meteo and karenca rule:** selecting a parcel loads meteo strip data, evaluates active karenca from merged treatments, disables `Berba` while karenca is active, and blocks/warns `Zastita` when wind/temperature/humidity thresholds are violated unless the user explicitly triggers meteo override.
- **work-timer rule:** the module exposes explicit `start/stop` work timing, keeps one active timer panel/sticky timer state, persists start/end times plus `TrajanjeMinuta`, and couples that timer output into both treatment save payload and derived labor-cost logic downstream.
- **smart dosage rule:** for `Zastita` and `Prihrana`, the article picker is filtered from current kooperant lager by type, then recommendation is calculated as `DozaPoHa × PovrsinaHa`, optionally rounded to whole packages and annotated with insufficient-stock warnings.
- **save model:** `agroSaveTretman()` writes a fully described local `tretmani` record first (timestamps, timer window, start/end geo, karenca fields, dosage fields, sync metadata), invalidates cache, then attempts `syncTretmani()` online and finally reloads history from merged state.
- **history rule:** `agroLoadIstorija()` renders non-deleted merged treatment rows sorted by `datum`, then `updatedAt*`, then `clientRecordID`.

#### 6.3.5e Kooperant `fiskalni.js` / Fiskalni Lager Contract
The supplied fiscal-receipt / fiskalni-lager flow makes the following frontend behavior canonical:

- **module identity:** this is the active kooperant-side Fiskalni Lager module, not only a passive receipt parser; its purpose is to transform verified fiscal lines into private lager-aware artikl rows for Knjiga Polja.
- **scan strategy:** the preferred path is native `BarcodeDetector` + live camera stream; if unsupported, the UI falls back to photo capture/input with server-side QR extraction via `parseFiskalniImage`.
- **image normalization:** photo fallback resizes the image to roughly 1024px max dimension and JPEG quality ~0.85 before upload, to reduce payload while preserving QR readability.
- **parse pipeline:** a scanned or decoded verification URL is sent to `parseFiskalni`; successful parse returns receipt meta plus line items and match-confidence hints.
- **mapping UI:** `renderFiskalniResult()` renders one selectable row per parsed line, shows exact/fuzzy/manual match status and allows manual artikl assignment or a special `__NEW__` private-article path.
- **private article rule:** `fiskalniCreateNewArtikal()` creates a temporary `PRIV-*` item ID only inside the staged fiscal line; it does not write into master `Artikli` and is intentionally private to fiscal storage.
- **save rule:** `fiskalniSaveToLager()` saves only checked rows, hard-blocks checked rows that still have no resolved `artikalID`, posts the selected lines to `saveFiskalni`, and then optionally sends learned manual mappings to `saveFiskalniMapiranje` in fire-and-forget mode.
- **cancel/reset:** `fiskalniCancel()` clears staged receipt metadata/items and hides the result surface without writing.

#### 6.3.5f Kooperant `knjiga-polja.js` Contract
The supplied field-book module clarifies these canonical data and calculation rules:

- **working set:** `kpData` contains `proizvodnja`, `tretmani`, `troskovi` and `lager`, while `_kpLoaded` marks whether the view has already been initialized.
- **production derivation:** production is derived from `stammdaten.kartice` by parsing rows whose `Opis` begins with `Otkup`, skipping `UKUPNO`, and extracting `VrstaVoca`, `Klasa` and `Kolicina` from the description text.
- **treatment source:** treatment rows come from `getTretmaniCached(false)` when available, otherwise from `getTretmani`.
- **cost merge rule:** local IndexedDB `troskovi` rows are merged with server `getTroskovi` rows through `mergeOfflineRecords` when available, with a fallback local-wins merge.
- **auto labor rule:** `kpCalcRadnaSnaga()` derives synthetic `radna_snaga` troškovi from treatment duration and config parameter `CenaRadaSat`; these rows are marked `_auto` and remain derived, not authoritative source records.
- **bilans rule:** `kpLoadBilans()` computes `proizvodnja − agrohemija − troškovi`, where agrohemija cost is treatment consumption multiplied by `WAC` or `CenaPoJedinici` from current lager rows.
- **write path:** `kpSaveTrosak()` is local-first — it writes a pending `troskovi` record to IndexedDB. Online `syncTrosak` is active in GAS and must return batch results; the UI preserves local/error metadata until server confirmation.
- **UI structure:** the module owns pregled/proizvodnja/troškovi/lager sections, section switching through `showKnjigaSection(...)`, KPI synchronization from rendered bilans rows and a separate seasonal `kpRenderPotrosnja()` summary derived from treatment consumption.


#### 6.3.5g Kooperant `parcele.js` Contract
The supplied kooperant parcel/GIS module makes these runtime and rendering rules canonical:

- **parcel map ownership:** `loadParcele()` owns the kooperant parcel list + map screen and initializes one Leaflet map instance per shell lifetime, with `_parceleLoaded` used as the screen-level load guard.
- **style contract:** polygon parcels use dedicated base and selected styles (`kooperantParcelStyle`, `kooperantSelectedParcelStyle`), while point-only parcels fall back to highlighted circle markers.
- **meteo prewarm rule:** the module prepopulates `window.meteoCache` directly from exported `stammdaten.meteoLatest` before doing any per-parcel API fallback, so parcel cards can often render meteo state without immediate network fetch.
- **geo read contract:** parcel geometry still comes from `getParcelGeo(parcelaId)` per parcel; polygons are rendered from `PolygonGeoJSON`, otherwise valid `Lat`/`Lng` are rendered as point markers.
- **popup contract:** every mapped parcel popup shows `KatBroj`, `Kultura`, `PovrsinaHa`, `KatOpstina`, `GGAPStatus`, raw `ParcelaID` and an explicit button into `openParcelaDetail(...)`.
- **CSP-safe action rule:** parcel list cards, popup detail buttons and expert-panel toggles now use delegated `data-action` hooks instead of inline event handlers, so GIS/detail interactions remain compatible with the stricter script CSP.
- **selection semantics:** `focusParcel(...)` synchronizes list click → map focus → popup open → parcel highlight → parcel detail navigation, making the parcel detail screen the canonical cross-link target for GIS interactions.
- **inline meteo rule:** card-level meteo uses a 6-hour `METEO_CACHE_TTL`, shows compact current/risk/spray/3-day forecast state and adds an expandable expert panel for soil moisture, soil temperature, ET₀, UV and solar radiation when those fields exist.
- **detail-screen contract:** `openParcelaDetail(...)` is an async cross-module entrypoint that ensures `loadKnjigaPolja()` has run, then renders parcel detail tabs for `osnovno`, `meteo`, `radovi`, `troskovi` and `proizvodnja`.
- **detail data-source rule:** parcel detail `radovi`, `troskovi` and `proizvodnja` are rendered from already-aggregated kooperant datasets (`kpData`, cached meteo, stammdaten), not from one-off dedicated parcel endpoints.
- **search/filter contract:** the list view provides client-side text filtering plus a kultura filter built from the kooperant’s exported parcel set.
- **cache invalidation:** `invalidateParceleCache()` resets only the loaded-flag, meaning the next open of the screen re-runs parcel rendering and geo/meteo initialization.

#### 6.3.5h Kooperant `pregled.js` Contract
The supplied kooperant home/dashboard aggregator clarifies these canonical orchestration rules:

- **aggregator role:** `loadPregled(forceRefresh)` is not a standalone data source; it is a cache-backed orchestrator over `kartica.js`, `knjiga-polja.js`, `agromere.js`, `parcele.js`, `koopinfo.js` and kooperant sync metadata.
- **cache window:** `pregledCache` has a 30-second TTL and is invalidated explicitly through `invalidatePregledCache()`.
- **build contract:** `buildPregledData()` assembles one composite home payload containing `hero`, `kpi`, `alerts`, `bilans`, `kartica` and `info` sections.
- **info source:** otkup status, working hours and season labels are derived directly from `stammdaten.config`.
- **parcel-alert rule:** the module scans kooperant parcels against preloaded `window.meteoCache` and raises alert cards from parcel risk items or from absence of valid spray windows.
- **today-work rule:** the home KPI for today’s work is computed from merged treatment records, while stale/older treatments can influence alert generation and suggested next actions.
- **financial-summary rule:** the kartica summary is derived either from cached `karticaCache` or from a fresh `getKartica` fetch with `UKUPNO` rows removed, mirroring the dedicated kartica screen.
- **bilans dependency:** when `kpData` is not yet ready, the dashboard triggers `loadKnjigaPolja()`/`kpFetchAll()` and then reuses those derived datasets for proizvodnja/agrohemija/troškovi/rezultat calculations.
- **sync-alert rule:** the dashboard inspects IndexedDB pending treatment rows through the canonical `tretmani` store and turns that into explicit sync-status alerts.
- **dynamic agronomy suggestion:** `getDynamicRadoviPredlog(...)` produces crop/month-specific next-work recommendations (e.g. for jabuka, višnja, šljiva) and then mutates them further when current meteo risk or missing spray windows should change the advice.
- **alert navigation contract:** `onPregledAlertClick(...)` deep-links from dashboard alerts into `parcele`, `agromere`, `koopinfo` or sync-related screens; parcel alerts can chain directly into `focusParcel(parcelaID)`.
- **home quick-action contract:** the module owns the quick-actions modal and its routing helpers (`goToNewRad`, `goToNewTrosak`, `goToScanRacun`, `goToKartica`, `goToKnjigaPolja`) for the kooperant shell.

#### 6.3.5i Management Legacy-Shell Retirement Contract
Legacy `mgmt-shell.js` is no longer part of the active frontend runtime.

- **runtime ownership:** Management routing and rendering are owned only by `mgmt-shell-v2.js`.
- **boot-helper migration:** `prefetchMgmtData()` and `populateMgmtKupciDropdown()` were migrated into `mgmt-shell-v2.js` so Management bootstrap no longer depends on the retired legacy shell.
- **legacy wrapper removal:** legacy `tab-mgmt` / `mgmtSubBar` compatibility wrapper is removed from the active entry shell.
- **architecture consequence:** Management no longer runs in a parallel-shell or legacy bridge mode.

#### 6.3.5j Management `mgmt-shell-v2.js` Contract
The supplied V2 management shell is now the canonical frontend owner of management root navigation:

- **shell state:** `window.mgmtShellState` tracks `activeRoot`, partner segment, koop/kup/otkup/agro sub-selection, dashboard period and mount status.
- **init contract:** `mgmtShellInit()` attempts dispatch preload, opens root `pregled` and then re-synchronizes role-nav visibility/active state.
- **explicit root model:** `showMgmtRoot(...)` toggles exactly one of `dashboard`, `pregled`, `dispecer`, `otkup`, `partneri` or `agro` and then delegates feature rendering to the corresponding root renderer/sub-router.
- **canonical DOM rule:** management content now lives directly in canonical V2 mount zones (`mgmtDispecerMount`, `mgmtOtkupMount`, `mgmtPartneriMount`, `mgmtAgroMount`); runtime DOM transplant via `mgmtMountLegacyBlocks()` is no longer part of the active shell.
- **sub-router ownership:** `showMgmtOtkupSub`, `showMgmtPartnerSegment`, `showMgmtKoopSub`, `showMgmtKupSub` and `showMgmtAgroSub` are the canonical routers for nested management content.
- **boot helper ownership:** management prefetch and kupac-dropdown bootstrap helpers are owned by the V2 shell itself.
- **canonical management read helpers:** overview/dashboard reads are normalized around shared accessors for otkup date/quantity and shared `mgmtData` slice extraction, reducing parser drift between Management root screens.
- **observability rule:** Management shell catch paths now log explicit runtime errors instead of silently swallowing failures.
- **overview root contract:** `mgmtRenderOverview()` combines `mgmtData` plus live dispatch runtime (`dpDem`, `dpPlans`, `dpKamioni`, `dpGetSup`) into KPI cards, alert lists, finance summaries and quick links.
- **dashboard root contract:** `mgmtRenderDashboard()` is period-aware (`today`, `7d`, `season`) and drives KPI cards, chart series, dispatch highlights, alerts, finance blocks and quick-link rendering using the same canonical field/date pipeline as the overview screen.
- **charting rule:** dashboard chart rendering uses `Chart.js`, with bar mode for `today` station-level aggregation and line mode for multi-day period series.
- **nav synchronization invariant:** management root selection must stay aligned with `.tab-btn.role-management` and the role-nav engine, making V2 shell state the canonical source for active management navigation.
- **dead-code cleanup:** shell-local bottom-nav helpers `showMgmtBottomRoot`, `updateMgmtBottomNavActive` and `updateMgmtBottomNavVisibility` are removed from the active Management shell.
- **state/runtime cleanup dependency:** Management shell now sits inside a cleaner app-wide runtime contract where `entityID` is canonical session identity and role sync/state orchestration is less coupled to legacy Otkupac-only assumptions.

#### 6.3.5k Management `dispecer.js` Contract
The supplied dispatch module makes the active management planning board explicit:

- **runtime state:** dispatch owns `dpDem`, `dpPlans`, `dpSel`, `dpKS`, `dpKamioni` and persisted truck-capacity map `dpKap` in `localStorage`.
- **multi-source truck model:** `dpInit()` merges truck identities from stammdaten drivers, saved capacities, live `KamionStatus`, assigned otkupi and active plans into one operative truck list.
- **three-column board:** the module renders supply (`dpRS`), transport (`dpRTr`) and demand (`dpRD`) as the canonical live planning board.
- **supply semantics:** `dpGetSup()` means today’s unassigned otkupi; `dpGetAsg()` means today’s otkupi already carrying a `VozacID`.
- **tap-to-plan workflow:** planning is an explicit 3-step selection state — truck (`dpTK`) → station (`dpTS`) → demand (`dpTD`) — surfaced through the active banner `dpBN(...)`.
- **save-plan rule:** `dpOK()` converts the current three-step selection into `saveDispecer(...)`, appends a local `planned` row, recomputes route, sets truck status to `utovar` and fire-and-forget updates `updateKamionStatus`.
- **plan lifecycle:** `dpChgPlanSt()` and `dpRmPlan()` keep local `dpPlans` synchronized with server `updateDispecer` / `removeDispecer`, while also recalculating truck route/status after every change.
- **truck-status rule:** truck status is separate from plan rows; `dpCS(...)` can push manual status changes into `updateKamionStatus` even outside plan creation.
- **war-room demand write path:** `dpAD()` writes new demand through `saveWarRoomDemand`, updates local demand state and then refreshes demand/KPI panels.
- **KPI contract:** `dpRK()` computes waiting kg, known truck count, demand kg and active-plan count directly from runtime state, not from a separate report export.

#### 6.3.5l Management `kooperanti.js` Contract
The supplied kooperant-management module defines three active partner views:

- **station-driven drilldown:** `populateMgmtStanice()` and `onMgmtStanicaChange()` drive the kooperant card flow by station first, kooperant second.
- **fallback station rule:** when `stammdaten.stanice` is empty, the station dropdown can still be built from `Kooperant.StanicaID` values.
- **kartica view:** `onMgmtKooperantChange()` loads one kooperant card either from prefetched `mgmtData.kartice` or `getMgmtKartica`, filters out `UKUPNO`, recomputes summary totals and renders detailed rows.
- **saldo view:** `loadMgmtKoopSaldo()` reads only `Opis === 'UKUPNO'` rows from management kartice as the active saldo summary model.
- **pregled view:** `loadMgmtKoopPregled()` / `renderMgmtKoopPregled()` rely on `mgmtData.saldoOMDetail`, optionally filter by station, compute cross-record totals (`kg`, `Vrednost`, `Isplaceno`, `AgroZaduzenje`, `Saldo`, `Ambalaza`) and then render one summary row per kooperant/station pair.
- **read-model dependency:** the management kooperant pregled is export/report driven and depends on `SaldoOMDetail`, not on direct transactional reconstruction in the browser.

#### 6.3.5m Management `kupci.js` Contract
The supplied buyer-management module makes these buyer-side read rules explicit:

- **faktura source:** `loadMgmtFakture()` prefers prefetched `mgmtData.fakture` and falls back to `getMgmtFakture(kupacID)` only when the bundled data is unavailable.
- **expandable invoice drilldown:** `toggleFakturaStavke(...)` lazily renders faktura stavke from prefetched `mgmtData.fakturaStavke` or the dedicated `getMgmtFakturaStavke` endpoint.
- **saldo buyer view:** `loadMgmtKupci()` renders one saldo summary row per buyer from `mgmtData.saldoKupci`.
- **delivered-goods view:** `loadMgmtPredato()` groups `predatoPoKupcu` by buyer, aggregates `Kolicina`, `Ambalaza` and `Vrednost`, and then renders nested rows by `VrstaVoca` / `Klasa` / broj prijemnica.
- **read-only invariant:** buyer management is fully read-only in the PWA; faktura creation/payment mutation remains outside this module.

#### 6.3.5n Management `stanice.js` Contract
The supplied station-management module clarifies the active procurement/station screens:

- **otkup list source:** `loadMgmtOtkupi()` prefers `mgmtData.otkupiAll`, filtering by `OTK-<StanicaID>` or `OtkupacID`, and falls back to `getMgmtOtkupiByStanica` when needed.
- **date filter rule:** optional `Od` / `Do` UI filters are applied client-side after records are materialized.
- **station KPI rule:** count, total kg, total value and distinct kooperant count are recomputed client-side from the filtered station otkup set.
- **station saldo view:** `loadMgmtSaldoOM()` reads from `mgmtData.saldoOM` and displays `Saldo`, `Avans` and `Isplaceno` per station.
- **roba-by-station view:** `loadMgmtOtkupPoOM()` groups exported `otkupPoOM` rows by station and aggregates total kg, ambalaža and vrednost, then renders detailed fruit/class/broj otkupa subrows.
- **export dependency:** station management screens are read-model consumers of bundled/exported data rather than direct document reconstruction.

#### 6.3.5o Management `agrohemija.js` / Izdavanje Contract
The supplied management agro-issuing module makes the current issuing flow explicit:

- **module identity:** this is the active Agrohemija Izdavanje module described in recent implementation notes — barcode-assisted issuing, parcel-aware dosage, printable otpremnica and signature-backed completion.
- **working state:** issuing uses `izdKorpa`, selected kooperant/name and one module-scoped recommendation quantity `izdPreporukaQty`.
- **dropdown boot:** `populateIzdDropdowns()` hydrates kooperanti and artikli selectors directly from stammdaten and binds artikl changes to dosage recalculation.
- **parcel-aware dosage:** `onIzdKooperantChange()` loads kooperant parcels into a multi-select list; `izdCalcPreporuka()` computes `DozaPoHa × ukupna odabrana površina`, optionally using package rounding, and `izdPrimeniPreporuku()` copies that into quantity input.
- **scan surfaces:** kooperant identity can come from QR (`startIzdKoopScan()`), while artikli can be added by camera barcode scanning or manual dropdown selection.
- **cart rule:** `izdDodajUKorpu()` merges duplicate artikli lines by incrementing quantity/value rather than appending duplicates; `izdRenderKorpa()` is the canonical cart renderer.
- **document boundary:** `izdZavrsi()` does not save immediately; it first opens a printable/signable agro otpremnica modal via `izdShowOtpremnica(...)`.
- **signature rule:** the current issuance completion flow supports dual-signature/document-confirmation semantics through the otpremnica modal before the issue is treated as finished.
- **issuing-save rule:** the active save path persists one issuing header row through backend `saveIzdavanje`, with `stavke` sent as the cart snapshot and parcel IDs serialized as a comma-joined list.
- **PDF rule:** the same modal can generate a PDF using `jsPDF` A5-sized issuance output and upload it through backend `uploadPdf`, following the same pattern as the otkupni-list PDF flow.
- **state-reset rule:** `izdReset()` clears cart, selected kooperant, parcel selection, note and barcode debounce state after successful finish/cancel.


#### 6.3.5p Otkupac `otkup-form.js` Contract
The supplied otkup-form module makes the field-procurement entry contract explicit:

- **UI bootstrap contract:** `initOtkupFormUI(...)` hydrates fruit, package and kooperant dropdowns, applies defaults, optionally preserves active selections and then binds wizard events only once.
- **mobile step flow:** the module treats `tab-otkup` as a guided 5-step mobile flow — kooperant → parcela/vozač → roba → cena/ambalaža → napomena — with smooth scroll progression between major blocks on mobile viewports only.
- **manual fallback rule:** QR scan is not mandatory; `onManualKooperantChange()` can set the same canonical kooperant state and then reveal/populate parcela options.
- **parcel visibility rule:** `populateParcelaDropdown()` hides the parcela selector entirely when the selected kooperant has no known parcels; parcela remains optional even when shown.
- **fruit-driven defaults:** the active package default is hardcoded by fruit normalization (`višnja` / `šljiva` → `12/1`; everything else → `6/1`), while default price still comes from `stammdaten.config` (`Cena{Vrsta}`).
- **validation boundary:** `validateOtkupInput(...)` currently requires kooperant, fruit, quantity, price, package type and package count; sorta, parcela, napomena and vozač remain optional.
- **local record shape:** `buildOtkupRecord(...)` creates one canonical local otkup row in IDB with `entityType = "otkup"`, `schemaVersion = 1`, sync metadata fields, optional `vozacID`, and a client-generated UUID-like `clientRecordID`.
- **save behavior:** `saveOtkup()` is local-first — write to IndexedDB, open the otkupni-list modal immediately, reset the form, refresh local queue/badge/stats, and then opportunistically trigger queue sync when online.
- **error UX rule:** field-level validation is surfaced through `.otk-field-error` blocks plus `is-invalid` classes, not through a separate validation summary component.

#### 6.3.5q Otkupac `otkup-more.js` Contract
The supplied “Više” module defines the current otkupac profile/signature/queue diagnostics surface:

- **scope:** the tab combines profile display, one persisted otkupac signature pad and a queue diagnostics panel.
- **signature persistence rule:** otkupac signature is stored locally in `localStorage` under an entity-scoped key (`otkupac-signature:{ENTITY_ID}`), not in IndexedDB and not on the backend.
- **trim/export contract:** `exportTrimmedSignature(...)` crops the handwritten region, normalizes stroke pixels to opaque black and background to transparent/empty PNG before persistence.
- **reusability invariant:** the signature saved here is the canonical otkupac signature source reused later by `otkupni-list.js` for on-screen display and PDF output.
- **queue diagnostics model:** `renderOtkupacMoreSyncStats()` reads all local otkup rows from `CONFIG.STORE_NAME`, computes pending/error/synced counters, derives last successful sync stamp, and renders only unresolved/pending cards.
- **manual sync surface:** `syncOtkupacFromMore()` is a UI wrapper that delegates to `syncQueueSafe()` when available, then refreshes queue diagnostics after completion.
- **empty-state rule:** when no local unresolved rows exist, the queue panel explicitly states that there are no local items waiting for synchronization.

#### 6.3.5r Otkupac `otkup-pregled.js` Contract
The supplied “Danas / Pregled” module makes the otkup daily overview contract explicit:

- **dual-source read model:** `loadOtkupPregled()` merges local IndexedDB otkup rows with server `getOtkupi(otkupacID)` rows into one normalized view-model.
- **merge precedence rule:** unsynced local rows, rows with `lastSyncError`, or locally newer rows override server versions; deleted local rows are filtered out from the final overview.
- **quick-filter model:** the screen supports `danas`, `juce`, `sve`, `bez_vozaca`, `problemi` and ad-hoc custom date ranges through `fldPregledOd` / `fldPregledDo`.
- **problem semantics:** a row is considered a “problem” whenever sync state is not `synced` or a non-empty `lastSyncError` exists.
- **sectioning rule:** rendered cards are grouped into `Danas bez vozača`, `Danas sa vozačem` and `Sync problemi` based on current filtered rows.
- **KPI recomputation:** count, total kg, total value and distinct kooperant count are recalculated entirely client-side from the filtered overview rows.
- **detail modal contract:** clicking a pregled card opens one reusable modal with normalized field/value grid and a direct handoff into `showOtkupniList(...)` for the same row.
- **status vocabulary:** the overview currently uses only three visual states — `Bez vozača`, `Dodeljen`, `Sync problem`.

#### 6.3.5s Otkupac `otkupni-list.js` Contract
The supplied otkupni-list module defines the current signed procurement-slip flow:

- **modal-first document flow:** `showOtkupniList(record)` builds one full-screen modal from the otkup row plus seller config (`SELLER_*`, `OtkupPDVStopa`, `OtkupRokIsplate`) and kooperant master data.
- **calculation rule:** otkupni-list value is recomputed client-side as `količina × cena`, with optional PDV compensation added into the displayed/payment total.
- **signature split:** kooperant signs inside the modal canvas (`sigKooperant`), while otkupac signature is pulled from the Više-tab `localStorage` snapshot.
- **mandatory signature boundary:** `saveOtkupniListWithSignatures(...)` requires a kooperant signature before persisting it back onto the existing local otkup row as `sigKooperant` + `signedAt`.
- **PDF contract:** `savePdfToDrive(...)` renders an A5 `jsPDF` document, embeds both signatures when available, and uploads the final base64 PDF through backend action `uploadPdf`.
- **seller-master dependency:** header/company identity in both modal and PDF is fully config-driven (`SELLER_NAME`, address, PIB, MB, account), not hardcoded in the module.
- **cleanup rule:** closing the modal destroys the temporary signature pad instance via `destroySignaturePad('sigKooperant')`.

#### 6.3.5t Otkupac `otpremnice.js` Contract
The supplied `otpremnice` module clarifies that the current PWA transport flow is a local driver-assignment layer over existing otkup rows:

- **UI-first scope:** the tab is rendered as three view states — root overview, assign-to-driver flow and success summary.
- **data model:** `loadOtpremaOverview()` merges local and server otkup rows exactly like pregled, but reinterprets them through transport-specific grouping (`today unassigned`, `older unassigned`, `today assigned`).
- **driver assignment semantics:** selecting a driver and confirming otprema does **not** create a separate otpremnica document in the browser; it mutates the existing otkup rows by writing `vozacID` / `vozacName` and resetting sync metadata to `pending`.
- **driver source rule:** driver can be chosen by QR (`VOZ-*` or JSON payload) or by fallback dropdown from `stammdaten.vozaci`.
- **selection model:** assignment uses one in-memory `selectedKeys` set keyed by canonical record key (`serverRecordID` else `clientRecordID`).
- **success transition:** `confirmOtpremaAssign()` persists the mutated rows locally, opens a dedicated success screen, updates badge state and opportunistically triggers queue sync when online.
- **assigned-date rule:** “Danas otpremljeno” is derived from assignment/update timestamps (`updatedAtClient` / `updatedAtServer` / `syncedAt`), not from a separate transport document date.
- **detail modal:** both assigned and unassigned cards reuse one shared otprema detail modal showing fruit, class, kg, ambalaža, driver and sync badges.
- **important architectural note:** despite the tab name, this module currently represents **driver assignment over otkup rows**, not canonical document-level `Otpremnica` creation.

#### 6.3.5u Otkupac `sync.js` Contract
The supplied otkupac sync module narrows the active procurement queue synchronization contract:

- **scope:** the file currently syncs only `CONFIG.STORE_NAME` otkup rows through backend action `sync`.
- **runtime gate:** sync aborts when `db` is missing, device is offline, or `window.appRuntime.syncInFlight` is already active.
- **state transition rule:** pending rows move `pending → syncing → synced` on successful confirmation and fall back to `pending` with diagnostic metadata on request failure, exception, or missing result.
- **result application:** successful backend rows update `serverRecordID`, `updatedAtServer`, `syncedAt`, `lastServerStatus` and clear `lastSyncError`; unsuccessful rows retain local ownership and diagnostics.
- **missing-result rule:** when a pending local row is absent from backend `json.results`, the module treats that as a failed confirmation (`Nema potvrde sa servera`) and returns the row to `pending`.
- **legacy fallback:** older backend responses shaped only as `{ success: true }` still mark all pending otkup rows as synced.
- **badge contract:** `updateSyncBadge()` derives the header badge from both connectivity and local queue state, with visible states `OFFLINE`, `SYNC...`, `ČEKA: n`, and `ONLINE`.
- **stats contract:** `updateStats()` computes only today’s pending/synced counts from local IndexedDB rows; it is a local operational signal, not a server truth source.
- **queue list contract:** `renderQueueList()` renders only `pending` and `syncing` rows, including inline server-error diagnostics when present.

#### 6.3.5v Vozac `zbirna.js` Contract
The supplied vozac zbirna module makes the driver-side aggregation flow explicit:

- **source model:** `loadVozacData()` first loads assigned otkupi from backend action `getVozacOtkupi`, then separately loads merged zbirne through `getMergedZbirneForVozac()`.
- **consumption rule:** otkup rows already referenced inside any existing zbirna via `otkupRecordIDs` are removed from the active vozac otpremnice pool, so the same otkup cannot be aggregated twice.
- **today-first UX:** the main zbirna screen shows only today’s still-unconsumed assigned otkupi grouped by stanica, while already-created zbirne are rendered in a separate historical list.
- **create flow:** `startZbirnaCreation()` opens a dedicated create view, populates kupac choices from `stammdaten.kupci` with `mgmtData.saldoKupci` fallback, and summarizes today’s available otkupi before confirmation.
- **aggregation rule:** `confirmZbirna()` creates one local `zbirne` record by summing class I / II kilograms, total ambalaža and concatenated `otkupRecordIDs` from today’s still-free otkupi.
- **local record contract:** new zbirne rows are stored in IndexedDB store `zbirne` with `entityType = "zbirna"`, `schemaVersion = 1`, sync metadata, a client-generated UUID-like `clientRecordID`, technical `serverRecordID` and separate business field `brojZbirne`.
- **merge rule:** local and server zbirne are reconciled through `mergeZbirneRecords()` which delegates to the generic `mergeOfflineRecords(...)` helper using `normalizeLocalZbirnaRecord(...)`.
- **sync engine:** `syncZbirne()` mirrors the same pending → syncing → synced lifecycle pattern as other entity sync flows but targets backend action `syncZbirna` and store `zbirne`.
- **duplication safety:** backend statuses `duplicate`, `existing`, `inserted`, `updated` and `synced` are all treated as successful confirmations for local zbirna rows.
- **business ID display rule:** `mapServerZbirnaRecord(...)`, `normalizeLocalZbirnaRecord(...)`, `confirmZbirna()` and `syncZbirne()` now preserve `brojZbirne` separately from `serverRecordID`, and driver UI renders `brojZbirne` before any technical server identifier.

#### 6.3.5w Vozac `transport.js` Contract
The supplied transport overview module clarifies the current driver transport read model:

- **dual-source read model:** `loadVozacTransport()` merges local IndexedDB store `zbirne` with backend action `getVozacZbirne` into one transport list.
- **normalization rule:** server zbirna rows are normalized into the same local/client shape used by the vozac queue, including sync metadata fields and aggregated class totals.
- **merge primitive:** transport overview uses `mergeTransportZbirne(local, server)` which is currently a thin specialization over `mergeOfflineRecords(local, server, normalizeLocalZbirnaRecord)`.
- **display semantics:** each transport row shows kupac/hladnjača, datum, total kilograms, ambalaža count, sync icon and optional `brojZbirne` as the preferred business-visible identifier, with `serverRecordID` only as technical fallback / diagnostics plus `lastSyncError` when relevant.
- **sync icon vocabulary:** `🔄` means `syncing`, `⏳` means `pending`, and `✅` means a clean synced/server-confirmed zbirna snapshot.
- **empty-state rule:** when neither local nor server zbirne exist, the tab explicitly renders `Nema transporta` instead of a blank screen.

#### 6.3.6 Render Model
- card-based UI
- bottom nav per role
- sub-navigation within large features
- modal overlays for detail views
- toast messaging
- fixed mobile nav and hidden desktop nav/sidebar hybrid
- sanitized dynamic rendering
- sticky action bars for mobile-heavy workflows such as otkup save, otprema confirm and agro timer flow
- shell-level modals declared in HTML and populated by feature scripts rather than injected ad hoc

#### 6.3.7 Sync Engine
- local-first writes to IDB
- action-based POST sync
- status updates per record
- conflict merge through `mergeOfflineRecords`
- connectivity-aware retries


#### 6.3.7a Kooperant `sync.js` Contract
The supplied kooperant sync module further narrows the active offline-first synchronization contract:

- **scope:** the kooperant sync surface owns both the treatment sync path for `tretmani` (`syncTretmani`) and the farm-expense sync path for `troskovi` (`syncTroskovi`), with `syncKooperantNow()` aggregating both results.
- **generic sync core:** `syncEntityStore(...)` is the active generic implementation and parameterizes `storeName`, backend `action`, runtime in-flight key and success label.
- **runtime gate:** sync aborts early when `db` is not ready, the device is offline or the same entity sync is already running; in-flight flags live under `window.appRuntime.kooperantSync`.
- **state transition rule:** pending rows move `pending → syncing → synced` on success, and revert to `pending` with diagnostic metadata on failure/exception.
- **attempt metadata:** each processed row tracks `syncAttemptAt`, `syncAttempts`, `lastSyncError`, `lastServerStatus`, and where returned by the backend also `serverRecordID`, `updatedAtServer` and `syncedAt`.
- **success interpretation:** the client treats backend results with `success = true` or statuses such as `synced`, `duplicate`, `existing`, `inserted` or `updated` as successful terminal confirmations.
- **missing-result rule:** if the backend omits a pending `clientRecordID` from `json.results`, that row is returned to `pending` with `lastServerStatus = "missing-result"`.
- **legacy fallback:** older backend responses shaped only as `{ success: true }` still mark all pending rows as synced through a `legacy-success` path.
- **manual sync surface:** `syncKooperantNow()` returns the canonical role-level sync object and includes child results for both `tretmani` and `troskovi`; `reason: "no-pending"` is valid when both stores have no pending rows.

#### 6.3.8 Service Worker Strategy
- cache names are explicit release versions and must be bumped when cached JS/CSS assets change; the v6.12 smoke branch used successive cache bumps through the PWA launch-hardening patches
- API calls bypass SW interception
- HTML uses network-first with cache fallback
- assets use cache-first with network fallback
- active precache includes the modular JS/CSS tree plus the currently required remote libraries used by the shell, including Leaflet and Chart.js
- install path no longer depends on a fail-all `cache.addAll(...)` pattern; asset caching is hardened through per-asset handling and cache-version updates
- offline map limitation remains where third-party tile/dependency coverage is still runtime-dependent
- service-worker completion alone is not full production readiness; manifest polish, update UX, cache pruning and explicit offline status surfaces remain open work

#### 6.3.7b Otkupac/Vozac Sync Convergence Note
The newly supplied otkupac and vozac sync snapshots make one broader client-sync pattern explicit:

- **per-entity store isolation:** otkup rows sync from `CONFIG.STORE_NAME`, kooperant agronomy entities sync from dedicated stores, and vozac zbirne sync from `zbirne`.
- **shared lifecycle shape:** all current entity sync flows use the same core metadata vocabulary — `syncStatus`, `syncAttemptAt`, `syncAttempts`, `lastSyncError`, `lastServerStatus`, `updatedAtServer`, `syncedAt`.
- **local authority during transit:** while a row is `pending` or `syncing`, the client intentionally preserves local authority in merged overview screens.
- **finally-refresh invariant:** sync flows are expected to refresh badges/lists after completion or failure so operational UI never stays on stale queue state.

#### 6.3.8a `sw.js` Offline-App-Shell Contract
The supplied service worker snapshot makes the current offline shell contract more concrete:

- **versioned cache contract:** the active shell cache is currently named `AgriX-v7`; activation deletes all older named caches.
- **precache scope:** the install phase precaches `index.html`, manifest, all local CSS/JS feature modules, icons, plus the currently required remote runtime libraries (`html5-qrcode`, `jsPDF`, `Leaflet`, `Chart.js`).
- **API bypass rule:** requests to `script.google.com` are intentionally not intercepted by the service worker, so business actions always go directly to the GAS backend.
- **HTML freshness rule:** documents use network-first with cache fallback, which keeps the shell update-friendly while still allowing offline startup.
- **asset rule:** non-document assets use cache-first with network fill-back, effectively making the app shell and local feature scripts offline-first.
- **root fallback:** when a document request fails, the worker falls back to cached `./index.html`.
- **operational caveat:** the current precache list covers local app modules and selected remote libs, but still depends on runtime network availability for some third-party assets not explicitly cached by the worker.
- **remaining readiness gap:** a user-visible offline banner, explicit update-available/apply-update mechanism, IndexedDB pruning policy and manifest hardening are still outside the completed SW contract.

### 6.4 Shared Utilities
| Utility | Purpose | Used by | Invariant |
|---|---|---|---|
| `apiFetch` / `apiPost` | wrapped API transport with auth/error handling | all PWA features | no raw `fetch()` for app actions |
| `escapeHtml` | output sanitization | all render paths | mandatory before `innerHTML` |
| `safeAsync` | async guard and consistent error toast behavior | feature modules | standard async wrapper |
| format helpers | localized dates and display formatting | all UIs | Serbian locale consistency |
| storage helpers | local session values only | auth/signature/device helpers | not for shared state |
| `mergeOfflineRecords` | conflict-aware local/server merge | queue, list views | pending/error local wins |

---


### 6.1.15 `modGoogleAuth` Active Business Capabilities

The current canonical desktop Google auth bridge exposes the following active capabilities:

- **one-time OAuth bootstrap:** `RunGoogleAuthSetup()` opens the Google consent URL in the browser, collects the authorization code through a desktop prompt, and exchanges it for access/refresh tokens.
- **token persistence contract:** the active write path stores `GOOGLE_ACCESS_TOKEN`, `GOOGLE_REFRESH_TOKEN` and `GOOGLE_TOKEN_EXPIRES_AT` via `SetConfigValue()` into `tblSEFConfig`, even though module comments describe them as config keys more generally.
- **auto-refresh behavior:** `GetAccessToken()` checks expiry with a 60-second buffer and automatically calls `RefreshAccessToken()` before any Google API call when required.
- **configured-auth gate:** `IsGoogleAuthConfigured()` treats a present refresh token as the minimum readiness signal for desktop↔Google integrations.
- **desktop Drive visibility rule:** desktop Google auth now requires a Drive scope broad enough to enumerate GAS-created spreadsheets in the shared master folder; using `drive.file` alone is insufficient for canonical master-sync listing of GAS-created `VOZ-*` files.
- **current scope requirement:** the active VBA Google integration scope is `https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive`, and changing scope requires a fresh OAuth consent/token bootstrap.
- **transport implementation:** token exchange and refresh use `WinHttp.WinHttpRequest.5.1` with explicit timeout settings and form-encoded POST bodies to the Google token endpoint.
- **helper surface:** the module also owns URL encoding, lightweight JSON value extraction and ISO-like expiry timestamp calculation used by downstream Google wrappers.

### 6.1.16 `modGoogleSheets` Active Business Capabilities

The current canonical desktop Google Sheets wrapper exposes the following active capabilities:

- **tab overwrite write path:** `WriteSheetData()` clears the target tab first and then writes a complete 1-based 2D array back starting from `A1`.
- **array serialization rule:** `BuildValuesJson()` serializes all outgoing cells as strings, with dates formatted as `yyyy-mm-dd`, so Google Sheets can infer values while preserving leading-zero identifiers.
- **read path:** `ReadSheetData()` retrieves a tab through the Sheets values API and parses the response back into a 1-based 2D VBA array via `ParseValuesJson()`.
- **spreadsheet provisioning:** `CreateSpreadsheet()`, `GetSpreadsheetID()` and `AddSheetTab()` are the approved provisioning/search helpers for desktop-created Google workbooks and tabs.
- **Drive placement helper:** `MoveFileToFolder()` relocates a newly created spreadsheet into the configured Google Drive folder using the Drive API.
- **auth dependency:** every public Google Sheets operation is gated through `modGoogleAuth.GetAccessToken()` and fails closed when no valid token can be obtained.

### 6.1.17 `modMasterSync` Active Business Capabilities

The current canonical desktop master-sync layer exposes the following active capabilities:

- **OTK sheet discovery:** `FindOTKSheets()` searches the configured PWA folder for Google spreadsheets whose names start with `OTK-` and returns ID/name pairs for import processing.
- **OTK master import gate:** `ImportOtkupFromPWA()` imports only rows whose Google-side `SyncStatus` equals `Synced`; all other rows are skipped as already handled or errored.
- **OTK transaction scope:** `ImportOtkupFromPWA_TX()` snapshots `tblOtkup` and `tblAmbalaza` before bulk import so that desktop-side master ingestion can be rolled back.
- **OTK row validation:** `ValidatePWAOtkup()` enforces presence of `KooperantID`, `VrstaVoca`, positive `Kolicina` and positive `Cena`, and verifies that the referenced kooperant exists in desktop master data.
- **OTK duplicate semantics:** `IsDuplicateInMaster()` treats `ClientRecordID` as the canonical deduplication key for imported field otkup rows; duplicate rows may still update missing `VozacID` through a dedicated repair path when applicable.
- **OTK writeback semantics:** successful imports are written back to the Google sheet as `Synced>Master`, while invalid or duplicate rows get `SyncError[:reason]` or `Duplicate`.
- **OTK sheet bootstrap:** `CreateOTKSheetsForAllStanice()` provisions one `OTK-{StanicaID}` spreadsheet per active station in the configured Google Drive folder and writes the current GAS-aligned 22-column header contract for field-side usage.
- **post-import transport assist:** `AutoCreateOtpremniceFromPWA()` groups imported otkupi without `OtpremnicaID` by `StanicaID + Datum + VozacID + Klasa`, creates one aggregated otpremnica per unique group, and back-links all member otkup rows to the new `OtpremnicaID`.
- **otpremnica numbering rule:** auto-created transport documents use the format `{StanicaNum}/{DDMM}-{seq}` where the sequence counter is derived from existing active otpremnice per station/date.
- **VOZ sheet discovery:** `FindVOZSheets()` searches the configured PWA folder for Google spreadsheets whose names start with `VOZ-` and returns ID/name pairs for import processing.
- **VOZ master import gate:** `ImportZbirneFromPWA()` imports only rows whose Google-side `SyncStatus` equals `Synced`; all other rows are skipped as already handled or errored.
- **VOZ transaction scope:** `ImportZbirneFromPWA_TX()` snapshots `tblZbirna`, `tblOtpremnica` and `tblOtkup` before bulk import so that desktop-side zbirna ingestion can be rolled back locally.
- **VOZ row validation:** `ValidatePWAZbirna()` enforces presence of `VozacID` and `KupacID`, verifies that the referenced kupac exists in desktop master data, and requires at least one positive class quantity.
- **VOZ duplicate semantics:** `IsDuplicateZbirnaInMaster()` treats `ClientRecordID` as the canonical deduplication key for imported zbirna rows.
- **BrojZbirne generation rule:** imported zbirna rows generate `BrojZbirne` inside desktop master using the numeric part of `VozacID` without leading zeros, `/`, `ddmmyy` date and optional `-2`, `-3`, ... daily sequence suffix per vozač.
- **Zbirna write semantics:** imported VOZ rows write `tblZbirna` in the active 16-column order `ZbirnaID | Datum | VozacID | BrojZbirne | KupacID | Hladnjaca | Pogon | VrstaVoca | SortaVoca | UkupnoKolicina | TipAmbalaze | UkupnoAmbalaze | Klasa | Stornirano | ClientRecordID | SyncSource`, with `SyncSource = "PWA"`.
- **cascade-link semantics:** after successful zbirna import, `LinkZbirnaToOtkupAndOtpremnica()` propagates generated `BrojZbirne` to matching `tblOtkup` and `tblOtpremnica` rows using comma-separated `OtkupRecordIDs` from the VOZ sheet.
- **VOZ writeback semantics:** successful zbirna imports are written back as `Synced>Master` through `Sheet1!F`, while `Sheet1!B` receives the canonical desktop `ZbirnaID`; invalid or duplicate rows get `SyncError[:reason]` or `Duplicate`.
- **rollback boundary:** `_TX` wrappers around master-sync protect local Excel tables only; Google Sheets writeback remains an external side effect and is not reverted by `clsTransaction` rollback.

### 6.1.18 `modStammdatenSync` Active Business Capabilities

The current canonical desktop export/sync layer for Google-facing read models exposes the following active capabilities:

- **stammdaten workbook lifecycle:** `SyncStammdatenToGoogle()` finds or creates the `Stammdaten` workbook inside the configured PWA folder, persists its spreadsheet ID, provisions missing tabs and exports 13 active tabs in one run.
- **export families:** active export functions include at minimum `ExportKooperanti`, `ExportKulture`, `ExportParcele`, `ExportConfig`, `ExportUsers`, `ExportFakture`, `ExportFakturaStavke`, `ExportSaldoOMDetail`, `ExportStanice`, `ExportKupci`, `ExportVozaci`, `ExportArtikli` and `ExportMagacinKoop`.
- **kartice workbook export:** `ExportKarticeToGoogle()` creates or reuses a dedicated `Kartice` workbook and writes combined per-kooperant financial-card rows generated from `ReportKarticaKooperanta()`.
- **management workbook export:** `ExportMgmtReports()` maintains a separate `MgmtReports` workbook and refreshes `SaldoOM`, `SaldoKupci`, `OtkupPoOM` and `PredatoPoKupcu` tabs.
- **derived aggregation semantics:** management exports aggregate desktop canonical data by station, buyer, kooperant, fruit type and class, and `ExportSaldoOMDetail()` explicitly combines otkup value, kooperant payouts and agro-warehouse debt into one OM detail balance read model.
- **auth/config gate:** all exports require valid Google OAuth setup plus a configured `GOOGLE_PWA_FOLDER_ID`.

### 6.1.19 Desktop Parcel Geo Point Helper Capabilities

The current canonical desktop parcel geo point helper exposes the following active capabilities:

- **UTM34→WGS84 conversion:** `ConvertUTM34ToLatLng()` converts UTM zone 34 coordinates into decimal latitude/longitude using the active ellipsoid constants embedded in the helper.
- **point save path:** `SaveParcelGeoPoint(rowIndex, nCoord, eCoord)` stores raw `N_Coord`/`E_Coord`, computed `Lat`/`Lng`, rounds coordinates to 6 decimals, sets `GeoStatus = "point"`, `GeoSource = "selenium"`, enables `MeteoEnabled = "Da"`, and stamps both geo-entered and updated timestamps.
- **point clear path:** `ClearParcelGeo(rowIndex)` clears all coordinate fields, resets `GeoStatus = "none"`, clears `GeoSource`, disables meteo (`"Ne"`), and updates the parcel modification timestamp.
- **table authority:** these helpers write directly to `tblParcele` through `UpdateCell()` and therefore operate inside the same desktop master-data authority as the rest of parcela management.

### 6.1.20 SEF HTTP Transport Capabilities

The current canonical desktop SEF HTTP transport layer exposes the following active capabilities:

- **submit path:** `SubmitUBLInvoice(ublXml, requestId)` sends UBL XML to `/api/publicApi/sales-invoice/ubl?requestId=...` using `WinHttp.WinHttpRequest.5.1`, `ApiKey` auth and optional `X-SEF-ENV`.
- **status path:** `GetInvoiceStatus(sefDocumentId)` queries `/api/publicApi/sales-invoice?invoiceId=...` and parses the latest external invoice status snapshot.
- **cancel/storno paths:** `CancelInvoiceOnSEF()` and `StornoInvoiceOnSEF()` post JSON bodies to dedicated `/cancel` and `/storno` endpoints and return normalized `clsSEFResponse` objects.
- **response normalization:** submit parsing treats HTTP `200/201/202` as successful send, `400/409/422` as business rejection, `429` as rate-limit and all other non-success cases as failure; status parsing preserves exact external values such as `SENT`, `NEW`, `DRAFT`, `ACCEPTED`, `REJECTED`, `CANCELLED`, `STORNO` and `ERROR`.
- **config dependency:** every SEF HTTP call requires `SEF_BASE_URL`, `SEF_API_KEY` and optionally `SEF_ENV` from `tblSEFConfig`.
- **debug/audit surface:** submit transport emits request/response debug markers and wraps failures into normalized HTTP-error style `clsSEFResponse` objects rather than leaving callers with raw transport exceptions.

### 6.1.21 SEF DTO / Payload Capabilities

The current canonical desktop SEF payload layer exposes the following active capabilities:

- **snapshot build:** `BuildSEFInvoiceDto(fakturaID)` loads invoice header data from `tblFakture`, buyer identity from `tblKupci`, seller identity from config and line-level delivery/invoice detail from `tblFakturaStavke` plus linked `tblPrijemnica`.
- **line semantics:** each outbound SEF line is derived from prijemnica-linked fruit context and carries `PrijemnicaID`, `BrojPrijemnice`, description, quantity, price, class, net, VAT and gross amount.
- **amount model:** current active logic computes line net as `Kolicina × Cena`, line VAT through `GetDefaultTaxPercent()`, line gross as `net + VAT`, and faktura totals as `TotalNet`, `TotalVat`, `TotalGross`; hardcoded VAT rates are not canonical.
- **dual serializer surface:** `SerializeSEFRequest()` produces a JSON snapshot primarily for debugging/inspection, while `SerializeUBLInvoice()` is the active external outbound format.
- **UBL enrichment:** UBL generation injects buyer/seller postal/tax/payment metadata from buyer master data and config keys, writes `IssueDate` from `dto.InvoiceDate`, writes `ActualDeliveryDate` from `dto.DeliveryDate`, defaults country to `RS`, defaults payment means code to `30`, and applies configurable due-days and note text.
- **payload identity:** `ComputePayloadHash()` generates the persisted outbound payload hash used by submission journaling and technical-retry reuse decisions.
- **delivery-date derivation:** `GetInvoiceDeliveryDate(fakturaID)` scans `tblFakturaStavke` rows for the exact faktura and resolves each `PrijemnicaID` to `tblPrijemnica.Datum`; the latest linked prijemnica date becomes `dto.DeliveryDate`.
- **local date guard:** `ValidateSEFDtoForUBL()` rejects payloads where `dto.DeliveryDate > dto.InvoiceDate` before `SubmitUBLInvoice` can be called.

### 6.1.22 `modSEFPersistence` Active Business Capabilities

The current canonical desktop SEF persistence layer exposes the following active capabilities:

- **read helpers:** `GetFakturaSEFWorkflowState()`, `GetFakturaSEFDocumentId()`, `GetLastSEFSubmissionID()`, `GetNextSEFVersionNo()` and `GetCurrentSEFVersionNo()` are the approved lookup surface for live SEF state on fakture.
- **state update helper:** `UpdateFakturaSEFState_Row()` validates the requested local transition, updates workflow/state fields on `tblFakture`, optionally writes `SEFStatus`, `SEFDocumentId`, payload hash, submission ID, version number and last-error fields, and maintains `PoslatNaSEF`, `SEFSentAt` and `SEFLastSyncAt` side effects where applicable.
- **refresh-only helper:** `UpdateFakturaSEFRefreshFields_Row()` updates `SEFStatus`, `SEFDocumentId` and last-error fields without performing a local workflow-state transition.
- **submission journaling:** `CreateSEFSubmission_Row()` allocates `SFS-*` rows and stores request format/body, payload hash, workflow state at submit and operator identity; `SaveSEFSubmissionResult_Row()` records HTTP/API outcome back into the same submission row.
- **event log journaling:** `AppendSEFEvent_Row()` appends operator-timestamped SEF lifecycle events into `tblSEFEventLog`.
- **lookup/report helpers:** the module exposes retrieval helpers for submissions/events per faktura, successful-submission detection, last-submission status lookup, request-body/payload-hash reuse and explicit clearing of a faktura’s last submission pointer.

### 6.1.23 SEF Outbound Orchestration Capabilities

The current canonical desktop SEF outbound orchestration layer exposes the following active capabilities:

- **three-phase submit orchestration:** `SendInvoiceToSEF_TX(fakturaID)` performs local preparation in one TX, transport-state transition plus submission-row creation in a second TX, the remote HTTP call outside TX, and final result persistence/state transition in a third TX.
- **retry semantics:** `ShouldReuseLastSubmission()` allows technical-failure retry to reuse the previous submission body and payload hash instead of generating a fresh submission, keeping request identity and payload continuity explicit.
- **request identity rule:** the active outbound submit always uses `requestId = submissionID`.
- **final local outcomes:** the module maps remote responses to `WF_SEF_ACCEPTED`, `WF_SEF_SENT`, `WF_SEF_REJECTED` or `WF_SEF_TECH_FAILED` and appends explicit SEF event-log entries for both HTTP send and HTTP response stages.
- **live submit baseline:** v6.6 live smoke confirms that a valid dummy faktura can move from `LOCAL_FINALIZED` to `SEF_SENT`, persist HTTP 200 / `SubmissionStatus=SENT`, store `SEFDocumentId` and survive repeated refresh without state corruption.
- **remote corrective actions:** `CancelInvoiceOnSEF_TX()` and `StornoInvoiceOnSEF_TX()` call dedicated SEF endpoints and then update faktura refresh fields plus event log without pretending these remote actions are ordinary submit transitions.
- **stuck-send recovery:** `RecoverStuckSEFSendingInvoice()` and `RecoverAllStuckSEFSendingInvoices()` provide dedicated repair entry points for invoices stranded in `WF_SEF_SENDING`.

### 6.1.24 SEF Status Refresh Capabilities

The current canonical desktop SEF refresh layer exposes the following active capabilities:

- **single-invoice refresh:** `RefreshSEFStatus_TX(fakturaID)` queries SEF by `SEFDocumentId`, records the exact returned external status and updates the local workflow only when the state machine genuinely changes.
- **dual-status model enforcement:** the refresh logic explicitly allows `SEFWorkflowState` and `SEFStatus` to differ; for example, a faktura may remain locally `WF_SEF_SENT` while externally being `SENT`, `NEW`, `DRAFT`, `STORNO` or `CANCELLED`.
- **batch refresh:** `RefreshPendingOutboundInvoices_TX()` iterates all fakture in `WF_SEF_SENT` or `WF_SEF_SYNC_ERROR`, refreshes them one by one and inserts a fixed 2-second pacing delay between calls.
- **sync-error handling:** failed refresh attempts move fakture into `WF_SEF_SYNC_ERROR` while preserving remote identifiers and last-error details for operator review.
- **accepted/rejected convergence:** positive refreshes converge the local state to `WF_SEF_ACCEPTED` or `WF_SEF_REJECTED` when SEF returns a final external status.

### 6.1.25 SEF Validation and State-Machine Capabilities

The current canonical desktop SEF validation/state-machine layer exposes the following active capabilities:

- **allowed-transition gate:** `ValidateAllowedTransition(oldState, newState)` is the approved guard for local workflow transitions and explicitly enumerates legal moves across draft, finalized, ready, sending, sent, accepted, rejected, technical-failed, sync-error and storno states.
- **sendability validation:** `ValidateFakturaForSEF()` checks faktura existence, required buyer/header fields, numeric/non-zero total, allowed source states, absence of an already successful SEF submission, presence of invoice lines, buyer readiness and core SEF config readiness.
- **payload validation:** `ValidateSEFPayload()` rejects empty payloads and payloads without an invoice identifier marker.
- **action gating:** `ValidateFakturaCanBeCancelledOnSEF()` allows cancel only in external statuses `DRAFT`, `NEW` or `ERROR`; `ValidateFakturaCanBeStorniranoOnSEF()` allows storno only in `SENT`, `ACCEPTED` or `REJECTED`.
- **rejected correction flow:** `PrepareRejectedInvoiceForResubmit()` moves a rejected faktura back to `WF_SEF_READY`, clears last submission linkage and appends an explicit recovery event.
- **status helpers:** `IsFinalSEFStatus()`, `IsPendingSEFStatus()` and `GetSEFDisplayStatus()` provide canonical UI/report helpers over the dual local/external SEF status model.

## 7. API and Endpoint Reference

Za **svaki aktivni endpoint / action** navesti:
- method
- auth
- caller
- request
- response
- side effects
- failure modes

| Action / Endpoint | Method | Auth | Called by | Writes | Response | Failure modes |
|---|---|---|---|---|---|---|
| `login` | POST | None | PWA auth | session/token state | token + role + entity context | invalid PIN, expired deploy mismatch |
| `getStammdaten` | GET | Token | bootstrap/all roles | none | 11 tabs + `meteoLatest` + `kartice` (+ other exported fields) | stale deploy, token failure |
| `getOtkupi` | GET | Token + entity scope | Otkupac / management | none | station otkup records | auth/scope mismatch |
| `getKartica` | GET | Token + entity | Kooperant | none | kartica lines/summary | cache or auth error |
| `getTretmani` | GET | Token + entity | Kooperant | none | treatment records | auth/data mismatch |
| `getTroskovi` | GET | Token + entity | Kooperant | none | cost records | auth/data mismatch |
| `getOprema` | GET | Token + entity | Kooperant | none | equipment presets/list | auth/data mismatch |
| `getParcelGeo` | GET / POST-public | None | Kooperant GIS | none | point/polygon geo | unauthenticated public exposure risk |
| `getParcelMeteo` | GET / POST-public | None | Kooperant GIS fallback | none | current meteo + risk | external API / upstream failure |
| `getVozacOtkupi` | GET | Token + vozac | Vozac | none | assigned otkupi | auth mismatch |
| `getVozacZbirne` | GET | Token + vozac | Vozac | none | created zbirne | auth mismatch |
| `getMgmtAll` | GET | Token + management | Management bootstrap | none | bundled management data | payload size / auth issues |
| `getDispecer` | GET | Token + management | Management dispatch | none | supply + demand + plans | auth/state mismatch |
| `getKamionStatus` | GET | Token + management | Management dispatch | none | truck status list | stale status/deploy |
| `getKooperantProizvodnja` | GET | Token + entity | Knjiga Polja | none | parsed production from Kartice | parse/layout mismatch |
| `getFiskalni` | GET | Token + entity | Kooperant fiskalni | none | fiskalni line items | auth/data mismatch |
| `sync` | POST | Token | Otkupac | OTK sheets | sync status + duplicate-aware result | partial row failure, duplicate logic |
| `syncTretmani` / `syncAgromere` | POST | Token | Kooperant | treatment sheet/store target | sync result | validation / auth error |
| `syncTrosak` | POST | Token + role/entity ownership | Kooperant / Management | `records[]`, `kooperantID` | batch `{ success, processed, succeeded, failed, results }` | active expense sync; idempotent by `ClientRecordID` |
| `syncOprema` | POST | Token | Kooperant | equipment store/sheet | sync result | validation / auth error |
| `syncZbirna` | POST | Token | Vozac | VOZ sheets | sync result | auth/validation error |
| `parseFiskalni` | POST | Token | Kooperant fiskalni | none directly | parsed receipt/items | invalid QR/journal parse/upstream failure |
| `parseFiskalniImage` | POST | Token | Kooperant fiskalni | temp Drive file | parsed receipt/items | QR decode failure, image failure |
| `saveFiskalni` | POST | Token | Kooperant fiskalni | FISKALNI-KOOP sheet | save result | duplicate/save error |
| `saveFiskalniMapiranje` | POST | Token | Kooperant fiskalni | FiskalniMapiranje | save ack | mapping conflict |
| `createArtikal` | POST | Token | fiscal/manual flows | Artikli sheet | created artikal record | misuse can pollute master artikli |
| `uploadPdf` | POST | Token | Otkupni List / PDF flows | Google Drive | file url/id | Drive/auth failure |
| `saveParcelPolygon` | POST | No auth documented | geo editor | parcel geo/polygon store | save ack | auth gap, invalid geometry |
| `saveDispecer` | POST | Token + management | dispatch | DispecerPlan | created plan | auth/data conflict |
| `updateDispecer` | POST | Token + management | dispatch | DispecerPlan | updated plan | stale plan / auth error |
| `removeDispecer` | POST | Token + management | dispatch | DispecerPlan | removal ack | stale plan / auth error |
| `saveWarRoomDemand` | POST | Token + management | demand UI | demand tab | created demand | legacy naming drift |
| `removeWarRoomDemand` | POST | Token + management | demand UI | demand tab | removal ack | legacy naming drift |
| `updateKamionStatus` | POST | Token | dispatch / truck updates | KamionStatus | updated status | stale or invalid status |
| `scheduledMeteoFetch` | Trigger | Scheduled | GAS time triggers | MeteoLatest, MeteoHistory | scheduled execution | rate limit / upstream API |

### 7.0.1 Current Router Concretization from Supplied `Code.gs`

The newly supplied GAS backend narrows the currently confirmed action surface to the following concrete router actions.
This subsection is the authoritative action list confirmed by the pasted backend snapshot, even where older architectural notes mention broader or legacy action names.

**Confirmed `doPost` actions:**
- `login`
- `getParcelGeo`
- `getParcelMeteo`
- `getParcelMeteoLatest`
- `getAllMeteoLatest`
- `saveParcelPolygon`
- `sync`
- `syncAgromere`
- `syncZbirna`
- `saveOtkupniListPdf` — disabled / `FEATURE_DISABLED`
- `uploadPdf`
- `saveWarRoomDemand`
- `removeWarRoomDemand`
- `updateDemandPrimljeno`
- `updateKamionStatus`
- `saveDispecer`
- `updateDispecer`
- `removeDispecer`
- `saveIzdavanje`
- `syncTretman`
- `syncTrosak` — active Kooperant/Management expense sync endpoint
- `syncOprema`
- `parseFiskalniImage`
- `parseFiskalni`
- `saveFiskalni`
- `saveFiskalniMapiranje`
- `createArtikal`

**Confirmed `doGet` actions:**
- `ping`
- `getParcelGeo`
- `getParcelMeteo`
- `getParcelMeteoLatest`
- `getAllMeteoLatest`
- `getStammdaten`
- `getOtkupi`
- `getKartica`
- `getAgromere`
- `getMgmtKartica`
- `getMgmtOtkupiByStanica`
- `getMgmtSaldoOM`
- `getMgmtSaldoKupci`
- `getMgmtOtkupPoOM`
- `getMgmtPredatoPoKupcu`
- `getMgmtAll`
- `getMgmtFakture`
- `getMgmtFakturaStavke`
- `getVozacOtkupi`
- `getVozacZbirne`
- `getWarRoomDemand`
- `getDispecer`
- `getKamionStatus`
- `getTretmani`
- `getOprema`
- `getKooperantProizvodnja`

### 7.0.2 Confirmed Backend Response and Failure Contracts

The supplied router confirms these response/failure rules:

- **uniform JSON envelope:** all router exits return JSON through `jsonResponse(obj)`.
- **batch sync summary:** list-oriented sync actions return `processed`, `succeeded`, `failed` and per-record `results` arrays.
- **idempotent success semantics:** duplicate/client-replayed sync records still return `success: true` with `status: 'existing'` when the row already exists.
- **auth failures:** the router returns structured unauthorized or forbidden payloads with `code: 401` or `code: 403` where explicit gating exists.
- **catch-all protection:** top-level `doPost`/`doGet` catch blocks return `{ success: false, error: err.message }` instead of throwing raw platform errors back to the client.
- **public-read bridge rule:** the frontend can keep one POST-first request contract because public parcel geo/meteo reads are also handled in `doPost` before token validation.

### 7.1 Authentication Endpoints
- `login` is the explicit session entrypoint.
- Session model is PIN + role + entity ID/token.
- Current token TTL in the supplied backend is 24h via `CacheService`, while refresh/renewal beyond that cache window remains a future-hardening area.

### 7.2 Otkup Endpoints
- `sync`
- `getOtkupi`
- supporting PDF upload via `uploadPdf`

### 7.3 Dispatch Endpoints
- `getDispecer`
- `getKamionStatus`
- `saveDispecer`
- `updateDispecer`
- `removeDispecer`
- `saveWarRoomDemand`
- `removeWarRoomDemand`
- `updateKamionStatus`

### 7.4 Financial Endpoints
- direct canonical financial posting stays primarily in VBA
- PWA financial read actions include `getKartica`
- fiscal private financial intake uses `parseFiskalni`, `parseFiskalniImage`, `saveFiskalni`, `saveFiskalniMapiranje`

### 7.5 GIS and Meteo Endpoints
- `getParcelGeo`
- `saveParcelPolygon`
- `getParcelMeteo`
- `getParcelMeteoLatest`
- `getAllMeteoLatest`
- `scheduledMeteoFetch`

### 7.6 Reports and KPI Endpoints
- `getMgmtAll`
- `getKooperantProizvodnja`
- implicit export-fed reads from stammdaten/mgmt reports

### 7.7 Sync Endpoints
- `sync`
- `syncTretman` / compatibility `syncAgromere`
- `syncTrosak` — active Kooperant/Management expense sync endpoint
- `syncOprema`
- `syncZbirna`

### 7.8 Admin / Config Endpoints
- no dedicated admin-control plane is formally documented beyond exports, config reads and deploy discipline
- `getStammdaten` acts as primary configuration/bootstrap feed
- `createArtikal` is the notable controlled write endpoint touching master catalog

### 7.9 External SEF HTTP Contracts

The desktop SEF integration also depends on the following active external HTTP contracts, executed directly from VBA rather than through GAS:

| External endpoint | Method | Auth | Called by | Writes | Response model | Failure modes |
|---|---|---|---|---|---|---|
| `/api/publicApi/sales-invoice/ubl?requestId={SEFSubmissionID}` | POST | `ApiKey` header + optional `X-SEF-ENV` | `SubmitUBLInvoice()` | external SEF invoice intake | normalized `clsSEFResponse` with `Success`, `Accepted`, `Rejected`, `apiStatus`, `sefDocumentId` | config missing, transport exception, `400/409/422` reject, `429` rate-limit, other HTTP failure |
| `/api/publicApi/sales-invoice?invoiceId={SEFDocumentId}` | GET | `ApiKey` header + optional `X-SEF-ENV` | `GetInvoiceStatus()` | none directly | normalized status snapshot including exact external status | config missing, transport exception, HTTP failure, malformed/partial response |
| `/api/publicApi/sales-invoice/cancel` | POST | `ApiKey` header + optional `X-SEF-ENV` | `CancelInvoiceOnSEF()` | remote cancel intent | normalized `clsSEFResponse` with cancel result | invalid status/action, missing comment, HTTP failure |
| `/api/publicApi/sales-invoice/storno` | POST | `ApiKey` header + optional `X-SEF-ENV` | `StornoInvoiceOnSEF()` | remote storno intent | normalized `clsSEFResponse` with storno result | invalid status/action, missing comment, HTTP failure |

---

## 8. State, Sync, and Offline Model

### 8.1 State Ownership
- `AppState` is the client in-memory authority for shared runtime UI state.
- shared runtime flags are normalized around one canonical runtime branch exposed through `window.appRuntime`, instead of a second private app-shell runtime object.
- legacy globals (`db`, `stammdaten`, `mgmtData`, `qrScanner`, `selectedMera`, `parcelExpertOpen`, `appRuntime`) remain compatibility aliases over canonical state rather than independent sources of truth.
- IndexedDB stores are local persistence for offline-first flows.
- Google Sheets remain shared operational state.
- Excel desktop remains canonical for formal documents, finance and master data.

### 8.2 IndexedDB Model
The currently supplied runtime opens IndexedDB through `openDB()` using `CONFIG.DB_NAME` and `CONFIG.DB_VERSION`, with `onupgradeneeded` provisioning the following active stores and deleting removed legacy kooperant store schemas during upgrade when present:

- `CONFIG.STORE_NAME` / otkup queue store keyed by `clientRecordID` with `syncStatus` and `datum` indices
- `CONFIG.STAMM_STORE` keyed by `key` for cached stammdaten blobs
- `zbirne` keyed by `clientRecordID` with `syncStatus` index
- `tretmani` keyed by `clientRecordID` with `syncStatus`, `datum` and `parcelaID` indices
- `troskovi` keyed by `clientRecordID` with `syncStatus` and `datum` indices

Operational IndexedDB access is normalized through `dbPut`, `dbGet`, `dbGetAll`, `dbGetByIndex` and `dbDelete`.

### 8.3 Sync Queue
General active runtime pattern:

1. read local pending/error records from IndexedDB
2. POST to the action-specific GAS endpoint
3. update local sync status, timestamps and last error
4. refresh badge/runtime flags in `finally`

Additional active rules:

- `syncQueueSafe()` is role-gated to `Otkupac`.
- background queue retry is started only when online and currently ticks every 60 seconds.
- successful stammdaten refresh rewrites the cached `key = all` snapshot and emits `stammdaten:updated` to invalidate feature caches.

### 8.4 Merge Rules
`mergeOfflineRecords` rules:

- server wins when local row is already synced and clean
- local wins when row is pending or has `lastSyncError`
- timestamp tie-break uses `updatedAtClient`
- row identity favors stable server ID then client ID

### 8.5 Retry Rules
- background retry on connectivity restore
- meteo fetch uses progressive retry
- SEF uses transition-aware retry/resubmit logic
- fiscal parser can retry via image/QR fallback paths

### 8.6 Idempotency Rules
- SEF uses request ID tied to submission ID
- sync endpoints are duplicate-aware
- bank import duplicate logic uses bank reference or statement-scoped composite key fallback
- Otkup append-only design minimizes edit conflict risk

### 8.7 Offline Guarantees
- app shell available offline
- cached stammdaten readable offline
- field and kooperant records can be queued offline
- local signatures and queue state survive session changes
- dispatch shared truth still requires online refresh for reliable planning state

### 8.8 localStorage Ban Exceptions
Allowed localStorage/session-only exceptions:

- auth token / session helper
- device ID
- otkupac signature asset
- temporary helper fallback such as capacity until stammdaten load

Not allowed:

- kamion status
- dispatch plans
- shared logistics state
- canonical business entities

### 8.9 Background Refresh / Scheduled Sync
- `startBackgroundSync()` every 60 seconds for Otkupac
- refresh on `online` event for all roles
- background stammdaten refresh after bootstrap
- scheduled meteo fetch 2–4x daily
- service worker handles app shell and asset caching, not API orchestration

---


### 8.10 Desktop↔Google Sync Bridge Rules

- Desktop Google integrations are gated by Google OAuth2 readiness and a valid refresh token; no Stammdaten, Kartice, MgmtReports or OTK master-sync flow should run without that prerequisite.
- `OTK-*` Google spreadsheets are treated as shared operational intake buffers, not canonical master storage; canonical ownership transfers to desktop only after a row is imported into `tblOtkup`.
- Desktop master-sync only ingests rows whose Google-side `SyncStatus` is exactly `Synced`; desktop then writes back `Synced>Master`, `Duplicate` or `SyncError[:reason]`.
- `ClientRecordID` is the active cross-system deduplication key for imported field otkupi.
- Auto-created otpremnice from PWA-imported otkupi are a desktop-side convenience layer and remain canonical only after the resulting `tblOtpremnica` row is created locally.

## 9. Role Architecture

### 9.1 Otkupac
- **landing screen:** `otkup`
- **visible tabs:** otkup, pregled, otpremnice, queue/više
- **writes allowed:** otkup records, VozacID assignment, pdf upload
- **restricted actions:** cannot do dispatch planning as authoritative management action; cannot change master data
- **critical invariants:** only Otkupac may set `VozacID` on field records; local-first queue required

### 9.2 Kooperant
- **landing screen:** `home`
- **visible tabs:** home, parcele, agromere, knjigapolja, more
- **writes allowed:** tretmani, troškovi, fiskalni private records
- **restricted actions:** cannot write master artikli, dispatch plans, canonical desktop documents
- **critical invariants:** private inputs stay in kooperant-scoped sheets; parcel and meteo views are read-heavy

### 9.3 Vozac
- **landing screen:** `zbirna`
- **visible tabs:** zbirna, transport
- **writes allowed:** zbirna sync / transport-side records
- **restricted actions:** no direct OTK write authority beyond driver role flows
- **critical invariants:** zbirna is aggregated from assigned transport records

### 9.4 Management
- **landing screen:** `pregled`
- **visible tabs:** dashboard, pregled, dispecer, otkup, partneri, agro
- **writes allowed:** dispatch plans, demand entries, kamion status updates, agro issuing documents
- **restricted actions:** may not directly assign `VozacID` to OTK rows in dispatcher flow
- **critical invariants:** planning-only boundary; two-level navigation architecture

### 9.5 Admin / Operator
- **landing screen:** Excel desktop forms / menus, including unified `frmDokumenta` shell
- **visible tabs:** n/a in PWA, all relevant desktop forms
- **writes allowed:** full master data and document chain, bank import, SEF, exports
- **restricted actions:** none within canonical desktop authority, but should follow TX and storno rules
- **critical invariants:** single-operator Excel file, backups, transactional discipline

---

## 10. Reports and Derived Views

| Report / View | Purpose | Source data | Calculation owner | Refresh trigger | Derived or canonical | Caveats |
|---|---|---|---|---|---|---|
| SaldoOM | station-level open balance | novac + otkup + agro flows | VBA/export | export/report run | derived | historical bug noted for avans edge cases |
| SaldoKupci | buyer balance | fakture + novac | VBA/export | export/report run | derived | depends on correct faktura mapping |
| OtkupPoOM | procurement by station | tblOtkup | VBA/export | export/report run | derived | aggregation/report semantics |
| PredatoPoKupcu | delivered goods by buyer | zbirna/prijemnica/faktura chain | VBA/export | export/report run | derived | depends on complete document chain |
| KPI dashboard | management quick overview | mgmt bundle + exports | PWA management | bootstrap/refresh | derived | caches and dispatch freshness matter |
| Kartica | per-kooperant financial card | novac + otkup + kartice export | VBA/export + PWA reader | export run / PWA cache | derived export | `UKUPNO` rows ignored in production parsing |
| MeteoLatest | current parcel meteo state | scheduled fetch + parcel config | GAS | scheduled | canonical current meteo read model | rate-limit / offline map caveats |
| MgmtReports | shared management export | desktop reports | VBA export | export run | derived | not canonical transaction source |

### 10.1 Management Reports
Management consumes overview, dashboard, dispatch, partner, saldo and agro views primarily from `getMgmtAll`, `MgmtReports`, `SaldoOMDetail` and the live dispatch runtime hydrated into `mgmtShellState` / `dp*` state.

### 10.2 Financial Reports
Key financial read models are:

- Kartica Kooperanta
- SaldoOM
- SaldoKupci
- Bank reconciliation candidates and open `tblBankaImport` staging queue
- invoice payment status views

### 10.3 Operational Dashboards
- Management dashboard with KPIs and alerts
- Kooperant home dashboard
- Otkupac today overview
- Dispatch 3-column board

### 10.4 KPI Calculations
Documented KPI families include:

- today otkupi / kg / waiting kg
- dispatch plans / kamioni / demand
- kooperant saldo and alerts
- knjiga polja result = proizvodnja − agrohemija − troškovi

### 10.5 Materialized / Derived Views
- `SaldoOMDetail`
- `MgmtReports`
- `Kartice`
- `MeteoLatest`
- local caches in PWA

---

## 11. Error Handling and Recovery

### 11.1 VBA Error Pattern
- business modules propagate errors
- `_TX` wrappers rollback and return empty string on failure
- forms catch and surface user feedback
- widespread `On Error Resume Next` is forbidden

### 11.2 GAS Error Pattern
- action router returns structured failure to client
- deploy/version drift is a recurring operational risk
- some endpoints still need better auth and remote logging coverage

### 11.3 PWA Error Pattern
- `safeAsync()` wraps async feature handlers
- `apiFetch` / `apiPost` centralize auth and transport errors
- UI feedback uses toasts, inline validation, sync badges and queue inspection

### 11.4 Recovery Procedures
- offline records remain queued until network returns
- duplicate-aware sync reduces double-posting
- dispatch plans can be updated/removed explicitly
- meteo retries on partial upstream failure
- fiscal parsing falls back from native detector to photo pipeline
- desktop document entry blocks on duplicate document numbers before `_TX` writes
- OM ulaz blocks on overpayment of open otkup or insufficient OM avans
- storno refreshes orphan-document warnings for dependent otpremnice/prijemnice waiting on a new zbirna
- bank-import reconciliation can be retried from staging while failed rows remain marked `Error` and deferred rows remain `Skip` without deleting source bank facts

### 11.5 Stuck-State Recovery
- `SEF_SENDING` with `SEFDocumentId` present → `RefreshSEFStatus_TX()` is the first-line recovery path because the remote document may already exist and only local completion is missing.
- `SEF_SENDING` without `SEFDocumentId` → `RecoverStuckSEFSendingInvoice()` moves the faktura to a recoverable technical-failure path so resend logic can be attempted safely.
- `SEF_TECH_FAILED` → outbound resend may reuse the last submission request body and payload hash rather than creating a brand-new logical submission body.
- `SEF_REJECTED` → `PrepareRejectedInvoiceForResubmit()` clears rejection-specific blocking fields and returns the faktura to `WF_SEF_READY` for corrected resend.
- `WF_SEF_SENT` and `WF_SEF_SYNC_ERROR` → `RefreshPendingOutboundInvoices_TX()` is the batch recovery path and intentionally paces refresh calls with a fixed wait between invoices.
- sync blocked rows outside SEF still remain visible in queue/error state

### 11.6 Logging, Journal, and Audit Trail
- desktop runtime now has three explicit local resilience layers: daily text logs (`modLogError`), per-table append journals plus startup backups (`modJournal`), and transactional rollback through `_TX` wrappers
- `AppendRow()` is the audit/journal hinge because every successful append immediately writes a CSV journal line with table headers and timestamp
- log, journal and backup helpers are intentionally best-effort and must never block the underlying business operation
- startup-time purge policies keep local `Log/`, `Journal/` and `Backup/` folders bounded to a 30-day retention window
- `CheckJournalForRecovery()` is the active desktop crash-warning mechanism for append-heavy tables; it is advisory, not a full replay engine
- SEF has the strongest built-in domain audit on desktop because `tblSEFSubmission` stores request/response payload context and `tblSEFEventLog` stores timestamped state/transport events per faktura
- full remote error logging remains open work
- audit trail is stronger on desktop append paths than on remote PWA/GAS paths and is still not a complete end-to-end active capability

| Scenario | Detection | Recovery path | Owner | Notes |
|---|---|---|---|---|
| SEF stuck sending | SEF state and missing completion | refresh / mark tech failed / resubmit | Operator | documented state machine |
| Crash after append before durable workbook save | today's journal count exceeds live table row count | operator review + possible manual reimport from CSV journal + backup comparison | Operator / support | `CheckJournalForRecovery()` warning path |
| Damaged workbook / bad save | startup backup exists under `Backup/` | restore most recent dated workbook copy | Operator / support | backup created at app start |
| Desktop runtime exception | dated log line in `Log/` | inspect source, err number and details; share log file for remote support | Operator / support | logging must stay non-blocking |
| Sync queue blocked | pending/error counts in PWA queue | retry on online / inspect errors | Field user + engineering | remote logging still weak |
| Partial sheet write | inconsistent sync status | duplicate-aware replay or operator review | Engineering / operator | OTK/VOZ import path sensitive |
| Offline replay conflict | merge sees pending/error local row | local wins until resolved | PWA sync engine | append-heavy flows reduce risk |

---

## 12. Security and Data Integrity

### 12.1 Authentication
- PIN-based login with token per session
- role and entity scoping are part of returned auth context

### 12.2 Authorization
- endpoint access is role/entity scoped
- dispatch requires management token
- driver views require vozac token
- kooperant private sheets remain entity-scoped

### 12.3 Input Validation
- PWA forms use inline validation and focused field corrections
- VBA modules validate business rules before save
- bank import and fiscal parse use duplicate checks
- bank mapping actions first validate that the staging row exists, is not stornirano, and has not already been processed or skipped
- meteo and geo flows validate parcel eligibility and geometry
- `frmDokumenta` validates required station/kupac/vozac context and positive numeric quantities before `_TX` saves
- module-level saves enforce mandatory keys: otkup requires `KooperantID + Kolicina`, otpremnica requires `StanicaID + VozacID + Kolicina`, zbirna requires `VozacID + BrojZbirne`, prijemnica requires `KupacID + BrojZbirne + Kolicina`
- dual-class desktop entry enables Klasa II fields only when toggled and requires positive qty/price for that branch
- `loadKartica()` strips `Opis = UKUPNO` export-total rows before rendering kooperant financial history.
- `kpSaveTrosak()` requires a selected category and positive amount before creating a local `troskovi` record.
- `fiskalniSaveToLager()` refuses checked fiscal rows that still have no resolved `artikalID`.
- digital agronom blocks `Berba` while parcel karenca is active and warns/blocks `Zastita` on unfavorable meteo unless the user explicitly confirms override.
- desktop zbirna save is gated by green kg+ambalaža validation, and desktop prijemnica surfaces shortage thresholds before save
- saved-state reconciliation and preview warnings are class-aware for kg and aggregate ambalaža-aware for package counts

### 12.4 Output Sanitization
- `escapeHtml()` is mandatory for user data into HTML
- raw `innerHTML` with unsanitized input is forbidden

### 12.5 Sensitive Data Rules
- auth/session tokens are local helper data only
- private kooperant fiscal records remain outside master artikli
- buyer/SEF credentials are held in config tables, not spread through client UI indiscriminately
- bank-reconciliation provenance is stored in `tblNovac.Napomena` as compact source metadata (`BIM`, reference, konto, opis/svrha excerpts), not as a mutable replacement for the original staging row

### 12.6 Audit / Soft Delete
- soft delete is core business delete strategy
- full audit trail is not yet complete across all layers
- SEF submissions/events already provide detailed event history for that domain

### 12.7 Backup and Restore
- file backup on app start
- CSV journal on TX writes
- restore/rotation exists conceptually in backup/recovery discipline
- remote/cloud layer still relies on Google platform durability plus re-export/re-sync

---

## 13. Known Issues

Ovde ulaze samo **aktivni** problemi.

| ID | Title | Affected layer | Impact | Workaround | Owner | Status |
|---|---|---|---|---|---|---|
| KI-001 | GAS deploy/version drift | GAS / Ops | new code may not be active until redeploy | redeploy after GAS changes | Engineering | Open |
| KI-002 | AutoSave after TX missing | VBA | Resolved by AR-002 central CommitTx autosave hook | v6.11 |
| KI-003 | Remaining MsgBox-in-module debt | VBA | business/data modules were coupled to UI feedback | resolved 2026-04-26 by migrating document/finance TX error paths to log/rollback/return and keeping user feedback in forms | Engineering | Closed |
| KI-004 | `saveParcelPolygon` intentional public exception | GAS / GIS | geo write endpoint remains pre-auth by product decision | restrict URL access operationally and revisit dedicated geo-editor auth | Engineering | Open |
| KI-005 | Google OAuth testing token limit | GAS / Google platform | 7-day token expiry in test mode | move app to production mode | Engineering | Open |
| KI-006 | Legacy `VozaciID` OTK column variants | Sheets / sync | import logic must support old column spelling | maintain compatibility on read | Engineering | Open |
| KI-007 | Full offline GIS still limited by live tile dependency | PWA SW / maps | app shell and map libraries cache locally, but map imagery still depends on external tile/network availability | keep online fallback and define an explicit offline tile strategy if field usage requires it | Engineering | Open |
| KI-008 | AGRO sheets not imported to master | PWA ↔ VBA | desktop master lacks full agromere import path | use PWA-side views/export workarounds | Engineering | Open |
| KI-009 | `UKUPNO` rows in Kartice export | export / knjiga polja | parser must skip totals to avoid false production | filter in PWA parser | Engineering | Open |
| KI-010 | MeteoHistory long-term scaling | Sheets / meteo | sheet growth risk over seasons | archive/split in future | Engineering | Open |
| KI-011 | `script.external_request` scope dependency | GAS | meteo/fiskalni/URL fetches depend on proper manifest scope | verify appsscript manifest | Engineering | Open |
| KI-012 | Dispatch capacity fallback still local helper | PWA dispatch | temporary capacity fallback can drift before stammdaten load | treat server data as authoritative | Engineering | Open |
| KI-013 | `WarRoomDemand` legacy naming | GAS / dispatch | naming drift complicates maintenance | keep compatibility, rename later | Engineering | Open |
| KI-014 | Open-Meteo scaling / rate-limit risk | GAS / meteo | higher parcel volumes may exceed comfortable free usage | batch fetch + API key | Engineering | Open |
| KI-015 | BarcodeDetector Safari limitations | PWA/iOS | native scanning unreliable in some standalone cases | photo fallback pipeline | Engineering | Open |
| KI-016 | `createArtikal` misuse risk | GAS / artikli | private fiskalni articles could leak into master catalog | enforce operator-controlled master rule | Engineering | Open |
| KI-019 | duplicate `findLegacyTabBtn` | PWA | maintenance drift / duplicate logic | deduplicate utility | Engineering | Open |
| KI-023 | Remote error logging rollout not fully verified | PWA / GAS / Ops | backend `ErrorLog` / `logClientError` path exists, but full field-device coverage still needs deployment and frontend wiring verification | deploy v6.2 GAS, verify PWA callers and ops review workflow | Engineering | Partially mitigated |
| KI-034 | IndexedDB migration / recovery plan is still minimal | PWA IndexedDB | failed upgrades or corrupted local stores still rely on manual clear/reset more than a controlled recovery path | add versioned migrate/reset discipline before launch | Engineering | Open |
| KI-035 | Sync result shape is still not fully unified across all role modules | PWA sync/runtime | cross-role observability and retry UX remain less consistent than the normalized API client | converge role sync modules on one shared result/status model | Engineering | Open |
| KI-036 | `style-src` still needs `'unsafe-inline'` because the shell still uses inline style attributes | PWA UI/security | script CSP is hardened, but style CSP is still weaker than the long-term target posture | gradually replace inline styles with class-based styling | Frontend | Open |
| KI-032 | Dual-class document partial-save risk | VBA dokumenta | Klasa I could commit while Klasa II failed in otpremnica/zbirna/prijemnica | resolved 2026-04-26 by introducing `SaveOtpremnicaMulti_TX`, `SaveZbirnaMulti_TX` and `SavePrijemnicaMulti_TX` | Engineering | Closed |
| KI-033 | OM/Kupci money-packaging partial-save risk | VBA dokumenta/novac | ambalaža, novac and status side-effects could commit separately | resolved 2026-04-26 by introducing atomic document-side wrappers for OM ulaz and kupci izlaz that use base `SaveNovac` inside one transaction | Engineering | Closed |
| KI-037 | Display-text entity lookup in `frmDokumenta` | VBA forms | duplicate kupac/stanica/faktura labels could bind saves to the wrong ID | resolved 2026-04-26 by introducing `modComboBinding` hidden-ID combos for `cmbOtkupnoMesto`, `cmbKupac` and open fakture | Engineering | Closed |
| KI-024 | ReportSaldoOM avans edge bug | VBA reports | kooperant with avans but no otkup in period may be omitted | resolved 2026-04-25 by removing premature `Exit Function` after `GetOtkupByStation` and removing inner `dict.count = 0` guard so Novac aggregation runs even when station has no Otkup in period | Engineering | Closed |
| KI-025 | Bank statement mathematical consistency verification missing | VBA bank import | imported statement rows are not yet cross-checked against statement totals/ending balance completeness | manual spot-check of extracted PDF statements | Engineering / operator | Open |
| KI-026 | SEF 429 backoff is not yet adaptive | VBA / SEF | rate-limited submit/refresh is now represented as `RATE_LIMITED` and can surface `Retry-After`, but automated retry scheduling is not implemented | manual retry after delay using surfaced Retry-After guidance | Engineering | Partially mitigated |
| KI-027 | SEF header-vs-line amount mismatch is only soft-checked | VBA / SEF payload | payload generation can flag suspicious totals without hard-blocking send | operator verification before submit | Engineering / operator | Open |
| KI-028 | `Workbook_Open` ran `ImportBankaInbox_TX` automatically and could leave Excel invisible on failure | VBA app lifecycle | operator could be locked out of an invisible Excel if the auto-import failed at boot | resolved 2026-04-25 — `Workbook_Open` now delegates to `StartApp`, auto-import removed from boot path, Excel visibility restored in EH | Engineering | Closed |
| KI-029 | `frmSplash` hardcoded version label `v2.1.0` while `APP_VERSION` advanced to v2.2.1 | VBA UI | splash branding diverged from actual release | resolved 2026-04-25 — splash label now derived from `APP_VERSION` constant | Engineering | Closed |
| KI-030 | `modReportModernUI` declared `Public Const UI_BG`/`UI_PANEL`/etc. with values conflicting with `modModernUI.Private UI_BG`, while no caller used the module | VBA theming | dead code with cross-module color collision in global scope | resolved 2026-04-25 — `modReportModernUI`, `modUI` and `frmModernUIIzvestaji` deleted; only `modTheme` and `modModernUI` remain | Engineering | Closed |
| KI-031 | `frmBankaImport` carried 8 `_Preview` helper functions duplicating `modBankaMapiranje` private helpers | VBA bank reconciliation | preview/save logic could drift across the two copies | resolved 2026-04-25 — 8 helpers in `modBankaMapiranje` made `Public`, form preview now calls them directly, ~290 lines of duplicate removed from the form | Engineering | Closed |
| KI-038 | Fakturisanje duplicate/partial-create risk | VBA fakturisanje | selected prijemnice could be double-fakturisane or faktura/stavke/prijemnica updates could partially persist | resolved 2026-04-26 by hardening `CreateFaktura_TX`, checked `AppendRow`, `RequireUpdateCell`, duplicate/storno guards and explicit `cmbFaktura` print selection | Engineering | Closed |
| KI-039 | Desktop otkup multi-step transaction split | VBA otkup | Klasa I, Klasa II, novac and avans could commit separately for one operator action | resolved 2026-04-26 by introducing `SaveOtkupMulti_TX` and moving form orchestration to one business wrapper | Engineering | Closed |
| KI-040 | Sledljivost unchecked link updates | VBA sledljivost | auto-link/manual-link could continue without confirming `OtpremnicaID` write success | resolved 2026-04-26 by introducing guarded `AutoLinkOtkupOtpremnica_TX` and `RequireUpdateCell` for critical link writes | Engineering | Closed |
| KI-041 | PWA-first traceability still needs optional workflow enhancements | PWA / VBA sledljivost | PWA will normally simplify linking by carrying driver/zbirna context, but future UX can better surface PWA suggestions in desktop repair forms | keep current canonical chain and repair form; treat PWA-suggested filtering/linking as explicit roadmap enhancement | Engineering | Open |
| KI-042 | Lightweight SEF JSON extraction | VBA / SEF client | simple key extraction now handles numeric-or-string IDs and tolerant simple booleans, but can still misread nested or escaped JSON response fields | keep response parsing limited to known simple SEF keys until a JSON parser is introduced | Engineering | Partially mitigated |
| KI-043 | SEF payload hash is lightweight fingerprint | VBA / SEF mapper | `ComputePayloadHash` is not cryptographic and should not be used as audit-grade proof | document as fingerprint; consider SHA-256 later if needed | Engineering | Open |
| KI-044 | SEF test procedures remain inside production modules | VBA / SEF | dev-only `Test_*` routines can clutter production modules | move to `modSEFTests` or mark dev-only before final packaging | Engineering | Open |
| KI-045 | `frmOtkupAPP.btnExit_Click` DisplayAlerts cleanup | VBA main shell | `Application.DisplayAlerts` should be restored defensively through `CleanExit` | add defensive cleanup block | Engineering | Open |
| KI-046 | `frmOtkupAPP.ResetHover` navButtons guard | VBA main shell UI | rare early mouse events can occur before navigation collection setup | guard `If navButtons Is Nothing Then Exit Sub` | Engineering | Open |
| KI-047 | `modSEFPersistance` spelling debt | VBA / SEF | module name is misspelled but compile-compatible | keep current name until a controlled rename migrates all references | Engineering | Open |
| KI-048 | SEF cancel final outcome not fully verified | VBA / SEF cancel | destructive cancel smoke now asserts Boolean service success and separates API success from final external status, but final business cancellation still needs controlled outcome evidence across allowed SEF states | treat cancel API success and final cancel-like external status as separate evidence | Engineering | Open |
| KI-049 | SEF storno already-STORNO guard needs test classification cleanup | VBA / SEF storno tests | validator correctly blocks storno for invoices already in external `STORNO`; v6.7 test suite classifies this as expected SKIP rather than FAIL | pre-check implemented in destructive test suite | Engineering | Closed |
| KI-050 | `StornoInvoiceOnSEF_TX` error handling can still mask original validator error | VBA / SEF service | validator message such as “Invoice cannot be storno on SEF in status: STORNO” could be replaced by generic/invalid-procedure errors if EH did not capture original Err before logging/rollback | v6.7 EH cleanup captures original `Err.Number`, `Err.Description` and `Err.Source` before rollback/logging | Engineering | Closed |
| KI-051 | SEF accepted/final buyer-side lifecycle not yet live-tested | VBA / SEF status sync | live submit/refresh reached `SENT`, but an external `ACCEPTED` final refresh still needs evidence | continue refreshing accepted-capable test invoice or run controlled accepted scenario | Engineering / operator | Open |
| KI-052 | Dev/test modules must remain non-operator surface | VBA / tests | `modBusinessFlowProTests`, `modSEFTests` and reset helpers are valuable regression tools but can mutate/destructively test data if invoked accidentally | keep dev-only naming, hide from UI and remove/lock before production packaging if needed | Engineering | Open |



---

## 14. Active Roadmap and Open Architecture Work

Ovde ulaze samo stavke koje su i dalje arhitektonski relevantne i otvorene.

| ID | Description | Why it matters | Affected modules | Dependencies | Target status |
|---|---|---|---|---|---|
| AR-001 | Remote error logging (`logError` / `logClientError` / GAS `ErrorLog` workbook) | observability across field devices | PWA, GAS, VBA | GAS backend path exists in active `Code.gs`; remaining work is deployment verification plus confirming frontend wiring across all PWA roles | Partially implemented / verify |
| AR-002 | AutoSave after transaction commit | Reduces crash-loss window on desktop master | VBA TX/commit path | Implemented centrally in clsTransaction.CommitTx via AutoSaveAfterCommit | P0 | Done v6.11 |
| AR-004 | Full endpoint verification and deploy hygiene | avoids silent non-working features after code changes | GAS router/deploy | release discipline | P0 |
| AR-005 | Offline map-tile strategy beyond shell asset caching | full GIS offline-readiness still needs a deliberate tile/offline-map approach even after self-hosting Leaflet assets | sw.js, parcel map stack, deploy discipline | asset + provider strategy | P1/P2 |
| AR-005a | Bank statement integrity verification | ensure imported PDF statements reconcile with statement totals and detect missing pages/rows | modBankaImport, modBankaImport_PdfText, reports | parser/math validation design | P1 |
| AR-006 | Final runtime-state and role-sync cleanup | stabilizes the modular client beyond the cleaned-up bootstrap shell and aligns sync/status semantics across roles | PWA UI/core, sync modules | code patches + shared result model | P2 |
| AR-006a | Queue newly created kooperant oprema offline instead of direct-online `syncOprema` only | removes one of the remaining non-uniform offline-first write paths in the kooperant agronomy stack | PWA agromere, GAS `syncOprema`, IndexedDB model | queue design + migration | P2 |
| AR-006b | Extract a shared parcel meteo/detail renderer instead of DOM copy/restore between map panel and parcel detail | reduces coupling between kooperant parcel overview and parcel detail screens and makes GIS/meteo rendering safer to evolve | PWA parcele, meteo UI helpers | frontend component refactor | P2 |
| AR-007 | PHI/Karenca control | agronomy safety and harvest blocking logic | artikli, agromere, knjiga polja | `Karenca` semantics | P2 |
| AR-008 | Foto dokumentacija | links field records to media evidence | PWA, Drive, sheets | upload flows | P2 |
| AR-009 | Notifications (frost/prices/treatment) | agronomy and ops alerting | meteo, PWA, messaging | provider integration | P2/P5 |
| AR-010 | PWA hladnjača / prijemnica flow | closes receiving digitization gap | new role/UI, master sync | receiving workflow design | P3 |
| AR-011 | Audit trail | operational trust and reviewability | VBA, GAS, reports | storage/log design | P3 |
| AR-012 | Kontni plan / bookkeeper export | accounting maturity for finance layer | novac, reports, export | business rules | P4 |
| AR-013 | SEF API integration completion / automation | strengthens invoice dispatch automation narrative | SEF modules + integrations | external API readiness | P5 |
| AR-013a | Adaptive SEF retry / rate-limit handling | reduces manual operator retries on `429` and large refresh batches | SEF HTTP + refresh orchestration | backoff strategy and queueing rules | P2/P3 |
| AR-013b | Proper SEF JSON response parser | replaces lightweight string extraction for nested/escaped response fields | modSEFClient | parser/library decision | P2 |
| AR-013c | Optional SHA-256 payload hash | upgrade SEF payload identity from lightweight fingerprint to cryptographic hash if audit requirements demand it | modSEFMapper, modSEFPersistance | CryptoAPI/library decision and migration handling | P2/P3 |
| AR-013d | Move SEF test routines to `modSEFTests` | separates dev-only procedures from production modules | SEF modules | test module organization | P2 |
| AR-013e | Adaptive SEF retry scheduler | uses `RATE_LIMITED` / `Retry-After` data to schedule retries instead of requiring manual retry | SEF client, service, status sync | queue/backoff strategy | P2/P3 |
| AR-013f | Main shell defensive cleanup | close remaining low-risk shell guard items such as `DisplayAlerts` cleanup and `ResetHover` guard | frmOtkupAPP | small form patch | P2 |
| AR-014 | Supabase/postgres migration exploration | reduces GAS boilerplate and scaling friction | full stack | major migration effort | P5 |
| AR-019a | Manifest hardening / installability polish | formalize `id`, `scope`, categories, display metadata and maskable icon purpose for production-grade install behavior | `manifest.json`, icons, entry shell | PWA packaging discipline | P2 |
| AR-020a | Offline UX banner + update mechanism | users need visible offline state and explicit app-update/apply-update flow rather than silent shell refresh behavior | `sw.js`, `index.html`, app bootstrap, toast/UI layer | service worker + shell coordination | P1/P2 |
| AR-019 | IndexedDB pruning / retention policy | prevent unbounded growth of local stores, stale synced rows and oversized caches on long-lived field devices | IndexedDB stores, sync engine, bootstrap | pruning rules + migration-safe cleanup | P2 |
| AR-021 | IndexedDB migration / recovery strategy | launch-ready local persistence needs a controlled recover/reset path when DB upgrades fail or stores drift | `src/js/services/db.js`, bootstrap, support playbooks | migration design + safe reset UX | P1 |
| AR-022 | Unified sync result and retry contract across roles | Otkupac, Kooperant and Vozac flows still need one shared runtime status vocabulary for observability and UX | sync modules, `app.js`, badge/reporting surfaces | shared sync result design | P1/P2 |
| AR-023 | Cross-role smoke test matrix for online/offline transitions | the shell now spans more roots and stricter CSP/SW behavior, so release confidence requires one mandatory operator checklist | PWA roles, auth, sync, SW, offline flows | test pack + release discipline | P0/P1 |
| AR-024 | Style-CSP hardening / inline-style reduction | script CSP is already strict, but style hardening still needs class-based cleanup to remove `'unsafe-inline'` from `style-src` | `index.html`, feature renderers, CSS structure | frontend cleanup | P2 |
| AR-020 | Token expiry / session-expiry UX | prevent silent 401 loops and clarify re-auth path after cached token/session expiration | GAS auth, `auth.js`, app bootstrap | token lifetime and logout UX rules | P1/P2 |
| AR-025 | PWA-assisted traceability suggestions | make the PWA-first model more visible in desktop repair UI without changing the canonical `OtpremnicaID` bridge | PWA otkup flow, MasterSync, `frmOtkupniBlokovi`, `modSledljivost` | explicit design for suggested zbirna/document context and operator override rules | P2 |
| AR-026 | Dedicated saldo/reporting module for cross-domain station balance | avoid expanding `modOtkup.GetSaldoByStation` into mixed accounting/reporting logic | reports, novac, banka, otkup, dokumenta | clear accounting rule for Banka/Novac/Isporuka deductions | P2 |
| AR-027 | Full form lifecycle cleanup pass | standardize safe `Activate`, QueryClose, navigation and debug-button removal across remaining forms while preserving working chrome patterns | remaining `frm*.frm` | compile and visual smoke tests | P1/P2 |


---| AR-030 | Complete SEF cancel/storno certification | cancel/storno are destructive fiscal workflows and need final-outcome verification, not only API smoke | `modSEFTests`, `modSEFService`, `modSEFStatusSync` | destructive test gating, valid SEF statuses, preserved EH errors | P1 |
| AR-031 | Move/lock dev-only test utilities before production packaging | prevents accidental invocation of destructive test/reset helpers by operators | `modSEFTests`, `modBusinessFlowProTests`, `modDevReset` | production packaging decision | P1 |
| AR-032 | Accepted-status live refresh evidence | closes final outbound lifecycle proof beyond SENT | `modSEFStatusSync`, `frmSEF`, test workbook | SEF test invoice that can be accepted by receiver/test environment | P1 |


## 15. Deprecated and Transitional Elements

Ovde dokumentovati ono što još postoji u sistemu, ali više nije target arhitektura.

| Element | Current status | Replacement | Removal plan | Notes |
|---|---|---|---|---|
| WarRoom / `wr*` naming | transitional compatibility | Dispečer / `dp*` naming | remove after code and data cleanup | some legacy demand naming still present |
| Monolithic PWA file | deprecated / no longer target architecture | modular PWA structure under `src/js/features|services|ui|utils` | keep only compatibility helpers until residual wrappers disappear | legacy compatibility remains but modular tree is canonical |
| Shared-state localStorage | deprecated / forbidden | IndexedDB + server sync | keep only allowed helper keys | architectural invariant now |
| Dispatch direct OTK write assumption | deprecated | planning-only dispatch + field-side assignment | keep forbidden by invariant | important boundary |
| Raw `fetch()` pattern | deprecated | `apiFetch` / `apiPost` | remove remaining occurrences | invariant |
| Runtime CDN library loading | deprecated | self-hosted `./vendor/` assets + SW cache | keep only local vendor loads in active shell | launch hardening completed in current shell |
| Management `tabs.js` root intercept | deprecated / removed from active router | `role-nav.js` + `showMgmtRoot(...)` | keep management ownership out of `showTab(...)` | non-management tabs router is now canonical |
| Raw unsanitized `innerHTML` | deprecated | `escapeHtml` + DOM helpers | remove remaining occurrences | invariant |
| `VozaciID` legacy OTK column | transitional compatibility | `VozacID` | remove after old sheets retired | read compatibility still needed |

---

## 16. Glossary

| Term | Meaning |
|---|---|
| Otkup | procurement intake of fruit from kooperant |
| Otpremnica | transport/shipping document from station toward buyer/cold storage |
| Zbirna | aggregate shipment document created from multiple transport records |
| Prijemnica | receiving document at cold storage / buyer side |
| OM | Otkupno mesto (procurement station / cash pool context) |
| SEF | Sistem elektronskih faktura |
| Kartica | chronological financial card / statement for kooperant |
| Dispečer | dispatch/logistics planning UI and data domain |
| Knjiga Polja | field book with production, costs, treatments and result |
| Fiskalni | fiscal receipt ingestion flow for private kooperant purchases |
| Stammdaten | exported shared master data tabs for PWA |
| MeteoLatest | current meteo read model per parcela |
| MeteoHistory | append-only meteo history |
| MgmtReports | exported management aggregate views |

---

## 17. Revision Metadata

### 17.1 What Changed in This Version
Ova verzija (**v6.9 final**) promoviše post-v6.8 business-core hardening u canonical snapshot. Fokus je na `modFaktura`, `modDokumenta`, `modOtkup` i proširenim regresionim testovima.

Uneto je naročito:

- `modFaktura` sada canonical veruje caller-u samo za `PrijemnicaID`; količina, cena, klasa i broj prijemnice se uvek čitaju iz `tblPrijemnica`.
- `UpdateFakturaStatus` sada radi dvosmerni recompute, čuva postojeći `DatumPlacanja`, čisti ga pri reopen-u i preskače stornirane fakture.
- `PrintFaktura` blokira aktivno štampanje stornirane fakture.
- Dodat je `RunFakturaSmokeSuite`; potvrđen je baseline od 18 pass / 0 fail.
- `modDokumenta` `_TX` wrapper-i čuvaju originalnu grešku pre logovanja/rollback-a, a `SaveZbirna` i `SavePrijemnica` propagiraju originalni uzrok greške.
- `modDokumenta` read helper-i (`GetOtpremniceByZbirna`, `GetOtpremniceByStation`, `GetZbirnaByKupac`, `GetPrijemniceByKupac`) vraćaju aktivne redove po defaultu kroz interni `ExcludeStornirano`.
- `modDokumenta` base writer-i imaju fail-fast input validaciju za obavezne ID-jeve/brojeve, klasu, količinu, cenu i ambalažu.
- Canonical ime kolone za vraćenu ambalažu u `tblPrijemnica` je dokumentovano kao `kolAmbVracena`; ne koristiti `KolAmbalazeVracena`.
- `modBusinessFlowProTests` je proširen testovima za dokumentnu input validaciju, stornirano read-helpere i dual-class wrapper-e.
- `modOtkup` dobija EH preservation, propagaciju originalne greške iz `SaveOtkup`, validaciju klase i `ExcludeStornirano` read helpere.
- `modBusinessFlowProSuite` je proširen otkup hardening testovima za invalid cenu, invalid klasu i stornirano exclusion u `GetOtkupByStation`/`GetOtkupByKooperant`.

### 17.2 Migration Notes

Ovaj reference update ne zahteva obaveznu Excel/GAS data migraciju.

Obavezne code/config napomene:

- Ako postoji konstanta za vraćenu ambalažu prijemnice, njena vrednost mora biti `kolAmbVracena`.
- Ako postoje testovi ili helper-i koji koriste `KolAmbalazeVracena`, treba ih zameniti sa `kolAmbVracena`.
- Posle primene v6.9 izmena potrebno je uraditi `Debug > Compile VBAProject`.
- Dev/test moduli (`modBusinessFlowProTests`, `modFakturaTests`, `modNovacTests`, `modSEFTests`) ostaju inženjerski alati i ne smeju biti dostupni operatoru kroz normalan UI.

### 17.3 Verification Evidence

Potvrđeni regression baseline za v6.9:

- `RunFakturaSmokeSuite`: 18 pass / 0 fail.
- `RunBusinessFlowProSuite`: proširen dokumenta/otkup suite; korisnik je potvrdio da svi testovi prolaze nakon korekcije `kolAmbVracena` imena kolone.
- Raniji v6.8 `RunNovacSmokeSuite` baseline ostaje važeći za `modNovac`.
- SEF v6.7/v6.8 baseline ostaje važeći za SEF send/refresh/status/UI hardening.

### 17.4 Editorial Validation Checklist
- [x] Dokument je samodovoljan i čitljiv bez starijih verzija.
- [x] Nema formulacija tipa “same as previous” ili “unchanged from earlier”.
- [x] Sve aktivne tabele su eksplicitno navedene.
- [x] Svi aktivni endpoint-i su eksplicitno navedeni.
- [x] Sve uloge i write-authority pravila su eksplicitno navedene.
- [x] Source-of-Truth Matrix je kompletan.
- [x] Svi aktivni known issues su navedeni.
- [x] Sve otvorene arhitektonske roadmap stavke su navedene.
- [x] SEF dual-status model je dokumentovan kao canonical.
- [x] SEF DeliveryDate/InvoiceDate UBL validation je dokumentovan.
- [x] SEF live submit + refresh idempotency evidence je dokumentovan.
- [x] Strict BrojZbirne traceability auto-link rule je dokumentovan.
- [x] `modFaktura` canonical prijemnica-value rule je dokumentovan.
- [x] `modDokumenta` input/EH/stornirano hardening je dokumentovan.
- [x] `modOtkup` validation/stornirano hardening je dokumentovan.
- [x] `tblPrijemnica.kolAmbVracena` canonical column name je dokumentovan.
- [x] Business-flow, Faktura, Novac i SEF test baseline je dokumentovan.
- [x] Deprecated elementi su odvojeni od aktivne arhitekture.

---

## Appendix A. Quick Review Gate

Pre objave nove verzije proveriti:

1. Da li novi član tima može da razume sistem samo iz ovog fajla?
2. Da li su sve aktivne tabele i sheetovi pobrojani?
3. Da li su svi write path-ovi eksplicitni?
4. Da li je za svaki domen jasno šta je source of truth?
5. Da li su svi aktivni endpoint-i navedeni?
6. Da li su svi aktivni known issues navedeni?
7. Da li postoji ijedna rečenica koja upućuje na stariji dokument umesto da prenese sadržaj?
8. Da li su novi VBA write path-ovi pokriveni `_TX` wrapperima, fail-fast guards i form-level user feedback-om?
9. Da li forme koriste hidden-ID combo binding i shared parser, umesto display-text lookup-a i lokalnih parser duplikata?
10. Da li je `AutoLinkOtkupOtpremnica` pokriven strict `BrojZbirne` regression testom?
11. Da li SEF live baseline uključuje submit, rejection persistence, local UBL date guard i repeated refresh?
12. Da li destructive SEF cancel/storno testovi imaju explicit config/user confirmation i da li se outcome ne preuveličava?
13. Da li `CreateFaktura` koristi canonical `tblPrijemnica` vrednosti umesto UI payload-a?
14. Da li `RunFakturaSmokeSuite`, `RunNovacSmokeSuite` i prošireni `RunBusinessFlowProSuite` prolaze?
15. Da li se u kodu/testovima koristi `tblPrijemnica.kolAmbVracena`, a ne `KolAmbalazeVracena`?

Ako je odgovor na pitanje 7 = da, dokument nije gotov.

---

## Appendix v6.11 — Pre-Launch Hardening Addendum

### A.v6.11.1 Persistence Rule

Every successful `clsTransaction.CommitTx` is now expected to trigger `AutoSaveAfterCommit`. The hook is centralized inside the transaction class instead of relying on each `_TX` wrapper to remember a save call. Debounce is allowed; lack of a save log on every commit is not an error if a recent save was already recorded.

### A.v6.11.2 HTTP Utility Rule

All desktop outbound HTTP clients must use `modHttpUtils.UrlEncode` and `modHttpUtils.JsonEscape` for request construction. New API clients must not introduce private ANSI `Asc`-based encoders.

### A.v6.11.3 SEF Transport Security Rule

Plain HTTP is invalid for SEF. `SEF_BASE_URL=http://...` must fail local validation before a network call is attempted.

### A.v6.11.4 Test Data Rule

Smoke/regression suites may create test fixtures only if they are rolled back, stornirano-marked, or otherwise prevented from producing active production-health failures. `ProductionHealthCheck Fail=0` is the final workbook launch gate.

### A.v6.11.5 Remaining P1 Work

- Replace manual SEF JSON parsing with VBA-JSON wrappers after parser regression coverage is established.
- Add explicit AutoSave suspend/resume around any confirmed high-volume batch loop if debounce proves insufficient.
- Keep `saveOtkupniListPdf` disabled until real implementation exists and passes smoke tests. `syncTrosak` is active as of v6.12.

