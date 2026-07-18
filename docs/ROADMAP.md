# AgriX / OtkupApp Architecture Roadmap

**Purpose:** Active and deferred architecture work extracted from AR/CL so the canonical reference stays focused on current behavior.

---

## 0. Purpose
This document owns future hardening and post-launch cleanup candidates. Active known issues and accepted risks are consolidated in `KNOWN_ISSUES.md`, with roadmap follow-up items tracked here.

---

## 1. P0 / Launch Blockers

NEEDS REVIEW: Confirm whether any items remain true launch blockers after v6.21.

---

## 2. P1 / Post-Launch Hardening

### 2.1 Shared `modDataAccessGuards`
Current duplicated exact-row guard patterns exist in `modBankaMapiranje` and Storno hardening surfaces. Target state is one shared exact-row/checked-access helper module used by finance, document and storno modules.
Consolidate exact-row and schema/write guards currently repeated locally in hardened modules such as `modBankaMapiranje` and `modStorno`.

### 2.2 SEF JSON Parser Migration
Replace/manual-hardening path for current manual SEF JSON parser with controlled VBA-JSON wrapper and parser regression tests.

### 2.3 GAS-Side `BrojZbirne` Duplicate Guard
Optional defensive duplicate guard for `(VozacID, BrojZbirne)` to cover multi-device same-driver offline collision.

### 2.4 `saveIzdavanje` Idempotency
Optional server-side idempotency by stable client issuance ID for management agrohemija issuing. Target state: retrying the same issuance after an ambiguous network response returns the existing row instead of appending another issue header.

### 2.5 Treatment Stock Decrement Contract
Define whether synced treatment consumption should decrement server-side stock and how storno, retry, idempotency and reconciliation work. Current contract treats treatment sync as evidence/history/karenca and does not decrement server-side `magacinkoop` stock.

### 2.6 UI / Business Logic Separation
Continue moving user prompts out of business modules where not directly part of form/UI behavior. `frmAgrohemija` may show UI feedback, but `modAgrohemija` and other business modules should raise errors/return results rather than controlling flow with `MsgBox`.


### 2.7 `modFaktura` Exact-Row Hardening
Status after v6.22: `PrintFaktura` and `UpdateFakturaStatus` duplicate-`FakturaID` guards are closed as current contract. Any remaining `CreateFaktura` source-prijemnica exact-row review is a separate hardening item and must not redesign the row-unique `PrijemnicaID` model.

### 2.8 Relink Regression Test
Add a targeted regression test for the orphan scenario where an old/stornirano prijemnica row exists and a new row with the same `BrojPrijemnice + Klasa` relinks the faktura stavka exactly once to the replacement `PrijemnicaID`.

### 2.9 Opening-Debt (`ART_POCETNI_DUG`) PWA / Sync Filter + Tests
Introduced in v6.41 (desktop „Pocetni dug" migration via reserved virtual article `ART_POCETNI_DUG`, booked as a `MAG_IZLAZ` row with `allowNoStock`). Open follow-ups:

- **PWA / sync filter (KI-006):** `modStammdatenSync.ExportMagacinKoop` exports all `MAG_IZLAZ` rows including `ART_POCETNI_DUG` (phantom `Kolicina = 1`, virtual `ArtikalID`). Audit `src/` (PWA `MagacinKoop`/`magacinkoop` consumer + Kooperant lager validation `agromere.js`) and decide: exclude `ART_POCETNI_DUG` from `ExportMagacinKoop` (preferred — opening debt is a desktop finance migration, not PWA stock) or confirm the PWA tolerates the synthetic row. Same audit applies to any other `MAG_IZLAZ` quantity consumer.
- **Tests (`modBusinessFlowProTests`):** add coverage for `BookPocetniDug`: book → `GetAgrohemijaDug` increases by the amount; the row is excluded from `GetMagacinStanje`; storno of the magacin row restores the debt; `ReportSaldoOM` AgroZaduzenje reflects the booking. Currently the financial booking helper has no automated test.
- **Seed placement:** `EnsureArtikalPocetniDug` lives in `modAgrohemija` (lazy-seed); evaluate moving it to `modSetup` alongside the other `Ensure*` schema/seed helpers per the CLAUDE.md code map.

---


### 2.7 Dedicated Geo Editor/Auth Model

Target state:

- Replace the current public/pre-auth `saveParcelPolygon` exception with an authenticated editor/session model.
- Gate polygon writes by role/entity ownership where possible.
- Preserve public read compatibility only where explicitly accepted.

### 2.8 True Selected-Parcel Google Upsert

Target state:

- Replace full `Parcele` export inside selected-parcel sync with a true one-row upsert if field data volume makes full export too slow.
- Keep canonical `ExportParcele` mapping as the source of truth for headers/field order.

### 2.10 Row-Index Geo Wrapper Retirement
After v6.22, `SaveParcelGeoPointByID[_TX]` and `ClearParcelGeoByID[_TX]` are canonical. Row-index wrappers may remain temporarily but should be retired once all callers are confirmed to use `ParcelaID`.

### 2.9 Polygon Clear / Geometry Lifecycle Contract

Target state:

- Decide whether point clear should remain point-only or whether a separate polygon-clear operation is needed.
- If polygon clear is introduced, define transaction, audit and sync behavior explicitly.



### 2.10 Endpoint Authorization Matrix Audit

Target state:

- Reconcile the `saveParcelPolygon` security-state conflict in the docs and deployed GAS code.
- Produce one canonical endpoint authorization matrix covering public, authenticated read, authenticated write, Management-only and disabled actions.
- Add automated/fixture smoke for 401/403/entity-mismatch behavior on every write/quota-sensitive endpoint.

### 2.11 Secret and Redaction Regression Suite

Target state:

- Add regression tests proving monitoring/client-error/logging paths redact tokens, secrets, authorization headers, SEF API keys, full XML/PDF/base64 payloads and password/PIN-like strings.
- Keep monitoring diagnostic output operationally useful without leaking sensitive payloads.

### 2.12 CSP and Local-State Audit

Target state:

- Maintain a small audit script/checklist that flags accidental external script dependencies, critical service-worker cache misses and shared operational state stored in `localStorage`.



### 2.13 Reports / Derived Views Audit

Target state:

- Produce a concise inventory of every exported/materialized report tab and the canonical source tables behind it.
- Add a freshness marker or last-export timestamp to Management/finance views where stale export data can mislead operators.
- Add fixture tests for `UKUPNO` row skipping, stornirano exclusion, dirty date/number KPI safety and required-column failures.

### 2.14 Report Ownership Cleanup

Target state:

- Remove any report code that performs hidden source-table correction during read/render.
- Keep report modules as read/aggregate/format layers unless a transaction wrapper explicitly owns the write.
- Document compatibility wrappers such as `ReportStanjePoDoabvljacu()` with removal criteria.



### 2.15 Desktop Helper / UI Boundary Cleanup

- Consolidate remaining duplicate `Nz*` and parser helpers into shared helper modules.
- Remove or wrap form-local update helpers in favor of `RequireUpdateCell`.
- Continue moving residual `MsgBox` control flow out of business modules and into forms/operator shells.
- Audit startup code to keep long-running imports/syncs behind explicit operator actions.


### 2.16 Data Architecture Inventory / Schema Contract Audit

Target state:

- Produce a machine-checkable inventory of required workbook tables, Google Sheets tabs and monitoring workbook tabs.
- Add a lightweight schema-audit script/checklist that validates required columns by domain before production launch.
- Split required vs optional columns explicitly for every domain module so optional columns do not become accidental silent defaults.

### 2.17 Source-of-Truth / Derived-View Cleanup

Target state:

- Audit every report/export/cache to confirm whether it is source-of-truth, transport state or derived read model.
- Add freshness metadata to exported views where stale data can mislead operators.
- Remove or isolate any hidden mutation in read/report code.



### 2.18 GAS Endpoint Contract Audit

Target state:

- Produce a machine-checkable endpoint authorization matrix from deployed `Code.gs`.
- Reconcile documented action tables with live handlers before production handoff.
- Add fixture smoke for route presence, disabled endpoints, 401/403/entity mismatch, master-sync soft-lock and schema drift.
- Close the `saveParcelPolygon` auth-state inconsistency with one explicit architecture decision.

### 2.19 GAS Fixture Sync Suite

Target state:

- Add isolated test-sheet/folder fixtures for `sync`, `syncZbirna`, `syncTretman`, `syncTrosak` and `syncOprema`.
- Verify idempotent retry, terminal-status preservation, partial failures and cleanup without touching production business sheets.



### 2.20 Google Sheets Schema Registry / Header Audit

Target state:

- Produce a machine-checkable inventory of required Google workbooks, tabs and headers.
- Verify live `Stammdaten`, OTK, VOZ, treatment, expense, fiskalni, Kartice, MgmtReports and monitoring workbook schemas against the active AR.
- Add fixture smoke for shifted headers, missing columns and Google Sheets date-coercion risks.
- Keep `SheetRegistry` / registry-assisted lookup behavior aligned with deployed GAS handlers.

### 2.21 Google Sheets External Writeback Boundary

Target state:

- Clarify every VBA path where local Excel transactions and Google Sheets writeback interact.
- Move irreversible Google writebacks outside misleading local transaction boundaries where feasible.
- Add operator-visible diagnostics for partial external writeback failure.
- Preserve the current explicit B/F/T writeback contract for VOZ and B/F contract for OTK.


## 3. P2 / Architecture Cleanup

### 3.1 Changelog Legacy Split
Move or compress v2.2–v6.9 entries after confirming no active migration notes remain hidden there.

### 3.2 Test Automation
Add focused tests around sync, merge, dedupe, stale recovery, BankaImport parser/import/mapiranje, document-chain relink edge cases, Agrohemija basket rollback, quantity semantics and treatment sync idempotency.

### 3.3 Monitoring Noise Tuning
Review pilot event volume and tune success/failure event granularity. Current architecture intentionally avoids helper/read/update row-level monitoring and uses TX-boundary/domain events; post-pilot tuning should preserve that boundary unless there is a concrete diagnostic need.

### 3.4 Endpoint Deprecation Cleanup
Confirm disabled/deprecated GAS endpoint list and remove or guard legacy surfaces safely.

### 3.5 Shared Parser / Formatter Cleanup
Consolidate duplicate parser/formatter helpers where current architecture allows.

---

## 4. Accepted Operational Risks

- Multi-device same-driver offline `BrojZbirne` collision is accepted under one-device-per-driver operational model unless GAS duplicate guard is implemented.
- File moves after DB commit cannot be rolled back; operator/manual recovery is required if post-commit move fails.
- Treatment save does not decrement server-side `magacinkoop` stock in the current contract.
- `saveIzdavanje` duplicate-submit protection is currently client-lock based; server idempotency by stable issuance ID is future hardening.

---


- Public geo/meteo read bridge and `saveParcelPolygon` public write exception are acknowledged GIS security gaps; protect deployment URL operationally until a dedicated auth model is implemented.
- Selected parcel sync may export the full `Parcele` tab internally; accepted while it reuses canonical mapping and remains fast enough.
- Geo clear intentionally clears point fields only and does not clear polygon data unless a separate polygon-clear contract is introduced.


- Endpoint authorization for `saveParcelPolygon` is inconsistent across source documentation and must be verified against deployed code before production handoff.
- Public geo/meteo reads, if retained, are accepted as read-only exposure surfaces and must not become public write paths.

## 5. Deferred Technical Debt

Populate from `ARCHITECTURE_CHANGELOG.md` ROADMAP sections during the next pass.

---

## 6. Needs Review

- Confirm whether selected-parcel sync should remain full `Parcele` export or become true row upsert.
- Confirm whether a separate polygon-clear feature is required.

- Confirm final `saveParcelPolygon` security state and target auth model.
- Confirm canonical Banka parser module name.
- Confirm legacy CL archive split boundaries.


## 7. Pass 15 Follow-Up Items

### 7.1 Finance Architecture Follow-Up

- Review `UpdateFakturaStatus` for full two-way recompute parity with `UpdateOtkupStatus` after payment removal/storno.
- Prefer `GetIsplataForOtkup()` naming in new code while keeping `GetUplataForOtkup()` as compatibility alias where required.
- Keep finance/dev smoke suites available to engineering but inaccessible from normal operator UI.

### 7.2 Ownership / Endpoint Follow-Up

- Resolve `saveParcelPolygon` auth-state conflict by checking deployed `Code.gs`.
- Consider generating endpoint matrix directly from route declarations/tests to prevent documentation drift.
- Consider a schema registry/checker for Google Sheets headers and role-specific transport sheets.

### 7.3 Documentation Follow-Up

- If the changelog grows beyond reviewable size again, split legacy entries into one archive file per major release range.
- Add line-by-line traceability only for sections that remain high-risk after human review.


---

## 8. Known Issues Follow-Up

The active known-issues register lives in `KNOWN_ISSUES.md`. Roadmap work must stay aligned with that register.

Current follow-up items:

- Resolve `saveParcelPolygon` authorization conflict after checking deployed `Code.gs`.
- Normalize Banka PDF parser module naming after checking exported VBA modules.
- Add optional GAS duplicate guard for `BrojZbirne`.
- Define treatment stock decrement contract if automatic consumption becomes active.
- Expand regression coverage for BankaImport negative paths, PWA stale recovery/dedupe and GAS fixture sync.

## 9. Pass 17 GO Closeout Follow-Up

### 9.1 GoogleSheets Helper Cleanup

Non-blocking cleanup candidates after quota/cache hardening:

- `modGoogleSheets.GoogleRetryDelayMs` appears unused after quota-aware retry cleanup.
- `modGoogleSheets.ExtractSheetIdByTitle` appears unused after sheetId cache implementation.

Remove only after compile/search confirms no callers.

### 9.2 Backup Tab Cleanup Policy

Phased Google Sheets replacement may leave a backup tab if backup deletion fails after the new target is already live.

Current rule: leftover backup cleanup is manual/operator-controlled. Do not add automatic deletion beyond the confirmed safe path without explicit approval.


---

## 9. v6.23 Follow-Up Items

### 9.1 PWA Otkup Read-Model Regression Fixture

Add an automated fixture covering:

- one VBA/master-only otkup in `OtkupiAll`;
- one PWA operational-only otkup in `OTK-ST-*`;
- one synced otkup present in both sources with shared `ServerRecordID` but different `ClientRecordID`;
- expected merged display count and ordering.

This is post-documentation test hardening. The v6.23 behavior was reported as browser-tested.


## 8. v6.24 Follow-Up Items

### 8.1 Target Repo Git Verification

Connect the actual AgriX/OtkupApp repository or provide a source export so the v6.24 UI/runtime work can be verified against real files instead of source summary only.

### 8.2 Final Design-System Sweep

Complete deferred cleanup:

- split green/gold accent semantics cleanly;
- remove remaining old radius aliases if present;
- replace inline display styles where safe;
- audit role CSS for duplicate primitives after `components_v2.css`.

### 8.3 Google Sheets Sentinel Audit

Confirm whether `modGoogleSheets` still uses `sheetId = 0` as a sentinel and replace with a non-colliding sentinel if needed.

### 8.4 Lazy-Load Regression Suite

Add smoke/regression checks for lazy-loaded jsPDF, Leaflet, Chart.js, Firebase intercom paths and service-worker runtime cache behavior.


---

## 10. Remediation Plan — Code Review + Functional Map Triage (2026-07-18)

**Register:** `KNOWN_ISSUES.md` §8 (AUD-001..029, TL-006..008). **Per-item detail:**
`docs/AUDIT_FM_TRIJAZA.md` (full triage of all 665 FM v35 risk rows — verdict/urgency/
fix/effort with file:line evidence). Waves are independently shippable; verify each with
the existing suites (`RunBusinessFlowProSuite`, storno/palete suites, `Test_BankParse`,
`RunProductionHealthCheck`) per `RELEASE_GATES.md`.

### 10.0 Wave 0 — dead weight (S)
Delete duplicate test modules `modNovacTest.bas`, `modFakturaTest.bas`,
`modLicenceTests.bas` (keep `*Tests`; one-time manual VBE removal — `ImportAllVBA` does
not prune) and dead `modBankaImportParserClipboard.bas` + `GetFileNameFromPath2` +
dead `GroupBySum`/`SumColumn`/`IzvestajTip` enum. [AUD-016; katalog FM-0027]

### 10.1 Wave 1 — P0 / data safety
1. Sheets JSON read: quote-aware strip, `\"` handling, `\uXXXX` decode + regression
   values (commas, quotes, diacritics). [AUD-001]
2. `ImportOtkupFromPWA_TX` → thin alias of `_Core(False)`, VOZ-style messaging. [AUD-002]
3. `RequireColumns` guard on positional financial inserts — `SaveNovac` first, then
   `AddCena`/`SaveZbirna`/`CreateFaktura` (first concrete targets of §2.1). [AUD-003]
4. `RollbackTx`: per-table trap, always `CleanUp`, then `LogErr`; add `Class_Terminate`
   safety. [AUD-004]
5. Hladnjača chain set: check `otpID`/`SaveZbirna_TX` results; set `outBrPrij` only after
   successful create; seed backfill numbers from existing prijemnice of the zbirna;
   propagate link failure into the chain warning. [AUD-005]
6. Storno "false success" chain (6 small fixes): context-guard the 5 unchecked branches;
   verify each paletni detach result; verify zbirna relink count; invariant existence
   check (`0=0` must not pass for nonexistent zbirna); sum-scan error flag; multi-match
   detection in `LookupActiveID`. [AUD-020]
7. Journal: today-vs-today comparison (`CreatedAt`); journal UPDATE lines from
   `UpdateCell`; storno-marker on rollback. [AUD-006]
8. `TryParseDateValue`: month bounds + round-trip validation. [AUD-007]

### 10.2 Wave 2 — P1 / functional fixes
- Finance: novac storno broj→`NovacID` resolution [AUD-008]; stornirane fakture out of
  `FillOpenFakture` [AUD-009]; avans target-active/target-owner/no-op guards [AUD-010];
  `CreateFaktura` kupac + `Count=1` [AUD-011]; `StornoNovac` → `UpdateOtkupStatus`
  (storno of payment must not hide debt) [AUD-021].
- frmDokumenta unos set: Kl.II checkbox hard block; mandatory ambalaza smer; visible
  malina auto-zbirna failure; latest-generation prefill. [AUD-022]
- Reports: revers print stornirano filter + tip separation [AUD-012]; station-attributed
  kooperant payouts; kartice "Početno stanje" row; consistent "nema prijema" state;
  per-vrsta payment allocation; explicit `Select Case` in Report* type dispatch
  [AUD-023]; frmIzvestaj freshness (`m_cur*` in status/print), loud lazy-report errors,
  valid zbirni tab matrix [AUD-024]; `MatchesFilter` unknown operator → `Err.Raise`
  [AUD-013].
- Banka: dedupe key + account number; 3+ candidate `Err.Raise` instead of subscript
  crash; direction guard in manual Map*; auto-map behind explicit action or surfaced
  result; preview/command source alignment; stale override clamp + final saldo
  revalidation before CSV. [AUD-014/025/026]
- Print: block reprint of stornirani otkup; `.UnMerge` in `FillFakturaSablon` cleanup;
  `KarticaDetalji_Clear` on tab switch. [AUD-027]
- Otkup UI: parcela culture comparison fix; blocking date re-lock failure; loud block
  linking (+ "Izgubljeni" includes unlinked); storno filter in panel price helpers;
  kooperant free-text resolver disambiguation. [AUD-028]
- Palete: either-kutije-or-kese UI validation; sledljivost "mogući izvori" labeling;
  `Preradjeno` guard on Reassign/Detach. [AUD-029]
- Dashboard KPI stornirano filter [AUD-015]; startup/lifecycle trio [AUD-017]; infra
  drift set [AUD-018].

### 10.3 Wave 3 — consolidation (P2; extends §2.15/§3.5)
- Promote `modGoogleSheets` retry/throttle helpers to `modHttpUtils`; switch
  `modMasterSync`'s 4 raw call sites.
- `modBankaParseUtils`: one canonical copy of the ~10 per-bank format helpers; keep
  per-bank block logic and `Select Case` dispatch untouched.
- Shared `BrutoUNeto(...)` for the 6 cloned conversions; move `SaveOMUlaz_TX` from
  `frmDokumenta` into `modDokumenta`.
- `NzBlank` in `modHelpers` + opportunistic clone retirement; case-insensitive stornirano
  compare (`ExcludeStornirano`/`CheckDuplicate`).
- Document-snapshot initiative for reprint drift (PDV rate, seller/legal, revers saldo)
  [TL-006]; positional-insert conversion to name-based writes across remaining sites
  [AUD-003 tail].
- `modTestHarness` for new suites + `LastRunFailedCount()` for the E2E gate; self-update
  manifest `files_count`.

### 10.4 Wave 4 — security/process (coordinated)
- Hash PWA PINs in `ExportUsers` (dual-column one-release migration); review JMBG export
  need. [AUD-019]
- Close `saveParcelPolygon` (KI-001) per §2.10/§2.18.
- Docs: fix AR/CL version metadata; mark `instructions/` drafts historical; update
  CLAUDE.md reference list (add FM + this register with defined roles; AR stays canonical
  contracts).
- `modVbaTools`/`modRelease` folders from `tblLocalConfig` (`VBA_SRC_PATH`).

### 10.5 FM continuation
- Triage delta for FM versions beyond v35 (same method: per-row verdict/urgency/fix,
  anchored to the FM version and code commit; diff-check that already-triaged entries
  did not change).
- Commit the FM into the repo split per file (`docs/functional-map/`); add a drift-check
  script (`git hash-object` vs per-entry Referentni SHA); resolve `modMarza` status (FM
  skipped as unused, but `frmMarza` is reachable from `frmOtkupAPP`).
