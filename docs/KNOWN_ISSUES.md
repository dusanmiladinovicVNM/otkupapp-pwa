# AgriX / OtkupApp Known Issues

**Version:** v6.24 documentation refactor / design-runtime supplement  
**Status:** Active known-issues register for documentation handoff  
**Companion to:** `ARCHITECTURE_REFERENCE.md`, `ARCHITECTURE_CHANGELOG.md`, `RELEASE_GATES.md`, `ROADMAP.md`  

---

## 0. Purpose

This file consolidates current known issues, unresolved review items and accepted operational risks that were previously scattered across the Architecture Reference, Roadmap, Release Gates and Changelog.

Rules:

- Active/current issues live here and are summarized in `ARCHITECTURE_REFERENCE.md` section 19.
- Historical fixed issues remain in `ARCHITECTURE_CHANGELOG.md`.
- Future remediation work lives in `ROADMAP.md`.
- Test/verification steps live in `RELEASE_GATES.md`.
- When source material conflicts, the issue remains `NEEDS REVIEW` until checked against the deployed code/workbook.

---

## 1. Current Launch-Relevant Issues

| ID | Issue | Status | Impact | Required action | Canonical location |
|---|---|---|---|---|---|
| KI-001 | `saveParcelPolygon` authorization state is inconsistent across source documents. Some material treats it as public/pre-auth; later changelog material says it is token-protected and Management-only. | NEEDS REVIEW | Possible security exposure if polygon writes are still public. | Verify deployed `Code.gs`; update endpoint matrix, security section and gates. | AR §§9, 14, 16, 19; RELEASE_GATES §9; ROADMAP |
| KI-002 | Specific production workbook readiness is not proven by documentation refactor alone. | OPEN | A code/docs package may look ready while workbook data still fails health checks. | Run `RunProductionHealthCheck` on the target workbook and resolve blocking failures. | AR §18; RELEASE_GATES §§1, 2, 14 |
| KI-003 | Runtime gates have not been executed as part of the documentation refactor. | OPEN | Documentation can be clean while VBA/GAS/PWA runtime is unverified. | Run compile, smoke, route health, PWA role smoke, monitoring and SEF gates. | RELEASE_GATES |
| KI-004 | Canonical Banka PDF parser module name differs across source material: `modBankaImport_PdfText` vs `modBankaImportParserPdfToText`. | NEEDS REVIEW | Confusion during code lookup, onboarding and future patching. | Verify actual exported `.bas` module name and normalize AR/CL wording. | AR §§8, 19 |
| KI-005 | Endpoint matrix is documented, but must be reconciled with deployed `Code.gs` before production handoff. | NEEDS REVIEW | Authorization table may drift from deployed GAS code. | Compare AR endpoint table with deployed handler/action list. | AR §9; RELEASE_GATES §3 |
| KI-006 | Opening-debt rows on reserved article `ART_POCETNI_DUG` (v6.41 „Pocetni dug" migration) are not filtered from the `MagacinKoop` PWA export (`ExportMagacinKoop`). They reach the PWA read model as a `MAG_IZLAZ` row with a synthetic `Kolicina = 1` and a virtual `ArtikalID`. | OPEN | PWA kooperant lager validation / management stock view may show a phantom 1-unit issue for an article that does not exist in the PWA catalog. Desktop debt is correct and unaffected. | Audit `src/` `MagacinKoop` consumer; exclude `ART_POCETNI_DUG` from `ExportMagacinKoop` (preferred) or confirm PWA tolerates the row. | AR §5.8, §13.1; ROADMAP §2.9 |

---

## 2. Accepted Operational Risks

| ID | Risk | Current acceptance | Mitigation / follow-up | Roadmap link |
|---|---|---|---|---|
| KR-001 | PWA-first `BrojZbirne` generation assumes one active device per driver. Multi-device same-driver offline collision remains possible. | Accepted operationally for launch unless business process changes. | Optional GAS duplicate guard for `(VozacID, BrojZbirne)`. | ROADMAP §2.3 |
| KR-002 | BankaImport file moves happen after DB commit and cannot be transactionally rolled back. | Accepted trade-off because DB commit must precede `Processed` file move. | Clear operator manual-recovery procedure for failed post-commit move. | AR §8; RELEASE_GATES §5 |
| KR-003 | Monitoring is best-effort and can be unavailable without blocking business transactions. | Accepted by design. | Keep local logs/journals/SEF tables as canonical operational state. | AR §15; RELEASE_GATES §7 |
| KR-004 | Treatment sync records evidence/history/karenca but does not automatically decrement server-side `magacinkoop` stock. | Accepted for current launch boundary. | Define server-side stock decrement contract before enabling automatic consumption. | ROADMAP §2.5 |
| KR-005 | Google Sheets writebacks and file-system moves are external side effects and do not roll back with Excel table snapshots. | Accepted architecture boundary. | Use deferred side effects, clear error reporting and manual recovery. | AR §§3, 10, 18 |

---

## 3. Known Technical Limitations

| ID | Limitation | Current state | Follow-up |
|---|---|---|---|
| TL-001 | SEF JSON parser remains manual/lightweight. | Active limitation. | Controlled VBA-JSON wrapper migration with parser regression tests. |
| TL-002 | Exact-row guard helpers are still partly local to modules such as `modBankaMapiranje`. | Active technical debt. | Consolidate into shared `modDataAccessGuards`. |
| TL-003 | Some residual business-layer `MsgBox` usage remains outside hardened blocks. | Active cleanup item. | Move operator messaging to forms/UI layer over time. |
| TL-004 | Automated regression coverage is still uneven across PWA/GAS/VBA negative paths. | Active limitation. | Expand fixture sync tests, PWA dedupe/stale recovery tests and BankaImport negative tests. |
| TL-005 | Legacy changelog entries v2.2–v6.9 are compacted/archived rather than fully normalized in the main CL. | Documentation trade-off. | Keep archive available; normalize only if legacy audit is needed. |
| TL-006 | Reprint uses current state: PDV nadoknada rate (`CFG_PDV_NADOKNADA_STOPA`), seller/legal config and revers/ambalaza saldo are re-read at print time — historical reprints drift when rates/master data/saldo change. | Active limitation. | Document-snapshot initiative (store rate/derived values per document); see ROADMAP §10.3 and `AUDIT_FM_TRIJAZA.md` FM-0031/0032/0033. |
| TL-007 | Partial avans allocation reduces the original bank-sourced `tblNovac` row and creates a split row — no append-only allocation lineage; bank statement no longer matches the original row 1:1. | Active limitation. | Consider allocation lineage (`OsirocenoOD`/parent ID on split rows) as part of finance hardening; see AUDIT register FM-0019 #2/#3. |
| TL-008 | Business numbers (`BrDok`, poziv na broj, `BrojZbirne`) act as cross-station/generation correlation keys in storno-by-broj, banka auto-map, correction relink and print grouping; collisions are handled per-flow, not via a global uniqueness contract. | Active limitation. | Per-flow scope guards (station/generation checks); see ROADMAP §10 and AUDIT register FM-0012/0013/0021/0023/0031. |

---

## 4. Closed / Historical Issues

Historical issues that are already fixed should not be duplicated here except when they still affect current architecture.

Examples that belong in `ARCHITECTURE_CHANGELOG.md` rather than this file:

- v6.22 fixed hardcoded `pdftotext.exe` path risk.
- v6.22 fixed static `%TEMP%\pdf_extract.txt` stale-output risk.
- v6.22 fixed `GetBankaImportRowByID` raw-row regression by preserving the legacy 1x10 semantic shape.
- v6.17 fixed ambalaža invalid `Smer` / negative quantity risks.
- v6.13 fixed stale `syncing` recovery and duplicate render issues.
- v6.11 introduced SEF HTTPS-only enforcement.

---

## 5. Review Checklist

Before production handoff:

- [ ] Resolve or explicitly accept every `NEEDS REVIEW` item in this file.
- [ ] Confirm `saveParcelPolygon` deployed authorization state.
- [ ] Confirm Banka PDF parser canonical module name.
- [ ] Run `RunProductionHealthCheck` on the target workbook.
- [ ] Run required runtime gates from `RELEASE_GATES.md`.
- [ ] Record any accepted unresolved risk with owner/date in `ROADMAP.md` or this file.

## Pass 17 Closeout Notes

The GO hardening closeout reduces several previously GO-blocking risks:

- Google full-tab writes no longer use target clear-before-write.
- Kartice no longer depends on `Sheet1` fallback behavior.
- Full Google/PWA sync cannot be green if PWA unlock fails.
- MasterSync document-chain links now require exact-row checks in the affected flows.
- `SaveNovac` append failure is now a hard error.
- `ProductionHealthCheck` now includes duplicate-key preflight checks.

Remaining issue status:

- Google backup tab leftovers after successful phased replacement are an accepted manual cleanup case, not automatic deletion.
- Runtime gates still need to be executed in the target workbook/GAS/PWA environment before production approval.


---

## Resolved by v6.22 Documentation Cut

- Faktura print/status duplicate-`FakturaID` guard is no longer treated as future-only roadmap; it is current contract for `PrintFaktura` and `UpdateFakturaStatus`.
- Parcel point save/clear no longer uses physical row index as canonical architecture identity; `ParcelaID` ByID APIs are current contract.

## Still Needs Code/Deployment Confirmation

- Confirm exported VBA callers use `SaveParcelGeoPointByID[_TX]` and `ClearParcelGeoByID[_TX]` for UI/front-door parcel geo mutations.
- Confirm row-index geo wrappers, if retained, delegate to ByID APIs.
- Confirm duplicate-`FakturaID` negative tests are present in the local smoke/regression set.


---

## Resolved by v6.23

### PWA otkup overview missed VBA/master-created rows

Resolved by documenting and adopting `MgmtReports/OtkupiAll` as the master otkup read projection for PWA Management and Otkupac views. `OTK-ST-*` / `OTK-*` remains operational queue state and is merged with `OtkupiAll` rather than replacing it.

### Duplicate display risk for synced PWA otkup across operational and master projection sources

Resolved by making `ServerRecordID` / `OtkupID` the first dedup key for otkup overview merge, before `ClientRecordID`.


## 7. v6.24 Active / Deferred Issues

### KI-v6.24-01 — Target Git repository not available through current connector

**Status:** Active documentation limitation  
**Impact:** Git verification of AgriX/OtkupApp-specific files could not be completed from the currently connected GitHub repo.

Observed through GitHub connector:

- available repository: `dusanmiladinovicVNM/handoverApp`;
- repository README describes an Apartment Handover app, not AgriX/OtkupApp;
- root app loads `css/tokens.css`, `css/layout.css`, `css/components.css`, `css/forms.css`, not the AgriX `base.css` / `components_v2.css` structure;
- service worker uses `handover-v1` and Handover app shell paths.

**Required resolution:** connect or provide the actual AgriX/OtkupApp repo/files before marking v6.24 as Git-verified.

### KI-v6.24-02 — `modGoogleSheets` `sheetId = 0` sentinel diagnosis status unknown

**Status:** Active follow-up  
**Impact:** The source summary reports a diagnosis and patch proposal for replacing sentinel `0` with `-1` in nine places, but implementation status is unknown.

**Required resolution:** inspect `modGoogleSheets.bas` and confirm whether the sentinel collision is patched.

### KI-v6.24-03 — Deferred design-system cleanup

**Status:** Accepted deferred cleanup  
**Items:**

- `--accent` green/gold split cleanup deferred to final design-system sweep;
- inline `style="display:none"` to `.is-hidden` migration remains P3/cosmetic and requires JS audit;
- remaining `var(--r-md)` / `var(--r-lg)` occurrences were reported as locally resolved but not pushed/verified;
- Otpremi hero description text remains as-is by user decision.

### KI-v6.24-04 — Source-summary verification vs runtime proof

**Status:** Documentation limitation  
**Impact:** The v6.24 package records the user-provided summary and performs connector-level repo sanity checking, but it does not independently run browser, VBA, GAS or service-worker tests.

**Required resolution:** run the v6.24 gates in the real target environment.


---

## 8. Code Audit Register — Architecture Review + Functional Map Triage (2026-07-18)

**Sources:** (a) independent 5-track architecture review of `src-vba`; (b) full per-item
triage of `AgriX_Functional_Map` v35 — **all 665 recorded risk rows** (FM-0002..FM-0034,
anchored to commit `a0bc9e2` / vba v2.21.0) individually verified against current code.
The complete per-item catalog (verdict / urgency / fix proposal / effort, with file:line
evidence) lives in **`docs/AUDIT_FM_TRIJAZA.md`** — that file is the canonical detail
source; this register carries only the deduplicated actionable set.

Triage outcome: 515 Tačno (77.4%) · 78 Delimično · 63 design-accepted · 6 refuted ·
3 not statically verifiable. Urgency: 1×P0, ~52 unique P1 defects, ~200 P2.
Severity calibrated for the single-writer desktop model (multi-user/CAS claims → P2).

Remediation plan: `ROADMAP.md` §10. AUD items reference FM entries for detail.

### 8.1 P0 / data safety (fix first)

| ID | Finding | Location | Detail |
|---|---|---|---|
| AUD-001 | Sheets JSON read corrupts values containing `", "` or escaped quotes; `\uXXXX` never decoded. Live OTK/VOZ import path. | `modGoogleSheets` `ParseValuesJson`/`SplitCsvJson` (~1750-1830) | Review |
| AUD-002 | `ImportOtkupFromPWA_TX` wraps whole batch in one TX while per-sheet Google writeback `Synced>Master` is not rollbackable — rollback permanently loses acknowledged rows. | `modMasterSync.bas:287-341` vs `:2165-2177` | Review |
| AUD-003 | Positional `Array(...)` inserts without order guard on financial/document rows: `SaveNovac` (P0), `AddCena`, `SaveZbirna`, `CreateFaktura`, `modOtkup`, otpremnica/zbirna, `modAmbalaza`, `modStornoContext`, invariant class-row builder. | `modNovac.bas:197-204` + katalog FM-0006/0011/0014/0015/0019/0033/0034 | Both |
| AUD-004 | `RollbackTx` has no error handling and no guaranteed `CleanUp` — failed restore leaves tables half-restored and Excel frozen; also no `Class_Terminate` safety. | `clsTransaction.cls:77-87` | Both (FM-0003 #2/#8) |
| AUD-005 | `AutoChainHladnjaca` chain set: `otpID`/`SaveZbirna_TX` results discarded (silent failure); `outBrPrij` exposed before successful create (recovery relinks to nonexistent doc); backfill splits class numbers of the same zbirna; link failure not propagated. | `modAutoHladnjaca.bas:150-243` | Both (FM-0010 #1/#2/#4/#5/#6) |
| AUD-006 | Journal/crash-recovery: today's journal lines compared against all-time row count (cannot fire on mature data); `UpdateCell` mutations never journaled; rolled-back inserts leave journal lines. | `modJournaling.bas:203-223`, `modDataAccess.bas:210` | Both (FM-0002 #2, FM-0003 #5) |
| AUD-007 | `TryParseDateValue` accepts impossible dates via `DateSerial` rollover — used by bank statement parsing. | `modParse.bas:59-73` | Review |
| AUD-020 | Storno "false success" chain: correction context failure does not block mutation (5 branches); paletni detach false-success; zbirna relink count ignored; invariant passes `0=0` for nonexistent zbirna; sum scan error becomes valid zero; `LookupActiveID` silently takes last duplicate. Combined effect: correction can report COMPLETED while the chain is broken. | `modStornoFlow.bas:399-776, 692, 1209-1216`, `modDokumentInvariant.bas:93-95, 178-203`, `modStorno.bas:1002-1010` | FM-0012/0013/0015 |

### 8.2 P1 / functional defects (confirmed in code)

| ID | Finding | Location | Detail |
|---|---|---|---|
| AUD-008 | Novac storno passes operator-entered business number where `NovacID` is required (every other doc type resolves broj→ID first); no document-level novac storno API. | `frmDokumenta.frm:3230` → `modStorno.bas:759` | FM-0018 #1, FM-0019 #7 |
| AUD-009 | `FillOpenFakture` keeps stornirane fakture selectable — uplata can be booked against a cancelled invoice. | `frmDokumenta.frm:3108-3144` | FM-0018 #2 |
| AUD-010 | `ApplyAvansToOtkup/Faktura`: no target-active (storno) and no target-owner check; no-op returns `True` (inflates batch counters in frmBankaExportPregled). | `modNovac.bas:1067-1204, 487-550` | FM-0019 #4/#5/#6/#11 |
| AUD-011 | `CreateFaktura`: no prijemnica-kupac ownership check; takes `rows(1)` without `Count=1` guard (extends ROADMAP §2.7). | `modFaktura.bas` validation loop | FM-0034 #1/#2 |
| AUD-012 | Revers print reconstruction: no stornirano filter and merges multiple ambalaza types (tip from first row, sum over all). | `frmIzvestaj.frm:2087-2164` | FM-0029 #4/#5/#16 |
| AUD-013 | `MatchesFilter` returns True for unknown operator — a typo silently removes the criterion; carries the whole `ExcludeStornirano` layer. | `modArrayUtils.bas:94-95` | FM-0027 #1 |
| AUD-014 | Opening `frmBankaImport` books money (`AutoMapStrongKeysBankaImport_TX` on Activate under `On Error Resume Next`); manual Kupac mapping always becomes avans (empty `FakturaID`); blok preview reads a different source than the command. | `frmBankaImport.frm:70-74, 332, 340-346 vs 595` | FM-0024 #1/#2/#3 |
| AUD-015 | Dashboard KPIs count stornirane rows; frmDokumenta twin excludes them — screens disagree after storno. | `frmOtkupAPP.frm:1356-1422` | Review |
| AUD-016 | Duplicate test modules shipped twice (~1,430 lines) — E2E gate `Application.Run` resolves ambiguously; dead `modBankaImportParserClipboard` (555 lines). | `modE2EReleaseGate.bas:34-75` | Review |
| AUD-017 | Startup/lifecycle trio: outer startup EH clears `Err` before logging; failed startup backup aborts remaining startup; pending autosave `OnTime` survives close (ghost reopen). | `ThisWorkbook.doccls:39-52`, `modMain.bas:91`, `modJournaling.bas:527` | Review |
| AUD-018 | Infra drift set: `GetConfigValue` exact vs `SetConfigValue` trim; legacy `tblConfig` still required; `modMonitoring` `ActiveWorkbook` fallback; `FindVOZSheets` no pagination; `BulkPushPendingForStanica` bypasses `UpdateCell`; `AppendRow` leaves phantom row on mid-write failure. | see catalog Blok A | Both |
| AUD-019 | Plaintext PWA PINs (all roles) and kooperant JMBG exported to Google Sheets while `ExportConfig` already filters credentials. | `modStammdatenSync.bas:2186-2379, 1327-1381` | Review |
| AUD-021 | `StornoNovac` does not refresh `UpdateOtkupStatus` for the linked otkup — storno of a payment permanently hides the debt from the payout list (`GetOpenOtkupi` trusts stale `Isplaceno`). | `modStorno.bas:781-785`, `modNovac.bas:974/1001` | FM-0019 #16, FM-0021 #10 |
| AUD-022 | frmDokumenta unos set: Kl.II zbirne validation depends on a checkbox (source Kl.II can be silently dropped); ambalaza smer not mandatory (0 toggles books legacy OM prijem via `Case Else`); malina auto-zbirna failure silent; prefill takes first generation of a reused number. | `frmDokumenta.frm:2216-2221, 1706+1303-1308, 907-912, 2694-2703` | FM-0018 #3/#4/#5/#8 |
| AUD-023 | Reports correctness set: kartice start from zero (period net shown as "saldo"); kooperant payouts not station-attributed in `ReportSaldoOM`; missing prijemnica renders as 0% in one report and 100% in another; kupac per-vrsta payment goes entirely to the invoice's first item's vrsta; `ReportProsecnaCena` fail-open branch reachable from zbirni mode. | `modIzvestaj.bas:116-121, 636-648, 1558-1577/2257-2262, 2030/2064`, `modNovac.bas:783-789` | FM-0028 #1/#3/#5/#6/#13 |
| AUD-024 | frmIzvestaj freshness set: status/print header read `txtDatum*` while data uses `m_cur*` (stale data presented under new period); lazy report errors fully silent (old list + green status); zbirni mode offers invalid report combinations. | `frmIzvestaj.frm:603-604/1446-1447, 702-705, 488-492` | FM-0029 #1/#2/#3/#14 |
| AUD-025 | Banka set: import dedupe key omits account number (multi-account transaction silently dropped); block resolver crashes on 3+ candidates (`ReDim 1 To 2` → subscript error rolls back whole AutoMapAll batch); manual mapping never checks direction (reachable from UI). | `modBankaImport.bas:783-798`, `modBankaMapiranje.bas:1457-1497`, Map* funkcije | FM-0022 #1, FM-0023 #8/#14 |
| AUD-026 | Banka export set: stale per-blok override survives reload without clamp and `GenerisiNalogeCSV` performs no final saldo revalidation — CSV can order more than the currently open amount. | `frmBankaExportPregled.frm:504-535, 967`, `modBankaExportPregled.bas:369-374` | FM-0020 #1/#2, FM-0021 #1 |
| AUD-027 | Print set: reprint of stornirani otkup possible through raw-table fallback; `FillFakturaSablon` cleanup lacks `.UnMerge` — next faktura with more stavke renders through stale merge; kartica detail print target survives tab switch (cross-tab print). | `modPrint.bas:334/354, 1896-1924`, `modKarticaDetalji.bas:27-30` + `frmIzvestaj.frm:737-749` | FM-0031 #3/#19, FM-0030 #1 |
| AUD-028 | Otkup UI set: parcela culture warning compares wrong fields (false alarms); failed date re-lock silent and non-blocking; post-save block linking best-effort and silent (unlinked block missing from "Izgubljeni"); stornirani blok feeds default price; free-text kooperant resolver merges same-named persons. | `frmOtkup.frm:1040-1048, 711-717, 1174-1176`, `modOtkupBlok.bas:1600-1628`, `modKooperant.bas:57-82` | FM-0007 #2/#3/#5, FM-0009 #4/#5, FM-0008 #1 |
| AUD-029 | Palete set: strict UI requires both kutije and kese while core allows either (legitimate prerada blocked); sledljivost prints full otkup quantities of the whole zbirna as if pallet-specific; Reassign/Detach mutate already-processed (prerađena) pallets while prerada keeps old input state. | `frmPalete.frm:262-281`, `modPaletniList.bas:2196-2308, 1197-1437, 1704` | FM-0017 #1, FM-0016 #1/#2 |

### 8.3 Recalibrated / refuted (see catalog for full list)

- "Multi-user race / cross-user CAS" claims (FM-0022/0023/0034 and others): single-writer-per-xlsm
  deployment; realistic variant is same-instance re-entry → hardening candidates (P2), not blockers.
  `ValidateBankaImportNotProcessed` already guards BIM re-booking.
- Refuted (6): avans snapshot coverage is complete (FM-0006 #6); `tblPaletaIstorija` does not exist
  (FM-0011 #9, FM-0012 #15); `Monitor_Event` cannot raise (FM-0015 #9); empty Dictionary `Keys`
  loop is safe (FM-0020 #20); revers hidden-state has no live wrong path (FM-0031 #5).
- Design-accepted (63 rows): documented in the catalog per entry; notable: saga model of the
  hladnjača chain, VeryHidden config sheets, deferred/best-effort monitoring.
