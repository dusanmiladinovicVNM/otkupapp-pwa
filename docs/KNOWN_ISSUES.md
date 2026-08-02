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

Triage outcome (v35 base): 515 Tačno (77.4%) · 78 Delimično · 63 design-accepted · 6 refuted ·
3 not statically verifiable. Urgency incl. v85 (§8.4) + v142 (§8.6) deltas: 2×P0 (AUD-001,
AUD-030), ~75 unique P1 defects, ~230 P2. Severity calibrated for the single-writer desktop
model (multi-user/CAS claims → P2). Delta coverage: v85 = FM-0035..0084 (§8.4), v142 =
FM-0085..0140 (§8.6); full per-item detail in `docs/AUDIT_FM_TRIJAZA.md` DEO II/DEO III.

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

### 8.4 Delta v85 — new findings (FM-0035..FM-0084, 2026-07-19)

Full per-item detail in `docs/AUDIT_FM_TRIJAZA.md` DEO II. Anchors: FM-0035..0075 = `f6313dc`,
FM-0076..0084 = `a0bc9e2`. Severity calibrated for single-writer desktop. Many delta rows map to
already-registered AUD-001/003/006/016/018/019 and KI-006 — those are not re-listed here.

**IMPORTANT (RF-03/RF-04 scope):** `main` advanced past the FM anchor to v2.24.0 (`58a5075`) via
storno PRs #134-137 (`modStornoFlow` +746 lines). The v85 storno entries (FM-0011..0015) predate
that work; RF-03/RF-04 must be re-verified against `origin/main` before implementation.

| ID | Sev | Finding | Location |
|---|---|---|---|
| AUD-030 | P0 | SEF client maps HTTP 409 → REJECTED; a duplicate/conflict permanently marks the faktura rejected while the document exists on SEF (retry reuses same requestId) — risk of duplicate/incorrect invoice toward the tax authority. | `modSEFClient.bas:473-476` |
| AUD-031 | P1 | SEF correctness cluster: (a) stornirana faktura is end-to-end sendable — validator never reads `Stornirano`, `frmSEF` combo unfiltered, `StornoFaktura` doesn't touch SEF workflow; (b) qty/price truncated to 2 decimals → arithmetically inconsistent UBL; (c) DueDate < IssueDate under force-today; (d) fail-soft `HasSuccessfulSEFSubmission` (EH→False) allows double submit; (e) stale DocumentID carried through resubmit. | `modSEFValidator.bas:130-159`, `modSEFMapper.bas:190/561/577/372/411`, `modSEFPersistance.bas:528-572`, `modSEFValidator.bas:406-414` |
| AUD-032 | P1 | SEF UX/lifecycle: `modSEFService` returns SubmissionID for REJECTED/TECH_FAILED → `frmSEF` shows "Faktura poslata" for failures; public `Test_Cancel/Storno…_TX` macros with legal side-effect in Alt+F8; blank/unknown status → silent "SENT"; `frmSEF` combo change doesn't reset shown context; recovery/refresh return True on API failure + false "Recovered" event each startup for SENDING+remote-terminal. | `modSEFService.bas:384/652-686/940-960`, `frmSEF.frm:279-283/454-458`, `modSEFStatusSync.bas:144-159/311`, `modSEFClient.bas:597-600` |
| AUD-033 | P1 | Authorization chain: a user with "Matični podaci" reaches the Admin panel (Očisti tabele, Migracija, VBA import/export, fleet publish with password from `modConfig`). Guard exists only for "Korisnici" (`modMaticniLookups.bas:254-259`); shell lets `frmStammdaten` through (`OblastZaFormu`=""); `modAdmin`/`modPodesavanja`/`ShowConfigSheet` have no own check. Fix: one `MozeAdministraciju` guard (pattern in `modSetup.bas:1429`). | `modMaticniLookups.bas:254-259`, `modAdmin.bas:39/200`, `modPodesavanja.bas:725`, `frmOtkupAPP.frm:1072-1077` |
| AUD-034 | P1 | Startup/integration: `Workbook_Open` never calls `AccessWasDenied` although the comment and runbook claim it does — deny relies solely on an unchecked `OnTime` close; false `STARTUP_SUCCESS` after deny; `frmOtkupAPP.btnBanka_Click` books money (auto-map on Activate) before the auth check. | `ThisWorkbook.doccls:15-35`, `modLicense.bas:626-628`, `frmOtkupAPP.frm:728 vs 1072` |
| AUD-035 | P1 | Self-update: phase 1 Removes a failed `.frm` while phase 2 never re-imports it → component vanishes; no download manifest completeness check → mixed-version code set. | `modSelfUpdate.bas:101-105/149/261-281` |
| AUD-036 | P1 | Cenovnik stale auto-price: `If c > 0 Then txtCena…` never clears the field, so a lookup miss leaves the previous product's price in the input. | `frmOtkup.frm:407-413`, `frmDokumenta.frm:583-591` |
| AUD-037 | P2 | Release/build guard hardening: `PublishReleaseToDrive` performs no guard (placeholder/`+dirty` build shippable, no disk↔workbook SHA cross-check); `AssertBlankBuild` scans only ListObjects, missing plain-range logs (`SETUP_LOG`, test logs) that carry machine/user/path. | `modRelease.bas:19-77`, `modBuildGuard.bas:29-41`, `tools/release.sh:61` |
| AUD-038 | P2 | Sync/IO hardening: `SetPWAMasterSyncLock` full-tab overwrite deletes `STANICA_LOCK_*` keys (asymmetric vs modStanicaLock RMW); non-atomic rename-pair in Sheets swap (target name absent between renames); `modDrive` find-error treated as not-found → duplicate release artifact + first-match self-update; empty-source → header-only cloud wipe. | `modGoogleSyncOrchestrator.bas:578-605`, `modGoogleSheets.bas:834-932`, `modDrive.bas:70-94`, `modStammdatenSync.bas:1324-1337` |
| AUD-039 | P2 | Test suites unsafe as shipped/gate: `modE2EReleaseGate` reports PASS on any non-throwing `Application.Run` (child suites swallow failures internally); `modBusinessFlowProTests`/E2E have no environment guard and are shipped to clients; hard-delete cleanup misses fakture headers + ambalaza ledger. | `modE2EReleaseGate.bas:74-82`, `modBusinessFlowProTests.bas:60-86/2278-2316` |

### 8.5 Delta v85 — recalibrated / refuted highlights

- Most "Kritično"-titled SEF DTO/validator rows (FM-0042..0044, FM-0040 first-match) are P3: mutable DTOs
  aren't mutated between build and serialize (single sync call), and every write path goes through strict
  `GetSingleRowIndexByKey` before HTTP, so first-match reads are fail-closed for the send flow.
- Auth fail-open rows (FM-0053 #2/#3/#9/#15) are documented opt-in design with `EnableAuth` anti-lockout
  bootstrap → P2 (the silent plaintext-PIN fallback deserves a signal; cheapest delta).
- WithEvents wrapper "stale-click/race" rows (FM-0056/0058/0060/0062) are neutralized: rebuild resets the
  wrapper collection, so the old event sink dies and the old button cannot fire → P3.
- Refuted side-details: FM-0083 #91.23 ("banka error-code constants unused" — they are used); the
  pre-registration note that `modBusinessFlowProTests` runs inside a rollback TX (it does not — that applies
  to `RunMasterSyncSmokeSuite`).

### 8.6 Delta v142 — new findings (FM-0085..FM-0140, 2026-07-20)

Full per-item detail in `docs/AUDIT_FM_TRIJAZA.md` DEO III. Anchor: `origin/main` v2.24.0
(`9fd7087`) — the v142 header claims `a0bc9e2` but lists v2.24.0-only files, so the whole delta
was verified against `origin/main`. 38 files triaged (blocks K1..K8); entries already covered in
v35/v85 were skipped. No new clean P0 — the strongest chain (FM-0093 E2E false-green) is latent
(gate not wired into `PublishReleaseToDrive`) and rolls up to AUD-039. Rows mapping to
AUD-002/003/007/016/017/018/034/037/039 are not re-listed.

| ID | Sev | Finding | Location |
|---|---|---|---|
| AUD-040 | P1 | Agrohemija price not booked as entered: `frmAgrohemija` snapshots the basket price but `btnZavrsiIzlaz` calls `SaveMagacin` **without** `overrideCena`, so the ledger re-reads the master price (input path passes it correctly — asymmetry proves the oversight). `modAgrohemija` also writes `Cena=0/Vrednost=0` silently when the price is non-numeric → understated dug. Fix: pass `m_KorpaIzlaz(i).cena` as `overrideCena`; require price>0 for real articles (except `ART_POCETNI_DUG`). **Fixed under RF-27:** izlaz passes `overrideCena`; `SaveMagacinCore` fails closed (`Err 4206`) when a real article resolves to `cena<=0`; `ValidateMagacinInput` adds artikal/kooperant existence checks (`Err 4207/4208`). **Typed errors now reach the operator** — the form calls `SaveMagacinCore` (raises) instead of the swallowing `SaveMagacin` wrapper, so the exact reason shows in the form EH (closes FM-0113 #4). Zero-value ULAZ is allowed only via an explicit `allowZeroValue` confirmation (documented free/corrective receipt); IZLAZ stays strict. **Parcela↔kooperant is now validated** for real izlaz (`;`-list split; each parcela must exist, belong to the passed kooperant, and be active via `COL_PAR_AKTIVNA`; `PRACENJE_PARCELA` ON → parcela required, OFF → empty allowed; `ART_POCETNI_DUG` exempt). Note: `tblArtikli`/`tblKooperanti` have no active-flag column in the schema, so an "active" check applies only to parcele. Regression-guarded by `modAgrohemijaTests.RunAgrohemijaSmokeSuite` (isolated: dev-guard + `modJournaling` test-mode + TX rollback — no journal/sheet/autosave trace). | `frmAgrohemija.frm` (basket + `btnZavrsi*`), `modAgrohemija.bas` (`SaveMagacinCore`/`ValidateMagacinInput`), `modAgrohemijaTests.bas`, `modJournaling.bas` (test-mode) |
| AUD-041 | P1 | Duplicate document-number generation: `GenerateBrojPrijemnice` error handler returns a valid-looking `1/ddmmyy` instead of hard-failing; `modMasterSync.GenerateBrojZbirne` is a parallel **row-count** generator (`seq=count+1`) that produces a duplicate on gaps (`1/ddmmyy` + `…-3` → `-3` again) instead of the canonical `SuggestNextBroj`/`MaxSeqFromTable`. Fix: EH → `""`; delegate ZBR generation to `SuggestNextBroj(KIND_ZBR,…)`. | `modBrojevi.bas:203`, `modMasterSync.bas:2887-2928` |
| AUD-042 | P1 | MasterSync wrong-write cluster: `TryUpdateVozacID` returns True even when `UpdateCell` fails → GS marked `Synced>Master` with empty `VozacID`; an invalid date silently becomes **today** on both paths (OTK/VOZ); a failed header write leaves a poison spreadsheet that the next run treats as existing. Fix: check the Boolean write result; strict date parse → `SyncError`; trash/temp-name the sheet until the header write succeeds. | `modMasterSync.bas:1773-1798`, `:1547-1550`, `:2598-2600`, `:476-494` |
| AUD-043 | P1 | MasterSync document-integrity cluster: auto-otpremnica groups by `Stanica\|Datum\|Vozac\|Klasa`, mixing vrste/cene/ambalaza into one otpremnica; VOZ `LinkZbirnaToOtkupAndOtpremnica` links by CRID without membership check (same vozač/datum/not-already-linked) and overwrites existing links with no "empty-or-identical" conflict guard. Fix: add `VrstaVoca\|SortaVoca\|Cena\|TipAmbalaze` to the key; validate membership; guard link writes. | `modMasterSync.bas:668-672`, `:2701-2771`, `:2759`, `:1894` |
| AUD-044 | P1 | `modIntegritet` false-green: `WriteErr` does not raise `m_totalIssues`, while the overlay title and MsgBox read only that counter → a run can report "0 neusklađenih" alongside GRESKA blocks; an in-memory read failure returns `Empty`, which is indistinguishable from PASS. Fix: `WriteErr` increments an `ErrorCount`; MsgBox shows INCOMPLETE when errors>0; typed `IntegrityRunResult` so `Empty` ≠ PASS. | `modIntegritet.bas:1304-1310`, `:84-85`, `:59`, `:90` |
| AUD-045 | P1 | Sledljivost incomplete trace shown as complete: `TraceByZbirna` filters by the helper `OtpremnicaID` instead of canonical `tblOtkup.BrojZbirne`, and normalizes the number inconsistently (auto-link `UCase$+Trim$` vs raw trace compare) → an otkup with `BrojZbirne` but empty `OtpremnicaID` drops out; `frmSledljivost`/printed PDF present the partial trace as complete (no incompleteness marker). No corruption (manual write goes through `ReassignOtkupToOtpremnica_TX`). Fix: direct `BrojZbirne` pass with `vbTextCompare`; typed trace result with `IsComplete`. | `modSledljivost.bas:540-544`, `:282` vs `:464`, `frmSledljivost.frm:462/499/543` |
| AUD-046 | P1 | Station-mirror missing-shadow: `modMalina.EnsureVozacMirrorForStanica` doesn't confirm the station exists, its `AppendRow=0`/EH swallow failures, and `modMasterSync.StampVozacFromStanicaForMalina` + `modAutoHladnjaca` unconditionally set `vozacID=stanicaID` → a document gets an FK with no `tblVozaci` row. Fix: one canonical `IsManagedStationMirror(id)` checking the `tblStanice`+`tblVozaci` pair; re-raise in Ensure EH; verify the mirror before stamping. | `modMalina.bas:61/68`, `modMasterSync.bas:836-838`, `modAutoHladnjaca.bas:114-118` |
| AUD-047 | P2 | `modProductionHealthCheck` SEF drift + false-green: the SEF status list uses a nonexistent `SEF_CANCELLED` and misses `SEF_REJECTED/SYNC_ERROR/TECH_FAILED` (drift vs `modConfig.bas:659-663`); parent check reports OK after a child FAIL in two spots (Google `:951`, soft-delete helper `:928`). Fix: use `WF_SEF_*` constants + a state matrix; gate the parent on child delta-counters. | `modProductionHealthCheck.bas:871`, `:951`, `:928` |
| AUD-049 | P1 | **Storno celog izvoda ne postoji** — a pojedinačni storno izvodnog reda je od RF-03 zabranjen (`modStorno.ResolveNovacForStorno` / `StornoNovac_TX`), pa pogrešno mapiran red iz izvoda trenutno **nema putanju ispravke**. `modBankaMapiranje` ima samo `Map*`/`AutoMap*`/`Skip*` — nema unmap/revert; `COL_BIM_STORNIRANO` se nigde ne upisuje, a `Obradjeno="Da"` trajno zaključava staging red (`ValidateBankaImportNotProcessed`). **Odlučeno (vlasnik, ovaj tred) — DVA ishoda:** (A) *mapiranje pogrešno, PDF ispravan* → staging redovi se vraćaju u „za obradu" (`Obradjeno` → prazno), izvod ostaje uvezen, PDF se NE uvozi ponovo; (B) *PDF korumpiran/loš* → staging redovi se gase (`Stornirano="Da"`) i izvod se uvozi ponovo iz ispravnog PDF-a. Operater bira ishod. **Provereno: putanja B je već podržana u čitaocima** — `IsDuplicateBankaImport` (`modBankaImport.bas:768`), `GetBankaImportOpen` (`modBankaMapiranje.bas:78`) i `ValidateBankaImportNotProcessed` svi poštuju `Stornirano`/`ExcludeStornirano`, a 4-nivo integritet se računa nad parsiranim fajlom (`modBankaImport.bas:502`), ne nad tabelom → ponovni uvoz se validira nezavisno od starih redova. Nedostaje samo upisivač flag-a. **Hazard (oba ishoda):** storno mora ATOMIČNO oboriti i `tblNovac` redove nastale iz tog izvoda; ako se staging ugasi a novac ostane aktivan, ponovni uvoz + mapiranje = **dvostruko knjiženje**. Identitet grupe = `BrojDokumenta` (banka uvek daje broj; `"IZVOD"` fallback uklonjen — `RequireIzvodBroj`). **Prepreka za implementaciju:** `ApplyAvansToFaktura`/`ApplyAvansToOtkup` destruktivno umanjuju `Uplata`/`Isplata` originalnog avans reda i prave split red, pa storno izvoda mora vratiti i te iznose — inače saldo ostaje pogrešan. Zaseban paket sa sopstvenim testovima. | `modBankaMapiranje.bas:65-118, 1590-1650`, `modBankaImport.bas:502, 741-801`, `modNovac.bas:640-675, 1210-1245`, `modStorno.bas` (ResolveNovacForStorno) |
| AUD-048 | P2 | `modStornoWarm` false lifecycle flags: `ScheduleStornoWarm` sets `m_warmScheduled=True` unconditionally under `On Error Resume Next` (believes it's scheduled after an `OnTime` failure); `CancelStornoWarm` sets it False unconditionally. Fix: set the flag only on `Err.Number=0`; `LogErr` on cancel failure. (Late fire is already harmless via the `StornoWarmTick` re-guard.) | `modStornoWarm.bas:51-54`, `:118-124` |
### 8.7 Delta v142 — recalibrated / refuted highlights

- The whole `modTheme` layer (FM-0134, all 7 rows tagged P1) touches **only colors** — `DisableField`/
  `DisableCombo` clearing the value is the intended mode-switch behavior at every call site (no
  value-preserving lock exists), so there is no live data-loss → P3. Cheapest real win: unify the
  storno/danger color standard (#6) after confirming the cream "soft storno" is intentional (v2.24
  "Vrati storno" undo treats storno as recoverable).
- `modMouseWheel` (FM-0140) P0/P1 tags are inflated for a cosmetic off-by-default scrollbar hook; the
  only real hardening is checking `UnhookWindowsHookEx` and not leaking the handle (P2, S).
- Banka parsers (FM-0128..0132) are recalibrated P2/P3: the shared 4-level import integrity turns any
  offset/shape drift into an import abort, not silent corruption. Residual the gate can't see: 0/0
  account-only rows (money-zero). Shared date/account/poziv fixes belong to RF-16.
- `frmMarza` (FM-0106) is legacy/unused → all rows Accepted/P3; no business risk unless revived.
- Refuted side-detail: FM-0137 #10 (`modStornoWarm` shutdown) — `Workbook_BeforeClose` already calls
  `ShutdownApp`→`StopStornoWarm`, so the "add to BeforeClose" recommendation is redundant.
