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
| AUD-001 | **RESENO (RF-14).** `ParseValuesJson` je prepisan u jedan stateful tokenizer nad CELIM dokumentom (`TryScanJsonDocument`): uklonjen globalni `Replace(", ", ",")` koji je brisao razmak u svakoj tekst-celiji, `\"` vise ne lomi quote tracking, `\uXXXX` se dekoduje (ChrW, izvor ostaje ASCII), red se ne cepa literalnim `Split(block, "],[")`. Uz to je parser **fail-closed**: skracen/neuravnotezen JSON, nevalidan escape i `values` koji nije niz vracaju gresku umesto parcijalnih redova. Novi seam-ovi `TryParseValuesJson`/`TryReadSheetData`/`TryGetSpreadsheetID`/`TryGetOrCreateSpreadsheetID` razdvajaju „prazno" od „greska" na svim putanjama koje posle citanja mutiraju podatke, generisu identitet ili rade get-or-create. Testovi: `RunSheetsJsonParserTests` (offline, 18 grupa) + integracioni u `RunMasterSyncSmokeSuite`. | `modGoogleSheets` `ParseValuesJson`/`TryScanJsonDocument`, `modMasterSync`, `modStanicaLock`, `modBrojevi`, `modStammdatenSync`, `modGoogleSyncOrchestrator`, `modProductionHealthCheck` | Fixed |
| AUD-002 | **RESENO (RF-14).** Uklonjena spoljna transakcija oko celog OTK batch-a u `ImportOtkupFromPWA_TX` (isti model kao `ImportZbirneFromPWA_TX` za VOZ); red-atomicnost ostaje na `ImportRowToTblOtkup_RowTX`. Pad kasnijeg sheeta vise ne ponistava vec uvezene i na Google-u ack-ovane redove ranijih sheetova. Test: `Test_MasterSyncCrossSheetFailureKeepsEarlierRows`. | `modMasterSync.bas` `ImportOtkupFromPWA_TX` / `ImportOtkupSheetLoop` | Fixed |
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
| AUD-009 | **RESENO (RF-05).** `FillOpenFakture` je prebacen na centralni read-model `modNovac.GetOpenFakture` (izbacuje stornirane, trazi `Neplaceno`, samo preostalo > 0; dopunjen kolonom `Datum` za prikaz) — forma vise ne duplira filter, pa test koji gadja read-model pada ako se filter ukloni. Posledica: fakture sa nestandardnim statusom ili bez preostalog duga vise nisu u listi. Test: `Test_OpenFaktureExcludeStornirano`. | `frmDokumenta.frm` FillOpenFakture, `modNovac.bas:690+` | FM-0018 #2 |
| AUD-010 | `ApplyAvansToOtkup/Faktura`: no target-active (storno) and no target-owner check; no-op returns `True` (inflates batch counters in frmBankaExportPregled). | `modNovac.bas:1067-1204, 487-550` | FM-0019 #4/#5/#6/#11 |
| AUD-011 | `CreateFaktura`: no prijemnica-kupac ownership check; takes `rows(1)` without `Count=1` guard (extends ROADMAP §2.7). | `modFaktura.bas` validation loop | FM-0034 #1/#2 |
| AUD-012 | **RESENO (RF-07).** `StampajReversAmbDok` je dobio (a) `ExcludeStornirano(d, TBL_AMBALAZA)` pre rekonstrukcije — stornirane noge vise ne ulaze u kolicinu na reversu; (b) **tip ambalaze kao deo kljuca** preko novog seam-a `modIzvestaj.ReversRedPripada` (DokumentID + DokumentTip + TipAmbalaze, trim + case-insensitive). Ranije se `tipAmb` uzimao sa PRVOG reda dokumenta dok su se kolicine sabirale preko SVIH tipova, pa je dokument sa 25 letvarica + 15 plasticnih davao jedan revers na 40 „letvarica". Tip stize sa IZABRANOG reda pregleda (`KarticaDetalji_CurrentAmbTip`, nov getter uz postojeci `KarticaDetalji_CurrentAmbDok`) — a red pregleda je vec grupisan po `DokumentID|TipAmbalaze` u `ReportAmbalazePojedinacni`, pa je poklapanje 1:1. Bez prosledjenog tipa (defanzivni fallback) uzima se tip prvog reda dokumenta i sabira se **samo** taj tip — nikad mesavina. (c) FM-0029 #16: vise od jedne noge po strani za isti tip (duplikat / vise generacija) se sada **prijavljuje operateru** (`vbYesNo`, odustajanje prekida stampu) umesto tihog sabiranja. Testovi: `T_ReversKljucPoTipu` u `RunIzvestajTests`. **Ostaje otvoreno (FM-0029 #17):** dokument bez datuma i dalje tiho dobija `Date` umesto greske.  **Review dopuna:** normalizacija tipa je bila **razlicita na dve putanje** — pregled je grupisao po sirovom `dokIDv & "|" & tipv`, a matcher poredio `Trim + vbTextCompare`, pa su „Letvarica" i „letvarica" davali DVA reda pregleda dok je svaki revers sabirao OBA (tiho vracanje istog mesanja). Uveden je jedan kanonski kljuc `modIzvestaj.AmbTipKljuc` (`UCase$(Trim$(...))`) koji koriste OBE putanje — `ReportAmbalazePojedinacni` grupisanje i `ReversRedPripada` match; DokumentID se na obe strane trimuje. Testovi: `AmbTipKljuc` asserti u `T_ReversKljucPoTipu`. **Review dopuna 2 (P1, deterministicki):** kanonizacija tipa nije bila dovoljna — `ReversRedPripada` definise identitet kao `DokumentID + DokumentTip + TipAmbalaze`, dok je `ReportAmbalazePojedinacni` grupisao SAMO po `DokumentID + TipAmbalaze` (nedostajao `DokumentTip`; zateceno stanje, ne regresija RF-07 — ali putanje su i dalje bile nesaglasne). To NIJE teorijski sudar: `modOtkup.SaveOtkup` na NORMALNOJ putanji upisuje **isti `otkupID` i isti `tipAmb`** pod DVA tipa dokumenta — primljene pune gajbe kao `DOK_TIP_OTKUP` (`modOtkup.bas:600/603`), izdate prazne kao `DOK_TIP_OM_IZLAZ_KOOP` (`:613/616`). Otkup koji primi 20 i izda 8 istih gajbica je time davao JEDAN red pregleda koji nosi tip PRVOG zapisa (`Otkup`), skriveni ref-kljuc `AMB|Otkup|<id>`, pa je „Stampaj dokument" uvek rutirao na `ReprintOtkupniListByOtkupID` — **revers `OM-Izlaz-Koop` nije imao svoj red i bio je nedostupan za stampu iz pregleda**. Grupni kljuc je sada pun: `Trim$(dokTipv) & "|" & Trim$(dokIDv) & "|" & AmbTipKljuc(tipv)`. Vidljiva promena: uz-otkup sa izdatom ambalazom daje DVA reda (Ulaz na `Otkup`, Izlaz na `OM-Izlaz-Koop`), svaki sa svojim ref-kljucem i svojom stampom; UKUPNO se ne menja. Gate: `T_E2E_AmbPregledRazdvajaTipDokumenta` (isti DokumentID + tip ambalaze, dva tipa dokumenta -> 2 reda + UKUPNO, dva RAZLICITA ref-kljuca, kolicine ostaju na svom dokumentu). **Cleanup (review 3):** storno provera je prebacena sa `ExcludeStornirano` na inline `AmbRedStorniran` sa IDENTICNIM pravilom (`FilterArray` `"<>"` `"Da"` -> `CStr` poredjenje, bez trima; kolona koje nema = nema storna) -- u one-shot print putanji je puna filtrirana kopija celog `tblAmbalaza` bila cist visak, jer petlja ionako prolazi ceo ledger. Ista odluka koju `ReportAmbalaza` vec nosi i dokumentuje za istu tabelu.| `frmIzvestaj.frm` `StampajReversAmbDok`, `modIzvestaj.ReversRedPripada`, `modKarticaDetalji.bas` | FM-0029 #4/#5/#16 |
| AUD-013 | `MatchesFilter` returns True for unknown operator — a typo silently removes the criterion; carries the whole `ExcludeStornirano` layer. **RF-06 provera (ostaje otvoreno, ne siri se scope):** prebrojani su SVI `clsFilterParam.Init` pozivi u `src-vba/` — svaki operator je literal iz podrzanog skupa (`"="`, `"<>"`, `"BETWEEN"`; `">="`/`"<="`/`"LIKE"` nemaju nijednog pozivaoca). Nijedna report putanja ne zavisi od `Case Else`, pa je grana danas **nedostizna u produkciji** (latentna zamka za nov kod, ne aktivan bug u izvestajima). Fix (`Case Else -> False` + log) pripada svom paketu jer dira ceo `ExcludeStornirano` sloj. | `modArrayUtils.bas:94-95` | FM-0027 #1 |
| AUD-014 | Opening `frmBankaImport` books money (`AutoMapStrongKeysBankaImport_TX` on Activate under `On Error Resume Next`); manual Kupac mapping always becomes avans (empty `FakturaID`); blok preview reads a different source than the command. | `frmBankaImport.frm:70-74, 332, 340-346 vs 595` | FM-0024 #1/#2/#3 |
| AUD-015 | Dashboard KPIs count stornirane rows; frmDokumenta twin excludes them — screens disagree after storno. | `frmOtkupAPP.frm:1356-1422` | Review |
| AUD-016 | Duplicate test modules shipped twice (~1,430 lines) — E2E gate `Application.Run` resolves ambiguously; dead `modBankaImportParserClipboard` (555 lines). | `modE2EReleaseGate.bas:34-75` | Review |
| AUD-017 | Startup/lifecycle trio: outer startup EH clears `Err` before logging; failed startup backup aborts remaining startup; pending autosave `OnTime` survives close (ghost reopen). | `ThisWorkbook.doccls:39-52`, `modMain.bas:91`, `modJournaling.bas:527` | Review |
| AUD-018 | Infra drift set: `GetConfigValue` exact vs `SetConfigValue` trim; legacy `tblConfig` still required; `modMonitoring` `ActiveWorkbook` fallback; ~~`FindVOZSheets` no pagination~~; `BulkPushPendingForStanica` bypasses `UpdateCell`; `AppendRow` leaves phantom row on mid-write failure. **Delimicno RESENO (RF-14):** `FindVOZSheets` je dobio `nextPageToken` petlju (deljeni `ExtractNextPageToken` sa `FindOTKSheets`) — preko 100 VOZ sheetova se vise ne gubi tiho; test `Test_VOZSheetListingPagination`. Ostale stavke ostaju otvorene. | see catalog Blok A | Both (paginacija Fixed) |
| AUD-019 | Plaintext PWA PINs (all roles) and kooperant JMBG exported to Google Sheets while `ExportConfig` already filters credentials. | `modStammdatenSync.bas:2186-2379, 1327-1381` | Review |
| AUD-021 | `StornoNovac` does not refresh `UpdateOtkupStatus` for the linked otkup — storno of a payment permanently hides the debt from the payout list (`GetOpenOtkupi` trusts stale `Isplaceno`). | `modStorno.bas:781-785`, `modNovac.bas:974/1001` | FM-0019 #16, FM-0021 #10 |
| AUD-022 | **RESENO (RF-05).** Sve cetiri stavke: (a) Kl.II u izvoru blokira snimanje zbirne kad je cekboks iskljucen (`ZbirnaIzvorImaKlasuII` + labela + hard-stop u `btnUnosZbr_Click`); (b) smer ambalaze je obavezan uz kolicinu — UI blokada + core guard (`Case Else` u `SaveOMUlaz_TX` vise ne knjizi legacy `Stanica ULAZ` nego odbija nepoznat smer); (c) pad malina auto-zbirne je vidljiv (hvata se i povrat i `Err`); (d) prefill uzima poslednju generaciju preko nove kolone `GeneracijaID` + anchor na `OldDocID` iz correction context-a, a Kl.I i Kl.II dolaze iz ISTE generacije. Testovi: `Test_ZbirnaKlasaIIGuard`, `Test_OMUlazSmerObavezan`, `Test_MalinaAutoZbirnaFailSignal`, `Test_PrefillBiraPoslednjuGeneraciju`, `Test_GeneracijaIDNaSavePutanji`, `Test_GeneracijaNePrelaziVlasnika`. | `frmDokumenta.frm`, `modDokumenta.bas` (generacija/prefill/OM ulaz), `modSetup.EnsureSledljivostSchema` | FM-0018 #3/#4/#5/#8 |
| AUD-023 | **RESENO (RF-06).** Svih pet stavki: (a) `ReportSaldoOM` isplate pripisuje stanici po `COL_NOV_OM_ID` REDA (isti kljuc koji `ReportIsplata("OM")` vec koristi), a ne po maticnoj stanici kooperanta -- isplata sa jednog OM-a vise ne ulazi u izvestaj drugog; red bez `OMID`-a pada na maticnu stanicu da novac ne nestane iz svih izvestaja (pokriva i #9). (b) `ReportKarticaKooperanta` i `ReportKarticaAmbalaze` sabiraju promet pre `datumOd` u red `POCETNO STANJE`; kolona Saldo je sada ZAVRSNI saldo, a UKUPNO zaduzenje/razduzenje ostaje promet perioda. (c) Nedostajuca prijemnica daje prazne brojke + oznaku `nema prijema` u OBA izvestaja (`ReportOtkupRobaOM` je davao 0/0.00%, `ReportManjak` 100%), i takvi redovi ne ulaze u UKUPNO manjak ni u osnovicu procenta. **Uz to (review nalaz): join zbirna<->prijemnica je bio po samom `BrojZbirne`**, a AUD-052 je vec dokazao da poslovni broj nije identitet -- dve aktivne zbirne istog broja (razlicit vozac/kupac) su sabirale tudju prijemnu kolicinu. `BuildManjakDict` je sada scoped na **stavku** dokumenta: `ZbirnaVlasnikKljuc` = `broj|vozac|kupac` (**ista definicija vlasnika koju koristi `modStorno.RequireJedanVlasnikPoBroju`**) + `ZbirnaStavkaKljuc` dodaje **`Klasa`**. Klasa je nuzna jer auto-lanac hladnjace vodi Klasu I i II kroz ceo lanac odvojeno (zasebna otpremnica/zbirna/prijemnica) uz isti broj, vozaca i kupca -- bez nje se prijem obe klase sabere pa taj zbir dodeli SVAKOJ klasi (u malina modu UKUPNO prijem 2x stvarni). To je ista granularnost koju `modAutoHladnjaca.KeyZbrKlasa` vec koristi za idempotenciju; `#V|` i dalje broji VLASNIKE bez klase, pa dve klase istog dokumenta ne obaraju fail-closed granu; `ReportManjak` vise ne duplira agregaciju prijemnica (zadrzava JEDAN red po dokumentu, ali prijem sabira po klasama koje red obuhvata), a `ReportOtkupRobaOM` razresava vlasnika preko vozaca otpremnice i racuna srazmeru UNUTAR klase. Kad se vlasnik ne moze dokazati (ili postoji prijemnica bez kompletnog vlasnika na broju sa vise vlasnika) red je **fail-closed**: oznaka `nejasan vlasnik`, bez brojke i van UKUPNO. Broj sa tacno jednim vlasnikom i dalje koristi agregat po broju, pa starije prijemnice bez popunjenog vlasnika ne regresiraju. (d) `modNovac` deli uplatu po fakturi SRAZMERNO stavkama (`BuildFakturaVrstaUdeoCache` + `RaspodeliPoUdelima`), umesto cele uplate na vrstu prve stavke; **raspodela se racuna u celim parama metodom najvecih ostataka** (largest remainder): svaki kljuc dobije floor svog idealnog udela, pa visak para ide redom na najveci ostatak. Time se drze OBE invarijante odjednom -- zbir delova je tacno `ZaokruziNovac(iznos)` (inace bi 100/3 dalo prikaz `33,33 x 3 = 99,99` uz UKUPNO `100,00`) **i nijedan deo nije negativan** (raniji oblik je poslednjem kljucu davao ostatak posle zaokruzivanja, pa je npr. 0,03 na 5 vrsta davalo poslednjoj vrsti `-0,01`; clamp na nulu bi razbio prvu invarijantu). Zaokruzivanje je half-up (`ZaokruziNovac`) jer je VBA `Round` banker's. (e) `ReportProsecnaCena`/`ReportManjak`/`ReportAmbalaza`/`ReportOtkupRoba` dobili eksplicitan `Select Case` sa `Case Else -> Empty` -- zbirni mod za Kooperante/Vozace vise ne dobija globalni izvestaj pod tudjim naslovom (#12/#13/#14). Uz to #10: `AGROHEMIJA (nerasporedjena)` ne ulazi u UKUPNO stanice (`tblMagacin` nema kolonu stanice, pa se isti iznos ponavljao po svakoj stanici). Testovi: `RunIzvestajTests` (`modIzvestajTests`) -- **tvrd gate** (podize gresku kad assert padne, po konvenciji uvedenoj u RF-14; ranije je samo stampao FAIL, pa je automatizovani pozivalac mogao da ga prijavi kao PASS). Pokriva ciste seam-ove `NovacRedPripadaStanici` / `KarticaRezultatSaPocetnim` / `KarticaAmbRezultatSaPocetnim` / `ManjakStavka` / `PrijemZaZbirnu` / `RaspodeliPoUdelima` + dispatch matricu, **plus tri end-to-end testa nad tabelama** (dve zbirne istog broja kod dva kupca sa razlicitim prijemnim kolicinama, iz ugla `ReportManjak` i `ReportOtkupRobaOM`; i Klasa I+II istog dokumenta -- 1000/900 i 200/150 -- gde svaka klasa mora dobiti SVOJ prijem a UKUPNO ostati 1050) -- seam testovi ne mogu da uhvate gresku u samom table-join-u. E2E rade u `clsTransaction` sa snapshot-om i uvek se rollback-uju. | `modIzvestaj.bas` (seam-ovi + `Report*`), `modHelpers.bas` `BuildManjakDict` + `ZbirnaVlasnikKljuc`/`ZbirnaStavkaKljuc`/`KlasaOrDefault`, `modAutoHladnjaca.bas` (`ClassOrDefault` delegira), `modNovac.bas` `GetUplataByVrsta` + `ZaokruziNovac`, `frmIzvestaj.frm` (prikaz oznake) | FM-0028 #1/#3/#5/#6/#9/#10/#12/#13/#14 + AUD-052 posledica u report sloju |
| AUD-024 | **RESENO (RF-07).** Sve tri stavke. (a) **Freshness (#1/#19):** `UpdateStatusLabel`, naslov `PrintIzvestaj`-a i PDF kartice (`PrintKarticaPDF`/`PrintKarticaAmbalazePDF`) citaju period iz `m_curOd`/`m_curDo` (kroz novi `PrikazanPeriod()`), nikad vise iz `txtDatum*` — stara lista ne moze da dobije nov period u zaglavlju. Uz to `txtDatumOd_Change`/`txtDatumDo_Change` postavljaju `m_periodDirty`, pa status odmah pise **„NIJE OSVEZENO — kliknite 'Prikazi'"** (uz period podataka koji su jos na ekranu) sve dok se izvestaj ne regenerise; `btnUnos_Click` cisti flag. Stampa pre ijednog generisanja je blokirana porukom (nema perioda za zaglavlje). (b) **Tihi `CleanFail` (#2):** greska lazy generisanja se pamti (`m_genFailMsg`/`m_genFailTab`), tab ostaje van `m_genTabs` (pokusace ponovo), a po izlasku iz handlera (posle `EndTableCache` + `ScreenUpdating=True`, ne pod `Resume`) `ShowGenFailure` **brise aktivnu listu**, cisti panel detalja, boji status crveno i prikazuje `MsgBox` sa opisom greske — umesto stare liste uz zelen status „N redova ucitano". (c) **Nevalidne kombinacije (#3):** vidljivost tabova vodi nova matrica `modIzvestaj.IzvestajTabDostupan(entitetTip, zbirni, pageIdx)` koja prati stvarni dispatch `Report*` funkcija; zbirni tabovi 5/6/7 vise nisu vidljivi svima (Vozac nema „Prosecna cena" jer `ReportProsecnaCena` nema vozacku granu; Kooperant nema nijedan zbirni tab). Kad kombinacija nema nijedan tab, status nosi vidljivu poruku umesto prazne forme bez objasnjenja. Mapiranje UI labele u kod entiteta je izdvojeno u `IzvestajEntitetKod` (bio dupliran `Select Case` u `btnUnos_Click`). #14 („0% manjka" za pending prijemnicu) je zatvoren jos u RF-06 oznakom `nema prijema`. Testovi: `T_TabMatrica` (28 provera, obuhvata i pojedinacni rezim da se matrica ne suzi) + `T_EntitetKod` u `RunIzvestajTests`; freshness i CleanFail su operater-smoke (UI stanje). **Ostaje otvoreno:** FM-0029 #11 (status broji UKUPNO/summary redove), #26 (redosled perioda nevalidiran), #9 (prazan `entitetID` neblokiran) — svi P2/P3.  **Review dopuna (isti paket):** (d) **Promena konteksta bez dostupnog entiteta.** `AutoRefresh` izlazi odmah kad novi entitet nema nijedan izbor (`cmbEntitet.ListIndex < 0`), pa je npr. prelaz sa „Kupac A / Ambalaza" na „Vozaci" bez aktivnog vozaca ostavljao **redove Kupca A** na ekranu, `m_genTabs` ih je vodio kao generisane, status je bio zelen, a stampa bi ih odstampala pod naslovom „Vozaci" — ista klasa greske kao (a), samo sto je period ispravan. Uveden je centralni `InvalidateReportContext` na **pocetku** `AutoRefresh`-a (jedini ulaz za svaku promenu entiteta/rezima/combo izbora, pa se poziva pre guarda): prazni `m_genTabs`, resetuje `m_hasGenerated`/`m_curTip`/`m_curID`/`m_curEntLabel`/`m_curEntName`, cisti SVE liste (`ClearAllReportLists`) i panel detalja, i osvezava status. Uz to je uveden print guard `AktivanTabGenerisan()` (`Not m_noValidTabs And m_hasGenerated And m_genTabs.Exists(CStr(mpReports.value))`) na obe stampe, a naslov stampe sada nosi **zapamceni identitet** (`m_curEntLabel`/`m_curEntName`), po istom principu po kom period nosi `m_curOd/m_curDo` — ranije je citao zivi `GetActiveEntitetTip()`/`cmbEntitet.value`. Isti zapamceni identitet (`m_curTip`) bira i skup zaglavlja kolona. Status dobio granu `Not m_hasGenerated` -> „Prvo kliknite 'Prikazi'" (prazna lista posle invalidacije NIJE rezultat upita, pa ne sme da pise „Nema podataka za izabran filter"). Vidljivost panela „Detalji" izdvojena u `SyncDetaljiVisibility` (deli je `mpReports_Change` i invalidacija — ciscenje lista okida `Change` evente koji panel mogu da pokazu na tudjem tabu). (e) **Matrica se razilazila sa core-om na jednoj kombinaciji:** `Kupac + Zbirni + Prosecna cena` je bio ponudjen, ali `ReportProsecnaCena("Kupac", "")` ide kroz `GetPrijemniceByKupac` koji **bezuslovno** dodaje filter `KupacID = entitetID` — zbirni rezim salje `""`, pa upit trazi prijemnice BEZ kupca i vraca prazno i kad prijemnice postoje. (`OM` nema taj problem: grana `Case "OM", ""` eksplicitno hvata prazan ID kao „svi".) Tab je **sklonjen** iz matrice. **Otvoreno poslovno pitanje (nije bug):** globalni prosek cene preko SVIH kupaca nije implementiran — to je nov izvestaj (koja cena, koje grupisanje), ne UI podesavanje; ako se odluci da treba, uklanja se `KupacID` filter za prazan ID i tab se vraca u matricu. Gate: `T_E2E_ProsecnaCenaZbirniKupac` seed-uje prijemnice **dva razlicita kupca** u sentinel prozoru i tvrdi da pojedinacno radi a zbirno vraca `Empty` — pa **pada u oba smera**: i ako matrica ponovo ponudi prazan tab, i ako core pocne da vraca podatke a tab ostane skriven.| `frmIzvestaj.frm` (`UpdateStatusLabel`/`ActiveReportList`/`PrikazanPeriod`/`UpdateReportMode`/`GenerateActivePage`/`ShowGenFailure`/`txtDatum*_Change`), `modIzvestaj.bas` (`IzvestajTabDostupan`/`IzvestajEntitetKod` + `IZV_TAB_*`), `modPoruke.bas` | FM-0029 #1/#2/#3/#14/#19 |
| AUD-025 | Banka set: import dedupe key omits account number (multi-account transaction silently dropped); block resolver crashes on 3+ candidates (`ReDim 1 To 2` → subscript error rolls back whole AutoMapAll batch); manual mapping never checks direction (reachable from UI). | `modBankaImport.bas:783-798`, `modBankaMapiranje.bas:1457-1497`, Map* funkcije | FM-0022 #1, FM-0023 #8/#14 |
| AUD-026 | Banka export set: stale per-blok override survives reload without clamp and `GenerisiNalogeCSV` performs no final saldo revalidation — CSV can order more than the currently open amount. | `frmBankaExportPregled.frm:504-535, 967`, `modBankaExportPregled.bas:369-374` | FM-0020 #1/#2, FM-0021 #1 |
| AUD-027 | **DELIMICNO RESENO (RF-07 — samo cross-tab print).** Print target panela „Detalji" (`mCurOtkupID`/`mCurOtpremnicaID`/`mCurAmbDokID`) je modul-stanje u `modKarticaDetalji`, nevezano za tab/listu/red, pa je red izabran na Kartici prezivljavao prelaz na Ambalazu i „Stampaj dokument" je stampao dokument SA DRUGOG taba. `mpReports_Change` sada zove postojeci `KarticaDetalji_Clear` na svaku promenu taba (isti helper koji `GenerateAmbalazeReport` vec koristi), pa target ne prezivljava tab-switch; `ShowGenFailure` ga cisti i posle neuspelog generisanja. Time je zatvoren i FM-0029 #15 (deljeni detail state sa drugog taba). **Ostaju otvorene dve stavke — RF-08:** reprint storniranog otkupa kroz raw fallback (`modPrint.bas:334/354`, FM-0031 #3) i `FillFakturaSablon` cleanup bez `.UnMerge` (`modPrint.bas:1896-1924`, FM-0031 #19). | `modPrint.bas:334/354, 1896-1924` (otvoreno), `frmIzvestaj.mpReports_Change` + `modKarticaDetalji.KarticaDetalji_Clear` (reseno) | FM-0031 #3/#19 (otvoreno), FM-0030 #1 + FM-0029 #15 (reseno) |
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
| AUD-041 | P1 | Duplicate document-number generation: `GenerateBrojPrijemnice` error handler returns a valid-looking `1/ddmmyy` instead of hard-failing; `modMasterSync.GenerateBrojZbirne` is a parallel **row-count** generator (`seq=count+1`) that produces a duplicate on gaps (`1/ddmmyy` + `…-3` → `-3` again) instead of the canonical `SuggestNextBroj`/`MaxSeqFromTable`. Fix: EH → `""`; delegate ZBR generation to `SuggestNextBroj(KIND_ZBR,…)`. **Fixed under RF-28:** `GenerateBrojPrijemnice` EH → `""` (aligned with `GenerateBrojDokumenta`/`GenerateBrojOtpremnice`); `GenerateBrojZbirne` body is now `SuggestNextBroj(KIND_ZBR, vozacID, datum, checkRemote:=False)` (MAX-seq + `ApplyMirrorPrefix` + `BrojZbirneExists` bump-loop), keeping only the "vozacID has no digits → `''`" guard. **Residual (accepted):** the ZBR fallback now inherits the `IsAutoBrojDokumenta` toggle — with auto-numbering OFF a PWA row that arrives with an empty `BrojZbirne` fails as `SyncError` instead of getting a locally generated number (loud, recoverable; PWA-generated numbers are unaffected). | `modBrojevi.bas` (`GenerateBrojPrijemnice` EH), `modMasterSync.bas` (`GenerateBrojZbirne`) |
| AUD-042 | P1 | MasterSync wrong-write cluster: `TryUpdateVozacID` returns True even when `UpdateCell` fails → GS marked `Synced>Master` with empty `VozacID`; an invalid date silently becomes **today** on both paths (OTK/VOZ); a failed header write leaves a poison spreadsheet that the next run treats as existing. Fix: check the Boolean write result; strict date parse → `SyncError`; trash/temp-name the sheet until the header write succeeds. **Fixed under RF-28:** (a) `TryUpdateVozacID` no longer returns a Boolean — it returns a classified outcome (`UPDATED`/`NOCHANGE`/`CONFLICT`/`FAILED`/`NOTFOUND`) plus a detail string, because the old single `False` merged "nothing to change" with "write failed" and "row missing". Only `NOCHANGE` still maps to the terminal `Duplicate` status; `FAILED`/`CONFLICT`/`NOTFOUND` now write `SyncError`, increment `outErrors` **and** set the fatal flag, so the top level ends in `Monitor_MasterSyncFail` instead of a green sync with an empty `VozacID`. A sheet carrying a *different* `VozacID` than master is treated as a data conflict (no write) rather than an ordinary duplicate. (b) new `IsParsableMasterSyncDate` (empty/text/time-only/pre-2000 rejected) is checked in `ValidatePWAOtkup` **and** `ValidatePWAZbirna` → the row becomes `SyncError`, and both import bodies (`ImportRowToTblOtkup`/`ImportRowToTblZbirna`) hard-fail instead of falling back to `Date()`; (c) create+header is now one unit (`CreateOTKSheetWithHeader`): on a failed header write the freshly created OTK sheet is sent to Drive trash (`DriveTrashFile`) so the next run cannot count the poison sheet as "existing" (unsuccessful trash is logged as a manual-cleanup instruction). **Fail paths are regression-tested end-to-end.** `UpdateCell`/`WriteSheetData` cannot be made to fail naturally, so `modMasterSync` carries a one-shot test seam (`TestHook_ArmFailSeam` with `VOZAC_WRITE`/`OTK_HEADER`, armed and consumed on first use — same pattern as `modAutoHladnjaca.mTestFailStep`, empty in production). Local suite covers the `FAILED` classification and that the seam is one-shot; `RunMasterSyncSmokeSuite` covers the real end-to-end: `Test_MasterSyncVozacUpdateFailureIsFatal` (real sheet → `ImportOneOTKSheet` → row written back as `SyncError`, `errors>=1`, fatal flag set) and `Test_MasterSyncPoisonSheetIsTrashed` (header fails → Drive lookup can no longer find the sheet, since the query filters `trashed=false`). | `modMasterSync.bas` (`TryUpdateVozacID`, `ValidatePWAOtkup`, `ValidatePWAZbirna`, `ImportRowToTblOtkup`, `ImportRowToTblZbirna`, `CreateOTKSheetWithHeader`), `modBusinessFlowProTests.bas`, `modGoogleSyncSmokeTests.bas` |
| AUD-043 | P1 | MasterSync document-integrity cluster: auto-otpremnica groups by `Stanica\|Datum\|Vozac\|Klasa`, mixing vrste/cene/ambalaza into one otpremnica; VOZ `LinkZbirnaToOtkupAndOtpremnica` links by CRID without membership check (same vozač/datum/not-already-linked) and overwrites existing links with no "empty-or-identical" conflict guard. Fix: add `VrstaVoca\|SortaVoca\|Cena\|TipAmbalaze` to the key; validate membership; guard link writes. **Fixed under RF-28:** (a) the group key is now `Stanica\|Datum\|Vozac\|Klasa\|VrstaVoca\|SortaVoca\|Cena\|TipAmbalaze` — the new segments are appended, so `parts(0..3)` keep their meaning and "metadata from the first row" is safe because every row of a group is identical on those fields; (b) new `RequireBrojZbirneNotConflicting` (write only when the field is empty or already identical, otherwise raise) guards both `LinkZbirnaToOtkupAndOtpremnica` and `LinkOtpremnicaToBrojZbirneStrict`, plus a membership check against the zbirna's vozač and business day. The zbirna is resolved by **primary key** — `ImportVOZRow_RowTX` now passes `outZbirnaID` and the link does `RequireSingleMasterSyncRow(TBL_ZBIRNA, COL_ZBR_ID, …)`, reading vozač/datum off that row (plus a sanity check that the row really carries the number being linked). The earlier `LookupValue(…, COL_ZBR_BROJ, …)` was first-match, so on a `BrojZbirne` collision — a documented multi-device risk — membership was validated against the *older, unrelated* zbirna while the new row sits at the end of the table. The day check is a real guard with a bounded window (`MASTER_SYNC_MEMBERSHIP_DAY_TOLERANCE = 1`): same day passes, the adjacent day passes with a `LogWarn` (legitimate post-midnight loading), anything further raises and rolls back the row — an unreadable date also raises rather than being treated as equal. **Open assumption:** the ±1 window assumes a zbirna is a single-day transport document (plus post-midnight loading). Not yet confirmed with the domain owner; if one zbirna may legitimately span several days, this rejects valid imports as `SyncError` — the fix is the single `MASTER_SYNC_MEMBERSHIP_DAY_TOLERANCE` constant (raise it, or drop the day check to a `LogWarn` and keep vozač as the only hard guard). An empty otkup `VozacID` is still allowed (Otprema tab not yet synced); a different non-empty one raises. Regression-guarded by the RF-28 tests in `modBusinessFlowProTests`: the grouping test varies **each** key segment separately against one baseline row (Cena, VrstaVoca, SortaVoca, TipAmbalaze → five combinations must yield five otpremnice, each carrying its own value, so a regression in a single segment cannot hide), plus link conflict, PK-vs-collision membership and the day window. To keep the auto-otpremnica test from sweeping unrelated rows, `AutoCreateOtpremniceFromPWA[_TX]` takes an optional `samoDatum` scope — same pattern as the existing `samoBrojOtp` on `AutoCreateZbirnaFromOtpremnice`; production callers pass nothing and behave as before. | `modMasterSync.bas` (`AutoCreateOtpremniceFromPWA`, `ImportVOZRow_RowTX`, `LinkZbirnaToOtkupAndOtpremnica`, `LinkOtpremnicaToBrojZbirneStrict`, `RequireBrojZbirneNotConflicting`, `TryMasterSyncDay`), `modBusinessFlowProTests.bas` |
| AUD-044 | P1 | `modIntegritet` false-green: `WriteErr` does not raise `m_totalIssues`, while the overlay title and MsgBox read only that counter → a run can report "0 neusklađenih" alongside GRESKA blocks; an in-memory read failure returns `Empty`, which is indistinguishable from PASS. Fix: `WriteErr` increments an `ErrorCount`; MsgBox shows INCOMPLETE when errors>0; typed `IntegrityRunResult` so `Empty` ≠ PASS. | `modIntegritet.bas:1304-1310`, `:84-85`, `:59`, `:90` |
| AUD-045 | P1 | Sledljivost incomplete trace shown as complete: `TraceByZbirna` filters by the helper `OtpremnicaID` instead of canonical `tblOtkup.BrojZbirne`, and normalizes the number inconsistently (auto-link `UCase$+Trim$` vs raw trace compare) → an otkup with `BrojZbirne` but empty `OtpremnicaID` drops out; `frmSledljivost`/printed PDF present the partial trace as complete (no incompleteness marker). No corruption (manual write goes through `ReassignOtkupToOtpremnica_TX`). Fix: direct `BrojZbirne` pass with `vbTextCompare`; typed trace result with `IsComplete`. | `modSledljivost.bas:540-544`, `:282` vs `:464`, `frmSledljivost.frm:462/499/543` |
| AUD-046 | P1 | Station-mirror missing-shadow: `modMalina.EnsureVozacMirrorForStanica` doesn't confirm the station exists, its `AppendRow=0`/EH swallow failures, and `modMasterSync.StampVozacFromStanicaForMalina` + `modAutoHladnjaca` unconditionally set `vozacID=stanicaID` → a document gets an FK with no `tblVozaci` row. Fix: one canonical `IsManagedStationMirror(id)` checking the `tblStanice`+`tblVozaci` pair; re-raise in Ensure EH; verify the mirror before stamping. **Fixed under RF-28:** `modMalina.IsManagedStationMirror` is the single validator (station row **and** vozač row with the same ID; deliberately a different question from `modBrojevi.IsStanicaMirrorVozac`, which only decides the `S` prefix). `EnsureVozacMirrorForStanica` now requires the station to exist and be unique (`FindRows` count), treats `AppendRow=0` as an error, and re-raises from EH like `BackfillVozacMirrorsForMalina`; an inactive station is a `LogWarn`, not a block (backfill iterates all stations). Both stampers ask before writing: `StampVozacFromStanicaForMalina` tries `Ensure` and skips the row with a `LogWarn` if no pair results (locally caught so one bad station can't roll back the whole `_TX`), and `AutoChainHladnjaca` aborts the chain with an operator warning instead of stamping an uncovered FK — the check runs whenever `vozacID = stanicaID`, not only when the fallback fills an empty `vozacID`, so a caller that passes `stanicaID` in directly cannot bypass the guard. Regression-guarded in `modBusinessFlowProTests.Test_MalinaVozacMirror` (pair check + Ensure re-raise for a nonexistent station + no shadow vozač created). | `modMalina.bas` (`IsManagedStationMirror`, `EnsureVozacMirrorForStanica`), `modMasterSync.bas` (`StampVozacFromStanicaForMalina`), `modAutoHladnjaca.bas` (`AutoChainHladnjaca`), `modBusinessFlowProTests.bas` |
| AUD-047 | P2 | `modProductionHealthCheck` SEF drift + false-green: the SEF status list uses a nonexistent `SEF_CANCELLED` and misses `SEF_REJECTED/SYNC_ERROR/TECH_FAILED` (drift vs `modConfig.bas:659-663`); parent check reports OK after a child FAIL in two spots (Google `:951`, soft-delete helper `:928`). Fix: use `WF_SEF_*` constants + a state matrix; gate the parent on child delta-counters. | `modProductionHealthCheck.bas:871`, `:951`, `:928` |
| AUD-049 | ~~P1~~ RESENO | **Storno celog izvoda** — implementirano u RF-03 grani (`modStorno.StornoIzvod_TX`). Pojedinacni storno izvodnog reda je zabranjen (`ResolveNovacForStorno`/`StornoNovac_TX`), a ceo izvod se stornira sa dva ishoda koja bira operater: **REMAP** (PDF ispravan, mapiranje pogresno → `Obradjeno` = "" , stavke nazad u „za obradu", izvod ostaje uvezen) i **REIMPORT** (PDF los/korumpiran → `Stornirano="Da"`, isti PDF se moze uvesti ponovo jer `IsDuplicateBankaImport` i `GetBankaImportOpen` rade nad `ExcludeStornirano`). Novac i staging padaju u ISTOJ transakciji (inace ponovni uvoz + mapiranje = dvostruko knjizenje). Avans-split naslednici nasledjuju BIM marker roditelja (`BuildAvansSplitNapomena`) pa ostaju vezani za izvod i pod drugim `BrojDokumenta`; identitet je ISKLJUCIVO po markeru (PK `BankaImportID`), nikad po broju/partneru (markerless heuristika uklonjena tokom review-a). Isti broj izvoda na vise racuna trazi „broj/racun". Post-merge review potvrdio bezbednost (fail-closed guardovi, 4-tabelni atomski TX, po-stavci rekonsilijacija uplata/isplata). Testovi: `modTestStorno` T23–T36. | `modStorno.bas` (sekcija IZVOD), `frmDokumenta.frm` (Case "Izvod (ceo)"), `modTestStorno.bas:T23-T27` |
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

### 8.8 RF-03 / RF-05 post-merge review — follow-ups (PR #167/#170, 2026-08)

Two independent code reviews of the merged RF-03 work (RF-03 core A/B/C/D + keš/virman channel
split + storno-izvoda AUD-049) confirmed every core fix correct and complete, the channel split
totals-preserving, and the storno-izvoda feature safe (identity by PK-marker, fail-closed read
guards, 4-table atomic TX, substantive tests T23–T36). Doctrine clean (ASCII, `.frx` untouched,
no new form `WithEvents`, no duplicate `Public`). Process note: the PR was over-scoped (~1252
lines / 3-4 concerns) — keš/virman and storno-izvoda merited their own PRs. Residual follow-ups
(non-blocking, PR already merged):

| ID | Sev | Finding | Location |
|---|---|---|---|
| AUD-050 | P2 | Reconciliation blind spot: `GetIzvodStornoBlokade` accumulates expected amounts only for `Obradjeno="DA"` staging items and skips reconciliation when the expected set is empty, but `CollectIzvodNovacIDs` reverses by marker regardless. A non-"DA" staging row still carrying active marked novac (inconsistent/tampered state) is reversed without an amount check — fails in the safe (un-book) direction, not a double-book, but the guard's reach is narrower than its "za svaku obradjenu stavku" claim. Fix: reconcile any marker with active novac, or assert non-"DA" items carry zero marked novac. | `modStorno.bas:1132/1144` (GetIzvodStornoBlokade) |
| AUD-051 | P3 | Test gap: no TX-level cross-account isolation test. T26 asserts only that the resolver rejects a bare ambiguous number; nothing seeds two accounts each with marked novac, storns one, and asserts the other account's NOVAC is untouched. Code is correct (strict racun filter), but the strongest invariant is unasserted. | `modTestStorno.bas` (add T-case), code `modStorno.bas:1332` |

| AUD-052 | P1 | Storno/ispravka i dalje polaze od **poslovnog broja**, ne od identiteta dokumenta. `RunPrijemnicaCorrection`/`RunZbirnaCorrection` primaju samo `broj`; `ScanPrijemnica` bira jedan PK preko number-only `LookupActiveID`; `StornoPrijemnicaByBroj_TX` / `StornoOtpremnicaBrojAtomic_TX` / `StornoZbirna` skupljaju **sve** aktivne redove tog broja. Broj nije globalno jedinstven (`GenerateBrojPrijemnice` racuna sekvencu po kupcu, x-deo je fiksno „1"), pa dva kupca istog dana dele `1/ddmmyy`. **Mitigacija u RF-05:** `RequireJedanVlasnikPoBroju` (modStorno, Public, fail-closed) odbija storno kad aktivni redovi broja pripadaju vise od jednog vlasnika — greska unutar transakcije, rollback, nijedan red se ne menja. Ukljucen u: `StornoPrijemnicaByBroj_TX`, `StornoOtpremnicaByBroj_TX`, core `StornoZbirna`, `StornoOtpremnicaBrojAtomic_TX` (modStornoFlow) i sve tri **kaskade** — dakle SIMPLE/ISPRAVKA/DUPLI, malina 1:1 i autohladnjaca lanac. Kaskade su poseban slucaj: ulaz im je `BrojZbirne`, a mutiraju `tblOtpremnica`/`tblPrijemnica`, pa provera vlasnika nad `tblZbirna` nije dovoljna. Zato `ResolveZbirnaChainScope` razresava vlasnika lanca (`VozacID`+`KupacID`) **jednom, pre prve mutacije** (prva kaskada obara zbirnu, pa bi kasnije razresavanje videlo „nema parenta"), a kaskade zatim mutiraju iskljucivo redove tog vlasnika. Bez aktivne zbirne uz aktivne nizvodne redove -> fail-closed (pripadnost nedokaziva); bez ijednog aktivnog reda -> idempotentan no-op. Test: `Test_StornoPoBrojuOdbijaDvaVlasnika`, `Test_StornoGuardNaSvimPutanjama`, `Test_StornoGuardUKaskadi`, `Test_StornoKaskadaScopePoLancu` (javni ulaz `StornoOtkupByBrDok_TX`, sve tri kaskade, tudji child pod istim `BrojZbirne`, fail-closed bez parenta). **Ostaje (RF-06+):** puni identitetski storno — `OldDocID → GeneracijaID → svi redovi te generacije` kroz `Scan*`/`Run*Correction`, cime broj prestaje da bude ulaz storna; tada se guard moze relaksirati u obicnu proveru. **Preostali rezidual (do RF-06+):** `StornoOtpremnicaCascade` filtrira po `VozacID` (otpremnica nema `KupacID` kolonu), pa isti-vozac + isti `BrojZbirne` preko dva kupca uz parcijalni storno moze zahvatiti tudju otpremnicu; zbirna/prijemnica (imaju kupca) su precizno scope-ovane. Ograniceno semom otpremnice; nestaje kada kaskade budu keyovane po `GeneracijaID`. | `modStornoFlow.bas` (Scan*/Run*Correction, `StornoOtpremnicaBrojAtomic_TX`), `modStorno.bas` (`StornoZbirna`, `Storno*ByBroj_TX`, `RequireJedanVlasnikPoBroju`) |
| AUD-053 | P3 | Kl.II hard-block guard fails OPEN on error: `ZbirnaIzvorImaKlasuII` EH returns `False`, and its dependency `ValidateZbirnaPreUnosa` swallows column-drift errors returning `sumaKgKlII=0`, so under schema drift in `tblOtpremnica` (KLASA/KOLICINA/KOL_AMB) the AUD-022 Kl.II last-line guard silently disables and Kl.II can again be lost. Narrow window (the write path usually fails on the same drift) but it undercuts a last-line-of-defense guard. Fix: fail-closed — block the save when the guard cannot prove absence of Kl.II. Found in RF-05 post-merge review (PR #170). | `modDokumenta.bas:854` (ZbirnaIzvorImaKlasuII EH), `:826-830` (ValidateZbirnaPreUnosa) |

Nits (no AUD): new user-facing storno strings use inline `ChrW(...)` rather than the `modPoruke`
catalog (matches existing `modStorno` convention; sources stay ASCII) · `CountActiveNovacByBroj`
early-returns `0` on blank input (harmless given callers; worth a clarifying comment) · two new
private numeric helpers (`NumOrZero`/`NzNum`) alongside the existing `Nz` family.

### 8.9 RF-28 post-fix review — systemic finding (2026-08)

The RF-28 MasterSync review surfaced BUG-1 (MED): an EH block that calls `LogErr` **before**
`Err.Raise Err.Number, …`. `LogErr` calls `LogError`, whose first executable line is
`On Error Resume Next` (`modLogError.bas:50`, plus `On Error GoTo 0` at `:90`) — every `On Error`
statement clears the global `Err` object — so the trailing `Err.Raise Err.Number` runs as
`Err.Raise 0`: the original error number/description are gone and, in an `On Error Resume Next`
caller, no non-zero error is seen. RF-28 fixed its own sites (capture `Err.Number`/`Err.Description`
into locals **before** `LogErr`, then re-raise). A codebase-wide scan confirms the same fragile
idiom is not local to MasterSync — it recurs across the SEF, banka and document layers.

| ID | Sev | Finding | Location |
|---|---|---|---|
| AUD-054 | P1 | Systemic defeated re-raise (same class as BUG-1 / AUD-046): **46** EH blocks on `main` run `LogErr`/`LogError` immediately followed by `Err.Raise Err.Number, …`. Because `LogError` (`modLogError.bas:50/90`) executes `On Error Resume Next` / `On Error GoTo 0`, each of which clears the global `Err`, the trailing `Err.Raise Err.Number` executes as `Err.Raise 0` — the error is swallowed and does not propagate (proven by RF-28 `Test_MalinaVozacMirror`, which asserted the re-raise and measured `Err.Number = 0` before the fix). Any caller relying on the re-raise to roll back a TX, mark a sync fatal, abort a batch, or show the failure to the operator instead proceeds as if the operation succeeded. RF-28 (in flight) removes 2 of the 46 (`modMalina.BackfillVozacMirrorsForMalina`, `modMasterSync.LinkZbirnaToOtkupAndOtpremnica`) and adds a correct re-raise to `modMalina.EnsureVozacMirrorForStanica`; **44 remain**, ~80% in the SEF (e-invoice / fiscal) layer. **Fix — mechanical, identical to BUG-1:** capture `errNum = Err.Number` / `errDesc = Err.Description` into locals before `LogErr`, then `Err.Raise errNum, SRC, errDesc` (with a fallback code when `errNum = 0`). **Severity is per-site:** P1 core is SEF + banka (a swallowed submit/validate/persist error can mark a fiscal or bank document processed when it actually failed); pure top-of-stack log-and-show sites (e.g. `PrintFaktura`, `modSledljivost`) are P2/P3; test-only (`modSEFTests`) is lowest. Phase 1 = triage each caller for propagation dependence; Phase 2 = apply the mechanical fix. Sites (post-RF-28, 44): `modSEFPersistance.bas` ×14 (52/77/170/214/294/354/401/486/610/639/665/691/722/810); `modSEFValidator.bas` ×8 (196/223/245/280/313/340/373/406); `modSEFMapper.bas` ×7 (259/328/790/818/875/902/1034); `modSEFClient.bas` ×4 (307/326/356/862); `modBankaImport.bas` ×2 (134/1260); `modPaletniList.bas` ×2 (873/2808); `modSEFStatusSync.bas` (438); `modSEFTests.bas` (757, test-only); `modConfig.bas` (918); `modDokumenta.bas` (2378); `modFaktura.bas` (508); `modGoogleSheets.bas` (72); `modSledljivost.bas` (263). Line numbers are the `Err.Raise Err.Number` line on `main`@9a21a69; re-scan with the grep in the fix PR before editing. | Root cause `modLogError.bas:50/90` (`On Error` resets `Err`); 44 EH blocks listed at left |

Doctrine note: the fix is source-safe (adds only ASCII locals, no `.frx`, no form `WithEvents`), and
each site keeps its existing log line — only the re-raise argument changes from `Err.Number` to the
captured local. No behavior change on the success path.
