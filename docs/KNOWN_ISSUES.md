# AgriX / OtkupApp Known Issues

**Version:** v6.23 documentation refactor / read-model convergence  
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
