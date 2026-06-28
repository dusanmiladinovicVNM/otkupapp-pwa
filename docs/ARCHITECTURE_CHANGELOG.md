# AgriX / OtkupApp Architecture Changelog

**Document Purpose:** Delta notes between canonical architecture snapshots  
**Companion to:** `ARCHITECTURE_REFERENCE.md`  
**Current Version:** v6.40  
**Last Updated:** 2026-06-22  
**Status:** Active changelog — v6.40 release rutina: `tools/release.sh|ps1` spaja pull→bump `APP_VERSION`→commit→tag→push→stamp u jednu komandu (ostaju 3 Excel klika); GAS `installMonitoringTriggers` sad uključuje i `rebuildMonitoringFleet` (auto rebuild Fleet taba na sat); `docs/RELEASE_PROCEDURE.md` dobio „Jedna komanda" + opcioni „Potpis VBA projekta (opciono)". v6.39 auto-verzija: `stamp-build` i `modBuildInfo` dobijaju `BUILD_VERSION` (`git describe --tags --always`; na tagu `vba-v2.2.1`, posle N commita se sam diže `vba-v2.2.1-3-gabc1234`); `modMonitoring`/`modLicense` šalju `buildVersion`; GAS `Events` + `Fleet` dobijaju kolonu `BuildVersion`. Jedini ručni „bump" je `git tag`; `APP_VERSION` ostaje gruba baza. v6.38 verzionisanje/fleet: novi modul `modBuildInfo` (`BUILD_SHA`/`BUILD_DATE`, stamp pri buildu preko `tools/stamp-build.sh|ps1`); `modMonitoring` i `modLicense` šalju `buildSha`/`buildDate`; GAS `Monitoring.gs` — `Events` dobija kolone `BuildSha`/`BuildDate` (`ensureHeader_` rekonsilijacija, stari redovi ostaju poravnati) + novi `Fleet` tab sa `rebuildMonitoringFleet()`/`getMonitoringFleet` za pregled „ko ima koju verziju"; release procedura R1–R3 u `docs/RELEASE_PROCEDURE.md`. v6.37 panel „Otkupni blokovi" (`modOtkupBlok`): „Cena po otpremnici" je sada **default** za nove blokove (pre-fill `txtCena`); **ručni override u `txtCena` se poštuje i nikad se ne pregazi** (uklonjen `ApplyCenaToOtpremnica`, reverz v6.27 „one price applies to all"). v6.36 prijemnica auto-štampa (`CFG_PRIJEMNICA_PRINT_MODE`) gejtovana na **default hladnjaču** (`frmDokumenta.btnUnosPrij`: štampa samo kad je `kupac == MALINA_DEFAULT_KUPAC`); eksterni kupci se ne štampaju automatski. v6.35 panel „Otkupni blokovi" (`modOtkupBlok`): količina (kg) se u panelu i obe liste (otpremnice + blokovi) prikazuje sa **fiksne 2 decimale** (`FmtKgDec` → `#,##0.00`, ista konvencija kao `frmOtkup.UpdateUkupnoKg`). v6.34 otkupni list: „Saldo ambalaze" je sada kumulativni entitetski saldo kooperanta (Početno stanje pre bloka + Izdato − Primljeno; novi `modAmbalaza.GetKooperantAmbOpening`). v6.33 dnevna/periodična specifikacija (datum od–do) + kolona „Otkupno mesto" u panelu „Otkupni blokovi" (`modOtkupBlok`). v6.32 otkupni list: ambalaža kao uokvirena tabelica (Primljena/Izdata/Saldo). v6.31 daje punu podršku za **Klasu II kroz ceo lanac** (zasebne gajbe #3 — ledger, otpremnica/zbirna/prijemnica, paletizacija), dozvoljava **unos samo Klase II bez Klase I** (#2), uvodi **bruto unos sa obaveznim gajbama** (#1) i **otpremnicu koja čuva neto + `BrutoKg`** (#5, panel poredi neto↔neto); MALINA posivljava Zbirna sekciju u `frmDokumenta`; auto-cena Klase II; PR57 numeracija prijemnice (`GenerateBrojPrijemnice`); fix regresije izdate ambalaže kod unosa samo Klase II

---

## 0. Changelog Contract

This document records differences between canonical architecture snapshots. It is not a replacement for `ARCHITECTURE_REFERENCE.md`.

Allowed content:

- added
- changed
- fixed
- deprecated
- removed
- known issues/regressions
- roadmap changes
- migration/data notes
- verification/acceptance gates

If something is current architecture, it must also be present in `ARCHITECTURE_REFERENCE.md`.

### 0.1 Writing Rule

Each changelog item should state what changed, why, the affected layer, whether the reference must be updated and whether migration is required.

### 0.2 Standard Version Entry Shape

Preferred structure:

1. Summary
2. Added
3. Changed
4. Fixed
5. Removed / Deprecated
6. Known Issues / Known Limitations
7. Roadmap
8. Migration / Data Notes
9. Verification / Acceptance Gates
10. Documentation Actions

Older preserved entries may use equivalent headings such as `VERIFIED / TO VERIFY`, `TESTED`, `DEFERRED` or `Compatibility / Scope Boundaries` when that avoids losing source detail.

---

## 1. Version Index

| Version | Date | Summary | Reference updated | Notes |
|---|---|---|---|---|
| v6.37 | 2026-06-22 | Panel „Otkupni blokovi" (`modOtkupBlok`): „Cena po otpremnici" je sada **default** cena za nove blokove (pre-fill `frmOtkup.txtCena`), a **ručni override u `txtCena` se poštuje i nikad se ne pregazi**. Uklonjen `ApplyCenaToOtpremnica` (i poziv iz `OtkupBlok_AfterUnos` posle snimanja i iz `OnCenaChanged`) — izmena cene po otpremnici više ne preračunava već unete blokove; reverz v6.27 „one price applies to all". `OtkupBlok_AfterUnos` i dalje vraća default u `txtCena` za sledeći blok (brzi unos). | Yes — AR §6 (Otkupni blokovi panel) | behavior-only; bez migracije; re-import `modOtkupBlok` (`frmOtkup`/`clsBlokUI` nepromenjeni) |
| v6.36 | 2026-06-22 | Prijemnica auto-štampa gejtovana na **default hladnjaču**: `frmDokumenta.btnUnosPrij` poziva `OutputPrijemnica` (po `CFG_PRIJEMNICA_PRINT_MODE`) **samo** kad je `kupac == MALINA_DEFAULT_KUPAC`; eksterni kupci se ne štampaju automatski (ručna re-štampa preko `PrintPrijemnica` i dalje radi za sve). | Yes — AR §5.12/§6.4 | presentation-only; re-import `frmDokumenta`; bez migracije |
| v6.35 | 2026-06-22 | Panel „Otkupni blokovi" (`modOtkupBlok`): količina (kg) sada **uvek 2 decimale** u panelu (sažetak Ukupno/U blokovima/Ostatak) i u obe liste — otpremnice (Količina/Ostatak) i blokovi (Količina + zbirni red). Jedna izmena u `FmtKgDec` (`#,##0.###` → `#,##0.00`); ista konvencija kao živi prikaz `frmOtkup.UpdateUkupnoKg`. Cene (`FmtRsd`) i ambalaža/cele gajbe (`FmtKg`) nepromenjeni. | No (uz sledeći snapshot) | presentation-only; bez migracije; re-import `modOtkupBlok` |
| v6.34 | 2026-06-21 | Otkupni list (`modPrint`): red **„Saldo ambalaze" → kumulativni entitetski saldo kooperanta** = početno stanje pre bloka (`modAmbalaza.GetKooperantAmbOpening`, čita ledger po redosledu upisa → ispravno i na re-print) + izdato − primljeno; dodat red **„Pocetno stanje"**; kutija ostaje **3 reda** (primerak ostaje 1/3 A4). Rešava v6.32 known-limit (saldo je bio per-dokument). | Yes — AR §5.12 | presentation/read-only; bez migracije; re-import `modPrint` + `modAmbalaza`; smoke: re-print starijeg bloka — saldo = početno + izdato − primljeno |
| v6.33 | 2026-06-21 | Panel „Otkupni blokovi" (`modOtkupBlok`): dnevna/periodična specifikacija (filter datum od–do, dugme „Stampaj po datumu") + kolona „Otkupno mesto"; akcioni red iznad listboxova; renderer refaktorisan u `RenderSpec`. Fix: Type mismatch po datumu; kg bez praznih decimala (`FmtKgDec`). | No (uz sledeći snapshot) | bez migracije; re-import `modOtkupBlok` |
| v6.32 | 2026-06-21 | Otkupni list (`modPrint`): ambalaža prikazana kao mala **uokvirena tabelica** — redovi Primljena / Izdata / **Saldo ambalaze** (saldo = primljena − izdata; novi `h("ambSaldo")`), levo od obračun-boksa; „Rok isplate" premešten na red oznake primerka (ostaje vidljiv kao mandatorni element). Bez dodavanja redova — primerak i dalje tačno 1/3 A4 (99 mm). | Yes — AR §5.12 | presentation-only; bez migracije; re-import `modPrint`; smoke: štampa otkupnog lista (tabelica primljena/izdata/saldo + staje u 1/3) |
| v6.31 | 2026-06-21 | Dvoklasni otkup — puna podrška za **Klasu II kroz ceo lanac**: Klasa II (drugi `tblOtkup`/dokument red, isti `BrDok`) nosi **svoju** količinu ambalaže `kolAmbII` kroz ledger, otpremnicu/zbirnu/prijemnicu i **paletizaciju** (#3, reverz starog „ambalaža samo na Klasi I"). **Unos samo Klase II** bez Klase I (otkup + dokumenta; `hasKlasaI = kolicinaI > 0`, bar jedna klasa) (#2). **Bruto unos**: ako ima količine bez gajbi → blok snimanja (bez gajbi se bruto ne pretvara u neto, tara se plaća kao voće) (#1). **Otpremnica čuva neto** + `tblOtpremnica.BrutoKg` (panel poredi neto↔neto; labele „bruto (neto X)") (#5). MALINA: posivljena Zbirna sekcija u `frmDokumenta` (`DisableFraZbirnaMalina` — i deca kontrole). Auto-cena Klase II na `chkDveKlase`. Storno celog dokumenta `StornoOtkupByBrDok_TX` (oba reda). PR57: `modBrojevi.GenerateBrojPrijemnice` + auto-predlog. Fix regresije: izdata ambalaža se gubila kod unosa samo Klase II (sad se rutira na red Klase II). | **Yes** — AR §6.1/§6.2/§6.4/§6.8/§6.9 + §5.4/§5.5; companion `AMBALAZA_MODEL.md` §3/§9 | re-import `frmOtkup`/`frmDokumenta`/`modOtkup`/`modDokumenta`/`modAutoHladnjaca`/`modStorno`/`modOtkupBlok`/`modConfig`/`modSetup`/`modPrint`/`modPaletniList`/`modPodesavanja`/`modBrojevi`/`modBusinessFlowProTests`; pokreni `EnsureDoradeSchema` (dodaje `tblOtpremnica.BrutoKg`); stari redovi nemaju `kolAmbII`/`BrutoKg` (prazno = neto, ambalaža samo Klasa I) — bez migracije |
| v6.30 | 2026-06-20 | Otkup desktop: novo polje „Izdata ambalaza" (OM→kooperant uz otkup) — double-entry `Kooperant Ulaz` + `Stanica Izlaz` pod `DOK_TIP_OM_IZLAZ_KOOP` (DokumentID = otkupID, bez vozača; reuse modela iz v6.29 / `AMBALAZA_MODEL.md`), perzistira u `tblOtkup.KolAmbIzdata`; `StornoOtkup` reversuje obe noge (deli otkupID). Otkupni list (`modPrint`) obogaćen: ambalaža blok prikazuje **Primljenu i Izdatu**, tabela stavki proširena na 8 kolona (Cena bez PDV / Cena s PDV / Kol. neto / Kol. bruto = neto + gajbice×tara iz `tblTipAmbalaze.TezinaGajbiceKg` / Vrednost neto), red „Objekat" (lokacija + br. registra; firmin config `SELLER_OBJEKAT_*`) ispod reda sa PIB, i vreme snimanja (`tblOtkup.VremeUnosa` = `Now()`) uz Datum. | Companion `AMBALAZA_MODEL.md` (§3/§9); main ref TBD | re-import `modConfig`/`modSetup`/`modOtkup`/`frmOtkup`/`modStorno`/`modPrint`/`modPodesavanja`; pokreni `EnsureDoradeSchema` (kreira `KolAmbIzdata` + `VremeUnosa`); popuni `SELLER_OBJEKAT_*` u Podešavanjima; postojeći otkupi nemaju izdatu/vreme (prazno); 2-stavke (Klasa I+II) + „Objekat" red su tesni za 1/3 A4 — proveri štampu |
| v6.29 | 2026-06-20 | Ambalaža direction model made consistently **entity-relative** (`Ulaz` = into the entity, `Izlaz` = out); the `Vozac` balance reads as the inverse transport counterparty. Write-side fix: `SavePrijemnica` booked the buyer backwards — now full received = `Kupac` `Ulaz`, returned empties = `Kupac` `Izlaz` (corrects the `Kupac` saldo too). `VozacAmbEffectiveSmer` inverts the `Stanica`/`Kupac` legs so otpremnica loads the driver (`Ulaz`) and prijemnica unloads him (`Izlaz`); a full route nets to 0. `DokumentTip = "Otkup"` has no vozač and is excluded (auto-hladnjača forced the mirror vozač onto every hladnjača otkup, double-charging). | Yes | re-import `modAmbalaza`/`modIzvestaj`/`modDokumenta`; existing prijemnica rows need `Smer` flip or re-seed |
| v6.28 | 2026-06-19 | `OtkupSablon` otkupni-list print switched to a one-third-A4 two-up layout: each primerak (poljoprivrednik + otkupljivac) is exactly 1/3 A4 (99 mm) in the top two thirds, bottom third blank, for the client's pre-perforated paper; prints 1:1 (Zoom 100, margins 0, no fit-to-page) with explicit row heights + a filler row so the copy boundaries land on the 99/198 mm perforations. Content/obračun/klauzula unchanged. | Yes | presentation-only; no migration; re-import `modPrint`; run otkupni-list print gate; calibrate `OL_TOP_MARGIN_TRIM_PT` if the printer ignores a 0 top margin |
| v6.27 | 2026-06-18 | Otkupni blokovi: optional `frmOtkup` panel for per-otpremnica blok entry — clicking an otpremnica pre-fills the existing otkup form (mesto/vrsta/sorta/vozač/broj zbirne/datum/cena), a single per-otpremnica price applies to all its blokovi, and the normal "Unos" auto-links the saved row(s) to the otpremnica via `OtkupBlok_AfterUnos` so remaining-quantity tracking updates without a manual Sledljivost auto-link. | Yes | no business-data migration; UI-only desktop add-on; re-import `modOtkupBlok`+`clsBlokUI`, add 2 guarded hook lines to `frmOtkup` (`AttachOtkupBlokPanel`, `OtkupBlok_AfterUnos`); opt-out via `OTKUP_BLOK_PANEL=NO` |
| v6.26 | 2026-06-18 | Document print/presentation layer: shared `modDocStyle` house style for otkupni/paletni/preradni lists, otkupni-list legal block (PDV-nadoknada klauzula + rok isplate via `OTKUP_KLAUZULA` / `OTKUP_ROK_ISPLATE`), and paletni/preradni redesign (vrsta voca as subtitle, dropped redundant Vrsta column, layout-version auto-rebuild). | Yes | presentation-only; no business-data migration; re-import `modDocStyle`/`modConfig`/`modPrint`/`modPaletniList`; templates regenerate; run otkupni/paletni/preradni print gates |
| v6.25 | 2026-06-16 | Paletni List (pallet) domain added as a canonical local Excel/VBA workflow plus a desktop-only setup mode. | Yes | no business-data migration; `EnsurePaletniListSchema` idempotent; run palletization/prerada/storno gates |
| v6.24 | 2026-05-26 | Systematic documentation supplement: VBA document numbering model, PWA design tokens/fonts/components_v2, Otkup/Otprema/Pregled UI redesign, service-worker/font cache discipline, runtime bugfixes, lazy-loading performance model and Google Sync diagnostic follow-ups. | Yes | no business-data migration; source summary reviewed; target AgriX Git repo still needs direct confirmation because connected GitHub repo resolves to handoverApp |
| v6.23 | 2026-05-18 | PWA otkup read-model convergence: `MgmtReports/OtkupiAll` becomes the master otkup projection for PWA Management/Otkupac display, merged with `OTK-ST-*` operational queue rows and deduped by `ServerRecordID` / `OtkupID` before `ClientRecordID`. | Yes | no historical business-data migration; browser smoke reported as tested; run Management/Otkupac merged-read/dedup gates |
| v6.22 | 2026-05-15 | Residual GO hardening closeout: Faktura duplicate-ID print/status guards, ParcelaID-based geo save/clear APIs, explicit Storno eligibility helper naming and v6.21 Google/MasterSync/Novac/HealthCheck hardening carried forward. | Yes | no historical business-data migration; run Faktura duplicate guard, Geo ByID, Storno eligibility and existing GO closeout gates |
| v6.21 | 2026-05-14/15 | Banka/Storno production hardening plus GO hardening closeout: Google Sheets staging/verify/replace writes, quota/cache hardening, named Kartice tab export, degraded PWA unlock handling, MasterSync exact-row document links, SaveNovac append hard-fail and ProductionHealthCheck duplicate-key preflight. | Yes | no production data migration; code-only VBA/GAS-sync/setup hardening; run parser/import/mapiranje/storno plus Google sync/MasterSync/Novac/HealthCheck gates |
| v6.20 | 2026-05-12 | Permanent full-document correction and BankaImport statement-integrity hardening: v6.17 full reference remains the base, user-added BankaImport saldo/integrity changes are canonical, and full v6.18 + v6.19 material is integrated into AR/CL without shortening. | Yes | no business data migration; add/confirm tblBankaImport saldo columns; run BankaImport statement-integrity smoke, v6.18 master-sync smoke and v6.19 geo/editor/KPI smoke |
| v6.19 | 2026-05-12 | Parcel geo/editor UX and data-integrity hardening: frmStammdaten inline geo UX, selected-parcel Google sync before polygon editor, transaction-backed geo save/clear and frmOtkupAPP KPI robustness. | Yes | no production data migration; VBA compile and geo/editor/KPI smoke required |
| v6.18 | 2026-05-11 | VBA/PWA/GAS master-sync guard: full-cycle sync orchestrator, SyncControl lock, GAS write blocking, PWA master-sync guard, soft-lock retry semantics and parcel geo pull before Stammdaten export. | Yes | no production data migration; GAS/PWA deploy plus master-sync/polygon/cache smoke required |
| v6.17 | 2026-05-09 | Document-chain launch hardening after v6.16: `modAmbalaza` fail-fast ledger validation and strict `Ulaz`/`Izlaz` semantics; `modSledljivost` auto-link monitoring, `COL_OTK_BROJ_ZBIRNE` usage and exact `OtkupID` update guard; `modDokumenta` save/relink/stornirano hardening for dual-class document flows; confirmed `PrijemnicaID` is row-unique while `BrojPrijemnice` groups class rows; BusinessFlowPro suite green 111/111. | Yes | no production data migration; VBA document-chain hardening; GAS/PWA unchanged |
| v6.16 | 2026-05-09 | Agrohemija and Digitalni Agronom launch-hardening: desktop `modAgrohemija`/`frmAgrohemija` validation, stock guard, rollback and monitoring alignment; PWA management agrohemija package-quantity/parcel-separator/submit-lock cleanup; Kooperant `agromere.js` real-quantity dosage and local lager validation. GAS `syncTretman` remains unchanged and accepted as current launch contract. | Yes | no production data migration; VBA + PWA code change required; GAS unchanged |
| v6.15 | 2026-05-07 | BrojZbirne authority shift to PWA-first deterministic generation: PWA generates `x/ddmmyy[-rb]` locally at confirmZbirna, GAS persists provided value (no GAS code change required), VBA ImportRowToTblZbirna reads from VOZ column 20 with GenerateBrojZbirne preserved as fallback for legacy/desktop-manual entry. Reverses v6.12/v6.14.1 VBA-owned generation rule. | Yes | no production data migration; PWA + VBA code change required; GAS unchanged |
| v6.14.1 | 2026-05-06 | Document-flow launch correction: otpremnica `BrojZbirne` optionality confirmed, VOZ writeback B/T split clarified, VOZ empty `ClientRecordID` guard, checked cascade-link and Boolean core import contract | Yes | no production data migration; code-only launch-hardening delta on top of v6.14 |
| v6.14 | 2026-05-06 | Production monitoring and observability layer: VBA best-effort monitoring client, GAS monitoring ingest, OtkupApp_Monitoring_PROD workbook, Health/Alerts/AuditCritical routing, SEF/finance/otkup/document/bank/masterdata event coverage and watchdog health checks | Yes | no production data migration; requires monitoring Script Properties, tblSEFConfig keys, GAS redeploy and watchdog trigger install |
| v6.13 | 2026-05-04 | PWA pre-launch P0 runtime hardening: app-shell/cache stability, unified render dedupe, single sync trigger entrypoint, bootstrap stale-syncing recovery and submit locks for critical save flows | Yes | no production data migration; unit tests deferred; saveParcelPolygon decision and selected final role smoke remain open |
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

## 2. Maintained Version Entries

The following entries are kept in the active changelog because they affect current architecture, launch gates, migration notes or recent production hardening.


## v6.43 — 2026-06-27

### Summary

Ručni šabloni (faktura, kartica kooperanta, sledljivost) prevaziđeni — sada su generisani house-style šabloni u modPrint grupi, istovetnog dizajna kao otkupni/paletni (Faza 3). **Menja štampani izgled ova tri dokumenta** (sadržaj/podaci verno reprodukovani; izgled prelazi na house style).

### Added

- `modPrint`: `EnsureFakturaSablon`/`FillFakturaSablon`, `EnsureKarticaSablon`/`FillKarticaSablon`, `EnsureSledljivostSablon`/`FillSledljivostSablon`. Generišu sheet od nule (`DocSellerHeader` + `DocTitleBlock` + stilizovana tabela), sa layout-version markerom (`H1` faktura/kartica, `N1` sledljivost — 12 kolona, landscape) kao `PaletaSablon`; auto-rebuild na promenu verzije.

### Changed

- `FakturaSablon`/`KarticaSablon`/`SledljivostSablon` više nisu ručni sheet-ovi koji pucaju ako ne postoje — auto-generišu se. Print logika izmeštena iz `frmSledljivost` (forme) u `modPrint`.
- Entry-point funkcije zadržavaju podatke/gardove, render delegiraju na modPrint: `modFaktura.PrintFaktura` (gardovi: single-row, storno, duplicate-FakturaID), `modIzvestaj.PrintKarticaPDF` (`ReportKarticaKooperanta`), `frmSledljivost.PrintTracePDF` (`TraceByZbirna` + prijemnica). Izlaz preko Faza2 dispečera (`*_PRINT_MODE`).
- Nova unikatna named-range imena: faktura `Fak*`, kartica `Kart*`, sledljivost `Sled*` (izbegava koliziju sa starim workbook imenima).

### Known Issues / Known Limitations

- Stara workbook imena ručnog `SledljivostSablon`-a (npr. `KupacNaziv`, `LOTBroj`) mogu ostati kao neaktivna `#REF` posle prvog auto-rebuild-a — bezopasno; očistiti kroz Name Manager po želji.
- `modFaktura.ClearFakturaStavkeArea` je sada mrtav (inertan) kod — kandidat za uklanjanje.
- `PrintKarticaAmbalazePDF` (kartica ambalaže, code-built `_KartAmbPrint`) namerno nedirnut — zaseban dokument.

### Verification / Acceptance Gates

- Statički: balans `Sub`/`Function`, `Select Case`, `With`, `For/Next` u svim diranim modulima; named-range konzistentnost Ensure↔Fill po dokumentu; gardovi fakture očuvani; encoding (Windows-1250) + LF, bez korupcije. **Excel proof obavezan** (izgled se ne može verifikovati bez Excela): odštampati/PDF-ovati po jednu fakturu, karticu i sledljivost.

### Migration / Data Notes

- Nema schema/data migracije. Re-import `modPrint`, `modFaktura`, `modIzvestaj`, `frmSledljivost`, `modDocStyle` → `Compile`. Prvi print svakog dokumenta rebuild-uje šablon (briše stari ručni sheet, generiše house-style); šabloni ne nose poslovne podatke pa je rebuild bezbedan.

Reference updated: Yes — AR §5.12 (generisani Faktura/Kartica/Sledljivost u modPrint, known-inconsistency rešen).


## v6.42 — 2026-06-27

### Summary

Izlazni dispečer obrazaca konsolidovan (Faza 2) i ručni dokumenti dobili konfigurabilan izlaz. Bez promene podrazumevanog ponašanja (defaulti čuvaju zatečeno).

### Added

- `modDocStyle.DocResolveMode(mode, defMode)` — normalizuje `*_PRINT_MODE` na `OFF/PRINT/PREVIEW/PDF`; prazno/nepoznato → `defMode` (po dokumentu). Centralizuje razliku u defaultu (otkupni/grupni/izdamb/kartica/sledljivost = PDF, prijemnica/paletni = OFF, faktura = PRINT) i typo handling.
- `modDocStyle.DocPrintWs(ws, mode)` — štampa/pregled napunjenog sheeta (`PREVIEW` → pregled, inače `PrintOut Copies:=1`).
- `modConfig` `CFG_FAKTURA_PRINT_MODE` / `CFG_KARTICA_PRINT_MODE` / `CFG_SLEDLJIVOST_PRINT_MODE`.

### Changed

- 5 `Output*` dispečera (otkupni/grupni/prijemnica/izdamb/paletni) svedeno na zajednički obrazac: `DocResolveMode` + spojene `PRINT`/`PREVIEW` grane preko `DocPrintWs` (jedan `Fill` umesto dva); `PDF` grana i dalje delegira na `Export*PDF` (čuva putanju folder+timestamp).
- `PrintFaktura` (default `PRINT`), `PrintKarticaKooperanta` (default `PDF`) i `frmSledljivost.PrintTracePDF` (default `PDF`) više nemaju zakovan izlaz — koriste isti obrazac. `PrintKarticaAmbalazePDF` namerno nedirnut (zaseban dokument).

### Verification / Acceptance Gates

- Statički: balans `Select Case`/`End Select` i `Sub`/`End Sub` u svim diranim modulima, jedinstvene definicije helpera, encoding (Windows-1250) + LF očuvani, bez `�` šuma. Excel smoke: otkupni/prijemnica/paletni/revers izlaze isto kao pre; faktura štampa (default), kartica/sledljivost PDF (default); provera da `*_PRINT_MODE = PREVIEW/PDF/OFF` menja izlaz.

### Migration / Data Notes

- Nema schema/data migracije. Re-import `modDocStyle`, `modConfig`, `modPrint`, `modPaletniList`, `modFaktura`, `modIzvestaj`, `frmSledljivost` → `Compile`. Novi config ključevi su opcioni (prazno = zatečeno ponašanje).

Reference updated: Yes — AR §5.12 (`DocResolveMode` / `DocPrintWs`, `*_PRINT_MODE` ključevi za sve obrasce).


## v6.41 — 2026-06-27

### Summary

Print / šablon DRY konsolidacija (Faza 1). Rasuta logika za štampu obrazaca svedena na zajednički `modDocStyle` sloj, bez promene poslovne logike ili izgleda dokumenata (presentation-only).

### Changed

- `modDocStyle.DocExportPdf(ws, pdfPath, openAfter)` — jedan helper umesto 11 skoro identičnih `ExportAsFixedFormat` poziva (modPrint, modPaletniList, modIzvestaj, modOtkupBlok, frmSledljivost). `modPrint:_Print` (`UsedRange` izvoz) namerno ostavljen.
- `modDocStyle.DocPageSetupThirdA4(ws, lastRow)` — četiri bajt-identična 1/3-A4 `PageSetup` bloka (otkupni / grupni / otpremnica / revers ambalaže) na jednom mestu; geometrija 99/198 mm perforacije ne može da drift-uje. Profil B (prijemnica / paletni / preradni) nedirnut (različite margine/kolone).
- `modConfig` `Public Const WS_*_SABLON` — imena šablon sheet-ova kao jedinstven izvor istine (26 hardkodovanih stringova → konstante, uz postojeći `TBL_*`/`COL_*` obrazac).

### Known Issues / Known Limitations

- `FakturaSablon` / `KarticaSablon` / `SledljivostSablon` su i dalje ručni sheet-ovi (greška ako ne postoje, bez `Ensure*` ni `H1` verzije) — kandidat za Fazu 3.
- Izlazni dispečer (`Output*` `Select Case PRINT/PREVIEW/PDF/OFF`) još nije konsolidovan (PDF grana re-fill-uje preko `Export*PDF`); odloženo jer nije čisto behavior-preserving.

### Verification / Acceptance Gates

- Statički: balans `Sub`/`End Sub` u `modDocStyle`, jedinstvene definicije helpera, ispravna arnost poziva, encoding (Windows-1250) i LF očuvani, bez `�` šuma u diffu. Excel smoke (operater): otkupni list, prijemnica, paletni/preradni, faktura, kartica, sledljivost — PDF/print izlaze isto kao pre.

### Migration / Data Notes

- Nema schema/data migracije. Re-import 8 modula (`modDocStyle`, `modConfig`, `modPrint`, `modPaletniList`, `modFaktura`, `modIzvestaj`, `modOtkupBlok`, `frmSledljivost`) → `Debug → Compile`.

Reference updated: Yes — AR §5.12 (modDocStyle `DocExportPdf` / `DocPageSetupThirdA4`, `WS_*_SABLON` konstante, known inconsistency za ručne šablone).


## v6.36 — 2026-06-22

### Summary

Prijemnica je već imala auto-štampu na „Unos" (`OutputPrijemnica` po `CFG_PRIJEMNICA_PRINT_MODE`: PDF/PRINT/PREVIEW/OFF), ali je štampala za **sve** kupce. Sada je gejtovana na **default hladnjaču** (firmin sopstveni magacin), jer se auto-prijemnica/štampa odnosi samo na prijem u tu hladnjaču.

### Changed

- `frmDokumenta.btnUnosPrij_Click`: poziv `OutputPrijemnica result` obmotan proverom `Len(MALINA_DEFAULT_KUPAC) > 0 And kupac == MALINA_DEFAULT_KUPAC`. Za eksterne kupce nema auto-štampe (bez obzira na `CFG_PRIJEMNICA_PRINT_MODE`). Ručna `PrintPrijemnica(prijemnicaID)` ostaje dostupna za sve.

### Verification / Acceptance Gates

- Statički: `frmDokumenta` Sub/Function balansiran; `defHlad` deklarisan jednom; `CFG_MALINA_DEFAULT_KUPAC` postoji.
- Smoke (Excel): prijemnica za default hladnjaču + `PRIJEMNICA_PRINT_MODE=PDF` → otvara PDF; prijemnica za drugog kupca → bez izlaza.

### Migration / Data Notes

- Re-import `frmDokumenta.frm`. Bez migracije/šeme.

Reference updated: Yes — AR §5.12 (template `PrijemnicaSablon` + gejt) i §6.4.


## v6.34 — 2026-06-21

### Summary

Otkupni list (`OtkupSablon` / `modPrint`): red **„Saldo ambalaze"** više nije per-dokument (primljena − izdata ovog otkupa), nego **kumulativni entitetski saldo kooperanta** = **početno stanje pre izrade bloka** + izdato (`Kooperant Ulaz`) − primljeno (`Kooperant Izlaz`). Dodat je i poseban red **„Pocetno stanje"** iznad Primljene/Izdate. Početno stanje se računa iz ledgera **po redosledu upisa** (sve kooperantove `Ulaz/Izlaz` stavke upisane pre prvog reda ovog bloka), pa je ispravno i kod **ponovne štampe starijeg bloka** (kasniji blokovi se ne uračunavaju). Rešava known-limitaciju iz v6.32. Bez promene ledgera — read/presentation-only.

### Added

- `modAmbalaza.GetKooperantAmbOpening(koopID, tipAmb, blockOtkupIDs)` → `Long`: entitetski saldo kooperanta (Ulaz +, Izlaz −) za dati tip ambalaže, sabran nad svim ledger redovima **upisanim pre prvog reda datog bloka**. Blok se identifikuje preko `DokumentID ∈ blockOtkupIDs` (obe noge — `Otkup` i `OM-Izlaz-Koop` — dele `DokumentID = otkupID`). `ExcludeStornirano`; schema guard preko `RequireColumnIndex`. Fallback (blok bez ledger redova, npr. legacy): sve kooperantove stavke kao početno.
- `modPrint.FillOtkupSablon`: `h("ambPocetno")` (`tipAmb & " x " & početno`), `h("ambPrijem")` i `h("ambIzdavanje")` (tip + broj gajbi tekućeg bloka, npr. „12/1 x 50", za kombinovani srednji red).

### Changed

- `modPrint.FillOtkupSablon`: `h("ambSaldo")` = `ambPoc + kolAmbIzd − kolAmb` (entitetski, „koliko gajbica kooperant drži/duguje"); pre toga je bilo `kolAmb − kolAmbIzd` (per-dokument, suprotan znak).
- `modPrint.WriteOtkupCopy`: leva ambalaža kutija ima **3 reda** (`ob..ob+2`): **Pocetno stanje** / **Primljeno + Izdato** (jedan red, dva inline `DocLabelVal` para u kol. 1 i 3) / **Saldo**. Desni obračun-boks (Osnovica/PDV/UKUPNO) ostaje `ob..ob+2`. Visine 13/13/14 pt; `usedPt += 40`; `rr = ob+3` — bez dodavanja reda, geometrija 1/3 A4 (99 mm) nepromenjena u odnosu na v6.32.

### Known Issues / Known Limitations

- Otkupni list prikazuje **jedan tip ambalaže** (iz prvog reda); ako kooperant istorijski drži više tipova, početno/saldo se odnose samo na taj tip.
- Nasleđeno iz v6.30: „Objekat" red + 2 stavke — primerak je već na granici 99 mm; ova izmena NE dodaje red (kutija ostaje 3 reda) pa ne pogoršava tesnoću, ali je vizuelno proveriti na štampi.

### Verification / Acceptance Gates

- VBA verifikovan statički (Function/End Function balans, Select Case/End Select balans, helper definisan jednom i pozvan jednom, `amb*` ključevi samo u `modPrint`). Finalni smoke (korisnik, Excel): re-printaj stariji blok — potvrdi da je Saldo = Početno + Izdato − Primljeno i da primerak staje u 1/3 A4.

### Migration / Data Notes

- Read/presentation-only; bez migracije. Re-import `modPrint` + `modAmbalaza`.

Reference updated: Yes (§5.12).


## v6.33 — 2026-06-21

### Summary

Panel „Otkupni blokovi" (`frmOtkup` / `modOtkupBlok`): pored postojeće specifikacije za **ručno izabrane otpremnice**, dodata je **dnevna / periodična specifikacija** (filter po datumu *od–do*). Iznad listboxova je novi **akcioni red**: levo (nad otpremnicama) polja **Od** / **Do** (pretpopunjena na današnji datum → klik odmah daje dnevnu specifikaciju) + dugme **„Stampaj po datumu"**; desno (nad blokovima) preseljena dugmad **Storniraj / Stampaj list / Biraj otpremnice**. Listboxovi su spušteni (`GRID_TOP` 104 → 120, zaglavlja 74 → 90). Na samoj specifikaciji je dodata kolona **Otkupno mesto** (posle „Broj otpremnice"). Renderer je refaktorisan u jedno jezgro (`RenderSpec`) koje filtrira po skupu `OtpremnicaID` **ILI** po opsegu datuma — bez dupliranja PDF logike.

### Added

- `modOtkupBlok.PrintSpecifikacijaPoDatumu(datumOd, datumDo)` + handler `PrintSpecOdDo` (čita Od/Do preko `TryParseDateValue`, validira opseg): dnevna (od=do) ili periodična specifikacija svih ne-storniranih otkup blokova čija je kolona `Datum` (`COL_OTK_DATUM`) u opsegu.
- `BuildPanel`: dinamičke kontrole `txtOtkBlokSpecOd` / `txtOtkBlokSpecDo` + dugme `btnOtkBlokSpecDatum` (akcija `"SPECDATUM"` u `OtkupBlok_OnButton`), wired preko `clsBlokUI`; smeštene u akcioni red iznad listboxova. `frmOtkup.frx` se ne dira.
- Specifikacija: nova kolona **Otkupno mesto**, vrednost = naziv stanice po redu (`COL_OTK_STANICA` → `TBL_STANICE.Naziv`, reuse `BuildLookup`).

### Changed

- `modOtkupBlok.PrintSpecifikacija(otpIDs)` je sada tanak omotač oko novog `RenderSpec(selSet, byDate, datumOd, datumDo, subtitle)`; ponašanje ručne selekcije otpremnica je nepromenjeno. Tabela sada ima 11 kolona (A–K); izlaz je sortiran po (Otkupno mesto, Datum) radi grupisanja.
- Layout panela: dugmad `Storniraj` / `Stampaj list` / `Biraj otpremnice` preseljena iz reda naslova u novi akcioni red (red 66); listboxovi spušteni (`GRID_TOP` 104 → 120).

### Fixed

- **Type mismatch (greška 13)** pri kliku na „Stampaj po datumu": `DateValue(datum)` u `PrintSpecifikacijaPoDatumu` i u filteru `RenderSpec` zamenjeno robusnim poređenjem serijskog broja `Int(CDbl(datum))` (bez parsiranja stringa → bez locale/Type zamke).
- **Količina (kg) se više ne zaokružuje** u panelu: nova `FmtKgDec` (`#,##0.###` — decimale samo kad su unete) za sve prikaze količine (lista otpremnica Količina/Ostatak, lista blokova Količina + zbir, sažetak kg, poruka o prekoračenju) i za kolonu Količina na PDF specifikaciji. Cene i ambalaža (cele gajbe) ostaju nepromenjene (`FmtKg`).

### Verification / Acceptance Gates

- VBA verifikovan statički (Sub/Function/With/Select balans, nema duplih definicija, indeksi kolona 1–11 dosledni). Finalni smoke (korisnik, Excel): u panelu „Otkupni blokovi" → „Stampaj po datumu" za danas i za period; potvrdi kolonu „Otkupno mesto" i da postojeća specifikacija po izboru otpremnica i dalje radi.

### Migration / Data Notes

- Nema migracije (koristi postojeće `tblOtkup.Datum` / `StanicaID`). Re-import `modOtkupBlok` (`clsBlokUI` nepromenjen).

Reference updated: No (UI/izveštaj u panelu; uneti u `ARCHITECTURE_REFERENCE.md` uz sledeći snapshot).


## v6.32 — 2026-06-21

### Summary

Otkupni list (`OtkupSablon` / `modPrint`): ambalaža je sada mala **uokvirena tabelica** sa tri reda — **Primljena**, **Izdata** i **Saldo ambalaze** — levo od uokvirenog obračun-boksa (ogledalo njegovog izgleda). Saldo = primljena − izdata za taj dokument. Nastavak na v6.30/v6.31 koji su uveli prikaz primljene/izdate ambalaže. Bez promene poslovne logike i **bez dodavanja redova** — geometrija 1/3 A4 (99 mm po primerku) je očuvana.

### Changed

- `modPrint.FillOtkupSablon`: dodat `h("ambSaldo") = tipAmb & " x " & CLng(kolAmb − kolAmbIzd)` (primljena = zbir gajbi po stavkama, izdata = `tblOtkup.KolAmbIzdata`).
- `modPrint.WriteOtkupCopy`: levi ambalaža blok (`ob..ob+2`) je sada Primljena / Izdata / **Saldo** sa `BorderAround` (kolone 1–4) i gornjom linijom na saldo redu; saldo red podebljan. „Rok isplate" premešten sa `ob+2` na red oznake primerka (kolona 5, font 8) da se napravi mesto za saldo bez dodavanja reda — mandatorni element ostaje vidljiv. Visine redova i `usedPt` nepromenjeni.

### Known Issues / Known Limitations

- Saldo je **per-dokument** (primljena − izdata ovog otkupa), ne kumulativni saldo kooperanta; za kumulativni bi se koristio `modAmbalaza.GetAmbalazeStanje(koopID, "Kooperant")`.
- Nasleđeno iz v6.30: sa „Objekat" redom + 2 stavke (Klasa I+II) primerak je već na granici 99 mm; ova izmena ne dodaje visinu ali ne otklanja tu tesnoću.

### Verification / Acceptance Gates

- VBA verifikovan statički (Function/Sub/With balans, `WriteOtkupCopy` arity = 7, `ambSaldo` definisan i korišćen). Finalni smoke (korisnik, Excel): odštampaj/pregledaj otkupni list, potvrdi tabelicu ambalaže (primljena/izdata/saldo) i da primerak staje u svoju trećinu.

### Migration / Data Notes

- Presentation-only; bez migracije. Re-import `modPrint`.

Reference updated: Yes (§5.12).


## v6.31 — 2026-06-21

### Summary

Dvoklasni otkup (Klasa I + Klasa II) dobija **punu, simetričnu podršku** za Klasu II kroz ceo lanac. Do sada je Klasa II bila „drugorazredna": delila je ambalažu sa Klasom I i nije ulazila u paletizaciju/dokumente sa svojim gajbama. Sada su Klasa I i Klasa II **dva ravnopravna `tblOtkup`/dokument reda** koji dele isti `BrDok`, a svaki nosi svoju količinu i svoju ambalažu. Uz to: dozvoljen je **unos samo Klase II** (npr. kada nema prve klase), **bruto unos** sa obaveznim gajbama (da se tara ne plati kao voće), i **otpremnica čuva neto** (+`BrutoKg`) da panel poredi neto↔neto. Spojen je PR57 (numeracija prijemnice) i posivljena je Zbirna sekcija u MALINA modu.

### Added

- **#3 — Klasa II nosi zasebnu ambalažu kroz ceo lanac.** Klasa II red sada prosleđuje svoju količinu gajbica (`kolAmbII`) u: ledger (`TrackAmbalaza` dvojni upis, kao Klasa I), otpremnicu/zbirnu/prijemnicu (`modDokumenta.Save*Multi_TX`), auto-hladnjača lanac (`modAutoHladnjaca.AutoChainHladnjaca`) i **paletizaciju** (`modPaletniList`). Reverz starog modela gde je ambalaža bila vezana samo za Klasu I. UI: zasebno polje za gajbe Klase II na **pola širine** (deli red sa „Kolicina ambalaze", ogledalo `txtKolicinaKLII`) u `frmOtkup` (`m_txtKolAmbalazeII`/`ShowKolAmbalazeII`) i `frmDokumenta` (3 runtime polja `m_txtKolAmbIIOtp/Zbr/Prij` preko `ambI.Parent.Controls.Add` — `.frx` netaknut).
- **#2 — Unos samo Klase II (bez Klase I).** I otkup (`frmOtkup`/`SaveOtkupMulti_TX`) i dokumenta (`SaveOtpremnicaMulti_TX`/`SaveZbirnaMulti_TX`/`SavePrijemnicaMulti_TX`) prihvataju unos samo Klase II; `hasKlasaI = (kolicinaI > 0)`, obavezna je bar jedna klasa (greške 1812/1103/1203/1303). Kad se unosi samo Klasa II, Količina I i Ambalaža I moraju biti prazni. Kes/avans i izdata ambalaža se tada vezuju na primarni red koji postoji (Klasa II).
- **#5 — Otpremnica čuva neto + `BrutoKg`.** Kad je bruto unos, otpremnica oduzme ambalažu i čuva **neto** (nova kolona `tblOtpremnica.BrutoKg` = `COL_OTP_BRUTO`, uz postojeće `tblOtkup.BrutoKg`/`tblPrijemnica.BrutoKg`), pa panel u bloku poredi **neto↔neto**. Panel labele u bruto modu prikazuju „bruto (neto X)" za sve tri vrednosti.
- **Storno celog dokumenta.** `modStorno.StornoOtkupByBrDok_TX(brDok)` stornira **sve** redove jednog otkupnog dokumenta (Klasa I + Klasa II) atomično (jedna transakcija). Koristi se iz `frmDokumenta` storno putanje i iz panela (`modOtkupBlok.StornoSelectedBlok`, sa fallback-om na `OtkupID`).
- **Auto-cena Klase II** u `frmDokumenta` na `chkDveKlase` (otpremnica/prijemnica) — `AutoFillCenaDok` predloži cenu Klase II iz cenovnika.
- **PR57 — numeracija prijemnice.** `modBrojevi.GenerateBrojPrijemnice(kupacID, datum)` (`MaxSeqFromTable` + `FormatBroj`) + auto-predlog broja u `frmDokumenta` (`RefreshBrojPrijSuggestion`). Dvoklasna prijemnica nosi **isti** broj (jedna prijemnica = jedan broj).
- Test `modBusinessFlowProTests.Test_OtkupClassIIAmbalaza` (registrovan) — pokriva da Klasa II knjiži svoju ambalažu.

### Changed

- **#1 — Bruto mod: gajbe su obavezne.** U bruto modu (i Klasa I i Klasa II), ako je uneta količina a broj gajbi je prazan → snimanje se **blokira**. Bez gajbi se bruto ne pretvara u neto, pa bi se tara (težina gajbica) platila kao voće.
- **MALINA — Zbirna sekcija posivljena** u `frmDokumenta` (`DisableFraZbirnaMalina`). Zbirne se u malina modu prave automatski iz otpremnica, pa je ručni unos onemogućen; pošto `Frame.Enabled = False` ne posivi decu kontrole u MSForms, funkcija iterira i **decu** frejma.
- `modPrint` (otkupni list): „Primljena ambalaža" = zbir `kolAmb` preko **svih** stavki dokumenta (Klasa I + II), ne samo Klase I.
- `modAutoHladnjaca` / `modDokumenta`: Klasa II prolazi kroz auto-lanac i paletizaciju sa svojim `kolAmbII`/`brutoKgII`; Klasa I je opciona (`If hasKlasaI Then`).

### Fixed

- **Regresija — izdata ambalaža se gubila kod unosa samo Klase II.** Od `d7ea1ee` (#2) je `SaveOtkup` za Klasu I obmotan u `If hasKlasaI Then`, a izdata ambalaža (`OM→kooperant`, `DOK_TIP_OM_IZLAZ_KOOP`) se upisuje **samo** unutar `SaveOtkup` Klase I. Kod unosa samo Klase II (`hasKlasaI = False`) taj poziv se preskakao → `kolAmbIzdata` se nije upisivao ni u `tblOtkup.KolAmbIzdata` ni u ledger. Sada se `kolAmbIzdata` rutira na red Klase II kad nema Klase I (po uzoru na postojeći `novacII` obrazac; `kolAmbIzdataII = 0` kod dvoklasnog → bez dvostrukog upisa). **Normalan/dvoklasni otkup nije bio pogođen** (izdata ostaje na redu Klase I).
- `frmDokumenta` otpremnica „samo Klasa II": uklonjen rani `Validacija` blok (`IsNumeric`/`val`) koji je blokirao unos pre izmenjene provere.
- Storno dvoklasnog dokumenta iz `frmDokumenta` više ne stornira samo jedan red (Klasu II) — `StornoOtkupByBrDok_TX` hvata oba reda (dele `BrDok`).

### Known Issues / Known Limitations

- Bruto-mod validacija „gajbe obavezne" je u UI sloju (`frmOtkup`/`frmDokumenta`); direktan poziv `Save*Multi_TX` sa bruto bez gajbi nije zaustavljen na nivou TX-a (UI je jedini ulaz).
- Postojeći (stari) otkupi/dokumenti nemaju `kolAmbII`/`BrutoKg` — prikazuju se kao Klasa-I-only / neto. Bez migracije (vidi dole).

### Verification / Acceptance Gates

- VBA se ne kompajlira u ovom okruženju — verifikovano **statički**: balans `Sub`/`Function` i `If`/`End If` po izmenjenom modulu, nema duplih `Public` definicija (`modOtkup` 9/9 Sub-Fn, 47/47 If). `modBusinessFlowProTests` zelen sa dodatim `Test_OtkupClassIIAmbalaza`.
- Finalni smoke (korisnik, Excel): (1) **samo Klasa II + izdata** → upis u `tblOtkup.KolAmbIzdata` + dvojni ledger; (2) **dvoklasni otkup** → obe klase paletizovane, obe gajbe kroz otpremnicu/zbirnu/prijemnicu; (3) **bruto bez gajbi** → blokira snimanje; (4) **storno dvoklasnog dokumenta** → oba reda stornirana; (5) **MALINA** → Zbirna sekcija posivljena; (6) **otpremnica bruto** → čuva neto, panel poredi neto↔neto.

### Migration / Data Notes

- Re-import (forme idu sa `.frx` parom): `frmOtkup.frm`, `frmDokumenta.frm`, `modOtkup.bas`, `modDokumenta.bas`, `modAutoHladnjaca.bas`, `modStorno.bas`, `modOtkupBlok.bas`, `modConfig.bas`, `modSetup.bas`, `modPrint.bas`, `modPaletniList.bas`, `modPodesavanja.bas`, `modBrojevi.bas`, `modBusinessFlowProTests.bas`.
- Pokreni **`EnsureDoradeSchema`** — dodaje `tblOtpremnica.BrutoKg` (uz ranije `tblOtkup.KolAmbIzdata`/`VremeUnosa`/`BrutoKg` i `tblPrijemnica.BrutoKg`), format `0.00`.
- `Debug → Compile VBAProject` (mora bez greške — proveri duple `Public` posle merge-a).
- Bez migracije podataka: stari redovi ostaju Klasa-I-only / neto (prazno `kolAmbII`/`BrutoKg`).

Reference updated: **Yes.** `ARCHITECTURE_REFERENCE.md` §6.1 (dvoklasni otkup — zasebni redovi sa istim `BrDok`, opciona Klasa I, bruto sa obaveznim gajbama, izdata + `VremeUnosa`), §6.2 i §6.4 (Klasa II nosi svoju ambalažu, bruto čuva neto + `BrutoKg`, `GenerateBrojPrijemnice`), §6.8 (ledger knjiži ambalažu po klasi), §6.9 (`StornoOtkupByBrDok_TX` — storno celog dokumenta), §5.4 i §5.5 (`KolAmbIzdata`/`VremeUnosa`/`BrutoKg` kolone). Companion `AMBALAZA_MODEL.md` §3/§9.


## v6.29 — 2026-06-20

### Summary

Corrected the ambalaža (packaging) **direction model** so the ledger `Smer` is consistently **entity-relative** (`Ulaz` = crates into that entity, `Izlaz` = out) and the driver (`Vozac`) balance reads as the inverse transport counterparty. Three parts: **(1) Write-side data fix** — `SavePrijemnica` booked the buyer side backwards (full crates received = `Izlaz`, returned empties = `Ulaz`); now full = `Kupac` `Ulaz` and returned = `Kupac` `Izlaz`, matching every other document, which also corrects the `Kupac` saldo/report. **(2) Vozač = inverse counterparty** (single-entry — the driver leg is derived on read): `VozacAmbEffectiveSmer` inverts the `Stanica` and `Kupac` legs so an otpremnica **loads** the driver (`Ulaz`) and a prijemnica **unloads** him (`Izlaz`); a complete otpremnica→prijemnica route nets to 0, an open otpremnica shows a positive saldo = crates still on the driver. **(3) Otkup has no vozač** — `DokumentTip = "Otkup"` (`Kooperant` procurement) is excluded from the driver balance; it was double-charging the same crates already on the otpremnica, made universal by auto-hladnjača forcing the mirror vozač onto every hladnjača otkup. Surfaced via the otkup-in-own-hladnjača case (buyer = the firm, "vozač" = the station).

### Added

- **OM issues empty packaging to a kooperant** (`frmDokumenta`, OM-Ulaz frame). A runtime toggle `tglIzdKoop` (added as a child of `fraOMUlaz` — `.frx` untouched) switches the frame's ambalaža direction: default **Prijem na OM** (`Stanica` `Ulaz`, unchanged) vs **Izdavanje kooperantu**, which books a **double leg** (`DOK_TIP_OM_IZLAZ_KOOP`, no vozač): `Kooperant` `Ulaz` (kooperant receives empties) **+ `Stanica` `Izlaz`** (OM discharged for the same amount) — both rows share `brojDok`/`DokumentTip` in one transaction. Double-entry is required here because two real entities move (OM and kooperant) and neither is derivable from the other's row (unlike the vozač). New constant `DOK_TIP_OM_IZLAZ_KOOP`; `SaveOMUlaz_TX` gained an `izdavanjeKoop` flag. The kooperant is resolved with `GetComboID` (real `KooperantID` from the combo's hidden column, not the display name) — also fixes the same latent bug on the OM-Ulaz money + open-otkupi paths. Storno of this type follows the existing OM-Ulaz gap (not in the storno combo).
- **Otkup now charges the OM (double-entry).** Otkup previously booked only `Kooperant` `Izlaz` (kooperant returns full) and never credited the OM. `modOtkup.SaveOtkup` and the PWA `modMasterSync` import now also book `Stanica` `Ulaz` (OM charged; same `brojDok`/`DOK_TIP_OTKUP`, no vozač), symmetric to OM-izdavanje — closing the OM↔kooperant loop (izdavanje OM−/koop+, otkup koop−/OM+ net out). Storno reverses both (by document). The vozač report is unaffected (otkup is excluded). Existing otkup rows keep the old single leg → re-seed or migrate to add the missing `Stanica` `Ulaz`.

### Fixed

- **Write-side data fix (`modDokumenta.SavePrijemnica`):** the prijemnica booked the buyer backwards. Now entity-relative — `kolAmb` (full crates the hladnjača receives from the zbirna) = `Kupac` `Ulaz`, `kolAmbVracena` (empties the hladnjača returns) = `Kupac` `Izlaz` (previously `Izlaz`/`Ulaz`). Also corrects the `Kupac` saldo (`GetAmbalazeStanje`) and the `Kupac` packaging report.
- **`modAmbalaza.VozacAmbEffectiveSmer(smer, entitetTip)` — driver = inverse counterparty** (single-entry; derived on read, not stored): `Stanica` and `Kupac` transport legs are sign-inverted (otpremnica `Izlaz` → driver `Ulaz`/load; prijemnica `Ulaz` → driver `Izlaz`/unload), `Kooperant` is left raw (and excluded anyway). A complete route nets to 0; an open otpremnica shows a positive saldo = crates still on the driver. Invalid `Smer` still fails fast.
- `modAmbalaza.GetVozacAmbSaldo`: reads `EntitetTip` + `DokumentTip`, skips `Otkup`, and routes each row through `VozacAmbEffectiveSmer` before bucketing `Izlaz`/`Ulaz`.
- `modIzvestaj.ReportAmbalaza` (`Vozac` pojedinačni + zbirni): passes an `isVozac` flag into `ReportAmbalazePojedinacni`/`ReportAmbalazeZbirni`, which apply the same inversion. Entity reports (`OM`/`Kupac`) are unchanged — `isVozac = False` keeps the raw `Smer`.
- **Otkup excluded from vozač custody (read-side).** `ReportAmbalaza` (vozač branch) adds a `DokumentTip <> "Otkup"` filter (`clsFilterParam` `<>`), and `GetVozacAmbSaldo` skips `Otkup` rows (`AmbText(colDokTip) = DOK_TIP_OTKUP → NextRow`). Fixes both existing and new data; `tblOtkup.VozacID` (zbirna grouping / traceability) and the auto-hladnjača writeback are untouched — only the driver **saldo** ignores otkup.

### Design Decisions

- **Uniform document flow for kooperant → own hladnjača (intentional, not a gap).** When a kooperant delivers directly into the firm's own cold storage, the otkup is still recorded through the normal otpremnica + prijemnica legs (buyer = the firm, "vozač" = the station) rather than through a special internal-transfer path. This deliberately trades a few redundant documents/rows for a single code path, one document layout to maintain, and no special case to remember — nothing is mis-entered, and with the vozač leg now netting to 0 those redundant rows no longer distort any saldo. A dedicated internal-transfer flow is therefore **not** planned; do not "fix" this by special-casing it.

### Known Issues / Known Limitations

- The driver-inverse rule keys on `EntitetTip`: `Stanica` and `Kupac` (transport legs) are inverted, `Kooperant` (otkup) has no vozač and is excluded. Matches every current booking. A future entity that is not a simple inverse transport counterparty would need the rule revisited.
- Write-side left as-is: `modOtkup:448` still tags the otkup ambalaza ledger row with the otkup's vozač (the PWA sync path `modMasterSync:1627` already passes an empty vozač). It no longer affects the vozač saldo (excluded on read), but the ledger field stays populated; a future optional cleanup could pass an empty vozač at `modOtkup` for consistency with the sync path.
- The exclusion assumes every buyer delivery has an otpremnica leg (auto-hladnjača always creates one), so dropping otkup does not lose the driver's "load"; a direct kooperant→buyer delivery with no otpremnica would need the rule revisited.

### Verification / Acceptance Gates

- VBA is not compiled in this environment; verified statically (Sub/Function balance, single `Public VozacAmbEffectiveSmer`, call arity: `ReportAmbalazeZbirni` = 6, `ReportAmbalazePojedinacni` = 9, helper = 2). Existing smoke tests keep passing — `modIzvestajTests` (ReportAmbalaza Vozac) and `modBusinessFlowProTests` (`GetVozacAmbSaldo` → `Not IsEmpty`) assert presence, not the saldo value.
- Final smoke test (user, in Excel): Izveštaj → Vozači → a vozač with one otpremnica + matching prijemnica for the same crates should show **Saldo 0** (an open/partial route still shows the real manjak). Auto-hladnjača case: a hladnjača otkup chain (otkup + auto otpremnica + prijemnica, all on the mirror vozač = `StanicaID`) should also show **Saldo 0** — the otkup legs no longer appear in / charge the vozač balance.

### Migration / Data Notes

- **Write-side change** (`modDokumenta.SavePrijemnica`) — re-import `modAmbalaza.bas`, `modIzvestaj.bas`, `modDokumenta.bas`. **Existing `tblAmbalaza` prijemnica rows keep the old (reversed) `Smer`**: either re-seed test data, or run a one-time migration flipping `Smer` (`Izlaz`↔`Ulaz`) on rows with `DokumentTip = "Prijemnica"`. Until then, historical prijemnica / `Kupac` / vozač numbers mix conventions. Other document types and schema are unchanged.

Reference updated: Yes (Ambalaza ledger read/saldo rules).


## v6.28 — 2026-06-19

### Summary

`OtkupSablon` otkupni-list print geometry switched to a one-third-A4 two-up layout. Each primerak (poljoprivrednik + otkupljivac) now occupies exactly 1/3 of A4 (99 mm) in the top two thirds; the bottom third stays blank, matching the client's pre-perforated paper (two perforations → three equal parts, two copies printed, third left empty). Document content, BRUTO→neto obračun and the legal PDV-nadoknada klauzula are unchanged — only layout geometry and print scaling changed.

### Changed

- `modPrint.FillOtkupSablon`: page setup now prints 1:1 (`Zoom = 100`, `TopMargin/BottomMargin/HeaderMargin/FooterMargin = 0`, `LeftMargin/RightMargin = 0.31"`, `CenterHorizontally`) instead of `FitToPagesWide/Tall = 1`, so explicit row heights map directly to millimetres and the two copies span exactly 198 mm. The printed scissor/cut line between the two copies was removed (the paper is pre-perforated; the copy boundary falls on the perforation).
- `modPrint.WriteOtkupCopy`: gained a `targetPt` parameter; sets explicit row heights (no AutoFit), compacts the copy to ~90 mm of content, and appends a filler row that pads each copy to exactly 99 mm. Seller-name/title/klauzula fonts and the stavke header labels were tightened to fit one third.
- New `modPrint` constants: `OL_THIRD_PT` (= 280.63 pt = 99 mm), `OL_TOP_SPACER_PT`, `OL_MIN_FILLER_PT`, `OL_TOP_MARGIN_TRIM_PT`.

### Known Issues / Known Limitations

- Layout is sized for 1–2 stavke per primerak; 3+ items reduce the bottom safety gap and can push a copy past 99 mm.
- Physical perforation alignment requires printing at 100% / Actual Size with the printer honouring a ~0 top margin. Printers that force a top margin shift content down — compensate by setting `OL_TOP_MARGIN_TRIM_PT` to that margin in points (T mm / 25.4 × 72).

### Verification / Acceptance Gates

- VBA is not compiled in this environment; verified statically (Sub/Function balance, no duplicate Public definitions, `WriteOtkupCopy` call arity = 7). Final smoke test (user, in Excel): run an otkupni-list print/preview and confirm each copy fits inside its third and the copy boundaries align to the 99 mm / 198 mm perforations.

### Migration / Data Notes

- Presentation-only; no business-data or schema migration. Re-import `modPrint`. `OtkupSablon` is cleared and fully redrawn each print, so it auto-updates.

Reference updated: Yes (§5.12).


## v6.27 — 2026-06-18

### Summary

Optional **Otkupni blokovi** entry panel for the desktop `frmOtkup`. It is not a new save path: it is a per-otpremnica entry aid that drives the existing otkup form and links each saved otkup row ("blok") to its otpremnica. UI-only desktop add-on; no new tables, no schema or business-data migration.

### Added

- `modOtkupBlok` (feature module, dynamic UI) + `clsBlokUI` (WithEvents wrapper for the dynamically created controls). Attached to `frmOtkup` via `AttachOtkupBlokPanel Me` in `UserForm_Initialize`; toggled on a button and hidden by default. Controls are built with `Controls.Add`, so `frmOtkup.frx` is unchanged.
- Otpremnice preview (middle) + otkupni-blokovi list (right, for the selected otpremnica), a per-otpremnica price field, and a live "Ukupno / U blokovima / Preostalo" summary computed as otpremnica `Kolicina` − Σ linked `tblOtkup.Kolicina`.
- Clicking an otpremnica pre-fills the existing `frmOtkup` controls — `cmbOtkupnoMesto`, `cmbVrstaVoca`, `cmbSortaVoca`, `cmbVozac` (via `SetComboByID` with a display-`(ID)` fallback for single-column combos), `txtBrojZbirne`, `txtDatum`, `txtCena` — so each blok needs only kooperant + količina. Hladnjača shown in the preview is resolved from `tblZbirna` via `BrojZbirne`.
- Otkup-list number (`txtBrojDokumenta`) is set from the **canonical `SuggestNextBroj`** (the same generator `frmOtkup` uses) for the selected OM + the otpremnica date, yielding `OM/ddmmyy[-N]` instead of today's date. The panel calls it **explicitly** in `PrefillLeftForm` rather than relying on `cmbOtkupnoMesto_Change`, because that event does not fire when consecutive otpremnice share the same OM (the number would otherwise keep the previous otpremnica's date). Normal (panel-closed) entry is unchanged — its number is already driven by the selected OM + `txtDatum`.
- `OtkupBlok_AfterUnos(result)` hook called from `frmOtkup.btnUnos_Click` after a successful `SaveOtkupMulti_TX`: links the returned `OtkupID`(s) (split on `" + "`) to the selected otpremnica via `COL_OTK_OTPREMNICA_ID` (transaction-backed `RequireUpdateCell`) and refreshes the panel — remaining-quantity tracking updates without a manual `modSledljivost` auto-link.
- Per-otpremnica price: one price applies to all of an otpremnica's blokovi (`ApplyCenaToOtpremnica` propagates `tblOtkup.Cena` across all linked rows); price is stored/displayed as BRUTO (PDV-nadoknada included), consistent with the otkupni-list print, and the blok list shows neto cena / vrednost / iznos PDV / ukupno.
- Opt-out config key `OTKUP_BLOK_PANEL` (`NO` hides the toggle button entirely).
- Panel actions/columns: the otpremnice list gets a **Preostalo** column (otpremnica `Kolicina` − Σ linked blokovi), a **Prikaz: Sve/Nezavrsene** filter (hide fully-covered otpremnice) and datum-descending sort; the blok list gets a **UKUPNO** totals row and per-row **Storniraj blok** (`StornoOtkup_TX`) and **Stampaj list** (`PrintOtkupniList`) buttons. `OtkupBlok_ConfirmUnos` (called from `btnUnos_Click` before the save) warns when the entered quantity exceeds Preostalo; after a blok that fills the otpremnica the panel auto-deselects it (Preostalo = 0) to prevent accidental mislinking of the next Unos.

### Changed

- `frmOtkup`: three guarded hook lines only — `AttachOtkupBlokPanel Me` (`UserForm_Initialize`), `If Not OtkupBlok_ConfirmUnos() Then Exit Sub` (in `btnUnos_Click`, before the save) and `OtkupBlok_AfterUnos result` (after `ClearOtkupFields`). No `.frx`/layout change; the existing single-row otkup entry is unchanged when the panel is closed.

### Notes / Migration

- No business-data or schema migration. Re-import `modOtkupBlok.bas` and `clsBlokUI.cls`; add the two hook lines to `frmOtkup` (or re-import `frmOtkup.frm` together with its unchanged `frmOtkup.frx`).
- Reuses canonical invariants: exact-row `OtpremnicaID` linking (no fuzzy key), checked writes via `RequireUpdateCell`, transaction rollback on failure.
- While the panel is open with an otpremnica selected, every "Unos" links to that otpremnica ("blok mode"); hide the panel for an unrelated direct otkup.

---

## v6.26 — 2026-06-18

### Summary

Document print/presentation layer for the otkup, paletni and preradni lists: a shared `modDocStyle` house style, the legally required otkup-block elements (PDV-nadoknada klauzula + rok isplate) on the otkupni list, and a paletni/preradni list redesign. Presentation-only; no business-data or document-chain change.

### Added

- `modDocStyle` shared print-styling module (all Public): `DocSellerHeader` (company name/address/PIB-MB-ziro from `SELLER_*` + optional logo + rule line), `DocTitleBlock` (descriptor + large title with rules), `DocLabelVal` (label + rich-text-bold value in one cell), `DocLogoPath` / `DocDrawLogo` (logo from `SELLER_LOGO_PATH` or `<workbook>\logo.png`/`.jpg`, silently skipped if absent), `DocConfigOr` and color helpers.
- Otkupni-list legal block: PDV-nadoknada **klauzula** (`OtkupKlauzulaDefault`, čl. 34 ZPDV) and **rok isplate**, both configurable; new constants `CFG_OTKUP_KLAUZULA` (`OTKUP_KLAUZULA`), `CFG_OTKUP_ROK` (`OTKUP_ROK_ISPLATE`), `OTKUP_ROK_DEFAULT` (`Po dogovoru`).
- `OtkupSablon`: styled two-up A4 otkupni list (two primerka — poljoprivrednik + otkupljivac — separated by a cut line), neto cena/vrednost with PDV nadoknada as a separate obračun line, framed obračun box, signatures.
- Layout-version marker (cell `H1`) in `PaletaSablon` / `PreradaSablon`: `Ensure*Sablon` rebuilds the sheet once when the marker is stale, so existing templates auto-upgrade; named ranges preserved.

### Changed

- `PaletaSablon` / `PreradaSablon` restyled to the shared house style (logo, two-line title, framed right-side summary with highlighted BRUTO/Neto, gray footer); sheet font Calibri.
- Paletni and preradni stavke tables drop the redundant `Vrsta` column (a pallet/prerada is one fruit type) → `Rb | Kooperant | Neto kg | Ambalaza` (Ambalaza merged across two columns). `Vrsta voca` becomes a large subtitle above the table via named ranges `PalVrsta` (pallet head) and `PreVrsta` (first traced otkup row).
- `modPrint` otkupni rendering uses the shared `modDocStyle` helpers; rows auto-fit, with fixed heights for the merged title and klauzula rows.

### Fixed

- Otkupni-list klauzula clipping: a trailing `EntireRow.AutoFit` collapsed the merged klauzula row to a single line; AutoFit moved to before content render so the fixed klauzula height survives.

### Migration / data notes

- `OtkupSablon` / `PaletaSablon` / `PreradaSablon` are generated render targets (no business data) and regenerate from tables. `PaletaSablon` / `PreradaSablon` auto-rebuild once on layout-version change (`H1`), replacing any manual sheet styling.
- Config is optional and uses strict exact-key match via `GetConfigValue`: `OTKUP_KLAUZULA`, `OTKUP_ROK_ISPLATE`, `SELLER_LOGO_PATH`; absent/empty falls back to built-in defaults.

### Verification / acceptance gates

- Re-import `modDocStyle`, `modConfig`, `modPrint`, `modPaletniList`; generate otkupni + paletni + preradni PDFs.
- Otkupni list renders the full PDV-nadoknada klauzula (not clipped) and the rok from `OTKUP_ROK_ISPLATE`.
- Paletni/preradni show `Vrsta voca` as a subtitle with no `Vrsta` table column; cell `H1` equals the current layout version after first regeneration.

### Documentation actions

- `ARCHITECTURE_REFERENCE.md` section 5.12 (Document Print / Presentation Layer) added: shared `modDocStyle`, template table, layout-version rule, and the `OTKUP_KLAUZULA` / `OTKUP_ROK_ISPLATE` / `SELLER_LOGO_PATH` config keys.


## v6.25 — 2026-06-16

### Summary

Paletni List (pallet) domain added as a canonical local Excel/VBA workflow, plus a desktop-only setup mode.

### Added

- `tblPaleta`, `tblPaletaStavka`, `tblPrerada`, `tblPreradaStavka`, `tblTipPalete`, `tblTipAmbalaze` schema via `EnsurePaletniListSchema` (idempotent column repair).
- Transactional palletization from receipt save: `PaletizePrijemnica` runs inside `SavePrijemnica_TX` / `SavePrijemnicaMulti_TX` before `CommitTx`; both wrappers now snapshot `tblPaleta` and `tblPaletaStavka`.
- `SavePrerada_TX` processing flow (snapshots `tblPrerada`, `tblPreradaStavka`, `tblPaleta`).
- Pallet/prerada print/PDF (`PaletaSablon`, `PreradaSablon`) as post-commit side effects.
- Minimal pallet status (`lblPaletaStatus`) and a manual "print incomplete pallets" button in `frmDokumenta`.
- Desktop-only setup mode: `CheckGoogleOAuthConfig` is gated on `IsCloudSyncEnabled()`; `EnableDesktopOnlyMode` / `EnableCloudSyncMode` toggles.
- Pallet management UI (PR #44): `frmPalete` (read-models `GetPaleteForGrid` / `GetPaletaStavkeForGrid` / `GetPaletaStavkeForGridMulti`, `ClosePaletaManual_TX`); `modPaletniListUI` Alt+F8 entries. Build guide in `docs/frmPalete-build.md` (controls live in a binary `.frx`).
- Paletni/prerada list itemized per otkup: `GetOtkupiZaPalete` traces pallet -> zbirne -> `tblOtkup` (via `TraceByZbirna`/OtpremnicaID), filtered to the pallet's klasa, deduped by `OtkupID`; columns Rb | Kooperant (sifra) | Vrsta | Neto | Ambalaza (KolAmbalaze x TipAmbalaze).
- Storno: `modStorno.StornoPaleta_TX` (marks pallet + its stavke, refuses processed pallets) and `StornoPrerada_TX` (marks prerada + stavke, returns pallets to stock).

### Changed

- `tblPaletaStavka` keys palletization by `PrijemnicaID` (row identity), not `BrojPrijemnice`.
- Pallet/prerada critical reads use `RequireColumnIndex`; critical writes use `RequireUpdateCell`; ID lookups fail on `0` and `>1`.
- `modPaletniList` business functions contain no `MsgBox`.

### Fixed

- Receipt save no longer runs the pallet projection post-commit under a swallowed `On Error Resume Next`; palletization is atomic with the receipt.
- Receipt save performance regression (recalc storm + printer-bound `PageSetup`): palletization runs under the TX manual-calc guard, and `PageSetup` is wrapped with `Application.PrintCommunication`.

### Migration / data notes

- `EnsurePaletniListSchema` adds missing columns to existing pallet tables in place; no receipt/finance data migration required.
- `tblPrerada` adopts `NetoUlazKg` + `NetoIzlazKg`; a legacy `NetoKolicina` column (from the pre-release prerada increment) is left orphaned and unused.

### Verification / acceptance gates

- `EnsurePaletniListSchema` runs clean on a fresh workbook.
- Both receipt TX wrappers snapshot `tblPaleta` and `tblPaletaStavka`; palletization before `CommitTx`; print/PDF after.
- `PrijemnicaID` present in `tblPaletaStavka`; duplicate active palletization rejected.


## v6.24 — 2026-05-26

### Summary

v6.24 documents the systematic PWA design/runtime and VBA numbering work performed after v6.23. The release is a documentation supplement over the v6.23 package; it does not remove the v6.23 PWA otkup read-model convergence.

Scope:

- VBA document numbering model;
- PWA design tokens and self-hosted font system;
- reusable `components_v2.css` design system;
- Otkup form redesign;
- Otkupni list modal redesign;
- Otprema tab redesign;
- Pregled / Danas tab redesign;
- service-worker/cache and font asset updates;
- migration bug fixes and hardcoded-text cleanup;
- lazy-loading performance model;
- VBA / Google Sync diagnostics and known follow-ups.

### Added

- [Layer: VBA / Document numbering] Added/confirmed canonical business document numbering model for `BrojDokumenta`, `BrojOtpremnice` and `BrojZbirne` using the `x/ddmmyy[-rb]` convention.
- [Layer: PWA / Design system] Added brand token set for forest/accent/gold/cream/text/border/shadow/radius variables.
- [Layer: PWA / Design system] Added legacy token aliases for compatibility with existing code paths.
- [Layer: PWA / Fonts] Added self-hosted Cormorant Garamond display font and DM Sans body font architecture, including Latin-ext coverage for Serbian characters.
- [Layer: PWA / Components] Added `components_v2.css` reusable primitives for app headers, app body, step wizard, cards, scan CTA, kooperant chips, fields, pickers, buttons, sticky bars, hero/stat views, pills, list heads and record cards.
- [Layer: PWA / Otkup] Added 5-step Otkup form state-machine model and current helper surface:
  - `bindKlasaPicker`;
  - `bindTipAmbalazePicker`;
  - `evaluateOtkupFormState`;
  - `applyOtkupFormState`;
  - `bindOtkupFormStateListeners`;
  - `bindOtkupFormUIEvents`.
- [Layer: PWA / Otkupni list] Added `.ol-*` modal design contract while preserving required `data-action` hooks.
- [Layer: PWA / Otprema] Added redesigned Otprema views, detail modal, truck hero card, driver chip, pending summary, toolbar, selected-card state and sticky `Utovari` bar.
- [Layer: PWA / Pregled] Added redesigned Pregled/Danas header, stats grid, pills, date range, `.danas-*` cards and detail modal.
- [Layer: PWA / Runtime] Added `lazy.js` / `lazyLoadScript()` as Promise-based idempotent lazy-loading helper.

### Changed

- [Layer: PWA / Otkup form] Otkup UI changed from plain form groups to staged wizard/field/picker design.
- [Layer: PWA / Otkup form] Driver selection moved out of the Otkup form and remains in Otprema.
- [Layer: PWA / Otkup form] Class options are reduced to canonical `I` and `II`.
- [Layer: PWA / Otkup form] Inline picker script logic is replaced by event delegation to preserve CSP compatibility.
- [Layer: PWA / Service worker] Service-worker cache discipline now includes redesigned CSS/font assets and requires cache-version bump on critical app-shell changes.
- [Layer: PWA / Runtime] Heavy vendor assets are moved away from initial synchronous loading and toward feature-owned lazy loading.
- [Layer: PWA / Intercom] Firebase compat libraries are loaded only by feature paths that need them instead of every role.
- [Layer: PWA / Hardcoded UI cleanup] Station/role header text moves toward runtime config values such as `CONFIG.ENTITY_NAME`.

### Fixed

- Fixed iOS safe-area viewport support through `viewport-fit=cover`.
- Fixed sticky save bar overlap with bottom navigation / safe-area inset.
- Fixed date-range input overflow using `minmax(0, 1fr)`.
- Fixed Latin-ext rendering for Serbian characters in Cormorant.
- Fixed missing `.danas-*` card CSS after accidental deletion.
- Fixed `sync-engine.js` duplicate/triple `const entityID` declaration reported in the source summary.
- Fixed GAS `blockWriteIfMasterSyncActive` caller signature issue reported in the source summary.
- Fixed OTK operational header mismatch (`22` vs `23` columns) reported in the source summary.
- Fixed QR icon registry mismatch and missing `more` / `mic` icons.
- Fixed missing `.btn-v2--ghost` / `.btn-v2--sm` styles.
- Fixed hardcoded `OBLAČINA` role-eyebrow text by moving to dynamic entity-name helpers.
- Fixed `CONFIG.USERNAME` availability by exposing it in config.
- Fixed Pregled filter UI mismatch between `.pill .is-active` and old JS selector expectations.
- Fixed problem badge class mismatch by mapping the problem kind to an existing error badge class.
- Fixed Serbian comma decimal parsing through `parseDecimalInput`.
- Removed polling `setInterval(applyOtkupFormState, 500)` in favor of event-driven updates.

### Removed / Deprecated

- Deprecated role-local one-off UI styling where a shared `components_v2.css` primitive exists.
- Deprecated inline picker scripts for Otkup form behavior.
- Deprecated default synchronous loading of heavy vendor scripts in the app head.
- Deprecated hardcoded station/role text where config values exist.
- Deprecated `Sheet1`/default-tab assumptions where named tabs are canonical in the architecture.

### Known Issues / Deferred

- Otpremi hero description text remains intentionally unchanged even though it may overstate automatic truck/capacity recognition.
- `--accent` green/gold token split cleanup remains deferred to a final design-system sweep.
- Inline `style="display:none"` to `.is-hidden` migration remains cosmetic/P3 and requires JS audit.
- Remaining `var(--r-md)` / `var(--r-lg)` occurrences in `features-otkup.css` were reported as locally resolved but not Git-confirmed.
- `modGoogleSheets` `sheetId = 0` sentinel collision diagnosis remains status unknown in this documentation pass.

### Migration / Data Notes

- No historical business-data migration is required.
- PWA users may require cache refresh/service-worker update after deployment.
- Font assets must be present wherever the PWA app shell is served.
- Lazy-loaded vendor scripts must remain available at their expected paths.

### Verification / Acceptance Gates

- Source summary reports the changes as completed from origin/main `ce90970`.
- Direct Git verification against the connected GitHub repository could not confirm the AgriX/OtkupApp files because the available connector repository resolves to an Apartment Handover app, not the AgriX frontend.
- Required target-repo gates are listed in `RELEASE_GATES.md` under `v6.24 PWA Design/Runtime and Numbering Gates`.

### Documentation Actions

- [x] Architecture Reference updated with v6.24 PWA design/runtime and VBA numbering supplement.
- [x] Changelog updated with v6.24 Added/Changed/Fixed/Deferred items.
- [x] Release gates updated.
- [x] Known issues and roadmap updated for unresolved Git-confirmation and cleanup items.

---

## v6.23 — 2026-05-18

### Summary
v6.23 corrects the PWA otkup read model so Management and Otkupac views read both canonical master otkup data and operational PWA queue data. `tblOtkup` remains the canonical master. `MgmtReports/OtkupiAll` is the master read projection for PWA display. `OTK-ST-*` / `OTK-*` remains the operational inbound/live queue for PWA-origin otkup records.

Browser testing was reported as completed for the affected role views.

### Added

- Added `MgmtReports/OtkupiAll` as the canonical PWA read-model projection for otkup records from `tblOtkup`.
- Added PWA display merge of `OtkupiAll + OTK-ST-*` / `OTK-*` operational rows.
- Added explicit dedup priority for merged otkup display:
  1. `ServerRecordID` / `OtkupID`;
  2. `ClientRecordID`;
  3. natural-key fallback.
- Added role contract that Management and Otkupac otkup views must show both VBA/master-created otkupi and PWA/operational otkupi.

### Changed

- PWA otkup overview no longer treats `OTK-ST-*` / `OTK-*` as the only source for otkup display.
- `OTK-ST-*` / `OTK-*` is clarified as operational inbox/live queue, not canonical master history.
- Management and Otkupac views now use a master projection plus operational queue merge.
- Dedup priority for otkup overview is changed to prefer `ServerRecordID` / `OtkupID` before `ClientRecordID`, because synced PWA rows can carry different `ClientRecordID` values across `OTK-ST-*` and `OtkupiAll` while sharing the same `ServerRecordID`.
- `MgmtReports` read-model scope now includes `OtkupiAll` for PWA otkup overview.

### Fixed

- Fixed Management not seeing VBA-created otkupi.
- Fixed Otkupac not seeing VBA-created otkupi.
- Fixed PWA otkup display depending on operational `OTK-ST-*` sheets even though those sheets are inbound/live queue state.
- Fixed operational otkup rows disappearing from display when `OtkupiAll` exists.
- Fixed potential duplicate display of synced PWA otkup where `ClientRecordID` differs across `OTK-ST-*` and `OtkupiAll` but `ServerRecordID` is shared.

### Migration / Data Notes

- No historical business-data migration is required.
- `MgmtReports/OtkupiAll` export must be present and readable by GAS/PWA.
- Existing `OTK-ST-*` / `OTK-*` operational sheets remain active and are not removed.
- The change is a read-model and display merge correction, not a canonical master-data redesign.

### Verification / Acceptance Gates

Reported browser-tested behavior:

- Management sees VBA-created otkupi.
- Management sees PWA-created/synced otkupi.
- Management Partneri/Kooperanti/Kupci/report surfaces load expected data.
- Otkupac sees VBA-created otkupi in its scope.
- Otkupac sees operational `OTK-ST-*` / `OTK-*` rows.
- Relevant roles see otkupi from both worlds: VBA/master + PWA/operational.

Required regression gates for future changes:

- same synced otkup appearing in both sources renders once by `ServerRecordID` / `OtkupID`;
- if `ServerRecordID` is missing, `ClientRecordID` fallback works;
- natural-key fallback does not hide legitimate distinct records.

### Documentation Actions

- [x] Architecture Reference updated with `OtkupiAll` master read projection and operational queue merge rules.
- [x] Changelog updated with v6.23 entry.
- [x] Release gates updated with v6.23 PWA otkup read-model gates.

---

## v6.22 — 2026-05-15

### Summary
v6.22 is the residual GO hardening/documentation cut after the v6.21 GO closeout package. It carries forward the v6.21 Banka/Storno, Google Sheets, MasterSync, Novac and HealthCheck hardening and adds the pieces that were not explicit enough in the prior AR/CL package:

- `modFaktura` duplicate-`FakturaID` guards for print and status recompute paths;
- `modGeoParcele` ParcelaID-based save/clear APIs with exact-row lookup;
- explicit `RequireStornoAllowed` / checked-write storno helper pattern naming;
- gates for Faktura duplicate guards and Geo ByID save/clear.

### Added

#### `modFaktura`

- Added current architecture rule that `PrintFaktura` must reject duplicate `FakturaID` instead of accepting the first matching row.
- Added current architecture rule that `UpdateFakturaStatus` must reject duplicate `FakturaID` before recomputing payment/status state.

#### `modGeoParcele`

- Added canonical ParcelaID-based geo APIs:

```vb
SaveParcelGeoPointByID
SaveParcelGeoPointByID_TX
ClearParcelGeoByID
ClearParcelGeoByID_TX
RequireSingleParcelaRow
```

- Added exact-row `ParcelaID` rule for geo save/clear:

```text
Count = 0  => missing parcela error
Count = 1  => allowed
Count > 1  => duplicate ParcelaID error
```

#### `modStorno`

- Added explicit helper-pattern name for storno eligibility checks:

```text
RequireStornoAllowed / CanStorno / LookupActiveID
```

### Changed

#### `modFaktura`

- `PrintFaktura` and `UpdateFakturaStatus` are no longer documented as future-only hardening for duplicate `FakturaID`; they are now current contract.
- The earlier v6.17 roadmap item for FakturaID-based print/status duplicate guards is treated as closed for those two surfaces.

#### `modGeoParcele`

- Canonical geo save/clear no longer relies on physical `rowIndex` as the architecture-level target identity.
- UI code should capture selected `ParcelaID` and call the ByID API.
- Row-index functions remain compatibility wrappers only.

### Fixed

- Fixed documentation gap where Faktura print/status duplicate-key protection was still shown as roadmap despite being part of the GO hardening closeout.
- Fixed geo architecture fragility by making `ParcelaID` the canonical save/clear target instead of physical row index.
- Fixed documentation gap around the explicit `RequireStornoAllowed` / checked-write storno eligibility pattern.

### Removed / Deprecated

- Row-index parcel geo save/clear is deprecated as the preferred architecture. It may remain as a compatibility wrapper only.
- No business-data structures are removed by this documentation cut.

### Migration / Setup Notes

- No historical business-data migration is required.
- Existing geo callers should be reviewed: UI/front-door save/clear paths should use `SaveParcelGeoPointByID[_TX]` and `ClearParcelGeoByID[_TX]`.
- Existing row-index wrappers may remain during transition if they delegate to the ByID APIs.

### Verification / Acceptance Gates

Required gates:

- `PrintFaktura` duplicate-`FakturaID` negative smoke.
- `UpdateFakturaStatus` duplicate-`FakturaID` negative smoke.
- `SaveParcelGeoPointByID[_TX]` success and rollback smoke.
- `ClearParcelGeoByID[_TX]` success and rollback smoke.
- `RequireSingleParcelaRow` missing/duplicate negative smoke.
- Storno eligibility helper smoke for missing/duplicate/already-stornirano target rows.

Carry forward v6.21 GO validation evidence: full Google/PWA sync green with `Geo=True`, `Otkup=True`, `Otpremnice=True`, `Zbirne=True`, `Stammdaten=True`, `Kartice=True`, `MgmtReports=True` and `PWA master sync lock=NO`.

### Documentation Actions

- [x] Architecture Reference updated to v6.22.
- [x] Changelog updated with v6.22 entry.
- [x] Release gates updated for Faktura duplicate guards and Geo ByID APIs.
- [x] Roadmap updated to mark FakturaID print/status guard as closed and row-index geo as transitional.

---

## v6.21 — 2026-05-14

### Summary
- v6.21 documents the Banka/Storno production-hardening work performed after v6.20 and the later GO hardening closeout for Google Sync, MasterSync, Novac and ProductionHealthCheck.
- The release is VBA/Excel desktop focused with a small workstation setup implication.
- The active scope includes `modStorno`, `modBankaImportParserPdfToText`, `modBankaImport`, `modBankaMapiranje`, `modSetup`, `Setup-OtkupApp.ps1`, `modGoogleSheets`, `modStammdatenSync`, `modGoogleSyncOrchestrator`, `modMasterSync`, `modNovac` and `modProductionHealthCheck`.
- The release does not change GAS/PWA business semantics, but hardens Google sync/write behavior and lock/unlock outcome classification.
- The release does not require production data migration.
- The release must not delete files or unrelated code. Only functions/blocks touched by this v6.21 hardening work are replaced or extended.

### ADDED
- [Layer: VBA / `modStorno`] Added hardened storno architecture aligned with `modNovac` / `modFaktura`.
  - What changed: storno operations now use fail-fast row guards, `RequireColumnIndex`, `RequireUpdateCell`, transaction rollback and monitoring patterns.
  - Why: `modStorno` had valid business coverage but lower engineering hardening than newer finance/document modules.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modStorno`] Added/standardized strict helper surface for storno operations.
  - What changed: hardened helper patterns cover required input, exact-row checks, stornirano marking, side-effect repair and TX monitoring.
  - Why: storno operations must not silently update a wrong row or continue after missing schema/update failures.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImportParserPdfToText`] Added configurable `pdftotext.exe` resolution.
  - What changed: the parser resolves `PDFTOTEXT_EXE_PATH` through local workstation configuration rather than a user-specific hardcoded path.
  - Why: PDF import must work on a clean machine after setup and must not be tied to one developer workstation.
  - Reference update required: Yes
  - Migration required: No, but local config must be set where the default tools path is not used.

- [Layer: VBA / `modBankaImportParserPdfToText`] Added unique temp txt generation and cleanup.
  - What changed: each PDF extract uses a unique temp txt file and deletes it defensively before/after extraction.
  - Why: a failed `pdftotext` run must not read stale output from an earlier import.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImportParserPdfToText`] Added `pdftotext` process exit-code verification.
  - What changed: `WScript.Shell.Run` return value is captured; non-zero exit code is a hard error.
  - Why: PDF extract failure must go through the error path instead of producing a false processed import.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImport`] Added explicit import outcome categories.
  - What changed: import flow distinguishes `imported`, `duplicate-only`, `parse error`, `integrity error`, `append error`, `schema error`, `extract error` and `unknown error`.
  - Why: operators and rollback/error paths must be able to distinguish parse/integrity/staging/extract failures.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImport`] Added deferred file move infrastructure.
  - What changed: successful PDF file moves are recorded in a pending list and executed only after DB commit.
  - Why: file-system moves cannot rollback with the Excel transaction snapshot; processed files must represent committed staging.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaMapiranje`] Added exact-row guard helper `RequireSingleRow`.
  - What changed: ID links can require exactly one matching row and fail on missing/duplicate keys.
  - Why: `FindRows(...).Count > 0` can link the wrong row when duplicate IDs exist.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaMapiranje`] Added strict `LinkNovacToOtkupStrict`.
  - What changed: `NovacID -> OtkupID` link now checks exact `NovacID`, exact `OtkupID`, uses `RequireUpdateCell`, then recomputes otkup status.
  - Why: finance-document links must not silently attach to the first matching row.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA / `modStorno`] Storno flows now fail fast for missing/duplicate document IDs.
  - What changed: ID-based storno no longer treats "first found row" as sufficient.
  - Why: storno must be deterministic and reversible through transaction rollback.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImportParserPdfToText`] `ExtractTextFromPdf` no longer uses a static temp output path.
  - What changed: static `%TEMP%\pdf_extract.txt` behavior is replaced by unique per-import temp names.
  - Why: stale temp content could create a false-success import.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImportParserPdfToText`] `ExtractTextFromPdf` now fails if `pdftotext` fails.
  - What changed: non-zero process exit code and missing temp output are hard errors.
  - Why: failed extract must be visible and must not be staged as processed data.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modSetup`] Local workstation config ownership clarified.
  - What changed: `PDFTOTEXT_EXE_PATH` belongs in `tblLocalConfig` and should use the existing public `GetLocalConfigValue` from `modSetup`.
  - Why: `tblConfig` is reserved for Google/PWA config and must not be repurposed for local workstation settings.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PowerShell / `Setup-OtkupApp.ps1`] Workstation setup implication documented.
  - What changed: the standard setup root `C:\OtkupApp` is the natural fallback base for local tools such as Poppler/pdftotext.
  - Why: clean-machine install should have a predictable tools path or a clear config override.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImport`] `SaveBankaImportRows` is now fail-fast.
  - What changed: required columns use `RequireColumnIndex`, invalid input is rejected, `GetNextID` empty result fails, and `AppendRow <= 0` fails.
  - Why: financial import staging must not partially or silently fail.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaImport`] Banka import file moves are deferred until after transaction commit.
  - What changed: `ImportOnePdfIntoBankaImport` stages data and records pending moves; `ImportBankaInbox_TX` commits before executing moves.
  - Why: file-system state must not claim success before the DB staging transaction is committed.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaMapiranje`] Manual and auto-map paths share exact-row integrity.
  - What changed: status updates and critical links use `RequireSingleRow`; link writes use `RequireUpdateCell`.
  - Why: duplicate/missing keys must abort both manual and automatic mapping consistently.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaMapiranje`] `GetBankaImportRowByID` implementation is hardened without changing its public contract.
  - What changed: internally it uses `RequireSingleRow` and `RequireColumnIndex`, but still returns the existing 1x10 semantic result shape.
  - Why: callers rely on `bim(1, 1)` through `bim(1, 10)` meaning specific business fields.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modBankaMapiranje`] `ValidateBankaImportNotProcessed` checks soft-delete state again.
  - What changed: `COL_BIM_STORNIRANO = "DA"` makes the row ineligible for mapping/skip flow.
  - Why: stornirani BankaImport rows must not be mapped.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA / `modBankaImportParserPdfToText`] Fixed hardcoded local `pdftotext.exe` path risk.
  - Symptom: import could work only on one machine.
  - Resolution: executable path is local-config driven with setup-based fallback.
  - Reference update required: Yes

- [Layer: VBA / `modBankaImportParserPdfToText`] Fixed stale temp file risk.
  - Symptom: a failed PDF extract could read old `%TEMP%\pdf_extract.txt`.
  - Resolution: unique temp file per import plus defensive cleanup.
  - Reference update required: Yes

- [Layer: VBA / `modBankaImportParserPdfToText`] Fixed missing exit-code validation.
  - Symptom: `pdftotext` failure could continue as if extraction succeeded.
  - Resolution: non-zero exit code raises an error.
  - Reference update required: Yes

- [Layer: VBA / `modBankaImport`] Fixed silent staging failure risk.
  - Symptom: missing columns or `AppendRow` failure could produce unclear/partial import behavior.
  - Resolution: `RequireColumnIndex`, `GetNextID` and `AppendRow` hard-fail rules.
  - Reference update required: Yes

- [Layer: VBA / `modBankaImport`] Fixed processed-file-before-commit risk.
  - Symptom: a PDF could be moved to `Processed` before the batch transaction was reliably committed.
  - Resolution: successful file moves are deferred until after `CommitTx`.
  - Reference update required: Yes

- [Layer: VBA / `modBankaMapiranje`] Fixed duplicate-key link risk.
  - Symptom: `FindRows(...).Count > 0` could link `NovacID`, `BankaImportID`, `OtkupID` or `FakturaID` to the first matching row despite duplicates.
  - Resolution: exact-row `RequireSingleRow` guards.
  - Reference update required: Yes

- [Layer: VBA / `modBankaMapiranje`] Fixed `GetBankaImportRowByID` contract regression.
  - Symptom: a hardened implementation returned the raw `tblBankaImport` row instead of the legacy 1x10 business shape.
  - Resolution: the function keeps the legacy 1x10 shape while using strict column lookups internally.
  - Reference update required: Yes

- [Layer: VBA / `modBankaMapiranje`] Fixed `MapBankaImportAsKooperantBlockCore` double-count bug.
  - Symptom: nested `If novID <> "" Then` could count one candidate twice, reduce the remaining amount twice and skip the hard error when `novID` was empty.
  - Resolution: empty `novID` is a hard error; link/count/remainder update each happen once.
  - Reference update required: Yes

- [Layer: VBA / `modBankaMapiranje`] Fixed stornirani BankaImport eligibility regression.
  - Symptom: hardened validation checked `Obradjeno` but no longer checked `Stornirano`.
  - Resolution: `COL_BIM_STORNIRANO` guard is restored.
  - Reference update required: Yes

### UNCHANGED (explicit)
- [Layer: GAS/PWA] No GAS or PWA business contract change is included in v6.21.
- [Layer: Data] No production data migration is required.
- [Layer: Documentation/process] Existing v6.20 content remains the baseline and must not be shortened.
- [Layer: Codebase] No files are deleted as part of this release.
- [Layer: Codebase] Unrelated functions/sections are not removed; only v6.21-touched functions/blocks are replaced or extended.

### KNOWN LIMITATIONS
- `RequireSingleRow` is local to `modBankaMapiranje` in this release. A shared `modDataAccessGuards` consolidation remains recommended.
- Some business-layer `MsgBox` usage remains outside the specific 6.21 hardened blocks. Full UI/business separation is a future cleanup item.
- File moves after DB commit cannot be rolled back. If post-commit file move fails, operator/manual recovery is required.

### ROADMAP
- [RM-v6.21-01] Move exact-row guard helpers into a shared `modDataAccessGuards` module.
  - Affected modules: `modBankaMapiranje`, `modStorno`, future finance/document hardening.
  - Target state: one canonical exact-row guard implementation.

- [RM-v6.21-02] Add focused smoke/regression tests for Banka parser/import/mapiranje.
  - Coverage targets: non-zero `pdftotext` exit code, missing BankaImport column, `AppendRow <= 0`, duplicate key mapiranje and deferred move batch failure.

- [RM-v6.21-03] Continue moving UI prompts out of business logic where not directly part of this patch.
  - Target state: business modules raise errors/return results; forms decide operator messaging.

### VERIFIED / TO VERIFY
- Compile VBA after replacing touched functions/blocks.
- Parser smoke: valid PDF with configured `PDFTOTEXT_EXE_PATH`.
- Parser negative smoke: wrong `PDFTOTEXT_EXE_PATH` and/or forced non-zero `pdftotext` exit code.
- Import staging smoke: missing `tblBankaImport` column fails immediately.
- Import staging smoke: forced `AppendRow <= 0` rolls back transaction.
- Deferred move smoke: batch failure before commit does not move any PDF to `Processed`.
- Mapiranje smoke: duplicate `BankaImportID`, `NovacID`, `OtkupID` and `FakturaID` fail.
- Mapiranje smoke: stornirani BankaImport row cannot be mapped.
- Contract smoke: `GetBankaImportRowByID` returns the legacy 1x10 semantic shape.
- Storno smoke: missing/duplicate target IDs fail and rollback works.

### Migration / Data Notes
- No production data migration is required.
- Workstations that import bank PDFs must have `PDFTOTEXT_EXE_PATH` configured in `tblLocalConfig` unless the standard setup tools path is used.
- Recommended local config row:

```text
Kljuc: PDFTOTEXT_EXE_PATH
Vrednost: C:\OtkupApp\Tools\poppler\Library\bin\pdftotext.exe
Opis: Putanja do pdftotext.exe za PDF bankarske izvode
```

### Documentation Actions
- [x] Version index updated to v6.21.
- [x] Canonical reference updated with v6.21 Banka/Storno hardening baseline.
- [x] Changelog updated with v6.21 added/changed/fixed/verification items.
- [x] No unrelated v6.20 content removed.

---

---

### v6.21 GO Hardening Closeout — Google Sync / MasterSync / Novac / HealthCheck

#### Added

##### `modGoogleSheets.bas`

- Added staging/verify/replace write model for `WriteSheetData`.
- Added sheetId cache per spreadsheet.
- Added Google HTTP retry handling for `429`, `500`, `502`, `503` and `504`.
- Added write-request throttling.
- Added long wait handling for quota-window `429`.
- Added cache-aware `AddSheetTab`.
- Added phased target replacement:
  - rename target to backup;
  - rename staging to target;
  - delete backup.
- Added cache maintenance after rename/delete operations.

##### `modStammdatenSync.bas`

- Added canonical Kartice tab constant:

```vba
Private Const KARTICE_TAB_NAME As String = "Kartice"
```

- Added `EnsureKarticeTabsBestEffort`.
- Kartice export now ensures named tab `Kartice` before writing.

##### `modGoogleSyncOrchestrator.bas`

- Added failed-unlock handling aligned with GAS/PWA lock TTL behavior.
- Failed unlock now produces partial/degraded result instead of green success.
- Operator messaging references temporary TTL-based recovery rather than permanent lockout.

##### `modMasterSync.bas`

- Added exact-row helper `RequireSingleMasterSyncRow`.
- Added strict document-chain helpers:
  - `LinkOtkupToOtpremnicaStrict`;
  - `LinkOtpremnicaToBrojZbirneStrict`;
  - `GetBrojZbirneForIDStrict`.

##### `modNovac.bas`

- Added hard failure behavior in `SaveNovac` for empty `GetNextID` and `AppendRow <= 0`.

##### `modProductionHealthCheck.bas`

- Added duplicate document-key preflight check for `OtkupID`, `OtpremnicaID`, `ZbirnaID`, `PrijemnicaID`, `FakturaID`, `NovacID`, `BankaImportID` and `ParcelaID`.

#### Changed

##### `modGoogleSheets.bas`

- `WriteSheetData` no longer clears target tab before writing.
- `WriteSheetData` now writes to staging first and only replaces target after successful write/verify.
- `AddSheetTab` now checks whether a tab already exists before attempting `addSheet`.
- `GetSheetIdByTitle` uses cached metadata and forced refresh only when needed.
- Google write operations now use retry and throttle.
- Target replacement uses phased rename to avoid Google Sheets title collision.

##### `modStammdatenSync.bas`

- `ExportKarticeToGoogle_Core` no longer writes to `"Sheet1"`.
- All Kartice write/header paths now use `KARTICE_TAB_NAME`.
- Kartice export now behaves like Stammdaten/MgmtReports named-tab exports.

##### `modGoogleSyncOrchestrator.bas`

- Full sync cannot be green if PWA unlock fails.
- Unlock failure is classified as degraded/partial because GAS/PWA stale lock TTL can recover automatically.

##### `modMasterSync.bas`

- `AutoCreateOtpremniceFromPWA` now uses `RequireColumnIndex` for critical columns.
- `AutoCreateOtpremniceFromPWA` validates every grouped `OtkupID` before creating/linking an `Otpremnica`.
- Post-save `OtkupID -> OtpremnicaID` linking now uses exact-row checks and `RequireUpdateCell`.
- VOZ/Zbirna link logic now uses exact-row guards for `ZbirnaID`, `ClientRecordID`, `OtkupID` and `OtpremnicaID`.

##### `modNovac.bas`

- `SaveNovac` no longer returns `""` silently when append fails.
- Direct `SaveNovac` callers now receive hard error on append failure.

#### Fixed

- Fixed Google Sheets target-clear data-loss risk by replacing clear-before-write with staging/verify/replace.
- Fixed quota pressure from excessive metadata lookups through sheetId cache, no-op existing-tab handling, quota-aware retry and throttling.
- Fixed Google Sheets title collision risk during replace through phased target/backup/staging rename.
- Fixed Kartice export using `"Sheet1"` while the target named tab was `Kartice`.
- Fixed PWA unlock false-green risk.
- Fixed MasterSync document-chain duplicate risk from `FindRows(...).Count > 0` style linking.
- Fixed critical `SaveNovac` append silent-failure risk.
- Added health-check duplicate-key preflight to catch corrupted IDs before operator workflows.

#### Removed / Cleaned Up

- Removed temporary Kartice/`Sheet1` fallback helpers after switching Kartice to named tab:
  - `GetFirstSheetId`;
  - `ResolveTargetSheetIdForReplace`.

#### Validation

Full PWA / Google sync completed successfully with:

```text
Geo=True
Otkup=True
Otpremnice=True
Zbirne=True
Stammdaten=True
Kartice=True
MgmtReports=True
```

PWA lock was released successfully:

```text
SetPWAMasterSyncLock | PWA master sync lock=NO
```

#### Follow-Up / Non-Blocking Cleanup

- `modGoogleSheets.GoogleRetryDelayMs` appears unused after quota-aware retry cleanup.
- `modGoogleSheets.ExtractSheetIdByTitle` appears unused after sheetId cache implementation.
- These helpers are not GO blockers and can be removed later if compile/search confirms no callers.


## v6.20 — 2026-05-12

### Summary
- v6.20 is the permanent full-document correction after v6.18 and v6.19 were produced too short.
- The Architecture Reference and Changelog must be treated as standing documents, not temporary delta notes.
- v6.17 full canonical reference remains the structural base because it was the last acceptable complete reference.
- The BankaImport statement-saldo and statement-integrity hardening added on top of v6.17 is canonical in v6.20.
- v6.18 master-sync guard content and v6.19 parcel geo/editor hardening content are integrated into v6.20 as permanent, self-contained documentation.

### ADDED
- [Layer: Documentation / Architecture Reference] Added permanent v6.20 baseline rule.
  - What changed: v6.20 explicitly states that AR and CL are permanent self-contained documents and that v6.18/v6.19 must not remain shortened delta-only references.
  - Why: engineering and operations must be able to use the latest AR/CL without opening older partial documents.
  - Reference update required: Yes
  - Migration required: No

- [Layer: Documentation / Architecture Reference] Integrated full v6.18 architecture material into v6.20.
  - What changed: master-sync guard, SyncControl lock, GAS write blocker, PWA guard, soft-lock behavior, parcel geo pull and smoke/acceptance criteria are included in the v6.20 reference.
  - Why: v6.18 introduced active architecture and cannot remain only as a shortened standalone snapshot.
  - Reference update required: Yes
  - Migration required: No

- [Layer: Documentation / Architecture Reference] Integrated full v6.19 architecture material into v6.20.
  - What changed: frmStammdaten geo UX, modGeoParcele service ownership, transaction-backed geo save/clear, selected parcel Google sync and KPI hardening are included in the v6.20 reference.
  - Why: v6.19 introduced active architecture and cannot remain only as a shortened standalone snapshot.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / BankaImport] Added statement-level saldo metadata to the canonical bank import staging schema.
  - What changed: `tblBankaImport` includes `PocetnoStanje`, `ZavrsnoStanje`, `UkupanDuguje` and `UkupanPotrazuje`.
  - Why: every staged bank transaction must retain the statement-level totals used to prove the statement parsed correctly.
  - Reference update required: Yes
  - Migration required: Add/confirm columns in workbooks that do not yet contain them.

- [Layer: VBA / BankaImport PDF parser] Added `ExtractIzvodSaldoPdfText()` statement saldo extraction.
  - What changed: the parser extracts the `STANJE` block anchored by `Prethodno stanje` and reads `PocetnoStanje`, `UkupanDuguje`, `UkupanPotrazuje`, `ZavrsnoStanje`, `BrojNalogaZaduzenje` and `BrojNalogaOdobrenje`.
  - Why: header-only parsing cannot prove that all transaction rows and amounts were parsed correctly.
  - Reference update required: Yes
  - Migration required: No production-data migration; code/schema readiness required.

- [Layer: VBA / BankaImport parser integrity] Added four statement-level integrity gates before staging.
  - What changed: `ParseBankaIzvodForImport()` rejects the whole statement if math consistency, parsed uplata total, parsed isplata total or transaction counts do not match the parsed bank saldo block.
  - Why: a partially parsed bank statement must not enter mapping, partner-map learning, novac creation or document linking.
  - Reference update required: Yes
  - Migration required: No historical data migration.

### CHANGED
- [Layer: VBA / BankaImport] `ParseBankaIzvodForImport()` is now fail-fast on missing saldo block.
  - What changed: the parser still requires `BrojIzvoda`, `DatumIzvoda` and `BrojRacuna`, and now also requires a parseable `STANJE` saldo block.
  - Why: a statement without extracted saldo cannot be reconciled against the parser result.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / BankaImport] Staging rows now carry statement-saldo metadata.
  - What changed: parsed transaction rows copy the statement-level saldo fields onto each staged `tblBankaImport` row.
  - Why: each staged transaction remains auditable even after filtering, mapping, export or later review.
  - Reference update required: Yes
  - Migration required: Add/confirm columns if missing.

- [Layer: Documentation / Versioning] v6.20 supersedes v6.18/v6.19 shortened documentation style.
  - What changed: v6.20 uses a full v6.17-based document plus integrated v6.18/v6.19 content rather than publishing only a narrow delta.
  - Why: AR and CL must remain stable operating documents, not release-note fragments.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: Documentation] Fixed shortened v6.18/v6.19 documentation regression.
  - Symptom: v6.18 and v6.19 were too short to function as permanent canonical AR/CL documents.
  - Resolution: v6.20 keeps the full v6.17 base, preserves BankaImport additions, and integrates complete v6.18 and v6.19 materials.
  - Reference update required: Yes

- [Layer: VBA / BankaImport] Closed partial bank-statement staging risk.
  - Symptom: if PDF text parsing missed or misread rows, the import could still stage rows that looked structurally valid.
  - Resolution: statement totals and transaction counts are checked before any row is staged.
  - Reference update required: Yes

### Acceptance Criteria
v6.20 can be accepted when:

- Architecture Reference v6.20 is self-contained and does not require v6.18/v6.19 files to understand active architecture.
- Changelog v6.20 contains v6.18, v6.19 and v6.20 entries.
- `tblBankaImport` contains `PocetnoStanje`, `ZavrsnoStanje`, `UkupanDuguje`, `UkupanPotrazuje`.
- BankaImport PDF parser rejects a statement when saldo math does not match.
- BankaImport PDF parser rejects a statement when parsed totals do not match bank-reported totals.
- BankaImport PDF parser rejects a statement when parsed counts do not match bank-reported counts.
- v6.18 full-cycle master-sync smoke passes.
- v6.19 geo/editor/KPI smoke passes.

### Migration / Data Notes
- No historical business-data migration is required.
- Workbooks must have the four new `tblBankaImport` saldo columns before new v6.20 BankaImport imports are accepted.
- Existing historical BankaImport rows may have blank saldo fields unless backfilled from original statements; blank historical saldo is not automatically a production data error.

### Launch Decision
Recommended status:

```text
Launch Candidate after BankaImport statement-integrity smoke, v6.18 master-sync smoke and v6.19 geo/editor/KPI smoke pass.
```

---

## Integrated Permanent Changelog Section — v6.19

The following content is embedded in the v6.20 changelog so the latest CL remains self-contained.

# AgriX / OtkupApp Changelog — v6.19

**Version:** v6.19  
**Date:** 2026-05-12  
**Status:** Launch-candidate geo/Stammdaten UX hardening  
**Architecture reference:** `AGRIX_ARCHITECTURE_REFERENCE_v6_19.md`  
**Scope:** Changes after v6.18.

---

## 0. Summary

v6.19 is a focused VBA hardening release for the master-data parcel geo workflow and operator shell KPI robustness.

Main outcomes:

- `frmStammdaten` parcel geo controls are visible only for an actually selected parcel.
- Geo actions no longer use blocking `MsgBox` prompts during normal geo flow.
- A non-blocking geo status label is now the canonical feedback surface for parcel geo actions.
- Geo save/clear preserve the selected parcel after list refresh.
- Polygon editor launch now first syncs the selected local parcel to Google, then opens the editor.
- `modGeoParcele` is promoted to a domain/service-level geo module, aligned with the transactional pattern used by higher-criticality business modules.
- `SaveParcelGeoPoint` and `ClearParcelGeo` now use transaction-backed `_TX` implementations with rollback support.
- `SyncSelectedParcelaToGoogle` provides the canonical service entry point for “selected parcel → Google → polygon editor”.
- `frmOtkupAPP` KPI helpers were hardened against type mismatch on dirty date/kg data.

v6.19 does not replace the v6.18 master-sync contract. It builds on it by making the parcel geo/editor workflow safer and more operator-friendly.

---

## 1. Added

### 1.1 [VBA] `modGeoParcele.SaveParcelGeoPoint_TX`

Added transaction-backed geo point save function:

```vb
Public Function SaveParcelGeoPoint_TX(ByVal rowIndex As Long, _
                                      ByVal nCoord As Double, _
                                      ByVal eCoord As Double) As Boolean
```

The function:

- validates `rowIndex`
- validates positive N/E coordinates
- converts UTM zone 34 coordinates to `Lat`/`Lng`
- starts `clsTransaction`
- snapshots `tblParcele`
- writes geo point fields through `RequireUpdateCell`
- commits on success
- rolls back on failure

Updated fields:

```text
N_Coord
E_Coord
Lat
Lng
GeoStatus
GeoSource
MeteoEnabled
DatumGeoUnosa
DatumAzuriranja
```

Canonical values on save:

```text
GeoStatus = point
GeoSource = selenium
MeteoEnabled = Da
```

**Reason:**  
The old implementation used multiple independent `UpdateCell` calls. If one write failed mid-sequence, `tblParcele` could contain a partially updated geo state.

**Impact:**  
Geo point save now follows the same transactional discipline expected from domain/business modules.

**Migration:**  
No data migration.

---

### 1.2 [VBA] `modGeoParcele.ClearParcelGeo_TX`

Added transaction-backed geo clear function:

```vb
Public Function ClearParcelGeo_TX(ByVal rowIndex As Long) As Boolean
```

The function:

- validates `rowIndex`
- starts `clsTransaction`
- snapshots `tblParcele`
- clears point fields through `RequireUpdateCell`
- updates geo status fields
- commits on success
- rolls back on failure

Canonical values on clear:

```text
GeoStatus = none
GeoSource = empty
MeteoEnabled = Ne
DatumAzuriranja = Now
```

**Reason:**  
Clearing geo data is also a multi-field mutation and must not leave partial state.

**Impact:**  
Geo clear becomes rollback-safe.

**Migration:**  
No data migration.

---

### 1.3 [VBA] Compatibility wrappers for geo save/clear

Existing public procedure names remain available:

```vb
Public Sub SaveParcelGeoPoint(ByVal rowIndex As Long, ByVal nCoord As Double, ByVal eCoord As Double)
Public Sub ClearParcelGeo(ByVal rowIndex As Long)
```

They now call their `_TX` equivalents and raise an error if the transaction function returns `False`.

**Reason:**  
Existing form code can keep calling stable public names while the implementation becomes transaction-safe.

**Impact:**  
No form-level rename required.

**Migration:**  
No migration.

---

### 1.4 [VBA] Selected parcel Google sync service

Added canonical geo service entry point:

```vb
Public Function SyncSelectedParcelaToGoogle(ByVal parcelaID As String) As Boolean
```

The function:

1. validates `ParcelaID`
2. confirms the parcel exists locally in `tblParcele`
3. calls the existing canonical parcel export wrapper:

```vb
SyncParceleToGoogle_Core(False)
```

4. verifies that the parcel exists in Google `Stammdaten / Parcele` after sync
5. returns `True` only when the editor can safely read the selected parcel online

**Reason:**  
The polygon editor reads the Google/online layer, not the Excel workbook. If the operator just saved coordinates locally, the editor must not open stale or empty Google data.

**Impact:**  
The editor launch flow becomes:

```text
selected tblParcele row
↓
SyncSelectedParcelaToGoogle
↓
Stammdaten / Parcele contains selected ParcelaID
↓
open parcel-draw.html?parcelaId=...
```

**Migration:**  
Requires `SyncParceleToGoogle_Core(False)` to exist in `modStammdatenSync` as a public wrapper around the existing canonical `ExportParcele` logic.

---

### 1.5 [VBA] Local/Google parcel existence checks

Added helper checks in `modGeoParcele`:

```vb
Private Function LocalParcelaExists(ByVal parcelaID As String) As Boolean
Private Function GoogleParcelaExists(ByVal parcelaID As String) As Boolean
Private Function FindHeaderColumnInArray(ByVal data As Variant, ByVal headerName As String) As Long
Private Function SafeArrayText(ByVal data As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
```

**Reason:**  
`SyncSelectedParcelaToGoogle` must fail closed if the selected parcel cannot be proven locally or in Google after sync.

**Impact:**  
Polygon editor is not opened against missing/stale Google data.

---

### 1.6 [VBA] `frmStammdaten` geo status surface

Added non-blocking geo status helpers:

```vb
Private Sub SetGeoStatus(ByVal message As String, Optional ByVal isError As Boolean = False)
Private Sub ClearGeoStatus()
Private Sub ResetGeoClearConfirm()
Private Function HasSelectedParcelaForGeo() As Boolean
```

Expected UI control:

```text
lblGeoStatus
```

**Reason:**  
Geo button flow should not be interrupted by `MsgBox` for normal recoverable states.

**Impact:**  
Operators see inline geo status/errors without modal interruption.

---

### 1.7 [VBA] Parcel selection preservation helpers

Added helpers to preserve selected parcel after list refresh:

```vb
Private Function GetSelectedParcelaID() As String
Private Sub ReselectParcelaInList(ByVal parcelaID As String)
```

**Reason:**  
`LoadList` rebuilds the ListBox and row map. After save/clear geo, the selected parcel must remain selected, otherwise geo controls can enter a confusing state.

**Impact:**  
After saving or clearing geo data, the same parcel remains selected and geo actions remain coherent.

---

## 2. Changed

### 2.1 [VBA] `frmStammdaten` geo control visibility

Changed geo controls to be visible only when all conditions are true:

```vb
Me.Tag = "Parcele"
m_SelectedRow > 0
lstData.ListIndex >= 0
```

Canonical helper:

```vb
Private Sub UpdateGeoControlsVisibility()
    SetGeoControlsVisible (Me.Tag = "Parcele" And _
                           m_SelectedRow > 0 And _
                           lstData.ListIndex >= 0)
End Sub
```

**Reason:**  
Geo buttons were visible across modes or when no concrete parcel was selected.

**Impact:**  
Geo actions are only available for a real selected parcel.

---

### 2.2 [VBA] `frmStammdaten` geo buttons no longer show normal-flow MsgBox prompts

Changed geo button UX from modal messages to inline status messages:

- `btnGeoOpen_Click`
- `btnGeoSave_Click`
- `btnGeoClear_Click`
- `btnPasteCoords_Click`
- `btnOpenMap_Click`
- `btnOpenPolygonEditor_Click`

**Reason:**  
Geo workflow is an operator micro-flow. Repeated modal popups slow down the workflow and create poor UX.

**Impact:**  
Normal geo issues such as empty clipboard, missing selection, invalid coordinates, no Lat/Lng, and sync failure are shown in `lblGeoStatus`.

---

### 2.3 [VBA] Geo clear confirmation changed from modal prompt to second-click confirmation

Changed delete/clear confirmation:

Old behavior:

```text
MsgBox Yes/No confirmation
```

New behavior:

```text
First click: sets button caption to "Potvrdi brisanje" and shows status warning.
Second click: clears geo data.
```

State flag:

```vb
Private mGeoClearConfirmPending As Boolean
```

**Reason:**  
Keeps geo flow non-modal while still protecting destructive action.

**Impact:**  
No blocking prompt for clear; accidental click still protected.

---

### 2.4 [VBA] Polygon editor launch flow

Changed `btnOpenPolygonEditor_Click` to require selected-parcel sync before opening the editor:

```vb
If Not SyncSelectedParcelaToGoogle(parcelaID) Then
    SetGeoStatus "Parcela nije sinhronizovana. Editor nije otvoren.", True
    Exit Sub
End If

If OpenParcelPolygonEditor(parcelaID) Then
    SetGeoStatus "Polygon editor otvoren.", False
End If
```

**Reason:**  
The editor must open the coordinates/row just saved in `tblParcele`, not stale/empty data currently in Google Sheets.

**Impact:**  
Operator action now has deterministic data flow:

```text
local selected parcel → Google Parcele tab → polygon editor
```

---

### 2.5 [VBA] `OpenGoogleMaps` / `OpenParcelPolygonEditor` return Boolean

Changed helper procedures from fire-and-forget `Sub` style to success-aware `Function` style:

```vb
Public Function OpenGoogleMaps(ByVal lat As Double, ByVal lng As Double) As Boolean
Public Function OpenParcelPolygonEditor(ByVal parcelaID As String) As Boolean
```

**Reason:**  
Callers need to show inline status without raising modal errors.

**Impact:**  
Geo button handlers can display success/failure in `lblGeoStatus`.

---

### 2.6 [VBA] `btnPasteCoords_Click` now requires selected parcel

Changed paste-coordinates flow to require `Parcele` mode and selected parcel:

```vb
If Me.Tag <> "Parcele" Then Exit Sub
If Not HasSelectedParcelaForGeo() Then Exit Sub
```

**Reason:**  
Paste coordinates is a parcel-specific action and should not run without a selected parcel.

**Impact:**  
More consistent geo guard behavior across all geo buttons.

---

### 2.7 [VBA] `frmOtkupAPP` KPI helpers hardened against dirty dates/numbers

Hardened KPI/date helper behavior for functions such as:

```vb
SumOtkupKgForDate
CountDocsForDate
RefreshSidebarKpi
```

Recommended safe helper pattern:

```vb
SafeDateKey
SafeKpiDouble
```

**Reason:**  
Dashboard KPI refresh produced runtime error `13 | Type mismatch` when a date/kg field contained an unexpected value, blank, error, or non-date text.

**Impact:**  
KPI widgets skip invalid rows and return safe zero/summary values instead of logging repeat runtime errors.

---

## 3. Fixed

### 3.1 [Geo UX] Geo buttons visible outside valid parcel selection

**Problem:**  
Geo buttons could remain visible in contexts where no parcel was selected.

**Fix:**  
`UpdateGeoControlsVisibility` now requires both `m_SelectedRow > 0` and `lstData.ListIndex >= 0` in `Parcele` mode.

**Result:**  
Geo UI is hidden until a concrete parcel is selected.

---

### 3.2 [Geo UX] Modal popup noise in geo workflow

**Problem:**  
Geo button flow used `MsgBox` for normal states such as missing data, invalid coordinates, empty clipboard, sync failure, and success confirmation.

**Fix:**  
Geo flow now uses `SetGeoStatus` / `lblGeoStatus` for inline feedback.

**Result:**  
Operator workflow is faster and less disruptive.

---

### 3.3 [Geo UX] Selected parcel lost after geo save/clear refresh

**Problem:**  
`LoadList` refreshes the ListBox and row map after geo save/clear. The same parcel could lose active selection.

**Fix:**  
Save/clear now captures selected `ParcelaID`, reloads the list, and reselects the same parcel.

**Result:**  
Selected parcel remains stable after geo mutation.

---

### 3.4 [Geo/Data] Polygon editor could open stale or empty Google parcel data

**Problem:**  
After saving coordinates locally in `tblParcele`, opening the polygon editor could still show stale data from Google `Stammdaten / Parcele`, or blank state if the row had not been exported.

**Fix:**  
`btnOpenPolygonEditor_Click` now calls `SyncSelectedParcelaToGoogle(parcelaID)` before opening the editor.

**Result:**  
The editor reads the current selected parcel state from Google.

---

### 3.5 [Geo/Data] Multi-field geo save/clear could be partial

**Problem:**  
Old geo save/clear used repeated `UpdateCell` calls without transaction rollback.

**Fix:**  
`SaveParcelGeoPoint_TX` and `ClearParcelGeo_TX` use `clsTransaction`, `AddTableSnapshot`, `RequireUpdateCell`, `CommitTx`, and rollback on failure.

**Result:**  
Geo point save/clear are atomic at the `tblParcele` table level.

---

### 3.6 [Dashboard] `frmOtkupAPP` KPI type mismatch

**Problem:**  
KPI refresh could log repeated `Type mismatch` errors from date/count/sum helper functions.

**Fix:**  
KPI helpers were made tolerant of invalid date/numeric values.

**Result:**  
Bad/blank/error cells no longer break sidebar KPI refresh.

---

## 4. Deprecated

### 4.1 Direct full Stammdaten export from polygon editor button

Deprecated operator flow:

```vb
SyncStammdatenToGoogle_Core(False)
```

from `btnOpenPolygonEditor_Click`.

Replacement:

```vb
SyncSelectedParcelaToGoogle(parcelaID)
```

**Reason:**  
The editor button needs a selected-parcel sync service, not a broad all-Stammdaten operator export.

---

### 4.2 Geo button MsgBox prompts for normal recoverable states

Deprecated pattern inside geo button handlers:

```vb
MsgBox "...", vbExclamation, APP_NAME
```

Replacement:

```vb
SetGeoStatus "...", True
```

**Reason:**  
Geo flow should be non-modal except for wider form-level critical errors.

---

### 4.3 Raw multi-field `UpdateCell` geo mutations

Deprecated pattern:

```vb
UpdateCell TBL_PARCELE, rowIndex, "N_Coord", ...
UpdateCell TBL_PARCELE, rowIndex, "E_Coord", ...
...
```

Replacement:

```vb
SaveParcelGeoPoint_TX
ClearParcelGeo_TX
```

**Reason:**  
Multi-field parcel geo mutations need rollback discipline.

---

## 5. Operational Notes

### 5.1 Deployment order

Recommended order:

1. Import updated `modGeoParcele`.
2. Confirm `modStammdatenSync.SyncParceleToGoogle_Core(False)` exists and compiles.
3. Import updated `frmStammdaten`.
4. Confirm `lblGeoStatus` exists on the form.
5. Compile all VBA.
6. Run geo smoke tests.
7. Export VBA modules and tag/archive v6.19.

---

### 5.2 Required form controls

`frmStammdaten` expects these geo controls to exist:

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

`lblGeoStatus` is the canonical non-modal status surface.

---

### 5.3 Geo smoke tests

Minimum smoke:

1. Open `frmStammdaten` with `Tag = "Parcele"`.
2. Verify geo controls are hidden before parcel selection.
3. Select a parcel.
4. Verify geo controls become visible.
5. Click `Otvori GeoData` / GeoSrbija and verify search text is copied.
6. Paste N/E coordinates.
7. Save geo point.
8. Verify same parcel remains selected.
9. Verify `Lat`/`Lng`, `GeoStatus`, `GeoSource`, `MeteoEnabled`, dates update in `tblParcele`.
10. Open Google Maps.
11. Open polygon editor.
12. Verify selected parcel exists in Google `Stammdaten / Parcele` before editor opens.
13. Clear geo with two-click confirmation.
14. Verify same parcel remains selected and point fields are cleared.

---

### 5.4 KPI smoke tests

Minimum smoke:

1. Add or simulate dirty date/kg values in relevant tables.
2. Open `frmOtkupAPP`.
3. Refresh sidebar KPI.
4. Verify no `13 | Type mismatch` log entries from `SumOtkupKgForDate` or `CountDocsForDate`.
5. Verify KPI shows safe zero or valid aggregate.

---

## 6. Known Issues / Remaining Risks

### KI-6.19-01 — Selected parcel sync currently depends on `SyncParceleToGoogle_Core`

**Risk:**  
`SyncSelectedParcelaToGoogle` depends on `SyncParceleToGoogle_Core(False)` being available and correctly implemented in `modStammdatenSync`.

**Mitigation:**  
Compile all VBA and run selected-parcel sync smoke.

---

### KI-6.19-02 — Selected parcel sync may still use full Parcele-tab export internally

**Risk:**  
The selected parcel service may still delegate to full `ExportParcele` through the existing canonical export path. This is intentionally chosen to avoid duplicate export mapping, but it is broader than a true one-row upsert.

**Mitigation:**  
Accepted for v6.19 because it reuses canonical mapping and was reported fast enough. Consider true row upsert only if field data volume makes full Parcele export too slow.

---

### KI-6.19-03 — Polygon overwrite safety still governed by v6.18 full-cycle rule

**Risk:**  
If local `PolygonGeoJSON` is empty and selected-parcel sync performs a full export, existing Google polygon data can still be overwritten unless the canonical export/pull discipline is respected.

**Mitigation:**  
Keep v6.18 rule: full-cycle sync must import Google parcel geo/polygon into master before outbound Stammdaten export. For editor pre-open selected sync, confirm expected behavior with real polygon smoke.

---

### KI-6.19-04 — `lblGeoStatus` missing from designer

**Risk:**  
If the form designer does not contain `lblGeoStatus`, inline status will not show. The code uses defensive `Controls("lblGeoStatus")`, but UX will be silent.

**Mitigation:**  
Verify form designer contains `lblGeoStatus`.

---

## 7. Acceptance Criteria

v6.19 can be accepted when:

- VBA compiles cleanly.
- `frmStammdaten` opens in all supported modes.
- Geo controls are hidden in non-`Parcele` modes.
- Geo controls are hidden in `Parcele` mode until a row is selected.
- Geo flow uses `lblGeoStatus`, not normal-flow `MsgBox` prompts.
- Save geo point writes all required fields transactionally.
- Clear geo writes all required fields transactionally.
- Save/clear rollback correctly on forced error.
- Save/clear preserve selected parcel after `LoadList`.
- `SyncSelectedParcelaToGoogle` returns `True` only when selected parcel exists in Google after sync.
- Polygon editor opens only after successful selected-parcel sync.
- `frmOtkupAPP` KPI refresh no longer logs type mismatch errors.

---

## 8. Launch Decision

Recommended status:

```text
Launch Candidate
```

Reason:

- parcel geo/editor operator flow is now coherent
- geo mutations are transaction-backed
- modal geo UX noise is removed
- selected parcel remains stable through list refresh
- editor no longer opens against stale/empty Google data
- KPI shell robustness improved

Final production launch should still wait for the geo smoke, polygon smoke, KPI smoke, and full v6.18 master-sync smoke.

---

## Integrated Permanent Changelog Section — v6.18

The following content is embedded in the v6.20 changelog so the latest CL remains self-contained.

# AgriX / OtkupApp Changelog — v6.18

**Version:** v6.18  
**Date:** 2026-05-11  
**Status:** Launch-candidate hardening  
**Architecture reference:** `AGRIX_ARCHITECTURE_REFERENCE_v6_18.md`  
**Scope:** Changes after v6.17.

---

## 0. Summary

v6.18 delivers a coordinated VBA/PWA/GAS master-sync guard and closes the major launch blocker around PWA writes during VBA master sync.

Main outcomes:

- VBA now owns one full-cycle PWA/Google sync entrypoint.
- VBA publishes a lock into `Stammdaten / SyncControl`.
- GAS exposes the lock through `getMasterSyncState`.
- GAS blocks write actions while master sync is active.
- PWA shows an operator-visible sync overlay.
- PWA no longer needs to treat `MASTER_SYNC_ACTIVE` as a true sync error.
- Local PWA work can continue in offline-first pending mode.
- Parcel geo/polygon data is pulled into master before Stammdaten export, preventing overwrites.
- Operator sync button has live progress feedback.

This version is a launch-candidate sync-hardening release.

---

## 1. Added

### 1.1 [VBA] Full PWA/Google sync orchestrator

Added canonical orchestration layer for full PWA/Google cycle.

Canonical procedures:

```vb
Public Sub SyncPWAFullCycle()
Public Function SyncPWAFullCycle_ForButton() As Boolean
Private Function SyncPWAFullCycle_Core(ByVal showMessages As Boolean) As Boolean
```

The orchestrator owns the ordered flow:

1. Google config/auth check
2. lock PWA writes
3. pull parcel geo from Google
4. import OTK/PWA data
5. auto-create/link Otpremnice
6. import VOZ/Zbirna data
7. export Stammdaten
8. export Kartice
9. export MgmtReports
10. unlock PWA writes
11. monitor result

**Reason:**  
The operator needs a single safe master caller instead of multiple separate sync buttons/procedures.

**Impact:**  
Manual sync should go through `SyncPWAFullCycle_ForButton`.

**Migration:**  
No data migration.

---

### 1.2 [VBA] PWA master-sync lock writer

Added lock writer that updates:

```text
Stammdaten / SyncControl
```

Canonical keys:

```text
MASTER_SYNC_LOCK
MASTER_SYNC_UPDATED_AT
MASTER_SYNC_MESSAGE
MASTER_SYNC_OWNER
```

Expected state during sync:

```text
MASTER_SYNC_LOCK = YES
MASTER_SYNC_OWNER = VBA
```

Expected state after sync:

```text
MASTER_SYNC_LOCK = NO
```

**Reason:**  
PWA and GAS need a shared signal that VBA master sync is active.

**Impact:**  
GAS and PWA can coordinate around VBA sync.

**Migration:**  
No data migration. `SyncControl` tab is created/updated as needed.

---

### 1.3 [VBA] Parcel geo pull before Stammdaten export

Added/imported function:

```vb
Public Function ImportParcelGeoFromGoogleToMaster() As Boolean
```

This pulls geo/polygon data from Google Stammdaten `Parcele` into `tblParcele`.

**Reason:**  
Polygons created in PWA/HTML/GAS flow could be overwritten by stale VBA Stammdaten export.

**Impact:**  
Full-cycle sync must run geo pull before outbound Stammdaten.

**Migration:**  
No structural migration, but final smoke required with real polygon edits.

---

### 1.4 [VBA] Core functions for orchestration

Added/standardized core functions:

```vb
Public Function SyncStammdatenToGoogle_Core(ByVal showMessages As Boolean) As Boolean
Public Function ExportKarticeToGoogle_Core(ByVal showMessages As Boolean) As Boolean
Public Function ExportMgmtReports_Core(ByVal showMessages As Boolean) As Boolean
Public Function ImportOtkupFromPWA_Core(ByVal showMessages As Boolean) As Boolean
Public Function ImportZbirneFromPWA_Core(ByVal showMessages As Boolean) As Boolean
```

Public UI subs remain thin wrappers.

**Reason:**  
The orchestrator must call sync/export/import logic without unnecessary UI popups.

**Impact:**  
Cleaner separation between UI and automation.

**Migration:**  
No data migration.

---

### 1.5 [VBA] Live sync progress log

Added progress/log mechanism for the PWA full-cycle sync.

Canonical orchestrator helper:

```vb
SyncProgress "message"
```

Expected form methods:

```vb
BeginPWASyncLog
AppendPWASyncLog
EndPWASyncLog
```

**Reason:**  
Long Google/VBA sync operations can look like Excel is frozen.

**Impact:**  
Operator sees what the system is doing.

**Migration:**  
No data migration.

---

### 1.6 [VBA] Scheduled auto-sync

Added scheduled sync support using `Application.OnTime`.

Canonical functions:

```vb
StartScheduledSync
StopScheduledSync
ScheduledSyncTick
```

Config key:

```text
SYNC_AUTO_INTERVAL_MIN
```

Rules:

```text
empty/0 = disabled
< 15 = disabled/rejected
>= 15 = enabled
```

**Reason:**  
Auto sync should run from timers after application start, not directly from `Workbook_Open`/`BeforeClose`.

**Impact:**  
App startup/shutdown stays responsive.

**Migration:**  
No data migration.

---

### 1.7 [GAS] Master sync state endpoint

Added public read action:

```js
getMasterSyncState
```

Returns:

```js
success
locked
stale
updatedAt
ageMin
message
```

**Reason:**  
PWA needs a safe read-only endpoint to know whether VBA sync is active.

**Impact:**  
PWA can show overlay and avoid write attempts during lock.

**Migration:**  
Deploy GAS update.

---

### 1.8 [GAS] Master sync write blocker

Added server-side guard:

```js
blockWriteIfMasterSyncActive(action)
isMasterSyncWriteAction(action)
```

Blocked write actions include:

```text
sync
syncAgromere
syncZbirna
syncTretman
syncOprema
syncTrosak
saveParcelPolygon
uploadPdf
saveWarRoomDemand
removeWarRoomDemand
updateDemandPrimljeno
updateKamionStatus
saveDispecer
updateDispecer
removeDispecer
saveIzdavanje
saveFiskalni
saveFiskalniMapiranje
createArtikal
```

Blocked response:

```json
{
  "success": false,
  "locked": true,
  "code": "MASTER_SYNC_ACTIVE"
}
```

**Reason:**  
Client-side guard is useful, but server-side enforcement is mandatory.

**Impact:**  
No PWA write can mutate Google data during active VBA master sync.

**Migration:**  
Deploy GAS update.

---

### 1.9 [PWA] `master-sync-guard.js`

Added runtime guard:

```text
src/js/utils/master-sync-guard.js
```

Exposes:

```js
getMasterSyncStateSafe(force)
ensureMasterSyncNotActive(context, options)
startMasterSyncGuardPolling()
```

Includes:

- state polling
- cache
- overlay
- manual re-check
- offline/unknown handling

**Reason:**  
PWA UI needs to know when master sync is active.

**Impact:**  
PWA can inform user and avoid unnecessary write attempts.

**Migration:**  
PWA deploy + service worker cache bump.

---

## 2. Changed

### 2.1 [PWA] `index.html` runtime order

Changed runtime include order to ensure `master-sync-guard.js` is loaded.

Required conceptual order:

```html
<script src="./src/js/services/api.js"></script>
<script src="./src/js/utils/master-sync-guard.js"></script>
<script src="./src/js/utils/async.js"></script>
<script src="./src/js/utils/sync-engine.js"></script>
```

**Reason:**  
The guard must exist before `withSubmitLock`/`syncStore` uses it.

**Impact:**  
Client guard now actually runs.

---

### 2.2 [PWA] `sync-engine.js` handling for `MASTER_SYNC_ACTIVE`

Changed behavior when GAS returns:

```js
code === 'MASTER_SYNC_ACTIVE'
```

Old behavior:

```text
treated as sync error
record could appear in sync errors
```

New behavior:

```text
record stays pending/retryable
lastSyncError is cleared
lastServerStatus = master-sync-active
failed count is not increased for this condition
```

Recommended local state:

```js
record.syncStatus = 'pending';
record.lastServerStatus = 'master-sync-active';
record.lastSyncError = '';
```

**Reason:**  
Master sync lock is temporary coordination state, not data error.

**Impact:**  
Operators do not see false sync errors for records created during VBA sync.

---

### 2.3 [PWA] Soft-lock policy

Changed recommended behavior from hard blocking all local work to soft-locking server writes.

New policy:

```text
Allow local IndexedDB capture.
Block/skip GAS upload while master sync is active.
Retry later.
```

**Reason:**  
Field work should not stop unnecessarily.

**Impact:**  
PWA can keep working during VBA sync; records upload later.

---

### 2.4 [PWA] Service worker cache

Changed service worker cache to include the new guard file and require version bump.

Required cached file:

```text
src/js/utils/master-sync-guard.js
```

**Reason:**  
Field devices must load the new guard.

**Impact:**  
Users may need refresh/update after deployment.

---

### 2.5 [VBA] `frmOtkupAPP` sync button

Changed sync button flow to call:

```vb
SyncPWAFullCycle_ForButton
```

instead of individual import/export routines.

**Reason:**  
One operator action must execute the full safe sync sequence.

**Impact:**  
Better operator flow and less chance of incomplete sync.

---

### 2.6 [VBA] Startup/shutdown sync behavior

Changed auto-sync logic so it belongs to app lifecycle:

```vb
StartApp -> StartScheduledSync
ShutdownApp -> StopScheduledSync
```

Not direct workbook-open/before-close business sync.

**Reason:**  
Avoid blocking workbook startup/shutdown.

**Impact:**  
More stable runtime.

---

### 2.7 [VBA] Stammdaten export empty dataset behavior

Changed multiple export functions to prefer header-only output for legitimate empty datasets.

Canonical helper:

```vb
WriteHeaderOnly(sheetID, tabName, headers...)
```

**Reason:**  
Empty reports/tables should not automatically mean failed sync.

**Impact:**  
Fewer false partial failures.

---

### 2.8 [VBA] Active flag normalization

Changed PWA exports to use normalized active-check logic.

Inactive values:

```text
NE
NO
FALSE
0
NEAKTIVAN
INACTIVE
```

All other values are active.

**Reason:**  
Tables use mixed active markers.

**Impact:**  
More consistent PWA Stammdaten output.

---

### 2.9 [VBA] Config export filtering

Changed PWA config export safety to exclude internal sync keys.

Excluded prefixes:

```text
GOOGLE_
SEF_
SYNC_
```

**Reason:**  
Runtime sync config and credentials must not be exposed as PWA business config.

**Impact:**  
Safer PWA Config export.

---

## 3. Fixed

### 3.1 [Sync] PWA record created during VBA sync looked like sync error

**Problem:**  
When PWA saved a record while VBA sync was active, the record could be saved locally but fail upload due to lock/server state and appear as sync error.

**Fix:**  
`MASTER_SYNC_ACTIVE` is now treated as temporary pending/retry state.

**Result:**  
The record can remain local and upload after master lock is released.

---

### 3.2 [Sync] Missing client-side master-sync guard include

**Problem:**  
`master-sync-guard.js` existed but was not linked in `index.html`.

**Fix:**  
Add the script include in correct runtime order.

**Result:**  
PWA overlay/guard now runs.

---

### 3.3 [Sync] Server-side write race during VBA master sync

**Problem:**  
Client-side checks alone are not sufficient. A PWA write could reach GAS during VBA sync.

**Fix:**  
GAS now checks master-sync lock before write dispatch.

**Result:**  
Server-side enforcement prevents Google mutation during active VBA master sync.

---

### 3.4 [Geo] Parcel polygon overwrite risk

**Problem:**  
PWA-created polygon in Google could be overwritten by stale VBA Stammdaten export.

**Fix:**  
Pull parcel geo from Google into master before Stammdaten export.

**Result:**  
PWA/GAS polygon edits are protected in the full-cycle flow.

---

### 3.5 [UX] Operator could think Excel is frozen during sync

**Problem:**  
Long sync had insufficient progress feedback.

**Fix:**  
Added live sync log/progress messages in `frmOtkupAPP`.

**Result:**  
Operator sees ongoing work.

---

### 3.6 [Runtime] Duplicate/incorrect runtime script loading risk

**Problem:**  
Runtime modules can behave incorrectly if critical files are duplicated or missing.

**Fix:**  
`master-sync-guard.js` is included once and `sync-engine.js` should not be duplicated.

**Result:**  
More deterministic PWA runtime.

---

## 4. Deprecated

### 4.1 Direct UI use of individual sync functions

Deprecated as operator flow:

```vb
ImportOtkupFromPWA
ImportZbirneFromPWA
SyncStammdatenToGoogle
ExportKarticeToGoogle
ExportMgmtReports
```

Replacement:

```vb
SyncPWAFullCycle_ForButton
```

Individual functions remain valid as internal/test/admin functions.

---

### 4.2 Treating `MASTER_SYNC_ACTIVE` as row sync error

Deprecated behavior:

```text
MASTER_SYNC_ACTIVE -> sync error
```

Replacement:

```text
MASTER_SYNC_ACTIVE -> pending/retry
```

---

### 4.3 Business sync from workbook events

Deprecated pattern:

```text
Workbook_Open directly runs business sync
Workbook_BeforeClose directly runs business sync
```

Replacement:

```text
StartApp starts scheduler
ShutdownApp stops scheduler
Application.OnTime controls background ticks
```

---

## 5. Operational Notes

### 5.1 Deployment order

Recommended order:

1. Deploy GAS.
2. Deploy PWA with `master-sync-guard.js` linked.
3. Bump service worker cache.
4. Import/export updated VBA modules.
5. Compile VBA.
6. Run full-cycle smoke.

### 5.2 Smoke tests required

Minimum smoke:

1. Start VBA full cycle.
2. Verify `SyncControl` lock becomes `YES`.
3. During lock, try PWA sync.
4. Verify no permanent sync error is created for lock condition.
5. Verify local PWA capture remains possible if soft-lock policy is kept.
6. Finish VBA sync.
7. Verify lock becomes `NO`.
8. Run PWA sync again.
9. Verify pending data uploads.
10. Run VBA full cycle again.
11. Verify uploaded data imports.
12. Edit polygon in PWA/HTML.
13. Run full cycle.
14. Verify polygon is not overwritten.

### 5.3 Manual troubleshooting

If PWA appears blocked, check:

```text
Stammdaten / SyncControl
MASTER_SYNC_LOCK
MASTER_SYNC_UPDATED_AT
MASTER_SYNC_MESSAGE
```

If timestamp is old, stale-lock logic should eventually allow writes, but manual inspection is still useful.

---

## 6. Known Issues / Remaining Risks

### KI-6.18-01 — Old PWA cache on field devices

**Risk:**  
Some devices may still run old service worker cache without `master-sync-guard.js`.

**Mitigation:**  

- bump cache
- force reload/update
- verify runtime file loaded

---

### KI-6.18-02 — Real polygon smoke still required

**Risk:**  
Geo protection is architecturally correct but must be validated with actual polygon edit.

**Mitigation:**  

- edit polygon in PWA/HTML
- confirm Google value
- run VBA full cycle
- confirm master received value
- confirm export did not wipe value

---

### KI-6.18-03 — Too broad GAS write blocking

**Risk:**  
A write action might be blocked even if it is harmless.

**Mitigation:**  
Accept for launch safety. Refine action list after field observation.

---

### KI-6.18-04 — Too strict local hard-block if enabled everywhere

**Risk:**  
If `withSubmitLock` blocks local save everywhere, operator productivity drops.

**Mitigation:**  
Keep soft-lock for ordinary otkup capture. Use hard-block only for workflows that require immediate server write.

---

### KI-6.18-05 — VBA module export drift

**Risk:**  
Excel VBA state may differ from repository/exported files.

**Mitigation:**  
After final test, export modules from Excel and archive/tag them.

---

## 7. Acceptance Criteria

v6.18 can be accepted when:

- VBA compiles cleanly.
- Full-cycle button runs.
- Lock is written and released.
- GAS blocks writes during active lock.
- PWA guard overlay appears.
- PWA does not convert lock condition into permanent sync error.
- PWA pending records upload after lock release.
- VBA imports those records in next cycle.
- Polygon data survives Stammdaten export.
- Service worker cache update is confirmed.

---

## 8. Launch Decision

Recommended status:

```text
Launch Candidate
```

Reason:

- main VBA/PWA/GAS race is controlled
- field work can continue
- no data loss expected from sync lock
- operator visibility improved
- geo overwrite risk controlled

Final launch should wait for the acceptance smoke listed above.

---

## v6.19 — 2026-05-12

The following content is embedded in the v6.20 changelog so the latest CL remains self-contained.

# AgriX / OtkupApp Changelog — v6.19

**Version:** v6.19  
**Date:** 2026-05-12  
**Status:** Launch-candidate geo/Stammdaten UX hardening  
**Architecture reference:** `AGRIX_ARCHITECTURE_REFERENCE_v6_19.md`  
**Scope:** Changes after v6.18.

---

## 0. Summary

v6.19 is a focused VBA hardening release for the master-data parcel geo workflow and operator shell KPI robustness.

Main outcomes:

- `frmStammdaten` parcel geo controls are visible only for an actually selected parcel.
- Geo actions no longer use blocking `MsgBox` prompts during normal geo flow.
- A non-blocking geo status label is now the canonical feedback surface for parcel geo actions.
- Geo save/clear preserve the selected parcel after list refresh.
- Polygon editor launch now first syncs the selected local parcel to Google, then opens the editor.
- `modGeoParcele` is promoted to a domain/service-level geo module, aligned with the transactional pattern used by higher-criticality business modules.
- `SaveParcelGeoPoint` and `ClearParcelGeo` now use transaction-backed `_TX` implementations with rollback support.
- `SyncSelectedParcelaToGoogle` provides the canonical service entry point for “selected parcel → Google → polygon editor”.
- `frmOtkupAPP` KPI helpers were hardened against type mismatch on dirty date/kg data.

v6.19 does not replace the v6.18 master-sync contract. It builds on it by making the parcel geo/editor workflow safer and more operator-friendly.

---

## 1. Added

### 1.1 [VBA] `modGeoParcele.SaveParcelGeoPoint_TX`

Added transaction-backed geo point save function:

```vb
Public Function SaveParcelGeoPoint_TX(ByVal rowIndex As Long, _
                                      ByVal nCoord As Double, _
                                      ByVal eCoord As Double) As Boolean
```

The function:

- validates `rowIndex`
- validates positive N/E coordinates
- converts UTM zone 34 coordinates to `Lat`/`Lng`
- starts `clsTransaction`
- snapshots `tblParcele`
- writes geo point fields through `RequireUpdateCell`
- commits on success
- rolls back on failure

Updated fields:

```text
N_Coord
E_Coord
Lat
Lng
GeoStatus
GeoSource
MeteoEnabled
DatumGeoUnosa
DatumAzuriranja
```

Canonical values on save:

```text
GeoStatus = point
GeoSource = selenium
MeteoEnabled = Da
```

**Reason:**  
The old implementation used multiple independent `UpdateCell` calls. If one write failed mid-sequence, `tblParcele` could contain a partially updated geo state.

**Impact:**  
Geo point save now follows the same transactional discipline expected from domain/business modules.

**Migration:**  
No data migration.

---

### 1.2 [VBA] `modGeoParcele.ClearParcelGeo_TX`

Added transaction-backed geo clear function:

```vb
Public Function ClearParcelGeo_TX(ByVal rowIndex As Long) As Boolean
```

The function:

- validates `rowIndex`
- starts `clsTransaction`
- snapshots `tblParcele`
- clears point fields through `RequireUpdateCell`
- updates geo status fields
- commits on success
- rolls back on failure

Canonical values on clear:

```text
GeoStatus = none
GeoSource = empty
MeteoEnabled = Ne
DatumAzuriranja = Now
```

**Reason:**  
Clearing geo data is also a multi-field mutation and must not leave partial state.

**Impact:**  
Geo clear becomes rollback-safe.

**Migration:**  
No data migration.

---

### 1.3 [VBA] Compatibility wrappers for geo save/clear

Existing public procedure names remain available:

```vb
Public Sub SaveParcelGeoPoint(ByVal rowIndex As Long, ByVal nCoord As Double, ByVal eCoord As Double)
Public Sub ClearParcelGeo(ByVal rowIndex As Long)
```

They now call their `_TX` equivalents and raise an error if the transaction function returns `False`.

**Reason:**  
Existing form code can keep calling stable public names while the implementation becomes transaction-safe.

**Impact:**  
No form-level rename required.

**Migration:**  
No migration.

---

### 1.4 [VBA] Selected parcel Google sync service

Added canonical geo service entry point:

```vb
Public Function SyncSelectedParcelaToGoogle(ByVal parcelaID As String) As Boolean
```

The function:

1. validates `ParcelaID`
2. confirms the parcel exists locally in `tblParcele`
3. calls the existing canonical parcel export wrapper:

```vb
SyncParceleToGoogle_Core(False)
```

4. verifies that the parcel exists in Google `Stammdaten / Parcele` after sync
5. returns `True` only when the editor can safely read the selected parcel online

**Reason:**  
The polygon editor reads the Google/online layer, not the Excel workbook. If the operator just saved coordinates locally, the editor must not open stale or empty Google data.

**Impact:**  
The editor launch flow becomes:

```text
selected tblParcele row
↓
SyncSelectedParcelaToGoogle
↓
Stammdaten / Parcele contains selected ParcelaID
↓
open parcel-draw.html?parcelaId=...
```

**Migration:**  
Requires `SyncParceleToGoogle_Core(False)` to exist in `modStammdatenSync` as a public wrapper around the existing canonical `ExportParcele` logic.

---

### 1.5 [VBA] Local/Google parcel existence checks

Added helper checks in `modGeoParcele`:

```vb
Private Function LocalParcelaExists(ByVal parcelaID As String) As Boolean
Private Function GoogleParcelaExists(ByVal parcelaID As String) As Boolean
Private Function FindHeaderColumnInArray(ByVal data As Variant, ByVal headerName As String) As Long
Private Function SafeArrayText(ByVal data As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
```

**Reason:**  
`SyncSelectedParcelaToGoogle` must fail closed if the selected parcel cannot be proven locally or in Google after sync.

**Impact:**  
Polygon editor is not opened against missing/stale Google data.

---

### 1.6 [VBA] `frmStammdaten` geo status surface

Added non-blocking geo status helpers:

```vb
Private Sub SetGeoStatus(ByVal message As String, Optional ByVal isError As Boolean = False)
Private Sub ClearGeoStatus()
Private Sub ResetGeoClearConfirm()
Private Function HasSelectedParcelaForGeo() As Boolean
```

Expected UI control:

```text
lblGeoStatus
```

**Reason:**  
Geo button flow should not be interrupted by `MsgBox` for normal recoverable states.

**Impact:**  
Operators see inline geo status/errors without modal interruption.

---

### 1.7 [VBA] Parcel selection preservation helpers

Added helpers to preserve selected parcel after list refresh:

```vb
Private Function GetSelectedParcelaID() As String
Private Sub ReselectParcelaInList(ByVal parcelaID As String)
```

**Reason:**  
`LoadList` rebuilds the ListBox and row map. After save/clear geo, the selected parcel must remain selected, otherwise geo controls can enter a confusing state.

**Impact:**  
After saving or clearing geo data, the same parcel remains selected and geo actions remain coherent.

---

## 2. Changed

### 2.1 [VBA] `frmStammdaten` geo control visibility

Changed geo controls to be visible only when all conditions are true:

```vb
Me.Tag = "Parcele"
m_SelectedRow > 0
lstData.ListIndex >= 0
```

Canonical helper:

```vb
Private Sub UpdateGeoControlsVisibility()
    SetGeoControlsVisible (Me.Tag = "Parcele" And _
                           m_SelectedRow > 0 And _
                           lstData.ListIndex >= 0)
End Sub
```

**Reason:**  
Geo buttons were visible across modes or when no concrete parcel was selected.

**Impact:**  
Geo actions are only available for a real selected parcel.

---

### 2.2 [VBA] `frmStammdaten` geo buttons no longer show normal-flow MsgBox prompts

Changed geo button UX from modal messages to inline status messages:

- `btnGeoOpen_Click`
- `btnGeoSave_Click`
- `btnGeoClear_Click`
- `btnPasteCoords_Click`
- `btnOpenMap_Click`
- `btnOpenPolygonEditor_Click`

**Reason:**  
Geo workflow is an operator micro-flow. Repeated modal popups slow down the workflow and create poor UX.

**Impact:**  
Normal geo issues such as empty clipboard, missing selection, invalid coordinates, no Lat/Lng, and sync failure are shown in `lblGeoStatus`.

---

### 2.3 [VBA] Geo clear confirmation changed from modal prompt to second-click confirmation

Changed delete/clear confirmation:

Old behavior:

```text
MsgBox Yes/No confirmation
```

New behavior:

```text
First click: sets button caption to "Potvrdi brisanje" and shows status warning.
Second click: clears geo data.
```

State flag:

```vb
Private mGeoClearConfirmPending As Boolean
```

**Reason:**  
Keeps geo flow non-modal while still protecting destructive action.

**Impact:**  
No blocking prompt for clear; accidental click still protected.

---

### 2.4 [VBA] Polygon editor launch flow

Changed `btnOpenPolygonEditor_Click` to require selected-parcel sync before opening the editor:

```vb
If Not SyncSelectedParcelaToGoogle(parcelaID) Then
    SetGeoStatus "Parcela nije sinhronizovana. Editor nije otvoren.", True
    Exit Sub
End If

If OpenParcelPolygonEditor(parcelaID) Then
    SetGeoStatus "Polygon editor otvoren.", False
End If
```

**Reason:**  
The editor must open the coordinates/row just saved in `tblParcele`, not stale/empty data currently in Google Sheets.

**Impact:**  
Operator action now has deterministic data flow:

```text
local selected parcel → Google Parcele tab → polygon editor
```

---

### 2.5 [VBA] `OpenGoogleMaps` / `OpenParcelPolygonEditor` return Boolean

Changed helper procedures from fire-and-forget `Sub` style to success-aware `Function` style:

```vb
Public Function OpenGoogleMaps(ByVal lat As Double, ByVal lng As Double) As Boolean
Public Function OpenParcelPolygonEditor(ByVal parcelaID As String) As Boolean
```

**Reason:**  
Callers need to show inline status without raising modal errors.

**Impact:**  
Geo button handlers can display success/failure in `lblGeoStatus`.

---

### 2.6 [VBA] `btnPasteCoords_Click` now requires selected parcel

Changed paste-coordinates flow to require `Parcele` mode and selected parcel:

```vb
If Me.Tag <> "Parcele" Then Exit Sub
If Not HasSelectedParcelaForGeo() Then Exit Sub
```

**Reason:**  
Paste coordinates is a parcel-specific action and should not run without a selected parcel.

**Impact:**  
More consistent geo guard behavior across all geo buttons.

---

### 2.7 [VBA] `frmOtkupAPP` KPI helpers hardened against dirty dates/numbers

Hardened KPI/date helper behavior for functions such as:

```vb
SumOtkupKgForDate
CountDocsForDate
RefreshSidebarKpi
```

Recommended safe helper pattern:

```vb
SafeDateKey
SafeKpiDouble
```

**Reason:**  
Dashboard KPI refresh produced runtime error `13 | Type mismatch` when a date/kg field contained an unexpected value, blank, error, or non-date text.

**Impact:**  
KPI widgets skip invalid rows and return safe zero/summary values instead of logging repeat runtime errors.

---

## 3. Fixed

### 3.1 [Geo UX] Geo buttons visible outside valid parcel selection

**Problem:**  
Geo buttons could remain visible in contexts where no parcel was selected.

**Fix:**  
`UpdateGeoControlsVisibility` now requires both `m_SelectedRow > 0` and `lstData.ListIndex >= 0` in `Parcele` mode.

**Result:**  
Geo UI is hidden until a concrete parcel is selected.

---

### 3.2 [Geo UX] Modal popup noise in geo workflow

**Problem:**  
Geo button flow used `MsgBox` for normal states such as missing data, invalid coordinates, empty clipboard, sync failure, and success confirmation.

**Fix:**  
Geo flow now uses `SetGeoStatus` / `lblGeoStatus` for inline feedback.

**Result:**  
Operator workflow is faster and less disruptive.

---

### 3.3 [Geo UX] Selected parcel lost after geo save/clear refresh

**Problem:**  
`LoadList` refreshes the ListBox and row map after geo save/clear. The same parcel could lose active selection.

**Fix:**  
Save/clear now captures selected `ParcelaID`, reloads the list, and reselects the same parcel.

**Result:**  
Selected parcel remains stable after geo mutation.

---

### 3.4 [Geo/Data] Polygon editor could open stale or empty Google parcel data

**Problem:**  
After saving coordinates locally in `tblParcele`, opening the polygon editor could still show stale data from Google `Stammdaten / Parcele`, or blank state if the row had not been exported.

**Fix:**  
`btnOpenPolygonEditor_Click` now calls `SyncSelectedParcelaToGoogle(parcelaID)` before opening the editor.

**Result:**  
The editor reads the current selected parcel state from Google.

---

### 3.5 [Geo/Data] Multi-field geo save/clear could be partial

**Problem:**  
Old geo save/clear used repeated `UpdateCell` calls without transaction rollback.

**Fix:**  
`SaveParcelGeoPoint_TX` and `ClearParcelGeo_TX` use `clsTransaction`, `AddTableSnapshot`, `RequireUpdateCell`, `CommitTx`, and rollback on failure.

**Result:**  
Geo point save/clear are atomic at the `tblParcele` table level.

---

### 3.6 [Dashboard] `frmOtkupAPP` KPI type mismatch

**Problem:**  
KPI refresh could log repeated `Type mismatch` errors from date/count/sum helper functions.

**Fix:**  
KPI helpers were made tolerant of invalid date/numeric values.

**Result:**  
Bad/blank/error cells no longer break sidebar KPI refresh.

---

## 4. Deprecated

### 4.1 Direct full Stammdaten export from polygon editor button

Deprecated operator flow:

```vb
SyncStammdatenToGoogle_Core(False)
```

from `btnOpenPolygonEditor_Click`.

Replacement:

```vb
SyncSelectedParcelaToGoogle(parcelaID)
```

**Reason:**  
The editor button needs a selected-parcel sync service, not a broad all-Stammdaten operator export.

---

### 4.2 Geo button MsgBox prompts for normal recoverable states

Deprecated pattern inside geo button handlers:

```vb
MsgBox "...", vbExclamation, APP_NAME
```

Replacement:

```vb
SetGeoStatus "...", True
```

**Reason:**  
Geo flow should be non-modal except for wider form-level critical errors.

---

### 4.3 Raw multi-field `UpdateCell` geo mutations

Deprecated pattern:

```vb
UpdateCell TBL_PARCELE, rowIndex, "N_Coord", ...
UpdateCell TBL_PARCELE, rowIndex, "E_Coord", ...
...
```

Replacement:

```vb
SaveParcelGeoPoint_TX
ClearParcelGeo_TX
```

**Reason:**  
Multi-field parcel geo mutations need rollback discipline.

---

## 5. Operational Notes

### 5.1 Deployment order

Recommended order:

1. Import updated `modGeoParcele`.
2. Confirm `modStammdatenSync.SyncParceleToGoogle_Core(False)` exists and compiles.
3. Import updated `frmStammdaten`.
4. Confirm `lblGeoStatus` exists on the form.
5. Compile all VBA.
6. Run geo smoke tests.
7. Export VBA modules and tag/archive v6.19.

---

### 5.2 Required form controls

`frmStammdaten` expects these geo controls to exist:

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

`lblGeoStatus` is the canonical non-modal status surface.

---

### 5.3 Geo smoke tests

Minimum smoke:

1. Open `frmStammdaten` with `Tag = "Parcele"`.
2. Verify geo controls are hidden before parcel selection.
3. Select a parcel.
4. Verify geo controls become visible.
5. Click `Otvori GeoData` / GeoSrbija and verify search text is copied.
6. Paste N/E coordinates.
7. Save geo point.
8. Verify same parcel remains selected.
9. Verify `Lat`/`Lng`, `GeoStatus`, `GeoSource`, `MeteoEnabled`, dates update in `tblParcele`.
10. Open Google Maps.
11. Open polygon editor.
12. Verify selected parcel exists in Google `Stammdaten / Parcele` before editor opens.
13. Clear geo with two-click confirmation.
14. Verify same parcel remains selected and point fields are cleared.

---

### 5.4 KPI smoke tests

Minimum smoke:

1. Add or simulate dirty date/kg values in relevant tables.
2. Open `frmOtkupAPP`.
3. Refresh sidebar KPI.
4. Verify no `13 | Type mismatch` log entries from `SumOtkupKgForDate` or `CountDocsForDate`.
5. Verify KPI shows safe zero or valid aggregate.

---

## 6. Known Issues / Remaining Risks

### KI-6.19-01 — Selected parcel sync currently depends on `SyncParceleToGoogle_Core`

**Risk:**  
`SyncSelectedParcelaToGoogle` depends on `SyncParceleToGoogle_Core(False)` being available and correctly implemented in `modStammdatenSync`.

**Mitigation:**  
Compile all VBA and run selected-parcel sync smoke.

---

### KI-6.19-02 — Selected parcel sync may still use full Parcele-tab export internally

**Risk:**  
The selected parcel service may still delegate to full `ExportParcele` through the existing canonical export path. This is intentionally chosen to avoid duplicate export mapping, but it is broader than a true one-row upsert.

**Mitigation:**  
Accepted for v6.19 because it reuses canonical mapping and was reported fast enough. Consider true row upsert only if field data volume makes full Parcele export too slow.

---

### KI-6.19-03 — Polygon overwrite safety still governed by v6.18 full-cycle rule

**Risk:**  
If local `PolygonGeoJSON` is empty and selected-parcel sync performs a full export, existing Google polygon data can still be overwritten unless the canonical export/pull discipline is respected.

**Mitigation:**  
Keep v6.18 rule: full-cycle sync must import Google parcel geo/polygon into master before outbound Stammdaten export. For editor pre-open selected sync, confirm expected behavior with real polygon smoke.

---

### KI-6.19-04 — `lblGeoStatus` missing from designer

**Risk:**  
If the form designer does not contain `lblGeoStatus`, inline status will not show. The code uses defensive `Controls("lblGeoStatus")`, but UX will be silent.

**Mitigation:**  
Verify form designer contains `lblGeoStatus`.

---

## 7. Acceptance Criteria

v6.19 can be accepted when:

- VBA compiles cleanly.
- `frmStammdaten` opens in all supported modes.
- Geo controls are hidden in non-`Parcele` modes.
- Geo controls are hidden in `Parcele` mode until a row is selected.
- Geo flow uses `lblGeoStatus`, not normal-flow `MsgBox` prompts.
- Save geo point writes all required fields transactionally.
- Clear geo writes all required fields transactionally.
- Save/clear rollback correctly on forced error.
- Save/clear preserve selected parcel after `LoadList`.
- `SyncSelectedParcelaToGoogle` returns `True` only when selected parcel exists in Google after sync.
- Polygon editor opens only after successful selected-parcel sync.
- `frmOtkupAPP` KPI refresh no longer logs type mismatch errors.

---

## 8. Launch Decision

Recommended status:

```text
Launch Candidate
```

Reason:

- parcel geo/editor operator flow is now coherent
- geo mutations are transaction-backed
- modal geo UX noise is removed
- selected parcel remains stable through list refresh
- editor no longer opens against stale/empty Google data
- KPI shell robustness improved

Final production launch should still wait for the geo smoke, polygon smoke, KPI smoke, and full v6.18 master-sync smoke.

---

---

## v6.18 — 2026-05-11

The following content is embedded in the v6.20 changelog so the latest CL remains self-contained.

# AgriX / OtkupApp Changelog — v6.18

**Version:** v6.18  
**Date:** 2026-05-11  
**Status:** Launch-candidate hardening  
**Architecture reference:** `AGRIX_ARCHITECTURE_REFERENCE_v6_18.md`  
**Scope:** Changes after v6.17.

---

## 0. Summary

v6.18 delivers a coordinated VBA/PWA/GAS master-sync guard and closes the major launch blocker around PWA writes during VBA master sync.

Main outcomes:

- VBA now owns one full-cycle PWA/Google sync entrypoint.
- VBA publishes a lock into `Stammdaten / SyncControl`.
- GAS exposes the lock through `getMasterSyncState`.
- GAS blocks write actions while master sync is active.
- PWA shows an operator-visible sync overlay.
- PWA no longer needs to treat `MASTER_SYNC_ACTIVE` as a true sync error.
- Local PWA work can continue in offline-first pending mode.
- Parcel geo/polygon data is pulled into master before Stammdaten export, preventing overwrites.
- Operator sync button has live progress feedback.

This version is a launch-candidate sync-hardening release.

---

## 1. Added

### 1.1 [VBA] Full PWA/Google sync orchestrator

Added canonical orchestration layer for full PWA/Google cycle.

Canonical procedures:

```vb
Public Sub SyncPWAFullCycle()
Public Function SyncPWAFullCycle_ForButton() As Boolean
Private Function SyncPWAFullCycle_Core(ByVal showMessages As Boolean) As Boolean
```

The orchestrator owns the ordered flow:

1. Google config/auth check
2. lock PWA writes
3. pull parcel geo from Google
4. import OTK/PWA data
5. auto-create/link Otpremnice
6. import VOZ/Zbirna data
7. export Stammdaten
8. export Kartice
9. export MgmtReports
10. unlock PWA writes
11. monitor result

**Reason:**  
The operator needs a single safe master caller instead of multiple separate sync buttons/procedures.

**Impact:**  
Manual sync should go through `SyncPWAFullCycle_ForButton`.

**Migration:**  
No data migration.

---

### 1.2 [VBA] PWA master-sync lock writer

Added lock writer that updates:

```text
Stammdaten / SyncControl
```

Canonical keys:

```text
MASTER_SYNC_LOCK
MASTER_SYNC_UPDATED_AT
MASTER_SYNC_MESSAGE
MASTER_SYNC_OWNER
```

Expected state during sync:

```text
MASTER_SYNC_LOCK = YES
MASTER_SYNC_OWNER = VBA
```

Expected state after sync:

```text
MASTER_SYNC_LOCK = NO
```

**Reason:**  
PWA and GAS need a shared signal that VBA master sync is active.

**Impact:**  
GAS and PWA can coordinate around VBA sync.

**Migration:**  
No data migration. `SyncControl` tab is created/updated as needed.

---

### 1.3 [VBA] Parcel geo pull before Stammdaten export

Added/imported function:

```vb
Public Function ImportParcelGeoFromGoogleToMaster() As Boolean
```

This pulls geo/polygon data from Google Stammdaten `Parcele` into `tblParcele`.

**Reason:**  
Polygons created in PWA/HTML/GAS flow could be overwritten by stale VBA Stammdaten export.

**Impact:**  
Full-cycle sync must run geo pull before outbound Stammdaten.

**Migration:**  
No structural migration, but final smoke required with real polygon edits.

---

### 1.4 [VBA] Core functions for orchestration

Added/standardized core functions:

```vb
Public Function SyncStammdatenToGoogle_Core(ByVal showMessages As Boolean) As Boolean
Public Function ExportKarticeToGoogle_Core(ByVal showMessages As Boolean) As Boolean
Public Function ExportMgmtReports_Core(ByVal showMessages As Boolean) As Boolean
Public Function ImportOtkupFromPWA_Core(ByVal showMessages As Boolean) As Boolean
Public Function ImportZbirneFromPWA_Core(ByVal showMessages As Boolean) As Boolean
```

Public UI subs remain thin wrappers.

**Reason:**  
The orchestrator must call sync/export/import logic without unnecessary UI popups.

**Impact:**  
Cleaner separation between UI and automation.

**Migration:**  
No data migration.

---

### 1.5 [VBA] Live sync progress log

Added progress/log mechanism for the PWA full-cycle sync.

Canonical orchestrator helper:

```vb
SyncProgress "message"
```

Expected form methods:

```vb
BeginPWASyncLog
AppendPWASyncLog
EndPWASyncLog
```

**Reason:**  
Long Google/VBA sync operations can look like Excel is frozen.

**Impact:**  
Operator sees what the system is doing.

**Migration:**  
No data migration.

---

### 1.6 [VBA] Scheduled auto-sync

Added scheduled sync support using `Application.OnTime`.

Canonical functions:

```vb
StartScheduledSync
StopScheduledSync
ScheduledSyncTick
```

Config key:

```text
SYNC_AUTO_INTERVAL_MIN
```

Rules:

```text
empty/0 = disabled
< 15 = disabled/rejected
>= 15 = enabled
```

**Reason:**  
Auto sync should run from timers after application start, not directly from `Workbook_Open`/`BeforeClose`.

**Impact:**  
App startup/shutdown stays responsive.

**Migration:**  
No data migration.

---

### 1.7 [GAS] Master sync state endpoint

Added public read action:

```js
getMasterSyncState
```

Returns:

```js
success
locked
stale
updatedAt
ageMin
message
```

**Reason:**  
PWA needs a safe read-only endpoint to know whether VBA sync is active.

**Impact:**  
PWA can show overlay and avoid write attempts during lock.

**Migration:**  
Deploy GAS update.

---

### 1.8 [GAS] Master sync write blocker

Added server-side guard:

```js
blockWriteIfMasterSyncActive(action)
isMasterSyncWriteAction(action)
```

Blocked write actions include:

```text
sync
syncAgromere
syncZbirna
syncTretman
syncOprema
syncTrosak
saveParcelPolygon
uploadPdf
saveWarRoomDemand
removeWarRoomDemand
updateDemandPrimljeno
updateKamionStatus
saveDispecer
updateDispecer
removeDispecer
saveIzdavanje
saveFiskalni
saveFiskalniMapiranje
createArtikal
```

Blocked response:

```json
{
  "success": false,
  "locked": true,
  "code": "MASTER_SYNC_ACTIVE"
}
```

**Reason:**  
Client-side guard is useful, but server-side enforcement is mandatory.

**Impact:**  
No PWA write can mutate Google data during active VBA master sync.

**Migration:**  
Deploy GAS update.

---

### 1.9 [PWA] `master-sync-guard.js`

Added runtime guard:

```text
src/js/utils/master-sync-guard.js
```

Exposes:

```js
getMasterSyncStateSafe(force)
ensureMasterSyncNotActive(context, options)
startMasterSyncGuardPolling()
```

Includes:

- state polling
- cache
- overlay
- manual re-check
- offline/unknown handling

**Reason:**  
PWA UI needs to know when master sync is active.

**Impact:**  
PWA can inform user and avoid unnecessary write attempts.

**Migration:**  
PWA deploy + service worker cache bump.

---

## 2. Changed

### 2.1 [PWA] `index.html` runtime order

Changed runtime include order to ensure `master-sync-guard.js` is loaded.

Required conceptual order:

```html
<script src="./src/js/services/api.js"></script>
<script src="./src/js/utils/master-sync-guard.js"></script>
<script src="./src/js/utils/async.js"></script>
<script src="./src/js/utils/sync-engine.js"></script>
```

**Reason:**  
The guard must exist before `withSubmitLock`/`syncStore` uses it.

**Impact:**  
Client guard now actually runs.

---

### 2.2 [PWA] `sync-engine.js` handling for `MASTER_SYNC_ACTIVE`

Changed behavior when GAS returns:

```js
code === 'MASTER_SYNC_ACTIVE'
```

Old behavior:

```text
treated as sync error
record could appear in sync errors
```

New behavior:

```text
record stays pending/retryable
lastSyncError is cleared
lastServerStatus = master-sync-active
failed count is not increased for this condition
```

Recommended local state:

```js
record.syncStatus = 'pending';
record.lastServerStatus = 'master-sync-active';
record.lastSyncError = '';
```

**Reason:**  
Master sync lock is temporary coordination state, not data error.

**Impact:**  
Operators do not see false sync errors for records created during VBA sync.

---

### 2.3 [PWA] Soft-lock policy

Changed recommended behavior from hard blocking all local work to soft-locking server writes.

New policy:

```text
Allow local IndexedDB capture.
Block/skip GAS upload while master sync is active.
Retry later.
```

**Reason:**  
Field work should not stop unnecessarily.

**Impact:**  
PWA can keep working during VBA sync; records upload later.

---

### 2.4 [PWA] Service worker cache

Changed service worker cache to include the new guard file and require version bump.

Required cached file:

```text
src/js/utils/master-sync-guard.js
```

**Reason:**  
Field devices must load the new guard.

**Impact:**  
Users may need refresh/update after deployment.

---

### 2.5 [VBA] `frmOtkupAPP` sync button

Changed sync button flow to call:

```vb
SyncPWAFullCycle_ForButton
```

instead of individual import/export routines.

**Reason:**  
One operator action must execute the full safe sync sequence.

**Impact:**  
Better operator flow and less chance of incomplete sync.

---

### 2.6 [VBA] Startup/shutdown sync behavior

Changed auto-sync logic so it belongs to app lifecycle:

```vb
StartApp -> StartScheduledSync
ShutdownApp -> StopScheduledSync
```

Not direct workbook-open/before-close business sync.

**Reason:**  
Avoid blocking workbook startup/shutdown.

**Impact:**  
More stable runtime.

---

### 2.7 [VBA] Stammdaten export empty dataset behavior

Changed multiple export functions to prefer header-only output for legitimate empty datasets.

Canonical helper:

```vb
WriteHeaderOnly(sheetID, tabName, headers...)
```

**Reason:**  
Empty reports/tables should not automatically mean failed sync.

**Impact:**  
Fewer false partial failures.

---

### 2.8 [VBA] Active flag normalization

Changed PWA exports to use normalized active-check logic.

Inactive values:

```text
NE
NO
FALSE
0
NEAKTIVAN
INACTIVE
```

All other values are active.

**Reason:**  
Tables use mixed active markers.

**Impact:**  
More consistent PWA Stammdaten output.

---

### 2.9 [VBA] Config export filtering

Changed PWA config export safety to exclude internal sync keys.

Excluded prefixes:

```text
GOOGLE_
SEF_
SYNC_
```

**Reason:**  
Runtime sync config and credentials must not be exposed as PWA business config.

**Impact:**  
Safer PWA Config export.

---

## 3. Fixed

### 3.1 [Sync] PWA record created during VBA sync looked like sync error

**Problem:**  
When PWA saved a record while VBA sync was active, the record could be saved locally but fail upload due to lock/server state and appear as sync error.

**Fix:**  
`MASTER_SYNC_ACTIVE` is now treated as temporary pending/retry state.

**Result:**  
The record can remain local and upload after master lock is released.

---

### 3.2 [Sync] Missing client-side master-sync guard include

**Problem:**  
`master-sync-guard.js` existed but was not linked in `index.html`.

**Fix:**  
Add the script include in correct runtime order.

**Result:**  
PWA overlay/guard now runs.

---

### 3.3 [Sync] Server-side write race during VBA master sync

**Problem:**  
Client-side checks alone are not sufficient. A PWA write could reach GAS during VBA sync.

**Fix:**  
GAS now checks master-sync lock before write dispatch.

**Result:**  
Server-side enforcement prevents Google mutation during active VBA master sync.

---

### 3.4 [Geo] Parcel polygon overwrite risk

**Problem:**  
PWA-created polygon in Google could be overwritten by stale VBA Stammdaten export.

**Fix:**  
Pull parcel geo from Google into master before Stammdaten export.

**Result:**  
PWA/GAS polygon edits are protected in the full-cycle flow.

---

### 3.5 [UX] Operator could think Excel is frozen during sync

**Problem:**  
Long sync had insufficient progress feedback.

**Fix:**  
Added live sync log/progress messages in `frmOtkupAPP`.

**Result:**  
Operator sees ongoing work.

---

### 3.6 [Runtime] Duplicate/incorrect runtime script loading risk

**Problem:**  
Runtime modules can behave incorrectly if critical files are duplicated or missing.

**Fix:**  
`master-sync-guard.js` is included once and `sync-engine.js` should not be duplicated.

**Result:**  
More deterministic PWA runtime.

---

## 4. Deprecated

### 4.1 Direct UI use of individual sync functions

Deprecated as operator flow:

```vb
ImportOtkupFromPWA
ImportZbirneFromPWA
SyncStammdatenToGoogle
ExportKarticeToGoogle
ExportMgmtReports
```

Replacement:

```vb
SyncPWAFullCycle_ForButton
```

Individual functions remain valid as internal/test/admin functions.

---

### 4.2 Treating `MASTER_SYNC_ACTIVE` as row sync error

Deprecated behavior:

```text
MASTER_SYNC_ACTIVE -> sync error
```

Replacement:

```text
MASTER_SYNC_ACTIVE -> pending/retry
```

---

### 4.3 Business sync from workbook events

Deprecated pattern:

```text
Workbook_Open directly runs business sync
Workbook_BeforeClose directly runs business sync
```

Replacement:

```text
StartApp starts scheduler
ShutdownApp stops scheduler
Application.OnTime controls background ticks
```

---

## 5. Operational Notes

### 5.1 Deployment order

Recommended order:

1. Deploy GAS.
2. Deploy PWA with `master-sync-guard.js` linked.
3. Bump service worker cache.
4. Import/export updated VBA modules.
5. Compile VBA.
6. Run full-cycle smoke.

### 5.2 Smoke tests required

Minimum smoke:

1. Start VBA full cycle.
2. Verify `SyncControl` lock becomes `YES`.
3. During lock, try PWA sync.
4. Verify no permanent sync error is created for lock condition.
5. Verify local PWA capture remains possible if soft-lock policy is kept.
6. Finish VBA sync.
7. Verify lock becomes `NO`.
8. Run PWA sync again.
9. Verify pending data uploads.
10. Run VBA full cycle again.
11. Verify uploaded data imports.
12. Edit polygon in PWA/HTML.
13. Run full cycle.
14. Verify polygon is not overwritten.

### 5.3 Manual troubleshooting

If PWA appears blocked, check:

```text
Stammdaten / SyncControl
MASTER_SYNC_LOCK
MASTER_SYNC_UPDATED_AT
MASTER_SYNC_MESSAGE
```

If timestamp is old, stale-lock logic should eventually allow writes, but manual inspection is still useful.

---

## 6. Known Issues / Remaining Risks

### KI-6.18-01 — Old PWA cache on field devices

**Risk:**  
Some devices may still run old service worker cache without `master-sync-guard.js`.

**Mitigation:**  

- bump cache
- force reload/update
- verify runtime file loaded

---

### KI-6.18-02 — Real polygon smoke still required

**Risk:**  
Geo protection is architecturally correct but must be validated with actual polygon edit.

**Mitigation:**  

- edit polygon in PWA/HTML
- confirm Google value
- run VBA full cycle
- confirm master received value
- confirm export did not wipe value

---

### KI-6.18-03 — Too broad GAS write blocking

**Risk:**  
A write action might be blocked even if it is harmless.

**Mitigation:**  
Accept for launch safety. Refine action list after field observation.

---

### KI-6.18-04 — Too strict local hard-block if enabled everywhere

**Risk:**  
If `withSubmitLock` blocks local save everywhere, operator productivity drops.

**Mitigation:**  
Keep soft-lock for ordinary otkup capture. Use hard-block only for workflows that require immediate server write.

---

### KI-6.18-05 — VBA module export drift

**Risk:**  
Excel VBA state may differ from repository/exported files.

**Mitigation:**  
After final test, export modules from Excel and archive/tag them.

---

## 7. Acceptance Criteria

v6.18 can be accepted when:

- VBA compiles cleanly.
- Full-cycle button runs.
- Lock is written and released.
- GAS blocks writes during active lock.
- PWA guard overlay appears.
- PWA does not convert lock condition into permanent sync error.
- PWA pending records upload after lock release.
- VBA imports those records in next cycle.
- Polygon data survives Stammdaten export.
- Service worker cache update is confirmed.

---

## 8. Launch Decision

Recommended status:

```text
Launch Candidate
```

Reason:

- main VBA/PWA/GAS race is controlled
- field work can continue
- no data loss expected from sync lock
- operator visibility improved
- geo overwrite risk controlled

Final launch should wait for the acceptance smoke listed above.

---

## v6.17 — 2026-05-09

### Summary
- v6.17 documents the document-chain hardening work performed after v6.16.
- The release is VBA/Excel desktop focused. It does not change GAS, PWA, Google Sheets transport contracts or SEF remote contracts.
- The active scope is `modAmbalaza`, `modSledljivost`, `modDokumenta` and the confirmed boundary for `modFaktura` / `frmFakturisanje`.
- The release confirms the row identity model for `tblPrijemnica`: `PrijemnicaID` is unique per physical row; `BrojPrijemnice` is the business document number and may group multiple class rows with different `PrijemnicaID` values.
- The Business Flow Professional suite passed after the changes with `111/111` pass, `0` fail and `0` skipped.

### ADDED
- [Layer: VBA / `modAmbalaza`] Added launch-hardening contract for packaging ledger writes:
  - `TrackAmbalaza` validates non-negative quantity.
  - quantity `0` is a legal no-op.
  - `tipAmb`, `entitetID` and `entitetTip` are required when quantity is positive.
  - `Smer` is accepted only as `Ulaz` or `Izlaz`.
  - `GetNextID` empty result and `AppendRow <= 0` are fail-fast errors.
  - schema reads use fail-fast column guards.
- [Layer: VBA / `modAmbalaza`] Added strict read semantics for packaging balances:
  - `GetAmbalazeStanje` treats `Ulaz` as `+Kolicina` and `Izlaz` as `-Kolicina`; unknown `Smer` is an error, not an implied direction.
  - `GetVozacAmbSaldo` treats driver balance as all active movements with matching `VozacID`; no `DokumentTip` filter is canonical.
  - open-ended date filters use independent `datumOd` / `datumDo` checks.
- [Layer: VBA / `modSledljivost`] Added success/failure monitoring around `AutoLinkOtkupOtpremnica_TX`:
  - success event after `CommitTx` is best-effort.
  - failure path emits `Monitor_Error` and `SLEDLJIVOST_AUTOLINK_FAIL` before rollback.
- [Layer: VBA / `modDokumenta`] Added dual-class-safe relink contract for orphaned faktura stavke:
  - `RelinkFakturaStavke(newPrijemnicaID, brojPrijemnice, Optional klasaFilter)` relinks only the class currently being recreated.
  - relink identity is `BrojPrijemnice + Klasa`, not plain `BrojPrijemnice`.
  - relink updates the new prijemnica row to `Fakturisano = Da` and sets `FakturaID`.
  - relink recomputes faktura status after a successful relink when `FakturaID` exists.

### CHANGED
- [Layer: VBA / `modSledljivost`] `AutoLinkOtkupOtpremnica` now uses the canonical `COL_OTK_BROJ_ZBIRNE` constant instead of a hardcoded `"BrojZbirne"` string.
- [Layer: VBA / `modSledljivost`] Updating `tblOtkup.OtpremnicaID` after an auto-link now requires exactly one matching `OtkupID` row. `FindRows` returning `Nothing`, zero rows or multiple rows is a data-integrity error.
- [Layer: VBA / `modDokumenta`] `SaveOtpremnica` now preserves error context through a local error handler and explicitly fails when `GetNextID` does not return an `OtpremnicaID`.
- [Layer: VBA / `modDokumenta`] `SavePrijemnica` now relies on the row returned by `AppendRow` instead of a follow-up `FindRows` lookup. The redundant post-append update of `KolAmbVracena` was removed because `BuildPrijemnicaRowData` already writes `COL_PRJ_KOL_AMB_VRACENA`.
- [Layer: VBA / `modDokumenta`] `GetVozacDokumenta` and `BuildZbirnaVrstaCache` now exclude stornirano otpremnica rows before returning/reporting active document data.
- [Layer: VBA / `modFaktura`] The active architecture decision was clarified: because `PrijemnicaID` is row-unique in the workbook, `CreateFaktura` may continue to use `PrijemnicaID` as its canonical row reference. A composite `PrijemnicaID + Klasa` redesign is not required.

### FIXED
- [Layer: VBA / `modAmbalaza`] Closed the silent-ledger risk where invalid `Smer` values could be interpreted as opposite directions by different read helpers.
- [Layer: VBA / `modAmbalaza`] Closed the risk of negative packaging quantity silently reversing balances.
- [Layer: VBA / `modAmbalaza`] Closed the `GetVozacAmbSaldo(datumOd only)` filter bug where an omitted `datumDo` could exclude valid rows.
- [Layer: VBA / `modSledljivost`] Closed the cross-zbirna link risk by keeping `BrojZbirne` inside the strict auto-link key and confirming cross-zbirna audit green.
- [Layer: VBA / `modDokumenta`] Closed the relink edge case where recreating Klasa I of a dual-class prijemnica could attempt to relink Klasa II before the Klasa II replacement row exists.

### Compatibility / Scope Boundaries
- No new modules are introduced.
- No GAS endpoint changes are included.
- No PWA code changes are included.
- No data migration is required.
- `Attribute VB_Name` headers should be preserved in exported `.bas` modules.
- `modAmbalaza` must not introduce a nested `TrackAmbalaza_TX` for existing document flows because caller transactions already snapshot `tblAmbalaza`.
- `modFaktura` remains `PrijemnicaID`-based because the workbook confirms one unique `PrijemnicaID` per class row. `BrojPrijemnice` is the grouping business number.

### Verification Evidence
- `RunBusinessFlowProSuite` passed `111/111`, failed `0`, skipped `0` after the changes.
- Evidence covered dual-class `Otpremnica`, `Zbirna`, `Prijemnica`, ambalaža side effects, faktura creation from both class rows, duplicate faktura prevention, invalid input rollback, stornirano exclusion, strict `BrojZbirne` auto-link and cross-zbirna audit.

### Known Limitations / Roadmap
- [RM-v6.17-01] `modFaktura` should still be hardened with exact-row guards:
  - `CreateFaktura`: `FindRows(TBL_PRIJEMNICA, COL_PRJ_ID, prijemnicaID)` should require `Count = 1`.
  - `PrintFaktura`: `FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)` should require `Count = 1`.
  - `UpdateFakturaStatus`: `FindRows(TBL_FAKTURE, COL_FAK_ID, fakturaID)` should require `Count = 1`.
  - This is hardening, not a redesign.
- [RM-v6.17-02] Add a targeted test for `RelinkFakturaStavke` orphan scenario: old/stornirano prijemnica row exists, new row with same `BrojPrijemnice + Klasa` is created, faktura stavka is relinked exactly once to the new `PrijemnicaID`.
- [RM-v6.17-03] Optional additional monitoring success events may be added to document TX wrappers, always best-effort after commit.

### Documentation Actions
- [x] Version index updated to v6.17.
- [x] Canonical reference updated with v6.17 document-chain delta.
- [x] `PrijemnicaID` / `BrojPrijemnice` identity model clarified.
- [x] v6.17 verification evidence documented.

---

---

## v6.16 — 2026-05-09

### Summary
- v6.16 documents the Agrohemija / Digitalni Agronom launch-hardening work performed after v6.15.
- The release does not introduce a new module or a new backend sync architecture.
- Desktop changes harden the existing `modAgrohemija` and `frmAgrohemija` surfaces.
- PWA changes harden the existing management `agrohemija.js` issuing flow and kooperant `agromere.js` treatment flow.
- GAS remains unchanged for this release. The current `syncTretman` route, `processTretmanRecord`, `getTretmaniForKooperant` and `buildBatchSyncResponse` are accepted as the active backend contract for treatment evidence sync.
- Automatic server-side lager decrement from a treatment remains outside the v6.16 launch scope.

### ADDED
- [Layer: VBA / `modAgrohemija`] Input validation and stock guard for warehouse saves.
  - What changed: `SaveMagacin` now uses a validation helper for required article, valid movement type, positive quantity and required kooperant on `MAG_IZLAZ`; outbound saves check current article stock before append.
  - Why: agrohemija issue rows must not create invalid or obviously negative stock movements.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modAgrohemija`] Transaction monitoring for magacin saves.
  - What changed: `SaveMagacin_TX` follows the existing transaction/monitoring pattern used by hardened modules: success emits `MAGACIN_SAVE_SUCCESS`, failure emits `MAGACIN_SAVE_FAIL` and `Monitor_Error`; success monitoring is best-effort after commit and must not turn a committed save into a false failure.
  - Why: warehouse/agrohemija saves are operationally important and must be visible in production diagnostics without introducing business risk.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `modAgrohemija`] Compatibility wrapper for supplier-state report spelling.
  - What changed: `ReportStanjePoDobavljacu()` is documented as the correct public spelling while the previous `ReportStanjePoDoabvljacu()` spelling remains as a compatibility surface.
  - Why: new code should use the correct name without breaking existing callers.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `frmAgrohemija`] Aggregated basket stock validation.
  - What changed: issue baskets are checked per article before commit, so multiple rows for the same article cannot pass individual checks and exceed total available stock when aggregated.
  - Why: form-level UX must catch whole-basket stock problems before transaction work starts.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA / Management `agrohemija.js`] Client issue record identity and submit-lock protection.
  - What changed: the management agro issuing flow now creates/stores a client-side issuance identifier for the modal payload and wraps final save through `withSubmitLock`.
  - Why: double-tap and ambiguous retry risk must be reduced on the client side while GAS `saveIzdavanje` remains unchanged.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA / Kooperant `agromere.js`] Treatment quantity/lager validation helper.
  - What changed: treatment save validates current input quantity against the locally loaded `stammdaten.magacinkoop` stock and writes the accepted value back into `agroState.kolicina` before saving.
  - Why: the treatment record must persist the same quantity that was validated in the UI.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: VBA / `modAgrohemija`] Schema guard discipline tightened.
  - What changed: required column reads in stock/report/debt/parcela paths use fail-fast schema guards such as `RequireColumnIndex`; optional columns remain explicit optional branches.
  - Why: silent zero-index reads are not acceptable in pre-launch warehouse/finance-adjacent code.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `ReportIzdavanjePoKooperantu`] Date filtering clarified.
  - What changed: open-ended filters are valid; `datumOd` and `datumDo` are evaluated independently instead of treating missing upper bound as date zero.
  - Why: report consumers must be able to request from-date-only or to-date-only ranges without accidentally filtering out all rows.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / `frmAgrohemija`] Explicit transaction rollback state.
  - What changed: form-level basket commits track whether a transaction has started and rollback in `EH` for both `btnZavrsiIzlaz_Click` and `btnZavrsiUlaz_Click`.
  - Why: a multi-line basket must commit atomically or leave no partial magacin rows.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA / Management `agrohemija.js`] Package quantity semantics corrected.
  - What changed: package rounding now makes the persisted quantity the real unit-of-measure quantity (`pakCount * pakovanje`), not the number of packages.
  - Why: desktop `frmAgrohemija`, PWA management issuing and kooperant agromere must share the same quantity semantics.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA / Management `agrohemija.js`] Parcel serialization standardized.
  - What changed: multiple selected parcel IDs in agro issuing are serialized with semicolon (`;`) to match the desktop agrohemija form.
  - Why: cross-layer reporting and future parsing should not depend on inconsistent comma/semicolon conventions.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA / Kooperant `agromere.js`] Smart dosage semantics corrected.
  - What changed: `agroCalcPreporuka()` now treats `finalQty` as the real quantity in the article unit of measure in all branches. If packaging exists, `finalQty = pakCount * pakovanje`; insufficient-stock warnings compare stock against that same real quantity.
  - Why: treatment records must not confuse package count with consumed quantity.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA / Kooperant `agromere.js`] Reset state cleanup clarified.
  - What changed: reset removes both `active` and `selected` visual classes from agro measure buttons and clears transient dosage/selection state.
  - Why: after a save or tab reload, stale visual selection must not imply active business state.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA / `SaveMagacin_TX`] False failure after committed transaction prevented.
  - Symptom: a monitoring exception after commit could incorrectly route a successful business save into the transaction error path.
  - Resolution: success monitoring is best-effort after commit and cannot invalidate a committed transaction.
  - Reference update required: Yes

- [Layer: VBA / `SaveMagacin` -> `SaveMagacin_TX`] Original failure reason preserved.
  - Symptom: validation failures such as insufficient stock could be collapsed into a generic `SaveMagacin nije uspeo` message.
  - Resolution: the service stores the last magacin failure reason so the transaction wrapper, logs and operator feedback can preserve the specific cause.
  - Reference update required: Yes

- [Layer: VBA / `frmAgrohemija`] Return navigation error label corrected.
  - Symptom: the return button error handler could log the wrong form/procedure name copied from another form.
  - Resolution: logging now identifies `frmAgrohemija.btnPovratak_Click`.
  - Reference update required: Yes

- [Layer: PWA / Management `agrohemija.js`] Modal reset and render guards tightened.
  - Symptom: stale otpremnica modal data or missing DOM targets could leave stale state or runtime errors after save/reset.
  - Resolution: reset clears modal data and dosage/cart state; render functions guard required DOM targets.
  - Reference update required: Yes

### UNCHANGED (explicit)
- [Layer: GAS / `syncTretman`] No change in v6.16.
  - Current route remains `syncTretman` with Kooperant/Management authorization, entity scoping, `records[]` validation, `withLock`, `processTretmanRecord` and `buildBatchSyncResponse`.
  - Current storage remains per-kooperant `TRETMAN-<KooperantID>` sheet.
  - Current idempotency remains keyed by `ClientRecordID` inside `processTretmanRecord`.
  - Current readback remains `getTretmaniForKooperant(kooperantID)` reading the same per-kooperant treatment sheet.

- [Layer: GAS / `saveIzdavanje`] No change in v6.16.
  - Management agro issuing continues to persist through the existing `saveIzdavanje` endpoint and existing `Izdavanje` sheet contract.
  - Client-side submit lock is the v6.16 mitigation for duplicate user submit; server-side `saveIzdavanje` idempotency remains a possible future hardening item.

### KNOWN LIMITATIONS
- Treatment save does not automatically decrement server-side `magacinkoop` stock. In v6.16 the treatment path is canonical as evidence/history/karenca sync, while lager truth remains supplied by management/exported `stammdaten.magacinkoop`.
- GAS `processTretmanRecord` is accepted unchanged. It validates core treatment fields and is idempotent by `ClientRecordID`; stricter server-side validation that `Zastita/Prihrana` must carry non-empty `ArtikalID` and positive `KolicinaUpotrebljena` remains recommended but is not part of v6.16 because GAS is intentionally unchanged.
- GAS `saveIzdavanje` is intentionally unchanged. Unknown-result retries of management agro issue saves are not fully server-idempotent unless/until a future backend change keys the endpoint by a stable client identifier.

### ROADMAP
- [RM-v6.16-01] Optional GAS-side `saveIzdavanje` idempotency by client issuance ID.
  - Affected modules: GAS `saveIzdavanje`, PWA management `agrohemija.js`.
  - Target state: retrying the same issuance after an ambiguous network response returns the existing row instead of appending another issue header.
  - Not a launch blocker for v6.16 if submit-lock smoke passes.

- [RM-v6.16-02] Optional server-side stock decrement contract for treatment consumption.
  - Affected modules: GAS `syncTretman`, management export, VBA/Excel `tblMagacin` and `magacinkoop` read model.
  - Target state: define whether a synced treatment consumes kooperant-issued stock, how storniranje works, and how retry/idempotency avoids double decrement.
  - Not a launch blocker for treatment evidence/karenca launch.

- [RM-v6.16-03] Optional stricter GAS validation for treatment articles and quantities.
  - Affected modules: GAS `processTretmanRecord`.
  - Target state: reject `Zastita` / `Prihrana` rows that do not include `ArtikalID` and positive `KolicinaUpotrebljena`.
  - Not included in v6.16 because GAS remains unchanged.

### VERIFIED / TO VERIFY
- VBA smoke: issue agrohemija with enough stock; verify all basket rows commit and monitoring success is best-effort.
- VBA smoke: issue the same article in multiple basket rows exceeding total stock; verify pre-commit aggregated validation blocks and no rows are saved.
- VBA smoke: force a second-line save failure; verify rollback removes all earlier basket rows.
- VBA smoke: `ReportIzdavanjePoKooperantu` with only `datumOd` and only `datumDo` returns expected ranges.
- PWA management smoke: dosage requiring 2.4 kg with 1 kg packaging saves/displays 3 kg, not 3 packages.
- PWA management smoke: selected parcel IDs serialize with `;`.
- PWA management smoke: double-click final save produces one user submit due to submit lock.
- Kooperant smoke: `agroCalcPreporuka` with packaging writes real JM quantity into `agroKolicina` and treatment record.
- Kooperant smoke: quantity greater than local `magacinkoop` stock is blocked before local treatment save.
- Kooperant smoke: offline treatment remains pending; online sync returns existing/inserted result and history reload shows one deduplicated treatment.
- GAS unchanged smoke: existing `syncTretman` path still returns `results[]` and PWA moves records out of pending.

### Migration / Data Notes
- No production-data migration.
- Existing treatment and issue rows are not rewritten.
- No repair/backfill for disposable test data.
- Service-worker `CACHE_NAME` must be bumped if the changed PWA files are part of the deployed app-shell cache.

### Documentation Actions
- [x] Version index updated to v6.16.
- [x] Canonical reference updated with v6.16 Agrohemija / Digitalni Agronom delta.
- [x] GAS unchanged boundary documented.
- [x] Known limitations and roadmap split documented.

---

## v6.15 — 2026-05-07

### Summary
- v6.15 reverses the v6.12 / v6.14.1 `BrojZbirne` ownership decision.
- `BrojZbirne` is now generated PWA-side at `confirmZbirna` time using deterministic local context (numeric VozacID, today's date, count of today's zbirne for that driver including soft-deleted).
- The format `x/ddmmyy[-rb]` is unchanged.
- VBA `GenerateBrojZbirne` is preserved as a fallback path for legacy/pre-rollout VOZ rows and for desktop manual entry through `frmDokumenta`.
- GAS `processZbirnaRecord` already accepted `record.brojZbirne` and wrote it to the `BrojZbirne` column; no GAS code change is required.
- The operational driver was that v6.12/v6.14.1 design required the operator to run MasterSync before the driver could see the business `BrojZbirne` and print zbirna at the last otkup station. This blocks transport handoff.

### Reversal Notice
- v6.12 § "VBA/GAS/PWA Zbirna Boundary" declared `BrojZbirne` ownership in VBA Master import.
- v6.14.1 hardened the VOZ writeback so column T received the VBA-generated business number.
- v6.15 explicitly reverses this primary ownership: the canonical generator is now PWA `confirmZbirna`. The VBA generator is retained but demoted to fallback role.
- The reversal is recorded here so future readers understand the intent of removing v6.12's "post-VBA ownership" wording from the canonical reference.

### CHANGED
- [Layer: PWA / Vozac] `confirmZbirna` now generates `BrojZbirne` locally before `dbPut`.
  - What changed: `confirmZbirnaUnlocked()` in `vozac/zbirna.js` computes `brojZbirne` from `CONFIG.ENTITY_ID` (numeric part), today's date in `ddmmyy` and the count of today's zbirne (including soft-deleted, so storno does not reclaim numbers) for the current driver.
  - Why: the driver must print the zbirna at the last otkup station before transporting goods to the buyer. The previous flow blocked printing until the operator ran MasterSync.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / MasterSync / VOZ Import] `ImportRowToTblZbirna` reads `BrojZbirne` from VOZ sheet first, falls back to `GenerateBrojZbirne` only when empty.
  - What changed: `ImportRowToTblZbirna` reads VOZ column 20 (`BrojZbirne`) into the local variable. If non-empty, the value is validated against the canonical format regex `^\d+/\d{6}(-\d+)?$` via the new private `IsValidBrojZbirneFormat` helper. If empty, it falls back to `GenerateBrojZbirne(vozacID, datum)` and logs the fallback at `WARN`.
  - Why: PWA-generated values must flow through unchanged; legacy/pre-rollout rows with empty BrojZbirne must still import cleanly.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / MasterSync / VOZ Import] `IsValidBrojZbirneFormat` private helper added.
  - What changed: regex-based format validator for canonical `x/ddmmyy[-rb]` form.
  - Why: PWA-generated values must be defended against client-side bugs producing malformed numbers; rejecting at import is safer than persisting malformed data.
  - Reference update required: Yes
  - Migration required: No

### UNCHANGED (explicit)
- [Layer: GAS / processZbirnaRecord] No change required.
  - The handler already reads `record.brojZbirne` from the payload and writes it to the `BrojZbirne` column of the VOZ sheet through the canonical row map. v6.15 simply means PWA now sends a populated value where it previously sent empty.
- [Layer: VBA / SaveZbirna, SaveZbirna_TX, SaveZbirnaMulti_TX] Signatures unchanged.
  - `brojZbirne` is already a parameter; the change in v6.15 is upstream of these functions.
- [Layer: VBA / WriteBackVOZSyncStatus] Unchanged.
  - Continues to write column B / `ServerRecordID` from `update(2)` and column T / `BrojZbirne` from `update(3)`. In the new flow most zbirne already carry `BrojZbirne` in column T from initial `processZbirnaRecord` write; the VBA writeback is now an idempotent reaffirmation in the typical case and a real backfill in the fallback case.
- [Layer: VBA / GenerateBrojZbirne, ExtractNumericVozacBroj] Unchanged.
  - Retained as fallback owner for legacy VOZ rows and desktop-manual entry through `frmDokumenta`.

### KNOWN LIMITATIONS
- Multi-device same-driver offline collision is theoretically possible because PWA computes the daily sequence locally without server consultation. In practice this is bounded by the operational reality of one device per driver and three trips per day. A defensive duplicate guard in `processZbirnaRecord` is documented as ROADMAP, not as launch requirement.
- Storno does not reclaim sequence numbers. A storno-and-recreate cycle within the same day produces a gap, which is the expected accounting-friendly behavior.

### KNOWN ISSUE
- [KI-v6.15-01] Multi-device same-driver offline collision risk.
  - Affected layer: PWA / GAS Zbirna sync boundary
  - Impact: if a single driver runs the PWA on two devices offline simultaneously, both can generate the same `x/ddmmyy[-rb]` and both syncs will currently succeed.
  - Workaround: operational. One device per driver is the supported configuration.
  - Reference update required: Yes

### ROADMAP
- [RM-v6.15-01] Optional GAS-side defensive duplicate guard for `BrojZbirne` per `(VozacID, BrojZbirne)`.
  - Why it matters: covers the multi-device edge case end-to-end.
  - Affected modules: GAS `processZbirnaRecord`.
  - Target state: reject second insertion with explicit `DUPLICATE_BROJ_ZBIRNE` error code so PWA can regenerate.
  - Not a launch blocker.

### VERIFIED / TO VERIFY
- PWA smoke: create zbirna offline, verify `brojZbirne` populated immediately, verify printable from last station before sync.
- PWA smoke: create two zbirne the same day, verify second is `x/ddmmyy-2`.
- PWA smoke: create one zbirna, soft-delete it, create new one, verify new is `x/ddmmyy-2` (storno does not reclaim).
- VBA smoke: import VOZ row carrying populated `BrojZbirne`, verify `tblZbirna.BrojZbirne` matches PWA value, no fallback path taken.
- VBA smoke: import legacy VOZ row with empty `BrojZbirne`, verify `GenerateBrojZbirne` fallback fires, WARN logged.
- VBA smoke: inject malformed `BrojZbirne` (e.g. `7/060526-` or `abc`), verify import is rejected with format error logged.
- Cross-system: verify `LinkZbirnaToOtkupAndOtpremnica` cascade still propagates the PWA-generated `BrojZbirne` into `tblOtkup` and `tblOtpremnica`.

### Migration / Data Notes
- No production-data migration.
- Existing VOZ rows with empty `BrojZbirne` (pre-rollout) continue to flow through the VBA fallback path.
- Existing `tblZbirna` rows are unaffected.

### Documentation Actions
- [ ] Section 1.7.8 of canonical reference updated to record the PWA-first generation rule and the VBA fallback role.
- [ ] New section 1.10 v6.15 BrojZbirne PWA-First Generation Delta added.
- [ ] Architecture invariant table updated where it references VBA-owned BrojZbirne generation.
- [ ] Section 5.4 Zbirna Flow updated to record PWA-side numbering step before sync.

---

## v6.14.1 — 2026-05-06

### Summary
- v6.14.1 is a document-flow launch-hardening correction on top of the v6.14 monitoring baseline.
- The delta clarifies that otpremnica can be created before zbirna and may initially have empty `BrojZbirne`.
- The VOZ import/writeback contract is tightened so Google column B receives the technical/master `ZbirnaID`, while Google column T receives the generated business `BrojZbirne`.
- No production-data migration is included.

### CHANGED
- [Layer: VBA / Dokumenta] Otpremnica `BrojZbirne` optionality clarified.
  - What changed: `ValidateOtpremnicaInput` must not reject empty `BrojZbirne`.
  - Why: otpremnica can be created before the vozač creates/imports the matching zbirna. `BrojZbirne` is populated later through the VOZ import cascade link.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / MasterSync / VOZ] VOZ import core contract tightened.
  - What changed: `ImportZbirneFromPWA_TX` must call a non-UI core import function returning `Boolean`, and must rollback instead of commit when row/import/writeback errors are present.
  - Why: the transaction wrapper snapshots `tblZbirna`, `tblOtpremnica` and `tblOtkup`, but it cannot make a correct commit/rollback decision if the public import entrypoint is only a message-showing `Sub`.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA / MasterSync / VOZ] VOZ cascade-link updates are traceability-critical.
  - What changed: updates from `LinkZbirnaToOtkupAndOtpremnica` to `tblOtkup.BrojZbirne` and `tblOtpremnica.BrojZbirne` must use checked/fail-fast update semantics.
  - Why: a successful zbirna import without propagated `BrojZbirne` breaks the canonical trace bridge from raw/PWA otkup into the document chain.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: VBA / MasterSync / VOZ Writeback] `BrojZbirne` writeback payload index.
  - Symptom: `WriteBackVOZSyncStatus` could write `update(2)` to both Google column B and Google column T, causing `BrojZbirne` to contain an internal `ZBR-*` ID instead of the generated business number.
  - Resolution: Google `Sheet1!B` receives `update(2)` / `ZbirnaID`; Google `Sheet1!T` receives `update(3)` / business `BrojZbirne`.
  - Reference update required: Yes

- [Layer: VBA / MasterSync / VOZ Idempotency] Empty `ClientRecordID` guard.
  - Symptom: a VOZ row with empty `ClientRecordID` could pass duplicate detection as non-duplicate because `IsDuplicateZbirnaInMaster("")` returns false.
  - Resolution: VOZ import must mark such rows as `SyncError:ClientRecordID missing` and skip import.
  - Reference update required: Yes

### VERIFIED / TO VERIFY
- Run VOZ document-flow smoke:
  - create/import PWA otkup;
  - auto-create otpremnica with empty `BrojZbirne`;
  - create/sync VOZ zbirna;
  - run `ImportZbirneFromPWA_TX`;
  - verify `tblZbirna.BrojZbirne` contains the generated business number;
  - verify `tblOtkup.BrojZbirne` and `tblOtpremnica.BrojZbirne` are populated;
  - verify Google VOZ `Sheet1!B = ZBR-*`;
  - verify Google VOZ `Sheet1!T = <business BrojZbirne>`, not `ZBR-*`;
  - verify Google VOZ `Sheet1!F = Synced>Master`.

### Migration / Data Notes
- No production-data migration.
- No repair/backfill for disposable test data.
- Existing otpremnice with empty `BrojZbirne` are valid pending-link documents until the matching zbirna import cascade populates them.

---

## v6.14 — 2026-05-06

### Summary
- v6.14 introduces the production monitoring and observability layer for AgriX / OtkupApp.
- The release connects VBA, GAS and the monitoring Google Sheets workbook through a best-effort event pipeline.
- The monitoring workbook `OtkupApp_Monitoring_PROD` now owns runtime `Health`, central `Events`, structured `Errors`, `SEFStatus`, `Backups`, `Alerts` and `AuditCritical` views.
- VBA monitoring covers app lifecycle, SEF, faktura, novac, otkup, document-chain failures, bank mapping, MasterSync, Stammdaten export and backup signals.
- GAS monitoring covers public ingest, authenticated monitoring, event normalization, sanitization, health updates, alert creation, watchdog checks and daily summaries.
- No production-data migration is included.

### ADDED
- [Layer: GAS/Monitoring] Dedicated `Monitoring.gs` observability layer.
  - What changed: monitoring ingest, normalization, routing, workbook initialization, alerts, health checks, watchdog and summary jobs are part of the active architecture.
  - Why: production/pilot diagnosis must be possible without opening local Excel logs or relying on user screenshots.
  - Reference update required: Yes
  - Migration required: No

- [Layer: Google Sheets/Monitoring] Canonical monitoring workbook `OtkupApp_Monitoring_PROD`.
  - What changed: monitoring storage is organized into `Health`, `Events`, `Errors`, `SyncStatus`, `SEFStatus`, `UserSessions`, `Backups`, `Alerts` and `AuditCritical` tabs.
  - Why: runtime health, errors, SEF state, backups, alerts and audit-critical events need a single operator-visible surface.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Monitoring] Best-effort VBA monitoring client.
  - What changed: VBA emits structured events to GAS through `MONITORING_ENDPOINT` and `MONITORING_SECRET` read from `tblSEFConfig`; app version is read from `modConfig.APP_VERSION`; device identity is included in every event.
  - Why: desktop production issues must be remotely diagnosable while preserving the existing local log/journal model.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/App Lifecycle] Startup monitoring coverage.
  - What changed: `Workbook_Open` / `StartApp` now emit `VBA_APP_OPEN`, `VBA_STARTAPP_START`, `VBA_STARTAPP_SUCCESS`, `JOURNAL_RECOVERY_WARN` and structured startup errors.
  - Why: app boot, startup failure and journal recovery warnings are operationally important during pilot and production support.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/SEF Observability] Detailed SEF operational monitoring.
  - What changed: SEF recovery, send and status-refresh flows now emit `SEF_*` events for start, success, rejection, failure, unknown/manual-review and stuck/recovery states.
  - Why: SEF failures can have tax/invoice impact and need immediate visibility, manual-review routing and audit trace.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Business TX Monitoring] Critical transaction-boundary monitoring.
  - What changed: monitoring is added at selected `_TX` boundaries for faktura, novac, avans allocation, otkup, multi-class otkup, document-chain failures, bank mapping, MasterSync and Stammdaten export.
  - Why: critical business outcomes should be visible remotely without logging every helper/read/write operation.
  - Reference update required: Yes
  - Migration required: No

- [Layer: GAS/Health] Active Health row completion.
  - What changed: watchdog checks now actively maintain `GAS API`, `Google Sheets DB`, `Auth`, `Backup` and `MasterData Sync` health rows in addition to event-driven components.
  - Why: Health must not contain empty rows for components that are checked by system state rather than frequent user events.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/MasterData] MasterSync and Stammdaten monitoring events.
  - What changed: `modMasterSync` emits `MASTERDATA_SYNC_SUCCESS` / `MASTERDATA_SYNC_FAIL`; `modStammdatenSync.SyncStammdatenToGoogle` emits `STAMMDATEN_SYNC_SUCCESS` / `STAMMDATEN_SYNC_FAIL`.
  - Why: `MasterData Sync` health must be driven by real import/export outcomes.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: GAS/Monitoring Config] Monitoring GAS constants use Script Property names.
  - Previous behavior: property constants could be confused with actual values.
  - New behavior: `MONITORING_PROP_SPREADSHEET_ID`, `MONITORING_PROP_ALERT_EMAIL` and `MONITORING_PROP_INGEST_SECRET` are property keys whose values live in Apps Script Project Settings.
  - Why: secrets and deployment-specific IDs must not be hardcoded in source.
  - Reference update required: Yes
  - Migration required: No data migration; Script Properties must be configured.

- [Layer: VBA/Monitoring Security] Monitoring debug and payload handling is privacy-bounded.
  - Previous behavior: debug/test output could expose full JSON bodies if left unguarded.
  - New behavior: debug output prints only redacted diagnostics; messages and payloads are sanitized/truncated before send.
  - Why: monitoring must not leak tokens, secrets, SEF keys, XML/PDF payloads or sensitive transport data.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Monitoring UX] Production monitoring timeout is short and best-effort.
  - Previous behavior: business operations could be exposed to long monitoring waits if monitoring used business-API-scale HTTP timeouts.
  - New behavior: production monitoring uses a short timeout, while debug connectivity tests may wait longer.
  - Why: monitoring must never noticeably block operator workflow.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/Observability Scope] Monitoring is explicitly narrow and non-replacing.
  - Previous behavior: remote monitoring was not a documented production layer.
  - New behavior: monitoring complements `LogErr`, journals, `AppendSEFEvent_Row`, SEF persistence and `ProductionHealthCheck`; it does not replace them.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: Monitoring/Health] Empty Health component rows.
  - Symptom: components such as `GAS API`, `Google Sheets DB`, `Backup`, `Auth` and `MasterData Sync` could remain empty until a matching event arrived.
  - Resolution: active watchdog checks now set explicit `OK`, `WARN` or `CRITICAL` snapshots for these rows.
  - Reference update required: Yes

- [Layer: Monitoring/Security] Secret/body leakage risk in debug testing.
  - Symptom: full JSON debug output could include `monitoringSecret` or other sensitive fields.
  - Resolution: debug output is redacted and payload/message sanitization is part of the monitoring contract.
  - Reference update required: Yes

- [Layer: MasterData Health] Missing master-data sync signal.
  - Symptom: `MasterData Sync` could remain unknown/empty even after real sync/export operations.
  - Resolution: `MASTERDATA_SYNC_*` and `STAMMDATEN_SYNC_*` events now update the `MasterData Sync` component.
  - Reference update required: Yes

### VERIFIED
- [Layer: VBA/GAS Monitoring] End-to-end ingest confirmed.
  - Evidence: VBA HTTP diagnostic returned HTTP 200 with `success: true`, generated `eventId`, timestamp, severity and `component: VBA Client`.
  - Acceptance: VBA -> GAS Web App -> `monitorPublic` -> Monitoring workbook write path is operational.

- [Layer: Monitoring Config] `tblSEFConfig` config contract confirmed.
  - Evidence: `MONITORING_ENDPOINT`, `MONITORING_SECRET`, `MONITORING_ENV`, `APP_VERSION` and local `DeviceId` are resolved by the VBA client.
  - Acceptance: monitoring config does not require hardcoded endpoint/secret/env values in VBA.

- [Layer: Monitoring Tests] VBA monitoring test suite established.
  - Evidence: `TestMonitoring_All`, `TestMonitoring_Config`, `TestMonitoring_HTTP`, `TestMonitoring_ErrorEvent`, `TestMonitoring_SEFUnknown`, `TestMonitoring_BackupSuccess` and `TestMonitoring_BackupFail` define the active smoke surface.
  - Acceptance: monitoring health/events/errors/SEF/backups/alerts/audit-critical routing is testable from VBA.

### DEFERRED
- [Layer: Monitoring/Noise Tuning] Post-pilot event-volume tuning.
  - What changed: selected critical TX boundaries are instrumented; noisy helper/read-level instrumentation remains intentionally excluded.
  - Why: pilot data should confirm whether success events for otkup/novac/bank mapping remain useful or need aggregation.
  - Reference update required: Yes
  - Migration required: No

### KNOWN ISSUE
- [Layer: Monitoring/Operational Setup] Monitoring depends on deployment configuration.
  - Current state: monitoring requires the GAS Web App deployment URL, Script Properties, `tblSEFConfig` values and triggers to be correctly installed.
  - Impact: missing configuration produces `DEV`, `WARN` or `CRITICAL` health states rather than a business-data migration problem.
  - Reference update required: Yes
  - Migration required: No

### Migration / Data Notes
- No production-data migration.
- Monitoring workbook tabs and headers are schema/setup artifacts, not business-data migration.
- Required operational setup: `MONITORING_SPREADSHEET_ID`, `MONITORING_ALERT_EMAIL`, `MONITORING_INGEST_SECRET`, `MONITORING_ENDPOINT`, `MONITORING_SECRET`, `MONITORING_ENV`, GAS Web App redeploy, watchdog trigger and daily summary trigger.

### Documentation Actions
- [x] Canonical reference updated for production monitoring architecture.
- [x] Changelog updated for v6.14.
- [x] Monitoring workbook tabs and component health model documented.
- [x] VBA event surface documented.
- [x] GAS ingest, watchdog, alerts and audit-critical routing documented.
- [x] Security/redaction and timeout rules documented.

---

## v6.13 — 2026-05-04

### Summary
- v6.13 closes the remaining PWA pre-launch P0 runtime-hardening issues after the v6.12 launch-smoke baseline.
- The release stabilizes app-shell/cache behavior, unified render dedupe, sync trigger ownership, stale `syncing` recovery and critical save submit-locking.
- No production-data migration is included.
- Minimal unit tests remain deferred as a separate P-1/post-launch safety-net item.

### ADDED
- [Layer: PWA/Merge + Render] Canonical alias-aware render dedupe helper `dedupeRecordsForRender(records, aliasesFn?)` in `src/js/utils/merge.js`.
  - What changed: key UI lists now use one shared dedupe helper before render.
  - Why: local/server merge, reconnect refresh, partial sync and stale-cache/fresh-response combinations must not render the same logical row twice.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Sync] Public sync request wrappers.
  - What changed: `requestRoleSync(reason)`, `requestOtkupSync(reason)` and `requestKooperantSync(reason)` are active wrapper entrypoints around role/store sync.
  - Why: manual, online, interval and post-save triggers must not call low-level sync functions directly.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Sync Recovery] Bootstrap stale-`syncing` recovery helpers.
  - What changed: `recoverStaleSyncingRecords(storeName)`, `recoverStaleSyncingStores(storeNames)` and `recoverStaleSyncingForCurrentRole(reason)` are exposed and used during bootstrap.
  - Why: rows stuck in `syncing` after a crash/reload/network interruption must become retryable before the first role render and sync badge update.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Runtime UX] Shared critical-submit lock helper `withSubmitLock(lockKey, fn, options)`.
  - What changed: the helper stores locks in `window.appRuntime.submitLocks`, can disable matching action buttons and clears locks in `finally`.
  - Why: double-tap or repeated clicks on field devices must not create duplicate local records.
  - Reference update required: Yes
  - Migration required: No

### CHANGED
- [Layer: PWA/App Shell] Runtime stability items were closed for role navigation and app-shell loading.
  - What changed: `tabs.js` guards `agroState` for non-Kooperant roles; `role-nav.js` uses canonical `cfg.type`; `db.js` is expected to load once; service-worker cache discipline was tightened.
  - Why: non-Kooperant roles, management navigation and mixed-cache deploys must not crash field usage.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Service Worker] Critical runtime asset deploy discipline was tightened.
  - What changed: `CACHE_NAME` must be bumped when critical JS/runtime assets change, and Leaflet marker image assets are part of the offline app-shell contract when maps are launch-relevant.
  - Why: deploys must not leave users with a mixed old/new runtime asset set.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Render Paths] Main list render paths now pass through unified dedupe.
  - What changed: dedupe is applied to Otkup queue, Otkup pregled, Vozač zbirna pregled, Kooperant treatment history, Otkup otprema overview and Otkup otprema assign runtime state.
  - Why: duplicate record display must be solved centrally rather than by feature-specific local maps.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Sync] `syncQueueSafe(reason)` is the single app-level sync trigger entrypoint.
  - What changed: online, interval, manual and post-save paths pass explicit reason strings to `syncQueueSafe(...)`; `runRoleSync(reason)` is now only a legacy alias to `requestRoleSync(reason)`.
  - Why: parallel sync attempts must be prevented at the trigger layer before reaching store-level sync.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Vozac Zbirna] Zbirna post-save sync now follows the app-level sync gate.
  - What changed: `confirmZbirnaUnlocked()` uses `syncQueueSafe('post-save')` instead of direct low-level `syncZbirne()`.
  - Why: the single sync-entrypoint rule must also apply to Vozač post-save flows.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Save Flow] Critical save functions were split into locked public wrappers and unlocked implementations.
  - What changed: `saveOtkup()`, `confirmZbirna()` and `agroSaveTretman()` call `withSubmitLock(...)` and delegate to `saveOtkupUnlocked()`, `confirmZbirnaUnlocked()` and `agroSaveTretmanUnlocked()`.
  - Why: business logic stays unchanged while double-submit protection is centralized.
  - Reference update required: Yes
  - Migration required: No

### FIXED
- [Layer: PWA/Render Dedupe] Alias-case duplicate rows.
  - Symptom: the same logical record could render twice when one copy had `{ serverRecordID, clientRecordID }` and another local copy had only `{ clientRecordID }`.
  - Resolution: `dedupeRecordsForRender(...)` groups identity aliases and keeps the local pending/syncing/error version over server-synced rows.
  - Reference update required: Yes

- [Layer: PWA/Sync] Parallel sync trigger race.
  - Symptom: manual sync plus online event/background interval/post-save could attempt overlapping sync operations.
  - Resolution: all triggers now go through `syncQueueSafe(reason)` and role/store wrappers with runtime in-flight flags.
  - Reference update required: Yes

- [Layer: PWA/Sync Recovery] Stale rows stuck as `syncing` before first sync.
  - Symptom: after interruption, rows could stay visually/internally stuck as `syncing` until another sync attempt started.
  - Resolution: role-aware stale recovery now runs during bootstrap, before role render and sync badge calculation.
  - Reference update required: Yes

- [Layer: PWA/Save UX] Double-tap duplicate local record risk.
  - Symptom: rapid repeated taps on critical save buttons could create more than one local record.
  - Resolution: shared submit locks now wrap Otkup save, Zbirna confirm and Kooperant treatment save flows.
  - Reference update required: Yes

### VERIFIED
- [Layer: PWA/Dedupe] Runtime coverage confirmed.
  - Evidence: `dedupeRecordsForRender` is available and used by Otkup queue, Otkup pregled, Zbirna pregled, treatment history, Otprema overview and Otprema assign state.
  - Acceptance: alias test returns one row `['local-pending']` in both input orders.

- [Layer: PWA/Sync] Runtime routing and parallel guard confirmed.
  - Evidence: `requestRoleSync`, `requestOtkupSync`, `requestKooperantSync` exist; online and interval use explicit reasons; Otkup post-save has no low-level `syncQueue()` fallback.
  - Acceptance: concurrent `syncQueueSafe('manual')` + `syncQueueSafe('online')` as Otkupac returned one normal/no-pending result and one `already-running / ALREADY_RUNNING`; runtime flags reset afterwards.

- [Layer: PWA/Stale Recovery] Bootstrap recovery confirmed.
  - Evidence: recovery helpers exist globally and `bootstrapApp()` calls `recoverStaleSyncingForCurrentRole('bootstrap')`.
  - Acceptance: manual stale test restored a `syncing` row to `pending` and set `lastServerStatus = 'stale-syncing-recovered'`.

- [Layer: PWA/Submit Lock] Submit-lock structure confirmed.
  - Evidence: `withSubmitLock` exists; `saveOtkup`, `confirmZbirna` and `agroSaveTretman` use it; `confirmZbirnaUnlocked` uses `syncQueueSafe('post-save')`.
  - Acceptance: Otkupac double-tap smoke created one record in the 5-minute test window.

### DEFERRED
- [Layer: Tests] Minimal unit test safety net.
  - What changed: unit tests for sync, merge, dedupe and stale recovery remain deferred.
  - Why: v6.13 is pre-launch runtime hardening; unit tests remain a P-1/post-launch safety item.
  - Reference update required: Yes
  - Migration required: No

- [Layer: PWA/Submit Lock Smoke] Remaining role double-tap smoke.
  - What changed: implementation is complete and Otkupac smoke passed; strict acceptance should still smoke Vozač `confirmZbirna` and Kooperant `agroSaveTretman`.
  - Why: both flows are structurally locked but still deserve role-specific field confirmation.
  - Reference update required: Yes
  - Migration required: No

### KNOWN ISSUE
- [Layer: GAS/Auth] `saveParcelPolygon` public write decision remains open unless explicitly accepted.
  - Current state: endpoint is a known public-write decision point.
  - Recommended outcome: lock under token before real production launch.
  - Reference update required: Yes
  - Migration required: No

- [Layer: VBA/VOZ Writeback] `BrojZbirne` writeback fix remains relevant if VBA document-flow launch is in scope.
  - Current state: `WriteBackVOZSyncStatus` must write column B / `ServerRecordID` from `update(2)` and column T / `BrojZbirne` from `update(3)`.
  - Why: `BrojZbirne` is the VBA-owned business document number, not the technical sync ID.
  - Reference update required: Already documented in v6.12/v6.13 reference
  - Migration required: No test-data backfill

### Migration / Data Notes
- No production-data migration.
- No repair/backfill for disposable test data.
- Service-worker cache version must be bumped before production deploy.

---

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

---

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

---

## 3. Legacy Changelog Summary

Detailed legacy entries are archived in `docs/archive/CHANGELOG_legacy_v2_to_v6_17.md`. The compact table below is retained in the active changelog because older versions contain important migration, source-of-truth and endpoint-history context.

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
| v6.9 | 2026-04-29 | Business-core hardening after v6.8: modFaktura canonical prijemnica values and status/print guards, modDokumenta EH/input/stornirano hardening, modOtkup validation/read-helper hardening, expanded Faktura and BusinessFlowPro regression suites | Yes | no schema migration; compile + RunFakturaSmokeSuite + RunBusinessFlowProSuite passed |
| v6.8 | 2026-04-29 | frmSEF operator-shell cleanup and modNovac finance hardening: guarded SEF form activation/send cleanup, destructive action confirmations, SaveNovac validation, avans split fail-fast, partner-map conflict guard, otkup payment recompute, reset-link TX and Novac smoke suite | Yes | no schema/data migration; compile + RunNovacSmokeSuite passed |
| v6.7 | 2026-04-28 | SEF P0/P1 follow-up hardening: external status persistence on submit, missing `SEFDocumentId` fail-fast, refresh helper convergence, EH preservation, transition matrix suite, stricter cancel/storno tests, mapper total consistency and lightweight parser hardening | Yes | no schema/data migration; compile + SEF offline/state-transition/refresh smoke required |
| v6.6 | 2026-04-27 | Desktop tested baseline: strict BrojZbirne traceability auto-link, professional business-flow regression suite, SEF live submit/refresh evidence, SEF DeliveryDate/InvoiceDate validation, destructive cancel/storno test scaffolding | Yes | no data migration; compile + business-flow + SEF smoke passed; cancel/storno final outcome still P1 |
| v6.5 | 2026-04-26 | SEF P0/P1 desktop hardening, Stammdaten update-guard convergence and frmOtkupAPP shell save/navigation cleanup | Yes | no data migration; compile OK; SEF smoke test required |


### 3.1 Legacy Handling Rule

- Do not copy full legacy release prose back into `ARCHITECTURE_REFERENCE.md`.
- If a legacy item is still active architecture, keep the timeless rule in AR and the historical explanation here/archive.
- If a legacy item is resolved, keep it only as changelog/archive context.

---

## 4. Changelog Maintenance Policy

- New entries are added at the top under `Maintained Version Entries`.
- Keep the latest release detailed enough for engineering handoff.
- Compress older releases only after preserving migration, breaking-change and acceptance-gate notes in archive.
- Do not restate full current architecture in CL; link conceptually to the AR section instead.
- When source documents conflict, mark `NEEDS REVIEW` instead of guessing.
