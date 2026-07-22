# Refaktor playbook — sprovođenje plana ispravki (jedan paket = jedna sesija)

**Izvori:** `docs/AUDIT_FM_TRIJAZA.md` (detalji svake stavke, dokazi fajl:linija) ·
`docs/KNOWN_ISSUES.md` §8 (AUD registar) · `docs/ROADMAP.md` §10 (Wave plan).
**Model rada:** svaki paket (RF-xx) se radi u NOVOM chatu, na svojoj grani, serijski
(paket se merge-uje u `main` pre početka sledećeg). Jedan paket ≈ jedna sesija.

---

## 1. Prompt šablon za novu sesiju

Zalepiti u nov chat (Claude Code na ovom repou — CLAUDE.md se učitava automatski):

```
Radi paket RF-XX iz docs/REFAKTOR_PLAYBOOK.md.
Pre izmena: pročitaj sekciju paketa u playbook-u, navedene AUD/FM stavke u
docs/AUDIT_FM_TRIJAZA.md i SVE module iz obima (cele, ne isečke).
Radi ISKLJUČIVO obim paketa — ništa van njega. Na kraju: statičke provere,
commit+push na granu claude/rf-XX-<slug>, git komande za preuzimanje,
numerisana Excel test checklista, i ažuriraj Status tabelu u playbook-u.
```

## 2. Radna pravila (važe za svaki paket)

1. **Grana:** `claude/rf-XX-<slug>` sa svežeg `main`-a. Serijski — ne počinjati novi
   paket dok prethodni nije merge-ovan (paketi dele fajlove).
2. **Obim je zakon.** Radi se samo ono što paket navodi. Ako se usput otkrije nov
   problem: NE popravljati — upisati kao novu stavku u `KNOWN_ISSUES.md` §8 (AUD-0xx)
   i nastaviti.
3. **Minimal delta** (CLAUDE.md doktrina): bez novih apstrakcija, bez preimenovanja,
   bez „usputnog sređivanja". Postojeći obrasci su šablon: `RequireColumns`/
   `RequireUpdateCell`, `_TX` + ne-TX core, `BuildPrijemnicaRowData` (upis po imenu),
   runtime paneli, `Poruka("KLJUC")` katalog.
4. **VBA ograničenja:** izvori ostaju 100% ASCII (dijakritika SAMO kroz
   `modPoruke.UpsertPoruke` + `ChrW`); `.frx` se ne dira; nove kontrole samo runtime
   (`Controls.Add` + WithEvents).
5. **Statičke provere pre commit-a:** `file` = „ASCII text" za svaki izmenjen VBA fajl;
   grep ne-ASCII = prazno; balans `Sub/Function/End`; nema duplih `Public` definicija
   (`grep -h "^Public " src-vba/*.bas | sort | uniq -d`); svaki novi `Poruka("X")` ima
   par u `UpsertPoruke`.
6. **Kraj svakog paketa:** commit sa jasnom porukom → push → git komande za
   preuzimanje grane → **numerisana Excel test checklista** (klik-po-klik, fokus samo
   na izmene paketa + navedeni regression suite) → ažurirati Status tabelu (§4).
7. **Merge u main:** tek posle korisnikovog Excel testa (`ImportAllVBA` → Compile →
   checklista). Posle merge-a: `tools/release.sh` po `docs/RELEASE_PROCEDURE.md` kada
   se skupi smisleni skup paketa (npr. posle RF-05, RF-08, RF-14).

## 3. Paketi (redosled izvršavanja)

> Detalji svake stavke (tačne linije, predlog, napor) su u `docs/AUDIT_FM_TRIJAZA.md`
> pod navedenim FM/AUD referencama — ovde je samo obim i definicija gotovog.

### RF-01 — Brisanje balasta [Wave 0 · S]
**Fajlovi:** `src-vba/` (brisanje), `modE2EReleaseGate.bas` (bez izmene — verifikacija).
**Obim:** obrisati `modNovacTest.bas`, `modFakturaTest.bas`, `modLicenceTests.bas`
(ostaju `*Tests` verzije), `modBankaImportParserClipboard.bas`; u `modBankaImport.bas`
obrisati `GetFileNameFromPath2` (poziv :1163 preusmeriti na original); u
`modArrayUtils.bas` obrisati mrtve `GroupBySum`/`SumColumn`; u `modIzvestaj.bas`
mrtav enum `IzvestajTip`; u `frmIzvestaj.frm` mrtve `UpdateUnosButtonState` i
`PrijemniceZaOtpremnicu`. [AUD-016; FM-0027 #6/#7/#34; FM-0029 #28/#29]
**Gotovo kad:** grep pokazuje 0 referenci na obrisano; E2E gate `Application.Run`
imena sada jednoznačna. **Checklista mora reći:** obrisane module ukloniti i ručno u
VBE jednom (ImportAllVBA ne briše komponente).

### RF-02 — modNovac finansijski guardovi [Wave 1 · S/M]
**Fajlovi:** `modNovac.bas`, `modCenovnik.bas`.
**Obim:** (1) `RequireColumns` za svih 17 kolona na vrhu `SaveNovac` (P0, AUD-003);
(2) `ApplyAvansToOtkup/Faktura`: target-active (storno) + target-owner provera +
`_TX` vraća primenjeni iznos umesto no-op `True` (AUD-010; FM-0019 #4/#5/#6/#11);
(3) `AddCena` presence guard (AUD-003). **Regression:** `RunNovacSmokeSuite`,
`RunBusinessFlowProSuite`.

### RF-03 — Storno „lažni uspeh" lanac [Wave 1 · M]
**Fajlovi:** `modStornoFlow.bas`, `modDokumentInvariant.bas`, `modStorno.bas`.
**Obim:** context-guard u 5 nezaštićenih grana; provera svakog paletnog detach
rezultata; provera zbirna relink count-a; invariant existence (`0=0` ne sme proći za
nepostojeću zbirnu); sum-scan error flag; `LookupActiveID` multi-match detekcija;
`StornoNovac` → čitaj `OtkupID` + `UpdateOtkupStatus` (+ snapshot `tblOtkup`)
(AUD-020, AUD-021; FM-0013 #1/#2/#3/#7, FM-0015 #1/#3, FM-0012 #12, FM-0019 #16).
**Regression:** `RunStornoTestSuite` (modTestStorno), `RunBusinessFlowProSuite`.

### RF-04 — modAutoHladnjaca lanac [Wave 1 · S/M]
**Fajlovi:** `modAutoHladnjaca.bas`.
**Obim:** `outBrPrij` postaviti tek posle uspešnog `SavePrijemnica_TX`; backfill
brojeve seedovati postojećim prijemnicama zbirne; proveriti `otpID` i `SaveZbirna_TX`
rezultat (prekid klase + warning); propagirati link grešku u warning lanca; `have`
dopuna posle uspeha (AUD-005; FM-0010 #1/#2/#4/#5/#6/#9). **Regression:** ručni
malina tok + `RunBusinessFlowProSuite`.

### RF-05 — frmDokumenta unos + storno set [Wave 1/2 · M]
**Fajlovi:** `frmDokumenta.frm`, `modDokumenta.bas`.
**Obim:** novac storno broj→`NovacID` rezolucija (kao za fakturu na :3222);
`FillOpenFakture` storno filter; Kl.II checkbox hard blokada kad izvor ima Kl.II;
obavezan smer ambalaže uz `kolAmb>0`; malina auto-zbirna vidljiv pad; prefill
poslednja generacija; `SumByBroj` storno filter; `SaveZbirna` presence guard
(AUD-008/009/022 + delovi AUD-003; FM-0018 #1-#5/#8/#20, FM-0011 #3).
**Regression:** `RunStornoTestSuite` + ručni unos otpremnica/zbirna/prijemnica.

### RF-06 — modIzvestaj ispravnost brojki [Wave 2 · M/L]
**Fajlovi:** `modIzvestaj.bas` (+ `modNovac.bas` per-vrsta cache).
**Obim:** isplate kooperanta vezati za stanicu (`COL_NOV_OM_ID`); kartice red
„Početno stanje"; „nema prijema" oznaka umesto 0%/100% nekonzistentnosti; kupac
per-vrsta raspodela uplate srazmerno stavkama; eksplicitni `Select Case` u Report*
dispatch-u (Else → Empty/greška) (AUD-023; FM-0028 #1/#3/#5/#6/#13/#14 + P2 stavke
#2/#9/#10/#11 po proceni sesije). **Regression:** `RunIzvestajTests` + uporedni
pregled izveštaja pre/posle na istim podacima (checklista mora dati očekivane razlike).

### RF-07 — frmIzvestaj freshness + revers [Wave 2 · M]
**Fajlovi:** `frmIzvestaj.frm`, `modKarticaDetalji.bas`.
**Obim:** status/štampa iz `m_curOd/m_curDo` (+ „nije osveženo" na izmenu datuma);
`CleanFail` čisti listu + vidljiva greška; zbirni tabovi 5/6/7 samo za validne tipove;
`StampajReversAmbDok`: `ExcludeStornirano` + tip ambalaže u match; `KarticaDetalji_Clear`
na promenu taba (AUD-024, AUD-012, deo AUD-027; FM-0029 #1-#5/#14/#15, FM-0030 #1).
**Regression:** ručni prolaz kroz tabove + revers štampa za slučaj sa storniranim redom.

### RF-08 — modFaktura + faktura štampa [Wave 2 · S/M]
**Fajlovi:** `modFaktura.bas`, `modPrint.bas`.
**Obim:** `CreateFaktura`: poređenje kupca prijemnice + `rows.Count=1` guard +
`CreateFaktura` na Private; `FillFakturaSablon` cleanup `.UnMerge`; blokada reprint-a
storniranog otkupa (ukinuti raw fallback :334) (AUD-011, AUD-027; FM-0034 #1/#2/#3,
FM-0031 #3/#19). **Regression:** `RunFakturaSmokeSuite`; ručno: faktura sa 3 stavke
posle fakture sa 1 stavkom (merge test).

### RF-09 — Banka import + mapiranje [Wave 2 · M]
**Fajlovi:** `modBankaImport.bas`, `modBankaMapiranje.bas`, `frmBankaImport.frm`.
**Obim:** dedupe ključ + broj računa; 3+ kandidata → `Err.Raise` (umesto subscript
pada); smer guard u Map* funkcijama; Activate auto-map → vidljiv rezultat ili iza
dugmeta; preview/command isti izvor za blok; upozorenje „ručno kupac = avans"
(AUD-014, AUD-025; FM-0022 #1, FM-0023 #8/#14, FM-0024 #1/#2/#3). **Regression:**
`Test_BankParse` po banci + ručni uvoz test izvoda + mapiranje.

### RF-10 — Banka export pregled [Wave 2 · S/M]
**Fajlovi:** `frmBankaExportPregled.frm`, `modBankaExportPregled.bas`.
**Obim:** stale override clamp u `PruneStaleOverrides`; finalna saldo revalidacija u
`GenerisiNalogeCSV`; brojanje stvarno primenjenih avansa (oslanja se na RF-02 iznos)
(AUD-026; FM-0020 #1/#2, FM-0021 #1). **Regression:** ručno: override → izmena
podataka → reload → CSV mora odbiti preplatu.

### RF-11 — frmOtkup / blokovi / kooperant [Wave 2 · M]
**Fajlovi:** `frmOtkup.frm`, `modOtkupBlok.bas`, `modKooperant.bas`.
**Obim:** parcela poređenje (sorta, ne vrsta); date re-lock blokirajući; glasan pad
link-a bloka + „Izgubljeni" uključuje prazan link; storno filter u
`ExistingBlokCena`/`BuildFirstBlokCena`; kooperant free-text disambiguacija kod >1
pogotka (AUD-028; FM-0007 #2/#3/#5, FM-0009 #4/#5, FM-0008 #1). **Regression:**
`RunBusinessFlowProSuite` + ručni unos otkupa sa panelom blokova.

### RF-12 — Palete [Wave 2 · S/M]
**Fajlovi:** `frmPalete.frm`, `modPaletniList.bas`.
**Obim:** UI validacija „bar jedna vrsta pakovanja" (kutije ILI kese); `Preradjeno`
guard na ulazu Reassign/Detach; sledljivost lista označena kao „mogući izvori"
(AUD-029; FM-0017 #1, FM-0016 #1/#2). **Regression:** `RunPaleteTestSuite`
(modTestPalete).

### RF-13 — Infra lifecycle [Wave 1/2 · M]
**Fajlovi:** `clsTransaction.cls`, `modJournaling.bas`, `modDataAccess.bas`,
`modParse.bas`, `ThisWorkbook.doccls`, `modMain.bas`, `modConfig.bas`,
`modMonitoring.bas`, `modStanicaLock.bas`, `modSetup.bas`.
**Obim:** `RollbackTx` per-tabela trap + garantovan `CleanUp` + `Class_Terminate`;
journal today-vs-today + `UpdateCell` journal + storno-marker na rollback; `AppendRow`
fantomski red; `TryParseDateValue` opseg+round-trip; startup trio (Err capture,
fail-soft backup, `FlushNow` na close); `GetConfigValue` trim; `TBL_CONFIG` van
required listi; monitoring `ActiveWorkbook` fallback; `BulkPushPendingForStanica` →
`UpdateCell` (AUD-004/006/007/017/018). **Regression:** restart aplikacije + namerno
izazvan rollback (checklista daje korake).

### RF-14 — MasterSync / JSON [Wave 1 · M]
**Fajlovi:** `modGoogleSheets.bas`, `modMasterSync.bas`.
**Obim:** JSON čitanje (strip samo van navodnika, `\"`, `\uXXXX`/`\t` dekod);
`ImportOtkupFromPWA_TX` → alias `_Core(False)`; `FindVOZSheets` paginacija
(AUD-001/002 + AUD-018 deo). **Regression:** pun sync ciklus sa test vrednostima
koje sadrže zapete, navodnike i dijakritiku (checklista ih navodi).

### RF-15…RF-19 — Wave 3 konsolidacije [P2 · po jedna sesija]
RF-15 HTTP helperi → `modHttpUtils` + 4 MasterSync call site-a; RF-16
`modBankaParseUtils` (~10 helpera, parseri netaknuti); RF-17 `BrutoUNeto` (6 klonova)
+ selidba `SaveOMUlaz_TX` u `modDokumenta`; RF-18 `NzBlank` + case-insensitive
storno compare; RF-19 `modTestHarness` + `LastRunFailedCount()` + self-update
`files_count`. Detalji: ROADMAP §10.3.

### RF-20 — Wave 4 bezbednost/proces [koordinisano — planirati posebno]
PIN hash + JMBG (VBA+GAS/PWA, migracioni prozor); `saveParcelPolygon` (KI-001);
sređivanje dokumentacije (AR/CL verzije, `instructions/` istorijski, CLAUDE.md
reference); `VBA_SRC_PATH` iz LocalConfig. Detalji: ROADMAP §10.4.

---

## 3b. Paketi iz delte v85 (SEF, startup-auth, sync, self-update, cenovnik)

> Detalji: `docs/AUDIT_FM_TRIJAZA.md` DEO II + `KNOWN_ISSUES.md` §8.4 (AUD-030…039).
> Sidra su `f6313dc`/`a0bc9e2` — pre rada re-bazirati na svež `main` (v2.24.0+).

### RF-21 — SEF correctness [P0/P1 · M]
**Fajlovi:** `modSEFClient.bas`, `modSEFValidator.bas`, `modSEFMapper.bas`,
`modSEFPersistance.bas`.
**Obim:** (P0) 409 izdvojiti iz REJECTED → `apiStatus="CONFLICT"` → TECH_FAILED/manual
(AUD-030); stornirana faktura sendable — `Stornirano` guard u `ValidateFakturaForSEF` +
filter u `frmSEF` combu; qty/price odvojeni `XmlQuantity`/`XmlUnitPrice` (3+ dec);
DueDate ≥ IssueDate uz force-today; `HasSuccessfulSEFSubmission` EH → re-raise
(fail-closed); resubmit čisti stari `SEFDocumentId`; stavke: samo aktivne (Stornirano≠DA,
OsirocenoOd prazno) (AUD-031). **Regression:** `RunSEFTestSuite` + demo SEF: duplicate
submit (409), faktura sa >2 dec, stara faktura (DueDate), storno pa pokušaj slanja.

### RF-22 — SEF UX/lifecycle [P1 · M]
**Fajlovi:** `frmSEF.frm`, `modSEFService.bas`, `modSEFStatusSync.bas`, (+ `modSEFTests`).
**Obim:** posle send-a `frmSEF` bira poruku po `SEFWorkflowState` (ne bezuslovno „Faktura
poslata"); `Test_Cancel/Storno…_TX` iz servisa → `modSEFTests` (ili SEF_ENV guard);
blank/unknown status → `UNKNOWN_STATUS` + manual review (ne SENT); `cmbFaktura_Change` →
`ClearSEFInfo`; recovery/refresh vraćaju rezultat, `frmSEF` ga prikazuje; batch summary
(found/recovered/failed). (AUD-032). **Regression:** ručno slanje odbijene fakture (poruka),
recovery zaglavljenog SENDING.

### RF-23 — Startup + authorization [P1 · S/M]
**Fajlovi:** `ThisWorkbook.doccls`, `modTrial.bas`/`modLicense.bas`, `modMaticniLookups.bas`,
`modAdmin.bas`, `modPodesavanja.bas`, `frmOtkupAPP.frm`, `frmStammdaten.frm`.
**Obim:** `Workbook_Open` → `If AccessWasDenied() Then Exit Sub` (pre `STARTUP_SUCCESS`)
(AUD-034); `MozeAdministraciju` guard: proširiti u `MaticniMenu_OnClick` na Admin/Podešavanja
+ vrh `BuildAdminPanel`/`AdminPanel_OnClick`/`BuildConfigEditor`/`ShowConfigSheet` (AUD-033);
`btnBanka_Click` auth guard pre importa (obrazac iz `btnSyncPWA_Click`) (AUD-034); PasswordChar
za „secret" polja u Podešavanjima; signal pri plaintext PIN fallbacku. **Regression:** login kao
ne-admin → probati Matične → Admin panel mora biti blokiran; deny licence → app se zatvara.

### RF-24 — Self-update hardening [P1 · M]
**Fajlovi:** `modSelfUpdate.bas`, `modRelease.bas`, `modBuildGuard.bas`, `modBuildInfo` gen.
**Obim:** faza 2 pokriva i failed `.frm` (ili ga ne Remove-uje u fazi 1); manifest `files_count`
+ abort na nepotpun download; `PublishReleaseToDrive` guard: placeholder/`+dirty` deny +
disk↔workbook `BUILD_SHA` cross-check; `AssertBlankBuild` skenira plain-range logove
(`SETUP_LOG`, test logovi) (AUD-035, AUD-037). **Regression:** simulirati prekinut download;
publish sa dirty radnim direktorijumom mora pasti.

### RF-25 — Sync/IO hardening [P2 · M]
**Fajlovi:** `modGoogleSyncOrchestrator.bas`, `modGoogleSheets.bas`, `modDrive.bas`,
`modStammdatenSync.bas`.
**Obim:** `SetPWAMasterSyncLock` → RMW obrazac (ne full-tab overwrite koji briše
`STANICA_LOCK_*`); dva rename-a u jedan `batchUpdate` (atomski swap); `ReadSheetData`
EMPTY≠ERROR (ByRef ok); `DriveFindInFolder` error≠not-found; empty-source cloud-wipe guard;
samostalni Parcele export → geo pull gate. (AUD-038). **Regression:** pun sync sa praznim
lokalom (ne sme obrisati cloud), paralelni station lock očuvan.

### RF-26 — Cenovnik + E2E gate [P1/P2 · S/M]
**Fajlovi:** `frmOtkup.frm`, `frmDokumenta.frm`, `modCenovnik.bas`, `modE2EReleaseGate.bas`,
`modBusinessFlowProTests.bas`.
**Obim:** stale auto-cena — očistiti polje pre lookup-a, prazno na 0 (AUD-036); `GetVazecaCena`
`Optional asOfDate` + `Datum` u schema guardu + UI datum fallback → greška; E2E gate: Boolean
`Core` po suite-u, `E2E_Pass` samo na `m_Failed=0`, ILI deprecirati modul; environment guard u
shipped test suite-ovima (AUD-039). **Regression:** unos otkupa za dva proizvoda uzastopno
(cena se ne prenosi); E2E gate mora prijaviti FAIL kad suite padne.

---

## 3c. Paketi iz delte v142 (agrohemija, sync-integritet, dijagnostika, sledljivost)

> Detalji: `docs/AUDIT_FM_TRIJAZA.md` DEO III + `KNOWN_ISSUES.md` §8.6 (AUD-040…048).
> Sidro je `origin/main` v2.24.0 (`9fd7087`) — pre rada re-bazirati na svež `main`.

### RF-27 — Agrohemija cena + validacija [P1 · S/M]
**Fajlovi:** `frmAgrohemija.frm`, `modAgrohemija.bas`.
**Obim:** izlaz prosleđuje `m_KorpaIzlaz(i).cena` kao `overrideCena` u `SaveMagacin` (simetrično
sa ulazom `:843`) → knjižena cena = snapshot korpe (AUD-040); `modAgrohemija` zahteva validnu
cenu > 0 za realne artikle (osim `ART_POCETNI_DUG`) umesto tihog `Cena=0/Vrednost=0`; referencijalne
provere u `ValidateMagacinInput` (postoji/aktivan artikal/koop/parcela↔koop). **Regression:** izlaz
artikla čija master cena ≠ cena u korpi → `tblMagacin` red mora nositi cenu iz korpe; izlaz sa
nenumeričkom cenom mora pasti, ne upisati 0.

### RF-28 — MasterSync integritet delte [P1 · M] (koordinisati sa RF-14)
**Fajlovi:** `modMasterSync.bas`, `modBrojevi.bas`, `modMalina.bas`, `modAutoHladnjaca.bas`.
**Obim:** `GenerateBrojPrijemnice` EH → `""` (ne `1/ddmmyy`) (AUD-041); `GenerateBrojZbirne`
delegira `SuggestNextBroj(KIND_ZBR,…)` umesto row-count (AUD-041); `TryUpdateVozacID` vraća True
samo ako write uspe (`RequireUpdateCell`) (AUD-042); strict datum parse → `SyncError` (ne današnji)
na oba puta (AUD-042); poison spreadsheet — temp-ime/rename ili trash na fail header-a (AUD-042);
auto-otpremnica ključ + `VrstaVoca|SortaVoca|Cena|TipAmbalaze`; VOZ link membership + „prazno ILI
isto" guard (AUD-043); canonical `IsManagedStationMirror` + re-raise u Ensure EH (AUD-046).
**Napomena:** deli `modMasterSync` sa RF-14 (JSON/paginacija) — raditi u istoj sesiji ili strogo
serijski uz re-bazu. **Regression:** sync batch sa 2 ista `BrojZbirne`, nevalidnim datumom, i
stanicom bez `tblVozaci` para — svaki mora dati `SyncError`, ne tihi upis.

### RF-29 — Integritet/health dijagnostika [P1/P2 · M]
**Fajlovi:** `modIntegritet.bas`, `modProductionHealthCheck.bas`.
**Obim:** `WriteErr` uvodi `ErrorCount`; overlay/MsgBox prikazuje INCOMPLETE kad errors>0; typed
`IntegrityRunResult` (Empty ≠ PASS) (AUD-044); health SEF lista koristi `WF_SEF_*` konstante +
state matrica (ukloniti nepostojeći `SEF_CANCELLED`); parent OK uslovljen child delta-brojačima
(`:951`, `:928`) (AUD-047). **Regression:** namerno pokvaren red (schema drift) → integritet mora
prijaviti GRESKA + broj, ne „0 neusklađenih"; health sa child FAIL ne sme dati parent OK.

### RF-30 — Sledljivost trace + sitni lifecycle [P1/P2 · S/M]
**Fajlovi:** `modSledljivost.bas`, `frmSledljivost.frm`, `modStornoWarm.bas`.
**Obim:** `TraceByZbirna` direktan `BrojZbirne` prolaz kroz `tblOtkup` + `UCase$(Trim$())`+
`vbTextCompare` (isto kao auto-link); typed trace rezultat (`IsComplete`); forma/PDF oznaka
„NEPOTPUN TRACE" (AUD-045); `modStornoWarm` flag tek po uspehu `OnTime` (`m_warmScheduled`),
`LogErr` na cancel fail (AUD-048). **Regression:** otkup sa `BrojZbirne` a praznim `OtpremnicaID`
mora ući u trag; PDF nepotpunog traga mora biti obeležen.

> **Banka (RF-16):** v142 FM-0128..0132 potvrđuju deljene rupe parsera (datum shape-only,
> račun bez normalizacije, poziv preuzak, 0/0-sa-računom) → hrane `modBankaParseUtils` (RF-16):
> kalendarska validacija datuma, opc. mod-97 računa, prošireni poziv obrazac, `zad>0 Xor odo>0`.

## 4. Status

| Paket | Naziv | Status | Grana | Napomena |
|---|---|---|---|---|
| RF-01 | Brisanje balasta | ⬜ | — | |
| RF-02 | modNovac guardovi | ⬜ | — | |
| RF-03 | Storno lanac | ⬜ | — | |
| RF-04 | AutoHladnjaca | ⬜ | — | |
| RF-05 | frmDokumenta set | ⬜ | — | |
| RF-06 | modIzvestaj brojke | ⬜ | — | |
| RF-07 | frmIzvestaj + revers | ⬜ | — | |
| RF-08 | Faktura + štampa | ⬜ | — | |
| RF-09 | Banka import/map | ⬜ | — | |
| RF-10 | Banka export | ⬜ | — | |
| RF-11 | Otkup UI | ⬜ | — | |
| RF-12 | Palete | ⬜ | — | |
| RF-13 | Infra lifecycle | ⬜ | — | |
| RF-14 | MasterSync/JSON | ⬜ | — | **koordinisati sa RF-28 (isti fajl)** |
| RF-15–19 | Konsolidacije | ⬜ | — | posle RF-14; RF-16 hrani v142 banka delta |
| RF-20 | Bezbednost/proces | ⬜ | — | planirati posebno |
| RF-21 | SEF correctness | ⬜ | — | **sadrži jedini nov P0 (409)** |
| RF-22 | SEF UX/lifecycle | ⬜ | — | posle RF-21 |
| RF-23 | Startup + authorization | ⬜ | — | **P1 auth lanac + AccessWasDenied** |
| RF-24 | Self-update hardening | ⬜ | — | |
| RF-25 | Sync/IO hardening | ⬜ | — | |
| RF-26 | Cenovnik + E2E gate | ⬜ | — | |
| RF-27 | Agrohemija cena + validacija | ⬜ | — | **jeftin high-value (cena≠knjižena)** |
| RF-28 | MasterSync integritet delte | ⬜ | — | spojiti sa RF-14 (isti fajl) |
| RF-29 | Integritet/health dijagnostika | ⬜ | — | |
| RF-30 | Sledljivost trace + lifecycle | ⬜ | — | |

**Redosled:** RF-01 → RF-02 → RF-23 → RF-21 → RF-27 → RF-03* → RF-04* → RF-05 →
RF-14+RF-28 → RF-06 → RF-07 → RF-08 → RF-22 → RF-09 → RF-10 → RF-11 → RF-12 → RF-13 →
RF-26 → RF-29 → RF-30 → RF-24 → RF-25 → RF-15+ → RF-20.
- **RF-27 podignut napred:** agrohemija cena≠knjižena je P1 finansijski/audit nalaz sa jeftinim
  (S) fixom — najbolji odnos vrednost/rizik u v142 delti.
- **RF-14 + RF-28 zajedno:** oba diraju `modMasterSync` (JSON/paginacija vs generatori/writeback);
  raditi u jednoj sesiji da se izbegne dupla re-baza istog fajla.
- **RF-23 i RF-21 podignuti napred:** RF-23 nosi P0-klasu bezbednosti (auth lanac do Admin
  panela) i P1 `AccessWasDenied`; RF-21 nosi jedini nov P0 (SEF 409). Oba pre storna.
- **RF-03*/RF-04* (storno):** OBAVEZNO re-verifikovati protiv `origin/main` (v2.24.0, storno
  PR #134–137 su prepravili `modStornoFlow` +746 linija) — deo nalaza je možda već rešen.
- Redosled se sme menjati — pravilo je samo: serijski, jedan po jedan, re-baza na svež `main`.
