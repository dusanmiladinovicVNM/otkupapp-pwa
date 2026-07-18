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
| RF-14 | MasterSync/JSON | ⬜ | — | |
| RF-15–19 | Konsolidacije | ⬜ | — | posle RF-14 |
| RF-20 | Bezbednost/proces | ⬜ | — | planirati posebno |

**Redosled:** RF-01 → RF-02 → RF-03 → RF-04 → RF-05 → RF-14 → RF-06 → RF-07 →
RF-08 → RF-09 → RF-10 → RF-11 → RF-12 → RF-13 → RF-15+ (RF-14 ranije jer je P0;
RF-06/07 posle da bi se izveštaji testirali nad već očišćenim podacima). Redosled se
sme menjati — pravilo je samo: serijski, jedan po jedan.
