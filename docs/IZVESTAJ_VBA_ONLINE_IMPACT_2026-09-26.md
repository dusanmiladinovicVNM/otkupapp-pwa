# Izveštaj — uticaj VBA izmena na online (GAS + PWA) sistem

**Datum:** 2026-09-26 · **Opseg izmena:** rad od 17.06.2026. naovamo · **Repo:** `otkupapp-pwa`

Dvodelni izveštaj:
- **Deo A** — analiza VBA izmena iz perioda **17–25.06.2026** i njihovog uticaja na online stranu (rađeno nad granom `e999ac0`, koja je u međuvremenu spojena u `main`).
- **Deo B** — **re-verifikacija** svih nalaza protiv **aktuelnog `main`-a** (`1095fdd`, 26.09.2026), tri meseca i 221 commit kasnije.

> **Metodologija.** Izvor je git istorija i čitanje koda; VBA nije kompajliran ni izvršen
> (Linux sesija — `tools/run_vba.py` se ne izvršava, pa se ponašanje prijavljuje kao
> statički verifikovano, ne kao „zeleno"). Deo A je rađen `git diff`-om `af308a3` (16.06)
> → `e999ac0` (25.06). Za Deo B je klon plitak (dostupno do ~16.09 / PR #344), pa su junski
> objekti van git-a — re-verifikacija je rađena **čitanjem trenutnog koda**, ne diff-om.
> Tvrdnje tipa „online niko ne čita X" počivaju na `grep` (0 pogodaka) po `gas/` i `src/js/`.

---

## Rezime (TL;DR)

Od ~180 commit-a u junu, na online stranu je tada zaista uticalo tek nekoliko stvari jer
**granica sinhronizacije seče većinu VBA izmena**: `tblAmbalaza` (ledger) se ne sinhronizuje,
a nove kolone van fiksne sync-šeme ostaju lokalne. Do septembra je online tok
(**predaja / otpremnica / zbirna**) suštinski prepravljen, pa je dobar deo junskih nalaza
**prevaziđen ili pomeren**, uz pojavu novih online površina.

Neto stanje na aktuelnom main-u: **2 nalaza REŠENA**, **~10 i dalje VAŽE**, **3 PROMENJENA**,
**1 PREVAZIĐEN**, plus **1 nova regresija** (`sw.js` keš) i **veliki nov online tok**
(`OtkupiAll`/`OtkupiAllStavke` split, `OtpremniceAll`, `PredajaID`) kojeg u junu nije bilo.
`APP_VERSION`: `2.2.2 → 2.28.4`.

---

## Deo A — VBA izmene 17–25.06 i uticaj na GAS + PWA

### A0. Ključni princip — gde je granica sinhronizacije

Većina „velikih" VBA izmena **ne prelazi online** jer:
- **`tblAmbalaza` (ledger) se NE sinhronizuje** — ni GAS ni PWA ne čitaju sirov ledger niti
  računaju saldo iz `Smer`. Saldo koji ide online je **pred-izračunat u VBA** iz
  `tblOtkup.KolAmbalaze`/`tblPrijemnica.KolAmbalaze`.
- **`tblOtkup`/`tblZbirna` se uvoze iz PWA kroz FIKSNU šemu** (OTK/VOZ). Nove kolone van te
  šeme ostaju lokalne.
- **GAS i PWA čitaju po IMENU kolone/polja** (`sheetToArray`/`headerIndexMap`, `r.Polje`) —
  dodavanje kolone na kraj je non-breaking; **nema pozicionog rizika** na online čitačima.

### A1. Izmene koje ZAISTA prelaze na online stranu

| Izmena (jun) | Šta ide online | Uticaj | Rizik |
|---|---|---|---|
| **`ExportOtkupiAll` 22 → 30 kolona** | 8 novih kolona na kraj `OtkupiAll` (MgmtReports): `BrojZbirne, OtpremnicaID, PrijemnicaID, BrojPrijemnice, KupacID, DatumPrijema, Primljeno, TransportStatus` + `BuildPrijemnicaIndexByBrojZbirne` join i izvedeni `TransportStatus` | „Pali" postojeću PWA logiku (`dispecer.js`, `kupci.js`) koja je do tada dobijala prazno | Nizak; traži re-export iz Excela |
| **Klasa II dobija realan `KolAmbalaze`** | Na Klasa-II redovima `KolAmbalaze` više nije `0` | **Jedina promena VREDNOSTI** sinhronizovanog polja — online ambalažni zbirovi rastu | Srednji |
| **„S" prefiks na `BrojZbirne`** | Oblik `1/250625 → S1/250625`; regex `^S?\d+/\d{6}(-\d+)?$` | Bezbedno — GAS drži `@`/tekst, string-equality, bez `parseInt`; PWA tretira kao string | Nizak |
| **Monitoring build otisak** | `monitorPublic` + `buildSha/buildVersion/buildDate` (novi `modBuildInfo`) | `Events` tab +3 kol., `Fleet` agregat, satni `rebuildMonitoringFleet` | Nizak; traži GAS redeploy (pozicioni `appendRow`) |
| **Malina auto-lanac + config ključevi** | Auto-zbirne, shadow-vozač; dodatni redovi u `Config` tabu | Sve `IsMalinaMode()`-gated, idempotentno; čita se po imenu | Nizak |

### A2. Izmene koje NE prelaze online (lokalno u Excelu)

| Izmena | Zašto ostaje lokalno |
|---|---|
| Otkup dvojni upis ambalaže (`Stanica Ulaz` + `Kooperant Izlaz`) | Redovi idu u `tblAmbalaza` (ne sinhr.); saldo online se računa iz `KolAmbalaze` |
| Ambalaza entitet-relativna + flip `Smer`-a | Isti razlog — **znak saldo-a promenjen, ali samo u Excelu** |
| Nove `tblOtkup` kolone `KolAmbIzdata`, `BrutoKg`, `VremeUnosa` | Nisu u OTK sync-šemi ni u `OtkupiAll` izvozu |
| Prijemnica bruto mod (neto u `Kolicina` + `BrutoKg`) | `Kolicina` ostaje neto → iznosi/saldo online nepromenjeni |
| `tblCenovnik` append-only | 0 referenci u sync sloju → PWA cenu vuče iz `config` (**buduća divergencija**) |

### A3. Novi VBA → GAS endpoint klijenti (licenca / verzija / build)

Sva imena polja i `action` stringovi se poklapaju **1:1** sa GAS handlerima.

| VBA modul | action | endpoint (config) | auth | GAS handler | GAS resursi |
|---|---|---|---|---|---|
| `modLicense` | `checkLicense` | `LICENSE_ENDPOINT`→`MONITORING_ENDPOINT` | nema (javno; ključ+otisak) | `checkLicense(data)` | sheet `Licenses` + `adminCreateLicense`; SP `LICENSE_HASH_SALT`, `LICENSE_TOKEN_SECRET` (auto) |
| `modUpdateGate` | `checkVersion` | `MONITORING_ENDPOINT` | `monitoringSecret` | `handleCheckVersionPublic_` | SP `VERSION_MIN/LATEST/ENFORCE/MESSAGE`, `MONITORING_INGEST_SECRET` |
| `modMonitoring` | `monitorPublic` (+build polja) | `MONITORING_ENDPOINT` | `monitoringSecret` | `normalizeMonitoringEvent_`/`appendMonitoringEvent_` | Monitoring spreadsheet: `Events` (+3 kol.), `Fleet` tab |

- **`modLicense`** — 3 hardverska otiska (`MachineGuid`, SMBIOS UUID, volume serial), fuzzy 2/3,
  anti-rollback (`LICENSE_HWM`), latch posle prve aktivacije. **Fail-open** dizajn.
- **`modUpdateGate`** — poredi `APP_VERSION` (ne `BUILD_VERSION`); `enforce=YES` blokira, inače warn.
- **`modBuildInfo`** — konstante `BUILD_SHA/VERSION/DATE`, auto iz `tools/stamp-build` (`git describe`).
- **`modBuildGuard` / `modTrial`** — čisto lokalni, bez mreže; gate-uju startup.

---

## Deo B — Re-verifikacija protiv aktuelnog main-a (26.09.2026)

Junska grana `e999ac0` je spojena u `main` i obrisana. Od tada je stiglo **221 commit-a**, uz
veliki refaktor online toka (PR-ovi S1c / S5-3 / S5-4b / #390 / #392 / #344) i novu arhitekturu
(`schema/schema.json` kao kanon, `docs/DOMEN/`, JS test-harness).

### B1. Master tabela — status svakog nalaza

Legenda: 🟢 VAŽI · 🟡 PROMENJENO · ⚫ PREVAZIĐENO · 🔴 REGRESIJA/otvoreno · ✅ REŠENO

| # | Nalaz (jun) | Status | Stvarnost na main-u |
|---|---|---|---|
| **PWA / online frontend** | | | |
| P1 | Multi-tenant `config.js` (silo, 3 tenanta) | 🟢 VAŽI | Nepromenjen; isti URL-ovi/DB_NAME model; stale komentar „POPUNI PLACEHOLDER" (URL-ovi su ipak popunjeni) |
| P2 | Per-klijent redirekti + tenant badge | 🟢 VAŽI | Nepromenjeno (`/bucaijoca` `/venivo` `/bukovik` → `/?t=`) |
| P3 | `sw.js` bumpovan na `AgriX-v25` | 🔴 REGRESIJA | I dalje `v25` posle velikog PWA refaktora; 47 lokalnih `<script src>` bez cache-busta, **cache-first**; bajt-identičan sw.js → nema reinstalacije → **star keširan JS kod povratnih klijenata** |
| **GAS endpointi** | | | |
| G1–G4 | `checkLicense`, `checkVersion`, Monitoring `Events`+`Fleet`+trigger, `bootstrapAgriXFolderTree` | 🟢 VAŽI | Svi prisutni; dirani u PR #344 (16.09) — stack revidiran posle juna |
| **VBA sync** | | | |
| S1 | `ExportOtkupiAll` 22→30 (ravna tabela) | ⚫ PREVAZIĐENO | Refaktor S1c: **`OtkupiAll` (28-kol zaglavlje) + `OtkupiAllStavke` (8-kol stavke)**; `Klasa/Kolicina/Cena/KolAmbalaze` prešli u stavke; `PredajaID` dodat u zaglavlje |
| S2 | Master-sync lock nepromenjen | 🟢 VAŽI | `SetPWAMasterSyncLock` preseljen u `modGoogleSyncOrchestrator`, protokol identičan (`SyncControl`/`MASTER_SYNC_LOCK`, TTL 10 min); nov potrošač istog taba: `modStanicaLock` |
| S3/S4 | `modGoogleAuth`/`modGoogleSheets` payload isti; monitoring šalje build polja | 🟢 VAŽI | Nepromenjeno / prisutno |
| **VBA data-model** | | | |
| D1 | Klasa II dobija realan `KolAmbalaze` | 🟡 PROMENJENO | `KolAmbalaze` je sada **polje stavke**; online stiže **agregatom** kroz `SaldoOMDetail` (čita PWA `kooperanti.js`); per-stavka se online ne čita |
| D2 | „S" prefiks `BrojZbirne` (bezbedan string) | 🟢 VAŽI | `ApplyMirrorPrefix` + regex prisutni; GAS/PWA tretiraju kao string |
| D3 | `tblAmbalaza` ledger nije online | 🟢 VAŽI | Nema izvoznog taba; jedina ref je rollback snapshot pri uvozu |
| D4 | `tblCenovnik` nije sinhronizovan (gap) | 🔴 otvoreno | 0 referenci u sync/GAS; cenovna divergencija VBA↔PWA ostaje |
| D5 | `KolAmbIzdata`/`BrutoKg`/`VremeUnosa` nisu u sync šemi | 🟡 PROMENJENO | `KolAmbIzdata` → sada u `OtkupiAll` zaglavlju (sinhr.); `BrutoKg` → u `OtkupiAllStavke`; `VremeUnosa` → **kolona obrisana** (zamenjena `CreatedAt`/`SourceCreatedAt`) |
| **VBA licenca/verzija/build** | | | |
| L1/L2 | `checkLicense`/`checkVersion` ugovor | 🟢 VAŽI | Poklapa se 1:1; nijedan field/action mismatch |
| L3 | Secret name trap (`MONITORING_SECRET` vs `MONITORING_INGEST_SECRET`) | 🟡 PROMENJENO (rizik VAŽI) | Junska premisa korigovana: VBA čita **samo** `MONITORING_SECRET` (drugo ime je samo komentar). Rizik „tihi 401 fail-open" i dalje postoji, ali dobio komentar-mapiranje + dijagnostiku (`Monitor_Test`/`CheckServerLink`) |
| L4 | Dupli `modLicenceTests.bas` (Ambiguous name) | ✅ REŠENO | Uklonjen; ostao samo `modLicenseTests.bas` |
| L5–L7 | `modBuildInfo` placeholder / gate poredi `APP_VERSION`; `modBuildGuard`/`modTrial` lokalni; ops zahtevi | 🟢 VAŽI | Nepromenjeno (`APP_VERSION`=2.28.4; guard/trial bez mreže; ops funkcije prisutne) |

### B2. Novo od juna (menja online sliku)

1. **🔴 Split `OtkupiAll`/`OtkupiAllStavke` je izvezen, ali online POLUispotrošen (najveći nalaz).**
   VBA piše stavke-tab, ali **nijedan GAS `getMgmt*` ni PWA modul ne čita `OtkupiAllStavke`**
   (grep = 0). Zbog dedupa u `mergeOtkupRows_` (master pobeđuje), čim se red mastruje, per-stavka
   polja (`Klasa/Kolicina/Cena/KolAmbalaze`) **nestaju iz read-modela** → `otkup-pregled.js` za
   mastrovane redove daje default (`klasa='I'`, ostalo 0). Live (neuvezeni) redovi ih još nose.
   Potreban je **nov čitač `OtkupiAllStavke` (join `OtkupID`↔`ServerRecordID`)** na GAS i PWA strani.
2. **Nova online površina za vozača: `OtpremniceAll` + `OtpremniceAllStavke`** (S5-4b). Vozač više
   ne dobija otkupne redove po `Otkup.VozacID` (mrtva veza), nego svoje otpremnice kao dokument;
   GAS čita oba taba; cena se namerno **ne** izvozi vozaču.
3. **`PredajaID` — identitet utovara** (S5-3). Nova kolona kroz `tblOtpremnica` → `OtkupiAll` →
   PWA → GAS (`processPredajaRecord`); jedan klik = jedan utovar; sprečava da parcijalni sync
   razbije utovar na dve otpremnice.
4. **`OtkupiAll` zaglavlje je sada IZVEDENI read-model „tekuće istine"** (review #390) — ne čita
   više mrtve `Otkup.VozacID/BrojZbirne/OtpremnicaID`, nego računa iz kanonskog lanca
   (`TekucaPredajaOtkupa`); `TransportStatus` izveden, prijemnica spojena po `BrojZbirne`.
   Dispečer/stanice rade nad header poljima koja i dalje postoje → dobili tačnije stanje.
5. **`validateLicenseToken` postoji ali je DORMANT** — HMAC verifikacija tokena nije zakačena ni
   na jedan endpoint (samo self-test). Token je zasad samo klijentski keš za offline grace;
   server-strana validacija VBA poziva **još nije aktivirana**.
6. **`schema.json` NE vodi sync kolone.** Kanon pokriva samo lokalne tabele; izvozne liste
   (`OtkupiAllKolone`, `OtkStavkeKolone`) su i dalje **ručni nizovi**. Dodavanje online kolone je
   ručni posao na 4 mesta (VBA niz + `Select Case`, GAS čitač, PWA čitač).

### B3. Živi rizici / gapovi na aktuelnom main-u (rangirano)

1. **`OtkupiAllStavke` se izvozi ali ga niko ne čita** → per-stavka podaci se gube iz mastrovanog
   read-modela (tiho degradiranje). Proveriti je li namerno (nov predaja-tok zamenjuje stari
   pregled) ili gap.
2. **`sw.js` zastareo keš** — posle velikog PWA refaktora obavezno bumpovati `CACHE_NAME`/`sw.js`
   pre deploy-a, inače povratni klijenti voze junski JS uz svež `index.html`.
3. **Secret mismatch = tihi fail-open** — različite vrednosti `MONITORING_SECRET` (Excel) ↔
   `MONITORING_INGEST_SECRET` (GAS) gase min-version gate i gube monitoring bez greške; posle
   setup-a pustiti `RunSetupHealthCheck`/`Monitor_Test`.
4. **`tblCenovnik` cenovna divergencija** — i dalje otvoren gap.
5. **`validateLicenseToken` dormant + `LIC_ENDPOINT_PINNED` prazan** — ako je server-strana
   validacija / anti-repoint bila namera, nije aktivirano.

---

## Deo C — Deployment / ops obaveze (pre paljenja online funkcija)

1. **GAS Web App redeploy** sa `Code.gs` (`checkLicense`/`checkVersion`, license blok) i
   `Monitoring.gs` (version gate, Fleet, `Events` +3 kol.) — **pre `LICENSE_ENABLED=YES`**;
   stari deploy vraća „Unknown action" → license gate može da zatvori svesku nevezanoj mašini.
2. **`MONITORING_SECRET` (Excel `tblSEFConfig`) == `MONITORING_INGEST_SECRET` (GAS Script
   Property)** kao identičan string; neusklađenost = tih 401 fail-open.
3. **Sheet `Licenses` + bar jedna licenca** (`adminCreateLicense`); `LICENSE_HASH_SALT`/
   `LICENSE_TOKEN_SECRET` se auto-generišu, ali se nikad ne smeju menjati posle prve aktivacije.
4. **Re-export iz Excela** (`SyncPWAFullCycle`) da se `OtkupiAll`/`OtkupiAllStavke` osveže.
5. **`installMonitoringTriggers()`** jednom (hourly `rebuildMonitoringFleet` + Fleet tab).
6. **Release kroz `tools/stamp-build`** inače Fleet prikazuje `0.0.0-dev`.
7. **`VERSION_ENFORCE=YES` je kill-switch cele flote** — držati `NO` dok se rollout ne potvrdi.

---

## Prilog — ključne lokacije (aktuelni main)

- **Sync:** `src-vba/modStammdatenSync.bas` (`OtkupiAllKolone` ~980, `ExportOtkupiAll` ~1034,
  `OtkupiAllStavkeRedovi` ~697, `ExportSaldoOMDetail` ~1447, `TekucaPredajaOtkupa` ~1006),
  `src-vba/modMasterSync.bas` (`OtkStavkeKolone` ~724, `IsValidBrojZbirneFormat` ~4575),
  `src-vba/modGoogleSyncOrchestrator.bas` (`SetPWAMasterSyncLock` ~731).
- **Model:** `src-vba/modConfig.bas` (`COL_OKS_*`, `COL_OTK_KOL_AMB_IZDATA`, `APP_VERSION` r.13),
  `src-vba/modBrojevi.bas` (`ApplyMirrorPrefix` ~320), `src-vba/modCenovnik.bas`, `schema/schema.json`.
- **Licenca/verzija:** `src-vba/modLicense.bas`, `modUpdateGate.bas`, `modMonitoring.bas`,
  `modBuildInfo.bas`.
- **GAS:** `gas/Code.gs` (`doPost` rute ~893–917, licenca ~6416–6811, `mergeOtkupRows_` ~1595,
  `sheetToArray` ~3690), `gas/Monitoring.gs` (`checkVersion`/`monitorPublic` ~191–231, Fleet ~635–708).
- **PWA:** `src/js/config.js` (TENANTS), `src/js/features/otkup/otkup-pregled.js` (~181–184),
  `sw.js` (`CACHE_NAME`), `src/js/features/management/{dispecer,kupci,kooperanti}.js`.
