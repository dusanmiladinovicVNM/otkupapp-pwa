# AgriX — Desktop Setup Reference

> **Svrha.** Jedinstveni izvor istine za dva pitanja:
> 1) Šta je sve od fajlova, foldera i podešavanja neophodno da AgriX radi na desktopu **u punom obimu**.
> 2) Kompletna procedura **setup / prvo pokretanje / instalacija** za novog korisnika — polazeći od
>    „blanko" `.xlsm` koji već ima podešen GAS URL + Google ID (to se radi na dev mašini).
>
> **Verzija koda:** `modConfig.APP_VERSION = 2.8.7`. Dokument odražava stanje nakon merge-a grane
> `claude/bank-pdf-gmail-downloader-76qfq8` u `main` (PR #108) — uključuje živi server-link
> health-check, `SetupPopplerInteractive`, banka PDF Gmail downloader i UI dugmad u Podešavanjima.
>
> **Vezani dokumenti:** `install/AgriX_Onboarding_Vodic_Novi_Klijent_v2.md` (operativni vodič, 47
> sekcija), `docs/production-runbook-banka-import-setup.md` (banka lanac od nule),
> `docs/SELF_UPDATE.md`, `docs/RELEASE_PROCEDURE.md`, `docs/licenciranje-po-uredjaju.md`.

---

## 0. Mentalni model — 3 nivoa

Sistem se pušta u rad kroz tri odvojena nivoa; ne mešati ih:

| Nivo | Ko / gde | Rezultat |
|---|---|---|
| **1. Dev priprema „blanko master"** | Ti, na dev mašini (`release.sh` → Excel `ImportAllVBA`) | `AgriX.xlsm` = **sav kod + sve tabele (shema) + config placeholderi**, prazne transakcione tabele |
| **2. Windows priprema po mašini** | `install/Setup-AgriX.ps1` (kod klijenta, admin) | Folderi, kopiran/unblock-ovan `.xlsm`, Poppler, sertifikat, Trusted Location, shortcut |
| **3. Aplikaciona provera po mašini** | `SetupNewPC` unutar Excela (prvo otvaranje) | `tblLocalConfig` putanje + health-check → `APP_SETUP_COMPLETED=DA` |

**Ključna činjenica:** `SetupNewPC` **ne pravi domenske tabele** — samo ih proverava
(`CheckRequiredTablesForSetup`). Pravi jedino foldere + `tblLocalConfig` + `tblPoruke`. Zato „blanko
master" **mora već imati kompletnu shemu**. To čuva `AssertBlankBuild` (`modBuildGuard.bas`): sve
tabele postoje, ali transakcione (`Otkup/Novac/Fakture/Otpremnice/Prijemnice`) su prazne (šifarnici
smeju biti seed-ovani).

---

## 1. Šta „blanko master" mora da nosi (dev-priprema)

### 1a. VBA kod

**Strogo neophodni za boot do dashboard-a** (`frmOtkupAPP`):
`ThisWorkbook.doccls`, `modMain`, `modConfig`, `modDataAccess`, `modSetup`, `modPoruke`, `modHelpers`,
`modArrayUtils`, `modTheme`, `modWindow`, `modLogError`, `modJournaling`, `modMonitoring`,
`modSchemaGuard`, `modBuildInfo`, `clsTransaction`, `frmSplash`, `frmOtkupAPP`.

**Kapije pri startu** (moraju da postoje/compile-uju, ali su **opt-in + fail-open** — na svežoj mašini
prolaze bez blokade): `modLicense` (+ `modTrial`), `modUpdateGate`, `modSelfUpdate`, `modAuth`
(+ `frmLogin`).

**Feature moduli** (app se digne i bez njih, ali „pun obim" ih traži): otkup/dokumenta (`modOtkup`,
`modOtkupBlok`, `modDokumenta`, `modBrojevi`, `modStorno`, `modSledljivost`), agrohemija
(`modAgrohemija`), ambalaža/cenovnik (`modAmbalaza`, `modCenovnik`), palete/print (`modPaletniList*`,
`modPrint`, `modDocStyle`), banka (`modBankaImport`, `modBankaImportParserPdfToText`,
`modBankaMapiranje`, `modBankaExportPregled`), novac/faktura/marža/izveštaji, SEF (`modSEF*`), GIS
(`modGeoParcele`), matični (`frmStammdaten`, `modMaticniLookups`), sync/PWA (`modGoogleAuth`,
`modMasterSync`, `modStammdatenSync`, `modGoogleSheets`, `modDrive`, `modGoogleSyncOrchestrator`,
`modStanicaLock`).

**Build/dev/test — u repo-u, ali se ne isporučuju logikom** (self-update ih preskače): `modRelease`,
`modVbaTools`, `modBuildGuard`, `modE2EReleaseGate`, `mod*Tests`. Prazni Phase-2/3 stub-ovi:
`modKvalitet`, `modHladnjaca`, `modML`, `modMeteo`.

> Nema Excel ribbon/CustomUI XML — ceo UI je full-screen UserForm `frmOtkupAPP` koji modeless-om
> učitava feature forme u panel; kontrola prava kroz `OpenContentForm` → `modAuth`.

### 1b. Sve tabele (shema)

`SetupNewPC` proverava da postoje (pada u „NE" ako fale): `tblConfig`, `tblSEFConfig`,
`tblBankaImport`, `tblPartnerMap`, `tblKooperanti`, `tblStanice`, `tblVozaci`, `tblKupci`,
`tblKulture`, `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakture`,
`tblFakturaStavke`, `tblNovac`, `tblAmbalaza`, `tblPoruke` (+ `tblLocalConfig` se pravi sam).

Za pun obim, pri pripremi templata pokreni jednokratne `Ensure*` makroe na masteru (idempotentni,
schema-drift safe): `EnsurePaletniListSchema` (palete/prerada/tipovi/kutije/kese/VrstaGP +
`GajbicaPoPaleti` na kulture), `EnsureCenovnikSchema` (`tblCenovnik`), `EnsureKorisniciSchema`
(`tblKorisnici`), `EnsureDoradeSchema` (soft-delete `Aktivan`, `JeHladnjaca`, bruto/decimalni
format…), `EnsureAuditColumns` (Created/Modified na svim glavnim tabelama), `EnsurePoruke` (katalog
poruka — **obavezan**, ceo UI ide kroz `Poruka("KLJUC")`).

### 1c. Config placeholderi (u `tblSEFConfig`)

Ovo je „gas i google ID". Minimalno za cloud režim (v. sekciju 4 za pun spisak): `GOOGLE_CLIENT_ID`,
`GOOGLE_CLIENT_SECRET`, `GOOGLE_PWA_FOLDER_ID`, `MONITORING_ENDPOINT`, `MONITORING_SECRET`,
`MONITORING_ENV`, `CLOUD_SYNC_ENABLED=YES`.

> ⚠ **Google ide u `tblSEFConfig`, NE `tblConfig`** — v. Known issues (sekcija 10, drift #1).

---

## 2. Neophodni folderi i fajlovi na desktopu

Root aplikacije = folder u kome stoji `.xlsm` (`GetDefaultRootPath = ThisWorkbook.path`,
`modSetup.bas`). Sve se pravi **pored radne sveske**:

```
<APP_ROOT> (npr. C:\AgriX\)
├─ AgriX.xlsm
├─ Tools\poppler\Library\bin\pdftotext.exe     ← Poppler (default, APP_PDFTOTEXT_RELATIVE_EXE_PATH)
├─ Backups\  Logs\  Journal\  Export\  Temp\  Secrets\   ← EnsureAppFolders
├─ Bank_Izvodi\Inbox\  \Processed\  \Error\    ← SetupBankFolders (LOKALNI, ne Drive!)
└─ (PDF izlazi) Otkupni listovi\ Prijemnice\ Otpremnice\ Revers ambalaze\
   Kartice kooperanata\ Paletni listovi\ Preradni listovi\ Specifikacije\ Izvestaji\
```

---

## 3. Google / cloud strana (ti pripremaš)

**Dva odvojena GAS deploja (dva URL-a, ne mešati):**

1. **Glavni AgriX GAS** (`gas/Code.gs`+`Monitoring.gs`+`DriveFolder.gs`) — Web App `/exec`,
   „Execute as: Me". Desktop ga zove preko `MONITORING_ENDPOINT` (i licenca preko istog). Traži
   Script Properties: **svih ~33 `AGRIX_*_FOLDER_ID`** (pokreni `bootstrapAgriXFolderTree()` u tom
   projektu), `MONITORING_INGEST_SECRET` (= desktop `MONITORING_SECRET`), version props
   (`VERSION_MIN/LATEST/ENFORCE`), license secrets (`LICENSE_HASH_SALT`, `LICENSE_TOKEN_SECRET`).
2. **Bank PDF Gmail Downloader** (`gas/bank-pdf-downloader/`) — na nalogu koji **prima mejlove banke**.
   Traži `BANK_IMPORT_CLIENTS_JSON` (sa ID-em `01_Bank` foldera + `bankSenders`), Gmail+Drive scope,
   Editor share `01_Bank`, dnevni trigger 07h. (Detalji: sekcija 6, korak 11.)

**Drive stablo (izvor istine = `DriveFolder.gs`):** jedno po klijentu, `AgriX_C00X_PROD/`:

```
AgriX_C00X_PROD                       → AGRIX_ROOT_FOLDER_ID
├─ 00_Inbox                           → AGRIX_INBOX_FOLDER_ID
│  ├─ 01_Bank        (banka PDF; = BANKA_DRIVE_SOURCE_PATH; NEMA zaseban prop)
│  ├─ Downloaded     (puller premešta povučene PDF-ove; NEMA prop)
│  ├─ Fiskalni  Uvoz  Manual
├─ 01_Sheets                          → AGRIX_SHEETS_FOLDER_ID
│  ├─ 01_Operational / 02_Master (Stammdaten) / 03_Reports (MgmtReports, Kartice) / 04_Archive
├─ 03_Documents  04_Export  05_Backup  06_Monitoring (ErrorLog)  07_Admin
```

Plus dva globalna, share-ovana foldera zakucana u `modConfig`: `REL_FOLDER_ID`
(`1zL7ronXQUsOY56p7rULsqrM1u1U8sxod`, AgriX_Release — self-update) i `BACKUP_FOLDER_ID`
(`199is7nQW3d4wfGX974AFTjpS4wo8itpl`, AgriX_Backup). **Stammdaten** sheet živi u `01_Sheets/02_Master`;
razrešava se **po imenu**, ne po zakucanom ID-u.

---

## 4. Podešavanja — dve tabele, ne mešati store

| | **`tblSEFConfig`** (GLOBAL, putuje sa `.xlsm`) | **`tblLocalConfig`** (PER-MAŠINA, ne putuje) |
|---|---|---|
| Kolone | `ConfigKey` / `ConfigValue` | `Kljuc` / `Vrednost` / `Opis` |
| Čita/piše | `GetConfigValue`/`SetConfigValue` (`modConfig`) | `GetLocalConfigValue`/`SetLocalConfigValue` (`modSetup`) |
| Editor | Matični podaci → Podešavanja (default store „sef") | ista forma, grupa „Banka / lokalno" (store „local") |
| Vidljivost | VeryHidden čim setup prođe zeleno; izlaz `Alt+F8 → ShowConfigSheet` | vidljiva radna tabela |

**`tblSEFConfig` — cloud/global (obavezno za cloud režim):** `GOOGLE_CLIENT_ID`,
`GOOGLE_CLIENT_SECRET`, `GOOGLE_PWA_FOLDER_ID`, `MONITORING_ENDPOINT`, `MONITORING_SECRET`,
`MONITORING_ENV`, `CLOUD_SYNC_ENABLED`. Auto-popunjeni posle 1. sync-a: `GOOGLE_REFRESH_TOKEN`,
`GOOGLE_STAMMDATEN_SHEET_ID`. Opciono: `GOOGLE_REPORTS_FOLDER_ID`, `GOOGLE_KARTICE_SHEET_ID`,
`GOOGLE_MGMT_SHEET_ID`, `SYNC_AUTO_INTERVAL_MIN`. Licenca: `LICENSE_ENABLED`, `LICENSE_ENDPOINT`
(fallback na `MONITORING_ENDPOINT`), `LICENSE_KEY`. SEF (opciono): `SEF_BASE_URL`, `SEF_API_KEY`,
`SEF_ENV`. Firma: `SELLER_*`. Feature flag-ovi: `MALINA_MODE`, `PRACENJE_PARCELA`, `PALETIRANJE`,
`AUTH_ENABLED`, `*_PRINT_MODE`, „Kompletna validacija unosa" (v2.8.6)…

**`tblLocalConfig` — per-mašina (default-e postavlja `SetupNewPC`):** `APP_SETUP_COMPLETED`,
`APP_ROOT_PATH` (+ `APP_BACKUP/LOG/JOURNAL/EXPORT/TEMP/SECRETS_PATH`),
`BANKA_INBOX/PROCESSED/ERROR_PATH`, `BANKA_DRIVE_SOURCE_PATH` (lokalna putanja Drive `01_Bank` —
**jedini ključ koji aktivira banka pull**), `BANKA_DRIVE_DOWNLOADED_PATH`, `BANKA_DRIVE_MAX_FILES`
(def 50), `BANKA_AUTO_IMPORT_ON_START` (def NE), `BANKA_ALLOWED_EXTENSIONS` (def pdf),
`PDFTOTEXT_EXE_PATH` (prazno = auto, `Tools\poppler` pored sveske).

---

## 5. Boot lanac i kapije (`StartApp`)

`Workbook_Open` (`ThisWorkbook.doccls`) → `StartApp` (`modMain.bas`). Redosled:

| # | Korak | Uslov / ponašanje |
|---|---|---|
| 1 | `InitApp` → `EnsurePoruke` + `ValidateAllTables` | Ako fale tabele: samo upozorenje „Pokrenite Setup" (**ne blokira**) |
| 2 | **Licenca** `AccessGateOrQuit` | Opt-in `LICENSE_ENABLED=YES` (ili već aktivirana mašina). **Fail-open** |
| 3 | **Min-verzija** `UpdateGateOrQuit` | Opt-in `MONITORING_ENDPOINT`+`SECRET`; GAS `checkVersion`. Fail-open offline |
| 4 | **Self-update** `CheckForUpdateOnOpen` | Opt-in `REL_FOLDER_ID`. Nova verzija → reimport + Exit |
| 5 | **Prijava** `modAuth.Login` | Opt-in `AUTH_ENABLED=YES` |
| 6 | **First-run kapija** | Ako `APP_SETUP_COMPLETED ≠ "DA"` (tblLocalConfig) → ponudi `SetupNewPC`. Fail-soft, jednom |
| 7 | `Visible=False` → `frmSplash` → backup/journal/purge → `RecoverAllStuckSEF` → `StartScheduledSync` → `frmOtkupAPP` | |

Sve kapije (2–5) su opt-in i fail-open: sveža mašina sa samo GAS+Google ih prolazi bez blokade.

---

## 6. Detaljna lista koraka (setup / prvo pokretanje / instalacija)

Legenda aktera: **[DEV]** ti na dev mašini · **[KLIJENT-GOOGLE]** ti, jednom po klijentu u Google-u ·
**[MAŠINA]** kod klijenta po računaru.

### FAZA 0 — „Blanko master" `.xlsm` [DEV]

1. `bash tools/release.sh 2.8.7` (bump `APP_VERSION`, tag `vba-v2.8.7`, push).
2. Excel (master): `Alt+F8 → ImportAllVBA` → `Debug → Compile VBAProject` (bez greške).
3. Jednokratni schema makroi (ako već nisu u masteru): `EnsurePaletniListSchema`,
   `EnsureCenovnikSchema`, `EnsureKorisniciSchema`, `EnsureDoradeSchema`, `EnsureAuditColumns`,
   `EnsurePoruke`.
4. `Alt+F8 → AssertBlankBuild` → mora **„BLANKO OK"** (transakcione tabele prazne).
5. `Alt+F8 → PublishReleaseToDrive` (kod + `version.json` u `AgriX_Release`) — za self-update flote.
6. (Opciono) VBE `Tools → Digital Signature` (potpiši), `Ctrl+S`, vrati `modBuildInfo` placeholder.
7. Verifikuj da master nosi: sav kod, **sve tabele**, i placeholdere u `tblSEFConfig`.

### FAZA 1 — Google okruženje po klijentu [KLIJENT-GOOGLE]

8. **Glavni GAS #2** (nalog ops/klijent): postavi `ROOT_FOLDER_ID` (ručno napravljen `AgriX_C00X_PROD`)
   → `Run → bootstrapAgriXFolderTree()` → pravi stablo i upisuje **33 `AGRIX_*_FOLDER_ID`** u Script
   Properties. Time nastaju i `00_Inbox/01_Bank` i `00_Inbox/Downloaded` (**bez zasebnog propa**, po
   dizajnu — `DriveFolder.gs`).
9. Script Properties glavnog GAS-a: `MONITORING_INGEST_SECRET` (≥16), `VERSION_MIN/LATEST/ENFORCE`,
   `LICENSE_HASH_SALT` + `LICENSE_TOKEN_SECRET`. Deploy Web App `/exec` („Execute as: Me").
10. Google Cloud: OAuth **Desktop** client → `GOOGLE_CLIENT_ID` + `GOOGLE_CLIENT_SECRET`.

11. **POVEZIVANJE DVA GAS-a (banka) — pivot je deljeni `01_Bank` folder.**
    Dva GAS-a se **ne povezuju API-jem** — povezuju se **istim `01_Bank` folderom**; isti folder ID se
    pojavljuje na tri mesta:
    - **11a.** Iz Drive-a **iskopiraj folder ID od `01_Bank`** (iz URL-a; nema ga u Script Properties
      glavnog GAS-a, po dizajnu).
    - **11b.** **Podeli `01_Bank` kao Editor** nalogu koji **prima mejlove banke** (na kom radi
      downloader GAS #1) → tek tada #1 sme da upisuje PDF-ove.
    - **11c.** **Instaliraj downloader GAS #1** (`AgriX_C00X_BankPdfDownloader`,
      `gas/bank-pdf-downloader/`): nalepi `Code.gs` + `appsscript.json` (V8, scopes:
      `mail.google.com`, `drive`, `script.scriptapp`).
    - **11d.** Script Properties GAS #1: `BANK_IMPORT_CLIENTS_JSON` sa `driveFolderId` = **isti
      `01_Bank` ID iz 11a** (+ `bankSenders`, `fileNamePrefix=C00X`, `searchDays`…). Time je „veza"
      gotova: #1 piše u folder koji živi u stablu #2.
    - **11e.** GAS #1: `Run → testGmailAccessOnly` (odobri Gmail scope) → `testBankPdfImportConfig` →
      `runBankPdfImportNow` (PDF-ovi padnu u `01_Bank`) → `setupDailyBankPdfImportTrigger` (07h).
      Istorija: `BANK_IMPORT_BACKFILL_FROM_DATE` + `runBankPdfImportBackfill`.

12. **Napuni `tblSEFConfig` u masteru** (⚠ NE `tblConfig`): `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`,
    `GOOGLE_PWA_FOLDER_ID` (= ID `01_Sheets/02_Master`), `MONITORING_ENDPOINT` (GAS `/exec`),
    `MONITORING_SECRET`, `MONITORING_ENV=PROD`, `CLOUD_SYNC_ENABLED=YES`. Licenca/SEF opciono.

### FAZA 2 — Install paket [DEV]

13. Spakuj (onboarding §25): `app/AgriX.xlsm` (potpisan, BLANKO OK) · `install/Setup-AgriX.ps1`
    · `tools/poppler/…/pdftotext.exe` (+ DLL) · `cert/AgriX-VBA-Publisher.cer` (samo javni `.cer`)
    · `docs/` · `manifest.json` + `checksums.sha256`.

### FAZA 3 — Windows instalacija po mašini [MAŠINA]

14. `Setup-AgriX.ps1` (kao admin): pravi `C:\AgriX` + podfoldere → kopira + **`Unblock-File`**
    xlsm → kopira `Tools\poppler` **pored** xlsm → import `.cer` u `CurrentUser\Root`+`TrustedPublisher`
    → **Excel Trusted Location** (`AllowSubfolders=1`) → Desktop shortcut.

### FAZA 4 — Drive for Desktop po mašini (banka transport cloud→disk) [MAŠINA]

15. Instaliraj **Google Drive for Desktop**, uloguj na nalog koji **vidi `01_Bank`** — vlasnički
    (ops/klijent) ili email-nalog kome je deljen (tada: Drive web → desni klik `01_Bank` → **Add
    shortcut to Drive → My Drive**, uđe pod `H:\My Drive\…`).
16. `00_Inbox` → desni klik → **Available offline** (da `pdftotext` čita realne bajtove, ne cloud
    placeholder).
17. Zapamti lokalnu putanju (npr. `H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank`) — upisuješ je u
    koraku 22.

### FAZA 5 — Prvo otvaranje + `SetupNewPC` [MAŠINA]

18. Otvori xlsm (shortcut). Boot prolazi kapije (sekcija 5) → **first-run kapija** ponudi `SetupNewPC`.
19. `SetupNewPC` (redosled): `EnsureAppFolders` (+ PDF podfolderi) → `SetupBankFolders` (lokalni
    `Bank_Izvodi\{Inbox,Processed,Error}`) → `CheckGoogleOAuthConfig` (iz `tblSEFConfig`; preskače se
    ako `CLOUD_SYNC_ENABLED=NO`) → `CheckSEFConfigForSetup` (SEF opcion) → provera tabela/kolona.
20. **Živi server-link (`CheckServerLink`) je ADVISORY:** proverava Google OAuth token +
    `DriveListFolder(GOOGLE_PWA_FOLDER_ID)` + `Monitor_Test` (ako `MONITORING_ENDPOINT`) +
    `BANKA_DRIVE_SOURCE_PATH` postoji. **Prikazuje NAPOMENU, ne obara „zeleno".** Zeleno →
    `APP_SETUP_COMPLETED=DA` + `HideConfigSheet` (sakrije `tblSEFConfig`).
21. **Poppler** (jedan od): `Alt+F8 → SetupPopplerInteractive` (ako je `Tools\poppler` pored xlsm →
    auto režim `PDFTOTEXT_EXE_PATH=""`; inače picker + `FindPdfToTextExe`) — ili Matični podaci →
    **Podešavanja → „Izaberi Poppler (pdftotext.exe)"** (bez Alt+F8).
22. **Upiši `BANKA_DRIVE_SOURCE_PATH`** = lokalna putanja iz koraka 17. Kroz **Podešavanja → grupa
    „Banka / lokalno" → inline „…" browse dugme** (nema više potrebe za Immediate).
23. Brza provera veze: `Alt+F8 → TestServerLink` (Google / GAS / banka Drive folder).

### FAZA 6 — Banka: šifarnici → uvoz → verifikacija [MAŠINA]

24. Preduslovi za AUTO-map: kooperanti/kupci imaju **`TekuciRacun`**; kooperanti imaju **`StanicaID`**
    (bez stanice se isplata ne knjiži).
25. Meni **„Banka uvoz izvoda"** (`ImportBankaInbox_WithDrivePull`): pull `01_Bank`→lokalni Inbox
    (original u `Downloaded`) → `pdftotext` → `tblBankaImport`; forma auto-mapira po jakim ključevima
    (poziv na broj / tekući račun) → dovrši ručno/„Auto sve" → `tblNovac`.
26. Verifikacija lanca: `BIM:<id>` trag u `tblNovac.Napomena`; par knjiženja. Dnevna rutina: GAS #1 07h
    → operater otvori uvoz → auto-map → provera brojača.

**Ceo banka lanac (posle 11+15):** Banka(email) → GAS #1 (Editor na `01_Bank`) → Drive `01_Bank`
(u stablu #2) → Drive for Desktop → lokalni `…\01_Bank` → VBA puller
(`PullBankPdfsFromDriveProduction`) → `Bank_Izvodi\Inbox` → `pdftotext` → `tblBankaImport` →
`tblNovac`.

### Minimalno za NOVOG korisnika

- **Nov klijent** (novo Google okruženje): koraci **8–12** (uklj. povezivanje dva GAS-a 11a–11e), pa
  **14–26**.
- **Nova mašina istog klijenta**: preskačeš Fazu 1; radiš **14–23** (+ 24–26 za banku).
- Za banku uvek treba **15–17 + 22** (Drive for Desktop + `BANKA_DRIVE_SOURCE_PATH`).

---

## 7. Install package (trenutno stanje)

**Verzionisano u `install/`:** `Setup-AgriX.ps1`, `AgriX_Onboarding_Vodic_Novi_Klijent_v2.md`,
`Priprema_pre_instalacije.txt`. **Build-artefakti (NISU u repo-u):** potpisan `AgriX.xlsm`,
`Tools\poppler\`, `AgriX-VBA-Publisher.cer`. Sklopljena struktura: onboarding §25.

**`Setup-AgriX.ps1` radi:** param `InstallRoot=C:\AgriX`, `ExcelVersion=16.0` → folderi
(`Backups/Logs/Journal/Export/Temp/Secrets` + `Bank_Izvodi\{Inbox,Processed,Error}`) → kopira xlsm
(throw ako fali) → `Unblock-File` → kopira `Tools\` (Poppler) i `docs\` (non-fatal) → import `.cer` u
`CurrentUser\Root`+`TrustedPublisher` → Trusted Location (`AllowSubfolders=1`) → Desktop shortcut →
piše `Logs\install-log.txt` + **PASS/FAIL** summary → osveženi next-steps (SetupNewPC,
SetupPopplerInteractive, Drive-for-Desktop / `BANKA_DRIVE_SOURCE_PATH`, `TestServerLink`).

**Konzistentno sa kodom:** `InstallRoot` = mesto xlsm-a → `APP_ROOT_PATH` = `C:\AgriX`; Poppler
layout = `APP_PDFTOTEXT_RELATIVE_EXE_PATH`; `Bank_Izvodi\{…}` = LOKALNI processed folderi (Drive izvor
`01_Bank` ide preko Drive for Desktop, ps1 ga s pravom ne dira).

**Preostalo (build-time, ne installer):** `manifest.json` / `checksums.sha256` (onboarding §24A.12/§25)
se generišu pri pakovanju paketa, ne u `Setup-AgriX.ps1`. Install-log, PASS/FAIL summary, `docs\`
kopija i osveženi next-steps su dodati u ovoj grani (v. §10).

---

## 8. Eksterne zavisnosti

- **Poppler `pdftotext.exe`** — samo za banka import. Razrešavanje (`ResolvePdfToTextExePath`):
  eksplicitni `PDFTOTEXT_EXE_PATH` → default `<sveska>\Tools\poppler\Library\bin\pdftotext.exe` → goli
  `pdftotext.exe` iz PATH-a. Parser je za **Komercijalnu banku**.
- **SEF / e-faktura** — potpuno **opciono**; runtime tvrdo staje ako fali ključ, setup preskače ako
  nije podešeno.
- **Office/Windows** — Excel 16.0 (2016+/M365), VBA7. Makroi kroz **potpis + Trusted Location** (NE
  globalno omogućavanje; „Trust access to VBA project" se izričito NE uključuje). COM (standardni):
  `WinHttp`, `MSXML2.ServerXMLHTTP`, `WScript.Shell`, `Scripting.FileSystemObject`, `ADODB.Stream`,
  `VBScript.RegExp`, WMI.

---

## 9. Licenciranje (per-uređaj, opt-in)

Sveža instalacija sa samo GAS+Google **radi bez licence** (fail-open). Za enforce: server-side
`adminCreateLicense`, u svesci `LICENSE_ENABLED=YES` + `LICENSE_ENDPOINT`, pa na mašini
`Alt+F8 → ActivateLicensePrompt` (nalepi ključ, treba internet prvi put). Dijagnostika:
`LicenseShowDevice`. Node-locked (fingerprint 2-od-3: MachineGuid/SMBIOS UUID/volume serial), radi
offline do `LICENSE_NEXT_CHECK`. Detalji: `docs/licenciranje-po-uredjaju.md`.

---

## 10. Poznati driftovi dokumentacije (known issues)

> Merodavan je kod (grana `bank-pdf-gmail-downloader`). Stavke 1, 2, 5, 6 su **usklađene u ovoj grani**
> (`install/` fajlovi ažurirani); ostaju zabeležene radi traga.

1. **[REŠENO u grani] Google OAuth lokacija.** `Priprema_pre_instalacije.txt` i onboarding
   §9/§15/§24A.11 su ranije upućivali na `tblConfig`; ispravljeno na **`tblSEFConfig`** (kod:
   `modSetup.CheckGoogleOAuthConfig` + runtime `modGoogleAuth` preko `GetConfigValue`). UI Podešavanja
   ionako piše u `tblSEFConfig`.
2. **[REŠENO u grani] Stari `AGRIX_BANK_IZVODI_FOLDER_ID`.** Izbačen iz onboarding §8/§11; aktuelno je
   `00_Inbox/01_Bank` (+ `Downloaded`), koji nema Script Property i čiji se ID uzima ručno (za Editor
   share + `BANK_IMPORT_CLIENTS_JSON.driveFolderId` + `BANKA_DRIVE_SOURCE_PATH`).
3. **`PDFTOTEXT_EXE_PATH` store.** Ide u `tblLocalConfig` (ne `tblSEFConfig`). Stariji build je pisao u
   SEFConfig → tiho ne radi; rešeno grupom „Banka/lokalno".
4. **Dve `APP_VERSION`.** `modConfig.APP_VERSION=2.8.7` (verzija koda za version-gate/self-update) ≠
   `tblSEFConfig.APP_VERSION` (npr. `1.0.0-C00X`, monitoring/fleet tag). Ne poistovećivati.
5. **[REŠENO u grani] ps1 §24A.10 spec.** `Setup-AgriX.ps1` sada piše `Logs\install-log.txt`,
   ispisuje PASS/FAIL summary, kopira `docs\`, i ima osvežene next-steps (SetupPopplerInteractive,
   Podešavanja „Izaberi Poppler"/„…", `TestServerLink`, povezivanje dva GAS-a).
6. **[REŠENO u grani] Povezivanje dva GAS-a + Drive for Desktop** dodato u onboarding §27/§39 kao
   prvorazredni korak (uz postojeći `docs/production-runbook-banka-import-setup.md`).

---

## 11. Reference (fajlovi / funkcije)

- Boot: `src-vba/ThisWorkbook.doccls` (`Workbook_Open`), `src-vba/modMain.bas` (`StartApp`,
  `InitApp`, `ValidateAllTables`).
- Setup: `src-vba/modSetup.bas` (`SetupNewPC`, `RunSetupHealthCheck`, `CheckServerLink`,
  `TestServerLink`, `SetupPopplerInteractive`, `EnsureAppFolders`, `SetupBankFolders`,
  `EnsureDataTable`, tblLocalConfig get/set).
- Config: `src-vba/modConfig.bas` (`APP_VERSION`, `REL_FOLDER_ID`, `BACKUP_FOLDER_ID`,
  `APP_PDFTOTEXT_RELATIVE_EXE_PATH`, `GetConfigValue`/`SetConfigValue`, `IsCloudSyncEnabled`);
  `src-vba/modPodesavanja.bas` (editor, `ConfigEditor_PickPoppler`, `ConfigEditor_PickFolderInto`,
  `HideConfigSheet`/`ShowConfigSheet`).
- Blank-build: `src-vba/modBuildGuard.bas` (`AssertBlankBuild`).
- Licenca/update: `src-vba/modLicense.bas`, `modTrial.bas`, `modUpdateGate.bas`, `modSelfUpdate.bas`,
  `modRelease.bas`.
- Banka: `src-vba/modBankaImport.bas` (`PullBankPdfsFromDriveProduction`, `ImportBankaInbox_*`),
  `modBankaImportParserPdfToText.bas` (`ResolvePdfToTextExePath`), `modBankaMapiranje.bas`,
  `frmBankaImport.frm`.
- GAS: `gas/Code.gs`, `gas/Monitoring.gs`, `gas/DriveFolder.gs` (`bootstrapAgriXFolderTree`),
  `gas/bank-pdf-downloader/` (`Code.gs`, `appsscript.json`, `README.md`).
- Build/release: `tools/release.sh`, `tools/stamp-build.sh`, `src-vba/modVbaTools.bas` (`ImportAllVBA`).
- Vodiči: `install/AgriX_Onboarding_Vodic_Novi_Klijent_v2.md`,
  `docs/production-runbook-banka-import-setup.md`, `docs/SELF_UPDATE.md`, `docs/RELEASE_PROCEDURE.md`.
