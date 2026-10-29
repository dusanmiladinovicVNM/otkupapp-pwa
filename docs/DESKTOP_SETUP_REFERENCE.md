# AgriX — Desktop Setup Reference (v2.28.4)

> **Svrha.** Praktičan, konsolidovan vodič za **setup / prvo pokretanje / instalaciju** desktop
> AgriX-a. Drži tačno ono što nigde drugde nije sažeto na jednom mestu; za dublje slojeve upućuje,
> ne duplira:
> - arhitektura/boot/invarijante → `docs/ARCHITECTURE_REFERENCE.md`
> - shema / vlasništvo upisa (A11) → `docs/DOMEN/` (`README.md`, `ARCHITECTURE_CONTRACT.md`)
> - banka import od nule (dva naloga / dva GAS-a) → `docs/production-runbook-banka-import-setup.md`
> - self-update / release → `docs/SELF_UPDATE.md`, `docs/RELEASE_PROCEDURE.md`
> - licenciranje → `docs/licenciranje-po-uredjaju.md`
> - tri config tabele + schema-iz-koda (pravila) → `.claude/rules/podaci-i-config.md`
>
> Puni ASCII nije uslov (ovo je `.md`). Verzija koda u trenutku pisanja: `modConfig.APP_VERSION = 2.28.4`.

---

## 1. Model — šema dolazi iz koda

`schema/schema.json` (KANON, u gitu) → `tools/gen_schema_module.py` → `src-vba/modSchema.bas`
(generisan) → `.xlsm` (posledica). **Sveska NIJE izvor istine.**

Posledice za pripremu:
- **`SetupNewPC` GRADI tabele** (`modSchema.EnsureAllTables`, `modSetup.bas:55`) iz kanona — ne samo što
  ih proverava. „Blanko master" više ne mora unapred da nosi kompletnu shemu.
- **Prazna radna sveska se pravi iz koda:** `python tools/make_dev_workbook.py --out <put>\AgriX_DEV.xlsm`
  (prazan `.xlsm` → uvoz svog VBA → `EnsureAllTables` → `tblLocalConfig`/`tblSEFConfig`, licenca off).
  Traži Windows + Excel + pywin32 + „Trust access to the VBA project". Koristi se i za **oporavak**
  sveske kojoj je VBA projekat korumpiran (kad `ImportAllVBA` merge pukne).
- Kanon je **47 tabela / ~656 kolona** (stariji komentari „41/590" su zastareli).

---

## 2. Tri config tabele — NE mešati store

| Tabela | Nosi | Get/Set |
|---|---|---|
| **`tblSEFConfig`** (global, putuje sa `.xlsm`) | Google/PWA, SEF, SELLER, MONITORING, LICENSE/TRIAL, `CLOUD_SYNC_ENABLED`, print-mode i poslovni toggle-i | `GetConfigValue` / `SetConfigValue` (`modConfig.bas:1169/1183`) |
| **`tblLocalConfig`** (per-mašina, ne putuje) | `APP_ROOT_PATH`, `APP_SETUP_COMPLETED`, `BANKA_*_PATH`, `BANKA_DRIVE_SOURCE_PATH`, `PDFTOTEXT_EXE_PATH` | `GetLocalConfigValue` / `SetLocalConfigValue` (`modSetup.bas:376/426`) |
| `tblConfig` | — LEGACY, mrtav (samo kolonski validiran ako postoji) | — |

**Dve česte greške (naučene):**
- ⚠ **Google/cloud ide u `tblSEFConfig`, NE `tblConfig`** — stari `tblConfig` čitač je uklonjen
  (`modSetup.bas:488`). `GOOGLE_PWA_FOLDER_ID` = ID foldera `01_Sheets/02_Master` (ne ceo `01_Sheets`).
- ⚠ **`PDFTOTEXT_EXE_PATH` i `BANKA_*` MORAJU u `tblLocalConfig`** — runtime ih čita iz Local; upisani u
  SEFConfig „tiho ne rade".

**Editor:** Matični podaci → **Podešavanja** (`modPodesavanja`). Anti-tamper: čim je setup zelen,
`tblSEFConfig` se sakrije (VeryHidden); izlaz u nuždi `Alt+F8 → ShowConfigSheet`. Grupa **„Banka / lokalno"**
piše u `tblLocalConfig` (store="local"), sve ostalo u `tblSEFConfig`. Poppler i folder polja imaju „…"
browse dugme.

---

## 3. Boot i first-run kapija

`ThisWorkbook.Workbook_Open` → `modMain.StartApp`. Redosled (sve kapije **opt-in + fail-open** — sveža
mašina prolazi bez blokade):

```
sakrij Excel → splash (faza BOOT) → licenca → self-update → min-verzija → prijava
→ FIRST-RUN kapija → schema-check → ljuska frmOtkupUI (faza APP)
```

- **First-run** (`modMain.bas:112`): ako `APP_SETUP_COMPLETED` (u `tblLocalConfig`) nije `"DA"` → ponudi
  `SetupNewPC`. Jednokratno; fail-soft.
- **Jedna forma `frmOtkupUI`**, četiri faze: `BOOT` (splash), `LOGIN` (prijava), `MINI` (kartica dok je
  Excel otkriven), `APP` (ljuska). Zasebne forme `frmSplash`/`frmOtkupAPP`/`frmLogin`/`frmExcelMini` više
  ne postoje.
- **Tačno 3 tačke gde se sveska (listovi) otkriva:** (1) odbijena kapija (licenca/verzija/login), (2)
  first-run `SetupNewPC` (FileDialog pickeri), (3) dugme **„Otvori Excel"** iza prava `OBL_OTVORI_EXCEL`.

---

## 4. Folderi i Poppler

Root aplikacije = **folder u kome stoji `.xlsm`** (`GetDefaultRootPath = ThisWorkbook.path`). Sve se pravi
pored radne sveske:

```
<APP_ROOT>  (npr. C:\AgriX\)
├─ AgriX.xlsm
├─ Tools\poppler\Library\bin\pdftotext.exe          ← Poppler (APP_PDFTOTEXT_RELATIVE_EXE_PATH)
├─ Backups\ Logs\ Journal\ Export\ Temp\ Secrets\    ← EnsureAppFolders
├─ Bank_Izvodi\{Inbox,Processed,Error}\              ← SetupBankFolders (LOKALNI)
└─ (9 PDF izlaza) Otkupni listovi\ Prijemnice\ Otpremnice\ Revers ambalaze\
   Kartice kooperanata\ Paletni listovi\ Preradni listovi\ Specifikacije\ Izvestaji\
```

**Poppler razrešavanje** (`modBankaImportParserPdfToText.ResolvePdfToTextExePath`): eksplicitni
`PDFTOTEXT_EXE_PATH` (tblLocalConfig) ima prioritet; inače default `<xlsm>\Tools\poppler\Library\bin\
pdftotext.exe` (relativno na `ThisWorkbook.path`, ne na perzistirani `APP_ROOT_PATH` — da premeštanje
paketa ne obori putanju); golo ime `pdftotext.exe` prolazi iz PATH-a. Podešavanje: `Alt+F8 →
SetupPopplerInteractive` ili Podešavanja → „Izaberi Poppler".

---

## 5. Procedura (setup / first-run / instalacija)

Sažeto; pune korake po klijentu drži `install/AgriX_Onboarding_Vodic_Novi_Klijent_v2.md`.

| Faza | Ko / gde | Suština |
|---|---|---|
| **0. Dev build** | dev mašina | `gen_schema_module.py --check` (kanon↔modSchema u koraku) → `make_dev_workbook.py` (ili `ImportAllVBA` nad zatečenom) → `Debug→Compile` → `AssertBlankBuild` (ručno, proverava da su transakcione tabele prazne) → potpiši `.xlsm` → `PublishReleaseToDrive` (flota) |
| **1. Google po klijentu** | `ops@agrix.rs` | root folder `AgriX_C00X_PROD` + GAS projekat `AgriX_C00X_GAS_PROD`; `bootstrapAgriXFolderTree` (stablo + 32 `AGRIX_*_FOLDER_ID` u Script Properties); OAuth Desktop client; napuni `tblSEFConfig` (Google/MONITORING/CLOUD_SYNC) |
| **2. Paket** | dev | `AgriX.xlsm` (potpisan, BLANKO OK) + `Setup-AgriX.ps1` + `Tools\poppler` + `AgriX-VBA-Publisher.cer` + `docs\` |
| **3. Windows** | kod klijenta (admin) | `Setup-AgriX.ps1` → `C:\AgriX` + podfolderi, `Unblock-File`, cert → Trusted Location, shortcut, `install-log.txt` + PASS/FAIL |
| **4. Drive for Desktop** | mašina | banka `01_Bank` → **Add shortcut to Drive**; `00_Inbox` → Available offline; upiši `BANKA_DRIVE_SOURCE_PATH` (Podešavanja → „Banka / lokalno") |
| **5. First-run** | mašina | otvori AgriX → prihvati `SetupNewPC` (gradi shemu + foldere + config; zeleno → `APP_SETUP_COMPLETED=DA`); Poppler; `TestServerLink` |
| **6. Banka** | mašina | v. sekciju 6 |

---

## 6. Banka — dva GAS-a + Drive for Desktop

Lanac (VBA **ne** čita mailbox direktno):

```
Banka (email) → GAS #1 "Bank PDF Downloader" (na nalogu koji PRIMA izvode; Editor na 01_Bank)
→ Drive 00_Inbox/01_Bank → Google Drive for Desktop → lokalni ...\01_Bank  (= BANKA_DRIVE_SOURCE_PATH)
→ PullBankPdfsFromDriveProduction → Bank_Izvodi\Inbox → ImportBankaInbox_TX → pdftotext
→ tblBankaImport → modBankaMapiranje → tblNovac
```

**Povezivanje dva GAS-a = isti folder ID od `01_Bank` na tri mesta** (folder nema Script Property, ID se
uzima ručno iz Drive URL-a):
1. `01_Bank` podeljen kao **Editor** nalogu koji prima izvode (na njemu radi GAS #1,
   `gas/bank-pdf-downloader/`).
2. Taj isti ID u `BANK_IMPORT_CLIENTS_JSON.driveFolderId` (GAS #1).
3. Lokalna putanja do `01_Bank` u `BANKA_DRIVE_SOURCE_PATH` (`tblLocalConfig`).

Ključni `tblLocalConfig` banka-ključevi: `BANKA_DRIVE_SOURCE_PATH` (prazno = pull isključen),
`BANKA_DRIVE_MAX_FILES` (50), `BANKA_DRIVE_MIN_FILE_AGE_SECONDS` (15), `BANKA_INBOX/PROCESSED/ERROR_PATH`,
`BANKA_AUTO_IMPORT_ON_START` (NE), `BANKA_ALLOWED_EXTENSIONS` (pdf). Auto-map ključevi: poziv na broj /
tekući račun / `StanicaID` (OM noga — kooperant mora imati `StanicaID`). Puni setup od nule:
`docs/production-runbook-banka-import-setup.md`.

---

## 7. Šema iz koda — `EnsureX` ↔ kanon (koegzistencija)

- **`EnsureAllTables`** (`modSchema.bas:235`) — kanonski motor: kreira/dopunjava sve iz `schema.json`;
  idempotentno; **ne popravlja redosled kolona**.
- **`EnsureRuntimeSchema`** (`modSetup.bas:1173`) — self-heal na svakom startu: PRVO preimenovanja
  (`PreimenujKolonuAko`), pa `EnsureAllTables` ako otisak odudara, pa ručne idempotentne dopune
  (`EnsureUtovarSchemaCore`, `EnsureStornoVezeSchemaCore`, `EnsureSledljivostSchema`,
  `PrimeniFormateKanona`…). Pojedinačni `EnsureX*Schema` (Cenovnik, PaletniList, Dorade, Korisnici,
  AuditColumns, Poruke) zovu se iz Admin panela / ručno — svi na kraju zovu `EnsureDataTable` (aditivno).
- **`SchemaReadyOrFail`** (`modSchema.bas:407`) — **tvrda kapija pred upis**: proverava postojanje,
  **redosled kolona** (kanon mora biti prefiks zaglavlja; upis je pozicion) i **format** (General tiho
  kvari „3/2026", vodeću nulu, 18-cifreni račun). Fail-hard preko `modSchemaGuard.RaiseSistemski`.
- **Runtime dopune idu NA KRAJ** i otisak (`SchemaFingerprintActual`) ih ignoriše (meri samo kanonski
  prefiks) — zato ne prave lažan drift; to je mehanizam koegzistencije koda i kanona.

Nova/izmenjena kolona = izmena `schema.json` pa `gen_schema_module.py`; **nikad obrnuto**. `--check` je CI
kapija; pred uvoz u zatečenu svesku `tools/schema_diff.py <sveska>`.

---

## 8. Verifikacija / health

- `SetupNewPC` zeleno → `APP_SETUP_COMPLETED=DA`. Ako nešto fali → `"NE"` + poruka.
- `Alt+F8 → TestServerLink` — Google / GAS / banka Drive folder (živi link; u setup-u je **advisory**, ne
  obara zeleno).
- `RunSetupHealthCheck` (read-only re-provera) + Admin panel (health/ensure/googleauth dugmad; AUTH brana).
- Desktop-only režim: `EnableDesktopOnlyMode` (`CLOUD_SYNC_ENABLED=NO`) — gasi Google sync/lock/numerisanje;
  **ne gasi licencu** (gejt je `LICENSE_ENABLED`).

---

## 9. Gde je šta (reference, da se ne duplira)

| Tema | Autoritativno |
|---|---|
| Boot/startup ugovor, invarijante | `docs/ARCHITECTURE_REFERENCE.md` |
| Shema (kanon, registar), vlasništvo upisa A11 | `docs/DOMEN/README.md`, `ARCHITECTURE_CONTRACT.md`, `WRITE_OWNERSHIP.json` |
| Tri config tabele, schema-iz-koda, redosled kolona | `.claude/rules/podaci-i-config.md` |
| Banka sloj (dispatch, jaki ključevi, saldo) | `.claude/rules/banka.md` + runbook banka-import-setup |
| Self-update / release | `docs/SELF_UPDATE.md`, `docs/RELEASE_PROCEDURE.md`, `.claude/rules/sync-i-self-update.md` |
| Licenciranje (node-lock) | `docs/licenciranje-po-uredjaju.md` |
| Operativni vodič po klijentu | `install/AgriX_Onboarding_Vodic_Novi_Klijent_v2.md` |
| Priprema pre instalacije | `install/Priprema_pre_instalacije.txt` |
