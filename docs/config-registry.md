# Config Registry — OtkupApp

> **Svrha:** Jedan izvor istine za SVE konfiguracione ključeve aplikacije.
> Trenutno su raspoređeni u 3 tabele (`tblSEFConfig`, `tblLocalConfig`, `tblConfig`)
> sa 3 različita pristupna API-ja, bez jasne logike. Ovaj dokument popisuje svaki
> ključ, gde se čita/piše, i u koju ciljnu tabelu treba da pređe.
>
> Status: **Faza 0** (popis i klasifikacija). Bez promena koda — ovo je osnova za
> migraciju opisanu u dnu dokumenta.
> Generisano: 2026-06-18.

---

## 1. Legenda

### Tip (ciljna klasa)
| Tip | Značenje | Sme u cloud / PWA? |
|---|---|---|
| **CONFIG** | Deljeni, ne-tajni poslovni/app config | DA |
| **SECRET** | Tajna (API ključ, token, lozinka) | **NIKAD** |
| **STATE** | Regenerabilni runtime handle / stanje (sheet ID, folder ID, lock) | Svejedno (nije tajna, ne unosi se ručno) |
| **LOCAL** | Vezano za konkretan računar (putanje, setup markeri) | **NIKAD** |
| **AUTH** | Identitet/nalozi za prijavu (osetljivo) | Samo izvedeno (Users tab) |
| **DEAD** | Mrtav ključ — nigde se ne čita ni piše | obrisati |

### Ciljne tabele
| Tabela | Drži | Sync |
|---|---|---|
| `tblConfig` | CONFIG + STATE + AUTH(izvedeno) | ceo sadržaj sme u cloud (bez ručnog filtera) |
| `tblSecrets` *(novo)* | SECRET | nikad |
| `tblLocalConfig` *(postoji)* | LOCAL | nikad |
| `—` | DEAD | obrisati |

### Trenutni pristupni API
| Funkcija | Tabela | Modul |
|---|---|---|
| `GetConfigValue` / `SetConfigValue` | `tblSEFConfig` | `modConfig.bas:486,500` |
| `GetLocalConfigValue` / `SetLocalConfigValue` | `tblLocalConfig` | `modSetup.bas:251,301` |
| `GetGoogleConfigValue` | `tblConfig` | `modSetup.bas:363` |
| `ConfigValue` / `SafeConfigValue` | `tblSEFConfig` (preko workbook-a) | `modMonitoring.bas:445,410` |

---

## 2. Ključni problemi (rezime)

1. **`tblSEFConfig` je de facto globalni config**, ne SEF. `GetConfigValue` uvek
   čita tu tabelu bez obzira na ključ (`modConfig.bas:490`). Samo ~10 od ~60 ključeva
   je stvarno SEF.
2. **🔴 Split-brain Google kredencijali:** setup proverava `tblConfig`
   (`modSetup.bas:477-479`), runtime čita `tblSEFConfig` (`modGoogleAuth.bas`, 5–6×).
   Moraju da postoje na oba mesta da bi radilo.
3. **🔴 Split-brain setup/health markeri:** `modSetup` piše u `tblLocalConfig`
   (`:69,73`), `modProductionHealthCheck` čita/piše u `tblSEFConfig` (`:801,919`).
   Nikad se ne vide.
4. **🟠 Tajne i javni PWA-config u istoj tabeli** → `ExportConfig` mora ručnu
   allow-listu (`IsPwaConfigKey`, `modStammdatenSync.bas:2028`) da tajne ne procure.
   Trenutno radi, ali krhko.
5. **🟡 Mrtvi ključevi:** `LICENSE_*` (7), `CLIENT_ID`, `CLIENT_NAME`, `ENV`,
   `APP_VERSION` (kao tabelarni ključ — koristi se konstanta).
6. **🟡 Pogrešno smešten:** `PDFTOTEXT_EXE_PATH` je u SEF listi, ali se čita iz
   `tblLocalConfig` (`modBankaImportParserPdfToText.bas:120`) — SEF kopija je mrtva.
7. **🟡 Redundantni ključevi:** `OtkupRokIsplate` vs `OTKUP_ROK_ISPLATE` (oba postoje).

---

## 3. Registar — `tblSEFConfig` (trenutno globalna kanta)

### 3a. Seller / matični podaci → `tblConfig` (CONFIG)
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `SELLER_NAME` | CONFIG | tblConfig | `modSEFMapper:104` (R); eksport `ExportConfig` | Naziv firme |
| `SELLER_PIB` | CONFIG | tblConfig | `modSEFMapper:105` | |
| `SELLER_MATICNI_BROJ` | CONFIG | tblConfig | `modSEFMapper:354` | |
| `SELLER_STREET` | CONFIG | tblConfig | `modSEFMapper:356` | |
| `SELLER_CITY` | CONFIG | tblConfig | `modSEFMapper:357` | |
| `SELLER_POSTAL_CODE` | CONFIG | tblConfig | `modSEFMapper:358` | |
| `SELLER_COUNTRY_CODE` | CONFIG | tblConfig | `modSEFMapper:359` | |
| `SELLER_ACCOUNT` | CONFIG | tblConfig | `modSEFMapper:360` | Tekući račun |
| `SELLER_EMAIL` | CONFIG | tblConfig | `modSEFMapper` | |
| `SELLER_LOGO_PATH` | LOCAL? | tblLocalConfig | `GetConfigValue` (1×) | **Putanja** — ako je apsolutna, vezana za mašinu → razmotri LOCAL |

### 3b. SEF e-faktura → `tblConfig` (osim API ključa)
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `SEF_BASE_URL` | CONFIG | tblConfig | `modSEFClient:283`, `modSEFValidator:263`, `modSEFTests:207` | |
| `SEF_API_KEY` | **SECRET** | **tblSecrets** | `modSEFClient:284`, `modSEFValidator:264`, `modSEFTests:208` | API ključ — nikad u cloud |
| `SEF_ENV` | CONFIG | tblConfig | `modSEFClient:285`, `modSEFTests:209` | DEMO/PROD |
| `SEF_PAYMENT_MEANS_CODE` | CONFIG | tblConfig | `modSEFMapper:362` | |
| `SEF_NOTE_DEFAULT` | CONFIG | tblConfig | `modSEFMapper:363` | |
| `SEF_PAYMENT_DUE_DAYS` | CONFIG | tblConfig | `modSEFMapper:776`, `modSEFTests:210` | |
| `SEF_FORCE_TODAY_ISSUE_DATE` | CONFIG | tblConfig | `modSEFMapper:408` | feature flag |
| `SEF_DEBUG_LOG` | CONFIG | tblConfig | `modSEFClient:359` | feature flag |
| `SEF_TEST_ALLOW_LIVE` | CONFIG | tblConfig | `modSEFTests:662` | test guard |
| `SEF_TEST_ALLOW_PROD` | CONFIG | tblConfig | `modSEFTests:663` | test guard (nije bio u originalnoj listi) |
| `SEF_TEST_ALLOW_CANCEL_STORNO` | CONFIG | tblConfig | `modSEFTests:1167` | test guard |

### 3c. Google OAuth / Drive
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `GOOGLE_CLIENT_ID` | CONFIG | tblConfig | `modGoogleAuth` (5× R, SEF) **+ `modSetup:477` (tblConfig)** | 🔴 split-brain |
| `GOOGLE_CLIENT_SECRET` | **SECRET** | **tblSecrets** | `modGoogleAuth` (6× R, SEF) **+ `modSetup:478` (tblConfig)** | 🔴 split-brain |
| `GOOGLE_ACCESS_TOKEN` | **SECRET** | **tblSecrets** | `modGoogleAuth:186,255` (W), R3 | runtime token |
| `GOOGLE_REFRESH_TOKEN` | **SECRET** | **tblSecrets** | `modGoogleAuth:188` (W), R3 | runtime token |
| `GOOGLE_TOKEN_EXPIRES_AT` | STATE | tblConfig | `modGoogleAuth:190` (W), R1 | |
| `GOOGLE_PWA_FOLDER_ID` | STATE | tblConfig | `modMasterSync`, `modStammdatenSync`, `modStanicaLock`, `modBrojevi`, `modGoogleSync*` (15× R) **+ `modSetup:479`** | 🔴 split-brain |
| `GOOGLE_STAMMDATEN_SHEET_ID` | STATE | tblConfig | `modStammdatenSync`, `modMasterSync`, `modStanicaLock`, `modGeoParcele` (R6/W4) | |
| `GOOGLE_KARTICE_SHEET_ID` | STATE | tblConfig | `modStammdatenSync:315` (W), R1 | |
| `GOOGLE_MGMT_SHEET_ID` | STATE | tblConfig | `modStammdatenSync:502` (W), R1 | |
| `GOOGLE_REPORTS_FOLDER_ID` | STATE | tblConfig | `modStammdatenSync:86` (R) | |

### 3d. Poslovni / PWA parametri → `tblConfig` (CONFIG)
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `OtkupAktivan` | CONFIG | tblConfig | allow-list eksport → PWA; GAS | da/ne otkup |
| `CenaVisnja` | CONFIG | tblConfig | `ExportConfig` preko `Cena*` prefiksa → PWA | „živ" preko prefiksa, ne direktno |
| `DefaultVrsta` | CONFIG | tblConfig | `ExportConfig`; GAS/PWA | |
| `DefaultSorta` | CONFIG | tblConfig | `GetConfigValue` (1×) + eksport | |
| `OtkupRokIsplate` | CONFIG | tblConfig | eksport → PWA | ⚠ duplikat sa `OTKUP_ROK_ISPLATE` |
| `OtkupPDVStopa` | CONFIG | tblConfig | eksport → PWA | |
| `DEFAULT_TIP_PALETE` | CONFIG | tblConfig | `modConfig` (CFG_DEFAULT_TIP_PALETE) | |
| `PALETA_PRINT_MODE` | CONFIG | tblConfig | `SetConfigValue` (W1) | |
| `OTKUP_PRINT_MODE` | CONFIG | tblConfig | `modConfig` (CFG_OTKUP_PRINT_MODE) | |
| `OTKUP_KLAUZULA` | CONFIG | tblConfig | `modConfig` (CFG_OTKUP_KLAUZULA) | |
| `OTKUP_ROK_ISPLATE` | CONFIG | tblConfig | `modConfig` (CFG_OTKUP_ROK) | ⚠ duplikat sa `OtkupRokIsplate` |
| `PDV_NADOKNADA_STOPA` | CONFIG | tblConfig | `modConfig` (CFG_PDV_NADOKNADA_STOPA) | |

### 3e. Management nalozi → `tblConfig` (AUTH)
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `MGMT_USER_1` | AUTH | tblConfig | `modStammdatenSync:2196` (prefiks `MGMT_USER`) → Users tab | login |
| `MGMT_USER_2` | AUTH | tblConfig | isto | |
| `MGMT_USER_3` | AUTH | tblConfig | isto (živ preko prefiksa) | |

### 3f. Monitoring / telemetrija
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `MONITORING_ENDPOINT` | CONFIG | tblConfig | `modMonitoring:421` | GAS Web App URL |
| `MONITORING_SECRET` | **SECRET** | **tblSecrets** | `modMonitoring:425,543` | redaktuje se u logovima (`:548`) |
| `MONITORING_ENV` | CONFIG | tblConfig | `modMonitoring:440` | |

### 3g. Sync / feature flagovi
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `CLOUD_SYNC_ENABLED` | CONFIG | tblConfig | `modConfig.IsCloudSyncEnabled:571` | glavni desktop-only prekidač |
| `SHEETS_SYNC_ENABLED` | CONFIG | tblConfig | `modProductionHealthCheck:800` | |
| `SYNC_AUTO_INTERVAL_MIN` | CONFIG | tblConfig | (proveriti čitača) | |
| `MASTER_SYNC_LOCK` | STATE | tblConfig | runtime lock | regenerabilno |

### 3h. Setup / health → `tblLocalConfig` (LOCAL) — TRENUTNO SPLIT
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `APP_SETUP_COMPLETED` | LOCAL | tblLocalConfig | piše `modSetup:69,81` (LOCAL); čita `modProductionHealthCheck:801` (**SEF**) | 🔴 split-brain |
| `APP_LAST_HEALTHCHECK_AT` | LOCAL | tblLocalConfig | piše `modSetup:73`+`modProductionHealthCheck:919` (**SEF**) | 🔴 split-brain |

### 3i. Putanja → `tblLocalConfig` (LOCAL)
| Ključ | Tip | Cilj | Čita / Piše | Napomena |
|---|---|---|---|---|
| `PDFTOTEXT_EXE_PATH` | LOCAL | tblLocalConfig | čita se iz `tblLocalConfig` (`modBankaImportParserPdfToText:120`) | SEF kopija je mrtva |

### 3j. MRTVI ključevi → obrisati
| Ključ | Tip | Napomena |
|---|---|---|
| `APP_VERSION` | DEAD | kao tabelarni ključ mrtav — koristi se konstanta `modConfig.bas:11`/`modMonitoring:431` |
| `CLIENT_ID` | DEAD | nigde se ne čita |
| `CLIENT_NAME` | DEAD | nigde se ne čita |
| `ENV` | DEAD | nigde se ne čita (koristi se `MONITORING_ENV`/`SEF_ENV`) |
| `LICENSE_ENABLED` | DEAD | feature nikad implementiran |
| `LICENSE_ENDPOINT` | DEAD | |
| `LICENSE_KEY` | DEAD | |
| `LICENSE_NEXT_CHECK` | DEAD | |
| `LICENSE_BOUND_PARTS` | DEAD | |
| `LICENSE_TOKEN` | DEAD | (ako se uvede licenciranje → SECRET) |
| `LICENSE_STATUS` | DEAD | |

---

## 4. Registar — `tblLocalConfig` (machine-local, uglavnom ispravno)

Sve LOCAL, ostaju gde jesu. Pristup: `GetLocalConfigValue/SetLocalConfigValue`.

| Ključ | Čita / Piše | Napomena |
|---|---|---|
| `APP_ROOT_PATH` | `modSetup:409,413,435` | koren app foldera |
| `APP_BACKUP_PATH` | `modSetup:417` | |
| `APP_LOG_PATH` | `modSetup:418` | |
| `APP_JOURNAL_PATH` | `modSetup:419` | |
| `APP_EXPORT_PATH` | `modSetup:420` | |
| `APP_TEMP_PATH` | `modSetup:421` | |
| `APP_SECRETS_PATH` | `modSetup:422` | folder za lokalne tokene (ako se koriste) |
| `BANKA_INBOX_PATH` | `modSetup:233`, `modBankaImport:922` | |
| `BANKA_PROCESSED_PATH` | `modSetup:234`, `modBankaImport:926` | |
| `BANKA_ERROR_PATH` | `modSetup:235`, `modBankaImport:930` | |
| `BANKA_AUTO_IMPORT_ON_START` | `modSetup:450` | da/ne |
| `BANKA_ALLOWED_EXTENSIONS` | `modSetup:454` | default „pdf" |
| `BANKA_DRIVE_SOURCE_PATH` | `modBankaImport:1104` | |
| `BANKA_DRIVE_DOWNLOADED_PATH` | `modBankaImport:1105` | |
| `BANKA_DRIVE_MAX_FILES` | `modBankaImport:1108` | |
| `BANKA_DRIVE_MIN_FILE_AGE_SECONDS` | `modBankaImport:1111` | |
| `APP_SETUP_COMPLETED` | `modSetup:69,81,137` | ⚠ vidi 3h — čita ga i SEF strana |
| `APP_SETUP_COMPLETED_AT` | `modSetup:70` | |
| `APP_SETUP_MACHINE_NAME` | `modSetup:71` | |
| `APP_SETUP_WINDOWS_USER` | `modSetup:72` | |
| `APP_LAST_HEALTHCHECK_AT` | `modSetup:73,82,116` | ⚠ vidi 3h |
| `PDFTOTEXT_EXE_PATH` | `modBankaImportParserPdfToText:120` | premestiti i SEF „kopiju" ovde (jedinstveno) |

---

## 5. Registar — `tblConfig` (trenutno samo Google, runtime je ignoriše)

Pristup: `GetGoogleConfigValue` (`modSetup.bas:363`). **Problem:** ovo je jedini
čitač — runtime aplikacija čita iste ključeve iz `tblSEFConfig`.

| Ključ | Čita | Napomena |
|---|---|---|
| `GOOGLE_CLIENT_ID` | samo `modSetup:477` | runtime gleda `tblSEFConfig` → 🔴 |
| `GOOGLE_CLIENT_SECRET` | samo `modSetup:478` | runtime gleda `tblSEFConfig` → 🔴 |
| `GOOGLE_PWA_FOLDER_ID` | samo `modSetup:479` | runtime gleda `tblSEFConfig` → 🔴 |

> Napomena: GAS strana ima poseban „Config" tab u Stammdaten Sheet-u (`gas/Code.gs:2902`)
> koji puni `ExportConfig` (`modStammdatenSync.bas:1954`) iz **`tblSEFConfig`** kroz
> allow-listu. To NIJE ista stvar kao VBA `tblConfig`.

---

## 6. Ciljni model i plan migracije

### Ciljni raspored (po dve ose: tajna? / vezano za mašinu?)
```
tblConfig      (deljeni, ne-tajni)  →  CONFIG + STATE + AUTH(izvedeno)  →  sme ceo u cloud
tblSecrets     (NOVO, tajne)        →  SECRET                          →  nikad iz workbook-a
tblLocalConfig (postoji, mašina)    →  LOCAL                           →  nikad sync
(obrisati)                          →  DEAD
```

Jedinstven pristup: **jedan config servis + namespace-ovani ključevi + kolona `Tip`**.
`Tip` vodi ponašanje: `SECRET` se nikad ne eksportuje, `LOCAL` se nikad ne sinhronizuje.
Time ručna `IsPwaConfigKey` allow-lista nestaje (eksportuje se cela `tblConfig`).

### Faze
- **Faza 0 (ovaj dokument):** popis + klasifikacija. ✅
- **Faza 1 — popravka 2 bug-a (bez restrukturiranja):**
  - P2/3c: ujednači Google kredencijale na jedno mesto. Preporuka: runtime je
    merodavan → `modSetup` da koristi `GetConfigValue` umesto `GetGoogleConfigValue`.
  - P3/3h: `APP_SETUP_COMPLETED` / `APP_LAST_HEALTHCHECK_AT` su LOCAL →
    `modProductionHealthCheck` da koristi `GetLocalConfigValue/SetLocalConfigValue`.
- **Faza 2:** jedinstven config ruter; stare `Get*ConfigValue` postaju tanki wrapperi.
- **Faza 3:** izdvoji SECRET u `tblSecrets`; ukloni `IsPwaConfigKey` allow-listu.
- **Faza 4:** idempotentna migraciona rutina (kopiraj → verifikuj → obriši staro);
  obriši DEAD ključeve (sekcija 3j).
- **Faza 5 (opciono):** normalizuj imena ključeva (`OtkupAktivan` → `OTKUP_AKTIVAN`),
  ukloni duplikat `OtkupRokIsplate`/`OTKUP_ROK_ISPLATE`, uz alias-mapu radi kompatibilnosti.

### Brojač (≈)
| Ciljna klasa | Broj ključeva |
|---|---|
| CONFIG | ~30 |
| SECRET | 5 (`SEF_API_KEY`, `GOOGLE_CLIENT_SECRET`, `GOOGLE_ACCESS_TOKEN`, `GOOGLE_REFRESH_TOKEN`, `MONITORING_SECRET`) |
| STATE | ~7 (Google sheet/folder ID-jevi, token expiry, sync lock) |
| AUTH | 3 (`MGMT_USER_*`) |
| LOCAL | ~22 (tblLocalConfig + `PDFTOTEXT_EXE_PATH`) |
| DEAD | 11 (`LICENSE_*` ×7, `CLIENT_ID`, `CLIENT_NAME`, `ENV`, `APP_VERSION`) |
