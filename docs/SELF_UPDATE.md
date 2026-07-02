# Self-update (Funkcija A) — auto-ažuriranje koda preko Drive-a

**Cilj:** postojeći klijent na `Workbook_Open` detektuje da postoji novija verzija,
ponudi ažuriranje, i — uz potvrdu — povuče nov kod sa Drive-a i uveze ga u sebe.
Podaci se ne diraju (isti `.xlsm`); šema se self-heal-uje kroz postojeći
`InitApp/ValidateAllTables` posle restarta. Ovo je **alternativa** distribuciji
celog `.xlsm` (vidi `RELEASE_PROCEDURE.md` R3) — za izmene **koda**, bez migracije.

> Izvori istine za kod ostaju isti (git → klijent, R1/R2). Self-update je samo
> transport: ono što se objavi u `AgriX_Release` mora biti `git` build (vidi dole).

---

## Arhitektura

```
[build masina]  PublishReleaseToDrive (modRelease)
      |  upload src-vba (.bas/.cls/.frm/.frx/.doccls) + version.json
      v
[Drive] AgriX_Release/   (REL_FOLDER_ID u modConfig)
      |  CheckForUpdateOnOpen cita version.json, poredi app_version
      v
[klijent]  RunSelfUpdate: backup -> download -> import -> save -> restart
```

| Sloj | Gde |
|---|---|
| Drive REST (download/list/upload) | `modDrive` (`DriveDownloadToFile`/`DriveListFolder`/`DriveUploadFile`) |
| Build objava | `modRelease.PublishReleaseToDrive` (BUILD-ONLY, kao `modBuildGuard`) |
| Folder ID-jevi | `modConfig` `REL_FOLDER_ID` / `BACKUP_FOLDER_ID` |
| Detekcija + import | `modSelfUpdate` |
| Poređenje verzija | `modUpdateGate.VersionCompare` (reuse) |
| OAuth / HTTP | `modGoogleAuth.GetAccessToken`, WinHttp (reuse) |

---

## Build strana

Posle `ImportAllVBA → Compile → AssertBlankBuild` (vidi `RELEASE_PROCEDURE.md`
korak 7b), pre `git checkout -- modBuildInfo.bas`:

```
Alt+F8 -> PublishReleaseToDrive
```

Šalje **sve** `src-vba` fajlove (kao sirove bajtove) + `version.json`:

```json
{ "app_version": "2.5.0", "build_version": "vba-v2.5.0",
  "build_sha": "abc1234", "build_date": "2026-..." }
```

`app_version` (iz `modConfig.APP_VERSION`) je komparator za „ima novije". `build_*`
su informativni (placeholder `0.0.0-dev` ako `stamp-build` nije pokrenut — vidi
`RELEASE_PROCEDURE.md`).

---

## Klijent strana — tok

1. `Workbook_Open → StartApp` (posle `UpdateGateOrQuit`) → `CheckForUpdateOnOpen()`:
   čita `version.json`, poredi `app_version` sa lokalnim `APP_VERSION`. Ako novije
   → `MsgBox vbYesNo`. **Fail-soft** (offline/nema manifesta → tiho dalje).
2. Na „Da": `StartApp` zakaže `Application.OnTime Now, "RunSelfUpdate"` + `Exit Sub`
   (import se NE radi u toku `Workbook_Open` stack-a — vidi zamke).
3. `RunSelfUpdate` (prazan stack):
   - **backup** `<folder>\Backup\AgriX_pre-update_*.xlsm` (lokalno; rollback),
   - **download** svih fajlova u `%TEMP%\AgriX_update`,
   - **`PrepareRuntimeForSelfUpdate`** (stop sync, release dinamičkih kontrola,
     unload formi),
   - **faza 1** `ImportFromFolder`: code-merge (`DeleteLines`+`AddFromString`)
     svih komponenti; moduli koje ne može idu u `failed`,
   - **faza 2** (ako ima `failed`): `Remove` njih → `OnTime +2s` →
     `RunSelfUpdatePhase2` ih `Import`-uje,
   - **save** + „zatvori i otvori".

---

## Naučene zamke (TEŠKO stečeno — NE eksperimentiši ponovo)

Runtime manipulacija VBProject-a je krhka; svaka stavka ispod je realan kvar koji
se desio i fix koji radi:

1. **Forme se NE smeju `Remove`+`Import`-ovati u runtime-u** → „Errors during load"
   (60061) + korupcija (Document Recovery). Forme idu **code-merge** (zameni samo
   code-behind; dizajn/`.frx` se NE menja).
2. **Member `Attribute` linije** (`Attribute m_x.VB_VarHelpID = -1` kod `WithEvents`)
   → `AddFromString` baca „Syntax error". `ExtractModuleCode` **strip-uje SVE
   `Attribute` linije** (header + member) + preskače `Begin/End` dizajn blok.
3. **Module-level `MSForms.` deklaracije** (`Private mBtn As MSForms.CommandButton`)
   → `AddFromString` bind-uje MSForms tip-biblioteku u toku COM edita →
   diskonektuje `CodeModule` (`-2147417848`). `DeleteLines` prođe; pada baš
   `AddFromString`. **Ti moduli MORAJU kroz `Import`** (rekreacija komponente —
   podnosi MSForms decls; radi i u `ImportAllVBA`). Trenutno: `modOtkupBlok`,
   `modKarticaDetalji`, `modPodesavanja`, `modMouseWheel`, `clsWheelList`
   (rutiranje je automatsko: greška u fazi 1 → `failed` → faza 2 `Import`;
   nema hardkodirane liste). NIJE stvar živih instanci — `Release`
   referenci NE pomaže.
4. **`VBComponents.Remove` je ODLOŽEN** u runtime-u (izvrši se tek kad makro
   završi). `Remove`+`Import` u istom makrou → **`modX1` duplikati** → „Ambiguous
   name". Zato **dvofazni**: faza 1 `Remove` (queued) → `OnTime +2s` (flush) →
   faza 2 `Import` (čist modul). (Ovo je i razlog što `ImportAllVBA` „radi iz
   drugog pokušaja".)
5. **Encoding:** svi fajlovi se prenose kao **sirovi bajtovi** (ADODB.Stream
   binarno, `alt=media` / `uploadType=media`) — bez transkodiranja. Izvori su
   ASCII (posle lokalizacije), pa nema rizika; binarni put je i dalje ispravan.
6. **`PrepareRuntimeForSelfUpdate`** (release dinamičkih panela + unload formi +
   `StopScheduledSync`) je **higijena pre `Remove`-a** (da forma ne drži kontrole
   tih modula) — NE rešava zamku #3 (to rešava `Import`).

---

## Preduslovi i ograničenja

- **„Trust access to the VBA project object model"** mora biti uključen na
  klijent mašini (`SelfVBAAccessible` prijavi ako fali). Operater ga uključi pri
  instalaciji. AV može da reaguje na prepisivanje VBProject-a → dodaj exception.
- **OAuth:** klijent mora imati Google auth (`GetAccessToken` čita token iz
  config tabele). Blanko build bez auth-a → `DriveSelfTest` (modDrive) za dijagnozu.
- **NE update-uje se:**
  - **dizajn formi** (`.frx`/statičke kontrole) — samo code-behind formi; za
    izmene dizajna → pun reinstall `.xlsm`;
  - **`modSelfUpdate`** (na call-stack-u) i **`modVbaTools`** (dev tool) —
    `SKIP_MODULES`; ako se menjaju baš oni → reinstall;
  - **nove forme / novi sheetovi** (faza 1 ih prijavi „Preskočeno, reinstall").

---

## Smoke test posle release-a (naročito kad se menjaju moduli sa MSForms decls)

Novi/izmenjeni moduli sa `module-level MSForms.` deklaracijama ili `WithEvents`
(npr. `modMouseWheel`, `clsWheelList`) idu kroz dvofazni `Remove`+`Import`
(zamka #3/#4). Posle release-a koji ih dira, na **kopiji** klijenta:

1. `PublishReleaseToDrive` sa izmenjenim modulima.
2. Na kopiji klijenta pokreni self-update (`Workbook_Open` → „Da").
3. Posle restarta `Alt+F11` → proveri da **nema duplikata** (`modMouseWheel1`,
   `clsWheelList1`, `modX1` …); duplikat = „Ambiguous name" = faza 2 pala.
4. `Debug → Compile VBAProject` → mora proći bez greške.
5. Otvori formu sa ListBox-om, upali točkić (Podešavanja ili `MouseWheel_On`),
   proveri scroll; otvori/zatvori VBE (ne sme freeze).
6. Rollback po potrebi: `Backup\AgriX_pre-update_*.xlsm`.

---

## Funkcija B (backlog)

Nedeljni `.xlsx` backup **podataka** u Drive `AgriX_Backup` (`BACKUP_FOLDER_ID`) —
proširenje `modJournaling.BackupFileOnStart`, reuse `modDrive.DriveUploadFile`.
Vidi `backlog/backlog_2.md`.
