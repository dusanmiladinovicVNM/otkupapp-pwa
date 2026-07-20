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
   - **`PrepareRuntimeForSelfUpdate`** (otkaz SVIH `OnTime` tikova — sync +
     `AutoSaveTick` + StanicaLock heartbeat; release dinamičkih kontrola;
     unload formi; events/screen OFF),
   - **faza 1** `ImportFromFolder`: **delta-skip** — komponenta čiji je kod
     identičan novom telu se NE dira; ostale idu code-merge
     (`DeleteLines`+`AddFromString`). U `failed` idu SAMO `.bas`/`.cls`;
     forma/sheet čiji merge padne dobija best-effort **rollback na stari kod**
     + „potreban reinstall" u izveštaju (nikad `Remove`!),
   - **faza 2** (ako ima `failed`): `Remove` njih (uz type-guard: samo
     std/class modul) → `OnTime +2s` → `RunSelfUpdatePhase2` ih `Import`-uje,
   - **save** + „zatvori i otvori"; `EnableEvents`/`ScreenUpdating` se
     **vraćaju na svakom izlazu** (uspeh/greška/prekid).

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
7. **Komponenta koja padne u fazi 1 sme u `Remove`/fazu 2 SAMO ako je `.bas`/`.cls`.**
   Ranije je SVAKA `failed` komponenta išla u `VBComponents.Remove` — uklanjanje
   FORME u runtime-u je zamka #1 (korupcija + Document Recovery = **crash Excela**),
   a faza 2 uvozi samo `.bas`/`.cls` pa bi forma i **trajno nestala** iz projekta.
   Ovo je bila glavna rupa za „self-update crashuje Excel" posle v2.16.1 (prvi
   release-i sa masivnim izmenama formi: `frmDokumenta` storno framework,
   `frmOtkupAPP` integritet overlay). Sada: `failedOut` filtrira po ekstenziji,
   `Remove` ima dodatni type-guard (`Type` 1/2), a forma čiji merge padne dobija
   **best-effort rollback** starog koda + „potreban reinstall" u izveštaju.
8. **`Application.OnTime` tikovi van sync-a** (`modJournaling.AutoSaveTick`,
   `modStanicaLock.HeartbeatStanicaLock` — 90s, `modStornoWarm.StornoWarmTick`)
   mogu da opale usred importa ili u prozoru između faza (dok su „tvrdi" moduli
   uklonjeni) → demand-compile polomljenog projekta; `AutoSaveTick`/`StornoWarm`
   bi uz to i **snimili polu-ažuriran fajl**. `PrepareRuntimeForSelfUpdate` sada
   otkazuje SVE (`StopAutoSaveTimer`, `StopHeartbeatTimer`, `StopStornoWarm`).
   **NB:** kad se u `main` doda nov `Application.OnTime` tajmer, MORA se dodati i
   njegov `Stop*` u `PrepareRuntimeForSelfUpdate` (StornoWarm je bio propušten
   jer je stigao posle prvog hardening rada).
9. **`EnableEvents`/`ScreenUpdating` se moraju VRATITI na svakom izlazu update
   toka** (`RestoreRuntimeAfterSelfUpdate`) — inače `Workbook_Open` ne opali pri
   sledećem otvaranju fajla u ISTOJ Excel instanci („zatvori i otvori" onda
   izgleda kao da je update ubio aplikaciju), a `Workbook_BeforeClose` higijena
   se preskoči. (Tokom prozora faza 1→2 events namerno OSTAJU off; vraća ih
   `RunSelfUpdatePhase2`.)
10. **Delta-skip:** komponenta čiji je kod bajt-za-bajt identičan novom telu
    (`SameCode`, ignoriše samo završne CR/LF) se **ne dira** — ranije se na svaki
    update prepisivao CEO projekat (~90 komponenti), pa je i najmanji release
    nosio pun COM-edit rizik. Posledica: faza 2 se sada dešava samo kad je neki
    „tvrd" modul stvarno izmenjen, a update velikog skoka verzija dira samo
    stvarno promenjene komponente.
11. **NOVE `WithEvents` deklaracije u FORMI — utvrđeni krivac za crash
    2.16.1→2.21.0.** Dodavanje event-sink deklaracija (`Private WithEvents x As
    MSForms.Y`) u deklaracioni blok POSTOJEĆE forme kroz code-merge je ista
    klasa kvara kao #3 (bind event interfejsa u toku COM edita), a pad merge-a
    forme je (pre guard-a #7) vodio u `Remove` forme = korupcija/crash.
    **Pravilo ubuduće:** event sink za runtime kontrole formi ide kroz
    **`clsUiSink`** (generički WithEvents omotač; forma ima `WireSink` helper +
    jedan Public `UiSinkEvent` dispatcher) ili kroz namensku klasu
    (`clsBlokUI` obrazac) — **NIKAD novi `Private WithEvents` u `.frm`**.
    Post-2.16.1 form-WithEvents (storno centar/finder/undo/nedovršeno/recovery
    u `frmDokumenta`; integritet overlay u `frmOtkupAPP`) su prebačeni na
    `clsUiSink`, čime se deklaracioni blok formi vratio na 2.16.1-kompatibilan
    oblik (samo inertni dodaci: plain `MSForms.` reference, `String`/`Boolean`,
    `As Object`). Zatečeni PRE-2.16.1 form-WithEvents su zamrznuti (klijenti ih
    već imaju — uklanjanje bi opet menjalo deklaracije).
12. **ATOMARNOST — nikad ne snimaj polu-nov projekat.** Update se snima SAMO pri
    punom uspehu (`SaveWorkbookVerified`: `Save` bez greške I `ThisWorkbook.Saved`).
    Svaki fatalni ishod → NE snima se, i **`AbortSelfUpdateClose` auto-zatvara
    svesku bez snimanja** (`Saved=True` + `Close SaveChanges:=False`) — tehnička
    garancija, ne oslanja se na to da operater neće `Ctrl+S` ni da OneDrive
    AutoSave neće upisati polu-nov projekat. Pošto se pre pune uspešnosti nikad ne
    snima, disk ostaje **stara ispravna verzija**. **Zašto je važno:** `APP_VERSION`
    živi u `modConfig.bas` (soft `.bas`, merge-uje se u fazi 1); da se snimi
    parcijalni projekat sa novim `APP_VERSION` a starom formom, `CheckForUpdateOnOpen`
    bi na sledećem startu video „ažurno" i **nikad više ne bi ponudio isti release**.
    Fatalno = forma ne može merge (`needsReinstall`), faza-2 `Import` padne ili se
    ne verifikuje (`ImportedOk`), stara komponenta **još postoji** (`Remove`
    nedovršen — inače tiho preskočena = mešan build), `imported <> expected`,
    `Save` ne uspe, ili **download nepotpun** (#13). (Higijena: `AbortSelfUpdateClose`
    radi jer `ShutdownApp`/`FlushNow` gledaju `.Saved`, koji je postavljen na True.)
13. **Download mora biti KOMPLETAN.** `DownloadReleaseFiles` vraća i „očekivano"
    (svi podržani fajlovi iz Drive listinga) i „preuzeto"; `RunSelfUpdate` prekida
    ako `preuzeto <> očekivano`. Ranije se gledalo samo `n = 0`, pa je i 1/95
    fajlova prolazilo kao validan release → parcijalan merge (npr. nov `modConfig`
    bez nove forme).
14. **Tvrde module PREPOZNAJ UNAPRED, ne kroz pali `AddFromString`.** `IsHardModuleBody`
    (module-level `WithEvents` ili `As MSForms.`, uz strip stringa/komentara da
    reč u komentaru ne da lažni pozitiv) rutira tvrde `.bas/.cls` pravo u fazu 2 —
    `AddFromString` (koji baš i diskonektuje `CodeModule`) se nad njima **nikad ne
    poziva**. Ranije su išli „error-driven" (prvo pao `AddFromString` pa u fazu 2),
    što je za NOV tvrd modul (`clsUiSink`) značilo instalaciju baš opasnim putem.
    Faza 2 dobija **tačnu listu** fajlova (iz Settinga), ne skenira ceo temp (inače
    `SKIP_MODULES` bypass / uvoz sirovo-palih ili dev modula).
15. **Startup watchdog za prekinut update.** Ako faza 2 nikad ne opali (Excel
    zatvoren, OnTime otkazan), `modConfig` je u memoriji nov ali disk je stara
    (nesnimljena) verzija, a `pending` marker ostaje u registru.
    `RecoverPendingSelfUpdate` (iz `StartApp`) čisti stale marker + temp i
    obavesti. **NE pokušava da „dovrši" fazu 2** nad starim projektom — to bi
    spojilo stari i nov kod. Marker se briše tek na **uspeh ili kontrolisan
    abort** (ne na početku faze 2), da crash usred importa ostavi trag.
16. **Multi-copy izolacija (naročito DEV test na kopiji).** Oba `Application.OnTime`
    poziva (`RunSelfUpdate`, `RunSelfUpdatePhase2`, `AbortSelfUpdateClose`) su
    **workbook-qualified** (`'Ime.xlsm'!Proc`) — inače Excel može da razreši proc
    u pogrešnoj otvorenoj kopiji. `phase2` registarsko stanje je **scope-ovano po
    workbook imenu** (`P2Section`) — dve kopije ne dele/gaze pending. `RunSelfUpdate`
    ide **PRE min-version enforce gate-a** u `StartApp`: inače bi `enforce=YES`
    ugasio baš klijenta kome update treba, pre nego što stigne do provere. (Ostaje
    i operativni stopgap `VERSION_ENFORCE=NO` dok se flota ne digne — sad manje
    kritičan.)
17. **Release na Drive-u mora biti KOMPLETAN snapshot** (build strana,
    `PublishReleaseToDrive`): `version.json` se objavljuje **tek pošto SVI code
    fajlovi uspešno stignu** (ako makar jedan padne → manifest se ne dira, release
    ostaje na staroj verziji); zastareli (obrisani) code fajlovi se **prune-uju**
    (`DriveTrashFile`) da ih klijent ne bi ponovo skidao; manifest nosi `files`
    listu (ime+veličina). Klijentski „preuzeto = očekivano" (#13) štiti od
    nepotpunog *download-a*, ali ne od nepotpune *objave* — zato oba kraja.
    (Sledeći korak ka pravom snapshot-u: SHA-256 po fajlu + versioned folderi.)
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
- **VAŽNO — distribucija ispravki samog updatera:** pošto je `modSelfUpdate` u
  `SKIP_MODULES`, ispravke self-update mehanizma (npr. hardening protiv crash-a)
  **ne stižu self-update-om**. Klijenti ih dobijaju jednokratno ručno:
  `ImportAllVBA` iz ažuriranog git klona na klijent mašini, ili zamena `.xlsm`
  novom kopijom (reinstall). Tek POSLE toga self-update opet sme da se koristi.

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

## Lokalni DEV test (najlakše — bez Drive-a, bez publish-a)

Da se self-update **engine** testira na svojoj mašini pre nego što bilo šta ode na
flotu: `Alt+F8 → **RunSelfUpdateDev**` (u `modSelfUpdate`). Code-merge-uje **ovu**
svesku iz **lokalnog `src-vba` foldera** (git klon) kroz **isti `RunSelfUpdateCore`**
kao pravi self-update (faza 1 + faza 2), samo bez Drive download-a, bez
`REL_FOLDER_ID` i bez Google auth-a.

> **Zašto ne `ImportAllVBA`?** `ImportAllVBA` rekreira komponente (`Import`) i
> **toleriše sve** — self-update ide **code-merge** (`DeleteLines`+`AddFromString`),
> a baš tamo su forme pucale. Zato je jedini validan test onaj koji koristi
> code-merge put; `RunSelfUpdateDev` to radi (deli jezgro sa produkcijom).

**Postupak:**
1. `git pull` na klonu (da `src-vba` ima verziju koju testiraš).
2. Otvori **KOPIJU** klijentske sveske (ne build-master; merge menja kod ove sveske).
3. `Alt+F8 → RunSelfUpdateDev` → u pickeru izaberi svoj `...\otkupapp-pwa\src-vba\`.
4. Backup se napravi sam; merge teče; na kraju „zatvori i otvori".
5. Posle restarta: `Alt+F11` → **nema duplikata** (`modX1`); `Debug → Compile` čist;
   otvori forme koje su menjane (Dokumenta „Storno", Integritet overlay…).
6. Idempotencija: pokreni `RunSelfUpdateDev` **drugi put** iz istog foldera →
   izveštaj mora reći „Ažurirano: 0, bez izmene: ~sve" (delta-skip radi).
7. Rollback po potrebi: `Backup\AgriX_pre-update_*.xlsm`.

> **GUARD (zaštita od slučajnog klika — NIJE bezbednosna granica):**
> `RunSelfUpdateDev` radi samo ako je izabran folder oblika git klona — ime
> `src-vba` + `.git` u roditeljskom folderu (`IsDevCloneFolder`). To sprečava da
> se iz `Alt+F8` slučajno pokrene nad proizvoljnim folderom i pokvari sveska.
> **Nije sigurnosna granica** — ko namerno napravi `…\fake\.git` + `…\fake\src-vba`
> prolazi. Za pravo razdvajanje bio bi potreban dev-only modul koji se ne
> objavljuje ili build-flag; ovde je svesno zadržan u `modSelfUpdate` (odluka
> operatera) uz ovaj lagani guard. Svako ko ionako može `Alt+F8` može i `Alt+F11`
> / `ImportAllVBA` (ista klasa mogućnosti). Prolazi kroz ISTI atomski
> `RunSelfUpdateCore` kao produkcija (bez Drive-a).

`modSelfUpdate` je u `SKIP_MODULES`, pa se DEV harness pri merge-u **ne prepisuje**
(ostaje aktivan), isto kao u produkciji. Za test i Drive transporta (download) →
i dalje `PublishReleaseToDrive` u **test** `AgriX_Release` folder (staged rollout).

---

## Funkcija B (backlog)

Nedeljni `.xlsx` backup **podataka** u Drive `AgriX_Backup` (`BACKUP_FOLDER_ID`) —
proširenje `modJournaling.BackupFileOnStart`, reuse `modDrive.DriveUploadFile`.
Vidi `backlog/backlog_2.md`.
