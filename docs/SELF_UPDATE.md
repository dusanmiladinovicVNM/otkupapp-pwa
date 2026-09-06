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

### F2 — versioned folderi + `current.json` (lanac poverenja)

Povrh flat kanala postoji **immutable snapshot po verziji** + kriptografski lanac.
Detaljan plan/status: `docs/SELF_UPDATE_SNAPSHOT_PLAN.md`.

```
AgriX_Release/                 (REL_FOLDER_ID)
  current.json                 <- pokazivač: app_version + release_folder_id + manifest_sha256
  version.json                 <- LEGACY (dual-write, za stare klijente)
  <flat .bas/.cls/...>         <- LEGACY (dual-write)
  releases/
    2.21.0/  manifest.json + svi src-vba fajlovi   (snapshot; bump-per-release)
    2.22.0/  ...                                    (retention: poslednjih 10)
```

Lanac: `current.json.manifest_sha256` → verifikuje bajtove `manifest.json` →
`manifest.files[].sha256` → verifikuje svaki skinuti fajl. **Bilo koji nesklad =
fail-closed** (prekid pre importa; ništa se ne snima).

- **Klijent** (`modSelfUpdate`): `GetRemoteAppVersion` prvo čita `current.json`
  (pa `version.json`); `ResolveReleaseSource` bira versioned folder i proverava
  `manifest_sha256` PRE ijednog download-a koda, inače **flat fallback** (nema
  `current.json`/`release_folder_id` → stari put). Novi klijent radi i sa starim
  publisher-om i obrnuto (dual-write) — nema „big bang" cutover-a.
- **Build** (`modRelease`): `PublishReleaseToDrive` upload-uje u OBA kanala, piše
  `manifest.json` (versioned) + `version.json` (flat), pa **na kraju** `current.json`
  (atomski „go live"; piše se samo ako je versioned manifest uspeo). `PruneOldReleases`
  drži 10 najnovijih. `RollbackReleaseTo` (Alt+F8) prepiše `current.json` na stariji
  `releases/<v>` (recompute `manifest_sha256`); `ListReleases` (Alt+F8) = pregled.
- **`manifest_sha256` je heš bajtova UPLOADOVANOG `manifest.json`** (posle
  `WriteReleaseTextFile`), koje klijent skida `alt=media` (sirovo) i hešuje —
  bajt-tačan round-trip (ista zamka kao per-file heš, #16).

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
   podnosi MSForms decls; radi i u `ImportAllVBA`). Trenutno **samo pet event-sink
   klasa**: `clsUiSink`, `clsFlatBtn`, `clsBlokUI`, `clsConfigBtn`, `clsAdminBtn`
   (rutiranje je automatsko: `IsHardModuleBody` pre-rutira, a greška u fazi 1
   → `failed` → faza 2 `Import`; nema hardkodirane liste). Veliki kontroleri
   (`modOtkupBlok`, `modPodesavanja`) su **soft** — module-level reference
   dinamičkih kontrola su `As Object`, a event routing je u sink klasama.
   NIJE stvar živih instanci — `Release` referenci NE pomaže.
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
   **Svi pozivi čišćenja su KASNO VEZANI** (`CallOptional` → `Application.Run`
   workbook-qualified), v. zamka #24.
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
   jer je stigao posle prvog hardening rada; `StopOtkupUITimers` je bio propušten
   jer je toast tajmer stigao sa ljuskom `frmOtkupUI`). Isto važi za novi runtime
   podsistem koji drži forme/sink-ove — treba mu `*_Release` u istoj listi
   (`Admin_Release` i `OtkupUI_Release` su bili propušteni iz istog razloga).
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
18. **SHA-256 verifikacija sadržaja (F1 — implementirano).** Manifest (`files[]`)
    nosi `sha256` (+ `size`) svakog fajla (`modRelease` reuse `modDrive.Sha256File`
    — isti `.NET SHA256Managed` kao PIN hash). Klijent je sada **manifest-driven**:
    skida samo fajlove iz `files[]` i **verifikuje SHA-256** svakog pre importa;
    nesklad (tiha korupcija/stale) → fajl se ne broji → `n <> expected` → fatalno
    (`AbortSelfUpdate`). **Fallback:** ako SHA-256 nije dostupan na mašini (retko;
    `Sha256File=""`, kao PIN plaintext fallback) → prisustvo/broj (kao pre F1), uz
    log. Stari publisher (manifest bez `files[]`) → legacy listing-download. Self-test:
    `Alt+F8 → Test_Sha256File`. **F2 (implementirano, `docs/SELF_UPDATE_SNAPSHOT_PLAN.md`):**
    versioned folderi `releases/<v>/` + `current.json` pokazivač + `manifest_sha256`
    lanac poverenja → snapshot + rollback.
19. **Updatable moduli NE smeju early-bind-ovati NOVE simbole iz frozen `modSelfUpdate`.**
    `modSelfUpdate` je u `SKIP_MODULES` → **ne stiže self-update-om**; klijent ga dobija
    tek ručnim bootstrap-om (`ImportAllVBA`/nov `.xlsm`). Star klijent koji se
    self-update-uje dobija **NOV `modMain` + STAR `modSelfUpdate`**. Ako nov `modMain`
    DIREKTNO (early-bound) zove NOV `modSelfUpdate` simbol koga star `modSelfUpdate`
    nema → **`Compile error: Sub or Function not defined`** obori CEO `StartApp` (iako
    je update „prošao uspešno"). Realan kvar: `modMain` je zvao **novi**
    `RecoverPendingSelfUpdate` → svaki star klijent koji se auto-update-uje crashuje na
    reopen-u. **Fix:** watchdog se zove **iznutra `CheckForUpdateOnOpen`** (isti modul,
    uvek razrešen), a `modMain` early-bind-uje SAMO **stabilne** `modSelfUpdate` simbole
    (`CheckForUpdateOnOpen`; `RunSelfUpdate` je i onako preko `OnTime` **stringa** =
    late-bound). **Pravilo:** nov cross-modul poziv u `modSelfUpdate` iz updatable
    modula = ILI ga sakrij iza postojećeg stabilnog simbola, ILI ga zovi late-bound
    (`Application.Run`) fail-soft.
20. **Forma/sheet sa module-level `WithEvents`/`MSForms` → `needsReinstall` (NE merge).**
    `AddFromString` takvog tela diskonektuje CodeModule (zamka #3), a forma se ne sme
    `Remove`+`Import`-ovati u runtime-u (zamka #1) — pa je jedini bezbedan ishod
    **reinstall** (fail-closed, ne polu-merge). Guard hvata i **novo** tvrdo telo i
    **zatečenu** tvrdu formu (`IsHardModuleBody(body) Or IsHardModuleBody(cur)`; radi i
    nad `CodeModule.Lines` koji koristi lone `vbCr`). **Posledica:** dok god forma ima
    ijedan module-level `WithEvents`, svaka njena izmena traži reinstall (ne self-update).
    **URAĐENO:** sve zatečene zamrznute `WithEvents` (`frmDokumenta`, `frmOtkupAPP`,
    `frmPalete`, `frmIzvestaj`, `frmAgrohemija`, `frmBankaExportPregled`) izmeštene su u
    `clsUiSink` → **nijedna forma više nema module-level `WithEvents`**, čime je
    **uklonjena klasa hard-crash-a** (crash 2.16.1→2.21.0 kad `AddFromString` naiđe na
    `WithEvents` u formi). **ALI forme i dalje NISU self-updatable — ostaju reinstall-only:**
    posle migracije zadržavaju module-level `As MSForms.*` reference (runtime kontrole —
    npr. `frmAgrohemija`: `Private m_btnPocetniDug As MSForms.CommandButton`), a
    `IsHardModuleBody` (pa i form-guard) **namerno** hvata i običnu `As MSForms.`
    deklaraciju → forma i dalje ide na **reinstall**. Empirijski potvrđeno:
    `RunSelfUpdateDev` nad izmenjenom `frmAgrohemija` = „Preskočeno (reinstall)". Guard
    **nije** „samo zaštita od regresije" — i dalje (ispravno) rutira svaku formu na
    reinstall; forme se distribuiraju **bootstrap-om** (`ImportAllVBA`/nov `.xlsm`), ne
    self-update-om. (Da postanu self-updatable trebao bi **procedure-level** merge; svesno
    se NE radi — whole-module `AddFromString` nad telom sa `As MSForms.` nije potvrđen kao
    bezbedan u formi.) `clsUiSink` proširen (`tgl.Change`/`lst.Click`/`btn.MouseMove`;
    `Bind` diže `Err.Raise` na nepodržan tip kontrole). **Pravilo dalje:** nikad ne vraćaj
    `WithEvents` u formu.
21. **Prazan modul (prazan stub `.bas/.cls`) → „same"/skip, NE fatalna greška.**
    `ExtractModuleCode` nad modulom bez tela (samo header) vrati prazno telo uz `Err=0`.
    Ranije je faza 1 tada dizala `vbObjectError+2801` („prazno telo") → modul je padao u
    `failed` → **forsirao fazu 2** (Remove+Import) na SVAKOM self-update-u i prikazivao
    „GRESKE: [-2147218703] prazno telo" (krivac: prazan orphan stub
    `clsSEFValidationResult.cls`). **Fix:** prazno telo uz `Err=0` = `„same"` (no-op) —
    prazan izvor NIKAD ne briše zatečen kod (fail-safe i protiv loše ekstrakcije nad
    ne-praznim fajlom); genuina greška ekstrakcije (`Err<>0`) i dalje → `failed` → faza
    2/reinstall. Mrtva `clsSEFValidationResult.cls` (0 referenci) **obrisana**.
22. **`IsHardModuleBody` ne vidi deklaraciju sa VIŠE RAZMAKA — rupa i ovde.**
    Detekcija traži doslovan niz `" AS MSFORMS."`, pa `Private x  As  MSForms.Label`
    (dva razmaka) prolazi kao **soft** modul i ide u `AddFromString`. VBE svoje
    eksporte normalizuje, pa se to ne može desiti round-tripom — može **samo ručnim
    editovanjem izvornog fajla**, što je tačno ono što programer radi.
    **Mereno 06.09.2026** kroz `ImportAllVBA` (PR #274): tako napisan `modArrayUtils`
    klasifikovan je kao soft; posle popravke (detekcija nad sažetim razmakom) isti
    fajl je klasifikovan kao **tvrd**. Popravka je prvo ušla samo u `modVbaTools`;
    `modSelfUpdate` je ostao sa rupom još jedan PR, a `modVbaTools` je pritom nosio
    komentar „isti obrazac stoji i u `modSelfUpdate`“ — koji više nije bio tačan.
    **Zatvoreno:** ista popravka je sada i u `modSelfUpdate` (`CollapseSpaces` +
    detekcija nad sažetim razmakom + sažimanje razmaka u `LowerOutsideStrings`).

    **Pouka koja je važnija od same rupe:** algoritam ima **dve privatne kopije**
    koje se ne mogu spojiti (`modSelfUpdate` je frozen bootstrap i ne sme da zavisi
    ni od čega što se update-uje). Komentar koji tvrdi paritet nije kapija — istruli
    tiho. Zato paritet sada čuva `tools/vba_parity_check.py` (u CI): šest procedura
    algoritma mora biti **kod-za-kod** isto u oba modula, uz korpus fikstura koji
    pina specifikaciju. Komentari smeju da se razlikuju, kod ne sme.

    **Poznat lažni pozitiv (svesno pinovan, NE popravljen):** `CodeLineUpper`
    odseca komentar, ali **ne** sadržaj stringa — pa `Private Const X As String =
    "Private WithEvents …"` klasifikuje modul kao **tvrd**. Greška je konzervativna
    (skuplji put kroz `Remove`+`Import`, nikad netačan ishod) i danas nema nijednog
    takvog modula u `src-vba`. Popravka bi bila izmena **semantike** detektora i
    mora u **obe** kopije istovremeno, uz merenje — ne usput.
23. **`AddFromString` nad tvrdim telom NIJE diskonektovao `CodeModule` — zamka #3
    se nije reprodukovala.** Isto merenje 06.09.2026: modul sa module-level
    `As MSForms.Label` primljen je kroz `AddFromString` **bez greške** (`Err=0`),
    a ne sa `-2147417848`. Prijavljuje se kao **NEREPRODUKOVANO**, bez zaključka i
    bez izmene modela — uslovi se razlikuju od originalnog nalaza (aplikacija nije
    bila podignuta, MSForms tip-biblioteka je već bila učitana, drugi Excel build).
    **Šta ovo NE znači:** da je podela soft/tvrd nepotrebna. Zamka #3 je nastala iz
    stvarnog crash-a i model se na osnovu jednog negativnog merenja ne menja.
    **Šta znači:** uslov pod kojim zamka #3 nastupa nije poznat tačno koliko smo
    mislili. Pre bilo kakvog opuštanja dvofaznog modela treba ponoviti merenje na
    živoj aplikaciji (forma podignuta, paneli izgrađeni) i na više Excel verzija.
24. **Čišćenje runtime-a ne sme biti compile zavisnost — zamka #19 naopako.**
    `PrepareRuntimeForSelfUpdate` je do sada zvao `OtkupBlok_Release`,
    `KarticaDetalji_Reset`, `MouseWheel_Off`… **direktno**. To je compile-time
    referenca iz **frozen** modula (`SKIP_MODULES`) na **updatable** module: čim
    self-update isporuči release u kome je jedna od tih procedura obrisana ili
    preimenovana, `modSelfUpdate` **prestaje da se kompajlira** — a on je jedini
    put kojim klijent dobija sledeću ispravku. Klijent bi ostao trajno zaključan
    na staroj verziji, i to tek posle uspešnog update-a (najgori mogući trenutak).
    Zamka #19 opisuje isti mehanizam u drugom smeru (nov `modMain` → star
    `modSelfUpdate`); ovo je smer frozen → updatable.
    **Fix:** sav teardown ide kroz `CallOptional` (`Application.Run` +
    `QualifiedProc`). Nepostojeća procedura = tih no-op. Lista imena je
    **kompatibilnosni ABI**, ne compile zavisnost — zato u njoj smeju (i treba da)
    ostanu legacy imena (`KarticaDetalji_Reset`, `MouseWheel_Off`) i posle brisanja
    tih modula iz `src-vba`: isti updater mora da očisti i STAR klijent koji ih još
    ima. Workbook-qualified je obavezno — golo `Application.Run "Proc"` sa dve
    otvorene kopije može da očisti tuđi runtime (zamka #16).
    Ista lista i isti redosled stoje u `modVbaTools.PrepareRuntimeForImport`;
    razlika između ta dva je bug (`StopOtkupUITimers`, `Admin_Release`,
    `OtkupUI_Release` su nedostajali u `modSelfUpdate` do ove izmene).
25. **`Err.Number = 0` posle `AddFromString` NIJE dokaz da je telo primenjeno.**
    To je tačno oblik u kome zamka #3 nastupa: COM diskonekt `CodeModule`-a može da
    ostavi modul sa **starim** (ili polovinim) kodom, a prolaz da se završi kao
    uspeh — pa `Save` overi nekonzistentan projekat. Merenje iz zamke #23 je ovo
    učinilo hitnijim, ne manje hitnim: `AddFromString` nad tvrdim telom je prošao
    **bez greške**, dakle „nema greške“ i „primenjeno“ nisu ista stvar.
    **Fix:** svaki soft upis nosi dokaz — `VerifyWritten` čita kod **nazad** iz
    projekta i poredi ga sa izvorom (`SameCode`). Nova komponenta uz to mora da
    prođe i `ImportedOk` (ime iz `VB_Name` + tip): `VBComponents.Add` ume da vrati
    `modX1` kad ime zauzme zaostala komponenta, pa bi se nov kod upisao **pored**
    starog umesto preko njega.
    Neuspeo dokaz ide **istim putem kao genuino pao upis** — još jedan prolaz, pa
    `.bas`/`.cls` u fazu 2 (`Remove`+`Import`), a forma/sheet u rollback +
    `needsReinstall` (nikad `Remove` forme, zamka #1). Nikad `"ok"`.
    **Zašto ovo ne pravi lažne padove:** ista `SameCode` već godinu dana odlučuje
    delta-skip nad istim parom (`CodeModule.Lines` vs `ExtractModuleCode`). Da
    proizvodi lažne razlike, svaki update bi re-merge-ovao SVE — a to je bio bug
    koji je zatvoren (zamka #10). Read-back koristi tačno taj, već dokazan par.
    Obrnuto isto važi: ako `SameCode` ne bi smela da se veruje kao dokaz upisa,
    ne bi smela ni kao osnov da se modul **ne dira**.
26. **Nijedan `Save` bez zavrsne provere celog release-a.**
    Pojedinačne provere hvataju svaka svoj korak — `VerifyWritten` dokazuje jedan
    upis, `ImportedOk` jedan `Import` — ali **nijedna ne hvata zbir**: komponentu
    koja je nestala *posle* svog uspešnog merge-a, tip koji se promenio, ili modul
    koji je drift-ovao u prozoru između faze 1 i faze 2. Takav projekat bi do sada
    bio **snimljen**.
    **Fix:** `VerifyReleaseProject(folder)` se zove na **oba** uspešna puta,
    neposredno pre `SaveWorkbookVerified`: za svaki izvorni fajl osim `SKIP_MODULES`
    proverava da komponenta postoji, da joj je tip tačan (`.bas`→1, `.cls`→2,
    `.frm`→3, `.doccls`→100), da je kod čitljiv i da `SameCode(projekat, izvor)`.
    Bilo koji problem → `AbortSelfUpdate` bez snimanja.

    **Smer provere je `izvor → projekat`, namerno jednosmerno.** Komponenta koja
    postoji u projektu a nema je u izvoru se **NE** prijavljuje. Self-update nije
    kanonski sinhronizator kao `ImportAllVBA`: postojeći klijenti nose legacy/stale
    module koje self-update istorijski ne briše, pa bi pravilo „višak = fatalno“
    oborilo update **svima**. Čišćenje zaostalih komponenti pripada `ImportAllVBA`
    (`RemoveStaleComponents`), ne updateru.

    **Prazan izvorni fajl ne opisuje komponentu** (isto pravilo kao zamka #21) —
    inače bi provera tražila ono što merge namerno preskače.

    **Pun tekst nalaza ide u `LogErr`, u poruku ide skraćen na 300 znakova.**
    `AbortSelfUpdate` stavlja poruku PRE uputstva operateru, a `MsgBox` tiho seče
    oko 1024 znaka — dugačak spisak drift-a bi progutao baš ono što operater mora
    da pročita.

    **Kapija je i statički čuvana:** `tools/vba_selfupdate_gates.py` (u CI) traži da
    svaki poziv `SaveWorkbookVerified` ima `VerifyReleaseProject` ranije u istoj
    proceduri, i da između njih stoji `AbortSelfUpdate`. Razlog: to je invarijanta
    **rasporeda** koda — put do nje se otvara tek kad neko doda **treći** uspešan
    izlaz i zaboravi kapiju, a tada nema crvenog testa, ima samo klijenta koji je
    snimio polu-nov projekat.
27. **Tvrda površina se ne održava komentarom nego kapijom.**
    Do sada je jedina evidencija toga šta je tvrdo bio komentar u zaglavlju
    `modSelfUpdate` — i on je nabrajao `modKarticaDetalji`, `modMouseWheel`,
    `clsWheelList`… module koji su u međuvremenu **obrisani**. Komentar koji niko
    ne izvršava istruli tiho; isti obrazac je već jednom ujelo kod zamke #22.
    **`tools/vba_hard_census.py` (u CI)** čita `src-vba/` i prijavljuje svaku tvrdu
    komponentu sa razlogom i tačnim redom. Pravila:

    | pravilo | pada kad |
    |---|---|
    | `TVRDA_FORMA` | `.frm` postane tvrd — takav update **nikad ne može da prođe** (ni `AddFromString`, zamka #3, ni `Remove`+`Import`, zamka #1) |
    | `TVRDA_DOCCLS` | document modul postane tvrd — ne može se `Remove`-ovati, pa za njega faza 2 ne postoji |
    | `TVRDA_BAS` | standardni modul postane tvrd — **bez izuzetka** |
    | `TVRDA_CLS` | nova tvrda klasa nije u `WHITELIST` |
    | `MRTAV_UNOS` | klasa iz `WHITELIST` više nije tvrda — inace whitelist truli u spisak imena bez značenja |

    `MRTAV_UNOS` je tu jer je baš to što je ovde pošlo naopako: lista koja se ne
    održava propusti prvo sledeće stvarno zaprljanje.
    Algoritam detekcije se **ne duplira** — census uvozi referentnu implementaciju
    iz `vba_parity_check`, koja je pod paritetnom kapijom sa obe VBA kopije. Treća
    kopija bi bila treća stvar koja može da divergira. Self-test tvrdi i da se
    objašnjenje (`hard_reason`) i odluka (`ref_is_hard`) slažu nad celim korpusom.

    Stanje u trenutku uvođenja kapije: **0 tvrdih `.bas`, 0 tvrdih `.frm`,
    0 tvrdih `.doccls`, 5 tvrdih `.cls`** (`clsAdminBtn`, `clsBlokUI`,
    `clsConfigBtn`, `clsFlatBtn`, `clsUiSink`) — mali `WithEvents` kernel, tačno
    ono što je i bila namera.
28. **Prazan `.doccls`/`.frm` stub kome fali komponenta NIJE „nova forma” — bio je
    doživotna blokada self-update-a.** Reprodukovano 06.09.2026 na živom klijentu
    (`AgriX_2.39.0_testVenivno.xlsm`), pri prvom ručnom `RunSelfUpdateDev`:

    ```
    Azuriranje NIJE moguce kroz self-update (forma/sheet zahteva reinstall).
    Azurirano: 1, bez izmene: 189, faza 2 (tvrdi): 0 (prolaza: 1)
    Preskoceno (novo, reinstall):
      skarticakoop.doccls
    ```

    `sKarticaKoop.doccls` ima **9 redova i svih 9 su header** — nijedan red koda.
    U `src-vba` je **42 od 43** `.doccls` takvo. `ExtractModuleCode` nad njima vrati
    prazno telo uz `Err=0`, a `extractOk` za `.doccls` prihvata i prazno
    (`Len(body) > 0 Or ext = "doccls"`). Kad komponente **nema** u svesci, tok je
    padao u završni `Else` → `"skip"` + `needsReinstall = True` → **fatalni abort**.

    **Posledica:** klijentu kome fali ijedan takav list self-update **nikad** ne može
    da prođe. Ne jednom — nikad, jer se prazan stub nema čime isporučiti. I gore:
    `AnyUpdatePending` je istu komponentu brojao kao „nova”, pa no-op kapija nikad
    nije opalila — svaki pokušaj je rušio **živ runtime** (pun teardown, unload svih
    formi) da bi završio abortom.

    **Fix (dva mesta, isto pravilo):** *prazan izvorni fajl ne opisuje komponentu*.
    - `ImportFromFolder`: `Len(body) = 0` uz nepostojeću komponentu → `"same"` (no-op).
      Forma/sheet koja **nosi kod** i ne postoji → i dalje `needsReinstall` (zamka #7/#20).
    - `AnyUpdatePending`: prazan izvor uz nepostojeću komponentu nije „nova komponenta”.

    Isto pravilo je već važilo za `.bas`/`.cls` (zamka #21) i za `VerifyReleaseProject`
    (zamka #26) — **nedostajalo je baš na putu koji odlučuje o reinstall-u.** Merge i
    završna provera su tvrdili suprotno jedno o drugom nad istim fajlom.

    **Zašto se stubovi NE brišu iz `src-vba`:** za `ImportAllVBA` prazan `.doccls`
    nosi informaciju „ovaj list postoji i nema kod” — brisanje iz izvora bi značilo
    reći da je list višak. Stubovi ostaju; tolerantan je merge, ne izvor.

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
(danas samo event-sink klase: `clsUiSink`, `clsFlatBtn`, `clsBlokUI`,
`clsConfigBtn`, `clsAdminBtn`) idu kroz dvofazni `Remove`+`Import`
(zamka #3/#4). Posle release-a koji ih dira, na **kopiji** klijenta:

1. `PublishReleaseToDrive` sa izmenjenim modulima.
2. Na kopiji klijenta pokreni self-update (`Workbook_Open` → „Da").
3. Posle restarta `Alt+F11` → proveri da **nema duplikata** (`clsUiSink1`,
   `modX1` …); duplikat = „Ambiguous name" = faza 2 pala.
4. `Debug → Compile VBAProject` → mora proći bez greške.
5. Otvori Otkup ekran i Podešavanja — paneli koje grade `modOtkupBlok` /
   `modPodesavanja` moraju da se izgrade i da dugmad reaguju (te module drže
   sink klase, pa se baš na njima vidi da je faza 2 prošla); otvori/zatvori VBE
   (ne sme freeze).
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
