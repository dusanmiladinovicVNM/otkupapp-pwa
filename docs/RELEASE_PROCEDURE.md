# Release & verzionisanje VBA koda (OtkupApp)

**Cilj:** kod kod klijenata se NE razilazi, i u svakom trenutku znaš koja
verzija / koji commit radi kod koga.

> Problem koji ovo rešava: kod živi u binarnom `.xlsm` kod svakog klijenta.
> Razilaženje nastaje kad se fix uradi direktno u klijentskom fajlu (i ostane
> samo tamo), ili kad različiti klijenti dobiju različit build a verzija ne
> razlikuje builde.

---

## TL;DR — tri pravila

- **R1 — Kod teče samo `git → klijent`.** Klijentski `.xlsm` je potrošan
  build-artefakt, nikad izvor.
- **R2 — 1 verzija = 1 commit = 1 tag** (`vba-vX.Y.Z`). Svaki ship bumpuje
  `APP_VERSION`, a `BUILD_SHA` pokazuje tačan commit i između tagova.
- **R3 — Podaci teku samo `klijent → nov fajl`** preko `modMigracija`. Kod i
  podaci nikad ne putuju zajedno.

## Izvori istine

| Šta | Gde |
|---|---|
| Kod | `src-vba/` u gitu |
| Verzija | `modConfig.APP_VERSION` |
| Build otisak | `modBuildInfo.BUILD_SHA` / `BUILD_VERSION` / `BUILD_DATE` (stamp pri buildu) |
| Build most (kod ↔ Excel) | `modVbaTools.ImportAllVBA` / `ExportAllVBA` |
| Migracija podataka | `modMigracija.MigrirajPodatkeIzStarog` |
| Self-heal šeme / fail-fast | `modSetup.Ensure*Schema` / `modSchemaGuard.RequireColumns` |
| Fleet (ko ima šta) | GAS `OtkupApp_Monitoring_PROD` → tab `Events` / `Fleet` |
| Distribucija (artefakt) | `builds\AgriX_x.x.x.xlsm` — **blanko** (samo kod, prazne tabele), isti za sve |
| Self-update kanal (postojeći klijenti) | Drive `AgriX_Release` (kod + `version.json`) ← `modRelease.PublishReleaseToDrive`; Drive helperi `modDrive` |
| Blanko provera (build-only) | `modBuildGuard.AssertBlankBuild` (ručno, pre `Save As`) |
| Min-version gate | GAS Script Properties `VERSION_MIN`/`VERSION_ENFORCE` · `modUpdateGate` |

---

## R1 — Kod teče samo `git → klijent`

- Izvor istine za kod je `src-vba/` u gitu. Klijentski `.xlsm` je **artefakt**.
- **Nikad ne ostavljaj fix samo u klijentskom fajlu.** Svaka izmena ide:
  git → `ImportAllVBA` u master `.xlsm` → isporuka.
- `ExportAllVBA` (obrnut smer) koristi se **isključivo za spašavanje /
  rekonsilijaciju** kad sumnjaš da je neko dirao klijenta (vidi dole).

## R2 — 1 verzija = 1 commit = 1 tag

Release koraci (po redu):

1. Završi i commit-uj sve izmene koda na feature grani; merge u `main`.
2. Bump `APP_VERSION` u `src-vba/modConfig.bas` (SemVer: patch za fix, minor za
   feature). Commit poruka npr. `release: vba v2.2.2`.
3. Tag na tom commitu:
   ```
   git tag vba-v2.2.2
   git push origin vba-v2.2.2
   ```
4. Stamp build otisak (upisuje HEAD sha u `src-vba/modBuildInfo.bas`, samo u
   working tree):
   ```
   bash tools/stamp-build.sh
   # ili na Windowsu:
   powershell -ExecutionPolicy Bypass -File tools\stamp-build.ps1
   ```
5. U master Excel fajlu: `Alt+F8` → `ImportAllVBA` (modul `modVbaTools`).
   (Prethodno postavi `FOLDER` konstantu na svoj `src-vba` put — vidi *Caveat*.)
6. `Debug → Compile VBAProject` — statička provera (nema duplih `Public`,
   balans `Sub`/`Function`). VBA se ne kompajlira u CI-u, ovo je obavezno ručno.
7. Snimi `.xlsm`, isporuči klijentima.
8. Vrati placeholder (stamp-ovanu vrednost **ne** commit-uješ):
   ```
   git checkout -- src-vba/modBuildInfo.bas
   ```

**Zašto i tag i SHA:** `APP_VERSION` sam ne razlikuje dva builda *između*
bumpova (lako zaboraviš da bumpuješ). `BUILD_SHA` pokazuje tačan commit uvek.
U fleet pregledu `0000000` znači **nestamp-ovan / dev build** (neko je build-ovao
bez `stamp-build`).

**`BUILD_VERSION` = auto verzija iz gita** (`git describe --tags --always`): na
tagu je čisto (`vba-v2.2.1`), a posle N commita se **sama diže**
(`vba-v2.2.1-3-gabc1234`). Zato je jedini ručni „bump" upravo `git tag` na
prekretnici (korak 3); između tagova verzija odražava stvarnost bez ijedne ručne
izmene. `APP_VERSION` u `modConfig` ostaje gruba baza koju zoveš klijentu.

## Jedna komanda (rutina) — sa OBAVEZNOM kapijom pre taga

```
bash tools/release.sh 2.2.2
# Windows:  powershell -ExecutionPolicy Bypass -File tools\release.ps1 2.2.2
# probni prolaz:  ... 2.2.2 --dry-run   /   ... 2.2.2 -DryRun
```

**Redosled je promenjen i to je suština.** Ranije je skripta radila bump → commit
→ push → tag → push, pa TEK ONDA rekla operateru da uradi Import i Compile;
behavior gate uopšte nije bio deo skripte. Tako je nastao `vba-v2.40.0`: tagovan i
objavljen uz pošteno zapisanu napomenu da testovi nisu ni pokrenuti. Poštena
napomena posle taga ne pomaže — tag već postoji.

Sada:

```
1  main + pull + cisto radno stablo
2  bump APP_VERSION U RADNO STABLO (bez commita)   <- kapija vrti BAS to
3  release gate: static / fixture / behavior / green / compile / external
4  commit + push                                    <- tek posle zelene kapije
5  anotiran tag sa verdiktom kapije + push
6  stamp build otisak
7  preostali Excel koraci (isporuka, ne provera)
```

Ako kapija padne: **nema commita, nema taga, nema push-a**, bump se vraća i radno
stablo ostaje čisto. Verdikt završava u poruci anotiranog taga — `git show
vba-v2.2.2` kasnije tačno kaže šta je bilo zeleno, šta izuzeto i nad kojim hashom
`src-vba`.

Kapija traži **Windows + Excel + pywin32**. Iz Linux sesije `behavior` i `compile`
padaju i release se zaustavlja — to je namerno, a ne kvar.

### Kad nešto stvarno ne može da se pokrene

```
bash tools/release.sh 2.2.2 --waive external --reason "SEF sandbox nedostupan 14.08."
powershell -File tools\release.ps1 2.2.2 -Waive external -WaiveReason "..."
```

Waiver **bez razloga se odbija**. Izuzeta kapija zadržava originalni status u
zapisu (`WAIVED (razlog) -- bilo je: FAIL ...`), pa se u tagu vidi i šta je tačno
zaobiđeno. `NOT_RUN` blokira isto kao `FAIL` — „nije pokrenuto" nije „prošlo".

Puna slika kapija i verdikata: `docs/TEST_PLATFORM.md`.

## R3 — Podaci: migracija, ne kopiranje koda

Kad isporučuješ NOVU verziju koda, ne nosiš tuđe podatke u njoj:

- Klijent dobije **prazan** novi `.xlsm` (samo kod, prazne tabele) — to je
  artefakt `builds\AgriX_x.x.x.xlsm` (ime prati `vba-vX.Y.Z`), **isti za sve**.
- **Blanko garancija (build-only):** pošto jedan fajl ide svima, gradi ga iz
  **praznog build-mastera** i pre `Save As` pokreni `Alt+F8 → AssertBlankBuild`
  (`modBuildGuard`) — nedestruktivno prijavi tabele s podacima. Ako master nosi
  podatke, oni bi iscureli svim klijentima. Guard je **ručan/build-only**, ne
  poziva se na `Workbook_Open` → kod klijenata nikad ne okida.
- `Alt+F8` → `MigrirajPodatkeIzStarog` (`modMigracija`): povuče **vrednosti** iz
  starog fajla **po imenu kolone** (preskače `tblRpt*`, merge-uje config tabele).
  Za preimenovane kolone između verzija koristi `StaroImeKolone` override.
- **Schema drift:** realne kolone se razlikuju po instalaciji. `modSetup.Ensure*Schema`
  self-heal-uje šemu, `modSchemaGuard.RequireColumns` fail-fast. Pre upisa proveri
  stvarne nazive kolona (`Alt+F8 → DebugKoloneTabele`).

---

## „Ko ima šta" — fleet inventory

Svaki klijent na `Workbook_Open` (`Monitor_AppOpen`) i pri proveri licence javi:
`deviceId` / `computerName`, `appVersion`, `buildVersion`, `buildSha`, `buildDate`.

- **Sirovo:** GAS `OtkupApp_Monitoring_PROD` → tab `Events`
  (kolone `AppVersion`, `BuildVersion`, `BuildSha`, `BuildDate`, `DeviceId`, `Timestamp`).
- **Agregat:** tab `Fleet` (po uređaju: poslednja `BuildVersion`/sha, kad poslednji
  put viđen, broj događaja). Puni se:
  - **automatski** — `installMonitoringTriggers()` (pokreni jednom u editoru) sad
    uključuje i `rebuildMonitoringFleet` na svaki sat;
  - **ručno** — `rebuildMonitoringFleet()` u editoru, ili action `getMonitoringFleet`
    (role management/admin/operator).
- **Klijent koji „ćuti"** u `Events`: proveri da li je monitoring uključen
  (`MONITORING_ENDPOINT` + `MONITORING_SECRET` u `tblConfig`).

---

## Min-version gate — „niko ne ostaje na staroj verziji"

Fleet ti pokaže ko zaostaje (detekcija); gate to i **sprovede** (prinuda). Na
`Workbook_Open` (`modMain.StartApp` → `modUpdateGate.UpdateGateOrQuit`) klijent
sinhrono pita GAS (`action=checkVersion`) koja je minimalna dozvoljena verzija i
poredi je sa svojom `APP_VERSION`.

- **Opt-in:** radi samo ako su `MONITORING_ENDPOINT` + `MONITORING_SECRET`
  podešeni u `tblSEFConfig` (isti uslov kao monitoring = „u floti si").
- **Fail-open:** nema interneta / server ćuti / greška → **propušta** (mreža
  nikad ne brick-uje korisnika). Pravi autoritet je server.
- **Dva nivoa:** `VERSION_ENFORCE=NO` (default) → klijent samo **upozori** i
  radi dalje; `=YES` → **blokira** start (isti mehanizam kao license blok).

**Podešavanje (GAS → Project Settings → Script Properties; menja se BEZ
redeploy-a):**

| Property | Primer | Značenje |
|---|---|---|
| `VERSION_MIN` | `2.2.0` | minimalna dozvoljena; **prazno = gate isključen** |
| `VERSION_LATEST` | `2.3.0` | samo za poruku korisniku |
| `VERSION_ENFORCE` | `NO` / `YES` | `NO` = upozorenje, `YES` = blok |
| `VERSION_MESSAGE` | (opciono) | custom tekst u dijalogu |

**Bezbedan rollout:** prvo objavi novu verziju, prati `Fleet` tab dok se flota
ne digne, pa **tek onda** podigni `VERSION_MIN` i (po potrebi) `VERSION_ENFORCE=YES`.
Ako prerano enforce-uješ minimum koji niko nema → zaključaš celu flotu (zato je
default `NO`). Poredi se `APP_VERSION` (gruba SemVer baza), ne git-describe sufiks.

---

## Caveat — `modVbaTools.FOLDER`

`ExportAllVBA`/`ImportAllVBA` imaju hardcode `FOLDER = "C:\put\do\src-vba\"`
(po mašini). Ako više ljudi/mašina radi import sa različitih putanja, kod se
lako razilazi. Preporuka: drži **jednu mašinu** kao „build" tačku i jedan
master `.xlsm`, ili izvedi `FOLDER` iz konfiguracije.

## Hitno spašavanje (sumnja da je klijent diran)

1. Na toj mašini `Alt+F8` → `ExportAllVBA` u privremen folder.
2. `git diff` prema `src-vba/` na commitu/tagu koji je klijent javio kao
   `BUILD_SHA` (`git show <sha>`).
3. Reši razliku **u gitu**, pa ponovi R2. Nikad ne ostavljaj fix samo u klijentu.

## Potpis VBA projekta (opciono)

Potpis makroa **nije** u `release.sh` (ne postoji podržan CLI za potpis VBA
projekta; uz to potpis mora **posle** `ImportAllVBA`, jer svaka izmena koda lomi
potpis). Ako ga uvodiš:

- **Kada:** poslednji korak u VBE pre snimanja (Excel korak 3): `Tools → Digital
  Signature → izaberi sertifikat`.
- **Sertifikat:** self-signed (besplatno, ali instaliraj kao *Trusted Publisher*
  na svaku mašinu, idealno preko GPO) ili komercijalni (€€ + token).
- **Korist:** tamper-detekcija (izmena koda kod klijenta lomi potpis), „samo
  potpisani makroi" security, zaobilazi Office „blocked macros" (MOTW) i AV/EDR.
- **Cena:** ručni korak svaki put + re-sign posle svake izmene; ne pokriva podatke
  ni schema drift; detekcija ≠ prevencija.

Preporuka: uvedi tek ako te Office blokira makroe pri deljenju fajla ili želiš
hardening „samo potpisani makroi". Inače su `BUILD_SHA` telemetrija +
`ExportAllVBA` / `git diff` dovoljni za detekciju izmena.

---

## Rad sa skriptom — po koracima

> Kratko, ali sve. Uz svaki korak piše GDE se radi.

### A) Priprema — SAMO JEDNOM
1. **[Git Bash]** Imaj klon repoa na mašini gde je Excel.
2. **[Excel]** U `modVbaTools` postavi `FOLDER` na `src-vba` putanju tog klona.
3. **[Browser/GAS]** Deploy `Monitoring.gs`, pa u editoru pokreni jednom: `installMonitoringTriggers()`.
4. **[Excel]** Jednom uradi `Alt+F8 → ImportAllVBA` (da uđe sva mašinerija).

### B) Svaki release (zameni `2.2.2` svojim brojem)
1. **[Git Bash]** Otvori Git Bash u folderu klona (desni klik → *Git Bash Here*), ili `cd /putanja/do/otkupapp-pwa`.
2. **[Git Bash]** `bash tools/release.sh 2.2.2`  *(pull → bump u radno stablo → **release kapija** → commit → anotiran tag `vba-v2.2.2` → push → stamp)*. Ako kapija padne, ovde staje i ništa nije tagovano.
3. **[Git Bash]** `cat src-vba/modBuildInfo.bas` → mora `BUILD_VERSION As String = "vba-v2.2.2"` (bez `+dirty`).
4. **[Excel]** Otvori **prazan build-master** `.xlsm` (master koji NE drži podatke — vidi R3 „Blanko garancija").
5. **[Excel]** `Alt+F8` → **ImportAllVBA** → Run.
6. **[Excel]** **Debug → Compile VBAProject** (mora bez greške). *Napomena: od
   uvođenja release kapije compile je već proveren u koraku 2 — ovo je potvrda nad
   build-master fajlom, ne prva provera.*
7. **[Excel]** `Alt+F8` → **AssertBlankBuild** → mora „BLANKO OK". Ako prijavi tabele s podacima → isprazni ih pa ponovi (taj fajl ide SVIMA).
7b. **[Excel]** `Alt+F8` → **PublishReleaseToDrive** (`modRelease`) → objavi `src-vba` kod + `version.json` u Drive folder `AgriX_Release` (kanal za self-update postojećih klijenata). Radi **tek pošto Compile prođe**, a **pre** koraka 9 (čita stamp-ovan `BUILD_*`). Preduslov: `REL_FOLDER_ID` postavljen u `modConfig.bas`.
8. **[Excel]** **File → Save As** → `builds\AgriX_2.2.2.xlsm` (ime prati `vba-v2.2.2`).
9. **[Git Bash]** `git checkout -- src-vba/modBuildInfo.bas` (placeholder; stamp se ne commit-uje).
10. **[bilo gde]** Pošalji `builds\AgriX_2.2.2.xlsm` klijentima (Drive / OneDrive / mejl).
11. **[Browser/GAS]** `OtkupApp_Monitoring_PROD` → tab **Fleet**: kad klijent otvori fajl, u redu vidiš `BuildVersion = vba-v2.2.2`.
12. **[uређivač]** Dopuni `docs/RELEASE_NOTES.md` — par rečenica šta je u ovom izdanju.
13. **[Browser/GAS]** *(opciono)* kad se flota digne na novu verziju: u Script Properties podigni `VERSION_MIN` (i `VERSION_ENFORCE=YES` ako želiš blok). Vidi „Min-version gate".

### Ako stane
- **„Radni direktorijum nije cist"** (korak 2) → commit-uj ili odloži izmene pa ponovi.
- **Push padne (mreža)** → ponovi `bash tools/release.sh 2.2.2` (preskoči gotovo, gurne ostatak).

---

## Distribucija preko self-update (od v2.6.1)

> **Prelaz:** `v2.6.0` je **poslednja** verzija distribuirana po starom (mejl +
> `MigrirajPodatkeIzStarog`) — jer stariji klijenti još nemaju self-update kod, ne
> mogu da ga povuku. `v2.6.0` uvodi Funkciju A (`modSelfUpdate`); od **`v2.6.1`**
> sve izmene **koda** idu preko self-update kanala, bez mejla. Detalji:
> `docs/SELF_UPDATE.md`.

### Jednokratno po klijentu (pri instalaciji v2.6.0)
Da bi self-update radio ubuduće, na svakoj mašini:
- uključi **„Trust access to the VBA project object model"** (File → Options →
  Trust Center → Trust Center Settings → Macro Settings);
- potvrdi da je **Google auth** podešen (self-update čita token; bez njega
  `Alt+F8 → DriveSelfTest` u `modDrive` prijavi prazan token);
- dodaj **AV exception** ako AV reaguje na prepisivanje VBA projekta.

### Svaki release od v2.6.1 (kod-only)
1. **[Git Bash]** razvoj na grani → merge u `main`.
2. **[Git Bash]** `bash tools/release.sh 2.6.1` *(pull → bump APP_VERSION → commit
   → tag `vba-v2.6.1` → push → stamp)*.
3. **[Excel — prazan build-master]** `Alt+F8 → ImportAllVBA` (ako baci form-grešku
   → import opet, dev quirk) → **Debug → Compile** (čisto) → `Alt+F8 →
   AssertBlankBuild` („BLANKO OK") → **`Alt+F8 → PublishReleaseToDrive`** (objavi
   kod + `version.json` u `AgriX_Release`).
4. **[Git Bash]** `git checkout -- src-vba/modBuildInfo.bas`.
5. **(opciono)** `Save As builds\AgriX_2.6.1.xlsm` — **samo za NOVE instalacije**;
   postojeći klijenti ga **ne** dobijaju (ažuriraju se sami).
6. **[GAS]** podigni `VERSION_LATEST = 2.6.1` (i, kad se flota digne, po potrebi
   `VERSION_MIN` — „Min-version gate").

**To je sve.** Klijent na sledećem `Workbook_Open` vidi „Postoji nova verzija
2.6.1 — ažurirati?" → „Da" → povuče i uveze (faza 1 code-merge + faza 2 `Import`),
napravi lokalni backup pre toga, pa traži restart. Bez mejla, bez migracije
(isti `.xlsm`; `Ensure*Schema` dovuče nove kolone posle restarta).

> **Staged rollout (preporuka):** posle `PublishReleaseToDrive` prvo testiraj na
> **jednom** klijentu; ako je dobro, ostali se sami ažuriraju. Za prinudu
> zaostalih → `VERSION_MIN`.

### Šta i dalje ide REINSTALL-om (mejl/Drive `.xlsm`, NE self-update)
Self-update povlači samo **kod**. Nov `.xlsm` (+ `MigrirajPodatkeIzStarog`) treba za:
- **izmene dizajna formi** (`.frx` / statičke kontrole) — self-update menja samo
  code-behind formi *(većina kontrola se gradi runtime-om `Controls.Add`, pa retko)*;
- **nove forme / nove sheetove** (faza 1 ih prijavi „Preskočeno, reinstall");
- izmene **`modSelfUpdate`** ili **`modVbaTools`** (skip-lista — kod koji se izvršava
  / dev tool).

> Promene **šeme** (nove kolone/tabele) NE traže reinstall — `modSetup.Ensure*Schema`
> ih self-heal-uje posle update-restarta.
