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

## Jedna komanda (rutina)

Koraci 1–4 su spojeni u skriptu (pokreni iz repo klona **na build mašini**):

```
bash tools/release.sh 2.2.2
# Windows:  powershell -ExecutionPolicy Bypass -File tools\release.ps1 2.2.2
```

Ona uradi: `checkout main` + `pull` → bump `APP_VERSION` → commit → `git tag
vba-v2.2.2` + push → `stamp-build`. Zatim ti ostaju **3 Excel klika** (Import →
Compile → Snimi) i `git checkout -- src-vba/modBuildInfo.bas`, što skripta ispiše
na kraju. Excel deo se ne automatizuje (cloud/CLI ne pokreću Excel).

## R3 — Podaci: migracija, ne kopiranje koda

Kad isporučuješ NOVU verziju koda, ne nosiš tuđe podatke u njoj:

- Klijent dobije **prazan** novi `.xlsm` (samo kod, prazne tabele).
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
2. **[Git Bash]** `bash tools/release.sh 2.2.2`  *(pull → bump APP_VERSION → commit → tag `vba-v2.2.2` → push → stamp)*
3. **[Git Bash]** `cat src-vba/modBuildInfo.bas` → mora `BUILD_VERSION As String = "vba-v2.2.2"` (bez `+dirty`).
4. **[Excel]** Otvori master `.xlsm`.
5. **[Excel]** `Alt+F8` → **ImportAllVBA** → Run.
6. **[Excel]** **Debug → Compile VBAProject** (mora bez greške).
7. **[Excel]** **Ctrl+S**.
8. **[Git Bash]** `git checkout -- src-vba/modBuildInfo.bas` (placeholder; stamp se ne commit-uje).
9. **[bilo gde]** Pošalji `.xlsm` klijentima (Drive / OneDrive / mejl).
10. **[Browser/GAS]** `OtkupApp_Monitoring_PROD` → tab **Fleet**: kad klijent otvori fajl, u redu vidiš `BuildVersion = vba-v2.2.2`.
11. **[uređivač]** Dopuni `docs/RELEASE_NOTES.md` — par rečenica šta je u ovom izdanju.

### Ako stane
- **„Radni direktorijum nije cist"** (korak 2) → commit-uj ili odloži izmene pa ponovi.
- **Push padne (mreža)** → ponovi `bash tools/release.sh 2.2.2` (preskoči gotovo, gurne ostatak).
