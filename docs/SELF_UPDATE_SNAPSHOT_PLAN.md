# Plan: SHA-256 + versioned release folderi (pravi snapshot self-update-a)

> Nadogradnja release kanala iz „vreće fajlova" (`AgriX_Release/` flat) u
> **immutable, hash-verifikovan snapshot po verziji**. Cilj: prava sledljivost
> (svaka verzija je zamrznut folder) + laka proverljivost (klijent verifikuje
> SHA-256 svakog fajla pre importa) + trivijalan rollback (pomeri pokazivač).
>
> Nadovezuje se na već urađenu atomarnost (`docs/SELF_UPDATE.md` zamke #12–#17).
> Hash-verifikacija je novi izvor „fatalno" u postojećem atomskom toku
> (`AbortSelfUpdate` → auto-close bez snimanja).

---

## Šta ovo kupuje (i zašto povrh već urađenog)

Već imamo: klijent broji „preuzeto = očekivano" (#13) i publisher ne objavljuje
`version.json` pri delimičnom uploadu + prune (#17). To hvata **nepotpun** skup
fajlova. Ne hvata:

- **tihu korupciju sadržaja** (upload/download vratio HTTP 200 ali bajtovi
  pogrešni/oštećeni; Drive/mreža/disk) — brojanje fajlova je slepo na sadržaj;
- **„koji tačno bajtovi su u verziji X"** — flat folder se menja u mestu, nema
  zamrznutog otiska po verziji;
- **rollback** — trenutno bi značio ponovni build+publish stare verzije.

SHA-256 po fajlu rešava (1). Versioned folderi + `current.json` pokazivač rešavaju
(2) i (3).

---

## Dizajn

### Layout na Drive-u

```
AgriX_Release/                         (REL_FOLDER_ID, postojeći)
  current.json                         <- JEDINI "šta je uživo" pokazivač (flat, root)
  version.json                         <- LEGACY (zadržan u tranziciji za stare klijente)
  releases/
    2.27.0/
      manifest.json
      modConfig.bas
      frmDokumenta.frm
      ... (svi src-vba fajlovi te verzije)
    2.28.0/
      manifest.json
      ...
  <legacy flat .bas/.cls/...>           <- LEGACY (zadržani u tranziciji)
```

Svaki `releases/<verzija>/` je **snapshot po verziji** (write-once _po konvenciji_:
bump APP_VERSION za svaki release). Re-objava iste verzije **prepisuje** snapshot —
`PublishReleaseToDrive` to detektuje (postoji `manifest.json`) i **upozori**
(vbYesNo) pre prepisa; retry _neuspele_ objave prolazi bez pitanja (manifest.json
se piše tek na kraju). Stare verzije ostaju (sledljivost) uz retention (drži 10).

### `manifest.json` (po verziji)

```json
{
  "app_version": "2.28.0",
  "build_version": "vba-v2.28.0",
  "build_sha": "abc1234",
  "build_date": "2026-07-21",
  "files": [
    { "name": "modConfig.bas",   "size": 12345, "sha256": "<hex>" },
    { "name": "frmDokumenta.frm", "size": 98765, "sha256": "<hex>" }
  ]
}
```

### `current.json` (pokazivač, na rootu — atomski „go live")

```json
{
  "app_version": "2.28.0",
  "release_folder_id": "<Drive folder id za releases/2.28.0>",
  "manifest_sha256": "<hex otiska manifest.json bajtova>"
}
```

### Lanac poverenja (klijent)

```
current.json  --(app_version novije?)-->  release_folder_id
     |  manifest_sha256
     v
manifest.json  --(hash == manifest_sha256?)-->  files[]
     |  per-file sha256
     v
svaki fajl  --(download -> hash == sha256 && size?)-->  OK
```

Bilo koji nesklad na bilo kom nivou = **fatalno** (ne importuj, ne snimaj,
`AbortSelfUpdate`). `manifest_sha256` iz `current.json` sprečava da oštećen/
podmetnut manifest prokrijumčari loše hešove.

---

## Reuse (grunt u postojećem kodu — nema novih zavisnosti)

| Potreba | Postojeće (reuse) | Novo |
|---|---|---|
| SHA-256 primitiv | **`modAuth.Sha256Hex`** — `.NET SHA256Managed`, dokazan u PIN hashu; ima self-test (`modSetup:1490` proverava `Sha256Hex("abc")`) i „SHA ne radi" detekciju | `Sha256File(path)` — ista `SHA256Managed.ComputeHash_2` nad **sirovim bajtovima fajla** (ADODB.Stream binarno, isti obrazac kao `DriveDownloadToFile`) |
| Byte I/O | `modDrive` ADODB.Stream (Type=1) već čita/piše bajtove | — |
| Drive REST | `modDrive`: `DriveNewHttp`, `DriveCreateEmpty`, `DriveEscapeQ`, `DriveListFolder`, `DriveUploadFile`, `DriveDownloadToFile`, `DriveFindInFolder`, `DriveTrashFile` (nov) | `DriveEnsureFolder(parentID, name)` (find-or-create; `mimeType=application/vnd.google-apps.folder`, po obrascu `DriveCreateEmpty`) |
| JSON | `ExtractJsonStringGoogle` (modGoogleAuth) za skalare; split-po-`"}"` obrazac iz `DriveListFolder` za `files[]` niz | `ParseManifestFiles(json)` → kolekcija {name,size,sha256} |
| Verzije | `modUpdateGate.VersionCompare` | — |
| Build meta | `modBuildInfo` `BUILD_*` | — |

**`Sha256File` ide u `modDrive`** (uz byte I/O koji verifikuje; deljen build+klijent;
bez novog modula). Skica (reuse `Sha256Hex` obrasca 1:1):

```vba
' modDrive
Public Function Sha256File(ByVal path As String) As String
    On Error GoTo EH
    Dim stm As Object, sha As Object, bytes() As Byte, hash() As Byte
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open: stm.LoadFromFile path
    bytes = stm.Read: stm.Close
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    hash = sha.ComputeHash_2((bytes))
    Dim i As Long, s As String
    For i = LBound(hash) To UBound(hash)
        s = s & Right$("0" & Hex$(hash(i) And &HFF), 2)
    Next i
    Sha256File = LCase$(s)
    Exit Function
EH:
    Sha256File = vbNullString      ' "" = SHA nedostupan/greska -> pozivalac odlucuje
End Function
```

---

## Fallback kad SHA-256 nije dostupan (retko, ali dokumentovano)

`Sha256File` vraća `""` ako `.NET SHA256Managed` fali (kao što `Sha256Hex` već
radi — PIN tada pada na plaintext). Politika za update:

- Ako klijent **može** SHA (99% mašina; već proveravano za PIN) → hash-verifikacija
  je **fatalna** pri neskladu (fail-closed).
- Ako klijent **ne može** SHA → fallback na postojeći **size + count** integritet
  (#13), uz upozorenje u logu. **Ne blokiramo update** zbog nedostupnog SHA
  (update je važniji od maksimalne verifikacije; presedan: PIN plaintext fallback).

Publisher (dev mašina) **mora** imati SHA (build se prekida ako `Sha256File`
vrati `""` — manifest bez hešova se ne objavljuje).

---

## Faze

### Faza 0 — Spike (mali, GATE za ostalo) · ~0.5 dana · ✅ IMPLEMENTIRANO
> `modDrive.Sha256File` + `Test_Sha256File` (Alt+F8, poredi sa SHA256("abc")).
> Operater pokrene `Test_Sha256File` na dev + jednoj klijent mašini pre oslanjanja.
- `Sha256File` na dev mašini: hash test-fajla == `sha256sum`/`certutil` referenca.
- Potvrdi `SHA256Managed` na **jednoj klijent mašini** (već poznato-radno za PIN;
  osloni se na `Alt+F8 → TestPinHash` koji već postoji).
- Odluka: potvrđen primitiv → dalje. (Primitiv je već produkcijski, rizik nizak.)

### Faza 1 — SHA-256 u manifestu + klijentska verifikacija (flat layout) · ~1 dan · ✅ IMPLEMENTIRANO
> `modRelease`: `sha256` u `files[]` (fail ako `Sha256File=""`). `modSelfUpdate`:
> `ParseManifestFiles` + manifest-driven download + per-file SHA verify + SHA-less
> fallback + legacy (bez `files[]`) fallback. Nesklad → `n<>expected` → `AbortSelfUpdate`.
Bez strukturne promene Drive-a; „laka proverljivost" odmah.
- **Build (`modRelease`)**: za svaki fajl dodaj `sha256` (reuse `Sha256File`) u
  `files[]` (već ima name+size). Prekini objavu ako neki `Sha256File=""`.
- **Klijent (`modSelfUpdate`)**: `DownloadReleaseFiles` postaje **manifest-driven**
  — čita `files[]`, skida SAMO te fajlove, i **verifikuje `sha256` (+ size)**
  svakog. Nesklad → fatalno (`AbortSelfUpdate`). Ako SHA nedostupan → size+count
  fallback. „očekivano = broj fajlova u manifestu" (ne folder listing).
- Rezultat: tiha korupcija sadržaja se hvata; layout ostaje flat pa **stari
  klijenti i dalje rade** (čitaju `version.json` + listing).

### Faza 2 — Versioned folderi + `current.json` pokazivač · ~1.5 dana · ✅ IMPLEMENTIRANO
> `modDrive.DriveEnsureFolder` (find-or-create). `modRelease`: dual-write u
> `releases/<APP_VERSION>/` (manifest.json) + flat (`version.json`), pa **na kraju**
> `current.json` (app_version + release_folder_id + `manifest_sha256`) — atomski flip.
> `modSelfUpdate`: `GetRemoteAppVersion` prvo `current.json`; `ResolveReleaseSource`
> (versioned + `manifest_sha256` provera PRE koda, inače flat fallback);
> `DownloadReleaseFiles` koristi razrešeni folder. `manifest_sha256` nesklad = prekid.
- **Drive (`modDrive`)**: `DriveEnsureFolder(parentID, name)` (find-or-create).
- **Build (`modRelease`)**:
  1. Izračunaj manifest (Faza 1) + `manifest_sha256`.
  2. `DriveEnsureFolder(REL_FOLDER_ID,"releases")` → `DriveEnsureFolder(releases,APP_VERSION)`.
  3. Upload svih fajlova u taj folder; upload `manifest.json`.
  4. **TEK na kraju** upiši `current.json` (atomski „go live"). Bilo koji raniji
     fail → `current.json` netaknut → stara verzija ostaje uživo.
  5. **Tranzicija:** i dalje piši flat + `version.json` (dual-write) dok se flota
     ne prebaci na novi updater.
- **Klijent (`modSelfUpdate`)**: `GetRemoteAppVersion` čita **`current.json`**
  (app_version + release_folder_id); download manifest iz tog foldera, verifikuj
  `manifest_sha256`, pa Faza-1 per-file verifikacija. **Fallback:** ako
  `current.json` ne postoji (stariji publisher) → stari put (`version.json` +
  flat). Tako novi klijent radi i pre i posle cutover-a.

### Faza 3 — Cutover + rollback/retention alati · ~0.5 dana · 🟡 DELIMIČNO
> Rollback/retention/list alati **IMPLEMENTIRANI** (v Faza 2 granu). **Cutover
> (gašenje dual-write-a) NIJE** — čeka da cela flota bude na novom updateru.
- ⏳ Kad je cela flota na novom updateru (proveri Fleet/Monitoring): publisher
  prestaje da dual-write-uje flat/`version.json` (opciono ih prune-uje). **NIJE još.**
- ✅ **Alati** (u `modRelease`, uz `PublishReleaseToDrive` — ne `modAdmin`, jer su
  build/publish-side i dele njegove helpere): `RollbackReleaseTo` = prepiši
  `current.json` na stariji `releases/<verzija>` (recompute `manifest_sha256` iz
  stvarnih bajtova; bez re-build/upload); `ListReleases` = versioned verzije +
  na koju pokazuje `current.json` (sledljivost).
- ✅ **Retention**: `PublishReleaseToDrive` posle objave Trash-uje `releases/<v>`
  preko poslednjih **10** (`PruneOldReleases`, SemVer sort).

---

## Migracija / interop (kritično — redosled)

Pošto je `modSelfUpdate` u `SKIP_MODULES`, novi klijentski kod stiže **samo ručno**
(`ImportAllVBA` / nov `.xlsm`). Zato:

1. Ship Faza-1+2 **klijentski** kod na SVE mašine (ručni bootstrap; već pravilo).
2. Publisher u tranziciji **dual-write** (flat `version.json` + versioned
   `current.json`), pa i „zaostali" klijent radi.
3. Kad Fleet pokaže da su svi na novoj verziji → Faza 3 cutover (flat se gasi).

Novi klijent radi i sa starim publisher-om (`current.json` chybí → fallback na
`version.json`), i sa novim → nema „big bang" prekida.

---

## Rizici i zamke

- **SHA primitiv** — najmanji rizik (već produkcijski za PIN); ipak Faza-0 spike
  potvrđuje file-varijantu.
- **Hash round-trip = bajt-tačan.** Publisher mora hešovati **fajl koji stvarno
  uploaduje** (posle `WriteReleaseTextFile`), ne in-memory string — inače ANSI/UTF
  neslaganje. Isto važi za `manifest_sha256`: heš bajtova uploadovanog
  `manifest.json`, koje klijent skida `alt=media` (sirovo) i hešuje.
- **Re-publish iste verzije** = write-once prekršen (klijent koji tu verziju već
  ima NEĆE povući izmenjene bajtove — `VersionCompare` ne vidi razliku). Zaštita:
  `PublishReleaseToDrive` **upozorava** (vbYesNo) ako `releases/<v>` već ima
  `manifest.json`; pravilo ostaje **bump APP_VERSION za svaki release**. (Hard-reject
  bi slomio retry neuspele objave, pa je namerno upozorenje, ne blokada.)
- **`current.json` čitanje fail-soft** (offline/korupcija → nema update, kao danas).
- **Rollback ne „vraća" već ažurirane klijente** (`VersionCompare` ne dozvoljava
  auto-downgrade) — pomaže samo onima koji još nisu povukli. Dokumentovati; za
  prinudni downgrade postoji zaseban put (reinstall / `VERSION_MIN`).
- **Drive konzistentnost:** read-after-write za sadržaj je ok; `current.json` flip
  vidljiv brzo. Ne oslanjati se na redosled listinga.

---

## Test / verifikacija (bez Excela u CI — ide na kopiji/dev)

Na TEST folderu (`AgriX_Release_TEST`), preko DEV harnessa gde može:

1. **Happy path:** publish → klijent pull → svi hešovi match → import → restart OK.
2. **Korupcija sadržaja:** ručno izmeni 1 bajt jednog fajla u `releases/<v>/`
   (ili podmetni pogrešan sha u manifest) → klijent MORA fatalno (auto-close, bez
   snimanja), disk stara verzija.
3. **Manifest korupcija:** pokvari `manifest.json` → `manifest_sha256` nesklad →
   fatalno pre ijednog download-a fajla.
4. **Delimičan publish:** obori upload jednog fajla → `current.json` se NE upiše →
   klijent i dalje vidi staru verziju (nema pola-nove).
5. **SHA nedostupan:** simuliraj `Sha256File=""` → fallback na size+count, update
   prolazi uz log-upozorenje.
6. **Rollback:** `RollbackReleaseTo(stara)` → nov klijent (koji još nije updejtovan)
   povuče staru; već-updejtovan ostaje (očekivano).
7. **Dual-write interop:** star klijent (samo `version.json`) i nov klijent
   (`current.json`) rade protiv istog foldera.

---

## Sažetak izmena po fajlu

| Fajl | Faza | Izmena | Status |
|---|---|---|---|
| `modDrive` | 1/2 | `Sha256File` (nov); `DriveTrashFile` (nov); `DriveEnsureFolder` (nov) | ✅ |
| `modRelease` | 1/2/3 | `sha256` u manifestu; versioned dual-write; `current.json` flip; `PruneOldReleases` (retention 10); `RollbackReleaseTo`, `ListReleases` (Alt+F8) | ✅ (cutover ⏳) |
| `modSelfUpdate` | 1/2 | `GetRemoteAppVersion`→`current.json`; `ResolveReleaseSource` (versioned + `manifest_sha256`); `DownloadReleaseFiles`→manifest-driven + hash verify; `ParseManifestFiles`; `DownloadNamedText`; SHA fallback | ✅ |
| `docs/SELF_UPDATE.md` | sve | trust chain, layout, fallback, migracija | ⏳ |

**Procena:** ~3.5–4 dana rada + smoke/fault-injection. Faze 1 i 2 su nezavisno
mergeable (Faza 1 daje verifikaciju odmah bez rizika layout-a).

> **Napomena o smeštaju alata:** plan je predviđao `RollbackReleaseTo`/`ListReleases`
> u `modAdmin` (klijent-side admin). Implementirani su u `modRelease` jer su
> **build/publish-side** (dele `WriteReleaseTextFile`, `DriveEnsureFolder`,
> `DriveUploadFile` sa `PublishReleaseToDrive`); klijent ih nikad ne poziva, kao ni
> `PublishReleaseToDrive`. Kohezivnije uz release tooling.

## Otvorene odluke (za operatera) — REŠENO
- ✅ Retention: **10** poslednjih `releases/<v>` (`PruneOldReleases`).
- ✅ `manifest_sha256` u `current.json` **sada** (heš bajtova uploadovanog `manifest.json`).
- ✅ Faza 1+2 idu u **ISTU granu** (`claude/selfupdate-excel-crash-edxnmk`, uz hardening).
