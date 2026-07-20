# Self-update smoke + fault-injection (fail-closed verifikacija)

> Cilj: dokazati da **nijedan fatalni ishod ne ostavi snimljenu polu-novu /
> neoperativnu aplikaciju**. Svaki namerni kvar mora zavrsiti isto: poruka →
> **auto-close bez snimanja** → disk ostaje stara ISPRAVNA verzija → na sledecem
> startu watchdog javi „prethodni update nije zavrsen" (podaci ocuvani).
>
> Radi se na **KOPIJI** klijenta. DEV deo (RunSelfUpdateDev) ne treba Drive/auth.
> Drive deo treba Google re-auth (`Alt+F8 → RunGoogleAuthSetup`) + **test** folder
> `AgriX_Release_TEST` (ne diraj produkcijski `AgriX_Release`).

## Priprema (jednom)
1. Sveza kopija 2.21.0 → `Alt+F8 → ImportAllVBA → Debug → Compile → snimi`.
2. Zatvori i ponovo otvori (app „ziv" — kao u produkciji posle updejta).

---

## A) DEV fault-injection (lokalno, bez Drive-a) — `RunSelfUpdateDev`

Svaki test: izmeni/pokvari u **kopiji** `src-vba` foldera (ili u klonu pa vrati),
pokreni `Alt+F8 → RunSelfUpdateDev`, uporedi sa ocekivanim.

| # | Kvar (kako injektovati) | Ocekivano (FAIL-CLOSED) |
|---|---|---|
| A0 | **Nista promenjeno** (isti kod) | „Nema izmena koda" — no-op, bez teardown-a, bez crash-a. ✓ (vec potvrdjeno) |
| A1 | **Selektivna izmena** — dodaj komentar u `modKooperant.bas`, snimi | „Azurirano: 1, bez izmene: ~89" — dira SAMO taj modul, restart. (Potvrda da delta radi kad treba.) |
| A2 | **Forma ne moze merge** — u `frmOtkup.frm` dodaj u deklaracije `Private WithEvents zzTest As MSForms.CommandButton` (bez ostalog), snimi | code-merge forme padne → `needsReinstall` → „Azuriranje NIJE moguce (forma/sheet reinstall)" → **auto-close bez snimanja**. Reopen: stara verzija, watchdog. (Ovo je BAS originalni uzrok crash-a — sad fail-closed.) |
| A3 | **Faza 2 verify padne** — iz `clsAdminBtn.cls` obrisi liniju `Attribute VB_Name = "clsAdminBtn"` + dodaj komentar (da ne bude delta-skip), snimi | phase 2 Import → `ImportedOk` (ime != baseName) → „2. faza NIJE uspela … (ime='...')" → **auto-close bez snimanja**. |
| A4 | **Save padne** — izmeni jedan modul (da update krene), pa u Explorer-u stavi `.xlsm` **Read-only** (Properties → Read-only), pokreni | update prodje merge ali `SaveWorkbookVerified` = False → „Azuriranje NIJE snimljeno (read-only)" → **auto-close bez snimanja**. (Skini read-only posle.) |
| A5 | **Prekid pre faze 2** — (tesko rucno) preskoci; pokriveno watchdog-om (#A2/#A3 ostave pending pa reopen javi „prethodni update nije zavrsen"). | — |

**Posle SVAKOG A-testa (osim A0/A1):** reopen kopije → mora „**Prethodni self-update
nije zavrsen … radna verzija je ocuvana**" (watchdog) i app radi na 2.21.0. Podaci
netaknuti. Zatim vrati pokvareni fajl u ispravno stanje pre sledeceg testa.

---

## B) Drive fault-injection (treba re-auth + test folder)

Priprema: `Alt+F8 → RunGoogleAuthSetup` (token je istekao); napravi Drive folder
`AgriX_Release_TEST`, uzmi njegov ID; u `modConfig` **privremeno** `REL_FOLDER_ID` =
TEST ID (vrati posle!). Build masina: `PublishReleaseToDrive` u TEST folder.

| # | Kvar | Ocekivano |
|---|---|---|
| B1 | **Cist happy-path** — publish → na klijentu podigni `APP_VERSION` remote (ili spusti lokalni) da ponudi update → „Da" | download svih iz manifesta + **SHA-256 verifikacija** svakog → merge → restart OK. |
| B2 | **SHA nesklad (korupcija sadrzaja)** — u TEST folderu izmeni 1 bajt jednog fajla (ili u `manifest.json`/`current.json` promeni jedan `sha256`) | klijent: download → hash != manifest → fajl se ne broji → `n<>expected` → **PREKID, nista se ne menja** (log: „SHA-256 NESKLAD"). |
| B3 | **Manifest korupcija (F2)** — pokvari `manifest.json` u versioned folderu (promeni bajt) tako da `manifest_sha256` iz `current.json` ne odgovara | klijent: `manifest_sha256` provera padne PRE ijednog download-a fajla → prekid. |
| B4 | **Stale fajl** — dodaj u folder `modXYZ.bas` koji NIJE u manifestu | manifest-driven download ga IGNORISE (ne uvozi ga); update prolazi normalno. |
| B5 | **Nepotpuna objava** — u `PublishReleaseToDrive` prekini (Esc) posle par fajlova | `version.json`/`current.json` se NE upisu (gate) → klijent i dalje vidi staru verziju. |
| B6 | **Rollback (F2)** — `Alt+F8 → RollbackReleaseTo` → unesi stariju verziju | `current.json` se prepise na stariji `releases/<v>`; klijent koji jos NIJE azuriran povuce staru; vec-azuran ostaje (VersionCompare). |

**Posle B-testova:** vrati `REL_FOLDER_ID` na produkcijski u `modConfig` (i ne
commit-uj TEST ID!).

---

## Kriterijum prolaza
- Nijedan A/B fatalni test ne sme da ostavi **snimljenu** izmenu (disk = stara verzija).
- `Alt+F11` posle svakog testa: **nema duplikata** (`Module1`, `clsX1`, `modX1`),
  `Debug → Compile` cist.
- Watchdog javi prekid na reopen (kad je faza 2 zapoceta).
- Podaci (tabele) uvek netaknuti.
