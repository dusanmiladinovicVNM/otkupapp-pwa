# OtkupApp — operativna pravila (Codebase Guardian)

> Ova pravila Claude Code učitava na startu svake sesije. Cilj: čuvati postojeću
> arhitekturu, sprečiti dupliranje, svesti izmene na najmanji potreban delta.

**Default stav:** `reuse > new` · `extend > duplicate` · `verify > conclude` ·
`inspect before propose` · `minimal change over idealized redesign`.

---

## 1) Pre svake izmene (obavezno)

1. **Reference-first.** Pogledaj izvore istine:
   - `docs/ARCHITECTURE_REFERENCE.md`, `docs/ARCHITECTURE_CHANGELOG.md`
   - `instructions/AGRIX_ARCHITECTURE_REFERENCE_FILLED_v6_12_DRAFT.md`
   - `instructions/DOMAIN_MODELS_REVIEW_DRAFT_v6_21_WITH_AGROHEMIJA.md`
2. **Pretraži postojeće** u `src-vba/` (VBA/Excel app) i `src/` (PWA) PRE nego što
   predložiš novi fajl / komponentu / hook / service / helper / tip / konstantu /
   validaciju / state / API sloj.
3. Ako ekvivalent (ili delimičan ekvivalent) postoji — **koristi ga ili proširi
   minimalno**. Novo uvodi SAMO uz jasan razlog (postojeće objektivno ne podržava
   scenario; proširenje bi napravilo veći tehnički dug). Ako nešto nije provereno,
   reci da nije provereno — ne popunjavaj rupe pretpostavkama.

## 2) Anti-duplication

Ne praviti paralelnu implementaciju za nešto što već postoji; ne uvoditi novi
naming ako naming pattern već postoji; ne praviti novi shared helper ako sličan
postoji; ne uvoditi novi sloj apstrakcije bez jasnog razloga („rule of three").

## 3) Mapa koda (gde šta živi — ne praviti paralele)

| Oblast | Gde |
|---|---|
| Tabele / kolone / konstante | `modConfig.bas` (`TBL_*`, `COL_*`) |
| Pristup podacima | `modDataAccess.bas` (`GetTableData/GetColumnIndex/UpdateCell/AppendRow/GetNextID/LookupValue`) |
| Filter/sort/util nad nizovima | `modArrayUtils.bas` (`FilterArray`, `SortArray`), `modHelpers.bas` (`Nz/NzToText/ExcludeStornirano/FillCmb`) |
| Maticni podaci (UI) | `frmMaticniPodaci` + `frmStammdaten` (`Select Case Me.Tag`) + `modMaticniLookups` (data-driven meni) |
| Otkup / dokumenta | `frmOtkup`+`modOtkup`, `frmDokumenta`+`modDokumenta` |
| Ambalaza ledger | `modAmbalaza` · **Cenovnik (append-only):** `modCenovnik` (`GetVazecaCena/AddCena`) |
| Cena — DVA modela (ne mešati) | single-current po artiklu = `tblArtikli.CenaPoJedinici` (inline `LookupValue`, agrohemija); append-only istorija za otkup voća = `tblCenovnik` |
| Dinamičke kontrole (bez `.frx`) | `Controls.Add` + WithEvents klasa (`clsBlokUI`/`modOtkupBlok`, `clsLookupMenuBtn`/`modMaticniLookups`) |
| Sync / PWA | `modStammdatenSync`, `modMasterSync`, `gas/` |
| Setup / šeme | `modSetup` (`EnsureDataTable`, `EnsurePaletniListSchema`, `EnsureCenovnikSchema`), dijagnostika `DebugKoloneTabele` |

## 4) VBA / Excel — specifična pravila (naučeno)

- **Ne zaključuj iz par linija.** Logika je raspoređena kroz module/forme/klase/
  evente — traži pun set relevantnih (`frm*`, `mod*`, `cls*`, `ThisWorkbook`) pre
  procene reuse-a / refaktora.
- **Šema tabela je izvor istine, ne kod.** Realne kolone se razlikuju po
  instalaciji (schema drift). PRE upisa proveri stvarne nazive kolona
  (`Alt+F8 → DebugKoloneTabele`). Primeri naučeni:
  - `tblStanice`: telefon je u koloni `Kontakt` (NE `Telefon`); kontakt = `Ime`/`Prezime`/`PIN`.
  - `tblKulture`: `KulturaID | VrstaVoca | SortaVoca | GajbicaPoPaleti` (NEMA `Aktivan`).
- **Pozicijski `AppendRow` zavisi od redosleda kolona** — bezbedan samo ako je
  redosled potvrđen. Za polja čiji redosled nije siguran koristi upis **po imenu**
  (`UpdateCell`/`GetColumnIndex`).
- **Kontrole formi su u binarnom `.frx`** (`.frm` ima samo `OleObjectBlob`). Nove
  kontrole se NE dodaju editovanjem teksta — dodaj ih u runtime-u
  (`Controls.Add` + WithEvents), kao postojeći `modOtkupBlok`/`clsBlokUI`.
- Pri merge-u: čist git-merge može dati VBA compile grešku (dupli `Public`
  `Sub`/`Function`/`Const` → „Ambiguous name"). Posle merge-a uradi
  `Debug → Compile VBAProject` i proveri duple definicije.
- **Encoding: VBA fajlovi (`.bas`/`.cls`/`.frm`) su jednobajtni ANSI —
  Windows-1250 (srednjoevropski; ima `š/ž/č/ć`) + LF — NE UTF-8.** Edit/Write alat
  ih pri snimanju prebaci u UTF-8 i **tiho ošteti** sve `š/ž/č/ć` i nemačke znake —
  uključujući `MsgBox` poruke vidljive korisniku (gubitak je nepovratan jer više
  bajtova mapira na isti `�`). Već se dešavalo (vidi `f08a0ee`). Pravilo (da se NE
  ispravlja stalno posle greške):
  - PRE editovanja proveri `file <fajl>` → „Non-ISO extended-ASCII" = jednobajtni ANSI.
  - Ako fajl ima ne-ASCII znake, **NE diraj ga Edit/Write alatom direktno.**
    Primeni izmenu **byte-preserving latin-1 round-trip skriptom** (Python
    `open(..., encoding='latin-1', newline='')` za čitanje i pisanje — latin-1 čuva
    bajtove bez obzira na 1250/1252; **očuvaj originalni EOL**, u ovom repo-u LF), a
    sopstvene dodatke drži **ASCII-only** (bez dijakritike, npr. „pocetno").
    Po potrebi: `git checkout HEAD -- <fajl>` pa re-apliciraj izmene tom skriptom.
  - POSLE izmene: `file` MORA i dalje da kaže „Non-ISO extended-ASCII" (ne „UTF-8"),
    a `git diff` sme da pokaže **samo namerne linije** (bez `�` šuma na netaknutim).
  - `.md` / `.js` / ostali UTF-8 fajlovi: Edit/Write je bezbedan (nema konverzije).

## 5) Verifikacija (CI ne pokreće Excel)

- VBA se ne kompajlira/pokreće u ovom okruženju. Verifikuj **statički**: balans
  `Sub`/`Function`/`Select Case`, nema duplih `Public` definicija, `git merge-tree`
  za konflikte. Finalni smoke-test radi korisnik u Excelu.
- Forme: izmene su u kodu; `.frx` se ne dira. Pri re-importu, `.frm` ide sa svojim
  `.frx` parom.

## 6) Git / PR

- Razvoj na zadatoj feature grani; commit poruke jasne i opisne. Ne praviti PR bez
  eksplicitnog zahteva. Pre merge-a u `main` proveriti konflikte (`git merge-tree`)
  i preklapanja fajlova.
- **Integracija ažuriranog `main`-a u feature granu = UVEK „Opcija 3":** `fetch` →
  proveri preklapanja + `git merge-tree` → **rebase lokalno** na `origin/main` →
  pokaži rezultat (log, diff vs `main`, statičke provere) → `push --force-with-lease`
  TEK po eksplicitnom odobrenju. Nikad force-push pre pokazivanja.
- **Posle kreiranja PR-a ka `main`:** podseti korisnika na release/verzionisanje —
  `tools/release.sh <verzija>` → Excel `ImportAllVBA` → `Compile` → snimi → ship →
  `Fleet` provera, da se novi `AgriX_OtkupApp.xlsm` pravilno verzioniše. Vidi
  `docs/RELEASE_PROCEDURE.md` i dopuni `docs/RELEASE_NOTES.md`.
- **Na kraju SVAKE izmene koda (posle commit/push):** UVEK daj git bash komandu za
  preuzimanje feature grane radi testa kroz `ImportAllVBA`. Lokalni klon je
  `~/Documents/GitHub/otkupapp-pwa` (= `ImportAllVBA` folder). Šablon:
  `cd ~/Documents/GitHub/otkupapp-pwa` → `git fetch origin <grana>` →
  `git checkout <grana>` → `git pull --ff-only origin <grana>`; zatim u Excelu
  `Alt+F8 → ImportAllVBA → Debug → Compile → snimi → test`.
- **Uz to, svaki rad završi kratkom test-checklistom direktno u chatu:** numerisani,
  konkretni koraci šta operater proba u Excelu (klik po klik + očekivani rezultat),
  fokusiran na ono što je u toj izmeni dodato/promenjeno. Kratko i praktično.

---

_Detaljnu „Codebase Guardian" doktrinu (reference-first, anti-duplication, format
odgovora po sekcijama) primenjivati i kad nije eksplicitno ponovljena u promptu._
