# Podela na .xlam (kod) + .xlsx (podaci) — implementacija

**Status:** implementirano u izvoru (`src-vba/`), čeka build/test na Windows + Excel.
**Cilj:** zajednički kod (`OtkupApp.xlam`) za sve klijente + poseban, **čist `.xlsx`**
(podaci + config) po klijentu. Program se ažurira zamenom `.xlam`-a, **bez diranja podataka**.

---

## 1. Rezultat

| Fajl | Sadržaj | Po klijentu? | Update |
|---|---|---|---|
| `OtkupApp.xlam` | sav VBA kod (moduli, klase, forme, šabloni se generišu iz koda) | ne — zajednički | zameni fajl |
| `Klijent_C00X.xlsx` | sve `tbl*` tabele (master + transakcije) + `tblConfig`/`tblSEFConfig`/`tblLocalConfig` + logo | da | nikad se ne dira pri update-u koda |

Klijentski `.xlsx` je **100% bez VBA**. Aplikaciju pokreće add-in koji sam prepoznaje
data-fajl po marker tabeli `tblConfig`.

---

## 2. Arhitektura

Isti izvorni kod podržava **dva režima** (detekcija: `ThisWorkbook.IsAddin`):

- **COMBINED** — jedan `.xlsm` (kod + podaci zajedno, kao do sada). `DataBook = CodeBook = ThisWorkbook`.
  Razvoj i dalje radi kao pre; ništa se ne kvari.
- **ENGINE** — kod je izgrađen kao `.xlam`, podaci su u `.xlsx`. `DataBook` = otvoreni klijentski `.xlsx`.

### Ključni pojmovi (`modContext.bas`)
- `DataBook` — workbook sa podacima/config-om. **Sav pristup podacima ide preko njega.**
- `CodeBook` — workbook sa kodom (sam `.xlam`). Za ugrađene asset-e (trenutno se ne koristi
  jer se svi „Sablon" sheetovi generišu iz koda u `DataBook`).
- `IsClientDataBook(wb)` — `wb` ima `tblConfig` i nije add-in.

### Pokretanje (`modBoot.bas` + `clsAppEvents.cls`)
- **ENGINE:** kad se Excel pokrene, `ThisWorkbook.Workbook_Open` add-in-a poziva
  `modBoot.InstallAppEvents`, koji kači `Application` event hook. Kada se otvori klijentski
  `.xlsx`, `clsAppEvents.App_WorkbookOpen` → `modBoot.BootFromData(wb)`. Zatvaranje →
  `App_WorkbookBeforeClose` → `ShutdownFromData`.
- **COMBINED:** `Workbook_Open` direktno zove `BootFromData(ThisWorkbook)`.
- `BootFromData` je idempotentno (guard preko `IsBoundDataBook`) i **prvo veže `DataBook`**,
  pa tek onda `Monitor_AppOpen` / `StartApp` / lock cleanup. Sva startup logika je centralizovana
  ovde (deo koda → ažurira se kroz `.xlam`).

---

## 3. Šta je promenjeno u kodu

### Novi fajlovi (`src-vba/`)
- `modContext.bas` — `DataBook`/`CodeBook`/`Bind/UnbindDataBook`/`IsClientDataBook`/`IsBoundDataBook`/`ResolveDataBook`.
- `modBoot.bas` — `InstallAppEvents`, `BootFromData`, `ShutdownFromData` (preuzeta startup/shutdown logika iz starog `ThisWorkbook`).
- `clsAppEvents.cls` — `WithEvents Application` hook (auto-detekcija klijentskog `.xlsx`).

### Izmenjeno
- `ThisWorkbook.doccls` — sada mode-aware (ENGINE: `InstallAppEvents`; COMBINED: `BootFromData`).
- **`ThisWorkbook` → `DataBook`** na svim mestima pristupa podacima/render-u/putanjama/snimanju:
  73 zamene u 20 fajlova + 3 u `modMonitoring` (telemetrijsko ime workbook-a → `DataBook.name`).
  - Keystone: `modDataAccess.GetTable` (`For Each ws In DataBook.Worksheets`) — kroz njega ide sav pristup tabelama.
  - `modSetup` (`FindListObject`, `GetOrCreateWorksheet`, `GetDefaultRootPath`, setup log).
  - Putanje/backup/log/journal: `modJournaling`, `modLogError`, `modDocStyle` (logo), `modBankaImportParserPdfToText`.
  - Snimanje: `modMain.SaveApp`, `frmOtkupAPP`, `modJournaling` (backup data-fajla).
  - Render/„Sablon" sheetovi (generišu se u `DataBook`, zbog `PrintPreview`): `modPrint`, `modPaletniList`, `modFaktura`, `modIzvestaj`, `frmSledljivost`.
- `modMonitoring.ConfigValue` — čita config iz `DataBook` (uklonjen `ActiveWorkbook` fallback).

Provera: jedini preostali `ThisWorkbook` u kodu su namerni (`modContext`, engine `doccls`, komentari, 1 telemetrijski string).

---

## 4. Build `.xlam`

```powershell
powershell -ExecutionPolicy Bypass -File build\Build-Engine.ps1
# -> dist\OtkupApp.xlam
```
Zahteva Excel + „Trust access to the VBA project object model" (skripta pokušava da postavi
`AccessVBOM=1`). Skripta importuje sve iz `src-vba/`, ubaci engine `ThisWorkbook` kod,
postavi `IsAddin=True`, snimi `.xlam`. Zatim **potpisati** `.xlam` postojećim VBA publisher
sertifikatom (VBE → Tools → Digital Signature).

## 5. Deploy / update
- `.xlam` se instalira jednom po PC-u (Excel Add-ins ili postojeća Trusted Location `C:\OtkupApp\`).
- **Update programa = zameni `OtkupApp.xlam`.** Klijentski `.xlsx` se ne dira.
- `Setup-OtkupApp.ps1` treba dopuniti da kopira/registruje `.xlam` (vidi TODO).

## 6. Novi klijent (`.xlsx`)
```powershell
powershell -ExecutionPolicy Bypass -File build\New-ClientFile.ps1 `
    -Source "C:\OtkupApp\OtkupApp.xlsm" -Output "C:\OtkupApp\Klijent_C00X.xlsx" -ClearConfigValues
```
Prazni sve poslovne `tbl*` (čuva config + kataloge), snima kao čist `.xlsx`. Operater zatim
popuni config (SELLER_*, SEF, Google…) i master podatke.

---

## 7. Pravilo za dalji razvoj
> **Nikad više `ThisWorkbook` za podatke/config/render.** Uvek `DataBook` (i `CodeBook` za
> eventualne ugrađene asset-e). `ThisWorkbook` je dozvoljen samo u `modContext` i engine `doccls`.

## 8. Rizici / TODO
- **`AccessVBOM`** mora biti uključen na build mašini (ne na klijentskim PC-evima).
- **Potpisivanje `.xlam`-a** nije automatizovano u skripti (PowerShell ne potpisuje VBA projekat lako) — uraditi u VBE ili alatom.
- **Više klijentskih `.xlsx` istovremeno otvoreno** nije podržano (single-operator dizajn); `ResolveDataBook` uzima prvi sa `tblConfig`.
- **Dupli Google secret** (zatečeni bug, nezavisan od podele): `modGoogleAuth` koristi `tblSEFConfig`, a `modSetup`/onboarding `tblConfig`. Ujednačiti na jednu tabelu.
- `Setup-OtkupApp.ps1` dopuniti za `.xlam` deploy + (opciono) auto-load preko registry-ja.

---

## 9. TEST CHECKLIST (Windows + Excel) — OBAVEZNO

> Izvor nije kompajliran u ovom okruženju (Linux, bez Excela). Pre produkcije:

**A. Regresija u COMBINED režimu (najbrža provera da ništa nije polomljeno)**
1. Build trenutni `src-vba` nazad u `.xlsm` (ili otvori postojeći radni `.xlsm` sa novim kodom).
2. VBE → Debug → **Compile** (mora proći bez greške).
3. Otvori → splash → `frmOtkupAPP`. Uradi: unos otkupa, štampa/PDF otkupnog lista,
   faktura, paletni list (i **PrintPreview**), izveštaj/kartica, banka import, save, exit.
4. Proveri da backup/journal/log nastaju **pored data-fajla**.

**B. ENGINE režim (.xlam + .xlsx)**
5. `Build-Engine.ps1` → `OtkupApp.xlam`; instaliraj kao add-in; potpiši.
6. `New-ClientFile.ps1` → `Klijent_TEST.xlsx`; popuni minimalni config (`tblConfig`/`tblSEFConfig`) i par master redova.
7. Zatvori Excel. Otvori **samo** `Klijent_TEST.xlsx` (dupli klik).
   - Očekivano: add-in se učita, `clsAppEvents` detektuje fajl, splash → forma, sve radi nad `.xlsx`-om.
8. Ponovi ključne tokove iz koraka 3 u ENGINE režimu (posebno **PrintPreview** i PDF export — render sheetovi se grade u `.xlsx`).
9. Zatvori `.xlsx` → proveri shutdown (lock release, log shutdown).
10. „Update" test: zameni `.xlam` novom verzijom; otvori isti `.xlsx`; podaci netaknuti, nova verzija radi.

**C. Edge**
11. Otvori `.xlsx` kad add-in nije instaliran → ne sme da krešuje Excel (samo se ne pokrene app).
12. Monitoring događaji nose ime klijentskog fajla (`DataBook.name`), ne `OtkupApp.xlam`.

## 10. Rollback
Podela je čisto aditivna na izvor + `ThisWorkbook→DataBook` (u COMBINED režimu identično ponašanje).
Rollback = `git revert` ove grane; `.xlsm` workflow nastavlja da radi.
