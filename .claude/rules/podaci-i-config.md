---
paths:
  - "src-vba/modConfig.bas"
  - "src-vba/modDataAccess.bas"
  - "src-vba/modSetup.bas"
  - "src-vba/modPodesavanja.bas"
  - "src-vba/modHelpers.bas"
  - "src-vba/modArrayUtils.bas"
  - "src-vba/sConfig.doccls"
---

# Podaci, šema tabela i config

> Preseljeno iz `CLAUDE.md` §3/§4.

## Gde šta živi (ne praviti paralele)

| Oblast | Gde |
|---|---|
| Tabele / kolone / konstante | `modConfig.bas` (`TBL_*`, `COL_*`) |
| Pristup podacima | `modDataAccess.bas` (`GetTableData` / `GetColumnIndex` / `UpdateCell` / `AppendRow` / `GetNextID` / `LookupValue`) |
| Filter/sort/util nad nizovima | `modArrayUtils.bas` (`FilterArray`, `SortArray`), `modHelpers.bas` (`Nz` / `NzToText` / `ExcludeStornirano` / `FillCmb`) |
| Setup / šeme | `modSetup` (`SetupNewPC`, `Ensure*Schema`; `SetupPopplerInteractive` / `SetupBankFoldersInteractive` pickeri; `RunSetupHealthCheck` uklj. živi `CheckServerLink` / `TestServerLink`), first-run kapija u `StartApp` (nudi `SetupNewPC` dok `APP_SETUP_COMPLETED != DA`), Admin dugmad `modAdmin` (health/googleauth/ensure), dijagnostika `DebugKoloneTabele` |

## Schema registar — jedno mesto za tabele koje aplikacija PRAVI

`modSetup.SchemaTables()` + `modSetup.SchemaOps()` su jedini izvor istine o tome
šta `Ensure*Schema` kreira. `Ensure*Schema` procedure su tanke ulazne tačke —
telo im je `ApplySchemaGroup(SG_*)`.

- **Nova tabela / nova kolona = jedan red u registru**, ne novi `EnsureDataTable`
  poziv u nekoj od procedura. Ako dodaješ poziv primitiva direktno u
  `Ensure*Schema`, staješ van registra — to je tačno ona fragmentacija koju je
  registar ukinuo.
- **Redosled redova = redosled primene, i to je semantika.** `EnsureColumnOnTable`
  dodaje kolonu na KRAJ tabele, a pozicijski `AppendRow` zavisi od redosleda
  kolona (v. dole). Ne preslaguj redove „radi urednosti".
- Dorade su namerno **dve** grupe (`SG_DORADE_SIFARNICI` / `SG_DORADE_DOKUMENTI`)
  jer se između njih zove `EnsureRuntimeSchema`; spajanje bi pomerilo runtime
  kolone ispred dorada na `tblKulture`.
- Registar je `Collection`, ne jedan `Array` literal: VBA dozvoljava najviše **25
  nastavaka reda** po naredbi.
- **Lokalni self-heal ne drži svoj spisak kolona.** Ako modul mora da popravi šemu
  usred toka (npr. `modPaletniList` pre snimanja prerade), zove
  `modSetup.EnsureTableSchema(TBL_X)` — repair-only, ne kreira tabelu koje nema.
  Lokalna kopija spiska bi ostala da leči zastarelu šemu kad se registar promeni.

### Leaf primitivi vraćaju ishod, ne oslanjaju se na `Err`

`EnsureDataTable` / `EnsureColumnOnTable` / `SetColumnNumberFormat` /
`BackfillColumn` vraćaju `SCHEMA_OK` / `SCHEMA_SKIPPED` / `SCHEMA_FAILED`.

Ne piši `On Error Resume Next` pa proveru `Err.Number` **oko poziva procedure koja
ima svoj `On Error Resume Next`** — `Err` ne preživi izlazak iz te procedure, pa je
provera mrtva i pad prolazi kao uspeh. Pozivalac čita **povratnu vrednost**.

`SCHEMA_SKIPPED` nije greška: tabele/kolone nema na ovoj instalaciji, ili je
tabela prazna pa nema data-ćelija za format. Da se to broji kao pad, svaki start
bi prijavljivao greške za funkcije koje se ne koriste. `ApplySchemaGroup` je
**jedino** mesto koje bira politiku „nastavi posle greške".

**Posle svake izmene registra ili refaktora `modSetup`-a:**

```bash
git show <ref>:src-vba/modSetup.bas > /tmp/old.bas
python3 tools/schema_diff.py /tmp/old.bas src-vba/modSetup.bas
```

Exit `0` = šema identična. Alat razume i stari zapis (inline `EnsureDataTable`
pozivi) i registar, pa radi preko refaktora. Koristi ga kad je izmena trebalo da
bude čist refaktor — ispuštena `COL_*` iz `Array` liste je tiha izmena šeme koja
se vidi tek kao prazna kolona na dokumentu kod klijenta.

## Šema tabela je izvor istine, ne kod

Realne kolone se razlikuju po instalaciji (schema drift). PRE upisa proveri
stvarne nazive kolona (`Alt+F8 → DebugKoloneTabele`). Naučeno:

- `tblStanice`: telefon je u koloni `Kontakt` (**NE** `Telefon`); kontakt =
  `Ime` / `Prezime` / `PIN`.
- `tblKulture`: `KulturaID | VrstaVoca | SortaVoca | GajbicaPoPaleti` (**NEMA**
  `Aktivan`).
- `tblOtkup` / `tblOtpremnica` / `tblPrijemnica` / `tblFakturaStavke`: količina je
  ASCII `Kolicina` (**NE** `Količina`); koristi `COL_*_KOLICINA`, ne hardkoduj
  dijakritiku (bio `RunProductionHealthCheck` bug).

**Pozicijski `AppendRow` zavisi od redosleda kolona** — bezbedan samo ako je
redosled potvrđen. Za polja čiji redosled nije siguran koristi upis **po imenu**
(`UpdateCell` / `GetColumnIndex`).

## TRI config tabele — ČITANJE i UPIS moraju u ISTU tabelu

Inače polje „ne radi" (tiho, bez greške).

| Tabela | Šta drži | API |
|---|---|---|
| `tblSEFConfig` | poslovni + **Google/PWA + SEF** kredencijali | `GetConfigValue` / `SetConfigValue` |
| `tblLocalConfig` | per-mašina: `PDFTOTEXT_EXE_PATH`, `BANKA_*_PATH`, `APP_SETUP_COMPLETED` | `GetLocalConfigValue` / `SetLocalConfigValue` |
| `tblConfig` | **legacy**, ne koristi se | — |

- Podešavanja editor rutira po `store` (`"sef"` / `"local"`) u `CfgAdd`; path polja
  imaju inline „…" browse dugme.
- Naučene greške: poppler upisan u SEFConfig a čitan iz Local;
  Google / `APP_SETUP_COMPLETED` čitani iz pogrešne tabele.
- `GetLocalConfigValue` na **praznu** vrednost vraća **default** — pa prazan
  `PDFTOTEXT_EXE_PATH` znači auto
  `<xlsm>\Tools\poppler\Library\bin\pdftotext.exe`.
