# Migracija postojećih klijenata na .xlam + .xlsx

Za klijente koji su **već počeli da pune bazu** u jednom `.xlsm`. Cilj: izvući njihove
postojeće podatke u čist `.xlsx` i preći na zajednički `OtkupApp.xlam`, **bez gubitka unosa**.

> Princip: `.xlsm` → `.xlsx` čuva SVE (sheetovi, tabele, podaci, config, formati, named ranges)
> i izbacuje samo VBA kod. Kod od sada dolazi iz `.xlam`-a. Original `.xlsm` se čuva kao backup.

---

## 0. Jednom, centralno (pre svih klijenata)
1. `build\Build-Engine.ps1` → `dist\OtkupApp.xlam`.
2. **Potpiši** `.xlam` istim VBA publisher sertifikatom (VBE → Tools → Digital Signature)
   koji `Setup-OtkupApp.ps1` već instalira na klijentskim PC-evima.
3. Ovaj `.xlam` ide na sve klijente (isti fajl za sve).

---

## 1. Po klijentu — cutover (jednokratno)

**1. Zatvori aplikaciju** na klijentskom PC-u (Exit u aplikaciji, Excel ugašen).
   Bitno: AutoSave snima posle svakog unosa, ali svejedno zatvori čisto da fajl bude
   otključan i poslednji unos na disku.

**2. Konvertuj `.xlsm` → `.xlsx`** (pravi backup + ne pokreće app):
```powershell
powershell -ExecutionPolicy Bypass -File build\Convert-XlsmToXlsx.ps1 `
    -Source "C:\OtkupApp\OtkupApp.xlsm" -Output "C:\OtkupApp\OtkupApp.xlsx"
```
Skript: backup-uje original u `C:\OtkupApp\_pre_split_backup\`, snimi `.xlsx`,
pa ga ponovo otvori i ispiše broj tabela + `tblConfig` redova radi provere.
Original `.xlsm` ostaje netaknut.

**3. Instaliraj / uključi `OtkupApp.xlam`** na PC-u:
   - kopiraj `.xlam` u `C:\OtkupApp\` (već je Trusted Location), i
   - uključi ga: Excel → File → Options → Add-ins → *Manage: Excel Add-ins* → Go… →
     Browse… → izaberi `C:\OtkupApp\OtkupApp.xlam` → OK (ostaje ✔ čekiran).

**4. Ažuriraj Desktop prečicu** da otvara `OtkupApp.xlsx` (umesto starog `.xlsm`).

**5. Otvori `OtkupApp.xlsx`** (dupli klik / prečica).
   - Add-in ga prepozna po `tblConfig`, splash → glavna forma. Aplikacija radi nad `.xlsx`-om.

**6. Verifikuj** (5 min):
   - podaci na mestu (otvori par formi: otkup, fakture, maticni);
   - probni **PrintPreview / PDF** (render sheetovi se grade u `.xlsx`);
   - pusti **RunProductionHealthCheck** (Alt+F8) — mora bez failova;
   - jedan probni unos → AutoSave → zatvori/otvori → unos i dalje tu.

**7. Tek po uspešnoj verifikaciji** ukloni stari `.xlsm` sa radne lokacije
   (backup u `_pre_split_backup\` ostaje kao sigurnost).

### Manuelna varijanta (bez skripta, za 1–2 klijenta)
Drži **Shift** dok otvaraš Excel/`.xlsm` (sprečava `Workbook_Open` da pokrene app) →
File → Save As → tip **Excel Workbook (\*.xlsx)** → potvrdi izbacivanje makroa.
Pre toga ručno kopiraj `.xlsm` u backup. Dalje isto: koraci 3–7.

---

## 2. Ažuriranje programa ubuduće (zbog čega je sve ovo)
**Zameni `OtkupApp.xlam` na PC-u novom verzijom. To je sve.**
- `.xlsx` (podaci) se NE dira.
- Ako nova verzija menja šemu tabela, dodaj `Ensure*Schema` migraciju u `modBoot.BootFromData`
  (pokreće se centralno iz koda, pa se i migracija isporučuje kroz `.xlam`).
- Distribucija: ručno kopiranje, login skripta, ili dopuni `Setup-OtkupApp.ps1` da gura `.xlam`.

---

## 3. Rollback (ako nešto ne valja posle cutover-a)
1. Ukloni/iščekiraj `OtkupApp.xlam` iz Add-ins.
2. Vrati `.xlsm` iz `_pre_split_backup\` na radnu lokaciju.
3. Vrati prečicu na `.xlsm`. Klijent radi kao pre (COMBINED režim i dalje funkcioniše).

---

## 4. Napomene / rizici
- **Schema drift između klijenata:** stariji `.xlsm`-ovi mogu imati malo drugačije tabele/kolone.
  `ValidateAllTables` (startup) i `RunProductionHealthCheck` to prijave. Po potrebi pokreni
  odgovarajući `Ensure*Schema` ili dopuni `BootFromData` migracijom pre masovnog cutover-a.
- **`tblLocalConfig` ostaje u `.xlsx`** i već je tačan za taj PC (putanje) — ne dirati.
- **Sheet dugmad sa dodeljenim makroom** (ako ih ima na listovima) postaju neaktivna u `.xlsx`
  jer nema VBA u fajlu; ova aplikacija koristi UserForme (pokreće ih `.xlam`), pa nije problem —
  ali proveri pri verifikaciji ako si negde stavljao dugmad na list.
- **Više klijentskih `.xlsx` istovremeno otvoreno** nije podržano (single-operator dizajn).
- Redosled učitavanja (add-in pre/posle `.xlsx`) je pokriven: `InstallAppEvents` skenira već
  otvorene fajlove, a `App_WorkbookOpen` hvata naknadno otvaranje.
