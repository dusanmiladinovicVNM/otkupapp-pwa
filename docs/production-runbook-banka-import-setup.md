# Osposobljavanje Banka importa — od nule

Status: **operativni runbook za puštanje bankarskog importa u rad kod novog
(ili postojećeg) klijenta.**

Aplikacija: **OtkupApp / AgriX**
Domen: **Gmail izvod → Drive → lokalni disk → `pdftotext` → `tblBankaImport` →
`tblNovac`**

Za incidente posle puštanja (mapiranje, avansi, dedupe) vidi
`docs/production-runbook-banka-novac.md`.

---

## 0. Arhitektura (ceo lanac)

```
Banka (email)
  │  [GAS downloader]  -> na nalogu koji PRIMA izvode
  ▼
Drive: deljeni folder 01_Bank  (u AgriX_C00X_PROD)
  │  [Google Drive for Desktop sync]
  ▼
Lokalni folder  (npr. H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank)
  │  PullBankPdfsFromDriveProduction   (src-vba/modBankaImport.bas)  -> lokalni Inbox, original u Downloaded
  ▼
Lokalni Inbox  (npr. C:\...\OtkupAPP\Bank_Izvodi\Inbox)
  │  ImportBankaInbox_TX  ->  ExtractTextFromPdf (pdftotext)  ->  ParseBankaIzvodForImport (Komercijalna)
  ▼
tblBankaImport  (staging)
  │  frmBankaImport: auto-map na otvaranje (poziv/racun) + rucno
  ▼
tblNovac  (finansijska knjiga)
```

Dva odvojena Google naloga / dva GAS-a (ne mešati):

| Nalog | GAS projekat | Kod | Uloga |
|---|---|---|---|
| nalog koji **prima izvode** | `AgriX_C00X_BankPdfDownloader` | `gas/bank-pdf-downloader/Code.gs` | **Gmail → Drive** (ovaj runbook) |
| **opsagrix** (AgriX ops) | `AgriX_C00X_GAS_PROD` | `gas/Code.gs` | PWA backend, vlasnik Drive stabla, monitoring — **ne dira se** |

Veza: folder `01_Bank` je u `AgriX_C00X_PROD` (vlasnik opsagrix), pa mora biti
**podeljen nalogu koji prima izvode kao Editor** ako downloader radi na tom
nalogu.

---

## Faza 0 — Kod u produkcioni `.xlsm`

VBA izmene za banka import moraju biti u radnom `.xlsm`.
```
cd ~/Documents/GitHub/otkupapp-pwa
git fetch origin <grana>
git checkout <grana> && git pull --ff-only origin <grana>
```
Excel: `Alt+F8 → ImportAllVBA → Debug → Compile → snimi`.

Relevantni moduli: `modBankaImport`, `modBankaImportParserPdfToText`,
`modBankaMapiranje`, `modNovac`, forma `frmBankaImport`.

---

## Faza 1 — GAS downloader (Gmail → Drive)

Na **nalogu koji STVARNO prima mejlove banke**. Kod i detalji:
`gas/bank-pdf-downloader/` (`Code.gs`, `appsscript.json`, README).

1. `script.google.com` → novi projekat → `AgriX_C00X_BankPdfDownloader` → nalepi `Code.gs`.
2. `Project Settings → Show "appsscript.json" manifest` → nalepi `appsscript.json` (V8 + oauthScopes: mail, drive, script.scriptapp).
3. `Project Settings → Script properties` → `BANK_IMPORT_CLIENTS_JSON`:
   ```json
   [{
     "clientId": "C001",
     "enabled": true,
     "driveFolderId": "<ID foldera 01_Bank, samo ID a ne URL>",
     "bankSenders": ["<mejl Komercijalne>"],
     "searchDays": 30,
     "maxThreadsPerRun": 100,
     "pageSize": 50,
     "fileNamePrefix": "C001",
     "archiveAfterSave": false,
     "markReadAfterSave": false
   }]
   ```
   (opciono `BANK_IMPORT_SECRET` za WebApp; `Run → printExampleBankImportProperties` ispiše primer.)
4. Podeli Drive folder `01_Bank` **tom Gmail nalogu kao Editor**.
5. `Run → testGmailAccessOnly` → odobri Gmail scope (authorize). Očekivano `threadCount >= 0`.
6. `Run → testBankPdfImportConfig` → proveri `folderName`, `query`, `sampleThreadCount` po klijentu.
7. `Run → runBankPdfImportNow` → PDF-ovi padnu u `01_Bank`.
8. `Run → setupDailyBankPdfImportTrigger` → dnevni okidač (07h).

**Backfill istorije** (npr. od 1. januara):
1. Script Properties: `BANK_IMPORT_BACKFILL_FROM_DATE = 2026-01-01` (i po želji `BANK_IMPORT_BACKFILL_TO_DATE`).
2. `Run → runBankPdfImportBackfill`.
3. Za velik opseg digni `maxThreadsPerRun`/`pageSize` u klijent JSON-u.

---

## Faza 2 — Drive → lokalni disk (Google Drive for Desktop)

> **Klijentska mašina se NIKAD ne loguje na vlasnički `ops@agrix`** (on drži foldere/GAS-ove SVIH firmi — multi-tenant izolacija). Zato klijent uvek pristupa svom `01_Bank`-u kroz **deljeni shortcut** (mail nalog te firme), pa je putanja uvek oblika `G:\.shortcut-targets-by-id\<id>\01_Bank` — **to je normalno, ne greška.**

Na **mašini gde radi Excel** (za svakog novog klijenta isto):
1. **Podeli** klijentskom mail nalogu (kao **Editor**) foldere **`01_Bank`** i **`Downloaded`** — oba su u `00_Inbox` (struktura `00_Inbox/{01_Bank, Downloaded}`). Editor je nužan: pull **čita + briše** original iz `01_Bank` i **piše kopiju** u `Downloaded`.
2. U Drive web-u desni klik na `00_Inbox` (ili `01_Bank`) → **Add shortcut to Drive → My Drive**. Drive for Desktop ga izloži kao `G:\.shortcut-targets-by-id\<id>\01_Bank`.
3. `BANKA_DRIVE_SOURCE_PATH` = ta putanja (folder picker je razreši sam). **`BANKA_DRIVE_DOWNLOADED_PATH` ostavi PRAZNO** → default se izračuna na `…\<id>\Downloaded` (= `00_Inbox\Downloaded`).
4. (Preporučeno, ne obavezno) `00_Inbox` → desni klik → **Available offline** — pull koristi `FileSystemObject` koji radi i online-only, ali offline smanjuje hidraciju/kašnjenje.

> **FSO na pull sloju (v2.12.0+, defanziva):** od **vba-v2.12.0** su pull file/folder operacije na `Scripting.FileSystemObject` (umesto legacy `Dir$`/`MkDir`/`Name`/`FileCopy`) — robusnije na Drive virtuelnim (`.shortcut-targets-by-id`, online-only) putanjama. **Napomena:** u produkciji je 75/76 skoro uvek bilo zbog **nepokrenutog `SetupNewPC`** (nasleđene dev putanje — vidi Faza 4.5), a ne zbog shortcut-a samog; FSO je dodatna sigurnost, ne primarni lek. Ako i dalje puca posle setup-a → `ImportAllVBA` na v2.12.0+.

---

## Faza 3 — Poppler (`pdftotext`)

Parser koristi lokalni `pdftotext.exe` (Poppler).
1. Ako nije instaliran: skini „poppler-windows" release, raspakuj (layout `...\poppler-XX\Library\bin\pdftotext.exe`), kopiraj **ceo** folder (treba mu prateći DLL-ovi).
2. **Preporučeni (auto-default) raspored:** preimenuj raspakovani folder u `poppler` i stavi ga pored radne sveske, tako da putanja bude `<folder sa OtkupApp.xlsm>\Tools\poppler\Library\bin\pdftotext.exe`. Tada `PDFTOTEXT_EXE_PATH` može ostati prazan — VBA računa default relativno na radnu svesku (`Setup-OtkupApp.ps1` ovo radi automatski kopiranjem `Tools\`).
3. **Alternativa (versioned):** ostavi folder kako je (`poppler-XX`) i zapamti punu putanju do `pdftotext.exe` — nju upisuješ eksplicitno u `PDFTOTEXT_EXE_PATH` (Faza 4). Za pronalazak: `where /R C:\Users\<user> pdftotext.exe`.

> **Najlakše (bez ručnog upisa):** Podešavanja → grupa „Banka / lokalno" → dugme **„…"** pored `PDFTOTEXT_EXE_PATH` (= `SetupPopplerInteractive`): ako je poppler pored xlsm-a upiše auto-režim (prazna vrednost = relativno na svesku), inače otvori folder picker i sam nađe `pdftotext.exe` (traži i u `\Library\bin`, `\bin`, `\poppler\Library\bin`).

---

## Faza 4 — VBA lokalna konfiguracija (`tblLocalConfig`)

Dva ekvivalentna načina — izaberi jedan:

**A) Matični podaci → Podešavanja → grupa „Banka / lokalno".** Sva polja te grupe
(`PDFTOTEXT_EXE_PATH`, `BANKA_DRIVE_SOURCE_PATH`, `BANKA_INBOX/PROCESSED/ERROR_PATH`,
`BANKA_DRIVE_*`, …) sada pišu direktno u `tblLocalConfig`. Najlakše za operatera —
svako path-polje ima **inline „…" dugme** (folder picker): `PDFTOTEXT_EXE_PATH` zove
`SetupPopplerInteractive`, `BANKA_*` putanje otvaraju folder picker i upišu izbor u
polje; klik **„Sačuvaj"** persistuje. (Int/lista polja — `BANKA_DRIVE_MAX_FILES`,
`..._MIN_FILE_AGE_SECONDS`, `BANKA_AUTO_IMPORT_ON_START`, `BANKA_ALLOWED_EXTENSIONS` —
nemaju „…", to su brojevi/opcije.)

**B) Immediate (`Ctrl+G`)** ili ručni upis u `tblLocalConfig` listu:
```vba
SetLocalConfigValue "BANKA_DRIVE_SOURCE_PATH", "H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank", "Drive izvor izvoda"
SetLocalConfigValue "PDFTOTEXT_EXE_PATH", "<puna putanja do pdftotext.exe>", "Poppler pdftotext"
' BANKA_INBOX_PATH / BANKA_PROCESSED_PATH / BANKA_ERROR_PATH su lokalni
' (default C:\...\OtkupAPP\Bank_Izvodi\{Inbox,Processed,Error}).
' Za backfill podigni kapu povlacenja:
SetLocalConfigValue "BANKA_DRIVE_MAX_FILES", "500", "Backfill kapacitet"
```

> **NAPOMENA:** `PDFTOTEXT_EXE_PATH` i `BANKA_*` putanje su per-mašina i žive u
> `tblLocalConfig`. Grupa „Banka / lokalno" u Podešavanjima ih rutira tamo; ostatak
> editora i dalje piše u `tblSEFConfig`. Ako `Tools\poppler` stoji pored `OtkupApp.xlsm`,
> `PDFTOTEXT_EXE_PATH` možeš i ostaviti prazan — default se računa relativno na radnu svesku.

Provera (mora vratiti vrednosti, ne prazno):
```vba
?Dir$(GetLocalConfigValue("BANKA_DRIVE_SOURCE_PATH","") & "\*.pdf")   ' ime .pdf iz 01_Bank
?Dir$(GetLocalConfigValue("PDFTOTEXT_EXE_PATH",""))                    ' pdftotext.exe
```

---

## Faza 4.5 — Podešavanje računara i provere veze

> **KRITIČNO — `SetupNewPC` je OBAVEZAN na svakoj novoj klijentskoj mašini (nije opciono).**
> `tblLocalConfig` (per-mašina putanje: `BANKA_DRIVE_SOURCE_PATH`, `BANKA_INBOX_PATH`/Processed/Error,
> `PDFTOTEXT_EXE_PATH`, i sam `APP_SETUP_COMPLETED`) **putuje UNUTAR distribuiranog `.xlsm`-a** sa build/dev
> mašine. Ako se `.xlsm` samo prekopira a `SetupNewPC` se NE pokrene, klijent **nasleđuje dev putanje** —
> a pošto je `APP_SETUP_COMPLETED=DA` već „upečen" u fajlu, first-run kapija se **preskače** i ništa ne
> upozori. Posledica: banka uvoz puca **greškom 75/76** jer pull gađa nepostojeću putanju sa dev mašine.
> **Fix:** `Alt+F8 → SetupNewPC` (ili obriši `APP_SETUP_COMPLETED` iz `tblLocalConfig` pa restart) → sve
> putanje se resetuju na ovu mašinu. **Ovo je bio stvarni uzrok „75/76 na Drive-u" u produkciji — ne
> shortcut putanja sama po sebi.**

- **Prvi start:** na otvaranju `.xlsm`, ako računar nije podešen (`APP_SETUP_COMPLETED != DA`
  u `tblLocalConfig`), aplikacija ponudi **„Ovaj računar nije podešen — pokrenuti
  podešavanje?"** → na „Da" pokrene `SetupNewPC`. Jednokratno (posle „zelenog" setup-a
  se više ne javlja). Ručno: **Admin → „Ensure…"** ili `Alt+F8 → SetupNewPC`.
- **`SetupNewPC`** pravi lokalne foldere (uklj. `Bank_Izvodi\{Inbox,Processed,Error}`),
  validira šeme/config i upiše `APP_SETUP_COMPLETED`. **SEF je opcion** (ako sva SEF
  polja prazna, provera se preskače — ne blokira „zeleno"). **Google** kredencijali se
  čitaju iz `tblSEFConfig` (isto kao runtime); desktop-only isključi Google proveru
  preko **`EnableDesktopOnlyMode`** (`CLOUD_SYNC_ENABLED=NO`).
- **Provere (UI dugmad, bez Alt+F8) — Matični podaci → Admin:**
  - **„Health check (setup)"** (`RunSetupHealthCheck`) — folderi, poppler, Google/SEF
    config **i živi server-link** (Google OAuth token, GAS monitoring, banka Drive folder).
    Server-link je advisory: u `SetupNewPC` je samo NAPOMENA, ne obara „zeleno" (offline
    ne blokira setup).
  - **„Production health check"** (`RunProductionHealthCheck`) — dublji audit šema/integriteta.
  - **„Google autorizacija"** (`RunGoogleAuthSetup`) — jednokratni Google login (samo za
    cloud/PWA sync; banka import ga NE traži, radi 100% lokalno).
- Brza provera samo veze: `Alt+F8 → TestServerLink`.

---

## Faza 5 — Šifarnici (preduslovi za AUTO mapiranje)

Bez ovoga import radi, ali auto-map slabo hvata (sve ide ručno).
1. **Kooperanti / kupci: unet `TekuciRacun`** (`COL_KOOP_TEKUCI_RACUN` / `COL_KUP_TEKUCI_RACUN`). Format nebitan (`205-...-XX`, gole cifre, sa/bez nula) — normalizuje se. → auto-map po računu (`PartnerKonto` sa izvoda).
2. **Kooperanti: dodeljen `StanicaID`** (`COL_KOOP_STANICA`). Bez stanice se isplata **ne knjiži** — red ostaje otvoren.
3. **Ubuduće, na nalozima za plaćanje: `poziv na broj` = broj otkupnog lista** (isplata → `COL_OTK_BR_DOK`) **/ broj fakture** (uplata → `COL_FAK_BROJ`). → auto-map direktno na otkup/fakturu.

> Za istorijske izvode gde poziv na broj nije broj otkupa/fakture i ime ima
> dodatak (npr. „, ROZINA"), **tekući račun je jedini auto-ključ** — otud važnost tačke 1.

---

## Faza 6 — Uvoz (PDF → `tblBankaImport`)

1. Meni **„Banka uvoz izvoda"** (= `ImportBankaInbox_WithDrivePull`: povuče Drive→lokalni Inbox uz proveru veličine/min-age, original u `Downloaded`, parsira `pdftotext`, upiše u `tblBankaImport`).
2. Alternativa iz Immediate: `ImportBankaInbox_TX` (bez pull-a, čita lokalni Inbox) ili `ImportOnePdfIntoBankaImport "<putanja>"` (jedan fajl, dijagnostika).
3. Backfill (≫50 fajlova): `BANKA_DRIVE_MAX_FILES` diže granicu; ponavljaj uvoz dok `01_Bank` ne ostane prazan.

> Uvoz je „sve ili ništa" po batch-u: jedan PDF koji ne prođe (sken, druga banka, saldo-integrity) rollback-uje ceo batch. Izoluj krivca sa `ImportOnePdfIntoBankaImport` ili `Diag_DumpPdfTextAroundStanje`.

---

## Faza 7 — Mapiranje (`tblBankaImport` → `tblNovac`)

Otvori formu **Banka uvoz izvoda** (`frmBankaImport`):
- **Na otvaranje** se automatski mapira sve po **jakim ključevima** (poziv na broj → otkup/faktura, tekući račun) — `AutoMapStrongKeysBankaImport_TX`. Dvosmislene ostaju otvorene (bez Error).
- Kartica **„Mapirano X / Ukupno Y"** = stvarno stanje (`Obradjeno=Da` / sve staged).
- Preostale: selektuj red → „Pregled automatskog mapiranja" pokaže predlog i izvor poklapanja (`tekuci racun` / `poziv na broj`) → **„Automatski mapiraj red"** ili **„Ručno mapiraj red"**; **„Skip"** za naknade/interne.
- **„Automatski mapiraj sve"** = pun cascade (uključuje ime/PartnerMap heuristiku).

Backfill saveti:
- Prvi veliki prolaz pokreni iz Immediate na **kopiji**: `?AutoMapStrongKeysBankaImport_TX` (vrati broj mapiranih), proveri par knjiženja, pa na produkciji.
- Ručna mapiranja pune `tblPartnerMap` (uči ime→partner) za buduće auto-poklapanje po imenu.

---

## Faza 8 — Verifikacija i dnevna rutina

Verifikacija (par stavki): lanac `BIM → NOV → otkup/faktura/avans`, BIM trag u `tblNovac.Napomena` (`BIM:<id>; Ref:...; Konto:...`). Detalji i incidenti: `docs/production-runbook-banka-novac.md`.

Dnevna rutina:
1. GAS trigger u 07h puni `01_Bank`.
2. Operater otvori „Banka uvoz izvoda" → (pull+import) → auto-map na otvaranju → dovrši ručno/„Auto sve".
3. Provera „Mapirano" brojača i otvorenih stavki.

---

## 9. Redosled zavisnosti

1. Faza 0 (VBA kod) i Faza 1 (GAS) mogu paralelno.
2. Faza 2–4 (sync + Poppler + putanje) pre Faze 6.
3. **Faza 5 (računi/stanice) pre Faze 7** — bez toga auto-map nema šta da hvata.

---

## 10. Troubleshooting

| Simptom | Uzrok | Rešenje |
|---|---|---|
| Uvoz „ne daje ništa", bez greške | `BANKA_INBOX_PATH`/`BANKA_DRIVE_SOURCE_PATH` ne pokazuje na `01_Bank` (prazan inbox) | Faza 4: postavi `BANKA_DRIVE_SOURCE_PATH` na lokalnu putanju `01_Bank`; `?Dir$(...\*.pdf)` mora vratiti fajl |
| `extract error: pdftotext.exe nije pronadjen` | Poppler ne postoji ili `PDFTOTEXT_EXE_PATH` prazan/pogrešan | Faza 3 + 4: instaliraj Poppler, `SetLocalConfigValue "PDFTOTEXT_EXE_PATH", ...` (u `tblLocalConfig`, ne SEFConfig) |
| Podešeno kroz Matične podatke, i dalje ne radi (STAR build) | stariji editor je `PDFTOTEXT_EXE_PATH` pisao u `tblSEFConfig`, a import čita `tblLocalConfig` | Ažuriraj build (`ImportAllVBA`) — grupa „Banka / lokalno" sada piše u `tblLocalConfig`; ili postavi preko `SetLocalConfigValue` |
| `STANJE blok / Izvod broj nije pronadjen` | Parser je za Komercijalnu banku; drugi format | Proveri banku; `Diag_DumpPdfTextAroundStanje`; prilagođavanje parsera je zaseban posao |
| Ceo batch rollback na jednom PDF-u | „sve ili ništa" import + jedan loš PDF | `ImportOnePdfIntoBankaImport "<fajl>"` da izoluješ; izbaci/reši taj PDF |
| Preview „Auto match: Nije pronađen" iako račun postoji | forma nije reimportovana (star preview) ili račun/stanica nisu uneti | `ImportAllVBA` (sa formom); Faza 5 (unesi `TekuciRacun`/`StanicaID`) |
| Isplata se ne knjiži, red ostaje otvoren | kooperant nema `StanicaID` | Dodeli stanicu kooperantu |
| Fajlovi u `01_Bank` su „online-only" | Drive for Desktop nije materijalizovao | Preporučeno „Available offline"; od v2.12.0 pull koristi FSO pa radi i online-only |
| Pull puca **greška 75** „Path/File access" / **76** „Path not found" na Drive putanji | **Najčešće: `SetupNewPC` nije pokrenut na klijentu** → `tblLocalConfig` nosi dev putanje iz `.xlsm`-a, pa pull gađa nepostojeći folder. Ređe: nema **Editor**-a na `01_Bank`/`Downloaded`, ili STAR build sa legacy file-op-ama | **1)** `Alt+F8 → SetupNewPC` (resetuje putanje na ovu mašinu) — vidi Faza 4.5 kritičnu napomenu. **2)** Proveri **Editor** na `01_Bank`+`Downloaded`. **3)** Ažuriraj build (od **v2.12.0** su pull op-e na FSO, robusnije na shortcut putanji) |
| GAS re-skida iste izvode / `Downloaded` se puni | filename-dedupe + VBA iseli fajl iz `01_Bank` | benigno (staging-dedupe čuva tačnost); opciono Gmail label/arhiviranje u GAS-u |

---

## 11. Referentne vrednosti (primer C001)

| Stavka | Vrednost |
|---|---|
| Banka | Komercijalna (računi počinju `205`) |
| Drive izvor (`BANKA_DRIVE_SOURCE_PATH`) | `H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank` |
| Drive „downloaded" | `H:\My Drive\AgriX_C001_PROD\00_Inbox\Downloaded` |
| Lokalni Inbox/Processed/Error | `C:\Users\<user>\Desktop\OtkupAPP\Bank_Izvodi\{Inbox,Processed,Error}` |
| Poppler (`PDFTOTEXT_EXE_PATH`) | `...\Tools\poppler-XX.XX.X\Library\bin\pdftotext.exe` |
| GAS Script Property | `BANK_IMPORT_CLIENTS_JSON` (driveFolderId = ID foldera `01_Bank`) |
