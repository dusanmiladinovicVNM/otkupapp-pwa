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

Na **mašini gde radi Excel**:
1. Google Drive for Desktop ulogovan na nalog koji vidi `01_Bank`. Ako je folder deljen sa drugog naloga → u Drive web-u desni klik na `01_Bank` → **Add shortcut to Drive → My Drive** (pa je pod `H:\My Drive\...`).
2. `00_Inbox` → desni klik → **Available offline** (da `pdftotext` čita realne bajtove, ne cloud placeholder).
3. Zapamti lokalnu putanju foldera (npr. `H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank`).

---

## Faza 3 — Poppler (`pdftotext`)

Parser koristi lokalni `pdftotext.exe` (Poppler).
1. Ako nije instaliran: skini „poppler-windows" release, raspakuj (layout `...\poppler-XX\Library\bin\pdftotext.exe`), kopiraj **ceo** folder (treba mu prateći DLL-ovi).
2. Nađi tačnu putanju (Command Prompt): `where /R C:\Users\<user> pdftotext.exe`.
3. Zapamti punu putanju do `pdftotext.exe` (za Fazu 4).

---

## Faza 4 — VBA lokalna konfiguracija (`tblLocalConfig`)

Immediate (`Ctrl+G`) ili upis u `tblLocalConfig` listu:
```vba
SetLocalConfigValue "BANKA_DRIVE_SOURCE_PATH", "H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank", "Drive izvor izvoda"
SetLocalConfigValue "PDFTOTEXT_EXE_PATH", "<puna putanja do pdftotext.exe>", "Poppler pdftotext"
' BANKA_INBOX_PATH / BANKA_PROCESSED_PATH / BANKA_ERROR_PATH su lokalni
' (default C:\...\OtkupAPP\Bank_Izvodi\{Inbox,Processed,Error}).
' Za backfill podigni kapu povlacenja:
SetLocalConfigValue "BANKA_DRIVE_MAX_FILES", "500", "Backfill kapacitet"
```

> **VAŽNO:** `PDFTOTEXT_EXE_PATH` postavi kroz `tblLocalConfig` (`SetLocalConfigValue`),
> **NE** kroz Matični podaci → Podešavanja — taj editor piše u `tblSEFConfig`, a
> bankarski import čita iz `tblLocalConfig`.

Provera (mora vratiti vrednosti, ne prazno):
```vba
?Dir$(GetLocalConfigValue("BANKA_DRIVE_SOURCE_PATH","") & "\*.pdf")   ' ime .pdf iz 01_Bank
?Dir$(GetLocalConfigValue("PDFTOTEXT_EXE_PATH",""))                    ' pdftotext.exe
```

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
| Podešeno kroz Matične podatke, i dalje ne radi | `PDFTOTEXT_EXE_PATH` upisan u `tblSEFConfig`, a import čita `tblLocalConfig` | Postavi ga preko `SetLocalConfigValue` |
| `STANJE blok / Izvod broj nije pronadjen` | Parser je za Komercijalnu banku; drugi format | Proveri banku; `Diag_DumpPdfTextAroundStanje`; prilagođavanje parsera je zaseban posao |
| Ceo batch rollback na jednom PDF-u | „sve ili ništa" import + jedan loš PDF | `ImportOnePdfIntoBankaImport "<fajl>"` da izoluješ; izbaci/reši taj PDF |
| Preview „Auto match: Nije pronađen" iako račun postoji | forma nije reimportovana (star preview) ili račun/stanica nisu uneti | `ImportAllVBA` (sa formom); Faza 5 (unesi `TekuciRacun`/`StanicaID`) |
| Isplata se ne knjiži, red ostaje otvoren | kooperant nema `StanicaID` | Dodeli stanicu kooperantu |
| Fajlovi u `01_Bank` su „online-only" | Drive for Desktop nije materijalizovao | Folder → Available offline |
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
