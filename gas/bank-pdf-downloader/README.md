# AgriX_C00X_BankPdfDownloader

Zaseban Apps Script projekat instaliran na **Google nalogu koji PRIMA mejlove
banke sa PDF izvodima**. Jedina odgovornost:

> Gmail PDF prilozi (od banke) → deljeni Drive folder `Bank_Izvodi`

Ne parsira PDF, ne piše u `tblBankaImport`, ne dira `tblNovac`. Sve nizvodno već
postoji u Excel/VBA klijentu i ostaje nepromenjeno.

## Gde se uklapa u postojeći lanac

```
Banka (email)
  └─[OVAJ PROJEKAT]→ Drive: Bank_Izvodi
        └─[Google Drive for Desktop sync]→ G:\My Drive\Bank_Izvodi
              └─ PullBankPdfsFromDriveProduction   (src-vba/modBankaImport.bas)
                   └─ ImportBankaInbox_TX           → tblBankaImport
                        └─ modBankaMapiranje        → tblNovac  (frmBankaImport)
```

Karika **Gmail → Drive** je jedino što je falilo; sve ostalo je već u repou
(runbook: `docs/production-runbook-banka-novac.md`).

## Model deploya (odluka)

- **Zaseban GAS projekat**, NE deo `AgriX_C00X_GAS_PROD` → nema `doPost` kolizije
  sa `gas/Code.gs`.
- **Samo dnevni time-trigger** → nema Web App / `doPost` / `BANK_IMPORT_SECRET`.
  (Web-app „skini sad" dugme iz Excela namerno izostavljeno; VBA ionako povlači iz
  Drive foldera pri svakom importu.)

## Script Properties (po klijentu — NE u izvoru)

Isti izvor ide na svaki `C00X`; razlike su u Script Properties
(`Project Settings → Script properties`):

| Property | Obavezno | Vrednost |
|---|---|---|
| `AGRIX_BANK_IZVODI_FOLDER_ID` | da | ID Drive foldera `Bank_Izvodi` koji se sinhronizuje na VBA bank inbox (`BANKA_INBOX_PATH` / `BANKA_DRIVE_SOURCE_PATH`). **Isti** folder ID koji glavni projekat drži pod ovim imenom (`gas/DriveFolder.gs` → `BANK_IZVODI`). |
| `BANK_PDF_SENDERS` | da | Email adrese banke, razdvojene zarezom/tačka-zarezom/razmakom, npr. `izvodi@banka.rs, noreply@komercijalna.rs`. |

> Folder `Bank_Izvodi` mora biti **podeljen sa ovim Gmail nalogom kao Editor**.

## Setup (jednom)

1. `script.google.com` → New project → preimenuj u `AgriX_C00X_BankPdfDownloader`.
2. Nalepi `Code.gs` iz ovog foldera.
3. Postavi Script Properties iz tabele gore.
4. `Run → testGmailAccessOnly` → odobri Gmail scope (authorize).
5. `Run → testBankPdfDownloaderConfig` → proveri `folderName`, `senders`, `query`,
   `sampleThreadCount`.
6. `Run → saveBankPdfsToProjectDrive` (ručno, jednom) → proveri da PDF-ovi padnu
   u `Bank_Izvodi`.
7. `Run → setupDailyBankPdfDownloader` → kreira dnevni trigger (07h).

## Dedupe (zašto Gmail label, ne samo ime fajla)

VBA strana **iseli** fajl iz foldera posle obrade (`Downloaded` / `Verarbeitet`),
pa provera „postoji li fajl po imenu" sama nije dovoljna — sledeći run bi ponovo
skinuo isti izvod. Zato:

- thread koji je sačuvao bar jedan PDF dobija label **`AgriX-BankPdfSaved`**;
- query isključuje taj label (`-label:AgriX-BankPdfSaved`) → izvod se skida tačno
  jednom, nezavisno od toga što ga VBA kasnije pomeri.

Provera po imenu fajla ostaje kao sekundarni štit unutar istog run-a.
Staging-dedup u Excelu (`IsDuplicateBankaImport`, `BrojDokumenta` + `BankaReferenz`)
je treći, nezavisni sloj.

## Integracioni ugovor (važno)

- Piši u **koren** foldera `Bank_Izvodi` (onaj čiji je ID u
  `AGRIX_BANK_IZVODI_FOLDER_ID`), **ne** u mesečne podfoldere (`2026/06_Jun`) —
  VBA puller čita jedan ravan folder i ne ulazi rekurzivno.
- Ime fajla uvek završava `.pdf` (garantovano u `buildStableDriveFileName_`) jer
  VBA filtrira `Dir$ "*.pdf"`.

## Funkcije

| Funkcija | Namena |
|---|---|
| `saveBankPdfsToProjectDrive` | glavna; dnevni trigger ili ručno |
| `setupDailyBankPdfDownloader` | kreira/zamenjuje dnevni trigger (07h) |
| `testBankPdfDownloaderConfig` | provera config-a + sample threadova |
| `testGmailAccessOnly` | minimalni Gmail scope smoke-test |
