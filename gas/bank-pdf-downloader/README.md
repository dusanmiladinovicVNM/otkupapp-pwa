# Bank PDF Gmail Downloader (multi-client)

Google Apps Script projekat instaliran na **Gmail nalogu koji STVARNO prima
mejlove banke** sa PDF izvodima. Jedina odgovornost:

> Gmail PDF prilozi (od banke) → deljeni Drive folder(i)

Ne parsira PDF, ne piše u `tblBankaImport`, ne dira `tblNovac`. Sve nizvodno
(parsiranje, saldo integritet, staging, mapiranje) ostaje u Excel/VBA klijentu.

> Ovo je **kanonska, deployovana verzija** (`Code.gs` = ono što stvarno radi na
> nalogu). Ne menjati u repou „radi lepote" — sinhronizovati sa GAS-om.

## Gde se uklapa u lanac

```
Banka (email)
  └─[OVAJ GAS]→ Drive: deljeni folder (npr. 01_Bank u AgriX_C001_PROD)
        └─[Google Drive for Desktop sync]→ H:\My Drive\AgriX_C001_PROD\00_Inbox\01_Bank
              └─ PullBankPdfsFromDriveProduction   (src-vba/modBankaImport.bas)
                   └─ ImportBankaInbox_TX           → tblBankaImport
                        └─ modBankaMapiranje        → tblNovac  (frmBankaImport)
```

## Fajlovi

| Fajl | Šta je |
|---|---|
| `Code.gs` | ceo skript (nalepiti u Apps Script projekat) |
| `appsscript.json` | manifest (timeZone, V8, oauthScopes) |

## Script Properties (per-instalacija)

`Project Settings → Script properties`:

| Property | Obavezno | Vrednost |
|---|---|---|
| `BANK_IMPORT_CLIENTS_JSON` | da | JSON **niz** klijenata (vidi dole). Multi-client: jedan GAS može puniti više foldera. |
| `BANK_IMPORT_SECRET` | samo za WebApp | dug random string; bez njega su WebApp pozivi isključeni. |
| `BANK_IMPORT_BACKFILL_FROM_DATE` | samo za backfill | `YYYY-MM-DD` (npr. `2026-01-01`). |
| `BANK_IMPORT_BACKFILL_TO_DATE` | opciono | `YYYY-MM-DD`; prazno = danas. |

`BANK_IMPORT_CLIENTS_JSON` primer:
```json
[
  {
    "clientId": "CLIENT_A",
    "enabled": true,
    "driveFolderId": "OVDE_ID_DELJENOG_FOLDERA",
    "bankSenders": ["izvodi@banka.rs", "noreply@banka.rs"],
    "searchDays": 7,
    "maxThreadsPerRun": 100,
    "pageSize": 50,
    "gmailQueryExtra": "",
    "fileNamePrefix": "CLIENT_A",
    "archiveAfterSave": false,
    "markReadAfterSave": false
  }
]
```
- `driveFolderId` = ID deljenog foldera (za C001 = folder `01_Bank`); **samo ID**, ne URL. Folder mora biti **podeljen ovom Gmail nalogu kao Editor**.
- `bankSenders` — mejlovi banke (ili koristi `gmailQueryExtra` za custom Gmail upit).
- `fileNamePrefix` — prefiks u imenu fajla (npr. `C001`), korisno kad jedan nalog puni više klijenata.
- `searchDays` — koliko dana unazad se pretražuje Gmail (**default 7**). Sa 6×/dan rasporedom (`BANK_IMPORT_TRIGGER_HOURS`) nov izvod se uhvati za par sati, pa ovaj prozor **nije za svežinu** nego je **outage buffer**: koliko dana GAS-nerada (pauza, kvota, istekla autorizacija) da se automatski nadoknadi pri oporavku. Ujedno **bounduje re-download churn** (vidi „Dedupe") — manji broj = manje duplikata u `Downloaded`. Preporuka **7** (nedelja); **3** = minimalan churn, i dalje weekend-safe. Ne stavljaj **1** (svaki prekid > 1 dan propušta). Ređi duži prekid: `runBankPdfImportBackfill` (eksplicitan opseg, ne zavisi od `searchDays`).
- Za pomoć oko vrednosti: `Run → printExampleBankImportProperties`.

## Funkcije

| Funkcija | Namena |
|---|---|
| `runBankPdfImportDaily` | dnevni trigger; svi `enabled` klijenti |
| `runBankPdfImportNow` | ručno iz editora (isto kao daily) |
| `runBankPdfImportBackfill` | backfill po `BANK_IMPORT_BACKFILL_FROM/TO_DATE` |
| `testBankPdfImportConfig` | provera config-a + Gmail/Drive (bez upisa) |
| `testGmailAccessOnly` | minimalni Gmail scope smoke-test |
| `setupDailyBankPdfImportTrigger` | kreira dnevne okidače (`BANK_IMPORT_TRIGGER_HOURS`, podrazumevano 07/08/10/12/14/16); re-run briše stare pa postavlja nove |
| `removeDailyBankPdfImportTrigger` | briše dnevni okidač |
| `doPost` | opcioni WebApp (`runNow` / `backfill` / `test`), štiti ga `BANK_IMPORT_SECRET` |
| `printExampleBankImportProperties` | ispiše primer Script Properties |

## Setup (jednom)

1. `script.google.com` → novi projekat → nalepi `Code.gs`.
2. `Project Settings → Show "appsscript.json" manifest` → nalepi `appsscript.json`.
3. Script Properties: `BANK_IMPORT_CLIENTS_JSON` (+ `BANK_IMPORT_SECRET` ako koristiš WebApp).
4. Deljeni folder podeli ovom Gmail nalogu kao **Editor**.
5. `Run → testGmailAccessOnly` → odobri Gmail scope (authorize).
6. `Run → testBankPdfImportConfig` → proveri `folderName`, `query`, `sampleThreadCount` po klijentu.
7. `Run → runBankPdfImportNow` → PDF-ovi padnu u folder(e).
8. `Run → setupDailyBankPdfImportTrigger` → dnevni okidači (07/08/10/12/14/16; menja se konstanta `BANK_IMPORT_TRIGGER_HOURS` na vrhu skripta, pa re-run ove funkcije).

## Backfill (istorija od npr. 1. januara)

1. Postavi `BANK_IMPORT_BACKFILL_FROM_DATE` (i po želji `..._TO_DATE`).
2. `Run → runBankPdfImportBackfill` → skine PDF-ove iz tog opsega (Gmail `after:`/`before:`).
3. Podigni `maxThreadsPerRun`/`pageSize` u klijent JSON-u ako je opseg velik.

## WebApp (opciono — „skini sad" iz VBA)

Deploy: `Execute as: Me`, `Who has access: Anyone with the link`. Poziv (JSON):
```json
{ "action": "runNow", "secret": "<BANK_IMPORT_SECRET>" }
```
ili `"action": "backfill"` sa `fromDate`/`toDate`/`clientId`, `"action": "test"`.

## Dedupe i integracioni ugovor

- Dedupe: stabilno ime fajla (`[prefix_]YYYY-MM-DD_<msgId>_attN_original.pdf`) + provera postojanja u folderu. Ime **uvek završava `.pdf`** (VBA puller filtrira `Dir$ "*.pdf"`).
- Piši u **koren** deljenog foldera (onaj koji se sinhronizuje na VBA bank-inbox putanju), ne u mesečne podfoldere.
- Staging-dedup u Excelu (`IsDuplicateBankaImport`, `BrojDokumenta`+`BankaReferenz`) je nezavisni drugi sloj.
- **Ponovno skidanje posle VBA povlačenja (očekivano ponašanje):** provera postojanja gleda **samo** taj koren folder. Kad `PullBankPdfsFromDriveProduction` povuče PDF, **premesti** ga u `Downloaded` (folder koji GAS nalog ne vidi), pa naredno pokretanje ne nađe fajl u korenu i **ponovo ga skine** iz Gmail-a (mejl je i dalje u `newer_than` prozoru). To je **bezopasno**: staging-dedup ga odbije, ne knjiži se dvaput. Cena je gomilanje duplikata u `Downloaded` (ime po msgID-u je isto → VBA `GetUniqueTargetPath` doda `_001`, `_002`, …). Ovo **već postoji i pri 1×/dan**; češći raspored samo umnožava pokušaje (i dalje sve odbijeno u stagingu). Svesno prihvaćeno umesto Gmail labela.
- **Ne usporava VBA:** `BankaCollectPdfFiles` enumeriše samo koren (`01_Bank`), **ne** `Downloaded`; jedini dodir je `GetUniqueTargetPath` probe, čija je dubina bounded na ~`searchDays` po fajlu. `Downloaded` sme slobodno da raste (storage-trivijalno). Ako ikad zasmeta: smanji `searchDays` (manje churn-a) ili periodično obriši stare fajlove iz `Downloaded` (to je samo arhiva već uvezenih PDF-ova).
