# AgriX onboarding vodič za novog klijenta

**Dokument:** Standardna priprema i instalacija AgriX sistema za novog klijenta  
**Model:** `ops@agrix.rs` + Google Drive/Sheets/GAS/PWA + lokalni Excel/VBA + SEF + bankarski import  
**Primena:** C001, C002, C003...  
**Status:** Operativni vodič zasnovan na uspešno stabilizovanom C001 putu  
**Ažurirano:** 2026-05-13
**Revizija:** v2 — dodati digitalni potpis, certifikat, Trusted Location, PS1 install pravila, release manifest

---

## 0. Svrha dokumenta

Ovaj dokument opisuje tačan postupak za pripremu novog AgriX klijenta.

Vodič je podeljen u dve faze:

1. **Faza 1 — priprema u kancelariji**
   - sve što možeš da uradiš bez prisustva klijenta;
   - Google Drive struktura;
   - Apps Script / GAS;
   - OAuth;
   - Sheets;
   - PWA povezivanje;
   - Excel/VBA konfiguracija;
   - full sync;
   - install package;
   - tehnički smoke test.

2. **Faza 2 — aktivnosti kod klijenta**
   - lokalna instalacija na računaru;
   - SEF setup;
   - bankarski email/PDF tok;
   - završni smoke test;
   - predaja sistema korisniku.

Cilj je da se kod svakog novog klijenta izbegnu greške koje su se pojavile tokom C001 setup-a, posebno:

- pogrešan folder za `Stammdaten`;
- pogrešan Google browser profil;
- stari spreadsheet ID-jevi za `Kartice` i `MgmtReports`;
- neusklađen `GAS Web App URL` i `PWA API_URL`;
- izostavljen `LoginLog`;
- nejasan redosled full sync-a i PWA testiranja;
- zaboravljen bankarski email/PDF tok;
- nepotpun install package.

---

## 1. Standardne oznake

Za svakog klijenta koristi se oznaka:

```text
CLIENT_ID = C00X
```

Primeri:

```text
C001
C002
C003
```

Za produkciju koristi:

```text
ENV = PROD
APP_VERSION = 1.0.0-C00X
```

Primer za drugog klijenta:

```text
CLIENT_ID = C002
ENV = PROD
APP_VERSION = 1.0.0-C002
```

Standardni nazivi:

```text
AgriX_C00X_PROD
AgriX_C00X_GAS_PROD
AgriX_C00X_Monitoring_PROD
AgriX_C00X_Install_v1.0.0
```

---

## 2. Osnovna arhitektura

Za sada, bez Google Workspace-a, koristi se ovaj model:

```text
ops@agrix.rs
  - Google Account napravljen preko Loopia adrese
  - vlasnik / admin Google Cloud projekta
  - vlasnik / admin GAS projekta
  - vlasnik / admin Drive/Sheets strukture
  - deployment kontrola

backup@agrix.rs
  - rezervni pristup
  - editor na Drive folderima
  - editor na Apps Script projektu

app.agrix.rs
  - univerzalni PWA frontend
  - config.js pokazuje na aktivni GAS Web App URL za trenutnog klijenta

OtkupApp.xlsm
  - lokalni Excel/VBA sistem
  - čita/piše Google Sheets preko OAuth-a
  - radi full sync
  - radi SEF
  - radi bankarski PDF import
```

Važno pravilo:

```text
Ne koristiti lični/multi-account Google browser profil za AgriX administraciju.
```

Za administraciju koristi poseban browser profil:

```text
Browser profil: AgriX OPS
Google nalog: ops@agrix.rs
```

U tom profilu radiš:

```text
Google Drive
Google Cloud Console
Apps Script
GAS deployment
PWA ping test
config.js proveru
```

---

# FAZA 1 — PRIPREMA U KANCELARIJI

---

## 3. Pre-start checklist

Pre nego što počneš sa novim klijentom, potvrdi:

```text
[ ] ops@agrix.rs Google Account radi.
[ ] backup@agrix.rs Google Account radi.
[ ] Imaš pristup Loopia emailovima.
[ ] Imaš pristup GitHub/repo kodu.
[ ] Imaš poslednji stabilan OtkupApp.xlsm.
[ ] Imaš poslednji stabilan GAS Code.gs / Monitoring.gs / DriveFolder.gs.
[ ] Imaš poslednji stabilan Setup-OtkupApp.ps1.
[ ] Imaš Poppler paket.
[ ] Imaš AgriX OPS browser profil.
[ ] Znaš novi CLIENT_ID.
```

---

## 4. Otvori interni klijentski zapis

Za svakog klijenta napravi interni zapis, na primer u password manager-u ili lokalnom sigurnom dokumentu:

```text
CLIENT_ID:
CLIENT_NAME:
ENV:
APP_VERSION:

Drive root:
GAS project:
GAS Web App URL:
PWA API_URL:

OAuth client ID:
OAuth client secret:

Stammdaten ID:
Kartice ID:
MgmtReports ID:
Monitoring ID:

SEF_API_KEY status:
Bank email status:

Install package:
Install date:
Notes:
```

Tajne vrednosti ne stavljati u GitHub.

Tajne vrednosti su:

```text
GOOGLE_CLIENT_SECRET
MONITORING_INGEST_SECRET
SEF_API_KEY
lozinke
backup codes
```

---

## 5. Google Drive struktura

Kao `ops@agrix.rs`, u Google Drive-u napravi root folder:

```text
AgriX_C00X_PROD
```

Unutra napravi kompletnu strukturu:

```text
AgriX_C00X_PROD/
  00_Inbox/
    Bank/
    Fiskalni/
    Uvoz/
    Manual/

  01_Sheets/
    01_Operational/
    02_Master/
    03_Reports/
    04_Archive/

  02_Bank_Izvodi/
    2026/
      01_Januar/
      02_Februar/
      03_Mart/
      04_April/
      05_Maj/
      06_Jun/
      07_Jul/
      08_Avgust/
      09_Septembar/
      10_Oktobar/
      11_Novembar/
      12_Decembar/

  03_Documents/
    Otkupni_Listovi/
    Otpremnice/
    Zbirne/
    Fakture/

  04_Export/
    Excel/
    PDF/
    CSV/
    API/

  05_Backup/
    Daily/
    Weekly/
    Before_Release/

  06_Monitoring/
    ErrorLog/
    Sync_Reports/
    Incidenti/
    Health_Checks/

  07_Admin/
    Config/
    Deployments/
    Access/
    Templates/
    Runbooks/
```

> **Automatizacija (preporučeno):** ručno pravljenje ovih podfoldera (ovaj korak) +
> skupljanje folder ID-jeva (§8) + upis u Script Properties (§11) možeš zameniti jednom
> GAS funkcijom. Napravi **samo root** `AgriX_C00X_PROD` (i podeli ga sa backup nalogom, §6),
> pa kad postaviš GAS projekat (§10) pokreni `bootstrapAgriXFolderTree` iz `DriveFolder.gs` —
> on napravi celo stablo i upiše svih 33 folder ID-ja odjednom. Ručni postupak ispod
> (§8, §11) ostaje kao referenca/fallback.

---

## 6. Share pristup za backup nalog

Root folder:

```text
AgriX_C00X_PROD
```

podeli sa:

```text
backup@agrix.rs
```

kao:

```text
Editor
```

Provera:

```text
[ ] Uloguj se kao backup@agrix.rs.
[ ] Otvori AgriX_C00X_PROD.
[ ] Napravi test Google Sheet.
[ ] Obriši test Google Sheet.
[ ] Potvrdi da backup nalog ima realan edit pristup.
```

---

## 7. Pravilo gde šta ide

Ovo je jedno od najvažnijih pravila.

```text
Stammdaten
→ 01_Sheets/02_Master

Kartice
→ 01_Sheets/03_Reports

MgmtReports
→ 01_Sheets/03_Reports

OTK-*
VOZ-*
AGRO-*
TRETMAN-*
OPREMA-*
TROSKOVI-*
FISKALNI-*
→ 01_Sheets/01_Operational

ErrorLog
LoginLog
→ 06_Monitoring/ErrorLog

Monitoring workbook
→ 06_Monitoring
```

Posebno važno:

```text
Stammdaten ne sme ostati direktno u 01_Sheets.
Mora biti u 01_Sheets/02_Master.
```

Ovo je bio uzrok C001 login problema: GAS je tražio `Stammdaten` u `02_Master`, dok ga je VBA napravio direktno u `01_Sheets`.

---

## 8. Prikupljanje folder ID-jeva

> Ako koristiš `bootstrapAgriXFolderTree` (vidi §5/§11), ovaj korak preskačeš — funkcija
> sama upiše sve ID-jeve u Script Properties. Ispod je ručni postupak.

Za svaki folder otvori ga u browseru i uzmi ID iz URL-a.

Primer:

```text
https://drive.google.com/drive/folders/1AbCdEfGh...
```

Folder ID je:

```text
1AbCdEfGh...
```

Popuni evidenciju:

```text
AGRIX_ROOT_FOLDER_ID =
AGRIX_INBOX_FOLDER_ID =

AGRIX_SHEETS_FOLDER_ID =
AGRIX_SHEETS_OPERATIONAL_FOLDER_ID =
AGRIX_SHEETS_MASTER_FOLDER_ID =
AGRIX_SHEETS_REPORTS_FOLDER_ID =
AGRIX_SHEETS_ARCHIVE_FOLDER_ID =

AGRIX_BANK_IZVODI_FOLDER_ID =

AGRIX_DOCUMENTS_FOLDER_ID =
AGRIX_DOC_OTKUPNI_LISTOVI_FOLDER_ID =
AGRIX_DOC_OTPREMNICE_FOLDER_ID =
AGRIX_DOC_ZBIRNE_FOLDER_ID =
AGRIX_DOC_FAKTURE_FOLDER_ID =

AGRIX_EXPORT_FOLDER_ID =
AGRIX_EXPORT_EXCEL_FOLDER_ID =
AGRIX_EXPORT_PDF_FOLDER_ID =
AGRIX_EXPORT_CSV_FOLDER_ID =
AGRIX_EXPORT_API_FOLDER_ID =

AGRIX_BACKUP_FOLDER_ID =
AGRIX_BACKUP_DAILY_FOLDER_ID =
AGRIX_BACKUP_WEEKLY_FOLDER_ID =
AGRIX_BACKUP_BEFORE_RELEASE_FOLDER_ID =

AGRIX_MONITORING_FOLDER_ID =
AGRIX_MONITORING_ERRORLOG_FOLDER_ID =
AGRIX_MONITORING_SYNC_REPORTS_FOLDER_ID =
AGRIX_MONITORING_INCIDENTI_FOLDER_ID =
AGRIX_MONITORING_HEALTH_CHECKS_FOLDER_ID =

AGRIX_ADMIN_FOLDER_ID =
AGRIX_ADMIN_CONFIG_FOLDER_ID =
AGRIX_ADMIN_DEPLOYMENTS_FOLDER_ID =
AGRIX_ADMIN_ACCESS_FOLDER_ID =
AGRIX_ADMIN_TEMPLATES_FOLDER_ID =
AGRIX_ADMIN_RUNBOOKS_FOLDER_ID =
```

---

## 9. Google Cloud / OAuth

Ako već postoji stabilan AgriX platform OAuth, koristi ga.

Standard:

```text
Google Cloud Project:
AgriX_PLATFORM_PROD

OAuth Client:
AgriX_VBA_Desktop_PROD

Type:
Desktop app
```

Potrebni API-jevi:

```text
[ ] Google Drive API
[ ] Google Sheets API
```

Sačuvaj:

```text
GOOGLE_CLIENT_ID
GOOGLE_CLIENT_SECRET
```

u password manager.

U Excel `tblConfig` će ići:

```text
GOOGLE_CLIENT_ID
GOOGLE_CLIENT_SECRET
```

---

## 10. Apps Script projekat

Kao `ops@agrix.rs` napravi Apps Script projekat:

```text
AgriX_C00X_GAS_PROD
```

U projekat ubaci najmanje:

```text
Code.gs
Monitoring.gs
DriveFolder.gs
```

U `DriveFolder.gs` moraju postojati helper-i za folder mapu:

```text
AGRIX_FOLDER_PROPS
getAgriXFolder_
getSpreadsheetByNameInFolder_
createSpreadsheetInFolder_
getOrCreateSpreadsheetInFolder_
getStammdatenSpreadsheet_
getOperationalSpreadsheet_
getReportSpreadsheet_
getMonitoringErrorLogSpreadsheet_
getLoginLogSpreadsheet_

AGRIX_FOLDER_TREE
getOrCreateChildFolder_
buildAgriXTree_
bootstrapAgriXFolderTree
```

`bootstrapAgriXFolderTree` je one-time bootstrap za novog klijenta: napravi celo Drive
stablo (§5) i upiše svih 33 folder ID-ja u Script Properties (§8 + §11) u jednom run-u.
Pokreni ga u GAS projektu **tog** klijenta (Script Properties su per-projekat); idempotentan je.

Obavezno dodati `getLoginLogSpreadsheet_()`:

```javascript
function getLoginLogSpreadsheet_() {
  return getOrCreateSpreadsheetInFolder_(
    'MONITORING_ERRORLOG',
    'LoginLog',
    ['Timestamp', 'Username', 'EntityID', 'Success', 'Message']
  );
}
```

Bez ove funkcije login može raditi, ali se `LoginLog` neće napraviti, jer login logging ne sme da sruši auth flow i greške se tiho progutaju.

---

## 11. Script Properties u Apps Script

> **Brži put:** umesto ručnog upisa folder ID-jeva ispod, pokreni `bootstrapAgriXFolderTree`
> (`DriveFolder.gs`, §10) — napravi stablo i upiše svih 33 `AGRIX_*_FOLDER_ID` propsa odjednom.
> Posle njega idi pravo na §13 (`debugAgriXFolders`) za proveru. Lista ispod je referenca šta
> mora da postoji (i za ručni upis). `MONITORING_*` tajne (§12) se i dalje upisuju ručno.

Idi na:

```text
Apps Script > Project Settings > Script Properties
```

Unesi sve folder ID-jeve.

Minimum koji mora postojati:

```text
AGRIX_ROOT_FOLDER_ID
AGRIX_INBOX_FOLDER_ID

AGRIX_SHEETS_FOLDER_ID
AGRIX_SHEETS_OPERATIONAL_FOLDER_ID
AGRIX_SHEETS_MASTER_FOLDER_ID
AGRIX_SHEETS_REPORTS_FOLDER_ID
AGRIX_SHEETS_ARCHIVE_FOLDER_ID

AGRIX_BANK_IZVODI_FOLDER_ID

AGRIX_DOCUMENTS_FOLDER_ID
AGRIX_DOC_OTKUPNI_LISTOVI_FOLDER_ID
AGRIX_DOC_OTPREMNICE_FOLDER_ID
AGRIX_DOC_ZBIRNE_FOLDER_ID
AGRIX_DOC_FAKTURE_FOLDER_ID

AGRIX_EXPORT_FOLDER_ID
AGRIX_EXPORT_EXCEL_FOLDER_ID
AGRIX_EXPORT_PDF_FOLDER_ID
AGRIX_EXPORT_CSV_FOLDER_ID
AGRIX_EXPORT_API_FOLDER_ID

AGRIX_BACKUP_FOLDER_ID
AGRIX_BACKUP_DAILY_FOLDER_ID
AGRIX_BACKUP_WEEKLY_FOLDER_ID
AGRIX_BACKUP_BEFORE_RELEASE_FOLDER_ID

AGRIX_MONITORING_FOLDER_ID
AGRIX_MONITORING_ERRORLOG_FOLDER_ID
AGRIX_MONITORING_SYNC_REPORTS_FOLDER_ID
AGRIX_MONITORING_INCIDENTI_FOLDER_ID
AGRIX_MONITORING_HEALTH_CHECKS_FOLDER_ID

AGRIX_ADMIN_FOLDER_ID
AGRIX_ADMIN_CONFIG_FOLDER_ID
AGRIX_ADMIN_DEPLOYMENTS_FOLDER_ID
AGRIX_ADMIN_ACCESS_FOLDER_ID
AGRIX_ADMIN_TEMPLATES_FOLDER_ID
AGRIX_ADMIN_RUNBOOKS_FOLDER_ID
```

---

## 12. Monitoring spreadsheet

U folderu:

```text
AgriX_C00X_PROD/06_Monitoring
```

napravi spreadsheet:

```text
AgriX_C00X_Monitoring_PROD
```

Uzmi spreadsheet ID.

U Script Properties dodaj:

```text
MONITORING_SPREADSHEET_ID = <ID>
MONITORING_ALERT_EMAIL = alerts@agrix.rs
MONITORING_INGEST_SECRET = <duga_tajna_vrednost>
```

`MONITORING_INGEST_SECRET` čuvati u password manager-u.

---

## 13. GAS debug funkcije

Pre deploy-a pokreni `debugAgriXFolders`.

Primer funkcije:

```javascript
function debugAgriXFolders() {
  const keys = Object.keys(AGRIX_FOLDER_PROPS);
  const out = [];

  keys.forEach(function(key) {
    try {
      const folder = getAgriXFolder_(key);
      out.push(key + ' OK: ' + folder.getName() + ' | ' + folder.getId());
    } catch (err) {
      out.push(key + ' FAIL: ' + err.message);
    }
  });

  Logger.log(out.join('\n'));
  return out;
}
```

Očekivanje:

```text
Svi folderi moraju biti OK.
Nema Missing Script Property.
Nema permission error-a.
```

Dodaj i test za `doGet`:

```javascript
function debugDoGetPing() {
  const res = doGet({
    parameter: {
      action: 'ping'
    }
  });

  Logger.log(res.getContent());
}
```

Očekivanje:

```json
{"success":true,"timestamp":"..."}
```

---

## 14. GAS deploy

Deploy:

```text
Deploy > New deployment > Web app
```

Podešavanja:

```text
Execute as: Me (ops@agrix.rs)
Who has access: Anyone
```

Kopiraj Web App URL:

```text
https://script.google.com/macros/s/AKfycb.../exec
```

Test:

```text
<GAS_WEB_APP_URL>?action=ping
```

Očekivanje:

```json
{"success":true,"timestamp":"..."}
```

Ako ping radi u jednom browser profilu, a ne radi u drugom, problem je Google multi-account session. Za AgriX koristi čist browser profil `AgriX OPS`.

---

## 15. Excel `tblConfig`

U `OtkupApp.xlsm`, u `tblConfig`, postavi:

```text
Kljuc                         Vrednost

GOOGLE_CLIENT_ID              <OAuth client id>
GOOGLE_CLIENT_SECRET          <OAuth client secret>

GOOGLE_PWA_FOLDER_ID          <ID od 01_Sheets/02_Master>
GOOGLE_REPORTS_FOLDER_ID      <ID od 01_Sheets/03_Reports>

GOOGLE_STAMMDATEN_SHEET_ID    prazno
GOOGLE_KARTICE_SHEET_ID       prazno
GOOGLE_MGMT_SHEET_ID          prazno

CLIENT_ID                     C00X
CLIENT_NAME                   <ime klijenta>
ENV                           PROD
```

Važno:

```text
GOOGLE_PWA_FOLDER_ID mora biti ID foldera 01_Sheets/02_Master.
```

Ne sme biti ID od:

```text
01_Sheets
```

Ako je `GOOGLE_PWA_FOLDER_ID` pogrešan, `Stammdaten` će nastati na pogrešnom mestu i PWA login će dati `System error`.

---

## 16. Excel `tblSEFConfig`

U `tblSEFConfig` postavi:

```text
ConfigKey             ConfigValue

SEF_BASE_URL          <SEF produkcioni URL>
SEF_ENV               PROD
SEF_DEBUG_LOG         NE
SEF_API_KEY           prazno do odlaska kod klijenta

MONITORING_ENDPOINT   <GAS Web App URL>
MONITORING_SECRET     <MONITORING_INGEST_SECRET>
MONITORING_ENV        PROD
APP_VERSION           1.0.0-C00X
CLIENT_ID             C00X
```

---

## 17. Prvi Stammdaten export

U Excelu pokreni:

```vb
SyncStammdatenToGoogle
```

Očekivanje:

```text
Export abgeschlossen: 13/13 Tabs
```

Zatim proveri Drive:

```text
AgriX_C00X_PROD/
  01_Sheets/
    02_Master/
      Stammdaten
```

Ako je `Stammdaten` nastao direktno u:

```text
01_Sheets/
```

to je greška.

Ispravka:

```text
[ ] Proveri GOOGLE_PWA_FOLDER_ID.
[ ] Postavi ga na ID od 01_Sheets/02_Master.
[ ] Premesti postojeći Stammdaten u 02_Master ili ponovi export.
[ ] Ne pravi nepotrebne duplikate.
```

---

## 18. Users tab

U `Stammdaten` proveri tab:

```text
Users
```

Header mora biti tačno:

```text
Username | PIN | Role | EntityID | DisplayName
```

Dozvoljene role:

```text
Management
Otkupac
Kooperant
Vozac
```

Minimalni test korisnici:

```text
Username     PIN    Role        EntityID     DisplayName
tkoop        9003   Kooperant   KOOP-90001   Test Kooperant
tstanica     9001   Otkupac     ST-90001     Test Stanica
tvozac       9002   Vozac       VOZ-90001    Test Vozac
admin        9999   Management  MGMT-001     Admin
```

Pravila:

```text
[ ] Username mora biti jedinstven.
[ ] PIN može biti broj ili tekst, ali mora odgovarati unosu.
[ ] EntityID mora postojati za Otkupac, Kooperant i Vozac.
[ ] Za Management preporuka je da EntityID takođe postoji, npr. MGMT-001.
```

---

## 19. Full sync

Pokreni:

```vb
RunFullPWAGoogleSyncCycle
```

Očekivanje:

```text
Geo=True
Otkup=True
Otpremnice=True
Zbirne=True
Stammdaten=True
Kartice=True
MgmtReports=True
```

Proveri da su fajlovi završili ovde:

```text
01_Sheets/02_Master/Stammdaten
01_Sheets/03_Reports/Kartice
01_Sheets/03_Reports/MgmtReports
```

Ako `Kartice` ili `MgmtReports` daju 403:

```text
[ ] GOOGLE_KARTICE_SHEET_ID = prazno
[ ] GOOGLE_MGMT_SHEET_ID = prazno
[ ] GOOGLE_REPORTS_FOLDER_ID = ID od 01_Sheets/03_Reports
[ ] Ponovo pokreni export Kartice/MgmtReports
```

---

## 20. PWA config

U `config.js` postavi:

```javascript
API_URL: 'https://script.google.com/macros/s/AKfycb.../exec',
APP_VERSION: '1.0.0-C00X'
```

Ne koristiti URL sa:

```text
/u/0/
/u/1/
/u/2/
```

U `sw.js` promeni cache name:

```javascript
const CACHE_NAME = 'AgriX-v1-C00X-prod1';
```

Deploy PWA na:

```text
https://app.agrix.rs
```

Provera:

```text
https://app.agrix.rs/src/js/config.js?v=c00x-prod1
```

Mora prikazati novi `API_URL`.

---

## 21. PWA login test

U AgriX OPS browser profilu otvori:

```text
https://app.agrix.rs?v=c00x-prod1
```

U Console proveri:

```javascript
CONFIG.API_URL
```

Mora biti novi GAS URL.

Direktni test:

```javascript
apiPostSafe('login', {
  username: 'tkoop',
  pin: '9003'
}, {
  includeToken: false
}).then(console.log)
```

Očekivanje:

```text
ok: true
data.success: true
data.role: Kooperant
data.entityID: KOOP-90001
```

Testiraj sve role:

```text
[ ] Kooperant
[ ] Otkupac
[ ] Vozac
[ ] Management
```

Ako dobiješ `System error`, proveri redom:

```text
[ ] Stammdaten je u 01_Sheets/02_Master.
[ ] Users tab postoji.
[ ] Headeri su tačni.
[ ] Role je validna.
[ ] EntityID postoji.
[ ] GAS je redeployovan kao New version.
[ ] PWA config.js pokazuje na pravi GAS URL.
```

---

## 22. LoginLog i ErrorLog test

Posle uspešnog login-a proveri da je napravljen:

```text
06_Monitoring/ErrorLog/LoginLog
```

Ako nije:

```text
[ ] Proveri da postoji getLoginLogSpreadsheet_().
[ ] Proveri da je AGRIX_MONITORING_ERRORLOG_FOLDER_ID tačan.
[ ] Redeploy GAS kao New version.
[ ] Ponovi login.
```

Test greške iz PWA console:

```javascript
apiPostSafe('logClientError', {
  errorAction: 'manualSmokeTest',
  message: 'C00X monitoring smoke test',
  details: 'manual test'
}).then(console.log)
```

Proveri:

```text
06_Monitoring/ErrorLog/ErrorLog
```

---

## 23. PWA funkcionalni smoke test

Testiraj minimalno:

```text
[ ] Kooperant login.
[ ] Kooperant vidi svoje ekrane.
[ ] Kooperant ne vidi management ekrane.

[ ] Otkupac login.
[ ] Otkupac vidi otkupne ekrane.
[ ] Otkupac može napraviti test otkup zapis.

[ ] Vozac login.
[ ] Vozac vidi vozačke ekrane.
[ ] Vozac može videti/pripremiti zbirne tokove ako su aktivni.

[ ] Management login.
[ ] Management vidi SaldoOM.
[ ] Management vidi SaldoKupci.
[ ] Management vidi Kartice.
[ ] Management vidi MgmtReports.
```

Test zapis označiti:

```text
TEST C00X - OBRISATI
```

---

## 24. Sync posle PWA test unosa

Posle PWA test unosa ponovo pokreni:

```vb
RunFullPWAGoogleSyncCycle
```

Proveri:

```text
[ ] Test OTK ulazi u VBA.
[ ] Kartice se osvežavaju.
[ ] MgmtReports se osvežava.
[ ] SyncControl lock se vraća na NO.
[ ] PWA i dalje radi posle sync-a.
```

---


## 24A. Digitalni potpis, makroi, Trusted Location i sigurnosni model

Ovaj deo je obavezan deo install package-a. Cilj nije da se Excel sigurnost nasilno isključi, nego da se AgriX workbook pokreće uredno, kontrolisano i bez ručnog kliktanja na `Enable Content` kod svakog klijenta.

### 24A.1 Pravilo za produkciju

Za produkcioni paket koristi sledeći model:

```text
[ ] OtkupApp.xlsm je digitalno potpisan.
[ ] Certifikat za proveru potpisa je u install package-u kao .cer fajl.
[ ] Setup-OtkupApp.ps1 instalira javni certifikat kod korisnika.
[ ] Setup-OtkupApp.ps1 dodaje C:\OtkupApp kao Excel Trusted Location.
[ ] Svi fajlovi iz paketa su unblocked.
[ ] Makroi se ne omogućavaju globalno za ceo Excel.
[ ] Ne koristi se opcija “Enable all macros”.
```

Dozvoljeno:

```text
Trusted Location za C:\OtkupApp
+
potpisan workbook
+
poznat publisher certifikat
```

Ne koristiti kao standard:

```text
Trust access to VBA project object model
Enable all macros
isključivanje Protected View globalno
ručnu promenu Excel Trust Center podešavanja bez dokumentovanja
```

---

### 24A.2 Tip certifikata

Za prve klijente možeš koristiti self-signed VBA certifikat, jer ti radiš inicijalnu instalaciju lično.

Preporučeni minimum za prve rollout-e:

```text
Certifikat: OtkupApp VBA Publisher
Namena: VBA project signing
Lokacija privatnog ključa: samo tvoj dev računar
U install package ide samo javni .cer, ne privatni ključ
```

Važno:

```text
Nikada ne ubacuj .pfx sa privatnim ključem u install package.
Nikada ne šalji privatni ključ klijentu.
U package ide samo .cer za Trusted Publisher / Root trust.
```

Dugoročno, kada sistem pređe na više klijenata, bolje je koristiti komercijalni code-signing certifikat, ali za prva 2–3 klijenta self-signed + lična instalacija je prihvatljiv operativni model.

---

### 24A.3 Kreiranje self-signed certifikata za VBA

Na dev računaru možeš koristiti `SelfCert.exe` koji dolazi uz Office. Tipične lokacije su:

```text
C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE
C:\Program Files (x86)\Microsoft Office\root\Office16\SELFCERT.EXE
```

Koraci:

```text
[ ] Pokreni SELFCERT.EXE.
[ ] Certificate name: OtkupApp VBA Publisher
[ ] Potvrdi kreiranje certifikata.
[ ] Certifikat ostaje u Current User / Personal store na dev računaru.
```

Ako koristiš komercijalni certifikat, preskačeš SelfCert i koristiš certifikat koji je instaliran u Windows certificate store.

---

### 24A.4 Potpisivanje OtkupApp.xlsm

Potpisivanje radiš tek kada je VBA kod spreman za release. Svaka izmena VBA koda posle potpisivanja poništava potpis.

Redosled:

```text
[ ] Otvori OtkupApp.xlsm na dev računaru.
[ ] VBA Editor: ALT + F11.
[ ] Debug > Compile VBAProject.
[ ] Ako compile ne prođe, ne potpisivati.
[ ] Tools > Digital Signature.
[ ] Choose.
[ ] Izaberi OtkupApp VBA Publisher.
[ ] Save workbook.
[ ] Zatvori Excel.
[ ] Ponovo otvori workbook i proveri da potpis nije pao.
```

Release pravilo:

```text
Poslednji korak pre pakovanja app/OtkupApp.xlsm je:
1. Compile VBA
2. Save
3. Digital Signature
4. Save
5. Close
6. Reopen smoke test
```

---

### 24A.5 Export javnog certifikata za install package

Na dev računaru otvori:

```text
certmgr.msc
```

Zatim:

```text
Current User
  Personal
    Certificates
      OtkupApp VBA Publisher
```

Export:

```text
[ ] Right click certifikat.
[ ] All Tasks > Export.
[ ] No, do not export the private key.
[ ] DER encoded binary X.509 (.CER) ili Base-64 encoded X.509 (.CER).
[ ] Naziv fajla: OtkupApp-VBA-Publisher.cer
```

Fajl ide u install package:

```text
AgriX_C00X_Install_v1.0.0/cert/OtkupApp-VBA-Publisher.cer
```

---

### 24A.6 Instalacija certifikata kod klijenta

`Setup-OtkupApp.ps1` treba da uveze javni certifikat u Current User store.

Preporučeni Current User model:

```powershell
$certPath = Join-Path $PackageRoot "cert\OtkupApp-VBA-Publisher.cer"

if (Test-Path $certPath) {
    Import-Certificate -FilePath $certPath -CertStoreLocation "Cert:\CurrentUser\TrustedPublisher" | Out-Null
    Import-Certificate -FilePath $certPath -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null
}
```

Napomena:

```text
CurrentUser ne traži nužno admin prava.
LocalMachine može tražiti admin prava.
Za prvi rollout koristi CurrentUser, jer instaliraš aplikaciju za konkretnog Windows korisnika.
```

---

### 24A.7 Excel Trusted Location

Trusted Location mora biti:

```text
C:\OtkupApp\
```

sa subfolderima.

Office 16.0 pokriva Office 2016 / 2019 / 2021 / Microsoft 365.

Primer registry upisa za Current User:

```powershell
$officeVersion = "16.0"
$trustedLocationName = "AgriX_OtkupApp"
$trustedLocationPath = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel\Security\Trusted Locations\$trustedLocationName"

New-Item -Path $trustedLocationPath -Force | Out-Null
New-ItemProperty -Path $trustedLocationPath -Name "Path" -Value "C:\OtkupApp\" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $trustedLocationPath -Name "AllowSubfolders" -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $trustedLocationPath -Name "Description" -Value "AgriX OtkupApp trusted location" -PropertyType String -Force | Out-Null
```

Ako Excel i dalje prikazuje macro warning:

```text
[ ] proveri da li workbook stvarno leži u C:\OtkupApp\
[ ] proveri registry Trusted Location
[ ] proveri da li je fajl blokiran iz interneta
[ ] proveri da li je potpis validan
[ ] proveri da li Excel koristi Office 16.0 path
```

---

### 24A.8 Unblock fajlova

Ako je install package skinut sa interneta ili kopiran sa USB-a, Windows može staviti Mark-of-the-Web. PS1 treba da uradi unblock za ceo paket i finalnu aplikaciju.

```powershell
Get-ChildItem -Path "C:\OtkupApp" -Recurse -File | ForEach-Object {
    try {
        Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue
    } catch {
        # best effort
    }
}
```

Posebno proveriti:

```text
[ ] C:\OtkupApp\OtkupApp.xlsm nije blocked
[ ] C:\OtkupApp\tools\poppler\bin\pdftotext.exe nije blocked
[ ] C:\OtkupApp\tools\poppler\bin\pdfinfo.exe nije blocked
```

---

### 24A.9 PowerShell execution policy

Za instalaciju koristi se procesni bypass, ne trajna promena sistema:

```powershell
powershell -ExecutionPolicy Bypass -File .\install\Setup-OtkupApp.ps1
```

Ovo ne menja trajno execution policy na računaru klijenta.

PS1 potpisivanje može biti uvedeno kasnije. Za prve rollout-e je dovoljno:

```text
[ ] PS1 dolazi iz tvog release paketa
[ ] pokrećeš ga lično
[ ] koristiš ExecutionPolicy Bypass samo za taj proces
[ ] zapisuješ install log
```

---

### 24A.10 Šta `Setup-OtkupApp.ps1` mora da radi

Minimalni installer mora da uradi sledeće:

```text
[ ] Detektuje package root.
[ ] Kreira C:\OtkupApp.
[ ] Kreira lokalne foldere:
    C:\OtkupApp\Bank_Izvodi\Inbox
    C:\OtkupApp\Bank_Izvodi\Processed
    C:\OtkupApp\Bank_Izvodi\Error
    C:\OtkupApp\Logs
    C:\OtkupApp\Backup
    C:\OtkupApp\Export

[ ] Kopira app/OtkupApp.xlsm u C:\OtkupApp.
[ ] Kopira tools/poppler u C:\OtkupApp\tools\poppler.
[ ] Kopira docs u C:\OtkupApp\docs.
[ ] Unblock svih fajlova.
[ ] Instalira OtkupApp-VBA-Publisher.cer ako postoji.
[ ] Dodaje C:\OtkupApp kao Excel Trusted Location.
[ ] Kreira Desktop shortcut.
[ ] Verifikuje pdftotext.exe.
[ ] Verifikuje da OtkupApp.xlsm postoji.
[ ] Piše install log u C:\OtkupApp\Logs\install-log.txt.
[ ] Na kraju ispisuje PASS/FAIL summary.
```

---

### 24A.11 Šta `SetupNewPC` mora da proveri posle PS1

`SetupNewPC` nije zamena za PS1. PS1 priprema Windows/fajlove, a `SetupNewPC` proverava aplikacionu konfiguraciju.

`SetupNewPC` mora da proveri:

```text
[ ] tblConfig postoji.
[ ] tblSEFConfig postoji.
[ ] tblLocalConfig postoji ili je kreiran.
[ ] BANKA_INBOX_PATH postoji.
[ ] BANKA_PROCESSED_PATH postoji.
[ ] BANKA_ERROR_PATH postoji.
[ ] POPPLER_PDFTOTEXT_PATH postoji i pokazuje na pdftotext.exe.
[ ] GOOGLE_CLIENT_ID postoji.
[ ] GOOGLE_CLIENT_SECRET postoji.
[ ] GOOGLE_PWA_FOLDER_ID postoji i pokazuje na 01_Sheets/02_Master.
[ ] GOOGLE_REPORTS_FOLDER_ID postoji i pokazuje na 01_Sheets/03_Reports.
[ ] MONITORING_ENDPOINT postoji.
[ ] CLIENT_ID postoji.
[ ] ENV = PROD.
[ ] APP_SETUP_COMPLETED = DA samo ako su obavezne stavke OK.
```

Ako nešto fali:

```text
APP_SETUP_COMPLETED = NE
SETUP_LOG mora jasno reći šta fali
```

---

### 24A.12 Release manifest i kontrola verzije

U svaki install package dodaj:

```text
manifest.json
release-notes.md
checksums.sha256
```

Minimalni `manifest.json`:

```json
{
  "clientId": "C00X",
  "env": "PROD",
  "appVersion": "1.0.0-C00X",
  "packageName": "AgriX_C00X_Install_v1.0.0",
  "createdAt": "YYYY-MM-DD",
  "requiresExcel": "Office 2016/2019/2021/365, version 16.0",
  "containsPoppler": true,
  "containsCertificate": true,
  "containsSignedWorkbook": true
}
```

Pre odlaska kod klijenta proveri:

```text
[ ] OtkupApp.xlsm u paketu je poslednja potpisana verzija.
[ ] APP_VERSION u tblSEFConfig odgovara manifestu.
[ ] PWA APP_VERSION odgovara release-u.
[ ] Setup-OtkupApp.ps1 je iz istog release paketa.
[ ] Poppler je prisutan.
[ ] Certifikat je prisutan.
```

---

### 24A.13 Anti-regression pravila iz C001

Za svakog novog klijenta proveri ove stavke pre PWA login testa:

```text
[ ] Stammdaten je u 01_Sheets/02_Master, ne direktno u 01_Sheets.
[ ] GOOGLE_PWA_FOLDER_ID = 01_Sheets/02_Master.
[ ] AGRIX_SHEETS_MASTER_FOLDER_ID = 01_Sheets/02_Master.
[ ] Kartice i MgmtReports su u 01_Sheets/03_Reports.
[ ] GOOGLE_REPORTS_FOLDER_ID = 01_Sheets/03_Reports.
[ ] Nema starih GOOGLE_KARTICE_SHEET_ID / GOOGLE_MGMT_SHEET_ID iz prethodnog klijenta.
[ ] Browser profil za AgriX admin je čist i prijavljen kao ops@agrix.rs.
[ ] GAS ping radi u AgriX OPS profilu.
[ ] PWA API_URL pokazuje na novi C00X GAS Web App URL.
[ ] LoginLog helper postoji.
```

---

## 25. Install package struktura

Pripremi folder:

```text
AgriX_C00X_Install_v1.0.0/
  app/
    OtkupApp.xlsm

  install/
    Setup-OtkupApp.ps1

  tools/
    poppler/
      bin/
        pdftotext.exe
        pdfinfo.exe
        ...

  cert/
    OtkupApp-VBA-Publisher.cer

  docs/
    PRE-INSTALL-C00X.md
    ON-SITE-C00X.md
    RUNBOOK-C00X.md

  manifest.json
  release-notes.md
  checksums.sha256
```

Provera:

```text
[ ] OtkupApp.xlsm ima ispravan tblConfig.
[ ] OtkupApp.xlsm ima ispravan tblSEFConfig.
[ ] modSetup postoji.
[ ] SetupNewPC postoji.
[ ] modBankaImport koristi local config paths.
[ ] Poppler postoji.
[ ] VBA compile prolazi.
[ ] Workbook je potpisan.
[ ] Javni certifikat je u cert/OtkupApp-VBA-Publisher.cer.
[ ] Setup-OtkupApp.ps1 instalira certifikat.
[ ] Setup-OtkupApp.ps1 dodaje Trusted Location.
[ ] Setup-OtkupApp.ps1 radi Unblock-File.
[ ] Setup-OtkupApp.ps1 postoji.
```

---

## 26. Lokalni install test u kancelariji

Na test Windows računaru ili čistom Windows profilu:

```text
[ ] Pokreni Setup-OtkupApp.ps1.
[ ] C:\OtkupApp postoji.
[ ] OtkupApp.xlsm je kopiran.
[ ] tools\poppler\bin\pdftotext.exe postoji.
[ ] Desktop shortcut radi.
[ ] Trusted Location radi.
[ ] Nema macro warning-a.
```

Otvori Excel i pokreni:

```vb
SetupNewPC
```

Očekivanje:

```text
APP_SETUP_COMPLETED = DA
```

Proveri:

```text
[ ] BANKA_INBOX_PATH
[ ] BANKA_PROCESSED_PATH
[ ] BANKA_ERROR_PATH
[ ] POPPLER_PDFTOTEXT_PATH
[ ] GOOGLE config
[ ] MONITORING config
[ ] SEF config
```

---

## 27. Bankarski email / PDF import pre-test

Za sada standardni tok:

```text
Banka / klijentov email
→ forwarding ili direktno slanje
→ mailbox koji VBA čita
→ VBA skida PDF
→ C:\OtkupApp\Bank_Izvodi\Inbox
→ ImportBankaInbox_TX
```

U kancelariji pripremi:

```text
[ ] znaš koji mailbox će VBA čitati
[ ] znaš IMAP/POP/Outlook tok ako se koristi
[ ] znaš kako će se podesiti forwarding
[ ] imaš test PDF izvod
[ ] Poppler radi
[ ] ImportBankaInbox_TX radi na test PDF-u
```

Test:

```text
[ ] Ubaci PDF u Bank_Izvodi\Inbox.
[ ] Pokreni ImportBankaInbox_TX.
[ ] PDF ode u Processed ili Error.
[ ] Ako je validan, redovi su u tblBankaImport.
```

---

## 28. Backup pre odlaska kod klijenta

Pre terena napravi backup:

```text
[ ] OtkupApp.xlsm baseline kopija.
[ ] tblConfig export ili screenshot.
[ ] tblSEFConfig export ili screenshot.
[ ] folder ID evidencija.
[ ] GAS Web App URL.
[ ] OAuth client ID/secret u password manager.
[ ] MONITORING_SECRET u password manager.
[ ] install package ZIP.
```

U Drive-u proveri:

```text
[ ] backup@agrix.rs ima pristup AgriX_C00X_PROD.
[ ] backup@agrix.rs ima pristup Apps Script projektu.
```

---

## 29. GO / NO-GO pre odlaska kod klijenta

Kod klijenta ideš samo ako je sve zeleno:

```text
[ ] GAS ping radi.
[ ] PWA login radi.
[ ] Sve 4 role su testirane.
[ ] Stammdaten je u 01_Sheets/02_Master.
[ ] Kartice je u 01_Sheets/03_Reports.
[ ] MgmtReports je u 01_Sheets/03_Reports.
[ ] Full sync prolazi.
[ ] PWA test unos ulazi u VBA.
[ ] LoginLog radi.
[ ] ErrorLog radi.
[ ] Poppler radi.
[ ] Bank PDF import radi.
[ ] Install package je testiran.
[ ] SetupNewPC završava APP_SETUP_COMPLETED = DA.
```

Ako bilo šta nije OK:

```text
NO-GO
```

---

# FAZA 2 — AKTIVNOSTI KOD KLIJENTA

---

## 30. Šta poneti kod klijenta

Na laptopu ili USB-u:

```text
AgriX_C00X_Install_v1.0.0
```

Unutra mora biti:

```text
[ ] app/OtkupApp.xlsm
[ ] install/Setup-OtkupApp.ps1
[ ] tools/poppler/bin/pdftotext.exe
[ ] tools/poppler/bin/pdfinfo.exe
[ ] cert/OtkupApp-VBA-Publisher.cer ako se koristi
[ ] docs/on-site checklist
```

U password manager-u moraš imati:

```text
[ ] ops@agrix.rs pristup
[ ] GAS Web App URL
[ ] OAuth client ID/secret
[ ] MONITORING_SECRET
[ ] folder ID evidencija
```

Ne držati tajne u običnom `.txt` fajlu na USB-u.

---

## 31. Početna provera kod klijenta

Pre instalacije proveri:

```text
[ ] Windows računar radi.
[ ] Excel je instaliran.
[ ] Internet radi.
[ ] PowerShell može da se pokrene.
[ ] Imaš pravo da kopiraš fajlove na C:\.
[ ] Možeš pristupiti https://app.agrix.rs.
[ ] Možeš otvoriti GAS ping URL.
[ ] Zastupnik / ovlašćeno lice je prisutno za SEF.
[ ] Dostupna je lična karta / pristup za SEF.
```

Ako nešto od ovoga nije ispunjeno, prvo rešiti preduslov.

---

## 32. Lokalna instalacija

Kopiraj install paket na računar klijenta.

Pokreni PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\install\Setup-OtkupApp.ps1
```

Proveri rezultat:

```text
[ ] C:\OtkupApp napravljen.
[ ] OtkupApp.xlsm kopiran.
[ ] Poppler kopiran.
[ ] pdftotext.exe postoji.
[ ] workbook je unblocked.
[ ] javni certifikat je instaliran u CurrentUser TrustedPublisher.
[ ] javni certifikat je instaliran u CurrentUser Root ako je self-signed.
[ ] Trusted Location dodat: C:\OtkupApp\.
[ ] Desktop shortcut napravljen.
[ ] install-log.txt postoji u C:\OtkupApp\Logs.
```

Ako PS1 failuje, ne otvarati aplikaciju dok se ne reši uzrok.

---

## 33. Prvo otvaranje Excel aplikacije

Pokreni preko desktop ikonice:

```text
OtkupApp
```

Proveri:

```text
[ ] Otvara C:\OtkupApp\OtkupApp.xlsm.
[ ] Nema macro warning-a.
[ ] Ako macro warning postoji, rešiti Trusted Location/certifikat pre nastavka.
```

---

## 34. SetupNewPC

U Excelu pokreni:

```vb
SetupNewPC
```

Očekivanje:

```text
[ ] LocalConfig postoji.
[ ] Svi lokalni folderi postoje.
[ ] BANKA_INBOX_PATH postoji.
[ ] BANKA_PROCESSED_PATH postoji.
[ ] BANKA_ERROR_PATH postoji.
[ ] POPPLER_PDFTOTEXT_PATH postoji.
[ ] Google config pronađen.
[ ] SEF config pronađen.
[ ] Monitoring config pronađen.
[ ] APP_SETUP_COMPLETED = DA.
```

Ako je:

```text
APP_SETUP_COMPLETED = NE
```

otvori setup log i reši prijavljene stavke.

---

## 35. Google / sync test kod klijenta

Pokreni:

```vb
RunFullPWAGoogleSyncCycle
```

Očekivanje:

```text
[ ] Nema OAuth greške.
[ ] Nema 403 permission greške.
[ ] Stammdaten export OK.
[ ] Kartice OK.
[ ] MgmtReports OK.
[ ] SyncControl lock se vraća na NO.
```

Ako se pojavi OAuth login, koristi pripremljeni tok i odgovarajući nalog.

---

## 36. PWA test na računaru klijenta

U browseru otvori:

```text
https://app.agrix.rs?v=c00x-onsite
```

Proveri:

```text
[ ] Login Kooperant.
[ ] Login Otkupac.
[ ] Login Vozac.
[ ] Login Management.
```

Ako je potrebno očistiti cache:

```text
[ ] Site settings > app.agrix.rs > Clear data.
[ ] Ponovo otvoriti app.agrix.rs.
```

---

## 37. PWA test na telefonu

Na telefonu:

```text
[ ] Otvori https://app.agrix.rs.
[ ] Login radi.
[ ] Add to Home Screen.
[ ] Otvori kroz ikonicu.
[ ] Login radi iz ikonice.
```

Ako postoji stara PWA instalacija:

```text
[ ] Obriši staru ikonicu.
[ ] Obriši site data za app.agrix.rs.
[ ] Otvori ponovo.
```

---

## 38. SEF setup

Sa zastupnikom / ovlašćenim licem:

```text
[ ] Uloguj se u SEF / eFaktura.
[ ] Otvori API / integracije.
[ ] Generiši ili kopiraj API key.
[ ] Potvrdi da se radi o PROD okruženju.
```

U `tblSEFConfig` upiši:

```text
SEF_API_KEY = <ključ>
```

Proveri:

```text
[ ] SEF_BASE_URL je ispravan.
[ ] SEF_ENV = PROD.
[ ] SEF_DEBUG_LOG = NE.
[ ] SEF_API_KEY nije prazan.
```

Pokreni SEF smoke test koji ne pravi poslovnu štetu.

---

## 39. Bankarski email setup kod klijenta

Sa klijentom proveri:

```text
[ ] Na koji email banka trenutno šalje izvode.
[ ] Da li banka može dodati novu adresu.
[ ] Da li treba forwarding.
[ ] Da li izvodi stižu kao PDF.
[ ] Da li PDF ima tekst, ne samo skeniranu sliku.
```

Podesi najjednostavniji tok:

```text
email sa izvodom
→ mailbox koji VBA čita
→ VBA download
→ C:\OtkupApp\Bank_Izvodi\Inbox
```

Test:

```text
[ ] Pošalji test email sa PDF izvodom.
[ ] Proveri da VBA može da ga skine.
[ ] PDF završi u lokalnom Inbox-u.
[ ] ImportBankaInbox_TX ga obrađuje.
```

---

## 40. Bank PDF import test

Na računaru klijenta:

```text
[ ] Ubaci validan PDF izvod u C:\OtkupApp\Bank_Izvodi\Inbox.
[ ] Pokreni ImportBankaInbox_TX.
[ ] Proveri tblBankaImport.
[ ] PDF ode u Processed ili Error.
```

Ako ne radi:

```text
[ ] Proveri POPPLER_PDFTOTEXT_PATH.
[ ] Proveri da pdftotext.exe postoji.
[ ] Proveri da PDF nije skenirana slika.
[ ] Proveri BANKA_*_PATH.
```

---

## 41. Monitoring test kod klijenta

Proveri:

```text
[ ] LoginLog dobija novi login.
[ ] ErrorLog radi.
[ ] Monitoring endpoint radi.
[ ] MONITORING_SECRET tačan.
```

Iz PWA konzole možeš poslati test:

```javascript
apiPostSafe('logClientError', {
  errorAction: 'onsiteSmokeTest',
  message: 'C00X onsite monitoring smoke test',
  details: 'manual onsite test'
}).then(console.log)
```

Proveri da se red pojavio u:

```text
06_Monitoring/ErrorLog/ErrorLog
```

---

## 42. Poslovni smoke test

Minimalno uradi:

```text
[ ] Otvori glavnu Excel formu.
[ ] Proveri kooperante.
[ ] Proveri kulture.
[ ] Proveri parcele.
[ ] Proveri stanice.
[ ] Proveri vozače.
[ ] Proveri kupce.
[ ] Napravi jedan test ili stvarni otkup.
[ ] Proveri PWA prikaz.
[ ] Pokreni sync.
[ ] Proveri report.
```

Ako se pravi stvarni unos:

```text
[ ] Prvi stvarni unos radiš zajedno sa korisnikom.
[ ] Proveri da je ušao u Excel.
[ ] Proveri da se vidi u PWA/reportu.
```

---

## 43. Predaja korisniku

Korisniku objasni:

```text
[ ] Aplikaciju otvara preko Desktop ikonice.
[ ] Ne premešta C:\OtkupApp.
[ ] Bankarske izvode koristi kroz dogovoreni tok.
[ ] Ne dira tblConfig.
[ ] Ne dira tblSEFConfig.
[ ] Ne dira tblLocalConfig.
[ ] Ako Excel pita za makroe, ne nastavlja nego zove podršku.
[ ] Ako PWA ne radi, prvo refresh, zatim podrška.
```

Na telefonu:

```text
[ ] Pokazati PWA ikonicu.
[ ] Pokazati login/logout.
[ ] Pokazati osnovni ekran prema roli.
```

---

## 44. Finalni zapis posle instalacije

U svojoj evidenciji upiši:

```text
CLIENT_ID:
Naziv klijenta:
Datum instalacije:
Računar:
Windows korisnik:
Excel verzija:
Install package:
APP_VERSION:
GAS Web App URL:
PWA API_URL:
Stammdaten ID:
Kartice ID:
MgmtReports ID:
SEF_API_KEY status:
PWA desktop login:
PWA mobile login:
Full sync:
Bank import:
Monitoring:
SEF smoke:
Napomene:
```

---

## 45. Finalni GO status

Klijent je pušten u rad samo ako je ovo OK:

```text
[ ] Excel aplikacija radi.
[ ] SetupNewPC = DA.
[ ] PWA radi na računaru.
[ ] PWA radi na telefonu.
[ ] Login radi za potrebne role.
[ ] Full sync prolazi.
[ ] SEF key podešen.
[ ] SEF smoke prolazi.
[ ] Bank import prolazi.
[ ] Monitoring radi.
[ ] Backup pristup postoji.
```

Ako SEF ili bank import nisu završeni, zapiši status:

```text
Sistem aktivan uz otvorene stavke:
- SEF pending
- bank email pending
- bank PDF pending
```

Ne označavati kompletan launch kao završen dok sve ključne stavke nisu potvrđene.

---

# 46. Kratka troubleshooting sekcija

## 46.1 PWA login vraća `System error`

Proveri:

```text
[ ] Stammdaten je u 01_Sheets/02_Master.
[ ] GAS getStammdatenSpreadsheet_ traži isti folder.
[ ] Users tab postoji.
[ ] Header je Username | PIN | Role | EntityID | DisplayName.
[ ] Role je validna.
[ ] EntityID postoji.
[ ] GAS je redeployovan kao New version.
```

Najčešći uzrok:

```text
Stammdaten je greškom napravljen u 01_Sheets umesto u 01_Sheets/02_Master.
```

## 46.2 Ping radi na jednom računaru, ne radi na drugom

Ako ping radi u jednom profilu/računaru, GAS je OK.

Problem je najčešće:

```text
Google multi-account browser session.
```

Rešenje:

```text
Koristi poseban browser profil AgriX OPS sa loginom samo na ops@agrix.rs.
```

## 46.3 Full sync daje 403 za Kartice/MgmtReports

Proveri:

```text
[ ] GOOGLE_KARTICE_SHEET_ID nije stari ID.
[ ] GOOGLE_MGMT_SHEET_ID nije stari ID.
[ ] Ako jesu stari, očisti ih.
[ ] GOOGLE_REPORTS_FOLDER_ID pokazuje na 01_Sheets/03_Reports.
[ ] Pokreni export ponovo.
```

## 46.4 Login radi, ali nema LoginLog fajla

Proveri da postoji:

```javascript
function getLoginLogSpreadsheet_() {
  return getOrCreateSpreadsheetInFolder_(
    'MONITORING_ERRORLOG',
    'LoginLog',
    ['Timestamp', 'Username', 'EntityID', 'Success', 'Message']
  );
}
```

Zatim redeploy GAS kao New version i ponovi login.

## 46.5 PWA pokazuje stari GAS

Proveri:

```text
[ ] https://app.agrix.rs/src/js/config.js?v=...
[ ] CONFIG.API_URL u browser console.
[ ] sw.js cache name je promenjen.
[ ] Clear site data za app.agrix.rs.
```

---

# 47. Zaključak

Za novog klijenta najbitnije je da se ne preskoče ova pravila:

```text
1. Koristiti AgriX OPS browser profil.
2. Stammdaten mora biti u 01_Sheets/02_Master.
3. Kartice i MgmtReports moraju biti u 01_Sheets/03_Reports.
4. OTK/VOZ operativni sheets idu u 01_Sheets/01_Operational.
5. GAS Script Properties moraju imati sve folder ID-jeve.
6. PWA API_URL mora biti isti GAS URL koji prolazi ping.
7. Full sync mora proći pre odlaska kod klijenta.
8. Install package mora biti testiran pre odlaska.
9. Kod klijenta se završavaju SEF, bankarski tok i finalni smoke test.
10. Launch nije završen dok PWA, Excel, SEF, banka i monitoring ne prođu test.
```
