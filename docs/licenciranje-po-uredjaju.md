# Licenciranje po uređaju (node-locked)

Licenca se prodaje **po računaru**. Jedan licencni ključ = jedna mašina.
Vezivanje (bind) živi **na serveru** (GAS), ne u `.xlsm` fajlu, pa kopiranje
fajla na drugi računar ne daje pristup.

## Kako radi

1. Kupac unese licencni ključ na svom računaru (VBA makro `ActivateLicensePrompt`).
2. VBA izračuna otisak mašine — `MachineGuid` + `SMBIOS UUID` + `volume serial`
   (`modLicense.GetDeviceParts`) — i pošalje sirove komponente GAS-u
   (`action: "checkLicense"`).
3. Server (`gas/Code.gs` → `checkLicense`):
   - **prva aktivacija** ključa veže taj otisak (`BoundParts` u sheetu `Licenses`);
   - svaki **drugi** računar sa istim ključem → `BOUND_OTHER` → blokada;
   - vraća **HMAC-potpisan token** + `graceDays`.
4. VBA kešira token i radi **offline** do isteka grace prozora, pa ponovo proverava.
   Kopiran fajl na drugom računaru lokalno ne poklapa otisak → blokira se i offline.

Otisci se na serveru **sole i heširaju** (SHA-256) — `Licenses` sheet ne drži
sirove hardverske ID-jeve. Fuzzy match (2 od 3 komponente) toleriše manju
promenu hardvera (nov disk, reinstall) bez lažnog lockout-a.

## Server setup (GAS, jednokratno)

Iz Apps Script editora pokreni ručno:

```js
adminCreateLicense('Naziv kupca', '');   // -> vrati npr. "ABCD-EFGH-JKLM-NPQR"
// trajanje: adminCreateLicense('Kupac', '2027-01-01T00:00:00Z');
```

Sheet `Licenses` se kreira automatski u `Stammdaten` spreadsheet-u. Tajne
(`LICENSE_HASH_SALT`, `LICENSE_TOKEN_SECRET`) se auto-generišu u Script Properties.

## Klijent setup (Excel, po računaru)

U `tblSEFConfig`:

| ConfigKey          | Vrednost                                  |
|--------------------|-------------------------------------------|
| `LICENSE_ENABLED`  | `YES`                                     |
| `LICENSE_ENDPOINT` | GAS Web App `/exec` URL (ili ostavi prazno → koristi `MONITORING_ENDPOINT`) |

Zatim na svakom računaru jednokratno: **Alt+F8 → `ActivateLicensePrompt`** i unesi ključ.

> `LICENSE_ENABLED` je **opt-in**: dok nije `YES`, provera je isključena
> (fail-open), pa uvođenje koda ne blokira postojeće instalacije.

## Svakodnevne operacije

| Zadatak                          | Akcija (GAS editor)                         |
|----------------------------------|---------------------------------------------|
| Nov ključ za kupca               | `adminCreateLicense('Kupac','')`            |
| **Prenos na nov računar**        | `adminResetLicenseBinding('KLJUC')` → kupac aktivira na novoj mašini |
| Privremeno uskrati pristup       | `adminSuspendLicense('KLJUC')`              |
| Vrati pristup                    | `adminActivateLicense('KLJUC')`             |
| Pročitaj otisak mašine (support) | VBA: **Alt+F8 → `LicenseShowDevice`**       |

Kill-switch (`SUSPENDED`/`EXPIRED`) se na mašini primeni najkasnije po isteku
offline grace prozora (`LICENSE_DEFAULT_GRACE_DAYS`, podrazumevano 7 dana).
Za bržu primenu smanji konstantu.

## Granica zaštite (pošteno)

Server-side node-locking zaustavlja **casual deljenje** („pošalji mi fajl") —
to je najveći deo realnog rizika. Ali ko otvori VBE može da izbaci poziv i radi
offline; to je univerzalni plafon svakog VBA-locka. Za tvrđu zaštitu kritični
podaci/obračun moraju da žive samo na serveru.
