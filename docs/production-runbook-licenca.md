# Production runbook: Licenciranje po uređaju (node-locked)

Status: **operativni runbook za uvođenje i podršku per-uređaj licenci** — incidenti tipa „aplikacija se ne otvara, traži licencu", „licenca je već aktivirana na drugom računaru", „kupac je promenio računar", „suspenduj/aktiviraj kupca".

Aplikacija: **OtkupApp / AgriX Excel/VBA master + GAS backend**
Domen: **Workbook startup → license gate → GAS `checkLicense` → device binding u `Licenses` (Stammdaten)**
Glavni kod: `src-vba/modLicense.bas`, `src-vba/modLicenseTests.bas`, `src-vba/modMain.bas` (gate), `src-vba/ThisWorkbook.doccls` (rani prekid), `gas/Code.gs` (`checkLicense`, admin, `runLicenseSelfTest`)
Vidi i: `docs/licenciranje-po-uredjaju.md` (kratak pregled za operatera)

---

## 1. Kada korisnik kaže problem

Tipični incidenti:

* „Aplikacija se ne otvara — iskoči poruka da licenca nije uneta."
* „Piše da je licenca već aktivirana na drugom računaru."
* „Promenio sam/zamenio računar i sad ne radi."
* „Reinstalirao sam Windows i blokiralo se."
* „Kupac nije platio — kako da mu privremeno uskratim pristup?"
* „Radi li offline?"

Minimalni podaci koje operator prikuplja:

```text
Licencni kljuc kupca:
Naziv kupca:
Poruka koju vidi (tacne reci):
Da li je menjao/reinstalirao racunar:
Otisak masine (Alt+F8 -> LicenseShowDevice):
LICENSE_ENABLED u tblSEFConfig (YES/NO):
Da li masina ima internet:
```

> Prvo pravilo: pre bilo kakvog reseta, pročitaj `status` poruku — ona kaže tačan uzrok (`BOUND_OTHER`, `SUSPENDED`, `EXPIRED`, `UNKNOWN_KEY`).

---

## 2. Arhitektura ukratko

* **Vezivanje živi na serveru** (GAS `Licenses` tab u Stammdaten), ne u `.xlsm`. Kopiranje fajla ne daje pristup.
* Klijent (`modLicense`) izračuna otisak (`MachineGuid|SMBIOS UUID|VolumeSerial`), pošalje sirove komponente GAS-u; server ih **soli + hešira** (SHA-256) i čuva samo heš. Fuzzy match **2 od 3** toleriše manju promenu hardvera.
* „Prva aktivacija veže": prvi računar koji aktivira ključ se veže; svaki drugi → `BOUND_OTHER`.
* Token (HMAC) + `LICENSE_NEXT_CHECK` daju **offline grace** (default 7 dana); prva aktivacija mora online.
* **Opt-in:** gate radi samo ako je `LICENSE_ENABLED = YES`.

---

## 3. Implementacija (deploy) — redosled

### 3.1 GAS (server)
1. Ubaci izmenjeni `gas/Code.gs` u Apps Script projekat (`ops@agrix.rs`).
2. Inicijalizuj jednom iz editora:
   ```js
   licenseSecret_('LICENSE_HASH_SALT');
   licenseSecret_('LICENSE_TOKEN_SECRET');
   ensureLicensesSheet_();
   ```
3. `runLicenseSelfTest` → View → Logs → mora **8/8 PASS**.
4. **Deploy > Manage deployments > Edit > Version: New version.**
   > ⚠️ Bez „New version" `/exec` servira stari kod i `checkLicense` nije živ.

### 3.2 Tajne (zabeleži u password manager)
`LICENSE_HASH_SALT` i `LICENSE_TOKEN_SECRET` su deployment secret (kao `MONITORING_SECRET`).
> ⚠️ Gubitak/promena salta → svi sačuvani heševi nevažeći → **sve mašine `BOUND_OTHER`** dok ih ručno ne resetuješ.

### 3.3 VBA (master `.xlsm`)
1. Uvezi `modLicense.bas` i `modLicenseTests.bas` u VBE.
2. Dva mala dodatka (ako ne re-importuješ cele module):
   * `modMain.StartApp`, posle `If Not m_Initialized Then InitApp`: `If Not LicenseGateOrQuit() Then Exit Sub`
   * `ThisWorkbook.Workbook_Open`, posle `StartApp`: `If LicenseWasDenied() Then Exit Sub`
3. Debug → Compile VBAProject (bez greške).
4. Alt+F8 → `TestLicense_All` → `FAIL=0`.
5. **Re-sign** VBA projekat publisher sertifikatom; **bump** `APP_VERSION` u `modConfig`.

### 3.4 Config u masteru (`tblSEFConfig`)
| ConfigKey | Vrednost u masteru |
|---|---|
| `LICENSE_ENABLED` | `NO` (pilot prvo) |
| `LICENSE_ENDPOINT` | prazno (koristi `MONITORING_ENDPOINT`) ili `/exec` URL |
| `LICENSE_KEY` | **prazno** (po mašini) |

> ⚠️ Ostavi **prazne**: `LICENSE_KEY`, `LICENSE_TOKEN`, `LICENSE_BOUND_PARTS`, `LICENSE_NEXT_CHECK`, `LICENSE_STATUS`. Inače master nosi tvoje test-vezivanje na sve kopije.

### 3.5 Pilot pa rollout
* Provizija ključeva: `adminCreateLicense('Kupac C001', '')`.
* Na test mašini `LICENSE_ENABLED=YES` → `ActivateLicensePrompt` → restart → mora normalno.
* Dokaz: kopiraj `.xlsm` na drugu mašinu → `BOUND_OTHER`.
* Tek onda rollout sa `LICENSE_ENABLED=YES`.

---

## 4. Dijagnostika incidenta (po `status`)

| Poruka / status | Uzrok | Rešenje |
|---|---|---|
| „Licencni ključ nije unet" | `LICENSE_KEY` prazan na toj mašini | Alt+F8 → `ActivateLicensePrompt`, unesi ključ |
| `BOUND_OTHER` „aktivirana na drugom računaru" | ključ već vezan za drugu mašinu **ili** kupac reinstalirao Windows (promenjeno ≥2/3 komponente) | Potvrdi legitimnost → `adminResetLicenseBinding('KLJUC')` → kupac ponovo `ActivateLicensePrompt` |
| `SUSPENDED` | `Status=SUSPENDED` u `Licenses` | Ako treba vratiti: `adminActivateLicense('KLJUC')` |
| `EXPIRED` | `ExpiresAt` prošao | Produži datum u `Licenses` ili nov ključ |
| `UNKNOWN_KEY` | ključ ne postoji / pogrešno ukucan | Proveri evidenciju; ključ je case/crtica-tolerantan ali mora postojati |
| „Aktivacija zahteva internet" | prva aktivacija bez mreže | Poveži na internet i pokreni ponovo |

Dijagnostika otiska: **Alt+F8 → `LicenseShowDevice`** (prikaže MachineGuid/UUID/VolSerial).
Server log: GAS error sheet (`source=LICENSE`) beleži `BOUND_OTHER`/odbijanja; uspesi idu u Logger.

---

## 5. Svakodnevne operacije (GAS editor)

```js
adminCreateLicense('Naziv kupca', '');     // nov kljuc (trajna)
adminCreateLicense('Kupac', '2027-01-01T00:00:00Z'); // sa istekom
adminResetLicenseBinding('KLJUC');         // prenos na nov racunar
adminSuspendLicense('KLJUC');              // uskrati pristup
adminActivateLicense('KLJUC');             // vrati pristup
```

> Kill-switch (`SUSPENDED`) se na mašini primeni najkasnije po isteku offline grace (`LICENSE_DEFAULT_GRACE_DAYS`, sad 7 dana). Za bržu primenu smanji konstantu (`modLicense` + `Code.gs`), redeploy + re-sign.

---

## 6. Brzi rollback / „ugasi licencu odmah"

Ako uvođenje pravi problem u produkciji:

* **Po mašini:** `LICENSE_ENABLED = NO` u `tblSEFConfig` → gate se preskače (fail-open).
* **Flotno:** vrati prethodni `.xlsm` build (sa `LICENSE_ENABLED=NO`).
* Server-side ne treba ništa gasiti — bez `LICENSE_ENABLED=YES` klijent ne zove `checkLicense`.

---

## 7. Ne zaboravi (sažeto)

* GAS: redeploy kao **New version** (inače stari kod).
* VBA: **re-sign** + **bump `APP_VERSION`** posle izmene.
* Master: `LICENSE_KEY` + 4 interna keša **prazni**.
* `LICENSE_HASH_SALT`/`LICENSE_TOKEN_SECRET`: **zabeleži** (gubitak = svi `BOUND_OTHER`).
* Prvo pokretanje mora **online** (~do 15s na hladnom GAS-u; tokom provere stoji „Proveravam licencu…" u status baru); posle radi offline do grace isteka.
* Reinstalacija Windowsa / zamena diska legitimno može da okine `BOUND_OTHER` → standardno rešenje je `adminResetLicenseBinding`.

---

## 8. Granica zaštite

Server-side node-locking zaustavlja **casual deljenje** („pošalji mi fajl") — najveći deo realnog rizika. Ko otvori VBE može da izbaci poziv i radi offline; to je univerzalni plafon svakog VBA-locka. Za tvrđu zaštitu kritični podaci/obračun moraju da žive samo na serveru.
