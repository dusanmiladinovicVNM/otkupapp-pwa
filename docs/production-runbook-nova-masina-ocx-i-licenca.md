# Production runbook: Nova mašina — `MSCOMCTL.OCX` + unos licence

Status: **operativni runbook za dva česta problema pri instalaciji na novu mašinu** —
(A) `MSCOMCTL.OCX` (Microsoft Windows Common Controls) se ne registruje / kontrola
ne učitava, i (B) kako na novoj mašini uneti/aktivirati licencu OtkupApp-a.

Aplikacija: **OtkupApp / AgriX Excel/VBA**
Vidi i: `docs/production-runbook-licenca.md` (detaljan licencni runbook),
`docs/licenciranje-po-uredjaju.md`, `install/AgriX_Onboarding_Vodic_Novi_Klijent_v2.md`,
`install/Setup-OtkupApp.ps1`.

---

## 0. Bitno pre svega: da li OtkupApp uopšte koristi `MSCOMCTL.OCX`?

> Provereno statički u repou (`src-vba/`): **commitovani izvor ne referencira
> `MSCOMCTL.OCX`.** Nijedna od 15 `.frm` formi nema `Object = "...MSCOMCTL.OCX"`
> liniju u headeru; nema `CreateObject("MSComctlLib...")`; svi „StatusBar/ProgressBar"
> pogoci su `Application.StatusBar` (Excel-ova ugrađena osobina), ne ActiveX kontrola.
> Runtime kontrole (`Controls.Add` u `clsBlokUI`/`modOtkupBlok`) su **native MSForms**
> kontrole (Label/TextBox/ComboBox/CommandButton) — one **ne traže** OCX ni licencu.

Ako ti svejedno `MSCOMCTL.OCX` pravi problem pri instalaciji, uzrok je skoro sigurno
jedan od ova tri (proveri **pre** nego što kreneš da registruješ OCX):

1. **Zaostala VBA referenca** (najčešće). U live `.xlsm` je u VBE
   `Tools → References` čekirana „**Microsoft Windows Common Controls 6.0 (SP6)**"
   iako se kontrola realno ne koristi. Na svežoj mašini gde OCX nije registrovan,
   ta referenca je „MISSING" i ruši kompajl/otvaranje (`Can't find project or
   library`), pa izgleda kao da app „zavisi" od OCX-a. → Rešenje: ako se kontrola
   ne koristi, **odčekiraj referencu** (Sekcija 5), problem nestaje zauvek.
2. **Kontrola dodata direktno u live workbook** koja nije izvezena u repo `.frm/.frx`
   (repo nije sinhron sa potpisanom produkcionom verzijom). → Onda OCX **jeste**
   prava zavisnost; idi na Sekciju 1–3 i ugradi registraciju u installer (Sekcija 5).
3. **Runtime kreirana MSCOMCTL kontrola** (`Controls.Add "MSComctlLib...."`) — traži
   i registraciju i **design-time licencu** u registru (Sekcija 4, „Error 429").

**Prvo dijagnostikuj koja je situacija** — to određuje da li uopšte treba da nosiš
i registruješ OCX, ili samo da skineš mrtvu referencu.

---

## A. Instalacija i registracija `MSCOMCTL.OCX` na novoj mašini

### A.1 Najčešći koren problema: **bitnost Office-a (32 vs 64)**

`MSCOMCTL.OCX` koji isporučuje Microsoft je **32-bitni** — ne postoji zvanična
64-bit verzija.

- **32-bit Office** (najčešće na terenu) → OCX radi.
- **64-bit Office** → 32-bit OCX **se ne može učitati** u Excel; forma sa tom
  kontrolom puca bez obzira na registraciju. Ako mašine variraju (negde 32-, negde
  64-bit Office), to je tipičan razlog „radi na jednoj, ne radi na drugoj".

Proveri bitnost: Excel → `File → Account → About Excel` (gore piše *32-bit* ili
*64-bit*). Ako je 64-bit, OCX nije rešenje — mora se zameniti kontrola native
MSForms ekvivalentom (vidi Sekciju 0, tačka 1/3).

### A.2 Nabavi tačan fajl

Uzmi `MSCOMCTL.OCX` iz **proverenog izvora i iste verzije** kao na dev/master
mašini (verzija mora da se poklapa — vidi A.6). Najpouzdanije: kopiraj fajl sa
mašine na kojoj app radi.

### A.3 Kopiraj na pravo mesto (zavisi od Windows bitnosti)

| Windows | Folder za 32-bit `MSCOMCTL.OCX` |
|---|---|
| 64-bit Windows | `C:\Windows\SysWOW64\MSCOMCTL.OCX`  (SysWOW64 drži **32-bit** binarne — da, ime zavarava) |
| 32-bit Windows | `C:\Windows\System32\MSCOMCTL.OCX` |

### A.4 Registruj kroz `regsvr32` — **kao Administrator**

Otvori **Command Prompt (Admin)** ili **PowerShell (Admin)**.

64-bit Windows (32-bit OCX → koristi 32-bit `regsvr32` iz SysWOW64):

```cmd
C:\Windows\SysWOW64\regsvr32.exe C:\Windows\SysWOW64\MSCOMCTL.OCX
```

32-bit Windows:

```cmd
regsvr32 C:\Windows\System32\MSCOMCTL.OCX
```

Uspeh = poruka `DllRegisterServer in MSCOMCTL.OCX succeeded.`

### A.5 Verifikacija

- Otvori OtkupApp → forma se učitava bez greške.
- Ili u VBE `Tools → References` — referenca više nije „MISSING".

### A.6 Česte greške

| Poruka | Uzrok | Rešenje |
|---|---|---|
| `The module "MSCOMCTL.OCX" failed to load` / `make sure the binary is stored at the specified path` | fajl ne postoji na putanji, pogrešna bitnost, ili nedostaje zavisnost | proveri putanju i bitnost (A.1/A.3); kopiraj ispravan fajl |
| `DllRegisterServer ... 0x8002801c` (*Error accessing the OLE registry*) | `regsvr32` nije pokrenut kao Administrator | pokreni elevated cmd (A.4) |
| `Object library invalid` / kontrole „nestanu" / pucanje posle Windows/Office update-a | **version mismatch** `MSCOMCTL.OCX` (poznati Office security update **MS12-027** i kasniji su menjali verziju OCX-a; workbook snimljen sa starom verzijom puca na mašini sa novom i obrnuto) | uskladi verziju OCX-a sa onom iz mastera (A.2); idealno: skini referencu/kontrolu ako se ne koristi |
| `Can't find project or library` pri otvaranju | MISSING referenca na neregistrovan OCX | registruj OCX (A.4) **ili** odčekiraj mrtvu referencu (Sekcija 5) |

---

## B. „License information for this component not found" / Run-time error **429**

Ovo je **licenca ActiveX kontrole** (ne licenca OtkupApp-a). Javlja se kad se
MSCOMCTL kontrola **kreira u runtime-u** (`Controls.Add`/`CreateObject`) na mašini
koja ima registrovan OCX, ali joj fali **design-time licencni ključ** u registru
(`HKEY_CLASSES_ROOT\Licenses\{GUID}`).

- **Ispravna registracija OCX-a kroz `regsvr32` (Sekcija A.4) upisuje i `Licenses`
  ključ** → u 90% slučajeva to reši i error 429. Ako si fajl samo kopirao bez
  `regsvr32`, ključ fali → 429.
- **Alternativa (najrobusnije):** kontrola koja je **postavljena na formu u
  design-time-u** nosi licencni blob u `.frx` fajlu forme i **ne traži** registarsku
  licencu na klijentu. Ako negde radiš `Controls.Add` MSCOMCTL kontrole, razmisli da
  je umesto toga staviš na formu (ili da pređeš na native MSForms kontrolu — vidi
  Sekciju 0).

---

## C. Unos / aktivacija licence OtkupApp-a na novoj mašini

> Ovo je OtkupApp node-locked licenca (jedan ključ = jedna mašina). Detaljan
> runbook i admin operacije: **`docs/production-runbook-licenca.md`**. Ovde je samo
> brzi „na novoj mašini" deo.

### C.1 Prva aktivacija na novoj mašini

1. U `tblSEFConfig` proveri:
   | ConfigKey | Vrednost |
   |---|---|
   | `LICENSE_ENABLED` | `YES` |
   | `LICENSE_ENDPOINT` | GAS `/exec` URL (ili **prazno** → koristi `MONITORING_ENDPOINT`) |
   | `LICENSE_KEY` | **prazno** dok ne aktiviraš (puni se kroz makro) |
2. Mašina mora imati **internet** (prva aktivacija je obavezno online; offline radi
   tek posle, do isteka grace prozora).
3. **Alt+F8 → `ActivateLicensePrompt`** → unesi licencni ključ kupca → Enter.
   (Ključ je tolerantan na velika/mala slova i crtice, ali mora postojati na serveru.)
4. **Restartuj** OtkupApp. Treba da se otvori normalno.

> Dijagnostika otiska mašine (za podršku): **Alt+F8 → `LicenseShowDevice`**
> (prikaže `MachineGuid` / `SMBIOS UUID` / `VolumeSerial`).

### C.2 Selidba licence sa stare na novu mašinu (zamena/reinstal računara)

Isti ključ na novoj mašini → server vraća `BOUND_OTHER` (vezan za staru). Postupak:

1. U GAS editoru: `adminResetLicenseBinding('KLJUC')` (oslobađa vezivanje).
2. Na novoj mašini: **Alt+F8 → `ActivateLicensePrompt`** → ponovo unesi isti ključ.
3. Restart.

> Reinstal Windowsa / zamena diska može legitimno okinuti `BOUND_OTHER` (promenjeno
> ≥2/3 hardverske komponente) — rešenje je isto: `adminResetLicenseBinding` pa
> ponovna aktivacija.

### C.3 Ako se app ne otvara da bi uneo ključ (fallback)

Otvori `.xlsm` sa **onemogućenim makroima** (drži `Shift` pri otvaranju ili klikni
*Disable Macros*), upiši `LICENSE_KEY` direktno u `tblSEFConfig`, snimi, pa otvori
normalno i pusti `ActivateLicensePrompt` da dovrši online proveru.

### C.4 Tipične poruke pri aktivaciji

| Poruka / status | Značenje | Akcija |
|---|---|---|
| „Licencni ključ nije unet" | `LICENSE_KEY` prazan | `ActivateLicensePrompt` |
| `BOUND_OTHER` | ključ vezan za drugu mašinu | `adminResetLicenseBinding('KLJUC')` → reaktiviraj |
| `UNKNOWN_KEY` | ključ ne postoji / pogrešno ukucan | proveri evidenciju ključeva |
| `SUSPENDED` / `EXPIRED` | server blokirao / istekla | `adminActivateLicense` / produži datum u `Licenses` |
| „Aktivacija zahteva internet" | prva aktivacija bez mreže | poveži internet, ponovi |

---

## 5. Trajno rešenje (da ne bude problem „pri svakoj instalaciji")

Cilj: skloni ručno petljanje sa OCX-om na svakom računaru.

### 5.1 Korak 1 — utvrdi da li je OCX **stvarno potreban** ili je mrtva referenca

Radi se u **live (potpisanom) `.xlsm`-u**, jer repo `.frm/.frx` ne moraju biti
sinhroni sa produkcionom verzijom. Tri provere:

1. **References:** `Alt+F11` (VBE) → `Tools → References`. Nađi
   „**Microsoft Windows Common Controls 6.0 (SP6)**" (ili sličnu „Common Controls").
   - Ako piše `MISSING:` ispred → referenca postoji ali OCX nije registrovan ovde.
   - Ako je čekirana bez „MISSING" → registrovan je na ovoj mašini.
2. **Compile test (presuđuje da li se koristi):**
   - `Debug → Compile VBAProject` (treba da prođe).
   - **Odčekiraj** „Microsoft Windows Common Controls" → `Debug → Compile` ponovo:
     - **Prođe čisto** → ništa je ne koristi → **mrtva referenca**, bezbedno je
       skloniti (ostavi odčekirano).
     - **Pukne** (`User-defined type not defined` / nepoznat tip kontrole) → OCX
       **jeste** u upotrebi; greška pokazuje gde. Vrati ček i idi na 5.3.
3. **Forme:** otvori svaki UserForm u design pregledu i traži ne-MSForms kontrole:
   `TreeView, ListView, ProgressBar, Toolbar, StatusBar, Slider, ImageList`
   (MSCOMCTL) ili `DTPicker, MonthView` (MSCOMCT2). Ako ih nema ni na jednoj formi
   i compile test iz tačke 2 prolazi → potvrđeno mrtva referenca.

### 5.2 Ako je **mrtva referenca** (najverovatnije, po nalazu iz repoa)

`Alt+F11` → `Tools → References` → **odčekiraj** Common Controls → `Debug → Compile
VBAProject` → snimi → **re-sign** publisher sertifikatom + **bump `APP_VERSION`** u
`modConfig` (isti release ritual kao u `docs/production-runbook-licenca.md`, deo „VBA
deploy"). Problem nestaje na svim mašinama bez ijednog OCX-a.

### 5.3 Ako OCX **jeste** zavisnost — ugradnja `regsvr32` u `Setup-OtkupApp.ps1`

> Preduslov bitnosti: 32-bit `MSCOMCTL.OCX` radi samo uz **32-bit Office** (Sekcija
> A.1). Na 64-bit Office-u registracija ne pomaže — tu se mora preći na native
> MSForms ekvivalent.

**Korak 1 — paket.** Dodaj fajl u install package: `tools/ocx/MSCOMCTL.OCX`
(pored postojećeg `tools/poppler/...`). U `manifest.json` po želji dodaj
`"containsOcx": true`.

**Korak 2 — gde tačno u PS1.** U `install/Setup-OtkupApp.ps1` blok ide **posle**
kreiranja desktop shortcut-a, a **pre** `Write-Host "=== OtkupApp setup completed ==="`
(trenutno red ~104–106). `$ScriptRoot` je već definisan na vrhu skripte (red ~24),
pa ga snippet koristi.

**Korak 3 — elevacija.** Pisanje u `SysWOW64`/`System32` i sistemska registracija
traže **admin**. Najčistije: ceo installer pokreni elevated:
```powershell
# desni klik PowerShell → Run as administrator, pa:
powershell -ExecutionPolicy Bypass -File .\install\Setup-OtkupApp.ps1
```
Snippet ispod detektuje da li je proces elevated: ako jeste, zove `regsvr32` direktno
(bez novog UAC-a); ako nije, koristi `-Verb RunAs` (jedan UAC prompt po registraciji).

**Korak 4 — snippet (zalepi na mesto iz Koraka 2):**

```powershell
# --- MSCOMCTL.OCX (32-bit) kopiranje + registracija ---
$ocxSource = Join-Path $ScriptRoot "tools\ocx\MSCOMCTL.OCX"
if (Test-Path $ocxSource) {
    # 64-bit Windows -> SysWOW64 (drzi 32-bit binarne); 32-bit Windows -> System32
    if (Test-Path (Join-Path $env:WINDIR "SysWOW64")) {
        $sysDir = Join-Path $env:WINDIR "SysWOW64"
    } else {
        $sysDir = Join-Path $env:WINDIR "System32"
    }
    $ocxTarget = Join-Path $sysDir "MSCOMCTL.OCX"
    $regsvr    = Join-Path $sysDir "regsvr32.exe"

    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent() `
        ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    try {
        Copy-Item $ocxSource $ocxTarget -Force
        if ($isAdmin) {
            & $regsvr "/s" $ocxTarget
        } else {
            Start-Process $regsvr -ArgumentList "/s `"$ocxTarget`"" -Verb RunAs -Wait
        }
        Write-Host "Registered MSCOMCTL.OCX: $ocxTarget"
    } catch {
        Write-Warning "MSCOMCTL.OCX registration failed: $($_.Exception.Message)"
    }
} else {
    Write-Host "MSCOMCTL.OCX not in package — skipping OCX registration."
}
```

**Korak 5 — provera posle PS1.** Otvori OtkupApp; forma sa kontrolom se učitava bez
`Can't find project or library` i bez Run-time 429. (Registracija kroz `regsvr32`
usput upiše i `Licenses` ključ — vidi Sekciju B.)

> Trenutni `Setup-OtkupApp.ps1` namerno radi u CurrentUser kontekstu (cert, Trusted
> Location). OCX registracija je jedini **per-machine** korak i zato traži admin —
> zato je drži opcionim (ceo `if (Test-Path $ocxSource)`): ako OCX nije u paketu,
> installer normalno preskoči taj deo.

---

## 6. Brza checklista — nova mašina

```text
OCX:
[ ] Bitnost Office-a proverena (32-bit za MSCOMCTL.OCX)
[ ] (ako treba) MSCOMCTL.OCX kopiran u SysWOW64 (64-bit Win) / System32 (32-bit Win)
[ ] regsvr32 pokrenut KAO ADMIN → "succeeded"
[ ] OtkupApp se otvara bez "Can't find project or library" / 429
[ ] ALTERNATIVA: mrtva VBA referenca odčekirana (ako se kontrola ne koristi)

Licenca OtkupApp:
[ ] tblSEFConfig: LICENSE_ENABLED=YES, LICENSE_ENDPOINT set (ili prazno→MONITORING_ENDPOINT)
[ ] Internet dostupan (prva aktivacija je online)
[ ] Alt+F8 -> ActivateLicensePrompt -> unet ključ
[ ] (selidba) adminResetLicenseBinding('KLJUC') pre reaktivacije
[ ] Restart -> app se otvara normalno
```
