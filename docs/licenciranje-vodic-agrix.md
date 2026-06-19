# Licenciranje — operativni vodič (AgriX)

> **Za koga:** AgriX osoblje (prodaja / podrška / instalacija). Cilj: da **bilo ko**
> može da izda/prenese/blokira licencu i da korisniku objasni šta da uradi.
>
> **Tehnički detalji** (kako radi iznutra, deploy): vidi
> `docs/licenciranje-po-uredjaju.md` i `docs/production-runbook-licenca.md`.
> Ovaj dokument je „kako se radi", ne „kako je napravljeno".

---

## 0. Ukratko (pročitaj prvo)

- **1 licencni ključ = 1 računar.** Ključ se „veže" za prvi računar koji se aktivira.
- **Vezivanje živi na serveru** (Google tabela `Licenses` u *Stammdaten*), **ne** u Excel fajlu.
  Zato kopiranje `.xlsm` na drugi računar **ne radi** — drugi računar daje drugačiji „otisak".
- **Ti (AgriX) kontrolišeš sve sa servera:** izdavanje, prenos na nov računar, blokadu.
- Korisnik na svojoj strani uradi **samo jednom**: unese ključ (potreban internet pri prvoj aktivaciji).

Tri pojma koja ćeš stalno koristiti:

| Pojam | Značenje |
|---|---|
| **Ključ (LicenseKey)** | npr. `ABCD-EFGH-JKLM-NPQR`. Daješ ga kupcu. |
| **Bind (vezivanje)** | otisak računara zalepljen za ključ na serveru pri prvoj aktivaciji. |
| **Otisak računara** | `MachineGuid + matična ploča + disk`. Jedinstven po računaru. |

---

## DEO A — AgriX strana (server: GAS editor)

> **Gde:** Google Apps Script projekat naloga `ops@agrix.rs` → otvori projekat → meni
> funkcija (gore), izaberi funkciju → **Run**. Rezultat vidiš u **View → Logs**
> (ili *Execution log*).
>
> **Prvi put ikad** (jednokratno za ceo sistem): vidi „Server setup" u
> `docs/production-runbook-licenca.md`. Ovde pretpostavljamo da je server već podešen.

### A1. Napravi novu licencu (nov kupac)

U editoru pokreni:

```js
adminCreateLicense('Naziv kupca', '')
```

- Vrati ključ (npr. `ABCD-EFGH-JKLM-NPQR`) — pročitaj ga u **View → Logs**.
- Sa rokom važenja: `adminCreateLicense('Kupac', '2027-01-01T00:00:00Z')` (prazno = trajna).
- Zapiši ključ uz kupca (CRM / tabela). Pošalji ga kupcu (vidi **DEO E**).

> Red se automatski upiše u `Licenses` sheet sa `Status = ACTIVE` i praznim `BoundParts`
> (još nije vezan).

### A2. Prenos na NOV računar (kupac promenio PC / reinstalirao Windows)

Ovo je **najčešći** zahtev. Bez ovog koraka nov računar dobija „već aktivirano na drugom".

```js
adminResetLicenseBinding('ABCD-EFGH-JKLM-NPQR')
```

- Očisti vezivanje (`BoundParts`) → ključ je opet slobodan.
- Reci kupcu da na **novom** računaru ponovo uradi aktivaciju (DEO B1).
- Stari računar time prestaje da radi (sledeća provera → „BOUND_OTHER").

### A3. Blokada / vraćanje pristupa (npr. neplaćanje)

```js
adminSuspendLicense('ABCD-EFGH-JKLM-NPQR')   // blokira
adminActivateLicense('ABCD-EFGH-JKLM-NPQR')  // vraća
```

- Blokada se na računaru primeni **najkasnije za ~3 dana** (offline grace), a odmah čim
  računar ode online.
- `adminActivateLicense` vraća `Status = ACTIVE` (ne dira vezivanje — isti računar nastavlja).

### A4. Provera stanja licence (`Licenses` sheet u Stammdaten)

Otvori `Licenses` tab i gledaj kolone:

| Kolona | Šta ti govori |
|---|---|
| `Status` | `ACTIVE` = radi; `SUSPENDED` = blokirano. |
| `BoundParts` | prazno = još nije aktivirano; popunjeno = vezano za računar. |
| `BoundAt` | datum prve aktivacije. |
| `LastSeen` | poslednji put kad je računar proverio licencu. |
| `LastDeviceInfo` | naziv računara (za prepoznavanje). |
| `ExpiresAt` | rok važenja (prazno = trajna). |

> Korisno za podršku: ako `LastSeen` star par dana → računar je offline ili se ne koristi.

---

## DEO B — Korisnikova strana (šta mu objasniti)

> Najbolje: **AgriX uradi aktivaciju daljinski/pri instalaciji.** Ako baš mora kupac sam,
> dole je tačno šta da klikne.

### B1. Prva aktivacija (jednokratno, treba internet)

1. Otvori OtkupApp.
2. **Alt + F8** (otvara listu makroa) → izaberi **`ActivateLicensePrompt`** → **Run**.
3. Unese licencni ključ koji je dobio od AgriX → OK.
4. Poruka „Licenca je uspešno aktivirana" = gotovo. Računar je sada vezan.

> Podešavanja licence (`LICENSE_ENABLED`, `LICENSE_ENDPOINT`) AgriX postavlja kroz
> **Matični podaci → Podešavanja** (tabela `tblSEFConfig` je skrivena). Izlaz u nuždi za
> AgriX: **Alt+F8 → `ShowConfigSheet`**.

### B2. Otisak računara za podršku

Ako treba da javiš serviseru koji je to računar:

- **Alt + F8 → `LicenseShowDevice`** → prikaže otisak + naziv računara. Pošalji screenshot AgriX-u.

### B3. Šta korisnik NE treba da dira

- Ne dira skrivena podešavanja, ne briše ključ, ne menja sistemski datum/sat
  (vraćanje sata unazad blokira rad — zaštita protiv varanja).

---

## DEO C — Tipične situacije (playbook)

**Nov kupac od nule**
1. `adminCreateLicense('Kupac','')` → ključ.
2. Instalacija + `LICENSE_ENDPOINT` i `LICENSE_ENABLED=YES` (Podešavanja).
3. `ActivateLicensePrompt` na kupčevom računaru → unese ključ.
4. Provera: `Licenses` → `BoundParts` popunjen, `Status=ACTIVE`.

**Kupac kupio nov računar / reinstalirao Windows**
1. `adminResetLicenseBinding('KLJUC')`.
2. Kupac na novom računaru: `ActivateLicensePrompt` → isti ključ.

**Kupac nije platio**
1. `adminSuspendLicense('KLJUC')` → blokira (do ~3 dana ili odmah online).
2. Posle naplate: `adminActivateLicense('KLJUC')`.

**Kupac kaže „ne radi mi, traži internet"**
- Prva aktivacija **mora** online. Posle toga radi offline do 3 dana pa se osveži.
- Ako je vezan računar i samo nema interneta → radi u grace prozoru; nije problem.

**„Licenca je već aktivirana na drugom računaru" (a kupac tvrdi da nije)**
- Najčešće: kupac je promenio računar/Windows bez reset-a. Uradi **A2** (reset binding).
- Ako sumnjaš na deljenje: proveri `LastDeviceInfo` / `BoundAt` u `Licenses`.

---

## DEO D — Poruke na ekranu → uzrok → akcija

| Poruka korisniku | Uzrok | Šta uraditi |
|---|---|---|
| „Licencni ključ nije unet na ovom računaru" | nema ključa | korisnik: `ActivateLicensePrompt` |
| „Licenca je već aktivirana na drugom računaru" | ključ vezan za drugi otisak (`BOUND_OTHER`) | AgriX: `adminResetLicenseBinding` pa nova aktivacija |
| „Licenca je suspendovana" | `Status=SUSPENDED` | AgriX: `adminActivateLicense` (posle naplate) |
| „Licenca je istekla" | prošao `ExpiresAt` | AgriX: napravi/produži licencu |
| „Licencni ključ nije prepoznat" | pogrešan ključ | proveri ključ (kopiraj tačno) |
| „Aktivacija licence zahteva internet" | nov/ne-vezan računar bez neta | poveži internet pa ponovo |
| „Licencni server trenutno nije dostupan" | prolazni problem servera/neta | pokušaj ponovo za par minuta |
| „Ne mogu pouzdano da očitam ovaj uređaj" | WMI/registry nedostupan | kontaktiraj AgriX (retko) |

---

## DEO E — Tekst za kupca (copy-paste u mejl/Viber)

> Poštovani,
>
> Vaš licencni ključ za OtkupApp je: **`ABCD-EFGH-JKLM-NPQR`**
>
> Aktivacija (jednokratno, potreban internet):
> 1. Otvorite OtkupApp.
> 2. Pritisnite **Alt + F8**, izaberite **ActivateLicensePrompt** i kliknite **Run**.
> 3. Unesite ključ iznad i potvrdite.
> 4. Poruka „Licenca je uspešno aktivirana" znači da je sve gotovo.
>
> Licenca važi za **jedan računar**. Ako menjate računar, javite nam da je prebacimo.
> Za pomoć: [telefon/mejl AgriX].

---

## Granica zaštite (budite iskreni prema sebi)

Sistem zaustavlja **prosto deljenje fajla** („pošalji mi OtkupApp") — to je najveći deo
rizika. Napredan korisnik koji otvori VBA editor može da zaobiđe lokalnu proveru; to je
plafon svake Excel/VBA zaštite. Prava (tvrda) zaštita bila bi da kritični podaci/obračun
žive samo na serveru. Za 99% kupaca, ovo što imamo je sasvim dovoljno.

---

## Brzi podsetnik (cheat-sheet)

| Hoću da… | Komanda / akcija |
|---|---|
| izdam ključ | `adminCreateLicense('Kupac','')` |
| prebacim na nov PC | `adminResetLicenseBinding('KLJUC')` + nova aktivacija |
| blokiram | `adminSuspendLicense('KLJUC')` |
| odblokiram | `adminActivateLicense('KLJUC')` |
| aktiviram kod kupca | Alt+F8 → `ActivateLicensePrompt` |
| pročitam otisak računara | Alt+F8 → `LicenseShowDevice` |
| otkrijem skriveni config | Alt+F8 → `ShowConfigSheet` (pa `HideConfigSheet`) |
| podesim LICENSE_*/ostalo | Matični podaci → Podešavanja |
