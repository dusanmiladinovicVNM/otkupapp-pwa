# Uputstvo — modul „Korisnici" (prijava i prava pristupa)

> Za operatere i administratora. Objašnjava kako se uključuje i koristi sistem
> korisnika, šta koja uloga može, i **šta se radi kad neko zaboravi PIN**
> (uključujući zaključavanje admina). Verzija: **vba-v2.8.0**.

---

## 1. Šta je ovo

Modul „Korisnici" uvodi **prijavu** u aplikaciju i **prava po oblasti**:

- **Admin** — ima **sva** prava (vidi i radi sve), upravlja korisnicima.
- **Korisnik** — admin mu odobrava pristup **po oblasti**. Vidi i otvara samo
  ono za šta ima „DA"; ostalo je zaključano.

**Oblasti (12):** Otkup · Dokumenta · Agrohemija · Izveštaji · Fakturisanje ·
Banka · Marža · Sledljivost · Matični podaci · Palete · Otvori Excel ·
Sinhronizuj PWA.

> **Sve je opciono (opt-in).** Dok administrator ne uključi prijavu, aplikacija
> radi tačno kao i pre — bez prijave i bez ograničenja. Nema rizika da se
> „sami zaključate".

---

## 2. Prvo uključivanje (radi administrator, jednokratno)

Redosled je važan (ugrađena je zaštita: prijava se **ne može** uključiti dok ne
postoji bar jedan aktivan admin).

1. **Napravi tabelu korisnika i audit kolone:**
   `Alt+F8 → EnsureKorisniciSchema` (ili jednostavno `Matični → Sistem → Admin → Ensure`).
2. **Napravi prvog admina:** `Alt+F8 → KreirajPrvogAdmina`
   (pita: korisničko ime, PIN, ime i prezime).
3. **Uključi prijavu:** `Alt+F8 → EnableAuth`.
4. Zatvori i ponovo otvori aplikaciju — sada traži prijavu.

> **PIN hashing je podrazumevano UKLJUČEN** — PIN-ovi se čuvaju kao heš, ne kao
> goli tekst. Po potrebi se isključuje sa `Alt+F8 → DisablePinHash`.

**Isključivanje prijave** (vraćanje na rad bez prijave): `Alt+F8 → DisableAuth`.

---

## 3. Prijava

Pri pokretanju (kad je prijava uključena) otvara se **prozor za prijavu**:
korisničko ime + **PIN** (maskiran zvezdicama).

- Ispravna prijava → ulazi se u aplikaciju; gore u traci piše
  **`Operator: Ime Prezime`**.
- Pogrešan PIN / Otkaz → aplikacija se **zatvara** (bezbednosno). Ponovo otvori
  i probaj opet.

---

## 4. Rad sa korisnicima (admin)

**Putanja:** `Matični podaci → Sistem → Korisnici` (sekciju vidi samo admin).

**Polja u editoru:**

| Polje | Opis |
|---|---|
| Korisničko ime | jedinstveno; njime se prijavljuje |
| Ime i prezime | prikazuje se gore kao „Operator" |
| PIN | pri **izmeni**: ostavi prazno = PIN se ne menja; upiši nešto = novi PIN |
| Uloga | padajuća lista: **Admin** ili **Korisnik** |
| Aktivan | padajuća lista: **DA** / **NE** (NE = ne može da se prijavi) |
| Stanica | padajuća lista (opciono) |
| **OBLASTI** (desno) | po jedan **DA/NE** za svaku oblast — pravo pristupa |

- **Dodavanje:** popuni polja → **Dodaj**.
- **Izmena:** klikni korisnika u listi → izmeni → **Izmeni**.
- **Uloga = Admin:** svih 9 oblasti se automatski postavi na **DA** i zaključa
  (admin ionako vidi sve).
- **Deaktivacija:** dugme **Deaktiviraj/Aktiviraj** (meko brisanje — korisnik
  ostaje u evidenciji, ali se ne može prijaviti).

---

## 5. Šta korisnik vidi / ne vidi

- Korisnik otvara samo ekrane za koje ima **DA** u svojoj oblasti.
- Pokušaj otvaranja zabranjene oblasti → poruka da nema pravo.
- Admin vidi i radi sve, bez obzira na DA/NE (bypass).
- Dok je prijava **isključena**, svi vide sve (kao pre).

---

## 6. Odjava / zamena korisnika u toku rada

Gore u traci piše `Operator: Ime Prezime   [Odjava]`.
**Klik na taj natpis** → potvrda → odjava → otvara se prijava za drugog
korisnika (isti tok kao paljenje aplikacije). Nema potrebe gasiti Excel.

- Uspešna prijava drugog → traka i prava se osvežavaju na novog korisnika.
- Otkaz/neuspeh → aplikacija se zatvara (kao na startu).

---

## 7. PIN — preporuke

- PIN je lični; ne deli se. Admin može u svakom trenutku postaviti nov PIN
  korisniku (vidi tačku 4).
- **PIN hashing je podrazumevano uključen** — PIN se čuva kao nečitljiv „otisak"
  (SHA-256), ne kao goli tekst. Ako SHA nije dostupan, sistem bezbedno pada na
  plaintext (bez zaključavanja). Resetovanje PIN-a radi isto (novi PIN se hešira).
  Isključiti se može sa `Alt+F8 → DisablePinHash`.

---

## 8. Scenariji: zaboravljen PIN i oporavak

| Situacija | Rešenje (ukratko) |
|---|---|
| **Korisnik zaboravi PIN** | Admin mu postavi nov (tačka 8.1) |
| **Admin zaboravi PIN, a postoji drugi admin** | Drugi admin resetuje prvom (tačka 8.1) |
| **Jedini admin zaboravi PIN (zaključavanje)** | Oporavak preko skrivenog config-a (tačka 8.2) |
| **Korisnik „deaktiviran"** | Admin ga ponovo aktivira (Aktivan = DA) |
| **Pogrešan PIN** | Aplikacija se zatvori; samo ponovo otvori i unesi tačan PIN |

### 8.1. Korisnik (ili admin) je zaboravio PIN — postoji admin koji može da uđe

1. Admin se prijavi.
2. `Matični → Sistem → Korisnici` → klikni korisnika u listi.
3. U polje **PIN** upiši **nov PIN** → **Izmeni**.
4. Korisnik se sada prijavljuje novim PIN-om.

> Stari PIN se ne može „pročitati" (naročito uz hashing) — samo se **postavlja nov**.

### 8.2. Jedini admin je zaboravio PIN (zaključavanje aplikacije)

Pošto je prijava uključena a niko ne može da uđe, aplikacija se zatvara na
neuspeloj prijavi. Oporavak (radi administrator/vlasnik, jednokratno):

1. Zatvori Excel.
2. Otvori `AgriX_OtkupApp.xlsm` **držeći taster `Shift`** dok se otvara — time se
   **preskače** automatska prijava (otvori se radna sveska bez prijavnog prozora).
   Ako pita za makroe, klikni **Omogući sadržaj / Enable Content**.
3. `Alt+F8 → ShowConfigSheet` — otkrije skriveni list **tblSEFConfig**.
4. U tom listu nađi red sa ključem **`AUTH_ENABLED`** i promeni vrednost u **`NO`**.
   Sačuvaj (`Ctrl+S`).
5. Zatvori pa **otvori normalno** (sada nema prijave — pun pristup).
6. `Matični → Sistem → Korisnici` → izaberi admina → upiši **nov PIN** → **Izmeni**.
7. Vrati zaštitu: `Alt+F8 → HideConfigSheet`, pa `Alt+F8 → EnableAuth`
   (ponovo uključi prijavu).

> **Napomena:** makro `DisableAuth` u ovoj situaciji **neće** raditi (traži admina
> koji je prijavljen), zato se ide preko `ShowConfigSheet` i ručnog `AUTH_ENABLED = NO`.

---

## 9. Migracija iz starog fajla

„Migracija iz starog fajla" (`Matični → Sistem → Admin → Migracija`) **prenosi i
korisnike** (`tblKorisnici`: imena, PIN, uloga, aktivan, stanica, prava po oblasti,
audit kolone) — mapiranjem po imenu kolone.

- Od v2.8.0 migracija **sama** napravi `tblKorisnici` (+ audit kolone) u novom
  fajlu pre kopiranja, pa korisnici prelaze i **bez ručnog „Ensure".**
- Ako stari fajl nema korisnike (starija verzija) — nema šta da se prenese; napravi
  admina nanovo (`KreirajPrvogAdmina`).

---

## 10. Dobra praksa (preporuke)

- **Najmanje dva admina** — ako jedan zaboravi PIN, drugi ga resetuje bez procedure 8.2.
- Zapiši admin PIN na sigurno mesto (sef/menadžer lozinki).
- **PIN hashing je već uključen** (ne ostavljaj plaintext osim ako baš moraš).
- Uvodi prijavu **tek kad napraviš admina** (zaštita to i traži).

---

## 11. Tehničke napomene

- **Prijavni prozor (frmLogin)** je posebna forma — stiže uz pun `.xlsm`
  (ne kroz auto-ažuriranje koda). Za aktivaciju prijave koristi distribuirani
  `.xlsm` koji je sadrži.
- **Audit trag:** kolone `CreatedAt/CreatedBy/ModifiedAt/ModifiedBy` na glavnim
  tabelama — vidi se ko je i kada uneo/izmenio red (`Alt+F8 → SetupAuditColumns`).
- **Gde su podaci:** korisnici su u tabeli `tblKorisnici`; prekidač prijave je
  `AUTH_ENABLED` u `tblSEFConfig` (skriven posle setup-a).

### Brza referenca — `Alt+F8` komande

| Komanda | Šta radi |
|---|---|
| `EnsureKorisniciSchema` | napravi/proveri tabelu korisnika |
| `KreirajPrvogAdmina` | napravi prvog admina (ime, PIN) |
| `EnableAuth` / `DisableAuth` | uključi / isključi prijavu |
| `EnablePinHash` / `DisablePinHash` | uključi / isključi heširanje PIN-a |
| `SetupAuditColumns` | dodaj audit kolone (ko/kada) |
| `ShowConfigSheet` / `HideConfigSheet` | otkrij / sakrij `tblSEFConfig` (oporavak) |

> **Preimenovano.** Komande koje prikazuju izveštaj sada počinju sa `Setup…`:
> `SetupAuditColumns`, `SetupPaletniListSchema`, `SetupDoradeSchema`,
> `SetupStornoVezeSchema` (ranije `Ensure…`). Stara `Ensure…` imena i dalje
> postoje, ali su sada **tiha jezgra za poziv iz koda** — ne prikazuju se u
> `Alt+F8` listi i ne javljaju rezultat. Traži komandu pod `Setup…`.
