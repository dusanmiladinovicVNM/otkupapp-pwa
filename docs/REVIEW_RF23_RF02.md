# Review — RF-23 (startup + authorization) i RF-02 (modNovac finansijski guardovi)

> **Datum:** 2026-07-22 · **Reviewer:** Claude Code (statička verifikacija) ·
> **Grane:** `claude/rf-23-startup-auth-u8i5ot`, `claude/rf-02-novac-guardovi-xdfmuq` ·
> **Baza:** `origin/main` @ `9fd7087`
>
> Metod: reference-first uporedba **traženo vs. urađeno** protiv izvora istine
> (`docs/REFAKTOR_PLAYBOOK.md`, `docs/AUDIT_FM_TRIJAZA.md`, `FM-0019`) + čitanje
> stvarnog koda na granama. VBA se ne kompajlira u CI (`CLAUDE.md §5`) — sve dole je
> **statička** provera; finalni smoke-test ostaje na operateru u Excelu.

---

## Zaključak

**Obe grane su korektno uradile svoj deklarisani obim. Nije nađen nijedan funkcionalni
defekt.** Sve tvrdnje iz commit poruka su potvrđene protiv koda. Poštuju se sva tvrda
pravila iz `CLAUDE.md` (100% ASCII, `.frx` nedirano, nema novih `Private WithEvents` u
formama, reuse postojećih obrazaca, bez novih `Public` definicija → nema „Ambiguous name").

> ⚠️ **Strukturni nalaz (nije bug, ali utiče na merge):** obe grane sede na velikoj
> zajedničkoj bazi (~40 fajlova, ~50 commita) koja **nije** u `main`. Stvarni RF-23/RF-02
> rad je **samo vršni commit** svake grane. Vidi sekciju „Struktura grana".

---

## Struktura grana (kontekst pre ocene)

| | RF-23 | RF-02 |
|---|---|---|
| Vršni (zadatak) commit | `1635a8a` — 7 fajlova, +60/−4 | `3397cb3` — 2 fajla, +97/−8 |
| Merge-base sa `main` | `9fd7087` (= trenutni `main` HEAD) | `9fd7087` (isto) |
| Zajednička baza ispod | PR #147 (RF-01) + PR #141 | PR #147 (RF-01) |

`main` nije odmakao — sve ispod vršnog commita je in-flight rad koji čeka svoj merge
(self-update hardening, `clsUiSink` migracija formi, RF-01 brisanje balasta, KPI/kartica
fixevi). **Posledica:** grane se ne mogu samostalno merge-ovati u `main` bez povlačenja
cele baze. Ovo je očekivano po serijskom playbook-u (`RF-01 → RF-02 → RF-23`), ali pre
merge-a treba ili prvo spustiti bazu kroz njen PR, ili **rebase-ovati samo zadatak-commit
na čist `main`**.

Dva zadatak-delta-a diraju **disjunktne fajlove**, pa se međusobno neće sudariti:

- **RF-23:** `ThisWorkbook.doccls`, `frmOtkupAPP.frm`, `modAdmin.bas`, `modAuth.bas`,
  `modMaticniLookups.bas`, `modPodesavanja.bas`, `modPoruke.bas`
- **RF-02:** `modCenovnik.bas`, `modNovac.bas`

---

## RF-23 — Startup + authorization (AUD-033/034)

### Traženo → urađeno

| # | Zahtev (playbook) | Status | Verifikacija |
|---|---|:--:|---|
| AUD-034a | `Workbook_Open`: `If AccessWasDenied() Then Exit Sub` pre `STARTUP_SUCCESS` | ✅ | Gard je **posle `StartApp`, pre `CleanupOrphanedLocks` i pre `VBA_STARTUP_SUCCESS`**. `AccessWasDenied()` postoji (`modLicense.bas:626`), čita `gAccessDenied` koji postavlja `DenyAccessAndScheduleClose`. |
| AUD-034b | `btnBanka_Click` auth gard pre importa (obrazac `btnSyncPWA_Click`) | ✅ | Verna kopija obrasca; `OBL_BANKA` postoji (`modConfig.bas:634`); gard je **pre** `ImportBankaInbox_WithDrivePull` (koji knjiži novac / auto-map). |
| AUD-033 | `MozeAdministraciju` gard: proširiti meni + tvrde brane | ✅ | Meni gard (`MaticniMenu_OnClick`) proširen na **Korisnici/Admin/Podešavanja** — tagovi se **tačno** poklapaju sa `MaticniSekcijeGrupisano()` (uklj. `ChrW(353)` za „š"). Tvrde brane na svih 5 ulaza: `BuildAdminPanel`, `AdminPanel_OnClick`, `BuildConfigEditor`, `ShowConfigSheet`. |
| Item 4 | `PasswordChar` za „secret" polja u Podešavanjima | ✅ | `If typ = "secret" Then tb.PasswordChar = "*"`. Postoje **3 realna** secret polja (`SEF_API_KEY`, `MONITORING_SECRET`, `GOOGLE_CLIENT_SECRET`) → nije mrtav kod. Maskira samo prikaz; `SaveConfigEditor` čita `.value` normalno. |
| Item 5 | Signal pri plaintext PIN fallback-u | ✅ | `LogWarn` (postoji, `modLogError.bas:107`) pod `On Error Resume Next` (fail-soft); **logika provere nepromenjena**. |

### Anti-lockout (najveći rizik izmene — korektno rešen)

```vba
Public Function MozeAdministraciju() As Boolean
    MozeAdministraciju = (Not AuthEnabled()) Or CurrentUserIsAdmin()
End Function
```

Kad je AUTH isključen (**default** stanje), vraća `True` → tvrde brane propuštaju →
**niko nije zaključan** iz Podešavanja/Admin-a u uobičajenom radu. Brane grizu tek kad je
AUTH uključen i korisnik nije admin.

### Defense-in-depth (potvrđen)

`frmStammdaten` rutira Tag `"Podešavanja"` → `BuildConfigEditor`, `"Admin"` →
`BuildAdminPanel`; `clsAdminBtn` → `AdminPanel_OnClick`. Čak i da se meni gard zaobiđe
(npr. `Alt+F8` → `ShowConfigSheet`), tvrde brane na ulaznim tačkama drže.

### Sitno (ne-defekt)

- Poruka `AUTH_MSG_SAMO_ADMIN_KORISNICI` je posle izmene **osirotela** (nigde se više ne
  referiše na grani; ostaje samo definisana u `UpsertPoruke`). Bezopasno — jedan
  neiskorišćen red u `tblPoruke`; može se obrisati kasnije.

---

## RF-02 — modNovac finansijski guardovi (AUD-003, AUD-010)

### Traženo → urađeno

| # | Zahtev (FM-0019) | Status | Verifikacija |
|---|---|:--:|---|
| #1 / AUD-003 | `RequireColumns` za svih 17 kolona pre pozicionog `AppendRow` u `SaveNovac` | ✅ | Lista **17 kolona se poklapa sa `Array(...)` po redosledu i broju**; sve `COL_*` konstante postoje; identičan obrazac kao `modOtkup.SaveOtkup`. |
| AUD-003 | `AddCena` presence guard | ✅ | 8 kolona ↔ 8 elemenata `Array(...)`, poklapaju se. |
| #4 | Avans na fakturu **drugog kupca** → greška | ✅ | `LookupValue(... COL_FAK_KUPAC)` vs `kupacID`; `Err.Raise` na neslaganje **i na nepostojeću** fakturu. |
| #5 | Avans na otkup **drugog kooperanta** → greška | ✅ | `otkData(r, COL_OTK_KOOPERANT)` vs `kooperantID`; `Err.Raise`. |
| #6 | Avans na **stornirani** otkup/fakturu → greška | ✅ | Oba puta; obe storno konstante = `"Stornirano"` (potvrđeno — nema schema drift-a). |
| #11 | `_TX` vraća stvarno primenjeni iznos (ByRef) | ✅ | `Optional ByRef appliedAmount As Double` na oba `_TX` i base-Sub-a; Boolean zadržan; **svi postojeći pozivaoci rade nepromenjeni** (param je Optional). |

### Sigurnost pozivalaca (novi `Err.Raise`)

Dva ne-TX pozivaoca (`modFaktura.CreateFaktura`, `modOtkup.SaveOtkupMulti_TX`) primenjuju
avans na **tek kreiran cilj istog vlasnika** → nova garda tu ne okida; oba svejedno imaju
rollback `EH`. Ručni tok (`frmBankaExportPregled`) ide kroz `_TX` sa rollback-om. Nema
regresije.

### Napomene (nisu defekti, ali vredi znati)

1. **No-op i dalje vraća Boolean `True`** — po dizajnu (merljivo preko `appliedAmount`).
   ALI: potrošač `frmBankaExportPregled` **nije ažuriran** da koristi `appliedAmount`, pa
   korisnički vidljiv bug „naduvan okCount" (**FM-0020 #2/#6**) **još nije rešen** — RF-02
   je isporučio samo *mehanizam*. To je u skladu sa obimom RF-02 (samo `modNovac` +
   `modCenovnik`), ali treba biti follow-up stavka.
2. **Asimetrija nepostojećeg cilja:** nepostojeća **faktura → `Err.Raise`**, nepostojeći
   **otkup → tiho `Exit Sub`** (`appliedAmount=0`). FM-0019 #11 je sugerisao grešku za oba.
   Minorno.
3. **Presence-garda ne hvata REDOSLED kolona** (samo nedostajuću/preimenovanu) — ista
   granica kao `SaveOtkup`; prihvatljivo za izabrani obrazac.

---

## Zajedničko za obe grane (statička provera prošla)

- ✅ Svi izmenjeni `.bas/.frm/.doccls` = **`ASCII text`**, 0 ne-ASCII bajtova.
- ✅ Balans `Sub`/`Function`/`End` čist (npr. `modNovac` 5/5 + 22/22).
- ✅ **Nema novih `Private WithEvents`** u formama (self-update crash-zamka #11).
- ✅ Nema novih `Public` definicija → nema „Ambiguous name" rizika pri merge-u.
- ✅ Nova poruka `AUTH_MSG_SAMO_ADMIN_SEKCIJA` ima par u `UpsertPoruke`, bez dijakritike u
  literalu (tekst se ne gradi kroz `ChrW` jer ga i ne sadrži).

---

## Šta NIJE (i ne može biti) verifikovano ovde

VBA se ne kompajlira/pokreće u ovom okruženju. Finalni smoke-test radi operater u Excelu:

- **RF-02:** `RunNovacSmokeSuite`, `RunBusinessFlowProSuite`.
- **RF-23:**
  1. Login kao ne-admin → Matične → **Admin/Podešavanja mora biti blokirano**.
  2. Deny licence/trial → **app se zatvara**, bez lažnog `VBA_STARTUP_SUCCESS`.
  3. AUTH off → **Podešavanja se normalno otvaraju** (anti-lockout).
  4. Banka import kao ne-admin bez `OBL_BANKA` → **blokirano pre uvoza**.
  5. Podešavanja → secret polja (SEF/Monitoring/Google) prikazana **maskirano** (`*`),
     a snimljena vrednost i dalje ispravna.

---

## Preporuka

1. Kod oba zadatka je **spreman** sa aspekta same izmene — nema traženih izmena pre merge-a.
2. Pre spuštanja u `main`: **rebase vršnog commita na čist `main`** (ili prvo merge-uj
   zajedničku bazu kroz njen PR), jer trenutno svaka grana povlači ~50 commita in-flight
   baze.
3. Zavedi follow-up: ažurirati `frmBankaExportPregled` da broji `appliedAmount>0` umesto
   Boolean-a (zatvara FM-0020 #2/#6 — RF-02 je za to već postavio mehanizam).
