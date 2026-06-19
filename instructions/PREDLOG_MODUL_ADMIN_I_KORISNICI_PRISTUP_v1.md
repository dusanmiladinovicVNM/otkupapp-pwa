# PREDLOG — Modul „Admin i korisnici sa ograničenjima pristupa" (v1, DRAFT)

> Status: **predlog / nije implementirano**. Pisano po Codebase Guardian doktrini
> (`reuse > new`, `extend > duplicate`, `minimal change`). Cilj: NE praviti
> paralelni auth sistem, već **proširiti postojeći** GAS + PWA sloj koji već nosi
> korisnike, role i tokene.

---

## 0) TL;DR + preporuke

- **Auth već postoji** i živi u `gas/Code.gs` (autoritet) + PWA (`src/js/services/auth.js`,
  `ui/role-nav.js`). Role su: `Management`, `Otkupac`, `Kooperant`, `Vozac`. Token
  model (48h), rate-limit (5 pokušaja / 15 min), `requireRole/requireEntity/forbiddenResponse`
  — sve postoji. **Ne diramo to, gradimo NA tome.**
- **Šta zaista fali:** (1) nema UI za upravljanje korisnicima — `Users` sheet se danas
  edituje **ručno** u Google Sheets; (2) nema „admin" odvojenog od `Management`;
  (3) nema *deaktivacije* korisnika (jedini način da se neko zaključa je brisanje reda);
  (4) restrikcije su grube (role + jedan `EntityID`) — nema per-feature ni multi-stanica
  opsega.
- **Preporučeni minimalni put (Faza 1):** Admin kao **flag** (kolona `Admin=YES` u `Users`),
  NE novi login-role → ne dira ~20 postojećih `role === 'Management'` provera. Dodati
  `Aktivan` kolonu (blok logina), `requireAdmin()` helper, 3 nove GAS akcije
  (`adminListUsers`/`adminUpsertUser`/`adminSetUserActive`) i **jedan** novi PWA ekran
  `features/management/korisnici.js` po uzoru na postojeće management ekrane.
- **VBA (Excel) je van opsega** — vidi §7. Tamo ne postoji pojam „korisnik/role";
  dodavanje bi napravilo paralelni sistem (anti-duplication).
- **Otvorene odluke** (§9) tražim da potvrdiš pre kodiranja: model admina (flag vs role),
  dubina restrikcija (samo on/off vs per-feature), i hash PIN-a (da/ne u v1).

---

## 1) Šta VEĆ postoji (inventar za reuse) — izvor istine

### 1.1 Backend — `gas/Code.gs` (autoritet za auth/role)

- **`Users` sheet** (u Stammdaten spreadsheet-u) = jedini izvor korisnika.
  Kolone se čitaju **po imenu** (schema-drift tolerantno): `Username | PIN | Role |
  EntityID | DisplayName` — `authenticateUser()` @ `Code.gs:1161`, header lookup
  preko `requireHeaderIndexFromArray(...)` @ `Code.gs:1190–1194`.
- **Login:** `authenticateUser(username, pin)` — normalizacija username-a, **plain-text**
  poređenje PIN-a (`Code.gs:1198`), rate-limit **5 pokušaja → 15 min blok** preko
  `CacheService` (`Code.gs:1176–1186`), logovanje preko `logLoginAttempt(...)`.
- **Whitelist rola:** `validateLoginUserConfig(role, entityID)` @ `Code.gs:1492` —
  dozvoljene role `['Management','Otkupac','Kooperant','Vozac']`; `EntityID` obavezan
  za sve osim `Management`.
- **Token:** `generateToken()` (UUID chain) @ `Code.gs:1235`; `saveToken(token, entityID, role)`
  @ `Code.gs:1244` (Cache + ScriptProperties, payload `{entityID, role, created, expiresAt}`);
  TTL `AUTH_TOKEN_TTL_MS = 48h` @ `Code.gs:1158`; `validateToken()` @ `Code.gs:1289`;
  `getTokenData()` @ `Code.gs:1332`; dnevni `purgeExpiredTokens()`.
- **AuthZ helperi** @ `Code.gs:283–300`:
  - `isManagement(tokenData)` — `role === 'Management'`
  - `requireRole(tokenData, allowedRoles)` — role ∈ niz
  - `requireEntity(tokenData, entityID)` — `tokenData.entityID === entityID`
  - `forbiddenResponse()` — `{success:false, error:'Nemate pristup', code:403}`
- **Dispatch:** `doPost(e)` @ `Code.gs:831`; `action === 'login'` je javno @ `Code.gs:842`;
  sve ostalo: `validateToken(data.token)` @ `:888` → `getTokenData()` @ `:892` →
  `handleAuthorizedRead(...)` @ `:898` (def @ `:303`). Pojedinačne akcije gejtovane
  `requireRole(...)` ili inline `tokenData.role !== 'Management'` (≈20 mesta).

### 1.2 Frontend — PWA (`src/`)

- **`services/auth.js`:** `showLoginScreen()` (username + 4-cifreni PIN), `doLogin()`
  (POST `login`, čuva u localStorage: `authToken, authExpiresAt, userRole, entityID,
  entityName, username`), `doLogout()`, `applyRoleVisibility()` (toggluje CSS klase
  `.role-otkupac / .role-kooperant / .role-vozac / .role-management`), `applyHeaderBranding()`.
- **`ui/role-nav.js`:** `getRoleNavConfig()` — per-role bottom-nav mapa (kooperant/otkupac/
  management/vozac); role se porede **lowercase**.
- **`config.js`:** `CONFIG.USER_ROLE / ENTITY_ID / ENTITY_NAME / USERNAME / TOKEN`, plus
  `isStoredAuthExpired()` (čisti sesiju kad token istekne).
- **Management shell:** `features/management/mgmt-shell-v2.js` — `window.mgmtShellState`
  (activeRoot, segmenti), `ensureMgmtSection(section, action, params)` (lazy fetch sekcije
  preko `apiFetch('action=...')`), `showMgmtRoot(tabKey)`.
- **Postojeći management ekrani** (`stanice.js`, `kooperanti.js`, `kupci.js`) su
  **read-only nadzor** — npr. `stanice.js` samo `getMgmtOtkupiByStanica` (`stanice.js:25`),
  nema write. Matični podaci se kreiraju u VBA Excel-u i sinkuju.
- **Reusable UI:** `ui/modal.js`, `ui/toast.js`, `styles/features-management.css`,
  `styles/auth.css`.

### 1.3 VBA — `src-vba/` (drugi kolosek, NE korisnički auth)

- `modLicense.bas` / `modTrial.bas` — licenca/trial **po uređaju** (machine fingerprint),
  ne po korisniku.
- `modGoogleAuth.bas` — OAuth ka Google-u, **app-level** (zajednički kredencijali), ne per-user.
- `modStanicaLock.bas` + `tblStanice` (`StanicaID, Naziv, Mesto, Telefon, Aktivan, Ime,
  Prezime, PIN`) — PIN je **po stanici** (ne po korisniku), za zaključavanje stanice.
- **Ne postoji** `tblKorisnici / tblRole / tblPrava` ni „admin" pojam.

---

## 2) Gap — šta tačno nedostaje za zahtev

| # | Nedostatak | Posledica danas |
|---|---|---|
| G1 | Nema UI za korisnike | `Users` sheet se edituje **ručno** u Google Sheets |
| G2 | Nema „Admin" odvojenog od `Management` | Svako Management može sve; nema ko „administrira korisnike" kao zasebno pravo |
| G3 | Nema deaktivacije | Da bi se neko zaključao, briše se red iz `Users` |
| G4 | Restrikcije su grube | Samo `role` + jedan `EntityID`; nema per-feature ni multi-stanica |
| G5 | PIN plain-text, nema admin-audita | Postoji samo login log; nema „ko je menjao korisnika" |

> Bitno: grube restrikcije **rade** (Otkupac vidi samo svoju stanicu, Kooperant svoju
> karticu, Vozac svoj transport, Management sve). Predlog ih **ne ruši** — dodaje finije
> opcije iznad njih.

---

## 3) Predlog — minimalni delta (extend, ne replace)

**Princip:** `Users` sheet ostaje jedini izvor istine; sve nove kolone su **append-only**
i čitaju se **po imenu** (kao postojeći kod). Nove GAS akcije idu kroz **postojeći**
`doPost → validateToken → getTokenData → requireRole/requireAdmin` lanac.

### 3.1 Model podataka — proširenje `Users` sheet-a (append-only)

Dodati kolone (redosled nebitan jer se čita po imenu):

| Kolona | Tip / vrednosti | Faza | Svrha |
|---|---|---|---|
| `Aktivan` | `YES` / `NO` (prazno = `YES`) | 1 | Blok logina bez brisanja reda (G3) |
| `Admin` | `YES` / `NO` | 1 | Pravo administracije korisnika (G2) |
| `Permisije` | CSV flagova npr. `FIN,DISPECER` | 2 | Per-feature restrikcije (G4) |
| `StaniceScope` | CSV `StanicaID`-jeva | 2 | Multi-stanica opseg za Otkupac/Management-lite (G4) |
| `PinHash` + `PinSalt` | string | 3 | Zamena plain PIN-a (G5), uz migraciju |
| `PromenioKorisnik` / `PromenjenoKad` | string / ISO | 1 | Trag izmene (lagani audit) |

*Opciono (samo ako zatreba „rule of three"):* zaseban `Prava` sheet (role → default
`Permisije`) kao šablon. U v1 **ne** uvodimo — držimo flagove na korisniku.

### 3.2 Backend (GAS) — sve aditivno

1. **Token payload + login** (`saveToken`, `authenticateUser`):
   - blok logina ako `Aktivan === 'NO'` → `{success:false, error:'Nalog je deaktiviran'}`;
   - u payload dodati `admin` (bool) i `permisije` (niz) i `staniceScope` (niz),
     da `getTokenData()` vraća prava bez novog čitanja sheet-a.
2. **Novi authz helperi** (uz postojeće, isti stil):
   ```js
   function isAdmin(td)            { return !!td && td.admin === true; }
   function requireAdmin(td)       { return isAdmin(td); }
   function requirePermission(td, flag) {
     return isManagement(td) || (td && Array.isArray(td.permisije) && td.permisije.indexOf(flag) >= 0);
   }
   ```
3. **Nove akcije** (u `handleAuthorizedRead` / write dispatch), sve iza `requireAdmin`:
   - `adminListUsers` → lista korisnika (bez PIN-a u odgovoru!);
   - `adminUpsertUser` → dodaj/izmeni (Username, Role, EntityID, DisplayName, Aktivan,
     Admin, Permisije, StaniceScope); validacija preko **postojećeg** `validateLoginUserConfig`;
   - `adminSetUserActive` → brza (de)aktivacija;
   - `adminResetPin` → set/replace PIN (Faza 3: hash).
   - Pristup do sheet-a preko **postojećeg** `requireHeaderIndexFromArray` obrasca (po imenu).
4. **Audit:** `logAdminAction(actor, action, target)` — mirror postojećeg login-log obrasca
   (novi `AdminLog` sheet). Faza 1 minimalno, Faza 3 puno.

> Napomena o `Management` proverama: pošto Admin ostaje **flag a ne role**, ~20 postojećih
> `tokenData.role !== 'Management'` provera se **ne dira**. Admin korisnik je i dalje
> `Management` (vidi sve) + ima `admin=true` (vidi ekran „Korisnici").

### 3.3 Frontend (PWA) — reuse management shell

- **Novi ekran:** `src/js/features/management/korisnici.js` — CRUD lista korisnika.
  Registruje se kao i ostale sekcije: `ensureMgmtSection('users', 'adminListUsers')`,
  render + `apiPost('adminUpsertUser', ...)`, `ui/modal.js` za formu, `ui/toast.js` za poruke.
- **Ulaz u meni:** dodati stavku „Korisnici" pod `partneri` segment (ili novi root) u
  `mgmt-shell-v2.js` + `role-nav.js`, vidljivo samo ako `CONFIG.IS_ADMIN`.
- **Vidljivost po pravu:** proširiti `applyRoleVisibility()` u `services/auth.js` sestrinskim
  obrascem — `.perm-fin`, `.perm-dispecer`… se gase/pale po `CONFIG.PERMISIJE`
  (isti princip kao postojeće `.role-*`).
- **`config.js`:** dodati `IS_ADMIN`, `PERMISIJE`, `STANICE_SCOPE` iz localStorage
  (login ih već može vratiti u `json`).

---

## 4) Mapa delte (koji fajlovi se diraju)

| Fajl | Tip izmene | Šta |
|---|---|---|
| `Users` sheet | data (append-only) | nove kolone iz §3.1 |
| `gas/Code.gs` | **aditivno** | `isAdmin/requireAdmin/requirePermission`, 3–4 `admin*` akcije, blok inactive u loginu, prošireni `saveToken` payload, `logAdminAction` |
| `src/js/features/management/korisnici.js` | **nov fajl** | jedini novi PWA ekran (CRUD) |
| `src/js/features/management/mgmt-shell-v2.js` | mala izmena | registracija sekcije/menija |
| `src/js/ui/role-nav.js` | mala izmena | stavka menija „Korisnici" za admina |
| `src/js/services/auth.js` | mala izmena | `.perm-*` vidljivost; čuvanje IS_ADMIN/PERMISIJE iz login odgovora |
| `src/js/config.js` | mala izmena | `IS_ADMIN, PERMISIJE, STANICE_SCOPE` |
| `index.html` | mala izmena | nav dugme + kontejner ekrana |
| `src/styles/features-management.css` | opciono | stil liste korisnika |

**Bez novih:** servisa, auth sloja, token mehanizma, tabele rola — sve se reuse-uje.

---

## 5) Tok (login → admin → restrikcija)

1. Admin se loguje (postojeći `doLogin`) → GAS vraća `role:Management, admin:true, permisije:[...]`.
2. PWA: `applyRoleVisibility()` pali `.role-management` + (novo) prikazuje „Korisnici" jer `IS_ADMIN`.
3. Admin u ekranu „Korisnici" radi `adminUpsertUser` → red u `Users` sheet-u (preko `requireAdmin`).
4. Obični korisnik se loguje → ako `Aktivan=NO` login blokiran; inače dobija `permisije` i
   `staniceScope`; GAS i PWA primenjuju restrikcije (server = autoritet, klijent = UX).

---

## 6) Bezbednost / hardening (fazirano)

- **F1:** blok logina za `Aktivan=NO`; admin akcije iza `requireAdmin`; `adminListUsers`
  **nikad** ne vraća PIN.
- **F2:** `requirePermission` na osetljivim akcijama (npr. finansije/`getMgmtSaldo*`,
  `saveDispecer`); `.perm-*` UI gašenje.
- **F3:** PIN hash (`Utilities.computeDigest` SHA-256 + per-user salt), migracija „na prvi
  login / admin reset" uz back-compat za postojeće plain PIN-ove; pun `AdminLog` audit;
  opcioni PIN expiry.

---

## 7) VBA (Excel) — van opsega (obrazloženje)

- VBA nema pojam „korisnik/role"; ima licencu/trial **po uređaju** i PIN **po stanici**.
- Per-user admin u VBA = **paralelni sistem** (krši anti-duplication §2 CLAUDE.md).
- Ako kasnije zatreba desktop per-user kontrola, izvor istine treba da ostane **isti GAS
  `Users`/token model**, a VBA da ga konzumira (ne nova `tblKorisnici`).

---

## 8) Faze isporuke (predlog redosleda)

- **Faza 1 (MVP, najmanji delta):** kolone `Aktivan`+`Admin`, blok inactive logina,
  `isAdmin/requireAdmin`, `adminListUsers/adminUpsertUser/adminSetUserActive`, PWA ekran
  „Korisnici". → **Zamena ručnog editovanja Google Sheet-a.**
- **Faza 2 (restrikcije):** `Permisije` + `requirePermission` + `.perm-*`; `StaniceScope`
  za multi-stanica.
- **Faza 3 (hardening):** PIN hash + migracija, `AdminLog` audit, PIN reset/expiry.

---

## 9) Otvorene odluke (molim potvrdu pre kodiranja)

1. **Model admina:** *flag* `Admin=YES` (preporuka — ne dira ~20 `Management` provera)
   **vs** nova login-role `'Admin'` (čistija separacija, ali širi delta).
2. **Dubina restrikcija u v1:** samo `Aktivan` (on/off) **vs** odmah i per-feature
   `Permisije` (preporuka: v1 = samo on/off, per-feature u Fazi 2).
3. **PIN hash u v1:** zadržati plain (kao sad) i hash u Fazi 3 (preporuka) **vs** odmah hash.
4. **Površina:** PWA+GAS (preporuka — tu su korisnici/role) — potvrdi da NE misliš na
   VBA Excel admin.

---

_Spreman sam da po potvrdi odluka iz §9 krenem od Faze 1 (najmanji delta), striktno
proširujući postojeći auth umesto novog sloja._
