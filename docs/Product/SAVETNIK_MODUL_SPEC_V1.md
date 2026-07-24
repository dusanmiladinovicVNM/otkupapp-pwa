# AgriX Savetnik — v1 implementaciona specifikacija (Opcija A: isti tenant)

- **Status:** radna specifikacija (draft) — razrada skice; NIJE isporučen kod
- **Datum:** 2026-07-24
- **Grana:** `claude/agrix-savetnik-modul-skica-275e7w`
- **Prethodi:** `docs/Product/SAVETNIK_MODUL_SKICA.md` (strateška skica; ovaj dokument je
  tehnička razrada **v1 = Opcija A**)
- **Opseg v1:** Savetnik i sva njegova gazdinstva su u **istom tenantu** (jedan GAS + jedan
  Drive + jedan IndexedDB). Pokriva interne agronomske službe (§206, §215) i savetnika
  jednog otkupljivača. Cross-silo (Opcija B/C) je **van v1** — vidi skicu §4.
- **Sve reference (`file:line`) verifikovane u kodu na dan pisanja.**

---

## 0. Vodeća načela (iz `CLAUDE.md`)

1. **Ne dirati postojeće kooperant tokove.** `tretman`/`trosak` model, `getTretmani`
   guard i `TRETMAN-*` sheme ostaju **netaknuti** (nula regresije, nula schema drift-a).
2. **Nov entitet `savet` je aditivni sloj**, ogledalo postojećeg `tretman` transporta.
3. **Veza plan↔izvršenje živi na `savet` strani** (`izvrsenjeTretmanID`), NE na tretmanu —
   tako se `TRETMAN_COLUMNS` ne menja (izbegnut schema drift na popunjenim sheet-ovima,
   `CLAUDE.md §4`; `ARCHITECTURE_REFERENCE.md §2.4`: „must not silently append missing
   columns to a populated production sheet").
4. **Reuse sync engine, IndexedDB API, GAS obrasce** — minimalne, izolovane izmene.

---

## 1. Komponente v1 (šta se dodaje)

| Sloj | Dodaje se | Postojeće koje se pozajmljuje |
|---|---|---|
| GAS auth | role `Savetnik`; `SAV-*` EntityID | `validateLoginUserConfig` (`Code.gs:1498`) |
| GAS veze | tab `AdvisorAssignments`; `getAssignments()`, `requireAssignedEntity()` | `requireEntity`/`isManagement` (`Code.gs:283-300`) |
| GAS write | action `syncSavet`; `processSavetRecord()` | `processTretmanRecord` (`Code.gs:2112`), `withLock`, `buildBatchSyncResponse` |
| GAS read | `getSaveti`, `getSavetnikPortfolio`, `getSavetnikTretmani` | `getTretmaniForKooperant` (`Code.gs:2476`), `getStammdaten` |
| Sheets | `SAVETI-<KooperantID>` po gazdinstvu | obrazac `TRETMAN-<KooperantID>` |
| PWA store | `saveti` (IndexedDB), `DB_VERSION 6→7` | `buildDbSchema`/`runDbMigrations` (`db.js:20-89`) |
| PWA sync | `syncSaveti` wrapper + savetnik grupisanje po gazdinstvu | `syncStore` (`sync-engine.js:227`), `sync.js` |
| PWA role | `Savetnik` u role-nav; „aktivno gazdinstvo" u sesiji | `role-nav.js`, `auth.js:174-194` |
| PWA UI (savetnik) | portfolio + forma naloga/preporuke | `management/kooperanti.js`, `agromere.js` |
| PWA UI (proizvođač) | površina „Nalozi i preporuke" | `parcele.js:736-771`, `pregled.js` |

---

## 2. Model podataka

### 2.1 GAS: `SAVETI_COLUMNS` (nov, ogledalo `TRETMAN_COLUMNS`)

Dodati uz ostale `*_COLUMNS` konstante (`Code.gs:97-158`):

```js
const SAVETI_COLUMNS = [
  'ClientRecordID',        // idempotency key (kao tretman)
  'ServerRecordID',        // generateEntityServerID('SAV', kooperantID)
  'CreatedAtClient',
  'UpdatedAtClient',
  'UpdatedAtServer',
  'SyncStatus',
  'AutorTip',              // 'savetnik' | 'management'
  'AutorId',               // SAV-xxxxx (atribucija; ostaje i posle opoziva)
  'KooperantID',           // ciljno gazdinstvo (== sheet sufiks)
  'ParcelaID',             // opciono
  'Obavezujuci',           // TRUE=radni nalog, FALSE=preporuka (§212)
  'Mera',                  // isti enum kao tretman: Zastita/Prihrana/Rezidba/Zalivanje/Berba
  'ArtikalID',
  'ArtikalNaziv',
  'DozaPreporucena',
  'JedinicaMere',
  'Rok',                   // ISO datum; osnova za kašnjenja (§213)
  'Naslov',
  'Opis',
  'Napomena',
  'StatusIzvrsenja',       // poslato|procitano|prihvaceno|u_toku|izvrseno|odbijeno|isteklo
  'IzvrsenjeTretmanID',    // clientRecordID tretmana koji ga je izvršio (veza plan→urađeno)
  'Odstupanje',            // TRUE/FALSE (§214)
  'OdstupanjeRazlog',
  'OdstupanjeKolicina',
  'ReceivedAt'
];
```

Napomene:
- **Vlasnik polja po fazi** (server-enforced, sekcija 3.4): savetnik piše „plan" polja
  (`Mera…Rok, Obavezujuci, opis`); proizvođač piše samo „izvršenje" polja
  (`StatusIzvrsenja, IzvrsenjeTretmanID, Odstupanje*`).
- `SAVETI-<KooperantID>` se kreira `getOrCreateSheet(sheetName, SAVETI_COLUMNS)`
  (`Code.gs:2946`), idempotencija po `ClientRecordID` — identično `processTretmanRecord`.

### 2.2 PWA: `savet` zapis (IndexedDB store `saveti`)

Prati konvencije `tretman` zapisa (`agromere.js:1044-1096`). camelCase na klijentu,
PascalCase kolone na serveru (mapiranje kao `agroMapServerTretman` `agromere.js:1291`).

```js
{
  clientRecordID, serverRecordID,
  createdAtClient, updatedAtClient, updatedAtServer, syncedAt,
  autorTip: 'savetnik', autorId,           // SAV-xxxxx
  kooperantID, parcelaID,
  obavezujuci: true,                        // §212
  mera, artikalID, artikalNaziv, dozaPreporucena, jedinicaMere,
  rok, naslov, opis, napomena,
  statusIzvrsenja: 'poslato',
  izvrsenjeTretmanID: '', odstupanje: false, odstupanjeRazlog: '', odstupanjeKolicina: null,
  // sync polja (identičan skup kao tretman):
  syncStatus: 'pending', syncAttempts: 0, syncAttemptAt: '', lastSyncError: '',
  lastServerStatus: '', deleted: false, entityType: 'savet', schemaVersion: 1
}
```

### 2.3 GAS: `AdvisorAssignments` (nov tab u `Stammdaten`)

Izvor **pristupa** i **naplate** (§198, §200). Kolone:

```
SavetnikID | KooperantID | Stanje | Tip | DatumOd | DatumDo | CreatedAt
```

- `SavetnikID` = `SAV-xxxxx` (== `EntityID` savetnika u `Users`).
- `Stanje` = `aktivno | pauzirano` (naplaćuje se samo `aktivno`; §198).
- `Tip` = dozvoljeni tipovi sadržaja (npr. `nalog+preporuka`).
- Provizionisanje: v1 popunjava operater (isti tok kao `Users`); standalone savetnik
  self-service je v2 (vidi otvoreno pitanje 2 u skici).

### 2.4 PWA: IndexedDB izmene (`db.js`)

Dodati `saveti` u `buildDbSchema()` (`db.js:20-60`) i podići `DB_VERSION 6→7`
(`config.js:101`). `runDbMigrations` je aditivan (`db.js:78-89`: iterira schema,
`ensureObjectStore` kreira samo nepostojeće) → **postojeći store-ovi ostaju netaknuti**.

```js
// dodati kao nov element niza u buildDbSchema():
{
  name: 'saveti',
  options: { keyPath: 'clientRecordID' },
  indexes: [
    { name: 'syncStatus', keyPath: 'syncStatus', options: { unique: false } },
    { name: 'kooperantID', keyPath: 'kooperantID', options: { unique: false } }, // particija po gazdinstvu
    { name: 'statusIzvrsenja', keyPath: 'statusIzvrsenja', options: { unique: false } },
    { name: 'rok', keyPath: 'rok', options: { unique: false } }
  ]
}
```

Index `kooperantID` je ključan: savetnik ima `saveti` za više gazdinstava u jednom store-u,
pa se sync i prikaz filtriraju po gazdinstvu (`dbGetByIndex(db,'saveti','kooperantID',id)`,
`db.js:357`).

---

## 3. GAS izmene (server)

### 3.1 Nova role `Savetnik`

`validateLoginUserConfig` (`Code.gs:1502`) — dodati u dozvoljeni skup:

```js
// bilo:  ['Management', 'Otkupac', 'Kooperant', 'Vozac']
if (!requireRole({ role: roleValue }, ['Management', 'Otkupac', 'Kooperant', 'Vozac', 'Savetnik'])) { ... }
```

`Savetnik` ima **obavezan** `EntityID` (`SAV-xxxxx`) — postojeći uslov
`roleValue !== 'Management' && !entityValue` (`Code.gs:1508`) to već pokriva.

### 3.2 Helper-i za veze (nov blok, uz `Code.gs:283-300`)

```js
function getAssignmentsForSavetnik(savetnikID) {
  // čita AdvisorAssignments iz Stammdaten; vraća [{kooperantID, stanje, tip}]
  // filtrira Stanje === 'aktivno'
}
function isAssignedEntity(tokenData, kooperantID) {
  if (isManagement(tokenData)) return true;                 // Mgmt bypass (kao svuda)
  if (tokenData.role !== 'Savetnik') return false;
  return getAssignmentsForSavetnik(tokenData.entityID)
           .some(a => String(a.kooperantID) === String(kooperantID));
}
function requireAssignedEntity(tokenData, kooperantID) { return isAssignedEntity(tokenData, kooperantID); }
```

### 3.3 Nova write grana `syncSavet` (uz `Code.gs:943` obrazac)

Jedna grana, **field-level autorizacija po roli** (savetnik piše plan; kooperant piše samo
izvršenje). Idempotentan upsert kao `processTretmanRecord`.

```js
if (data.action === 'syncSavet') {
  if (!requireRole(tokenData, ['Savetnik', 'Kooperant', 'Management'])) return forbiddenResponse();
  if (tokenData.role === 'Savetnik' && !requireAssignedEntity(tokenData, data.kooperantID)) return forbiddenResponse();
  if (tokenData.role === 'Kooperant' && !requireEntity(tokenData, data.kooperantID)) return forbiddenResponse();
  if (!Array.isArray(data.records)) return jsonResponse({ success:false, error:'records must be an array' });
  return jsonResponse(withLock(function() {
    const results = data.records.map(r => processSavetRecord(r, data.kooperantID, tokenData.role));
    return buildBatchSyncResponse(results);
  }));
}
```

### 3.4 `processSavetRecord(record, kooperantID, role)` (ogledalo `processTretmanRecord`)

Isti skelet kao `Code.gs:2112-2328`: `getOrCreateSheet('SAVETI-'+id, SAVETI_COLUMNS)`,
`ensureSheetColumns`, `headerIndexMap`, `requireHeaderIndex`, `findByColumn` po
`ClientRecordID`, `ServerRecordID = generateEntityServerID('SAV', kooperantID)`
(`Code.gs:5669`). **Jedina razlika** = field-level pravilo:

- `role === 'Savetnik' | 'Management'` → sme da upiše **plan** polja
  (`Obavezujuci, Mera, Artikal*, DozaPreporucena, Rok, Naslov, Opis, Napomena, ParcelaID`)
  i da **kreira** nov red.
- `role === 'Kooperant'` → sme da upiše **samo**
  `StatusIzvrsenja, IzvrsenjeTretmanID, Odstupanje, OdstupanjeRazlog, OdstupanjeKolicina`
  na **postojećem** redu; ako red ne postoji ili menja plan polja → odbij
  (`code: 'FIELD_NOT_ALLOWED'`). Ovo drži proizvođača kao „izvršioca", ne autora plana.

### 3.5 Nove read grane (uz `handleAuthorizedRead`, `Code.gs:303-585`)

Dodati **nove** grane — **ne dirati** postojeći `getTretmani`/`getTroskovi` guard (nula
regresije za kooperanta):

```js
if (action === 'getSaveti') {                 // i savetnik i proizvođač čitaju savete
  const kooperantID = data.kooperantID || '';
  const ok = (tokenData.role === 'Kooperant') ? requireEntity(tokenData, kooperantID)
                                               : isAssignedEntity(tokenData, kooperantID);
  if (!ok) return jsonResponse({ success:false, error:'Nemate pristup', code:403 });
  return jsonResponse(getSavetiForKooperant(kooperantID));   // čita SAVETI-<id>
}

if (action === 'getSavetnikTretmani') {       // savetnik čita tretmane dodeljenog gazd.
  const kooperantID = data.kooperantID || '';
  if (!isAssignedEntity(tokenData, kooperantID)) return jsonResponse({ success:false, error:'Nemate pristup', code:403 });
  return jsonResponse(getTretmaniForKooperant(kooperantID));  // REUSE postojeće čitanje
}

if (action === 'getSavetnikPortfolio') {      // lista dodeljenih gazdinstava + sažetak
  if (tokenData.role !== 'Savetnik' && !isManagement(tokenData)) return jsonResponse({ success:false, error:'Nemate pristup', code:403 });
  return jsonResponse(getSavetnikPortfolio(tokenData.entityID)); // assignments ⋈ stammdaten.kooperanti
}
```

> **Rešava read-asimetriju** (skica §4.2.1): `getSavetnikTretmani` **reuse-uje** postojeći
> `getTretmaniForKooperant` (`Code.gs:2476`), ali sa `isAssignedEntity` guardom umesto
> striktne jednakosti — bez diranja kooperantovog puta.

---

## 4. PWA izmene (klijent)

### 4.1 Role i sesija

- `Savetnik` u role-routingu: `getRoleNavConfig()` (`role-nav.js:5`) nova grana
  (`navId:'savetnikBottomNav'`, tabovi `portfolio/gazdinstvo/nalozi/vise`);
  `applyRoleVisibility` (`auth.js:174`) `.role-savetnik`.
- **Aktivno gazdinstvo** u sesiji: `CONFIG.ACTIVE_KOOPERANT_ID` (localStorage
  `activeKooperantID` — dozvoljeno kao device-pref, `ARCHITECTURE_REFERENCE.md §2.5`).
  Bira se iz portfolija; svi read/write pozivi savetnika koriste taj id kao `kooperantID`.

### 4.2 Sync wrapper (`features/savetnik/sync.js`, ogledalo `kooperant/sync.js`)

```js
window.syncSavetiForGazdinstvo = function (kooperantID) {
  return syncStore({
    storeName: 'saveti',
    action: 'syncSavet',
    inFlightKey: 'savetiInFlight',
    entityIdField: 'kooperantID',
    entityId: kooperantID,                 // NOVO: override (vidi 4.3)
    pendingFilter: r => r.kooperantID === kooperantID,   // NOVO: samo ovo gazdinstvo
    successLabel: 'Saveti sinhronizovani'
  });
};
window.syncSavetnikNow = async function () {
  const pending = await dbGetByIndex(db, 'saveti', 'syncStatus', 'pending');
  const gazdinstva = [...new Set(pending.map(r => r.kooperantID))];   // grupiši po gazdinstvu
  for (const id of gazdinstva) await window.syncSavetiForGazdinstvo(id);
};
```

Proizvođačeva strana koristi isti `saveti` store, ali su svi njegovi `pending` saveti sa
`kooperantID === CONFIG.ENTITY_ID` → proizvođač zove standardni `syncStore` bez override-a
(ista grana `syncSavet`, server prepoznaje rolu `Kooperant` i primenjuje field-level pravilo).

### 4.3 Minimalna ekstenzija `syncStore` (`sync-engine.js`)

Danas `entityID = CONFIG.ENTITY_ID || CONFIG.OTKUPAC_ID` (`sync-engine.js:242`) i uzima
**sve** pending (`:277`). Dve aditivne, opciono-uslovljene izmene:

1. `const entityID = opts.entityId || CONFIG.ENTITY_ID || CONFIG.OTKUPAC_ID;`
   (bez `opts.entityId` ponašanje je nepromenjeno — kooperant/otkupac rade kao dosad).
2. Ako je `opts.pendingFilter` prisutan, filtriraj `pending` tim predikatom pre
   `markPendingAsSyncing` (`:282`). Bez njega — nepromenjeno.

Nula uticaja na postojeće pozivaoce (`syncTretmani/syncTroskovi` ne prosleđuju nove opcije).

### 4.4 UI — savetnik (nova role-nav grupa; reuse Management shell)

- **Portfolio** (`features/savetnik/portfolio.js`): lista iz `getSavetnikPortfolio`;
  po gazdinstvu badge-evi (aktivni nalozi, kašnjenja `rok<danas`, odstupanja) — reuse
  brojača iz `pregled.js:157-169`. Klik → set `activeKooperantID` → detalj.
- **Detalj gazdinstva**: reuse read-only prikaza po parceli (`parcele.js:736-771`) preko
  `getSavetnikTretmani`/`getSaveti`; + „Novi nalog / preporuka".
- **Forma naloga** (`features/savetnik/nalog-form.js`): varijanta `agromere` wizarda
  **bez tajmera/GPS-a** (plan, ne izvršenje): parcela → mera → artikal → `dozaPreporucena`
  (reuse `agroCalcPreporuka` `agromere.js:609`) → `obavezujuci` toggle → `rok` → snimi
  (`dbPut(db,'saveti',record)`) → `syncSavetiForGazdinstvo(activeKooperantID)`.
- **Pregled izvršenja**: planirano (`savet.dozaPreporucena`) vs urađeno
  (`tretman.dozaPrimenjena` preko `izvrsenjeTretmanID`); odstupanja.

### 4.5 UI — proizvođač (role `Kooperant`, minimalan dodatak)

- **Površina „Nalozi i preporuke"** (`features/kooperant/nalozi.js`): pull `getSaveti`
  (kooperantID = svoj), lokalni `saveti` store; lista sa statusom; akcije:
  - **Prihvati / Odbij** (obavezujući) → `statusIzvrsenja` + `syncSavet`.
  - **Evidentiraj izvršenje** → deep-link u `agromere` preselektovano iz saveta (u duhu
    `goToNewRadFromParcela` `parcele.js:964`); po snimanju tretmana klijent postavi
    `savet.izvrsenjeTretmanID = tretman.clientRecordID`, `statusIzvrsenja='izvrseno'` →
    queue `syncSavet`.
  - **Prijavi odstupanje** → `odstupanje/odstupanjeRazlog` → `syncSavet`.
- **U detalju parcele**: saveti za tu parcelu (uz postojeći prikaz radova/troškova).
- Ostatak Pro aplikacije **nepromenjen**.

---

## 5. Ključni tokovi (sekvence)

**A. Savetnik šalje nalog**
1. Savetnik → Portfolio → izabere gazdinstvo (`activeKooperantID`).
2. Forma → `savet` (`statusIzvrsenja='poslato'`, `syncStatus='pending'`) → `dbPut('saveti')`.
3. `syncSavetiForGazdinstvo(id)` → `syncSavet` → `SAVETI-<id>` (idempotent) →
   `serverRecordID`, `syncStatus='synced'`.
4. (opc.) RTDB signal „nov nalog" ka gazdinstvu (presedan `intercom-monitor.js`).

**B. Proizvođač primi i izvrši**
1. Kooperant app pull `getSaveti` (svoj id) → lokalni `saveti`.
2. „Prihvati" → `statusIzvrsenja='prihvaceno'` → `syncSavet` (rola Kooperant, field-level).
3. „Evidentiraj izvršenje" → `agromere` preselektovano → snimi `tretman` (`syncTretman`,
   nepromenjeno) → klijent veže `savet.izvrsenjeTretmanID`, `='izvrseno'` → `syncSavet`.
4. Savetnik pull `getSaveti`/portfolio → vidi `izvrseno` + `dozaPrimenjena` vs `dozaPreporucena`.

**C. Odstupanje (§214)**
1. Kooperant → „Prijavi odstupanje" + razlog → `odstupanje=true` → `syncSavet`.
2. Savetnik dobija upozorenje (alert lista + opc. RTDB) na portfoliju.

**D. Opoziv veze (§228)**
1. Operater postavi `AdvisorAssignments.Stanje='pauzirano'` (ili obriše red).
2. `isAssignedEntity` pada → savetnik gubi read/write za to gazdinstvo **odmah**.
3. `SAVETI-<id>` i `TRETMAN-<id>` ostaju kod proizvođača (njegovi su; §183, §201);
   `AutorId` čuva atribuciju istorije.

---

## 6. Autorizaciona matrica (v1)

| Operacija | Savetnik | Kooperant (vlasnik) | Management | Ostali |
|---|---|---|---|---|
| `syncSavet` — plan polja | ✔ ako `assigned` | ✘ | ✔ | ✘ |
| `syncSavet` — izvršenje polja | ✘ | ✔ na svoj | ✔ | ✘ |
| `getSaveti` | ✔ ako `assigned` | ✔ svoj | ✔ | ✘ |
| `getSavetnikTretmani` | ✔ ako `assigned` | (koristi `getTretmani`) | ✔ | ✘ |
| `getSavetnikPortfolio` | ✔ (svoj SAV-id) | ✘ | ✔ | ✘ |
| `getTretmani` (postojeće) | ✘ (nepromenjeno) | ✔ svoj | ✘ (nepromenjeno) | ✘ |

`✘` = `forbiddenResponse()` / `code:403`. Guard je uvek iz tokena (`tokenData.role`,
`tokenData.entityID`), nikad iz klijentskog inputa.

---

## 7. Migracije i kompatibilnost

- **IndexedDB:** `DB_VERSION 6→7`; `runDbMigrations` aditivno kreira `saveti`
  (`db.js:78-89`). Stari podaci netaknuti; nema data-loss puta.
- **Sheets:** `SAVETI-<id>` se kreira on-demand pri prvom `syncSavet`; nula uticaja na
  `TRETMAN-*`/`TROSKOVI-*`. `Stammdaten` dobija nov tab `AdvisorAssignments` (nov tab, ne
  menja postojeće).
- **Stari kooperant klijent (pre v1)** koji nije dobio update: ne poznaje `saveti` store i
  površinu „Nalozi" — ali ništa mu se ne lomi (ne dobija saveti dok se ne update-uje).
  Savetnik zavisi od toga da su gazdinstva na v1 klijentu → uslov aktivacije veze.
- **`checkVersion` gate** (`Code.gs:853`) i PWA `sw.js` cache-bust: podići `APP_VERSION`
  (`config.js:104`) i verzionisati SW kao pri svakom PWA release-u.

---

## 8. Acceptance kriterijumi

1. Savetnik login (role `Savetnik`, `SAV-*`) prolazi; ne-savetnik ne dobija savetnik nav.
2. Savetnik vidi **samo** `aktivno`-dodeljena gazdinstva; ne-dodeljeno gazdinstvo →
   `getSaveti/getSavetnikTretmani` vraća `403`.
3. Kreiran nalog stiže u `SAVETI-<id>`; idempotentan (dupli sync ne pravi duplikat —
   `findByColumn` po `ClientRecordID`).
4. Proizvođač vidi nalog, „Prihvati" i „Evidentiraj izvršenje" rade offline pa se
   sinhronizuju; `izvrsenjeTretmanID` veže tretman; status → `izvrseno`.
5. Proizvođač NE može da izmeni plan polja (server `FIELD_NOT_ALLOWED`); savetnik NE može
   da izmeni tuđe (ne-dodeljeno) gazdinstvo.
6. Odstupanje se propagira; savetnik ga vidi kao alert.
7. Opoziv veze (`pauzirano`) → savetnik odmah gubi pristup; zapisi ostaju kod proizvođača.
8. **Nula regresije:** postojeći kooperant `getTretmani`/`syncTretman`/`syncTrosak` tokovi
   rade identično (isti guard, isti store, isti sheet).

## 9. Test checklist (operater, ručno)

1. U `Users` dodaj savetnika (`Savetnik`, `SAV-90001`); u `AdvisorAssignments` veži 2
   gazdinstva `aktivno`. Login kao savetnik → očekuj portfolio sa 2 gazdinstva.
2. Kreiraj „radni nalog" (Zastita, artikal, doza, rok) za gazd. #1 → proveri red u
   `SAVETI-<KOOP #1>` i `syncStatus=synced`.
3. Login kao kooperant #1 → „Nalozi i preporuke" prikazuje nalog → „Prihvati" → „Evidentiraj
   izvršenje" (snimi tretman) → proveri `TRETMAN-<#1>` red i savet status `izvrseno`.
4. Isključi net na koraku 3 (offline), pa uključi → proveri da se saveti/tretman
   sinhronizuju (pending→synced).
5. Kao savetnik pokušaj otvoriti gazd. koje NIJE dodeljeno (ručno pozovi `getSaveti`) →
   očekuj `403`.
6. Prijavi odstupanje kao kooperant → kao savetnik proveri alert.
7. Postavi vezu na `pauzirano` → savetnik više ne vidi to gazdinstvo; kooperantu podaci ostaju.
8. Regresija: kao „obični" kooperant (bez savetnika) uradi tretman i trošak kao ranije →
   sve radi identično.

---

## 10. Redosled rada (PR-ovi)

1. **PR1 — GAS temelj (bez UI):** role `Savetnik`, `AdvisorAssignments` + helper-i,
   `SAVETI_COLUMNS`, `processSavetRecord`, `syncSavet`, read grane. Testirati preko
   direktnih `apiPost` poziva.
2. **PR2 — PWA data sloj:** `DB_VERSION 6→7` + `saveti` store, `syncSaveti` wrapper,
   `syncStore` ekstenzija (`entityId`/`pendingFilter`).
3. **PR3 — Savetnik UI:** role-nav + portfolio + forma naloga (reuse agromere/mgmt).
4. **PR4 — Proizvođač UI:** površina „Nalozi i preporuke" + deep-link izvršenja + odstupanje.
5. **PR5 — Polir:** RTDB signali, alerti kašnjenja, pregled planirano-vs-urađeno, proba/cap
   (§209–210), read-only posle probe.

Svaki PR: `git merge-tree` provera, statička provera (nema duplih `Public`/simbola u GAS-u),
smoke-test u realnom tenantu (`RELEASE_GATES.md`).

---

## 11. Otvorena pitanja specifična za implementaciju

1. **Provizionisanje `SAV-*` i veza bez desktopa** (standalone savetnik) — v1 pretpostavlja
   operatera; self-service registracija je v2 (skica §12/2).
2. **Naplatni događaj** — životni ciklus `AdvisorAssignments.Stanje` (proba→aktivno→pauza)
   kao osnov obračuna po gazdinstvu (§198); gde se broji „aktivno" na dan fakture.
3. **Skalabilnost `withLock`** (`Code.gs:5307`) — jedan lock po tenantu; masovni savetnik
   sync serijalizuje. Za veliki portfolio: batch po gazdinstvu (već predviđeno u 4.2) +
   opc. throttle.
4. **RTDB pravila** za savetnik↔gazdinstvo signal (autorizacija po entitetu).
5. **Finalni enum statusa** i da li `procitano` treba (vs. samo `poslato→prihvaceno`).
6. **Proba/cap ≤10** (§210) — gde se enforce-uje (broj `aktivno` veza po savetniku pri
   dodeli).

---

_Sve `file:line` reference verifikovane pri pisanju. Pre kodiranja: `DebugKoloneTabele`
analog na serveru (proveri stvarne nazive kolona `Users`/`Stammdaten`), i potvrdi da
`ensureSheetColumns` ne dira popunjene `TRETMAN-*` sheet-ove (ne menjamo ih ni ovde)._
