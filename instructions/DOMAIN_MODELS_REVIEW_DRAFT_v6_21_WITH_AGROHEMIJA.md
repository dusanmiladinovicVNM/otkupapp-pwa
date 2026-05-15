# AgriX / OtkupApp — Domain Models Review Draft

**Status:** Review draft only — not yet approved for insertion into the documentation package  
**Target use:** Human review before merging into `ARCHITECTURE_REFERENCE.md`  
**Current baseline:** v6.21 documentation refactor work  
**Reviewer:** Dušan / architecture owner  

> This document intentionally describes business/domain models in detail.  
> It is not a changelog. It should be read and corrected before being merged into the final package.

---

## 0. Review Rules

### 0.1 Purpose

This draft defines the current domain models that should be explicit in the Architecture Reference:

1. Roba model
2. Novac model
3. Faktura + SEF model
4. Sync model
5. Sledljivost model
6. Ambalaža model
7. Agrohemija / Digitalni Agronom model

### 0.2 Confirmation Required

Items marked `NEEDS CONFIRMATION` must be reviewed before this content is inserted into the package.

### 0.3 Canonical Writing Style

The final AR version should describe:

- what is true now;
- ownership and source of truth;
- IDs and relationships;
- allowed writes;
- side effects;
- rollback/recovery behavior;
- current known limitations.

It should not describe version history except through short references to the changelog.

---

# 1. Roba Model

## 1.1 Scope

The Roba model describes the physical goods moving through the otkup/document chain.

In the current architecture, “roba” primarily appears as fruit/procurement goods represented through:

- `Kultura` / fruit type / variety / crop context;
- `Klasa`;
- `Kolicina`;
- `Cena`;
- `Vrednost`;
- `ParcelaID` where available;
- document rows in `tblOtkup`, `tblOtpremnica`, `tblZbirna`, `tblPrijemnica`, `tblFakturaStavke`.

`NEEDS CONFIRMATION:` Confirm whether “Roba model” in final AR should mean only procured fruit/goods, or whether it should also include Agrohemija `Artikli` and warehouse stock. This draft treats Agrohemija stock as separate and only references it as a parallel article/stock model.

## 1.2 Source of Truth

The canonical desktop source of truth for procured goods is the Excel/VBA workbook.

The field-origin path is:

```text
PWA/desktop capture
  -> OTK/VOZ transport or direct desktop write
  -> tblOtkup
  -> tblOtpremnica / tblZbirna
  -> tblPrijemnica
  -> tblFakturaStavke
```

PWA and Google Sheets can originate or transport goods data, but the canonical long-lived business-document chain is in the desktop workbook.

## 1.3 Goods Identity

A goods row is not identified by a single standalone “RobaID”.

Instead, goods identity is contextual and document-bound:

| Level | Identity |
|---|---|
| Field/procurement capture | `OtkupID` / `ClientRecordID` / `ServerRecordID` |
| Shipment / transport | `OtpremnicaID`, `ZbirnaID`, `BrojZbirne` |
| Receipt / buyer-side document | `PrijemnicaID`, `BrojPrijemnice`, `Klasa` |
| Invoice line | `FakturaStavkaID` linked to `PrijemnicaID` |
| Trace view | `Zbirna -> Otpremnica -> Otkup -> Kooperant/Parcela` |

The final AR should make clear that `PrijemnicaID` is row-unique. `BrojPrijemnice` is a business grouping number and may group more than one physical row when multiple classes exist.

## 1.4 Quantity, Price, and Class Model

Goods rows carry:

- class (`Klasa`);
- quantity (`Kolicina`);
- price (`Cena`);
- derived value (`Kolicina × Cena`);
- fruit/culture context;
- station / kooperant / buyer / driver context depending on the document.

Dual-class capture is supported.

For desktop otkup, `frmOtkup` can capture Klasa I and optional Klasa II as one operator action. That operator action must save through `SaveOtkupMulti_TX`, not through separate independent transactions.

For document flows, Klasa I and Klasa II document rows must commit or rollback together when they are part of one operator action.

## 1.5 Kultura / Sorta / Parcela Context

The goods model is tied to master data:

- `Kultura` / fruit type / variety is selected in PWA or desktop.
- `ParcelaID` may enrich the row where field-origin data is available.
- Desktop `frmOtkup` supports master-data cascades:
  - `VrstaVoca -> SortaVoca`;
  - `OtkupnoMesto -> Kooperanti`;
  - `Kooperant -> Parcele`.
- Selecting a parcela can auto-fill fruit context from parcela kultura.

The final AR should distinguish:

```text
Goods/crop identity = Kultura / Sorta / Klasa
Field origin = ParcelaID
Business document identity = OtkupID / OtpremnicaID / ZbirnaID / PrijemnicaID / FakturaID
```

## 1.6 Otkup as Goods Entry

`tblOtkup` is the canonical starting point for procured goods in the desktop chain.

`SaveOtkup()` must:

- allocate an `OTK-*` ID;
- validate required inputs;
- validate class;
- resolve `KulturaID` when possible;
- persist optional `ParcelaID`;
- persist optional `BrojZbirne`;
- write the row as the beginning of the procurement chain.

Desktop otkup may also create side effects:

- packaging movement through `TrackAmbalaza`;
- cash payout / money row where entered;
- avans allocation when applicable.

These must be part of the same high-level business operation in `SaveOtkupMulti_TX`.

## 1.7 Otpremnica / Zbirna / Prijemnica Goods Movement

Goods movement from procurement into buyer-facing receipt/invoice flow is represented through document chain rows:

```text
Otkup
  -> Otpremnica
  -> Zbirna
  -> Prijemnica
  -> FakturaStavke
```

Important current rules:

- Otpremnica may be created before Zbirna and may initially have empty `BrojZbirne`.
- `BrojZbirne` is the business transport number and is PWA-first in normal driver flow.
- VBA `GenerateBrojZbirne` remains fallback for legacy/desktop-manual cases.
- `PrijemnicaID` is the row-level invoiceable receipt identity.
- `BrojPrijemnice + Klasa` is the relink identity for class-specific relink cases.

## 1.8 Roba and Faktura Lines

Faktura line values must not be trusted from UI payload.

`CreateFaktura()` uses selected `PrijemnicaID` values, then derives invoice line values from `tblPrijemnica`:

- `Kolicina`;
- `Cena`;
- `Klasa`;
- `BrojPrijemnice`.

This prevents caller/UI tampering from changing invoice values.

## 1.9 Roba Model Invariants

The final AR should include these invariants:

1. A goods row is document-contextual, not a free-floating item row.
2. `tblOtkup` is the desktop canonical procurement entry point.
3. Invoiceable goods are represented by `tblPrijemnica` rows.
4. Faktura lines must derive quantity/price/class from canonical prijemnica rows.
5. Dual-class operator actions commit/rollback atomically.
6. Stornirano rows must be excluded from active goods/reporting views.
7. Traceability must preserve the path from invoice/receipt back to procurement and, where available, parcela.

---

# 2. Novac Model

## 2.1 Scope

The Novac model describes all money movement and money allocation in the system.

It covers:

- manual finance entry;
- kooperant/supplier payout;
- buyer payment;
- OM/station cash flow;
- avans;
- BankaImport / BankaMapiranje reconciliation;
- payment status recomputation for otkup and faktura.

## 2.2 Source of Truth

`tblNovac` is the canonical money ledger.

Money can enter the system through:

1. manual desktop entry;
2. desktop otkup with cash payout;
3. bank import staging and mapping;
4. buyer payment;
5. supplier/kooperant payout;
6. OM/station cash movement;
7. avans allocation flows.

## 2.3 Ledger Direction

A valid `tblNovac` row must have exactly one positive money direction:

```text
Uplata  > 0 and Isplata = 0
or
Isplata > 0 and Uplata  = 0
```

Invalid cases:

- both directions positive;
- both directions empty/zero;
- negative money amounts;
- missing `Tip`;
- missing required partner/entity context for the money type.

## 2.4 Partner / Entity Context

A money row may be linked to:

- `KupacID`;
- `KooperantID`;
- `OMID`;
- `FakturaID`;
- `OtkupID`;
- BankaImport source context.

The model must make clear that not every money row is immediately linked to a final settlement target. Some rows are avans and become allocated later.

## 2.5 Buyer-Side Money Model

Buyer-side money flow:

```text
buyer payment
  -> match to open Faktura when unique
  -> otherwise record buyer avans
  -> allocate avans to Faktura later
  -> recompute Faktura status
```

Canonical behaviors:

- `GetOpenFakture()` defines open buyer invoice candidates by outstanding remainder, not by status text alone.
- `ApplyAvansToFaktura[_TX]()` allocates buyer avans to individual fakture.
- Full avans consumption links the existing avans row.
- Partial avans consumption reduces the original avans row and creates a linked split row.
- `UpdateFakturaStatus()` recomputes payment state from active linked uplata.

## 2.6 Supplier / Kooperant-Side Money Model

Supplier/kooperant-side money flow:

```text
supplier/kooperant payment
  -> match to one or more open Otkup rows
  -> create/link money rows to specific OtkupID
  -> if excess remains, record kooperant avans
  -> recompute Otkup status
```

Canonical behaviors:

- `GetOpenOtkupi()` defines open payable candidates by outstanding remainder.
- `ApplyAvansToOtkup[_TX]()` allocates avans to individual otkup rows.
- Full avans consumption links the existing avans row.
- Partial avans consumption reduces the original avans row and creates a linked split row.
- `UpdateOtkupStatus()` recomputes paid state from linked `Isplata`.

## 2.7 BankaImport and Novac

BankaImport is not itself the final finance ledger.

Flow:

```text
bank PDF / bank statement
  -> tblBankaImport staging
  -> BankaMapiranje classification
  -> tblNovac ledger row(s)
  -> Faktura/Otkup/OM/avans link
```

For buyer payments:

- if one unique faktura is resolved, mapping creates buyer payment linked to `FakturaID`;
- if no unique faktura exists, mapping creates buyer avans.

For kooperant block payments:

- mapping finds open otkup candidates;
- one bank payment may be split across multiple otkup rows;
- each allocation must create/link one `tblNovac` row to a specific `OtkupID`;
- remaining amount becomes kooperant avans.

Each reconciled `tblNovac` row must preserve traceability to the BankaImport source, including generated note/metadata such as `BIM:<id>` and relevant bank reference/match reason context.

## 2.8 Status Recompute

Status recomputation must be bidirectional.

For otkup:

- sufficient linked `Isplata` sets `Isplaceno`;
- `DatumIsplate` is filled if empty;
- insufficient linked payment clears the paid state and date.

For faktura:

- sufficient active linked uplata closes the faktura;
- `DatumPlacanja` is filled only on first closure;
- if payment is removed/stornirano/relinked and the amount becomes insufficient, the faktura reopens and payment date is cleared according to the active recompute rule.

## 2.9 Partner Map

`tblPartnerMap` stores learned exact bank-side partner mapping.

Rules:

- identical existing mapping = idempotent success;
- conflicting mapping for the same bank partner name = fail-fast;
- learned mapping must not silently override an existing partner/entity/OM mapping.

## 2.10 Novac Transaction and Rollback Model

`SaveNovac_TX()` snapshots:

- `tblNovac`;
- `tblFakture`;
- `tblOtkup`.

It delegates to `SaveNovac()` for append-oriented row creation.

Any operation that modifies money links and dependent document statuses must commit or rollback together.

External effects, such as bank file moves or Google writebacks, are not part of the Excel table rollback model.

## 2.11 Novac Model Invariants

1. `tblNovac` is the canonical money ledger.
2. Money rows must have exactly one direction: `Uplata` or `Isplata`.
3. Avans is not an error; it is unallocated money.
4. Buyer avans allocates to `FakturaID`.
5. Kooperant/supplier avans allocates to `OtkupID`.
6. Partial avans consumption must split/reduce rows explicitly.
7. Otkup and faktura payment statuses are derived from active linked money, not from manual status-only flags.
8. Stornirano rows are excluded from live finance aggregates.
9. BankaImport mapping must leave audit trace back to the staged bank row.

---

# 3. Faktura + SEF Model

## 3.1 Scope

The Faktura + SEF model describes invoice creation, invoice lines, buyer payment status, electronic invoice submission and SEF lifecycle state.

It covers:

- `tblFakture`;
- `tblFakturaStavke`;
- `tblPrijemnica`;
- `tblNovac`;
- `tblSEFSubmission`;
- `tblSEFEventLog`;
- SEF local workflow state;
- external SEF API status.

## 3.2 Faktura Source of Truth

`tblFakture` is the canonical invoice header table.

`tblFakturaStavke` is the canonical invoice line table.

Invoice line data is based on `tblPrijemnica`, not UI-entered line payload.

## 3.3 Faktura Creation

`CreateFaktura_TX()` must snapshot:

- `tblFakture`;
- `tblFakturaStavke`;
- `tblPrijemnica`;
- `tblNovac`.

It delegates to `CreateFaktura()` and must rollback:

- faktura header;
- faktura lines;
- prijemnica linkage;
- buyer-avans side effects.

## 3.4 Invoiceability Rules

Faktura creation must fail fast when selected prijemnice include:

- duplicate `PrijemnicaID`;
- stornirana prijemnica;
- already-fakturisana prijemnica;
- prijemnica already carrying `FakturaID`.

## 3.5 Canonical Faktura Line Values

`CreateFaktura()` trusts caller-selected stavke only for `PrijemnicaID`.

It derives from `tblPrijemnica`:

- `Kolicina`;
- `Cena`;
- `Klasa`;
- `BrojPrijemnice`.

This rule protects invoice totals from UI/caller tampering.

## 3.6 Faktura Payment Status

`UpdateFakturaStatus()` recomputes status from active linked payments.

Rules:

- sufficient active uplata closes the faktura;
- `DatumPlacanja` is preserved if already populated;
- `DatumPlacanja` is filled only on first closure;
- insufficient active uplata reopens the faktura and clears payment date;
- stornirana faktura rows are skipped.

## 3.7 Print Boundary

`PrintFaktura()` must not print active stornirana invoices as normal active invoices.

Any archival reprint of stornirana invoices must be a separate, clearly marked workflow.

## 3.8 SEF Local vs External State

The SEF model separates:

| Field / concept | Meaning |
|---|---|
| `SEFWorkflowState` | local/internal process control |
| `SEFStatus` | latest exact external SEF/API status |
| `SEFDocumentId` | stable SEF document identity |
| `tblSEFSubmission` | request/response submission journal |
| `tblSEFEventLog` | lifecycle/audit event log |

`SEFStatus` must store the external API status, not internal workflow constants.

## 3.9 SEF Submission Model

`SendInvoiceToSEF_TX(fakturaID)` uses a multi-phase model:

```text
1. local preparation transaction
2. transition/submission-row transaction
3. remote SEF HTTP call outside transaction
4. final result persistence transaction
```

The remote HTTP call is outside the Excel transaction scope.

Local state and local result persistence are transactional; the remote call itself cannot be rolled back.

## 3.10 SEF Submission Identity

The request identity rule is:

```text
requestId = submissionID
```

Technical-failure retry may reuse the previous submission body and payload hash when `ShouldReuseLastSubmission()` allows it.

## 3.11 Successful Submit Guard

A successful HTTP/API response without `SEFDocumentId` is invalid for successful workflow progression.

It must be normalized into a technical failure with evidence such as `MISSING_SEF_DOCUMENT_ID`.

It must not produce a stable successful state such as `WF_SEF_SENT` or `WF_SEF_ACCEPTED`.

## 3.12 SEF DTO / UBL Validation

Before outbound submission:

- header totals must match line totals within tolerance;
- net, VAT and gross totals must be consistent;
- tax percent must come from the active tax configuration source;
- hardcoded 10% VAT is not canonical;
- `DeliveryDate` must not be after `InvoiceDate`;
- config must validate `SEF_BASE_URL`, `SEF_API_KEY` and HTTPS scheme.

## 3.13 SEF Refresh and Final-State Protection

Refresh operations may update external `SEFStatus` without changing local `SEFWorkflowState`.

Rules:

- repeated refresh of compatible states must be idempotent;
- final local states must not move backwards because external API returns pending-like values;
- `SEF_SYNC_ERROR` may recover through allowed transition when later refresh returns valid final status;
- stuck `SEF_SENDING` rows are recovered through guarded recovery helpers.

## 3.14 SEF Corrective Actions

Remote corrective actions include:

- cancel;
- storno.

These are not ordinary submit transitions.

They call dedicated SEF endpoints and then update refresh fields/event logs.

## 3.15 Faktura + SEF Invariants

1. Faktura lines derive from `tblPrijemnica`.
2. `PrijemnicaID` is row-unique and invoiceable.
3. Duplicate/stornirano/already-fakturisano prijemnice are blocked.
4. Faktura payment status is derived from active linked novac.
5. `SEFWorkflowState` and `SEFStatus` have different meanings and must not be collapsed.
6. `SEFDocumentId` is mandatory evidence for successful remote identity.
7. SEF HTTP calls are outside Excel transaction scope.
8. SEF submission and event history must remain auditable in `tblSEFSubmission` and `tblSEFEventLog`.

---

# 4. Sync Model

## 4.1 Scope

The Sync model describes synchronization between:

- PWA IndexedDB;
- GAS API;
- Google Sheets transport layer;
- Excel/VBA desktop master.

It covers:

- `ClientRecordID`;
- `ServerRecordID`;
- `SyncStatus`;
- role-specific local stores;
- role-specific Google Sheets;
- MasterSync import/writeback;
- offline-first retry behavior.

## 4.2 Layer Ownership

| Layer | Ownership |
|---|---|
| PWA | local capture, offline queue, local sync state, submit locks |
| IndexedDB | local source for pending/failed/synced client rows |
| GAS | auth, endpoint authorization, idempotent append/update, batch response |
| Google Sheets | transport/projection state, per-role sheet rows |
| VBA/Excel | canonical desktop master import, document/finance/SEF ownership |

## 4.3 Identity Fields

### ClientRecordID

`ClientRecordID` is the client-generated idempotency identity for PWA-created records.

It is required to prevent duplicate append on retry.

### ServerRecordID

`ServerRecordID` is the technical server/master identity written back after GAS/VBA accepts or imports a row.

It must not be used as a business document number.

### Business Document Numbers

Examples:

- `BrojZbirne`;
- `BrojPrijemnice`;
- `BrojFakture`.

Business numbers must remain separate from technical sync IDs.

## 4.4 SyncStatus

`SyncStatus` describes row state across local/GAS/Sheets/VBA boundaries.

Common states include:

- `pending`;
- `syncing`;
- `synced`;
- `duplicate`;
- `existing`;
- `inserted`;
- `updated`;
- `SyncError`;
- `Synced>Master`.

`NEEDS CONFIRMATION:` Confirm the exact canonical list of Google-side `SyncStatus` strings used by current `Code.gs` and VBA MasterSync.

## 4.5 Canonical PWA Sync Result Shape

Role sync entrypoints return a normalized shape:

```js
{
  ok: true | false,
  role: 'Otkupac' | 'Kooperant' | 'Vozac',
  synced: number,
  failed: number,
  results: [],
  reason: '',
  code: '',
  partial: false
}
```

All app-level role sync paths should converge to this shape.

## 4.6 Single Sync Entrypoint

All sync triggers route through:

```js
syncQueueSafe(reason)
```

Supported reasons:

- `manual`;
- `online`;
- `interval`;
- `post-save`.

Low-level sync functions must not be called directly by UI triggers.

## 4.7 Offline-First Behavior

PWA capture must work offline.

Offline-created records:

1. are written to IndexedDB;
2. receive `ClientRecordID`;
3. remain `pending`;
4. sync later when network/backend is available;
5. update local state from backend result.

If a sync request fails, all attempted batch records must return to retryable `pending` state with diagnostics, not remain permanently stuck.

## 4.8 Stale Syncing Recovery

Records can become stuck in:

```js
syncStatus: 'syncing'
```

after refresh, crash, browser kill, deploy, or network interruption.

Bootstrap must recover stale syncing records for the current role before role render/sync badge calculation.

Recovered rows become:

```js
syncStatus: 'pending'
lastServerStatus: 'stale-syncing-recovered'
```

## 4.9 Render Dedupe

Before rendering key UI lists, merged local/server data must pass through alias-aware dedupe.

Identity aliases:

```text
srv:<serverRecordID>
cli:<clientRecordID>
```

Local priority records win over stale server echoes when they are pending, syncing or carry a sync error.

## 4.10 Role Sync Scope

### Otkupac

Syncs otkup records through OTK sheet/API path.

### Kooperant

Syncs:

- tretmani;
- troskovi.

### Vozač

Syncs zbirne.

`BrojZbirne` is generated PWA-side at `confirmZbirna` time in the normal flow.

### Management

Management is primarily read/report/planning and endpoint-driven.

Normal role sync may return `no-sync-for-role`.

## 4.11 Google Sheets Transport

Google Sheets are transport/projection state, not the ultimate business source of truth for desktop finance/document logic.

Important writeback rules:

- OTK sheet column B receives `ServerRecordID`;
- OTK sheet column F receives `SyncStatus`;
- VOZ sheet column B receives `ServerRecordID` / `ZbirnaID`;
- VOZ sheet column F receives `SyncStatus`;
- VOZ sheet column T receives business `BrojZbirne`.

`ServerRecordID` must not be copied into `BrojZbirne`.

## 4.12 MasterSync Guard

During desktop MasterSync:

- GAS write endpoints can be blocked;
- PWA checks master-sync state;
- `MASTER_SYNC_LOCK` / `SyncControl` prevents unsafe concurrent writes;
- soft-lock responses must be retryable, not destructive.

## 4.13 Sync Invariants

1. PWA is offline-first.
2. `ClientRecordID` is the idempotency anchor.
3. `ServerRecordID` is a technical sync/master ID.
4. Business document numbers are separate from technical sync IDs.
5. Sync retries must not duplicate rows.
6. Failed/incomplete sync attempts must return rows to retryable state.
7. MasterSync lock must prevent unsafe concurrent writes.
8. Google Sheets writeback is external and cannot be rolled back by Excel table transactions.

---

# 5. Sledljivost Model

## 5.1 Scope

The Sledljivost model describes traceability from procurement to shipment, receipt, invoice and back to origin context.

It covers:

- `tblOtkup`;
- `tblOtpremnica`;
- `tblZbirna`;
- `tblPrijemnica`;
- `tblFakturaStavke`;
- `ParcelaID`;
- `BrojZbirne`;
- `OtpremnicaID`;
- `modSledljivost`.

## 5.2 Canonical Trace Chain

The canonical trace chain is:

```text
Faktura
  -> FakturaStavke
  -> Prijemnica
  -> Zbirna / Otpremnica
  -> Otkup
  -> Kooperant
  -> Parcela, when available
```

Reverse transport trace is:

```text
BrojZbirne
  -> Zbirna
  -> Otpremnica
  -> Otkup
  -> Kooperant / BPG / Parcela
```

## 5.3 Trace Bridge

`tblOtkup.OtpremnicaID` is the desktop bridge from raw/PWA otkup rows into the canonical document chain.

`BrojZbirne` is part of the preferred trace key and prevents same-day same-driver cross-zbirna collisions.

## 5.4 Auto-Link Model

`AutoLinkOtkupOtpremnica()` scans active otkup rows with blank `OtpremnicaID`.

Primary strict key:

```text
StanicaID | Datum | VozacID | Klasa | BrojZbirne
```

Legacy fallback key:

```text
StanicaID | Datum | VozacID | Klasa
```

Fallback is allowed only where `BrojZbirne` is genuinely missing on the relevant side and the candidate is unique.

Cross-`BrojZbirne` links are invalid and must remain unresolved for manual repair.

## 5.5 Checked Update Rule

Updating `tblOtkup.OtpremnicaID` after auto-link requires exactly one matching `OtkupID` row.

These are data-integrity errors:

- no match;
- duplicate `OtkupID`;
- failed update.

Link writes use checked/fail-fast update semantics.

## 5.6 Manual Unresolved Queue

`GetUnlinkedOtkupi()` is the canonical read model for procurement rows still lacking `OtpremnicaID`.

It should return compact fields required for manual review.

## 5.7 TraceByZbirna

`TraceByZbirna(brojZbirne)` walks active relations and returns enriched trace output:

```text
Zbirna
  -> Otpremnica
  -> Otkup
  -> Kooperant
  -> BPG
  -> Parcela metadata
  -> Klasa
  -> shipment context
```

## 5.8 PWA-First / VBA-Fallback Rule

PWA is the preferred field source for modern field-created otkup blocks and driver/zbirna context.

VBA remains:

- fallback;
- repair layer;
- canonical desktop chain builder;
- manual/exception entry surface.

## 5.9 Sledljivost Invariants

1. Traceability links must be deterministic.
2. `BrojZbirne` is part of the preferred trace key.
3. Cross-zbirna auto-linking is invalid.
4. Ambiguous links remain unresolved; they are not guessed.
5. `COL_OTK_BROJ_ZBIRNE` is the canonical constant; hardcoded `"BrojZbirne"` should not be used in `modSledljivost`.
6. Auto-link TX must rollback on failure.
7. Monitoring is best-effort after commit and fail-path before rollback.

---

# 6. Ambalaža Model

## 6.1 Scope

The Ambalaža model describes packaging movement and balances.

It covers:

- `tblAmbalaza`;
- `modAmbalaza`;
- `TrackAmbalaza`;
- `GetAmbalazeStanje`;
- `GetVozacAmbSaldo`;
- packaging side effects from otkup, otpremnica, prijemnica and other document flows.

## 6.2 Source of Truth

`tblAmbalaza` is the canonical packaging ledger.

It stores movement rows, not just current balances.

Current balances are derived by summing active ledger movements.

## 6.3 Movement Direction

`Smer` accepts only:

```text
Ulaz
Izlaz
```

Semantics:

```text
Ulaz  = +Kolicina
Izlaz = -Kolicina
```

Unknown direction values are fail-fast errors.

They must not be interpreted as inbound/outbound by default.

## 6.4 Movement Validation

`TrackAmbalaza` rules:

- `kolicina = 0` is a legal no-op;
- `kolicina < 0` is invalid;
- positive movements require:
  - `tipAmb`;
  - `entitetID`;
  - `entitetTip`;
- `Smer` must be `Ulaz` or `Izlaz`;
- `GetNextID` empty result is a hard error;
- `AppendRow <= 0` is a hard error.

## 6.5 Transaction Boundary

`TrackAmbalaza` is a base writer used by larger document transactions.

It must not introduce nested `TrackAmbalaza_TX` for existing document flows.

Caller transactions such as otkup, otpremnica, prijemnica and kupci izlaz must snapshot `tblAmbalaza` where packaging side effects exist.

## 6.6 Otkup Packaging Side Effect

During procurement save, when `KolAmbalaze > 0`, `SaveOtkup()` writes a packaging movement through `TrackAmbalaza`.

The current model records this as part of the otkup business operation and rollback boundary.

`NEEDS CONFIRMATION:` Confirm business wording for whether this movement is best described as “ambalaža leaves kooperant-side balance” or “ambalaža issued/used at procurement”. The code-level direction is `Izlaz` in the current documentation.

## 6.7 Balance Read Model

`GetAmbalazeStanje` computes balance by entity and `TipAmbalaze`:

```text
Stanje = sum(Ulaz) - sum(Izlaz)
```

`ExcludeStornirano` is required before active packaging balance reads.

## 6.8 Driver Packaging Balance

`GetVozacAmbSaldo` computes driver balance from all active movements matching `VozacID`.

It does not filter by `DokumentTip` unless a future business rule explicitly introduces that filter.

Open-ended date filters are supported:

- `datumOd` only;
- `datumDo` only;
- both;
- neither.

## 6.9 Ambalaža Invariants

1. `tblAmbalaza` is a ledger, not just a balance table.
2. `Ulaz` and `Izlaz` are strict allowed directions.
3. Unknown direction is a data error.
4. Negative packaging quantity is invalid.
5. Zero quantity is a no-op.
6. Active balances are derived from non-stornirano movements.
7. Document flows must snapshot `tblAmbalaza` when they create packaging side effects.
8. Driver balance is all active movements by `VozacID`, not document-type-filtered.
---

# 7. Agrohemija / Digitalni Agronom Model

## 7.1 Scope

The Agrohemija / Digitalni Agronom model describes agricultural chemical articles, warehouse movements, kooperant-issued stock, treatment evidence, treatment quantity semantics, and the boundary between stock truth and agronomical history.

It covers two related but separate domains:

1. **Agrohemija warehouse / article-stock model** — physical stock of agricultural chemical articles issued to or received from suppliers/kooperants.
2. **Digitalni Agronom / treatment model** — kooperant field-treatment evidence, dosage calculation, karenca/history, and local stock validation during treatment entry.

`NEEDS CONFIRMATION:` Confirm whether final AR should keep Agrohemija as a separate model, as this draft does, or fold Agrohemija `Artikli` into a broader Roba/Inventory model. This draft keeps procured fruit/goods and Agrohemija stock separate because their identity, ledger, and business lifecycle are different.

## 7.2 Source of Truth

The canonical desktop source of truth for Agrohemija warehouse movements is the Excel/VBA workbook.

Primary desktop ownership:

- `modAgrohemija` owns warehouse business logic.
- `frmAgrohemija` owns operator UI and basket workflow.
- `tblMagacin` is the canonical warehouse movement ledger.
- `tblArtikli` / article master data define article identity and unit context.
- kooperant-issued stock projections are exported to PWA through Stammdaten / `MagacinKoop`-style read models.

PWA ownership:

- Management PWA can prepare barcode/parcel-aware Agrohemija issuing and signature-backed otpremnica-style confirmation.
- Kooperant PWA can enter treatment evidence and validate treatment quantity against locally loaded issued stock.
- PWA is not the canonical warehouse ledger owner.

GAS ownership:

- `syncTretman` persists treatment evidence/history.
- `saveIzdavanje` persists management issuing through the existing endpoint boundary.
- GAS does not become the desktop warehouse ledger owner in the current architecture.

## 7.3 Article Identity and Unit Semantics

Agrohemija articles are represented as article/master-data records, not as otkup fruit/goods rows.

Article identity should be understood through:

- `ArtikalID`;
- article name / barcode where available;
- unit of measure (`JM` / article unit);
- packaging size where available;
- supplier / dobavljač context where applicable;
- stock movement history in `tblMagacin`.

The final AR should explicitly distinguish:

```text
Procured fruit/goods = Otkup / Prijemnica / FakturaStavke chain
Agrohemija articles = ArtikalID / tblMagacin / issued-stock / treatment-evidence chain
```

Both are physical goods, but they are not the same domain model.

## 7.4 Warehouse Movement Ledger

`tblMagacin` is a movement ledger, not only a current-balance table.

Canonical warehouse movement types:

```text
MAG_ULAZ
MAG_IZLAZ
```

`MAG_ULAZ` represents warehouse receipt/increase.

`MAG_IZLAZ` represents issue/outbound movement, usually toward a kooperant and optionally parcel/treatment context.

Current article stock is derived from active, non-stornirano ledger rows. The system should not rely on manually maintained stock totals when a ledger-derived balance is available.

## 7.5 Desktop `modAgrohemija` Contract

`modAgrohemija` remains the canonical desktop business module for Agrohemija warehouse operations.

`SaveMagacin` owns single-row warehouse journal creation for:

- `MAG_ULAZ`;
- `MAG_IZLAZ`.

`SaveMagacin` must validate:

- required `ArtikalID` / article identity;
- valid movement type;
- positive quantity;
- required kooperant for `MAG_IZLAZ`;
- available stock before outbound issue;
- required schema through fail-fast column guards.

`SaveMagacin` must not silently create invalid stock movement rows.

`SaveMagacin_TX` is the single-row transaction wrapper. It snapshots `tblMagacin`, delegates to `SaveMagacin`, commits on success, and rolls back on hard failure.

Monitoring around `SaveMagacin_TX` is best-effort after commit. Monitoring failure must not turn a committed warehouse save into a false business failure.

## 7.6 Desktop `frmAgrohemija` Basket Model

`frmAgrohemija` owns operator UX for warehouse/agrohemija operations.

It maintains two in-memory baskets:

```text
m_KorpaIzlaz
m_KorpaUlaz
```

Basket commits use one explicit transaction over `tblMagacin`.

Invariant:

```text
one operator basket finish action
=> all basket rows commit
or
=> all basket rows rollback
```

Issue baskets must run an aggregated pre-commit stock check by `ArtikalID`.

This is required because multiple basket lines for the same article can individually look valid but collectively exceed current stock.

Optional UX validation may also block adding a line that would make the current basket exceed available stock, but the pre-commit aggregated check is the canonical safety gate.

Multiple parcel IDs are serialized with semicolon (`;`) separators.

Business modules should raise/return errors. `frmAgrohemija` may show operator-facing `MsgBox` feedback because it is the UI layer.

## 7.7 Current Stock Read Model

`GetMagacinStanje()` is the canonical desktop read model for current article stock.

It must derive stock from active ledger movements and exclude stornirano rows.

Basic stock direction:

```text
MAG_ULAZ  => +Kolicina
MAG_IZLAZ => -Kolicina
```

Unknown movement type is a data-integrity error and should not be silently interpreted.

Reports and UI stock displays should read through this model or an explicitly documented projection of this model.

## 7.8 Kooperant-Issued Stock Projection

Kooperant-issued stock is exposed to PWA as a read/projection model, not as the canonical source of warehouse truth.

Typical projection purpose:

- show kooperant what has been issued;
- validate treatment quantity locally;
- support Digitalni Agronom treatment entry;
- support treatment history and karenca context.

`NEEDS CONFIRMATION:` Confirm the final canonical name and exact shape of the kooperant stock projection used by PWA (`MagacinKoop`, `magacinkoop`, Stammdaten tab/header shape, or equivalent current implementation name).

## 7.9 PWA Management Agrohemija Issuing Model

The Management PWA Agrohemija issuing flow supports article/barcode/parcel-aware issue preparation and signable/printable issue confirmation.

Current model:

- recommendation quantity represents real quantity in the article unit of measure;
- packaging can round the recommendation up to a package-compatible quantity;
- the saved quantity must remain real unit-of-measure quantity, not package count;
- parcel IDs use semicolon (`;`) serialization;
- `izdZavrsi()` opens the printable/signable confirmation modal and does not directly persist final business state;
- final save is protected by `withSubmitLock`;
- modal payload carries a stable client-side issuance identity for display/PDF and future idempotency compatibility;
- `izdReset()` clears cart, selected kooperant, selected article/quantity, parcel list, notes, recommendation state, modal data, and barcode debounce state.

Packaging invariant:

```text
rawQty    = calculated required real quantity
pakovanje = package size in article unit
pakCount  = ceil(rawQty / pakovanje)
finalQty  = pakCount * pakovanje
```

`finalQty` is the saved quantity.

`pakCount` may be displayed as explanatory package information but must not replace `finalQty` as the business quantity.

## 7.10 PWA Kooperant Digitalni Agronom Treatment Model

The Kooperant Digitalni Agronom flow records field-treatment evidence.

Treatment records may include:

- kooperant identity;
- parcela identity;
- treatment date;
- treatment type such as zaštita / prihrana;
- article identity where applicable;
- quantity used;
- unit context;
- note/evidence fields;
- karenca/history context;
- local sync identity (`ClientRecordID`).

The PWA treatment save path must validate treatment quantity against locally loaded kooperant-issued stock where article/stock data is available.

The value persisted in the treatment record must be the same accepted/validated real quantity shown in the UI.

The public save function should be a submit-lock wrapper:

```text
agroSaveTretman()
  -> withSubmitLock("agro:tretman:save", agroSaveTretmanUnlocked)
```

This prevents double-click/double-tap duplicate local treatment records.

## 7.11 GAS Treatment Sync Boundary

`syncTretman` is the active GAS endpoint for treatment evidence/history sync.

Current backend contract:

- action: `syncTretman`;
- allowed roles: Kooperant / Management;
- Kooperant caller must match the scoped kooperant identity;
- payload contains `records[]`;
- processing is protected by `withLock(...)`;
- each record is processed by `processTretmanRecord(...)`;
- idempotency is based on `ClientRecordID`;
- backend returns the canonical batch sync response shape.

Treatment storage remains per-kooperant treatment sheet:

```text
TRETMAN-<KooperantID>
```

Readback uses the current treatment read helper for the scoped kooperant.

## 7.12 GAS `saveIzdavanje` Boundary

`saveIzdavanje` remains the existing Management-side issuing endpoint boundary.

Current model:

- Management-only business operation;
- used by the PWA management Agrohemija issuing flow;
- persists issuing data through the existing GAS/sheet contract;
- client-side submit lock reduces duplicate user-submit risk;
- server-side idempotency by stable client issuance ID is a roadmap item unless confirmed implemented.

`NEEDS CONFIRMATION:` Confirm whether current deployed `saveIzdavanje` is already server-idempotent by client issuance ID. If not, keep this as accepted limitation / roadmap.

## 7.13 Treatment vs Stock Decrement Boundary

Current treatment sync is accepted as evidence/history/karenca sync.

It does not automatically become the canonical warehouse stock-decrement operation unless a future architecture decision explicitly defines:

- whether treatment consumption decrements kooperant-issued stock server-side;
- how treatment storno/correction restores or adjusts stock;
- how retry/idempotency prevents double decrement;
- whether desktop or GAS owns the consumption ledger.

Current invariant:

```text
warehouse issue / stock truth = Agrohemija warehouse model
field treatment evidence      = Digitalni Agronom treatment model
```

Treatment may validate against local issued stock, but validation is not the same as canonical stock decrement.

## 7.14 Reports and Derived Views

Agrohemija reporting includes at least:

- current article stock;
- issuing by kooperant;
- supplier/dobavljač stock state;
- kooperant-issued stock projection;
- treatment history;
- karenca/agronomical views where applicable.

`ReportIzdavanjePoKooperantu()` supports open-ended date filters:

- `datumOd` only;
- `datumDo` only;
- both;
- neither.

`ReportStanjePoDobavljacu()` is the correct public spelling.

The older `ReportStanjePoDoabvljacu()` spelling may remain as a compatibility wrapper, but new code should not call it.

## 7.15 Agrohemija Invariants

1. Agrohemija articles are separate from procured fruit/goods unless final AR explicitly chooses a broader inventory model.
2. `tblMagacin` is a ledger of stock movements, not only a balance table.
3. `MAG_ULAZ` increases stock and `MAG_IZLAZ` decreases stock.
4. Outbound issue requires available stock.
5. Multi-line issue/receipt baskets commit or rollback atomically.
6. Issue basket stock validation must aggregate quantities by `ArtikalID` before commit.
7. Package rounding must save real article quantity, not package count.
8. Treatment records must persist the validated real quantity.
9. `syncTretman` is treatment evidence/history sync, not automatic warehouse decrement unless future architecture says so.
10. `saveIzdavanje` idempotency must remain explicitly documented as implemented, accepted limitation, or roadmap.
11. PWA local issued-stock validation is a UX/safety check, not the canonical warehouse ledger.
12. Stornirano warehouse rows must be excluded from active stock read models.


---

# 8. Suggested AR Insertion Plan

After reviewer approval, insert this content as follows:

| Draft section | Target AR section |
|---|---|
| Roba Model | New `6A. Roba / Goods Model` or inside `6. Document Flow Architecture` before Otkup |
| Novac Model | Replace/expand `7. Finance Architecture` |
| Faktura + SEF Model | Split between `6.5 Faktura` and `6.6 SEF Submission Flow` |
| Sync Model | Expand `11. PWA Architecture`, `9. GAS API Architecture`, and `10. Google Sheets Data Layer` |
| Sledljivost Model | Replace/expand `6.7 Sledljivost` |
| Ambalaža Model | Replace/expand `6.8 Ambalaža Ledger` |
| Agrohemija / Digitalni Agronom Model | Replace/expand `13. Agrohemija / Digitalni Agronom` |

Recommended approach:

1. Keep this file as a standalone reviewer document first.
2. Review business wording and `NEEDS CONFIRMATION` items.
3. Approve or edit each model.
4. Merge into AR in one controlled future package pass.
5. Add model-specific gates to `RELEASE_GATES.md` only after model wording is confirmed.

---

# 9. Open Confirmation List

## 9.1 Roba Meaning

Confirm whether `Roba model` means:

```text
A. only fruit/procurement goods in otkup/document/faktura chain
B. fruit/procurement goods + Agrohemija articles/stock
C. a broader inventory model
```

## 9.2 Ambalaža Direction Wording

Confirm the business-language wording for otkup packaging side effect.

The current code-level model says otkup packaging writes `TrackAmbalaza(..., "Izlaz", kooperantID, "Kooperant", ..., DOK_TIP_OTKUP)`.

Need confirmation of business wording for final AR.

## 9.3 SyncStatus Canonical List

Confirm exact current Google/VBA/PWA status strings in code.

## 9.4 Faktura + SEF Details

Confirm whether the final AR should include all SEF workflow constants or keep them in SEF implementation docs and only keep model-level rules in AR.

## 9.5 Novac Type Vocabulary

Confirm exact canonical `Tip` / novac constants that should be listed in the final AR.



## 9.6 Agrohemija Scope

Confirm whether Agrohemija should remain a separate domain model or be folded into a broader Roba/Inventory model.

Recommended current wording: keep it separate, because fruit/procurement goods and Agrohemija articles have different identity, ledger, and lifecycle rules.

## 9.7 Kooperant Stock Projection Name

Confirm exact canonical PWA/Stammdaten projection name and headers for kooperant-issued Agrohemija stock.

Candidate names seen in documentation/code wording include:

```text
MagacinKoop
magacinkoop
Stammdaten MagacinKoop projection
```

## 9.8 saveIzdavanje Idempotency

Confirm whether `saveIzdavanje` is currently server-idempotent by stable client issuance ID.

If not implemented, keep it as roadmap / accepted limitation and rely on client submit-lock as the current launch mitigation.

## 9.9 Treatment Stock Decrement

Confirm whether treatment sync is intentionally evidence/history-only or whether any current deployed path decrements kooperant-issued stock.

Recommended current wording: treatment sync is evidence/history/karenca; stock decrement is not automatic unless explicitly implemented and documented.
