# AgriX / OtkupApp — Procena zrelosti sistema (Maturity Assessment)

**Datum procene:** 2026-06-26
**Verzija arhitekture u trenutku procene:** AR v6.31 (canonical) · release `vba-v2.4.0`
**Metod:** statička analiza koda + pregled dokumenata stanja (AR, ROADMAP, KNOWN_ISSUES, RELEASE_GATES, RELEASE_NOTES, git istorija)
**Companion:** `ARCHITECTURE_REFERENCE.md`, `KNOWN_ISSUES.md`, `ROADMAP.md`, `RELEASE_GATES.md`

> Ovo je snapshot procene na navedeni datum. Otvorene stavke su linkovane na svoje
> kanonske registre (KNOWN_ISSUES / ROADMAP). Procena je vezana za **stanje koda i
> dokumenata**, ne za starost git repozitorijuma — interno verzionisanje (v2.2 → v6.31)
> i obim koda pokazuju sistem koji postoji znatno duže od git istorije.

---

## 0. Zaključak (TL;DR)

AgriX je **funkcionalno zreo, operativno ozbiljno postavljen sistem u kasnoj fazi
(production-bound / late-beta)** — nije prototip, ali ni „mirno potpisan GO".
Po inženjerskom procesu je iznad proseka za VBA/Excel klasu alata; glavni jaz je što
**runtime gateovi još nisu izvršeni na ciljnom workbook-u** i postoji nekoliko
otvorenih security/QA stavki.

---

## 1. Obim sistema (mereno)

| Metrika | Vrednost |
|---|---|
| VBA `.bas` moduli | 83 |
| VBA `.cls` klase | 11 |
| VBA forme (`.frm`) | 16 |
| Ukupno linija VBA | ~75.000 |
| PWA + GAS JavaScript | ~19.000 linija |
| Test moduli (VBA) | 20 |
| Test rutine (`Test*` Sub) | ~132 |
| Production runbookovi | 15 |
| Release tagovi | do `vba-v2.4.0` |
| Interna arhitekturna verzija | v6.31 |

**Pokriveni domeni:** Otkup → Otpremnica → Zbirna → Prijemnica → Faktura → SEF,
Novac/avansi, Banka import/export (PDF/clipboard parser), Agrohemija, Ambalaža/Palete
ledger, Sledljivost/karenca, Geo/parcele (GIS), Licenciranje po uređaju, Meteo, ML,
Monitoring/Fleet, Sync (Google / MasterSync / Stammdaten).

**PWA role:** Otkupac, Kooperant, Vozač, Management — sa offline queue, service worker,
render dedupe i submit-lock mehanizmima.

---

## 2. Scorecard po dimenzijama

| Dimenzija | Ocena | Obrazloženje |
|---|---|---|
| Funkcionalni obim | **5.0 / 5** | Pokriven ceo poslovni lanac + finansije + SEF e-faktura + GIS + licenca |
| Arhitektura / slojevitost | **4.5 / 5** | Jasni slojevi: `modConfig` / `modDataAccess` / `clsTransaction` (rollback), `modSchemaGuard`, `modJournaling`, `modMonitoring`. Anti-duplication doktrina aktivno održavana |
| Release engineering | **4.5 / 5** | `tools/release.sh`, `modBuildInfo` (SHA/verzija), Fleet telemetrija, min-version gate (`modUpdateGate`), blanko-build guard |
| Dokumentacija | **5.0 / 5** | Canonical AR + CHANGELOG + ROADMAP + KNOWN_ISSUES + RELEASE_GATES + 15 runbookova |
| Test / QA automatizacija | **3.0 / 5** | 20 test modula, ~132 rutine, E2E gate, `RunProductionHealthCheck`. Pokrivenost neravnomerna (TL-004); CI ne pokreće Excel → finalni smoke manuelan |
| Launch readiness | **3.0 / 5** | Otvoreni `NEEDS REVIEW` / `OPEN` blokeri (vidi §4) |
| Security | **3.0 / 5** | `saveParcelPolygon` auth nerešen (KI-001); endpoint matrica neusaglašena sa deployovanim `Code.gs` (KI-005) |

**Agregat (neponderisan):** ~4.0 / 5 → *zreo sistem sa otvorenim launch-readiness repom.*

---

## 3. Šta vuče zrelost GORE

- **Realne isporuke + fleet telemetrija.** Instalacioni folderi (`bucaijoca`, `bukovik`,
  `venivo`) = aktivne instalacije; Fleet pregled „ko ima koju verziju" preko GAS
  `Events`/`Fleet` + `rebuildMonitoringFleet`.
- **Transakcioni model sa journal/backup/autosave** (`clsTransaction` + `modJournaling`)
  i schema guard (`modSchemaGuard`) — ozbiljnije od tipične VBA aplikacije.
- **Disciplinovan release proces** — gateovi selektovani po izmenjenoj površini
  (VBA / GAS / PWA / Monitoring), accepted-risks registar sa vlasnikom/datumom.
- **Bogata operativna dokumentacija** — 15 production runbookova za realne incident scenarije
  (banka/novac, document-chain, fiskalni lager, GAS auth, GIS, licenca, offline sync, SEF…).

## 4. Šta vuče zrelost DOLE (otvoreni rizici)

| ID | Stavka | Status | Link |
|---|---|---|---|
| KI-002 | Production readiness ciljnog workbook-a nije dokazan | OPEN | KNOWN_ISSUES §1 |
| KI-003 | Runtime gateovi nisu izvršeni (compile/smoke/route/PWA/monitoring/SEF) | OPEN | RELEASE_GATES |
| KI-001 | `saveParcelPolygon` autorizacija nedosledna između izvora | NEEDS REVIEW | AR §§9,14,16,19 |
| KI-005 | Endpoint matrica neusaglašena sa deployovanim `Code.gs` | NEEDS REVIEW | AR §9 |
| KI-004 | Naziv canonical Banka PDF parser modula nejasan u izvorima | NEEDS REVIEW | AR §§8,19 |
| TL-001 | SEF JSON parser ostaje manuelan/lagan | Active limit | ROADMAP §2.2 |
| TL-002 | Exact-row guardi još lokalni (`modBankaMapiranje`, `modStorno`) | Tech debt | ROADMAP §2.1 |
| TL-003 | Residualni business-layer `MsgBox` | Cleanup | ROADMAP §2.6 |
| TL-004 | Neravnomerna negativna test pokrivenost (PWA/GAS/VBA) | Active limit | — |

**Strukturni rizik:** schema drift po instalaciji — realne kolone variraju, izvor istine
je šema (ne kod); pozicijski `AppendRow` bezbedan samo uz potvrđen redosled kolona.

**CI ograničenje:** Excel se ne kompajlira/pokreće u CI okruženju → verifikacija je
statička, finalni smoke radi operater u Excelu.

---

## 5. Put do „GO" (production handoff)

Iz `KNOWN_ISSUES.md` §5 — blokeri između trenutnog stanja i čistog handoff-a:

- [ ] Razrešiti/eksplicitno prihvatiti svaki `NEEDS REVIEW` (KI-001, KI-004, KI-005)
- [ ] Potvrditi deployovano `saveParcelPolygon` auth stanje (KI-001)
- [ ] Potvrditi canonical naziv Banka PDF parser modula (KI-004)
- [ ] `RunProductionHealthCheck` čist na ciljnom workbook-u (KI-002)
- [ ] Izvršiti obavezne runtime gateove iz `RELEASE_GATES.md` (KI-003)
- [ ] Upisati svaki prihvaćeni rizik sa vlasnikom/datumom u ROADMAP ili KNOWN_ISSUES

---

## 6. Metodološka napomena

Git istorija (203 commita od 2026-06-19) **nije pokazatelj zrelosti** ovog sistema:
git repo je novijeg datuma za aplikaciju koja interno ide do v6.31 / `vba-v2.4.0`.
Procena je namerno vezana za stanje koda, dokumenata i otvorenih registara —
ne za starost repozitorijuma.
