# 09A — Portfolio proizvoda i launch-readiness (razrada QA Decision Log-a)

**Status:** Review
**Vlasnik:** osnivač AgriX-a
**Poslednje ažuriranje:** 2026-07-24 (v2 — produbljena analiza: code-verifikacija + tržište/konkurencija)
**Sidro koda:** `origin/main` v2.24.0 (`9fd7087`)
**Povezani dokumenti:** `09_QA_DECISION_LOG.md`, `07_PRODUCT_PORTFOLIO.md`, `07A_PRODUCT_STATUS_MATRIX.csv`, `07B_ENTERPRISE_OPERATING_MODES.md`, `08_PRODUCT_ROADMAP.md`, `03_CUSTOMERS_AND_JOBS.md`, `04_MARKET.md`, `05_COMPETITION.md`, `05A_COMPETITOR_EVIDENCE.md`, `05B_INFOSYS_REPLACEMENT_GTM.md`, `06_POSITIONING.md`, `../PLAN_SANACIJE.md`, `../REFAKTOR_PLAYBOOK.md`, `../KNOWN_ISSUES.md` §8, `../RELEASE_GATES.md`

> **Cilj poglavlja.** Razraditi `09_QA_DECISION_LOG.md` (260 komercijalnih odluka) u **prodajni portfolio** i dati **launch-readiness presudu** — gde je najviše posla da _sve_ bude spremno za lansiranje.
>
> **v2 dodaje dubinu:** (1) **code-verifikaciju** top blokera protiv stvarnog `src-vba` koda — registar defekata je delom **zastareo u odnosu na deploy v2.24.0**; (2) per-defekt razradu svih 9 P0 + P1 klastera sa scenarijem otkaza i procenom napora; (3) tržište/konkurenciju/pozicioniranje sa realnim brojevima.
>
> **Ovaj dokument NE duplira izvore** — spaja komercijalni portfolio (`09`, `07`, `03–06`) i inženjerski backlog (`KNOWN_ISSUES §8`, `PLAN_SANACIJE`, `REFAKTOR_PLAYBOOK`) u jednu presudu.

---

## 1. Potvrđene polazne činjenice

**Proizvod / klijenti (FACT):**
- **3 klijenta** u produkciji na Desktop-u (silosi u repou: `venivo/`, `bucaijoca/`, `bukovik/`); management koristi PWA.
- Prosečna firma: **~10 stanica**, 1 desktop korisnik (single-active-user, QA #137), **~100 kooperanata**.
- Kod: `origin/main` **v2.24.0**; funkcionalno bogat (storno kaskade, recovery paneli, palete, izveštaji, self-update, ASCII lokalizacija, fleet monitoring).
- PWA zrelost (broj linija koda): **Management ~5.100** · **Kooperant/Gazdinstvo ~4.850** · **Otkupac ~3.710** · **Vozač ~735** (najtanji — poklapa se sa „osnovni Vozač je Core, napredni Dispatch opcion").

**Tržište (MEASURED, APR 22.07.2026, delatnosti 1039+4631):**
- **1.516 aktivnih firmi**; medijana prihoda **23,0M RSD**, 90. percentil **398,3M RSD**.
- Koncentracija: **Top 100 firmi = 66,8% prihoda** (Top 50 = 52,7%). → account-based prodaja na Top 100–300, ne masovni marketing.

**Inženjering (triaž FM v35→v85→v142 + review):** 48 AUD stavki → **30 RF paketa**; nominalno **9 P0, ~34 P1, 5 P2**. **Urađen 1 (RF-27).** Runtime gate-ovi **nisu izvršeni** na ciljnom workbook-u (KI-002/003).

**Cene (jedini konkretni brojevi):** Gazdinstvo **Basic 19€ / Pro 39€**; prvih 50 Partner uz paket hladnjače. Sve ostalo su _pravila_, ne iznosi (`10_PRICING` nije napisan).

---

## 2. Razrada QA log-a → komercijalna arhitektura

Jedan proizvod (QA #59: nema small/mid/large izdanja); razlike kroz **pakete + module + broj stanica + konfiguraciju**. Pet stubova nad **Platform Services**:

```
                    ┌─────────────────────────────────────────────┐
                    │        AgriX Platform Services              │  (u svakom plaćenom proizvodu)
                    │  auth · offline queue · sync/MasterSync ·   │
                    │  storno/audit · monitoring · self-update ·  │
                    │  backup/recovery · release gates            │
                    └─────────────────────────────────────────────┘
   ┌───────────────────────┐   ┌──────────────────┐   ┌──────────────────┐
   │   AgriX ENTERPRISE     │   │  AgriX GAZDINSTVO │   │   AgriX SAVETNIK  │
   │  (glavni B2B proizvod) │   │  Partner /        │   │  1 agronom →      │
   │  Paketi: Desktop /     │   │  Basic 19€ /      │   │  više gazdinstava │
   │  Mobile                │   │  Pro 39€          │   │  (po gazdinstvu)  │
   │  Moduli: SEF · Banka · │   └──────────────────┘   └──────────────────┘
   │  Dispatch · Hladnjača  │            ▲
   └───────────────────────┘            │ compliance add-on
                               ┌──────────────────┐
                               │    AgriX GGAP    │  (Enterprise dodatak; kupuje hladnjača za mrežu kooperanata)
                               └──────────────────┘
```

> **Napomena o imenovanju (reconciliacija):** stariji `02/02A/07` treći stub zovu **GGAP**; najnoviji `09`/`09A` (24.07) tri _prodajna_ stuba definišu kao **Enterprise / Gazdinstvo / Savetnik**, a **GGAP je demotovan na Enterprise compliance add-on** (QA #75, §8). Važi novija formulacija.

### 2.1 Cenovni model (izvučen iz QA log-a)

| Pravilo | Odluka | QA ref |
|---|---|---|
| Osnovna jedinica | Godišnja pretplata **po pravnom licu** (nema mesečno / per-user / per-uređaj) | #25/26/29/31 |
| Stanice | Osnovni paket do **5 stanica**; svaka >5 = ista fiksna godišnja (nema tier-ova) | #121/122 |
| Cena stanice | Ista u Desktop i Mobile; razlika je u ceni osnovnog paketa | #123 |
| Mobile multiplikator | Desktop Otkup + Mobile ≥ **2× Desktop Otkup** | #126/127 |
| Moduli (SEF/Banka/Dispatch) | Fiksno **po pravnom licu**, isto bez obzira na paket, plaća se jednom za sve instance | #124/125/135 |
| Proizvodni dodatak | Fiksno **tek posle standardizacije**; 1 pogon = 1 dodatak; dodatni pogon = dodatna Desktop instanca | #130/131/132/243/244 |
| Gazdinstvo | Samo godišnja; Basic 19€ / Pro 39€; prvih 50 Partner uz paket hladnjače | #159/161/172 |
| Savetnik | Po aktivnom gazdinstvu; **javna cena se ne objavljuje** | #198/225 |
| GGAP | Fiksno po pravnom licu, pokriva sve kooperante; nema per-user | §8 |
| Custom rad | **Time-and-materials**, jedna satnica, procena + max budžet, prekoračenje uz pisano odobrenje | #107–110 |
| Javne cene | Desktop/Mobile **rasponi**, proizvodnja **raspon**, stanica >5 **tačan iznos**, Gazdinstvo **tačno**; Savetnik ne | #156–159/225 |

### 2.2 Core vs modul

- **Core (u Desktop-u):** otkup + dokumenta, prijemnice, fakture (kreiranje/evidencija), ambalaža + repromaterijal, **agrohemija (kompletan tok)**, skladište/WMS/palete/sledljivost, standardni izveštaji, **ručni unos novca**, kartice/salda, Management PWA (QA #3/12/14/16/17).
- **Odvojeni moduli:** SEF, Banka (automatizacija), Dispatch (napredni), Hladnjača/Proizvodnja.
- **Granica ka ERP-u:** AgriX **nije** računovodstveni ERP — bez glavne knjige, PDV-a, završnog računa, zarada; **koegzistira** sa BizniSoft/Pantheon (QA #72–74).

---

## 3. ⚠️ Code-verifikacija: registar je delom ZASTAREO

**Ključni nalaz v2.** Registar (`KNOWN_ISSUES §8` / `PLAN_SANACIJE`) sidren je na FM v35 (`a0bc9e2`), a `main` je odmakao na **v2.24.0**. Provera **stvarnog `src-vba` koda** pokazuje da su neki „otvoreni P0/P1" **već ispravljeni**:

| Stavka | Registar kaže | Kod (v2.24.0) kaže | Presuda |
|---|---|---|---|
| **AUD-030** SEF 409→REJECTED | „jedini nov aktivan **P0**" | `modSEFClient.bas:491-501` ima `Case 409 → CONFLICT` **sa eksplicitnim AUD-030 fix komentarom**; retry reuse-uje isti requestId (idempotentno) | ❌ **već ispravljeno** — SEF nema više P0 |
| **AUD-033** auth lanac do Admin panela | P1/„P0-klasa bezbednosti" | `MozeAdministraciju()` čuva `BuildAdminPanel`, **svaku** admin akciju, `ShowConfigSheet` (`modAdmin.bas:45-48,211-214`, `modPodesavanja.bas:740-743`) | ❌ **već implementirano** |
| **AUD-034** btnBanka pre auth | P1 | auth provera **prethodi** booking importu (`frmOtkupAPP.frm:729-734` pre `:751`) | ❌ **već ispravljeno** |
| **AUD-014** Banka knjiži na otvaranje forme | P1 | `frmBankaImport.frm:72` `UserForm_Activate` → `AutoMapStrongKeysBankaImport_TX` pod `On Error Resume Next`, tiho commit-uje | ✅ **potvrđeno — živ, opasan** |
| **AUD-001** Sheets JSON read | P0 | `modGoogleSheets.bas:1757-1830` global `Replace(", ",",")` unutar navodnika, bez `\"`/`\uXXXX`; **živ OTK/VOZ import** | ✅ **potvrđeno — živ P0** |
| **AUD-041/042** MasterSync wrong-write/dup# | P1 | `TryUpdateVozacID` vraća True na tihi `UpdateCell` fail; `GenerateBrojZbirne`=rowcount+1 (dup posle gap/storno); loš datum→danas | ✅ **potvrđeno — sva 3 živa** |
| **KI-006** ART_POCETNI_DUG PWA fantom | OPEN | `GetMagacinStanje` ga izuzima (`modAgrohemija.bas:435`), `ExportMagacinKoop` **ne** (`modStammdatenSync.bas:1966-1982`) → curi kao `MAG_IZLAZ Kolicina=1` | ✅ **potvrđeno** |

**Dve reinterpretacije koje menjaju plan:**
1. **AUD-033/034 nisu „nedostajući guard" nego config-default.** `MozeAdministraciju()` je namerno **anti-lockout** — vraća `True` kad je `AUTH_ENABLED` isključen. Prava izloženost je **„AUTH se isporučuje ugašen"** (na instalaciji bez uključenog auth-a, svako je admin). Vrednost RF-23 se pomera sa „dodaj guard" na **„odluči default uključenog auth-a + signaliziraj plaintext-PIN fallback"**.
2. **AUD-030 (SEF P0) je zatvoren u kodu** → SEF modul je **bliži prodaji** nego što registar tvrdi; ostaje P1 SEF korektnost/UX klaster (AUD-031/032, nije nezavisno re-verifikovan u kodu).

> **Posledica:** stvarno-otvoreni P0-klasa blokeri su **manje brojni i koncentrisani na PWA→centrala sync kičmu** (JSON + MasterSync) **+ Banka forma-na-otvaranju + storno/finansije lanci + AUTH-default**, a **ne** na SEF/auth kako registar sugeriše. Uz to, **RF-03/RF-04 (storno) se MORAJU re-verifikovati protiv v2.24.0** — storno PR-ovi #134-137 su prepisali `modStornoFlow` (+746 linija), pa su delovi AUD-020/AUD-005 možda već rešeni. Neto: realan open-P0 broj je verovatno **niži od 9**.

---

## 4. Portfolio proizvoda — prodajna tabela

| # | Prodajna celina | Sadržaj | Cena (model) | Impl | Dokaz | Launch gate (gde je posao) |
|---|---|---|---|---|---|---|
| 1 | **Platform Services** | auth, sync/MasterSync, storno/audit, monitoring, self-update, backup, gates | ugrađeno | Implemented | Production-proven (malo klijenata) | Runtime gate-ovi neizvršeni (KI-002/003); self-update component-loss (RF-24); TX rollback (RF-13) |
| 2 | **AgriX Desktop (Core)** | otkup→dokument→prijem→faktura→izveštaj + Mgmt PWA | godišnja / pravno lice + stanice >5 | Implemented | Production-proven (3 klijenta) | **Najveća koncentracija:** P0 data-safety + finansije/storno/izveštaji P1 + health-check green |
| 3 | **AgriX Mobile** | Desktop + PWA Otkupac + Vozač (osnovni) | ≥ 2× Desktop Otkup | Implemented | Limited production | **Strateški najvredniji:** JSON (RF-14) + MasterSync (RF-28) — glavna diferencijacija |
| 4 | **SEF modul** | slanje/status/storno izlaznih, preuzimanje ulaznih, povezivanje | godišnja / pravno lice | Implemented | Limited evidence | **P0 zatvoren (kod).** Ostaje P1 korektnost/UX (RF-21/22): stornirana faktura sendable, truncation, „poslata" za grešku |
| 5 | **Banka modul** | uvoz izvoda, povezivanje, rasknjižavanje, avansi, nalozi | godišnja / pravno lice | Implemented | Limited evidence | **Knjiži novac na otvaranje forme** (AUD-014, potvrđeno) + dedupe/crash (AUD-025) + over-order (AUD-026) → RF-09/10 |
| 6 | **Dispatch (napredni)** | raspoređivanje vozila/vozača, rute, kapaciteti, dispečer | godišnja / pravno lice | Implemented | Limited evidence | Osnovni Vozač = Mobile-Core; napredni dispatch = malo dokaza, discovery |
| 7 | **Hladnjača/Proizvodnja** | prerada, palete sveže/prerađene, lager, sledljivost; (2027: nalozi, norme, linije) | godišnja **posle standardizacije** | Impl/Partial (palete u prod.) | Limited evidence | Palete correctness (RF-12); pun proizvodni sistem = **najveći NOV build 2027**, nije launch-blocker sad |
| 8 | **AgriX Gazdinstvo** | kartica, parcele/GIS, tretmani/karenca, troškovi, agrohemija, prognoza | Basic 19€ / Pro 39€ | Impl/Partial (~4.850 l koda) | **Pilot** | **Nije kod — validacija:** activation, 30/90/180 retention, WTP (Track G) |
| 9 | **AgriX Savetnik** | 1 agronom → više gazdinstava | po aktivnom gazdinstvu | Planned (cilj 2027) | — | Izgraditi osnovnu verziju; nije launch-kritično 2026 |
| 10 | **AgriX GGAP** | evidencije/dokazi/rokovi/audit-readiness | godišnja / pravno lice | Planned/Discovery | Unvalidated | **Not for sale**; domain owner + pilot; post-2027 |

**Hardver (zaseban readiness):** Kiosk (`Controlled rollout`), Tablet (`Optional hardware`), Termalna štampa (`Controlled rollout`). Nezrelost **ne obara** status PWA aplikacije.

---

## 5. Osa A — Inženjerska spremnost (correctness / release truth)

### 5.1 P0 tabela — 9 nominalnih (kalibrisano za single-writer desktop)

| ID | Defekt | Scenario otkaza (podatak → pogrešan ishod) | Blokira | Napor | RF |
|---|---|---|---|---|---|
| **AUD-001** | Sheets JSON read kvari vrednosti sa `", "`/`\"`; `\uXXXX` nedekodiran | PWA otkup gde polje sadrži `"Petrović, Mile"` → `SplitCsvJson` deli usred vrednosti → kolone se pomere → pogrešan kooperant/iznos u `tblOtkup` na živom OTK/VOZ putu | PWA-sync Mobile | **S** (RF-14 M ukupno) | RF-14 |
| **AUD-002** | Ceo import batch u jednoj Excel TX, ali per-sheet Google writeback `Synced>Master` nije rollback-abilan | Red 5 padne → Excel TX vrati sve, ali redovi 1–4 već označeni `Synced>Master` na Google strani (nepovratno) → potvrđeni redovi **trajno izgubljeni** | PWA-sync Mobile | **M** | RF-14 |
| **AUD-003** | Pozicijski `Array(...)`+`AppendRow` na 17-kol `tblNovac` bez order-guard (`SaveNovac`=P0) | Instalacija sa schema drift (umetnuta/pomerena kolona) → `SaveNovac` upiše `Iznos` u pogrešnu kolonu → **tiho korumpiran novčani ledger**; svi saldi/izveštaji čitaju korupt | Desktop Core / Finance | **S** | RF-02 |
| **AUD-004** | `RollbackTx` bez EH i bez garantovanog `CleanUp`; nema `Class_Terminate` | Bilo koja `_TX` op vrati; `RestoreTable` pukne usred petlje → `CleanUp` preskočen → `EnableEvents`/manual-calc ostaju off, tabele **polu-vraćene, Excel zamrznut** | Platform (323 BeginTx sajta) | **S** | RF-13 |
| **AUD-005** | AutoChainHladnjača: `otpID`/`SaveZbirna_TX` rezultati odbačeni; `outBrPrij` pre uspešnog kreiranja | `SavePrijemnica_TX` padne ali `outBrPrij` već setovan → recovery relink na **nepostojeći broj**; lanac javi „COMPLETED" a slomljen | Palete-Hladnjača | **S** | RF-04 (re-verify) |
| **AUD-006** | Journal recovery: današnje linije poređene sa **all-time** brojem redova; `UpdateCell` mutacije nikad ne journal-uju | Na zreloj svesci (hiljade redova) recovery **ne može da se okine**; crash usred `UpdateCell` je nepovratan | Platform | **M** | RF-13 |
| **AUD-007** | `TryParseDateValue` prima nemoguće datume kroz `DateSerial` rollover | Izvod sa `45.13.2025` → rollover u validan datum → bankarska transakcija knjižena sa **pogrešnim datumom** → pogrešan period u saldu | Banka | **S** | RF-13 |
| **AUD-020** | Storno „lažni uspeh" lanac (5 grana; paletni detach false-success; invariant `0=0` za nepostojeću zbirnu) | Storno korekcije: `Detach` vrati 0 na grešci ali `ok=True` → korekcija „COMPLETED" a palete i dalje vezane za otkazani dok | Desktop Core / Storno | **M** | RF-03 (re-verify) |
| **AUD-030** | ~~SEF 409→REJECTED~~ | **VEĆ ISPRAVLJENO u kodu** (§3) — `Case 409→CONFLICT` | ~~SEF~~ | — | RF-21 (samo P1 ostaje) |

> Realno-otvoreni P0: **AUD-001/002/003/004/006/007** čvrsto; **AUD-005/020** uz re-verifikaciju protiv v2.24.0 (možda delom rešeni); **AUD-030 zatvoren**. Koncentracija: **PWA-sync (RF-14) + Platform TX/journal (RF-13) + Finance-pozicijski (RF-02)**.

### 5.2 P1 klasteri po oblasti (sažeto; puni detalj gore/u `KNOWN_ISSUES §8`)

- **Finance/Novac (RF-02/05/08):** uplata na storniranu fakturu (AUD-009); avans na tuđi/stornirani cilj + no-op `True` naduvava brojače (AUD-010); novac storno prima broj umesto `NovacID` (AUD-008); `CreateFaktura` bez kupac-ownership + `Count=1` (AUD-011); **storno plaćanja skriva dug** iz payout liste (AUD-021).
- **MasterSync (RF-28, deli fajl sa RF-14):** dup `BrojZbirne` (AUD-041); `Synced>Master` sa praznim VozacID + loš datum→danas (AUD-042); otpremnica meša vrste/cene (AUD-043); station-mirror FK bez `tblVozaci` reda (AUD-046).
- **SEF (RF-21/22):** stornirana faktura potpuno sendable; qty/price truncation → nekonzistentan UBL; double-submit; „Faktura poslata" za REJECTED/TECH_FAILED (AUD-031/032). *(nije nezavisno code-verifikovano — po registru)*
- **Banka (RF-09/10):** **AUD-014 knjiži na otvaranje** (potvrđeno); dedupe bez broja računa + crash na 3+ kandidata (AUD-025); CSV over-order (AUD-026).
- **Izveštaji (RF-06/07/08):** kartice kreću od nule (AUD-023); freshness — stari podaci pod novim periodom + tihe lazy greške (AUD-024); revers bez storno filtera + meša ambalaže (AUD-012); reprint storniranog + `.UnMerge` (AUD-027); sledljivost parcijalni trace kao kompletan (AUD-045).
- **Otkup UI (RF-11/26):** parcela false-alarm; nevezan blok nevidljiv iz „Izgubljeni"; stornirani blok daje default cenu (AUD-028); **cenovnik stale cena** — miss ostavlja cenu prethodnog artikla (AUD-036).
- **Palete (RF-12):** UI traži i kutije i kese dok core dozvoljava jedno; sledljivost štampa količine cele zbirne; Reassign/Detach diraju prerađene palete (AUD-029).
- **Agrohemija (RF-27 ✅):** cena≠knjižena — **rešeno** (grana `claude/rf-27-agrohemija-cena`).
- **Auth/Startup (RF-23):** AUD-033/034 — **već implementirano**; ostaje **AUTH-default + plaintext-PIN signal** (§3).
- **Self-update/Infra (RF-24/25/29/30):** component-loss pri prekidu (AUD-035); integritet/health false-green (AUD-044/047); plaintext PWA PIN + JMBG u Sheets (AUD-019, RF-20); empty-source → cloud wipe (AUD-038).

**Dve nepakovane P1** (nemaju RF): **AUD-013** (`MatchesFilter` nepoznat operator → tiho ukloni kriterijum, `modArrayUtils.bas:94-95`) i **AUD-015** (dashboard broji stornirane, neslaganje sa frmDokumenta). Treba im dom ili eksplicitno odlaganje.

### 5.3 Napor i sekvenca

**Tally (30 paketa):** ~10× S/S–M, ~13× M, 2× M/L–L (RF-06, RF-20), 5× P2 konsolidacije. **Ukupno ≈ 28 dev-sesija / 13 release milestone-a ≈ 10–14 nedelja** pri 2–3 paketa/nedelji. **Kritičan put M0–M4 ≈ 4–5 nedelja** nosi svih 9 P0 + top P1 lance (auth, SEF, finansije, storno, sync) — ~40% napora, gotovo sva bezbednosna vrednost. **Usko grlo je operater smoke-turnaround, ne kod.**

**Redosled (playbook §4):** `RF-01 → RF-02 → RF-23 → RF-21 → RF-27✅ → RF-03* → RF-04* → RF-05 → RF-14+RF-28 → RF-06 → RF-07 → RF-08 → RF-22 → RF-09 → RF-10 → RF-11 → RF-12 → RF-13 → RF-26 → RF-29 → RF-30 → RF-24 → RF-25 → RF-15+ → RF-20`

**Zašto front-load:** RF-27 (najjeftiniji high-value, urađen); RF-23+RF-21 pre storna (bezbednost + jedini nominalni P0 — iako §3 pokazuje da su delom zatvoreni, paket i dalje nosi AUTH-default + SEF P1); RF-02 prvi jer RF-10 zavisi od njegovog „applied amount".

**Deljeni fajlovi (moraju se serijalizovati):** `modMasterSync.bas` = **RF-14 + RF-28** (raditi zajedno, M4); `modNovac.bas` = RF-02/03/05/06; `frmOtkup.frm` = RF-11 + RF-26; `frmDokumenta.frm` = RF-05 + RF-26; `modStorno.bas` = RF-03 + RF-05. **Nezavisno isporučivi:** RF-01, RF-12, RF-29, RF-30, RF-24, RF-25, RF-15…19.

**Re-verifikacija (kritično):** RF-03/RF-04 protiv v2.24.0 — storno PR-ovi #134-137 (+746 l) su možda već zatvorili delove AUD-020/005. Zadrži potvrđene, izbaci rešene, pre pisanja koda.

### 5.4 Release-truth gate-ovi (KI-002/003 — OPEN, obavezni pre „Standard offer")

1. **Compile gate** (RELEASE_GATES §2.1) — svi dirani moduli kompajliraju; `RequireColumns/RequireUpdateCell` vidljivi; nema stale form-handler referenci.
2. **Smoke — BusinessFlowPro** (§2.2) — E2E `Otkup→Otpremnica→Zbirna→Prijemnica→Faktura→SEF`; dvoklasni dokumenti atomski; duplikat-faktura prevencija; stornirano izuzeto.
3. **Route health** (§3.1) — `runGasRouteHealthCheck()`; svaka aktivna akcija ima handler; disabled → `FEATURE_DISABLED`.
4. **PWA role smoke** (§4.2–4.5) — Otkupac (offline→pending→synced, badge, double-tap lock), Kooperant (dedupe, stale syncing→pending), Vozač (`brojZbirne` pre sync, storno ne reciklira sekvencu), Management (no-sync-for-role).
5. **Monitoring** (§7), **SEF gate** (§6 — hvata AUD-030/031/032 regresije), **`RunProductionHealthCheck`** (§2.10 — 0 failure-a; fixtures rollback-ovani).

**Upozorenje:** release gating **ne sme** da se oslanja na `modE2EReleaseGate` dok RF-26 ne popravi result-contract (AUD-039 false-green); dotad manualni smoke suite-ovi + `RunProductionHealthCheck`. Otvoreno van RF: **KI-001** (`saveParcelPolygon` auth — proveriti deploy `Code.gs`), **KI-006** (fantom u `ExportMagacinKoop` — potvrđeno).

---

## 6. Osa B — Komercijalna spremnost

### 6.1 Tržište (`04_MARKET`)
- **Registry TAM = 1.516 aktivnih firmi** (ne tvrditi javno „1.516 kupaca"). **SAM ≈ 150–300** (hipoteza, low-med). **SOM:** 12–18 mes. bazno **8–12**; 3–4 god. bazno **40–60** (200 je stretch — **finansijski model ne sme koristiti 200 kao bazu**).
- **Kiosk sizing:** ~10 stanica/firmi → 10 firmi = **100 terminala**, 50 = 500, 100 = 1.000 (nabavka samo protiv potpisanih firmi).
- **Gazdinstvo B2B2C:** ~100 kooperanata/firmi, **5% konverzija (hipoteza)** → 10 firmi ≈ 50 plaćenih.
- **Sezonalnost (FACT):** otkup je sezonski intenzivan; **bez onboardinga u špicu**; in-season veliki release = nesrazmeran reputacioni rizik. Svaki deal nosi kulturu + početak sezone + **poslednji bezbedan datum onboardinga**.

### 6.2 ICP / Anti-ICP (`03`)
- **Sweet spot: 6–15 stanica** (hipoteza CUS-H01). ICP scoring 100 poena — **process breadth koji AgriX zatvara = najveći ponder (20)**. Tiers: A 75–100 (pun demo), B 60–74 (prvo blokeri), C 45–59 (samo uz referencu), D <45 (ne prodavati aktivno).
- **Process fit > prihod.** Anti-ICP: traži trajni fork; neograničen custom u licenci; nema data/impl ownera; pun rollout pred sezonu bez pilota; odbija backup/monitoring/kontrole; traži garanciju rezultata/GGAP sertifikata; kupuje isključivo po najnižoj ceni. **Jedan diskvalifikator obara skor.**

### 6.3 Konkurencija + Infosys wedge (`05/05A/05B`)
- Hijerarhija: **Infosys #1 → SOFTEK+KRUNET (24 relevantne reference) → Yuteam (Vojvodina) → generički ERP / Excel / status quo**. Tržište **nije greenfield**.
- **Najjači dokaz u celom planu (FACT): 2 postojeća klijenta prešla sa Infosys-a** (oba ~150M RSD). → „praktično iskustvo sa migracijom" je dozvoljena tvrdnja; opšta superiornost tek posle win-intervjua.
- **Kvalifikovana lista od 9 Infosys-replacement naloga** (`05B`): Talas 1 — **FRIGO-PAUN** (>1.000 proizvođača), **BUDIM GRAD** (1,4 mlrd), **FRUCOM FOOD**, **AS-AGRO 99** (3,5 mlrd), **AGRO-SUNCOKRET**, **FRIGO BRAĆA MITROVIĆ**; Talas 2 — MAGIC BERRY, FRIGOMIL, MALINA PROIZVOD. **Ne masovni outreach na svih 30.**
- Generički ERP (BizniSoft/Pantheon/Minimax): **koegzistirati, ne napadati** — AgriX pokriva ulazni tok robe koji oni ne. Izgubljen lead (~400M) otišao lokalnom programeru **zbog vremena odziva** → speed-to-contact je konkurentska funkcija.
- **Teško kopirati:** povezan Enterprise–Gazdinstvo–GGAP data flywheel, real-season know-how, deljeni proizvod bez forkova, mrežni efekat. **Lako kopirati:** jedan ekran/izveštaj/PWA forma → **ne pozicionirati se na jednu funkciju.**

### 6.4 Pozicioniranje (`06`)
- Centralna poruka: *„Otkupci i vozači unose podatke tamo gde posao nastaje, a AgriX automatski puni centralnu bazu za kontrolu, fakture i izveštaje."* Operater: od **prepisivača ka kontroloru**. **Nije „ERP za poljoprivredu"** — koegzistira.
- **Dozvoljeno** (uz tačan scope): „podatak se ne prepisuje ako je već unet na terenu"; „radi uz postojeći ERP"; **„tri firme koriste AgriX, dva klijenta prešla sa Infosys-a"**. **Zabranjeno** (bez dokaza): „eliminiše sve greške", „nikad ne gubi podatke", „radi bez interneta u svakoj funkciji", „podržava svaki printer", „potpuna zamena ERP-a", „potpun GGAP compliance".
- **Demo mora pratiti realan tok** stanica→sync→centrala→faktura→izveštaj (ne spisak ekrana). **Enterprise nema pilot** — samo dummy demo (QA #255–257).

### 6.5 Cene — što nedostaje
- **Konkretni brojevi postoje samo za Gazdinstvo** (19€/39€). Sve ostalo su _pravila_ (§2.1); `10_PRICING` (apsolutni iznosi) **nije napisan** → **#1 komercijalni blocker**.
- WTP nepotvrđen: dispečer+finansijske automatizacije podižu WTP više od izveštaja (CUS-H02); 15–20 struktuiranih WTP intervjua po segmentu (`04` §21).

### 6.6 Unit-economics inputi (FACT, ali model nenapisan)
3 klijenta · ~10 stanica · ~100 kooperanata · 1 desktop korisnik · 5% konverzija (hipoteza) · 19/39€ · onboarding ~1 dan → cilj 0,5 · **~1 support poziv/nedeljno** (~15 min trivijalno / sati kad je bug) · SLA: RTO 24h, kritičan odziv 1h · tim: osnivač (+možda 1 dev) + 1 support/impl. **`11_UNIT_ECONOMICS` i `12_FINANCIAL_MODEL` nisu napisani** → ARR/stanici, gross margin, CAC, LTV nedefinisani.

### 6.7 Komercijalni gapovi (rang)
1. **Apsolutni cenovnik** (`10`) — bez ovoga nema prave ponude. *Genuinely missing.*
2. **Unit economics + finansijski model** (`11`/`12`). *Missing.*
3. **Standardizovan onboarding checklist** (Track C) — danas founder-only ~1 dan; + free→paid cutover neodlučen. *Namera postoji, artefakt (checklista) ne.*
4. **Ugovori/legal/SLA** (`20`) — pravila razbacana u QA #40–57/181–189; formalni okvir ne.
5. **Sales playbook + GTM + partneri** (`14/15`) — sirovi inputi (ICP scoring, Infosys lista) postoje; proces ne.

> **Reconciliacija:** README status-tabela („sve nije započeto") je **zastarela** — 02–08 su napisani (CHANGELOG). Rast: **40–60 realno (3–4 god.)**, ne 200.

---

## 7. Launch-readiness po proizvodu (presuda)

| Proizvod | Presuda | Najveći posao |
|---|---|---|
| **AgriX Desktop (Core)** | Prodaje se postojećima; **ne još „Standard offer"** dok P0 + finansije/storno/izveštaji P1 + health-gate ne prođu | Osa A 5.1 + 5.2 (finance/storno/izveštaji) |
| **AgriX Mobile** | Strateški core; **najvredniji inženjerski posao** = JSON+MasterSync (RF-14+RF-28) | Osa A: PWA-sync kičma |
| **SEF modul** | **P0 zatvoren u kodu** → bliži prodaji nego što registar tvrdi; ostaje P1 (RF-21/22) pre „Standard" | RF-21/22 (korektnost/UX) |
| **Banka modul** | ⛔ Ne „Standard" — **knjiži novac na otvaranje forme** (potvrđeno) | RF-09/10 |
| **Dispatch (napredni)** | Optional; osnovni Vozač ide u Mobile | — |
| **Hladnjača/Proizvodnja** | Palete se prodaju (Controlled); pun sistem = 2027 build | RF-12 + 2027 roadmap |
| **Gazdinstvo** | Pilot only; **validacija pre skaliranja** (retention/WTP), ne kod | Track G metrike |
| **Savetnik** | Osnovna verzija cilj 2027 | — |
| **GGAP** | Not for sale; discovery; domain owner | Post-2027 |

---

## 8. Odluke

1. **Ground-truth (kod) ima prednost nad registrom** kad se razilaze — registar je delom zastareo vs v2.24.0. AUD-030/033/034 su **zatvoreni**; ne prikazivati ih kao otvorene blokere.
2. **P0 data-safety (Osa A 5.1) ima apsolutni prioritet** — blokira Core i module.
3. **PWA-sync (RF-14+RF-28) je strateški najvredniji posao** — nosi glavnu diferencijaciju; raditi kao jednu sesiju (isti fajl).
4. **Banka se ne sme prodavati kao „Standard"** dok AUD-014 (knjiži na otvaranje) ne padne. **SEF** je bliži — ostaje P1 klaster (RF-21/22).
5. **RF-03/RF-04 re-verifikovati protiv v2.24.0** pre koda (storno rewrite #134-137).
6. **Runtime gate-ovi + `RunProductionHealthCheck` = obavezan exit-uslov** (KI-002/003).
7. **#1 komercijalni preduslov = apsolutni cenovnik** (`10`); **#2 = standardizovan onboarding**.
8. **Infosys-replacement (9 naloga) je primarni GTM wedge**; obaviti 2 win-intervjua sa migriranim klijentima → battlecard + migration playbook.
9. **Finansijski model koristi 40–60 (3–4 god.) kao bazu**, ne 200.

---

## 9. KPI i pragovi

**Inženjering:** 0 izgubljenih potvrđenih zapisa; 0 nekontrolisanih canonical duplikata; ≥99% sync uspeha bez dev intervencije; otvoreni RF paketi (29/30); ciljni workbook 0 health failure-a; kritičan put M0–M4 zatvoren.
**Komercijala:** fiksiran javni raspon Desktop/Mobile + tačan iznos stanice >5; ≥1 onboarding po checklisti (izmereno vreme/firmi); 2 Infosys win-intervjua; support cost održiv za ~8 firmi.

---

## 10. Rizici

| Rizik | Uticaj | Mitigacija |
|---|---|---|
| „Zeleno" bez izvršenih runtime gate-ova | Latentan P0 u produkciji | KI-002/003 hard exit-gate; health-check na ciljnom workbook-u |
| Banka knjiži na otvaranje forme (AUD-014) | Pogrešno/tiho knjiženje novca | RF-09 pre Banka prodaje; dotad modul van „Standard" |
| AUTH se isporučuje ugašen (§3) | Svako je admin na instalaciji bez auth-a | RF-23: odluka o default-u + signal plaintext-PIN fallback-a |
| Registar zastareo → duplo raditi već-rešeno | Straćen kapacitet | Re-verifikacija RF-03/04 (i drugih) protiv v2.24.0 pre koda |
| Prespora prodaja/standardizacija (QA #93/94/98) | Uska grla 30–50 klijenata | Osa B #1 (cenovnik) + #3 (onboarding) pre skaliranja |
| Sporost odziva na lead (izgubljen ~400M lead) | Gubitak posla ka lokalnom programeru | SLA za inbound: instant ack + ljudski kontakt |

---

## 11. Otvorena pitanja (QA §9)

1. Apsolutni cenovnici (Enterprise, Mobile, stanice, moduli).
2. Trenutak free → paid onboarding za nove Enterprise klijente.
3. Pravna pravila za povlačenje saglasnosti proizvođača za deljenje podataka.
4. Formalni uslovi partnerske provizije i atribucije lead-a.
5. Redosled post-2027: GGAP, Savetnik marketplace, multi-Enterprise.

---

## 12. Akcije

| # | Akcija | Vlasnik | Rok / gate |
|---|---|---|---|
| 1 | **Re-verifikovati registar protiv v2.24.0** (RF-03/04 + potvrditi §3 zatvorene stavke) pre koda | osnivač | pre M2 |
| 2 | RF sekvenca M0–M4 (kritičan put: RF-01/02/23/21 → RF-14+28) | osnivač (+dev) | ~4–5 nedelja |
| 3 | Runtime gate-ovi + `RunProductionHealthCheck` na ciljnom workbook-u | osnivač | pre svakog „Standard offer" |
| 4 | **Banka van „Standard"** dok AUD-014 ne padne; **SEF** označiti „P0 zatvoren, P1 u toku" | osnivač | odmah |
| 5 | `10_PRICING_AND_PACKAGING.md` — apsolutni iznosi | osnivač | pre sezone |
| 6 | Standardizovan onboarding checklist (Track C) | osnivač | 30.11.2026 |
| 7 | 2 Infosys win-intervjua → battlecard + migration playbook; Talas 1 outreach (9 naloga) | osnivač | Q3 2026 |
| 8 | Dodeliti dom nepakovanim AUD-013/AUD-015 ili ih eksplicitno odložiti | osnivač | pre M3 |

---

## 13. Izvori i datum provere

- Komercijala: `09_QA_DECISION_LOG`, `07/07A/07B`, `03`, `04`, `05/05A/05B`, `06`, `README`, `CHANGELOG` — 2026-07-24.
- Inženjering: `../PLAN_SANACIJE`, `../REFAKTOR_PLAYBOOK` (RF status), `../KNOWN_ISSUES §8`, `../ROADMAP §10`, `../RELEASE_GATES` — 2026-07-24.
- **Code-verifikacija (v2):** `modSEFClient.bas:491-501`, `frmBankaImport.frm:72`, `modGoogleSheets.bas:1757-1830`, `modMasterSync.bas:1773-1798/2887-2928`, `modMaticniLookups.bas:256-263`, `modAdmin.bas:45-48/211-214`, `modAgrohemija.bas:435`, `modStammdatenSync.bas:1966-1982` — verifikovano protiv `origin/main` v2.24.0.

## 14. Change log

| Datum | Izmena |
|---|---|
| 2026-07-24 | Inicijalna verzija: razrada QA log-a u portfolio + launch-readiness (dve ose). |
| 2026-07-24 (v2) | Produbljeno: (1) code-verifikacija — registar delom zastareo (AUD-030/033/034 zatvoreni u v2.24.0); (2) per-defekt P0/P1 sa scenarijima i naporom (~28 dev-sesija, kritičan put 4–5 ned.); (3) tržište/ICP/konkurencija/Infosys wedge sa realnim brojevima; (4) release-truth gate-ovi; (5) reconciliacije (README stale, rast 40–60, Savetnik vs GGAP). |
