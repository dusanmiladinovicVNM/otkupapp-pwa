# PLAN SANACIJE — AgriX / OtkupApp (konsolidovani izvršni program)

**Datum:** 2026-07-20
**Verzija plana:** 1.0
**Sidro koda:** `origin/main` v2.24.0 (`9fd7087`)

> Ovo je **program-nivo** plan: milestone-i, gejtovi, sekvenca, procena obima i rizik.
> **Detalji po paketu** (scope, fajlovi, regression) žive u `REFAKTOR_PLAYBOOK.md` (RF-01…RF-30).
> **Nalazi po stavci** su u `AUDIT_FM_TRIJAZA.md` (DEO I/II/III); **actionable registar** u
> `KNOWN_ISSUES.md` §8 (AUD-001…048). Ovaj dokument NE duplira te izvore — vezuje ih u izvršni plan.

---

## 1. Šta se sanira (potvrđeni inventar)

Triaža svih FM verzija (v35 → v85 → v142) + nezavisni review, kalibrisano za **single-writer
desktop** model. Ukupno **48 dedupliranih AUD stavki** → **30 RF paketa**.

| Klasa | Broj | AUD stavke |
|---|---|---|
| **P0** (bezbednost podataka) | 9 | AUD-001, 002, 003, 004, 005, 006, 007, 020, 030 |
| **P1** (funkcionalni defekti) | 34 | AUD-008…019, 021…029, 031…036, 040…046 |
| **P2** (hardening / dug) | 5 | AUD-037, 038, 039, 047, 048 |
| Konsolidacije (tehnički dug) | — | RF-15…19 (deljeni helperi, klonovi) |

**Realno stanje (pošteno):** dokument je činjenično vrlo precizan; nula sistemskih opovrgnutih
nalaza. Ali je **težina sistematski precenjena** — od stotina „P0/P1" naslova, na single-writer
modelu ostaje **9 stvarnih P0** i **~34 P1 klastera**. Jedini nov **aktivan P0** iz cele delte je
**AUD-030** (SEF 409 → REJECTED). Ostali „P0" su ili već pokriveni, ili SWMR-uslovljeni (Prihvaćeno).

---

## 2. Načela izvršenja (obavezna za svaki paket)

1. **Serijski, jedan paket po sesiji.** Bez paralelnih grana nad istim fajlom.
2. **Re-baza na svež `main` pre SVAKOG paketa.** `fetch` → `git merge-tree` provera → rebasе →
   pokaži rezultat → `push --force-with-lease` tek po odobrenju (CLAUDE.md §6, „Opcija 3").
3. **Minimal-delta.** `reuse > new`, `extend > duplicate`. Nema idealizovanog redizajna.
4. **Schema-first.** Pre svakog upisa `DebugKoloneTabele` — realne kolone se razlikuju po
   instalaciji. Upis **po imenu** (`GetColumnIndex`/`UpdateCell`/`RequireUpdateCell`), ne pozicijski.
5. **ASCII disciplina.** VBA izvori ostaju 100% ASCII; dijakritika samo kroz `modPoruke.UpsertPoruke`
   + `ChrW`. `.frx` se NE dira kao tekst; nove kontrole runtime (`Controls.Add` + WithEvents).
6. **Svaki paket = jedan Excel release-kandidat.** `ImportAllVBA` → `Debug → Compile` → smoke →
   ship. Paket nije „gotov" dok operater ne odradi smoke checklistu.
7. **Re-verifikacija pre storna.** RF-03/RF-04 ciljaju v35 linije; `main` je otišao na v2.24.0
   (storno PR #134–137, `modStornoFlow` +746). **Obavezno re-verifikovati protiv `origin/main`** —
   deo nalaza je možda već rešen; potvrđene zadržati, rešene skinuti.

---

## 3. Milestone-i (release vozovi)

30 paketa grupisano u **13 shippable milestone-a**. Redosled prati `REFAKTOR_PLAYBOOK.md` §4
(P0-klasa bezbednosti i jedini nov P0 podignuti napred). Svaki milestone se zatvara Excel
release-om + Fleet publish-om.

| M | Naziv | Paketi | Zatvara (AUD) | Obim | Zašto tu |
|---|---|---|---|---|---|
| **M0** | Higijena | RF-01 | (deo AUD-016) | S | Balast/mrtvi moduli — smanjuje import šum, čist teren. |
| **M1** | Pristup i startup | RF-02, RF-23 | 003(deo), 033, 034 | S+M | **P0-klasa bezbednosti:** korisnik „Matični" stiže do Admin panela; `AccessWasDenied` se ne poziva. |
| **M2** | SEF ispravnost + agrohemija cena | RF-21, RF-27 | 030, 031, 040 | M+S | **Jedini nov aktivan P0 (409).** Agrohemija cena≠knjižena (jeftin P1, high-value). |
| **M3** | Storno (re-verify vs v2.24.0) | RF-03\*, RF-04\*, RF-05 | 020, 005, 008–010 | M+M+M | Storno „lažni uspeh" lanac; hladnjača; frmDokumenta. **Prvo re-verifikacija.** |
| **M4** | Sync jezgro | RF-14 + RF-28 | 001, 002, 018, 041, 042, 043, 046 | M+M (1 sesija) | JSON read korupcija + ZBR duplikati + writeback + otpremnica + stanica-mirror. **Isti fajl `modMasterSync` — jedna sesija.** |
| **M5** | Izveštaji i faktura | RF-06, RF-07, RF-08 | 011–013, 003(faktura) | M/L+M+S | Ispravnost brojki, freshness/revers, faktura + štampa. |
| **M6** | SEF UX + banka import/export | RF-22, RF-09, RF-10 | 032, 007, 014–015 | M+M+S | SEF lifecycle poruke; banka datum-parse (P0 AUD-007), mapiranje, export. |
| **M7** | Otkup UI i palete | RF-11, RF-12 | 021–024 | M+S | frmOtkup/blokovi/kooperant; palete. |
| **M8** | Infra lifecycle + cenovnik + E2E | RF-13, RF-26 | 004, 006, 017, 036, 039 | M+S | `RollbackTx`/journal; stale auto-cena; E2E gate result-contract (AUD-039 latentan). |
| **M9** | Dijagnostika + sledljivost | RF-29, RF-30 | 044, 045, 047, 048 | M+S | Integritet/health „lažni zeleni"; nepotpun trag prikazan kao kompletan. |
| **M10** | Distribucija (self-update + sync IO) | RF-24, RF-25 | 035, 037, 038 | M+M | Self-update component-loss + publish guard; sync/IO (lock RMW, atomski swap, empty-source guard). |
| **M11** | Konsolidacije (P2 dug) | RF-15, RF-16, RF-17, RF-18, RF-19 | (deljeni helperi) | 5×~S | HTTP helperi, **`modBankaParseUtils` (hrani v142 banka)**, `BrutoUNeto`, `NzBlank`, test harness. |
| **M12** | Bezbednost/proces (koordinisano) | RF-20 | 019, KI-001 | L | PIN hash + JMBG (VBA+GAS/PWA migracija), `saveParcelPolygon`, docs cleanup. **Planirati posebno.** |

\* RF-03/RF-04 = obavezna re-verifikacija protiv `origin/main` pre implementacije.

### Redosled (linearno)
`M0 → M1 → M2 → M3 → M4 → M5 → M6 → M7 → M8 → M9 → M10 → M11 → M12`

### Kritični put (ako se ide samo po vrednosti/riziku)
**M0–M4 nose ~40% obima ali gotovo svu bezbednosnu vrednost** (svih 9 P0 + najvažniji P1 lanci:
auth, SEF, finansije, storno, sync). Preporuka: **M0–M4 front-load** (cilj ~4–5 nedelja), pa ostalo
po kapacitetu.

---

## 4. Procena obima i tempo (realno, sa pretpostavkama)

**Jedinica:** „dev-sesija" = jedan fokusiran paket (kod + statička provera), pre operater smoke-a.

| Veličina | Paketi | Dev-sesija/paket |
|---|---|---|
| S / S–M | RF-01, 02, 04, 08, 10, 12, 23, 26, 27, 30 (10) | ~0.75 |
| M | RF-03, 05, 07, 09, 11, 13, 14, 21, 22, 24, 25, 28, 29 (13) | ~1.0 |
| M/L, L | RF-06, RF-20 (2) | ~1.5–2 |
| P2 konsolidacije | RF-15…19 (5) | ~0.75 |

**Ukupno ≈ 28 dev-sesija · 13 release-a.**

**Tempo — realno usko grlo je operater smoke turnaround, ne pisanje koda.**
- Pri **2–3 paketa nedeljno** (kod + smoke + release): **~10–14 nedelja (~3–3.5 meseca)** za ceo program.
- **M0–M4 (kritični put): ~4–5 nedelja** ako operater smoke ide isti/sledeći dan.
- M12 (PIN hash migracija) je izdvojen — traži koordinisani prozor (VBA+GAS+PWA istovremeno), ne
  ulazi u tempo gornjeg reda.

**Pretpostavke:** jedan developer serijski; operater dostupan za smoke u roku od 1–2 dana po paketu;
nema velikih paralelnih feature-a na `main`-u koji bi terali česte re-baze; single build-mašina.

---

## 5. Verifikacija i gejtovi (CI NE pokreće Excel)

Tri nivoa, svaki obavezan:

**A) Statički gejt (po paketu, u repo-u — automatizovljivo):**
- Balans `Sub`/`Function`/`Select Case`; nema duplih `Public` definicija („Ambiguous name").
- `file` = „ASCII text" na svim izmenjenim VBA izvorima; grep ne-ASCII = prazno.
- Svaki nov `Poruka("KLJUC")` ima par u `UpsertPoruke` (0 orphan-a).
- `git merge-tree` protiv `main` = bez konflikata.

**B) Operater smoke gejt (po paketu, u Excelu — čovek):**
- `ImportAllVBA` → `Debug → Compile VBAProject` (mora čisto) → snimi.
- Regression checklista paketa (već u `REFAKTOR_PLAYBOOK.md`, klik-po-klik + očekivani rezultat).
- Paket se NE zatvara dok smoke ne prođe.

**C) Milestone release gejt (pre Fleet publish-a):**
- `RunProductionHealthCheck` + `RunSetupHealthCheck` na ciljnoj svesci — bez blocking FAIL-a.
- Ciljani suite po temi: SEF milestone → `RunSEFTestSuite`; sync → MasterSync smoke; storno → storno testovi.
- `tools/release.sh <verzija>` → `ImportAllVBA` → `Compile` → snimi → ship → **Fleet provera** da se
  novi `AgriX_OtkupApp.xlsm` pravilno verzioniše (vidi `RELEASE_PROCEDURE.md`, dopuni `RELEASE_NOTES.md`).

> **Napomena o AUD-039 (E2E gate):** gate danas prijavljuje PASS na svaki non-throw i **nije pozvan**
> u `PublishReleaseToDrive` (latentan). Do RF-26 (koji popravlja result-contract) **release gejt se ne
> oslanja na E2E gate** — koristi ručne suite + health check gore.

---

## 6. Registar rizika izvršenja

| # | Rizik | Verovatnoća/uticaj | Mitigacija |
|---|---|---|---|
| R1 | **Schema drift** — realne kolone ≠ kod (po instalaciji) | Vis/Vis | `DebugKoloneTabele` pre upisa; upis po imenu; `RequireColumnIndex`. |
| R2 | **`.frx` binarni** — tekst-edit kvari formu | Sred/Vis | Nikad ne editovati `.frx` kao tekst; kontrole runtime; caption u dizajneru. |
| R3 | **Encoding regresija** — ne-ASCII u VBA izvoru | Sred/Vis | ASCII-only; `modPoruke`+`ChrW`; `file` provera posle svake izmene. |
| R4 | **Merge „Ambiguous name"** — dupli `Public` posle merge-a | Sred/Sred | `git merge-tree` pre; `Debug → Compile` posle svakog import-a. |
| R5 | **Storno drift** — RF-03/04 ciljaju stare linije | Vis/Sred | Obavezna re-verifikacija protiv v2.24.0 pre koda; skinuti već-rešeno. |
| R6 | **Deljeni fajl** — RF-14 i RF-28 diraju `modMasterSync` | Vis/Sred | Spojiti u jednu sesiju (M4); izbeći duplu re-bazu. |
| R7 | **Self-update / single build-mašina** — loš publish ruši flotu | Nis/Vis | RF-24 publish guard (placeholder/dirty deny, SHA cross-check) PRE bilo kog fleet push-a; Fleet provera posle. |
| R8 | **Nema automatskog Excel testa** — regresija promakne | Sred/Vis | Operater smoke je obavezan gejt; precizne checkliste; ciljani suite na milestone-u. |
| R9 | **Promena poslovnog modela** — pojava pravog multi-device | Nis/Vis | SWMR-Prihvaćeno stavke dokumentovane u katalogu; re-otvoriti ako se model promeni. |

---

## 7. Definicija završenosti (DoD) — po paketu

Paket je „gotov" tek kad je **sve** ispunjeno:
1. Reference-first: pročitani izvori istine + postojeća implementacija (bez pretpostavki).
2. Samo minimal-delta; `reuse > new`; bez novog sloja bez „rule of three".
3. Statički gejt (§5A) zelen.
4. Operater smoke checklista (§5B) prošla u Excelu.
5. `KNOWN_ISSUES.md` AUD status prebačen (Open → Fixed) + `ARCHITECTURE_REFERENCE.md` §19 sync;
   fiksirano ide u `ARCHITECTURE_CHANGELOG.md`.
6. `RELEASE_NOTES.md` / changelog unos.
7. Re-baziran na svež `main`; `force-with-lease` tek po prikazu i odobrenju.

---

## 8. Praćenje

- **Status po paketu:** `REFAKTOR_PLAYBOOK.md` §4 (tabela RF-01…RF-30 + kolona Grana/Status).
- **Status po milestone-u:** tabela dole (održavati uz svaki release).
- **Zatvaranje nalaza:** `KNOWN_ISSUES.md` §8 (AUD status).

| Milestone | Status | Release verzija | Datum | Napomena |
|---|---|---|---|---|
| M0 Higijena | ✅ gotov | v2.28.2 | 2026-07 | RF-01 (PR #147) |
| M1 Pristup i startup | ✅ gotov | — | 2026-07 | RF-02 (#148) + RF-23 (#149); P0-klasa bezbednosti |
| M2 SEF + agrohemija cena | ✅ gotov | — | 2026-07/08 | RF-21 (#152, P0 409) + RF-27 (#154) |
| M3 Storno | 🟡 1/3 | — | 2026-08 | RF-03 ✅ (#167, +AUD-049 storno izvoda); ostaju RF-04, RF-05. re-verify vs aktuelni main |
| M4 Sync jezgro | ⬜ | — | — | RF-14+RF-28 zajedno |
| M5 Izveštaji i faktura | 🟡 1/3 | — | 2026-08 | RF-06 🟢 PR #175 (AUD-023 zatvoren; `Compile` + `RunIzvestajTests` zeleni, ostaje uporedni pregled izveštaja); ostaju RF-07, RF-08 |
| M6 SEF UX + banka | ⬜ | — | — | |
| M7 Otkup UI i palete | ⬜ | — | — | |
| M8 Infra + cenovnik + E2E | ⬜ | — | — | |
| M9 Dijagnostika + sledljivost | ⬜ | — | — | |
| M10 Distribucija | ⬜ | — | — | |
| M11 Konsolidacije | ⬜ | — | — | P2 dug |
| M12 Bezbednost/proces | ⬜ | — | — | koordinisani prozor |

---

## 9. Preporučeni prvi korak

**M0 (RF-01, balast) kao zagrevanje**, pa odmah **M1 (RF-02 + RF-23)** — jer nosi P0-klasu
bezbednosti (auth lanac do Admin panela) i finansijske guardove uz najmanji obim. Alternativa za
najbrži „win": **RF-27 (agrohemija cena)** može ići kao samostalni S-fix pre svega ostalog — jedan
argument (`overrideCena`), potvrđen asimetrijom ulaz/izlaz, direktan finansijski/audit efekat.
