# 09A — Portfolio proizvoda i launch-readiness (razrada QA Decision Log-a)

**Status:** Review
**Vlasnik:** osnivač AgriX-a
**Poslednje ažuriranje:** 2026-07-24
**Sidro koda:** `origin/main` v2.24.0 (`9fd7087`)
**Povezani dokumenti:** `09_QA_DECISION_LOG.md` (izvor odluka), `07_PRODUCT_PORTFOLIO.md`, `07A_PRODUCT_STATUS_MATRIX.csv`, `07B_ENTERPRISE_OPERATING_MODES.md`, `08_PRODUCT_ROADMAP.md`, `08A_ROADMAP_MILESTONES.csv`, `../PLAN_SANACIJE.md`, `../REFAKTOR_PLAYBOOK.md`, `../KNOWN_ISSUES.md` §8, `../RELEASE_GATES.md`

> **Cilj poglavlja.** Razraditi `09_QA_DECISION_LOG.md` (260 komercijalnih odluka) u **prodajni portfolio** i, za svaku prodajnu celinu, dati **launch-readiness presudu** — gde je najviše posla da _sve_ bude spremno za lansiranje.
>
> **Ovaj dokument NE duplira izvore.** Inženjerski registar defekata živi u `KNOWN_ISSUES.md` §8 (AUD-001…048), izvršni program u `PLAN_SANACIJE.md`, paketi u `REFAKTOR_PLAYBOOK.md` (RF-01…RF-30), komercijalne odluke u `09_QA_DECISION_LOG.md`. Ovde se ta dva sveta **spajaju** u jednu tabelu: _šta prodajemo_ × _šta blokira prodaju_.

---

## 1. Potvrđene polazne činjenice

- Tri klijenta koriste Desktop u produkciji; management koristi PWA za pregled i kontrolu. (`Master Plan/README.md`)
- Plan naredne sezone: ~5 novih firmi, ~8 ukupno; preko 10 nije cilj. Cilj 2027: 10–20 aktivnih pravnih lica (QA #249).
- Prosečna firma: ~10 stanica, 1 desktop korisnik, ~100 kooperanata.
- Gazdinstvo cene: **Basic 19 EUR / Pro 39 EUR**; prvih 50 Partner naloga uz paket hladnjače (QA #159/#161, README).
- Očekivanje: proizvodni modul koristi >80% Enterprise klijenata (QA #250).
- Inženjerski triaž (FM v35→v85→v142 + nezavisni review): **48 dedupliranih AUD stavki → 30 RF paketa**; kalibrisano za single-writer desktop → **9 stvarnih P0, ~34 P1 klastera, 5 P2**. Jedini _nov aktivan_ P0 je AUD-030 (SEF 409). (`PLAN_SANACIJE.md` §1)
- **Status sanacije: samo RF-27 (agrohemija cena) je urađen** (grana, pre-merge). RF-01…RF-30 (osim 27) su ⬜ otvoreni. (`REFAKTOR_PLAYBOOK.md` §4)
- **Runtime gate-ovi nisu izvršeni** u ciljnom workbook/GAS/PWA okruženju (KI-002, KI-003) — dokumentacija može biti čista dok runtime nije proveren (rizik false-green).

---

## 2. Razrada QA log-a → komercijalna arhitektura

QA log definiše **jedan proizvod** (QA #59: nema small/mid/large izdanja), a razlike se rešavaju **paketima + modulima + brojem stanica + konfiguracijom**. Ekosistem ima pet stubova:

```
                    ┌─────────────────────────────────────────────┐
                    │        AgriX Platform Services              │  (u svakom plaćenom proizvodu)
                    │  auth · offline queue · sync/MasterSync ·   │
                    │  storno/audit · monitoring · self-update ·  │
                    │  backup/recovery · release gates            │
                    └─────────────────────────────────────────────┘
   ┌───────────────────────┐   ┌──────────────────┐   ┌──────────────────┐
   │   AgriX ENTERPRISE     │   │  AgriX GAZDINSTVO │   │   AgriX SAVETNIK  │
   │  (glavni B2B proizvod) │   │  (proizvođač/     │   │  (agronom vodi    │
   │                        │   │   kooperant)      │   │   više gazdinstava)│
   │  Paketi:               │   │  Partner /        │   │  po aktivnom      │
   │   • AgriX Desktop      │   │  Basic 19€ /      │   │  gazdinstvu       │
   │   • AgriX Mobile       │   │  Pro 39€          │   └──────────────────┘
   │  Moduli (odvojeno):    │   └──────────────────┘
   │   • SEF · Banka        │            ▲
   │   • Dispatch (napredni)│            │ compliance sloj
   │   • Hladnjača/Proizv.  │   ┌──────────────────┐
   └───────────────────────┘   │    AgriX GGAP    │  (Enterprise dodatak; kupuje hladnjača za mrežu)
                               └──────────────────┘
```

### 2.1 Cenovni model (izvučen iz QA log-a)

| Pravilo | Odluka | QA ref |
|---|---|---|
| Osnovna jedinica | Godišnja pretplata **po pravnom licu** (nema mesečno, nema per-user/uređaj) | #25/#26/#29/#31 |
| Stanice | Osnovni paket do **5 stanica**; svaka preko 5 = ista fiksna godišnja cena (nema tier-ova) | #121/#122 |
| Cena stanice | Ista u Desktop i Mobile; razlika je u ceni osnovnog paketa | #123 |
| Mobile multiplikator | Desktop Otkup + Mobile ≥ **2× Desktop Otkup** | #126/#127 |
| Moduli | SEF/Banka/Dispatch = fiksna godišnja **po pravnom licu**, ista bez obzira na paket, plaća se jednom i važi kroz sve instance | #124/#125/#135 |
| Proizvodni dodatak | Fiksna godišnja **tek posle standardizacije**; jedan pogon = jedan dodatak; dodatni pogon = dodatna Desktop instanca | #130/#131/#132/#243/#244 |
| Gazdinstvo | Samo godišnja; Basic 19€, Pro 39€; prvih 50 Partner uz paket hladnjače | #159/#161/#172 |
| Savetnik | Po aktivnom gazdinstvu; javna cena se **ne objavljuje** (individualna ponuda) | #198/#225 |
| GGAP | Fiksna godišnja po pravnom licu, pokriva sve kooperante; nema per-user | §8 |
| Održavanje | Bug-fix u scope-u + bezbednost + ažuriranja = u pretplati; novi zahtevi odvojeno (T&M) | #23/#24/#107 |

### 2.2 Šta je Core, a šta poseban modul

- **Core (uključeno u Desktop):** otkup + osnovna dokumenta, prijemnice, fakture (kreiranje/evidencija), ambalaža + repromaterijal, agrohemija (kompletan tok zaliha/dugovanja/doziranja), skladište/WMS/palete/sledljivost, standardni izveštaji, **ručni unos novca**, kartice i salda, Management PWA (QA #3/#12/#14/#16/#17/#18/#30/#31).
- **Odvojeno plaćeni moduli:** SEF, Banka (automatizacija), Dispatch (napredni), Hladnjača/Proizvodnja dodatak (QA #4/#128).
- **Granica ka ERP-u:** AgriX **nije** računovodstveni ERP; ne pokriva glavnu knjigu, PDV, završni račun, zarade (QA #72/#73/#74).

---

## 3. Portfolio proizvoda — prodajna tabela

Kolone: **Impl** (implementacioni status, 07A) · **Dokaz** (evidence, 07A) · **Komerc.** (dozvoljena ponuda, 07/09) · **Launch gate** (šta blokira „Standard offer").

| # | Prodajna celina | Sadržaj | Cena (model) | Impl | Dokaz | Komerc. | Launch gate (gde je posao) |
|---|---|---|---|---|---|---|---|
| 1 | **Platform Services** | auth, sync/MasterSync, storno/audit, monitoring, self-update, backup, release gates | ugrađeno | Implemented | Production-proven (malo klijenata) | u svakom paketu | Runtime gate-ovi neizvršeni (KI-002/003); release-truth (RF-26 E2E, RF-24 self-update) |
| 2 | **AgriX Desktop (Enterprise Core)** | otkup→dokument→prijem→faktura→izveštaj + Management PWA | godišnja / pravno lice + stanice >5 | Implemented | Production-proven (3 klijenta) | Standard / Controlled | **Najveća koncentracija posla:** P0/P1 correctness (9 P0, ~34 P1) + health-check green |
| 3 | **AgriX Mobile** | Desktop + PWA Otkupac + PWA Vozač (osnovni) | ≥ 2× Desktop Otkup | Implemented | Limited production | Standard / Controlled | **Strateški najvredniji posao:** end-to-end teren→centrala sync gate (Track B) + sync hardening |
| 4 | **SEF modul** | slanje/status/storno izlaznih, preuzimanje ulaznih, povezivanje | godišnja / pravno lice | Implemented | Limited evidence | Optional extension | **NIJE spremno:** jedini nov P0 (AUD-030 409→REJECTED) + P1 klaster (RF-21/22) |
| 5 | **Banka modul** | uvoz izvoda, povezivanje uplata, rasknjižavanje, avansi, nalozi | godišnja / pravno lice | Implemented | Limited evidence | Optional extension | **NIJE spremno:** knjiži novac na otvaranje forme + dedupe/crash (RF-09/10; AUD-014/025/026/034) |
| 6 | **Dispatch (napredni)** | raspoređivanje vozila/vozača, rute, kapaciteti, dispečerski pregled | godišnja / pravno lice | Implemented | Limited evidence | Optional extension | Osnovni Vozač je Mobile-Core; napredni dispatch = malo dokaza, discovery-scoped |
| 7 | **Hladnjača/Proizvodnja dodatak** | prerada, palete sveže/prerađene, lager, sledljivost; (2027: radni nalozi, norme, linije) | godišnja **posle standardizacije** | Implemented/Partial (palete u prod.) | Limited evidence | Controlled rollout | Palete correctness (RF-12/AUD-029); **pun proizvodni sistem = najveći NOV build (2027), nije launch-blocker sad** |
| 8 | **AgriX Gazdinstvo** (Partner/Basic/Pro) | kartica prema hladnjači, parcele/GIS, tretmani/karenca, troškovi, agrohemija, prognoza | Basic 19€ / Pro 39€ (godišnje) | Implemented/Partial | **Pilot evidence** | **Pilot only** | **Nije primarno kod — validacija:** activation, 30/90/180 retention, WTP (Track G) pre skaliranja |
| 9 | **AgriX Savetnik** | 1 agronom → više gazdinstava; planovi/nalozi u Pro naloge | po aktivnom gazdinstvu | Planned (osnovna v. cilj 2027) | — | još nije u prodaji | Izgraditi osnovnu verziju; nije launch-kritično za sezonu 2026 |
| 10 | **AgriX GGAP** | evidencije/dokazi/rokovi/audit-readiness/export | godišnja / pravno lice | Planned/Discovery | Unvalidated | **Not for sale** | Domain owner + standard + pilot; post-2027 (Track H) |

**Delivery komponente sa zasebnim readiness statusom** (07 §6, QA #48-alt): Kiosk režim (`Controlled rollout`), Tablet paket (`Optional hardware`), Termalna štampa (`Controlled rollout`). Njihova nezrelost **ne obara** status PWA aplikacije (07B §6, ROADMAP §3.3).

---

## 4. Gde ima NAJVIŠE posla — dve ose spremnosti

„Launch ready" ima dve nezavisne ose. Obe moraju proći; trenutno je **inženjerska osa uska grla za postojeće proizvode**, a **komercijalna osa je uska grla za _standardizovanu_ ponudu**.

### 4.1 Osa A — Inženjerska spremnost (correctness / release truth)

Rangirano po koncentraciji posla × prioritetu. Detalji: `PLAN_SANACIJE.md`, `REFAKTOR_PLAYBOOK.md`, `KNOWN_ISSUES.md` §8.

| Rang | Blok posla | Zašto blokira lansiranje | Paketi (status) |
|---:|---|---|---|
| **1** | **Enterprise Core data-safety (P0)** | Sheets JSON read korupcija (AUD-001), import rollback gubi potvrđene redove (AUD-002), pozicijski upisi u finansijske/dokumentne redove (AUD-003), storno „lažni uspeh" lanac (AUD-020), RollbackTx bez CleanUp (AUD-004), hladnjača-lanac tihi otkaz (AUD-005), journal recovery (AUD-006), datum rollover (AUD-007). **Sve ispod Core-a i svakog paketa.** | RF-01→02→03→04→14 (⬜) |
| **2** | **PWA-led sync (Mobile, strateški core)** | MasterSync klaster: duplikat brojeva, wrong-write „Synced" sa praznim VozacID, mešanje otpremnica, station-mirror bez `tblVozaci` reda (AUD-041…046); JSON (AUD-001). Ovo je **glavna diferencijacija** koja se prodaje. | RF-14 + RF-28 (⬜, isti fajl — zajedno) |
| **3** | **SEF modul** | Jedini nov **aktivan P0** (AUD-030: 409→REJECTED → trajno pogrešna/duplirana faktura ka poreskoj); + P1 klaster (stornirana faktura sendable, truncation, double-submit, „Faktura poslata" za grešku) (AUD-031/032). Modul se **ne sme prodavati** dok ovo ne prođe. | RF-21 (P0) → RF-22 (⬜) |
| **4** | **Startup + autorizacija** | Korisnik sa „Matični podaci" dolazi do Admin panela (brisanje tabela, migracija, fleet publish); `Workbook_Open` ne zove `AccessWasDenied` (AUD-033/034); `saveParcelPolygon` auth nepotvrđen (KI-001). Bezbednosna P0-klasa. | RF-23 (⬜) |
| **5** | **Finansije (Banka + avans/storno/novac)** | Banka knjiži novac na otvaranje forme (AUD-014/034), dedupe ispušta multi-account + crash na 3+ kandidata (AUD-025), export može naručiti više od otvorenog (AUD-026); storno plaćanja skriva dug (AUD-021); avans no-op naduvava brojače (AUD-010). | RF-09/10 + RF-02 (⬜) |
| **6** | **Izveštaji + faktura + palete** | Ekrani se ne slažu posle storna (dashboard vs frmDokumenta, AUD-015); kartice kreću od nule (AUD-023); freshness (stari podaci pod novim periodom, AUD-024); reprint storniranog (AUD-027); palete blokiraju legitimnu preradu (AUD-029). | RF-06/07/08/11/12 (⬜) |
| **7** | **Release truth / dijagnostika** | E2E gate false-green (AUD-039), integritet/health false-green (AUD-044/047), self-update gubitak komponente (AUD-035), cenovnik stale cena (AUD-036). Bez ovoga „zeleno" ne znači spremno. | RF-24/26/29/30 (⬜) |
| **8** | **Konsolidacije (tehnički dug)** | Deljeni helperi (HTTP, banka parser, BrutoUNeto), pozicijski→imenski upisi. Smanjuje budući rizik, nije blocker. | RF-15…19, RF-20 (⬜) |

> **Presuda ose A:** najviše posla je u **Enterprise Core correctness + PWA sync** (rang 1–2) — to je launch gate za oba fleg-proizvoda (Desktop + Mobile). Redosled izvršenja (playbook §4): `RF-01 → RF-02 → RF-23 → RF-21 → RF-27✅ → RF-03 → RF-04 → RF-14+RF-28 → RF-06 → RF-07 → RF-08 → RF-22 → RF-09 → RF-10 → RF-11 → RF-12 → RF-13 → RF-26 → RF-29 → RF-30 → RF-24 → RF-25 → RF-15+ → RF-20`. Od 30 paketa **urađen je 1**. Uz to, **runtime gate-ovi + `RunProductionHealthCheck` na ciljnom workbook-u** (KI-002/003) su obavezni pre bilo kog „Standard offer".

### 4.2 Osa B — Komercijalna spremnost (packaging / GTM / isporuka)

Master Plan status tabela: gotovo sva poglavlja su **„nije započeto"**. Za _standardizovanu, ponovljivu_ ponudu (ne ad-hoc prodaju postojećim klijentima) najviše posla je ovde:

| Rang | Blok | Stanje | QA otvoreno pitanje |
|---:|---|---|---|
| **1** | **Cene — apsolutni iznosi** | `10_PRICING` nije započeto. QA log daje _pravila_ (odnose), ali ne i brojeve za Enterprise/Mobile/stanice/module. Bez ovoga nema prave ponude. | §9.3 |
| **2** | **Unit economics + finansijski model** | `11`/`12` nije započeto. ARR/stanici, support cost, gross margin, hardver marža. | — |
| **3** | **Standardizovan onboarding** | Track C (ROADMAP): danas founder-only, ~1 dan; cilj pola dana. Bez checkliste bez skrivenih koraka → ne skalira na 8 firmi. Kada se naplaćuje početni onboarding (free→paid cutover). | §9.4 |
| **4** | **Ugovori / pravni okvir** | `20_LEGAL` nije započeto. SLA (RTO 24h, kritičan odziv 1h), vlasništvo podataka, silo, izvoz pri prestanku, povlačenje saglasnosti za deljenje. | §9.1 |
| **5** | **Sales playbook + GTM + partneri** | `14`/`15` nije započeto; demo model (dummy, bez pilota — QA #255), reference. Formalni uslovi partnerske provizije/atribucije. | §9.5 |

> **Presuda ose B:** za sezonu 2026 sa postojećim + ~5 novih klijenata, najkritičnije je **(1) fiksirati cenovnik** i **(3) standardizovati onboarding** — to su preduslovi za _ponovljivu_ prodaju. Ostalo (finansijski model, legal, sales playbook) je nužno za skaliranje ka 10–20, ali ne blokira prve ugovore.

---

## 5. Launch-readiness po proizvodu (sažeta presuda)

| Proizvod | Presuda | Najveći posao |
|---|---|---|
| **AgriX Desktop (Core)** | Prodaje se postojećim klijentima; **nije još „Standard offer"** dok P0/P1 + health gate ne prođu | Osa A rang 1, 4, 5, 6 |
| **AgriX Mobile** | Strateški core; productizacija u toku (Track B) | Osa A rang 2 (sync) |
| **SEF modul** | ⛔ **Ne prodavati** dok RF-21 (P0) ne prođe | Osa A rang 3 |
| **Banka modul** | ⛔ Ne kao „Standard"; knjiži novac na otvaranje forme | Osa A rang 5 |
| **Dispatch (napredni)** | Optional; osnovni Vozač ide u Mobile; napredni čeka dokaz | — |
| **Hladnjača/Proizvodnja** | Palete se prodaju (Controlled); pun sistem = 2027 build | RF-12 + 2027 roadmap |
| **Gazdinstvo** | Pilot only; **validacija pre skaliranja** (ne kod) | Osa A: platforma stabilna; onda Track G metrike |
| **Savetnik** | Osnovna verzija cilj 2027; nije launch-kritično | — |
| **GGAP** | Not for sale; discovery; domain owner | Post-2027 |

---

## 6. Odluke

1. **Portfolio je jedan proizvod + paketi + moduli** (QA #59). Ovaj dokument je kanonski _spoj_ komercijalnog portfolija i inženjerskog backlog-a; detalji ostaju u izvornim dokumentima.
2. **Nijedna celina ne prelazi u „Standard offer" pre nego što prođe readiness gate iz 07 §12** (end-to-end tok + compile/smoke/regression + realni produkcioni dokaz + monitoring/recovery + onboarding + support boundary + cena).
3. **P0 data-safety (Osa A rang 1) ima apsolutni prioritet** — blokira Core i sve module (ROADMAP §5 stop-pravilo).
4. **SEF i Banka se ne smeju prodavati kao „Standard"** dok RF-21 (P0) i RF-09/10 ne prođu; do tada `Optional extension` uz jasnu ogradu.
5. **PWA-led sync (RF-14+RF-28) je strateški najvredniji inženjerski posao** jer nosi glavnu diferencijaciju (teren→centrala bez ponovnog unosa).
6. **Runtime gate-ovi + `RunProductionHealthCheck` na ciljnom workbook-u su obavezan izlazni uslov** (KI-002/003) — „čista dokumentacija" nije dokaz.
7. **Cenovnik (apsolutni iznosi) i standardizovan onboarding su komercijalni preduslovi** za ponovljivu prodaju naredne sezone.

---

## 7. KPI i pragovi

**Inženjerski (po ROADMAP §9 / 07 §12):**
- 0 izgubljenih potvrđenih zapisa; 0 nekontrolisanih canonical duplikata.
- ≥ 99% sync uspeha bez developerske intervencije.
- Svi P0 zatvoreni root-cause analizom; broj otvorenih RF paketa (trenutno 29/30).
- Ciljni workbook: 0 health failure-a.
- Ručne centralne korekcije na 100 otkupa ispod dogovorenog praga.

**Komercijalni:**
- Fiksiran javni cenovni raspon (Desktop/Mobile) + tačan iznos stanice >5.
- ≥ 1 onboarding izveden po checklisti, izmereno vreme/firmi.
- Support cost održiv za planiranih ~8 firmi.

---

## 8. Rizici

| Rizik | Uticaj | Mitigacija |
|---|---|---|
| „Zeleno" bez izvršenih runtime gate-ova | Lansiranje sa latentnim P0 | KI-002/003 kao hard exit-gate; health-check na ciljnom workbook-u |
| SEF poslat pogrešno/duplirano ka poreskoj (AUD-030) | Pravni/finansijski | RF-21 pre svake SEF prodaje; do tada modul van „Standard" |
| Sanacija sporija od sezone | Feature freeze blokira ispravke | 30-dnevni sezonski freeze (ROADMAP §3.6); RF paketi nezavisno isporučivi |
| Prespora prodaja/standardizacija (QA #93/#94/#98) | Uska grla na 30–50 klijenata | Osa B rang 1/3 pre skaliranja |
| Custom zahtevi razvlače roadmap (QA #98/#103) | Odlaganje Core-a | Pisana namera + eksplicitno šta se odlaže (QA #104/#106) |

---

## 9. Otvorena pitanja (nasleđena iz QA §9)

1. Apsolutni cenovnici (Enterprise, Mobile, stanice, moduli, Gazdinstvo, Savetnik).
2. Tačan trenutak prelaska free → paid onboarding za nove Enterprise klijente.
3. Pravna pravila za povlačenje saglasnosti proizvođača za dodatno deljenje podataka.
4. Formalni uslovi partnerske provizije i atribucije lead-a.
5. Redosled post-2027 inicijativa: GGAP, Savetnik marketplace, multi-Enterprise.

---

## 10. Akcije

| # | Akcija | Vlasnik | Rok / gate |
|---|---|---|---|
| 1 | Izvršiti RF sekvencu po playbook §4, počev od RF-01/02/23/21 | osnivač (+dev) | Track A closeout 31.08.2026 |
| 2 | Pokrenuti runtime gate-ove + `RunProductionHealthCheck` na ciljnom workbook-u | osnivač | pre svakog „Standard offer" |
| 3 | Track B end-to-end sync gate + RF-14+RF-28 | dev | 31.10.2026 |
| 4 | `10_PRICING_AND_PACKAGING.md` — apsolutni iznosi | osnivač | pre naredne sezone |
| 5 | Standardizovan onboarding checklist (Track C) | osnivač | 30.11.2026 |
| 6 | SEF/Banka označiti `Optional extension` sa ogradom dok RF-21/09/10 ne prođu | osnivač | odmah |

---

## 11. Izvori i datum provere

- `09_QA_DECISION_LOG.md` (260 odluka), `07_PRODUCT_PORTFOLIO.md`, `07A_PRODUCT_STATUS_MATRIX.csv`, `07B_ENTERPRISE_OPERATING_MODES.md`, `08_PRODUCT_ROADMAP.md`, `08A_ROADMAP_MILESTONES.csv` — verifikovano 2026-07-24.
- `../PLAN_SANACIJE.md`, `../REFAKTOR_PLAYBOOK.md` (RF-01…30 status), `../KNOWN_ISSUES.md` §8 (AUD-001…048), `../ROADMAP.md` §10, `../RELEASE_GATES.md` — verifikovano 2026-07-24.
- Sidro koda: `origin/main` v2.24.0 (`9fd7087`).

## 12. Change log

| Datum | Izmena |
|---|---|
| 2026-07-24 | Inicijalna verzija: razrada QA log-a u portfolio + launch-readiness (dve ose), presuda po proizvodu. Spaja komercijalni portfolio i inženjerski AUD/RF backlog bez dupliranja izvora. |
