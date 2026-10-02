# Hardening review — 20 tačaka na trenutnom `main`-u

**Datum pregleda:** 2026-10-02  
**Repo:** `dusanmiladinovicVNM/otkupapp-pwa`  
**Pregledani branch:** `main`  
**Main head u trenutku pregleda:** `25e482c1df45809f2198f913c65fc88c00139f92`  
**Poslednji uračunati merge:** PR #400 — `feat(ambalaza): AMB-10b-1 -- knjiga dobija pisca, bez cutovera`

> Ovo je source-level review protiv aktuelnog GitHub `main`-a.  
> Nije dokaz da je konkretna `.xlsm` sveska prošla `ImportAllVBA`, ručni VBA compile, `RunBusinessFlowProSuite`, `RunProductionHealthCheck` ili E2E release gate.

## Sažetak verdikta

U odnosu na prethodni pregled, `main` je značajno ojačan kroz merge-ovane rezove #394–#400:

- S5-5b: otkup žica kao zaglavlje + stavke kroz VBA/GAS/PWA.
- Schema migracije i format ugovor su ozbiljno ojačani.
- AMB-10 je prešao iz domenskog modela u ugovor + kanon + pisac knjige.
- PWA/GAS sloj sada ima JS harness i CI kapije.
- `make_fixture` pin-uje `OTKUP_BRUTO_UNOS = NO`, čime je zatvoren KI-008, ali bruto grana ostaje namenski nemerena.

Procena stanja 20 hardening tačaka:

```text
Rešeno / praktično rešeno:  9 / 20
Delimično rešeno:           8 / 20
Nije rešeno:                3 / 20
```

Najkraći opis trenutnog stanja:

```text
Hardening nivo:        značajno jači nego pre #394–#400
Production proof:      još nije kompletan
Najveći preostali dug: shared exact-row guard + MsgBox/UI boundary + runtime gates
AMB-10 stanje:         odličan model + pisac, ali bez produkcionog cutovera
JS/PWA/GAS stanje:     prvi put ozbiljno pod CI kapijom
```

---

## Status 20 tačaka

| # | Tačka | Status | Nalaz |
|---:|---|---|---|
| 1 | Kritičan `UpdateCell` → `RequireUpdateCell` | Delimično | Jaki tokovi koriste `RequireUpdateCell`, ali nema dokaza da je 100% svih starih `UpdateCell` poziva pokriveno. |
| 2 | Svaki kritičan lookup exact-one | Delimično jače | Ambalaža ugovor sada fail-closed razrešava `Tip + ID` naloga i odbija duplikate. Faktura/Zbirna već imaju lokalne guardove. Ali Known Issues i dalje kaže da su exact-row guard helperi delom lokalni. |
| 3 | Zabraniti update preko filtriranih array-ja | Delimično | Smer je bolji kroz kanonske čitače i fail-closed modele, ali ne postoji globalni statički enforcement. |
| 4 | Centralni `RequireSingleRow` | Nije rešeno | I dalje nema zajedničkog `modDataAccessGuards.RequireSingleRow`. Lokalni `RequireSingle*` obrasci postoje. |
| 5 | Svaki `_TX` snapshotuje sve menjane tabele | Delimično | Jakim postojećim tokovima je dobro. AMB pisac sam kaže da AMB-INV-08 — dokaz da upis živi u TX izvornog dokumenta — ide uz cutover `10b-2`. |
| 6 | Zabraniti business write iz UserForm handlera | Delimično | Arhitektonski smer je dobar, ali nema globalnog dokaza da UI više nikad ne piše domain state direktno. |
| 7 | Smanjiti `On Error Resume Next` van best-effort zona | Delimično | Monitoring/file/Drive best-effort zone su uglavnom razgraničene. Nije dokazano globalno očišćeno. |
| 8 | Standardni error contract po modulu | Delimično jače | AMB uvodi imenovane error brojeve za deficit, identitet, knjiga-ulaz/kvar itd. Dobar obrazac, ali ne globalno zatvoren. |
| 9 | Nema `MsgBox` u business modulima | Nije rešeno | Known Issues i dalje vodi residual business-layer `MsgBox` usage kao aktivan cleanup. |
| 10 | Svaki kritičan `AppendRow` proveriti | Praktično rešeno u novim kritičnim tokovima | AMB novi pisci proveravaju `AppendRow <= 0` za `tblAmbalaza`, `tblAmbalazaDokument` i red knjige. Ne dokazuje svaki stari `AppendRow`, ali nova kritična površina je ispravna. |
| 11 | `GetNextID` nije jedina zaštita | Rešeno | Transakcioni identiteti sve više idu preko opaque `NewEntityID`; duplicate document keys su deo health/gate priče. |
| 12 | Hardening `modVbaTools` lokalnih putanja | Nije rešeno | Hardcoded lokalne putanje i dalje postoje u build/dev alatima. Rizik je izolovan, ali tačka nije zatvorena. |
| 13 | Razdvojiti dev/test module od production build-a | Delimično | Test/gate sistem je znatno jači, ali production packaging nije formalno dokazano čist iz samog `main`-a. |
| 14 | Centralni test runner / regression kapije | Rešeno jače | Postoje VBA gate-ovi i sada JS harness sa `--self-test`. Static workflow pokreće JS sintaksu, `npm ci`, JS harness i self-test. |
| 15 | Testovi čiste za sobom | Delimično | Fixture je jači; testovi i dalje nisu univerzalno clean-room. |
| 16 | `ProductionHealthCheck` kao gate | Rešeno kao contract | Gate postoji i obavezan je u release dokumentima, ali nije dokaz da je konkretna workbook instanca prošla. |
| 17 | External side effects posle commit-a ili recovery | Delimično | BankaImport je dobro rešen. AMB eksplicitno imenuje da AMB-INV-08 mora biti statički dokazano uz pozivna mesta. Google/file-system side effects ostaju accepted boundary. |
| 18 | Monitoring best-effort | Rešeno | Monitoring ostaje best-effort i ne sme obarati business operaciju. |
| 19 | Secrets/config redaction sistematski | Delimično | Monitoring redaktuje glavne tajne; nema dokaza da su svi ostali log/debug putevi sistematski pokriveni. |
| 20 | Svaki accepted risk ima recovery proceduru | Delimično jače | Known Issues je bolji i KI-008 je zatvoren. I dalje postoje accepted risks i technical limitations koji traže follow-up. |

---

## Glavni pozitivni pomaci od prethodnog snapshot-a

### 1. Schema hardening je ozbiljno jači

`modSchema` sada eksplicitno tretira redosled kolona kao deo šeme. `SchemaReadyOrFail`, format ugovor i `VerifySchema` direktno zatvaraju klasu bugova gde pozicioni upis tiho šalje vrednosti u pogrešne kolone.

To jača:

- tačku 10 (`AppendRow` / pozicioni upis),
- tačku 14 (regression gates),
- tačku 16 (release health),
- delimično tačku 5 (TX/write safety).

### 2. AMB-10 više nije samo dokumentacija

Uveden je `modAmbalazaUgovor`:

- zatvorene liste naloga,
- zatvorene liste vrsta kretanja,
- fail-closed razrešavanje naloga,
- provera prenosa,
- storno-svestan doprinos obavezi.

Važna granica: modul sam kaže da ništa ne piše u tabele i da produkcioni cutover nije urađen u tom koraku.

### 3. AMB pisac postoji, ali cutover nije završen

`modAmbalaza.PrenesiAmbalazu` uvodi novi writer model:

- schema guard,
- `RequireAmbPrenos`,
- idempotency kroz zbir zahteva,
- deficit handling,
- upis u append-only knjigu.

Ali trenutni rez ne dira devet pozivnih mesta. `TrackAmbalaza` i dalje piše stari oblik; čitaoci i produkcioni put idu u `10b-2`.

### 4. JS/PWA/GAS sloj više nije bez kapije

Uveden je JS harness:

- green run,
- `--self-test`,
- sabotaže u memoriji,
- CI izvršenje kroz Node 20 i `npm ci`.

To ne zatvara svu PWA/GAS regresionu pokrivenost, ali uklanja prethodni problem da je taj sloj bio uglavnom „pročitan, ne meren“.

### 5. KI-008 je zatvoren, uz imenovani preostali dug

`make_fixture` sada pin-uje:

```text
OTKUP_BRUTO_UNOS = NO
```

Time je zatvorena petorka padova koja je dolazila od nasledjenog donor config-a.

Preostali dug: bruto grana ostaje nemerena i traži poseban test koji sam uključuje `OTKUP_BRUTO_UNOS`.

---

## Glavni preostali dugovi

### A. Shared exact-row guard

Trenutno postoje dobri lokalni obrasci, ali ne i jedan kanonski helper.

Predlog:

```text
modDataAccessGuards.RequireSingleRow(tableName, idColumn, idValue, sourceName) As Long
```

Minimalni ugovor:

```text
0 redova  -> Err.Raise, missing
1 red     -> return row index
2+ redova -> Err.Raise, duplicate/ambiguous
```

Ovo treba uvoditi postepeno, prvo u najkritičnije module.

### B. Business-layer `MsgBox`

Known Issues i dalje priznaje residual business-layer `MsgBox` usage.

Ovo nije samo estetika. U headless/test/automation kontekstu `MsgBox` je potencijalno trajno visenje. Operator messaging treba gurati u UI/form layer, dok business moduli treba da vraćaju razlog ili dižu imenovanu grešku.

### C. Runtime proof

Ovaj dokument nije zamena za runtime kapije.

Pre production handoff-a i dalje je potreban stvarni lanac:

```text
ImportAllVBA
Debug > Compile VBAProject
RunBusinessFlowProSuite
RunProductionHealthCheck
RunE2EReleaseGate_v610
GAS route/smoke checks
PWA/JS CI green na istom SHA
```

### D. AMB-10b-2

AMB-10 je kvalitetno modelovan, ali trenutno stanje je prelazno:

- novi writer postoji,
- ugovor postoji,
- schema postoji,
- ali produkciona pozivna mesta i čitaoci nisu cutover-ovani.

Sledeći AMB posao mora dokazati:

```text
svaki od devet poziva knjiženja živi unutar transakcije izvornog dokumenta
svaki takav TX snapshotuje tblAmbalaza / relevantne tabele
čitaoci ne mešaju stari i novi oblik reda
storno dobija kontra-stav model, ne soft-delete stare knjige
```

---

## Preporučen redosled sledećih poteza

1. **Završiti AMB-10b-2**  
   Pozivna mesta + čitaoci + dokaz AMB-INV-08. Ne širiti domen dok knjiga ne prođe kroz stvarne tokove.

2. **Uvesti shared exact-row guard**  
   Ne kao masovni refaktor, nego prvo kroz najopasnije tokove: faktura, prijemnica/zbirna, banka mapiranje, storno.

3. **Statički audit `UpdateCell(` i `AppendRow(`**  
   Napraviti alat/grep listu i ručno klasifikovati: checked, intentionally unchecked, best-effort, legacy/dead.

4. **Business `MsgBox` cleanup**  
   Prvo pronaći sve pozive u business modulima; zatim prebacivati operator poruke u UI layer.

5. **Runtime gate snapshot**  
   Zabeležiti konkretan SHA + workbook fixture/donor + rezultate compile/BFP/health/E2E. Bez toga se source-level zaključak ne sme zvati production proof.

---

## Finalni verdict

Trenutni `main` je sada bliže **9/20** nego ranijih **7/20** hardening tačaka.

Najvažniji kvalitetni skok nije samo broj zatvorenih tačaka, nego promena kulture:

```text
sve manje pravila stoji kao komentar,
sve više pravila ima kapiju,
sabotažu,
fixture,
ili CI harness.
```

Ipak, sistem još nije 20/20 hardening dok se ne zatvore:

```text
shared exact-row guard,
business-layer MsgBox cleanup,
AMB-10 produkcioni cutover,
stvarni runtime proof nad workbook-om.
```
