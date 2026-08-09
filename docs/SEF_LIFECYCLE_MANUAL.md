# SEF — uputstvo za rad i test-lista (RF-22 / v2.37.0)

Status: **operativno uputstvo + smoke lista za paket RF-22 (AUD-032).**
Aplikacija: **OtkupApp**, ekran **SEF upravljanje** (`frmSEF`).
Verzija u kojoj je uvedeno: `vba-v2.37.0`.

> Kada koristiti koji dokument:
> - **ovaj fajl** — „šta koja poruka/dugme sada znači" i test-lista za v2.37.0;
> - `docs/production-runbook-sef-slanje-faktura.md` — incident runbook „ne mogu
>   da pošaljem fakturu" (dijagnostika, tabele, SQL-nivo provere);
> - `docs/KNOWN_ISSUES.md` §8.2 (AUD-032) — istorija nalaza i zašto je odlučeno
>   ovako;
> - `docs/RELEASE_NOTES.md` → `vba-v2.37.0` — sažetak za korisnika.

---

## 1. Dva stanja koja se ne smeju mešati

Svaka faktura nosi **dva** odvojena podatka, i oba se vide na SEF ekranu:

| Polje | Ko ga postavlja | Šta znači |
|---|---|---|
| `SEFWorkflowState` | **naša aplikacija** | dokle je naš proces stigao (`SEF_READY`, `SEF_SENDING`, `SEF_SENT`, `SEF_ACCEPTED`, `SEF_REJECTED`, `SEF_TECH_FAILED`, `SEF_STORNO`, `SEF_SYNC_ERROR`, `SEF_UNKNOWN`) |
| `SEFStatus` | **SEF** (verbatim odgovor) | šta SEF kaže o dokumentu (`Approved`, `Rejected`, `Sent`, `Seen`, `Mistake`, `Storno`, `Cancelled`, `Paid`…) |
| `SEFDocumentId` | **SEF**, jednom | broj dokumenta na SEF-u. **Postoji samo ako je SEF stvarno primio dokument** i ne briše se kad provera statusa padne |

Pravilo koje je uvedeno u v2.37.0: **`SEFStatus` je promenljiv, `SEFDocumentId` nije.**
Zato se odluka „sme li da se šalje" vezuje za `SEFDocumentId`, a ne za status —
pad mreže pri proveri statusa prepiše status, ali ne briše broj dokumenta.

---

## 2. Kako se SEF status prevodi u značenje

Jedan spisak (`ClassifySEFExternalStatus`) koriste prikaz, boja, dugmad, batch i
sve provere — pa ne mogu da se raziđu.

| SEF status | Klasa | Značenje | Lokalno stanje posle „Osveži status" |
|---|---|---|---|
| `Approved`, `Accepted` | ACCEPTED | kupac odobrio | `SEF_ACCEPTED` |
| `Rejected` | REJECTED | kupac odbio | `SEF_REJECTED` |
| `New`, `Draft`, `Sending`, `Sent`, `Seen` | PENDING | još nema odluke | `SEF_SENT` |
| `Storno` | STORNO | dokument storniran | `SEF_STORNO` |
| `Cancelled`, `Deleted` | TERMINAL | dokument otkazan/obrisan na SEF-u | `SEF_SENT` (izlazak iz „šalje se"; lokalnog `CANCELLED` stanja **nema** — namerno, vidi §6) |
| `Mistake` | SEND_FAILED | **greška prilikom slanja dokumenta** | `SEF_TECH_FAILED` |
| `Paid`, `OverDue`, `Archived` | INFO | ne govori o odluci kupca, ali dokazuje da dokument nije više u slanju | `SEF_SENT` |
| `Error` | ERROR | tehnička greška | iz `SEF_SENDING` → `SEF_UNKNOWN`; iz `SEF_SENT` → `SEF_SYNC_ERROR`; ostalo netaknuto |
| prazno / bilo šta drugo | UNKNOWN | nepoznat status → **ručna provera** | isto kao ERROR; u `SEFStatus` se upisuje `UNKNOWN_STATUS` |

**Šta ćete primetiti u praksi:** kod prihvaćene fakture u polju statusa sada piše
`APPROVED` (ranije je program poznavao samo `Accepted`).

---

## 3. Poruke posle klika „Pošalji na SEF"

Ranije je posle svakog slanja pisalo „Faktura poslata". Sada poruka prati ishod:

| Ishod | Poruka | Šta uraditi |
|---|---|---|
| `SEF_SENT` | „Faktura poslata na SEF." + SubmissionID | ništa — sačekati odluku kupca |
| `SEF_ACCEPTED` | faktura prihvaćena | ništa |
| `SEF_REJECTED` | **žuto** upozorenje da je SEF ODBIO fakturu | pročitati „Poslednja greška" → ispraviti → „Pripremi za ponovno slanje" |
| `SEF_TECH_FAILED` | tehnička greška pri slanju | vidi §6 (zavisi da li dokument postoji na SEF-u) |
| bilo šta drugo | nepoznat ishod → ručna provera | proveriti na SEF portalu |

Stanje fakture se u svim slučajevima **snimi pre poruke** — poruka ne odlučuje
ni o čemu, samo prijavljuje ono što je već upisano.

---

## 4. Kada je koje dugme aktivno

| Dugme | Aktivno kada |
|---|---|
| **Pošalji na SEF** / „Retry slanje na SEF" | workflow je `LOCAL_FINALIZED` / `SEF_READY` / `SEF_TECH_FAILED` **i** `SEFDocumentId` je prazan. Odbijena faktura sme samo iz `SEF_READY` (tj. posle pripreme). Caption je „Retry…" samo iz `SEF_TECH_FAILED` |
| **Osveži status** | `SEF_SENT`, `SEF_SYNC_ERROR`, `SEF_UNKNOWN`, i `SEF_TECH_FAILED` **ako postoji** `SEFDocumentId` |
| **Pripremi za ponovno slanje** | workflow `SEF_REJECTED` |
| **Otkaži slanje na SEF** | status `Draft`, `New`, `Mistake`, `Error` |
| **Storniraj u SEF-u** | status `Sent`, `Accepted`, `Approved`, `Rejected` |
| **Recover sending** | workflow `SEF_SENDING` |

Dugmad i kapije (validator) koriste **iste** funkcije (`CanSendSEFInvoice`,
`CanCancelSEFStatus`, `CanStornoSEFStatus`), pa forma više ne može da ponudi
akciju koju kapija odbija. Kad je slanje blokirano, poruka kaže šta je sledeći
korak (otkazivanje / storniranje / provera na portalu).

---

## 5. Batch akcije i recovery — kako čitati rezultat

- **„Osveži sve Pending"** → sažetak `Scanned / Refreshed / Unresolved /
  SkippedTerminal / Failed`.
- **„Recover sve sending"** → sažetak `Found / Recovered / NotRecovered / Failed`.
- **„Recover sending"** (jedna faktura) → „Recovery završen" **samo** ako je
  faktura stvarno izašla iz `SEF_SENDING`; inače eksplicitno „Recovery NIJE
  uspeo".
- **„Osveži status"** → ako poziv ka SEF-u padne, poruka je „SEF status NIJE
  osvežen… ručna provera", ne uspeh.

Brojači više ne računaju „nije puklo" kao uspeh.

---

## 6. Tipične situacije (šta operater radi)

**a) SEF je odbio fakturu (`Rejected` / workflow `SEF_REJECTED`).**
„Poslednja greška" → ispraviti fakturu → **„Pripremi za ponovno slanje"** →
„Pošalji na SEF". Priprema u istoj transakciji razdužuje i prethodni submission,
pa slanje više ne pada na proveru duplikata. Ako fakturu blokira neka *ranija*
uspešna predaja, priprema **staje sa jasnom porukom** — tada ide ručna provera na
portalu, ne ponovno slanje.

**b) Status je `Mistake` (workflow `SEF_TECH_FAILED`, dokument postoji na SEF-u).**
Putanja je **„Otkaži slanje na SEF" + ručna provera**, a **ne** ponovno slanje.
Dugme za slanje je namerno ugašeno: da li SEF prihvata ponovnu predaju istog
dokumenta nije utvrđeno, a slanje fakture je pravni čin — ne pretpostavlja se.
Dugme ostaje ugašeno i posle uspešnog otkazivanja (status `Cancelled`) i posle
pada mreže pri proveri statusa.

**c) Tehnički pad pri slanju, bez `SEFDocumentId`.**
Dokument nije ni stigao do SEF-a → dugme se zove „Retry slanje na SEF" i radi.

**d) Faktura zaglavljena u `SEF_SENDING`.**
„Recover sending" ili automatski recovery na startu. Faktura izlazi iz „šalje se"
i kad je na SEF-u stornirana / otkazana / plaćena / arhivirana, i kad provera
statusa padne (tada ide u `SEF_UNKNOWN`, odakle „Osveži status" i dalje radi).
Više ne postoji slučaj u kom se pri **svakom** pokretanju upisuje lažan zapis o
oporavku.

**e) Status `UNKNOWN_STATUS`.**
SEF nije vratio status. Faktura ide na **ručnu proveru** (žuto u pregledu,
upozorenje u monitoringu); lokalno stanje se **ne** pomera napred. Ponoviti
„Osveži status" kasnije, ili proveriti na portalu.

**f) `Cancelled` / `Deleted` — zašto nema lokalnog `SEF_CANCELLED`.**
Namerna odluka: otkazivanje/brisanje je **external-terminal-only** — beleži se u
`SEFStatus`, a lokalni workflow se samo izvlači iz „šalje se". Lokalni state
machine nema `WF_SEF_CANCELLED` i ne uvodi se u ovom paketu.

---

## 7. Instalacija ove verzije (obavezni koraci)

```bash
cd ~/Documents/GitHub/otkupapp-pwa
git fetch origin claude/rf-22-sef-ux-lifecycle-kzzzvn
git checkout claude/rf-22-sef-ux-lifecycle-kzzzvn
git pull --ff-only origin claude/rf-22-sef-ux-lifecycle-kzzzvn
```

Opciono, pre Excela (ne traži Excel):

```bash
python3 tools/check-sef-asserts.py     # ocekivano: assert-vs-code mismatches: 0
```

U Excelu:

1. `Alt+F8 → ImportAllVBA`
2. `Debug → Compile VBAProject` — **mora čisto**
3. `Alt+F8 → EnsurePoruke` — **obavezno**, 7 novih ključeva
4. Snimi

> Re-import obuhvata: `frmSEF` (+ `.frx` par), `modSEFService`, `modSEFStatusSync`,
> `modSEFClient`, `modSEFValidator`, `modSEFPersistance`, `modConfig`, `modPoruke`,
> `modSEFTests`.
> Bez koraka 3 poruke pišu `[SEF_MSG_...]` — to je tada očekivano, ne bug.

---

## 8. Test-lista za v2.37.0

Legenda statusa: ✅ potvrđeno na operaterskoj mašini · ⬜ nije još izvršeno.

### A) Automatski gate — prvo ovo

| # | Radnja | Očekivano | Status |
|---|---|---|---|
| A1 | `Alt+F8 → RunSEFTestSuite` | `Failed=0`. Suite **ne zove pravi SEF**. Ako ijedna provera padne, makro se zaustavi greškom (tvrd gate) | ✅ all green |
| A2 | Journal folder pored radne sveske | **Nema** novih CSV zapisa sa `TEST-SEF-*` (dva testa rade nad tabelama, ali sve poništavaju uz ugašen journal) | ⬜ |
| A3 | Zatvori pa otvori radnu svesku | **Nema** upozorenja o mogućem gubitku podataka | ⬜ |
| A4 | Provera da su tabele netaknute | U `tblFakture`/`tblSEFSubmission` nema redova `TEST-SEF-RESUB-*`, `TEST-SEF-STORNO-*` | ⬜ |
| A5 | `Alt+F8` lista makroa | `Test_CancelInvoiceOnSEF_TX` i `Test_StornoInvoiceOnSEF_TX` **nisu tu**; nema ni `Test_SendInvoiceToSEF_TX`, `Test_RecoverStuckSEFSendingInvoice`, `Test1/Test2_RefreshSEFStatus_TX`. `RunSEFCancelLiveSuite` / `RunSEFStornoLiveSuite` **jesu** (iza tri kapije) | ⬜ |

**Ako A1 padne — stani.** Sve ostalo nema smisla dok gate nije zelen.

### B) Forma bez slanja (bilo koja sveska)

| # | Radnja | Očekivano | Status |
|---|---|---|---|
| B1 | SEF ekran → izaberi fakturu → „Učitaj fakturu" | Prikažu se podaci | ⬜ |
| B2 | **Izaberi drugu fakturu — bez klika na „Učitaj"** | Status, SEF broj dokumenta, verzija i event log se **isprazne** | ⬜ |
| B3 | Prođi kroz fakture u raznim stanjima | Nema pada; dugmad se pale/gase po stanju (§4) | ⬜ |

### C) Demo SEF — ishodi slanja

> **Ne raditi na produkciji.** Potvrditi da je `SEF_ENV` demo pre početka.

| # | Radnja | Očekivano | Status |
|---|---|---|---|
| C1 | Pošalji ispravnu fakturu | „Faktura poslata na SEF." + SubmissionID | ⬜ |
| C2 | Pošalji fakturu koju SEF odbija | **Žuto** „SEF je ODBIO fakturu (REJECTED)…", **ne** „poslata". Workflow `SEF_REJECTED`, „Pripremi za ponovno slanje" aktivno | ⬜ |
| C3 | Kupac odobri fakturu → „Osveži status" | `SEFStatus` = **`APPROVED`**, zeleno; workflow `SEF_ACCEPTED`; „Storniraj u SEF-u" aktivan | ⬜ |
| C4 | Tehnički pad (prekini mrežu **tokom** slanja) | Workflow `SEF_TECH_FAILED`, dugme **„Retry slanje na SEF" aktivno i radi** | ⬜ |

### D) Ključni scenariji iz review-a

| # | Radnja | Očekivano | Status |
|---|---|---|---|
| D1 | Faktura sa statusom `Mistake` | Workflow `SEF_TECH_FAILED`, status **crven**, „Otkaži slanje" **aktivan**, dugme za slanje **ugašeno** | ⬜ |
| D2 | D1 → isključi mrežu → „Osveži status" | „SEF status NIJE osvežen… ručna provera". Dugme za slanje **ostaje ugašeno** | ⬜ |
| D3 | D1 → „Otkaži slanje na SEF" (uspešno) | `SEFStatus` = `CANCELLED`; dugme za slanje **i dalje ugašeno** | ⬜ |
| D4 | **Storniraj** fakturu (iz `SENT` ili `ACCEPTED`) | **Workflow = `SEF_STORNO`** ← *ključno, ranije je ostajalo `SEF_SENT`*; `SEFStatus` = `STORNO` | ⬜ |
| D5 | Posle D4 → „Osveži sve Pending" | Ta faktura se **više ne obrađuje** (terminalna) | ⬜ |
| D6 | Faktura odbijena **tek pri proveri statusa** → „Pripremi za ponovno slanje" → „Pošalji na SEF" | **Slanje prolazi** (ranije: greška o duplikatu). Posle pripreme `SEFStatus` je i dalje `REJECTED` | ⬜ |
| D7 | Zaglavljena `SEF_SENDING`, na SEF-u stornirana → restart Excela **dva puta** | Prvi start je prebaci (završi u `SEF_STORNO`) i javi „Recovery završen"; drugi start je **više ne nalazi** | ⬜ |
| D8 | „Recover sending" uz simuliran pad | „Recovery **NIJE** uspeo — i dalje `SEF_SENDING`". U event log-u **nema** „Recovered" | ⬜ |
| D9 | „Osveži sve Pending" / „Recover sve sending" | Poruka nosi sažetak (`Scanned/Refreshed/Unresolved/…` odn. `Found/Recovered/NotRecovered/Failed`) | ⬜ |

### E) Regresija — mora raditi kao pre

| # | Radnja | Očekivano | Status |
|---|---|---|---|
| E1 | Fakturisanje, otkup, štampa | Nepromenjeno — paket ne dira ništa van SEF-a | ⬜ |
| E2 | Startup aplikacije | Bez novih poruka; SEF recovery i dalje non-blocking | ⬜ |

---

## 9. Šta je do sada verifikovano

| Provera | Kako | Rezultat |
|---|---|---|
| `RunSEFTestSuite` (12 testova, uklj. matricu 12 stanja × 9 klasa) | Excel, operaterska mašina | **prošlo, all green** |
| `tools/check-sef-asserts.py` (assert-i vs produkcioni izvor) | CI/lokalno, bez Excela | 0 mismatch-eva |
| Statičke provere (ASCII-only, balans `Sub`/`Function`, nema modul-level deklaracija posle prve procedure, nema duplih `Public`) | skripte | čisto |
| Operaterski smoke A2–E2 | ručno | **nije još izvršeno** |

**D4 i D7 su najvažniji preostali koraci** — nalazi iz poslednje dve runde
review-a i jedini deo koji nije izvršen ni u kom obliku van statičkih provera.

Ako C3 pokaže nešto drugo od `APPROVED`, zabeležiti tačan string — klasifikacija
je pisana po zvaničnom `SalesInvoiceStatus` enum-u; nepoznat status je
fail-closed (ide na ručnu proveru), pa ne škodi, ali treba da uđe u spisak u
`ClassifySEFExternalStatus`.
