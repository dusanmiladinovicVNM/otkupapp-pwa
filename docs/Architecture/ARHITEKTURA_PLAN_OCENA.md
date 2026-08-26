# Ocena plana „idealni AgriX VBA codebase" + korigovani plan

- **Status:** Analiza / predlog smera. NIJE implementirano.
- **Datum:** 2026-08-26
- **Predmet:** predlog slojevite arhitekture (Host → Presentation → Application →
  Domain → Repository → Infrastructure) sa Sync-om kao izolovanim subsistemom.
- **Metod:** merenje nad `src-vba/` (213 fajla, ~152.700 linija), bez pokretanja
  Excela. Sve brojke u dokumentu su izmerene, ne procenjene.

> **Verifikacioni status:** ovaj dokument ne menja kod. Ništa u njemu nije
> potvrđeno kroz `run_vba.py` (traži Windows + Excel). Brojke su statičke —
> `grep`/`wc` nad izvorom, reproducibilne komandama u §8.

---

## 0. Kratak sud

Smer je tačan. **Redosled je pogrešan, a dijagnoza promašuje najskuplji problem
u codebase-u.**

Plan je napisan kao da AgriX nema Application sloj i kao da UI piše direktno u
tabele. Merenje pokazuje da je **prvo netačno**, a **drugo netačno za upis** i
tačno samo za čitanje. Istovremeno, dva stvarna problema — **12 pisaca nad
`tblOtkup`** i **namerna duplikacija poslovnih pravila između legacy formi i
novih `*Unos` modula** — u planu se ne pominju.

Zbog toga bi plan izvršen po svom redosledu (Application → Domain → Repository)
potrošio najveći deo rada na preimenovanje sloja koji već postoji, a ostavio
netaknutim ono što danas proizvodi bagove.

| | Ocena |
|---|---|
| Ciljna slika (§ „Moj konačni target") | **Dobra** — zadržati |
| Dijagnoza trenutnog stanja | **Slaba** — 4 materijalne greške (M1–M4) |
| Redosled implementacije (10 koraka) | **Loš** — invertovati |
| Predlog fizičke organizacije (`00_Host/`…) | **Odbaciti** — skup, bez efekta |
| CI dependency rules (§21) | **Najbolja stavka u planu** — uraditi prvo |

---

## 1. Šta je izmereno

### 1.1 Presentation → DataAccess

Tvrdnja plana: `UI → DataAccess` i `modScr* → TBL_*` treba ukinuti kao obrazac.

Izmereno:

| Modul | LOC | `TBL_` | `COL_` | `LookupValue` | `AppendRow`/`UpdateCell` |
|---|---|---|---|---|---|
| `modScrDokumenti` | 2322 | 75 | 149 | 20 | **0** |
| `modOtkupUI` | 7520 | 41 | 22 | 3 | **0** |
| `modScrAgro` | 1738 | 5 | 4 | 4 | **0** |
| `modScrBankaUvoz` | 1773 | 4 | 3 | 1 | **0** |
| `modScrFakture` | 1387 | 2 | 2 | 1 | **0** |
| `modScrStorno` | 1349 | 0 | 1 | 1 | **0** |
| `modScrPalete` | 1106 | 0 | 0 | 0 | **0** |
| `modUiScreens` | 341 | 0 | 0 | 0 | **0** |

**Nijedan ekranski modul ne piše u tabele preko `modDataAccess`.** Upis već ide
kroz `*_TX` procedure. Ono što je ostalo je **read-path**: `LookupValue` za
prikaz i `COL_` konstante za indeksiranje grida.

To nije „UI zna bazu" — to je **nedostatak Query sloja**, što plan sam ispravno
prepoznaje u §14 (CQRS-lite), ali ga stavlja na 14. mesto umesto da shvati da je
to *jedini* preostali Presentation dug.

### 1.2 Application sloj — postoji

Tvrdnja plana: „Ovo trenutno najviše nedostaje."

Izmereno — postoji, pod drugim imenom:

```
modOtkupUnos    356 LOC   NoviOtkupUnos / OtkupValidiraj / OtkupUpisi
modDokUnos     1062 LOC   Novi{Otpremnica,Zbirna,Prijemnica}Unos / *Validiraj / *Upisi
modNovacUnos    572 LOC
modAgroUnos     640 LOC
                ----
                2630 LOC
```

Oblik `Novi*Unos` → `*Validiraj` → `*Upisi` je tačno *command → validate →
execute* koji plan predlaže kao `clsKreirajOtkupCmd` → `modAppOtkup.Kreiraj`.

Pored toga: **85 `*_TX` procedura u 18 modula**. To su use-case-ovi imenovani po
mehanizmu (transakcija) umesto po nameri (`KreirajOtkup`). Semantički su
Application sloj.

Uvođenje `modApp*` pored ovoga daje **treće imenovanje istog sloja**, ne novi
sloj.

### 1.3 Vlasništvo upisa — problem koji plan ne vidi

Iz `docs/DOMEN/WHO_WRITES.md` (generisano, `tools/who_writes.py`):

| Tabela | Broj pisaca | Moduli |
|---|---|---|
| `tblOtkup` | **12** | `modAutoHladnjaca`, `modBankaMapiranje`, `modDokumenta`, `modMasterSync`, `modNovac`, `modOtkup`, `modOtkupBlok`, `modSetup`, `modSledljivost`, `modStorno`, `modStornoFlow`, `modStornoRecovery` |
| `tblFakture` | **9** | … |
| `tblNovac` | **7** | … |
| `tblAmbalaza` | **6** | … |
| `tblZbirna` | **5** | … |

Ovo je klasa buga na koju `CLAUDE.md` §2 eksplicitno upozorava: *„isto polje
često piše više modula, pa zakrpa na jednom mestu ostavlja ostale."*

Repository sloj rešava **tačno ovo** — jedan pisač po tabeli. Plan ga svrstava
na **5. mesto od 10**, iza dva koraka koja uglavnom preimenuju postojeće.

### 1.4 Legacy duplikacija — najveći aktivni dug, van plana

`docs/UI_MIGRACIJA_KATALOG.md`, §0.1:

> **„Legacy zadržava svoju kopiju te logike — namerno.”** `frmOtkup` i
> `frmDokumenta` ostaju potpuno operativni dok novi UI ne bude umeo sve; do tada
> se pravilo menja u zajedničkom modulu pa **ručno preslikava** u legacy.

Potvrđeno merenjem:

```
frmOtkup.frm     1308 LOC   poziva modOtkupUnos:  0 puta
frmDokumenta.frm 6500 LOC   poziva modDokUnos:    0 puta   (36 direktnih *_TX poziva)
```

Dakle **7808 linija forme drži paralelnu kopiju pravila unosa**, bez ijednog
poziva ka novom putu. Svaka izmena poslovnog pravila mora ručno u dva mesta.

Posledica za plan: „izvuci pure Domain pravila" (korak 4) u ovom stanju znači
izvlačenje iz **dve divergentne kopije**. Ili se radi dvaput, ili se legacy tiho
razilazi. Ovo mora biti zatvoreno **pre** Domain rada, ne posle.

### 1.5 Sync i SyncControl

Plan (§12) tvrdi „više modula direktno piše SyncControl". Izmereno — **dva**:

- `modStanicaLock` (35 pomena) — **već ima read-merge-write sa fail-closed
  ponašanjem** (`TryReadSyncControlAsDict`, linije 531–577: ako čitanje padne,
  upis se preskače da se tab ne prepiše nepotpuno).
- `modGoogleSyncOrchestrator` (5 pomena) — sopstveni `WriteSheetData` (linija 611).

Problem je realan ali je **jednodnevni**, ne arhitektonski stub. Plan mu daje
težinu koju nema.

`modMasterSync` (4065 LOC) jeste prenatrpan kako plan kaže — čita Drive, parsira
JSON, validira, radi row-level TX, linkuje, radi writeback. Ali već ima
`ImportRowToTblOtkup_RowTX`, idempotency preko `IsDuplicateInMaster`, i
`TestHook_*` seam-ove. To je dobra osnova za razdvajanje, ne za rewrite.

### 1.6 Alat za enforcement — već postoji

`tools/vba_check.py` = **2128 LOC**, sa **self-testom koji dokazuje u oba smera**
(`self_test()`, „katalog sabotaža"), i vezan na `PostToolUse` hook.

To znači da su CI dependency rules iz §21 plana — **najjeftinija stavka u celom
planu**. Nije potrebna nova infrastruktura; potrebna je nova `check_*` funkcija
u postojećem checkeru.

---

## 2. Materijalne greške u planu

### M1 — „Application sloj najviše nedostaje"

Netačno (§1.2). Postoji 2630 LOC u `*Unos` modulima + 85 `*_TX` procedura.

**Posledica ako se ne ispravi:** korak 2 i 3 plana („uvedi `modAppOtkup`",
„prebaci `*Unos` iza tih API-ja") postaju wrapper oko wrappera. Veliki diff,
nula promene ponašanja, i treće ime za isti sloj u codebase-u koji već ima dva.

**Ispravka:** ne uvoditi `modApp*`. **Preimenovati postojeće** `modOtkupUnos` →
`modAppOtkup` *ako se već dira taj fajl*, i uvesti pravilo da nova use-case
procedura ide tamo. Sufiks `_TX` zadržati — on nosi informaciju („ova procedura
otvara transakciju") koju `modApp` prefiks ne nosi.

### M2 — Repository na 5. mestu

Repository je jedini korak koji direktno obara metriku iz `WHO_WRITES.md`.
Sve pre njega je preraspodela imena.

**Ispravka:** Repository je **prvi strukturni korak** posle instrumentacije.

### M3 — „Domain tests — no workbook"

VBA nema host-free runtime. Kod se izvršava isključivo unutar Excela;
`tools/run_vba.py` traži Windows + Excel + `pywin32`. Domain test bez workbook-a
je **fizički neizvodljiv**, bez obzira na čistoću modula.

Ono što jeste izvodljivo i vredno: **Domain test bez table fixture-a** — bez
sejanja tabela, bez rollback-a, bez `clsTransaction`. To je razlika između testa
koji traje 40 ms i testa koji traje sekunde, i to je stvarni dobitak.

**Ispravka:** preformulisati cilj u „Domain testovi bez fixture-a i bez
transakcije". Ne obećavati host-free.

### M4 — Interne kontradikcije oko Repository klasa

Plan u §5 odbacuje interface-e („ne bih silovao jezik", `Implements` „jednog
dana"), a u §20 traži **fake repository** za Application testove.

U VBA bez `Implements` fake se ne može ubaciti — `clsOtkupRepository` je
konkretan tip, `Dim r As clsOtkupRepository` ne prima drugu klasu. Dakle §20
zahteva ono što §5 odbija.

Dodatno: u celom `src-vba/` ima **0 `Implements`** i svega 14 klasa, nijedna
`PredeclaredId`. Uvođenje 5+ repository klasa nosi lifetime management
(ko ih pravi, gde žive, da li se keširaju) koji standardni moduli nemaju.

**Ispravka:** `modRepo*` **standardni moduli**, ne klase. Vrednost Repository
sloja je *ekskluzivno vlasništvo nad `TBL_`/`COL_` konstantama i upisom*, a to
`vba_check` može da nametne nad standardnim modulom isto kao nad klasom — i
tada `modRepo*` postaje jedini fake seam koji ionako već koristiš
(`TestHook_*` obrazac, prisutan u `modMasterSync`, `modBrojevi`).

### M5 — Fizička taksonomija `00_Host/` … `90_Tests/`

VBA namespace je **ravan**. Izmereno: **2222 `Public` simbola** u `.bas`
fajlovima. Folderi u VBE-u ne postoje kao koncept izvoza — `src-vba/` je flat po
konstrukciji, a `.frm` fajlovi moraju u commit sa `.frx` parom (`CLAUDE.md` §3).

Masovno preimenovanje ~200 modula: veliki diff, rizik od kolizije `Public`
imena, i nula uticaja na dependency graf.

**Ispravka:** zadržati flat `src-vba/`. Sloj se izražava **prefiksom imena**
(`modScr*`, `modApp*`, `modDom*`, `modRepo*`, `modQry*`, `modSync*`), a granica
se nameće u `vba_check`, ne u folderu. Prefiks je već ustaljen (`modScr*` radi).

### M6 — `modOtkupUI` tretiran kao gotova ljuska

Plan: „`modOtkupUI` već funkcioniše kao shell".

Izmereno: **7520 LOC, 84 `Public` procedure, 41 `TBL_` + 22 `COL_` reference**
(KPI agregacija nad `TBL_OTKUP`/`TBL_OTPREMNICA`, `FillComboDisplayID` nad
`TBL_STANICE`/`TBL_VOZACI`, partner mape nad `TBL_KOOPERANTI`, lista zbirnih nad
`TBL_ZBIRNA`).

Ljuska je `modUiScreens` (341 LOC, čist registry, 0 `TBL_`). `modOtkupUI` je
ljuska **plus** najveća pojedinačna količina ekranske data-logike u codebase-u.

**Ispravka:** `modOtkupUI` je klijent Query sloja br. 1, ne „gotovo".

### M7 — Plan ne sleće u canonical dokument

`docs/ARCHITECTURE_REFERENCE.md` §0.3: *„Ako nešto nije navedeno u ovom
dokumentu, ne smatra se canonical arhitekturom dok ne bude eksplicitno
uneseno."*

Plan koji ne uđe tamo nije obavezujući ni za ljude ni za agente.

**Ispravka:** svaka faza koja se prihvati završava upisom u
`ARCHITECTURE_REFERENCE.md` + red u `ARCHITECTURE_CHANGELOG.md`. Odluke o
smeru (Repository vlasništvo, ID vs broj) idu kao ADR — nastavak na 0001/0002.

### M8 — Nema izlazne metrike ni po jednoj fazi

10 koraka, nijedan merljiv kriterijum „gotovo".

**Ispravka:** `WHO_WRITES.md` je već generisan, mehanički brojač. „Pisača po
tabeli" je metrika faze Repository. Za Presentation: broj `TBL_`/`COL_`
referenci u `modScr*`. Obe se dobijaju iz `grep`-a i mogu u CI.

---

## 3. Šta u planu ostaje netaknuto (dobro je)

1. **Ciljni dijagram** (`Host → Presentation → {Application | Query} → Domain →
   Repository → DataAccess → Excel`) — tačan i za VBA izvodljiv.
2. **CQRS-lite** (§14): read ne mora kroz Domain. Ovo je jedini deo plana koji
   pravilno oslovljava stvarni Presentation dug (§1.1).
3. **Business ID ≠ broj dokumenta** (§8). Već je delom kodifikovano u
   ADR-0001/0002 i `GeneracijaID` modelu. Vredi ga podići u nepovredivo pravilo.
4. **`AllocateNewNumber` / `ObserveExistingNumber`** (§9). Danas `modBrojevi`
   (626 LOC) radi `MaxSeqFromTable`, `MaxSeqFromGoogleSheet`,
   `BrojZbirneExists`, mirror prefiks i cache — pet strategija u jednom modulu.
   Razdvajanje na *alociraj* vs *upamti viđeno* je čista dobit, naročito za
   naknadni papirni unos.
5. **CI dependency rules** (§21) — najbolja stavka, i najjeftinija (§1.6).
6. **„Ovo nije rewrite, nego reorganizacija ownership-a"** — tačan okvir.

---

## 4. Korigovani plan

Princip: **prvo ono što se može nametnuti mašinski i što obara merljivu
metriku; imenovanje poslednje.**

### Faza 0 — Instrumentacija (pre ijednog pomeranja koda)

Bez izmene poslovne logike. Cilj: da svaka sledeća faza ima crveno/zeleno.

1. Proširiti `tools/vba_check.py` novom proverom `SLOJ`, tabelarno definisanom:

   ```
   modScr*, modOtkupUI   : zabranjeno AppendRow, UpdateCell, GetNextID
   modDom*               : zabranjeno TBL_, COL_, Range, Worksheet, ListObject,
                           MsgBox, WinHttp, modDataAccess, modGoogle*
   modRepo*              : zabranjeno frm, modScr, modApp
   modDataAccess         : zabranjeno modScr, frm, modApp, modDom
   ```

   Uvesti sa **whitelistom postojećih prekršaja** (baseline fajl), da suite ostane
   zelena, a svaki *nov* prekršaj pada. Bez baseline-a se pravilo ne može uvesti
   nad 152k linija.
2. Self-test za `SLOJ` po postojećem obrascu „dokaz u oba smera" (`CLAUDE.md` §5
   to traži jer se menja sam checker).
3. `tools/who_writes.py` → dodati izlaz `--max-writers N` koji vraća exit 2 ako
   neka tabela ima više pisaca od praga. Prag se spušta fazama.

**Izlazna metrika:** `vba_check` prijavljuje `SLOJ` prekršaje; baseline zamrznut.
**Rizik:** nizak — ne dira runtime.

### Faza 1 — Ugasiti legacy duplikaciju

Ovo je preduslov za sve što dira poslovna pravila (§1.4).

4. Za svaki režim koji `modOtkupUnos`/`modDokUnos` već pokrivaju: preusmeriti
   `frmOtkup`/`frmDokumenta` da **zovu isti put**, umesto svoje kopije.
   Ne dirati `.frx`; ne dodavati `WithEvents` (`CLAUDE.md` §3).
5. Kad režim ima jedan put — obrisati mrtvu kopiju iz forme.
6. Tek kad novi UI pokrije režim u celini (po `UI_MIGRACIJA_KATALOG.md`) —
   ukloniti legacy ekran.

**Izlazna metrika:** `grep -c 'modDokUnos' frmDokumenta.frm` > 0; broj `*_TX`
poziva iz `frmDokumenta.frm` pada sa 36.
**Rizik:** visok — dira aktivan produkcioni put. Traži `run_vba` suite na
Windowsu, ne samo `vba_check`.

### Faza 2 — Repository: jedan pisač po tabeli

Najveća pojedinačna dobit (§1.3).

7. `modRepoOtkup` — standardni modul, jedini vlasnik `TBL_OTKUP` + `COL_OTK_*`
   upisa. API po nameri: `InsertOtkup`, `SetOtpremnica`, `SetBrojZbirne`,
   `MarkStornirano`.
8. Redom prevoditi 12 pisača `tblOtkup` na taj API. Svaka konverzija je zaseban
   commit sa svojim testom.
9. Ponoviti za `tblFakture` (9), `tblNovac` (7), `tblAmbalaza` (6), `tblZbirna` (5).
10. Kad tabela ima jednog pisača — dići prag u `who_writes.py`.

**Izlazna metrika:** `tblOtkup` 12 → 1 pisač. Ostale ≤ 2.
**Rizik:** srednji. Mehanička konverzija, ali `modStorno`/`modStornoFlow` pišu iz
transakcionog konteksta — Repository mora da bude TX-neutralan (ne otvara
sopstvenu transakciju).

> **Otvoreno pitanje za odluku:** da li Repository sme da poziva
> `clsTransaction.AddTableSnapshot`, ili to ostaje isključivo pozivaocu.
> Preporuka: **ostaje pozivaocu** — inače Repository postaje vlasnik transakcionog
> scope-a, što je Application posao (plan §7 to ispravno kaže).

### Faza 3 — Query sloj (jedini preostali Presentation dug)

11. `modQryDokumenti` — vraća već formatirane redove grida. `modScrDokumenti`
    gubi 20 `LookupValue` i 149 `COL_` referenci.
12. `modQryOtkup` — KPI agregacije koje danas žive u `modOtkupUI`
    (`SumKgForDate`, `CountForDate`) + lookup liste (`FillComboDisplayID`).
13. Zatvoriti `SLOJ` pravilo za `modScr*`: skinuti baseline za `TBL_`/`COL_`.

**Izlazna metrika:** `TBL_`+`COL_` u `modScr*` → 0; u `modOtkupUI` → 0.
**Rizik:** nizak — read-only put, greška se vidi odmah na ekranu.

### Faza 4 — Domain gde se isplati (ne svuda)

Ne izvlačiti Domain iz svih 8 poslovnih modula. Izvući gde je pravilo gusto i
testirano bez tabela:

14. `modDomStorno` — iz `modStornoImpact` (500 LOC, već 0 Excel objekata) i
    `modStorno` (2412 LOC): `CanStorno`, `ValidateOwnership`, `ValidateGeneration`,
    `ValidateCascade`.
15. `modDomDokument` — iz `modDokumentInvariant` (670 LOC, 1 Excel referenca):
    lifecycle i ownership pravila.
16. `modDomOtkup` — bruto→neto, klase, ambalaža, cena.

Ostalo (`modNovac`, `modFaktura`, `modDokumenta`) **ostaviti** dok ne postoji
konkretan povod. `modDokumenta` (4079 LOC, 290 `TBL_`) je najskuplji za
razdvajanje i najmanje se menja.

**Izlazna metrika:** `SLOJ` pravilo za `modDom*` prolazi bez baseline izuzetka;
Domain testovi rade bez sejanja tabela.
**Rizik:** srednji.

### Faza 5 — Sync, brojevi, SyncControl

17. `modSyncControl` — jedini vlasnik `SyncControl` taba. Preseliti
    `TryReadSyncControlAsDict` + `ApplySyncControlUpdates` iz `modStanicaLock`
    (već fail-closed), i prevesti `modGoogleSyncOrchestrator` na njega.
    *Jednodnevno, uraditi ranije ako se sync ionako dira.*
18. `modBrojevi` → `AllocateNewNumber` / `ObserveExistingNumber`. `NumberRegistry`
    (`Kind | EntityID | BusinessDate | LastSeq`) za cloud namespace; lokalni
    counter za offline. Naknadni papirni unos samo `Observe`.
19. `modMasterSync` → `parse DTO` odvojiti od `ImportRowToTblOtkup`; import zove
    `modAppOtkup` umesto da sam piše. `TestHook_*` seam-ovi ostaju.

**Izlazna metrika:** `grep -l SyncControl src-vba/*.bas` → 1 fajl;
`modMasterSync` ne poziva `AppendRow` nad poslovnim tabelama.
**Rizik:** visok za #19 — sync ima fail-closed ponašanje koje se ne sme oslabiti.

### Faza 6 — Imenovanje (poslednje, opciono)

20. `modOtkupUnos` → `modAppOtkup` itd., **samo uz izmenu koja ionako dira fajl**.
    Nikad kao zaseban „rename commit" preko 200 modula.

### Odbačeno iz originalnog plana

| Stavka | Razlog |
|---|---|
| `00_Host/` … `90_Tests/` folderi | VBA namespace je ravan; prefiks + `vba_check` daju isto (M5) |
| `clsOtkupRepository` i sl. klase | bez `Implements` ne daju fake seam koji §20 traži; `modRepo*` je jeftiniji (M4) |
| `clsKreirajOtkupCmd` command klase | `Object`/`Dictionary` payload koji `modOtkupUnos` već koristi radi isto |
| „Domain tests — no workbook" | fizički neizvodljivo u VBA (M3) |
| Novi `modApp*` pored `*Unos` | treće ime za isti sloj (M1) |

---

## 5. Redosled — original vs korigovan

| # | Original | Korigovano |
|---|---|---|
| 1 | Presentation granica | **CI `SLOJ` pravilo + baseline** |
| 2 | Uvedi `modApp*` | **Ugasi legacy duplikaciju** |
| 3 | Prebaci `*Unos` iza App API-ja | **Repository: `tblOtkup` 12 → 1 pisač** |
| 4 | Izvuci Domain | **Repository: ostale tabele** |
| 5 | Prva 3–4 Repository-ja | **Query sloj (`modScr*` → 0 `TBL_`)** |
| 6 | Očisti `modDataAccess` | **Domain gde se isplati (storno, dokument, otkup)** |
| 7 | Razdvoj `modMasterSync` | **`modSyncControl` + `modBrojevi` split** |
| 8 | Pojednostavi numbering | **`modMasterSync` → Application** |
| 9 | Centralizuj `SyncControl` | **Imenovanje, uz postojeće izmene** |
| 10 | CI dependency rules | *(izvršeno u koraku 1)* |

Ključna inverzija: **CI ide sa 10. na 1. mesto, Repository sa 5. na 3.,
Application preimenovanje sa 2. na 9.**

Razlog: pravilo koje mašina nameće važi i za agenta i za čoveka od prvog dana.
Sloj koji se samo preimenuje ne sprečava nijedan bag.

---

## 6. Šta ovo NE rešava

Pošteno, da se ne prodaje više nego što daje:

- **Ne smanjuje `modTest.bas` (7943 LOC)** ni potrebu za Windows/Excel runnerom.
- **Ne dira `.frx`** — nove kontrole i dalje idu runtime-om.
- **Ne rešava schema drift** — šema tabela ostaje izvor istine po instalaciji
  (`CLAUDE.md` §3). Repository sloj čak *povećava* važnost te provere, jer
  centralizuje pretpostavke o kolonama na jedno mesto.
- **Ne ubrzava PWA sync** — samo mu daje jedan ulaz umesto direktnog upisa.

---

## 7. Preporuka

Prihvatiti ciljnu sliku, odbaciti redosled i fizičku taksonomiju.

Ako se radi samo jedna faza — **Faza 0**. Ona je najjeftinija, nema runtime
rizik, i menja ponašanje svake buduće sesije (ljudske i agentske) time što
prekršaj sloja pada po imenu umesto da se otkrije u code review-u.

Ako se radi druga — **Faza 2** (Repository), jer jedina obara broj koji danas
proizvodi bagove.

Fazu 1 (legacy duplikacija) tretirati kao **preduslov**, ne kao fazu: dok stoji,
svaka izmena poslovnog pravila plaća dvostruko.

---

## 8. Reprodukcija brojki

```bash
# pisači po tabeli
python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md

# Presentation coupling
grep -c 'TBL_[A-Z_]*'  src-vba/modScrDokumenti.bas   # 75
grep -c 'COL_[A-Z_]*'  src-vba/modScrDokumenti.bas   # 149
grep -cE '\b(AppendRow|UpdateCell)\b' src-vba/modScr*.bas src-vba/modOtkupUI.bas  # 0

# legacy duplikacija
grep -cE 'modOtkupUnos|OtkupUpisi'  src-vba/frmOtkup.frm       # 0
grep -cE 'modDokUnos|OtpremnicaUpisi' src-vba/frmDokumenta.frm # 0
grep -cE '[A-Za-z_]+_TX\b' src-vba/frmDokumenta.frm            # 36

# Application sloj koji "ne postoji"
wc -l src-vba/mod{Otkup,Dok,Novac,Agro}Unos.bas                # 2630
grep -hoE '^(Public|Private)? ?(Sub|Function) [A-Za-z_]+_TX' src-vba/*.bas | wc -l  # 85

# VBA ograničenja
grep -rn '^Implements ' src-vba/ | wc -l                       # 0
ls src-vba/*.cls | wc -l                                       # 14
grep -hoE '^Public (Sub|Function|Const|Type|Enum) +[A-Za-z_]+' src-vba/*.bas | wc -l  # 2222

# SyncControl pisci
grep -lic 'SyncControl' src-vba/*.bas                          # 2 fajla
```
