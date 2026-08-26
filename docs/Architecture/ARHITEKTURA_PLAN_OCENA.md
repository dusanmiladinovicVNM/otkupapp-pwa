# Ocena plana „idealni AgriX VBA codebase" + korigovani plan

- **Status:** Analiza / predlog smera. NIJE implementirano.
- **Verzija:** v2 (2026-08-26) — ugradjene korekcije autora plana; `_TX`
  klasifikacija izmerena, ne pretpostavljena; plan prepravljen u v3 redosled.
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

#### `*_TX` — mereno, ne pretpostavljeno

U v1 ovog dokumenta stajalo je da je svih 85 `*_TX` procedura Application sloj.
**To je bilo preširoko.** `_TX` označava *vlasništvo nad transaction boundary-jem*,
ne *poslovni use-case*. Merenje:

| | Broj |
|---|---|
| Ukupno `*_TX` | 85 |
| ...od toga `Test*_TX` (test helperi, ne produkcija) | 13 |
| **Produkcionih** | **72** |
| ...od toga **transakcioni omotač** oko blizanca bez `_TX` | **39 (54%)** |
| ...od toga samostalna operacija | 33 |

Tipičan omotač (`modOtkup.SaveOtkup_TX`) je doslovno:

```vba
tx.BeginTx
tx.AddTableSnapshot TBL_OTKUP
tx.AddTableSnapshot TBL_AMBALAZA
SaveOtkup_TX = SaveOtkup(...)          ' <- ceo posao je ovde
If SaveOtkup_TX = "" Then Err.Raise ...
tx.CommitTx
```

Dakle **54% produkcionih `_TX` je mehanizam, ne namera.** Application sloj
postoji, ali **nije koherentan niti eksplicitno ograničen** — što je bitno
drugačija tvrdnja od „ne postoji" (v1) i od „svih 85 je Application" (takođe v1).

Uvođenje `modApp*` pored ovoga i dalje daje **treće imenovanje istog sloja**.

#### Nuspojava: 31 javna vrata pored transakcije

Od 39 omotača, blizanac bez `_TX` je **`Public` u 31 slučaju** (`SaveOtkup`,
`SaveZbirna`, `StornoOtkup`, `ApplyAvansToOtkup`, …), a `Private` u 8. Znači da
je poslovni upis pozivljiv i **mimo** snapshot/rollback puta.

> **Provereno, i nalaz je uži nego što je prvo izgledalo:** stvarni pozivaoci
> non-TX blizanaca su **poslovni moduli** (`modDokumenta`, `modStorno`,
> `modNovac`, `modBankaMapiranje`) koji komponuju unutar *spoljne* transakcije —
> to je legitiman obrazac.
> **Nijedan `modScr*` ni `frmOtkupUI` ne zaobilazi `_TX`.** Prvi `grep` je to
> naizgled pokazao, ali su pogoci bili (a) `modScrDokumenti` koji ima **sopstvene
> `Private` funkcije istog imena** (`SaveOtpremnica`, `SaveZbirna`,
> `SavePrijemnica` — VBA name shadowing), i (b) komentari.

Dakle: **nije demonstriran bag.** Nalaz je da je transakciona granica
**konvencija sa 31 otvorenih vrata, koja danas niko ne koristi pogrešno, ali to
nije mašinski provereno.** Po `CLAUDE.md` §2 to se prijavljuje kao provera koja
nedostaje, ne kao ispravka koja se gura.

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

Netačno (§1.2). Postoji 2630 LOC u `*Unos` modulima + 33 samostalne `*_TX`
operacije.

**Precizna formulacija (ispravka v1):** nije da Application sloj *nedostaje* —
on **postoji ali nije koherentan ni ograničen**. Dva imenovanja (`*Unos` i
`*_TX`), od kojih drugo u 54% slučajeva označava samo transakciju.

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

**Precizacija cilja (ispravka v1):** cilj **nije** „jedan poslovni pisač po
tabeli". Poslovnih pisača i dalje treba da bude više — `modStorno` stornira,
`modMasterSync` uvozi, `modDokumenta` vezuje. Cilj je **jedan fizički write
gateway po tabeli**:

```
modOtkup      -> RepoOtkup.Insert
modStorno     -> RepoOtkup.MarkStornirano
modMasterSync -> RepoOtkup.InsertImported
modDokumenta  -> RepoOtkup.SetOtpremnica
                        |
                        v
                  modRepoOtkup          <- JEDINI koji sme AppendRow/UpdateCell
                        |                  nad TBL_OTKUP
                        v
                  modDataAccess
```

Merljiva metrika je zato **broj modula koji zovu `AppendRow`/`UpdateCell` nad
datom tabelom**, ne broj modula koji tabelu poslovno menjaju. `who_writes.py`
već meri prvo (signal `direct`), pa metrika ne traži nov alat.

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

**Ispravka:** `modRepo*` **standardni moduli**, ne klase.

**Ali obrazloženje iz v1 je bilo pola netačno i ispravlja se.** v1 je tvrdio da
standardni modul daje „isti fake seam" kao klasa. **Ne daje.** Standardni modul
nije objekat — `Set repo = fakeRepo` ne postoji, pa se ne može injektovati.

Tačno obrazloženje je drugo: **fake seam nam trenutno nije dovoljno vredan da
opravda dodatnu kompleksnost.** Standardni modul daje ono zbog čega Repository
sloj i uvodimo — *ekskluzivno vlasništvo nad `TBL_`/`COL_` i upisom*, mašinski
proverivo — po ceni od nula lifetime managementa.

Ako DI jednog dana zaista zatreba, put postoji i nije zatvoren:

```
IRepoOtkup  (Implements)
     |
     +-- clsExcelRepoOtkup
     +-- clsFakeRepoOtkup
```

Za 3–4 ključna aggregate-a, ne za 30. Uslov za taj korak je **demonstrirana
potreba** (Application test koji se drugačije ne može napisati), ne estetika.

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

### M9 — Repository može da se izrodi u „lepši DataAccess"

*(Nalaz autora plana, prihvaćen — najozbiljnija zamerka korigovanom planu.)*

Ako `modRepo*` dobije generički API, ceo posao je preimenovanje:

```vba
' LOSE -- ovo je UpdateCell sa prefiksom
RepoOtkup.UpdateColumn otkupID, COL_OTK_OTPREMNICA_ID, otpID
RepoOtkup.SetField otkupID, "BrojZbirne", broj
```

```vba
' DOBRO -- namera, ne mehanika
RepoOtkup.LinkToOtpremnica otkupID, otpID
RepoOtkup.AssignZbirna otkupID, brojZbirne
RepoOtkup.MarkStornirano otkupID, razlog
```

Razlika nije stilska. Kod generičkog API-ja invarijanta „kada se postavlja
`BrojZbirne`, mora se proveriti owner" **nema gde da živi** — ostaje razbacana
po 12 pozivalaca, tačno kao danas. Kod semantičkog API-ja ima tačno jedno mesto.

**Ovo je mašinski proverivo**, pa ne mora ostati dobra namera. Pravilo za
`vba_check`, uz `SLOJ`:

```
modRepo*: Public procedura NE SME
  - da ima parametar imena colName / columnName / fieldName
  - da joj ime sadrzi Column | Cell | Field | SetValue | UpdateRow
```

Bez ovog pravila je Faza „Repository" najskuplji no-op u planu.

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

## 4. Plan v3

> **v3 = v1 (originalni plan) + v2 (ova ocena) + korekcije autora plana.**
> Ključne izmene u odnosu na v2: `SyncControl` izlazi iz faza kao hotfix, legacy
> konvergencija i Repository se **prepliću** umesto da se nižu, numbering se
> odvaja kao nezavisan subsistem.

Princip: **prvo ono što se nameće mašinski i obara merljivu metriku; imenovanje
poslednje.**

### HOTFIX — `SyncControl` write ownership

*Nije faza.* Dve putanje pišu isti tab, jedna whole-tab replace-om (§1.5). To je
**data-safety bug**, ne arhitektura, i nema razloga da čeka Fazu 5.

- `modSyncControl` — jedini vlasnik taba. Preseliti `TryReadSyncControlAsDict` +
  `ApplySyncControlUpdates` iz `modStanicaLock` (već fail-closed), prevesti
  `modGoogleSyncOrchestrator` sa sopstvenog `WriteSheetData`.

**Metrika:** `grep -lic SyncControl src-vba/*.bas` → 1 fajl.
**Rizik:** nizak. **Obim:** ~1 dan.

### FAZA 0 — Instrumentacija

Bez izmene poslovne logike. Cilj: da svaka sledeća faza ima crveno/zeleno.

1. `vba_check.py` → nova provera **`SLOJ`**, tabelarno:

   ```
   modScr*, modOtkupUI : zabranjeno AppendRow, UpdateCell, GetNextID
   modDom*             : zabranjeno TBL_, COL_, Range, Worksheet, ListObject,
                         MsgBox, WinHttp, modDataAccess, modGoogle*
   modRepo*            : zabranjeno frm, modScr, modApp
   modDataAccess       : zabranjeno modScr, frm, modApp, modDom
   ```

   Uvesti sa **baseline fajlom postojećih prekršaja**: stari prolaze, **svaki nov
   pada**. Bez baseline-a se pravilo ne može uvesti nad 152k linija.

2. `vba_check.py` → provera **`REPO_API`** (M9): `modRepo*` ne sme generički
   API. Uvodi se **pre** prvog `modRepo*` modula, dok je baseline prazan.

3. `vba_check.py` → provera **`TX_VRATA`**: blizanac bez `_TX` mora biti
   `Private`, ili njegov pozivalac mora biti unutar transakcije. Danas 31 javnih
   vrata (§1.2). Baseline zamrzava zatečeno; nova vrata padaju.

4. Self-test za sve tri, po postojećem obrascu „dokaz u oba smera" — `CLAUDE.md`
   §5 to izričito traži kad se menja sam checker.

5. `who_writes.py` → `--max-writers N`, exit 2 iznad praga. Prag se spušta fazama.

**Metrika:** tri nove provere zelene sa zamrznutim baseline-om.
**Rizik:** nizak — ne dira runtime. **Ovo je jedina faza koja se isplati čak i
ako se ostatak plana nikad ne uradi.**

### FAZE 1A / 1B — prepliću se, ne nižu

*(Korekcija v2, koji ih je nizao 1 → 2.)* Legacy konvergencija je **visok runtime
rizik**; Repository Otkup je **mehanički**. Nema razloga da mehanički dobitak
čeka rizičan posao. Dve linije rada idu paralelno, po jedan PR:

```
PR1   vba_check: SLOJ + REPO_API + TX_VRATA + baseline
PR2   modRepoOtkup: Insert / LinkToOtpremnica / AssignZbirna / MarkStornirano
PR3   modOtkup        -> repo
PR4   modMasterSync   -> repo
PR5   frmOtkup: jedan vec migriran use-case -> modOtkupUnos
PR6   modStorno       -> repo
PR7   frmDokumenta: sledeci migriran use-case -> modDokUnos
...
```

**1A — legacy konvergencija.** Samo za use-case koje `*Unos` moduli **već**
pokrivaju. Ne dirati `.frx`; ne dodavati `WithEvents` (`CLAUDE.md` §3). Kad režim
ima jedan put — obrisati mrtvu kopiju iz forme.

*Metrika:* `grep -c 'modDokUnos' frmDokumenta.frm` > 0; broj `*_TX` poziva iz
`frmDokumenta.frm` pada sa 36. *Rizik:* **visok** — traži `run_vba` na Windowsu,
`vba_check` ovde ne dokazuje ništa.

**1B — `modRepoOtkup`.** Jedini fizički write gateway za `TBL_OTKUP`.
Semantički API (M9). **TX-neutralan** — ne otvara sopstvenu transakciju.

*Metrika:* moduli koji zovu `AppendRow`/`UpdateCell` nad `TBL_OTKUP`: 12 → 1.
*Rizik:* srednji — `modStorno`/`modStornoFlow` pišu iz transakcionog konteksta.

> **Odluka koja mora pasti pre PR2:** sme li Repository da zove
> `clsTransaction.AddTableSnapshot`?
> **Preporuka: ne.** Ostaje pozivaocu, inače Repository postaje vlasnik
> transaction scope-a — a to je Application posao (originalni plan §7 to
> ispravno kaže).

### FAZA 2 — Repository za ostale tabele

`tblFakture` (9), `tblNovac` (7), `tblAmbalaza` (6), `tblZbirna` (5). Isti
obrazac, isti semantički API, prag u `who_writes.py` se spušta posle svake.

**Metrika:** sve poslovne tabele ≤ 2 fizička write gateway-a.

### FAZA 3 — Query sloj

Jedini preostali Presentation dug (§1.1): UI ne menja bazu, ali **zna previše o
njenoj strukturi**.

- `modQryDokumenti` — vraća **read model**, ne formatiran grid.
- `modQryOtkup` — KPI agregacije koje danas žive u `modOtkupUI`
  (`SumKgForDate`, `CountForDate`) + lookup liste (`FillComboDisplayID`).

> **Granica koju ne treba preći** *(nalaz autora plana, prihvaćen):* Query vraća
> **podatak**, ne izgled. `DocumentListRow{Datum, Broj, Partner, Kolicina,
> Status}` — a ne širinu kolone, bold, pill ili redosled. Inače se coupling samo
> premesti sa baze na grid, i `modQry*` postane drugi `modScr*`.
>
> ```
> modQryDokumenti  -> read model   (sta)
> modScrDokumenti  -> formatiranje (kako)
> ```

**Metrika:** `TBL_`+`COL_` u `modScr*` → 0; u `modOtkupUI` → 0. Skinuti `SLOJ`
baseline za `modScr*`. *Rizik:* nizak — read-only, greška se vidi na ekranu.

### FAZA 4 — Domain gde ima ROI (ne svuda)

Ne izvlačiti Domain iz svih 8 poslovnih modula:

- `modDomStorno` — iz `modStornoImpact` (500 LOC, već 0 Excel objekata) i
  `modStorno`: `CanStorno`, `ValidateOwnership`, `ValidateGeneration`,
  `ValidateCascade`.
- `modDomDokument` — iz `modDokumentInvariant` (670 LOC, 1 Excel referenca).
- `modDomOtkup` — bruto→neto, klase, ambalaža, cena.

`modNovac`, `modFaktura`, `modDokumenta` **ostaviti** dok ne postoji povod.
`modDokumenta` (4079 LOC, 290 `TBL_`) je najskuplji za razdvajanje i najmanje se
menja.

**Metrika:** `SLOJ` za `modDom*` prolazi bez baseline izuzetka; Domain testovi
rade **bez sejanja tabela i bez `clsTransaction`** (§M3 — ne bez workbook-a).

### FAZA 5 — Sync adapter → postojeći Application put

`modMasterSync` (4065 LOC): razdvojiti `parse DTO` od `ImportRowToTblOtkup`.
Import zove `modOtkupUnos`/`modAppOtkup` umesto da sam piše. `TestHook_*`
seam-ovi ostaju.

**Metrika:** `modMasterSync` ne zove `AppendRow` nad poslovnim tabelama.
**Rizik:** visok — sync ima fail-closed ponašanje koje se ne sme oslabiti.

### FAZA 6 — Numbering, nezavisno

*Ne meša se sa Repository refaktorom.* `modBrojevi` (626 LOC) danas radi pet
strategija u jednom modulu (`MaxSeqFromTable`, `MaxSeqFromGoogleSheet`,
`BrojZbirneExists`, mirror prefiks, cache).

```
AllocateNewNumber(...)      -- cloud namespace: NumberRegistry, atomic increment
                               offline centralni VBA: lokalni counter
ObserveExistingNumber(...)  -- naknadni papirni unos: LastSeq = max(LastSeq, n)
```

`NumberRegistry`: `Kind | EntityID | BusinessDate | LastSeq`. Može kad god,
nezavisno od ostalih faza.

### FAZA 7 — Imenovanje, oportunistički

`modOtkupUnos` → `modAppOtkup`, **samo uz izmenu koja ionako dira fajl**. Nikad
kao zaseban rename commit preko 200 modula.

Sufiks `_TX` **zadržati** — nosi informaciju („ova procedura otvara transakciju")
koju `modApp` prefiks ne nosi, i na koju se veže `TX_VRATA` provera.

### Odbačeno iz originalnog plana

| Stavka | Razlog |
|---|---|
| `00_Host/` … `90_Tests/` folderi | VBA namespace je ravan; prefiks + `vba_check` daju isto (M5) |
| `clsOtkupRepository` i sl. | fake seam koji §20 traži ne dobija se ni klasom bez `Implements`, a sam seam trenutno ne vredi dodatnu kompleksnost (M4) |
| `clsKreirajOtkupCmd` command klase | `Object`/`Dictionary` payload koji `modOtkupUnos` već koristi radi isto; typed command tek kad se dokaže problem (npr. `kooperantID` / `KooperantId` / `koopID` drift) |
| „Domain tests — no workbook" | fizički neizvodljivo u VBA (M3); cilj je bez fixture-a |
| Novi `modApp*` pored `*Unos` | treće ime za isti sloj (M1) |

---

## 5. Redosled — tri verzije

| # | v1 (originalni plan) | v2 (prva ocena) | **v3 (usaglašeno)** |
|---|---|---|---|
| — | — | — | **HOTFIX: `SyncControl`** |
| 1 | Presentation granica | CI `SLOJ` | **CI: `SLOJ` + `REPO_API` + `TX_VRATA`** |
| 2 | Uvedi `modApp*` | Legacy duplikacija | **1A legacy ↔ 1B `modRepoOtkup` (preplitanje)** |
| 3 | `*Unos` iza App API-ja | `tblOtkup` 12 → 1 | **Repo: Fakture / Novac / Ambalaza / Zbirna** |
| 4 | Izvuci Domain | Repo ostale tabele | **Query sloj (`modScr*` → 0 `TBL_`)** |
| 5 | Prva 3–4 Repository-ja | Query sloj | **Domain: Storno, Dokument, Otkup** |
| 6 | Očisti `modDataAccess` | Domain gde se isplati | **Sync adapter → Application** |
| 7 | Razdvoj `modMasterSync` | `modSyncControl` + brojevi | **Numbering (nezavisno)** |
| 8 | Pojednostavi numbering | `modMasterSync` → App | **Imenovanje, oportunistički** |
| 9 | Centralizuj `SyncControl` | Imenovanje | — |
| 10 | CI dependency rules | — | — |

Šta se promenilo od v2: **`SyncControl` izlazi iz faza** (bugfix, ne
arhitektura), **legacy i Repository se prepliću** umesto da se nižu, **numbering
postaje nezavisan**, i **dve nove CI provere** (`REPO_API`, `TX_VRATA`) ulaze u
Fazu 0 — jer se obe uvode jeftino dok je baseline prazan, a skupo posle.

---

## 5a. Sažetak: šta je AgriX-u zaista potrebno

Najkraća formulacija, i ona koju vredi zapamtiti umesto celog dokumenta:

> **AgriX-u ne treba još slojeva. Treba mu ownership nad slojevima koji već
> faktički postoje.**

Izmereno stanje u prilog tome:

| Sloj | Postoji? | Šta nedostaje |
|---|---|---|
| Presentation write separation | **da** | read coupling (`TBL_`/`COL_` u `modScr*`) |
| Application-ish sloj | **da**, 2630 LOC + 33 operacije | koherentnost i granica; 39 omotača mešaju mehanizam sa namerom |
| Transaction boundaries | **da**, `clsTransaction` + 72 `_TX` | 31 javnih vrata pored granice, mašinski neprovereno |
| Sync idempotency | **da** (`ClientRecordID`, row-TX, fail-closed) | jedan ulaz umesto direktnog upisa |
| Storno dekompozicija | **da**, 6 modula | pure Domain jezgro |
| **Repository** | **ne** | **12 fizičkih pisača nad `tblOtkup`** |
| **Query** | **ne** | 149 `COL_` referenci u jednom ekranu |
| Enforcement granica | **ne** | `vba_check` ih ne zna |

Dve prazne linije u koloni „postoji" (Repository, Query) + neenforce-ovane
granice = ceo posao. Sve ostalo je ownership i imenovanje.

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

**HOTFIX `SyncControl`** ide odmah i van plana — data-safety, ~1 dan.

Ako se radi samo jedna faza — **Faza 0**. Najjeftinija, bez runtime rizika, i
menja ponašanje svake buduće sesije (ljudske i agentske): prekršaj sloja pada po
imenu umesto da se otkrije u code review-u. `REPO_API` i `TX_VRATA` moraju ući
tu, dok im je baseline prazan — posle prvog `modRepo*` modula su znatno skuplje.

Ako se radi druga — **1B (`modRepoOtkup`)**, jer jedina obara broj koji danas
proizvodi bagove, uz semantički API (M9) bez kojeg je ceo posao preimenovanje.

**1A (legacy duplikacija) ne blokira 1B.** Tretirati ih kao dve paralelne linije
PR-ova; legacy je visok runtime rizik i ne sme da drži mehanički dobitak.

Ono što se **ne sme** raditi prvo: Domain extraction. Dok legacy forme drže
paralelnu kopiju pravila, svaki izvučen Domain invariant važi za jedan od dva
puta.

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

# _TX klasifikacija: omotac vs samostalna operacija
for p in $(grep -hoE '^(Public |Private )?(Sub|Function) +[A-Za-z_]+_TX' src-vba/*.bas \
           | awk '{print $NF}' | grep -v '^Test' | sort -u); do
  base=${p%_TX}
  grep -qhE "^(Public |Private )?(Sub|Function) +${base}\b" src-vba/*.bas \
    && echo "omotac: $p"
done | wc -l                                                   # 39 od 72

# VBA ograničenja
grep -rn '^Implements ' src-vba/ | wc -l                       # 0
ls src-vba/*.cls | wc -l                                       # 14
grep -hoE '^Public (Sub|Function|Const|Type|Enum) +[A-Za-z_]+' src-vba/*.bas | wc -l  # 2222

# SyncControl pisci
grep -lic 'SyncControl' src-vba/*.bas                          # 2 fajla
```
