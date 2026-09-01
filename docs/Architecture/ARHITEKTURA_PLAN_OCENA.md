# Ocena plana „idealni AgriX VBA codebase" + korigovani plan

- **Status:** Analiza / predlog smera. NIJE implementirano.
- **Verzija:** v4 (2026-09-01) — premereno nad `origin/main` posle **109
  commita i +14.278 linija** u `src-vba/`. Tri nalaza iz v2/v3 **opovrgnuta
  merenjem**; `who_writes.py` popravljen. v2/v3 istorija u §9.
- **Datum:** 2026-08-26, revidiran 2026-09-01
- **Predmet:** predlog slojevite arhitekture (Host → Presentation → Application →
  Domain → Repository → Infrastructure) sa Sync-om kao izolovanim subsistemom.
- **Odluke izvedene iz ovog dokumenta:** `docs/adr/0003-repository-granica-i-izuzeci.md`
  (Repository granica: vlasništvo upisa, transakcioni scope, bootstrap izuzetak).
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

## 0a. Šta se promenilo od v3 (premereno 2026-09-01)

v2/v3 su merени nad `466df52`. `origin/main` je od tada otišao **109 commita**
napred, sa **+14.278 linija u `src-vba/`** i tri nova ekranska modula. Premereno
stanje **obara tri nalaza iz v3**.

### N1 — „12 fizičkih pisača nad `tblOtkup`" je bilo netačno. Pravi broj je 4.

v3 je tu brojku uzeo iz `WHO_WRITES.md`, ne shvatajući da ona sabira **dva
različita signala**: `tx` (`AddTableSnapshot` — „operacija menja ovu tabelu u
svojoj transakciji") i `direct` (`AppendRow`/`UpdateCell` — stvarni upis).
Za write-gateway metriku važi samo drugi.

Premereno, **fizički pisci (produkcija, oba oblika poziva):**

| Tabela | Fizičkih pisača | Moduli |
|---|---|---|
| `tblOtkup` | **4** | `modOtkup`, `modDokumenta`, `modMasterSync`, `modSetup`¹ |
| `tblZbirna` | **3** | `modDokumenta`, `modMasterSync`, `modDokumentInvariant` |
| `tblKorisnici` | **2** | `modAuth`, `modSetup` |
| **ostalih 18 tabela** | **1** | — |

¹ `modSetup` je „new PC setup / health-check"; upis je idempotentan backfill
(`samo prazne`), ne poslovni put.

**Posledica za plan:** `tblFakture`, `tblNovac`, `tblAmbalaza`, `tblOtpremnica`,
`tblPrijemnica` **već imaju tačno jedan fizički pisač**. Repository sloj je de
facto ~85% gotov. Faza „Repository" nije posao nad 5 tabela i 12 modula, nego
**nad 3 tabele i 6 modula**.

### N2 — `who_writes.py` je promašivao 43% mesta upisa. Popravljeno.

```python
# bilo -- trazi RAZMAK posle imena, pa ne vidi poziv sa zagradom
DIRECT_RE = re.compile(r'\b(?:AppendRow|UpdateCell)\s+(TBL_\w+|"(\w+)")', re.I)
```

`AppendRow` i `UpdateCell` su `Function`, pa se najčešće zovu `r = AppendRow(TBL_X, red)`.
`modMasterSync` ima **0** pogodaka na stari izraz, a stvarno radi
`AppendRow(TBL_OTKUP, rowData)` na liniji 1950.

| | Stari izraz | Ispravljen |
|---|---|---|
| mesta upisa | 32 | **57** (promašaj 43%) |
| `(modul, tabela)` parova | 18 | **37** |

Nevidljivih 19 parova uključivalo je `modMasterSync → TBL_OTKUP`,
`modFaktura → TBL_FAKTURE`, `modNovac → TBL_NOVAC`. Signal `direct` je bio
skoro prazan, pa je tabela u praksi prikazivala samo `tx`.

Najjasnija posledica: `tblSEFConfig` je u staroj tabeli stajao kao
**„0 pisača — samo testovi"**, a piše ga `modConfig`.

To je bag u **generisanom izvoru istine** na koji `CLAUDE.md` §2 izričito upućuje
pre izmene pravila upisa. Ispravljen u ovom commitu; `WHO_WRITES.md` regenerisan.

### N3 — Ciljno stanje Presentation sloja **već postoji u repou**, dvaput.

Tri nova ekrana od v3, pisana **bez ijednog `SLOJ` pravila**:

| Modul | LOC | `TBL_` | `COL_` | lookup | upis |
|---|---|---|---|---|---|
| `modScrSledljivost` | 1916 | **0** | **0** | **0** | 0 |
| `modScrBankaNalozi` | 1580 | **0** | **0** | **0** | 0 |
| `modScrIzvestaji` | 2903 | 51 | 33 | 17 | 0 |

Dva od tri su **savršeno čista**. Ne zato što ih je pravilo nateralo — nego zato
što delegiraju podatak poslovnom modulu:

```
modScrSledljivost --> modSledljivost      (TraceByZbirna, GetOtpremnicaKandidati...)
modScrBankaNalozi --> modBankaExportPregled (22 poziva), modNovac
             oba --> modOtkupUI  SAMO kao UI kit (ShowToast, GridCell, NewFieldG)
```

I — što je najvažnije za granicu iz §M-Query — **razdvajanje „šta" od „kako"
već je tačno postavljeno.** Kolonska specifikacija sa širinama živi u samom
ekranu (`modScrSledljivost.SlKoloneZaListu` → `"OTKUI_HD_DATUM||date|60|1"`), a
poslovni modul vraća samo podatak.

**Posledica za plan:** Faza „Query sloj" nije *uvođenje* novog sloja. To je
**„dovedi `modScrDokumenti` i `modScrIzvestaji` na obrazac koji `modScrSledljivost`
već koristi"** — sa dve radne referentne implementacije u repou.

### N4 — `SLOJ` baseline je 283 linije, a dva od tri pravila ga ne traže

v3 je tvrdio da baseline treba, ali ga nikad nije izmerio:

| Pravilo | Prekršaja danas | Baseline? |
|---|---|---|
| `modScr*`/`modOtkupUI` ne sme `AppendRow`/`UpdateCell`/`GetNextID` | **0** | **ne treba** |
| `modDataAccess` ne sme uzvodno (`modScr`/`modApp`/`modDom`/`frm`) | **0** | **ne treba** |
| `modScr*`/`modOtkupUI` ne sme `TBL_`/`COL_` | 283 u 7 modula | da |

Od 283, **264 je u tri fajla** (`modScrDokumenti` 166, `modScrIzvestaji` 52,
`modOtkupUI` 46).

Dakle dva pravila mogu **odmah, tvrdo, bez ikakvog baseline-a** — ona samo
zaključavaju ono što je već istina i sprečavaju povratak. To je najjeftiniji
posao u celom planu.

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

## 4. Plan v4

> v4 = v3 + premeravanje nad tekućim `main`. Tri izmene sledе direktno iz §0a:
> Repository se **smanjuje** (3 tabele, ne 5), Query se **prekvalifikuje** iz
> „uvedi sloj" u „primeni postojeći obrazac", a Faza 0 se **cepa** na deo bez
> baseline-a (odmah) i deo sa baseline-om.

### PR0 — `SyncControl` write ownership (hotfix, van plana)

Dve putanje pišu isti tab, jedna whole-tab replace-om (§1.5). Data-safety, ne
arhitektura. `modSyncControl` kao jedini vlasnik; `TryReadSyncControlAsDict` +
`ApplySyncControlUpdates` sele se iz `modStanicaLock` (već fail-closed).

*Metrika:* `grep -lic SyncControl src-vba/*.bas` → 1. *Rizik:* nizak. *~1 dan.*

### PR1 — `SLOJ`, tvrdi deo: bez baseline-a

Dva pravila imaju **nula prekršaja danas** (§0a/N4), pa idu odmah i tvrdo:

```
modScr*, modOtkupUI : zabranjeno AppendRow / UpdateCell / GetNextID
modDataAccess       : zabranjeno modScr* / modApp* / modDom* / frm*
```

Ne čiste ništa — **zaključavaju ono što je već istina.** Bez njih se stanje
održava disciplinom, a §0a/N3 pokazuje da disciplina ne drži uvek
(`modScrIzvestaji`).

Uz njih ide `REPO_API` (M9), takođe praznog baseline-a jer `modRepo*` još ne
postoji — a posle prvog takvog modula je znatno skuplje uvesti.

*Metrika:* tri provere zelene, baseline fajl **ne postoji**.
*Rizik:* nizak — ne dira runtime. *Obavezan „dokaz u oba smera" (`CLAUDE.md` §5).*

### PR2 — metrika vlasništva upisa

- `who_writes.py` — **urađeno u ovom commitu** (§0a/N2).
- `--max-writers N`, exit 2 iznad praga, uz `direct` signal (ne `tx`).

*Metrika:* prag postavljen na zatečeno (4), spušta se sa Fazom 1.

### FAZA 1 — Repository, ali samo gde stvarno treba

**Premereno: samo 3 tabele imaju >1 fizičkog pisača** (§0a/N1). Ostalih 18 su
već jednopisačke — tamo nema šta da se radi.

```
tblOtkup      4 -> 1    modOtkup, modDokumenta, modMasterSync, modSetup
tblZbirna     3 -> 1    modDokumenta, modMasterSync, modDokumentInvariant
tblKorisnici  2 -> 1    modAuth, modSetup
```

Dakle **`modRepoOtkup` i `modRepoZbirna`** — ne 20 repozitorijuma. Semantički
API (M9), TX-neutralan.

> **ODLUČENO — `ADR-0003`:** Repository **ne sme** `BeginTx` ni
> `AddTableSnapshot`; scope ostaje pozivaocu (Application). To ne uvodi novo
> pravilo nego kodifikuje zatečeno: od 88 procedura sa `BeginTx`, **87** deklariše
> snapshot u istoj proceduri, a `AddTableSnapshot` bez `BeginTx` ima **0**.
> (Jedini `BeginTx` bez snapshota je `modSEFValidator.ValidateFakturaCanBe-`
> `StorniranoOnSEF`, namerna sonda na ugnežđenu transakciju, ne upis.)
>
> **Nova klasa greške koju odluka uvodi:** `modRepo*.Insert` pozvan van
> transakcije piše bez rollback zaštite. Zato uz Fazu 1 idu provere `REPO_TX` i
> `REPO_POZIV` (ADR-0003, „Sledeći koraci").

> **ODLUČENO — `ADR-0003`:** `modSetup` i `modMigracija` su **imenovan izuzetak**
> (spisak, ne obrazac). Njihov upis je bootstrap admin naloga i jednokratna
> migracija, ne poslovni događaj; `modMigracija` uopšte ne piše kroz
> `modDataAccess`. Nov **poslovni** upis u `modSetup` i dalje pada na `SLOJ`.
>
> **Cena, prihvaćena:** `tblKorisnici` ostaje trajno na 2 fizička pisca, pa se
> Faza 1 svodi na **`tblOtkup` i `tblZbirna`**.

*Metrika:* sve poslovne tabele → 1 fizički pisač. *Rizik:* srednji.
*Obim: mnogo manji nego što je v3 procenio.*

### FAZA 2 — Ekranska disciplina (bivši „Query sloj")

Ne uvodi se sloj. **Primenjuje se obrazac koji `modScrSledljivost` i
`modScrBankaNalozi` već koriste** (§0a/N3) na dva ekrana koja su ostala:

```
modScrDokumenti  166 linija TBL_/COL_   -> podatak trazi od modDokumenta/modQry*
modScrIzvestaji   52 linija             -> od modIzvestaj
modOtkupUI        46 linija             -> KPI i lookup liste izlaze iz ljuske
```

To je **264 od 283 linije baseline-a**. Granica je već dokazana u repou:
poslovni modul vraća podatak, ekran drži kolonsku specifikaciju sa širinama.

*Metrika:* `TBL_`+`COL_` u `modScr*` → 0, u `modOtkupUI` → 0; **baseline fajl se
briše**, pravilo postaje tvrdo kao ona iz PR1. *Rizik:* nizak — read-only.

### FAZA 3 — Legacy konvergencija

Najveći runtime rizik, nepromenjen od v3 i **jedina faza koju v4 ne smanjuje**:
`frmOtkup` (1308) + `frmDokumenta` (6500) drže paralelnu kopiju pravila unosa,
sa **0 poziva** ka `modOtkupUnos`/`modDokUnos`.

Ide **paralelno** sa Fazama 1–2, ne pre i ne posle: mehanički dobici ne smeju da
čekaju visokorizičan posao.

*Metrika:* `*_TX` poziva iz `frmDokumenta.frm` pada sa 36.
*Rizik:* **visok** — traži `run_vba` na Windowsu; `vba_check` ovde ne dokazuje ništa.

### FAZA 4 — `TX_VRATA` i Domain

Tek sada, jer oba zavise od prethodnog:

- `TX_VRATA`: blizanac bez `_TX` mora biti `Private` ili pozivan iz transakcije.
  Danas 31 javnih vrata (§1.2). **Ne pre Faze 3** — legacy forme su deo slike.
- `modDomStorno`, `modDomDokument`, `modDomOtkup` — samo gde ima ROI (§v3).
  `modNovac`, `modFaktura`, `modDokumenta` se ne diraju bez povoda.

*Metrika:* Domain testovi bez sejanja tabela i bez `clsTransaction` (ne bez
workbook-a — M3).

### FAZA 5 — Sync adapter i numbering

Nepromenjeno od v3, i dalje nezavisno jedno od drugog:

- `modMasterSync`: `parse DTO` odvojiti od upisa; import zove `modOtkupUnos`.
  **Sada je jasniji ulog:** `modMasterSync` je jedan od 4 fizička pisača
  `tblOtkup` (§0a/N1), pa se ova faza i Faza 1 dodiruju — raditi ih u istom PR-u
  za `tblOtkup`.
- `modBrojevi` → `AllocateNewNumber` / `ObserveExistingNumber` + `NumberRegistry`.

### FAZA 6 — Imenovanje, oportunistički

`modOtkupUnos` → `modAppOtkup` samo uz izmenu koja ionako dira fajl. Sufiks `_TX`
ostaje — nosi informaciju na koju se veže `TX_VRATA`.

---

## 5. Redosled — četiri verzije

| v1 (originalni plan) | v2 | v3 | **v4 (premereno)** |
|---|---|---|---|
| — | — | HOTFIX SyncControl | **PR0 SyncControl** |
| Presentation granica | CI `SLOJ` | CI (3 provere) | **PR1 `SLOJ` tvrdi deo — bez baseline-a** |
| Uvedi `modApp*` | Legacy | legacy ↔ Repo | **PR2 `who_writes` popravka + prag** |
| `*Unos` iza App API-ja | `tblOtkup` 12→1 | Repo ostale | **F1 Repo: 3 tabele (ne 5)** |
| Izvuci Domain | Repo ostale | Query | **F2 Ekranska disciplina (264 linije)** |
| Prva 3–4 Repository-ja | Query | Domain | **F3 Legacy — paralelno sa F1/F2** |
| Očisti `modDataAccess` | Domain | Sync | **F4 `TX_VRATA` + Domain** |
| Razdvoj `modMasterSync` | SyncControl+brojevi | Numbering | **F5 Sync + numbering** |
| Numbering | `modMasterSync` | Imenovanje | **F6 Imenovanje** |
| SyncControl | Imenovanje | — | — |
| CI rules | — | — | — |

**Šta v4 menja u odnosu na v3:**

1. **Faza 0 se cepa.** Dva `SLOJ` pravila imaju prazan baseline i idu odmah,
   tvrdo (PR1). Treće čeka Fazu 2.
2. **Repository se smanjuje sa 5 tabela na 3**, jer je 18 tabela već
   jednopisačko — v3 je brojao pogrešan signal.
3. **Query se prekvalifikuje** iz „uvedi sloj" u „primeni obrazac koji dva
   ekrana već koriste".
4. **`TX_VRATA` se pomera iza legacy konvergencije** — legacy forme su deo te
   slike, pa bi baseline pre Faze 3 bio meren nad stanjem koje se menja.
5. **`who_writes.py` popravljen**, jer je metrika cele Faze 1 zavisila od
   signala koji je promašivao 43%.

---

## 5a. Sažetak: šta je AgriX-u zaista potrebno

> **AgriX-u ne treba još slojeva. Treba mu ownership nad slojevima koji već
> faktički postoje.**

Premereno stanje (2026-09-01) tu tezu **pojačava**, jer je AgriX bliže cilju nego
što su v2/v3 procenili:

| Sloj | Postoji? | Šta stvarno nedostaje |
|---|---|---|
| Presentation write separation | **da**, 0 upisa u svim `modScr*` | ništa — samo zaključati pravilom |
| Presentation read separation | **delimično** | 283 linije `TBL_`/`COL_`, 264 u 3 fajla |
| Query obrazac | **da**, 2 radna ekrana | primeniti na `modScrDokumenti`/`modScrIzvestaji` |
| Application-ish sloj | **da**, 2630 LOC + 33 operacije | koherentnost; 39 omotača meša mehanizam i nameru |
| Transaction boundaries | **da** | 31 javnih vrata, mašinski neprovereno |
| **Repository** | **18 od 21 tabele već** | **3 tabele: `tblOtkup`, `tblZbirna`, `tblKorisnici`** |
| Sync idempotency | **da** | jedan ulaz umesto direktnog upisa |
| Enforcement granica | **ne** | `vba_check` ih ne zna |
| Legacy jedinstvenost | **ne** | 7808 LOC paralelne kopije pravila |

Najveći preostali posao **nije Repository** — to je 3 tabele. Najveći je i dalje
**legacy duplikacija**, jedina stavka koju nijedno premeravanje nije smanjilo.

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

**PR0 `SyncControl`** ide odmah i van plana — data-safety, ~1 dan.

Ako se radi samo jedan PR — **PR1**. Dva `SLOJ` pravila imaju **nula prekršaja
danas**: ne čiste ništa, nego zaključavaju stanje koje je već postignuto, po ceni
od jedne `check_` funkcije. `modScrIzvestaji` (§0a/N3) je dokaz da se bez
pravila stanje vraća unazad i u novom kodu.

Ako se radi drugi — **F2 (ekranska disciplina)**, jer je 264 od 283 linije u tri
fajla, obrazac je već dokazan u repou dvaput, i rizik je nizak.

**Repository (F1) je pao u prioritetu** posle premeravanja: 18 od 21 tabele je
već jednopisačko, ostaju tri. To je i dalje vredan posao, ali nije ono što danas
proizvodi najviše rizika.

Ono što se **ne sme** raditi prvo, nepromenjeno kroz sve četiri verzije: Domain
extraction. Dok `frmOtkup` i `frmDokumenta` drže paralelnu kopiju pravila, svaki
izvučen Domain invariant važi za jedan od dva puta.

---

## 9. Istorija verzija

| Verzija | Šta je donela | Šta je od nje opovrgnuto |
|---|---|---|
| **v1** | originalni plan: Host→Presentation→Application→Domain→Repository→Infra | dijagnoza trenutnog stanja (M1–M9); redosled; folder taksonomija |
| **v2** | prvo premeravanje: Application postoji, UI ne piše, 12 pisača, legacy duplikacija | „svih 85 `_TX` je Application"; „jedan pisač po tabeli"; obrazloženje za `modRepo*` |
| **v3** | korekcije autora + `_TX` klasifikacija (39/72 omotača), M9 semantički API, preplitanje faza | „12 fizičkih pisača" (bio pogrešan signal); veličina Repository faze; „uvedi Query sloj" |
| **v4** | premereno nad `main` +109 commita: Repo 3 tabele, Query obrazac već postoji ×2, baseline 283 linije, `who_writes.py` popravljen | — |

Metod je kroz sve četiri isti i vredi ga zadržati: **tvrdnja koja nije izmerena
se ne upisuje kao nalaz.** Tri puta je merenje oborilo zaključak koji je zvučao
tačno — jednom u v1, jednom u v2, jednom u v3.

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

# v4: fizicki pisci po tabeli (OBA oblika poziva -- ovo je metrika Faze 1)
grep -rhoE '\b(AppendRow|UpdateCell)\s*\(?\s*TBL_\w+' src-vba/*.bas \
  | grep -oE 'TBL_\w+' | sort | uniq -c | sort -rn        # tblOtkup: 4 modula
python3 tools/who_writes.py --out docs/DOMEN/WHO_WRITES.md   # posle popravke N2

# v4: ekrani na tekucem main
for f in src-vba/modScr*.bas src-vba/modOtkupUI.bas; do
  echo "$(basename $f) TBL_=$(grep -c 'TBL_[A-Z_]' $f) COL_=$(grep -c 'COL_[A-Z_]' $f)"
done                        # modScrSledljivost / modScrBankaNalozi = 0/0

# v4: SLOJ baseline
grep -c -E '(AppendRow|UpdateCell|GetNextID)' src-vba/modScr*.bas   # sve 0
grep -hc -E '(TBL_|COL_)[A-Z_]+' src-vba/modScr*.bas src-vba/modOtkupUI.bas \
  | paste -sd+ | bc                                                 # 283
```
