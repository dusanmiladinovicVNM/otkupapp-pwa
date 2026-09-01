# ADR-0003 — Repository granica: vlasništvo upisa, transakcioni scope, bootstrap izuzetak

- **Status:** Prihvaćeno (smer); sprovođenje nije počelo — prvi `modRepo*` modul još ne postoji.
- **Datum:** 2026-09-01
- **Kontekst grane:** `claude/agrix-vba-architecture-ooahib` (ocena plana slojevite arhitekture)
- **Gradi na:** ADR-0001 (nepromenljivost izdatih dokumenata), ADR-0002 (append-only)
- **Vezano:** `docs/Architecture/ARHITEKTURA_PLAN_OCENA.md` (plan v4, Faza 1),
  `docs/DOMEN/WHO_WRITES.md`, `.claude/rules/podaci-i-config.md`

## Kontekst

Plan v4 uvodi Repository sloj (`modRepo*`) kao **jedini fizički write gateway** po
tabeli. Premeravanje je pokazalo da je posao manji nego što se mislilo: od 21
poslovne tabele **18 već ima tačno jednog fizičkog pisca**, a rade se samo tri —
`tblOtkup` (4), `tblZbirna` (3), `tblKorisnici` (2).

Dve odluke moraju pasti **pre prvog `modRepo*` modula**, jer se posle njega
ugrađuju u API i menjaju se skupo:

1. Kako se tretiraju `modSetup` i `modMigracija`, koji pišu poslovne tabele ali
   nisu poslovni put.
2. Da li Repository sme da otvara/deklariše transakciju.

## Odluka

### A. `modSetup` i `modMigracija` su **eksplicitan, imenovan izuzetak**

Ne idu kroz Repository. Razlog: njihov upis nije poslovni događaj nego
**bootstrap i jednokratna migracija**.

Mereno stanje koje odluku opravdava:

| Modul | Fizički upis | Priroda |
|---|---|---|
| `modSetup` | `TBL_KORISNICI` ×9, `TBL_OTKUP` ×1 | kreiranje **admin naloga** na novom računaru (`ULOGA_ADMIN`); jedan idempotentan backfill `BrojOtpremnice` (`samo prazne`) |
| `modMigracija` | **0** kroz `modDataAccess` | jednokratna migracija iz starog fajla, radi **direktno nad Excel objektnim modelom** (bez `AppendRow`/`UpdateCell`) |

Da su ovi upisi išli kroz Repository, `modRepoOtkup` bi dobio operacije tipa
„popravi zatečeno" i „napravi prvog admina", koje nemaju veze sa poslovnim
jezikom domena i zagadile bi API koji ADR upravo pokušava da drži semantičkim.

**Granice izuzetka — izuzetak je uzak i imenovan, ne wildcard:**

- Važi **samo** za `modSetup` i `modMigracija`, kao **spisak imena** u `vba_check`
  pravilu, nikad kao obrazac `mod(Setup|Migracija|...)*` niti kao „infrastrukturni
  moduli smeju".
- Pokriva **samo** bootstrap, jednokratnu migraciju i idempotentan backfill.
- **Nov poslovni upis u `modSetup` i dalje pada** na `SLOJ` proveri — izuzetak se
  odnosi na modul, ali ne oprašta proširenje njegove uloge. Ako `modSetup` sutra
  dobije operaciju koja knjiži poslovni događaj, ona ide kroz Repository ili se
  seli iz `modSetup`.

### B. Repository **ne sme** da zove `BeginTx` ni `AddTableSnapshot`

Vlasništvo transakcionog scope-a ostaje **pozivaocu** (Application), ne
Repository-ju.

```
modAppOtkup / *Unos / *_TX          <- OTVARA transakciju i DEKLARISE tabele
    tx.BeginTx
    tx.AddTableSnapshot TBL_OTKUP
    tx.AddTableSnapshot TBL_AMBALAZA
        |
        +--> modRepoOtkup.Insert          <- SAMO pise; ne zna za tx
        +--> modRepoAmbalaza.Record
    tx.CommitTx
```

Razlog nije stil. `clsTransaction.RollbackTx` može da vrati **samo one tabele
koje su deklarisane snapshotom**. Ako Repository sam deklariše svoju tabelu, onda:

- operacija koja piše dve tabele dobija dva nezavisna scope-a umesto jednog, pa
  delimičan rollback postaje moguć;
- `clsTransaction.BeginTx` **puca na ugnežđenu transakciju**, pa bi Repository
  pozvan iz već otvorene transakcije rušio poziv;
- vlasništvo scope-a se seli u sloj koji ne zna poslovni redosled, a jedino
  Application zna koje tabele jedna operacija menja kao celinu.

**Ovo ne uvodi novo pravilo — kodifikuje zatečeno.** Mereno nad `src-vba/`
(produkcija, bez testova):

| | Broj |
|---|---|
| Procedura koje zovu `BeginTx` | 88 |
| ...od toga deklarišu `AddTableSnapshot` u **istoj** proceduri | **87** |
| `AddTableSnapshot` **bez** `BeginTx` u istoj proceduri | **0** |

Jedini izuzetak je `modSEFValidator.ValidateFakturaCanBeStorniranoOnSEF`, koji
`BeginTx` koristi **namerno kao sondu** — oslanja se na to da `BeginTx` puca na
ugnežđenu transakciju, da bi utvrdio da li je već u njoj. Nosi komentar koji to
objašnjava. To nije upis i ne narušava pravilo.

Dakle invarijanta „ko otvara transakciju, taj deklariše tabele" danas važi
**100%**. Odluka je štiti od jedinog sloja koji bi je prirodno prekršio.

## Posledice

**Prihvaćene, izgovorene naglas:**

- **`tblKorisnici` ostaje trajno na 2 fizička pisca** (`modAuth`, `modSetup`).
  Metrika Faze 1 se za tu tabelu ne zatvara na 1, i to nije propust nego odluka.
  Faza 1 se svodi na **`tblOtkup` i `tblZbirna`**.
- **Backfill `BrojOtpremnice` u `modSetup` ostaje van `modRepoOtkup`.** Ko traži
  „ko sve piše `COL_OTK_BROJ_OTPREMNICE`" mora da pogleda i taj put —
  `WHO_WRITES.md` ga prikazuje, pa je vidljiv.
- **`modRepo*.Insert` je pozivljiv samo iz otvorene transakcije.** Poziv van nje
  piše bez rollback zaštite. Ovo je **nova klasa greške koju uvodi sama odluka**
  i mora da dobije proveru (v. „Sledeći koraci"), inače se granica oslanja na
  disciplinu — a `modScrIzvestaji` je pokazao da disciplina ne drži uvek.

**Dobijeno:**

- Repository API ostaje semantički (`LinkToOtpremnica`, `MarkStornirano`), bez
  „popravi zatečeno" operacija — bez toga bi Faza 1 bila preimenovanje
  `UpdateCell`-a, što plan v4 označava kao M9.
- Transakcioni scope ostaje na jednom mestu po operaciji, pa `RollbackTx` i dalje
  vraća **sve** što je operacija dirala.

## Sledeći koraci

Ne blokiraju ovaj ADR; ulaze u `vba_check` uz Fazu 1:

1. **`SLOJ` izuzetak kao imenovan spisak** — `modSetup`, `modMigracija`, ništa
   drugo; obrazac je zabranjen.
2. **`REPO_TX`** — `modRepo*` ne sme sadržati `BeginTx` ni `AddTableSnapshot`.
   Trivijalna provera, prazan baseline dok `modRepo*` ne postoji.
3. **`REPO_POZIV`** — poziv `modRepo*.<upis>` mora biti u proceduri koja je
   otvorila transakciju ili je dokazivo pozvana iz nje. Ovo je teže od (2) i
   može krenuti kao upozorenje umesto greške.
4. **Obavezan „dokaz u oba smera"** za sve tri (`CLAUDE.md` §5) — menja se sam
   checker.

## Alternativa (odbijena)

**„Repository sam otvara transakciju kad je nema."** Zvuči udobno — poziv radi u
oba konteksta. Odbijeno jer:

- `clsTransaction.BeginTx` puca na ugnežđenu transakciju, pa bi Repository morao
  da **detektuje** da li je već u njoj; ta detekcija je danas sonda-sa-greškom
  (`modSEFValidator`), a ne stanje koje se čisto čita;
- operacija nad dve tabele bi dobila dva scope-a i delimičan rollback;
- sakriva od pozivaoca da je nešto transakciono, pa `AddTableSnapshot` spisak
  prestaje da bude čitljiv izvor „šta ova operacija menja" — a `who_writes.py`
  upravo od njega gradi signal `tx`.
