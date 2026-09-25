# Refaktor: imenovani upis umesto pozicionog

> **Ulaz za svaku sesiju koja radi na ovom refaktoru.** Cilj: nijedan pisac ne zna poziciju kolone.
> Kanon (`schema/schema.json`) je registar **imena**; živa sveska je registar **pozicija**; pisac ne
> vidi ni jedno ni drugo nego predaje imenovane vrednosti.
>
> Merenje je na `2cd86eb8` (merge PR #390). Ponašanje u Excelu je **NEVERIFIKOVANO** — `run_vba`
> traži Windows + Excel + `pywin32` i u web sesiji se ne izvršava.

---

## 0. Ispravka merenja

Ranija procena „161–163 mesta upisa" je bila **pogrešna**. To je bio broj pojavljivanja tokena
`AppendRow` u `src-vba/`, što uključuje 62 komentarske linije, dve definicije i `AppendRow = 0` /
`AppendRow = newRow.index` dodele u samom telu.

Stvaran broj, merеn parserom koji spaja prelomljene redove i izbacuje komentare, string literale i
definicione redove:

| Mera | Vrednost |
|---|---|
| Poziva `AppendRow` (+ lokalni `PalAppendRow`) | **51** |
| — produkcionih | **36** u 19 modula |
| — u test modulima | 15 |
| Mesta koja grade goli `Array(...)` | 16 |
| **Od toga niz kraći od broja kolona** | **14** (11 produkcionih, 3 testna) |
| Mesta gde je niz duži od broja kolona | 0 |
| Kapija arnosti | **0** |

Obim posla je dakle **36 produkcionih mesta**, ne 163. To menja rez: ovo je posao za 5–6 PR-ova,
ne za desetine.

---

## 1. Pre-flight verdikt

Po `.claude/skills/pre-flight/SKILL.md` §1. Dokaz uz svaku osu.

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | **GAP** (lokalizovan) | Dva pisca se oslanjaju na nezapisanu konvenciju „prazno znači X": `modAmbalaza.bas:175-184` šalje 10 vrednosti u `tblAmbalaza` (16 kolona) i `Stornirano` nikad ne upisuje; `modFaktura.bas:299` šalje 21 u `tblFakture` (30) i `IzdatoStatus` ostavlja prazan. Imenovani upis prisiljava odluku: upisati eksplicitno ili izostaviti ključ. **Blokira samo PR-ove B2 i B4**, ne ceo plan. |
| `IDENTITY` | `N/A` | Nijedan identitet se ne menja. Ni jedan `*ID` ne dobija nov izvor. Napomena: prvi argument mora ostati `TBL_` konstanta (atribucija, ne identitet — v. `WRITERS`). |
| `CARDINALITY` | `N/A` | Jedan niz = jedan red, pre i posle. Nijedan pisac ne prelazi iz 1:1 u 1:N. |
| `INVARIANTS/OWNER` | **PROVEN** | `python tools/who_writes.py --check-ownership` → exit 0, 34 tabele, „nema novih pisaca". Vlasnik upisa ostaje isti modul; menja se samo kako gradi red. |
| `WRITERS` | **PROVEN** (uz imenovanu zamku) | 36 produkcionih mesta popisano po fajlu (§0). **Zamka:** `tools/who_writes.py:81` `MUTATE_RE = r'\w*(?:AppendRow\|UpdateCell\|DeleteRow)\s*[\s(]\s*(TBL_\w+\|"(\w+)")'` — prefiks je slobodan, **sufiks strog**. Ime koje se *nastavlja* posle `AppendRow` (`AppendRowByName`, `AppendRowNamed`) kapiji je **nevidljivo**, pa bi A11 ostao zelen ne mereći ništa. Zato ime **mora da se završava** na `AppendRow`. |
| `DOWNSTREAM` | **PROVEN** | **Nijedno mesto u `src-vba/` ne čita tabelarni red literalnim indeksom kolone** — svi čitaoci idu po imenu (`RequireColumnIndex` 1368 pojava, `GetColumnIndex` 1057). Zato je kvar nevidljiv: `modSEFMapper.bas:37-41` čita `FakturaID/KupacID/BrojFakture/Datum/Iznos` po imenu i **verno pošalje pomerenu vrednost u SEF** — van sveske, u državni sistem. Drugi potrošač istog niza: `modJournaling.WriteJournalRow` (v. §2.4). Treći pozicioni put van sveske: `modGoogleSheets.AppendRowToSheet` (`:1056`, 3 produkciona pozivaoca). |
| `CAPABILITY` | `N/A` za zamenu upisa | Nijedna sposobnost se ne dodaje ni uklanja; `docs/DOMEN/MAPA_SPOSOBNOSTI.md` se ne dira. **Ne važi za fazu D3** (brisanje mrtve površine) — tamo je `CAPABILITY` posebna osa. |
| `ACCEPTANCE CONTRACT` | **PROVEN** (plan dokaza) | §5. |
| `PLATFORM` | **UNMEASURED** | Ne zna se šta `lo.ListRows.Add` ostavlja u kolonama koje petlja ne popuni — `ListRows.Add` nasleđuje formule i autofill tabele, pa „kraći niz ostavlja prazno" **nije dokazano**. Traži sondu na Windows + pravom Excelu. Do tada se piše kao neizmereno. |
| `LANDING` | **PROVEN** | Grana `claude/great-hawking-91ti26` = `origin/main` = `2cd86eb8`, radno stablo čisto. **PR #391 je otvoren nad S5** — `modDataAccess`/`schema.json` ne smeju paralelno s njim (`CLAUDE.md` §6, serijski redosled za schema/setup i centralne primitive). `tools/` sme sa kodom; `.claude/` ide **zasebnim process PR-om**. |

### BUSINESS EVENTS

`EVENTS: N/A — čist refaktor sloja upisa.` Nijedan datum, status ni tabela poslovnog događaja se ne
menja. Izuzetak je `DOMAIN GAP` iznad: odluka o `Stornirano` i `IzdatoStatus` **jeste** semantika
statusa i rešava se kao odluka operatera pre PR-ova B2/B4, ne u kodu.

### ⚠ Signali

- **⚠ DOMAIN NOT CLOSED** — dve konvencije „prazno znači X" (v. `DOMAIN`).
- **⚠ DOWNSTREAM RISK** — pomerena kolona danas putuje u SEF i u Google Sheets bez ijedne kapije.
- **⚠ PLATFORM UNKNOWN** — ponašanje `ListRows.Add` nad nepopunjenim kolonama nije izmereno.
- **⚠ FALSE-GREEN RISK** — kapija koja se prijavljuje kao `AppendRow = 0` je **nemerljiva** na 11
  poziva koji povratnu vrednost ne čitaju; u `dokaz.py` bi ispala kao „ne obara ništa", nerazlučivo
  od mrtve sabotaže. **Zato kapija mora da diže `Err.Raise`, ne da vraća 0.**

---

## 2. Izabran projekat

Tri nezavisna predloga su merena; izabran je **koegzistentni imenovani upis nad postojećim
`modSchemaGuard` obrascem**, uz presađene delove iz druga dva. Obrazloženje izbora:

**Zašto ne generisani per-tabela potpisi.** Predlog sa 46 generisanih funkcija
(`AppendOtkup(OtkupID:=..., Datum:=...)`) daje najviše kompajl-vremenske provere, ali ima dva obarača:
(1) ime `StrogUpisReda` **ne sadrži** `AppendRow`, pa `MUTATE_RE` postaje slep i A11 kapija ostaje
zelena ne mereći ništa — a ako bi se ime uskladilo, `who_writes` bi pripisao **sve 46 tabela**
generisanom modulu i mapa vlasništva se raspada; (2) 643 kolone u 46 potpisa je generisan fajl reda
veličine celog `modSchema.bas`, a compile hvata grešku samo ako pisac stvarno koristi imenovane
argumente.

**Zašto se kanon ne koristi za poziciju.** Kanon je autoritet za **imena i njihov skup**, ne za
poziciju u živoj svesci. Sveska korisnika može imati drugačiji redosled (self-heal dodaje kolone koje
falе, ne preuređuje postojeće). Zato imenovani upis razrešava ime → indeks nad **živim headerom**
(`RequireColumnIndex`), a kanon koristi samo za širinu i za „postoji li ta kolona uopšte".

### 2.1 Mehanizam već postoji — proširuje se, ne izmišlja

`reuse > new`: `modSchemaGuard.bas:281` već ima

```vb
Public Sub SetRowValueByColumn(ByRef rowData() As Variant, _
                               ByVal tableName As String, _
                               ByVal columnName As String, _
                               ByVal value As Variant, _
                               ByVal sourceName As String)
```

i koristi se **90 puta** — `modDokumenta` 63, `modOtkup` 24, `modSchemaGuard` 2, testovi 1. Dakle dva
najkanonskija modula **već pišu po imenu**. Fali samo (a) kapija širine i (b) append koji taj red
bezbedno položi.

### 2.2 Javna površina (sve u `modSchemaGuard.bas`, ASCII)

Deklaraciona sekcija, uz postojeće `ERR_SIS_OD`/`ERR_SIS_DO` (`:21-22`), **pre prve procedure**.
Zauzeti kodovi `RaiseSistemski` su 1, 2, 10–13, 20–25; blok 30–33 je slobodan:

```vb
Public Const ERR_SIS_ARNOST As Long = 30
Public Const ERR_SIS_KANON  As Long = 31
Public Const ERR_SIS_NIJE_NIZ As Long = 32
Public Const ERR_SIS_NEPOZNATA_KOLONA As Long = 33
```

Nove procedure:

```vb
' Prazan red pune sirine tabele, spreman za SetRowValueByColumn.
Public Sub PrazanRedTabele(ByRef rowData() As Variant, _
                           ByVal tableName As String, _
                           ByVal sourceName As String)

' Imenovan upis: validira SVE pre lo.ListRows.Add, pa polozi red.
' Ime se ZAVRSAVA na AppendRow da ga tools/who_writes.py MUTATE_RE vidi.
Public Function ImenovaniAppendRow(ByVal tableName As String, _
                                   ByRef rowData() As Variant, _
                                   ByVal sourceName As String) As Long

' Kapija sirine, upotrebljiva i samostalno.
Public Sub RequireRedArnost(ByVal tableName As String, _
                            ByRef rowData() As Variant, _
                            ByVal sourceName As String)
```

`SetRowValueByColumn` i `TabelaBrojKolona` ostaju **nepromenjeni**.

### 2.3 Redosled u telu je obavezujući

Presudan nalaz merenja: danas se red dodaje na `modDataAccess.bas:283`
(`Set newRow = lo.ListRows.Add`), a `ErrHandler` na `:304` samo uradi `Debug.Print` i vrati 0 —
**`newRow.Delete` ne postoji**. Svaka provera ubačena *posle* `Add`-a proizvodi duh-red pri svakom
odbijanju.

```
1. GetTable(tableName) Is Nothing        -> RaiseSistemski   (danas: tiho vrati 0)
2. IsArray(rowData), 1-D                 -> ERR_SIS_NIJE_NIZ
3. sirina kanona vs sirina sveske         -> ERR_SIS_KANON (poruka imenuje OBE sirine)
4. UBound-LBound+1 <> sirina              -> ERR_SIS_ARNOST
5. ---- TEK SADA ----  Set newRow = lo.ListRows.Add
6. petlja punjenja BEZ Application.Min klampa
7. greska u petlji -> newRow.Delete -> Err.Raise dalje
8. StampRowAudit / WriteJournalRow / InvalidateTableCache kao danas
```

### 2.4 Drugi potrošač istog niza

`modJournaling.WriteJournalRow` (`:116`) ispisuje CSV red petljom `LBound..UBound` **bez klampa i bez
dopune**, dok header (`:100`, `GetTableHeaders`) ispisuje **sva** imena kolona tabele. Za svih 14
mesta sa kraćim nizom žurnalni CSV **već danas** ima manje polja nego imena u zaglavlju. Ceo `Sub` je
pod `On Error Resume Next` (`:13`), pa se ne vidi.

To nije posledica ovog refaktora nego postojeći kvar revizijskog traga koji refaktor **zatvara sam
po sebi** (pun red = pun CSV). Ulazi u plan kao provera, ne kao zasebna ispravka.

---

## 3. Rez na PR-ove

Pravilo kroz sve: **svaki PR ostavlja projekat kompajlabilnim i sam po sebi zelen.**

### Faza A — temelj, ne menja ponašanje

**A1 · `ImenovaniAppendRow` + kapija + merilo napretka**
`modSchemaGuard.bas` dobija tri procedure iz §2.2 i četiri `ERR_SIS_*` konstante.
**Nijedan produkcioni pisac se ne dira** — jedini pozivaoci su testovi.
`tools/popis_citalaca.py` dobija grupu `pozicioni_upis` sa pragom **36** (prag pada svakim PR-om,
pada i iznad i ispod praga, kao postojeće grupe).
Suite `imenovani-upis` u `modTest` tvrdi: pun red prihvaćen · kraći **odbijen po imenu** · duži
**odbijen po imenu** · nepoznato ime kolone odbijeno · širina kanona ≠ širina sveske odbijena ·
**odbijen upis ne ostavlja red u tabeli**.

> `tools/` sme u isti PR sa kodom; `.claude/` ne sme (`CLAUDE.md` §6).

### Faza B — prevod pisaca, prag pada

Redosled je po polumeru štete, a ne po lakoći:

| PR | Moduli | Mesta | Prag posle | Napomena |
|---|---|---|---|---|
| **B1** | `modConfig`, `modKooperant`, `modMaticniKorisnici`, `modMaticniUnos`, `modStornoContext`, `modStornoZurnal`, `modSetup`, `modBankaImport` | 8 | 28 | matični/infra, po jedno mesto; `modMaticniUnos.DodajPrazanRed` je već imenovana granica alata (`MUTATE_DYN_RE`) i mora takva ostati |
| **B2** | `modCenovnik` (8/11), `modAgrohemija` (13/17), `modAmbalaza` (10/16) | 3 | 25 | **`modCenovnik` je najrizičnije mesto u repou**: nema `SchemaReadyOrFail`, a `:124` nosi *rukom održavan komentar* sa redosledom kolona. `modAmbalaza` **čeka odluku o `Stornirano`** |
| **B3** | `modUtovar` | 4 | 21 | `:1265-1267` komentar **eksplicitno** računa na tiho dopunjavanje kratkog niza — prevod mora da imenuje šta ide u 8 kolona koje danas ostaju nedirnute |
| **B4** | `modFaktura` (21/30, 9/16), `modNovac` (17/26, 4/8) | 4 | 17 | **`modFaktura` čeka odluku o `IzdatoStatus`**. `modNovac.bas:204-220` ima `RequireColumns` sa 17 imena i komentar „redosled prati `Array(...)` ispod" — rukom održavana kopija redosleda, ista bolest |
| **B5** | `modDokumenta` (10), `modOtkup` (2), `modPaletniList` (2), `modSEFPersistance` (2), `modMalina` (1) | 17 | **0** | kanonski pisci dokumenata; već koriste `SetRowValueByColumn` 87×, pa je prevod najmehaničkiji uprkos broju |

### Faza C — zatvaranje

**C1 · Kapija u `AppendRow` kao zaštitna mreža**
Tek kad je `pozicioni_upis` = 0: `AppendRow` dobija `RequireRedArnost` **pre** `ListRows.Add`,
`newRow.Delete` u `ErrHandler`, i gubi `Application.Min` klamp. Ovde se proverava i `modJournaling`
(pun red = pun CSV).
**Kapija diže `Err.Raise`, ne vraća 0** — inače je nemerljiva na 11 poziva koji povratnu vrednost ne
čitaju (⚠ FALSE-GREEN).

**C2 · 15 testnih mesta** na imenovani upis; 3 testna kratka niza prestaju da postoje.

### Faza D — ostale popravke (nezavisne, posle C, mogu paralelno međusobno)

**D1 · `schema.json` v3: tip, obaveznost, unique, FK**
Kompatibilno unazad — `columns` ostaje lista imena, dodaje se opcion `columnMeta`. Generator emituje
`SchemaKolonaObavezna` / `SchemaKolonaTip` / `SchemaObavezneKolone`.
**`required` se NE sme upaliti pre nego što faza B završi** — upaljeno ranije, svih 11 kratkih
pozicionih upisa postaje tvrd pad, uključujući fakturu, novac i utovar.

**D2 · Integritet na kanonsko članstvo**
Izmereno: `RunAllChecks` zove 21 proveru; **11 spaja preko `BrojZbirne`**, **0 preko kanonskog
članstva**, a `tblOtpremnicaIzvori` i `tblZbirnaIzvori` imaju **0 čitanja** u `modIntegritet`.
Prvi korak nije prepravka postojećih nego **prva provera koja uopšte postoji** nad članstvom: sirota
veza, dupli par, članstvo na stornirano zaglavlje, članstvo na nepostojeći red, dva aktivna članstva
za isto dete. Zatim A1/A2/B1/B4/B5/B5b na `ZbirnaID` umesto na broj.
Usput: `INTEGRITET_PROVERE.md` dokumentuje B2/B3 kojih u `RunAllChecks` **nema**, a komentari
`modIntegritet.bas:18,221,225` upućuju na `modSledljivost` — modul obrisan u S3e-1.

**D3 · Kapija mrtve javne površine**
Izmereno: 427 kandidata bez reference van svog modula, ali **siguran nalaz je 114** (funkcija ili
`Sub` sa najmanje jednim obaveznim parametrom), od toga 107 produkcionih. **`modTheme`: 49/49 javnih
procedura mrtvo, 983 linije, nula referenci van modula.** 13 modula ima 100% mrtvu javnu površinu.
**78 `Sub`-ova bez parametara NE ULAZI u brisanje** — sveska nije u repou, pa dugme u `.xlsm` iz
koda nije vidljivo; traži se potvrda operatera po imenu. `modDokumenta.GeneracijaIDZaBroj` i
`ApplyNovaGeneracijaID` čekaju S6 (generacija na prijemnici) i brišu se s njim.

**D4 · Ostali pozicioni putevi**
`modGoogleSheets.AppendRowToSheet` (3 pozivaoca, van sveske → Google Sheets) i 3 `Array` u
`modSetup.EnsureDataTable` koja se **razilaze sa kanonom**. Isti princip, drugi polumer.

---

## 4. Šta plan NE radi

- Ne dira `.frx` i ne uvodi `Private WithEvents` (`CLAUDE.md` §3).
- Ne menja vlasništvo upisa ni jedne tabele (A11 ostaje kakav je).
- Ne menja semantiku ni jednog upisa — izuzev dve konvencije iz `DOMAIN GAP`, koje su **odluka
  operatera**, ne odluka koda.
- Ne uvodi nov modul. Sve u `modSchemaGuard` i `modDataAccess`.
- Ne menja pravilo „nove kolone idu na kraj" dok faza B nije završena. **Trenutak kada to pravilo
  prestaje da bude invarijanta zbog koje se pada je PR C1**, i tada se `CLAUDE.md` §3 menja —
  zasebnim process PR-om.

---

## 5. ACCEPTANCE CONTRACT

Plan dokaza, ne dokaz — kod još ne postoji (`pre-flight` §1).

**Šta važi kad se završi**
Nijedan produkcioni pisac ne prosleđuje pozicioni niz. `pozicioni_upis` = 0. Kolona ubačena u sredinu
tabele **ne može** da pošalje vrednost u pogrešnu kolonu, jer se ime razrešava nad živim headerom.
Nepun red pada po imenu pre nego što red uopšte nastane.

**Šta mora ostati netaknuto**
`python tools/who_writes.py --check-ownership` exit 0, ista lista vlasnika. `gen_schema_module.py
--check` otisak nepromenjen kroz fazu B. `modTest` 199 testova ne pada. 2 golden scenarija nepromenjena.
`RequireColumnIndex`/`GetColumnIndex` semantika nedirnuta.

**Koji edge case mora proći**
Sveska sa **drugačijim redosledom** kolona od kanona: upis mora da ide u tačnu kolonu po imenu.
Sveska sa kolonom koja u kanonu ne postoji: širina se razilazi → pada po imenu, ne piše tiho.

**Koji negativan slučaj mora biti odbijen**
Kraći red · duži red · red sa nepoznatim imenom kolone · red koji nije niz · sveska uža ili šira od
kanona. Svaki od njih diže `Err.Raise` sa kodom iz `ERR_SIS_*` i **ne ostavlja red u tabeli**.

**Čime se dokazuje**
`python tools/vba_check.py` (radi bez Excela) · `python tools/run_vba.py --suite imenovani-upis` ·
`python tools/who_writes.py --check-ownership` · `python tools/gen_schema_module.py --check` ·
`python tools/popis_citalaca.py --check`.
**Dokaz u oba smera je obavezan** (§5 `CLAUDE.md` — menja se checker): pokvari `RequireRedArnost` →
`dokaz.py` mora da javne pad **po imenu te tvrdnje** → vrati → zeleno. Nove sabotaže u rezu:
`python tools/dokaz.py imenovani-upis`; pun katalog (598 unosa) pred release.
`Alt+F11 → Debug → Compile VBAProject` je ručna kapija pred svaki merge.

---

## 6. Odluke koje čekaju operatera

| ID | Pitanje | Blokira |
|---|---|---|
| `UPIS-KONV-01` | `tblAmbalaza.Stornirano` se danas **ne upisuje** (`modAmbalaza.bas:175-184`); čitaoci računaju „prazno = nije stornirano". Da li imenovani upis piše eksplicitnu vrednost ili izostavlja ključ? | PR B2 |
| `UPIS-KONV-02` | `tblFakture.IzdatoStatus` ostaje prazan (`modFaktura.bas:299`); konvencija je „prazno = IZDATO". Isto pitanje. | PR B4 |
| `UPIS-KONV-03` | `modUtovar.bas:1265-1267` računa na tiho dopunjavanje 8 kolona. Šta u njih ide? | PR B3 |
| `UPIS-PLATFORM-01` | Šta `lo.ListRows.Add` ostavlja u kolonama koje petlja ne popuni (formule, autofill)? Traži sondu na Windows + Excelu. | ne blokira; menja formulaciju u §5 |

---

## 7. Preuzimanje grane

```bash
cd ~/Documents/GitHub/otkupapp-pwa
git fetch origin claude/great-hawking-91ti26
git checkout claude/great-hawking-91ti26
git pull origin claude/great-hawking-91ti26
python tools/vba_check.py
```
