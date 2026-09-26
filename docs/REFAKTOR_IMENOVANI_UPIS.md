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
| `DOMAIN` | **ZATVOREN** (v. §2.4) | Prvobitno prijavljen kao GAP. Zapravo je odluka **vec doneta i zapisana** u kanonskom piscu: `modOtkup.bas:1444-1446` pise `IzdatoStatus` EKSPLICITNO, uz komentar "nov model se ne oslanja na legacy konvenciju prazno = IZDATO". Periferija to samo nije ispratila. Dva pisca se oslanjaju na nezapisanu konvenciju „prazno znači X": `modAmbalaza.bas:175-184` šalje 10 vrednosti u `tblAmbalaza` (16 kolona) i `Stornirano` nikad ne upisuje; `modFaktura.bas:299` šalje 21 u `tblFakture` (30) i `IzdatoStatus` ostavlja prazan. Imenovani upis prisiljava odluku: upisati eksplicitno ili izostaviti ključ. **Blokira samo PR-ove B2 i B4**, ne ceo plan. |
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

## 2. Obrazac ne treba projektovati -- postoji i dokazan je

**Ispravka prvobitne verzije ovog plana.** Prvobitno je ovde stajao izbor izmedju tri predloga novog
API-ja (`ImenovaniAppendRow` i sl.). To je bilo nepotrebno: obrazac postoji, koristi se 89 puta i
**prosao je stvaran test u produkcionoj istoriji**.

### 2.1 Kanonski obrazac

`modOtkup.BuildOtkupHeaderRowData` (`src-vba/modOtkup.bas:1408`):

```vb
Dim colCount As Long
colCount = TabelaBrojKolona(TBL_OTKUP)
If colCount <= 0 Then Err.Raise vbObjectError + 1883, SRC, "..."

Dim rowData() As Variant
ReDim rowData(0 To colCount - 1)          ' PUNA sirina, uvek

SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_ID, otkupID, SRC
SetRowValueByColumn rowData, TBL_OTKUP, COL_OTK_DATUM, datum, SRC
' ... po imenu za svako polje
```

Tri svojstva koja resavaju ceo problem:
1. **Puna sirina** iz `TabelaBrojKolona` -- niz nikad nije kraci, pa `Application.Min` klamp nema sta
   da odseca i nijedna kolona ne ostaje nedirnuta slucajno.
2. **Upis po imenu** -- pozicija se razresava iz zive sveske, pa je redosled nebitan.
3. **Bez "ako kolona postoji"** -- kanonska kolona koja fali **pada**, ne postaje prazno polje.
   Komentar na `:1437` to izricito obrazlaze: tiho preskakanje bi sakrilo bas drift zbog kog kapija
   postoji.

### 2.2 Dokaz da obrazac radi

`018be597` (S1d, 17.09.2026) obrisao je **8 kolona iz SREDINE `tblOtkup`** (37 -> 29 kolona;
`Kolicina` sa pozicije 8, `Cena` sa 9, `KolAmbalaze` sa 11, `Novac` sa 14 -- sve iza pozicije 8 se
pomerilo). U tom commitu **`src-vba/modOtkup.bas` nije diran nijednom linijom**, a pisac je nastavio
da radi ispravno.

To je najjaci moguci dokaz: obrazac je prosao tacno onaj dogadjaj zbog kog se ovaj refaktor radi.

### 2.3 Sta se onda radi

Ne projektuje se nista novo. Periferni pisci se prevode **na postojeci obrazac** -- `Build*RowData`
helper sa `ReDim` pune sirine i `SetRowValueByColumn` po imenu. Nijedna nova javna procedura, nijedan
nov modul, nijedna izmena `modDataAccess.AppendRow` dok prevod ne zavrsi.

## 3. Rez na PR-ove

### Konačno merenje (tri revizije; ova važi)

Klasifikacija po pozivu, sa razrešenim helperom (i kad se zove inline kao drugi argument, i kad ide
kroz promenljivu):

| Stanje | Mesta |
|---|---|
| **Imenovano** — puna širina + razrešenje po imenu | **18** |
| **Pozicionо** — goli `Array(...)` | **21** |
| Ukupno produkcionih | 39 |

**Jezgro je 100% imenovano.** Nijedno pozicionо mesto nije u `modDokumenta` ni `modOtkup` — dakle
pisci `tblOtkup`, `tblOtpremnica`, `tblZbirna`, njihovih stavki **i tabela članstva**
(`tblZbirnaIzvori` preko `BuildZbirnaIzvorRowData`) svi rade po imenu.

Preostalih 21 je periferija: `modPaletniList` 4, `modUtovar` 4, `modNovac` 2, `modSEFPersistance` 2,
`modFaktura` 2, `modAgrohemija` 2, te po jedno `modStornoZurnal`, `modStornoContext`, `modCenovnik`,
`modSetup`, `modAmbalaza`.

### Tri imenovana idioma već postoje

Merenje je našlo **tri nezavisna** načina na koja repo već piše po imenu — zato se ništa ne
projektuje, nego se bira jedan:

| Idiom | Primer | Baza niza |
|---|---|---|
| `Build*RowData` + `SetRowValueByColumn` | `modOtkup.bas:1408`, `modDokumenta` (7 helpera, 63 upotrebe) | **0**-bazna (`rowData(colIndex - 1)`, `modSchemaGuard.bas:289`) |
| `ReDim 1 To count` + `RequireColumnIndex` kao indeks | `modUtovar.bas:157-163`, `modMalina.bas` | **1**-bazna |
| lokalni `Set*Cell` omotač | `modAgrohemija.bas:403` | varira |

**Posledica za kapiju arnosti:** baza niza **nije uniformna**. Formula koja pretpostavlja
`LBound = 0` je pogrešna za `modUtovar`/`modMalina`, a formula koja pretpostavlja 1 je pogrešna za
`SetRowValueByColumn`. Kapija mora da normalizuje `LBound`, ne da ga pretpostavi.

Za prevod periferije preporučen je **prvi idiom** — ima najviše upotreba, jedini je dokazan stvarnim
događajem (§2.2), i drži pravilo „kanonska kolona koja fali pada".

### Redosled

| PR | Sadržaj | Mesta | Zašto tu |
|---|---|---|---|
| **P1** | `modJournaling.WriteJournalRow` dopunjava red do širine zaglavlja | — | **Jedini PRISUTAN kvar u skupu.** `:116` ispisuje CSV bez dopune, `:100` ispisuje sva imena kolona; za pisce sa kraćim nizom žurnalni CSV **danas** ima manje polja od zaglavlja, pod `On Error Resume Next`. Nezavisno od svega ostalog, ~10 linija |
| **P2** | `modCenovnik`, `modFaktura`, `modNovac`, `modAgrohemija`, `modAmbalaza` | 8 | Kraći nizovi. `modCenovnik` prvi — nema `SchemaReadyOrFail`, a `:124` nosi **rukom održavan komentar sa redosledom kolona**. `modFaktura` uz njega jer `tblFakture` čita SEF |
| **P3** | `modPaletniList`, `modUtovar`, `modSEFPersistance`, `modStornoZurnal`, `modStornoContext`, `modSetup` | 13 | Mehanički. `modUtovar` traži odluku `UPIS-KONV-03` |
| **P4** | Kapija arnosti u `AppendRow` + `newRow.Delete` u `ErrHandler` | — | **Zaštitna mreža, ne otkriće.** Tek posle P2/P3. PRE `lo.ListRows.Add` (`:283`), `Err.Raise` a ne `return 0`, i sa **normalizovanim `LBound`** |
| **P5** | 15 testnih mesta | 15 | Bez pritiska |

`tools/popis_citalaca.py` dobija grupu `pozicioni_upis`, prag **21** i pada svakim PR-om.

### Ograničenja koja važe za svaki plan nad ovim slojem

Devet nezavisnih ocena predloga (3–7/10, sve sa fatalnom zamerkom) ostavilo je pet činjenica koje
vrede bez obzira šta se radi:

1. **Ime novog mutatora mora da prođe pravi `MUTATE_RE` pre nego što se izabere.** Token je
   `UpdateCell`, **ne** `UpdateRow` — pa bi `ImenovaniUpdateRow(TBL_NOVAC, ...)` bio **nevidljiv** za
   A11 kapiju. Izvršeno nad pravim regexom: `ImenovaniAppendRow(TBL_CENOVNIK` → MATCH,
   `ImenovaniUpdateRow(TBL_NOVAC` → NO MATCH.
2. **Tabela mora ostati PRVI argument kao `TBL_` konstanta u istoj logičkoj liniji**, inače upis pada
   u `MUTATE_DYN_RE` („mapa ne može da pripiše").
3. **`RaiseSistemski` kodovi 1, 2, 10–13, 20–25 su zauzeti** (izmereno; 20–25 su u
   `modDokumenta.bas:4186,4195,4444,4454,4631,4639,4644`). Slobodno od 30.
4. **`tools/vba_check.py` već statički hvata** „Sub or Function not defined" (`:444`) i „Wrong number
   of arguments" (`:461`), bez Excela. Jeftina nezavisna dobit: proširiti `collect_arities` (`:333`)
   da skuplja i **imena** parametara, pa validirati svaki `ime:=` poziv u CI-ju.
5. **Ne dodavati četvrtu definiciju „ispravnog prefiksa".** `modSchema.PrefiksNeslaganje:546` je
   `Private`, a komentar na `:544` izričito kaže da dve kapije ne smeju da razviju različite
   definicije. Postojeći helper se koristi, ne duplira.

> **Šta ovaj plan NIJE dobio:** kritika kompletnosti i adversarna kritika rizika migracije nisu se
> izvršile — pale su na session limit dva puta. Plan nije prošao nezavisnu proveru na propuste.

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

---

## 8. Značaj u ovom trenutku — iskren rang

Merenje je oborilo prvobitnu procenu značaja ovog posla. Redosled po stvarnoj koristi, ne po tome
koji je nalaz prvi našao:

| # | Posao | Značaj | Zašto |
|---|---|---|---|
| 1 | **Integritet nad kanonskim članstvom** (`tblOtpremnicaIzvori`, `tblZbirnaIzvori`) | **VISOK** | `RunAllChecks` zove 21 proveru; **11 spaja preko `BrojZbirne`, 0 preko članstva**, a tabele članstva imaju **nula čitanja** u `modIntegritet`. Najnoviji i najmanje testiran deo modela — ono što je poslednjih 100 PR-ova izgradilo — **nema ni jedan sken integriteta**. Za redosled kolona postoji pravilo (`CLAUDE.md` §3) i dokazan obrazac pisca; za članstvo ne postoji ništa |
| 2 | **`modJournaling` CSV** | SREDNJI | Jedini **prisutan** kvar iz ovog skupa, ~10 linija. Revizijski trag je već nepotpun |
| 3 | **Prevod 21 perifernog pisca** | SREDNJE‑NIZAK | Mehanički, po postojećem obrascu. Nijedna od tih tabela nije u planu skraćivanja za S3e-2 ni S6, pa okidač ne dolazi skoro |
| 3b | **`vba_check`: validacija imena parametara u `ime:=` pozivima** | SREDNJE‑NIZAK | Radi bez Excela, u CI-ju; `collect_arities` (`tools/vba_check.py:333`) već skuplja arnost, treba mu samo lista imena. Hvata klasu koju danas hvata jedino ručni Compile |
| 4 | **Kapija arnosti** | NIZAK | Zaštitna mreža posle #3, ne otkriće. Ne može samostalno: 7 pisaca bi palo istog trenutka |
| 5 | **Tipovi / FK / unique u kanonu** | NIZAK | Upis po imenu već uklanja rizik pozicije. Tipovi hvataju drugu klasu koja još nije ugrizla |
| 6 | **Brisanje mrtve površine** | NIZAK | Higijena. `modTheme` 49/49 mrtvo, 983 linije — ali 78 `Sub`-ova bez parametara ne sme u brisanje bez potvrde operatera (sveska nije u repou) |

**Šta NIJE rizik, iako je zvučalo tako.** Jezgro modela je već imenovano i **dokazano stvarnim
događajem** (§2.2). Pomeraj kolone u `tblOtkup`/`tblOtpremnica`/`tblZbirna` piscima više ne može
ništa, a to su jedine tabele koje S3e-2 skraćuje.
