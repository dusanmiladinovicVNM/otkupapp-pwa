---
paths:
  - "src-vba/mod*Tests.bas"
  - "src-vba/modTest*.bas"
  - "src-vba/frmOtkup.frm"
  - "src-vba/modOtkupUI.bas"
  - "src-vba/modScrDokumenti.bas"
  - "src-vba/clsTest*.cls"
  - "tools/vba_check.py"
  - "tools/vba_gate.py"
  - "tools/run_vba.py"
  - "tools/release_gate.py"
  - "tools/sabotaza.py"
  - "tools/make_fixture.py"
  - "tools/dump_schema.py"
  - "tests/suite_manifest.json"
  - "tests/golden/*"
---

<!-- frmOtkup.frm / modOtkupUI.bas / modScrDokumenti.bas su u paths namerno: u
     njima su test seam-ovi (§4). Bez ovoga agent koji menja formu ne bi znao da
     postoje. -->

# Verifikacija — politika

> Ovde su samo **pravila**. Kako se harness koristi (fixture, golden, pisanje
> testa, zamke sabotaže, trijaža padova): `docs/EXCEL_TEST_HARNESS.md`.
> Zašto su pravila baš ovakva — incidenti i cena:
> `docs/engineering/postmortems/2026-08-verifikacija.md`.

## 1) Tri nivoa — šta se kada pušta

| Nivo | Kada | Komanda |
|---|---|---|
| **FAST** | posle svake VBA izmene; automatski kroz `PostToolUse` hook | `python tools\vba_check.py` |
| **TARGETED** | za feature ili bug — suite koja pokriva to područje | `python tools\run_vba.py --suite <ime>` |
| **FULL** | pred release i za rizične izmene u jezgru | `python tools\run_vba.py` |

`FULL` je **namerna komanda**, ne automatika. Nije u hook-u i ne sme da se vrati
u hook: 11 suite-ova uz podizanje Excela na svakom zaustavljanju je desktop
sesiju činilo neupotrebljivom.

**Centralni verdikt nezavisan od mašine daje CI** (`.github/workflows/static.yml`):
JSON validnost `settings*.json`, `vba_check`, `who_writes --check`, `vba_gate
--manifest-check` i self-testovi alata. CI ne pokreće Excel i nikad neće.

## 2) `tools/vba_check.py` — radi svuda, i u Linux sesiji

Exit `0` = čisto, `2` = ima nalaza. **Obavezan pre commita** svake VBA izmene.

| Provera | Šta hvata |
|---|---|
| `ASCII` | ne-ASCII bajt u `.bas`/`.cls`/`.frm`/`.doccls` |
| `DEKLARACIJA` | modul-level `Const`/`Dim`/`Declare`/`Type`/`Enum` posle prve procedure |
| `REZERVISANO` | ime koje se case-insensitive poklapa sa VBA ključnom reči (`eNum`) |
| `DUPLIKAT` | isti `Public Sub/Function/Const` u dva modula → „Ambiguous name" |
| `DUPLIKAT_LOKALNI` | isto ime dva puta u **istom** modulu → modul se ne kompajlira |
| `PORUKA` | `Poruka("KLJUC")` bez para u `modPoruke.UpsertPoruke` |
| `NEDEFINISAN` | poziv procedure koja nigde nije definisana |
| `ARNOST` | poziv sa pogrešnim brojem argumenata |
| `ZAKLONJENO` | lokalni skalar zvan sa zagradom → „Expected array" |

`NEDEFINISAN`/`ARNOST` su namerno uske (samo `.bas`, samo poziv u poziciji
naredbe) — lažan nalaz je gori od propuštenog. **Poziv u izrazu (`x = Foo(1)`)
se ne proverava**: bez tipova se poziv funkcije ne razlikuje od indeksiranja
niza. To je poznata rupa, ne previd — pokušaj proširenja je dao 406 lažnih
nalaza (ime funkcije unutar string literala, između ostalog). Uzak izuzetak od `DUPLIKAT`-a
postoji za ugovor ekrana (`Scr_*` u `modScr*`), a od `DUPLIKAT_LOKALNI`-og za
`Property Get/Let/Set` trojku. Ne kompajlira VBA: ne hvata tip-greške ni
nedeklarisane promenljive, a u `.frm`/`.cls` ne radi `NEDEFINISAN`/`ARNOST`.

**`ZAKLONJENO` postoji zbog rupe koju su ostale tri provere ostavile.** Lokalno
ime koje se poklapa sa imenom funkcije **zaklanja** je unutar te procedure (VBA je
case-insensitive), pa poziv postaje indeksiranje:

```vb
Public Function StornoIzvrsi(..., ByRef poruka As String, ...)
    poruka = Poruka("STORNO_MSG_OK")      ' Expected array
```

Dvanaest takvih poziva je živelo u dve procedure od `v6-ui-119` do `v6-ui-141`.
Nije ih videla **nijedna** postojeća kapija: suite ne, jer VBA kompajlira
proceduru **tek kad se pozove** a te dve je zvao samo UI; `ARNOST`/`NEDEFINISAN`
ne, jer je poziv u poziciji izraza; CI ne, jer ne pokreće Excel. Našao ih je
operater ručnim `Debug → Compile`.

Provera je namerno **uža** od „ime zaklanja funkciju": skalar eksplicitnog tipa se
u VBA ne može indeksirati nikako, pa je `ime(` uz `Dim ime As String` uvek greška.
Izostavljeni su `Variant` (može nositi niz), nizovi, objekti (default member) i —
najvažnije — sadržaj string literala i komentara: 14 od prvih 20 nalaza ovog
obrasca bilo je ime unutar teksta, ista klasa lažnih nalaza zbog koje je širenje
`ARNOST`-a odbijeno.

**Dve provere duplikata nisu ista provera.** `DUPLIKAT` gleda globalni imenski
prostor (isto `Public` ime u dva modula); duplo ime unutar jednog modula mu je
nevidljivo. A ono obara compile isto tako — i **ne prijavljuje se kao compile
greška** nego kao `Cannot run the macro` na bilo kom makrou, jer modul koji se ne
kompajlira obara ceo projekat. Simptom ne pokazuje na krivca.

`python tools\vba_check.py --self-test` vrti slučajeve koji **moraju** da zapište
i legalan VBA koji **ne sme**. Vrti se i u CI-ju: zelen checker nad čistim repoom
ne razlikuje „nema greške" od „provera ništa ne meri".

Slučajevi idu **kroz `check_file()`, istu funkciju koju zove CLI**, a jedan ide
kroz ceo `main()` nad pravim fajlom na disku. Namerno: self-test koji zove
proveru direktno dokazuje da funkcija radi, ali ne i da je CLI zove — otkačen
jedan red tada ostavlja i repo-run i self-test zelene, a checker isključen. Isti
oblik greške kao placebo test, samo u alatu.

## 3) `tools/run_vba.py` — SAMO Windows + Excel + `pywin32`

Ne pokreće se na Linux/macOS — ni u web sesiji ni u CI. Tamo se testovi ponašanja
**ne izvršavaju uopšte**, pa se VBA izmena koja dira ponašanje prijavljuje kao
**neverifikovana**, nikad kao zelena.

- **Katalog je `tests/suite_manifest.json`, ne kod** — koja suite postoji, u
  kojoj je kapiji (`dev`/`pr`/`release`), da li podiže grešku (`raises`), koliko
  provera najmanje mora da prijavi (`min_asserts`) i koliki joj je timeout. Ne
  prepisivati ga nigde. `python tools\vba_gate.py --manifest-check` proverava oba
  smera: upisana suite mora da postoji u kodu, a nova ulazna tačka koju niko ne
  poziva mora da bude upisana — ili u `suites`, ili u `unlisted` sa razlogom.
- `raises: true` znači da suite podiže grešku pa je runner vidi kao crvenu.
  `false` piše samo u Immediate i runner je prijavljuje kao **`blind`** — to znači
  „prošla bez greške", **ne** „sve provere prošle". **Nova suite mora biti gate.**
- Verdikt ne dolazi iz toga da li `Run()` pukne, nego iz izveštaja pored sveske
  (`last_run.txt`, `last_run_banka.txt`, `suite_results.txt`). **Nema fajla = pad**,
  a prijavljen `FAIL > 0` obara run i kad `Run()` nije pukao.
- **`NOT_RUN` nije prolaz**, ni za suite ni za ceo run. Run u kome nijedan skup
  nije dao `PASS` izlazi sa 2.
- **`--no-import` nikad ne upisuje last-green marker** — sveska tada nosi tuđi
  kod, a hash bi opisivao repo.
- **Fixture je gitignored, pa ga `git checkout` NE menja.** Posle prelaska na
  drugu granu na disku ostaje sveska prethodne i testovi padaju *na podacima*, a
  pad izgleda kao regresija koda. Runner to sada hvata sam: `make_fixture` piše
  potpis podataka u `otkup_test.sig`, `run_vba` ga poredi pre Excela i staje uz
  komandu za regeneraciju. Bez važećeg potpisa run **ne prolazi** — jedini izlaz
  je svestan `--ignore-fixture-sig`. **Crveno posle prelaska grane — prvo
  regeneriši fixture, pa tek onda traži krivca u kodu.** Detalji:
  `docs/EXCEL_TEST_HARNESS.md`.
- `COMPILE NEJASNO` ne obara run kad suite-ovi idu. **Compile ostaje ručna kapija
  pred release:** `Alt+F11 → Debug → Compile VBAProject`.
- Kapije: `--gate dev` (podrazumevano), `--gate pr` (pun lokalni set),
  `--gate release` (+ external ugovori). U `pr`/`release` je i **CLEANUP
  blokirajući** (zaostao `TST-`/`SVT-`/`BIT-` red obara run); `--no-enforce-cleanup`
  je izlaz dok detektor ne bude dokazan u oba smera.
- **Verdikt nije jedan**: `STATIC`, `COMPILE`, `SCHEMA`, `BEHAVIOR`, `COUNTS`,
  `CLEANUP`, `EXTERNAL` — svaki svoj red. Proces, identitet run-a i Definition of
  Done: `docs/TEST_PLATFORM.md`.

## 4) Test seam-ovi koje produkcioni kod nosi

Ako menjaš `frmOtkup`, `modOtkupUI` ili `modScrDokumenti`, znaj da ovo postoji i
zašto — inače ćeš ih „počistiti":

- `ClearForm` / `ParseDatum` / `ParcelaID` su **`Public`**, ne `Private` — test ih
  zove direktno, bez vožnje celog upisa.
- **Tri `SetFocus`-a** su iza `If Not IsTestMode()`. Forma bez `.Show` ne može da
  primi fokus, a u nevidljivom Excelu `SetFocus` **ne puca nego trajno visi**. U
  produkciji je `IsTestMode()` uvek `False`.
- `modScrDokumenti.Scr_OtpTestSet` postoji samo za test i **tvrdo je gejtovan** —
  van test-režima ne radi ništa.

Šta `modTest` NE pokriva: mrežu, storno, i sam `Save*_TX` (transakcioni upis
pokrivaju `RunStornoTestSuite` i `RunBusinessFlowProSuite`). Forma se gradi bez
`.Show`, pa `UserForm_Activate` nikad ne ide.

## 5) Kapija ide i u writer, ne samo u modul unosa

Modul unosa proverava nad **snimkom iz trenutka kad je lista punjena**. Između
punjenja i potvrde stanje se može promeniti — drugi unos, uvoz izvoda, drugi
pozivalac istog writer-a — pa bi prošla prevelika isplata. Zato kritična
poslovna kapija (vlasništvo + aktivnost + **preračunat** preostali iznos) stoji
i u writer-u, koji se zove i iz legacy `frmDokumenta`, bez ijedne UI provere.

Obrazac je zajednički za `ApplyAvansToOtkup`, `IsplataBlokProblem` i
`UplataFakturaProblem`: **jedna implementacija, dva pozivaoca** — core je diže
kao grešku, modul unosa je vraća kao poruku uz polje. Ne praviti drugu kopiju
pravila u modulu.

Isto važi za **nerazrešen izbor**: ukucano a nerazrešeno ime, blok ili faktura
(`ListIndex = -1` uz vidljiv tekst) mora da **zaustavi** dokument, a ne da tiho
postane isplata otkupnom mestu ili avans. Prazno polje i dalje prolazi — to su
dva različita stanja.

## 6) Sabotaža — kada je obavezna

Dvosmerni dokaz (pokvari → provera pukne **po imenu** → vrati → zeleno) traži se:

- kad **dodaješ ili menjaš sam test ili checker** — zelena suite koja nikad nije
  pokazana crvena ne dokazuje da išta meri;
- kod **naročito kritične poslovne invarijante**, gde hoćeš dokaz da test nije
  placebo.

**Za običnu funkcionalnu izmenu se ne traži.** Katalog postojećih sabotaža je
`python tools/sabotaza.py --lista` — skripta je izvor istine, ne prepisivati ovde.
Zamke pri pisanju nove (kraj reda, uvlačenje, vraćanje): `docs/EXCEL_TEST_HARNESS.md`.

## 7) Šta ostaje na operateru

Finalni smoke-test u Excelu (klik po klik), izgled forme, štampa i PDF, ponašanje
nad pravim podacima. Zato svaki rad završi kratkom numerisanom test-checklistom u
chatu. Checklista **nije** zamena za test koji je moguće napisati.
