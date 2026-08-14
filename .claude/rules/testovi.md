---
paths:
  - "src-vba/mod*Tests.bas"
  - "src-vba/modTest*.bas"
  - "src-vba/frmOtkup.frm"
  - "src-vba/modOtkupUI.bas"
  - "src-vba/modScrDokumenti.bas"
  - "tools/vba_check.py"
  - "tools/run_vba.py"
  - "tools/sabotaza.py"
  - "tools/make_fixture.py"
  - "tools/dump_schema.py"
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
JSON validnost `settings*.json`, `vba_check`, `who_writes --check`. CI ne pokreće
Excel i nikad neće.

## 2) `tools/vba_check.py` — radi svuda, i u Linux sesiji

Exit `0` = čisto, `2` = ima nalaza. **Obavezan pre commita** svake VBA izmene.

| Provera | Šta hvata |
|---|---|
| `ASCII` | ne-ASCII bajt u `.bas`/`.cls`/`.frm`/`.doccls` |
| `DEKLARACIJA` | modul-level `Const`/`Dim`/`Declare`/`Type`/`Enum` posle prve procedure |
| `REZERVISANO` | ime koje se case-insensitive poklapa sa VBA ključnom reči (`eNum`) |
| `DUPLIKAT` | isti `Public Sub/Function/Const` u dva modula → „Ambiguous name" |
| `PORUKA` | `Poruka("KLJUC")` bez para u `modPoruke.UpsertPoruke` |
| `NEDEFINISAN` | poziv procedure koja nigde nije definisana |
| `ARNOST` | poziv sa pogrešnim brojem argumenata |

Poslednje dve su namerno uske (samo `.bas`, samo poziv u poziciji naredbe) — lažan
nalaz je gori od propuštenog. Uzak izuzetak od `DUPLIKAT`-a postoji za ugovor
ekrana (`Scr_*` u `modScr*`). Ne kompajlira VBA: ne hvata tip-greške ni
nedeklarisane promenljive, a u `.frm`/`.cls` ne radi `NEDEFINISAN`/`ARNOST`.

## 3) `tools/run_vba.py` — SAMO Windows + Excel + `pywin32`

Ne pokreće se na Linux/macOS — ni u web sesiji ni u CI. Tamo se testovi ponašanja
**ne izvršavaju uopšte**, pa se VBA izmena koja dira ponašanje prijavljuje kao
**neverifikovana**, nikad kao zelena.

- **Katalog `SUITES` u `tools/run_vba.py` je jedini izvor istine** — koja suite
  postoji, da li je `gate` i da li je u punom setu. Ne prepisivati ga nigde.
- `gate: True` podiže grešku pa je runner vidi kao crvenu. `gate: False` piše samo
  u Immediate i runner je prijavljuje kao **`blind`** — to znači „prošla bez
  greške", **ne** „sve provere prošle". **Nova suite mora biti `gate`.**
- Verdikt ne dolazi iz toga da li `Run()` pukne, nego iz `last_run.txt` pored
  sveske. **Nema fajla = pad.**
- `COMPILE NEJASNO` ne obara run kad suite-ovi idu. **Compile ostaje ručna kapija
  pred release:** `Alt+F11 → Debug → Compile VBAProject`.

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
