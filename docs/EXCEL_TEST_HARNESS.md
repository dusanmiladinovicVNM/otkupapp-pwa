# Excel test harness — `tools/run_vba.py`

Headless `import src-vba/` → `Debug > Compile` → test suite, bez ijednog klika.
Radi **samo na Windows mašini sa instaliranim Excelom** (COM). Na Linux/macOS
nema `VBProject` interfejsa — tamo postoji samo `tools/vba_check.py` (statičke
provere, radi svuda).

## Jednokratna priprema

```bash
pip install pywin32
```

Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings**
→ čekiraj **„Trust access to the VBA project object model"**. Bez toga skripta ne
može da uveze module (`VBProject` je zaključan).

## Upotreba

```bash
python tools/run_vba.py --compile-only     # najbrže i najstabilnije
python tools/run_vba.py                    # + podrazumevani set suite-ova
python tools/run_vba.py --suite RunBankaImportTestSuite
python tools/run_vba.py --all
python tools/run_vba.py --workbook "C:\putanja\AgriX_OtkupApp.xlsm"
```

Izlazni kod: `0` = zeleno, `2` = palo (compile greška, pala `gate` suite, ili
neočekivan modalni dijalog).

## Koja sveska se koristi

| Slučaj | Sveska |
|---|---|
| bez `--workbook` | `tests/fixtures/otkup_test.xlsm` |
| tog fajla nema | skripta ga **sama napravi** kao praznu `.xlsm` |
| `--workbook PUT` | ta sveska |

Sveska se **uvek kopira u temp folder** pre rada — original se ne dira. Zato je
bezbedno proslediti i pravi radni `AgriX_OtkupApp.xlsm`.

`tests/fixtures/` je u `.gitignore` — fixture je lokalni artefakt, ne ide u repo.

### Zašto je prazna sveska dovoljna za compile

Nijedan modul ne referencira sheet `CodeName` rano-vezano — sve `SHT_*` konstante
u `modConfig` su **string literali** (`"sOtkup"`, `"sConfig"`…), a nijedan
`.doccls` nema `Public` član koji bi neko spolja zvao. Sheet `.doccls`-ovi se u
praznoj svesci uredno preskoče (`SKIP ... (nema komponente 'sOtkup')`) i compile
i dalje pokriva ceo `src-vba/`.

### Kada prazna sveska NIJE dovoljna

Test suite čitaju `tblOtkup`, `tblKooperanti`, `tblArtikli`… Nad praznom sveskom
padaju na podacima, ne na kodu. Za suite:

```bash
python tools/run_vba.py --suite RunIzvestajTests --workbook "C:\...\AgriX_OtkupApp.xlsm"
```

## Zašto skripta ne zove `modVbaTools.ImportAllVBA`

`ImportAllVBA` ima **hardkodiran folder**
(`C:\Users\Dusan\Documents\GitHub\otkupapp-pwa\src-vba\`) i završava se
`MsgBox`-om. Oboje je smrt za headless run — `MsgBox` u nevidljivom Excelu visi
zauvek (COM poziv ne puca, samo stoji).

Import se zato radi istom logikom preko COM-a, sa folderom izvedenim iz lokacije
skripte. Pravila su identična:

- `.doccls` se **merge-uje** u postojeću komponentu (ne briše se pa uvozi)
- `.bas` / `.cls` / `.frm` se uklone pa uvezu
- `.frm` mora imati `.frx` par, inače se preskače
- `modVbaTools` se preskače (on je taj koji se izvršava)

Modalne dijaloge zatvara watchdog nit (prozori klase `#32770` u Excel procesu,
klik na podrazumevano dugme). Svaki uhvaćen dijalog se prijavljuje u izveštaju —
neočekivan dijalog je nalaz, ne šum.

## Compile: pravo rešenje je statičko, ne headless

Posle četiri pokušaja headless compile gate-a (istorija ispod) ostaje zaključak:
**dve najčešće compile greške u ovom projektu ne traže Excel.**

| Greška | Hvata je |
|---|---|
| „Sub or Function not defined" | `vba_check.py` → `NEDEFINISAN` |
| „Wrong number of arguments" | `vba_check.py` → `ARNOST` |
| „Ambiguous name detected" | `vba_check.py` → `DUPLIKAT` |

To radi u milisekundama, na svakoj platformi, i vrti se kao PostToolUse hook —
dakle greška stiže pre commita, a ne posle importa u Excel. Namerno je usko
(samo `.bas`, samo poziv u poziciji naredbe), jer je lažan nalaz u hook-u gori od
propuštenog.

Šta i dalje traži Excel: tipovi, nedeklarisane promenljive, greške u `.frm`/`.cls`.
Za to ostaje `Alt+F11 → Debug → Compile VBAProject`.

## Kako se compile stvarno meri (i zašto ne preko menija)

Prva verzija je zaključivala iz `Enabled` stanja stavke **Debug → Compile
VBAProject** („posle uspešnog compile-a postaje siva"). To **ne radi** u
nevidljivom Excelu: VBE osvežava enabled-stanje kontrola tek kad se meni iscrta,
pa stavka ostaje aktivna i kad je projekat uredno kompajliran. Rezultat je bio
`COMPILE NEJASNO` na svakom pokretanju — a, gore, `NEJASNO` je prolazilo kao
`REZULTAT: ZELENO`.

Druga verzija je verdikt tražila od **probe-a**: doda se modul sa funkcijom koja
vraća `42`, pozove se, i ako `Application.Run` prođe — projekat se kompajlira.
Ni to ne radi: **VBA kompajlira lenjo**. Prevede modul koji se zove i ono što taj
modul stvarno referencira. Probe modul ne referencira ništa, pa je namerno
ubačena greška u `modConfig` prošla netaknuta — treće lažno zeleno.

Ispostavilo se i zašto meni-compile nije prijavio grešku: **`Execute()` bez
vidljivog VBE prozora ne uradi ništa.** Kontrola „Compile VBAProject" je tada
disabled, `Execute` tiho prođe i nikakav dijalog se ne pojavi — pa je izostanak
dijaloga izgledao kao uspeh.

Treći pokušaj je VBE prozor otvarao **minimizovan** („da ne skače preko ekrana").
Minimizovan prozor nije aktivan — kontrola je ostala mrtva, `Execute` je opet
tiho prošao. To se videlo iz dijagnostike: `dialogs = []` **i na namerno
slomljenom kodu**, gde je compile error dijalog morao da iskoči.

Sada:

1. **VBE prozor se otvara u normalnom stanju i dobija fokus**, projekat se
   postavlja kao aktivan.
2. **Pokušaj 1** — `FindControl(578).Execute()`.
3. **Pokušaj 2** (samo ako prvi nije ništa uradio) — `SendKeys` `Alt+D`, `C` u
   VBE prozor. Šalje se isključivo ako `AppActivate` potvrdi da je VBE prozor
   zaista aktivan, da tastatura ne odleti u tuđu aplikaciju.
4. **Verdikt, tim redom:** uhvaćen compile dijalog → `FAIL`; kontrola postala
   disabled → `OK`; sve ostalo → `NEJASNO`.
5. **Probe** ostaje kao dodatna kontrola (hvata projekat koji uopšte nije
   izvršiv), ali nije izvor verdikta.

Ispis uvek sadrži sirovo stanje svakog pokušaja (`com_*`, `keys_*`,
`enabled_before`, `probe`), da se sledeći pogrešan ishod ne mora pogađati.

**`NEJASNO` znači da compile nije izvršen, a ne da je pao** — u tom slučaju
uradi ručno `Alt+F11 → Debug → Compile VBAProject`. Izlazni kod je i tada 2:
nepoznat ishod je nalaz, ne zeleno.

Pravilo: **sve što nije eksplicitno `OK` je pad** (izlazni kod 2). Alat koji ne
zna ishod mora to da kaže glasno, ne da ćuti u zeleno.

## `gate` vs „blind" suite

Katalog `SUITES` u skripti nosi zastavicu `gate`:

- `gate: True` — suite **podiže grešku** kad provera padne → runner je vidi kao crvenu
- `gate: False` — rezultat postoji samo u Immediate prozoru → runner je prijavljuje
  kao **`blind`**, što znači „prošla bez greške", a **ne** „sve provere prošle"

Kad pišeš novu suite, napravi je `gate` i upiši je u katalog. **Katalog `SUITES` u
`tools/run_vba.py` je jedini izvor istine** — koja suite postoji, da li je `gate`
i da li je u punom setu (`default: True`). Ne prepisivati ga u dokumentaciju.

Konverzija blind → gate, po uzoru na `modTestBanka.ERR_BIT_SUITE_FAILED`:

1. `Private Const ERR_X_SUITE_FAILED As Long = vbObjectError + <slobodan>` u
   deklaracionu sekciju (zauzeti offseti: 2900, 2950, 2960–2963, 3010–3012, 3100)
2. posle završnog izveštaja: `If mFail > 0 Then Err.Raise ERR_X_SUITE_FAILED, ...`
3. u `EH`: prebroj prekid kao pad (`Fail "SUITE prekinut..."`) pa podigni

**Prođi i rane izlaze:** uslov koji tiho radi `Exit Sub` pre prve provere mora da
podigne grešku sa porukom koja počinje `suite NIJE pokrenut:` — inače runner
„nije se pokrenulo" vidi kao `OK`. Poznata otvorena rupa:
`RunBankaImportTestSuite` (rani `Exit Sub` kad `tblBankaImport` / `tblOtkup` ne
postoje).

**Ne proširivati na `--all`**: među `Run*` procedurama nisu sve testovi
(`RunSelfUpdate`, `RunGoogleAuthSetup`), a deo traži mrežu ili live SEF nalog.

## Provera bez Excela

```bash
python3 tools/run_vba.py --self-test
```

Radi svuda (i u Claude Code sesiji na Linuxu). Proverava da strip VBA header-a ne
propušta header u kod — greška koja je jednom već prošla neopaženo i videla se
tek kao `[break]` u naslovu VBE prozora na Windows mašini.

## Ako zapne

- Skripta ima **tvrdi prekid** (`--timeout`, default 600 s): ako Excel prestane da
  odgovara, proces se ubija i run pada sa `FATAL`. Bez toga COM poziv ne puca —
  samo stoji zauvek.
- **Break mode** (`[break]` u naslovu VBE prozora) znači da je VBA stao usred
  izvršavanja. Najčešći uzrok je bio pokvaren import (header u kodu → naredba
  `End`), a drugi klik na **Debug** u dijalogu greške. Watchdog sada nikad ne
  klika `Debug` ni `Help` — bira `End` / `OK` / `Cancel`.
- `Ctrl+C`, pa proveri Task Manager za zaostali `EXCEL.EXE` — skripta pokreće
  zaseban proces (`DispatchEx`), pa ne dira tvoj otvoreni Excel.
- `--keep` ostavlja temp kopiju sveske da možeš da je otvoriš i vidiš stanje.
- Ako `import` prijavi puno `SKIP ... (nema komponente ...)` nad **pravom**
  sveskom — to znači da sveska nema te listove, tj. da nije AgriX radna sveska.

## Pisanje testa

- **Nov test:** `RunOne n` u `RunAllTests`, plus grana u `TestName` i `InvokeTest`.
  Poziv je direktan (ne `Application.Run`) da bi VBA morao da kompajlira i test i
  sve što on referencira — odatle stiže compile signal.
- **Forma bez prikaza:** `Set f = New frmOtkup`, pa odmah `f.Controls.Count` (bez
  toga se `Initialize` ne okine). Bez `.Show`. `modTestMode.SetTestMode True` gasi
  sve što čeka operatera; kad naiđeš na `MsgBox`/`InputBox` na testiranoj putanji,
  gard ide istim oblikom.
- **Čišćenje ide u `EH` granu, ne na zelenu putanju.** `CleanupPosleTesta` se zove
  iz `EH`, a `Err` se čita **pre** njega (`OtkupUI_Release` je pod
  `On Error Resume Next`, što briše `Err`). Test koji padne inače ostavlja `mFrm`,
  keš i aktivnu otpremnicu sledećem testu — i onda jedan uzrok daje dva pada.
- **Polja se postavljaju kroz `ApplyPrefill`**, ne pisanjem u kontrolu: direktan
  upis u `fgDatum` okine `OnDatumChanged`, a on traži stanica-lock i predlog broja
  **sa pitanjem Google-u** — mreža u testu.
- **Golden za novi UI ne postoji i ne treba.** `DumpKontrole` nad `frmOtkupUI`
  uhvatio bi i `titDatum` (`FmtDatumPun(Now)`), pa bi golden padao svakog sledećeg
  dana. Legacy forma ima fiksne `.frx` kontrole i tu je snapshot smislen.

Test seam-ovi koje produkcioni kod nosi zbog ovoga (`Public` umesto `Private`,
`IsTestMode` gardovi oko `SetFocus`, `Scr_OtpTestSet`) popisani su u
`.claude/rules/testovi.md` §4 — tamo, jer ih mora videti i onaj ko menja formu, a
ne samo onaj ko piše test.

## Fixture i golden

`tests/fixtures/otkup_test.xlsm` je lokalan artefakt (`.gitignore`), pravi ga
`tools/make_fixture.py` iz **donor** sveske.

> **Kad se `SEED` dict u `make_fixture.py` promeni, fixture se MORA
> regenerisati** — inače testovi padaju na podacima kojih nema. Donor može biti i
> **postojeći fixture**: on nosi punu šemu i nema VBA, a generator ionako briše sve
> redove pre sejanja. Izlaz mora biti druga putanja (donor = izlaz se odbija), pa se
> fajl posle premesti:
>
> ```powershell
> python tools\make_fixture.py --donor tests\fixtures\otkup_test.xlsm --out tests\fixtures\otkup_test_new.xlsm --force
> ```
>
> **`--out` je tu obavezan.** Generator odbija donor koji je isti fajl kao
> izlaz, pa `--donor testsixtures\otkup_test.xlsm --force` bez `--out`
> ne radi — ta komanda je već dva puta napisana u pregledima kao da radi.

- Donor daje samo strukturu; spisak kolona se **ne** zakucava u Python (šema
  tabela je izvor istine). Podaci su 100% sintetički, u transakciji koja se uvek
  poništava — nijedan klijentski podatak ne može da završi u golden fajlu na
  GitHub-u.
- Generator **uklanja sav VBA kod iz donora**: modul zaostao iz starijeg donora se
  izvršava i, ako nosi `Public` ime koje postoji i u svežem kodu, daje „Ambiguous
  name" → `Cannot run the macro`, poruka koja ne liči na compile grešku. Za sveske
  kroz `--workbook` ne briše ništa, nego prijavljuje `ORPHAN` red.
- Šemu donora ispisuje `tools/dump_schema.py` (samo čitanje).

### Potpis fixture-a — zašto `git checkout` nije dovoljan

Fixture je gitignored, pa ga **`git checkout` ne menja**. Posle prelaska na drugu
granu na disku ostaje sveska prethodne: testovi padaju **na podacima**, a pad
izgleda kao regresija koda. To je već pojelo pola sata trijaže — četiri crvena
testa nad ispravnim kodom.

Zato generator pored sveske ostavlja `tests/fixtures/otkup_test.sig` sa hash-om
posejanih podataka, a `run_vba.py` ga poredi **pre podizanja Excela**:

| Stanje | Šta radi `run_vba` |
|---|---|
| potpis se slaže | tiho nastavlja |
| potpis se razlikuje | **staje, exit 2**, ispiše komandu za regeneraciju |
| potpisa nema | **staje, exit 2** — sveska od generatora pre ovog sistema |
| `--workbook` (tuđa sveska) | ne dira ništa |
| `--ignore-fixture-sig` | nastavlja svesno, uz poruku |

Odsustvo potpisa je **fail-closed** namerno: sveska bez njega ne može da se
proveri, a to je baš prvi run na svakoj zatečenoj mašini — tačno onaj u kome se
incident i desio. Upozorenje bi ga propustilo jednom po mašini, što je isto kao
da provere nema. Prazna auto-sveska ovde ne stiže; nju pravi grana iznad poziva.

Potpis pokriva **samo deklarativne podatke** (`SEED`, config, `FIXTURE_DATE`,
`KEEP_ROWS`). Izmena logike upisa (`add_row`, `strip_rows`, `upsert_config`) ili
sadržaja tabela koje se čuvaju iz donora (`KEEP_ROWS`) mu je **nevidljiva**. Za
to postoji ručna poluga — `FIXTURE_FORMAT_VERSION` u `make_fixture.py`, koja
ulazi u hash: kad se promeni semantika generatora, podigne se za jedan i svi
fixture-i postaju ustajali. Jeftinije i tačnije nego hashirati ceo `.py`, koji bi
tražio regeneraciju i na izmenu komentara.

Potpis se **briše na početku** build-a i piše tek posle uspešnog `Save`: neuspeo
build (donor bez kolone → `SEMA:`) prepiše svesku, pa bi zadržan stari `.sig`
tvrdio da je fixture svež. Bolje „nema potpisa" nego lažan.

Svestan run nad starim podacima: `--ignore-fixture-sig`. To je jedini način da
run prođe bez važećeg potpisa — nema tihe grane.

`tests/golden/*.txt` idu u git. Kad golden ne postoji, test ga upiše i **padne** —
nov golden mora proći ljudski pregled pre nego što postane merilo. Dva pravila:
**ASCII** (`DumpKontrole` escape-uje dijakritiku u `\uXXXX`; VBA `Print #` piše u
ANSI stranu koja `ć` nema) i **LF** (`.gitattributes` drži `eol=lf`, inače suite
pada na svakom svežem klonu na Windows-u).

## Sabotaža — kako se radi

Kada je obavezna, pravilo je u `.claude/rules/testovi.md` §6. Mehanika:

```bash
python tools/sabotaza.py --lista
python tools/sabotaza.py clear-datum          # primeni jednu
python tools/run_vba.py --suite RunAllTests   # ocekuj FAIL po IMENU tog testa
python tools/sabotaza.py --vrati              # vrati
```

Koja sabotaža obara koji test i sa kojom tvrdnjom — **`--lista`**, ne prepisivati
nigde; skripta je izvor istine.

Za legacy formu radi se ručno u `ClearOtkupFields` (dodaj `txtDatum.value = ""`,
`txtBrojZbirne.value = ""`, ukloni `cmbKooperant.value = ""`), revert je
`git checkout -- src-vba/frmOtkup.frm`.

> **Kad sabotaža obori VIŠE testova, to ne mora biti curenje stanja.**
> `zbirna-vozac` i `prijemnica-kupac` obaraju i `T_ScrSave_RutaPoRezimu`, jer taj
> test dokazuje rutu time što prazan dokument staje na **prvom pravilu svog tipa**.
> `blok-ostatak-snapshot` obara **tri** (kapija, put unosa, ruta), a `blok-tudj-om`
> **dva** (kapija i writer) — isto pravilo je namerno provereno na više nivoa, pa
> njegovo uklanjanje mora da se vidi na svakom.
> Razlika u odnosu na pravo curenje je merljiva: svaki pad ima **svoju poruku i
> svoju tvrdnju**, a ne `Err.Number=0` sa praznim opisom. Prvo proveri izolaciju,
> pa tek onda proizvod.

> **Sidro sabotaže je deo koda koji sabotira.** `clear-zbirna` se razvezalo čim je
> `ClearForm` dobio `fgNovac` u spisak polja — skripta bi tiho prijavila „sidro
> nije jednoznačno" tek pri sledećem pokretanju, a do tada bi izgledalo da je dokaz
> i dalje važeći. **Kad menjaš red koji je nečije sidro, promeni i sidro, pa ponovo
> pokaži crveno.**

Tri zamke koje skripta rešava, i koje važe za svaki sličan zahvat nad izvorom:

1. **Kraj reda** — `src-vba` je CRLF na Windows-u, LF na Linuxu. Sidro sa zakucanim
   `\n` ne pogodi ništa, skripta tiho ne uradi ništa, run prođe nad neizmenjenim
   fajlom i izgleda kao da sabotaža „nije oborila" suite. Detektuj
   (`nl = '\r\n' if '\r\n' in s else '\n'`) i tvrdi `assert s.count(old) == 1`.
2. **Uvlačenje** — sidro se poredi od početka reda; inače isti niz pogađa dva mesta.
3. **Vraćanje** — `git checkout --` briše i nesnimljene izmene koje sa sabotažom
   nemaju veze (jednom je pojelo test seam-ove). `--vrati` radi obrnutu zamenu.

`parse-cdate` pada na tvrdnji „godina van poslovnog opsega" (`11.08.1899`) —
jedina tvrdnja koja razlikuje `CDate` od determinističkog parsera **na DMY
mašini**. Razliku na MDY mašini ne pokriva nijedan test i to se ne prijavljuje kao
pokriveno.

## Trijaža masovnih padova

```bash
python tools/run_vba.py --suite X --keep
```

Snimi temp kopiju, pa `tools/read_test_log.py <temp>/otkup_test.xlsm` grupiše
padove po temi i razlogu.
