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

<!-- frmOtkup.frm je u paths namerno: u njoj su meta sva tri testa ponasanja
     (ClearOtkupFields), test seam (Public umesto Private) i IsTestMode gard.
     Bez ovoga agent koji menja formu ne bi ni znao da to postoji.
     Isto vazi za modOtkupUI.bas (ClearForm / ParseDatum / ParcelaID su Public
     zbog testa, dva SetFocus-a su iza IsTestMode) i modScrDokumenti.bas
     (Scr_OtpTestSet postoji samo za test i tvrdo je gejtovan). -->


# Verifikacija: šta se stvarno može proveriti

> CI i dalje ne pokreće Excel. Postoje dva alata: jedan radi svuda i gleda izvor,
> drugi traži Windows + Excel i gleda **ponašanje**. Definicija gotovog je u
> CLAUDE.md §5 — zeleno nad ispravnim **i dokazano crveno** nad namerno pokvarenim
> kodom, izlaz priložen.

## 1) `tools/vba_check.py` — radi i u Claude Code sesiji (Linux/macOS)

```bash
python3 tools/vba_check.py                 # sve nad src-vba/
python3 tools/vba_check.py src-vba/modX.bas
python3 tools/vba_check.py --hook          # tiho kad je čisto
```

Exit `0` = čisto, `2` = ima nalaza. Mehanizuje pravila iz §4 koja je do sada
hvatao tek operater u VBE-u:

| Provera | Šta hvata | Odakle pravilo |
|---|---|---|
| `ASCII` | ne-ASCII bajt u `.bas`/`.cls`/`.frm`/`.doccls` | §4 encoding (`f08a0ee`) |
| `DEKLARACIJA` | modul-level `Const`/`Dim`/`Declare`/`Type`/`Enum` posle prve procedure | §4 (RF-07) |
| `REZERVISANO` | ime koje se case-insensitive poklapa sa VBA ključnom reči | §4 (RF-06, `eNum`) |
| `DUPLIKAT` | isti `Public Sub/Function/Const` u dva modula → „Ambiguous name" | §4 merge |
| `PORUKA` | `Poruka("KLJUC")` bez para u `modPoruke.UpsertPoruke` | §4 katalog |
| `NEDEFINISAN` | poziv procedure koja nigde nije definisana → „Sub or Function not defined" | compile |
| `ARNOST` | poziv sa pogrešnim brojem argumenata → „Wrong number of arguments" | compile |

Poslednje dve pokrivaju **dve najčešće compile greške** u ovom projektu — one
zbog kojih je i pravljen headless compile gate koji se nije dao ukrotiti. Ovde se
hvataju bez Excela. Namerno su uske (samo `.bas`, samo poziv u poziciji naredbe),
jer je lažan nalaz u hook-u gori od propuštenog.

**Ovo je obavezan korak pre commita** svake VBA izmene. Ne prijavljuj izmenu kao
gotovu dok `vba_check` nije zelen — to je jedini deo §5 koji više nije „na
poverenje".

Ograničenje: ne kompajlira VBA. Ne hvata tip-greške, pogrešnu arnost, nepostojeći
simbol. Za to i dalje treba VBE.

## 2) `tools/run_vba.py` — SAMO Windows + Excel + `pywin32`

Import `src-vba/` → `Debug > Compile` → test suite, headless. Traži COM, pa se
**ne pokreće na Linux/macOS** — ni u Claude Code sesiji na webu, ni u GitHub
Actions. Tamo se testovi ponašanja **ne izvršavaju uopšte**, a `Stop` hook prolazi
tiho: sesija može da se završi „zeleno" bez ijedne provere ponašanja. To nije
verifikacija i ne sme se tako prijaviti — VBA izmena koja dira ponašanje ide na
Windows mašinu ili ostaje neverifikovana.

Jedini deo koji radi svuda: `python3 tools/run_vba.py --self-test` — provera da
strip VBA header-a ne propušta header u kod. **Pokreni ga posle svake izmene
`_read_code_body`/import logike**; ta greška je jednom prošla neopaženo i videla
se tek kao `[break]` u VBE-u.

```powershell
python tools\run_vba.py --compile-only     # najbrže i najstabilnije
python tools\run_vba.py                    # + podrazumevani set suite-ova
python tools\run_vba.py --suite RunBankaImportTestSuite
python tools\run_vba.py --all
```

Ne zove `modVbaTools.ImportAllVBA` (hardkodiran folder + završni `MsgBox` = smrt
za headless), nego ponavlja njegovu logiku preko COM-a. Modalne dijaloge zatvara
watchdog.

Sveska: bez `--workbook` ide `tests/fixtures/otkup_test.xlsm`. Ako ga nema, skripta
napravi **praznu** `.xlsm` — dovoljno za compile, ali **ne i za suite**: prazna
sveska nema tabele, pa `New frmOtkup` pukne na 1004 već u `UserForm_Initialize`
(`FillCmb ... TBL_KULTURE` i `FillComboDisplayID ... TBL_STANICE` nisu pod
`On Error`). Za suite napravi pravi fixture — `tools/make_fixture.py`, §4.
Original se nikad ne dira — radi se nad temp kopijom.
Detalji: `docs/EXCEL_TEST_HARNESS.md`.

Compile verdikt **ne obara** run kad suite-ovi idu: `COMPILE NEJASNO` se i dalje
ispisuje, ali odgovor nose testovi (da bi se `RunAllTests` uopšte pokrenuo, VBA
mora da kompajlira `modTest` i sve što on referencira). Eksplicitan compile
`FAIL` i dalje pada, kao i `NEJASNO` uz `--compile-only`, gde je probe jedini
izvor istine.

## 3) `gate` vs „blind" suite — bez ovoga se lako pogrešno zaključi

Suite sa `gate: True` **podiže grešku** kad provera padne, pa je runner vidi kao
crvenu. Suite sa `gate: False` rezultat piše samo u Immediate prozor — runner je
prijavljuje kao **`blind`**, što znači „prošla bez greške", a **NE** „sve provere
prošle".

| gate (crveno se vidi) | blind (rezultat samo u Immediate) |
|---|---|
| `RunAllTests`² | `RunNovacSmokeSuite` (12) |
| `RunIzvestajTests` | `RunProductionHealthCheck` |
| `RunSheetsJsonParserTests` | `TestMonitoring_All` |
| `RunBankaImportTestSuite` | |
| `RunFakturaSmokeSuite` | |
| `RunStornoTestSuite` | |
| `Test_StornoCentar_All` | |
| `RunPaleteTestSuite` | |
| `RunAgrohemijaSmokeSuite` | |
| `RunGoogleSyncSmokeSuite`¹ | |
| `RunMasterSyncSmokeSuite`¹ | |
| `RunSEFTestSuite`¹ | |

¹ Nije u podrazumevanom setu — traži mrežu ili live SEF nalog.
² Verdikt ne dolazi iz toga da li `Run()` pukne — `modTest` hvata grešku po testu
da jedan pad ne obori ostale — nego iz `last_run.txt` pored sveske. Nema fajla =
pad. Vidi §4.

Kad pišeš NOVU suite, napravi je `gate` (`Err.Raise` na pad) i upiši je u
`SUITES` katalog u `tools/run_vba.py`. Nova „blind" suite je test koji niko neće
videti kad pukne.

### Kako se blind prevodi u gate

Blind suite **već broji** padove (`mFail`) — samo ne podiže grešku, pa runner
vidi „prošlo bez greške". Konverzija je tri linije, po uzoru na
`modTestBanka.ERR_BIT_SUITE_FAILED`:

1. `Private Const ERR_X_SUITE_FAILED As Long = vbObjectError + <slobodan>` u
   deklaracionu sekciju (zauzeti offseti: 2900, 2950, 2960–2963, 3010–3012, 3100)
2. posle završnog izveštaja: `If mFail > 0 Then Err.Raise ERR_X_SUITE_FAILED, ...`
3. u `EH`: prebroj prekid kao pad (`Fail "SUITE prekinut..."`) pa podigni —
   prekinuta suite nije „nije se desilo" nego pad

Pa `gate: True` i `default: True` u `SUITES` katalogu — hook se **ne dira**, jer
pušta goli `run_vba.py` i katalog je jedini izvor istine.

**Konverzija nije gotova bez dvosmernog dokaza** (§5): obori namerno jednu proveru
(`Chk False, "SABOTAZA"`), pokaži `exit 2` sa imenom te suite, pa vrati i pokaži
zeleno.

> **Zamka pri pisanju sabotaže skriptom:** `src-vba` se na Windows-u checkout-uje
> kao **CRLF**, a na Linuxu kao LF. Sidro sa zakucanim `\n` neće pogoditi ništa i
> skripta tiho ne uradi ništa — pa run prođe nad neizmenjenim fajlom i izgleda
> kao da sabotaža „nije oborila" suite. Uvek detektuj:
> `nl = '\r\n' if '\r\n' in s else '\n'`, i tvrdi `assert s.count(old) == 1`.

## 4) `modTest` — suite koja pada na PONAŠANJU, ne na sintaksi

`vba_check` hvata sintaksu; `modTest` hvata izmenu koja se uredno kompajlira, a
menja ponašanje. Tri testa nad `frmOtkup.ClearOtkupFields` (tu bug i živi):

| Test | Šta drži |
|---|---|
| `T_PosleSnimanja_ZadrzavaKontekstOtpremnice` | datum se posle snimanja NE briše (+ pun snapshot forme) |
| `T_PosleSnimanja_ZadrzavaZbirnu` | broj zbirne ostaje, i drugi blok dobija istu zbirnu |
| `T_ClearForm_BrisePartnera` | `cmbKooperant` se BRIŠE (obrnut smer od prva dva) |

I tri nad novim UI-jem (`modOtkupUI`), jer legacy se **ne gasi** — obe kopije
postoje namerno (`docs/UI_MIGRACIJA_KATALOG.md`), pa obe nose svoj test:

| Test | Šta drži |
|---|---|
| `T_ParseDatum_Ugovor` | prazno/necitljivo je `0`; `d.m.yyyy` se čita kao DMY bez `CDate`; trailing tačka se skida; nemoguć datum se **odbija**, ne preliva (`30.02` → `2.3`, mesec 13 → januar sledeće godine) |
| `T_ParcelaID_IzSkriveneKolone` | ID parcele dolazi iz **skrivene druge kolone**, ne iz prikaznog teksta; sakriveno polje ne šalje parcelu u dokument |
| `T_ClearForm_Ugovor` | ista tri ponašanja kao legacy trojka (datum ostaje, zbirna ostaje, partner se briše) + razlika novog UI-ja: **bez** aktivne otpremnice datum se vraća na danas |

**Šta ovi testovi NE pokrivaju:** ništa iznad `ClearForm` — put upisa (`modOtkupUnos`
/ `modDokUnos`), mreža, storno. Forma se gradi bez `.Show` (`New frmOtkupUI` pa
`Controls.count`), pa `UserForm_Activate` — raspored, `GoFullScreen`, punjenje
mreže — nikad ne ide.

### Tri seam-a koja novi UI nosi zbog ovih testova

Isti oblik kao u `frmOtkup` (§2 u `.claude/rules/otkup-i-dokumenta.md`):

- `ClearForm` / `ParseDatum` / `ParcelaID` su **`Public`**, ne `Private` — test ih
  zove direktno, bez vožnje celog upisa (stanica-lock, PDF, auto-lanac hladnjače).
- **Tri `SetFocus`-a** (dva u `ClearForm`, jedan na kraju `ApplyPrefill`) su iza
  `If Not IsTestMode()`. Forma koja nije `.Show`-ovana ne može da primi fokus, a u
  nevidljivom Excelu `SetFocus` **ne puca nego trajno visi**. U produkciji je
  `IsTestMode()` uvek `False`.
- `modScrDokumenti.Scr_OtpTestSet` — suprotan smer od `Scr_OtpOtkazi` i **jedini**
  način da test dobije aktivnu otpremnicu (produkcija je bira klikom na red, što
  traži učitanu mrežu). Tvrdo gejtovan: van test-režima ne radi ništa.

Test polja postavlja kroz `ApplyPrefill`, ne pisanjem u kontrolu: direktan upis u
`fgDatum` okine `OnDatumChanged`, a on traži stanica-lock i predlog broja **sa
pitanjem Google-u** — mreža u testu. `ApplyPrefill` je isti put kojim polja stižu
i u produkciji (izbor otpremnice) i jedini koji ide pod `mLoading`.

**Golden snapshot za novi UI ne postoji i ne treba da postoji.** `DumpKontrole`
nad `frmOtkupUI` uhvatio bi i `titDatum` (`FmtDatumPun(Now)`), pa bi golden padao
svakog sledećeg dana. Legacy forma ima fiksne `.frx` kontrole i tu je snapshot
smislen; runtime forma sa vremenom u natpisu nije ista stvar.

```powershell
python tools/make_fixture.py --donor "<put>\AgriX_2.28.4.xlsm"   # jednom
python tools/run_vba.py --suite RunAllTests                       # samo ove tri
```

### Akceptaciona komanda — goli poziv, ~1050 provera

`--suite RunAllTests` vrti samo `modTest` (tri testa nad legacy formom + tri nad
novim UI-jem). Pun gate je ceo podrazumevani set, i to je ono što pušta `Stop`
hook:

```powershell
python tools/run_vba.py
```

Izmereno: `EXIT=0`, 11 suite-ova, i **bez `BLIND` reda u ispisu** — u
podrazumevanom setu nema više nijedne suite bez verdikta.

Nema više eksplicitne liste: katalog `SUITES` u `tools/run_vba.py` je jedini izvor
istine. Nova suite ulazi u gate time što je upisana tamo sa `default: True` —
hook se ne dira. **Ne proširivati na `--all`**: među `Run*` procedurama nisu sve
testovi (`RunSelfUpdate`, `RunGoogleAuthSetup`), a deo traži mrežu ili live SEF
nalog.

Izmereno na operaterskoj mašini (`EXIT=0`, svih devet zeleno):

| Suite | Provera |
|---|---|
| `RunBankaImportTestSuite` | 189 |
| `RunStornoTestSuite` | 181 |
| `RunPaleteTestSuite` | 97 |
| `Test_StornoCentar_All` | 88 |
| `RunSheetsJsonParserTests` | 72 |
| `RunFakturaSmokeSuite` | 35 |
| `RunBusinessFlowProSuite` | 336 |
| `RunAgrohemijaSmokeSuite` | 25 |
| `TestLicense_All` | 23 |
| `RunAllTests` | 3 |
| `RunIzvestajTests` | ne prijavljuje broj |
| **ukupno** | **~1050** + `RunIzvestajTests` |

> Merenje je **starije od tri testa novog UI-ja** (`RunAllTests` ih sada vrti
> šest). Brojevi se ne prepravljaju napamet — red se ispravlja tek posle
> pokretanja na Windows mašini.

Sve rade nad **sintetičkim** fixture-om — suite koje diraju tabele seju sebi
podatke u transakciji koja se uvek poništava (`SVT-*`, `BIT-*`, `TST-*`), pa im
prava radna sveska nije potrebna.

**U podrazumevanom setu nema više nijedne blind suite.** Ostale su van njega:
`RunNovacSmokeSuite` (12), `RunProductionHealthCheck` i `TestMonitoring_All`.
Recept za konverziju je iznad, u §3.

Za trijažu masovnih padova: `run_vba.py --suite X --keep` zadrži temp kopiju i
**snimi je**, pa `tools/read_test_log.py <temp>/otkup_test.xlsm` grupiše padove po
temi i po razlogu. (Bez snimanja bi kopija ostala u stanju pre rana i trijaža bi
čitala stariji, tuđi run — što se jednom i desilo.)

### Šema se mora podići pre suite-ova — inače „regresija" koje nema

Fixture nastaje iz **starijeg donora** (npr. 2.28.4), a kod je noviji. Kolone
dodate u međuvremenu ne postoje dok se ne pokrene schema upgrade, pa suite koje
ih diraju padaju masovno — a izgleda kao regresija u proizvodu.

`RunBusinessFlowProSuite` je na tome davao `Total=310 | Passed=163 | Failed=147`.
Posle `EnsureRuntimeSchema` prolazi **100%**.

Zato `run_vba.py` sada **uvek** pusti `EnsureRuntimeSchema` posle importa a pre
suite-ova, i ispiše `SCHEMA OK` / `SCHEMA FAIL`. Rutina je idempotentna
(`EnsureColumnOnTable` je no-op kad kolona postoji). Pala priprema šeme obara run
i kad su sve suite zelene — rezultati nad nepripremljenom šemom nisu merodavni.

Redosled je bitan i nije proizvoljan: **posle importa**, jer schema pravila
dolaze iz svežeg koda, ne iz onoga što je u donoru. Isto traži i komentar na vrhu
`modTestStornoCentar`.

> Ovo je bio i najskuplji promašaj u ovom poslu: 147 padova je izgledalo kao
> nalaz o proizvodu, a bila je nepripremljena sveska.

### „Suite se nije pokrenuo" nije „prošlo"

Uz konverziju para palete/agrohemija zatvorene su **četiri putanje lažnog
zelenog** — tihi `Exit Sub` pre nego što ijedna provera krene, koji je runner
video kao `OK`:

| Suite | Uslov koji je tiho izlazio |
|---|---|
| `RunPaleteTestSuite` | paletiranje isključeno u Podešavanjima |
| `RunPaleteTestSuite` | zatečen `TST-` ostatak od prekinutog run-a |
| `RunPaleteTestSuite` | operater odustao na potvrdi |
| `RunAgrohemijaSmokeSuite` | dev-guard odbijen |

Sve sada podižu grešku sa porukom koja počinje `suite NIJE pokrenut:`, pa se u
ispisu razlikuje od pale provere. **Kad pišeš ili konvertuješ suite, prođi i rane
izlaze** — pala provera je glasna, a suite koji se nije ni pokrenuo je tih.

Ista rupa i dalje stoji u `RunBankaImportTestSuite` (rani `Exit Sub` kad
`tblBankaImport`/`tblOtkup` ne postoje) — poznata, nije zatvorena.

Izostavljena su tačno dva: `TestLicense_All` (ne može da se pokrene — v. dole) i
`Test_StornoCentar_All` (blind, rezultat samo u Immediate, troši vreme bez
verdikta). Čim se `TestLicense_All` raščisti, cela lista se briše i ostaje goli
`python tools/run_vba.py`.

**Do ove verzije nijedna od tih suite nije se pokretala kroz `run_vba.py`
uopšte** — compile probe je vraćao `NEJASNO`, `rc = 2` je padao pre suite petlje i
petlja se nikad nije dosegla. Suite su postojale samo kao ručni `Alt+F8`.

**Dokazano u oba smera** (bez toga suite ne znači ništa — vidi PR #181, četiri
puta zeleno-ali-nedokazano-crveno). Sabotaža se radi u `ClearOtkupFields`, revert
je `git checkout -- src-vba/frmOtkup.frm`:

| Sabotaža | Očekivano |
|---|---|
| dodaj `txtDatum.value = ""` | `FAIL T_PosleSnimanja_ZadrzavaKontekstOtpremnice` |
| dodaj `txtBrojZbirne.value = ""` | `FAIL T_PosleSnimanja_ZadrzavaZbirnu` |
| ukloni `cmbKooperant.value = ""` | `FAIL T_ClearForm_BrisePartnera` |

Svaka sabotaža obara i snapshot iz prvog testa — to je namerno, snapshot hvata i
polja koja niko nije tražio da se provere.

### Sabotaže novog UI-ja — `tools/sabotaza.py`

Sedam sabotaža nad `modOtkupUI`, svaka obara **tačno jedan** test i po imenu:

```bash
python tools/sabotaza.py --lista
python tools/sabotaza.py clear-datum          # primeni jednu
python tools/run_vba.py --suite RunAllTests   # ocekuj FAIL po imenu
python tools/sabotaza.py --vrati              # vrati
```

| Sabotaža | Šta kvari | Očekivano |
|---|---|---|
| `parse-tacka` | ukloni skidanje trailing tačke | `FAIL T_ParseDatum_Ugovor` |
| `parse-cdate` | vrati `IsDate`/`CDate` umesto determinističkog parsera | `FAIL T_ParseDatum_Ugovor` |
| `parcela-tekst` | čitaj ID iz prikaznog teksta (`CB.text`) | `FAIL T_ParcelaID_IzSkriveneKolone` |
| `parcela-vidljivost` | ukloni proveru vidljivosti polja | `FAIL T_ParcelaID_IzSkriveneKolone` |
| `clear-datum` | vraćaj datum na danas i uz aktivnu otpremnicu | `FAIL T_ClearForm_Ugovor` |
| `clear-zbirna` | dodaj `fgBrZbir` u listu polja koja se prazne | `FAIL T_ClearForm_Ugovor` |
| `clear-partner` | prestani da brišeš partnera | `FAIL T_ClearForm_Ugovor` |

`parse-cdate` pada na tvrdnji „godina van poslovnog opsega" (`11.08.1899`) —
namerno, jer je to jedina tvrdnja koja razlikuje `CDate` od determinističkog
parsera **na DMY mašini**. Ostale (`01.02.2026`, `30.02.2026`) na operaterskoj
mašini daju isti rezultat u oba slučaja; tu razliku bi pokazala tek MDY mašina.
To se ne prijavljuje kao pokriveno.

**Tri zamke koje skripta rešava** (sve tri su već jednom ujele, treća u ovoj
sesiji):

1. **Kraj reda** — `src-vba` je CRLF na Windows-u, LF na Linuxu. Sidro sa
   zakucanim `\n` ne pogodi ništa, skripta tiho ne uradi ništa, run prođe nad
   neizmenjenim fajlom i izgleda kao da sabotaža „nije oborila" suite.
2. **Uvlačenje** — sidro se poredi **od početka reda**. Bez toga je
   `    mFrm...cbKupac.value = ""` (4 razmaka) podniz istog reda uvučenog za 8, pa
   je isto sidro pogađalo dva različita mesta.
3. **Vraćanje** — `git checkout --` vraća fajl na `HEAD`, pa briše i nesnimljene
   izmene koje sa sabotažom nemaju veze. `--vrati` radi **obrnutu zamenu**.

**Kako se dodaje test:** `RunOne n` u `RunAllTests`, plus grana u `TestName` i
`InvokeTest`. Poziv je direktan (ne `Application.Run`) da bi VBA morao da
kompajlira i test i sve što on referencira — odatle stiže compile signal.

**Forma bez prikaza:** `Set f = New frmOtkup`, pa odmah `f.Controls.Count` (bez
toga se `Initialize` ne okine). Bez `.Show`. `modTestMode.SetTestMode True` gasi
sve što čeka operatera — trenutno `SetFocus` u `ClearOtkupFields`; kad naiđeš na
`MsgBox`/`InputBox` na testiranoj putanji, gard ide istim oblikom.

### Golden snapshot-i

`tests/golden/*.txt`, idu u git. Kad golden ne postoji, test ga upiše i **padne** —
nov golden mora proći ljudski pregled pre nego što postane merilo. Dva pravila
koja su već jednom slomila suite:

- **ASCII.** `DumpKontrole` escape-uje dijakritiku u `\uXXXX`. VBA `Print #` piše u
  ANSI kodnu stranu koja `ć` nema, pa bi round-trip bio gubitav a poruka o razlici
  besmislena („golden [Vrsta voca] vs tekuci [Vrsta voca]").
- **LF.** `.gitattributes` drži `tests/golden/*.txt` na `eol=lf`; bez toga git na
  Windows-u konvertuje u CRLF i suite pada na svakom svežem klonu.

### Fixture

`tests/fixtures/otkup_test.xlsm` je lokalan artefakt (`.gitignore`), pravi ga
`tools/make_fixture.py` iz **donor** sveske. Donor daje samo strukturu — osnovnu
šemu ne pravi nijedan kod (`Ensure*` rutine u `modSetup` samo dodaju kolone na
postojeće tabele), pa bi zakucavanje spiska kolona u Python bilo drugi izvor
istine (CLAUDE.md §4). Podaci su 100% sintetički: nijedan klijentski podatak ne
može da završi u golden fajlu koji ide na GitHub.

Generator uz to **uklanja sav VBA kod iz donora** (standardni/klasni/form moduli;
document moduli ostaju jer se ne mogu ukloniti). Kod u fixture-u je balast — svež
`src-vba/` se uvozi na svakom pokretanju — ali balast koji ujeda: import prepisuje
samo ono što repo **ima**, pa modul zaostao iz starijeg donora ostaje i izvršava
se. Ako nosi `Public` ime koje postoji i u svežem kodu → „Ambiguous name" → VBA
odbija da pokrene makro iz njega, uz poruku `Cannot run the macro` koja ne liči na
compile grešku. Donor 2.28.4 je nosio **131** takav modul.

Za sveske prosleđene kroz `--workbook` driver ih ne briše (tuđa sveska), nego
prijavljuje kao `ORPHAN` red u ispisu.

Šemu donora ispisuje `tools/dump_schema.py` (samo čitanje, ne dira svesku) —
batch varijanta onoga što `modSetup.DebugKoloneTabele` radi interaktivno.

### Zatečeni padovi u punom setu (NISU iz ove suite)

**Nema ih više.** Goli `python tools/run_vba.py` je `EXIT=0` nad celim
podrazumevanim setom. Oba zatečena pada su zatvorena, i nijedan nije bio ono na
šta je ličio:

- **`RunBankaImportTestSuite` / `T13`** (`PASS=186 FAIL=1`) — rešen u #183. Pao je
  **test vektor, ne produkcija**: `600.005` se u `Double`-u čuva ispod pola pare,
  pa ga zaokruživanje korektno spušta na `600.00`.
- **`TestLicense_All`** — „Cannot run the macro". Bio je **zaostali duplikat u
  svesci**, ne compile greška u modulu. Fixture je nasleđivao 131 VBA modul iz
  donora 2.28.4; jedan je nosio ime koje postoji i u svežem kodu → „Ambiguous
  name" → VBA odbija da pokrene makro. Zato je ručno pokretanje prolazilo (druga
  sveska), a driver padao (fixture), dok je `vba_check` bio zelen s pravom
  (duplikata u repou nema). Rešeno tako što `make_fixture.py` sada uklanja SAV kod
  iz donora — vidi §4, „Fixture".

> **Moja hipoteza da je u pitanju compile greška u `modLicenseTests` bila je
> netačna.** Stajala je označena kao nepotvrđena i oborena je pokretanjem. Beleži
> se jer je pokazala pravu pouku: pitanje „nad kojom sveskom" nije se postavljalo
> nigde, pa su tri tačna signala zajedno izgledala kontradiktorno.

### Stop hook

`.claude/hooks/vba-test.sh` pušta suite na kraju sesije kad je `src-vba/` diran (u
radnom stablu ili u poslednjem commit-u). Bez `pywin32`/Excela prolazi **tiho** —
u Linux sesiji ostaje samo `vba_check` kroz PostToolUse.

## 5) Šta i dalje ostaje na operateru

Finalni smoke-test u Excelu (klik po klik). Zato svaki rad završi kratkom,
numerisanom test-checklistom u chatu — vidi `.claude/rules/git-i-release.md`.
