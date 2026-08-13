---
paths:
  - "src-vba/mod*Tests.bas"
  - "src-vba/modTest*.bas"
  - "src-vba/frmOtkup.frm"
  - "tools/vba_check.py"
  - "tools/run_vba.py"
  - "tools/make_fixture.py"
  - "tools/dump_schema.py"
  - "tests/golden/*"
---

<!-- frmOtkup.frm je u paths namerno: u njoj su meta sva tri testa ponasanja
     (ClearOtkupFields), test seam (Public umesto Private) i IsTestMode gard.
     Bez ovoga agent koji menja formu ne bi ni znao da to postoji. -->


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
| `RunAllTests`² | `RunBusinessFlowProSuite` (337) |
| `RunIzvestajTests` | `TestLicense_All`³ (23) |
| `RunSheetsJsonParserTests` | `RunNovacSmokeSuite` (12) |
| `RunBankaImportTestSuite` | `RunProductionHealthCheck` |
| `RunFakturaSmokeSuite` | `TestMonitoring_All` |
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
³ Pada nezavisno od ove suite i nije popravljano — vidi „Zatečeni padovi" u §4.

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

Pa `gate: True` u `SUITES` i u listu u `.claude/hooks/vba-test.sh`.

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

```powershell
python tools/make_fixture.py --donor "<put>\AgriX_2.28.4.xlsm"   # jednom
python tools/run_vba.py --suite RunAllTests                       # samo ove tri
```

### Akceptaciona komanda — gate je ~690 provera, ne tri

`--suite RunAllTests` vrti samo tri nova testa. Pravi gate su **svi gate suite-ovi
iz podrazumevanog seta**, i to je ono što pušta `Stop` hook:

```powershell
python tools/run_vba.py --suite RunAllTests --suite RunIzvestajTests ^
    --suite RunSheetsJsonParserTests --suite RunBankaImportTestSuite ^
    --suite RunFakturaSmokeSuite --suite RunStornoTestSuite ^
    --suite Test_StornoCentar_All --suite RunPaleteTestSuite ^
    --suite RunAgrohemijaSmokeSuite
```

Izmereno na operaterskoj mašini (`EXIT=0`, svih devet zeleno):

| Suite | Provera |
|---|---|
| `RunBankaImportTestSuite` | 189 |
| `RunStornoTestSuite` | 181 |
| `RunPaleteTestSuite` | 97 |
| `Test_StornoCentar_All` | 88 |
| `RunSheetsJsonParserTests` | 72 |
| `RunFakturaSmokeSuite` | 35 |
| `RunAgrohemijaSmokeSuite` | 25 |
| `RunAllTests` | 3 |
| `RunIzvestajTests` | ne prijavljuje broj |
| **ukupno** | **690** + `RunIzvestajTests` |

Sve rade nad **sintetičkim** fixture-om — suite koje diraju tabele seju sebi
podatke u transakciji koja se uvek poništava (`SVT-*`, `BIT-*`, `TST-*`), pa im
prava radna sveska nije potrebna.

**Ostalo u blind stanju: ~35 provera** — `TestLicense_All` (23),
`RunNovacSmokeSuite` (12), plus `RunProductionHealthCheck` i `TestMonitoring_All`.
Recept je iznad, u §3.

### `RunBusinessFlowProSuite` — 147 palih provera, zatečeno

Konvertovana je u `gate` (verdikt u `EndRun`, koji zovu sva četiri `Run*` runnera
tog modula), ali je **van podrazumevanog seta i van `Stop` hook-a** dok se ne
trijažira:

```
Total=310 | Passed=163 | Failed=147
```

**Te provere su padale i ranije** — suite je bila `blind`, pa je runner prijavljivao
„prošla bez greške" dok je skoro polovina padala. Konverzija ih nije napravila nego
otkrila; dokaz: sabotaža jedne provere pomera brojač za tačno `+1` (147 → 148 →
147), dakle brojanje je suite-ovo i nepromenjeno.

Uzrok **nije utvrđen**. Dve hipoteze, obe neproverene: suite traži master podatke
koje sintetički fixture nema (seje svoje kroz `SeedBusinessFlowProMasterData`, ali
može zavisiti i od zatečenog config-a), ili je deo provera stvarno u regresiji.
Trijaža ide kroz Immediate prozor posle
`python tools/run_vba.py --suite RunBusinessFlowProSuite --keep`.

Dok se to ne razreši, suite se pokreće ručno i njen verdikt je vidljiv — ali ne
obara svaku sesiju.

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

Šemu donora ispisuje `tools/dump_schema.py` (samo čitanje, ne dira svesku) —
batch varijanta onoga što `modSetup.DebugKoloneTabele` radi interaktivno.

### Zatečeni padovi u punom setu (NISU iz ove suite)

Jedan jedini: **`TestLicense_All` — „Cannot run the macro".** Makro postoji
(`modLicenseTests.bas:18`, `Public Sub`) i import prolazi bez primedbe, pa je
najverovatnije compile greška u `modLicenseTests` (VBA kompajlira lenjo i odbija
da pokrene makro iz modula koji ne prolazi). **Nije potvrđeno** — za potvrdu treba
`Alt+F11 → Debug → Compile` ručno. Suite je `blind`, pa ni da se pokrene ne bi
davala verdikt; jedina šteta je što obara goli `run_vba.py`.

Zato hook i akceptacija idu kroz **eksplicitnu listu gate suite-ova** (v. gore),
a ne kroz goli poziv. Čim se `TestLicense_All` raščisti, lista se briše i ostaje
goli `run_vba.py`.

> `RunBankaImportTestSuite` je bio drugi pad na prvom pokretanju (`T13`,
> `PASS=186 FAIL=1`) i **rešen je u #183** — pao je test vektor, ne produkcija:
> `600.005` se u `Double`-u čuva ispod pola pare, pa ga zaokruživanje korektno
> spušta na `600.00`. Ne vodi se više kao zatečen pad.

### Stop hook

`.claude/hooks/vba-test.sh` pušta suite na kraju sesije kad je `src-vba/` diran (u
radnom stablu ili u poslednjem commit-u). Bez `pywin32`/Excela prolazi **tiho** —
u Linux sesiji ostaje samo `vba_check` kroz PostToolUse.

## 5) Šta i dalje ostaje na operateru

Finalni smoke-test u Excelu (klik po klik). Zato svaki rad završi kratkom,
numerisanom test-checklistom u chatu — vidi `.claude/rules/git-i-release.md`.
