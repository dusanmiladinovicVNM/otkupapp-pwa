---
paths:
  - "src-vba/mod*Tests.bas"
  - "src-vba/modTest*.bas"
  - "tools/vba_check.py"
  - "tools/run_vba.py"
  - "tools/make_fixture.py"
  - "tools/dump_schema.py"
  - "tests/golden/*"
---

# Verifikacija: šta se stvarno može proveriti

> CLAUDE.md §5 kaže „CI ne pokreće Excel". To i dalje važi, ali od sada postoje
> dva alata koja pomeraju granicu — jedan radi svuda, drugi samo na Windows
> mašini sa Excelom.

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

Import `src-vba/` → `Debug > Compile` → test suite, headless. U ovoj sesiji se
**ne može pokrenuti** (nema COM-a); to je alat za operatera i za Windows dev
mašinu.

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
| `RunAllTests`² | `Test_StornoCentar_All` |
| `RunIzvestajTests` | `TestLicense_All`³ |
| `RunSheetsJsonParserTests` | `RunStornoTestSuite` |
| `RunBankaImportTestSuite` | `RunPaleteTestSuite` |
| `RunFakturaSmokeSuite` | `RunNovacSmokeSuite` |
| `RunGoogleSyncSmokeSuite`¹ | `RunBusinessFlowProSuite` |
| `RunMasterSyncSmokeSuite`¹ | `RunAgrohemijaSmokeSuite` |
| `RunSEFTestSuite`¹ | `RunProductionHealthCheck`, `TestMonitoring_All` |

¹ Nije u podrazumevanom setu — traži mrežu ili live SEF nalog.
² Verdikt ne dolazi iz toga da li `Run()` pukne — `modTest` hvata grešku po testu
da jedan pad ne obori ostale — nego iz `last_run.txt` pored sveske. Nema fajla =
pad. Vidi §4.
³ Pada nezavisno od ove suite i nije popravljano — vidi „Zatečeni padovi" u §4.

Kad pišeš NOVU suite, napravi je `gate` (`Err.Raise` na pad) i upiši je u
`SUITES` katalog u `tools/run_vba.py`. Nova „blind" suite je test koji niko neće
videti kad pukne.

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
python tools/run_vba.py --suite RunAllTests                       # exit 0 / 2
```

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

Prvo pokretanje punog seta dalo je dva pada nezavisna od `modTest`. Jedan je rešen,
jedan stoji:

- ~~`RunBankaImportTestSuite` — `T13`~~ **REŠENO u #183.** Pao je **test vektor, ne
  produkcija**: `600.005` se u `Double`-u čuva kao `600.00499999...`, dakle ispod
  pola pare, pa ga half-up zaokruživanje korektno spušta na `600.00`.
  `ValidateNalogSaldo` je bio u pravu. Vektor je zamenjen jednoznačnim `600.006`, a
  za vrednost tačno na pola pare se sada tvrdi invarijanta (iznos u fajlu ne prelazi
  otvoreno) umesto smera zaokruživanja.
- `TestLicense_All` — „Cannot run the macro". Makro **postoji**
  (`modLicenseTests.bas:18`, `Public Sub`) i import prolazi bez primedbe, pa je
  najverovatnije compile greška u `modLicenseTests` (VBA kompajlira lenjo, pa
  odbija da pokrene makro iz modula koji ne prolazi). **Nije potvrđeno** — za
  potvrdu treba `Alt+F11 → Debug → Compile` ručno.

Dok `TestLicense_All` stoji, akceptaciona komanda za ovu suite je
`--suite RunAllTests`, a ne goli poziv.

**Kad se i to reši**, u `.claude/hooks/vba-test.sh` se `--suite RunAllTests` menja
golim pozivom i hook počinje da vrti ceo podrazumevani set — blizu 300 provera pod
gate-om umesto tri. To je jedan red izmene i glavni razlog da se `TestLicense_All`
ne ostavi da visi.

### Stop hook

`.claude/hooks/vba-test.sh` pušta suite na kraju sesije kad je `src-vba/` diran (u
radnom stablu ili u poslednjem commit-u). Bez `pywin32`/Excela prolazi **tiho** —
u Linux sesiji ostaje samo `vba_check` kroz PostToolUse.

## 5) Šta i dalje ostaje na operateru

Finalni smoke-test u Excelu (klik po klik). Zato svaki rad završi kratkom,
numerisanom test-checklistom u chatu — vidi `.claude/rules/git-i-release.md`.
