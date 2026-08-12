---
paths:
  - "src-vba/mod*Tests.bas"
  - "src-vba/modTest*.bas"
  - "tools/vba_check.py"
  - "tools/run_vba.py"
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

Sveska: bez `--workbook` ide `tests/fixtures/otkup_test.xlsm`, a ako ga nema,
skripta ga sama napravi kao **praznu** `.xlsm` — za compile je to dovoljno (nijedan
modul ne referencira sheet `CodeName` rano-vezano). Suite-ovima trebaju podaci, pa
im prosledi pravu radnu svesku. Original se nikad ne dira — radi se nad temp
kopijom. Detalji: `docs/EXCEL_TEST_HARNESS.md`.

## 3) `gate` vs „blind" suite — bez ovoga se lako pogrešno zaključi

Suite sa `gate: True` **podiže grešku** kad provera padne, pa je runner vidi kao
crvenu. Suite sa `gate: False` rezultat piše samo u Immediate prozor — runner je
prijavljuje kao **`blind`**, što znači „prošla bez greške", a **NE** „sve provere
prošle".

| gate (crveno se vidi) | blind (rezultat samo u Immediate) |
|---|---|
| `RunIzvestajTests` | `Test_StornoCentar_All` |
| `RunSheetsJsonParserTests` | `TestLicense_All` |
| `RunBankaImportTestSuite` | `RunStornoTestSuite` |
| `RunFakturaSmokeSuite` | `RunPaleteTestSuite` |
| `RunGoogleSyncSmokeSuite`¹ | `RunNovacSmokeSuite` |
| `RunMasterSyncSmokeSuite`¹ | `RunBusinessFlowProSuite` |
| `RunSEFTestSuite`¹ | `RunAgrohemijaSmokeSuite` |
| | `RunProductionHealthCheck`, `TestMonitoring_All` |

¹ Nije u podrazumevanom setu — traži mrežu ili live SEF nalog.

Kad pišeš NOVU suite, napravi je `gate` (`Err.Raise` na pad) i upiši je u
`SUITES` katalog u `tools/run_vba.py`. Nova „blind" suite je test koji niko neće
videti kad pukne.

## 4) Šta i dalje ostaje na operateru

Finalni smoke-test u Excelu (klik po klik). Zato svaki rad završi kratkom,
numerisanom test-checklistom u chatu — vidi `.claude/rules/git-i-release.md`.
