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
     njima su test seam-ovi (Public umesto Private, IsTestMode gardovi,
     Scr_OtpTestSet). Bez ovoga agent koji menja formu ne bi znao da postoje.
     Vidi §4, „Tri seam-a". -->


# Verifikacija: šta se stvarno može proveriti

> Definicija gotovog je u CLAUDE.md §5: zeleno nad ispravnim **i dokazano crveno**
> nad namerno pokvarenim kodom, izlaz priložen.
>
> CI ne pokreće Excel. Postoje dva alata: jedan radi svuda i gleda izvor, drugi
> traži Windows + Excel i gleda **ponašanje**.
>
> Zašto su pravila baš ovakva — incidenti, promašene hipoteze, cena:
> `docs/engineering/postmortems/2026-08-verifikacija.md`. Ovde su samo pravila.

## 1) `tools/vba_check.py` — radi i u Claude Code sesiji (Linux/macOS)

```bash
python3 tools/vba_check.py                 # sve nad src-vba/   (Windows: python tools\vba_check.py)
python3 tools/vba_check.py src-vba/modX.bas
python3 tools/vba_check.py --hook          # tiho kad je čisto
```

Exit `0` = čisto, `2` = ima nalaza.

| Provera | Šta hvata |
|---|---|
| `ASCII` | ne-ASCII bajt u `.bas`/`.cls`/`.frm`/`.doccls` |
| `DEKLARACIJA` | modul-level `Const`/`Dim`/`Declare`/`Type`/`Enum` posle prve procedure |
| `REZERVISANO` | ime koje se case-insensitive poklapa sa VBA ključnom reči (`eNum`) |
| `DUPLIKAT` | isti `Public Sub/Function/Const` u dva modula → „Ambiguous name" |
| `PORUKA` | `Poruka("KLJUC")` bez para u `modPoruke.UpsertPoruke` |
| `NEDEFINISAN` | poziv procedure koja nigde nije definisana → „Sub or Function not defined" |
| `ARNOST` | poziv sa pogrešnim brojem argumenata → „Wrong number of arguments" |

**Obavezan korak pre commita** svake VBA izmene. Poslednje dve su namerno uske
(samo `.bas`, samo poziv u poziciji naredbe) — lažan nalaz u hook-u je gori od
propuštenog. Uzak izuzetak od `DUPLIKAT`-a postoji za ugovor ekrana (`Scr_*` u
`modScr*`): ljuska ih zove isključivo kvalifikovano i kasno vezano.

Ne kompajlira VBA: ne hvata tip-greške ni nedeklarisane promenljive, a u
`.frm`/`.cls` ne radi `NEDEFINISAN`/`ARNOST` (nasleđeni članovi se zovu bez
kvalifikatora). Za to treba Excel.

## 2) `tools/run_vba.py` — SAMO Windows + Excel + `pywin32`

Import `src-vba/` → `Debug > Compile` → `EnsureRuntimeSchema` → suite, headless.
Traži COM, pa se **ne pokreće na Linux/macOS** — ni u web sesiji ni u CI. Tamo se
testovi ponašanja ne izvršavaju uopšte i `Stop` hook prolazi tiho: sesija može da
se završi „zeleno" bez ijedne provere ponašanja. **To nije verifikacija i ne sme
se tako prijaviti** — VBA izmena koja dira ponašanje ide na Windows mašinu ili
ostaje neverifikovana.

```powershell
python tools\run_vba.py --compile-only     # najbrže i najstabilnije
python tools\run_vba.py                    # pun set: ceo podrazumevani katalog
python tools\run_vba.py --suite RunAllTests
```

- Radi nad **temp kopijom**; original se nikad ne dira. Bez `--workbook` ide
  `tests/fixtures/otkup_test.xlsm`. Ako fixture-a nema, pravi **praznu** `.xlsm` —
  dovoljno za compile, ne i za suite (`New frmOtkup` pukne na 1004 u
  `UserForm_Initialize`). Za suite napravi pravi fixture, §4.
- `python3 tools/run_vba.py --self-test` je jedini deo koji radi svuda — provera
  da strip VBA header-a ne propušta header u kod. **Pokreni ga posle svake izmene
  `_read_code_body`/import logike.**
- **Šema se podiže posle importa a pre suite-ova** (schema pravila dolaze iz
  svežeg koda, ne iz donora). Rutina je idempotentna. Pala priprema šeme obara run
  i kad su sve suite zelene — rezultati nad nepripremljenom sveskom nisu merodavni.
- Compile verdikt: `COMPILE NEJASNO` **ne obara** run kad suite-ovi idu (da bi se
  `RunAllTests` pokrenuo, VBA mora da kompajlira `modTest` i sve što on
  referencira). Eksplicitan `FAIL` pada uvek, kao i `NEJASNO` uz `--compile-only`.
- Ne zove `modVbaTools.ImportAllVBA` (hardkodiran folder + završni `MsgBox` = smrt
  za headless), nego ponavlja njegovu logiku preko COM-a. Dijaloge zatvara
  watchdog. Detalji: `docs/EXCEL_TEST_HARNESS.md`.

## 3) `gate` vs „blind" suite

Suite sa `gate: True` **podiže grešku** kad provera padne, pa je runner vidi kao
crvenu. Suite sa `gate: False` rezultat piše samo u Immediate — runner je
prijavljuje kao **`blind`**, što znači „prošla bez greške", a **NE** „sve provere
prošle".

**Katalog `SUITES` u `tools/run_vba.py` je jedini izvor istine** — koja suite
postoji, da li je `gate` i da li je u punom setu (`default: True`). Ne prepisivati
ga ovde ni bilo gde drugde. Van punog seta su blind ostaci
(`RunNovacSmokeSuite`, `RunProductionHealthCheck`, `TestMonitoring_All`) i ono što
traži mrežu ili live SEF nalog. **Ne proširivati na `--all`**: među `Run*`
procedurama nisu sve testovi (`RunSelfUpdate`, `RunGoogleAuthSetup`).

Verdikt brzog seta ne dolazi iz toga da li `Run()` pukne — `modTest` hvata grešku
po testu da jedan pad ne obori ostale — nego iz `last_run.txt` pored sveske. **Nema
fajla = pad.**

**Nova suite mora biti `gate`.** Blind suite je test koji niko neće videti kad
pukne. Konverzija, po uzoru na `modTestBanka.ERR_BIT_SUITE_FAILED`:

1. `Private Const ERR_X_SUITE_FAILED As Long = vbObjectError + <slobodan>` u
   deklaracionu sekciju (zauzeti offseti: 2900, 2950, 2960–2963, 3010–3012, 3100)
2. posle završnog izveštaja: `If mFail > 0 Then Err.Raise ERR_X_SUITE_FAILED, ...`
3. u `EH`: prebroj prekid kao pad (`Fail "SUITE prekinut..."`) pa podigni

Pa `gate: True` + `default: True` u katalogu. **Prođi i rane izlaze**: uslov koji
tiho radi `Exit Sub` pre prve provere mora da podigne grešku sa porukom koja
počinje `suite NIJE pokrenut:` — inače runner „nije se pokrenulo" vidi kao `OK`.
Poznata otvorena rupa: `RunBankaImportTestSuite` (rani `Exit Sub` kad
`tblBankaImport` / `tblOtkup` ne postoje).

Konverzija nije gotova bez dvosmernog dokaza — §4, „Sabotaža".

## 4) `modTest` — suite koja pada na PONAŠANJU, ne na sintaksi

`vba_check` hvata sintaksu; `modTest` hvata izmenu koja se uredno kompajlira, a
menja ponašanje. Legacy forma se **ne gasi** (`docs/UI_MIGRACIJA_KATALOG.md`), pa
obe kopije nose svoj test.

Tri nad `frmOtkup.ClearOtkupFields`:

| Test | Šta drži |
|---|---|
| `T_PosleSnimanja_ZadrzavaKontekstOtpremnice` | datum se posle snimanja NE briše (+ pun snapshot forme) |
| `T_PosleSnimanja_ZadrzavaZbirnu` | broj zbirne ostaje, i drugi blok dobija istu zbirnu |
| `T_ClearForm_BrisePartnera` | `cmbKooperant` se BRIŠE (obrnut smer od prva dva) |

Tri nad novim UI-jem (`modOtkupUI`):

| Test | Šta drži |
|---|---|
| `T_ParseDatum_Ugovor` | prazno/nečitljivo je `0`; `d.m.yyyy` se čita kao DMY bez `CDate`; trailing tačka se skida; nemoguć datum se **odbija**, ne preliva (`30.02` → `2.3`) |
| `T_ParcelaID_IzSkriveneKolone` | ID parcele dolazi iz **skrivene druge kolone**, ne iz prikaznog teksta; sakriveno polje ne šalje parcelu u dokument |
| `T_ClearForm_Ugovor` | ista tri ponašanja kao legacy trojka + razlika novog UI-ja: **bez** aktivne otpremnice datum se vraća na danas |

Pet nad **upisom zbirne i prijemnice** (`modDokUnos` + ruta u `modScrDokumenti`,
v6-ui-116). Ovi **ne grade formu**: pravilo unosa živi u modulu bez ijedne
kontrole, pa se tamo i proverava — brzo i bez stanja koje ostaje za sobom.

| Test | Šta drži |
|---|---|
| `T_ZbirnaValidiraj_TraziVozaca` | vozač je entitet niza zbirne (Z3a) i **prva** provera; sa vozačem zbirnu zaustavlja tek kupac |
| `T_ZbirnaValidiraj_MoraDaSeSlazeSaOtpremnicama` | zbirna je poklopac nad otpremnicama: kg **i** ambalaža moraju da se poklope (`ValidateZbirnaPreUnosa`); zbirna bez ambalaže ne prolazi; **kapija ne zavisi od `VALIDACIJA_UNOSA`** |
| `T_PrijemnicaValidiraj_TraziKupca` | kupac je **prva** provera (kod zbirne je vozač); broj zbirne je obavezan |
| `T_BrutoNeto_PoRezimu` | prijemnica: uneti bruto → `BrutoKg`, u `Kolicina` ide neto, po klasama zasebno. Zbirna: bruto→neto **nema i ne sme da ga dobije** (`tblZbirna` nema `BrutoKg`; ona zbraja već netirane otpremnice) |
| `T_ScrSave_RutaPoRezimu` | `Scr_Save` vodi F2–F7 u njihov modul (dokaz: svaki staje na pravilu koje je **isključivo njegovo**), a nepokriven režim (F8 storno) i dalje vraća `OTKUI_TODO_NEVEZANO` |

Tri nad **upisom novca i ambalaže** (`modNovacUnos`, v6-ui-117). Isti obrazac:
pravilo živi u modulu bez ijedne kontrole, pa se tamo i proverava.

| Test | Šta drži |
|---|---|
| `T_IsplataValidiraj_TipNovcaPoIzboru` | tip novca po izboru primaoca/bloka/prekidača — sve četiri grane (`KesOtkupacKoop`, `VirmanFirmaKoop`, `VirmanAvansKoop`, `KesFirmaOtkupac`); iznos ne preko ostatka bloka ni preko OM avansa; primalac-otkupno-mesto je entitet novca i odbacuje blok; broj dokumenta je kapija **oba smera** `VALIDACIJA_UNOSA` |
| `T_UplataValidiraj_FakturaOdlucujeTip` | izabrana faktura → `KupciUplata` + napomena sa brojem; bez nje → `KupciAvans`; uplata ne preko **trenutnog** preostalog iznosa; faktura drugog kupca i nepostojeća faktura se odbijaju |
| `T_ReversValidiraj_SmerJeObavezan` | smer je obavezan (prazan je ranije tiho knjižio „OM prima od vozača"); količina i tip ambalaže idu pre smera; kooperantski smer ne prima kupca; firma↔OM traži vozača **i bez** stroge validacije; prevod segmenta u `koopSmer` |

Tri nad **kapijama koje UI ne može da odbrani** (pregled PR #190). Ovo su
pravila oko kojih je ranije bilo moguće da suite bude zelen a novac ode na
pogrešno mesto:

| Test | Šta drži |
|---|---|
| `T_IsplataBlokGuard_VlasnistvoITrenutniOstatak` | `IsplataBlokProblem`: blok mora postojati, ne biti storniran, pripadati **tom** kooperantu i **tom** otkupnom mestu, a neisplaćeni ostatak se čita iz podataka **sada** — ne iz snimka koji je ekran poslao (helper zato namerno šalje lažnih 999999) |
| `T_NerazresenIzbor_NeProlaziKaoPrazno` | ukucano a nerazrešeno ime/blok/faktura (`ListIndex = -1` uz vidljiv tekst) **zaustavlja** dokument umesto da tiho postane isplata otkupnom mestu / avans kooperantu / avans kupca; obrnut smer: **prazno** polje i dalje prolazi |
| `T_WriterGuard_OdbijaTudjBlok` | `SaveOMUlaz_TX` odbija nemoguću kombinaciju **bez ijedne UI provere** — zove se direktno, kao što ga zove i legacy `frmDokumenta`; odbijen upis ne ostavlja red u `tblNovac` |

> **Zašto writer, kad modul već proverava.** Modul proverava nad snimkom iz
> trenutka kad je lista punjena. Između punjenja i potvrde stanje se može
> promeniti — drugi unos, uvoz izvoda, drugi pozivalac istog writer-a — pa bi
> prošla prevelika isplata. Isti obrazac postoji u `ApplyAvansToOtkup`
> (target-owner + target-active + preračunat preostali iznos) i sada ga dele
> `IsplataBlokProblem` i `UplataFakturaProblem`: jedna implementacija, dva
> pozivaoca — core je diže kao grešku, modul je vraća kao poruku uz polje.

**Šta NE pokrivaju:** iznad `ClearForm` — mrežu i storno; a od puta upisa samo
**provere i bruto→neto**, ne i sam `Save*_TX` (transakcioni upis pokrivaju
`RunStornoTestSuite` i `RunBusinessFlowProSuite`). Forma se gradi bez `.Show`, pa
`UserForm_Activate` (raspored, `GoFullScreen`, punjenje mreže) nikad ne ide.

### Tri seam-a koja kod nosi zbog ovih testova

- `ClearForm` / `ParseDatum` / `ParcelaID` su **`Public`**, ne `Private` — test ih
  zove direktno, bez vožnje celog upisa (stanica-lock, PDF, auto-lanac hladnjače).
- **Tri `SetFocus`-a** (dva u `ClearForm`, jedan na kraju `ApplyPrefill`) su iza
  `If Not IsTestMode()`. Forma koja nije `.Show`-ovana ne može da primi fokus, a u
  nevidljivom Excelu `SetFocus` **ne puca nego trajno visi**. U produkciji je
  `IsTestMode()` uvek `False`.
- `modScrDokumenti.Scr_OtpTestSet` — jedini način da test dobije aktivnu
  otpremnicu. Tvrdo gejtovan: van test-režima ne radi ništa.

Polja se postavljaju kroz `ApplyPrefill`, ne pisanjem u kontrolu: direktan upis u
`fgDatum` okine `OnDatumChanged`, a on traži stanica-lock i predlog broja **sa
pitanjem Google-u**.

### Pisanje testa

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
- **Golden za novi UI ne postoji i ne treba.** `DumpKontrole` nad `frmOtkupUI`
  uhvatio bi i `titDatum` (`FmtDatumPun(Now)`), pa bi golden padao svakog sledećeg
  dana. Legacy forma ima fiksne `.frx` kontrole i tu je snapshot smislen.

### Akceptaciona komanda

```powershell
python tools\make_fixture.py --donor "<put>\AgriX_2.28.4.xlsm"   # jednom
python tools\run_vba.py --suite RunAllTests                      # brzi set (modTest)
python tools\run_vba.py                                          # pun set
```

`--suite RunAllTests` vrti samo `modTest` — to je **brzi set** i to je ono što
pušta `Stop` hook. Pun set (goli poziv) pušta se **namerno**, pred commit ili
release: 11 suite-ova, **~1055 provera**, `EXIT=0` i bez `BLIND` reda — mereno na
operaterskoj mašini 14.08.2026. Brojevi po suite-u nisu ovde prepisani; stoje u
ispisu rana, a katalog je u `run_vba.py`.

Za trijažu masovnih padova: `run_vba.py --suite X --keep`, **snimi** temp kopiju,
pa `tools/read_test_log.py <temp>/otkup_test.xlsm` grupiše padove po temi i
razlogu.

### Sabotaža — dvosmerni dokaz

Radi se **jednom, kad se test piše ili menja**, ne pri svakom ranu:

```bash
python tools/sabotaza.py --lista
python tools/sabotaza.py clear-datum          # primeni jednu
python tools/run_vba.py --suite RunAllTests   # ocekuj FAIL po IMENU tog testa
python tools/sabotaza.py --vrati              # vrati
```

Dvadeset devet sabotaža: sedam nad `modOtkupUI`, sedam nad putem upisa F3/F4
(`modDokUnos`, `modScrDokumenti`), osam nad putem upisa F5/F6/F7
(`modNovacUnos`, ruta isplate) i sedam nad kapijama vlasništva i trenutnog
ostatka (`modNovac`, `modDokumenta`). Koja obara koji test i sa kojom tvrdnjom
— **`--lista`**, ne prepisivati ovde; skripta je izvor istine. Prvih četrnaest
potvrđeno 14.08.2026 (`TESTS=11 FAIL=0`), osam iz v6-ui-117 nad `TESTS=14`,
sedam iz v6-ui-118 nad `TESTS=17 FAIL=0`.

**Dve sabotaže namerno obaraju više od jednog testa** i to je tačan nalaz, ne
curenje stanja: `blok-ostatak-snapshot` obara tri (kapija, put unosa, ruta), a
`blok-tudj-om` dva (kapija i writer). Isto pravilo je namerno provereno na više
nivoa, pa njegovo uklanjanje mora da se vidi na svakom. Razlika u odnosu na
pravo curenje je i dalje merljiva: svaki pad ima **svoju** poruku i svoju
tvrdnju, a ne `Err.Number=0` sa praznim opisom.

> **Sidro sabotaže je deo koda koji sabotira.** `clear-zbirna` se razvezalo čim
> je `ClearForm` dobio `fgNovac` u spisak polja — skripta bi tiho prijavila
> „sidro nije jednoznačno" tek pri sledećem pokretanju, a do tada bi izgledalo
> da je dokaz i dalje važeći. Kad menjaš red koji je nečije sidro, promeni i
> sidro, pa ponovo pokaži crveno.

Za legacy formu sabotaža se radi ručno u `ClearOtkupFields` (dodaj
`txtDatum.value = ""`, `txtBrojZbirne.value = ""`, ukloni `cmbKooperant.value = ""`),
revert je `git checkout -- src-vba/frmOtkup.frm`. Svaka od njih obara i snapshot
iz prvog testa — namerno, snapshot hvata i polja koja niko nije tražio da proveri.

> **Kad sabotaža obori DVA testa, to ne mora biti curenje stanja.**
> `zbirna-vozac` i `prijemnica-kupac` obaraju i `T_ScrSave_RutaPoRezimu`, jer taj
> test dokazuje rutu time što prazan dokument staje na **prvom pravilu svog tipa**
> — ukloniš li to pravilo, menja se i ono što ruta vidi. Razlika u odnosu na pravo
> curenje (#186) je merljiva: ovde drugi pad ima **svoju poruku i svoju tvrdnju**,
> a ne `Err.Number=0` sa praznim opisom. Prvo proveri izolaciju, pa tek onda
> proizvod.

`parse-cdate` pada na tvrdnji „godina van poslovnog opsega" (`11.08.1899`) —
jedina tvrdnja koja razlikuje `CDate` od determinističkog parsera **na DMY
mašini**. Razliku na MDY mašini ne pokriva nijedan test i to se ne prijavljuje kao
pokriveno.

Tri zamke koje skripta rešava, i koje važe za svaki sličan zahvat nad izvorom:

1. **Kraj reda** — `src-vba` je CRLF na Windows-u, LF na Linuxu. Sidro sa zakucanim
   `\n` ne pogodi ništa, skripta tiho ne uradi ništa, run prođe nad neizmenjenim
   fajlom i izgleda kao da sabotaža „nije oborila" suite. Detektuj
   (`nl = '\r\n' if '\r\n' in s else '\n'`) i tvrdi `assert s.count(old) == 1`.
2. **Uvlačenje** — sidro se poredi od početka reda; inače isti niz pogađa dva mesta.
3. **Vraćanje** — `git checkout --` briše i nesnimljene izmene koje sa sabotažom
   nemaju veze (jednom je pojelo test seam-ove). `--vrati` radi obrnutu zamenu.

### Fixture i golden

`tests/fixtures/otkup_test.xlsm` je lokalan artefakt (`.gitignore`), pravi ga
`tools/make_fixture.py` iz **donor** sveske.

> **Kad se `FIXTURE` dict u `make_fixture.py` promeni, fixture se MORA
> regenerisati** — inače testovi padaju na podacima kojih nema. Donor može biti
> i **postojeći fixture**: on nosi punu šemu i nema VBA, a generator ionako
> briše sve redove pre sejanja. Izlaz mora biti druga putanja (donor = izlaz se
> odbija), pa se fajl posle premesti:
> `python tools\make_fixture.py --donor tests\fixtures\otkup_test.xlsm --out tests\fixtures\otkup_test_new.xlsm --force`

- Donor daje samo strukturu; spisak kolona se **ne** zakucava u Python (šema tabela
  je izvor istine — CLAUDE.md §4). Podaci su 100% sintetički, u transakciji koja se
  uvek poništava — nijedan klijentski podatak ne može da završi u golden fajlu na
  GitHub-u.
- Generator **uklanja sav VBA kod iz donora**: modul zaostao iz starijeg donora se
  izvršava i, ako nosi `Public` ime koje postoji i u svežem kodu, daje „Ambiguous
  name" → `Cannot run the macro`, poruka koja ne liči na compile grešku. Za sveske
  kroz `--workbook` ne briše ništa, nego prijavljuje `ORPHAN` red.
- Šemu donora ispisuje `tools/dump_schema.py` (samo čitanje).

`tests/golden/*.txt` idu u git. Kad golden ne postoji, test ga upiše i **padne** —
nov golden mora proći ljudski pregled pre nego što postane merilo. Dva pravila:
**ASCII** (`DumpKontrole` escape-uje dijakritiku u `\uXXXX`; VBA `Print #` piše u
ANSI stranu koja `ć` nema) i **LF** (`.gitattributes` drži `eol=lf`, inače suite
pada na svakom svežem klonu na Windows-u).

### Hook-ovi

**PostToolUse** (`.claude/hooks/vba-check.sh`) posle svakog `Edit`/`Write`:

- nad `.bas`/`.cls`/`.frm`/`.doccls` pušta `vba_check` nad tim fajlom;
- nad `.claude/settings*.json` proverava da je i dalje validan JSON — fajl koji ne
  prođe validaciju Claude Code odbacuje **u celini**, pa padnu sva permission
  pravila i oba hook-a odjednom.

**Stop** (`.claude/hooks/vba-test.sh`) na kraju sesije:

1. JSON provera istih fajlova — **na svakom Stop-u**, jer merge/rebase ne prolazi
   kroz `Edit`, a spoj dve grane koje dopisuju na kraj istog niza prolazi bez
   konflikt markera;
2. `who_writes.py --check` (instant, bez Excela) kad je `src-vba/` diran;
3. **brzi set** `--suite RunAllTests` + **žig** u `.git/vba-test-stamp` (HEAD +
   hash nekomitovanog diffa nad `src-vba/`). Isto stanje se ne proverava dvaput;
   žig se piše samo na zeleno, pa se pao set ponavlja dok se ne popravi.

Bez `pywin32`/Excela korak 3 prolazi **tiho** — u Linux sesiji ostaju samo koraci 1
i 2. **Pun set nije u hook-u** nego je namerna komanda pred commit/release; nova
suite u katalogu ulazi u pun set, a ne u hook.

Obe skripte biraju interpreter **probom** (`"$PY" -c ""`), ne preko `command -v`:
na Windows-u `python3` postoji u PATH-u kao Microsoft Store alias koji ne pokreće
ništa. Kad menjaš hook, prvi test je da li uopšte pukne nad namerno pokvarenim
fajlom — tih hook izgleda isto kao čist repo.

## 5) Šta i dalje ostaje na operateru

Finalni smoke-test u Excelu (klik po klik), izgled forme, štampa i PDF, ponašanje
nad pravim podacima. Zato svaki rad završi kratkom, numerisanom test-checklistom u
chatu — vidi `.claude/rules/git-i-release.md`. Checklista **nije** zamena za test
koji je moguće napisati.
