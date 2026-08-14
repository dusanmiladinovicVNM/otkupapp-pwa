# Postmortem: verifikacija i test harness, avgust 2026.

Istorija incidenata izvučena iz `.claude/rules/testovi.md` (14.08.2026). Pravila
koja iz njih slede žive u tom fajlu; ovde je **zašto** — arheologija koja ne mora
da se učitava u svakoj sesiji.

Čita se kad neko pita „zašto je pravilo baš takvo" ili kad se isti obrazac ponovi.

---

## 1) Nepripremljena šema izgledala je kao regresija u proizvodu

`RunBusinessFlowProSuite` je davao `Total=310 | Passed=163 | Failed=147`. Ličilo
je na masovni nalaz o proizvodu; bila je nepripremljena sveska.

Fixture nastaje iz **starijeg donora** (2.28.4), kod je noviji. Kolone dodate u
međuvremenu ne postoje dok se ne pokrene schema upgrade. Posle
`EnsureRuntimeSchema` suite prolazi 100%.

**Pravilo koje je ostalo:** `run_vba.py` uvek pusti `EnsureRuntimeSchema` posle
importa a pre suite-ova, i pala priprema šeme obara run i kad su sve suite zelene.
Redosled nije proizvoljan — schema pravila dolaze iz svežeg koda, ne iz donora.

> Najskuplji promašaj u ovom poslu: 147 padova čitano kao nalaz o proizvodu.

## 2) „Cannot run the macro" — zaostali duplikat u svesci, ne compile greška

`TestLicense_All` nije mogao da se pokrene. Ručno pokretanje je prolazilo, driver
je padao — tri tačna signala koja su zajedno izgledala kontradiktorno, jer se
pitanje **„nad kojom sveskom"** nije postavljalo nigde.

Fixture je nasleđivao **131** VBA modul iz donora; jedan je nosio `Public` ime koje
postoji i u svežem kodu → „Ambiguous name" → VBA odbija da pokrene makro. Poruka
ne liči na compile grešku. `vba_check` je s pravom bio zelen: duplikata u repou
nema.

**Pravilo:** `make_fixture.py` uklanja sav VBA kod iz donora; za tuđe sveske
(`--workbook`) prijavljuje `ORPHAN` red umesto da briše.

> Hipoteza da je u pitanju compile greška u `modLicenseTests` bila je **netačna**.
> Stajala je označena kao nepotvrđena i oborena je pokretanjem.

## 3) Jedna sabotaža obarala je dva testa — kriva je bila izolacija, ne proizvod

`parcela-tekst` i `parcela-vidljivost` obarale su svoj test **i još**
`T_ClearForm_Ugovor`, sa `Err.Number=0` i praznim opisom.

`T_ParcelaID_IzSkriveneKolone` čistio je za sobom tek poslednjom linijom
(`ReleaseOtkupUIForm`) — pa test koji **padne** nikad do nje ne stigne.
`mFrm`/`Btns`/keš u `modOtkupUI` i aktivna otpremnica u `modScrDokumenti` ostaju, i
sledeći test gradi ekran nad ostacima prethodnog. Čišćenje je stajalo na **zelenoj
putanji**.

**Pravilo:** `CleanupPosleTesta` se zove iz `EH` grane, a `Err` se čita **pre**
njega (`OtkupUI_Release` je ceo pod `On Error Resume Next`, što briše `Err`). Pad
bez opisa prijavljuje `Err.Number` — „`FAIL T_X`" bez razloga koštalo je dva rana
dijagnostike.

> Ista pouka kao kod šeme: **drugi pad u ispisu ne mora biti drugi nalaz.** Kad
> jedna sabotaža obori dva testa, prvo proveri izolaciju, pa tek onda proizvod.

## 4) Četiri puta zeleno-ali-nedokazano-crveno (PR #181)

Suite koja je zelena nad ispravnim kodom, a nije pokazana crvena nad pokvarenim,
ne dokazuje da išta meri. U PR #181 je to bio ishod četiri puta zaredom.

**Pravilo (CLAUDE.md §5):** nova ili izmenjena provera nosi dokaz u oba smera —
namerno pokvari, pokaži pad **po imenu**, vrati, pokaži zeleno.

## 5) Četiri putanje lažnog zelenog: tihi `Exit Sub`

Runner je „suite se nije pokrenula" video kao `OK`:

| Suite | Uslov koji je tiho izlazio |
|---|---|
| `RunPaleteTestSuite` | paletiranje isključeno u Podešavanjima |
| `RunPaleteTestSuite` | zatečen `TST-` ostatak od prekinutog run-a |
| `RunPaleteTestSuite` | operater odustao na potvrdi |
| `RunAgrohemijaSmokeSuite` | dev-guard odbijen |

**Pravilo:** rani izlaz podiže grešku sa porukom koja počinje `suite NIJE
pokrenut:`. Pala provera je glasna; suite koji se nije ni pokrenuo je tih.

Do te verzije **nijedna** od tih suita nije se pokretala kroz `run_vba.py` uopšte —
compile probe je vraćao `NEJASNO`, `rc = 2` je padao pre suite petlje, i petlja se
nikad nije dosegla. Suite su postojale samo kao ručni `Alt+F8`.

## 6) `T13` u banci: pao je test vektor, ne produkcija

`RunBankaImportTestSuite` `PASS=186 FAIL=1`, rešeno u #183. `600.005` se u
`Double`-u čuva ispod pola pare, pa ga zaokruživanje korektno spušta na `600.00`.

**Pouka:** pre nego što se pad proglasi nalazom o proizvodu, proveri vektor.

## 7) Stop hook je puštao pun gate — sesija neupotrebljiva (do 14.08.2026)

Hook je puštao ceo podrazumevani set: 11 suite-ova, ~1050 provera, uz podizanje
Excela na **svakom** Stop-u. Uz to je grana „poslednji commit" bila zamka: čim
jednom commit-uješ izmenu u `src-vba/`, `git diff HEAD~1 HEAD` ostaje neprazan do
kraja sesije — pa se pun set vrteo i na turnovima gde se samo razgovaralo.

**Pravilo:** hook pušta brzi set + žig u `.git/vba-test-stamp` (HEAD + hash
nekomitovanog diffa). Pun set je namerna komanda pred commit/release. Cena provere
mora biti proporcionalna riziku izmene.

## 8) Jedan zarez je oborio ceo `.claude/settings.json` (14.08.2026)

Merge `5b3777c` spojio je dve grane koje su obe dopisivale na kraj `allow` niza —
**bez ijednog konflikt markera**. Rezultat je ostao bez zareza:
`Expecting ',' delimiter: line 30 column 7`.

Claude Code fajl koji ne prođe validaciju odbacuje **u celini**: pola dana nije
važilo nijedno od 29 permission pravila ni jedan od dva hook-a. Otud odobrenje za
svaku komandu i PostToolUse koji se nije palio.

**Pravilo:** JSON gate na dva ulaza — PostToolUse (kad taj fajl menja Edit) i Stop
(svaki Stop, jer merge/rebase ne prolazi kroz Edit). Detalji u `testovi.md`.

## 9) `python3` na Windows-u je Microsoft Store alias — PostToolUse hook je bio mrtav (14.08.2026)

`command -v python3` uspeva (`.../WindowsApps/python3`), a poziv ispiše „Python was
not found" i vrati `rc=49`. Oba hook-a su birala interpreter preko `command -v`, pa
su birala **mrtav** interpreter:

- `vba-check.sh` nije uspevao ni da isparsira payload → `file_path` prazan → tih
  `exit 0` nad **svakim** fajlom. Dokazano: `.bas` sa ne-ASCII bajtom kroz hook
  daje `EXIT=0`, a `python tools/vba_check.py --hook` nad istim fajlom `EXIT=2`
  sa nalazom `ASCII`.
- `vba-test.sh` bi na istom mestu prijavljivao `WHO_WRITES.md je zastareo` — lažan
  blokirajući nalaz — a zatim tiho preskakao suite (`import win32com` pada isto).

**Pravilo (rešeno u `211c0048`, PR #188):** interpreter se bira **probom**
(`"$PY" -c ""`), ne preko `command -v`.

> Isti obrazac kao permission pravila pisana u Bash/Linux obliku za PowerShell
> mašinu: **infrastruktura pisana bez smoke testa u stvarnom okruženju.** Kad se
> piše hook, prvi test je da li uopšte pukne nad namerno pokvarenim fajlom.
