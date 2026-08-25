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
## 10) Sam dvosmerni dokaz je istrunuo — deset mrtvih sidara (25.08.2026)

Pravilo iz §4 traži da se posle izmene pusti **ceo** dvosmerni dokaz i tvrdi da je
broj crvenih jednak broju sabotaža. Nad 222 sabotaže to traje oko dva i po sata,
pa se u praksi puštao podskup — obično onaj koji dira tekući posao.

Kad je prvi put pušten preko svega što tekuća izmena dira (PR #226), ispalo je da
**deset sabotaža ne može ni da se primeni**: kod ispod njihovog sidra je odavno
popravljen, pa se sidro više ne nalazi. Za tih deset tvrdnji dokaza više nije ni
bilo.

Zašto je prošlo neprimećeno: sabotaža sa zastarelim sidrom **ne javlja** „test je
prošao" — javlja „nisam našla sidro", i to na `stderr`, usred izlaza koji traje
pola sata. U petlji to izgleda isto kao zeleno.

| Nalaz | Koliko |
|---|---|
| sidro zastarelo ili dvosmisleno | **10 od 222** |
| sabotaža koja ne obara ništa | 1 (`ekran-curi-greska`) |
| očekivano ime testa ne postoji | 1 (test preimenovan) |
| dve sabotaže dele tvrdnju (zamka 5) | 1 par |

### Jedna „mrtva" sabotaža to nije bila — izveštaj je pucao

`parcela-tekst` je dva puta prijavljena kao „ne obara ništa". Zapravo je uredno
obarala svoju tvrdnju, ali je `run_vba.py` **pucao pri ispisu**: poruka o padu
nosi ono što je test video, a to je bio `ChrW(183)` iz prikaznog teksta parcele —
znak koji cp1252 konzola ne ume da ispiše. Run bi završio `Traceback`-om **umesto**
linijom `FAIL <test> -- <tvrdnja>`, pa je petlja videla nula padova.

**Pravilo:** izveštaj o rezultatu ne sme da pukne zbog jednog znaka. Ispis je sada
otporan (`errors="replace"`), jer je alternativa da crven test izgleda kao mrtva
provera.

### Šta je promenjeno

**1. Jeftina polovina pravila je sada statička.** `python tools/sabotaza.py
--proveri-sidra` proverava, bez Excela i za sekundu: da se svako sidro nalazi
**tačno jednom** (istim poređenjem od početka reda koje koristi i sam alat), da
očekivani test postoji, i zamke 4, 5, 7 i 8 iz kataloga.

**2. Vezano za `vba_check`.** Provera ide kroz `PostToolUse` hook posle svake VBA
izmene — a baš VBA izmena je ono što obara sidro. Ko popravi kod, odmah vidi koju
je sabotažu time obesmislio.

**3. Pun dokaz je dobio alat:** `python tools/dokaz.py [filter]`. Do sada je bio
skripta iz scratchpada, pa se i nije puštao ceo.

**4. Kriterijum je izoštren.** Sabotaža **sme** da obori više testova — široka
izmena to i radi. Ne sme da ne obori **svoj**. Tekst tvrdnje u katalogu je
dokumentacija (često parafraza) i ne obara dokaz; ime testa je obavezno tačno.

**5. Priznati nalaz ima ime — i to ime baš tog pada.** `POZNATI_NALAZI` u
`sabotaza.py` drži nalaze koji
imaju vlasnika a ne mogu se zatvoriti bez izmene testa. Ispisuju se kao
upozorenje i ne obaraju gejt — crvena provera koju svi nauče da preskoče ne čuva
ništa. Upis koji više ništa ne pokriva je **isto nalaz**, pa spisak ne može tiho
da raste.

Vrednost upisa je **početak baš te poruke**, ne njena vrsta. Prva verzija je za
pun dokaz upisivala golo `PALA DRUGA TVRDNJA` — a to je cela *kategorija*: svaka
buduća, sasvim druga pogrešna tvrdnja u istom testu bila bi tiho progutana kao
poznata i dokaz bi završio zeleno. Prefiks zato nosi i ime tvrdnje koja stvarno
pada, prepisano iz izmerenog izlaza. Provereno u oba smera: tačan prefiks →
`POZNATO`; bilo koja druga poruka → `PROBLEM` (i, uz to, upis se prijavi kao
mrtav).

### Četiri rupe u samom alatu, nađene u review-u

Prva verzija ove mašinerije imala je četiri problema — i sva četiri su bila u
delu koji **treba da dokazuje** da su testovi živi:

**1. Provera nije išla kroz hook, iako je to bila cela poenta.** Bila je vezana
za `not args.paths`, a `PostToolUse` hook zove baš `vba_check.py --hook <fajl>` —
dakle sa putanjom. Katalog se kroz hook nikad nije proveravao; video bi ga tek
CI, posle dvadeset izmena. Katalog nema veze sa tim koji je fajl dat: sidra
pokrivaju ceo `src-vba`, pa se sada proverava uvek (0,14 s, tiho kad je čisto).

**2. Dokaz je mogao da bude lažno pozitivan nad crvenom bazom.** Alat je ispisivao
baseline ali ga nije **tvrdio**. Test koji već pada iz trećeg razloga proglasio bi
svaku sabotažu nad sobom dokazanom — uključujući onu koja ne radi ništa. Sada je
zelena baza **kapija**: nije zelena → `rc=2`, bez ijedne mutacije.

> Razlika je suštinska: „posle mutacije postoji crveno" nije isto što i „mutacija
> je izazvala crveno".

**3. Čišćenje nije bilo fail-safe.** Mutacija se vraćala tek posle uspešnog run-a;
timeout Excela ili `Ctrl+C` usred prolaza od dva i po sata ostavljao je **namerno
pokvaren** radni izvor. Sada ide kroz `finally`, uz poređenje potpisa celog
`src-vba` posle svakog vraćanja — ako se ne poklopi, dokaz staje odmah, jer bi sve
mereno posle toga išlo nad pokvarenim kodom.

**4. Banka-suite se nije mogla ni izmeriti.** Njeni detalji idu u Immediate
prozor, a opis podignute greške ne preživi COM granicu — `pywin32` vidi golo
`Exception occurred`. Alat je zato video samo „3 provere palo", pa je za te
sabotaže tvrdnja bila „nešto je palo".

Mereno, ispalo je gore nego što je izgledalo: `run_vba` je detalje čitao **samo**
kad `Run()` ne pukne, a ta suite pad prijavljuje **greškom** — pa se rezultat
nikad nije ni čitao. Svaka sabotaža nad njom čitala bi se kao „ne obara ništa",
što je lažni negativ.

Sada i ta suite piše `last_run_banka.txt` (isti format kao `modTest`), rezultat se
čita **i kad suite pukne greškom**, a identitet se vadi iz stabilnog prefiksa
tvrdnje (`T21 izabran placen blok: ...`). Provereno da nije prazno: kad se unos
namerno usmeri na drugi test, alat kaže `NE OBARA SVOJ TEST, nego: T21`.

### Još dva, iz drugog kruga review-a

**5. Dve suite sa rezultat-fajlom prepisivale su jedna drugu.** Rezultat je išao u
jedan zajednički `report["tests"]`, pa je banka (koja ide kasnije) prepisivala
`RunAllTests`. Full run je umeo da završi ovako:

```
SUITE   FAIL   RunAllTests
SUITE   OK     RunBankaImportTestSuite
TESTS   196 ukupno, 0 palo
```

— izlaz koji **sam sebi protivreči**, a ime palog testa je nestalo. Izlazni kod je
i dalje bio crven, ali dijagnostika je lagala. Sada svaka suite ima svoj slot i
svoj označeni red (`TESTS   RunAllTests: 115 ukupno, 1 palo`).

**6. Dokaz je prihvatao pogrešnu tvrdnju u pravom testu.** Alat je proveravao da
je pao **njen test**, a razliku u tekstu tvrdnje prijavljivao kao „parafraza — u
redu". To vraća zamku 6: `AssertEq` puca na **prvom** padu, pa sabotaža koja usput
obori raniju, uzgrednu tvrdnju ostavlja ciljanu **neizvršenom** — a izlaz i dalje
nosi ime pravog testa.

> Pravi test + pogrešna tvrdnja = **crven** dokaz, ne zelen.

Peti član n-torke je time prestao da bude komentar i postao **merena vrednost**:
mora se naći u poruci koja je pala, i to **samo među porukama njenog testa**.

Čim je pravilo postalo strogo, izmerilo je **sedam** neusklađenih od trinaest.
Pet je bila razlika u rečima. Dva su bila prava nalaza — sabotaža obara
**preduslov**, pa ciljana tvrdnja ne dođe na red: `relink-ignorise-generaciju` i
`f8-identitet-po-broju`. Ni jedan se ne može zatvoriti bez izmene testa ili
fikstуre (uža sabotaža bi u prvom slučaju obarala tuđu tvrdnju, a u drugom —
mereno — ne obara ništa). Zapisani su u `POZNATI_NALAZI_DOKAZ`, sa razlogom.

Cena je poštena: tekstovi za sabotaže koje nisu skoro puštane nisu usklađeni sa
onim što stvarno pada, pa će ih alat prijaviti čim se puste. To nije regresija
nego prvi put da se ta razlika uopšte meri.

### Šta ovo ne rešava

Da li sabotaža stvarno nešto obara zna **samo** pun dokaz. Statička provera hvata
mrtvo sidro, ne mrtvu tvrdnju: `ekran-curi-greska` je imala ispravno sidro i
uredno se primenjivala — a suite je ostajala zelena.
