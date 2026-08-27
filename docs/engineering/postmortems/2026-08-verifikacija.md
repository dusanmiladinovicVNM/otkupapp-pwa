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
fiksture (uža sabotaža bi u prvom slučaju obarala tuđu tvrdnju, a u drugom —
mereno — ne obara ništa). Zapisani su u `POZNATI_NALAZI_DOKAZ`, sa razlogom.

Cena je poštena: tekstovi za sabotaže koje nisu skoro puštane nisu usklađeni sa
onim što stvarno pada, pa će ih alat prijaviti čim se puste. To nije regresija
nego prvi put da se ta razlika uopšte meri.

### Šta ovo ne rešava

Da li sabotaža stvarno nešto obara zna **samo** pun dokaz. Statička provera hvata
mrtvo sidro, ne mrtvu tvrdnju: `ekran-curi-greska` je imala ispravno sidro i
uredno se primenjivala — a suite je ostajala zelena.
## 11) „Nedostaje kolona" nad sveskom u kojoj kolona postoji (25.08.2026)

U logu radne sveske:

```
ERROR | modStornoFlow.ZbirnaBrojJeDvos | Nedostaje kolona 'VozacID' u tabeli 'tblZbirna'.
```

`tools/dump_schema.py` nad **tom istom** sveskom pokazuje da `VozacID` u
`tblZbirna` **postoji**. Poruka je, dakle, opisivala stanje koje nije tačno.

Nije reprodukovano. Ovaj zapis postoji zato što su na putu do toga **dve moje
dijagnoze redom oborene merenjem**, i to je korisniji nalaz od same poruke.

### Prva dijagnoza: schema drift — oborena čitanjem sveske

Prvo objašnjenje je bilo „ta sveska nema kolonu". Otpalo za minut, čim je
puštena šema. Pouka je banalna i skupa: **kad poruka govori o šemi, pročitaj
šemu** pre nego što se krene u kod.

### Druga dijagnoza: progutana greška ostaje živa — oborena sabotažom

`GetColumnIndex` je tražio kolonu ovako:

```vb
On Error Resume Next
GetColumnIndex = lo.ListColumns(colName).index
On Error GoTo 0
```

`ListColumns(ime)` za nepostojeću kolonu **diže grešku 9**, ne vraća nulu — pa
sam zaključio da `Err` posle toga ostaje živ i curi pozivaocu. Napisao sam
zamenu (prolazak kroz zaglavlje), test i sabotažu koja vraća zatečeni oblik.

**Sabotaža nije oborila ništa.** Razlog je u jeziku: u VBA **svaki** `On Error`
iskaz — uključujući `On Error GoTo 0` — resetuje `Err`. Curenja nije bilo.

Izmena je zato **povučena**: ostala bi kao popravka bez reprodukcije i bez
merljive razlike, tačno ono što `CLAUDE.md` §2 zabranjuje. Dve tvrdnje o
zatečenom ponašanju su ostale u testu, izričito označene kao **bez sabotaže**,
jer ih zatečeni kod već zadovoljava.

### Šta je stvarno urađeno

Poruka sada nosi i **zaglavlje koje je videla**:

```
Nedostaje kolona 'VozacID' u tabeli 'tblZbirna'.
Vidjeno zaglavlje: ZbirnaID, Datum, VozacID, BrojZbirne, ... (+21).
```

Time isti tekst prestaje da opisuje tri različita stanja — „kolone nema",
„tabele nema" i „zaglavlje je drugačije" sada se razlikuju **iz same poruke**,
bez ponovnog pokretanja.

### Treći put ista bolest — u samoj dijagnostici

Prva verzija pomoćne funkcije držala je **sve** pod jednim `On Error Resume Next`.
Pad čitanja tabele bi zato prijavila kao „tabela nije nadjena", a pad čitanja
zaglavlja kao „prazno" — dakle **opet** bi grešku predstavila kao stanje šeme,
samo jedan nivo niže, i to u funkciji čiji je ceo smisao da ta dva razlikuje.

Nađeno u review-u. Sada se posle svakog rizičnog koraka `Err` **čita** i, ako je
postavljen, poruka kaže da čitanje nije uspelo, uz broj greške.

### Poruka odgovara na pitanje koje se iz nje traži

Spisak imena je ograničen na 12 (poruka ide u log i u dijalog), pa bi kolona iza
te granice ostala nevidljiva baš u poruci koja treba da kaže da li postoji. Zato
se **tražena** kolona traži kroz celo zaglavlje i poruka to kaže izričito:

```
Vidjeno zaglavlje: ZbirnaID, Datum, VozacID, ... (+21).
Trazena kolona VIDJENA, pozicija 3.
```

Ako traženje kaže nula, a svež prolaz je vidi — uzrok nije šema nego put do nje.

### Simptom je sada merljiv

Test 117 ga izaziva kroz keš (`KesKoloneTestSet`, tvrdo gejtovan): nula se podmetne
za kolonu koja **postoji**, i tvrdi se da preživi ceo prozor i da poruka to ume da
razlikuje od stvarnog nedostatka. Time se **ne** tvrdi da je keš uzrok prvog
neuspeha — samo da jednom zapamćena nula ostaje do kraja prozora.

### Šta ostaje otvoreno, i zašto se ne popravlja naslepo

Nula iz `GetColumnIndex` se **kešira** za ceo `BeginTableCache` prozor
([modDataAccess.bas:148](../../../src-vba/modDataAccess.bas)), a
`InvalidateTableCache` čisti `mTableCache` i `mExclCache` — **`mColCache` ne**.
Jedan trenutan neuspeh bi tako postao trajan za ceo prozor, a kapije koje su
fail-closed (`ZbirnaBrojJeDvosmislenIkad` na grešci vraća `True`) na to staju:
storno nad zbirnom se zaustavlja uz „broj je dvosmislen".

To je **pojačivač**, ne uzrok — i dalje ne znam zašto je prvo traženje palo. Kad
se poruka sledeći put pojavi, nosiće zaglavlje i time reći da li je uzrok u
šemi, u tabeli ili u samom čitanju.

---

## 12) Isto truljenje, drugo polje — 119 zastarelih tvrdnji (26.08.2026)

§10 je zatvorio **sidra**: unos u katalogu čije je sidro zastarelo hvata se sada
za sekundu (`--proveri-sidra`), umesto posle dva i po sata punog prolaza. Isti
unos ima još jedno polje koje zastareva na isti način — **tvrdnju** — i ono nije
imalo nikakvu proveru.

### Šta je bilo

`dokaz.py` traži da se deklarisan tekst nađe u poruci koja je pala; ako se ne
nađe, javlja `PALA DRUGA TVRDNJA`. Tekst tvrdnje se u testu menja pri svakoj
doradi, a katalog o tome ne zna ništa — pa zastari **tiho**.

Alat time laže u **oba** smera:

- javlja grešku nad sabotažom koja radi savršeno;
- a sabotažu koja stvarno obara **tuđu** tvrdnju niko više ne čita, jer je alat
  naučio da laje.

Druga polovina je gora. `--proveri-sidra` postoji baš zato što provera koja
nikad nije pokazana crvenom ne dokazuje ništa; ovde je provera bila crvena
**stalno**, što je isti ishod drugim putem.

### Kako je izmereno, a ne procenjeno

| Korak | Nalaz |
|---|---|
| statički: tekst se ne nalazi u izvoru | 133 od 251 |
| ...neosetljivo na veličinu slova | **120** |
| ...vezano za telo **svog** testa | 123 (120 nigde + 3 u tuđem testu) |
| uzorak od 8 pušten kroz `dokaz.py` | **6** javilo `PALA DRUGA TVRDNJA` |
| puna žetva svih 123 | **119** zastarelih · 3 lažna pozitiva · 1 mrtva |

Prvi broj (133) je bio **pogrešan i nije prijavljen kao nalaz** — razlika je bila
samo u veličini slova, a `dokaz.py` poredi neosetljivo. Isto važi za tri unosa
koja se statički ne nalaze a u radu se poklapaju: tvrdnja im je sklopljena u radu.

### Popravka: mereno, ne pogađano

Katalog ne sme da dobije tekst „koji najviše liči". Zato:

- **žetva je izvor istine za to KOJA tvrdnja pada** — 123 prolaza `dokaz.py`-ja,
  svaki ispisuje poruku koja je stvarno pala;
- kad poruka **nije odsečena** (`dokaz.py` seče na 120 znakova), prefiks je pun
  tekst tvrdnje i upisuje se **doslovno**;
- tek kad jeste odsečena, pun tekst se traži u izvoru testa.

Doslovan upis je bitan zbog tvrdnji sklopljenih u radu:
`"Storno / " & tip & " cita svoju tabelu"` daje **različite** poruke za `OTKUP` i
`FAKTURA`. Skraćivanje na zajednički literal bi dve različite tvrdnje spojilo u
jednu — i to sam u prvom pokušaju i uradio (v. niže).

### Dve moje greške usput, obe poučne

**1. Naivan regex nad literalima.** `"([^"]{8,})"` u telu testa spaja zatvoreni
navodnik jednog stringa sa otvorenim sledećeg, pa nad
`AssertEq CLng(pre("brojPak")), 1&, "3 l trazi..."` vraća `)), 1&, ` kao da je
tekst. Rezultat: 50 „nema kandidata" koji su zapravo postojali. Lek je
tokenizator, ne bolji regex.

**2. Skraćivanje je spojilo dve tvrdnje u jednu.** Prvi prolaz je za tri para
upisao zajednički literal (`"Storno / "`), pa je `--proveri-sidra` odmah javio
**zamka 5** — „dve sabotaže koje test ne razlikuje". Provera je uhvatila grešku
koju je napravila popravka, u istom potezu. Da je nije bilo, katalog bi ostao
zeleno pokvaren.

### Pravilo: rupa se meri rečima, ne procentom

Provera prihvata dva oblika: tekst je doslovno u telu, ili su literali jednog
izraza u njemu **redom**, a između njih stoji nešto što liči na vrednost —
najviše tri reči po rupi.

Procenat pokrivenosti je probao pa odbačen: `" u dijalogu ide BEZ oznake"` je 51%
svoje tvrdnje, pa bi pao kroz prag od 60% — a isti prag bi primio slučajno
poklapanje kratkog literala u dugačkoj tuđoj tvrdnji. Broj reči modeluje ono što
se stvarno dešava: na tom mestu je bila **jedna vrednost**.

### Provera je prvo obećavala više nego što meri (iz review-a)

Prva verzija je za „doslovni" slučaj tražila tekst u **celom telu procedure**.
Telo sadrži i kod i komentare, pa bi kroz nju prošlo i:

```
tvrdnja = "AssertEq nosiDok, True"          ' komad KODA
tvrdnja = "recenica iz komentara ..."       ' komentar
```

`dokaz.py` poredi sa **porukom koja je pala**, a poruka može biti samo string
literal — pa bi ishod bio zeleno statički, `PALA DRUGA TVRDNJA` u prolazu. Tačno
rupa koju je posao trebalo da zatvori.

Gore od same rupe: `_literali` je **već vraćao** literale, a `_tela_testova` ih
je bacao i čuvao celo telo. Podatak je postojao; put do njega nije.

### ...a prvi dokaz te ispravke ništa nije dokazivao

Ovo je deo koji vredi najviše. Kad su dodata dva slučaja („tekst postoji samo kao
kod", „tekst postoji samo u komentaru"), oni su **prošli i sa vraćenim starim
ponašanjem** — dakle nisu merili ništa.

Razlog: self-test je svoje „telo" sklapao **sam**, pa je zaobilazio baš
`_tela_testova`, funkciju u kojoj je rupa i bila. Dokazivao je da
`_tvrdnja_pripada` pretražuje ono što **dobije**, a pitanje je bilo **šta
dobija**.

Lek je struktura, ne još jedan slučaj: izdvojen je `_telo_podaci(telo)`, koji
sada koriste **i** pravi put **i** self-test. Sabotaža te jedne funkcije odmah
obara oba slučaja, po imenu.

To je ista pouka koju `vba_check` već nosi u komentaru — self-test mora da ide
**kroz** produkcionu putanju, a ne pored nje. Ovde se pokazalo da važi i za
podatak koji putanja proizvodi, ne samo za funkciju koja ga troši.

### ...pa je i suženje bilo preširoko: literal nije isto što i poruka

Drugi krug review-a je pokazao da „tekst je u nekom string-literalu tog testa"
još uvek nije isto što i „tekst je poruka koju `dokaz.py` može da vidi".
Literal može da bude i **očekivana vrednost** ili obična dodela:

```vb
status = "blok drugog kooperanta se odbija"
AssertEq rezultat, "Placeno", "status fakture je ispravan"
```

Katalog sa `"Placeno"` bi statički prošao, a u prolazu dao `PALA DRUGA TVRDNJA`.

Provera zato ne gleda literale procedure nego **poslednji argument** assertion
poziva. Skup je uzak i pravilan — u sve četiri primitive oba harness-a poruka je
poslednji argument:

| Primitiva | Gde | Poziva |
|---|---|---|
| `AssertEq actual, expected, poruka` | `modTest` | 1176 |
| `ChkEq act, exp, nm` | `modTestBanka` | 88 |
| `Chk cond, nm` | `modTestBanka` | 82 |
| `ChkEqD act, exp, nm` | `modTestBanka` | 26 |

**Nezavisna potvrda popravke:** sužavanje sa „bilo koji literal" na „poruka
tvrdnje" nije oborilo **nijedan** od 251 unosa. Da je neki od 119 prepisanih
tekstova bio slučajno poklapanje sa nekim drugim literalom, ovde bi ispao.

Test koji nema nijednu prepoznatu poruku ne prolazi tiho — njegova tvrdnja pada
kao zastarela, pa nova assert primitiva ne može da otvori rupu neprimećeno.

### ...i to je bio treći krug: parser je imao dve rupe

„Poslednji argument assertion poziva" nije dovoljno ako se taj argument pogrešno
izdvoji. Obe rupe su reprodukovane na **postojećem** `modTest.bas`, ne sintetički.

**1. Lažne spoljne zagrade.** Skidale su se čim ostatak počne `(` i završi `)`:

```vb
AssertEq (modStornoDok.StornoIzvrsiMod(CStr(nisu(i)), "1/TEST", "", _
          SV_MODE_ISPRAVKA, False, False) Is Nothing), True, _
         "framework ne izvrsava nista nad: " & CStr(nisu(i))
```

To **nisu iste** zagrade. Posle skidanja dubina padne na −1, zarezi više nisu na
vrhu, argumenti se ne razdvoje — i ceo poziv prođe kao „poruka". Mereno:
`_tvrdnja_pripada("1/TEST", …)` je vraćalo **`True`**.

**2. Literali unutar ugnježdenih poziva.** `"|"` u `Split(x, "|")` i `"0.00"` u
`Format$(x, "0.00")` jesu literali, ali se **nikad ne ispisuju**. Mereno:
`_tvrdnja_pripada("|", T_Storno_UgovorIRadnje)` je vraćalo **`True`**.

Model koji to rešava je jednostavniji od pokušaja da se razumeju `Split`,
`Format$` i ostali: poruka se deli po **`&` na vrhu**, i statični su samo
operandi koji su **sami ceo literal**. Sve ostalo je rupa.

```
"Storno / " & tip & " cita svoju tabelu"   ->  literal · rupa · literal
"lista " & Split(x, "|")(0) & " ima ..."   ->  literal · rupa · literal
```

Spoljne zagrade se sada skidaju samo uz izričito `Call Ime(...)`, i to tek kad se
prva `(` stvarno zatvara na kraju.

### I jedan lažni pozitiv koji sam sam napravio

Posle sužavanja je iskočio `storno-nema-dok`. Nije bio stvaran: deklarisano je
`"kapija zaustavlja nepostojeci dokument"`, a poruka je
`"kapija zaustavlja nepostojeci dokument, tip " & CStr(tip)` — dakle deklarisan
tekst je **prefiks fragmenta**, i `dokaz.py` ga nalazi kao podniz poruke (žetva
mu je i dala `OK`).

Brzi put zato gleda i **pojedinačne fragmente** poruke, ne samo pune poruke.
Katalog sme da nosi prepoznatljiv deo, kao i do sada.

### Kako je dokazano da suženja grizu

| Sabotaža | Šta padne |
|---|---|
| bezuslovno skidanje spoljnih zagrada | **pozitivan** slučaj: ispravna poruka iza lažnih zagrada više nije prepoznata |
| svi literali izraza, ne samo operandi | slučajevi `Split` i `Format$` |
| **obe zajedno** | oblik iz review-a: literal iz prvog argumenta postaje tvrdnja |

Prva je pozitivan slučaj namerno: pod novim pravilom operanda naivno skidanje
više ne pravi lažno zeleno nego **gubi poruku** — a lažna uzbuna u hook-u je
skuplja od propusta, jer uči da se checker preskače.

Treći red je oblik koji je review reprodukovao: nastaje iz **obe** stare grane
zajedno, pa ga nijedna sabotaža sama ne proizvodi.

### Gde provera živi, i zašto baš tu

U `_nalazi` — jedinoj putanji kroz koju prolaze sva statička pravila kataloga,
kojoj `--self-test` podmeće izmišljene unose i koju `vba_check` zove posle
**svake** VBA izmene. Zaseban `--proveri-tvrdnje` bi bio još jedna putanja koju
neko mora da se seti da pusti, a baš VBA izmena je ono što tvrdnju čini
zastarelom.

Jedanaest novih self-test slučajeva (24 ukupno, bilo 13): zastarela tvrdnja,
tvrdnja koja pripada **drugom** testu, prazna tvrdnja, i tekst koji u testu
postoji samo kao **kod**, samo u **komentaru**, samo u **string promenljivoj**,
kao **očekivana vrednost** `AssertEq`-a, kao **literal iz prvog argumenta**, kao
**separator** u `Split`-u i kao **format-spec** u `Format$`-u — plus jedan
**pozitivan**: ispravna poruka iza lažnih spoljnih zagrada mora da se prepozna.

### Dva nalaza koja NISU zataškana

| Nalaz | Zašto stoji |
|---|---|
| `zbirna-vlasnik-samo-kupac` deli tvrdnju sa `oporavak-cilj-po-broju` | žetva pokazuje da obe obore **istu** poruku — test ih stvarno ne razlikuje |
| `ljuska-rez-bez-potvrde` ne obara ništa | `PostaviRez` piše pa **čita nazad** do tri puta; sabotaža svodi na jedan upis, što se u testu ne vidi jer prvi upis tamo uvek uspe |

Oba su upisana u `POZNATI_NALAZI` sa razlogom i vlasnikom, kao i zatečeni
`stale-parent-po-broju`. Druga nema „tačan tekst" koji bi se upisao: invarijanta
je otporna na *flaky* upis, pa je merljiva samo nad lažnom kontrolom koja prvi
upis odbija.

### Pun prolaz je izvršen — i našao je ono što statička provera ne može

Pušten je posle popravke, nad svih 251 sabotažom (27.08.2026):

```
crvenih: 247 / sabotaza: 251 (priznatih: 2)
izvor pre/posle: 2c1ee801fc99fe2a / 2c1ee801fc99fe2a -> IDENTICAN
```

Migracija je time potvrđena — **nijedan** od 119 prepisanih tekstova nije pao. Ali
prolaz je našao **osam** unosa koje statička provera **provereno ne vidi**, i
nijedan od njih nije bio u žetvi (svi su bili među 128 koji su se već poklapali):

| Klasa | Koliko | Zašto statička provera ne pomaže |
|---|---|---|
| `PALA DRUGA TVRDNJA` | 5 | tekst **jeste** tvrdnja tog testa, ali sabotaža obara **drugu** — obično preduslov koji pukne pre nje (zamka 6) |
| `NE OBARA NISTA` | 3 | sabotaža se uredno primeni, a nijedna tvrdnja ne padne |

Imena: `parse-cdate`, `bruto-prijemnica`, `guard-samo-aktivni-vlasnici`,
`completion-ne-prevezuje`, `brojac-nije-opcion` (prva klasa); `uvid-guta-necitljivo`,
`identitet-degradira-na-broj`, `paleta-klik-otvara` (druga).

> **Druga klasa je zatvorena** u `v2.81.0` — v. §14. Ostaje pet iz prve i jedan
> zastareo priznat upis. Uz njih i jedan
**zastareo priznat nalaz** — `POZNATI_NALAZI_DOKAZ['relink-ignorise-generaciju']`
više ne pokriva ništa, što znači da je nalaz koji opisuje u međuvremenu zatvoren.

To je **potvrda ograničenja**, ne iznenađenje: §12 je i pisao da provera tvrdi
samo da je tekst tvrdnja tog testa, a ne da je to tvrdnja koju ta sabotaža obara.
Sada se zna i **koliko** to ograničenje košta: osam od 251, oko tri odsto.

### Šta ostaje otvoreno

**Tih osam unosa nije popravljeno.** Svaki traži svoje merenje — pet ih traži ili
novu tvrdnju ili užu sabotažu koja ne obara preduslov, tri traže odgovor na
pitanje da li invarijanta uopšte može da se meri. To je zaseban posao, ne dodatak
ovom.

**Provera i dalje ne zna šta sabotaža stvarno obara.** To zna jedino `dokaz.py`, i
sada je izmereno da ta razlika nije teorijska.

---

## 13) `vba_check` nije video nedeklarisanu modul-promenljivu (27.08.2026)

Zapisano kao nalaz sa strane u `UI_MIGRACIJA_KATALOG` §16.6, zatvoreno ovde.

### Šta je bilo

Patch skripta piše fajl tek kad **svi** parovi zamena prođu, pa je pad na drugom
paru otkotrljao i prvi. Drugi patch je zatim upisao kod koji koristi
`m_BlokoviOk`, a deklaracije nije bilo.

`Option Explicit` to hvata — ali **tek pri compile-u**, a compile je ručna kapija
pred release. U međuvremenu:

```
vba_check: cisto        <- zeleno
run_vba:   visi         <- Excel stoji u [break], bez ijedne poruke
```

Najskuplji mogući kanal za grešku koja se vidi statički.

### Zašto ne pun undefined-variable checker

Zato što bi nad Excel objektnim modelom, kontrolama forme i `Enum` članovima
davao lavinu lažnih uzbuna — a lažna uzbuna u hook-u je **gora** od propuštenog
nalaza, jer uči da se checker preskače. To pravilo `vba_check` već nosi zapisano.

Provera je zato vezana za **konvenciju imenovanja** (`mFoo`, `m_Foo`), kojom se u
ovom projektu zovu modul-promenljive. Izmereno pre pisanja: **585** takvih
deklaracija u **68** fajlova, i nijedna se ne deli između modula.

### Nula lažnih uzbuna, i to mereno

Nad celim zatečenim kodom (195 fajlova) pravilo daje **0** nalaza. Do te nule se
stiglo kroz dve moje greške, obe uhvaćene istim merenjem:

| Greška | Posledica | Lek |
|---|---|---|
| `DEKL_POCETAK` je gutao `Private **Sub**` | ime procedure se nikad ne zapamti, pa je **svaki** poziv event handlera bio „nalaz" — 54 lažne uzbune | potpis se proverava **pre** deklaracije, uz negativan lookahead |
| komentar je trošio i prelom reda | brojevi redova su klizili za jedan po komentaru — nalaz pokazuje na tuđi red | prelom reda prepisuje spoljna petlja |

Prva je našla i sama sebe: 54 nalaza nad kodom koji se uredno kompajlira ne mogu
biti ništa drugo nego mana provere.

### Treća greška: doseg je bio izgubljen (iz review-a)

Prve dve su nađene merenjem. Treću je našao pregled, i bila je najozbiljnija —
jer je vraćala baš onu klasu zbog koje pravilo postoji.

Imena su se skupljala u **jedan ravan skup za ceo fajl**. Zato je ovo prolazilo
kao čisto:

```vb
Private Sub A()
    Dim mState As Boolean       ' lokalno u A
End Sub

Private Sub B()
    mState = True               ' NIJE deklarisano -- VBA nece prevesti
End Sub
```

Lokalni `Dim` u `A` legalizovao je `mState` kroz ceo modul. Isto je važilo za
**parametar** procedure `A`. Oba oblika su izmerena pre popravke i oba su davala
**0 nalaza** nad kodom koji se ne kompajlira.

To nije egzotičan slučaj: sam PR je proglasio legalnim da `m` prefiks nosi i
parametar, i lokalni `Dim`, i `Static` — pa se ne može reći „to ionako ne radimo".

Doseg se sada poštuje na dva nivoa, koliko i VBA traži:

| Nivo | Šta ulazi |
|---|---|
| **globalno** | deklaraciona sekcija (`Dim`/`Private`/`Public`/`Const`/`WithEvents`) + imena **svih** procedura |
| **po proceduri** | njeni parametri + njeni `Dim`/`Static`/`Const` |

**Suženje nije ništa pokvarilo:** i posle njega je **0** nalaza nad svih 195
fajlova, a rekonstrukcija incidenta i dalje daje 1.

### Dokaz je rekonstrukcija stvarnog incidenta

Ne sintetički fixture: uzet je `frmDokumenta.frm` sa `origin/main` i uklonjen je
**taj jedan red**.

| Fajl | Nalaza |
|---|---|
| zdrav | **0** |
| bez `Private m_BlokoviOk As Boolean` | **1**, imenuje `m_BlokoviOk` |

Uz to četrnaest self-test slučajeva (96 ukupno, bilo 82), od kojih je **jedanaest
nula** — svaki legalan oblik koji bi mogao da zapišti: ime procedure po istoj
konvenciji, parametar, parametri prelomljeni preko više redova, višestruka
deklaracija u jednom `Dim`-u, `Const`, `WithEvents`, kvalifikovano ime, ime u
komentaru, ime u tekstu.

### Dvosmerni dokaz

| Sabotaža | Šta padne |
|---|---|
| pravilo nije priključeno na `check_file` | **tri** slučaja koja traže nalaz |
| `Private Sub` opet prolazi kao deklaracija | **tri** slučaja: ime procedure i **oba** parametarska |
| doseg se gubi — lokalno postaje globalno | **oba** scope slučaja, a kontrolni „modul-nivo pokriva obe" ostaje zelen |

Drugi red je pošten po cenu urednosti: ime procedure i lista parametara dolaze iz
**iste** grane, pa vraćanje starog oblika gasi obe. Prvi pokušaj dokaza je zbog
toga prijavio „palo je i nešto drugo" — i to je bila greška u očekivanju, ne u
pravilu.

### Šta ovo NE pokriva

**Promenljiva van konvencije** (`blokoviOk` umesto `mBlokoviOk`) se ne hvata.
Pravilo je namerno usko: pokriva oblik koji je u ovom projektu standard za
modul-stanje, a to je i oblik koji je incident i proizveo.

**Lokalna promenljiva bez `Dim`** se ne hvata iz istog razloga — osim ako slučajno
nosi `m` prefiks.

---

## 14) Tri sabotaže koje ništa nisu obarale — tri različita razloga (27.08.2026)

Pun prolaz iz §12 našao je tri sabotaže koje se uredno primene a ništa ne padne.
Ispostavilo se da su to **tri različite bolesti**, i da je samo jedna od njih ono
što ime „mrtva sabotaža" sugeriše.

### Dve su merile tuđu kapiju

`identitet-degradira-na-broj` i `uvid-guta-necitljivo` gađaju kapije u
`modStornoImpact`. Obe su bile deklarisane nad tvrdnjom koju obara **ranija
sekcija** `BuildStornoImpact`-a, pa uklanjanje ciljane kapije nije menjalo ništa.

Ovo nije zaključeno čitanjem nego **mereno**: testu je privremeno dodata tvrdnja
`AssertEq m("greska"), ""`, pa je pad ispisao ko stvarno obara uvid.

| Sabotaža | Ko je stvarno obarao | Sekcija |
|---|---|---|
| `identitet-degradira-na-broj` | `modStornoFlow.PkPoIdentitetu` | `chain` / `blocks` / `flags` |
| `uvid-guta-necitljivo` | `modStornoFlow.CountActive` | `flags` |

Obe kapije već imaju svoje sabotaže drugde, pa nisu bile nepokrivene — nepokrivena
je bila **ciljana** kapija.

**Lek nije bio prepisati tvrdnju nego naći stanje u kome ciljana kapija JESTE
jedina koja odlučuje.** Oba postoje, i oba su poučna:

*Identitet.* `IdoviGeneracije` traži generaciju kroz **celu tabelu**, a
`PrijemniceIDPoIdentitetu` traži **broj i generaciju**. Zato postoji stanje u kome
prva prođe a druga ne — **generacija koja pripada drugom broju**. Tu se meri baš
kapija u `ImpactPalete`, i bez nje bi palete bile pročitane po broju, dakle tuđe,
unutar modela koji se posle označava kao valid.

*Drift šeme.* Raniji test gasi `PrijemnicaID` — ali po **baš toj** koloni filtrira
i `CountActive`. `PaletaID` u `BuildStornoImpact` čita **jedino**
`GetPaleteImpactByField`, pa tek njen drift meri strogost paletne sekcije.

### Treća uopšte nije bila mrtva — nije se kompajlirala

`paleta-klik-otvara` je pisala `Scr_Event = OtvoriStavke(...)` iz tela
`ObradiDogadjaj`. To je dodela imenu **tuđe** procedure, dakle compile error.

Posledica: Excel stane u `[break]`, suite se ne pokrene, `dokaz.py` ne vidi nijednu
palu tvrdnju — i prijavi **`NE OBARA NISTA`**. Isto što i mrtva sabotaža.

```
SUITE FAIL RunAllTests (67.7s)  (-2147352567, 'Exception occurred.', ...)
TESTS RunAllTests: suite nije upisala last_run.txt (nije stigla do kraja)
```

Ispravka je jedan red — dodela ide u `ObradiDogadjaj`. Ali razlika je velika:
prva vrsta se popravlja u jednom redu, druga traži rad nad testom, a **izveštaj ih
ne razlikuje**.

### Zato provera, a ne samo popravka

Novo statičko pravilo: zamena koja dodeljuje imenu procedure iz istog fajla, a
sidro joj **nije** u toj proceduri, je nalaz. Uže je nego što zvuči — pokriva tačno
oblik koji ne može da se prevede, a ne pokušava da bude kompajler.

Pušteno nad **originalnim** zapisom iz kataloga, ne nad izmišljenim:

```
KATALOG: paleta-klik-otvara: zamena dodeljuje imenu tudje procedure 'Scr_Event'
```

Od 251 zamene, njih 112 nečemu dodeljuje — i nijedna druga nije pogrešna.

### Šta ovo NE pokriva

**Druge vrste compile grešaka u zamenama** se i dalje vide tek kroz Excel u
`[break]`: nedostajuća zagrada, pogrešan broj argumenata, tip koji se ne slaže.
Pravilo pokriva jedan oblik — onaj koji se stvarno dogodio.

**Zamena koja dodeljuje imenu procedure iz DRUGOG modula** se ne hvata; traži se
samo u fajlu koji se sabotira.

