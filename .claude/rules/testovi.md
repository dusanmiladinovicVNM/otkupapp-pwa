---
paths:
  - "src-vba/mod*Tests.bas"
  - "src-vba/modTest*.bas"
  - "src-vba/modOtkupUI.bas"
  - "src-vba/modScr*.bas"
  - "src-vba/modUiScreens.bas"
  - "src-vba/modUiPanel.bas"
  - "src-vba/modMaticniEkran.bas"
  - "src-vba/modPodesavanja.bas"
  - "src-vba/modPrint.bas"
  - "tools/vba_check.py"
  - "tools/run_vba.py"
  - "tools/vba_parity_check.py"
  - "tools/vba_selfupdate_gates.py"
  - "tools/vba_hard_census.py"
  - "tools/who_writes.py"
  - "tools/gen_schema_module.py"
  - "tools/sabotaza.py"
  - "tools/make_fixture.py"
  - "tools/dump_schema.py"
  - "tests/golden/*"
---

<!-- modOtkupUI.bas / modScr*.bas su u paths namerno: u njima su test
     seam-ovi (§4). Bez ovoga agent koji ih menja ne bi znao da postoje.
     modScr* je glob, ne samo modScrDokumenti (Scr_OtpTestSet): kuke seam-a
     pada iz KI-016 stoje u 30 clanova ugovora, u svih 14 modScr* modula.
     Iz istog razloga i modUiScreens / modUiPanel / modMaticniEkran /
     modPodesavanja / modPrint: seam-ovi i tumac ishoda Run-a iz KI-011 i
     KI-014, a u modUiScreens i seam i kanal pada iz KI-016 (§4).
     frmOtkup.frm je otisao odavde sa formom (korak 2, §27.10).

     tests/golden/* trenutno ne pogadja nijedan fajl: jedini golden je bio snimak
     kontrola te forme. Mehanizam (AssertSnapshot / DumpKontrole u modTest) ostaje,
     pa unos stoji za sledeci golden. -->

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
JSON validnost `settings*.json`, `vba_check`, `who_writes --check`, tri kapije
nad self-update motorom (§8), i dve kapije nad šemom i vlasništvom (§9). CI ne
pokreće Excel i nikad neće.

**Šest** od tih alata ima i `--self-test` korak u istom workflow-u:
`vba_check`, `who_writes`, `run_vba`, `vba_parity_check`,
`vba_selfupdate_gates`, `vba_hard_census`. To nije udvajanje — zelen checker
nad čistim repoom ne razlikuje „nema greške" od „provera ništa ne meri".

> `gen_schema_module.py` **nema** self-test, i to je poznata rupa a ne izuzetak:
> `--check` koji nikad nije pokazan crven ne dokazuje da poredi otisak. Ranija
> formulacija („svaki od tih alata") je tvrdila suprotno.

## 2) `tools/vba_check.py` — radi svuda, i u Linux sesiji

Exit `0` = čisto, `2` = ima nalaza. **Obavezan pre commita** svake VBA izmene.

| Provera | Šta hvata |
|---|---|
| `ASCII` | ne-ASCII bajt u `.bas`/`.cls`/`.frm`/`.doccls` |
| `DEKLARACIJA` | modul-level `Const`/`Dim`/`Declare`/`Type`/`Enum` posle prve procedure |
| `REZERVISANO` | ime koje se case-insensitive poklapa sa VBA ključnom reči (`eNum`) |
| `DUPLIKAT` | isti `Public Sub/Function/Const` u dva modula → „Ambiguous name" |
| `DUPLIKAT_LOKALNI` | isto ime dva puta u **istom** modulu → modul se ne kompajlira |
| `PORUKA` | `Poruka("KLJUC")` bez para u `modPoruke.UpsertPoruke` |
| `NEDEFINISAN` | poziv procedure koja nigde nije definisana |
| `ARNOST` | poziv sa pogrešnim brojem argumenata |
| `ZAKLONJENO` | lokalni skalar zvan sa zagradom → „Expected array" |
| `KRAJ_REDA` | LF umesto CRLF u VBA izvoru → `Import` ne prepozna formu |
| `CLAN_FORME` | `frmX.Clan` gde forma nema taj javni član ni kontrolu |
| `DUPLI_LOKAL` | isto ime dvaput u **istoj** proceduri (`Dim src` uz `Const SRC`, parametar pa `Dim`) |
| `SEMA_REGISTAR` | `TBL_*` konstanta koje nema u registru `modSchema` — `EnsureAllTables` je ne bi napravio, a `VerifySchema` ne bi prijavio da fali |

**`Poruka()` nikad ne vraća prazno.** Nepoznat ključ daje `"[KLJUČ]"`, pa je
`Len(Poruka(k)) = 0` provera koja **ne može da padne**. Tvrdnja „ključ postoji u
katalogu" mora da meri tu oznaku (`PorukaNedostaje` u `modTest`), inače je
placebo — dve takve su živele zelene dok ih sabotaža nije razotkrila
(`UI_MIGRACIJA_KATALOG.md` §27.15). Statički `PORUKA` hvata samo
`Poruka("LITERAL")`; ključ sastavljen u kodu (`"OTKUI_OBL_" & oblast`) ili
sakriven u stringu (`Scr_Meta`) mu je **nevidljiv** i traži tvrdnju u testu.

`NEDEFINISAN`/`ARNOST` su namerno uske (samo `.bas`, samo poziv u poziciji
naredbe) — lažan nalaz je gori od propuštenog. **Poziv u izrazu (`x = Foo(1)`)
se ne proverava**: bez tipova se poziv funkcije ne razlikuje od indeksiranja
niza. To je poznata rupa, ne previd — pokušaj proširenja je dao 406 lažnih
nalaza (ime funkcije unutar string literala, između ostalog).

> **Posledica koju plaćaš ti, ne checker.** Kad dodaš parametar javnoj
> proceduri, `ARNOST` **neće** naći zastarele pozive u `AssertEq ...` — oni su u
> poziciji izraza. Suite tada ne padne nego **visi**: `run_vba` javlja
> „Exception occurred", Excel ostaje u `[break]`, a pravi razlog je
> „Argument not optional" u dijalogu. Posle svake izmene potpisa ručno pretraži
> pozive po imenu.

Uzak izuzetak od `DUPLIKAT`-a postoji za ugovor ekrana (`Scr_*` u `modScr*`), a
od `DUPLIKAT_LOKALNI`-og za `Property Get/Let/Set` trojku. Ne kompajlira VBA: ne
hvata tip-greške ni nedeklarisane promenljive, a u `.frm`/`.cls` ne radi
`NEDEFINISAN`/`ARNOST`.

**`ZAKLONJENO` postoji zbog rupe koju su ostale tri provere ostavile.** Lokalno
ime koje se poklapa sa imenom funkcije **zaklanja** je unutar te procedure (VBA je
case-insensitive), pa poziv postaje indeksiranje:

```vb
Public Function StornoIzvrsi(..., ByRef poruka As String, ...)
    poruka = Poruka("STORNO_MSG_OK")      ' Expected array
```

Dvanaest takvih poziva je živelo u dve procedure od `v6-ui-119` do `v6-ui-141`.
Nije ih videla **nijedna** postojeća kapija: suite ne, jer VBA kompajlira
proceduru **tek kad se pozove** a te dve je zvao samo UI; `ARNOST`/`NEDEFINISAN`
ne, jer je poziv u poziciji izraza; CI ne, jer ne pokreće Excel. Našao ih je
operater ručnim `Debug → Compile`.

Provera je namerno **uža** od „ime zaklanja funkciju": skalar eksplicitnog tipa se
u VBA ne može indeksirati nikako, pa je `ime(` uz `Dim ime As String` uvek greška.
Izostavljeni su `Variant` (može nositi niz), nizovi, objekti (default member) i —
najvažnije — sadržaj string literala i komentara: 14 od prvih 20 nalaza ovog
obrasca bilo je ime unutar teksta, ista klasa lažnih nalaza zbog koje je širenje
`ARNOST`-a odbijeno.

**`KRAJ_REDA` gleda BAJTOVE, ne dekodirane linije** — do ostalih provera je kraj
reda već izgubljen. `.frm` sa LF krajem reda `VBComponents.Import` **ne prepozna
kao formu**: uveze je kao standardni modul sa zaglavljem (`VERSION 5.00`,
`Begin {C62A...}`) kao **kodom**. Taj modul se ne kompajlira, a modul koji se ne
kompajlira obara ceo projekat — pa svaki makro javi `Cannot run the macro`, ne
compile grešku. Jedan `eol=lf` u `.gitattributes` je tako oborio svih 11
suite-ova dok je `vba_check` ostao zelen. Kraj reda sada drži `.gitattributes`
(`eol=crlf` za `.frm`/`.bas`/`.cls`/`.doccls`), a ova provera je kapija koja radi
**i u CI-ju, bez Excela**.

**`CLAN_FORME` postoji zbog rupe koju `NEDEFINISAN` ne pokriva.** Ta provera ne
radi nad `.frm`, a VBA kompajlira modul forme **tek kad se pozove** — pa poziv na
član koji ne postoji preživi i `vba_check` i celu suite. Zatečen slučaj:
`OpenContentFormPublic` obrisan iz `frmOtkupAPP` uz pogrešnu pretpostavku o
jedinom pozivaocu, dok ga je `frmMaticniPodaci` i dalje zvao — **FULL je bio
zelen na svih 11 suite-ova**, a našao ga je tek ručni `Debug → Compile`.

Uska je iz istog razloga kao `NEDEFINISAN`: gleda samo imena koja odgovaraju
postojećoj formi, preskače ime deklarisano kao promenljiva ili parametar u tom
fajlu (`Dim frmX As ...` — bez toga `modTheme` daje pet lažnih nalaza na
`frmBelow`), i pušta uzak spisak ugrađenih članova `UserForm`-a.

**Dve provere duplikata nisu ista provera.** `DUPLIKAT` gleda globalni imenski
prostor (isto `Public` ime u dva modula); duplo ime unutar jednog modula mu je
nevidljivo. A ono obara compile isto tako — i **ne prijavljuje se kao compile
greška** nego kao `Cannot run the macro` na bilo kom makrou, jer modul koji se ne
kompajlira obara ceo projekat. Simptom ne pokazuje na krivca.

`python tools\vba_check.py --self-test` vrti slučajeve koji **moraju** da zapište
i legalan VBA koji **ne sme**. Vrti se i u CI-ju: zelen checker nad čistim repoom
ne razlikuje „nema greške" od „provera ništa ne meri".

Slučajevi idu **kroz `check_file()`, istu funkciju koju zove CLI**, a jedan ide
kroz ceo `main()` nad pravim fajlom na disku. Namerno: self-test koji zove
proveru direktno dokazuje da funkcija radi, ali ne i da je CLI zove — otkačen
jedan red tada ostavlja i repo-run i self-test zelene, a checker isključen. Isti
oblik greške kao placebo test, samo u alatu.

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
- **Fixture je gitignored, pa ga `git checkout` NE menja.** Posle prelaska na
  drugu granu na disku ostaje sveska prethodne i testovi padaju *na podacima*, a
  pad izgleda kao regresija koda. Runner to sada hvata sam: `make_fixture` piše
  potpis podataka u `otkup_test.sig`, `run_vba` ga poredi pre Excela i staje uz
  komandu za regeneraciju. Bez važećeg potpisa run **ne prolazi** — jedini izlaz
  je svestan `--ignore-fixture-sig`. **Crveno posle prelaska grane — prvo
  regeneriši fixture, pa tek onda traži krivca u kodu.** Detalji:
  `docs/EXCEL_TEST_HARNESS.md`.
- `COMPILE NEJASNO` ne obara run kad suite-ovi idu. **Compile ostaje ručna kapija
  pred release:** `Alt+F11 → Debug → Compile VBAProject`.

## 4) Test seam-ovi koje produkcioni kod nosi

Ako menjaš `modOtkupUI`, bilo koji `modScr*`, `modUiScreens`, `modUiPanel`,
`modMaticniEkran`, `modPodesavanja` ili `modPrint`, znaj da ovo postoji i zašto —
inače ćeš ih „počistiti":

- `ClearForm` / `ParseDatum` / `ParcelaID` su **`Public`**, ne `Private` — test ih
  zove direktno, bez vožnje celog upisa.
- **Tri `SetFocus`-a** su iza `If Not IsTestMode()`. Forma bez `.Show` ne može da
  primi fokus, a u nevidljivom Excelu `SetFocus` **ne puca nego trajno visi**. U
  produkciji je `IsTestMode()` uvek `False`.
- `modScrDokumenti.Scr_OtpTestSet` postoji samo za test i **tvrdo je gejtovan** —
  van test-režima ne radi ništa.

**Zajedničko svim test seam-ovima ispod.** Seam sme samo da **pokvari, zabrani
ili posmatra** — nema vrednosti kojom dodaje pristup ili sposobnost. Dejstvo je
vezano za test režim **pri svakom pozivu**, ne samo pri postavljanju: postavljač
odbija rad van test režima, a kuka ponovo pita `IsTestMode()`. Zato seam koji
ostane postavljen postaje inertan čim `RunAllTests` vrati prethodni režim. Unutar
suite ga gasi `modTest.ResetSeamova` — na početku `RunAllTests` i posle testa
koji **padne** (`RunOne` → `CleanupPosleTesta`); test koji prođe gasi svoj seam
sam. **Nov seam koji nosi stanje ide i u `ResetSeamova`**, inače test koji pukne
pre čišćenja ostavi seam upaljen i obori tuđi test.

### Ishod `Application.Run` u ugovoru ekrana (KI-011, KI-014, KI-016)

- **Ugovor `Application.Run` ima dva nivoa.** Kod članova koje štiti `UI-PAD-01`
  (`Scr_Rows`, `Scr_Event`, `Scr_Build` — v. KI-016 ispod) čitalac
  (`ScrGridData`, `ScrEvent`, `ScrBuild`) **pre** poziva prazni kanal pada kroz
  `ScrPadPreuzmi`, a posle `Run`-a **najpre** ponovo preuzima kanal. Neprazan
  kanal znači da je član uhvatio izuzetak, i on ima **prednost** nad vraćenom
  vrednošću i nad `Err`. Tek kad kanal ne prijavi pad, `ScrRunIshod` tumači
  vraćenu vrednost i zaostali `Err` po `UI-ISHOD-01`. `ScrBuild` je Sub, pa
  posle kanala zadržava svoje tumačenje `Err`-a (dole). Test člana kome je
  skinuta zaštita zove član **direktno**: neuhvaćen izuzetak kroz `Run` otvara
  modalni VBA dijalog (KI-016 ispod).
- **`modUiScreens.ScrRunIshod` je drugi nivo: tumač vraćene vrednosti i `Err`-a**
  (invarijanta `UI-ISHOD-01`). Član se **nije izvršio** samo kad uz `Empty`
  stoji greška: 1004 (`SCR_RUN_NEMA`, člana nema) ili druga (`SCR_RUN_PAO`).
  Vraćena vrednost uz grešku u `Err` je **zaostatak** pomoćnog koraka, ne pad
  člana (`SCR_RUN_VRATIO`); `Empty` uz `Err = 0` je član koji je vratio `Empty`
  (`SCR_RUN_PRAZNO`) — šta to znači odlučuje čitalac (lista: „nema liste";
  `ScrSave`: neuspeh). Do KI-014 je svaki čitalac `Err` čitao po svom; upis
  reversa je tako javljao „Radnja nije uspela" posle uspešnog snimanja, pa je
  operater isti revers snimio dvaput (KI-011). **Nov čitalac ugovora ide kroz
  `ScrRunIshod`** (a nad članom iz `UI-PAD-01` prvo kroz kanal), ne kroz svoj
  `Err.Number`.
- Zove se **odmah** posle `Run`-a (posle kanala, gde ga ima). Namerno **nema
  `On Error`** (obrisao bi `Err` koji pozivalac još čita), a `v` je **`ByRef`**
  (`Scr_Rows` vraća celu mrežu). Ne „sređuj" ni jedno ni drugo. Isto važi za
  `ScrPadPreuzmi`: ni on nema `On Error`.
- `ScrBuild` (Sub) posle kanala i dalje čita `Err` direktno: Sub uz zaostatak
  izgleda isto kao nepokrenut član (sonda 14.10.2026). Imenovano ograničenje, ne
  previd.
- **`modUiScreens.ScrZaostalaGreska`** je javna promenljiva koju **nijedan
  produkcioni kod ne čita** — piše je `ScrRunIshod` (uz WARN u log), a tvrde je
  testovi 226, 227 i 236. Nije mrtav kod: bez nje se zaostatak guta bez traga, a
  sabotaža `ki014-pomocnik-guta-zaostatak` obara baš tu tvrdnju. Nosi
  **poslednji** zaostatak; prazni je samo `ScrSave`, na početku upisa. **Nije**
  test seam — piše se i u produkciji, pa je `ResetSeamova` ne dira.

### Seam-ovi puta upisa (KI-011)

- **`modOtkupUI.CommitDokumentTest`** vozi ceo potez „Sačuvaj" kroz
  `CommitDokument` — isti put kojim ide klik operatera, ne kopija. Ishod upisa
  (lažna greška, čišćenje forme, osvežavanje liste) se odlučuje **baš tu**; test
  samo nad `ScrSave` meri pomoćnik, ne poziv.
- **`modOtkupUI.Dijalog`** u test režimu **beleži** tekst umesto `MsgBox`-a, a
  `DijalogTest` ga vraća. `MsgBox` u nevidljivom Excelu visi do timeout-a — isti
  razlog kao `SetFocus`. Oba su tvrdo gejtovana; `CommitDokumentTest` sam prazni
  zabeležen dijalog, pa `ResetSeamova` tu nema šta da gasi.
- **`modPrint.StampaTestSet`** (`"pada"` / `"curi"`) kvari izvoz reversa tačno
  tamo gde je kvar nastao — posle foldera, pre izvoza. `"curi"` radi ono što je
  `LogSetup` radio pre KI-011: greška pod `Resume Next`, izlaz van handlera, koji
  je ne briše. Gasi ga `ResetSeamova` (`StampaTestSet ""`).

### Seam-ovi zaostatka u ugovoru ekrana (KI-014)

- **`modUiScreens.ScrCuriTestSet(clan)`** bira član ugovora koji „curi"; prazno
  gasi seam. Dve kuke dele to stanje:
  - **`ScrCuriAkoTest(clan)`** — član **vrati** rezultat i ostavi grešku u `Err`
    (oblik koji je sonda izmerila: proba pod `Resume Next`, izlaz van handlera);
  - **`ScrPraznoAkoTest(clan)`** — član vrati `Empty` uz `Err = 0`, oblik „ekran
    nema listu". Danas ga nijedan ekran ne proizvodi u normalnom toku, pa bez
    ove kuke grana `SCR_RUN_PRAZNO` u `ScrGridData` nema slučaj — prva tvrdnja o
    njoj je bila placebo (stavka 140 u `docs/STANJE_REFAKTORA.md`).
- Kuke stoje u **pravim** članovima, pa testovi 236–238 idu kroz pravi
  `Application.Run`, ne kroz kopiju: `modMaticniEkran.Redovi` (`REDOVI`,
  `REDOVI_PRAZNO`, `REDOVI_BEZ_LISTE`), `Dogadjaj` (`DOGADJAJ`), `ImaNesacuvano`
  (`NESACUVANO`) i `modPodesavanja.Podesavanja_ImaNesacuvano`
  (`PANEL_NESACUVANO`).
- **Kuka stoji na izlazu člana, POSLE poslednjeg `On Error` / `Err.Clear`.**
  Pomerena iznad, `Err.Clear` pojede podmetnutu grešku i seam postane inertan —
  test tada meri čist izlaz.
- Van test režima i za drugi član kuka ne radi ništa i **ne dira `Err`**. Gasi je
  `ResetSeamova` (`ScrCuriTestSet ""`); stanje je zajedničko, pa time gasi i
  `ScrPraznoAkoTest`.

### Seam-ovi pada člana ugovora (KI-016)

- **Zašto postoje.** Grešku koju član ugovora **digne** `Application.Run` ne
  vraća pozivaocu — ni pod `On Error Resume Next`, ni pod `On Error GoTo`, ni iz
  događaja prikazane forme. Otvara se runtime dijalog sa ugašenim „Continue", a
  „End" resetuje projekat: forma nestaje, Excel ostaje sakriven (sonda
  15.10.2026, stavka 142 u `docs/STANJE_REFAKTORA.md`). Jedina brana koju
  platforma dozvoljava je handler **u samom članu** (invarijanta `UI-PAD-01`).
- **Kanal pada nije seam, nego produkcioni kod.** Handler člana zove
  `modUiScreens.ScrPadZabelezi(clan)`, koji pad upiše u kanal i ERROR u log.
  Zapis nosi ime člana, broj i opis greške, i izvor kad nosi ime (korak funkcije
  liste); tačan oblik drži komentar nad `ScrPadZabelezi`, ne ovaj fajl.
  `ScrPadPreuzmi` vraća zapis i **briše** ga. Redosled čitanja je prvi nivo
  ugovora (gore). Zašto dvaput: pre `Run`-a čitalac odbacuje tuđ pad (član
  pozvan direktno, mimo čitaoca, ostavlja ga u kanalu), a ugnježden čitalac
  (radnja koja osveži mrežu) tako preuzme svoj. Zašto pre vrednosti: vrednost pad
  ne može da nosi — posle uhvaćenog pada je `Err = 0`, `Scr_Event` (Boolean)
  „vrati" `False`, `Scr_Build` (Sub) izgleda kao uspešna gradnja, a `Scr_Rows`
  vrati `Empty` kao „nema liste" ili listu koju je već dodelio.
- **`modUiScreens.ScrPadTestSet(clan)`** bira član koji „pada"; prazno gasi seam.
  Kuka **`ScrPadAkoTest(clan)`** tada digne grešku 9. Tvrdo je gejtovan:
  postavljač odbija rad van test režima, a kuka ponovo pita `IsTestMode()`. Van
  test režima i za drugi član kuka ne radi ništa i **ne dira `Err`**.
- **Kuka stoji na VRHU člana, odmah posle `On Error GoTo EH`** — obrnuto od kuke
  iz KI-014, koja stoji na izlazu. Tako podmetnut pad gađa branu koja pokriva
  celo telo. Iznad `On Error GoTo EH` pad prolazi mimo brane: direktan poziv tada
  izgleda kao član bez brane, a kroz `Run` je dijalog. Sabotaže
  `ki016-*-pusta-pad` sidre baš taj par redova. Kuku nosi 30 članova (obim iz
  stavke 142):
  - 13 `Scr_Rows` — svi osim ANALIZE, koja nema put do `Err.Raise`;
  - 4 `Scr_Event` matičnih (MatKorisnici, MatPakovanje, MatPartneri, MatRoba).
    Ostali imaju svoju zaštitu (osam sopstveni handler, Dokumenti
    `Resume Next`), a ANALIZA ga nema;
  - 13 `Scr_Build` — svaki koji postoji (Dokumenti ga nemaju).

  Druga kuka, `"modScrIzvestaji.Scr_Build#usred"`, stoji odmah iza povratka na
  branu posle lokalnog `Resume Next` segmenta (test 243). Član koji otvori
  `Resume Next` segment vraća se na `On Error GoTo EH`, ne na `GoTo 0` — kuka na
  vrhu to ne meri.
- Gasi ga `ResetSeamova` (`ScrPadTestSet ""`), koji **prazni i kanal**
  (`ScrPadPreuzmi`): zaostao zapis bi sledećem čitaocu javio tuđ pad.
- **Test člana BEZ brane zove član DIREKTNO, nikad kroz `Application.Run`.**
  Greška dignuta kroz `Run` i u testu otvara dijalog: `run_vba` visi do
  timeout-a, `dokaz.py` to beleži kao `TIMEOUT`, ne kao crveno, i niko ne sazna
  koji član je pustio pad. Direktno, greška stiže u test i on je crven **po
  imenu**. „Bez brane" je i član čiju branu skida sabotaža (`ki016-*-pusta-pad`
  gađa Fakture `Scr_Rows`, MatRoba `Scr_Event` i Storno `Scr_Build`), pa nova
  sabotaža brane bira član koji se u testovima ne zove kroz `Run`.
- **Testovi 239–243** i kako svaki poštuje to pravilo:
  - **239** — svaki od 30 članova hvata svoj pad i upisuje ga u kanal pod svojim
    imenom. Zove ih **direktno** (`Scr_Build` dobija `Nothing`, jer je kuka prva
    naredba tela) i posle svakog preuzme kanal, pa nalaz ne prelazi na sledeći;
  - **240** — pad liste je neuspeh liste, ne „nema liste", sa izvorom koji
    operater vidi. Ide kroz pravi `Run` (`GridTestLoad`, MatPartneri
    `Scr_Rows`), A → B → A: pad, obično čitanje, pa pad zabeležen mimo čitaoca
    koji ne sme da obori sledeće čitanje;
  - **241** — pad radnje je vidljiv na sva tri ulaza do `Scr_Event`: radnja
    ekrana, prekidač liste i izbor reda (`UiEvent`, MatPartneri `Scr_Event`);
  - **242** — pad gradnje ne prebacuje ekran, a sledeći ulazak gradi zonu
    ispočetka (`IdiNaEkran`, ANALIZA `Scr_Build`);
  - **243** — pad **usred** gradnje (kuka `#usred`, Izveštaji) ne ostavlja ni
    zonu ni omotače. Isti član mora i kroz `Run`, pa test prvo **direktno**
    izmeri da brana važi, a ljusku vozi samo ako važi — bez brane padne po imenu
    umesto da visi.

  240–242 gađaju članove koje **nijedna sabotaža brane ne skida** (MatPartneri,
  ANALIZA). Svih pet testova seam gasi sami; ako test pukne pre toga, gasi ga
  `ResetSeamova`.
- **Direktan poziv člana iz obima više ne diže grešku.** Vraća `Err = 0`, a pad
  je u kanalu. Test koji je pad merio kao `Err` posle direktnog poziva meri kanal
  (`PadDnevnika`, test 162, stavka 143). Zapis ostaje u kanalu i posle direktnog
  poziva: test ga preuzme sam, a čitalac ga ionako odbaci pre svog `Run`-a (240,
  poslednja tvrdnja).
- **Nov član iz obima** — nov `Scr_Rows` ili `Scr_Build`, ili `Scr_Event` bez
  sopstvenog handlera — po `UI-PAD-01` nosi branu, kuku i `ScrPadZabelezi` u
  handleru, i ide u spisak 239. Test nabraja članove po imenu, pa član van
  spiska niko ne meri.

Šta `modTest` pokriva, a šta ne: meri ugovore ekrana i UI ponašanje, uključujući
mrežu (npr. 236 i 240 čitaju listu kroz pravi `Scr_Rows` i produkcioni put
mreže). **Ne zamenjuje poslovne suite**: storno i sam `Save*_TX` (transakcioni
upis) pokrivaju `RunStornoTestSuite` i `RunBusinessFlowProSuite`. Forma se gradi
bez `.Show`, pa `UserForm_Activate` nikad ne ide.

## 5) Kapija ide i u writer, ne samo u modul unosa

Modul unosa proverava nad **snimkom iz trenutka kad je lista punjena**. Između
punjenja i potvrde stanje se može promeniti — drugi unos, uvoz izvoda, drugi
pozivalac istog writer-a — pa bi prošla prevelika isplata. Zato kritična
poslovna kapija (vlasništvo + aktivnost + **preračunat** preostali iznos) stoji
i u writer-u. Razlog je prvo bio legacy `frmDokumenta`, koji je writer zvao bez
ijedne UI provere; forma je otišla u koraku 2, ali pravilo ostaje — writer i
dalje ima pozivaoce mimo ekrana (uvoz izvoda, storno, poslovne suite).

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

## 8) Tri kapije nad self-update motorom

Sve tri su čist Python, rade u CI-ju i u web sesiji, i sve tri imaju
`--self-test`. Postoje zato što motor koji se sam ažurira ima invarijante koje
**suite ne može da izmeri**: greška se ne vidi kao crven test nego kao klijent
koji je snimio polu-nov projekat i više ne može da se ažurira.

| Alat | Čuva | Pada kad |
|---|---|---|
| `tools/vba_parity_check.py` | da obe kopije HARD/SOFT algoritma nose **isti kod** | `modSelfUpdate` i `modVbaTools` divergiraju, ili korpus fikstura ne daje očekivan verdikt |
| `tools/vba_selfupdate_gates.py` | da se `Save` **ne dosegne bez dokaza** | postoji put do `SaveWorkbookVerified` bez `VerifyReleaseProject` pre njega, ili bez `AbortSelfUpdate` između |
| `tools/vba_hard_census.py` | da **tvrda površina** ostane mala i namerna | tvrd `.frm`/`.doccls`/`.bas`, nova tvrda `.cls` van `WHITELIST`, ili mrtav unos u `WHITELIST`-u |

### Zašto paritet uopšte postoji

HARD/SOFT algoritam (`IsHardModuleBody`, `CanonCode`, `SameCode` i pratioci) živi
u **dve privatne kopije**. Spojiti ih se ne može: `modSelfUpdate` je frozen
bootstrap (`SKIP_MODULES`) i ne sme da zavisi ni od čega što se update-uje.

Kopije su **već divergirale**. PR #274 je popravio detekciju razmaka u
`modVbaTools`, `modSelfUpdate` je ostao sa rupom, a `modVbaTools` je pritom nosio
komentar „isti obrazac stoji i u `modSelfUpdate`" — koji više nije bio tačan
(zamka #22 u `docs/SELF_UPDATE.md`). **Komentar ne može da bude kapija; exit kod
može.**

Census **ne duplira** algoritam nego uvozi referentnu implementaciju iz
`vba_parity_check`. Treća kopija bi bila treća stvar koja može da divergira.

### Šta ove kapije NE dokazuju

Da se Python referenca ponaša isto kao VBA. Dokazuju da obe VBA kopije nose isti
kod i da taj kod odgovara pisanoj specifikaciji. Ponašanje nad živim VBE-om se
meri **ručno** — `RunSelfUpdateDev` nad **kopijom** prave sveske
(`docs/SELF_UPDATE.md`, „Lokalni DEV test"). Ta razlika se prijavljuje, ne
prećutkuje.

### `MRTAV_UNOS` — zašto whitelist mora da boli kad zastari

`WHITELIST` u `vba_hard_census.py` nabraja namerne event-sink klase. Pravilo
`MRTAV_UNOS` obara CI kad klasa iz liste **više nije** tvrda. To nije pedanterija:
lista koja se ne održava postaje spisak imena bez značenja, a prvo sledeće
stvarno zaprljanje prođe neopaženo. Isti obrazac je već jednom ujeo — komentar u
zaglavlju `modSelfUpdate` je nabrajao module (`modKarticaDetalji`, `modMouseWheel`,
`clsWheelList`) koji su u međuvremenu obrisani, a niko to nije primetio jer
komentar ništa ne obara.

## 9) Dve kapije nad šemom i vlasništvom

Obe su čist Python, rade u CI-ju i u web sesiji, i obe hvataju stanje koje
**suite ne može da izmeri** — jer greška ne postoji dok se ne pojavi na tuđoj
svesci ili u tuđem modulu.

| Alat | Čuva | Pada kad |
|---|---|---|
| `tools/gen_schema_module.py --check` | da su kanon (`schema/schema.json`) i `modSchema.bas` u koraku | neko izmeni jedno bez drugog — sveska bi se lečila po zastareloj šemi, a otisak na startu prijavljivao drift koji niko ne ume da objasni |
| `tools/who_writes.py --check-ownership` | ugovor A11 — ko sme da piše koju tabelu | pojavi se pisač van `WRITE_OWNERSHIP.json` |

### Zašto ownership kapija meri baš mutatore

`AddTableSnapshot` znači „moja transakcija mora da ume da vrati ovu tabelu", ne
„ja sam pišem". Koordinator sme da snapshotuje tuđu tabelu i pozove API njenog
vlasnika. Kapija zato broji **mutatore**: `AppendRow`, `UpdateCell` i
`DeleteRow`, svaki i u `Require` obliku.

Prva verzija je merila pomešano, a `RequireUpdateCell` joj je bio **nevidljiv**:
regex je tražio granicu reči pred `UpdateCell`, a u `RequireUpdateCell` je nema.
**220 poziva** je prolazilo neopaženo — uključujući `modSEFPersistance` nad
`tblFakture`, koji u mapi uopšte nije postojao. Ista klasa greške kao placebo
test: kapija zelena, a meri pola stvarnosti.

### Otisak šeme — zašto nad prefiksom, a ne nad celim zaglavljem

`SchemaCheckOnStart` poredi otisak sveske sa kanonom. Meri se **kanonski
prefiks** zaglavlja, iz dva razloga:

1. `modSetup.EnsureRuntimeSchema` dodaje kolone na svakom startu i one idu **na
   kraj**. Legitimne su — samo ih kanon još ne drži — pa bi otisak nad celim
   zaglavljem prijavljivao trajan lažan drift. Provera koju operater nauči da
   ignoriše ne štiti ništa.
2. Pozicioni upis zavisi **tačno** od prefiksa: sve iza poslednje kanonske
   kolone ne može da pomeri nijednu vrednost koju pisci šalju.

Time se otisak i `VerifySchema` poklapaju **po konstrukciji**. Kad se ne poklope,
to više nije drift u svesci nego kvar u kodu — i poruka tako i glasi.

### Šta ove kapije NE dokazuju

Da je kanon **tačan**. Dokazuju da su kanon, modul i sveska međusobno u koraku.
Da li je sam kanon ispravan model pokazuju domenski testovi (`modTest` 197–199) i
`docs/DOMEN/DOCUMENT_HEADER_LINES.md`.
