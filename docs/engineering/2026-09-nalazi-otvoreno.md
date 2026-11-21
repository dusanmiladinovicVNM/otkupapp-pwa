# Otvoreni nalazi: storno, self-update, Drive, migracija

Registar nalaza koji **još stoje** protiv `main @ c60fd9de` (21.11.2026). Nastao je
re-verifikacijom liste od 23 nalaza (revizija avgust 2026) posle refaktora
S3c-S5-3b, koji je identitet storna prebacio sa labele/broja na `ZbirnaID` i
kanonsko članstvo.

| Revizija | Baza | Promena |
|---|---|---|
| 26.09.2026 | `1095fdd` | prvi registar: 14 nalaza + 2 reziduala + 6 strukturnih |
| 20.10.2026 | `8e970d9` | +387 commit-a (nov model ambalaže, S6 prijemnica). **Nijedna stavka registra nije zatvorena.** `#17` i Strukturno ispravljeni nadole, `#9` i `#20` preformulisani, `#6` i `R1` izoštreni |
| 29.10.2026 | `b552892` | +96 commit-a (S6 koraci 4–8). **`#4` zatvoren** — prvi zatvoren nalaz u registru. `R1` pola zatvoreno, `#9` i `#21b` suženi, Strukturno pogoršano |
| 21.11.2026 | `c60fd9de` | +558 commit-a (S7–S9). **`#21b` i `R1` zatvoreni, `#6` simptom zatvoren.** `#9` sužen na primitivu. Drive/self-update klaster nedirnut treći period |

Čita se pre nego se otvori rez u `modStorno*`, `modSelfUpdate`, `modDrive` ili
`modMigracija` — da se ne krpi ono što je već zatvoreno i ne preskoči ono što
nije. Numeracija je **zadržana iz originalne liste**, pa `#4` ovde je `#4` tamo;
brojevi koji se ne pojavljuju su verifikovano zatvoreni i namerno se ne ponavljaju.

## Nivo verifikacije ovog dokumenta

| | |
|---|---|
| Metod | statičko čitanje izvora na `main @ c60fd9de` |
| FAST | `python tools/vba_check.py` → čisto (195 fajlova) |
| Šema | `python tools/gen_schema_module.py --check` → u koraku (otisak `F22278A1`) |
| Vlasništvo | `python tools/who_writes.py --check-ownership` → nema novih pisaca (38 tabela) |
| TARGETED / FULL | **NIJE izvršeno** — `run_vba` traži Windows + Excel + `pywin32` |

Nijedna tvrdnja ovde **nije bihevioralno verifikovana**. Svaka je čitanje koda sa
navedenom putanjom i linijom. Gde je zaključak izveden a ne izmeren, tako i stoji.

## Gde ide pažnja, a gde ne

`modDrive.bas` i `modSelfUpdate.bas` **nisu dirani ni jednom u 1041 commit-a**,
kroz sva tri perioda:

| Period | Commit-a | `git log <opseg> -- modDrive.bas modSelfUpdate.bas` |
|---|---|---|
| `1095fdd..8e970d9` | 387 | prazan |
| `8e970d9..b552892` | 96 | prazan |
| `b552892..c60fd9de` | 558 | prazan |

Svih sedam nalaza u tom klasteru — `#5`, `#7`, `#8`, `#13`, `#14`, `#19` i
self-update polovina `#18` — zato stoje **bajt-identično**, sa istim brojevima
linija kao 26.09. `#5` je jedan red, i prošlo je skoro dva meseca.

Storno, ambalaža, prijemnica i faktura su u istom vremenu dobili nov model
knjiženja i refaktor do S9, i tamo se registar stvarno pomerio: **četiri od
petnaest stavki zatvoreno, sve četiri u tom sloju.** Nije u pitanju da su nalazi
iz Drive/self-update klastera odbijeni; nisu bili ni pogledani. To je raspored
prioriteta, i vredi ga videti kao raspored a ne kao ocenu.

> Vredi reći i obrnuto: tri revizije pokazuju da se **ono što je u putu aktivnog
> refaktora zatvara samo od sebe**, a ono što nije ne zatvara se nikako. Nalaz
> van putanje posla ne čeka prioritet — čeka da neko odluči da ga uzme.

---

## P0 — propušta pogrešan podatak tiho

### #4 Faza B kaskade bez verifikacije — ZATVOREN 29.10.

Jedini zatvoren nalaz u registru. Ostaje zapisan jer pokazuje kako se ova klasa
zatvara: **uklanjanjem dvofaznosti**, ne dodavanjem brojača.

Faza B više ne postoji — `grep "Faza B: paletne stavke"` nad `modStornoFlow.bas`
ne vraća nijedan pogodak, a `DetachOsirocenePaletaStavke_TX` se iz kaskade više ne
zove. Palete se sada skidaju **unutar transakcije pozivaoca**:

`PonistiZbirnaChain_TX` (`modStornoFlow.bas:2178-2205`) zove `StornoPrijemnica(id,
tx, True, palK)` i sabira `palUkupno`, pa `res("pals") = palUkupno` i
`res("ok") = True` stoje **posle** `tx.CommitTx`. `StornoPrijemnica`
(`modStorno.bas:477`) prima `ByRef outPaletnihStavki` i delegira
`modPaletniList.SkiniPaleteZaStornoPrijemnice(tx, prijemnicaID, SRC)`
(`modPaletniList.bas:2032`).

Ta primitiva zatvara sva tri otvora iz nalaza:

| Otvor iz nalaza | Kako je zatvoren |
|---|---|
| callee guta grešku (`EH: ... = 0`) | **nema `On Error` ni `EH:` granu** — svaki put diže grešku |
| petlja preko prijemnica nije atomična | traži `tx` pozivaoca i diže grešku bez njega — jedna transakcija za ceo lanac |
| `expPals` se ne računa | broj dolazi **iz same mutacije** (`outPaletnihStavki`), ne iz nezavisnog prebrojavanja |

Uz to je prešla na identitet: bira po `COL_PALS_PRIJEMNICA_ID`
(`:2053`), ne po labeli — čime je zatvorena i paletna polovina
`R1`. Review #438 je dodao i dve kapije koje ranije nisu postojale: paletna stavka
koja pokazuje na paletu koje nema obara storno **po imenu**, a prerađena paleta se
odbija pre dijaloga.

`res("ok") = True` je time nedosežan na neuspeh: svaki pad ide kroz `Err.Raise` →
`tx.RollbackTx`, a razlog stiže do operatera (`StornoRazlogOperateru`).

> Nalaz je bio pisan kao „prebroj `expPals`, uporedi". Zatvoren je bolje od
> predloga: dvofaznost koja je zahtevala brojanje je uklonjena, pa brojanje nije
> trebalo.

### #5 modSelfUpdate: prozor faza1 → faza2

`modSelfUpdate.bas:203-229` — posle `proj.VBComponents.Remove` (`:213`) projekat
je polomljen u memoriji. Sledi upis registry state-a, `Application.OnTime Now +
TimeSerial(0, 0, 2)` (`:229`), `Exit Sub`. Dve sekunde sa polomljenim projektom i
**postavljenim dirty flag-om**.

`ThisWorkbook.Saved` se dira samo na abort putanjama: `:545` (`AbortSelfUpdate`) i
`:568` (`AbortSelfUpdateClose`). `PrepareRuntimeForSelfUpdate` gasi VBA `OnTime`
tajmere (`StopAutoSaveTimer` među njima), **ali ne Excelov cloud AutoSave**. Za
OneDrive/SharePoint klijenta u tom prozoru Excel sam upiše polu-nov build preko
ispravne kopije na disku.

**Minimalna zakrpa:** `ThisWorkbook.Saved = True` pre `Exit Sub` na `:229`.

Lek za cloud AutoSave već postoji u repou — `modMigracija.bas:81-97` radi tačno
to, sa late-bind obrascem koji izbegava compile grešku na Excel-u pre 2016. Isti
obrazac se prenosi.

> Ovo je klasa koju self-update **ne može sam da isporuči**: ispravka živi u
> modulu koji se ažurira mehanizmom koji se ispravlja. Vidi `modSelfUpdateBoot` u
> Strukturno.

---

## P1 — funkcija ne radi, ili garancija ne važi

### #6 Kaskada van žurnala — simptom ZATVOREN 21.11., koren stoji

**Simptom je zatvoren, i to brisanjem.** S9-2 je obrisao oba mesta koja su čitala
kolone bez pisca, i u komentarima navodi **tačno ovaj nalaz** kao razlog:

- `OtkupBlockDeadParentByID` — obrisana; `modStornoZurnal.bas:312` kaže: „čitala je
  `Otkup.OtpremnicaID` i `Otkup.BrojZbirne`, kolone bez pisca, pa je na kanonskim
  podacima uvek vraćala `""`".
- `FreeOtkupBloksInline` — obrisan; `modStornoFlow.bas:2149` kaže: „brisao je staru
  vezu `Otkup.OtpremnicaID` — kolonu koju kanonski pisac ne piše, pa je na kanonskim
  podacima bio no-op".

Zamena nije drugi čitalac nego **konstrukcija**: u kanonu to stanje ne nastaje —
`StornoOtkup` odbija člana aktivne otpremnice, a pisac članstva odbija storniran
otkup. Potez u lancu čuva `LancaRazlog` (`modStornoZurnal.bas:141`).

Time je i „šest čitalaca, nula pisaca" iz revizije 20.10. otišlo: `COL_OTK_BROJ_ZBIRNE`
i `COL_OTK_OTPREMNICA_ID` sada nemaju **ni jednog čitaoca ni pisca** u celom
`src-vba/`. Ostale su u `schema/schema.json` — dve mrtve kolone u kanonu. To je
higijena šeme, ne tiha greška, i tako je i treba tretirati.

**Koren stoji.** Kaskada je još van žurnala: `BeginStornoOp`
(`modStornoZurnal.bas:42`) ima **jednog živog pozivaoca**, `modStorno.bas:154`
(otkup). Zbirna, otpremnica i prijemnica kaskade ne otvaraju žurnal operaciju, pa
kaskada nije undo-abilna kao celina. `JournalCell` ima tri pisca:
`modStorno.bas:155` (`tblOtkup.Stornirano`), `modNovac.bas:2410` i `:2560`.

Ovo je sada **jedini otvoren deo `#6`**, i on nije zakrpa nego odluka: ili kaskada
dobije operaciju u žurnalu, ili se „Vrati storno" za lanac svesno odbija porukom
koja to kaže.

### #7 modDrive: nema retry-ja

Nula pogodaka za retry / backoff / 429 / attempt / `Sleep` u celom `modDrive.bas`.
Jedan 429 ili 5xx na jednom od ~100 fajlova ruši ceo update; istekao token isto.

### #8 `DriveListFolder` odseca na 1000 — ostalo u legacy grani

Odsecanje stoji: `modDrive.bas:250`, `pageSize=1000`, nula pogodaka za
`pageToken`/`nextPageToken`. Nije propust nego **dokumentovana pretpostavka** —
`modDrive.bas:234` je i imenuje: „Jedna strana (pageSize=1000); AgriX_Release ima
~100 src-vba fajlova." Nalaz je da pretpostavka **nije iznuđena**: ništa ne tvrdi
da listing nije odsečen.

Posledica je promenjena na **živom** kanalu: manifest-driven put broji `expected`
iz `files.Keys` manifesta, a fajl kojeg listing nije dao → `LogErr "manifest fajl
nije na Drive-u"` → `n < expected` → **fatalni prekid** (`modSelfUpdate.bas:133`).

Tiho-pogrešan `expected` preživeo je **samo u LEGACY grani** (flat publisher bez
`files[]`), gde se `expected` još broji iz `dict.Keys`
(`modSelfUpdate.bas:766-773`). Tu odsecanje i dalje tiho spušta očekivani broj.

### #9 `LookupActiveID` na duplikat tiho uzima poslednji — sužen na primitivu

**Primitiva nepromenjena.** `modStorno.bas:1694` prolazi **sve** redove i pamti
**poslednji** aktivan pogodak: petlja na `:1725-1732` prepisuje `resultId` do kraja
tabele, bez `Exit For` i bez detekcije duplikata.

**Komentar ZATVOREN 21.11.** Oba netačna komentara („uzima prvi pogodak") nema više
nigde u `src-vba/` — `grep` ne vraća pogodak. Nisu ispravljeni nego uklonjeni sa
kodom nad kojim su stajali, isto kao drugi od njih u reviziji 29.10.

**Pozivaoci po broju spali sa tri na dva** (storno tok više nije među njima):

| Mesto | Tabela | Beleška |
|---|---|---|
| `modStorno.bas:814`, `:816` | `tblNovac` | po `BrojDokumenta`, bez kapije |
| `modDokumentInvariant.bas:215` | prosleđena tabela | čita `IzdatoStatus` po broju |

`modDokUnos` i `modStornoDok` pozivaoci, koje su prve dve revizije navodile, otišli
su sa S6–S9 refaktorom. Treći pozivalac (`modStornoFlow.bas:1423`) bira po
`COL_OTP_ID`, dakle po jedinstvenom ključu — nije izložen.

Nalaz time pada sa P1 na **tihu primitivu sa dva izložena pozivaoca**. Vredi je
zatvoriti kad se dira `tblNovac`, ne zbog sebe same.

### #13 `Sha256File` puca na 0 bajtova — ostala globalna degradacija

Funkcija nepromenjena: `modDrive.bas:145-167`. Na fajlu od 0 bajtova `stm.Read`
vraća `Null`, dodela u `Byte()` pukne, `EH` vrati `""`. **Izvedeno čitanjem, nije
izvršeno** — VBA se u Linux sesiji ne pokreće.

Po fajlu je potrošač sada fail-closed: nesklad → `LogErr "SHA-256 NESKLAD"` →
`n < expected` → fatalno.

Ostaje **globalna** degradacija, i ona je prijavljena a ne tiha: `Sha256Available`
(`modSelfUpdate.bas`) probuje 3-bajtnim fajlom, i na neuspeh ceo download pada na
prisustvo/broj uz notu „SHA-256 nedostupan ... fallback na prisustvo/broj". Update
svejedno ide dalje — lanac poverenja je tada samo najavljen, ne proveren.

Nezavisno od toga, primitiva konflatira dva stanja: „prazan fajl" i „SHA
nedostupan" oba izlaze kao `""`.

### #14 Nema brisanja komponenti

Nema nikakve logike za brisanje komponente koje u release-u više nema.

`VerifyReleaseProject` prolazi kroz **release fajlove** i za svaki proverava da
postoji u projektu. Nikad ne prolazi kroz **projekat** da nađe komponentu bez
release fajla. Modul obrisan iz repoa zato ostaje kod klijenta zauvek, **i provera
to ne prijavi**.

**Najmanji delta:** proširi `VerifyReleaseProject` da prođe i kroz
`proj.VBComponents` i **prijavi** ekstra komponentu (tip 1/2, bez `SKIP_MODULES`).
Brisanje je odvojena odluka; prijava je jeftina i zatvara tišinu.

> Jedini nalaz u registru koji **raste sa vremenom**: svaki obrisan modul je
> trajan rep kod klijenta.

---

## P2 — performanse, higijena, rast

### #16 `GetNedovrseno` bez batch keša

`modStornoRecovery.bas:51` — više prolaza kroz tabele bez `BeginTableCache`.

Frekvencija je popravljena (`Scr_Brojac` se zove iz `RefreshFromData`, ne pri
svakom crtanju sidebara — `modScrOporavak.bas:162`), pa je hitnost mala. Sam
prolaz je još neoptimizovan.

### #17 `tblStornoVeze` i `tblStornoZurnal` bez retencije

Dve append-only tabele, nula pogodaka za retenc / arhiv / prune.

**Ispravka 20.10.** Prva revizija je tvrdila da žurnal „raste najbrže od svega u
svesci". To **nije bilo tačno** ni tada, a sada je i manje: `JournalCell` ima dva
živa pisca — `modStorno.bas:151` (`tblOtkup.Stornirano`) i `modNovac.bas:2410`
(`tblNovac.OtkupID`). Revers je izašao iz žurnala u `10e-2`. Dakle oko dve ćelije
po stornu otkupa, ne „najbrže od svega".

`tblStornoVeze` raste po ispravci, što je red veličine ređe od storna. Nalaz zato
ostaje otvoren kao **higijena bez roka**, ne kao rastući rizik.

### #18 Backup folder bez retencije

Oba pisca upisuju u `<path>\Backup` sa timestamp imenom i **nijedan ne čisti**:

- `modSelfUpdate.bas:644` `MakePreUpdateBackup` → `AgriX_pre-update_*.xlsm`
- `modMigracija.bas:708` `BackupPreMigracije` → `AgriX_pre-migracija_*.xlsm`

**Predlog:** jedan zajednički `OcistiStareBackupe(dir, 3)`, zvan sa oba mesta
posle uspešnog `SaveCopyAs`. U tree-u ga nema.

### #19 `DownloadNamedText` fiksna temp putanja

`modSelfUpdate.bas:820` — `Environ$("TEMP") & "\AgriX_dl_" & fileName`, bez PID-a
ni GUID-a. Dve otvorene kopije AgriX-a se sudaraju nad istim fajlom.

Scenario je u istom modulu već priznat: `QualifiedProc` postoji baš zato što dve
otvorene kopije moraju da razlikuju `OnTime` cilj.

### #20 `Trim$` nekonzistentnost u `BeginStornoOp`

`modStornoZurnal.bas:42` zahteva non-blank **samo** `docType`; `broj` nikad.
Uporedba je golo `StrComp(broj, mBroj, vbTextCompare)` bez `Trim$`, i `mBroj =
broj` čuva netrimovano.

**Suženo 20.10.** Revers je u `10e-2` izašao iz žurnala (storno mu je sada
kontra-stav nad `AmbDokID`-em), pa je `BeginStornoOp` ostao sa **jednim živim
pozivaocem**: `modStorno.bas:154`. Šalje netrimovano — `brDok =
NzToText(LookupValue(...))` na `:153`, a `NzToText` (`modHelpers.bas:249`) ne
trimuje.

Od dve posledice ostaje jedna i po:

- **lažno „mešanje operacija"** — traži dva ugnježdena poziva istog dokumenta sa
  različito razmaknutim brojem. Sa jednim pozivaocem i jednim `otkupID` po pozivu
  put je sada uzak; nije dokazano zatvoren, ali nije ni lako dosežan.
- **op sa praznim brojem** — stoji nepromenjeno: `broj` nije ni `Trim$`-ovan ni
  tražen kao non-blank, pa otkup red bez `BrojDokumenta` otvara operaciju sa
  praznim brojem, a `LatestOpFor(docType, broj)` je posle traži po tom istom
  praznom broju.

### #21 Ostao samo `CanStorno`

**(a) Stale snapshot — ČISTO** od prve revizije; ne dira se.

**(b) Filter `IsStorniranoValue` — ZATVOREN 21.11., brisanjem potprograma.**
`StornoFakturaStavke` više ne postoji. S7 rez 2 (`FAK-ODL-03`) ga je obrisao jer je
linija aktivna **ako i samo ako je aktivna njena faktura**, pa je storno zaglavlja
ujedno storno linija — nema upisa `Stornirano` po liniji, pa nema šta da se
re-markira. `modStorno.bas:1750` to i kaže: „Do reza 2 je ovde stajao
`StornoFakturaStavke`, koji je svakoj liniji pisao `Stornirano`".

Ostalo je samo `OslobodiUtovareFakture` (`modStorno.bas:1755`), koje **ne piše
`Stornirano`** — oslobađa marker utovara, jer je činjenica utovara stvarna a ne
izvedena. `relUt` dikt štiti od duplog poziva unutar prolaza, a
`ReleaseUtovarFromFaktura` je fail-closed na dupli ID i oslobađa samo utovar koji
tvrdi baš tu fakturu.

**(c) `CanStorno` guta schema greške — STOJI, nepromenjeno.**
`modStorno.bas:1672`: `EH: ... CanStorno = False`. Schema drift, nečitljiva tabela i
„nije dozvoljeno" izlaze kao ista vrednost. Pozivalac nema način da ih razlikuje.
Ovo je sada jedini otvoren deo `#21`.

---

## Reziduali koje je re-verifikacija otkrila

Nisu iz originalne liste — izašli su iz nalaza koji su inače zatvoreni.

### R1 Kaskada je birala PO BROJU — ZATVOREN 21.11.

Paletna polovina je zatvorena 29.10. (prelaz na `COL_PALS_PRIJEMNICA_ID`),
prijemnička 21.11.: `ActivePrijIDsByZbirna` je zamenjena
`ActivePrijIDsByZbirnaID` (`modStornoFlow.bas:1896`), koja bira **identitetom** i
jeste **fail-closed** — zbirna čiji se skup prijemnica ne zna (aktivna prijemnica
bez `PrijemnicaID`-a) diže grešku, uz obrazloženje koje vredi citirati: „skup koji
kaskada stornira, a pregled obećava, ne sme da se pogađa".

> **Ispravka moje prognoze.** Revizija 29.10. je tvrdila: „ne može se zatvoriti
> istim potezom — `tblPrijemnica` ne nosi `ZbirnaID`, veza je `BrojZbirne`".
> Premisa je bila tačna (kolone i dalje nema), **zaključak pogrešan**: zatvoreno je
> kroz **javnog kanonskog čitaoca** (`modAutoHladnjaca.AktivnePrijemniceZbirnih`,
> ivica E4), ne kroz novu kolonu. Pretpostavio sam da identitet mora da stoji na
> redu da bi se po njemu biralo; dovoljno je da postoji čitalac koji ga zna.

### R2 Migracija ne čisti zatečenu aktivaciju u NOVOM šablonu

Dokumentovano u samom kodu, `modMigracija.bas:649`: „ovo ne CISTI eventualno
zatecenu aktivaciju u NOVOM sablonu (target-clean)". Licenca u odredišnom šablonu
preživi migraciju.

---

## Strukturno — granice, ne bagovi

| Stavka | Stanje |
|---|---|
| `modSelfUpdateBoot` (~80 redova, frozen) | **ne postoji** — i dalje razlog zašto `#5` self-update ne može sam da isporuči |
| `Sha256File` van frozen granice | **stoji** — verifikator lanca poverenja je u *updatable* `modDrive` |
| `modStornoFlow` podela na tri modula | **nije** — **2435 linija** (29.10. bilo 2513, 20.10. bilo 2431, 26.09. bilo 2585) |
| Generici u `modDataAccess` | **nema ih, i ima ih sve više** — `CountActive` (`modStornoFlow.bas:2331`), `DistinctActiveValues` (`:2370`), `NzTx` (`:2420`), `FindRowIndexByID` (`modPaletniList.bas:1348`), a S7–S9 su dodali `NzTx` i u `modDokumentInvariant.bas:156` i `modTestStorno.bas:1057`. **Šest privatnih kopija u četiri modula, nula javnih.** |
| `modMigracija` dry-run režim | **nema** — nula pogodaka za `dryRun` |
| `modMigracija` verzijska svesnost | **stub stoji** — `StaroImeKolone = novoIme` (`modMigracija.bas:471-472`); `APP_VERSION` starog fajla se ne čita |

**Kretanje `modStornoFlow` kroz četiri revizije:** 2585 → 2431 → 2513 → **2435**.
Fajl se već tri meseca drži oko 2500 i **nije podeljen**. Oscilacija od ±80 nije
trend ni u jednom smeru — to je refaktor koji unosi i iznosi, nad modulom čija
veličina stoji na mestu.

Metrika koja se **jeste** pogoršala u ovoj reviziji su generici: sa četiri privatne
kopije na šest, u četiri modula umesto dva, i dalje nijedna javna. S7–S9 su dvaput
rešili isti problem lokalno. `NzTx` je sada u `modStornoFlow`,
`modDokumentInvariant` i `modTestStorno` — tri identična tela.

---

## Predložen redosled

**Izmenjen 21.11.** `#21b` i `R1` su zatvoreni, `#6` je ostao bez simptoma, `#9` je
spao na primitivu. Redosled se time **prevrnuo**: posle tri revizije ništa iz
storno sloja nije na vrhu — sve što je bilo u putanji refaktora zatvorilo se samo,
a ostalo je ono čega refaktor ne dotiče.

1. **`#5`** — jedan red (`ThisWorkbook.Saved = True` pre `Exit Sub` na
   `modSelfUpdate.bas:230`). Najjeftinija stavka u registru, jedina koja može da
   ostavi klijentu polomljen fajl, i **1041 commit-a nedirnuta**. Posle tri
   revizije ovo nije više „nije stiglo na red" nego odluka da se ne radi.
2. **`#21c`** — `CanStorno` razdvoji schema grešku od „nije dozvoljeno". Mala,
   tiha, i jedini preostali deo `#21`.
3. **`#14`** — prijava ekstra komponente u `VerifyReleaseProject`. Jedini nalaz
   koji **raste sa vremenom**: svaki obrisan modul je trajan rep kod klijenta, a
   S7–S9 su obrisali više modula.
4. **`#6` koren** — odluka, ne zakrpa: kaskada dobije operaciju u žurnalu, ili se
   „Vrati storno" za lanac svesno odbija porukom koja to kaže.
5. **`#20`** — `Trim$` i non-blank na `broj` u `BeginStornoOp`; „op sa praznim
   brojem" još stoji.
6. **`#9`** — primitiva; zatvoriti kad se dira `tblNovac`.
7. **Dve mrtve kolone u šemi** (`tblOtkup.BrojZbirne`, `.OtpremnicaID`) — nula
   čitalaca, nula pisaca. Higijena kanona, uz prvi sledeći rez šeme.

`#7`, `#8`, `#13`, `#19` su jedan koherentan rez u `modDrive`/`modSelfUpdate`
(retry + paging + temp putanja + 0-bajt grana) i vrede više zajedno nego
pojedinačno.

---

## Obrazac zatvaranja — potvrđen četiri puta

Sva četiri zatvorena nalaza zatvorena su **uklanjanjem strukture, ne dodavanjem
provere** — i ni jedan onako kako je nalaz predlagao:

| Nalaz | Predlog iz nalaza | Kako je stvarno zatvoren |
|---|---|---|
| `#4` | „prebroj `expPals`, uporedi" | uklonjena dvofaznost — palete u transakciju pozivaoca, brojač nepotreban |
| `#21b` | „dodaj `IsStorniranoValue` filter" | obrisan potprogram — linija je aktivna iff je aktivna faktura, nema šta da se re-markira |
| `R1` | „čeka `ZbirnaID` na `tblPrijemnica`" | javni kanonski čitalac (ivica E4), bez nove kolone |
| `#6` simptom | „odluči: kolone odu ili dobiju pisca" | obrisani oba čitaoca; stanje u kanonu ne nastaje po konstrukciji |

Tri stvari koje iz toga slede za ostatak registra:

- **Minimalna zakrpa iz nalaza je hipoteza, ne recept.** Tamo gde nalaz predlaže
  proveru nad lošom strukturom, prvo pitanje je može li struktura da otpadne.
  Registar je `#21b` i `#4` sam označio kao „isti rod" — i oba su zatvorena tim
  putem, ne filterom.
- **Moje prognoze o tome *kako* će se nešto zatvoriti promašile su dvaput**
  (`#4` „~10 redova", `R1` „ne bez nove kolone"). Nalaz vredi kao opis kvara;
  procena rešenja u njemu je slabiji deo i tako je treba čitati.
- **Nalaz van putanje refaktora ne zatvara se sam.** Četiri zatvorena su sva bila
  u sloju koji se aktivno menjao. Sedam nalaza u `modDrive`/`modSelfUpdate` nije
  ni pogledano kroz 1041 commit, uključujući jedan koji je jedan red.

---

## Obim dokaza pri zatvaranju## Obim dokaza pri zatvaranju

Svaka stavka odavde koja menja ponašanje nosi test u `modTest` (CLAUDE.md §5), ne
checklistu. Za `#4`, `#20` i `#21c` traži se i **dokaz u oba smera** — to su
poslovne invarijante, a zelena suite koja nikad nije pokazana crvena ne dokazuje
da išta meri. U rezu se pušta samo filtriran `python tools/dokaz.py <prefiks>`,
pun katalog pred release.
