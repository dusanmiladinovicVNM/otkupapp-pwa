# Otvoreni nalazi: storno, self-update, Drive, migracija

Registar nalaza koji **još stoje** protiv `main @ 8e970d9` (20.10.2026). Nastao je
re-verifikacijom liste od 23 nalaza (revizija avgust 2026) posle refaktora
S3c-S5-3b, koji je identitet storna prebacio sa labele/broja na `ZbirnaID` i
kanonsko članstvo.

| Revizija | Baza | Promena |
|---|---|---|
| 26.09.2026 | `1095fdd` | prvi registar: 14 nalaza + 2 reziduala + 6 strukturnih |
| 20.10.2026 | `8e970d9` | +387 commit-a (nov model ambalaže, S6 prijemnica). **Nijedna stavka registra nije zatvorena.** `#17` i Strukturno ispravljeni nadole, `#9` i `#20` preformulisani, `#6` i `R1` izoštreni |

Čita se pre nego se otvori rez u `modStorno*`, `modSelfUpdate`, `modDrive` ili
`modMigracija` — da se ne krpi ono što je već zatvoreno i ne preskoči ono što
nije. Numeracija je **zadržana iz originalne liste**, pa `#4` ovde je `#4` tamo;
brojevi koji se ne pojavljuju su verifikovano zatvoreni i namerno se ne ponavljaju.

## Nivo verifikacije ovog dokumenta

| | |
|---|---|
| Metod | statičko čitanje izvora na `main @ 8e970d9` |
| FAST | `python tools/vba_check.py` → čisto (191 fajl) |
| Šema | `python tools/gen_schema_module.py --check` → u koraku (otisak `106AC27A`) |
| Vlasništvo | `python tools/who_writes.py --check-ownership` → nema novih pisaca (36 tabela) |
| TARGETED / FULL | **NIJE izvršeno** — `run_vba` traži Windows + Excel + `pywin32` |

Nijedna tvrdnja ovde **nije bihevioralno verifikovana**. Svaka je čitanje koda sa
navedenom putanjom i linijom. Gde je zaključak izveden a ne izmeren, tako i stoji.

## Gde ide pažnja, a gde ne

U 387 commit-a između dve revizije `modDrive.bas` i `modSelfUpdate.bas` **nisu
dirani ni jednom** (`git log 1095fdd..8e970d9 -- src-vba/modDrive.bas
src-vba/modSelfUpdate.bas` je prazan). Svih sedam nalaza u tom klasteru — `#5`,
`#7`, `#8`, `#13`, `#14`, `#19` i self-update polovina `#18` — zato stoje
**bajt-identično**, sa istim brojevima linija kao 26.09.

Storno i ambalaža su u istom periodu dobili ceo nov model knjiženja i S6, i tamo
se registar stvarno pomerio. Nije u pitanju da su nalazi iz Drive/self-update
klastera odbijeni — nisu bili ni pogledani. To je raspored prioriteta, i vredi ga
videti kao raspored a ne kao ocenu.

---

## P0 — propušta pogrešan podatak tiho

### #4 Faza B kaskade bez verifikacije

`modStornoFlow.bas:2116-2124` — petlja sumira `DetachOsirocenePaletaStavke_TX` u
`res("pals")`, pa `res("ok") = True` **bezuslovno**.

Callee guta svoju grešku: `modPaletniList.bas:2231-2234` na `EH` vraća `0`.
Neuspela prijemnica zato vrati nulu bez ijedne reči. Per-prijemnica TX je
atomičan, ali **petlja preko više prijemnica nije** — prijemnica 1 prođe,
prijemnica 2 padne, operater dobije pun uspeh i broj koji laže. `expPals` se
nigde ne računa.

**Minimalna zakrpa:** prebroj očekivane paletne stavke pre Faze B, uporedi sa
`res("pals")`, i na razliku ne postavljaj `ok` — otvori MANUAL kontekst.

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

### #6 Kaskada van žurnala — koren stoji, simptom neutralisan

**Koren stoji.** `BeginStornoOp` ima **samo dva živa pozivaoca**:
`modStorno.bas:169` (otkup). Zbirna, otpremnica i
prijemnica kaskade ne otvaraju žurnal operaciju, pa kaskada nije undo-abilna kao
celina.

**Simptom je otišao za kanon, ne za zatečeno.** `OtkupBlockDeadParentByID`
(`modStornoRecovery.bas:222`) čita `Otkup.OtpremnicaID` i `Otkup.BrojZbirne`. Te
kolone **više niko ne piše** — jedina referenca je `FreeOtkupBloksInline`
(`modStornoFlow.bas:1843`) koja ih *prazni*; row owner-i iz `WRITE_OWNERSHIP.json`
(`modDokumenta`, `modMasterSync`, `modOtkup`) ih nemaju. Za kanonske redove oba su
prazna → garda vraća `""` → undo prolazi.

Zatečeni redovi koji još nose stare linkove **i dalje se blokiraju**, sa porukom
koja krivi stanje koje je taj storno napravio.

**Napomena za sledeći rez šeme (izoštreno 20.10.).** `tblOtkup.BrojZbirne` i
`tblOtkup.OtpremnicaID` su kolone **bez pisca, ali sa živim čitaocima**:

| Čitalac | Šta radi |
|---|---|
| `modScrDokumenti.bas:1180` | čita `BrojZbirne` za prikaz operateru |
| `modScrIzvestaji.bas:2117`, `:2217` | izveštaj grupiše po njima |
| `modIntegritet.bas:1496-1497` | provera integriteta sudi po njima |
| `modStornoRecovery.bas:228-229` | undo garda (gore) |

Kolona koju niko ne piše, a četiri mesta čitaju, ne vraća „nema podatka" nego
**prazno kao da je podatak** — pa prikaz pokaže prazno polje, izveštaj grupiše u
prazan koš, a provera integriteta i garda prođu. Ovo je sada **veći deo nalaza od
same undo garde**: garda je jedan od četiri čitaoca.

Odluka je jedna od dve, i treba je doneti svesno: ili kolone odu kroz
`schema.json` pa čitaoci padnu glasno, ili dobiju pisca. Treće stanje — koje je
trenutno — je najgore od oba.

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

### #9 `LookupActiveID` na duplikat tiho uzima poslednji

Primitiva nepromenjena: `modStorno.bas:1588-1635` prolazi **sve** redove i pamti
**poslednji** aktivan pogodak, bez detekcije duplikata.

Glavni ulaz je zatvoren — storno tok za zbirnu ne ulazi kroz nju: `ZbrIdPoBroju`
(`modStornoFlow.bas:755`) diže grešku na `n <> 1`, a `PkPoIdentitetu`
(`modStornoFlow.bas:795`) izlazi prazan kad `VlasniciPoBroju(...).count > 1`,
**pre** poziva.

Nezaštićeni pozivaoci ostaju:

| Mesto | Tabela |
|---|---|
| `modDokUnos.bas:1007` | `tblPrijemnica` po `BrojZbirne` |
| `modStorno.bas:819`, `:821` | `tblNovac` |
| `modStornoDok.bas:409` | `tblFakture` |

**Uz to, netačan komentar na DVA mesta** (revizija 20.10. našla drugo):

- `modStornoDok.bas:406` — „LookupActiveID uzima prvi pogodak"
- `modStornoFlow.bas:649` — „Zamena za `LookupActiveID` ..., koji uzima PRVI aktivan"

Uzima **poslednji** (`modStorno.bas:1619-1626`: petlja prepisuje `resultId` do
kraja tabele, bez `Exit For`). Drugi komentar je gori od prvog: stoji nad
`PkPoIdentitetu`, dakle nad funkcijom koja postoji baš zato što `LookupActiveID`
nije dovoljno dobar — i pogrešno imenuje ono od čega brani. Oba se menjaju u istom
rezu, nulti rizik.

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

`modStornoRecovery.bas:33` — više prolaza kroz tabele bez `BeginTableCache`.

Frekvencija je popravljena (`Scr_Brojac` se zove iz `RefreshFromData`, ne pri
svakom crtanju sidebara — `modScrOporavak.bas:134-136`), pa je hitnost mala. Sam
prolaz je još neoptimizovan.

### #17 `tblStornoVeze` i `tblStornoZurnal` bez retencije

Dve append-only tabele, nula pogodaka za retenc / arhiv / prune.

**Ispravka 20.10.** Prva revizija je tvrdila da žurnal „raste najbrže od svega u
svesci". To **nije bilo tačno** ni tada, a sada je i manje: `JournalCell` ima dva
živa pisca — `modStorno.bas:170` (`tblOtkup.Stornirano`) i `modNovac.bas:1910`
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
pozivaocem**: `modStorno.bas:169`. Šalje netrimovano — `brDok =
NzToText(LookupValue(...))` na `:168`, a `NzToText` (`modHelpers.bas:249`) ne
trimuje.

Od dve posledice ostaje jedna i po:

- **lažno „mešanje operacija"** — traži dva ugnježdena poziva istog dokumenta sa
  različito razmaknutim brojem. Sa jednim pozivaocem i jednim `otkupID` po pozivu
  put je sada uzak; nije dokazano zatvoren, ali nije ni lako dosežan.
- **op sa praznim brojem** — stoji nepromenjeno: `broj` nije ni `Trim$`-ovan ni
  tražen kao non-blank, pa otkup red bez `BrojDokumenta` otvara operaciju sa
  praznim brojem, a `LatestOpFor(docType, broj)` je posle traži po tom istom
  praznom broju.

### #21 Dva od tri dela još stoje

**(a) Stale snapshot — ČISTO, ne dira se.** `StornoPrijemnica` hvata `fakturaID`
pre mutacije, a mutacija dira samo `Stornirano`; `StornoPaleta` čita `palData` pre
`MarkRowStornirano` i petlja po stavkama filtrira `Not IsStorniranoValue`.

**(b) `StornoFakturaStavkeAndReleasePrijemnice` bez `IsStorniranoValue` filtera —
STOJI.** `modStorno.bas:1672` poredi samo `colFakID`. Već stornirane stavke se
re-markiraju i `ReleasePrijemnicaFromFaktura` se vrti ponovo nad njima.

**(c) `CanStorno` guta schema greške — STOJI.** `modStorno.bas:1566-1586`:
`EH: ... CanStorno = False`. Schema drift, nečitljiva tabela i „nije dozvoljeno"
izlaze kao ista vrednost. Pozivalac nema način da ih razlikuje.

---

## Reziduali koje je re-verifikacija otkrila

Nisu iz originalne liste — izašli su iz nalaza koji su inače zatvoreni.

### R1 Prijemnice i palete se u kaskadi još biraju PO BROJU

Otpremnice idu po identitetu (`ActiveOtpIDsByZbirna` →
`ZbrClanoviPoStanju(zbirnaID)`), ali `ActivePrijIDsByZbirna`
(`modStornoFlow.bas:1799`) i `DistinctActiveValues` nad `COL_PALS_BROJ_ZBIRNE`
(`:1985`) biraju **po `BrojZbirne`**.

Nizvodni tok je zato zaštićen **kapijom** (`ZbirnaScopeRazlog`, fail-closed nad
dvosmislenim brojem), ne identitetom. Kapija je tačna, ali je slabija garancija:
odbija operaciju umesto da je suzi na svoje.

**Izoštreno 20.10.** S6 je u `tblPaletaStavka` doneo i `PrijemnicaID` i
`PrijemnicaStavkaID` — dakle kanonski identitet **stoji na redu**. Paletni motor ga
svejedno ne koristi za izbor: `DetachOsirocenePaletaStavke_TX`
(`modPaletniList.bas:2162`) prima `oldBroj` i bira po `COL_PALS_BROJ_PRIJ`
(`:2174`).

To nije regresija nego **modulski obrazac**: `ReassignPaleteToPrijemnica_TX`
(`:1518`) i `EvaluatePaletaReassign` (`:1998`) takođe primaju `oldBroj`, a
`PrijemnicaID` se u istom modulu piše (`:1378`, `:1842`) i čita (`:1694`, `:2077`).
Identitet je tu, motor ga nosi uz sebe i bira po labeli.

Posledica za `#4`: dok se Faza B hrani brojevima prijemnica, njena verifikacija ne
može biti stroža od toga čime bira. Ako se `#4` zatvara ozbiljno, prelaz na
`PrijemnicaID` je isti rez, ne sledeći.

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
| `modStornoFlow` podela na tri modula | **nije** — **2431 linija** (26.09. je bilo 2585, avgustovska revizija ga je merila na ~1800) |
| Generici u `modDataAccess` | **nema ih** — `CountActive` (`modStornoFlow.bas:2317`), `DistinctActiveValues` (`:2361`), `NzTx` (`:2416`) još su `Private` u `modStornoFlow` |
| `modMigracija` dry-run režim | **nema** — nula pogodaka za `dryRun` |
| `modMigracija` verzijska svesnost | **stub stoji** — `StaroImeKolone = novoIme` (`modMigracija.bas:471-472`); `APP_VERSION` starog fajla se ne čita |

**Ispravka 20.10.** Revizija od 26.09. je `modStornoFlow` nazvala „jedinom
metrikom u registru koja se pogoršala". Ta tvrdnja **više ne stoji**: fajl je od
tada spao sa 2585 na 2431 liniju (−154), jer je `10e-2` iz njega izneo grane
starog storna reversa. Nije podeljen i dalje je najveći modul u storno sloju, ali
trend je obrnut, i to treba reći isto onako jasno kao što je rečeno pogoršanje.

Nijedna metrika u registru se u ovoj reviziji nije pogoršala.

---

## Predložen redosled

**Izmenjen 20.10.** Prethodni redosled je `#4` i `#5` vezivao u jedno („manje od
sata"). Razvezuju se: `#5` je i dalje jedan red, a `#4` je potcenjen (v. niže).

1. **`#5`** — jedan red (`ThisWorkbook.Saved = True` pre `Exit Sub` na
   `modSelfUpdate.bas:230`). Najjeftinija stavka u registru, i jedina koja može da
   ostavi klijentu polomljen fajl. Skoro mesec dana stoji nedirnuta.
2. **`#9` oba komentara** (`modStornoDok.bas:406`, `modStornoFlow.bas:649`) —
   dva reda, nulti rizik. Ide uz bilo koji rez u tim fajlovima.
3. **`#21b` + `#21c`** — `IsStorniranoValue` filter i `CanStorno` koji guta schema
   grešku. Oba mala, oba tiha.
4. **`#4` zajedno sa `R1`** — prethodna revizija je `#4` zvala „~10 redova". To je
   bilo **potcenjeno**: dok Faza B bira po broju prijemnice, verifikacija ne može
   biti stroža od izbora. Kao pošten rez to je prelaz na `PrijemnicaID` plus
   brojanje, ne deset redova.
5. **`#6` odluka o šemi** — dve kolone bez pisca sa četiri čitaoca. Odluka, ne
   zakrpa.
6. **`#14`** — prijava ekstra komponente u `VerifyReleaseProject`. Jedini nalaz
   koji raste sa vremenom.

`#7`, `#8`, `#13`, `#19` su jedan koherentan rez u `modDrive`/`modSelfUpdate`
(retry + paging + temp putanja + 0-bajt grana) i vrede više zajedno nego
pojedinačno.

---

## Obim dokaza pri zatvaranju

Svaka stavka odavde koja menja ponašanje nosi test u `modTest` (CLAUDE.md §5), ne
checklistu. Za `#4`, `#20` i `#21c` traži se i **dokaz u oba smera** — to su
poslovne invarijante, a zelena suite koja nikad nije pokazana crvena ne dokazuje
da išta meri. U rezu se pušta samo filtriran `python tools/dokaz.py <prefiks>`,
pun katalog pred release.
