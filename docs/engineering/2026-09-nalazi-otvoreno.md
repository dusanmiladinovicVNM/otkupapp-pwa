# Otvoreni nalazi: storno, self-update, Drive, migracija

Registar nalaza koji **još stoje** protiv `main @ 1095fdd` (26.09.2026). Nastao je
re-verifikacijom liste od 23 nalaza (revizija avgust 2026) posle refaktora
S3c-S5-3b, koji je identitet storna prebacio sa labele/broja na `ZbirnaID` i
kanonsko članstvo.

Čita se pre nego se otvori rez u `modStorno*`, `modSelfUpdate`, `modDrive` ili
`modMigracija` — da se ne krpi ono što je već zatvoreno i ne preskoči ono što
nije. Numeracija je **zadržana iz originalne liste**, pa `#4` ovde je `#4` tamo;
brojevi koji se ne pojavljuju su verifikovano zatvoreni i namerno se ne ponavljaju.

## Nivo verifikacije ovog dokumenta

| | |
|---|---|
| Metod | statičko čitanje izvora na `main @ 1095fdd` |
| FAST | `python tools/vba_check.py` → čisto (189 fajlova, 607 sabotaža, 0 nalaza +2 poznatih) |
| Šema | `python tools/gen_schema_module.py --check` → u koraku (otisak `88E04EC5`) |
| TARGETED / FULL | **NIJE izvršeno** — `run_vba` traži Windows + Excel + `pywin32` |

Nijedna tvrdnja ovde **nije bihevioralno verifikovana**. Svaka je čitanje koda sa
navedenom putanjom i linijom. Gde je zaključak izveden a ne izmeren, tako i stoji.

---

## P0 — propušta pogrešan podatak tiho

### #4 Faza B kaskade bez verifikacije

`modStornoFlow.bas:2204-2212` — petlja sumira `DetachOsirocenePaletaStavke_TX` u
`res("pals")`, pa `res("ok") = True` **bezuslovno**.

Callee guta svoju grešku: `modPaletniList.bas:2171-2174` na `EH` vraća `0`.
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
`modStorno.bas:165` (otkup) i `modStorno.bas:1920` (revers). Zbirna, otpremnica i
prijemnica kaskade ne otvaraju žurnal operaciju, pa kaskada nije undo-abilna kao
celina.

**Simptom je otišao za kanon, ne za zatečeno.** `OtkupBlockDeadParentByID`
(`modStornoRecovery.bas:372`) čita `Otkup.OtpremnicaID` i `Otkup.BrojZbirne`. Te
kolone **više niko ne piše** — jedina referenca je `FreeOtkupBloksInline`
(`modStornoFlow.bas:2021`) koja ih *prazni*; row owner-i iz `WRITE_OWNERSHIP.json`
(`modDokumenta`, `modMasterSync`, `modOtkup`) ih nemaju. Za kanonske redove oba su
prazna → garda vraća `""` → undo prolazi.

Zatečeni redovi koji još nose stare linkove **i dalje se blokiraju**, sa porukom
koja krivi stanje koje je taj storno napravio.

**Napomena za sledeći rez šeme:** `tblOtkup.BrojZbirne` i `tblOtkup.OtpremnicaID`
su sada kolone bez pisca. Ili ih očisti kroz `schema.json`, ili im ostavi čitaoca
svesno — trenutno su tiha zamka baš za `OtkupBlockDeadParentByID`.

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

Primitiva nepromenjena: `modStorno.bas:1549-1596` prolazi **sve** redove i pamti
**poslednji** aktivan pogodak, bez detekcije duplikata.

Glavni ulaz je zatvoren — storno tok za zbirnu ne ulazi kroz nju: `ZbrIdPoBroju`
(`modStornoFlow.bas:956`) diže grešku na `n <> 1`, a `PkPoIdentitetu`
(`modStornoFlow.bas:996`) izlazi prazan kad `VlasniciPoBroju(...).count > 1`,
**pre** poziva.

Nezaštićeni pozivaoci ostaju:

| Mesto | Tabela |
|---|---|
| `modDokUnos.bas:998` | `tblPrijemnica` po `BrojZbirne` |
| `modStorno.bas:780`, `:782` | `tblNovac` |
| `modStornoDok.bas:377` | `tblFakture` |

**Uz to, netačan komentar:** `modStornoDok.bas:374` kaže „LookupActiveID uzima
prvi pogodak". Uzima **poslednji**. Komentar aktivno obmanjuje sledećeg čitaoca i
menja se u istom rezu — nulti rizik.

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

`modStornoRecovery.bas:37` — više prolaza kroz tabele bez `BeginTableCache`.

Frekvencija je popravljena (`Scr_Brojac` se zove iz `RefreshFromData`, ne pri
svakom crtanju sidebara — `modScrOporavak.bas:134-136`), pa je hitnost mala. Sam
prolaz je još neoptimizovan.

### #17 `tblStornoVeze` i `tblStornoZurnal` bez retencije

Dve append-only tabele, nula pogodaka za retenc / arhiv / prune. Žurnal je
cell-level, pa raste najbrže od svega u svesci.

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

`modStornoZurnal.bas:41` zahteva non-blank **samo** `docType`; `broj` nikad.
Uporedba je golo `StrComp(broj, mBroj, vbTextCompare)` bez `Trim$`, i `mBroj =
broj` čuva netrimovano.

Pozivaoci šalju netrimovano:

- `modStorno.bas:164` — `brDok = NzToText(LookupValue(...))`, a `NzToText`
  (`modHelpers.bas:249`) **ne trimuje**
- `modStorno.bas:1893` — `RequireNonBlank brDok` takođe ne trimuje

Posledice: lažno „mešanje operacija" nad istim dokumentom, i op sa **praznim
brojem** za unbound blok.

### #21 Dva od tri dela još stoje

**(a) Stale snapshot — ČISTO, ne dira se.** `StornoPrijemnica` hvata `fakturaID`
pre mutacije, a mutacija dira samo `Stornirano`; `StornoPaleta` čita `palData` pre
`MarkRowStornirano` i petlja po stavkama filtrira `Not IsStorniranoValue`.

**(b) `StornoFakturaStavkeAndReleasePrijemnice` bez `IsStorniranoValue` filtera —
STOJI.** `modStorno.bas:1633` poredi samo `colFakID`. Već stornirane stavke se
re-markiraju i `ReleasePrijemnicaFromFaktura` se vrti ponovo nad njima.

**(c) `CanStorno` guta schema greške — STOJI.** `modStorno.bas:1527-1547`:
`EH: ... CanStorno = False`. Schema drift, nečitljiva tabela i „nije dozvoljeno"
izlaze kao ista vrednost. Pozivalac nema način da ih razlikuje.

---

## Reziduali koje je re-verifikacija otkrila

Nisu iz originalne liste — izašli su iz nalaza koji su inače zatvoreni.

### R1 Prijemnice i palete se u kaskadi još biraju PO BROJU

Otpremnice idu po identitetu (`ActiveOtpIDsByZbirna` →
`ZbrClanoviPoStanju(zbirnaID)`), ali `ActivePrijIDsByZbirna`
(`modStornoFlow.bas:1977`) i `DistinctActiveValues` nad `COL_PALS_BROJ_ZBIRNE`
(`:2148`) biraju **po `BrojZbirne`**.

Nizvodni tok je zato zaštićen **kapijom** (`ZbirnaScopeRazlog`, fail-closed nad
dvosmislenim brojem), ne identitetom. Kapija je tačna, ali je slabija garancija:
odbija operaciju umesto da je suzi na svoje.

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
| `modStornoFlow` podela na tri modula | **nije** — i narastao na **2585 linija** (revizija ga je merila na ~1800), *posle* brisanja ~800 redova okvira modova. Neto rast ~1600. |
| Generici u `modDataAccess` | **nema ih** — `CountActive` (`modStornoFlow.bas:2471`), `DistinctActiveValues` (`:2515`), `NzTx` (`:2570`) još su `Private` u `modStornoFlow` |
| `modMigracija` dry-run režim | **nema** — nula pogodaka za `dryRun` |
| `modMigracija` verzijska svesnost | **stub stoji** — `StaroImeKolone = novoIme` (`modMigracija.bas:471-472`); `APP_VERSION` starog fajla se ne čita |

`modStornoFlow` na 2585 linija je **jedina metrika u registru koja se pogoršala.**

---

## Predložen redosled

1. **`#4` + `#5`** — ~10 redova i 1 red. Zajedno zatvaraju sve što je ostalo od
   „pogrešan podatak izlazi tiho" u P0. Manje od sata.
2. **`#21b` + `#21c`** — `IsStorniranoValue` filter i `CanStorno` koji guta schema
   grešku. Oba mala, oba tiha.
3. **`#9` komentar na `modStornoDok.bas:374`** — nulti rizik, sprečava pogrešan
   zaključak sledećeg čitaoca. Ide uz bilo koji rez u tom fajlu.
4. **`#14`** — prijava ekstra komponente u `VerifyReleaseProject`. Jedini nalaz
   koji raste sa vremenom.
5. **`#6` napomena o šemi** — `tblOtkup.BrojZbirne` / `.OtpremnicaID` bez pisca,
   uz prvi sledeći rez šeme.

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
