# Otvoreni nalazi: storno, self-update, Drive, migracija

Registar nalaza koji **još stoje** protiv `main @ b552892` (29.10.2026). Nastao je
re-verifikacijom liste od 23 nalaza (revizija avgust 2026) posle refaktora
S3c-S5-3b, koji je identitet storna prebacio sa labele/broja na `ZbirnaID` i
kanonsko članstvo.

| Revizija | Baza | Promena |
|---|---|---|
| 26.09.2026 | `1095fdd` | prvi registar: 14 nalaza + 2 reziduala + 6 strukturnih |
| 20.10.2026 | `8e970d9` | +387 commit-a (nov model ambalaže, S6 prijemnica). **Nijedna stavka registra nije zatvorena.** `#17` i Strukturno ispravljeni nadole, `#9` i `#20` preformulisani, `#6` i `R1` izoštreni |
| 29.10.2026 | `b552892` | +96 commit-a (S6 koraci 4–8). **`#4` zatvoren** — prvi zatvoren nalaz u registru. `R1` pola zatvoreno, `#9` i `#21b` suženi, Strukturno pogoršano |

Čita se pre nego se otvori rez u `modStorno*`, `modSelfUpdate`, `modDrive` ili
`modMigracija` — da se ne krpi ono što je već zatvoreno i ne preskoči ono što
nije. Numeracija je **zadržana iz originalne liste**, pa `#4` ovde je `#4` tamo;
brojevi koji se ne pojavljuju su verifikovano zatvoreni i namerno se ne ponavljaju.

## Nivo verifikacije ovog dokumenta

| | |
|---|---|
| Metod | statičko čitanje izvora na `main @ b552892` |
| FAST | `python tools/vba_check.py` → čisto (191 fajl) |
| Šema | `python tools/gen_schema_module.py --check` → u koraku (otisak `87AB4F9B`) |
| Vlasništvo | `python tools/who_writes.py --check-ownership` → nema novih pisaca (36 tabela) |
| TARGETED / FULL | **NIJE izvršeno** — `run_vba` traži Windows + Excel + `pywin32` |

Nijedna tvrdnja ovde **nije bihevioralno verifikovana**. Svaka je čitanje koda sa
navedenom putanjom i linijom. Gde je zaključak izveden a ne izmeren, tako i stoji.

## Gde ide pažnja, a gde ne

`modDrive.bas` i `modSelfUpdate.bas` **nisu dirani ni jednom u 483 commit-a**,
kroz oba perioda:

| Period | Commit-a | `git log <opseg> -- modDrive.bas modSelfUpdate.bas` |
|---|---|---|
| `1095fdd..8e970d9` | 387 | prazan |
| `8e970d9..b552892` | 96 | prazan |

Svih sedam nalaza u tom klasteru — `#5`, `#7`, `#8`, `#13`, `#14`, `#19` i
self-update polovina `#18` — zato stoje **bajt-identično**, sa istim brojevima
linija kao 26.09. `#5` je jedan red.

Storno, ambalaža i prijemnica su u istom periodu dobili nov model knjiženja i S6
do koraka 8, i tamo se registar stvarno pomerio — uključujući prvi zatvoren
nalaz. Nije u pitanju da su nalazi iz Drive/self-update klastera odbijeni; nisu
bili ni pogledani. To je raspored prioriteta, i vredi ga videti kao raspored a ne
kao ocenu.

---

## P0 — propušta pogrešan podatak tiho

### #4 Faza B kaskade bez verifikacije — ZATVOREN 29.10.

Jedini zatvoren nalaz u registru. Ostaje zapisan jer pokazuje kako se ova klasa
zatvara: **uklanjanjem dvofaznosti**, ne dodavanjem brojača.

Faza B više ne postoji — `grep "Faza B: paletne stavke"` nad `modStornoFlow.bas`
ne vraća nijedan pogodak, a `DetachOsirocenePaletaStavke_TX` se iz kaskade više ne
zove. Palete se sada skidaju **unutar transakcije pozivaoca**:

`PonistiZbirnaChain_TX` (`modStornoFlow.bas:2170-2201`) zove `StornoPrijemnica(id,
tx, True, palK)` i sabira `palUkupno`, pa `res("pals") = palUkupno` i
`res("ok") = True` stoje **posle** `tx.CommitTx`. `StornoPrijemnica`
(`modStorno.bas:484`) prima `ByRef outPaletnihStavki` i delegira
`modPaletniList.SkiniPaleteZaStornoPrijemnice(tx, prijemnicaID, SRC)`
(`modPaletniList.bas:2111`).

Ta primitiva zatvara sva tri otvora iz nalaza:

| Otvor iz nalaza | Kako je zatvoren |
|---|---|
| callee guta grešku (`EH: ... = 0`) | **nema `On Error` ni `EH:` granu** — svaki put diže grešku |
| petlja preko prijemnica nije atomična | traži `tx` pozivaoca i diže grešku bez njega — jedna transakcija za ceo lanac |
| `expPals` se ne računa | broj dolazi **iz same mutacije** (`outPaletnihStavki`), ne iz nezavisnog prebrojavanja |

Uz to je prešla na identitet: bira po `COL_PALS_PRIJEMNICA_ID`
(`modPaletniList.bas:2132`), ne po labeli — čime je zatvorena i paletna polovina
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

### #6 Kaskada van žurnala — koren stoji, simptom neutralisan

**Koren stoji.** `BeginStornoOp` ima **tačno jednog živog pozivaoca**:
`modStorno.bas:150` (otkup). Zbirna, otpremnica i prijemnica kaskade ne otvaraju
žurnal operaciju, pa kaskada nije undo-abilna kao celina.

> Revizija 20.10. je ovde rekla „samo dva živa pozivaoca" a navela jedan —
> nedoslednost nastala kad je revers izašao iz žurnala a tekst nije dopunjen do
> kraja. Prebrojano 29.10.: jedan.

**Simptom je otišao za kanon, ne za zatečeno.** `OtkupBlockDeadParentByID`
(`modStornoRecovery.bas:269`) čita `Otkup.OtpremnicaID` i `Otkup.BrojZbirne`. Te
kolone **više niko ne piše** — jedina referenca je `FreeOtkupBloksInline`
(`modStornoFlow.bas:1907`, prazni na `:1924`) koja ih *prazni*; row owner-i iz `WRITE_OWNERSHIP.json`
(`modDokumenta`, `modMasterSync`, `modOtkup`) ih nemaju. Za kanonske redove oba su
prazna → garda vraća `""` → undo prolazi.

Zatečeni redovi koji još nose stare linkove **i dalje se blokiraju**, sa porukom
koja krivi stanje koje je taj storno napravio.

**Napomena za sledeći rez šeme (izoštreno 20.10.).** `tblOtkup.BrojZbirne` i
`tblOtkup.OtpremnicaID` su kolone **bez pisca, ali sa živim čitaocima**:

| Čitalac | Šta radi |
|---|---|
| `modIntegritet.bas:564`, `:1496-1497` | provera integriteta sudi po njima |
| `modScrIzvestaji.bas:2143`, `:2243` | izveštaj grupiše po njima |
| `modScrDokumenti.bas:2531` | mapira kolonu za ekran |
| `modStornoDok.bas:1059` | mapira kolonu za prefill |
| `modStornoFlow.bas:2264` | bira blokove po `OtpremnicaID` |
| `modStornoRecovery.bas:275-276` | undo garda (gore) |

Kolona koju niko ne piše, a šest mesta čita, ne vraća „nema podatka" nego
**prazno kao da je podatak** — pa prikaz pokaže prazno polje, izveštaj grupiše u
prazan koš, a provera integriteta i garda prođu. Ovo je sada **veći deo nalaza od
same undo garde**: garda je jedan od šest čitaoca. Direktno čitanje za prikaz, koje je
revizija 20.10. navela, u međuvremenu je otišlo — ali su ostala mapiranja kolone,
pa ekran i prefill i dalje stižu do nje.

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

Primitiva nepromenjena: `modStorno.bas:1583-1630` prolazi **sve** redove i pamti
**poslednji** aktivan pogodak, bez detekcije duplikata.

Glavni ulaz je zatvoren — storno tok za zbirnu ne ulazi kroz nju: `ZbrIdPoBroju`
(`modStornoFlow.bas:763`) diže grešku na `n <> 1`, a `PkPoIdentitetu`
(`modStornoFlow.bas:795`) izlazi prazan kad `VlasniciPoBroju(...).count > 1`,
**pre** poziva.

Nezaštićeni pozivaoci ostaju:

| Mesto | Tabela |
|---|---|
| `modDokUnos.bas:1011` | `tblPrijemnica` po `BrojZbirne` |
| `modStorno.bas:814`, `:816` | `tblNovac` |
| `modStornoDok.bas:416` | `tblFakture` |

**Uz to, netačan komentar — opet na JEDNOM mestu.** `modStornoDok.bas:413` kaže
„LookupActiveID uzima prvi pogodak". Uzima **poslednji**: petlja prepisuje
`resultId` do kraja tabele, bez `Exit For`.

**Suženo 29.10.** Drugi netačan komentar, koji je revizija 20.10. našla na
`modStornoFlow.bas:649`, **otišao je sa `PkPoIdentitetu`** — ta funkcija je
zamenjena (`modStornoFlow.bas:772` opisuje zašto: „identitet je čitala kao
GENERACIJU"). Nije ispravljen nego uklonjen zajedno sa kodom nad kojim je stajao.
Ostaje jedan red, nulti rizik.

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

`modStornoRecovery.bas:34` — više prolaza kroz tabele bez `BeginTableCache`.

Frekvencija je popravljena (`Scr_Brojac` se zove iz `RefreshFromData`, ne pri
svakom crtanju sidebara — `modScrOporavak.bas:139`), pa je hitnost mala. Sam
prolaz je još neoptimizovan.

### #17 `tblStornoVeze` i `tblStornoZurnal` bez retencije

Dve append-only tabele, nula pogodaka za retenc / arhiv / prune.

**Ispravka 20.10.** Prva revizija je tvrdila da žurnal „raste najbrže od svega u
svesci". To **nije bilo tačno** ni tada, a sada je i manje: `JournalCell` ima dva
živa pisca — `modStorno.bas:151` (`tblOtkup.Stornirano`) i `modNovac.bas:1910`
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
pozivaocem**: `modStorno.bas:150`. Šalje netrimovano — `brDok =
NzToText(LookupValue(...))` na `:149`, a `NzToText` (`modHelpers.bas:249`) ne
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

**(b) Filter `IsStorniranoValue` još nema — STOJI, suženo 29.10.** Potprogram je
preimenovan u `StornoFakturaStavke` (`modStorno.bas:1641`) i `:1669` i dalje poredi
samo `colFakID`, bez provere da je red već storniran. Već stornirane stavke se
zato re-markiraju.

Posledica je prepolovljena: `ReleasePrijemnicaFromFaktura` je **izašao iz ovog
potprograma** — S6 korak 5 je fakturu prebacio na `PrijemnicaStavkaID` a
fakturisanost učinio izvedenom, pa nema markera koji bi se dvaput resetovao. Ostaje
samo utovar grana: `ReleaseUtovarFromFaktura` se može ponovo pozvati nad već
storniranom stavkom. `relUt` dikt štiti od duplog poziva **unutar jednog prolaza**,
ne od ponovljenog storna iste fakture.

**(c) `CanStorno` guta schema greške — STOJI.** `modStorno.bas:1561-1581`:
`EH: ... CanStorno = False`. Schema drift, nečitljiva tabela i „nije dozvoljeno"
izlaze kao ista vrednost. Pozivalac nema način da ih razlikuje.

---

## Reziduali koje je re-verifikacija otkrila

Nisu iz originalne liste — izašli su iz nalaza koji su inače zatvoreni.

### R1 Prijemnice i palete se u kaskadi još biraju PO BROJU

Otpremnice idu po identitetu (`ActiveOtpIDsByZbirna` →
`ZbrClanoviPoStanju(zbirnaID)`), ali `ActivePrijIDsByZbirna`
(`modStornoFlow.bas:1880`) i `DistinctActiveValues` nad `COL_PALS_BROJ_ZBIRNE`
(`:1985`) biraju **po `BrojZbirne`**.

Nizvodni tok je zato zaštićen **kapijom** (`ZbirnaScopeRazlog`, fail-closed nad
dvosmislenim brojem), ne identitetom. Kapija je tačna, ali je slabija garancija:
odbija operaciju umesto da je suzi na svoje.

**Paletna polovina ZATVORENA 29.10.** Revizija 20.10. je tvrdila da paletni motor
nosi `PrijemnicaID` uz sebe a bira po labeli, i da je prelaz na identitet isti rez
kao `#4`. **Tako je i ispalo**: `SkiniPaleteZaStornoPrijemnice`
(`modPaletniList.bas:2111`) bira po `COL_PALS_PRIJEMNICA_ID` (`:2132`), i stigla je
zajedno sa zatvaranjem `#4`.

**Prijemnička polovina stoji.** `ActivePrijIDsByZbirna` (`modStornoFlow.bas:1880`)
i dalje prima `brojZbirne` i bira po labeli; isto `CountActive` nad
`COL_PALS_BROJ_ZBIRNE` (`:2371`). Zbirna → prijemnica veza je i dalje po broju, pa
nizvodni tok ostaje zaštićen **kapijom** (`ZbirnaScopeRazlog`), ne identitetom.

Razlika prema paletnoj polovini je u šemi: `tblPaletaStavka` nosi `PrijemnicaID`,
a `tblPrijemnica` **ne nosi `ZbirnaID`** — veza je `BrojZbirne`. Dok se to ne
promeni, ovo se ne može zatvoriti istim potezom; kapija je ispravan izbor do tada.

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
| `modStornoFlow` podela na tri modula | **nije** — **2513 linija** (20.10. bilo 2431, 26.09. bilo 2585; avgustovska revizija merila ~1800) |
| Generici u `modDataAccess` | **nema ih, i ima ih više** — `CountActive` (`modStornoFlow.bas:2399`), `DistinctActiveValues` (`:2443`), `NzTx` (`:2498`) su `Private` u `modStornoFlow`, a S6 je dodao i `FindRowIndexByID` kao `Private` u `modPaletniList` (`:1386`) — četvrta privatna kopija generika |
| `modMigracija` dry-run režim | **nema** — nula pogodaka za `dryRun` |
| `modMigracija` verzijska svesnost | **stub stoji** — `StaroImeKolone = novoIme` (`modMigracija.bas:471-472`); `APP_VERSION` starog fajla se ne čita |

**Kretanje `modStornoFlow` kroz tri revizije:** 2585 (26.09.) → 2431 (20.10.) →
**2513 (29.10.)**. Pad od −154 koji je revizija 20.10. prijavila bio je tačan, ali
privremen: S6 koraci 4–8 su vratili +82. Neto kroz oba perioda je −72, dakle fajl
se drži oko 2500 i **nije podeljen**.

Ovo je jedina metrika u registru koja se u ovoj reviziji pogoršala, i treba je
čitati zajedno sa prethodnim redom: kad S6 doda sposobnost, ona padne u
`modStornoFlow`, jer tri modula na koja bi se podelio ne postoje. Generici idu u
istom smeru — `FindRowIndexByID` je četvrta privatna kopija, ne prva javna.

---

## Predložen redosled

**Izmenjen 29.10.** `#4` je zatvoren, pa izlazi iz redosleda. `#9` je spao na
jedan red. Ostatak je nepromenjen jer se nad njim nije radilo.

1. **`#5`** — jedan red (`ThisWorkbook.Saved = True` pre `Exit Sub` na
   `modSelfUpdate.bas:230`). Najjeftinija stavka u registru, i jedina koja može da
   ostavi klijentu polomljen fajl. **Pet nedelja i 483 commit-a** stoji nedirnuta,
   u modulu koji za to vreme nije otvoren ni jednom.
2. **`#9` komentar** (`modStornoDok.bas:413`) — jedan red, nulti rizik.
3. **`#21b` + `#21c`** — `IsStorniranoValue` filter i `CanStorno` koji guta schema
   grešku. Oba mala, oba tiha.
4. **`#6` odluka o šemi** — dve kolone bez pisca sa šest čitaoca. Odluka, ne
   zakrpa.
5. **`#14`** — prijava ekstra komponente u `VerifyReleaseProject`. Jedini nalaz
   koji raste sa vremenom.
6. **`R1` prijemnička polovina** — čeka `ZbirnaID` na `tblPrijemnica`; ne zatvara
   se bez tog reza u šemi.

`#7`, `#8`, `#13`, `#19` su jedan koherentan rez u `modDrive`/`modSelfUpdate`
(retry + paging + temp putanja + 0-bajt grana) i vrede više zajedno nego
pojedinačno.

---

## Šta je naučeno zatvaranjem `#4`

Prvi zatvoren nalaz vredi pročitati kao obrazac, jer je zatvoren **drugačije od
onoga što je nalaz predlagao**.

Nalaz je rekao „prebroj `expPals`, uporedi" — dakle dodaj proveru nad dvofaznim
tokom. Rez je umesto toga **uklonio dvofaznost**: palete su ušle u transakciju
pozivaoca, callee je ostao bez `On Error` grane, a broj je počeo da dolazi iz same
mutacije. Posle toga brojač koji je nalaz tražio nije imao šta da proverava.

Dve stvari koje iz toga slede za ostatak registra:

- **Minimalna zakrpa iz nalaza nije uvek najbolji rez.** Tamo gde nalaz predlaže
  proveru nad lošom strukturom, pitanje je prvo može li struktura da otpadne.
  `#21b` je isti rod: filter nad petljom koja re-markira, ili petlja koja ne može
  da re-markira.
- **Identitet i verifikacija dolaze zajedno.** `#4` se nije mogao zatvoriti
  ozbiljno dok je Faza B birala po labeli — što je revizija 20.10. i predvidela.
  Isto važi za `R1` prijemničku polovinu i `#9`: dok se bira po broju, provera ne
  može biti stroža od izbora.

---

## Obim dokaza pri zatvaranju

Svaka stavka odavde koja menja ponašanje nosi test u `modTest` (CLAUDE.md §5), ne
checklistu. Za `#4`, `#20` i `#21c` traži se i **dokaz u oba smera** — to su
poslovne invarijante, a zelena suite koja nikad nije pokazana crvena ne dokazuje
da išta meri. U rezu se pušta samo filtriran `python tools/dokaz.py <prefiks>`,
pun katalog pred release.
