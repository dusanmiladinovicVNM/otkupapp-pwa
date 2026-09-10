# Refaktor: dokument = header + stavke

> Status: **PR0–PR2 mergovani, PR3 u reviziji; od PR4 nadalje je plan.** Tačno
> stanje po stavkama: §14 „PR-ovi".
>
> Kontekst: **nema migracije i nema legacy podataka** — sezona je prošla,
> nijedan klijent nije na starom programu, novi korisnici kreću sa novom šemom.
> Zato se stari model **briše**, ne prevodi.
>
> Ovaj fajl je i model (šta dokumenti postaju) i plan (kojim redom). Kad se
> implementira, model deo prelazi u `docs/DOMEN/DOCUMENT_HEADER_LINES.md`, a ovaj
> fajl ostaje kao zapis odluka.

---

## 1) Problem, u jednoj rečenici

`jedan fizički red = jedan dokument` nije tačno: dvoklasni dokument su dva reda sa
dva ID-a, pa se logički dokument posle **rekonstruiše** — kroz `GeneracijaID`,
poslovni broj, vlasnički scope i resolvere. Sve to je kompenzacija za nepostojeće
zaglavlje.

Najjasniji dokaz da logički dokument postoji a nema gde da živi: `Save*Multi_TX`
vraća string `"OTK-1 + OTK-2"`, koji se onda **parsira** na devet mesta
(`Split(x, " + ")` u `modAmbalaza`, `modAutoHladnjaca`, `modDokUnos` ×3,
`modOtkupBlok`, `modPrint` ×3). Konkatenirani string je de facto identifikator
dokumenta.

---

## 2) Pravilo podele header/stavka — mehaničko, ne procena

> **Polje koje u postojećem `Save*Multi_TX` potpisu stoji JEDNOM ide na header.
> Polje koje stoji kao par I/II ide na stavku.**

Potpis je već napisana specifikacija: autor je pre svake klase odlučio da li
vrednost varira. Time podela nije nova poslovna odluka — samo se zapisuje u šemu.

Jedan izuzetak, izveden iz **ponašanja** a ne iz potpisa: `Fakturisano` /
`FakturaID` na prijemnici. `GetPrijemniceByKupac(samoNefakturisano)`
(`modDokumenta.bas:2288`) filtrira po redu, dakle po klasi — korisnik danas može
da fakturiše Klasu I bez Klase II. Da ta polja odu na header, tiho bi se izgubila
delimična fakturisanost. **Stavka.**

### Šta stavka nosi

| Dokument | Kolone stavke | Napomena |
|---|---|---|
| Otkup | `RedniBroj`, `Klasa`, `Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg` | |
| Otpremnica | `RedniBroj`, `Klasa`, `Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg` | |
| Zbirna | `RedniBroj`, `Klasa`, `Kolicina`, `KolAmbalaze` | zbirna **nema** cenu (`modDokUnos.bas:422`) |
| Prijemnica | `RedniBroj`, `Klasa`, `Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg`, `Fakturisano`, `FakturaID` | |

### Šta ostaje na headeru iako „deluje kao stavka"

`VrstaVoca`, `SortaVoca`, `TipAmbalaze`, `KulturaID`, `ParcelaID` — svi stoje
**jednom** u potpisu, dakle dokument-level. Posledica koja se mora poštovati:

> **Dokument ima jednu vrstu, jednu sortu i jedan tip ambalaže. Stavke se
> razlikuju samo po klasi.** Mešovit dokument nije novo dozvoljeno stanje.

Ovo nije ograničenje koje refaktor uvodi — to je ograničenje koje refaktor
**čuva**. Da su vrsta/sorta otišle na stavku, model bi strukturno dozvolio
mešovitu zbirnu koju invarijanta, štampa i faktura nikad nisu videle.

### `Isplaceno` / `DatumIsplate` → header (i to je ispravka, ne izmena)

Danas: `MarkOtkupIsplacen` (`modNovac.bas:1240`) računa `Kolicina × Cena`
**jednog reda** i poredi sa uplatama vezanim za taj `OtkupID`. Ali gotovina se
upisuje samo na primarni red — `SaveNovac(otkupID:=primaryID)`
(`modOtkup.bas:279`), gde je `primaryID` Klasa I ako postoji. Rezultat: kod
dvoklasnog otkupa red Klase II **nikad ne dobija `Isplaceno`**, iako je kooperant
plaćen u celosti.

Posle refaktora: vrednost dokumenta = `SUM(stavke.Kolicina × stavke.Cena)`,
`tblNovac.OtkupID` pokazuje na header, `Isplaceno` je jedno polje na headeru.
Primary-row hack nestaje. Ovo je jedini slučaj gde refaktor menja zatečeno
ponašanje — i menja ga zato što je zatečeno ponašanje pogrešno.

---

## 3) Ciljna šema

Prefiksi po postojećoj konvenciji (`PLS-` za `tblPaletaStavka`): `OKS-`, `OPS-`,
`ZBS-`, `PRS-`.

### `tblOtkup` (header)

```
OtkupID           PK, "OTK-"
Datum
KooperantID       FK
StanicaID         FK
KulturaID
VrstaVoca
SortaVoca
ParcelaID         FK
TipAmbalaze
KolAmbIzdata      dokument-level (OM izdao prazne kooperantu)
VozacID           FK
BrojDokumenta     poslovni broj -- LABELA
Novac             snapshot isplacene gotovine
PrimalacNovca
OtpremnicaID      FK, nullable   <- ODLUKA, v. 3.1
ZbirnaID          FK, nullable   <- zamenjuje BrojZbirne (denorm)
Isplaceno
DatumIsplate
VremeUnosa
Stornirano
IspravkaOdID / ZamenjenSaID / CorrectionID / IzdatoStatus
CreatedAt / CreatedBy / ModifiedAt / ModifiedBy
```

### `tblOtkupStavke`

```
OtkupStavkaID     PK, "OKS-"
OtkupID           FK, obavezan
RedniBroj
Klasa
Kolicina
Cena
KolAmbalaze
BrutoKg
```

### `tblOtpremnica` (header)

```
OtpremnicaID  PK "OTP-" | Datum | StanicaID | VozacID | BrojOtpremnice
VrstaVoca | SortaVoca | TipAmbalaze
ZbirnaID      FK, nullable   <- NOVO, zamenjuje BrojZbirne
Stornirano | trace | audit
```

### `tblOtpremnicaStavke`

```
OtpremnicaStavkaID PK "OPS-" | OtpremnicaID FK | RedniBroj
Klasa | Kolicina | Cena | KolAmbalaze | BrutoKg
```

### `tblZbirna` (header)

```
ZbirnaID PK "ZBR-" | Datum | VozacID | BrojZbirne | KupacID
Hladnjaca | Pogon | VrstaVoca | SortaVoca | TipAmbalaze
Stornirano | trace | audit
```

### `tblZbirnaStavke`

```
ZbirnaStavkaID PK "ZBS-" | ZbirnaID FK | RedniBroj
Klasa | Kolicina | KolAmbalaze
```

### `tblPrijemnica` (header)

```
PrijemnicaID PK "PRJ-" | Datum | KupacID | VozacID | BrojPrijemnice
VrstaVoca | SortaVoca | TipAmbalaze | KolAmbVracena
ZbirnaID     FK, nullable   <- NOVO, zamenjuje BrojZbirne
Stornirano | trace | audit
```

### `tblPrijemnicaStavke`

```
PrijemnicaStavkaID PK "PRS-" | PrijemnicaID FK | RedniBroj
Klasa | Kolicina | Cena | KolAmbalaze | BrutoKg
Fakturisano | FakturaID
```

### `tblFakturaStavke` — jedina izmena

```
PrijemnicaStavkaID   FK   <- NOVO, kanonski izvor
PrijemnicaID              <- ostaje denormalizovan (read-modeli, poruke)
```

### PK/FK graf

```
tblOtkup ──1:N──> tblOtkupStavke
   │ OtpremnicaID (nullable)
   │ ZbirnaID (nullable, denorm)
   v
tblOtpremnica ──1:N──> tblOtpremnicaStavke
   │ ZbirnaID (nullable)
   v
tblZbirna ──1:N──> tblZbirnaStavke
   ^
   │ ZbirnaID
tblPrijemnica ──1:N──> tblPrijemnicaStavke
                            ^
                            │ PrijemnicaStavkaID
                       tblFakturaStavke ──N:1──> tblFakture

tblAmbalaza.DokumentID  -> header ID (Otkup / Otpremnica / Prijemnica)
tblNovac.OtkupID        -> header ID
tblPaletaStavka         -> ZbirnaID (bilo BrojZbirne)
```

### 3.1) `Otkup.OtpremnicaID` — header, i to je ODLUKA

Danas se piše po fizičkom (klasnom) redu — `modDokumenta.bas:4266`,
`modMasterSync.bas:2381` — pa šema **dozvoljava** da dve klase istog bloka odu na
dve otpremnice. Bez postojećih podataka to se više ne može pročitati iz baze.

Odluka: **header**. Otkupni blok je „jedan otkup od jednog kooperanta, na jednom
otkupnom mestu, jednog dana" (`docs/DOMEN/README.md` §1) i fizički ide na jednu
otpremnicu. Ako se ikad pojavi potreba za delimičnom alokacijom, uvodi se
eksplicitna alokaciona tabela — ne rasplinjava se FK na stavku „za svaki slučaj".

Isto važi za `ZbirnaID` na otkupu: header, denormalizovan (nasleđen od
otpremnice), i **ne** koristi se kao kanonska membership veza. Kanonska
membership je uvek `Otpremnica.ZbirnaID`.

---

## 4) Poslovni broj je labela

Posle refaktora `BrojDokumenta` / `BrojOtpremnice` / `BrojZbirne` /
`BrojPrijemnice` smeju samo za: pretragu, prikaz, generator sledećeg broja,
duplicate-validaciju pri unosu.

Ne smeju za: FK, storno identitet, izbor „svih redova dokumenta", rekalkulaciju
roditelja, correction identitet, ownership guard.

UI mreža i picker nose stvarni `DocumentID` u skrivenoj koloni; akcija radi nad
ID-em, nikad nad tekstom broja.

---

## 5) Šema iz koda — `modSchema.bas` (novo)

Danas šema osnovnih tabela **ne postoji nigde u kodu**. Doslovno iz
`tools/make_fixture.py`:

> „osnovna šema (sheetovi + ListObject-i sa kolonama) ne postoji nigde u kodu —
> Ensure\* rutine u modSetup samo DODAJU kolone na postojeće tabele, a spiskovi
> kolona osnovnih tabela žive isključivo u .xlsm."

To se menja. Cilj: **sveska se može obrisati do praznog fajla, kod je ponovo
napravi.**

### 5.1 Sadržaj

Jedan modul, jedna deklaracija po tabeli:

```vba
' TabelaSpec: ime | sheet | kolone | nosi Stornirano | nosi audit | nosi trace
Private Function SpecOtkup() As Variant
    SpecOtkup = Array(TBL_OTKUP, "Otkup", Array( _
        COL_OTK_ID, COL_OTK_DATUM, COL_OTK_KOOPERANT, ...))
End Function
```

Registar svih ~51 tabele + 4 nove. Bootstrap se **ne kuca ručno**:
`tools/dump_schema.py <donor> --json` već ispisuje punu šemu, pa se iz tog JSON-a
generiše VBA deklaracija jednim prolazom. Posle toga sveska prestaje da bude izvor
istine za šemu.

### 5.2 API

| Procedura | Šta radi |
|---|---|
| `EnsureAllTables()` | kreira nedostajuće tabele, dopunjava nedostajuće kolone. Idempotentno. |
| `VerifySchema() As Collection` | vraća listu odstupanja (tabela nema, kolona nema, višak kolone). Ništa ne menja. |
| `SchemaReadyOrFail(sourceName)` | tvrda kapija: ako fali tabela ili kolona bez koje se dokument ne može upisati — diže grešku sa imenom. |

`EnsureDataTable` iz `modSetup` (`modSetup.bas:1785`) je već dokazan obrazac —
tako su rođeni `tblUtovar`/`tblUtovarStavke` i `tblStornoVeze`/`tblStornoZurnal`.
`modSchema` ga koristi, ne zamenjuje.

### 5.3 Dve stvari koje se moraju popraviti u istom koraku

**a) Tiho padanje.** `EnsureRuntimeSchema` otvara sa `On Error Resume Next`
(`modSetup.bas:1213`). Pozvana procedura bez sopstvenog handlera propušta grešku
gore, gde bude progutana. Sa starim modelom to je bila degradacija (fali kolona);
sa novim znači **da model dokumenta ne postoji**. Svaka tabela ide kroz guard sa
tragom, po uzoru na `EnsureKolonaSaTragom` (`modSetup.bas:1286`) — pad jedne se
zapiše i ne zaustavlja ostale.

**b) Readiness kapija pre prvog upisa.** Ako `tblOtkupStavke` fali, unos se ne sme
pustiti pa da pukne na `AppendRow`-u. `SchemaReadyOrFail` se zove na ulazu u svaki
`Create*_TX`.

### 5.4 Statička kapija

Novo `vba_check` pravilo `SEMA_REGISTAR`, po uzoru na postojeće `STORNO_REGISTAR`:
svaka `TBL_*` konstanta mora biti imenovana u `modSchema` registru. Tabela koja
postoji kao konstanta a nije u registru = pad CI-ja. Sabotaža uz pravilo obavezna.

### 5.5 Uz šemu idu još dve stvari u istom PR-u

**`NewEntityID(prefix)`** — jedna fabrika opaque ID-eva za ceo sistem. Novi
transakcioni PK je `PREFIX-<32 hex>`; matični podaci zadržavaju `KOOP-123` format.
Detalji i obrazloženje: `DOCUMENT_HEADER_LINES.md` §2.

**`WRITE_OWNERSHIP.json` + `who_writes.py --check-ownership`** (ugovor A11) —
imenovana lista modula koji smeju da pišu svaku domen-tabelu; nov pisač obara CI.
Ovo ne traži nove module ni slojeve, samo jedan ulaz po tabeli. Meta:
`tblOtkup` sa 12 pisača na ≤ 3.

Oba idu u temeljni PR jer ih je posle prvog dokumenta skuplje uvesti nego pre.

### 5.6 Posledica za CLAUDE.md §3

Pravilo „Šema tabela je izvor istine, ne kod" se **obrće**. Novo pravilo:

> Registar u `modSchema` je izvor istine. Sveska koja odstupa je drift i
> `VerifySchema` ga prijavljuje.

Izmena `CLAUDE.md` ide **zasebnim process PR-om** (CLAUDE.md §6), ne zajedno sa
kodom.

---

## 6) Write boundary

Po jedan javni ulaz po dokumentu, koji vraća **jedan** ID:

```vba
Public Function CreateOtkup_TX(ByRef h As Object, ByVal stavke As Collection) As String
Public Function CreateOtpremnica_TX(ByRef h As Object, ByVal stavke As Collection) As String
Public Function CreateZbirna_TX(ByRef h As Object, ByVal stavke As Collection) As String
Public Function CreatePrijemnica_TX(ByRef h As Object, ByVal stavke As Collection) As String
```

Obrazac je već u repou i radi: `CreateFaktura_TX(kupacID, stavke As Collection)`
(`modFaktura.bas:11`) — `_TX` drži transakciju i monitoring, `Private` core radi
posao, kompletna prevalidacija pre ijednog upisa, `RequireColumnIndex` fail-fast.
Kopira se, ne izmišlja.

**DTO:** `Scripting.Dictionary` za header, `Collection` diktova za stavke. Bez
novih klasa, bez nasleđivanja, bez generičkog repozitorijuma.

**Adapter:** `modOtkupUnos` / `modDokUnos` i dalje čitaju F1–F4 polja iz forme i
prave DTO. Cutover površina je iznenađujuće mala — **po jedan stvarni callsite po
dokumentu**:

| Funkcija | Jedini produkcioni callsite |
|---|---|
| `SaveOtkupMulti_TX` | `modOtkupUnos.bas:275` |
| `SaveOtpremnicaMulti_TX` | `modDokUnos.bas:262` |
| `SaveZbirnaMulti_TX` | `modDokUnos.bas:533` |
| `SavePrijemnicaMulti_TX` | `modDokUnos.bas:934` |

**Bez compatibility wrappera.** Nema legacy pozivalaca; stari `Save*Multi_TX` i
`Save*` core funkcije se brišu u istom PR-u u kom sleti zamena.

**Transakcija** snapshotuje sve što operacija dira. Merene današnje granice:

| Dokument | Tabele u `AddTableSnapshot` danas | Dodaje se |
|---|---|---|
| Zbirna | `tblZbirna` | `tblZbirnaStavke` |
| Otpremnica | `tblOtpremnica`, `tblAmbalaza` | `tblOtpremnicaStavke` |
| Otkup | `tblOtkup`, `tblAmbalaza`, `tblNovac` | `tblOtkupStavke` |
| Prijemnica | `tblPrijemnica`, `tblAmbalaza`, `tblFakturaStavke`, `tblFakture`, `tblPaleta`, `tblPaletaStavka` | `tblPrijemnicaStavke` |

Napomena: `clsTransaction.RestoreTable` **diže grešku** na neslaganje broja kolona.
Zato `EnsureAllTables` nikad ne sme da se pozove unutar otvorene transakcije.

---

## 7) Storno

```vba
StornoZbirna_TX(zbirnaID)
```

1. `RequireSingleHeader(TBL_ZBIRNA, COL_ZBR_ID, zbirnaID)`
2. provera da li je storno dozvoljen
3. `Stornirano = "Da"` na headeru
4. dokument-level posledice (ambalaža, novac, detach dece)
5. commit

Stavke **nemaju** svoj `Stornirano` — line-level storno ne postoji u domenu.
`ExcludeStornirano` se nad `tbl*Stavke` ne zove; te tabele idu u `BEZ_STORNA`
registar u `modSchemaGuard`, uz komentar zašto (nisu matični podaci, nego deca
čiji status drži roditelj). Bez tog unosa `vba_check` pravilo `STORNO_REGISTAR`
pada — i tako i treba.

Čitač aktivnih stavki: `stavke(docID)` filtrira po FK, a aktivnost dokumenta se
pita headeru. Jedna funkcija po dokumentu, ne ponavljati join na 40 mesta.

Žurnal storna identifikuje operaciju preko `DocumentID`; poslovni broj ostaje kao
display podatak u zapisu.

### 7.1) Storno otpremnice → **rekalkulacija zbirne** (odluka, 10.09.2026)

```
StornoOtpremnica(otpremnicaID)
  1. otpremnica + njena ambalaza -> Stornirano
  2. AKO ima ZbirnaID:  rekalkulisi zbirnu na preostale AKTIVNE otpremnice
  3. AKO vise nijedna ne ostane: stornirati i zbirnu
```

Zbirna **jeste** agregat svojih otpremnica (`DOCUMENT_HEADER_LINES.md` §6.2), pa
je rekalkulacija jedina opcija koja tu definiciju drži tačnom. Kaskada bi
oborila zbirnu i kad na njoj ima drugih aktivnih otpremnica; zabrana bi
blokirala legitimnu ispravku jedne otpremnice. Puna argumentacija i zatečeno
stanje: `docs/DOMEN/GOLDEN_SCENARIJI.md` §10.

Tri stvari koje ovaj korak menja u zatečenom kodu:

| Danas | Posle |
|---|---|
| malina mod kaskadira, ostali modovi ne rade ništa (`modStorno.bas:370`) | jedno pravilo; malina prestaje da bude poseban slučaj — prazna zbirna se stornira, što je isti ishod |
| rekalkulacija ide **po broju** i staje na dvosmislen broj (`ZbirnaMutRazlog`) | po `ZbirnaID`; ta kapija nema više posao |
| `modStornoFlow.RunSimpleStornoOtpremnica` sadrži tačno ovo pravilo i **nema nijednog produkcionog pozivaoca** | pravilo živi u writeru; mrtva kopija se briše |

Pravilo ide u **writer**, ne u flow sloj — inače ga zaobiđe svaki drugi ulaz,
što je tačno ono što se i desilo. Nizvodni dijalog kad postoji prijemnica ili
paleta (`CorrectionNeedsDialog`) ostaje nepromenjen.

#### Šta znači da klasa nestane iz keša

Otvoreno pitanje koje rekalkulacija otvara, i koje mora biti rešeno **pre** nego
što se PR4 napiše:

```
Zbirna ima:  I = 400,  II = 600
stornira se POSLEDNJA otpremnica Klase II
posle:       I = 400,  II = ?
```

`tblZbirnaStavke` nema `Stornirano`, a writer zabranjuje količinu 0 — dakle
„II = 0" nije legalno stanje.

**Odluka: rekalkulator briše red keša koji više nema izvor.** Stavke su izvedeni
keš (§4.3); kad klasa nestane iz izvora, nestaje i iz keša. Line-level storno se
**ne uvodi** — status i dalje drži header, a `BEZ_STORNA` registar ostaje tačan.
Brisanje ide u istoj transakciji kao i rekalkulacija ostalih klasa.

Alternativa (ostaviti red sa nulom) bila bi gora na dva načina: pravila bi
razliku između „nije bilo Klase II" i „bila pa nestala" tamo gde je izvor već
nosi, i probila bi sopstveno pravilo da količina mora biti veća od nule.

---

## 8) Zbirna invarijanta

Pravilo se **ne menja**: zbirna = zbir svojih aktivnih otpremnica, po klasi, hard.
Menja se samo identitet preko kog se računa.

```
pre:   SumOtpremniceByKlasa(brojZbirne)   -- join po BrojZbirne
posle: SumOtpremniceByKlasa(zbirnaID)     -- join po Otpremnica.ZbirnaID
```

1. headeri otpremnica sa `ZbirnaID = X`, aktivni
2. njihove stavke
3. suma po klasi
4. poredi sa `tblZbirnaStavke` gde `ZbirnaID = X`

Isti `BrojZbirne` na drugom `ZbirnaID` više nije problem integriteta — pa
`RequireJedanVlasnikIkadPoBroju`, `historicalOwnerCount`, `activeLogicalCount` i
cela `ZbirnaIdent` mašinerija prestaju da imaju posao.

Time se zatvara i rupa koju `docs/DOMEN/ZBR_IDENTITET.md` §3 sam priznaje:
*„prijemnica može da se veže na pogrešan red i nijedan sloj to ne prijavljuje"*.
Sa FK po ID-u to stanje strukturno ne postoji.

---

## 9) Ispravke (correction)

Append-only + storno + reizdavanje ostaje. Identitet postaje ID-based:

| Danas | Posle |
|---|---|
| `IspravkaOd` (nosi **broj**) | `IspravkaOdID` |
| `ZamenjenSa` (nosi **broj**) | `ZamenjenSaID` |
| `CorrectionID` | ostaje |

Nova verzija dobija **nov** `DocumentID` i kad poslovni broj ostaje isti. Time
`GeneracijaID` gubi i poslednji posao — razlikovanje originala od ispravke.

---

## 10) Štampa i read-modeli

```
pre:   OutputOtkupniList("OTK-1 + OTK-2")
posle: OutputOtkupniList("OTK-1")
```

Loader: header po ID → sve stavke tog ID → dokument. Nijedan print sloj više ne
sme da sastavlja dokument traženjem istog poslovnog broja.

Devet `Split(x, " + ")` mesta se **briše**, ne prilagođava.

---

## 11) Brisanja

Ovo je pola vrednosti refaktora i mora biti u planu eksplicitno, ne „kad
stignemo".

### 11.1 Identitetska mašinerija — briše se cela

`modDokumenta.bas:71-760` (~690 linija):

- `Public Type ZbirnaIdent` (14 polja)
- `ZbirnaIdentResolve` (136 linija)
- `ZbirnaNovUnosRazlog`, `ZbirnaSmeNovUnos`
- `ZbirnaMutacijaPoBrojuRazlog`
- `ZbirnaRoditeljRazlog`, `ZbirnaRoditeljOK`
- konstante `ZBR_MUT_*`, `ZBR_INT_*`, `ZBR_RES_*`

Dalje, po repou:

- `NewGeneracijaID`, `ApplyGeneracijaID`, `ApplyNovaGeneracijaID`,
  `GeneracijaIDZaBrojArr`, `GeneracijaPoID`
- `ZbirnaGeneracijaZaBroj`, `ZbirnaJedinaGeneracijaIkadZaBroj`
- `RedJeIzabranogDokumenta` (`modStorno.bas:87`) i sva 4 poziva
- `StornoOtkupByBrDok_TX`, `StornoOtpremnicaByBroj_TX`,
  `StornoPrijemnicaByBroj_TX`, `StornoZbirna(brojZbirne, ...)`
- `RequireJedanVlasnikPoBroju`, `RequireJedanVlasnikIkadPoBroju`, `VlasniciPoBroju`
- `FindSingleActiveRow`, `ZbirnaVlasnikKljuc`
- `BackfillDeteZbirnaGeneracija`
- `PoveziDeteNaZbirnu`, `ZavrsiVezuOtpremniceNaZbirnu` (postaju običan FK upis)
- kolone `COL_GENERACIJA_ID`, `COL_DETE_ZBIRNA_GEN` i njihovi
  `EnsureKolonaSaTragom` pozivi

Mera: **143 produkcione + 83 test linije** pominju `GeneracijaID` / `ZbirnaIdent`
/ `RedJeIzabranogDokumenta`.

### 11.2 Stari writeri

`SaveOtkup`, `SaveOtkupMulti_TX`, `SaveOtkup_TX`, `SaveOtpremnica`,
`SaveOtpremnicaMulti_TX`, `SaveZbirna`, `SaveZbirnaMulti_TX`, `SavePrijemnica`,
`SavePrijemnica_TX`, `SavePrijemnicaMulti_TX`.

### 11.3 Kolone koje nestaju

`tblOtpremnica.BrojZbirne`, `tblPrijemnica.BrojZbirne`, `tblOtkup.BrojZbirne`
(zamenjuje `ZbirnaID`), `tblPaletaStavka.BrojZbirne`, `tblOtkup.BrojOtpremnice`
(denorm poslovni ključ, `modSetup.bas:1270`), `GeneracijaID` i
`ZbirnaGeneracijaID` sa svih 6+4 tabele.

### 11.4 Testovi koji se brišu, ne portuju

Ceo `ZBR-IDENT-01` / `ZBR-MUT-01` / `ZBR-NORM-02` / `ZBR-CHILD-01` acceptance set
(`docs/DOMEN/ZBR_IDENTITET.md` §9, §13, §14, §15) — testira mehanizam koji
prestaje da postoji. Isto i sabotaže vezane za njih (`tools/sabotaza.py`, katalog
u `vba_check.check_katalog_sabotaza`): sabotaža bez koda koji sabotira je crvena
kapija koja meri prazno.

### 11.5 Dokumentacija

`docs/DOMEN/ZBR_IDENTITET.md` se **ne briše**. Dobija zaglavlje
`SUPERSEDED — v. DOCUMENT_HEADER_LINES.md` i pasus zašto je `GeneracijaID` bio
ispravno rešenje za multi-row model. Istorija odluke ostaje čitljiva.

`docs/DOMEN/README.md` §1 i §2 se ažuriraju (lanac, invarijanta po ID-u).

---

## 12) Testovi

Baseline: čist `main` je **186/4** (četiri testa već padaju). Svaka tvrdnja o
zelenom se poredi sa tim, ne sa nulom.

Obavezni scenariji:

| Test | Tvrdnja |
|---|---|
| `DveKlase_JedanID` | dvoklasni upis → jedan header, dve stavke, jedan vraćen ID |
| `SamoKlasaI` / `SamoKlasaII` | jedan header, jedna stavka |
| `BrojNijeIdentitet` | dva dokumenta sa istim brojem: storno jednog ne dira drugi, print jednog ne čita drugi, invarijanta jednog ne vidi drugi |
| `ZbirnaFK` | otpremnica sa `ZbirnaID=X` ulazi u invarijantu; ista `BrojZbirne` na `ZbirnaID=Y` ne ulazi |
| `StornoPoID` | jedan storno headera = jedan logički dokument |
| `IspravkaID` | original i ispravka imaju različite ID-eve i vezu `IspravkaOdID` |
| `PrintDvoklasni` | dvoklasni dokument se štampa kao jedan sa dve stavke |
| `FakturaStavkaSource` | faktura referencira tačnu `PrijemnicaStavkaID`, ne pogrešnu klasu |
| `NovacBezPrimary` | dvoklasni otkup: `Isplaceno` na headeru, bez dupliranja isplate |
| `AmbalazaStorno` | storno poništava sva packaging kretanja bez oslanjanja na dva stara row ID-a |
| `SemaSamoLeci` | obrisana `tblOtkupStavke` → `EnsureAllTables` je vraća; `VerifySchema` je pre toga prijavio |
| `SemaKapija` | obrisana tabela → `CreateOtkup_TX` pada sa imenom tabele, ne na `AppendRow`-u |
| `AutoHladnjaca` | postojeći auto-lanac funkcionalno identičan |

Dokaz u oba smera (pokvari → pukne **po imenu** → vrati → zeleno) obavezan za:
`SemaKapija`, `BrojNijeIdentitet`, `ZbirnaFK`, `NovacBezPrimary` — kritične
poslovne invarijante i nov checker.

**Fixture:** `tests/fixtures/otkup_test.xlsm` se regeneriše. Redosled: donor →
jedan prolaz `EnsureAllTables` nad **kopijom** → `tools/make_fixture.py --donor
<kopija>`. Bez toga nove tabele u fixture-u ne postoje i nijedan novi test se ne
može ni napisati.

`run_vba` traži Windows + Excel + pywin32. Iz web sesije se izmena ponašanja
prijavljuje kao **neverifikovana**, nikad kao zelena.

---

## 13) Statičke kapije

Ne „repo-wide search treba da pokaže", nego imenovana `vba_check` pravila sa
sabotažama:

| Pravilo | Tvrdnja |
|---|---|
| `SEMA_REGISTAR` | svaka `TBL_*` konstanta je u `modSchema` registru |
| `NEMA_GENERACIJE` | nijedan produkcioni modul ne pominje `GeneracijaID` / `ZbirnaGeneracijaID` |
| `NEMA_ID_PLUS_ID` | nigde `Split(..., " + ")` nad ID stringom; nijedan `Create*_TX` ne vraća konkatenaciju |
| `NEMA_BROJA_KAO_FK` | `BrojZbirne` se ne koristi kao join ključ (dozvoljen samo u prikazu / generatoru / validaciji unosa) |
| `STORNO_REGISTAR` | postojeće; dopunjuje se sa 4 nove `tbl*Stavke` u `BEZ_STORNA` |

`who_writes.py --check` se regeneriše posle svakog dokumenta (nova tabela = novi
pisci).

---

## 13a) Kapija je merila stil pisanja poziva, ne vlasništvo (nađeno u PR3)

Pre nego što je PR3 dodao nov pisač nad `tblZbirna`, provera je pokazala da A11
kapija **ne vidi polovinu upisa**. `AppendRow` je funkcija i pola koda je zove
kao funkciju:

```vba
AppendRow TBL_ZBIRNA, rowData          ' naredba  -- kapija je videla
n = AppendRow(TBL_ZBIRNA, rowData)     ' funkcija -- kapija NIJE videla
```

Regex je tražio razmak posle imena mutatora. Posledica: **22 poziva nevidljivo**,
među njima produkcioni upisi nad `tblZbirna` (`modDokumenta`, `modMasterSync`),
`tblOtkup` (`modOtkup`, `modMasterSync`), `tblPrijemnica`, `tblOtpremnica`,
`tblNovac`, `tblFakturaStavke` — a **sedam tabela** (`tblCenovnik`,
`tblKooperanti`, `tblMagacin`, `tblPartnerMap`, `tblSEFEventLog`,
`tblStornoZurnal`, `tblVozaci`) uopšte nije bilo u registru vlasništva. Kapija je
sve to vreme bila **zelena**.

Isti kvar kao raniji `RequireUpdateCell` (nema granice reči pre `UpdateCell`) —
dva puta ista bolest, oba puta nevidljiva, oba puta nađena slučajno. Zato oblik
poziva sada ima **sopstvene slučajeve**: `who_writes.py --self-test`, pozitivni i
negativni, u CI-ju.

`WRITE_OWNERSHIP.json` je re-baseline-ovan. Dodati pisači **nisu novi** — bili su
neizmereni; baseline je bio zamrznut prema slepom skeneru, pa je zamrzao
nepotpunu stvarnost. Ništa nije uklonjeno.

**Treća rupa istog roda, nađena u reviziji PR3:** skener je čitao **red po red**,
pa mu je prelomljen poziv bio nevidljiv:

```vba
n = AppendRow( _
        TBL_ZBIRNA, rowData)
```

Takav oblik danas u `src-vba/` ne postoji, ali kapija ne sme da zavisi od toga
gde je neko prelomio red. Sada se VBA nastavci (` _`) spajaju pre regexa —
pažljivo, jer bi naivna verzija otvorila **novu** rupu: komentar koji se završava
sa ` _` progutao bi sledeću liniju i sakrio pravi `AppendRow` ispod sebe. I to
ima svoj slučaj u `--self-test`.

> Pouka koja važi i za ostatak refaktora: kapija koja nikad nije pokazana crvena
> ne dokazuje da išta meri — a kapija koja stoji na jednom regexu meri tačno
> onoliko oblika koliko je taj regex video kad je pisan.

---

## 14) Redosled

Merena cena po dokumentu:

| Dokument | Tabela u TX | Produkcionih pisaca | Redosled |
|---|---|---|---|
| Zbirna | 1 | 5 | **1.** |
| Otpremnica | 2 | 4 | **2.** |
| Otkup | 3 | 12 | **3.** |
| Prijemnica | 6 | 4 | **4.** |

Zbirna prva: najmanja transakcija, a u njoj živi **cela** kompenzaciona mašinerija
koju brišemo. Prijemnica poslednja: šest tabela, i njene stavke hrane fakturu.

### PR-ovi

**PR0 je gotov** — model za sva četiri dokumenta pre implementacije bilo kog:
`docs/DOMEN/ARCHITECTURE_CONTRACT.md` (A1–A12, četiri kapije, Pre-Flight) i
`docs/DOMEN/DOCUMENT_HEADER_LINES.md` (grain, PK/FK, kardinaliteti, izvor istine,
invarijante). Razlog za to: refaktor Zbirne pre nego što je poznat grain
Prijemnice je lokalna optimizacija jednog dela lanca — tačno način na koji je
`GeneracijaID` i nastao.

| # | Sadržaj | Zavisi od |
|---|---|---|
| 0 | ✅ **Ugovor + model** — `ARCHITECTURE_CONTRACT.md`, `DOCUMENT_HEADER_LINES.md` | — |
| 1 | ✅ **Temelj**: `modSchema` registar svih tabela + `VerifySchema` + `SchemaReadyOrFail`; `NewEntityID` fabrika; `WRITE_OWNERSHIP.json` + `who_writes.py --check-ownership`; pravilo `SEMA_REGISTAR` + self-test. **Bez ijedne nove tabele.** | 0 |
| 2 | ✅ **Kanon šeme u gitu** (`schema/schema.json` → `gen_schema_module.py` → `modSchema.bas`) + tri CI kapije; **golden mreža, 12 zaključanih scenarija**; testovi `SemaSamoLeci` / `SemaKapija` / `PrefiksNijeString` | 1 |
| 3 | ✅ **Zbirna header+stavke**: `tblZbirnaStavke`, `Otpremnica.ZbirnaID`, `CreateZbirna_TX(h, izvorOtpremnice, outGreska, ocekivano)` — stavke se **izvode iz izvornih otpremnica**, membership ide u **istoj** transakciji, opaque `ZbirnaID` i `ZbirnaStavkaID` oba fail-closed. **Aditivno** — produkcija još ide starim putem, golden 12/0 nepromenjen. Uz to: A11 kapija je bila slepa na funkcijski i na prelomljen oblik `AppendRow` (v. §13a) | 2 |
| 4 | **Zbirna cutover**: invarijanta po ID-u, `StornoZbirna_TX(id)`, `RecalculateZbirna_TX(id)`, **rekalkulacija zbirne pri stornu otpremnice (§7.1)**, print, izveštaji, testovi. **Briše `ZbirnaIdent*`, `ZbirnaGeneracija*` i mrtvu `RunSimpleStornoOtpremnica`.** Registruje golden D1. | 3 · **§7.1 odlučen** |
| 5 | **Otkup header+stavke**: `tblOtkupStavke`, `CreateOtkup_TX` | 4 |
| 6 | **Otkup integracije**: ambalaža na header, novac na header, `Isplaceno` izvedeno, storno, ispravka, print, auto-hladnjača | 5 |
| — | **KAPIJA ODLUKE** — v. §14.1 | 6 |
| 7 | **Otpremnica** header+stavke + cutover | 6 |
| 8 | **Prijemnica** header+stavke + cutover | 7 |
| 9 | **Faktura**: `FakturaStavka.PrijemnicaStavkaID` | 8 |
| 10 | **Paleta**: `PaletaStavka.PrijemnicaStavkaID` | 8 |
| 11 | **Sledljivost kao graf** nad eksplicitnim FK; ukloniti heuristički AutoLink | 10 |
| 12 | **E2E + brisanje**: `COL_GENERACIJA_ID`, `COL_DETE_ZBIRNA_GEN`, `*ByBroj_TX`, svih 9 `Split(" + ")`, mrtvi testovi i sabotaže; pravila `NEMA_GENERACIJE` / `NEMA_BROJA_KAO_FK` / `NEMA_ID_PLUS_ID`; `ZBR_IDENTITET.md` → superseded | 11 |
| — | `CLAUDE.md` §3 (obrtanje pravila o izvoru istine šeme) | zaseban process PR |

PR 12 je ključan i **ne sme se preskočiti**: dok `GeneracijaID` postoji kao živ
runtime mehanizam, refaktor nije završen — to je dual identity model, gori od
sadašnjeg.

### 14.1) Kapija odluke posle PR 6

Posle Zbirne i Otkupa — dva najreprezentativnija slajsa — donosi se formalna
odluka: **nastavak u mestu** ili **novo stablo koda**. Kriterijumi su merljivi,
sa alatima koji već postoje.

**Nastavljamo u mestu ako:**

| Metrika | Kako se meri | Prag |
|---|---|---|
| Pisaca nad `tblOtkup` | `who_writes.py` | 12 → ≤ 3 |
| Linija u jezgru upisa otkupa | `wc -l` nad `CreateOtkup_TX` + core | kraće od zbira `SaveOtkup*` danas |
| Resolver/fallback logika | `grep` za resolve/fallback/scope u jezgru | 0 |
| Identitetski testovi | broj `T_ZBR*` / generacija testova | pada, ne raste |
| Reuse UI/SEF/Banka/izveštaja | koliko modula je moralo da se prepiše | većina samo adaptirana |
| Dual model | postoji li putanja koja piše i staro i novo | **ne sme postojati** |

**Prelazimo na novo stablo ako:**

- više od polovine relevantnih pozivalaca mora da se prepiše bez reuse-a
- svaki legacy modul traži compatibility facade
- novi model mora da živi paralelno sa starim kroz više faza
- test corpus postaje neupotrebljiv
- čisto jezgro se ne može izolovati od stare šeme

Ako `CreateOtkup_TX` — sa headerom, stavkama, ambalažom i novcem — ispadne kratka
i čitljiva operacija, pitanje rewrite-a je zatvoreno. Ako se i tada pojave
resolveri, scope, historical ownership i dual mode, imamo **empirijski** dokaz da
šema nije bila jedini problem, i tada je `/agrix2` opravdan — sa dva dokazana
slajsa kao specifikacijom, umesto sa osećajem.

---

## 15) Backlog — namerno van opsega

| Stavka | Zašto ne sada |
|---|---|
| **PWA / MasterSync ingest** | radi se isključivo VBA. Nalaz koji čeka: PWA šalje **jedan record = jedna klasa = ceo dokument**, sa svežim `brojDokumenta` po svakom snimanju (`src/js/features/otkup/otkup-form.js:643`, `:493`). Ingest postaje 1 record → 1 header + 1 stavka; **nema heurističkog grupisanja i ne treba eksterni Document UID**. `ClientRecordID` ide na header, a `IsDuplicateInMaster` (`modMasterSync.bas:1824`) mora da se prepokaže na header tabelu — inače se svaki PWA dokument reimportuje. |
| **Self-update** | van opsega po dogovoru |
| **App / Repo / Qry slojevi** | **Ne paralelno sa refaktorom** — pokvarilo bi kapiju odluke iz §14.1: dve promenljive odjednom znače da se ne može reći da li je čist ishod zasluga šeme ili slojeva. Uz to, App sloj već postoji neimenovan (`mod*Unos` prima DTO rečnik, `NoviOtpremnicaUnos`), a enforcement daje A11 allowlist, ne ime modula. Jedini sloj koji stvarno nedostaje je **Qry** (`modDokumenta`: 15 javnih čitača pored 21 mesta upisa) — ali dobar deo tih čitača postoji da rekonstruiše dokument po broju i **umire u PR 12**. Revidirati **posle PR 12**, kad se zna koji čitači preživljavaju. Do tada: čitanja u novim writer-ima idu iza imenovanih funkcija, ne inline skenova. |
| **Delimična alokacija Otkup → Otpremnica** | nema poslovnog zahteva; ako se pojavi — eksplicitna alokaciona tabela |
| **Mešovit dokument (više vrsta u jednom)** | vrsta/sorta ostaju header; nije zahtev |

---

## 16) Definicija gotovog

- „ZBR-123 je jedan red u `tblZbirna`, njene klase su redovi u `tblZbirnaStavke`."
- „`Otpremnica.ZbirnaID = ZBR-123` je prava veza."
- „PRJ-789 je jedna prijemnica bez obzira ima li jednu ili dve klase."
- „Faktura stavka zna tačnu `PrijemnicaStavkaID`."
- „Storno prima `DocumentID`. Štampa prima `DocumentID`. Invarijanta prima `ZbirnaID`."
- „Broj dokumenta je labela."
- „`GeneracijaID` ne postoji."
- „Sveska se može obrisati — kod je vrati."

Sve dok bilo koja od ovih rečenica nije tačna, refaktor nije završen.
