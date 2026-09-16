# Refaktor: dokument = header + stavke

> Status: **PR0–PR2 mergovani, PR3 u reviziji; od Otkup skele nadalje je plan.** Tačno
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
| Otkup | `RedniBroj`, `Klasa`, `Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg` | `Kolicina` je uvek **neto**; `BrutoKg` samo kod bruto unosa (§4.1d) |
| Otpremnica | `RedniBroj`, `Klasa`, `Kolicina`, `KolAmbalaze`, `BrutoKg` | **bez `Cena`** — izvedeni dokument, izvori mogu imati različite cene (§13b) |
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
`tblNovac.OtkupID` pokazuje na header, a **`Isplaceno` uopšte nije polje** nego
read-model:

```
Placeno   = SUM(tblNovac vezan na OtkupID)
Isplaceno = (Vrednost - Placeno) <= 0
```

Primary-row hack nestaje, a `modNovac` prestaje da bude pisač `tblOtkup`.
Imenovane zamene za dva zatečena čitaoca: `DOCUMENT_HEADER_LINES.md` §4.1c.
Ovo je jedini slučaj gde refaktor menja zatečeno ponašanje — i menja ga zato što
je zatečeno ponašanje pogrešno.

---

## 3) Ciljna šema

Prefiksi po postojećoj konvenciji (`PLS-` za `tblPaletaStavka`): `OKS-`, `OPS-`,
`ZBS-`, `PRS-`.

### `tblOtkup` (header)

```
OtkupID           PK, "OTK-"
BrojDokumenta     poslovni broj -- LABELA
Datum
KooperantID       FK
StanicaID         FK
ParcelaID         FK
KulturaID         FK -- razresen, NIKAD fabrikovan
VrstaVoca
SortaVoca
TipAmbalaze
KolAmbIzdata      dokument-level (OM izdao prazne kooperantu)
ClientRecordID    eksterni identitet (PWA)
SyncSource        poreklo zapisa
Stornirano
IzdatoStatus
IspravkaOdID / ZamenjenSaID / CorrectionID
CreatedAt / CreatedBy / ModifiedAt / ModifiedBy
SourceCreatedAt   vreme nastanka na izvoru (PWA); desktop koristi CreatedAt
```

Nema: `VozacID` (pripada otpremnici) · `Isplaceno` / `DatumIsplate` (izvedeno iz
`tblNovac`) · `VremeUnosa` (udvajanje sa `CreatedAt`) · `Novac` /
`PrimalacNovca` (keš ne ulazi kroz otkup) · `OtpremnicaID` / `ZbirnaID` /
`BrojOtpremnice` / `BrojZbirne` (pripadnost je `tblOtpremnicaIzvori`) ·
`Klasa` / `Kolicina` / `Cena` / `KolAmbalaze` / `BrutoKg` (stavka) ·
`GeneracijaID` / `ZbirnaGeneracijaID`.

Obrazloženje po koloni i imenovane zamene: `DOCUMENT_HEADER_LINES.md` §4.1c.

### `tblOtkupStavke`

```
OtkupStavkaID     PK, "OKS-"
OtkupID           FK, obavezan
RedniBroj
Klasa
Kolicina          UVEK NETO kg -- zamrznuto pri izdavanju
Cena              STVARNO PRIMENJENA cena (cenovnik je samo predlog)
KolAmbalaze
BrutoKg           zamrznut original -- popunjen SAMO kad je unos bio bruto
CreatedAt / CreatedBy / ModifiedAt / ModifiedBy
```

Težina ambalaže se koristi samo u trenutku nastanka; `BrutoKg` i `Kolicina` se
nikad ne rekalkulišu iz `tblTipAmbalaze` (`DOCUMENT_HEADER_LINES.md` §4.1d).

### `tblOtpremnica` (header)

```
OtpremnicaID  PK "OTP-" | Datum | StanicaID | VozacID | BrojOtpremnice
VrstaVoca | SortaVoca | TipAmbalaze
Cena          predlog za prefill blokova -- NE-FINANSIJSKO polje (S13b)
Stornirano | trace | audit
```

Bez `ZbirnaID`: pripadnost zbirnoj zna **`tblZbirnaIzvori`**, ne kolona na
otpremnici.

### `tblOtpremnicaStavke`

```
OtpremnicaStavkaID PK "OPS-" | OtpremnicaID FK | RedniBroj
Klasa | Kolicina | KolAmbalaze | BrutoKg
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

### Tabele članstva — sastav VERZIJE dokumenta

```
tblZbirnaIzvori        ZbirnaIzvorID PK "ZBI-" | ZbirnaID FK | OtpremnicaID FK
tblOtpremnicaIzvori    OtpremnicaIzvorID PK "OPI-" | OtpremnicaID FK | OtkupID FK
```

Odgovaraju na pitanje koje kolona ne može: *„od kojih je tačno dokumenata ova
verzija bila sastavljena"* — bez gledanja trenutnog stanja sistema.

**Nepromenljivost počinje pri izdavanju, ne pri upisu** (A15):

| Stanje roditelja | Članstvo |
|---|---|
| `DRAFT` | **promenljivo** — izvori se dodaju i sklanjaju slobodno |
| `IZDATO` / `PROSLEDJENO` | **zamrznuto** — nova verzija dobija svoje redove |

Puno obrazloženje: `ARCHITECTURE_CONTRACT.md` **A15**.

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
   ^
   │ OtkupID
tblOtpremnicaIzvori          <── SASTAV verzije otpremnice
   v
tblOtpremnica ──1:N──> tblOtpremnicaStavke
   ^
   │ OtpremnicaID
tblZbirnaIzvori              <── SASTAV verzije zbirne
   v
tblZbirna ──1:N──> tblZbirnaStavke

Pripadnost drze ISKLJUCIVO tabele *Izvori. Nema pratecih kolona na deci --
ni Otkup.OtpremnicaID ni Otpremnica.ZbirnaID.
   ^
   │ ZbirnaID
tblPrijemnica ──1:N──> tblPrijemnicaStavke
                            ^
                            │ PrijemnicaStavkaID
                       tblFakturaStavke ──N:1──> tblFakture

tblAmbalaza.DokumentID  -> header ID (Otkup / Otpremnica / Prijemnica)
tblNovac.OtkupID        -> header ID
tblPaletaStavka         -> PrijemnicaStavkaID (bilo BrojZbirne)
```

### 3.1) Pripadnost se ne drži kolonom — ni na otkupu ni na otpremnici

Zatečeno stanje: `Otkup.OtpremnicaID` se piše po fizičkom (klasnom) redu
(`modDokumenta.bas:4266`, `modMasterSync.bas:2381`), a `Otpremnica.BrojZbirne` je
labela u ulozi veze.

**Odluka: obe kolone nestaju.** Pripadnost živi u tabelama članstva:

```
tblOtpremnicaIzvori   OtpremnicaID + OtkupID
tblZbirnaIzvori       ZbirnaID     + OtpremnicaID
```

> **Dvaput ispravljena formulacija, i vredi zapisati zašto.**
>
> Prvo je ovde stajalo „kanonska membership je uvek `Otpremnica.ZbirnaID`" — to
> je palo na scenariju sa sestrama (A15): posle ispravke jedne otpremnice, one
> koje se nisu menjale pripadaju **i** staroj **i** novoj verziji zbirne, a jedan
> FK može da pokaže samo jednu.
>
> Zatim je kolona zadržana kao **pokazivač** („gde je sada"), uz test koji
> dokazuje da se poklapa sa članstvom. I to je palo: pokazivač je jedno jeftinije
> čitanje po ceni cele nove klase problema — drift između kanona i keša, provera
> tog drifta, snapshot još jedne tabele, još jedan upis i još dva testa. Bez
> produkcionih podataka nema nikoga kome se to plaća.
>
> Odgovor na „gde je sada" računa se iz članstva
> (`modDokumenta.AktivnaZbirnaZaOtpremnicu`).

Delimična alokacija — da jedna otkupna stavka delimično završi u više otpremnica
— i dalje **nije** modelovana; ako se pojavi, ide zasebna tabela sa `Kg`
(`DOCUMENT_HEADER_LINES.md` §3.2). Članstvo i alokacija nisu isti pojam.

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
Public Function CreateZbirna_TX(ByVal h As Object, _
                                ByVal izvorOtpremnice As Collection, _
                                ByVal ocekivano As Collection, _
                                Optional ByRef outGreska As String) As String

' automatski tok -- bez nezavisnog ocekivanja, i to izricito
Public Function CreateZbirnaIzIzvora_TX(ByVal h As Object, _
                                        ByVal izvorOtpremnice As Collection, _
                                        Optional ByRef outGreska As String) As String
Public Function CreatePrijemnica_TX(ByRef h As Object, ByVal stavke As Collection) As String
```

Obrazac je već u repou i radi: `CreateFaktura_TX(kupacID, stavke As Collection)`
(`modFaktura.bas:11`) — `_TX` drži transakciju i monitoring, `Private` core radi
posao, kompletna prevalidacija pre ijednog upisa, `RequireColumnIndex` fail-fast.
Kopira se, ne izmišlja.

**Stavke se ne primaju — izvode se.** Zbirna i Otpremnica su izvedeni
dokumenti, pa njihovi writeri primaju **izvorne dokumente**, a stavke računaju iz
njih (A13). Otkup, kao primarna činjenica, i dalje prima stavke.

**DTO:** `Scripting.Dictionary` za header, `Collection` ID-eva za izvore,
`Collection` diktova za očekivano. Bez novih klasa, bez nasleđivanja, bez
generičkog repozitorijuma.

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
  2. nadji zbirne cije clanstvo (tblZbirnaIzvori) sadrzi ovu otpremnicu
     i koje NISU stornirane

     zbirna je DRAFT   -> rekalkulisi je in-place iz preostalih izvora
     zbirna je IZDATO  -> stara ostaje NEPROMENJENA i biva superseded;
                          nastaje NOVA verzija (nov ZbirnaID, nov BrojZbirne,
                          IspravkaOdID, isti CorrectionID, novi izvori i stavke)

  3. AKO vise nijedna otpremnica ne ostane: stara se stornira BEZ naslednika
```

Pošto lanac danas nema draft fazu, u praksi važi druga grana. **Nema in-place
rekalkulacije izdatog dokumenta** (A13); puno obrazloženje i tri opcije:
`GOLDEN_SCENARIJI.md` §10.

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
što se cutover napiše:

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
posle: SumOtpremniceByKlasa(zbirnaID)     -- join po tblZbirnaIzvori
```

1. `tblZbirnaIzvori` gde `ZbirnaID = X` → **tačni `OtpremnicaID`-evi te verzije**
2. njihove stavke
3. suma po klasi
4. poredi sa `tblZbirnaStavke` gde `ZbirnaID = X`

> **Tabela članstva je jedini zapis pripadnosti** — kolone na otpremnici nema.
> Da postoji, pomerala bi se pri ispravci sa `ZBR18` na `ZBR19` i invarijanta
> stare više se ne bi mogla reprodukovati. Ovako sastav svake verzije ostaje
> proverljiv (A15).

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

Nova verzija dobija **nov `DocumentID` i nov poslovni broj** (A9). Time
`GeneracijaID` gubi i poslednji posao — razlikovanje originala od ispravke.

> Ranija formulacija je glasila „nov `DocumentID` **i kad poslovni broj ostaje
> isti**", što je ostavljalo prostor da dve verzije dele broj. Za lanac
> `Otkup → Otpremnica → Zbirna` to nije opcija: broj je labela koju operater vidi
> na papiru, pa dva papira sa istim brojem i različitim sadržajem nisu razlučiva
> izvan sistema.

Kolone se u cutover-u **stvarno preimenuju** (`IspravkaOd` → `IspravkaOdID`), ne
pretumače: kolona koja se zove `IspravkaOd` a nosi ID je tačno vrsta
dvosmislenosti koju refaktor uklanja.

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
- `PoveziDeteNaZbirnu`, `ZavrsiVezuOtpremniceNaZbirnu` — **nestaju**; pripadnost se upisuje u `tblZbirnaIzvori` kroz writer, ne kolonom na detetu
- kolone `COL_GENERACIJA_ID`, `COL_DETE_ZBIRNA_GEN` i njihovi
  `EnsureKolonaSaTragom` pozivi

Mera: **143 produkcione + 83 test linije** pominju `GeneracijaID` / `ZbirnaIdent`
/ `RedJeIzabranogDokumenta`.

### 11.2 Stari writeri

`SaveOtkup`, `SaveOtkupMulti_TX`, `SaveOtkup_TX`, `SaveOtpremnica`,
`SaveOtpremnicaMulti_TX`, `SaveZbirna`, `SaveZbirnaMulti_TX`, `SavePrijemnica`,
`SavePrijemnica_TX`, `SavePrijemnicaMulti_TX`.

> **Odluka 15.09.2026 (§14.7):** `SaveOtpremnica` i `SaveOtpremnicaMulti_TX`
> posle PR7 ostaju samo za testove i brišu se u PR8, zajedno sa zbirnim tokovima
> (presedan `SaveOtkup_TX` iz PR6).

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

Baseline: `main` je **zelen na svih 13 suite-ova** (`RunAllTests` 202/0,
`BusinessFlowPro` 608/0, `RunGoldenSuite` 12/0). Ranije je ovde stajalo „186/4 —
četiri testa već padaju"; ta četiri su u međuvremenu popravljena, pa se svaka
tvrdnja o zelenom sada poredi sa **nulom padova**.

Obavezni scenariji:

| Test | Tvrdnja |
|---|---|
| `DveKlase_JedanID` | dvoklasni upis → jedan header, dve stavke, jedan vraćen ID |
| `SamoKlasaI` / `SamoKlasaII` | jedan header, jedna stavka |
| `BrojNijeIdentitet` | dva dokumenta sa istim brojem: storno jednog ne dira drugi, print jednog ne čita drugi, invarijanta jednog ne vidi drugi |
| `ZbirnaClanstvo` | otpremnica upisana u `tblZbirnaIzvori(ZBR-X, OTP-A)` ulazi u sastav `ZBR-X`; ista `BrojZbirne` na drugoj zbirnoj ne menja ništa |
| `StornoPoID` | jedan storno headera = jedan logički dokument |
| `IspravkaID` | original i ispravka imaju različite ID-eve i vezu `IspravkaOdID` |
| `PrintDvoklasni` | dvoklasni dokument se štampa kao jedan sa dve stavke |
| `FakturaStavkaSource` | faktura referencira tačnu `PrijemnicaStavkaID`, ne pogrešnu klasu |
| `NovacBezPrimary` | dvoklasni otkup ima **jedan** `OtkupID`; vrednost = `SUM(stavke)`; isplata se vezuje za taj header; read-model vraća isplaćeno/neisplaćeno. **Nema persistentnog `Isplaceno`** |
| `AmbalazaStorno` | storno poništava sva packaging kretanja bez oslanjanja na dva stara row ID-a |
| `SemaSamoLeci` | obrisana `tblOtkupStavke` → `EnsureAllTables` je vraća; `VerifySchema` je pre toga prijavio |
| `SemaKapija` | obrisana tabela → `CreateOtkup_TX` pada sa imenom tabele, ne na `AppendRow`-u |
| `AutoHladnjaca` | postojeći auto-lanac funkcionalno identičan |

#### Mreža za Otkup skelu — imenovano, pre writer-a

Pre-Flight je platio odluke koje ništa još ne meri. Skela ih mora zaključati:

| Test | Tvrdnja |
|---|---|
| `KulturaSeNeFabrikuje` | nerazrešena `(Vrsta, Sorta)` **pada**; ne nastaje `"Vrsta-Sorta"` string |
| `KulturaDvosmislenaPada` | dva pogotka su greška, ne „uzmi prvi" |
| `ParcelaPripadaKooperantu` | tuđa parcela obara upis |
| `BrutoUnosCuvaBrutoINeto` | `BrutoKg` = tačno uneto, `Kolicina` = izračunat neto |
| `NetoUnosNeIzmisljaBruto` | neto unos ostavlja `BrutoKg` prazno |
| `CenaOverrideJeDozvoljen` | cena različita od cenovnika prolazi; `Cena <= 0` pada |
| `DveKlaseJedanHeader` | dvoklasni blok = **jedan** `OtkupID` + dve stavke |
| `DuplaKlasaPada` | dve stavke iste klase su greška |
| `LosaDrugaStavkaRollback` | pad na drugoj stavci ne ostavlja header ni prvu |
| `PrazanOtkupIDFailClosed` | `NewEntityID` vrati `""` → upis odbijen |
| `PrazanOtkupStavkaIDFailClosed` | isto za `OKS-`, sa header-om već upisanim → rollback |
| `HeaderNeNosiLinePolja` | nov writer ostavlja `Kolicina` / `Cena` / `Klasa` / `KolAmbalaze` / `BrutoKg` / `VozacID` / `Isplaceno` / `DatumIsplate` / `VremeUnosa` **prazne** |
| `KolAmbIzdataJeHeader` | polje je na headeru i preživi oba klasna reda |


**Dopuna iz revizije skele.** Prva mreža je merila ono što je Pre-Flight
odlučio, ali je propustila ono što je writer odlučio **umesto** domena — a to se
vidi tek nad gotovim kodom:

| Test | Tvrdnja |
|---|---|
| `NepoznatKljucUStavciPada` | stavka ima **zatvoren** spisak ključeva; `BruttoKg` ne prolazi kao „nije uneto" |
| `KooperantMoraPostojati` | `KooperantID` je FK, ne string |
| `StanicaMoraPostojati` | `StanicaID` je FK — **i** otkup na nematičnoj stanici prolazi |
| `RedosledKlasaJeKanonski` | ulaz `II, I` daje `I=RB1` — `RedniBroj` nosi dokument, ne redosled poziva |
| `SamoKlasaII` | jednoklasni blok samo druge klase je legitiman dokument |
| `SortaPraznaSamoUzKulturuBezSorte` | prazna sorta prolazi tačno uz kulturu bez sorte; ključ koji fali je greška |
| `TipAmbalazeVezujeSvakaAmbalaza` | obavezan kad ima primljene **ili izdate** ambalaže; bez ambalaže prazan prolazi |

> Tri od njih mere **odsustvo** pooštravanja (`SamoKlasaII`, `SortaPrazna…`,
> `TipAmbalaze…` slučaj (a)). Takav test je lako napisati kao zelen bez sadržaja,
> pa svaki nosi i suprotan slučaj u istom telu — inače bi kapija koja **uvek**
> odbija izgledala isto kao kapija koja radi.

> **Zašto „prazno", a ne „kolone nema".** U skeli te kolone **još postoje** — stari
> writer ih puni i brišu se tek u cutover-u. Tvrdnja `kanonska pozicija = 0`
> (oblik `Test_PR3_OtpremnicaNemaZbirnaID`) postaje moguća **posle** Otkup
> cutover-a, i tada zamenjuje ovu. Do tada bi bila trajno crvena.

Dokaz u oba smera (pokvari → pukne **po imenu** → vrati → zeleno) obavezan za:
`SemaKapija`, `BrojNijeIdentitet`, `ZbirnaClanstvo`, `NovacBezPrimary` — kritične
poslovne invarijante i nov checker.

**Fixture:** `tests/fixtures/otkup_test.xlsm` se regeneriše. Redosled: donor →
jedan prolaz `EnsureAllTables` nad **kopijom** → `tools/make_fixture.py --donor
<kopija>`. Bez toga nove tabele u fixture-u ne postoje i nijedan novi test se ne
može ni napisati.

`run_vba` traži Windows + Excel + pywin32. Iz web sesije se izmena ponašanja
prijavljuje kao **neverifikovana**, nikad kao zelena.

---

### Otpremnica skela (PR5) — pre-flight verdikt

Kapija `pre-flight` je pokrenuta pre ijedne linije koda. **Nije sve zeleno**, i to
je bio smisao:

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | **GAP** | draft-first je *glavni* desktop tok otpremnice (`modOtkupBlok.LinkOtkupIDsToOtpremnica`), a specificiran writer zna samo „izvori → izdato" — v. §9 modela |
| `IDENTITY` | **RISK** | `OtpIdZaBroj` (`modScrDokumenti:502`) razrešava broj u ID preko `LookupValue`, koji vraća **prvi** pogodak (`modDataAccess:608`) |
| `CARDINALITY` | PROVEN | Otkup → Otpremnica N:1 promenljiva; `ReassignOtkupToOtpremnica_TX` (`modDokumenta:5382`) dokazuje premeštanje |
| `INVARIANTS/OWNER` | **GAP** | danas **nijedno** pravilo članstva: reassign proverava samo da cilj postoji i nije storniran |
| `WRITERS` | PROVEN | `row_owner` = `modDokumenta`; 4 schema pisca; 3 produkciona poziva `SaveOtpremnica_TX` (`modAutoHladnjaca:213,258`, `modMasterSync:877`) |
| `DOWNSTREAM` | PROVEN | `Otkup.OtpremnicaID`: **39** ne-test korišćenja, **15** modula, ~~**5 pisača**~~ → **6** (ispravljeno u PR7 pre-flight-u: `modSledljivost` je promašen jer mu je poziv prelomljen u dva reda) → kolona ostaje do PR7 |
| `EVENTS` | PROVEN | fizički: roba napušta otkupno mesto · poslovni: otpremnica nastaje · finansijski: **ne postoji** — `Otpremnica.Cena` je prefill predlog (§13b), ne obračun |
| `CAPABILITY` | N/A | skela je aditivna, nijedna sposobnost se ne seli |
| `PLATFORM` | N/A | nema novog Excel/COM ponašanja |
| `LANDING` | **RISK** | PR4 (#306) još nije merge-ovan; PR5 bi bio stacked nad njim |

`GAP` na `DOMAIN` i `INVARIANTS/OWNER` znači: **nema produkcionog koda** dok se te
dve stvari ne zaključaju. Pravila članstva su zaključana u §4.2a modela. Ostaje
draft — jedina odluka koja menja **oblik API-ja**, pa ne sme da se izabere usput.

#### Postmortem: pre-flight je merio kod, ne već donete odluke

Prva verzija skele otpremnice je zaključala **pogrešan** domenski model: draft
bez stavki, stavke izvedene tek pri izdavanju. Test `DraftNemaStavke` ga je i
učvrstio.

Odgovor je sve vreme stajao u **§13b istog ovog fajla**, u odeljku koji doslovno
kaže da mora biti rešen *pre nego što se napiše* `CreateOtpremnica_TX`:
očekivano su stavke drafta, povezano je `SUM(izvori)`, finalizacija zahteva
jednakost.

Kapija `pre-flight` je odrađena — i vratila je `GAP` na dve ose — ali je merila
**zatečeni kod** (39 čitalaca, dva pisca, `LookupValue` prvi pogodak), a ne
**već donete odluke**. Osa `DOMAIN` je zaključena čitanjem `DOCUMENT_HEADER_LINES`
§4.2 i §9, u kojima tog pravila nema.

> **Pravilo koje iz ovoga sledi:** `DOMAIN` je zatvoren tek kad su pročitana
> **oba** izvora — model dokumenta *i* odeljak plana koji nosi otvorene odluke za
> taj dokument. „Nema toga u modelu" nije dokaz da odluka nije doneta.

Cena greške bila bi vidljiva tek na cutover-u: očekivanje danas živi na
`Otpremnica.Kolicina`, koja u ciljnom modelu odlazi na stavku, pa bi četiri
sposobnosti panela ostale bez izvora (`DOCUMENT_HEADER_LINES` §4.2a).

---

#### Mreža za Otpremnica skelu — imenovano, pre writer-a

| Test | Tvrdnja |
|---|---|
| `JedanBrojJedanHeader` | dvoklasna otpremnica = **jedan** `OtpremnicaID` + dve stavke |
| `StavkeSuIzvedene` | zbir po klasi dolazi iz otkupnih stavki; podmetnute stavke se ne primaju |
| `ClanstvoUIstojTransakciji` | pad pri upisu člana ne ostavlja header |
| `IzvorNeSmeDvaPutaAktivno` | otkup već u aktivnoj otpremnici se odbija |
| `StorniranIzvorNeUlazi` | storniran otkup se odbija |
| `DveStaniceNeProlaze` | izvori sa dve stanice — greška, ne „uzmi prvu" |
| `DveVrsteNeProlaze` | isto za `(VrstaVoca, SortaVoca)` |
| `BezIzvoraNeProlazi` | otpremnica bez ijednog otkupa nije isporuka |
| `VozacSePrima` | vozač dolazi sa headera, izvori o njemu ne govore ništa |
| `PrazanIDFailClosed` | `OPS-` i `OPI-` prazan → upis odbijen, rollback |
| `HeaderNeNosiLinePolja` | `Kolicina` / `KolAmbalaze` / `Klasa` / `BrutoKg` ostaju prazne |
| `OtkupOtpremnicaIDNetaknut` | skela **ne** dira staru kolonu — 39 čitalaca je i dalje na njoj |

Poslednji je jedini te vrste do sada: tvrdi da nova skela **nije** promenila staro
polje. Bez njega bi „aditivno" bila namera, ne mereno svojstvo.

Uz njih idu i četiri koje nosi odluka o draft-u:

| Test | Tvrdnja |
|---|---|
| `DraftNosiOcekivanje` | draft **ima** stavke — ono što je operater prijavio; `povezano` je još 0 |
| `NapredakPoKlasi` | `očekivano / povezano / preostalo` po klasi, kroz dodavanje izvora |
| `IzdavanjeTraziJednakost` | manjak i višak po klasi oba obaraju izdavanje |
| `IzdavanjeRevalidiraIzvore` | `dodaj → storno izvora → izdaj` **pada** (TOCTOU) |
| `BrutoSeNeSabiraParcijalno` | jedan izvor bez bruta → stavka ostaje **bez** bruta |
| `KulturaSeSlaziSaIzvorima` | izvor druge kulture odbijen; draft je zna od otvaranja |
| `DraftNothingOcekivanjePada` | `Nothing` je privatan signal automatskog puta — ručni ulaz ga odbija |
| `UpdateStaniceSaPostojecimIzvoromPada` | izmena zaglavlja ne sme da pokvari već validno članstvo |
| `UpdateKultureSaPostojecimIzvoromPada` | isto za kulturu, uz kontrolu da bezopasna izmena i dalje prolazi |
| `DvaTipaAmbalazeNeUlazeUDraft` | homogen `TipAmbalaze` **pri dodavanju**, ne tek pri izdavanju |
| `NeizdatOtkupNeUlazi` | veza traži `IzdatoStatus = IZDATO`, ne „red koji slučajno ima stavke" |
| `TipAmbalazeJeHeaderCinjenica` | očekivana ambalaža bez tipa odbijena; DRAFT nosi tip **pre** ijednog izvora |
| `IzvorBezGajbiNeOdredjujeTip` | otkup sa `KolAmbIzdata` a bez gajbi na stavkama prolazi uprkos drugom tipu |
| `ClanstvoNaNepostojeciOtkupPada` | read-model **pada**, ne računa manji zbir |
| `DupliParUClanstvuPada` | dupli par je integritet, ne „jedan član" |

> Poslednja dva pišu korupciju mimo writer-a, pa **čiste je pre tvrdnji**:
> korumpiran red truje globalni loader za svaki sledeći test u istom prolazu.
> Oba imaju i kontrolu da read-model posle čišćenja opet radi — inače bi test
> dokazao samo da nešto puca.
| `ClanstvoMutabilnoUDraftu` | dodaj → ukloni → dodaj; uklonjen izvor je **slobodan** za drugu otpremnicu |
| `PosleIzdavanjaClanstvoZamrznuto` | `Dodaj`/`Ukloni`/ponovno izdavanje — sva tri odbijena |
| `StariOtkupNeUlazi` | otkup bez `tblOtkupStavke` (stari pisač) ne može u kanonsku otpremnicu |

> `ClanstvoMutabilnoUDraftu` je jedini koji dokazuje da `DeleteRow` **stvarno**
> briše: da ostavlja tombstone, kapija „već u aktivnoj otpremnici" bi uklonjeni
> izvor i dalje držala, i drugi `Dodaj` bi pao.

---

### Otkup cutover (PR6) — pre-flight verdikt

Prvi dokument kod kog nov pisač postaje **jedini put**. Do sada je sve bilo
aditivno; od ovog PR-a nadalje zeleni golden više ne dokazuje „nisam ništa
pomerio", nego da je ponašanje **namerno** promenjeno tačno tamo gde treba.

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN | §4.1–4.1g zaključani kroz PR4; keš je odlučen (v. ispod) |
| `IDENTITY` | PROVEN | `CreateOtkup_TX` daje jedan `OtkupID` po bloku; `Split(" + ")` gubi razlog postojanja |
| `CARDINALITY` | PROVEN | 1 blok → 1 header → N stavki |
| `INVARIANTS/OWNER` | **GAP → plan** | A11 cilj je `modOtkup` sam; danas **16 upisa u 8 modula** |
| `WRITERS` | PROVEN | **jedan** produkcioni poziv starog pisca: `modOtkupUnos:275` |
| `DOWNSTREAM` | PROVEN | **131** korišćenje kolona koje umiru, ~20 modula (tabela ispod) |
| `EVENTS` | PROVEN | fizički: roba primljena na otkupnom mestu · poslovni: otkupni list nastaje · finansijski: **ne kroz ovaj dokument** (v. ispod) |
| `CAPABILITY` | **treba red** | štampa otkupnog lista, panel blokova, auto-hladnjača — svaka mora završiti kao `MIGRATED`, ne „kod još postoji" |
| `PLATFORM` | N/A | nema novog Excel/COM ponašanja |
| `LANDING` | RISK | stacked nad #307 |

#### Mereno: šta cutover zapravo dira

```
pisac koji se menja      1   modOtkupUnos:275 (jedini produkcioni poziv)
Split(" + ") potrosaci   4   OTKUPNI: modAmbalaza:359, modAutoHladnjaca:182,
                             modOtkupBlok:1439, modPrint:594
                         5   PRIJEMNICA (modDokUnos, modPrint:1090/1359) -- PR9
A11 konsolidacija       16   upisa u 8 modula -> modOtkup API
```

> Prvo merenje je reklo „9" i nije razdvojilo dokumente. Pet od njih parsira
> **`prijemnicaIDs`**, ne `otkupIDs` — oni padaju u PR9, ne ovde.

**Tri od četiri otkupna `Split`-a su bezbolna:** `Split("OTK-1", " + ")` vraća
niz od jednog elementa, pa `modAmbalaza`, `modOtkupBlok` i `modPrint` nastavljaju
da rade i posle prelaska — postaju samo besmisleni, i brišu se kao čišćenje.

#### ⚠ Auto-hladnjača: jedini `Split` koji deli PO KLASI

`modAutoHladnjaca:182` iz `"ID1 + ID2"` vadi `idI` i `idII` i svaku klasu vodi
kroz **svoj** lanac: otpremnica → zbirna → prijemnica → `LinkOtkupRedNaDokument`.
Ugovor je izričit u samom kodu (`:337`):

> „Prazan `OtkupID` NIJE legitimno *nema šta da se veže*… dovde se stiže sa
> ID-jem za svaku aktivnu klasu. Prazno ⇒ prekršen ugovor."

Posle prelaska dvoklasni blok daje **jedan** ID, pa `idII` postaje `""` i noga
Klase II se prijavljuje kao **neuspeh veze** — upozorenje operateru na svakom
dvoklasnom unosu u hladnjaču.

Dublje od toga: lanac pravi **otpremnicu po klasi**, a jedan otkup red može da
drži **jedan** `OtpremnicaID`. Veza je time strukturno gubitna dok otpremnica ne
pređe na header+stavke (PR7).

To nije bug u lancu nego **sudar dva modela** — i pitanje je za operatera, jer
menja opseg PR6. Odluka se upisuje ovde pre nego što korak 2 nastavi.

| Kolona koja umire | Ne-test korišćenja | Modula |
|---|---:|---:|
| `Kolicina` | 47 | 20 |
| `Cena` | 33 | 14 |
| `KolAmbalaze` | 24 | 11 |
| `Klasa` | 20 | 15 |
| `BrutoKg` | 7 | 5 |
| **ukupno** | **131** | |

#### Keš NE ulazi u pisca — i to je merenje, ne pretpostavka

Red 6 u tabeli PR-ova kaže „novac na header", što se lako čita kao „`CreateOtkup_TX`
mora da piše `tblNovac`". Model kaže suprotno, i to je već odlučeno: §4.1b briše
`Novac` / `PrimalacNovca`, a §6.1 kaže da **keš uopšte ne ulazi kroz otkupni list** —
put koji stvarno postavlja `Isplaceno` je avans. Golden `B1`/`B4` su ranije uklonjeni
baš iz tog razloga.

Legacy `SaveOtkupMulti_TX` ipak snapshot-uje `tblNovac` i ima granu za keš — **mrtav
kod**, i briše se zajedno sa pisačem.

> Ovo je prvi put da je pravilo iz §13c radilo unapred: da su pročitane samo tabela
> PR-ova ili samo model, PR6 bi dobio pisca koji knjiži keš i test koji taj mrtav
> kod čuva.

Ostaje samo **ambalaža**: `TrackAmbalaza` po klasi za primljene gajbe i obrnut smer
za `KolAmbIzdata` (`modOtkup:1294-1310`).

#### Redosled unutar PR-a

Jedan PR (odluka operatera), ali commit-i idu ovim redom — svaki je celina koja se
može čitati zasebno:

```
1  pisac kompletan     ambalaza + StornoOtkup_TX nad headerom + ispravka
                       jos ADITIVNO: golden 12/0 mora ostati nepromenjen
2  jedini put          modOtkupUnos:275 -> CreateOtkup_TX
                       SaveOtkupMulti_TX obrisan, 9x Split(" + ") pada
                       OVDE golden SME da se promeni -- i mora se objasniti
3  citaoci             131 koriscenje -> read-model nad tblOtkupStavke
4  kanon               kolone van schema.json; tvrdnja postaje "kolone nema"
5  A11                 8 pisaca -> modOtkup API, ratchet na cilj
```

Korak 2 je jedini koji menja ponašanje bez mreže ispod sebe — zato korak 1 mora
biti zelen i dokazan **pre** njega.

#### ⚠ `SaveOtkupMulti_TX` se NE može obrisati u PR6

Korak 2 je predviđao brisanje starog pisca, a korak 5 pun ugovor za
`VrednostOtkupa` — koji bez toga ne može, jer je stari pisac poslednji
proizvođač otkupa **bez stavki**. Migracija je pokušana i **izmerena**.

Devet fixture poziva se prevodi mehanički. Ali od 12 tvrdnji koje posle toga
padnu, **šest nije mehaničko** — one opisuju živo ponašanje koje tek PR7 menja:

```
Test_FullDocumentChainHappyPath
    Otkup(klasa I).OtpremnicaID  = otpI
    Otkup(klasa II).OtpremnicaID = otpII
```

Dva otkup reda, svaki na **svojoj** otpremnici. Jedan header drži **jedan**
`OtpremnicaID`, pa je tvrdnja strukturno nemoguća dok otpremnica ne pređe na
header+stavke i dok pripadnost ne postane `tblOtpremnicaIzvori`.

Isto važi za hladnjački lanac (`RunHladnjacaChain`,
`Test_HladnjacaChainHappyPath`, `Test_HladnjacaChainLinkFailureIsReported`):
lanac je u PR6 **pauziran**, pa testovi koji tvrde da se kompletira ne mogu da
prođu — oni mere sposobnost koja je namerno isključena.

| Tvrdnja koja pada | Zašto |
|---|---|
| `Otkup class I/II linked to matching otpremnica` | per-class veza — **PR7** |
| `TraceByZbirna returns rows` | posledica gornje |
| `Hladnjaca lanac ... NEPOTPUN` ×2 | lanac pauziran — **PR7** |
| `Hladnjaca lanac: otkup red povezan sa otpremnicom` | per-class veza — **PR7** |
| `Hladnjaca lanac: otkup red nosi BrojZbirne` | lanac pauziran — **PR7** |

Preostalih šest (AutoLink fixture, `FindOtkupIDByBrojAndKlasa`) jesu mehaničke.

**Odluka: stari pisac ostaje do PR7.** Alternativa bi bila ugasiti šest tvrdnji
koje pokrivaju živo ponašanje — a test koji se tiho isključi je gori od testa
koji nedostaje, jer izgleda kao pokriće.

Posledica: `VrednostOtkupa` ostaje bez punog ugovora do istog trenutka, i to je
imenovano u samom kodu, ispod funkcije.

> Pokušaj je vraćen u celini (`git checkout`), a ne ostavljen polovičan. Grana
> je posle toga ponovo zelena: `BusinessFlowPro 945/0`.

---

#### Mreža za Otkup cutover — imenovano, pre writer-a

| Test | Tvrdnja |
|---|---|
| `AmbalazaIdeNaHeader` | primljene gajbe se knjiže po klasi, izdate obrnutim smerom; zbir odgovara stavkama |
| `StornoJednimID` | dvoklasni blok se stornira **jednim** pozivom nad `OtkupID`, ne dva puta po klasi |
| `IspravkaJeNovaVerzija` | korekcija pravi nov `OtkupID` + `IspravkaOdID` + `ZamenjenSaID`, stara ostaje storniran fakt (A13) |
| `NovacBezPrimary` | vrednost = `SUM(stavke)`; **nema** persistentnog `Isplaceno`; keš ne ulazi kroz otkupni list |
| `JedanIDBezSplita` | `CreateOtkup_TX` vraća jedan ID; nijedan potrošač ne parsira `" + "` |
| `KoloneNema` | `kanonska pozicija = 0` za `Kolicina`/`Cena`/`Klasa`/`KolAmbalaze`/`BrutoKg` — zamenjuje „nov writer ih ostavlja prazne" iz PR4 |
| `PanelCitaStavke` | „Ostatak", prekoračenje i sažetak čitaju stavke, ne header |
| `StampaCitaStavke` | otkupni list štampa klase iz `tblOtkupStavke` |
| `AutoHladnjacaJedanBlok` | auto-lanac dobija jedan `OtkupID` i ne deli po klasi |

Dokaz u oba smera obavezan za `StornoJednimID`, `IspravkaJeNovaVerzija`,
`NovacBezPrimary` i `KoloneNema` — sve četiri su kritične poslovne invarijante ili
menjaju premisu zatečenog testa.

#### Golden mreža prestaje da bude „nepromenjena"

Do sada je `RunGoldenSuite 12/0 nepromenjen` bio dokaz da skela ništa nije pomerila.
U koraku 2 to više ne važi: scenariji koji prolaze kroz otkupni list **moraju** da
se promene, jer se menja broj redova po bloku. Svaka promena goldena mora da nosi
obrazloženje u commit-u; golden koji se promenio bez objašnjenja je regresija koja
je prošla kao osvežavanje.

---

### Otpremnica cutover (PR7) — pre-flight verdikt

> **Ponovo izmeren 15.09.2026 — v. §14.7.** Tabele ispod su stanje od 12.09:
> linije su se pomerile, LANDING rizik je otpao (#308 je merge-ovan), a ugovor i
> capability mapa imaju rupe koje §14.7 imenuje. Granica PR7/PR8 i ugovor
> odlučeni su istog dana — v. §14.7, „Odluke operatera“.

Merenja pre ijedne linije koda, po `.claude/skills/pre-flight`. Rađeno **dok PR6
čeka merge** — spec ne zavisi od ishoda njegovog review-a.

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | **PROVEN** | §4.2 (grain: jedna isporuka sa otkupnog mesta), §4.2a/b, §3.1 (pripadnost nije kolona), A15 (verzionisano članstvo) |
| `EVENTS` | **PROVEN** | v. dole |
| `IDENTITY` | **PROVEN** | `OtpremnicaID` opaque; broj je labela scoped po stanici; članstvo u `tblOtpremnicaIzvori` |
| `CARDINALITY` | **PROVEN** | 1 otpremnica ← N otkupa; 1 otkup → najviše **jedno aktivno** članstvo (A15); istorijski više |
| `INVARIANTS/OWNER` | **PROVEN** | `AktivnoOtpClanstvoPoKanonu` diže grešku na dva aktivna zapisa (`modDokumenta:3417`); `IzdajOtpremnicu_TX` revalidira izvore |
| `WRITERS` | **PROVEN** | v. dole — i tu je prvi nalaz |
| `DOWNSTREAM` | **PROVEN** | 141 pojava četiri kolone koje odlaze (§14.6), 16 produkcionih modula čita `Otkup.OtpremnicaID` |
| `CAPABILITY` | **PROVEN** | tri pauzirane sposobnosti + AutoLink + GlobalGAP; v. mapu |
| `ACCEPTANCE CONTRACT` | v. dole | plan dokaza, ne dokaz |
| `PLATFORM` | N/A | nema Excel/COM nepoznanice; sve je nad `ListObject`-ima koji već rade |
| `LANDING` | **RISK** | zavisi od merge-a #308; v. dole |

#### ⚠ NALAZ 1: plan kaže 5 pisača, ima ih 6

Red 7 tabele PR-ova glasi „**briše `Otkup.OtpremnicaID`** sa svih 5 pisača".
Mereno — šest:

```
modAutoHladnjaca:385   modDokumenta:6883   modMasterSync:2598
modOtkupBlok:1480      modSledljivost:255  modStornoFlow:2569
```

Šesti (`modSledljivost`, AutoLink) je promašen jer mu je poziv **prelomljen u dva
reda** — tačno slepa mrlja koju §13a već opisuje za `AppendRow`. Ironija je
potpuna: promašen je baš onaj pisač koji ceo PR7 treba da ukine, jer AutoLink je
heuristika koju eksplicitno članstvo zamenjuje.

> Pouka za ubuduće: brojevi u planu se mere skriptom, ne grep-om po jednom redu.

#### ⚠ NALAZ 2: otpremnica ima ISTI oblik pisca koji je otkup upravo izgubio

```
modDokUnos:262  ->  SaveOtpremnicaMulti_TX  ->  SaveOtpremnica x2 (po klasi)
                                            ->  vraca "ID1 + ID2"
```

To je linija-po-liniju isti obrazac kao obrisani `SaveOtkupMulti_TX`. Znači i isti
posao, i ista zamka: pozivalac koji string parsira. Uz njega `SaveOtpremnica_TX`
(jednoklasni) ima **3 produkciona pozivaoca** — `modAutoHladnjaca` ×2 (pauziran) i
`modMasterSync` ×1 — i **41 test poziva**.

Ukupno: **44 mesta**, od kojih su 4 produkciona. Isti razred posla kao korak 3 u
PR6, samo veći test rep.

#### ⚠ NALAZ 3: tri `Split(" + ")` nad OTKUP ID-evima su već mrtva

`CreateOtkup_TX` vraća **jedan** ID od PR6, pa ovi više nikad ne cepaju ništa:

```
modAmbalaza:359      Split(blockOtkupIDs, " + ")
modAutoHladnjaca:190 Split(otkupIDs, " + ")
modOtkupBlok:1460    Split(otkupIDs, " + ")
```

Nisu bug (jedan element = ceo string), ali su **mrtav aparat koji izgleda živ**.
Brišu se u PR7 uz svoje pozivaoce; preostalih pet (`modDokUnos` ×3, `modPrint` ×2)
je prijemnica, dakle PR9.

#### BUSINESS EVENTS

| Događaj | Kada | Šta nastaje | Šta može bez sledećeg |
|---|---|---|---|
| **fizički** | roba napušta otkupno mesto i ulazi u vozilo | ništa u bazi po sebi | otprema bez papira ne sme postojati |
| **poslovni** | otpremnica prelazi u `IZDATO` | dokument sa sastavom; otkupi postaju njeni članovi | **DRAFT sme da stoji** koliko treba — sastav se gradi postepeno |
| **finansijski** | **nijedan** | — | otpremnica **ne stvara ni dug ni potraživanje** |

Treći red je važan i lako se previdi: kooperant je plaćen po **otkupu**, kupac
plaća po **fakturi**. Otpremnica je isključivo dokument kretanja robe. Zato PR7
**ne sme** da dira `tblNovac` — ako se pojavi potreba, to je `DOMAIN GAP`, ne
implementacioni detalj.

Datum otpremnice je datum **fizičke otpreme**, ne dan unosa i ne dan izdavanja.

#### CAPABILITY MAP — tri pauze koje PR7 mora da podigne

> **Odluka 15.09.2026 (§14.7):** auto-lanac hladnjače ostaje `PAUZIRAN` do PR8
> (ne `MIGRATED`); GlobalGAP `REPLACED` prelazi u PR8; ručna zbirna F3 dobija
> imenovanu pauzu do PR8.

| Sposobnost | Stanje danas | PR7 |
|---|---|---|
| Panel napretka bloka | `NapredakBlokaDostupan() = False` (`modOtkupBlok:1572`), tri gejta | `MIGRATED` na `GetOtpremnicaProgress` |
| Auto-lanac hladnjače | poziv ugašen u `modOtkupUnos:388`, kod netaknut | `MIGRATED` — lanac prima jedan `OtkupID` i pravi članstvo |
| Wiring correction API-ja | `IspravkaOtkupa_TX` bez produkcionog pozivaoca | `MIGRATED` — stari `OtkupID` mora da otputuje kroz prefill |
| `AutoLinkOtkupOtpremnica` | povezuje 0 | **`INTENTIONALLY REMOVED`** — heuristika koju članstvo zamenjuje |
| GlobalGAP sledljivost | `TraceByZbirna` vraća `"NEMA"` za nov otkup | `REPLACED` — čita članstvo, ne `Otkup.OtpremnicaID` |

Četvrti red je jedini `INTENTIONALLY REMOVED` i traži izričitu potvrdu: dugme
„Auto-poveži" nestaje sa ekrana sledljivosti, jer povezivanje prestaje da bude
pogađanje.

#### ACCEPTANCE CONTRACT — plan dokaza

> **Odluka 15.09.2026 (§14.7) menja ovaj plan dokaza:** §14.2 tvrdnje 6, 7 i
> zbirni deo 3 prelaze u PR8, H2 takođe; golden ostaje na test-only pisaču
> (cilj: 12/0 bez promene snapshot-a); A11 za `tblOtkup` je 8 → 1.

**Šta će važiti:** svih **7** tvrdnji iz PR7 acceptance mreže (§14.2) zeleno po
imenu · `tblOtpremnicaIzvori` pokazuje na prave `OtkupID`-eve · `Otkup.OtpremnicaID`
i `Otkup.BrojZbirne` **obrisani** iz kanona · `Cena → PredlogCena` sa svih 8
čitalaca · A11: pisaca nad `tblOtkup` **9 → 1**.

**Šta mora ostati netaknuto:** golden mreža 12/0 sa **nepromenjenim** snapshot-ima
· `CreateOtkup_TX` ponašanje · storno kaskada · A13 kapija iz PR6.

**Edge koji mora proći:** H2 (`CorrectionSestre`) — `OTP2` i `OTP3` pripadaju
sastavu **i** stare i nove zbirne. Ako to ne prođe bez gubitka istorije, model
veze nije dovoljan i to je ceo razlog zbog kog tabela članstva postoji.

**Šta mora biti odbijeno:** dva aktivna članstva za isti otkup · izmena sastava
`IZDATE` otpremnice · otkup bez stavki kao izvor.

**Čime se dokazuje:** `RunBusinessFlowProSuite` + `RunGoldenSuite`; sabotaža nad
svakom novom kapijom; `who_writes --check-ownership` kao brojčani dokaz za 9 → 1.

#### ⚠ PR7 NE SME SAMO DA OTKLJUČA ZAJEDNIČKU KAPIJU

PR6 je ceo izvedeni lanac stavio iza **jedne** kapije:

```vb
Public Function IzvedeniLanacIzPwaDostupan() As Boolean
    IzvedeniLanacIzPwaDostupan = False
End Function
```

Za PR6 je to ispravno — maksimalno fail-closed, jer sva tri legacy ulaza pišu
nazad na zaglavlje otkupa. **Za PR7 nije.** Kapija pokriva tri koraka koji se
oslobađaju u **dva različita PR-a**:

| Korak | Oslobađa |
|---|---|
| auto-Otpremnica iz PWA | **PR7** |
| Malina auto-Zbirna iz Otpremnice | **PR8** |
| VOZ/Zbirna import + legacy backlink | **PR8** |

Ako PR7 samo napiše `= True`, zajedno sa novom otpremnicom se **istog trenutka
vraćaju i dva legacy zbirna koraka** koji i dalje pišu `Otkup.BrojZbirne` — i
kontaminacija koju je PR6 upravo zatvorio se vraća na mala vrata.

**Obavezno u PR7:** kapija se **deli**, ne otključava.

```
PR7:   PwaOtpremnicaDostupna = True
       PwaZbirnaDostupna     = False

PR8:   PwaZbirnaDostupna     = True
```

Svaki od tri ulaza tada gleda **svoju** kapiju. Test `Test_PWA_IzvedeniLanacJePauziran`
se u PR7 razdvaja na isti način: otpremnica prolazi, dva zbirna ulaza i dalje
padaju po imenu.

#### Ugovorna nijansa CRID poređenja (zabeleženo, ne otvoreno)

`PwaIstiSadrzaj` sada poredi `BrojDokumenta` **uslovno**: učestvuje samo kad ga
PWA izričito pošalje, jer prazan incoming broj znači da ga je master generisao
lokalno.

Time se hvata `77 → 78` (konflikt), ali **ne** i `77 → prazno`: taj slučaj se iz
trenutnog zaglavlja ne može razlikovati od „broj je oduvek bio prazan pa ga je
master dodelio". Za doslovan ugovor „isti payload" trebalo bi pamtiti **da li je
klijent poslao broj**, a to danas nigde ne stoji.

**Ne otvara se u PR6.** Ako PWA posle rollout-a garantuje da jednom dodeljen broj
ne nestaje iz istog `ClientRecordID`, trenutni model je praktično dovoljan. Ovde
stoji da se kasnije ne bi mislilo da se iz zaglavlja može dokazati nešto što ne
može.

#### LANDING RISK

PR7 dira `modDokumenta`, `modOtkupBlok`, `modAutoHladnjaca`, `modSledljivost` i
`modDokUnos` — **sve fajlove koje PR6 već menja**. Grana se zato otvara tek kad se
#308 merge-uje; rad nad njegovom granom bi bio stacked PR koji propada ako se bazni
merge-uje prvi (poznata zamka).

Do tada je PR7 **spec-only** — ovaj odeljak.

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

| Dokument | Tabela u TX | Produkcionih pisaca | Skela | **Cutover** |
|---|---|---|---|---|
| Zbirna | 1 | 3 | **1.** ✅ | **3.** |
| Otkup | 3 | 9 | 2. | **1.** |
| Otpremnica | 2 | 1 | 3. | **2.** |
| Prijemnica | 6 | 4 | 4. | 4. |

**Skela i cutover više nisu isti redosled**, i to je posledica A15.

Skela (nove tabele + writer, aditivno) sme bilo kojim redom — Zbirna je bila
prva jer ima najmanju transakciju, i to je i dalje bio dobar izbor.

**Cutover** ide **uzvodno-nadole**: dokument sme da postane jedini put tek kad
svi dokumenti na koje pokazuje po ID-u imaju svoj header identitet. `tblZbirnaIzvori`
pokazuje na `OtpremnicaID`, a otpremnica danas nema jedan identitet — v. „Zašto
je redosled promenjen posle PR3" ispod tabele PR-ova.

Prijemnica ostaje poslednja: šest tabela, i njene stavke hrane fakturu.

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
| 3 | ✅ **Zbirna header+stavke**: `tblZbirnaStavke`, **`tblZbirnaIzvori`**, `CreateZbirna_TX` / `CreateZbirnaIzIzvora_TX` — stavke se **izvode iz izvornih otpremnica**, membership ide u **istoj** transakciji, opaque `ZbirnaID` / `ZbirnaStavkaID` / `ZbirnaIzvorID` svi fail-closed; **`tblZbirnaIzvori`** nosi verzionisano članstvo (A15). **Aditivno** — stari pisač je i dalje jedini put, golden 12/0 nepromenjen. Uz to: A11 kapija je bila slepa na funkcijski i na prelomljen oblik `AppendRow` (v. §13a) | 2 |
| 4 | ✅ **Otkup header+stavke** (skela): `tblOtkupStavke`, `CreateOtkup_TX(h, stavke, outGreska)`, opaque `OtkupID` po **bloku**, ne po klasi. Target šema po §4.1c–f: bez `VozacID` / `Isplaceno` / `DatumIsplate` / `VremeUnosa`; `KulturaID` prima, ne razrešava. **Bez PWA adaptera** — v. napomenu ispod | 3 · **spec zaključan** |
| 5 | ✅ **Otpremnica header+stavke** (skela): `tblOtpremnicaStavke`, **`tblOtpremnicaIzvori`**, **sedam ulaza** — `CreateOtpremnicaDraft_TX(h, očekivano)` / `Update` / `Dodaj` / `Ukloni` / `GetOtpremnicaProgress` / `IzdajOtpremnicu_TX` + jednopotezni `CreateOtpremnicaIzIzvora_TX`. **Stavke drafta su očekivanje** (§13b), izdavanje traži `očekivano = povezano` i revalidira izvore. Otpremnica ima **persistentan `DRAFT`**, za razliku od otkupa. Uz to: prvi **meren** put brisanja reda (`DeleteRow` + A11 kapija) | 4 · **spec zaključan** |
| 6 | ✅ **Otkup cutover + integracije** (PR #308, merge 12.09.2026): ambalaža i novac na header, `Isplaceno` **izvedeno pa obrisano**, storno, ispravka (A9) + A13 kapija, print, PWA ingest. Nov pisač je jedini put. Auto-hladnjača, panel bloka i **PWA auto-otpremnica** pauzirani do 7; reader sweep izmeren i podeljen (§14.6) | 5 |
| — | ✅ **KAPIJA ODLUKE — ZATVORENA 13.09.2026: nastavak u mestu** (u mestu 3 · novo stablo 0 · nejasno 3; kriterijumi zamenjeni merljivima) — v. §14.1 | 6 |
| 7 | 🟡 **pre-flight 15.09 (§14.7) — granica odlučena (A): zbirni tokovi i auto-lanac hladnjače pauzirani do PR8, `SaveOtpremnica*` samo za testove. Pre koda: mali PR za kvarove 2/3/9 (✅ #334), ponovljen popis sa proverom (🟡 7b 16.09: premeren, nezavisno 6/12 celina), odluke o F2.** **Otpremnica cutover**: `tblOtpremnicaIzvori` pokazuje na prave `OtkupID`-eve; propagacija ispravke naniže; panel prelazi na `GetOtpremnicaProgress`; **briše `Otkup.OtpremnicaID`** sa svih **6** pisača (ne 5 — v. PR7 pre-flight, NALAZ 1); **rename `Cena` → `PredlogCena`** sa čitaocima (§13b) | 6 |
| 8 | **Zbirna cutover**: invarijanta preko `tblZbirnaIzvori` (sada nad **pravim** `OtpremnicaID`-evima), `StornoZbirna_TX(id)`, storno otpremnice po §7.1, **propagacija ispravke = nova verzija (A13)**, print, izveštaji. **Briše `ZbirnaIdent*`, `ZbirnaGeneracija*` i mrtvu `RunSimpleStornoOtpremnica`.** Registruje goldene D1, H1, H2. **Iz PR7 preuzima (odluka 15.09, §14.7):** §14.2 tvrdnje 6, 7 i zbirni deo 3, edge H2, podizanje pauze zbirnih tokova (F3, malina, VOZ) i auto-lanca hladnjače, brisanje test-only `SaveOtpremnica*` i po-klasnih kolona `tblOtpremnica`, izmenu golden scenarija A4 | 7 · **§7.1, A13–A15 odlučeni** |
| 9 | **Prijemnica** header+stavke + izvori + cutover | 8 |
| 10 | **Faktura**: `FakturaStavka.PrijemnicaStavkaID` | 9 |
| 11 | **Paleta**: `PaletaStavka.PrijemnicaStavkaID` | 9 |
| 12 | **Sledljivost kao graf** nad eksplicitnim FK; ukloniti heuristički AutoLink | 11 |
| 13 | **E2E + brisanje**: `COL_GENERACIJA_ID`, `COL_DETE_ZBIRNA_GEN`, `*ByBroj_TX`, svih 10 `Split(" + ")`, mrtvi testovi i sabotaže; pravila `NEMA_GENERACIJE` / `NEMA_BROJA_KAO_FK` / `NEMA_ID_PLUS_ID`; `ZBR_IDENTITET.md` → superseded | 12 |
| — | `CLAUDE.md` §3 (obrtanje pravila o izvoru istine šeme) | zaseban process PR |

Završni korak (red 13) je ključan i **ne sme se preskočiti**: dok `GeneracijaID` postoji kao živ
runtime mehanizam, refaktor nije završen — to je dual identity model, gori od
sadašnjeg.

#### PWA adapter nije u skeli — i to je izbor, ne previd

Roadmap je kratko nosio `ImportOtkupPWA_TX` u redu 4, dok `DOCUMENT_HEADER_LINES`
§7 i backlog istovremeno kažu da je PWA ingest van opsega. Tri mesta, dve
tvrdnje.

**Odluka: skela nosi samo `CreateOtkup_TX`.** PWA adapter ima smisla tek kad se
zaista testira idempotency (isti `ClientRecordID` dvaput → jedan header, jedna
stavka) i kad stvarno **zamenjuje** `modMasterSync`, a ne stoji pored njega.
Polovičan omotač koji ne zamenjuje ništa je treći put do istog upisa.

Ide u **Otkup cutover**, zajedno sa uklanjanjem `modMasterSync`-ovog direktnog
upisa. Tada i test `PWAReimport` postaje merljiv.

### Zašto je redosled promenjen posle PR3

Prvobitni plan je išao **Zbirna → Otkup → Otpremnica**, jer je zbirna najmanja
transakcija. Verzionisano članstvo (A15) je to obesmislilo:

> `tblZbirnaIzvori` beleži `ZbirnaID → OtpremnicaID`. Ali otpremnica danas **nema
> jedan identitet** — `SaveOtpremnicaMulti_TX` pravi red po klasi i vraća
> `"OTP-1 + OTP-2"`. Kanonsko članstvo bi time zapisivalo *koji fizički klasni
> redovi* čine zbirnu, a A15 traži *koje verzije poslovnih otpremnica*.

To nisu iste stvari. Zato važi pravilo koje generiše redosled:

> **Dokument sme u cutover tek kad svi dokumenti na koje pokazuje po ID-u imaju
> svoj header identitet.**

Skele (nove tabele + writer, aditivno) smeju bilo kojim redom — PR3 je to i
uradio. **Cutover** ide uzvodno-nadole: `Otkup → Otpremnica → Zbirna`.

Cena promene je nula danas: `tblZbirnaIzvori` još nema nijednog produkcionog
pisca, pa nema ni jednog reda sa pogrešnim grain-om. Da je Zbirna otišla u
cutover pre Otpremnice, kanonsko članstvo bi se punilo identitetom za koji već
znamo da je pogrešan.

### 13b) Dve odluke pre Otpremnice

Obe moraju biti rešene pre nego što se napiše `CreateOtpremnica_TX`.

#### „Očekivano" nema svoju tabelu — to su stavke drafta

Realan tok je: operater prvo otvori otpremnicu, pa unosi otkupne listove pod
njom, gledajući `očekivano / povezano / preostalo`. Gde živi „očekivano":

```
DRAFT OtpremnicaStavke        ono sto je operater UNEO (ocekuje)
SUM(tblOtpremnicaIzvori -> OtkupStavke)   ono sto je POVEZANO
preostalo = ocekivano - povezano

FINALIZE: zahteva  ocekivano = povezano
IZDATO:   stavke se zamrzavaju (A13)
```

**Bez dodatne `Expected` tabele.** Stavke drafta *jesu* očekivanje; pri
finalizaciji prestaju to da budu i postaju sadržaj verzije. To je isti prelaz
koji A5/A13 već opisuju („keš dok je draft, činjenica kad je izdato"), samo
gledan sa ulazne strane.

#### `Cena` ne ide na `OtpremnicaStavke`

Pitanje: ako jedna otpremnica sabira pet otkupnih listova iste klase sa
**različitim cenama**, šta znači jedna `Cena` na stavci?

Mereno u zatečenom kodu — `Otpremnica.Cena` danas **nije agregat**, nego
**seed za prefill blokova**: `modOtkupBlok.bas:688` i `modScrDokumenti.bas:1067`
prvo pitaju `ExistingBlokCena(otpID)`, pa tek ako je 0 padaju na `Otpremnica.Cena`.
Uz to je čitaju izveštaji i štampa.

**Odluka:**

| | |
|---|---|
| `OtpremnicaStavka.Cena` | **ne postoji.** Vrednost dokumenta je `SUM(izvorne otkupne stavke)` |
| `Otpremnica.Cena` (header) | ostaje, ali **preimenovana u `PredlogCena`** — predlog cene za blokove, izričito **ne-finansijsko polje** |
| zabrana | nigde se vrednost otpremnice ne računa kao `Kolicina × Cena` |

Poslednja tačka nije teorijska: `modDokumenta.CalculateManjakByOtpremnica`
(`:3618`) danas čita **i** `Kolicina` **i** `Cena` sa otpremnice. Posle refaktora
bi ta cifra mogla da se ne slaže sa zbirom izvornih otkupa. Otpremnica cutover
mora da odluči iz čega se `manjak` računa — iz izvora, ne iz denormalizovanog
para.

Da je `Cena` prosto nasleđena „jer je legacy `SaveOtpremnica` ima", model bi
dobio drugi izvor istine za novac.

#### `ocekivano` ne sme da ostane `Optional`

Danas: `Optional ByVal ocekivano As Collection`, a `RequireOcekivanoSeSlaze`
izlazi na `Nothing`. To znači da se kontrola **može isključiti time što se ne
prosledi** — a poslovno pravilo je da ono što je operater otkucao mora da se
poredi sa izvedenim.

Istovremeno postoje legitimni automatski tokovi (auto-hladnjača, malina) gde
nezavisnog ručnog očekivanja **nema**.

**Odluka:** razdvojiti namere u dva javna ulaza nad istim `Private` core-om.

```
CreateZbirna_TX            ocekivano OBAVEZNO   (operater unosi papirnu zbirnu)
CreateZbirnaIzIzvora_TX    bez ocekivanog       (izricito "derived-only")
```

Ne dva writera — dva **potpisa** koji izražavaju nameru. Kontrola se time ne
može isključiti slučajno; može se samo odabrati drugi ulaz, i to se vidi na
callsite-u.

**Zašto tek u cutover-u, a ne sada:** skela nema nijednog produkcionog
pozivaoca, pa danas ništa ne može da je isključi. Kad se pojavi prvi, potpis
mora već biti podeljen.

#### Pre prvog mutable-DRAFT članstva: A11 mora da meri i brisanje

A14/A15 kažu da je članstvo promenljivo dok je dokument `DRAFT`. Uklanjanje
izvora iz drafta znači **fizičko brisanje reda** članstva.

Ali `who_writes.py` meri `AppendRow` / `UpdateCell` / `RequireUpdateCell` — **ne
i brisanje**. Kod koji radi `lo.ListRows(i).Delete` bio bi A11 kapiji nevidljiv,
isto kao što su ranije bili funkcijski i prelomljeni `AppendRow`.

> To bi bila **treća** pojava iste klase rupe. Prve dve su nađene slučajno.

**Odluka:** pre nego što se napiše prvi API koji menja članstvo drafta, moraju
postojati **oba**:

1. kanonski `Delete`/membership API (brisanje ne ide direktno po `ListRows`);
2. `who_writes.py` koji brisanje meri kao mutaciju, sa slučajevima u
   `--self-test` — i za `lo.ListRows(i).Delete` i za novi API.

#### Referencijalni integritet članstva (P2)

`AktivnoClanstvoPoKanonu` proverava da otpremnica nema dva aktivna zapisa, ali
**ne** proverava da `ZbirnaID` iz zapisa zaista postoji u `tblZbirna`. Orphan
zapis se time tretira kao aktivan.

Ishod je fail-closed — takva otpremnica se ne može ponovo upotrebiti — pa nije
blokada za skelu. Ali pre cutover-a mora u health/invariant mrežu:

```
za svaki red tblZbirnaIzvori:
    ZbirnaID     postoji TACNO jednom u tblZbirna
    OtpremnicaID postoji TACNO jednom u tblOtpremnica
```

Isto važi za `tblOtpremnicaIzvori` kad nastane.

#### Correction polja u šemi se moraju preimenovati, ne pretumačiti

A9 govori o `IspravkaOdID` / `ZamenjenSaID`, a kanon (`schema/schema.json`) i
dalje fizički nosi `IspravkaOd` / `ZamenjenSa` / `CorrectionID`.

Dok stari writer koristi zatečeni model to se ne dira. Ali cutover mora ta polja
**stvarno preimenovati** — nije dovoljno reći „ovo sad znači ID". Kolona koja se
zove `IspravkaOd` a nosi ID je tačno vrsta dvosmislenosti koju refaktor uklanja.

---

### 14.1) Kapija odluke posle Otkup cutover-a

> **ZATVORENA 13.09.2026. Odluka: NASTAVAK U MESTU.** Mereno, ne procenjeno —
> svih šest kriterijuma, jedan po jedan, sa komandama. Rezultat:
> **u mestu 3 · novo stablo 0 · nejasno 3.**
>
> Nijedan kriterijum ne pokazuje na novo stablo, i **nijedan od pet okidača** iz
> liste „Prelazimo na novo stablo ako" nije izmeren: nema compatibility facade-a,
> nema dual-write putanje, nema paralelnog života dva modela, test corpus je
> upotrebljiv, jezgro jeste izolovano od stare šeme.

#### Šta je odlučilo

| Kriterijum | Izmereno | Verdikt |
|---|---|---|
| Pisaca nad `tblOtkup` | 8, ne 1 — **ali presek kolona je prazan skup**: `modOtkup` piše 16 kolona sadržaja, ostalih sedam piše tačno 5, sve denormalizovane veze ka tuđem dokumentu. Nula upisa u `Kolicina/Cena/Klasa/Kooperant/Datum/Stanica` van legacy `SaveOtkup`. Tri od sedam su strukturno mrtva na nov dokument. | **u mestu** |
| Resolver/fallback u jezgru | **0.** Naivan grep daje 6 — svih 6 su komentari koji objašnjavaju odsustvo. Resolver živi u UI/adapter sloju i fail-closed je (`If koliko <> 1 Then` → greška). | **u mestu** |
| Dual model | **nema ga.** Tri upisa u `tblOtkup` ukupno; `SaveOtkup` piše stari plosnat red ali je **write-dead** (nijedan produkcioni pozivalac — postoji da test može napraviti zaglavlje *bez* stavki, oblik koji nove kapije moraju da odbiju). Nov pisac stare kolone ostavlja **prazne**, ne ogleda ih — §14.6 to izričito traži. | **u mestu** |
| Linije u jezgru | „core" nije definisan: dva razumna čitanja daju **−41%** i **+16%**. | nejasno |
| Identitetski testovi | prefiks `T_ZBR*` **ne postoji u repou** (0 pogodaka na svakom sidru); meri se nad Zbirna lancem, a presečen je Otkup. | nejasno |
| Reuse pozivalaca | 10 : 5 nad *dodirnutim* modulima, ali **14 nedirnutih čita kolone koje nov pisac nikad ne puni** — diff od nule broji kao „savršen reuse". | nejasno |

Najjači pojedinačni dokaz nije broj nego **mehanizam**: prelaz 9 → 8 pisaca nije
nastao prepisivanjem nego brisanjem kolona `Isplaceno` i `DatumIsplate` iz kanona.
**Pisac nestaje sa kolonom.** Isti potez nad preostalih 5 vezivnih kolona vodi
8 → 1 mehanički, bez ijednog novog sloja — a taj put je već jednom pređen, na ovoj
istoj tabeli, u ovom istom ciklusu.

#### Šta merenje NIJE potvrdilo, i mora se reći

Kriterijumi su prikazivali refaktor **gotovijim nego što jeste.** Upis je presečen,
**čitaoci nisu prešli**:

- `modNovac.IsplataBlokProblem` računa vrednost bloka kao `Kolicina × Cena` sa
  **zaglavlja** — a te kolone su na nov dokument prazne. Posledica: `preostalo`
  ispada 0 i **svaka isplata na nov otkup se odbija**. Zovu ga dva živa mesta
  (`SaveOMUlaz_TX`, ekran novca).
- 14 nedirnutih modula čita kolone koje nov pisac ne puni: `modBankaMapiranje`
  (uplata se knjiži kao avans), pet izveštaja u `modIzvestaj` (nula).

**To nije argument za novo stablo** — novo stablo te čitaoce takođe ne bi prevelo.
Ali znači da je **konverzija čitalaca stvaran, nepopisan posao**, i ona ulazi u
plan kao imenovana stavka umesto da se podrazumeva.

#### Kriterijumi se ZAMENJUJU, ne relaksiraju

Tri „nejasno" nisu neodlučnost nego **defekti kriterijuma**. Za sledeći slajs
(Otpremnica, PR7) kapija nosi merljive pragove:

| Umesto | Novi prag | Danas |
|---|---|---|
| „kraće od zbira `SaveOtkup*`" | jezgro slajsa **ne sme biti >20% veće** od jezgra `CreateZbirna_TX` (greenfield brat) | 695 vs 770 → **10% manje** |
| „broj `T_ZBR*` pada" | poimeničan spisak testova koji **moraju nestati**, upisan **pre** početka slajsa | — |
| „većina modula adaptirana" | populacija se **popisuje** (svi moduli koji čitaju tabelu na baznom commitu); nedirnut modul koji čita nepopunjenu kolonu broji se kao **odložen**, ne kao reuse | 17 konvertovano, **14 odloženo** |
| „dual model" u jednom redu | dva reda: **dual WRITE** (prag 0, ispunjen) i **dual READ** — produkcioni čitaoci linijskih polja sa zaglavlja (prag 0, **danas nije ispunjen**) | — |

Granica „jezgra" se takođe upisuje, jer bez nje merenje nije ponovljivo:
**jezgro = `CreateX_TX` + tranzitivno zatvaranje do `Monitor_*`/`LogError`,
uključujući `ApplyAvansToOtkup`, isključujući deljenu infrastrukturu
`modDataAccess`/`modSchema`.**

---

#### Original kriterijuma (pre zamene, 13.09.2026)


Posle Otkup cutover-a i skele za Otpremnicu i Zbirnu (tabela PR-ova: red 6)
donosi se formalna
odluka: **nastavak u mestu** ili **novo stablo koda**.

> Kapija je posle promene redosleda **jača nego ranije**: Otkup je dokument sa
> najviše pisača (danas 9) i najviše integracija, pa se kriterijumi mere na
> najtežem slajsu umesto na najlakšem. Kriterijumi su merljivi,
sa alatima koji već postoje.

**Nastavljamo u mestu ako:**

| Metrika | Kako se meri | Prag |
|---|---|---|
| Pisaca nad `tblOtkup` | `who_writes.py` | **9 → 1** (`modOtkup`); `modSetup` ostaje samo `schema_owner` |
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

### 14.2) PR7 acceptance mreza — tvrdnje koje je Otkup cutover preselio

Ove tvrdnje su **do Otkup cutover-a bile zelene**, a posle njega se ne mogu
odrzati bez vracanja starog modela. Nisu obrisane nego **preseljene**: PR7
(Otpremnica cutover) ih preuzima nad `tblOtpremnicaIzvori`, gde veza vise nije
pogodjena iz kolona nego upisana.

**PR7 nije gotov dok svaka od njih ne bude zelena — po imenu.**

> **Odluka 15.09.2026 (§14.7):** tvrdnje **6** i **7** i zbirni deo tvrdnje **3**
> prelaze u PR8. PR7 dokazuje 1, 2, 4, 5 i tvrdnju 3 suženu na vezu otkup →
> otpremnica preko `tblOtpremnicaIzvori`.

#### Zasto je veza uopste pukla

`AutoLinkOtkupOtpremnica` (`modSledljivost`) **pogadja** vezu kljucem
`StanicaID + Datum + VozacID + Klasa + BrojZbirne` **nad `tblOtkup`**. Nov pisac
ne pise nijedno od tri: vozac je cinjenica otpremnice, klasa cinjenica stavke, a
broj zbirne je labela tudjeg dokumenta (A2). Kljuc zato vise nikad ne pogadja, pa
`Otkup.OtpremnicaID` ostaje prazan — a na njemu stoji ceo nizvodni citac:

| Sta | Gde | Stanje posle cutover-a |
|---|---|---|
| `AutoLinkOtkupOtpremnica_TX` | dugme „Auto-povezi", `modScrSledljivost:432` | povezuje 0 (toast pokazuje „Povezano: 0") |
| `TraceByZbirna` | GlobalGAP sledljivost, `modSledljivost:519` | vraca `Empty` za nov otkup |
| `StampajSledljivostZbirne` | `modIzvestaj:5966` | vraca `"NEMA"` — **ne stampa prazan list** |
| `GetKooperantiZaZbirnu` | paletni list, `modPaletniList:2614` | prazno polje kooperanata |
| auto-lanac hladnjace | `modAutoHladnjaca` | **pauziran** u `modOtkupUnos` (odluka operatera) |

Privremen citac se **ne pravi** — to je izricita odluka: kolona
`Otkup.OtpremnicaID` i kolona `Otkup.BrojZbirne` su bas ono sto refaktor brise
(S11.3), pa bi shim bio rad u pogresnom smeru.

#### Preseljene tvrdnje — čuva se ISHOD, ne staro ime

> **Ispravljeno posle review-a.** Prva verzija ove tabele je tvrdnje prenela
> doslovno, pa je tražila da isti otkup bude član „otpremnice Klase I" **i**
> „otpremnice Klase II". To bi u PR7 ponovo uvelo **dokument po klasi** — tačno
> model koji se uklanja. Ime stare tvrdnje ne sme da zaključa pogrešan grain.

Ciljni grain (§4.2, i PR5 API to već tako radi — `CreateOtpremnicaIzIzvora_TX`
prima **kolekciju izvora** i sam izvodi stavke):

```
OTK1  ├ I   400 kg          OTP1  ├ I   400 kg
      └ II  600 kg                └ II  600 kg

tblOtpremnicaIzvori:  OTP1 -> OTK1        JEDAN red, ne dva
```

| # | Ishod koji PR7 mora da dokaže | Izvorna tvrdnja |
|---|---|---|
| 1 | dvoklasni otkup je član **tačno jedne** otpremnice | `Otkup class I linked to matching otpremnica` |
| 2 | ta otpremnica nosi **obe klase kao svoje stavke**, sa istim količinama | `Otkup class II linked to matching otpremnica` |
| 3 | sledljivost čita **članstvo**, ne `Otkup.OtpremnicaID` | `TraceByZbirna returns rows` |
| 4 | povezivanje je **upis**, ne pogađanje — nema heuristike po `Stanica+Datum+Vozac+Klasa` | `Positive autolink links exact unique scenario` |
| 5 | dva **aktivna** članstva za isti otkup su tvrda greška (A15) | `Auto-link must NOT link otkup with different BrojZbirne` |
| 6 | hladnjački lanac vezuje dokument za **jednu** otpremnicu koju je sam napravio | `Hladnjaca lanac: otkup red povezan sa otpremnicom` |
| 7 | broj zbirne se **čita kroz lanac članstva**, ne prepisuje na otkup | `Hladnjaca lanac: otkup red nosi BrojZbirne` |

Red 5 je namerno preformulisan: stara tvrdnja je čuvala da heuristika ne „precuri"
na tuđu zbirnu. Kad povezivanje prestane da bude pogađanje, precurivanja nema —
ostaje jača invarijanta koju `AktivnoOtpClanstvoPoKanonu` već drži.

Red 6 isto: **jedna** otpremnica, ne „obe svoje otpremnice". Lanac koji za jedan
otkup pravi dva izvedena dokumenta po klasi je stari model.

#### Sta u medjuvremenu stoji umesto njih

Gubitak nije tih — svako mesto nosi **imenovano merenje** koje pukne cim se stari
model vrati:

| Test | Tvrdnja koja stoji danas |
|---|---|
| `Test_FullDocumentChainHappyPath` | `Cutover: otkup se vise ne nalazi po (broj, klasa)`, `Cutover: auto-link ne povezuje header+stavke otkup`, `Cutover: TraceByZbirna je prazna bez auto-link veze` |
| `Test_AutoLinkNeVidiHeaderStavkeOtkup` (nov) | `Auto-link slep: zaglavlje ne nosi Vozaca` / `... ne nosi Klasu` / `savrsen par OSTAJE nepovezan` |
| `Test_HladnjacaChainHappyPath` | `vraca SAMO poznat cutover gap`, `zaglavlje nosi SAMO otpremnicu Klase I (gubitna veza)`, `otpremnica Klase II postoji ali nije u zaglavlju` |
| `Test_StornoKaskadaScopePoLancu` | `Kaskada scope: lanac vraca SAMO poznat cutover gap` |

Sabotaza je meren dokaz, ne tvrdnja: vracanje `VozacID` i `Klase` na zaglavlje
obara **6** tih provera po imenu; obaranje back-linka lanca obara **3**.

#### Obrisani testovi (ne sele se)

| Test | Zasto |
|---|---|
| `Test_OtkupAtomicMultiClassSave` | merio „appends exactly two rows" — bas oblik koji se uklanja; zive tvrdnje nose `Test_OTK_HeaderIStavke` i `Test_OTK_LosaDrugaStavkaRollback` |
| `Test_OtkupClassIIAmbalaza` | merio `KolAmbalaze` po klasi na zaglavlju; zamenjuje ga `Test_OTK_AmbalazaIdeNaDokument` (jedan dvojni upis po dokumentu) |
| `Test_AutoLinkPositiveUniqueMatch` | tvrdnja preseljena (red 4); sam test meri pogadjanje, koje prestaje da postoji |
| `Test_AutoLinkMustNotCrossBrojZbirne` | tvrdnja preseljena (red 5), isti razlog |

#### Sta je ostalo od starih pisaca

`SaveOtkupMulti_TX` je **obrisan** — sa njim i posledji pisac koji je jedan unos
pretvarao u dva reda `tblOtkup`.

`SaveOtkup_TX` **namerno ostaje**, i to nije previd: nijedan produkcioni put ga
ne zove, ali je jedini posten nacin da test napravi **zaglavlje bez stavki** —
oblik koji nove kapije moraju da odbiju (`Test_OTK_VrednostBezStavkiPada`,
`Test_OTP_StariOtkupNeUlazi`). Alternativa bi bila `AppendRow` iz testa, sto
duplira znanje o semi i cini test pisacem tabele (A11). Odlazi u koraku 7,
zajedno sa kolonama koje puni.

#### Zatecen nalaz usput (ne dira se u ovom PR-u)

`OtkupIdsByBrDok` (`modScrDokumenti:668`) trazi otkupe **samo po
`BrojDokumenta`**, bez stanice. Posle `RequireBrojJedinstven`, cija je oblast
`StanicaID + Datum + Broj`, dve stanice smeju istog dana imati isti broj — pa bi
stampa spojila dva razlicita dokumenta. Nalaz je **zatecen** (i pre refaktora je
broj bio slobodan tekst), ne uveden; ide u citalacki prolaz, korak 7.

---

### 14.3) Read-model isplata: status je izveden, ne kesiran (korak 4)

`tblOtkup.Isplaceno` i `tblOtkup.DatumIsplate` bile su **keš** koji je održavao
`modNovac.UpdateOtkupStatus`. Keš je imao dva problema, i drugi je stariji od
refaktora:

1. vrednost je računao kao `Kolicina x Cena` **sa zaglavlja**. Posle prelaska na
   header + stavke to je uvek nula — pa nijedan nov otkup ne bi nikada bio
   označen kao isplaćen, tiho i bez ijedne greške.
2. tačnost mu je zavisila od toga da ga **svaki** pisac `tblNovac` pozove. Bilo
   je četiri takva mesta (`modNovac` ×2, `modBankaMapiranje`, `modDokumenta`) i
   peto u `modStorno`. Šesto koje bi zaboravilo poziv dalo bi otkup koji je
   plaćen a izgleda otvoren, ili obrnuto.

`UpdateOtkupStatus` je **obrisan**, sa svih pet poziva. Status je sada izveden:

```
otvoreno = VrednostOtkupa(id) - SUM(isplate za id)
```

`GetOpenOtkupi` tako i računa; kolone nemaju pisca i brišu se u koraku 7.

#### Pun ugovor `VrednostOtkupa`

Ugovor je do sada bio **imenovano nepotpun** (komentar ispod funkcije): kapije bi
oborile 10–33 tvrdnje jer su postojala dva pisca zaglavlja bez stavki. Oba su
zatvorena — stari pisac obrisan (korak 3), PWA ide kroz `CreateOtkup_TX`
(korak 2) — pa ugovor sada stoji ceo:

| # | Kapija | Zašto nije kozmetika |
|---|---|---|
| 1 | prazan `OtkupID` | računa se vrednost ničega |
| 2 | zaglavlje **tačno jednom** | dva reda znače da `OtkupID` nije više identitet |
| 3 | svaka stavka brojčana i **> 0** | isto pravilo koje pisac traži na upisu (`modOtkup:252/261`); citalac koji ga ne drži tiše je od istine |
| 4 | bar jedna stavka | nula je legitiman odgovor samo kad stavke postoje a zbir im je nula — inače `ApplyAvansToOtkup` čita 0 kao „nema šta da se plati" |

#### Dva neispravna reda, dva različita odgovora

Razlika je **merena**, ne stilska:

| Red | Odgovor `GetOpenOtkupi` | Zašto |
|---|---|---|
| bez `OtkupID` | **ostaje u listi**, prisilno | ne može se vrednovati ali se ne sme ni izgubiti — to je kvar FM-0021 #5; imenuje ga `BuildBlokIsplataList` sa `ERR_ISPLATA_PRAZAN_OTKUPID` |
| sa `OtkupID`, bez stavki | **pada po imenu** (review #334, P1) | do tog review-a se preskakao, uz brojač i jedan log red. I to je tišina: red nestane iz liste za isplatu kao da je plaćen, dok ga mreža istovremeno prikazuje sa 0 kg i pilulom „plaćeno“. Kapija je sada **jedna** -- `modOtkup.StavkeOtkupaRedovi` -- pa ovaj čitalac pada tamo gde padaju i izveštaji |

Ono što se **ne radi** ni u jednom slučaju: računanje vrednosti sa zaglavlja. To
bi bio compatibility sloj i vraćao bi 0 za svaki nov dokument.

#### Dokumentski ugovor bulk čitaoca (review #334, P1)

`VrednostOtkupa` drži ugovor za **jedan** dokument. Čitaoci koji vrednuju **sve**
dokumente odjednom (mreža, izveštaji, KPI, rang, lista za isplatu, kandidati
banke) išli su kroz slabiju kopiju pravila, pa je dokument bez stavki izlazio
kao **0 kg i 0 dinara** — a u mreži i kao **„plaćeno“**: `duguje` se računao iz
zbira stavki kojih nema, pa je ispadao 0, a `PayCode(0, plaćeno > 0)` vraća
PLAĆENO. Header-only korupcija je tako izgledala kao uredno zatvoren dokument.

Od #334 svi idu kroz **jedan** prolaz — `modOtkup.StavkeOtkupaRedovi` — koji drži:

| # | Kapija | Broj | Zašto nije kozmetika |
|---|---|---|---|
| 1 | stavka ima neprazan `OtkupID` | 1907 | stavka bez dokumenta se ne može vrednovati; ranije tiho preskočena, pa je zbir bio manji nego što jeste |
| 2 | stavka ima **zaglavlje** | 1908 | siroče se ne sme sabrati ni u čiji zbir |
| 3 | zaglavlje sa `OtkupID`-em ima **bar jednu stavku** | 1909 | „nema ključa u rečniku“ je davalo 0 kg / 0 din i pilulu „plaćeno“ |
| 4 | `Kolicina` i `Cena` brojčane i > 0 | 1905/1906 | isto pravilo koje pisac traži na upisu |
| 5 | `Klasa` je I ili II | 1830 | **ista procedura** koju zove pisac (`RequireValidOtkupClass`) |
| 6 | `KolAmbalaze` brojčana, ≥ 0, ceo broj | 1919/1920/1886 | pisac traži isto; čitalac je nebrojčanu vrednost tiho čitao kao 0, pa su gajbe nestajale |
| 7 | zaglavlje se nalazi **tačno jednom** | 1921 | čitaoci iteriraju zaglavlja, pa dva reda sa istim ID-em isti teret broje **dvaput** — a sa različitim `KooperantID` ga pripišu **dvojici** kooperanata |

Nijedan čitalac više nema granu `If dict.Exists(id) … Else 0`: ključ se traži
kroz `modOtkup.ZbirStavkiZaOtkup` (odnosno `modNovac.VrednostOtkupaIzDikta`),
koji nedostatak prijavljuje **po imenu**. To je druga brana — izvor već pada —
ali grana koja nulu vraća kao podatak ne sme da postoji nigde.

Isti prolaz koristi i put novca: `BuildVrednostDictByOtkup` više nema svoju
kopiju pravila, pa lista za isplatu i kandidati banke padaju na istom mestu na
kom padaju izveštaji. Ranije je banka nedostajuću vrednost čitala kao 0, blok
je ispadao iz kandidata, a uplata se knjižila kao **avans** — tiho.

**Jedan rod je namerno izuzet, i ima imenovanog vlasnika:** **prazan `OtkupID`
na zaglavlju** — red ostaje u listi za isplatu (FM-0021 #5) i imenuje ga
`BuildBlokIsplataList` (`ERR_ISPLATA_PRAZAN_OTKUPID`); čitaoce vrednosti obara
`ZbirStavkiZaOtkup` na mestu upotrebe, pa ni tamo nije nula.

**Dupli `OtkupID` je u drugom krugu review-a prešao iz izuzetka u kapiju.**
Prvi predlog je bio da ga i dalje drže samo finansijske kapije
(`ERR_ISPLATA_DUPLI_OTKUPID`, `VrednostOtkupa` 1902, `BuildOtkupOwnerIndex`).
To je bilo pogrešno iz jednog konkretnog razloga: te kapije čuvaju **svoje**
putanje, a izveštaji, KPI i mreža iteriraju zaglavlja — dva reda sa istim ID-em
tamo daju **2× kg** nad jednom fizičkom stavkom, a sa različitim `KooperantID`
iste kilograme pripišu dvojici. Tolerancija u novom centralnom čitaocu je zato
uklonjena; `RequireJedinstvenZaglavljeOtkupa` pada pre agregacije.

Cena je izgovorena: **jedan** duplirani `OtkupID` bilo gde — uključujući
istorijski, plaćeni — zaustavlja svakog čitaoca vrednosti dok se podatak ne
ispravi. Prihvatljivo je jer nema produkcionih podataka, `NewEntityID` je
opaque pa ga nijedan pisac ne može proizvesti, a `modProductionHealthCheck`
ostaje tolerantan baš zato da takav red **imenuje** umesto da padne na njemu.
`ERR_ISPLATA_DUPLI_OTKUPID` ne gubi dokaz: `modTestBanka` **T15** je dokazuje
direktno nad `BuildOpenAmountDict`, a **T17** sada tvrdi raniji, jači pad.

Regresija: `Test_OTK_ZaglavljeBezStavkiObaraCitaoce` (dokument se isplati do
kraja, pa mu se stavke obrišu — mreža, saldo OM i lista za isplatu moraju pasti
po imenu, a pilula ne sme reći „plaćeno“) i
`Test_OTK_StavkaBezZaglavljaObaraCitaoce` (siroče, klasa, `KolAmbalaze`). Oba
čiste za sobom (`clsTransaction` + rollback), kao i četiri zatečena testa koja
su ostavljala zaglavlje bez stavki.

#### Fixture je morao da postane dokument

`tools/make_fixture.py` je sejao 32 `tblOtkup` zaglavlja i **nijednu** stavku —
tabela `tblOtkupStavke` nije ni postojala u donoru (pravio ju je runtime
self-heal). Posle prelaska čitalaca na stavke svih 32 su bila „nije dokument", pa
je pola testova banke ostalo bez ijednog reda.

Generator sada pravi tabelu (`ENSURE_TABLES`, kolone iz kanona) i **izvodi**
stavke iz zaglavlja. Kolone `Kolicina/Cena/Klasa/KolAmbalaze` na zaglavlju se
namerno ne brišu — čita ih još petnaestak modula, to je korak 7 — pa fixture nosi
iste brojeve na oba mesta, a merodavna je stavka.

Isto pravilo je moralo u seed-ove testova: **„isplaćen" sada znači *plaćen***, ne
„označen kao plaćen". `SeedOtkupIsplacen` (banka) i `SeedOtkupPlacen` (storno)
zato knjiže i pokrivajuću isplatu u `tblNovac`.

#### Tri tvrdnje koje su nestale, i zašto

| Tvrdnja | Odakle | Zašto ne može da ostane |
|---|---|---|
| `T18 (3): zatvoren red bez OtkupID NE obara pregled` | `modTestBanka` | „zatvoren" se čitalo iz kolone; status je sada izveden iz plaćanja, a plaćanje se vezuje **preko** `OtkupID`-a — red bez njega ne može biti plaćen. Stanje više ne postoji, pa ostaje jedno pravilo umesto dva |
| `ResetNovacOtkupLink recomputes otkup as unpaid` | `modNovacTests` | prepisano na `ResetNovacOtkupLink vraca otkup medju otvorene obaveze` — ista zaštita, merena tamo gde operater gleda |
| `T27: blok vise nije isplacen` / `datum isplate ocisceni` | `modTestStorno` | prepisano na `posle storna blok je opet otvoren`. Tvrdnja nad kolonom bez pisca bila bi zelena i kad blok ostane nevidljiv — tačno kvar koji test sprečava |

#### Šta stoji umesto njih

| Test | Šta meri |
|---|---|
| `Test_OTK_VrednostPunUgovor` | sve četiri kapije, **po poruci** — pad iz drugog razloga ne dokazuje ništa |
| `Test_OTK_StatusIsplateJeIzveden` | nov dokument je otvoren; delimična isplata ga **ne** zatvara; puna ga zatvara |
| `Test_ResetNovacOtkupLinkRecomputesStatus` | skidanje veze vraća obavezu u listu |
| `T27_StornoIzvodaOsvezavaOtkup` | storno izvoda vraća dug |

Sabotaža (dokaz da kapije mere): gašenje pravila „stavka > 0" obara 1 proveru po
imenu, gašenje provere „zaglavlje tačno jednom" 1, a povratak računa na zaglavlje
obara 2 — uključujući „nov dokument je otvorena obaveza", tj. baš tihu rupu zbog
koje je keš i obrisan.

#### Zdravstvena provera više ne meri keš

`Check_OtkupPaymentConsistency` je merio **samo redove koje je keš označio** —
preplata na neoznačenom redu se nije ni videla. Sada čita stavke i prijavljuje
preplatu (`isplaceno > vrednost`) nad svakim redom, plus upozorenje za redove bez
stavki. `Check_KooperantOtkupReconciliation` isto: zbir po kooperantu je išao iz
zaglavlja, pa bi se posle cutover-a „slagao" — na nuli.

Da kolone ne dobiju novog pisca čuva **statička** kapija (`who_writes
--check-ownership`, A11), ne runtime provera: zatečen red koji još nosi staru
vrednost nije kvar, a runtime kapija nad njim bi vikala na podatke umesto na kod.

---

### 14.4) Lokalni correction API za otkup (korak 5)

A9 traži vezu **po ID-u**; zatečeni aparat je drži **poslovnim brojem**
(`modStornoFlow.StampIspravkaTrace`) i zato ne razlikuje dve verzije istog
dokumenta. Otkup je prvi koji prelazi.

#### Kolone su preimenovane, ne udvojene

Na `tblOtkup`: `IspravkaOd → IspravkaOdID`, `ZamenjenSa → ZamenjenSaID`. Kolona
koja se zove `IspravkaOd` a nosi ID bila bi gora od obe. Ostale tri tabele
(Otpremnica, Zbirna, Prijemnica) još nose broj-oblik i prelaze u PR7/PR8.

Preimenovanje nosi tri posledice koje se lako previde:

1. `modSchema` poredi i **poziciju** (upis je pozicioni), pa dodavanje novog imena
   pored starog ne pomaže — zatečena sveska bi i dalje bila odbijena na poziciji
   32. Zato `modSetup.PreimenujKolonuAko` menja ime **na mestu**, čuvajući i
   poziciju i podatke, i radi samo kad staro ime postoji a novo ne.
2. `modSetup.EnsureSledljivostSchema` je kolone dodavao kroz petlju nad šest
   tabela — `tblOtkup` je morao da izađe iz te petlje, inače bi self-heal vraćao
   ime koje je kanon upravo uklonio.
3. `tools/make_fixture.py` gradi fixture iz **donora**, ne pokretanjem aplikacije,
   pa je dobio isti aparat (`RENAME_COLS`) i potpis koji ga pokriva.

Sadržaj se **ne prevodi**: na `tblOtkup` te kolone nikad nisu ni pisane
(`StampIspravkaTrace` se zove samo za druge tri), pa je svaka zatečena vrednost
prazna.

#### `IspravkaOtkupa_TX` — jedna transakcija

```
OTK-101 / broj 17   Stornirano=Da,  ZamenjenSaID = OTK-202
      v
OTK-202 / broj 18   IspravkaOdID = OTK-101
oba nose isti CorrectionID
```

Redosled unutar transakcije je **meren**: storno ide **prvi**, da bi
`ResetNovacOtkupLink` oslobodio vezan novac pre nego što novi dokument kroz
`ApplyAvansToOtkup` uopšte potraži avans.

`CorrectionID` je **pravi** ID iz `tblStornoVeze` (`CreateCorrectionContext` +
`CompleteCorrectionContext` u istoj transakciji), ne broj koji pisac izmisli —
kolona je deklarisana kao veza na tu tabelu, a izmišljen ID bi bio viseći
pokazivač. Test to i tvrdi (`RowExists(TBL_STORNO_VEZE, ...)`).

Tri kapije, sve merene **porukom** a ne samo padom:

| Kapija | Poruka imenuje |
|---|---|
| isti broj kao stari | pravilo o **novom broju** (A9) — ne jedinstvenost |
| dokument već zamenjen | **postojećeg naslednika** („ispravlja se POSLEDNJA verzija") |
| izvor stornirano | storno |

Redosled prve dve je takođe meren: već ispravljen dokument je **uvek** i
storniran, pa bi storno-kapija prva uhvatila oba slučaja i rekla manje korisnu
istinu. Specifičnija ide prva.

`PoslednjaVerzijaOtkupa` prati `ZamenjenSaID` do kraja lanca, sa brojem koraka
ograničenim brojem redova — ciklus u podacima ne sme da zavrti čitaoca.

#### ⚠ NALAZ: novac se **ne** prenosi na ispravku

Očekivanje je bilo suprotno i upisano je u pre-flight kao `PROVEN` na osnovu
čitanja samo prve polovine lanca. **Merenje kaže drugačije**, i test to sada
tvrdi:

```
StornoOtkup -> ResetNovacOtkupLink -> isplati se prazni OtkupID
ApplyAvansToOtkup -> uzima SAMO Tip = NOV_VIRMAN_AVANS_KOOP   (modNovac:1624)
```

Odvezana isplata tipa `VirmanFirmaKoop` zato ostaje nevidljiva **i** za dug (nema
`OtkupID`) **i** za avans (pogrešan tip — isti filter je i u
`GetKooperantUnallocatedAvans:1982` i `BuildKooperantUnallocatedAvansDict:2031`).
Vidi se još samo u kartici kooperanta (`modIzvestaj:2507`), gde ulazi u saldo.

Novac dakle **nije izgubljen**, ali jeste ispao iz svake mašinerije koja odlučuje
šta se plaća — operater vidi novi dokument kao pun dug, a plaćeni iznos nigde
među avansima.

**Ovo nije uvedeno ispravkom** — isto radi običan `StornoOtkup_TX` i radio je
oduvek. Ispravka ga samo čini lakše dostižnim. Test ga tvrdi kao **zatečeno**
stanje (`Test_OTK_IspravkaNeGubiNovac`), da se ne bi tumačilo kao osobina novog
pisca; kad se donese odluka o prenosu, test se **okreće**.

#### Writer nema produkcionog pozivaoca u ovom PR-u

Merena odluka, ne propust. Da bi ekran zvao ovaj put, stari `OtkupID` mora da
otputuje od prefill-a do pisca. Danas prefill ide kroz spec string
(`PrefillIzStorniranog` → `modOtkupUI.ApplyPrefill:8452`) čiji se ključevi mapiraju
**na kontrole** — ID nema gde. Rešenje bi bilo modul-stanje u ljusci, tj. **isti
oblik in-memory veze** (`mPendingRelinkOldPrij`) koji ovaj korak treba da ukine.

Wiring zato ide sa PR7, kad se `PrefillIzStorniranog` ionako prepisuje sa
po-klasnih redova. Isti obrazac kao `CreateOtpremnicaDraft_TX` u PR5.

Sabotaža: gašenje A9 kapije o broju obara 5 provera po imenu, dozvola drugog
naslednika 1, izostanak `ZamenjenSaID` 4.

---

### 14.5) A13 kapija: izdat roditelj se ne menja ispod ruke (korak 6)

Izdata otpremnica je **papir sa sastavom**. Ispravka jednog njenog izvora nije
lokalna: otpremnica mora dobiti novu verziju sa novim brojem, a za njom i zbirna
(H1, `GOLDEN_SCENARIJI.md` §12). Ta propagacija je PR7 — do tada se staje glasno.

```
otkup u IZDATOJ otpremnici  ->  IspravkaOtkupa_TX PADA, imenuje otpremnicu
otkup u DRAFT otpremnici    ->  ispravka PROLAZI
otkup bez otpremnice        ->  ispravka PROLAZI
```

Tiha alternativa bi bila najgora: nov otkup, stara otpremnica netaknuta, i izdat
papir koji više ne opisuje robu koju nosi.

**DRAFT se namerno ne blokira.** Članstvo drafta je mutabilno po dogovoru, a
`IzdajOtpremnicu_TX` revalidira izvore pri izdavanju — storniran otkup ne može da
prođe kroz izdavanje. Kapija koja bi blokirala i draft izgledala bi isto zeleno,
a oduzela bi operateru ispravku dokumenta koji još niko nije video. Test zato meri
**obe** strane granice.

#### Pripadnost se pita, ne izvodi

Otkup nema kolonu koja bi rekla kojoj otpremnici pripada — to je §3.1, i namerno.
`modDokumenta` je zato dobio dva javna čitača nad kanonskim članstvom:

| Ulaz | Vraća | Napomena |
|---|---|---|
| `OtpremnicaZaOtkup(otkupID)` | `OtpremnicaID` ili `""` | dva aktivna članstva **podižu grešku**, ne vraćaju jedno |
| `OtpremnicaJeIzdata(otpremnicaID)` | `True` samo za `IZDATO` | prazan status **nije** „izdato" — nov pisac ga upisuje eksplicitno, pa je prazno polje drift |

Javni ulaz postoji baš zato da pisac otkupa ne dobije **drugu** implementaciju
istog pitanja — dupliranje je ono što je kanonsko članstvo trebalo da ukine.

Sabotaža: gašenje provere izdatosti obara **6** tvrdnji po imenu — uključujući
„odbijena ispravka nije upisala nijedan red", koja meri da je odbijanje potpuno, a
ne polovično.

---

### 14.6) Korak 7 je izmeren i podeljen (ne odložen)

„Reader sweep → brisanje kolona" je u planu stajao kao **jedna** stavka. Merenje
kaže da je to dve — i da se granica poklapa sa granicom PR-ova, ne sa procenom.

Popis (skripta broji pojave `COL_OTK_*` u kodu, bez komentara):

| Kolone | Pojava | Zašto odlaze | Kad su stvarno slobodne |
|---|---:|---|---|
| `OtpremnicaID`, `BrojZbirne`, `VozacID`, `BrojOtpremnice` | 141 | članstvo (A15) / labela (A2) | **PR7** — čitaoci su po-klasna otpremnica |
| `GeneracijaID` | 69 | identitet; §11.1 briše celu mašineriju | **PR8** (zbirna) |
| `Kolicina`, `Cena`, `Klasa`, `KolAmbalaze` | 170 | stavke (§4.1) | čitaoci su **isti moduli** kao gore |
| `Isplaceno`, `DatumIsplate` | 12 | keš izvedenog statusa (korak 4) | **sada** |
| `Novac`, `PrimalacNovca` | 9 | `tblNovac` | skoro — v. dole |

Ukupno **401 pojava u 27 produkcionih modula**. Od toga **210** pripada
čitaocima koje PR7/PR8 ionako prepisuju: sweep sada značio bi pisati čitaoce
protiv modela koji se uklanja — ista greška zbog koje privremeni čitač za
`TraceByZbirna` nije napravljen.

**Odluka operatera: u PR6 ide samo ono što je stvarno slobodno.**

#### Obrisano

`tblOtkup.Isplaceno` i `tblOtkup.DatumIsplate` — posle koraka 4 nemaju **nijednog
pisca** i nijednog živog čitaoca. Otisak šeme: `EC04DB7C → A410CF67`, 648 → 646
kolona.

Uz njih je otišao i `modOtkup.GetSaldoByStation`: sabirao je `Kolicina`, `Novac` i
`KolAmbalaze` **sa zaglavlja** po kooperantu — tri kolone koje nov pisac ne piše.
Da ga je iko zvao, vraćao bi nule. Grep po celom `src-vba` daje samo redove unutar
same funkcije: mrtav čitač mrtve kolone, pa se briše a ne prepisuje.

`Novac` i `PrimalacNovca` **ostaju**: njihov jedini živi čitač je po-klasni lister
u `modDokumenta:6073`, koji je PR7. Ostala dva su mrtva (`GetSaldoByStation`,
sada obrisan) i pauzirana (`modOtkupBlok` prefill).

#### Sveska mora da IZGUBI kolonu, ne samo kanon

Ovo je posledica koju je lako prevideti: kolona obrisana iz **sredine** kanona
pomera sve iza sebe, a `AppendRow` piše **pozicijski**. Zatečena sveska sa viškom
tada šalje vrednosti u pogrešna polja — `modSchema` to hvata i staje, ali sveska
ostaje neupotrebljiva dok se višak ne ukloni.

Zato su dva mesta dobila migraciju:

| Gde | Šta radi |
|---|---|
| `modSetup.ObrisiKolonuAko` | briše na startu aplikacije, **samo po imenu iz uskog spiska** |
| `tools/make_fixture.py` `DROP_COLS` | isto, jer generator gradi iz donora a ne pokretanjem aplikacije |

Brisanje je **jedini destruktivan korak** u self-heal-u, pa je i najuži: nema
petlje po „sve što nije u kanonu" — `EnsureRuntimeSchema` legitimno dodaje kolone
na kraj pre nego što ih kanon preuzme, pa bi takva petlja brisala tekući rad.

#### Destruktivan put je dobio meru

`ObrisiKolonuAko` i `PreimenujKolonuAko` do sada **nije izvršio niko**: fixture je
migriran generatorom, pa je VBA put ostao mrtav. Kod koji briše kolonu a nikad
nije pokrenut je najgora vrsta nemerene odbrane.

`Test_OTK_SelfHealMigracijeKolona` radi nad kolonom koju sam doda **na kraj**
tabele — višak na kraju ne pomera nijednu kanonsku poziciju (otisak se računa nad
kanonskim prefiksom), pa ni pad testa ne ostavlja svesku u lošem stanju. Test
pokriva: preimenovanje, idempotenciju, brisanje, brisanje nepostojeće kolone, i
slučaj **oba imena odjednom** (stanje koje pravi međuverzija).

Sabotaža: gašenje brisanja obara 4 tvrdnje po imenu, pretvaranje preimenovanja u
dodavanje 5.

> **Jedna kapija namerno ne bije.** Gašenje provere „već migrirano" ne obara
> nijednu tvrdnju — mereno. Excel sam odbija drugu `ListColumn` istog imena, pa se
> ishod ne menja. Kapija štedi `LogError` na **svakom** startu takve sveske, ne
> podatak; to je razlog zbog kog ostaje, i tako je i imenovana u kodu.

---

### 14.7) PR7 pre-flight, ponovo izmeren — granica PR7/PR8 i ugovor koji protivreči sam sebi (15.09.2026)

> Pre-flight „Otpremnica cutover (PR7)" (gore) pisan je 12.09, pre merge-a #308;
> posle njega je merge-ovan još 21 PR. Ovde je ponovo izmeren na `main`
> `a173c134` i dopunjen sa dva ulaza koja §14.1 traži **pre** slajsa: popisom
> čitalaca i poimeničnim spiskom testova. Oba su u prilogu
> **`docs/REFAKTOR_PR7_POPIS.md`**.
>
> **Verdikt posle odluka operatera (15.09):** granica PR7/PR8 je odlučena (A),
> ugovor PR7 je usklađen, a capability mapa dopunjena — v. „Odluke operatera“.
> PR7 **još ne kreće u kod**: pre njega idu mali PR za kvarove 2, 3 i 9, ponovljen
> popis sa nezavisnom proverom i odluke iz „Još otvoreno“ (ambalaža otpremnice je
> `DOMAIN GAP`).

Oznake: **✔** ručno provereno čitanjem koda · **◐** potvrdio nezavisni
proveravač · **○** jednoprolazna klasifikacija (prilog, „Kako je mereno").

#### Verdikt po osama

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN, **jedan GAP** | §4.2, A15. **GAP — ambalaža otpremnice:** jedini `TrackAmbalaza` sa `DOK_TIP_OTPREMNICA` je legacy `SaveOtpremnica` (`modDokumenta:399`) ✔; ulazi skele izlaz gajbi stanica→vozač ne knjiže ✔, a odluka ko ga knjiži u novom modelu nije nađena ○ — **otvoreno do odluke pre F2 dela PR7** |
| `EVENTS` | PROVEN | tabela od 12.09 važi; otvoreno je samo *kada* se knjiži izlaz ambalaže (nastanak ili izdavanje) — isto pitanje kao gore |
| `IDENTITY` | PROVEN | `OtpremnicaZaOtkup` (`modDokumenta:3423`) i `OtpremnicaJeIzdata` (`:3439`) postoje ✔ |
| `CARDINALITY` | PROVEN | 1 otkup → najviše jedno aktivno članstvo; posledica za golden je tačka 1 ugovora, ispod |
| `INVARIANTS/OWNER` | PROVEN | `AktivnoOtpClanstvoPoKanonu`; A11 allowlist za `tblOtkup` skraćen **9 → 8** u ovom PR-u — `modNovac` više ne piše tu tabelu ✔ |
| `WRITERS` | PROVEN | 18 upisnih mesta četiri vezne kolone u 7 modula (prilog). `Otkup.OtpremnicaID` i dalje piše istih šest modula iz NALAZA 1, na pomerenim linijama: `modAutoHladnjaca:394`, `modDokumenta:6942`, `modMasterSync:2777`, `modOtkupBlok:1480`, `modSledljivost:256`, `modStornoFlow:2639`. `SaveOtpremnica*` ima 4 produkciona poziva: `modDokUnos:269`, `modMasterSync:919`, `modAutoHladnjaca:230/:275` ✔ |
| `DOWNSTREAM` | PROVEN — **odluka 15.09** | granica PR7/PR8 = opcija A: zbirni tokovi imenovano pauzirani do PR8, zbirni čitaoci odloženi po §14.1 (v. „Odluke operatera“) |
| `CAPABILITY` | PROVEN — **odluka 15.09** | ručna zbirna F3 i auto-lanac hladnjače: `PAUZIRAN` do PR8; GlobalGAP sledljivost zbirne: PR8; ručno „Poveži“ i „Preuzmi“ izgubljeni blok: `MIGRATED` na članstvo u PR7, nad izdatom otpremnicom odbijeno po A15 ○; prefill ispravke otpremnice: `MIGRATED` u PR7 ✔; izvoz za PWA menadžment: zaseban pre-flight (kvar 4) |
| `ACCEPTANCE CONTRACT` | PROVEN — **odluka 15.09** | PR7 dokazuje §14.2 tvrdnje 1, 2, 4, 5 i suženu 3; tvrdnje 6, 7 i zbirni deo 3 → PR8; golden na test-only pisaču (cilj: 12/0 bez promene snapshot-a); H2 → PR8; A11 `tblOtkup` 8 → 1 |
| `PLATFORM` | N/A | nema Excel/COM nepoznanice |
| `LANDING` | PROVEN | #308 merge-ovan 12.09 — rizik od 12.09 je otpao ✔. `modDokumenta`, `modStornoFlow` i `modOtkup` su posle #308 menjani u po 6 commit-a, pa grana ide od `main` ✔ |

#### ⚠ Glavni nalaz: PR7 kako je napisan sam gasi zbirnu — samo bez imena

Skela ne piše na zaglavlje otpremnice ni `BrojZbirne` ni linijska polja ✔
(`BuildOtpremnicaHeaderRowData`, `modDokumenta:2724`; `COL_OTP_BROJ_ZBIRNE` se ne
pojavljuje ni u jednom ulazu skele). A sve živo što pravi ili proverava zbirnu
nalazi izvore **po `Otpremnica.BrojZbirne` i sabira po-klasne redove**. Kad skela
postane jedini put:

| Tok | Posle PR7 | |
|---|---|---|
| ručna zbirna F3 (`ZbirnaValidiraj` → `ValidateZbirnaPreUnosa`, `modDokumenta:3742`) | pada na „validacija nije prošla" — zbir izvora je 0, a poruka ne kaže da je reč o pauzi | ○ |
| invarijanta i rekalkulacija zbirne (`SumOtpremniceByKlasa`, `modDokumentInvariant:42`) | tiho: pad kolone proguta EH i vrati nule; čim bilo koji put upiše `BrojZbirne` na nov header, rekalkulacija upiše 0 kg u zbirnu, a invarijanta kaže OK | ○ |
| ciljni pisac zbirne `CreateZbirna_TX` (PR3 skela) | pada na praznoj klasi izvora | ○ |
| prijemnica F4 | podrazumevano BLOK — zbirne nema | ○ |
| auto-lanac hladnjače, ako se prevede samo korak otpremnice | `ZavrsiVezuOtpremniceNaZbirnu` upiše `BrojZbirne` na nov header → rekalkulacija nulira zbirnu | ○ |
| ispravka otpremnice koja ima zbirnu (`CompleteOtpremnicaIspravka`, `modStornoFlow:633`, `:780-787`) | nova zbirna je prazna i različita od stare → stara se stornira ili rekalkuliše bez nove otpremnice, a završetak se prijavi kao uspeh | ○, nereprodukovano |

Opcije (pun tekst i brojevi: prilog):

| | Šta | Operater gubi | Krši |
|---|---|---|---|
| **A** | PR7 **imenovano pauzira** zbirne tokove do PR8, obrazac PR6: kapija se deli (8 postojećih mesta), F3 dobija svoju; zbirnih čitalaca u PR7 nema | za nove podatke između PR7 i PR8: zbirnu, prijemnicu, palete, fakturu, sledljivost zbirne | nijedan princip modela; **menja acceptance PR7** (§14.2 tvrdnje 3, 7 i zbirni deo 6 prelaze u PR8) |
| **B** | privremen join po `Otpremnica.BrojZbirne`, zbirni čitaoci na stavke otpremnice | ništa | §14.2 („privremen čitač se ne pravi"), §4, §3.1/A15; 11 procedura se radi dvaput, tvrdnje 3 i 7 zelene preko labele |
| **C** | PR7 i PR8 zajedno | ništa | redosled PR-ova; obim PR8 nije pre-flight-ovan; najduža grana |
| **D** | PR7 nosi i izvođenje pisca zbirne iz `tblOtpremnicaStavke`; čitaoci zbirne ostaju za PR8 iza kapija | storno i ispravku otpremnice u izdatoj zbirnoj; F3 menja oblik bez spec-a | dual READ na `tblZbirna`; lista pauza nije izmerena |
| **E** | legacy F2 pisač ostaje do PR8 | ništa | **dual WRITE** — okidač za novo stablo iz §14.1; odbačeno |

**Preporuka merenja je A**, uz tri uslova: (1) pauza je imenovana i glasna, sa
testom po imenu, umesto zatečenog „validacija nije prošla"; (2) acceptance se
menja izričito — tvrdnje 3, 7 i zbirni deo 6 u PR8, auto-lanac hladnjače ostaje
**ceo** pauziran do PR8, GlobalGAP `REPLACED` u PR8; (3) golden i test rep se
odlučuju izričito (tačka 2 ugovora). Ako prozor bez zbirne na `main`-u nije
prihvatljiv, izbor je **C**, ne B ni D. **Izbor je operaterov.**

#### ⚠ Ugovor PR7 protivreči sam sebi

1. **„Golden 12/0 sa nepromenjenim snapshot-ima."** A4 deli **jedan** otkup od
   1000 kg na dve otpremnice od 400 i 600 (`modGoldenTests:1386-1388`) ✔. U novom
   modelu otkup je član najviše jedne aktivne otpremnice, a delimična alokacija
   nije modelovana (§3.1) — menja se **scenario**, ne adapter. Isto važi za
   A2/F2/G1, gde otpremnica Klase II nosi gajbe kojih u otkupu nema ○. §12 traži
   obrazloženje svake promene goldena u commit-u; golden se, dakle, **menja**.
2. **„Stari pisač se briše" (§11.2)** naspram goldena i ~57 test poziva
   po-klasne otpremnice ○. Ili pisač ostaje samo za testove do PR8 (presedan
   `SaveOtkup_TX` iz PR6), ili se sve prepisuje odmah. Oba ne mogu.
3. **„Edge koji mora proći: H2 (`CorrectionSestre`)."** Red 8 tabele PR-ova kaže
   da H1 i H2 registruje **PR8**, a u `modGoldenTests` ne postoje ✔.

Uz to: `SaveOtkup_TX`, jedini način da test napravi zaglavlje bez stavki, traži
kolone `VozacID`/`BrojZbirne`/`OtpremnicaID` ○ — pada čim ih PR7 obriše, pa
`Test_OTK_VrednostBezStavkiPada` i `Test_OTP_StariOtkupNeUlazi` traže drugi put.

#### ⚠ Kvarovi koji žive danas — posledica PR6, nezavisni od granice

Plan i komentar uz `NapredakBlokaDostupan` (`modOtkupBlok:1565-1573`) tvrde da je
napredak bloka pauziran. **Pauza stoji samo u mrtvom panelu:** kapija je na
`:225`, `:269` i `:1410` ✔, a `AttachOtkupBlokPanel` nema pozivaoca ○. Živa ljuska
zove iste zbirove bez kapije.

| # | Kvar | Dokaz | |
|---|---|---|---|
| 1 | traka otpremnice u F1: „u blokovima 0, ostatak = cela otpremnica"; upozorenje na prekoračenje ne ide; otpremnica nikad ne izlazi iz „otvorene" | `modScrDokumenti:171-172` → `modOtkupBlok.SumKolByOtp:1576` / `SumAmbByOtp:1617` sabiraju `tblOtkup.Kolicina`/`KolAmbalaze` po `Otkup.OtpremnicaID`, bez kapije ✔; isto `:1088`, `:2115` ◐ | ✔ |
| 2 | pill plaćanja u mreži otkupa: „plaćeno" posle prve delimične isplate (duguje = Kolicina × Cena sa zaglavlja = 0) | `modScrDokumenti:1768-1769`, `:1810-1823`, `PayCode:1609` | ◐ |
| 3 | vrednost otkupa u izveštajima je 0: saldo OM, kartica kooperanta, otkupne liste, prosečna cena, zbirni OM, roba po OM | `modIzvestaj:589-590`, `:895-896`, `:1522-1523`, `:2720`, `:3453-3454`, `:3788-3789` čitaju `Kolicina`/`Cena` zaglavlja ✔ | ✔ ◐ |
| 4 | izvoz za PWA menadžment šalje prazne `Kolicina`/`Cena`/`Klasa` za nove otkupe, a `TransportStatus` računa iz praznih veza | `modStammdatenSync:671-680` (`ExportOtkupiAll`), `:554-557` (`ExportOtkupPoOM`) ✔ | ✔ |
| 5 | sledljivost prijavljuje lažan KG-RAZLIKA za blokove novih otkupa | `modIzvestaj.SledBlokSumMapa:4963` | ◐ |
| 6 | završetak ispravke otpremnice iz F2 **uvek** ide u MANUAL: `OtpremnicaUpisi` predaje ID-eve (`"OTP-a + OTP-b"`, `modDokumenta:165/:210`) kao `newBroj`, a `CompleteOtpremnicaIspravka` traži po `BrojOtpremnice` | `modDokUnos:311` → `:1189` → `modStornoFlow:613/:625` ✔ | ✔, **nije reprodukovano** — testovi zovu `CompleteOtpremnicaIspravka` direktno sa brojem, pa put iz F2 nije meren |
| 7 | admin health check traži `Isplaceno`/`DatumIsplate`, obrisane u PR6 → FAIL na ispravnoj svesci | `modProductionHealthCheck:124` ✔ | ✔, nije pokrenuto |
| 8 | KPI „OM saldo" čita kolonu 5 = agro zaduženje, a saldo je kolona 6 (nezavisno od refaktora) | `modOtkupUI:7590-7591` naspram `modIzvestaj:845-846` ✔ | ✔ |
| 9 | testovi slaganja izveštaja (`T_Izv_SlaganjeOtkupOM`, `_SlaganjeKartica`, `_RangKooperanata`, `_ZbirniSadrzaj`) zeleni su jer porede zaglavlje sa zaglavljem — za nov dokument oba daju 0 | `modTest`, oko `:11988-13015` | ○ |

Produkcionih podataka nema, pa su ovo kvarovi dev sveske, ne kod korisnika.
Stavke 1–5 i 9 su upravo „konverzija čitalaca" koju je §14.1 imenovao kao
nepopisan posao; 6–8 su zatečeni i ne zavise od PR7.

#### ⚠ Mine za PR7 koje pre-flight od 12.09 nije video

- **Tvrdnje „prazno" postaju placebo kad se kolona obriše.** `LookupValue` za
  nepostojeću kolonu vraća `Empty` (`modDataAccess:634-636`) ✔, pa „OTK header:
  Kolicina prazna" i „Cutover: auto-link ne povezuje…" ostaju zelene i ne mere
  ništa. PR7 ih briše po imenu ili zamenjuje tvrdnjom o kanonskoj poziciji
  (`Pr3KanonskaPozicija(...) = 0`).
- **Kapije koje na obrisanu kolonu tiho propuštaju.**
  `FirstLiveOtpremnicaForBlocks` izlazi sa `""` kad kolone nema, a
  `OtkupBlockDeadParent*` preskače proveru (`modStornoFlow:2060`,
  `modStornoRecovery:267-315`) ○. Kolone se ne brišu pre nego što se ta mesta
  prepišu na članstvo — inače kapija nestane bez ijedne crvene tvrdnje.
- **Alat popisa je slep.** `tools/otkup_kolone_popis.py`, izvor brojeva iz §14.6,
  broji samo `COL_OTK_*` ✔ — ne vidi `COL_OTP_*`, literale imena kolona, omotače
  ni prosleđene indekse. Van njega je kritičar našao još 18 produkcionih mesta
  (health check, prefill ispravke, trace kolone otpremnice i njihov self-heal,
  izgubljeni blokovi, ručno „Poveži", izvoz za PWA) i 8 stavki van VBA
  (`make_fixture.py`, sidra sabotaža, GAS/PWA). Prag „dual READ = 0" meri se nad
  prilogom, ne nad tim alatom.
- **Preimenovanje trace kolona otpremnice** (`IspravkaOd` → `IspravkaOdID`) traži
  istu granu u `modSetup.EnsureSledljivostSchema` koju je PR6 dodao za `tblOtkup`,
  i u `make_fixture.py` `RENAME_COLS` — inače start vraća staro ime na kraj
  tabele ○.
- **`VremeUnosa`** nije svrstan ni u PR7 ni u PR8 (§14.6); piše ga samo legacy
  `SaveOtkup` ○.

#### Zastarele tvrdnje u planu

| Tvrdnja | Sada |
|---|---|
| §14.2: `GetKooperantiZaZbirnu` (`modPaletniList:2614`) je živ čitalac paletnog lista | **mrtva** — `Private`, nijedan poziv; živ čitalac je `GetOtkupiZaPalete` (`:983`, `:3171`) ✔ |
| §14.2: `StampajSledljivostZbirne` na `modIzvestaj:5966` | `:6072` ✔ |
| §14.2: `OtkupIdsByBrDok` na `modScrDokumenti:668` | `:712` ✔ |
| NALAZ 3: tri `Split(" + ")` nad otkup ID-evima | dva — `modAutoHladnjaca:199`, `modOtkupBlok:1460`; `modAmbalaza` ga više nema ✔ |
| CAPABILITY: poziv auto-hladnjače ugašen na `modOtkupUnos:388` | `:391` ✔ |
| ACCEPTANCE: A11 `tblOtkup` 9 → 1 | 8 → 1 ✔ |
| komentar `NapredakBlokaDostupan`: brojevi napretka se ne prikazuju | prikazuju se, iz praznih kolona — kvar 1 ✔ |
| LANDING: grana čeka #308 | merge-ovan 12.09 ✔ |

#### Odluke operatera (15.09.2026)

Paket preporuke je prihvaćen u celini. Uz svaku odluku stoji posledica koju PR7
mora da sprovede — ne samo izbor.

**1. Granica PR7/PR8 — opcija A: imenovana pauza zbirnih tokova do PR8.**

- Kapija `IzvedeniLanacIzPwaDostupan` se **deli** (8 postojećih mesta):
  `PwaOtpremnicaDostupna = True` u PR7, `PwaZbirnaDostupna = False` do PR8.
- Ručna zbirna F3 dobija **svoju** imenovanu pauzu, sa testom po imenu — poruka
  kaže „pauzirano do PR8“, ne „validacija nije prošla“.
- Zbirni čitaoci po `Otpremnica.BrojZbirne` (invarijanta, rekalkulacija, F8
  zbirne, integritet, izveštaji zbirne) se u PR7 **ne diraju**: nad novim
  podacima nemaju šta da rade i prelaze u PR8. Po §14.1 vode se kao
  **odloženi**, ne kao reuse.
- Otkupne grane pauziranih tokova — upisi u `Otkup.OtpremnicaID` / `BrojZbirne` /
  `VozacID` u hladnjačkom lancu, malina backfill-u i VOZ vezivanju — nestaju **već
  u PR7**, jer te kolone odlaze; zbirni deo tih tokova ostaje iza kapije ○.
- Propagacija ispravke u PR7 ide **otkup → otpremnica**; otpremnica → zbirna (H1)
  je PR8.
- Prihvaćen gubitak: između merge-a PR7 i PR8 za nove podatke nema zbirne,
  prijemnice, paleta, fakture ni sledljivosti zbirne. Produkcionih podataka nema,
  pa prozor postoji samo na `main`-u.

**2. Stari pisač otpremnice — samo za testove do PR8** (presedan `SaveOtkup_TX` iz PR6).

- `SaveOtpremnica_TX` i `SaveOtpremnicaMulti_TX` u PR7 gube produkcione pozivaoce
  (F2 `modDokUnos:269`, PWA `modMasterSync:919`); pauzirani hladnjački lanac nema
  produkcionog pozivaoca. Nov produkcioni pozivalac starog pisača je greška koju
  PR7 meri, ne podrazumeva.
- Brišu se u **PR8**, zajedno sa zbirnim tokovima.
- Po-klasne kolone `tblOtpremnica` (`Kolicina`, `Klasa`, `KolAmbalaze`, `BrutoKg`,
  `BrojZbirne`) zato **ostaju u kanonu do PR8** — nose ih test-only pisač i
  pauzirani zbirni čitaoci. Živi čitaoci otpremnice ih u PR7 više ne čitaju
  (prelaze na `tblOtpremnicaStavke`), pa je dual READ nad živim putem 0.
- Golden ostaje na test-only pisaču: **cilj PR7 je `RunGoldenSuite` 12/0 bez
  promene snapshot-a.** Nije izmereno — ako se snapshot ipak pomeri, to je nalaz
  koji se objašnjava, ne osvežava. Golden se menja jednom, u PR8 (scenario A4,
  gajbe Klase II u A2/F2/G1, registracija D1/H1/H2).
- Nov put PR7 dokazuje `RunBusinessFlowProSuite` nad `CreateOtpremnica*`, ne golden.

**3. Auto-lanac hladnjače — ceo pauziran do PR8.**

- Capability mapa: `MIGRATED` → **`PAUZIRAN` do PR8**. Polovičan prevod (samo
  korak otpremnice) bi kroz `ZavrsiVezuOtpremniceNaZbirnu` upisao `BrojZbirne` na
  nov header, a rekalkulacija bi nulirala zbirnu.
- Tvrdnja o pauzi (`Test_OTK_EkranPauziraAutoLanac`, testovi hladnjačkog lanca) u
  PR7 se **ne okreće** ○.
- GlobalGAP sledljivost zbirne (`REPLACED`) prelazi u PR8.

**4. Acceptance PR7 — usklađen odlukama 1–3.**

- PR7 dokazuje §14.2 tvrdnje **1, 2, 4, 5** i tvrdnju **3 suženu** na: veza
  otkup → otpremnica čita `tblOtpremnicaIzvori`, ne `Otkup.OtpremnicaID`.
- Tvrdnje **6 i 7** i **zbirni deo tvrdnje 3** (`TraceByZbirna`) prelaze u **PR8**.
  Tvrdnja 6 prelazi cela, a ne samo zbirni deo kako stoji u preporuci iznad:
  hladnjački lanac ostaje ceo pauziran.
- Edge H2 (`CorrectionSestre`) prelazi u PR8, usklađeno sa redom 8 tabele PR-ova.
- A11 `tblOtkup` **8 → 1** ostaje cilj PR7: svih 18 upisnih mesta vezne kolone
  pišu kolone koje PR7 briše.

**5. Kvarovi koji žive danas.**

| Kvar | Gde se rešava | Zašto |
|---|---|---|
| 2 pill plaćanja · 3 vrednost otkupa u izveštajima (saldo OM, kartica, otkupne liste, prosečna cena, zbirni OM — bez „roba po OM“, čiji manjak ide po otpremnici) · 9 testovi slaganja izveštaja | **mali PR pre PR7** — ✅ #334 (+ P1 iz review-a: jedan fail-closed prolaz za sve čitaoce vrednosti, §14.3) | čitaju Kolicina × Cena sa zaglavlja otkupa i ne zavise od otpremnice; vrednost je na `tblOtkupStavke` (obrazac `VrednostOtkupa`, §14.3) |
| 1 napredak bloka u F1 · 5 KG-RAZLIKA u sledljivosti | **PR7** | čitaju `Otkup.OtpremnicaID`, koji PR7 zamenjuje članstvom |
| 4 izvoz za PWA menadžment | **zaseban mali pre-flight** | menja oblik izvoznih redova koje čitaju GAS i PWA |
| 6 ispravka otpremnice iz F2 uvek u MANUAL | **PR7** | PR7 prepisuje F2 put; test koji vozi baš taj put ide u PR7 |
| 7 health check · 8 KPI „OM saldo“ | **zasebni mali zadaci**, paralelno | ne dodiruju refaktor |

#### Još otvoreno — pre F2 dela PR7

1. **Ambalaža otpremnice (`DOMAIN GAP`):** ko i kada knjiži izlaz gajbi
   stanica → vozač — pri nastanku drafta ili pri izdavanju.
2. **F2 prelaz:** gde ide `cenaII` kad zaglavlje nosi jednu `PredlogCena`; ostaje
   li kucani bruto; ko pravi radnju „Izdaj“ (`IzdajOtpremnicu_TX` nema pozivaoca).
3. **Zaglavlje bez stavki u testovima:** `SaveOtkup_TX` pada čim PR7 obriše
   kolone, pa `Test_OTK_VrednostBezStavkiPada` i `Test_OTP_StariOtkupNeUlazi`
   traže drugi put.
4. **PWA vozač** (`src/js/features/vozac/zbirna.js`): izvor liste otpremnica za
   zbirnu nije praćen.

#### Redosled do koda PR7

1. ✅ merge ovog pre-flight-a (#333);
2. ✅ mali PR: kvarovi 2, 3 i 9 — čitaoci vrednosti otkupa na stavke (#334), uz KPI „danas“;
   review P1 zatvoren u istom PR-u: dokumentski ugovor u jednom prolazu, bez
   grane „nema ključa = 0“ u ijednom čitaocu (§14.3);
3. 🟡 popis ponovo meren na novom `main`-u, sa nezavisnom proverom svih celina —
   posle koraka 2, jer on menja deo popisa (15.09 je provera stigla za 1 od 11);
   **16.09 (7b): popis premeren na `c2be85e8`, nezavisna provera stigla za 6 od 12 celina** — v. „7b“ ispod;
4. odluke iz „Još otvoreno“;
5. PR7 kod.

Paralelno i bez blokiranja: kvarovi 7 i 8, pre-flight za kvar 4.

#### 7b — popis premeren na `c2be85e8` (16.09.2026) — 🟡 nezavisna provera 6 od 12 celina

Popis (`docs/REFAKTOR_PR7_POPIS.md`) je ponovo izmeren posle #334 i sada ima alat: **`tools/popis_citalaca.py`** (samo čita
`src-vba`; osnovne grupe sidara iste kao 15.09, proširene `x_*` za literale, prosleđene indekse, `SaveOtkup*`, `VremeUnosa`,
trace kolone; graf poziva sa imenovanim ulaznim tačkama i kapijama pauze; red DUAL READ). Isti alat meri kraj PR7.

| Šta | Rezultat |
|---|---|
| delta #334 | osnovna sidra PROD 510 → **490** (−24 nestala, +4 nova ključa kolona u `RedoviZaTip`, sve u `otk_linija`); TEST 235 → 232; nijedna procedura popisa nije promenila status |
| P1 | alat + presude 15.09 vezane po sadržaju (486/490) + čitanje svega što je #334 dirao; kalibracija alata nad `a173c134`: 487/510 |
| P2 (slepi čitači) | **6 od 12 celina, 338 od 490 mesta**: saglasno 257, rešeno čitanjem koda **81** (19 rešenja R-NN), nerešeno **0**; pobednik P2 11 · 15.09 3 · treće 4 · P1 1. Jedno razrešenje prvog prolaza oboreno. |
| nije stiglo | `otp_brojzbirne` (62 mesta), `otp_pisci` (90 mesta), `testovi_1`–`testovi_3` (102 procedure) i kritičar pokrivenosti — limit sesije 16.09; te celine nose samo P1 |

**Kapija §14.1 za PR7 (zamenjeni pragovi):**

| Prag | Verdikt | Dokaz |
|---|---|---|
| jezgro ≤ 120% `CreateZbirna_TX` | **nije primenljiv kako je zapisan** | POP7-03: jedino čitanje koje ponavlja 695 je „unutar modula“, a ono za jednopotezni pisac otpremnice daje 1075 naspram 770 (+39,6%) — na greenfield skeli iz PR5, pre ijednog reda PR7. Po tekstu §14.1 (preko modula) Otkup bi 13.09 bio +50%. |
| spisak testova koji nestaju | **otvoren** | spisak 15.09 važi jednoprolazno; P2 test celina nije stigla. P1: POP7-09, placebo `Test_OTK_CitaociCitajuStavke` |
| populacija na baznom commitu | **izmerena** | 28 produkcionih modula, 26 sa živim mestom; tabela po modulu u prilogu |
| dual READ = 0 | **baseline** | 155 referenci u produkciji; **102 živih u 49 procedura** (ZIV_UI 98, ZIV_MAKRO 4), PAUZIRAN 7, SAMO_TEST 12, MRTAV 34; prag se dokazuje alatom i prepisom mesta, ne zelenom suite-om posle brisanja kolone (POP7-10) |
| A11 | cilj 8 → 1 stoji | `tblOtkup` 8 pisaca; `Otkup.OtpremnicaID` 6 (POP7-08) |

**Zatečeni nalazi:** NALAZ 1 (6 pisaca) **potvrđen** (POP7-08) · `OtkupIdsByBrDok` **potvrđen**, sada AUD-057
(POP7-06) · pauza `NapredakBlokaDostupan` samo u mrtvom panelu **potvrđena** u oba prolaza (POP7-07) ·
`SaveOtkup_TX` **potvrđen i dopunjen**: 6 testova, a #334 već pokazuje drugi put za zaglavlje bez stavki (POP7-09).

**Novo za odluku operatera pre PR7 koda:**

1. **Prag jezgra** (POP7-03): zapisati ulaz (jednopotezni ili svi ulazi skele) i čitanje, ili zameniti prag rastom skele u
   PR7 (npr. ≤ 20% naspram današnjih 1075/1373).
2. **Odluke 4 i 5 se sudaraju** (POP7-15) — tri nezavisna čitača: `TraceByZbirna` je PR8, a spaja otkupe po
   `Otkup.OtpremnicaID` koju PR7 briše; posle brisanja PDF sledljivosti zbirne kaže „NEMA“, a paletni list tiho gubi kooperante.
   PR7 mora bar da prevede taj spoj na `tblOtpremnicaIzvori` ili da izričito pauzira te izlaze.
3. **Završetak 7b:** P2 za `otp_brojzbirne` (62 mesta), `otp_pisci` (90 mesta), `testovi_1`–`testovi_3` (102 procedure) i kritičar pokrivenosti. Ulazi i brief su spremni; ništa od ovoga ne zahteva nov kod.

Živi kvarovi van PR7 dobili su redove u `docs/KNOWN_ISSUES.md` §8.10: AUD-055 (kvar 7, POP7-04), AUD-056 (kvar 8,
POP7-05), AUD-057 (POP7-06). Klasifikacija 15.09 je starija od odluka operatera; u popisu je 19
rešenja sa imenovanim pobednikom, a zamene koje protivreče odlukama 2/3 su označene (POP7-02).

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
- „`tblZbirnaIzvori` je jedini zapis pripadnosti; kolone `Otpremnica.ZbirnaID` nema, a „gde je sada" se računa iz članstva."
- „PRJ-789 je jedna prijemnica bez obzira ima li jednu ili dve klase."
- „Faktura stavka zna tačnu `PrijemnicaStavkaID`."
- „Storno prima `DocumentID`. Štampa prima `DocumentID`. Invarijanta prima `ZbirnaID`."
- „Broj dokumenta je labela."
- „`GeneracijaID` ne postoji."
- „Sveska se može obrisati — kod je vrati."

Sve dok bilo koja od ovih rečenica nije tačna, refaktor nije završen.
