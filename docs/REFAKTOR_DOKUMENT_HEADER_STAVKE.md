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
| 7 | 🟡 **pre-flight 15.09 (§14.7) — granica A, pauze i test-only pisci oboreni 16.09 (§14.7 „Odluke operatera 16.09“): legacy se ne čuva živim, čuva se mapa sposobnosti. Pre koda: mali PR za kvarove 2/3/9 (✅ #334), ponovljen popis sa proverom (🟡 7b 16.09: premeren, nezavisno 6/12 celina), odluke o F2.** **Otpremnica cutover**: `tblOtpremnicaIzvori` pokazuje na prave `OtkupID`-eve; propagacija ispravke naniže; panel prelazi na `GetOtpremnicaProgress`; **briše `Otkup.OtpremnicaID`** sa svih **6** pisača (ne 5 — v. PR7 pre-flight, NALAZ 1); **rename `Cena` → `PredlogCena`** sa čitaocima (§13b) | 6 |
| 8 | ~~**Zbirna cutover**: invarijanta preko `tblZbirnaIzvori` (sada nad **pravim** `OtpremnicaID`-evima), `StornoZbirna_TX(id)`, storno otpremnice po §7.1, **propagacija ispravke = nova verzija (A13)**, print, izveštaji. **Briše `ZbirnaIdent*`, `ZbirnaGeneracija*` i mrtvu `RunSimpleStornoOtpremnica`.** Registruje goldene D1, H1, H2. **Iz PR7 preuzima (odluka 15.09, §14.7 — oboreno 16.09, slajsovi se seku po novom modelu):** §14.2 tvrdnje 6, 7 i zbirni deo 3, edge H2, podizanje pauze zbirnih tokova (F3, malina, VOZ) i auto-lanca hladnjače, brisanje test-only `SaveOtpremnica*` i po-klasnih kolona `tblOtpremnica`, izmenu golden scenarija A4~~ — **zamenjeno §14.9** | 7 · **§7.1, A13–A15 odlučeni** |
| 9 | ~~**Prijemnica** header+stavke + izvori + cutover~~ — **zamenjeno §14.9** | 8 |
| 10 | ~~**Faktura**: `FakturaStavka.PrijemnicaStavkaID`~~ — **zamenjeno §14.9** | 9 |
| 11 | ~~**Paleta**: `PaletaStavka.PrijemnicaStavkaID`~~ — **zamenjeno §14.9** | 9 |
| 12 | ~~**Sledljivost kao graf** nad eksplicitnim FK; ukloniti heuristički AutoLink~~ — **zamenjeno §14.9** | 11 |
| 13 | ~~**E2E + brisanje**: `COL_GENERACIJA_ID`, `COL_DETE_ZBIRNA_GEN`, `*ByBroj_TX`, svih 10 `Split(" + ")`, mrtvi testovi i sabotaže; pravila `NEMA_GENERACIJE` / `NEMA_BROJA_KAO_FK` / `NEMA_ID_PLUS_ID`; `ZBR_IDENTITET.md` → superseded~~ — **zamenjeno §14.9** | 12 |
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

#### Odluke operatera (15.09.2026) — premisa oborena 16.09 (v. „Odluke operatera 16.09“)

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

#### Redosled do koda PR7 (15.09 — zamenjen 16.09)

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

#### 7b — popis premeren na `c2be85e8` (16.09.2026) — zamenjen odlukama 16.09

> **Zamenjeno istog dana** (v. „Odluke operatera 16.09“ ispod): PR/pauza klasifikacija i preostale provere se ne dovršavaju. Ostaju alat, spisak mesta kao spisak za brisanje, AUD-055..057 i nalazi koji opisuju sposobnosti.

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

#### Odluke operatera (16.09.2026) — legacy je samo mapa sposobnosti

> Obaraju premisu odluka od 15.09 (1, 2, 3). Nema produkcije ni podataka koje treba štititi; radi se iznova.

**Pravilo.** Legacy kod ne mora da radi dok novi kod na novom modelu ne proradi, i ne pravi se ništa što ga čuva između faza:
imenovane pauze, podela kapija, test-only pisci, kolone ostavljene u kanonu za pauzirane čitaoce, mostovi, dvojni putevi.
**Apsolutno se čuva samo mapa sposobnosti:** svaka sposobnost koju operater danas ima (radnja, ekran, izveštaj, PDF, izvoz,
sync, makro) mora imati mesto u novom modelu i PR koji je vraća.

**Jedno tehničko ograničenje:** posle svakog PR-a projekat mora da se kompajlira — jedan modul koji se ne kompajlira obara ceo
`run_vba`, pa se ni novi kod ne može testirati. Legacy se zato **briše**, ne ostavlja polomljen.

| Odluka 15.09 | 16.09 |
|---|---|
| 1 — granica A: zbirni tokovi imenovano pauzirani do PR8, kapija se deli, F3 sa svojom pauzom i testom | **oborena** — bez pauza i podele kapije; sposobnost koja pukne vodi se u mapi kao „prekinuto do PRn“, ne u kodu |
| 2 — `SaveOtpremnica*` samo za testove do PR8, po-klasne kolone ostaju u kanonu, golden nepromenjen | **oborena** — stari pisci i po-klasne kolone se brišu u slajsu otpremnice; testovi legacy mehanizma se brišu; golden se preformuliše na novom modelu, do tada je imenovano isključen |
| 3 — auto-lanac hladnjače ceo pauziran do PR8 | **oborena** — lanac je stavka mape; kod i aparat pauze (ugašen poziv, poruke o pauzi, fail-closed pending relink) se brišu, lanac se gradi nad novim modelom |
| 4 — tvrdnje §14.2 raspoređene PR7/PR8 po granici pauze | tvrdnje ostaju ugovor novog modela; raspoređuju se po slajsu novog modela |
| 5 — kvarovi 1/5/6 u PR7, kvar 4 zaseban pre-flight | kao kvarovi nevažni (nema operatera); ostaju kao stavke mape. AUD-055..057 su tačni, nisu hitni |

**Kapija §14.1 za sledeći slajs.** Prag jezgra se **uklanja** — merio je trošak rada u mestu naspram rada iz nule, a novi model se
ionako piše iznova (POP7-03). „Dual READ = 0“ ostaje u jačem obliku: kolone starog modela izbačene iz kanona, njihove konstante
obrisane (kompajler meri), `tools/popis_citalaca.py` bez preostale reference. „Spisak testova koji nestaju“ postaje: za svaki test
legacy mehanizma u mapu se upisuje **ishod** koji je tvrdio, a test se briše. Sukob odluka 4 i 5 (POP7-15) više ne postoji.

**7b je zamenjen.** `docs/REFAKTOR_PR7_POPIS.md` ostaje spisak za brisanje i istorijat; PR/pauza klasifikacija i preostale provere
(`otp_brojzbirne`, `otp_pisci`, test celine, kritičar u obliku 15.09) se ne dovršavaju — kritičar se preliva u mapu sposobnosti.
#334 nije bačen rad: čitaoci vrednosti otkupa sa stavki su već novi model.

**Novi redosled do koda:**

1. **mapa sposobnosti** (`docs/DOMEN/MAPA_SPOSOBNOSTI.md`) — prikuplja se spolja po promptu, spaja i proverava pokrivenost ulaznih tačaka alatom;
2. **odluke domena** koje traže sposobnosti: ambalaža otpremnice (`DOMAIN GAP`), `cenaII` naspram jedne `PredlogCena`, kucani bruto, ko pravi radnju „Izdaj“, PWA vozač i otpremnica;
3. **slajsovi po novom modelu** — svaki briše legacy svog dela (pisce, kolone, čitaoce) i ostavlja projekat koji se kompajlira; otpremnica i zbirna smeju zajedno, jer je razlog da budu odvojene (živa zbirna između njih) nestao; tabela PR-ova se piše ponovo;
4. kod.

---

### 14.8) Odluke domena posle mape sposobnosti (17.09.2026)

Ulaz: `docs/DOMEN/MAPA_SPOSOBNOSTI.md` (387 sposobnosti). Odlučio operater; oznaka **[podrazumevano]** znači da
nije posebno pitano i važi dok operater ne kaže drugačije.

| # | Pitanje | Odluka | Posledica za slajs |
|---|---|---|---|
| 1 | Ambalaža otpremnice (`DOMAIN GAP`): kada se knjiži izlaz gajbi stanica → vozač | **Pri izdavanju.** Količina = zbir `KolAmbalaze` sa stavki izdate otpremnice | `TrackAmbalaza` ide u `IzdajOtpremnicu_TX` / `CreateOtpremnicaIzIzvora_TX`, ne u nacrt; izmena nacrta ne dira `tblAmbalaza` |
| 2 | `cenaII` kad zaglavlje nosi jednu `PredlogCena` | **Predlog cene po klasi na stavci otpremnice**; zaglavlje nema cenu | očekivana stavka: `Klasa`, `Kolicina`, `KolAmbalaze`, `PredlogCena`; prefill otkupa uzima cenu klase; `Otpremnica.Cena`/`PredlogCena` se briše sa zaglavlja |
| 3 | Kucani bruto | **Ostaje, `BrutoKg` na stavci** otkupa, otpremnice i prijemnice | bruto→neto (tara po gajbici) po klasi na stavci; isti prekidač `OTKUP_BRUTO_UNOS`; `brutoKgI/II` sa zaglavlja nestaju |
| 4 | Ko izdaje otpremnicu (`IzdajOtpremnicu_TX` nema pozivaoca) | **Operater dugmetom „Izdaj“** kad je ostatak 0; auto-lanac i PWA koriste jednopotezni `CreateOtpremnicaIzIzvora_TX` | nova radnja na ekranu DOKUMENTI; nacrt se menja do izdavanja |
| 5 | Dodela vozača iz PWA (E-058; danas `vozacID` na redu otkupa, završava kao terminalni `Duplicate`) | **Dodela pravi otpremnicu**: pri uvozu na desktop postaje izdata otpremnica (stanica, vozač, izabrani otkupi) | PWA čuva dodelu kao zapis dodele, ne kao polje otkupa; uvoz zove `CreateOtpremnicaIzIzvora_TX` |
| 6 | Zbirna vozača iz PWA (E-044, E-063, E-023) | **Vozač vidi svoje izdate otpremnice; zbirna iz PWA stiže kao dokument** sa izvorima = te otpremnice | GAS servira otpremnice po `VozacID` otpremnice (ne `OTK-*` redove); uvoz `VOZ-*` gradi zbirnu sa `tblZbirnaIzvori` |
| 7 | Banka: blok raspodele uplate je poslovni broj otkupa (D-035..D-037) | **Broj samo na nalogu, veza po ID-u**: poziv na broj ostaje poslovni broj, pri mapiranju se jednom razreši u `OtkupID` | mapiranje i otvoreno-po-bloku rade nad ID-em; nema traženja po broju posle razrešenja |
| 8 | ANALIZA / marža (FM-0106) — ekran prazan, marža ne postoji | **Posle refaktora** (ispravljeno 17.09.2026; prvi odgovor je bio „u slajsu fakture“) | nije obaveza nijednog slajsa; gradi se kao nova funkcija nad novim modelom kad svi dokumenti pređu |
| 9 | Makroi bez provere prava (`SetupNewPC`, `RunSelfUpdate`, `PublishReleaseToDrive`, `RollbackReleaseTo`, `OcistiTabele`, `MigrirajPodatkeIzStarog`, `OpenExcel`/`CloseExcel`) | **Ostaje kako jeste** — Alt+F8 je alat održavanja | nema brane u makroima; brana ostaje na ekranima |
| 10 | Provera zdravlja sa ručnim spiskom kolona (`Check_CoreTablesAndColumns`) | **[podrazumevano]** spisak kolona iz kanona `schema.json`, kao `Check_SchemaRegistry` | ide u prvi slajs koji briše kolonu; zatvara i AUD-055 |
| 11 | Provere integriteta starog modela (15 od 22) | **[podrazumevano]** brišu se sa starim modelom; slajs dokumenta dodaje proveru svoje invarijante (zaglavlje = zbir stavki, izvori aktivni) | nema „prevođenja“ starih provera |
| 12 | Popravke podataka starog modela (F-061, F-090..F-093) | **[podrazumevano]** brišu se, nisu sposobnost (pravilo „bez migracija i backfill-a“) | v. `MAPA_SPOSOBNOSTI.md` „Nije sposobnost“ |
| 13 | Oblik otkupa između desktopa i PWA (OTK sheet, `OtkupiAll`) — PWA red nosi jednu klasu | **Zaglavlje + zaseban tab stavki; PWA zapis dobija niz stavki.** PWA i GAS se **ne diraju u S1** — ostaje instrukcija (§14.10) | VBA izvoz/push prelazi na nov oblik u S1c; GAS/PWA u S5 |

**Ostaje otvoreno:** nijedna odluka domena iz liste „Još otvoreno“ (§14.7). Tačka 3 te liste (testovi sa zaglavljem
bez stavki) nije domen nego posao slajsa otkupa. Sledeći korak: nova tabela PR-ova (slajsova) po novom modelu, sa
mapom kao spiskom obaveznih ishoda i ovim odlukama kao ugovorom.

### 14.9) Nova tabela slajsova (17.09.2026) — zamenjuje redove 7–13 u tabeli PR-ova

**Pravila slajsa** (iz §14.7 „Odluke operatera 16.09“ i §14.8):

- Slajs = jedan dokument (ili jedna veza) **do kraja**: pisac, čitaoci, ekran, štampa, izveštaji, storno/ispravka,
  sync i izvozi tog dokumenta. Na kraju slajsa kolone i procedure starog modela koje on pokriva su **obrisane**.
- Posle svakog PR-a projekat se kompajlira. Legacy kod koji slajs ne prenosi a koji bi pukao se **briše**, ne
  pauzira; sposobnost koju je nosio ostaje zapisana u mapi i u koloni „Sadržaj“ piše koji je slajs vraća („vraća Sx“).
- **Prag slajsa je merljiv:** `python tools/popis_citalaca.py` — grupe iz kolone „Briše“ imaju **0** PROD mesta
  (danas izmereno na `main` `0e6e3acd`, PROD bez MRTAV). DUAL READ = 0.
- **Spisak sposobnosti po slajsu je ulaz za pre-flight**, izveden iz `MAPA_SPOSOBNOSTI.md` (redovi sa presudom
  da/delimično, 120 ukupno) po tabeli i ekranu. Pre-flight slajsa ga proverava red po red i dopisuje koje su
  sposobnosti vraćene i kojim testom; nijedna ne sme ostati bez slajsa.
- Svaki slajs dodaje proveru pravila svog dokumenta (§14.8 t. 11), ne prevodi stare provere integriteta.

| # | Slajs | Sadržaj | Odluke §14.8 | Briše (prag = 0) | Sposobnosti (ulaz za pre-flight) | Zavisi od |
|---|---|---|---|---|---|---|
| S1 | **Otkup do kraja** (stavke su jedini izvor) | svi čitaoci linijskih polja otkupa prelaze na `tblOtkupStavke`: izveštaji, panel bloka, storno prefill, KPI (AUD-056), štampa, izvozi `ExportOtkupPoOM` / `ExportOtkupiAll` (deo otkupa) / `ExportSaldoOMDetail`, push OTK redova (`BuildOTKSheetRowForOtkup`), GAS/PWA pregled otkupa i menadžment; provera zdravlja iz kanona (AUD-055) | 3, 10 | `otk_linija` (91), `x_saveotkup` (35, stari `SaveOtkup_TX`), `x_vreme_unosa` (3), `x_indeks` (2), `x_literal` u health (13); kolone `Otkup.Kolicina/Cena/Klasa/KolAmbalaze/...` | A-015, B-036 (otkup), C-019, D-057, E-024, E-026, E-027, E-030, E-039, E-040, E-041, E-042, E-053, E-056, E-057, E-070, E-072, E-074, E-075, E-076, E-077, F-089 | PR6 |
| S2 | **Banka po ID-u** | mapiranje izvoda i nalozi: poziv na broj se jednom razreši u `OtkupID`, blok i otvoreno po bloku rade nad ID-em | 7 | traženje otkupa po poslovnom broju u `modBankaMapiranje` / `modNovac` | D-032, D-033, D-035, D-036, D-037 | S1 |
| S3 | **Otpremnica cutover — desktop** | F2 kroz nacrt + očekivanje po klasi (sa `PredlogCena` i `BrutoKg` na stavci) + dugme „Izdaj“; ambalaža pri izdavanju; panel blokova i izgubljeni blokovi nad `tblOtpremnicaIzvori`; storno i ispravka otpremnice; auto-hladnjača i malina auto-otpremnica kroz `CreateOtpremnicaIzIzvora_TX`; štampa i izveštaji otpremnice. PWA/sync putevi koji čitaju `Otkup.OtpremnicaID/VozacID` se **brišu** (vraća S5) | 1, 2, 3, 4 | `otp_linija` (101), `otp_cena` (9), `otp_stari_pisac` (68, `SaveOtpremnica*`, AutoLink), `otk_veze` za `OTPREMNICA_ID` / `VOZAC` / `BROJ_OTPREMNICE`; kolone `Otkup.OtpremnicaID/VozacID/BrojOtpremnice`, linijska polja i `Cena` na `tblOtpremnica` | A-001, A-002, A-011, A-012, A-014, A-018..A-029, B-004, B-005, B-010, B-013, B-014, B-022, B-023, B-027, B-032, B-033, B-039, B-041, B-042, B-046, B-047, C-005, C-008, C-010, C-025, C-059 | S1 |
| S4 | **Zbirna cutover** (stari PR8) | F3 kroz `CreateZbirnaIzIzvora_TX`; storno zbirne i propagacija ispravke = nova verzija (A13); malina auto-zbirna; štampa i izveštaji zbirne; relink u storno toku nad ID-em | 3 | `otp_brojzbirne` (56), `otk_veze` za `BROJ_ZBIRNE`, `x_trace` (12), `GeneracijaID`, `ZbirnaIdent*`, `*ByBroj_TX` zbirne | A-017, A-030, B-001, B-006, B-012, B-015, B-024, B-026, B-028, B-038, B-044, C-004, C-007, C-011, C-016, C-024, C-038, C-060 | S3 |
| S5 | **PWA i sync na novom modelu** | PWA/GAS otkup kao zaglavlje + stavke po instrukciji §14.10; dodela vozača iz PWA → izdata otpremnica; GAS vozaču servira otpremnice po `VozacID` otpremnice; `VOZ-*` zbirna → zbirna sa izvorima; auto-lanac bez kapije pauze; badge sync-a bez „degradirano“ grane | 5, 6 | `pauza` (13), `IzvedeniLanacIzPwaDostupan`, `NapredakBlokaDostupan`, `TryUpdateVozacID` na otkupu | E-001, E-003, E-019..E-023, E-035, E-044, E-058, E-063, E-064, E-065 | S4 |
| S6 | **Prijemnica** header + stavke + izvori + cutover | F4; hladnjača auto-prijemnica po zbirnoj (ne po `BrojZbirne\|Klasa`); ambalaža vraćena; štampa, izveštaji, izvoz | 3 | linijska polja `tblPrijemnica`, `Prijemnica.BrojZbirne` kao veza, `split_plus` (7) | A-031, B-029, B-045, C-006, C-014, C-015, C-033, C-039 | S4 |
| S7 | **Faktura** | `FakturaStavka.PrijemnicaStavkaID`; SEF; kartica kupca | — | veza fakture preko broja prijemnice | C-003, D-031 i redovi D1 koji čitaju prijemnicu (pre-flight) | S6 |
| S8 | **Paleta i prerada** | `PaletaStavka.PrijemnicaStavkaID`; usklađivanje paleta po ID-u prijemnice | — | `PaletaStavka.BrojZbirne` kao veza, broj prijemnice kao veza | C-040, C-048, C-050, C-051 | S6 |
| S9 | **Sledljivost kao graf + brisanje** | sledljivost nad eksplicitnim FK; provere pravila novog modela umesto 15 starih provera integriteta; pravila `NEMA_GENERACIJE` / `NEMA_BROJA_KAO_FK` / `NEMA_ID_PLUS_ID`; 49 pomoćnih procedura iz „Pokrivenost“ u mapi potvrđeno 0; `ZBR_IDENTITET.md` → superseded | 11 | ostatak `x_literal` (integritet), sve grupe popisa = 0 | C-029, C-031, C-032, C-035, C-036, C-037, C-041, C-043, C-044, F-087, F-088 | S7, S8 |

**Van refaktora:** marža i ekran ANALIZA (§14.8 t. 8). **Nije slajs:** makroi bez brane (§14.8 t. 9).

**Sledeći korak:** pre-flight S1 (skill `pre-flight`), jedna sesija.

### 14.10) Pre-flight S1 „Otkup do kraja“ (17.09.2026)

Mereno na `main` `0a2f5e5b` (kod isti kao `0e6e3acd`), `popis_citalaca.py --json`, grupe `otk_linija`,
`x_saveotkup`, `x_vreme_unosa`, `x_indeks`, `x_literal` u `modProductionHealthCheck`: **163 PROD mesta u 56
procedura** (ZIV_UI 86 · ZIV_MAKRO 7 · PAUZIRAN 5 · SAMO_TEST 46 · MRTAV 19) + **67 TEST mesta** (najviše
`modBusinessFlowProTests` 32, `modTestStornoCentar` 16, `modTestBanka` 8).

#### Glavni nalaz — čitaoci danas čitaju prazna polja

Jedini pisac otkupa `CreateOtkup_TX` → `BuildOtkupHeaderRowData` (`modOtkup.bas:1379`) na zaglavlje piše
`TipAmbalaze` i `KolAmbIzdata`, a **ne** `Kolicina`, `Cena`, `Klasa`, `KolAmbalaze`, `BrutoKg`, `Novac`,
`PrimalacNovca`, `VremeUnosa` — te kolone su i dalje u kanonu (`schema/schema.json` `tblOtkup`). Svaki živi čitalac
iz spiska ispod zato za svaki nov otkup danas vidi prazno/0 (npr. `modStammdatenSync.ExportOtkupPoOM:554-557`,
`modOtkupBlok.SumKolByOtp:1584`, `modScrDokumenti.ColKolicina:1348` za mod OTKUP). To je dozvoljeno stanje po
§14.7 (legacy ne mora da radi), ali znači da S1 ne „čuva“ ništa — **vraća** sposobnosti.

#### Ispravka praga iz §14.9

Grupa `otk_linija` meša činjenice zaglavlja i stavke. `COL_OTK_TIP_AMB` (8 živih) i `COL_OTK_KOL_AMB_IZDATA`
(6 živih) su **H** po `DOCUMENT_HEADER_LINES.md` (redovi „`VrstaVoca`, `SortaVoca`, `TipAmbalaze`“ i
„`KolAmbIzdata`“) i ostaju. Prag S1 je: **`COL_OTK_KOLICINA/CENA/KLASA/KOL_AMB/BRUTO/NOVAC/PRIMALAC/VREME_UNOSA`
i `SaveOtkup(_TX)` = 0**, a konačni dokaz je da su te konstante **obrisane iz `modConfig`** i kolone iz kanona —
kompajler je tada checker. Alat `popis_citalaca.py` dobija podelu grupe (`otk_stavka` / `otk_header`); izmena
checkera traži dokaz u oba smera (CLAUDE.md §5).

#### Verdikt po osama

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | **PROVEN** | otkup = zaglavlje + 1..2 stavke (`DOCUMENT_HEADER_LINES.md` §4.1); količina, cena, klasa, gajbe, bruto su **L**; `Novac`/`PrimalacNovca` se brišu (§4.1b); bruto na stavci (§14.8 t. 3) |
| `IDENTITY` | **PROVEN** za desktop · **GAP rešen odlukom** za PWA | desktop: `OtkupID` putuje, stavke po `OtkupStavkaID`. PWA red (`BuildOTKSheetRowForOtkup`, `modStanicaLock.bas:548`) nosi identitet po otkupu i JEDNU klasu → odluka §14.8 t. 13 |
| `CARDINALITY` | **PROVEN** | 1 otkup : 1..2 stavke; čitaoci po klasi (izvoz po Stanica+Vrsta+Klasa, storno prefill, štampa) moraju iterirati **stavke**, ne zbir |
| `INVARIANTS/OWNER` | **PROVEN** | „otkup bez stavki pada“ drži `StavkeOtkupaRedovi` (`modOtkup.bas:973`) i `ZbirStavkiZaOtkup` (`modOtkup.bas:1252`), iz #334 |
| `WRITERS` | **PROVEN** | jedini produkcioni pisac `CreateOtkup_TX` (`modOtkup.bas:71`); `SaveOtkup`/`SaveOtkup_TX` su SAMO_TEST (35 mesta u `modOtkup`, pozivaoci u testovima) — brišu se, testovi prelaze na `CreateOtkup_TX` |
| `DOWNSTREAM` | **PROVEN** (spisak) | 56 procedura ispod; nizvodno van VBA: OTK sheet i `OtkupiAll` → GAS `getOtkupiForOtkupac:2768`, `mergeOtkupRows_:1521`, PWA pregled/kartica/menadžment (E-039..E-042, E-053..E-057, E-070..E-077) |
| `CAPABILITY` | **PROVEN** (spisak) | 22 sposobnosti S1 provereno po presudi u mapi — v. dole |
| `ACCEPTANCE CONTRACT` | **plan** | v. dole |
| `PLATFORM` | N/A | nema novog Excel/COM ponašanja |
| `LANDING` | **rizik: šema** | brisanje kolona menja kanon i `modSchema` → S1d ide serijski, sam; ostali PR-ovi S1 ne diraju šemu |

`EVENTS: N/A` — S1 ne menja kada nastaje otkup, ambalaža ni novac; menja samo odakle se čitaju činjenice koje već postoje.

#### Reuse (ne pisati nov sloj)

`modOtkup.StavkeOtkupaRedovi:973` (redovi stavki, pada na otkup bez stavki) · `ZbirStavkiPoOtkupu:1180` →
`Array(kg, vrednost, gajbe, klase)` · `ZbirStavkiZaOtkup:1252` (nedostajući ključ = greška) ·
`VrednostOtkupa:874`. Već ih koriste `modIzvestaj` (`:595`, `:906`, `:1270`, `:1544`). Čitalac po klasi ili po
bruto koristi `StavkeOtkupaRedovi`; zbir po dokumentu koristi `ZbirStavkiPoOtkupu`.

#### Sposobnosti S1 — provera

- **Potvrđeno S1 (20):** B-036 (storno prefill po klasi), C-019, E-024, E-026, E-027, E-030, E-039, E-040, E-041,
  E-042, E-053, E-056, E-057, E-070, E-072, E-074, E-075, E-076, E-077, F-089 — sve čitaju linijska polja otkupa.
- **Potvrđeno S1 uz dopunu praga:** A-015 i D-057 zavise od `Split(ID, " + ")` (`modScrDokumenti.bas:725`,
  `modAmbalaza.bas:343`) — to je više `OtkupID`-ova istog broja iz modela „red po klasi“. Deo `split_plus` koji
  nosi **otkup** (`modPrint`, `modOtkupBlok`, `modAmbalaza`, `modScrDokumenti.OtkupIdsByBrDok`) ide u S1; deo
  otpremnice/prijemnice (`modDokUnos`, `modAutoHladnjaca`) ostaje u S3/S6.
- **Premešteno:** nijedna.
- **PWA/GAS strana** (E-039..E-042, E-053..E-057, E-070..E-077): po §14.8 t. 13 S1 **ne dira** `gas/` ni `src/`;
  VBA izvoz prelazi na nov oblik, a PWA/GAS prilagođavanje je instrukcija ispod i posao S5.

#### Mesta bez sposobnosti S1 (čija su)

- **MRTAV — briše S1a:** `modDokumenta.GetStorniraniByTip`, `modHelpers.CheckVerwaisteDokumente`,
  `modMarza.AggregateOtkupByVrsta(Filtered)` (marža je van refaktora, §14.8 t. 8), `modOtkupBlok.BuildFirstBlokCena`,
  `modOtkupBlok.PrefillOtkupFromStornirano`, `modScrDokumenti.ColumnSpec`.
- **Otpremnica (S3), ali čitaju količinu otkupa:** `modOtkupBlok.SumKolByOtp/SumAmbByOtp/SumBrutoByOtp/BuildNapisanoByOtp/ExistingBlokCena/LoadBlokovi/RenderSpec`,
  `modDokumenta.GetLostOtkupBlokovi`, `modSledljivost.AutoLinkOtkupOtpremnica/GetUnlinkedOtkupi/TraceByZbirna`,
  `modStornoFlow.GetStornoBlockRows`, `modPaletniList.GetOtkupiZaPalete`, `modMasterSync.AutoCreateOtpremniceFromPWA`
  (PAUZIRAN). S1 im menja **samo izvor količine/cene/klase** (stavke); vezu `Otkup.OtpremnicaID` ne dira — to je S3.
  Pauzirani `AutoCreateOtpremniceFromPWA` se **briše** (vraća S5).
- **Šema/health:** `modSetup.EnsureDoradeSchema` (dodaje stare kolone — briše se deo za otkup),
  `Check_CoreTablesAndColumns`, `Check_GoogleSyncMasterSchema` → spisak iz kanona (§14.8 t. 10). Literali
  `Check_OtkupOtpremnicaCrossZbirnaLinks` i `Check_DocumentSoftDeleteReferences` su veze (`OtpremnicaID`/`BrojZbirne`) → S3/S4.

#### Podela S1 na PR-ove (svaki se kompajlira)

| PR | Sadržaj | Prag posle PR-a |
|---|---|---|
| **S1a** | brisanje `SaveOtkup`/`SaveOtkup_TX` i MRTAV procedura iznad; testovi koji grade otkup prelaze na `CreateOtkup_TX` (fixture helper jedan, u `modTest`); `popis_citalaca` podela grupe `otk_stavka`/`otk_header` sa dokazom u oba smera | `x_saveotkup` = 0; MRTAV mesta S1 = 0 |
| **S1b** | desktop čitaoci na stavke: `modScrDokumenti` (kolone moda OTKUP, `RedoviZaTip`, `RowsBlokovi`, `OtkupIdsByBrDok` bez `" + "`), `modStornoDok` prefill po stavkama, `modStornoZurnal.OtkupReissueDupExists`, `modPrint.FillOtkupSablon`, `modIzvestaj` (`ReportKarticaKooperanta`, `ReportOtkupRobaOM`, `ReportSledljivost*` + `x_indeks`), `modAmbalaza` početno stanje po `OtkupID`, izvor količine u `modOtkupBlok`/`modDokumenta`/`modSledljivost`/`modStornoFlow`/`modPaletniList`; KPI `SaldoOMUkupno` (AUD-056) | `otk_stavka` živih u tim modulima = 0 |
| **S1c** | sync i izvozi, samo VBA: `ExportOtkupPoOM`, `ExportOtkupiAll` (deo otkupa), `ExportSaldoOMDetail`, push `BuildOTKSheetRowForOtkup` → **nov oblik** (zaglavlje + stavke, §14.8 t. 13); `PwaIstiSadrzaj` poredi stavku; brisanje pauziranog `AutoCreateOtpremniceFromPWA`; health iz kanona (AUD-055) | `otk_stavka` = 0 u celom PROD; `x_literal` health za otkup = 0 |
| **S1d** | brisanje kolona `Kolicina, Cena, Klasa, KolAmbalaze, BrutoKg, Novac, PrimalacNovca, VremeUnosa` iz `tblOtkup` u kanonu + `gen_schema_module.py` + brisanje konstanti `COL_OTK_*` iz `modConfig`; `EnsureDoradeSchema` bez tih kolona | **kompajlira se bez tih konstanti** = dokaz nula čitalaca |

#### Acceptance contract (plan dokaza)

- **Važi posle S1:** otkup sa dve klase (I 100 kg × 50, II 40 kg × 30) se u mreži DOKUMENTI, štampi, izveštaju
  otkupa po OM, kartici kooperanta i izvozu `OtkupPoOM` vidi kao dve klase sa tačnim kg i vrednošću 6200 —
  testovi u `modBusinessFlowProTests` nad `CreateOtkup_TX`, po jedan po čitaocu grupe (mreža, štampa, izveštaj, izvoz).
- **Ostaje netaknuto:** `TipAmbalaze` i `KolAmbIzdata` na zaglavlju; vrednost otkupa u novcu/banci (#334 testovi zeleni);
  `Otkup.OtpremnicaID/VozacID/BrojZbirne` (S3/S4).
- **Edge:** otkup sa jednom klasom; storno prefill otkupa sa dve klase vraća obe stavke; ponovna štampa starog broja.
- **Negativan:** otkup bez stavki u bilo kom čitaocu S1 **pada po imenu** (`StavkeOtkupaRedovi` / `ZbirStavkiZaOtkup`),
  nikad 0 — test po jedan za izvoz i za štampu.
- **Merenje:** `popis_citalaca.py` grupa `otk_stavka` = 0 posle S1c; S1d kompajl bez konstanti; `gen_schema_module.py --check`,
  `who_writes.py --check` i `--check-ownership`; `run_vba` puna suite (menja se jezgro čitanja).

#### Instrukcija za PWA i GAS (§14.8 t. 13 — ne radi se u S1, radi se u S5)

1. **OTK sheet po stanici** postaje dva taba: `OTK` (red po otkupu: `ClientRecordID`, `ServerRecordID` = `OtkupID`,
   datum, stanica, kooperant, vrsta, sorta, `TipAmbalaze`, `KolAmbIzdata`, parcela, `BrojDokumenta`, sync kolone) i
   `OTK_STAVKE` (red po stavci: `OtkupStavkaID`, `OtkupID`/`ClientRecordID` roditelja, `RedniBroj`, `Klasa`,
   `Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg`). Redosled kolona zapisati na JEDNOM mestu (danas se dva mesta
   ručno poklapaju: `BuildOTKSheetRowForOtkup` i `BuildOTKOperationalHeaders_`, nalaz E-5).
2. **PWA zapis otkupa** dobija niz `stavke[]` (1..2); forma otkupa dozvoljava drugu klasu; vrednost = zbir stavki
   (`otkupni-list.js:254`, kartica, pregled, knjiga polja `kpParseOpisOtkupa:145` čitaju stavke, ne jedan red).
3. **GAS `doPost action=sync`** prima zaglavlje + stavke u jednom zapisu i upisuje oba taba atomski po
   `ClientRecordID`; `mergeOtkupRows_:1521` i sadržajni fallback ključ (`gas/Code.gs:2857-2868`) porede zaglavlje +
   skup stavki, ne jedan red.
4. **`OtkupiAll` / menadžment** (`getMgmtAll`, `getOtkupiForOtkupac:2768`): isti oblik (zaglavlje + stavke);
   agregati po OM/klasi računaju iz stavki.
5. **VBA uvoz** (`ImportRowToTblOtkup`) čita `OTK_STAVKE` i zove `CreateOtkup_TX(h, stavke)` sa svim stavkama;
   do S5 PWA šalje jednu klasu po zapisu i uvoz pravi otkup sa jednom stavkom (danas radi, E-011 „ne“).
6. **Između S1c i S5 PWA prikaz desktop otkupa može biti pogrešan** (VBA piše nov oblik, GAS/PWA čitaju stari) —
   dozvoljeno po §14.7; S5 to zatvara.

**Otvoreno:** ništa od domena.

#### S1a — urađeno (17.09.2026)

- **Obrisano:** `modOtkup.SaveOtkup_TX`, `SaveOtkup`, `GetKooperantNazivForNovac` (bez pozivaoca);
  zatvoreni mrtvi lanci `modHelpers.CheckVerwaisteDokumente`, `modDokumenta.GetStorniraniGrupisano` →
  `GetStorniraniByTip`, **ceo modul `modMarza`** (javni `ReportMarza*` bez pozivaoca; ostatak su bili nedostupni privatni helperi; marža je van refaktora, §14.8 t. 8),
  `modOtkupUI.ColSpecIdx` → `modScrDokumenti.ColumnSpec`. Posle brisanja grep ne nalazi nijedan poziv.
- **Testovi:** 8 poziva starog pisca u `modBusinessFlowProTests` prešlo je na `CreateOtkup_TX`
  (`OtkHeader`/`OtkStavka`, `NoviOtkupFixture`); zaglavlje bez stavki pravi nov
  `OtkupBezStavkiFixture` (synthetic anomaly: kanonski otkup pa brisanje stavki u transakciji testa);
  deo `Test_BKTX_VlasnikOsaOdbijaTudjuStanicu` koji je merio stari pisac je obrisan. Review #352 (P1):
  `CreateOtkup_TX` zove `ApplyAvansToOtkup` (`modOtkup.bas:710`), pa svaki test koji se vraća rollback-om a pravi
  otkup snapshot-uje i **`TBL_NOVAC`** (`Test_OtkupReadHelpersExcludeStornirano`, `Test_OTP_StariOtkupNeUlazi`,
  `Test_BKTX_VlasnikOsaOdbijaTudjuStanicu`, `Test_OTK_VrednostBezStavkiPada`); `WHO_WRITES.md` regenerisan.
- **Alat:** `popis_citalaca.py` grupa **`x_otk_stavka`** (`COL_OTK_KOLICINA/CENA/KLASA/KOL_AMB/BRUTO/NOVAC/PRIMALAC`,
  bez `TIP_AMB`/`KOL_AMB_IZDATA`). Dokaz u oba smera: sabotaža `COL_OTK_KOLICINA` → `x_otk_stavka` 74→75 i
  `otk_linija` 91→92; sabotaža `COL_OTK_KOL_AMB_IZDATA` → `x_otk_stavka` 74 (ne raste), `otk_linija` 92; vraćeno 74/91.
- **Prag posle S1a:** `x_saveotkup` = 0; `x_otk_stavka` PROD 74 (ZIV_UI 55 · ZIV_MAKRO 7 · PAUZIRAN 4 · MRTAV 8).
- **Premešteno u S1b:** mrtvi lanac panela u `modOtkupBlok` (`LoadOtpremnice` ima 6 poziva unutar modula,
  `BuildFirstBlokCena`, `OfferHladnjacaIspravka` → `PrefillOtkupFromStornirano`) — traži proveru celog starog
  panela bloka, nije zatvoren lanac.
- **Verifikacija:** `vba_check` čisto, `who_writes --check` ažuran, `--check-ownership` bez novih pisaca,
  `gen_schema_module --check` u koraku. Prvi krug: `RunBusinessFlowProSuite` 1360/0 i ručni Compile čist.
  **Merge kapija (review #352):** pun `python tools/run_vba.py` 12/12 + `Debug → Compile VBAProject` posle review ispravki.

#### S1b — odluka operatera i spisak (17.09.2026)

Merenje posle S1a (`x_otk_stavka`, PROD, van sync/izvoza i `modSetup`): 32 mesta u 20 procedura.

**Odluka:** čitaoci polja stavki koji žive u procedurama **starog modela** (veza preko `Otkup.OtpremnicaID`,
poređenje po klasi reda, pojam „duplikat broj+klasa“) se u S1 **ne prevode nego brišu**, zajedno sa granom
pozivaoca koja ih koristi. Sposobnost koju su nosili vraća S3 (otpremnica, izgubljeni blokovi, storno blokova),
S4 (zbirna) ili S9 (sledljivost, integritet); do tada ne radi (§14.7). S1d ostaje posle S1c.

**S1b-1 — brisanje (lanac pozivalaca izmeren grep-om):**

| Briše se | Pozivaoci (procedura) — šta se uklanja iz njih | Vraća |
|---|---|---|
| `modSledljivost.AutoLinkOtkupOtpremnica` + `AutoLinkOtkupOtpremnica_TX` | `modScrSledljivost:432` (radnja auto-link na ekranu SLEDLJIVOST); testovi `modBusinessFlowProTests` `:584`, `:1338`, `:13612`, `:13720` | nijedan (heuristika; nov model ima eksplicitne izvore) |
| `modSledljivost.TraceByZbirna` | `modIzvestaj.StampajSledljivostZbirne:6105`, `modPaletniList.GetKooperantiZaZbirnu:2635`, `modPaletniList.GetOtkupiZaPalete:2712`; test `Test_FullDocumentChainHappyPath:590` | S9 |
| `modSledljivost.GetUnlinkedOtkupi` | `modIntegritet.Chk_B2_UnlinkedOtkupi:234` (provera odlazi, §14.8 t. 11) | S3 (provera novog modela) |
| `modDokumenta.GetLostOtkupBlokovi` | `modIntegritet.Chk_B3_IzgubljeniBlokovi:251`, `modOtkupBlok.LoadLostBlokovi:1034`, `modScrDokumenti.RowsIzgubljeni:623` (lista IZGUBLJENI), `modStornoRecovery.GetNedovrseno:71` (vrsta reda `IZGUBLJEN_BLOK`) | S3 |
| `modStornoFlow.GetStornoBlockRows` | `modScrStorno.StornirajBlokoveAko:1265` (storno blokova uz storno otpremnice); testovi `modTest.T_BlokoviF8_PoIdentitetu:3723-3771` | S3 |
| `modStornoZurnal.OtkupReissueDupExists` | `modStornoZurnal.UndoOperation_TX:171` (provera duplikata broj+klasa pri undo) | nijedan (pojam nestaje: otkup je jedan dokument) |

**S1b-2 — prevođenje na stavke** (sposobnost ostaje): `modIzvestaj.ReportOtkupRobaOM:2745`,
`ReportSledljivostLanac:5124-5133`, `ReportSledljivostProblemi:5688-5692` (ako ne padnu uz `TraceByZbirna`),
`modPaletniList.GetOtkupiZaPalete` (ako ostane), `modScrDokumenti.ColKlasa/ColKolicina/ColKolAmb/ColCena` +
`RedoviZaTip` + `RowsBlokovi` (mreža), `modStornoDok.Col*ZaPrefill` (storno prefill po stavkama, B-036), živi deo
`modOtkupBlok` (`SumKolByOtp`, `SumAmbByOtp`, `SumBrutoByOtp`, `BuildNapisanoByOtp`, `ExistingBlokCena`, `RenderSpec`,
`LoadBlokovi`, `KoopPrometYear`) i mrtvi lanac panela (`LoadOtpremnice`, `BuildFirstBlokCena`,
`OfferHladnjacaIspravka` → `PrefillOtkupFromStornirano`) koji se briše. Izvor: `modOtkup.ZbirStavkiPoOtkupu` /
`StavkeOtkupaRedovi`; bruto samo iz `StavkeOtkupaRedovi`.

#### S1b-1 — urađeno (17.09.2026)

**Obrisane procedure starog modela i grane pozivalaca** (grep posle brisanja: nula poziva):

| Obrisano | Uklonjeno kod pozivalaca | Vraća |
|---|---|---|
| `modSledljivost.AutoLinkOtkupOtpremnica(_TX)` | dugme „Poveži automatski“ na ekranu SLEDLJIVOST (`scrSlAuto`, `AutoPovezi`), 3 ključa poruka | nijedan |
| `modSledljivost.TraceByZbirna`, `modIzvestaj.StampajSledljivostZbirne` | meta ZBIRNA na SLEDLJIVOST sada javlja neuspeh štampe; `modPaletniList.GetOtkupiZaPalete` i `GetKooperantiZaZbirnu` obrisani — paletni i preradni list nemaju stavke po otkupu | S8 / S9 |
| `modSledljivost.GetUnlinkedOtkupi` | provera integriteta B2 | S3 (provera novog modela) |
| `modDokumenta.GetLostOtkupBlokovi` | lista IZGUBLJENI i radnja „preuzmi“ na DOKUMENTI (`RowsIzgubljeni`, `LostGridCols`, `PreuzmiBlok`), provera B3, vrsta `IZGUBLJEN_BLOK` u Oporavku, `modOtkupBlok.LoadLostBlokovi`; 9 ključeva poruka | S3 |
| `modStornoFlow.GetStornoBlockRows` | storno otkupnih blokova posle storna dokumenta (`modScrStorno.StornirajBlokoveAko`); sekcija blokova u uvidu storna je prazna kolekcija (`modStornoImpact`); 4 ključa poruka | S3 |
| `modStornoZurnal.OtkupReissueDupExists` | provera duplikata (broj, klasa) pri undo | nijedan |

**Testovi obrisani** (merili su obrisano ponašanje): `Test_AutoLinkNeVidiHeaderStavkeOtkup` (BFP), deo
`Test_FullDocumentChainHappyPath` o AutoLink/TraceByZbirna, tvrdnja šablona u `T_Sled_MeteSledljivosti`,
`modTest` 50 `T_BlokoviF8_PoIdentitetu`, 51 `T_StorniranSibling_ZadrzavaSvojBlok`, 65
`T_StornoImpact_BlokSekcijaDriftJeInvalidna`, 67 `T_StornoImpact_PrijemnicaBlokDriftJeInvalidan`, 71
`T_StornoBlokovi_PodrazumevanoNijedan` (pao u punom prolazu: lista blokova u stornu je sada prazna). Registar
`modTest` prenumerisan bez rupa (kapija `REGISTAR`). Katalog sabotaža: uklonjeno 6 unosa koji su gađali obrisan kod
(`sledljivost-sablon-dvosmislen-broj`, `uvid-blok-sekcija-guta`, `uvid-blok-zbirna-guta`, `uvid-blok-prijemnica-guta`,
`blokovi-po-broju`, `blockcount-po-broju`, `blokovi-svi-oznaceni`, `blokovi-oznake-prezive-izbor`,
`blok-status-ne-prati-izbor`).

**Merenje:** `x_otk_stavka` PROD 74 → **63**; `otk_linija` 91 → 79; DUAL READ živih 112 → 100.
**Ostaje za S3 čišćenje:** pomoćne procedure bloka u `modStornoFlow`/`modScrStorno` (`ActiveBlocksForFlow`,
`BlokOznacen`, `StornoSelectedBlocks_TX`, `BlockStornoDriftReason`) i prekidač izgubljenih u starom panelu
`modOtkupBlok` (`ToggleLostMode`) — kompajliraju se, ali više nemaju ulaz.

**Verifikacija:** `vba_check` čisto; `who_writes --check` (regenerisan) i `--check-ownership`; `gen_schema_module --check`.
**Pre merge-a:** pun `python tools/run_vba.py` + `Debug → Compile VBAProject` (menjaju se ekrani i testovi).

#### S1b-2 — urađeno (17.09.2026)

**Prevedeno na stavke** (`modOtkup.ZbirStavkiPoOtkupu` / `ZbirStavkiZaOtkup` / `StavkeOtkupaRedovi`; otkup bez stavki pada po imenu):

| Mesto | Šta čita sada |
|---|---|
| `modOtkupBlok.SumKolByOtp`, `SumAmbByOtp`, `BuildNapisanoByOtp` | kg i gajbe blokova otpremnice = zbir stavki (veza `Otkup.OtpremnicaID` ostaje do S3) |
| `modOtkupBlok.ExistingBlokCena` | cena prve stavke prvog bloka (predlog za nov blok) |
| `modOtkupBlok.RenderSpec` (specifikacija) | kg i vrednost reda sa stavki; neto cena reda = prosek (dve klase nemaju jednu cenu) |
| `modScrDokumenti` kolone moda OTKUP + `RowsBlokovi` | opis kolona imenuje kolone stavke (`COL_OKS_*`); blokovi: kg, gajbe, vrednost sa stavki |
| `modStornoDok.PrefillIzStorniranog` (B-036) | otkup je jedan red zaglavlja; `StavkeOtkupaZaPrefill` daje kol/amb/cena po klasi i `dveklase` |
| `modIzvestaj.ReportOtkupRobaOM` | kg blokova po otpremnici iz `modOtkupBlok.BuildNapisanoByOtp` (račun se ne duplira) |
| `modIzvestaj.ReportSledljivostLanac`, `ReportSledljivostProblemi`, `SledBlokSumMapa` | kg i klase otkupa sa stavki (`SledKgStavki`) |

**Obrisano:** ceo stari panel „Otkupni blokovi“ u `modOtkupBlok` (ulaz `AttachOtkupBlokPanel` bez pozivaoca: `LoadOtpremnice`,
`LoadBlokovi`, `BuildFirstBlokCena`, `OfferHladnjacaIspravka` → `PrefillOtkupFromStornirano`, `ToggleLostMode`, `KoopPrometYear`,
`SumBrutoByOtp`, `NapredakBlokaDostupan` i ostalo — modul 2230 → 548 linija) i njegov event-omotač `clsBlokUI.cls`
(izbačen i iz `WHITELIST`-a `vba_hard_census.py`); kapija pauze `NapredakBlokaDostupan` izbačena iz `popis_citalaca.py`.

**AUD-056 zatvoren:** KPI saldo OM čita kolonu 6 (`modOtkupUI.SaldoIzIzvestajaOM`).

**Testovi:** `Test_OTK_PanelNapredakJePauziran` (merio pauzu) → `Test_OTK_BilansOtpremniceSaStavki` (dvoklasni blok: kg 1000,
gajbe 20, napisano 1000, predlog cene 50); nov `Test_OTK_PrefillStornaDveKlaseSaStavki` (obe stavke u prefill-u); nov
`modTest` 213 `T_KpiSaldoOM_CitaKolonuSalda` + sabotaža `kpi-saldo-om-kolona-agro`.

**Merenje:** `x_otk_stavka` PROD 63 → **22** (ZIV_UI 18: `modStammdatenSync` izvozi 11, `modStanicaLock.BuildOTKSheetRowForOtkup` 4,
`modSetup` 3 — S1c/S1d; PAUZIRAN 4: `AutoCreateOtpremniceFromPWA` — briše S1c); `otk_linija` 79 → 37.

**Verifikacija:** `vba_check` čisto; `who_writes --check` / `--check-ownership`; `gen_schema_module --check`; `vba_hard_census`,
`vba_selfupdate_gates`, `vba_parity_check` zeleni. **Pre merge-a:** pun `python tools/run_vba.py` + `Debug → Compile VBAProject`.

#### S1b-3 — review #355: bez mosta preko `Otkup.OtpremnicaID` (17.09.2026)

**Pravilo (review #355, P1):** novi model otkupa se NE sme čitati kroz staru vezu `Otkup.OtpremnicaID` — takav čitalac je
most koji S3 opet menja (`tblOtpremnicaIzvori`). Sposobnost koja stoji samo na toj vezi se **briše**, ne prevodi; vraća je S3/S9.
Ostaje ono što ne zavisi od veze: mreža otkupa, vrednost i plaćanje, prefill ispravke, rang kooperanata, AUD-056.

**Obrisano:**

| Šta | Gde | Vraća |
|---|---|---|
| Radni sto otpremnice u F1: liste OTPREMNICE i BLOKOVI, aktivna otpremnica, traka bilansa (`zOtp`), upozorenje na prekoračenje, vezivanje bloka posle unosa (`LinkOtkupIDsToOtpremnica`), specifikacija blokova (i po datumu), izlazak iz konteksta pri promeni OM, izuzetak „datum otpremnice ostaje“ u `ClearForm` | `modScrDokumenti`, `modOtkupUI`, `modOtkupBlok`, `modPrint` (šablon specifikacije) | S3 |
| Bilans po otpremnici: `SumKolByOtp`, `SumAmbByOtp`, `BuildNapisanoByOtp`, `ExistingBlokCena`, `ExistingBlokZbirna` | `modOtkupBlok` (ostaje samo `KoopRangRows`, 210 linija) | S3 |
| Kolone „kg blokova“ i „razlika“ u Roba po OM — **prazne**, oblik rezultata isti | `modIzvestaj.ReportOtkupRobaOM`, `modScrIzvestaji` | S3 |
| Ceo ekran SLEDLJIVOST (`modScrSledljivost.bas`) i izveštaji `ReportSledljivostLanac/Problemi/Mete/Dokumenti` sa `Sled*` pomoćnicima; šablon sledljivosti u `modPrint` | lanac je kretao od otkupa preko `Otkup.OtpremnicaID` | S9 |

**Prefill (review #355, P1):** `modStornoDok.StavkeOtkupaZaPrefill` čita `modOtkup.StavkeOtkupaRedovi` (kanonska granica);
otkup bez stavki, nevažeća klasa ili dve stavke iste klase padaju po imenu, pa `PrefillIzStorniranog` ne vraća delimičan spec.
Negativni test `Test_OTK_PrefillStornaBezStavkiPada` (kontrola sa stavkama, pa brisanje stavki u transakciji → prazan spec).

**Testovi:** obrisani `Test_OTK_BilansOtpremniceSaStavki` (BFP) i `modTest` `T_Sled_*` (10) — mere obrisano;
`T_UtovarB_SledIStornoKapije` → `T_UtovarB_StornoKapije` (deo o lancu obrisan, storno kapije ostaju; poslednji u `RunOne`);
`T_ClearForm_Ugovor` meri ugovor bez otpremnice; iz testa čipova obrisan deo o listi otpremnica. Registar prenumerisan.
Katalog sabotaža: 37 unosa sledljivosti uklonjeno; `clear-datum` preusmeren na novo pravilo.

**Ostaci za S3/S9 (bez ulaza, ne čitaju vezu):** konstante `WS_SPECIFIKACIJA_SABLON` (obrisana) / `WS_SLEDLJIVOST_SABLON`,
`CFG_SPECIFIKACIJA_PRINT_MODE`, `CFG_SLEDLJIVOST_PRINT_MODE`, `OBL_SLEDLJIVOST`, ikona `IC_SLEDLJ`, ključevi poruka ekrana;
pisci veze `Otkup.OtpremnicaID` (`modSledljivost.ReassignOtkupToOtpremnica_TX`, `modDokumenta`, `modAutoHladnjaca`) — S3.

#### S1c — urađeno: izvozi i push na zaglavlje + stavke (17.09.2026)

**Kanon:** `modOtkup.StavkeOtkupaRedovi` nosi i `OtkupStavkaID` (kol. 7) i `BrutoKg` (kol. 8; prazno = neto), na kraju niza.
Svaki izvoz čita tu granicu ili `ZbirStavkiPoOtkupu`/`ZbirStavkiZaOtkup` — zaglavlje bez stavki pada po imenu, nikad 0 kg.

| Šta | Kako sada | Gde |
|---|---|---|
| `OtkupPoOM` | zbir po Stanica + Vrsta + **klasa stavke**; `BrojOtkupa` broji dokumente, ne stavke; oblik taba isti | `modStammdatenSync.OtkupPoOMRedovi` |
| `OtkupiAll` | red po **zaglavlju**; `Klasa/Kolicina/Cena/KolAmbalaze` izbačeni, dodat `KolAmbIzdata`; raspored kolona na jednom mestu (`OtkupiAllKolone`) | `modStammdatenSync.ExportOtkupiAll` |
| `OtkupiAllStavke` (nov tab MgmtReports) | red po stavci, raspored `modMasterSync.OtkStavkeKolone` | `OtkupiAllStavkeRedovi`; izvoz 5 → 6 tabova |
| `SaldoOMDetail` | kg/vrednost/gajbe po kooperantu iz `ZbirStavkiZaOtkup` | `OtkupSaldoPoKooperantu` |
| Push ka stanici | stavke → tab `OTK_STAVKE`, pa zaglavlje → `Sheet1` sa **praznim** linijskim poljima; zaglavlje je oznaka završenog push-a; slanje stavki je **idempotentno po `OtkupStavkaID`** (ugovor ispod) | `modStanicaLock.BulkPushPendingForStanica`, `BuildOTKSheetRowForOtkup` (po imenu), `EnsureOtkStavkeTab` |
| Raspored OTK kolona | **jedno mesto**: `modMasterSync.OtkZaglavljeKolone` / `OtkStavkeKolone`; `BuildOTKOperationalHeaders_` i graditelj reda čitaju odatle (nalaz E-5 zatvoren) | `modMasterSync` |
| `PwaIstiSadrzaj` | već poredi stavku (fail-closed, tačno jedna) — bez izmene | — |
| Health `healthprod` | `Check_CoreTablesAndColumns` i `Check_GoogleSyncMasterSchema` traže kolone **iz kanona** (`HealthRequireKanon` → `modSchema.SchemaTableColumns`), i za `tblOtkupStavke` | AUD-055 zatvoren |
| `modSetup.EnsureDoradeSchema` | bez formata `Otkup.Kolicina` i bez `Otkup.BrutoKg` (bruto je na stavci) | — |

**Raspored `Sheet1` OTK taba ostaje 23 kolone:** to je PWA ugovor do S5 (PWA puni, `ImportOneOTKSheet` čita poziciono `GS_*`).
Između S1c i S5 PWA prikaz desktop otkupa nema kg/cenu u `Sheet1` ni u `OtkupiAll` — dozvoljeno po tački 6 instrukcije iznad.

**Obrisano (vraća S5):** `modMasterSync.AutoCreateOtpremniceFromPWA` i `_TX` (E-020, pauzirano), pisac veze
`LinkOtkupToOtpremnicaStrict` (ostao bez pozivaoca); korak 3 ciklusa prijavljuje pauzu bez poziva. Testovi:
`Test_RF28_AutoOtpremnicaNeMesaArtikle` (merio obrisano grupisanje) i ulaz „OTP“ u `Test_PWA_IzvedeniLanacJePauziran`.

**Novi testovi (BFP):** `Test_OTK_IzvozDveKlaseIzStavki` (I 100 × 50 + II 40 × 30 = 6200 kroz OtkupPoOM, OtkupiAllStavke,
SaldoOMDetail i push), `Test_OTK_IzvozBezStavkiPada` (kontrola, pa brisanje stavki u transakciji → sva četiri puta padaju po `OtkupID`).

**Merenje:** `popis_citalaca.py` `x_otk_stavka` 22 → **0** u PROD; `x_literal` 15 → 7 (ostatak su veze
`BrojZbirne`/`OtpremnicaID` u `modIntegritet` i dve health provere — S3/S4, ne otkup).

**Ugovor `OTK_STAVKE` (review #357, P1/P2):** tab ima **tačno jedan red po `OtkupStavkaID`**. Pisac pre slanja jednom
pročita tab (`OtkStavkeIndeksIzTaba`), ne šalje stavku koja već postoji sa istim sadržajem, a isti ID sa drugačijim
sadržajem je **konflikt**: ne šalje se ništa od tog otkupa (ni stavke ni zaglavlje). Ponovljen push posle mrežnog pada
zato ne može da promeni količinu. Naslov taba mora biti tačno `OtkStavkeKolone` istim redom, inače push staje
(pisanje je poziciono). Čitalac u S5 sme dupli `OtkupStavkaID` da tretira kao kvar, ne kao zbir. Indeks važi za jedan
prolaz: push radi pod lock-om stanice, a PWA do S5 ne piše u `OTK_STAVKE` (kad S5 uvede PWA pisca, ugovor se proširuje).
Test `Test_OTK_PushStavkiIdempotentan`: prva stavka prođe, druga padne, retry → 2 reda, 140 kg; konflikt bez upisa;
pogrešan redosled naslova pada. Sabotaže `push-stavke-retry-dupla`, `push-stavke-naslov-bez-provere`.

Sledeći korak: **S1d**.

#### S1d — urađeno: linijske kolone obrisane iz `tblOtkup` (17.09.2026)

**Kanon:** iz `schema/schema.json` (`tblOtkup`) obrisano 8 kolona — `Kolicina`, `Cena`, `Klasa`, `KolAmbalaze`, `BrutoKg`
(polja stavke, `tblOtkupStavke`), `Novac`, `PrimalacNovca` (pripadaju `tblNovac`), `VremeUnosa` (zamenjen sa
`CreatedAt`/`SourceCreatedAt`); `modSchema.bas` regenerisan (37 → 29 kolona, otisak `5EC25CB1`). Iz `modConfig` obrisane
konstante `COL_OTK_KOLICINA/CENA/KLASA/KOL_AMB/BRUTO/NOVAC/PRIMALAC/VREME_UNOSA` — **kompajl bez njih je dokaz nula čitalaca**.

**Zatečena sveska:** kolone su obrisane iz SREDINE kanona, a upis je pozicioni (`SchemaReadyOrFail` poredi po indeksu).
Self-heal na startu (`modSetup.EnsureSledljivostSchema`, grana `tblOtkup`) ih briše kroz postojeći `ObrisiKolonuAko`, isto kao
`Isplaceno`/`DatumIsplate` u PR6; `run_vba` zove `EnsureRuntimeSchema` posle uvoza. Podatak u njima se gubi namerno (nema
produkcije; istina je na stavkama). `EnsureDoradeSchema` više ne dodaje `VremeUnosa`.

**Testovi (samo fixture-i, bez izmene ponašanja):** seed-ovi otkupa u `modTestBanka` (4), `modTestStorno` (2),
`modTestStornoCentar` (16), `modFakturaTests`, `modNovacTests` i RF-28 fixture (`AppendRF28OtkupFixture` bez `klasa`/`cena`)
ne pišu obrisane kolone — nijedan produkcioni čitalac ih više nije čitao, pa se merenje ne menja.
`Test_OTK_HeaderNeNosiLinePolja` sada tvrdi da **kolone ne postoje** (ne da su prazne); tvrdnje „zaglavlje ne nosi
količinu“ u 4 testa obrisane (pokriva ih prethodni). Obrisan `FindOtkupIDByBrojAndKlasa` i dve tvrdnje „ne nalazi se po
(broj, klasa)“. `Test_HladnjacaChainLinkFailureIsReported` sada nalazi otkup po broju — ranije je tražio po klasi, dobijao
prazan ID i tvrdnja „otkup NIJE povezan“ je prolazila nad praznim ključem (lažno zeleno). `OtkMrezaRed` čita kolone mreže
kroz `COL_OKS_*`.

**Merenje:** `popis_citalaca.py` `x_otk_stavka` 0 (PROD i TEST), `x_vreme_unosa` 0; `gen_schema_module.py --check` u koraku;
`who_writes --check` / `--check-ownership` 0. Preostalih 15 `otk_linija` su polja koja ostaju na zaglavlju (`TipAmbalaze`, `KolAmbIzdata`).

**S1 nije završen bez S1e** (review #358): S1d fizički završava stari data grain, ali storno i štampa još nose model
„više zaglavlja po klasi“. Sledeći korak: **S1e**.

#### S1e — identitet otkupa: UI → `OtkupID` → mutacija (17.09.2026)

**Pre-flight (review #358):**

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN | otkup = jedno zaglavlje + 1..2 stavke (§14.10 S1d); broj je labela, jedinstven tek po (OM, dan) — `KIND_OTK` |
| `IDENTITY` | PROVEN posle S1e | F1 i F8 nose nevidljiv `OtkupID` (`modScrDokumenti.IdKolonaTipa("OTKUP") = COL_OTK_ID`); `GeneracijaID` otkupni pisac ne upisuje, pa je F8 za OTKUP stvarno radio po broju |
| `CARDINALITY` | PROVEN | jedan dokument = jedan `OtkupID`; nema „svih redova broja“ |
| `INVARIANTS/OWNER` | PROVEN | `modStorno.StornoOtkup` (zurnal op, ambalaža, novac) nepromenjen; kaskada autohladnjače pripada `StornoOtkup_TX` |
| `WRITERS` | PROVEN | isti pisac (`StornoOtkup`); menja se samo ulaz. `StornoOtkupByBrDok_TX` i `StornoSelectedBlocks_TX` obrisani |
| `DOWNSTREAM` | PROVEN | zurnal op i dalje nosi broj kao LABELU (`BeginStornoOp` u `StornoOtkup`), `RowID` = `OtkupID`; undo u Oporavku ide po `OperationID`. F8 za OTKUP nema tok ispravke ni uvid (`TipUFlowDoc` = "") |
| `CAPABILITY` | PROVEN | štampa i storno iz F1, storno iz F8, ponuda ispravke autohladnjače, reprint iz izveštaja — ostaju; dodatni storno blokova (bez UI pozivaoca od S1b-1) vraća S3 |
| `ACCEPTANCE` | plan → test | `T_OtkupStornoPoID_NeDiraTudjeOM`: isti broj na dva OM-a, bez ID-a odbija, sa ID-em B stornira B a A ostaje; sabotaža `otkup-storno-po-broju` |
| `PLATFORM` | N/A | nema nove Excel/COM pretpostavke; nevidljiva kolona je postojeći mehanizam F8 |
| `LANDING` | stack na S1d (#358) | PR se otvara posle merge-a #358 |

**Izmene:**

| Tačka review-a | Šta | Gde |
|---|---|---|
| 1. F1 red nosi `OtkupID` | `Scr_Rows` traži kolonu identiteta za OTKUP; `RowAction` čita `OtkupID` iz nje (`IdentKolonaIndeks`, deljen sa F8); štampa `OutputOtkupniList otkupID`; prazan ID = „nema reda“ | `modScrDokumenti` |
| 2. kanonski storno po ID-u | `StornoOtkup_TX(otkupID)` nosi i kaskadu autohladnjače (stanica i `BrojZbirne` sa istog zaglavlja); `StornoOtkupByBrDok_TX` obrisan | `modStorno` |
| 3. `HladnjacaLanac` po ID-u | stanica i zbirna se čitaju po `OtkupID`; ponuda ispravke prefiluje po `OtkupID` | `modScrDokumenti` |
| 4. F8 `docID` = `OtkupID` | `StornoRazlog`/`StornoIzvrsi` za OTKUP: `OtkupAktivanPoID(docID)`, prazan ID se ne razrešava po broju | `modStornoDok` |
| 5. brisanje broj-ulaza | `OtkupIdsByBrDok`, `StornoOtkupByBrDok_TX`, `StornoSelectedBlocks_TX` (bez UI pozivaoca, grupisao po broju) | — |
| — | reprint iz izveštaja više ne širi na sve redove istog broja | `modPrint.ReprintOtkupniListByOtkupID` |
| 6. testovi dva zaglavlja | obrisani `Test_StornoJournalDualClass_Auto`, `Test_StornoJournalPartialClass_Auto`, `Test_StornoJournalEmptyBrDok_Auto`, `Test_StornoSelectedBlocks_Auto`; `T_OtkupBezGeneracije_NeStorniraTudjeOM` → `T_OtkupStornoPoID_NeDiraTudjeOM`; F8 seam testovi šalju (broj, `OtkupID`); `Test_StornoKaskadaScopePoLancu` i golden D3 storniraju po ID-u | testovi |

**Ostaje (svesno, van S1e):** zurnal ključ operacije otkupa je i dalje (tip, broj) kao labela — `LatestOpFor`/`UndoStorno_TX(tip, broj)`
(samo makro `Test_UndoStorno`) traže po broju; UI oporavka ide po `OperationID`. Testovi „isti broj, dva dokumenta“
(`ReusedBroj`, `DeadParentOtherGen`) ostaju: mere ponovno korišćen broj, ne model po klasi. → S9.

**S1 je završen** (S1a–S1e, PR-ovi #354–#359). Sledeći slajs po §14.9: **S2 — banka po ID-u**.

### 14.11) S2 — banka po ID-u (17.09.2026)

**Pre-flight (skill `pre-flight`):**

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN | posle S1 otkup = jedno zaglavlje po dokumentu; „blok“ u banci = otkupni list = jedan dokument. `BrojDokumenta` je **dnevni niz po otkupnom mestu** (`modBrojevi`, `KIND_OTK` → `MaxSeqFromTable(TBL_OTKUP, COL_OTK_BR_DOK, COL_OTK_DATUM, COL_OTK_STANICA, …)`, oblik `7/150326`) |
| `IDENTITY` | GAP → zatvoren u S2 | combo bloka je nosio `(BrojDokumenta + StanicaID)`, pa je pisac iz toga **ponovo tražio** dokument — display vrednost kao zamena za identitet, plus „scope“ da par uopšte bude jednoznačan |
| `CARDINALITY` | PROVEN | po `(broj, otkupno mesto)` postoji najviše jedan dokument; skup >1 je ili drugo otkupno mesto (legitimno) ili anomalija podataka — u oba slučaja **nije povod za raspodelu**, nego za ručno |
| `INVARIANTS/OWNER` | PROVEN | `SaveNovac` + `LinkNovacToOtkupStrict` nepromenjeni; dodata kapija vlasništva i storna u pisaču (obrazac `modNovac.ApplyAvansToOtkup`) |
| `WRITERS` | PROVEN | isti pisci; menja se samo ulaz (`brojBloka` → `otkupID`) |
| `DOWNSTREAM` | PROVEN | nalozi (`modBankaExportPregled.BuildBlokIsplataList`) i primena avansa (`modNovac.ApplyAvansToOtkup`, `modScrBankaNalozi.BnAvansNadIDovima`) **već rade po `OtkupID`**; broj ostaje labela na nalogu |
| `CAPABILITY` | PROVEN | D-032, D-033 netaknuti; D-035 greedy raspodela preko više kandidata **REPLACED**; D-036 i D-037 sačuvani (v. niže) |
| `ACCEPTANCE` | testovi | `T03_DvosmislenPozivNeObaraBatch`, `T11_RucniKooperantBezIzboraBloka`, `T21_IzabranPlacenBlokNijeAvans`, **novi** `T24_BlokTudjegKooperantaIStorniran`, `T_BankaUvoz_RucnoMapiranjePravila`, `Test_BIM_NovOtkupJeOtvorenBlok` |
| `PLATFORM` | N/A | nema novih Excel/COM pretpostavki |
| `LANDING` | čisto | grana od `main` posle #358/#359 |

**Izmene:**

| Šta | Gde |
|---|---|
| Lista blokova nosi `OtkupID` (kolona 1), broj i otkupno mesto su **prikaz** | `modBankaMapiranje.GetBlokoviZaBimMapiranje`, `modScrBankaUvoz.PuniCiljCombo` |
| Poziv na broj se razrešava **jednom**: `BimOtkupIzBroja` / `BimOtkupIzPozivaNaBroj`; 0 pogodaka = avans (namerno), >1 = `ERR_BMAP_MANUAL_REQUIRED` | `modBankaMapiranje` |
| Pisac prima `OtkupID`: `MapBankaImportAsKooperantBlockCore(bimID, koop, otkupID, …)`; na blok ide najviše njegov dug, ostatak u avans | `modBankaMapiranje` |
| Nova kapija pisca: dokument postoji, **nije storniran**, pripada **tom** kooperantu i **ima otkupno mesto** (`PotvrdiOtkupZaKooperanta`, `ERR_BMAP_BLOK_TUDJ` / `_STORNIRAN` / `_BEZ_OM`) | `modBankaMapiranje` |
| **Vlasništvo OM-a:** vezan red `tblNovac` nosi `StanicaID` **dokumenta**, ne matično mesto kooperanta; avans (i višak) nose matično mesto — odluka izrečena i merena | `MapBankaImportAsKooperantBlockCore` |
| „Otvoreno“ po dokumentu na jednom mestu: `BimOtvorenoNaOtkupu` (stavke − isplate); `BimOtkupBezOtvorenog` zamenjuje `BimBlokBezOtvorenih` | `modBankaMapiranje` |
| Okidač potvrde: **isplata veća od duga na bloku** (`BimOtkupTraziPotvrdu`), umesto „3+ otvorenih stavki“ | `modBankaMapiranje`, `modScrBankaUvoz.PitajZaPodelu` / `TekstPodele` |
| **Saglasnost je argument pisca**, ne UI konvencija: `dozvoliVisakKaoAvans` (podrazumevano **ne**); bez nje višak = `ERR_BMAP_VISAK_BEZ_POTVRDE` pre ijednog upisa. Auto put je prosleđuje izričito (tamo je avans namerno pravilo) | `MapBankaImportAsKooperantBlock*` |
| Obrisano: `GetOtkupCandidatesForKooperantBlock`, `PlanBlokRaspodela`, `SortKandidatiPoOtvorenomDesc`, `MAX_BLOK_KANDIDATA`, `BimScopeKolona`, `BimBlokTraziPotvrdu`, `TryResolveOtkupForKooperant` (mrtav), `BuScopeNedostaje`, `IzabranaStanicaCilja`, `ScopeIzbora`, `Scr_BuScopeBlokaTest`, `Scr_BuStopBezOmTest`, poruke `OTKUI_*_BU_BLOK_BEZ_OM` | — |

**Šta je scope bio i zašto ga više nema:** otkupno mesto je u mapiranje uvedeno zato što `(kooperant, broj)` nije bio jednoznačan.
Kad red liste nosi `OtkupID`, dvosmislenosti nema — pa nema ni scope-a, ni kapije „blok bez otkupnog mesta“, ni schema-drift
grane u kojoj scope tiho otpada. Tri stanja praznog stringa iz `BuScopeNedostaje` nestaju sa uzrokom.

**Višak u avans je odluka, ne ostatak deljenja (review #360, drugi P1):** ekran pita nad stanjem iz trenutka **prikaza**, a pisac dug računa u trenutku **upisa**. Dok saglasnost nije bila argument, pisac je mogao da napravi avans koji njegov pozivalac nikad nije odobrio — dovoljno je da se dug u međuvremenu smanji (druga isplata, ispravka stavki). Sada `dozvoliVisakKaoAvans` putuje od mesta odluke do pisca, podrazumevano je **ne**, a bez nje se ne piše ništa i stavka izvoda ostaje otvorena da operater dobije pitanje sa tačnim brojevima (`T25`, sabotaža `banka-writer-visak-bez-potvrde`).

**Vlasništvo nije isto što i identitet (review #360, P1):** pisac je znao tačan `OtkupID`, ali je `OMID` uzimao iz `tblKooperanti.StanicaID` — matičnog mesta. Za kooperanta koji predaje na dva mesta to daje red sa **tačnim** `OtkupID`-em i **pogrešnim** `OMID`-em, pa saldo tuđeg otkupnog mesta nosi kupovinu. Vezani red sada nosi `StanicaID` dokumenta (`T03`: `OTK-B@OM-1B → Novac.OMID = OM-1B`), a dokument bez otkupnog mesta se odbija (`ERR_BMAP_BLOK_BEZ_OM`, `T24`) — time je kapija „blok bez OM“ prešla sa ekrana (gde je bila deo scope-a) na mesto gde se piše. Avans i višak ostaju na matičnom mestu: nisu vezani ni za jedan dokument i mogu se kasnije primeniti na blok bilo kog mesta.

**Popravljeno usput (identitet, ne kozmetika):** automatski put nije imao otkupno mesto, pa je isti broj na dva otkupna mesta
ulazio u **jednu** raspodelu — jedna isplata na dva poslovna lanca. Sada takav red ide operateru (`T03`).

**Sposobnosti:** D-035 gubi greedy raspodelu preko više kandidata (kardinalnost koja ju je pravila nestala je u S1) — ostaje
„na blok ide njegov dug, višak u avans“. D-036 („ceo iznos kao avans, vežem kasnije“) je **sačuvan**, ali ga sada pokreće višak
preko duga, a ne broj kandidata; bez te izmene bi jedini put do te radnje nestao sa okidačem. D-037 (izabran plaćen blok = STOP)
radi nad `OtkupID`-em, na oba mesta (ekran i pisac).

**Merenje:** `COL_OTK_BR_DOK` u bankarskom putu ostaje na tri mesta i nijedno nije traženje dokumenta po broju:
razrešenje na granici (`OtkupIDoviPoBroju`), prikaz u listi blokova, i `TryResolveKooperantByOtkupPoziv` (poziv na broj →
**kooperant**, ne dokument). `modNovac.GetOpenOtkupi` broj samo prenosi kao labelu.

Sledeći korak: **S3 — otpremnica cutover (desktop)**.

### 14.12) S3 — otpremnica cutover, rez na pet koraka (17.09.2026)

Merenje pre reza (`popis_citalaca.py` na `main` `5610bfc9`): `otp_linija` 85 PROD mesta, `otp_stari_pisac` 49,
`otp_cena` 9, `otk_veze` 63 — ~190 mesta u 18 modula. Ali **živa produkciona ulazna tačka starog pisca je tačno
jedna**: `modDokUnos.OtpremnicaUpisi` → `SaveOtpremnicaMulti_TX`. Sve ostalo su čitaoci, pauzirani putevi ili test
fixture. Nova skela (`CreateOtpremnicaDraft_TX`, `IzdajOtpremnicu_TX`, `CreateOtpremnicaIzIzvora_TX`,
`GetOtpremnicaProgress`) postojala je **bez ijednog produkcionog pozivaoca**.

| korak | sadržaj | stanje |
|---|---|---|
| **S3a** | F2 otvara **nacrt**: `OtpremnicaUpisi` → `CreateOtpremnicaDraft_TX`; očekivanje po klasi sa `PredlogCena`; ambalaža se knjiži **pri izdavanju**; malina auto-zbirna pauzirana | ovaj PR |
| **S3b** | čitaoci na stavke: F2 mreža, štampa, izveštaji (i prazne kolone koje čekaju ovaj slajs), invarijante; **panel blokova nad `tblOtpremnicaIzvori`** i radnja „Izdaj“ | |
| **S3c** | identitet: storno i ispravka otpremnice po `OtpremnicaID`; nevidljiva kolona F2 = `COL_OTP_ID`; brisanje `*ByBroj_TX` | |
| **S3d** | auto lanci kroz `CreateOtpremnicaIzIzvora_TX`; otvaranje A13 kapije u `IspravkaOtkupa_TX` | |
| **S3e** | brisanje: `Otkup.OtpremnicaID/VozacID/BrojOtpremnice`, linijska polja i `Cena` na `tblOtpremnica`, `SaveOtpremnica*` | |

**Zašto „Izdaj" nije u S3a.** Odluka §14.8 t. 4 kaže „operater dugmetom kad je ostatak 0“ — a ostatak se vidi tek u
panelu blokova, i tek tamo postoji izvor koji izdavanje traži (`OtpIzdaj` odbija otpremnicu bez izvora). Radnja nad
redom bi uz to tražila `OtpremnicaID` u nevidljivoj koloni, a tu kolonu danas čita i `modScrStorno`
(`IdentKolonaIndeks`) očekujući **`GeneracijaID`** — promena bi mu tiho podmetnula drugi identitet. Zato izdavanje
ide zajedno sa panelom (S3b) i identitetom (S3c), a ne ranije.

#### Pre-flight S3a

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN | otpremnica = zaglavlje + očekivanje po klasi + izvori; odluke §14.8 t. 1–4 |
| `IDENTITY` | PROVEN | `OtpremnicaUpisi` vraća **`OtpremnicaID`**, ne broj; ekran ga čuva pod `polja("otpremnicaID")`, a operateru prikazuje broj |
| `CARDINALITY` | PROVEN | dvoklasan unos = **jedan** header + dve stavke (pre: dva reda pod istim brojem) |
| `INVARIANTS/OWNER` | PROVEN | očekivano = povezano pri izdavanju; ambalaža knjižena jednom, u istoj transakciji (`tx.AddTableSnapshot TBL_AMBALAZA`) |
| `EVENTS` | PROVEN | fizički (gajbe odlaze) = poslovni (dokument nastaje) = **izdavanje**; finansijskog nema — `PredlogCena` se ne knjiži |
| `DOWNSTREAM` | PROVEN | izmereno: `otp_cena` 5 živih čitalaca zaglavlja i `otp_linija` 47 od sada čitaju **prazno**; vraća S3b |
| `CAPABILITY` | PROVEN | A-029 MIGRATED; E-022 (malina auto-zbirna) **pauzirana glasno**; A-011/A-012/A-018..A-028 i dalje odsutni od S1b-3 |
| `ACCEPTANCE` | testovi | `Test_OTP_F2OtvaraNacrt`, `Test_OTP_PredlogCeneJePoKlasi`, `Test_OTP_AmbalazaSeKnjiziPriIzdavanju`, `Test_OTP_MalinaZbirnaPauzirana` + tri sabotaže |
| `PLATFORM` | N/A | nema novih Excel/COM pretpostavki |
| `LANDING` | čisto | grana od `main` `5610bfc9`; menja kanon → `gen_schema_module --check` i obe `who_writes` kapije |

#### Izmene S3a

| Šta | Gde |
|---|---|
| `PredlogCena` na `tblOtpremnicaStavke` (kanon + `modSchema`), `COL_OPS_PREDLOG_CENA` | `schema/schema.json`, `modConfig` |
| Očekivana stavka prima `PredlogCena`; ključ `Cena` na **zaglavlju se odbija** | `modDokumenta.OtpOcekKljucPoznat`, `OtpHdrKljucPoznat` |
| Zaglavlje više ne piše `COL_OTP_CENA` | `BuildOtpremnicaHeaderRowData`, `OtpIzmeniDraft` |
| Ambalaža pri izdavanju, kao zbir stavki | **nova** `OtpKnjiziAmbalazu`, zvana iz `OtpIzdaj` |
| F2 otvara nacrt; kultura se razrešava u adapteru | `modDokUnos.OtpremnicaUpisi` |
| Malina auto-zbirna pauzirana uz poruku | isto, `DOKUNOS_MSG_ZBIRNA_PAUZIRANA` |
| **Ispravka otpremnice pauzirana** (review #361 P1) — `ZavrsiIspravkuAko` uklonjen | isto, `DOKUNOS_MSG_OTP_ISPRAVKA_PAUZIRANA` |
| Ekran čuva ID, prikazuje broj | `modScrDokumenti.SaveOtpremnica` |

**Merenje posle:** `otp_stari_pisac` — sva tri tela starog pisca (`SaveOtpremnica`, `SaveOtpremnica_TX`,
`SaveOtpremnicaMulti_TX`, 42 mesta) prešla su iz `ZIV_UI` u **`PAUZIRAN`**: nemaju više nijednog živog pozivaoca.
Preostala 4 `ZIV_UI` mesta u toj grupi su **sudar imena** — ekranska `modScrDokumenti.SaveOtpremnica` (adapter F2,
ne pisac) koju popis hvata po imenu `SaveOtpremnica*`. Prag S3e („grupa = 0") mora da računa na to: ili se ekranska
rutina preimenuje, ili popis dobije izuzetak za `modScr*`.

**Šta S3a namerno NE radi:** nacrt se ne može izdati dok S3b ne vrati vezivanje blokova. To nije regresija — veza
`Otkup.OtpremnicaID` iz F1 obrisana je u S1b-3, pa nijedna otpremnica ni danas nema izvore.

#### Ispravka otpremnice — pauzirana, ne prevedena (review #361, P1)

Prva verzija S3a je posle otvaranja nacrta i dalje zvala `ZavrsiIspravkuAko FLOW_DOC_OTPREMNICA`. To nije zatvaranje
konteksta nego **pisac starog modela**: `CompleteOtpremnicaIspravka` → `ReassignOtkupToOtpremnica_TX` upisuje
`Otkup.OtpremnicaID` i `BrojZbirne`, pa rekalkuliše ili stornira zbirnu. Nad upravo otvorenim nacrtom to je bilo
pogrešno dvaput:

1. nov model bi se vezivao **starom vezom** — tačno most koji se po §14.7 ne pravi;
2. dokument koji **nema nijedan izvor** i nije `IZDATO` bio bi proglašen zamenom izdate otpremnice, a correction
   kontekst zatvoren. Lifecycle je: `DRAFT → članstvo → očekivano = povezano → IZDATO → tek onda zamena.`

Poziv je uklonjen. Sposobnost **B-038** je `PAUZIRAN` i vraća se u **S3c**, nad `OtpremnicaID`-em i
`tblOtpremnicaIzvori`, bez ijednog dodira `Otkup.OtpremnicaID`. Operater dobija poruku da ispravka i dalje čeka na
ekranu Oporavak — tiho preskakanje bi značilo da misli da je završena.

#### Stari pisac: zašto tek S3e (review #361, P2)

Izmereno posle S3a: `SaveOtpremnica` (76), `SaveOtpremnica_TX` (78), `SaveOtpremnicaMulti_TX` (155) — **309 linija
bez ijednog živog produkcionog pozivaoca**. Brisanje sada nije mali diff: 54 poziva u `modBusinessFlowProTests`,
1 u `modGoldenTests`, i **2 u `modAutoHladnjaca`** — a taj drugi je pauzirani auto-lanac čiju sudbinu (vraća se kroz
`CreateOtpremnicaIzIzvora_TX` ili se briše) odlučuje **S3d**. Brisanje sada bi tu odluku nametnulo prerano.

**Pravilo od S3a:** nijedan nov kod ni nov test ne sme da zove `SaveOtpremnica*`. Kandidat za mehaničku kapiju u
S3b: `popis_citalaca.py` da nauči prag po grupi (`otp_stari_pisac` ne sme da raste), umesto pravila u dokumentu.

Sledeći korak: **S3b — čitaoci otpremnice na stavke + panel blokova**.

### 14.13) S3b-1 — citaoci otpremnice na stavke (18.09.2026)

**S3b je rezan na dva.** §14.12 ga je opisao kao jedan korak („citaoci + panel blokova + Izdaj"), ali to su dve
razlicite vrste posla i dve razlicite vrste rizika: preusmeravanje **postojecih** citalaca na kanonski izvor, i
**nov ekran** sa tri produkciona pisca (`DodajOtpremnicaIzvor_TX`, `UkloniOtpremnicaIzvor_TX`, `IzdajOtpremnicu_TX`).
Uz to panel trazi `OtpremnicaID` u nevidljivoj koloni reda, a tu kolonu danas cita `modScrStorno` ocekujuci
`GeneracijaID` — isti razlog zbog kojeg „Izdaj" nije usao u S3a. Zato:

| korak | sadrzaj | stanje |
|---|---|---|
| **S3b-1** | citaoci na stavke: F2 mreza, stampa, izvestaji, invarijanta zbirne, lista i uvid storna, prefill ispravke | ovaj PR |
| **S3b-2** | **panel blokova** nad `tblOtpremnicaIzvori` + radnja **„Izdaj"** (odluka §14.8 t. 4); vraca A-011, A-012, A-018..A-028 | |

#### Pre-flight S3b-1

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN | otpremnica = zaglavlje (Datum, StanicaID, VozacID, KulturaID, Broj, TipAmbalaze, Vrsta, Sorta) + ocekivanje po klasi na `tblOtpremnicaStavke` |
| `IDENTITY` | PROVEN | svaki nov citalac trazi `OtpremnicaID`, ne broj (`ZbirStavkiZaOtpremnicu`, `StavkeZaOtpremnicu`); broj ostaje labela |
| `CARDINALITY` | PROVEN | jedno zaglavlje : N stavki. Mreza crta **jedan red** po dokumentu, izvestaj po otkupnom mestu **jedan red po klasi** (manjak se razresava kroz kljuc stavke zbirne, koji nosi klasu) |
| `INVARIANTS/OWNER` | PROVEN | „zbirna = zbir svojih aktivnih otpremnica, po klasi" (`modDokumentInvariant`) sada sabira stavke; vlasnik upisa se ne menja (`who_writes --check-ownership`: nema novih pisaca) |
| `EVENTS` | N/A | read-only slajs — nijedan poslovni dogadjaj se ne pomera |
| `WRITERS` | N/A | nijedan `_TX` nije dodat ni izmenjen |
| `DOWNSTREAM` | PROVEN | merenje: `otp_cena` zivih **5 → 0**, `otp_linija` zivih **59 → 24**; preostala 24 su cinjenice ZAGLAVLJA (Vrsta, Sorta, TipAmbalaze, KulturaID) i tri kozmeticka reda u `modSetup` koja odlaze u S3e |
| `CAPABILITY` | PROVEN | nijedna sposobnost ne menja status — S3b-1 **vraca brojeve** koje je S3a ispraznio. Jedini izuzetak je F-090, a on je vec u tabeli „nije sposobnost" |
| `ACCEPTANCE` | testovi | `Test_OTP_MrezaCitaStavke`, `Test_OTP_ZaglavljeBezStavkiObaraCitaoce`, `Test_OTP_IzvestajOMRedPoKlasi`, `Test_OTP_InvarijantaSabiraStavke`, `Test_OTP_PrefillIspravkeCitaStavke` + cetiri sabotaze |
| `PLATFORM` | N/A | nema novih Excel/COM pretpostavki |
| `LANDING` | **trazi regeneraciju fixture-a** | v. „Fixture" nize — bez nje suite ne moze da prodje |

#### Izmene

| Sta | Gde |
|---|---|
| Kanonski citaoci stavki: `StavkeOtpremniceRedovi`, `ZbirStavkiPoOtpremnici`, `ZbirStavkiZaOtpremnicu`, `StavkeOtpremnicePoDokumentu`, `StavkeZaOtpremnicu` | **novo**, `modDokumenta` |
| F2 mreza: `Col*` za OTPREMNICA na `COL_OPS_*`; jedna `ovStav` staza za oba tipa | `modScrDokumenti.RedoviZaTip` |
| Stampa: jedan red = jedna stavka (pre: jedan red iz zaglavlja) | `modPrint.FillOtpremnicaSablon` |
| Izvestaji: vozac zbirno, vozac po vrsti, otkupno mesto (red po klasi) | `modIzvestaj` |
| Invarijanta zbirne i provera pre unosa; lista siroceta | `modDokumentInvariant.SumOtpremniceByKlasa`, `modDokumenta.ValidateZbirna*`, `GetVerwaisteOtpremnice` |
| Audit B4a/B5b: kolicina iz stavki, **meko** (audit ne sme da padne na kvaru koji prijavljuje) | `modIntegritet` |
| F8: lista za storno i uvid pre storna | `modStornoFlow.AddStornoDocs2`, `modStornoImpact.SumActiveOtpStavke` |
| Prefill ispravke po uzoru na otkup (`StavkeOtkupaZaPrefill`) | **nova** `modStornoDok.StavkeOtpremniceZaPrefill` |
| Backfill prijemnica hladnjace: javni ulaz **PAUZIRAN** | `modAutoHladnjaca.BackfillPrijemniceHladnjaca` |
| **Kapija** `popis_citalaca.py --check`: prag zivih mesta po grupi | `tools/popis_citalaca.py` |

#### Ugovor citaoca: zaglavlje bez stavki PADA

Isti ugovor koji otkup nosi od S14.7 (`modOtkup.StavkeOtkupaRedovi`), i iz istog razloga (review #334, P1):
zaglavlje bez stavki nije dokument kolicine nula nego **nije dokument**. Pisac ga ne moze napraviti —
`OtpUpisiOcekivano` odbija prazno ocekivanje — pa citalac koji nedostajuci kljuc procita kao nulu laze o
dokumentu koji ne postoji.

Jedina razlika u odnosu na otkup: **`PredlogCena` sme da bude prazna.** Ona je predlog, ne knjizena cena
(odluka §14.8 t. 2), pa „nije predlozena" nije kvar. Upisana nula jeste — i odbija se.

Dva mesta namerno citaju **meko**, i to su jedina dva: audit integriteta (`modIntegritet`) i lista za storno.
Oba postoje da NABROJE pokvarene redove; tvrd pad bi ih ucinio slepim tacno tamo gde su potrebni. Kolona
kolicine tamo ostaje **prazna**, ne nula.

#### Kapija umesto pravila u dokumentu (review #361, P2)

`popis_citalaca.py --check` nosi prag zivih mesta po grupi (`PRAGOVI`). Prag se **spusta** kad slajs skine
citaoce i nikad se ne podize bez odluke u planu; merenje **ispod** praga takodje pada, jer bi zastareo prag
pustio grupu da naraste nazad bez ijednog crvenog — isti rod greske kao sidro sabotaze koje vise ne pokazuje ni
na sta.

Dokazano u oba smera: jedan nov poziv `SaveOtpremnica_TX` u `modScrDokumenti` digne `otp_stari_pisac` sa
**4 na 27** (ceo dosad mrtav pisac ozivi) i `--check` vrati exit 1 sa imenima; vracanje izvora vraca exit 0.

Stanje pragova 18.09.2026: `otp_stari_pisac` 4 (sudar imena — ekranski adapter `modScrDokumenti.SaveOtpremnica`,
ne pisac), `otp_linija` 24 (cinjenice zaglavlja + `modSetup`), `otp_cena` **0**.

#### Fixture — zasto mora da se regenerise

Zateceni `tests/fixtures/otkup_test.xlsm` ima 28 redova `tblOtpremnica` i **nijednu** stavku: tabele
`tblOtpremnicaStavke` u njemu uopste nema (pravi je self-heal na startu, prazna). Posle ovog PR-a svaki citalac
kolicine otpremnice trazi stavke, pa bi ti redovi obarali suite — ne zbog koda nego zbog **test podataka u
starom modelu**.

`tools/make_fixture.py` zato izvodi `tblOtpremnicaStavke` iz `tblOtpremnica`, isto kao sto od S1 izvodi
`tblOtkupStavke` iz `tblOtkup`: jedna stavka po zaglavlju, `PredlogCena` prazna kad zaglavlje nema cenu.
**Zaglavlja se NE spajaju po broju** iako fixture drzi dva reda istog `BrojOtpremnice` — to su dva dokumenta sa
razlicitih stanica i cetiri scenarija se oslanjaju bas na to; spajanje bi bilo izmena scenarija, ne prenos
podataka u nov model. Kolone `Klasa/Kolicina/KolAmbalaze/Cena` na zaglavlju ostaju do S3e; do tada fixture nosi
iste brojeve na oba mesta, a merodavna je stavka.

To nije migracija podataka (tih nema, §14.7) nego **autorstvo test podataka u novom modelu**.

Sledeci korak: v. §14.14 (S3b-1 je proširen), pa **S3b-2**.

### 14.14) S3b-1 proširen: stari pisac obrisan, F3 pauziran do S4, F4 do S6 (19.09.2026)

**Nalaz iz prvog punog prolaza S3b-1** (`run_vba`, 19.09): 27 padova u BFP-u, 10 od 12 golden scenarija, 5 u storno suite-u,
2 u izveštajima. Uzrok je jedan: stari pisac `SaveOtpremnica*` je i dalje pravio otpremnice **samo sa zaglavljem**,
a zvali su ga testovi — 50 poziva u BFP-u, 1 u golden testovima, plus direktni seed u tri suite-a. Prvi takav
dokument obara strogi čitač za ceo prolaz.

**Drugi nalaz, dublji, i propust iz S3a:** od #361 **nijedan živi put ne vezuje otpremnicu za zbirnu**. Svi pisci
`Otpremnica.BrojZbirne` su obrisani ili pauzirani (stari pisac, auto-lanac hladnjače, malina auto-zbirna, uvoz VOZ),
a F2 nacrt namerno ne šalje `BrojZbirne`. F3 ima tvrdu kapiju „zbir zbirne = zbir vezanih otpremnica“
(`ZbirnaSeSlazeSaIzvorom`), pa od S3a nijedna zbirna ne može da se sačuva. Za njom staje i F4, koja traži
postojeću zbirnu. S3a je glasno pauzirao auto-zbirnu za malinu, ali ručnu nije ni pogledao. Mapa je netačno
tvrdila da su A-030 i A-031 živi.

**Odluka (operater, 19.09.2026):** (a) F3 i F4 su **glasno pauzirani** — F3 do **S4**, F4 do **S6**
(ispravljeno u review-u #362, v. §14.15); (b) testovi starog lanca zbirne se
**brišu** i upisuju ovde kao spisak koji S4 mora da vrati. **Ne prepravljaju se ručnom vezom `BrojZbirne`**, jer
bi tako bili zeleni nad vezom koju produkcija ne ume da napravi (⚠ FALSE-GREEN RISK).

**Redosled slajsova se ne menja.** Zbirna u novom modelu sabira izdate otpremnice, a izdavanje dolazi sa panelom
blokova. Redosled ostaje S3b-2 → S4, a lanac posle F2 stoji do S4.

#### Izmene

| Šta | Gde |
|---|---|
| `SaveOtpremnicaMulti_TX`, `SaveOtpremnica_TX`, `SaveOtpremnica`, `ValidateOtpremnicaInput` **obrisani** | `modDokumenta` |
| Auto-lanac hladnjače obrisan (`AutoChainHladnjaca`, backfill prijemnica, test seam), ostaju samo oznake hladnjače i relink stanje | `modAutoHladnjaca` (675 → 58 linija) |
| `CalculateProsekGajbe` (po broju otpremnice) obrisan — samo test ga je zvao | `modDokumenta` |
| **F3 (do S4) i F4 (do S6) pauzirani** pre svih provera, uz poruke `DOKUNOS_ERR_ZBIRNA_PAUZIRANA` / `DOKUNOS_ERR_PRIJEMNICA_PAUZIRANA` | `modDokUnos.ZbirnaValidiraj`, `PrijemnicaValidiraj` |
| `SumOtpremniceByKlasa` **više ne guta grešku**. Progutana je davala nule, a `RecalculateZbirnaFromOtpremnice_TX` je tada upisivao 0 kg u zbirnu (storno T01) | `modDokumentInvariant` |
| Nacrt nove zbirne (`CreateZbirna_TX`) čita **stavke** otpremnice. Da li prima samo IZDATU odlučuje S4 | `modDokumenta.CreateZbirna` |
| Ekranski adapter F2 preimenovan u `SnimiOtpremnicu` (bio je sudar imena sa starim piscem) | `modScrDokumenti` |
| Prag `otp_stari_pisac` **4 → 0** | `tools/popis_citalaca.py` |

#### Testovi koji ostaju, prebačeni na novi oblik

- `Test_PR3_*` (19, nacrt nove zbirne): `Pr3Otpremnica` pravi otpremnicu **produkcionim** piscem
  (`CreateOtpremnicaDraft_TX`). Druga vrsta = druga kultura (`TEST_KUL_BEZ_SORTE_ID`).
- `Test_OTP_BrojZauzetPoStaniciIDanu`: broj po (stanica, dan) nad novim piscem. Dvoklasna otpremnica je **jedno**
  zaglavlje.
- `modTestStorno`: seed piše zaglavlje i stavku. `BrojZbirne` ostaje veza starog modela za storno okvir;
  prelazi u S3c (storno po ID-u) i S4 (članstvo).
- `modIzvestajTests` E2E: seed je novog oblika, a „Klasa I+II istog dokumenta“ je sada **jedno zaglavlje sa dve
  stavke**, koje izveštaj razvija u dva reda.
- `modTest`: `T_ZbirnaUnos_PauziranDoS4` i `T_PrijemnicaUnos_PauziranDoS4` mere da pauza stoji **pre** svih
  provera; `T_ScrSave_RutaPoRezimu` dokazuje rutu porukom pauze. Registar je prenumerisan 1..200, a redosled
  izvršavanja je proveren po imenu.

#### Spisak za S4 i S6 — obrisani testovi i scenariji koje moraju da vrate

Kolona „vraća“ kaže koji slajs (review #362: F4 i sve što je prijemnica je **S6**, ne S4).

Tela su u git istoriji (`fc06fa77`, poslednji commit pre brisanja); golden snimci su u `tests/golden/` istog commita.

| Tema | Obrisano | Šta je tvrdilo | vraća |
|---|---|---|---|
| **Zbir zbirne = zbir otpremnica** | `T_ZbirnaValidiraj_MoraDaSeSlazeSaOtpremnicama`; storno T01, T02, T20; sabotaže `zbirna-kapija`, `zbirna-kapija-strogo` | kg i gajbe po klasi moraju da se slože; kapija ne zavisi od `VALIDACIJA_UNOSA`; storno ili poništenje otpremnice rekalkuliše zbirnu i ne obara deljenu | S4 |
| **Redosled provera F3** | `T_ZbirnaValidiraj_TraziVozaca`; sabotaža `zbirna-vozac` | vozač je prva provera zbirne | S4 |
| **Redosled provera F4** | `T_PrijemnicaValidiraj_TraziKupca`; sabotaža `prijemnica-kupac` | kupac je prva provera prijemnice | S6 |
| **Bruto → neto** | `T_BrutoNeto_PoRezimu`; sabotaže `bruto-prijemnica`, `bruto-prijemnica-neto`, `bruto-zbirna` | prijemnica zamrzava bruto po klasi (S6); zbirna bruto NEMA (S4) | S4 + S6 |
| **Prijemnica se vezuje samo na jednoznačnu zbirnu** (A13, A14) | `T_Prijemnica_VezujeSeSamoNaJednoznacnu`; sabotaže `zbirna-f4-nije-vezan`, `zbirna-f4-pusta-tudjeg-vlasnika` | nema zbirne / dvosmislena / tuđ vlasnik → odbijeno po imenu | S6 |
| **Identitet deteta zbirne** (ZBR-CHILD-01) | `Test_ZBR_DeteNosiGeneracijuRoditelja`, `…BackfillNeVezeStaroDeteNaNovuGeneraciju`, `…KaskadaNeDiraDecuDrugogDokumenta`, `…RezimJeZaCeluOperacijuNePoTabeli`, `…IspravkaPodNovimBrojem`, `…IspravkaVezeSvojuDecuNeTudju`, `…TudjaGeneracijaNeOtvaraKapiju`, `…MutacijaPoBrojuStajeNaDvaDokumenta`, `Test_GeneracijaIDNaSavePutanji`; 10 sabotaža ZBR-CHILD | kaskada dira SVOJU decu, ne svu pod brojem; ispravka uzima identitet starog dokumenta; mutacija po broju staje na dva dokumenta. **U S4 se ovo prevodi na `tblZbirnaIzvori`, ne na generaciju.** | S4 |
| **Storno kaskada lanca** | `Test_StornoGuardNaSvimPutanjama`, `Test_StornoGuardUKaskadi`, `Test_StornoKaskadaScopePoLancu`, `Test_DokumentaReadHelpersExcludeStornirano` | kaskada ne obara tuđi lanac pod istim brojem; čitači izuzimaju stornirane | S4 (zbirna), S6 (prijemnica) |
| **Druga klasa** | `Test_ZbirnaKlasaIIGuard`, `Test_DualClassDocumentWrappers` | izvor sa klasom II blokira unos bez „Dve klase“ | S4 |
| **Ekran zbirne → pisac → tabela** (MIG-001) | `Test_ZbirnaEkranNosiOdrediste` | hladnjača i pogon iz rečnika stižu u `tblZbirna` | S4 |
| **Prosek gajbe** | `Test_ProsekGajbeExcludesStornirano` | prosek bez storniranih; `CalculateProsekGajbeByZbirna` (živ u UI) ostao bez testa | S4 |
| **Malina auto-zbirna** (E-022) | `Test_MalinaAutoZbirnaFromOtpremnice` | 1:1 otpremnica → zbirna | S4 |
| **Auto-lanac hladnjače** (A-014) | 5 `Test_HladnjacaChain*`, 2 `Test_BackfillHladnjaca*`; sabotaža `autochain-ne-dovrsava-vezu-otpremnice` | pad koraka zaustavlja lanac; broj prijemnice se deli po zbirnoj. Vraća se samo ako S3d/S4 odluče da se lanac vraća | S3d/S4/S6 |
| **Pun lanac do fakture (golden)** | A1, A2, A3, A4, A5, B2, B3, D1, D2, F2, G1 | baseline celog lanca, kardinalnost, kalo, avans, storno fakturisane prijemnice, delimično fakturisanje po klasi. **Golden se snima iznova tek kad downstream slajsovi postoje** — lanac do fakture traži S4, S6 i fakturu | posle S6 |
| **Stari pisac** | `Test_FullDocumentChainHappyPath`, tri `Test_InvalidOtpremnica*`; ručni `CreateSEFLive*` | validacija starog pisca. Nema šta da se vrati: novi pisac ima svoje testove (`Test_OTP_*`) | — |

`modTestStornoCentar` i ostatak `modTestStorno` i dalje seju vezu `BrojZbirne`. Zeleni su jer ne čitaju količinu
kroz strog čitač. Prelaze u S3c/S4.

#### Merenje posle

`otp_stari_pisac` **0** (PROD i TEST), `otp_cena` živih 0, `otp_linija` živih 24 (činjenice zaglavlja i `modSetup`).
`vba_check`: 480 sabotaža (20 obrisano zajedno sa svojim testovima), nula nalaza. `RunAllTests` je sada 200 testova
(bilo 203), golden 2 scenarija (bilo 12).

**Kapija `--check` dokazana i u drugom smeru, sama od sebe:** kad je stari pisac obrisan, grupa je pala na 0, a
prag je još bio 4. `--check` je vratio `PRAG ZASTAREO` sve dok prag nije spušten.

### 14.15) Review #362 — predlog cene nije vrednost, nacrt nije otpremljena roba, F4 je S6 (19.09.2026)

Review na `379688e3` (pun `run_vba` je na tom commitu bio zelen: 200/0, BFP 1185/0) dao je NO-GO sa tri P1.
Sva tri su tačna.

**P1 — `PredlogCena` je ponovo postala finansijska cifra.** Odluka je jasna (tabela odluka o `PredlogCena`):
predlog je **ne-finansijsko** polje i „nigde se vrednost otpremnice ne računa kao `Kolicina × Cena`“. S3b-1 je
uprkos tome uveo `vrednost = Σ Kolicina × PredlogCena` u `ZbirStavkiPoOtpremnici`, dva izveštaja vozača su tu
cifru prikazivala kao vrednost, a štampa ju je uzimala kao cenu, osnovicu, nadoknadu i ukupno. Ispravka:

- `ZbirStavkiPoOtpremnici` više nema vrednost. Mesto (1) je `Null`, ne 0, da slučajna upotreba padne.
- **Vrednost otpremnice je vrednost njenih izvora**: `VrednostIzvoraPoOtpremnici` sabira `Kolicina × Cena`
  otkupnih stavki članova (`tblOtpremnicaIzvori`). Strog pristupnik (`VrednostIzvoraZaOtpremnicu`) odbija
  izdatu otpremnicu bez izvora po imenu.
- Izveštaji vozača računaju vrednost iz izvora. Štampa uzima **prosečnu cenu izvora po klasi** (vrednost/kg).
- Mreža F2 **nema kolonu vrednosti** za otpremnicu: nacrt nema izvore, a i izmišljena cifra i nula bi lagale.

**P1 — nacrt je ulazio u otpremljenu robu.** „Roba po vozaču“ i „roba po OM“ su brojale svaku nestorniranu
otpremnicu, pa je nacrt od 1000 kg, bez ijednog izvora, već bio otpremljena roba. Ispravka: jedno pravilo
`IzdatoStatusJeIzdato` (i `OtpremnicaJeIzdata` ga sada koristi). Operativni čitaoci — oba izveštaja vozača,
izveštaj po OM i štampa — broje **samo IZDATO**. Štampa odbija nacrt porukom `PRINT_OTP_NIJE_IZDATA`, a ne
opštim „nije pronađena“. Mreža F2 i dalje vidi nacrte, jer baš tu operater radi sa njima.

**P1 (strateški) — F4 je S6, ne S4.** Kanonski redosled: S4 = Zbirna cutover, S6 = Prijemnica header + stavke
+ izvori + F4. Pauza F4, njena poruka, test (`T_PrijemnicaUnos_PauziranDoS6`) i spisak iz §14.14 su prebačeni.
Stavke prijemnice idu u **S6**, a pun lanac (golden) se vraća tek kad downstream slajsovi postoje. Inače bi S4
morao ili da oživi staru prijemnicu, ili da uradi pola S6.

**Fixture.** Da operativni izveštaji nad fixture-om ne mere 0 = 0, `make_fixture.py` izvodi
`tblOtpremnicaIzvori` iz stare veze `Otkup.OtpremnicaID` (ista činjenica, zapisana na drugom mestu). Otpremnica
sa izvorom dobija `IzdatoStatus = IZDATO`: 18 izvora, 17 izdatih. Otpremnica bez izvora ostaje bez statusa, pa
nije otpremljena. `modTest` nijednom ne zove `IspravkaOtkupa_TX` nad fixture otkupima, pa se kapija A13 ne dira.
Potpis fixture-a se menja, pa je **potrebna regeneracija**.

**Testovi:**
- `Test_OTP_OtpremljenoJeSamoIzdato`: nacrt 600 + 400 ne ulazi ni u robu po vozaču ni po OM; posle izvora i
  izdavanja ulazi tačno 1000 kg i dva reda po OM; stornirana izdata izlazi.
- `Test_OTP_VrednostIzIzvoraNePredlogCene`: dva bloka iste klase po 50 i 40 din u otpremnici sa predlogom 999;
  vrednost je 23000, ne 500 × 999.
- `Test_OTP_MrezaCitaStavke`: mreža nema kolonu vrednosti.
- `Test_OTP_IzvestajOMRedPoKlasi`: otpremnica se pravi iz izvora, dakle izdata.
- `modTest` (slaganje izveštaja): ručni prolaz koristi istu definiciju „otpremljeno“ (IZDATO + sirove stavke),
  a ne zaglavlje.

Sabotaže: `otp-nacrt-je-otpremljena-roba` i `otp-vrednost-iz-predlog-cene` su nove. `izvestaji-roba-vozaci-storno`
je prebačena na `Test_OTP_OtpremljenoJeSamoIzdato`: nad starim fixture redovima više ne može da ugrize, jer
stornirana nema aktivno članstvo, pa strog čitač vrednosti padne pre poređenja kilaže. Helper testa vraća grešku
kao tekst, pa tvrdnja pada po imenu. Ukupno 482.

`T_PrijemnicaUnos_PauziranDoS6` i poruka `DOKUNOS_ERR_PRIJEMNICA_PAUZIRANA` sada kažu „dok prijemnica ne pređe
na nov model“.

#### Drugi krug review-a #362 (19.09.2026)

Pun `run_vba` na `c6e47370` je bio zelen: `RunAllTests` 200/0, BFP 1200/0, Storno 200/0, golden 2/2. Commit
`c6e47370` je popravio grešku koju je `82cbb480` uveo u mrežu F2. `ColCena("OTPREMNICA")` je postao `""`, pa se
`Case ColCena(mk)` poklapao sa svakom kolonom bez izvorne kolone (status). `Null` na mestu vrednosti je to oborio
glasno; nula bi tiho upisala broj u pilulu statusa.

**P1 — `PROSLEDJENO` je izdato.** Pravilo `IzdatoStatusJeIzdato` je priznavalo samo `IZDATO`. Kad sync ubuduće
prebaci otpremnicu u `PROSLEDJENO`, isti fizički dokument bi nestao iz robe po vozaču i robe po OM, a štampa bi
ga odbila. Pravilo je sada eksplicitno:

| status | izdat |
|---|---|
| `DRAFT` | ne |
| `IZDATO` | da |
| `PROSLEDJENO` | da |
| prazno | ne |
| nepoznato | ne |

Stari ugovor `modDokumentInvariant.DocIsIssued` („sve osim DRAFT“, prazno = izdato) **namerno nije ponovo
upotrebljen**. Stari lanac nije imao nacrt, a u novom modelu otpremnica nastaje kao nacrt, pa prazan status
nije dokaz izdavanja. Test `Test_OTP_IzdatoStatusPravilo` pokriva svih pet stanja; sabotaža je
`otp-prosledjeno-nije-izdato`.

**P2 — jedna stavka po klasi.** `StavkeOtpremniceRedovi` sada odbija dve stavke iste klase na istoj otpremnici,
isto kao pisac (`OtpUpisiOcekivano`). Bez toga bi pokvaren dokument u izveštaju bio sabran kao 2 × I, a u štampi
dao dva reda iste klase. Test `Test_OTP_DveStavkeIsteKlaseObaraCitaoce` (sintetička anomalija u transakciji koja
se vraća); sabotaža `otp-citalac-pusta-dve-iste-klase`. Ukupno 484 sabotaže.

#### Treći krug review-a #362 — F8 storno otpremnice po `OtpremnicaID`-u (19.09.2026)

**P1.** Nov nacrt nema `GeneracijaID`, a skrivena kolona F8 za otpremnicu je bila `COL_GENERACIJA_ID`. Zato je
kolona bila prazna, a preflight, uvid i storno su išli **po broju**. Dve stanice istog dana legalno nose isti broj,
pa je uvid sabirao oba dokumenta. Pisac bi dvosmislen broj odbio, ali operater je gledao posledice tuđeg dokumenta.

**Mereno pre koda.** F8 okvir ispravke je identitet otpremnice čitao kao generaciju na oko 12 mesta u tri modula:
uvid (zaglavlje, palete, faktura), `PreviewOtpremnica`, `CorrectionNeedsDialog`, lanac i zastavice, i sva tri
moda (`RunOtpremnicaCorrection`). Sadržaj modova je lanac preko `Otpremnica.BrojZbirne`, a tu vezu od S3a ne piše
nijedan živi put. Prevod okvira na ID bi značio čitanje novog modela kroz staru vezu.

**Odluka (operater, 19.09.2026): varijanta B, isti rez kao S1e za otkup.**

| Šta | Gde |
|---|---|
| Skrivena kolona identiteta otpremnice je `COL_OTP_ID` | `modScrDokumenti.IdKolonaTipa` |
| Otpremnica **nije framework tip** u F8 — samo običan storno | `modStornoDok.TipUFlowDoc` |
| Preflight i izvršenje po `OtpremnicaID`-u: prazan ID se ne pogađa po broju, izvršenje ide kroz postojeći `StornoOtpremnica_TX(id)` | `modStornoDok.StornoRazlog`, `StornoIzvrsi`, nova `OtpremnicaAktivnaPoID` |
| Potvrda imenuje stanicu, dan i kg **baš izabranog** dokumenta (kg iz njegovih stavki) | nova `modStornoDok.OtpremnicaOpis` |
| Zbir kg u uvidu ne napušta zadat identitet: bez kolone generacije strict diže grešku, inače je kilaža nepoznata, a ne zbir po broju | `modStornoImpact.SumActiveOtpStavke` |

Modovi ISPRAVKA/DUPLI/PONIŠTENJE/REŠI KASNIJE za otpremnicu su **PAUZIRANI do S3c/S4** (B-022..B-025).
Ispravka je na završetku pauzirana još od S3a (B-038). Kod okvira ostaje dok ga S3c ne prevede na izvore ili
ne obriše. Zato su testovi koji ga zovu direktno ostali: test 45 sada zove `RunOtpremnicaCorrection` direktno
umesto kroz F8. `StornoOtpremnicaByBroj_TX` ostaje, jer ga zove taj okvir (`RunSimpleStornoOtpremnica`).

Testovi: `Test_OTP_F8StornoPoID` (isti broj na ST1/ST2 → red nosi `OtpremnicaID` → preflight po ID-u, bez ID-a
odbijen → potvrda pokazuje samo B kg → storno dira samo B, A ostaje aktivna); `T_FrameworkIspravke_SamoTriTipa`
(ranije `…SamoCetiriTipa`): otpremnica je na listi „obični“. Sabotaže: `framework-otpremnica-vracen`,
`otp-f8-identitet-generacija`, `otp-f8-storno-po-broju`; `framework-otkup` je preusmerena na zbirnu. Ukupno 487.

#### Četvrti krug review-a #362 — izvor aktivne zbirne se ne stornira (19.09.2026)

Pun prolaz na `7d3d3c3a` je bio zelen (BFP 1220/0). Review je našao P1: `StornoOtpremnica_TX` nije proveravao
kanonsko članstvo, pa je F8 posle trećeg kruga mogao da stornira otpremnicu koja je izvor **aktivne** zbirne
(`tblZbirnaIzvori`). Zbirna bi ostala aktivna, sa članstvom koje pokazuje na storniran izvor. To je povreda
A13/A15.

**Popravka je kapija, ne kaskada.** U jezgru `modStorno.StornoOtpremnica`, pre prve mutacije: ako je otpremnica
član aktivne zbirne (`AktivnaZbirnaZaOtpremnicu`), storno se odbija i imenuje zbirnu. Kapija je u jezgru, a ne
samo u `StornoOtpremnica_TX`, jer jezgro zovu i put po broju i kaskade starog okvira. Stari lanac (veza
`BrojZbirne`) je ne dotiče, jer se članstvo čita isključivo iz kanona. F8 isti razlog daje **pre** potvrde
(`STORNO_ERR_OTP_IZVOR_ZBIRNE`). Šta storno izvora znači posle S4 (zamena zbirne, nova verzija, kaskada) odlučuje
S4.

Test `Test_OTP_IzvorAktivneZbirneSeNeStornira`: kanonska otpremnica → kanonska zbirna preko izvora → F8 odbija i
imenuje zbirnu → pisac odbija → otpremnica i zbirna ostaju aktivne, članstvo netaknuto. Sabotaža:
`otp-storno-izvora-aktivne-zbirne` (ukupno 488).

**Poznato, a namerno nedirano (review P2, uspavano):** `StornoOtpremnicaByBroj_TX` i deo `modStornoFlow` nose stari
model „klasa I/II = dva zaglavlja istog broja“, a `SumActiveOtpStavke` u starom okviru uvida `docID` čita kao
`GeneracijaID`. Otpremnica je iz tog okvira izbačena, pa su obe grane uspavane. S3c ih briše ili zamenjuje logikom
po ID-u.

### 14.16) S3b-2a — radni sto otpremnice nad kanonom, „Izdaj“ i izmena nacrta (19.09.2026)

**Rez S3b-2 (odluka operatera, 19.09.2026).** Radni sto koji je S1b-3 obrisao sa stare veze bio je ~1400 linija
u četiri modula. Zato je S3b-2 rezan na dva:

| korak | sadržaj | stanje |
|---|---|---|
| **S3b-2a** | kapija storna otkupa (review #362, P1); radni sto u F1 (liste OTPREMNICE i BLOKOVI, aktivna otpremnica, traka, prekoračenje, vezivanje posle unosa, veži/ukloni nad redom, „Izdaj“, napuštanje pri promeni OM, datum aktivne otpremnice u `ClearForm`); izmena nacrta u F2 | ovaj PR |
| **S3b-2b** | štampa specifikacije blokova (A-018, A-019, A-021) i lista nevezanih blokova (A-025) | |

**Odluka (operater, 19.09.2026): povezano ≠ očekivano rešava izmena nacrta u F2.** `IzdajOtpremnicu_TX` traži
povezano = očekivano po klasi, tačno. Pre S3b-2 nijedan živi put nije menjao očekivanje nacrta, pa bi nacrt sa
prekoračenjem (ili sa greškom u F2) ostao zaglavljen. Izjednačavanje pri izdavanju je odbijeno: očekivanje bi
postalo formalnost uz jednu potvrdu.

#### Pre-flight S3b-2a

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN | otpremnica = zaglavlje + očekivanje po klasi + izvori (§14.8 t. 1–4); izdaje operater kad je ostatak 0 |
| `IDENTITY` | PROVEN | red liste OTPREMNICE/BLOKOVI i F2 nosi ID u nevidljivoj poslednjoj koloni (prioritet 4); izbor, vezivanje, uklanjanje, izdavanje i izmena idu po ID-u, nikad po broju |
| `CARDINALITY` | PROVEN | otkup je u najviše jednoj aktivnoj otpremnici (pisac + `AktivnoOtpClanstvoPoKanonu`); aktivna otpremnica na radnom stolu je jedna i uvek NACRT |
| `INVARIANTS/OWNER` | PROVEN | A13/A15: otkup u sastavu aktivne otpremnice se ne stornira — kapija u jezgru `modStorno.StornoOtkup` (P1 review #362); sastav i izdavanje piše samo `modDokumenta` |
| `WRITERS` | PROVEN | nijedan nov pisac: `DodajOtpremnicaIzvor_TX`, `UkloniOtpremnicaIzvor_TX`, `IzdajOtpremnicu_TX`, `UpdateOtpremnicaDraft_TX` dobijaju prvog živog pozivaoca; ekran ne piše tabelu |
| `EVENTS` | PROVEN | fizički (roba i gajbe odlaze) = poslovni (dokument važi) = izdavanje; finansijskog nema |
| `DOWNSTREAM` | PROVEN | izdata otpremnica ulazi u izveštaje i štampu (S3b-1 čitaju samo IZDATO); nacrt ne |
| `CAPABILITY` | PROVEN | A-011, A-012, A-022..A-024, A-027, A-028 vraćeni, A-020 zamenjen radnjom `vezi`; A-018, A-019, A-021, A-025 → S3b-2b |
| `ACCEPTANCE` | testovi | v. ispod |
| `PLATFORM` | N/A | runtime kontrole istim obrascem kao pre S1b-3 (`BuildOtpTraka`), bez novih `WithEvents` |
| `LANDING` | čisto | grana od `main` `a8d3bd1a` |

#### Izmene

| Šta | Gde |
|---|---|
| Kapija: otkup u sastavu aktivne otpremnice (nacrt ili izdata) se ne stornira; F1 i F8 kažu razlog pre potvrde | `modStorno.StornoOtkup`, `modStornoDok.StornoRazlog`, `modScrDokumenti.RowAction` |
| Liste OTPREMNICE (čip „otvorene“ = nacrti; očekivano/povezano/ostatak iz stavki, status) i BLOKOVI (sastav kroz strog čitač) | `modScrDokumenti.RowsOtpremnice`, `RowsBlokovi`; nov `modDokumenta.IzvoriOtpremnice` |
| Izbor aktivne otpremnice — samo NACRT; prefill sa cenom po klasi (blokovi, pa `PredlogCena`) | `AktivirajOtpremnicu`, `NacrtRazlog`, `PrefillSpec`/`PrefillZaglavlja`, `CenaKlase` |
| Traka: ukupno / u blokovima / ostatak / cena; semafor **po klasi** (+20 u I i −20 u II nije „spremna“) | `Scr_OtpInfo`; ljuska `BuildOtpTraka`/`RefreshOtpTraka` |
| Upis bloka: pitanje o prekoračenju po klasi, pa vezivanje; pad vezivanja se kaže imenom, otkup ostaje upisan | `Scr_Save`, `PrekoracenjeOpis`, `VeziZaAktivnu` |
| Radnje: veži (SVI), ukloni i izdaj (BLOKOVI); posle izdavanja ekran izlazi iz konteksta | `RowAction`, `UkloniIzAktivne`, `IzdajAktivnu` |
| Promena OM napušta otpremnicu; `ClearForm` zadržava datum aktivne otpremnice | ljuska `NapustiOtpremnicu`, `ClearForm` |
| Izmena nacrta u F2: klik na nacrt popuni formu, snimanje menja taj nacrt; pražnjenje forme i promena režima otkazuju izmenu | `IzaberiNacrtZaIzmenu`, `PrefillNacrta`, `SnimiOtpremnicu`; `modDokUnos.OtpremnicaIzmeniNacrt` (+ `OtpremnicaNacrtIzUnosa`, jedno mesto za upis i izmenu) |

#### Testovi i sabotaže

`Test_OTK_IzvorAktivneOtpremniceSeNeStornira` (DRAFT i IZDATO roditelj; kontrola: van sastava storno prolazi),
`Test_OTP_RadniStoBiraSamoNacrt`, `Test_OTP_RadniStoVeziTrakaIzdaj` (traka „u toku“ → „prekoračenje“ →
„spremna“; izdavanje odbijeno pa prolazi), `Test_OTP_RadniStoListe` (ID u poslednjoj koloni, brojke iz kanona,
BLOKOVI = samo sastav), `Test_OTP_IzmenaNacrtaF2` (kroz `Scr_Save`: nov nacrt ne nastaje, članstvo ostaje,
izdavanje posle izmene prolazi), `T_ClearForm_Ugovor` (datum aktivne otpremnice ostaje). Sabotaže: +11,
`clear-datum` preusmeren na `ClearForm` (sidro je posle vraćanja `NapustiOtpremnicu` pokazivalo na nju);
ukupno 499.

**Prag popisa `otp_linija` 24 → 36 (odluka u ovom koraku).** Svih 12 novih živih mesta su činjenice ZAGLAVLJA,
nijedno linijsko polje: radni sto čita `Vrsta`/`Sorta` (lista) i `Vrsta`/`Sorta`/`TipAmbalaze` (prefill, jedno
mesto za F1 i F2) — +5; pisci koji su tek sada dobili živog pozivaoca (`OtpRequireIzvorValjan`,
`OtpKnjiziAmbalazu`, `OtpIzmeniDraft`) — +7. Grupa meša linijska polja sa činjenicama zaglavlja; razdvajanje
ide u S3e, kad se linijske kolone brišu.

**Ručna provera (ne može se automatizovati):** izgled trake i semafora; klik na red liste OTPREMNICE (prefill,
prelazak na BLOKOVI); upis bloka u F1 sa izabranom otpremnicom — blok je odmah u BLOKOVI, a pri prekoračenju
stiže pitanje; potvrda „Izdaj“; klik na nacrt u F2 i snimanje izmene. Vezivanje posle unosa je u `Scr_Save`
tri linije oko `VeziZaAktivnu`; sama funkcija je pokrivena testom, a spoj kroz ceo upis F1 nije (upis traži
stanica-lock i štampu).

**Poznato, a namerno nedirano:** posle snimanja bloka `ClearForm` vraća cenu iz cenovnika (`AutoFillCena`), ne
sa otpremnice — isto ponašanje kao pre S1b-3.

#### Review #363, prvi krug — izdavanje čita strogo (19.09.2026)

**P1.** Read-model (`GetOtpremnicaProgress`) i izdavanje (`OtpIzdaj`) su očekivano i povezano čitali iz **sirovih**
tabela (`OtpUcitajOcekivano`, `OtpUcitajPovezano`), mimo kanonskih strogih čitača, a nebrojčana vrednost je
postajala nula (`OtpDbl`). Dve stavke iste klase su se sabirale: nacrt sa 400 + 100 kg klase I i izvorom od 500 kg
prolazio je jednakost i postajao IZDATO. To je dokument koji `StavkeOtpremniceRedovi` odbija kao korumpiran.
`Test_OTP_DveStavkeIsteKlaseObaraCitaoce` je merio samo put mreže, pa je suite izgledala kao da je invarijanta
zatvorena.

**Popravka bez nove runde validacija:** oba učitavanja agregiraju preko postojećih strogih čitača —
`StavkeOtpremniceRedovi` i `modOtkup.StavkeOtkupaRedovi`. Read-model i izdavanje sada drže isti ugovor stavki kao
mreža, štampa i izveštaji. Test `Test_OTP_IzdavanjeCitaStrogo`: anomalija u transakciji koja se vraća → read-model
pada po imenu → pisac odbija sa istim razlogom, status ostaje DRAFT; kontrola posle vraćanja. Sabotaža
`izdavanje-guta-korupciju` (ukupno 500).

**P2, popravljen jer je kod nov u ovom PR-u:** izmena nacrta u F2 je prefill sastavljala pod `On Error Resume Next`,
pa je pad strogog čitača mogao da ostavi otvorenu izmenu nad delimičnom formom. Sada `OtvoriIzmenuNacrta` sastavlja
formu **pre** otvaranja; pad čitača ostavlja izmenu zatvorenu (isti test), a ekran prazni formu prethodne izmene da je
sledeće snimanje ne bi upisalo kao nov nacrt.

**Prvi pun prolaz (`2726edd2`): izmena nacrta je odbijala SOPSTVENI broj.** F2 validacija (`modDokUnos.OtpremnicaValidiraj`) je broj proveravala kroz `BrojZauzetUNizu` bez `izuzmiID`, a pisac izmene (`OtpIzmeniDraft`) izuzima svoj red po ID-u. `Test_OTP_IzmenaNacrtaF2` je pao sa „broj je već izdat“, a kao zauzimač je naveden baš nacrt koji se menja. Popravka: unos nosi `izmenaOtpID`, a validacija izuzima samo taj red, isto kao pisac. Test sada meri i drugi smer: tuđi broj istog niza (stanica, dan) se i dalje odbija, a izmena ostaje otvorena. Sabotaža `izmena-nacrta-sopstveni-broj` (ukupno 501). Ostale suite su u tom prolazu bile zelene; testovi strogog izdavanja su prošli (ugnježdena transakcija pisca u test-transakciji radi).

**P2, ostaje za pre S5 (backlog):** `OtpRequireIzvorValjan` prima otkup kao izvor samo kad je status tačno `IZDATO`,
a životni ciklus je `IZDATO → PROSLEDJENO`. Danas nijedan lokalni put ne prebacuje otkup u `PROSLEDJENO`
(`CreateOtkup_TX` piše `IZDATO`), pa nije živ kvar. Pre S5 (PWA sync) izvor mora da prizna i `PROSLEDJENO`, sa istom
eksplicitnom semantikom kao `IzdatoStatusJeIzdato`. Stari `DocIsIssued` se ne koristi, jer prazno stanje tumači
drugačije.

#### Review #363, drugi krug — čitač otkupa drži ugovor pisca (19.09.2026)

**P1.** Posle prvog kruga izdavanje čita izvore kroz `modOtkup.StavkeOtkupaRedovi`. Taj čitač je proveravao
pravila jedne stavke, ali ne i dva pravila koja `CreateOtkup_TX` drži nad dokumentom:
- najviše jedna stavka po klasi;
- bruto koji nije manji od neta.

Isti oblik baga je zato ostao na drugoj strani jednačine. Izvor sa 400 + 100 kg klase I prolazio je kao
500 = očekivano i postajao IZDATO, a bruto manji od neta čitač je tiho pretvarao u „neto“.

**Popravka je u ugovoru čitača, ne u `OtpIzdaj`.** `StavkeOtkupaRedovi` sada odbija:
- dve stavke iste klase na istom otkupu;
- nebrojčan ili negativan `BrutoKg`;
- `BrutoKg` veći od nule a manji od `Kolicina`.

Prazno ili `0` je „unet neto“, isto kako pisac čita nulu (`OtkStavkaBrojOpcion` → ne upisuje je). Pravila tako
dobijaju svi čitaoci otkupa odjednom: mreža, izveštaji, izvozi, `VrednostIzvoraPoOtpremnici`, read-model i
izdavanje.

Test `Test_OTP_IzdavanjeCitaIzvorStrogo`:
- anomalija na izvoru u transakciji koja se vraća, jednom dve iste klase, jednom bruto manji od neta;
- read-model pada po imenu, pisac odbija, status ostaje DRAFT;
- kontrola posle vraćanja.

Sabotaže `otk-citalac-dve-iste-klase` i `otk-citalac-bruto-manji` (ukupno 503).

**P2, zapisano, ne menja se u ovom PR-u.** Komanda jednog dokumenta (`IzdajOtpremnicu_TX`) validira **celu** tabelu
stavki otpremnica i otkupa, jer strogi čitači važe za ceo skup. Pokvaren nepovezan otkup zato može da obori
izdavanje ispravne otpremnice. To nije kvar podataka (fail-closed), nego granica agregata; v. backlog §15.

### 14.17) S3b-2b — specifikacija blokova i blokovi bez otpremnice (19.09.2026)

Drugi deo reza S3b-2 (§14.16): sposobnosti koje je S1b-3 obrisao zajedno sa radnim stolom, jer su stajale na
vezi `Otkup.OtpremnicaID` — štampa specifikacije (A-018, A-019, A-021), opseg datuma nad listom otpremnica
(ostatak A-022) i lista blokova bez otpremnice (A-025).

**Odluka operatera (19.09.2026): A-025 je lista NEVEZANIH, ne samo „izgubljenih“.** Stara lista je brojala samo
blok čija je otpremnica stornirana (`GetLostOtkupBlokovi`: „nije vezan → nije izgubljen“). U kanonu blok ostaje bez
otpremnice na tri načina — upisan bez izabrane otpremnice, uklonjen iz nacrta, oslobođen stornom otpremnice — a
radnja `veži` važi za sva tri. Prva dva slučaja se do sada nisu videla nigde. Kolona **„bila u“** razlikuje treći:
nosi broj stornirane otpremnice (članstvo stornirane ostaje kao istorija).

#### Pre-flight S3b-2b

| Osa | Status | Dokaz |
|---|---|---|
| `DOMAIN` | PROVEN (uz odluku iznad) | otpremnica = zaglavlje + očekivanje + izvori (§14.8); specifikacija je spisak onoga što je otišlo, pa se štampa **samo izdata** — isto pravilo kao `OutputOtpremnicaPDF` (review #362, P1) |
| `IDENTITY` | PROVEN | oznake u ljusci se ključaju po nevidljivoj koloni `OTKUI_HD_IDENT` (`modOtkupUI.RowKeyAt`), ne više po broju iz prve kolone; stari `OtpIdZaBroj` (broj → ID) se **ne vraća**; „po datumu“ čita prikazanu mrežu (`GridBrojRedova` + `GridCell`), ne mapu brojeva; red liste NEVEZANI nosi `OtkupID` |
| `CARDINALITY` | PROVEN | otkup je u najviše jednoj aktivnoj otpremnici (`AktivnoOtpClanstvoPoKanonu`, 1305); otkupno mesto bloka = otkupno mesto otpremnice (`OtpRequireIzvorValjan`); red specifikacije = **stavka** izvora, a najviše je jedna po klasi (`StavkeOtkupaRedovi`, 1924) |
| `INVARIANTS/OWNER` | PROVEN | nova invarijanta ne nastaje; čitanja idu kroz kanonske stroge čitače, pa specifikacija ne može da odštampa sastav koji bi izdavanje odbilo |
| `WRITERS` | N/A | samo čitanje; `veži` iz nove liste ide kroz postojeći `VeziZaAktivnu` → `DodajOtpremnicaIzvor_TX` |
| `EVENTS` | N/A | štampa i lista, bez poslovnog događaja |
| `DOWNSTREAM` | PROVEN | od PDF-a ne zavisi ništa; `SPECIFIKACIJA_PRINT_MODE` i folder za PDF su ostali u kanonu; **B-041** (vrsta „izgubljen blok“ u OPORAVKU) ide u **S3c**, uz storno otpremnice po ID-u |
| `CAPABILITY` | PROVEN | A-018, A-019, A-021, A-025 i ostatak A-022 vraćeni (v. mapu) |
| `PLATFORM` | N/A | polja OD/DO istim obrascem kao pre S1b-3a (`NewTxt` u `zGrid`, promena kroz postojeći `UiChange`), bez novih `WithEvents`; PDF ostaje ručna provera |
| `LANDING` | čisto | grana od `main` `0ceccea1` |

#### Izmene

| Šta | Gde |
|---|---|
| Podaci specifikacije: red = stavka izvora; strogo (nacrt, stornirana, nepoznata, dupla, izdata bez izvora, storniran izvor → greška po imenu) | nov `modPrint.SpecifikacijaBlokovaRedovi` |
| Šablon (14 kolona, nova kolona **Klasa**) i izlaz po `SPECIFIKACIJA_PRINT_MODE` | nov `modPrint.EnsureSpecifikacijaSablon` / `FillSpecifikacijaSablon` / `PrintSpecifikacijaBlokova`; `WS_SPECIFIKACIJA_SABLON` vraćen u `modConfig` |
| Radnje liste otpremnica: `mark`, `spec` (označene ili izabrana), `specdat` (cela prikazana lista) | `modScrDokumenti.Scr_Radnje`, `SpecZaIzbor`, `SpecPoDatumu`, `StampajSpecifikaciju`, `IzdateZaSpecifikaciju` |
| Oznake se ključaju po identitetu reda; nova vrednost `2` u ugovoru radnji („označeni ili izabrani“) | `modOtkupUI.RowKeyAt`, `IdentKolonaMreze`, `RunRowAction`, `RefreshRowActions` |
| Opseg datuma OD/DO iznad liste koja ga prijavi (peto polje reda u `Scr_Liste` = `opseg`) | `modOtkupUI` (gradnja, raspored, `ListaImaOpseg`, `ShowOpseg`, `GridDatOd`/`GridDatDo`), `modScrDokumenti.RowsOtpremnice`, `DatGranica` |
| Lista NEVEZANI + radnja `veži`; skup i kolona „bila u“ iz kanona | `modScrDokumenti.RowsNevezani`; nov `modDokumenta.NevezaniOtkupi` (+ `BivseOtpremniceIzvora`, `AktivnoClanstvoOtpremnica`) |

**Zbirna i kupac na specifikaciji** čitaju se iz kanona zbirne (`AktivnaZbirnaZaOtpremnicu`), pa su **prazni dok
S4 ne vrati F3**. Stara veza `Otpremnica.BrojZbirne` se ne čita — pravilo „novi model se ne čita kroz staru vezu“.

#### Testovi i sabotaže

- `Test_OTP_SpecifikacijaBlokova`: blok sa dve klase daje **dva reda sa svojim cenama** (prosek bi ih izjednačio),
  PDV nadoknada je izdvojena i delovi se sabiraju u ukupno, red nosi broj otpremnice, broj bloka i klasu; blok van
  izabranih otpremnica ne ulazi; nacrt, nepoznata i stornirana otpremnica padaju po imenu; šablon ima kolonu Klasa,
  kolone broja su tekst, a red UKUPNO sabira odštampano.
- `Test_OTP_NevezaniBlokovi`: sva tri načina da blok ostane bez otpremnice su u listi, „bila u“ nosi broj stornirane;
  član nacrta i storniran blok nisu; vezivanje iz liste uklanja blok iz nje.
- `T_Otp_OpsegIOznake` (nad pravom mrežom): ključ oznake je `OtpremnicaID`, ne broj; „po datumu“ uzima sve prikazane
  redove; opseg na dan bez otpremnica prazni listu, granica je uključiva, a nepotpun datum nije granica.
- Sabotaže: +8 (ukupno **511**).
- Prag `popis_citalaca` se ne pomera: specifikacija čita **činjenice otkupa** (vrsta, sorta, stanica), ne linijska
  polja otpremnice.

**Ručna provera (ne može se automatizovati):** izgled PDF-a specifikacije (14 kolona, landscape, red UKUPNO);
„Izaberi više“ → označavanje → „Štampaj specifikaciju“; „Po datumu“ nad opsegom; prekidač NEVEZANI i `veži` sa
izabranim nacrtom; polja OD/DO se sklanjaju kad u redu radnji nema mesta.

#### Review #364, prvi krug — bulk čitač članstva drži ugovor čitača dokumenta (19.09.2026)

**P1.** Specifikacija i lista nevezanih čitaju članstvo **jednim prolazom**, kroz nov javni ulaz
`AktivnoClanstvoOtpremnica` → `AktivnoOtpClanstvoPoKanonu`. Taj prolaz je bio **slabiji** od čitača jednog dokumenta
(`OtpClanovi`): prazne ID-eve je preskakao, a postojanje dokumenata nije proveravao — držao je samo globalnu
jedinstvenost aktivnog članstva.

Posledica je ista klasa greške koju ovaj refaktor uklanja — čitalac korupciju pretvara u **uredan poslovni odgovor**:

- članstvo na **nepostojeći otkup** ulazi u mapu, ali ispada iz `zag` (zaglavlja se čitaju iz `tblOtkup`), pa
  specifikacija odštampa PDF **bez tog izvora**. Kapija „izdata bez izvora“ to ne vidi, jer je zadovoljava **validan
  sibling** izvor;
- članstvo na **nepostojeću otpremnicu** nije u skupu storniranih, pa se broji kao aktivno: slobodan blok se
  proglašava zauzetim i **nestaje** sa liste NEVEZANI.

**Popravka je u kanonskom bulk čitaču, ne u `modPrint`.** `AktivnoOtpClanstvoPoKanonu` sada drži isti skup pravila
kao `OtpClanovi` nad jednim dokumentom: članstvo bez `OtpremnicaID`-a ili bez `OtkupID`-a, roditelj koji ne postoji
tačno jednom, dete koje ne postoji tačno jednom, isti par dvaput i otkup u dve aktivne otpremnice — sve su **tvrde
greške**. Postojanje se meri jednim prolazom po tabeli (`BrojRedovaPoID`), ne `RequireTacnoJedan`-om po članu.
Ugovor tako dobijaju svi bulk čitaoci odjednom: specifikacija, lista nevezanih, `VrednostIzvoraPoOtpremnici`,
`OtpremnicaZaOtkup` (kapija storna) i budući.

Test `Test_OTP_ClanstvoBulkStrogo` pravi korupciju **mimo pisca**, uz **validan sibling izvor** (bez njega bi i stara
kapija pukla, pa bag ne bi bio dokazan): članstvo na nepostojeći otkup obara specifikaciju po imenu, članstvo na
nepostojeću otpremnicu obara listu nevezanih; kontrola pre i posle čišćenja. Sabotaže `clanstvo-bulk-dete` i
`clanstvo-bulk-roditelj` (ukupno **513**).

**P2, zapisano u backlog, ne menja se ovde.** Specifikacija čita **činjenice zaglavlja otkupa** (broj, datum,
kooperant, stanica, vrsta, sorta) direktno iz `tblOtkup`, dok stavke idu kroz strog čitač. Naknadno pokvaren
`BrojDokumenta` tako završi kao prazan broj bloka na papiru umesto kao pad. Pravila pisca se ne prepisuju u
`modPrint` — traži se (ili gradi) kanonski strog čitač zaglavlja otkupa; v. §15.

### 14.18) S3c — ispravka izdate otpremnice u jednom potezu, brisanje okvira modova (20.09.2026)

**Odluke operatera (20.09.2026), pre koda:**

1. **Ispravka izdate otpremnice je JEDNA radnja i JEDNA transakcija.** U F1, nad izabranom izdatom otpremnicom:
   storno stare + nov **NACRT** koji nasleđuje zaglavlje, očekivanje i sve izvore. Operater doradi očekivanje i izda.
2. **DUPLI, PONIŠTENJE i REŠI KASNIJE za otpremnicu se brišu.** DUPLI u kanonu radi tačno ono što i običan storno
   (blokovi se oslobađaju sami i vide se u listi „Bez otpremnice“), a PONIŠTENJE i REŠI KASNIJE nemaju o čemu da
   odluče dok F3 (S4) i F4 (S6) ne postoje — vraćaju se tada, nad kanonom.

**Zašto jedan potez, a ne stari dvokorak.** Stari tok je stornirao odmah, pa čekao da operater snimi zamenu:
između ta dva koraka je postojao prozor u kome je stara oborena a nove nema. Zbog tog prozora su i postojali
pending kontekst (`tblStornoVeze`), završetak po snimanju (B-037/B-038) i MANUAL zadaci (B-039) za slučaj da se u
međuvremenu nešto pomeri. U jednoj transakciji prozora nema: padne li bilo koja kapija, ništa nije stornirano.
Zato je B-039 `INTENTIONALLY REMOVED`, a ne prevedeno.

**Šta nova otpremnica nasleđuje, a šta ne**

| Nasleđuje | Ne nasleđuje |
|---|---|
| datum, stanica, vozač, kultura, tip ambalaže | **broj** — storno ne oslobađa broj (A9), nova dobija sledeći slobodan iz niza (stanica, dan) |
| očekivanje **doslovno** (klasa, količina, gajbe, predlog cene) | status — nova je NACRT, ne izdata |
| sve izvore (članstvo 1:1, kroz `OtpUpisiClanstvo` i njegove kapije) | knjiženje ambalaže — storno stare ga vraća, nova ga knjiži tek pri izdavanju |

Očekivanje se **prepisuje**, ne izvodi iz izvora: izjednačavanje očekivanog sa povezanim odbijeno je još u
S3b-2a (očekivanje bi postalo formalnost). Ako je baš očekivanje bilo pogrešno, menja se u F2 — nacrt se ispravlja
izmenom, a ne novim stornom.

Trag ide po **identitetu**: `tblOtpremnica` dobija `IspravkaOdID` i `ZamenjenSaID` (isti par koji `tblOtkup` ima od
S1e). Stare kolone `IspravkaOd`/`ZamenjenSa` nose **broj** i niko ih više ne piše — brišu se u S3e.

**Nalaz usput: kapija koja je čitala mrtvu vezu.** `modStornoFlow.BlockStornoDriftReason` odbija storno čekiranog
bloka koji je u aktivnoj otpremnici. Pitala je `Otkup.OtpremnicaID` — kolonu koju od S3a **ne piše nijedan živi
put** — pa je uvek vraćala „bezbedno je“ i odbijanje nikad nije stizalo do operatera. Sada čita kanon
(`modDokumenta.OtpremnicaZaOtkup`) i na grešci strogog čitača odgovara **fail-closed**.

**Obrisano (okvir modova za otpremnicu, ~800 linija):** `RunOtpremnicaCorrection`, `CompleteOtpremnicaIspravka`,
`RunSimpleStornoOtpremnica`, `StornoOtpremnicaBrojAtomic_TX`, `StornoOtpremnicaByBroj_TX`, `ScanOtpremnica`,
`GetOtpremnicaIDsByBroj`, `PreviewOtpremnica`, `SumActiveOtpStavke` i sve grane `FLOW_DOC_OTPREMNICA` u uvidu
(`GetChainFlags`, `BuildPonistenjePosledice`, `CorrectionNeedsDialog`, `GetStornoChainRows`, `ActiveBlocksForFlow`,
`modStornoImpact`). Tip `FLOW_DOC_OTPREMNICA` **ostaje** — F8 je i dalje lista i stornira otpremnicu (B-014).

**Testovi i kapije**

- `Test_OTP_IspravkaIzdate` — nacrt se odbija po imenu; izdata daje nov NACRT sa **novim** brojem, istim
  očekivanjem i oba bloka; trag `IspravkaOdID`/`ZamenjenSaID` po identitetu; **pad posle storna ostavlja staru
  AKTIVNOM** (zaglavlje pokvareno mimo pisca, pa generator broja nema niz).
- `Test_OTP_KapijaBlokaPoKanonu` — razlog imenuje aktivnu otpremnicu, slobodan blok prolazi, PONIŠTENJE prolazi.
- Sabotaže: `ispravka-nacrt-prolazi`, `ispravka-bez-clanstva`, `ispravka-ne-stornira-staru`, `ispravka-trag-po-broju`,
  `ispravka-pad-ostavlja-storniranu`, `kapija-bloka-po-staroj-vezi`. Obrisano osam sabotaža čija je meta nestala
  (ukupno **511**).
- Testovi obrisanih modova su **obrisani, ne prepravljeni**: `modTest` 201 → **197** (registar prenumerisan),
  `modTestStorno` bez T11/T16/T21/T22. Dve tvrdnje o samom okviru (uvid nad nestalim identitetom, tekst efekta iz
  kataloga) **premerene su nad zbirnom** — okvir je i dalje njen, a tvrdnja nije bila o otpremnici.
- Prag `otp_linija` 36 → **38**: ispravka prepisuje zaglavlje, pa čita `KulturaID` i `TipAmbalaze`. Obe činjenice
  ostaju u kanonu i posle S3e; grupa ih ne razlikuje od osuđenih linijskih polja dok se ne podeli (§15).
- Šema: `tblOtpremnica` + `IspravkaOdID`, `ZamenjenSaID` (na kraj). **Zatečena sveska ih dobija kroz self-heal.**

**Van opsega, zapisano:** vrsta „izgubljen blok“ na ekranu OPORAVAK (B-041) i njen brojač (B-042) idu u **S3c-2** —
čitač (`modDokumenta.NevezaniOtkupi`) već postoji, ali je OPORAVAK svoja površina i ne staje uz ovaj rez.

### 14.19) S3c-2 — „izgubljen blok“ na ekranu OPORAVAK, nad kanonom (20.09.2026)

Vrsta `IZGUBLJEN_BLOK` u listi `Nedovršeno` (B-041) i njen brojač uz stavku menija (B-042) vraćeni su —
ali nad kanonskim članstvom, a ne nad `Otkup.OtpremnicaID` na kojem je stajao obrisani `GetLostOtkupBlokovi`.

**Ko je „izgubljen“.** Ne svaki blok bez otpremnice. Blok upisan bez izabrane otpremnice je **normalno stanje**:
čeka na radnom stolu i vidi se u F1, lista „Bez otpremnice“ (A-025, S3b-2b). U OPORAVAK ulazi samo onaj koji je
**bio** u otpremnici pa ga je njen storno oslobodio — posao koji je neko započeo i ostavio. Merilo je zato zapis
istorije („bila u“), a ne sama nevezanost; skup računa `modDokumenta.NevezaniOtkupi`, isti čitač kojim radni sto
zna šta je slobodno, pa lista ne može da pokaže blok koji je u stvari zauzet.

**Dedup se ne deli sa osirotelim dokumentima.** Broj bloka i broj prijemnice su dva **različita niza istog oblika**
(`1/ddmmgg`), pa bi zajednički `seen` sakrio red zbog tuđeg broja.

**Pad strogog čitača se ne guta.** Članstvo se čita strogo (od review-a #364), a ovo je ekran koji postoji da
nabroji ono što nije u redu — tiho kraća lista bila bi najgori mogući ishod baš ovde. Zato greška daje **vidljiv
red** sa statusom `GRESKA` i porukom čitača.

**Radnja nad redom je pokazivač, ne mutacija:** „Otkup (F1), lista Bez otpremnice: Veži za otpremnicu“. Vezivanje
ostaje kod kanonskog pisca i identiteta (`OtkupID`), tamo gde i pripada; OPORAVAK je pregled.

**Testovi:** `Test_OPO_IzgubljenBlok` — oslobođen stornom je u listi i **opis imenuje storniranu** · nikad vezan i
član aktivnog nacrta **nisu** · storniran blok izlazi · brojač menija = broj redova liste (B-042) · pokvareno
članstvo (sirov upis mimo pisca) daje **red sa greškom**, pa se posle čišćenja gubi. Sabotaže
`oporavak-blok-nikad-vezan` i `oporavak-blok-guta-gresku` (ukupno **513**).

**Zapisano, ne rađeno ovde (review #366, P2):** lista `Nedovršeno` nudi radnju po **listi** (jedno `danger` dugme `odbaci`), pa ga operater dobija i nad redom koji ga ne prima — uključujući ovaj nov. Mutacije nema (`OdbaciIspravku` odbija red bez `CorrectionID`-a), ali UI nudi radnju za koju unapred zna da nije primenljiva. Read-model to već zna (`actionCode`), samo ga `RowsNedovrseno` ne prenosi u mrežu. Rez traži i izmenu **ugovora ljuske** (`trebaRed` ne ume „zavisi od vrste reda“), pa ide kao svoj korak → §15.

**Time je S3c zatvoren.** Sledeći je S3d.

### 14.20) S3d-1 — auto-lanac hladnjače: kanonski korak otpremnice, aktivacija u S6 (20.09.2026)

**Specifikacija (operater, 20.09.2026).** Auto-lanac postoji **samo za otkupno mesto podešeno kao hladnjača**.
Robu do hladnjače kooperant dovozi **sam**, pa taj prevoz nema firminog vozača — zato je na hladnjačkoj stanici
**vozač-ogledalo (`VozacID = StanicaID`) obavezan, bez obzira na režim**. (Ogledalo na *ostalim* stanicama je
zasebna stvar: malina režim, u kome svaka stanica sama dovozi robu.) Roba se meri **na prijemu u hladnjaču**, pa je
ceo lanac **1:1 sa otkupnim listom, po klasi**:

```
OTK  I 700/70  II 120/12
 └─ OTP  I 700/70  II 120/12
     └─ ZBR  I 700/70  II 120/12
         └─ PRJ  I 700/70  II 120/12
```

i **jedan blok = jedan lanac** (nema agregiranja više blokova u jednu automatsku otpremnicu).

**Review #367: NO-GO, tri P1.** Prva verzija je bila odbijena i prepravljena:

| Nalaz | Šta je bilo | Ispravka |
|---|---|---|
| **automatika nije bila obavezna** | ekran je prvo vezivao blok za aktivan ručni nacrt, pa tek onda zvao lanac; lanac je na članstvo izlazio — dakle ručni nacrt je gasio obaveznu automatiku, a blok postajao član tuđeg dokumenta sa sasvim drugom kilažom | grananje ide **pre** ručnog vezivanja (`LanacVaziZaBlok`); provera članstva ostaje, ali kao **zaštita od dupliranja** (retry), ne kao način da ručni tok pobedi |
| **ogledalo se tražilo, a nije moglo da nastane** | `EnsureVozacMirrorForStanica` je i sam iza `IsMalinaMode()`, pa van malina režima ogledalo za hladnjaču **nijedan put ne bi napravio** — lanac bi tamo uvek stajao | uslov proširen: ogledalo se pravi kad je malina režim **ili** je stanica hladnjača; lanac ga traži kroz taj kanonski idempotentan upis i odlučuje tek po **ponovljenoj** proveri. (Review je predlagao da se zahtev veže za malina režim — operater je precizirao da je vezan za **hladnjaču**, pa je tako i urađeno.) |
| **palila se polovična automatika** | pravila se samo otpremnica, a poruka je upućivala na ručni unos zbirne i prijemnice — koji je **nemoguć** (F3 i F4 tvrdo odbijaju) | lanac ima **prekidač** i default je **OFF do S6**; poruka o ručnom nastavku je obrisana |

**Jedan autoritet nad aktivacijom.** Postojao je prekidač `AUTO_PRIJEMNICA_HLADNJACA` u Podešavanjima („Auto
otpremnica+zbirna+prijemnica (OM=hladnjača)“) koji **nijedan red koda nije čitao**, a nov put ga je ignorisao — dva
gospodara. Sada: **stanica kaže KOJI blok** ide u lanac, **prekidač kaže DA LI lanac radi**
(`modAutoHladnjaca.LanacUkljucen`). Default OFF ostaje dok lanac ne ume da završi ceo `OTK → OTP → ZBR → PRJ`;
polovičan lanac je gori od nikakvog, jer operater ostaje sa izdatom otpremnicom i bez ijednog puta napred.

**Drugi krug review-a #367 — routing ide pre SVIH pravila ručnog toka.** Prva ispravka je grananje stavila posle
upisa, ali je pre njega ostalo pitanje o **prekoračenju aktivnog ručnog nacrta** (`PotvrdiPrekoracenje` →
`GetOtpremnicaProgress(mOtpID)`). Hladnjački blok tom nacrtu ne pripada, a operater bi na „Ne“ izgubio **ceo upis**
zbog ograničenja dokumenta sa kojim blok nema veze. Zato sada postoji i router **po stanici**
(`LanacVaziZaStanicu`), koji odlučuje pre nego što blok uopšte postoji; pitanje se postavlja samo kad se
**pouzdano** zna da blok ide ručnim tokom — i „ide u lanac“ i „ne može da se utvrdi“ ga preskaču.

**Oporavak hladnjačkog bloka nije ručno vezivanje.** Ako lanac padne, blok ostaje nevezan i vidi se u „Bez
otpremnice“ — ali odatle ga je jedan klik („Veži“) mogao smestiti u ručnu otpremnicu i time zaobići obavezan lanac.
Kapija je sada na **granici radnje** (`VeziZaAktivnu` odbija hladnjački blok i blok za koji se put ne zna), a lista
je dobila radnju **„Ponovi auto-lanac“**. Obe radnje stoje u redu radnji i svaka na svojoj granici odbija blok koji
joj ne pripada — red još ne može da nosi svoje radnje (ugovor ljuske, §15).

**Jedno pravilo, ne dva.** `AutoLanacHladnjaca` je čitao stanicu kroz `IsHladnjacaStanica` (fail-open, za prikaz),
dok je router čitao strogo — dve kopije istog pravila koje bi se razišle prvom izmenom. Sada oba koriste isti strog
primitiv (`HladnjacaStrogo`).

**Prekidač je PRIVREMEN, i to je zapisano.** Dok traje refaktor on znači „implementacija je dovoljno kompletna da
sme da se pusti“ — tehnička kapija, ne poslovna opcija. Kad S6 zatvori ceo lanac, stanje `JeHladnjača = DA` uz
`AUTO_PRIJEMNICA_HLADNJACA = NE` postaje **zabranjeno specifikacijom** (za hladnjaču je automatika obavezna), pa
prekidač ili nestaje ili postaje interni/deployment safety prekidač koji normalan tok ne koristi. To je **izlazni
uslov S6**, ne preporuka — v. §15.

**IZLAZNI USLOVI S6 — tvrde kapije, ne backlog:**

1. svaki sledeći dokument **izvodi** stavke iz svog kanonskog roditelja (`CreateAutoZbirnaIzOtpremnice_TX(otpID)`,
   `CreateAutoPrijemnicaIzZbirne_TX(zbrID)`) — lanac im **ne prosleđuje prepisane brojeve**. Tako je 1:1 posledica
   modela, a ne tri vrednosti koje mogu da se raziđu;
2. **transakciona i recovery politika mora biti definisana pre paljenja**: „OTP uspeo, ZBR uspeo, PRJ pao“ ne sme da
   ostane neodlučeno stanje;
3. jednakost **po klasi** na sva četiri nivoa je domenska invarijanta i nosi svoj test;
4. **idempotencija se mora preispitati.** Danas `AutoLanacHladnjaca` izlazi čim otpremnica postoji — tačno dok
   je lanac samo `OTK → OTP`. Sa `ZBR` i `PRJ` isti red postaje zamka: posle ishoda „OTP uspeo, ZBR pao“
   ponovljen poziv izlazi odmah i lanac se **nikad ne dovrši**. Bira se **A)** ceo lanac kao jedna atomska
   transakcija ili **B)** nastavljiv lanac (OTP postoji → proveri/nastavi ZBR…). Napomena stoji i **na tom redu
   u kodu**, da je ne propusti onaj ko pali lanac.

**Testovi:** `Test_HLD_AutoLanacOtpremnica` — isključen prekidač ne pušta blok u lanac i ne pravi ništa · uključen:
izdata otpremnica sa vozačem-ogledalom, kilaža i ambalaža 1:1 **i za dve klase** · **hladnjački blok ide u svoj
lanac i kad je aktivan ručni nacrt** (jezgro odluke) · ponovljen poziv ne pravi drugu otpremnicu · obična stanica i
storniran blok se ne diraju · bez ogledala lanac staje i imenuje stanicu. Sabotaže `lanac-bez-prekidaca`,
`lanac-dupli-poziv`, `lanac-i-na-obicnoj-stanici`, `lanac-storniran-blok` (ukupno **517**).

**Dokazni jaz, zapisan kao jaz (review #367, P2):** ceo `Scr_Save` scenario — aktivan ručni nacrt sa malim
ostatkom, pa hladnjački blok veći od njega, **bez pitanja o prekoračenju** i sa sopstvenim lancem — nije
vožen testom, jer `Scr_Save` povlači ostatak UI/print putanje. Mereni su **suđenje** (`LanacVaziZaStanicu`,
`LanacVaziZaBlok`) i **ishod** (lanac, kapije), a ne njihov redosled u ekranu; taj red je u ručnoj listi.

**Nije mereno testom, ide u ručnu proveru:** da ekran zaista grana pre vezivanja (`Scr_Save` nosi štampu otkupnog
lista, pa se u headless prolazu ne vozi) — isti dogovor kao za vezivanje posle unosa iz S3b-2a.

**Sledeće:** S3d-2 — A13 kapija za NACRT (atomska zamena članstva) i radnja „Ispravi“ nad blokom (B-040).

### 14.21) S3d-2 — A13 kapija za nacrt i „Ispravi“ nad blokom (21.09.2026)

Ispravka otkupa je od S1 imala **pisca bez ijednog živog pozivaoca** (B-040), a A13 kapija ju je odbijala za svaki
blok koji je u otpremnici — i za nacrt i za izdatu — uz poruku „nije dostupno do PR7“, dakle uputstvo bez izlaza.
S3d-2 zatvara oboje.

**Nacrt prolazi, izdata ne.**

| Stanje roditelja | Ishod | Zašto |
|---|---|---|
| **nema roditelja** | ispravka radi kao i do sada | ništa izvedeno ne zavisi od bloka |
| **NACRT** | prolazi uz **atomsku zamenu članstva** u istoj transakciji | članstvo nacrta *jeste* mutabilno; stari izvor izlazi, naslednik ulazi i prolazi iste kapije (stanica, kultura, ambalaža, slobodan) |
| **IZDATO** | odbijeno, ali poruka imenuje **put koji postoji** | sastav izdatog dokumenta je istorijska činjenica (A13): prvo ispravka **otpremnice** (S3c) — nastaje nacrt sa istim blokovima — pa onda ispravka bloka u njemu |

Zamena je `modDokumenta.ZameniOtpremnicaIzvor` — **core koji radi unutar tuđe transakcije**, po istom obrascu kao
`modStorno.StornoOtpremnica`. Pozivalac drži snapshot `tblOtpremnicaIzvori`; bez njega bi pad posle uklanjanja
starog izvora ostavio nacrt **bez ijednog izvora**. Redosled je merljiv, ne stilski: stari izlazi **pre** nego što
naslednik uđe, inače bi kapija „izvor sme da bude u tačno jednoj aktivnoj otpremnici“ videla oba i odbila sam posao.

**Kapije žive u jednom izvoru.** `modOtkup.IspravkaOtkupaRazlog` vraća „“ ili rečenicu za operatera; ekran je pita
**pre** nego što operater počne da kuca zamenu, a `IspravkaOtkupa_TX` istu funkciju diže kao grešku — pisac ne sme
da veruje da je iko pitao. Dva spiska istih pravila bi se razišla prvom izmenom.

**Ekran (B-040).** Radnja **„Ispravi“** stoji nad redom u listama `SVI` i `BLOKOVI`. Forma se puni starim blokom
(`PrefillIzStorniranog`, koji broj namerno izostavlja — ispravka dobija **nov** broj, A9), a sledeće snimanje pravi
zamenu umesto novog dokumenta. Ispravka se gasi **tek posle uspeha** (na grešku operater popravi polja i snimi
ponovo), a prazna forma ili promena režima je otkazuju — inače bi sledeći „nov“ unos tiho postao zamena.

**Ispravka ne prolazi kroz routing posle upisa:** članstvo naslednika je već odlučeno u pisčevoj transakciji, pa ni
hladnjački lanac ni ručno vezivanje nemaju šta da odluče.

#### Review #368 — četiri P1 i dve odluke koje su zaključane

| Nalaz | Šta je bilo | Ispravka |
|---|---|---|
| **glavni scenario nije mogao da prođe** | jezgro `modStorno.StornoOtkup` odbija storno bloka koji je u sastavu aktivne otpremnice — a ispravka ga zove **pre** zamene članstva, pa se do zamene nikad nije stizalo | zamena je podeljena na dva koraka: `IzvadiIzvorIzNacrta` **pre** storna, `UvediIzvorUNacrt` **posle** upisa. Kapija nije zaobiđena — posle vađenja je istina da blok nije ni u jednom sastavu. Storno ostaje pre upisa, jer tako oslobađa novac koji naslednik preuzima |
| **pokvarena ispravka otpremnice iz #365** | isti ključ radnje `ispravi` za dve različite stvari; lista otpremnica nema `OtkupID`, pa ju je nov strážar sa `otkupID = ""` odbijao — a u istom `Select`-u su postojala **dva** `Case "ispravi"`, gde drugi nikad ne dobija red | ključevi su razdvojeni: `ispravblok` (blok) i `ispravi` (otpremnica) |
| **prekoračenje se pitalo i za ispravku** | ispravka **zamenjuje**, ne dodaje: nacrt koji očekuje 60 i ima izvor 60 posle ispravke na 55 ima povezano 55, ne 115 — pitanje bi tražilo potvrdu za količinu koja ne postoji, a „Ne“ bi prekinuo legitimnu ispravku | ispravka se zna **pre** pitanja i preskače ga, kao i hladnjački blok i neizvestan put |
| **`GoTo` je preskakao routing i za blok bez roditelja** | „članstvo je već odlučeno“ važi samo kad je stari **bio** u nacrtu; slobodan hladnjački blok bi tako ostao van **obaveznog** lanca | odluka je izdvojena u `RutaPosleUpisa` — deterministična i merljiva bez forme |

**Ugovor rute posle upisa** (`modScrDokumenti.RutaPosleUpisa`):

```
ispravka bloka koji je BIO u nacrtu   ->  nista   (pisac je clanstvo vec preneo)
ispravka SLOBODNOG hladnjackog bloka  ->  LANAC   (lanac je obavezan i za nju)
ispravka slobodnog obicnog bloka      ->  nista   (naslednik ostaje slobodan)
nov unos                              ->  LANAC ili NACRT, po pravilima
put se ne zna                         ->  nista + RAZLOG (fail-closed)
```

Ispravka slobodnog običnog bloka namerno **ne** ulazi u nacrt koji je slučajno otvoren: original nije bio ni u jednom dokumentu, a ispravka menja dokument — ne njegovu pripadnost.

**Preflight sudi po tačnom statusu.** `OtpremnicaJeIzdata` vraća `False` i za nacrt i za prazan/nepoznat status, pa bi preko nje pokvaren roditelj prošao kao „može“, a pisac bi pukao tek na `RequireOtpDraft`. Sada preflight traži **baš `DRAFT`**; izdata dobija put, a nepoznat status svoju rečenicu.

**Zaključana odluka: ispravka izvora NE prepisuje očekivanje nacrta.** Očekivanje je ono što je operater prijavio; da ga dete tiho menja, nestala bi razlika između „prijavljeno“ i „stvarno doneto“ — a upravo ona zaustavlja izdavanje. Test zato meri da posle zamene 60 → 55 ostaje ostatak 5 kg / 1 gajba i da se nacrt **ne izdaje**. Sabotaža `ispravka-prepisuje-ocekivanje` to obara.

**Testovi:** `Test_OTK_IspravkaBlokaUNacrtu` — blok u nacrtu: nacrt pokazuje na naslednika, storniran blok više nije
izvor, nacrt ima tačno jedan izvor · blok izdate: odbijen, razlog **imenuje tu otpremnicu i put** · **pad ne
razmontira nacrt** (naslednik na drugoj stanici pukne posle storna i upisa; transakcija vraća sve) · slobodan blok
radi kao i do sada. Sabotaže `ispravka-bloka-bez-zamene-clanstva`, `ispravka-bloka-izdata-prolazi`,
`ispravka-bloka-bez-snapshota` (ukupno **522**).

**Nije mereno testom:** da radnja „Ispravi“ zaista puni formu i da sledeće snimanje ide kroz zamenu — `Scr_Save`
nosi štampu otkupnog lista, pa se u headless prolazu ne vozi. Mereni su **kapije** i **pisac**; ekranski red je u
ručnoj listi, isto kao kod izmene nacrta.

**Time je S3d zatvoren.** Sledeći je S3e: brisanje `Otkup.OtpremnicaID`, `VozacID`, `BrojOtpremnice` i linijskih
polja zaglavlja otpremnice, uz podelu grupe `otp_linija`.

### 14.22) S3e-1 — mrtav kod stare veze i podela grupa u popisu (21.09.2026)

**Merenje je oborilo pretpostavku koraka.** Plan je S3e (brisanje `Otkup.OtpremnicaID`, `VozacID`,
`BrojOtpremnice` i linijskih polja zaglavlja otpremnice) stavio **pre** S4 i S5. Popis čitalaca kaže suprotno:
poslednji živi čitaoci tih kolona su **kaskada zbirne** (`AutoCreateZbirnaFromOtpremnice`,
`LinkZbirnaToOtkupAndOtpremnica`, `FreeOtkupBloksInline`, `ProizvodjacByZbirna`) i **OTK list za PWA**
(`BuildOTKSheetRowForOtkup`, `ExportOtkupiAll`, `StampVozacFromStanicaForMalina`) — dakle S4 i S5.

**Odluka operatera (21.09.2026):** S3e se deli. Sada ide samo ono što nema nijednog živog pozivaoca, plus podela
grupa u popisu; kolone se brišu u **S3e-2**, kad brojevi padnu na nulu posle S4 i S5.

**Obrisano (bez ijednog živog pozivaoca, provereno i grepom i `vba_check`-om):**

| Celina | Držala | Zašto je mrtva |
|---|---|---|
| `modDokumenta.ReassignOtkupToOtpremnica_TX` | `Otkup.OtpremnicaID` | poslednji pozivalac (`CompleteOtpremnicaIspravka`) obrisan u S3c |
| `modDokumenta.CalculateManjakByOtpremnica` | `Otpremnica.Kolicina`, `Cena` | bez pozivaoca |
| `modSetup.BackfillOtkupBrojOtpremnice` | `Otkup.BrojOtpremnice` | backfill kolone koja se briše |
| **ceo `modSledljivost.bas`** | `Otpremnica.Kolicina`, `Klasa` | jedini javni ulaz (`GetOtpremnicaKandidatiZaOtkup`) ostao bez pozivaoca kad je S1b-3 obrisao ekran SLEDLJIVOST |

**Nalaz o samom postupku brisanja.** `SetOtkupBrojOtpremnice` je prvo obrisan kao mrtav, pa **vraćen**: `vba_check`
je prijavio `NEDEFINISAN` poziv iz `modStornoFlow.FreeOtkupBloksInline` (kaskada poništenja zbirne). Moj grep ga je
propustio jer poziv nosi komentar na kraju reda, a filter je odbacivao svaki red sa `'`. Zapisano namerno: brisanje
„mrtvog“ koda po grepu bez kapije je tačno ta klasa greške, i jedino ju je checker uhvatio.

**Podela grupa u popisu — prag mora da meri nešto što SME na nulu.**

| Bilo | Postalo | Cilj |
|---|---|---|
| `otk_veze` (OTPREMNICA_ID, VOZAC, BROJ_OTPREMNICE, BROJ_ZBIRNE) | `otk_veza_otp` (prve tri) + `otk_brojzbirne` | prva → **0 u S3e-2**; druga umire sa S4 |
| `otp_linija` (linijska polja **+** činjenice zaglavlja) | `otp_linija` (Kolicina, Klasa, KolAmbalaze, BrutoKg) + `otp_zaglavlje` (TipAmbalaze, Sorta, Vrsta, KulturaID) | prva → **0 u S3e-2**; druga ostaje u kanonu |

Pragovi: `otp_stari_pisac` 0 · `otk_veza_otp` **13** · `otp_linija` **3** (bilo 38 dok je grupa nosila i zaglavlje)
· `otp_cena` **0** (dostignuto — poslednji čitalac obrisan ovde).

**`otp_zaglavlje` namerno nema prag.** Vrsta, sorta, tip ambalaže i kultura ostaju u kanonu; prag nad njima pravi
trenje pri svakoj novoj funkciji (S3c ga je već morao podići 36 → 38 zbog ispravke otpremnice) a ne štiti ništa —
ta grupa nema cilj nula.

**Vlasništvo upisa očišćeno.** `tblOtkup.row_owner` je nabrajao `modAutoHladnjaca`, `modOtkupBlok` i `modSetup`, a
nijedan od njih ga više ne piše (lanac obrisan u S3b-1, panel u S1b-2, backfill ovde). Spisak dozvola koji nabraja
module bez upisa ne štiti ništa — sada su uklonjeni, pa bi slučajan upis iz njih oborio A11.

**Sledeće:** S4 (zbirna na kanon, vraća F3), pa S5 (PWA sync), pa **S3e-2** — brisanje kolona, kad popis pokaže nulu.

### 14.23) S4-1 — sadržaj zbirne se čita sa stavki (21.09.2026)

**Merenje pre reza.** S4 („Zbirna cutover") je prevelik za jedan PR, pa je pre bilo kakve izmene
izmereno šta stvarno stoji:

| Činjenica | Broj | Posledica |
|---|---|---|
| Kanonski pisac (`CreateZbirna_TX` / `…IzIzvora_TX`) postoji od PR3 | — | S4 ne gradi pisca |
| Kanonske tabele `tblZbirnaStavke`/`tblZbirnaIzvori` čita **samo** `modDokumenta` | 0 | **nijedan kanonski čitalac nije postojao** |
| Pisac **namerno** ostavlja prazna `UkupnoKolicina`/`UkupnoAmbalaze`/`Klasa` | 48 mesta ih čita | kanonska zbirna se štampa i prikazuje kao **prazna** |
| Živi čitaoci stare veze `Otpremnica.BrojZbirne` | 39 PROD | 16 u `modStornoFlow` (okvir modova) |
| Živa ulazna tačka starog pisca `SaveZbirnaMulti_TX` | **jedna** (`modDokUnos.ZbirnaUpisi`) | isti oblik kao S3a kod otpremnice |
| Živi putevi koji **prave** zbirnu | **nijedan** | F3, malina auto-zbirna i VOZ uvoz su sva tri pauzirana |

Zbog trećeg reda **redosled „F3 pa čitaoci" pada**: da je pauza skinuta prvo, operater bi dobio zbirnu
koja u bazi postoji a na štampi i u izveštajima je prazna — tačno ona lažno-zelena sposobnost koju
§14.14 zabranjuje. **Odluka operatera (21.09.2026):** prvo čitaoci (ovaj korak), pa F3 (S4-2), pa
okvir storna/ispravke (S4-3), pa malina auto-zbirna (S4-4).

**Granica reza je po pitanju koje čitalac postavlja**, ne po modulu:

- **„Šta piše na ovoj zbirnoj?"** (sadržaj) → S4-1, čita `tblZbirnaStavke`.
- **„Koji dokumenti vise o broju Z?"** (članstvo) → S4-3, zajedno sa okvirom — koji se po odluci
  ZBR-KANON-03 **briše, ne prevodi**.

#### Urađeno

| Celina | Gde |
|---|---|
| Strogi kanonski čitači: `StavkeZbirneRedovi`, `ZbirStavkiPoZbirni`, `StavkeZbirnePoDokumentu`, `StavkeZaZbirnu`, `ZbirnaPoKlasi`, `IzvoriZbirne`, `AktivnoZbrClanstvoPoKanonu` | `modDokumenta` |
| Liste F8: kilaža, gajbe i klasa zbirne idu kroz isti mehanizam stavki koji otkup i otpremnica već koriste (`ovStav`) | `modScrDokumenti` |
| Ciljna lista Oporavka (`RowsAktivneZbirne`) — kilaža po ID-u dokumenta; `RowsAktivni` dobio `kgPoDok`, prijemnica ostaje starim putem | `modScrOporavak` |
| Uvid pred storno (`ZbirnaKgZaUvid`) i prefill stornirane zbirne (`StavkeZbirneZaPrefill`) | `modStornoImpact`, `modStornoDok` |
| Izveštaj „zbirni po vozaču" | `modIzvestaj` |
| Integritet B7: umesto „UkupnoKolicina = 0" (što je kod svake kanonske zbirne tačno) meri **zbirnu bez kilaže na stavkama** | `modIntegritet` |
| Ekranski adapter F3 preimenovan `SaveZbirna` → `SnimiZbirnu` (sudar imena sa piscem, isti rez kao S3b-1) | `modScrDokumenti` |
| Obrisano mrtvo: `CalculateManjak`, `GetAktivneZbirne` (popis: MRTAV) | `modDokumenta` |
| `tblZbirnaStavke` se **izvodi iz `tblZbirna`** — fixture se MORA regenerisati | `tools/make_fixture.py` |
| Domenski ugovor: **ZBR-KANON-01/02/03** | `docs/DOMEN/README.md` |

#### Odluka domena: ZBR-KANON-03 (operater, 21.09.2026)

Kad se izvorna otpremnica promeni ili stornira, zbirna **ne** dobija prepravku u mestu nego **novu
verziju** — storno stare + nova sa preostalim izvorima, jedan potez, jedna transakcija, trag po ID-u.
Isto pravilo koje A13 drži za otpremnicu od S3c. `docs/DOMEN/README.md` je do sada tvrdio suprotno
(„popravi otpremnicu pa rekalkuliši zbirnu"); ispravljeno.

Posledica za S4-3: brišu se `RecalculateZbirnaFromOtpremnice_TX`, `ApplyKlasaRecalc` i grane
`modStornoFlow`-a koje prevezuju decu po broju — **uključujući test
`Test_ZbirnaRecalcInPlace_Auto`, koji tvrdi baš ono što je ova odluka ukinula**.

#### Šta NIJE u S4-1 i zašto

| Ostaje | Razlog | Vraća |
|---|---|---|
| `ValidateZbirnaInvariant`, `SumZbirnaByKlasa`, `RecalculateZbirnaFromOtpremnice_TX` | pitaju za **članstvo** (zbir otpremnica po broju), a i brišu se po ZBR-KANON-03 | S4-3 |
| ceo okvir ispravke/poništenja u `modStornoFlow` (27 mesta stare veze) | isto — briše se, ne prevodi | S4-3 |
| Manjak (`CalculateManjakPreview`, `BuildManjakDict`, `ReportManjak`) i prosek gajbe | spajaju zbirnu i **prijemnicu** po broju; prijemnica prelazi u S6, a polovična konverzija bi ostavila isti join | S6 |
| Stari pisac `SaveZbirna*` | jedini pozivalac je iza pauze F3 | S4-2 |

#### KAPIJA ZA S4-2 — F3 se ne odmrzava dok identitet zbirne nije `ZbirnaID`

> **Ovo je tvrda kapija, ne stavka liste.** Review #370 je premestio ovaj nalaz iz S4-3 u **preduslov
> S4-2**, i to je ispravno: između „F3 radi" i „identitet je sređen" ne sme da postoji nijedan commit.

`modScrDokumenti.IdKolonaTipa("ZBIRNA")` je **`GeneracijaID`**, a kanonski pisac generaciju **ne
upisuje** (`BuildZbirnaHeaderRowData` to i kaže). Dakle kanonska zbirna izgleda ovako:

```
ZbirnaID      = ZBR-7f...      <- identitet
BrojZbirne    = 12/210926      <- poslovna labela
GeneracijaID  = ""             <- ljuska ovo uzima kao identitet reda
```

Dok je F3 pauziran, to ništa ne laže — nijedan živi put ne pravi kanonsku zbirnu. **Počinje da laže
tačno u trenutku kad F3 proradi:** operater napravi dokument, ljuska mu ne nađe stabilan ID i radnje
padnu nazad na `BrojZbirne` — dakle na „poslovni broj = identitet", što je baš ono što ceo refaktor
uklanja (S1e za otkup, #362 za otpremnicu).

**Ista kontradikcija je već u kodu, na dva mesta:**

| Tvrdnja | Gde |
|---|---|
| „GeneracijaID se **ne** piše — identitet je `ZbirnaID`" | `modDokumenta.BuildZbirnaHeaderRowData` |
| „aktivna zbirna **bez** `GeneracijaID` je integritetska greška" | `modIntegritet.Chk_B9_ZbirnaBezGeneracije` |

Obe ne mogu biti kanon. Pravac je jasan i **rešenje NIJE dodati `GeneracijaID` kanonskom piscu** —
time bi dokument opet imao dva identiteta:

```
ZbirnaID      = identitet kanonskog dokumenta
GeneracijaID  = legacy mehanizam koji se uklanja
BrojZbirne    = poslovna labela
```

**S4-2 počinje ovim, pre nego što dotakne pauzu:**

1. `IdKolonaTipa("ZBIRNA")` → `COL_ZBR_ID`;
2. `Chk_B9` usklađen sa tim ugovorom — redefinisan na legacy redove ili uklonjen; kanonska zbirna sa
   validnim `ZbirnaID`-em **ne sme** biti proglašena pokvarenom;
3. okvir storna prestaje da `docID` poredi sa `GeneracijaID`-em (danas to radi
   `modStornoImpact`/`modStornoFlow`);
4. **tek onda** `ZbirnaValidiraj` gubi pauzu.

Ne traži zaseban PR — može biti prvi commit S4-2. Traži da bude **prvi**.

#### Kapije

`vba_check` (sabotaža **526 → 529**: dve iste klase, zaglavlje bez stavki, lista čita zaglavlje),
obe `who_writes`, `gen_schema_module --check`, `popis_citalaca --check` sa dve nove grupe:
`zbr_linija` **30** i `zbr_stari_pisac` **29**. Obe idu na nulu u S4-2/S4-3 — prag je rok, ne opis.

Novi testovi: `Test_ZBR_SadrzajCitaStavkeNeZaglavlje` (zaglavlje je prazno **i** mreža ipak pokazuje
pun iznos — obe strane iste tvrdnje) i `Test_ZBR_CitalacStavkiDrziUgovor` (četiri sintetičke
anomalije, svaka pada po imenu, pa se posle vraćanja meri da je čitalac opet čist).

**Sledeće:** S4-2 — **prvo identitet (kapija gore)**, pa F3 nad kanonom (ekran bira izdate
otpremnice, pauza pada, stari pisac se briše).

### 14.24) S4-2a — identitet zbirne je `ZbirnaID` (21.09.2026)

**Ovo je kapija iz §14.23, isporučena pre F3.** Review #370 ju je premestio iz S4-3 u preduslov S4-2, uz
obrazloženje koje merenje potvrđuje: dok je F3 pauziran ništa ne laže, ali bi između „F3 radi" i
„identitet je sređen" postojao prozor u kom operater pravi dokument bez stabilnog ključa.

**Zašto S4-2 ide u dva PR-a.** Merenje pred rez: aparatura generacije ima **22 reference samo u
`modDokumenta`**, plus `modSetup` (4), `modStornoFlow` (5), `modMasterSync`, `modHelpers`,
`modPaletniList`, `modIntegritet` — a `ZBR-CHILD-01` je vezuje za **decu** (prijemnica, paleta), koja
ostaju na starom modelu do S6. Potpuno uklanjanje generacije zato nije posao ovog slajsa. Rez je
uži i tačno pokriva kapiju: **identitet zbirne u ljusci, u stornu i u integritetu**.

#### Urađeno

| Celina | Bilo | Postalo |
|---|---|---|
| Nevidljiva kolona identiteta (F8) | `GeneracijaID` — kod kanonskog dokumenta **prazna** | `ZbirnaID` |
| `modStorno.StornoZbirna` / `_TX` | `(BrojZbirne, GeneracijaID)` | `(ZbirnaID)` — bira **tačno jedan** red |
| F8 dispečer storna | `StornoZbirna_TX(broj, docID)` | provera `ZbirnaAktivnaPoID` + `StornoZbirna_TX(docID)` |
| Uvid pred storno | `HLI` po paru (broj, generacija) | `HLZ` po `ZbirnaID`; `ZbirnaKgZaUvid(zbirnaID)` |
| `Chk_B9` | „aktivna zbirna bez `GeneracijaID`" | „bez `ZbirnaID`-a **ili sa duplim**" |

**Stari okvir se sam prevodi.** `modStornoFlow` i dalje radi po (broj, generacija) i prevodi ih u ID
kroz `ZbrIdIliGreska` — **fail-closed**: nerazrešen identitet diže grešku umesto da padne na broj.
Prevod stoji na strani okvira, ne u jezgru, pa okvir može da nestane u S4-3 bez ijedne izmene u jezgru.

**Kapija nad vlasnicima broja je obrisana, ne zaobiđena.** `RequireJedanVlasnikPoBroju` je štitila
izbor **po broju**; izbora po broju više nema, pa štiti nešto što ne postoji. Umesto nje jezgro diže
grešku ako dva reda nose isti `ZbirnaID` — to je jedina dvosmislenost koja je u novom modelu moguća.

#### Zašto rešenje nije bilo „dodaj generaciju kanonskom piscu"

Time bi dokument opet imao **dva** identiteta, a ceo refaktor ide u suprotnom smeru:

```
ZbirnaID      = identitet kanonskog dokumenta
GeneracijaID  = legacy mehanizam, umire sa okvirom (S4-3) i decom (S6)
BrojZbirne    = poslovna labela
```

#### Kapije

`vba_check` (sabotaža **529 → 530**), obe `who_writes`, `gen_schema_module`, `popis_citalaca`
(pragovi nepromenjeni — merenje je PROD, a identitet se ne meri brojem referenci).

Nov test `Test_ZBR_LjuskaNosiZbirnaID` meri **oba kraja** iste tvrdnje: da red u mreži nosi baš
`ZbirnaID` (i da kanonski dokument generaciju nema), i da storno tim ID-em obori baš taj dokument.
Bez druge polovine bi kolona mogla da nosi tačnu vrednost koju niko ne koristi.

Dva zatečena testa su preimenovana jer im je ime tvrdilo staru premisu:
`T_Zbirna_ZaglavljePoGeneracijiKaskadaStaje` → `…PoIDKaskadaStaje`,
`T_Integritet_VidiDvosmislenBrojIPraznuGeneraciju` → `…IPrazanIdentitet`. Tvrdnje se **ne menjaju** —
dvosmislen broj i dalje ne sme da odlučuje koji dokument pada; menja se čime se dokument imenuje.

#### Review #371 (NO-GO, tri P1) — identitet je bio presečen na pola

Prvi prolaz je promenio **krajeve** lanca a ne i sredinu, pa je ista vrednost u dva sloja imala dva
značenja. Merenje iz review-a:

| P1 | Šta je bilo | Ispravka |
|---|---|---|
| `StornoZbirna_TX` je posle preimenovanja parametra i dalje slao `brojZbirne` u monitoring — uz `Option Explicit` to je **compile blocker**, koji statički CI ne vidi | zeleni CI ≠ projekat se kompajlira | monitoring dobija `zbirnaID` |
| Preflight `StornoRazlog` je za ZBIRNU zvao `AktivanPoIdentitetu`, koji `docID` tumači kao **generaciju** — kanonska zbirna (generacija prazna) je dobijala „nema nestorniranog dokumenta" i do izvršenja se nije ni stizalo | preflight i izvršenje čitali istu vrednost različito | preflight zove `ZbirnaAktivnaPoID`, isti čitač kao izvršenje |
| Ceo okvir ispravke (`ScanZbirna`, `StornoZbirnaIDetach_TX`, `PonistiZbirnaChain_TX`) je primao `ZbirnaID` a prosleđivao ga kao `gen` | tip/semantika presečeni | okvir prima **`ZbirnaID`**; generaciju **izvodi** iz njega (`GenZaZbirnu`) i koristi je samo za legacy scoping dece |

**Smer je sada jedan**, i to je poenta:

```
ZbirnaID --> mutacija zaglavlja        (direktno)
         --> legacy scoping dece       (izvedena generacija)
```

a ne obrnuto — `ZbirnaID` tretiran kao generacija pa tražen nazad `ZbirnaID`, što bi vratilo
sekundarni identitet kao autoritet.

**Zrno prevoda je fail-closed.** Jedna legacy generacija legitimno pokriva **dva** reda `tblZbirna`
(Klasa I i II starog modela), a `StornoZbirna` po ID-u obara **tačno jedan** — `Keys()(0)` bi od
logičkog dokumenta napravio proizvoljan red i ostavio drugu klasu aktivnom. `ZbrIdIzGeneracije`
prevodi samo kad je pogodak jednoznačan; `ZbrIdPoBroju` isto, za zatečene putanje koje nose samo broj.

**Zatvorena je i rupa koju je otkrila zastarela sabotaža:** kad identitet stiže gotov iz ljuske, niko
više nije proveravao da postoji. `RequireZbirnaPostoji` to radi u strict režimu, i sabotaža sada čuva
baš tu kapiju.

**Nov integration test** `Test_ZBR_StornoKrozLjuskuPogadjaSvojDokument` ide putem kojim ide operater —
preflight → izbor moda → izvršenje — nad **dva dokumenta istog broja** (različiti vozači). Raniji test
je merio krajeve lanca i baš zato nije video pokvarenu sredinu.

#### Review #371, drugi krug — par `(BrojZbirne, ZbirnaID)` nije bio proveren

Prvi krug je dao ispravne **tipove** (broj = labela, ID = identitet), ali par niko nije validirao. Bio
je moguć poziv `(broj = A, zbirnaID = B)`, a posledica nije teorijska:

```
StornoZbirnaIDetach_TX:
    StornoZbirna(zbrID)                 -> stornira ZAGLAVLJE B
    DetachOtpremniceInline(broj, ...)   -> odvezuje DECU A
                                        -> HEADER A ostaje aktivan
```

Identitet i članstvo se opet raziđu — baš klasa problema koju refaktor uklanja.

**Ispravka ide korak dalje od validacije para: broj se čita IZ identiteta.** `RequireZbirnaPar` vraća
kanonski `BrojZbirne` pročitan iz zaglavlja, i nizvodno se koristi **on**; prosleđeni broj je samo
provera zastarelog izbora. Tako labela nizvodno nema autoritet ni kad je tačna.

```
zbirnaID -> zaglavlje -> kanonski BrojZbirne -> scoping dece
```

**Prost storno ne ide kroz okvir**, pa par ima svoju kapiju na granici komande (`ZbirnaParOK` u
`StornoRazlog` i `StornoIzvrsi`) sa porukom o **osvežavanju liste** — zastareo izbor je jedini realan
izvor takvog para, pa poruka govori operateru šta da uradi, ne šta je pokvareno.

**Poredak kapija je bitan.** `PonistiZbirnaChain_TX` razrešava ID tek **posle** svoje kapije
dvosmislenosti; tvrd prevod na vrhu je gutao informativnu poruku („broj je pripadao više vlasnika") i
operater bi dobio generički neuspeh. Kaskada iz prijemnice zato koristi **meki** prevod generacije
(`ZbrIdIzGeneracijeAko`).

**Uvid** razrešava identitet jednom, na ulazu (`ZbrIdUvid`): zadat ID se proverava (postoji i nosi baš
taj broj), a bez njega se ide po broju — i to samo dok je jednoznačan.

**Testovi koji su govorili starim identitetom su prevedeni, ne obrisani:** `ZBR-F4` bira dokument
`ZbirnaID`-em umesto generacijom (tvrdnja ista — ko kaže *koji* dokument dira, prolazi i kad broj nose
dva), a prost storno zbirne u `modTest` šalje ID. Nov `Test_ZBR_BrojIIdentitetMorajuBitiIstiDokument`
meri ukršten par i tvrdi da **ništa** nije dirnuto, uz pozitivnu kontrolu da ispravan par prolazi.

**Sledeće:** S4-2b — F3 nad kanonom (ekran bira izdate otpremnice, pauza pada, stari pisac se briše).

### 14.25) S4-2b — zbirna dobija NACRT (21.09.2026)

**Odluka operatera (21.09.2026), na pitanje kako F3 bira izvore:**

> „Treba da postoji isti princip kao i za otpremnicu. Zbirna može da se unese prvo, i onda da se
> validira unosom otpremnica dok ne bude potpuno pokrivena. A treba i da može bez toga, direktno
> odabirom otpremnica koje je čine. Svakako u F3 mora da postoji pregled svih zbirnih, kao što u F2
> mora da postoji pregled svih otpremnica, kao što u F1 mora da postoji pregled svih otkupnih listova.
> To je conditio sine qua non."

Tri stvari, i sve tri su obavezne:

1. **Nacrt pa pokrivanje** — zbirna nastaje PRE nego što se zna od čega je sastavljena. Operater
   najavi šta nosi, pa dodaje izdate otpremnice dok najava ne bude pokrivena.
2. **Jedan potez** — direktno biranje otpremnica; očekivanje se izvodi iz njih
   (`CreateZbirnaIzIzvora_TX`, postoji od PR3).
3. **Pregled svih zbirnih u F3 ostaje** — lista dokumenata svog tipa je uslov bez kog se ne može, na
   sva tri ekrana.

**Rez:** ovaj PR je **pisac**, ne ekran. Lifecycle mora da postoji pre nego što ekran ima šta da zove,
a mešanje to dvoje je isto što je S3a/S3b-2a već razdvojilo kod otpremnice.

#### Urađeno — lifecycle nacrta

| Ulaz | Šta radi |
|---|---|
| `CreateZbirnaDraft_TX(h, ocekivano)` | zaglavlje **DRAFT** + stavke (najava po klasi). `Nothing` kao očekivanje se odbija — to je privatan signal jednopoteznog puta |
| `DodajZbirnaIzvor_TX` / `UkloniZbirnaIzvor_TX` | članstvo nacrta; uklanjanje traži DRAFT |
| `IzdajZbirnu_TX` | revalidacija svih izvora, pa **najava = povezano po klasi**, pa status IZDATO |
| `ZbirnaJeIzdata`, `ZbrClanovi` | čitači stanja; prazno članstvo je kod nacrta uredno, kod izdate kvar |

**Izvor zbirne je IZDATA otpremnica** — odluka koju je §14.14 ostavila S4. Nacrt otpremnice je najava,
ne roba koja je otišla, pa ne može biti deo prevoznog spiska. `PROSLEDJENO` se računa kao izdato
(review #362): sync ne menja činjenicu da je roba otpremljena.

**Ostale kapije drže da je zbirna JEDAN transport JEDNOG vozača:** isti vozač, i ista vrsta/sorta/tip
ambalaže koje je nacrt najavio. Polje koje zaglavlje **nije dalo** se ne poredi — tada ga definiše
izvor. Tako nacrt ne može da pokupi tuđu robu, a ne mora unapred da zna sve.

**Revalidacija pri izdavanju**, isti razlog kao kod otpremnice: između „dodaj" i „izdaj" prođe vreme,
pa se izvor u međuvremenu može stornirati ili ispraviti. Provera i upotreba moraju biti u istom
trenutku — inače je to TOCTOU.

**Zbirna ne knjiži ambalažu.** Gajbe su knjižene pri izdavanju otpremnice i knjižiće se ponovo pri
prijemu (S6); zbirna je prevozni spisak, ne promena stanja gajbi. Zato `IzdajZbirnu_TX` nema snapshot
`tblAmbalaza` — a to je i razlika prema `IzdajOtpremnicu_TX`, koja ga ima.

#### Kapije

`vba_check` (sabotaža **531 → 534**: pokrivenost, izvor mora biti izdat, izvor ne sme u dve zbirne),
obe `who_writes`, `gen_schema_module`, `popis_citalaca`.

Tri nova testa mere luk, ne pojedinačne pozive: `Test_ZBR_NacrtPaPokrivanje` (nacrt nije izdat →
dodavanje izvora ga ne izdaje → izdavanje menja status), `Test_ZBR_NepokrivenNacrtSeNeIzdaje`
(odbijanje **imenuje** najavu i povezano, pa se dopunom istog nacrta izdavanje dobije — bez te druge
polovine bi tvrdnja bila zelena i da izdavanje uvek odbija) i `Test_ZBR_IzvorMoraBitiIzdatISlobodan`.

#### Review #372 (NO-GO, četiri P1) — pisac nije imao jedan ugovor

Sve četiri su u **novom** modelu, ne u legacy-u:

| P1 | Šta je bilo | Ispravka |
|---|---|---|
| Jednopotezni `CreateZbirna` **nije** proveravao da je izvor izdat, dok put nacrta jeste | ista DRAFT otpremnica odbijena na jednom ulazu, primljena na drugom — i zbirna nastane IZDATA iz robe koja nije otišla | jedna definicija: `RequireOtpValidanIzvorZbirne`, koju zovu **oba** ulaza |
| Nacrt je vrstu/sortu/tip ambalaže čitao iz zaglavlja, a `HdrProveriKljuceve` ih **izričito ne dozvoljava** | svaki nacrt je nastajao prazan i takav se **izdavao** — zbirna bez vrste i sorte, bez ijedne greške | **prvi izvor ih definiše** (`ZbrPreuzmiCinjenice`), sledeći mora da se poklopi; `DodajZbirnaIzvor_TX` zato snapshotuje i `tblZbirna` |
| `RequireCeoBroj` meri samo celobrojnost, pa je **−10 gajbi** prolazilo | pisac pravi dokument koji strogi čitalac odbija | eksplicitna provera `< 0` |
| `NewEntityID` za stavku i članstvo nije proveravan, iako `CreateZbirna` to radi | red bez identiteta prolazi kroz commit — nijedna kasnija radnja ne može da ga pogodi | provera na oba mesta, dokazana seam-om `NewEntityIDPadniTest` |

**Poenta drugog nalaza nije bila prazno polje nego pogrešan izvor istine.** Vrsta, sorta i tip
ambalaže su činjenica **robe**, a robu donosi izvor — zaglavlje ih zato i ne prima. Nacrt kreće
prazan i to je ispravno; kvar je bio što ih niko posle nije popunio.

Četiri nova testa, po jedan na svaki nalaz. Onaj o dva ulaza meri **isti izvor kroz oba** — jer je
kvar bio upravo u razlici među njima. Sabotaža **534 → 537**.

**Sledeće:** S4-2c — ulazna kapija (§14.26), pa ekrani.

### 14.26) S4-2c/1 — ulazna kapija nacrta (22.09.2026)

**Ovo je kapija koju je review #372 postavio pred ekrane, a ne sam ekran.** Doslovno: „odlučiti
lifecycle ZBR DRAFT-a: 1. dodati `UpdateZbirnaDraft_TX` 2. odlučiti šta se dešava sa izvedenim
Vrsta/Sorta/TipAmb kada membership padne na 0 3. **tek onda** F3 forma nad postojećim nacrtom."
Redosled nije kozmetika: forma nad nacrtom koji se ne može izmeniti nije forma nego čarobnjak u
jednom smeru, a ekran koji prikazuje vrstu bez ijednog izvora prikazuje tvrdnju koju niko ne drži.

**1) `UpdateZbirnaDraft_TX(zbirnaID, h, ocekivano)`** — izmena **najave**, članstvo netaknuto. Isti
obrazac koji otpremnica ima od S3b-1 (`UpdateOtpremnicaDraft_TX` → `OtpIzmeniDraft`), uz tri razlike
koje dolaze iz domena zbirne:

| | Otpremnica | Zbirna |
|---|---|---|
| Vrsta/Sorta | iz `KulturaID`, header činjenica — menja se izmenom | **ne postoje kao polja zaglavlja**, izmena ih ne dira (ZBR-KANON-04) |
| Niz brojeva | (stanica, dan) | (vozač, dan) |
| Revalidacija članstva | izvor je otkupni blok | izvor je **izdata otpremnica** |

Očekivanje se piše **iznova** (`ZbrObrisiOcekivano` + `ZbrUpisiOcekivano`), ne dopunjava: klasa koja
je nestala iz najave mora da nestane i iz stavki, inače nacrt meri prema klasi koju više ne tvrdi —
i nastaju dve stavke iste klase, dokument koji strog čitalac odbija.

Dve kapije koje nisu očigledne, i obe imaju svoju sabotažu:

- **Nacrt sme da ZADRŽI svoj broj, a ne sme da preuzme tuđi.** Provera zauzetosti gleda niz
  (vozač, dan); naivno primenjena na izmenu, odbila bi i nacrt koji broj samo zadržava — pa bi svaka
  ispravka kilaže tražila i promenu broja. Zato se sopstveni red izuzima **po `ZbirnaID`-u**
  (`RequireBrojSlobodanUNizu(..., izuzmiID)`), ne po datumu.
- **Izmena zaglavlja revalidira postojeće članstvo.** Nacrt vozača A sa članom vozača A, prebačen na
  vozača B, nosio bi člana koga `Dodaj` nikad ne bi primio. Izdavanje bi to na kraju uhvatilo, ali
  invarijanta ne sme da bude prekršena **između dva klika** — ekran u međuvremenu uredno prikazuje
  sastav koji ne postoji.

**2) Odluka operatera (22.09.2026): članstvo na nuli briše izvedene činjenice** → **ZBR-KANON-04**
(`docs/DOMEN/README.md`). Razmotrene su tri opcije; odbijene su „ostaju kao ograničenje" (operater
nikad nije izabrao to ograničenje niti ga vidi) i „operater ih unosi sam" (obrnula bi odluku review-a
#372 — one su činjenica robe, ne zaglavlja).

Pravilo je **uslovno**, pa se i meri u oba smera: uklonjen jedan od dva izvora → činjenice **ostaju**;
uklonjen i poslednji → **brišu se**. Drugu stranu čuva sabotaža `zbirna-brise-cinjenice-i-sa-clanovima`:
kapija koja briše uvek prošla bi test koji meri samo prazan slučaj. Dokaz da je brisanje stvarno, a ne
kozmetika u prikazu: ispražnjen nacrt prima otpremnicu **druge vrste**, koju bi pre toga
`ZbrRequireIstiAko` odbio.

`UkloniZbirnaIzvor_TX` zato od sada snima i `tblZbirna` — uklanjanje menja i zaglavlje.

**Nedokazano, i prijavljeno kao takvo:** danas ništa ne može da padne **posle** čišćenja činjenica
(ono je poslednji korak transakcije), pa nijedan test ne može da natera rollback baš tog upisa.
Snapshot je tu jer je tačan, ne zato što ga tvrdnja pokriva — S4-3 dodaje korake iza njega i tada
postaje merljiv.

**Review #373, P1 — isti kvar koji je #372 već jednom rešio, na drugoj imenici.** `CreateZbirnaDraft_TX`
je odbijao prazno očekivanje, `UpdateZbirnaDraft_TX` je proveravao samo da kolekcija nije `Nothing`.
Prazna kolekcija je zato prolazila: `ZbrObrisiOcekivano` obriše sve stavke, `ZbrUpisiOcekivano` odradi
nula iteracija, i commit ostavi **zaglavlje bez ijedne stavke** — dokument koji
`RequireZaglavljaZbirneSaStavkama` proglašava korumpiranim. Kako je strog čitalac **registarski**, jedan
takav nacrt obara i čitanje svih ostalih zbirnih.

Zakrpa nije otišla u `UpdateZbirnaDraft_TX` nego u **jezgro kroz koje prolaze oba ulaza**
(`ZbrUpisiOcekivano`), pa „valjano očekivanje" više nema dve definicije. Fail-fast iz `ZbrNapraviDraft`
je **obrisan**, ne dupliran — dve provere iste stvari su tačno ono što je kvar i napravilo.

> **Obrazac vredi zapamtiti:** u #372 je spojena *jedna definicija valjanog izvora*, pa je u #373
> odmah nastalo *dve definicije valjanog očekivanja*. Pravilo koje se proverava u wrapper-u umesto u
> jezgru vraća se kroz svaki nov ulaz. Sledeći ulaz nad nacrtom (`CreateZbirnaIzIzvora_TX` u S4-4, F3
> u S4-2c/2) mora da prođe kroz iste helper-e, ne pored njih.

Pet novih testova, sedam sabotaža (**537 → 544**). **Nijedna linija ekrana u ovom PR-u** — F3 forma,
pregled svih zbirnih, radni sto izvora u F2 i brisanje starog pisca idu u S4-2c/2.

### 14.27) S4-2c/2a — stari pisac zbirne je obrisan (22.09.2026)

S4-2c/2 je isečen na **2a: brisanje starog pisca** i **2b: ekrani**. Razlog je merenje, ne ukus:
popis je za grupu `zbr_stari_pisac` pokazivao 45 pogodaka, ali regex broji i komentare — **stvarnih
pozivalaca je bilo dva**, oba iza eksplicitne pauze:

| Pozivalac | Kapija | Provereno |
|---|---|---|
| `modDokUnos.ZbirnaUpisi` (F3) | `ZbirnaValidiraj` vraća poruku o pauzi **pre svih provera** | `modDokUnos.bas`, S3a |
| `modMasterSync.AutoCreateZbirnaFromOtpremnice` (malina) | `IzvedeniLanacIzPwaDostupan() = False`, raise na ulazu | `modMasterSync.bas` |

**Obrisano:** `SaveZbirnaMulti_TX`, `SaveZbirna_TX`, `SaveZbirna`, `BuildZbirnaRowData`,
`ValidateZbirnaInput` (siroče), `modDokUnos.ZbirnaUpisi`, `AutoCreateZbirnaFromOtpremnice` +
`BackfillOtkupBrojZbirneByOtpremnica`. Nuspojava koja se ne vidi iz naslova: brisanje backfill-a
spustilo je i grupu `otk_brojzbirne` (27 → 25 živih), jer je on bio jedan od pisaca te veze.

**`AutoCreateZbirnaFromOtpremnice_TX` je OSTAO** — bez tela, sa glasnim raise-om. Orkestrator ga
zove iza iste kapije; da smo ga obrisali, malina operater bi dobio tišinu umesto razloga. Isti
postupak koji je S1c primenio na `AutoCreateOtpremniceFromPWA`.

**Test-strana je bila veći posao od produkcione.** 11 testova je staro pisca koristilo kao preduslov.
Podela nije bila „koliko ih ima" nego **šta koji tvrdi**:

- **Tvrdnja je bila sam pisac (4 testa) → prešli na kanonski nacrt.** Nevalidna klasa, storniran broj
  istog vozača (A9), tuđi vlasnik broja, i mapiranje reda po imenu kolone. Ove tvrdnje **nisu smele**
  da ostanu na fixture-u: fixture ne proverava ništa, pa bi im tvrdnja postala placebo — a suite bi
  ostao zelen. Jedna tvrdnja je **namerno nestala**: „dvoklasna zbirna je dva reda" u kanonu nema
  predmet (jedno zaglavlje, dve stavke).
- **Tvrdnja je nizvodna (7 testova + golden) → fixture `ZbrZateceniRed`.** Manjak, generacija palete,
  otvorene fakture, storno po broju. Njima treba **red**, ne pisac.

Fixture pravi **zatečeni** oblik (kilaža/klasa/gajbe na zaglavlju + stavka) i pečati generaciju istim
scope-om (broj + vozač + kupac). To je namerno: čitaoci koji te kolone još čitaju (`zbr_linija`,
27 živih) dobili bi od kanonskog seed-a **nulu**, pa bi test „prolazio" ne merivši ništa. Fixture i ti
čitaoci nestaju **zajedno** u S3e-2.

Isti razlog drži golden: `GldZbirnaZaVozaca` je postao seed, pa `tests/golden/G2_isti_broj_dva_dokumenta.txt`
**ostaje nepromenjen**. Golden fajl koji se menja uz refaktor prestaje da bude sidro.

**Sposobnost koja je nestala, i to upisano:** `Test_MalinaAutoZbirnaFailSignal` je obrisan sa svojom
rutinom — merio je kapije tela kojeg više nema. U katalog (`UI_MIGRACIJA_KATALOG.md` §5, stavke 7 i 8)
upisano je `INTENTIONALLY REMOVED (vraća se u S4-4)` za malina auto-zbirnu i `REPLACED (u toku)` za F3
upis. „Kod još postoji" nije dokaz da operater ima funkciju — ni obrnuto.

**Popis:** `zbr_stari_pisac` **29 → 0** (prag spušten na 0), `zbr_linija` **30 → 27**,
`otk_brojzbirne` 27 → 25.

**Sabotaža 544 → 546.** Stari pisac **nije imao nijedno sidro** — izmereno pre brisanja, pa se pokriće nije izgubilo. Ali kapije broja (zauzet broj, tuđi vlasnik niza) preselile su se u `ZbrNapraviDraft`, gde pokrića nije bilo, pa su dobile svoja dva sidra. Oba anchor-a mora da uključe i sledeću liniju (`Dim zbirnaID As String`): isti blok poziva stoji i u `CreateZbirna`, pa bi kraće sidro pogodilo dva mesta.

**Nalaz o alatu, ponovo:** `vba_check` je ostao zelen i posle brisanja četiri javne funkcije koje se i
dalje zovu iz tri druga modula — statička kapija ne vidi pozive kroz module. Preostali pozivi su nađeni
grep-om. Isti obrazac kao review #371: **statički zeleno nije „projekat se kompajlira"**.

**Nalaz iz prvog prolaza: rollback je vracao POLA dokumenta.** Osam BFP testova je snimalo
`tblZbirna` a ne i `tblZbirnaStavke`. Rollback bi vratio zaglavlje i ostavio stavku -- siroce koje
strog citalac prijavljuje u **tudjem** testu, pa je jedan pokvaren red dao 17 padova sa porukom koja
na uzrok ne pokazuje. Dopunjeno na svih osam mesta; produkcioni rollback-ovi nad `tblZbirna`
(`modStorno`, `modMasterSync`, `modDokumentInvariant`) **ne pisu stavke**, pa je asimetrija bila
iskljucivo test-strana. Isto pravilo koje je S4-1 vec primenio na golden, storno, palete i izvestaje.

> Pravilo koje iz ovoga sledi: **ko snima zaglavlje u rollback, snima i njegove stavke.** Dokument
> je zaglavlje + stavke; vratiti samo jedno znaci proizvesti korupciju, ne ponistiti izmenu.

**P3 iz review-a, svesno ODLOZEN:** komentari u `modDokumenta` jos opisuju stari svet
(`identitet logickog dokumenta je GeneracijaID`, `aktivan red MORA da nosi generaciju`). Popravljati
tekst tranzicionog koda **pre** brisanja samog framework-a napravilo bi vecu zabunu nego sto resava.
Ide u **S4-3**, zajedno sa kodom koji opisuje: `GeneracijaID` ZBR identity framework se brise, i
njegovi komentari sa njim.

### 14.28) S4-2c/2b-1 — jezgro i unos za ekrane zbirne (22.09.2026)

**Merenje je oborilo premisu s kojom sam ušao.** Radni sto nacrta **nije u F2 nego u F1**: F2 je samo
forma nacrta plus lista svih otpremnica, a liste `SVI/BLOKOVI/NEVEZANI`, radnje po redu
(`vezi`/`ukloni`/`izdaj`) i traka napretka žive u F1 — jer su izvori otpremnice baš otkupni blokovi,
a oni su F1-ov predmet. Po toj simetriji radni sto **zbirne** pripada **F2**. Operaterova rečenica
„radni sto za izvore u F2, kao blokovi u F1" nije bila preferencija nego tačan opis arhitekture.

Zato je 2b isečen po **sloju**, ne po ekranu: **2b-1 = jezgro + modul unosa** (sve što je dokazivo
testom), **2b-2 = ekrani** (F3 forma, pregled, F2 radni sto, skidanje pauze). Ekran se ne može
automatski testirati — ostaje klik-checklista — pa sve što JESTE dokazivo ulazi zeleno pre njega.

**Dodato u jezgro (`modDokumenta`):**

| | Šta | Ogledalo |
|---|---|---|
| `GetZbirnaProgress(zbirnaID)` | najavljeno / povezano / preostalo po klasi | `GetOtpremnicaProgress` |
| `NevezaneOtpremnice()` | otpremnice koje čekaju zbirnu | `NevezaniOtkupi` |
| `BivseZbirneIzvora` | istorija: brojevi storniranih zbirnih | `BivseOtpremniceIzvora` |

Dve razlike koje nisu kozmetika. Prva: `NevezaneOtpremnice` nudi **samo IZDATE** otpremnice. Nacrt je
najava, ne roba koja je otišla; da se nudi, operater bi ga izabrao **sa ponuđenog spiska**, vezao ga, i
tek bi ga izdavanje zbirne odbilo — porukom o dokumentu koji mu je sam program ponudio. **Kapija koja
odbija tek na kraju je lošija od spiska koji ne laže.** Druga: `GetZbirnaProgress` vraća **uniju**
najavljenih i povezanih klasa i zove **isti par čitača** koji `ZbrIzdaj` koristi za jednakost — da
prikaz i kapija ne mogu da se raziđu.

**Modul unosa (`modDokUnos`) — `ZbirnaValidiraj` je prepisan, ne odmrznut.** Nestalo je troje:

1. **poređenje sa izvorom po `BrojZbirne`** (`ZbirnaBrojIzvora`, `ZbirnaSeSlazeSaIzvorom` — obrisani).
   Članstvo je zapis, ne labela (ZBR-KANON-01), pa pokrivenost meri **izdavanje** nad vezanim
   otpremnicama, a ne unos nad pogođenim brojem;
2. **traženje vrste i sorte.** One su činjenica robe koju donosi prvi izvor (ZBR-KANON-04) — tražiti
   ih od operatera značilo bi tražiti podatak koji pisac odbija. **F3 forma ih gubi u 2b-2**;
3. **`GeneracijaID` kapije** (`ZbirnaIdentResolve` → `ZbirnaGatePoruka`, obrisana). Identitet je
   `ZbirnaID` od S4-2a, a taj okvir umire u S4-3.

Ostalo je: polja bez kojih dokument ne postoji, bar jedna klasa sa kilažom, nenegativna ambalaža, i
**broj** — suđen **istim alatom kojim ga sudi pisac** (`modBrojevi.BrojOdgovaraKontekstu`,
`BrojZauzetUNizu`), po pravilu „jedna implementacija, dva pozivaoca" (`.claude/rules/testovi.md` §5).
`ZbirnaValidiraj` je dobio opcioni `zbirnaID`: pri **izmeni** nacrt sme da zadrži svoj broj, pa se
sopstveni red izuzima — isto kao u `ZbrIzmeniDraft`.

Novi ulazi: `ZbirnaUpisi` (→ `CreateZbirnaDraft_TX`) i `ZbirnaIzmeniNacrt` (→ `UpdateZbirnaDraft_TX`),
oba kroz **jedan** prevodilac `ZbirnaNacrtIzUnosa`, da upis i izmena ne bi različito čitali ista polja.

**Nalaz usput:** F3 je u katalogu poruka opisivao **pogrešan lanac** — podnaslov „Objedinjavanje
prijemnica u jednu zbirnu", hint „prijemnice čekaju novu zbirnu". Zbirna objedinjuje **otpremnice**;
prijemnica je korak posle nje. Ispravljeno, jer bi nov ekran nosio pogrešnu definiciju dokumenta na sebi.

**Pauza je sada na TAČNO jednom mestu** — `modScrDokumenti.SnimiZbirnu`. Validator više nije pauziran
(pa je merljiv), a operater i dalje ne može da upiše zbirnu dok 2b-2 ne isporuči ekran.

**Dug koji ostaje za 2b-2, upisan da se ne zaboravi:** `NoviZbirnaUnos` još nosi ključeve
`vrsta`/`sorta`/`tipAmb` koje niko više ne čita. Brisanje sada ne bi pomoglo — ekran ih i dalje puni,
pa bi ih `Dictionary` tiho vratio; odlaze zajedno sa poljima forme.

**Review #375, P1 — adapter je POPRAVLJAO unos umesto da ga prenese.** Dva kvara iste klase:

| Unos | Šta je adapter radio | Posledica |
|---|---|---|
| `KolAmbalaze = 20.5` | `L()` → `CLng(20.5)` = **20** | `RequireCeoBroj` u piscu meri vrednost koju operater nije uneo |
| `I = -5`, `II = 100` | prevodilac je klasu I preskočio jer nije `> 0` | pisac dobije uredan **II-only** dokument, minus tiše nestane |

Ni jedno ni drugo nije UX propust nego **gubitak podatka**: unos je semantički promenjen, a nijedna
kapija to ne može da vidi jer original do nje ne stigne. Isti obrazac koji je #372 našao kod izvora i
#373 kod očekivanja — ovog puta na granici **ekran → kanonski DTO**.

Rez: ambalaža ide kao **`Double`**, a prevodilac prenosi **prisustvo**, ne sud o vrednosti — klasa
ulazi u najavu kad je operater za nju bilo šta uneo (`kol <> 0 Or amb <> 0`). Validator je dobio
poruke (negativna kilaža, decimalne gajbe), ali **pisac ostaje poslednja tvrda kapija**.

Razlika koja se čuva i koju test dokazuje: **`0` znači „te klase nema"** (II-only nacrt je legitiman),
**`-5` znači „nevalidan podatak"**. Svođenje ta dva na isto stanje je bio ceo kvar.

**Review #375, P2 — validator je puštao stanje koje pisac odbija.** `Kolicina = 0` uz `KolAmb > 0`
je prolazilo kroz unos, a pisac ga je odbijao (`Kolicina <= 0`). Nema kvara podatka, ali operater bi
razlog video **tek posle upisa** — a ovaj sloj postoji baš zato da ga vidi uz polje. Pravilo je sada
izričito: klasa postoji → `Kolicina > 0`; `Kolicina = 0` → i `KolAmb` mora biti 0. Isto za klasu II kad
je prekidač uključen. Pisac i dalje sudi isto — ovde je poruka, tamo tvrda kapija.

**Review #375, drugi P1 — prekidač „dve klase" je bio TVRDNJA, a adapter ga je tretirao kao FILTER.**
`dveKlase = True` uz praznu II klasu je tiše postajalo **jednoklašna zbirna**: prevodilac je II
prenosio samo „ako ima kilažu ili gajbe". Moj sopstveni komentar iznad tog koda je tvrdio suprotno —
da je prekidač izbor operatera da ta klasa postoji. Komentar je bio tačan, kod nije.

Rez: unutar prekidača II klasa se prenosi **bezuslovno**, pa prazna stigne do pisca i padne na
`Kolicina <= 0`. Validator daje istu tvrdnju uz polje (`OTKUNOS_ERR_KOLICINA_II`). Pozitivna kontrola
je netaknuta: `I = 0/0` uz `dveKlase` i `II = 100/5` je i dalje legitiman II-only nacrt — klasa I se
ne zahteva, poštuje se **značenje prekidača**, ne prisustvo podataka.

> Tri kruga, ista klasa greške na tri mesta: **pravilo primenjeno na pogrešnom sloju**. Zato je
> granica sada izričita: **adapter PRENOSI · validator OBJAŠNJAVA · pisac PRESUĐUJE.**

Uz to je zatvoren i P3: fokus za ambalažu i celobrojnost se razdvaja po klasi (`kolAmb` / `kolAmbII`).
Validator koji pokaže na pogrešno polje šalje operatera da popravlja ono što nije pokvareno.

**Prvi pun prolaz: dva pada u `RunAllTests`, oba tačna posledica reza — i oba u suite koju BFP ne
pokriva.** `T_ZbirnaUnos_PauziranDoS4` je merio pauzu **u validatoru**, a ona se preselila na ekran.

Jedan od ta dva pada je otkrio nešto vrednije od sebe: `T_ScrSave_RutaPoRezimu` je **rutu F3 dokazivao
porukom pauze**. Pauza je sada na granici ekrana, pa bi `Scr_Save` vratio istu poruku i da poziv
**nikad ne stigne do modula unosa** — tvrdnja bi ostala zelena nad pokvarenom rutom. Ruta se sada
dokazuje porukom koju vraća **samo** `ZbirnaValidiraj` (broj zbirne).

> Pravilo: kad se kapija preseli, testovi koji su je koristili kao **posrednu** meru prestaju da mere
> ono što tvrde — i to se ne vidi kao pad nego kao lažno zeleno.

`T_ZbirnaUnos_PauziranDoS4` → `T_ZbirnaUnos_PauzaJeNaEkranu`, i meri **obe** polovine: validator više
nije pauziran i stvarno meri, a ekran i dalje odbija upis i imenuje pauzu.

**Drugi pun prolaz je bio ZELEN, ali se brojka nije poklopila — i to je bio nalaz.** Prethodni prolaz:
BFP **1668**; ovaj: **1671**; između njih BFP fajl nije diran. Razlika od tri tvrdnje je pokazala na
`Test_ZBR_NapredakPokrivanja`: seed druge klase je u prvom prolazu vratio prazno, a
`If Len(otpII) = 0 Then Exit Sub` je **ćutke preskočio ostatak testa** — tačno tri tvrdnje o **uniji
klasa**, koje su i bile poenta tog testa. Suite je bio zelen jer test nije ni izmeren.

Isti obrazac je zatečen na **21 mestu** u ZBR bloku (većina iz S4-2b): svaki neuspeo preduslov je tiho
prekidao test. Sva su prevedena u glasnu tvrdnju pre izlaza.

> **Ukupan broj tvrdnji je merenje, ne ukras.** Neobjašnjena razlika između dva prolaza nad istim
> fajlom znači da je neki test preskočio deo sebe — i to se **ne vidi kao pad**.

Pet novih testova, osam sabotaža (**546 → 554**). **Nijedna linija ekrana.**

### 14.29) S4-2c/2b-2a — F3 piše (22.09.2026)

**Pauza F3 je skinuta.** Ekran je od S3a vraćao poruku o pauzi; sada ide kroz `modDokUnos` do
kanonskog nacrta. `SnimiZbirnu` **prevodi polja i ništa ne sudi** — validacija je u `ZbirnaValidiraj`,
kapija u piscu.

| Šta | Gde | Ogledalo |
|---|---|---|
| nevidljiva kolona `ZbirnaID` u mreži F3 | `Scr_Rows` | `OTPREMNICA` od S3b-2 |
| klik na red otvara izmenu nacrta | `Scr_Event` → `IzaberiZbirnuZaIzmenu` | `IzaberiNacrtZaIzmenu` |
| kapija „izdato se ne menja" pre forme | `ZbrNacrtRazlog` | `NacrtRazlog` |
| forma iz dokumenta | `PrefillZbirnaNacrta` | `PrefillNacrta` |
| snimanje pravi **ili** menja nacrt | `SnimiZbirnu` | `SnimiOtpremnicu` |

**Tri kapije koje sam sebi postavio pre koda, i sve tri drže:** red mreže nosi **`ZbirnaID`**, ne broj —
isti broj smeju da nose dva vozača, pa bi klik po broju otvarao tuđi dokument · **nijedna provera ne
živi u ekranu** · napredak se ne računa u ljusci (dolazi u 2b-2b iz `GetZbirnaProgress`).

**Šta je otišlo sa pauzom:** ključ `DOKUNOS_ERR_ZBIRNA_PAUZIRANA` (poruka koju više niko ne vraća) i
mrtvi ključevi `vrsta`/`sorta`/`tipAmb` iz `NoviZbirnaUnos` — dug upisan u §14.28, sada zatvoren.

**Šta OSTAJE nedorečeno, i to je ulazni uslov za 2b-2b:** polja vrste, sorte i tipa ambalaže u formi F3
još postoje i operater sme da ih kuca, a ništa se ne upisuje — one su činjenica robe koju donosi prvi
izvor. Prefill ih **prikazuje** da bi forma govorila istinu o dokumentu, ali polje koje prima unos a
ne čuva ga je tvrdnja bez pokrića. Odlaze sa radnim stolom.

**Posledica reza:** posle 2b-2a operater pravi i menja nacrt, ali ga **ne može izdati** — izvori se
vezuju tek u 2b-2b. Nacrt bez izvora ništa ne kvari i storno postoji.

`T_ZbirnaUnos_*` je u tri reza merio tri stvari — pauzu u validatoru, pauzu na ekranu, pa rad validatora.
To nije lutanje nego **zapis gde je kapija živela**; ime testa prati kapiju, ne obrnuto.

**Review #376, P2 — PR je tvrdio dokaz koji nije postojao.** `Test_ZBR_EkranPraviIMenjaNacrt` je zvao
`OtvoriIzmenuZbirne` **direktno, sa ID-em koji je već držao u ruci** — pa je preskočio tačno onaj spoj
koji ovaj rez uvodi:

```
Scr_Rows (F3) -> nevidljiva kolona ZbirnaID -> GridCell -> Scr_Event "row:n"
              -> IzaberiZbirnuZaIzmenu -> OtvoriIzmenuZbirne -> Scr_Save -> pisac
```

Takav test bi ostao **zelen** i da mreža prestane da nosi identitet, i da se čita pogrešna kolona. Gore
od toga: komentar u testu je tvrdio „drugi nacrt ISTOG vozača — meta za pogrešno pogađanje", a oba
nacrta su imala **različite brojeve** — scenario kojim se prelaz na stabilan ID opravdava nije bio ni
konstruisan.

`T_ZbirnaKlik_OtvaraSvojDokument` (modTest, gde postoje forma i transakcija) prelazi ceo spoj onim
putem kojim ide operater, nad **dva dokumenta pod ISTIM `BrojZbirne`, različiti vozači**: klik na red
drugog mora da otvori baš njega, a snimanje da ostavi prvi netaknut. Dve sabotaže gađaju baš tu
granicu — mreža bez identiteta, i čitanje kolone **broja** umesto identiteta.

> Pravilo: **test koji sam sebi doda ključ ne meri bravu.** Kad rez uvodi spoj, dokaz mora da počne sa
> one strane sa koje počinje operater.

Tri nova testa, pet sabotaža (**554 → 559**).

### 14.30) S4-2c/2b-2b — radni sto zbirne u F2 (22.09.2026)

**Zbirna je prvi put ceo tok:** najava u F3 → pokrivanje u F2 → izdavanje. Simetrija je ista koju F1/F2
već imaju: dokument se **pravi** u svojoj formi, a **pokriva** u formi svog izvora. Izvori zbirne su
izdate otpremnice, a one su predmet F2 — zato radni sto stoji tamo.

| Lista u F2 | Šta pokazuje | Radnje |
|---|---|---|
| `SVI` | sve otpremnice (zatečeno) | `veži` uz izabran nacrt |
| `ZBIRNE` | **nacrte** zbirnih | klik bira aktivan nacrt |
| `IZVORI` | otpremnice u sastavu | `ukloni`, `izdaj` |
| `NEVEZANE` | izdate otpremnice bez zbirne | `veži` |

**Nisu pravljene nove mreže.** Kolone, zbirovi po stavkama i nevidljiva kolona identiteta dolaze iz
`RedoviZaTip` — istog čitaoca koji puni glavne liste; `RedoviZaSkup` samo **bira koji redovi ostaju**.
Zato lista izvora i lista dokumenata ne mogu da pokazu različite brojeve za isti dokument. Kolona
kilaže se prepoznaje **po tipu** iz opisa kolone, ne po poziciji.

**Ekran ne nosi nijedno pravilo.** Sve radnje idu u kanonski pisac i vraćaju **njegov** razlog:
`DodajZbirnaIzvor_TX`, `UkloniZbirnaIzvor_TX`, `IzdajZbirnu_TX`. „Izvor mora biti izdata otpremnica"
stoji u `RequireOtpValidanIzvorZbirne`, ne ovde.

**Ključevi radnji su svoji** (`vezizbr`/`uklonizbr`/`izdajzbr`) da se ne bi sudarili sa istoimenim
radnjama nad otpremnicom u F1 — ista kontrola, drugi predmet.

Dokaz je podeljen po tome šta se čime može dokazati: **klik-put** (red → nevidljivi `ZbirnaID` → izbor
nacrta → prelazak na izvore) meri `T_ZbirnaRadniSto_BiraSvojNacrt` u `modTest`, gde postoji forma, opet
nad **dva dokumenta pod istim brojem**; **veživanje, uklanjanje i izdavanje** meri
`Test_ZBR_RadniStoVezeIIzdaje` u BFP, gde sve ide u transakciji. Potvrda izdavanja (`MsgBox`) je
operaterova, ne logika — zato se izdavanje meri kroz `IzdajAktivnuZbirnu`, a ne kroz `RowAction`.

**Kapija je uhvatila zastarelo sidro:** posle ovog reza je isti red (`GridCell(red, IdentKolonaIndeks("ZBIRNA"))`)
postojao na **dva** mesta — u F3 izmeni i u F2 izboru — pa je sabotaža `zbirna-klik-po-broju` prestala da
bude jednoznačna. Oba sidra sada nose i sledeći red, a novo mesto je dobilo **svoju** sabotažu.

**Review #377, P1 — strog čitalac upotrebljen na pogrešnom lifecycle grain-u.** Radni sto je sastav
čitao kroz `IzvoriZbirne`, a ta dva čitaoca imaju **različit ugovor**:

| Čitalac | Nad čim | Prazno znači |
|---|---|---|
| `ZbrClanovi` | nacrt | **uredno stanje** — još nije pokriven |
| `IzvoriZbirne` | izdata | **kvar** — diže grešku |

Posledica je bila crvena mreža na potpuno ispravnom stanju: svaki tek napravljen nacrt rušio je listu
**odmah po izboru**, a uklanjanje poslednjeg izvora isto. Strogi čitalac ostaje strog — greška je bila u
tome čime je radni sto čitao.

> Treći put u ovom slajsu ista klasa: **pravilo (ili čitalac) na pogrešnom sloju.** #375 adapter koji
> popravlja unos, #375 prekidač kao filter, #377 strog čitalac nad nacrtom.

**Zašto test to nije uhvatio:** `T_ZbirnaRadniSto_BiraSvojNacrt` je posle klika proveravao da stanje
kaže `IZVORI`, ali **nije ponovo učitao mrežu** — a produkciona ljuska to radi automatski. Nov test
čita redove **direktno kroz `Scr_Rows`**, mimo `ScrGridData` koji grešku guta (`On Error Resume Next`).

**Review #377, P2 — izvorni identitet nije bio dokazan.** `Test_ZBR_RadniStoVezeIIzdaje` je
`VeziZaAktivnuZbirnu` zvao direktno. Nov `Test_ZBR_RadniStoVezePoIdentitetu` vozi ceo spoj
(red → nevidljivi `OtpremnicaID` → `RowAction` → pisac) nad **dve izdate otpremnice pod istim brojem**
— broj otpremnice je jedinstven po stanici i **danu**, pa isti broj na dva dana jesu dva dokumenta.
Testovi rade bez forme: `ShowToast` izlazi kad forme nema, pa je ceo put merljiv u BFP.

**P3** (lista `SVI` nudi `Veži` i nad nacrtom otpremnice) ostaje kako ga je reviewer rangirao: pisac je
bezbedno odbija, a `SVI` je namerno „sve". Ide uz sledeći rez, sa trakom napretka.

Četiri nova testa, tri sabotaže (**559 → 562**).

**Ostaje za sledeći rez:** traka napretka u F2 (`GetZbirnaProgress` — čitalac postoji od 2b-1, prikaz ne)
i uklanjanje polja vrste/sorte/tipa ambalaže iz F3, čime se zatvara poslednji P3 iz review-a #376.

### 14.31) S4-2c/2b-2c-1 — F3 prestaje da traži ono što ne nosi (22.09.2026)

Forma F3 je tražila **cenu**, **tip ambalaže** i prikazivala **vrednost** — a `tblZbirna` kolonu `Cena`
uopšte nema, tip ambalaže je činjenica robe koju donosi prvi izvor (ZBR-KANON-04), a vrednost bi bez
cene uvek bila nula. **Polje koje prima unos a nigde ga ne čuva je tvrdnja interfejsa bez pokrića**, a
nula uz robu izgleda kao podatak.

Sve tri idu kroz postojeći `FldShow` mehanizam po režimu — bez `.frx` izmene i bez nove kontrole.

**P3 iz review-a #377 zatvoren:** lista `SVI` više ne nudi `Veži`. Ona je namerno sveobuhvatna, pa
sadrži i **nacrte** otpremnica; pisac ih odbija, ali odbiti **posle klika** znači ponuditi operateru
nešto što će se sigurno odbiti.

**Dokle ta tvrdnja seže (review #379, P3):** `NevezaneOtpremnice` filtrira po **stanju dokumenta** —
izdata, nestornirana, slobodna. Ne filtrira po **odnosu prema aktivnom nacrtu**: otpremnica drugog
vozača ili druge vrste je i dalje u spisku, a `ZbrRequireIzvorValjan` je odbija. Spisak je dakle „sve
što **može** da bude izvor", ne još „sve što može da bude izvor **ove** zbirne". Sužavanje na aktivan
nacrt je zaseban rez — upisan u backlog, da tvrdnja u dokumentaciji ne bude jača od koda.

**Drugi P3 (`mZbrID` preživljava izlazak iz F2) se NE zatvara brisanjem stanja — i to je nalaz.**
Radni sto zbirne je u F2, a njena forma u F3; radni sto otpremnice je u F1, a forma u F2. Kontekst
**mora** da preživi prelazak između ta dva ekrana, inače se tok prekida na svakom koraku. Problem nije
da stanje živi predugo nego da **nije dovoljno vidljivo** — a to rešava traka napretka, ne čišćenje.
Zato ta stavka prelazi u 2c-2 i nestaje sa njom.

**Review #379, P1 — sakriven je nosač zajedno sa sadržajem.** `fgCena` nije samo cena: u istom okviru
žive `segKlasa1`/`segKlasa2`, **jedini operaterski put do dvoklasne najave**. Sakrivanjem okvira se kroz
F3 više nije mogla napraviti zbirna sa I i II klasom — a pisac je izričito podržava. **Kontrola koja
postoji u nevidljivom roditelju ne postoji za operatera.** Odluka „zbirna nema cenu" se ne menja; gase
se **kutije** cene (`KlasaCenaPoRezimu`), a natpis okvira se svodi na klasu.

**Review #379, P2 — test nije izvršavao put koji tvrdi.** Prvi pokušaj je postavljao `ActiveMode` i
zvao `GridRenderTest` — a to je `LayoutGrid` + `RenderGrid`, dakle **mreža, ne forma**. Vidljivost polja
postavlja `ApplyFormFields`, do koga se stize samo kroz `SelectMode`. Test je zato merio formu koja je
ostala u stanju iz gradnje (F1). Sada ide kroz `modOtkupUI.SelectMode f, "F3"` / `"F1"`.

> Treći put u ovom lancu ista klasa: **dokaz presecen pola koraka prerano.** #376 test je sam sebi
> dodao ključ, #377 nije ponovo učitao mrežu, #379 nije prešao režim. Zajedničko im je da svaki put
> ostaje ZELENO — pa razlika između „prolazi" i „meri" nije vidljiva iz rezultata.

Dva nova testa, dve sabotaže (**562 → 564**). Vidljivost polja se ne može automatski izmeriti —
ide na operatersku checklistu (`.claude/rules/testovi.md` §7).

**Ostaje za 2c-2:** traka napretka iz `GetZbirnaProgress` (uz odluku šta pokazuje četvrta grupa mera,
jer zbirna nema cenu) i vrsta/sorta iz kontekstne zone, koje traže raspored — oba diraju ljusku.

### 14.32) S4-2c/2b-2c-2 — traka napretka zbirne (22.09.2026)

Traka iznad forme sada se crta i u **F2**, nad aktivnim nacrtom zbirne. Čita **`GetZbirnaProgress`**,
a on isti par čitača koji `ZbrIzdaj` koristi za jednakost — pa „pokriveno" na ekranu i „pokriveno" na
kapiji **ne mogu da se raziđu**.

**Odluka operatera (22.09.2026): četvrta mera je BROJ IZVORA.** Za otkup i otpremnicu je to cena;
zbirna je nema (`tblZbirna` tu kolonu ni nema), a pokrivenost zbirne **i jeste pitanje članstva** — pa
je broj izvora jedina mera koja prirodno zauzima to mesto. Odbijeno: sakriti četvrtu grupu (prazan
prostor, a operater i dalje mora u listu da vidi ima li nacrt ijedan izvor) i prikazati odredište
(činjenica zaglavlja koja se ne menja dok operater radi — traka postoji za ono što se menja).

**Natpise šalje EKRAN, ne ljuska** — traka je dobila **14. polje**: četiri ključa poruka, zarezom
razdvojena. Ljuska ostaje glupa: ne zna šta je u kom režimu predmet rada i **ne sme da pogađa**. Ekran
koji ih ne pošalje ponaša se kao pre, pa F1 nije dirnut. Uz četvrtu meru ide i pravilo prikaza: kad
ekran pošalje svoje natpise, ta grupa je **ceo broj bez podnaslova** — decimale i „po otpremnici" su
osobina cene, ne mere.

Prve tri grupe takođe menjaju wording: „u blokovima" nema smisla za zbirnu čiji su izvori otpremnice.

**Semafor je isti ugovor** kao kod otpremnice: `-1` neka klasa **prekoračena**, `0` sve na nuli
(spremna za izdavanje), `1` u toku. Test ga meri **u oba smera** — prazan nacrt, pokriven, i višak —
jer je to jedini broj zbog kog traka i postoji.

**Kapija je treći put uhvatila zastarelo sidro:** semafor sada postoji na dva mesta (otpremnica i
zbirna), pa `traka-prekoracenje-nevidljivo` više nije bilo jednoznačno. Oba sidra nose i sledeći red,
a novo mesto je dobilo **svoju** sabotažu.

**Review #381, P2 — `IIf` evaluira OBE grane.** `IIf(imaKljuceve, CStr(kljucevi(0)), "OTKUI_OTP_UKUPNO")`
je nad praznim nizom pucao **i kad je uslov False**. `RefreshOtpTraka` počinje sa `On Error Resume Next`,
pa se greška gutala, dodela preskočila, i **F1 traka je ostajala bez natpisa** — brojevi bez zaglavlja,
i to tiho. Tvrdnja „F1 nije dirnut" nije bila tačna.

Izbor natpisa je zato izdvojen u `TrakaNatpisi`, koja **uvek** vrati četiri ključa; spec pogrešne dužine
ili sa praznim članom se odbija **u celosti** — pola natpisa je gore od nijednog, jer izgleda kao podatak.

**Review #381, drugi P2 — traka je mogla da kaže SPREMNA nad stanjem koje izdavanje odbija.**
`GetZbirnaProgress` meri količine; `ZbrIzdaj` pre toga **revalidira izvore** (storniran, više nije izdat,
tuđi vozač, druga vrsta/sorta/tip ambalaže). Kad se već vezana otpremnica stornira, brojevi ostaju isti
— 400 je i dalje 400 — pa je semafor bio zelen, a izdavanje je padalo.

Rez: `ZbrIzvoriNevaljaniRazlog` je **jedna implementacija sa dva pozivaoca** — `ZbrIzdaj` je diže kao
grešku, traka je pokazuje kao **stanje**. Isti obrazac koji je #372 uveo za izvor i #373 za očekivanje.

> Četvrti put u ovom slajsu: **dva mesta racunaju isti sud.** Kad god ekran i kapija odgovaraju na isto
> pitanje, odgovor mora da ima jedno telo — inace se raziju tiho, a ekran je taj koji laze.

**Review #381, drugi krug — ista greška, pomerena za jedan red.** Prva popravka je sklonila `IIf` iz
**dodela** natpisa i vratila ga **u argument**: `TrakaNatpisi(IIf(UBound(p) >= 13, p(13), ""))`. VBA i
dalje evaluira obe grane, F1 i dalje šalje tačno 13 polja, `On Error Resume Next` i dalje guta
`Subscript out of range` — pa je F1 traka i dalje ostajala bez natpisa, ili sa **tuđim** koji su ostali
iz F2. Pozivno mesto sada nema nijedan uslovni izraz nad poljem koje možda ne postoji:

```vb
Dim kljucevi As Variant, imaKlj As Boolean, spec As String
If UBound(p) >= 13 Then spec = CStr(p(13))
kljucevi = TrakaNatpisi(spec, imaKlj)
```

**Važniji nalaz je zašto je pobegla dvaput: test je merio POMOĆNIK, a bug je bio u POZIVU.**
`Test_ZBR_TrakaNatpisi` je zvao `TrakaNatpisi` direktno — zelen, tačan, i potpuno slep za red iznad
sebe. Zato `modOtkupUI` dobija seam `TrakaRefreshTest`, a `modTest` test **201
`T_Traka_NatpisiPoRezimu`**, koji ide putem operatera kroz pravo pozivno mesto: **F1 (otpremnica, 13
polja) → F2 (zbirna, 14 polja) → nazad F1**. Treći korak je onaj koji vredi — meri da se natpisi
**vraćaju**, a ne da su samo jednom bili tačni. Sabotaža `traka-cita-polje-koje-f1-ne-salje` reprodukuje
tačno zatečeni bug i obara ga po imenu.

**Review #381, P3 — pola odluke je bilo nevidljivo.** Pozivalac je sam računao `imaKlj` („ekran je poslao
spec"), a `TrakaNatpisi` je odvojeno odlučivala da li je spec **valjan**. Pokvaren spec je zato dobijao
podrazumevane natpise **ali custom formatiranje** četvrte mere (ceo broj bez podnaslova). Sada
`TrakaNatpisi` vraća i `prihvacen`, pa jedna odluka nosi oboje.

Četiri nova testa ukupno u slajsu, pet sabotaža (**564 → 569**); `modTest` 200 → **201**.

### 14.33) S4-3a — ispravka zbirne umire, ne seli se (22.09.2026)

**Merenje je oborilo tri tvrdnje plana pre ijednog reda koda.**

1. **`ZbrIdIliGreska` ne postoji** — nula pogodaka u `src-vba/`. §14.26 je tvrdio da okvir kroz nju
   prevodi (broj, generacija) u ID. Red u planu je tvrdnja, ne dokaz.
2. **„Članstvo" je već isporučeno.** Plan ga vodi kao S4-3 posao, ali `NevezaneOtpremnice` već ide kroz
   `AktivnoClanstvoPoKanonu` + `BivseZbirneIzvora`: otpremnica stornirane zbirne se **već** vraća u
   ponudu, sa istorijom. Isporučeno u S4-2c/2b-1.
3. **Kapija nad osiročenom decom već postoji** — `DUPLI` atomarno stornira + odvežuje otpremnice i
   diže **MANUAL zapis** kad ostanu prijemnica ili palete. Druga kapija bi bila drugi autoritet nad
   istim pravilom.

**Odluka operatera (22.09.2026): ispravka zbirne se ODLAŽE do S6.** ZBR-KANON-03 traži jedan potez
(storno + nova iz istih izvora), po ogledalu `IspravkaOtpremnice_TX`. Ali zbirnu vezuju **prijemnice**,
kolonom `BrojZbirne` — `ZbirnaID` im nije strani ključ nigde u šemi — a nova zbirna dobija **nov broj**
(A9: storno ne oslobađa broj). Svaka prijemnica bi ostala siroče. Odbijeno: pisati relink po broju koji
S6 odmah briše. Do tada F8 nad zbirnom nudi `DUPLI` i `PONIŠTENJE` — obe imenovane, nijedna polovična.

**Obrisano:** `SV_MODE_ISPRAVKA` za zbirnu (`modScrStorno.AkcijeRacun`, `RunZbirnaCorrection`),
`CompleteZbirnaIspravka` i njena grana u `modDokUnos.ZavrsiIspravkuAko`, `RelinkOtpremniceToZbirna_TX`,
`RecalcOrStornoEmptyZbirna_TX` (bio je bez ijednog pozivaoca), `RecalculateZbirnaFromOtpremnice_TX`,
`ApplyKlasaRecalc`, `ValidateZbirnaInvariant`, `SumZbirnaByKlasa`, `IsZbirnaConsistent`,
`ValidateOtpremnicaZbirnaImpact` (takođe bez pozivaoca), audit trojka oko rekalkulacije, i šest
privatnih pomoćnika koji su ostali bez posla.

**Najvažniji nalaz nije brisanje nego šta je invarijanta merila.** `ValidateZbirnaInvariant` je
poredila zaglavlje `tblZbirna` sa zbirom otpremnica **po `BrojZbirne`**. Pod kanonom pisac **ne upisuje
ni jedno ni drugo** — članstvo je zapis (`tblZbirnaIzvori`, po `ZbirnaID`), sadržaj je na stavkama, a
zaglavlje ostaje namerno prazno. Obe strane su bile nule, pa je racun **uvek** javljao `OK`. To je
stajalo i u uvidu pred storno i u golden snimku (`tests/golden/D3_*.txt:19`). Provera koja ne može da
padne nije provera, a izgleda kao da jeste — pa je uklonjena iz oba.

Uvid pred storno je umesto nje dobio **kanonski** red: `ZbirnaStavkeTekst` čita `modDokumenta.ZbirnaPoKlasi`
(isti čitač koji koriste liste i štampa) i piše stvarne kilograme po klasi.

**Test nije otišao uz mod.** `T_ZamenaZbirne_NeDiraDecuTudje` je meren kroz `ISPRAVKA`, ali tvrdnja
(„radnja koja dira decu staje dok broj nose dva aktivna dokumenta") važi i za `DUPLI`, koji decu dira
isto tako. Preusmeren, ne obrisan — da je otišao sa modom, kapija bi ostala bez ijednog merenja.

**Nov test `T_Zbirna_NemaIspravku` (modTest 113)** meri **oba kraja**: zbirna ne nudi `ISPRAVKU` ali
i dalje nudi `DUPLI` i `PONIŠTENJE`, a **prijemnica ispravku i dalje nudi**. Bez drugog kraja bi prazan
red odluke — pokvaren ekran — prošao kao zelen.

**Kapije su uhvatile četiri zastarela sidra** (tri sabotaže bez koda, jedna bez tvrdnje) i **rupu u
numeraciji testova**; `popis_citalaca` je javio da je prag `zbr_linija` zastareo (**27 → 17**) —
merenje ispod praga pada isto kao iznad, jer zastareo prag pušta grupu da naraste nazad bez ijednog
crvenog. Sabotaže **569 → 567**.

### 14.34) S4-4 — malina auto-zbirna nad kanonom (23.09.2026)

Sposobnost se **vraća**, ne prevodi. Staro telo (izvučeno iz `f437661c~1`) je šlo **po redu**
`tblOtpremnica`, uzimalo `Kolicina/Klasa/KolAmbalaze` sa **zaglavlja** — kolone koje od S3b-1 nijedan
pisac ne puni — **nije gledalo `IzdatoStatus`**, i pisalo staru vezu `Otpremnica.BrojZbirne` plus
backfill `Otkup.BrojZbirne`.

**Jedno jezgro, dva pozivaoca (odluka operatera, 23.09.2026).** `AutoZbirnaZaOtpremnicu(otpID)` radi nad
**jednom izdatom** otpremnicom; zovu ga kuka na **izdavanju** (`modScrDokumenti.IzdajAktivnu`) i **batch
prolaz** iz sync orkestratora, za otpremnice koje stignu kroz PWA i kuku nikad ne prođu.

**Odluka šta sme se ne računa u jezgru** — pita se `modDokumenta.NevezaneOtpremnice`, **ista lista koju
operater vidi u F2**. Drugo pravilo na tom mestu značilo bi da ekran i automatika mogu da se raziju:
automatika bi vezala otpremnicu koju spisak ne nudi, ili obrnuto. Ta lista već drži sva tri uslova
(IZDATA, nestornirana, bez aktivnog članstva), pa je jezgro **idempotentno po konstrukciji** — uslov, ne
udobnost, jer batch ume da stigne pre kuke.

**Zbirna dobija SVOJ broj (odluka operatera).** Stari kod je pisao `ApplyMirrorPrefix(vozacID, BrojOtpremnice)`
i sam komentar je to vodio kao **dug**: numerički deo je pripadao **stanici**, a vlasnik niza zbirne je
**vozač** — prolazilo je samo dok je vozač doslovno mirror-stanica. Sada `SuggestNextBroj(KIND_ZBR, vozacID,
datum)`, koji mirror prefiks `S` primenjuje sam, pa se **izgled** broja u malini ne menja — menja se **čiji
je niz**. Dug zatvoren; sidro je bilo `Test_BKTX_ZbirnaTudjegVlasnikaOdbijena`.

**`checkRemote:=False` — auto-zbirna ne pita Google.** Podrazumevano `SuggestNextBroj` gleda i udaljeni
list. Za batch to znači mrežni poziv **po dokumentu**, a automatika koja zavisi od mreže pada **tiho**
(generator grešku guta i vraća prazno) — tako je i pao prvi lokalni prolaz. Bezbedno je jer je VOZ/zbirna
uvoz pauziran; **dug za S5**: kad se uvoz vrati, udaljena osa mora nazad u račun.

**Kapija lanca je razdvojena.** `IzvedeniLanacIzPwaDostupan` je jedna kapija nad celim izvedenim lancem —
tacna dok je auto-zbirna pisala `Otkup.BrojZbirne` nazad na zaglavlje. Kanonska to ne radi, pa je dobila
svoju (`AutoZbirnaDostupna`); **VOZ/zbirna uvoz ostaje pauziran**. Kapija koja pokriva više nego što mora
zaustavlja i ono što je popravljeno, a onda se otvara „u paketu" — tiho puštajući i ono što nije.

#### Nalaz: hladnjački lanac NIJE dobio svoj ZBR korak, i to je namerno

Zaglavlje `modAutoHladnjaca` kaže da se zbirna u S4 dodaje u **istu** funkciju (`AutoLanacHladnjaca`). Ali
ta funkcija nosi **tvrdu kapiju** koju sama imenuje: rani izlazak `If Len(OtpremnicaZaOtkup(otkupID)) > 0
Then Exit Function` je tačan dok je lanac samo OTK→OTP, a postaje zamka čim dobije ZBR — posle ishoda
„OTP uspeo, ZBR pao" ponovljen poziv izlazi odmah i lanac se **nikad ne dovrši**. Izbor između **(A)
atomski lanac** i **(B) nastavljiv lanac** je izlazni uslov S6 (plan §14.20), izričitno „ne stvar ukusa".

Zato S4-4 dira **samo malina tok**. Kad hladnjački lanac dobije ZBR, mora da zove **isto jezgro** —
idempotencija je ono što sprečava dve zbirne za istu otpremnicu na stanici koja je i hladnjača i malina.

#### Placebo tvrdnja koju je sabotaža razotkrila

Prva verzija idempotentne tvrdnje merila je samo da drugi poziv ne vrati `ZbirnaID`. **Sabotaža koja skida
kapiju jezgra nije oborila ništa** — jer i bez nje pisac (`ZbrRequireIzvorValjan`) odbija već vezanu
otpremnicu. Tvrdnja je merila **tuđu** kapiju. Prava razlika je u **tišini**: sa kapijom se drugi poziv ne
desi, bez nje pisac pukne i vrati razlog — koji operater vidi kao grešku posle sasvim normalnog ponovnog
izdavanja. Tvrdnja sada meri `outGreska`. Isti obrazac kao `dvoslojna kapija: sabotaža ne grize`.

Uz to: EH tog testa je radio `On Error Resume Next` **pre** čitanja `Err`, pa je prvi pravi pad prijavio kao
`FATAL 0` bez opisa — uništivši jedini trag. Popravljeno.

Sabotaže **566 → 568**, BFP **1783 → 1792** (+9, sravnjeno po stavkama).

**Review #383 — tri P2 na granicama, sve tri ista klasa: podaci se promene, a sistem tvrdi drugo.**

**P2 #1 — delimičan uspeh se prijavljivao kao neuspeh.** Izdavanje i auto-zbirna su **dva događaja**, i
drugi sme da padne posle prvog **commit-a**. Dok je `IzdajAktivnu` vraćala samo „razlog", pozivalac je svaki
neprazan odgovor čitao kao „ništa se nije promenilo“: mrežu nije osvežavao, a otpremnica je već bila
`IZDATO`. Teži podslučaj: `AutoZbirnaUpis` diže grešku **pre pisca** (prazan kupac, prazan vozač, nema
broja), a jezgro to nije hvatalo — izuzetak je stizao u EH ekrana i postajao **„Otpremnica nije izdata"**,
laž o događaju koji se desio, na osnovu koje bi operater ponovio radnju.

Rez ima tri dela: jezgro **ne diže grešku nego je vraća** (`"" + outGreska`), `IzdajAktivnu` izlazi sa
`outIzdata` („primarna mutacija se desila“), a pozivalac **osvežava i kad je poruka greška**. EH ekrana
bira uvod poruke po `outIzdata`.

**P2 #2 — `LogErr` briše `Err` pre re-raise-a.** Batch je radio `LogErr SRC` pa `Err.Raise Err.Number, ...`.
Repo taj invariant već nosi napisan u `modAutoHladnjaca` („Opis se cita PRE LogErr-a“), a orkestrator
odlučuje da li je korak pao **baš po `Err.Number`** posle `On Error Resume Next` — pa je re-raise sa
obrisanim `Err`-om odnosio i broj i razlog, a s njima i signal.

**P2 #3 — pisac zbirne nije proveravao da veze postoje.** Zbirna je jedina od tri dokumenta gledala samo
`Len() > 0`; otpremnica i otkup odavno traže `RequireTacnoJedan`, uz komentar koji doslovno kaže isto:
*neprazan string nije dokaz da red postoji*. Put unutra: `MALINA_DEFAULT_KUPAC` sa typo-om — `LookupValue`
za hladnjaču vraća prazno **bez greške**, i nastaje **finalna `IZDATO` zbirna sa `KupacID`-em bez pokrića**.
Kapija je smeštena u **pisca**, ne u auto-put, i tri mesta koja su čitala zaglavlje zbirne sada dele
`ZbrHdrCitajIProveri` — pa F3, uvoz i automatika ne mogu da se raziđu.

Sabotaža je usput popravila i sam dokaz: prva verzija tvrdnje „jezgro ne diže grešku“ padala je kao
**FATAL**, dakle po imenu testa a ne po tvrdnji koja to meri. Poziv sada ide pod `On Error Resume Next` i
meri `Err.Number`, pa sabotaža obara baš tu tvrdnju.

Sabotaže **568 → 571**, BFP **1793 → 1807** (+14).

### 14.35) S4-3b — identitet zbirne je ZbirnaID, i to je bio živ kvar (23.09.2026)

**Ušao sam da čistim mrtav kod, a našao kvar koji sam sam napravio.**

`ZbirnaIdentResolve` je tvrdio: *„aktivan red MORA da nosi generaciju; prazna nije alternativni oblik
identiteta nego **integritetska greška**"*. Komentar je imenovao dva pisca koji je pečate:
`modMasterSync` (pauziran) i `modDokumentInvariant` — **koji je S4-3a obrisao**. Kanonski pisac je nikad
nije ni pisao.

Posledica: svaka kanonski napravljena zbirna obarala je kapiju sa `ZBR_MUT_INTEGRITET` **pre** nego
što se mod uopšte bira — pa `DUPLI` i `PONIŠTENJE` **nisu radili ni za jedan dokument koji aplikacija
danas pravi**. A baš njih je S4-3a ponudio kao zamenu za ukinutu ispravku. Poruka je operatera još
slala na B9 proveru koja to ne može da popravi.

To je doslovno obrazac iz `brisanje-pisca-veze-proveri-kapije`: obrisan pisac, a kapija koja je tražila
ono što je on pisao ostala je da stoji.

**Zašto nijedan test nije pao:** `T_ZamenaZbirne_NeDiraDecuTudje` meri na fixture redovima
`ZBI-KASK-1/2`, a oni **nose** `GeneracijaID`. Kapija je bila merena isključivo na legacy podacima;
kanonski dokument ide drugom granom koju nijedan test nije dodirivao.

**Rez:** osa identiteta u `ZbirnaIdentResolve` je `ZbirnaID`, ne generacija. `activeLogicalCount` i
`historicalLogicalCount` broje različite **ID-eve**; polja DTO-a su preimenovana
(`selectedZbirnaID`, `historicalOnlyZbirnaID`), a sa njima i pristupnici
(`ZbirnaIDZaBroj`, `ZbirnaJedanIDIkadZaBroj`). Prazan **ZbirnaID** je i dalje integritetska greška —
ali takvu grešku pisac ne može da napravi (`NewEntityID`), pa je to kapija nad zatečeno pokvarenim
redom. Isto meri `Chk_B9`.

**A20 je okrenut, ne obrisan.** Tvrdnja je bila „red bez generacije je greška"; sada je „red bez
generacije se rešava normalno, po ZbirnaID-u". Injekcija je ista, očekivanje suprotno.

**Poruka koja je lagala o broju.** `DUPLI` je javljao „0 otpremnica vraćeno" i za zbirnu koja je imala
izvore: brojao je samo staru vezu po `BrojZbirne`, koju kanonska otpremnica ne nosi. Članstvo se sada
broji **pre** storna (posle njega ga `AktivnoClanstvoPoKanonu` više ne vidi — i to je baš ono što ga
oslobađa), pa se kanonski članovi i stara veza **sabiraju**.

#### Šta sam pokušao i povukao

Obrisao sam i **scoping dece po generaciji** u `StornoZbirnaIDetach_TX` i `PonistiZbirnaChain_TX` —
logika „smem uže, jer sva deca nose generaciju roditelja". Za kanonske podatke je inertna (generacije
nema), pa je delovala kao mrtav kod. **Nije:** storno suite je pao **9 provera**, jer njegovi seed-ovi
generacije pišu, i ti testovi mere baš tu relaksaciju.

Vraćeno. Relaksacija košta ništa dok generacija nema, a njeno uklanjanje je posao koji ide **zajedno
sa fixture-om** — dakle uz S5 (PWA uvoz prelazi na kanon) i S6 (deca prestaju da vise o broju). Upisano
u backlog. Mrtvo je obrisano samo ono što je **dokazano** mrtvo: `ZbirnaGeneracijaPripadaBroju`, nula
pozivalaca.

#### Kapije

Nova sabotaža `zbirna-ident-opet-po-generaciji` vraća osu identiteta na generaciju i obara A20;
`kanonska-zbirna-ne-sme-da-se-razveze` čini svaki broj dvosmislenim i obara nov
`Test_ZBR_KanonskaSmeDaSeRazveze`. Stara `zbirna-ident-greska-kao-none` je **zamenjena, ne
preimenovana**: posle reza njena grana se iz A20 slučaja više ne dostiže, pa bi bila placebo.

### 14.36) S4-3c — osa se dovršava: trag na detetu nosi identitet (23.09.2026)

**Pitanje operatera je zatvorilo dilemu:** *„pošto nema legacy korisnika ni podataka sa genID, da li
čuvamo nešto što nikad neće biti u upotrebi?"* Lanac je bio potpuno cirkularan — kod se čuvao jer ga
testovi mere, a testovi su ga merili jer je kod postojao. Jedini preostali proizvođač generacije zbirne
bio je **test-pečat** (`PecatiGeneracijuAkoZbirna`), napravljen da zadovolji pravilo koje je S4-3b
obrisao.

**Polustanje je već proizvelo kvar.** Otkad `ZbirnaIDZaBroj` vraća identitet, pisci u trag na detetu
upisuju **ID** — a kaskada poništenja ga je čitala kao **generaciju** i prevodila kroz
`IdoviGeneracije`. Prevod bi tražio ID među generacijama, nikad ga ne našao i **tiše** vratio prazno:
identitet se gubi, kaskada pada nazad na broj.

#### Šta je rez uradio

| Korak | Rez |
|---|---|
| **0 · vidljivost** | `RunStornoTestSuite` nije imala `result_file`, pa su njeni padovi ostajali **bez imena**. Zato sam u S4-3b brisanje povukao umesto da ga razumem. Sada piše `last_run_storno.txt`, isti format koji BFP već koristi |
| **1 · čitaoci** | kaskada koristi trag direktno; `ZbrIdIzGeneracije`, `ZbrIdIzGeneracijeAko`, `GenZaZbirnu` obrisani |
| **2 · ime kolone** | `ZbirnaGeneracijaID` → **`ZbirnaRoditeljID`** (kanon, 4 tabele) |
| **3 · scoping i pečat** | scoping **prehranjen**, ne obrisan; oba test-pečata obrisana |
| **4 · backfill** | `BackfillDeteZbirnaGeneracija` obrisan — nema šta da migrira |

#### Dve greške koje su testovi uhvatili, a ja ne bih

**Scoping nije bio mrtav — bio je pogrešno hranjen.** U S4-3b sam ga uklonio kao mrtav kod. Ali i
`SvaAktivnaDecaNoseGeneraciju` i `SuziDecuNaGeneraciju` rade nad **tragom na detetu**, ne nad
`tblZbirna.GeneracijaID`. Trebalo je promeniti samo **šta je scope**: umesto generacije zaglavlja,
njegov `ZbirnaID`. Sposobnost „storniraj jedan od dva dokumenta pod istim brojem" tako ostaje živa,
na kanonskoj osi.

**Ime kolone zamalo da počne da laže o nečem gorem.** Prvo sam je nazvao prosto `ZbirnaID`, po FK
konvenciji. `Test_PR3_OtpremnicaNemaZbirnaID` je pao, i s pravom: *„otpremnica NEMA kolonu koja pokazuje
na zbirnu — pripadnost je u `tblZbirnaIzvori`"*. Članstvo ima **tačno jedan kanal**; kolona imenovana
`ZbirnaID` na detetu predstavljala bi se kao drugi — baš backlink koji je refaktor ukinuo. Ime bi
prestalo da laže o generaciji i počelo da laže o članstvu. `ZbirnaRoditeljID` kaže šta jeste:
denormalizovan pokazivač na roditelja, za decu koja još vise o `BrojZbirne` — i umire sa njima u S6.

> Treća, moja: članove sam prvo brojao kroz `IzvoriZbirne`, koja **diže grešku** nad zbirnom bez
> izvora. Rušila je celu transakciju **pre** storna i obarala 9 storno provera koje sa članstvom nemaju
> veze. Sada `ZbrClanovi`, koji prazno članstvo tretira kao legitimno.

#### Kapije

Sabotaže **572 → 574**: `trag-deteta-opet-generacija` (vraća poslednjeg pisca na generaciju) i
`scoping-dece-bez-identiteta` (izbor prestaje da bude scoped) — obe obaraju imenovanu tvrdnju.
`RunAllTests` **200/0** · BFP zeleno · Storno **163/0**. Otisak šeme `6BDD0C1D` → `64BD33C7`.

### 14.37) S5 — pre-flight slajsa i S5-1 „auto-otpremnica iz PWA otkupa“ (23.09.2026)

**Pravilo koje važi za ceo S5** (operater, 23.09.2026): *cilj je idealna VBA arhitektura; PWA se
prilagođava njoj kasnije, nikako obrnuto.* Ne prave se adapteri, prevodioci ni kolone koje postoje samo
da bi zatečen PWA payload nastavio da radi. Gde kanon i zatečeni PWA ugovor stoje u sukobu, VBA dobija
kanonski oblik, a **šta PWA mora da šalje** se zapisuje kao nizvodni zahtev — i, ako neka PWA sposobnost
zbog toga privremeno ne radi, to se kaže glasno (pauza sa imenom), ne krpi.

#### Merenje je isprav​ilo plan na dva mesta

| Plan je tvrdio | Mereno na `main` `1ef05b36` |
|---|---|
| PWA otkup ingest treba prebaciti na zaglavlje + stavke (§15 backlog) | **Već urađeno u S1c**: `ImportRowToTblOtkup_RowTX` ide kroz `CreateOtkup_TX`, `IsDuplicateInMaster` čita `tblOtkup.ClientRecordID`, `Test_PWA_IngestPraviHeaderIStavku` stoji. Backlog stavka je zastarela |
| `Otkup.VozacID` je živa veza koju treba migrirati | Pišu je **tačno dva mesta**, oba u `modMasterSync`, **oba pod pauzom**. Kolona je mrtva |
| `PROSLEDJENO` kao izvor otpremnice je „obavezno pre S5“ | `IZDATO_PROSLEDJENO` **ne piše nijedan put**. Nije živ kvar nego dva odgovora na isto pitanje — zatvoreno jednim redom |

#### Rez na četiri sesije

| # | Sadržaj | Stanje |
|---|---|---|
| **S5-1** | malina auto-otpremnica nad kanonom (koraci 2b + 3 ciklusa) | ovaj rez |
| **S5-2** | VOZ/zbirna uvoz nad `CreateZbirnaIzIzvora_TX`; `LinkZbirnaToOtkupAndOtpremnica` nestaje | ⏳ |
| **S5-3** | dodela vozača iz PWA (E-019) na otpremnicu; `IzvedeniLanacIzPwaDostupan` i „DEGRADIRANO“ grana obrisani; `Otkup.VozacID/OtpremnicaID/BrojOtpremnice` iz kanona; most preko starog backlinka u `ActiveOtpIDsByZbirna` umire | ⏳ |
| **S5-4** | GAS/PWA strana (E-044, E-058) — vozaču se servira **otpremnica po `Otpremnica.VozacID`**, a ne OTK red po `Otkup.VozacID` | ⏳ |

#### Šta je S5-1 uradio

**Klasa više nije ključ grupisanja — i to je ceo poent reza.** Zatečena auto-otpremnica je grupisala
`StanicaID|Datum|VozacID|Klasa`, jer je klasa bila polje zaglavlja otkupa. Dva bloka istog dana sa istog
otkupnog mesta davala su **dva dokumenta**. Klasa je sada stavka, pa isti ulaz daje **jedan dokument sa
dve stavke**. Ključ grupe je ono što zaglavlje otpremnice nosi i što pisac traži da bude isto za sve
izvore: `StanicaID`, `Datum`, `KulturaID`, `TipAmbalaze`.

Vozač nije u ključu — u malini je **posledica** stanice (ogledalo), ne nezavisan podatak. Zato je korak
2b („`VozacID := StanicaID` na `tblOtkup`“) **obrisan**, a ne preveden: pečat je pripremao grupisanje po
koloni koja umire.

`TipAmbalaze` **jeste** u ključu iako ga pisac poredi samo za izvore koji stvarno nose gajbe. Posledica je
namerno stroža od minimuma: otkup sa deklarisanim tipom a bez ijedne gajbe dobija svoju grupu. To
proizvodi eventualno jedan dokument više — nikad dokument sa pogrešnim tipom, i nikad upis koji pisac
odbije.

| Korak | Rez |
|---|---|
| **1 · jedno telo za vozača-ogledala** | `modMalina.VozacOgledaloZaStanicu` — pravilo koje su hladnjački lanac i malina auto-otpremnica nosili u kopiji. Pravilo nije „pozovi `Ensure`“ nego „odlučuje **ponovljena provera** para“: `Ensure` re-raise-uje, a njegov Boolean kaže samo da li je baš on upisao red |
| **2 · jezgro i batch** | `AutoOtpremnicaDostupna()` (sopstvena kapija, po uzoru na `AutoZbirnaDostupna` iz S4-4) + `AutoCreateOtpremniceFromPWA_TX(samoOtkupID, outGreske)` |
| **3 · delimičan uspeh nije pad** | grupa je svoja transakcija, pa što je prošlo — prošlo je. Razlozi se **imenuju** (`outGreske`), jer „0 kreirano“ bez razloga operateru ne kaže šta da popravi |
| **4 · jedno pravilo „šta je izdato“** | `OtpRequireIzvorValjan` zove `IzdatoStatusJeIzdato` umesto svoje kopije; `PROSLEDJENO` je izdato i za pisca, ne samo za čitače |

#### Dva kvara koja sam sam napravio, a našli testovi

**1 · Normalizacija je iscurila u upis.** Zaglavlje sam gradio **iz ključa grupe**, a ključ je normalizovan
na velika slova jer služi poređenju. Otpremnica je dobijala `TEST GAJBA` umesto `Test Gajba`. Nijedna
kapija to ne vidi — `RequireIstoPolje` poredi `vbTextCompare` — pa bi razlika izašla tek na štampi i u
izveštajima ambalaže, kao tip koji nigde drugde ne postoji. **Normalizacija služi poređenju, nikad
upisu.**

**2 · Pad grupe i pad prolaza nisu ista stvar.** Batch je **svaki** izuzetak pretvarao u `outGreske`, pa
bi ciklus za sistemski pad javio „deo otkupa je ostao bez otpremnice" — a nijedna grupa ne bi ni bila
pokušana. Sastavljanje grupa čita članstvo **strogim** čitačem (`NevezaniOtkupi`), koji nad pokvarenim
zapisom diže grešku; to je pad **koraka**, ne ishod grupe. Uz to je orkestratorova grana `errNum <> 0`
bila **mrtva** — funkcija grešku nikad nije puštala do nje. Sada re-raise-uje, pa ciklus staje pre
outbound sync-a, kao i kod VOZ koraka. Pad jedne **grupe** i dalje ne obara ostale.

#### Kapije

Sabotaže **576 → 581**: `auto-otpremnica-blok-po-blok` (svaki blok svoja grupa), `auto-otpremnica-bez-tipa-u-kljucu`,
`auto-otpremnica-normalizacija-u-upis` (vraća baš gornji kvar 1), `auto-otpremnica-guta-kvar` (grupa bez
otpremnice prođe u tišini), `izvor-otpremnice-opet-samo-izdato`. **Dokazano u oba smera: 5/5 crvenih,
izvor vraćen bit-identično.**

BFP **1837 → 1870** (+33) · `RunAllTests` **200/0** · Storno **163/0** · Banka **241/0** · Palete **97** ·
Agrohemija **25**.

**Nalaz o kapijama:** `vba_check` proverava da je tvrdnja **podniz** literala, a `dokaz.py` za BFP traži
**tačan i statički** tekst — pa je pet unosa prošlo za 5 sekundi, a pun dokaz ih je posle ~20 minuta
prijavio kao `NE OBARA SVOJ TEST`, iako su svi bili crveni i svi na pravoj tvrdnji. U backlogu §15.

#### Ostaje otvoreno posle S5-1

- **Van malina režima auto-otpremnice još nema** — nema izvora vozača dok E-019 ne pređe na otpremnicu
  (S5-3). Korak se prijavljuje kao pauza sa razlogom, ne kao uspeh sa nulom.
- **Grana „nema vozača-ogledala“ je u malina režimu praktično nedostižna** (`Ensure` napravi ogledalo za
  svaku postojeću stanicu, a stanica otkupa je FK-provere​na). Merena je **fault injection-om** — otkupu se
  prepisuje `StanicaID` na nepostojeću stanicu — što je i realan povod (stanica uklonjena iz matičnih
  podataka dok njeni otkupi još žive). Ista grana u hladnjačkom lancu ima svoj test.


## 15) Backlog — namerno van opsega

| Stavka | Zašto ne sada |
|---|---|
| **Hladnjačka otpremnica može upasti u malina batch pre S6** (review #383, P3) | batch uzima **sve** slobodne izdate otpremnice kad je malina mod uključen, a hladnjački auto-lanac takođe pravi odmah izdatu otpremnicu. Ako bi se `AUTO_PRIJEMNICA_HLADNJACA` uključio **pre** S6, batch bi toj otpremnici naknadno napravio zbirnu **mimo** `AutoLanacHladnjaca` — a ovaj rez HLD ZBR korak namerno odlaže zbog odluke atomic-vs-resumable. Danas je lanac OFF do S6, pa nije živ put. **Izlazni uslov S6:** granica se zatvara tako što hladnjački lanac zove **isto jezgro** (`AutoZbirnaZaOtpremnicu`), pa idempotencija rešava preklapanje — ili tako što batch isključi otpremnice hladnjačkih stanica |
| ~~Scoping dece po generaciji još stoji u storno okviru~~ (**zatvoreno u S4-3c** — prehranjen na `ZbirnaID`, ne obrisan) | `StornoZbirnaIDetach_TX` i `PonistiZbirnaChain_TX` računaju „smem uže, jer sva aktivna deca nose generaciju roditelja". Za kanonske podatke je **inertno** — generaciju ne piše nijedan živ pisac — ali nije mrtvo: storno fixture je piše u seed-ovima, pa brisanje obori **9 provera** koje tu relaksaciju mere. Uklanja se **zajedno sa fixture-om**: uz S5 (PWA uvoz prelazi na kanon, prestaje jedini pisac generacije) i S6 (deca prestaju da vise o `BrojZbirne`). Do tada košta ništa |
| **PWA / MasterSync ingest** | radi se isključivo VBA. Nalaz koji čeka: PWA šalje **jedan record = jedna klasa = ceo dokument**, sa svežim `brojDokumenta` po svakom snimanju (`src/js/features/otkup/otkup-form.js:643`, `:493`). Ingest postaje 1 record → 1 header + 1 stavka; **nema heurističkog grupisanja i ne treba eksterni Document UID**. `ClientRecordID` ide na header, a `IsDuplicateInMaster` (`modMasterSync.bas:1824`) mora da se prepokaže na header tabelu — inače se svaki PWA dokument reimportuje. |
| **Self-update** | van opsega po dogovoru |
| **App / Repo / Qry slojevi** | **Ne paralelno sa refaktorom** — pokvarilo bi kapiju odluke iz §14.1: dve promenljive odjednom znače da se ne može reći da li je čist ishod zasluga šeme ili slojeva. Uz to, App sloj već postoji neimenovan (`mod*Unos` prima DTO rečnik, `NoviOtpremnicaUnos`), a enforcement daje A11 allowlist, ne ime modula. Jedini sloj koji stvarno nedostaje je **Qry** (`modDokumenta`: 15 javnih čitača pored 21 mesta upisa) — ali dobar deo tih čitača postoji da rekonstruiše dokument po broju i **umire u PR 12**. Revidirati **posle PR 12**, kad se zna koji čitači preživljavaju. Do tada: čitanja u novim writer-ima idu iza imenovanih funkcija, ne inline skenova. |
| **Delimična alokacija Otkup → Otpremnica** | nema poslovnog zahteva; ako se pojavi — eksplicitna alokaciona tabela |
| **Mešovit dokument (više vrsta u jednom)** | vrsta/sorta ostaju header; nije zahtev |
| **Otkup u statusu `PROSLEDJENO` kao izvor otpremnice** (review #363, P2) | danas nije živ put (`CreateOtkup_TX` piše `IZDATO`); **obavezno pre S5**: `OtpRequireIzvorValjan` priznaje `IZDATO` i `PROSLEDJENO` (semantika `IzdatoStatusJeIzdato`, ne `DocIsIssued`) |
| **Granica agregata za komande nad jednim dokumentom** (review #363, drugi krug, P2) | strogi čitači (`StavkeOtpremniceRedovi`, `StavkeOtkupaRedovi`) validiraju ceo skup, pa komanda jednog dokumenta (`IzdajOtpremnicu_TX`, `GetOtpremnicaProgress`) pada i zbog nepovezanog pokvarenog dokumenta. Fail-closed, nije kvar podataka. Kandidat: strogi čitač po dokumentu (ciljna otpremnica + njeni izvori) za komande, a skup za izveštaje i mreže |
| **Strog čitač ZAGLAVLJA otkupa za čitaoce koji ga sastavljaju** (review #364, P2) | specifikacija blokova čita `BrojDokumenta`, `Datum`, `KooperantID`, `StanicaID`, `VrstaVoca`, `SortaVoca` direktno iz `tblOtkup`, a `CreateOtkup_TX` za te činjenice drži jače invarijante (broj i datum obavezni, kooperant i stanica postoje, kultura usklađena). Naknadno pokvareno zaglavlje zato daje prazan broj bloka na papiru umesto pada. Pravila pisca se **ne prepisuju** u `modPrint`: u sledećem prolazu proveriti postoji li kanonski strog čitač zaglavlja, pa ga koristiti — isti rez kao `StavkeOtkupaRedovi` za stavke |
| **`NEDOVRSENO` nudi radnju po LISTI, a ne po REDU** (review #366, P2) | `Scr_Radnje` za tu listu vraća jedno `danger` dugme (`odbaci`), pa ga operater dobija i nad redom koji ga ne prima — `IZGUBLJEN_BLOK`, osirotela prijemnica, red sa greškom. Mutacije nema: `OdbaciIspravku` odbija red bez `CorrectionID`-a i još to i zabeleži. Problem je što UI **nudi** radnju za koju unapred zna da nije primenljiva, i što je jedini put do prave radnje rečenica u koloni „akcija“. Read-model to već zna — `GetNedovrseno` nosi `actionCode` (`CONTEXT` / `PRIJ` / `PAL` / `BLOK`) — ali ga `RowsNedovrseno` **ne prenosi u mrežu**, pa ljuska nema čime da bira. Od S3d-1 je isti nedostatak vidljiv i u listi „Bez otpremnice“: red nudi i **„Veži“** i **„Ponovi auto-lanac“**, a svaka radnja tek na svojoj granici odbije blok koji joj ne pripada (podaci su bezbedni, UX nije). Rez: nevidljiva kolona sa `actionCode`-om + radnje po redu (`CONTEXT` → Odbaci ispravku, `BLOK` → otvori F1/Bez otpremnice, `PRIJ`/`PAL` → Preveži, `GRESKA` → bez mutacione radnje). To je **ugovor ljuske**, ne samo ovaj ekran: `trebaRed` danas zna samo „treba red / ne treba / označeni“, a ovde treba „zavisi od vrste reda“. Zato ide kao svoj rez, ne uz S3c-2 |
| **Rollback `tblAmbalaza` u ispravci nije dokazan testom** (review #365, P2) | `IspravkaOtpremnice_TX` snimi sve četiri tabele (`tblOtpremnica`, `…Stavke`, `…Izvori`, `tblAmbalaza`), a storno stare vraća gajbe koje je njeno izdavanje knjižilo. Test atomarnosti (`Test_OTP_IspravkaIzdate`) meri da stara ostaje AKTIVNA kad ispravka padne, i sabotaža `ispravka-pad-ostavlja-storniranu` to obara — ali **nijedna tvrdnja ne meri stanje ambalaže posle rollback-a**. Implementacija izgleda ispravno; nedokazano je nedokazano. Rez: tvrdnja nad zbirom gajbi pre i posle pale ispravke + sabotaža koja skida `AddTableSnapshot TBL_AMBALAZA` (danas bi prošla neprimećeno) |
| **`AUTO_PRIJEMNICA_HLADNJACA` ne sme da preživi S6 kao poslovna opcija** (review #367) | dok traje refaktor prekidač je legitimna **tehnička** kapija: „lanac sme da se pusti“. Ali za hladnjaču je automatika **obavezna**, pa kombinacija `JeHladnjača = DA` + `AUTO = NE` posle S6 opisuje stanje koje specifikacija zabranjuje — a korisnik bi ga podesio u dva klika. **Izlazni uslov S6:** obrisati podešavanje, ili ga pretvoriti u interni/deployment prekidač van normalnog toka (i tako ga opisati u Podešavanjima). Ne ostavljati dva autoriteta nad istim pravilom |
| **Strog čitalac je registarski, ne dokumentarni** (review #370, P2) | `StavkeZbirneRedovi` (kao i čitači otkupa i otpremnice) validira **celu tabelu**, pa jedna pokvarena istorijska zbirna obori čitanje svih ostalih — i liste u kojima te zbirne nema. Fail-closed je ovde namerno izabran i ostaje, ali se vredi razdvojiti: **registarski audit → globalno strogo**, **jedan dokument → strogo u opsegu tog dokumenta**. Rez važi za sva tri tipa odjednom, pa ne ide unutar jednog slajsa |
| **`who_writes` ne prijavljuje MRTAV UNOS u `WRITE_OWNERSHIP.json`** (nalaz S3e-1) | spisak dozvola je nabrajao tri modula koja tabelu odavno ne pišu, a kapija je ćutala: `--check-ownership` proverava samo da je **svaki pisac naveden**, ne i da **svaki naveden piše**. `vba_hard_census` isti problem rešava pravilom `MRTAV_UNOS`. Rez: isto pravilo u `who_writes.py`, pa spisak ne može da istruli neprimećeno |
| **`NEVEZANE` nije sužena na AKTIVNI nacrt** (review #379, P3) | čitalac filtrira po stanju dokumenta (izdata, nestornirana, slobodna), ali ne po odnosu prema izabranoj zbirnoj — otpremnica drugog vozača ili druge vrste/sorte/tipa ambalaže ostaje u ponudi, a `ZbrRequireIzvorValjan` je odbija. Nema kvara podataka (pisac je fail-closed), ali je to isti obrazac koji smo već jednom zatvorili za nacrte. Rez: `NevezaneOtpremnice(zbirnaID)` koja sužava po vozaču i preuzetim činjenicama kad nacrt postoji — i test sa **nekompatibilnom** otpremnicom, jer današnji test meri samo ime liste |
| **Aktivan nacrt (`mZbrID`) preživljava izlazak iz F2** (review #377, P3) | radni sto ostaje izabran i posle promene režima, pa se operater može vratiti u F2 i ne primetiti da je kontekst još tu. Nije integritetski problem — kontekst je vidljiv kroz aktivnu listu i naslov mreže, a pisac i dalje drži sve kapije; isti obrazac postoji i kod otpremnice (`mOtpID`). Pripada **usability sweep-u** nad radnim stolovima, ne kanonskom cutover-u — i tada se rešava za **oba** stola odjednom, ne samo za zbirnu |
| **Lista `SVI` u F2 nudi `Veži` i nad NACRTOM otpremnice** (review #377, P3) | pisac je bezbedno odbija (`RequireOtpValidanIzvorZbirne`), pa nema kvara podataka — ali je to isto ono što `NevezaneOtpremnice` namerno izbegava: nuditi operateru nešto što će pisac odbiti. `SVI` je namerno sveobuhvatna lista, pa se rešava uz sledeći rez (traka napretka + čišćenje polja F3) |
| **`vba_check` pusta PODNIZ tamo gde `dokaz.py` trazi TACAN tekst** (nalaz 23.09.2026) | katalog sabotaza za BFP mora da nosi **doslovan** tekst tvrdnje, jer ta suite ispisuje naziv tvrdnje umesto imena Sub-a — tvrdnja je jedina adresa. `vba_check` proverava samo da je tvrdnja **podniz** nekog literala u imenovanom testu, pa je pet novih unosa proslo za 5 sekundi, a pun dokaz ih je posle ~20 minuta prijavio kao `NE OBARA SVOJ TEST` — iako je svih pet bilo crveno i svih pet na pravoj tvrdnji. Jeftina kapija pusta ono sto skupa odbija, pa povratna informacija stize dvadeset minuta kasnije. Rez: za suite sa `result_file`-om `vba_check` da trazi **tacan i staticki** tekst (tvrdnja sa `&` u sebi nije adresa). Ide uz PR nad `tools/` zajedno sa pravilom vidljivosti, ne uz feature |
| **`vba_check` ne vidi VIDLJIVOST pozvanog imena** (nalaz 22.09.2026) | treći compile-pad u jednoj sesiji koji statička kapija propusti: #371 preimenovan parametar, #374 obrisane javne funkcije koje se još zovu, #376 poziv **`Private` procedure iz drugog modula** (`GetValueByKey` je privatan u `modBusinessFlowProTests`). Svaki put ishod nije pad nego **Excel koji visi do timeout-a** (`run-vba visi = compile greska`), pa je dijagnoza skupa. Rez: pravilo koje za svako `Ime(` proveri da je ime u istom modulu ili `Public` negde; filtriranje lokalnih deklaracija i komentara je obavezno, inache je šum neupotrebljiv (mereno: 20 lažnih pogodaka bez filtera). Ide kao svoj mali PR nad `tools/`, ne uz feature |
| **`modOtkup.VrednostOtkupa` ne drži ceo ugovor stavki** (review #363, drugi krug) | čitač vrednosti JEDNOG otkupa (banka, novac) proverava samo kg i cenu > 0, ne klasu, jedinstvenost klase ni gajbe. Otpremnica ga ne koristi. Uskladiti sa `StavkeOtkupaRedovi` kad se dira novac |

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
