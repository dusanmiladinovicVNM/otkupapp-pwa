# Domain Model + Schema v1 — lanac dokumenata

> Ciljni model posle refaktora „header + stavke". Ugovor koji ga uokviruje:
> `ARCHITECTURE_CONTRACT.md`. Plan isporuke: `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md`.
>
> Status: **model usvojen; implementacija u toku.** Ovo je i dalje specifikacija,
> ne opis koda — osim tamo gde red kaže drugačije.
>
> | Dokument | Stanje |
> |---|---|
> | Zbirna | **tabele i pisač postoje** (PR3, aditivno): `tblZbirnaStavke`, `tblZbirnaIzvori`, `CreateZbirna_TX` / `CreateZbirnaIzIzvora_TX`. Nov pisač još nema pozivaoca — stari (`SaveZbirnaMulti_TX`) je jedini put; čitaoci, invarijanta i storno idu u Zbirna cutover |
> | Otpremnica / Otkup / Prijemnica | specifikacija |
>
> Kontekst: nema legacy transakcionih podataka. Zatečena šema **nema pravo veta**
> nad ovim modelom. Gde postojeći kod ne podržava model — kod se adaptira ili
> briše, ne model.

---

## 1) Kako je model izveden

Tri izvora, po prioritetu:

1. **Postojeći potpisi `Save*Multi_TX`** — podela header/stavka. Polje koje u
   potpisu stoji jednom je dokument-level; polje koje stoji kao par I/II je
   line-level. Potpis je već napisana specifikacija.
2. **Postojeće ponašanje** — tamo gde potpis ne odgovara (npr. `Fakturisano`).
3. **Eksplicitna odluka** — samo gde prva dva ne daju odgovor. Svaka takva odluka
   je dole označena kao **ODLUKA** sa obrazloženjem.

Ništa nije pretpostavljeno bez oznake.

---

## 2) Identitet — dva sloja

### 2.1 Tehnički PK: opaque, immutable

Nove transakcione tabele dobijaju neprozirne ID-eve:

```
OtkupID = "OTK-" & 32 hex karaktera
```

Razlog: desktop, PWA i budući sistemi moraju da mogu da naprave identitet **bez
koordinacije**. Današnji `GetNextID` (`modDataAccess.bas:392`) skenira celu tabelu
i uzima `max + 1` — O(n) po upisu i strukturno nesiguran kod više uređaja.

Jedna fabrika za ceo sistem:

```vba
Public Function NewEntityID(ByVal prefix As String) As String
```

Nijedan modul ne pravi svoju GUID logiku. Implementacija: `CoCreateGuid` kroz
`Declare` u **deklaracionoj sekciji** modula (CLAUDE.md §3). Bez crtica i vitičastih
zagrada — 32 hex znaka, da string ostane kratak u Variant nizovima.

### 2.2 Matični podaci ostaju čitljivi

`KOOP-123`, `ST-1`, `VOZ-7`, `KUP-4` **ne menjaju format**. PWA/GAS ih već tretiraju
kao stabilne identitete i nema razloga da se diraju.

Sistem dakle svesno ima **dva formata ID-a**:

| Sloj | Format | Ko pravi |
|---|---|---|
| Matični podaci | `PREFIX-<broj>` | `GetNextID`, nepromenjeno |
| Transakcioni dokumenti i stavke | `PREFIX-<32 hex>` | `NewEntityID` |

Ovo je namerno i zapisano ovde da ga niko kasnije ne „harmonizuje".

### 2.3 Poslovni broj

`BrojDokumenta` / `BrojOtpremnice` / `BrojZbirne` / `BrojPrijemnice` ostaju u
današnjem formatu (`12/090926`) i ostaju **labele** (A2).

---

## 3) Kardinaliteti — izvedeni iz koda

| Veza | Kardinalitet | Izvor | Nosilac FK |
|---|---|---|---|
| Otkup → OtkupStavke | 1:N | model | `OtkupStavka.OtkupID` |
| Otkup → Otpremnica | **N:1, promenljiva** | `docs/DOMEN/README.md` §1 („više blokova → jedna otpremnica"); `ReassignOtkupToOtpremnica_TX` (`modDokumenta.bas:4266`) dokazuje da se blok može premestiti | **`tblOtpremnicaIzvori`** |
| Otpremnica → Zbirna | **N:1, promenljiva** | jedna `BrojZbirne` kolona danas; `RelinkOtpremniceToZbirna_TX` | **`tblZbirnaIzvori`** |
| Zbirna → Prijemnica | **1:N** (namera 1:1, ali **nije tvrdo**) | v. §3.1 | `Prijemnica.ZbirnaID` |
| PrijemnicaStavka → FakturaStavka | **1:1, puna količina** | `CreateFaktura` uzima celu `Kolicina` reda; `IsPrijemnicaAvailableForFaktura` sprečava drugo fakturisanje (`modFaktura.bas:232, 896`) | `FakturaStavka.PrijemnicaStavkaID` |
| PrijemnicaStavka → PaletaStavka | 1:N | `tblPaletaStavka` već nosi `PrijemnicaID` **i** `Klasa`/`VrstaVoca`/`SortaVoca` → grain je klasa, ne dokument | `PaletaStavka.PrijemnicaStavkaID` |

### 3.1) Zbirna → Prijemnica: pravilo je meko, i to je nalaz

Ovo je kardinalitet koji je bio označen kao otvoren. Kod ima odgovor, i odgovor je
nijansiran.

`modDokUnos.PrijemnicaValidiraj` nosi komentar *„1 zbirna = 1 prijemnica"* i
poruku `DOKUNOS_ASK_DUPLA_PRIJ_3`: *„Pravilo je 1 zbirna = 1 prijemnica, pa ovo liči
na DUPLI UNOS. Ipak snimiti?"* — ali to je **`MsgBox` sa Yes/No, ne tvrda kapija**.
Operater sme da nastavi.

Dakle:

- **Namera:** 1:1
- **Sprovedeno:** 1:N sa upozorenjem
- **Model:** 1:N (`Prijemnica.ZbirnaID`), jer FK u tom smeru prirodno nosi oba

Ne pretvarati ovo u tvrdo 1:1 u ovom refaktoru — to bi bila izmena poslovnog
ponašanja koja nije potrebna za identitet. Upozorenje ostaje, samo se pita po
`ZbirnaID` umesto po `BrojZbirne`.

### 3.2) Članstvo i alokacija su dva različita pojma

Lako se pomešaju jer oba „vezuju otkup za otpremnicu". Nisu isto:

| | Pitanje na koje odgovara | Status |
|---|---|---|
| **`tblOtpremnicaIzvori`**<br>`(OtpremnicaID, OtkupID)` | *Koji otkupi čine ovu **verziju** otpremnice?* | **potrebno sada** (A15) |
| **`tblOtpremnicaAlokacije`**<br>`(OtpremnicaStavkaID, OtkupStavkaID, Kg)` | *Koliko je kilograma iz ovog otkupa otišlo na ovu otpremnicu?* | **samo ako se pojavi poslovni zahtev** |

**Članstvo je celobrojno i obavezno:** otkup pripada otpremnici ili ne pripada.
Ono postoji zato što ispravka pravi novu verziju, a nepromenjene sestre moraju
ostati vidljive u sastavu **obe** — to nema veze sa deljenjem količina.

**Delimična alokacija** — da jedna `OtkupStavka` delimično završi u više
otpremnica — danas nema poslovni zahtev i **ne modelujemo je**. Ako se pojavi,
uvodi se zasebna tabela sa `Kg`; ne rasplinjava se članstvo „za svaki slučaj",
i ne dodaje se `Kg` u tabelu članstva.

> Ranija verzija ovog odeljka je koristila ime `tblOtpremnicaIzvori` za
> **alokacionu** tabelu i time ga sudarila sa članstvom. Alokacija se od sada
> zove `tblOtpremnicaAlokacije`.

---

## 4) Tabele — grain, kolone, izvor istine

Legenda: **H** = header, **S** = stavka. `→` = FK.

### 4.1 `tblOtkup` — **grain: jedan otkupni blok**

Jedan otkup od jednog kooperanta, na jednom otkupnom mestu, jednog dana
(`README.md` §1).

| Kolona | Napomena |
|---|---|
| `OtkupID` | PK, `OTK-<hex>` |
| `BrojDokumenta` | labela, scoped po otkupnom mestu |
| `Datum`, `KooperantID`, `StanicaID`, `ParcelaID`, `KulturaID` | → matični |
| `VrstaVoca`, `SortaVoca`, `TipAmbalaze` | H — u potpisu stoje jednom |
| `KolAmbIzdata` | H — OM izdao prazne kooperantu; **stvarna činjenica sa otkupnog lista** |
| `ClientRecordID`, `SyncSource` | eksterni identitet i poreklo (PWA); v. §7 i §4.1c |
| `Stornirano`, `IzdatoStatus` | lifecycle |
| `IspravkaOdID`, `ZamenjenSaID`, `CorrectionID` | correction |
| audit ×4 | |
| ~~`VozacID`~~ | **ne postoji** — vozač pripada Otpremnici; v. §4.1c |
| ~~`Novac`, `PrimalacNovca`~~ | **BRIŠU SE** — keš se ne vezuje za otkupni list; v. §4.1b |
| ~~`Isplaceno`, `DatumIsplate`~~ | **ne postoje** — izvedeno iz `tblNovac`; v. §4.1c i §6.1 |
| ~~`VremeUnosa`~~ | **ne postoji** — `CreatedAt` / `SourceCreatedAt`; v. §4.1c |
| ~~`OtpremnicaID`~~ | **ne postoji** — pripadnost zna `tblOtpremnicaIzvori` |
| ~~`ZbirnaID`~~ | **ne postoji** — pripadnost zna `tblZbirnaIzvori` preko otpremnice |
| ~~`BrojOtpremnice`, `BrojZbirne`~~ | **ne postoje** — broj nije veza (A2) |
| ~~`Klasa`, `Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg`~~ | **stavka**, ne header |

**`tblOtkupStavke`** — grain: **jedna klasa jednog bloka**

| Kolona | Semantika |
|---|---|
| `OtkupStavkaID` | PK `OKS-` |
| `OtkupID` → | FK, obavezan |
| `RedniBroj`, `Klasa` | |
| `Kolicina` | **uvek NETO kg**, zamrznuto pri izdavanju |
| `BrutoKg` | zamrznut originalni bruto — popunjen **samo** kad je unos bio bruto |
| `Cena` | **stvarno primenjena** cena tog dokumenta |
| `KolAmbalaze` | |

**Vlasnik upisa (A11): samo `modOtkup`.** Danas 9 pisaca
(`modAutoHladnjaca`, `modDokumenta`, `modMasterSync`, `modNovac`, `modOtkup`,
`modOtkupBlok`, `modSetup`, `modSledljivost`, `modStornoFlow`) — svi ostali idu
kroz API `modOtkup`-a.

> `modSetup` je u zatečenom stanju pisač zbog backfill-a `BrojOtpremnice`. Bez
> podataka koje treba dopuniti taj kod nema posao i briše se u cutover-u;
> `modSetup` ostaje **`schema_owner`**, ne `row_owner`. `WRITE_OWNERSHIP.json`
> već nosi `cilj: ["modOtkup"]` — dokument je bio taj koji je kasnio.

#### 4.1c Četiri kolone koje odlaze, i šta ih zamenjuje

Brisanje kolone bez imenovanog naslednika je način da se obori ekran koji ju je
čitao. Zato svaka nosi zamenu, izmerenu nad zatečenim kodom.

**`VozacID` — vozač nije činjenica otkupa.**

U trenutku nastanka otkupnog lista vozač često nije ni poznat; desktop ga dobija
iz izabrane otpremnice, a PWA tek naknadno bira koji listovi idu u koju
otpremnicu i kod kog vozača. Jedno otkupno mesto istog dana ima tri otpremnice sa
tri vozača — otpremnica je **transportni agregat**, otkup nije.

```
OTPREMNICA 17  Vozac = VOZ-3        OTPREMNICA 18  Vozac = VOZ-7
   ├── OTK-101                          ├── OTK-103
   ├── OTK-102                          └── OTK-104
   └── OTK-108
```

Zamena: `Otpremnica.VozacID` + `tblOtpremnicaIzvori`. PWA sme da nosi izabranog
vozača kroz svoj tok, ali podatak sleće na otpremnicu, ne kao kopija na svakom
otkupu.

**`Isplaceno` / `DatumIsplate` — read-model, ne kolona.**

```
Vrednost  = SUM(stavke.Kolicina x stavke.Cena)
Placeno   = SUM(tblNovac vezan na OtkupID)
Preostalo = Vrednost - Placeno
Isplaceno = (Preostalo <= 0)
```

Mereni čitaoci danas i njihova zamena:

| Čitalac | Šta radi | Posle |
|---|---|---|
| `modNovac.GetOpenOtkupi:1355` | `If CStr(data(i, colIspl)) = STATUS_ISPLACENO Then GoTo NextCount` | uslov se **računa**; `BuildIsplataDictByOtkup()` već postoji u istom modulu |
| `modProductionHealthCheck.Check_OtkupPaymentConsistency:485` | poredi kolonu sa `tblNovac` | **briše se** — postoji samo da uhvati neslaganje kolone i knjige; bez kolone nema šta da se ne slaže |
| `modNovac:1277,1280,1285,1286` | `RequireUpdateCell` nad `tblOtkup` | **nestaje** — time `modNovac` prestaje da bude pisač otkupa (A11) |

`DatumIsplate` nema nijednog produkcionog čitaoca — samo testove.

**`VremeUnosa` — dva naslednika, ne jedan.**

Čita ga `modPrint:591`, na samom otkupnom listu. Ne može prosto da nestane.

| Tok | Nosilac vremena |
|---|---|
| desktop unos | `CreatedAt` — isto značenje, kolona je bila udvajanje |
| PWA uvoz | `CreatedAt` je vreme **uvoza**; vreme unosa na terenu je druga činjenica → **`SourceCreatedAt`** |

Štampa čita naslednika po toku. Dvosmisleno „vreme unosa" ne ostaje.

---

### 4.1d Bruto/neto i cena — zamrznute činjenice

Ovde se ne uvodi nov mehanizam; zapisuje se onaj koji `OtkupValidiraj` već ima.

```
Kolicina = UVEK NETO kg

unos NETO   ->  BrutoKg = prazno
unos BRUTO  ->  BrutoKg = TACNO ono sto je korisnik uneo
                Kolicina = izracunat neto (bruto - tara)
```

> **Težina ambalaže se koristi samo u trenutku nastanka dokumenta.** `BrutoKg` i
> `Kolicina` izdate verzije su **zamrznute činjenice** i nikad se ne
> rekalkulišu iz `tblTipAmbalaze`.
>
> Ako danas `BrutoKg 1100`, `KolAmbalaze 100`, tara `1 kg` daju `Kolicina 1000`,
> a za dve godine master težina gajbice postane `1.2 kg` — istorijski dokument
> ostaje `1100 / 1000`. Zasebna `TaraKg` kolona nije potrebna: `BrutoKg`,
> `Kolicina` i `KolAmbalaze` već nose dovoljno istorije.

Isto važi za cenu:

| | |
|---|---|
| `Cenovnik.Cena` | **predlog** — autofill u formu |
| `OtkupStavka.Cena` | **stvarno primenjena** cena tog izdatog dokumenta |

Operater sme da je pregazi. Writer nema pravo da traži `Cena = Cenovnik.Cena`;
njegovo pravilo je `Cena > 0`. Promena cenovnika **ne menja** već izdat otkup —
štampa i danas računa iz cene sa samog otkupa, što je ta semantika.

---

### 4.1e Lifecycle i pripadnost

**Otkup nema persistentan `DRAFT`.** Forma jeste njegov draft:

```
operater unosi -> koriguje -> Unos -> OTKUP nastaje kao IZDATO
```

Otpremnica i Zbirna imaju pravi persistentan `DRAFT` jer se njihovo članstvo
gradi postepeno; otkup tu potrebu nema.

Posledica (A13): ispravka nikad ne menja snimljen otkup.

```
OTK-101 / broj 17
      v correction
OTK-202 / broj 18,  IspravkaOdID = OTK-101
```

**Pripadnost otpremnici je isključivo `tblOtpremnicaIzvori`.** Jedan otkup sme
istorijski da pripada i staroj i novoj verziji otpremnice, ali u datom trenutku
najviše **jednoj aktivnoj**. Dva aktivna članstva su integrity failure (A15).

Delimična alokacija (`OTK-17` 40% → `OTP-A`, 60% → `OTP-B`) **nije modelovana** i
ne uvodi se dok ne postoji poslovni zahtev; v. §3.2.

Time se zatvara i poslednje otvoreno pitanje iz §9: *„sme li `Otkup.OtpremnicaID`
da se razlikuje po klasi"*. Ne — kolone nema, a članstvo je na nivou jednog
otkup **headera**.

---

### 4.2 `tblOtpremnica` — **grain: jedna isporuka sa otkupnog mesta**

`OtpremnicaID` (PK `OTP-`), `BrojOtpremnice` (labela, scoped po stanici), `Datum`,
`StanicaID`, `VozacID`, `VrstaVoca`, `SortaVoca`, `TipAmbalaze`, `Cena`
(ne-finansijski predlog, §13b plana), `Stornirano`, trace ×4, audit ×4.

**Bez `ZbirnaID`.** Pripadnost zbirnoj zna `tblZbirnaIzvori`; „na kojoj je
aktivnoj zbirnoj sada" računa `modDokumenta.AktivnaZbirnaZaOtpremnicu`.

**`tblOtpremnicaStavke`** — grain: **jedna klasa jedne otpremnice**
`OtpremnicaStavkaID` (PK `OPS-`), `OtpremnicaID` →, `RedniBroj`, `Klasa`,
`Kolicina`, `KolAmbalaze`, `BrutoKg`.

> **Bez `Cena`.** Otpremnica je izvedeni dokument: njena vrednost je zbir
> izvornih otkupnih stavki, koje mogu imati **različite cene**. Jedna `Cena` na
> stavci bi bila drugi izvor istine za novac. Zatečena `Otpremnica.Cena` nije
> agregat nego **predlog za prefill otkupnih blokova**
> (`modOtkupBlok.bas:688`) — ostaje na headeru kao izričito ne-finansijsko
> polje. Puno obrazloženje: `REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §13b.

**`tblOtpremnicaIzvori`** — grain: **jedan otkup u sastavu jedne verzije otpremnice**
`OtpremnicaIzvorID` (PK `OPI-`), `OtpremnicaID` →, `OtkupID` →, audit ×4.

> Isti obrazac i isto pravilo kao `tblZbirnaIzvori` (A15): promenljivo dok je
> otpremnica `DRAFT`, zamrznuto pri izdavanju. `Otkup.OtpremnicaID` u ciljnom
> modelu **ne postoji** — pripadnost zna isključivo ova tabela.

**Vlasnik upisa:** `modDokumenta` (ili nov `modOtpremnica`). Danas 4 pisca.

---

### 4.3 `tblZbirna` — **grain: jedan transport ka kupcu/hladnjači**

`ZbirnaID` (PK `ZBR-`), `BrojZbirne` (labela, scoped po vozaču), `Datum`,
`VozacID`, `KupacID`, `Hladnjaca`, `Pogon`, `VrstaVoca`, `SortaVoca`,
`TipAmbalaze`, `Stornirano`, trace ×4, audit ×4.

**`tblZbirnaStavke`** — grain: **jedna klasa jedne zbirne**
`ZbirnaStavkaID` (PK `ZBS-`), `ZbirnaID` →, `RedniBroj`, `Klasa`, `Kolicina`,
`KolAmbalaze`.

**`tblZbirnaIzvori`** — grain: **jedna otpremnica u sastavu jedne verzije zbirne**
`ZbirnaIzvorID` (PK `ZBI-`), `ZbirnaID` →, `OtpremnicaID` →, audit ×4.

> **Nepromenljivost počinje pri izdavanju, ne pri upisu** (A15):
>
> | Stanje zbirne | Članstvo |
> |---|---|
> | `DRAFT` | **promenljivo** — izvori se dodaju i sklanjaju slobodno |
> | `IZDATO` / `PROSLEDJENO` | **zamrznuto** — nova verzija dobija svoje redove, stara zadržava svoje |
>
> Zato tabela nema `Stornirano` (storno verzije ne briše njen sastav) i stoji u
> `BEZ_STORNA`.
>
> **`Otpremnica.ZbirnaID` ne postoji.** Pripadnost zna isključivo ova tabela;
> „na kojoj je *aktivnoj* zbirnoj otpremnica sada" računa
> `modDokumenta.AktivnaZbirnaZaOtpremnicu`. Razlog za tabelu umesto kolone:
> sestre koje se nisu menjale pripadaju i staroj i novoj verziji, a **jedna FK
> kolona bi mogla da pokaže samo jednu** (A15).

> **Zbirna nema cenu.** `tblZbirna` je nikad nije imala i `SaveZbirnaMulti_TX` je
> ne prima (`modDokUnos.bas:422`). Ne dodavati je.

**Izvor istine:** zbirna je **agregat** — otpremnice su izvor. Ali „izvedeno"
prestaje da važi kad dokument bude izdat:

| Stanje dokumenta | Šta su stavke |
|---|---|
| `DRAFT` | **keš** (A5) — izvode se iz izvora, invarijanta §6.2 to dokazuje pri svakoj izmeni |
| `IZDATO` / `PROSLEDJENO` | **sadržaj te verzije** — istorijska činjenica, ne prepisuje se (A13) |

Kad se izvor promeni posle izdavanja, ne menja se ovaj dokument nego se pravi
**nova verzija** (A13). Lanac dokumenata danas nema draft fazu, pa je u praksi
svaki dokument odmah izdat — što znači da je drugi red pravilo, a prvi
priprema za trenutak kad UI dobije „otvoren dokument".

**Vlasnik upisa (A11): samo `modDokumenta`.** Danas 3 pisca
(`modDokumentInvariant`, `modDokumenta`, `modMasterSync`).

> Ranija verzija je ovde pisala „`modDokumenta` + `modDokumentInvariant`
> (rekalkulacija)", što je bilo u suprotnosti sa `WRITE_OWNERSHIP.json`, gde je
> `cilj` samo `modDokumenta`. Kontradikcija se razrešava u korist registra:
>
> **`modDokumentInvariant` računa i validira; `modDokumenta` jedini fizički
> piše.** `RecalculateZbirna_TX` živi u `modDokumenta` i prima izračunat
> rezultat. To je čisto A11 vlasništvo — jedan pisač, jedan ulaz — umesto dva
> modula koji pišu istu tabelu po svojim pravilima.

> **Stanje posle PR3.** Sve tri tabele postoje u kanonu i u svesci, a writer
> piše header, stavke **i članstvo** u jednoj transakciji. `ZbirnaID` je opaque
> (`NewEntityID`), `GeneracijaID` se ne piše.
>
> Dva javna ulaza, jedno jezgro:
> `CreateZbirna_TX(h, izvor, ocekivano, outGreska)` za ručni unos (očekivano je
> **obavezno**) i `CreateZbirnaIzIzvora_TX(h, izvor, outGreska)` za automatski
> tok. Kontrola „očekivano vs izvedeno" se time ne može isključiti time što se
> argument ne prosledi.
>
> **Količine se ne primaju — izvode se.** Writer prima *izvorne otpremnice* i
> računa stavke iz njih; `VrstaVoca` / `SortaVoca` / `TipAmbalaze` takođe dolaze
> iz otpremnica, koje moraju biti saglasne. Prva verzija je primala gotove
> stavke, pa je bilo legalno napraviti zbirnu od izmišljenih 400+600 kg **bez
> ijedne otpremnice** — keš koji se ne slaže sa izvorom, i to kroz kanonski
> writer. Opcioni argument `ocekivano` nosi ono što je operater otkucao i služi
> samo kao unakrsna provera protiv izvedenog.
>
> **Članstvo ide u istoj transakciji.** Da writer ne upisuje `tblZbirnaIzvori`,
> cutover bi morao „napravi zbirnu; commit; pa poveži otpremnice" — a pad drugog
> koraka ostavlja zbirnu bez izvora.
>
> Header koji taj pisač napravi **namerno ostavlja `UkupnoKolicina`,
> `UkupnoAmbalaze` i `Klasa` prazne** — to su kolone koje u ovom modelu ne
> postoje; količina živi na stavci. Prazno je tačan odgovor („ne pitaj header za
> količinu"), i test to zaključava da neko u cutover-u ne bi „za svaki slučaj" upisao i
> zbir na header i time napravio dva izvora istine za istu vrednost.
>
> Kolone se brišu u Zbirna cutover-u, zajedno sa `ZbirnaIdent*` / `ZbirnaGeneracija*`.

---

### 4.4 `tblPrijemnica` — **grain: jedan prijem robe na odredištu**

`PrijemnicaID` (PK `PRJ-`), `BrojPrijemnice` (labela, scoped **po kupcu**),
`Datum`, `KupacID`, `VozacID`, `VrstaVoca`, `SortaVoca`, `TipAmbalaze`,
`KolAmbVracena` (H — u potpisu stoji jednom), `ZbirnaID` →, `Stornirano`,
trace ×4, audit ×4.

**`tblPrijemnicaStavke`** — grain: **jedna klasa jednog prijema**
`PrijemnicaStavkaID` (PK `PRS-`), `PrijemnicaID` →, `RedniBroj`, `Klasa`,
`Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg`, **`Fakturisano`**, **`FakturaID`**.

> `Fakturisano`/`FakturaID` su **line-level**, i to je jedini izuzetak od pravila
> „polje iz potpisa". Razlog je ponašanje: `GetPrijemniceByKupac(samoNefakturisano)`
> (`modDokumenta.bas:2288`) filtrira po redu, dakle po klasi — korisnik danas
> fakturiše Klasu I bez Klase II. Na headeru bi se delimična fakturisanost tiho
> izgubila.

**Vlasnik upisa:** `modDokumenta` + `modFaktura` (samo `Fakturisano`/`FakturaID`,
kroz imenovani API). Danas 4 pisca.

---

### 4.5 `tblFakturaStavke` — jedina izmena

Dodaje se `PrijemnicaStavkaID` → kao **kanonski** izvor. `PrijemnicaID` ostaje
denormalizovan (read-modeli, poruke, SEF mapiranje).

Količina/cena/klasa ostaju snapshot na stavci — faktura mora biti istorijski
stabilan dokument.

---

### 4.6 Ledgeri — nepromenjen koncept, promenjen target

**`tblAmbalaza`** ostaje ledger; saldo se izvodi pri čitanju
(`docs/AMBALAZA_MODEL.md`). Menja se samo šta `DokumentID` pokazuje:
**header ID**, ne fizički klasni red. Time nestaje parsiranje `"OTK-1 + OTK-2"` u
`GetKooperantAmbOpening` (`modAmbalaza.bas:359`).

**`tblNovac`** ostaje ledger. `OtkupID` pokazuje na **header**. Time nestaje
primary-row hack (`SaveNovac(otkupID:=primaryID)`, `modOtkup.bas:279`).

> Alokacioni model (`tblNovacAlokacije`) je **van opsega v1**. Model ostavlja
> prostor, ne gradi ga.

**`tblPaletaStavka`** — FK prelazi sa `PrijemnicaID` na `PrijemnicaStavkaID`, jer
red već nosi `Klasa`/`VrstaVoca`/`SortaVoca`, tj. grain mu je klasa. Kolone
`BrojPrijemnice` i `BrojZbirne` se brišu.

---

## 5) PK/FK graf

```
KOOPERANT ─┐
PARCELA   ─┤
KULTURA   ─┤
STANICA   ─┼──> tblOtkup ──1:N──> tblOtkupStavke
VOZAC     ─┘        ^
                    │ OtkupID
              tblOtpremnicaIzvori          <== SASTAV verzije otpremnice (KANON)
                    v
              tblOtpremnica ──1:N──> tblOtpremnicaStavke
                    ^
                    │ OtpremnicaID
               tblZbirnaIzvori             <== SASTAV verzije zbirne (KANON)
                    v
                tblZbirna ──1:N──> tblZbirnaStavke   [sadrzaj verzije]
                    ^
                    │ ZbirnaID (1:N -- namera 1:1, meko)
              tblPrijemnica ──1:N──> tblPrijemnicaStavke

  Pripadnost drze ISKLJUCIVO tabele *Izvori. U ciljnom modelu nema kolona
  Otkup.OtpremnicaID ni Otpremnica.ZbirnaID -- "gde je sada" se racuna.
                                            │
                          ┌─────────────────┼─────────────────┐
                          │ 1:1             │ 1:N             │
                          v                 v                 │
                  tblFakturaStavke   tblPaletaStavka          │
                          │                 │                 │
                          v                 v                 │
                    tblFakture         tblPaleta              │
                          │                                   │
                          v                                   │
                        tblSEF                                │
                                                              │
  tblAmbalaza.DokumentID ──> header ID (Otkup/Otpremnica/Prijemnica)
  tblNovac.OtkupID       ──> Otkup header ID
```

---

## 6) Invarijante

### 6.1 Otkup — vrednost i isplata

```
VrednostOtkupa(OtkupID) = SUM(stavke.Kolicina * stavke.Cena)
Isplaceno(OtkupID) = "Da"  <=>  SUM(tblNovac za OtkupID) >= VrednostOtkupa
```

`Isplaceno` je **izvedeno** (A5), računa se nad headerom, i računa se **jednom po
dokumentu**.

> Raniji tekst je ovde opisivao „primary-row bug": gotovina se upisuje samo na
> Klasu I, pa Klasa II nikad ne dobije `Isplaceno`. **To nije bug nego mrtav
> kod** — keš uopšte ne ulazi kroz otkupni list (§4.1b). Putanja koja stvarno
> postavlja `Isplaceno` je avans, i ona radi ispravno (golden B2/B3).
> Ostaje da vrednost dokumenta posle refaktora bude `SUM(stavke)`.

**`Isplaceno` nije polje.** Ranija verzija ovog odeljka je govorila „jedno polje
na headeru" -- to je bilo pola koraka. Kolona i knjiga mogu da se raziju, i
`modProductionHealthCheck.Check_OtkupPaymentConsistency` postoji bas zato da to
uhvati. Bez kolone nema sta da se ne slaze, pa i ta provera odlazi (S4.1c).

### 6.2 Zbirna = zbir svojih aktivnih otpremnica

```
za svaku klasu K:
  SUM(OtpremnicaStavke.Kolicina)
      gde OtpremnicaID IN (tblZbirnaIzvori gde ZbirnaID = X)
  ==
  ZbirnaStavke.Kolicina
      gde ZbirnaID = X i Klasa = K
```

KG po klasi → **hard**. Ambalaža po klasi → **hard**; ukupno je posledica.

> **Ispravka ranije formulacije** („ambalaža ukupno hard, po klasi soft").
> Zbirna fizički nosi gajbice po klasama, a kg i gajbice zajedno služe za
> zatvaranje i kontrolu transporta — neslaganje po klasi je stvarna greška, ne
> zaokruživanje. Zatečeni writer to i potvrđuje: `SaveZbirnaMulti_TX` prima
> `ukupnoAmb` **i** `ukupnoAmbII`, dakle ambalaža je oduvek bila per-klasa
> podatak.
>
> `RequireOcekivanoSeSlaze` u novom writeru već poredi obe veličine po klasi;
> ovim se dokument usklađuje sa kodom, a ne obrnuto.

> **Spaja se preko `tblZbirnaIzvori`** — to je jedini zapis pripadnosti, kolone
> na otpremnici nema. Da postoji, pomerala bi se pri ispravci sa `ZBR18` na
> `ZBR19` i invarijanta stare `ZBR18` više se ne bi mogla reprodukovati. Ovako
> sastav svake verzije ostaje proverljiv i za istorijski dokument (A15).
>
> Filtriranje po `aktivna` važi samo dok je zbirna `DRAFT`. Za izdatu verziju se
> uzimaju **tačno one otpremnice koje su u njoj bile** — njihov kasniji storno je
> razlog za novu verziju (A13), ne za menjanje ove.

Ista `BrojZbirne` na drugom `ZbirnaID` **više nije problem integriteta**.

### 6.3 Homogenost dokumenta

```
Sve stavke jednog dokumenta dele Vrsta, Sorta i TipAmbalaze
(oni su na headeru -- stavke se razlikuju SAMO po klasi).
```

Writer odbija stavku koja bi to prekršila, sa imenovanom greškom. Bez ove kapije
model strukturno dozvoljava mešovit dokument, koji invarijanta, štampa i faktura
nikad nisu videle.

### 6.4 Bar jedna stavka

Dokument bez stavki se ne upisuje. `Kolicina > 0` i `Cena > 0` po stavci (osim
zbirne, koja cenu nema).

### 6.5 Storno je dokument-level

`Stornirano` postoji **samo na headeru**. Stavke ga nemaju — line-level storno ne
postoji u domenu. `tbl*Stavke` idu u `BEZ_STORNA` registar `modSchemaGuard`-a.

### 6.6 Faktura ne prelazi izvor

```
FakturaStavka.Kolicina == PrijemnicaStavka.Kolicina   (puna kolicina, 1:1)
PrijemnicaStavka sme biti fakturisana najvise jednom
FakturaStavka.PrijemnicaStavkaID mora postojati i biti aktivna
Prijemnica.KupacID == Faktura.KupacID
```

### 6.7 Ambalaža je ledger

Saldo se **nikad** ne čuva kao kolona. Izvodi se iz `Ulaz − Izlaz` po entitetu i
tipu.

---

## 7) Eksterni identitet (PWA) — model, implementacija kasnije

Nalaz iz koda, zapisan da se ne bi ponovo istraživao:

PWA šalje **jedan record = jedna klasa = ceo dokument**. `buildOtkupRecord`
(`src/js/features/otkup/otkup-form.js:643`) emituje jedan `klasa`, jednu
`kolicina`, jednu `cena`, svež `clientRecordID` **i svež `brojDokumenta` po svakom
snimanju** (`generateBrojDokumenta`, `:493`).

Posledice za model:

- Ingest je **1 record → 1 header + 1 stavka**. Nema grupisanja po broju.
- Ne treba eksterni Document UID — `clientRecordID` već identifikuje dokument.
- `ClientRecordID` stoji na **headeru** (`tblOtkup`), ne na stavci.
- `IsDuplicateInMaster` (`modMasterSync.bas:1824`) skenira `tblOtkup.ClientRecordID`
  — ostaje tačno, jer header ostaje u `tblOtkup`.

Implementacija je van opsega ovog refaktora (radi se isključivo VBA). Model je
ovde da adapter kasnije ne izmišlja pravilo.

---

## 8) Šta model NE radi

- Ne uvodi `tblDokumenti` / EAV
- Ne uvodi generički Repository/App framework
- Ne uvodi `Cmd`/`Qry`/`Rules` slojeve
- Ne uvodi alokacione tabele bez poslovnog zahteva
- Ne menja format ID-a matičnih podataka
- Ne pretvara „1 zbirna = 1 prijemnica" iz mekog u tvrdo pravilo
- Ne dodaje cenu na zbirnu

---

## 9) Otvoreno

Ništa od gornjeg nije pretpostavka. Jedina stavka koja čeka odluku van koda:

| Pitanje | Zašto kod ne odgovara | Predlog |
|---|---|---|
| Da li `Otkup.OtpremnicaID` sme da se razlikuje po klasi | Šema je dozvoljavala, ali nema podataka koji bi rekli da li se dešavalo | **Ne** — header, po §3.2 |

Ako se ne slažeš sa tim predlogom, to je jedino mesto u modelu koje se menja.

---

## 10) Šta ovaj model zamenjuje

`docs/DOMEN/ZBR_IDENTITET.md` postaje **SUPERSEDED**. `GeneracijaID` je bio
ispravno rešenje za multi-row model dokumenta — surogat logičkog identiteta koji
je nedostajao. Sa headerom taj identitet postoji direktno, pa `GeneracijaID`,
`ZbirnaIdent`, resolveri i vlasničke kapije nemaju posao.

Fajl se ne briše — istorija odluke ostaje čitljiva.
