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
| `SourceCreatedAt` | vreme nastanka **na izvoru** (PWA `GS_CREATED_AT`); prazno za desktop |
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
| desktop unos | `CreatedAt` — isto značenje, kolona je bila udvajanje; `SourceCreatedAt` ostaje prazno |
| PWA uvoz | `CreatedAt` je vreme **uvoza**; vreme unosa na terenu je druga činjenica → **`SourceCreatedAt`** |

Štampa ne bira po toku nego po popunjenosti:

```
PrikazVremena = SourceCreatedAt  ako postoji
                CreatedAt        inace
```

`SourceCreatedAt` nije izmišljen podatak — PWA sheet ga već nosi kao
`GS_CREATED_AT`. Dvosmisleno „vreme unosa" ne ostaje.

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

### 4.1f Matični podaci — FK-ovi, kultura i parcela

**Sve četiri veze ka matičnim podacima su pravi FK-ovi.** Neprazan string
nije dokaz da red postoji, a dokument sa slomljenim FK-om izgleda ispravno sve
dok ga neko ne spoji sa matičnim podacima — a to je po pravilu izveštaj ili
isplata.

```
KooperantID  -> tblKooperanti   TACNO jedan red
StanicaID    -> tblStanice      TACNO jedan red
KulturaID    -> tblKulture      TACNO jedan red   (+ snapshot vrsta/sorta)
ParcelaID    -> tblParcele      TACNO jedan red   (+ vlasnistvo), opciono
```

Nula pogodaka znači da veza pokazuje na nešto čega nema; dva i više da se ne
zna na šta pokazuje. Oba su tvrda greška.

**`StanicaID` se NE izvodi iz kooperanta.** To su dve različite činjenice, i
kod ih drži razdvojene:

| Polje | Šta je | Ko ga postavlja |
|---|---|---|
| `tblKooperanti.StanicaID` | **matično** otkupno mesto kooperanta | šifarnik; banka po njemu razvrstava uplate (`modBankaMapiranje:609,1049`) |
| `tblOtkup.StanicaID` | mesto **gde je otkup obavljen** | desktop: zaključana sesija (`modStanicaLock`, `gActiveStanica`), operater bira `cmbOtkupnoMesto` (`modOtkupBlok:729`) |

PWA ingest (`modMasterSync:1951`) uzima stanicu iz kooperanta **samo zato što
nema sesiju** — to je fallback jednog adaptera, ne pravilo domena. Zato
`Kooperant.StanicaID = Otkup.StanicaID` **nije** invarijanta: isti kooperant
sme da preda robu na drugoj stanici, i `ChangeStanica` postoji baš zato što
operater menja stanicu unutar iste sesije.

**`KulturaID` se RAZREŠAVA, nikad ne fabrikuje.**

Zatečeno stanje je gore nego što izgleda — fabrikuje se na **dva** mesta, i to
jedno od njih nije uvoz nego sam desktop writer:

```
modOtkup.bas:556      kulturaID = LookupValue(tblKulture, "VrstaVoca", vrsta, "KulturaID")
modOtkup.bas:559      If Len(kulturaID) = 0 Then kulturaID = vrsta & "-" & sorta
modMasterSync.bas:1959-1960   isti obrazac
```

Oba traže **samo po `VrstaVoca`** (sorta se ignoriše), a kad ne nađu — sklope
string koji izgleda kao FK a ne pokazuje ni na šta. Takav „ID" onda uđe u
dokument i preživi zauvek.

Ciljno pravilo:

```
(VrstaVoca, SortaVoca)  ->  TACNO jedan KulturaID

0 pogodaka   -> GRESKA
2+ pogodaka  -> GRESKA
1 pogodak    -> koristi taj ID
```

**Razrešavanje radi adapter, ne writer.** `CreateOtkup_TX` prima gotov
`KulturaID`; desktop i PWA adapter su ti koji iz UI vrednosti dolaze do matičnog
podatka. Writer proverava dvoje: da FK postoji, i da se snapshot
`VrstaVoca`/`SortaVoca` na dokumentu slaže sa tom kulturom.

Razlog za tu podelu: writer koji sam radi lookup mora da poznaje UI semantiku
(šta znači prazna sorta, šta se radi sa razmacima), a to je tačno mesto na kom
je fabrikovanje i nastalo.

**Parcela mora pripadati kooperantu — HARD.**

```
ako je ParcelaID zadat:
    Parcela.KooperantID = Otkup.KooperantID     obavezno
```

Tuđa parcela ne prolazi kanonski writer. Neslaganje **kulture** parcele ostaje
`warning` sa override-om, kao danas — za tvrdo pravilo tu nema dovoljno osnova, a
operater ima legitimne slučajeve.

---

### 4.1g Kada je prazno legitimno — sorta i tip ambalaže

Dva polja smeju da budu prazna, i to **ne odlučuje writer**. Ako writer traži
više nego domen, tiho je pooštrio poslovno pravilo — a to je ista klasa greške
kao i da ga je olabavio, samo se prijavljuje kao „ne mogu da snimim".

```
SortaVoca     kljuc OBAVEZAN, vrednost sme prazna
              prazna prolazi TACNO kad je i sama kultura bez sorte
              (pravilo je vec tu: snapshot mora da odgovara kulturi, S4.1f)

TipAmbalaze   kljuc OBAVEZAN, vrednost sme prazna
              obavezan kad SUM(stavke.KolAmbalaze) > 0 ILI KolAmbIzdata > 0
```

Oba pravila su **merena nad zatečenim ekranom**, ne izmišljena:
`modOtkupUnos:120` traži sortu samo kad je `IsValidacijaUnosa()` uključena, a
`modOtkupUnos:158` traži tip ambalaže kad `kolAmb > 0 Or kolAmbII > 0 Or
kolAmbIzd > 0` — dakle **i zbog izdate**. Izdata ambalaza bez tipa je gajba koja
je otišla kooperantu a ne zna se koja.

**Zašto ključ mora da postoji i kad vrednost sme da bude prazna:** bez toga se
tipfeler u imenu polja ne razlikuje od namerno praznog polja. Isti razlog drži
zatvoren spisak ključeva — na headeru i na **stavci**. Na stavci je to jedina
odbrana za `BrutoKg`: ostala polja su obavezna pa tipfeler u njima padne sam od
sebe, a `BruttoKg` bi se samo ignorisao i bruto unos bi tiho postao neto.

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
`StanicaID`, `VozacID`, **`KulturaID`**, `VrstaVoca`, `SortaVoca`, `TipAmbalaze`,
**`PredlogCena`**, `Stornirano`, trace ×4, audit ×4.

> **`Cena` → `PredlogCena` je deo ciljne šeme, ne kozmetika.** §13b je već odlučio
> da polje ostaje ali „preimenovano u ono što jeste"; ime `Cena` je semantički
> mamak, a legacy štampa tu kolonu već koristi kao pravu finansijsku cenu. Skela
> je **ne preimenuje** — rename je posao Otpremnica cutover-a, zajedno sa
> čitaocima. Do tada nov pisač piše u zatečenu `Cena`, a spec nosi ciljno ime.

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

**`BrutoKg` se ne sabira parcijalno.** Bruto je poznat samo kad je unos bio bruto;
prazno polje **nije nula**. Da se sabira kao nula, otpremnica bi dobila fizički
nemoguć red — jedan izvor `500 bruto / 480 neto`, drugi `300 neto` bez bruta, i
rezultat bi bio `Kolicina 780, BrutoKg 500`, dakle bruto manji od neta. Pravilo:

```
OtpremnicaStavka.BrutoKg = SUM(izvori) samo ako SVAKA izvorna stavka te klase
                           ima poznat bruto; cim je jedna prazna -> PRAZNO
```

Ne rekonstruiše se iz težine ambalaže: tara je poznata samo u trenutku nastanka
otkupa (§4.1d).

**Vlasnik upisa:** `modDokumenta` (ili nov `modOtpremnica`). Danas 4 pisca.

---

### 4.2a Članstvo i izvedena polja

Otpremnica je **izvedeni dokument**, isto kao zbirna. Zato njen kanonski writer
prima **izvore**, a ne stavke:

```
  stavke     OCEKIVANJE  ono sto je operater prijavio da otpremnica nosi;
                         upisuju se ODMAH, sa draftom
  clanstvo   POVEZANO    SUM nad otkupnim stavkama clanova
  preostalo  = ocekivano - povezano, po klasi
  header     PRIMLJEN    Datum, StanicaID, VozacID, BrojOtpremnice,
                         KulturaID + snapshot VrstaVoca/SortaVoca/TipAmbalaze
  izdavanje  zahteva ocekivano = povezano, pa ZAMRZAVA stavke
```

**Stavke drafta su OCEKIVANJE, ne izvedeni keš.** To nisu dva izvora istine nego
**dve različite činjenice**: šta je vozač/operater prijavio da nosi, naspram šta
su otkupni listovi dokumentovali. Njihov *mismatch* je koristan poslovni signal —
on je i razlog zašto panel postoji.

Mereno u zatečenom kodu: „očekivano" danas živi na `Otpremnica.Kolicina`, i to
čitaju **četiri** sposobnosti panela:

| Mesto | Sposobnost |
|---|---|
| `modOtkupBlok:223` | upozorenje na prekoračenje pri unosu bloka |
| `modOtkupBlok:262` | auto-deselekcija kad `Preostalo` padne na 0 |
| `modOtkupBlok:500` | kolona „Ostatak" + filter „samo nezavršene" |
| `modOtkupBlok:1384` | sažetak `Ukupno / Napisano / Preostalo` |

Pošto `Kolicina` u ciljnom modelu **odlazi sa headera na stavku**, očekivanje bez
stavki drafta nema gde da živi — i sve četiri bi pale na cutover-u.

Pri izdavanju stavke **prestaju** da budu očekivanje i postaju sadržaj verzije;
isti prelaz koji A5/A13 opisuju, gledan sa ulazne strane (`REFAKTOR` §13b).

**`StanicaID` i `KulturaID` se PRIMAJU, pa proveravaju** — ne izvode se. Draft
nastaje pre ijednog izvora, a izvođenje nad praznim skupom nije definisano.

Kod kulture to nije samo pitanje praznog skupa nego **smera podataka**: zatečeni
glavni tok ide *otpremnica → otkup*, ne obrnuto. Operater klikne otpremnicu, a
ona **prefiluje** formu otkupa stanicom, vrstom, sortom, vozačem i cenom
(`modOtkupBlok.PrefillLeftForm:706-730`). Otpremnica koja kulturu saznaje tek pri
izdavanju ne bi imala čime da prefiluje prvi otkup — a taj otkup treba da je od
nje i dobije.

`KulturaID` je relaciona istina, `VrstaVoca`/`SortaVoca`/`TipAmbalaze` su njen
snapshot na dokumentu — isti par kao na otkupu (§4.1f). Svaki izvor koji se posle
doda mora da se slaže sa stanicom **i** sa `KulturaID`.

#### Dva ulaza, jedan core

Otpremnica **ima persistentan `DRAFT`** — za razliku od otkupa, kod koga je forma
draft (§4.1e). To nije izbor implementacije nego zatečeni glavni desktop tok:
panel napravi otpremnicu praznu i blokovi se kače naknadno
(`modOtkupBlok.LinkOtkupIDsToOtpremnica`).

```
CreateOtpremnicaDraft_TX(h, ocekivano)   -> OTP-...  DRAFT + stavke ocekivanja
UpdateOtpremnicaDraft_TX(otpID, h, ocekivano)      samo DRAFT

DodajOtpremnicaIzvor_TX(otpID, otkupID)            samo DRAFT
UkloniOtpremnicaIzvor_TX(otpID, otkupID)           samo DRAFT

GetOtpremnicaProgress(otpID)             -> ocekivano / povezano / preostalo
                                            po klasi

IzdajOtpremnicu_TX(otpID)                -> REVALIDIRA izvore,
                                            zahteva ocekivano = povezano,
                                            zamrzne stavke, IZDATO

CreateOtpremnicaIzIzvora_TX(h, izvori)      jedan potez, JEDNA transakcija;
                                            ocekivanje IZVEDENO iz izvora
```

Jednopotezni ulaz sme da izvede očekivanje iz samih izvora **jer tu nezavisnog
operaterskog očekivanja nema** — auto-lanac i PWA ne prijavljuju šta nose, oni to
znaju. Ručni tok bez očekivanja bi ostao bez svoje jedine kontrole.

Zato „očekivanje izvodim kasnije" ostaje **privatan** signal: `CreateOtpremnicaDraft_TX`
odbija `Nothing` tvrdo. Da ga prosleđuje u core, ručni ulaz bi umeo da napravi
`DRAFT` bez ijedne stavke očekivanja — dokument koji nema šta da meri, a izgleda
ispravno. Prazna `Collection` je **nešto drugo** i takođe je odbijena, ali sa
svojim razlogom.

#### Izmena zaglavlja revalidira već upisano članstvo

`UpdateOtpremnicaDraft_TX` sme da menja `StanicaID` i `KulturaID` — a upravo po
njima se sudi da li izvor sme da uđe. Draft sa stanicom `ST1` i članom sa `ST1`,
prebačen na `ST2`, nosio bi člana koga `Dodaj` **nikad ne bi primio**.

Izdavanje bi to kasnije uhvatilo, ali invarijanta ne sme da bude prekršena
**između dva klika**: `GetOtpremnicaProgress` u međuvremenu uredno računa
nevalidno članstvo. Zato izmena, posle upisa novog zaglavlja, revalidira **svakog**
postojećeg člana prema novim vrednostima; pad rollback-uje ceo update, pa staro
zaglavlje i staro očekivanje ostaju netaknuti.

#### Izdavanje revalidira, ne veruje `Dodaj`-u

Između `Dodaj` i `Izdaj` prolazi vreme — u panelu i po nekoliko sati. Otkup se u
međuvremenu može stornirati ili ispraviti. Zato `IzdajOtpremnicu_TX` ponavlja
**ceo** skup provera nad svakim članom, nezavisno od toga šta je `Dodaj` proverio:
postoji tačno jednom, nije storniran, ista stanica, ista kultura, i kanonsko
članstvo ga i dalje veže baš za **ovu** otpremnicu.

Bez toga je `DRAFT → dodaj OTK1 → storno OTK1 → Izdaj` izdavao dokument iz
storniranog izvora. Klasičan TOCTOU: provera i upotreba nisu u istom trenutku.

Jednopotezni ulaz nije druga implementacija nego **isti core**: auto-lanac
(`modAutoHladnjaca`) i PWA prave otpremnicu bez ijednog međukoraka, pa bi ih tri
poziva naterala da drže tuđe stanje.

**`DRAFT` je jedino stanje u kom se članstvo menja** (A14, A15). Posle izdavanja
`Dodaj`/`Ukloni` dižu grešku — sastav izdatog dokumenta je istorijska činjenica,
a izmena je nova verzija (A13). Izdavanje je i trenutak u kom stavke nastaju:
draft ih **nema**, jer bi inače postojale dve istine o istoj količini — jedna u
stavkama, druga u članstvu koje se još menja.

**Izvedeno, ne primljeno** — isti razlog kao kod zbirne (PR3): polje koje writer
prima a moglo je da izračuna je drugi izvor istine, i tiho se razilazi sa prvim.
Neslaganje među izvorima time postaje **greška pri izvođenju**, ne zaseban
validator koji neko može da zaobiđe.

Tvrde kapije za ulazak otkupa u otpremnicu:

| Pravilo | Zašto |
|---|---|
| otkup postoji i **nije storniran** | storniran dokument nije roba |
| otkup je **`IzdatoStatus = IZDATO`** | veza traži *izdat dokument*, ne bilo koji red koji slučajno ima stavke |
| otkup **nije već u drugoj aktivnoj otpremnici** | isto pravilo kao `AktivnoClanstvoPoKanonu` za zbirnu (A15) |
| svi izvori sa **iste stanice** | grain je *jedna isporuka **sa otkupnog mesta*** — header nosi jedan `StanicaID`, pa dve stanice ne mogu ni da se predstave |
| svi izvori iste **`KulturaID`** | header nosi jednu kulturu; vrsta/sorta se poklapaju posledično |
| svi izvori istog **`TipAmbalaze`** | header nosi jedan tip, pa *20 plastičnih + 30 drvenih gajbi* nije 50 gajbi |
| bar jedan izvor **pri izdavanju** | otpremnica bez ijednog otkupa nije isporuka; prazan `DRAFT` je legitiman |

**Sve se proveravaju pri `Dodaj`, ne tek pri izdavanju.** Razlog nije urednost
nego read-model: `GetOtpremnicaProgress` do izdavanja uredno računa ono što u
članstvu stoji, pa bi nehomogen draft prikazivao broj koji semantički ne znači
ništa — zbir dve različite gajbe.

`VozacID` je **na headeru i prima se** — vozač je odluka otpreme, ne svojstvo
otkupa (§4.1c). Zato izvori o vozaču ne govore ništa i nema šta da se poklapa.

---

### 4.2b Šta je mereno pre skele

**1. Jedna poslovna otpremnica su danas N redova, i dva pisca ih vezuju
različito.** `modAutoHladnjaca:213,258` zove `SaveOtpremnica_TX` **dvaput** sa
istim `brOtp` — jedan red po klasi, dva različita `OtpremnicaID`. Ali:

```
modAutoHladnjaca   Klasa I -> otpID     Klasa II -> otpID2    (podela po klasi)
modOtkupBlok:1425  SVI otkupi bloka  -> jedan mActiveOtpID    (bez podele)
```

Isti pojam („koja otpremnica nosi ovaj otkup") ima **dva različita odgovora**
zavisno od toga koji ga je pisač upisao. To nije rubni slučaj nego posledica
toga što otpremnica nema jedan identitet.

**2. Broj → ID razrešenje bira PRVI red.** `modScrDokumenti:502`
(`OtpIdZaBroj`) radi `LookupValue(tblOtpremnica, BrojOtpremnice, broj,
OtpremnicaID)`, a `LookupValue` (`modDataAccess:608`) vraća **prvi pogodak i
izlazi**. Kada broj nosi dva reda, izbor je proizvoljan. Rezultat ide u
`PrintSpecifikacija` → `RenderSpec`, koji filtrira otkupe po `OtpremnicaID`.

> **Status: nije reprodukovano.** Iz koda sledi da specifikacija štampana za
> otpremnicu čije su klase razdvojene (auto-hladnjača put) prikazuje samo jednu
> klasu. Nije izmereno nad podacima, pa se vodi kao **rizik**, ne kao nalaz.
> Header+stavke ga uklanja bez zasebne ispravke: broj tada nosi jedan red.

**3. `Otkup.OtpremnicaID` se NE briše u skeli.** Mereno: **39** ne-test
korišćenja u **15** modula, od toga **5 pisača** (`modAutoHladnjaca:374`,
`modDokumenta:5413`, `modMasterSync:2434`, `modOtkupBlok:1459`,
`modStornoFlow:2569`).

> Razlika u odnosu na `Otpremnica.ZbirnaID`, koji je u PR3 obrisan odmah: tamo
> je čitalaca bilo malo i svi su imali imenovanog naslednika. Ovde bi brisanje
> u skeli oborilo izveštaje, štampu i storno tok. Kolona odlazi u **PR7**,
> zajedno sa svojim pisačima.

**4. Danas ne postoji nijedno pravilo članstva.**
`ReassignOtkupToOtpremnica_TX` (`modDokumenta:5382`) proverava samo da cilj
postoji i da nije storniran — ni stanicu, ni vrstu/sortu, ni da li je otkup već
negde. Pravila iz §4.2a su zato **nova**, ne prepisana.

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
KULTURA   ─┼──> tblOtkup ──1:N──> tblOtkupStavke
STANICA   ─┘        ^
                    │ OtkupID
              tblOtpremnicaIzvori          <== SASTAV verzije otpremnice (KANON)
                    v
                  VOZAC ────> tblOtpremnica ──1:N──> tblOtpremnicaStavke
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

**Obaveza za cutover: `SourceCreatedAt` mora biti vreme, ne tekst.** Skela ga
prima kao `String` i upisuje bez provere, jer u skeli nema pošiljaoca — jedini
pisač je test. U trenutku kad adapter počne da ga puni, provera formata ide uz
njega: kolonu čita štampa (§4.1c, `modPrint:591` nasleđuje `VremeUnosa`), pa bi
proizvoljan string završio na otkupnom listu kao vreme.

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

**Za Otkup i Zbirnu: ništa.** Sve što je ranije stajalo ovde je odlučeno.

Poslednja stavka — *„sme li `Otkup.OtpremnicaID` da se razlikuje po klasi"* — je
zatvorena time što **kolone nema**: pripadnost je `tblOtpremnicaIzvori`, na nivou
jednog otkup headera (§4.1e). Pitanje je prestalo da postoji, nije odgovoreno.

Otvoreno ostaje samo ono što po redosledu tek dolazi:

| Pitanje | Kad se rešava |
|---|---|
| iz čega se računa `manjak` posle uklanjanja `Otpremnica.Cena` iz obračuna | Otpremnica cutover (§13b) |
| da li draft-first tok dobija `CreateZbirnaDraft_TX` / `CreateOtpremnicaDraft_TX` | **Otpremnica skela (PR5)** — v. ispod |

> **Za otpremnicu ovo pitanje nije odloživo kao za otkup.** Kod otkupa je forma
> njegov draft, pa persistentnog `DRAFT`-a nema (§4.1e). Kod otpremnice je
> draft-first **glavni desktop tok**: panel napravi otpremnicu, pa se blokovi
> naknadno kače na `mActiveOtpID` (`modOtkupBlok.LinkOtkupIDsToOtpremnica`).
> Writer koji ume samo „izvori → izdato" nema gde da primi taj tok, a cutover
> otpremnice je PR7 — dakle pitanje stiže za dva koraka, ne „nekad".

---

## 10) Šta ovaj model zamenjuje

`docs/DOMEN/ZBR_IDENTITET.md` postaje **SUPERSEDED**. `GeneracijaID` je bio
ispravno rešenje za multi-row model dokumenta — surogat logičkog identiteta koji
je nedostajao. Sa headerom taj identitet postoji direktno, pa `GeneracijaID`,
`ZbirnaIdent`, resolveri i vlasničke kapije nemaju posao.

Fajl se ne briše — istorija odluke ostaje čitljiva.
