# Domain Model + Schema v1 — lanac dokumenata

> Ciljni model posle refaktora „header + stavke". Ugovor koji ga uokviruje:
> `ARCHITECTURE_CONTRACT.md`. Plan isporuke: `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md`.
>
> Status: **model usvojen, nije implementirano.** Ovo je specifikacija za
> implementaciju, ne opis koda.
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
| Otkup → Otpremnica | **N:1, promenljiva** | `docs/DOMEN/README.md` §1 („više blokova → jedna otpremnica"); `ReassignOtkupToOtpremnica_TX` (`modDokumenta.bas:4266`) dokazuje da se blok može premestiti | `Otkup.OtpremnicaID` (header) |
| Otpremnica → Zbirna | **N:1, promenljiva** | jedna `BrojZbirne` kolona danas; `RelinkOtpremniceToZbirna_TX` | `Otpremnica.ZbirnaID` |
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

### 3.2) Delimična alokacija — namerno NE modelujemo

Pitanje „može li jedna `OtkupStavka` delimično da završi u više Otpremnica"
danas nema poslovni zahtev, a šema ga ne podržava (`Otkup.OtpremnicaID` je jedno
polje). **ODLUKA:** ostaje prost FK na headeru otkupa. Ako se potreba pojavi,
uvodi se eksplicitna tabela `tblOtpremnicaIzvori (OtpremnicaStavkaID,
OtkupStavkaID, Kg)` — ne rasplinjava se FK „za svaki slučaj".

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
| `Datum`, `KooperantID`, `StanicaID`, `VozacID`, `ParcelaID`, `KulturaID` | → matični |
| `VrstaVoca`, `SortaVoca`, `TipAmbalaze` | H — u potpisu stoje jednom |
| `KolAmbIzdata` | H — OM izdao prazne kooperantu |
| `Novac`, `PrimalacNovca` | H — snapshot gotovine |
| `Isplaceno`, `DatumIsplate` | H — **izvedeno** iz `tblNovac` vs `SUM(stavke.Kolicina × Cena)`; v. §6.1 |
| `OtpremnicaID` | → `tblOtpremnica`, nullable |
| `ZbirnaID` | → `tblZbirna`, nullable, **denormalizovano** (nasleđeno od otpremnice) — nije kanonska membership |
| `VremeUnosa`, `Stornirano` | |
| `IspravkaOdID`, `ZamenjenSaID`, `CorrectionID`, `IzdatoStatus` | |
| `ClientRecordID` | eksterni identitet (PWA); v. §7 |
| audit ×4 | |

**`tblOtkupStavke`** — grain: **jedna klasa jednog bloka**

`OtkupStavkaID` (PK `OKS-`), `OtkupID` →, `RedniBroj`, `Klasa`, `Kolicina`,
`Cena`, `KolAmbalaze`, `BrutoKg`.

**Vlasnik upisa (A11):** `modOtkup`, `modSetup`.
Svi ostali (`modNovac`, `modStorno`, `modSledljivost`, `modAutoHladnjaca`,
`modBankaMapiranje`, `modMasterSync`, `modOtkupBlok`, `modStornoFlow`,
`modStornoRecovery`, `modDokumenta`) idu kroz API `modOtkup`-a. **Danas ih je 12.**

---

### 4.2 `tblOtpremnica` — **grain: jedna isporuka sa otkupnog mesta**

`OtpremnicaID` (PK `OTP-`), `BrojOtpremnice` (labela, scoped po stanici), `Datum`,
`StanicaID`, `VozacID`, `VrstaVoca`, `SortaVoca`, `TipAmbalaze`,
`ZbirnaID` → nullable, `Stornirano`, trace ×4, audit ×4.

**`tblOtpremnicaStavke`** — grain: **jedna klasa jedne otpremnice**
`OtpremnicaStavkaID` (PK `OPS-`), `OtpremnicaID` →, `RedniBroj`, `Klasa`,
`Kolicina`, `Cena`, `KolAmbalaze`, `BrutoKg`.

**Vlasnik upisa:** `modDokumenta` (ili nov `modOtpremnica`). Danas 4 pisca.

---

### 4.3 `tblZbirna` — **grain: jedan transport ka kupcu/hladnjači**

`ZbirnaID` (PK `ZBR-`), `BrojZbirne` (labela, scoped po vozaču), `Datum`,
`VozacID`, `KupacID`, `Hladnjaca`, `Pogon`, `VrstaVoca`, `SortaVoca`,
`TipAmbalaze`, `Stornirano`, trace ×4, audit ×4.

**`tblZbirnaStavke`** — grain: **jedna klasa jedne zbirne**
`ZbirnaStavkaID` (PK `ZBS-`), `ZbirnaID` →, `RedniBroj`, `Klasa`, `Kolicina`,
`KolAmbalaze`.

> **Zbirna nema cenu.** `tblZbirna` je nikad nije imala i `SaveZbirnaMulti_TX` je
> ne prima (`modDokUnos.bas:422`). Ne dodavati je.

**Izvor istine:** zbirna je **agregat** — otpremnice su izvor. Njene stavke su
**keš** (A5), i invarijanta §6.2 to dokazuje pri svakoj izmeni.

**Vlasnik upisa:** `modDokumenta` + `modDokumentInvariant` (rekalkulacija).
Danas 5 pisaca.

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
VOZAC     ─┘        │
                    │ OtpremnicaID (N:1, promenljiva)
                    v
              tblOtpremnica ──1:N──> tblOtpremnicaStavke
                    │
                    │ ZbirnaID (N:1, promenljiva)
                    v
                tblZbirna ──1:N──> tblZbirnaStavke        [agregat / kes]
                    ^
                    │ ZbirnaID (1:N -- namera 1:1, meko)
              tblPrijemnica ──1:N──> tblPrijemnicaStavke
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

> Ovo ispravlja postojeći bug: danas se `Isplaceno` računa po redu
> (`modNovac.bas:1240`) ali se gotovina upisuje samo na primarni red
> (`modOtkup.bas:279`), pa red Klase II nikad ne dobije `Isplaceno` iako je
> kooperant plaćen u celosti.

### 6.2 Zbirna = zbir svojih aktivnih otpremnica

```
za svaku klasu K:
  SUM(OtpremnicaStavke.Kolicina)  gde Otpremnica.ZbirnaID = X i aktivna
  ==
  ZbirnaStavke.Kolicina           gde ZbirnaID = X i Klasa = K
```

KG po klasi → **hard**. Ambalaža ukupno → hard, po klasi → soft.
(Nepromenjeno pravilo; menja se samo ključ spajanja — `ZbirnaID` umesto
`BrojZbirne`.)

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
