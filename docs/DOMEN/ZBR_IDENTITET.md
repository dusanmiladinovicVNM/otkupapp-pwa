# ZBR-IDENT-01 / ZBR-PARENT-01 — identitet zbirne i vezivanje prijemnice

> **Status:** §§1–3 opisuju **zatečeni kod** (svaka tvrdnja nosi izvor sa brojem
> linije). §§4–7 su **ugovor koji još NIJE implementiran** — to je plan za L1,
> ne opis stanja. Ne čitaj ih kao opis ponašanja.
>
> Ovaj fajl je odgovor na `KNOWN_ISSUES.md` KI-007, koji traži da se invarijanta
> ZBR-IDENT-01 definiše pre nego što se dira core.

## 1) Tri stvari koje se lako pomešaju

|  | Šta je | Primer |
|---|---|---|
| **fizički red** | red u `tblZbirna` | Klasa I; Klasa II |
| **logički dokument** | `GeneracijaID` | jedna zbirna, dva reda |
| **poslovni broj** | `BrojZbirne` | ono što picker prikazuje |

```
fizički red  ->  GeneracijaID  ->  logički dokument
```

`GeneracijaID` je **perzistiran, neproziran identitet** logičkog dokumenta.
`broj + VozacID + KupacID` je **scope** u kome se generacija nalazi ili kuje —
nije sam identitet. Isto pravilo grupisanja već stoji u
`modScrOporavak.bas:770`: generacija kad postoji, inače broj + pun vlasnik.

**Broj je labela, ne identitet.** `SaveZbirnaMulti_TX` zove `SaveZbirna` dvaput
sa **istim** `broj/VozacID/KupacID` — dvoklasna zbirna je JEDAN dokument na DVA
reda.

## 2) Šta kod danas stvarno garantuje

**Svaki upisan red dobija `GeneracijaID`.** Tri i samo tri writer-a rade
`AppendRow` u `tblZbirna`, i sva tri odmah zovu `ApplyGeneracijaID`:

| Writer | AppendRow | ApplyGeneracijaID |
|---|---|---|
| `modDokumenta.SaveZbirna` | `modDokumenta.bas:633` | `:636` |
| `modMasterSync` (PWA import) | `modMasterSync.bas:3175` | `:3180` |
| `modDokumentInvariant` (rekalkulacija) | `modDokumentInvariant.bas:412` | `:421` |

Lanac garancija: `RequireColumnIndex(COL_GENERACIJA_ID)` pada glasno ako kolone
nema · `modSetup.EnsureSledljivostSchema` (`modSetup.bas:1255`) je pravi na
svakom startu · `NewGeneracijaID` diže grešku ako `GetNextID` vrati prazno ·
upis ide kroz `RequireUpdateCell`.

**`GeneracijaIDZaBrojArr` isključuje stornirane.** `modDokumenta.bas:909`:

```vb
If IsArray(data) Then data = ExcludeStornirano(data, tableName)
```

Ta jedna linija nosi **dva** ponašanja i ne sme da nestane:

- **Multi-klasa:** Klasa I se snima prva (ništa aktivno u scope-u → nova
  generacija), pa Klasa II — red Klase I je tada aktivan i u scope-u → generacija
  se **ponovo koristi**. Tako dva reda dobijaju jednu generaciju.
- **Ispravka:** posle storna je jedini red u scope-u storniran → isključen →
  kuje se **nova** generacija. Original i ispravka su **različiti** logički
  dokumenti. Tvrdi to `modTest` `T_IstiBrojRazliciteGeneracije_NijeIstiDokument`.

**Broj je jedinstven redovnim putem.** Format `x/ddmmyy[-rb]` gde je `x`
numerički deo vozača; `SuggestNextBroj` za `ZBR` uz to vrti
`Do While BrojZbirneExists(...)` nad celom tabelom.

## 3) Rupa — gde jedinstvenost niko ne čuva

Dvosmislen broj **ne nastaje** redovnim putem. Nastaje ručnim unosom sa ugašenim
auto-brojem, uvozom, ili ispravkom u tabeli. Na tim putevima:

- `BrojZbirneExists` je `Private` i zove se **samo iz predloga**, ne pri upisu.
- `modDokUnos.ZbirnaValidiraj:427` zove
  `CheckDuplicate(TBL_ZBIRNA, COL_ZBR_BROJ, ...)`, ali `CheckDuplicate`
  (`modDataAccess.bas:432`) **preskače stornirane** i poredi **sirovo**
  (`CStr(...) = searchValue` — bez `Trim`, case-sensitive). Storno-pa-ponovna-
  upotreba od drugog vlasnika prolazi, a i `" 5/070926 "` prolazi pored
  `"5/070926"`.
- `modDokUnos.PrijemnicaValidiraj:577-597` → `ZbirnaPostoji`
  (`modDokumenta.bas:329`) odgovara samo na „postoji li taj broj" — ne na „koji
  dokument".
- `modStorno.RequireJedanVlasnikIkadPoBroju` (`modStorno.bas:2467`) je ista
  kapija za mutaciju, ali njen `VlasniciPoBroju` (`:2500`) poredi broj
  **case-sensitive** (`:2523`).

Posledica: ako duplikat nastane mimo generatora, prijemnica može da se veže na
pogrešan red i **nijedan sloj to ne prijavljuje**.

---

# Ugovor (NIJE IMPLEMENTIRANO — plan za L1)

## 4) Resolver

```
normalizedBroj                     String    ' Trim$, poredjenje vbTextCompare
integrityStatus                    OK | INTEGRITY_ERROR
activeLogicalCount                 Long      ' distinct non-empty GeneracijaID, aktivni redovi
activeOwnerCount                   Long      ' distinct (VozacID,KupacID), aktivni redovi
historicalOwnerCount               Long      ' distinct (VozacID,KupacID), SVI redovi IKAD
scopeProvided                      Boolean
historicalOwnerIsScope             Boolean   ' True samo kad je historicalOwnerCount = 1
                                             ' i taj vlasnik = prosledjeni scope
matchingScopeActiveLogicalCount    Long
resolutionStatus                   NONE | UNIQUE | OWNER_MISMATCH | CURRENT_AMBIGUOUS
selectedGeneracijaID               String    ' popunjeno samo kad je UNIQUE
selectedVozacID / selectedKupacID  String
```

Dve **nezavisne** dimenzije, dva pitanja:

| Dimenzija | Pitanje | Gleda stornirane? |
|---|---|---|
| `resolutionStatus` | „Da li je broj **sada** jednoznačno razrešiv u jedan logički dokument?" | ne |
| `historicalOwnerCount` | „Da li je broj **ikad** pripadao više od jednog vlasnika?" | da |

`UNIQUE` uz `historicalOwnerCount = 2` je **validno stanje, ne kontradikcija**:
broj je danas jednoznačan, ali ga je u prošlosti držao i drugi vlasnik.

`activeLogicalCount` i `activeOwnerCount` se **ne izvode jedan iz drugog**:
multi-klasa daje `1 / 1` iz dva reda, dva zasebna unosa istog vozača daju `2 / 1`.

`integrityStatus = INTEGRITY_ERROR` (aktivan red sa praznim `GeneracijaID`) gasi
sve ostalo — `resolutionStatus` se ne računa, svi potrošači blokiraju. Identitet
se **ne pogađa** iz broja i vlasnika.

## 5) `ZBR-ACTIVE-NUMBER-01` — kapija za kreiranje (F3)

Radi nad `normalizedBroj`. **Ne** delegira poređenje na `VlasniciPoBroju`, koji
je case-sensitive.

```
integrityStatus = INTEGRITY_ERROR                     -> BLOCK
activeLogicalCount > 0                                -> BLOCK   (bez izuzetka za vlasnika)
activeLogicalCount = 0 AND historicalOwnerCount = 0   -> ALLOW
activeLogicalCount = 0 AND historicalOwnerIsScope     -> ALLOW   (ispravka / re-entry)
inace                                                 -> BLOCK
```

| Aktivni log. dok. | Istorija vlasnika | Kandidat | Ishod |
|---|---|---|---|
| 0 | 0 | bilo koji | ALLOW |
| 0 | 1, isti | isti | ALLOW — ispravka |
| 0 | 1, drugi | bilo koji | BLOCK |
| 0 | >1 | bilo koji | BLOCK |
| >0 | bilo šta | bilo ko | BLOCK |

**Zašto strogo, bez izuzetka za istog vlasnika.** Provereno je da to ništa ne
lomi: `ZbirnaValidiraj` se zove **tačno jednom**, iz `modScrDokumenti.bas:868`, i
tek posle njega ide `ZbirnaUpisi` → `SaveZbirnaMulti_TX` → dva `SaveZbirna` —
validator nikad ne vidi red koji je sam upravo napisao. Uređivanja zbirne u mestu
nema; ispravka je storno pa nov unos, pa je tada `activeLogicalCount = 0`.

**Opseg kapije: ovo je UI kapija na F3, ne invarijanta tabele.** `modMasterSync`
(import) i `modDokumentInvariant` (rekalkulacija) upisuju ne prolazeći kroz
`ZbirnaValidiraj`. Zapisano namerno, da se niko kasnije ne osloni na to da je
tabela zaštićena.

## 6) Čitanje u F4 (prijemnica → roditelj)

| Stanje | Ishod |
|---|---|
| `INTEGRITY_ERROR` | hard block |
| `NONE` | postojeći BLOK/UPOZORENJE (`ZbirnaPostoji` + `PrijemnicaZbirnaBlokira()`), bez izmene |
| `CURRENT_AMBIGUOUS` | hard block, **bez** traženja kanonskog roditelja |
| `OWNER_MISMATCH` | hard block |
| `UNIQUE` | prolaz, `selectedGeneracijaID` je roditelj |

## 7) Odluke

| # | Odluka |
|---|---|
| D1 | Roditeljstvo preko `GeneracijaID`, ne preko broja. |
| D2 | `CheckDuplicate` se **ne menja** — nastavlja da preskače stornirane, pa ispravka-workflow na svih 6 tipova ostaje. Rupa se zatvara **pored** njega, ZBR-specifičnom kapijom. |
| D3 | Otpremnica **nije dete** zbirne. `OtpremnicaValidiraj` se ne dira. |
| D4 | Nema legacy fallbacka na `broj + vlasnik` kao identitet. AgriX se distribuira novim klijentima sa čistom bazom; prazan `GeneracijaID` je integritetska greška, ne alternativni oblik identiteta. |

## 8) Intervencije

**I1 — kanonski read-model aktivnih brojeva.** Distinct `normalizedBroj` iz
aktivnih zbirnih **∪** distinct `normalizedBroj(BrojZbirne)` iz aktivnih
prijemnica. Broj koji referencira aktivna prijemnica je zauzet i kad mu je
zbirna-red storniran.

**I2 — razrešavanje roditelja.** Poziva se **samo** kad je
`resolutionStatus = UNIQUE`. Kod `CURRENT_AMBIGUOUS` / `OWNER_MISMATCH` /
`INTEGRITY_ERROR` se **ne traži** kanonski roditelj. Biranje „najverovatnijeg"
roditelja iz dvosmislenog skupa je tiho pogađanje i zabranjeno je ugovorom.

## 9) Acceptance testovi

Svaki test tvrdi **svoj preduslov** pre glavne tvrdnje.

| # | Ulaz | Tvrdnja |
|---|---|---|
| A1 | `broj = ""` | `NONE`, svi count-ovi 0. F4 zadržava postojeći BLOK/UPOZORENJE. |
| A2 | broj koji ne postoji | `NONE`, `historicalOwnerCount = 0`. |
| A3 | jedna aktivna zbirna, scope se poklapa | `UNIQUE`; `1 / 1 / 1`, `historicalOwnerIsScope = True`. |
| A4 | jedna aktivna zbirna, **drugi** scope | `OWNER_MISMATCH`; `matchingScopeActiveLogicalCount = 0`, `selectedGeneracijaID = ""`. |
| A5 | dve aktivne zbirne, dva vlasnika | `CURRENT_AMBIGUOUS`; `activeLogicalCount = 2`, `activeOwnerCount = 2`. |
| A6 | multi-klasa | Preduslov: 2 fizička reda, **ista** `GeneracijaID`. Tvrdnja: `activeLogicalCount = 1`, `activeOwnerCount = 1`. |
| A7 | jedini red storniran | `activeLogicalCount = 0` → `NONE`; `historicalOwnerCount = 1`. |
| A8 | ikad dva vlasnika, aktivan jedan | `UNIQUE` **i** `historicalOwnerCount = 2` istovremeno. F3 kapija: BLOCK. |
| A9 | `Scr_Save` sa brojem koji drži aktivan drugi vlasnik | Broj redova u `tblZbirna` **nepromenjen**. |
| A10 | `Scr_Save`, legitimna zbirna | +1 red (odn. +2 za multi-klasu). |
| A11 | `ZbirnaValidiraj:427` → `CheckDuplicate` | Netaknut; stornirani se preskaču. Brana za D2. |
| A12 | re-entry posle storna, **isti** vlasnik | Prolazi kroz kapiju. Dobija **NOVU** `GeneracijaID` — posledica `modDokumenta.bas:909`, ne zahtev ugovora. Preduslov: generacija posle ≠ generacija stornirane. |
| A13 | prijemnica | `PrijemnicaValidiraj:577-597` nepromenjen na obe grane. |
| A14 | otpremnica | `OtpremnicaValidiraj` nema poziv ka resolveru. Brana za D3. |
| A15 | aktivan `5/070926` vlasnik A, kandidat `" 5/070926 "` vlasnik **B** | BLOCK. Preduslov: isti par **prolazi** kroz sirovo poređenje u `CheckDuplicate`. |
| A16 | resolver nad varijantama istog broja (razmaci, case) | Isti `normalizedBroj`, isti count-ovi, isti `selectedGeneracijaID`. |
| A17 | dva aktivna log. dok. **istog** vlasnika | `CURRENT_AMBIGUOUS`; `2 / 1`. Dvosmislenost nije pitanje vlasnika. |
| A18 | aktivan `5/070926` vlasnik A, kandidat `" 5/070926 "` vlasnik **A** | **BLOCK.** Preduslov: normalizovani brojevi jednaki i vlasnik isti. |
| A20 | aktivan red sa praznim `GeneracijaID` | `INTEGRITY_ERROR`; resolver **ne pogađa** identitet; `IntegritetUkupno` +1. |
| A21 | import / rekalkulacija | Ne prolaze kroz `ZbirnaValidiraj`; red ipak dobija `GeneracijaID`. Fiksira granicu kapije iz §5. |

A19 je **povučen** (bio je legacy `LogicalZbirnaKey` fallback) — v. D4. Broj se
ne reciklira, da se stariji zapisi ne bi pogrešno čitali.

## 10) Van opsega, imenovano

- **`ZBR-NORM-02`** — `VlasniciPoBroju` (`modStorno.bas:2523`) poredi broj
  case-sensitive. Nije deo L1; resolver ga zaobilazi sopstvenom normalizacijom.
- **MIG-005b** — dupla stavka dvoklasne zbirne u pickeru; blokiran je na ovome,
  jer ispravna de-duplikacija grupiše po logičkom dokumentu, a picker danas nosi
  samo broj. V. `UI_MIGRACIJA_KATALOG.md` §28.1f.
- Nema izmene `CheckDuplicate`, `OtpremnicaValidiraj`, `GeneracijaIDZaBrojArr`.
- Važe opšta pravila: bez novih `Private WithEvents`, `.frx` se ne dira, VBA
  izvor 100% ASCII, korisnički tekst kroz `modPoruke`.

## 11) Reference

| Tema | Gde |
|---|---|
| Poznato ograničenje koje ovo zatvara | `docs/KNOWN_ISSUES.md` KI-007 |
| Zašto de-duplikacija po broju nije rešenje | `docs/UI_MIGRACIJA_KATALOG.md` §28.1f |
| Pravilo grupisanja u praksi | `modScrOporavak.bas:770` |
| Kapija za mutaciju po broju (isti obrazac) | `modStorno.RequireJedanVlasnikIkadPoBroju` |
| Storno nije brisanje | `docs/DOMEN/README.md` §2 |
