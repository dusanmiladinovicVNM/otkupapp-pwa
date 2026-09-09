# ZBR-IDENT-01 / ZBR-PARENT-01 — identitet zbirne i vezivanje prijemnice

> **KI-007 je zatvoren:** resolver, F3 prevencija (`ZBR-ACTIVE-NUMBER-01`), I1
> read-model, F4 parent guard (I2), MasterSync detekcija i MIG-005b (picker)
> **jesu** implementirani.
>
> **Status (`v6-ui-224`):** §§1–3 opisuju **zatečeni kod** (svaka tvrdnja nosi izvor
> sa brojem linije). §§4–6 i I1/I2 iz §8 su **implementirani**. §9 kaže koji su
> acceptance testovi napisani, a koji čekaju i zašto.
>
> **§2 je bio nepotpun:** tabela tri writer-a je tvrdila da sva tri zovu
> `ApplyGeneracijaID`, i to je tačno opisivalo kod — ali ne i to da je za jednog
> od njih (`modMasterSync`) nasleđivanje generacije **pogrešno**. V. §11b.
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
`AppendRow` u `tblZbirna`, i sva tri ga odmah **pečate** — ali ne istom rutinom:
prva dva generaciju **nasleđuju** u svom opsegu, `modMasterSync` je **kuje**
(v. §11b):

| Writer | AppendRow | ApplyGeneracijaID |
|---|---|---|
| `modDokumenta.SaveZbirna` | `modDokumenta.bas:633` | `:636` |
| `modMasterSync` (PWA import) | `modMasterSync.bas:3175` | `ApplyNovaGeneracijaID` — **uvek nova**, v. §11b |
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
historicalLogicalCount             Long      ' distinct GeneracijaID, SVI redovi IKAD
                                             ' MERI SE, NE BLOKIRA -- v. par.13
scopeProvided                      Boolean
historicalOwnerIsScope             Boolean   ' True samo kad je historicalOwnerCount = 1
                                             ' i taj vlasnik = prosledjeni scope
matchingScopeActiveLogicalCount    Long
resolutionStatus                   NONE | UNIQUE | OWNER_MISMATCH | CURRENT_AMBIGUOUS
selectedGeneracijaID               String    ' popunjeno samo kad je UNIQUE
selectedVozacID / selectedKupacID  String
brojUAktivnojPrijemnici            Boolean  ' I1, par.8
```

**Fail-closed, dodato u implementaciji:** kad `integrityStatus` nije `OK`,
`resolutionStatus` **nikad nije `NONE`** — postavlja se `CURRENT_AMBIGUOUS`.
`NONE` je jedina vrednost koja negde znači „sme se" (kapija u §5), pa se ne sme
dobiti iz greške. `integrityStatus` i dalje kaže *zašto*.

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
integrityStatus = INTEGRITY_ERROR                     -> BLOCK  (INTEGRITET)
activeLogicalCount > 0                                -> BLOCK  (AKTIVNA)
activeLogicalCount = 0 AND historicalOwnerCount = 0
        AND brojUAktivnojPrijemnici                   -> BLOCK  (SIROCE)
activeLogicalCount = 0 AND historicalOwnerCount = 0   -> ALLOW
activeLogicalCount = 0 AND historicalOwnerIsScope     -> ALLOW  (ispravka / re-entry)
inace                                                 -> BLOCK  (TUDJ)
```

Kapija vraća **kod razloga**, ne poruku: `modDokumenta` je sloj podataka i ne nosi
korisnički tekst. Prevod je jedini prelaz, u `modDokUnos.ZbirnaGatePoruka`. Uzrok
se ne stapa u jednu poruku — zauzet sada, zauzet ikad, drži ga prijemnica i
pokvaren podatak su četiri različita poteza za operatera.

**`SIROCE` grana je dodata u implementaciji** i dolazi iz I1: broj koji drži
aktivna prijemnica a nijedna zbirna nikad nije slobodan — nova zbirna pod njim
tiho bi postala roditelj tuđe prijemnice.

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
| `UNIQUE` + `historicalOwnerCount > 1` | **hard block** |
| `UNIQUE` + `historicalOwnerCount <= 1` | prolaz, `selectedGeneracijaID` je roditelj |

Tabela živi na **jednom mestu** — `modDokumenta.ZbirnaRoditeljRazlog`;
`ZbirnaRoditeljOK` je tanak omotač nad njom. Vraća **kod razloga**, a prevod u
tekst je u `modDokUnos.ZbirnaRoditeljPoruka` — isti obrazac kao F3 kapija.

**Zašto tvrda blokada, a ne `UPOZORENJE`.** `PRIJEMNICA_ZBIRNA_PROVERA` bira
politiku za *„zbirne nema"* — stanje koje operater može da zna unapred (zbirna tek
stiže). Dvosmislen ili tuđ dokument nije to: potvrda ne pomaže, jer ekran ne može
ni da ponudi **koji** je pravi. Zato se ta politika ne dira, a nove grane blokiraju
bezuslovno.

Kapija stoji u **`Else` grani** iza `ZbirnaPostoji` — kad zbirne nema, do nje se i
ne stiže, pa je „`NONE` → postojeće ponašanje" iz tabele doslovno tačno.

**Istorija je deo bezbednosti, ne samo sadašnje stanje.** `UNIQUE` danas uz broj
koji je *ikad* držalo više vlasnika i dalje nije bezbedan roditelj: prijemnica
čuva **samo `BrojZbirne`**, pa svaka nizvodna operacija po broju može da zahvati i
tuđe. `modStorno.RequireJedanVlasnikIkadPoBroju` postoji baš zbog toga. Uslov pada
tek kad prijemnica dobije pravi FK na generaciju (`Prijemnica.ZbirnaGeneracijaID`).

> **Ispravka regresije.** Ugovor v3 je ovaj red imao; v4 ga je ispustio pri
> usvajanju stroge F3 matrice, bez oznake. `ZbirnaRoditeljOK` je verno
> implementirao v4 — kod nije odlutao od ugovora, ugovor je odlutao od sebe.
> Vraćeno pre nego što je I2 dobio ijednog pozivaoca.

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

**Pokriveno** (`modTest`, `RunAllTests`): A1–A5, A7, A8, A13, A15–A18, A20 —
testovi `T_ZbirnaIdent_BrojSeRazresavaUDokument`,
`T_ZbirnaKapija_AktivanBrojNeSmeDvaput` i
`T_Prijemnica_VezujeSeSamoNaJednoznacnu`, plus **devet** sabotaža.

**Pokriveno** (`modBusinessFlowProTests`, `RunBusinessFlowProSuite`): A21 —
`Test_ZBR_ImportDvaUredjajaNeStapaDokumente` ide kroz **pravi** uvoz
(`TestHook_ImportZbirnaRowPWA` → `ImportRowToTblZbirna`), plus sabotaža
`mastersync-nasledjuje-tudju-generaciju`.

**Čeka:** A6, A9–A12, A14 — sve traže **upis** (`Scr_Save`,
`SaveZbirnaMulti_TX`), pa idu u BFP suite, ne u `RunAllTests`.
A17 više ne čeka: par „dva dokumenta istog vlasnika" fixture nema, ali ga test
pravi sam (privremeno izjednači vozača para) i vraća.

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
| A21 | import: dva `ClientRecordID`-a, **isti** vozač, kupac i broj | Ne prolazi kroz `ZbirnaValidiraj` (fiksira granicu kapije iz §5), oba reda opstaju, ali dobijaju **različite** `GeneracijaID` → `2 / 1`, `CURRENT_AMBIGUOUS`, F4 blokira, B8 prijavljuje. |

A19 je **povučen** (bio je legacy `LogicalZbirnaKey` fallback) — v. D4. Broj se
ne reciklira, da se stariji zapisi ne bi pogrešno čitali.

## 10) Van opsega, imenovano

- **`ZBR-NORM-02`** — **urađeno** (`v6-ui-226`, §14): kapije nad poslovnim brojem
  porede po **jedinstvenoj semantici** (`Trim` + `vbTextCompare`).
- **MIG-005b** — **urađen** (`v6-ui-223`): `FillZbirneCombo` de-duplikuje po
  `GeneracijaID`, pa dvoklasna zbirna daje jednu stavku. Dva **različita**
  dokumenta pod istim brojem i dalje stoje dvaput — namerno; F4 takav broj odbija
  (`CURRENT_AMBIGUOUS`), a B8 ga prijavljuje. Test
  `T_Zbirne_PickerJednaStavkaPoDokumentu`. V. `UI_MIGRACIJA_KATALOG.md` §28.1f.
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

---

## 11b) MasterSync — ingest, pa detekcija

**Ista provera, druga posledica.** F3 i F4 su **komande** operatera: konflikt tamo
znači „ne radi to", pa se unos odbija. PWA import je **ingest već nastale
činjenice sa terena**, i tu odbijanje nije simetrično.

| | F3 / F4 | MasterSync |
|---|---|---|
| Šta je unos | namera operatera | činjenica sa terena |
| Ko može da izabere drugi broj | operater, odmah | niko |
| Cena odbijanja | operater kucne drugi broj | `Err.Raise` u `tx` → **rollback celog sync reda** |

`ImportZbirnaRow_TX` na prazan povratak diže grešku unutar transakcije, pa bi
blokada **izgubila podatak** — a koliziju ostavila neprijavljenu. `KNOWN_ISSUES`
**KR-001** multi-device koliziju `BrojZbirne` već *prihvata* kao rizik, sa
mitigacijom na GAS strani.

Zato `ImportRowToTblZbirna` red **upisuje**, pa zove
`PrijaviKolizijuBrojaZbirne` — koja meri **stanje koje je import ostavio** (zove
se posle pečaćenja identiteta, ne pre) i piše `LogWarn`. Ta procedura **nikad ne
diže grešku**: pad detekcije unutar transakcije oborio bi baš onaj upis koji
treba da sačuva.

Trajni trag je u `modIntegritet`: **B8** (broj nosi više aktivnih dokumenata) i
**B9** (aktivna zbirna bez `GeneracijaID`). Log se izgubi, nalaz ostaje.

### Zašto MasterSync **uvek kuje** novu generaciju (`v6-ui-224`)

Do `v6-ui-224` je import zvao `ApplyGeneracijaID`, koji generaciju **nasleđuje**
od aktivnog reda istog broja u istom opsegu (`VozacID + KupacID`). To je tačno za
`SaveZbirnaMulti_TX` i `modDokumentInvariant` — oni **jedan** dokument pišu u
**dva** reda (Kl. I i Kl. II), pa drugi red mora u generaciju prvog.

Za import nije. PWA obe klase sabira u **jedan** red (`Klasa = "I/II"`), a
`IsDuplicateZbirnaInMaster` odbija već uvezen `ClientRecordID` **pre** upisa —
svaki red koji stigne do `AppendRow` je dokument koji nikad nije viđen. Isti
vozač i isti kupac **ne znače** isti dokument; to je tačno **A17**, i to je baš
`KR-001` scenario: dva uređaja offline dodele isti broj istom vozaču i kupcu.

Nasleđivanje je zato uništavalo činjenicu koju detekcija treba da vidi — i to
**pre** nego što bi je iko izmerio:

| Posledica | Zašto |
|---|---|
| `PrijaviKolizijuBrojaZbirne` ćuti | `activeLogicalCount` broji **generacije**; stopljene daju 1 |
| **B8** nema nalaz | presudu uzima od istog resolvera |
| **F4** pušta prijemnicu | `resolutionStatus = UNIQUE`, `historicalOwnerCount = 1` |
| **jedan storno obori oba dokumenta** | `StornoZbirna` redove bira po generaciji (`RedJeIzabranogDokumenta`), a kapiju `RequireJedanVlasnikPoBroju` **preskače** kad je generacija zadata |
| operater ih ne razlikuje ni u listi | skrivena kolona identiteta u storno gridu je baš `COL_GENERACIJA_ID` (`modScrDokumenti.IdKolonaTipa`) |

Poslednja dva reda su teža od prva tri: nije reč o propuštenoj detekciji nego o
**pogrešnoj identifikaciji na svakoj nizvodnoj radnji po identitetu**.

Pravilo je zato: pisac koji **jedan dokument deli na više redova** nasleđuje
(`ApplyGeneracijaID`); pisac koji **svaki red piše kao zaseban dokument** kuje
(`ApplyNovaGeneracijaID`). Bezbednost ne dolazi od toga što je I2 ušao ranije,
nego od toga što dve terenske činjenice ne postaju jedan identitet — tek onda broj
**stvarno** postane dvosmislen, pa ga F4 fail-closed odbija.

## 12) Fixture i ZBR-IDENT-01

**Do `v6-ui-221` fixture nije poštovao invarijantu.** `tools/make_fixture.py` nije
upisivao `GeneracijaID` ni na jedan red, a tri reda (`ZBI-TEST-1`, `ZBI-TEST-2`,
`ZBI-TEST-STOR`) nisu imala ni `KupacID`. To je stanje koje produkcija **ne može
da proizvede**: `ValidateZbirnaInput` odbija zbirnu bez kupca, a sva tri writer-a
odmah pečate validan `GeneracijaID`.

Posledica je bila da se **svaki postojeći broj čita kao `INTEGRITY_ERROR`**, pa je
I2 bio blokiran — `T_BrutoNeto_PoRezimu` tvrdi `AssertEq resP, ""` nad
`FX_ZBIRNA`, i tvrda blokada u `PrijemnicaValidiraj` oborila bi ga bez obzira gde
se postavi.

**Sada svih 24 reda `tblZbirna` nose `GeneracijaID` i `KupacID`.**

| Odluka | Zašto |
|---|---|
| Format `GEN-00000` | `GetNextID` parsira **numerički** sufiks posle prefiksa; nenumerička generacija ostavlja `maxNum` pogrešan, pa bi sledeći upis kovao već zauzetu vrednost |
| Jedna generacija po redu | u ovom fixture-u nijedan par ne deli `broj + vozač + kupac`, pa dvoklasne zbirne nema. **A6** se zato meri kroz `SaveZbirnaMulti_TX` u BFP suite-i |
| Anomalije se prave **u testu** | **A17** (dva aktivna dokumenta istog vlasnika) i **A20** (aktivan red bez generacije) su *fault injection* koji `modTest` postavlja i vraća — fixture ostaje validno produkciono stanje |
| Testovi više **ne pečate** generacije | pečatiranje je bilo zaobilaženje pogrešnog fixture-a; sada se očekivana vrednost **čita iz reda**, pa tvrdnja ne zavisi ni od rednog broja reda u `make_fixture.py` |

### Granica: invarijanta je ZBR, ne „svaka dokument-tabela"

Širenje na `tblOtpremnica` bi **odmah** oborilo
`T_ZavrsetakIspravke_NeDegradiraOldDocID`, koji kao preduslov tvrdi:

```vb
AssertEq modDokumenta.GeneracijaPoID(TBL_OTPREMNICA, COL_OTP_ID, "OTP-LEG-A"), "", _
         "preduslov: zatecen dokument NEMA generaciju"
```

Taj red namerno nema generaciju, jer test meri **degradaciju na poslovni broj** kad
je nema. `ZBR-IDENT-01` se zato drži `tblZbirna`; ostale tabele su zaseban razgovor.

### Kolona ne postoji u donoru — zato `ENSURE_COLS`

`GeneracijaID` **nije deo šeme sveske**: pravi je `modSetup.EnsureSledljivostSchema`
na svakom startu aplikacije, u radnoj kopiji. Donor je zato nema, pa je prvi
pokušaj sejanja pao sa `SEMA: tblZbirna: donor nema kolone ['GeneracijaID']`.

Generator za to već ima mehanizam — `ENSURE_COLS` dograđuje kolone koje donor
nema, *pre* sejanja, isto što `EnsureColumnOnTable` radi na startu. Dodat je
`"tblZbirna": ["GeneracijaID"]`.

Namerno **samo `tblZbirna`**: `EnsureSledljivostSchema` je dodaje na šest tabela,
ali invarijanta se drži zbirne — `tblOtpremnica` namerno zadržava `OTP-LEG-A` bez
generacije (v. granicu iznad).

### Regeneracija

Fixture je artefakt, ne repo sadržaj — posle izmene `make_fixture.py` mora se
ponovo napraviti, inače `run_vba` staje na proveri ustajalosti:

```bash
python tools/make_fixture.py --donor "<put/do/donora.xlsm>" --force
```

## 13) `ZBR-MUT-01` — vlasnička i dokumentna dvosmislenost nisu isto (`v6-ui-225`)

Deca zbirne — otpremnica, prijemnica, paletna stavka, denormalizovan otkup —
nose **samo `BrojZbirne`**. Generacije nemaju. Zato svaka rutina koja decu bira
po broju zahvata **sve** dokumente tog broja, ma koliko ih bilo.

Kapija za to je postojala (`modStornoFlow.ZbirnaBrojJeDvosmislenIkad`, šest
poziva) i merila je **vlasnike**:

```vb
VlasniciPoBroju(...).count > 1
```

**Komentar uz svaki od tih šest poziva opisivao je dokumentnu dvosmislenost**
(*„dva aktivna dokumenta istog broja delila bi otpremnice, pa bi se odvezale i
tuđe"*), a mera je bila vlasnička. Dva pojma se poklapaju samo dok jedan vlasnik
**ne može** da ima dva dokumenta pod istim brojem — a `v6-ui-224` je baš to
učinio dostižnim (`KR-001`, dva uređaja offline; A17 oblik `2 / 1`).

| | Šta meri | Kad je opasno |
|---|---|---|
| `historicalOwnerCount > 1` | broj je **ikad** prešao granicu vlasništva | storniran vlasnik i dalje ima aktivnu decu |
| `activeLogicalCount > 1` | broj **sada** nosi više dokumenata | samo aktivni konkurišu za decu |

Kapija je sada `modDokumenta.ZbirnaMutacijaPoBrojuRazlog` i blokira na **oba**,
sa različitim razlogom (`ZBR_MUT_VISE_VLASNIKA` / `ZBR_MUT_VISE_DOKUMENATA` /
`ZBR_MUT_INTEGRITET`) — uzrok se ne stapa, jer su to tri različita poteza za
operatera.

### Šta je propuštalo, po putanji

| Putanja | Zaglavlje | Deca |
|---|---|---|
| `StornoZbirnaIDetach_TX` (SIMPLE) | tačno, po generaciji | `DetachOtpremniceInline` **prazni `BrojZbirne` deci OBA dokumenta** |
| `PonistiZbirnaChain_TX` | `gen` je bio **mrtav parametar** — prosleđivan i nikad korišćen | otpremnice/prijemnice skupljane po golom broju |
| `Run/CompleteZbirnaCorrection` | — | relink i rekalkulacija po broju zahvataju oba |
| `RunOtpremnicaCorrection` (roditelj) | — | ista kapija nad roditeljskom zbirnom |

### Zašto `historicalLogicalCount` **ne** blokira

Meri se i stoji u DTO-u, ali nije uslov. Razlog je unutar ovog istog ugovora:
`ZbirnaNovUnosRazlog` ima ALLOW granu

```
activeLogicalCount = 0 AND historicalOwnerIsScope  ->  ALLOW  (ispravka / re-entry)
```

a `GeneracijaIDZaBrojArr` isključuje stornirane — pa **svaki redovan re-entry kuje
novu generaciju**. Ispravljena zbirna pod istim brojem zato stoji kao
`historicalLogicalCount = 2, historicalOwnerCount = 1`. Blokada na toj vrednosti
bi značila da F3 kaže „smeš ponovo pod ovim brojem", a storno „ne smeš ga više
dirati" — kontradikcija u istom ugovoru. `modTest` test 34
(`T_IstiBrojRazliciteGeneracije_NijeIstiDokument`) to stanje tvrdi kao legitimno.

Uz to je opasnost uža nego što izgleda: `DetachOtpremniceInline` **prazni broj**
na deci, pa posle prostog storna deca stornirane generacije taj broj više i ne
nose — za mutaciju po broju ne konkurišu. Stanje „stornirana generacija sa
aktivnom decom pod istim brojem" ne nastaje redovnim putem, nego ručnom izmenom
u tabeli; tada ga hvata `integrityStatus` / B9, ne ova kapija.

**Trajno rešenje** je da deca nose `ZbirnaGeneracijaID`. Dok ga nemaju, kapija je
jedino što stoji između mutacije po labeli i tuđeg dokumenta.

### Verifikacija

| Šta | Gde |
|---|---|
| `2 aktivna dokumenta / 1 vlasnik` → SIMPLE i ISPRAVKA staju, deca ostaju vezana | `modBusinessFlowProTests` `Test_ZBR_MutacijaPoBrojuStajeNaDvaDokumenta` (stanje pravi **pravi uvoz**, dva `ClientRecordID`-a) |
| Dvoklasna zbirna (dva reda, **jedna** generacija) se i dalje stornira | ista, negativna kontrola |
| Fail-closed na sopstvenu grešku | `modTest` `T_KapijaZbirne_FailClosedNaSvojuGresku` (schema drift) |
| Sabotaža | `kapija-mutacije-broji-samo-vlasnike` |

**Nije zasebno mereno:** prosleđivanje `gen` u `StornoZbirna` iz
`PonistiZbirnaChain_TX`. Kapija iznad više ne pušta dva aktivna dokumenta, a
`StornoZbirna` preskače već stornirane redove — pa kroz podržane putanje razlike
u ponašanju nema. Ispravka je precizna, ne merljiva; sabotaža za nju bi bila
zelena bez obzira na kod, pa nije ni dodata.

## 14) `ZBR-NORM-02` — jedinstvena semantika poređenja broja (`v6-ui-226`)

Poslovni broj je **labela**: dolazi iz Excel ćelije (razmaci) ili iz ručnog unosa
(case). Isti ključ je imao **tri** normalizacije:

| Nivo | Gde | Šta radi |
|---|---|---|
| pun | `ZbirnaPostoji`, `BrojZbirneExists`, `ZbirnaIdentResolve`, `AktivniBrojeviZbirne` | `Trim` + `vbTextCompare` |
| samo `Trim` | `VlasniciPoBroju`, `LookupActiveID`, `DistinctActiveValues` | case-sensitive |
| sirov | `CheckDuplicate` | ni trim ni case — imenovano u §3 |

Dok je kapija za mutaciju bila vlasnička (`VlasniciPoBroju`), kapija i akter su
računali **isto** — oba na drugom nivou. `ZBR-MUT-01` je kapiju prebacio na
resolver (prvi nivo), pa su se razišli.

**Pravilo: kapija sme da bude ŠIRA od aktera, nikad uža.** Šira kapija
preblokira — glasno i bezbedno; uža bi pustila radnju koja zahvata više nego što
je izmereno. Zato `BrojJednak` koriste **odlučivači**, a mutatori
(`DetachOtpremniceInline`, `RelinkOtpremniceToZbirna_TX`,
`RedJeIzabranogDokumenta`, `ActiveOtpIDsByZbirna`, `ActivePrijIDsByZbirna`)
namerno ostaju uži: proširiti njih značilo bi dirati redove koje danas ne diraju,
a to je izmena ponašanja koja traži svoj dokaz po putanji.

**Asimetrija u `DistinctActiveValues`** je usput ispravljena: ćelija je bila
trimovana, `filterVal` nije, pa bi netrimovan pozivalac tiho dobio prazan skup —
u `CompleteZbirnaIspravka` to znači „nema paleta za prevezivanje". Sva tri
zatečena pozivaoca šalju trimovanu vrednost (`GetCorrectionField` trimuje), pa
**kvar nije bio živ** — ali jeste bio zamka.

**Verifikacija:** `modTest` `T_BrojKapija_IstoZaSvakiCase` meri sva tri
odlučivača, bez ijednog upisa, nad fixture brojem `ZB-TEST-KASK` koji nosi slova.
Preduslovi tvrde da tačan case daje ne-nula rezultat, inače bi „isto kao tačan
case" bilo zeleno i kad obe grane vrate nulu. Sabotaže: `vlasnici-poredi-case`,
`lookup-aktivnog-poredi-case`, `deca-po-broju-poredi-case` — po jedna na svaki
odlučivač, da se ne može desiti da dva budu prebačena a treći ostane star.

**`BrojJednak` nije jedini komparator, i ne treba da bude.** `ZbirnaPostoji`,
`BrojZbirneExists`, `ZbirnaIdentResolve` i `AktivniBrojeviZbirne` zadržavaju
svoja inline poređenja — **ista semantika**, drugi zapis. Prebacivati i njih samo
radi jednog izvora istine znači dirati stabilan kod bez poslovne koristi i sa
istim blast radiusom kao prava izmena. Ono što je popravljeno su odlučivači koji
su radili **drugačije**, ne oni koji su radili isto na svoj način.

**Ostaje otvoreno:** `CheckDuplicate` (§3) i mutatori. Trajno rešenje za oboje je
isto kao za `ZBR-MUT-01` — `ZbirnaGeneracijaID` na deci, pa poređenje po broju
prestane da bude identitetsko pitanje.

## 15) `ZBR-CHILD-01` — generacija roditelja na detetu (`v6-ui-227`, faze 1–2)

Trajno rešenje za ono što `ZBR-MUT-01` (§13) drži kapijom, i za ono što
`ZBR-NORM-02` (§14) drži pravilom „kapija ⊇ akter". Oboje postoji **samo zato što
deca zbirnu nose kao labelu**, ne kao identitet.

`ZbirnaGeneracijaID` na `tblOtpremnica`, `tblPrijemnica`, `tblPaletaStavka` i
`tblOtkup` (denorm) je ta veza.

### Zašto se ne peča pri upisu deteta

Prva pretpostavka — „dete pri nastanku zna roditelja" — **ne važi**.
`modAutoHladnjaca.bas:213` snima otpremnicu, a zbirnu tek na `:221`. U malina i
hladnjača lancu **dete redovno nastaje pre roditelja**.

Zato invarijanta nije „uvek popunjeno", nego:

> `ZbirnaGeneracijaID` na detetu je **prazan**, ili jednak generaciji zbirne kojoj
> dete pripada — i menja se **u koraku sa `BrojZbirne`**, uključujući brisanje.

**Prazno je legitimno** i znači „roditelj još nije razrešen". Čitalac tada pada na
broj, tačno kao pre ove kolone. To je ono što čini postepenu migraciju mogućom.

### Pravilo: nikad ne pogađaj kad već znaš

Identitet roditelja se uzima **najbližim poznatim putem**, a razrešavanje po broju
je poslednja opcija — ne prva:

| Situacija | Odakle generacija |
|---|---|
| Neposredni roditelj je nosi | **kopira se od njega** (paleta ← prijemnica) |
| Poznat je konkretan `ZbirnaID` | čita se **iz tog reda** (`GeneracijaPoID`) |
| Nema nijednog kanonskog identiteta | tek tada `ZbirnaGeneracijaZaBroj`, fail-closed |

Prva verzija ovog koraka je to prekršila na tri mesta, i sva tri su bila
**tiho pogrešna**:

- `LinkZbirnaToOtkupAndOtpremnica` je imao `ZbirnaID` i bacao ga da bi pitao
  labelu. U `KR-001` koliziji (dva aktivna dokumenta pod istim brojem)
  razrešavanje po broju vrati **prazno** — dakle veza bi izostala baš tamo gde je
  najpotrebnija.
- `AddStavka` je imao `PrijemnicaID` i pitao globalno „koja je zbirna **sada** pod
  ovim brojem". Posle storna + re-entry prijemnica ostaje na `GEN-A`, a njena
  paleta bi dobila `GEN-B` — razbijena sledljivost unutar jednog lanca.
- `modAutoHladnjaca` je imao upravo kreiran `ZbirnaID` i nije završio vezu, pa je
  otpremnica **trajno** ostajala prazna (v. sledeći odeljak).

### Jedan put, u oba smera

| | |
|---|---|
| `PoveziDeteNaZbirnu` | upisuje broj **i** generaciju, u istom potezu |
| `OdveziDeteOdZbirne` | briše oboje |
| `ZbirnaGeneracijaZaBroj` | broj → generacija, **fail-closed**: prazno za sve što nije `UNIQUE` |

Dva odvojena upisa bi se pre ili kasnije razišla — neko doda putanju koja
postavlja broj a zaboravi generaciju, i dete ostane sa **tuđom** generacijom.
To je gore od prazne: prazna bar znači „ne znam".

`ZbirnaGeneracijaZaBroj` se zove **jednom po broju, ne po redu** —
`ZbirnaIdentResolve` čita celu `tblZbirna`, pa bi poziv u petlji nad decom bio
O(n·m). Petlje uzimaju generaciju jednom i prosleđuju je.

### Šta je urađeno, šta nije

**Faza 1** — kolona (`EnsureSledljivostSchema`), choke point, i **svih 16
produkcionih pisaca** kroz njega. **Nijedan čitalac nije diran.**

**Faza 2** — `modSetup.BackfillDeteZbirnaGeneracija`: jednokratno, idempotentno
(samo prazni redovi), van `EnsureRuntimeSchema` jer je skupo po startu.

Telo je u **`BackfillDeteZbirnaGeneracija_Core(showMessages, popunjeno, preskoceno)`**;
javna procedura je samo operaterski ulaz sa `MsgBox`-om. Seam nije kozmetika —
`MsgBox` u automatskoj suite visi, pa je backfill bez njega bio **nepozvan ni iz
jednog testa**. To se videlo tek dvosmernim dokazom: sabotaža koja mu je menjala
kriterijum izbora nije obarala ništa, jer je menjala red koda koji se ne izvršava.
**Pokrivena primitiva nije pokriven pozivalac** — `ZbirnaJedinaGeneracijaIkadZaBroj`
je imala svoju tvrdnju, a jedini pisac koji je zove nije imao nijednu.

**Kriterijum je ISTORIJSKI, ne tekući** — i to je razlika koja čuva sledljivost.
Backfill zove `ZbirnaJedinaGeneracijaIkadZaBroj`, ne `ZbirnaGeneracijaZaBroj`:

```
GEN-A | ZB-10 | vlasnik X | STORNIRANO
GEN-B | ZB-10 | vlasnik X | AKTIVNO      <- resolver kaže UNIQUE = GEN-B
OTP-A | BrojZbirne = ZB-10 | generacija prazna
```

To stanje §5 **izričito dozvoljava** (re-entry istog vlasnika posle storna).
`OTP-A` je istorijski dete `GEN-A`; „sada" bi mu upisalo `GEN-B` i napravilo
**lažnu sledljivost** — gore od prazne kolone, jer prazna bar ne tvrdi ništa.

Popunjava se samo broj koji je **ikad** nosio jednu generaciju. Stornirana jedina
generacija se sme upisati: ako je pod tim brojem ikad postojala samo jedna,
identitet je poznat bez obzira na današnje stanje.

### Auto-lanac: veza se završava, ne ostavlja

U auto-lancu otpremnica nastaje **pre** zbirne, pa joj je generacija tada prazna —
tačno u tom trenutku. Ali bez dopune ostala bi tako **zauvek**, i faza 3 („koristi
generaciju kad je nose svi redovi") nad novim podacima nikad ne bi postala tačna
bez ručnog backfill-a. `ZavrsiVezuOtpremniceNaZbirnu` zato završava vezu odmah po
nastanku zbirne, čitajući generaciju **iz njenog PK-a**.

**Faza 3 (`v6-ui-228`)** — pet odlučivača (`DetachOtpremniceInline`,
`ActiveOtpIDsByZbirna`, `ActivePrijIDsByZbirna`, `RelinkOtpremniceToZbirna_TX`,
`DistinctActiveValues`) više ne dira svu decu pod brojem, nego samo svoju.

Čvor je `modDokumenta.SuziDecuNaGeneraciju` — pandan write choke point-u:

```
trazena generacija prazna          -> kandidati nepromenjeni
bilo koji kandidat bez generacije  -> kandidati nepromenjeni
inace                              -> samo oni koji se poklapaju
```

**Sve-ili-ništa, ne hibrid.** „Poklapa se ili je prazno" bi isti prazan red
ubacilo u skup **oba** dokumenta pod tim brojem — dupli detach, pogrešan račun.

**Izbor po broju namerno NIJE preuzet.** Četiri odlučivača porede broj tačno
(`Trim$(CStr(..)) = broj`), peti kanonski (`BrojJednak`). Da je čvor preuzeo i to,
ona četiri bi se **tiho proširila** — akter širi od zatečenog je smer koji §14
zove opasnim. Nesklad ostaje zatečen i imenovan, nije usput „popravljen".

#### Šta je faza 3 zapravo zatvorila

Nije bila samo priprema. Ovo stanje postoji **danas**:

```
A: Z-10, vlasnik X, STORNIRANA, OTP-A jos AKTIVNA
B: Z-10, vlasnik X, AKTIVNA
```

Creation path je `modStornoDok` `STIP_ZBIRNA` → `modStorno.StornoZbirna_TX`, koji
snapshot-uje **samo `tblZbirna`** i stornira zaglavlje — decu ne dira.

Kapija `ZBR-MUT-01` to ne zaustavlja: istorijska grana broji **vlasnike**
(`ikadVl` po `ZbirnaVlasnikKljuc`), pa re-entry istog vozača i kupca daje 1, a
aktivnih dokumenata je takođe 1. Kapija pušta, a `Detach` po broju odvezuje i decu
`A`. Sužavanje po generaciji to zatvara strukturno.

**Zaostatak, imenovan:** dok backfill ne prođe, fallback grana i dalje nosi tu
rupu. Zatvara je faza 4, prelaskom kapije sa `historicalOwnerCount` na
`historicalLogicalCount`.

#### Granica „sve-ili-ništa" je OPERACIJA, ne tabela

`SuziDecuNaGeneraciju` odlučuje nad **jednim** skupom, a poslovna mutacija dira
više tabela. Kad svaka odlučuje sama, jedna kaskada zna da bude pola scoped a
pola po broju:

```
ZB-X / GEN-A:  OTP-A -> GEN-A      PRJ-A -> ""      <- legacy
ZB-X / GEN-B:  OTP-B -> GEN-B      PRJ-B -> GEN-B

ponisti GEN-B:
  otpremnice  svi popunjeni -> suzi -> OTP-A prezivi
  prijemnice  jedan prazan  -> broj  -> PRJ-A STORNIRANA
```

Dokument `GEN-A` završi **polovično poništen**, što je gore od oba čista režima.

Zato `SvaAktivnaDecaNoseGeneraciju` računa odluku **jednom**, nad svim tabelama
koje ta operacija bira po broju, pa se svim selektorima prosledi isti režim:

| operacija | tabele u odluci |
|---|---|
| `DetachOtpremniceInline` | `tblOtpremnica` + `tblOtkup` |
| `PonistiZbirnaChain_TX` | `tblOtpremnica` (+ `tblPrijemnica`, `tblPaletaStavka` kad `ownsChain`) |
| `CompleteZbirnaIspravka` | `tblOtpremnica` + `tblOtkup` + `tblPrijemnica` |

Odlučivač poredi kroz `BrojJednak` iako četiri pozivaoca porede tačno. `BrojJednak`
je širi, pa je njegov skup kandidata **nadskup** stvarnog — ako svi u nadskupu nose
generaciju, nosi je i svaki podskup. Greška ide samo u stranu „ne sužavaj".

#### Ispravka uzima identitet, ne pogađa ga

Lifecycle je: context sa `OldDocID` → `StornoZbirna_TX` → operater snimi novu →
`CompleteZbirnaIspravka` → relink. **U trenutku relinka stara zbirna više nije
aktivna**, pa je razrešavanje po broju tu najgore što se može uraditi:

```
GEN-A  broj X  STORNIRANA   <- dokument koji se ispravlja
GEN-B  broj X  AKTIVNA      <- nastao u medjuvremenu

ZbirnaGeneracijaZaBroj("X")  ->  GEN-B
```

Relink bi tada **precizno izabrao pogrešan dokument**: prevezao bi tuđu decu, a
svoju ostavio. To je gore od stanja pre faze 3, gde je prevozio obe.

`GeneracijaPoID(TBL_ZBIRNA, COL_ZBR_ID, oldDocID)` radi nad PK-om, kome storno ne
smeta. Kanonski ID je sačuvan **pre** storna — isto pravilo kao u §15: *nikad ne
pogađaj kad već znaš*.

#### Jedan ishod koji se menja

`RelinkOtpremniceToZbirna_TX` sada sužava **izvor**. Kad pod starim brojem aktivnu
decu ima samo *drugi* dokument, relink vrati 0, a pozivalac to (preko
`CountActive` po broju) vidi kao neuspeh i obeleži ispravku **MANUAL**. Ranije bi
prevezao tuđu decu bez reči. Fail-closed umesto tihe štete, ali jeste nova
`MANUAL` tamo gde je ranije „prolazilo".

**Faza 4 (ne u ovom koraku)** — `ZbirnaMutacijaPoBrojuRazlog` prestaje da blokira
`activeLogicalCount > 1` kad sva deca tog broja nose generaciju. Tek tada je
`ZBR-MUT-01` rešen strukturno, a ne kapijom.

**Korist stiže u fazi 4.** Faze 1–2 su trošak bez vidljive promene — to je
svesno plaćeno da bi koraci bili odvojivo dokazivi.

### Kapija mora da gleda isto što pisac piše

Uvođenje kolone je **oslabilo jednu zatečenu kapiju, a da je niko nije dirao.**

`modMasterSync.RequireBrojZbirneNotConflicting` je gledao samo `BrojZbirne`.
Dok je `PoveziDeteNaZbirnu` pisao samo broj, upis pod **istim** brojem bio je
idempotentan — ista vrednost preko sebe. Otkad pisac piše i generaciju, isti taj
put menja **roditelja** deteta:

```
Zbirna A: Broj = ZB-10, Gen = GEN-A     <- dete je već ovde
Zbirna B: Broj = ZB-10, Gen = GEN-B     <- drugi uređaj, isti broj (§5 dozvoljava)

kapija:   "ZB-10" == "ZB-10"  -> prolazi
pisac:    GEN-A -> GEN-B      -> tiho premešten vlasnik
```

To je `ZBR-MUT-01` naopako: **kapija (broj) uža od aktera (broj + generacija)**.
Regresiju je uveo upis, ne kapija — što je i razlog da se pravilo formuliše kao
*„kapija i pisac gledaju isti ključ"*, a ne kao spisak provera.

Guard je zato `RequireZbirnaVezaNotConflicting`, sa matricom:

| postojeći broj | postojeća gen. | novo (broj/gen.) | ishod |
|---|---|---|---|
| prazan | prazna | `X` / `GEN-A` | ALLOW |
| `X` | prazna | `X` / `GEN-A` | ALLOW — završava nerazrešenu vezu |
| `X` | `GEN-A` | `X` / `GEN-A` | ALLOW — idempotentno |
| `X` | `GEN-A` | `X` / `GEN-B` | **BLOCK** |
| `X` | `GEN-A` | `X` / prazna | **BLOCK** — znanje se ne briše |
| `X` | bilo šta | `Y` / bilo šta | **BLOCK** |
| prazan | `GEN-A` | bilo šta | **BLOCK** — integritet |

**Prepisivanje roditelja postoji**, ali kroz ispravku i prevez, koji su
operaterske komande. Zato zabrana **nije** u `PoveziDeteNaZbirnu`: choke point
mora da ostane upotrebljiv za te putanje. Ingest zatečene činjenice nije mesto
za promenu vlasništva dokumenta — ista podela komanda/ingest kao u §13.

### Verifikacija

| Šta | Gde |
|---|---|
| Kolona postoji na sve četiri tabele, i nije ista kao generacija samog dokumenta | `modTest` `T_DeteZbirne_ImaKolonuGeneracije` (bez upisa) |
| Roditelj jednoznačan → dete nosi njegovu generaciju; roditelja nema → **prazno**; odvezivanje briše oboje; **storniran roditelj → prazno** | `modBusinessFlowProTests` `Test_ZBR_DeteNosiGeneracijuRoditelja` |
| Sabotaže | `dete-ne-nosi-generaciju-roditelja`, `odvez-ostavlja-generaciju`, `dete-pogadja-generaciju-po-broju` |

Grana sa **storniranim** roditeljem postoji zato što razdvaja *razrešavanje* od
*pogađanja*: kad zbirne uopšte nema, i naivni `LookupValue` po broju vrati prazno,
pa bi sabotaža koja uvodi pogađanje prošla neprimećeno. Sa storniranim roditeljem
pogađanje vraća njegovu generaciju, a tačan odgovor je prazno.
