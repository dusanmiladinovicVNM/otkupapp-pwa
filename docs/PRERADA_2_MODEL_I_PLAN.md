# Prerada 2.0 — proizvodno jezgro: rafinisan model i plan

- **Status:** predlog za odluku (nije implementirano; ništa od ovoga još nije u kodu)
- **Datum:** 2026-09-02
- **Grana:** `claude/prerada-2-production-core-stfebv`
- **Ulaz:** predlog „N ulaznih lager jedinica → procesna šarža → M izlaznih" (59 tačaka),
  stanje `main` posle #246, otvoreni PR-ovi
  [#247](https://github.com/dusanmiladinovicVNM/otkupapp-pwa/pull/247) i
  [#248](https://github.com/dusanmiladinovicVNM/otkupapp-pwa/pull/248).
- **Vezano:** `docs/DOMEN/README.md`, `docs/adr/0001-*.md`, `docs/adr/0002-*.md`,
  `docs/AMBALAZA_MODEL.md`, `docs/UI_MIGRACIJA_KATALOG.md` §25, `docs/Master Plan/09_QA_DECISION_LOG.md`
  (odluke 66, 67).

> Cilj dokumenta: da predlog sroči u oblik koji **ova** baza koda može da nosi bez
> big-banga — isti obrasci (writer `*_TX` + snapshot, `Stornirano` registar, saldo koji
> se izvodi pri čitanju, ekran u ljusci, test + sabotaža), isti dokumentacioni izvori
> istine. Gde se od predloga odstupa, piše zašto.

---

## 0) Presuda u jednoj rečenici

Predlog je **tačan u jezgru** (procesna šarža kao događaj, N→M kroz junction tabele,
klasifikovan izlaz i masa-balans kao invarijanta, tip procesa kao matični podatak,
redosled procesa koji se ne hardkoduje, storno sa nizvodnom kapijom, dva koraka
OTVORENA→ZAVRŠENA) i **preširok u obimu** za jedan zahvat (komore, pozicije, tuneli,
senzori, smene, radnici, HACCP, specifikacije, kalibracije, lotovi, rezervacije).
Ovaj plan zadržava jezgro u potpunosti, a periferiju raspoređuje u faze koje se
oslanjaju na već donete odluke (Master Plan 66: smene/radnici/integracije su
„minimalni scope proizvodnje 2027"; 67: integracije samo za odobrene uređaje).

---

## 1) Šta zatičemo (činjenice, ne pretpostavke)

### 1.1 Legacy prerada na `main`

| Šta | Gde | Činjenica |
|---|---|---|
| Ulazna lager jedinica | `tblPaleta` + `tblPaletaStavka` | sveža paleta; header je **izvedena projekcija** iz aktivnih stavki (ADR-0002 §5); `Preradjeno` = `Da`/prazno; broj se resetuje po godini |
| Prerada | `tblPrerada` + `tblPreradaStavka`, `SavePrerada_TX` (`modPaletniList.bas:2872`) | **N celih paleta → 1 izlaz** (kutije/kese/tip GP). Paleta se troši **cela** (`Preradjeno=Da`, `:2973`); nema parcijalnog ulaza, nema više izlaza, nema tipa procesa, nema otpada/gubitka |
| Kolone `tblPrerada` | `modConfig.bas:321-347` | `PreradaID, BrojPrerade, Godina, Datum, NetoUlazKg, NetoIzlazKg, BrojKutija, BrojKesa, TezinaPaleteKg, BrutoKg, AmbalazaKg, TipKutije, TipKese, TipGotovogProizvoda, Napomena, CreatedAt, Stornirano` |
| Integritet | `modIntegritet` A5, D1, D2 | A5: `NetoUlaz = Σ stavke`, `izlaz ≤ ulaz`; D1: `Preradjeno=Da` bez aktivne stavke; D2: stavka ka nevalidnoj paleti |
| Storno | `modStorno.StornoPrerada` / `StornoPaleta` | prerada vraća `Preradjeno`; prerađena paleta se ne stornira dok se ne stornira prerada |
| Matični podaci proizvoda | `tblKulture` (VrstaVoca, SortaVoca, GajbicaPoPaleti, pragovi), `KLASA_I/II`, `tblVrstaGotovihProizvoda` (TipGotovogProizvoda, Aktivan) | proizvod je **rascepljen**: sveže = kultura + klasa; gotovo = tekstualni tip GP. Nema `tblProizvodi` |
| Objekat | `tblStanice.JeHladnjaca` | „objekat" postoji kao stanica-hladnjača; `TBL_HLADNJACA` konstanta postoji, tabela se nigde ne koristi |
| Komore / tuneli / oprema / smene / senzori / HACCP / uzorci | — | **ne postoje nigde u kodu** (ni VBA ni PWA). „Senzori koje smo planirali" i „uzorci i HACCP" su u Master Planu, ne u kodu |
| UI | `modScrPalete` lista `NOVAPRERADA` (v6-ui-159) | kvačice u mreži + polja u zoni; legacy `frmPalete` ostaje operativan (pravilo §5 `otkup-i-dokumenta.md`) |
| Sledljivost | `modIzvestaj.ReportSledljivostLanac`, `SledGpMape` | **linearan** LANAC (30 kolona posle #248); paleta→prerada join `PaletaID` preko `tblPreradaStavka`; mape „paletaID → dict(preradaID)" — struktura već trpi 1→N, ne i N→M |

### 1.2 PR #247 (`claude/izvestaji-matrix-test-fix`)

Jedan fajl, 9/5 linija: ispravka dva tvrđenja u `modIzvestajTests` posle #245.
**Nema dodira sa preradom.** Važno samo kao **baseline**: `RunIzvestajTests` je van
`RunAllTests` puta i #248 ga ne broji (163/0 je `RunAllTests`). Pre početka Faze A
oba PR-a treba da su u `main`-u da bi FULL suite bio zelen kao polazna tačka.

### 1.3 PR #248 (`claude/sledljivost-gp-lanac`, v6-ui-189 / vba-v2.91.0) — ono što Prerada 2.0 mora da poštuje

Dvadeset fajlova, 21 commit, 11 revizionih krugova. Ključne odluke koje su tamo već
donete i koje 2.0 **nasleđuje, ne preispituje**:

1. **Prerada = proizvodni lot = „paleta gotovog proizvoda".** Nema `tblGPPaleta`;
   katalog §25.8 izričito: „`GPPaletaID` denorm se može dodati kasnije bez loma".
2. **Utovar je consumption event lagera GP robe** (`tblUtovar`, `tblUtovarStavke` sa
   `PreradaID | KolicinaKg | CenaKg | BrojKutija | BrojKesa`, `tblPrevoznici`).
   Parcijalna prodaja je legalna. **Na stanju = `NetoIzlazKg − Σ aktivno utovareno`**,
   jedna funkcija za mrežu, writer i storno (`modUtovar.UtovarenoPoPreradi`).
3. **Model B:** utovar sme da postoji pre fakture (`CreateUtovarCore` +
   `CreateFakturaIzUtovara`); 1 utovar = 1 faktura; kupac utovara = kupac fakture;
   SEF `DeliveryDate = Utovar.DatumUtovara`; količinski dokaz 1:1 u oba smera.
4. **Storno lanac:** faktura → oslobađa utovar; utovar (samo nefakturisan) → vraća
   stanje; **prerada sa aktivnim utovarom se NE stornira** (fail-closed).
5. **Sledljivost stanja:** `prodato GP` / `delimicno prodato` / `utovareno, ceka
   fakturu` / `delimicno utovareno` / `preradjeno` / `prodato svezo` / `u hladnjaci`.
6. **Rok trajanja po vrsti GP** (`tblVrstaGotovihProizvoda.RokMeseci`), obrazac
   utovarne liste (LAYOUT_VER 2), `UTOVAR_PRINT_MODE`, `BackfillUtovariIzGPFaktura`
   kao **eksplicitna** (Alt+F8), ne automatska migracija.
7. Konvencije koje 2.0 kopira 1:1: `EnsureDataTable` idempotentno + self-heal u
   `EnsureRuntimeSchema` + audit kolone za nove tabele; `STORNO_TABELE` registar;
   `AuditableTables()`; pre-validacija **svih** stavki pre ijednog upisa; fixture
   `ENSURE_TABLES`; sabotaže po imenu; `who_writes --check`.

**Posledica za redosled:** #248 dira `modConfig`, `modSetup`, `modStorno`,
`modIzvestaj`, `modScrFakture`, `modTest`, `modPoruke`, `modSchemaGuard` — sve fajlove
koje i 2.0 mora da dira. Kod se za Fazu A **ne piše dok #248 nije u `main`-u**; grana
2.0 se rebase-uje na taj `main` (pravilo „Opcija 3", `git-i-release.md`).

---

## 2) Presude nad predlogom — tačka po tačka

| # predloga | Presuda | Rafinman i razlog (vezano za kod) |
|---|---|---|
| 1–3, 8 (`tblProcesSarze`, `Ulazi`, `Izlazi`, N→M) | **Prihvata se** | Tri tabele, tačno kako je predloženo, uz naše konvencije imena/ID-jeva (§3). Junction tabele nisu overengineering — `tblPreradaStavka` je već junction, samo jednosmerna |
| 4 (`tblLagerJedinice`) | **Prihvata se, sa preciznom kompatibilnošću** | LJ je **jedini ključ grafa**. Legacy paleta / legacy GP lot dobijaju LJ red **lenjo, pri prvom ulasku u proces, u istoj TX** (materijalizacija sa `IzvorTip/IzvorID`). Prodajni izlaz 2.0 dobija LJ **i** projekciju u `tblPrerada` (Faza B) da bi stek #248 (utovar/faktura/SEF/sledljivost) radio nepromenjen; projekcija se gasi u Fazi D. Odluka D1 |
| 4 (`KgTrenutno` cache) | **Odbija se** | Kućno pravilo: saldo se **izvodi pri čitanju** (ambalaža ledger, paleta header je projekcija, `UtovarenoPoPreradi`). Cache = klasa drift buga koju test hvata tek posle nastanka. `RaspolozivoKg` je jedna funkcija |
| 5 (parcijalni ulaz) | **Prihvata se** | `KgUlaz` na ulaznom redu; `Preradjeno=Da` na legacy paleti se postavlja **tek kad raspoloživo padne na 0** (čitaoci `Preradjeno` ostaju tačni: D1, `SledGpMape`, mreža paleta) |
| 9–11 (balans, tipovi izlaza, randman) | **Prihvata se** | `TipIzlaza ∈ {GLAVNI, NUSPROIZVOD, OTPAD, GUBITAK}`. `GUBITAK` je **eksplicitan red**, ne izračunat ostatak — `Σ izlaza = Σ ulaza ± tolerancija` je onda jednostavna, proverljiva tvrdnja (modIntegritet P1). Randmani su čitanje, ne kolone |
| 12–13 (tip procesa, bez hardkodovanog toka) | **Prihvata se** | `tblTipoviProcesa` sa `Sifra` kao ključem (isti obrazac kao `TipGotovogProizvoda`, `TipKutije`). Jedina kapija toka je opciona `DozvoljenaUlaznaForma` na tipu |
| 14–15 (`tblProizvodi`, klasa odvojena) | **Prihvata se, sa seed-om** | `tblProizvodi` se **puni iz postojećih** `tblKulture` (forma SVEZE) i `tblVrstaGotovihProizvoda` (`IzvorKljuc` čuva mapiranje, `RokMeseci` se ne duplira nego čita preko mapiranja). `Klasa` i `Kalibracija` su kolone LJ. `tblCenovnik` (ključ Vrsta+Sorta+Klasa) se ne dira |
| 16–17 (kalibracije, specifikacije) | **Odlaže se** (Faza F) | `Kalibracija` v1 = slobodan tekst na LJ (predlog i sam ostavlja slobodan opis). Tabela tek kad se pojavi upit nad njom |
| 18 (`tblZamrzavanje` subtype) | **Menja se** | Subtype tabela po tipu procesa = po jedna nova tabela za svaki novi tip. Umesto toga `tblProcesParametri (SarzaID, Kljuc, Vrednost, Jedinica)`; tip procesa deklariše koje ključeve traži. To je i mehanizam za §58 „različit UI, isti data model" |
| 19–22 (tuneli, senzori, temp. robe vs vazduha) | **Tunel = red u `tblOprema`** (Faza A); senzori Faza F | `TipOpreme=TUNEL`, `KapacitetKg`. Temperatura robe ulaz/izlaz = parametri šarže (ručno merenje). Senzori: odluka 67 — samo odobreni uređaji; nema integracije u v1 |
| 23, 41–42 (dokumenti) | **Prihvata se, 2 od 5 u v1** | Procesni list (generički, sa blokom parametara → pokriva i „List zamrzavanja") i Paletni list LJ. Transfer list Faza E; packing list po potrebi. Obrazac po kućnom obrascu `Ensure*Sablon` + `LAYOUT_VER` + `*_PRINT_MODE` |
| 24–28 (komore, pozicije, istorija lokacije, transfer) | **Odlaže se** (Faza E) | Nezavisan sloj od procesnog jezgra; `StanicaID` (hladnjača) je jedini „objekat" koji kod danas poznaje. Lokacija se dodaje na LJ kao istorija (ledger), nikad kao jedina `KomoraID` kolona — prihvata se princip §26 |
| 29 (`tblOprema`) | **Prihvata se** (Faza A) | Jedna tabela, `TipOpreme` + `KapacitetKg`; ne pravi se `tblTuneli` |
| 30–31 (smene, radnici, produktivnost) | **Odlaže se** (Faza F, 2027 po odluci 66) | `Kreirao/Kreirano` = postojeće audit kolone (`StampRowAudit`), ne nove. `OdgovorniRadnik` v1 = tekst |
| 32–34 (pakovanje je proces, tipovi LJ) | **Prihvata se** | `TipJedinice ∈ {PALETA, BULK, BLOK, CISTERNA, KONTEJNER}`; pakovanje na LJ preko postojećih šifarnika `tblKutije`/`tblKese` |
| 35 (status LJ) | **Menja se** | Skladišti se samo `Stornirano` (registar). `POTROSENA / OTPREMLJENA / BLOKIRANA` se **izvode** iz ledgera (ulazi procesa, utovar, blokade) |
| 36–40 (quality hold, uzorci, HACCP, kontrole, metal detektor) | **Odlaže se** (Faza D) | `tblBlokadeRobe` ulazi u `RaspolozivoKg` kao odbitak; `tblProcesKontrole` generička; `IzvorTip+IzvorID` polimorfija prihvaćena **samo** za uzorke/kontrole, ne za graf |
| 43–47 (graf, forward/backward, lot) | **Prihvata se** (Faza C) | Graf se gradi jednim prolazom (S5 pravilo); LANAC ostaje linearan po otkupnom bloku sa multi-vrednosnom kolonom; nova lista „Genealogija" po LJ. `LotBroj` = opciona kolona LJ, `tblLotovi` Faza F |
| 48–49 (rezervacije, utovar) | **Utovar već postoji (#248)**; rezervacije van obima | Utovar mora naučiti LJ (Faza D). Rezervacije prodaje nisu planirane |
| 50 (4 faze migracije) | **Prihvata se, prilagođeno** | A: `tblPrerada + SarzaID + LagerJedinicaID`; B: novi writer-i + projekcija; C: novi ekran + graf; D: backfill **eksplicitan** (Alt+F8, kao `BackfillUtovariIzGPFaktura`), ne automatski |
| 51–54 (kapije, dva koraka, storno, korekcija) | **Prihvata se** | Kapije u writer-u (`testovi.md` §5). Korekcija = storno + nova šarža (ADR-0001), bez `CorrectionProcess` tabele. Odluka D2: ulaz se knjiži pri OTVARANJU |
| 55–58 (ekran) | **Prilagođava se ljusci** | Ljuska ima **jednu deljenu mrežu + zonu**, ne tri kolone. Tri zone predloga postaju tri liste jednog ekrana (`ULAZ`, `IZLAZ`, `SARZE`) sa korpom u stanju modula (obrazac `modScrFakture`). Živi balans u KPI traci zone; „ZAVRŠI" ugašen van tolerancije |
| 59 (izveštaji) | **Prihvata se** kao Faza C/D čitanja | Randman, gubici, stanje proizvodnje, genealogija — sve su čitanja nad tri tabele |

---

## 3) Model podataka v1 (Faza A)

Imena po kućnoj konvenciji: PascalCase bez dijakritike, ID sufiks `ID`, šifarnici
ključani tekstom (`Sifra`), brojevi dokumenata `Broj + Godina` sa resetom po godini,
`Stornirano` na svakoj dokument-tabeli, audit kolone kroz `EnsureAuditColumns`.

### 3.1 Matični podaci

**`tblTipoviProcesa`** — `Sifra | Naziv | MenjaProizvod | ZahtevaOpremu |
DozvoljenaUlaznaForma | ObavezniParametri | Aktivan`
Seed: `PRANJE, SORTIRANJE, KALIBRACIJA, PREBIRANJE, ZAMRZAVANJE, IZBIJANJE_KOSTICE,
PAKOVANJE, PREPAKIVANJE, PASIRANJE, BLOK, ODMRZAVANJE, PRERADA_LEGACY`.
`ObavezniParametri` = lista ključeva odvojena `;` (npr. za ZAMRZAVANJE:
`OPREMA;VREME_ULAZ;VREME_IZLAZ;TEMP_ROBE_ULAZ;TEMP_ROBE_IZLAZ;CILJNA_TEMP`).

**`tblProizvodi`** — `ProizvodID | VrstaVoca | Naziv | Forma | Prodajni | IzvorTip |
IzvorKljuc | Aktivan`
- `Forma ∈ {SVEZE, SMRZNUTO, BLOK, PIRE, BULK}`; `Prodajni=Da` za ono što ide na utovar.
- Seed (idempotentan, u `EnsureProizvodnjaSchema`): iz `tblKulture` po `VrstaVoca`
  (`IzvorTip=KULTURA`, forma SVEZE) i iz `tblVrstaGotovihProizvoda`
  (`IzvorTip=VGP`, `IzvorKljuc=TipGotovogProizvoda`, `Prodajni=Da`). `RokMeseci` se
  **ne kopira** — čita se sa VGP preko `IzvorKljuc` (jedan izvor istine).
- Klasa i kalibracija **nisu** deo proizvoda (kolone LJ).

**`tblOprema`** — `OpremaID | StanicaID | TipOpreme | Naziv | KapacitetKg | Aktivan`
(`TipOpreme ∈ {TUNEL, LINIJA, IZBIJAC, PAKERICA, KALIBRATOR, PASIRKA, METAL_DETEKTOR, OSTALO}`).

Unos matičnih podataka: kroz postojeći data-driven meni `frmMaticniPodaci` /
`modMaticniLookups` (bez novih formi, bez `.frx`).

### 3.2 Lager jedinica

**`tblLagerJedinice`** — `LagerJedinicaID | BrojJedinice | Godina | TipJedinice |
ProizvodID | Klasa | Kalibracija | KgPocetno | LotBroj | TipKutije | BrojKutija |
TipKese | BrojKesa | TezinaPaleteKg | BrutoKg | DatumNastanka | StanicaID |
IzvorSarzaID | IzvorTip | IzvorID | Napomena | Stornirano` + audit.

- `IzvorTip ∈ {SARZA, PALETA, PRERADA}`: `SARZA` = rođena u procesu; `PALETA` /
  `PRERADA` = **materijalizovan** legacy red (pokazivač na `tblPaleta.PaletaID`,
  odnosno `tblPrerada.PreradaID`).
- Za `IzvorTip=PALETA` fizička masa je **živa** `tblPaleta.NetoKg` (header sme da se
  menja kroz Reassign/Adjust dok paleta nije ušla u proces); `KgPocetno` je snimak za
  audit. Za `SARZA` i `PRERADA` fizička masa = `KgPocetno` (= `NetoIzlazKg`).
- **Nema kolone `Status`.** Izvedeno: `POTROSENA` (raspoloživo ≈ 0 kroz ulaze/utovar),
  `BLOKIRANA` (Faza D), `STORNIRANA` (`Stornirano=Da`).

### 3.3 Procesna šarža

**`tblProcesSarze`** — `SarzaID | BrojSarze | Godina | TipProcesa | StanicaID |
OpremaID | DatumVremePocetak | DatumVremeKraj | Status | OdgovorniRadnik | Napomena |
Stornirano` + audit.
- `Status ∈ {OTVORENA, ZAVRSENA}`; storno je `Stornirano=Da` (registar), ne treće stanje.
- Bez `Voce/Klasa/Kg/PaletaID` u zaglavlju — tačno kako predlog kaže.
- Bez snimljenih `UlazKg/IzlazKg`: izvode se; `modIntegritet` P1 ih proverava.

**`tblProcesUlazi`** — `ProcesUlazID | SarzaID | LagerJedinicaID | KgUlaz | Napomena |
Stornirano` + audit.

**`tblProcesIzlazi`** — `ProcesIzlazID | SarzaID | LagerJedinicaID | ProizvodID |
Klasa | Kalibracija | KgIzlaz | TipIzlaza | Napomena | Stornirano` + audit.
- `TipIzlaza ∈ {GLAVNI, NUSPROIZVOD, OTPAD, GUBITAK}`.
- `LagerJedinicaID` je prazan za `OTPAD` i `GUBITAK` (nema jedinice koja bi se
  dalje pratila); za `GLAVNI`/`NUSPROIZVOD` obavezan.
- `ProizvodID/Klasa/Kalibracija` su **denormalizovani snimak** sa LJ (obrazac
  `BrojPrerade` na utovarnoj stavci u #248) — izveštaj po šarži ne mora u join.

**`tblProcesParametri`** — `ParametarID | SarzaID | Kljuc | Vrednost | Jedinica |
Stornirano` + audit. Nosi zamrzavanje (vremena, temperature, tunel), Brix, itd.

### 3.4 Most ka legacy tabelama (Faza A šema, Faza B upis)

- `tblPrerada` **+ `SarzaID` + `LagerJedinicaID`** (na kraj, `EnsureColumnOnTable`).
  Red sa `SarzaID` je **projekcija** prodajnog izlaza 2.0: `NetoIzlazKg=KgIzlaz`,
  `TipGotovogProizvoda = Proizvod.IzvorKljuc`, kutije/kese/bruto sa LJ,
  `NetoUlazKg` prazno, **bez** `tblPreradaStavka` redova.
- `tblUtovarStavke` **+ `LagerJedinicaID`** (prazno u Fazi B; puni se u Fazi D kad
  utovar nauči LJ).
- Registar: šest novih tabela u `STORNO_TABELE` (`vba_check` pravilo `STORNO_REGISTAR`),
  matične u `BEZ_STORNA`, sve u `AuditableTables()`, self-heal u `EnsureRuntimeSchema`.

### 3.5 ID-jevi i brojevi

`GetNextID` prefiksi (provera preklapanja `Left$`): `SRZ-` (šarža), `PUL-` (ulaz),
`PIZ-` (izlaz), `PPR-` (parametar), `LJ-` (jedinica), `PRZ-` (proizvod), `OPR-` (oprema).
`BrojSarze` = `maxN+1` unutar godine (mirror `GenerateBrojPrerade`), prikaz
`SORT 12/2026`; jedan brojač za sve tipove (odluka D4).

---

## 4) Invarijante i writer-i (Faza B, modul `modProizvodnja.bas`)

### 4.1 Jedna funkcija raspoloživog

```
RaspolozivoKg(ljID) = FizickoKg(lj)
                    - Σ aktivnih tblProcesUlazi.KgUlaz za lj
                    - UtovarenoKg(lj)         ' legacy PRERADA: UtovarenoPoPreradi(IzvorID)
                    - BlokiranoKg(lj)         ' Faza D, do tada 0
```
Mape se grade **jednim prolazom** (`PotrosenoPoJedinici()`, isti obrazac kao
`UtovarenoPoPreradi`) i dele ih mreža, writer i storno kapija — nikad tri kopije.

Posledica na #248 (mali, obavezan dodir u Fazi B): `GetGPZaFakturisanjeForGrid` i
`CreateUtovarCore` odbijaju i **procesno potrošene** kg legacy GP lota
(prepakivanje/blok od GP lota), inače je moguća dvostruka upotreba iste robe.

### 4.2 `OtvoriSarzu_TX(tipProcesa, stanicaID, opremaID, pocetak, ulazi, parametri)`

Snapshot: `tblProcesSarze, tblProcesUlazi, tblProcesParametri, tblLagerJedinice,
tblPaleta` (materijalizacija + eventualno `Preradjeno`).
Kapije **pre prvog upisa**, redom: tip procesa postoji i aktivan; stanica postoji;
oprema (ako tip traži) postoji, aktivna, iste stanice; svaki ulaz postoji **tačno
jednom** (`RequireSingleRowIndexByKey` obrazac), nije storniran, nije dupliran u
listi; `KgUlaz > 0` i `≤ RaspolozivoKg`; forma proizvoda ulaza dozvoljena za tip;
obavezni parametri tipa prisutni. Tek onda: LJ materijalizacija za legacy ulaze,
header (`Status=OTVORENA`), ulazi, parametri; legacy paleta dobija `Preradjeno=Da`
ako joj je raspoloživo palo na 0.

**Ulaz se knjiži pri otvaranju** (odluka D2): tunel radi 18 h, roba ne sme za to
vreme da bude „raspoloživa" drugom procesu ili utovaru. Ulazi su posle otvaranja
nepromenljivi — promena = storno otvorene šarže (uvek dozvoljen, nema izlaza) + nova.

### 4.3 `ZavrsiSarzu_TX(sarzaID, kraj, izlazi, parametri)`

Snapshot: gornje + `tblProcesIzlazi` + `tblPrerada` (projekcija).
Kapije: šarža postoji tačno jednom, `OTVORENA`, nije stornirana; svaki izlaz ima
proizvod (aktivan), `KgIzlaz > 0`, tip izlaza iz skupa; `GLAVNI` bar jedan;
**balans**: `|Σ ulaza − Σ izlaza| ≤ tolerancija` (`PROCES_BALANS_TOLERANCIJA_KG`,
default 0,5 kg kao A5 — odluka D3). Ekran nudi „Knjiži razliku kao tehnološki
gubitak" koje **dodaje red** `GUBITAK`, ne sakriva razliku (§10 predloga).
Upis: LJ za svaki `GLAVNI/NUSPROIZVOD` (`IzvorTip=SARZA`), izlazi, parametri,
`DatumVremeKraj`, `Status=ZAVRSENA`, projekcija u `tblPrerada` za `Prodajni=Da`.

### 4.4 `StornoProcesSarza_TX(sarzaID)` (u `modStorno`, kućni obrazac)

`RequireStornoAllowed` → kapija nizvodnog: nijedna izlazna LJ nema aktivan
`tblProcesUlazi`, aktivnu utovarnu stavku (preko projektovanog `PreradaID` dok traje
Faza B) ni blokadu. Ako ima — greška sa imenom potomka (fail-closed, kao
`StornoPrerada` ERR+53). Inače: izlazi + LJ + projekcija + ulazi + parametri +
header `Stornirano=Da`; legacy paleti se vraća `Preradjeno` samo ako više nema
**nijednog** aktivnog ulaza koji je troši. `StornoPrerada` nad projektovanim redom
(`SarzaID` neprazan) **odbija** sa uputom na storno šarže.

### 4.5 Zaštita legacy motora paleta

`ReassignPaleteToPrijemnica_TX`, `DetachOsirocenePaletaStavke_TX`,
`AdjustPaletaGajbiceZaPrijemnicu_TX` danas gledaju samo `IsPaletaPreradjena`
(AUD-029 već beleži da diraju prerađene palete). Dodaje se `PaletaUProcesu(palID)`
(ima materijalizovanu LJ sa aktivnim ulazom) na **ista** mesta — paleta koja je
delimično ušla u proces se ne prevezuje.

### 4.6 `modIntegritet` — nove provere

- **P1** balans po završenoj šarži (van tolerancije).
- **P2** ulaz veći od raspoloživog (posledica ručnog diranja tabela).
- **P3** projekcija ≠ LJ (`tblPrerada.SarzaID` red čiji `NetoIzlazKg` ≠ `KgIzlaz`).
- **A5** preskače redove sa `SarzaID` (ulaz živi u procesu, ne u stavkama).

---

## 5) Sledljivost kao graf (Faza C)

- `modSledljivost`: `ProcesGraf()` gradi `parents(lj)` i `children(lj)` jednim
  prolazom kroz ulaze/izlaze; legacy karike (`tblPreradaStavka`) se učitavaju kao
  ivice istog grafa (paleta → legacy prerada).
- **LANAC** ostaje jedan red po otkupnom bloku; kolona „Pal. gotovog proizv."
  postaje multi-vrednosna („3 sarze / 5 LJ"), „Stanje" dobija `u procesu` i `prerađeno
  (2.0)`; fail-closed pravila #248 ostaju.
- Nova lista **„Genealogija"** na ekranu Sledljivost: unos = LJ / broj šarže / broj
  prerade / broj palete; izlaz unapred (procesi → LJ → utovari → fakture → kupci) i
  unazad (šarže → ulazne LJ → palete → prijemnice → zbirne → otkupi → kooperanti →
  parcele). To je recall funkcija; PDF preko postojećeg `SledljivostSablon` obrasca.
- `SledGpMape` u Fazi B dobija minimalnu granu: paleta → (ulaz) → šarža → (izlaz) →
  projektovani `PreradaID`, da otkup čija je roba prošla 2.0 ne bi pisao „u hladnjaci".

---

## 6) Ekran `PROIZVODNJA` (Faza C, `modScrProizvodnja.bas`)

Registruje se u `modUiScreens.ScrRows` (`PROIZVODNJA | modScrProizvodnja |
OTKUI_NAV_PROIZVODNJA | ikonica | OPERACIJE | OBL_PALETE`) — kasno vezano, klijent bez
modula vidi prigušenu stavku. Oblast prava je `OBL_PALETE` (ko sme palete, sme i
preradu) — `modAuth.OblastiList` se ne dira.

| Lista | Mreža | Zona |
|---|---|---|
| `SARZE` | pregled šarži (broj, tip, oprema, datum, ulaz kg, izlaz kg, balans, status) | KPI traka; radnje: Procesni list, Završi, Storno |
| `ULAZ` | lager jedinice sa **raspoloživo kg** (izvedeno), poreklo, kvačica + polje „kg za ulaz (prazno = sve)" — obrazac `NOVAPRERADA` + „Kol. za utovar" iz #248 | polja zaglavlja: tip procesa, stanica, oprema, početak; parametri po tipu (za ZAMRZAVANJE: tunel, vreme ulaz/izlaz, temperature) — pale se po tipu, isti obrazac kao `PoljaPrerade`; dugme **Otvori šaržu** |
| `IZLAZ` | korpa izlaza otvorene šarže (proizvod, klasa, kalibracija, kg, pakovanje, tip izlaza) | polja za dodavanje reda; **KPI: ULAZ · GLAVNI · NUSPROIZVOD · OTPAD · GUBITAK · NERASPOREĐENO**; „Knjiži gubitak"; **Završi** ugašeno dok nije u toleranciji |
| `LAGER` | pregled LJ (raspoloživo, poreklo, šarža) | radnje: Paletni list LJ, Genealogija (skok na Sledljivost) |

Korpa i izabrana šarža žive u stanju modula (`mKorpa*`, kao `modScrFakture`);
ekran **ne računa i ne upisuje** — sve ide kroz §4 writere. Test seam-ovi
`Scr_*Test` iza `IsTestMode()`. Sve poruke kroz `modPoruke.UpsertPoruke`.
Lista `NOVAPRERADA` na ekranu Palete postaje tanak wrapper: „Preradi izabrane" =
`OtvoriSarzu_TX(PRERADA_LEGACY, cele palete)` + `ZavrsiSarzu_TX(1 GLAVNI izlaz)`; legacy
`frmPalete` i `SavePrerada_TX` ostaju netaknuti do Faze D (pravilo §5).

---

## 7) Dokumenti

| Dokument | Faza | Obrazac |
|---|---|---|
| **Procesni list** (zaglavlje, ulazi, izlazi, balans, parametri, potpis) | C | `ProcesSablon` u `modPrint` (`EnsureProcesSablon`/`FillProcesSablon`, `LAYOUT_VER` marker), ključ `PROCES_PRINT_MODE` u `modPodesavanja` (default PDF, kao utovar) |
| List zamrzavanja | C | isti obrazac; blok parametara nosi tunel/vremena/temperature; poseban naslov po tipu |
| Paletni list LJ | C | `LjSablon` (paleta gotove robe: proizvod, klasa, lot, pakovanje, kg, poreklo — šarža i ulazne palete) |
| Transfer list | E | uz lokacije |
| Packing list | po potrebi | — |

---

## 8) Plan po fazama

Svaka faza = zaseban PR, zelena FAST + TARGETED (na Windows-u), `who_writes`
regenerisan, `Compile` ručna kapija, smoke checklista operatera. `.claude/` se ne
dira u feature PR-u.

| Faza | Sadržaj | Fajlovi (novi / dirani) | Kapija |
|---|---|---|---|
| **0 — odluke i baseline** | ovaj dokument; odluke D1–D5; merge #247 i #248; rebase grane | `docs/` | FULL suite zelen na `main` |
| **A — šema + matični podaci** | 7 novih tabela + 3 kolone (§3), seed proizvoda/tipova, registri, self-heal, matični unos, fixture `ENSURE_TABLES` | novi: `modProizvodnja.bas` (šema+seed); dirani: `modConfig`, `modSetup`, `modSchemaGuard`, `modMaticniLookups`, `tools/make_fixture.py` | `vba_check`; test „šema prisutna i idempotentna" |
| **B — writer-i + read model** | `RaspolozivoKg`/`PotrosenoPoJedinici`, `OtvoriSarzu_TX`, `ZavrsiSarzu_TX`, `StornoProcesSarza_TX`, materijalizacija, projekcija, zaštita paleta (§4.5), utovar „na stanju − proces", A5/P1–P3, minimalna grana `SledGpMape` | dirani: `modProizvodnja`, `modStorno`, `modPaletniList`, `modUtovar`, `modFaktura`, `modIntegritet`, `modIzvestaj` | nova suite **`RunProizvodnjaTestSuite`** (`gate: True` u `SUITES`); sabotaže za balans i storno kapiju (kritične invarijante → dvosmerni dokaz obavezan) |
| **C — ekran + dokumenti + graf** | `modScrProizvodnja`, wrapper `NOVAPRERADA`, `ProcesSablon`, `LjSablon`, `PROCES_PRINT_MODE`, `ProcesGraf`, LANAC multi-kolona, „Genealogija" | novi: `modScrProizvodnja`; dirani: `modUiScreens`, `modPoruke`, `modPrint`, `modPodesavanja`, `modScrPalete`, `modSledljivost`, `modIzvestaj`, `modScrSledljivost`, `modTest` | testovi ugovora ekrana (`T_Proiz_UgovorEkrana`, korpa, balans KPI), `T_Sled_Genealogija`; smoke operatera |
| **D — migracija + kvalitet + gašenje projekcije** | `BackfillSarzeIzLegacyPrerada` (Alt+F8, opciono, idempotentno: 1 prerada → `PRERADA_LEGACY`, ulazi = stavke, 1 GLAVNI, razlika = `GUBITAK` sa napomenom LEGACY); `tblBlokadeRobe` u `RaspolozivoKg`; `tblProcesKontrole` (temp, Brix, metal detektor → blokada na FAIL); **utovar/faktura/SEF/sledljivost uče `LagerJedinicaID`**, projekcija u `tblPrerada` prestaje | `modProizvodnja`, `modUtovar`, `modFaktura`, `modSEFMapper`, `modIzvestaj`, `modScrFakture`, `modSetup` | isti #248 testovi prošireni na LJ ključ; sabotaže #248 ostaju zelene |
| **E — lokacije** | `tblKomore`, `tblPozicije`, `tblLagerLokacija` (istorija), `TransferLager_TX`, Transfer list, komora = tip (ne hardkodovan) | novi: `modLokacije.bas`; ekran `LAGER` dobija lokaciju | suite proširena |
| **F — 2027 obim (odluka 66/67)** | smene/radnici/produktivnost, senzori i tuneli kao izvori merenja, `tblKalibracije`, `tblSpecifikacije`, `tblLotovi`, kalibracija kao entitet | — | — |

Procena (relativna, ne kalendarska): A ≈ 1 PR-krug; B ≈ 2–3 kruga (writer-i +
zaštita paleta su srce); C ≈ 2–3 kruga (ekran u ljusci je najskuplji deo, iskustvo
#248: 11 revizionih krugova za GP prodaju); D ≈ 2 kruga (dodir na svež #248 kod).

---

## 9) Odluke koje treba potvrditi pre Faze A

| # | Pitanje | Preporuka | Alternativa i cena |
|---|---|---|---|
| **D1** | Kako prodajni izlaz 2.0 stiže do utovara/fakture/SEF-a? | LJ + **projekcija** u `tblPrerada` (Faza B), gašenje u Fazi D | odmah generalizovati stek #248 na `LagerJedinicaID`: dodir `modUtovar`, `modFaktura`, `modSEFMapper`, `modIzvestaj`, `modScrFakture`, testovi 160–163 i 14+ sabotaža — na kodu koji je tek stabilizovan i još nije u `main`-u |
| **D2** | Kad se ulaz knjiži? | pri **OTVARANJU** šarže (rezervacija u vremenu trajanja procesa) | pri završetku: kraće, ali roba u tunelu izgleda raspoloživa |
| **D3** | Tolerancija balansa | apsolutna kg iz `Podešavanja` (`PROCES_BALANS_TOLERANCIJA_KG`, default 0,5) | procenat od ulaza; kombinacija (max od oba) |
| **D4** | Numeracija šarže | jedan brojač po godini, prikaz sa šifrom tipa | brojač po tipu (SORT-125 / ZAM-155) — više brojača, više mesta za grešku |
| **D5** | Da li `NOVAPRERADA` na Paletama odmah ide kroz 2.0 writer (Faza C) | da — jedan put upisa | ostaviti `SavePrerada_TX` do Faze D: dva paralelna writera iste robe |

---

## 10) Rizici i ono što ovaj plan ne rešava

- **Dvostruki zapis u Fazi B** (LJ + projekcija u `tblPrerada`) je namerna prelazna
  kopija sa rokom (Faza D). Provera P3 čuva da se ne raziđu; `who_writes` će
  pokazati `modProizvodnja` kao novog pisca `tblPrerada` — očekivano.
- **Preklapanje sa #248** u `modUtovar`/`modFaktura` (§4.1) je mali ali obavezan
  dodir: bez njega legacy GP lot koji uđe u prepakivanje ostaje „na stanju".
- **Excel performanse:** graf i raspoloživo su mape jednim prolazom; ipak, tri
  tabele više po svakom `GetTableData` na ekranu zahtevaju `BeginTableCache` u
  `Scr_Rows` kao na Sledljivosti.
- **Verifikacija u web sesiji ne postoji:** `run_vba` traži Windows + Excel. Svaka
  faza se prijavljuje kao **neverifikovana** dok operater ne pokrene suite;
  `vba_check` je jedino što Linux daje.
- Ne rešava: rezervacije prodaje, cenu GP-a po LJ (ostaje unos operatera kao u
  #248), PWA/GAS sync ovih tabela (nema ga ni za palete danas).

---

## 11) Šta se traži od suite-a (definicija gotovog po fazi)

- Pad **po imenu** za svaku kapiju writer-a (ulaz ne postoji / storniran / dupliran /
  > raspoloživo / forma nedozvoljena / parametar fali / balans van tolerancije / bez
  GLAVNI izlaza / storno sa potomkom).
- Rollback dokaz: namerna greška posle upisa izlaza vraća **sve** tabele iz snapshota
  (obrazac `mTestFailPosleRelease` iz #248).
- N→M scenario iz predloga (1.000 kg → 520/300/150/20/10) i merge (3 → 1) prolaze,
  raspoloživo posle svake operacije tačno na 0,01 kg; parcijalni ulaz ostavlja
  ostatak; `Preradjeno` se pali tek na 0.
- Sabotaže (`tools/sabotaza.py`): `proces-balans-placebo` (tolerancija ignorisana),
  `proces-storno-potomak` (kapija uklonjena), `proces-raspolozivo-bez-utovara`
  (utovar izostavljen iz odbitka) — svaka obara tačno svoj test i vraća se
  bit-identično.
