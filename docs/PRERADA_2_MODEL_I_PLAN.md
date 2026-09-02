# Prerada 2.0 — proizvodno jezgro: model, specifikacija i plan isporuke

- **Status:** plan usvojen na nivou odluka D1–D5 (2026-09-02); **nije implementirano**.
  Ovaj dokument je specifikacija za implementaciju, ne opis koda.
- **Verzija dokumenta:** 2 (v1 je bila analiza sa otvorenim odlukama; ovde su odluke
  unete i plan je razrađen do nivoa PR-ova, funkcija, tabela, testova i sabotaža).
- **Grana:** `claude/prerada-2-production-core-stfebv`
- **Preduslovi u `main`-u:** [#247](https://github.com/dusanmiladinovicVNM/otkupapp-pwa/pull/247)
  i [#248](https://github.com/dusanmiladinovicVNM/otkupapp-pwa/pull/248) spojeni;
  FULL suite zelen.
- **Vezano:** `docs/DOMEN/README.md`, `docs/adr/0001-*.md`, `docs/adr/0002-*.md`,
  `docs/AMBALAZA_MODEL.md`, `docs/UI_MIGRACIJA_KATALOG.md` §25,
  `docs/EXCEL_TEST_HARNESS.md`, `docs/INTEGRITET_PROVERE.md`,
  `docs/Master Plan/09_QA_DECISION_LOG.md` (odluke 66, 67).

---

## 0) Sažetak

Prerada 2.0 uvodi **procesnu šaržu** kao proizvodni događaj: *N ulaznih lager jedinica
→ šarža → M izlaznih lager jedinica*, sa klasifikovanim izlazom (glavni proizvod,
nusproizvod, otpad, tehnološki gubitak) i masa-balansom kao invarijantom writer-a.
**Lager jedinica** (`tblLagerJedinice`) postaje jedini ključ robe u lageru i u
sledljivosti: sveža paleta i legacy „paleta gotovog proizvoda" (`tblPrerada`) se
predstavljaju kao lager jedinice, a prodajni stek iz #248 (utovar → faktura → SEF →
sledljivost) se **odmah** prevodi na taj ključ (odluka D1). Isporuka ide u sedam
PR-ova (A, B1, B2, C1, C2, D, E), svaki sa sopstvenom definicijom gotovog,
suite-om i sabotažama. Komore, senzori, smene, HACCP, specifikacije i lotovi su
kasnije faze, u skladu sa odlukama 66 i 67 Master Plana.

---

## 1) Usvojene odluke

| # | Odluka | Posledica |
|---|---|---|
| **D1** | **Prodajni stek ide odmah na `LagerJedinicaID`.** Nema projekcije 2.0 izlaza u `tblPrerada`. | Faza B1 je refaktor #248 koda (`modUtovar`, `modFaktura`, `modSEFMapper`, `modIzvestaj`, `modScrFakture`, `modStorno`, `modSetup`, `clsSEFLine`, testovi 160–163, ~20 sabotaža) na jedan ključ. Sve legacy prerade dobijaju lager jedinicu (materijalizacija, deterministička). |
| **D2** | Ulaz se knjiži pri **otvaranju** šarže. | Roba u tunelu nije raspoloživa ni utovaru ni drugom procesu. Ulazi su posle otvaranja nepromenljivi; promena = storno otvorene šarže (uvek dozvoljen) + nova. |
| **D3** | Tolerancija balansa je **apsolutni kg** iz Podešavanja, `PROCES_BALANS_TOLERANCIJA_KG`, default **0,5**. | Isti prag kao `modIntegritet` A5. Razlika iznad praga blokira „Završi"; razlika ispod praga se **ne skriva** — prijavljuje se kao `NERASPOREDJENO` u izveštaju šarže. |
| **D4** | **Jedan brojač** `BrojSarze` po godini za sve tipove; prikaz `SORT 12/2026`. | Mirror `GenerateBrojPrerade`; tip je prefiks u prikazu, ne deo broja. |
| **D5** | Lista „Nova prerada" na ekranu Palete ide kroz 2.0 writer od Faze C. | Posle C1 `SavePrerada_TX` zove samo legacy `frmPalete`; posle D se gasi i tamo. |

Odbačeno u v1 i ostaje odbačeno: `KgTrenutno` keš na jedinici; `Status` kolona na
jedinici; `tblZamrzavanje` subtype; `tblTuneli` kao posebna tabela; `CorrectionProcess`.

---

## 2) Stanje koda na koje se plan oslanja (činjenice)

| Oblast | Činjenica | Referenca |
|---|---|---|
| Legacy prerada | `SavePrerada_TX`: N **celih** paleta → 1 izlaz; `Preradjeno=Da`; bez tipa procesa, otpada, gubitka | `modPaletniList.bas:2872-2990` |
| `tblPrerada` kolone | `PreradaID, BrojPrerade, Godina, Datum, NetoUlazKg, NetoIzlazKg, BrojKutija, BrojKesa, TezinaPaleteKg, BrutoKg, AmbalazaKg, TipKutije, TipKese, TipGotovogProizvoda, Napomena, CreatedAt, Stornirano` | `modConfig.bas:321-347` |
| Paleta | header = projekcija iz aktivnih stavki; broj po godini; `Preradjeno` čitaju D1, `SledGpMape`, mreža paleta, Reassign/Detach/Adjust kapije (`IsPaletaPreradjena`) | `modPaletniList.bas:2434`, `modIntegritet.bas:652` |
| #248 model | `tblUtovar` (+8 prevoz kolona), `tblUtovarStavke(PreradaID, BrojPrerade, KolicinaKg, BrojKutija, BrojKesa, CenaKg)`, `tblPrevoznici`, `tblFakturaStavke(+PreradaID, BrojPrerade, UtovarID)`, `tblVrstaGotovihProizvoda(+RokMeseci)` | `modSetup.EnsureUtovarSchemaCore`, `modConfig` (#248) |
| #248 stanje | `UtovarenoPoPreradi()` — jedna mapa za mrežu, writer, storno; na stanju = `NetoIzlazKg − Σ` | `modUtovar.bas:53` (#248) |
| #248 SEF | XOR `PrijemnicaID`/`PreradaID` po stavci; `SellersItemIdentification = PreradaID`; `ValidateGpUtovarZaSEF` (utovar tačno jedan, aktivan, `Fakturisano=Da`, tvrdi fakturu, kg 1:1 u oba smera, kupac = kupac) | `modSEFMapper.bas:175-360` (#248) |
| #248 sledljivost | `SledGpMape`, `SledUtovarPoPreradi`, ključ dokaza `UtovarID|FakturaID|PreradaID → kg`; stanja `SLED_ST_*` | `modIzvestaj.bas:4404-4560, 5436-5447` (#248) |
| #248 storno | faktura oslobađa utovar; utovar samo nefakturisan; prerada sa utovarom se ne stornira (`ERR_STORNO_BASE+53`); poslednji zauzet kod `+70` | `modStorno.bas` (#248) |
| #248 ekran | `modScrFakture` liste `ZAFAKT | GOTOVA | UTOVARI | FAKTURE | SEF`; identitet GP reda kolona 9 (`FK_GP_KOL_ID`), dostupnost 10; korpa `mKorpa` u stanju modula; seam `Scr_FkKorpaTestDodajGP(preradaID, …)` | `modScrFakture.bas:85-125, 2048` (#248) |
| Ljuska | registar `modUiScreens.ScrRows`, kasno vezivanje `Application.Run`; ugovor `Scr_Meta/Build/Layout/Rows/Liste/Lista/Radnje/Event/Save/ResetCache/Cipovi/NaslovDopuna/Brojac`; polja `NewFieldG` sa prefiksom `scr`; mreža `GridCell`; `ShowToast` | `modUiScreens.bas:1-60`, `modOtkupUI.bas:4300, 5153, 6344` |
| Matični unos | `modMaticniLookups.MaticniSekcije` (data-driven meni) + `frmStammdaten` (`Select Case Me.Tag` → `m_TableName`) | `modMaticniLookups.bas:33`, `frmStammdaten.frm:1036-1072` |
| Podešavanja | `CfgAdd c, grupa, kljuc, natpis, tip`; grupe „Otkup / dokumenta", „Štampa", … | `modPodesavanja.bas:187` |
| Integritet | postojeći blokovi A1–A5, B2–B7, C1–C5, D1–D2; A5 = `NetoUlaz = Σ stavke`, `izlaz ≤ ulaz` | `modIntegritet.bas` |
| Greške | `modPaletniList` 7320–7344, `modCenovnik` 7701–7705, `modUtovar` 1730–1779, `modSetup` 9310–9321; **7400–7499 slobodno** | grep |
| Verifikacija | `SUITES` u `tools/run_vba.py` (nova suite `gate: True`); fixture `ENSURE_TABLES` + potpis; `tools/sabotaza.py` katalog (381 na #248); CI: `vba_check`, `who_writes --check` | `.claude/rules/testovi.md`, `docs/EXCEL_TEST_HARNESS.md` |
| Nepostojeće | komore, tuneli, oprema, smene, senzori, HACCP, uzorci, lotovi — ni u VBA ni u PWA | grep |

---

## 3) Ciljna arhitektura

```
                MATIČNI                    LAGER                     PROCES
  tblTipoviProcesa  tblProizvodi     tblLagerJedinice  <──────  tblProcesIzlazi
  tblOprema         (seed iz         (jedini ključ robe)         tblProcesUlazi  ──>  tblProcesSarze
                     tblKulture +        ▲   ▲                    tblProcesParametri
                     tblVrstaGP)        │   │ materijalizacija (IzvorTip/IzvorID)
                                        │   └──── tblPaleta (sveža; lenjo, pri ulasku u proces)
                                        └──────── tblPrerada (legacy GP lot; eagerno, sve)

  PRODAJA (#248, na LJ ključu od B1):  tblUtovarStavke.LagerJedinicaID ─> tblUtovar ─> tblFakture/Stavke ─> SEF
  SLEDLJIVOST (C2):  graf  paleta ─> LJ ─> šarža ─> LJ ─> … ─> utovar ─> faktura ─> kupac
```

Moduli (novi su podebljani):

| Modul | Uloga |
|---|---|
| **`modProizvodnja.bas`** | šema + seed (`EnsureProizvodnjaSchema`), materijalizacija, `RaspolozivoPoJedinici`, writer-i `OtvoriSarzu_TX` / `ZavrsiSarzu_TX`, read-modeli za mreže, `LjOznaka`, `BalansSarze` |
| **`modScrProizvodnja.bas`** | ekran `PROIZVODNJA` u ljusci (4 liste) |
| **`modTestProizvodnja.bas`** | suite `RunProizvodnjaTestSuite` |
| `modStorno` | `StornoProcesSarza_TX`; `StornoPrerada` uči LJ |
| `modUtovar`, `modFaktura`, `modSEFMapper`, `clsSEFLine`, `modScrFakture` | prodajni stek na LJ ključu (B1) |
| `modIzvestaj`, `modSledljivost`, `modScrSledljivost` | graf porekla, genealogija, LANAC multi-kolona (C2) |
| `modPaletniList` | `SavePrerada_TX` +LJ (A); kapija `PaletaUProcesu` u Reassign/Detach/Adjust (B2); wrapper `NOVAPRERADA` (C1) |
| `modIntegritet` | provere P1–P6 |
| `modPrint` | `ProcesSablon`, `LjSablon` (C1) |
| `modConfig`, `modSetup`, `modSchemaGuard`, `modPodesavanja`, `modPoruke`, `modMaticniLookups`, `frmStammdaten` (samo code-behind), `modUiScreens`, `modOtkupUI` (ikonica) | konstante, registri, unos matičnih, ključevi poruka |
| `tools/make_fixture.py`, `tools/sabotaza.py`, `tools/run_vba.py`, `tests/schema_donor.json` | harness |

---

## 4) Model podataka

Konvencije: PascalCase bez dijakritike; ID prefiks kroz `GetNextID`; brojevi
dokumenata `Broj + Godina` sa resetom po godini; `Stornirano` na dokument-tabelama;
audit kolone (`CreatedAt/By`, `ModifiedAt/By`) kroz `EnsureAuditColumns`; nove kolone
na postojećim tabelama **na kraj** (`EnsureColumnOnTable`), upis **po imenu**
(`RequireUpdateCell` posle `AppendRow`), nikad pozicijski za nove kolone.

### 4.1 `tblTipoviProcesa` (matični, `BEZ_STORNA`)

| Kolona | Tip | Obavezno | Opis |
|---|---|---|---|
| `Sifra` | tekst | da (PK) | `SORTIRANJE`, `ZAMRZAVANJE`, … |
| `Naziv` | tekst | da | prikaz |
| `MenjaProizvod` | Da/Ne | da | informativno (izveštaj) |
| `ZahtevaOpremu` | Da/Ne | da | kapija writer-a: `OpremaID` obavezan |
| `DozvoljenaUlaznaForma` | tekst `;` lista | ne | prazno = sve; kapija: forma proizvoda svakog ulaza mora biti u listi |
| `ObavezniParametri` | tekst `;` lista | ne | ključevi `tblProcesParametri` koje writer traži pri otvaranju |
| `Aktivan` | Da/Ne | da | |

Seed (idempotentno po `Sifra`):

| Sifra | MenjaProizvod | ZahtevaOpremu | DozvoljenaUlaznaForma | ObavezniParametri |
|---|---|---|---|---|
| `PRANJE` | Ne | Ne | `SVEZE` | |
| `SORTIRANJE` | Da | Da | `SVEZE;SMRZNUTO` | |
| `KALIBRACIJA` | Da | Da | `SVEZE;SMRZNUTO` | |
| `PREBIRANJE` | Da | Da | `SVEZE;SMRZNUTO` | |
| `ZAMRZAVANJE` | Da | Da | `SVEZE` | `VREME_ULAZ;VREME_IZLAZ;TEMP_ROBE_ULAZ;TEMP_ROBE_IZLAZ;CILJNA_TEMP` |
| `IZBIJANJE_KOSTICE` | Da | Da | `SVEZE;SMRZNUTO` | |
| `PAKOVANJE` | Da | Ne | `SMRZNUTO;BULK` | |
| `PREPAKIVANJE` | Ne | Ne | | |
| `PASIRANJE` | Da | Da | `SVEZE;SMRZNUTO` | `BRIX` |
| `BLOK` | Da | Da | `SMRZNUTO;PIRE;BULK` | |
| `ODMRZAVANJE` | Da | Ne | `SMRZNUTO` | |
| `PRERADA_LEGACY` | Da | Ne | | (koristi wrapper `NOVAPRERADA` i migracija D) |

### 4.2 `tblProizvodi` (matični, `BEZ_STORNA`)

| Kolona | Tip | Obavezno | Opis |
|---|---|---|---|
| `ProizvodID` | `PRZ-00001` | da (PK) | |
| `VrstaVoca` | tekst | da | isti domen kao `tblKulture.VrstaVoca` |
| `Naziv` | tekst | da | „Smrznuta malina Rolend", „Višnja BK I" |
| `Forma` | tekst | da | `SVEZE / SMRZNUTO / BLOK / PIRE / BULK` |
| `Prodajni` | Da/Ne | da | `Da` = sme na utovar |
| `IzvorTip` | tekst | ne | `KULTURA / VGP / RUCNO` |
| `IzvorKljuc` | tekst | ne | `VrstaVoca` (KULTURA) ili `TipGotovogProizvoda` (VGP) — most ka `RokMeseci` i legacy nazivima |
| `Aktivan` | Da/Ne | da | |

Seed (idempotentno po `IzvorTip+IzvorKljuc`): po jedan `SVEZE` proizvod za svaku
`VrstaVoca` iz `tblKulture` (`Prodajni=Ne`); po jedan proizvod za svaki
`tblVrstaGotovihProizvoda.TipGotovogProizvoda` (`Forma=SMRZNUTO`, `Prodajni=Da`,
`VrstaVoca` prazno ako se ne može izvesti — operater dopunjava). `RokMeseci` se **ne
kopira**; čita se `VGP.RokMeseci` preko `IzvorKljuc`, pa globalni default.
`Klasa` i `Kalibracija` **nisu** atributi proizvoda.

### 4.3 `tblOprema` (matični, `BEZ_STORNA`)

| Kolona | Tip | Obavezno | Opis |
|---|---|---|---|
| `OpremaID` | `OPR-00001` | da (PK) | |
| `StanicaID` | FK `tblStanice` | da | objekat (hladnjača) |
| `TipOpreme` | tekst | da | `TUNEL / LINIJA / IZBIJAC / PAKERICA / KALIBRATOR / PASIRKA / METAL_DETEKTOR / OSTALO` |
| `Naziv` | tekst | da | „Tunel 2" |
| `KapacitetKg` | broj | ne | informativno; upozorenje (ne blokada) kad `Σ ulaza > kapacitet` |
| `Aktivan` | Da/Ne | da | |

### 4.4 `tblLagerJedinice` (dokument-tabela, `STORNO_TABELE`, `AuditableTables`)

| Kolona | Tip | Obavezno | Opis |
|---|---|---|---|
| `LagerJedinicaID` | `LJ-00001` | da (PK) | jedini ključ robe |
| `BrojJedinice`, `Godina` | broj, broj | da | za `IzvorTip=SARZA` sopstveni brojač po godini; za materijalizovane = broj i godina porekla |
| `TipJedinice` | tekst | da | `PALETA / BULK / BLOK / CISTERNA / KONTEJNER` |
| `ProizvodID` | FK | da | |
| `Klasa` | tekst | ne | `I`, `II`, `industrija`… |
| `Kalibracija` | tekst | ne | slobodan opis v1 („18–24 mm") |
| `KgPocetno` | broj | da | fizička masa pri nastanku; za `PALETA` poreklo — snimak (živa masa je `tblPaleta.NetoKg`) |
| `LotBroj` | tekst | ne | komercijalni lot, slobodan v1 |
| `TipKutije`, `BrojKutija`, `TipKese`, `BrojKesa` | tekst/broj | ne | pakovanje (šifarnici `tblKutije`/`tblKese`) |
| `TezinaPaleteKg`, `BrutoKg` | broj | ne | |
| `DatumNastanka` | datum | da | osnova roka trajanja |
| `StanicaID` | FK | da | |
| `IzvorTip` | tekst | da | `SARZA / PALETA / PRERADA` |
| `IzvorID` | tekst | da | `SarzaID` / `PaletaID` / `PreradaID` |
| `Napomena` | tekst | ne | |
| `Stornirano` | Da/prazno | da | |

Bez `Status` kolone. Izvedeno pri čitanju: `POTROSENA` (raspoloživo ≤ 0,01),
`BLOKIRANA` (Faza D), `STORNIRANA`.

**Materijalizacija:**
- `PRERADA`: **eagerno**, za **sve** redove `tblPrerada` (i stornirane — LJ nasleđuje
  `Stornirano`, da istorijske utovarne/fakturne stavke uvek imaju jedinicu). Puni:
  `BrojJedinice/Godina = BrojPrerade/Godina`, `TipJedinice=PALETA`,
  `ProizvodID = tblProizvodi(IzvorTip=VGP, IzvorKljuc=TipGotovogProizvoda)`,
  `KgPocetno=NetoIzlazKg`, pakovanje/bruto/težina sa prerade, `DatumNastanka=Datum`,
  `StanicaID` = stanica iz `tblStanice.JeHladnjaca` ako je jedinstvena, inače prazno
  (P6 prijavljuje). Prerada bez `TipGotovogProizvoda` ili sa nepoznatim tipom dobija LJ
  bez `ProizvodID` → **nije dostupna za utovar** (fail-closed, P4 prijavljuje).
- `PALETA`: **lenjo**, u `OtvoriSarzu_TX`, pri prvom ulasku palete u proces:
  `ProizvodID = tblProizvodi(KULTURA, VrstaVoca)`, `Klasa = tblPaleta.Klasa`,
  `KgPocetno = NetoKg` (snimak), `DatumNastanka = tblPaleta.Datum`.

### 4.5 `tblProcesSarze` (`STORNO_TABELE`, audit)

| Kolona | Tip | Obavezno | Opis |
|---|---|---|---|
| `SarzaID` | `SRZ-00001` | da (PK) | |
| `BrojSarze`, `Godina` | broj, broj | da | D4 |
| `TipProcesa` | FK `tblTipoviProcesa.Sifra` | da | |
| `StanicaID` | FK | da | |
| `OpremaID` | FK | uslovno | obavezno ako tip `ZahtevaOpremu` |
| `DatumVremePocetak` | datum-vreme | da | |
| `DatumVremeKraj` | datum-vreme | pri završetku | |
| `Status` | tekst | da | `OTVORENA / ZAVRSENA` |
| `OdgovorniRadnik` | tekst | ne | v1 tekst; `CreatedBy` je audit |
| `Napomena` | tekst | ne | |
| `Stornirano` | | da | |

Bez `UlazKg/IzlazKg` kolona — `BalansSarze` ih izvodi; P1 proverava.

### 4.6 `tblProcesUlazi` (`STORNO_TABELE`, audit)

`ProcesUlazID (PUL-) | SarzaID | LagerJedinicaID | KgUlaz | Napomena | Stornirano`

### 4.7 `tblProcesIzlazi` (`STORNO_TABELE`, audit)

`ProcesIzlazID (PIZ-) | SarzaID | LagerJedinicaID | ProizvodID | Klasa | Kalibracija |
KgIzlaz | TipIzlaza | Napomena | Stornirano`

- `TipIzlaza ∈ {GLAVNI, NUSPROIZVOD, OTPAD, GUBITAK}`; `LagerJedinicaID` prazan za
  `OTPAD`/`GUBITAK`, obavezan inače.
- `ProizvodID/Klasa/Kalibracija` su snimak sa jedinice (kao `BrojPrerade` na utovarnoj
  stavci) — izveštaj po šarži bez join-a.

### 4.8 `tblProcesParametri` (`STORNO_TABELE`, audit)

`ParametarID (PPR-) | SarzaID | Kljuc | Vrednost | Jedinica | Stornirano`

Standardni ključevi: `VREME_ULAZ`, `VREME_IZLAZ`, `TEMP_ROBE_ULAZ`, `TEMP_ROBE_IZLAZ`,
`CILJNA_TEMP`, `TEMP_TUNELA_MIN`, `TEMP_TUNELA_MAX` (ručno do senzora), `BRIX`.
Ključ slobodan; tip procesa deklariše obavezne.

### 4.9 Izmene postojećih tabela

| Tabela | Kolona | Faza | Svrha |
|---|---|---|---|
| `tblPrerada` | `+LagerJedinicaID` | A | obrnuti pokazivač na materijalizovanu LJ (denorm; P4 čuva da nije prazan) |
| `tblPrerada` | `+SarzaID` | D | oznaka „migrirana u šaržu" (backfill) |
| `tblUtovarStavke` | `+LagerJedinicaID` | A (šema), B1 (ključ) | prodajni grain na LJ; `PreradaID/BrojPrerade` ostaju kao legacy denorm i prikaz |
| `tblFakturaStavke` | `+LagerJedinicaID` | A (šema), B1 (ključ) | dokaz `UtovarID|FakturaID|LagerJedinicaID`; SEF identitet |
| `tblLagerJedinice` | `+BlokiranoRazlog` (ne) | — | ne; blokade su ledger `tblBlokadeRobe` (D) |

Backfill (deterministički, idempotentan, u `EnsureProizvodnjaSchema` **i** kao
self-heal u `EnsureRuntimeSchema`): za svaki red `tblUtovarStavke`/`tblFakturaStavke`
sa `PreradaID` i praznim `LagerJedinicaID` upiši `LJ` iz mape `PreradaID → LJ`. Nema
izmišljanja: mapa je 1:1 iz materijalizacije. Ovo **sme** u self-heal (za razliku od
`BackfillUtovariIzGPFaktura`) jer ne legalizuje siročad — stavka bez `PreradaID`
ostaje bez LJ i P5 je prijavljuje.

### 4.10 Registri i konstante (`modConfig`, `modSchemaGuard`, `modSetup`)

- `TBL_TIPOVI_PROCESA, TBL_PROIZVODI, TBL_OPREMA, TBL_LAGER_JEDINICE, TBL_PROCES_SARZE,
  TBL_PROCES_ULAZI, TBL_PROCES_IZLAZI, TBL_PROCES_PARAMETRI` + `COL_*` po tabeli
  (prefiksi `TPR_`, `PRZ_`, `OPR_`, `LJ_`, `SRZ_`, `PUL_`, `PIZ_`, `PPR_`).
- `STORNO_TABELE` += `tblLagerJedinice, tblProcesSarze, tblProcesUlazi,
  tblProcesIzlazi, tblProcesParametri`; `BEZ_STORNA` += `tblTipoviProcesa, tblProizvodi,
  tblOprema`. (`vba_check` pravilo `STORNO_REGISTAR` ne pušta `ExcludeStornirano` nad
  neregistrovanom tabelom.)
- `AuditableTables()` += pet dokument-tabela; `EnsureRuntimeSchema` dopunjava audit
  kolone kao za utovar (#248 revizija #10).
- Konstante skupova: `PRZ_FORMA_*`, `LJ_TIP_*`, `SRZ_STATUS_OTVORENA/ZAVRSENA`,
  `PIZ_TIP_GLAVNI/NUSPROIZVOD/OTPAD/GUBITAK`, `LJ_IZVOR_SARZA/PALETA/PRERADA`.
- ID prefiksi: `SRZ-`, `PUL-`, `PIZ-`, `PPR-`, `LJ-`, `PRZ-`, `OPR-` (provera:
  `GetNextID` poredi `Left$`, pa `PRZ-` i `PRS-` ne kolidiraju; svaki je u svojoj tabeli).

---

## 5) Invarijante (izvor istine za testove i integritet)

| # | Invarijanta | Čuva | Proverava |
|---|---|---|---|
| I1 | `RaspolozivoKg(lj) = FizickoKg − Σ aktivnih ulaza − Σ aktivnog utovara − blokirano ≥ 0` | jedina funkcija `RaspolozivoPoJedinici` (deli je mreža, writer, storno, SEF) | P2 |
| I2 | Za `ZAVRSENA` šaržu: `|Σ KgUlaz − Σ KgIzlaz (sva četiri tipa)| ≤ tolerancija` | `ZavrsiSarzu_TX` | P1 |
| I3 | `ZAVRSENA` šarža ima bar jedan `GLAVNI` izlaz sa jedinicom | `ZavrsiSarzu_TX` | P1 |
| I4 | Jedinica koja ima aktivnog potomka (ulaz procesa / utovarnu stavku / blokadu) ne može se stornirati, ni posredno kroz storno šarže/prerade koja ju je proizvela | `StornoProcesSarza`, `StornoPrerada`, `StornoPaleta` | P3 |
| I5 | Paleta sa materijalizovanom LJ koja ima aktivan ulaz ne sme kroz Reassign/Detach/Adjust | `PaletaUProcesu` kapija u `modPaletniList` | P3 |
| I6 | Svaka aktivna prerada ima tačno jednu LJ (`IzvorTip=PRERADA`), `KgPocetno = NetoIzlazKg` | materijalizacija | P4 |
| I7 | Svaka aktivna utovarna/fakturna GP stavka nosi `LagerJedinicaID` koji postoji tačno jednom | B1 writer-i | P5 |
| I8 | `Preradjeno=Da` na paleti ⇔ raspoloživo ≤ 0,01 kg (ili legacy prerada) | `OtvoriSarzu_TX`, storno | D1 (postojeći) |
| I9 | Ulazi šarže su nepromenljivi posle otvaranja; izlazi posle završetka | nema writer-a koji ih menja; korekcija = storno + nova | — |
| I10 | Forma proizvoda ulaza ∈ `DozvoljenaUlaznaForma` tipa; obavezni parametri prisutni; oprema obavezna ako tip traži | `OtvoriSarzu_TX` | — |

---

## 6) Writer-i i read-model (`modProizvodnja.bas`)

### 6.1 Zajednički obrazac

- `*_TX`: `clsTransaction` + `AddTableSnapshot` za **svaku** tabelu koju piše (i uslovno
  za legacy tabele koje u datoj svesci ne moraju postojati); pre-validacija **svih**
  stavki pre prvog upisa; `RequireColumnIndex` kao fail-fast šema-kapija; upis novih
  kolona po imenu; `Monitor_Event` po uspehu, `Monitor_Error` + `RollbackTx` u EH;
  greška se re-raise-uje (UI je hvata i pretvara u toast).
- Greške: opseg `vbObjectError + 7400..7499`:

| Kod | Kapija |
|---|---|
| 7401 | tip procesa ne postoji / neaktivan |
| 7402 | stanica ne postoji |
| 7403 | oprema obavezna a nije data / ne postoji / neaktivna / druga stanica |
| 7404 | nema ulaza |
| 7405 | ulaz ne postoji tačno jednom (`RequireSingleRowIndexByKey` obrazac) |
| 7406 | ulaz storniran |
| 7407 | ulaz dupliran u listi |
| 7408 | `KgUlaz ≤ 0` ili nije numerički |
| 7409 | `KgUlaz > RaspolozivoKg` |
| 7410 | forma ulaza nije dozvoljena za tip |
| 7411 | obavezan parametar nedostaje |
| 7412 | paleta se ne može materijalizovati (nema `VrstaVoca` proizvod) |
| 7420 | šarža ne postoji tačno jednom |
| 7421 | šarža nije `OTVORENA` |
| 7422 | šarža stornirana |
| 7423 | nema izlaza / nema `GLAVNI` |
| 7424 | proizvod izlaza ne postoji / neaktivan |
| 7425 | `KgIzlaz ≤ 0` |
| 7426 | tip izlaza van skupa |
| 7427 | balans van tolerancije (poruka nosi `ulaz`, `Σ izlaza`, `razlika`, `tolerancija`) |
| 7428 | pakovanje neispravno (tip kutije/kese nepoznat; broj < 0) |
| 7429 | kraj pre početka |

- Monitor događaji: `SARZA_OTVORENA`, `SARZA_ZAVRSENA` (entityType `Sarza`);
  storno kroz `MonitorStornoSuccess SRC, "Sarza", id`.

### 6.2 Potpisi

```
' Jedna mapa: LagerJedinicaID -> raspolozivo kg. Jedan prolaz kroz tblLagerJedinice,
' tblProcesUlazi, tblUtovar(Stavke), tblBlokadeRobe (D). Legacy PALETA poreklo cita
' zivu tblPaleta.NetoKg.
Public Function RaspolozivoPoJedinici() As Object
Public Function RaspolozivoKg(ByVal ljID As String) As Double        ' omotac nad mapom
Public Function PotrosenoPoJedinici() As Object                      ' samo ulazi procesa
Public Function PaletaUProcesu(ByVal paletaID As String) As Boolean  ' I5 kapija
Public Function LjOznaka(ByVal ljID As String) As String             ' "PRE 51/2026" | "PAL 31/2026" | "LJ 12/2026"
Public Function LjRokTrajanja(ByVal ljID As String) As Variant       ' DatumNastanka + RokMeseci(VGP) | globalni; Empty kad nema
Public Function BalansSarze(ByVal sarzaID As String) As Variant      ' Array(ulaz, glavni, nus, otpad, gubitak, nerasporedjeno, randmanKom, recovery)

' ulazi:   Collection of Array(izvorTip, id, kgUlaz)   izvorTip in {"LJ","PALETA"}
' parametri: Collection of Array(kljuc, vrednost, jedinica)
Public Function OtvoriSarzu_TX(ByVal tipProcesa As String, ByVal stanicaID As String, _
        ByVal opremaID As String, ByVal pocetak As Date, ByVal ulazi As Collection, _
        ByVal parametri As Collection, Optional ByVal odgovorni As String = "", _
        Optional ByVal napomena As String = "") As String              ' SarzaID

' izlazi: Collection of Array(tipIzlaza, proizvodID, klasa, kalibracija, kg, tipJedinice, _
'                             tipKutije, brojKutija, tipKese, brojKesa, tezinaPaleteKg, brutoKg, lotBroj, napomena)
Public Function ZavrsiSarzu_TX(ByVal sarzaID As String, ByVal kraj As Date, _
        ByVal izlazi As Collection, ByVal parametri As Collection) As String   ' SarzaID

' modStorno:
Public Function StornoProcesSarza_TX(ByVal sarzaID As String) As Boolean
Public Function StornoProcesSarza(ByVal sarzaID As String) As Boolean

' read-model za mreze (1-bazirani 2D nizovi, obrazac GetPreradeForGrid / GetUtovariForGrid):
Public Function GetSarzeForGrid(Optional ByVal god As Long = 0) As Variant
Public Function GetLagerJediniceForGrid(Optional ByVal samoRaspolozive As Boolean = True) As Variant
Public Function GetSarzaUlaziForGrid(ByVal sarzaID As String) As Variant
Public Function GetSarzaIzlaziForGrid(ByVal sarzaID As String) As Variant
Public Function GetSarzaParametriForGrid(ByVal sarzaID As String) As Variant
```

### 6.3 `OtvoriSarzu_TX` — tok

1. Snapshot: `tblProcesSarze, tblProcesUlazi, tblProcesParametri, tblLagerJedinice,
   tblPaleta`.
2. Kapije 7401–7412 nad **svim** ulazima (mapa raspoloživog izračunata **jednom**;
   dupli ulaz iste jedinice u listi = 7407, ne sabiranje).
3. Kapacitet opreme: `Σ ulaza > KapacitetKg` → upozorenje u poruci povratka (ne blokada).
4. Materijalizacija `PALETA` ulaza koji nemaju LJ (`IzvorTip=PALETA`, `IzvorID=PaletaID`).
5. Header (`Status=OTVORENA`, `BrojSarze`), ulazi, parametri.
6. Za svaku paletu-poreklo: ako `RaspolozivoKg ≤ 0,01` → `Preradjeno=Da` (I8).
7. Commit; `Monitor_Event SARZA_OTVORENA`.

### 6.4 `ZavrsiSarzu_TX` — tok

1. Snapshot: `tblProcesSarze, tblProcesIzlazi, tblProcesParametri, tblLagerJedinice`.
2. Kapije 7420–7429; balans (D3) računat nad **stvarnim** ulazima iz tabele, ne nad
   parametrom.
3. LJ za svaki `GLAVNI`/`NUSPROIZVOD` (`IzvorTip=SARZA`, `IzvorID=SarzaID`,
   `BrojJedinice` sopstveni brojač, `DatumNastanka=kraj`, `StanicaID` šarže).
4. Izlazi (snimak proizvoda/klase/kalibracije), parametri, `DatumVremeKraj`,
   `Status=ZAVRSENA`.
5. Commit; `Monitor_Event SARZA_ZAVRSENA`.

Ekran nudi radnju „Knjiži razliku kao tehnološki gubitak" koja **dodaje red u korpu**
(`GUBITAK`, kg = razlika); writer ne dodaje ništa sam.

### 6.5 Storno matrica (posle B2)

| Entitet | Preduslov (fail-closed) | Efekat | Snapshot |
|---|---|---|---|
| Šarža `OTVORENA` | nema | ulazi + parametri + header `Stornirano`; paletama-poreklu `Preradjeno` se briše ako više nema aktivnih ulaza | sarze, ulazi, parametri, paleta |
| Šarža `ZAVRSENA` | nijedna izlazna LJ nema aktivan ulaz procesa / utovarnu stavku / blokadu (`ERR_STORNO_BASE+80`, poruka imenuje potomka) | + izlazi + izlazne LJ `Stornirano` | + izlazi, lagerJedinice |
| Legacy prerada (`StornoPrerada`) | postojeće (#248: utovar) **+** njena LJ nema aktivan ulaz procesa (`+81`) | postojeće + LJ `Stornirano` | + lagerJedinice |
| Paleta (`StornoPaleta`) | postojeće (`Preradjeno`) **+** `PaletaUProcesu=False` (`+82`) | postojeće + LJ porekla `Stornirano` ako postoji | + lagerJedinice |
| Utovar | postojeće (#248) | postojeće; stanje LJ se vraća **izvođenjem** (ništa se ne piše na LJ) | postojeće |
| Faktura | postojeće (#248) | postojeće | postojeće |

Kodovi `ERR_STORNO_BASE + 80..89` (poslednji zauzet na #248 je `+70`).

---

## 7) Prodajni stek na LJ ključu (Faza B1) — obim refaktora

Cilj: **isto ponašanje, jedan ključ.** Svaki test i sabotaža iz #248 ostaju i moraju
proći nad LJ ključem. Po modulu:

| Modul | Šta se menja | Ostaje |
|---|---|---|
| `modUtovar` | `UtovarenoPoPreradi` → **`UtovarenoPoJedinici`** (ključ `LagerJedinicaID`); `UtovarenoKgPrerade` → `UtovarenoKgJedinice`; `CreateUtovarCore`: stavka = `Array(lagerJedinicaID, kg, cena)`, kapije nad LJ (postoji tačno jednom, nije stornirana, ima `ProizvodID` sa `Prodajni=Da`, `kg ≤ RaspolozivoKg`), upis `LagerJedinicaID` **i** `PreradaID/BrojPrerade` (iz `IzvorID` kad je `IzvorTip=PRERADA`, inače prazno); `UtsPakovanja` čita pakovanje sa LJ; `GetUtovariForGrid` kolona „roba" = `LjOznaka`; `PrintUtovar`: Lot = `LjOznaka` (ili `LotBroj` ako postoji), Proizvod = `Proizvod.Naziv`, Rok = `LjRokTrajanja`, Bruto = LJ; `CreateFakturaIzUtovara`: FST nosi `LagerJedinicaID`; `AktivnihFstZaUtovar` nepromenjen | model B, cene, prevoz, plomba, migracija |
| `modFaktura` | `GetGPZaFakturisanjeForGrid`: izvor **`tblLagerJedinice`** (`Prodajni=Da`, nestornirane): `1 LJ | 2 Oznaka | 3 Proizvod | 4 Klasa | 5 Datum | 6 NaStanju | 7 Kutije | 8 Kese | 9 Dostupna | 10 BrojFakture`; `PrintFaktura` mapa proizvoda po LJ; `BrojacIdova` nad LJ | šablon fakture, `gpFaktura` grana |
| `modSEFMapper` | XOR `PrijemnicaID`/**`LagerJedinicaID`** po stavci; `gpKg` po LJ; `SellersItemIdentification = LagerJedinicaID`; naziv = `Proizvod.Naziv`; `ValidateGpUtovarZaSEF` kg mape po LJ; `clsSEFLine` + `lagerJedinicaID` (JSON snapshot dobija `"LagerJedinicaID"`, `"PreradaID"` ostaje radi starih snapshot-a) | `DeliveryDate`, kupac invariant |
| `modScrFakture` | lista `GOTOVA`: identitet = LJ (kolona `FK_GP_KOL_ID`), naslovi „Oznaka / Proizvod / Klasa"; korpa ključ LJ; `FkDodajGP(ljID, …)`; seam `Scr_FkKorpaTestDodajGP(ljID, …)`; lista `UTOVARI` kolona roba = oznaka | čipovi, hint, radnje, prevoz polja |
| `modIzvestaj` | `SledGpMape`: `prePoPaleti` → **`ljPoPaleti`** (paleta → LJ: preko `IzvorID` za materijalizovane, preko `tblPreradaStavka` za legacy prerade → njihove LJ); `SledUtovarPoPreradi` → `SledUtovarPoJedinici`; ključ dokaza `UtovarID|FakturaID|LagerJedinicaID`; mete (`ReportSledljivostMete`) `PRERADA` → LJ (štampa = paletni list LJ kad postoji, inače preradni list) | stanja `SLED_ST_*`, oznake, fail-closed pravila, NEPOTPUNI klase |
| `modStorno` | `StornoPrerada`: kapija utovara preko `UtovarenoKgJedinice(LJ)`; LJ `Stornirano`; `ReleaseUtovarFromFaktura` nepromenjen | ostalo |
| `modSetup` | `BackfillUtovariIzGPFaktura` upisuje i `LagerJedinicaID` | |
| `modPrint` | `FillUtovarSablon` čita LJ polja | LAYOUT_VER +1 samo ako se raspored menja (ne mora) |
| Testovi | 160–163 prevedeni na LJ (fixture daje LJ redove); +1 test: `T_LJ_MaterijalizacijaLegacy` (prerada → LJ 1:1, `KgPocetno`, stornirana → stornirana, utovar/faktura stavke dobile LJ, prerada bez tipa → nedostupna) | tvrdnje |
| Sabotaže | 20 sabotaža sa `utovar-gp-*`, `sef-gp-*`, `sledljivost-gp-*`, `storno-gp-*`, `faktura-gp-*`, `fakture-gp-*` prefiksom: sidra premeštena na nove nazive funkcija; svaka mora i dalje da obara **isti** test | katalog |

Napomena o SEF-u: `SellersItemIdentification` za legacy lotove menja vrednost sa
`PRE-xxxxx` na `LJ-xxxxx` **samo za buduće fakture**; poslate fakture su nepromenljive
(snapshot). Zabeležiti u `docs/SEF_LIFECYCLE_MANUAL.md`.

---

## 8) Sledljivost kao graf (Faza C2)

### 8.1 Graf

- Čvorovi: LJ (sve), palete bez LJ (listovi porekla), legacy prerade (kroz svoju LJ),
  utovari, fakture.
- Ivice (jedan prolaz, S5): `paleta → LJ` (materijalizacija), `LJ → šarža`
  (`tblProcesUlazi`), `šarža → LJ` (`tblProcesIzlazi`), `paleta → prerada-LJ`
  (`tblPreradaStavka`, legacy), `LJ → utovar` (`tblUtovarStavke`), `utovar → faktura`.
- `modSledljivost.ProcesGraf()` vraća dve mape `parents(id) → dict`, `children(id) → dict`
  (id = `LJ-…`, `PAL-…`, `SRZ-…`, `UT-…`, `FAK-…`); `LjPreci(id)`, `LjPotomci(id)` = BFS
  sa zaštitom od ciklusa (ciklus = neusaglašenost, prijavljuje se, ne petlja).
- Fail-closed: dupli ID u bilo kojoj tabeli → čvor se **ne** povezuje i red dobija
  oznaku `SLED_OZN_VEZA`; stornirane ivice se ne crtaju.

### 8.2 LANAC (ekran Sledljivost, lista `LANAC`)

- Kolona 28 „Palete" nepromenjena. Kolona 29 „Pal. gotovog proizv." postaje
  multi-vrednosna: `SORT 12/2026 -> LJ 3/2026, LJ 4/2026` ili `3 sarze / 5 LJ` kad je
  više; kolona 30 „Stanje" dobija `u procesu` (`SLED_ST_U_PROCESU`, prioritet između
  `preradjeno` i `u hladnjaci`) i `preradjeno` važi i za 2.0 izlaze.
- SearchRefs (kol. 27) += brojevi šarži i oznake LJ.

### 8.3 Nova lista `GENEALOGIJA` na ekranu Sledljivost

- Ulaz: polje u zoni (oznaka LJ / broj šarže / broj prerade / broj palete / broj
  utovara / broj fakture) — razrešenje po istoj listi kandidata kao pretraga LANCA.
- Mreža: `Smer (unazad/unapred) | Nivo | Vrsta (paleta/šarža/LJ/utovar/faktura) | Broj |
  Proizvod | Klasa | Kg | Datum | Partner (kooperant/kupac) | Oznaka`.
- Radnje: „Dokument" (štampa karike: paletni list / procesni list / paletni list LJ /
  utovarna lista / faktura), „Recall lista (PDF)" (sve karike unapred od izabranog
  čvora, kroz postojeći `SledljivostSablon`).
- Test: `T_Sled_Genealogija` (split 1→3, merge 3→1, dva nivoa, storno ivica se ne
  crta, dupli ID fail-closed).

---

## 9) Ekran `PROIZVODNJA` (Faza C1, `modScrProizvodnja.bas`)

### 9.1 Registracija

`modUiScreens.ScrRows`: `"PROIZVODNJA|modScrProizvodnja|OTKUI_NAV_PROIZVODNJA|" & IC_PROIZ &
"|OPERACIJE|" & OBL_PALETE` (posle `PALETE`). `IC_PROIZ` = novi MDL2 kod u `modOtkupUI`.
Oblast prava `OBL_PALETE` — `modAuth.OblastiList` se ne dira. Kasno vezivanje: klijent
bez modula vidi prigušenu stavku.

`Scr_Meta`: `kljuc=PROIZVODNJA|naslov=OTKUI_NAV_PROIZVODNJA|sub=OTKUI_SCRPRZ_SUB|lista=OTKUI_SCRPRZ_LISTA|oblik=lista|upis=da`.

### 9.2 Liste (`Scr_Liste`)

| Ključ | Mreža (kolone) | Zona | Radnje nad redom | Čipovi |
|---|---|---|---|---|
| `SARZE` | `Broj | Tip | Oprema | Početak | Kraj | Ulaz kg | Glavni kg | Nus kg | Otpad kg | Gubitak kg | Razlika | Status | [SarzaID skriven]` | KPI: otvorene / završene u periodu / ukupan ulaz / ukupan glavni | `Procesni list`, `Otvori u IZLAZ` (samo OTVORENA), `Storniraj` | `sve / otvorene / zavrsene / stornirane` |
| `ULAZ` | `☐ | Oznaka | Vrsta | Proizvod | Klasa | Poreklo (šarža/prijemnica) | Raspoloživo kg | kg za ulaz | Datum | [ID skriven] | [IzvorTip skriven]` — izvor: `GetLagerJediniceForGrid` ∪ palete bez LJ | polja: `scrPrzTip` (cmb), `scrPrzStanica` (cmb), `scrPrzOprema` (cmb, filtrira po stanici i tipu), `scrPrzPocetak` (txt datum-vreme), `scrPrzOdgovorni`, `scrPrzNap`; blok parametara (`scrPrzParVremeUlaz`, `…VremeIzlaz`, `…TempRobeUlaz`, `…TempRobeIzlaz`, `…CiljnaTemp`, `…Brix`) — vidljivi po `ObavezniParametri` tipa; KPI: `IZABRANO n | ULAZ kg`; dugme **`scrPrzOtvori`** | `Označi / Odznači`, polje kg (prazno = celo raspoloživo; obrazac „Kol. za utovar") | `sve / sveže / smrznuto / GP / u procesu` |
| `IZLAZ` | korpa izlaza aktivne OTVORENE šarže: `Tip izlaza | Proizvod | Klasa | Kalibracija | Kg | Tip jedinice | Kutije | Kese | Bruto | Lot | [idx]` | polja: `scrPrzIzTip` (cmb), `scrPrzIzProizvod` (cmb), `scrPrzIzKlasa`, `scrPrzIzKalib`, `scrPrzIzKg`, `scrPrzIzTipJed` (cmb), `scrPrzIzTipKut`/`Kut`, `scrPrzIzTipKes`/`Kes`, `scrPrzIzTezPal`, `scrPrzIzBruto`, `scrPrzIzLot`; dugme `scrPrzDodaj`; **KPI balans:** `ULAZ | GLAVNI | NUSPROIZVOD | OTPAD | GUBITAK | NERASPOREDJENO`; dugme `scrPrzGubitak` („Knjiži razliku kao gubitak"); dugme **`scrPrzZavrsi`** (ugašeno dok `|NERASPOREDJENO| > tolerancija` ili nema GLAVNI) | `Ukloni red` | — |
| `LAGER` | `Oznaka | Vrsta | Proizvod | Klasa | Kalib. | Raspoloživo | Fizičko | U procesu | Utovareno | Datum | Rok | Poreklo | Stanica | [ID]` | KPI: jedinice / kg raspoloživo / kg u procesu | `Paletni list LJ`, `Genealogija` (skok na Sledljivost sa prefilom) | `sve / raspolozive / potrosene / GP / sveze` |

Stanje modula: `mLista`, `mSarzaID` (aktivna), `mUlazi` (dict ID → kg), `mIzlazi`
(Collection), `mParametri`, `mCombosPunjeni`, `mStep`. Kombo ponude: tipovi procesa
(aktivni), stanice (`JeHladnjaca`), oprema, proizvodi (aktivni), tipovi jedinica,
kutije/kese (`GetKutijeOptions`/`GetKeseOptions` — postojeće).

Ekran **ne računa i ne upisuje**: balans u KPI računa `modProizvodnja.BalansKorpe(ulazKg,
izlazi)` (ista aritmetika kao writer, jedna implementacija). `Scr_Save` ruta: na `ULAZ`
→ `OtvoriSarzu_TX`; na `IZLAZ` → `ZavrsiSarzu_TX`; posle uspeha `OutputProcesList`,
`Scr_ResetCache`, toast.

Test seam-ovi (gejtovani `IsTestMode`): `Scr_PrzTestSet(kljucListe)`,
`Scr_PrzUlazTestDodaj(id, kg)`, `Scr_PrzIzlazTestDodaj(...)`, `Scr_PrzBalansTest()`.

### 9.3 Poruke (`modPoruke.UpsertPoruke`, prefiks `OTKUI_PRZ_`)

Navigacija/naslovi: `OTKUI_NAV_PROIZVODNJA`, `OTKUI_SCRPRZ_SUB`, `OTKUI_SCRPRZ_LISTA`,
`OTKUI_PRZ_L_SARZE/ULAZ/IZLAZ/LAGER`. Hederi mreža: `OTKUI_HPRZ_*` (po koloni).
Polja: `OTKUI_PRZ_TIP`, `…_STANICA`, `…_OPREMA`, `…_POCETAK`, `…_ODG`, `…_NAP`,
`…_PAR_*`, `…_IZ_*`. Validacije: `OTKUI_PRZ_V_TIP`, `…_V_ULAZ_NEMA`, `…_V_KG`,
`…_V_RASP`, `…_V_FORMA`, `…_V_PARAM`, `…_V_GLAVNI`, `…_V_BALANS`. Potvrde/toasts:
`OTKUI_PRZ_ASK_OTVORI`, `…_ASK_ZAVRSI`, `…_ASK_STORNO`, `…_OTVORENA`, `…_ZAVRSENA`,
`…_ERR`. Čipovi: `OTKUI_PRZ_CIP_*`. Radnje: `OTKUI_BTN_PRZ_*`.

### 9.4 Izmene ostalih ekrana

| Ekran | Izmena | Faza |
|---|---|---|
| Palete (`modScrPalete`) | lista `NOVAPRERADA`: „Preradi izabrane" → `OtvoriSarzu_TX(PRERADA_LEGACY, cele palete)` + `ZavrsiSarzu_TX(1 GLAVNI, proizvod iz `scrPreGP`, pakovanje iz polja; razlika ulaz−izlaz → `GUBITAK` red automatski uz potvrdu)`; lista `PRERADE` ostaje pregled legacy prerada + šarži tipa `PRERADA_LEGACY` (kolona „Izvor") | C1 (D5) |
| Fakturisanje (`modScrFakture`) | lista `GOTOVA` na LJ (§7) | B1 |
| Sledljivost (`modScrSledljivost`) | lista `GENEALOGIJA`; LANAC kolone (§8) | C2 |
| Storno centar (`modScrStorno`) | ne dira se: storno šarže/LJ pripada ekranu Proizvodnja (kao paleta/prerada ekranu Palete — katalog, red „Storno palete i prerade") | — |
| Matični podaci | tri nove sekcije u `MaticniSekcije` + `frmStammdaten` `Select Case` (`TBL_TIPOVI_PROCESA`, `TBL_PROIZVODI`, `TBL_OPREMA`); samo code-behind, bez `.frx` | A |

---

## 10) Dokumenti (`modPrint`, Faza C1)

| Dokument | List | Named ranges | Sadržaj | Režim |
|---|---|---|---|---|
| **Procesni list** | `ProcesSablon` (`WS_PROCES_SABLON`), `EnsureProcesSablon`/`FillProcesSablon`, `LAYOUT_VER="1"` u `H1` | `PrzBroj, PrzTip, PrzDatum, PrzOprema, PrzStanica, PrzOdg, PrzUlazTab, PrzIzlazTab, PrzUlazUk, PrzGlavni, PrzNus, PrzOtpad, PrzGubitak, PrzRazlika, PrzRandman, PrzRecovery, PrzParamTab, PrzNap, PrzPotpis` | zaglavlje; ulazi `Rb | Oznaka | Proizvod | Klasa | Kg | Poreklo`; izlazi `Rb | Oznaka | Proizvod | Klasa | Kalib. | Tip | Kg | Pakovanje | Lot`; balans; parametri `Kljuc | Vrednost | Jedinica`; kontrole (D) | `PROCES_PRINT_MODE` (default `PDF`), folder `PDF_DIR_PROCES = "Procesni listovi"` |
| List zamrzavanja | isti obrazac; naslov „LIST ZAMRZAVANJA" kad je tip `ZAMRZAVANJE`; blok parametara nosi tunel, vremena, temperature, trajanje (izvedeno) | | | isti ključ |
| **Paletni list LJ** | `LjSablon` (`WS_LJ_SABLON`), `EnsureLjSablon`/`FillLjSablon` | `LjOznaka, LjProizvod, LjKlasa, LjKalib, LjKg, LjPak, LjLot, LjDatum, LjRok, LjPorekloTab, LjStanica, LjPotpis` | jedinica + poreklo: šarža i ulazne jedinice/palete sa kg (`LjPreci` nivo 1; opciono pun lanac kad `PRERADA_SLEDLJIVOST_DETALJ`) | `LJ_PRINT_MODE` (default `PDF`), folder `Paletni listovi LJ` |
| Utovarna lista | postojeći `UtovarSablon` čita LJ (§7) | | | postojeći |

Auto-izlaz posle `ZavrsiSarzu_TX`: `OutputProcesList` (best-effort, `On Error Resume
Next`, ne obara potvrdu) + po jedan paletni list LJ za svaki `GLAVNI` izlaz tipa
`PALETA` (kao `PaletniListOutputClosed`).

---

## 11) Konfiguracija (`modPodesavanja`)

| Ključ | Grupa | Tip | Default | Svrha |
|---|---|---|---|---|
| `PROCES_BALANS_TOLERANCIJA_KG` | **Proizvodnja** (nova grupa) | broj | `0,5` | D3; sanity 0–50 |
| `PROCES_PRINT_MODE` | Štampa | `list:PDF;PRINT;PREVIEW;OFF` | `PDF` | procesni list |
| `LJ_PRINT_MODE` | Štampa | isto | `PDF` | paletni list LJ |
| `PROIZVODNJA_AKTIVNA` | Proizvodnja | `DA/NE` | `NE` | prikaz ekrana i wrapper-a `NOVAPRERADA` (D5) — klijent bez modula Hladnjača/Proizvodnja ne vidi ekran; **ne** gejtuje B1 (jedan ključ nema prekidač) |

Čitanje kroz `ConfigFlag` / `GetConfigValue` obrazac (`IsProizvodnjaAktivna()` u
`modConfig`, kao `IsPaletiranje`).

---

## 12) Integritet (`modIntegritet`, novi blok „P")

| ID | Provera | Kolone izlaza |
|---|---|---|
| P1 | `ZAVRSENA` šarža: balans van tolerancije, ili bez `GLAVNI` izlaza, ili izlaz `GLAVNI/NUS` bez LJ | `SarzaID, Broj, UlazKg, IzlazKg, Razlika, Razlog` |
| P2 | jedinica sa `Raspolozivo < −0,01` (prekomerna potrošnja: ulazi + utovar > fizičko) | `LJ, Oznaka, Fizicko, UProcesu, Utovareno, Razlika` |
| P3 | stornirana jedinica/šarža sa aktivnim potomkom; paleta prevezana/korigovana dok je u procesu (`Istorija` zapis posle prvog ulaza) | `ID, Vrsta, Potomak, Razlog` |
| P4 | prerada bez LJ / sa dve LJ / `KgPocetno ≠ NetoIzlazKg` / LJ bez `ProizvodID` (tip GP nepoznat) | `PreradaID, Broj, LJ, Razlog` |
| P5 | utovarna ili fakturna GP stavka bez `LagerJedinicaID` ili sa nepostojećom LJ | `StavkaID, Tabela, LJ, Razlog` |
| P6 | LJ bez `StanicaID`; šarža sa opremom druge stanice; tip procesa neaktivan a korišćen | `ID, Razlog` |
| A5 | ostaje za legacy prerade; **preskače** redove sa `SarzaID` (posle D) | — |

`docs/INTEGRITET_PROVERE.md` dobija sekciju P; `RunProductionHealthCheck` ih uključuje.

---

## 13) Migracija i kompatibilnost

### 13.1 `EnsureProizvodnjaSchema` (Alt+F8; zove ga i `EnsurePaletniListSchema`)

1. `EnsureDataTable` × 8 (kolone tačno po §4; postojeća tabela = dopuna kolona).
2. `EnsureColumnOnTable`: `tblPrerada.LagerJedinicaID`, `tblUtovarStavke.LagerJedinicaID`,
   `tblFakturaStavke.LagerJedinicaID`.
3. Seed tipova procesa; seed proizvoda (§4.2).
4. **Materijalizacija legacy prerada** (§4.4) + upis `tblPrerada.LagerJedinicaID`.
5. Backfill `LagerJedinicaID` na utovarnim/fakturnim stavkama (§4.9).
6. Audit kolone na novim tabelama; `LogSetup` sažetak (n tabela, n LJ, n stavki).

Sve idempotentno (ponovni poziv = 0 promena). Test: `T_Prz_SemaIdempotentna`.

### 13.2 Self-heal (`EnsureRuntimeSchema`)

Koraci 1, 2, 4, 5, 6 (deterministički, jeftini: jedan prolaz po tabeli). Seed tipova (3)
takođe (ključan za writer). Razlog: klijent koji dobije kod self-update-om mora moći da
proda GP lot **odmah**, a to od B1 traži LJ — isti argument kao za utovar šemu (#248,
katalog §25.9/§25.15).

### 13.3 Faza D — `BackfillSarzeIzLegacyPrerada` (Alt+F8, **opciono**, idempotentno)

Za svaku aktivnu preradu bez `SarzaID`: šarža `PRERADA_LEGACY` (`Status=ZAVRSENA`,
početak = kraj = `Datum`), ulazi = `tblPreradaStavka` (materijalizovane palete, cele,
`KgUlaz=NetoKg` stavke), izlaz `GLAVNI` = postojeća LJ prerade, razlika `NetoUlaz −
NetoIzlaz` → `GUBITAK` sa napomenom `LEGACY (neevidentiran otpad)`; `tblPrerada.SarzaID`
= nova šarža. Ništa se ne izmišlja preko onoga što je evidentirano; **ne ide** u
self-heal (kao `BackfillUtovariIzGPFaktura`). Posle D: `SavePrerada_TX` se gasi i u
legacy `frmPalete` (dugme zove wrapper).

### 13.4 Ono što se **ne** menja

`tblPaleta`/`tblPaletaStavka` šema i paletizacija; `PaletizePrijemnica`; `tblCenovnik`;
`tblKulture`; `tblVrstaGotovihProizvoda` (ostaje izvor `RokMeseci`); SEF snapshot-i
poslatih faktura; PWA/GAS sync (ove tabele se ne sinhronizuju, kao ni palete danas).

### 13.5 Rollback

Šema je append-only (nikad se ne briše tabela ni kolona). Kod svakog PR-a se može
vratiti (`git revert`) bez štete po podatke: LJ redovi za legacy su pokazivači;
`LagerJedinicaID` kolone ostaju prazne/neiskorišćene za stari kod (koji ih ne čita).
Jedini nepovratni podatak su šarže i LJ tipa `SARZA` nastale u 2.0 — posle B2 vraćanje
koda ostavlja te redove nevidljivim, pa se **B2 ne vraća** nego popravlja unapred.

---

## 14) Verifikacija

### 14.1 Nivoi

| Nivo | Šta | Kada |
|---|---|---|
| FAST | `python tools\vba_check.py` (ASCII, deklaracije, duplikati, `PORUKA`, `STORNO_REGISTAR`, `ZAKLONJENO`) | posle svake VBA izmene, i u CI |
| TARGETED | `python tools\run_vba.py --suite RunProizvodnjaTestSuite` (+ `RunPaleteTestSuite`, `RunStornoTestSuite`, `RunAllTests` po fazi) | po PR-u |
| FULL | `python tools\run_vba.py` | pred merge svakog PR-a i pred release |
| Dokaz | `python tools/dokaz.py proizvodnja utovar-gp sef-gp sledljivost-gp storno-gp` | pred merge B1, B2, C2 |
| Compile | `Alt+F11 → Debug → Compile VBAProject` — ručna kapija | pred merge |
| Smoke | numerisana checklista operatera po PR-u (§15) | pred merge |

U Linux/web sesiji **ništa osim FAST nije izvršivo**; svaka izmena ponašanja se
prijavljuje kao neverifikovana dok Windows run ne prođe.

### 14.2 Nova suite `RunProizvodnjaTestSuite` (`modTestProizvodnja.bas`, `gate: True`, `dialogs: True`, `default: True`)

Obrazac `modTestPalete`: `SeedMasterData`, `TstAppend`, `Chk/ChkEq/ChkEqD`,
`ReportResults`, rollback po testu. Testovi (ime = tvrdnja):

| # | Test | Tvrdnje |
|---|---|---|
| T01 | `T01_SemaIdempotentna` | drugi `EnsureProizvodnjaSchema` = 0 novih redova/kolona; seed tipova i proizvoda prisutan; `STORNO_TABELE` registrovane |
| T02 | `T02_MaterijalizacijaLegacyPrerade` | 3 prerade (aktivna sa tipom, aktivna bez tipa, stornirana) → 3 LJ; `KgPocetno`, `Stornirano`, `ProizvodID` prazan za bez-tipa; `tblPrerada.LagerJedinicaID` popunjen; utovarna/fakturna stavka dobila LJ |
| T03 | `T03_RaspolozivoJednaFunkcija` | LJ 1.000 kg, ulaz 600, utovar 100 → 300; paleta poreklo čita živi `NetoKg`; negativno = P2 |
| T04 | `T04_OtvoriKapije` | svaka kapija 7401–7412 pada **po kodu** (12 tvrdnji); nijedan red nije upisan posle pada (rollback dokaz kroz brojanje redova) |
| T05 | `T05_OtvoriUlazParcijalan` | paleta 1.000 kg, ulaz 600 → LJ materijalizovana, `Preradjeno` prazno, raspoloživo 400; drugi ulaz 400 → `Preradjeno=Da` |
| T06 | `T06_ZavrsiSplit1naN` | 1.000 → 520/300/150 GLAVNI+NUS, 20 OTPAD, 10 GUBITAK: 3 LJ, balans 0, `BalansSarze` randman 0,52 / recovery 0,99 |
| T07 | `T07_ZavrsiMergeNna1` | 400+350+250 → 1 LJ 1.000; sve tri ulazne jedinice potrošene |
| T08 | `T08_ZavrsiBalansTolerancija` | razlika 0,6 kg uz toleranciju 0,5 → 7427; 0,4 → prolazi i `NERASPOREDJENO=0,4` u `BalansSarze`; tolerancija se čita iz config-a (promena config-a menja ishod) |
| T09 | `T09_ZavrsiKapije` | 7420–7429 po kodu; bez GLAVNI → 7423; rollback: namerna greška posle upisa 2. LJ vraća sve (obrazac `mTestFailPosleRelease`) |
| T10 | `T10_StornoOtvorene` | vraća raspoloživo i `Preradjeno`; ulazi stornirani |
| T11 | `T11_StornoZavrseneBezPotomka` | izlazne LJ stornirane; ulazi vraćeni |
| T12 | `T12_StornoBlokiraPotomak` | izlazna LJ u drugoj šarži → `+80`; izlazna LJ na utovaru → `+80`; posle storna potomka prolazi |
| T13 | `T13_PaletaUProcesuKapija` | `ReassignPaleteToPrijemnica_TX` / `Detach` / `Adjust` nad paletom sa aktivnim ulazom → greška; posle storna šarže prolazi |
| T14 | `T14_LegacyStornoUciLJ` | `StornoPrerada` nad preradom čija je LJ u procesu → `+81`; `StornoPaleta` nad paletom u procesu → `+82` |
| T15 | `T15_IntegritetP1doP6` | svaka provera nalazi tačno svoj sabotirani red; čist skup = 0 nalaza |
| T16 | `T16_TipProcesaPravila` | forma nedozvoljena → 7410; oprema obavezna → 7403; obavezan parametar → 7411; kapacitet → upozorenje bez blokade |
| T17 | `T17_LjOznakaIRok` | `PRE 51/2026`, `PAL 31/2026`, `LJ 1/2026`; rok = datum + VGP.RokMeseci, fallback globalni, Empty bez ijednog |

### 14.3 `RunAllTests` (`modTest`) — ekran i sledljivost

| # | Test | Faza |
|---|---|---|
| 160–163 | postojeći GP testovi prevedeni na LJ ključ | B1 |
| 164 | `T_LJ_ProdajniStekNaLJ` — grid GOTOVA iz LJ, korpa po LJ, utovar+faktura+SEF snapshot nose LJ, legacy lot i 2.0 lot na istoj fakturi | B1 |
| 165 | `T_Prz_UgovorEkrana` — `Scr_Meta`, 4 liste, radnje po listi, čipovi, identitet skriven i poslednji | C1 |
| 166 | `T_Prz_KorpaIBalansKPI` — dodaj/ukloni izlaz, KPI = `BalansKorpe`, „Završi" ugašeno van tolerancije, „Knjiži gubitak" dodaje red | C1 |
| 167 | `T_ZonaPrz_PoljaIRaspored` — parametarska polja se pale po tipu; `scr` prefiks; `LayoutFieldInner` | C1 |
| 168 | `T_Prz_WrapperNovaPrerada` — Palete/NOVAPRERADA kroz 2.0: šarža `PRERADA_LEGACY`, 1 GLAVNI, GUBITAK = razlika | C1 |
| 169 | `T_Sled_Genealogija` | C2 |
| 170 | `T_Sled_LanacMultiSarza` — kolona 29 multi, stanje `u procesu`, refs | C2 |

### 14.4 Sabotaže (`tools/sabotaza.py`)

Obavezan dvosmerni dokaz (kritične invarijante + izmenjen checker/test):

| Sabotaža | Šta kvari | Obara |
|---|---|---|
| `proizvodnja-balans-placebo` | tolerancija ignorisana u `ZavrsiSarzu` | T08 |
| `proizvodnja-balans-bez-gubitka` | `GUBITAK` izostavljen iz Σ izlaza | T06, P1 |
| `proizvodnja-raspolozivo-bez-utovara` | utovar izostavljen iz odbitka | T03 |
| `proizvodnja-raspolozivo-bez-procesa` | ulazi izostavljeni iz odbitka | T03, T05 |
| `proizvodnja-storno-potomak-prolazi` | kapija `+80` uklonjena | T12 |
| `proizvodnja-paleta-u-procesu-prolazi` | `PaletaUProcesu` vraća False | T13 |
| `proizvodnja-preradjeno-uvek` | `Preradjeno=Da` i na parcijalu | T05 |
| `proizvodnja-rollback-bez-lj` | snapshot `tblLagerJedinice` uklonjen iz `ZavrsiSarzu_TX` | T09 |
| `proizvodnja-materijalizacija-bez-storna` | stornirana prerada dobija aktivnu LJ | T02, P4 |
| `proizvodnja-forma-placebo` | `DozvoljenaUlaznaForma` ne proverava | T16 |
| `utovar-gp-kljuc-prerada` | `CreateUtovarCore` čita stanje po `PreradaID` umesto LJ | 164 |
| `sef-gp-identitet-prerada` | `SellersItemIdentification` vraćen na `PreradaID` | 164 |
| `sledljivost-graf-storno-ivica` | stornirani ulaz crta ivicu | 169 |
| `sledljivost-graf-dupli-id` | dupli LJ ID se povezuje | 169 |
| `sledljivost-lanac-multi-prvi` | kolona 29 uzima samo prvu šaržu | 170 |
| `prz-ekran-zavrsi-uvek` | dugme Završi upaljeno van tolerancije | 166 |

Postojećih 20 GP sabotaža ostaju i moraju obarati **isti** test posle B1.

### 14.5 Fixture (`tools/make_fixture.py`, `tests/schema_donor.json`)

- `ENSURE_TABLES` += 8 tabela sa kolonama; `ENSURE_COLS` += tri `LagerJedinicaID`.
- Generator emituje LJ redove za sve fixture prerade (isti algoritam kao
  materijalizacija — implementiran u Python-u **i** proveren testom T02 nad VBA
  materijalizacijom; razilaženje = pad).
- Novi lanci: `SRZ-SLED-S` (split 1→3, sa GUBITAK), `SRZ-SLED-M` (merge 3→1),
  `SRZ-SLED-O` (otvorena, roba u tunelu), `SRZ-SLED-P` (parcijalan ulaz sa palete),
  `LJ-SLED-U` (2.0 lot na utovaru + fakturi), `SRZ-SLED-X` (stornirana).
- Config u fixture-u: `PROCES_BALANS_TOLERANCIJA_KG=0,5`, `PROCES_PRINT_MODE=OFF`,
  `LJ_PRINT_MODE=OFF`, `PROIZVODNJA_AKTIVNA=DA`.
- Potpis se regeneriše; `run_vba` staje bez važećeg potpisa (postojeći mehanizam).

### 14.6 CI

`vba_check` + `who_writes --check` (mapa dobija `modProizvodnja` kao pisca
`tblLagerJedinice`, `tblProcesSarze/Ulazi/Izlazi/Parametri`, `tblPaleta`, `tblPrerada`;
`modUtovar`/`modFaktura` kao pisce `LagerJedinicaID` kolona — regenerisati u svakom PR-u).

---

## 15) Plan isporuke po PR-ovima

Redosled je strog (svaki PR gradi na prethodnom u `main`-u). Svaki PR: jedna grana,
rebase na `main` pre push-a (Opcija 3), release notes + katalog §26 dopuna, smoke
checklista u opisu PR-a. `.claude/` izmene (npr. `rules/proizvodnja.md`) idu u **zaseban
process PR** posle PR-A.

| PR | Naslov | Sadržaj | Fajlovi | Preduslov | Definicija gotovog |
|---|---|---|---|---|---|
| **PR-0** | docs: model i plan | ovaj dokument | `docs/` | — | ✔ (ova grana) |
| **PR-A** | Prerada 2.0 — šema, matični podaci, lager jedinica za legacy | §4 tabele/kolone/registri; seed; materijalizacija + backfill (§13.1, 13.2); `SavePrerada_TX` pravi LJ za novu preradu; `StornoPrerada` stornira LJ; matični unos 3 sekcije; `IsProizvodnjaAktivna`; P4, P5, P6; `modTestProizvodnja` T01, T02, T17; fixture `ENSURE_TABLES` | `modConfig`, `modSetup`, `modSchemaGuard`, `modProizvodnja` (novo), `modPaletniList`, `modStorno`, `modIntegritet`, `modMaticniLookups`, `frmStammdaten` (code), `modPodesavanja`, `modTestProizvodnja` (novo), `tools/make_fixture.py`, `tools/run_vba.py`, `tests/schema_donor.json`, `docs/` | #247, #248 u `main` | FAST čist; `RunProizvodnjaTestSuite` 3/0; `RunAllTests` 163/0 i `RunPaleteTestSuite` bez promene; dokaz `proizvodnja-materijalizacija-bez-storna`; compile; smoke A |
| **PR-B1** | Prodajni stek na `LagerJedinicaID` | §7 u celosti; test 164; sabotaže `utovar-gp-kljuc-prerada`, `sef-gp-identitet-prerada`; 20 GP sabotaža premeštena sidra | `modUtovar`, `modFaktura`, `modSEFMapper`, `clsSEFLine`, `modScrFakture`, `modIzvestaj`, `modStorno`, `modSetup`, `modPrint`, `modTest`, `tools/sabotaza.py`, `tools/make_fixture.py`, `docs/SEF_LIFECYCLE_MANUAL.md` | PR-A | `RunAllTests` 164/0; `RunSEFTestSuite`, `RunFakturaSmokeSuite`, `RunStornoTestSuite` zeleni; pun dokaz `utovar-gp sef-gp sledljivost-gp storno-gp faktura-gp fakture-gp` (22/22); compile; smoke B1 (legacy lot i dalje se prodaje, štampa, šalje na SEF; storno lanac isti) |
| **PR-B2** | Procesni writer-i i storno šarže | §5, §6, §6.5; `PaletaUProcesu` u tri writer-a `modPaletniList`; utovar „na stanju" odbija procesnu potrošnju (već kroz `RaspolozivoPoJedinici`); P1–P3; T03–T16; 8 sabotaža `proizvodnja-*` | `modProizvodnja`, `modStorno`, `modPaletniList`, `modUtovar` (jedna funkcija stanja), `modIntegritet`, `modTestProizvodnja`, `tools/sabotaza.py`, `tools/make_fixture.py` (lanci `SRZ-SLED-*`) | PR-B1 | `RunProizvodnjaTestSuite` 17/0; `RunPaleteTestSuite` + T13; `RunStornoTestSuite`; dokaz `proizvodnja` 10/10; compile; **nema ekrana** — smoke kroz `Alt+F8` test-rutinu nad fixture-om |
| **PR-C1** | Ekran Proizvodnja + procesni list + paletni list LJ | §9 (osim GENEALOGIJA), §10, §11; wrapper `NOVAPRERADA` (D5); poruke; testovi 165–168; sabotaža `prz-ekran-zavrsi-uvek` | `modScrProizvodnja` (novo), `modUiScreens`, `modOtkupUI` (ikonica), `modPoruke`, `modPrint`, `modPodesavanja`, `modConfig`, `modScrPalete`, `modTest`, `docs/UI_MIGRACIJA_KATALOG.md` §26 | PR-B2 | `RunAllTests` 168/0; compile; **smoke C1** (§15.1) nad pravim podacima operatera; `PROIZVODNJA_AKTIVNA=NE` sakriva ekran i wrapper |
| **PR-C2** | Sledljivost kao graf | §8; testovi 169–170; 3 sabotaže `sledljivost-graf-*`, `sledljivost-lanac-multi-prvi` | `modSledljivost`, `modIzvestaj`, `modScrSledljivost`, `modPrint` (recall lista), `modPoruke`, `modTest` | PR-C1 | `RunAllTests` 170/0; dokaz `sledljivost sledljivost-gp`; smoke C2 |
| **PR-D** | Migracija legacy prerada, blokade, kontrole | §13.3; `tblBlokadeRobe` (`BlokadaID | LagerJedinicaID | Kg | Razlog | DatumOd | DatumDo | IzvorTip | IzvorID | Odobrio | Stornirano`) u `RaspolozivoPoJedinici`; `tblProcesKontrole` (`KontrolaID | SarzaID | TipKontrole | DatumVreme | Vrednost | Jedinica | LimitMin | LimitMax | Rezultat | Radnik | Stornirano`) + `METAL_DETEKTOR FAIL → blokada`; lista `KONTROLE` na ekranu; A5 skip; gašenje `SavePrerada_TX` u `frmPalete` | `modProizvodnja`, `modSetup`, `modConfig`, `modSchemaGuard`, `modScrProizvodnja`, `frmPalete` (code), `modPaletniList`, `modIntegritet`, `modPrint` (kontrole na procesnom listu), testovi | PR-C2 | suite + testovi blokade/kontrola; migracija nad kopijom pravih podataka pokazana operateru pre pokretanja |
| **PR-E** | Lokacije | `tblKomore`, `tblPozicije`, `tblLagerLokacija` (istorija), `TransferLager_TX`, Transfer list, kolona lokacije na `LAGER` listi i paletnom listu LJ | `modLokacije` (novo), `modScrProizvodnja`, `modPrint`, `modSetup`, … | PR-D | suite; smoke |
| **F** | 2027 obim | smene/radnici/produktivnost; senzori (odluka 67); `tblKalibracije`, `tblSpecifikacije`, `tblLotovi` | — | — | zaseban plan |

Procena u revizionim krugovima (iskustvo #248 = 11 krugova za GP prodaju):
A ≈ 2, B1 ≈ 3, B2 ≈ 3, C1 ≈ 3–4, C2 ≈ 2, D ≈ 2, E ≈ 2. Kritični put: A → B1 → B2 → C1.
B1 i B2 se ne mogu paralelizovati (B2 zavisi od jedne funkcije stanja iz B1); C1 i C2
mogu ići paralelno na različitim modulima posle B2, uz zajednički rebase.

### 15.1 Smoke checklista PR-C1 (operater, klik po klik)

1. `Alt+F11 → Compile` bez greške; `Alt+F8 → EnsureProizvodnjaSchema` → poruka sa
   brojem tabela i materijalizovanih LJ; ponovni poziv → 0 promena.
2. Podešavanja → grupa Proizvodnja: `PROIZVODNJA_AKTIVNA=DA`; sidebar OPERACIJE
   pokazuje „Proizvodnja".
3. Matični podaci: tipovi procesa (12 seed), proizvodi (iz kultura + vrsta GP),
   oprema (uneti „Tunel 1", stanica hladnjača).
4. Proizvodnja / ULAZ: mreža pokazuje palete iz lagera sa raspoloživim kg; označiti dve
   palete, kod jedne upisati 600 kg; tip `SORTIRANJE`, oprema linija; „Otvori šaržu" →
   toast sa brojem; lista SARZE pokazuje OTVORENA; Palete: paleta sa 600 nema
   `Preradjeno`, cela ima.
5. IZLAZ: dodati 3 reda (Rolend I 520 kg PALETA sa pakovanjem, Bruh 300, Griz 150),
   OTPAD 20; KPI pokazuje NERASPOREDJENO 10; „Završi" ugašeno; „Knjiži gubitak" → red
   GUBITAK 10 → „Završi" upaljeno → potvrda → PDF procesnog lista + 3 paletna lista LJ.
6. LAGER: 3 nove jedinice, raspoloživo = kg; ulazne palete: 400 raspoloživo / 0.
7. Fakturisanje / Gotova roba: 3 nove jedinice vidljive (oznaka `LJ n/2026`), legacy
   lotovi i dalje kao `PRE n/2026`; „Napravi utovar" 200 kg sa `LJ`; utovarna lista
   pokazuje lot, proizvod, rok; Fakturiši; SEF pregled: identitet stavke `LJ-…`.
8. Storno šarže sa utovarenim izlazom → odbijeno sa imenom jedinice; storno fakture →
   storno utovara → storno šarže prolazi; palete vraćene (raspoloživo puno, `Preradjeno`
   prazno).
9. Palete / Nova prerada: stari tok (cele palete, kutije/kese, tip GP) → nastaje šarža
   `PRERADA_LEGACY` sa GUBITAK = razlika; preradni list i paletni list LJ.
10. Sledljivost / LANAC: red otkupa čija je paleta prošla sortiranje pokazuje
    `SORT n/2026 -> LJ …`, stanje `preradjeno`; posle utovara `delimicno prodato`.
11. `PROIZVODNJA_AKTIVNA=NE` → ekran nestaje; Palete/Nova prerada radi legacy put.
12. `RunProductionHealthCheck`: blok P bez nalaza.

---

## 16) Rizici i mitigacije

| Rizik | Verovatnoća | Uticaj | Mitigacija |
|---|---|---|---|
| B1 refaktor #248 koda unese regresiju u prodaju/SEF | srednja | visok | B1 je **samo** promena ključa, bez feature-a; 163 testa + 20 sabotaža ostaju; SEF snapshot test poredi JSON pre/posle za legacy lot (isto osim novog polja) |
| Materijalizacija na klijentu sa prljavim `tblPrerada` (dupli `PreradaID`, prazan tip) | srednja | srednji | fail-closed: LJ bez proizvoda nije prodajna; P4 prijavljuje; `BrojacIdova` guard već postoji |
| Legacy paleta menjana (Adjust/Reassign) posle ulaska u proces | niska | visok | I5 kapija u sva tri writer-a + P3; `Istorija` kolona palete beleži pokušaj |
| Excel performanse na ekranu (8 tabela više) | srednja | srednji | `BeginTableCache`/`EndTableCache` u `Scr_Rows`; mape jednim prolazom; nema `LookupValue` po redu (S5) |
| Konflikti pri rebase-u na #248 module | visoka | nizak | strog redosled PR-ova; novi kod u novim modulima gde god je moguće; konstante na kraj blokova |
| Operater unosi izlaz bez pakovanja pa utovarna lista nema broj kutija | srednja | nizak | pravilo #248 ostaje: obrazac radije prazan nego pogrešan |
| Tolerancija 0,5 kg premala za višnju/pire (sok, para) | srednja | nizak | config po instalaciji; `GUBITAK` red je legitiman izlaz, ne greška |
| Compile greške tipa „Duplicate declaration" / zaklonjeno ime (mine iz #248) | srednja | srednji | `vba_check ZAKLONJENO`, jedinstvena imena promenljivih po proceduri, compile pre svakog push-a |

---

## 17) Šta ovaj plan namerno ne pokriva

Rezervacije prodaje (`AvailableKg − ReservedKg`), specifikacije kupaca, kalibracije kao
entitet, lotovi kao entitet, senzori i automatsko očitavanje temperature, smene i
radnici, PWA/GAS sync proizvodnih tabela, cena GP po jedinici (ostaje unos operatera
pri utovaru, #248). Sve navedeno ima mesto u modelu (LJ, parametri, kontrole) i ne
zahteva promenu ključeva kad dođe na red.

---

## Dodatak A — mapa 59 tačaka predloga na plan

| Tačke | Gde u planu |
|---|---|
| 1–3, 8 | §4.5–4.7, §6 |
| 4, 34, 35 | §4.4 (bez `KgTrenutno`, bez `Status`) |
| 5–7 | §6.3, T05–T07 |
| 9–11 | §5 I2, §6.4, `BalansSarze`, P1 |
| 12–13 | §4.1, I10 |
| 14–15 | §4.2 (seed), LJ `Klasa/Kalibracija` |
| 16–17 | §17 (F) |
| 18–22 | §4.8 parametri, §4.3 oprema; senzori F |
| 23, 41–42 | §10 |
| 24–28 | PR-E |
| 29 | §4.3 |
| 30–31 | F; `OdgovorniRadnik` tekst, audit kolone |
| 32–33 | §4.4 pakovanje, tip `PAKOVANJE` |
| 36–40 | PR-D |
| 43–47 | §8; `LotBroj` na LJ |
| 48–49 | #248 (utovar), §7 |
| 50 | §13, PR-A/B1/B2/D |
| 51–54 | §6.1–6.5, ADR-0001 |
| 55–58 | §9 |
| 59 | `BalansSarze`, `GetSarzeForGrid`, `LAGER` lista; izveštaji po tipu/opremi/stanici u PR-C1 kao čitanja |
