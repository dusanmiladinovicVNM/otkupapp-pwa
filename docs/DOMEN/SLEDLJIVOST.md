# Sledljivost kao graf — identitet, ivice, invarijante i odluke S9

> **Izvor:** pre-flight S9 (plan `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.59,
> `docs/STANJE_REFAKTORA.md` stavka 288). Mereno nad `main` = `d19c2dae`, dok S8 još traje
> (rez 1b je PR `#459`, otvoren). Linije koda u ovom fajlu važe za taj commit.
>
> **Odluke `SLED-ODL-01..08` su PREDLOZI.** Operater ih nije doneo. Dok ne budu donete,
> važi verdikt i GO / HOLD iz plana §14.59. Invarijante `SLED-INV-xx` imenuju šta kod danas
> drži, šta mu fali i ko je vlasnik posle S9.

## 1) Šta je već odlučeno (ne otvara se ponovo)

| Pravilo | Gde |
|---|---|
| Poslovni broj je **labela**: sme za pretragu, prikaz, generator i proveru duplikata pri unosu; ne sme za FK, izbor dokumenta, storno ni spoj | `ARCHITECTURE_CONTRACT.md` A2 |
| Pripadnost drže **isključivo** tabele `*Izvori`; kolona `Otkup.OtpremnicaID` i `Otpremnica.ZbirnaID` u ciljnom modelu nema, a „gde je sada" se računa | `DOCUMENT_HEADER_LINES.md` §5; `ARCHITECTURE_CONTRACT.md` (referencijalni integritet članstva) |
| Članstvo je zapisano **po verziji**: na pitanje „od kojih je dokumenata ova verzija bila sastavljena" se odgovara bez gledanja trenutnog stanja | `ARCHITECTURE_CONTRACT.md` (A13, članstvo po verziji) |
| G3 — sledljivost bez pogađanja: lanac unazad ide samo kroz ID/FK graf | `ARCHITECTURE_CONTRACT.md`, „Četiri tvrde kapije" |
| Identitet zbirne je `ZbirnaID` (S4-2a), prijemnice `PrijemnicaID` (S6 korak 4), stavke prijemnice `PrijemnicaStavkaID` (S6 korak 2); generacija nije identitet | plan §14.24, §14.55 |
| Faktura zna tačnu `PrijemnicaStavkaID` (S7) | `docs/DOMEN/FAKTURA.md`; plan §14.57 |
| Paletna stavka zna tačnu `PrijemnicaStavkaID`; sledljivost na paleti (C-040 i spisak kooperanata na paletnom i preradnom listu) **prelazi u S9** | `docs/DOMEN/PALETE.md`, `PAL-INV-01`, `PAL-ODL-07` |
| Stare provere integriteta se **ne prevode**: brišu se sa starim modelom, a slajs dodaje proveru pravila novog modela | plan §14.8, odluka 11 |
| Prerada 2.0 (lager jedinica, šarža, genealogija) je **zaseban plan, nije implementiran**; njena faza C2 je pisana nad modulima koji su u međuvremenu obrisani | `docs/PRERADA_2_MODEL_I_PLAN.md` §8 |

## 2) Graf — ivice, kardinalnost i čitač

Čvor grafa je **dokument ili stavka sa svojim ID-om**. Ivica postoji samo tamo gde je u kanonu
zapisan ID drugog čvora. Broj, datum i naziv su osobine čvora, nikad ivica.

| # | Ivica | Nosilac u kanonu | Kardinalnost | Kanonski čitač danas (`d19c2dae`) | Strog? |
|---|---|---|---|---|---|
| E1 | otkup → stavke otkupa | `tblOtkupStavke.OtkupID` | 1:N | `modOtkup.StavkeOtkupaRedovi:1132` | da |
| E2 | otkup ↔ otpremnica | `tblOtpremnicaIzvori` | N:1; otkup je u najviše jednoj **aktivnoj** otpremnici, istorijski u više | `modDokumenta.IzvoriOtpremnice:5523` (→ `OtpClanovi:5358`), `OtpremnicaZaOtkup:5505`, mapa `AktivnoClanstvoOtpremnica:5533` | da |
| E3 | otpremnica ↔ zbirna | `tblZbirnaIzvori` | N:1; isto pravilo aktivne verzije | `IzvoriZbirne:2883`, `AktivnaZbirnaZaOtpremnicu:1527`, `AktivnoClanstvoZbirni:1544` | da |
| E4 | zbirna ↔ prijemnica | `tblPrijemnica.ZbirnaRoditeljID` | **1:N meko** (namera 1:1, F4 pita „dupla prijemnica?" i pušta) | **javnog čitača nema**; privatno `modAutoHladnjaca.AktivnePrijemnicePoZbirni:539` (po ID-u, strogo) | samo privatno |
| E5 | prijemnica → stavke | `tblPrijemnicaStavke.PrijemnicaID` | 1:N, najviše jedna po klasi | `modDokumenta.StavkePrijemniceRedovi:7340`, `StavkePrijemnicePoDokumentu:7631` | da |
| E6 | stavka prijemnice → linija fakture | `tblFakturaStavke.PrijemnicaStavkaID` | najviše jedna aktivna linija po stavci (puna količina) | `modFaktura.FakturaLinijePoStavci:984` | da |
| E7 | stavka prijemnice → paletne stavke | `tblPaletaStavka.PrijemnicaStavkaID` | 1:N; paleta nosi stavke više prijemnica (N:M na nivou palete) | **javnog čitača po `PrijemnicaStavkaID` nema** (`modPaletniList:193`, `:2263`, `:2659` su privatni ili u transakciji) | ne |
| E8 | paleta → prerada | `tblPreradaStavka.PaletaID` | paleta u najviše jednoj aktivnoj preradi; prerada N paleta | nema; inline u `modStorno.StornoPrerada:1517`, `modPaletniList.FillPreradaSablon:3132` | ne |
| E9 | prerada → utovar | `tblUtovarStavke.PreradaID` | N:M, **delimično po kg** (zbir ≤ `NetoIzlazKg`) | `modUtovar.UtovarenoPoPreradi:53` (meko, bez ID-a utovara) | ne |
| E10 | utovar → GP faktura | `tblFakturaStavke.UtovarID` / `PreradaID`, `tblUtovar.FakturaID` | NEPROVERENO | `modFaktura.AktivnihFstZaUtovar:373`; `modSEFMapper.GpFakturaRazlog:354` | brojač da |
| — | otkup → kooperant, parcela; prijemnica → kupac | kolone zaglavlja | N:1 | kolone | — |

**Nisu ivice** (labele ili kolone bez pisca, §4 `SLED-INV-03`): `Otkup.OtpremnicaID`,
`Otkup.BrojOtpremnice`, `Otkup.BrojZbirne`, `Otkup.VozacID`, `Otkup.ZbirnaRoditeljID`,
`Otpremnica.BrojZbirne`, `Otpremnica.ZbirnaRoditeljID`, `Prijemnica.BrojZbirne`,
`PaletaStavka.BrojZbirne`, `PaletaStavka.BrojPrijemnice`, svaki `GeneracijaID` i broj-oblik
traga ispravke (`IspravkaOd`, `ZamenjenSa`).

**Jedini postojeći hod po ID-u kroz više karika:**
- `modStammdatenSync.TekucaPredajaOtkupa:1006`: otkup → otpremnica → zbirna;
- `modPrint.SpecifikacijaBlokovaRedovi:3250`: otpremnica → otkupi → kooperant.

Sve ostalo što danas hoda lancem ide bar jednu kariku po broju (plan §14.59, N4 i N5).

## 3) Identitet

**`SLED-IDENT-01` — graf hoda samo ivicama E1–E10.** Broj koji operater ukuca u pretragu
razrešava se **jednom** u skup ID-eva. Kad broj nose dva dokumenta, prikazuju se oba, sa
vlasnikom niza (`ARCHITECTURE_CONTRACT.md` A2). Ne bira se prvi pogodak. Od razrešenja dalje
putuje samo ID: do čvora, do štampe dokumenta (C-043) i do papira.
- **Danas:** ekran sledljivosti ne postoji (obrisan u S1b-3b, `b4d74093`).
- **Preostali hodovi lancem idu po broju:**
  - prijem iz izvoza za PWA, „prva je dovoljna": `modStammdatenSync.BuildPrijemnicaIndexByBrojZbirne:1209`;
  - detalj izveštaja: `modScrIzvestaji.DodajPrijemniceZbirne:2347`;
  - storno zbirne: `ActivePrijIDsByZbirna:1880`, `ScanZbirna:2367`;
  - integritet: `PrijemnicaZbirnaMap:1382`.
- **Vlasnik posle S9:** jezgro grafa (rez S9-1) i kapija `NEMA_BROJA_KAO_FK` (rez S9-3).

**`SLED-IDENT-02` — identitet ne zavisi od generacije.** `GeneracijaID` nema produkcionog
pisca:
- `ApplyGeneracijaID` (`modDokumenta:6315`) zove samo `modGoldenTests:1313`;
- `NewGeneracijaID` zove samo `GeneracijaIDZaBrojArr`, kome spolja niko ne pristupa.

Kolonu ipak čita šest produkcionih mesta (plan §14.59, N6). Jedno od njih bira **cilj
mutacije**: Oporavak „Prevezi prijemnicu" prenosi generaciju umesto `ZbirnaID`, pa pisac
razrešava ciljnu zbirnu po broju (N7). **Vlasnik posle S9:** rez S9-4 i kapija
`NEMA_GENERACIJE`.

## 4) Invarijante

**`SLED-INV-01` — svaka ivica razrešava tačno jednom.** Za svaki red nosioca ivice (E2–E10)
moraju da važe dve stvari:
- roditelj i dete postoje **tačno jednom**;
- aktivno dete ima aktivnog roditelja (zbirna, prijemnica, linija fakture, paletna stavka).

Nula ili dva su kvar, isto kao dva aktivna članstva.
- **Danas, pri upisu:** drže je pisci.
  - E2: `OtpRequireIzvorValjan`;
  - E3: `ZbrRequireIzvorValjan`;
  - E4: `ZbirnaRoditeljPrijemnice:6922`, `ZBR-PARENT-01`;
  - E6: `UgovorLinijaFaktureNalazi`;
  - E7: posle S8 reza 2b, `PAL-INV-01`.
- **Danas, pri čitanju:** strogi čitači E1–E3, E5 i E6.
- **Revizija je nema:** `modIntegritet` ne čita nijednu tabelu `*Izvori`, a E4 ne proverava
  ni jedna provera.
- **Vlasnik posle S9:** jezgro grafa (čvor koji ne razrešava se ne vezuje i dobija oznaku, ne
  pogađa se) i integritet (rez S9-6).

**`SLED-INV-02` — revizija meri model, ne kolone bez pisca.** Provera koja čita kolonu bez
pisca je ili uvek tiha ili uvek glasna, i obe su laž. Danas su takve tri:
- A1 je uvek tih: poredi 0 sa 0;
- B5b prijavljuje **svaku** aktivnu otpremnicu;
- health 7 je uvek OK.

Plan §14.59, N2. **Vlasnik posle S9:** rez S9-6, po §14.8 t. 11. Briše se, ne prevodi se.

**`SLED-INV-03` — kolona bez živog pisca nema živog čitaoca.** Kad poslednji pisac kolone
nestane, u istom slajsu nestaje i čitalac. Ako to ne može u istom slajsu, čitalac se gasi pre
nego što kolona izađe iz kanona. Ovo je pravilo koje su S3–S6 ostavili otvorenim:
- plan :5995 (S5-3b): „brisanje kolona ide u svoj rez, zajedno sa prijemnicom i paletom u S6";
- S6 je izbrisao kolone prijemnice, ali ne i otkupa, otpremnice i zbirne.

**Vlasnik posle S9:** rez S9-5 (kanon + self-heal), kapija `popis_citalaca.py --check` (prag 0
po grupi) i statička pravila iz §13 plana.

## 5) Odluke — PREDLOZI (operater ih nije doneo)

| ID | Pitanje | Predlog | Zašto | Blokira |
|---|---|---|---|---|
| **`SLED-ODL-01`** | Šta znači „poreklo" kad zbirna ide na dve prijemnice (E4 1:N), a paleta nosi stavke više prijemnica (E7 N:M)? | **Nadskup bez alokacije**, sa natpisom „moguće poreklo". Unazad: prijemnica → svi otkupi aktivnih otpremnica njene zbirne. Unapred: otkup → sve prijemnice, fakture i palete te zbirne. Tačno samo gde je ivica 1:1 ili po kg (E9) | Alokacije nema i ne gradi se (§15 plana, „delimična alokacija"). Za opoziv (recall) je nadskup bezbedan smer, a za sertifikaciju je pošten samo ako je imenovan | S9-1 |
| **`SLED-ODL-02`** | Da li unazad suziti po klasi (stavka prijemnice klase II → samo stavke otkupa klase II)? | **Ne sužavati.** Klasa se prikazuje uz svaku kariku, a izvori se ne odsecaju | Kupac sme da preklasira robu pri prijemu. Suženje bi tada dalo prazan ili pogrešan skup, a nijedan pisac ne meri da je klasa očuvana kroz lanac | S9-1 |
| **`SLED-ODL-03`** | Verzije: šta graf crta posle storna i ispravke? | Podrazumevano **samo aktivne ivice**. Istorija ide kroz `tblStornoVeze` (`OldDocID → NewDocID`), kao poseban prikaz, ne u LANAC. Isto pravilo ima Prerada 2.0 (§8.1: „stornirane ivice se ne crtaju") | Članstvo po verziji već postoji. Mešanje verzija u jedan lanac bi sabralo robu dvaput | S9-1 |
| **`SLED-ODL-04`** | Koje sposobnosti ekrana SLEDLJIVOST se vraćaju, i u kom obliku? | **Vraćaju se:** LANAC (C-031, C-035, C-036, C-044, C-045), PARCELE (C-032), A4 lanca (C-041), štampa dokumenta reda (C-043), dokument sledljivosti zbirne za GlobalGAP (C-039) i izbor dokumenta perioda (C-040, `PAL-ODL-07`). **NEPOTPUNI (C-033/C-034)** se zamenjuje listom INTEGRITET iz reza S9-6, jer to su isti nalazi. **C-037/C-038** (ručno i auto vezivanje otkupa na otpremnicu) su **REPLACED** radnim stolom F2 (radnja „vezi", `modScrDokumenti:1284` → `VeziZaAktivnu:646` → `modDokumenta.DodajOtpremnicaIzvor_TX:3514`). **C-042** (štampa liste) ide kroz zajednički put liste | Palete će se slabo koristiti, a Prerada 2.0 dolazi u januaru 2027 (`PAL-ODL-01`). Minimum koji odgovara na „sa kojih parcela je roba sa ove fakture" | S9-7 |
| **`SLED-ODL-05`** | Spisak kooperanata na paletnom i preradnom listu (`PAL-ODL-07`) | Kroz **isto** jezgro grafa (E7 unazad), kao nadskup po `SLED-ODL-01`, sa natpisom „moguće poreklo" | Jedno mesto koje zna poreklo; papir i ekran ne smeju dati dva odgovora | S9-7 |
| **`SLED-ODL-06`** | Granica sa Preradom 2.0 | S9 gradi jezgro grafa sa ivicama E1–E10 nad današnjim tabelama: E1–E6, E9 i E10 u rezu S9-1, a palete (E7, E8) u S9-7, nad strogim čitačem iz S8 2b. Prerada 2.0 C2 **proširuje isto jezgro** ivicama LJ i šarže, a ne pravi svoje. Plan Prerade 2.0 §8 se pri S9-1 ispravlja: `modSledljivost` iz tog plana je obrisan (`c79935be`), a ime je slobodno | reuse > new; dva grafa bi dala dva porekla | S9-1 (ime i API) |
| **`SLED-ODL-07`** | Izvoz za PWA (`ExportOtkupiAll`): koju prijemnicu otkup pokazuje kad zbirna ima dve? | Prijemnice po `ZbirnaRoditeljID`. Kad ih je više: **najranija po datumu**, uz broj prijemnica u istom polju. Ugovor kolona GAS/PWA se ne menja u S9 | Danas „prva je dovoljna" po **broju** (`modStammdatenSync:1245`), a to je i 1:N i kolizija broja. Promena oblika za PWA je S5 teren | S9-3 |
| **`SLED-ODL-08`** | Labele na detetu: `Prijemnica.BrojZbirne` i broj-oblik traga ispravke (`IspravkaOd`/`ZamenjenSa`) | **`Prijemnica.BrojZbirne` ostaje kao snapshot labele za papir i pretragu**, kao `FakturaStavka.BrojPrijemnice` u S7, ali **nijedan spoj** (kapija `NEMA_BROJA_KAO_FK`). Broj-oblik traga se briše tamo gde nema produkcionog pisca (otpremnica, zbirna, novac). Na prijemnici ostaje dok papir ne pređe na `tblStornoVeze` | Papir prijemnice nosi broj zbirne (`modPrint.FillPrijemnicaSablon:1472`). Za identitet ništa ne zavisi od broj-oblika: jedini pisac bira red po ID-u (`StampIspravkaTrace:1431`) i u produkciji ga zove samo prijemnica (`modDokUnos:1339`), a čitaoci su samo naslov papira | S9-3, S9-5 |

## 6) Šta S9 namerno NE radi

- Ne gradi alokaciju (koja kila otkupa je u kojoj kili prijema). Ako zahtev dođe, alokacija
  se gradi kao eksplicitna tabela (§15 plana).
- Ne menja ugovor kolona GAS/PWA (S5 teren). VBA izvoz menja samo izvor vrednosti.
- Ne dira paletu i preradu van čitanja (S8; Prerada 2.0).
- Ne pravi drugi graf za Preradu 2.0 (`SLED-ODL-06`).
