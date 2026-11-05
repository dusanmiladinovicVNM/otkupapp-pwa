# Sledljivost kao graf — identitet, ivice, invarijante i odluke S9

> **Izvor:** pre-flight S9 (plan `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` §14.64,
> `docs/STANJE_REFAKTORA.md` stavka 324). Mereno nad `main` = `57ab47ed`, posle zatvaranja
> S8 (stavka 322). Linije koda u ovom fajlu važe za taj commit.
>
> Prvi pre-flight (PR `#461`, grana `claude/sharp-cray-iyhy1m`, mereno nad `d19c2dae` dok je S8
> trajao) je ovde ponovo izmeren i zamenjen. Odluke iz njega važe.
>
> **Odluke (operater, 04.11.2026): sve `SLED-ODL-01..08` su DONETE** (§5), uz pravilo klase
> `SLED-KLASA-01` (varijanta b, §5a). Invarijante `SLED-INV-xx` imenuju šta kod danas drži, šta
> mu fali i ko je vlasnik posle S9.

## 1) Šta je već odlučeno (ne otvara se ponovo)

| Pravilo | Gde |
|---|---|
| Poslovni broj je **labela**: sme za pretragu, prikaz, generator i proveru duplikata pri unosu; ne sme za FK, izbor dokumenta, storno ni spoj | `ARCHITECTURE_CONTRACT.md` A2 |
| Pripadnost drže **isključivo** tabele `*Izvori`; kolona `Otkup.OtpremnicaID` i `Otpremnica.ZbirnaID` u ciljnom modelu nema, a „gde je sada" se računa | `DOCUMENT_HEADER_LINES.md` §5 |
| Članstvo je zapisano **po verziji** (A13) | `ARCHITECTURE_CONTRACT.md` |
| G3 — sledljivost bez pogađanja: lanac unazad ide samo kroz ID/FK graf | `ARCHITECTURE_CONTRACT.md`, „Četiri tvrde kapije" |
| Identitet zbirne je `ZbirnaID` (S4-2a), prijemnice `PrijemnicaID` (S6 korak 4), stavke prijemnice `PrijemnicaStavkaID` (S6 korak 2); generacija nije identitet | plan §14.24, §14.55 |
| Faktura zna tačnu `PrijemnicaStavkaID` (S7) | `docs/DOMEN/FAKTURA.md` |
| Paletna stavka zna samo tačnu `PrijemnicaStavkaID` (S8 rez 2, `PAL-INV-01`); broj prijemnice, broj zbirne i trag zbirne čitaju se iz zaglavlja prijemnice | `docs/DOMEN/PALETE.md`, `PAL-ODL-03`, `PAL-ODL-13` |
| Sledljivost na paleti (C-040 i spisak kooperanata na paletnom i preradnom listu) pripada S9 | `PAL-ODL-07` |
| C002 ostaje na izdanju pre S8 reza 2b; posledica: ni izmene S9 ne idu u C002 | `PAL-ODL-12` (stavka 311) |
| Stare provere integriteta se **ne prevode**: brišu se sa starim modelom, a slajs dodaje proveru pravila novog modela | plan §14.8, odluka 11 |
| Prerada 2.0 (lager jedinica, šarža, genealogija) je **zaseban plan, nije implementiran**; njena faza C2 je pisana nad obrisanim modulima | `docs/PRERADA_2_MODEL_I_PLAN.md` §8 |

## 2) Graf — ivice, kardinalnost i čitač

Čvor grafa je **dokument ili stavka sa svojim ID-om**. Ivica postoji samo tamo gde je u kanonu
zapisan ID drugog čvora. Broj, datum i naziv su osobine čvora, nikad ivica.

| # | Ivica | Nosilac u kanonu | Kardinalnost | Kanonski čitač (`57ab47ed`) | Strog javni? |
|---|---|---|---|---|---|
| E1 | otkup → stavke otkupa | `tblOtkupStavke.OtkupID` | 1:N | `modOtkup.StavkeOtkupaRedovi:1132` | da |
| E2 | otkup ↔ otpremnica | `tblOtpremnicaIzvori` | N:1; otkup je u najviše jednoj **aktivnoj** otpremnici, istorijski u više | `modDokumenta.IzvoriOtpremnice:5523`, `OtpremnicaZaOtkup:5505`, mapa `AktivnoClanstvoOtpremnica:5533` | da |
| E3 | otpremnica ↔ zbirna | `tblZbirnaIzvori` | N:1, isto pravilo aktivne verzije | `IzvoriZbirne:2883`, `AktivnaZbirnaZaOtpremnicu:1527`, mapa `AktivnoClanstvoZbirni:1544` | da |
| E4 | zbirna ↔ prijemnica | `tblPrijemnica.ZbirnaRoditeljID` | **1:N meko** (namera 1:1; F4 pita „dupla prijemnica?" i pušta). C002: 5 od 894 zbirnih ima više od jedne aktivne prijemnice | samo **privatno** `modAutoHladnjaca.AktivnePrijemnicePoZbirni:539` (po ID-u, strogo) | ne |
| E5 | prijemnica → stavke | `tblPrijemnicaStavke.PrijemnicaID` | 1:N, najviše jedna po klasi | `modDokumenta.StavkePrijemniceRedovi:7340`, `StavkePrijemnicePoDokumentu:7631` | da |
| E6 | stavka prijemnice → linija fakture | `tblFakturaStavke.PrijemnicaStavkaID` | najviše jedna aktivna linija po stavci | `modFaktura.FakturaLinijePoStavci:984` | da |
| E7 | stavka prijemnice → paletne stavke | `tblPaletaStavka.PrijemnicaStavkaID` | 1:N; paleta nosi stavke više prijemnica (N:M na nivou palete) | **čitača nema**. Pravilo postoji kao javni predikat `modPaletniList.PalInv01Razlog:2726` (pisac ga drži tvrdo kroz `RequireLinijaPalete:2752`, integritet meko u C1, `modIntegritet:626`); svi čitači po ključu su privatni ili u transakciji (`:196`, `:2461`, `:2861`) | ne |
| E8 | paleta → prerada | `tblPreradaStavka.PaletaID` | paleta u najviše jednoj aktivnoj preradi; prerada N paleta | nema; inline u `modStorno.StornoPrerada:1525` i `modPaletniList.FillPreradaSablon:3266` | ne |
| E9 | prerada → utovar | `tblUtovarStavke.PreradaID` | N:M, **delimično po kg** (zbir ≤ `NetoIzlazKg`) | `modUtovar.UtovarenoPoPreradi:53` vraća samo kg po preradi, bez ID-a utovara (meko) | ne |
| E10 | utovar → GP faktura | `tblUtovar.FakturaID`; `tblFakturaStavke.UtovarID`/`PreradaID` | utovar → najviše jedna aktivna GP faktura (pisac `modUtovar:1633`, kapija `AktivnihFstZaUtovar:373` na `:1497` i `modStorno:1461`). Faktura → N utovara: **NIJE izmereno** | brojač `AktivnihFstZaUtovar`; pravilo linije `modSEFMapper.GpFakturaRazlog:354` | brojač da |
| — | otkup → kooperant, parcela; prijemnica → kupac | kolone zaglavlja | N:1 | kolone | — |

**Nisu ivice** (labele ili kolone bez pisca, §4 `SLED-INV-03`): `Otkup.OtpremnicaID`,
`Otkup.BrojOtpremnice`, `Otkup.BrojZbirne`, `Otkup.VozacID`, `Otkup.ZbirnaRoditeljID`,
`Otpremnica.BrojZbirne`, `Otpremnica.ZbirnaRoditeljID`, `Prijemnica.BrojZbirne`, svaki
`GeneracijaID` i broj-oblik traga ispravke (`IspravkaOd`, `ZamenjenSa`).

**Postojeći hodovi po ID-u kroz više karika:**
- `modStammdatenSync.TekucaPredajaOtkupa:1006`: otkup → otpremnica → zbirna;
- `modPrint.SpecifikacijaBlokovaRedovi:3250`: otpremnica → otkupi → kooperant.

Sve ostalo što danas hoda lancem ide bar jednu kariku po broju ili po koloni bez pisca (plan
§14.64, N2 i N4).

## 3) Identitet

**`SLED-IDENT-01` — graf hoda samo ivicama E1–E10.** Broj koji operater ukuca u pretragu
razrešava se **jednom** u skup ID-eva. Kad broj nose dva dokumenta, prikazuju se oba, sa
vlasnikom niza (`ARCHITECTURE_CONTRACT.md` A2). Ne bira se prvi pogodak. Od razrešenja dalje
putuje samo ID: do čvora, do štampe dokumenta (C-043) i do papira.
- **Danas:** ekran sledljivosti ne postoji (obrisan u S1b-3b, `b4d74093`).
- **Hodovi lancem po broju zbirne** (21 živo mesto, plan §14.64 merenje 1): izveštaji, manjak,
  detalj, izvoz za PWA (`BuildPrijemnicaIndexByBrojZbirne:1193`, „prva je dovoljna"), storno
  zbirne i prijemnice, poruka posle storna, F4 i Oporavak.
- **Vlasnik posle S9:** jezgro grafa (rez S9-1) i kapija `NEMA_BROJA_KAO_FK` kao grupa popisa sa
  pragom 0 (rez S9-3).

**`SLED-IDENT-02` — identitet ne zavisi od generacije.** `GeneracijaID` nema produkcionog
pisca: `ApplyGeneracijaID` (`modDokumenta:6315`) zove samo `modGoldenTests:1313`, a
`NewGeneracijaID` samo `GeneracijaIDZaBrojArr`, kome niko ne pristupa. U C002 je kolona prazna
u svakom aktivnom redu otkupa, otpremnice, zbirne i novca (izmereno, plan §14.64 merenje 5).
Kolonu ipak čitaju četiri živa produkciona mesta (plan §14.64, N6), a jedno od njih bira
**cilj mutacije**: Oporavak
„Prevezi prijemnicu" prenosi generaciju umesto `ZbirnaID`, pa pisac razrešava ciljnu zbirnu po
broju (N7). **Vlasnik posle S9:** rez S9-4 i kapija `NEMA_GENERACIJE`.

## 4) Invarijante

**`SLED-INV-01` — svaka ivica razrešava tačno jednom.** Za svaki red nosioca ivice (E2–E10)
moraju da važe dve stvari:
- roditelj i dete postoje **tačno jednom**;
- aktivno dete ima aktivnog roditelja (zbirna, prijemnica, linija fakture, paletna stavka).

Nula ili dva su kvar, isto kao dva aktivna članstva.
- **Danas, pri upisu:** drže je pisci — E2 `OtpRequireIzvorValjan:4673`, E3
  `ZbrRequireIzvorValjan:2213`, E4 `ZbirnaRoditeljPrijemnice:6922` (`ZBR-PARENT-01`), E6
  `UgovorLinijaFaktureNalazi:1106`, E7 `RequireLinijaPalete:2752` (`PAL-INV-01`).
- **Danas, pri čitanju:** strogi čitači E1–E3, E5 i E6.
- **Revizija je nema:** `modIntegritet` ne čita nijednu tabelu `*Izvori`, a E4 ne proverava
  nijedna provera. U C002 je 6 od 905 aktivnih prijemnica bez `ZbirnaRoditeljID`, a lista
  INTEGRITET to ne prijavljuje (B5 meri broj, ne ID; plan §14.64, N14).
- **Vlasnik posle S9:** jezgro grafa (čvor koji ne razrešava se ne vezuje i dobija oznaku, ne
  pogađa se) i integritet (rez S9-5).

**`SLED-INV-02` — revizija meri model, ne kolone bez pisca.** Provera koja čita kolonu bez
pisca je ili uvek tiha ili uvek glasna, i obe su laž. Izmereno nad kopijom C002 sa kodom
`57ab47ed` (plan §14.64, merenje 5):
- **A1 je uvek tih:** 0 nalaza;
- **B5b je uvek glasan:** 921 nalaz, tačno koliko ima aktivnih otpremnica;
- health 7 je uvek OK (izvedeno iz koda, nije pokretan).

**Vlasnik posle S9:** rez S9-5, po §14.8 t. 11. Briše se, ne prevodi se.

**`SLED-INV-03` — kolona bez živog pisca nema živog čitaoca.** Kad poslednji pisac kolone
nestane, u istom slajsu nestaje i čitalac. Ako to ne može u istom slajsu, čitalac se gasi pre
nego što kolona izađe iz kanona. S3–S6 su ovo ostavili otvorenim: brisanje kolona otkupa,
otpremnice i zbirne je odloženo „u svoj rez" i nikad nije došlo. **Vlasnik posle S9:** rez
S9-6 (kanon i self-heal), kapija `popis_citalaca.py --check` (prag 0 po grupi).

## 5) Odluke — DONETE (operater, 04.11.2026)

| ID | Pitanje | Odluka | Zašto | Rez |
|---|---|---|---|---|
| **`SLED-ODL-01`** | Šta znači „poreklo" kad zbirna ide na dve prijemnice (E4 1:N), a paleta nosi stavke više prijemnica (E7 N:M)? | **Nadskup bez alokacije**, sa natpisom „moguće poreklo". Unazad: prijemnica → svi otkupi aktivnih otpremnica njene zbirne. Unapred: otkup → sve prijemnice, fakture i palete te zbirne. Tačno je samo gde je ivica 1:1 ili po kg (utovar). Klasa po `SLED-KLASA-01` | Alokacije nema i ne gradi se (§15 plana). Za opoziv robe je nadskup bezbedan smer, a za sertifikaciju je pošten samo ako je imenovan | S9-1 |
| **`SLED-ODL-02`** | Da li unazad suziti po klasi? | **Varijanta (b), `SLED-KLASA-01`:** hladnjački lanac se filtrira po klasi; kod ostalih prijemnica se ništa ne odseca, a otkupi iste klase idu prvi i označeni su | §5a | S9-1 |
| **`SLED-ODL-03`** | Verzije: šta graf crta posle storna i ispravke? | **Samo aktivne veze.** Istorija storna i ispravki ide iz `tblStornoVeze` (`OldDocID → NewDocID`), kao poseban prikaz, ne u LANAC | Mešanje verzija u jedan lanac bi sabralo robu dvaput | S9-1 |
| **`SLED-ODL-04`** | Koje sposobnosti ekrana SLEDLJIVOST se vraćaju? | **Vraćaju se:** LANAC (C-031, C-035, C-036, C-044, C-045), PARCELE (C-032), A4 lanca (C-041), dokument iza reda (C-043), GlobalGAP dokument zbirne (C-039), izbor dokumenta perioda (C-040). **NEPOTPUNI (C-033/C-034)** postaje lista INTEGRITET. **C-037/C-038** su REPLACED radnim stolom F2 (radnja „vezi" → `DodajOtpremnicaIzvor_TX`). **C-042** ide zajedničkim putem štampe liste | Minimum koji odgovara na „sa kojih parcela je roba sa ove fakture" | S9-7 |
| **`SLED-ODL-05`** | Spisak kooperanata na paletnom i preradnom listu (`PAL-ODL-07`) | Kroz **isti** graf, kao nadskup po `SLED-ODL-01`, sa natpisom „moguće poreklo" | Jedno mesto koje zna poreklo; papir i ekran ne smeju dati dva odgovora | S9-7 |
| **`SLED-ODL-06`** | Granica sa Preradom 2.0 | **Jedan graf, modul `modSledljivost`** (ime je slobodno: modul iz plana Prerade 2.0 je obrisan u `c79935be`). Prerada 2.0 C2 ga proširuje ivicama LJ i šarže, a ne pravi svoj. `PRERADA_2_MODEL_I_PLAN.md` §8 se ispravlja u S9-1 | reuse > new; dva grafa bi dala dva porekla | S9-1 |
| **`SLED-ODL-07`** | Izvoz za PWA (`ExportOtkupiAll`): koju prijemnicu otkup pokazuje kad zbirna ima dve? | Prijemnice po `ZbirnaRoditeljID`. Kad ih je više: **najranija po datumu**, uz broj prijemnica. Ugovor kolona GAS/PWA se ne menja. Operater: **PWA se verovatno neće koristiti, AgriX VBA će najverovatnije biti samo desktop**, pa je ovo minimalna popravka niskog prioriteta | Danas „prva je dovoljna" po **broju** (`modStammdatenSync:1245`) | S9-3 |
| **`SLED-ODL-08`** | Labele na detetu: `Prijemnica.BrojZbirne` i broj-oblik traga ispravke | **`Prijemnica.BrojZbirne` ostaje** kao snapshot labele za papir i pretragu, ali **nijedan spoj** po njemu (kapija `NEMA_BROJA_KAO_FK`). Broj-oblik traga (`IspravkaOd`/`ZamenjenSa`) se briše gde nema produkcionog pisca (otpremnica, zbirna, novac); na prijemnici ostaje dok papir ne pređe na `tblStornoVeze` | Jedini pisac traga (`StampIspravkaTrace:1453`) u produkciji zove samo prijemnica (`modDokUnos:1339`); za identitet od broj-oblika ne zavisi ništa | S9-3, S9-6 |

## 5a) Klasa — `SLED-KLASA-01`

**Izmereno nad `57ab47ed`** (isto kao nad `d19c2dae`):

| Karika | Da li se klasa čuva | Dokaz |
|---|---|---|
| otkup → otpremnica | **da, po konstrukciji**: izdavanje traži jednakost kg i gajbi po klasi između stavki otpremnice i zbira stavki izvora | `modDokumenta.OtpIzdaj:4888` → `OtpRequireJednakost:5184` |
| otpremnica → zbirna | **da, po konstrukciji** | `ZbrIzdaj:2427` → `ZbrRequireJednakost:2513` |
| zbirna → prijemnica (hladnjački lanac) | **da**: lanac je 1:1 po klasi na sva četiri nivoa | `HLD-LANAC-01` |
| zbirna → prijemnica (ostalo) | **ne**: „klasifikacija je procena pri prijemu"; pisac poredi broj, kupca i vozača, a klasu ne | `PRJ-ZBR-02` (odluka 27.10.2026); `ZbirnaRoditeljPrijemnice:6922` |
| prijemnica → paleta | **da**: paletna stavka je stavka prijemnice | `PAL-INV-01` |

Zašto se ne filtrira ni „kad se klase slažu": zbirna I 500 / II 300 i prijemnica I 450 / II 300
mogu biti 50 kg kala na I, ali i 50 kg premeštenih iz I u II uz 50 kg kala na II. Podaci ta dva
slučaja ne razlikuju, a filter bi u drugom izostavio stvarne izvore klase I. Za opoziv robe i
sertifikaciju je izostavljanje opasnija greška od viška.

**`SLED-KLASA-01`** (pravilo za jezgro grafa i sve njegove čitaoce: ekran, papir, spisak
kooperanata na paletnom listu):
- **hladnjački lanac:** unazad samo otkupi iste klase, jer je 1:1 po klasi dokazano piscem;
- **ostale prijemnice:** unazad svi otkupi aktivnih otpremnica zbirne. Otkupi iste klase kao
  stavka prijemnice idu prvi, sa oznakom „ista klasa". Ostali dolaze posle njih, sa oznakom
  „druga klasa — moguće preklasiranje pri prijemu";
- **unapred, od otkupa:** sve prijemnice zbirne, sa istim oznakama po klasi stavke prijemnice.

Nijedan izvor se ne odseca po klasi van hladnjačkog lanca.

## 6) Šta S9 namerno NE radi

- Ne gradi alokaciju (koja kila otkupa je u kojoj kili prijema). Ako zahtev dođe, alokacija se
  gradi kao eksplicitna tabela (§15 plana).
- Ne menja ugovor kolona GAS/PWA. VBA izvoz menja samo izvor vrednosti.
- Ne menja pisce paleta i prerade (S8 je zatvoren; Prerada 2.0 je zaseban plan).
- Ne pravi drugi graf za Preradu 2.0 (`SLED-ODL-06`).
- Ne ide u C002 (`PAL-ODL-12`: C002 ostaje na izdanju pre S8 reza 2b).
