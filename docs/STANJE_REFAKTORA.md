# Stanje refaktora „dokument = header + stavke“

> **Ulaz za svaku novu sesiju.** Kratko i ažurno — čita se umesto celog plana. Pun plan i istorijat odluka:
> `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` (odluke po datumu u §14.x; važeće: §14.7 „Odluke operatera 16.09“).
> Ažurira se na kraju svakog koraka, u istom commit-u.

**Ažurirano:** 16.09.2026, `main` `a826268` (posle #344, mapa F2).

## Pravila koja važe (16.09.2026)

- **Nema produkcije ni podataka koje treba štititi — radi se iznova.** Bez migracija, backfill-a, fallback-a.
- **Legacy kod ne mora da radi između faza.** Ne gradi se ništa što ga čuva živim: pauze, podela kapija, test-only pisci,
  kolone ostavljene za pauzirane čitaoce, mostovi.
- **Apsolutno se čuva samo mapa sposobnosti** — sve što operater danas može da uradi ili dobije.
- **Jedino tehničko ograničenje:** posle svakog PR-a projekat se kompajlira (inače pada ceo `run_vba`). Legacy se briše,
  ne ostavlja polomljen.
- **Jedna sesija = jedan korak.** Masovne mehaničke provere se rade spolja, po promptu.

## Gde smo

| Stavka | Status |
|---|---|
| PR0–PR6 (ugovor, šema iz koda, skele Zbirna/Otkup/Otpremnica, Otkup cutover) | ✅ |
| Kapija odluke §14.1 (nastavak u istom repou) | ✅ |
| #333 pre-flight PR7 · #334 čitaoci vrednosti otkupa na stavke | ✅ |
| #335 alat `tools/popis_citalaca.py` + odluke 16.09 | ✅ |
| **Mapa sposobnosti** | ✅ A–F gotove (`docs/DOMEN/mapa_sposobnosti_ulazi/`); sledeći korak = spajanje u `docs/DOMEN/MAPA_SPOSOBNOSTI.md` sa ispravkama iz liste ispod |
| Odluke domena (ispod) | ⏳ |
| Nova tabela PR-ova po novom modelu | ⏳ |
| Kod slajsova (otpremnica, zbirna, prijemnica, faktura, paleta, sledljivost, brisanje) | ⏳ |

## Sledeći korak: mapa sposobnosti

0. Gotovo: **A Dokumenti** — `docs/DOMEN/mapa_sposobnosti_ulazi/A.md` (38 sposobnosti, 27 NEPROVERENO).
0b. Gotovo: **B Storno/Oporavak** — `docs/DOMEN/mapa_sposobnosti_ulazi/B.md` (48 sposobnosti, 20 NEPROVERENO).
0c. Gotovo: **C Izveštaji/Sledljivost/Palete** — `docs/DOMEN/mapa_sposobnosti_ulazi/C.md` (62 sposobnosti, 44 sa NEPROVERENO).
0d. Gotovo: **D Fakture/Banka/Novac/Agro/Analiza** — `docs/DOMEN/mapa_sposobnosti_ulazi/D.md` (73 sposobnosti, 16 sa NEPROVERENO).
0e. Gotovo: **E Sync/PWA/GAS/izvozi** — `docs/DOMEN/mapa_sposobnosti_ulazi/E.md` (81 sposobnost, 69 sa NEPROVERENO, 6 pauziranih).
0f. Gotovo: **F1 Matični podaci i prijava** — `docs/DOMEN/mapa_sposobnosti_ulazi/F.md`
   (52 sposobnosti, 18 sa NEPROVERENO; sve presude „ne“ — F1 je jedina oblast bez zavisnosti od starog modela).
0g. Gotovo: **F2 Admin/podešavanja/setup/ažuriranje/integritet/health/makroi** — isti fajl, sekcije F2a–F2f
   (43 sposobnosti F-053…F-095, 42 sa NEPROVERENO; 206 makroa razvrstano). Presude „da“: integritet (15 od 22
   provere), health (`Check_CoreTablesAndColumns`, `Check_OtkupOtpremnicaCrossZbirnaLinks`,
   `Check_DocumentSoftDeleteReferences`, `Check_GoogleSyncMasterSchema`), `BackfillPrijemniceHladnjaca`,
   `BackfillOtkupBrojOtpremnice`, `BackfillDeteZbirnaGeneracija`, `PaletaAdjust_Prompt`, migracija iz starog fajla.
1. Ulazi su gotovi: `A.md` … `F.md` u `docs/DOMEN/mapa_sposobnosti_ulazi/`.
2. **Sledeći korak:** sesija koja ih spaja u `docs/DOMEN/MAPA_SPOSOBNOSTI.md`, unosi ispravke iz liste ispod,
   poravna kolonu `KO` (ima je samo F) i proverava pokrivenost ulaznih tačaka alatom
   (`python tools/popis_citalaca.py --procedura modul.Procedura` za lanac i status). Ulazne fajlove ne menjati.

## Ispravke za PR spajanja mape (iz pregleda izlaza A–F)

Mehanički pregled (reference, format presude, imena na linijama) je urađen; ovo su preostale ispravke koje sesija
spajanja unosi u `MAPA_SPOSOBNOSTI.md`. Ulazne fajlove ne menjati.

- **A:** dopune — A-015 štampa po broju bez stanice (AUD-057); A-013 KPI čita kolonu 5 (AUD-056);
  A-020 `Reassign` piše i `Otkup.BrojZbirne` (`modDokumenta.bas:6943-6945`).
- **B:** B-034 `modScrStorno.bas:1370-1377` je van fajla → `StornirajBlokoveAko:1255` (~1300-1312);
  `tblPaletaStavke` → `tblPaletaStavka`; opsezi `modStornoFlow.bas:477-516`, `:518-576` i `modStorno.StornoOtkup:220-241`
  suziti na ≤20 linija; presuda B-009 = „ne“, B-036 = „da“.
- **C:** neescapovan `|` lomi redove C-005, C-006, C-019, C-026, C-029, C-031; C-037 dopuniti — `Reassign` piše
  `BrojZbirne`; C-046 ostaje NEPROVERENO.
- **D:** `modUiScreens.ScrRedovi` → `ScrRows` (`:113`); presude „vidi D-xxx“ u D-003, D-036, D-037, D-044, D-045, D-046,
  D-048 zameniti sa da/ne/delimično.
- **E:** presude „oblast E-0xx“ (E-030, E-041, E-042, E-065, E-072, E-074, E-076, E-077) zameniti sa da/ne/delimično
  (npr. E-030 = da, push zavisi od E-027); spojiti GAS/PWA parove u jednu sposobnost sa dve implementacije:
  E-050/E-079, E-051/E-080, E-043/E-067, E-044/E-062, E-018/E-047/E-068; `otkupni-list.js:237` ne sadrži `signedAt`;
  u NEPROVERENO razlog „web sesija“ → „traži Google/OAuth“.
- **F1:** tabela ima dodatnu kolonu `KO` (pri spajanju poravnati sa ostalim oblastima); `tblVrstaGP` →
  `tblVrstaGotovihProizvoda`; linije: `T_Faza_PrijavaNeGradiLjusku:6405` → `:6409`, `PrimeniNovaPrava:4610` → `:4619`,
  `AlatkaSme:4479` → `:4480`, `NazadUAplikaciju:296` → `:421`; opsezi >20 linija suziti: `modBusinessFlowProTests.bas:1716-1737`,
  `modLicenseTests.bas:46-105`, `modMain.StartApp:136-158` i `:215-236`, `modUiScreens.bas:161-183`,
  `modMaticniEkran.bas:1088-1126`. Nalazi (provereni): F1 bez ijedne zavisnosti od starog modela; makroi
  `modMain.OpenExcel`/`CloseExcel` zaobilaze pravo `OtvoriExcel`; kolona statusa se proba nad sveskom umesto iz kanona
  (`tblTipAmbalaze`, `tblTipPalete`, `tblKutije`, `tblKese` imaju `Aktivan`), `RokMeseci` se ne nudi na ekranu; parcela
  status „Da“ pri unosu vs „Aktivan/Neaktivan“; ručna geo tačka se beleži kao `GeoSource = "selenium"` (`modGeoParcele.bas:78`).
- **F2:** nema mehaničkih ispravki iz ovog prolaza. Nalazi (provereni): `Check_CoreTablesAndColumns:109` nosi
  ručno kucan spisak kolona i traži linijska polja na zaglavljima — već pada na ispravnoj svesci (AUD-055), dok
  `Check_OtkupPaymentConsistency:505` i `Check_KooperantOtkupReconciliation:600` u istom modulu vrednost već
  čitaju sa stavki; 15 od 22 provere integriteta postoji samo zbog starog modela (`BrojZbirne` kao veza u četiri
  tabele, linijska polja, `Otkup.OtpremnicaID`), a sedam koje presuđuju „ne“ su palete i prerada — jedine celine
  već po novom modelu; `BackfillPrijemniceHladnjaca:439` radi po ključu `BrojZbirne|Klasa`; `PaletaAdjust_Prompt`
  vezuje palete na prijemnicu po poslovnom broju, ne po `PrijemnicaID`; tvrdu branu administracije nose samo
  `modAdmin` i `modPodesavanja` — `SetupNewPC`, `RunSelfUpdate`, `PublishReleaseToDrive`, `RollbackReleaseTo`,
  `OcistiTabele` i `MigrirajPodatkeIzStarog` se iz Alt+F8 pokreću bez provere prava.
- **F2:** reference i format čisti (43 sposobnosti, F-053..F-095). Pri spajanju: popravka podataka starog modela nije
  sposobnost korisnika (pravilo „bez migracija i backfill-a“) — `MigrirajPodatkeIzStarog` (F-061),
  `BackfillDeteZbirnaGeneracija` (F-091), `BackfillOtkupBrojOtpremnice` (F-092), `BackfillPrijemniceHladnjaca` (F-090) i
  naknadno usklađivanje paleta po broju prijemnice (F-093) prebaciti u listu za brisanje, ne u mapu; `OcistiTabele`
  (F-062) ostaje. Provereni nalazi: `Check_CoreTablesAndColumns` ručno traži linijska polja na zaglavljima
  (`modProductionHealthCheck.bas:120-137`) i pašće na svakom slajsu — spisak treba da dođe iz kanona; 15 od 22 provere
  integriteta postoje samo zbog starog modela; makroi `SetupNewPC`, `RunSelfUpdate`, `PublishReleaseToDrive`,
  `RollbackReleaseTo`, `OcistiTabele`, `MigrirajPodatkeIzStarog` nemaju proveru prava.
- **Nalazi E za redosled slajseva:** `ExportOtkupiAll` hrani ceo PWA menadžment i pregled otkupca a čita zaglavlje;
  `btnSync` uvek javlja neuspeh dok je izvedeni lanac pauziran; dodela vozača iz PWA završava kao terminalni
  `Duplicate` (`modMasterSync.bas:1775`); ekran vozača filtrira po `Otkup.VozacID` (`gas/Code.gs:1962`).

## Otvorene odluke domena (posle mape)

- ambalaža otpremnice: ko i kada knjiži izlaz gajbi stanica → vozač (`DOMAIN GAP`);
- `cenaII` kad zaglavlje otpremnice nosi jednu `PredlogCena`;
- ostaje li kucani bruto;
- ko pravi radnju „Izdaj“ (`IzdajOtpremnicu_TX` nema pozivaoca);
- PWA vozač i otpremnica (`src/js/features/vozac/zbirna.js`).

## Alati i kapije

- Popis starog modela i DUAL READ: `python tools/popis_citalaca.py` (prag slajsa: nula živih referenci na kolone starog modela).
- Pre push-a: `python tools/vba_check.py`, `python tools/who_writes.py --check` i `--check-ownership`,
  `python tools/gen_schema_module.py --check`; ponašanje: `python tools/run_vba.py --suite <ime>`.
- Poznati živi kvarovi van refaktora: `docs/KNOWN_ISSUES.md` AUD-055..057.
