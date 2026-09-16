# Stanje refaktora „dokument = header + stavke“

> **Ulaz za svaku novu sesiju.** Kratko i ažurno — čita se umesto celog plana. Pun plan i istorijat odluka:
> `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` (odluke po datumu u §14.x; važeće: §14.7 „Odluke operatera 16.09“).
> Ažurira se na kraju svakog koraka, u istom commit-u.

**Ažurirano:** 16.09.2026, `main` `7ef2f7c2` (posle #337).

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
| **Mapa sposobnosti** | ⏳ u toku — oblasti A, B, C, D i E gotove (`docs/DOMEN/mapa_sposobnosti_ulazi/`), ostaje F |
| Odluke domena (ispod) | ⏳ |
| Nova tabela PR-ova po novom modelu | ⏳ |
| Kod slajsova (otpremnica, zbirna, prijemnica, faktura, paleta, sledljivost, brisanje) | ⏳ |

## Sledeći korak: mapa sposobnosti

0. Gotovo: **A Dokumenti** — `docs/DOMEN/mapa_sposobnosti_ulazi/A.md` (38 sposobnosti, 27 NEPROVERENO).
0b. Gotovo: **B Storno/Oporavak** — `docs/DOMEN/mapa_sposobnosti_ulazi/B.md` (48 sposobnosti, 20 NEPROVERENO).
0c. Gotovo: **C Izveštaji/Sledljivost/Palete** — `docs/DOMEN/mapa_sposobnosti_ulazi/C.md` (62 sposobnosti, 44 sa NEPROVERENO).
0d. Gotovo: **D Fakture/Banka/Novac/Agro/Analiza** — `docs/DOMEN/mapa_sposobnosti_ulazi/D.md` (73 sposobnosti, 16 sa NEPROVERENO).
0e. Gotovo: **E Sync/PWA/GAS/izvozi** — `docs/DOMEN/mapa_sposobnosti_ulazi/E.md` (81 sposobnost, 69 sa NEPROVERENO, 6 pauziranih).
1. Korisnik pušta `docs/DOMEN/PROMPT_MAPA_SPOSOBNOSTI.md` po oblasti (A Dokumenti · B Storno/Oporavak ·
   C Izveštaji/Sledljivost/Palete · D Fakture/Banka/Novac/Agro/Analiza · E Sync/PWA/GAS/izvozi · F Admin/makroi/integritet/setup)
   i čuva izlaze kao `A.md` … `F.md`.
2. Sesija: spaja ih u `docs/DOMEN/MAPA_SPOSOBNOSTI.md` i proverava pokrivenost ulaznih tačaka alatom
   (`python tools/popis_citalaca.py --procedura modul.Procedura` za lanac i status).

## Ispravke za PR spajanja mape (iz pregleda izlaza A–E)

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
