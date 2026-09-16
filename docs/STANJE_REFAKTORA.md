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
