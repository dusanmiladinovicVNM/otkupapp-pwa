# Stanje refaktora „dokument = header + stavke“

> **Ulaz za svaku novu sesiju.** Kratko i ažurno — čita se umesto celog plana. Pun plan i istorijat odluka:
> `docs/REFAKTOR_DOKUMENT_HEADER_STAVKE.md` (odluke po datumu u §14.x; važeće: §14.7 „Odluke operatera 16.09“).
> Ažurira se na kraju svakog koraka, u istom commit-u.

**Ažurirano:** 17.09.2026 (mapa sposobnosti spojena).

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
| **Mapa sposobnosti** | ✅ spojena u `docs/DOMEN/MAPA_SPOSOBNOSTI.md` (387 sposobnosti; ulazi A–F ostaju u `docs/DOMEN/mapa_sposobnosti_ulazi/`) |
| Odluke domena (ispod) | ⏳ |
| Nova tabela PR-ova po novom modelu | ⏳ |
| Kod slajsova (otpremnica, zbirna, prijemnica, faktura, paleta, sledljivost, brisanje) | ⏳ |

## Sledeći korak: odluke domena

1. Mapa je gotova: `docs/DOMEN/MAPA_SPOSOBNOSTI.md` — 387 sposobnosti (A 38 · B 48 · C 62 · D 73 · E 76 · F 90),
   sa presudom zavisnosti od starog modela po redu, nalazima za redosled slajseva, NEPROVERENO i pokrivenošću
   (49 živih pomoćnih procedura starog modela bez pomena — lista u mapi).
2. Ispravke iz pregleda A–F primenjene su u `MAPA_SPOSOBNOSTI.md` (spojeno 5 GAS/PWA parova, izbačeno 5 popravki
   podataka starog modela: F-061, F-090..F-093).
3. **Sledeće:** odluke domena (ispod), pa nova tabela PR-ova po novom modelu, sa mapom kao spiskom obaveznih ishoda.

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
