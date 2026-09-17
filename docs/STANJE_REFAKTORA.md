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
| Odluke domena | ✅ §14.8 (17.09.2026) |
| Nova tabela slajsova | ✅ §14.9 (17.09.2026) |
| Kod slajsova (otpremnica, zbirna, prijemnica, faktura, paleta, sledljivost, brisanje) | ⏳ |

## Sledeći korak: S3b — čitaoci otpremnice + panel blokova (posle provere S3a u Excelu)

1. Mapa: `docs/DOMEN/MAPA_SPOSOBNOSTI.md`. Odluke: plan §14.8. Slajsovi: §14.9. Pre-flight, S1a, S1b: §14.10.
2. **S1b-1 spojen** (#354). **S1b-2 urađen** (§14.10 „S1b-2 — urađeno“): desktop čitaoci otkupa na stavkama, stari panel
   `modOtkupBlok` i `clsBlokUI` obrisani, AUD-056 zatvoren; `x_otk_stavka` 63 → 22. Spojen (#355).
3. **S1b-3 urađen** (review #355, §14.10 „S1b-3“): bez mosta preko `Otkup.OtpremnicaID` — F1 radni sto otpremnice,
   bilans po otpremnici i ekran SLEDLJIVOST obrisani (vraćaju S3/S9); prefill ispravke fail-closed. Spojen (#356).
4. **Pravilo:** novi model se ne čita kroz staru vezu — takva sposobnost se briše, ne prevodi.
5. **S1c urađen** (§14.10 „S1c“): izvozi `OtkupPoOM`/`OtkupiAll`(+`OtkupiAllStavke`)/`SaldoOMDetail` i push ka stanici
   (`OTK_STAVKE`) čitaju kanonske stavke; raspored OTK kolona na jednom mestu; `AutoCreateOtpremniceFromPWA` obrisan;
   health iz kanona (AUD-055); `x_otk_stavka` = 0; push `OTK_STAVKE` idempotentan po `OtkupStavkaID`. Spojen (#357).
6. **S1d urađen** (§14.10 „S1d“): 8 kolona obrisano iz `tblOtkup` u kanonu i konstante `COL_OTK_*` iz `modConfig`;
   zatečena sveska ih gubi kroz self-heal (`ObrisiKolonuAko`). Spojen (#358).
7. **S1e urađen** (review #358, §14.10 „S1e“): identitet otkupa UI → `OtkupID` → mutacija — F1 red nosi `OtkupID`,
   štampa/storno/F8/hladnjača po ID-u; `StornoOtkupByBrDok_TX`, `OtkupIdsByBrDok`, `StornoSelectedBlocks_TX` obrisani;
   testovi sa dva zaglavlja po klasi obrisani. Spojen (#359). **S1 završen.**
8. **S2 urađen** (§14.11): banka radi po `OtkupID`-u — lista blokova nosi ID, poziv na broj se razreši JEDNOM
   (dvosmislen = ručno, ne raspodela preko dva otkupna mesta), pisač prima ID i dobija kapiju vlasništva i storna;
   `GetOtkupCandidatesForKooperantBlock`, `PlanBlokRaspodela` i ceo scope otkupnog mesta obrisani. Pre merge-a:
   pun `run_vba` + Compile.
9. **S2 spojen** (#360, `5610bfc9`) posle dva review kruga: vezan red novca nosi otkupno mesto **dokumenta**, a višak
   preko duga sme u avans samo uz **saglasnost pozivaoca** (`dozvoliVisakKaoAvans`, default False).
10. **S3 rez na pet koraka** (§14.12): živa ulazna tačka starog pisca otpremnice bila je **jedna**
    (`modDokUnos.OtpremnicaUpisi`), sve ostalo su čitaoci i pauzirani putevi.
11. **S3a urađen** (§14.12): F2 otvara **nacrt** (`CreateOtpremnicaDraft_TX`); `PredlogCena` je po klasi na stavci, a
    zaglavlje je više ne prima (ključ se odbija); **ambalaža se knjiži pri izdavanju**, ne na nacrtu; malina
    auto-zbirna pauzirana uz poruku (vraća S4). Stari pisac (`SaveOtpremnica*`) ostao bez živog pozivaoca —
    briše se u S3e. Pre merge-a: pun `run_vba` + Compile.
    **Review #361 P1:** uklonjen i `ZavrsiIspravkuAko FLOW_DOC_OTPREMNICA` — ispravka otpremnice (B-038) je
    **pauzirana**, ne prevedena: stari tok je preko `ReassignOtkupToOtpremnica_TX` pisao `Otkup.OtpremnicaID` i
    proglašavao nacrt bez izvora zamenom izdate otpremnice. Vraća S3c, po ID-u.
    **Pravilo od S3a:** nijedan nov kod ni test ne sme da zove `SaveOtpremnica*`.
12. **Sledeće:** S3b — čitaoci otpremnice na stavke + panel blokova nad `tblOtpremnicaIzvori` i radnja „Izdaj“.

## Alati i kapije

- Popis starog modela i DUAL READ: `python tools/popis_citalaca.py` (prag slajsa: nula živih referenci na kolone starog modela).
- Pre push-a: `python tools/vba_check.py`, `python tools/who_writes.py --check` i `--check-ownership`,
  `python tools/gen_schema_module.py --check`; ponašanje: `python tools/run_vba.py --suite <ime>`.
- Poznati živi kvarovi van refaktora: `docs/KNOWN_ISSUES.md` AUD-055..057.
