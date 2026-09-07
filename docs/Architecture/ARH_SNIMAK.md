# Snimak arhitektonskih mera

> **Generisan fajl -- ne menjaj rukom.**
> `python3 tools/arh_snimak.py --out docs/Architecture/ARH_SNIMAK.md`

Mereno: **2026-09-07**, commit **bde867f**.

Zakljucci i obrazlozenja su u `ARHITEKTURA_PLAN_OCENA.md` -- ovde su samo
brojke. Kad se razidju, vazi ovaj fajl: plan se pise rukom, snimak se meri.

## Obim

| Mera | Vrednost |
|---|---|
| Linija u `src-vba/` | 160.724 |
| Fajlova | 190 |
| Formi (`.frm`) | 1 |
| Klasa (`.cls`) | 11 |
| `Public` simbola u `.bas` | 2.752 |
| `Implements` | 0 |
| Modula koji pominju `SyncControl` | 2 |

## Fizicki pisci po tabeli

Moduli koji zovu `AppendRow`/`UpdateCell` nad tabelom. Ovo je metrika
Repository faze -- ne broj modula koji tabelu poslovno menjaju.

| Tabela | Pisaca | Moduli |
|---|---|---|
| `TBL_OTKUP` | **4** | `modDokumenta`, `modMasterSync`, `modOtkup`, `modSetup` |
| `TBL_KORISNICI` | **3** | `modAuth`, `modMaticniKorisnici`, `modSetup` |
| `TBL_ZBIRNA` | **3** | `modDokumentInvariant`, `modDokumenta`, `modMasterSync` |
| `TBL_FAKTURA_STAVKE` | **2** | `modFaktura`, `modUtovar` |
| `TBL_FAKTURE` | **2** | `modFaktura`, `modUtovar` |

**Tabela sa vise od jednog pisca: 5 od 24.**

## Sloj prikaza

`upis` mora ostati 0 u svakom redu -- to nadgleda `SLOJ_UPIS` u `vba_check`.
`TBL_`/`COL_` je preostali dug: ekran zna imena kolona. Broje se POJAVE,
ne linije -- jedna linija zna da nosi tabelu i tri kolone.

| Modul | LOC | `TBL_` | `COL_` | upis |
|---|---|---|---|---|
| `modScrDokumenti` | 2322 | 75 | 194 | 0 |
| `modScrIzvestaji` | 2903 | 52 | 34 | 0 |
| `modOtkupUI` | 8949 | 43 | 25 | 0 |
| `modScrFakture` | 2334 | 15 | 13 | 0 |
| `modScrOporavak` | 859 | 2 | 20 | 0 |
| `modScrAgro` | 1738 | 5 | 4 | 0 |
| `modScrBankaUvoz` | 1998 | 4 | 5 | 0 |
| `modScrPalete` | 1106 | 0 | 7 | 0 |
| `modScrStorno` | 1349 | 0 | 6 | 0 |
| `modScrAnaliza` | 68 | 0 | 0 | 0 |
| `modScrBankaNalozi` | 1580 | 0 | 0 | 0 |
| `modScrMatKorisnici` | 107 | 0 | 0 | 0 |
| `modScrMatPakovanje` | 83 | 0 | 0 | 0 |
| `modScrMatPartneri` | 86 | 0 | 0 | 0 |
| `modScrMatRoba` | 83 | 0 | 0 | 0 |
| `modScrSledljivost` | 1971 | 0 | 0 | 0 |

**Zbir `TBL_`+`COL_` u sloju prikaza: 504.**

## Transakcije

| Mera | Vrednost |
|---|---|
| `*_TX` ukupno | 87 |
| ...test helperi | 10 |
| ...produkcionih | 77 |
| ...od toga transakcioni omotac oko blizanca bez `_TX` | **42** (54%) |
| ...od toga blizanac je `Public` (vrata pored granice) | 32 |

| Invarijanta ADR-0003 B | Vrednost |
|---|---|
| Procedura sa `BeginTx` | 95 |
| ...deklarise `AddTableSnapshot` u istoj proceduri | 94 |
| **`AddTableSnapshot` bez `BeginTx` -- mora biti 0** | **0** |

