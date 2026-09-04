---
paths:
  - "src-vba/frm*.frm"
  - "src-vba/cls*.cls"
  - "src-vba/modOtkupBlok.bas"
  - "src-vba/modMaticniLookups.bas"
  - "src-vba/modPaletniListUI.bas"
  - "src-vba/mod*UI.bas"
---

# Forme, `.frx` i runtime kontrole

> Preseljeno iz `CLAUDE.md` §3/§4. Učitava se kad se dira UI sloj.

## Koje forme ostaju

Cilj plana iz `docs/UI_MIGRACIJA_KATALOG.md` §27 su **četiri** forme:
`frmOtkupUI` (ljuska), `frmLogin`, `frmSplash`, `frmExcelMini`.

Posle koraka 6 ih je **sedam** — uz te četiri:

| Forma | Zašto je još tu |
|---|---|
| `frmSEF` | SEF upravljanje (`PrepareResubmit`, batch radnje, event log) nije preneto ni na jedan ekran (§8.7). Čeka **ekran**, ne korak. Ulaz je `frmOtkupAPP.btnInvoicing`, oblast `OBL_FAKTURISANJE`. |
| `frmMarza` | čeka ekran koji je zamenjuje (`modScrMarza` ne postoji) |
| `frmOtkupAPP` | poslednja — host za `frmSEF` i preostale `ReturnToDashboard` pozive |

**Te tri se drže međusobno**, i to je jedini razlog zašto ijedna još stoji:
`frmSEF` zove `frmOtkupAPP.ReturnToDashboard` i otvara se njegovim dugmetom, pa
host ne može pre nje. Nijedna od tri ne čeka „korak" nego **kod koji treba
napisati** — dva ekrana i gašenje dve legacy grane zatvaranja u `modAdmin` i
`modPodesavanja`.

Pravilo koje iz toga sledi: **nova kontrola ne ide u legacy formu**, nego na
ekran ljuske (`modScr*`). Izmena legacy forme se radi samo kad je pravilo iz
§5 `otkup-i-dokumenta.md` izričito traži, ili u koraku koji tu formu uklanja.

## `.frx` se ne dira kao tekst

Kontrole formi žive u binarnom `.frx` (`.frm` ima samo `OleObjectBlob`). Nove
kontrole se **NE** dodaju editovanjem teksta — dodaju se u runtime-u
(`Controls.Add` + WithEvents), kao postojeći `modOtkupBlok`/`clsBlokUI`.

Caption koji živi samo u `.frx` menja se u dizajneru forme, ne u kodu.
Pri re-importu `.frm` uvek ide sa svojim `.frx` parom — `ImportAllVBA` preskače
formu bez `.frx` para.

## Dinamičke kontrole (bez `.frx`)

`Controls.Add` + WithEvents klasa: `clsBlokUI`/`modOtkupBlok`. Drugi primer je
bio `clsLookupMenuBtn`/`modMaticniLookups` — klasa je otišla u koraku 5 sa
`frmMaticniPodaci`, jedinom formom koja je taj meni gradila.

Za form-hostovane runtime kontrole koristi **generički `clsUiSink`** (`WireSink`
+ `UiSinkEvent` dispatcher; vidi `frmOtkupAPP`).

## NOVE `Private WithEvents` deklaracije u FORMAMA su ZABRANJENE

Dodavanje event-sink deklaracije u formu **lomi self-update code-merge te forme**
(krivac za crash 2.16.1 → 2.21.0; `docs/SELF_UPDATE.md` zamka #11).

Zatečeni pre-2.16.1 form-WithEvents su bili **zamrznuti** — spisak se samo
skraćivao, red po red, kako je koja forma odlazila: `frmDokumenta` (korak 2),
`frmAgrohemija` / `frmPalete` / `frmIzvestaj` (korak 3), `frmBankaExportPregled`
(korak 4).

**Spisak je sada PRAZAN.** Zabrana iznad time ne slabi nego postaje apsolutna:
nijedna preostala forma nema nijednu `WithEvents` deklaraciju, i tako mora da
ostane. Izuzetka više nema — svaki novi form-`WithEvents` je nov kvar, ne
nastavak zatečenog stanja.

Runtime kontrole u formama idu kroz omotače: `clsFlatBtn` (ljuska), `clsUiSink`
(generički sink), `clsBlokUI` (`modOtkupBlok`).

## Matični podaci (UI)

**Ljuska (važi):** ekrani `MAT_PARTNERI` / `MAT_ROBA` / `MAT_PAKOVANJE` /
`MAT_KORISNICI` (`modMaticni*` + `modScrMat*`) i paneli `MAT_PODESAVANJA` /
`MAT_ADMIN` (`modUiPanel`). Ulaz je prekidač sekcija u zaglavlju ljuske; pravo
na oblast `MaticniPodaci` odlučuje šta je od toga dostupno.

**Legacy je otišao u koraku 5** (§27.13): `frmMaticniPodaci`, `frmStammdaten`,
`clsStmBtn` i `clsLookupMenuBtn`. `modMaticniLookups` **ostaje** — ljuska koristi
`MaticniSekcije`, `MaticniSekcijeGrupisano` i `MaticniMenu_Release`; otišla je
samo njegova polovina koja je gradila stari meni.
