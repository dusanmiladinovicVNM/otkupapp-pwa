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

Posle plana iz `docs/UI_MIGRACIJA_KATALOG.md` §27 u projektu ostaju **četiri**:
`frmOtkupUI` (ljuska), `frmLogin`, `frmSplash`, `frmExcelMini`. Sve ostale se
penzionišu po koracima §27.3; inventar i uslov po formi je §27.2.

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

`Controls.Add` + WithEvents klasa: `clsBlokUI`/`modOtkupBlok`,
`clsLookupMenuBtn`/`modMaticniLookups`.

Za form-hostovane runtime kontrole koristi **generički `clsUiSink`** (`WireSink`
+ `UiSinkEvent` dispatcher; vidi `frmOtkupAPP`).

## NOVE `Private WithEvents` deklaracije u FORMAMA su ZABRANJENE

Dodavanje event-sink deklaracije u formu **lomi self-update code-merge te forme**
(krivac za crash 2.16.1 → 2.21.0; `docs/SELF_UPDATE.md` zamka #11).

Zatečeni pre-2.16.1 form-WithEvents su **zamrznuti** — ne diraj i ne dodaji nove:

- `frmAgrohemija` — „Pocetni dug"
- `frmPalete`
- `frmBankaExportPregled`
- `frmIzvestaj`

**Ovaj spisak se samo SKRAĆUJE.** `frmDokumenta` je već otišao (korak 2, §27.10),
pa je red obrisan. Preostale četiri su u planu penzionisanja (§27.2):
`frmAgrohemija` / `frmPalete` / `frmIzvestaj` u koraku 3, `frmBankaExportPregled`
u koraku 4. Red se briše kad forma ode, i to u process PR-u — ne uz izmenu koda.

Kad spisak ostane prazan, zabrana iznad i dalje važi: nijedna od četiri forme
koje ostaju nema ni jednu `WithEvents` deklaraciju, i tako mora i da ostane —
runtime kontrole u njima idu kroz omotače (`clsFlatBtn`, `clsUiSink`).

## Matični podaci (UI)

**Ljuska (važi):** ekrani `MAT_PARTNERI` / `MAT_ROBA` / `MAT_PAKOVANJE` /
`MAT_KORISNICI` (`modMaticni*` + `modScrMat*`) i paneli `MAT_PODESAVANJA` /
`MAT_ADMIN` (`modUiPanel`). Ulaz je prekidač sekcija u zaglavlju ljuske; pravo
na oblast `MaticniPodaci` odlučuje šta je od toga dostupno.

**Legacy (odlazi u koraku 5):** `frmMaticniPodaci` + `frmStammdaten`
(`Select Case Me.Tag`) + `modMaticniLookups` (data-driven meni), zajedno sa
`clsStmBtn` i `clsLookupMenuBtn` ako ostanu bez pozivaoca.
