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
+ `UiSinkEvent` dispatcher; vidi `frmDokumenta`/`frmOtkupAPP`).

## NOVE `Private WithEvents` deklaracije u FORMAMA su ZABRANJENE

Dodavanje event-sink deklaracije u formu **lomi self-update code-merge te forme**
(krivac za crash 2.16.1 → 2.21.0; `docs/SELF_UPDATE.md` zamka #11).

Zatečeni pre-2.16.1 form-WithEvents su **zamrznuti** — ne diraj i ne dodaji nove:

- `frmDokumenta` — `m_tgl*` / storno-pregled / recovery dugmad
- `frmAgrohemija` — „Pocetni dug"
- `frmPalete`
- `frmBankaExportPregled`
- `frmIzvestaj`

## Matični podaci (UI)

`frmMaticniPodaci` + `frmStammdaten` (`Select Case Me.Tag`) + `modMaticniLookups`
(data-driven meni).
