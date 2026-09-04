---
paths:
  - "src-vba/modAgrohemija.bas"
  - "src-vba/modAgrohemijaTests.bas"
  - "src-vba/modAgroUnos.bas"
  - "src-vba/modAmbalaza.bas"
  - "src-vba/modCenovnik.bas"
  - "src-vba/sAmbalaza.doccls"
---

# Agrohemija / magacin, ambalaža i cene

> Preseljeno iz `CLAUDE.md` §3.

## Agrohemija / magacin

`modAgroUnos` + `modAgrohemija`: `SaveMagacin` piše ledger `MAG_ULAZ` /
`MAG_IZLAZ`, stanje kroz `GetMagacinStanje`, dug kroz `GetAgrohemijaDug`.
Ekran je `AGRO` (`modScrAgro`); `frmAgrohemija` je otišla u koraku 3 (§27.11).

> **AUD-040 stoji na granici unosa**, ne u jezgru: `modAgroUnos` mora da
> prosledi korpa cenu kao `overrideCena` i `allowZeroValue` iz korpe.
> `SaveMagacinCore` unit-test to ne hvata — kvar je bio u pozivaocu. Meri ga
> `Test_UnosWiresBasketPrice` (`modAgrohemijaTests`), nad izvorom modula.

- **Izlaz opciono bez parcele** kad je `PRACENJE_PARCELA` OFF (`IsPracenjeParcela`,
  isti flag koji čita i ekran dokumenata; smart-doza se tada preskače).
- **Početni dug kooperanta (migracija)** = rezervisani virtuelni artikal
  `ART_POCETNI_DUG` (`modConfig`) + `BookPocetniDug` →
  `SaveMagacin(... allowNoStock:=True)`. Artikal je izuzet iz combo-lista i iz
  `GetMagacinStanje` — **NE dirati to izuzimanje**.
- **PWA `ExportMagacinKoop` ga još NE izuzima** (KI-006, `docs/KNOWN_ISSUES.md`).

Smoke suite: `RunAgrohemijaSmokeSuite` (bez fail-gate-a — vidi
`.claude/rules/testovi.md`).

## Ambalaža

`modAmbalaza` — ledger istog tipa.

## Cena — DVA modela, ne mešati

| Model | Gde | Za šta |
|---|---|---|
| single-current po artiklu | `tblArtikli.CenaPoJedinici` (inline `LookupValue`) | agrohemija |
| append-only istorija | `tblCenovnik` → `modCenovnik` (`GetVazecaCena` / `AddCena`) | otkup voća |

Ne uvoditi treći model i ne čitati cenu voća iz `tblArtikli` (ni obrnuto).
