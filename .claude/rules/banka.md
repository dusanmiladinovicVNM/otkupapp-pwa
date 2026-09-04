---
paths:
  - "src-vba/modBanka*.bas"
  - "src-vba/frmBanka*.frm"
  - "src-vba/modNovac*.bas"
  - "src-vba/modTestBanka.bas"
  - "gas/bank-pdf-downloader/**"
---

# Banka — import izvoda i nalozi za isplatu

> Preseljeno iz `CLAUDE.md` §3 (v2.39.0). Sadržaj nepromenjen — samo se sada
> učitava kad se dira banka sloj, umesto u svakoj sesiji.

## Banka import (izvodi)

`modBankaImport` (pull + `ImportBankaInbox_TX`), **multi-bank dispatch**
`DetectBank` + `Select Case` u `ParseBankaIzvodForImport` (deljeni 4-nivo
integritet + 17-kol staging; parser po banci — `modBankaImportParserPdfToText` =
Komercijalna, `modBankaProCredit`, `modBankaHalk`, `modBankaAlta` (`190-`,
fingerprint naslov „IZVOD BR."); svi preko `pdftotext`/Poppler), mapiranje
`modBankaMapiranje` → `tblNovac`, ekran `BANKA_UVOZ` (`modScrBankaUvoz`).

**Uvoz pokreće ULAZAK u ekran**, ne dugme: u legacy meniju ga je pokretao klik
na Banka, a u ljusci je ulazak taj klik (`Scr_Aktiviraj`). Zato je uvoz
**izričita radnja operatera** — automatska usmeravanja (start, preusmeravanje
posle zamene operatera) prosleđuju `ActivateScreen … , False` i kuku NE zovu.
To je ugovor, ne stil: `RELEASE_GATES` §85 i `ARCHITECTURE_REFERENCE` §288/§440.

- **Jaki ključevi** — `poziv na broj` = otkup/faktura, `tekući račun` — od
  v2.38.0/RF-09 iza dugmeta „Mapiraj jake ključeve (N)", **NE** na `_Activate`;
  `_Activate` samo prebroji preko read-only `CountStrongKeyReadyBankaImport`.
- Dedupe ključ uključuje **broj računa**.
- `Map*` imaju **smer guard** (`RequireBimSmer`).
- Blok sa 3+ otvorenih stavki diže `ERR_BMAP_MANUAL_REQUIRED` koju batch guta
  **po redu** (`AutoMapBankaImportRowBatch`), ne obara ceo `AutoMapAll`.
- Datumi izvoda se validiraju pre staging-a (`TryParseDateValue` round-trip,
  AUD-007).

Testovi: `RunBankaImportTestSuite` (`modTestBanka`) — ima tvrd fail-gate, vidi
`.claude/rules/testovi.md`.
GAS `gas/bank-pdf-downloader/` (Gmail → Drive).
Runbook: `docs/production-runbook-banka-import-setup.md`; novi parser:
`docs/development-banka-parser.md`.

## Banka nalozi (isplate)

`modScrBankaNalozi` + `modBankaExportPregled` — pregled otvorenih blokova,
per-blok „Isplatiti" override; runtime combo (`.frx` se ne dira) — „Kooperant"
filter radi i na unos i kao dd, substring, prune override-a protiv PUNE liste;
„Sa računa" = izbor računa firme (do 4 zasebna polja `BANKA_NALOG_RACUN_1..4` u
Podešavanjima, spojena kroz `BankaNalogRacuniCSV`; legacy `;`-lista
`BANKA_NALOG_RACUNI` + `SELLER_ACCOUNT` kao fallback), prikaz banke
`BankaNazivZaRacun`.

- **CSV nalozi za prenos:** `GenerisiNalogeCSV` → `Nalozi za banku\` (platilac
  `SELLER_NAME`/`SELLER_ACCOUNT`; **poziv na broj = broj bloka** → auto-map pri
  uvozu izvoda; šifra/svrha `BANKA_NALOG_*`, Podešavanja grupa „Banka / nalozi").
- **PDF specifikacija isplata:** `PrintIsplataSpecifikacija` →
  `modPrint.FillIsplataSpecSablon` (`ISPLATA_SPEC_PRINT_MODE`, default PDF).
- **Vezivanje virman avansa:** dugmad „Primeni avans na blok"/„(sel.)" →
  postojeći `ApplyAvansToOtkup_TX` (dotad samo auto pri snimanju novog otkupa u
  `modOtkup`; sad i za već otvorene blokove).
- **Bez upisa u `tblNovac` za isplate** — isplata se knjiži tek uvozom izvoda
  (avans upis je zaseban, veže `OtkupID` na postojeći `NOV_VIRMAN_AVANS_KOOP`).

### Saldo je fail-closed (v2.39.0 / RF-10, AUD-026)

Override preživljava reload ali se pri svakom rebuild-u usklađuje sa otvorenim
(`ClampOverridesToOpen` — nestao/zatvoren blok → briše, veći → spušta, **manji
ostaje**), a `GenerisiNalogeCSV` pred upis čita **svež** saldo i kroz
`ValidateNalogSaldo` odbija CEO fajl kad ijedan nalog prelazi otvoreno (razlog
kroz `outOdbijeno`; blok kog nema među otvorenima = otvoreno 0).

Iznosi se porede u **cent-domenu** (`ZaokruziNovac`, half-up), **bez epsilon
tolerancije** — prag `+ 0.01` je propuštao preplatu od punog centa; isto pravilo
za unos u formi i `CsvIznos`. **NE uvoditi novu granicu ni nov helper.**

Identitet primaoca se **NE** rešava preko `BuildLookupDict` („prvi pojav
pobeđuje" nad sirovom tabelom) — `BuildOtkupOwnerIndex` daje vlasnika samo za
jednoznačne `OtkupID`-eve, jer bi kod duplikata čiji je jedan red isplaćen (pa
nevidljiv u otvorenima) nalog otišao POGREŠNOM kooperantu.

Isti zahtev važi i za **`KooperantID`** (`BuildKooperantTekuciRacunCache` →
`ERR_ISPLATA_DUPLI_KOOPERANTID`): `TekuciRacun` je `RacunPrimaoca`, a
health-check duplicate provera **ne pokriva** `KooperantID`.

Saldo-mapa je **striktna na `OtkupID`** (`BuildOpenAmountDict` diže
`ERR_ISPLATA_DUPLI_OTKUPID` / `ERR_ISPLATA_PRAZAN_OTKUPID`) — `GetOpenOtkupi`
vraća red po red, pa bi assignment pustio saldo jednog bloka da odobri nalog
drugog; **NE vraćati na „poslednji pobeđuje"**.

„Primeni avans" ide kroz `modNovac.ApplyAvansToOtkup`, koji od v2.39.0 odbija
**dupli `OtkupID`** (target mora biti jednoznačan — `FindRows(...)(1)` je ranije
značilo „prvi red pobeđuje"); guard je u core-u, ne u formi. Avansi se broje po
**stvarno proknjiženom iznosu** (`ApplyAvansToOtkup_TX` `ByRef`, RF-02) — `True`
sam po sebi ne znači da je nešto vezano.
