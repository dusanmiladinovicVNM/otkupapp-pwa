---
paths:
  - "src-vba/modBanka*.bas"
  - "src-vba/modScrBankaUvoz.bas"
  - "src-vba/modScrBankaNalozi.bas"
  - "src-vba/modNovac*.bas"
  - "src-vba/modTestBanka.bas"
  - "gas/bank-pdf-downloader/**"
---

# Banka — import izvoda i nalozi za isplatu

> Preseljeno iz `CLAUDE.md` §3 (v2.39.0) — učitava se kad se dira banka sloj,
> umesto u svakoj sesiji. Dopunjeno doplatom kooperantu (PR #460,
> `docs/DOMEN/DOPLATA.md`).

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
  v2.38.0/RF-09 iza dugmeta „Jaki ključevi", **nikad na ulazak u ekran**. Broj
  spremnih stavki daje read-only `CountStrongKeyReadyBankaImport` i stoji u
  **potvrdi**; na nuli se batch ne pokreće. Natpis dugmeta **nema `(N)`**: natpis
  radnje u ljusci je ključ kataloga poruka (`Scr_Radnje`), pa broj u njega ne
  staje bez izmene ugovora radnji. (Do v2.39 je ovde pisalo da natpis nosi `(N)`
  i da `_Activate` prebrojava — oboje je bilo netačno posle uklanjanja forme.)
- **Poziv na broj kooperanta ima tri vrste cilja** (DOP-ODL-30): broj otkupnog
  bloka, broj dokumenta doplate (`<broj bloka>-0z`) i broj specifikacije doplate
  (`x/ddmmyy-00y`). Sva tri su samo cifre, `/` i `-`, pa ih oblik ne razlikuje:
  `ResolveStrongKooperantByKeys` knjiži samo kad poziv pogodi **tačno jednu**
  vrstu, a `PozivViseVrsta` u `BimOtkupIzPozivaNaBroj` i `BimEfektivnaSpec` diže
  `ERR_BMAP_MANUAL_REQUIRED`. **NE vraćati „prvi pogodak pobeđuje"** — isplata
  doplate bi tiho završila na otkupnom bloku. Specifikaciju pisac
  (`MapBankaImportAsKooperantBlockCore`) raspoređuje na njene otvorene dokumente
  od najstarijeg bloka; novac se uvek vezuje **po dokumentu** (`DoplataID`),
  nikad po specifikaciji.
- **PRIKAZ I PISAC DELE JEDAN PLAN.** `BimAutoPlan` / `BimAutoPlanIzVrednosti`
  vraća `smer | partnerTip | partnerID | ciljTip | ciljID | izvor` **bez ijednog
  upisa**. Pisac ga izvršava (`AutoMapIncomingKupac`,
  `AutoMapOutgoingKooperantOrOM` više ne odlučuju, samo pišu), mreža ga prikazuje
  u koloni Predlog, potvrda ga čita naglas. **Nova grana automatskog mapiranja
  ide u plan, ne u pisca** — inače prikaz i akcija odu u različitim pravcima.
  Cena tog razlaza je izmerena: kolona je pisala „avans kupca" dok je `Auto`
  knjižio na fakturu (§27.14a).
- `samoJaki` **zaustavlja** lanac, ne filtrira ga posle računanja — inače greška
  u čitanju `tblStanice` obara i batch kome ta tabela ne treba.
- Čip i značka „jaki ključevi" znače **strogo jak ključ** i moraju se slagati sa
  `CountStrongKeyReadyBankaImport`. To je drugo pitanje od „šta će `Auto` uraditi"
  i ne spaja se s njim; čita se iz istog plana (`izvor = JAK`) da dve kopije
  istog uslova ne postoje.
- Radnja koja **knjiži novac nad jednim redom pita za potvrdu i imenuje cilj**.
  Plan bez cilja se ne prećutkuje: potvrda kaže da red ide na ručno.
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
  Doplata ide kao **specifikacija umesto njenih dokumenata** (DOP-ODL-27): jedan
  red i jedan nalog, poziv = broj specifikacije, iznos = otvoreni ostatak
  specifikacije (`BuildBlokIsplataList`). Dokument doplate bez specifikacije
  (knjižen pre P3) ide pod svojim brojem.
- **PDF specifikacija isplata:** `PrintIsplataSpecifikacija` →
  `modPrint.FillIsplataSpecSablon` (`ISPLATA_SPEC_PRINT_MODE`, default PDF).
- **Vezivanje virman avansa:** dugmad „Primeni avans na blok"/„(sel.)" →
  postojeći `ApplyAvansToOtkup_TX` (dotad samo auto pri snimanju novog otkupa u
  `modOtkup`; sad i za već otvorene blokove). Za doplatu: `ApplyAvansToKO_TX`
  (dokument) i `ApplyAvansToSpec_TX` (specifikacija — raspodela kao kod izvoda,
  u jednoj transakciji). Petlja vezivanja je zajednička (`VeziSlobodanAvans`),
  pa se pravilo slobodnog avansa ne kopira.
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

Novac vezan za dokument doplate (`tblNovac.DoplataID`) **nije slobodan avans**
(`JeSlobodanAvansKooperanta` gleda i tu kolonu). Storno doplate ga **odvezuje,
ne stornira** — virman ponovo postaje slobodan avans (DOP-05, isto pravilo kao
kod otkupa, `docs/DOMEN/ODLUKA_NOVAC_PRI_STORNU.md`).
