---
paths:
  - "src-vba/modConfig.bas"
  - "src-vba/modDataAccess.bas"
  - "src-vba/modSchemaGuard.bas"
  - "src-vba/modSetup.bas"
  - "src-vba/modPodesavanja.bas"
  - "src-vba/modHelpers.bas"
  - "src-vba/modArrayUtils.bas"
  - "src-vba/sConfig.doccls"
  - "src-vba/modSchema.bas"
  - "schema/schema.json"
  - "tools/gen_schema_module.py"
  - "tools/schema_diff.py"
  - "docs/DOMEN/WRITE_OWNERSHIP.json"
  - "tools/who_writes.py"
---

# Podaci, šema tabela i config

> Preseljeno iz `CLAUDE.md` §3/§4.

## Gde šta živi (ne praviti paralele)

| Oblast | Gde |
|---|---|
| Tabele / kolone / konstante | `modConfig.bas` (`TBL_*`, `COL_*`) |
| Pristup podacima | `modDataAccess.bas` (`GetTableData` / `GetColumnIndex` / `UpdateCell` / `AppendRow` / **`DeleteRow`** / `GetNextID` / `LookupValue`) |
| Filter/sort/util nad nizovima | `modArrayUtils.bas` (`FilterArray`, `SortArray`), `modHelpers.bas` (`Nz` / `NzToText` / `ExcludeStornirano` / `FillCmb`) |
| Setup / šeme | `modSetup` (`SetupNewPC`, `Ensure*Schema`; `SetupPopplerInteractive` / `SetupBankFoldersInteractive` pickeri; `RunSetupHealthCheck` uklj. živi `CheckServerLink` / `TestServerLink`), first-run kapija u `StartApp` (nudi `SetupNewPC` dok `APP_SETUP_COMPLETED != DA`), Admin dugmad `modAdmin` (health/googleauth/ensure), dijagnostika `DebugKoloneTabele` |

## Šema dolazi iz koda — `schema/schema.json` je kanon

> **Obrnuto od pravila koje je ovde stajalo do PR #302.** Do tada su spiskovi
> kolona osnovnih tabela živeli **isključivo u `.xlsm`** — pa se prazna sveska
> nije mogla rekonstruisati, a obrisana kolona se videla tek kao pad upisa
> satima kasnije. Doslovno iz `tools/make_fixture.py`: *„osnovna šema … ne
> postoji nigde u kodu"*.

```
schema/schema.json          <- KANON, u gitu
        |  tools/gen_schema_module.py
        v
src-vba/modSchema.bas       <- generisan artefakt, ne menja se rukom
        |  EnsureAllTables / VerifySchema / SchemaReadyOrFail
        v
.xlsm                       <- posledica
```

| Kad | Šta |
|---|---|
| Menjaš šemu | izmeni `schema/schema.json`, pa `python tools/gen_schema_module.py` |
| Pred commit | `python tools/gen_schema_module.py --check` (CI kapija) |
| Pred uvoz u zatečenu svesku | `python tools/schema_diff.py "<sveska>"` |
| Sveska odstupa | `Alt+F8 → EnsureAllTables` (samo **dodaje**; ne briše i ne premešta) |
| Inspekcija tuđe sveske | `tools/dump_schema.py` — **samo čitanje**, nikad izvor kanona |

**Redosled kolona je deo šeme.** `AppendRow` piše **poziciono**
(`modOtkup.SaveOtkup` gradi goli `Array(...)` sa 22 vrednosti), pa kolona
ubačena u sredinu tiho šalje vrednosti u pogrešne kolone — gore od pada upisa.
Nove kolone idu **na kraj**. Otisak (`SchemaCheckOnStart`) računa nad
**uređenim** kanonskim prefiksom, pa preraspored vidi; `schema_diff` na razliku
redosleda **blokira uvoz**. `EnsureAllTables` redosled **ne popravlja** —
premeštanje kolone u tabeli sa podacima bi pomerilo vrednosti.

**Provera na startu je fail-soft, kapija pred upisom je tvrda.**
`StartApp` zove `SchemaCheckOnStart` (otisak, ~10 ms) i samo loguje — pogrešna
šema ne sme da zaključa aplikaciju usred sezone. Tvrdo staje
`modSchema.SchemaReadyOrFail`, pred sam upis, gde pogrešan redosled stvarno može
da pošalje vrednosti u pogrešne kolone.

**Šta i dalje važi:** instalacije se razlikuju (schema drift). Razlika je što se
drift sada **meri i leči iz koda**, umesto da se pretpostavlja. PRE upisa i dalje
proveri stvarne nazive kolona (`Alt+F8 → DebugKoloneTabele`). Naučeno:

- `tblStanice`: telefon je u koloni `Kontakt` (**NE** `Telefon`); kontakt =
  `Ime` / `Prezime` / `PIN`.
- `tblKulture`: `KulturaID | VrstaVoca | SortaVoca | GajbicaPoPaleti` (**NEMA**
  `Aktivan`).
- `tblOtkup` / `tblOtpremnica` / `tblPrijemnica` / `tblFakturaStavke`: količina je
  ASCII `Kolicina` (**NE** `Količina`); koristi `COL_*_KOLICINA`, ne hardkoduj
  dijakritiku (bio `RunProductionHealthCheck` bug).

**Pozicijski `AppendRow` zavisi od redosleda kolona** — bezbedan samo ako je
redosled potvrđen. Za polja čiji redosled nije siguran koristi upis **po imenu**
(`UpdateCell` / `GetColumnIndex`).

Od PR #302 to više nije samo upozorenje: redosled je u kanonu, otisak ga meri, a
`SchemaReadyOrFail` staje pred upis. Ali pravilo ostaje — kapija štiti od
**zatečene** sveske, ne od novog koda koji pogrešno složi niz.

## Vlasništvo nad upisom (A11)

`docs/DOMEN/WRITE_OWNERSHIP.json` imenuje ko sme da piše koju tabelu.
`python tools/who_writes.py --check-ownership` obara CI na svakog novog pisca.

- **`row_owner`** sme da menja poslovne redove; **`schema_owner`** sme da napravi
  tabelu ili kolonu, ali ne i red. `modSetup` sme da napravi `tblOtkup` — ne sme
  da upiše otkup.
- **Snapshot nije vlasništvo.** `AddTableSnapshot` znači „moja transakcija mora
  da ume da vrati ovu tabelu". Kapija meri **mutatore**: `AppendRow`,
  `UpdateCell` i `DeleteRow`, svaki i sa `Require` prefiksom.
- Lista je **račna**: zamrznuto zatečeno stanje, pa hvata **širenje**. Skraćuje
  se kroz PR-ove ka `cilj`-u. Ne proširuj je da bi prošao — zovi API vlasnika.

Pun ugovor: `docs/DOMEN/ARCHITECTURE_CONTRACT.md`.

## Brisanje reda je RAZRED UPISA, ne pomoćna radnja

`modDataAccess.DeleteRow(tblName, rowIndex)` postoji od PR #307. Do tada se
poslovni red **nikad nije brisao** — samo označavao `Stornirano = Da`.

**`Stornirano` i dalje važi za dokumente.** Otkup, otpremnica, zbirna,
prijemnica i novac se ne brišu: append-only + storno je ceo model sledljivosti
(A9, A13). Danas `DeleteRow` zovu **dva** mesta, oba nad DRAFT otpremnicom
(`modDokumenta`): stavka očekivanja i red članstva, oba uklonjena **pre
izdavanja**. Obrazloženje stoji uz sam primitiv: izvor uklonjen pre izdavanja
nikad nije bio deo dokumenta, pa tombstone ne bi čuvao ništa — samo bi naterao
svakog čitača sastava da filtrira redove koji nikad nisu važili.

- **Zovi ga kroz `RequireDeleteRow`, ne golo.** `DeleteRow` vraća `False` (ne
  diže grešku) kad tabele nema, kad je prazna ili kad je indeks van opsega;
  golo, neprovereno brisanje tada **tiho ne uradi ništa**.
- **Jednoznačnost je ODVOJENA provera.** `RequireDeleteRow` **ne** proverava
  koliko redova odgovara ključu — on prima **indeks**. Kad indeks dolazi iz
  pretrage, pre njega ide `RequireTacnoJedan`; inače je „nađi pa obriši" isti
  kvar kao „prvi pogodak pobeđuje" kod FK-ova (AUD-026).
- **Indeks stari.** Posle jednog brisanja svi indeksi iza njega se pomeraju. U
  petlji se ide **unazad**, ili se indeksi razrešavaju iznova.
- **Primitiv postoji zbog A11, ne zbog udobnosti.** Kapija meri upise po
  **imenu mutatora** (`who_writes.py`); `lo.ListRows(i).Delete` sakriven u telu
  modula bio bi mutacija koju registar vlasništva ne vidi.

Pre nego što dodaš treće mesto koje briše: `BEZ_STORNA` u `modSchemaGuard`
nabraja tabele kod kojih storno **ne postoji kao koncept** (šifarnici, stavke,
tabele članstva). Tabela koja nije na toj listi se ne briše bez odluke u
`docs/DOMEN/`.

## TRI config tabele — ČITANJE i UPIS moraju u ISTU tabelu

Inače polje „ne radi" (tiho, bez greške).

| Tabela | Šta drži | API |
|---|---|---|
| `tblSEFConfig` | poslovni + **Google/PWA + SEF** kredencijali | `GetConfigValue` / `SetConfigValue` |
| `tblLocalConfig` | per-mašina: `PDFTOTEXT_EXE_PATH`, `BANKA_*_PATH`, `APP_SETUP_COMPLETED` | `GetLocalConfigValue` / `SetLocalConfigValue` |
| `tblConfig` | **legacy**, ne koristi se | — |

- Podešavanja editor rutira po `store` (`"sef"` / `"local"`) u `CfgAdd`; path polja
  imaju inline „…" browse dugme.
- Naučene greške: poppler upisan u SEFConfig a čitan iz Local;
  Google / `APP_SETUP_COMPLETED` čitani iz pogrešne tabele.
- `GetLocalConfigValue` na **praznu** vrednost vraća **default** — pa prazan
  `PDFTOTEXT_EXE_PATH` znači auto
  `<xlsm>\Tools\poppler\Library\bin\pdftotext.exe`.
