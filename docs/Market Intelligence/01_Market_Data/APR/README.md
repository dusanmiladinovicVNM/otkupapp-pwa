# APR tržišni podaci

Ovaj direktorijum sadrži APR izvore i reproduktivan pipeline za procenu AgriX tržišta.

## Direktorijumi

- `Raw Data/` — originalni APR snapshot-i; ne menjati ručno.
- `Clean Data/` — podaci obogaćeni finansijama i pripremljeni za validaciju.
- `Processed/` — validirani podaci, quality izveštaji i analitički međurezultati.
- `Reports/` — tržišni workbook-i i Markdown sažeci za `04_MARKET.md`.
- `Scripts/` — Python pipeline.

Postojeći fajl `apr_veleprodaja_voca_povrca_prerada_konzervisanje_2024.xlsx` može ostati kao raniji dataset. Novi pipeline generiše standardizovane fajlove sa datumom i metadata JSON zapisom.

## Pipeline

1. `01_extract_companies.py` — preuzima APR registar i filtrira šifre delatnosti.
2. `02_enrich_financials.py` — dodaje APR finansijske podatke i normalizuje novčane vrednosti u RSD.
3. `03_clean_validate.py` — normalizuje podatke, proverava finansijsku jedinicu, klasifikuje statuse i pravi data-quality izveštaj.
4. `04_analyze_market.py` — pravi tržišni Excel i Markdown sažetak samo iz potvrđenih RSD vrednosti.

Za orkestraciju se koristi `run_pipeline.py`.

## Instalacija

Iz `Scripts` foldera:

```bash
python -m pip install -r requirements.txt
```

## Režimi rada

### Full — godišnje online osvežavanje

```bash
python run_pipeline.py
```

Isto kao:

```bash
python run_pipeline.py --mode full
```

Podrazumevane šifre su:

- `1039` — ostala prerada i konzervisanje voća i povrća;
- `4631` — trgovina na veliko voćem i povrćem.

Druge šifre:

```bash
python run_pipeline.py --mode full --codes 1039 4631 4621
```

### Offline — postojeći Excel bez APR API-ja

Za standardni Clean Data fajl:

```bash
python run_pipeline.py --mode offline --input "../Clean Data/apr_companies_financials_1039_4631_2026-07-22.xlsx"
```

Offline režim radi validaciju i analizu bez interneta.

### Report — ponovna analiza validiranog skupa

```bash
python run_pipeline.py --mode report
```

Koristi najnoviji `apr_market_validated_*.xlsx` iz `Processed/`.

## Finansijske jedinice — kritično pravilo

APR finansijski izveštaji iskazuju novčane podatke u **hiljadama dinara**. Pipeline zato koristi sledeći model:

- izvorne APR vrednosti čuva u kolonama `*_apr_000_rsd`;
- glavne novčane kolone (`ukupni_prihodi`, `kapital`, `poslovna_imovina`, dobitak i gubitak) normalizuje množenjem sa `1.000`;
- validirani i analitički fajlovi koriste isključivo pune **RSD** vrednosti;
- metadata čuva `source_financial_unit`, `normalized_financial_unit` i primenjeni multiplier;
- `04_analyze_market.py` odbija dataset bez potvrđene jedinice `RSD`.

Za stariji APR Clean Data fajl metadata automatski aktivira legacy konverziju:

```bash
python run_pipeline.py --mode offline --input "../Clean Data/apr_companies_financials_1039_4631_2026-07-22.xlsx"
```

Ako ulaz nema metadata, jedinica se može eksplicitno navesti:

```bash
python run_pipeline.py --mode offline --input "../Clean Data/neki_fajl.xlsx" --financial-unit thousand-rsd
```

Podržane vrednosti:

- `auto` — metadata/detekcija; podrazumevano;
- `rsd` — ulaz je već u punim dinarima;
- `thousand-rsd` — ulaz je u hiljadama dinara i množi se sa `1.000`.

## TLS politika

Pipeline prvo pokušava punu TLS verifikaciju. Ako samo `openapi.apr.gov.rs` vrati SSL validation grešku, zahtev se automatski ponavlja bez verifikacije, uz upozorenje i metadata zapis.

Strogi režim:

```bash
python run_pipeline.py --strict-tls
```

Eksplicitno insecure pokretanje:

```bash
python run_pipeline.py --insecure
```

Fallback nije dozvoljen za druge hostove.

## Status klasifikacija

APR statusi trenutno dolaze na ćirilici. Pipeline radi Unicode normalizaciju, transliteraciju i proverava negativna stanja pre aktivnog statusa:

1. stečaj;
2. likvidacija i prinudna likvidacija;
3. neaktivan, brisan, ugašen ili prestao;
4. aktivan ili registrovan;
5. ostalo / nepoznato.

Validirani workbook sadrži:

- `Status Values` — originalne APR vrednosti i broj firmi;
- `Status Categories` — zbir po kategoriji;
- `Quality Issues` — nevalidni identifikatori, duplikati i neklasifikovani statusi.

Ako postoje redovi, a aktivnih firmi je nula, pipeline prekida obradu. Analiza takođe odbija prazan tržišni scope.

## Izlazi

### Raw Data

- `apr_companies_<codes>_<date>.xlsx`
- odgovarajući `.metadata.json`

### Clean Data

- `apr_companies_financials_<codes>_<date>.xlsx`
- odgovarajući `.metadata.json`

### Processed

- `apr_market_validated_<codes>_<date>.xlsx`
- `apr_data_quality_<codes>_<date>.xlsx`
- metadata JSON

### Reports

- `apr_market_analysis_<codes>_<date>.xlsx`
- `apr_market_summary_<codes>_<date>.md`
- metadata JSON

Analitički workbook sadrži summary, statuse, prometne razrede, geografiju, delatnosti, top 200 firmi, koncentraciju prihoda i kompletan analizirani skup.

## Preporučeni ritam

- `full` — novi godišnji APR snapshot;
- `offline` — iteracije nad postojećim Excelom i pravilima;
- `report` — menjaju se samo agregacije ili format izveštaja.

## Važna ograničenja

- Šifre `1039` i `4631` ne obuhvataju nužno sve organizovane otkupljivače.
- Prihod nije direktna mera broja stanica, kooperanata, logistike ili GGAP potrebe.
- Financial-statements endpoint se trenutno tretira kao jedan zapis po firmi; višegodišnja istorija nije potvrđena.
- Svaka nova APR status vrednost mora biti pregledana u `Status Values`.
- Rezultati su ulaz za TAM/SAM/SOM, ali nisu automatski konačan TAM, SAM ili SOM.
- TLS režim, finansijska jedinica, pravila i hash vrednosti moraju ostati u metadata fajlovima.

## Legacy skripte

`AllCompanies.py` i `FindAllFin.py` su zadržane kao kompatibilni wrapper-i. Novi razvoj koristi numerisane skripte.

## Privatnost

Pre javnog postavljanja proveriti da dataset ne sadrži osetljive lične ili ugovorne podatke. Matični broj pravnog lica koristi se za deduplikaciju i spajanje APR podataka.
