# APR tržišni podaci

Ovaj direktorijum sadrži APR izvore i reproduktivan pipeline za procenu AgriX tržišta.

## Direktorijumi

- `Raw Data/` — originalni APR snapshot-i; ne menjati ručno.
- `Clean Data/` — podaci obogaćeni finansijama i pripremljeni za validaciju.
- `Processed/` — validirani podaci, quality izveštaji i analitički međurezultati.
- `Reports/` — tržišni workbook-i i Markdown sažeci za `04_MARKET.md`.
- `Scripts/` — Python pipeline.

Postojeći fajl `apr_veleprodaja_voca_povrca_prerada_konzervisanje_2024.xlsx` može ostati kao raniji očišćeni dataset. Novi pipeline generiše standardizovane fajlove sa datumom i metadata JSON zapisom.

## Pipeline

Pipeline ima četiri koraka:

1. `01_extract_companies.py` — preuzima APR registar i filtrira šifre delatnosti.
2. `02_enrich_financials.py` — dodaje finansijske podatke iz APR financial-statements endpointa.
3. `03_clean_validate.py` — normalizuje podatke i pravi data-quality izveštaj.
4. `04_analyze_market.py` — pravi tržišni Excel i Markdown sažetak.

Za kompletno pokretanje koristi se `run_pipeline.py`.

## Instalacija

Iz `Scripts` foldera:

```bash
python -m pip install -r requirements.txt
```

## Kompletno pokretanje

```bash
python run_pipeline.py
```

Podrazumevane šifre su:

- `1039` — ostala prerada i konzervisanje voća i povrća;
- `4631` — trgovina na veliko voćem i povrćem.

Druge šifre se mogu proslediti:

```bash
python run_pipeline.py --codes 1039 4631 4621
```

Ako APR endpoint lokalno ne prolazi TLS proveru, privremeni workaround je:

```bash
python run_pipeline.py --insecure
```

`--insecure` ne treba koristiti kao podrazumevanu opciju jer isključuje proveru HTTPS sertifikata.

## Pojedinačni koraci

```bash
python 01_extract_companies.py
python 02_enrich_financials.py
python 03_clean_validate.py
python 04_analyze_market.py
```

Svaki korak automatski bira najnoviji očekivani ulaz. Eksplicitni input može se zadati sa `--input`.

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

Analitički workbook sadrži:

- osnovni summary;
- prometne razrede;
- geografsku raspodelu;
- raspodelu po delatnosti;
- top 200 firmi;
- koncentraciju prihoda;
- kompletan analizirani skup.

## Važna ograničenja

- Šifre `1039` i `4631` ne obuhvataju nužno sve organizovane otkupljivače.
- Prihod nije direktna mera broja stanica, kooperanata, logistike ili GGAP potrebe.
- Financial-statements endpoint se trenutno tretira kao jedan zapis po firmi. Višegodišnja istorija nije potvrđena ovim pipeline-om.
- Pravilo za aktivan status mora se proveriti prema stvarnim vrednostima u APR datasetu.
- Rezultati su ulaz za TAM/SAM/SOM model, ali nisu sami po sebi konačan TAM, SAM ili SOM.

## Legacy skripte

`AllCompanies.py` i `FindAllFin.py` su zadržane kao kompatibilni wrapper-i. Novi razvoj treba da koristi numerisane skripte.

## Metadata pravilo

Uz svaki generisani dataset čuvaju se:

- vreme generisanja;
- izvorni URL;
- obuhvaćene šifre;
- broj redova;
- poznata ograničenja;
- SHA-256 ulaznog i/ili izlaznog fajla.

## Privatnost

Pre javnog postavljanja proveriti da dataset ne sadrži osetljive lične ili ugovorne podatke. Matični broj pravnog lica koristi se za deduplikaciju i spajanje APR podataka.
