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
3. `03_clean_validate.py` — normalizuje podatke, klasifikuje statuse i pravi data-quality izveštaj.
4. `04_analyze_market.py` — pravi tržišni Excel i Markdown sažetak.

Za orkestraciju se koristi `run_pipeline.py`.

## Instalacija

Iz `Scripts` foldera:

```bash
python -m pip install -r requirements.txt
```

## Režimi rada

### 1. Full — godišnje online osvežavanje

```bash
python run_pipeline.py
```

Isto kao:

```bash
python run_pipeline.py --mode full
```

Radi sva četiri koraka: APR registar, finansije, validaciju i analizu.

Podrazumevane šifre su:

- `1039` — ostala prerada i konzervisanje voća i povrća;
- `4631` — trgovina na veliko voćem i povrćem.

Druge šifre se mogu proslediti:

```bash
python run_pipeline.py --mode full --codes 1039 4631 4621
```

### 2. Offline — postojeći Excel bez APR API-ja

Za postojeći 2024 fajl:

```bash
python run_pipeline.py --mode offline --input "../Clean Data/apr_veleprodaja_voca_povrca_prerada_konzervisanje_2024.xlsx"
```

Offline režim radi samo:

1. čišćenje i validaciju;
2. tržišnu analizu;
3. generisanje `Processed` i `Reports` izlaza.

Ne zahteva internet i ne poziva APR API.

### 3. Report — ponovna analiza validiranog skupa

```bash
python run_pipeline.py --mode report
```

Koristi najnoviji `apr_market_validated_*.xlsx` iz `Processed/` i ponovo generiše tržišne izveštaje.

## TLS politika

APR endpoint na pojedinim Windows/Python instalacijama može izazvati `CERTIFICATE_VERIFY_FAILED`.

Podrazumevana politika pipeline-a je:

1. prvo pokušaj sa punom TLS verifikacijom;
2. ako samo `openapi.apr.gov.rs` vrati SSL validation grešku, automatski ponovi zahtev bez verifikacije;
3. ispiši jasno upozorenje;
4. zapiši `tls_fallback_used`, originalnu grešku i stvarni TLS režim u metadata JSON.

Fallback nije dozvoljen za druge hostove.

Za strogi režim koji mora pasti ako sertifikat nije validan:

```bash
python run_pipeline.py --strict-tls
```

Za dijagnostiku i eksplicitno pokretanje bez prvog secure pokušaja:

```bash
python run_pipeline.py --insecure
```

`--insecure` i `--strict-tls` se međusobno isključuju.

## Status klasifikacija i zaštita od lažnog izveštaja

`03_clean_validate.py` prvo proverava negativna stanja, a zatim aktivna:

1. stečaj;
2. likvidacija;
3. neaktivan, brisan, ugašen ili prestao;
4. aktivan ili registrovan;
5. ostalo / nepoznato.

Validirani workbook sadrži posebne sheet-ove:

- `Status Values` — svaka originalna APR status vrednost i broj firmi;
- `Status Categories` — zbir po normalizovanoj kategoriji;
- `Quality Issues` — uključuje i neklasifikovane statuse.

Ako dataset ima redove, a klasifikacija vrati nula aktivnih firmi, validacija se prekida nakon što sačuva dijagnostičke sheet-ove. `04_analyze_market.py` takođe odbija da generiše izveštaj iz praznog scope-a. Time pipeline više ne može tiho da proizvede validno izgledajući izveštaj sa svim pokazateljima jednakim nuli.

## Pojedinačni koraci

```bash
python 01_extract_companies.py
python 02_enrich_financials.py
python 03_clean_validate.py
python 04_analyze_market.py
```

Svaki korak automatski bira najnoviji očekivani ulaz. Eksplicitni input može se zadati sa `--input` tamo gde je podržan.

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
- statusnu raspodelu;
- prometne razrede;
- geografsku raspodelu;
- raspodelu po delatnosti;
- top 200 firmi;
- koncentraciju prihoda;
- kompletan analizirani skup.

## Preporučeni operativni ritam

- `full` režim: kada se pravi novi godišnji APR snapshot;
- `offline` režim: za rad na postojećem Excelu i iteracije analitičkih pravila;
- `report` režim: kada se menjaju samo agregacije ili format izveštaja.

Time se APR API ne poziva pri svakoj analizi.

## Važna ograničenja

- Šifre `1039` i `4631` ne obuhvataju nužno sve organizovane otkupljivače.
- Prihod nije direktna mera broja stanica, kooperanata, logistike ili GGAP potrebe.
- Financial-statements endpoint se trenutno tretira kao jedan zapis po firmi. Višegodišnja istorija nije potvrđena ovim pipeline-om.
- Svaka nova ili promenjena APR status vrednost mora biti pregledana u sheet-u `Status Values`.
- Rezultati su ulaz za TAM/SAM/SOM model, ali nisu sami po sebi konačan TAM, SAM ili SOM.
- TLS fallback omogućava dostupnost APR podataka, ali metadata mora ostati sačuvana kao dokaz stvarnog režima preuzimanja.

## Legacy skripte

`AllCompanies.py` i `FindAllFin.py` su zadržane kao kompatibilni wrapper-i. Novi razvoj treba da koristi numerisane skripte.

## Metadata pravilo

Uz svaki generisani dataset čuvaju se:

- vreme generisanja;
- izvorni URL;
- obuhvaćene šifre;
- broj redova;
- TLS režim i eventualni fallback;
- statusna pravila i najčešće stvarne status vrednosti;
- poznata ograničenja;
- SHA-256 ulaznog i/ili izlaznog fajla.

## Privatnost

Pre javnog postavljanja proveriti da dataset ne sadrži osetljive lične ili ugovorne podatke. Matični broj pravnog lica koristi se za deduplikaciju i spajanje APR podataka.
