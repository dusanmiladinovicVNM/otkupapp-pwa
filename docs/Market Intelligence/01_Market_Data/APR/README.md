# APR tržišni podaci

Ovaj direktorijum sadrži APR izvore i reproduktivni pipeline za procenu AgriX tržišta.

## Gde ide postojeći Excel

Originalni Excel fajl sa firmama i prometom po godinama postaviti u:

`Raw Data/`

Primer naziva:

`apr_veleprodaja_voca_povrca_YYYY-MM-DD.xlsx`

## Direktorijumi

- `Raw Data/` — originalni APR Excel i drugi izvorni fajlovi; ne menjati ručno.
- `Clean Data/` — normalizovani, deduplikovani i dokumentovani podaci.
- `Processed/` — segmentacija, top firme, prometni razredi, geografija i drugi izlazi.
- `Reports/` — TAM, SAM, SOM, zaključci i grafikoni za Master Plan.
- `Scripts/` — Python skripte za extraction, cleaning, enrichment i analizu.

## Obavezni metadata podaci

Uz svaki novi dataset dokumentovati:

- datum preuzimanja;
- APR izvor i način ekstrakcije;
- obuhvaćene šifre delatnosti;
- obuhvaćene godine;
- opis kolona;
- pravila deduplikacije;
- poznata ograničenja;
- verziju skripte ili commit koji ga je proizveo.

## Privatnost

Pre postavljanja proveriti da Excel ne sadrži osetljive lične podatke ili podatke koji ne treba da budu u javnom tehničkom repozitorijumu. Matični broj firme može biti koristan za deduplikaciju, ali odluku o njegovom javnom čuvanju treba doneti eksplicitno.
