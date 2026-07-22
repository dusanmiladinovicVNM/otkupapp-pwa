# 01 — Market Data

Centralno mesto za strukturisane tržišne podatke.

## Izvori

- `APR/` — registri firmi, finansijski podaci i rezultati skripti.
- `Statistics/` — RZS, Eurostat, FAOSTAT i druge statistike.
- `Government/` — ministarstva, uprave, registri i javni izvori.
- `Other Sources/` — ostali proverljivi izvori koji ne pripadaju prethodnim grupama.

## Pravilo obrade

Za svaki važan izvor koristiti slojeve:

- `Raw Data/` — originalni fajlovi, bez ručne izmene;
- `Clean Data/` — normalizovani i deduplikovani podaci;
- `Processed/` — analitički izlazi;
- `Reports/` — zaključci, tabele i grafikoni;
- `Scripts/` — kod za preuzimanje, čišćenje i analizu.
