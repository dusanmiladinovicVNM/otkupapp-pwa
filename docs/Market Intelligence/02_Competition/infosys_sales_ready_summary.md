# Infosys replacement — sales-ready target summary

**Status:** Evidence-based account prioritization  
**Input:** `infosys_wide_enrichment.csv`  
**Važno:** tačan APR identitet nije isto što i potvrđen AgriX process fit.

## Sažetak

| Pokazatelj | Broj |
|---|---:|
| Jedinstveni matched pravni subjekti | 30 |
| Spremni za account research | 9 |
| Prvo potvrditi proces | 2 |
| Adjacent discovery | 7 |
| Hold — slab ili konfliktan process evidence | 11 |
| Isključeni zbog neaktivnog statusa | 1 |
| Identiteti za ručnu proveru | 10 |
| Bez bezbednog identiteta | 8 |

## Kvalifikovani research queue

Ova tabela trenutno sadrži 11 računa: devet algoritamski spremnih za account research i dva kojima prvo treba potvrditi proces. Broj nije veštački popunjen do 20 firmama iz `adjacent` ili `hold` grupa.

| # | Firma | MB | Prihod RSD | Zaposleni | Status prioriteta | Score |
|---:|---|---|---:|---:|---|---:|
| 1 | DRUŠTVO SA OGRANIČENOM ODGOVORNOŠĆU ZA PROIZVODNJU, PROMET I USLUGE BUDIM GRAD, BUDILOVINA | 06994857 | 1.398.804.000 | 59 | ready_for_account_research | 89 |
| 2 | POLJOPRIVREDNO PREDUZEĆE FRIGO-PAUN  DRUŠTVO SA OGRANIČENOM ODGOVORNOŠĆU POŽEGA | 06353096 | 881.724.000 | 31 | ready_for_account_research | 89 |
| 3 | FRUCOM FOOD društvo sa ograničenom odgovornošću Arilje | 20991348 | 775.708.000 | 66 | ready_for_account_research | 89 |
| 4 | AGRONOM FIT DOO POŽEGA | 20804734 | 462.733.000 | 9 | ready_for_account_research | 89 |
| 5 | ФРИГО БРАЋА МИТРОВИЋ доо ПИЛИЦА | 20932457 | 452.388.000 | 19 | ready_for_account_research | 89 |
| 6 | MAGIC BERRY FRUITS D.O.O. PILATOVIĆI | 22022989 | 211.140.000 | 2 | ready_for_account_research | 85 |
| 7 | DRUŠTVO ZA PROIZVODNJU I PROMET FRIGOMIL, DRUŠTVO SA OGRANIČENOM ODGOVORNOŠĆU KOTRAŽA | 17394037 | 83.422.000 | 8 | ready_for_account_research | 83 |
| 8 | AС-АГРО 99 д.о.о. Банатско Ново Село | 21225452 | 3.528.730.000 | 18 | ready_for_account_research | 78 |
| 9 | PREDUZEĆE ZA PROIZVODNJU, PRERADU I USLUGE U POLJOPRIVREDI MALINA PROIZVOD DOO, KOSJERIĆ (VAROŠ) | 20141948 | 18.765.000 | 1 | ready_for_account_research | 75 |
| 10 | Privredno društvo MASTER FRIGO d.o.o. Zlatibor | 06285783 | 382.735.000 | 2 | process_validation_first | 68 |
| 11 | MALUS JABUKA DRUŠTVO SA OGRANIČENOM ODGOVORNOŠĆU ZA PROIZVODNJU TRGOVINU I USLUGE, UDOVICE | 17363921 | 7.315.000 | 3 | process_validation_first | 49 |

## Pravila korišćenja

1. `ready_for_account_research` znači da su identitet, aktivan status i procesni signal dovoljno jaki za dublji research — ne za automatski hladni kontakt.
2. `process_validation_first` zahteva potvrdu broja stanica, kooperanata, dokumenata i stvarnog otkupnog toka.
3. `adjacent_discovery_only` ne ulazi u Enterprise outbound bez posebnog product discovery-ja.
4. `hold_until_process_proven` čuva konkurentski dokaz, ali ga ne pretvara u prodajni target.
5. Reference ne dokazuju da firma i dalje koristi Infosys.
6. Spoljni account research može promeniti algoritamski status; takve korekcije vode se u `infosys_account_research.csv` i ne brišu izvorni pipeline rezultat.
