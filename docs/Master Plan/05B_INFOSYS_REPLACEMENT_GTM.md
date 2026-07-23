# 05B — Infosys replacement GTM qualification

**Status:** Evidence-based qualification layer  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-23  
**Primarni izvori:** `infosys_wide_enrichment.csv`, `infosys_wide_enrichment.metadata.json`

---

## 1. Zašto je potreban drugi sloj

Wide APR enrichment je uspešno proširio identity matching sa dve šifre delatnosti na kompletan APR registar.

Rezultat:

| Pokazatelj | Broj |
|---|---:|
| Visokopotencijalni referentni redovi | 49 |
| Matched redovi | 31 |
| Manual review | 10 |
| Unmatched | 8 |
| Jedinstveni matched matični brojevi | 30 |
| APR registry records pregledano | 133.802 |

Ovaj rezultat je identity evidence, ne automatska prodajna lista.

`CRITICAL DISTINCTION`:

- tačan pravni identitet potvrđuje koja firma je navedena u referentnoj listi;
- APR šifra i javni podaci pomažu da se proceni današnja delatnost;
- ni jedno ni drugo samo po sebi ne potvrđuje da firma danas koristi Infosys;
- ni jedno ni drugo samo po sebi ne potvrđuje otkupni, kooperantski ili hladnjačarski proces;
- prodajni target nastaje tek kada su identitet, aktivan status i process fit dovoljno jaki.

---

## 2. Šta wide enrichment dokazuje

1. Infosys reference obuhvataju firme van šifara `1039` i `4631`.
2. Replacement pool uključuje voće, ratarstvo, kooperaciju, skladištenje, mlekare, duvan i druge susedne tokove.
3. APR šifra ne sme biti jedini eliminacioni kriterijum, ali mora biti signal za proveru.
4. Širi registar je podigao identity coverage sa 8 matched redova na 31.
5. Preostali problem više nije nalaženje pravnog lica, već dokazivanje aktuelnog sistema i stvarnog AgriX process fit-a.

---

## 3. Novi veliki i strateški računi

### AS-AGRO 99

- MB `21225452`;
- aktivan status;
- šifra `4621`;
- prihod 2025 približno `3.528.730.000 RSD`;
- 18 zaposlenih;
- referentna kategorija: poljoprivreda i kooperative.

`INFERENCE`: AS-AGRO 99 je snažan Tier A research kandidat za ratarstvo/kooperaciju. Veliki prihod ne dokazuje broj stanica ili kooperanata, ali opravdava ozbiljan account research.

### BUDIM GRAD

- MB `06994857`;
- šifra `1039`;
- prihod 2025 približno `1.398.804.000 RSD`;
- 59 zaposlenih;
- founder-confirmed hladnjača i Infosys referenca.

### FRIGO-PAUN, AGRO-SUNCOKRET i FRUCOM FOOD

Ostaju potvrđeni veliki domaći replacement računi sa direktnim ili vrlo jakim procesnim signalom.

### SIROGOJNO COMPANY

- potencijalni APR kandidat ima prihod približno `7.725.769.000 RSD` i 202 zaposlena;
- identity status je `manual_review` zbog razlike između izvornog zapisa `SIROGOJNO CO DOO — Rupeljevo` i APR pravnog identiteta.

`DECISION`: SIROGOJNO ne ulazi u prodajni talas pre ručne potvrde identiteta i aktuelnog Infosys odnosa.

---

## 4. Zašto matched nije isto što i sales-ready

Wide rezultat sadrži tačne ili veoma verovatne identitete čija APR delatnost ne potvrđuje cilj AgriX-a:

| Referenca | APR šifra | Signal |
|---|---:|---|
| MASTER FRIGO | 2825 | naziv ukazuje na `frigo`, ali registrovana delatnost ne potvrđuje otkup/hladnjaču |
| DRIM | 2550 | metaloprerađivačka delatnost — procesni konflikt |
| FILOS | 2849 | proizvodnja mašina/alata — procesni konflikt |
| PARAMUN | 4730 | trgovina motornim gorivima — procesni konflikt |
| LADJEVAC | 4941 | drumski transport — procesni konflikt |
| LATONA | 7911 | turistička agencija — procesni konflikt |
| ZLATIBORSKI EKO AGRAR | 7022 | naziv je agrarni, ali delatnost traži procesnu potvrdu |
| ALLIANCE ONE TOBACCO | 4621 | identitet je relevantan, ali firma je u likvidaciji |

Ove stavke moraju ostati u evidence datasetu, ali ne smeju automatski ući u outbound ili top target listu.

---

## 5. Sales-readiness klasifikacija

Novi pipeline `Scripts/build_infosys_sales_ready_targets.py` deduplikuje matched redove po matičnom broju i koristi sledeće statuse:

### `ready_for_account_research`

Aktivan pravni subjekt sa potvrđenim identitetom i direktnim procesnim signalom. Ovo znači da je opravdan dublji research, ne automatski kontakt.

### `process_validation_first`

Identitet i referenca su jaki, ali šifra ili javni podaci ne dokazuju broj stanica, kooperanata, dokumentacioni tok ili otkup.

### `adjacent_discovery_only`

Mlekarstvo i drugi susedni segmenti koji mogu biti budući proizvodni pravac, ali nisu automatski Enterprise outbound.

### `hold_until_process_proven`

Identitet je potvrđen, ali delatnost ili klasifikacija trenutno konfliktuju sa AgriX procesima.

### `exclude`

Neaktivan subjekt, likvidacija, stečaj ili drugo negativno stanje.

---

## 6. Ranking princip

Target score ne sme biti prost prihod. Kombinuje:

1. process-fit status;
2. aktivan pravni status;
3. sigurnost identiteta;
4. prihod kao proxy sposobnosti plaćanja i složenosti;
5. zaposlene kao slab dodatni proxy;
6. početni geografski klaster AgriX-a;
7. potrebu za ručnom potvrdom.

`DECISION`: process fit ima veću težinu od prihoda. Mala hladnjača sa više otkupnih stanica može biti bolji target od mnogo veće firme čija delatnost nema otkupni tok.

---

## 7. Operativni izlazi

Pokretanje:

```bash
python "docs/Market Intelligence/02_Competition/Scripts/build_infosys_sales_ready_targets.py"
```

Generiše:

- `infosys_sales_ready_targets.csv`;
- `infosys_sales_ready_targets.xlsx`;
- `infosys_sales_ready_summary.md`.

Excel sadrži:

- summary;
- deduplikovani Top 20;
- sve matched račune;
- identity manual-review queue;
- unmatched queue.

---

## 8. Sledeći poslovni korak

Nakon generisanja sales-ready izlaza, za prvih 10–20 računa treba prikupiti:

- potvrdu da li još koriste Infosys;
- broj otkupnih mesta ili lokacija;
- broj kooperanata;
- kulture/proizvode;
- tipove dokumenata i sezonski volumen;
- postojeći ERP/knjigovodstvo;
- odgovornu osobu i decision-maker-a;
- switching trigger;
- mogući izvor preporuke;
- dozvoljeni sledeći kontakt.

Tek nakon toga account prelazi iz `research` u stvarni CRM prospect.

---

## 9. Zaključak

Infosys replacement tržište je potvrđeno kao poseban GTM kanal. Wide APR enrichment je identifikovao 30 jedinstvenih pravnih lica, ali tržišna vrednost nije u broju match-eva. Vrednost je u pravilnom odvajanju:

1. potvrđenog identiteta;
2. aktivnog pravnog statusa;
3. stvarnog otkupnog ili kooperantskog procesa;
4. aktuelnog Infosys odnosa;
5. realnog switching trigger-a.

AgriX ne treba da kontaktira svih 30 firmi. Treba prvo detaljno istražiti mali broj računa kod kojih su sva četiri signala najjača: process fit, veličina, geografska blizina i replacement verovatnoća.
