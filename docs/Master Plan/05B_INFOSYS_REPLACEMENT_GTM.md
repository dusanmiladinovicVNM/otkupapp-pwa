# 05B — Infosys replacement GTM qualification

**Status:** Evidence-based qualification layer  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-23  
**Primarni izvori:** `infosys_wide_enrichment.csv`, `infosys_wide_enrichment.metadata.json`, `infosys_account_research.csv`

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
- referentna kategorija: poljoprivreda i kooperative;
- javno je eksplicitno navedeno da se bavi otkupom, skladištenjem i prodajom žitarica i uljarica.

`INFERENCE`: AS-AGRO 99 je snažan Tier A research kandidat za ratarstvo/kooperaciju. Veliki prihod nije dovoljan sam po sebi, ali javni procesni dokaz opravdava ozbiljan discovery.

### BUDIM GRAD

- MB `06994857`;
- šifra `1039`;
- prihod 2025 približno `1.398.804.000 RSD`;
- 59 zaposlenih;
- founder-confirmed hladnjača i Infosys referenca;
- javni izvori potvrđuju otkup, preradu i zamrzavanje voća.

### FRIGO-PAUN

Zvanični sajt potvrđuje otkup, preradu, pakovanje i izvoz voća, kao i ugovore sa više od 1.000 proizvođača. To je trenutno najjači javni signal velike mreže kooperanata u potvrđenom Infosys replacement skupu.

### FRUCOM FOOD

Javni izvori potvrđuju više lokacija, veliki skladišni i dnevni zamrzivački kapacitet i formalno praćenje proizvođača radi sledljivosti. Račun zahteva Enterprise account mapu, uključujući lokalni operativni tim i vlasničku grupu.

### AGRO-SUNCOKRET

- MB `20724064`;
- prihod 2025 približno `813.105.000 RSD`;
- 18 zaposlenih;
- aktuelni APR-derived i poslovni izvori vode firmu pod šifrom `1039`;
- javno je navedeno deset poslovnih jedinica;
- stariji javni profil navodi šifru `4621`, pa postoji konflikt istorijskih podataka.

`DECISION`: AGRO-SUNCOKRET ulazi u prvi research talas zbog veličine, aktuelnog procesnog signala i broja poslovnih jedinica, ali pre outreach-a mora se potvrditi stvarni otkupni tok, namena lokacija i aktuelna Infosys instalacija.

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

### AGRONOM FIT — korekcija algoritamskog statusa

Pipeline ga je označio kao `ready_for_account_research` zbog kategorije i šifre `4621`. Spoljna provera njegovog sopstvenog sajta, međutim, potvrđuje pre svega:

- poljoprivredne apoteke;
- pesticide;
- seme i sadni materijal;
- mineralna i organska đubriva;
- navodnjavanje;
- maloprodaju i veleprodaju inputa.

Nije pronađen javni dokaz otkupa, kooperantske mreže ili prijema poljoprivrednih proizvoda.

`DECISION`: AGRONOM FIT se operativno vodi kao `process_validation_first`, bez outbound-a dok se ne potvrdi stvarni otkupni ili kooperantski use case.

`RULE`: šifra `4621` sama nije dovoljan dokaz otkupa. Može opisivati trgovinu semenom, hranom za životinje ili drugim agro inputima bez mreže proizvođača.

---

## 5. Sales-readiness klasifikacija

Pipeline `Scripts/build_infosys_sales_ready_targets.py` deduplikuje matched redove po matičnom broju i koristi sledeće statuse:

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

`OVERRIDE RULE`: spoljno potvrđen procesni dokaz ima prednost nad algoritamskim statusom, ali se izvorni pipeline rezultat ne briše. Korekcije se vode u `infosys_account_research.csv`.

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

Dodatni spoljno validirani sloj:

- `infosys_account_research.csv`;
- `infosys_account_research_summary.md`;
- `infosys_win_interview_template.md`;
- `infosys_migration_discovery_checklist.md`.

Excel sadrži:

- summary;
- kvalifikovani research queue do 20 računa — trenutno 11, bez veštačkog popunjavanja adjacent/hold firmama;
- sve matched račune;
- identity manual-review queue;
- unmatched queue.

---

## 8. Preporučeni account-research talasi

### Talas 1 — dubinski research

1. FRIGO-PAUN;
2. BUDIM GRAD;
3. FRUCOM FOOD;
4. AS-AGRO 99;
5. AGRO-SUNCOKRET;
6. FRIGO BRAĆA MITROVIĆ.

### Talas 2 — verifikacija i verovatno kraći ciklus

7. MAGIC BERRY FRUITS;
8. FRIGOMIL;
9. MALINA PROIZVOD.

### Procesna provera pre outreach-a

10. AGRONOM FIT.

Za svaki račun prikupiti:

- potvrdu da li još koristi Infosys;
- broj otkupnih mesta, silosa ili hladnjača;
- broj kooperanata/proizvođača;
- kulture i trajanje sezone;
- tipove dokumenata i sezonski volumen;
- postojeći ERP/knjigovodstvo;
- operativnog sponsora i decision-maker mapu;
- switching trigger;
- format i rizike migracije;
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

AgriX ne treba da kontaktira svih 30 firmi. Prvi operativni fokus je devet spolja validiranih account-research kandidata, uz AGRONOM FIT kao dodatnu procesnu proveru. Nakon toga sledi Infosys migration discovery i battlecard zasnovan na dva postojeća AgriX klijenta koji su već prešli sa Infosys-a.
