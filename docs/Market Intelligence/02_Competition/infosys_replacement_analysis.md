# Infosys replacement analiza — prvi istraživački talas

**Status:** Initial evidence-based prioritization  
**Datum:** 2026-07-23  
**Izvori:** `infosys_replacement_targets.csv`, `infosys_first_wave_accounts.csv`  
**APR scope:** samo šifre delatnosti `1039` i `4631`

---

## 1. Izvršni sažetak

Infosys replacement pipeline trenutno obrađuje 49 visokopotencijalnih referentnih redova:

| Pokazatelj | Broj |
|---|---:|
| Visokopotencijalni referentni redovi | 49 |
| Prioritet A: voće, kooperative i duvan | 40 |
| Prioritet B: mlekare | 9 |
| Sigurno povezani redovi | 8 |
| Redovi za ručnu proveru | 5 |
| Nepovezani u uskom APR skupu | 36 |
| Sigurna stopa povezivanja | 16,3% |
| Pokrivenost uključujući manual review | 26,5% |

Osam sigurno povezanih redova predstavlja sedam jedinstvenih pravnih lica, jer `MLEKARA MORAVICA DOO ARILJE` i `MORAVICA DOO` iz Infosys izvora vode do istog MB `07635826`.

`LIMITATION`: 36 `unmatched` redova ne znači da firme ne postoje u APR-u niti da nisu aktivne. Znači samo da identitet nije bezbedno potvrđen u datasetu ograničenom na šifre `1039` i `4631`.

---

## 2. Potvrđeni jedinstveni domaći računi

| Prioritet | Referenca | MB | Prihod RSD | Zaposleni | Segment |
|---:|---|---|---:|---:|---|
| 1 | BUDIM GRAD | 06994857 | 1.398.804.000 | 59 | voće i hladnjače |
| 2 | FRIGO-PAUN | 06353096 | 881.724.000 | 31 | voće i hladnjače |
| 3 | AGRO-SUNCOKRET | 20724064 | 813.105.000 | 18 | poljoprivreda/kooperacija — proveriti proces |
| 4 | FRUCOM FOOD | 20991348 | 775.708.000 | 66 | voće i hladnjače |
| 5 | MORAVICA | 07635826 | 124.165.000 | 13 | mlekarstvo — adjacent discovery |
| 6 | FRIGOMIL | 17394037 | 83.422.000 | 8 | voće i hladnjače |
| 7 | MALINA PROIZVOD | 20141948 | 18.765.000 | 1 | voće/malina |

Sedam jedinstvenih potvrđenih pravnih lica zajedno imaju približno:

- `4.095.693.000 RSD` prihoda;
- `196` zaposlenih.

Ovi zbirni pokazatelji opisuju identifikovane firme, ne tržišni prihod dostupan AgriX-u.

---

## 3. Šta rezultat dokazuje

1. Infosys reference obuhvataju i veoma velike račune: četiri potvrđene firme imaju više od 775 miliona RSD godišnjih prihoda.
2. Replacement tržište nije ograničeno na preliminarni ICP signal od 100–500 miliona RSD.
3. Prihod nije dovoljan za rangiranje: `MALINA PROIZVOD` ima mali finansijski proxy, ali veoma direktan AgriX process fit.
4. Zapadna i Centralna Srbija ostaju glavni početni klasteri: Brus, Požega, Užice, Arilje, Kotraža i Kosjerić.
5. Mlekarstvo treba voditi kao adjacent discovery, ne kao automatski Enterprise fit.
6. Infosys replacement GTM mora imati poseban migration playbook, jer AgriX već ima dve realne migracije sa tog sistema.

---

## 4. Manual review grupa

### MASTER FRUITS Beograd

APR kandidat ima:

- MB `21202266`;
- prihod 2024: `3.286.002.000 RSD`;
- 117 zaposlenih;
- šifru `1039`.

Naziv je tačan, ali izvorna lokacija `Milićevo Selo` nije usklađena sa registrovanim sedištem Beograd–Palilula. Ne promovirati u potvrđen target pre provere da li je Milićevo Selo operativna lokacija iste firme.

### MAGIC BERRY FRUITS

Verovatni kandidat:

- MB `22022989`;
- prihod 2025: `211.140.000 RSD`;
- lokacija/opština Požega.

Dobar kandidat za ručnu potvrdu pravnog identiteta.

### PROBAR FRUIT

Verovatni kandidat:

- MB `21185477`;
- prihod 2025: `31.608.000 RSD`;
- 2 zaposlena.

Direktan sektorski fit, ali komercijalni prioritet zavisi od broja stanica, kooperanata i dokumentacionog obima.

### RIVAMIL

Verovatni kandidat:

- MB `17299158`;
- prihod 2025: `13.996.000 RSD`;
- 3 zaposlena.

Direktan voćarski fit, ali zahteva proveru da li veličina operativnog toka opravdava Enterprise prodajni napor.

### MASTER FRUITS Srebrenica

Ovo je regionalni BiH signal i ne sme se spajati sa srpskim MB `21202266` samo na osnovu identičnog naziva. Vodi se kao zaseban regionalni account.

---

## 5. Prvi istraživački talas

### A1 — strateški računi

1. BUDIM GRAD;
2. FRIGO-PAUN;
3. FRUCOM FOOD;
4. AGRO-SUNCOKRET.

Kriterijumi:

- potvrđen identitet;
- značajan prihod;
- direktan ili verovatan otkupni/kooperantski tok;
- lokacija u relevantnom klasteru;
- Infosys reference evidence.

### A2 — direktan fit, manji ekonomski proxy

5. FRIGOMIL;
6. MALINA PROIZVOD.

Ove firme ne treba automatski spustiti samo zbog manjeg prihoda. Broj stanica, kooperanata i sezonski dokumentacioni obim mogu ih učiniti boljim AgriX fitom od veće, ali jednostavnije firme.

### B — adjacent discovery

7. MORAVICA.

Mlekarstvo zahteva zasebnu proveru procesa i ne sme izazvati fork proizvoda pre dokaza o ponovljivom tržištu.

### Verify before ranking

8. MASTER FRUITS Beograd;
9. MAGIC BERRY FRUITS;
10. PROBAR FRUIT;
11. RIVAMIL.

---

## 6. Obavezni account-research podaci

Za svaki račun prvog talasa dopuniti:

- da li i dalje koristi Infosys;
- naziv i verziju Infosys modula;
- broj otkupnih stanica/lokacija;
- broj kooperanata;
- glavne kulture i sezonalnost;
- godišnji broj otkupnih dokumenata;
- postojeći računovodstveni ERP;
- izvozne i GGAP zahteve;
- switching trigger;
- osobu koja poseduje problem;
- implementacioni rok pre sledeće sezone;
- dozvoljen i primeren kontaktni kanal.

`DECISION`: referentna lista konkurenta nije sama po sebi razlog za agresivan outbound. Prvo se potvrđuju identitet, aktuelni sistem i procesni fit.

---

## 7. Sledeći tehnički korak

Za 36 `unmatched` redova treba izgraditi širi legal-entity enrichment koji nije ograničen na šifre `1039` i `4631`:

1. pronaći pravni identitet po nazivu i lokaciji;
2. potvrditi matični broj;
3. povući aktivan status, šifru delatnosti, prihod i zaposlene;
4. odvojiti Srbiju od regionalnih referenci;
5. deduplikovati po matičnom broju;
6. tek zatim računati komercijalni prioritet.

Do tada se `candidate_*` polja koriste samo kao dijagnostika, nikada kao potvrđen identitet.
