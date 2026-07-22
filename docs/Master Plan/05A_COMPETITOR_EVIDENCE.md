# 05A — Konkurentski dokazni signali: Infosys i Yuteam

**Status:** Review — evidence addendum  
**Vlasnik:** osnivač AgriX-a  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `04_MARKET.md`, `05_COMPETITION.md`, `14_GO_TO_MARKET.md`, `15_SALES_PLAYBOOK.md`  
**Primarni dokazni skup:** `docs/Market Intelligence/02_Competition/`

---

## 1. Svrha

Ovaj dodatak beleži konkurentske signale koji menjaju prioritete istraživanja:

1. Infosys kao potvrđen izvor dve AgriX migracije;
2. 114 agro i prehrambenih stavki izdvojenih ili naknadno potvrđenih u Infosys referentnom universe-u;
3. 49 visokopotencijalnih Infosys replacement računa;
4. Yuteam kao dobavljač sa 12 navedenih referenci u agraru, zadrugama, kooperaciji i mlinarstvu.

Dokument ne tvrdi tržišni udeo, aktuelnost svih instalacija niti funkcionalnu superiornost bilo kog sistema bez dodatne provere.

---

## 2. Infosys — najjači trenutni replacement signal

`FACT — founder-confirmed business event`: dva postojeća AgriX klijenta prethodno su koristila Infosys sistem.

`FACT`: oba računa imaju približno 150 miliona RSD godišnjih prihoda.

`FOUNDER ASSESSMENT`: Infosys je jedan od najvećih konkurenata u segmentu.

Ovaj dokaz je vredniji od obične javne reference jer potvrđuje:

- da je zamena incumbent sistema moguća;
- da je AgriX već pobedio u najmanje dve realne odluke;
- da korisnici imaju iskustvo dovoljno za direktno poređenje;
- da postoji replacement market, a ne samo greenfield tržište.

### 2.1 Reference universe

Osnovni klasifikovani Excel sadrži 113 agro i prehrambenih stavki. Osnivač je dodatno potvrdio da se u originalnoj Infosys listi nalazi i `DOO BUDIM GRAD BUDILOVINA`, koji je izostao iz filtriranog Excela.

Ažurirani universe:

| Pokazatelj | Broj |
|---|---:|
| Ukupno Infosys agro/prehrambenih referenci | 114 |
| Visok AgriX potencijal | 49 |
| Srednji potencijal | 49 |
| Nizak potencijal | 16 |
| Prioritet A: voće, kooperative i duvan | 40 |
| Prioritet B: mlekare | 9 |

`INFERENCE`: Infosys replacement pool je dovoljno veliki da bude zaseban GTM kanal, a ne samo ad hoc odgovor kada se javi nezadovoljan korisnik.

### 2.2 BUDIM GRAD

`FACT — founder-confirmed`: BUDIM GRAD se nalazi u Infosys referentnoj listi i jeste hladnjača.

Javno potvrđeni podaci:

- pravni naziv: `DOO BUDIM GRAD BUDILOVINA`;
- MB: `06994857`;
- PIB: `101140298`;
- lokacija: Budilovina, Brus;
- šifra delatnosti: `1039`;
- status: aktivan;
- prihod 2025: približno `1.399.742.000 RSD`;
- broj zaposlenih 2025: `59`;
- poslovni tok: otkup svežeg voća, prerada i zamrzavanje.

`INFERENCE`: Infosys baza uključuje i značajne Tier A račune. BUDIM GRAD je višestruko veći od dosadašnjeg preliminarnog ICP revenue signala od oko 150–400 miliona RSD i potvrđuje da prihod ne treba koristiti kao gornju granicu ICP-a.

`DECISION`: BUDIM GRAD ulazi u prioritet A replacement listu.

### 2.3 Obavezni win interview

Za oba migrirana klijenta treba strukturisano dokumentovati:

1. koliko dugo su koristili Infosys;
2. koje procese su vodili u njemu;
3. šta je pokrenulo traženje alternative;
4. koje funkcionalne ili operativne praznine su osećali;
5. koliki je bio switching cost;
6. šta je presudilo u korist AgriX-a;
7. koje su prednosti AgriX-a potvrdili nakon korišćenja;
8. šta je u migraciji bilo teško;
9. u čemu je prethodni sistem bio bolji;
10. šta bi moglo izazvati povratak ili churn.

`DECISION`: Infosys dobija prioritet 1 u product teardown-u, win/loss analizi i pozicioniranju.

---

## 3. Reproduktivni Infosys replacement pipeline

Kreirani su:

- `infosys_agro_references.csv` — osnovnih 113 klasifikovanih stavki;
- `infosys_manual_reference_additions.csv` — founder-confirmed dodaci, trenutno BUDIM GRAD;
- `Scripts/build_infosys_replacement_targets.py` — spajanje sa APR podacima i generisanje replacement liste.

Pipeline:

1. spaja osnovne i ručno potvrđene reference;
2. podrazumevano bira visokopotencijalne račune;
3. koristi matični broj kada je poznat;
4. ostale firme poredi po normalizovanom nazivu i lokaciji;
5. odvaja `matched`, `manual_review` i `unmatched`;
6. pravi CSV i Excel izlaz;
7. ne prihvata slab fuzzy match kao činjenicu.

`DECISION`: target-account dataset mora čuvati match score, metod povezivanja i potrebu za ručnom proverom.

---

## 4. Yuteam — ratarstvo, zadruge i kooperacija

Dostavljena referentna lista pod naslovom `RUGE I KOOPERACIJE` sadrži:

1. DOO Raca — Zrenjanin;
2. DOO Banija Agrar — Pivnice;
3. DOO Velisavljev — Botoš;
4. DOO Romić Agrar — Kikinda;
5. Agropromet Keser DOO — Kikinda;
6. Nikša Agrar — Kljaićevo;
7. PP Borac AD — Šurjan;
8. ZZ Zrenjanin — Zrenjanin;
9. ZZ Yuko — Begejci;
10. ZZ Elemir — Eelemir, prema dostavljenom zapisu;
11. ZZ Žitarice — Kać;
12. Mlin Banatski klas — Bavanište.

Strukturni signali:

- 12 referenci u poljoprivrednom i kooperantskom prostoru;
- najmanje pet zemljoradničkih zadruga;
- više agrarnih privrednih društava;
- jedan mlin;
- dominantan geografski klaster Vojvodine;
- signal da relevantno tržište nije ograničeno na voće, hladnjače i APR šifre `1039`/`4631`.

`INFERENCE`: Yuteam je ozbiljan adjacent konkurent za buduće širenje AgriX-a prema ratarstvu, žitaricama, zadrugama i organizovanoj kooperaciji.

---

## 5. Ažurirana konkurentska hijerarhija

### Prioritet 1 — potvrđen replacement konkurent

**Infosys**

Razlog: dve stvarne migracije ka AgriX-u, 114 izdvojenih agro/prehrambenih referenci i 49 visokopotencijalnih replacement računa.

### Prioritet 2 — direktni specijalizovani konkurenti u voću i otkupu

**SOFTEK** i **KRUNET**

Razlog: zajedno najmanje 24 javno navedene direktno relevantne reference.

### Prioritet 3 — adjacent agrar, zadruge i kooperacija

**Yuteam**

Razlog: 12 referenci u Vojvodini, agraru, zadrugama i mlinarstvu.

### Prioritet 4 — horizontalne i projektne alternative

- generički ERP;
- lokalni programer;
- interni Excel/Access;
- status quo.

---

## 6. Tržišne implikacije

1. `FACT`: AgriX već ima najmanje dve replacement pobede protiv Infosys-a.
2. `INFERENCE`: prelazak sa konkurenta može biti važniji kanal od prodaje firmama bez softvera.
3. `INFERENCE`: BUDIM GRAD potvrđuje da replacement tržište uključuje i firme sa prihodima većim od milijardu dinara.
4. `INFERENCE`: trenutni APR scan ne obuhvata dovoljno ratarstvo, zadruge, mlekare i duvan.
5. `DECISION`: ICP ne sme eliminisati firmu zbog šifre delatnosti kada procesni signali potvrđuju otkup ili kooperaciju.
6. `DECISION`: prethodni sistem postaje obavezno CRM polje za svaki lead.
7. `DECISION`: svaki novi klijent mora imati dokumentovan switching trigger i razlog pobede.
8. `DECISION`: migracija sa Infosys-a postaje poseban sales i onboarding playbook.

---

## 7. Sledeći istraživački koraci

### Infosys

- pokrenuti APR matching pipeline;
- ručno proveriti fuzzy i unmatched rezultate;
- sprovesti dva win interview-a;
- popisati funkcije koje su klijenti koristili;
- utvrditi migracioni proces i format podataka;
- proveriti pricing, održavanje, support i deployment;
- napraviti battlecard zasnovan samo na dokazima.

### Yuteam

- spojiti 12 referenci sa APR podacima;
- izdvojiti zadruge, mlinove i agrarne DOO kao zasebne subsegmente;
- istražiti specifične tokove žitarica, skladištenja i kooperacije;
- proceniti koliko AgriX može ući u ovaj segment bez stvaranja forka proizvoda.

---

## 8. Zaključak

Infosys je trenutno najvažniji konkurent za razumevanje zato što AgriX protiv njega već ima dve stvarne pobede, a referentni universe pokazuje široku prisutnost u ciljnom i susednim sektorima. BUDIM GRAD potvrđuje da Infosys pokriva i značajne hladnjače, ne samo male ili srednje firme.

Sledeća konkurentska prednost neće nastati iz duže liste funkcija, već iz preciznog odgovora na tri pitanja:

1. zašto korisnik napušta postojeći sistem;
2. zašto bira AgriX;
3. kako AgriX migraciju i prvu sezonu čini manje rizičnim od ostanka ili izbora drugog dobavljača.
