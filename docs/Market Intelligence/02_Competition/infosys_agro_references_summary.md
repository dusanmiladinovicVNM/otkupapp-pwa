# Infosys — agro i prehrambeni reference universe

**Status:** Initial classified evidence set  
**Datum obrade:** 2026-07-23  
**Izvori:** `infosys_agro_firme(1).xlsx` i founder-confirmed dodatak BUDIM GRAD  
**Strukturisani podaci:** `infosys_agro_references.csv`, `infosys_manual_reference_additions.csv`

---

## 1. Šta dataset predstavlja

Osnovni Excel sadrži 113 stavki iz šire Infosys referentne liste koje su izdvojene kao agro, poljoprivredne ili prehrambene firme i organizacije. Osnivač je naknadno potvrdio da se u originalnoj Infosys referentnoj listi nalazi i `DOO BUDIM GRAD BUDILOVINA`, koji nije bio obuhvaćen filtriranim Excelom.

Ukupan evidence universe zato sada sadrži **114 stavki**.

Klasifikacija je zasnovana na:

- nazivu firme ili organizacije;
- lokaciji;
- prepoznatljivom sektoru;
- ručno dodeljenom AgriX potencijalu;
- founder-confirmed poslovnom kontekstu;
- javnim identifikacionim i poslovnim dokazima za BUDIM GRAD.

`LIMITATION`: lista za većinu stavki ne sadrži matični broj, šifru delatnosti, datum implementacije, naziv Infosys modula niti potvrdu da firma danas koristi Infosys. Stavka u listi nije automatski aktivna instalacija ili direktan AgriX target.

---

## 2. Osnovni pokazatelji

| Pokazatelj | Broj |
|---|---:|
| Ukupno izdvojenih i potvrđenih stavki | 114 |
| Jasno ili naknadno potvrđeno agro/prehrambeno | 90 |
| Potrebna dodatna provera | 24 |
| Visok AgriX potencijal | 49 |
| Srednji AgriX potencijal | 49 |
| Nizak AgriX potencijal | 16 |
| Različite navedene lokacije | 48 |

---

## 3. Raspodela po kategorijama

| Kategorija | Broj |
|---|---:|
| Voće, povrće i hladnjače | 26 |
| Žitarice, pekarstvo i konditorska industrija | 24 |
| Poljoprivreda i kooperative | 13 |
| Opšta prehrambena industrija | 11 |
| Mlekarstvo | 9 |
| Vino, pivo i ostala pića | 9 |
| Agro-inputi i mehanizacija | 7 |
| Stočarstvo, ribarstvo i veterina | 7 |
| Meso i prerada mesa | 5 |
| Instituti i strukovne organizacije | 2 |
| Duvan | 1 |

---

## 4. Visok AgriX potencijal

Svih 49 stavki označenih kao visok potencijal pripadaju sledećim grupama:

| Kategorija | Broj | Prioritet |
|---|---:|---|
| Voće, povrće i hladnjače | 26 | A |
| Poljoprivreda i kooperative | 13 | A |
| Duvan | 1 | A |
| Mlekarstvo | 9 | B — tek nakon product-fit provere |
| **Ukupno** | **49** | |

`INFERENCE`: Infosys ima dokazanu ili istorijsku prisutnost upravo u sektorima koji su najbliži AgriX Enterprise tržištu. Replacement segment je značajno širi od dva već migrirana klijenta.

`DECISION`: ovih 49 stavki čine seed listu za Infosys replacement-market istraživanje, ali se ne targetiraju pre potvrde pravnog identiteta, aktivnog statusa, procesa i aktuelnog sistema.

---

## 5. BUDIM GRAD — potvrđeni Tier A račun

`FACT — founder-confirmed`: BUDIM GRAD se nalazi u Infosys referentnoj listi i jeste hladnjača.

Javno potvrđeni identifikacioni podaci:

| Polje | Vrednost |
|---|---|
| Pravni naziv | DOO BUDIM GRAD BUDILOVINA |
| Matični broj | 06994857 |
| PIB | 101140298 |
| Lokacija | Budilovina, Brus |
| Šifra delatnosti | 1039 |
| Status | Aktivan |
| Delatnost | ostala prerada i konzervisanje voća i povrća |
| Prihod 2025. | 1.399.742.000 RSD |
| Zaposleni 2025. | 59 |

Firma se javno opisuje kao porodična kompanija koja se bavi otkupom svežeg voća i preradom u zamrznuto voće, sa značajnim izvozom.

`DECISION`: BUDIM GRAD ulazi u prioritet A replacement listu i predstavlja dokaz da Infosys baza ne obuhvata samo male ili srednje kupce, već i firmu sa prihodima većim od milijardu dinara.

---

## 6. Geografski signali

Najzastupljenije lokacije u osnovnom Excelu:

| Lokacija | Broj stavki |
|---|---:|
| Čačak | 19 |
| Požega | 12 |
| Zlatibor | 7 |
| Arilje | 6 |
| Beograd | 6 |
| Užice | 5 |
| Kragujevac | 4 |
| Novi Beograd | 3 |
| Čajetina | 3 |

BUDIM GRAD dodaje Brus/Budilovinu kao novi potvrđeni visokopotencijalni klaster signal.

Visokopotencijalne reference su posebno vidljive u Požegi, Arilju, Čačku, Zlatiboru, Bajinoj Bašti, Kosjeriću, Kotraži, Užicu i okolini, a sada i u Brusu.

`INFERENCE`: Infosys ima jaku referentnu gustinu u Zapadnoj i Centralnoj Srbiji, što se preklapa sa početnim AgriX tržišnim fokusom i povećava verovatnoću replacement prodaje.

Dataset sadrži i regionalne reference u Pljevljima, Srebrenici i Zvorniku.

---

## 7. Reproduktivni replacement pipeline

Skripta:

`Scripts/build_infosys_replacement_targets.py`

radi sledeće:

1. spaja osnovnih 113 klasifikovanih stavki sa founder-confirmed dodacima;
2. podrazumevano bira 49 visokopotencijalnih računa;
3. koristi poznati matični broj kada postoji;
4. ostale firme poredi sa validiranim APR skupom po normalizovanom nazivu i lokaciji;
5. rezultate deli na `matched`, `manual_review` i `unmatched`;
6. generiše CSV i Excel replacement listu;
7. ne prihvata slab fuzzy match kao činjenicu.

BUDIM GRAD koristi poznati MB `06994857`, pa ne zavisi od fuzzy podudaranja naziva.

---

## 8. Najvažniji strateški zaključci

1. Infosys nije samo jedan od konkurenata; njegova lista pokazuje široku agro/prehrambenu bazu.
2. Dva stvarna prelaska na AgriX potvrđuju da je deo te baze osvojiv.
3. Najvredniji replacement pool trenutno je 49 visokopotencijalnih računa.
4. BUDIM GRAD potvrđuje da replacement pool uključuje i srednje preduzeće sa oko 1,4 milijarde RSD prihoda.
5. APR šifre `1039` i `4631` propuštaju mlekare, duvan, kooperative i druge procese koji se pojavljuju u Infosys bazi.
6. Regionalna prodaja treba da počne od klastera gde postoje reference, procesna sličnost i mogućnost preporuke.
7. AgriX treba da dokumentuje migracioni alat i playbook za prelazak sa Infosys-a, jer je migracija već potvrđen prodajni use case.

---

## 9. Obavezna validacija pre targetiranja

Za svaku od 49 visokopotencijalnih stavki dopuniti:

- tačan pravni naziv;
- matični broj;
- aktivan status;
- prihod i broj zaposlenih;
- šifru delatnosti;
- broj stanica/lokacija;
- broj kooperanata;
- kulture i proizvode;
- modul ili proces koji Infosys pokriva;
- da li se sistem i dalje koristi;
- približnu godinu implementacije;
- switching trigger;
- ownera i championa;
- dozvoljeni sledeći kontakt.
