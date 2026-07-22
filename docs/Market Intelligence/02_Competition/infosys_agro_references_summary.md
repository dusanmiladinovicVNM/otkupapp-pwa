# Infosys — agro i prehrambeni reference universe

**Status:** Initial classified evidence set  
**Datum obrade:** 2026-07-23  
**Izvor:** `infosys_agro_firme(1).xlsx`, izveden iz Infosys referentne liste  
**Strukturisani podaci:** `infosys_agro_references.csv`

---

## 1. Šta dataset predstavlja

Excel sadrži 113 stavki iz šire Infosys referentne liste koje su izdvojene kao agro, poljoprivredne ili prehrambene firme i organizacije.

Klasifikacija je zasnovana na:

- nazivu firme ili organizacije;
- lokaciji;
- prepoznatljivom sektoru;
- ručno dodeljenom AgriX potencijalu.

`LIMITATION`: lista ne sadrži matični broj, šifru delatnosti, datum implementacije, naziv modula niti potvrdu da firma danas koristi Infosys. Stavka u listi nije automatski aktivna instalacija ili direktan AgriX target.

---

## 2. Osnovni pokazatelji

| Pokazatelj | Broj |
|---|---:|
| Ukupno izdvojenih stavki | 113 |
| Jasno agro/prehrambeno iz naziva | 89 |
| Potrebna dodatna provera | 24 |
| Visok AgriX potencijal | 48 |
| Srednji AgriX potencijal | 49 |
| Nizak AgriX potencijal | 16 |
| Različite navedene lokacije | 47 |

---

## 3. Raspodela po kategorijama

| Kategorija | Broj |
|---|---:|
| Voće, povrće i hladnjače | 25 |
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

Svih 48 stavki označenih kao visok potencijal pripadaju sledećim grupama:

| Kategorija | Broj |
|---|---:|
| Voće, povrće i hladnjače | 25 |
| Poljoprivreda i kooperative | 13 |
| Mlekarstvo | 9 |
| Duvan | 1 |
| **Ukupno** | **48** |

`INFERENCE`: Infosys ima dokazanu ili istorijsku prisutnost upravo u sektorima koji su najbliži AgriX Enterprise tržištu. Replacement segment je značajno širi od dva već migrirana klijenta.

`DECISION`: ovih 48 stavki postaju seed lista za Infosys replacement-market istraživanje, ali se ne kontaktiraju pre potvrde pravnog identiteta, aktivnog statusa, procesa i aktuelnog sistema.

---

## 5. Geografski signali

Najzastupljenije navedene lokacije:

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

Visokopotencijalni klasteri posebno se vide u:

- Požegi;
- Arilju;
- Čačku;
- Zlatiboru;
- Bajinoj Bašti;
- Kosjeriću;
- Kotraži;
- Užicu i okolini.

`INFERENCE`: Infosys ima jaku referentnu gustinu u Zapadnoj i Centralnoj Srbiji, što se preklapa sa početnim AgriX tržišnim fokusom i povećava verovatnoću replacement prodaje.

Dataset sadrži i regionalne reference u Pljevljima, Srebrenici i Zvorniku.

`INFERENCE`: Infosys reference pružaju prvi konkretan signal da regionalni konkurentski prostor postoji i van Srbije, ali tri lokacije nisu dovoljne za procenu regionalnog tržišnog udela.

---

## 6. Najvažniji strateški zaključci

1. Infosys nije samo jedan od konkurenata; njegova javna lista pokazuje široku agro/prehrambenu bazu.
2. Dva stvarna prelaska na AgriX potvrđuju da je deo te baze osvojiv.
3. Najvredniji replacement pool nije svih 113 stavki, već prvo 48 visokopotencijalnih.
4. APR šifre `1039` i `4631` propuštaju mlekare, duvan, kooperative i druge procese koji se pojavljuju u Infosys bazi.
5. Regionalna prodaja treba da počne od klastera gde postoje reference, procesna sličnost i mogućnost preporuke.
6. AgriX treba da dokumentuje migracioni alat i playbook za prelazak sa Infosys-a, jer je migracija sada potvrđen prodajni use case.

---

## 7. Obavezna validacija pre targetiranja

Za svaku od 48 visokopotencijalnih stavki dopuniti:

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

---

## 8. Prioritetna radna lista

### Prioritet A

- 25 firmi iz kategorije voće, povrće i hladnjače;
- 13 firmi iz kategorije poljoprivreda i kooperative;
- 1 duvanski račun.

### Prioritet B

- 9 mlekara, nakon provere koliko se njihovi tokovi mogu pokriti bez forka proizvoda.

### Prioritet C

- srednji potencijal: žitarice, prehrambena proizvodnja, meso, piće i pekarstvo;
- koristiti za adjacent-market istraživanje, ne za trenutni generički outbound.

### Isključiti iz prvog talasa

- agro-inpute bez otkupnog/kooperantskog toka;
- institute i strukovne organizacije;
- veterinarske i druge organizacije bez potvrđenog AgriX procesa;
- svih 24 stavke označene kao `Potrebna provera` dok se pravni identitet i delatnost ne potvrde.
