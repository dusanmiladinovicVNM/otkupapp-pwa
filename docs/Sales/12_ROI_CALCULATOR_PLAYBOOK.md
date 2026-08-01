# AgriX ROI Calculator Playbook

**Status:** DRAFT v1 — VALIDATION  
**Datum:** 30.07.2026.  
**Vlasnik:** osnivač AgriX-a  
**Svrha:** Standardizovati kako se procenjuju, proveravaju i predstavljaju ekonomski efekti AgriX-a bez nerealnih obećanja, dvostrukog računanja koristi ili predstavljanja procene kao garantovanog rezultata.

---

## 1. Osnovni princip

ROI kalkulator nije prodajni trik i ne služi da matematički „dokaže“ kupovinu. Njegova uloga je da:

1. prevede potvrđeni poslovni problem u proverljive ulaze;
2. prikaže konzervativan raspon mogućeg efekta;
3. odvoji direktno merljivu korist od kvalitativne koristi i smanjenja rizika;
4. učini pretpostavke vidljivim kupcu;
5. pokaže kada investicija nema dovoljno jak poslovni osnov;
6. postavi baseline za kasnije merenje stvarnih rezultata.

**DECISION:** ROI se nikada ne predstavlja kao garantovan rezultat. Svaki izlaz je procena zasnovana na ulazima koje je potvrdio kupac i na jasno označenim pretpostavkama.

---

## 2. Kada se ROI model koristi

ROI model se koristi tek kada su poznati najmanje:

- stvarni trenutni proces;
- relevantan obim rada;
- glavni problem i posledica;
- očekivani scope AgriX-a;
- uloge koje učestvuju u procesu;
- period u kom se vrednost posmatra;
- osoba koja može da potvrdi ključne ulaze.

Ne koristi se:

- pre discovery-ja;
- kao zamena za kvalifikaciju;
- kada kupac ne može ili ne želi da potvrdi nijedan ulaz;
- kada je problem isključivo „želimo moderniji softver“ bez poslovne posledice;
- kada scope, cena ili način implementacije još nisu dovoljno poznati;
- kada bi model zahtevao izmišljanje podataka.

---

## 3. Nivoi pouzdanosti ulaza

Svaki ulaz dobija oznaku:

- **A — VERIFIED:** izmeren iz sistema, evidencije, uzorka dokumenata ili vremenskog praćenja;
- **B — CONFIRMED:** potvrdio vlasnik procesa ili više sagovornika, ali nije sistemski izmeren;
- **C — ESTIMATED:** razumna procena kupca ili AgriX-a, uz vidljivo objašnjenje;
- **D — PLACEHOLDER:** privremena vrednost koja ne sme u finalni business case.

Finalni ROI ne sme biti označen kao „validated“ ako:

- više od 30% ukupne procenjene koristi potiče iz C ulaza;
- bilo koji ključni ulaz ostane D;
- glavni finansijski efekat zavisi od jedne nepotvrđene pretpostavke;
- korist uključuje dvostruko računanje istog vremena ili iste greške.

---

## 4. Obavezni scenario model

Svaki ROI prikazuje najmanje tri scenarija:

| Scenario | Princip |
|---|---|
| Konzervativni | najniži razuman efekat, sporije usvajanje i samo direktno dokazive koristi |
| Osnovni | najverovatniji efekat na osnovu potvrđenih ulaza |
| Gornji | ostvariv, ali ne garantovan efekat uz dobro usvajanje i stabilan scope |

Komercijalna prezentacija počinje konzervativnim scenarijem. Gornji scenario se ne koristi kao headline vrednost.

---

## 5. Vremenski horizont

Standardni horizonti:

- **sezona** — kada se efekat vezuje za konkretnu otkupnu sezonu;
- **12 meseci** — podrazumevani poslovni pogled;
- **24 meseca** — koristi se za kumulativnu vrednost i stabilizaciju usvajanja;
- **36 meseci** — samo kada je pricing i scope dovoljno stabilan za takav prikaz.

Model mora jasno razdvojiti:

- jednokratne troškove;
- godišnje ili mesečne troškove;
- sezonske troškove;
- jednokratnu korist;
- ponavljajuću korist;
- korist koja se javlja tek posle implementacije i usvajanja.

---

## 6. Ukupan trošak investicije

### 6.1. Troškovi AgriX-a

Uključuju, prema važećem cenovniku i potvrđenom scope-u:

- licencu ili pretplatu;
- module;
- dodatne firme i instance;
- stanice i Mobile komponente;
- implementaciju;
- migraciju;
- obuku;
- odobren custom razvoj;
- hardver kada je deo ponude;
- druge eksplicitno ugovorene stavke.

Cene se preuzimaju isključivo iz važećeg izvora istine `AgriX_Cenovnik_2027.html` i potvrđene ponude.

### 6.2. Interni troškovi kupca

Po potrebi uključuju:

- vreme ključnih korisnika za radionice i obuku;
- pripremu i čišćenje podataka;
- internu IT podršku;
- privremeni paralelni rad;
- procesne promene;
- putne ili infrastrukturne troškove.

**DECISION:** Interni troškovi kupca ne kriju se da bi ROI izgledao bolje.

### 6.3. Formula ukupne investicije

`Ukupna investicija = jednokratni AgriX troškovi + periodični AgriX troškovi u horizontu + procenjeni interni troškovi kupca`

---

## 7. Stubovi poslovne vrednosti

### 7.1. Ušteda administrativnog vremena

Primene:

- manje ručnog prepisivanja;
- manje duplog unosa;
- brže generisanje dokumenata;
- manje traženja podataka;
- brže usaglašavanje stanica i centrale;
- manje ručnih izveštaja.

Formula po aktivnosti:

`Godišnja vrednost vremena = broj izvršenja × ušteđeni minuti / 60 × potpuno opterećena satnica`

Obavezni ulazi:

- aktivnost;
- broj izvršenja po periodu;
- trenutno vreme;
- očekivano buduće vreme;
- broj ljudi;
- satnica;
- izvor ulaza;
- nivo pouzdanosti.

Ne sme se računati celo oslobođeno vreme kao finansijska ušteda ako ono realno neće:

- smanjiti prekovremeni rad;
- smanjiti potrebu za dodatnim angažovanjem;
- omogućiti veći obim bez dodatne osobe;
- biti preusmereno na eksplicitno vredniji posao.

Zato model razdvaja:

- **kapacitet oslobođen**;
- **direktno monetizovana ušteda**.

### 7.2. Izbegnuti trošak dodatnog zapošljavanja

Formula:

`Izbegnuti trošak = verovatnoća dodatnog zapošljavanja × godišnji potpuno opterećeni trošak uloge`

Koristi se samo kada kupac potvrdi da bi bez promene procesa realno angažovao dodatnu osobu, sezonskog radnika ili eksternog saradnika.

### 7.3. Smanjenje grešaka i korektivnog rada

Primeri:

- pogrešan unos;
- nedostajući dokument;
- duplikat;
- neusaglašena količina;
- pogrešno vezana stanica ili dobavljač;
- korekcije koje zahtevaju više ljudi;
- ponovna izrada dokumenata.

Formula:

`Vrednost smanjenja grešaka = broj grešaka × prosečan interni trošak greške × očekivani procenat smanjenja`

Trošak greške može uključiti:

- vreme ispravke;
- vreme kontrole;
- dodatnu komunikaciju;
- ponovno štampanje ili obradu;
- dokaziv direktni finansijski gubitak.

Ne uključuje se hipotetička kazna ili ekstremni gubitak bez realne učestalosti i dokazive veze.

### 7.4. Brže zatvaranje i izveštavanje

Formula:

`Vrednost = broj ciklusa × ušteđeni sati po ciklusu × relevantna satnica`

Dodatno se kvalitativno beleži:

- koliko ranije uprava dobija pregled;
- koje odluke tada može doneti;
- koje zavisnosti ili kašnjenja nestaju.

Vrednost „bolje odluke“ ne monetizuje se bez potvrđenog mehanizma i podataka.

### 7.5. Kontrola robe, ambalaže i sledljivost

Moguće direktne komponente:

- manje izgubljene ili neusklađene ambalaže;
- manje ručnih popisa i usaglašavanja;
- brže rešavanje sporova;
- manje otpisanih ili neobjašnjenih razlika;
- manje vremena za rekonstrukciju toka robe.

Formula za potvrđene gubitke:

`Vrednost smanjenja gubitka = istorijski godišnji gubitak × konzervativni procenat izbegavanja`

Ako istorijski gubitak nije dokumentovan, korist ostaje kvalitativna ili se prikazuje kao zasebna sensitivity pretpostavka.

### 7.6. Veći obim bez proporcionalnog rasta administracije

Formula:

`Vrednost kapaciteta = dodatni obim koji postojeći tim može obraditi × doprinos po jedinici`,

ali samo kada:

- postoji realan demand ili plan rasta;
- administracija je dokazano ograničenje;
- doprinos po jedinici je potvrđen;
- efekat nije već uračunat kroz izbegnuto zapošljavanje.

### 7.7. Smanjenje operativnog rizika

Rizik se prikazuje odvojeno od direktnog ROI-ja.

Standardna formula očekivane vrednosti:

`Očekivani godišnji gubitak = verovatnoća događaja × finansijska posledica`

`Očekivano smanjenje rizika = očekivani godišnji gubitak × procenjeni procenat smanjenja`

Koristi se samo kada postoje:

- istorijski događaji;
- dokumentovana učestalost;
- realna posledica;
- jasna veza između AgriX kontrole i smanjenja rizika.

Bez toga se rizik prikazuje opisno: nizak, srednji ili visok poslovni značaj, bez lažne monetizacije.

---

## 8. Faktor realizacije koristi

Teorijska ušteda se ne uzima u celosti. Uvodi se faktor realizacije:

`Realizovana korist = bruto procenjena korist × faktor usvajanja × faktor procesne discipline × faktor pokrivenosti scope-a`

Tipični rasponi za početni model:

- usvajanje: 50–90%;
- procesna disciplina: 60–95%;
- pokrivenost scope-a: 50–100%.

Ovi rasponi nisu univerzalne činjenice. Moraju se prilagoditi kupcu i označiti kao pretpostavke dok ne postoje stvarni podaci.

---

## 9. Ramp-up po periodima

Konzervativni model pretpostavlja da korist ne nastaje u punom iznosu prvog dana.

Primer ramp-up modela:

| Period | Udeo pune koristi |
|---|---:|
| Implementacija | 0% |
| Prvi operativni mesec | 25–50% |
| Drugi mesec | 50–75% |
| Stabilizovan rad | 75–100% |

Za sezonske procese koristi se raspored prema stvarnoj fazi sezone, a ne generički mesečni model.

---

## 10. Ključni izlazi

### 10.1. Neto korist

`Neto korist = realizovana poslovna korist − ukupna investicija`

### 10.2. ROI procenat

`ROI % = (realizovana poslovna korist − ukupna investicija) / ukupna investicija × 100`

### 10.3. Payback period

`Payback period = ukupna investicija / prosečna mesečna realizovana korist`

Kod sezonskog poslovanja payback se izražava i kao:

- broj meseci;
- deo sezone;
- broj punih sezona.

### 10.4. Benefit-cost ratio

`BCR = realizovana poslovna korist / ukupna investicija`

### 10.5. Trogodišnji TCO i net benefit

Za 36 meseci:

`TCO36 = svi jednokratni + svi periodični + interni troškovi u 36 meseci`

`NetBenefit36 = kumulativna realizovana korist − TCO36`

---

## 11. Pravila protiv dvostrukog računanja

Ista korist ne sme biti uračunata kroz više kategorija.

Tipični duplikati:

- ušteda vremena i izbegnuto zapošljavanje;
- manje grešaka i isto vreme korekcije;
- dodatni kapacitet i prihod koji proizlazi iz istog oslobođenog vremena;
- izbegnuti gubitak robe i isti događaj kao risk reduction;
- brže izveštavanje i generička „bolja odluka“ bez dodatnog dokaza.

Za svaku korist mora postojati polje `overlap check` i referenca na povezane stavke.

---

## 12. Sensitivity analiza

Obavezno se testiraju najmanje tri najuticajnija ulaza.

Najčešći kandidati:

- broj dokumenata ili transakcija;
- ušteđeno vreme po transakciji;
- satnica;
- procenat smanjenja grešaka;
- faktor usvajanja;
- procenat smanjenja gubitka;
- trajanje implementacije.

Prikaz:

| Ulaz | Niska vrednost | Osnovna | Visoka | Uticaj na net benefit |
|---|---:|---:|---:|---:|
| Primer: ušteda minuta | 2 | 4 | 6 | izračunato |

Model mora pokazati i break-even vrednost:

- minimalan broj transakcija;
- minimalna ušteda minuta;
- minimalni procenat usvajanja;
- maksimalan dozvoljeni trošak implementacije;
- minimalni period korišćenja.

---

## 13. Conservative haircut

Kada postoji visok stepen neizvesnosti, na procenjenu korist primenjuje se dodatni haircut od 10–40%.

Haircut je obavezan kada:

- baseline potiče iz malog uzorka;
- sagovornici daju značajno različite procene;
- proces zavisi od više eksternih sistema;
- planirano usvajanje nije potvrđeno;
- koristi se nova ili nevalidirana funkcionalnost;
- deo scope-a još nije definitivno potvrđen.

Razlog i procenat haircuta moraju biti zapisani.

---

## 14. Šta ulazi u headline business case

Headline business case sme da sadrži samo:

- direktne uštede zasnovane na A ili B ulazima;
- konzervativni ili osnovni scenario;
- troškove iz potvrđenog scope-a;
- jasno naveden horizont;
- vidljivo ograničenje i pretpostavke.

Ne ulaze kao headline:

- reputaciona korist;
- generička digitalizacija;
- subjektivni osećaj kontrole;
- neprovereni budući prihod;
- maksimalni scenario;
- ekstremni rizici bez istorijskog osnova;
- funkcije koje nisu potvrđene kao deo scope-a.

---

## 15. Prezentacija kupcu

Preporučeni redosled:

1. potvrda problema i obima;
2. pregled ulaza koje je kupac dao;
3. razlikovanje činjenica i pretpostavki;
4. konzervativni scenario;
5. osnovni scenario;
6. trošak i vremenski raspored;
7. payback i break-even;
8. sensitivity analiza;
9. kvalitativne koristi i rizici odvojeno;
10. dogovor šta treba dodatno proveriti.

Dozvoljena formulacija:

> „Na osnovu podataka koje smo zajedno uneli, konzervativni scenario pokazuje procenjeni raspon. Ovo nije garancija rezultata; najveća neizvesnost je u usvajanju i stvarnoj uštedi vremena po dokumentu.“

Nedozvoljena formulacija:

> „AgriX će vam sigurno uštedeti ovaj iznos i isplatiti se za tri meseca.“

---

## 16. ROI Discovery pitanja

### Obim

- Koliko dokumenata, prijema, proizvođača i stanica obrađujete po sezoni?
- Koliko ljudi učestvuje u procesu?
- Koliko puta se isti podatak prepisuje?
- Koliki je vršni dnevni obim?

### Vreme

- Koliko traje obrada jednog tipičnog dokumenta?
- Koliko traje dnevno ili nedeljno usaglašavanje?
- Koliko se vremena troši na traženje podataka i korekcije?
- Koliko prekovremenog rada postoji u špicu?

### Greške

- Koje greške se najčešće javljaju?
- Koliko često?
- Ko ih otkriva i ispravlja?
- Koliko ljudi i vremena je potrebno po grešci?
- Postoje li direktni finansijski gubici?

### Kontrola i rizik

- Koji podaci kasne do uprave?
- Koji sporovi ili neusaglašenosti se ponavljaju?
- Koliki su istorijski gubici ambalaže, robe ili dokumentacije?
- Koji događaj bi imao najveću poslovnu posledicu?

### Rast

- Da li trenutna administracija ograničava broj stanica, kultura ili dobavljača?
- Da li biste bez promene sistema morali da angažujete dodatne ljude?
- Da li postoji potvrđen plan rasta?

---

## 17. ROI Assumption Register

Za svaki model vodi se tabela:

| ID | Ulaz | Vrednost | Jedinica | Izvor | Pouzdanost | Scenario | Vlasnik potvrde | Datum |
|---|---|---:|---|---|---|---|---|---|

Dodatna obavezna polja:

- formula u kojoj se koristi;
- da li postoji overlap;
- datum poslednje provere;
- komentar kupca;
- status: open, confirmed, rejected, replaced.

---

## 18. CRM ROI Record

CRM zapis sadrži:

- opportunity ID;
- datum modela i verziju;
- poslovni horizont;
- potvrđeni scope;
- ukupnu investiciju;
- konzervativnu, osnovnu i gornju korist;
- net benefit po scenariju;
- ROI % po scenariju;
- payback po scenariju;
- benefit-cost ratio;
- tri najosetljivija ulaza;
- broj A/B/C/D ulaza;
- ključni haircut;
- glavna kvalitativna korist;
- glavni rizik;
- osoba koja je potvrdila ulaze;
- datum review razgovora;
- odluka i sledeći korak.

Poverljivi finansijski podaci kupca ne čuvaju se u javnom repozitorijumu.

---

## 19. ROI Quality Score

Maksimum: 20 bodova.

| Kriterijum | Bodovi |
|---|---:|
| Potvrđen baseline i obim | 0–2 |
| Trošak obuhvata sve relevantne stavke | 0–2 |
| Direktne koristi imaju formulu i izvor | 0–2 |
| Činjenice i procene su razdvojene | 0–2 |
| Postoje tri scenarija | 0–2 |
| Uključen faktor realizacije i ramp-up | 0–2 |
| Proveren overlap | 0–2 |
| Urađena sensitivity analiza | 0–2 |
| Vidljivi rizici i ograničenja | 0–2 |
| Kupac je pregledao i potvrdio ulaze | 0–2 |

Tumačenje:

- 17–20: spremno za business-case razgovor;
- 13–16: upotrebljivo uz vidljive rezerve;
- 9–12: samo radni model;
- 0–8: ne koristiti u prodajnoj odluci.

---

## 20. Red flags

- svi ulazi dolaze od prodavca;
- jedna osoba nagađa kompletan proces;
- najveća korist zavisi od nepotvrđenog rasta;
- model računa 100% teorijske uštede;
- interni troškovi kupca su izostavljeni;
- isti efekat se računa kroz više kategorija;
- gornji scenario se predstavlja kao očekivani;
- rizik se monetizuje bez istorijskih podataka;
- ROI se koristi da prikrije slab product fit;
- proračun se menja dok ne pokaže željeni payback;
- kupcu se ne pokazuje assumption register;
- kalkulator koristi zastarele cene.

---

## 21. Zabranjene prakse i tvrdnje

Zabranjeno je:

- garantovati procenat uštede, ROI ili payback;
- izmišljati benchmark za sektor;
- koristiti rezultate jednog klijenta kao automatsku pretpostavku za drugog;
- prikazivati maksimalni scenario kao „realan“ bez dokaza;
- računati oslobođeno vreme kao gotovinsku uštedu bez mehanizma realizacije;
- skrivati troškove migracije, obuke ili paralelnog rada;
- računati funkcionalnost koja nije ugovorena ili stabilna;
- koristiti lažnu preciznost, npr. rezultat na dve decimale kada su ulazi grube procene;
- tvrditi da regulatorni ili operativni rizik potpuno nestaje;
- predstaviti procenu kao nezavisnu finansijsku analizu.

---

## 22. Validacioni plan

Prvih deset ROI modela koristi se za validaciju:

- koliko lako kupci daju pouzdane ulaze;
- koje kategorije koristi najčešće imaju stvarni dokaz;
- koji ulazi najviše utiču na odluku;
- koliko se procene razlikuju od post-implementation rezultata;
- koji haircut je realan po tipu kupca;
- da li je payback relevantniji po mesecima ili sezonama;
- gde se najčešće pojavljuje dvostruko računanje;
- koje metrike treba automatski prikupljati iz AgriX-a.

Posle svake završene implementacije, kada postoji dozvola i odgovarajući podaci, porede se:

- projected conservative;
- projected base;
- actual result;
- razlog odstupanja;
- promena scope-a;
- nivo usvajanja;
- validnost početnih pretpostavki.

---

## 23. Operativna checklist-a

### Pre izrade

- [ ] Discovery je završen.
- [ ] Scope je dovoljno poznat.
- [ ] Identifikovan je vlasnik ulaza.
- [ ] Prikupljen je baseline.
- [ ] Cene su iz važećeg izvora istine.

### Tokom izrade

- [ ] Svaki ulaz ima izvor i nivo pouzdanosti.
- [ ] Izračunata su tri scenarija.
- [ ] Uključen je ramp-up.
- [ ] Uključen je faktor realizacije.
- [ ] Proveren je overlap.
- [ ] Uključeni su svi relevantni troškovi.
- [ ] Urađena je sensitivity analiza.
- [ ] Izračunat je break-even.

### Pre prezentacije

- [ ] Headline koristi konzervativni ili osnovni scenario.
- [ ] Pretpostavke i ograničenja su vidljivi.
- [ ] Kvalitativna korist je odvojena od finansijske.
- [ ] Nema garantovanih tvrdnji.
- [ ] Kupac može da menja i proverava ulaze.

### Posle razgovora

- [ ] Zapisane su osporene pretpostavke.
- [ ] Dodeljeni su vlasnici provere.
- [ ] Model je verzionisan.
- [ ] CRM zapis je ažuriran.
- [ ] Dogovoren je sledeći korak.

---

## 24. Veze sa drugim dokumentima

- `03_BUYING_PROCESS.md` — decision criteria i business case faza;
- `04_SALES_PROCESS.md` — stage advancement i forecast;
- `05_DISCOVERY_PLAYBOOK.md` — prikupljanje procesa, posledica i success criteria;
- `08_DEMO_PLAYBOOK.md` — dokazivanje fit-a za koristi koje model pretpostavlja;
- `09_OBJECTION_HANDLING.md` — cena, vrednost i budžet;
- `10_NEGOTIATION_PLAYBOOK.md` — scope i komercijalni uslovi;
- `11_CASE_STUDIES_PLAYBOOK.md` — actual results i proof hierarchy;
- `AgriX_Cenovnik_2027.html` — jedini izvor istine za cene.

---

## 25. Hipoteze za validaciju

- **HYPOTHESIS:** kod manjih otkupljivača ušteda vremena i izbegavanje dodatne administracije biće glavni finansijski stub;
- **HYPOTHESIS:** kod više stanica najveću vrednost ima centralna kontrola, usaglašavanje i sledljivost;
- **HYPOTHESIS:** payback izražen u sezonama biće razumljiviji od godišnjeg ROI procenta;
- **HYPOTHESIS:** konzervativni scenario povećava poverenje više nego agresivna procena;
- **HYPOTHESIS:** vlasnici će kvalitativnu kontrolu smatrati važnijom od dela lako merljivih ušteda;
- **HYPOTHESIS:** stvarni post-implementation podaci će zahtevati niže faktore realizacije u prvoj sezoni nego u drugoj.

---

## 26. Definition of Done

ROI model za konkretnu priliku je završen kada:

- svi ključni ulazi imaju izvor i nivo pouzdanosti;
- cena odgovara potvrđenom scope-u;
- prikazana su najmanje tri scenarija;
- uračunati su realizacija, ramp-up i svi relevantni troškovi;
- uklonjena su preklapanja;
- urađeni su sensitivity i break-even;
- kupac je pregledao ključne pretpostavke;
- ograničenja su eksplicitna;
- postoji verzionisan ROI record i sledeći korak.

Segment ostaje `DRAFT v1 — VALIDATION` dok se model ne uporedi sa stvarnim rezultatima više implementacija.