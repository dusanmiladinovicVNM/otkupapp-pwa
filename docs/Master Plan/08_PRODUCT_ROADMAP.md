# 08 — Product roadmap AgriX-a

**Status:** Review  
**Vlasnik:** osnivač AgriX-a  
**Horizont:** 23.07.2026–31.12.2027  
**Poslednje ažuriranje:** 2026-07-23  
**Povezani dokumenti:** `02_STRATEGY.md`, `03_CUSTOMERS_AND_JOBS.md`, `06_POSITIONING.md`, `07_PRODUCT_PORTFOLIO.md`, `07A_PRODUCT_STATUS_MATRIX.csv`, `09_TECHNOLOGY_STRATEGY.md`, `10_PRICING_AND_PACKAGING.md`, `14_GO_TO_MARKET.md`, `16_ONBOARDING_AND_IMPLEMENTATION.md`, `../ROADMAP.md`, `../KNOWN_ISSUES.md`, `../RELEASE_GATES.md`

---

## 1. Svrha

Ovaj dokument određuje:

- kojim redosledom AgriX ulaže razvojni kapacitet;
- šta mora biti završeno pre naredne pune sezone;
- koje funkcionalnosti prelaze iz `pilot` u `standard offer`;
- koje poslovne i tehničke gate-ove mora da prođe svaka faza;
- šta se svesno odlaže;
- kada se razvoj zaustavlja radi stabilizacije;
- kako se štiti Enterprise Core od širenja scope-a;
- koliko kapaciteta ostaje za Gazdinstvo i GGAP;
- koji merljivi dokazi odlučuju o nastavku, promeni ili gašenju inicijative.

Roadmap nije spisak funkcija koje bi bilo dobro imati. On je redosled ograničenih investicionih odluka u uslovima u kojima razvoj i dalje dominantno vodi osnivač.

---

## 2. Glavna roadmap odluka

Redosled prioriteta je:

1. **istinitost podataka i release-a**;
2. **ponovljiv Enterprise Core**;
3. **standardizovan onboarding i migracija**;
4. **Field Operations pilot: PWA Otkupac, kiosk i štampa**;
5. **stabilizacija kroz realnu sezonu**;
6. **Gazdinstvo activation/retention dokaz**;
7. **GGAP discovery i ograničeni pilot**;
8. **širenje na nove segmente i region tek kada prethodni sloj ima dokaz.**

`DECISION`: nijedna nova funkcija nema prioritet nad otvorenim P0 problemom koji može da ugrozi podatke, dokumente, finansije, autorizaciju ili istinitost release gate-a.

---

## 3. Roadmap principi

### 3.1 Core pre širine

AgriX već ima veliku funkcionalnu širinu. Najveći rizik nije nedostatak još jednog modula, već da postojeća širina:

- nije jednako pouzdana u svim tokovima;
- zavisi od osnivača tokom onboarding-a ili incidenta;
- nije dovoljno jasno upakovana;
- nije prošla realan sezonski dokaz u novom Field Operations modelu.

Zato se roadmap ne meri brojem novih funkcija, već brojem celina koje su prešle iz `implemented` u `production-proven` i `standard offer`.

### 3.2 Gate pre datuma

Datumi u ovom dokumentu su ciljni prozori. Faza se ne završava zato što je istekao kalendarski rok, već kada su ispunjeni acceptance kriterijumi.

### 3.3 Jedan codebase

Nijedna roadmap inicijativa ne sme da stvori trajni klijentski fork. Novi zahtev mora biti:

- reusable funkcija;
- konfiguracija;
- feature flag;
- adapter sa jasnim ugovorom;
- ili odbijen/odložen.

### 3.4 Sezonski freeze

Najmanje 30 dana pre kritične sezone pilot-klijenta uvodi se feature freeze za pogođene tokove. Nakon freeze-a dozvoljeni su samo:

- P0/P1 ispravke;
- release-gate popravke;
- konfiguracija;
- dokumentacija i obuka;
- promene koje smanjuju, a ne povećavaju, operativni rizik.

### 3.5 Dokaz pre skaliranja

Jedan uspešan demo ili interna proba nije production evidence. Za sezonske tokove traži se najmanje jedan realan pilot kroz relevantan poslovni ciklus.

---

## 4. Polazno stanje na 23.07.2026.

### 4.1 Komercijalni status

| Celina | Trenutni status |
|---|---|
| Enterprise Core | `Standard offer` |
| Management PWA | production kanal u okviru Enterprise Core-a |
| Finance & Regulatory | `Optional extension` |
| Logistics & Fleet | `Optional extension`, ograničen dokaz |
| Agrohemija i kooperantska zaduženja | `Optional extension`, ograničen dokaz |
| Advanced Warehouse / Pallets / Traceability | discovery-scoped ekstenzija |
| PWA Otkupac | `Pilot only` |
| Kiosk i termalna štampa | `Pilot only` |
| Gazdinstvo Partner/Basic/Pro | kontrolisani pilot |
| AgriX GGAP | discovery; `Not for sale` kao završen proizvod |

### 4.2 Aktivni tehnički rizik

Postojeći audit i known-issues registri sadrže aktivne P0/P1 nalaze. Posebno su roadmap relevantni:

- Sheets JSON parsing koji može pogrešno pročitati vrednosti sa zarezima, navodnicima ili Unicode escape sekvencama;
- SEF statusni tok u kojem određeni odgovor može biti pogrešno mapiran;
- runtime gate-ovi koji moraju biti izvršeni u pravom workbook/GAS/PWA okruženju;
- autorizaciona nedoumica oko `saveParcelPolygon`;
- phantom opening-debt red u PWA magacinskoj projekciji;
- neujednačena negativna/regression pokrivenost kritičnih tokova.

`DECISION`: poslovni roadmap ne proglašava proizvod spremnijim od tehničkog release evidence-a.

---

## 5. Faza 0 — Core safety i release truth

**Ciljni prozor:** 23.07.2026–31.08.2026  
**Primarni cilj:** ukloniti ili eksplicitno prihvatiti svaki launch-relevant P0 i dokazati da release gate meri stvarno stanje.

### 5.1 Obavezni ishodi

1. Rebazirati aktivne P0 nalaze na trenutni `main` i potvrditi da i dalje postoje.
2. Ispraviti potvrđene P0 data-safety i SEF statusne defekte.
3. Izvršiti relevantne VBA, GAS, PWA, monitoring i production-health gate-ove u realnom okruženju.
4. Razrešiti deployed authorization stanje `saveParcelPolygon`.
5. Rešiti ili izolovati `ART_POCETNI_DUG` iz PWA stock projekcije.
6. Uvesti jedinstven zapis rezultata release-a: verzija, commit, workbook, environment, datum, gate rezultati i poznati rizici.
7. Razdvojiti testirano u kodu od testirano u produkcionom workbook-u.

### 5.2 Exit gate

Faza 0 je završena tek kada:

- nema otvorenog nepotvrđenog P0 nalaza;
- svaki preostali P0 ima eksplicitnu odluku, owner-a i kratkoročni containment;
- `RunProductionHealthCheck` ima nula failure-a na target workbook-u;
- obavezni compile/smoke/gate rezultati su sačuvani;
- nema false-green release ugovora;
- poznati rizici su ažurirani i povezani sa roadmap stavkama.

### 5.3 Stop pravilo

Dok Faza 0 nije završena:

- nema razvoja novih GGAP funkcija;
- nema novih Gazdinstvo premium funkcija;
- nema širenja Field Operations scope-a osim popravki potrebnih za pilot gate;
- nema nove banke/ERP integracije bez ugovorenog klijenta i posebnog kapaciteta.

---

## 6. Faza 1 — Ponovljiv Enterprise Core

**Ciljni prozor:** 01.08.2026–31.10.2026  
**Primarni cilj:** standardni klijent može biti konfigurisan, migriran, obučen i pušten u rad bez skrivenih koraka koje zna samo osnivač.

Faza 1 može delimično teći paralelno sa Fazom 0 samo za dokumentaciju, konfiguraciju i onboarding rad koji ne povećava tehnički scope.

### 6.1 Proizvodni ishodi

- canonical minimalni Enterprise Core scope;
- standardni tenant/client configuration paket;
- feature-flag matrica;
- standardna nova sezona i multi-company konfiguracija;
- jasno označene podržane i nepodržane varijante procesa;
- stabilni default izveštaji;
- standardan production workbook build i deployment paket;
- operativno merljiv monitoring i health pregled po klijentu.

### 6.2 Onboarding ishodi

- discovery obrazac;
- data-mapping šablon;
- migracioni input format;
- setup checklist za računar, Google, PWA i pristupe;
- standardna obuka po ulozi;
- go-live checklist;
- acceptance zapis;
- rollback i recovery procedura;
- handoff prema support-u;
- ciljano vreme standardnog onboarding-a do približno pola dana nakon pripreme podataka.

### 6.3 Infosys migration paket

Bez čekanja na završne win intervjue mogu se standardizovati:

- inventory izvornog sistema;
- data-export zahtevi;
- mapiranje matičnih podataka;
- otvorena stanja;
- dokumenti i istorija;
- cutover datum;
- paralelni rad;
- acceptance reconciliation;
- rollback odluka;
- migraciona cena kao posebna usluga.

Win intervjui će kasnije dopuniti razloge prelaska, battlecard i prodajnu poruku.

### 6.4 Exit gate

- najmanje jedan novi ili ponovljeni onboarding izveden po checklisti;
- nema undocumented founder-only koraka;
- svi standardni podaci imaju ulazni format i validaciju;
- podrška može da dijagnostikuje tipične probleme iz logova/monitoringa;
- onboarding vreme i greške su izmereni;
- odstupanja od standarda su eksplicitno evidentirana.

---

## 7. Faza 2 — Field Operations pilot readiness

**Ciljni prozor:** 01.09.2026–28.02.2027  
**Primarni cilj:** PWA Otkupac, kiosk i termalna štampa spremni su za kontrolisani realni pilot, ali još nisu `Standard offer`.

### 7.1 PWA Otkupac obavezni scope

- identifikacija i izbor kooperanta;
- stanica, kultura, klasa, cena i ambalaža;
- bruto/neto i podržane varijante otkupa;
- offline unos;
- stabilan local queue;
- retry i stale-sync recovery;
- idempotency i duplicate protection;
- jasan status `pending/syncing/synced/error`;
- dokument/print payload;
- controlled correction/storno proces;
- import u desktop bez ručnog popravljanja.

### 7.2 Kiosk i štampa

- odobren jedan standardni tablet model ili jasno definisan minimalni profil;
- odobren jedan standardni termalni printer;
- stabilan način štampe;
- kiosk zaključavanje;
- remote support pristup;
- rezervni uređaj;
- recovery posle gubitka veze, restarta i neuspele štampe;
- asset evidencija, garancija i zamena;
- poseban hardverski trošak i marža.

### 7.3 Pilot dizajn

Pilot mora unapred definisati:

- jednu imenovanu firmu;
- 1–3 stanice u prvom talasu;
- jednu glavnu kulturu/sezonu;
- maksimalan dnevni obim;
- fallback proceduru;
- osobu odgovornu kod klijenta;
- AgriX owner-a;
- dnevni monitoring tokom kritičnog perioda;
- success i abort kriterijume.

### 7.4 Exit gate za početak pilota

- end-to-end test bez ručnog popravljanja baze;
- offline/retry/idempotency regression green;
- print success na odobrenom hardveru;
- operator može da vidi i razume sync grešku;
- podrška može remote da dijagnostikuje uređaj;
- duplicate i data-loss testovi prolaze;
- pilot ugovor jasno kaže da je tok pilot;
- fallback i rollback su probani pre prvog realnog otkupa.

---

## 8. Faza 3 — Realni sezonski Field Operations pilot

**Ciljni prozor:** prema kulturi pilot-klijenta, okvirno mart–jul 2027.  
**Primarni cilj:** dokazati pouzdanost pod realnim pritiskom, ne dodavati širinu.

### 8.1 KPI-jevi pilota

- broj otkupa;
- prosečno i p95 vreme unosa;
- procenat offline unosa;
- sync success rate bez ručne intervencije;
- duplicate rate;
- data-loss rate;
- print success rate;
- broj blokirajućih incidenata;
- support minuti po stanici nedeljno;
- vreme oporavka;
- broj ručnih desktop korekcija;
- zadovoljstvo vagača i centralnog operatera;
- spremnost klijenta da plati standardnu cenu.

### 8.2 Minimalni pragovi za razmatranje Standard offer statusa

Radni pragovi koji se potvrđuju pre pilota:

- nula izgubljenih potvrđenih poslovnih zapisa;
- nula nekontrolisanih duplikata koji uđu u canonical master;
- najmanje 99% uspešnih sinhronizacija bez developerske intervencije;
- najmanje 98% uspešnih štampi iz prvog ili kontrolisanog retry pokušaja;
- svi P0 incidenti zatvoreni sa root-cause analizom;
- support opterećenje dovoljno nisko za planirani broj stanica;
- klijent prihvata nastavak produkcionog korišćenja.

Ovi pragovi su početna radna pretpostavka i menjaju se samo uz dokumentovanu odluku, ne retroaktivno radi proglašenja pilota uspešnim.

### 8.3 Feature freeze

Tokom pilota:

- ne dodavati nove poslovne varijante osim ako blokiraju ugovoreni tok;
- ne menjati podatkovni ugovor bez migracije i regression testa;
- svaki incident se klasifikuje kao proizvod, konfiguracija, hardver, mreža ili obuka;
- sve ručne intervencije se evidentiraju jer skrivaju stvarni support cost.

---

## 9. Faza 4 — Field Operations productization

**Ciljni prozor:** 30–90 dana nakon završetka pilota  
**Primarni cilj:** odlučiti da li PWA Otkupac i kiosk paket prelaze u `Standard offer`, ostaju pilot ili se scope smanjuje.

### 9.1 GO uslovi

- pilot KPI-jevi prolaze dogovorene pragove;
- onboarding stanice je dokumentovan;
- postoji podržani hardware bill of materials;
- poznat je support cost po stanici;
- postoji cenovni model;
- postoji zamena uređaja i incident SLA;
- release gate je automatizovan koliko je realno;
- najmanje jedna druga firma može da prođe isti setup bez izmene glavnog koda.

### 9.2 NO-GO / HOLD uslovi

- sistem zahteva stalnu developersku intervenciju;
- štampa ili sync nisu dovoljno pouzdani;
- svaka stanica zahteva posebnu verziju;
- support cost poništava cenu;
- klijent koristi samo mali deo vrednosti;
- canonical desktop i PWA podaci često zahtevaju ručno pomirenje.

### 9.3 Moguće odluke

1. `Standard offer` kao dodatak Enterprise-u;
2. `Pilot only` još jednu sezonu;
3. standardizovati samo PWA unos bez hardverskog paketa;
4. zadržati kiosk/štampu samo za odobrene modele;
5. privremeno stopirati i vratiti kapacitet na Enterprise Core.

---

## 10. Faza 5 — Gazdinstvo validation

**Ciljni prozor:** priprema od Q4 2026, kontrolisani test tokom 2027.  
**Primarni cilj:** dokazati da korisnici aktiviraju i zadržavaju proizvod pre većeg ulaganja u nove premium funkcije.

Gazdinstvo se razvija paralelno samo u ograničenom kapacitetu dok Field Operations i Enterprise Core ne prođu ključne gate-ove.

### 10.1 Prvi dokazni scope

- Partner nalog povezan sa hladnjačom;
- kartica i saldo;
- parcele i kultura;
- tretmani;
- osnovni troškovi;
- osnovna sezonska slika;
- jednostavan onboarding;
- pouzdan offline/sync tok.

### 10.2 Ne prioritetizovati pre dokaza

- širok AI/Digitalni agronom scope;
- kompleksne premium automatizacije;
- veliki broj spoljašnjih integracija;
- funkcije koje povećavaju agronomsku odgovornost bez stručnog owner-a;
- Pro paket zasnovan samo na pretpostavljenoj vrednosti.

### 10.3 KPI gate

Pre skalirane komercijalizacije izmeriti:

- pozvani → registrovani;
- registrovani → prvi unos;
- 30/90/180-day active rate;
- broj aktivnih parcela;
- broj tretmana/troškova po aktivnom korisniku;
- retenciju između sezona;
- Partner → Basic/Pro konverziju;
- willingness-to-pay;
- support cost;
- uticaj na kvalitet Enterprise podataka.

### 10.4 Odluka posle validacije

- skalirati Partner kao B2B2C kanal;
- prodavati Basic/Pro direktno;
- zadržati samo funkcije koje poboljšavaju odnos hladnjača–kooperant;
- ili stopirati komercijalno širenje dok se ne pronađe jači use case.

---

## 11. Faza 6 — GGAP discovery i pilot

**Najraniji početak punog discovery-ja:** nakon što postoji owner kapacitet i nije ugrožena sezonska spremnost Enterprise/Field Operations toka.  
**Primarni cilj:** potvrditi sadržaj, odgovornost, kupca i ekonomiku pre produkcionog razvoja.

### 11.1 Preduslovi

- izabran standard i verzija;
- stručni domain owner;
- imenovani pilot-klijent;
- kompletna lista evidencija, dokaza i procedura;
- mapiranje svakog polja na Enterprise/Gazdinstvo izvor ili ručni unos;
- definisane uloge, odobravanja i odgovornost;
- pravne i marketinške granice;
- pricing hipoteza i support-cost model.

### 11.2 Dozvoljeni prvi scope

- readiness dashboard;
- lista nedostajućih dokaza;
- rokovi i odgovorne osobe;
- ograničen broj automatski popunjenih evidencija;
- prilog/dokaz i provenance;
- kontrolisani export pilot paketa.

### 11.3 Stop pravilo

Ne graditi punu biblioteku listi, automatizacija i audit workflow-a pre potvrde:

- da kupac plaća;
- da automatsko popunjavanje stvarno smanjuje rad;
- da sadržaj može stručno da se održava;
- da support i liability ne prevazilaze prihod.

---

## 12. Opcione Enterprise ekstenzije

Finance, Logistics, Agrohemija i Advanced Warehouse ne razvijaju se ravnopravno u svakom kvartalu.

### 12.1 Pravilo prioriteta

Ekstenzija dobija roadmap kapacitet kada postoji najmanje jedan od sledećih uslova:

- aktivni klijent ima blocking problem;
- kvalifikovani prodajni račun ima visok process fit i finansira reusable rad;
- promena uklanja ozbiljan support ili data-risk;
- promena je preduslov za Field Operations/Gazdinstvo/GGAP gate;
- više klijenata traži isti tok.

### 12.2 Poseban adapter

Nova banka, ERP, vaga ili printer integracija ulazi u roadmap samo uz:

- definisan format i owner-a spoljnog sistema;
- testne podatke;
- finansiran razvoj;
- acceptance kriterijume;
- maintenance model;
- reusable arhitekturu;
- fallback.

---

## 13. Kapacitet osnivača i tima

Dok osnivač ostaje glavni developer, preporučeni capacity budget je:

### Period pre završetka Faze 0

| Oblast | Udeo kapaciteta |
|---|---:|
| P0/P1 correctness, release i incident rizik | 50% |
| Enterprise productization i onboarding | 25% |
| Field Operations pilot readiness | 20% |
| Gazdinstvo/GGAP discovery | 5% |

### Posle završetka Faze 0, pre sezonskog pilota

| Oblast | Udeo kapaciteta |
|---|---:|
| Enterprise Core i P1 stabilnost | 30% |
| Onboarding/migracija/support enablement | 20% |
| Field Operations | 35% |
| Gazdinstvo validation | 10% |
| GGAP discovery | 5% |

### Tokom kritične sezone

| Oblast | Udeo kapaciteta |
|---|---:|
| Stabilnost, monitoring i incidenti | 50% |
| Pilot support i korekcije | 30% |
| Onboarding postojećeg standardnog scope-a | 15% |
| Novi discovery | 5% |

Ovo nisu timesheet kvote, već zaštita od tihog preuzimanja kapaciteta od strane novih ideja i pojedinačnih klijentskih zahteva.

---

## 14. Uloga dodatnog developera

Prvi dodatni developer ne treba odmah da dobije najširi domen. Prioritetni ownership redosled:

1. test automation i release tooling;
2. PWA/GAS regression i fixture okruženje;
3. jasno ograničene P1 popravke;
4. dokumentovani adapteri i onboarding alati;
5. Field Operations stabilizacija;
6. tek kasnije samostalni proizvodni domen.

Osnivač zadržava:

- arhitektonske odluke;
- kritični document/finance model;
- product prioritization;
- ključne klijentske discovery odluke;
- release approval dok ownership i review nisu dokazano preneti.

---

## 15. Portfolio decision gates

| Celina | Sledeći gate | Dokaz potreban za napredovanje |
|---|---|---|
| Enterprise Core | repeatable onboarding | novi klijent bez founder-only koraka |
| Finance & Regulatory | reusable extension | podržan format, regression i poznat maintenance cost |
| Logistics & Fleet | production evidence | realan rad, poznat support cost i bez ručnog pomirenja |
| Agrohemija | process correctness | tačno zaduženje/cena/storno i jasna odgovornost preporuke |
| Advanced Warehouse | scoped repeatability | potvrđen use case bez tvrdnje da je generički WMS/MES |
| PWA Otkupac | seasonal pilot | realni obim, offline, sync, duplicate i print KPI |
| Kiosk/štampa | approved hardware package | standardni modeli, recovery, zamena i marža |
| Gazdinstvo | retention/WTP | activation, 30/90/180-day retention i support cost |
| GGAP | paid controlled pilot | domain owner, sadržaj, provenance i liability granice |

---

## 16. Roadmap anti-prioriteti

Do završetka osnovnih gate-ova ne prioritetizovati:

- rewrite platforme bez merljivog problema;
- generički ERP scope;
- generički WMS/TMS/MES;
- regionalnu lokalizaciju bez domaće ponovljivosti;
- neograničen broj novih bankarskih/ERP adaptera;
- premium Gazdinstvo funkcije bez activation dokaza;
- pun GGAP proizvod bez stručnog owner-a;
- nove AI funkcije koje ne rešavaju potvrđen proces;
- kozmetički redizajn ispred data-safety i sezonske pouzdanosti;
- custom zahtev koji koristi samo jedan klijent i ne finansira održavanje.

---

## 17. Kvartalni decision review

Na kraju svakog kvartala pregledati:

1. otvorene P0/P1 probleme;
2. production gate rezultate;
3. onboarding vreme;
4. support sate i incidente;
5. aktivne klijente i stanice;
6. korišćenje modula;
7. Field Operations pilot KPI;
8. Gazdinstvo activation/retention;
9. prihod i marginu po proizvodnoj celini;
10. founder capacity i bus factor;
11. nove zahteve koji pokušavaju da naprave fork;
12. da li inicijativu ubrzati, zadržati, smanjiti ili zaustaviti.

Svaki review završava jednom od odluka:

- `CONTINUE`;
- `ACCELERATE`;
- `HOLD`;
- `REDUCE SCOPE`;
- `STOP`;
- `MOVE TO STANDARD OFFER`.

---

## 18. KPI-jevi roadmap-a

### Core i release

- otvoreni P0/P1 po verziji;
- procenat mandatory gate-ova koji se izvršavaju automatski ili checklistom;
- false-green incidenti;
- escaped defects;
- vreme od greške do detekcije i oporavka.

### Productization

- onboarding sati po firmi;
- broj ručnih founder-only koraka;
- procenat konfiguracije bez izmene koda;
- vreme pripreme nove sezone;
- support sati po firmi;
- broj reusable naspram client-only zahteva.

### Field Operations

- sync, duplicate, data-loss i print KPI;
- support po stanici;
- vreme postavljanja nove stanice;
- pilot → standard conversion.

### Gazdinstvo i GGAP

- activation, retention, WTP i support cost;
- procenat automatski izvedenih podataka;
- broj eliminisanih ponovnih unosa;
- pilot revenue i gross margin.

---

## 19. Predložene odluke

### PRD-001 — Core safety pre novog scope-a

Otvoreni P0 data-safety, statusni ili authorization problem ima prednost nad novom funkcionalnošću.

### PRD-002 — Enterprise Core ostaje primarni komercijalni proizvod

Najveći deo kapaciteta do sledeće sezone ide na pouzdanost, ponovljiv onboarding, migraciju i Field Operations koji povećava vrednost Enterprise-a.

### PRD-003 — Field Operations kroz jednu kontrolisanu sezonu

PWA Otkupac, kiosk i termalna štampa ne prelaze u standardnu prodaju bez realnog sezonskog pilota i definisanih KPI pragova.

### PRD-004 — Gazdinstvo se validira kroz ponašanje korisnika

Nove premium funkcije ne zamenjuju dokaz activation-a, retencije, willingness-to-pay-a i support cost-a.

### PRD-005 — GGAP počinje sadržajem i odgovornošću, ne kodom

Pun razvoj počinje tek kada postoje stručni owner, standard/verzija, pilot-klijent, mapiranje podataka i ekonomska hipoteza.

### PRD-006 — Sezonski freeze je obavezan

Kritični tokovi se zamrzavaju pre sezone; kasne funkcije se odlažu umesto da ugroze pouzdanost.

### PRD-007 — Stopiranje je legitimna roadmap odluka

Inicijativa koja ne prolazi dokazni ili ekonomski gate smanjuje scope ili se zaustavlja, bez opravdavanja prethodno uloženog rada.

---

## 20. Neposredne naredne akcije

1. Rebazirati i potvrditi dva aktivna P0 nalaza na trenutnom `main`.
2. Napraviti jedan release-evidence zapis za trenutni produkcioni build.
3. Izvršiti target-workbook Production Health Check i relevantne runtime gate-ove.
4. Zaključati standardni Enterprise Core onboarding scope.
5. Izabrati jednog Field Operations pilot-klijenta i 1–3 početne stanice.
6. Zaključati podržani tablet/printer kandidat i failure/recovery test plan.
7. Definisati Field Operations KPI baseline i success/abort pragove.
8. Početi merenje onboarding i support vremena po klijentu.
9. Ograničiti Gazdinstvo rad na activation/retention dokazni scope.
10. GGAP zadržati u discovery-ju do izbora stručnog owner-a i pilot-klijenta.
