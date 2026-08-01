# AgriX KPI Dashboard Playbook

**Status:** DRAFT v1 — VALIDATION  
**Datum:** 30.07.2026.  
**Vlasnik:** osnivač AgriX-a  
**Svrha:** Definisati mali, pouzdan i operativno koristan skup komercijalnih KPI-jeva koji pokazuje kvalitet tržišnog nastupa, pipeline-a, forecast-a, prodajnog procesa, handoff-a i ranih signala zadržavanja i širenja klijenata.

---

## 1. Osnovni princip

Dashboard nije zbir što većeg broja grafikona. Njegova uloga je da odgovori na pet pitanja:

1. Da li dolazimo do pravih naloga i ljudi?
2. Da li se prilike kvalitetno kreću kroz proces?
3. Da li je pipeline dovoljan i realan?
4. Da li forecast može da se koristi za odluke?
5. Da li dobijeni poslovi prelaze u uspešnu implementaciju, korišćenje i širenje?

**DECISION:** Nijedan KPI se ne koristi kao cilj izolovano od kvaliteta. Veći broj poziva, emailova, leadova ili prilika nije uspeh ako ne povećava broj kvalitetnih razgovora, kvalifikovanih prilika i održivih klijenata.

---

## 2. Hijerarhija metrika

### 2.1. Leading indicators

Pokazuju da li se rade aktivnosti koje mogu proizvesti rezultat:

- target accounts researched;
- novi relevantni kontakti;
- kvalitetni prvi razgovori;
- discovery sastanci;
- demo sastanci;
- dogovoreni mutual next steps;
- prilike sa potvrđenim champion-om;
- prilike sa uključenim economic buyer-om;
- proposal review sastanci.

### 2.2. Pipeline indicators

Pokazuju količinu i kvalitet aktivnih prilika:

- broj i vrednost prilika po stage-u;
- weighted pipeline;
- pipeline coverage;
- stage conversion;
- stage aging;
- stalled rate;
- close-date slippage;
- opportunity quality score;
- CRM completeness.

### 2.3. Lagging indicators

Pokazuju ostvareni komercijalni rezultat:

- Closed Won broj i vrednost;
- win rate;
- sales cycle;
- prosečna vrednost posla;
- revenue po izvoru;
- forecast accuracy;
- no-decision rate;
- implementation acceptance;
- activation i adoption signali;
- expansion i renewal signali.

### 2.4. Guardrail metrics

Sprečavaju optimizaciju pogrešne stvari:

- procenat nekvalifikovanih prilika otvorenih duže od praga;
- popusti ili odstupanja od cenovnika;
- obećanja van potvrđenog scope-a;
- won poslovi bez handoff-a;
- rework zbog lošeg discovery-ja;
- churn/risk signal posle prodaje;
- CRM aktivnosti bez meaningful engagement-a.

---

## 3. Standard perioda i segmentacije

Dashboard prikazuje najmanje:

- tekuću nedelju;
- tekući mesec;
- rolling 90 dana;
- sezonu;
- godinu do danas.

Svaki relevantan KPI mora moći da se segmentira po:

- izvoru;
- kulturi ili tržišnom segmentu;
- ICP tier-u;
- veličini naloga;
- proizvodnom paketu;
- novom poslu, expansion-u ili renewal-u;
- vlasniku prilike;
- inbound/outbound/referral kanalu.

Mali uzorci se jasno označavaju. Procenat bez prikazanog broja slučajeva nije dovoljan.

---

## 4. Activity dashboard

### 4.1. Aktivnosti koje se broje

Broje se samo aktivnosti sa jasnim poslovnim značenjem:

- personalizovan prvi kontakt;
- razgovor sa relevantnom ulogom;
- discovery;
- demo;
- scope review;
- proposal review;
- negotiation meeting;
- reference ili referral razgovor;
- dogovoreni naredni korak.

Automatski poslati emailovi, pokušaji poziva bez razgovora i generički bulk outreach prikazuju se odvojeno i ne tretiraju kao engagement.

### 4.2. Ključni activity KPI-jevi

| KPI | Formula / definicija |
|---|---|
| Target account coverage | obrađeni prioritetni nalozi / planirani prioritetni nalozi |
| Contact connect rate | ostvareni razgovori / validni pokušaji kontakta |
| First meeting rate | prvi sastanci / target nalozi kontaktirani u periodu |
| Discovery completion rate | završeni kvalitetni discovery sastanci / zakazani discovery sastanci |
| Demo progression rate | demo sastanci koji vode u potvrđen sledeći korak / svi demo sastanci |
| Mutual next-step rate | aktivne prilike sa korakom, datumom i vlasnikom / sve aktivne prilike |

Activity target se postavlja tek posle dovoljnog uzorka. U početnoj fazi važniji je odnos aktivnosti i napredovanja nego apsolutna norma.

---

## 5. Funnel i conversion dashboard

### 5.1. Stage conversion

`Stage conversion = broj prilika koje su prešle u narednu fazu / broj prilika koje su ušle u posmatranu fazu`

Obavezno se prikazuje:

- conversion po stage-u;
- broj prilika u imenitelju;
- median vreme do prelaska;
- conversion po izvoru i ICP tier-u;
- razlog izlaska iz procesa.

### 5.2. Ključne konverzije

- Lead → Qualified;
- Qualified → Discovery Complete;
- Discovery → Demo;
- Demo → Solution Fit;
- Solution Fit → Proposal;
- Proposal → Negotiation;
- Negotiation → Closed Won;
- Qualified Opportunity → Closed Won.

### 5.3. Win rate

`Win rate = Closed Won / (Closed Won + Closed Lost)`

No-decision se prikazuje zasebno, ali i u dodatnom pokazatelju:

`Decision win rate = Closed Won / (Closed Won + Closed Lost + No Decision)`

Time se sprečava lažno poboljšanje win rate-a zatvaranjem slabih prilika kao deferred ili ostavljanjem bez odluke van metrike.

---

## 6. Pipeline value i coverage

### 6.1. Unweighted pipeline

Zbir vrednosti svih otvorenih prilika u dozvoljenim aktivnim stage-ovima.

### 6.2. Weighted pipeline

`Weighted pipeline = suma(vrednost prilike × stage probability)`

Stage probability se ne tretira kao individualna procena prodavca. Početno se koristi operativna matrica, a kasnije se kalibriše stvarnim istorijskim conversion podacima.

### 6.3. Pipeline coverage

`Coverage = kvalifikovani pipeline za period / cilj za isti period`

Coverage se prikazuje odvojeno za:

- ukupni pipeline;
- qualified pipeline;
- late-stage pipeline;
- Commit;
- novi posao;
- expansion.

**DECISION:** Jedan univerzalni coverage target se ne propisuje pre nego što postoje stabilni podaci o win rate-u, sales cycle-u i sezonskoj raspodeli odluka.

### 6.4. Pipeline concentration risk

Prikazuje se:

- udeo najveće prilike u ukupnom pipeline-u;
- udeo tri najveće prilike;
- zavisnost od jednog segmenta, kulture ili izvora;
- zavisnost od jednog decision window-a.

Visok pipeline koji zavisi od jedne prilike nije zdrav pipeline.

---

## 7. Velocity i sales cycle

### 7.1. Sales cycle

Meri se najmanje:

- od prvog meaningful kontakta do Closed Won/Lost;
- od Qualified Opportunity do Closed Won/Lost;
- vreme po stage-u.

Koriste se median i percentili, ne samo prosek, jer nekoliko ekstremno dugih prilika može iskriviti sliku.

### 7.2. Pipeline velocity

`Pipeline velocity = broj kvalifikovanih prilika × prosečna vrednost × win rate / prosečan sales cycle`

Formula se koristi za trend i poređenje segmenata, ne kao samostalna finansijska prognoza.

### 7.3. Stage aging

Za svaki stage prikazuje se:

- median dana;
- 75. percentil;
- broj prilika iznad aging praga;
- vrednost prilika iznad praga;
- poslednja meaningful aktivnost;
- sledeći korak.

Stalled rate:

`Stalled rate = aktivne prilike bez meaningful napretka duže od praga / sve aktivne prilike`

---

## 8. Forecast dashboard

### 8.1. Forecast kategorije

- Pipeline;
- Best Case;
- Commit;
- Closed.

### 8.2. Forecast accuracy

Za period se prikazuju:

`Forecast accuracy = 1 - |forecast - actual| / actual`

Kada je actual nula, greška se prikazuje apsolutno i ne koristi se navedena formula.

Odvojeno se meri:

- početni forecast perioda;
- forecast na polovini perioda;
- poslednji forecast pre zatvaranja;
- Commit accuracy;
- Best Case conversion.

### 8.3. Slippage

`Slippage rate = prilike čiji je close date pomeren van perioda / prilike planirane za zatvaranje u periodu`

Beleži se:

- broj pomeranja iste prilike;
- razlog;
- stage u trenutku pomeranja;
- da li je postojao stvarni kupčev commitment;
- da li je opportunity trebalo ranije vratiti u nižu kategoriju.

### 8.4. Forecast calibration

Forecast je koristan tek kada se redovno porede:

- prognoza;
- ostvarenje;
- razlog odstupanja;
- vrsta pristrasnosti;
- kvalitet dokaza u trenutku prognoze.

Ne kažnjava se iskreno smanjenje forecast-a kada se pojavi novi dokaz. Kažnjava se zadržavanje nerealne prognoze bez dokaza.

---

## 9. Source quality dashboard

Za svaki izvor prikazuju se:

- broj leadova;
- broj relevantnih kontakata;
- broj kvalifikovanih prilika;
- pipeline value;
- Closed Won;
- win rate;
- median sales cycle;
- prosečna vrednost posla;
- no-decision rate;
- trošak izvora kada je poznat;
- prihod i gross contribution kada su dostupni.

Izvori uključuju najmanje:

- organski sajt/SEO;
- direktna mreža;
- referral;
- outbound;
- partneri i savetodavna tela;
- događaji;
- postojeći klijenti/expansion.

Lead volume bez kvalifikovanih prilika nije kvalitetan izvor.

---

## 10. Qualification i opportunity quality

Prikazuju se:

- procenat prilika sa potvrđenim problemom;
- procenat sa merljivom posledicom;
- procenat sa decision process-om;
- procenat sa economic buyer-om;
- procenat sa champion-om;
- procenat sa mutual next step-om;
- procenat sa potvrđenim timing-om;
- prosečan PACT score;
- prosečan CRM quality score.

Opportunity quality mora moći da se uporedi sa kasnijim win/loss ishodom. Cilj je otkriti koji kriterijumi stvarno predviđaju uspeh, a koji samo izgledaju profesionalno.

---

## 11. Win, loss i no-decision dashboard

Obavezno se prikazuju:

- broj i vrednost Won/Lost/No Decision;
- primarni razlog;
- sekundarni razlog;
- konkurent ili status quo;
- stage izlaska;
- izvor;
- ICP tier;
- sales cycle;
- cena/scope kategorija;
- uključenost champion-a i economic buyer-a;
- kvalitet discovery-ja;
- lesson learned.

Standardne kategorije gubitka:

- no priority;
- no budget;
- timing/sezona;
- status quo;
- konkurent;
- internal build;
- scope mismatch;
- implementation risk;
- vendor risk;
- decision process failure;
- price/value gap;
- lost contact;
- disqualified by AgriX.

No-decision se ne skriva unutar „timing“ ako ne postoji konkretan datum reaktivacije i potvrđen razlog odlaganja.

---

## 12. CRM hygiene dashboard

Minimalni pokazatelji:

- aktivne prilike bez next step-a;
- prilike sa next step datumom u prošlosti;
- prilike bez meaningful aktivnosti;
- close date u prošlosti;
- prilike bez iznosa ili scope-a u fazi gde su obavezni;
- Commit bez svih entry kriterijuma;
- duplikati naloga i kontakata;
- aktivnosti bez ishoda;
- Closed Lost bez razloga;
- Closed Won bez handoff-a;
- deferred prilike bez datuma reaktivacije.

CRM completeness:

`Completeness = popunjena obavezna polja / sva obavezna polja za trenutni stage`

Dashboard prikazuje completeness po opportunity owner-u i stage-u, ali se koristi za coaching i kvalitet procesa, ne za birokratsko popunjavanje bez stvarne informacije.

---

## 13. Implementation handoff dashboard

Komercijalni uspeh ne završava potpisom.

Prikazuju se:

- procenat Won poslova sa kompletnim handoff-om;
- vreme od potpisa do prihvatanja handoff-a;
- broj otvorenih nejasnoća;
- scope mismatch incidenti;
- obećanja koja implementacija nije očekivala;
- nedostajući podaci ili odluke kupca;
- planirani i stvarni početak;
- rework izazvan prodajnim procesom;
- vreme do prvog operational milestone-a.

**DECISION:** Closed Won koji nije prihvaćen od implementacije ostaje vidljiv kao „Won — handoff pending“, ne kao potpuno završen komercijalni rezultat.

---

## 14. Early customer health, retention i expansion signali

Dok ne postoji dovoljan broj renewal ciklusa, koriste se rani signali:

- završena implementacija;
- aktivirani korisnici;
- stvarno korišćeni ključni procesi;
- broj otvorenih kritičnih problema;
- response/resolution trend;
- sponsor engagement;
- ostvareni milestone-i;
- procena vrednosti posle sezone;
- interesovanje za dodatne firme, stanice ili module;
- reference readiness;
- renewal risk signal.

Expansion pipeline se vodi odvojeno od novog posla, ali se povezuje sa customer health stanjem. Ne otvara se expansion opportunity samo zato što dodatni modul postoji.

---

## 15. Executive dashboard — minimalni pregled

Jedna izvršna strana treba da sadrži najviše:

1. Closed Won i cilj za period;
2. qualified pipeline i coverage;
3. Commit i forecast accuracy;
4. stage conversion;
5. median sales cycle;
6. stalled pipeline value;
7. win/loss/no-decision;
8. source quality;
9. CRM quality;
10. handoff i early health signal.

Svaki KPI mora imati:

- trenutnu vrednost;
- prethodni period;
- cilj ili očekivani raspon kada postoji;
- trend;
- objašnjenje značajnog odstupanja;
- vlasnika sledeće akcije.

---

## 16. Operativni dashboard-i

### 16.1. Dnevni pogled

- overdue next steps;
- današnji sastanci i priprema;
- prilike bez odgovora kupca;
- close dates u narednih 14 dana;
- novi inbound zahtevi;
- handoff blockers.

### 16.2. Nedeljni pipeline review

- stage movement;
- nove kvalifikovane prilike;
- stalled i aging;
- forecast promene;
- close-date slippage;
- top rizici;
- next-week commitments;
- prilike za disqualification.

### 16.3. Mesečni management review

- funnel conversion;
- source performance;
- velocity;
- win/loss/no-decision;
- forecast accuracy;
- ICP i segment performance;
- pricing/scope odstupanja;
- handoff quality;
- potrebna promena procesa ili poruke.

### 16.4. Kvartalna/season review

- trendovi uz dovoljan uzorak;
- kalibracija stage probability;
- revizija aging pragova;
- customer evidence;
- case study kandidati;
- ROI pretpostavke vs stvarni rezultati;
- strateška preraspodela fokusa.

---

## 17. Alert pravila

Dashboard treba da označi najmanje:

- Commit bez kupčevog eksplicitnog commitment-a;
- close date pomeren više od dva puta;
- late-stage priliku bez aktivnosti duže od praga;
- opportunity bez next step-a;
- pipeline coverage ispod dogovorenog raspona;
- više od 30% pipeline-a u jednoj prilici;
- rast no-decision rate-a;
- pad conversion-a iz discovery-ja u demo;
- rast proposal slippage-a;
- Won bez handoff-a;
- handoff sa neodobrenim obećanjem;
- customer health red signal.

Alert zahteva akciju, vlasnika i rok. Crvena boja bez procesa odgovora nema operativnu vrednost.

---

## 18. Data governance

Za svaki KPI dokumentuju se:

- naziv;
- poslovna definicija;
- formula;
- izvor podataka;
- vlasnik;
- učestalost osvežavanja;
- dozvoljeni filteri;
- tretman null vrednosti;
- istorijska verzija definicije;
- poznata ograničenja.

Promena definicije KPI-ja ne sme retroaktivno menjati interpretaciju istorije bez oznake verzije.

Obavezna pravila:

- stage history se čuva;
- close-date history se čuva;
- forecast snapshot se čuva po periodu;
- izbrisane prilike se ne koriste za popravljanje rezultata;
- Lost i No Decision ostaju u analitičkom skupu;
- ručne korekcije imaju audit trag.

---

## 19. KPI Quality Score

Dashboard dobija po jedan bod za svaki uslov:

1. definicije su dokumentovane;
2. postoje jasni izvori podataka;
3. periodi su dosledni;
4. mali uzorci su označeni;
5. stage history postoji;
6. forecast snapshot postoji;
7. Won/Lost/No Decision su kompletni;
8. activity i meaningful activity su razdvojeni;
9. pipeline quality je vidljiv;
10. aging i slippage su vidljivi;
11. source quality je vidljiv;
12. handoff je vidljiv;
13. customer health signali su vidljivi;
14. svaki alarm ima akciju;
15. dashboard se koristi u redovnom review-u;
16. odluke iz review-a se evidentiraju.

Tumačenje:

- 14–16: operativno pouzdan;
- 11–13: upotrebljiv uz korekcije;
- 8–10: visok rizik pogrešnih zaključaka;
- manje od 8: reporting postoji, management system ne postoji.

---

## 20. Zabranjeni obrasci

Ne sme se:

- slaviti broj aktivnosti bez ishoda;
- prikazivati procenat bez broja slučajeva;
- menjati stage da bi pipeline izgledao bolje;
- ostavljati izgubljene prilike otvorene;
- koristiti upper ROI scenario kao ostvareni prihod;
- sabirati new business i expansion bez oznake;
- računati pokušaj kontakta kao engagement;
- tretirati weighted pipeline kao siguran prihod;
- prikazivati forecast bez snapshot-a;
- koristiti proseke kada distribucija zahteva median;
- kažnjavati disqualification kvalitetnih razloga;
- optimizovati ljude za CRM popunjenost umesto za kvalitet informacija;
- skrivati loš handoff iza Closed Won rezultata.

---

## 21. Validacioni plan

Prva formalna verzija KPI sistema validira se kroz:

- najmanje 30 zatvorenih prilika ili jednu punu sezonu;
- poređenje stage probability sa stvarnim conversion-om;
- poređenje forecast kategorija sa ostvarenjem;
- analizu najmanje 10 Won, 10 Lost i 10 No Decision ishoda kada obim to dozvoli;
- proveru koji qualification faktori stvarno predviđaju win;
- proveru aging pragova;
- proveru uticaja izvora na kvalitet i cycle;
- proveru veze CRM quality score-a sa rezultatom;
- proveru handoff kvaliteta i rework-a;
- reviziju svih KPI-jeva koji ne vode ka konkretnoj odluci.

KPI koji tokom dva uzastopna review ciklusa ne menja nijednu odluku, pitanje ili akciju kandidat je za uklanjanje.

---

## 22. Veza sa ostalim dokumentima

Ovaj dokument operacionalizuje podatke iz:

- `04_SALES_PROCESS.md`;
- `05_DISCOVERY_PLAYBOOK.md`;
- `08_DEMO_PLAYBOOK.md`;
- `09_OBJECTION_HANDLING.md`;
- `10_NEGOTIATION_PLAYBOOK.md`;
- `11_CASE_STUDIES_PLAYBOOK.md`;
- `12_ROI_CALCULATOR_PLAYBOOK.md`;
- `13_CRM_PIPELINE_PLAYBOOK.md`.

Annual Sales Calendar definiše sezonske ciljeve, očekivani timing i review ritam. KPI dashboard mora te ciljeve pratiti bez mešanja sezonskih i vansezonskih perioda.