# AgriX Website ↔ Sales Alignment Review

**Status:** ACTION REQUIRED  
**Datum pregleda:** 02.08.2026.  
**Vlasnik:** osnivač AgriX-a  
**Svrha:** Uporediti javne poruke, tvrdnje, pakete, CTA-e i prodajne motion-e sa AgriX Commercial Operating System-om i zvaničnim komercijalnim izvorima istine.

---

## 1. Pregledani izvori

### Javni sajt

- `https://agrix.rs/` — sadržaj koji odgovara početnoj platforma stranici; direktni fetch `index-dd.html` nije bio dostupan, pa su provereni javna indeksirana verzija i izvorni `index.html` u repozitorijumu sajta;
- `https://agrix.rs/otkup-dd.html`;
- `https://agrix.rs/gazdinstvo-dd.html`.

### Sales source of truth

- `01_MARKET_POSITIONING.md`;
- `03_BUYING_PROCESS.md`;
- `04_SALES_PROCESS.md`;
- `04A_FAST_TRACK_SALES_MOTION.md`;
- `06_EMAIL_SEQUENCES.md`;
- `08_DEMO_PLAYBOOK.md`;
- `09_OBJECTION_HANDLING.md`;
- `10_NEGOTIATION_PLAYBOOK.md`;
- `11_CASE_STUDIES_PLAYBOOK.md`;
- `12_ROI_CALCULATOR_PLAYBOOK.md`;
- `13_CRM_PIPELINE_PLAYBOOK.md`;
- `15_ANNUAL_SALES_CALENDAR.md`;
- `AgriX_Cenovnik_2027.html`;
- `README.md` i komercijalne odluke na koje upućuje.

---

## 2. Izvršni zaključak

Osnovna tržišna priča sajta je uglavnom usklađena sa Sales strategijom:

- jedan povezani poslovni tok;
- podatak nastaje na mestu rada;
- teren, centrala, dokumentacija, logistika i management nisu odvojeni sistemi;
- AgriX se predstavlja kao više od alata za jedan dokument;
- kupac se poziva na razgovor i demonstraciju zasnovanu na procesu.

Međutim, javne `-dd` stranice trenutno sadrže više **P0 komercijalnih kontradikcija** i **nevalidiranih tvrdnji** koje direktno krše pravila cenovnika, proof hierarchy-ja i Case Studies/ROI playbook-a.

Najkritičniji problem nije stil, već to što kupac na sajtu može dobiti drugačiji komercijalni model od onoga koji Sales mora da ponudi.

---

## 3. Severity model

- **P0 — odmah ispraviti:** pravna, cenovna, reputaciona ili prodajna kontradikcija;
- **P1 — visoki prioritet:** poruka stvara pogrešno očekivanje o scope-u, spremnosti ili procesu kupovine;
- **P2 — srednji prioritet:** slabi kredibilitet, konverziju ili jasnoću;
- **P3 — optimizacija:** poboljšanje nakon ispravke viših prioriteta.

---

# 4. P0 — komercijalne kontradikcije

## 4.1 Mesečno plaćanje i „bez vezivanja“ naspram zvaničnog godišnjeg modela

### Website tvrdnja

Na `otkup-dd.html`:

- paketi se opisuju kao „od dogovora mesečno“;
- navodi se: „Implementacija i obuka su uključeni. Bez vezivanja na ugovor — plaćate mesečno.“

### Zvanični Sales model

`AgriX_Cenovnik_2027.html` definiše:

- cene bez PDV-a **na godišnjem nivou**;
- godišnju pretplatu;
- plaćanje pre početka sezone;
- ugovore koji se sklapaju pre sezone;
- tačno definisane Desktop i Mobile pakete, module, stanice i instance.

### Rizik

Kupac može opravdano očekivati:

- mesečnu pretplatu;
- raskid bez ugovorne obaveze;
- neograničeno uključenu implementaciju i obuku;
- individualno dogovaranje cene.

Sales bi zatim morao da povuče javno obećanje ili napravi nedozvoljeni izuzetak.

### Obavezna akcija

Website mora koristiti isti model kao cenovnik:

> Godišnja pretplata, plaćanje pre sezone. Cena zavisi od paketa, modula, broja aktivnih stanica i instanci. Početna instalacija/povezivanje i pet sati onboarding obuke uključeni su; dodatni rad se obračunava prema važećem cenovniku.

Ne koristiti „bez vezivanja na ugovor“ dok to nije eksplicitna, usklađena poslovna odluka u cenovniku, ugovoru i finansijskom modelu.

---

## 4.2 Website paketi nisu isti kao zvanični paketi

### Website

`otkup-dd.html` koristi:

- Osnovni;
- Standard;
- Po dogovoru;
- neobjavljenu formulaciju „od dogovora“.

### Source of truth

Cenovnik koristi:

- AgriX Desktop — 500 € godišnje;
- AgriX Mobile — 1.500 € godišnje;
- Desktop all-in — 1.200 €;
- Mobile all-in — 2.200 €;
- eksplicitne module i obračunske jedinice;
- stanice preko pet — 50 € godišnje;
- pravilo dodatne instance −50% prema definisanim uslovima.

### Rizik

- lead ne zna da li se paketi sa sajta mapiraju na cenovnik;
- sales qualification i proposal konfiguracija počinju iz različitih kategorija;
- „po meri“ stvara očekivanje neograničenog custom proizvoda;
- „od dogovora“ je u konfliktu sa odlukom da nema individualnih pregovaračkih cena.

### Obavezna akcija

Na sajtu koristiti zvanične nazive ili eksplicitnu mapu:

- Desktop;
- Mobile;
- All-in varijante;
- dodatni moduli;
- dodatne stanice i instance;
- „poseban scope“ samo za procenjeni razvoj/usluge, ne za promenu osnovnih cena.

---

## 4.3 „Gazdinstvo gratis uz svaki paket Otkupa“ nije zvanični model

### Website tvrdnja

`otkup-dd.html` više puta navodi:

- `+ AgriX Gazdinstvo (gratis)`;
- „Bez dodatnog troška — Gazdinstvo dolazi gratis uz svaki paket Otkupa.“

### Source of truth

Cenovnik definiše:

- Gazdinstvo Basic: maloprodaja 19 €, kanalska 10 €;
- Gazdinstvo Pro: maloprodaja 39 €, kanalska 20 €;
- prvih 50 **Basic** naloga partner dobija bez naknade;
- proizvođač ima jedan Pro nalog i on ostaje plaćena pretplata po odgovarajućem modelu.

### Rizik

Website proširuje ograničenu komercijalnu pogodnost na sve Gazdinstvo naloge i sve pakete.

### Obavezna akcija

Do usklađivanja koristiti preciznu formulaciju:

> Uz Enterprise partnerstvo dostupno je do 50 Gazdinstvo Basic naloga bez naknade. Gazdinstvo Pro i dodatni nalozi obračunavaju se prema važećem cenovniku.

Ako se donese nova odluka da je drugi bundle besplatan, istovremeno menjati:

1. cenovnik HTML;
2. šablon ponude;
3. ugovor;
4. finansijski model;
5. website.

---

## 4.4 GGAP se prikazuje kao redovna komponenta ponude

### Website

Početna strana predstavlja AgriX kao platformu za otkup, proizvodnju i GlobalGAP.

`otkup-dd.html` u najširem paketu navodi:

- „GGAP dokumentacija (AgriX GGAP)“;
- bez jasne oznake da nije deo standardne ponude.

`gazdinstvo-dd.html` tvrdi da firma dobija „spremnost za GGAP bez naknadnog skupljanja“.

### Source of truth

Zvanični cenovnik navodi:

- GGAP — `na upit`;
- od 1.000 €;
- samo uz potvrdu obima;
- nije deo standardne ponude;
- ne ugovara se bez prethodne potvrde obima.

Sales README dodatno zahteva vidljivu oznaku:

> „na upit, uz potvrdu obima — nije deo standardne ponude“.

### Rizik

Kupac može da razume da je GGAP:

- spreman standardni modul;
- automatski deo Premium/Enterprise paketa;
- uključen u cenu;
- garantovana usklađenost ili audit readiness.

### Obavezna akcija

Na svim javnim mestima uz GGAP dodati:

> Na upit, uz potvrdu obima. Nije deo standardne ponude i ne predstavlja garanciju sertifikacije.

Ne koristiti GGAP kao noseću platformsku tvrdnju dok komercijalni readiness nije formalno promenjen.

---

## 4.5 Placeholder reference, citati i slike su javno vidljivi

Na `otkup-dd.html` i `gazdinstvo-dd.html` postoje javno vidljivi elementi poput:

- „Klijent 02“, „Klijent 03“, „Klijent 04“;
- `[Ime klijenta]`;
- `[Ime proizvođača]`;
- „Placeholder — zameniti pravim citatom + imenom + fotografijom“;
- tekst koji nalaže da se fotografija zameni pravom slikom klijenta.

### Konflikt sa Sales pravilima

`11_CASE_STUDIES_PLAYBOOK.md` zabranjuje:

- izmišljene reference;
- neodobrene citate;
- korišćenje imena, logotipa ili fotografije bez dozvole;
- predstavljanje unapred napisane promotivne priče kao dokaza.

### Rizik

Ovo je direktan signal da stranica nije produkciono spremna i može ozbiljno smanjiti poverenje baš kod kupaca koji već procenjuju vendor risk.

### Obavezna akcija

Dok ne postoje odobrene reference:

- ukloniti ceo placeholder blok iz javnog DOM-a;
- koristiti dokaz proizvoda, procesni screenshot ili anoniman proof card sa stvarnim izvorom;
- ne prikazivati prazne logotipe ili izmišljene pozicije klijenata.

---

# 5. P0/P1 — nevalidirane ili apsolutne tvrdnje

## 5.1 Otkup metričke tvrdnje

Na `otkup-dd.html` pojavljuju se:

- „30s od QR-a do otkupnog lista“;
- „0 poziva tipa ‘koliki mi je saldo’“;
- „1 klik do sledljivosti od parcele do hladnjače“;
- „100% rad i kada nema interneta“;
- „Otkupni list spreman za 30 sekundi“;
- citat o sedam dana sređivanja papira i trenutnoj vidljivosti otpreme.

### Sales standard

Market Positioning, Case Studies i ROI dokumenti zahtevaju:

- definiciju metrike;
- izvor;
- uzorak i period;
- kontekst;
- jasno razdvajanje product capability, procene i klijentskog rezultata;
- zabranu apsolutnih tvrdnji bez dokaza.

### Procena

- `30s` može biti product-performance tvrdnja samo ako postoji ponovljiv test scenario;
- `1 klik` mora precizno definisati početno stanje i akciju;
- `0 poziva` je rezultat klijenta i ne sme biti generičko obećanje;
- `100% rad i kada nema interneta` je apsolutna i verovatno preširoka tvrdnja: offline capability mora biti vezan za tačno potvrđene workflow-e, uređaj, lokalne podatke i sinhronizaciju;
- izmišljeni testimonial se ne sme prikazivati ni kao vizuelni placeholder na javnoj stranici.

### Preporučena zamena dok nema dokaza

- „Otkupni list nastaje neposredno iz terenskog unosa, bez naknadnog prepisivanja.“
- „Potvrđeni terenski workflow-i podržavaju rad bez stalne veze i naknadnu sinhronizaciju.“
- „Sledljivost povezuje evidentirane podatke od kooperanta/parcele do prijema, prema aktivnom scope-u.“
- „Kooperant može da vidi potvrđene kartice i salda bez dodatnog poziva firmi, kada je ta funkcija uključena i podaci su ažurni.“

---

## 5.2 Homepage apsolutne metrike

Početna stranica koristi:

- „0 duplog unosa“;
- „30s od unosa do spremnog zapisa“;
- „100% više preglednosti“.

### Problem

Ove brojke nisu vezane za:

- konkretan proizvod;
- workflow;
- baseline;
- uzorak;
- klijenta;
- metod merenja.

„100% više preglednosti“ nema stabilnu operativnu definiciju i ne može se dokazati u sadašnjem obliku.

### Akcija

Dok ne postoji Claim Register sa dokazima, zameniti kvalitativnim, preciznim porukama:

- jedan unos se ponovo koristi kroz povezane korake;
- dokument nastaje iz već evidentiranog događaja;
- management dobija centralni pregled obuhvaćenih procesa;
- manje paralelnih evidencija i ručnog usaglašavanja.

---

## 5.3 Gazdinstvo agronomske i automatizacione tvrdnje

`gazdinstvo-dd.html` navodi:

- alarm za grad, mraz i bolesti „pre nego što se dogode“;
- automatski izračunate prozore za prskanje;
- aplikacija „preporučuje tačnu dozu“ po deklaraciji preparata;
- sistem automatski prepoznaje parcelu;
- radni sati i lokacija mere se automatski;
- „karenca pod kontrolom“ i jasan odgovor „smete ili ne smete“;
- prognoza za tačne koordinate;
- „firma vidi spremnost mreže za audit“.

### Procena

Ove tvrdnje mogu biti veoma vredne, ali Sales dokumentacija trenutno nema claim-validation standard specifičan za:

- izvore meteo podataka;
- geolokacionu tačnost;
- podržane kulture i bolesti;
- model predikcije i njegove granice;
- ažurnost deklaracija preparata;
- odgovornost korisnika za dozu, karencu i primenu;
- GPS/background ograničenja uređaja;
- razliku između evidencije i sertifikacione spremnosti.

### Obavezna akcija

Pre javne upotrebe svaku tvrdnju klasifikovati kao:

- `LIVE / VERIFIED`;
- `BETA`;
- `LIMITED COVERAGE`;
- `PLANNED`;
- `PROHIBITED UNTIL VALIDATED`.

Posebno ne koristiti „tačna doza“, „bolest pre nego što se dogodi“ i „smete/ne smete“ bez jasno definisanog izvora, ograničenja i pravne/agronomske formulacije.

---

# 6. P1 — positioning i scope gap-ovi

## 6.1 Homepage je širi od preporučenog tržišnog wedge-a

Homepage hero:

> „Kontrola celog poljoprivrednog lanca. U jednom sistemu.“

Sales positioning kao dugoročnu kategoriju podržava vertikalni operativni sistem, ali preporučuje da tržišni ulaz ostane razumljiv i dovoljno uzak.

### Rizik

„Ceo poljoprivredni lanac“ može značiti:

- proizvodnju;
- savetovanje;
- otkup;
- preradu;
- skladište;
- prodaju;
- logistiku;
- finansije;
- compliance;
- sve kulture i sve tipove firmi.

To povećava kategorijsku ambiciju brže od proof kapitala.

### Preporuka

Homepage može zadržati ekosistem, ali hero treba da vodi najjačim potvrđenim wedge-om:

> Operativni sistem za organizovani otkup i povezivanje otkupnih mesta sa centralom.

Gazdinstvo i GGAP zatim se prikazuju kao povezani proizvodi sa sopstvenim statusom i publikom.

---

## 6.2 Otkup hero je upečatljiv, ali sužava proizvod na jedan QR scenario

Hero:

> „Tri QR koda. Ceo lanac do banke.“

Prednost:

- konkretan;
- vizuelan;
- lako pamtljiv;
- demonstrira end-to-end vezu.

Rizik:

- Desktop-only kupac može pomisliti da proizvod nema smisla bez Mobile/QR toka;
- broj „tri“ zavisi od tačnog procesa;
- svi klijenti ne koriste identičan vozač/destinacija workflow;
- „do banke“ može implicirati punu bankarsku automatizaciju u svakom paketu.

Preporuka:

Koristiti ga kao dokazni scenario za Mobile, ne kao univerzalnu kategorijsku tvrdnju:

> Jedan Mobile scenario: od QR kooperanta i preuzimanja robe do prijema, dokumentacije i bankarskog modula.

---

## 6.3 Website ne razlikuje standardnu funkcionalnost od opcionih modula

Otkup stranica u jedinstven tok uključuje:

- Dispatch;
- Hladnjaču/Proizvodnju;
- SEF;
- Banku;
- GGAP;
- integracije;
- Gazdinstvo.

Cenovnik ih tretira kao različite pakete, module ili posebne ponude.

### Akcija

Svaki feature blok označiti:

- uključeno u Desktop;
- uključeno u Mobile;
- dodatni modul;
- samo uz Mobile;
- na upit;
- posebna usluga/integracija.

Website ne sme da ostavi utisak da najširi prikaz predstavlja svaki paket.

---

# 7. P1 — nedostajući prodajni motion za Gazdinstvo

`gazdinstvo-dd.html` predstavlja potpuno drugačiji go-to-market model od Enterprise Sales sistema:

- self-service registracija;
- dvominutni onboarding;
- 14 dana besplatne probe;
- bez kartice;
- godišnja cena po hektaru;
- direktna fizička lica/proizvođači;
- Premium konsultativni model za firme.

Postojeći Commercial Operating System gotovo u potpunosti opisuje konsultativnu B2B prodaju za hladnjače, otkupne firme i složenije organizacije.

### Zaključak

Gazdinstvo zahteva poseban PLG/inside-sales motion, najmanje sa:

- visitor → signup → activated → trial engaged → converted → retained lifecycle-om;
- definicijom activation event-a;
- trial email/onboarding sekvencom;
- in-app support modelom;
- Basic/Pro entitlement mapom;
- payment i renewal pravilima;
- churn i inactive-user signalima;
- jasnim handoff-om kada gazdinstvo postane Enterprise/partner lead;
- razlikovanjem individualnog korisnika i kanalskog naloga preko hladnjače/savetnika.

Dok taj motion ne postoji, website obećava prodajni i onboarding sistem koji Sales dokumentacija ne kontroliše.

---

# 8. P2 — credibility i copy problemi

## 8.1 Produkcioni placeholder tekst

Ukloniti:

- uputstva dizajneru/developeru;
- placeholder imena;
- prazne reference;
- lažne inicijale;
- nepostojeće fotografije.

## 8.2 Jezičke i tehničke greške

Na pregledanim stranicama postoje formulacije poput:

- „skená“ / „skení“;
- „100%rad“;
- „Sami sati radni sati“;
- mešanje `real-time`, „u istom trenutku“ i apsolutne svežine bez definicije;
- neujednačeno korišćenje Management/management/uprava.

Potrebna je završna language QA revizija pre objave.

## 8.3 Bezbednost je navedena bez proof pack-a

Poruka:

> „Vaši podaci. Vaša kontrola. Svaka firma radi u odvojenom i zaštićenom okruženju.“

je korisna, ali zahteva vezu ka konkretnom dokazu:

- gde se podaci čuvaju;
- ko ima pristup;
- backup;
- recovery;
- tenant separation;
- export i vlasništvo nad podacima;
- incident/support procedura.

Generička stranica „više o bezbednosti“ ne sme sadržati placeholder ili neodobrene pravne tvrdnje.

## 8.4 Nedostaje anti-ICP

Sales positioning jasno definiše za koga AgriX nije dobar izbor. Website uglavnom prikazuje samo maksimalni benefit.

Dodati blok:

> AgriX verovatno nije potreban ako imate jedan jednostavan proces koji već pouzdano radi, ne želite da standardizujete podatke ili očekujete neograničene individualne izmene bez vlasnika implementacije.

Ovo povećava kredibilitet i smanjuje nekvalitetne upite.

## 8.5 Nedostaje implementation-risk sadržaj

Website prikazuje rezultat, ali malo objašnjava:

- kako se definiše scope;
- kada se implementira;
- ko priprema podatke;
- šta je uključeno;
- šta se događa kada rok pred sezonu nije realan;
- kako izgleda obuka i handoff;
- kako AgriX koegzistira sa postojećim ERP-om.

Sales dokumenti pokazuju da je implementation anxiety jedan od glavnih buying rizika. Website bi trebalo da ga obradi pre CTA-a.

---

# 9. Alignment po stranici

## 9.1 Početna stranica

### Dobro usklađeno

- platforma povezuje procese;
- podaci nastaju u radu;
- management dobija pregled;
- postoje odvojeni proizvodi Otkup, Gazdinstvo i GGAP;
- CTA nudi razgovor/prezentaciju.

### Potrebna promena

1. suziti hero ili jasno označiti Otkup kao najzreliji tržišni wedge;
2. ukloniti apsolutne metrike;
3. GGAP označiti kao `na upit`;
4. dodati stvaran proof i anti-ICP;
5. dodati proces saradnje i implementation-risk FAQ;
6. ne koristiti platformsku tvrdnju kao dokaz da su svi proizvodi jednako komercijalno spremni.

## 9.2 Otkup stranica

### Dobro usklađeno

- end-to-end tok;
- teren → logistika → dokumentacija → management;
- offline-first vrednost;
- personae Otkupac, Dispečer, Vlasnik;
- CTA prvo poziva na razgovor, zatim demo;
- jaka demonstrabilna Mobile priča.

### Potrebna promena

1. zameniti website pakete zvaničnim Desktop/Mobile modelom;
2. uskladiti godišnje plaćanje i ugovor;
3. precizirati uključenu obuku i dodatne satnice;
4. ispraviti Gazdinstvo bundle;
5. GGAP označiti kao `na upit`;
6. ukloniti placeholders;
7. ukloniti ili dokazati 30s/0/1 klik/100% tvrdnje;
8. označiti opcione module;
9. razdvojiti standardni proizvod od integracija i prilagođenih izveštaja.

## 9.3 Gazdinstvo stranica

### Dobro usklađeno

- jasna publika;
- vrednost nastanka evidencije tokom rada;
- povezivanje parcela, mera, lagera, troškova i bilansa;
- mogućnost povezivanja proizvođača sa firmom;
- jasan CTA za samostalni ulazak i poseban razgovor za organizatore proizvodnje.

### Potrebna promena

1. uskladiti Basic/Pro nazive i cene sa cenovnikom — website koristi Basic/Standard/Premium i cenu po hektaru, dok cenovnik koristi Basic/Pro i fiksnu godišnju cenu po nalogu;
2. potvrditi da 14-day trial i self-service onboarding zaista postoje kao operativno podržan motion;
3. klasifikovati sve agronomske smart funkcije prema readiness-u;
4. dodati disclaimer i granice za dozu, karencu, bolesti i meteo;
5. ukloniti placeholder reference;
6. ne predstavljati audit/GGAP spremnost kao automatski rezultat;
7. definisati poseban Gazdinstvo PLG funnel i support model.

---

# 10. Claim Register — obavezna kontrola

Pre sledećeg website release-a napraviti tabelu:

| Claim ID | Stranica | Tvrdnja | Tip | Status | Dokaz | Ograničenje | Vlasnik | Dozvoljen tekst |
|---|---|---|---|---|---|---|---|---|
| WEB-O-01 | Otkup | 30s do otkupnog lista | performance | unverified | test potreban | uređaj/scenario | Product | TBD |
| WEB-O-02 | Otkup | 100% offline rad | capability | prohibited as written | workflow test | samo potvrđeni tokovi | Product | ograničena formulacija |
| WEB-O-03 | Otkup | 0 poziva za saldo | outcome | prohibited without case study | client data | kontekst klijenta | Sales | samo imenovana/anonimna studija |
| WEB-G-01 | Gazdinstvo | tačna preporučena doza | agronomic | validation required | deklaracije + logika | user responsibility | Product/Legal | TBD |
| WEB-G-02 | Gazdinstvo | predviđa bolesti | predictive | validation required | model + coverage | kultura/region | Product | TBD |
| WEB-C-01 | Platforma | 100% više preglednosti | outcome | prohibited | nema definiciju | — | Marketing | ukloniti |

Nijedna brojčana, regulatorna, sigurnosna, agronomska ili klijentska tvrdnja ne ide u produkciju bez Claim ID-a.

---

# 11. Source-of-truth redosled

Kada postoji konflikt, važi sledeći prioritet:

1. eksplicitna poslovna odluka;
2. `AgriX_Cenovnik_2027.html` za cenu i komercijalni model;
3. ugovor i pravni dokumenti;
4. potvrđena product/readiness dokumentacija;
5. Sales playbook;
6. website copy;
7. pojedinačna prezentacija ili usmena formulacija.

Website nikada ne postaje izvor istine samo zato što je javno objavljen.

---

# 12. Preporučeni release gate za website

Pre deploy-a svake komercijalne stranice:

- [ ] paketi i cene odgovaraju cenovniku;
- [ ] nema individualne ili neodobrene cenovne logike;
- [ ] GGAP i drugi uslovni proizvodi imaju tačnu oznaku;
- [ ] svaki broj ima Claim ID i dokaz;
- [ ] nema placeholder citata, logotipa, imena ili slika;
- [ ] opcioni moduli su označeni;
- [ ] product readiness je potvrđen;
- [ ] implementation i support obim su tačni;
- [ ] CTA odgovara stvarnom sales motion-u;
- [ ] pravni i privacy tekst nema placeholder polja;
- [ ] language QA je završen;
- [ ] Sales i Product owner odobrili su release.

---

# 13. Prioritetni plan korekcije

## P0 — pre javnog korišćenja `-dd` stranica

1. godišnje/mesečno plaćanje i ugovor;
2. zvanični paketi i cena;
3. Gazdinstvo bundle;
4. GGAP oznaka;
5. placeholder reference;
6. apsolutne i nedokazane metrike;
7. Basic/Pro Gazdinstvo model.

## P1 — naredni copy release

1. uži homepage wedge;
2. opcioni moduli i scope mapiranje;
3. implementation-risk sekcija;
4. anti-ICP;
5. security proof pack;
6. poseban Gazdinstvo PLG motion;
7. agronomski Claim Register.

## P2 — optimizacija

1. A/B test hero poruka;
2. CTA po nivou namere;
3. proof cards;
4. persona landing blokovi;
5. case-study integracija;
6. source tracking i conversion analiza.

---

# 14. Konačna ocena usklađenosti

| Oblast | Ocena | Zaključak |
|---|---:|---|
| Osnovni problem narrative | 9/10 | veoma dobro usklađen |
| End-to-end proizvodna priča | 8/10 | jaka, ali često prikazuje maksimalni scope kao standardni |
| ICP i persona relevantnost | 8/10 | dobra za Otkup; Gazdinstvo traži poseban motion |
| CTA i buying journey | 7/10 | dobar razgovor/demo CTA, ali Gazdinstvo uvodi nepokriven PLG funnel |
| Cenovna usklađenost | 2/10 | direktan konflikt sa zvaničnim cenovnikom |
| Proof i claims | 2/10 | placeholders i nevalidirane apsolutne tvrdnje |
| Product/readiness granice | 4/10 | standardni, opcioni i `na upit` scope nisu dovoljno razdvojeni |
| Implementation-risk komunikacija | 4/10 | znatno slabija od Sales dokumentacije |
| Ukupna trenutna spremnost `-dd` stranica | 4/10 | dobra strateška osnova, ali nisu spremne za javnu produkciju bez P0 korekcija |

---

## 15. Definition of Done

Website i Sales smatraju se usklađenim kada:

- nema cenovne ili ugovorne kontradikcije;
- svi paketi mapiraju na zvanični cenovnik;
- GGAP i uslovni proizvodi su pravilno označeni;
- nema javnih placeholder referenci;
- svaka brojčana i outcome tvrdnja ima dokaz;
- standardni i opcioni scope su razdvojeni;
- Gazdinstvo ima potvrđen self-service/PLG operating model ili se CTA prilagodi stvarnom readiness-u;
- Claim Register i website release gate postanu deo redovne procedure;
- najmanje jedna osoba iz Sales-a i jedna iz Product/Delivery-ja odobre svaki komercijalni release.
