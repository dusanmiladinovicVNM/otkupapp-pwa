# AgriX Fast Track Sales Motion

**Status:** DRAFT v1 — VALIDATION  
**Datum:** 02.08.2026.  
**Vlasnik:** osnivač AgriX-a  
**Povezani dokumenti:** `03_BUYING_PROCESS.md`, `04_SALES_PROCESS.md`, `05_DISCOVERY_PLAYBOOK.md`, `08_DEMO_PLAYBOOK.md`, `10_NEGOTIATION_PLAYBOOK.md`, `13_CRM_PIPELINE_PLAYBOOK.md`, `15_ANNUAL_SALES_CALENDAR.md`

---

## 1. Svrha

Fast Track je dodatni prodajni motion za standardne, manje složene AgriX prilike kod kojih bi puni enterprise ritam stvarao više pre-sales rada nego što je potrebno za sigurnu odluku i implementaciju.

Fast Track:

- **ne zamenjuje** Standard ili Complex motion;
- **ne briše** nijedan postojeći stage, kriterijum ili komercijalni guardrail;
- koristi isti kanonski stage model iz `04_SALES_PROCESS.md`;
- dozvoljava da se više stage exit kriterijuma završi tokom istog razgovora;
- smanjuje broj sastanaka i količinu dokumentacije samo kada su scope, rizik i buying committee jednostavni;
- automatski prelazi u Standard motion kada se pojavi složenost koja ne može bezbedno da se obradi ubrzanim putem.

Osnovno pravilo:

> Fast Track ubrzava redosled rada, ali ne preskače dokaz.

---

## 2. Kanonske faze

Fast Track koristi iste faze kao ceo AgriX Commercial Operating System:

| Faza | Naziv |
|---|---|
| S0 | Target Account |
| S1 | Connected |
| S2 | Qualified Problem |
| S3 | Discovery |
| S4 | Solution Evaluation |
| S5 | Risk Alignment |
| S6 | Proposal Review |
| S7 | Decision / Approval |
| S8 | Closed Won |
| S9 | Closed Lost |
| SN | Nurture |

Nije dozvoljeno praviti posebne Fast Track stage oznake, drugačije forecast kategorije ili paralelni pipeline.

Fast Track opportunity može, na primer, tokom jednog sastanka preći iz S2 u S4, ali CRM mora zabeležiti:

- dokaz za izlazak iz S2;
- discovery zaključke za S3;
- rezultat evaluacije i gap-ove za S4;
- vreme ulaska i izlaska iz svake faze, makar faze bile završene istog dana.

---

## 3. Sales motion polje

Svaka opportunity dobija polje `sales_motion`:

- `FAST_TRACK`;
- `STANDARD`;
- `COMPLEX`.

Motion se bira prema složenosti odluke i isporuke, ne prema želji da se posao zatvori brže.

Promena iz `FAST_TRACK` u `STANDARD` ili `COMPLEX` je dozvoljena u bilo kom trenutku i ne smatra se neuspehom. Promena u suprotnom smeru dozvoljena je samo ako je prethodno procenjena složenost dokazano nestala, a ne zato što kupac traži prečicu.

---

## 4. Obavezni eligibility kriterijumi

Fast Track je dozvoljen samo kada su **svi** sledeći uslovi ispunjeni:

1. kupac je jedno pravno lice ili je komercijalni i operativni scope jedne instance potpuno jasan;
2. početni scope obuhvata najviše jednu do dve otkupne stanice ili drugi uporedivo jednostavan deployment;
3. bira se postojeći standardni paket i potvrđeni moduli iz važećeg cenovnika;
4. nema kritičnog custom razvoja pre go-live-a;
5. nema nove, neproverene integracije sa ERP-om, bankom, SEF-om ili trećim sistemom;
6. migracija je mala, standardna ili nije potrebna;
7. economic buyer je direktno uključen ili je put odobrenja kratak i potpuno poznat;
8. problem owner i budući ključni korisnik dostupni su za kombinovani discovery/demo;
9. nema skrivenog procurement, legal, IT ili porodičnog veta koji tek treba mapirati;
10. implementacioni termin je realan i postoji dovoljno kapaciteta;
11. cena se formira direktno iz objavljenog cenovnika bez individualnih popusta;
12. kupac prihvata standardizovan proces i ne traži trajni privatni fork.

Fast Track nije nagrada za visok interes. To je motion za nisku složenost.

---

## 5. Automatski izlazak iz Fast Track-a

Opportunity odmah prelazi u `STANDARD` ili `COMPLEX` kada se pojavi bilo šta od sledećeg:

- više pravnih lica, instanci ili složena vlasnička struktura;
- više od dve stanice u početnom scope-u uz različite procese;
- novi ERP ili integracioni zahtev;
- custom razvoj koji utiče na kritični tok;
- nepoznat economic buyer ili nepoznat approval path;
- više funkcionalnih ili tehničkih evaluatora;
- potreba za formalnim ROI modelom;
- migracija iz više nepouzdanih izvora;
- poseban SLA, ugovorne izmene ili posebna dinamika plaćanja;
- visok vendor-continuity ili security rizik;
- pilot sa više lokacija ili složenim acceptance kriterijumima;
- neusaglašeni stakeholderi;
- implementation timing koji ugrožava sezonu;
- buyer traži prečicu uprkos otvorenom kritičnom pitanju.

Kada se motion promeni, ne otvara se nova opportunity. Postojeći zapis nastavlja kroz isti kanonski stage model.

---

## 6. Ciljni Fast Track ritam

Fast Track teži završetku evaluacije kroz dva do četiri kupčeva razgovora. To nije obećani rok niti obavezni broj sastanaka.

### Razgovor 1 — kvalifikacija i mini-discovery

**Trajanje:** približno 20–30 minuta.  
**Tipičan stage put:** S1 → S2, a kada je dovoljno dokaza i početak S3.

Obavezno potvrditi:

- trigger i razlog zašto se tema razmatra sada;
- problem i najmanje jednu posledicu;
- trenutni proces u osnovnim crtama;
- broj firmi, stanica i korisnika;
- postojeće sisteme;
- desired outcome;
- economic buyer i process owner;
- sezonski rok;
- verovatni standardni paket;
- da li postoji bilo koji Fast Track disqualifier.

Ishod:

- zakazan kombinovani discovery/demo;
- prelazak u Standard/Complex;
- SN Nurture;
- S9 Closed Lost/Disqualified.

### Razgovor 2 — kombinovani discovery i ciljani demo

**Trajanje:** približno 45–60 minuta.  
**Tipičan stage put:** završetak S3 i S4.

Učesnici:

- problem owner;
- economic buyer kada je moguće;
- stvarni ključni korisnik ili osoba koja poznaje proces.

Tok:

1. potvrda cilja i scope-a razgovora;
2. prolazak jednog stvarnog procesa od početka do kraja;
3. dva do tri kritična izuzetka;
4. 3–5 success criteria;
5. ciljani demo samo relevantnog workflow-a;
6. fit/gap klasifikacija;
7. potvrda da nema skrivenog tehničkog ili organizacionog rizika;
8. preliminarni paket, moduli i implementacioni obim;
9. sledeća odluka sa datumom.

Demo nije opšti obilazak proizvoda.

### Razgovor 3 — scope, rizik i proposal review

**Trajanje:** približno 20–40 minuta.  
**Tipičan stage put:** S5 → S6.

S5 i S6 mogu biti završeni u istom razgovoru samo kada:

- svi ulazi za cenu potiču iz standardnog cenovnika;
- nema otvorenog kritičnog gap-a;
- nema custom procene;
- implementation outline je kratak i jasan;
- kupac je pre sastanka dobio ili na sastanku dobija tačan scope;
- relevantna osoba može da potvrdi razumevanje obima i odgovornosti.

Ponuda se i u Fast Track-u predstavlja. Samo slanje PDF-a nije S6.

### Razgovor 4 — decision / contracting

**Trajanje:** prema potrebi.  
**Tipičan stage put:** S7 → S8 ili S9/SN.

Obavezno potvrditi:

- finalni scope i cenu;
- šta nije uključeno;
- odluku i potpisni korak;
- datum početka;
- project owner-a;
- prve obaveze kupca;
- implementation handoff.

Ako kupac može formalno da odluči tokom trećeg razgovora i svi kriterijumi su ispunjeni, poseban četvrti razgovor nije obavezan.

---

## 7. Fast Track PACT minimum

Opportunity ne ulazi u aktivni Fast Track bez:

- **Problem:** konkretno trenje ili cilj;
- **Authority:** direktan decision maker ili potpuno poznat kratki approval path;
- **Consequence:** operativna, finansijska, regulatorna ili upravljačka posledica;
- **Timing:** realan decision date i najkasniji bezbedan implementation start.

Fast Track ne sme da postane put za preskakanje authority ili consequence discovery-ja.

---

## 8. Progressive CRM minimum

Fast Track smanjuje administraciju progresivnim unosom, ali ne ukida source of truth.

### Do S2

Obavezno:

- account i relevantan kontakt;
- `sales_motion = FAST_TRACK`;
- source;
- trigger;
- problem;
- consequence;
- problem owner;
- economic buyer poznat/nepoznat;
- okvirni scope;
- sezonski rok;
- next step, owner i date.

### Do izlaska iz S4

Dodati:

- current-state sažetak;
- success criteria;
- fit zaključak;
- gap klasifikaciju;
- broj firmi, stanica, korisnika i uređaja;
- paket i module — preliminarno;
- top risk;
- buying committee;
- implementation deadline.

### Do S6

Dodati:

- finalni scope;
- cenu i izvor cenovnika;
- implementacione pretpostavke;
- out-of-scope;
- decision process;
- decision date;
- proposal version i review ishod.

### Pre S8

Dodati:

- formalno prihvatanje;
- ugovorenu vrednost;
- project owner-e;
- handoff paket;
- kickoff ili prvi onboarding korak;
- razlog dobitka i ključni dokaz.

Nepoznato polje se označava `UNKNOWN`; ne popunjava se pretpostavkom.

---

## 9. Fast Track scope guardrails

Dozvoljeno:

- standardni Desktop ili Mobile paket;
- potvrđeni standardni moduli;
- standardna implementacija;
- mali broj stanica;
- jednostavan import početnih podataka;
- objavljena cena;
- standardni ugovor i način plaćanja.

Nije dozvoljeno bez prelaska u Standard/Complex:

- „brzo ćemo dodati“ kritičnu funkciju;
- neprocenjena integracija;
- složen custom izveštaj kao uslov kupovine;
- poseban SLA;
- neograničena migracija ili obuka;
- pilot bez cilja i decision review-a;
- usmeni izuzetak od ugovora ili cenovnika;
- implementacija tokom kritičnog perioda samo zato što je kupac spreman da plati.

---

## 10. Cena i ponuda

Fast Track ne uvodi novu cenu niti popust.

Važe:

- `AgriX_Cenovnik_2027.html` kao izvor istine;
- objavljene jedinice obračuna;
- zabrana individualnih i pregovaračkih popusta;
- promena ukupne cene samo kroz stvarnu promenu scope-a;
- druga i naredne instance samo prema objavljenom pravilu;
- custom razvoj van osnovne ponude i tek posle procene.

Ponuda mora jasno navesti:

- paket;
- module;
- firme/instance;
- stanice;
- implementaciju i obuku koja je stvarno uključena;
- period i način plaćanja;
- out-of-scope;
- pretpostavke kupca;
- datum važenja;
- implementation window.

---

## 11. Implementacioni capacity gate

Pre proposal review-a proverava se:

- raspoloživ termin;
- potrebno vreme konfiguracije;
- spremnost podataka;
- dostupnost ključnog korisnika;
- obuka;
- buffer pre sezone;
- support opterećenje postojećih klijenata.

Status:

- **GREEN:** Fast Track rollout može pouzdano da se prihvati;
- **YELLOW:** moguć je samo uz precizan termin, manji scope ili dodatni buffer;
- **RED:** ne prodaje se obećanje go-live-a u traženom periodu; nudi se naredni prozor, manji legitimni scope ili SN Nurture.

Fast Track nikada ne zaobilazi capacity gate.

---

## 12. Fast Track SLA

Početni standardi za validaciju:

- odgovor na inbound: isti radni dan;
- recap prvog razgovora: isti ili naredni radni dan;
- kombinovani discovery/demo: idealno u roku od 7 radnih dana;
- scope i otvorena pitanja: najkasnije 2 radna dana posle evaluacije;
- standardna ponuda: do 2 radna dana nakon potvrđenog scope-a;
- proposal review: zakazuje se pre slanja ili zajedno sa slanjem ponude;
- decision follow-up: prema kupčevom potvrđenom datumu.

SLA nije razlog da se pošalje nepotpuna ili neproverena ponuda.

---

## 13. Forecast pravila

Fast Track koristi iste forecast kategorije:

- Pipeline;
- Best Case;
- Commit;
- Closed.

Brzina prilike ne povećava automatski forecast confidence.

Commit je dozvoljen samo u S7 kada postoje:

- economic buyer;
- prihvaćen finalni scope i cena;
- poznat potpisni korak;
- decision date u periodu;
- rešeni materijalni prigovori;
- kupčev eksplicitni commitment;
- realan implementation slot.

---

## 14. Fast Track quality check

Pre S6 sva pitanja moraju imati odgovor `DA`:

- Da li je Fast Track eligibility i dalje ispunjen?
- Da li su problem i posledica potvrđeni?
- Da li je economic buyer direktno uključen ili approval path potpuno poznat?
- Da li je realni proces dovoljno mapiran?
- Da li su success criteria potvrđeni?
- Da li je demo potvrdio kritični workflow?
- Da li su svi gap-ovi klasifikovani?
- Da li je scope standardan i merljiv cenovnikom?
- Da li je implementation capacity GREEN ili odobren YELLOW?
- Da li je proposal review zakazan?
- Da li postoji mutual next step?

Jedan odgovor `NE` vraća opportunity na nezavršenu fazu ili je prebacuje u Standard/Complex motion.

---

## 15. Primeri

### Primer A — dozvoljen Fast Track

Jedna hladnjača, jedno pravno lice, jedna stanica, standardni Desktop ili Mobile scope, vlasnik direktno odlučuje, nema custom razvoja ni nove integracije, podaci su spremni i postoji realan predsezonski termin.

### Primer B — nije Fast Track

Dve povezane firme, pet stanica sa različitim procesima, BizniSoft integracija, poseban tok ambalaže, migracija iz više Excel fajlova i zahtev za poseban SLA.

Motion: `COMPLEX`.

### Primer C — počeo kao Fast Track, prešao u Standard

Jedno pravno lice i jedna stanica, ali tokom demo-a finansije zahtevaju novu dvosmernu ERP integraciju kao uslov odluke.

Opportunity ostaje ista, stage se vraća na najniži nezavršen zadatak, `sales_motion` se menja u `STANDARD`, a integracija dobija vlasnika i validation plan.

---

## 16. Zabranjeni obrasci

- označiti složenu priliku kao Fast Track radi kraćeg sales cycle-a;
- preskočiti discovery zato što vlasnik „već zna šta želi“;
- prikazati generički demo i odmah poslati ponudu;
- tretirati javni cenovnik kao dovoljan scope;
- preći u S6 bez proposal review termina;
- obećati custom razvoj bez procene;
- koristiti usmeni interes kao Commit;
- prodati implementation rok bez capacity gate-a;
- držati Fast Track priliku otvorenom kada eligibility više ne postoji;
- otvoriti novi opportunity samo da bi se sakrio povratak u raniju fazu.

---

## 17. Validacioni plan

Prvih 10 Fast Track prilika prati:

- razlog izbora motion-a;
- broj kupčevih razgovora;
- vreme od S2 do odluke;
- procenat prilika prebačenih u Standard/Complex;
- najčešći razlog eskalacije;
- broj ponuda bez dodatnog scope pojašnjenja;
- proposal-to-won conversion;
- implementation rework nastao zbog nedovoljnog discovery-ja;
- forecast tačnost;
- zadovoljstvo kupca procesom odluke;
- stvarni pre-sales sati po dobijenom poslu.

Fast Track v1 se ne smatra potvrđenim dok najmanje pet dobijenih poslova ne prođe implementation handoff bez materijalnog iznenađenja koje je prodaja morala ranije da otkrije.

---

## 18. Definition of Done

Fast Track prelazi u `DONE v1` kada:

- koristi isti kanonski stage model kao Standard i Complex motion;
- CRM ima obavezno `sales_motion` polje;
- eligibility i escalation pravila se primenjuju;
- najmanje 10 prilika je vođeno kroz motion;
- najmanje pet je stiglo do implementacije;
- nije nastao sistemski rework zbog preskočenog discovery-ja ili risk alignment-a;
- stage conversion, sales cycle i pre-sales sati mogu da se porede sa Standard motion-om;
- postoji dokumentovana v2 revizija zasnovana na stvarnim podacima.
